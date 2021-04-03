Attribute VB_Name = "mdlLNCASY"
Option Explicit

Private mobjUSBKEY As Object     '辽宁省数字签名 华大 KEYUSBKEYACTIVE.USBKeyActiveCtrl.1
Private mobjMSScriptCtl As Object    'MSScriptControl.ScriptControl.1 微软提供脚本控件 用到javaScript中encodeURI方法获取URL串
Private mblnInit As Boolean
Private mstrLastPwd As String          '缓存输入的密码
Private mintLogin As Integer

'20170817 SM2算法集成  辽宁嘉鸿
Private mbytModel           As Byte             '0-RSA算法;1-SM2算法
Private mobjKeyManager      As Object           '证书管理对象
Private mobjCert            As Object           '证书对象
Private mobjKeyStore        As Object           'UKey操作类KeyStore
Private mobjKeySealArray    As Object
Private mobjKeySeal         As Object           '签章类
Private mobjKeyGateOper     As Object
Private mobjKeyDetector     As Object           'JHKey.KeyDetector.1.1
Private Enum E_Model
    E_RSA = 0
    E_SM2 = 1
End Enum

Public Function LNCA_Initialize() As Boolean
    '功能:创建辽宁CA控件对象
    
    Dim intRet As Integer
    Dim varTmp As Variant
    
    On Error GoTo errH
   
        If mblnInit Then LNCA_Initialize = True: Exit Function
        
        gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys) '读取URL 服务器
        'gstrPara = "http://218.25.86.214:2010/ssoworker"  '测试地址
        If gstrPara = "" Then
            MsgBox "没有配置认证服务器地址，请到【公共参数设置】先配置:" & vbCrLf & vbTab & "系统号100,参数号90000" & _
                    vbCrLf & vbTab & "参数值格式""http://218.25.86.214:2010/ssoworker""", vbInformation, gstrSysName
            Exit Function
        End If
        varTmp = Split(gstrPara, G_STR_SPLIT)
        gudtPara.strSignURL = varTmp(0)
        If UBound(varTmp) >= 1 Then
            mbytModel = Val(varTmp(1) & "")
        Else
            mbytModel = E_RSA
        End If
        
        If mbytModel = E_RSA Then   'RSA
            Set mobjUSBKEY = CreateObject("USBKEYACTIVE.USBKeyActiveCtrl.1") '签名对象
            Set mobjMSScriptCtl = CreateObject("MSScriptControl.ScriptControl.1")
            mobjMSScriptCtl.Language = "JavaScript"
        Else                    'SM2
            Set mobjKeyManager = CreateObject("JHKey.KeyManager.1")
            Set mobjCert = CreateObject("JHKey.Cert.1")
            Set mobjKeyStore = CreateObject("JHKey.KeyStore.1")
            Set mobjKeySealArray = CreateObject("JHKey.SealArray.1")
            Set mobjKeySeal = CreateObject("JHKey.Seal.1")
            Set mobjKeyGateOper = CreateObject("JHKey.GateOper.1")
            Set mobjKeyDetector = CreateObject("JHKey.KeyDetector.1.1")
            Call mobjKeyGateOper.SetTimeout(10)
            Call mobjKeyGateOper.SetURL(gudtPara.strSignURL)
            
        End If
        
        gstrLogins = ""
        mblnInit = True
        LNCA_Initialize = True
        Exit Function
errH:
     MsgBox "创建接口部件失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function LNCA_RegCert(arrCertInfo As Variant) As Boolean
        '功能：提供在HIS数据库中注册个人证书的必要信息,用于新发放或更换证书,,需要插入USB-Key
        '返回：arrCertInfo作为数组返回证书相关信息
        '      0-ClientSignCertCN:客户端签名证书公共名称(姓名),用于注册证书时程序验证身份
        '      1-ClientSignCertDN:客户端签名证书主题(每人唯一)
        '      2-ClientSignCertSN:客户端签名证书序列号(每证书唯一)
        '      3-ClientSignCert:客户端签名证书内容
        '      4-ClientEncCert:客户端加密证书内容
        '      5-签名图片文件名,空串表示没有签名图片
        
        Dim strCertSn As String, strCertTime As String, strCertUserName As String, strCertDN As String
        Dim strSigCert As String, i As Integer, strCACert As String, lngOk As Long
        Dim strPicPath As String
        On Error GoTo errH
    
100     For i = LBound(arrCertInfo) To UBound(arrCertInfo)
102         arrCertInfo(i) = ""
        Next
    
104     If GetCertList(strCertUserName, strCertSn, strSigCert, strCertDN, strPicPath) Then
106         arrCertInfo(0) = strCertUserName
108         arrCertInfo(1) = strCertDN
110         arrCertInfo(2) = strCertSn
112         arrCertInfo(3) = strSigCert
            arrCertInfo(4) = ""
113         arrCertInfo(5) = strPicPath
            
124         LNCA_RegCert = True
        End If

        Exit Function
errH:
126     MsgBox "获取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName

End Function

Public Function LNCA_CheckCert(ByRef blnReDo As Boolean, Optional ByRef strCert As String) As Boolean
'功能：读取USB进行设备初始化并登录
    Dim strKey As String, strPIN As String, strUserName As String
    Dim strCertName As String, strCertDN As String, strPicPath As String
    Dim strCertSn As String
    Dim strCertUserID As String    '包含身份证号信息
    Dim strDate As String
    Dim udtUser As USER_INFO
    Dim strCertID As String
    Dim blnOk As Boolean
    Dim blnRet As Boolean
    
    On Error GoTo errH

     '获取证书信息同时检查Key盘是否插入
    If Not GetCertList(strCertName, strCertSn, strCert, strCertDN, strPicPath, strCertUserID, strDate) Then
        LNCA_CheckCert = False: Exit Function
    End If
    
    '未注册在当前用户名下的Key
    If mUserInfo.strUserID = "" Then
        MsgBox "您的身份证号为空,请联系管理员到人员管理中录入！", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    ElseIf UCase(strCertUserID) <> UCase(mUserInfo.strUserID) Then
        MsgBox "您的身份证号：" & _
                   vbCrLf & vbTab & "【" & mUserInfo.strUserID & "】" & vbCrLf & _
                   "当前证书唯一标识:" & _
                   vbCrLf & vbTab & "【" & strCertUserID & "】" & vbCrLf & _
                   "用户身份证号与当前证书唯一标识不相等,不能使用！", vbInformation, gstrSysName
        Exit Function
    End If
    'CA首次签名时会自动弹出密码框
    '登录验证
    If InStr(gstrLogins & "|", "|" & strCertSn & "|") > 0 Then '首次验证通过后，下次不在继续验证
        blnOk = True
    Else
        If Not GetCertLogin(strCert, strDate) Then
            blnOk = False
        Else
            blnOk = True
            If InStr(gstrLogins & "|", "|" & strCertSn & "|") = 0 Then gstrLogins = gstrLogins & "|" & strCertSn
        End If
    End If

    If blnOk Then
        '判断是否需要更新注册证书
        udtUser.strName = strCertName
        udtUser.strSignName = strCertName
        udtUser.strUserID = strCertUserID
        udtUser.strCertSn = strCertSn
        udtUser.strCertDN = strCertDN
        udtUser.strCert = strCert
        udtUser.strPicPath = strPicPath
        udtUser.strPicCode = ""

        If IsUpdateRegCert(udtUser, strDate, blnReDo) Then
            LNCA_CheckCert = True
        Else
            LNCA_CheckCert = False
        End If
    End If
    
    Exit Function
errH:
     MsgBox "检查USBKEY失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Private Function GetCertList(Optional ByRef strName As String = "-1", Optional ByRef strCertSn As String = "-1", Optional ByRef strCert As String = "-1", _
                Optional ByRef strCertDN As String = "-1", Optional strPicPath As String = "-1", _
                Optional strUserID As String = "-1", Optional strEndDate As String = "-1") As Boolean
'功能:获取证书信息
'-出参
'    strName 证书持有者姓名
'   strCertSN 证书唯一标识
'   strCert 签名证书
'   strCertDN 证书描述信息  证书注册用到
'   strPicPath 证书图片保存位置

    Dim strPic As String
    Dim strMsg As String
    Dim strEnd As String
    Dim lngRet As Long
    Dim strPIN As String
    Dim i As Integer
    
    On Error GoTo errH
    If Not LNCA_Initialize() Then Exit Function

    If mbytModel = E_RSA Then
        '输入密码
        If mstrLastPwd <> "" Then strPIN = mstrLastPwd
        If strPIN = "" Then
            If Not frmPassword.ShowMe(strPIN) Then Exit Function
        End If
        
        If strPIN = "" Then
           MsgBox "请输入证书密码！", vbOKOnly + vbInformation, gstrSysName
           Exit Function
        Else
            If mintLogin >= 8 Then
                MsgBox "已经输入了" & mintLogin & "次错误密码，超过了最大输入次数！", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
            On Error Resume Next
            lngRet = mobjUSBKEY.MNGInit(strPIN)
            If Err.Number <> 0 Then
                MsgBox "请您插入KEY盘！", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
            Err.Clear: On Error GoTo 0
            
            If lngRet = 0 Then
               mstrLastPwd = strPIN
            Else
                mintLogin = mintLogin + 1
                MsgBox "证书密码可能不正确，您已经输入了" & mintLogin & "次密码，还可以输入" & 8 - mintLogin & "次!", vbOKOnly + vbInformation, gstrSysName
                mstrLastPwd = ""
                Exit Function
            End If
        End If
        Call mobjUSBKEY.MNGLogin
    
        If strCertSn <> "-1" Then strCertSn = mobjUSBKEY.MNGGetSignCertSN '唯一标识号
        If strCert <> "-1" Then strCert = mobjUSBKEY.MNGGetSignCert()    '获取签名证书
        If strName <> "-1" Then strName = mobjUSBKEY.MNGGetSignCertCN()          ''获取名称
        If strCertDN <> "-1" Then strCertDN = mobjUSBKEY.MNGGetSignCertDN      '描述
        
        If strPicPath <> "-1" Then
            strPic = mobjUSBKEY.MNGGetSESCount '法人签名|个人章
            strPic = Split(strPic, "|")(0)
            strPic = mobjUSBKEY.MNGReadSESealByLabelEx(strPic)         '获取签章图片BASE64
            If strPic <> "" Then
                strPicPath = SaveBase64ToFile("gif", strCertSn, strPic) '返回印章图片数据的BASE64转换成图片文件并返回图片位置
            Else
                strMsg = "读取签名图片失败！"
                GoTo msgINFO
            End If
        End If
        If strUserID <> "-1" Then
            strUserID = mobjUSBKEY.MNGGetSignCertDN_OUa   '身份证号
            If strUserID <> "" And Len(strUserID) >= 18 Then
                strUserID = Right(strUserID, 18)
            Else
                strMsg = "读取身份证号码失败！"
                GoTo msgINFO
            End If
        End If
        '获取客户端证书有效期截止时间
        If strEndDate <> "-1" Then
            strEndDate = mobjUSBKEY.MNGGetSignCertEndValidityTime()
            strEndDate = CDate(Format(strEndDate, "YYYY-MM-DD HH:MM:SS"))
        End If
    ElseIf mbytModel = E_SM2 Then
        lngRet = mobjKeyDetector.EnumUKey()
        If lngRet = 0 Then
            MsgBox "请您插入KEY盘！", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        ElseIf lngRet = 1 Then
            Call mobjKeyManager.EnumKeyStore               '调用EnumKeyStore枚举出所有证书
            '默认选取SM2的key签名
            lngRet = mobjKeyManager.GetCertCount()
            For i = 0 To lngRet - 1 '调用GetCertCount得到全部证书的个数
                mobjCert.SetCert (mobjKeyManager.GetCert(i)) '轮询获得所有的证书
                If mobjCert.CertUsage = 1 Then 'CertUsage == 1:签名证书,2:加密证书。CertType == 1:RSA,2:SM2
                    Call mobjKeyManager.InitKeyStoreByIndex(i, mobjKeyStore)
                    Exit For
                End If
            Next
        Else
            Call mobjKeyManager.EnumKeyStore               '调用EnumKeyStore枚举出所有证书
            '显示证书列表框，用于让用户选取证书。
            '第一个参数，表示显示的证书类型，1：RSA，2：SM2，3：RSA和SM2都显示;
            '第二个参数表示证书用途，1：签名，2：加密，3：签名加密均可
            '0-用户选择了证书;1-未选择证书
            lngRet = mobjKeyManager.ShowCertsDlg(3, 1)
            If lngRet <> 0 Then
                strMsg = "未选择任何证书！"
                GoTo msgINFO
            End If
            Call mobjKeyManager.GetSelectedCert(mobjCert)    '根据用户的选择得到指定的证书
            '根据用户选择的证书，初始化UKey操作类KeyStore 签名时直接调用签名接口
            Call mobjKeyManager.InitKeyStore(mobjKeyStore)
            'mobjCert.CertType证书类型，1:RSA，2:SM2
        End If
        strName = mobjCert.CertCN
        If strCertSn <> "-1" Then strCertSn = mobjCert.CertSN
        If strCertDN <> "-1" Then strCertDN = mobjCert.CertSubject
        If strCert <> "-1" Then strCert = mobjCert.Body
        If strUserID <> "-1" Then strUserID = Mid(mobjCert.CertOuA, 2)
        If strEndDate <> "-1" Then
            strEndDate = mobjCert.CertNotAfter
        End If

        If mobjCert.CertType = 2 Then  'SM2 允许缓存密码
            '输入密码
            If mstrLastPwd <> "" Then strPIN = mstrLastPwd
            If strPIN = "" Then
                If Not frmPassword.ShowMe(strPIN) Then Exit Function
            End If
                
            If strPIN = "" Then
               MsgBox "请输入证书密码！", vbOKOnly + vbInformation, gstrSysName
               Exit Function
            Else
                Call mobjKeyStore.SetWorkPin(strPIN)  '添加默认PING码，使PIN码框不再弹出，达到静默操作的目的（只对SM2Key有效。RSA的PIN码框由各驱动厂家各自实现，无法控制）
                lngRet = mobjKeyStore.SignData("123")    '0-成功;非0-失败
                If lngRet = 0 Then
                    mstrLastPwd = strPIN
                Else
                    mstrLastPwd = ""
                    Exit Function
                End If
            End If
        End If
        
        '取印章：根据所选证书，查询到证书所在Key里的所有印章，存入印章数组SealArray
        If strPicPath <> "-1" Then
            Call mobjKeyManager.InitSealStore(mobjKeySealArray)
            lngRet = mobjKeySealArray.GetSealCount()         '得到印章数组所存印章个数
            If lngRet = 0 Then
                strMsg = "所选证书无对应的印章！"
                GoTo msgINFO
            End If
            For i = 0 To lngRet - 1
                Call mobjKeySealArray.GetSeal(i, mobjKeySeal)     '从印章数组中取得印章
                strPic = mobjKeySeal.getpic()              '得到印章图片的base64数据
                If strPic <> "" Then
                    strPicPath = SaveBase64ToFile("gif", strCertSn, strPic) '返回印章图片数据的BASE64转换成图片文件并返回图片位置
                    Exit For
                Else
                    strMsg = "读取签名图片失败！"
                    GoTo msgINFO
                End If
            Next
        End If
    End If
    GetCertList = True
    Exit Function

msgINFO:
    If strMsg <> "" Then
        MsgBox strMsg, vbInformation, gstrSysName
        Exit Function
    End If
errH:

500  MsgBox "获取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Private Function GetCertLogin(ByVal strSignCert As String, ByVal strEnd As String) As Boolean
    Dim blnRet As Boolean
    Dim strSignResult As String
    Dim strToken As String, strParameter As String, strRandom As String, strMsg As String
    Dim strRet As String
    Dim datEnd As Date
    Dim intDay As Integer
    Dim lngRet As Long
    
    On Error GoTo errH
    If mbytModel = E_RSA Then
        '获取随机数
        strRandom = HttpPost(gudtPara.strSignURL, "cmd=getrand", responseText)   '获取随机数返回值: {"ret":1,"errinfo":"","rand":"3XO9JCXVJ6LF05M51165"}
        strRandom = GetSubString(strRandom, "rand")
        
        '随机数签名
        strSignResult = mobjUSBKEY.MNGSignData(strRandom, Len(strRandom))      '控件签名
        If strSignResult = "" Then
            MsgBox "随机数签名失败！", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    
        '随机数签名验证
        strParameter = "cmd=sm2certlogin" & "&rand=" & EnCodeURL(strRandom) & "&cert=" & EnCodeURL(strSignCert) & "&signed=" & EnCodeURL(strSignResult)  '服务器验证KEY 签名结果
        strToken = HttpPost(gudtPara.strSignURL, strParameter, responseText)  '
        strRet = GetSubString(strToken, "ret")
        If strRet <> "1" Then
            MsgBox "证书登录验证失败！", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    ElseIf mbytModel = E_SM2 Then
        'Call mobjKeyStore.SetWorkPin(mstrLastPwd)
        '       重复设置密码导致登陆验证返回值=9 提示失败 故注释
        lngRet = mobjKeyGateOper.ReqCertLogin(mobjKeyStore, "123")  '证书登陆
        If lngRet <> 0 Then
            MsgBox "证书登录验证失败！", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
    '获取客户端证书有效期截止时间
    datEnd = CDate(strEnd)
    '验证客户端证书有效期剩余天数
    intDay = Int(CDbl(datEnd) - CDbl(Now))
    
    If (intDay <= 30 And intDay > 0 And Not gblnShow) Then
        MsgBox "您的证书还有" & intDay & "天过期。", vbInformation + vbOKOnly, gstrSysName
        gblnShow = True
        GetCertLogin = True
    ElseIf (intDay <= 0) Then
        MsgBox "您的证书已过期 " & Abs(intDay) & " 天。", vbInformation + vbOKOnly, gstrSysName
        GetCertLogin = False
    End If
        
    GetCertLogin = True
    Exit Function
errH:
    MsgBox "登录服务器验证失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Function LNCA_Sign(ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String, _
        ByRef strTimeStampCode As String, ByRef blnReDo As Boolean) As Boolean
'签名
        Dim strParameter As String, strMsg As String, strDate As String, strRet As String
        Dim strSignCert As String
        Dim intRet As Integer
        Dim blnCheck As Boolean, blnRet As Boolean
        Dim datTime As Date
        
        On Error GoTo errH
        blnCheck = LNCA_CheckCert(blnReDo, strSignCert)
        If blnReDo Then Exit Function
1100    If blnCheck Then
            If mbytModel = E_RSA Then
                '验证当前USB是否是签名用户的，并获取签名证书
                strSource = EncodeBase64String(strSource) '源文中包含特殊字符串需要加密转换
                strSignData = mobjUSBKEY.MNGSignData(strSource, Len(strSource))        '控件签名
                If strSignData <> "" Then
                    '存储网关服务器
                    datTime = gobjComLib.zlDatabase.Currentdate()
                    strDate = Format(datTime, "yyyyMMddhhmmss")
                    strTimeStamp = Format(datTime, "yyyy-MM-dd HH:mm:ss")
                    strParameter = "cmd=insert_sign_record" & "&appid=" & EnCodeURL("100") & "&docid=" & EnCodeURL("100") & "&docname=" & _
                    EnCodeURL("ZLHIS") & "&textinfo=" & EnCodeURL(strSource) & "&signdata=" & EnCodeURL(strSignData) & "&signcert=" & _
                    EnCodeURL(strSignCert) & "&signdate=" & EnCodeURL(strDate)
                    strRet = HttpPost(gudtPara.strSignURL, strParameter, responseText)
                    blnRet = GetSubString(strRet, "ret") = "1"
                    If Not blnRet Then strMsg = "签名失败！"
                Else
                    strMsg = "签名失败！"
                    blnRet = False
1112            End If
            ElseIf mbytModel = E_SM2 Then
                'mobjKeyStore在LNCA_CheckCert 已经实例化
                Call mobjKeyStore.SetWorkPin(mstrLastPwd)
                intRet = mobjKeyStore.SignData(strSource)   '0-成功;非0-失败
                If intRet = 0 Then
                    strSignData = mobjKeyStore.GetSignData()               '得到签名数据
                Else
                    strMsg = "签名失败！"
                    blnRet = False
                End If
                '签名后存储网关服务器
                datTime = gobjComLib.zlDatabase.Currentdate()
                strDate = Format(datTime, "yyyyMMddhhmmss")
                strTimeStamp = Format(datTime, "yyyy-MM-dd HH:mm:ss")
                intRet = mobjKeyGateOper.ReqUploadMedRecord("01", "appid", "docid", "docname", strSource, strSignData, strSignCert, strDate)
                blnRet = (intRet = 0)
                If Not blnRet Then strMsg = "签名失败！"
            End If
        Else
            strMsg = "签名失败！"
            blnRet = False
        End If
        If strMsg <> "" Then
            MsgBox strMsg, vbInformation, gstrSysName
        End If
                
        LNCA_Sign = blnRet
        Exit Function
errH:
114     MsgBox "签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function LNCA_VerifySign(ByVal strCert As String, ByVal strSignData As String, ByVal strSource As String) As Boolean
'验证签名
'
    Dim strParameter As String
    Dim strRet As String
    Dim blnRet As Boolean
    Dim strMsg As String
    Dim lngRet As Long
    
    On Error GoTo errH
    If mbytModel = E_RSA Then
        strSource = EncodeBase64String(strSource)
        strCert = strCert & "123"
        strParameter = "cmd=verifysm2" & "&text=" & EnCodeURL(strSource) & "&cert=" & EnCodeURL(strCert) & "&signed=" & EnCodeURL(strSignData)
        strRet = HttpPost(gudtPara.strSignURL, strParameter, responseText)
        blnRet = GetSubString(strRet, "ret") = "1"    '返回值=1验证签名成功
    ElseIf mbytModel = E_SM2 Then
        '参数 签名结果;签名原文;预留;服务器证书
        lngRet = mobjKeyGateOper.ReqVerifySig(strSignData, strSource, 0, strCert)
        blnRet = lngRet = 0
    End If
    If blnRet Then    '验证签名失败
        strMsg = "验证成功，该电子签名数据有效！"
    Else
        strMsg = "验签失败！"
    End If
    
    If strMsg <> "" Then
        MsgBox strMsg, vbInformation, gstrSysName
    End If
        
    LNCA_VerifySign = blnRet
    
    Exit Function
errH:
104     MsgBox "验证签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Public Sub LNCA_UnLoadObj()
    If mbytModel = E_RSA Then
        Set mobjUSBKEY = Nothing
        Set mobjMSScriptCtl = Nothing
    Else
        Set mobjCert = Nothing
        Set mobjKeyGateOper = Nothing
        Set mobjKeyManager = Nothing
        Set mobjKeySeal = Nothing
        Set mobjKeySealArray = Nothing
        Set mobjKeyStore = Nothing
        Set mobjKeyDetector = Nothing
    End If
    
    mblnInit = False
End Sub

Private Function EnCodeURL(ByVal strUrl As String) As String
'功能:将传人字符串按UTF编码方式转换成十六进制的转义序列
'说明：encodeURI-javaScript方法不会对 ASCII 字母和数字进行编码，也不会对这些 ASCII 标点符号进行编码： - _ . ! ~ * ' ( )
    Dim i As Long
    Dim strChar As String
    Dim intAsc As Integer
    Dim strRet As String
    
    For i = 1 To Len(strUrl)
        strChar = Mid(strUrl, i, 1)
        intAsc = Asc(strChar)
        If intAsc >= 0 And intAsc <= 127 Then
           strChar = "%" & Hex(intAsc)
        Else
            strChar = mobjMSScriptCtl.Eval("encodeURI(""" & strChar & """)")
        End If
        strRet = strRet & strChar
    Next
    
    EnCodeURL = strRet
End Function

Public Function LNCA_GetPara() As Boolean
'设置服务器地址
    Dim arrTmp As Variant
    
    On Error GoTo errH
    gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys)
    If gstrPara = "" Then gstrPara = "http://218.25.86.214:2010/ssoworker"
    arrTmp = Split(gstrPara, G_STR_SPLIT)
    If UBound(arrTmp) > 0 Then
        gudtPara.strSignURL = arrTmp(0)
        gudtPara.bytSignVersion = Val(arrTmp(1) & "")
    Else
        gudtPara.strSignURL = arrTmp(0)  '
    End If

    Exit Function
errH:
    MsgBox "读取参数失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Function LNCA_SetParaStr() As String
    LNCA_SetParaStr = gudtPara.strSignURL & G_STR_SPLIT & gudtPara.bytSignVersion
End Function

Private Function GetSubString(ByVal strSource As String, ByVal strNode As String) As String
'功能:获取返回字符串中某个节点值
'参数:strSource -传人字符串 {"ret":1,"errinfo":"","rand":"3XO9JCXVJ6LF05M51165"}
'    strNode-标识要获取的节点名称
    Dim arrMain As Variant
    Dim arrSub As Variant
    Dim strRet As String
    Dim i As Long
    
    arrMain = Split(strSource, ",")
    For i = LBound(arrMain) To UBound(arrMain)
        Select Case UCase(strNode)
        Case UCase("rand"), UCase("token")
            If InStr(UCase(arrMain(i)), UCase(strNode)) > 0 Then
                arrSub = Split(arrMain(i), ":")
                strRet = Mid(arrSub(1), 2)
                strRet = left(strRet, Len(strRet) - 2)
                Exit For
            End If
        Case UCase("ret")
            If InStr(UCase(arrMain(i)), UCase(strNode)) > 0 Then
                arrSub = Split(arrMain(i), ":")
                strRet = arrSub(1)
                Exit For
            End If
        Case UCase$("errinfo")
            If InStr(UCase(arrMain(i)), UCase(strNode)) > 0 Then
                arrSub = Split(arrMain(i), ":")
                strRet = Mid(arrSub(1), 2)
                strRet = left(strRet, Len(strRet) - 1)
                Exit For
            End If
        End Select
    Next
    GetSubString = strRet
End Function





