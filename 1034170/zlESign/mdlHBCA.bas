Attribute VB_Name = "mdlHBCA"
Option Explicit

'河北邯郸市第三医院   河北CA
Private mblnInit As Boolean         '是否已初始化成功
Private mCertMgr As Object          'HebcaP11XLib.certMgr
Private mSignCert As Object         'HebcaP11XLib.cert
Private mFormSeal As Object         'FormSealCtrl1 电子签章控件
Private mSVSClient As Object        'SVS_SOFT_COMLib.SvsVerify '定义并实例化SVS客户端组件
Private mblnTs As Boolean           '是否启用时间戳

Private Const M_STR_LICENCE As String = "amViY55oZWKcZmhlnWxhaGViY2GXGmJjYWhlYnGH1QQ5GcNqnW6z3vohVnE+nTJr"
Private Const M_STR_SUMMARY As String = "[SUMMARY]"

Public Function HBCA_InitObject() As Boolean
'----------------------------------------------------------------------------------------------------------------------------------
'功能:初始化电子签名需要用到对象
'返回:True-初始化成功;False-初始化失败
'编制:余伟节
'日期:2015-08-31
'----------------------------------------------------------------------------------------------------------------------------------
    Dim arrList As Variant
    
    If mblnInit Then HBCA_InitObject = True: Exit Function
    On Error GoTo errH
    '参数信息:IP|端口号|是否启用时间戳(0-不启用/1-启用)
1   gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys, , "121.28.49.158&&&5000&&&0")  '读取配置内容
    If gstrPara = "" Then
        Err.Raise -1, , "配置文件读取失败，请到启用电子签名接口处设置。"
        Exit Function
    End If
    arrList = Split(gstrPara, "&&&")
    If UBound(arrList) <> 2 Then
        MsgBox "签名服务器地址配置格式有误,请到启用电子签名接口处设置。", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
99  mblnTs = (Val(CStr(Split(gstrPara, G_STR_SPLIT)(2))) = 1)
100 If mCertMgr Is Nothing Then Set mCertMgr = CreateObject("HebcaP11X.CertMgr.1")
110 If mSignCert Is Nothing Then Set mSignCert = CreateObject("HebcaP11X.Cert.1")
120 If mFormSeal Is Nothing Then Set mFormSeal = CreateObject("HebcaFormSeal.FormSealCtrl.1")
130 If mSVSClient Is Nothing Then Set mSVSClient = CreateObject("Svs_soft_com.SvsVerify.1")
    gstrLogins = ""
    mblnInit = True
    HBCA_InitObject = True
    Exit Function
errH:
500 MsgBox "创建河北CA接口部件失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & vbNewLine & _
            Err.Description, vbQuestion, gstrSysName
End Function

Public Function HBCA_RegCert(arrCertInfo As Variant) As Boolean
'----------------------------------------------------------------------------------------------------------------------------------
'功能:提供在HIS数据库中注册个人证书的必要信息,用于新发放或更换证书,需要插入USB-Key
'返回：arrCertInfo作为数组返回证书相关信息
'      0-ClientSignCertCN:客户端签名证书公共名称(姓名),用于注册证书时程序验证身份
'      1-ClientSignCertDN:客户端签名证书主题(每人唯一)
'      2-ClientSignCertSN:客户端签名证书序列号(每证书唯一)
'      3-ClientSignCert:客户端签名证书内容
'      4-ClientEncCert:客户端加密证书内容
'      5-签名图片文件名,空串表示没有签名图片
'      6-时间戳证书
'      7-签章信息
'编制:余伟节
'日期:2015-08-31
'----------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strCertName As String
    Dim strCertDN As String, strUserID As String, strSealB64 As String
    Dim strCertSn As String, strTSCert As String
    Dim strSignCert As String, strPic As String
    Dim strEncCert As String
    
    On Error GoTo errH
    
    If mFormSeal Is Nothing Then Set mFormSeal = CreateObject("HebcaFormSeal.FormSealCtrl.1")
    
100     For i = LBound(arrCertInfo) To UBound(arrCertInfo)
102         arrCertInfo(i) = ""
        Next

104     If GetCertList(strCertName, strCertSn, strSignCert, strUserID, strSealB64, strTSCert, strPic) Then
106         arrCertInfo(0) = strCertName
108         arrCertInfo(1) = strCertDN
110         arrCertInfo(2) = strCertSn
112         arrCertInfo(3) = strSignCert
113         arrCertInfo(4) = strEncCert
            arrCertInfo(5) = strPic
            arrCertInfo(6) = strTSCert
            arrCertInfo(7) = strSealB64
114
124         HBCA_RegCert = True
        End If

        Exit Function
errH:
126     MsgBox "获取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
   
        
End Function

Private Function GetCertList(ByRef strName As String, ByRef strCertSn As String, ByRef strSignCert As String, _
                ByRef strUserID As String, Optional ByRef strSealBase64 As String, Optional ByRef strTSCert As String, _
                Optional ByRef strPicFile As String, Optional ByRef strPic As String) As Boolean
'----------------------------------------------------------------------------------------------------------------------------------
'功能:获取证书信息
'参数:strName-证书用户名
'     strCertSn-证书序列号
'     strSignCert-证书内容
'     strUserID-用户唯一标识  截取身份证号
'     strSealBase64-签章BASE64
'     strTSCert-时间戳证书BASE64
'返回：
'编制:余伟节
'日期:2015-08-31
'----------------------------------------------------------------------------------------------------------------------------------
    Dim intCount As Integer
    Dim strSignData As String
    Dim strSealSn As String
    Dim objPic As Picture
    
    intCount = mFormSeal.GetSealCount()
    
    If intCount < 1 Then
        MsgBox "请您插入Key！", vbInformation, gstrSysName
        Exit Function
    Else
        mCertMgr.Licence = M_STR_LICENCE
        Set mSignCert = mCertMgr.SelectSignCert
        'CN=持有者姓名
        strName = mSignCert.GetSubjectItem("cn")
        'strSignCert = mSignCert.GetCertB64        '得到签名证书内容
'        strCertDN = mSignCert.GetSubjectItem("DN")
        '获取数字证书的唯一标识,用于和用户建立绑定
        strUserID = mSignCert.GetCertExtensionByOid("1.2.156.112586.1.4") '"2@6021SF0130637201507090001"
        strUserID = Mid(strUserID, 10) '获取身份证号
        '获取证书信息  验证签名需要证书信息：
        '签章BASE64编码,签章证书,时间戳证书BASE64编码
        strSignData = mFormSeal.SignAndSealWithoutTimeStampCert("测试20150901", "", 0, True, mblnTs)
        If strSignData = "" Then MsgBox "随机数签名失败！", vbInformation, gstrSysName: Exit Function
        '获取签章的SN\签章Bases64
        strSealSn = mFormSeal.GetSelectedSeal()
        strSealBase64 = mFormSeal.GetSeal(strSealSn) '获取章的B64
        strPic = mFormSeal.GetSealPicFromB64(strSealBase64)
        
        strPicFile = SaveBase64ToFile("gif", strSealSn, strPic)
        '将图片转换成指定bmp格式
'        Set objPic = LoadPicture(strPicFile)
'        SavePicture objPic, strPicFile
'
        '获取证书各项信息，证书的SN和证书的有效期需要存入数据库
        strSignCert = mFormSeal.GetCert(strSealSn) '获取证书
        Set mSignCert = mCertMgr.CreateCertFromB64(strSignCert)
        strCertSn = mSignCert.GetSerialNumber '证书SN
        'dCertDate:=mSignCert.NotAfter    有效期
         '时间戳
        If mblnTs Then
            strTSCert = mFormSeal.GetTimeStampCert '获取时间戳证书内容
        End If
    End If
        
    GetCertList = True
    Exit Function
End Function

Public Function HBCA_Sign( _
    ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String, _
    ByRef strTimeStampCode As String, ByRef blnReDo As Boolean) As Boolean
'----------------------------------------------------------------------------------------------------------------------------------
'功能:签名
'参数:strSource-数据源
'     strSignData-签名值
'     strTimeStamp-时间戳值
'     strTimeStampCode-时间戳信息
'返回：True/false
'编制:余伟节
'日期:2015-08-31
'----------------------------------------------------------------------------------------------------------------------------------
        '签名
        Dim strTiemRequest As String
        Dim strTmp As String
        Dim strSealSn As String
        Dim strSealBase64 As String
        Dim blnCheck As Boolean
        
        On Error GoTo errH
        blnCheck = HBCA_CheckCert(blnReDo)
        If blnReDo Then Exit Function
100     If blnCheck Then                '验证当前USB是否是签名用户的，并获取签名证书
            '调用SignAndSealWithoutTimeStampCert对原文进行盖章，可以对原文数据进行组织，
104         strSource = mCertMgr.util.HashText(strSource, 1)
120         strSignData = mFormSeal.SignAndSealWithoutTimeStampCert(strSource, "", 0, True, mblnTs)
121         If strSignData = "" Then MsgBox "签名失败,签名值为空！", vbInformation, gstrSysName: Exit Function
            If mblnTs Then
130             strTimeStampCode = mFormSeal.GetTimeStamp() '获取时间戳信息
140             strTimeStamp = mFormSeal.GetTimeStampInfoByB64(strTimeStampCode, "time")
                If strTimeStampCode = "" Then
                    MsgBox "签名失败,时间戳B64为空！", vbInformation, gstrSysName
                    Exit Function
                End If
            Else
150             strTimeStamp = CStr(gobjComLib.zlDatabase.Currentdate)
            End If
        Else
            MsgBox "签名失败！", vbInformation, gstrSysName
            Exit Function
        End If
        If Trim(strSignData) <> "" Then strSignData = M_STR_SUMMARY & strSignData  '此标识[SUMMARY]用于验证签名时区分按原文验证签名还是按摘要验证签名
        HBCA_Sign = True
        Exit Function
errH:
114     MsgBox "签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function HBCA_VerifySign(ByVal strSignData As String, ByVal strSource As String, ByVal strTimeStampCode As String, _
                            ByVal strCert As String, ByVal strTSCert As String, ByVal strSealCert As String) As Boolean
'----------------------------------------------------------------------------------------------------------------------------------
'功能:验证签名
'参数:
'   strSignData-签名值
'   strSource-源文
'   strTimeStampCode-时间戳信息
'   strCert-证书内容
'   strTSCert-时间戳证书内容
'   strSealCert-签章证书内容
'返回：True/false
'编制:余伟节
'日期:2015-08-31
'----------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    Dim lngRet As Long
    Dim blnOk As Boolean
    On Error GoTo errH
    
    If mFormSeal Is Nothing Then Set mFormSeal = CreateObject("HebcaFormSeal.FormSealCtrl.1")
    strTmp = ""

    '电子签章验证
    If UCase(left(strSignData, Len(M_STR_SUMMARY))) = M_STR_SUMMARY Then
        '按摘要签名时验证签名时需要取摘要
        strSignData = Mid(strSignData, Len(M_STR_SUMMARY) + 1)
        mCertMgr.Licence = M_STR_LICENCE
        strSource = mCertMgr.util.HashText(strSource, 1)
    End If

1    Call mFormSeal.VerifyAndShowSeal(strSealCert, strCert, strSource, 0, strSignData, IIf(mblnTs, 0, -1), strTimeStampCode, strTSCert, 0)
2    lngRet = mFormSeal.GetVerifyResult()

    If lngRet = 0 Then
        strTmp = "签章验证成功！"
        blnOk = True
    Else
        strTmp = "签章验证失败！"
        blnOk = False
    End If

    If strTmp <> "" Then
        MsgBox strTmp, vbOKOnly + vbInformation, gstrSysName
    End If
    
10    HBCA_VerifySign = blnOk
    Exit Function
errH:
104     MsgBox "验证签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Function HBCA_CheckCert(ByRef blnReDo As Boolean) As Boolean
'----------------------------------------------------------------------------------------------------------------------------------
'功能：读取USB进行设备初始化并登录
'参数:
'   出参:blnRedo-True：重新注册证书成功,False-未重新注册证书
'返回：True/false
'编制:余伟节
'日期:2015-08-31
'----------------------------------------------------------------------------------------------------------------------------------
        
        Dim strCertUserID As String, strPIN As String, strUserName As String
        Dim strWebUrl As String, intDate   As Integer
        Dim strCertName As String, strCertSn As String, strCert As String, strCertDN As String
        Dim strTSCert As String, strSealCode As String, strPic As String, strDate As String
        Dim blnOk As Boolean
        Dim udtUser As USER_INFO
        
        On Error GoTo errH
100     If Not mblnInit Then
            Call HBCA_InitObject
            If Not mblnInit Then
102             MsgBox "部件未初始化！"
                Exit Function
            End If
        End If
         '获取证书信息同时检查Key盘是否插入
        If Not GetCertList(strCertName, strCertSn, strCert, strCertUserID, strSealCode, strTSCert, , strPic) Then
            HBCA_CheckCert = False: Exit Function
        End If
        '未注册在当前用户名下的Key
        If mUserInfo.strUserID = "" Then
            MsgBox "您的身份证号为空,请联系管理员到人员管理中录入！", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        ElseIf strCertUserID <> mUserInfo.strUserID Then
            MsgBox "您的身份证号：" & _
                   vbCrLf & vbTab & "【" & mUserInfo.strUserID & "】" & vbCrLf & _
                   "当前证书唯一标识:" & _
                   vbCrLf & vbTab & "【" & strCertUserID & "】" & vbCrLf & _
                   "用户身份证号与当前证书唯一标识不相等,不能使用！", vbInformation, gstrSysName
            Exit Function
        End If
        
        '登录验证
        If InStr(gstrLogins & "|", "|" & strCertSn & "|") > 0 Then '首次验证通过后，下次不在继续验证
            blnOk = True
        Else
            If Not GetCertLogin() Then
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
            udtUser.strEncCert = ""
            udtUser.strCertID = ""
            udtUser.strPicCode = strPic
            udtUser.strTSCert = strTSCert
            udtUser.strSealCode = strSealCode
            '获取已经注册证书的有效结束日期
                '获取证书各项信息，证书的SN和证书的有效期需要存入数据库
            Set mSignCert = mCertMgr.CreateCertFromB64(mUserInfo.strCert)
            strDate = mSignCert.NotAfter
            If IsUpdateRegCert(udtUser, strDate, blnReDo) Then
                HBCA_CheckCert = True
            Else
                HBCA_CheckCert = False
            End If
        End If
    
     
    
        Exit Function
errH:
124     MsgBox "检查USBKEY失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Private Function GetCertLogin() As Boolean
'----------------------------------------------------------------------------------------------------------------------------------
'功能：登录验证
'参数:
'返回：True/false
'编制:余伟节
'日期:2015-08-31
'----------------------------------------------------------------------------------------------------------------------------------
    Dim strText As String
    Dim strMsg As String
    Dim lngRetVal As Long
    Dim strSignData As String
    Dim strCertB64 As String
    Dim strSealSn As String
    Dim strDate As String
    Dim intDay As Integer
    Dim strIP As String, lngPort As Long
    Dim arrTmp As Variant
    
    On Error GoTo errH
    If mCertMgr Is Nothing Then Set mCertMgr = CreateObject("HebcaP11X.CertMgr.1") '实例化P11组件的CertMgr类
    If mSVSClient Is Nothing Then Set mSVSClient = CreateObject("Svs_soft_com.SvsVerify.1")
    If mFormSeal Is Nothing Then Set mFormSeal = CreateObject("HebcaFormSeal.FormSealCtrl.1")
    strText = "hebca2013" '原始字符串

    mCertMgr.Licence = M_STR_LICENCE

    Set mSignCert = mCertMgr.SelectSignCert  '得到签名证书对象
    strSignData = mSignCert.SignText(strText, 1)   '进行数字签名,将签名值存放到signdata
    strCertB64 = mSignCert.GetCertB64         '得到签名证书内容
    'gstrPara = "121.28.49.158&&&5000"
    arrTmp = Split(gstrPara, G_STR_SPLIT)     'IP&&&端口号 "121.28.49.158", 5000
    strIP = arrTmp(0): lngPort = Val(CStr(arrTmp(1)))
    lngRetVal = mSVSClient.InitialVerify(strIP, lngPort) '初始化SVS客户端

    Dim r As Boolean
    
    If lngRetVal < 0 Then
        MsgBox "无法连接SVS服务器!", vbInformation, gstrSysName
        Exit Function
    End If
    
    lngRetVal = mSVSClient.VerifyCertSign(-1, 0, strText, Len(strText), strCertB64, strSignData, 1, lngRetVal)     '验证
    Select Case lngRetVal
        Case 0
            strMsg = "验证成功"
        Case 1
            strMsg = "您的证书未生效!"
        Case 2
            strMsg = "您的证书已经过期!"
        Case 4
            strMsg = "您的证书非河北CA颁发!"
        Case 1002
            strMsg = "您的证书非河北CA颁发!"
        Case 7
            strMsg = "您的证书已经被吊销!"
        Case -6406
            strMsg = "签名验证失败,请重试!"
    End Select
    If strMsg <> "验证成功" Then
        MsgBox "错误信息:" & strMsg, vbInformation, gstrSysName
        Exit Function
    End If
'
'     strSignData = mFormSeal.SignAndSealWithoutTimeStampCert("测试20150901", "", 0, True, True)
'
'     '获取签章的SN\签章Bases64
'     strSealSn = mFormSeal.GetSelectedSeal()
'     strSignCert = mFormSeal.GetCert(strSealSn) '获取证书
'     Set mSignCert = mCertMgr.CreateCertFromB64(strSignCert)
     strDate = mSignCert.NotAfter  '有效期
    If strDate <> "" Then
    '验证客户端证书有效期剩余天数
        intDay = CheckValidaty(CDate(strDate))
    
        If (intDay <= 30 And intDay > 0 And Not gblnShow) Then
            MsgBox "您的证书还有" & intDay & "天过期。", vbOKOnly + vbInformation, gstrSysName
            gblnShow = True
        ElseIf (intDay <= 0) Then
            MsgBox "您的证书已过期 " & Abs(intDay) & " 天。", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    GetCertLogin = True
    Exit Function
errH:
    MsgBox "登录验证失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Function HBCA_GetPara() As Boolean
'设置服务器地址
    Dim arrList As Variant
    
    On Error GoTo errH
    gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys)
    If gstrPara = "" Then gstrPara = "121.28.49.158&&&5000&&&0" '参数信息:IP&&&端口号&&&是否启用时间戳(0-不启用/1-启用)
    If gstrPara <> "" Then
        arrList = Split(gstrPara, "&&&")
        If UBound(arrList) = 2 Then
             gudtPara.strSIGNIP = Trim(arrList(0))
             gudtPara.strSignPort = Trim(arrList(1))
             gudtPara.blnISTS = (Val(arrList(2)) = 1)
        End If
    End If
    Exit Function
errH:
    MsgBox "读取参数失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Function HBCA_SetParaStr() As String
    HBCA_SetParaStr = gudtPara.strSIGNIP & "&&&" & gudtPara.strSignPort & "&&&" & IIf(gudtPara.blnISTS, "1", "0")
End Function

Public Sub HBCA_UnloadObj()
'----------------------------------------------------------------------------------------------------------------------------------
'功能:卸载对象
'返回:无
'编制:余伟节
'日期:2015-08-31
'----------------------------------------------------------------------------------------------------------------------------------
    Set mCertMgr = Nothing
    Set mSVSClient = Nothing
    Set mSignCert = Nothing
    Set mFormSeal = Nothing
    mblnInit = False
End Sub


