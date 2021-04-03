Attribute VB_Name = "mdlSHCA"
Option Explicit
'上海CA中心功能模块
Private mblnInit As Boolean         '是否已初始化成功
Private mLastPWD As String          '缓存输入的密码

Private SHCA_Client As Object       '证书部件
Private mLogin As Long              '输入密码错误次数
Public Enum SH_Version
    V_SEH = 0
    V_ESE = 1
End Enum

Public Function SHCA_InitObj() As Boolean
    '证书部件初始化
        Dim progID As String
        
        On Error GoTo errH
102     SHCA_InitObj = mblnInit
104     If mblnInit Then Exit Function
        mLastPWD = ""
        If Not SHCA_GetPara(1) Then Exit Function
108     Set SHCA_Client = CreateObject("SafeEngineCOM.SafeEngineCtl")
        If gudtPara.bytSignVersion = V_SEH Then
            Call SHCA_Client.SEH_InitialSession(2, "", "", 0, 2, "", "") '初始化CA接口
                Else
            Call SHCA_Client.ESE_InitialSession(2, "", "", 0, 2, "", "") '初始化CA接口
        End If
        If SHCA_Client.errorCode <> 0 Then
            GoTo errH
        End If
114     SHCA_InitObj = True
    
116     mblnInit = SHCA_InitObj
        mLogin = 0
        Exit Function
errH:
118     MsgBox "创建接口部件失败！" & vbNewLine & Err.Description, vbQuestion, gstrSysName
    
End Function

Public Function SHCA_RegCert(arrCertInfo As Variant) As Boolean
        '功能：提供在HIS数据库中注册个人证书的必要信息,用于新发放或更换证书,,需要插入USB-Key
        '返回：arrCertInfo作为数组返回证书相关信息
        '      0-ClientSignCertCN:客户端签名证书公共名称(姓名),用于注册证书时程序验证身份
        '      1-ClientSignCertDN:客户端签名证书主题(每人唯一)
        '      2-ClientSignCertSN:客户端签名证书序列号(每证书唯一)
        '      3-ClientSignCert:客户端签名证书内容
        '      4-ClientEncCert:客户端加密证书内容
        '      5-签名图片文件名,空串表示没有签名图片
        
        Dim strKeyId As String, strCertTime As String, strCertUserName As String, strCertDN As String, strCertSn As String
        Dim strSigCert As String, i As Integer, strCACert As String, lngOk As Long
        Dim strPicData As String
        On Error GoTo errH
    
100     For i = LBound(arrCertInfo) To UBound(arrCertInfo)
102         arrCertInfo(i) = ""
        Next
    
104     If GetCertList(strCertUserName, strKeyId, strSigCert, strCertSn) Then
106         arrCertInfo(0) = strCertUserName
108         arrCertInfo(1) = GetCertDN(strSigCert)
110         arrCertInfo(2) = strCertSn
112         arrCertInfo(3) = strSigCert
124         SHCA_RegCert = True
        End If

        Exit Function
errH:
126     MsgBox "获取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName

End Function
Public Function GetCertDN(strCert As String) As String
    Dim strCertDN As String
    Dim strCN As String, strO As String, strOU As String, strS As String, strL As String, strC As String, strE As String
    strC = SHCA_Client.SEH_GetCertDetail(strCert, 13)
    strO = SHCA_Client.SEH_GetCertDetail(strCert, 14)
    strOU = SHCA_Client.SEH_GetCertDetail(strCert, 15)
    strS = SHCA_Client.SEH_GetCertDetail(strCert, 16)
    strCN = SHCA_Client.SEH_GetCertDetail(strCert, 17)
    strL = SHCA_Client.SEH_GetCertDetail(strCert, 18)
    strE = SHCA_Client.SEH_GetCertDetail(strCert, 19)
    strCertDN = IIf(strS = "", "", "S=" & strS & ",") & IIf(strL = "", "", "L=" & strL & ",") & IIf(strO = "", "", "O=" & strO & ",") _
    & IIf(strOU = "", "", "OU=" & strOU & ",") & IIf(strCN = "", "", "CN=" & strCN & ",") & IIf(strE = "", "", "E=" & strE)
    GetCertDN = strCertDN
End Function
Public Function SHCA_Sign(ByVal strCurrCertSn As String, ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String, ByRef blnReDo As Boolean) As Boolean
        '签名
        Dim strSigCert As String
        Dim blnCheck As Boolean

        On Error GoTo errH
100     blnCheck = SHCA_CheckCert("", "", blnReDo)
        If blnReDo Then Exit Function
        If blnCheck Then
            '证书ID进行签名
            If gudtPara.bytSignVersion = V_SEH Then
                strSignData = SHCA_Client.SEH_SignData(strSource, 3)
            Else
                strSignData = SHCA_Client.ESE_SignData(strSource, "")
            End If
            If strSignData <> "" And SHCA_Client.errorCode = 0 Then
                 SHCA_Sign = True
            Else
                MsgBox "签名失败！" & ValidateCertView(SHCA_Client.errorCode)
            End If
        Else
            If mLastPWD = "" Then
                Exit Function
            Else
                MsgBox "签名失败！", vbInformation, "电子签名部件"
            End If
        End If
        Exit Function
errH:
114     MsgBox "签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function SHCA_VerifySign(ByVal strCurrCertSn As String, ByVal strSignData As String, ByVal strSource As String, ByVal strSignCert As String) As Boolean
        '验证签名
        Dim strTmp As String
        On Error GoTo errH
        If gudtPara.bytSignVersion = V_SEH Then
            Call SHCA_Client.SEH_InitialSession(2, "", "", 0, 2, "", "") '初始化CA接口
            Call SHCA_Client.SEH_VerifySignData(strSource, 3, strSignData, strSignCert)
        Else
            Call SHCA_Client.ESE_InitialSession(2, "", "", 0, 2, "", "") '初始化CA接口
            Call SHCA_Client.ESE_VerifySignData(strSource, "", strSignData, strSignCert)
        End If
        If SHCA_Client.errorCode = 0 Then
             MsgBox "验证签名成功！"
        Else
             MsgBox "验证签名失败！" & ValidateCertView(SHCA_Client.errorCode)
        End If
        Exit Function
errH:
104     MsgBox "验证签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function SHCA_GetPara(Optional ByVal bytFunc As Byte)
    Dim arrList As Variant
    
    On Error GoTo errH
    gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys)  '读取URLs 固定读取ZLHIS'gstrPara = "0&&&0"   '0-RSA;1-SM2&&&0-不启用签章图片;1-启用签章
    If Val(gstrPara) = 1 Then
        gudtPara.bytSignVersion = V_ESE
    ElseIf Val(gstrPara) = 0 Then
        gudtPara.bytSignVersion = V_SEH
    End If
    SHCA_GetPara = True
    Exit Function
errH:
    MsgBox "读取参数失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Function SHCA_SetParaStr() As String
    SHCA_SetParaStr = IIf(gudtPara.bytSignVersion = 0, "0", "1")
End Function

Public Function SHCA_CheckCert(ByVal strCurrCertSn As String, Optional ByRef strSigCert As String, Optional ByRef blnReDo As Boolean) As Boolean
        '功能：读取USB进行设备初始化并登录
        Dim strKey As String, strPIN As String, strUserName As String, strCertSn As String, strDate As String
        Dim strWebUrl As String, intDate   As Integer
        Dim blnRet As Boolean
        Dim udtUser As USER_INFO
        Dim intPoint As Integer
        On Error GoTo errH
         If Not SHCA_InitObj() Then
102         MsgBox "部件未初始化！"
            Exit Function
        End If
104     If Not GetCertList(strUserName, strKey, strSigCert, strCertSn) Then Exit Function
        intPoint = InStr(strKey, "F")
        If mUserInfo.strUserID = "" Then
            MsgBox "您的身份证号为空,请联系管理员到人员管理中录入！", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        ElseIf Mid(strKey, intPoint + 2) <> mUserInfo.strUserID Then
            MsgBox "您的身份证号：" & _
                       vbCrLf & vbTab & "【" & mUserInfo.strUserID & "】" & vbCrLf & _
                       "当前证书唯一标识:" & _
                       vbCrLf & vbTab & "【" & Mid(strKey, intPoint + 2) & "】" & vbCrLf & _
                       "用户身份证号与当前证书唯一标识不相等,不能使用！", vbInformation, gstrSysName
            Exit Function
        End If
110     If mLastPWD <> "" Then strPIN = mLastPWD
112     If strPIN = "" Then
114         If Not frmPassword.ShowMe(strPIN) Then Exit Function
        End If
        
116     If Not GetCertLogin(strKey, strPIN, strSigCert, intDate, strWebUrl) Then
118         strPIN = ""
            blnRet = False
        Else
            blnRet = True
        End If
        
        If blnRet Then
            '判断是否需要更新注册证书
            udtUser.strName = strUserName
            udtUser.strSignName = strUserName
            udtUser.strUserID = Mid(strKey, intPoint + 2) 'SF+身份证号
            udtUser.strCertSn = strCertSn
            udtUser.strCertDN = GetCertDN(strSigCert)
            udtUser.strCert = strSigCert
            udtUser.strEncCert = ""
            udtUser.strCertID = strKey
            '获取已经注册证书的有效结束日期 日期格式:axBJCASecCOMV21 这个版本解析出来的都是2015/09/15
            If gudtPara.bytSignVersion = V_SEH Then
                strDate = SHCA_Client.SEH_GetCertValidDate(mUserInfo.strCert)
            Else
                strDate = SHCA_Client.ESE_GetCertValidDate(mUserInfo.strCert)
            End If
            If IsUpdateRegCert(udtUser, strDate, blnReDo) Then
                blnRet = True
            Else
                blnRet = False
            End If
        End If
     
        mLastPWD = strPIN
        SHCA_CheckCert = blnRet
        Exit Function
errH:
124     MsgBox "检查USBKEY失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Public Sub SHCA_UloadObj()
    Set SHCA_Client = Nothing
    mblnInit = False
End Sub
'----- 以下是内部函数

''' 获取客户端证书列表
''' 返回boolean
Private Function GetCertList(ByRef strName As String, ByRef strUniqueID As String, ByRef strCert As String, Optional ByRef strCertSn As String) As Boolean
    '-入参:无
    '-出参
    'strName :      保存接口返回的证书所有者姓名
    'strUniqueID:   保存接口返回的证书所有者唯一标识
    'strCert:       保存接口返回的签名证书
    Dim strPassas As String
    On Error GoTo errH
    
    If gudtPara.bytSignVersion = V_SEH Then
        SHCA_Client.SEH_InitialSession 2, "", "", 0, 2, "", "" '初始化CA接口
        strCert = SHCA_Client.SEH_GetSelfCertificate(10, "com1", "")
        If SHCA_Client.errorCode <> 0 Then
            MsgBox ValidateCertView(SHCA_Client.errorCode)
            Exit Function
        End If
        
        strName = SHCA_Client.SEH_GetCertDetail(strCert, 17)
        If SHCA_Client.errorCode <> 0 Then
            MsgBox ValidateCertView(SHCA_Client.errorCode)
            Exit Function
        End If
        
        strUniqueID = SHCA_Client.SEH_GetCertInfoByOID(strCert, "1.2.156.112570.148")
        If SHCA_Client.errorCode <> 0 Then
            MsgBox ValidateCertView(SHCA_Client.errorCode)
            Exit Function
        End If
        
        strCertSn = SHCA_Client.SEH_GetCertDetail(strCert, 2)
        If SHCA_Client.errorCode <> 0 Then
            MsgBox ValidateCertView(SHCA_Client.errorCode)
            Exit Function
        End If
    ElseIf gudtPara.bytSignVersion = V_ESE Then
        Call SHCA_Client.ESE_InitialSession(2, "", "", 0, 2, "", "")
        strCert = SHCA_Client.ESE_GetSelfCertificate(36, "com1")
        If SHCA_Client.errorCode <> 0 Then
            MsgBox ValidateCertView(SHCA_Client.errorCode)
            Exit Function
        End If
    
        strName = SHCA_Client.ESE_GetCertDetail(strCert, 17)
        If SHCA_Client.errorCode <> 0 Then
            MsgBox ValidateCertView(SHCA_Client.errorCode)
            Exit Function
        End If
        
        strUniqueID = SHCA_Client.ESE_GetCertInfoByOID(strCert, "1.2.156.112570.148")
        If SHCA_Client.errorCode <> 0 Then
            MsgBox ValidateCertView(SHCA_Client.errorCode)
            Exit Function
        End If
        
        strCertSn = SHCA_Client.ESE_GetCertDetail(strCert, 2)
        If SHCA_Client.errorCode <> 0 Then
            MsgBox ValidateCertView(SHCA_Client.errorCode)
            Exit Function
        End If
    End If
    GetCertList = True
    Exit Function
errH:
    GetCertList = False
End Function

Private Function GetCertLogin(ByVal strUniqueID As String, ByVal strPassword As String, ByVal strCert As String, ByRef dDate As Integer, ByRef strWebserviceUrl As String) As Boolean
    '- 入参
    'strUniqueID : 证书唯一标识
    'strPassword : 证书密码
    'strWebserviceUrl:签名服务器地址，即为证书验证
    '- 出参
    'dDate       : 返回证书有效时间
    On Error GoTo errH
    Dim result As Boolean
    If SHCA_Client Is Nothing Then Set SHCA_Client = CreateObject("SafeEngineCOM.SafeEngineCtl")
    If (strPassword = "") Then
        MsgBox "请输入证书密码！"
    Else
        '证书安全登录
        'result:0:成功
        'result:非0:不成功
        If mLogin >= 8 Then
            MsgBox "已经输入了" & mLogin & "次错误密码，超过了最大输入次数！"
            Exit Function
        End If
        If gudtPara.bytSignVersion = V_SEH Then
            Call SHCA_Client.SEH_InitialSession(27, "com1", strPassword, 0, 27, "com1", "") '初始化CA接口(密码)
        Else
            Call SHCA_Client.ESE_InitialSession(36, "com1", strPassword, 0, 36, "com1", "") '初始化CA接口(密码)
        End If
        If SHCA_Client.errorCode = 0 Then
             '验证证书结果信息表示
            If gudtPara.bytSignVersion = V_SEH Then
                Call SHCA_Client.SEH_VerifyCertificate(strCert)
            Else
                Call SHCA_Client.ESE_VerifyCertificate(strCert)
            End If
            If SHCA_Client.errorCode = 0 Then
                
                '获取客户端证书有效期截止时间
                If gudtPara.bytSignVersion = V_SEH Then
                    dDate = SHCA_Client.SEH_GetCertValidDate(strCert)
                Else
                    dDate = SHCA_Client.ESE_GetCertValidDate(strCert)
                End If
                If (dDate <= 30 And dDate > 0) And Not gblnShow Then
                    MsgBox "您的证书还有" & dDate & "天过期"
                    gblnShow = True
                    GetCertLogin = True
                ElseIf (dDate <= 0) Then
                    MsgBox "您的证书已过期 " & Abs(dDate) & " 天"
                    GetCertLogin = False
                Else
                    GetCertLogin = True
                End If
            Else
               MsgBox "验证证书错误！" & ValidateCertView(SHCA_Client.errorCode)
            End If
        Else
            mLogin = mLogin + 1
            MsgBox "初始登陆错误！" & ValidateCertView(SHCA_Client.errorCode)
        End If
       
    End If
    Exit Function
errH:
    mLogin = mLogin + 1
    MsgBox "证书密码可能不正确，您已经输入了" & mLogin & "次密码，还可以输入" & 8 - mLogin & "次!"
    GetCertLogin = False
End Function

''' <summary>
''' 验证证书结果信息表示
''' </summary>
''' <remarks></remarks>
Private Function ValidateCertView(retValidateCert) As String
    Dim strErrorMsg As String
    Select Case retValidateCert
        Case 0
            strErrorMsg = ""
        Case -2113667072:
            strErrorMsg = "装载动态库错误(-2113667072)"
            
        Case -2113667071:
            strErrorMsg = "内存分配错误(-2113667071)"
            
        Case -2113667070:
            strErrorMsg = "读私钥设备错误(-2113667070)"
            
        Case -2113667069:
            strErrorMsg = "私钥密码错误(-2113667069)"
            
        Case -2113667068:
            strErrorMsg = "读证书链设备错误(-2113667068)"
            
        Case -2113667067:
            strErrorMsg = "证书链密码错误(-2113667067)"
            
        Case -2113667066:
            strErrorMsg = "读证书设备错误(-2113667066)"
            
        Case -2113667065:
            strErrorMsg = "证书密码错误(-2113667065)"
            
        Case -2113667064:
            strErrorMsg = "私钥超时(-2113667064)"
            
        Case -2113667063:
            strErrorMsg = "缓冲区太小(-2113667063)"
            
        Case -2113667062:
            strErrorMsg = "初始化环境错误(-2113667062)"
            
        Case -2113667061:
            strErrorMsg = "清除环境错误(-2113667061)"
            
        Case -2113667060:
            strErrorMsg = "数字签名错误(-2113667060)"
            
        Case -2113667059:
            strErrorMsg = "验证签名错误(-2113667059)"
            
        Case -2113667058:
            strErrorMsg = "摘要错误(-2113667058)"
            
        Case -2113667057:
            strErrorMsg = "证书格式错误(-2113667057)"
            
        Case -2113667056:
            strErrorMsg = "数字信封错误(-2113667056)"
            
        Case -2113667055:
            strErrorMsg = "从LDAP获取证书错误(-2113667055)"
            
        Case -2113667054:
            strErrorMsg = "证书已过期(-2113667054)"
            
        Case -2113667053:
            strErrorMsg = "获取证书链错误(-2113667053)"
            
        Case -2113667052:
            strErrorMsg = "证书链格式错误(-2113667052)"
            
        Case -2113667051:
            strErrorMsg = "验证证书链错误(-2113667051)"
            
        Case -2113667050:
            strErrorMsg = "证书已废除(-2113667050)"
            
        Case -2113667049:
            strErrorMsg = "CRL格式错误(-2113667049)"
            
        Case -2113667048:
            strErrorMsg = "连接OCSP服务器错误(-2113667048)"
            
        Case -2113667047:
            strErrorMsg = "OCSP请求编码错误(-2113667047)"
            
        Case -2113667046:
            strErrorMsg = "OCSP回包错误(-2113667046)"
            
        Case -2113667045:
            strErrorMsg = "OCSP回包格式错误(-2113667045)"
            
        Case -2113667044:
            strErrorMsg = "OCSP回包过期(-2113667044)"
            
        Case -2113667043:
            strErrorMsg = "OCSP回包验证签名错误(-2113667043)"
            
        Case -2113667042:
            strErrorMsg = "证书状态未知(-2113667042)"
            
        Case -2113667041:
            strErrorMsg = "对称加解密错误(-2113667041)"
            
        Case -2113667040:
            strErrorMsg = "获取证书信息错误(-2113667040)"
            
        Case -2113667039:
            strErrorMsg = "获取证书细目错误(-2113667039)"
            
        Case -2113667038:
            strErrorMsg = "获取证书唯一标识错误(-2113667038)"
            
        Case -2113667037:
            strErrorMsg = "获取证书扩展项错误(-2113667037)"
            
        Case -2113667036:
            strErrorMsg = "PEM编码错误(-2113667036)"
            
        Case -2113667035:
            strErrorMsg = "PEM解码错误(-2113667035)"
            
        Case -2113667034:
            strErrorMsg = "产生随机数错误(-2113667034)"
            
        Case -2113667033:
            strErrorMsg = "PKCS12参数错误(-2113667033)"
            
        Case -2113667032:
            strErrorMsg = "私钥格式错误(-2113667032)"
            
        Case -2113667031:
            strErrorMsg = "公私钥不匹配(-2113667031)"
            
        Case -2113667030:
            strErrorMsg = "PKCS12编码错误(-2113667030)"
            
        Case -2113667029:
            strErrorMsg = "PKCS12格式错误(-2113667029)"
            
        Case -2113667028:
            strErrorMsg = "PKCS12解码错误(-2113667028)"
            
        Case -2113667027:
            strErrorMsg = "非对称加解密错误(-2113667027)"
            
        Case -2113667026:
            strErrorMsg = "OID格式错误(-2113667026)"
            
        Case -2113667025:
            strErrorMsg = "LDAP地址格式错误(-2113667025)"
            
        Case -2113667024:
            strErrorMsg = "LDAP地址错误(-2113667024)"
            
        Case -2113667023:
            strErrorMsg = "连接LDAP服务器错误(-2113667023)"

        Case -2113667022:
            strErrorMsg = "LDAP绑定错误(-2113667022)"
            
        Case -2113667021:
            strErrorMsg = "没有OID对应的扩展项(-2113667021)"
            
        Case -2113667020:
            strErrorMsg = "获取证书级别错误(-2113667020)"
            
        Case -2113667019:
            strErrorMsg = "读取配置文件错误(-2113667019)"
            
        Case -2113667018:
            strErrorMsg = "私钥未载入(-2113667018)"
            
  ' 以下错误用于登录
        Case -2113666824:
            strErrorMsg = "无效的登录凭证(-2113666824)"
            
        Case -2113666823:
            strErrorMsg = "参数错误(-2113666823)"
            
        Case -2113666822:
            strErrorMsg = "不是服务器证书(-2113666822)"
            
        Case -2113666821:
            strErrorMsg = "登录错误(-2113666821)"
            
        Case -2113666820:
            strErrorMsg = "证书验证方式错误(-2113666820)"
            
        Case -2113666819:
            strErrorMsg = "随机数验证错误(-2113666819)"
            
        Case -2113666818:
            strErrorMsg = "与单点登录客户端代理通信(-2113666818)"
    End Select
    ValidateCertView = strErrorMsg
End Function





