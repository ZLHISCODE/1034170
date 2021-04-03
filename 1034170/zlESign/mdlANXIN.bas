Attribute VB_Name = "mdlANXIN"

Option Explicit
'吉林国投安信CA         七台河人民医院
Private mobjJLClient As Object   'JITComVCTKEx.JITVCTKEx.1
Private mobjJLServer As Object  'JITClientCOMAPI.JITClientProc.1
Private mobjCertInfo As Object    '签章 JITCertActiveX.CertInfo.1

'Private mobjJLClient As New JITComVCTKExLib.JITVCTKEx
'Private mobjJLServer As New JITClientCOMAPILib.JITClientProc
'Private mobjCertInfo As New JITCertActiveXLib.CertInfo    '签章
Private mblnInit As Boolean

Private mstrPWD As String          '缓存输入的密码
Private mIntPwd As Integer          '最多允许输入8次

Private Const M_STR_PARA As String = "<?xml version=""1.0"" encoding=""gb2312""?><authinfo><liblist>" & _
        "<lib type=""CSP"" version=""1.0"" dllname="""" ><algid val=""SHA1"" sm2_hashalg=""sm3""/></lib>" & _
        "<lib type=""SKF"" version=""1.1"" dllname=""SERfR01DQUlTLmRsbA=="" ><algid val=""SHA1"" sm2_hashalg=""sm3""/></lib>" & _
        "<lib type=""SKF"" version=""1.1"" dllname=""U2h1dHRsZUNzcDExXzMwMDBHTS5kbGw="" ><algid val=""SHA1"" sm2_hashalg=""sm3""/></lib>" & _
        "<lib type=""SKF"" version=""1.1"" dllname=""QU5YSU5Dc3AxMV8zMDAwR01BLmRsbA=="" ><algid val=""SHA1"" sm2_hashalg=""sm3""/></lib>" & _
        "</liblist><checkkeytimes><item times=""3"" ></item></checkkeytimes></authinfo>"

Public Function ANXIN_InitObj() As Boolean
     '证书部件初始化
    Dim lngRet As Long
    Dim strTSAIP As String
    Dim varTmp As Variant
100     If glngSign > 1 Then ANXIN_InitObj = True: Exit Function
        On Error Resume Next
102     If mobjJLClient Is Nothing Then
104         Set mobjJLClient = CreateObject("JITComVCTKEx.JITVCTKEx.1")
106         If Err.Number <> 0 Then
108             MsgBox "创建安信签名对象【JITComVCTK_S.dll】失败！请检查该控件是否安装并注册。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
110     Err.Clear
112     If mobjJLServer Is Nothing Then
114         Set mobjJLServer = CreateObject("JITClientCOMAPI.JITClientProc.1")
116         If Err.Number <> 0 Then
118             MsgBox "创建安信签名对象【JITClientCOMAPI.dll】失败！请检查该控件是否安装并注册。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
120     Err.Clear
122     If mobjCertInfo Is Nothing Then
124         Set mobjCertInfo = CreateObject("JITCertActiveX.SZZSCertInfo.1")
126         If Err.Number <> 0 Then
128             MsgBox "创建安信签章对象【JITCertActiveX.dll】失败！请检查该控件是否安装并注册。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
130     Err.Clear: On Error GoTo 0
        On Error GoTo errH
        '参数信息:是否启用时间戳[0-不启用;1-启用]&&&签名服务器IP&&&签名服务器端口&&&时间戳IP&&&时间戳端口
        '第一位 签名服务器，第二位时间戳服务器，第三位网关。安信就这3个硬件。如果没有硬件就是“000”，只有签名服务器就是“100”
        '为了兼容以前如果第一个参数=0;则代表只有签名服务器就是“100” ;=1代表启用签名服务器启用时间戳服务器 就是"110"第三位网关暂时未启用，预留参数
        '连接时间戳服务器
        'gstrPara = "1&&&175.17.252.155&&&8000&&&175.17.252.156&&&8000" '七台河人民医院
        'gstrPara = "000&&&175.17.252.155&&&8000&&&175.17.252.156&&&8000" '七台河妇幼医院
        
132     gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys, , "") '读取配置内容
        
134     If gstrPara = "" Then
136         Err.Raise -1, , "当前系统【" & glngSys & "】没有配置电子签名参数,请到启用电子签名接口处设置。"
            Exit Function
        End If
138     If UBound(Split(gstrPara, G_STR_SPLIT)) <> 4 Then
140         MsgBox "电子签名参数值设置有误,请检查。" & vbCrLf & _
                "当前参数值:" & gstrPara & vbCrLf & _
                "正确格式:签名服务器IP&&&时间戳服务器IP", vbInformation, gstrSysName
            Exit Function
        Else
142         Call ANXIN_GetPara
        End If
144     lngRet = mobjJLClient.Initialize(M_STR_PARA)
146     If Not GetErrorInfo("Initialize") Then Exit Function
148     mblnInit = True
150     mIntPwd = 8
152     mstrPWD = ""
154     ANXIN_InitObj = True
        
        Exit Function
errH:
156  MsgBox "创建接口部件失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbExclamation, gstrSysName
End Function

Public Function ANXIN_Sign(ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String, ByRef strTimeStampCode As String, Optional ByRef blnReDo As Boolean) As Boolean
'签名
    Dim lngRet As Long
    Dim strErr As String
    Dim strHash As String
    
    Dim blnCheck As Boolean
        On Error GoTo errH
        blnCheck = ANXIN_CheckCert(blnReDo)
        If blnReDo Then Exit Function
100     If blnCheck Then                 '验证当前USB是否是签名用户的，并获取签名证书
            '证书ID进行签名
            lngRet = mobjJLClient.SetPinCode(mstrPWD)
110         strSignData = mobjJLClient.DetachSignStr("", strSource)      'DetachSignStr-不带原文签名;AttachSignStr-带原文签名
            If Not GetErrorInfo("DetachSignStr") Then Exit Function
            If strSignData <> "" Then
                If gudtPara.blnISTS Then
                    If Not ConnectToTsaServer() Then Exit Function
                    strHash = StringSHA1(strSource)
                    strTimeStampCode = mobjJLServer.TsaSign("", 1, strHash)            '申请时间戳 传入签名值过长，签名时比较耗时,故采用固定值
                    strTimeStamp = mobjJLServer.VerifyTsaSign(strTimeStampCode)
                    Call mobjJLServer.FinalizeServerConnectEx    '断开时间戳服务器连接
                    If strTimeStampCode = "" Then MsgBox "获取时间戳失败！", vbInformation, gstrSysName: Exit Function
                    '日期格式化
                    strTimeStamp = Mid(strTimeStamp, 1, 14)
                    strTimeStamp = String14ToDate(strTimeStamp, strErr)
                    If strErr <> "" Then MsgBox strErr, vbInformation, gstrSysName: Exit Function
                    '转东八区时间
                    strTimeStamp = Format(DateAdd("h", 8, strTimeStamp), "YYYY-MM-DD HH:MM:SS")
                Else
                    strTimeStamp = Format(gobjComLib.zlDatabase.Currentdate & "", "yyyy-MM-dd HH:mm:ss")
                End If
            Else
                MsgBox "签名失败！", vbInformation, gstrSysName
                Exit Function
            End If
112
        Else
            MsgBox "签名失败！", vbInformation, gstrSysName
            Exit Function
        End If
        ANXIN_Sign = True
        Exit Function
errH:
114     MsgBox "签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbExclamation, gstrSysName
End Function

Public Function ANXIN_VerifySign(ByVal strSource As String, ByVal strSignData As String, ByVal strTimeStampCode As String) As Boolean
'功能;验证签名
'参数:strSignData -签名值
     Dim blnRet As Boolean
     Dim lngRet As Long
     Dim strTS As String
     On Error GoTo errH
     
    If gudtPara.blnIsSign Then
        '服务器验签
        If Not ConnectToSignServer() Then Exit Function
        lngRet = mobjJLServer.VerifyDetachedSign(strSignData, strSource) '服务器验证数据 不带原文签名:VerifyDetachedSign(string, string);带原文签名  VerifyAttachedSign
        If lngRet <> 0 Then
            MsgBox "签名验证失败:" & mobjJLServer.GetErrorMessage(lngRet), vbInformation, gstrSysName
            Call mobjJLServer.FinalizeServerConnectEx   '断开签名服务器链接
            Exit Function
        End If
        Call mobjJLServer.FinalizeServerConnectEx   '断开签名服务器链接
    Else
        Call mobjJLClient.VerifyAttachedSign(strSignData)  '客户端验证签名
        lngRet = mobjJLClient.GetErrorCode()
        If lngRet <> 0 Then
            MsgBox "验证签名失败，错误码：" & lngRet & " 错误信息：" & mobjJLClient.GetErrorMessage(lngRet), vbInformation, gstrSysName
            Exit Function
        End If
    End If

    '连接时间戳服务器
    If gudtPara.blnISTS Then
        If Not ConnectToTsaServer() Then Exit Function
        strTS = mobjJLServer.VerifyTsaSign(strTimeStampCode)
        If strTS = "" Then
              MsgBox "时间戳验证失败！", vbInformation, gstrSysName
              Call mobjJLServer.FinalizeServerConnectEx   '断开时间戳服务器
              Exit Function
        End If
        Call mobjJLServer.FinalizeServerConnectEx   '断开时间戳服务器
    End If
    MsgBox "验证成功，该电子签名数据有效!", vbInformation, gstrSysName
    
     ANXIN_VerifySign = True
     Exit Function
errH:
104     MsgBox "验证签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbExclamation, gstrSysName
End Function


Public Function ANXIN_CheckCert(ByRef blnReDo As Boolean) As Boolean
    '功能：读取USB进行设备初始化并登录
    Dim strKeySN As String, strUserID As String, strUserName As String, strCertDN As String
    Dim strDate As String
    Dim arrDN As Variant
    Dim udtUser As USER_INFO
    Dim blnRet As Boolean
    Dim i As Long
    
    On Error GoTo errH
    If Not GetCertList(strKeySN, strUserName, strCertDN, , strUserID) Then Exit Function
    If mUserInfo.strUserID = "" Then
        MsgBox "您的身份证号为空,请联系管理员到人员管理中录入！", vbOKOnly + vbInformation, gstrSysName
        Exit Function
     ElseIf mUserInfo.strUserID <> strUserID Then
        MsgBox "该证书未注册在您的名下，不能使用！"
        Exit Function
    End If
    
    '判断是否需要更新注册证书
    udtUser.strName = strUserName
    udtUser.strSignName = strUserName
    udtUser.strUserID = strUserID '身份证号
    udtUser.strCertSn = strKeySN
    udtUser.strCertDN = strCertDN
    udtUser.strCert = ""
    udtUser.strEncCert = ""
    udtUser.strCertID = ""
    udtUser.strPicPath = ""
    arrDN = Split(mUserInfo.strCertDN, ",")     'CN=王二小U3294, O=七台河人民医院, L=七台河市, S=黑龙江省, C=CN, 有效日期=
    For i = 0 To UBound(arrDN)
        If Trim(arrDN(i)) Like "有效日期*" Then
            strDate = Trim(Split(arrDN(i), "=")(1))
            Exit For
        End If
    Next
    If IsUpdateRegCert(udtUser, strDate, blnReDo) Then
        blnRet = True
    Else
        blnRet = False
    End If

    ANXIN_CheckCert = blnRet
    Exit Function
errH:
     MsgBox "检查USBKEY失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbExclamation, gstrSysName
End Function


Public Function ANXIN_RegCert(arrCertInfo As Variant) As Boolean
    '功能：提供在HIS数据库中注册个人证书的必要信息,用于新发放或更换证书,,需要插入USB-Key
    '返回：arrCertInfo作为数组返回证书相关信息
'      0-ClientSignCertCN:客户端签名证书公共名称(姓名),用于注册证书时程序验证身份
'      1-ClientSignCertDN:客户端签名证书主题(每人唯一)
'      2-ClientSignCertSN:客户端签名证书序列号(每证书唯一)
'      3-ClientSignCert:客户端签名证书内容
'      4-ClientEncCert:客户端加密证书内容
'      5-签名图片文件名,空串表示没有签名图片

        Dim strKeyId As String, strCertUserName As String, strCertDN As String, strPicPath As String
        Dim i As Integer
        On Error GoTo errH

        For i = LBound(arrCertInfo) To UBound(arrCertInfo)
             arrCertInfo(i) = ""
        Next

        If GetCertList(strKeyId, strCertUserName, strCertDN, strPicPath) Then
            arrCertInfo(0) = strCertUserName
            arrCertInfo(1) = strCertDN
            arrCertInfo(2) = strKeyId
            arrCertInfo(5) = strPicPath
            ANXIN_RegCert = True
        End If

        Exit Function
errH:
     MsgBox "获取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbExclamation, gstrSysName

End Function

Private Function GetCertList(Optional ByRef strUniqueID As String = "-1", Optional ByRef strName As String = "-1", Optional ByRef strCertDN As String = "-1", _
    Optional ByRef strPicPath As String = "-1", Optional ByRef strUserID As String = "-1", Optional ByRef strDate As String = "-1") As Boolean
'功能:获取安信证书详情
'strUserID-身份证号
    Dim lngRet As Long
    Dim strRet As String
    Dim datCurrent As Date
    Dim arrDN As Variant
    Dim i As Long
    Dim strKeyCount As String
    Dim strPic As String, strPIN As String
    Dim strTmp As String
    Dim lngDay As Long
    On Error GoTo errH
        
    If Not mblnInit Then
        lngRet = mobjJLClient.Initialize(M_STR_PARA)
        If Not GetErrorInfo("Initialize") Then Exit Function
        mblnInit = True
    End If
    lngRet = mobjJLClient.SetCertChooseType(1)
    lngRet = mobjJLClient.SetCert("SC", "", "", "", "CN = AnXin SM2 CA,O = AnXin CA,C = CN", "")
    If Not GetErrorInfo("SetCert") Then Exit Function
    strDate = mobjJLClient.GetCertInfo("SC", 6, "")     '有效日期
    If IsDate(strDate) Then
        '检查证书是否过期
        lngDay = CheckValidaty(strDate)
        If (lngDay <= 30 And lngDay > 0 And Not gblnShow) Then
            MsgBox "您的证书还有" & lngDay & "天过期", vbInformation, gstrSysName
            gblnShow = True
        ElseIf (lngDay <= 0) Then
            MsgBox "您的证书已过期 " & Abs(lngDay) & " 天"
            Exit Function
        End If
    End If
    If strUniqueID <> "-1" Then strUniqueID = mobjJLClient.GetCertInfo("SC", 2, "")     '证书序列号
    If strCertDN <> "-1" Or strName <> "-1" Then
        strCertDN = mobjJLClient.GetCertInfo("SC", 0, "") 'CN=王二小U3294, O=七台河人民医院, L=七台河市, S=黑龙江省, C=CN, 有效日期=
        If strCertDN <> "" Then
            arrDN = Split(strCertDN, ",")
            For i = 0 To UBound(arrDN)
                If Trim(arrDN(i)) Like "CN*" Then
                    strName = Trim(Split(arrDN(i), "=")(1))
                    Exit For
                End If
            Next
        End If
        strCertDN = strCertDN & ", 有效日期=" & strDate
    End If
    
    If strUserID <> "-1" Then
        strUserID = ""
        strTmp = mobjJLClient.GetCertInfo("SC", 7, "1.2.86.11.7.1")  '身份证号需要转ASCII :31 16 a0 14 13 12 34 33 32 35 30 33 31 39 38 36 30 31 31 32 36 32 31 35
        If Not GetErrorInfo("GetCertInfo") Then Exit Function
        If strTmp <> "" Then
            arrDN = Split(strTmp, " ")
            For i = 6 To UBound(arrDN)    '前6个字符为前缀
                strUserID = strUserID & Chr(Val("&H" & arrDN(i)))
            Next
        End If
    End If
    
    If mstrPWD = "" Then
CheckPWD:
        If Not frmPassword.ShowMe(mstrPWD, 6, 16) Then Exit Function
        lngRet = mobjCertInfo.VerifyUserPin("ANXIN3KGM", mstrPWD)
        'VB调试的时候单步跟踪返回乱码；直接运行返回正确字符串'{"RetryCount":"0","VerifyValue":"1"}
        '后调整成老版本 返回1-成功；0-失败
        '首次验证密码
        If lngRet = 0 Then
            mIntPwd = mIntPwd - 1
            mstrPWD = ""
            If mIntPwd > 0 Then
                MsgBox "验证密码失败,您还有" & mIntPwd & "次密码重试机会!", vbInformation + vbOKOnly, gstrSysName
                GoTo CheckPWD
            Else
                MsgBox "验证密码次数过多,请找管理员解锁。", vbInformation + vbOKOnly, gstrSysName
                Exit Function                     '密码错误重置密码
            End If
        Else
            mIntPwd = 8
        End If
    End If

    '获得Key数量
    If strPicPath <> "-1" Then
        '获取签章耗时,签名时不读取，只在注册的时候获取
        'strKeyCount = [{"KeyName":"安信电子钥匙 ","KeyType":"ANXIN3KGM","UsbKeySerialNumber":"AX00010415"},{"KeyName":"安信电子钥匙","KeyType":"ANXIN3KGM","UsbKeySerialNumber":"AX00010414"}]
        strKeyCount = mobjCertInfo.GetKeyCount("ANXIN3KGM")
        strRet = strKeyCount 'VB调试的时候单步跟踪返回乱码；直接运行返回正确字符串
        If strRet <> "" Then
            If UBound(Split(strRet, "},{")) = 0 Then
                strPic = mobjCertInfo.ReadImageData("ANXIN3KGM", mstrPWD)
                If Len(strPic) > 1 Then
                    strPicPath = SaveBase64ToFile("gif", strUniqueID, strPic)
                Else
                    MsgBox "读取签章信息失败！", vbInformation, gstrSysName
                    Exit Function
                End If
            ElseIf Val(strKeyCount) > 0 Then
                MsgBox "请选择唯一的KEY盘插入！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If

    GetCertList = True
    Exit Function
errH:
    MsgBox "获取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbExclamation, gstrSysName
End Function

Public Function ANXIN_GetSeal() As String
'获取签章图片
    Dim strPicPath As String
    Call GetCertList(, , , strPicPath)
    ANXIN_GetSeal = strPicPath
End Function

Private Function GetErrorInfo(ByVal strName As String) As Boolean
    Dim lngRet As Long

    On Error GoTo errH
    lngRet = mobjJLClient.GetErrorCode  'lngRet -536870826 密码不对;-536870823  指定的密码太长或太短
    If lngRet <> 0 Then
        MsgBox "调用接口：【" & strName & "】后出错,错误描述:" & vbCrLf & mobjJLClient.GetErrorMessage(lngRet), vbInformation, gstrSysName
        Exit Function
    End If
    GetErrorInfo = True
    Exit Function
errH:
    MsgBox "获取错误描述！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbExclamation, gstrSysName
End Function

Private Function ConnectToTsaServer() As Boolean
    Dim lngRet As Long

    On Error GoTo errH

    lngRet = mobjJLServer.InitServerConnectEx(gudtPara.strTSIP, CInt(gudtPara.strTSPort))
    If lngRet <> 0 Then
        MsgBox mobjJLServer.GetErrorMessage(lngRet), vbInformation, gstrSysName
        Exit Function
    End If
    lngRet = mobjJLServer.SetServerUriEx("/signserver/service/xml")
    If lngRet <> 0 Then
        MsgBox mobjJLServer.GetErrorMessage(lngRet), vbInformation, gstrSysName
        Call mobjJLServer.FinalizeServerConnectEx '终止服务器连接，以释放连接句柄
        Exit Function
    End If
    ConnectToTsaServer = True
    Exit Function
errH:
    MsgBox "连接时间戳服务器！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbExclamation, gstrSysName
End Function

Private Function ConnectToSignServer() As Boolean
    Dim lngRet As Long

    On Error GoTo errH
    lngRet = mobjJLServer.InitServerConnectEx(gudtPara.strSIGNIP, CInt(gudtPara.strSignPort))
    If lngRet <> 0 Then
        MsgBox mobjJLServer.GetErrorMessage(lngRet), vbInformation, gstrSysName  '连接服务器失败
        Exit Function
    End If
    lngRet = mobjJLServer.SetServerUriEx("/signserver/service/xml")
    If lngRet <> 0 Then
        MsgBox mobjJLServer.GetErrorMessage(lngRet), vbInformation, gstrSysName
        Call mobjJLServer.FinalizeServerConnectEx '终止服务器连接，以释放连接句柄
        Exit Function
    End If
    lngRet = mobjJLServer.SetCertAliasEx("")  '设置服务器签名时的签名证书标识,空为默认证书
    If lngRet <> 0 Then
        MsgBox mobjJLServer.GetErrorMessage(lngRet), vbInformation, gstrSysName
        Call mobjJLServer.FinalizeServerConnectEx '终止服务器连接，以释放连接句柄
        Exit Function
    End If
    ConnectToSignServer = True
    Exit Function
errH:
    MsgBox "连接签名戳服务器！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbExclamation, gstrSysName
End Function


Public Function ANXIN_GetPara() As Boolean
    Dim arrList As Variant
    
    On Error GoTo errH
    If gstrPara = "" Then gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys, , "") '读取URLs 固定读取ZLHIS 系统默认100
    If gstrPara = "" Then gstrPara = "110&&&175.17.252.155&&&8000&&&175.17.252.156&&&8000"   '格式是否启用设备[000-都不启用;111-都启用]&&&签名服务器IP&&&签名服务器端口&&&时间戳IP&&&时间戳端口
    arrList = Split(gstrPara, "&&&")
    If UBound(arrList) >= 4 Then
        If Len(arrList(0)) = 3 Then
            gudtPara.blnIsSign = Mid(arrList(0), 1, 1) = "1"
            gudtPara.blnISTS = Mid(arrList(0), 2, 1) = "1"
        Else
            gudtPara.blnISTS = Val(arrList(0)) = 1
            gudtPara.blnIsSign = True
        End If
        gudtPara.strSIGNIP = arrList(1)
        gudtPara.strSignPort = arrList(2)
        gudtPara.strTSIP = arrList(3)
        gudtPara.strTSPort = arrList(4)
    Else
        gudtPara.blnISTS = True
        gudtPara.blnIsSign = True
        gudtPara.strSIGNIP = "175.17.252.155"
        gudtPara.strSignPort = "8000"
        gudtPara.strTSIP = "175.17.252.156"
        gudtPara.strTSPort = "8000"
    End If
    Exit Function
errH:
    MsgBox "读取参数失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Function ANXIN_SetParaStr() As String
    With gudtPara
        ANXIN_SetParaStr = IIf(.blnIsSign, "1", "0") & IIf(.blnISTS, "1", "0") & "0" & G_STR_SPLIT & IIf(Trim(.strSIGNIP) = "", "175.17.252.155", .strSIGNIP) & G_STR_SPLIT & IIf(Trim(.strSignPort) = "", "8000", .strSignPort) & _
            G_STR_SPLIT & IIf(Trim(.strTSIP) = "", "175.17.252.156", .strTSIP) & G_STR_SPLIT & IIf(Trim(.strTSPort) = "", "8000", .strTSPort)
    End With
End Function

Public Sub ANXIN_UnLoadObj()
    On Error Resume Next
    Set mobjJLServer = Nothing
    Set mobjCertInfo = Nothing
    Call mobjJLClient.Finalize
    Set mobjJLClient = Nothing
End Sub




