VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHNCA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'河南商丘项目
Private mobjHNCACTRL As Object
Private mobjHNCATS As Object
Private mblnTs As Boolean
Private mbytFunc As Byte   '0-RSA算法,1-SM2算法

Private Declare Function doPostEx Lib "HNCACertTools6.dll" (ByVal strUrl As String, ByVal strAPPCode As String, ByVal strSource As String, _
        ByVal strSignData As String, ByVal strCertPUB As String, ByVal strGBK As String) As Integer

Public Function HNCA_CheckCert(ByRef blnReDo As Boolean) As Boolean
'功能：读取USB进行设备初始化并登录
    Dim strKey As String, strPIN As String, strUserName As String
    Dim strCertName As String, strCertDN As String, strPicPath As String
    Dim strCertSn As String
    Dim strCertUserID As String    '包含身份证号信息
    Dim strDate As String
    Dim udtUser As USER_INFO
    Dim strCert As String, strCertID As String
    Dim blnOk As Boolean
    Dim blnRet As Boolean
    
    On Error GoTo errH
1000
     '获取证书信息同时检查Key盘是否插入
    If Not GetCertList(strCertName, strCertSn, strCert, strCertDN, strPicPath, strCertUserID) Then
        HNCA_CheckCert = False: Exit Function
    End If
    '未注册在当前用户名下的Key
    If mbytFunc = 0 Then
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
    Else
        If mUserInfo.strCertSn <> strCertSn Then
            MsgBox "该证书未注册在您的名下，不能使用！", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
1100
    '登录验证
    If InStr(gstrLogins & "|", "|" & strCertSn & "|") > 0 Then '首次验证通过后，下次不在继续验证
        blnOk = True
    Else
        If Not GetCertLogin(strCert, strCertDN) Then
            blnOk = False
        Else
            blnOk = True
            If InStr(gstrLogins & "|", "|" & strCertSn & "|") = 0 Then gstrLogins = gstrLogins & "|" & strCertSn
        End If
    End If
1200
    If blnOk And mbytFunc = 0 Then
        '判断是否需要更新注册证书
        udtUser.strName = strCertName
        udtUser.strSignName = strCertName
        udtUser.strUserID = strCertUserID
        udtUser.strCertSn = strCertSn
        udtUser.strCertDN = strCertDN
        udtUser.strCert = strCert
        udtUser.strPicPath = strPicPath
        udtUser.strPicCode = ""
        '获取已经注册证书的有效结束日期 日期格式:2005年12月23日 18:09:31
        strDate = mobjHNCACTRL.GetCertInfo(6, "")
        strDate = Format(strDate, "YYYY-MM-DD HH:MM:SS")
        If IsUpdateRegCert(udtUser, strDate, blnReDo) Then
            HNCA_CheckCert = True
        Else
            HNCA_CheckCert = False
        End If
    Else
        HNCA_CheckCert = blnOk
    End If
1300
    
    Exit Function
errH:
     MsgBox "检查USBKEY失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Function HNCA_RegCert(arrCertInfo As Variant) As Boolean
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

            
124         HNCA_RegCert = True
        End If

        Exit Function
errH:
126     MsgBox "获取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName

End Function

Public Function HNCA_Sign(ByVal strCurrCertSn As String, ByVal strSource As String, ByRef strSignData As String, _
        ByRef strTimeStamp As String, ByRef strTimeStampCode As String, ByRef blnReDo As Boolean) As Boolean
        '签名
        Dim strTSURL As String, strUserName As String, strUserPWD As String
        Dim strRequest As String
        Dim lngRet As Long, intPort As Integer, intSSLPort As Integer
        Dim blnCheck As Boolean, blnRet As Boolean
        Dim strMsg As String
        Dim arrTmp As Variant
        
10      On Error GoTo errH
20      blnCheck = HNCA_CheckCert(blnReDo)
30      If blnReDo Then Exit Function
40      If blnCheck Then
            '验证当前USB是否是签名用户的，并获取签名证书
50          If mbytFunc = 0 Then
60              strSignData = mobjHNCACTRL.Sign(strCurrCertSn, strSource)    '签名值不含数据源
70          Else
80              strSignData = mobjHNCACTRL.SM2Sign("", strSource)    '签名值不含数据源
90          End If
100         If strSignData <> "" Then
110             If mblnTs Then
                    '由原文申请时间戳
120                 lngRet = -1
130                 strRequest = mobjHNCATS.SOF_CreateTimeStampRequestByPlain(strSource)
140                 lngRet = mobjHNCATS.SOF_GetLastError()
150                 If (lngRet <> 0) Then
160                     strMsg = "创建时间戳请求失败，失败代码是：" & lngRet
170                     blnRet = False
180                 Else
190                     lngRet = -1
200                     strTimeStampCode = mobjHNCATS.SOF_CreateTimeStampResponse(strRequest)
210                     lngRet = mobjHNCATS.SOF_GetLastError()
220                     If (lngRet <> 0) Then
230                         strMsg = "创建时间戳响应失败，失败代码是：" & lngRet
240                         blnRet = False
250                     Else
260                         strTimeStamp = mobjHNCATS.SOF_GetTimeStampInfo(strTimeStampCode, 1) '格式;2015年10月20日 02:18:45.534  GMT
270                         strTimeStamp = Format(Mid(strTimeStamp, 1, 20), "YYYY-MM-DD HH:MM:SS")
280                         blnRet = True
290                     End If
300                 End If
310             Else
320                 blnRet = True
330                 strTimeStamp = Format(gobjComLib.zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
340             End If
350         Else
360             strMsg = "签名失败！"
370             blnRet = False
380         End If
390     Else
400         strMsg = "签名失败！"
410         blnRet = False
420     End If
430     If strMsg <> "" Then
440         MsgBox strMsg, vbInformation, gstrSysName
450     End If
                
460     HNCA_Sign = blnRet
470     Exit Function
errH:
480     MsgBox "签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function HNCA_VerifySign(ByVal strCert As String, ByVal strSignData As String, ByVal strSource As String, ByVal strTimeStampCode As String) As Boolean
'验证签名
'
        Dim strSigCert As String, strMsg As String
        Dim lngRet As Long
        Dim blnRet As Boolean
10      On Error GoTo errH
     
        '获取客户端签名(公钥)证书Base64编码 客户端验证签名,不含数据源的验证签名
20      If mbytFunc = 0 Then
30          lngRet = mobjHNCACTRL.VerifySign(strSignData, strSource, strCert)
40      Else
50          lngRet = mobjHNCACTRL.SM2Verify(strSignData, strSource, strCert)
60      End If
70      If lngRet = 0 Then   '验证签名成功
80          If mblnTs And strTimeStampCode <> "" Then
                '验证时间戳
90              lngRet = -1
100             lngRet = mobjHNCATS.SOF_VerifyTimeStampByPlain(strSource, strTimeStampCode)
110         End If
120         If (lngRet = 0) Then
130             strMsg = "验签成功！"
140             blnRet = True
150         Else
160             strMsg = "时间戳验证失败，错误代码：" & lngRet
170             blnRet = False
180         End If
190     Else
200         strMsg = "验签失败！"
210         blnRet = False
220     End If
            
230     If strMsg <> "" Then
240         MsgBox strMsg, vbInformation, gstrSysName
250     End If
260     HNCA_VerifySign = blnRet
        
270     Exit Function
errH:
280     MsgBox "验证签名失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function


Private Function GetCertList(ByRef strName As String, ByRef strCertSn As String, ByRef strCertPUB As String, _
                ByRef strCertDN As String, Optional strPicPath As String, Optional strUserID As String) As Boolean
'功能:获取证书信息
'-出参
'    strName 证书持有者姓名
'   strCertSN 证书唯一标识
'   strCertPUB 证书公钥  验证签名用到
'   strCertDN 证书描述信息  证书注册用到
'   strPicPath 证书图片保存位置

    Dim strCertCN As String, strPic As String
    Dim lngRet As Long, intCnIndex As Integer
    Dim strPara As String, strSign As String
    
    On Error GoTo errH

10    If mobjHNCACTRL Is Nothing Then Set mobjHNCACTRL = CreateObject("HNCACTRL.HNCACtrlCtrl.1")
      If mobjHNCATS Is Nothing Then Set mobjHNCATS = CreateObject("TSAACTIVEXVER4.TsaActiveXver4Ctrl.1")    '时间戳对象
      
    '0 表示只有一张证书时也弹出证书选择框，1 表示只有一张证书时将不弹出证书选择框，默认值为0
20    mobjHNCACTRL.SetCertChooseType (1)
    '弹出HNCA颁发的签名证书列表
    If mbytFunc = 0 Then
30      lngRet = mobjHNCACTRL.SetCert("SC", "", "", "", "CN=HNCA,O=HNCA,L=ZZ,S=HN,C=CN", "")
    Else
31      lngRet = mobjHNCACTRL.SetCertEx("SC", "", "", "", "", "", "SM2")
    End If
    '判断是否初始化成功
    If lngRet <> 0 Then
        MsgBox "初始化客户端证书错误：" & mobjHNCACTRL.GetErrorMessage(), vbOKOnly, "河南CA"
        Exit Function
    End If

    '以下为业务系统中从主题(DN)项
40     strCertDN = mobjHNCACTRL.GetCertInfo(0, "")
            
    '以下为业务系统中从主题(DN)项中提取CN项

50      intCnIndex = InStr(strCertDN, "CN=") '获取"CN="项位置
51      strCertCN = Mid(strCertDN, intCnIndex) '将CN项前的信息去除
52      intCnIndex = InStr(strCertCN, ",") '第1个逗号所在位置
53      If (intCnIndex <> -1) Then
54          strCertCN = Mid(strCertCN, 1, intCnIndex - 1)
55      End If
56      strCertCN = Mid(strCertCN, 4)
        '证书CN项 姓名
57      strName = Trim(strCertCN)

        '获取证书序列号,SN项目，每个证书唯一标识
60      strCertSn = mobjHNCACTRL.GetCertInfo(2, "")
        '获取客户端签名(公钥)证书Base64编码
80      strCertPUB = mobjHNCACTRL.GetCertInfo(8, "")
        '获取扩展项（根据项目确定扩展域的参数）
81      strPara = IIf(mbytFunc = 0, "1.2.86.11.7.12", "1.2.86.11.7.11")
82      strUserID = mobjHNCACTRL.GetCertInfo(7, strPara)  '身份证号
        If mbytFunc = 0 Then
90          strPic = mobjHNCACTRL.LoadSealFromKey(strCertSn, "111111")   '返回印章数据的BASE64
        Else
91          strPara = "测试"
92          strSign = mobjHNCACTRL.SignSealForSM2(strCertSn, strPara)
93          lngRet = mobjHNCACTRL.VerifySealForSM2(strSign, strPara)
94          strPic = mobjHNCACTRL.GetSealForSM2(strSign, lngRet)
        End If
        If strPic <> "" Then
100        If mbytFunc = 0 Then strPic = mobjHNCACTRL.GetBmpFromSealData(strPic, 2) '返回印章图片数据的BASE64
120        strPicPath = SaveBase64ToFile("bmp", strCertSn, strPic) '返回印章图片数据的BASE64转换成图片文件并返回图片位置
        Else
            MsgBox "读取图片失败！请检查是否插上key盘！", vbInformation, gstrSysName
            Exit Function
        End If
    
130     GetCertList = True
        Exit Function
errH:
500     MsgBox "获取证书信息失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Private Function GetCertLogin(ByVal strCertPUB As String, ByVal strCertDN As String) As Boolean
    Dim blnRet As Boolean
    Dim strSignData As String
    Dim strUrl As String, strAPPCode As String, strRand As String, strMsg As String
    Dim strTSIP As String, strUserName As String, strUserPWD As String
    Dim lngRet As Long, intPort As Integer, intSSLPort As Integer
    Dim arrTmp As Variant
10      On Error GoTo errH
20
30       arrTmp = Split(gstrPara, "|")
        'gstrPara= http://218.28.16.104:7080/CAServer/servlet/CertChechServlet|218.28.16.104|8080|443|时间戳（0-不启用;1-启用）|签名算法(0-RSA\1-SM2)"
40      lngRet = -1
50      strUrl = CStr(Trim(arrTmp(0)))
60      strAPPCode = "1101"
70      strTSIP = CStr(Trim(arrTmp(1)))
80      intPort = CInt(arrTmp(2))
90      intSSLPort = CInt(arrTmp(3))
100     strUserName = "abc"
110     strUserPWD = "123"
        
120
        '获取随机数
130     Randomize
140     strRand = CStr(Int((9001) * Rnd + 1000))


        '随机数签名
150     If mbytFunc = 0 Then
160         strSignData = mobjHNCACTRL.AttachSign(strCertDN, strRand)
170     Else
180         strSignData = mobjHNCACTRL.SM2Sign("", strRand)
190     End If
        
200     If strSignData = "" Then
210         strMsg = mobjHNCACTRL.GetErrorMessage()  ' 获取错误信息
220         MsgBox "随机数签名失败:" + strMsg, vbOKOnly, gstrSysName
230         Exit Function
240     End If
250
260     strMsg = ""
        '发送服务器身份验证
270     lngRet = doPostEx(strUrl, strAPPCode, strRand, strSignData, strCertPUB, "GBK")

280     strMsg = ""

290     If (lngRet = 0) Then
300         blnRet = True
310     ElseIf (lngRet < 0) Then
            '随机数验签失败
320         blnRet = False
330         strMsg = "登录服务器验签失败！"
340     Else
350         blnRet = True
360         strMsg = "登录服务器验签成功！您的证书还剩：" & lngRet & " 天过期！"
370     End If
380     If strMsg <> "" Then
390        MsgBox strMsg, vbInformation, gstrSysName
400     End If
        '验证失败直接退出
410     If Not blnRet Then Exit Function
420
        '时间戳设置系统参数
430     If mblnTs Then
440         strMsg = ""
450         lngRet = mobjHNCATS.SetSystemParameter(strTSIP, intPort, intSSLPort)
460         If (lngRet = 0) Then        'MessageBox.Show("设置成功！")
                
                '2.设置哈希算法，如果不调用默认是sha1
470             Call mobjHNCATS.SetMD(1)
                
                '3.用户登录
480             If (mobjHNCATS.ISLogin = 1) Then
490                 blnRet = True
500             Else
510                 lngRet = -1
520                 lngRet = mobjHNCATS.LoginByUsername(strUserName, strUserPWD)    '
530                 If (lngRet = 0) Then
540                     blnRet = True
550                 Else
560                     blnRet = False
570                     strMsg = "登录时间戳服务器失败！错误代码为：" & lngRet
580                 End If
590             End If
600         Else
610             strMsg = "设置时间戳服务器的系统参数失败，错误代码：" & lngRet
620             blnRet = False
630         End If
                
640         If strMsg <> "" Then
650            MsgBox strMsg, vbInformation, gstrSysName
660         End If
670     End If
680
690       GetCertLogin = blnRet
700       Exit Function
errH:
710       MsgBox "登录服务器验证失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Function HNCA_Initialize() As Boolean
    '功能:创建河南CA控件对象
    
    On Error GoTo errH
    
1000
        Set mobjHNCACTRL = CreateObject("HNCACTRL.HNCACtrlCtrl.1") '签名对象
1001
        Set mobjHNCATS = CreateObject("TSAACTIVEXVER4.TsaActiveXver4Ctrl.1")  '时间戳对象
1100
        gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys) '读取URL 服务器
        'gstrPara = "http://218.28.16.104:7080/CAServer/servlet/CertChechServlet|218.28.16.104|8080|443|时间戳(0-不启用/1-启用)|签名算法(0-RSA\1-SM2)"
        If gstrPara = "" Then
            MsgBox "没有配置签名服务器地址，请到启用电子签名接口处设置。", vbInformation, gstrSysName
            Exit Function
        Else
            If UBound(Split(gstrPara, "|")) <> 5 Then
               MsgBox "签名服务器地址配置有误，请到启用电子签名接口处设置。" & _
                    vbCrLf & vbTab & "参数值格式""服务器URL|TSIP|TSPORT|SSLPORT|是否启用时间戳(0-不启用/1-启用)|签名算法(0-RSA\1-SM2)""", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        mblnTs = Val(Split(gstrPara, "|")(4)) = 1
        mbytFunc = CByte(Split(gstrPara, "|")(5))
1500
        HNCA_Initialize = True
        Exit Function

errH:
     MsgBox "创建接口部件失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbQuestion, gstrSysName
End Function

Private Sub Class_Terminate()
    Set mobjHNCACTRL = Nothing
    Set mobjHNCATS = Nothing
End Sub

Public Function HNCA_GetPara() As Boolean
'设置服务器地址
    
    On Error GoTo errH
    gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys)  '读取URLs 固定读取ZLHIS 系统默认100
    'gstrPara = "http://218.28.16.104:7080/CAServer/servlet/CertChechServlet|218.28.16.104|8080|443|时间戳(0-不启用/1-启用)|签名算法(0-RSA\1-SM2)"
    If gstrPara = "" Then gstrPara = "http://218.28.16.104:7080/CAServer/servlet/CertChechServlet|218.28.16.104|8080|443|0|0"
    If gstrPara <> "" Then
        gudtPara.strSignURL = gstrPara
    End If
    Exit Function
errH:
    MsgBox "读取参数失败！" & vbNewLine & "第" & CStr(Erl()) & "行 " & Err.Description, vbInformation, gstrSysName
End Function

Public Function HNCA_SetParaStr() As String
    HNCA_SetParaStr = gudtPara.strSignURL
End Function




