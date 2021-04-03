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
'����������Ŀ
Private mobjHNCACTRL As Object
Private mobjHNCATS As Object
Private mblnTs As Boolean
Private mbytFunc As Byte   '0-RSA�㷨,1-SM2�㷨

Private Declare Function doPostEx Lib "HNCACertTools6.dll" (ByVal strUrl As String, ByVal strAPPCode As String, ByVal strSource As String, _
        ByVal strSignData As String, ByVal strCertPUB As String, ByVal strGBK As String) As Integer

Public Function HNCA_CheckCert(ByRef blnReDo As Boolean) As Boolean
'���ܣ���ȡUSB�����豸��ʼ������¼
    Dim strKey As String, strPIN As String, strUserName As String
    Dim strCertName As String, strCertDN As String, strPicPath As String
    Dim strCertSn As String
    Dim strCertUserID As String    '��������֤����Ϣ
    Dim strDate As String
    Dim udtUser As USER_INFO
    Dim strCert As String, strCertID As String
    Dim blnOk As Boolean
    Dim blnRet As Boolean
    
    On Error GoTo errH
1000
     '��ȡ֤����Ϣͬʱ���Key���Ƿ����
    If Not GetCertList(strCertName, strCertSn, strCert, strCertDN, strPicPath, strCertUserID) Then
        HNCA_CheckCert = False: Exit Function
    End If
    'δע���ڵ�ǰ�û����µ�Key
    If mbytFunc = 0 Then
        If mUserInfo.strUserID = "" Then
            MsgBox "��������֤��Ϊ��,����ϵ����Ա����Ա������¼�룡", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        ElseIf strCertUserID <> mUserInfo.strUserID Then
            MsgBox "��������֤�ţ�" & _
                       vbCrLf & vbTab & "��" & mUserInfo.strUserID & "��" & vbCrLf & _
                       "��ǰ֤��Ψһ��ʶ:" & _
                       vbCrLf & vbTab & "��" & strCertUserID & "��" & vbCrLf & _
                       "�û�����֤���뵱ǰ֤��Ψһ��ʶ�����,����ʹ�ã�", vbInformation, gstrSysName
            Exit Function
        End If
    Else
        If mUserInfo.strCertSn <> strCertSn Then
            MsgBox "��֤��δע�����������£�����ʹ�ã�", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
1100
    '��¼��֤
    If InStr(gstrLogins & "|", "|" & strCertSn & "|") > 0 Then '�״���֤ͨ�����´β��ڼ�����֤
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
        '�ж��Ƿ���Ҫ����ע��֤��
        udtUser.strName = strCertName
        udtUser.strSignName = strCertName
        udtUser.strUserID = strCertUserID
        udtUser.strCertSn = strCertSn
        udtUser.strCertDN = strCertDN
        udtUser.strCert = strCert
        udtUser.strPicPath = strPicPath
        udtUser.strPicCode = ""
        '��ȡ�Ѿ�ע��֤�����Ч�������� ���ڸ�ʽ:2005��12��23�� 18:09:31
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
     MsgBox "���USBKEYʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Public Function HNCA_RegCert(arrCertInfo As Variant) As Boolean
        '���ܣ��ṩ��HIS���ݿ���ע�����֤��ı�Ҫ��Ϣ,�����·��Ż����֤��,,��Ҫ����USB-Key
        '���أ�arrCertInfo��Ϊ���鷵��֤�������Ϣ
        '      0-ClientSignCertCN:�ͻ���ǩ��֤�鹫������(����),����ע��֤��ʱ������֤����
        '      1-ClientSignCertDN:�ͻ���ǩ��֤������(ÿ��Ψһ)
        '      2-ClientSignCertSN:�ͻ���ǩ��֤�����к�(ÿ֤��Ψһ)
        '      3-ClientSignCert:�ͻ���ǩ��֤������
        '      4-ClientEncCert:�ͻ��˼���֤������
        '      5-ǩ��ͼƬ�ļ���,�մ���ʾû��ǩ��ͼƬ
        
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
126     MsgBox "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName

End Function

Public Function HNCA_Sign(ByVal strCurrCertSn As String, ByVal strSource As String, ByRef strSignData As String, _
        ByRef strTimeStamp As String, ByRef strTimeStampCode As String, ByRef blnReDo As Boolean) As Boolean
        'ǩ��
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
            '��֤��ǰUSB�Ƿ���ǩ���û��ģ�����ȡǩ��֤��
50          If mbytFunc = 0 Then
60              strSignData = mobjHNCACTRL.Sign(strCurrCertSn, strSource)    'ǩ��ֵ��������Դ
70          Else
80              strSignData = mobjHNCACTRL.SM2Sign("", strSource)    'ǩ��ֵ��������Դ
90          End If
100         If strSignData <> "" Then
110             If mblnTs Then
                    '��ԭ������ʱ���
120                 lngRet = -1
130                 strRequest = mobjHNCATS.SOF_CreateTimeStampRequestByPlain(strSource)
140                 lngRet = mobjHNCATS.SOF_GetLastError()
150                 If (lngRet <> 0) Then
160                     strMsg = "����ʱ�������ʧ�ܣ�ʧ�ܴ����ǣ�" & lngRet
170                     blnRet = False
180                 Else
190                     lngRet = -1
200                     strTimeStampCode = mobjHNCATS.SOF_CreateTimeStampResponse(strRequest)
210                     lngRet = mobjHNCATS.SOF_GetLastError()
220                     If (lngRet <> 0) Then
230                         strMsg = "����ʱ�����Ӧʧ�ܣ�ʧ�ܴ����ǣ�" & lngRet
240                         blnRet = False
250                     Else
260                         strTimeStamp = mobjHNCATS.SOF_GetTimeStampInfo(strTimeStampCode, 1) '��ʽ;2015��10��20�� 02:18:45.534  GMT
270                         strTimeStamp = Format(Mid(strTimeStamp, 1, 20), "YYYY-MM-DD HH:MM:SS")
280                         blnRet = True
290                     End If
300                 End If
310             Else
320                 blnRet = True
330                 strTimeStamp = Format(gobjComLib.zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
340             End If
350         Else
360             strMsg = "ǩ��ʧ�ܣ�"
370             blnRet = False
380         End If
390     Else
400         strMsg = "ǩ��ʧ�ܣ�"
410         blnRet = False
420     End If
430     If strMsg <> "" Then
440         MsgBox strMsg, vbInformation, gstrSysName
450     End If
                
460     HNCA_Sign = blnRet
470     Exit Function
errH:
480     MsgBox "ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function HNCA_VerifySign(ByVal strCert As String, ByVal strSignData As String, ByVal strSource As String, ByVal strTimeStampCode As String) As Boolean
'��֤ǩ��
'
        Dim strSigCert As String, strMsg As String
        Dim lngRet As Long
        Dim blnRet As Boolean
10      On Error GoTo errH
     
        '��ȡ�ͻ���ǩ��(��Կ)֤��Base64���� �ͻ�����֤ǩ��,��������Դ����֤ǩ��
20      If mbytFunc = 0 Then
30          lngRet = mobjHNCACTRL.VerifySign(strSignData, strSource, strCert)
40      Else
50          lngRet = mobjHNCACTRL.SM2Verify(strSignData, strSource, strCert)
60      End If
70      If lngRet = 0 Then   '��֤ǩ���ɹ�
80          If mblnTs And strTimeStampCode <> "" Then
                '��֤ʱ���
90              lngRet = -1
100             lngRet = mobjHNCATS.SOF_VerifyTimeStampByPlain(strSource, strTimeStampCode)
110         End If
120         If (lngRet = 0) Then
130             strMsg = "��ǩ�ɹ���"
140             blnRet = True
150         Else
160             strMsg = "ʱ�����֤ʧ�ܣ�������룺" & lngRet
170             blnRet = False
180         End If
190     Else
200         strMsg = "��ǩʧ�ܣ�"
210         blnRet = False
220     End If
            
230     If strMsg <> "" Then
240         MsgBox strMsg, vbInformation, gstrSysName
250     End If
260     HNCA_VerifySign = blnRet
        
270     Exit Function
errH:
280     MsgBox "��֤ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function


Private Function GetCertList(ByRef strName As String, ByRef strCertSn As String, ByRef strCertPUB As String, _
                ByRef strCertDN As String, Optional strPicPath As String, Optional strUserID As String) As Boolean
'����:��ȡ֤����Ϣ
'-����
'    strName ֤�����������
'   strCertSN ֤��Ψһ��ʶ
'   strCertPUB ֤�鹫Կ  ��֤ǩ���õ�
'   strCertDN ֤��������Ϣ  ֤��ע���õ�
'   strPicPath ֤��ͼƬ����λ��

    Dim strCertCN As String, strPic As String
    Dim lngRet As Long, intCnIndex As Integer
    Dim strPara As String, strSign As String
    
    On Error GoTo errH

10    If mobjHNCACTRL Is Nothing Then Set mobjHNCACTRL = CreateObject("HNCACTRL.HNCACtrlCtrl.1")
      If mobjHNCATS Is Nothing Then Set mobjHNCATS = CreateObject("TSAACTIVEXVER4.TsaActiveXver4Ctrl.1")    'ʱ�������
      
    '0 ��ʾֻ��һ��֤��ʱҲ����֤��ѡ���1 ��ʾֻ��һ��֤��ʱ��������֤��ѡ���Ĭ��ֵΪ0
20    mobjHNCACTRL.SetCertChooseType (1)
    '����HNCA�䷢��ǩ��֤���б�
    If mbytFunc = 0 Then
30      lngRet = mobjHNCACTRL.SetCert("SC", "", "", "", "CN=HNCA,O=HNCA,L=ZZ,S=HN,C=CN", "")
    Else
31      lngRet = mobjHNCACTRL.SetCertEx("SC", "", "", "", "", "", "SM2")
    End If
    '�ж��Ƿ��ʼ���ɹ�
    If lngRet <> 0 Then
        MsgBox "��ʼ���ͻ���֤�����" & mobjHNCACTRL.GetErrorMessage(), vbOKOnly, "����CA"
        Exit Function
    End If

    '����Ϊҵ��ϵͳ�д�����(DN)��
40     strCertDN = mobjHNCACTRL.GetCertInfo(0, "")
            
    '����Ϊҵ��ϵͳ�д�����(DN)������ȡCN��

50      intCnIndex = InStr(strCertDN, "CN=") '��ȡ"CN="��λ��
51      strCertCN = Mid(strCertDN, intCnIndex) '��CN��ǰ����Ϣȥ��
52      intCnIndex = InStr(strCertCN, ",") '��1����������λ��
53      If (intCnIndex <> -1) Then
54          strCertCN = Mid(strCertCN, 1, intCnIndex - 1)
55      End If
56      strCertCN = Mid(strCertCN, 4)
        '֤��CN�� ����
57      strName = Trim(strCertCN)

        '��ȡ֤�����к�,SN��Ŀ��ÿ��֤��Ψһ��ʶ
60      strCertSn = mobjHNCACTRL.GetCertInfo(2, "")
        '��ȡ�ͻ���ǩ��(��Կ)֤��Base64����
80      strCertPUB = mobjHNCACTRL.GetCertInfo(8, "")
        '��ȡ��չ�������Ŀȷ����չ��Ĳ�����
81      strPara = IIf(mbytFunc = 0, "1.2.86.11.7.12", "1.2.86.11.7.11")
82      strUserID = mobjHNCACTRL.GetCertInfo(7, strPara)  '����֤��
        If mbytFunc = 0 Then
90          strPic = mobjHNCACTRL.LoadSealFromKey(strCertSn, "111111")   '����ӡ�����ݵ�BASE64
        Else
91          strPara = "����"
92          strSign = mobjHNCACTRL.SignSealForSM2(strCertSn, strPara)
93          lngRet = mobjHNCACTRL.VerifySealForSM2(strSign, strPara)
94          strPic = mobjHNCACTRL.GetSealForSM2(strSign, lngRet)
        End If
        If strPic <> "" Then
100        If mbytFunc = 0 Then strPic = mobjHNCACTRL.GetBmpFromSealData(strPic, 2) '����ӡ��ͼƬ���ݵ�BASE64
120        strPicPath = SaveBase64ToFile("bmp", strCertSn, strPic) '����ӡ��ͼƬ���ݵ�BASE64ת����ͼƬ�ļ�������ͼƬλ��
        Else
            MsgBox "��ȡͼƬʧ�ܣ������Ƿ����key�̣�", vbInformation, gstrSysName
            Exit Function
        End If
    
130     GetCertList = True
        Exit Function
errH:
500     MsgBox "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
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
        'gstrPara= http://218.28.16.104:7080/CAServer/servlet/CertChechServlet|218.28.16.104|8080|443|ʱ�����0-������;1-���ã�|ǩ���㷨(0-RSA\1-SM2)"
40      lngRet = -1
50      strUrl = CStr(Trim(arrTmp(0)))
60      strAPPCode = "1101"
70      strTSIP = CStr(Trim(arrTmp(1)))
80      intPort = CInt(arrTmp(2))
90      intSSLPort = CInt(arrTmp(3))
100     strUserName = "abc"
110     strUserPWD = "123"
        
120
        '��ȡ�����
130     Randomize
140     strRand = CStr(Int((9001) * Rnd + 1000))


        '�����ǩ��
150     If mbytFunc = 0 Then
160         strSignData = mobjHNCACTRL.AttachSign(strCertDN, strRand)
170     Else
180         strSignData = mobjHNCACTRL.SM2Sign("", strRand)
190     End If
        
200     If strSignData = "" Then
210         strMsg = mobjHNCACTRL.GetErrorMessage()  ' ��ȡ������Ϣ
220         MsgBox "�����ǩ��ʧ��:" + strMsg, vbOKOnly, gstrSysName
230         Exit Function
240     End If
250
260     strMsg = ""
        '���ͷ�����������֤
270     lngRet = doPostEx(strUrl, strAPPCode, strRand, strSignData, strCertPUB, "GBK")

280     strMsg = ""

290     If (lngRet = 0) Then
300         blnRet = True
310     ElseIf (lngRet < 0) Then
            '�������ǩʧ��
320         blnRet = False
330         strMsg = "��¼��������ǩʧ�ܣ�"
340     Else
350         blnRet = True
360         strMsg = "��¼��������ǩ�ɹ�������֤�黹ʣ��" & lngRet & " ����ڣ�"
370     End If
380     If strMsg <> "" Then
390        MsgBox strMsg, vbInformation, gstrSysName
400     End If
        '��֤ʧ��ֱ���˳�
410     If Not blnRet Then Exit Function
420
        'ʱ�������ϵͳ����
430     If mblnTs Then
440         strMsg = ""
450         lngRet = mobjHNCATS.SetSystemParameter(strTSIP, intPort, intSSLPort)
460         If (lngRet = 0) Then        'MessageBox.Show("���óɹ���")
                
                '2.���ù�ϣ�㷨�����������Ĭ����sha1
470             Call mobjHNCATS.SetMD(1)
                
                '3.�û���¼
480             If (mobjHNCATS.ISLogin = 1) Then
490                 blnRet = True
500             Else
510                 lngRet = -1
520                 lngRet = mobjHNCATS.LoginByUsername(strUserName, strUserPWD)    '
530                 If (lngRet = 0) Then
540                     blnRet = True
550                 Else
560                     blnRet = False
570                     strMsg = "��¼ʱ���������ʧ�ܣ��������Ϊ��" & lngRet
580                 End If
590             End If
600         Else
610             strMsg = "����ʱ�����������ϵͳ����ʧ�ܣ�������룺" & lngRet
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
710       MsgBox "��¼��������֤ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Public Function HNCA_Initialize() As Boolean
    '����:��������CA�ؼ�����
    
    On Error GoTo errH
    
1000
        Set mobjHNCACTRL = CreateObject("HNCACTRL.HNCACtrlCtrl.1") 'ǩ������
1001
        Set mobjHNCATS = CreateObject("TSAACTIVEXVER4.TsaActiveXver4Ctrl.1")  'ʱ�������
1100
        gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys) '��ȡURL ������
        'gstrPara = "http://218.28.16.104:7080/CAServer/servlet/CertChechServlet|218.28.16.104|8080|443|ʱ���(0-������/1-����)|ǩ���㷨(0-RSA\1-SM2)"
        If gstrPara = "" Then
            MsgBox "û������ǩ����������ַ���뵽���õ���ǩ���ӿڴ����á�", vbInformation, gstrSysName
            Exit Function
        Else
            If UBound(Split(gstrPara, "|")) <> 5 Then
               MsgBox "ǩ����������ַ���������뵽���õ���ǩ���ӿڴ����á�" & _
                    vbCrLf & vbTab & "����ֵ��ʽ""������URL|TSIP|TSPORT|SSLPORT|�Ƿ�����ʱ���(0-������/1-����)|ǩ���㷨(0-RSA\1-SM2)""", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        mblnTs = Val(Split(gstrPara, "|")(4)) = 1
        mbytFunc = CByte(Split(gstrPara, "|")(5))
1500
        HNCA_Initialize = True
        Exit Function

errH:
     MsgBox "�����ӿڲ���ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function

Private Sub Class_Terminate()
    Set mobjHNCACTRL = Nothing
    Set mobjHNCATS = Nothing
End Sub

Public Function HNCA_GetPara() As Boolean
'���÷�������ַ
    
    On Error GoTo errH
    gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys)  '��ȡURLs �̶���ȡZLHIS ϵͳĬ��100
    'gstrPara = "http://218.28.16.104:7080/CAServer/servlet/CertChechServlet|218.28.16.104|8080|443|ʱ���(0-������/1-����)|ǩ���㷨(0-RSA\1-SM2)"
    If gstrPara = "" Then gstrPara = "http://218.28.16.104:7080/CAServer/servlet/CertChechServlet|218.28.16.104|8080|443|0|0"
    If gstrPara <> "" Then
        gudtPara.strSignURL = gstrPara
    End If
    Exit Function
errH:
    MsgBox "��ȡ����ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Public Function HNCA_SetParaStr() As String
    HNCA_SetParaStr = gudtPara.strSignURL
End Function



