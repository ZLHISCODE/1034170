Attribute VB_Name = "mdlANXIN"

Option Explicit
'���ֹ�Ͷ����CA         ��̨������ҽԺ
Private mobjJLClient As Object   'JITComVCTKEx.JITVCTKEx.1
Private mobjJLServer As Object  'JITClientCOMAPI.JITClientProc.1
Private mobjCertInfo As Object    'ǩ�� JITCertActiveX.CertInfo.1

'Private mobjJLClient As New JITComVCTKExLib.JITVCTKEx
'Private mobjJLServer As New JITClientCOMAPILib.JITClientProc
'Private mobjCertInfo As New JITCertActiveXLib.CertInfo    'ǩ��
Private mblnInit As Boolean

Private mstrPWD As String          '�������������
Private mIntPwd As Integer          '�����������8��

Private Const M_STR_PARA As String = "<?xml version=""1.0"" encoding=""gb2312""?><authinfo><liblist>" & _
        "<lib type=""CSP"" version=""1.0"" dllname="""" ><algid val=""SHA1"" sm2_hashalg=""sm3""/></lib>" & _
        "<lib type=""SKF"" version=""1.1"" dllname=""SERfR01DQUlTLmRsbA=="" ><algid val=""SHA1"" sm2_hashalg=""sm3""/></lib>" & _
        "<lib type=""SKF"" version=""1.1"" dllname=""U2h1dHRsZUNzcDExXzMwMDBHTS5kbGw="" ><algid val=""SHA1"" sm2_hashalg=""sm3""/></lib>" & _
        "<lib type=""SKF"" version=""1.1"" dllname=""QU5YSU5Dc3AxMV8zMDAwR01BLmRsbA=="" ><algid val=""SHA1"" sm2_hashalg=""sm3""/></lib>" & _
        "</liblist><checkkeytimes><item times=""3"" ></item></checkkeytimes></authinfo>"

Public Function ANXIN_InitObj() As Boolean
     '֤�鲿����ʼ��
    Dim lngRet As Long
    Dim strTSAIP As String
    Dim varTmp As Variant
100     If glngSign > 1 Then ANXIN_InitObj = True: Exit Function
        On Error Resume Next
102     If mobjJLClient Is Nothing Then
104         Set mobjJLClient = CreateObject("JITComVCTKEx.JITVCTKEx.1")
106         If Err.Number <> 0 Then
108             MsgBox "��������ǩ������JITComVCTK_S.dll��ʧ�ܣ�����ÿؼ��Ƿ�װ��ע�ᡣ", vbInformation, gstrSysName
                Exit Function
            End If
        End If
110     Err.Clear
112     If mobjJLServer Is Nothing Then
114         Set mobjJLServer = CreateObject("JITClientCOMAPI.JITClientProc.1")
116         If Err.Number <> 0 Then
118             MsgBox "��������ǩ������JITClientCOMAPI.dll��ʧ�ܣ�����ÿؼ��Ƿ�װ��ע�ᡣ", vbInformation, gstrSysName
                Exit Function
            End If
        End If
120     Err.Clear
122     If mobjCertInfo Is Nothing Then
124         Set mobjCertInfo = CreateObject("JITCertActiveX.SZZSCertInfo.1")
126         If Err.Number <> 0 Then
128             MsgBox "��������ǩ�¶���JITCertActiveX.dll��ʧ�ܣ�����ÿؼ��Ƿ�װ��ע�ᡣ", vbInformation, gstrSysName
                Exit Function
            End If
        End If
130     Err.Clear: On Error GoTo 0
        On Error GoTo errH
        '������Ϣ:�Ƿ�����ʱ���[0-������;1-����]&&&ǩ��������IP&&&ǩ���������˿�&&&ʱ���IP&&&ʱ����˿�
        '��һλ ǩ�����������ڶ�λʱ���������������λ���ء����ž���3��Ӳ�������û��Ӳ�����ǡ�000����ֻ��ǩ�����������ǡ�100��
        'Ϊ�˼�����ǰ�����һ������=0;�����ֻ��ǩ�����������ǡ�100�� ;=1��������ǩ������������ʱ��������� ����"110"����λ������ʱδ���ã�Ԥ������
        '����ʱ���������
        'gstrPara = "1&&&175.17.252.155&&&8000&&&175.17.252.156&&&8000" '��̨������ҽԺ
        'gstrPara = "000&&&175.17.252.155&&&8000&&&175.17.252.156&&&8000" '��̨�Ӹ���ҽԺ
        
132     gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys, , "") '��ȡ��������
        
134     If gstrPara = "" Then
136         Err.Raise -1, , "��ǰϵͳ��" & glngSys & "��û�����õ���ǩ������,�뵽���õ���ǩ���ӿڴ����á�"
            Exit Function
        End If
138     If UBound(Split(gstrPara, G_STR_SPLIT)) <> 4 Then
140         MsgBox "����ǩ������ֵ��������,���顣" & vbCrLf & _
                "��ǰ����ֵ:" & gstrPara & vbCrLf & _
                "��ȷ��ʽ:ǩ��������IP&&&ʱ���������IP", vbInformation, gstrSysName
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
156  MsgBox "�����ӿڲ���ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbExclamation, gstrSysName
End Function

Public Function ANXIN_Sign(ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String, ByRef strTimeStampCode As String, Optional ByRef blnReDo As Boolean) As Boolean
'ǩ��
    Dim lngRet As Long
    Dim strErr As String
    Dim strHash As String
    
    Dim blnCheck As Boolean
        On Error GoTo errH
        blnCheck = ANXIN_CheckCert(blnReDo)
        If blnReDo Then Exit Function
100     If blnCheck Then                 '��֤��ǰUSB�Ƿ���ǩ���û��ģ�����ȡǩ��֤��
            '֤��ID����ǩ��
            lngRet = mobjJLClient.SetPinCode(mstrPWD)
110         strSignData = mobjJLClient.DetachSignStr("", strSource)      'DetachSignStr-����ԭ��ǩ��;AttachSignStr-��ԭ��ǩ��
            If Not GetErrorInfo("DetachSignStr") Then Exit Function
            If strSignData <> "" Then
                If gudtPara.blnISTS Then
                    If Not ConnectToTsaServer() Then Exit Function
                    strHash = StringSHA1(strSource)
                    strTimeStampCode = mobjJLServer.TsaSign("", 1, strHash)            '����ʱ��� ����ǩ��ֵ������ǩ��ʱ�ȽϺ�ʱ,�ʲ��ù̶�ֵ
                    strTimeStamp = mobjJLServer.VerifyTsaSign(strTimeStampCode)
                    Call mobjJLServer.FinalizeServerConnectEx    '�Ͽ�ʱ�������������
                    If strTimeStampCode = "" Then MsgBox "��ȡʱ���ʧ�ܣ�", vbInformation, gstrSysName: Exit Function
                    '���ڸ�ʽ��
                    strTimeStamp = Mid(strTimeStamp, 1, 14)
                    strTimeStamp = String14ToDate(strTimeStamp, strErr)
                    If strErr <> "" Then MsgBox strErr, vbInformation, gstrSysName: Exit Function
                    'ת������ʱ��
                    strTimeStamp = Format(DateAdd("h", 8, strTimeStamp), "YYYY-MM-DD HH:MM:SS")
                Else
                    strTimeStamp = Format(gobjComLib.zlDatabase.Currentdate & "", "yyyy-MM-dd HH:mm:ss")
                End If
            Else
                MsgBox "ǩ��ʧ�ܣ�", vbInformation, gstrSysName
                Exit Function
            End If
112
        Else
            MsgBox "ǩ��ʧ�ܣ�", vbInformation, gstrSysName
            Exit Function
        End If
        ANXIN_Sign = True
        Exit Function
errH:
114     MsgBox "ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbExclamation, gstrSysName
End Function

Public Function ANXIN_VerifySign(ByVal strSource As String, ByVal strSignData As String, ByVal strTimeStampCode As String) As Boolean
'����;��֤ǩ��
'����:strSignData -ǩ��ֵ
     Dim blnRet As Boolean
     Dim lngRet As Long
     Dim strTS As String
     On Error GoTo errH
     
    If gudtPara.blnIsSign Then
        '��������ǩ
        If Not ConnectToSignServer() Then Exit Function
        lngRet = mobjJLServer.VerifyDetachedSign(strSignData, strSource) '��������֤���� ����ԭ��ǩ��:VerifyDetachedSign(string, string);��ԭ��ǩ��  VerifyAttachedSign
        If lngRet <> 0 Then
            MsgBox "ǩ����֤ʧ��:" & mobjJLServer.GetErrorMessage(lngRet), vbInformation, gstrSysName
            Call mobjJLServer.FinalizeServerConnectEx   '�Ͽ�ǩ������������
            Exit Function
        End If
        Call mobjJLServer.FinalizeServerConnectEx   '�Ͽ�ǩ������������
    Else
        Call mobjJLClient.VerifyAttachedSign(strSignData)  '�ͻ�����֤ǩ��
        lngRet = mobjJLClient.GetErrorCode()
        If lngRet <> 0 Then
            MsgBox "��֤ǩ��ʧ�ܣ������룺" & lngRet & " ������Ϣ��" & mobjJLClient.GetErrorMessage(lngRet), vbInformation, gstrSysName
            Exit Function
        End If
    End If

    '����ʱ���������
    If gudtPara.blnISTS Then
        If Not ConnectToTsaServer() Then Exit Function
        strTS = mobjJLServer.VerifyTsaSign(strTimeStampCode)
        If strTS = "" Then
              MsgBox "ʱ�����֤ʧ�ܣ�", vbInformation, gstrSysName
              Call mobjJLServer.FinalizeServerConnectEx   '�Ͽ�ʱ���������
              Exit Function
        End If
        Call mobjJLServer.FinalizeServerConnectEx   '�Ͽ�ʱ���������
    End If
    MsgBox "��֤�ɹ����õ���ǩ��������Ч!", vbInformation, gstrSysName
    
     ANXIN_VerifySign = True
     Exit Function
errH:
104     MsgBox "��֤ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbExclamation, gstrSysName
End Function


Public Function ANXIN_CheckCert(ByRef blnReDo As Boolean) As Boolean
    '���ܣ���ȡUSB�����豸��ʼ������¼
    Dim strKeySN As String, strUserID As String, strUserName As String, strCertDN As String
    Dim strDate As String
    Dim arrDN As Variant
    Dim udtUser As USER_INFO
    Dim blnRet As Boolean
    Dim i As Long
    
    On Error GoTo errH
    If Not GetCertList(strKeySN, strUserName, strCertDN, , strUserID) Then Exit Function
    If mUserInfo.strUserID = "" Then
        MsgBox "�������֤��Ϊ��,����ϵ����Ա����Ա������¼�룡", vbOKOnly + vbInformation, gstrSysName
        Exit Function
     ElseIf mUserInfo.strUserID <> strUserID Then
        MsgBox "��֤��δע�����������£�����ʹ�ã�"
        Exit Function
    End If
    
    '�ж��Ƿ���Ҫ����ע��֤��
    udtUser.strName = strUserName
    udtUser.strSignName = strUserName
    udtUser.strUserID = strUserID '���֤��
    udtUser.strCertSn = strKeySN
    udtUser.strCertDN = strCertDN
    udtUser.strCert = ""
    udtUser.strEncCert = ""
    udtUser.strCertID = ""
    udtUser.strPicPath = ""
    arrDN = Split(mUserInfo.strCertDN, ",")     'CN=����СU3294, O=��̨������ҽԺ, L=��̨����, S=������ʡ, C=CN, ��Ч����=
    For i = 0 To UBound(arrDN)
        If Trim(arrDN(i)) Like "��Ч����*" Then
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
     MsgBox "���USBKEYʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbExclamation, gstrSysName
End Function


Public Function ANXIN_RegCert(arrCertInfo As Variant) As Boolean
    '���ܣ��ṩ��HIS���ݿ���ע�����֤��ı�Ҫ��Ϣ,�����·��Ż����֤��,,��Ҫ����USB-Key
    '���أ�arrCertInfo��Ϊ���鷵��֤�������Ϣ
'      0-ClientSignCertCN:�ͻ���ǩ��֤�鹫������(����),����ע��֤��ʱ������֤���
'      1-ClientSignCertDN:�ͻ���ǩ��֤������(ÿ��Ψһ)
'      2-ClientSignCertSN:�ͻ���ǩ��֤�����к�(ÿ֤��Ψһ)
'      3-ClientSignCert:�ͻ���ǩ��֤������
'      4-ClientEncCert:�ͻ��˼���֤������
'      5-ǩ��ͼƬ�ļ���,�մ���ʾû��ǩ��ͼƬ

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
     MsgBox "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbExclamation, gstrSysName

End Function

Private Function GetCertList(Optional ByRef strUniqueID As String = "-1", Optional ByRef strName As String = "-1", Optional ByRef strCertDN As String = "-1", _
    Optional ByRef strPicPath As String = "-1", Optional ByRef strUserID As String = "-1", Optional ByRef strDate As String = "-1") As Boolean
'����:��ȡ����֤������
'strUserID-���֤��
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
    strDate = mobjJLClient.GetCertInfo("SC", 6, "")     '��Ч����
    If IsDate(strDate) Then
        '���֤���Ƿ����
        lngDay = CheckValidaty(strDate)
        If (lngDay <= 30 And lngDay > 0 And Not gblnShow) Then
            MsgBox "����֤�黹��" & lngDay & "�����", vbInformation, gstrSysName
            gblnShow = True
        ElseIf (lngDay <= 0) Then
            MsgBox "����֤���ѹ��� " & Abs(lngDay) & " ��"
            Exit Function
        End If
    End If
    If strUniqueID <> "-1" Then strUniqueID = mobjJLClient.GetCertInfo("SC", 2, "")     '֤�����к�
    If strCertDN <> "-1" Or strName <> "-1" Then
        strCertDN = mobjJLClient.GetCertInfo("SC", 0, "") 'CN=����СU3294, O=��̨������ҽԺ, L=��̨����, S=������ʡ, C=CN, ��Ч����=
        If strCertDN <> "" Then
            arrDN = Split(strCertDN, ",")
            For i = 0 To UBound(arrDN)
                If Trim(arrDN(i)) Like "CN*" Then
                    strName = Trim(Split(arrDN(i), "=")(1))
                    Exit For
                End If
            Next
        End If
        strCertDN = strCertDN & ", ��Ч����=" & strDate
    End If
    
    If strUserID <> "-1" Then
        strUserID = ""
        strTmp = mobjJLClient.GetCertInfo("SC", 7, "1.2.86.11.7.1")  '���֤����ҪתASCII :31 16 a0 14 13 12 34 33 32 35 30 33 31 39 38 36 30 31 31 32 36 32 31 35
        If Not GetErrorInfo("GetCertInfo") Then Exit Function
        If strTmp <> "" Then
            arrDN = Split(strTmp, " ")
            For i = 6 To UBound(arrDN)    'ǰ6���ַ�Ϊǰ׺
                strUserID = strUserID & Chr(Val("&H" & arrDN(i)))
            Next
        End If
    End If
    
    If mstrPWD = "" Then
CheckPWD:
        If Not frmPassword.ShowMe(mstrPWD, 6, 16) Then Exit Function
        lngRet = mobjCertInfo.VerifyUserPin("ANXIN3KGM", mstrPWD)
        'VB���Ե�ʱ�򵥲����ٷ������룻ֱ�����з�����ȷ�ַ���'{"RetryCount":"0","VerifyValue":"1"}
        '��������ϰ汾 ����1-�ɹ���0-ʧ��
        '�״���֤����
        If lngRet = 0 Then
            mIntPwd = mIntPwd - 1
            mstrPWD = ""
            If mIntPwd > 0 Then
                MsgBox "��֤����ʧ��,������" & mIntPwd & "���������Ի���!", vbInformation + vbOKOnly, gstrSysName
                GoTo CheckPWD
            Else
                MsgBox "��֤�����������,���ҹ���Ա������", vbInformation + vbOKOnly, gstrSysName
                Exit Function                     '���������������
            End If
        Else
            mIntPwd = 8
        End If
    End If

    '���Key����
    If strPicPath <> "-1" Then
        '��ȡǩ�º�ʱ,ǩ��ʱ����ȡ��ֻ��ע���ʱ���ȡ
        'strKeyCount = [{"KeyName":"���ŵ���Կ�� ","KeyType":"ANXIN3KGM","UsbKeySerialNumber":"AX00010415"},{"KeyName":"���ŵ���Կ��","KeyType":"ANXIN3KGM","UsbKeySerialNumber":"AX00010414"}]
        strKeyCount = mobjCertInfo.GetKeyCount("ANXIN3KGM")
        strRet = strKeyCount 'VB���Ե�ʱ�򵥲����ٷ������룻ֱ�����з�����ȷ�ַ���
        If strRet <> "" Then
            If UBound(Split(strRet, "},{")) = 0 Then
                strPic = mobjCertInfo.ReadImageData("ANXIN3KGM", mstrPWD)
                If Len(strPic) > 1 Then
                    strPicPath = SaveBase64ToFile("gif", strUniqueID, strPic)
                Else
                    MsgBox "��ȡǩ����Ϣʧ�ܣ�", vbInformation, gstrSysName
                    Exit Function
                End If
            ElseIf Val(strKeyCount) > 0 Then
                MsgBox "��ѡ��Ψһ��KEY�̲��룡", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If

    GetCertList = True
    Exit Function
errH:
    MsgBox "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbExclamation, gstrSysName
End Function

Public Function ANXIN_GetSeal() As String
'��ȡǩ��ͼƬ
    Dim strPicPath As String
    Call GetCertList(, , , strPicPath)
    ANXIN_GetSeal = strPicPath
End Function

Private Function GetErrorInfo(ByVal strName As String) As Boolean
    Dim lngRet As Long

    On Error GoTo errH
    lngRet = mobjJLClient.GetErrorCode  'lngRet -536870826 ���벻��;-536870823  ָ��������̫����̫��
    If lngRet <> 0 Then
        MsgBox "���ýӿڣ���" & strName & "�������,��������:" & vbCrLf & mobjJLClient.GetErrorMessage(lngRet), vbInformation, gstrSysName
        Exit Function
    End If
    GetErrorInfo = True
    Exit Function
errH:
    MsgBox "��ȡ����������" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbExclamation, gstrSysName
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
        Call mobjJLServer.FinalizeServerConnectEx '��ֹ���������ӣ����ͷ����Ӿ��
        Exit Function
    End If
    ConnectToTsaServer = True
    Exit Function
errH:
    MsgBox "����ʱ�����������" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbExclamation, gstrSysName
End Function

Private Function ConnectToSignServer() As Boolean
    Dim lngRet As Long

    On Error GoTo errH
    lngRet = mobjJLServer.InitServerConnectEx(gudtPara.strSIGNIP, CInt(gudtPara.strSignPort))
    If lngRet <> 0 Then
        MsgBox mobjJLServer.GetErrorMessage(lngRet), vbInformation, gstrSysName  '���ӷ�����ʧ��
        Exit Function
    End If
    lngRet = mobjJLServer.SetServerUriEx("/signserver/service/xml")
    If lngRet <> 0 Then
        MsgBox mobjJLServer.GetErrorMessage(lngRet), vbInformation, gstrSysName
        Call mobjJLServer.FinalizeServerConnectEx '��ֹ���������ӣ����ͷ����Ӿ��
        Exit Function
    End If
    lngRet = mobjJLServer.SetCertAliasEx("")  '���÷�����ǩ��ʱ��ǩ��֤���ʶ,��ΪĬ��֤��
    If lngRet <> 0 Then
        MsgBox mobjJLServer.GetErrorMessage(lngRet), vbInformation, gstrSysName
        Call mobjJLServer.FinalizeServerConnectEx '��ֹ���������ӣ����ͷ����Ӿ��
        Exit Function
    End If
    ConnectToSignServer = True
    Exit Function
errH:
    MsgBox "����ǩ������������" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbExclamation, gstrSysName
End Function


Public Function ANXIN_GetPara() As Boolean
    Dim arrList As Variant
    
    On Error GoTo errH
    If gstrPara = "" Then gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys, , "") '��ȡURLs �̶���ȡZLHIS ϵͳĬ��100
    If gstrPara = "" Then gstrPara = "110&&&175.17.252.155&&&8000&&&175.17.252.156&&&8000"   '��ʽ�Ƿ������豸[000-��������;111-������]&&&ǩ��������IP&&&ǩ���������˿�&&&ʱ���IP&&&ʱ����˿�
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
    MsgBox "��ȡ����ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
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




