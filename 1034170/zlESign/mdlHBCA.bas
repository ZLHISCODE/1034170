Attribute VB_Name = "mdlHBCA"
Option Explicit

'�ӱ������е���ҽԺ   �ӱ�CA
Private mblnInit As Boolean         '�Ƿ��ѳ�ʼ���ɹ�
Private mCertMgr As Object          'HebcaP11XLib.certMgr
Private mSignCert As Object         'HebcaP11XLib.cert
Private mFormSeal As Object         'FormSealCtrl1 ����ǩ�¿ؼ�
Private mSVSClient As Object        'SVS_SOFT_COMLib.SvsVerify '���岢ʵ����SVS�ͻ������
Private mblnTs As Boolean           '�Ƿ�����ʱ���

Private Const M_STR_LICENCE As String = "amViY55oZWKcZmhlnWxhaGViY2GXGmJjYWhlYnGH1QQ5GcNqnW6z3vohVnE+nTJr"
Private Const M_STR_SUMMARY As String = "[SUMMARY]"

Public Function HBCA_InitObject() As Boolean
'----------------------------------------------------------------------------------------------------------------------------------
'����:��ʼ������ǩ����Ҫ�õ�����
'����:True-��ʼ���ɹ�;False-��ʼ��ʧ��
'����:��ΰ��
'����:2015-08-31
'----------------------------------------------------------------------------------------------------------------------------------
    Dim arrList As Variant
    
    If mblnInit Then HBCA_InitObject = True: Exit Function
    On Error GoTo errH
    '������Ϣ:IP|�˿ں�|�Ƿ�����ʱ���(0-������/1-����)
1   gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys, , "121.28.49.158&&&5000&&&0")  '��ȡ��������
    If gstrPara = "" Then
        Err.Raise -1, , "�����ļ���ȡʧ�ܣ��뵽���õ���ǩ���ӿڴ����á�"
        Exit Function
    End If
    arrList = Split(gstrPara, "&&&")
    If UBound(arrList) <> 2 Then
        MsgBox "ǩ����������ַ���ø�ʽ����,�뵽���õ���ǩ���ӿڴ����á�", vbOKOnly + vbInformation, gstrSysName
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
500 MsgBox "�����ӱ�CA�ӿڲ���ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & vbNewLine & _
            Err.Description, vbQuestion, gstrSysName
End Function

Public Function HBCA_RegCert(arrCertInfo As Variant) As Boolean
'----------------------------------------------------------------------------------------------------------------------------------
'����:�ṩ��HIS���ݿ���ע�����֤��ı�Ҫ��Ϣ,�����·��Ż����֤��,��Ҫ����USB-Key
'���أ�arrCertInfo��Ϊ���鷵��֤�������Ϣ
'      0-ClientSignCertCN:�ͻ���ǩ��֤�鹫������(����),����ע��֤��ʱ������֤���
'      1-ClientSignCertDN:�ͻ���ǩ��֤������(ÿ��Ψһ)
'      2-ClientSignCertSN:�ͻ���ǩ��֤�����к�(ÿ֤��Ψһ)
'      3-ClientSignCert:�ͻ���ǩ��֤������
'      4-ClientEncCert:�ͻ��˼���֤������
'      5-ǩ��ͼƬ�ļ���,�մ���ʾû��ǩ��ͼƬ
'      6-ʱ���֤��
'      7-ǩ����Ϣ
'����:��ΰ��
'����:2015-08-31
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
126     MsgBox "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
   
        
End Function

Private Function GetCertList(ByRef strName As String, ByRef strCertSn As String, ByRef strSignCert As String, _
                ByRef strUserID As String, Optional ByRef strSealBase64 As String, Optional ByRef strTSCert As String, _
                Optional ByRef strPicFile As String, Optional ByRef strPic As String) As Boolean
'----------------------------------------------------------------------------------------------------------------------------------
'����:��ȡ֤����Ϣ
'����:strName-֤���û���
'     strCertSn-֤�����к�
'     strSignCert-֤������
'     strUserID-�û�Ψһ��ʶ  ��ȡ���֤��
'     strSealBase64-ǩ��BASE64
'     strTSCert-ʱ���֤��BASE64
'���أ�
'����:��ΰ��
'����:2015-08-31
'----------------------------------------------------------------------------------------------------------------------------------
    Dim intCount As Integer
    Dim strSignData As String
    Dim strSealSn As String
    Dim objPic As Picture
    
    intCount = mFormSeal.GetSealCount()
    
    If intCount < 1 Then
        MsgBox "��������Key��", vbInformation, gstrSysName
        Exit Function
    Else
        mCertMgr.Licence = M_STR_LICENCE
        Set mSignCert = mCertMgr.SelectSignCert
        'CN=����������
        strName = mSignCert.GetSubjectItem("cn")
        'strSignCert = mSignCert.GetCertB64        '�õ�ǩ��֤������
'        strCertDN = mSignCert.GetSubjectItem("DN")
        '��ȡ����֤���Ψһ��ʶ,���ں��û�������
        strUserID = mSignCert.GetCertExtensionByOid("1.2.156.112586.1.4") '"2@6021SF0130637201507090001"
        strUserID = Mid(strUserID, 10) '��ȡ���֤��
        '��ȡ֤����Ϣ  ��֤ǩ����Ҫ֤����Ϣ��
        'ǩ��BASE64����,ǩ��֤��,ʱ���֤��BASE64����
        strSignData = mFormSeal.SignAndSealWithoutTimeStampCert("����20150901", "", 0, True, mblnTs)
        If strSignData = "" Then MsgBox "�����ǩ��ʧ�ܣ�", vbInformation, gstrSysName: Exit Function
        '��ȡǩ�µ�SN\ǩ��Bases64
        strSealSn = mFormSeal.GetSelectedSeal()
        strSealBase64 = mFormSeal.GetSeal(strSealSn) '��ȡ�µ�B64
        strPic = mFormSeal.GetSealPicFromB64(strSealBase64)
        
        strPicFile = SaveBase64ToFile("gif", strSealSn, strPic)
        '��ͼƬת����ָ��bmp��ʽ
'        Set objPic = LoadPicture(strPicFile)
'        SavePicture objPic, strPicFile
'
        '��ȡ֤�������Ϣ��֤���SN��֤�����Ч����Ҫ�������ݿ�
        strSignCert = mFormSeal.GetCert(strSealSn) '��ȡ֤��
        Set mSignCert = mCertMgr.CreateCertFromB64(strSignCert)
        strCertSn = mSignCert.GetSerialNumber '֤��SN
        'dCertDate:=mSignCert.NotAfter    ��Ч��
         'ʱ���
        If mblnTs Then
            strTSCert = mFormSeal.GetTimeStampCert '��ȡʱ���֤������
        End If
    End If
        
    GetCertList = True
    Exit Function
End Function

Public Function HBCA_Sign( _
    ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String, _
    ByRef strTimeStampCode As String, ByRef blnReDo As Boolean) As Boolean
'----------------------------------------------------------------------------------------------------------------------------------
'����:ǩ��
'����:strSource-����Դ
'     strSignData-ǩ��ֵ
'     strTimeStamp-ʱ���ֵ
'     strTimeStampCode-ʱ�����Ϣ
'���أ�True/false
'����:��ΰ��
'����:2015-08-31
'----------------------------------------------------------------------------------------------------------------------------------
        'ǩ��
        Dim strTiemRequest As String
        Dim strTmp As String
        Dim strSealSn As String
        Dim strSealBase64 As String
        Dim blnCheck As Boolean
        
        On Error GoTo errH
        blnCheck = HBCA_CheckCert(blnReDo)
        If blnReDo Then Exit Function
100     If blnCheck Then                '��֤��ǰUSB�Ƿ���ǩ���û��ģ�����ȡǩ��֤��
            '����SignAndSealWithoutTimeStampCert��ԭ�Ľ��и��£����Զ�ԭ�����ݽ�����֯��
104         strSource = mCertMgr.util.HashText(strSource, 1)
120         strSignData = mFormSeal.SignAndSealWithoutTimeStampCert(strSource, "", 0, True, mblnTs)
121         If strSignData = "" Then MsgBox "ǩ��ʧ��,ǩ��ֵΪ�գ�", vbInformation, gstrSysName: Exit Function
            If mblnTs Then
130             strTimeStampCode = mFormSeal.GetTimeStamp() '��ȡʱ�����Ϣ
140             strTimeStamp = mFormSeal.GetTimeStampInfoByB64(strTimeStampCode, "time")
                If strTimeStampCode = "" Then
                    MsgBox "ǩ��ʧ��,ʱ���B64Ϊ�գ�", vbInformation, gstrSysName
                    Exit Function
                End If
            Else
150             strTimeStamp = CStr(gobjComLib.zlDatabase.Currentdate)
            End If
        Else
            MsgBox "ǩ��ʧ�ܣ�", vbInformation, gstrSysName
            Exit Function
        End If
        If Trim(strSignData) <> "" Then strSignData = M_STR_SUMMARY & strSignData  '�˱�ʶ[SUMMARY]������֤ǩ��ʱ���ְ�ԭ����֤ǩ�����ǰ�ժҪ��֤ǩ��
        HBCA_Sign = True
        Exit Function
errH:
114     MsgBox "ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function HBCA_VerifySign(ByVal strSignData As String, ByVal strSource As String, ByVal strTimeStampCode As String, _
                            ByVal strCert As String, ByVal strTSCert As String, ByVal strSealCert As String) As Boolean
'----------------------------------------------------------------------------------------------------------------------------------
'����:��֤ǩ��
'����:
'   strSignData-ǩ��ֵ
'   strSource-Դ��
'   strTimeStampCode-ʱ�����Ϣ
'   strCert-֤������
'   strTSCert-ʱ���֤������
'   strSealCert-ǩ��֤������
'���أ�True/false
'����:��ΰ��
'����:2015-08-31
'----------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    Dim lngRet As Long
    Dim blnOk As Boolean
    On Error GoTo errH
    
    If mFormSeal Is Nothing Then Set mFormSeal = CreateObject("HebcaFormSeal.FormSealCtrl.1")
    strTmp = ""

    '����ǩ����֤
    If UCase(left(strSignData, Len(M_STR_SUMMARY))) = M_STR_SUMMARY Then
        '��ժҪǩ��ʱ��֤ǩ��ʱ��ҪȡժҪ
        strSignData = Mid(strSignData, Len(M_STR_SUMMARY) + 1)
        mCertMgr.Licence = M_STR_LICENCE
        strSource = mCertMgr.util.HashText(strSource, 1)
    End If

1    Call mFormSeal.VerifyAndShowSeal(strSealCert, strCert, strSource, 0, strSignData, IIf(mblnTs, 0, -1), strTimeStampCode, strTSCert, 0)
2    lngRet = mFormSeal.GetVerifyResult()

    If lngRet = 0 Then
        strTmp = "ǩ����֤�ɹ���"
        blnOk = True
    Else
        strTmp = "ǩ����֤ʧ�ܣ�"
        blnOk = False
    End If

    If strTmp <> "" Then
        MsgBox strTmp, vbOKOnly + vbInformation, gstrSysName
    End If
    
10    HBCA_VerifySign = blnOk
    Exit Function
errH:
104     MsgBox "��֤ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Public Function HBCA_CheckCert(ByRef blnReDo As Boolean) As Boolean
'----------------------------------------------------------------------------------------------------------------------------------
'���ܣ���ȡUSB�����豸��ʼ������¼
'����:
'   ����:blnRedo-True������ע��֤��ɹ�,False-δ����ע��֤��
'���أ�True/false
'����:��ΰ��
'����:2015-08-31
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
102             MsgBox "����δ��ʼ����"
                Exit Function
            End If
        End If
         '��ȡ֤����Ϣͬʱ���Key���Ƿ����
        If Not GetCertList(strCertName, strCertSn, strCert, strCertUserID, strSealCode, strTSCert, , strPic) Then
            HBCA_CheckCert = False: Exit Function
        End If
        'δע���ڵ�ǰ�û����µ�Key
        If mUserInfo.strUserID = "" Then
            MsgBox "�������֤��Ϊ��,����ϵ����Ա����Ա������¼�룡", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        ElseIf strCertUserID <> mUserInfo.strUserID Then
            MsgBox "�������֤�ţ�" & _
                   vbCrLf & vbTab & "��" & mUserInfo.strUserID & "��" & vbCrLf & _
                   "��ǰ֤��Ψһ��ʶ:" & _
                   vbCrLf & vbTab & "��" & strCertUserID & "��" & vbCrLf & _
                   "�û����֤���뵱ǰ֤��Ψһ��ʶ�����,����ʹ�ã�", vbInformation, gstrSysName
            Exit Function
        End If
        
        '��¼��֤
        If InStr(gstrLogins & "|", "|" & strCertSn & "|") > 0 Then '�״���֤ͨ�����´β��ڼ�����֤
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
            '�ж��Ƿ���Ҫ����ע��֤��
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
            '��ȡ�Ѿ�ע��֤�����Ч��������
                '��ȡ֤�������Ϣ��֤���SN��֤�����Ч����Ҫ�������ݿ�
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
124     MsgBox "���USBKEYʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function

Private Function GetCertLogin() As Boolean
'----------------------------------------------------------------------------------------------------------------------------------
'���ܣ���¼��֤
'����:
'���أ�True/false
'����:��ΰ��
'����:2015-08-31
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
    If mCertMgr Is Nothing Then Set mCertMgr = CreateObject("HebcaP11X.CertMgr.1") 'ʵ����P11�����CertMgr��
    If mSVSClient Is Nothing Then Set mSVSClient = CreateObject("Svs_soft_com.SvsVerify.1")
    If mFormSeal Is Nothing Then Set mFormSeal = CreateObject("HebcaFormSeal.FormSealCtrl.1")
    strText = "hebca2013" 'ԭʼ�ַ���

    mCertMgr.Licence = M_STR_LICENCE

    Set mSignCert = mCertMgr.SelectSignCert  '�õ�ǩ��֤�����
    strSignData = mSignCert.SignText(strText, 1)   '��������ǩ��,��ǩ��ֵ��ŵ�signdata
    strCertB64 = mSignCert.GetCertB64         '�õ�ǩ��֤������
    'gstrPara = "121.28.49.158&&&5000"
    arrTmp = Split(gstrPara, G_STR_SPLIT)     'IP&&&�˿ں� "121.28.49.158", 5000
    strIP = arrTmp(0): lngPort = Val(CStr(arrTmp(1)))
    lngRetVal = mSVSClient.InitialVerify(strIP, lngPort) '��ʼ��SVS�ͻ���

    Dim r As Boolean
    
    If lngRetVal < 0 Then
        MsgBox "�޷�����SVS������!", vbInformation, gstrSysName
        Exit Function
    End If
    
    lngRetVal = mSVSClient.VerifyCertSign(-1, 0, strText, Len(strText), strCertB64, strSignData, 1, lngRetVal)     '��֤
    Select Case lngRetVal
        Case 0
            strMsg = "��֤�ɹ�"
        Case 1
            strMsg = "����֤��δ��Ч!"
        Case 2
            strMsg = "����֤���Ѿ�����!"
        Case 4
            strMsg = "����֤��Ǻӱ�CA�䷢!"
        Case 1002
            strMsg = "����֤��Ǻӱ�CA�䷢!"
        Case 7
            strMsg = "����֤���Ѿ�������!"
        Case -6406
            strMsg = "ǩ����֤ʧ��,������!"
    End Select
    If strMsg <> "��֤�ɹ�" Then
        MsgBox "������Ϣ:" & strMsg, vbInformation, gstrSysName
        Exit Function
    End If
'
'     strSignData = mFormSeal.SignAndSealWithoutTimeStampCert("����20150901", "", 0, True, True)
'
'     '��ȡǩ�µ�SN\ǩ��Bases64
'     strSealSn = mFormSeal.GetSelectedSeal()
'     strSignCert = mFormSeal.GetCert(strSealSn) '��ȡ֤��
'     Set mSignCert = mCertMgr.CreateCertFromB64(strSignCert)
     strDate = mSignCert.NotAfter  '��Ч��
    If strDate <> "" Then
    '��֤�ͻ���֤����Ч��ʣ������
        intDay = CheckValidaty(CDate(strDate))
    
        If (intDay <= 30 And intDay > 0 And Not gblnShow) Then
            MsgBox "����֤�黹��" & intDay & "����ڡ�", vbOKOnly + vbInformation, gstrSysName
            gblnShow = True
        ElseIf (intDay <= 0) Then
            MsgBox "����֤���ѹ��� " & Abs(intDay) & " �졣", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    GetCertLogin = True
    Exit Function
errH:
    MsgBox "��¼��֤ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Public Function HBCA_GetPara() As Boolean
'���÷�������ַ
    Dim arrList As Variant
    
    On Error GoTo errH
    gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys)
    If gstrPara = "" Then gstrPara = "121.28.49.158&&&5000&&&0" '������Ϣ:IP&&&�˿ں�&&&�Ƿ�����ʱ���(0-������/1-����)
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
    MsgBox "��ȡ����ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Public Function HBCA_SetParaStr() As String
    HBCA_SetParaStr = gudtPara.strSIGNIP & "&&&" & gudtPara.strSignPort & "&&&" & IIf(gudtPara.blnISTS, "1", "0")
End Function

Public Sub HBCA_UnloadObj()
'----------------------------------------------------------------------------------------------------------------------------------
'����:ж�ض���
'����:��
'����:��ΰ��
'����:2015-08-31
'----------------------------------------------------------------------------------------------------------------------------------
    Set mCertMgr = Nothing
    Set mSVSClient = Nothing
    Set mSignCert = Nothing
    Set mFormSeal = Nothing
    mblnInit = False
End Sub


