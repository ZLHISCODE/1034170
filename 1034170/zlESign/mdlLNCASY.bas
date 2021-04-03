Attribute VB_Name = "mdlLNCASY"
Option Explicit

Private mobjUSBKEY As Object     '����ʡ����ǩ�� ���� KEYUSBKEYACTIVE.USBKeyActiveCtrl.1
Private mobjMSScriptCtl As Object    'MSScriptControl.ScriptControl.1 ΢���ṩ�ű��ؼ� �õ�javaScript��encodeURI������ȡURL��
Private mblnInit As Boolean
Private mstrLastPwd As String          '�������������
Private mintLogin As Integer

'20170817 SM2�㷨����  �����κ�
Private mbytModel           As Byte             '0-RSA�㷨;1-SM2�㷨
Private mobjKeyManager      As Object           '֤��������
Private mobjCert            As Object           '֤�����
Private mobjKeyStore        As Object           'UKey������KeyStore
Private mobjKeySealArray    As Object
Private mobjKeySeal         As Object           'ǩ����
Private mobjKeyGateOper     As Object
Private mobjKeyDetector     As Object           'JHKey.KeyDetector.1.1
Private Enum E_Model
    E_RSA = 0
    E_SM2 = 1
End Enum

Public Function LNCA_Initialize() As Boolean
    '����:��������CA�ؼ�����
    
    Dim intRet As Integer
    Dim varTmp As Variant
    
    On Error GoTo errH
   
        If mblnInit Then LNCA_Initialize = True: Exit Function
        
        gstrPara = gobjComLib.zlDatabase.GetPara(90000, glngSys) '��ȡURL ������
        'gstrPara = "http://218.25.86.214:2010/ssoworker"  '���Ե�ַ
        If gstrPara = "" Then
            MsgBox "û��������֤��������ַ���뵽�������������á�������:" & vbCrLf & vbTab & "ϵͳ��100,������90000" & _
                    vbCrLf & vbTab & "����ֵ��ʽ""http://218.25.86.214:2010/ssoworker""", vbInformation, gstrSysName
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
            Set mobjUSBKEY = CreateObject("USBKEYACTIVE.USBKeyActiveCtrl.1") 'ǩ������
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
     MsgBox "�����ӿڲ���ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function LNCA_RegCert(arrCertInfo As Variant) As Boolean
        '���ܣ��ṩ��HIS���ݿ���ע�����֤��ı�Ҫ��Ϣ,�����·��Ż����֤��,,��Ҫ����USB-Key
        '���أ�arrCertInfo��Ϊ���鷵��֤�������Ϣ
        '      0-ClientSignCertCN:�ͻ���ǩ��֤�鹫������(����),����ע��֤��ʱ������֤���
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
            
124         LNCA_RegCert = True
        End If

        Exit Function
errH:
126     MsgBox "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName

End Function

Public Function LNCA_CheckCert(ByRef blnReDo As Boolean, Optional ByRef strCert As String) As Boolean
'���ܣ���ȡUSB�����豸��ʼ������¼
    Dim strKey As String, strPIN As String, strUserName As String
    Dim strCertName As String, strCertDN As String, strPicPath As String
    Dim strCertSn As String
    Dim strCertUserID As String    '�������֤����Ϣ
    Dim strDate As String
    Dim udtUser As USER_INFO
    Dim strCertID As String
    Dim blnOk As Boolean
    Dim blnRet As Boolean
    
    On Error GoTo errH

     '��ȡ֤����Ϣͬʱ���Key���Ƿ����
    If Not GetCertList(strCertName, strCertSn, strCert, strCertDN, strPicPath, strCertUserID, strDate) Then
        LNCA_CheckCert = False: Exit Function
    End If
    
    'δע���ڵ�ǰ�û����µ�Key
    If mUserInfo.strUserID = "" Then
        MsgBox "�������֤��Ϊ��,����ϵ����Ա����Ա������¼�룡", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    ElseIf UCase(strCertUserID) <> UCase(mUserInfo.strUserID) Then
        MsgBox "�������֤�ţ�" & _
                   vbCrLf & vbTab & "��" & mUserInfo.strUserID & "��" & vbCrLf & _
                   "��ǰ֤��Ψһ��ʶ:" & _
                   vbCrLf & vbTab & "��" & strCertUserID & "��" & vbCrLf & _
                   "�û����֤���뵱ǰ֤��Ψһ��ʶ�����,����ʹ�ã�", vbInformation, gstrSysName
        Exit Function
    End If
    'CA�״�ǩ��ʱ���Զ����������
    '��¼��֤
    If InStr(gstrLogins & "|", "|" & strCertSn & "|") > 0 Then '�״���֤ͨ�����´β��ڼ�����֤
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
        '�ж��Ƿ���Ҫ����ע��֤��
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
     MsgBox "���USBKEYʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Private Function GetCertList(Optional ByRef strName As String = "-1", Optional ByRef strCertSn As String = "-1", Optional ByRef strCert As String = "-1", _
                Optional ByRef strCertDN As String = "-1", Optional strPicPath As String = "-1", _
                Optional strUserID As String = "-1", Optional strEndDate As String = "-1") As Boolean
'����:��ȡ֤����Ϣ
'-����
'    strName ֤�����������
'   strCertSN ֤��Ψһ��ʶ
'   strCert ǩ��֤��
'   strCertDN ֤��������Ϣ  ֤��ע���õ�
'   strPicPath ֤��ͼƬ����λ��

    Dim strPic As String
    Dim strMsg As String
    Dim strEnd As String
    Dim lngRet As Long
    Dim strPIN As String
    Dim i As Integer
    
    On Error GoTo errH
    If Not LNCA_Initialize() Then Exit Function

    If mbytModel = E_RSA Then
        '��������
        If mstrLastPwd <> "" Then strPIN = mstrLastPwd
        If strPIN = "" Then
            If Not frmPassword.ShowMe(strPIN) Then Exit Function
        End If
        
        If strPIN = "" Then
           MsgBox "������֤�����룡", vbOKOnly + vbInformation, gstrSysName
           Exit Function
        Else
            If mintLogin >= 8 Then
                MsgBox "�Ѿ�������" & mintLogin & "�δ������룬������������������", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
            On Error Resume Next
            lngRet = mobjUSBKEY.MNGInit(strPIN)
            If Err.Number <> 0 Then
                MsgBox "��������KEY�̣�", vbOKOnly + vbInformation, gstrSysName
                Exit Function
            End If
            Err.Clear: On Error GoTo 0
            
            If lngRet = 0 Then
               mstrLastPwd = strPIN
            Else
                mintLogin = mintLogin + 1
                MsgBox "֤��������ܲ���ȷ�����Ѿ�������" & mintLogin & "�����룬����������" & 8 - mintLogin & "��!", vbOKOnly + vbInformation, gstrSysName
                mstrLastPwd = ""
                Exit Function
            End If
        End If
        Call mobjUSBKEY.MNGLogin
    
        If strCertSn <> "-1" Then strCertSn = mobjUSBKEY.MNGGetSignCertSN 'Ψһ��ʶ��
        If strCert <> "-1" Then strCert = mobjUSBKEY.MNGGetSignCert()    '��ȡǩ��֤��
        If strName <> "-1" Then strName = mobjUSBKEY.MNGGetSignCertCN()          ''��ȡ����
        If strCertDN <> "-1" Then strCertDN = mobjUSBKEY.MNGGetSignCertDN      '����
        
        If strPicPath <> "-1" Then
            strPic = mobjUSBKEY.MNGGetSESCount '����ǩ��|������
            strPic = Split(strPic, "|")(0)
            strPic = mobjUSBKEY.MNGReadSESealByLabelEx(strPic)         '��ȡǩ��ͼƬBASE64
            If strPic <> "" Then
                strPicPath = SaveBase64ToFile("gif", strCertSn, strPic) '����ӡ��ͼƬ���ݵ�BASE64ת����ͼƬ�ļ�������ͼƬλ��
            Else
                strMsg = "��ȡǩ��ͼƬʧ�ܣ�"
                GoTo msgINFO
            End If
        End If
        If strUserID <> "-1" Then
            strUserID = mobjUSBKEY.MNGGetSignCertDN_OUa   '���֤��
            If strUserID <> "" And Len(strUserID) >= 18 Then
                strUserID = Right(strUserID, 18)
            Else
                strMsg = "��ȡ���֤����ʧ�ܣ�"
                GoTo msgINFO
            End If
        End If
        '��ȡ�ͻ���֤����Ч�ڽ�ֹʱ��
        If strEndDate <> "-1" Then
            strEndDate = mobjUSBKEY.MNGGetSignCertEndValidityTime()
            strEndDate = CDate(Format(strEndDate, "YYYY-MM-DD HH:MM:SS"))
        End If
    ElseIf mbytModel = E_SM2 Then
        lngRet = mobjKeyDetector.EnumUKey()
        If lngRet = 0 Then
            MsgBox "��������KEY�̣�", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        ElseIf lngRet = 1 Then
            Call mobjKeyManager.EnumKeyStore               '����EnumKeyStoreö�ٳ�����֤��
            'Ĭ��ѡȡSM2��keyǩ��
            lngRet = mobjKeyManager.GetCertCount()
            For i = 0 To lngRet - 1 '����GetCertCount�õ�ȫ��֤��ĸ���
                mobjCert.SetCert (mobjKeyManager.GetCert(i)) '��ѯ������е�֤��
                If mobjCert.CertUsage = 1 Then 'CertUsage == 1:ǩ��֤��,2:����֤�顣CertType == 1:RSA,2:SM2
                    Call mobjKeyManager.InitKeyStoreByIndex(i, mobjKeyStore)
                    Exit For
                End If
            Next
        Else
            Call mobjKeyManager.EnumKeyStore               '����EnumKeyStoreö�ٳ�����֤��
            '��ʾ֤���б���������û�ѡȡ֤�顣
            '��һ����������ʾ��ʾ��֤�����ͣ�1��RSA��2��SM2��3��RSA��SM2����ʾ;
            '�ڶ���������ʾ֤����;��1��ǩ����2�����ܣ�3��ǩ�����ܾ���
            '0-�û�ѡ����֤��;1-δѡ��֤��
            lngRet = mobjKeyManager.ShowCertsDlg(3, 1)
            If lngRet <> 0 Then
                strMsg = "δѡ���κ�֤�飡"
                GoTo msgINFO
            End If
            Call mobjKeyManager.GetSelectedCert(mobjCert)    '�����û���ѡ��õ�ָ����֤��
            '�����û�ѡ���֤�飬��ʼ��UKey������KeyStore ǩ��ʱֱ�ӵ���ǩ���ӿ�
            Call mobjKeyManager.InitKeyStore(mobjKeyStore)
            'mobjCert.CertType֤�����ͣ�1:RSA��2:SM2
        End If
        strName = mobjCert.CertCN
        If strCertSn <> "-1" Then strCertSn = mobjCert.CertSN
        If strCertDN <> "-1" Then strCertDN = mobjCert.CertSubject
        If strCert <> "-1" Then strCert = mobjCert.Body
        If strUserID <> "-1" Then strUserID = Mid(mobjCert.CertOuA, 2)
        If strEndDate <> "-1" Then
            strEndDate = mobjCert.CertNotAfter
        End If

        If mobjCert.CertType = 2 Then  'SM2 ����������
            '��������
            If mstrLastPwd <> "" Then strPIN = mstrLastPwd
            If strPIN = "" Then
                If Not frmPassword.ShowMe(strPIN) Then Exit Function
            End If
                
            If strPIN = "" Then
               MsgBox "������֤�����룡", vbOKOnly + vbInformation, gstrSysName
               Exit Function
            Else
                Call mobjKeyStore.SetWorkPin(strPIN)  '���Ĭ��PING�룬ʹPIN����ٵ������ﵽ��Ĭ������Ŀ�ģ�ֻ��SM2Key��Ч��RSA��PIN����ɸ��������Ҹ���ʵ�֣��޷����ƣ�
                lngRet = mobjKeyStore.SignData("123")    '0-�ɹ�;��0-ʧ��
                If lngRet = 0 Then
                    mstrLastPwd = strPIN
                Else
                    mstrLastPwd = ""
                    Exit Function
                End If
            End If
        End If
        
        'ȡӡ�£�������ѡ֤�飬��ѯ��֤������Key�������ӡ�£�����ӡ������SealArray
        If strPicPath <> "-1" Then
            Call mobjKeyManager.InitSealStore(mobjKeySealArray)
            lngRet = mobjKeySealArray.GetSealCount()         '�õ�ӡ����������ӡ�¸���
            If lngRet = 0 Then
                strMsg = "��ѡ֤���޶�Ӧ��ӡ�£�"
                GoTo msgINFO
            End If
            For i = 0 To lngRet - 1
                Call mobjKeySealArray.GetSeal(i, mobjKeySeal)     '��ӡ��������ȡ��ӡ��
                strPic = mobjKeySeal.getpic()              '�õ�ӡ��ͼƬ��base64����
                If strPic <> "" Then
                    strPicPath = SaveBase64ToFile("gif", strCertSn, strPic) '����ӡ��ͼƬ���ݵ�BASE64ת����ͼƬ�ļ�������ͼƬλ��
                    Exit For
                Else
                    strMsg = "��ȡǩ��ͼƬʧ�ܣ�"
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

500  MsgBox "��ȡ֤����Ϣʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
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
        '��ȡ�����
        strRandom = HttpPost(gudtPara.strSignURL, "cmd=getrand", responseText)   '��ȡ���������ֵ: {"ret":1,"errinfo":"","rand":"3XO9JCXVJ6LF05M51165"}
        strRandom = GetSubString(strRandom, "rand")
        
        '�����ǩ��
        strSignResult = mobjUSBKEY.MNGSignData(strRandom, Len(strRandom))      '�ؼ�ǩ��
        If strSignResult = "" Then
            MsgBox "�����ǩ��ʧ�ܣ�", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    
        '�����ǩ����֤
        strParameter = "cmd=sm2certlogin" & "&rand=" & EnCodeURL(strRandom) & "&cert=" & EnCodeURL(strSignCert) & "&signed=" & EnCodeURL(strSignResult)  '��������֤KEY ǩ�����
        strToken = HttpPost(gudtPara.strSignURL, strParameter, responseText)  '
        strRet = GetSubString(strToken, "ret")
        If strRet <> "1" Then
            MsgBox "֤���¼��֤ʧ�ܣ�", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    ElseIf mbytModel = E_SM2 Then
        'Call mobjKeyStore.SetWorkPin(mstrLastPwd)
        '       �ظ��������뵼�µ�½��֤����ֵ=9 ��ʾʧ�� ��ע��
        lngRet = mobjKeyGateOper.ReqCertLogin(mobjKeyStore, "123")  '֤���½
        If lngRet <> 0 Then
            MsgBox "֤���¼��֤ʧ�ܣ�", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
    '��ȡ�ͻ���֤����Ч�ڽ�ֹʱ��
    datEnd = CDate(strEnd)
    '��֤�ͻ���֤����Ч��ʣ������
    intDay = Int(CDbl(datEnd) - CDbl(Now))
    
    If (intDay <= 30 And intDay > 0 And Not gblnShow) Then
        MsgBox "����֤�黹��" & intDay & "����ڡ�", vbInformation + vbOKOnly, gstrSysName
        gblnShow = True
        GetCertLogin = True
    ElseIf (intDay <= 0) Then
        MsgBox "����֤���ѹ��� " & Abs(intDay) & " �졣", vbInformation + vbOKOnly, gstrSysName
        GetCertLogin = False
    End If
        
    GetCertLogin = True
    Exit Function
errH:
    MsgBox "��¼��������֤ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Public Function LNCA_Sign(ByVal strSource As String, ByRef strSignData As String, ByRef strTimeStamp As String, _
        ByRef strTimeStampCode As String, ByRef blnReDo As Boolean) As Boolean
'ǩ��
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
                '��֤��ǰUSB�Ƿ���ǩ���û��ģ�����ȡǩ��֤��
                strSource = EncodeBase64String(strSource) 'Դ���а��������ַ�����Ҫ����ת��
                strSignData = mobjUSBKEY.MNGSignData(strSource, Len(strSource))        '�ؼ�ǩ��
                If strSignData <> "" Then
                    '�洢���ط�����
                    datTime = gobjComLib.zlDatabase.Currentdate()
                    strDate = Format(datTime, "yyyyMMddhhmmss")
                    strTimeStamp = Format(datTime, "yyyy-MM-dd HH:mm:ss")
                    strParameter = "cmd=insert_sign_record" & "&appid=" & EnCodeURL("100") & "&docid=" & EnCodeURL("100") & "&docname=" & _
                    EnCodeURL("ZLHIS") & "&textinfo=" & EnCodeURL(strSource) & "&signdata=" & EnCodeURL(strSignData) & "&signcert=" & _
                    EnCodeURL(strSignCert) & "&signdate=" & EnCodeURL(strDate)
                    strRet = HttpPost(gudtPara.strSignURL, strParameter, responseText)
                    blnRet = GetSubString(strRet, "ret") = "1"
                    If Not blnRet Then strMsg = "ǩ��ʧ�ܣ�"
                Else
                    strMsg = "ǩ��ʧ�ܣ�"
                    blnRet = False
1112            End If
            ElseIf mbytModel = E_SM2 Then
                'mobjKeyStore��LNCA_CheckCert �Ѿ�ʵ����
                Call mobjKeyStore.SetWorkPin(mstrLastPwd)
                intRet = mobjKeyStore.SignData(strSource)   '0-�ɹ�;��0-ʧ��
                If intRet = 0 Then
                    strSignData = mobjKeyStore.GetSignData()               '�õ�ǩ������
                Else
                    strMsg = "ǩ��ʧ�ܣ�"
                    blnRet = False
                End If
                'ǩ����洢���ط�����
                datTime = gobjComLib.zlDatabase.Currentdate()
                strDate = Format(datTime, "yyyyMMddhhmmss")
                strTimeStamp = Format(datTime, "yyyy-MM-dd HH:mm:ss")
                intRet = mobjKeyGateOper.ReqUploadMedRecord("01", "appid", "docid", "docname", strSource, strSignData, strSignCert, strDate)
                blnRet = (intRet = 0)
                If Not blnRet Then strMsg = "ǩ��ʧ�ܣ�"
            End If
        Else
            strMsg = "ǩ��ʧ�ܣ�"
            blnRet = False
        End If
        If strMsg <> "" Then
            MsgBox strMsg, vbInformation, gstrSysName
        End If
                
        LNCA_Sign = blnRet
        Exit Function
errH:
114     MsgBox "ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
End Function

Public Function LNCA_VerifySign(ByVal strCert As String, ByVal strSignData As String, ByVal strSource As String) As Boolean
'��֤ǩ��
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
        blnRet = GetSubString(strRet, "ret") = "1"    '����ֵ=1��֤ǩ���ɹ�
    ElseIf mbytModel = E_SM2 Then
        '���� ǩ�����;ǩ��ԭ��;Ԥ��;������֤��
        lngRet = mobjKeyGateOper.ReqVerifySig(strSignData, strSource, 0, strCert)
        blnRet = lngRet = 0
    End If
    If blnRet Then    '��֤ǩ��ʧ��
        strMsg = "��֤�ɹ����õ���ǩ��������Ч��"
    Else
        strMsg = "��ǩʧ�ܣ�"
    End If
    
    If strMsg <> "" Then
        MsgBox strMsg, vbInformation, gstrSysName
    End If
        
    LNCA_VerifySign = blnRet
    
    Exit Function
errH:
104     MsgBox "��֤ǩ��ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbQuestion, gstrSysName
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
'����:�������ַ�����UTF���뷽ʽת����ʮ�����Ƶ�ת������
'˵����encodeURI-javaScript��������� ASCII ��ĸ�����ֽ��б��룬Ҳ�������Щ ASCII �����Ž��б��룺 - _ . ! ~ * ' ( )
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
'���÷�������ַ
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
    MsgBox "��ȡ����ʧ�ܣ�" & vbNewLine & "��" & CStr(Erl()) & "�� " & Err.Description, vbInformation, gstrSysName
End Function

Public Function LNCA_SetParaStr() As String
    LNCA_SetParaStr = gudtPara.strSignURL & G_STR_SPLIT & gudtPara.bytSignVersion
End Function

Private Function GetSubString(ByVal strSource As String, ByVal strNode As String) As String
'����:��ȡ�����ַ�����ĳ���ڵ�ֵ
'����:strSource -�����ַ��� {"ret":1,"errinfo":"","rand":"3XO9JCXVJ6LF05M51165"}
'    strNode-��ʶҪ��ȡ�Ľڵ�����
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





