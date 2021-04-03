VERSION 5.00
Begin VB.Form frmUserLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����Ա��¼"
   ClientHeight    =   2205
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4170
   Icon            =   "frmUserLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   4170
   StartUpPosition =   2  '��Ļ����
   Begin VB.ComboBox cboServer 
      Height          =   300
      Left            =   1950
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   1050
      Width           =   1920
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   -360
      TabIndex        =   9
      Top             =   1455
      Width           =   5025
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "�޸�����(&M)"
      Height          =   350
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "�����˴��޸�����"
      Top             =   1710
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2865
      TabIndex        =   7
      Top             =   1710
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1755
      TabIndex        =   6
      Top             =   1710
      Width           =   1100
   End
   Begin VB.TextBox txtPassWord 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1950
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   630
      Width           =   1920
   End
   Begin VB.TextBox txtUser 
      Height          =   300
      Left            =   1950
      TabIndex        =   1
      Top             =   195
      Width           =   1920
   End
   Begin VB.Image imgFlag 
      Height          =   720
      Left            =   315
      Picture         =   "frmUserLogin.frx":1CFA
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Lbl������ 
      AutoSize        =   -1  'True
      Caption         =   "������"
      Height          =   180
      Left            =   1320
      TabIndex        =   4
      Top             =   1110
      Width           =   540
   End
   Begin VB.Label Lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "����"
      Height          =   180
      Left            =   1500
      TabIndex        =   2
      Top             =   690
      Width           =   360
   End
   Begin VB.Label Lbl�û��� 
      AutoSize        =   -1  'True
      Caption         =   "�û���"
      Height          =   180
      Left            =   1320
      TabIndex        =   0
      Top             =   255
      Width           =   540
   End
End
Attribute VB_Name = "frmUserLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'�����и�ʽ��

'zlhis.exe �˵�
'zlhis.exe �û���/����        �����������Ҫ��������ת��
'zlhis.exe �û��� ����
'zlhis.exe �û��� ���� �˵�
Private mblnFirst As Boolean  'ΪTrue��ʾ�Ѿ�������ʾ��
Private mintTimes As Integer  '��¼���Դ���
Private mblnת�� As Boolean     '��ʾ����������Ƿ�Ϊ���ݿ����룬�Ƿ���Ҫ��ת��
Private mcolServer As New Collection  '������������б�
Private mblnOk As Boolean  '�Ƿ��¼�ɹ��ɹ�
Private mblnUAAddUser As Boolean

Private mobjHttp As New XMLHTTP
Private mstrPostData As String
Private mstr���� As String
Private mstrUserURL As String
Private mstrSamlAssertion As String
Private mstrError As String
Private mblnZLUA As Boolean
Private mstrAppID As String
Private mstrZLUAUser As String
  
Private Sub cmdOK_Click()
    Dim strNote             As String
    Dim strUsername         As String
    Dim strServerName       As String
    Dim strPassword         As String
    Dim blnTransPassword    As Boolean
    Dim strError            As String
    Dim strServer As String
    
    On Error GoTo errH

    
    SetConState False
    If Not CheckInput(strUsername, strPassword, strServerName) Then
        SetConState
        Exit Sub
    End If
    mintTimes = mintTimes + 1
    
    If UCase(strUsername) = "SYS" Or UCase(strUsername) = "SYSTEM" Then
        blnTransPassword = False
    Else
        blnTransPassword = mblnת��
    End If
    
    If Not OraDataOpen(strServerName, strUsername, IIf(blnTransPassword, TranPasswd(strPassword), strPassword)) Then
        txtPassWord.Text = ""
        mblnOk = False
        If mblnZLUA = True Then mblnUAAddUser = True
        txtPassWord.SetFocus
        SetConState
        Exit Sub
    Else
        '����Ƿ��л��˷�����
        strServer = GetSetting("ZLSOFT", "ע����Ϣ\��½��Ϣ", "SERVER")
        If UCase(strServer) <> UCase(cboServer.Text) Then
            ClearComponent
        End If
        
        gobjRelogin.DBUser = UCase(strUsername)
        If strUsername = strPassword Then
            MsgBox "��¼�û�����������ͬ��������ϵͳ��ȫҪ�����������޸����롣", vbInformation, gstrSysName
            cmdModify_Click
            SetConState
            Exit Sub
        End If
        '������븴�Ӷ��Ƿ����Ҫ��
        If Not CheckPWDComplex(gcnOracle, strPassword) Then
            cmdModify_Click
            SetConState
            Exit Sub
        End If
        
        '�Ƿ������������
        If CheckPwdExpiry = True Then
            cmdModify_Click
            SetConState
            Exit Sub
        End If
    End If
    '����SQL Trace
    '-----------------------------------------------
    strNote = SetSQLTrace(strServerName)
    If strNote <> "" Then
        MsgBox "������SQL Trace����!" & vbCrLf & "���ٽ���ļ�:" & strNote & vbCrLf & _
                "�����Oracle������udumpĿ¼��,����100M��ֹͣд��.", vbInformation, "��ʾ"
    End If
    If UCase(strServerName) = "RBO" Then
        SetRunWithRBO
    End If
    '�ӿڵ��ã��ŵ�Trace����֮��
    '-----------------------------------------------
    '1.���������¼���ZLUA�˻�
    If mblnUAAddUser = True And mstrUserURL <> "" Then
        mstr���� = SoapEnvelope("AddUserAppInfo", mstrZLUAUser, mstrAppID, txtUser.Text & "/" & txtPassWord.Text & "@" & cboServer.Text, mstrSamlAssertion)
        Call PostData(mstrUserURL, "AddUserAppInfo", mstr����, 5)
        mblnUAAddUser = False
    End If
    
    '2.�°没�����Զ��������򡢵���̨����Ҫ���û���������(�û���������룬zlbrw�����л�ʹ��)
    gobjRelogin.InputUser = strUsername
    gobjRelogin.InputPwd = strPassword
    gobjRelogin.ServerName = strServerName
    gobjRelogin.IsTransPwd = blnTransPassword
    '�޸�ע���
    SaveSetting "ZLSOFT", "ע����Ϣ\��½��Ϣ", "USER", strUsername
    SaveSetting "ZLSOFT", "ע����Ϣ\��½��Ϣ", "SERVER", strServerName
    
    mblnOk = True
    Unload Me
    Exit Sub
errH:
    If mintTimes > 3 Then
        MsgBox "�������ε�¼ʧ�ܣ�ϵͳ���Զ��˳�", vbInformation, gstrSysName
        cmdCancel_Click
    Else
        MsgBox Err.Description, vbInformation, gstrSysName
        SetConState
    End If
End Sub

Private Sub SetRunWithRBO()
'���ܣ���ǰ�Ự��RBO�Ż���ģʽ����SQL���
    Dim strSQL As String
    strSQL = "alter session set optimizer_mode=rule"
    On Error Resume Next
    gcnOracle.Execute strSQL
    If Err.Number = 0 Then
        MsgBox "�����õ�ǰ�Ự��RBO�Ż���ģʽ���У�", vbInformation, gstrSysName
    End If
End Sub


Private Function SetSQLTrace(ByVal strServerName As String) As String
'����:����100046�¼�����SQL Trace����
'����:Trc�ļ���
    Dim strSQL As String, strLevel As String, strFile As String
    Dim rsTmp As ADODB.Recordset
    
    strServerName = UCase(strServerName)
    
    If strServerName Like "SQLTRACE*" Then
        On Error Resume Next
        strSQL = "alter session set timed_statistics=true"
        gcnOracle.Execute strSQL
        strSQL = "alter session set max_dump_file_size='100M'"
        gcnOracle.Execute strSQL
        Err.Clear
        
        '����Trc�ļ�����
        strFile = GetTrcFile(gobjRelogin.DBUser)
        
        strLevel = "12"
        If Replace(strServerName, "SQLTRACE", "") = "4" Then
            strLevel = "4"
        ElseIf Replace(strServerName, "SQLTRACE", "") = "8" Then
            strLevel = "8"
        ElseIf Replace(strServerName, "SQLTRACE", "") = "12" Then
            strLevel = "12"
        End If
        strSQL = "alter session set events '10046 trace name context forever ,level " & strLevel & "'"
        gcnOracle.Execute strSQL
        If Err.Number = 0 Then
            SetSQLTrace = strFile
            
            strSQL = "Select 1 From zlreginfo Where ��Ŀ='TRACE�ļ�'"
            Set rsTmp = gcnOracle.Execute(strSQL)
            
            If rsTmp.RecordCount > 0 Then
                strSQL = "Update zlreginfo Set ���� ='" & strFile & "' Where ��Ŀ='TRACE�ļ�'"
            Else
                strSQL = "Insert Into zlreginfo (��Ŀ,����) Values ('TRACE�ļ�','" & strFile & "')"
            End If
            gcnOracle.Execute strSQL

        End If
    End If
End Function


Private Function GetTrcFile(ByVal strUsername As String) As String
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim strFile As String
        
    On Error Resume Next
    strFile = "ZL_" & strUsername
    strSQL = "alter session set tracefile_identifier='" & strFile & "'"
    gcnOracle.Execute strSQL
    If Err.Number <> 0 Then     '��������,˵������traceidʧ��,����Ĭ�ϵ�traceFile����
        strFile = ""
    Else
        strFile = "_" & strFile
    End If
    
    strSQL = "Select Lower(Sys_Context('userenv', 'instance_name')) || '_ora_' || p.Spid || '" & "_" & strFile & ".trc' As Trace_File" & vbNewLine & _
                    "From V$session S, V$process P" & vbNewLine & _
                    "Where s.Paddr = p.Addr And s.Sid = Userenv('sid') And s.Audsid = Userenv('Sessionid')"
    Set rsTmp = OpenSQLRecord(strSQL, "��ȡTraceFile����")
    
    If rsTmp.RecordCount > 0 Then
        GetTrcFile = rsTmp!Trace_File
    End If
    
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdModify_Click()
    Dim strUsername As String
    Dim strPassword As String
    Dim strServerName As String
    Dim strNote As String
    
    On Error GoTo InputError
    '------�����û��Ƿ�oracle�Ϸ��û�----------------
    strUsername = Trim(txtUser.Text)
    strPassword = Trim(txtPassWord.Text)
    strServerName = Trim(cboServer.Text)
    
    '��Ч�ַ���Ч��
    If Len(Trim(txtUser.Text)) = 0 Then
        strNote = "�������û���"
        txtUser.SetFocus
        GoTo InputError
    End If
    
    If Len(strUsername) <> 1 Then
        If Mid(strUsername, 1, 1) = "/" Or Mid(strUsername, 1, 1) = "@" Or Mid(strUsername, Len(strUsername) - 1, 1) = "/" Or Mid(strUsername, Len(strUsername) - 1, 1) = "@" Then
            txtUser.SetFocus
            strNote = "�û�������"
            SetConState
            Exit Sub
        End If
    End If
    
    If Trim(strPassword) <> "" And Len(strPassword) <> 1 Then
        If Mid(strPassword, Len(strPassword) - 1, 1) = "/" Or Mid(strPassword, Len(strPassword) - 1, 1) = "@" Or Mid(strPassword, 1, 1) = "/" Or Mid(strPassword, 1, 1) = "@" Then
            If txtPassWord.Enabled Then txtPassWord.SetFocus
            strNote = "�������"
            GoTo InputError
        End If
    End If
    If Trim(strServerName) <> "" Then
        If Mid(strServerName, Len(strServerName) - 1, 1) = "/" Or Mid(strServerName, Len(strServerName) - 1, 1) = "@" Or Mid(strServerName, 1, 1) = "/" Or Mid(strServerName, 1, 1) = "@" Then
            strNote = "�������Ӵ�����"
            cboServer.SetFocus
            GoTo InputError
        End If
    End If
    
    '�����ַ���
    Dim intPos As Integer
    intPos = InStr(strUsername, "@")
    If intPos > 0 Then
        strServerName = Mid(strUsername, intPos + 1)
        strUsername = Mid(strUsername, 1, intPos - 1)
    End If
    
    intPos = InStr(strUsername, "/")
    If intPos > 0 Then
        strPassword = Mid(strUsername, intPos + 1)
        strUsername = Mid(strUsername, 1, intPos - 1)
    End If
    
    intPos = InStr(strPassword, "@")
    If intPos > 0 Then
        strServerName = Mid(strPassword, intPos + 1)
        strPassword = Mid(strPassword, 1, intPos - 1)
    End If
    
    If FrmChangePass.ShowMe(Me, strUsername, strPassword, strServerName, mblnת��) Then
        txtPassWord.Text = strPassword
        cboServer.Text = strServerName
        If cmdOK.Enabled Then Call cmdOK_Click
    Else
        txtPassWord.SetFocus
    End If
    Exit Sub
InputError:
    If strNote <> "" Then
        MsgBox strNote, vbInformation, gstrSysName
    Else
        MsgBox Err.Description, vbInformation, gstrSysName
    End If
End Sub

Private Sub Form_Activate()
    Dim LngStyle As Long
    
    If mblnFirst = False Then
        
        If InStr(gstrCommand, "=") <= 0 And InStr(gstrCommand, "&") <= 0 Then
            '���õ�ǰ��������������ʾ
            LngStyle = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
            LngStyle = LngStyle Or WinStyle
            Call SetWindowLong(Me.hwnd, GWL_EXSTYLE, LngStyle)
            
            ShowWindow Me.hwnd, 0 '������
            ShowWindow Me.hwnd, 1 '����ʾ
        
            If Trim(txtUser.Text) = "" Then
                cmdOK.Default = False
                txtUser.SetFocus
            Else
                txtPassWord.SetFocus
            End If
        End If
        
        mblnFirst = True
        If Trim(txtUser.Text) <> "" And Trim(txtPassWord.Text) <> "" Then Call cmdOK_Click
    End If
    If InStr(gstrCommand, "=") > 0 And InStr(gstrCommand, "&") = 0 Then Me.Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl.Name = "txtPassWord" Then
            Call cmdOK_Click
        Else
            SendKeys "{Tab}"
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim ArrCommand
    Dim i As Integer
    
    HookDefend txtPassWord.hwnd
    Call LoadServer
    
    On Error GoTo errH
    txtUser.Text = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ\��½��Ϣ", Key:="USER", Default:="")
    cboServer.Text = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ\��½��Ϣ", Key:="SERVER", Default:="")
    
    Call ApplyOEM_Picture(Me, "Icon")
    
    If InStr(gstrCommand, "=") > 0 And InStr(gstrCommand, "&") = 0 Then
        Me.Hide
    Else
        '������һ��Ļ�����������ʾfrmSplash���壬�ڿ������뷨������£�����Դ���򣬲�����ʾ��¼���ڣ�VBֻ���쳣��ֹ�˳�
        SetActiveWindow Me.hwnd
    End If
        
    '��������в��������û��������룬����䲢ִ��
    If gstrCommand <> "" And InStr(gstrCommand, "&") = 0 Then
        ArrCommand = Split(gstrCommand, " ")
        If UBound(ArrCommand) >= 1 Then
            If InStr(ArrCommand(0), "=") <= 0 Then
                Me.txtUser.Text = ArrCommand(0)
                Me.txtPassWord.Text = ArrCommand(1)
            End If
        ElseIf UBound(ArrCommand) = 0 Then
            '�������/����ʾͬʱ�������û��������룬�������벻��Ҫ����ת��
            If InStr(1, ArrCommand(0), "/") <> 0 And InStr(1, ArrCommand(0), ",") = 0 Then
                Me.txtUser.Text = Split(ArrCommand(0), "/")(0)
                Me.txtPassWord.Text = Split(ArrCommand(0), "/")(1)
                mblnת�� = False
            End If
        End If
    End If
    Exit Sub
errH:
    If CStr(gstrCommand) <> "" Then MsgBox CStr(Erl()) & "�г��ִ������ֶ���¼��" & vbNewLine & Err.Description, vbQuestion
End Sub

Private Sub GetFocus(ByVal TxtBox As TextBox)
    With TxtBox
        If Trim(TxtBox.Text) = "" Then Exit Sub
        .SelStart = 0
        .SelLength = Len(TxtBox.Text)
    End With
End Sub

Private Sub cboServer_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        '�س������д���
        If KeyAscii <> vbKeyBack Then
            Call AppendText(KeyAscii)
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjHttp = Nothing
    Set mcolServer = Nothing
End Sub

Private Sub txtUser_Change()
    If Not mblnFirst Then Exit Sub
    cmdOK.Default = False
End Sub

Private Sub txtUser_GotFocus()
    If Me.ActiveControl Is txtUser Then
        OpenImeByName
        GetFocus txtUser
    End If
End Sub

Private Sub txtPassWord_GotFocus()
    GetFocus txtPassWord
End Sub

Private Sub cboServer_GotFocus()
    If Me.ActiveControl Is cboServer Then
        OpenImeByName
        If Trim(cboServer.Text) <> "" Then
            With cboServer
                .SelStart = 0
                .SelLength = Len(.Text)
            End With
        End If
    End If
End Sub

Private Sub SetConState(Optional ByVal BlnState As Boolean = True)
    cmdCancel.Enabled = BlnState
    cmdModify.Enabled = BlnState
    cmdOK.Enabled = BlnState
End Sub

Private Sub LoadServer()
'���ܣ��������صķ������б�
    Dim strPath As String, strFile As String, lngFile As Integer
    Dim strLine As String, lngPos As Long
    Dim strServer As String, strComputer As String, strSID As String
    Dim arrTmp As Variant
    Dim rsOraHome As ADODB.Recordset
    Dim intVersion As Integer, intTimes As Integer, intServer As Integer
    Dim i As Long, blnRead As Boolean
    Dim lngBeforeNum As Long, lngAfterNum As Long
    Dim lngFirstPos As Long, lngLastPos As Long
    Dim strChr As String, arrSer() As String
    
    cboServer.Clear
    strPath = Environ("TNS_ADMIN")
    If strPath <> "" Then
        strFile = strPath & "\tnsnames.ora" 'Oracle 8i����
        If Dir(strFile) = "" Then
            strFile = strPath & "\NET80\ADMIN\tnsnames.ora" 'Oracle 8
        End If
        If Not gobjFile.FileExists(strFile) Then strFile = ""
    End If
    
    If strFile = "" Then
        Set rsOraHome = New ADODB.Recordset
        With rsOraHome
            .fields.Append "Name", adVarChar, 256 'Name
            .fields.Append "VerSion", adInteger  '�汾
            .fields.Append "Times", adInteger '�ڼ��ΰ�װ
            .fields.Append "Server", adInteger '1-������,2-�ͻ���
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open
            '1:��ȡ64λ��32Ŀ¼���Զ���λ��SOFTWARE\Wow6432Node\Oracle 2����ȡ32λ��32λĿ¼
            arrTmp = GetAllSubKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Oracle")
            If TypeName(arrTmp) = "Empty" Then
                If Is64bit Then
                    cboServer.ToolTipText = "û���ҵ�ע�����HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Oracle��"
                Else
                    cboServer.ToolTipText = "û���ҵ�ע�����HKEY_LOCAL_MACHINE\SOFTWARE\Oracle��"
                End If
            Else
                For i = LBound(arrTmp) To UBound(arrTmp)
                    If UCase(arrTmp(i)) Like "KEY_ORA*HOME*" Then
                        intVersion = 0: intTimes = 0:  intServer = 1
                        If GetOraInfoByRegKey(arrTmp(i), intVersion, intTimes, intServer) Then
                            .AddNew Array("Name", "VerSion", "Times", "Server"), Array("\" & arrTmp(i), intVersion, intTimes, intServer)
                            .Update
                        End If
                    End If
                Next
                If UBound(arrTmp) <> -1 Then ''����Ŀ¼������Oracle_Home��Ϣ��Ĭ�϶�ȡ���
                    .AddNew Array("Name", "VerSion", "Times", "Server"), Array("", 0, 0, 1): .Update
                End If
                .Sort = "VerSion Desc,Times Desc,Server"
                Do While Not .EOF
                    strPath = ""
                    strPath = GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Oracle" & !Name, "ORACLE_HOME")
                    If strPath = "" And !Name & "" = "" Then
                        Call GetKeyValue(HKEY_LOCAL_MACHINE, "SOFTWARE\Oracle", "ORA_CRS_HOME")
                    End If
                    If strPath <> "" Then
                        strFile = strPath & "\network\ADMIN\tnsnames.ora" 'Oracle 8i����
                        If Dir(strFile) <> "" Then Exit Do
                        strFile = strPath & "\NET80\ADMIN\tnsnames.ora" 'Oracle 8
                        If Not gobjFile.FileExists(strFile) Then Exit Do
                    End If
                    strFile = ""
                    .MoveNext
                Loop
            End If
        End With
    End If
    If strFile = "" Then Exit Sub
    
    cboServer.ToolTipText = "�������б���Դ:" & strFile
    lngFile = FreeFile()
    Open strFile For Input Access Read As lngFile
    Set mcolServer = Nothing
    
    '��ȡtnsnames.ora�ļ��е�����
    Do Until EOF(lngFile)
        Input #lngFile, strLine
        strLine = ConvertStr(strLine)
         If strLine <> "" And Left(strLine, 1) <> "#" Then  '���к�ע���в�ȡ
            strServer = strServer & strLine
         End If
    Loop
    
    lngPos = 1
    Do While lngPos <> Len(strServer)   'ѭ��ÿһ���ַ�
        lngPos = lngPos + 1
        strChr = Mid(strServer, lngPos, 1)
            
        If strChr = "(" Then
            If lngFirstPos = 0 Then
                lngFirstPos = lngPos    'ȡ��һ�������ŵ�λ����Ϊ���ŵĿ�ʼλ��
            End If
            
            lngBeforeNum = lngBeforeNum + 1
        ElseIf strChr = ")" Then
            lngAfterNum = lngAfterNum + 1
        End If
        
        '��������( �ͷ����� )�ĸ������,˵��ǰ��������ƥ��,����ɾ�������е�����
        If lngBeforeNum = lngAfterNum And lngBeforeNum <> 0 Then
            lngLastPos = lngPos 'ȡ���һ��λ����Ϊ�����ŵ���ֹλ��
            strServer = Replace(strServer, Mid(strServer, lngFirstPos, lngLastPos - lngFirstPos + 1), "")   'ȥ�������м������
            lngPos = 1
            lngBeforeNum = 0: lngAfterNum = 0
            lngFirstPos = 0: lngLastPos = 0
        End If
    Loop
    Close #lngFile
    
    If InStr(1, strServer, "(") > 0 Or InStr(1, strServer, ")") = 0 Then '
        arrSer = Split(strServer, "=")
        For i = 0 To UBound(arrSer)
            If arrSer(i) <> "" Then
                mcolServer.Add Array(arrSer(i), strComputer, strSID)
                cboServer.AddItem arrSer(i)
            End If
        Next
    End If
End Sub
Private Function GetOraInfoByRegKey(ByVal strOraHome As String, ByRef intVer As Integer, ByRef intTimes As Integer, ByRef intServer As Integer) As Boolean
'����:ͨ��OracleHome����ȡOracle��Ϣ
    Dim arrTmp As Variant
    Dim i As Long, blnRetrun As Boolean
    'KEY_OraDb11g_home1_32bit
    'Key_Ora*�汾Home_32Bit
    'Key_Ora*�汾_Home*
    arrTmp = Split(UCase(strOraHome), "_")
    For i = 1 To UBound(arrTmp)
        If arrTmp(i) Like "HOME*" Then
            intTimes = ValEx(arrTmp(2))
            blnRetrun = True
        ElseIf arrTmp(i) Like "*HOME*" Then
            intTimes = Val(Mid(arrTmp(1), InStr(UCase(arrTmp(1)), "HOME") + 4))
            blnRetrun = True
        End If
        If arrTmp(i) Like "ORADB*" Then
            intVer = ValEx(Mid(arrTmp(1), 6))
            intServer = 1
            blnRetrun = True
        ElseIf arrTmp(i) Like "ORACLIENT*" Then
            intVer = ValEx(Mid(arrTmp(1), 10))
            intServer = 2
            blnRetrun = True
        ElseIf arrTmp(i) Like "*CLIENT*" Then
            intServer = 2
            intVer = ValEx(arrTmp(i))
            blnRetrun = True
        End If
    Next
    GetOraInfoByRegKey = blnRetrun
End Function

Private Sub AppendText(KeyAscii As Integer)
'���ܣ���TextBox�ؼ���Text׷�����ݣ������ݵ�ǰText��ֵ���б��м������õ�������Ŀ
'������KeyAscii    ��ǰ�İ���
    Dim strTemp As String
    Dim strInput As String
    Dim lngIndex As Long, lngStart As Long
    Dim varItem As Variant
    
    '���ȵ�ǰ�û�������ַ�
    If KeyAscii < 0 Or InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.", UCase(Chr(KeyAscii))) > 0 Then
        '�����ַ�ֻ�������֡�Ӣ�ĺͺ���
        strInput = Chr(KeyAscii)
        KeyAscii = 0
    End If
    
    With cboServer
        '��¼�ϴεĲ����λ��
        lngStart = .SelStart + IIf(strInput <> "", 1, 0)
        '���ŵõ��û�������ɺ��ı����г��ֵ�����
        strInput = Mid(.Text, 1, .SelStart) & strInput & Mid(.Text, .SelStart + .SelLength + 1)
    End With
    '���ݼ�������ݵõ����ܵ��б���
    strTemp = ""
    For Each varItem In mcolServer
        If UCase(varItem(0)) Like UCase(strInput & "*") Then
            strTemp = varItem(0)
        End If
    Next
    If strTemp <> "" Then
        cboServer.Text = strTemp
        cboServer.SelStart = Len(strInput)
        cboServer.SelLength = 100
    Else
        cboServer.Text = strInput
        cboServer.SelStart = lngStart
    End If

End Sub

Private Sub ClearComponent()
'���ܣ�--���ע���[��������]--��Ϊ��ͬ�����ݿ����ʹ�õ�ϵͳ�Ͱ汾��ͬ
    If mblnFirst = True Then '����ʱ�Կؼ��ĸ�ֵ����������
        SaveSetting "ZLSOFT", "ע����Ϣ", "��������", ""
    End If
End Sub

Private Function ReadINIToRec(ByVal strFile As String) As ADODB.Recordset
'���ܣ���ָ��INI�����ļ������ݶ�ȡ����¼����
'���أ�Nothing�����"��Ŀ,����"�ļ�¼��,����ͬһ��Ŀ�����ж�������
    Dim rsTmp As New ADODB.Recordset
    Dim objINI As TextStream
    
    Dim strItem As String, strText As String
    Dim strLine As String
            
    rsTmp.fields.Append "��Ŀ", adVarChar, 200
    rsTmp.fields.Append "����", adVarChar, 200
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    Set objINI = gobjFile.OpenTextFile(strFile, ForReading)
    Do While Not objINI.AtEndOfStream
        strLine = Replace(objINI.ReadLine, vbTab, " ")
        strItem = Trim(Mid(strLine, InStr(strLine, "[") + 1, InStr(strLine, "]") - InStr(strLine, "[") - 1))
        strText = Trim(Mid(strLine, InStr(strLine, "]") + 1))
        If strItem <> "" And strText <> "" Then
            rsTmp.AddNew
            rsTmp!��Ŀ = strItem
            rsTmp!���� = strText
            rsTmp.Update
        End If
    Loop
    
    objINI.Close
    
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    
    Set ReadINIToRec = rsTmp
End Function


Private Function SoapEnvelope(ByVal strMethod As String, ByVal parm1 As String, ByVal parm2 As String, ByVal parm3 As String, ByVal samlAssertion As String) As String
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strEnvelope As String
    
    SoapEnvelope = strEnvelope

    On Error GoTo Errhand
    
    strEnvelope = ""
    
    strEnvelope = strEnvelope & "<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:Item=""http://tempuri.org/"">"
    
    If samlAssertion <> "" Then
        strEnvelope = strEnvelope & "<soapenv:Header>"
        strEnvelope = strEnvelope & "<wsse:Security xmlns:wsu=""http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd"" xmlns:wsse=""http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd"">"
        strEnvelope = strEnvelope & samlAssertion
        strEnvelope = strEnvelope & "</wsse:Security>"
        strEnvelope = strEnvelope & "</soapenv:Header>"
    End If
    
    strEnvelope = strEnvelope & "<soapenv:Body>"
    strEnvelope = strEnvelope & "<Item:" & strMethod & ">"
    Select Case strMethod
    Case "GetSAMLResponseByArtifact"
        strEnvelope = strEnvelope & "<Item:artifact>" & parm1 & "</Item:artifact>"
    Case "AddUserAppInfo"
        strEnvelope = strEnvelope & "<Item:account>" & parm1 & "</Item:account>"
        strEnvelope = strEnvelope & "<Item:appID>" & parm2 & "</Item:appID>"
        strEnvelope = strEnvelope & "<Item:appInfo>" & parm3 & "</Item:appInfo>"
    End Select
    strEnvelope = strEnvelope & "</Item:" & strMethod & ">"
    strEnvelope = strEnvelope & "</soapenv:Body>"
    strEnvelope = strEnvelope & "</soapenv:Envelope>"
    
    
    SoapEnvelope = strEnvelope
   
    Exit Function
Errhand:
    
End Function

Private Function PostData(ByVal strPostURL As String, _
                        ByVal strMethod As String, _
                        ByVal strPostContent As String, _
                        Optional ByVal intSendWaitTime As Integer = 30) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngWaitTimeOut As Long
    Dim oXmlDoc As Object
    Dim strPostCookie As String
    
    On Error GoTo Errhand
        
    If UCase(Left(strPostURL, 4)) <> "HTTP" Then strPostURL = "http://" & strPostURL
    strPostCookie = "ASPSESSIONIDAQACTAQB=HKFHJOPDOMAIKGMPGBJJDKLJ;"
    
    strPostCookie = Replace(strPostCookie, Chr(32), "%20")
    With mobjHttp
        Call .Open("POST", strPostURL, True)
        Select Case strMethod
        Case "GetSAMLResponseByArtifact"
            Call .setRequestHeader("SOAPAction", "http://tempuri.org/ISSOService/GetSAMLResponseByArtifact")
        Case "AddUserAppInfo"
            Call .setRequestHeader("SOAPAction", "http://tempuri.org/IAccountService/AddUserAppInfo")
        End Select
        Call .setRequestHeader("Content-Length", LenB(strPostContent))
        Call .setRequestHeader("Content-Type", "text/xml; charset=utf-8")
        Call .send(strPostContent)
    End With
    lngWaitTimeOut = 0
'    lngSecondNumber = 30 '��ʱ������
    Do
        DoEvents
        Call Wait(10)
        lngWaitTimeOut = lngWaitTimeOut + 1
    Loop Until (mobjHttp.readyState = 4 Or lngWaitTimeOut >= 100 * intSendWaitTime)
    
    If mobjHttp.readyState = 4 Then
        Set oXmlDoc = CreateObject("MSXML2.DOMDocument")

        oXmlDoc.Load mobjHttp.ResponseXML
        If oXmlDoc.xml = "" Then
            mstrError = mobjHttp.responseText
            PostData = False
        Else
            mstrPostData = oXmlDoc.xml
            PostData = True
        End If
    Else
        mstrError = mobjHttp.responseText
        PostData = False
    End If
    Exit Function
    
Errhand:
    mstrError = Err.Description
End Function


Private Sub Wait(tt)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim t, t1, t2, i
    t = tt
    If t > 10 Then
        t1 = Int(t / 10)
        t2 = t - t1 * 10
        For i = 1 To t1
            Call Sleep(10)
            DoEvents
        Next i
        If t2 > 0 Then Call Sleep(t2)
    Else
        If t > 0 Then Call Sleep(t)
    End If
End Sub

Private Sub ClearValues()
    '�������
    mblnFirst = False
    mintTimes = 1
    mblnת�� = True
    mblnOk = False
    mblnUAAddUser = False
    
    mstrPostData = ""
    mstr���� = ""
    mstrUserURL = ""
    mstrSamlAssertion = ""
    mstrError = ""
    mblnZLUA = False
    mstrAppID = ""
    mstrZLUAUser = ""
End Sub

Public Function ShowMe() As Boolean
    '�������
    Call ClearValues
    Me.Show vbModal
    ShowMe = mblnOk
End Function

Public Function Docmd(ByVal strCmd As String) As Boolean
    Dim ArrCommand
    Dim ArrCommandPortal
    Dim objSoap As Object
    Dim objDoc As Object
    Dim rsIni As ADODB.Recordset
    Dim strIp As String
    Dim strList As String
    Dim strResult As String
    Dim i As Integer
    Dim strPortURL As String
    Dim ResponseXML As Object
    Dim ResponseNode As Object
    Dim strArtifact����� As String
    Dim strStatus As String
    Dim arrSamlAssertion() As String
    Dim strSoapPost As String
    Dim strErr As String
    Dim strAppStart As String
    On Error GoTo Errhand
    '�������
    Call ClearValues
    'ZLUA��¼
    strAppStart = gobjFile.GetParentFolderName(App.Path)
    If Len(strCmd) > 0 And InStr(strCmd, ",") = 0 And InStr(gstrCommand, "&") > 0 Then
        
        If Not gobjFile.FileExists(strAppStart & "\" & "ZLUA.ini") Then
            MsgBox "δ�ҵ�" & strAppStart & "\" & "ZLUA.ini���޷���ȡ�����ļ�", vbInformation + vbOKOnly, "��ʾ"
            GoTo Errhand
        End If
        Set rsIni = ReadINIToRec(strAppStart & "\" & "ZLUA.ini")
        rsIni.Filter = ""
        rsIni.Filter = "��Ŀ='PortURL'"
        strPortURL = rsIni("����").Value
        rsIni.Filter = ""
        rsIni.Filter = "��Ŀ='UserURL'"
        mstrUserURL = rsIni("����").Value
        rsIni.Filter = "��Ŀ='AppID'"
        mstrAppID = rsIni("����").Value
        
        strArtifact����� = Split(gstrCommand, "&")(0)
        
        If Trim(strPortURL) = "" Then
            MsgBox "�����õ����¼�����ַ", vbInformation + vbOKOnly, "��ʾ"
        ElseIf (Trim(mstrUserURL) = "") Then
            MsgBox "�������˻������ַ", vbInformation + vbOKOnly, "��ʾ"
        Else
            '����httprequest��ʽ-----------------
            mstr���� = SoapEnvelope("GetSAMLResponseByArtifact", strArtifact�����, "", "", "")
            Call PostData(strPortURL, "GetSAMLResponseByArtifact", mstr����, 5)
            strSoapPost = mstrPostData
            strSoapPost = Replace(strSoapPost, "&gt;", ">")
            strSoapPost = Replace(strSoapPost, "&lt;", "<")
            
            '-------------
            '����XML�ı����ݲ��ж��Ƿ񷵻���ȷ��֤���
            If strSoapPost <> "" Then
                Set objDoc = CreateObject("MSXML2.DOMDocument")
                Call objDoc.loadXML(strSoapPost)
                Set ResponseXML = objDoc.documentElement
                Set ResponseNode = ResponseXML.selectSingleNode(".//samlp:StatusCode")
                strStatus = ResponseNode.Attributes(0).Text
                If strStatus <> "" Then
                    Select Case strStatus
                    Case "urn:oasis:names:tc:SAML:2.0:status:Success"
                        '��������ɹ�
                        '��ȡ��¼��Ϣ:�û���/����/������
                        Set ResponseNode = ResponseXML.selectSingleNode(".//saml:AttributeValue")
                        If ResponseNode Is Nothing Then
                            strStatus = ""
                        Else
                            strStatus = ResponseNode.Text
                        End If
                        
                        '��ȡZLUA�˻���
                        Set ResponseNode = ResponseXML.selectSingleNode(".//saml:NameID")
                        mstrZLUAUser = ResponseNode.Text
                        
                        Set ResponseNode = ResponseXML.selectSingleNode(".//saml:Assertion")
                        mstrSamlAssertion = ResponseNode.xml
                        '�����ϢΪ�գ�����ʾ��¼��Ϣ�򣬲����ýӿ��ϴ���Ϣ�Ա��´γɹ���ȡ
                        mblnZLUA = True
                        If Trim(strStatus) = "" Then
                            mblnUAAddUser = True
                            '--�������ZLUA�û��˻�
                        Else
                            If InStr(strStatus, "/") > 0 And InStr(strStatus, "@") > 0 And InStr(strStatus, "/") < InStr(strStatus, "@") Then
                               Me.txtUser.Text = Mid(strStatus, 1, InStr(strStatus, "/") - 1)
                               Me.txtPassWord.Text = Mid(strStatus, InStr(strStatus, "/") + 1, InStr(strStatus, "@") - InStr(strStatus, "/") - 1)
                               Me.cboServer.Text = Mid(strStatus, InStr(strStatus, "@") + 1)
                            End If
                            If Trim(txtUser.Text) <> "" And Trim(txtPassWord.Text) <> "" Then cmdOK_Click
                        End If
                    Case Else
                        '��������ʧ�ܣ����»�ȡ����������Ϣ
                        Set ResponseNode = ResponseXML.selectSingleNode(".//samlp:StatusMessage")
                        strStatus = ResponseNode.Text
                        strErr = "������Ϣ��" & strStatus
                        GoTo Errhand
                    End Select
                End If
            End If
            
        End If
    End If

    '�����¼
    ReDim ArrCommandPortal(0)
    If InStr(strCmd, ",") > 0 Then
        If objSoap Is Nothing Then
            Set objSoap = CreateObject("MSSOAP.SoapClient30")
        End If
        
        If Err.Number <> 0 Then
            Screen.MousePointer = 0
            Err.Clear
            MsgBox "�޷�����SOAP����", vbOKOnly + vbInformation, "��ʾ"
            Set objSoap = Nothing
            GoTo Errhand
        End If
        If Not gobjFile.FileExists(strAppStart & "\" & "Portal.ini") Then
            MsgBox "δ�ҵ� " & strAppStart & "\" & "Portal.ini ·��", vbInformation + vbOKOnly, "��ʾ"
            GoTo Errhand
        End If
        Set rsIni = ReadINIToRec(strAppStart & "\" & "Portal.ini")
        rsIni.Filter = ""
        rsIni.Filter = "��Ŀ='IP'"
        strIp = rsIni("����").Value
        rsIni.Filter = ""
        rsIni.Filter = "��Ŀ='List'"
        strList = rsIni("����").Value
        '��ǰ��ʧ��10.35.10����
        ArrCommandPortal = Split(strCmd, ",")
    End If
    
    ArrCommand = Split(strCmd, " ")
    
    If UBound(ArrCommandPortal) > 0 Then
        Call objSoap.MSSoapInit("http://" & strIp & "/" & strList & "?wsdl")
        strResult = objSoap.getZLSSORet(ArrCommandPortal(0), ArrCommandPortal(1))
        If strResult <> "" And InStr(strResult, "/") > 0 And InStr(strResult, "@") > 0 And InStr(strResult, "/") < InStr(strResult, "@") Then
           Me.txtUser.Text = Mid(strResult, 1, InStr(strResult, "/") - 1)
           Me.txtPassWord.Text = Mid(strResult, InStr(strResult, "/") + 1, InStr(strResult, "@") - InStr(strResult, "/") - 1)
           Me.cboServer.Text = Mid(strResult, InStr(strResult, "@") + 1)
        End If
        mblnת�� = True
        If Trim(txtUser.Text) <> "" And Trim(txtPassWord.Text) <> "" Then cmdOK_Click
    ElseIf InStr(ArrCommand(0), "=") > 0 And InStr(ArrCommand(0), "&") = 0 Then
        '�������������õ���̨��¼�ĸ�ʽ
        For i = LBound(ArrCommand) To UBound(ArrCommand)
            If UCase(ArrCommand(i)) Like "USER=*" Then
                Me.txtUser.Text = Split(ArrCommand(i), "=")(1)
            ElseIf UCase(ArrCommand(i)) Like "PASS=*" Then
                Me.txtPassWord.Text = Split(ArrCommand(i), "=")(1)
            ElseIf UCase(ArrCommand(i)) Like "SERVER=*" Then
                Me.cboServer.Text = Split(ArrCommand(i), "=")(1)
            ElseIf UCase(ArrCommand(i)) Like "ONLYONE=*" Then
                If Split(ArrCommand(i), "=")(1) = "1" Then
                    If App.PrevInstance = True Then
                        MsgBox "�����ظ������������"
                        End
                    End If
                End If
            End If
        Next
        If Trim(txtUser.Text) <> "" And Trim(txtPassWord.Text) <> "" Then Call cmdOK_Click
    End If
    Docmd = mblnOk
    Set objSoap = Nothing
    Exit Function
Errhand:
    If strErr <> "" Then
        MsgBox strErr, vbInformation + vbOKOnly, "��ʾ"
        strErr = ""
    Else
        If Err.Number <> 0 Then
            MsgBox Err.Description, vbInformation + vbOKOnly, "��ʾ"
        End If
    End If
    Set objSoap = Nothing
    Err.Clear
End Function

Private Function GetXMLVersion() As String
    
    Dim varXMLVersion As Variant
    Dim strXMLVer As String
    Dim intLoop As Integer
    Dim objXML As Object
    
    On Error GoTo Errhand
        
    varXMLVersion = Split("6.0,4.0", ",")
    
    On Error Resume Next
    For intLoop = 0 To UBound(varXMLVersion)
        Err = 0
        Set objXML = CreateObject("MSXML2.DOMDocument." & varXMLVersion(intLoop))
        If Err = 0 Then
            strXMLVer = varXMLVersion(intLoop)
            Exit For
        End If
    Next
    On Error GoTo Errhand
    
    If strXMLVer = "" Then
        MsgBox "����MSXML2.DOMDocument����ʧ��"
        Exit Function
    End If
    
    GetXMLVersion = strXMLVer
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
Errhand:
    MsgBox Err.Description
End Function

Private Sub txtUser_LostFocus()
    Call UpdateUser
End Sub

Private Sub txtUser_Validate(Cancel As Boolean)
    Call UpdateUser
End Sub

Private Sub UpdateUser()
On Error GoTo errH
    If IsNumeric(txtUser.Text) Then
        txtUser.Text = "U" & txtUser.Text
    End If
    Exit Sub
errH:
    MsgBox Err.Description, vbCritical, gstrSysName
    Err.Clear
End Sub

Private Function CheckInput(ByRef strUsername As String, ByRef strPassword As String, ByRef strServerName As String) As Boolean
'����:����û������룬������������ֵ
    '�����ַ���
    Dim intPos As Integer, strNote As String
    
    On Error GoTo InputError
    '------�����û��Ƿ�oracle�Ϸ��û�----------------
    strUsername = Trim(txtUser.Text)
    strPassword = Trim(txtPassWord.Text)
    strServerName = Trim(cboServer.Text)
    
    '��Ч�ַ���Ч��
    If Len(Trim(txtUser.Text)) = 0 Then
        strNote = "�������û���"
        txtUser.SetFocus
        GoTo InputError
    End If
    
    If Len(strUsername) <> 1 Then
        If Mid(strUsername, 1, 1) = "/" Or Mid(strUsername, 1, 1) = "@" Or Mid(strUsername, Len(strUsername) - 1, 1) = "/" Or Mid(strUsername, Len(strUsername) - 1, 1) = "@" Then
            txtUser.SetFocus
            strNote = "�û�������"
            SetConState
            Exit Function
        End If
    End If
    
    If Trim(strPassword) <> "" And Len(strPassword) <> 1 Then
        If Mid(strPassword, Len(strPassword) - 1, 1) = "/" Or Mid(strPassword, Len(strPassword) - 1, 1) = "@" Or Mid(strPassword, 1, 1) = "/" Or Mid(strPassword, 1, 1) = "@" Then
            If txtPassWord.Enabled Then txtPassWord.SetFocus
            strNote = "�������"
            GoTo InputError
        End If
    End If
    If Trim(strServerName) <> "" Then
        If Mid(strServerName, Len(strServerName) - 1, 1) = "/" Or Mid(strServerName, Len(strServerName) - 1, 1) = "@" Or Mid(strServerName, 1, 1) = "/" Or Mid(strServerName, 1, 1) = "@" Then
            strNote = "�������Ӵ�����"
            cboServer.SetFocus
            GoTo InputError
        End If
    End If
    
    intPos = InStr(strUsername, "@")
    If intPos > 0 Then
        strServerName = Mid(strUsername, intPos + 1)
        strUsername = Mid(strUsername, 1, intPos - 1)
    End If
    
    intPos = InStr(strUsername, "/")
    If intPos > 0 Then
        strPassword = Mid(strUsername, intPos + 1)
        strUsername = Mid(strUsername, 1, intPos - 1)
    End If
    
    intPos = InStr(strPassword, "@")
    If intPos > 0 Then
        strServerName = Mid(strPassword, intPos + 1)
        strPassword = Mid(strPassword, 1, intPos - 1)
    End If
    If Len(Trim(strPassword)) = 0 Then
        strNote = "����������"
        GoTo InputError
    End If
    CheckInput = True
    Exit Function
InputError:
    If strNote <> "" Then
        MsgBox strNote, vbExclamation, gstrSysName
    End If
End Function

Private Function CheckPwdExpiry() As Boolean
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim dtExpiryDate As Date
    Dim dtNow As Date
    Dim intDiff As Integer
    
    strSQL = "Select EXPIRY_DATE From User_Users Where UserName=User"
    Set rsData = OpenSQLRecord(strSQL, "���������Ч")
    
    If rsData.BOF = False Then
        If IsNull(rsData("EXPIRY_DATE").Value) = True Then
            CheckPwdExpiry = False
            Exit Function
        End If
        dtExpiryDate = Format(rsData("EXPIRY_DATE").Value, "YYYY-MM-DD HH:MM:SS")
        '�жϹ��������뵱ǰ�����������
        dtNow = Format(Currentdate, "YYYY-MM-DD HH:MM:SS")
       
        intDiff = DateDiff("d", dtNow, dtExpiryDate)
        
        If intDiff > 7 Then
            CheckPwdExpiry = False
            Exit Function
        End If
        
        If intDiff > 3 And intDiff <= 7 Then
            '��ʾ�޸�����
            If MsgBox("������Ч�ڻ���" & intDiff & "��,�Ƿ������޸�����?", vbQuestion + vbYesNo, "�����������") = vbYes Then
                CheckPwdExpiry = True
            Else
                CheckPwdExpiry = False
                Exit Function
            End If
        ElseIf intDiff <= 3 Then
            CheckPwdExpiry = True
            MsgBox "������Ч�ڻ���" & intDiff & "�죬���������޸����롣", vbInformation
        Else
            CheckPwdExpiry = False
            Exit Function
        End If
    End If
End Function

Private Function ConvertStr(ByVal strSource As String) As String
    '����:ȥ���ַ����Ŀո�\���з�,��ת��Ϊ��д
    
    strSource = UCase(strSource)
    strSource = Replace(strSource, " ", "")
    strSource = Replace(strSource, vbNewLine, "")
    strSource = Replace(strSource, vbCr, "")
    strSource = Replace(strSource, vbLf, "")
    strSource = Replace(strSource, vbTab, "")
    strSource = Replace(strSource, vbBack, "")
    ConvertStr = strSource
End Function

