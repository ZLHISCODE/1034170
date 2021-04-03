VERSION 5.00
Begin VB.Form frmProcConfigure 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4995
   Icon            =   "frmProcConfigure.frx":0000
   LinkTopic       =   "��������"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtPort 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2190
      MaxLength       =   20
      TabIndex        =   4
      Top             =   2175
      Width           =   2625
   End
   Begin VB.TextBox txtServerIP 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2190
      MaxLength       =   20
      TabIndex        =   2
      Top             =   1410
      Width           =   2625
   End
   Begin VB.TextBox txtSID 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2190
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1800
      Width           =   2625
   End
   Begin VB.TextBox txtUserName 
      Height          =   300
      Left            =   2190
      TabIndex        =   0
      Top             =   615
      Width           =   2625
   End
   Begin VB.TextBox txtUserPwd 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2190
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1005
      Width           =   2625
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2145
      TabIndex        =   8
      Top             =   2880
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3330
      TabIndex        =   7
      Top             =   2880
      Width           =   1100
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "��������(&T)"
      Height          =   350
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "�����˴���������"
      Top             =   2865
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   -75
      TabIndex        =   5
      Top             =   2610
      Width           =   5310
   End
   Begin VB.Image imgFlag 
      Height          =   720
      Left            =   225
      Picture         =   "frmProcConfigure.frx":000C
      Top             =   615
      Width           =   720
   End
   Begin VB.Label lblPort 
      AutoSize        =   -1  'True
      Caption         =   "�˿ں�"
      Height          =   180
      Left            =   1545
      TabIndex        =   14
      Top             =   2205
      Width           =   540
   End
   Begin VB.Label LblIP 
      AutoSize        =   -1  'True
      Caption         =   "���ݷ�����IP"
      Height          =   180
      Left            =   1005
      TabIndex        =   13
      Top             =   1470
      Width           =   1080
   End
   Begin VB.Label lblSID 
      AutoSize        =   -1  'True
      Caption         =   "���ݿ�ʵ��"
      Height          =   180
      Left            =   1185
      TabIndex        =   12
      Top             =   1860
      Width           =   900
   End
   Begin VB.Label lblMarks 
      BackStyle       =   0  'Transparent
      Caption         =   "����ҳ��չʾ�����ӵ����ݿ⣬����д��Ӧ���ݿ���û��������롢IP�Լ�ʵ����"
      Height          =   390
      Left            =   240
      TabIndex        =   11
      Top             =   150
      Width           =   4590
   End
   Begin VB.Label lblUserName 
      AutoSize        =   -1  'True
      Caption         =   "�����û�"
      Height          =   180
      Left            =   1365
      TabIndex        =   10
      Top             =   675
      Width           =   720
   End
   Begin VB.Label lblPwd 
      AutoSize        =   -1  'True
      Caption         =   "�û�����"
      Height          =   180
      Left            =   1365
      TabIndex        =   9
      Top             =   1065
      Width           =   720
   End
End
Attribute VB_Name = "frmProcConfigure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrUserName As String
Private mstrUserPwd As String
Private mstrServerIP As String
Private mstrSID As String
Private mstrPort As String
Private mstrConnection As String
Private mobjMain As Object
Private mblnOK As Boolean
Private mblnLocal As Boolean
Private mclsCiph As clsCipher
Private mcnThis As ADODB.Connection

Public Event AfterConn(ByVal cnOracle As ADODB.Connection)

Public Function ShowConfigure(ByVal objMain As Object) As Boolean
    Set mobjMain = objMain
    Me.Show 1, mobjMain
    ShowConfigure = mblnOK
End Function

Private Function OraDataOpen(ByVal strServerIP As String, ByVal strSID As String, ByVal strUserName As String, ByVal strUserPwd As String, ByVal strPort As String) As Boolean
    '------------------------------------------------
    '���ܣ� ��ָ�������ݿ�
    '������
    '   strServerName�������ַ���
    '   strUserName���û���
    '   strUserPwd������
    '���أ� ���ݿ�򿪳ɹ�������true��ʧ�ܣ�����false
    '------------------------------------------------
    Dim strSQL As String
    Dim strError As String
    Dim cnOracle As New ADODB.Connection
    Dim strServer As String
    
    On Error Resume Next
    err = 0
    DoEvents
    strServer = "(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=" & strServerIP & ")(PORT = " & strPort & ")))(CONNECT_DATA=(SERVICE_NAME=" & strSID & ")))"
    With cnOracle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServer, strUserName, strUserPwd
        If err <> 0 Then
            '���������Ϣ
            strError = err.Description
            If InStr(strError, "�Զ�������") > 0 Then
                MsgBox "���Ӵ��޷��������������ݷ��ʲ����Ƿ�������װ��", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "�޷���������������" & vbCrLf & "������Oracle�������Ƿ���ڸñ�������������������ַ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "�޷����ӣ�����������ϵ�Oracle�����������Ƿ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE���ڳ�ʼ�����ڹرգ����Ժ����ԡ�", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBox "ORACLE�����ã������������ݿ�ʵ���Ƿ�������", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBox "�û�" & UCase(strUserName) & "�Ѿ���¼���������ظ���¼(�Ѵﵽϵͳ�����������¼��)��", vbExclamation, gstrSysName
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBox "�����û�������������ָ�������޷���¼��", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-28000") > 0 Then
                MsgBox "�����û��Ѿ������ã��޷���¼��", vbInformation, gstrSysName
            Else
                MsgBox strError, vbInformation, gstrSysName
            End If
            
            OraDataOpen = False
            Exit Function
        End If
    End With
    Set mcnThis = cnOracle
    
    err = 0
    On Error GoTo errHand
    
    OraDataOpen = True
    Exit Function
    
errHand:
    MsgBox err.Description, vbCritical, Me.Caption
    OraDataOpen = False
    err = 0
End Function

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()

    Dim rsSQL As ADODB.Recordset
    Dim clsCiph As New clsCipher
    Dim lngRow As Long

    On Error GoTo errHand
    
    If TestConnect = False Then Exit Sub
    
    Call gclsBase.SQLRecord(rsSQL)

    mstrConnection = mstrUserName & "/" & mstrUserPwd & "/" & mstrSID & "/" & mstrServerIP & "/" & mstrPort
    mstrConnection = clsCiph.Cipher("zLw09OewKKO1`;owEWO-=,./w[]wwqq3##=``44314325", mstrConnection)
    
    '���ַ���&���ŵĴ���
    mstrConnection = Replace(mstrConnection, "&", "' || chr(38) || '")
    gstrSQL = "Zl_Zlprocedureconnect_Update('��������','" & mstrConnection & "')"
    Call gclsBase.SQLRecordAdd(rsSQL, gstrSQL)
    
    If SQLRecordExecute(rsSQL) Then
        mblnOK = True
        RaiseEvent AfterConn(mcnThis)
        Unload Me
    End If
    Set clsCiph = Nothing
    Exit Sub
errHand:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Function TestConnect() As Boolean

    Dim strUserName As String
    Dim strServerIP As String
    Dim strPassword As String
    Dim strSID As String
    Dim strPort As String
    Dim strNote As String
    Dim strPwdTmp As String
    
    On Error GoTo InputError
    
    '------�����û��Ƿ�oracle�Ϸ��û�----------------
    strUserName = Trim(txtUserName.Text)
    strPassword = Trim(txtUserPwd.Text)
    strServerIP = Trim(txtServerIP.Text)
    strSID = Trim(txtSID.Text)
    strPort = Trim(txtPort.Text)
    
    '��Ч�ַ���Ч��
    If Len(Trim(txtUserName.Text)) = 0 Then
        strNote = "�������û���"
        txtUserName.SetFocus
        GoTo InputError
    End If
    
    If Len(Trim(txtUserPwd.Text)) = 0 Then
        strNote = "����������"
        txtUserName.SetFocus
        GoTo InputError
    End If
    
    If Len(Trim(txtSID.Text)) = 0 Then
        strNote = "���������ݿ�ʵ��"
        txtSID.SetFocus
        GoTo InputError
    End If
    
    If Len(Trim(txtServerIP.Text)) = 0 Then
        strNote = "���������ݿ�IP"
        txtServerIP.SetFocus
        GoTo InputError
    End If
    
    If Len(Trim(txtPort.Text)) = 0 Then
        strNote = "������˿ں�"
        txtServerIP.SetFocus
        GoTo InputError
    End If
    
    If Len(strUserName) <> 1 Then
        If Mid(strUserName, 1, 1) = "/" Or Mid(strUserName, 1, 1) = "@" Or Mid(strUserName, Len(strUserName) - 1, 1) = "/" Or Mid(strUserName, Len(strUserName) - 1, 1) = "@" Then
            txtUserName.SetFocus
            strNote = "�û�������"
            GoTo InputError
        End If
    End If
    If Trim(strPassword) <> "" And Len(strPassword) <> 1 Then
        If Mid(strPassword, Len(strPassword) - 1, 1) = "/" Or Mid(strPassword, Len(strPassword) - 1, 1) = "@" Or Mid(strPassword, 1, 1) = "/" Or Mid(strPassword, 1, 1) = "@" Then
            txtUserPwd.SetFocus
            strNote = "�������"
            GoTo InputError
        End If
    End If
    
    strPwdTmp = strPassword
    strUserName = UCase(strUserName)
    If UCase(strUserName) <> "SYSTEM" And UCase(strUserName) <> "SYS" Then
        strPassword = TranPasswd(strPassword)
    End If

    If OraDataOpen(strServerIP, strSID, strUserName, strPassword, strPort) Then
        mstrUserName = strUserName
        mstrUserPwd = strPwdTmp
        mstrServerIP = strServerIP
        mstrSID = strSID
        mstrPort = strPort
                
        TestConnect = True
    End If
    
    Exit Function
    
InputError:
    If strNote <> "" Then
        MsgBox strNote, vbExclamation, gstrSysName
    End If
    Exit Function
    
End Function

Private Sub cmdTest_Click()
    If TestConnect = True Then
        MsgBox "�������ӳɹ��������������ʸ����ݿ⡣", vbInformation + vbOKOnly, gstrSysName
    End If
    
'    If OraDataOpen(strServerIP, strSID, strUserName, strPassword, strPort) Then
'        mstrUserName = strUserName
'        mstrUserPwd = strPassword
'        mstrServerIP = strServerIP
'        mstrSID = strSID
'        mstrPort = strPort
'        cmdOK.Enabled = True
'        MsgBox "�������ӳɹ��������������ʸ����ݿ⡣", vbInformation + vbOKOnly, gstrSysName
'        Exit Sub
'    Else
'        cmdOK.Enabled = False
'    End If
End Sub

Private Sub Form_Activate()

    On Error GoTo errHand
    Dim strConnection As String
    Dim strCon() As String
    
    Set mclsCiph = New clsCipher
    strConnection = gclsBase.GetOraConn("��������")
    If strConnection <> "" Then
        strConnection = mclsCiph.Decipher("zLw09OewKKO1`;owEWO-=,./w[]wwqq3##=``44314325", strConnection)
        strCon = Split(strConnection, "/")
        txtUserName.Text = strCon(0)
        txtUserPwd.Text = strCon(1)
        txtSID.Text = strCon(2)
        txtServerIP.Text = strCon(3)
        txtPort.Text = strCon(4)
    End If
    Set mclsCiph = Nothing
    txtUserPwd.SetFocus
    Exit Sub
errHand:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Sub Form_Load()
    txtUserName.Text = GetSetting(appName:="ZLSOFT", Section:="ע����Ϣ\��½��Ϣ", Key:="USER", Default:="")
    txtPort.Text = "1521"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (mclsCiph Is Nothing) Then
        Set mclsCiph = Nothing
    End If
End Sub

Private Sub txtPort_GotFocus()
    gclsBase.TxtSelAll txtPort
End Sub

Private Sub txtServerIP_GotFocus()
    gclsBase.TxtSelAll txtServerIP
End Sub

Private Sub txtSID_GotFocus()
    gclsBase.TxtSelAll txtSID
End Sub

Private Sub txtUserName_GotFocus()
    gclsBase.TxtSelAll txtUserName
End Sub

Private Sub txtUserPwd_GotFocus()
    gclsBase.TxtSelAll txtUserPwd
End Sub

