VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRole 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "��ɫ��Ȩ����"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8760
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   Picture         =   "frmRole.frx":0000
   ScaleHeight     =   6240
   ScaleWidth      =   8760
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picHLine 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   1500
      MousePointer    =   7  'Size N S
      ScaleHeight     =   90
      ScaleWidth      =   5835
      TabIndex        =   17
      Top             =   3900
      Width           =   5835
   End
   Begin VB.CheckBox chkOnlyShowNOGroups 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ֻ��ʾδ�����ɫ(&B)"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   5250
      TabIndex        =   4
      Top             =   570
      Value           =   1  'Checked
      Width           =   2040
   End
   Begin VB.CommandButton cmdNewGroup 
      Caption         =   "�½���(&N)"
      Height          =   350
      Left            =   7575
      TabIndex        =   6
      Top             =   945
      Width           =   1155
   End
   Begin VB.CommandButton cmdGrantAll 
      Caption         =   "�ָ�Ȩ��"
      Height          =   350
      Left            =   7560
      TabIndex        =   9
      Top             =   1995
      Width           =   1155
   End
   Begin VB.CommandButton cmdGrant 
      Caption         =   "��Ȩ(&G)"
      Height          =   350
      Left            =   7575
      TabIndex        =   8
      Top             =   1650
      Width           =   1155
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "����(&C)��"
      Height          =   350
      Left            =   7560
      TabIndex        =   10
      Top             =   2355
      Width           =   1155
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "���ӽ�ɫ(&A)"
      Height          =   350
      Left            =   7575
      TabIndex        =   7
      Top             =   1305
      Width           =   1155
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   3345
      TabIndex        =   3
      Top             =   540
      Width           =   1680
   End
   Begin MSComctlLib.TreeView tvwGroups 
      Height          =   2955
      Left            =   135
      TabIndex        =   1
      Top             =   870
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   5212
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   295
      Style           =   7
      ImageList       =   "img16"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   -465
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRole.frx":803A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdUser 
      Cancel          =   -1  'True
      Caption         =   "�޸Ľ�ɫ����Ȩ�û�"
      Height          =   350
      Left            =   6585
      TabIndex        =   15
      Top             =   3975
      Width           =   1875
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "�޸�ģ���ʹ��Ȩ��"
      Height          =   350
      Left            =   4560
      TabIndex        =   14
      Top             =   3975
      Width           =   1875
   End
   Begin VB.ComboBox cmbSystem 
      Height          =   300
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   4020
      Width           =   3150
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   -285
      Top             =   1755
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRole.frx":85D4
            Key             =   "Role"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwRole 
      Height          =   2970
      Left            =   2925
      TabIndex        =   5
      Top             =   885
      Width           =   4410
      _ExtentX        =   7779
      _ExtentY        =   5239
      Arrange         =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      Icons           =   "ils32"
      SmallIcons      =   "ils32"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "_��ɫ"
         Object.Tag             =   "��ɫ"
         Text            =   "��ɫ"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Grantee"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Admin_Option"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvwModule 
      Height          =   1605
      Left            =   135
      TabIndex        =   16
      Top             =   4575
      Width           =   8580
      _ExtentX        =   15134
      _ExtentY        =   2831
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ɫ��Ȩ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   105
      TabIndex        =   18
      Top             =   150
      Width           =   1440
   End
   Begin VB.Line lnRole 
      BorderColor     =   &H80000001&
      BorderWidth     =   2
      X1              =   1605
      X2              =   9165
      Y1              =   -195
      Y2              =   -195
   End
   Begin VB.Label lblRoleGroup 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "��ɫ����Ϣ"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   135
      TabIndex        =   0
      Top             =   630
      Width           =   1530
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2940
      TabIndex        =   2
      Top             =   600
      Width           =   360
   End
   Begin VB.Line lnModuel 
      BorderColor     =   &H80000001&
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   1050
      X2              =   8760
      Y1              =   -150
      Y2              =   -150
   End
   Begin VB.Label lblModule 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����Ȩģ��"
      Height          =   180
      Left            =   135
      TabIndex        =   11
      Top             =   4320
      Width           =   900
   End
   Begin VB.Label lblSystem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ȩϵͳ"
      Height          =   180
      Left            =   135
      TabIndex        =   12
      Top             =   4065
      Width           =   720
   End
   Begin VB.Menu mnuPopu 
      Caption         =   "�����˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuPopuAddGroups 
         Caption         =   "�½���(&N)"
      End
      Begin VB.Menu mnuPopuModify 
         Caption         =   "�޸���(&M)"
      End
      Begin VB.Menu mnuPopuDeleteGroups 
         Caption         =   "ɾ����(&D)"
      End
   End
   Begin VB.Menu mnuPopuRole 
      Caption         =   "�����˵���ɫ"
      Visible         =   0   'False
      Begin VB.Menu mnuPopuRoleAdd 
         Caption         =   "���ӽ�ɫ(&N)"
      End
      Begin VB.Menu mnuPopuRoleDelete 
         Caption         =   "ɾ����ɫ(&M)"
      End
      Begin VB.Menu mnuPopuRoleMove 
         Caption         =   "��ɫ�Ƶ�(&M)��"
         Begin VB.Menu mnuPopuRoleMoveGroups 
            Caption         =   "��1"
            Index           =   0
         End
      End
   End
End
Attribute VB_Name = "frmRole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mrsRole As New ADODB.Recordset
Private mblnFirst As Boolean
Private mblnMoveTop As Boolean
Private msngPreHeigt As Single

Private mfrmGrant As frmRoleGrant

Private Sub chkOnlyShowNOGroups_Click()
    'ֻ��ʾδ������Ľ�ɫ
    If tvwGroups.SelectedItem Is Nothing Then Exit Sub
    tvwGroups_NodeClick tvwGroups.SelectedItem
End Sub

Private Sub chkOnlyShowNOGroups_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub cmdAdd_Click()
    Dim cnTemp As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim strRoleName As String
    Dim blnLimited As Boolean
    Dim lst As ListItem
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    '���û�ӵ�еĽ�ɫ�����ﵽ148��ʱ���û���¼ʱ����ʾ����
    gstrSQL = "Select Count(*) as ���� From DBA_Role_Privs Where Grantee='" & gstrUserName & "'"
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open gstrSQL, gcnOracle, adOpenKeyset, adLockReadOnly
    If Nvl(rsTemp!����, 0) >= 148 Then
        If Not gblnOwner Then
            MsgBox "��ɫ�����Ѵﵽ������ƣ����������ӡ�", vbInformation, gstrSysName
            Exit Sub
        Else
            '�����߽�ɫ�����ﵽ����ʱ������Systeme�û�����
            If Not CheckRushHours("��ǰʱ�δ���ҵ��߷��ڣ��½���ɫ���ܻ��ϵͳʹ�����һ��Ӱ�죬�Ƿ����") Then
                Exit Sub
            End If
            'SYSTEM�������Ľ�ɫ��������������
            gstrSQL = "Select Count(*) as ���� From DBA_Role_Privs Where Grantee='SYSTEM'"
            Set rsTemp = New ADODB.Recordset
            rsTemp.CursorLocation = adUseClient
            rsTemp.Open gstrSQL, gcnOracle, adOpenKeyset, adLockReadOnly
            If Nvl(rsTemp!����, 0) >= 148 Then
                MsgBox "��ɫ�����Ѵﵽ������ƣ����������ӡ�", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        blnLimited = True
    End If
    
    strRoleName = frmNameEdit.GetName(name��ɫ)
    If strRoleName = "" Then Exit Sub
    strRoleName = "ZL_" & UCase(Trim(strRoleName))
    
    Set cnTemp = gcnOracle
    If blnLimited And gblnOwner Then
        Set gcnSystem = GetConnection("SYSTEM")
        If gcnSystem Is Nothing Then Exit Sub
        Set cnTemp = gcnSystem
    End If
    
    On Error Resume Next
    cnTemp.Execute "Create Role " & strRoleName & " Not Identified"
    
    If err <> 0 Then
        MsgBox "��������������������߽�ɫ�����������ݿ�Ĳ�������" & vbCrLf & _
                "(���޸����ݿ���������������ɫ��Ŀ)�����½�ɫ����ʧ�ܡ�", vbExclamation, gstrSysName
    Else
        On Error GoTo errHandle
        Call GrantSpecialToRole(cnTemp, strRoleName, False, "", True)
        If tvwGroups.SelectedItem Is Nothing Then
        ElseIf tvwGroups.SelectedItem.Key <> "Root" Then
            '���˺�:20070615����
            '���̲���:zlTools.b_Rolegroupmgr.RoletoRolegroup
            '        ����_In In ZlRolegroups.����%Type,
            '        ��ɫ_In In ZlRolegroups.��ɫ%Type := Null
            gstrSQL = "zlTools.b_Rolegroupmgr.RoleToRoleGroup("
            gstrSQL = gstrSQL & "'" & Mid(tvwGroups.SelectedItem.Key, 2) & "',"
            gstrSQL = gstrSQL & "'" & strRoleName & "')"
            ExecuteProcedure gstrSQL, Me.Caption
        End If
        Set lst = lvwRole.ListItems.Add(, strRoleName, Mid(strRoleName, 4), "Role", "Role")
        lst.Selected = True
        Call FillModule
    End If
    Call SetEnable
    
    Exit Sub
errHandle:
    MsgBox "����" & err.Number & vbCrLf & vbCrLf & err.Description, vbInformation, gstrSysName
End Sub

Private Sub DeleteRole()
    '------------------------------------------------------------------------------------------------------------------------------------------------
    '����:ɾ����ɫ
    '����:���˺�
    '����:2007/06/15
    '------------------------------------------------------------------------------------------------------------------------------------------------
    Dim strRoleName As String
    Dim intIndex As Integer
    
    On Error GoTo errHandle
    
    If lvwRole.SelectedItem Is Nothing Then Exit Sub
    
    strRoleName = lvwRole.SelectedItem.Key
    intIndex = lvwRole.SelectedItem.Index
    If MsgBox("���Ҫɾ����ɫ��" & strRoleName & "����", vbDefaultButton2 Or vbQuestion Or vbYesNo, gstrSysName) = vbNo Then Exit Sub
    
    Screen.MousePointer = 11
    If lvwRole.SelectedItem.SubItems(1) = UCase(gstrUserName) _
        And lvwRole.SelectedItem.SubItems(2) = "YES" Then
        gcnOracle.Execute "Drop Role " & strRoleName
    Else
        Set gcnSystem = GetConnection("SYSTEM")
        If gcnSystem Is Nothing Then Exit Sub
        gcnSystem.Execute "Drop Role " & strRoleName
    End If
    
    gstrSQL = "zlTools.b_Rolegroupmgr.Role_Delete('" & UCase(strRoleName) & "')"
    ExecuteProcedure gstrSQL, Me.Caption
    
    lvwRole.ListItems.Remove intIndex
    If lvwRole.ListItems.Count > 0 Then
        If intIndex > lvwRole.ListItems.Count Then
            intIndex = lvwRole.ListItems.Count
        End If
        lvwRole.ListItems(intIndex).Selected = True
    End If
    Call FillModule
    Call SetEnable
    Screen.MousePointer = 0
    Exit Sub
errHandle:
    Screen.MousePointer = 0
    MsgBox "����" & err.Number & vbCrLf & vbCrLf & err.Description, vbInformation, gstrSysName
End Sub

Private Sub cmdCopy_Click()
    '------------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ƽ�ɫ
    '����:���˺�
    '����:2007/06/15
    '------------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSourceRole As String
    Dim cnTemp As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim strRoleName As String
    Dim blnLimited As Boolean
    
    On Error GoTo errHandle
    
    If lvwRole.SelectedItem Is Nothing Then Exit Sub
    
    '���û�ӵ�еĽ�ɫ�����ﵽ148��ʱ���û���¼ʱ����ʾ����
    gstrSQL = "Select Count(*) as ���� From DBA_Role_Privs Where Grantee='" & gstrUserName & "'"
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open gstrSQL, gcnOracle, adOpenKeyset, adLockReadOnly
    If Nvl(rsTemp!����, 0) >= 148 Then
        If Not gblnOwner Then
            MsgBox "��ɫ�����Ѵﵽ������ƣ����������ӡ�", vbInformation, gstrSysName
            Exit Sub
        Else
            If Not CheckRushHours("��ǰʱ�δ���ҵ��߷��ڣ���ɫ���ƿ��ܻ��ϵͳʹ�����һ��Ӱ�죬�Ƿ����") Then
                Exit Sub
            End If
            '�����߽�ɫ�����ﵽ����ʱ������Systeme�û�����
            'SYSTEM�������Ľ�ɫ��������������
            gstrSQL = "Select Count(*) as ���� From DBA_Role_Privs Where Grantee='SYSTEM'"
            Set rsTemp = New ADODB.Recordset
            rsTemp.CursorLocation = adUseClient
            rsTemp.Open gstrSQL, gcnOracle, adOpenKeyset, adLockReadOnly
            If Nvl(rsTemp!����, 0) >= 148 Then
                MsgBox "��ɫ�����Ѵﵽ������ƣ����������ӡ�", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        blnLimited = True
    End If
    
    strSourceRole = lvwRole.SelectedItem.Key
    strRoleName = frmNameEdit.GetName(name��ɫ)
    If strRoleName = "" Then Exit Sub
    strRoleName = "ZL_" & UCase(Trim(strRoleName))
 
    Set cnTemp = gcnOracle
    If blnLimited And gblnOwner Then
        Set gcnSystem = GetConnection("SYSTEM")
        If gcnSystem Is Nothing Then Exit Sub
        Set cnTemp = gcnSystem
    End If
 
    If Not CopyRole(cnTemp, strSourceRole, strRoleName) Then Exit Sub
    
    '������Ȩ
    Call RoleGrant(strRoleName)
    Call tvwGroups_NodeClick(tvwGroups.SelectedItem)
    
    Dim strKey As String
    
    strKey = lvwRole.SelectedItem.Key
    err = 0: On Error Resume Next
    lvwRole.ListItems(strRoleName).Selected = True
    If err = 0 Then
        lvwRole.ListItems(strKey).Selected = False
    End If
    Call SetEnable
    
    Exit Sub
errHandle:
    MsgBox "����" & err.Number & vbCrLf & vbCrLf & err.Description, vbInformation, gstrSysName
End Sub

Private Function CopyRole(cnTemp As ADODB.Connection, ByVal strSourceRole As String, ByVal strTargetRole As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ɫȨ�޳��µĽ�ɫȨ��
    '����:strSourceRole-Դ��ɫ
    '     strTargetRole-Ŀ���ɫ
    '����:���Ƴɹ�,����true,����False
    '����:���˺�
    '����:2007/07/18
    '------------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    err = 0: On Error Resume Next
    cnTemp.Execute "Create Role " & strTargetRole & " Not Identified"
    If err <> 0 Then
        MsgBox "���������������������Ŀ���ɫ�����������ݿ�Ĳ�������" & vbCrLf & _
                "(���޸����ݿ���������������ɫ��Ŀ)�����½�ɫ����ʧ�ܡ�", vbExclamation, gstrSysName
        Exit Function
    End If
    err = 0: On Error GoTo ErrHand:
    Call GrantSpecialToRole(cnTemp, strTargetRole, False, "", True)
    '����:zlTools.b_Rolegroupmgr.Role_Copy
    '    Դ��ɫ_In   In zlRoleGrant.��ɫ%Type,
    '    Ŀ���ɫ_In In zlRoleGrant.��ɫ%Type
    gstrSQL = "zlTools.b_Rolegroupmgr.Role_Copy("
    gstrSQL = gstrSQL & "'" & UCase(strSourceRole) & "',"
    gstrSQL = gstrSQL & "'" & UCase(strTargetRole) & "')"
    ExecuteProcedure gstrSQL, Me.Caption
    CopyRole = True
    Exit Function
ErrHand:
    Call ShowErrHand
End Function

Private Function RoleGrant(ByVal str��ɫ As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ָ���Ľ�ɫ������Ȩ
    '����:str��ɫ-��ɫ
    '����:���Ƴɹ�,����true,����False
    '����:���˺�
    '����:2007/07/18
    '------------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    Me.MousePointer = vbHourglass
    On Error GoTo errHandle
    

    '����Ȩ�ޱ�����д��Ȩ��
    Dim objclsPrivilege As New clsPrivilege
    Call objclsPrivilege.InitOracle(gcnOracle)
    Call objclsPrivilege.ReGrantPrivilege(str��ɫ)
    Set objclsPrivilege = Nothing
    
    MousePointer = 0
    RoleGrant = True
    Exit Function
errHandle:
    MousePointer = 0
    MsgBox "��ǰ�û���Ȩ�޲�����ɱ�������", vbInformation, gstrSysName
End Function

Private Sub cmdGrant_Click()
    If Not CheckRushHours("��ǰʱ�δ���ҵ��߷��ڣ���ɫ��Ȩ���ܻ��ϵͳʹ�����һ��Ӱ�죬�Ƿ����") Then
        Exit Sub
    End If
    If mfrmGrant Is Nothing Then
        Set mfrmGrant = New frmRoleGrant
    End If
    If mfrmGrant.GrantToRole(lvwRole.SelectedItem.Key) = True Then
        Call FillModule
    End If
End Sub

Private Sub cmdGrantAll_Click()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim lst As ListItem
    
    If Not CheckRushHours("��ǰʱ�δ���ҵ��߷��ڣ��ָ�Ȩ�޿��ܻ��ϵͳʹ�����һ��Ӱ�죬�Ƿ����") Then
        Exit Sub
    End If
    If MsgBox("������������н�ɫ������Ȩ������Ҫ��һ��ʱ��(Լ1.5��������)��" & vbCrLf & "�Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    On Error GoTo errh
    On Error Resume Next
    '���ȴ��������ڵĽ�ɫ
    strSQL = "Select Distinct r.��ɫ" & vbNewLine & _
            "From Zlsystems s, Zlrolegrant r" & vbNewLine & _
            "Where s.��� = r.ϵͳ And s.������ = User And r.��ɫ Not In (Select Granted_Role From User_Role_Privs)"

    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, App.Title)
    If err.Number <> 0 Then err.Clear
    If Not rsTmp Is Nothing Then
        Do While Not rsTmp.EOF
            '���������ڵĽ�ɫ
            gcnOracle.Execute "Create Role " & rsTmp!��ɫ & " Not Identified"
            If err.Number = 0 Then
                '�����ɹ����������ӵ��б���
                Set lst = lvwRole.ListItems.Add(, rsTmp!��ɫ & "", Mid(rsTmp!��ɫ & "", 4), "Role", "Role")
                Call GrantSpecialToRole(gcnOracle, rsTmp!��ɫ, False, "", True)
            Else
                err.Clear
            End If
            rsTmp.MoveNext
        Loop
    End If
    If err.Number <> 0 Then err.Clear
    On Error GoTo errh
    '��ʼ��Ȩ
    Call ReGrantToRole(gcnOracle, "", True)
    '��ʾ��Ȩ�嵥
    If Not lst Is Nothing Then
        lst.Selected = True
        Call FillModule
    End If
    MsgBox "���н�ɫ������Ȩ��ɣ�", vbInformation, gstrSysName
    MousePointer = 0
    Exit Sub
errh:
    If 0 = 1 Then
        Resume
    End If
    MousePointer = 0
    MsgBox "��ǰ�û���Ȩ�޲�����ɱ�������", vbInformation, gstrSysName
End Sub

Private Sub cmdModify_Click()
    frmProgPriv.ProgPriv cmbSystem.ItemData(cmbSystem.ListIndex)
End Sub
Private Sub cmdNewGroup_Click()
    Dim strGroupsName As String
    Dim lst As ListItem
    Dim objNode As Node
ReDo:
    strGroupsName = frmNameEdit.GetName(name����)
    
    If strGroupsName = "" Then Exit Sub
    If ActualLen(strGroupsName) > 30 Then
        MsgBox "������Ľ�ɫ���Ʋ��ܴ���30���ַ���15������,����!", vbDefaultButton1 + vbInformation, gstrSysName
        GoTo ReDo:
    End If
    strGroupsName = UCase(Trim(strGroupsName))
    
    '���˺�:20070615����
    '���̲���:zlTools.b_Rolegroupmgr.Rolegroup_Add(����_In In ZlRolegroups.����%Type)
    err = 0: On Error GoTo ErrHand:
    gstrSQL = "zlTools.b_Rolegroupmgr.Rolegroup_Add("
    gstrSQL = gstrSQL & "'" & UCase(strGroupsName) & "')"
    ExecuteProcedure gstrSQL, Me.Caption
        
    Set objNode = tvwGroups.Nodes.Add("Root", 4, "K" & strGroupsName, strGroupsName, 1, 1)
    'objNode.Selected = True
    objNode.Tag = strGroupsName
    Call LoadMenu(strGroupsName, strGroupsName)
    
    Call FillModule
    Call SetEnable

'     '������
'     '������:��Tvw������һ���µĽ�ɫ����
'     Dim objNode As Node
'     Dim int��� As Integer
'     Dim str���� As String
'ReDo:
'    Err = 0: On Error Resume Next
'    int��� = int��� + 1
'    str���� = "�½���:" & int���
'     Set objNode = tvwGroups.Nodes.Add(, "Root", str����, str����)
'     If Err <> 0 Then
'        Err.Clear: On Error GoTo 0
'        GoTo ReDo
'     End If
'     Err = 0
'    objNode.Tag = "1"
'    objNode.Selected = True
'    tvwGroups.SetFocus
'    tvwGroups.LabelEdit
    Exit Sub
ErrHand:
        Call ShowErrHand
End Sub
Private Sub ShowErrHand()
    '------------------------------------------------------------------------------------------
    '����:��ȡ������Ϣ
    '------------------------------------------------------------------------------------------
    Dim strNote As String, lngErrNum As Long
    If gcnOracle.Errors.Count <> 0 Then
        strNote = gcnOracle.Errors(0).Description
        If InStr(UCase(strNote), "[ZLSOFT]") > 0 Then
            '��־����
            lngErrNum = gcnOracle.Errors(0).NativeError
            MsgBox Split(strNote, "[ZLSOFT]")(1), vbExclamation, App.Title
            Exit Sub
        Else
            MsgBox "ע��:" & vbCrLf & "�����������´���:" & vbCrLf & err.Description, vbExclamation, App.Title
        End If
    Else
        MsgBox "ע��:" & vbCrLf & "�����������´���:" & vbCrLf & err.Description, vbExclamation, App.Title
    End If
End Sub
Private Sub cmdUser_Click()
    If Not CheckRushHours("��ǰʱ�δ���ҵ��߷��ڣ��޸Ľ�ɫ����Ȩ�û����ܻ��ϵͳʹ�����һ��Ӱ�죬�Ƿ����") Then
        Exit Sub
    End If
    If lvwRole.SelectedItem Is Nothing Then Exit Sub
    Call frmRoleUser.ShowEdit(lvwRole.SelectedItem.Text)
End Sub


Private Sub Form_Activate()
    Dim lngTop As Long
    mblnMoveTop = False
    If mblnFirst = False Then Exit Sub
    '���Ի�����
    lngTop = Val(GetSetting("ZLSOFT", "����ģ��\������������\��ɫ����", "PicHLine_TOP", "4170"))
    picHLine.Top = lngTop
    mblnMoveTop = True
    Call Form_Resize
    mblnMoveTop = False
    mblnFirst = False
End Sub

Private Sub Form_Load()
   Dim rsTemp As New ADODB.Recordset
   Dim lngTop As Long
   
    '�жϸ��û��ܷ񴴽���ɫ
    gstrSQL = _
        " Select 1 From User_Sys_Privs Where Privilege='CREATE ROLE'" & _
        " Union" & _
        " Select 1 From Role_Sys_Privs Where Privilege='CREATE ROLE'"
    Call OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    cmdAdd.Enabled = rsTemp.RecordCount > 0
    
    Call Getע����
    Call FillRollGroup
    Call FillSystem
    mblnFirst = True
End Sub

Private Sub cmbSystem_Click()
    'ϵͳ��Ϊ��ʱ�����޸�
    cmdModify.Enabled = cmbSystem.ItemData(cmbSystem.ListIndex) <> 0
    
    Call FillModule
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '���Ի�����
    SaveSetting "ZLSOFT", "����ģ��\������������\��ɫ����", "PicHLine_TOP", picHLine.Top
    If Not mfrmGrant Is Nothing Then
        Set mfrmGrant = Nothing
    End If
End Sub

Private Sub lvwRole_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call FillModule
    Call SetEnable
End Sub

Private Sub Form_Resize()
    Dim sngTemp As Single
    On Error Resume Next
    
    Me.lnRole.X2 = Me.ScaleWidth
    Me.lnModuel.X2 = Me.ScaleWidth
    
    If picHLine.Top > Me.ScaleHeight - cmdUser.Height * 2 - 100 Then
        If Me.ScaleHeight - picHLine.Height - cmdUser.Height * 2 - 100 < cmdUser.Height * 3 + 100 Then
            picHLine.Top = cmdUser.Height * 3 + 100
        Else
            picHLine.Top = Me.ScaleHeight - picHLine.Height - cmdUser.Height * 2 - 100
        End If
    ElseIf picHLine.Top < Me.ScaleTop Then
         picHLine.Top = cmdUser.Height * 3 + 100
    Else
        
        If mblnMoveTop = False And msngPreHeigt <> 0 Then
            If Me.ScaleHeight - msngPreHeigt < cmdUser.Height * 3 + 100 Or Me.ScaleHeight - msngPreHeigt > Me.ScaleHeight Or msngPreHeigt < 0 Then
                Me.picHLine.Top = Me.ScaleHeight - cmdUser.Height * 3 + 100
            Else
                Me.picHLine.Top = Me.ScaleHeight - msngPreHeigt '(cmbSystem.Height + lvwModule.Height + lblModule.Height + 100)
            End If
        End If
    End If
    
    Me.lnModuel.Y1 = Me.picHLine.Top + (lblModule.Height \ 2) + picHLine.Height
    Me.lnModuel.Y2 = Me.lnModuel.Y1
    
    'Me.lblModule.Top = Me.lnModuel.Y1 - (lblModule.Height \ 2)
    Me.lvwRole.Height = Me.picHLine.Top - Me.lvwRole.Top
    Me.tvwGroups.Height = Me.picHLine.Top - tvwGroups.Top
    
    Me.cmbSystem.Top = Me.picHLine.Top + picHLine.Height   'Me.lblSystem.Top - (cmbSystem.Height - lblSystem.Height) \ 2
     
    Me.lblSystem.Top = cmbSystem.Top + (cmbSystem.Height - Me.lblSystem.Height) \ 2 ' Me.picHLine.Top + picHLine.Height ' Me.lblModule.Top + Me.lblModule.Height + 50 + (cmbSystem.Height - lblSystem.Height) \ 2  '  Me.picHLine.Top + 3150 - 2880
    'Me.cmbSystem.Top = Me.lblSystem.Top - (cmbSystem.Height - lblSystem.Height) \ 2       ' Me.picHLine.Top + 3090 - 2880
    lblModule.Top = cmbSystem.Top + cmbSystem.Height + 50
    
    With cmdUser
        .Top = Me.cmbSystem.Top - (.Height - cmbSystem.Height) \ 2
        .Left = Me.ScaleWidth - .Width - 50
    End With
    
    With cmdModify
        .Top = cmdUser.Top
        .Left = cmdUser.Left - .Width - 25
    End With
    With cmbSystem
        If cmdModify.Left - .Left - 100 < 0 Then
            .Width = 0
        Else
            .Width = cmdModify.Left - .Left - 100
        End If
    End With
    
    With Me.lvwModule
        .Top = lblModule.Top + lblModule.Height ' cmdModify.Top + cmdModify.Height + 50
        If Me.ScaleHeight - .Top < 0 Then
             .Height = 0
        Else
            .Height = Me.ScaleHeight - .Top - 50
        End If
        .Width = ScaleWidth - 50 - .Left
    End With
    
    With cmdNewGroup
        .Left = ScaleWidth - .Width - 50
    End With
    cmdAdd.Left = cmdNewGroup.Left
    cmdGrant.Left = cmdNewGroup.Left
    cmdGrantAll.Left = cmdNewGroup.Left
    cmdCopy.Left = cmdNewGroup.Left
    
    With lvwRole
        If cmdNewGroup.Left - 50 - .Left < 0 Then
            .Width = 0
        Else
            .Width = cmdNewGroup.Left - 50 - .Left
        End If
    End With
    With chkOnlyShowNOGroups
        .Left = cmdNewGroup.Left - .Width
        If .Left < lvwRole.Left Then
            .Left = lvwRole.Left
        End If
    End With
    
    With txtSearch
        If chkOnlyShowNOGroups.Left - 100 - .Left < 0 Then
            .Width = 0
        Else
            .Width = chkOnlyShowNOGroups.Left - 100 - .Left
        End If
    End With
    Me.picHLine.Left = 0: Me.picHLine.Width = Me.ScaleWidth
    msngPreHeigt = Me.ScaleHeight - picHLine.Top
End Sub

'Private Sub FillRole()
'    Dim rsTemp As New ADODB.Recordset
'
'    rsTemp.CursorLocation = adUseClient
'
'    '�жϸ��û��ܷ񴴽���ɫ
'    gstrSQL = "Select 1 from User_Sys_privs where privilege='CREATE ROLE' " & _
'        "union Select 1 from Role_Sys_privs where privilege='CREATE ROLE'"
'
'    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
'    cmdAdd.Enabled = rsTemp.RecordCount > 0
'    cmdDelete.Enabled = cmdAdd.Enabled
'    rsTemp.Close
'
'
'
'    '��ʾ���Խ�����Ȩ�Ľ�ɫ
'    If gblnDBA = True Then
'        gstrSQL = "select * from DBA_Roles where Upper(Role) Like 'ZL_%'"
'    Else
'        gstrSQL = "select Granted_Role as Role from user_Role_privs " & _
'            "where Granted_Role Like 'ZL_%'" 'ADMIN_OPTION='YES'ѡ����Բ���
'    End If
'    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
'    lvwRole.ListItems.Clear
'    Do Until rsTemp.EOF
'        lvwRole.ListItems.Add , rsTemp("Role"), Mid(rsTemp("Role"), 4), "Role", "Role"
'        rsTemp.MoveNext
'    Loop
'    If lvwRole.ListItems.Count > 0 Then
'        lvwRole.ListItems(1).Selected = True
'    Else
'        cmdGrant.Enabled = False
'    End If
'    rsTemp.Close
'    Call SetEnable
'End Sub

Private Sub FillSystem()
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    rsTemp.CursorLocation = adUseClient
    
    '��ʾ�������е�ϵͳ

    Set rsTemp = zlGetRegSystems
    cmbSystem.Clear
    Do Until rsTemp.EOF
        cmbSystem.AddItem rsTemp("����") & " v" & rsTemp("�汾��") & "��" & rsTemp("���") & "��"
        cmbSystem.ItemData(cmbSystem.NewIndex) = rsTemp("���")
        If rsTemp("������") = UCase(gstrUserName) And cmbSystem.ListIndex < 0 Then
            cmbSystem.ListIndex = cmbSystem.NewIndex
        End If
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    '������ϵͳ�ǳ���̶���
    If (gobjRegister.zlRegTool And 2) = 2 Then cmbSystem.AddItem "�Զ��屨��"
    cmbSystem.AddItem "��������"
    cmbSystem.AddItem "ȡ������"
    cmbSystem.AddItem "��������"
    If cmbSystem.ListIndex < 0 Then cmbSystem.ListIndex = 0
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Sub FillModule()
    Dim rsTemp As New ADODB.Recordset
    Dim lst As ListItem
    Dim strRole As String
    
    If cmbSystem.ListCount = 0 Then Exit Sub
    
    LockWindowUpdate lvwModule.hwnd
    
    lvwModule.ColumnHeaders.Clear
    lvwModule.ListItems.Clear
    If Not lvwRole.SelectedItem Is Nothing Then
        strRole = lvwRole.SelectedItem.Key
    End If
    '�����б���
    With lvwModule.ColumnHeaders
        If cmbSystem.Text = "��������" Then
            lblModule.Caption = "�ɹ���ı����"
            .Add , , "�����", "1200"
            .Add , , "����ϵͳ", "2100"
            .Add , , "˵��", "2500"
        ElseIf cmbSystem.Text = "ȡ������" Then
            lblModule.Caption = "�ɵ��õĺ���"
            .Add , , "������", "1200"
            .Add , , "������", "1500"
            .Add , , "����ϵͳ", "2100"
            .Add , , "˵��", "2500"
        ElseIf cmbSystem.Text = "��������" Then
            lblModule.Caption = "����Ȩ�Ļ�������"
            .Add , , "���", "800"
            .Add , , "����", "1500"
            .Add , , "˵��", "2100"
            .Add , , "��Ȩ����", "1500"
        Else
            lblModule.Caption = "����Ȩģ��"
            .Add , , "���", "800"
            .Add , , "����", "1500"
            .Add , , "˵��", "2100"
            .Add , , "��Ȩ����", "1500"
        End If
    End With
    lnModuel.X1 = lblModule.Left + lblModule.Width
    
    If strRole = "" Then
        '��ɫΪ�գ��˳�
        LockWindowUpdate 0
        Exit Sub
    End If
    
    If cmbSystem.Text = "��������" Then
        '��ʾ�ý�ɫ�ܷ��ʵĻ�����
        gstrSQL = "select T.ϵͳ,T.����,T.˵�� from " & _
                "(SELECT S.����||'��'||S.���||'��' as ϵͳ,S.������,B.����,B.˵�� FROM zlSystems S,zlBaseCode B where B.ϵͳ=S.���) T,USER_TAB_PRIVS R " & _
                "where T.������=R.OWNER And T.����=R.TABLE_NAME And R.GRANTEE='" & strRole & _
                "' And R.PRIVILEGE in ('SELECT','INSERT','UPDATE','DELETE') " & _
                "GROUP BY T.ϵͳ,T.����,T.˵�� " & _
                "Having Count(R.PRIVILEGE) = 4"
        rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
        Do Until rsTemp.EOF
            Set lst = lvwModule.ListItems.Add(, , rsTemp("����"))
            lst.SubItems(1) = rsTemp("ϵͳ")
            lst.SubItems(2) = IIf(IsNull(rsTemp("˵��")), "", rsTemp("˵��"))
            rsTemp.MoveNext
        Loop
    ElseIf cmbSystem.Text = "ȡ������" Then
        '��ʾ�ý�ɫ�ܷ��ʵĻ�����
        gstrSQL = "select S.����||'��'||S.���||'��' as ϵͳ,S.������,F.������,F.������,F.˵�� " & _
                  " from zlSystems S,zlFunctions F,USER_TAB_PRIVS R " & _
                  " where  F.ϵͳ=S.��� And S.������=R.OWNER And Upper(F.������)=R.TABLE_NAME And R.GRANTEE='" & strRole & "' And R.PRIVILEGE ='EXECUTE'"
        rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
        Do Until rsTemp.EOF
            Set lst = lvwModule.ListItems.Add(, , rsTemp("������"))
            lst.SubItems(1) = rsTemp("������")
            lst.SubItems(2) = rsTemp("ϵͳ")
            lst.SubItems(3) = IIf(IsNull(rsTemp("˵��")), "", rsTemp("˵��"))
            rsTemp.MoveNext
        Loop
    ElseIf cmbSystem.Text = "��������" Then
        '��ʾ�ý�ɫ�ܷ��ʵĻ�������
        gstrSQL = "select P.���,P.����,P.˵��,R.���� from " & _
                "zlRoleGrant R,zlPrograms P " & _
                "where R.ϵͳ is Null And P.���=R.��� And R.��ɫ='" & strRole & _
                "'  And P.ϵͳ is Null And P.���<100 And P.���� is Null " & _
                " Order By P.���"
        rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
        On Error Resume Next
        Do Until rsTemp.EOF
            Set lst = lvwModule.ListItems.Add(, "C" & rsTemp("���"), rsTemp("���"))
            If err <> 0 Then
                err.Clear
                If rsTemp("����") <> "����" Then
                    Set lst = lvwModule.ListItems("C" & rsTemp("���"))
                    lst.SubItems(3) = IIf(lst.SubItems(3) = "", "", lst.SubItems(3) & ",") & rsTemp("����")
                End If
            Else
                lst.SubItems(1) = rsTemp("����")
                lst.SubItems(2) = IIf(IsNull(rsTemp("˵��")), "", rsTemp("˵��"))
                If rsTemp("����") <> "����" Then
                    lst.SubItems(3) = rsTemp("����")
                End If
            End If
            rsTemp.MoveNext
        Loop
    Else
        '��ʾ�ý�ɫ�ܷ��ʵ�ģ��
        gstrSQL = "select P.���,P.����,P.˵��,R.���� from " & _
                "zlRoleGrant R,zlPrograms P " & _
                "where nvl(R.ϵͳ,0)=nvl(P.ϵͳ,0) And P.���=R.��� And P.���>=100 And R.��ɫ='" & strRole & "'  And " & _
                IIf(cmbSystem.Text = "�Զ��屨��", " P.ϵͳ is Null ", " P.ϵͳ=" & cmbSystem.ItemData(cmbSystem.ListIndex)) & _
                " Order By P.���"
        rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
        On Error Resume Next
        Do Until rsTemp.EOF
            Set lst = lvwModule.ListItems.Add(, "C" & rsTemp("���"), rsTemp("���"))
            If err <> 0 Then
                err.Clear
                If rsTemp("����") <> "����" Then
                    Set lst = lvwModule.ListItems("C" & rsTemp("���"))
                    lst.SubItems(3) = IIf(lst.SubItems(3) = "", "", lst.SubItems(3) & ",") & rsTemp("����")
                End If
            Else
                lst.SubItems(1) = rsTemp("����")
                lst.SubItems(2) = IIf(IsNull(rsTemp("˵��")), "", rsTemp("˵��"))
                If rsTemp("����") <> "����" Then
                    lst.SubItems(3) = rsTemp("����")
                End If
            End If
            rsTemp.MoveNext
        Loop
    End If
    
    LockWindowUpdate 0
End Sub

Private Sub SetEnable()
'���ø�����ť��Enable����
    Dim blnHave As Boolean
    Dim i As Long
    Dim lstItem As ListItem
    blnHave = Not lvwRole.SelectedItem Is Nothing
    mnuPopuModify.Enabled = tvwGroups.SelectedItem.Key <> "Root"
    chkOnlyShowNOGroups.Enabled = Not mnuPopuModify.Enabled
    mnuPopuDeleteGroups.Enabled = mnuPopuModify.Enabled
    'cmdDelete.Enabled = cmdAdd.Enabled And blnHave
    cmdGrant.Enabled = blnHave
    cmdUser.Enabled = blnHave
    cmdCopy.Enabled = blnHave
    cmdGrantAll.Enabled = (gblnOwner = True)
    mnuPopuRoleDelete.Enabled = blnHave
    blnHave = False
    For Each lstItem In lvwRole.ListItems
        If lstItem.Selected = True Then
            blnHave = True
            Exit For
        End If
    Next
    For i = 1 To mnuPopuRoleMoveGroups.UBound
          mnuPopuRoleMoveGroups(i).Enabled = blnHave
        If UCase(mnuPopuRoleMoveGroups(i).Tag) = UCase(tvwGroups.SelectedItem.Key) Then
            mnuPopuRoleMoveGroups(i).Enabled = False
        End If
    Next
End Sub


Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = True
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'�������ڵ��ã�ʵ�־���Ĵ�ӡ����
'���û�пɴ�ӡ�ģ�������һ���յĽӿ�

'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Dim objPrint As zlPrintLvw
    
    If lvwRole.SelectedItem Is Nothing Then Exit Sub
    Set objPrint = New zlPrintLvw
    objPrint.Title.Text = "��ɫȨ��"
    Set objPrint.Body.objData = lvwModule
    objPrint.UnderAppItems.Add "��ɫ��" & lvwRole.SelectedItem.Text
    objPrint.UnderAppItems.Add "��Ȩϵͳ��" & cmbSystem.Text
    objPrint.BelowAppItems.Add "��ӡʱ�䣺" & Format(CurrentDate, "yyyy��MM��dd��")
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrViewLvw objPrint, 1
          Case 2
              zlPrintOrViewLvw objPrint, 2
          Case 3
              zlPrintOrViewLvw objPrint, 3
      End Select
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If

End Sub

Private Sub lvwRole_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDelete
        Call mnuPopuRoleDelete_Click
    End Select
End Sub

Private Sub lvwRole_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        '����ͼ��
        If lvwRole.SelectedItem Is Nothing Then Exit Sub
        Set lvwRole.DragIcon = lvwRole.SelectedItem.CreateDragImage
        lvwRole.Drag 1
    End If
End Sub

Private Sub lvwRole_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Exit Sub
    PopupMenu mnuPopuRole
End Sub

Private Sub mnuPopuAddGroups_Click()
    Call cmdNewGroup_Click
End Sub

Private Sub mnuPopuDeleteGroups_Click()
     'ɾ����
     Call DeleteRoleGroups
     Call LoadMenus
End Sub

Private Sub mnuPopuModify_Click()
    '����
    If tvwGroups.SelectedItem.Key <> "Root" Then
        Call tvwGroups.StartLabelEdit
    End If
End Sub

Private Sub mnuPopuRoleAdd_Click()
    Call cmdAdd_Click
End Sub

Private Sub mnuPopuRoleDelete_Click()
    'ɾ����ɫ
    If cmdAdd.Enabled = False Then Exit Sub
    Call DeleteRole
End Sub

Private Sub mnuPopuRoleMoveGroups_Click(Index As Integer)
    Dim str���� As String
    If mnuPopuRoleMoveGroups(Index).Tag = "" Then Exit Sub
    str���� = UCase(Mid(mnuPopuRoleMoveGroups(Index).Tag, 2))
    
    If str���� = UCase("oot") Or str���� = "���н�ɫ" Then
        If str���� = Mid(tvwGroups.SelectedItem.Key, 2) Then Exit Sub
        If MsgBox("����Ҫ����ɫ��" & lvwRole.SelectedItem.Text & "...�� �Ƴ�������?", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
        str���� = ""
    Else
        If str���� = Mid(tvwGroups.SelectedItem.Key, 2) Then Exit Sub
        If MsgBox("����Ҫ����ɫ��" & lvwRole.SelectedItem.Text & "...�� �ƶ����顰" & str���� & "������?", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
    End If
    If MoveToGroups(str����) = False Then Exit Sub
    Call SetEnable
End Sub

Private Sub picHLine_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Me.picHLine.BackColor = &H8000000F: Me.picHLine.Top = Me.picHLine.Top + y
End Sub

Private Sub picHLine_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.picHLine.BackColor = Me.BackColor
    If Me.picHLine.Top < 2500 Then Me.picHLine.Top = 2500
    If Me.picHLine.Top > Me.ScaleHeight - 1500 Then Me.picHLine.Top = Me.ScaleHeight - 1500
    mblnMoveTop = True
    Call Form_Resize
    mblnMoveTop = False
End Sub
 
Private Sub tvwGroups_AfterLabelEdit(Cancel As Integer, NewString As String)
    err = 0: On Error GoTo ErrHand:
    Dim strKey As String
    strKey = UCase(Mid(tvwGroups.SelectedItem.Key, 2))
    If strKey = NewString Then Exit Sub
    '����:zlTools.b_Rolegroupmgr.Rolegroup_Delete(
    '    ����_Old_In In ZlRolegroups.����%Type,
    '    ����_New_In In ZlRolegroups.����%Type
    gstrSQL = "zlTools.b_Rolegroupmgr.Rolegroup_Rename("
    gstrSQL = gstrSQL & "'" & strKey & "',"
    gstrSQL = gstrSQL & "'" & UCase(NewString) & "')"
    ExecuteProcedure gstrSQL, Me.Caption
    tvwGroups.SelectedItem.Key = "K" & NewString
    Call LoadMenus
    
    Exit Sub
ErrHand:
    Cancel = True
    Call ShowErrHand
End Sub
Private Sub tvwGroups_BeforeLabelEdit(Cancel As Integer)
    If Me.tvwGroups.SelectedItem.Key = "Root" Then
        Cancel = True
    End If
End Sub
Private Sub DeleteRoleGroups()
    '---------------------------------------------------------------------------------------------------------
    '����:ɾ����
    '����:���˺�
    '����:2007/06/15
    '---------------------------------------------------------------------------------------------------------
    Dim strRoleGroupName As String
    Dim intIndex As Integer
     
    If tvwGroups.SelectedItem Is Nothing Then Exit Sub
    If tvwGroups.SelectedItem.Key = "Root" Then Exit Sub
    
    strRoleGroupName = tvwGroups.SelectedItem.Text
    intIndex = tvwGroups.SelectedItem.Index
    
    If MsgBox("���Ҫɾ����" & strRoleGroupName & "���Ľ�ɫ����", vbDefaultButton2 Or vbQuestion Or vbYesNo, gstrSysName) = vbNo Then Exit Sub
    
    err = 0: On Error GoTo ErrHand:
    '����:zlTools.b_Rolegroupmgr.Rolegroup_Delete(����_In In ZlRolegroups.����%Type)
    gstrSQL = "zlTools.b_Rolegroupmgr.Rolegroup_Delete("
    gstrSQL = gstrSQL & "'" & UCase(strRoleGroupName) & "')"
    ExecuteProcedure gstrSQL, Me.Caption
    
    tvwGroups.Nodes.Remove intIndex
    If tvwGroups.Nodes.Count > 0 Then
        If intIndex > tvwGroups.Nodes.Count Then intIndex = tvwGroups.Nodes.Count
        tvwGroups.Nodes(intIndex).Selected = True
        tvwGroups.Nodes(intIndex).EnsureVisible
    End If
    If tvwGroups.SelectedItem Is Nothing Then
        Call FillRoleData("")
    Else
        Call tvwGroups_NodeClick(tvwGroups.SelectedItem)
    End If

    Call SetEnable
    Exit Sub
ErrHand:
    Call ShowErrHand
End Sub

Private Sub tvwGroups_DragDrop(Source As Control, x As Single, y As Single)
    Dim str���� As String, str��ɫ As String, intIndex As Integer
    Dim lstItem As ListItem
    Dim strKeys As String
    Dim arrVar As Variant
    Dim i As Long
     err = 0: On Error GoTo ErrHand:
     
    If Source Is lvwRole And Not tvwGroups.DropHighlight Is Nothing Then
        intIndex = -1
        str���� = Mid(tvwGroups.DropHighlight.Key, 2)
        Set tvwGroups.DropHighlight = Nothing
        tvwGroups.DropHighlight = tvwGroups.SelectedItem

        If str���� = "oot" Then
            If str���� = Mid(tvwGroups.SelectedItem.Key, 2) Then Exit Sub
            If MsgBox("����Ҫ����ɫ��" & Source.SelectedItem.Text & "...�� �Ƴ�������?", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
            str���� = ""
        Else
            If str���� = Mid(tvwGroups.SelectedItem.Key, 2) Then Exit Sub
            If MsgBox("����Ҫ����ɫ��" & Source.SelectedItem.Text & "...�� �ƶ����顰" & str���� & "������?", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
        End If

        gcnOracle.BeginTrans
        strKeys = ""
        For Each lstItem In lvwRole.ListItems
            If lstItem.Selected Then
                If intIndex < 0 Then
                    intIndex = lstItem.Index
                End If
                str��ɫ = lstItem.Key
                strKeys = strKeys & "'" & lstItem.Key

                If MoveToRoleGroup(str����, str��ɫ) = False Then
                    gcnOracle.RollbackTrans
                    Exit Sub
                End If
            End If
        Next
        gcnOracle.CommitTrans
        If strKeys <> "" Then strKeys = Mid(strKeys, 2)
        If tvwGroups.SelectedItem.Key <> "Root" Or (tvwGroups.SelectedItem.Key = "Root" And chkOnlyShowNOGroups.value = 1) Then
            arrVar = Split(strKeys, "'")
            For i = 0 To UBound(arrVar)
                lvwRole.ListItems.Remove arrVar(i)
            Next

            If lvwRole.ListItems.Count > 0 Then
                If intIndex > lvwRole.ListItems.Count Then intIndex = lvwRole.ListItems.Count
                lvwRole.ListItems(intIndex).Selected = True
            End If
        End If
        Call FillModule
    End If
    Call SetEnable
    tvwGroups.Refresh
     
    Set tvwGroups.DropHighlight = Nothing
    Exit Sub
ErrHand:
    Set tvwGroups.DropHighlight = Nothing
    Call ShowErrHand
End Sub
Private Function MoveToGroups(ByVal str���� As String) As Boolean
    '-------------------------------------------------------------------------------------------------------------------
    '����:��ָ����ɫ�ƶ�������
    '����:str��-�Ƶ��������
    '�ƶ��ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2007/06/19
    '-------------------------------------------------------------------------------------------------------------------
    Dim intIndex As Integer
    Dim strKeys  As String
    Dim lstItem As ListItem
    Dim str��ɫ As String
    Dim arrVar As Variant
    Dim i As Long
    MoveToGroups = False
    intIndex = -1
    gcnOracle.BeginTrans
    strKeys = ""
    For Each lstItem In lvwRole.ListItems
        If lstItem.Selected Then
            If intIndex < 0 Then
                intIndex = lstItem.Index
            End If
            str��ɫ = lstItem.Key
            strKeys = strKeys & "'" & lstItem.Key
            If MoveToRoleGroup(str����, str��ɫ) = False Then
                gcnOracle.RollbackTrans
                Exit Function
            End If
        End If
    Next
    gcnOracle.CommitTrans
    If strKeys <> "" Then strKeys = Mid(strKeys, 2)
    If tvwGroups.SelectedItem.Key <> "Root" Or (tvwGroups.SelectedItem.Key = "Root" And chkOnlyShowNOGroups.value = 1) Then
        arrVar = Split(strKeys, "'")
        For i = 0 To UBound(arrVar)
            lvwRole.ListItems.Remove arrVar(i)
        Next
        If lvwRole.ListItems.Count > 0 Then
            If intIndex > lvwRole.ListItems.Count Then intIndex = lvwRole.ListItems.Count
            lvwRole.ListItems(intIndex).Selected = True
        End If
    End If
    MoveToGroups = True
End Function
Private Function MoveToRoleGroup(ByVal str�� As String, str��ɫ As String) As Boolean
    '-------------------------------------------------------------------------------------------------------------------
    '����:��ָ���Ľ�ɫ�Ƶ�����
    '����:str��-�Ƶ��������
    '     str��ɫ-ָ���Ľ�ɫ
    '�ƶ��ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2007/06/19
    '-------------------------------------------------------------------------------------------------------------------
    err = 0: On Error GoTo ErrHand:
    
    '���˺�:20070615����
    '���̲���:zlTools.b_Rolegroupmgr.RoletoRolegroup
    '        ����_In In ZlRolegroups.����%Type,
    '        ��ɫ_In In ZlRolegroups.��ɫ%Type := Null
    gstrSQL = "zlTools.b_Rolegroupmgr.RoletoRolegroup("
    gstrSQL = gstrSQL & IIf(str�� = "", "Null", "'" & UCase(str��) & "'") & ","
    gstrSQL = gstrSQL & "'" & UCase(str��ɫ) & "')"
    ExecuteProcedure gstrSQL, Me.Caption
    MoveToRoleGroup = True
    Exit Function
ErrHand:
    Call ShowErrHand
End Function

Private Sub tvwGroups_DragOver(Source As Control, x As Single, y As Single, State As Integer)
      Dim objOver As Node
      If Source Is lvwRole Then
           Set objOver = tvwGroups.HitTest(x, y)
            If Not objOver Is Nothing Then
                Set tvwGroups.DropHighlight = objOver
                tvwGroups.DropHighlight.EnsureVisible
            Else
                Set tvwGroups.DropHighlight = Nothing
            End If
      End If
End Sub

Private Sub tvwGroups_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDelete
        mnuPopuDeleteGroups_Click
    Case vbKeyF2
        '����
        If tvwGroups.SelectedItem.Key <> "Root" Then
            Call tvwGroups.StartLabelEdit
        End If
    End Select
End Sub

Private Sub tvwGroups_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Exit Sub
    PopupMenu mnuPopu
End Sub

Private Sub tvwGroups_NodeClick(ByVal Node As MSComctlLib.Node)
    '---------------------------------------------------------------------------------------------------------
    '��ȡ��Ӧ�Ľ�ɫȨ��
    '---------------------------------------------------------------------------------------------------------
    Dim strKey As String
    Call SetEnable
    strKey = IIf(Node.Key = "Root", "", Mid(Node.Key, 2))
    Call FillRoleData(strKey)
    Call FillModule
    Call SetEnable
End Sub

Private Function FillRoleData(ByVal str���� As String) As Boolean
    '---------------------------------------------------------------------------------------------------------
    '����:��������,��ȡ��ɫ��Ϣ,����䵽lvw��
    '����:str����:<>""ָ������,=""��ʾ���н�ɫ
    '����:���سɹ�,����true,���򷵻�False
    '---------------------------------------------------------------------------------------------------------
    Dim rsGroups As New ADODB.Recordset
    Dim objItem As ListItem
    Dim blnGroups As Boolean '�Ƿ�Ҫ������
    Dim strFiler  As String
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    strFiler = UCase(Trim(txtSearch.Text))
     
    '��ʾ���Խ�����Ȩ�Ľ�ɫ
    If gblnDBA = True Then
        gstrSQL = _
            " Select User as Grantee,'YES' as Admin_Option,Substr(A.Role,4) as RoleName," & _
            " A.Role,zlSpellCode(Substr(A.Role,4)) as ����  " & _
            " From DBA_Roles A,zlTools.zlRoleGroups B " & _
            " Where Upper(A.Role) Like 'ZL_%' And A.Role=B.��ɫ" & _
            IIf(str���� = "", "(+)", " And B.����='" & str���� & "' And B.��ɫ is Not Null") & _
            " Order by A.Role"
    Else
        '�����ߺ�SYSTEM�е�ZL��ɫ
        gstrSQL = _
            " Select UserName As Grantee,A.Admin_Option,Substr(A.Granted_Role ,4) as RoleName," & _
            " A.Granted_Role as Role,zlSpellCode(Substr(A.Granted_Role ,4)) as ���� " & _
            " From User_Role_Privs A,zlRoleGroups B " & _
            " Where A.Granted_Role Like 'ZL_%' And A.Granted_Role=B.��ɫ" & _
            IIf(str���� = "", "(+)", " And B.����='" & str���� & "' And B.��ɫ is Not Null") & _
            " Union ALL" & _
            " Select A.Grantee,A.Admin_Option,Substr(A.Granted_Role ,4) as RoleName," & _
            " A.Granted_Role as Role,zlSpellCode(Substr(A.Granted_Role ,4)) as ���� " & _
            " From DBA_Role_Privs A,zlRoleGroups B " & _
            " Where A.Granted_Role Like 'ZL_%' And A.Granted_Role=B.��ɫ" & _
            IIf(str���� = "", "(+)", " And B.����='" & str���� & "' And B.��ɫ is Not Null") & _
            " And A.Grantee='SYSTEM' And Not Exists(" & _
                " Select 1 From DBA_Role_Privs X Where X.Granted_Role=A.Granted_Role And Grantee='" & gstrUserName & "')" & _
            " Order by Role"
    End If
    Call OpenRecordset(mrsRole, gstrSQL, Me.Caption)
    If strFiler <> "" Then
        mrsRole.Filter = "RoleName Like '" & strFiler & "%' or ���� Like '" & strFiler & "%'"
    End If
    
    blnGroups = False
    If chkOnlyShowNOGroups.Enabled And chkOnlyShowNOGroups.value = 1 Then
        'ֻ��ʾδ�������
        strSQL = "SELECT ��ɫ FROM zlRoleGroups where ��ɫ is Not Null "
        OpenRecordset rsGroups, strSQL, Me.Caption
        blnGroups = True
    End If
    
    lvwRole.ListItems.Clear
    Do Until mrsRole.EOF
        Set objItem = Nothing
        If blnGroups Then
            rsGroups.Filter = "��ɫ='" & UCase(Nvl(mrsRole!Role)) & "'"
            If rsGroups.EOF Then
                Set objItem = lvwRole.ListItems.Add(, Nvl(mrsRole!Role), Nvl(mrsRole!RoleName), "Role", "Role")
            End If
        Else
            Set objItem = lvwRole.ListItems.Add(, Nvl(mrsRole!Role), Nvl(mrsRole!RoleName), "Role", "Role")
        End If
        If Not objItem Is Nothing Then
            objItem.SubItems(1) = Nvl(mrsRole!Grantee)
            objItem.SubItems(2) = Nvl(mrsRole!Admin_Option)
        End If
        mrsRole.MoveNext
    Loop
    
    If lvwRole.ListItems.Count > 0 Then
        lvwRole.ListItems(1).Selected = True
    Else
        cmdGrant.Enabled = False
    End If
    mrsRole.Filter = 0
    Call SetEnable
    FillRoleData = True
    
    Exit Function
errHandle:
    MsgBox "����" & err.Number & vbCrLf & vbCrLf & err.Description, vbInformation, gstrSysName
End Function

Private Function SearchRole(ByVal strFilter As String) As Boolean
    '----------------------------------------------------------------------------------------------------------------------------
    '����:���ǳ���Ӧ�Ľ�ɫ
    '����:strFilter-���˴�
    '����:�ɹ�,����ture,���򷵻�False
    '----------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsGroups As New ADODB.Recordset
    Dim blnGroups As Boolean '�Ƿ�Ҫ������
    SearchRole = True
    If mrsRole Is Nothing Then Exit Function
    If mrsRole.State <> 1 Then Exit Function
    If mrsRole.RecordCount = 0 Then Exit Function
    
    blnGroups = False
    If chkOnlyShowNOGroups.Enabled And chkOnlyShowNOGroups.value = 1 Then
        'ֻ��ʾδ�������
        strSQL = "SELECT ��ɫ FROM zlRoleGroups where ��ɫ is Not Null "
        OpenRecordset rsGroups, strSQL, Me.Caption
        blnGroups = True
    End If
    
    strFilter = UCase(strFilter)
    SearchRole = False
    If strFilter = "" Then
    Else
        mrsRole.Filter = "RoleName Like '" & strFilter & "%' or ���� Like '" & strFilter & "%'"
    End If
    lvwRole.ListItems.Clear
    Do Until mrsRole.EOF
        If blnGroups Then
            rsGroups.Filter = "��ɫ='" & UCase(Nvl(mrsRole!Role)) & "'"
            If rsGroups.EOF Then
                lvwRole.ListItems.Add , Nvl(mrsRole!Role), Nvl(mrsRole!RoleName), "Role", "Role"
            End If
        Else
            lvwRole.ListItems.Add , Nvl(mrsRole!Role), Nvl(mrsRole!RoleName), "Role", "Role"
        End If
        mrsRole.MoveNext
    Loop
    If lvwRole.ListItems.Count > 0 Then
        lvwRole.ListItems(1).Selected = True
    Else
        cmdGrant.Enabled = False
    End If
    Call SetEnable
    mrsRole.Filter = 0
    SearchRole = True
End Function

Private Sub txtSearch_Change()
    Call SearchRole(Trim(txtSearch.Text))
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}"
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Or KeyAscii = Asc("*") Or KeyAscii = Asc("_") Then
        KeyAscii = 0
    End If
End Sub
Private Sub FillRollGroup()
    '--------------------------------------------------------------------------------------------
    '����:���ؽ�ɫ��
    '����:���˺�
    '����:2007/06/15
    '--------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    Dim objNode As Node
    gstrSQL = "Select distinct ���� From zlRoleGroups"
    Call OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    With tvwGroups
        .Nodes.Clear
        Set objNode = tvwGroups.Nodes.Add(, 4, "Root", "���н�ɫ", 1, 1)
        objNode.Selected = True
        objNode.Expanded = True
        objNode.Sorted = True
        Call LoadMenu("���н�ɫ", "Root")
        Do While Not rsTemp.EOF
            Set objNode = tvwGroups.Nodes.Add("Root", 4, "K" & Nvl(rsTemp!����), Nvl(rsTemp!����), 1, 1)
            objNode.Tag = Nvl(rsTemp!����)
            objNode.Sorted = True
            Call LoadMenu(Nvl(rsTemp!����), Nvl(rsTemp!����))
            rsTemp.MoveNext
        Loop
    End With
    Call tvwGroups_NodeClick(Me.tvwGroups.SelectedItem)
End Sub
Private Sub LoadMenu(ByVal strTittle As String, ByVal strTag As String)
    '����:���ز˵�
        Dim varMenu As Variant
        Dim intCount As Integer
        
        Set varMenu = mnuPopuRoleMoveGroups
        intCount = varMenu.Count
        Load varMenu(intCount)
        varMenu(intCount).Caption = strTittle
        If strTag = "Root" Then
            varMenu(intCount).Tag = UCase(strTag)
        Else
            varMenu(intCount).Tag = UCase("K" & strTag)
        End If
        varMenu(intCount).Visible = True
        mnuPopuRoleMove.Visible = True
        varMenu(0).Visible = False
End Sub
Private Sub LoadMenus()
    Dim objNode As Node
    Call UnLoadMenus
    For Each objNode In tvwGroups.Nodes
        Call LoadMenu(objNode.Text, objNode.Key)
    Next
End Sub
Private Sub UnLoadMenus()
    '����:��ж�˵�
        Dim varMenu As Variant
        Dim intCount As Integer
        Set varMenu = mnuPopuRoleMoveGroups
        mnuPopuRoleMoveGroups(0).Visible = True
        mnuPopuRoleMove.Visible = True
        For intCount = 1 To mnuPopuRoleMoveGroups.UBound
            Unload varMenu(intCount)
        Next
        
End Sub

