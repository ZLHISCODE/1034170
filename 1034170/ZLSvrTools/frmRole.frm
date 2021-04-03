VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRole 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "角色授权管理"
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
      Caption         =   "只显示未分组角色(&B)"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   5250
      TabIndex        =   4
      Top             =   570
      Value           =   1  'Checked
      Width           =   2040
   End
   Begin VB.CommandButton cmdNewGroup 
      Caption         =   "新建组(&N)"
      Height          =   350
      Left            =   7575
      TabIndex        =   6
      Top             =   945
      Width           =   1155
   End
   Begin VB.CommandButton cmdGrantAll 
      Caption         =   "恢复权限"
      Height          =   350
      Left            =   7560
      TabIndex        =   9
      Top             =   1995
      Width           =   1155
   End
   Begin VB.CommandButton cmdGrant 
      Caption         =   "授权(&G)"
      Height          =   350
      Left            =   7575
      TabIndex        =   8
      Top             =   1650
      Width           =   1155
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "复制(&C)…"
      Height          =   350
      Left            =   7560
      TabIndex        =   10
      Top             =   2355
      Width           =   1155
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "增加角色(&A)"
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
      Caption         =   "修改角色的授权用户"
      Height          =   350
      Left            =   6585
      TabIndex        =   15
      Top             =   3975
      Width           =   1875
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "修改模块的使用权限"
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
         Key             =   "_角色"
         Object.Tag             =   "角色"
         Text            =   "角色"
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
      Caption         =   "角色授权管理"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "角色组信息"
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
      Caption         =   "搜索"
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
      Caption         =   "已授权模块"
      Height          =   180
      Left            =   135
      TabIndex        =   11
      Top             =   4320
      Width           =   900
   End
   Begin VB.Label lblSystem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "授权系统"
      Height          =   180
      Left            =   135
      TabIndex        =   12
      Top             =   4065
      Width           =   720
   End
   Begin VB.Menu mnuPopu 
      Caption         =   "弹出菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuPopuAddGroups 
         Caption         =   "新建组(&N)"
      End
      Begin VB.Menu mnuPopuModify 
         Caption         =   "修改组(&M)"
      End
      Begin VB.Menu mnuPopuDeleteGroups 
         Caption         =   "删除组(&D)"
      End
   End
   Begin VB.Menu mnuPopuRole 
      Caption         =   "弹出菜单角色"
      Visible         =   0   'False
      Begin VB.Menu mnuPopuRoleAdd 
         Caption         =   "增加角色(&N)"
      End
      Begin VB.Menu mnuPopuRoleDelete 
         Caption         =   "删除角色(&M)"
      End
      Begin VB.Menu mnuPopuRoleMove 
         Caption         =   "角色移到(&M)…"
         Begin VB.Menu mnuPopuRoleMoveGroups 
            Caption         =   "组1"
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
    '只显示未分配组的角色
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
    
    '当用户拥有的角色数量达到148个时，用户登录时会提示错误
    gstrSQL = "Select Count(*) as 数量 From DBA_Role_Privs Where Grantee='" & gstrUserName & "'"
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open gstrSQL, gcnOracle, adOpenKeyset, adLockReadOnly
    If Nvl(rsTemp!数量, 0) >= 148 Then
        If Not gblnOwner Then
            MsgBox "角色数量已达到最大限制，不能再增加。", vbInformation, gstrSysName
            Exit Sub
        Else
            '所有者角色数量达到限制时，借助Systeme用户创建
            If Not CheckRushHours("当前时段处于业务高峰期，新建角色可能会对系统使用造成一定影响，是否继续") Then
                Exit Sub
            End If
            'SYSTEM所创建的角色，不授予所有者
            gstrSQL = "Select Count(*) as 数量 From DBA_Role_Privs Where Grantee='SYSTEM'"
            Set rsTemp = New ADODB.Recordset
            rsTemp.CursorLocation = adUseClient
            rsTemp.Open gstrSQL, gcnOracle, adOpenKeyset, adLockReadOnly
            If Nvl(rsTemp!数量, 0) >= 148 Then
                MsgBox "角色数量已达到最大限制，不能再增加。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        blnLimited = True
    End If
    
    strRoleName = frmNameEdit.GetName(name角色)
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
        MsgBox "由于重名或命名错误或者角色数超过了数据库的参数限制" & vbCrLf & _
                "(可修改数据库启动参数调整角色数目)，导致角色增加失败。", vbExclamation, gstrSysName
    Else
        On Error GoTo errHandle
        Call GrantSpecialToRole(cnTemp, strRoleName, False, "", True)
        If tvwGroups.SelectedItem Is Nothing Then
        ElseIf tvwGroups.SelectedItem.Key <> "Root" Then
            '刘兴宏:20070615加入
            '过程参数:zlTools.b_Rolegroupmgr.RoletoRolegroup
            '        组名_In In ZlRolegroups.组名%Type,
            '        角色_In In ZlRolegroups.角色%Type := Null
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
    MsgBox "错误：" & err.Number & vbCrLf & vbCrLf & err.Description, vbInformation, gstrSysName
End Sub

Private Sub DeleteRole()
    '------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:删除角色
    '编制:刘兴宏
    '日期:2007/06/15
    '------------------------------------------------------------------------------------------------------------------------------------------------
    Dim strRoleName As String
    Dim intIndex As Integer
    
    On Error GoTo errHandle
    
    If lvwRole.SelectedItem Is Nothing Then Exit Sub
    
    strRoleName = lvwRole.SelectedItem.Key
    intIndex = lvwRole.SelectedItem.Index
    If MsgBox("真的要删除角色“" & strRoleName & "”吗？", vbDefaultButton2 Or vbQuestion Or vbYesNo, gstrSysName) = vbNo Then Exit Sub
    
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
    MsgBox "错误：" & err.Number & vbCrLf & vbCrLf & err.Description, vbInformation, gstrSysName
End Sub

Private Sub cmdCopy_Click()
    '------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:复制角色
    '编制:刘兴宏
    '日期:2007/06/15
    '------------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSourceRole As String
    Dim cnTemp As ADODB.Connection
    Dim rsTemp As ADODB.Recordset
    Dim strRoleName As String
    Dim blnLimited As Boolean
    
    On Error GoTo errHandle
    
    If lvwRole.SelectedItem Is Nothing Then Exit Sub
    
    '当用户拥有的角色数量达到148个时，用户登录时会提示错误
    gstrSQL = "Select Count(*) as 数量 From DBA_Role_Privs Where Grantee='" & gstrUserName & "'"
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open gstrSQL, gcnOracle, adOpenKeyset, adLockReadOnly
    If Nvl(rsTemp!数量, 0) >= 148 Then
        If Not gblnOwner Then
            MsgBox "角色数量已达到最大限制，不能再增加。", vbInformation, gstrSysName
            Exit Sub
        Else
            If Not CheckRushHours("当前时段处于业务高峰期，角色复制可能会对系统使用造成一定影响，是否继续") Then
                Exit Sub
            End If
            '所有者角色数量达到限制时，借助Systeme用户创建
            'SYSTEM所创建的角色，不授予所有者
            gstrSQL = "Select Count(*) as 数量 From DBA_Role_Privs Where Grantee='SYSTEM'"
            Set rsTemp = New ADODB.Recordset
            rsTemp.CursorLocation = adUseClient
            rsTemp.Open gstrSQL, gcnOracle, adOpenKeyset, adLockReadOnly
            If Nvl(rsTemp!数量, 0) >= 148 Then
                MsgBox "角色数量已达到最大限制，不能再增加。", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        blnLimited = True
    End If
    
    strSourceRole = lvwRole.SelectedItem.Key
    strRoleName = frmNameEdit.GetName(name角色)
    If strRoleName = "" Then Exit Sub
    strRoleName = "ZL_" & UCase(Trim(strRoleName))
 
    Set cnTemp = gcnOracle
    If blnLimited And gblnOwner Then
        Set gcnSystem = GetConnection("SYSTEM")
        If gcnSystem Is Nothing Then Exit Sub
        Set cnTemp = gcnSystem
    End If
 
    If Not CopyRole(cnTemp, strSourceRole, strRoleName) Then Exit Sub
    
    '重新授权
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
    MsgBox "错误：" & err.Number & vbCrLf & vbCrLf & err.Description, vbInformation, gstrSysName
End Sub

Private Function CopyRole(cnTemp As ADODB.Connection, ByVal strSourceRole As String, ByVal strTargetRole As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:拷贝角色权限成新的角色权限
    '参数:strSourceRole-源角色
    '     strTargetRole-目标角色
    '返回:复制成功,返回true,否则False
    '编制:刘兴宏
    '日期:2007/07/18
    '------------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    err = 0: On Error Resume Next
    cnTemp.Execute "Create Role " & strTargetRole & " Not Identified"
    If err <> 0 Then
        MsgBox "由于重名或命名错误或者目标角色数超过了数据库的参数限制" & vbCrLf & _
                "(可修改数据库启动参数调整角色数目)，导致角色增加失败。", vbExclamation, gstrSysName
        Exit Function
    End If
    err = 0: On Error GoTo ErrHand:
    Call GrantSpecialToRole(cnTemp, strTargetRole, False, "", True)
    '过程:zlTools.b_Rolegroupmgr.Role_Copy
    '    源角色_In   In zlRoleGrant.角色%Type,
    '    目标角色_In In zlRoleGrant.角色%Type
    gstrSQL = "zlTools.b_Rolegroupmgr.Role_Copy("
    gstrSQL = gstrSQL & "'" & UCase(strSourceRole) & "',"
    gstrSQL = gstrSQL & "'" & UCase(strTargetRole) & "')"
    ExecuteProcedure gstrSQL, Me.Caption
    CopyRole = True
    Exit Function
ErrHand:
    Call ShowErrHand
End Function

Private Function RoleGrant(ByVal str角色 As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------------------
    '功能:对指定的角色重新授权
    '参数:str角色-角色
    '返回:复制成功,返回true,否则False
    '编制:刘兴宏
    '日期:2007/07/18
    '------------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    Me.MousePointer = vbHourglass
    On Error GoTo errHandle
    

    '授予权限表中填写的权限
    Dim objclsPrivilege As New clsPrivilege
    Call objclsPrivilege.InitOracle(gcnOracle)
    Call objclsPrivilege.ReGrantPrivilege(str角色)
    Set objclsPrivilege = Nothing
    
    MousePointer = 0
    RoleGrant = True
    Exit Function
errHandle:
    MousePointer = 0
    MsgBox "当前用户的权限不能完成本操作。", vbInformation, gstrSysName
End Function

Private Sub cmdGrant_Click()
    If Not CheckRushHours("当前时段处于业务高峰期，角色授权可能会对系统使用造成一定影响，是否继续") Then
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
    
    If Not CheckRushHours("当前时段处于业务高峰期，恢复权限可能会对系统使用造成一定影响，是否继续") Then
        Exit Sub
    End If
    If MsgBox("本操作会对所有角色重新授权，这需要花一定时间(约1.5分钟左右)，" & vbCrLf & "是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    On Error GoTo errh
    On Error Resume Next
    '首先创建不存在的角色
    strSQL = "Select Distinct r.角色" & vbNewLine & _
            "From Zlsystems s, Zlrolegrant r" & vbNewLine & _
            "Where s.编号 = r.系统 And s.所有者 = User And r.角色 Not In (Select Granted_Role From User_Role_Privs)"

    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, App.Title)
    If err.Number <> 0 Then err.Clear
    If Not rsTmp Is Nothing Then
        Do While Not rsTmp.EOF
            '创建不存在的角色
            gcnOracle.Execute "Create Role " & rsTmp!角色 & " Not Identified"
            If err.Number = 0 Then
                '创建成功，将其增加到列表中
                Set lst = lvwRole.ListItems.Add(, rsTmp!角色 & "", Mid(rsTmp!角色 & "", 4), "Role", "Role")
                Call GrantSpecialToRole(gcnOracle, rsTmp!角色, False, "", True)
            Else
                err.Clear
            End If
            rsTmp.MoveNext
        Loop
    End If
    If err.Number <> 0 Then err.Clear
    On Error GoTo errh
    '开始授权
    Call ReGrantToRole(gcnOracle, "", True)
    '显示授权清单
    If Not lst Is Nothing Then
        lst.Selected = True
        Call FillModule
    End If
    MsgBox "所有角色重新授权完成！", vbInformation, gstrSysName
    MousePointer = 0
    Exit Sub
errh:
    If 0 = 1 Then
        Resume
    End If
    MousePointer = 0
    MsgBox "当前用户的权限不能完成本操作。", vbInformation, gstrSysName
End Sub

Private Sub cmdModify_Click()
    frmProgPriv.ProgPriv cmbSystem.ItemData(cmbSystem.ListIndex)
End Sub
Private Sub cmdNewGroup_Click()
    Dim strGroupsName As String
    Dim lst As ListItem
    Dim objNode As Node
ReDo:
    strGroupsName = frmNameEdit.GetName(name组名)
    
    If strGroupsName = "" Then Exit Sub
    If ActualLen(strGroupsName) > 30 Then
        MsgBox "你输入的角色名称不能大于30个字符或15个汉字,请检查!", vbDefaultButton1 + vbInformation, gstrSysName
        GoTo ReDo:
    End If
    strGroupsName = UCase(Trim(strGroupsName))
    
    '刘兴宏:20070615加入
    '过程参数:zlTools.b_Rolegroupmgr.Rolegroup_Add(组名_In In ZlRolegroups.组名%Type)
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

'     '新增组
'     '方法是:在Tvw中增加一个新的角色名称
'     Dim objNode As Node
'     Dim int序号 As Integer
'     Dim str组名 As String
'ReDo:
'    Err = 0: On Error Resume Next
'    int序号 = int序号 + 1
'    str组名 = "新建组:" & int序号
'     Set objNode = tvwGroups.Nodes.Add(, "Root", str组名, str组名)
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
    '功能:获取错误信息
    '------------------------------------------------------------------------------------------
    Dim strNote As String, lngErrNum As Long
    If gcnOracle.Errors.Count <> 0 Then
        strNote = gcnOracle.Errors(0).Description
        If InStr(UCase(strNote), "[ZLSOFT]") > 0 Then
            '日志变量
            lngErrNum = gcnOracle.Errors(0).NativeError
            MsgBox Split(strNote, "[ZLSOFT]")(1), vbExclamation, App.Title
            Exit Sub
        Else
            MsgBox "注意:" & vbCrLf & "操作发生如下错误:" & vbCrLf & err.Description, vbExclamation, App.Title
        End If
    Else
        MsgBox "注意:" & vbCrLf & "操作发生如下错误:" & vbCrLf & err.Description, vbExclamation, App.Title
    End If
End Sub
Private Sub cmdUser_Click()
    If Not CheckRushHours("当前时段处于业务高峰期，修改角色的授权用户可能会对系统使用造成一定影响，是否继续") Then
        Exit Sub
    End If
    If lvwRole.SelectedItem Is Nothing Then Exit Sub
    Call frmRoleUser.ShowEdit(lvwRole.SelectedItem.Text)
End Sub


Private Sub Form_Activate()
    Dim lngTop As Long
    mblnMoveTop = False
    If mblnFirst = False Then Exit Sub
    '个性化设置
    lngTop = Val(GetSetting("ZLSOFT", "公共模块\服务器管理工具\角色管理", "PicHLine_TOP", "4170"))
    picHLine.Top = lngTop
    mblnMoveTop = True
    Call Form_Resize
    mblnMoveTop = False
    mblnFirst = False
End Sub

Private Sub Form_Load()
   Dim rsTemp As New ADODB.Recordset
   Dim lngTop As Long
   
    '判断该用户能否创建角色
    gstrSQL = _
        " Select 1 From User_Sys_Privs Where Privilege='CREATE ROLE'" & _
        " Union" & _
        " Select 1 From Role_Sys_Privs Where Privilege='CREATE ROLE'"
    Call OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    cmdAdd.Enabled = rsTemp.RecordCount > 0
    
    Call Get注册码
    Call FillRollGroup
    Call FillSystem
    mblnFirst = True
End Sub

Private Sub cmbSystem_Click()
    '系统不为空时才能修改
    cmdModify.Enabled = cmbSystem.ItemData(cmbSystem.ListIndex) <> 0
    
    Call FillModule
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '个性化设置
    SaveSetting "ZLSOFT", "公共模块\服务器管理工具\角色管理", "PicHLine_TOP", picHLine.Top
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
'    '判断该用户能否创建角色
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
'    '显示可以进行授权的角色
'    If gblnDBA = True Then
'        gstrSQL = "select * from DBA_Roles where Upper(Role) Like 'ZL_%'"
'    Else
'        gstrSQL = "select Granted_Role as Role from user_Role_privs " & _
'            "where Granted_Role Like 'ZL_%'" 'ADMIN_OPTION='YES'选项可以不加
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
    
    '显示可以所有的系统

    Set rsTemp = zlGetRegSystems
    cmbSystem.Clear
    Do Until rsTemp.EOF
        cmbSystem.AddItem rsTemp("名称") & " v" & rsTemp("版本号") & "（" & rsTemp("编号") & "）"
        cmbSystem.ItemData(cmbSystem.NewIndex) = rsTemp("编号")
        If rsTemp("所有者") = UCase(gstrUserName) And cmbSystem.ListIndex < 0 Then
            cmbSystem.ListIndex = cmbSystem.NewIndex
        End If
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    '有两种系统是程序固定的
    If (gobjRegister.zlRegTool And 2) = 2 Then cmbSystem.AddItem "自定义报表"
    cmbSystem.AddItem "基础工具"
    cmbSystem.AddItem "取数函数"
    cmbSystem.AddItem "基础编码"
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
    '更新列表项
    With lvwModule.ColumnHeaders
        If cmbSystem.Text = "基础编码" Then
            lblModule.Caption = "可管理的编码表"
            .Add , , "编码表", "1200"
            .Add , , "所属系统", "2100"
            .Add , , "说明", "2500"
        ElseIf cmbSystem.Text = "取数函数" Then
            lblModule.Caption = "可调用的函数"
            .Add , , "函数名", "1200"
            .Add , , "中文名", "1500"
            .Add , , "所属系统", "2100"
            .Add , , "说明", "2500"
        ElseIf cmbSystem.Text = "基础工具" Then
            lblModule.Caption = "已授权的基础工具"
            .Add , , "序号", "800"
            .Add , , "标题", "1500"
            .Add , , "说明", "2100"
            .Add , , "授权功能", "1500"
        Else
            lblModule.Caption = "已授权模块"
            .Add , , "序号", "800"
            .Add , , "标题", "1500"
            .Add , , "说明", "2100"
            .Add , , "授权功能", "1500"
        End If
    End With
    lnModuel.X1 = lblModule.Left + lblModule.Width
    
    If strRole = "" Then
        '角色为空，退出
        LockWindowUpdate 0
        Exit Sub
    End If
    
    If cmbSystem.Text = "基础编码" Then
        '显示该角色能访问的基础表
        gstrSQL = "select T.系统,T.表名,T.说明 from " & _
                "(SELECT S.名称||'（'||S.编号||'）' as 系统,S.所有者,B.表名,B.说明 FROM zlSystems S,zlBaseCode B where B.系统=S.编号) T,USER_TAB_PRIVS R " & _
                "where T.所有者=R.OWNER And T.表名=R.TABLE_NAME And R.GRANTEE='" & strRole & _
                "' And R.PRIVILEGE in ('SELECT','INSERT','UPDATE','DELETE') " & _
                "GROUP BY T.系统,T.表名,T.说明 " & _
                "Having Count(R.PRIVILEGE) = 4"
        rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
        Do Until rsTemp.EOF
            Set lst = lvwModule.ListItems.Add(, , rsTemp("表名"))
            lst.SubItems(1) = rsTemp("系统")
            lst.SubItems(2) = IIf(IsNull(rsTemp("说明")), "", rsTemp("说明"))
            rsTemp.MoveNext
        Loop
    ElseIf cmbSystem.Text = "取数函数" Then
        '显示该角色能访问的基础表
        gstrSQL = "select S.名称||'（'||S.编号||'）' as 系统,S.所有者,F.函数名,F.中文名,F.说明 " & _
                  " from zlSystems S,zlFunctions F,USER_TAB_PRIVS R " & _
                  " where  F.系统=S.编号 And S.所有者=R.OWNER And Upper(F.函数名)=R.TABLE_NAME And R.GRANTEE='" & strRole & "' And R.PRIVILEGE ='EXECUTE'"
        rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
        Do Until rsTemp.EOF
            Set lst = lvwModule.ListItems.Add(, , rsTemp("函数名"))
            lst.SubItems(1) = rsTemp("中文名")
            lst.SubItems(2) = rsTemp("系统")
            lst.SubItems(3) = IIf(IsNull(rsTemp("说明")), "", rsTemp("说明"))
            rsTemp.MoveNext
        Loop
    ElseIf cmbSystem.Text = "基础工具" Then
        '显示该角色能访问的基础工具
        gstrSQL = "select P.序号,P.标题,P.说明,R.功能 from " & _
                "zlRoleGrant R,zlPrograms P " & _
                "where R.系统 is Null And P.序号=R.序号 And R.角色='" & strRole & _
                "'  And P.系统 is Null And P.序号<100 And P.部件 is Null " & _
                " Order By P.序号"
        rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
        On Error Resume Next
        Do Until rsTemp.EOF
            Set lst = lvwModule.ListItems.Add(, "C" & rsTemp("序号"), rsTemp("序号"))
            If err <> 0 Then
                err.Clear
                If rsTemp("功能") <> "基本" Then
                    Set lst = lvwModule.ListItems("C" & rsTemp("序号"))
                    lst.SubItems(3) = IIf(lst.SubItems(3) = "", "", lst.SubItems(3) & ",") & rsTemp("功能")
                End If
            Else
                lst.SubItems(1) = rsTemp("标题")
                lst.SubItems(2) = IIf(IsNull(rsTemp("说明")), "", rsTemp("说明"))
                If rsTemp("功能") <> "基本" Then
                    lst.SubItems(3) = rsTemp("功能")
                End If
            End If
            rsTemp.MoveNext
        Loop
    Else
        '显示该角色能访问的模块
        gstrSQL = "select P.序号,P.标题,P.说明,R.功能 from " & _
                "zlRoleGrant R,zlPrograms P " & _
                "where nvl(R.系统,0)=nvl(P.系统,0) And P.序号=R.序号 And P.序号>=100 And R.角色='" & strRole & "'  And " & _
                IIf(cmbSystem.Text = "自定义报表", " P.系统 is Null ", " P.系统=" & cmbSystem.ItemData(cmbSystem.ListIndex)) & _
                " Order By P.序号"
        rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
        On Error Resume Next
        Do Until rsTemp.EOF
            Set lst = lvwModule.ListItems.Add(, "C" & rsTemp("序号"), rsTemp("序号"))
            If err <> 0 Then
                err.Clear
                If rsTemp("功能") <> "基本" Then
                    Set lst = lvwModule.ListItems("C" & rsTemp("序号"))
                    lst.SubItems(3) = IIf(lst.SubItems(3) = "", "", lst.SubItems(3) & ",") & rsTemp("功能")
                End If
            Else
                lst.SubItems(1) = rsTemp("标题")
                lst.SubItems(2) = IIf(IsNull(rsTemp("说明")), "", rsTemp("说明"))
                If rsTemp("功能") <> "基本" Then
                    lst.SubItems(3) = rsTemp("功能")
                End If
            End If
            rsTemp.MoveNext
        Loop
    End If
    
    LockWindowUpdate 0
End Sub

Private Sub SetEnable()
'设置各个按钮的Enable属性
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
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = True
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'供主窗口调用，实现具体的打印工作
'如果没有可打印的，就留下一个空的接口

'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    Dim objPrint As zlPrintLvw
    
    If lvwRole.SelectedItem Is Nothing Then Exit Sub
    Set objPrint = New zlPrintLvw
    objPrint.Title.Text = "角色权限"
    Set objPrint.Body.objData = lvwModule
    objPrint.UnderAppItems.Add "角色：" & lvwRole.SelectedItem.Text
    objPrint.UnderAppItems.Add "授权系统：" & cmbSystem.Text
    objPrint.BelowAppItems.Add "打印时间：" & Format(CurrentDate, "yyyy年MM月dd日")
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
        '设置图标
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
     '删除组
     Call DeleteRoleGroups
     Call LoadMenus
End Sub

Private Sub mnuPopuModify_Click()
    '更名
    If tvwGroups.SelectedItem.Key <> "Root" Then
        Call tvwGroups.StartLabelEdit
    End If
End Sub

Private Sub mnuPopuRoleAdd_Click()
    Call cmdAdd_Click
End Sub

Private Sub mnuPopuRoleDelete_Click()
    '删除脚色
    If cmdAdd.Enabled = False Then Exit Sub
    Call DeleteRole
End Sub

Private Sub mnuPopuRoleMoveGroups_Click(Index As Integer)
    Dim str组名 As String
    If mnuPopuRoleMoveGroups(Index).Tag = "" Then Exit Sub
    str组名 = UCase(Mid(mnuPopuRoleMoveGroups(Index).Tag, 2))
    
    If str组名 = UCase("oot") Or str组名 = "所有角色" Then
        If str组名 = Mid(tvwGroups.SelectedItem.Key, 2) Then Exit Sub
        If MsgBox("你真要将角色“" & lvwRole.SelectedItem.Text & "...” 移出该组吗?", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
        str组名 = ""
    Else
        If str组名 = Mid(tvwGroups.SelectedItem.Key, 2) Then Exit Sub
        If MsgBox("你真要将角色“" & lvwRole.SelectedItem.Text & "...” 移动到组“" & str组名 & "”里吗?", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
    End If
    If MoveToGroups(str组名) = False Then Exit Sub
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
    '过程:zlTools.b_Rolegroupmgr.Rolegroup_Delete(
    '    组名_Old_In In ZlRolegroups.组名%Type,
    '    组名_New_In In ZlRolegroups.组名%Type
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
    '功能:删除组
    '编制:刘兴宏
    '日期:2007/06/15
    '---------------------------------------------------------------------------------------------------------
    Dim strRoleGroupName As String
    Dim intIndex As Integer
     
    If tvwGroups.SelectedItem Is Nothing Then Exit Sub
    If tvwGroups.SelectedItem.Key = "Root" Then Exit Sub
    
    strRoleGroupName = tvwGroups.SelectedItem.Text
    intIndex = tvwGroups.SelectedItem.Index
    
    If MsgBox("真的要删除“" & strRoleGroupName & "”的角色组吗？", vbDefaultButton2 Or vbQuestion Or vbYesNo, gstrSysName) = vbNo Then Exit Sub
    
    err = 0: On Error GoTo ErrHand:
    '过程:zlTools.b_Rolegroupmgr.Rolegroup_Delete(组名_In In ZlRolegroups.组名%Type)
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
    Dim str组名 As String, str角色 As String, intIndex As Integer
    Dim lstItem As ListItem
    Dim strKeys As String
    Dim arrVar As Variant
    Dim i As Long
     err = 0: On Error GoTo ErrHand:
     
    If Source Is lvwRole And Not tvwGroups.DropHighlight Is Nothing Then
        intIndex = -1
        str组名 = Mid(tvwGroups.DropHighlight.Key, 2)
        Set tvwGroups.DropHighlight = Nothing
        tvwGroups.DropHighlight = tvwGroups.SelectedItem

        If str组名 = "oot" Then
            If str组名 = Mid(tvwGroups.SelectedItem.Key, 2) Then Exit Sub
            If MsgBox("你真要将角色“" & Source.SelectedItem.Text & "...” 移出该组吗?", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
            str组名 = ""
        Else
            If str组名 = Mid(tvwGroups.SelectedItem.Key, 2) Then Exit Sub
            If MsgBox("你真要将角色“" & Source.SelectedItem.Text & "...” 移动到组“" & str组名 & "”里吗?", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) <> vbYes Then Exit Sub
        End If

        gcnOracle.BeginTrans
        strKeys = ""
        For Each lstItem In lvwRole.ListItems
            If lstItem.Selected Then
                If intIndex < 0 Then
                    intIndex = lstItem.Index
                End If
                str角色 = lstItem.Key
                strKeys = strKeys & "'" & lstItem.Key

                If MoveToRoleGroup(str组名, str角色) = False Then
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
Private Function MoveToGroups(ByVal str组名 As String) As Boolean
    '-------------------------------------------------------------------------------------------------------------------
    '功能:将指定角色移动到组中
    '参数:str组-移到组的组名
    '移动成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/06/19
    '-------------------------------------------------------------------------------------------------------------------
    Dim intIndex As Integer
    Dim strKeys  As String
    Dim lstItem As ListItem
    Dim str角色 As String
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
            str角色 = lstItem.Key
            strKeys = strKeys & "'" & lstItem.Key
            If MoveToRoleGroup(str组名, str角色) = False Then
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
Private Function MoveToRoleGroup(ByVal str组 As String, str角色 As String) As Boolean
    '-------------------------------------------------------------------------------------------------------------------
    '功能:将指定的角色移到组中
    '参数:str组-移到组的组名
    '     str角色-指定的角色
    '移动成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/06/19
    '-------------------------------------------------------------------------------------------------------------------
    err = 0: On Error GoTo ErrHand:
    
    '刘兴宏:20070615加入
    '过程参数:zlTools.b_Rolegroupmgr.RoletoRolegroup
    '        组名_In In ZlRolegroups.组名%Type,
    '        角色_In In ZlRolegroups.角色%Type := Null
    gstrSQL = "zlTools.b_Rolegroupmgr.RoletoRolegroup("
    gstrSQL = gstrSQL & IIf(str组 = "", "Null", "'" & UCase(str组) & "'") & ","
    gstrSQL = gstrSQL & "'" & UCase(str角色) & "')"
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
        '更名
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
    '获取相应的角色权限
    '---------------------------------------------------------------------------------------------------------
    Dim strKey As String
    Call SetEnable
    strKey = IIf(Node.Key = "Root", "", Mid(Node.Key, 2))
    Call FillRoleData(strKey)
    Call FillModule
    Call SetEnable
End Sub

Private Function FillRoleData(ByVal str组名 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------
    '功能:根据组名,获取角色信息,并填充到lvw中
    '参数:str组名:<>""指定组名,=""表示所有角色
    '返回:加载成功,返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------
    Dim rsGroups As New ADODB.Recordset
    Dim objItem As ListItem
    Dim blnGroups As Boolean '是否要过滤组
    Dim strFiler  As String
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    strFiler = UCase(Trim(txtSearch.Text))
     
    '显示可以进行授权的角色
    If gblnDBA = True Then
        gstrSQL = _
            " Select User as Grantee,'YES' as Admin_Option,Substr(A.Role,4) as RoleName," & _
            " A.Role,zlSpellCode(Substr(A.Role,4)) as 简码  " & _
            " From DBA_Roles A,zlTools.zlRoleGroups B " & _
            " Where Upper(A.Role) Like 'ZL_%' And A.Role=B.角色" & _
            IIf(str组名 = "", "(+)", " And B.组名='" & str组名 & "' And B.角色 is Not Null") & _
            " Order by A.Role"
    Else
        '所有者和SYSTEM中的ZL角色
        gstrSQL = _
            " Select UserName As Grantee,A.Admin_Option,Substr(A.Granted_Role ,4) as RoleName," & _
            " A.Granted_Role as Role,zlSpellCode(Substr(A.Granted_Role ,4)) as 简码 " & _
            " From User_Role_Privs A,zlRoleGroups B " & _
            " Where A.Granted_Role Like 'ZL_%' And A.Granted_Role=B.角色" & _
            IIf(str组名 = "", "(+)", " And B.组名='" & str组名 & "' And B.角色 is Not Null") & _
            " Union ALL" & _
            " Select A.Grantee,A.Admin_Option,Substr(A.Granted_Role ,4) as RoleName," & _
            " A.Granted_Role as Role,zlSpellCode(Substr(A.Granted_Role ,4)) as 简码 " & _
            " From DBA_Role_Privs A,zlRoleGroups B " & _
            " Where A.Granted_Role Like 'ZL_%' And A.Granted_Role=B.角色" & _
            IIf(str组名 = "", "(+)", " And B.组名='" & str组名 & "' And B.角色 is Not Null") & _
            " And A.Grantee='SYSTEM' And Not Exists(" & _
                " Select 1 From DBA_Role_Privs X Where X.Granted_Role=A.Granted_Role And Grantee='" & gstrUserName & "')" & _
            " Order by Role"
    End If
    Call OpenRecordset(mrsRole, gstrSQL, Me.Caption)
    If strFiler <> "" Then
        mrsRole.Filter = "RoleName Like '" & strFiler & "%' or 简码 Like '" & strFiler & "%'"
    End If
    
    blnGroups = False
    If chkOnlyShowNOGroups.Enabled And chkOnlyShowNOGroups.value = 1 Then
        '只显示未分配的组
        strSQL = "SELECT 角色 FROM zlRoleGroups where 角色 is Not Null "
        OpenRecordset rsGroups, strSQL, Me.Caption
        blnGroups = True
    End If
    
    lvwRole.ListItems.Clear
    Do Until mrsRole.EOF
        Set objItem = Nothing
        If blnGroups Then
            rsGroups.Filter = "角色='" & UCase(Nvl(mrsRole!Role)) & "'"
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
    MsgBox "错误：" & err.Number & vbCrLf & vbCrLf & err.Description, vbInformation, gstrSysName
End Function

Private Function SearchRole(ByVal strFilter As String) As Boolean
    '----------------------------------------------------------------------------------------------------------------------------
    '功能:过虑出相应的角色
    '参数:strFilter-过滤串
    '返回:成功,返回ture,否则返回False
    '----------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsGroups As New ADODB.Recordset
    Dim blnGroups As Boolean '是否要过滤组
    SearchRole = True
    If mrsRole Is Nothing Then Exit Function
    If mrsRole.State <> 1 Then Exit Function
    If mrsRole.RecordCount = 0 Then Exit Function
    
    blnGroups = False
    If chkOnlyShowNOGroups.Enabled And chkOnlyShowNOGroups.value = 1 Then
        '只显示未分配的组
        strSQL = "SELECT 角色 FROM zlRoleGroups where 角色 is Not Null "
        OpenRecordset rsGroups, strSQL, Me.Caption
        blnGroups = True
    End If
    
    strFilter = UCase(strFilter)
    SearchRole = False
    If strFilter = "" Then
    Else
        mrsRole.Filter = "RoleName Like '" & strFilter & "%' or 简码 Like '" & strFilter & "%'"
    End If
    lvwRole.ListItems.Clear
    Do Until mrsRole.EOF
        If blnGroups Then
            rsGroups.Filter = "角色='" & UCase(Nvl(mrsRole!Role)) & "'"
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
    '功能:加载角色组
    '编制:刘兴宏
    '日期:2007/06/15
    '--------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    Dim objNode As Node
    gstrSQL = "Select distinct 组名 From zlRoleGroups"
    Call OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    With tvwGroups
        .Nodes.Clear
        Set objNode = tvwGroups.Nodes.Add(, 4, "Root", "所有角色", 1, 1)
        objNode.Selected = True
        objNode.Expanded = True
        objNode.Sorted = True
        Call LoadMenu("所有角色", "Root")
        Do While Not rsTemp.EOF
            Set objNode = tvwGroups.Nodes.Add("Root", 4, "K" & Nvl(rsTemp!组名), Nvl(rsTemp!组名), 1, 1)
            objNode.Tag = Nvl(rsTemp!组名)
            objNode.Sorted = True
            Call LoadMenu(Nvl(rsTemp!组名), Nvl(rsTemp!组名))
            rsTemp.MoveNext
        Loop
    End With
    Call tvwGroups_NodeClick(Me.tvwGroups.SelectedItem)
End Sub
Private Sub LoadMenu(ByVal strTittle As String, ByVal strTag As String)
    '功能:加载菜单
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
    '功能:拆卸菜单
        Dim varMenu As Variant
        Dim intCount As Integer
        Set varMenu = mnuPopuRoleMoveGroups
        mnuPopuRoleMoveGroups(0).Visible = True
        mnuPopuRoleMove.Visible = True
        For intCount = 1 To mnuPopuRoleMoveGroups.UBound
            Unload varMenu(intCount)
        Next
        
End Sub

