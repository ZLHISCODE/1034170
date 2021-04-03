VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMgrUserGrant 
   Caption         =   "管理工具授权"
   ClientHeight    =   7395
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7575
   Icon            =   "frmMgrUserGrant.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   7575
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdMove 
      Height          =   495
      Index           =   1
      Left            =   3390
      Picture         =   "frmMgrUserGrant.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton cmdMove 
      Height          =   495
      Index           =   0
      Left            =   3960
      Picture         =   "frmMgrUserGrant.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3360
      Width           =   375
   End
   Begin MSComctlLib.TreeView tvwGranted 
      Height          =   5895
      Left            =   4440
      TabIndex        =   9
      Top             =   960
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   10398
      _Version        =   393217
      Sorted          =   -1  'True
      Style           =   7
      Checkboxes      =   -1  'True
      ImageList       =   "Img16"
      Appearance      =   1
   End
   Begin MSComctlLib.TreeView tvwNoGrant 
      Height          =   5895
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   10398
      _Version        =   393217
      Sorted          =   -1  'True
      Style           =   7
      Checkboxes      =   -1  'True
      ImageList       =   "Img16"
      Appearance      =   1
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "查找用户(&F)"
      Height          =   350
      Left            =   6120
      TabIndex        =   2
      Top             =   65
      Width           =   1215
   End
   Begin VB.TextBox txtFind 
      Height          =   300
      Left            =   4200
      TabIndex        =   1
      Top             =   90
      Width           =   1815
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4980
      TabIndex        =   3
      Top             =   6975
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6255
      TabIndex        =   4
      Top             =   6975
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   -360
      TabIndex        =   0
      Top             =   525
      Width           =   10110
   End
   Begin MSComctlLib.ImageList Img16 
      Left            =   3600
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   39
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":1A5E
            Key             =   "自动提醒"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":82C0
            Key             =   "系统装卸管理"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":EB22
            Key             =   "数据转移"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":15384
            Key             =   "用户注册管理"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":1BBE6
            Key             =   "系统升迁管理"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":22448
            Key             =   "系统参数管理"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":28CAA
            Key             =   "运行日志管理"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":2F50C
            Key             =   "错误日志管理"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":35D6E
            Key             =   "系统运行选项"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":3C5D0
            Key             =   "对象检查修复"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":42E32
            Key             =   "数据导出"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":49694
            Key             =   "站点文件收集"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":4FEF6
            Key             =   "编译无效对象"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":56758
            Key             =   "后台作业管理"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":5CFBA
            Key             =   "数据导入"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":6381C
            Key             =   "数据调入"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":6A07E
            Key             =   "数据清除"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":708E0
            Key             =   "数据调出"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":77142
            Key             =   "运行状态监控"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":7D9A4
            Key             =   "置换安装脚本"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":84206
            Key             =   "站点部件升级"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":8AA68
            Key             =   "报表管理"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":912CA
            Key             =   "函数管理"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":97B2C
            Key             =   "用户授权管理"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":9E38E
            Key             =   "角色授权管理"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":A4BF0
            Key             =   "菜单重组规划"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":AB452
            Key             =   "站点运行控制"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":B1CB4
            Key             =   "权限管理"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":B258E
            Key             =   "装卸管理"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":B2E68
            Key             =   "数据管理"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":B3742
            Key             =   "运行管理"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":B3CDC
            Key             =   "专项工具"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":B45B6
            Key             =   "DBA工具"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":BAE18
            Key             =   "空间管理"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":C167A
            Key             =   "SQL性能"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":C7EDC
            Key             =   "会话解锁"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":CE73E
            Key             =   "外键索引"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":D4FA0
            Key             =   "SQL跟踪"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMgrUserGrant.frx":DB802
            Key             =   "数据库性能"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "对“梁唐彬”进行授权处理。"
      Height          =   180
      Left            =   960
      TabIndex        =   7
      Top             =   150
      UseMnemonic     =   0   'False
      Width           =   3090
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgOne 
      Height          =   480
      Left            =   300
      Picture         =   "frmMgrUserGrant.frx":E2064
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblModul 
      AutoSize        =   -1  'True
      Caption         =   "可授权功能(&A)"
      Height          =   180
      Left            =   210
      TabIndex        =   6
      Top             =   660
      Width           =   1170
   End
   Begin VB.Label lblGranted 
      AutoSize        =   -1  'True
      Caption         =   "已授权功能(&G)"
      Height          =   180
      Left            =   4335
      TabIndex        =   5
      Top             =   660
      Width           =   1170
   End
End
Attribute VB_Name = "frmMgrUserGrant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrUser As String
Private mstrProg As String
Private mstrAccount As String '为空表示新用户授权
Private mblnOK As Boolean

Public Function GrantToProg(ByVal strAccount As String, ByVal strUser As String, ByVal strProg As String) As Boolean
    mstrUser = strUser
    mstrAccount = strAccount
    mstrProg = strProg
    mblnOK = False
    Me.Show 1
    GrantToProg = mblnOK
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Call FindPersonnel
End Sub

Private Sub MoveProg(objMoveIn As TreeView, objMoveOut As TreeView)
    Dim i As Long, y As Long
    Dim strDel As String, Node As Node
    
    For i = objMoveOut.Nodes.Count To 1 Step -1
        err = 0
        On Error Resume Next
        If objMoveOut.Nodes(i).Checked And Not objMoveOut.Nodes(i).Parent Is Nothing Then
            If err = 0 Then
                err = 0
                If objMoveIn.Nodes(objMoveOut.Nodes(i).Parent.Key).Key <> "" Then
                    If err <> 0 Then
                        '新增父项
                        Set Node = objMoveIn.Nodes.Add(, , objMoveOut.Nodes(i).Parent.Key, objMoveOut.Nodes(i).Parent.Text, objMoveOut.Nodes(i).Parent.Image, objMoveOut.Nodes(i).Parent.SelectedImage)
                        Node.Expanded = objMoveOut.Nodes(i).Parent.Expanded
                        Node.Checked = objMoveOut.Nodes(i).Parent.Checked
                        Node.ForeColor = objMoveOut.Nodes(i).Parent.ForeColor
                    End If
                     '新增子项
                    Set Node = objMoveIn.Nodes.Add(objMoveOut.Nodes(i).Parent.Key, tvwChild, objMoveOut.Nodes(i).Key, objMoveOut.Nodes(i).Text, objMoveOut.Nodes(i).Image, objMoveOut.Nodes(i).SelectedImage)
                    Node.Expanded = objMoveOut.Nodes(i).Expanded
                    Node.Checked = objMoveOut.Nodes(i).Checked
                    Node.ForeColor = objMoveOut.Nodes(i).ForeColor
                    '删除子项
                    If objMoveOut.Nodes(i).Parent.Children = 1 Then
                        objMoveOut.Nodes.Remove objMoveOut.Nodes(i).Parent.Index
                    Else
                        objMoveOut.Nodes.Remove i
                    End If
                    
                End If
                On Error GoTo 0
            End If
        End If
    Next
End Sub

Private Sub cmdMove_Click(Index As Integer)
    If Index = 0 Then
        Call MoveProg(tvwGranted, tvwNoGrant)
    ElseIf Index = 1 Then
        Call MoveProg(tvwNoGrant, tvwGranted)
    End If
End Sub

Private Sub cmdOK_Click()
'功能：授权
    Dim i As Integer, strProg As String
    Dim StrJiami() As Byte
    Dim strPwText As String
    Dim rsTemp As New ADODB.Recordset
    
    If mstrAccount = "" Then
        MsgBox "请先查找需要授权的用户。", vbInformation, Me.Caption
        If txtFind.Visible Then txtFind.SetFocus
        Exit Sub
    End If
    For i = 1 To tvwGranted.Nodes.Count
        If Not tvwGranted.Nodes(i).Parent Is Nothing Then
            strProg = strProg & "," & Trim(Mid(tvwGranted.Nodes(i).Key, 2))
        End If
    Next
    strProg = Mid(strProg, 2)
    '功能加密
    If strProg <> "" Then
        Call DES_Encode(StrConv(strProg, vbFromUnicode), StrJiami, gobjRegister.zlRegInfo("单位名称", False, 0))
        strPwText = FuncByteTo16Code(StrJiami)
    End If
    On Error GoTo errHandle
    gstrSQL = "Select 1 From zlMgrGrant Where 用户名='" & mstrAccount & "'"
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    If rsTemp.RecordCount > 0 Then
        If strPwText = "" Then
            gstrSQL = "Delete zlMgrGrant Where 用户名='" & mstrAccount & "'"
        Else
            gstrSQL = "Update zlMgrGrant Set 功能='" & strPwText & "' Where 用户名='" & mstrAccount & "'"
        End If
    Else
        gstrSQL = "Insert into zlMgrGrant(用户名,功能) values('" & mstrAccount & "','" & strPwText & "')"
    End If
    gcnOracle.Execute gstrSQL
    '更新管理员账户信息
    rsTemp.Close
    gstrSQL = "Select 1 From zlRegInfo where 项目='管理员'"
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
     If rsTemp.RecordCount > 0 Then
        gstrSQL = "Update zlRegInfo Set 内容='" & gstrUserName & "' Where 项目='管理员'"
    Else
        gstrSQL = "Insert into zlRegInfo(项目,内容) values('管理员','" & gstrUserName & "')"
    End If
    gcnOracle.Execute gstrSQL
    '验证码
    strPwText = ""
    ReDim Preserve StrJiami(0)
    If gstrPassword <> "" Then
        Call DES_Encode(StrConv(gstrPassword, vbFromUnicode), StrJiami, gobjRegister.zlRegInfo("单位名称", False, 0))
        strPwText = FuncByteTo16Code(StrJiami)
    End If
    rsTemp.Close
    gstrSQL = "Select 1 From zlRegInfo where 项目='验证码'"
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
     If rsTemp.RecordCount > 0 Then
        gstrSQL = "Update zlRegInfo Set 内容='" & strPwText & "' Where 项目='验证码'"
    Else
        gstrSQL = "Insert into zlRegInfo(项目,内容) values('验证码','" & strPwText & "')"




    End If
    gcnOracle.Execute gstrSQL
    mblnOK = True
    Unload Me
    Exit Sub
errHandle:
    MsgBox "[" & err.Number & "]" & err.Description, vbExclamation, Me.Caption
End Sub
Private Sub Form_Load()
    If mstrAccount = "" Then
        lblNote.Caption = "请先输入用户名、人员姓名或简码。"
        txtFind.Visible = True
        cmdFind.Visible = True
    Else
        lblNote.Caption = "正在对""" & mstrUser & """进行管理工具授权。"
        txtFind.Visible = False
        cmdFind.Visible = False
    End If

    Call FillProg
End Sub

Private Sub FillProg()
'功能：填充功能
    Dim rsTemp As New ADODB.Recordset
    Dim strProg As String, Node As Node
    Dim i As Long
    
    On Error GoTo errHandle
    '显示该用户具有的角色
    gstrSQL = "Select /*+Rule */ a.编号,a.标题,A.上级,Column_Value as 权限" & vbNewLine & _
            "From zlSvrTools A, (Select Column_Value From Table(Cast(f_Str2list('" & mstrProg & "') As Zltools.t_Strlist))) C" & vbNewLine & _
            "Where  a.编号 = c.Column_Value(+) And A.标题<>'管理工具授权'" & vbNewLine & _
            "Order By a.编号"

    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    Do Until rsTemp.EOF
        With IIf(rsTemp!权限 & "" = "", tvwNoGrant, tvwGranted)
            '上级两边都加
            If IsNull(rsTemp("上级")) Then
                Set Node = tvwNoGrant.Nodes.Add(, , "D" & rsTemp("编号"), "【" & rsTemp("编号") & "】" & rsTemp("标题"))
                tvwNoGrant.Nodes("D" & rsTemp("编号")).Sorted = True
                tvwNoGrant.Nodes("D" & rsTemp("编号")).Expanded = True
                tvwNoGrant.Nodes("D" & rsTemp("编号")).ForeColor = &HFF0000
                On Error Resume Next
                Node.Image = Img16.ListImages.Item(rsTemp!标题 & "").Index
                err.Clear: On Error GoTo errHandle
                Set Node = tvwGranted.Nodes.Add(, , "D" & rsTemp("编号"), "【" & rsTemp("编号") & "】" & rsTemp("标题"))
                tvwGranted.Nodes("D" & rsTemp("编号")).Sorted = True
                tvwGranted.Nodes("D" & rsTemp("编号")).Expanded = True
                tvwGranted.Nodes("D" & rsTemp("编号")).ForeColor = &HFF0000
                On Error Resume Next
                Node.Image = Img16.ListImages.Item(rsTemp!标题 & "").Index
                err.Clear: On Error GoTo errHandle
            Else
                Set Node = .Nodes.Add("D" & rsTemp("上级"), tvwChild, "C" & rsTemp("编号"), rsTemp("标题"))
                .Nodes("C" & rsTemp("编号")).Sorted = True
                Node.Checked = False
                On Error Resume Next
                Node.Image = Img16.ListImages.Item(rsTemp!标题 & "").Index
                err.Clear: On Error GoTo errHandle
            End If
        End With
        rsTemp.MoveNext
    Loop
    '删除没有子项的分类
    For i = tvwNoGrant.Nodes.Count To 1 Step -1
        If tvwNoGrant.Nodes(i).Children = 0 And tvwNoGrant.Nodes(i).Parent Is Nothing Then
            tvwNoGrant.Nodes.Remove i
        End If
    Next
    For i = tvwGranted.Nodes.Count To 1 Step -1
        If tvwGranted.Nodes(i).Children = 0 And tvwGranted.Nodes(i).Parent Is Nothing Then
            tvwGranted.Nodes.Remove i
        End If
    Next
    Exit Sub
errHandle:
    MsgBox "[" & err.Number & "]" & err.Description, vbExclamation, Me.Caption
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    fraLine.Width = Me.Width
    cmdOK.Move Me.Width - cmdCancel.Width - cmdOK.Width - 400, Me.Height - cmdOK.Height - 650
    cmdCancel.Move cmdOK.Left + cmdOK.Width + 100, cmdOK.Top
    tvwNoGrant.Width = Me.Width \ 2 - 885
    tvwGranted.Width = Me.Width \ 2 - 885
    tvwNoGrant.Height = cmdOK.Top - tvwNoGrant.Top - 100
    tvwGranted.Height = tvwNoGrant.Height
    tvwGranted.Left = tvwNoGrant.Left + tvwNoGrant.Width + 1185
    cmdMove(1).Left = tvwNoGrant.Left + tvwNoGrant.Width + 150
    cmdMove(0).Left = cmdMove(1).Left + cmdMove(1).Width + 150
    lblGranted.Left = tvwGranted.Left
End Sub

Private Sub tvwGranted_NodeCheck(ByVal Node As MSComctlLib.Node)
    Call NodeCheckMode(Node, tvwGranted)
End Sub

Private Sub tvwNoGrant_NodeCheck(ByVal Node As MSComctlLib.Node)
     Call NodeCheckMode(Node, tvwNoGrant)
End Sub

Private Sub NodeCheckMode(ByRef Node As MSComctlLib.Node, ByRef objtvwThis As TreeView)
'功能：让树表选中父节点，自动选中所有子节点，选中所有子节点，父节点也选中
    Dim i As Long
    Dim blnIsNothing As Boolean
    
    LockWindowUpdate objtvwThis.hwnd
    If Node.Parent Is Nothing Then
        For i = Node.Index + 1 To objtvwThis.Nodes.Count
            If Not objtvwThis.Nodes(i).Parent Is Nothing And objtvwThis.Nodes(i).ForeColor <> &H80000010 Then
                If objtvwThis.Nodes(i).Parent.Key = Node.Key Then
                    objtvwThis.Nodes(i).Checked = Node.Checked
                End If
            End If
        Next
    Else
        For i = Node.Parent.Index + 1 To objtvwThis.Nodes.Count
            If Not objtvwThis.Nodes(i).Parent Is Nothing And objtvwThis.Nodes(i).ForeColor <> &H80000010 Then
                If objtvwThis.Nodes(i).Parent.Key = Node.Parent.Key Then
                    If Not objtvwThis.Nodes(i).Checked = Node.Checked Then blnIsNothing = True
                End If
            End If
        Next
        If blnIsNothing Then
            Node.Parent.Checked = False
        Else
            Node.Parent.Checked = Node.Checked
        End If
    End If
    LockWindowUpdate 0
End Sub

Private Sub txtFind_GotFocus()
    txtFind.SelStart = 0: txtFind.SelLength = Len(txtFind.Text)
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call FindPersonnel
    End If
End Sub

Private Sub FindPersonnel()
'功能：查找人员
    Dim rsTemp As New Recordset
    Dim objPoint As POINTAPI
    
    If txtFind.Text = "" Then Exit Sub
    gstrSQL = "Select b.用户名, c.姓名, c.简码, d.名称 As 部门名称" & vbNewLine & _
            "From  Zlmgrgrant A,上机人员表 B, 人员表 C, 部门表 D, 部门人员 E" & vbNewLine & _
            "Where a.用户名(+) = b.用户名 And b.人员id = c.Id And c.Id = e.人员id And d.Id = e.部门id And A.用户名 is null And e.缺省 = 1 And B.用户名 <> '" & gstrUserName & "'" & _
            " And(b.用户名 like '" & UCase(Trim(txtFind.Text)) & "%' Or c.姓名 Like '" & UCase(Trim(txtFind.Text)) & "%' Or c.简码 Like '" & UCase(Trim(txtFind.Text)) & "%' Or c.编号=' & UCase(Trim(txtFind.Text)) & ')" & _
            " Order By c.姓名"
    Set rsTemp = New ADODB.Recordset
    OpenRecordset rsTemp, gstrSQL, Me.Caption
    If rsTemp.RecordCount = 0 Then
        MsgBox "您查找的用户不存在，或是已经拥有了权限，请检查。", vbInformation, Me.Caption
        If txtFind.Visible Then txtFind.SetFocus: Call txtFind_GotFocus
        Exit Sub
    End If
    Call ClientToScreen(txtFind.hwnd, objPoint)
    
    If frmSelectList.ShowSelect(Nothing, rsTemp, "用户名,900,0,1;姓名,900,0,1;简码,650,0,0;部门名称,1500,0,1", objPoint.X * 15 - 30, objPoint.y * 15 + cmdFind.Height - 30, txtFind.Width + cmdFind.Width + 1300, 3000, "", "查找人员", , , True) = False Then
        If txtFind.Visible Then txtFind.SetFocus: Call txtFind_GotFocus
        rsTemp.Filter = 0
        Exit Sub
    Else
        txtFind.Text = rsTemp!姓名 & ""
        mstrAccount = rsTemp!用户名 & ""
        mstrUser = rsTemp!姓名 & ""
    End If
End Sub
