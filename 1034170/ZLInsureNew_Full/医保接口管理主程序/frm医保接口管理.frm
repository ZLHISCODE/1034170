VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frm医保接口管理 
   Caption         =   "医保接口管理"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8715
   Icon            =   "frm医保接口管理.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   8715
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.ImageList imgProp 
      Left            =   3690
      Top             =   3780
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
            Picture         =   "frm医保接口管理.frx":1CFA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgModul 
      Left            =   3690
      Top             =   1170
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
            Picture         =   "frm医保接口管理.frx":4D7C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgInterface 
      Left            =   30
      Top             =   750
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保接口管理.frx":5BCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保接口管理.frx":6E50
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwInterface 
      Height          =   4635
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   8176
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgInterface"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "序号"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "名称"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "医保部件名称"
         Object.Width           =   2646
      EndProperty
   End
   Begin MSComctlLib.ImageList imgTbrBlack 
      Left            =   3660
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保接口管理.frx":80D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保接口管理.frx":82EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保接口管理.frx":8506
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保接口管理.frx":8958
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保接口管理.frx":8B72
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgTbrColor 
      Left            =   3090
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保接口管理.frx":8D8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保接口管理.frx":90DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保接口管理.frx":9430
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保接口管理.frx":9882
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保接口管理.frx":9A9C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrTool 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   1244
      BandCount       =   1
      FixedOrder      =   -1  'True
      _CBWidth        =   8715
      _CBHeight       =   705
      _Version        =   "6.7.8988"
      Child1          =   "tbrTool"
      MinHeight1      =   645
      Width1          =   1575
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrTool 
         Height          =   645
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   8595
         _ExtentX        =   15161
         _ExtentY        =   1138
         ButtonWidth     =   820
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgTbrBlack"
         HotImageList    =   "imgTbrColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "安装"
               Key             =   "Install"
               Object.ToolTipText     =   "安装医保接口部件"
               Object.Tag             =   "安装"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "卸载"
               Key             =   "Uninstall"
               Object.ToolTipText     =   "卸载医保接口部件"
               Object.Tag             =   "卸载"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "split0"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "启用"
               Key             =   "Start"
               Object.ToolTipText     =   "启用"
               Object.Tag             =   "启用"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "split1"
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Exit"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   7
      Top             =   5340
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   635
      SimpleText      =   $"frm医保接口管理.frx":9CB6
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm医保接口管理.frx":9CFD
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView lvw调试模块 
      Height          =   2385
      Left            =   3630
      TabIndex        =   6
      Top             =   930
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   4207
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgModul"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "模块"
         Object.Width           =   2646
      EndProperty
   End
   Begin MSComctlLib.ListView lvw属性栏 
      Height          =   1815
      Left            =   3630
      TabIndex        =   4
      Top             =   3540
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   3201
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imgProp"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "说明"
         Object.Width           =   8819
      EndProperty
   End
   Begin VB.Label lblNote 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      Caption         =   "接口调试"
      ForeColor       =   &H8000000E&
      Height          =   180
      Index           =   0
      Left            =   3630
      TabIndex        =   5
      Top             =   750
      Width           =   5040
   End
   Begin VB.Label lblNote 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      Caption         =   "属性栏"
      ForeColor       =   &H8000000E&
      Height          =   180
      Index           =   1
      Left            =   3630
      MousePointer    =   7  'Size N S
      TabIndex        =   3
      Top             =   3360
      Width           =   5040
   End
   Begin VB.Image imgSplit 
      Height          =   4605
      Left            =   3540
      MousePointer    =   9  'Size W E
      Top             =   750
      Width           =   75
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuInterface 
      Caption         =   "接口(&I)"
      Begin VB.Menu mnuInterfaceInstall 
         Caption         =   "安装(&I)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuInterfaceUninstall 
         Caption         =   "卸载(&U)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuInterfaceSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInterfaceStart 
         Caption         =   "启用(&S)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "帮助主题(&H)"
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&WEB上的中联"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "发送反馈(&M)..."
         End
      End
      Begin VB.Menu mnuHelpLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "frm医保接口管理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintInsure As Integer
Private mstrInsureUser As String
Private mstrInsureTablespace As String
Private mstrInsureName As String
Private mstrDemo As String
Private mstrComponent As String
Private mstrPath As String
Private mstrSQL As String
Private mstrUser As String              '用户名
Private mstrServer As String            '主机串
Private mblnDBA As Boolean              'DBA用户或所有者

Private mobjTest() As Object
Private mstrTest() As String
Private mobjConfigure As Object

Private mblnMove As Boolean
Private Type 坐标
    x As Double
    y As Double
End Type
Private Type_Scale As 坐标
'只有所有者或DBA才有权进行医保接口管理的安装、卸载
'普通用户只允许进行接口的调试

Private Sub Form_Load()
    mstrUser = GetSetting("ZLSOFT", "注册信息\登陆信息", "USER", "")
    mstrServer = GetSetting("ZLSOFT", "注册信息\登陆信息", "SERVER", "")
    
    mblnDBA = IsDBA()
    mnuInterfaceInstall.Visible = mblnDBA
    mnuInterfaceUninstall.Visible = mblnDBA
    mnuInterfaceSplit1.Visible = mblnDBA
    tbrTool.Buttons("Install").Visible = mblnDBA
    tbrTool.Buttons("Uninstall").Visible = mblnDBA
    tbrTool.Buttons("split0").Visible = mblnDBA
    
    Call LoadInterface
End Sub

Private Sub LoadInterface()
    Dim lvwItem As ListItem
    Dim rsInsure As New ADODB.Recordset
    On Error GoTo ErrHand
    
    '装入已注册医保接口的数据
    mstrSQL = " Select A.序号,A.名称,B.部件 As 医保部件,Nvl(B.启用,0) 启用" & _
              " From 保险类别 A,zlInsureComponents B" & _
              " Where A.序号=B.险类" & _
              " Order By A.序号"
    Call zlDatabase.OpenRecordset(rsInsure, mstrSQL, "装入已注册的医保接口")
    
    With rsInsure
        lvwInterface.ListItems.Clear
        Do While Not .EOF
            Set lvwItem = lvwInterface.ListItems.Add(, "K_" & !序号, !序号, , !启用 + 1)
            lvwItem.SubItems(1) = Nvl(!名称)
            lvwItem.SubItems(2) = Nvl(!医保部件)
            lvwItem.Tag = !启用 + 1
            .MoveNext
        Loop
    End With
    
    '如果有，则调用点击事件，显示详细信息，否则将相关控件及按钮设置为不允许选择状态
    If Me.lvwInterface.ListItems.Count <> 0 Then
        Call lvwInterface_ItemClick(lvwInterface.ListItems(1))
    Else
        Call SetEnabled(False)
    End If
    Me.stbThis.Panels(2).Text = ""
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    
    With Me.lvwInterface
        .Width = imgSplit.Left
        .Height = Me.ScaleHeight - stbThis.Height - .Top
    End With
    With lblNote(0)
        .Top = lvwInterface.Top
        .Left = imgSplit.Left + imgSplit.Width
        .Width = Me.ScaleWidth - .Left
    End With
    With Me.lvw调试模块
        .Top = lblNote(0).Top + lblNote(0).Height
        .Left = lblNote(0).Left
        .Width = lblNote(0).Width
        .Height = lblNote(1).Top - .Top
    End With
    With lblNote(1)
        .Left = lblNote(0).Left
        .Width = lblNote(0).Width
    End With
    With Me.lvw属性栏
        .Top = lblNote(1).Top + lblNote(1).Height
        .Left = lblNote(0).Left
        .Width = lblNote(0).Width
        .Height = Me.ScaleHeight - stbThis.Height - .Top
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim intIndex As Integer, intCount As Integer
    On Error Resume Next
    
    '关闭所有对象
    intCount = UBound(mobjTest)
    If Err <> 0 Then intCount = -1
    
    For intIndex = 0 To intCount
        Call mobjTest(intIndex).CloseWindows
        Set mobjTest(intIndex) = Nothing
    Next
    Call CloseWindows
End Sub

Private Sub imgSplit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnMove = (Button = 1)
    If Not mblnMove Then Exit Sub
    
    Type_Scale.x = x
    Type_Scale.y = y
End Sub

Private Sub imgSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim dblLeft As Double
    If Not mblnMove Then Exit Sub
    
    dblLeft = imgSplit.Left + x - Type_Scale.x
    If dblLeft < 1000 Or dblLeft > Me.ScaleWidth - 1000 Then Exit Sub
    
    With imgSplit
        .Move .Left + x - Type_Scale.x
    End With
    Call Form_Resize
End Sub

Private Sub imgSplit_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnMove = False
End Sub

Private Sub lblNote_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index = 0 Then Exit Sub
    mblnMove = (Button = 1)
    If Not mblnMove Then Exit Sub
    
    Type_Scale.x = x
    Type_Scale.y = y
End Sub

Private Sub lblNote_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim dblTop As Double
    If Index = 0 Then Exit Sub
    If Not mblnMove Then Exit Sub
    
    dblTop = lblNote(Index).Top + y - Type_Scale.y
    If dblTop - lblNote(0).Top < 1500 Then Exit Sub
    If Me.ScaleHeight - stbThis.Height - (lblNote(Index).Top + y - Type_Scale.y) - lblNote(1).Height < 1500 Then Exit Sub
    
    With lblNote(Index)
        .Move .Left, .Top + y - Type_Scale.y
    End With
    Call Form_Resize
End Sub

Private Sub lblNote_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    mblnMove = False
End Sub

Private Sub lvwInterface_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lvwInterface
        .Sorted = False
        .SortKey = ColumnHeader.Index - 1
        .SortOrder = IIf(.SortOrder = lvwDescending, lvwAscending, lvwDescending)
        .Sorted = True
    End With
End Sub

Private Sub lvwInterface_DblClick()
    Call lvwInterface_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub lvwInterface_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim lng险类 As Long
    Dim IntDO As Integer, intCount As Integer
    Dim arrItem
    Dim lvwItem As ListItem
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    If lvwInterface.ListItems.Count = 0 Then Exit Sub
    If lvwInterface.SelectedItem Is Nothing Then Exit Sub
    
    '显示该注册医保接口的详细信息
    lng险类 = Mid(Item.Key, 3)
    mnuInterfaceStart.Enabled = (Item.Tag = 1)
    tbrTool.Buttons("Start").Enabled = mnuInterfaceStart.Enabled
    
    '>>取支持的模块
    Me.lvw调试模块.ListItems.Clear
    mstrSQL = "Select A.序号,A.标题,A.说明" & _
        " From zlPrograms A,zlInsureModuls B" & _
        " Where A.序号=B.序号 And B.险类=" & lng险类 & _
        " Order by A.序号"
    Call zlDatabase.OpenRecordset(rsTemp, mstrSQL, "医保接口管理")
    With rsTemp
        '医保接口基础模块，也可用于调试
        Do While Not .EOF
            Set lvwItem = Me.lvw调试模块.ListItems.Add(, "K_" & !序号, Nvl(!标题), , 1)
            lvwItem.Tag = "zl9Insure"
            .MoveNext
        Loop
        '医保接口相关模块
        Set lvwItem = Me.lvw调试模块.ListItems.Add(, "K_1111", "门诊挂号", , 1)
        lvwItem.Tag = "zl9RegEvent"
        Set lvwItem = Me.lvw调试模块.ListItems.Add(, "K_1121", "门诊收费", , 1)
        lvwItem.Tag = "zl9OutExse"
        Set lvwItem = Me.lvw调试模块.ListItems.Add(, "K_1131", "入院登记", , 1)
        lvwItem.Tag = "zl9Inpatient"
        Set lvwItem = Me.lvw调试模块.ListItems.Add(, "K_1132", "入出院管理", , 1)
        lvwItem.Tag = "zl9Inpatient"
        Set lvwItem = Me.lvw调试模块.ListItems.Add(, "K_1133", "住院记帐", , 1)
        lvwItem.Tag = "zl9InExse"
        Set lvwItem = Me.lvw调试模块.ListItems.Add(, "K_1137", "住院结算", , 1)
        lvwItem.Tag = "zl9InExse"
    End With
    
    '>>取支持库说明
    lvw属性栏.ListItems.Clear
    mstrSQL = "Select 文件名,说明 From zlInsureBase Where 险类=" & lng险类
    Call zlDatabase.OpenRecordset(rsTemp, mstrSQL, "医保接口管理")
    With rsTemp
        Do While Not .EOF
            Call Me.lvw属性栏.ListItems.Add(, "K_" & .AbsolutePosition, Nvl(!文件名) & "," & Nvl(!说明, "无"), , 1)
            .MoveNext
        Loop
    End With
    
    '>>取支持业务说明
    mstrSQL = "Select 业务,描述 From zlInsureOperation Where 险类=" & lng险类
    Call zlDatabase.OpenRecordset(rsTemp, mstrSQL, "医保接口管理")
    
    With rsTemp
        If .RecordCount <> 0 Then
            For IntDO = 1 To 4
                .Filter = "业务=" & IntDO
                Do While Not .EOF
                    If Trim(Nvl(!描述)) <> "" Then
                        arrItem = Split(!描述, "|")
                        For intCount = 0 To UBound(arrItem)
                            Call lvw属性栏.ListItems.Add(, "K_" & lvw属性栏.ListItems.Count + 1, arrItem(intCount), , 1)
                        Next
                    End If
                    .MoveNext
                Loop
            Next
        End If
    End With
    
    '设置当前选择的医保
    SaveSetting "ZLSOFT", "公共全局", "是否支持医保", "Yes"
    SaveSetting "ZLSOFT", "公共全局", "医保类别", lng险类
    
    '设置各控件及按钮的状态
    Call SetEnabled(True)
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub lvwInterface_KeyDown(KeyCode As Integer, Shift As Integer)
    If lvwInterface.ListItems.Count = 0 Then Exit Sub
    If lvwInterface.SelectedItem Is Nothing Then Exit Sub
    
    If MsgBox("你确定要启用该医保接口吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysname) = vbNo Then Exit Sub
    Call mnuInterfaceStart_Click
End Sub

Private Sub lvw调试模块_DblClick()
    Call lvw调试模块_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub lvw调试模块_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intIndex As Integer
    If lvw调试模块.ListItems.Count = 0 Then Exit Sub
    If lvw调试模块.SelectedItem Is Nothing Then Exit Sub
    
    Call FindObject
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuHelpTitle_Click()
    '
End Sub

Private Sub mnuHelpWebHome_Click()
    '中联主页
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    '发送反馈
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub mnuInterfaceInstall_Click()
    If Not InitConfigure Then Exit Sub
    Me.stbThis.Panels(2).Text = "正在安装医保接口..."
    If Not mobjConfigure.I_Install(mstrServer) Then
        Me.stbThis.Panels(2).Text = ""
        Exit Sub
    End If
    Call LoadInterface
    MsgBox "医保接口部件安装成功！", vbInformation, gstrSysname
End Sub

Private Sub mnuInterfaceStart_Click()
    Dim lng险类 As Long
    
    If lvwInterface.ListItems.Count = 0 Then Exit Sub
    If lvwInterface.SelectedItem Is Nothing Then Exit Sub
    If lvwInterface.SelectedItem.Tag = 2 Then Exit Sub
    
    lng险类 = Mid(lvwInterface.SelectedItem.Key, 3)
    mstrSQL = "ZL_ZLINSURECOMPONENTS_START(" & lng险类 & ")"
    gcnOracle.Execute mstrSQL, , adCmdStoredProc
    
    Call LoadInterface
End Sub

Private Sub mnuInterfaceUninstall_Click()
    Dim intInsure As Integer
    If lvwInterface.SelectedItem Is Nothing Then Exit Sub
    intInsure = lvwInterface.SelectedItem
    
    If Not InitConfigure Then Exit Sub
    Me.stbThis.Panels(2).Text = "正在卸载医保接口..."
    If Not mobjConfigure.I_UnInstall(intInsure) Then
        Me.stbThis.Panels(2).Text = ""
        Exit Sub
    End If
    
    Me.stbThis.Panels(2).Text = ""
    lvwInterface.ListItems.Remove lvwInterface.SelectedItem.Key
    '如果有，则调用点击事件，显示详细信息，否则将相关控件及按钮设置为不允许选择状态
    If Me.lvwInterface.ListItems.Count <> 0 Then
        Call lvwInterface_ItemClick(lvwInterface.ListItems(1))
    Else
        Call SetEnabled(False)
        lvw调试模块.ListItems.Clear
        lvw属性栏.ListItems.Clear
    End If
    
    MsgBox "医保接口部件卸载成功！", vbInformation, gstrSysname
End Sub

Private Sub tbrTool_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Install"
        Call mnuInterfaceInstall_Click
    Case "Uninstall"
        Call mnuInterfaceUninstall_Click
    Case "Start"
        Call mnuInterfaceStart_Click
    Case "Exit"
        Call mnuFileExit_Click
    Case "Help"
        Call mnuHelpTitle_Click
    End Select
End Sub

Private Sub SetEnabled(ByVal BlnState As Boolean)
    mnuInterfaceUninstall.Enabled = BlnState
    tbrTool.Buttons("Uninstall").Enabled = BlnState
End Sub

Private Function RegistFile() As Boolean
    Const strRegist As String = "Regist.txt"
    '检查注册文件的合法性
    RegistFile = True
End Function

Private Sub FindObject()
    '查找是否已创建指定的部件，如果未创建则创建
    Dim strClass As String
    Dim strObject As String
    Dim lngModul As Long
    Dim blnExist As Boolean
    Dim objTest As Object
    Dim intIndex As Integer, intCount As Integer
    Const lngSys As Long = 100
    On Error Resume Next
    
    strObject = UCase(lvw调试模块.SelectedItem.Tag)
    lngModul = Val(Mid(lvw调试模块.SelectedItem.Key, 3))
    strClass = strObject & ".cls" & Mid(strObject, 4)
    
    Err = 0
    intCount = UBound(mobjTest)
    If Err <> 0 Then intCount = -1
    
    For intIndex = 0 To intCount
        If mstrTest(intIndex) = UCase(strObject) Then
            blnExist = True
            Exit For
        End If
    Next
    
    '创建部件
    If blnExist = False Then
        If Not objTest Is Nothing Then
            Call objTest.CloseWindows
            Set objTest = Nothing
        End If
        
        Err = 0
        Set objTest = CreateObject(strClass)
        If Err <> 0 Then
            MsgBox "无法创建该部件，请确认是否已安装！", vbInformation
            Exit Sub
        End If
        
        ReDim Preserve mobjTest(intIndex) As Object
        ReDim Preserve mstrTest(intIndex) As String
        Set mobjTest(intIndex) = objTest
        mstrTest(intIndex) = UCase(strObject)
    End If
    
    On Error GoTo ErrHand
    Call mobjTest(intIndex).CodeMan(lngSys, lngModul, gcnOracle, Nothing, "ZLHIS")
    
    Me.WindowState = 1
    Exit Sub
ErrHand:
    MsgBox Err.Description, vbInformation, gstrSysname
End Sub

Private Function IsDBA() As Boolean
    Dim rsTemp As New ADODB.Recordset
    '判断传入的用户是不是所有者或DBA用户
    
    mstrSQL = "SELECT 1 FROM DUAL " & _
            " WHERE EXISTS(SELECT 1 FROM ZLSYSTEMS WHERE 所有者='" & UCase(mstrUser) & "')"
    Call zlDatabase.OpenRecordset(rsTemp, mstrSQL, "判断该用户是不是所有者或DBA用户")
    IsDBA = (rsTemp.RecordCount <> 0)
End Function

Private Function InitConfigure() As Boolean
    If mobjConfigure Is Nothing Then
        On Error Resume Next
        Err = 0
        Set mobjConfigure = CreateObject("zl9I_Configure.clsI_Configure")
        If Err <> 0 Then
            MsgBox "主要组件丢失，无法完成医保接口部件的安装或卸载！", vbInformation, gstrSysname
            Exit Function
        End If
        Call mobjConfigure.InitOracle(gcnOracle)
    End If
    
    InitConfigure = True
End Function
