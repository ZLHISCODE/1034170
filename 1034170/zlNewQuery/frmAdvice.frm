VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmAdvice 
   Caption         =   "播放广告设置"
   ClientHeight    =   6165
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9330
   Icon            =   "frmAdvice.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvw 
      Height          =   2520
      Left            =   45
      TabIndex        =   5
      Top             =   855
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   4445
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils32"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "名称"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "序号"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "类型"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "宽度(像素)"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "高度(像素)"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "修改日期"
         Object.Width           =   3175
      EndProperty
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   1080
      Top             =   4470
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvice.frx":164A
            Key             =   "pic"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvice.frx":249C
            Key             =   "ico"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvice.frx":32EE
            Key             =   "swf"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   345
      Top             =   4470
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvice.frx":4140
            Key             =   "pic"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvice.frx":4F92
            Key             =   "swf"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvice.frx":5DE4
            Key             =   "ico"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMenuHot 
      Left            =   6780
      Top             =   405
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvice.frx":6C36
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvice.frx":6E56
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvice.frx":7076
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvice.frx":7296
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvice.frx":74B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvice.frx":7A10
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvice.frx":7F6A
            Key             =   "View"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvice.frx":8186
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvice.frx":83A6
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   6210
      Top             =   405
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvice.frx":85C6
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvice.frx":87E6
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvice.frx":8A06
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvice.frx":8C26
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvice.frx":8E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvice.frx":93A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvice.frx":98FA
            Key             =   "View"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvice.frx":9B16
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdvice.frx":9D36
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3555
      ScaleHeight     =   315
      ScaleWidth      =   6435
      TabIndex        =   3
      Top             =   810
      Width           =   6435
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   180
         Left            =   105
         TabIndex        =   4
         Top             =   60
         Width           =   105
      End
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   9330
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   720
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   720
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   9210
         _ExtentX        =   16245
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsMenuHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   13
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "预览"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "打印"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "增加"
               Key             =   "增加"
               Object.ToolTipText     =   "增加"
               Object.Tag             =   "增加"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "删除"
               Object.ToolTipText     =   "删除"
               Object.Tag             =   "删除"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_2"
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "上移"
               Key             =   "上移"
               Object.ToolTipText     =   "播放顺序上移"
               Object.Tag             =   "上移"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "下移"
               Key             =   "下移"
               Object.ToolTipText     =   "播放顺序下移"
               Object.Tag             =   "下移"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_3"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查看"
               Key             =   "查看"
               Object.ToolTipText     =   "播放广告查看方式"
               Object.Tag             =   "查看"
               ImageIndex      =   7
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "大图标"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "小图标"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "列表"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "详细资料"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_4"
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "帮助"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   5805
      Width           =   9330
      _ExtentX        =   16457
      _ExtentY        =   635
      SimpleText      =   $"frmAdvice.frx":9F56
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmAdvice.frx":9F9D
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11377
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
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
   Begin zl9NewQuery.ctlPicture picBack 
      Height          =   2310
      Left            =   3555
      TabIndex        =   6
      Top             =   1275
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   4075
   End
   Begin VB.Image picX 
      Height          =   1530
      Left            =   3150
      MousePointer    =   9  'Size W E
      Top             =   945
      Width           =   210
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFilePreView 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileUpdatePage 
         Caption         =   "更新查询页面(&U)"
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditNew 
         Caption         =   "增加(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "删除(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditUp 
         Caption         =   "播放顺序上移(&U)"
      End
      Begin VB.Menu mnuEditDown 
         Caption         =   "播放顺序下移(&D)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "标准按钮(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "文本标签(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuviewsplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "大图标(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "小图标(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "列表(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "详细资料(&D)"
         Checked         =   -1  'True
         Index           =   3
      End
      Begin VB.Menu mnuViewSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewPos 
         Caption         =   "广告位置(&P)"
      End
      Begin VB.Menu mnuViewSplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "刷新(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "帮助主题(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&Web上的中联"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "中联论坛(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "发送反馈(&K)..."
         End
      End
      Begin VB.Menu mnuHelpSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "frmAdvice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnFist As Boolean
Private mvarSvrPicRange As String           '保存增加图片的范围
Private mvarSvrPicType As String            '保存增加图片的类型

Private Sub Form_Activate()
    If mblnFist = False Then Exit Sub
    mblnFist = False
    
    Call Form_Resize
    DoEvents
    
    Call mnuViewRefresh_Click
End Sub

Private Sub Form_Load()
    mblnFist = True
    
    RestoreWinState Me, App.ProductName
    
    Call mnuViewIcon_Click(lvw.View)
    
    Call ReadRegister
    Call ModulePrivs
    
    picX.MousePointer = 9
    picX.Width = 45
    mvarSvrPicRange = ""
    mvarSvrPicType = ""
    
End Sub

Private Sub Form_Resize()
    '根据窗体状态,调整窗体中各控件的显示位置
    
    Dim sglCbrH As Single
    Dim sglStbH As Single
    
    On Error Resume Next
    sglCbrH = IIf(cbrThis.Visible, cbrThis.Height, 0)
    sglStbH = IIf(stbThis.Visible, stbThis.Height, 0)
    
    Call ResizeControl(lvw, 0, sglCbrH, picX.Left, Me.ScaleHeight - sglStbH - sglCbrH)
    Call ResizeControl(picX, picX.Left, lvw.Top, picX.Width, lvw.Height)
    Call ResizeControl(picTitle, picX.Left + picX.Width, lvw.Top + 15, Me.ScaleWidth - picX.Left - picX.Width, picTitle.Height)
    Call ResizeControl(picBack, picTitle.Left, picTitle.Top + picTitle.Height + 15, picTitle.Width, Me.ScaleHeight - picTitle.Top - picTitle.Height - sglStbH - 15)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call WriteRegister
    SaveWinState Me, App.ProductName
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)
    lblTitle.Caption = "图片名称:" & Item.Text & " 宽度:" & Item.SubItems(1) & " 高度:" & Item.SubItems(2)
    picBack.Tag = Item.Tag
    Call ShowPicture
    Call AdjustEnabled
End Sub

Private Sub lvw_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And mnuEdit.Visible Then Me.PopupMenu mnuEdit, 2
End Sub

Private Sub mnuEditDown_Click()
    '将当前的项目向上移一行，同时更新数据库

    Dim svrAry(6) As String
    Dim intPre As Long
    Dim strSQL(3) As String
    
    If lvw.SelectedItem Is Nothing Then Exit Sub
    
    intPre = lvw.SelectedItem.Index + 1
    
    If intPre < lvw.ListItems.Count + 1 Then
        strSQL(0) = "zl_咨询广告序列_update(" & Val(Mid(lvw.SelectedItem.Key, 2)) & ",-1)"
        strSQL(1) = "zl_咨询广告序列_update(" & Val(Mid(lvw.ListItems(intPre).Key, 2)) & "," & Val(Mid(lvw.SelectedItem.Key, 2)) & ")"
        strSQL(2) = "zl_咨询广告序列_update(-1," & Val(Mid(lvw.ListItems(intPre).Key, 2)) & ")"
        
        On Error GoTo errHand
        gcnOracle.BeginTrans
                
        Call zlDatabase.ExecuteProcedure(strSQL(0), Me.Caption)
        Call zlDatabase.ExecuteProcedure(strSQL(1), Me.Caption)
        Call zlDatabase.ExecuteProcedure(strSQL(2), Me.Caption)
        
        gcnOracle.CommitTrans
        On Error GoTo 0
        
        svrAry(0) = lvw.ListItems(intPre).Text
        svrAry(1) = lvw.ListItems(intPre).SubItems(2)
        svrAry(2) = lvw.ListItems(intPre).SubItems(3)
        svrAry(3) = lvw.ListItems(intPre).SubItems(4)
        svrAry(5) = lvw.ListItems(intPre).Tag
        
        lvw.ListItems(intPre).Text = lvw.SelectedItem.Text
        lvw.ListItems(intPre).SubItems(2) = lvw.SelectedItem.SubItems(2)
        lvw.ListItems(intPre).SubItems(3) = lvw.SelectedItem.SubItems(3)
        lvw.ListItems(intPre).SubItems(4) = lvw.SelectedItem.SubItems(4)
        lvw.ListItems(intPre).Tag = lvw.SelectedItem.Tag
        
        lvw.SelectedItem.Text = svrAry(0)
        lvw.SelectedItem.SubItems(2) = svrAry(1)
        lvw.SelectedItem.SubItems(3) = svrAry(2)
        lvw.SelectedItem.SubItems(4) = svrAry(3)
        lvw.SelectedItem.Tag = svrAry(5)
        
        lvw.ListItems(intPre).Selected = True
        Call AdjustEnabled
    End If
    Exit Sub
errHand:
    
    gcnOracle.RollbackTrans
    
    If ErrCenter() = -1 Then Resume
    
End Sub

Private Sub mnuEditNew_Click()
    Dim lngKey As Long
    Dim lngNO As Long
    
    If frmPicSelect.OpenPictureBox(Me, "广告增加", "2;0;1;4;9", lngKey, mvarSvrPicRange, mvarSvrPicType) Then
        '更新图片显示
        
'        lngNo = NextValue("咨询广告序列", "序号")
'        gstrSQL = "zl_咨询广告序列_insert(" & lngNo & "," & lngKey & ")"
'
'        On Error GoTo errHand
'        gcnOracle.BeginTrans
'        gcnOracle.Execute gstrSQL, , adCmdStoredProc
'        gcnOracle.CommitTrans
'        On Error GoTo 0
'        Call mnuViewRefresh_Click
    End If
    Exit Sub
errHand:
    
    If ErrCenter() = -1 Then Resume Next
    
End Sub

Private Sub mnuEditDelete_Click()
    Dim vIndex As Long
    
    If lvw.SelectedItem Is Nothing Then Exit Sub
    
    If MsgBox("您确认要移除广告[" & lvw.SelectedItem.Text & "]吗？", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Exit Sub
    
    On Error GoTo errHand

    'gcnOracle.Execute "zl_咨询广告序列_delete(" & Val(Mid(lvw.SelectedItem.Key, 2)) & ")", , adCmdStoredProc
    gstrSQL = "zl_咨询广告序列_delete(" & Val(Mid(lvw.SelectedItem.Key, 2)) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    vIndex = lvw.SelectedItem.Index
    lvw.ListItems.Remove lvw.SelectedItem.Index
    Call NextLvwPos(lvw, vIndex)
    
    If Not (lvw.SelectedItem Is Nothing) Then
        Call lvw_ItemClick(lvw.SelectedItem)
        Call AdjustOrder(lvw, 1)
    Else
        picBack.Cls
    End If
    Call AdjustEnabled
    Call LoadStatus
    
    Exit Sub
errHand:
    If ErrCenter() = -1 Then Resume
    
End Sub

Private Sub mnuEditUp_Click()
    '将当前的项目向上移一行，同时更新数据库
    
    Dim svrAry(6) As String
    Dim intPre As Long
    Dim strSQL(3) As String
    
    intPre = lvw.SelectedItem.Index - 1
    
    If intPre > 0 Then
    
        strSQL(0) = "zl_咨询广告序列_update(" & Val(Mid(lvw.SelectedItem.Key, 2)) & ",-1)"
        strSQL(1) = "zl_咨询广告序列_update(" & Val(Mid(lvw.ListItems(intPre).Key, 2)) & "," & Val(Mid(lvw.SelectedItem.Key, 2)) & ")"
        strSQL(2) = "zl_咨询广告序列_update(-1," & Val(Mid(lvw.ListItems(intPre).Key, 2)) & ")"
        
        On Error GoTo errHand
        gcnOracle.BeginTrans
        'gcnOracle.Execute strSQL(0), , adCmdStoredProc
        'gcnOracle.Execute strSQL(1), , adCmdStoredProc
        'gcnOracle.Execute strSQL(2), , adCmdStoredProc
        
        Call zlDatabase.ExecuteProcedure(strSQL(0), Me.Caption)
        Call zlDatabase.ExecuteProcedure(strSQL(1), Me.Caption)
        Call zlDatabase.ExecuteProcedure(strSQL(2), Me.Caption)
        
        gcnOracle.CommitTrans
        On Error GoTo 0
        
        svrAry(0) = lvw.ListItems(intPre).Text
        svrAry(1) = lvw.ListItems(intPre).SubItems(2)
        svrAry(2) = lvw.ListItems(intPre).SubItems(3)
        svrAry(3) = lvw.ListItems(intPre).SubItems(4)
        svrAry(5) = lvw.ListItems(intPre).Tag
        
        lvw.ListItems(intPre).Text = lvw.SelectedItem.Text
        lvw.ListItems(intPre).SubItems(2) = lvw.SelectedItem.SubItems(2)
        lvw.ListItems(intPre).SubItems(3) = lvw.SelectedItem.SubItems(3)
        lvw.ListItems(intPre).SubItems(4) = lvw.SelectedItem.SubItems(4)
        lvw.ListItems(intPre).Tag = lvw.SelectedItem.Tag
        
        lvw.SelectedItem.Text = svrAry(0)
        lvw.SelectedItem.SubItems(2) = svrAry(1)
        lvw.SelectedItem.SubItems(3) = svrAry(2)
        lvw.SelectedItem.SubItems(4) = svrAry(3)
        lvw.SelectedItem.Tag = svrAry(5)
        
        lvw.ListItems(intPre).Selected = True
        Call AdjustEnabled
    End If
    Exit Sub
errHand:
    
    gcnOracle.RollbackTrans
    
    If ErrCenter() = -1 Then Resume
    
End Sub

Private Sub mnuFileExcel_Click()
    Call PrintObject(3)
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub



Private Sub mnuFilePreView_Click()
    Call PrintObject(2)
End Sub

Private Sub mnuFilePrint_Click()
    Call PrintObject(1)
End Sub

Private Sub mnuFilePrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub mnuFileUpdatePage_Click()
    Call gfrmMain.FrameDefault.RefreshPage
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTopic_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
    Dim i As Long
    
    For i = 0 To 3
        mnuViewIcon(i).Checked = False
    Next
    mnuViewIcon(Index).Checked = True
    
    lvw.View = Index
End Sub

Private Sub mnuViewPos_Click()
    Call frmPosSample.ShowPageSample("广告图片")
End Sub

Private Sub mnuViewRefresh_Click()
    Dim svrKey As String
    
    svrKey = SaveLvwItem(lvw)
    Call LoadAdviceList
    Call RestoreLvwItem(lvw, svrKey)
    Call AdjustEnabled
    Call LoadStatus
    
    If lvw.SelectedItem Is Nothing Then Exit Sub
    DoEvents
    
    Call lvw_ItemClick(lvw.SelectedItem)
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    Call Form_Resize
End Sub


Private Sub mnuViewToolText_Click()
    Dim i As Long
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To tbrThis.Buttons.Count
        tbrThis.Buttons(i).Caption = IIf(mnuViewToolText.Checked, tbrThis.Buttons(i).Tag, "")
    Next
    cbrThis.Bands(1).MinHeight = tbrThis.Height
    Call Form_Resize
    
End Sub

Private Sub ShowPicture()
    '显示图形包括图片及FLASH电影
    gstrSQL = "select 图形,宽度,高度,类型 from 咨询图片元素 where 序号=[1]"
    
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(picBack.Tag))
    If gRs.BOF = False Then
        Call picBack.ShowPictureByField(gRs!图形, gRs!宽度 * Screen.TwipsPerPixelX, gRs!高度 * Screen.TwipsPerPixelY, IIf(IsNull(gRs!类型), 0, gRs!类型))
    End If
End Sub

Private Sub picX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    picX.Left = picX.Left + X
    If picX.Left < 1500 Then picX.Left = 1500
    If Me.Width - picX.Left - picX.Width < 1500 Then picX.Left = Me.Width - picX.Width - 1500
    
    Form_Resize
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "预览"
        Call mnuFilePreView_Click
    Case "打印"
        Call mnuFilePrint_Click
    Case "增加"
        Call mnuEditNew_Click
    Case "删除"
        Call mnuEditDelete_Click
    Case "上移"
        Call mnuEditUp_Click
    Case "下移"
        Call mnuEditDown_Click
    Case "查看"
        If lvw.View < 3 Then
            Call mnuViewIcon_Click(lvw.View + 1)
        Else
            Call mnuViewIcon_Click(0)
        End If
    Case "帮助"
        Call mnuHelpTopic_Click
    Case "退出"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then Me.PopupMenu Me.mnuViewTool, 2
End Sub

Private Sub LoadAdviceList()
    '加载已设置好的广告
    
    Dim Itmx As ListItem
    Dim lngNO As Long
    
    On Error GoTo errHand
    
    lvw.ListItems.Clear
    picBack.Cls
    
    gstrSQL = "select A.图片序号,A.序号,B.名称,B.类型,B.宽度,B.高度,B.修改日期 from 咨询广告序列 A,咨询图片元素 B where A.图片序号=B.序号 order by A.序号"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If gRs.BOF = False Then
        While Not gRs.EOF
            lngNO = lngNO + 1
            Set Itmx = lvw.ListItems.Add(, "K" & gRs!序号, IIf(IsNull(gRs!名称), "", gRs!名称), "pic", "pic")
            Itmx.Tag = IIf(IsNull(gRs!图片序号), 0, gRs!图片序号)
            Itmx.SubItems(1) = lngNO
            Select Case IIf(IsNull(gRs!类型), 0, gRs!类型)
            Case 0
                Itmx.SubItems(2) = "图片"
                Itmx.Icon = "pic"
            Case 1
                Itmx.SubItems(2) = "图标"
                Itmx.Icon = "ico"
            Case 2
                Itmx.SubItems(2) = "Flash"
                Itmx.Icon = "swf"
            End Select
            Itmx.SmallIcon = Itmx.Icon
                        
            Itmx.SubItems(3) = IIf(IsNull(gRs!宽度), "", gRs!宽度)
            Itmx.SubItems(4) = IIf(IsNull(gRs!高度), "", gRs!高度)
            Itmx.SubItems(5) = IIf(IsNull(gRs!修改日期), "", gRs!修改日期)
            gRs.MoveNext
        Wend
    End If
    Exit Sub
errHand:
    If ErrCenter() = -1 Then Resume
    Call SaveErrLog
End Sub

Private Sub tbrThis_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Text
    Case "大图标"
        Call mnuViewIcon_Click(0)
    Case "小图标"
        Call mnuViewIcon_Click(1)
    Case "列表"
        Call mnuViewIcon_Click(2)
    Case "详细资料"
        Call mnuViewIcon_Click(3)
    End Select
End Sub

Private Sub PrintObject(ByVal intMode As Byte)
    '---------------------------------------------------
    '功能：    根据屏幕组织表上附加项目，打印预览
    '参数：
    '     intMode: 2表示预览 1打印 3输出到EXCEL
    '返回：
    '---------------------------------------------------
    
    Dim objPrint As New zlPrintLvw
    Dim objRow As New zlTabAppRow

    If lvw.SelectedItem Is Nothing Then Exit Sub

    If UserInfo.姓名 = "" Then Call GetUserInfo

    objPrint.Title = "播放广告设置清单"
    objPrint.BelowAppItems.Add "打印人:" & UserInfo.姓名
    objPrint.BelowAppItems.Add "打印时间:" & Format(zlDatabase.Currentdate, "YYYY年MM月DD日")
    objPrint.Footer = "第[页码]页;;"

    Set objPrint.Body.objData = lvw

    If intMode = 1 Then
        Select Case zlPrintAsk(objPrint)
        Case 1
            zlPrintOrViewLvw objPrint, 1
        Case 2
            zlPrintOrViewLvw objPrint, 2
        Case 3
            zlPrintOrViewLvw objPrint, 3
        End Select
    Else
        zlPrintOrViewLvw objPrint, intMode
    End If

End Sub

Private Sub ModulePrivs()
    '根据模块权限,处理功能项的隐藏或显示;权限有:增删改
    
'    mnuEdit.Visible = True
'    If InStr(gstrPrivs, "增删改") = 0 Then
'        mnuEdit.Visible = False
'
'        tbrThis.Buttons("增加").Visible = False
'        tbrThis.Buttons("删除").Visible = False
'        tbrThis.Buttons("Split_2").Visible = False
'
'        tbrThis.Buttons("上移").Visible = False
'        tbrThis.Buttons("下移").Visible = False
'        tbrThis.Buttons("Split_3").Visible = False
'    End If
    
End Sub

Private Sub AdjustEnabled()
    '调整功能菜单或按钮的可用状态
    
    mnuFilePreView.Enabled = True
    mnuFilePrint.Enabled = True
    mnuFileExcel.Enabled = True
    mnuEditDelete.Enabled = True
    mnuEditNew.Enabled = True
    mnuEditUp.Enabled = True
    mnuEditDown.Enabled = True
    
    If lvw.ListItems.Count = 0 Then
        mnuFilePreView.Enabled = False
        mnuFilePrint.Enabled = False
        mnuFileExcel.Enabled = False
    End If
    
    If lvw.SelectedItem Is Nothing Then
        mnuEditDelete.Enabled = False
        mnuEditDown.Enabled = False
        mnuEditUp.Enabled = False
    Else
        If lvw.SelectedItem.Index - 1 <= 0 Then mnuEditUp.Enabled = False
        If lvw.SelectedItem.Index + 1 > lvw.ListItems.Count Then mnuEditDown.Enabled = False
    End If
                
    tbrThis.Buttons("预览").Enabled = mnuFilePreView.Enabled
    tbrThis.Buttons("打印").Enabled = mnuFilePrint.Enabled
    tbrThis.Buttons("增加").Enabled = mnuEditNew.Enabled
    tbrThis.Buttons("删除").Enabled = mnuEditDelete.Enabled
    tbrThis.Buttons("下移").Enabled = mnuEditDown.Enabled
    tbrThis.Buttons("上移").Enabled = mnuEditUp.Enabled
End Sub

Private Sub ReadRegister()
    '读取注册表信息
    
    If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then Exit Sub
    
    picX.Left = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\界面\" & Me.Name, "picX位置", 2385)
    
End Sub

Private Sub WriteRegister()
    '将信息写回注册表
    
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\界面\" & Me.Name, "picX位置", picX.Left
End Sub

Private Sub LoadStatus()
    '刷新状态栏显示
    
    If lvw.ListItems.Count > 0 Then
        stbThis.Panels(2).Text = "当前共有" & lvw.ListItems.Count & "个要播放的广告！"
    Else
        stbThis.Panels(2).Text = "当前没有要播放的广告！"
    End If
End Sub

Public Sub AddPicture(ByVal lngKey As Long)
    Dim lngNO As Long
    Dim Itmx As ListItem
    
    On Error GoTo errHand1
    
    lngNO = NextValue("咨询广告序列", "序号")
    gstrSQL = "zl_咨询广告序列_insert(" & lngNO & "," & lngKey & ")"
    
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    
    On Error GoTo 0
    On Error GoTo Handerr
    gstrSQL = "select A.图片序号,A.序号,B.名称,B.类型,B.宽度,B.高度,B.修改日期 from 咨询广告序列 A,咨询图片元素 B where A.图片序号=B.序号 and B.序号=[1]"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    If gRs.BOF = False Then
        Set Itmx = lvw.ListItems.Add(, "K" & lngNO, IIf(IsNull(gRs!名称), "", gRs!名称), "pic", "pic")
        Itmx.Tag = IIf(IsNull(gRs!图片序号), 0, gRs!图片序号)
        Itmx.SubItems(1) = lvw.ListItems.Count
        Select Case IIf(IsNull(gRs!类型), 0, gRs!类型)
        Case 0
            Itmx.SubItems(2) = "图片"
            Itmx.Icon = "pic"
        Case 1
            Itmx.SubItems(2) = "图标"
            Itmx.Icon = "ico"
        Case 2
            Itmx.SubItems(2) = "Flash"
            Itmx.Icon = "swf"
        End Select
        Itmx.SmallIcon = Itmx.Icon
                    
        Itmx.SubItems(3) = IIf(IsNull(gRs!宽度), "", gRs!宽度)
        Itmx.SubItems(4) = IIf(IsNull(gRs!高度), "", gRs!高度)
        Itmx.SubItems(5) = IIf(IsNull(gRs!修改日期), "", gRs!修改日期)
        
        Itmx.Selected = True
        
    End If
    Exit Sub
Handerr:
    If ErrCenter() = 1 Then Resume
    Exit Sub
errHand1:
    
    If ErrCenter() = 1 Then Resume
    
End Sub



Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

