VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmAdvice 
   Caption         =   "���Ź������"
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
         Text            =   "����"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "���"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "����"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "����(����)"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "�߶�(����)"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "�޸�����"
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
            Name            =   "����"
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
               Caption         =   "Ԥ��"
               Key             =   "Ԥ��"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "ɾ��"
               Object.ToolTipText     =   "ɾ��"
               Object.Tag             =   "ɾ��"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_2"
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����˳������"
               Object.Tag             =   "����"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����˳������"
               Object.Tag             =   "����"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_3"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�鿴"
               Key             =   "�鿴"
               Object.ToolTipText     =   "���Ź��鿴��ʽ"
               Object.Tag             =   "�鿴"
               ImageIndex      =   7
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "��ͼ��"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Сͼ��"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "�б�"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "��ϸ����"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_4"
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
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
            Text            =   "��������"
            TextSave        =   "��������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾����"
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
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFilePreView 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileUpdatePage 
         Caption         =   "���²�ѯҳ��(&U)"
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditNew 
         Caption         =   "����(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "ɾ��(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditUp 
         Caption         =   "����˳������(&U)"
      End
      Begin VB.Menu mnuEditDown 
         Caption         =   "����˳������(&D)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuviewsplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "��ͼ��(&G)"
         Index           =   0
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "Сͼ��(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "�б�(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "��ϸ����(&D)"
         Checked         =   -1  'True
         Index           =   3
      End
      Begin VB.Menu mnuViewSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewPos 
         Caption         =   "���λ��(&P)"
      End
      Begin VB.Menu mnuViewSplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "ˢ��(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "��������(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&Web�ϵ�����"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "������̳(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&K)..."
         End
      End
      Begin VB.Menu mnuHelpSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
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
Private mvarSvrPicRange As String           '��������ͼƬ�ķ�Χ
Private mvarSvrPicType As String            '��������ͼƬ������

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
    '���ݴ���״̬,���������и��ؼ�����ʾλ��
    
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
    lblTitle.Caption = "ͼƬ����:" & Item.Text & " ����:" & Item.SubItems(1) & " �߶�:" & Item.SubItems(2)
    picBack.Tag = Item.Tag
    Call ShowPicture
    Call AdjustEnabled
End Sub

Private Sub lvw_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And mnuEdit.Visible Then Me.PopupMenu mnuEdit, 2
End Sub

Private Sub mnuEditDown_Click()
    '����ǰ����Ŀ������һ�У�ͬʱ�������ݿ�

    Dim svrAry(6) As String
    Dim intPre As Long
    Dim strSQL(3) As String
    
    If lvw.SelectedItem Is Nothing Then Exit Sub
    
    intPre = lvw.SelectedItem.Index + 1
    
    If intPre < lvw.ListItems.Count + 1 Then
        strSQL(0) = "zl_��ѯ�������_update(" & Val(Mid(lvw.SelectedItem.Key, 2)) & ",-1)"
        strSQL(1) = "zl_��ѯ�������_update(" & Val(Mid(lvw.ListItems(intPre).Key, 2)) & "," & Val(Mid(lvw.SelectedItem.Key, 2)) & ")"
        strSQL(2) = "zl_��ѯ�������_update(-1," & Val(Mid(lvw.ListItems(intPre).Key, 2)) & ")"
        
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
    
    If frmPicSelect.OpenPictureBox(Me, "�������", "2;0;1;4;9", lngKey, mvarSvrPicRange, mvarSvrPicType) Then
        '����ͼƬ��ʾ
        
'        lngNo = NextValue("��ѯ�������", "���")
'        gstrSQL = "zl_��ѯ�������_insert(" & lngNo & "," & lngKey & ")"
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
    
    If MsgBox("��ȷ��Ҫ�Ƴ����[" & lvw.SelectedItem.Text & "]��", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Exit Sub
    
    On Error GoTo errHand

    'gcnOracle.Execute "zl_��ѯ�������_delete(" & Val(Mid(lvw.SelectedItem.Key, 2)) & ")", , adCmdStoredProc
    gstrSQL = "zl_��ѯ�������_delete(" & Val(Mid(lvw.SelectedItem.Key, 2)) & ")"
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
    '����ǰ����Ŀ������һ�У�ͬʱ�������ݿ�
    
    Dim svrAry(6) As String
    Dim intPre As Long
    Dim strSQL(3) As String
    
    intPre = lvw.SelectedItem.Index - 1
    
    If intPre > 0 Then
    
        strSQL(0) = "zl_��ѯ�������_update(" & Val(Mid(lvw.SelectedItem.Key, 2)) & ",-1)"
        strSQL(1) = "zl_��ѯ�������_update(" & Val(Mid(lvw.ListItems(intPre).Key, 2)) & "," & Val(Mid(lvw.SelectedItem.Key, 2)) & ")"
        strSQL(2) = "zl_��ѯ�������_update(-1," & Val(Mid(lvw.ListItems(intPre).Key, 2)) & ")"
        
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
    Call frmPosSample.ShowPageSample("���ͼƬ")
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
    '��ʾͼ�ΰ���ͼƬ��FLASH��Ӱ
    gstrSQL = "select ͼ��,����,�߶�,���� from ��ѯͼƬԪ�� where ���=[1]"
    
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(picBack.Tag))
    If gRs.BOF = False Then
        Call picBack.ShowPictureByField(gRs!ͼ��, gRs!���� * Screen.TwipsPerPixelX, gRs!�߶� * Screen.TwipsPerPixelY, IIf(IsNull(gRs!����), 0, gRs!����))
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
    Case "Ԥ��"
        Call mnuFilePreView_Click
    Case "��ӡ"
        Call mnuFilePrint_Click
    Case "����"
        Call mnuEditNew_Click
    Case "ɾ��"
        Call mnuEditDelete_Click
    Case "����"
        Call mnuEditUp_Click
    Case "����"
        Call mnuEditDown_Click
    Case "�鿴"
        If lvw.View < 3 Then
            Call mnuViewIcon_Click(lvw.View + 1)
        Else
            Call mnuViewIcon_Click(0)
        End If
    Case "����"
        Call mnuHelpTopic_Click
    Case "�˳�"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then Me.PopupMenu Me.mnuViewTool, 2
End Sub

Private Sub LoadAdviceList()
    '���������úõĹ��
    
    Dim Itmx As ListItem
    Dim lngNO As Long
    
    On Error GoTo errHand
    
    lvw.ListItems.Clear
    picBack.Cls
    
    gstrSQL = "select A.ͼƬ���,A.���,B.����,B.����,B.����,B.�߶�,B.�޸����� from ��ѯ������� A,��ѯͼƬԪ�� B where A.ͼƬ���=B.��� order by A.���"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If gRs.BOF = False Then
        While Not gRs.EOF
            lngNO = lngNO + 1
            Set Itmx = lvw.ListItems.Add(, "K" & gRs!���, IIf(IsNull(gRs!����), "", gRs!����), "pic", "pic")
            Itmx.Tag = IIf(IsNull(gRs!ͼƬ���), 0, gRs!ͼƬ���)
            Itmx.SubItems(1) = lngNO
            Select Case IIf(IsNull(gRs!����), 0, gRs!����)
            Case 0
                Itmx.SubItems(2) = "ͼƬ"
                Itmx.Icon = "pic"
            Case 1
                Itmx.SubItems(2) = "ͼ��"
                Itmx.Icon = "ico"
            Case 2
                Itmx.SubItems(2) = "Flash"
                Itmx.Icon = "swf"
            End Select
            Itmx.SmallIcon = Itmx.Icon
                        
            Itmx.SubItems(3) = IIf(IsNull(gRs!����), "", gRs!����)
            Itmx.SubItems(4) = IIf(IsNull(gRs!�߶�), "", gRs!�߶�)
            Itmx.SubItems(5) = IIf(IsNull(gRs!�޸�����), "", gRs!�޸�����)
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
    Case "��ͼ��"
        Call mnuViewIcon_Click(0)
    Case "Сͼ��"
        Call mnuViewIcon_Click(1)
    Case "�б�"
        Call mnuViewIcon_Click(2)
    Case "��ϸ����"
        Call mnuViewIcon_Click(3)
    End Select
End Sub

Private Sub PrintObject(ByVal intMode As Byte)
    '---------------------------------------------------
    '���ܣ�    ������Ļ��֯���ϸ�����Ŀ����ӡԤ��
    '������
    '     intMode: 2��ʾԤ�� 1��ӡ 3�����EXCEL
    '���أ�
    '---------------------------------------------------
    
    Dim objPrint As New zlPrintLvw
    Dim objRow As New zlTabAppRow

    If lvw.SelectedItem Is Nothing Then Exit Sub

    If UserInfo.���� = "" Then Call GetUserInfo

    objPrint.Title = "���Ź�������嵥"
    objPrint.BelowAppItems.Add "��ӡ��:" & UserInfo.����
    objPrint.BelowAppItems.Add "��ӡʱ��:" & Format(zlDatabase.Currentdate, "YYYY��MM��DD��")
    objPrint.Footer = "��[ҳ��]ҳ;;"

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
    '����ģ��Ȩ��,��������������ػ���ʾ;Ȩ����:��ɾ��
    
'    mnuEdit.Visible = True
'    If InStr(gstrPrivs, "��ɾ��") = 0 Then
'        mnuEdit.Visible = False
'
'        tbrThis.Buttons("����").Visible = False
'        tbrThis.Buttons("ɾ��").Visible = False
'        tbrThis.Buttons("Split_2").Visible = False
'
'        tbrThis.Buttons("����").Visible = False
'        tbrThis.Buttons("����").Visible = False
'        tbrThis.Buttons("Split_3").Visible = False
'    End If
    
End Sub

Private Sub AdjustEnabled()
    '�������ܲ˵���ť�Ŀ���״̬
    
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
                
    tbrThis.Buttons("Ԥ��").Enabled = mnuFilePreView.Enabled
    tbrThis.Buttons("��ӡ").Enabled = mnuFilePrint.Enabled
    tbrThis.Buttons("����").Enabled = mnuEditNew.Enabled
    tbrThis.Buttons("ɾ��").Enabled = mnuEditDelete.Enabled
    tbrThis.Buttons("����").Enabled = mnuEditDown.Enabled
    tbrThis.Buttons("����").Enabled = mnuEditUp.Enabled
End Sub

Private Sub ReadRegister()
    '��ȡע�����Ϣ
    
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then Exit Sub
    
    picX.Left = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\����\" & Me.Name, "picXλ��", 2385)
    
End Sub

Private Sub WriteRegister()
    '����Ϣд��ע���
    
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\����\" & Me.Name, "picXλ��", picX.Left
End Sub

Private Sub LoadStatus()
    'ˢ��״̬����ʾ
    
    If lvw.ListItems.Count > 0 Then
        stbThis.Panels(2).Text = "��ǰ����" & lvw.ListItems.Count & "��Ҫ���ŵĹ�棡"
    Else
        stbThis.Panels(2).Text = "��ǰû��Ҫ���ŵĹ�棡"
    End If
End Sub

Public Sub AddPicture(ByVal lngKey As Long)
    Dim lngNO As Long
    Dim Itmx As ListItem
    
    On Error GoTo errHand1
    
    lngNO = NextValue("��ѯ�������", "���")
    gstrSQL = "zl_��ѯ�������_insert(" & lngNO & "," & lngKey & ")"
    
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    
    On Error GoTo 0
    On Error GoTo Handerr
    gstrSQL = "select A.ͼƬ���,A.���,B.����,B.����,B.����,B.�߶�,B.�޸����� from ��ѯ������� A,��ѯͼƬԪ�� B where A.ͼƬ���=B.��� and B.���=[1]"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
    If gRs.BOF = False Then
        Set Itmx = lvw.ListItems.Add(, "K" & lngNO, IIf(IsNull(gRs!����), "", gRs!����), "pic", "pic")
        Itmx.Tag = IIf(IsNull(gRs!ͼƬ���), 0, gRs!ͼƬ���)
        Itmx.SubItems(1) = lvw.ListItems.Count
        Select Case IIf(IsNull(gRs!����), 0, gRs!����)
        Case 0
            Itmx.SubItems(2) = "ͼƬ"
            Itmx.Icon = "pic"
        Case 1
            Itmx.SubItems(2) = "ͼ��"
            Itmx.Icon = "ico"
        Case 2
            Itmx.SubItems(2) = "Flash"
            Itmx.Icon = "swf"
        End Select
        Itmx.SmallIcon = Itmx.Icon
                    
        Itmx.SubItems(3) = IIf(IsNull(gRs!����), "", gRs!����)
        Itmx.SubItems(4) = IIf(IsNull(gRs!�߶�), "", gRs!�߶�)
        Itmx.SubItems(5) = IIf(IsNull(gRs!�޸�����), "", gRs!�޸�����)
        
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
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub
