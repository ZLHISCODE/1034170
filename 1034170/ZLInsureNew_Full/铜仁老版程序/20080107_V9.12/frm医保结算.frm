VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmҽ������ 
   Caption         =   "ҽ���������"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9525
   Icon            =   "frmҽ������.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picSplitV 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2850
      Left            =   6150
      MousePointer    =   9  'Size W E
      ScaleHeight     =   2850
      ScaleWidth      =   90
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2190
      Width           =   90
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   5445
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmҽ������.frx":08CA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11748
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
            AutoSize        =   2
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
   Begin VB.PictureBox picSplitH 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   -60
      MousePointer    =   7  'Size N S
      ScaleHeight     =   90
      ScaleWidth      =   9195
      TabIndex        =   6
      Top             =   3960
      Width           =   9195
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   5220
      Top             =   420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":115C
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":1376
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":1590
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":17AA
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":19C4
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":1BDE
            Key             =   "Merge"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":22D8
            Key             =   "View"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":29D2
            Key             =   "Patis"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":30CC
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":32E6
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":3500
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":371A
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   5850
      Top             =   420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":3934
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":3B4E
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":3D68
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":3F82
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":419C
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":43B6
            Key             =   "Merge"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":4AB0
            Key             =   "View"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":51AA
            Key             =   "Patis"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":58A4
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":5ABE
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":5CD8
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmҽ������.frx":5EF2
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   720
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   1270
      BandCount       =   2
      ForeColor       =   -2147483635
      FixedOrder      =   -1  'True
      _CBWidth        =   9525
      _CBHeight       =   720
      _Version        =   "6.7.8988"
      Child1          =   "tbrThis"
      MinWidth1       =   4995
      MinHeight1      =   660
      Width1          =   810
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Caption2        =   "�������"
      Child2          =   "cmb����"
      MinWidth2       =   1995
      MinHeight2      =   300
      Width2          =   1935
      NewRow2         =   0   'False
      AllowVertical2  =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   660
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   1164
         ButtonWidth     =   820
         ButtonHeight    =   1164
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Preview"
               Description     =   "Ԥ��"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Description     =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Edit_1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Find"
               Description     =   "����"
               Object.ToolTipText     =   "����ҽ���ʻ�"
               Object.Tag             =   "����"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "����"
               Object.ToolTipText     =   "��ǰ��������"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Quit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox cmb���� 
         Height          =   300
         Left            =   7440
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   210
         Width           =   1995
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh��¼_S 
      Height          =   2805
      Left            =   60
      TabIndex        =   3
      Top             =   1110
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   4948
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frmҽ������.frx":610C
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh��ϸ 
      Height          =   1335
      Left            =   0
      TabIndex        =   5
      Top             =   4050
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   2355
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frmҽ������.frx":6426
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh�ֵ� 
      Height          =   1335
      Left            =   4710
      TabIndex        =   8
      Top             =   4020
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   2355
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frmҽ������.frx":6740
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.TabStrip tab���� 
      Height          =   345
      Left            =   30
      TabIndex        =   7
      Top             =   750
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   609
      TabWidthStyle   =   2
      Style           =   2
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�շ�"
            Key             =   "K1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
            Key             =   "K2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Ԥ��"
            Key             =   "K3"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFileSplitSet 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileSplitPrint 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFileSplitExcel 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileBalance 
         Caption         =   "���ᷢƱ��Ϣ(&F)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBill 
         Caption         =   "��ӡ��������(&B)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileDetail 
         Caption         =   "������ϸ(&D)"
      End
      Begin VB.Menu mnuFileBatch 
         Caption         =   "��ϸ������ӡ(&B)"
      End
      Begin VB.Menu mnuFileSplitReport 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileQuit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Visible         =   0   'False
      Begin VB.Menu mnuEditClinic 
         Caption         =   "���ﱨ��(&C)"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuEditIn_Patient 
         Caption         =   "סԺ����(&P)"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "����(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditXE 
         Caption         =   "����޶�(&X)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditView 
         Caption         =   "����(&V)"
      End
   End
   Begin VB.Menu mnuBalance 
      Caption         =   "����(&B)"
      Visible         =   0   'False
      Begin VB.Menu mnuBalanceBill 
         Caption         =   "��ȡ���㵥(&B)"
      End
      Begin VB.Menu mnuBalanceCollect 
         Caption         =   "��ȡ�����(&C)"
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
         Begin VB.Menu mnuViewTool_1 
            Caption         =   "-"
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
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "����(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuView_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "ˢ��(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "��������(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&WEB�ϵ�����"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&K)..."
         End
      End
      Begin VB.Menu mnuHelp_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmҽ������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private Enum ��¼Enum
    col��¼ID = 0
    col���ݺ� = 1
    col���� = 2
    col���� = 3
    col����ID = 4
    colסԺ�� = 5
    col���� = 6
    col��� = 7
    col�Ա� = 8
    col���� = 9
    col���� = 10
    col����Ա���� = 11
    col�Ǽ�ʱ�� = 12
    col���˱�־ = 13
    col�����ʻ� = 14
    col�������� = 15
    colʵ������ = 16
    col����ͳ�� = 17
    colͳ�ﱨ�� = 18
End Enum

Private Enum ��ϸEnum
    det�շ���� = 0
    det�շ�ϸĿ = 1
    det��� = 2
    det��λ = 3
    det���� = 4
    det���� = 5
    detʵ�ս�� = 6
    detͳ���� = 7
    detҽ������ = 8
    det�������� = 9
    det���� = 10
    det״̬ = 11
End Enum

Private mblnLoad As Boolean                     '��һ������

Private mint���� As Integer
Private mint���� As Integer
Private mdatBegin As Date, mdatEnd As Date
Private mstrCardCond As String

Dim msngStartX As Single, msngStartY As Single    '�ƶ�ǰ����λ��
Private mrs�����¼ As New ADODB.Recordset
Private mcol���� As New Collection              '����ҽ����������������

Private Sub cbr_HeightChanged(ByVal NewHeight As Single)
    Call ResizeForm(NewHeight)
End Sub

Private Sub cmb����_Click()
    With cmb����
        If mint���� = .ItemData(.ListIndex) Then Exit Sub
        mint���� = .ItemData(.ListIndex)
        mnuFileBalance.Visible = False
        If mint���� = TYPE_������ Then
            mnuFileBalance.Visible = True
            mnuBalance.Visible = True
        End If
    End With
    Call Ȩ�޿���
    Call FillList
End Sub

Private Sub Form_Activate()
    If mblnLoad = True Then
        mdatBegin = CDate(Format(zlDatabase.Currentdate, "yyyy-MM-dd"))
        mdatEnd = CDate(Format(mdatBegin, "yyyy-MM-dd") & " 23:59:59")
        mstrCardCond = ""
        
        'ǿ����ʾ
        msh��ϸ.Visible = False
        '��ʾ��¼
        Call tab����_Click
    End If
    mblnLoad = False
End Sub

Private Sub Form_Load()
    mblnLoad = True
    mint���� = -1
    mint���� = -1
    
    mstrPrivs = gstrPrivs
    zlControl.CboSetHeight cmb����, 3600
    Call InitTable
    
    RestoreWinState Me, App.ProductName
    Call Ȩ�޿���
End Sub

Private Sub Form_Resize()
    Call ResizeForm(cbr.Height)
End Sub

Private Sub ResizeForm(ByVal cbrHeight As Single)
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    
    sngTop = IIf(cbr.Visible, cbr.Top + cbrHeight, 0)
    sngBottom = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    
    If Me.WindowState = 1 Then Exit Sub
    
    tab����.Left = ScaleLeft
    tab����.Width = ScaleWidth
    tab����.Top = sngTop
    
    msh��¼_S.Left = ScaleLeft
    msh��¼_S.Width = ScaleWidth
    msh��¼_S.Top = tab����.Top + tab����.Height
    
    If picSplitH.Visible = False Then
        '����ʾԤ����¼ʱû����ϸ
        msh��¼_S.Height = IIf(sngBottom - msh��¼_S.Top > 0, sngBottom - msh��¼_S.Top, 0)
        Exit Sub
    End If
    
    
    picSplitH.Left = ScaleLeft
    picSplitH.Width = ScaleWidth
    picSplitH.Top = msh��¼_S.Top + msh��¼_S.Height
    
    msh��ϸ.Left = ScaleLeft
    msh��ϸ.Top = picSplitH.Top + picSplitH.Height
    msh��ϸ.Height = IIf(sngBottom - msh��ϸ.Top > 0, sngBottom - msh��ϸ.Top, 0)
    
    msh�ֵ�.Left = IIf(ScaleWidth - msh�ֵ�.Width > 0, ScaleWidth - msh�ֵ�.Width, 0)
    picSplitV.Left = msh�ֵ�.Left - picSplitV.Width
    If msh�ֵ�.Visible = False Then
        '����ʾ�շѼ�¼ʱ��û�зֵ�ͳ������
        msh��ϸ.Width = IIf(ScaleWidth - msh��ϸ.Left > 0, ScaleWidth - msh��ϸ.Left, 0)
        Exit Sub
    Else
        msh��ϸ.Width = IIf(picSplitV.Left - msh��ϸ.Left > 0, picSplitV.Left - msh��ϸ.Left, 0)
    End If
    
    msh�ֵ�.Top = msh��ϸ.Top
    msh�ֵ�.Height = msh��ϸ.Height
    picSplitV.Top = msh��ϸ.Top
    picSplitV.Height = msh��ϸ.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub mnuBalanceBill_Click()
    '��ʽ��1-����;2-����涨��;3-סԺ
    '��ȡ�������˵Ľ��㵥
    Const strBill As String = "ZL1_INSIDE_1605_10"
    Dim lng����ID As Long, lng����ID As Long
    Dim intҵ������ As Integer
    Dim strҵ�����к� As String
    On Error GoTo ErrHand
    
    lng����ID = Val(msh��¼_S.TextMatrix(msh��¼_S.Row, col����ID))
    If lng����ID = 0 Then Exit Sub
    lng����ID = Val(msh��¼_S.TextMatrix(msh��¼_S.Row, col��¼ID))
    If lng����ID = 0 Then Exit Sub
    
    Select Case mint����
    Case TYPE_������
        If Not ��ȡ���㵥_������(lng����ID, lng����ID, intҵ������, strҵ�����к�) Then Exit Sub
        '������Ԥ��
        Call ReportOpen(gcnOracle, glngSys, strBill, Me, "ҵ�����к�=" & strҵ�����к�, "ReportFormat=" & intҵ������, 1)
    Case Else
        Exit Sub
    End Select
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub mnuBalanceCollect_Click()
    '��ʽ��1-����;2-����涨��;3-סԺ
    '��ȡ������ܱ�����ҽԺ��ӡ�����ĵĶ��ʵ���
    Const strBill As String = "ZL1_INSIDE_1605_11"
    On Error GoTo ErrHand
    
    Select Case mint����
    Case TYPE_������
        If Not ��ȡ�����_������() Then Exit Sub
        
        '������Ԥ��
        Call ReportOpen(gcnOracle, glngSys, strBill, Me)
    Case Else
        Exit Sub
    End Select
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub mnuEditClinic_Click()
    If frm���ñ���.ShowMe(1) And mint���� = 1 Then Call FillList
End Sub
    
Private Sub mnuEditDelete_Click()
    Dim lng��¼id As Long, lng����ID As Long, lng����ID As Long
    Dim intסԺ���� As Integer, lng��� As Long, lng��ǰ��� As Long
    Dim cur�ʻ����_��� As Currency, curͳ���ۼ�_��� As Currency
    Dim cur�ʻ�֧�� As Currency, cur����ͳ�� As Currency '��������ʻ���סԺ����ͳ��
    Dim int��Ժ As Integer, int��Ժ As Integer, int��ǰסԺ����_��Ժ As Integer, int��ǰסԺ����_��Ժ As Integer
    Dim bln��Ժ As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim rsTest As New ADODB.Recordset
    On Error GoTo ErrHand
    
    If mint���� = 3 Then Exit Sub
    If Trim(msh��¼_S.TextMatrix(msh��¼_S.Row, col���˱�־)) = "��" Then
        MsgBox "���ܶ����ϵ���ִ�����ϲ�����ֻ�ܶ��շ����͵ĵ��ݽ������ϣ���", vbInformation, gstrSysName
        Exit Sub
    End If
    gstrSQL = "Select Count(*) Records From ���ս����¼ Where ����=25 And ֧��˳���=" & Val(msh��¼_S.TextMatrix(msh��¼_S.Row, col��¼ID))
    Call OpenRecordset(rsTemp, Me.Caption)
    If rsTemp!Records = 1 Then
        MsgBox "�õ����ѱ����ϣ������ٴ����ϣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If MsgBox("��ȷ��Ҫ���ϸý����¼��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    '����
    lng��ǰ��� = Format(zlDatabase.Currentdate, "yyyy")
    lng��¼id = Val(msh��¼_S.TextMatrix(msh��¼_S.Row, col��¼ID))
    If lng��¼id = 0 Then Exit Sub
    lng����ID = zlDatabase.GetNextId("���˽��ʼ�¼")
    
    gcnOracle.BeginTrans
    '���渺�ı��ս����¼
    With rsTemp
        gstrSQL = "Select * From ���ս����¼ Where ����=" & gintInsure & " And ��¼ID=" & lng��¼id
        Call OpenRecordset(rsTemp, Me.Caption)
        intסԺ���� = NVL(!סԺ����, 1) - 1
        lng����ID = !����ID
        lng��� = !���
        cur�ʻ�֧�� = NVL(!�����ʻ�֧��, 0)
        cur����ͳ�� = NVL(!����ͳ����, 0)
        
        'ֻ���������һ�εĽ����¼
        gstrSQL = "Select Nvl(סԺ�����ۼ�,0) ��Ժ,Nvl(��ԺסԺ����,0) ��Ժ From �ʻ������Ϣ Where ���=" & lng��ǰ��� & " ANd ����ID =" & lng����ID
        Call OpenRecordset(rsTest, Me.Caption)
        If rsTest.EOF Then
            int��ǰסԺ����_��Ժ = 0
            int��ǰסԺ����_��Ժ = 0
        Else
            int��ǰסԺ����_��Ժ = rsTest!��Ժ - IIf(mint���� = 2, 1, 0)
            int��ǰסԺ����_��Ժ = rsTest!��Ժ - IIf(mint���� = 2, 1, 0)
        End If
        If lng��� <> Format(zlDatabase.Currentdate, "yyyy") Then
            MsgBox "���ܳ���������ȵĵ��ݣ�", vbInformation, gstrSysName
            gcnOracle.RollbackTrans
            Exit Sub
        End If
        
        gstrSQL = "Select * From ���ս����¼ Where ����=" & gintInsure & " And ��¼ID=" & lng��¼id
        Call OpenRecordset(rsTemp, Me.Caption)
        gstrSQL = "zl_���ս����¼_insert(" & !���� & "," & lng����ID & ",25," & !����ID & "," & _
            lng��� & ",0,0,0,0," & NVL(!סԺ����, 0) & "," & -1 * NVL(!����, 0) & ",0," & -1 * NVL(!ʵ������, 0) & "," & _
            -1 * NVL(!�������ý��, 0) & "," & -1 * NVL(!ȫ�Ը����, 0) & "," & -1 * NVL(!�����Ը����, 0) & "," & -1 * NVL(!����ͳ����, 0) & "," & -1 * NVL(!ͳ�ﱨ�����, 0) & ",0," & _
            0 & "," & -1 * cur�ʻ�֧�� & ",'" & lng��¼id & "',null,null,null,null,'" & gstrUserName & "')" '֧��˳����������汻�����ļ�¼ID
        Call ExecuteProcedure("����סԺ��������")
        '���渺�ı��ձ�����¼(����)
        gstrSQL = "Select * From ���ձ�����¼ Where ��¼ID=" & lng��¼id
        Call OpenRecordset(rsTemp, Me.Caption)
        Do While Not .EOF
            gstrSQL = "ZL_���ձ�����¼_INSERT(" & !���� & "," & lng����ID & "," & _
            "'" & !������� & "','" & !�������� & "'," & !ͳ��ȶ� & "," & _
            "" & !��׼���� & "," & !��׼���� & "," & -1 * !�����ܶ� & "," & -1 * !�����ܶ� & ")"
            Call ExecuteProcedure("������౨������")
            bln��Ժ = (!���� = 1)
            .MoveNext
        Loop
        
        '������ǳ������һ�ν����¼����ֹ!
        If mint���� = 2 Then
            If bln��Ժ Then
                If int��ǰסԺ����_��Ժ <> intסԺ���� Then
                    MsgBox "ֻ���������һ�ν���ĵ��ݣ�", vbInformation, gstrSysName
                    gcnOracle.RollbackTrans
                    Exit Sub
                End If
            Else
                If int��ǰסԺ����_��Ժ <> intסԺ���� Then
                    MsgBox "ֻ���������һ�ν���ĵ��ݣ�", vbInformation, gstrSysName
                    gcnOracle.RollbackTrans
                    Exit Sub
                End If
            End If
        End If
        
        gstrSQL = "Select * From ���ս������ Where ����ID=" & lng��¼id
        Call OpenRecordset(rsTemp, Me.Caption)
        '���渺�ķֵ�������¼
        Do While Not .EOF
            gstrSQL = "ZL_���ս������_INSERT(" & lng����ID & "," & !���� & "," & -1 * !����ͳ���� & "," & -1 * !ͳ�ﱨ����� & "," & !���� & ")"
            Call ExecuteProcedure("����ֵ�������ϸ")
            .MoveNext
        Loop
    End With
        
    'ȡ������ʻ���������ͳ��
    gstrSQL = " Select Nvl(�ʻ������ۼ�,0) �ʻ����,Nvl(����ͳ���ۼ�,0) ����ͳ��" & _
              " ,Nvl(סԺ�����ۼ�,0) ��Ժ,Nvl(��ԺסԺ����,0) ��Ժ" & _
              " From �ʻ������Ϣ " & _
              " Where ���=" & lng��ǰ��� & " And ����ID=" & lng����ID ' gComInfo_üɽ.����ID
    Call OpenRecordset(rsTemp, Me.Caption)
    
    cur�ʻ����_��� = 0
    curͳ���ۼ�_��� = 0
    int��Ժ = 0: int��Ժ = 0
    If Not rsTemp.EOF Then
        cur�ʻ����_��� = NVL(rsTemp!�ʻ����, 0)
        curͳ���ۼ�_��� = NVL(rsTemp!����ͳ��, 0)
        int��Ժ = rsTemp!��Ժ
        int��Ժ = rsTemp!��Ժ
    End If
    If bln��Ժ Then
        int��Ժ = int��ǰסԺ����_��Ժ
    Else
        int��Ժ = int��ǰסԺ����_��Ժ
    End If
    '�����¸����ʻ��������Զ������ʻ�����˱���������
'    cur�ʻ����_��� = cur�ʻ����_��� + cur�ʻ�֧��
    '�������ʻ����㷨��סԺһ����ֻ�����ʻ�֧�������Ҳ���ڽ���ͳ����Դ˴�Ҫ�ж�һ��
    curͳ���ۼ�_��� = curͳ���ۼ�_��� - IIf(mint���� = 1, 0, cur����ͳ��)
    
    gstrSQL = "zl_�ʻ������Ϣ_Insert(" & lng����ID & ",25," & lng��� & _
              "," & cur�ʻ����_��� & ",0," & curͳ���ۼ�_��� & ",0," & int��Ժ & "," & int��Ժ & ")"
    Call ExecuteProcedure("����סԺ����")
    
    '���ʻ�����
    If cur�ʻ�֧�� <> 0 Then
        If Not �¸����ʻ�(lng����ID, cur�ʻ�֧��) Then Exit Sub 'gComInfo_üɽ.����ID
    End If
    
    gcnOracle.CommitTrans
    Call FillList
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    gcnOracle.RollbackTrans
End Sub

Private Sub mnuEditIn_Patient_Click()
    If frm���ñ���.ShowMe(1, True) And mint���� = 2 Then Call FillList
End Sub

Private Sub mnuEditView_Click()
    Dim lng����ID  As Long
    
    If mint���� = 1 Or mint���� = 2 Then
        lng����ID = Val(msh��¼_S.TextMatrix(msh��¼_S.Row, col��¼ID))
        If lng����ID = 0 Then Exit Sub
        Call frm���ñ���.ShowMe(2, (mint���� = 2), lng����ID)
    End If
End Sub

Private Sub mnuEditXE_Click()
  '��Ҫ¼�����������޶�
    Dim lng����ID As Long
    Dim strIdentify As String
    Dim bytType As Byte
    Dim clsҽ�� As New clsInsure
    Dim lng���� As Long
    Dim lng��¼id As Long
    Dim int���� As Long
    Dim frmMain As New frmIdentify����
    
    lng��¼id = Val(msh��¼_S.TextMatrix(msh��¼_S.Row, col��¼ID))
    If lng��¼id = 0 Then Exit Sub
    
    lng����ID = Val(msh��¼_S.TextMatrix(msh��¼_S.Row, col����ID))
    If lng����ID = 0 Then Exit Sub
    
    int���� = Val(Mid(tab����.SelectedItem.Key, 2))
    
    strIdentify = frmMain.GetPatient(9, lng����ID, int����, lng��¼id)
    
    mint���� = 0
    
    If strIdentify <> "" Then
        tab����_Click
    End If
End Sub

Private Sub mnuFileBalance_Click()
    Dim lng����ID As Long, lng��¼id As Long
    Dim strҽԺ���� As String, strҵ�����к� As String
    Dim rsTemp As New ADODB.Recordset
    'ֻ��������·ҽ�����ڸù��ܣ����Ե��ýӿڻ�ȡĳ�ν������Ϣ��������ʱ�����Թ���ӡ֮��
    lng����ID = Val(msh��¼_S.TextMatrix(msh��¼_S.Row, col����ID))
    lng��¼id = Val(msh��¼_S.TextMatrix(msh��¼_S.Row, col��¼ID))
    If lng����ID = 0 Then Exit Sub
    If lng��¼id = 0 Then Exit Sub
    
    gstrSQL = "Select ҽԺ���� From ������� Where ���=" & TYPE_������
    Call OpenRecordset(rsTemp, "��ȡҽԺ����")
    strҽԺ���� = NVL(rsTemp!ҽԺ����)
    If Trim(strҽԺ����) = "" Then
        MsgBox "��������ҽԺ�������ʹ�øù��ܣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    gstrSQL = "Select * From ���ս����¼ Where ����=" & TYPE_������ & " And ��¼ID=" & lng��¼id
    Call OpenRecordset(rsTemp, "��ȡҵ�����к�")
    If rsTemp.EOF Then
        MsgBox "û���ҵ��κν����¼��", vbInformation, gstrSysName
        Exit Sub
    End If
    strҵ�����к� = NVL(rsTemp!֧��˳���)
    If Trim(strҵ�����к�) = "" Then
        MsgBox "���ս������ݴ����޷���������ҵ�����кŲ���Ϊ�գ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '20031228:���:�������ID
    Call GetBalance(lng����ID, lng��¼id, strҵ�����к�, strҽԺ����)
End Sub

Private Sub mnuFileBill_Click()
    Dim lng����ID  As Long
    lng����ID = Val(msh��¼_S.TextMatrix(msh��¼_S.Row, col��¼ID))
    If lng����ID = 0 Then Exit Sub
    Call zl9Report.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1605_" & mint����, Me, "����=" & mint����, "��¼ID=" & lng����ID, 2)
End Sub

Private Sub mnuHelpTitle_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub mnuFileQuit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuViewFind_Click()
    If frmҽ���������.GetTimeScope(mdatBegin, mdatEnd, mstrCardCond, Me) = True Then
        Call FillList
    End If
End Sub

Private Sub mnuViewRefresh_Click()
    Call FillList
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = Not stbThis.Visible
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    cbr.Visible = Not cbr.Visible
    mnuViewToolText.Enabled = Not mnuViewToolText.Enabled
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim i As Integer
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To tbrThis.Buttons.Count
        tbrThis.Buttons(i).Caption = IIf(mnuViewToolText.Checked, tbrThis.Buttons(i).Tag, "")
    Next
    cbr.Bands(1).MinHeight = tbrThis.ButtonHeight
    Form_Resize
End Sub

Private Sub msh��¼_S_EnterCell()
    'ѡ��ĳ���ʻ�,����ȡ�����Ϣ
    Call FillDetail
End Sub

Private Sub msh��¼_S_KeyDown(KeyCode As Integer, Shift As Integer)
    If mint���� = TYPE_�Ĵ�üɽ Then
        If KeyCode = vbKeyReturn Then Call mnuEditView_Click
    End If
End Sub

Private Sub msh��¼_S_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim strSort As String
    
    If Button = 1 Then
        '����ͷ����
        If msh��¼_S.MouseRow = 0 Then
            strSort = msh��¼_S.TextMatrix(0, msh��¼_S.MouseCol)
            If strSort = "סԺ��" And mint���� = 1 Then strSort = "�����"
            
            If strSort = "" Then Exit Sub
            If mrs�����¼.Sort = strSort Then
                mrs�����¼.Sort = strSort & " DESC"
            Else
                mrs�����¼.Sort = strSort
            End If
            Call ������(msh��¼_S, mrs�����¼)
        End If
    
    End If
End Sub

Private Sub tab����_Click()
    Dim int���� As Integer
    Dim sngHeight As Single
    Call Ȩ�޿���
    int���� = Val(Mid(tab����.SelectedItem.Key, 2))
    If mint���� = int���� Then Exit Sub
    
    mint���� = int����
    
    Select Case mint����
        Case 1 '�շ�
            msh�ֵ�.Visible = False
            picSplitV.Visible = False
            
            If msh��ϸ.Visible = False Then
                'ǰһ��״̬����ʾԤ��
                msh��ϸ.Visible = True
                picSplitH.Visible = True
                
                sngHeight = ScaleHeight - IIf(stbThis.Visible = False, 0, stbThis.Height) - msh��¼_S.Top
                
                If sngHeight - msh��¼_S.Height < 1000 Then
                    msh��¼_S.Height = msh��¼_S.Height / 2
                End If
            End If
        Case 2 '����
            msh�ֵ�.Visible = True
            picSplitV.Visible = True
            If msh��ϸ.Visible = False Then
                'ǰһ��״̬����ʾԤ��
                msh��ϸ.Visible = True
                picSplitH.Visible = True
                msh��¼_S.Height = msh��¼_S.Height / 2
            End If
            
        Case 3 '
            picSplitH.Visible = False
            msh��ϸ.Visible = False
            msh�ֵ�.Visible = False
            picSplitV.Visible = False
    End Select
    '���µ���
    Call Form_Resize
    '��ʾ����
    Call FillList
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFileQuit_Click
        Case "Find"
            mnuViewFind_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Preview"
            mnuFilePreview_Click
        Case "Help"
            mnuHelpTitle_Click
    End Select
End Sub

Private Sub tbrThis_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Sub mnuFileDetail_Click()
    Dim lng����ID As Long
    
    lng����ID = Val(msh��¼_S.TextMatrix(msh��¼_S.Row, col��¼ID))
    If lng����ID <> 0 Then
        Call zl9Report.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1605", Me, "����=" & mint����, "ID=" & lng����ID, 1)
    End If
End Sub

Private Sub mnuFileBatch_Click()
    Dim lngRow As Integer, int���� As Integer
    Dim lng����ID As Long
    
    '����������¼
    For lngRow = 1 To msh��¼_S.Rows - 1
        lng����ID = Val(msh��¼_S.TextMatrix(lngRow, col��¼ID))
        If lng����ID <> 0 Then
            Call zl9Report.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1605", Me, "����=" & mint����, "ID=" & lng����ID, 1)
        End If
    Next
End Sub

Private Sub mnuFileExcel_Click()
    Call subPrint(3)
End Sub

Private Sub mnuFilePreview_Click()
    Call subPrint(2)
End Sub

Private Sub mnuFilePrint_Click()
    Call subPrint(1)
End Sub

Private Sub mnuFilePrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub subPrint(bytStyle As Byte)
    '���ܣ�������б�
    '������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim bytR As Byte, intRow As Integer
    
    If gstrUserName = "" Then GetUserInfo
    intRow = msh��¼_S.Row
    
    '��ͷ
    objOut.Title.Text = "ҽ�����˷��ý����嵥��" & tab����.SelectedItem.Caption & "��"
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '����
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate, "yyyy��MM��DD��")
    objOut.BelowAppRows.Add objRow
    
    '����
    Set objOut.Body = msh��¼_S
    
    '���
    msh��¼_S.Redraw = False
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    msh��¼_S.Redraw = True
    
    msh��¼_S.Row = intRow
    msh��¼_S.Col = 0: msh��¼_S.ColSel = msh��¼_S.Cols - 1
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage Me.hwnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo Me.hwnd
End Sub

Private Sub picSplitH_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        msngStartY = y
    End If
End Sub

Private Sub picSplitH_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = picSplitH.Top + y - msngStartY
        If sngTemp > msh��¼_S.Top + 1000 And (msh��ϸ.Top + msh��ϸ.Height) - (sngTemp + picSplitH.Height) > 1000 Then
            picSplitH.Top = sngTemp
            msh��¼_S.Height = picSplitH.Top - msh��¼_S.Top
            Form_Resize
        End If
    End If
End Sub

Private Sub picSplitV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        msngStartX = x
    End If
End Sub

Private Sub picSplitV_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = picSplitV.Left + x - msngStartX
        If sngTemp > msh��ϸ.Left + 1000 And ScaleWidth - (sngTemp + picSplitV.Width) > 1000 Then
            picSplitV.Left = sngTemp
            msh�ֵ�.Width = ScaleWidth - (sngTemp + picSplitV.Width)
            Form_Resize
        End If
    End If
End Sub

Private Function FillList() As Boolean
    '��ȡ�����ʻ�(���Ȩ������,����������ֶ�)������
    Dim strBegin As String
    Dim strEnd As String
    
    strBegin = "to_date('" & Format(mdatBegin, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS') "
    strEnd = "to_date('" & Format(mdatEnd, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS') "
    
    If mrs�����¼.State = adStateOpen Then mrs�����¼.Close
    
    MousePointer = vbHourglass
    On Error GoTo errHandle
    
    Call GetSpecialSQL(mint����, strBegin, strEnd)
    
    mrs�����¼.Sort = ""
    Call OpenRecordset(mrs�����¼, Me.Caption)
    Call ������(msh��¼_S, mrs�����¼)
    
    Call FillDetail
    FillList = True
    MousePointer = vbDefault
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    MousePointer = vbDefault
End Function

Private Sub ������(mshBind As MSHFlexGrid, rsBind As ADODB.Recordset)
    Dim lngCol As Long
    
    '���ʻ�����װ��FLEXGRID��������
    If mshBind Is msh��¼_S Then
        Call Init��¼Table '���ڲ�ͬ�������������ݺܴ�̶��ϲ�ͬ������ÿ�ζ���ʼ��
    End If
    
    With mshBind
        If rsBind.RecordCount <> 0 Then
            Set .DataSource = rsBind
            DoEvents
            .Col = 0
            .Row = .FixedRows - 1
            .ColSel = .Cols - 1
            .RowSel = .Row
            .FillStyle = flexFillRepeat
            .CellAlignment = 4
            .FillStyle = flexFillSingle
            .AllowBigSelection = False
            .Row = .FixedRows: .Col = 0
            .ColSel = .Cols - 1: .RowSel = .Row
            
        Else
            Set .DataSource = Nothing
            .Rows = 2
            For lngCol = 0 To .Cols - 1
                .TextMatrix(1, lngCol) = ""
            Next
        End If
        
        If mshBind Is msh��¼_S Then
            '���ض������
            If mcol����("K" & mint����) = "0" Then
                .ColWidth(col����) = 0
            Else
                If .ColWidth(col����) = 0 Then
                    .ColWidth(col����) = 1000
                End If
            End If
        End If
    End With
End Sub

Private Sub Init��¼Table()
    Dim lngCol As Integer
    
    '���ø�ʽ
    With msh��¼_S
        .Rows = 2
        .Cols = 19 'Ϊ������һЩ��������
        For lngCol = 0 To .Cols - 1
            .ColPosition(lngCol) = 0
        Next
        
        .TextMatrix(0, col��¼ID) = "����"
        .TextMatrix(0, col���ݺ�) = "���ݺ�"
        .TextMatrix(0, col����) = "����"
        .TextMatrix(0, col����) = "����"
        .TextMatrix(0, col����ID) = "����ID"
        .TextMatrix(0, col����) = "����"
        .TextMatrix(0, col���) = "���"
        .TextMatrix(0, col�Ա�) = "�Ա�"
        .TextMatrix(0, col����) = "����"
        .TextMatrix(0, col���˱�־) = "���˱�־"
        .TextMatrix(0, col�����ʻ�) = "�����ʻ�"
        .ColWidth(col��¼ID) = 0
        .ColWidth(col���ݺ�) = 900
        .ColWidth(col����) = 0
        .ColWidth(col����) = 900
        .ColWidth(col����ID) = 800
        .ColWidth(col����) = 800
        .ColWidth(col���) = 600
        .ColWidth(col�Ա�) = 400
        .ColWidth(col����) = 400
        .ColWidth(col���˱�־) = 855
        .ColWidth(col�����ʻ�) = 930
        
        .ColWidth(colסԺ��) = 800
        .ColWidth(col����Ա����) = 1200
        .ColWidth(col�Ǽ�ʱ��) = 1200
        Select Case mint����
            Case 1 '-�շ�
                .Cols = 16
                .TextMatrix(0, colסԺ��) = "סԺ��"
                .TextMatrix(0, col����) = "��������"
                .TextMatrix(0, col����Ա����) = "�շ�Ա"
                .TextMatrix(0, col�Ǽ�ʱ��) = "�շ�ʱ��"
                .TextMatrix(0, col��������) = "��������"
                
                .ColWidth(col����) = 1200
                .ColWidth(col��������) = 930
                
                '�ı�ĳЩ�е���ʾ˳��
                .ColPosition(col�����ʻ�) = col����Ա����
                .ColPosition(col��������) = col����Ա����
            Case 2 '-���㣨����סԺ���㡢����������㣩
                .TextMatrix(0, colסԺ��) = "�����"
                .TextMatrix(0, col����) = "��������"
                .TextMatrix(0, col����Ա����) = "������"
                .TextMatrix(0, col�Ǽ�ʱ��) = "����ʱ��"
                .TextMatrix(0, col��������) = "��������"
                .TextMatrix(0, colʵ������) = "ʵ������"
                .TextMatrix(0, col����ͳ��) = "����ͳ��"
                .TextMatrix(0, colͳ�ﱨ��) = "ͳ�ﱨ��"
                
                .ColWidth(col����) = 0
                .ColWidth(col��������) = 930
                .ColWidth(colʵ������) = 1120
                .ColWidth(col����ͳ��) = 930
                .ColWidth(colͳ�ﱨ��) = 930
            
                '�ı�ĳЩ�е���ʾ˳��
                .ColPosition(col�����ʻ�) = col����Ա����
                .ColPosition(col��������) = col����Ա����
                .ColPosition(colʵ������) = col�Ǽ�ʱ��
                .ColPosition(col����ͳ��) = col�Ǽ�ʱ�� + 1
                .ColPosition(colͳ�ﱨ��) = col�Ǽ�ʱ�� + 1
            Case 3 '-Ԥ��
                .Cols = 15
                .TextMatrix(0, colסԺ��) = "סԺ��"
                .TextMatrix(0, col����) = "����"
                .TextMatrix(0, col����Ա����) = "�տ���"
                .TextMatrix(0, col�Ǽ�ʱ��) = "�տ�ʱ��"
                
                .ColWidth(col����) = 1200
                
                '�ı�ĳЩ�е���ʾ˳��
                .ColPosition(col�����ʻ�) = col����Ա����
        End Select
        For lngCol = 0 To .Cols - 1
            .TextMatrix(1, lngCol) = ""
            .ColAlignmentFixed(lngCol) = 4
        Next
        
        .Col = 0
        .ColSel = .Cols - 1
    End With
End Sub

Private Function FillDetail()
    Dim lngCount As Long, lng����ID As Long
    Dim rsTemp As New ADODB.Recordset
    
    Call SetMenu
    
    If mint���� = 3 Then
        'Ԥ��������
        Exit Function
    End If
    
    '��������Ϣ
    msh��ϸ.Rows = 2
    msh�ֵ�.Rows = 2
    For lngCount = 0 To msh��ϸ.Cols - 1
        msh��ϸ.TextMatrix(1, lngCount) = ""
    Next
    For lngCount = 0 To msh�ֵ�.Cols - 1
        msh�ֵ�.TextMatrix(1, lngCount) = ""
    Next
    
    lng����ID = Val(msh��¼_S.TextMatrix(msh��¼_S.Row, col��¼ID))
    If lng����ID = 0 Then
        Exit Function
    End If
    
    '��ȡ�����¼����ϸ����
    gstrSQL = _
        " Select A.NO,C.���,B.����,B.���,A.���㵥λ as ��λ," & _
        " Ltrim(To_Char(Avg(Nvl(A.����,1)*decode(A.��¼״̬,2,-1,1)*A.����),'999990.000')) as ����, " & _
        " Ltrim(To_Char(Sum(A.��׼����),'999990.000')) as ����, " & _
        " Ltrim(To_Char(Sum(decode(A.��¼״̬,2,-1,1)*A.ʵ�ս��),'999990.00')) as ʵ�ս��, " & _
        " Ltrim(To_Char(Sum(decode(A.��¼״̬,2,-1,1)*A.ͳ����),'999990.00')) as ͳ����, " & _
        IIf(mint���� = 2, " Ltrim(To_Char(Sum(A.���ʽ��),'999990.00')) as ���ʽ��, ", "") & _
        " E.���� as ҽ������,B.�������� as ����," & _
        " Decode(A.��¼״̬,2,'��','��') as ����" & _
        " From ���˷��ü�¼ A,�շ�ϸĿ B,�շ���� C,����֧������ E " & _
        " Where A.�շ�ϸĿID=B.ID and A.�շ����=C.���� " & _
        "       And A.���մ���ID=E.ID(+) And A.����ID=" & lng����ID & _
        " Group by mod(A.��¼����,10),A.NO,Decode(A.�۸񸸺�,NULL,A.���,A.�۸񸸺�),A.��¼״̬,A.�շ����,C.���,B.����,B.���,A.���㵥λ,B.��������,E.����" & _
        " Order by A.NO,Decode(A.�۸񸸺�,NULL,A.���,A.�۸񸸺�)"
    
    Call OpenRecordset(rsTemp, Me.Caption)
    Call ������(msh��ϸ, rsTemp)
    
    If mint���� = 1 Then
        '�շѲ�����
        Exit Function
    End If
    
    '��ȡ�����¼�ķֵ�����
    If mint���� = TYPE_�Ĵ�üɽ Then
        gstrSQL = _
            " Select D.����," & _
            "   Ltrim(To_Char(A.����ͳ����,'999990.00')) as ����ͳ��, " & _
            "   Ltrim(To_Char(A.ͳ�ﱨ�����,'999990.00')) as ͳ�ﱨ��, " & _
            "   Ltrim(To_Char(A.����,'999990.00')) as ���� " & _
            " From ���ս������ A,���ս����¼ B,�����ʻ� C,���շ��õ� D " & _
            " Where B.��¼ID=" & lng����ID & " and B.����=2 And B.����=" & mint���� & _
            "   And B.����ID=C.����ID and C.����=B.���� and D.����=C.���� and D.����=C.���� " & _
            "   And A.����ID=" & lng����ID & "and A.����=D.����(+) "
    Else
        gstrSQL = _
            " Select D.����," & _
            "   Ltrim(To_Char(decode(E.��¼״̬,2,-1,1)*A.����ͳ����,'999990.00')) as ����ͳ��, " & _
            "   Ltrim(To_Char(decode(E.��¼״̬,2,-1,1)*A.ͳ�ﱨ�����,'999990.00')) as ͳ�ﱨ��, " & _
            "   Ltrim(To_Char(A.����,'999990.00')) as ���� " & _
            " From ���ս������ A,���ս����¼ B,�����ʻ� C,���շ��õ� D,���˽��ʼ�¼ E " & _
            " Where E.ID=B.��¼ID And B.��¼ID=" & lng����ID & " and B.����=2 And B.����=" & mint���� & _
            "   And B.����ID=C.����ID and C.����=B.���� and D.����=C.���� and D.����=C.���� " & _
            "   And A.����ID=" & lng����ID & "and A.����=D.����(+) "
    End If
    If rsTemp.State = adStateOpen Then rsTemp.Close
    Call OpenRecordset(rsTemp, Me.Caption)
    Call ������(msh�ֵ�, rsTemp)
End Function

Private Sub InitTable()
    Dim lngCol As Integer
    
    '���ø�ʽ
    With msh��ϸ
        .Rows = 2
        .Cols = 12 'Ϊ������һЩ��������
        .TextMatrix(0, det�շ����) = "�շ����"
        .TextMatrix(0, det�շ�ϸĿ) = "�շ�ϸĿ"
        .TextMatrix(0, det���) = "���"
        .TextMatrix(0, det��λ) = "��λ"
        .TextMatrix(0, det����) = "����"
        .TextMatrix(0, det����) = "����"
        .TextMatrix(0, detʵ�ս��) = "ʵ�ս��"
        .TextMatrix(0, detͳ����) = "ͳ����"
        .TextMatrix(0, detҽ������) = "ҽ������"
        .TextMatrix(0, det��������) = "��������"
        .TextMatrix(0, det����) = "����"
        .TextMatrix(0, det״̬) = "״̬"
        
        .ColWidth(det�շ����) = 600
        .ColWidth(det�շ�ϸĿ) = 1000
        .ColWidth(det���) = 900
        .ColWidth(det��λ) = 600
        .ColWidth(det����) = 800
        .ColWidth(det����) = 800
        .ColWidth(detʵ�ս��) = 930
        .ColWidth(detͳ����) = 930
        .ColWidth(detҽ������) = 800
        .ColWidth(det��������) = 800
        .ColWidth(det����) = 600
        .ColWidth(det״̬) = 600
        For lngCol = 0 To .Cols - 1
            .TextMatrix(1, lngCol) = ""
            .ColAlignmentFixed(lngCol) = 4
        Next
        
        .Col = 0
        .ColSel = .Cols - 1
    End With
    
    With msh�ֵ�
        .Rows = 2
        .Cols = 4 'Ϊ������һЩ��������
        .TextMatrix(0, 0) = "���õ�"
        .TextMatrix(0, 1) = "����ͳ��"
        .TextMatrix(0, 2) = "ͳ�ﱨ��"
        .TextMatrix(0, 3) = "����"
        .ColWidth(0) = 1200
        .ColWidth(1) = 1000
        .ColWidth(2) = 1000
        .ColWidth(3) = 800
        For lngCol = 0 To .Cols - 1
            .TextMatrix(1, lngCol) = ""
            .ColAlignmentFixed(lngCol) = 4
        Next
        
        .Col = 0
        .ColSel = .Cols - 1
    End With
End Sub

Private Sub Ȩ�޿���()
    If InStr(mstrPrivs, "�����") = 0 Then
        mnuFileBatch.Visible = False
        mnuFileDetail.Visible = False
        mnuFileSplitReport.Visible = False
    End If
    If InStr(mstrPrivs, "����") <> 0 Then
        mnuEdit.Visible = True
        mnuFileBill.Visible = True
    Else
        mnuEdit.Visible = False
        mnuFileBill.Visible = False
    End If
    
    If mint���� = TYPE_���������� Or mint���� = TYPE_������ Then
        mnuEditXE.Visible = True
        mnuEditClinic.Visible = mnuEdit.Visible
        mnuEditIn_Patient.Visible = mnuEdit.Visible
        mnuEditDelete.Visible = mnuEdit.Visible
        mnuEditView.Visible = mnuEdit.Visible
        mnuEdit.Visible = True
    End If
    
End Sub

Private Sub SetMenu()
    Dim blnData As Boolean
        
    blnData = (mrs�����¼.RecordCount > 0)
    stbThis.Panels(2).Text = "��ǰ����" & mrs�����¼.RecordCount & "��ҽ���ʻ�"

    tbrThis.Buttons("Print").Enabled = blnData
    tbrThis.Buttons("Preview").Enabled = blnData
    mnuFilePreview.Enabled = blnData
    mnuFilePrint.Enabled = blnData
    mnuFileExcel.Enabled = blnData
    
    mnuFileBatch.Enabled = blnData And (mint���� = 2)
    mnuFileDetail.Enabled = mnuFileBatch.Enabled
End Sub

Public Sub ShowForm(frmParent As Form)
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "select ���,����,nvl(��������,0) as �������� from ������� where nvl(�Ƿ��ֹ,0)<>1 order by ���"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "û�п��ñ�����𣬲���ʹ�ñ����ܡ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If frmҽ������.Visible = True Then
        frmҽ������.Show
        Exit Sub
    End If
    
    Set mcol���� = New Collection
    
    With cmb����
        .Clear
        Do Until rsTemp.EOF
            .AddItem rsTemp("����")
            .ItemData(.NewIndex) = rsTemp("���")
            mcol����.Add Val(rsTemp("��������")), "K" & rsTemp("���")
            If rsTemp("���") = gintInsure Then
                '��ǰҽ����
                'ʹ��API�����Բ�����Click�¼�
                zlControl.CboSetIndex .hwnd, .NewIndex
            End If
            
            rsTemp.MoveNext
        Loop
        If .ListIndex < 0 Then
            'ʹ��API�����Բ�����Click�¼�
            zlControl.CboSetIndex .hwnd, 0
        End If
        
        mint���� = .ItemData(.ListIndex)
        If mint���� = TYPE_������ Then
            mnuFileBalance.Visible = True
            mnuBalance.Visible = True
        End If
    End With
    
    
    frmҽ������.Show , frmParent
End Sub

Private Sub GetSpecialSQL(ByVal intType As Integer, ByVal strBegin As String, ByVal strEnd As String)
    Select Case intType
        Case 1 '-�շ�
            Select Case mint����
            Case TYPE_������
                gstrSQL = _
                    "Select A.����ID,A.NO as ���ݺ�,E.���� as ����,D.����,A.����ID,A.��ʶ�� as �����,Ltrim(A.����) as ����,F.���� as ���,A.�Ա�,A.����,B.���� as ��������," & _
                    "   A.����Ա���� as �շ�Ա,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as �շ�ʱ��,Decode(A.��¼״̬,2,'��','��') as ���˱�־," & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.�����ʻ�֧��),'999990.00')) as �����ʻ�," & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.�������ý��),'999990.00')) as ��������, " & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.ȫ�Ը����),'999990.00')) as ȫ�Է�, " & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.�����Ը����),'999990.00')) as �����Ը�, " & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.����ͳ����),'999990.00')) as ����ͳ��,Decode(C.�����Ը����,14,'��������',11,'��������','��ͨ����') �������,C.��ע ����" & _
                    " From ���˷��ü�¼ A,���ű� B,���ս����¼ C,�����ʻ� D,��������Ŀ¼ E,������Ⱥ F" & _
                    " Where A.��¼���� = 1 And A.����Ա���� IS NOT NULL AND A.��������ID = B.ID(+) And A.�Ǽ�ʱ��>=" & strBegin & " and A.�Ǽ�ʱ��<=" & strEnd & _
                    "       and A.���=1 and A.����ID=C.��¼ID and C.����=1 and C.����=" & mint���� & _
                    "       and A.����ID=D.����ID and D.����=" & mint���� & IIf(mstrCardCond = "", "", " ANd D.ҽ����='" & mstrCardCond & "'") & " And D.����=E.���� and D.����=E.��� and D.����=F.���� and D.��ְ=F.��� " & _
                    " Group by A.����ID,A.NO,E.����,D.����,A.����ID,A.��ʶ��,A.����,A.�Ա�,A.����,B.����,A.����Ա����,A.�Ǽ�ʱ��,A.��¼״̬,F.����,Decode(C.�����Ը����,14,'��������',11,'��������','��ͨ����'),C.��ע" & _
                    " Order by A.�Ǽ�ʱ�� Desc,A.NO Desc"
            Case TYPE_�Ĵ�üɽ
                gstrSQL = _
                    "Select C.��¼ID,H.NO as ���ݺ�,E.���� as ����,D.����,D.����ID,'' as �����,Ltrim(A.����) as ����,F.���� as ���,A.�Ա�,A.����,H.�������� as ��������," & _
                    "   H.������ as ����Ա,To_Char(C.����ʱ��,'YYYY-MM-DD HH24:MI:SS') as ����ʱ��,Decode(sign(C.�������ý��),-1,'��','��') as ���˱�־," & _
                    "   Ltrim(To_Char(Sum(C.�����ʻ�֧��),'999990.00')) as �����ʻ�," & _
                    "   Ltrim(To_Char(Sum(C.�������ý��),'999990.00')) as ��������, " & _
                    "   Ltrim(To_Char(Sum(C.ȫ�Ը����),'999990.00')) as ȫ�Է�, " & _
                    "   Ltrim(To_Char(Sum(C.�����Ը����),'999990.00')) as �����Ը�, " & _
                    "   Ltrim(To_Char(Sum(C.����ͳ����),'999990.00')) as ����ͳ��,Ltrim(To_Char(Sum(C.�����Ը����),'999990.00')) as �����Ը����,G.���� ����,D.����֤�� ְ������" & _
                    " From ������Ϣ A,(Select H.*,B.���� �������� From ���˷��ü�¼ H,���ű� B Where H.��¼���� = 1 And H.����Ա���� IS NOT NULL And H.���=1 AND H.��������ID = B.ID(+) ) H," & _
                    "      ���ս����¼ C,�����ʻ� D,��������Ŀ¼ E,������Ⱥ F,(Select * From ���ղ��� Where ����=" & mint���� & ") G" & _
                    " Where H.����ID(+)=C.��¼ID " & _
                    "       And A.����ID=D.����ID And C.����ID=G.ID(+) And C.����ID=D.����ID And C.����ʱ��>=" & strBegin & " and C.����ʱ��<=" & strEnd & _
                    "       And C.����=1 and C.����=" & mint���� & IIf(mstrCardCond = "", "", " And D.ҽ����='" & mstrCardCond & "'") & " and D.����=C.���� And D.����=E.���� and D.����=E.��� and D.����=F.���� and D.��ְ=F.��� " & _
                    " Group by C.��¼ID,H.NO,E.����,D.����,D.����ID,A.����,A.�Ա�,A.����,H.��������,H.������,C.����ʱ��,Decode(sign(C.�������ý��),-1,'��','��'),F.����,G.����,D.����֤��" & _
                    " Order by C.����ʱ�� Desc"
            Case TYPE_����������, TYPE_������
                'ԭ���̲���:
                 '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
                 "   �ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,
                 '   �������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,
                 '   ����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN,�����ʻ�֧��_IN,"
                 '   ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
                 '������ֵ����Ϊ:
                 '       ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN, _
                 '       dbl�����ʻ����,dblͳ��֧���ۼ�,dbl��������֧��,dbl�����ʻ�֧��,סԺ����_IN,����_IN,dbl���շ�Χ���Ը�,ʵ������_IN
                 '       �������ý��_IN,dbl����ͳ��֧��,dbl����ͳ���Ը�,
                 '       dbl����ͳ��֧��,dbl����ͳ���Ը�,dbl�ǲ�������֧��,�����Ը����_IN,dbl�����ʻ�֧��
                 '       ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
            
                gstrSQL = _
                    "Select A.����ID,A.NO as ���ݺ�,E.���� as ����,D.����,A.����ID,A.��ʶ�� as �����,Ltrim(A.����) as ����,F.���� as ���,A.�Ա�,A.����,B.���� as ��������," & _
                    "   A.����Ա���� as �շ�Ա,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as �շ�ʱ��,Decode(A.��¼״̬,2,'��','��') as ���˱�־," & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.�������ý��),'999990.00')) as ��������, " & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.����),'999990.00')) as ����, " & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.�����ʻ�֧��),'999990.00')) as �����ʻ�," & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.ȫ�Ը����),'999990.00')) as ����ͳ��֧��, " & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.�����Ը����),'999990.00')) as ����ͳ���Ը�, " & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.����ͳ����),'999990.00')) as ����ͳ��֧��, " & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.ͳ�ﱨ�����),'999990.00')) as ����ͳ���Ը�, " & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.�ۼƽ���ͳ��),'999990.00')) as ��������֧��, " & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.���Ը����),'999990.00')) as �ǲ�������֧��, " & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.�ۼ�ͳ�ﱨ��),'999990.00')) as �����ʻ�֧��, " & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.�ⶥ��),'999990.00')) as ���շ�Χ���Ը�," & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.����޶�),'999990.00')) as ����޶�" & _
                    " From ���˷��ü�¼ A,���ű� B,���ս����¼ C,�����ʻ� D,��������Ŀ¼ E,������Ⱥ F" & _
                    " Where A.��¼���� = 1 And A.����Ա���� IS NOT NULL AND A.��������ID = B.ID(+) And A.�Ǽ�ʱ��>=" & strBegin & " and A.�Ǽ�ʱ��<=" & strEnd & _
                    "       and A.���=1 and A.����ID=C.��¼ID and C.����=1 and C.����=" & mint���� & _
                    "       and A.����ID=D.����ID and D.����=" & mint���� & IIf(mstrCardCond = "", "", " And D.ҽ����='" & mstrCardCond & "'") & " And D.����=E.���� and D.����=E.��� and D.����=F.���� and D.��ְ=F.��� " & _
                    " Group by A.����ID,A.NO,E.����,D.����,A.����ID,A.��ʶ��,A.����,A.�Ա�,A.����,B.����,A.����Ա����,A.�Ǽ�ʱ��,A.��¼״̬,F.����" & _
                    " Order by A.�Ǽ�ʱ�� Desc,A.NO Desc"
            Case Else
                gstrSQL = _
                    "Select A.����ID,A.NO as ���ݺ�,E.���� as ����,D.����,A.����ID,A.��ʶ�� as �����,Ltrim(A.����) as ����,F.���� as ���,A.�Ա�,A.����,B.���� as ��������," & _
                    "   A.����Ա���� as �շ�Ա,To_Char(A.�Ǽ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as �շ�ʱ��,Decode(A.��¼״̬,2,'��','��') as ���˱�־," & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.�����ʻ�֧��),'999990.00')) as �����ʻ�," & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.�������ý��),'99999990.00')) as ��������, " & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.ȫ�Ը����),'999990.00')) as ȫ�Է�, " & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.�����Ը����),'999990.00')) as �����Ը�, " & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.����ͳ����),'99999990.00')) as ����ͳ�� " & _
                    " From ���˷��ü�¼ A,���ű� B,���ս����¼ C,�����ʻ� D,��������Ŀ¼ E,������Ⱥ F" & _
                    " Where A.��¼���� = 1 And A.����Ա���� IS NOT NULL AND A.��������ID = B.ID(+) And A.�Ǽ�ʱ��>=" & strBegin & " and A.�Ǽ�ʱ��<=" & strEnd & _
                    "       and A.���=1 and A.����ID=C.��¼ID and C.����=1 and C.����=" & mint���� & _
                    "       and A.����ID=D.����ID and D.����=" & mint���� & IIf(mstrCardCond = "", "", " And D.ҽ����='" & mstrCardCond & "'") & " And D.����=E.���� and D.����=E.��� and D.����=F.���� and D.��ְ=F.��� " & _
                    " Group by A.����ID,A.NO,E.����,D.����,A.����ID,A.��ʶ��,A.����,A.�Ա�,A.����,B.����,A.����Ա����,A.�Ǽ�ʱ��,A.��¼״̬,F.����" & _
                    " Order by A.�Ǽ�ʱ�� Desc,A.NO Desc"
            End Select
        Case 2 '-���㣨����סԺ���㡢����������㣩
            Select Case mint����
            Case TYPE_������
                gstrSQL = _
                    "Select A.ID as ����ID,A.NO as ���ݺ�,E.���� as ����,D.����,A.����ID,B.סԺ��,B.����,F.���� as ���,B.�Ա�,B.����,'' as ����," & _
                    "   A.����Ա���� as ������,To_Char(A.�շ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as ����ʱ��,Decode(A.��¼״̬,2,'��','��') as ���˱�־," & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.�����ʻ�֧��),'9999999990.00')) as �����ʻ�," & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.�������ý��),'999990.00')) as ��������, " & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.ȫ�Ը����),'999990.00')) as ȫ�Է�, " & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.�����Ը����),'999990.00')) as �����Ը�, " & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.����ͳ����),'999990.00')) as ����ͳ��," & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.����),'999990.00')) as ����," & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.ʵ������),'999990.00')) as ʵ������," & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.ͳ�ﱨ�����),'999990.00')) as ͳ�ﱨ��," & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.�����Ը����),'999990.00')) as ���޽��,C.��ע ����" & _
                    " From ���˽��ʼ�¼ A,������Ϣ B,���ս����¼ C,�����ʻ� D,��������Ŀ¼ E,������Ⱥ F " & _
                    " Where A.����ID=C.����ID And A.ID=C.��¼ID  And A.�շ�ʱ��>=" & strBegin & " and A.�շ�ʱ��<=" & strEnd & _
                    "       and C.����=2  and C.����ID=B.����ID and B.����ID=D.����ID and D.����=" & mint���� & IIf(mstrCardCond = "", "", " And D.ҽ����='" & mstrCardCond & "'") & " And D.����=E.���� and D.����=E.��� and D.����=F.���� and D.��ְ=F.��� " & _
                    " Group by A.ID,A.NO,E.����,D.����,A.����ID,B.סԺ��,B.����,B.�Ա�,B.����,A.����Ա����,To_Char(A.�շ�ʱ��,'YYYY-MM-DD HH24:MI:SS'),A.��¼״̬,F.����,C.��ע" & _
                    " Order by To_Char(A.�շ�ʱ��,'YYYY-MM-DD HH24:MI:SS') Desc,A.NO Desc"
            Case TYPE_�Ĵ�üɽ
                gstrSQL = _
                    "Select C.��¼ID,'' as ���ݺ�,E.���� as ����,D.����,D.����ID,'' as �����,Ltrim(A.����) as ����,F.���� as ���,A.�Ա�,A.����,'' as ��������," & _
                    "   C.������ as ����Ա,To_Char(C.����ʱ��,'YYYY-MM-DD HH24:MI:SS') as ����ʱ��,Decode(sign(C.�������ý��),-1,'��','��') as ���˱�־," & _
                    "   Ltrim(To_Char(Sum(C.�����ʻ�֧��),'999990.00')) as �����ʻ�," & _
                    "   Ltrim(To_Char(Sum(C.�������ý��),'999990.00')) as ��������, " & _
                    "   Ltrim(To_Char(Sum(C.ȫ�Ը����),'999990.00')) as ȫ�Է�, " & _
                    "   Ltrim(To_Char(Sum(C.�����Ը����),'999990.00')) as �����Ը�, " & _
                    "   Ltrim(To_Char(Sum(C.����ͳ����),'999990.00')) as ����ͳ��,Ltrim(To_Char(Sum(C.�����Ը����),'999990.00')) as �����Ը����,Ltrim(To_Char(Sum(C.ͳ�ﱨ�����),'999990.00')) ͳ�ﱨ�����,D.����֤�� ְ������" & _
                    " From ������Ϣ A,���ս����¼ C,�����ʻ� D,��������Ŀ¼ E,������Ⱥ F,(Select * From ���ղ��� Where ����=" & mint���� & ") G" & _
                    " Where A.����ID=D.����ID And C.����ID=G.ID(+) And C.����ID=D.����ID And C.����ʱ��>=" & strBegin & " and C.����ʱ��<=" & strEnd & _
                    "       and C.����=2 and C.����=" & mint���� & IIf(mstrCardCond = "", "", " And D.ҽ����='" & mstrCardCond & "'") & " and D.����=C.���� And D.����=E.���� and D.����=E.��� and D.����=F.���� and D.��ְ=F.��� " & _
                    " Group by C.��¼ID,E.����,D.����,D.����ID,A.����,A.�Ա�,A.����,C.������,C.����ʱ��,Decode(sign(C.�������ý��),-1,'��','��'),F.����,D.����֤��" & _
                    " Order by C.����ʱ�� Desc"
            Case TYPE_����������, TYPE_������
                'ԭ���̲���:
                 '   ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN," & _
                 "   �ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,�ۼƽ���ͳ��_IN,�ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,
                 '   �������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,
                 '   ����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN,�����ʻ�֧��_IN,"
                 '   ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
                 '������ֵ����Ϊ:
                 '       ����_IN  ,��¼ID_IN,����_IN,����ID_IN,���_IN, _
                 '       dbl�����ʻ����,dblͳ��֧���ۼ�,dbl��������֧��,dbl�����ʻ�֧��,סԺ����_IN,����_IN,dbl���շ�Χ���Ը�,ʵ������_IN
                 '       �������ý��_IN,dbl����ͳ��֧��,dbl����ͳ���Ը�,
                 '       dbl����ͳ��֧��,dbl����ͳ���Ը�,dbl�ǲ�������֧��,�����Ը����_IN,dbl�����ʻ�֧��
                 '       ֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN
                gstrSQL = _
                    "Select A.ID as ����ID,A.NO as ���ݺ�,E.���� as ����,D.����,A.����ID,B.סԺ��,B.����,F.���� as ���,B.�Ա�,B.����,'' as ����," & _
                    "   A.����Ա���� as ������,To_Char(A.�շ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as ����ʱ��,Decode(A.��¼״̬,2,'��','��') as ���˱�־," & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.�������ý��),'999990.00')) as ��������, " & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.����),'999990.00')) as ����, " & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.�����ʻ�֧��),'999990.00')) as �����ʻ�," & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.ȫ�Ը����),'999990.00')) as ����ͳ��֧��, " & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.�����Ը����),'999990.00')) as ����ͳ���Ը�, " & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.����ͳ����),'999990.00')) as ����ͳ��֧��, " & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.ͳ�ﱨ�����),'999990.00')) as ����ͳ���Ը�, " & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.�ۼƽ���ͳ��),'999990.00')) as ��������֧��, " & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.���Ը����),'999990.00')) as �ǲ�������֧��, " & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.�ۼ�ͳ�ﱨ��),'999990.00')) as �����ʻ�֧��, " & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.�ⶥ��),'999990.00')) as ���շ�Χ���Ը�" & _
                    " From ���˽��ʼ�¼ A,������Ϣ B,���ս����¼ C,�����ʻ� D,��������Ŀ¼ E,������Ⱥ F " & _
                    " Where A.����ID=C.����ID And A.ID=C.��¼ID  And A.�շ�ʱ��>=" & strBegin & " and A.�շ�ʱ��<=" & strEnd & _
                    "       and C.����=2  and C.����ID=B.����ID and B.����ID=D.����ID and D.����=" & mint���� & IIf(mstrCardCond = "", "", " And D.ҽ����='" & mstrCardCond & "'") & " And D.����=E.���� and D.����=E.��� and D.����=F.���� and D.��ְ=F.��� " & _
                    " Group by A.ID,A.NO,E.����,D.����,A.����ID,B.סԺ��,B.����,B.�Ա�,B.����,A.����Ա����,To_Char(A.�շ�ʱ��,'YYYY-MM-DD HH24:MI:SS'),A.��¼״̬,F.����" & _
                    " Order by To_Char(A.�շ�ʱ��,'YYYY-MM-DD HH24:MI:SS') Desc,A.NO Desc"
            Case Else
                gstrSQL = _
                    "Select A.ID as ����ID,A.NO as ���ݺ�,E.���� as ����,D.����,A.����ID,B.סԺ��,B.����,F.���� as ���,B.�Ա�,B.����,'' as ����," & _
                    "   A.����Ա���� as ������,To_Char(A.�շ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as ����ʱ��,Decode(A.��¼״̬,2,'��','��') as ���˱�־," & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.�����ʻ�֧��),'9999999990.00')) as �����ʻ�," & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.�������ý��),'99999990.00')) as ��������, " & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.ȫ�Ը����),'999990.00')) as ȫ�Է�, " & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.�����Ը����),'999990.00')) as �����Ը�, " & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.����ͳ����),'99999990.00')) as ����ͳ��," & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.����),'999990.00')) as ����," & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.ʵ������),'999990.00')) as ʵ������," & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.ͳ�ﱨ�����),'999990.00')) as ͳ�ﱨ��," & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(C.�����Ը����),'999990.00')) as ���޽��" & _
                    " From ���˽��ʼ�¼ A,������Ϣ B,���ս����¼ C,�����ʻ� D,��������Ŀ¼ E,������Ⱥ F " & _
                    " Where A.����ID=C.����ID And A.ID=C.��¼ID  And A.�շ�ʱ��>=" & strBegin & " and A.�շ�ʱ��<=" & strEnd & _
                    "       and C.����=2  and C.����ID=B.����ID and B.����ID=D.����ID and D.����=" & mint���� & IIf(mstrCardCond = "", "", " And D.ҽ����='" & mstrCardCond & "'") & " And D.����=E.���� and D.����=E.��� and D.����=F.���� and D.��ְ=F.��� " & _
                    " Group by A.ID,A.NO,E.����,D.����,A.����ID,B.סԺ��,B.����,B.�Ա�,B.����,A.����Ա����,To_Char(A.�շ�ʱ��,'YYYY-MM-DD HH24:MI:SS'),A.��¼״̬,F.����" & _
                    " Order by To_Char(A.�շ�ʱ��,'YYYY-MM-DD HH24:MI:SS') Desc,A.NO Desc"
            End Select
        Case 3 '-Ԥ��
            Select Case mint����
            Case Else
                gstrSQL = _
                    "Select A.ID as ����ID,A.NO as ���ݺ�,E.���� as ����,D.����,A.����ID,B.סԺ��,B.����,F.���� as ���,B.�Ա�,B.����,C.���� as ����," & _
                    "   A.����Ա���� as �տ���,To_Char(A.�տ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as �տ�ʱ��,Decode(A.��¼״̬,2,'��','��') as ���˱�־," & _
                    "   Ltrim(To_Char(decode(A.��¼״̬,2,-1,1)*Sum(A.���),'9999999990.00')) as �����ʻ�" & _
                    "   From ����Ԥ����¼ A,������Ϣ B,���ű� C,�����ʻ� D,��������Ŀ¼ E,������Ⱥ F" & _
                    " Where A.��¼����=1 And A.����ID=B.����ID And A.����ID=C.ID(+) " & _
                    "       and A.���㷽ʽ='�����ʻ�' and A.�տ�ʱ��>=" & strBegin & " and A.�տ�ʱ��<=" & strEnd & _
                    "       and B.����ID=D.����ID and D.����=" & mint���� & IIf(mstrCardCond = "", "", " And D.ҽ����='" & mstrCardCond & "'") & " And D.����=E.���� and D.����=E.��� and D.����=F.���� and D.��ְ=F.��� " & _
                    " Group by A.ID,A.NO,E.����,D.����,A.����ID,B.סԺ��,B.����,B.�Ա�,B.����,C.����," & _
                    "     A.����Ա����,To_Char(A.�տ�ʱ��,'YYYY-MM-DD HH24:MI:SS'),A.��¼״̬,F.����" & _
                    " Order by �տ�ʱ�� Desc,���ݺ� Desc"
            End Select
    End Select
End Sub
