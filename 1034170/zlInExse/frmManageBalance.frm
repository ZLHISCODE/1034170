VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmManageBalance 
   AutoRedraw      =   -1  'True
   Caption         =   "���˽��ʴ���"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9675
   Icon            =   "frmManageBalance.frx":0000
   KeyPreview      =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imgGray 
      Left            =   1035
      Top             =   1635
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":08CA
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":0AE4
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":0CFE
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":0F18
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":1132
            Key             =   "View"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":18AC
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":1AC6
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":1CE0
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":1EFA
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":2114
            Key             =   "Adjust"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":232E
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":2548
            Key             =   "mzBalance"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":2C42
            Key             =   "zyBalance"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":333C
            Key             =   "RollingCurtain"
            Object.Tag             =   "RollingCurtain"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   270
      Top             =   1965
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":CCD3
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":CEED
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":D107
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":D321
            Key             =   "Del"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":D53B
            Key             =   "View"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":DCB5
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":DECF
            Key             =   "Filter"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":E0E9
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":E303
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":E51D
            Key             =   "Adjust"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":E737
            Key             =   "Modi"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":E951
            Key             =   "mzBalance"
            Object.Tag             =   "mzBalance"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":F04B
            Key             =   "zyBalance"
            Object.Tag             =   "zyBalance"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmManageBalance.frx":F745
            Key             =   "RollingCurtain"
            Object.Tag             =   "RollingCurtain"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picVsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   7410
      MousePointer    =   9  'Size W E
      ScaleHeight     =   1695
      ScaleWidth      =   45
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4170
      Width           =   45
   End
   Begin VB.PictureBox picHsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   15
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   9675
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4140
      Width           =   9675
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   9675
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinHeight1      =   720
      Width1          =   810
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbr 
         Height          =   720
         Left            =   30
         TabIndex        =   5
         Top             =   30
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   15
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
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "MzBalance"
               Object.ToolTipText     =   "����������ʴ���"
               Object.Tag             =   "����"
               ImageKey        =   "mzBalance"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "סԺ"
               Key             =   "ZyBalance"
               Description     =   "����"
               Object.ToolTipText     =   "����סԺ���ʴ���"
               Object.Tag             =   "סԺ"
               ImageKey        =   "zyBalance"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Del"
               Description     =   "����"
               Object.ToolTipText     =   "����ǰѡ�е�������"
               Object.Tag             =   "����"
               ImageKey        =   "Del"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "View"
               Description     =   "����"
               Object.ToolTipText     =   "���ĵ�ǰ���ݵ�����"
               Object.Tag             =   "����"
               ImageKey        =   "View"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Filter"
               Description     =   "����"
               Object.ToolTipText     =   "��������������ɸѡ��¼"
               Object.Tag             =   "����"
               ImageKey        =   "Filter"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��λ"
               Key             =   "Go"
               Description     =   "��λ"
               Object.ToolTipText     =   "��λ�����������ļ�¼��"
               Object.Tag             =   "��λ"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "�շ�����"
               Object.Tag             =   "����"
               ImageKey        =   "RollingCurtain"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "SplitRollingCurtain"
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "����"
               Object.ToolTipText     =   "��ǰ��������"
               Object.Tag             =   "����"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Quit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   5844
      Width           =   9672
      _ExtentX        =   17066
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmManageBalance.frx":FE3F
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11986
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshMoney 
      Height          =   1665
      Left            =   7470
      TabIndex        =   2
      Top             =   4185
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   2937
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
      MouseIcon       =   "frmManageBalance.frx":106D3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   1665
      Left            =   0
      TabIndex        =   1
      Top             =   4185
      Width           =   7380
      _ExtentX        =   13018
      _ExtentY        =   2937
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
      MouseIcon       =   "frmManageBalance.frx":109ED
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   3315
      Left            =   90
      TabIndex        =   0
      Top             =   825
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   5847
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
      MouseIcon       =   "frmManageBalance.frx":10D07
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFile_PrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFile_PreView 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFile_Print 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFile_Excel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileMoneyEnum 
         Caption         =   "�ֽ�㳮(&E)"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRollingCurtain 
         Caption         =   "�շ�����(&M)"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuFileRollingCurtainSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileInsure 
         Caption         =   "�������(&I)"
      End
      Begin VB.Menu mnuFileLocalSet 
         Caption         =   "��������(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileLocalSet_ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_quit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEdit_MzBalance 
         Caption         =   "���ﲡ�˽���(&M)"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEdit_ZyBalance 
         Caption         =   "סԺ���˽���(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEdit_BalanceBat 
         Caption         =   "������;����(&T)"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuEdit_BalanceUnit 
         Caption         =   "��Լ��λ����(&U)"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuEdit_Due 
         Caption         =   "Ӧ�տ����(&Y)"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuEdit_RefundDeposit 
         Caption         =   "����˿�(&R)"
      End
      Begin VB.Menu mnuEditSplitMZ 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditMzToZy 
         Caption         =   "�������תסԺ(&Z)"
      End
      Begin VB.Menu mnuEditmzXZ 
         Caption         =   "תסԺ��������(&X)"
      End
      Begin VB.Menu mnuEditYbVerfy 
         Caption         =   "ҽ��У��(&C)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEdit_0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Del 
         Caption         =   "��������(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit_View 
         Caption         =   "���ĵ���(&V)"
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdit_Print 
         Caption         =   "�ش����Ʊ��(&R)"
      End
      Begin VB.Menu mnuEdit_Print_Supplemental 
         Caption         =   "�������Ʊ��(&B)"
      End
      Begin VB.Menu mnuEdit_PrintDetail 
         Caption         =   "��ӡ������ϸ(&L)"
      End
      Begin VB.Menu mnuEditPatiPrint 
         Caption         =   "�����˲������Ʊ��(&P)"
      End
      Begin VB.Menu mnuEditSplitWriteCard 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditWriteCard 
         Caption         =   "������Ϣд��(&W)"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "����(&R)"
      Visible         =   0   'False
      Begin VB.Menu mnuReportItem 
         Caption         =   "-"
         Index           =   0
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
         Begin VB.Menu mnuView_Tlb_1 
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
      Begin VB.Menu mnuView_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFilter 
         Caption         =   "����(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewGo 
         Caption         =   "��λ(&G)"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuView_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefeshOption 
         Caption         =   "ˢ�·�ʽ(&O)"
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "������Ҫˢ������(&1)"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "��������ʾ�Ƿ�ˢ��(&2)"
            Index           =   1
         End
         Begin VB.Menu mnuViewRefeshOptionItem 
            Caption         =   "�������Զ�ˢ������(&3)"
            Index           =   2
         End
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewreFlash 
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
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "������̳(&F)"
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
Attribute VB_Name = "frmManageBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mstrPrivs As String
Private mlngModul As Long
Private mrsList As ADODB.Recordset  '�����б�
Private mrsDetail As ADODB.Recordset
Private mrsMoney As ADODB.Recordset
Private Type Type_SQLCondition
    Default As Boolean          '�Ƿ���ȱʡ���룬��ʱû������ֵ,ȱʡֵ��mstrFilter��
    DateB As Date
    DateE As Date
    NOB As String
    NOE As String
    FactB As String
    FactE As String
    InPatientID As String
    OutExseID As String
    Patient As String
    Operator As String
    Flag As Byte
    str��Դ As String   '0000:����;סԺ;���;����
End Type
Private SQLCondition As Type_SQLCondition

Private mstrFilter As String
Private mblnMax As Boolean
Private mblnGo As Boolean, mlngGo As Long
Private mlngCurRow As Long, mlngTopRow As Long
Private mblnNOMoved As Boolean '��¼��ǰѡ��ĵ����Ƿ����ں����ݱ���
Private mbytType As Byte    '��¼����1-���ʼ�¼,2-�����¼�¼,3-����ԭ��¼
Private mobjInPati As Object
Private mbln��������  As Boolean
Private mstrWriteCardTypeIDs As String   '��ǰ���������п����ID
Private mstrPrivs_RollingCurtain As String  '�շ����ʹ���Ȩ��
Private mblnҽ��У�� As Boolean 'ҽ����ҪУ��

Private Sub Form_Activate()
    Call InitLocPar(mlngModul)
    Call mshList_GotFocus
End Sub

Private Sub mnuEdit_BalanceBat_Click()
    gblnOK = False
    frmBalanceBat.Show GetModuleType, Me
        
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����ѱ仯,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
        End If
    End If
End Sub

Private Sub mnuEdit_BalanceUnit_Click()
    If frmBalanceUnit.ShowMe(Me, 0, 0, False) = False Then Exit Sub
    If mnuViewRefeshOptionItem(1).Checked Then
        If MsgBox("��ǰ�����ѱ仯,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            mnuViewReFlash_Click
        End If
    ElseIf mnuViewRefeshOptionItem(2).Checked Then
        mnuViewReFlash_Click
    End If
End Sub

Private Sub mnuEdit_Due_Click()
    frmManageDue.mstrPrivs = mstrPrivs
    frmManageDue.mlngModul = mlngModul
    frmManageDue.Show 0, Me
End Sub



Private Sub mnuEdit_PrintDetail_Click()
    Dim strNo As String
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    If strNo = "" Then
        MsgBox "��ǰû�е��ݿ��Դ�ӡ֤����", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�Ƿ���ת������ݱ���
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNo, 7, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '��ʱ��ת���������ݱ�
    End If
    
    Call PrintDetail
End Sub

Private Sub mnuEdit_RefundDeposit_Click()
'---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����˿�
'---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objFun          As Object
    
    Err = 0: On Error Resume Next
    Set objFun = CreateObject("zl9Patient.clsPatient")
    If Err <> 0 Then Exit Sub
    
    If objFun.RefundDeposit(glngSys, gcnOracle, Me, gstrDBUser) = False Then
        Set objFun = Nothing
        Exit Sub
    End If
    Set objFun = Nothing
End Sub

Private Sub mnuEditMzToZy_Click()
    '�������תסԺ����:33635
    If InStr(1, mstrPrivs, ";�������תסԺ;") = 0 Then Exit Sub
    mnuEditMzToZy.Visible = InStr(1, mstrPrivs, ";�������תסԺ;") > 0
    If mobjInPati Is Nothing Then
        Err = 0: On Error Resume Next
        Set mobjInPati = CreateObject("zl9InPatient.clsInPatient")
        
        If Err <> 0 Then
            MsgBox "ע��:" & vbCrLf & "   סԺ���˲���(zl9InPatient)����ʧ��,����ϵͳ����Ա��ϵ!"
            Exit Sub
        End If
    End If
    'zlOutFeeToInFee(
    '   ByVal frmMain As Object, ByVal cnMain As ADODB.Connection, _
    '   ByVal lngSys As Long, ByVal lngModule As Long, ByVal strPrivs As String, strDBUser As String, _
    '   ByVal lng����ID As Long, intPatientRange As Integer)
    Call mobjInPati.zlOutFeeToInFee(Me, gcnOracle, glngSys, mlngModul, mstrPrivs, gstrDBUser, 0, 0)
End Sub

Private Sub mnuEditmzXZ_Click()
    If InStr(mstrPrivs, ";תסԺ��������;") = 0 Or mbln�������� Then Exit Sub
    If frmFeeRefundment.zlShowEdit(Me, 2, mlngModul, mstrPrivs) = False Then
        Exit Sub
    End If
End Sub

Private Sub mnuEditPatiPrint_Click()
    '�����˲���Ʊ��:56283
    If frmMakeupPrintBill.zlRePrintBill(Me, mlngModul, mstrPrivs) = False Then Exit Sub
    If mnuViewRefeshOptionItem(1).Checked Then
        If MsgBox("��ǰ�����ѱ仯,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            mnuViewReFlash_Click
        End If
    ElseIf mnuViewRefeshOptionItem(2).Checked Then
        mnuViewReFlash_Click
    End If
End Sub

Private Sub mnuEditWriteCard_Click()
    Dim lngCardTypeID As Long, strExpend As String, lng����ID As Long
    Dim lng����ID As Long, strNo As String
    Dim bytFunc As Byte
    
    With mshList
        strNo = .TextMatrix(.Row, GetColNum("���ݺ�"))
        lng����ID = Val(.TextMatrix(.Row, GetColNum("����ID")))
        lng����ID = Val(.TextMatrix(.Row, GetColNum("����ID")))
        bytFunc = IIf(Val(.TextMatrix(.Row, GetColNum("��־"))) = 1, 0, 1)
    End With
    '����:��סԺ��Ϣд�뿨��
    '����:56615
    If mstrWriteCardTypeIDs = "" Then Exit Sub
    If bytFunc = 0 Then '������ʷ���
        If InStr(mstrPrivs, ";������Ϣд��;") = 0 Then Exit Sub
    Else
        If InStr(mstrPrivs, ";סԺ��Ϣд��;") = 0 Then Exit Sub
    End If
    
    If strNo = "" Then
        MsgBox "��ǰû�е��ݿ�������д����", vbInformation, gstrSysName
        Exit Sub
    End If

    If InStr(1, mstrWriteCardTypeIDs, ",") = 0 Then lngCardTypeID = Val(mstrWriteCardTypeIDs)
    Call WriteInforToCard(Me, mlngModul, mstrPrivs, gobjSquare.objSquareCard, lngCardTypeID, bytFunc, lng����ID, lng����ID)
End Sub
Private Function IsYbBalanceCheck(ByVal lng����ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ƿ�ҽ������У��
    '����:���˺�
    '����:2015-05-07 16:18:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    On Error GoTo errHandle
    strSql = "" & _
    "   Select ���㷽ʽ,��� From ���ս�����ϸ " & _
    "   Where ����id = [1] And ���㷽ʽ<>'�ֽ�' " & _
    "           And ��־=1 and Rownum <2"  'ҽ���ܿصĹ��̶̹�д����һ��"�ֽ�"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "���ս������", lng����ID)
    IsYbBalanceCheck = Not rsTmp.EOF
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub mnuEditYbVerfy_Click()
    Dim strNo As String, lng����ID As Long
    Dim int��¼״̬ As Long, str����Ա���� As String
    Dim blnYb As Boolean, blnThreeDeposit As Boolean
    Dim lng����ID As Long, intԤ����� As Integer
    
    With mshList
        lng����ID = Val(.TextMatrix(.Row, GetColNum("����ID")))
        strNo = .TextMatrix(.Row, GetColNum("���ݺ�"))
        int��¼״̬ = Val(.TextMatrix(.Row, GetColNum("��¼״̬")))
        blnYb = .TextMatrix(.Row, GetColNum("ҽ��")) <> ""
        str����Ա���� = .TextMatrix(.Row, GetColNum("����Ա"))
    End With
    If lng����ID = 0 Then
        MsgBox "������У�ԵĽ��ʵ�!", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    If Not blnYb Then
        MsgBox "��ǰ���ʵ�����ҽ�����㵥�ݣ�������ҽ��У�Ե����!", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    If int��¼״̬ <> 1 Then
        MsgBox "���ϵĵ��ݣ�������ҽ��У��!", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    If str����Ա���� <> UserInfo.���� Then
        MsgBox "��ǰ���ʵ��ǲ���Ա��" & str����Ա���� & "�������ĵ��ݣ�ֻ�ܶ��Լ��ĵ��ݽ���ҽ��У��!", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    If IsYbBalanceCheck(lng����ID) = False Then
        MsgBox "��ǰ���ʵ����ý���У�Բ���!", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    
    If frmMedicareReckoning.ShowMeFromOut(Me, mstrPrivs, lng����ID, blnThreeDeposit, lng����ID, intԤ�����) = False Then Exit Sub
    
    If blnThreeDeposit Then
        frmBalanceDeposit.ShowMe Me, mlngModul, lng����ID, lng����ID, True, False, "", "", intԤ�����
    End If
    
    If mnuViewRefeshOptionItem(1).Checked Then
      If MsgBox("��ǰ�����ѱ仯,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
          mnuViewReFlash_Click
       End If
    ElseIf mnuViewRefeshOptionItem(2).Checked Then
       mnuViewReFlash_Click
    End If
End Sub

Private Sub mnuFileInsure_Click()
    gclsInsure.InsureSupport
End Sub

Private Sub mnuFileLocalSet_Click()
    frmSetExpence.mlngModul = mlngModul
    frmSetExpence.mstrPrivs = mstrPrivs
    frmSetExpence.mbytInFun = 1
    frmSetExpence.Show 1, Me
End Sub

Private Sub mnuFileMoneyEnum_Click()
    Call frmMoneyEnum.ShowMe(Me)
End Sub
 

Private Sub mnuFileRollingCurtain_Click()
   Call zlExecuteChargeRollingCurtain(Me)
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Dim strNo As String
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    If strNo = "" Then
        Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me)
    Else
        With mshList
            Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
                "����ID=" & .TextMatrix(.Row, GetColNum("����ID")), _
                "סԺ��=" & .TextMatrix(.Row, GetColNum("סԺ��")), _
                "����ID=" & .TextMatrix(.Row, GetColNum("����ID")), _
                "NO=" & strNo, _
                "��¼״̬=" & mbytType)
        End With
    End If
End Sub

Private Sub mnuViewFilter_Click()
    frmBalanceFilter.Show 1, Me
    If gblnOK Then
        With frmBalanceFilter
            mstrFilter = .mstrFilter
            
            SQLCondition.Default = False
            SQLCondition.DateB = .dtpBegin.Value
            SQLCondition.DateE = .dtpEnd.Value
            SQLCondition.NOB = .txtNOBegin.Text
            SQLCondition.NOE = .txtNoEnd.Text
            SQLCondition.FactB = .txtFactBegin.Text
            SQLCondition.FactE = .txtFactEnd.Text
            SQLCondition.InPatientID = Trim(.txtסԺ��.Text)
            SQLCondition.Patient = gstrLike & UCase(.txt����.Text) & "%"
            SQLCondition.Operator = NeedName(.cbo����Ա.Text)
            SQLCondition.OutExseID = Trim(.txtClinic.Text)
            SQLCondition.str��Դ = .mstr��Դ
            
            If .chkType(0).Value = 1 And .chkType(1).Value = 1 Then
                SQLCondition.Flag = 0
            ElseIf .chkType(0).Value = 1 Then
                SQLCondition.Flag = 1
            ElseIf .chkType(1).Value = 1 Then
                SQLCondition.Flag = 2
            End If
        End With
        
        mnuViewReFlash_Click
    End If
End Sub

Private Sub mshDetail_EnterCell()
    mshDetail.ForeColorSel = mshDetail.CellForeColor
End Sub

Private Sub mshDetail_GotFocus()
    Call SetActiveList(mshDetail)
End Sub

Private Sub mshList_DblClick()
    If mshList.MouseRow = 0 Then Exit Sub
    If mnuEdit_View.Enabled Then mnuEdit_View_Click
End Sub

Private Sub mshList_EnterCell()
    Dim lng����ID As Long
    Dim strNo As String, int��Դ As Integer
    Dim blnYb As Boolean, blnMzBalance As Boolean, blnZyBalance As Boolean
    Dim bln��ͨ���� As Boolean, blnҽ������ As Boolean
    Dim blnУ�� As Boolean, bytFunc As Byte
    
    lng����ID = Val(mshList.TextMatrix(mshList.Row, GetColNum("����ID")))
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    mbytType = Val(mshList.TextMatrix(mshList.Row, GetColNum("��¼״̬")))
    blnYb = mshList.TextMatrix(mshList.Row, GetColNum("ҽ��")) <> ""
    blnУ�� = Trim(mshList.TextMatrix(mshList.Row, GetColNum("У��"))) <> ""
    bytFunc = IIf(Val(mshList.TextMatrix(mshList.Row, GetColNum("��־"))) = 1, 0, 1)
    
    bln��ͨ���� = InStr(mstrPrivs, ";��ͨ���˽���;") > 0
    blnҽ������ = InStr(mstrPrivs, ";���ս���;") > 0
    blnMzBalance = InStr(1, mstrPrivs, ";������ý���;") > 0 And blnҽ������
    blnZyBalance = InStr(1, mstrPrivs, ";סԺ���ý���;") > 0 And blnҽ������
    
    mnuEditYbVerfy.Visible = False
    If blnYb And lng����ID <> 0 And mbytType = 1 Then
        blnYb = blnУ��
        mnuEditYbVerfy.Visible = blnYb And (blnMzBalance Or blnZyBalance)
    End If
    
    If mshList.Row = 0 Or lng����ID = 0 Then
        mnuEdit_PrintDetail.Enabled = False
        mnuEdit_Print_Supplemental.Enabled = False
        mnuEdit_Print.Enabled = False
        mnuEdit_Del.Enabled = False
        tbr.Buttons("Del").Enabled = False
        mnuEditWriteCard.Enabled = False
        Exit Sub
    End If
    
    mnuEditWriteCard.Enabled = (bytFunc = 0 And InStr(mstrPrivs, ";������Ϣд��;") > 0) _
                            Or (bytFunc = 1 And InStr(mstrPrivs, ";סԺ��Ϣд��;") > 0)
    int��Դ = Val(mshList.TextMatrix(mshList.Row, GetColNum("��־")))
    
    mlngGo = mshList.Row
    mlngCurRow = mshList.Row: mlngTopRow = mshList.TopRow
    
    Call ShowDetail(lng����ID, , strNo, int��Դ)
    Call ShowMoney(lng����ID, , strNo)
            
    mnuEdit_PrintDetail.Enabled = mbytType = 1
    mnuEdit_Print_Supplemental.Enabled = mbytType = 1 And Trim(mshList.TextMatrix(mshList.Row, GetColNum("Ʊ�ݺ�"))) = ""
    mnuEdit_Print.Enabled = mbytType = 1
    mnuEdit_Del.Enabled = mbytType = 1 And Not blnУ��
    tbr.Buttons("Del").Enabled = mbytType = 1 And Not blnУ��
    
    mshList.ForeColorSel = mshList.CellForeColor
End Sub

Private Sub mshList_GotFocus()
    Call SetActiveList(mshList)
End Sub

Private Sub SetActiveList(obj As Object)
    If obj Is mshList Then
        mshList.BackColorSel = &HC0C0C0
        mshDetail.BackColorSel = &HE0E0E0
        mshMoney.BackColorSel = &HE0E0E0
    ElseIf obj Is mshDetail Then
        mshList.BackColorSel = &HE0E0E0
        mshDetail.BackColorSel = &HC0C0C0
        mshMoney.BackColorSel = &HE0E0E0
    ElseIf obj Is mshMoney Then
        mshList.BackColorSel = &HE0E0E0
        mshDetail.BackColorSel = &HE0E0E0
        mshMoney.BackColorSel = &HC0C0C0
    End If
End Sub

Private Sub mshList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete And mnuEdit_Del.Enabled And mnuEdit_Del.Visible Then Call mnuEdit_Del_Click
End Sub

Private Sub mshList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuEdit, 2
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3
            'ʼ�մӵ�ǰ�п�ʼ
            If mnuViewGo.Enabled Then Call SeekBill(False)
        Case vbKeyReturn
            If mnuEdit_View.Enabled Then mnuEdit_View_Click
        Case vbKeyEscape
            mblnGo = False
    End Select
End Sub

Private Sub mnuEdit_Del_Click()
    Dim strNo As String
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    
    If strNo = "" Then
        MsgBox "��ǰû�е��ݿ������ϣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    frmBalance.mlngModul = mlngModul
    frmBalance.mstrPrivs = mstrPrivs
    frmBalance.mbytInState = 0
    frmBalance.mstrInNO = strNo
    frmBalance.Show GetModuleType, Me
        
    If mnuViewRefeshOptionItem(1).Checked Then
        If MsgBox("��ǰ�����Ѹ��ĵ����嵥����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            mnuViewReFlash_Click
        End If
    ElseIf mnuViewRefeshOptionItem(2).Checked Then
        mnuViewReFlash_Click
    End If
End Sub

Private Sub mnuHelpTitle_Click()
ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub mnuEdit_zyBalance_Click()
    On Error Resume Next
    Err.Clear
    
    frmBalance.mlngModul = mlngModul
    frmBalance.mstrPrivs = mstrPrivs
    frmBalance.mbytInState = 0
    frmBalance.mbytFunc = 1
    frmBalance.Show GetModuleType, Me
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����ѱ仯,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
        End If
    End If
End Sub
Private Sub mnuEdit_MzBalance_Click()
    On Error Resume Next
    Err.Clear
    frmBalance.mlngModul = mlngModul
    frmBalance.mstrPrivs = mstrPrivs
    frmBalance.mbytInState = 0
    frmBalance.mbytFunc = 0
    frmBalance.Show GetModuleType, Me
    If gblnOK Then
        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����ѱ仯,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
        End If
    End If
End Sub

Private Sub mnuEdit_View_Click()
    Dim strNo As String, lngPatientID As Long, lng����ID As Long
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    If strNo = "" Then
        MsgBox "��ǰû�е��ݿ��Բ��ģ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    On Error Resume Next
    Err.Clear
    
    lngPatientID = Val(mshList.TextMatrix(mshList.Row, GetColNum("����ID")))
    lng����ID = Val(mshList.TextMatrix(mshList.Row, GetColNum("����ID")))
    If lngPatientID = 0 Then
      '��ʾ��������
        Call frmBalanceUnit.ShowMe(Me, 1, lng����ID, IIf(mbytType = 2, True, False), mblnNOMoved)
    Else
        '��ʾ��������
        frmBalance.mlngModul = mlngModul
        frmBalance.mstrPrivs = mstrPrivs
        frmBalance.mbytInState = 1
        frmBalance.mblnViewCancel = IIf(mbytType = 2, True, False)
        frmBalance.mstrInNO = strNo
        frmBalance.mblnNOMoved = mblnNOMoved
        frmBalance.mlngBillID = Val(mshList.TextMatrix(mshList.Row, 0))
        frmBalance.Show GetModuleType, Me
    End If
End Sub

Private Sub mnuFile_quit_Click()
    Unload Me
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuViewReFlash_Click()
    ShowBills mstrFilter
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
    Dim i As Long
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For i = 1 To tbr.Buttons.Count
        tbr.Buttons(i).Caption = IIf(mnuViewToolText.Checked, tbr.Buttons(i).Tag, "")
    Next
    cbr.Bands(1).minHeight = tbr.ButtonHeight
    Form_Resize
End Sub

Private Sub mshMoney_EnterCell()
    mshMoney.ForeColorSel = mshMoney.CellForeColor
End Sub

Private Sub mshMoney_GotFocus()
    Call SetActiveList(mshMoney)
End Sub

Private Sub picHsc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If mshList.Height + Y < 1000 Or mshDetail.Height - Y < 1000 Then Exit Sub
        picHsc.Top = picHsc.Top + Y
        mshList.Height = mshList.Height + Y
        mshDetail.Top = mshDetail.Top + Y
        mshDetail.Height = mshDetail.Height - Y
        picVsc.Top = picVsc.Top + Y
        picVsc.Height = picVsc.Height - Y
        mshMoney.Top = mshMoney.Top + Y
        mshMoney.Height = mshMoney.Height - Y
        Me.Refresh
    End If
End Sub

Private Sub picHsc_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then mshList.SetFocus
End Sub

Private Sub picVsc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If mshDetail.Width + X < 1000 Or mshMoney.Width - X < 1000 Then Exit Sub
        picVsc.Left = picVsc.Left + X
        mshDetail.Width = mshDetail.Width + X
        mshMoney.Left = mshMoney.Left + X
        mshMoney.Width = mshMoney.Width - X
        Me.Refresh
    End If
End Sub

Private Sub picVsc_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then mshList.SetFocus
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Quit"
            mnuFile_quit_Click
        Case "Go" '��λ
            mnuViewGo_Click
        Case "Filter" '����
            mnuViewFilter_Click
        Case "View"
            mnuEdit_View_Click
        Case "ZyBalance"
            mnuEdit_zyBalance_Click
        Case "MzBalance"
            mnuEdit_MzBalance_Click
        Case "Del"
            mnuEdit_Del_Click
        Case "Print"
            mnuFile_Print_Click
        Case "Preview"
            mnuFile_PreView_Click
        Case "����"
            mnuFileRollingCurtain_Click
        Case "Help"
            mnuHelpTitle_Click
    End Select
End Sub

Private Sub tbr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool, 2
End Sub

Private Sub mnuFile_Excel_Click()
    Call OutputList(3)
End Sub

Private Sub mnuFile_PreView_Click()
    Call OutputList(2)
End Sub

Private Sub mnuFile_Print_Click()
    Call OutputList(1)
End Sub

Private Sub mnuFile_PrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub OutputList(bytStyle As Byte)
'���ܣ�������б�
'������bytStyle=1-��ӡ,2-Ԥ��,3-�����Excel
    Dim objOut As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim bytR As Byte, intRow As Integer
    
    intRow = mshList.Row
    
    '��ͷ
    objOut.Title.Text = "���˽��ʵ����嵥"
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '����
    With frmBalanceFilter
        objRow.Add "ʱ�䣺" & Format(.dtpBegin.Value, .dtpBegin.CustomFormat) & " �� " & Format(.dtpEnd.Value, .dtpEnd.CustomFormat)
        objRow.Add "���ʣ�" & IIf(mbytType = 2, "���ϵ���", "���ʵ���")
        objOut.UnderAppRows.Add objRow
    End With
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��")
    objOut.BelowAppRows.Add objRow
    
    '����
    mshList.Redraw = False
    Set objOut.Body = mshList
    
    '���
    If bytStyle = 1 Then
        bytR = zlPrintAsk(objOut)
        Me.Refresh
        If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    Else
        zlPrintOrView1Grd objOut, bytStyle
    End If
    
    mshList.Row = intRow
    mshList.Col = 0: mshList.ColSel = mshList.Cols - 1
    mshList.Redraw = True
End Sub

Private Sub mnuHelpWebHome_Click()
    zlHomePage hWnd
End Sub

Private Sub mnuHelpWebMail_Click()
    zlMailTo hWnd
End Sub

Private Sub SetMenu(blnUsed As Boolean)
'���ܣ��������޼�¼���ò˵�����״̬
    mnuFile_Print.Enabled = blnUsed
    mnuFile_PreView.Enabled = blnUsed
    mnuFile_Excel.Enabled = blnUsed
    tbr.Buttons("Print").Enabled = blnUsed
    tbr.Buttons("Preview").Enabled = blnUsed
    
    mnuEdit_Del.Enabled = blnUsed
    mnuEdit_View.Enabled = blnUsed
    mnuEdit_Print.Enabled = blnUsed
    mnuEdit_PrintDetail.Enabled = blnUsed
    mnuEdit_Print_Supplemental.Enabled = blnUsed
    tbr.Buttons("Del").Enabled = blnUsed
    tbr.Buttons("View").Enabled = blnUsed

    mnuViewGo.Enabled = blnUsed
    tbr.Buttons("Go").Enabled = blnUsed
End Sub
Private Sub Ȩ�޿���()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:Ȩ�޿���
    '����:���˺�
    '����:2011-09-20 23:27:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnMzBalance As Boolean, blnZyBalance As Boolean '�������,סԺ����
    Dim bln��ͨ���� As Boolean, blnҽ������ As Boolean, bln���� As Boolean
    Dim blnHavePrivs As Boolean
    
    On Error GoTo errHandle
    
    mnuEditMzToZy.Visible = InStr(1, mstrPrivs, ";�������תסԺ;") > 0 '33635
    mnuEditmzXZ.Visible = InStr(mstrPrivs, ";תסԺ��������;") > 0 And Not mbln��������
    mnuEditSplitMZ.Visible = InStr(mstrPrivs, ";תסԺ��������;") > 0 And Not mbln�������� Or InStr(1, mstrPrivs, ";�������תסԺ;") > 0
    
    bln��ͨ���� = InStr(mstrPrivs, ";��ͨ���˽���;") > 0
    blnҽ������ = InStr(mstrPrivs, ";���ս���;") > 0
    blnMzBalance = InStr(1, mstrPrivs, ";������ý���;") > 0 And (bln��ͨ���� Or blnҽ������)
    mnuEdit_MzBalance.Visible = blnMzBalance
    mnuEdit_BalanceUnit.Visible = blnMzBalance '��Լ��λ����
    tbr.Buttons("MzBalance").Visible = blnMzBalance
        
    blnZyBalance = InStr(1, mstrPrivs, ";סԺ���ý���;") > 0 And (bln��ͨ���� Or blnҽ������)
    mnuEdit_ZyBalance.Visible = blnZyBalance
    tbr.Buttons("ZyBalance").Visible = blnZyBalance
    
    mnuEdit_Print.Visible = (blnZyBalance Or blnMzBalance) And InStr(mstrPrivs, ";�ش�Ʊ��;") > 0 '�ش�Ʊ��
    '52328
    mnuEdit_Print_Supplemental.Visible = (blnZyBalance Or blnMzBalance) And InStr(mstrPrivs, ";����Ʊ��;") > 0        '����Ʊ��
    '����:56283
    mnuEditPatiPrint.Visible = (blnZyBalance Or blnMzBalance) And InStr(mstrPrivs, ";����Ʊ��;") > 0        '����Ʊ��
    
    mnuEdit_Due.Visible = InStr(mstrPrivs, ";Ӧ�տ����;") > 0
    
    mnuEdit_BalanceBat.Visible = InStr(mstrPrivs, ";������;����;") > 0
    
    '���ʷָ�
    mnuEdit_0.Visible = blnMzBalance Or blnZyBalance Or InStr(mstrPrivs, ";Ӧ�տ����;") > 0 Or InStr(mstrPrivs, ";������;����;") > 0
    
    bln���� = InStr(mstrPrivs, ";��������;") > 0 And (bln��ͨ���� Or blnҽ������)
    mnuEdit_Del.Visible = bln����
    tbr.Buttons("Del").Visible = bln����
    
    '�շ����ʹ���
    blnHavePrivs = InStr(mstrPrivs_RollingCurtain, ";����;") > 0
    mnuFileRollingCurtain.Visible = blnHavePrivs
    mnuFileRollingCurtainSplit.Visible = blnHavePrivs
    tbr.Buttons("����").Visible = blnHavePrivs
    tbr.Buttons("SplitRollingCurtain").Visible = blnHavePrivs
    
    mnuEditSplitWriteCard.Visible = (InStr(mstrPrivs, ";סԺ��Ϣд��;") > 0 Or InStr(mstrPrivs, ";������Ϣд��;") > 0) _
                                    And mstrWriteCardTypeIDs <> ""
    mnuEditWriteCard.Visible = (InStr(mstrPrivs, ";סԺ��Ϣд��;") > 0 Or InStr(mstrPrivs, ";������Ϣд��;") > 0) _
                                And mstrWriteCardTypeIDs <> ""
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub
Private Sub Form_Load()
    Dim i As Long
    
    mstrPrivs = gstrPrivs
    mlngModul = glngModul
    mstrPrivs_RollingCurtain = ";" & GetPrivFunc(glngSys, 1506) & ";"
    
    Call RestoreWinState(Me, App.ProductName)
    mbln�������� = Val(zlDatabase.GetPara("����ת�������˷�", glngSys, 1131)) = 1
    mstrWriteCardTypeIDs = ""
    If Not gobjSquare Is Nothing Then
        If Not gobjSquare.objSquareCard Is Nothing Then
            mstrWriteCardTypeIDs = gobjSquare.objSquareCard.zlGetAvailabilityWriteCardType
        End If
    End If
    
    'ˢ�·�ʽ
    For i = 0 To mnuViewRefeshOptionItem.UBound
        If i = Val(zlDatabase.GetPara("ˢ�·�ʽ", glngSys, mlngModul, 2)) Then
            mnuViewRefeshOptionItem(i).Checked = True
        Else
            mnuViewRefeshOptionItem(i).Checked = False
        End If
    Next
    '���������˰�ش�ӡ����
    If gobjTax Is Nothing Then
        On Error Resume Next
        Set gobjTax = CreateObject("zl9TaxBill.clsTaxBill")
        If Err.Number = 0 And Not gobjTax Is Nothing Then
            gblnTax = gobjTax.zlTaxUseable(2)
        End If
        On Error GoTo 0
    End If
    
    '����������Ʊ�ݴ�ӡ����
    On Error Resume Next
    gblnBillPrint = False
    Set gobjBillPrint = CreateObject("zlBillPrint.clsBillPrint")
    If Not gobjBillPrint Is Nothing Then
        gblnBillPrint = gobjBillPrint.zlInitialize(gcnOracle, glngSys, glngModul, UserInfo.���, UserInfo.����)
    End If
    On Error GoTo 0
    
    
    'Ȩ������
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, mstrPrivs, "ZL" & glngSys \ 100 & "_INSIDE_1137_1")
    Call Ȩ�޿���
    
    Call ClearErrInvoice
    

    'ˢ��ʱȱʡ��������(������),����ʱ����ʾ�κμ�¼
    mstrFilter = " And A.�շ�ʱ�� Between Trunc(sysdate) And Trunc(sysdate+1) " & _
                 " And A.����Ա����||''=[1] And A.��¼״̬=1"
    SQLCondition.Default = True
    SQLCondition.Operator = UserInfo.����
            
    Call SetHeader
    Call SetDetail
    Call SetMoney
    Call SetMenu(False)
    
    stbThis.Panels(2).Text = "��ˢ���嵥�����ù�������"
End Sub


Private Sub ClearErrInvoice()
'���ܣ��������Ա�ϴ��쳣�˳�ʱֻ����ʵ��Ʊ�Ŷ�û��ʵ�ʴ�ӡ�ĵ��ݵĽ��ʼ�¼�е�Ʊ�ݺ�
    Dim rsTmp As ADODB.Recordset, strSql As String, i As Long
 
    strSql = "Select A.NO" & vbNewLine & _
            "From ���˽��ʼ�¼ A," & vbNewLine & _
            "     (Select Max(NO) NO From ���˽��ʼ�¼ Where �շ�ʱ�� > Sysdate - 1 And ����Ա���� = [1]) B" & vbNewLine & _
            "Where A.NO = B.NO And A.ʵ��Ʊ�� Is Not Null And Not Exists (Select 1 From Ʊ�ݴ�ӡ���� C Where C.NO = B.NO)"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.����)
    If rsTmp.RecordCount > 0 Then
        strSql = "Zl_Ʊ����ʼ��_Update('" & rsTmp!NO & "','',3)"
        Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    End If

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    Dim cbrH As Long, staH As Long
    Dim sngVsc As Single, sngHsc As Single

    On Error Resume Next
    
    If WindowState = 1 Then Exit Sub
    
    mshList.MousePointer = 0
    
    '����ؼ���Ⱥ͸߶�
    cbrH = IIf(cbr.Visible, cbr.Height, 0)
    staH = IIf(stbThis.Visible, stbThis.Height, 0)
    sngVsc = mshDetail.Height / (mshDetail.Height + mshList.Height)
    sngHsc = mshMoney.Width / (mshMoney.Width + mshDetail.Width)
    
    If mblnMax Then
        sngVsc = 0.3: sngHsc = 0.2
        mblnMax = False
    End If
    If Me.WindowState = 2 Then mblnMax = True
    
    mshList.Left = Me.ScaleLeft
    mshList.Top = Me.ScaleTop + cbrH
    mshList.Width = Me.ScaleWidth
    mshList.Height = (Me.ScaleHeight - cbrH - staH - picHsc.Height) * (1 - sngVsc)
    
    picHsc.Top = mshList.Top + mshList.Height
    picHsc.Left = 0
    picHsc.Width = mshList.Width
    
    mshDetail.Left = 0
    mshDetail.Top = picHsc.Top + picHsc.Height
    mshDetail.Height = Me.ScaleHeight - cbrH - staH - picHsc.Height - mshList.Height
    mshDetail.Width = (Me.ScaleWidth - picVsc.Width) * (1 - sngHsc)
    
    picVsc.Top = mshDetail.Top
    picVsc.Left = mshDetail.Left + mshDetail.Width
    picVsc.Height = mshDetail.Height
    
    mshMoney.Top = mshDetail.Top
    mshMoney.Left = picVsc.Left + picVsc.Width
    mshMoney.Height = mshDetail.Height
    mshMoney.Width = Me.ScaleWidth - picVsc.Width - mshDetail.Width
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
        
    mbytType = 0
    mstrFilter = ""
    Unload frmBalanceFilter
    Unload frmBalanceGo
    Call SaveWinState(Me, App.ProductName)
    '33635
    Set mobjInPati = Nothing
    'ˢ�·�ʽ
    For i = 0 To mnuViewRefeshOptionItem.UBound
        If mnuViewRefeshOptionItem(i).Checked Then
            zlDatabase.SetPara "ˢ�·�ʽ", i, glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
            Exit For
        End If
    Next
    
    If Not gobjBillPrint Is Nothing Then
        Call gobjBillPrint.zlTerminate
        Set gobjBillPrint = Nothing
    End If
End Sub

Private Sub mnuViewGo_Click()
    frmBalanceGo.Show 1, Me
    If gblnOK Then Call SeekBill(frmBalanceGo.optHead)
End Sub

Private Sub SeekBill(blnHead As Boolean)
    Dim i As Long, int��Դ As Integer
    Dim blnFill As Boolean
    Dim strNo As String
    
    Screen.MousePointer = 11
    mblnGo = True
    stbThis.Panels(2).Text = "���ڶ�λ���������ĵ���,��ESC��ֹ ..."
    Me.Refresh
    
    For i = IIf(blnHead, 1, mlngGo) To mshList.Rows - 1
        DoEvents
        
        '�Ƚ�����
        blnFill = True
        With frmBalanceGo
            If .txtNO.Text <> "" Then
                blnFill = blnFill And mshList.TextMatrix(i, GetColNum("���ݺ�")) = .txtNO.Text
            End If
            If .txtFact.Text <> "" Then
                blnFill = blnFill And mshList.TextMatrix(i, GetColNum("Ʊ�ݺ�")) = .txtFact.Text
            End If
            If .txtסԺ��.Text <> "" Then
                blnFill = blnFill And mshList.TextMatrix(i, GetColNum("סԺ��")) = .txtסԺ��.Text
            End If
            If .txt����.Text <> "" Then
                blnFill = blnFill And UCase(mshList.TextMatrix(i, GetColNum("����"))) Like "*" & UCase(.txt����.Text) & "*"
            End If
        End With
        
        '�������˳�
        If blnFill Then
            mlngGo = i + 1
            mshList.Row = i: mshList.TopRow = i
            mshList.Col = 0: mshList.ColSel = mshList.Cols - 1
            
            strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
            int��Դ = Val(mshList.TextMatrix(mshList.Row, GetColNum("��־")))
            Call ShowDetail(mshList.TextMatrix(mshList.Row, GetColNum("����ID")), , strNo, int��Դ)
            Call ShowMoney(mshList.TextMatrix(mshList.Row, GetColNum("����ID")), , strNo)
            
            stbThis.Panels(2).Text = "�ҵ�һ����¼"
            Screen.MousePointer = 0: Exit Sub
        End If
        
        '��ESCȡ��
        If mblnGo = False Then
            stbThis.Panels(2).Text = "�û�ȡ����λ����"
            Screen.MousePointer = 0: Exit Sub
        End If
    Next
    mlngGo = 1
    stbThis.Panels(2).Text = "�Ѷ�λ���嵥β��"
    Screen.MousePointer = 0
End Sub

Private Function GetColNum(strHead As String) As Integer
    Dim i As Long
    For i = 0 To mshList.Cols - 1
        If mshList.TextMatrix(0, i) = strHead Then GetColNum = i: Exit Function
    Next
End Function

Private Sub mnuEdit_Print_Click()
    Call PrintBill(0)
End Sub

Private Sub mnuEdit_Print_Supplemental_Click()
    Call PrintBill(1)
End Sub

Private Sub PrintBill(bytMode As Byte)
'���ܣ���ǰ�տ��¼���´�ӡһ��Ʊ��
'bytMode=0-�ش�,1-����
    Dim strNo As String, lng����ID As Long, blnMediCare As Boolean, bytFlag As Byte '���ﻹ��סԺ
    Dim intInsure As Integer
    Dim lng����ID As Long, bytFunc As Byte
    
    strNo = mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
    If strNo = "" Then
        MsgBox "��ǰû�е��ݿ����ش�Ʊ�ݣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    lng����ID = Val(mshList.TextMatrix(mshList.Row, GetColNum("����ID")))
    lng����ID = Val(mshList.TextMatrix(mshList.Row, GetColNum("����ID")))
    bytFunc = IIf(Val(mshList.TextMatrix(mshList.Row, GetColNum("��־"))) = 1, 0, 1)
    
    '�Ƿ���ת������ݱ���
    If mblnNOMoved Then
        If Not ReturnMovedExes(strNo, 7, Me.Caption) Then Exit Sub
        mblnNOMoved = False  '��ʱ��ת���������ݱ�
    End If
    
     '����Ȩ��
    If bytMode = 0 Then
        If Not BillOperCheck(7, mshList.TextMatrix(mshList.Row, GetColNum("����Ա")), _
            CDate(mshList.TextMatrix(mshList.Row, GetColNum("�շ�ʱ��"))), "�ش�") Then Exit Sub
    Else
        If Trim(mshList.TextMatrix(mshList.Row, GetColNum("Ʊ�ݺ�"))) <> "" Then
            MsgBox "��ǰ�����Ѵ�ӡ��Ʊ��,���ܽ��в���", vbExclamation, gstrSysName
            Exit Sub
        End If
    End If
    
    intInsure = BalanceExistInsure(strNo, bytFlag)
    If RePrintBalance(strNo, Me, lng����ID, intInsure) Then
    
        '��ҽһ��ͨд����85950
        Call WriteInforToCard(Me, mlngModul, mstrPrivs, gobjSquare.objSquareCard, 0, bytFunc, lng����ID, lng����ID)

        If mnuViewRefeshOptionItem(1).Checked Then
            If MsgBox("��ǰ�����Ѹ��ĵ����嵥����,Ҫˢ���嵥��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mnuViewReFlash_Click
            End If
        ElseIf mnuViewRefeshOptionItem(2).Checked Then
            mnuViewReFlash_Click
        End If
    End If
End Sub

Private Sub mshList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshList.MouseRow = 0 Then
        mshList.MousePointer = 99
    Else
        mshList.MousePointer = 0
    End If
End Sub

Private Sub mshList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    
    lngCol = mshList.MouseCol
    
    If Button = 1 And mshList.MousePointer = 99 Then
        If mshList.TextMatrix(0, lngCol) = "" Then Exit Sub
        If mshList.TextMatrix(1, GetColNum("���ݺ�")) = "" Then Exit Sub
        If mshList.ColWidth(lngCol) = 0 Then Exit Sub
        If mrsList Is Nothing Then Exit Sub
        
        Set mshList.DataSource = Nothing

        mrsList.Sort = mshList.TextMatrix(0, lngCol) & IIf(mshList.ColData(lngCol) = 0, "", " DESC")
        mshList.ColData(lngCol) = (mshList.ColData(lngCol) + 1) Mod 2
        
        Call ShowBills(, True)
    End If
End Sub

Private Sub SetHeader()
    Dim strHead As String
    Dim i As Long
    
    '����65105:�������������Ϣ
    strHead = "����ID,1,0|��־,1,0|ҽ��,4,500|���ݺ�,4,850|Ʊ�ݺ�,4,850|����ID,1,750|�����,1,750|סԺ��,1,750|����,4,800|�Ա�,4,500|����,4,500|�ѱ�,4,750|��ʼ����,4,1000|��������,4,1000|���ʽ��,7,850|����Ա,4,800|�շ�ʱ��,4,1850|У��,4,500|��;����,4,800|��¼״̬,1,0"
    
    With mshList
        .Redraw = False
        
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshList, App.ProductName & "\" & Me.Name)
        
        .RowHeight(0) = 320
        
        '�ָ��ϴ���
        If mlngCurRow = 0 Then mlngCurRow = 1
        If mlngTopRow = 0 Then mlngTopRow = 1
        If mlngCurRow <= .Rows - 1 Then
            .Row = mlngCurRow
        Else
            .Row = .Rows - 1
        End If
        If mlngTopRow <= .Rows - 1 Then
            .TopRow = mlngTopRow
        Else
            .TopRow = .Row
        End If
        
        .Col = 0: .ColSel = .Cols - 1
        
        Call mshList_EnterCell

        .Redraw = True
    End With
End Sub

Private Sub ShowBills(Optional strIF As String, Optional blnSort As Boolean)
'����:��������ȡ�����б�(���˹���)
'����:strIF=��"AND"��ʼ��������
'     blnSort=�����¶�ȡ����,��������ʾ�����������
    Dim i As Long, j As Long, k As Long, strTable As String, rsTemp As ADODB.Recordset
    Dim strSql As String, str��Դ As String, strWhere As String
    
    
    On Error GoTo errH
    If Not blnSort Then
        Call zlCommFun.ShowFlash("���ڶ�ȡ�����б�,���Ժ� ...", Me)
        DoEvents
        Me.Refresh
        If SQLCondition.str��Դ = "" Then SQLCondition.str��Դ = "1111" '0000:����;סԺ;���;����
            
        str��Դ = ""
        For i = 1 To Len(SQLCondition.str��Դ)
            If Mid(SQLCondition.str��Դ, i, 1) = 1 Then
                str��Դ = str��Դ & "," & Choose(i, 1, 2, 4, 3)  '1-����;2-סԺ;3-����(���￨�ȶ�����շ�);4-���
            End If
        Next
        If str��Դ <> "" Then str��Դ = Mid(str��Դ, 2)
        If str��Դ = "" Then str��Դ = "-1"
        
        
        strTable = "" & _
        "   Select A.ID ,1 as סԺ��־,0 as �����־,A.NO,A.ʵ��Ʊ��,A.����ID,B.����ID as ���ò���ID,A.��ʼ����,A.��������,A.��¼״̬,B.���ʽ��,A.����Ա����,A.�շ�ʱ��,A.��;����,A.ԭ�� as ��Լ��λ,A.��������,B.��ҳID " & _
        "   From ���˽��ʼ�¼ A,סԺ���ü�¼ B,������Ϣ C " & _
        "   Where A.ID=B.����ID and  B.����ID=C.����ID" & _
                IIf(SQLCondition.str��Դ = "1111", "", " And Instr(',' || [11] || ',', ',' || Nvl(B.�����־,0) || ',') > 0 ") & strIF
        
        
        Select Case SQLCondition.str��Դ
        Case "1010", "1000", "0010"  '����
            strTable = Replace(strTable, "סԺ���ü�¼", "������ü�¼")
            strTable = Replace(strTable, "B.��ҳID", "Null As ��ҳID")
            strTable = Replace(strTable, "1 as סԺ��־,0 as �����־", "0 as סԺ��־,1 as �����־")
        Case "0101", "0001", "0100" 'סԺ
            '�Ѿ�����
        Case Else '�����סԺ
            strTable = strTable & vbCrLf & " Union ALL " & vbCrLf & Replace(Replace(Replace(strTable, "סԺ���ü�¼", "������ü�¼"), "1 as סԺ��־,0 as �����־", "0 as סԺ��־,1 as �����־"), "B.��ҳID", "Null As ��ҳID")
        End Select
        
        If frmBalanceFilter.mblnDateMoved Then
            strTable = strTable & vbCrLf & " Union ALL " & vbCrLf & Replace(Replace(Replace(strTable, "���˽��ʼ�¼", "H���˽��ʼ�¼"), "סԺ���ü�¼", "HסԺ���ü�¼"), "������ü�¼", "H������ü�¼")
        End If
        
        '����65105,������:�������ʱ����ʾ�����
        strSql = _
        " Select A.ID ����ID,decode(Max(סԺ��־),1,decode(max(�����־),1,3,2),1) as ��־ ,Decode(P.����,NULL,Decode(C.����,NULL,NULL,'��'),'��') as ҽ��,A.NO as ���ݺ�,A.ʵ��Ʊ�� as Ʊ�ݺ�," & _
        "        Decode(A.����ID,Null,' ',A.����ID) ����ID,Decode(Nvl(A.��������,0),2,' ',Decode(A.����ID,Null,' ',C.�����)) �����,Decode(A.����ID,Null,' ',Nvl(P.סԺ��,C.סԺ��)) סԺ��," & _
        "        Decode(A.����ID,Null,nvl(A.��Լ��λ,Q.����),C.����) ����,Decode(A.����ID,Null,' ',C.�Ա�) �Ա�," & _
        "        Decode(A.����ID,Null,' ',C.����) ����,Decode(A.����ID,Null,' ',Nvl(P.�ѱ�,C.�ѱ�)) as �ѱ�," & _
        "        To_Char(A.��ʼ����,'YYYY-MM-DD') as ��ʼ����,To_Char(A.��������,'YYYY-MM-DD') as ��������," & _
        "        To_Char(Sum(Decode(A.��¼״̬,2,-1,1) *A.���ʽ��),'999999999" & gstrDec & "') as ���ʽ��," & _
        "        A.����Ա���� as ����Ա,To_Char(A.�շ�ʱ��,'YYYY-MM-DD HH24:MI:SS') as �շ�ʱ��," & _
        "        ' '  as У��," & _
        "        Decode(Nvl(A.��;����,0),1,'��',' ') ��;����,Max(A.��¼״̬) as ��¼״̬" & _
        " From ( " & strTable & ") A,������Ϣ C,������ҳ P,��Լ��λ Q,��Ա�� N" & _
        " Where  A.���ò���ID=C.����ID And A.����Ա����=N.����  " & _
        "        And A.���ò���ID=P.����ID(+) And Nvl(A.��ҳID,0)=P.��ҳID(+) And C.��ͬ��λID=Q.ID(+)" & _
        "       And (N.վ��='" & gstrNodeNo & "' Or N.վ�� is Null)" & vbNewLine & _
        " Group by A.ID,Decode(P.����,NULL,Decode(C.����,NULL,NULL,'��'),'��'),A.NO,A.ʵ��Ʊ��,Decode(A.����ID,Null,' ',A.����ID),Decode(Nvl(A.��������,0),2,' ',Decode(A.����ID,Null,' ',C.�����)),Decode(A.����ID,Null,' ',Nvl(P.סԺ��,C.סԺ��))," & _
        "           Decode(A.����ID,Null,nvl(A.��Լ��λ,Q.����),C.����),Decode(A.����ID,Null,' ',C.�Ա�),Decode(A.����ID,Null,' ',C.����),Decode(A.����ID,Null,' ',Nvl(P.�ѱ�,C.�ѱ�))," & _
        "           To_Char(A.��ʼ����,'YYYY-MM-DD'),To_Char(A.��������,'YYYY-MM-DD')," & _
        "           A.����Ա����,To_Char(A.�շ�ʱ��,'YYYY-MM-DD HH24:MI:SS'),Decode(Nvl(A.��;����,0),1,'��',' ')"
        
        strSql = strSql & " Order by �շ�ʱ�� Desc,���ݺ� Desc"
        
        With SQLCondition
            If .Default Then
                Set mrsList = zlDatabase.OpenSQLRecord(strSql, Me.Caption, .Operator, "", "", "", "", "", 0, "", "", "", str��Դ)
            Else
                Set mrsList = zlDatabase.OpenSQLRecord(strSql, Me.Caption, .DateB, .DateE, .NOB, .NOE, .FactB, .FactE, Val(.InPatientID), .Patient, .Operator, .OutExseID, str��Դ)
            End If
        End With
    End If
    
    mshList.Redraw = False
    mshList.Clear
    mshList.Rows = 2
    
    mshDetail.Clear
    mshDetail.Rows = 2
    
    mshMoney.Clear
    mshMoney.Rows = 2
    
    If mrsList.EOF Then
        stbThis.Panels(2).Text = "��ǰ����û�й��˳��κε���"
        Call SetMenu(False)
    Else
        Set mshList.DataSource = mrsList
        stbThis.Panels(2) = "�� " & mrsList.RecordCount & " �ŵ���"
        Call SetMenu(True)
    End If
    
    '������ɫ
    If SQLCondition.Flag = 2 Then
        mshList.ForeColor = &HC0
    Else
        mshList.ForeColor = ForeColor
        k = GetColNum("��¼״̬")
        For i = 1 To mshList.Rows - 1
            If Val(mshList.TextMatrix(i, k)) = 2 Then
                '�˷Ѽ�¼�ú�ɫ
                mshList.Row = i
                For j = 0 To mshList.Cols - 1
                    mshList.Col = j
                    mshList.CellForeColor = &HC0
                Next
            ElseIf Val(mshList.TextMatrix(i, k)) = 3 Then
                '�����˹��ѵ�����ɫ
                mshList.Row = i
                For j = 0 To mshList.Cols - 1
                    mshList.Col = j
                    mshList.CellForeColor = &HC00000
                Next
            End If
            strSql = "Select 1 From ���ս�����ϸ Where ����ID=[1] And ��־=1 And ���㷽ʽ <> '�ֽ�'"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, Val(mshList.TextMatrix(i, GetColNum("����ID"))))
            If Not rsTemp.EOF Then
                mshList.TextMatrix(i, GetColNum("У��")) = "��"
            Else
                mshList.TextMatrix(i, GetColNum("У��")) = ""
            End If
        Next
    End If
    
    Call SetHeader '�˹����Ѱ���Call SetDetail,Call SetMoney
    
    If mshList.Row = 0 Or mshList.TextMatrix(mshList.Row, GetColNum("����ID")) = "" Then
        Call SetDetail
        Call SetMoney
    End If
    
    If Not blnSort Then Call zlCommFun.StopFlash
    mshList.Redraw = True
    Me.Refresh
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    mshList.Redraw = True
End Sub

Private Sub ShowDetail(Optional lng����ID As Long, Optional blnSort As Boolean, Optional strNo As String, Optional int��Դ As Integer = 2)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ʾ��ϸ���ü�¼
    '��Σ�int��Դ-1-����;2-סԺ;3-�����סԺ
    '���Σ�
    '���أ�
    '���ƣ����˺�
    '���ڣ�2010-03-08 19:34:16
    '˵����
    '------------------------------------------------------------------------------------------------------------------------

    Dim i As Long, j As Long, strSql As String, strDec As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    
    If Not blnSort Then
        
        If frmBalanceFilter.mblnDateMoved Then
            '���ʵ�ͨ������ID����ü�¼����,���ָ��NO�ŵĽ��ʵ��ں󱸱���,��ôν��ʵĵ��ݺ�һ���ں󱸱���
            'һ�Ž��ʵ������ļ��ʵ�������ͬʱ�����߱�ͺ󱸱���
            mblnNOMoved = zlDatabase.NOMoved("���˽��ʼ�¼", strNo, , , Me.Caption)
        Else
            mblnNOMoved = False   '����Ҫ����һ��
        End If
        
        strDec = gstrDec
        If lng����ID <> 0 Then
            Select Case int��Դ
            Case 1 '����
                strSql = "Select Max(Length(Abs(���ʽ��) - Trunc(Abs(���ʽ��))))-1 declen From " & IIf(mblnNOMoved, "H", "") & "������ü�¼ Where ����ID=[1]"
            Case 2 'סԺ
                strSql = "Select Max(Length(Abs(���ʽ��) - Trunc(Abs(���ʽ��))))-1 declen From " & IIf(mblnNOMoved, "H", "") & "סԺ���ü�¼ Where ����ID=[1]"
            Case Else
                
                strSql = "Select Length(Abs(���ʽ��) - Trunc(Abs(���ʽ��)))  as  declen From " & IIf(mblnNOMoved, "H", "") & "������ü�¼ Where ����ID=[1] Union ALL " & _
                         "Select Length(Abs(���ʽ��) - Trunc(Abs(���ʽ��)))   as  declen  From " & IIf(mblnNOMoved, "H", "") & "סԺ���ü�¼ Where ����ID=[1]"
                strSql = "Select Max(declen)-1 as declen  From ( " & strSql & ")"
            End Select
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID)
            If rsTmp.RecordCount > 0 Then
                If Len(strDec) < Len("0." & String(rsTmp!declen, "0")) Then
                    strDec = "0." & String(rsTmp!declen, "0")
                End If
            End If
        End If
        
        Select Case int��Դ
        Case 1  '����
            strSql = " (Select ����ID,NO,���,��������ID,�շ�ϸĿID,�����־,0 as ��ҳID,�վݷ�Ŀ,Ӥ����,���ʽ��,����ʱ�� From " & IIf(mblnNOMoved, "H", "") & "������ü�¼ A where A.����ID=[1] ) A "
            'strSQL = IIf(mblnNOMoved, "H", "") & "������ü�¼ A "
        Case 2  'סԺ
            strSql = IIf(mblnNOMoved, "H", "") & "סԺ���ü�¼ A"
        Case Else '�����סԺ
            strSql = " (Select ����ID,NO,���,��������ID,�շ�ϸĿID,�����־,0 as ��ҳID,�վݷ�Ŀ,Ӥ����,���ʽ��,����ʱ�� From " & IIf(mblnNOMoved, "H", "") & "������ü�¼ A where A.����ID=[1] Union ALL " & _
                       " Select ����ID,NO,���,��������ID,�շ�ϸĿID,�����־,��ҳID,�վݷ�Ŀ,Ӥ����,���ʽ��,����ʱ�� From " & IIf(mblnNOMoved, "H", "") & "סԺ���ü�¼ A where A.����ID=[1] )  A"
        End Select
        
        strSql = _
        "   Select Decode(�����־,1,'����',4,'����','��'||Nvl(A.��ҳID,0)||'��') as סԺ," & _
        "         A.NO as ���ݺ�,Nvl(B.����,'δ֪') as ��������,Nvl(E.����,D.����) as ��Ŀ," & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "E1.���� as ��Ʒ��,", "") & _
        "       A.�վݷ�Ŀ as ��Ŀ,Decode(Nvl(A.Ӥ����,0),0,'','��') as Ӥ����," & _
        "       To_Char(" & IIf(mbytType = 2, "-1*", "") & "A.���ʽ��,'999999999" & strDec & "') as ���ʽ��," & _
        "       To_Char(A.����ʱ��,'YYYY-MM-DD HH24:MI:SS') as ����ʱ��" & _
        " From " & strSql & ",���ű� B,�շ���ĿĿ¼ D,�շ���Ŀ���� E" & _
                IIf(gTy_System_Para.bytҩƷ������ʾ = 2, ",�շ���Ŀ���� E1", "") & _
        " Where A.��������ID=B.ID(+) And A.�շ�ϸĿID=D.ID" & _
        "       And A.�շ�ϸĿID=E.�շ�ϸĿID(+) And E.����(+)=1 And E.����(+)=" & IIf(gTy_System_Para.bytҩƷ������ʾ = 1, 3, 1) & _
                IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "       And A.�շ�ϸĿID=E1.�շ�ϸĿID(+) And E1.����(+)=1 And E1.����(+)=3", "") & _
        "       And A.����ID=[1]" & _
        " Order by סԺ Desc,����ʱ�� Desc,���ݺ� Desc,A.���"
        Set mrsDetail = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID)
    End If
    
    mshDetail.Redraw = False
    mshDetail.Clear
    mshDetail.Rows = 2
    mshDetail.ForeColor = IIf(mbytType = 2, &HC0, ForeColor)

    If Not mrsDetail.EOF Then Set mshDetail.DataSource = mrsDetail
    
    '������ɫ
    If mbytType = 2 Then
        '�˷�ֱ��Ϊ��ɫ
        mshDetail.ForeColor = &HC0
    Else
        'ԭʼ�����˹���Ϊ��ɫ
        mshDetail.ForeColor = ForeColor
        If mbytType = 3 Then
            For i = 1 To mshDetail.Rows - 1
                mshDetail.Row = i
                For j = 0 To mshDetail.Cols - 1
                    mshDetail.Col = j
                    mshDetail.CellForeColor = &HC00000
                Next
            Next
        End If
    End If
    
    Call SetDetail
    mshDetail.Redraw = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    mshDetail.Redraw = True
End Sub

Private Sub SetDetail()
    Dim strHead As String
    Dim i As Long
    
    strHead = "סԺ,4,750|���ݺ�,4,850|��������,1,850|��Ŀ,1,1800" & IIf(gTy_System_Para.bytҩƷ������ʾ = 2, "|��Ʒ��,1,1600", "") & "|��Ŀ,1,850|Ӥ����,4,650|���ʽ��,7,850|����ʱ��,1,1850"
    
    With mshDetail
        .Redraw = False
        
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshDetail, App.ProductName & "\" & Me.Name)
        '���˺�:27990 2010-02-22 17:34:32
        For i = 0 To .Cols - 1
            If .TextMatrix(0, i) = "��Ʒ��" Then
                If gTy_System_Para.bytҩƷ������ʾ = 2 Then
                    If .ColWidth(i) <= 0 Then .ColWidth(i) = 1600
                Else
                    .ColWidth(i) = 0
                End If
            End If
        Next
                
        .RowHeight(0) = 320
        
        .Row = 1: .Col = 0: .ColSel = .Cols - 1
        Call mshDetail_EnterCell
        
        .Redraw = True
    End With
End Sub

Private Sub ShowMoney(Optional lng����ID As Long, Optional blnSort As Boolean, _
    Optional strNo As String)
    Dim i As Long, strSql As String
    On Error GoTo errH
    
    '�����ǰ���ʵ��ں󱸱���,������صĽ��ʷ�ʽһ���ں󱸱���
    If Not blnSort Then
        strSql = "" & _
        " Select Decode(Substr(��¼����,Length(��¼����),1),1,'��Ԥ��',2,'����') as ����," & _
        "       NO as ���ݺ�,To_Char(" & IIf(mbytType = 2, "-1*", "") & "��Ԥ��,'FM9999999990.00999') as ���," & _
        "       ���㷽ʽ,�������" & _
        " From ����Ԥ����¼ " & _
        " Where ����ID=[1] And ��Ԥ�� <> 0 " & _
        " Order by ���� Desc,NO Desc,���㷽ʽ"
        If mblnNOMoved Then strSql = Replace(strSql, "����Ԥ����¼", "H����Ԥ����¼")
        Set mrsMoney = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng����ID)
    End If
    
    mshMoney.Clear
    mshMoney.Rows = 2
    mshMoney.ForeColor = mshList.ForeColor
    If Not mrsMoney.EOF Then Set mshMoney.DataSource = mrsMoney
    Call SetMoney
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetMoney()
    Dim strHead As String
    Dim i As Long
    
    strHead = "����,4,650|���ݺ�,4,850|���,7,850|���㷽ʽ,1,850|�������,1,1000"
    With mshMoney
        .Redraw = False
        
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
        
        If Not Visible Then Call RestoreFlexState(mshMoney, App.ProductName & "\" & Me.Name)
        
        .RowHeight(0) = 320
        
        .Row = 1: .Col = 0: .ColSel = .Cols - 1
        Call mshMoney_EnterCell
        
        .Redraw = True
    End With
End Sub

Private Sub mshDetail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshDetail.MouseRow = 0 Then
        mshDetail.MousePointer = 99
    Else
        mshDetail.MousePointer = 0
    End If
End Sub

Private Sub mshDetail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    
    lngCol = mshDetail.MouseCol
    
    If Button = 1 And mshDetail.MousePointer = 99 Then
        If mshDetail.TextMatrix(0, lngCol) = "" Then Exit Sub
        If mshDetail.TextMatrix(1, 0) = "" Then Exit Sub
        If mrsDetail Is Nothing Then Exit Sub
        
        Set mshDetail.DataSource = Nothing

        mrsDetail.Sort = mshDetail.TextMatrix(0, lngCol) & IIf(mshDetail.ColData(lngCol) = 0, "", " DESC")
        mshDetail.ColData(lngCol) = (mshDetail.ColData(lngCol) + 1) Mod 2
        
        Call ShowDetail(, True)
    End If
End Sub

Private Sub mshMoney_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mshMoney.MouseRow = 0 Then
        mshMoney.MousePointer = 99
    Else
        mshMoney.MousePointer = 0
    End If
End Sub

Private Sub mshMoney_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngCol As Long
    
    lngCol = mshMoney.MouseCol
    
    If Button = 1 And mshMoney.MousePointer = 99 Then
        If mshMoney.TextMatrix(0, lngCol) = "" Then Exit Sub
        If mshMoney.TextMatrix(1, 0) = "" Then Exit Sub
        If mrsMoney Is Nothing Then Exit Sub
        
        Set mshMoney.DataSource = Nothing

        mrsMoney.Sort = mshMoney.TextMatrix(0, lngCol) & IIf(mshMoney.ColData(lngCol) = 0, "", " DESC")
        mshMoney.ColData(lngCol) = (mshMoney.ColData(lngCol) + 1) Mod 2
        
        Call ShowMoney(, True)
    End If
End Sub

Private Sub PrintDetail()
'���ܣ�������б�
    Dim objOut As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim bytR As Byte, intRow As Integer
    
    intRow = mshDetail.Row
    
    '��ͷ
    objOut.Title.Text = "���˽��ʵ�����ϸ"
    objOut.Title.Font.Name = "����_GB2312"
    objOut.Title.Font.Size = 18
    objOut.Title.Font.Bold = True
    
    '����
    With frmBalanceFilter
        objRow.Add "���ݺţ�" & mshList.TextMatrix(mshList.Row, GetColNum("���ݺ�"))
        objRow.Add "���ʷ�Χ��" & mshList.TextMatrix(mshList.Row, GetColNum("��ʼ����")) & " �� " & mshList.TextMatrix(mshList.Row, GetColNum("��������"))
        objOut.UnderAppRows.Add objRow
    
        Set objRow = New zlTabAppRow
        objRow.Add "סԺ�ţ�" & mshList.TextMatrix(mshList.Row, GetColNum("סԺ��"))
        objRow.Add "������" & mshList.TextMatrix(mshList.Row, GetColNum("����"))
        objOut.UnderAppRows.Add objRow
    End With
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ�ˣ�" & UserInfo.����
    objRow.Add "��ӡ���ڣ�" & Format(zlDatabase.Currentdate(), "yyyy��MM��dd��")
    objOut.BelowAppRows.Add objRow
    
    '����
    mshDetail.Redraw = False
    Set objOut.Body = mshDetail
    
    bytR = zlPrintAsk(objOut)
    Me.Refresh
    If bytR <> 0 Then zlPrintOrView1Grd objOut, bytR
    
    mshDetail.Row = intRow
    mshDetail.Col = 0: mshDetail.ColSel = mshDetail.Cols - 1
    mshDetail.Redraw = True
End Sub

Private Sub mnuViewRefeshOptionItem_Click(Index As Integer)
    Dim i As Long
    For i = 0 To mnuViewRefeshOptionItem.UBound
        mnuViewRefeshOptionItem(i).Checked = i = Index
    Next
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

