VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmDrugPlanList 
   Caption         =   "ҩƷ�ƻ�����"
   ClientHeight    =   5895
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9120
   Icon            =   "frmDrugPlanList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5895
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picSeparate_s 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   30
      MousePointer    =   7  'Size N S
      ScaleHeight     =   300
      ScaleWidth      =   4815
      TabIndex        =   5
      Top             =   2790
      Width           =   4815
      Begin VB.Label lblRange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ѯ��Χ:1999��8��12����1999��9��12��"
         Height          =   180
         Left            =   0
         TabIndex        =   6
         Top             =   120
         Width           =   3690
      End
   End
   Begin VB.CommandButton Cmd���� 
      Caption         =   "����(&V)"
      Height          =   350
      Left            =   5160
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1100
   End
   Begin ComCtl3.CoolBar cbrTool 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   9120
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tlbTool"
      MinWidth1       =   4995
      MinHeight1      =   720
      Width1          =   4995
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Caption2        =   "�ⷿ"
      Child2          =   "cboStock"
      MinHeight2      =   300
      Width2          =   1995
      NewRow2         =   0   'False
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   7740
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   1290
      End
      Begin MSComctlLib.Toolbar tlbTool 
         Height          =   720
         Left            =   165
         TabIndex        =   1
         Top             =   30
         Width           =   6840
         _ExtentX        =   12065
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsCold"
         HotImageList    =   "ilsHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   17
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "PrintView"
               Description     =   "Ԥ��"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "Print"
               Description     =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "PrintSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Add"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   3
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "next"
                     Text            =   "�������ڼƻ�"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "period"
                     Text            =   "�������ڼƻ�"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�޸�"
               Key             =   "Modify"
               Description     =   "�޸�"
               Object.ToolTipText     =   "�޸�"
               Object.Tag             =   "�޸�"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ɾ��"
               Key             =   "Delete"
               Description     =   "ɾ��"
               Object.ToolTipText     =   "ɾ��"
               Object.Tag             =   "ɾ��"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "EditSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "���"
               Key             =   "Verify"
               Description     =   "���"
               Object.ToolTipText     =   "���"
               Object.Tag             =   "���"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Check"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "���"
               Key             =   "Clear"
               Description     =   "���"
               Object.ToolTipText     =   "���"
               Object.Tag             =   "���"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ȡ��"
               Key             =   "Cancel"
               Description     =   "ȡ�����"
               Object.ToolTipText     =   "ȡ��"
               Object.Tag             =   "ȡ��"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "VerifySeparate"
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Search"
               Description     =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "ˢ��"
               Key             =   "Refresh"
               Description     =   "ˢ��"
               Object.ToolTipText     =   "ˢ��"
               Object.Tag             =   "ˢ��"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FindSeparate"
               Style           =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "Help"
               Description     =   "��������"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "Exit"
               Description     =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   11
            EndProperty
         EndProperty
         MouseIcon       =   "frmDrugPlanList.frx":014A
      End
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   5535
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDrugPlanList.frx":0464
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11007
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
   Begin MSComctlLib.ImageList ilsCold 
      Left            =   0
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":0CF8
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":0F18
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":1138
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":1354
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":1574
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":1794
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":19B0
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":1BCC
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":1DE6
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":1F40
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":215C
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":237C
            Key             =   "Cancel"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":2596
            Key             =   "Check"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHot 
      Left            =   600
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":2C90
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":2EB0
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":30D0
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":32EC
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":350C
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":372C
            Key             =   "Verify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":3948
            Key             =   "Restore"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":3B64
            Key             =   "Search"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":3D7E
            Key             =   "Refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":3ED8
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":40F8
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":4318
            Key             =   "Cancel"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugPlanList.frx":4532
            Key             =   "Check"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   1005
      Left            =   0
      TabIndex        =   7
      Top             =   1200
      Width           =   6255
      _cx             =   11033
      _cy             =   1773
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16769992
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDrugPlanList.frx":4C2C
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   0   'False
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfDetail 
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   3240
      Width           =   5655
      _cx             =   9975
      _cy             =   1720
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16053482
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDrugPlanList.frx":4CA1
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   1
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   0   'False
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
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
      Begin VB.Menu mnuFileBillPrint 
         Caption         =   "���ݴ�ӡ(&B)"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuFileBillPreview 
         Caption         =   "����Ԥ��(&L)"
      End
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFileLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileParameter 
         Caption         =   "��������(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "�༭(&E)"
      Begin VB.Menu mnuEditAdd 
         Caption         =   "����"
         Begin VB.Menu mnuEditAddPlan 
            Caption         =   "�������ڼƻ�(Ĭ��)(&N)"
            Index           =   0
         End
         Begin VB.Menu mnuEditAddPlan 
            Caption         =   "�������ڼƻ�(&P)"
            Index           =   1
         End
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "�޸�(&M)"
      End
      Begin VB.Menu mnuEditDel 
         Caption         =   "ɾ��(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditVerify 
         Caption         =   "���(&V)"
      End
      Begin VB.Menu mnuEditCheck 
         Caption         =   "����(&R)"
      End
      Begin VB.Menu mnuEditClear 
         Caption         =   "���(&S)"
      End
      Begin VB.Menu mnuEditCancel 
         Caption         =   "ȡ��(&C)"
      End
      Begin VB.Menu mnuEditLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditExeAmount 
         Caption         =   "�޸�ִ������(&E)"
      End
      Begin VB.Menu mnuEditExport 
         Caption         =   "�ɹ��ƻ�����(&X)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEditDisplay 
         Caption         =   "�鿴����(&W)"
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
         Begin VB.Menu mnuViewToolText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewColDefine 
         Caption         =   "������(&C)"
      End
      Begin VB.Menu mnuViewLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewSearch 
         Caption         =   "����(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewLine4 
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
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "������̳(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&M)..."
         End
      End
      Begin VB.Menu mnuHelpLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmDrugPlanList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngMode As Long
Private mstrFind As String
Private mintPreCol As Integer           'ǰһ�ε���ͷ��������
Private mintsort As Integer             'ǰһ�ε���ͷ������
Private mblnBootUp As Boolean
Private mlastRow As Long                '�ϴε������
Private mstrPrivs As String
Private mint�۸���ʾ As Integer         '0:��ʾ�ɱ���;  1:��ʾ�ۼ�;  2:��ʾ�ɱ����ۼ�
Private mblnViewCost As Boolean             '�鿴�ɱ��� true-���Բ鿴�ɱ��� false-�����Բ鿴�ɱ���
Private Const MStrCaption As String = "ҩƷ�ƻ�����"

Private strStart As Date
Private strEnd As Date
Private strVerifyStart As Date
Private strVerifyEnd As Date
Private mlng�ⷿID As Long  '�ⷿid
Private mintUnit As Integer             '��λϵ����1-�ۼ�;2-����;3-סԺ;4-ҩ��
Private mintPlanPoint As Integer        'ȫԺ�ƻ�����վ�� 0-Ҫ��վ�㣬1-����վ��

'�Ӳ�������ȡҩƷ�۸����������С��λ��
Private mintShowCostDigit As Integer            '�ɱ���С��λ��
Private mintShowPriceDigit As Integer           '�ۼ�С��λ��
Private mintShowNumberDigit As Integer          '����С��λ��
Private mintShowMoneyDigit As Integer           '���С��λ��

Private mstrCostFormat As String
Private mstrPriceFormat As String
Private mstrNumberFormat As String
Private mstrMoneyFormat As String

Private Type Type_SQLCondition
    strNO��ʼ As String
    strNO���� As String
    date����ʱ�俪ʼ As Date
    date����ʱ����� As Date
    date���ʱ�俪ʼ As Date
    date���ʱ����� As Date
    date����ʱ�俪ʼ As Date
    date����ʱ����� As Date
    str������ As String
    str����� As String
    str������ As String
    lng�ƻ����� As Long
    lng���Ʒ��� As Long
    lngҩƷ As Long
End Type

Private SQLCondition As Type_SQLCondition
'�������������
Private Function CheckDepend() As Boolean
    Dim rsDepend As New Recordset
    
    On Error GoTo errHandle
    CheckDepend = False
    
    gstrSQL = "SELECT DISTINCT a.id, a.���� " _
            & "FROM ��������˵�� c, �������ʷ��� b, ���ű� a " _
            & "Where (a.վ�� = [1] Or a.վ�� is Null) And c.�������� = b.���� " _
            & "  AND Instr('HIJKLMN',b.����,1) > 0 " _
            & "  AND a.id = c.����id " _
            & "  AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' " _
            & IIf(IsHavePrivs(mstrPrivs, "���пⷿ"), "", " And a.ID IN (Select ����ID From ������Ա Where ��ԱID=[2])")
    Set rsDepend = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption, gstrNodeNo, UserInfo.�û�ID)
    
    If rsDepend.EOF Then
        MsgBox "û������ҩ�����ʵĲ���,��鿴���Ź���", vbInformation, gstrSysName
        rsDepend.Close
        Exit Function
    End If
            
    With cboStock
        .Clear
        '��վ��ʱ��������ȫԺ
        '0-Ҫ��վ�㣬1-����վ��
        If (gstrNodeNo = "-" Or gstrNodeNo = "0") Or mintPlanPoint = 1 Then
            .AddItem "ȫԺ"
            .ItemData(.NewIndex) = 0
        End If
        
        Do While Not rsDepend.EOF
            
            .AddItem rsDepend!����
            .ItemData(.NewIndex) = rsDepend!id
            If rsDepend!id = UserInfo.����ID Then
                .ListIndex = .NewIndex
            End If
            rsDepend.MoveNext
        Loop
        rsDepend.Close
        
        If .ListIndex = -1 Then
            If Not IsHavePrivs(mstrPrivs, "���пⷿ") Then
                MsgBox "�㲻��ҩ��������Ա�Ҳ��������пⷿ��Ȩ�ޣ����ܽ��룡", vbInformation, gstrSysName
                Unload Me
                Exit Function
            End If
            .ListIndex = 0
        End If
    End With
    CheckDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cboStock_Click()
    If mblnBootUp Then
        mlng�ⷿID = Me.cboStock.ItemData(Me.cboStock.ListIndex)
    
        Call GetDrugDigit(mlng�ⷿID, MStrCaption, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
        
        '��֯��ʽ����
        mstrCostFormat = "'999999999990." & String(mintShowCostDigit, "0") & "'"
        mstrPriceFormat = "'999999999990." & String(mintShowPriceDigit, "0") & "'"
        mstrNumberFormat = "'999999999990." & String(mintShowNumberDigit, "0") & "'"
        mstrMoneyFormat = "'999999999990." & String(mintShowMoneyDigit, "0") & "'"
        mnuViewRefresh_Click
    End If
End Sub

Private Sub cbrTool_Resize()
    Form_Resize
End Sub

Private Sub GetList(ByVal StrFind As String)
    Dim rsList As New Recordset
    Dim n As Integer
    Dim strSQL As String
     
    On Error GoTo errHandle
    Call zlCommFun.ShowFlash("��������ҩƷ�����¼,���Ժ� ...", Me)
    DoEvents
    Screen.MousePointer = vbHourglass
    
    vsfList.Redraw = flexRDNone
    strSQL = ""
    If SQLCondition.lngҩƷ <> 0 Then
        strSQL = ", ҩƷ�ƻ����� C "
        StrFind = " And a.Id = c.�ƻ�id and C.ҩƷid=[15] " & StrFind
    End If
    
    gstrSQL = " SELECT a.NO, a.ID, DECODE(a.�ƻ�����,0,'��ʱ',1,'�¶ȼƻ�',2,'���ȼƻ�',3,'��ȼƻ�','�ܼƻ�') AS �ƻ����� ," & _
        "a.�ڼ�,DECODE(A.���Ʒ���, 0, '�����������', 1, '����ͬ�����β��շ�', 2, '�ٽ��ڼ�ƽ�����շ�', 3, 'ҩƷ����������շ�', 4, 'ҩƷ�����������շ�', '�Զ���������շ�') AS ���Ʒ��� ," & _
        "a.������,TO_CHAR(a.��������,'YYYY-MM-DD HH24:MI:SS') AS ��������, a.�����, " & _
        "TO_CHAR(a.�������,'YYYY-MM-DD HH24:MI:SS') AS �������,a.������,TO_CHAR(a.��������,'YYYY-MM-DD HH24:MI:SS') AS ��������, b.���� ����ҩ��, a.����˵�� " & _
        " FROM ҩƷ�ɹ��ƻ� A, ���ű� B " & strSQL & _
        " WHERE a.ҩ��ID = b.ID(+) And NVL(a.�ⷿID,0)+0= [11] " & StrFind & _
        " ORDER BY A.NO DESC "

    Set rsList = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption, _
        SQLCondition.strNO��ʼ, _
        SQLCondition.strNO����, _
        SQLCondition.date����ʱ�俪ʼ, _
        SQLCondition.date����ʱ�����, _
        SQLCondition.date���ʱ�俪ʼ, _
        SQLCondition.date���ʱ�����, _
        SQLCondition.str������, _
        SQLCondition.str�����, _
        SQLCondition.lng�ƻ�����, _
        SQLCondition.lng���Ʒ���, _
        mlng�ⷿID, _
        SQLCondition.date����ʱ�俪ʼ, _
        SQLCondition.date����ʱ�����, _
        SQLCondition.str������, _
        SQLCondition.lngҩƷ)
    Set vsfList.DataSource = rsList
    With vsfList
        If .rows = 1 Then
            .rows = .rows + 100
            .Row = 1
            .TopRow = 1
            .rows = .rows - 99
        End If
        .Row = 1
        .Col = 0
        .ColSel = .Cols - 1
        
        For n = 0 To .Cols - 1
            .ColKey(n) = .TextMatrix(0, n)
            .FixedAlignment(n) = flexAlignCenterCenter
        Next
    End With
    SetListColWidth
    
    '�Ƿ���ʾ����ҩ����
    Call View����ҩ��(rsList)
    
    vsfList.Redraw = flexRDDirect
    Call zlCommFun.StopFlash
    Screen.MousePointer = vbDefault
    SetEnable
    staThis.Panels(2).Text = "��ǰ����" & rsList.RecordCount & "�ŵ���"
    rsList.Close
    Call vsfList_EnterCell
    vsfList.Row = 1
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub View����ҩ��(ByVal rsTmp As ADODB.Recordset)
'���ܣ����ҩ��ID������Ϣ����ȷ���Ƿ���ʾҩ����Ϣ����
    vsfList.ColHidden(vsfList.ColIndex("����ҩ��")) = True
    If rsTmp.RecordCount <= 0 Then Exit Sub
    With rsTmp
        .MoveFirst
        Do While Not .EOF
            If Nvl(!����ҩ��) <> "" Then
                vsfList.ColHidden(vsfList.ColIndex("����ҩ��")) = False
                Exit Sub
            End If
            .MoveNext
        Loop
    End With
End Sub

'��ͷ�п��ʼ
Private Sub SetListColWidth()
    Dim intCol As Integer
    
    With vsfList
        .ColAlignment(.ColIndex("NO")) = flexAlignLeftCenter
        .ColAlignment(.ColIndex("�ƻ�����")) = flexAlignLeftCenter
        .ColAlignment(.ColIndex("�ڼ�")) = flexAlignLeftCenter
        
        If mblnBootUp = False Then
            For intCol = 0 To .Cols - 1
                .ColWidth(intCol) = 1500
            Next
        End If
        .ColWidth(1) = 0
        
    End With
End Sub

'����Ȩ�����ò�ͬ����ʾ��Ŀ
Private Sub SetVisable()
    'ҩƷ�ƻ���������Ȩ�ޣ��������á����������пⷿ���Ǽǡ��޸ġ�ɾ�������ա������ȡ����ˡ����ˡ�ȡ�����ˡ��޸�ִ������

    If Not IsHavePrivs(mstrPrivs, "����") Then
        mnuEditAdd.Visible = False
        tlbTool.Buttons("Add").Visible = False
    End If
    
    If Not IsHavePrivs(mstrPrivs, "�޸�") Then
        mnuEditModify.Visible = False
        tlbTool.Buttons("Modify").Visible = False
    End If
    
    If Not IsHavePrivs(mstrPrivs, "ɾ��") Then
        mnuEditDel.Visible = False
        tlbTool.Buttons("Delete").Visible = False
         '��û�����б༭Ȩ��ʱ���Ѳ˵��͹������ϵ���Ӧ�ķָ������Ρ�
        If mnuEditAdd.Visible = False And mnuEditModify.Visible = False Then
            mnuEditLine1.Visible = False
            tlbTool.Buttons("EditSeparate").Visible = False
        End If
    End If
    
    If Not IsHavePrivs(mstrPrivs, "���") Then
        mnuEditVerify.Visible = False
        tlbTool.Buttons("Verify").Visible = False
    End If
    
    If Not IsHavePrivs(mstrPrivs, "���") Then
        mnuEditClear.Visible = False
        tlbTool.Buttons("Clear").Visible = False
    End If
    
    If Not IsHavePrivs(mstrPrivs, "ȡ�����") And Not IsHavePrivs(mstrPrivs, "ȡ������") Then
        mnuEditCancel.Visible = False
        tlbTool.Buttons("Cancel").Visible = False
        If mnuEditVerify.Visible = False And mnuEditClear.Visible = False Then
            mnuEditLine2.Visible = False
            tlbTool.Buttons("VerifySeparate").Visible = False
        End If
    End If
    
    If Not IsHavePrivs(mstrPrivs, "�ɹ��ƻ���ӡ") Then
        mnuFileBillPrint.Visible = False
        mnuFileBillPreview.Visible = False
    End If

    If Not IsHavePrivs(mstrPrivs, "����") Then
        mnuEditCheck.Visible = False
        tlbTool.Buttons("Check").Visible = False
    End If
    
    If Not IsHavePrivs(mstrPrivs, "�޸�ִ������") Then
        mnuEditExeAmount.Visible = False
    End If
End Sub


Private Sub Cmd����_Click()
    Call mnuEditDisplay_Click
End Sub

Private Sub Form_Activate()
    If vsfList.Visible Then
        vsfList.SetFocus
        vsfList.Row = 1
        vsfDetail.Row = IIf(vsfDetail.rows > 1, 1, 0)
    End If
End Sub

Private Sub Form_Load()
    '�ָ�����
    Dim strStart As String
    Dim strEnd As String
    Dim StrFind As String
    Dim dateCurrentDate As Date
    Dim int��ѯ���� As Integer
    
    mblnBootUp = False
    mlngMode = glngModul
    mstrPrivs = gstrprivs
    
    mblnViewCost = IsHavePrivs(mstrPrivs, "�鿴�ɱ���")
    mint�۸���ʾ = Val(zlDatabase.GetPara("�۸���ʾ��ʽ", glngSys, ģ���.ҩƷ�ƻ�))
    mintPlanPoint = Val(zlDatabase.GetPara("ȫԺ�ƻ�����վ��", glngSys, mlngMode, 0))
    
    If Not CheckDepend Then
        Unload Me
        Exit Sub
    End If
    
    On Error Resume Next
    'ʵ�����ɹ�ƽ̨�ӿ�
    If gobjDrugPurchase Is Nothing Then
        Set gobjDrugPurchase = CreateObject("zlDrugPurchase.clsDrugPurchase")
    End If
    Err.Clear
    On Error GoTo 0
    If Not gobjDrugPurchase Is Nothing Then
        mnuEditExport.Visible = True
    End If
    
    mlng�ⷿID = Me.cboStock.ItemData(Me.cboStock.ListIndex)
    
    Call GetDrugDigit(mlng�ⷿID, MStrCaption, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
    
    '��֯��ʽ����
    mstrCostFormat = "'999999999990." & String(mintShowCostDigit, "0") & "'"
    mstrPriceFormat = "'999999999990." & String(mintShowPriceDigit, "0") & "'"
    mstrNumberFormat = "'999999999990." & String(mintShowNumberDigit, "0") & "'"
    mstrMoneyFormat = "'999999999990." & String(mintShowMoneyDigit, "0") & "'"
    
    mlastRow = 0
    SetVisable  '����Ȩ�����ò�ͬ����ʾ��Ŀ
    
    strVerifyStart = "1901-01-01"
    strVerifyEnd = "1901-01-01"
    
    dateCurrentDate = zlDatabase.Currentdate
    int��ѯ���� = Val(zlDatabase.GetPara("��ѯ����", glngSys, mlngMode, 7)) - 1
    strStart = Format(DateAdd("d", -int��ѯ����, dateCurrentDate), "yyyy-MM-dd")
    strEnd = Format(dateCurrentDate, "yyyy-MM-dd")
    
    StrFind = " AND A.������� is Null And A.�������� Between [3] And [4] "
    SQLCondition.date����ʱ�俪ʼ = CDate(Format(strStart, "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date����ʱ����� = CDate(Format(strEnd, "yyyy-mm-dd") & " 23:59:59")
    
    mstrFind = StrFind
    
    lblRange.Caption = "��ѯ��Χ:" & Format(dateCurrentDate, "yyyy��MM��dd��") & "��" & Format(dateCurrentDate, "yyyy��MM��dd��")
    GetList (mstrFind)  '�г�����ͷ
   
    RestoreWinState Me, App.ProductName, MStrCaption
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
    mblnBootUp = True
End Sub

Private Sub Form_Resize()
    '����λ������
    
    On Error Resume Next
    
    If Me.WindowState = 1 Then Exit Sub
    
    cbrTool.Bands(1).MinHeight = tlbTool.Height
    With cbrTool
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth - .Left
    End With
    
    With picSeparate_s
        .Height = 300
        .Left = 0
        .Width = cbrTool.Width
    End With
    
    With vsfList
        .Top = IIf(cbrTool.Visible, cbrTool.Height, 0)
        .Left = 0
        .Width = cbrTool.Width
        .Height = picSeparate_s.Top - .Top
    End With
    
    With Cmd����
        .Left = Me.ScaleWidth - .Width - 100
        .Top = vsfList.Top + vsfList.Height + 30
        .ZOrder
    End With
    
    With vsfDetail
        .Top = picSeparate_s.Top + picSeparate_s.Height + 100
        .Left = 0
        .Height = ScaleHeight - .Top - IIf(staThis.Visible, staThis.Height, 0)
        .Width = cbrTool.Width
    End With
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, MStrCaption
End Sub

Private Sub mnuEditAddPlan_Click(Index As Integer)
    Dim strNo As String
    Dim blnSuccess As Boolean

    strNo = ""
    '����
    Select Case Index
        Case 0 '�������ڼƻ�
            frmDrugPlanCard.ShowCard Me, strNo, 1, blnSuccess, cboStock.ItemData(cboStock.ListIndex)
        Case 1 '�������ڼƻ�
            frmDrugPlanCard.ShowCard Me, strNo, 1, blnSuccess, cboStock.ItemData(cboStock.ListIndex), 1
    End Select
    
    If blnSuccess = True Then
        mnuViewRefresh_Click
    End If
    
End Sub

Private Sub mnuEditCancel_Click()
    'ȡ�����
    Dim lngBillId As Long
    Dim intRow As Integer
    Dim intReturn As Integer
    Dim intType As Integer
    
    With vsfList
        
        On Error GoTo errHandle
        intRow = .Row
        lngBillId = .TextMatrix(intRow, 1)
        intType = IIf(.TextMatrix(intRow, .ColIndex("������")) = "", 0, 1)
        intReturn = MsgBox("��ȷʵҪȡ����˵��ݺ�Ϊ��" & .TextMatrix(.Row, 0) & "���Ĳɹ��ƻ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)

        If intReturn = vbYes Then
            gstrSQL = "zl_ҩƷ�ƻ�����_Cancel(" & lngBillId & "," & intType & ")"
            
            If gstrSQL = "" Then Exit Sub
            Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)
            
        End If
    End With
    
    Call mnuViewRefresh_Click
    Exit Sub

errHandle:
    Exit Sub
End Sub

Private Sub mnuEditCheck_Click()
    '����
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    With vsfList
        strNo = .TextMatrix(.Row, 0)
        frmDrugPlanCard.ShowCard Me, strNo, 5, blnSuccess, cboStock.ItemData(cboStock.ListIndex)
    End With
    
    If blnSuccess = True Then
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditClear_Click()
    '���
    Dim lngBillId As Long
    Dim intRow As Integer
    Dim strTitle As String
    Dim intReturn As Integer
    Dim intRecord As Integer
     
    With vsfList
        
        On Error GoTo errHandle
        intRow = .Row
        lngBillId = .TextMatrix(intRow, 1)
        intReturn = MsgBox("��ȷʵҪ������ݺ�Ϊ��" & .TextMatrix(.Row, 0) & "���Ĳɹ��ƻ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
        intRecord = .rows - 1
        If intReturn = vbYes Then
            gstrSQL = "zl_ҩƷ�ƻ�����_DELETE('" & lngBillId & "')"
            
            If gstrSQL = "" Then Exit Sub
            Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)
            
            intRecord = intRecord - 1
            If .rows > 2 Then
                .RemoveItem intRow
            ElseIf .rows = 2 Then
                .rows = 3
                .RemoveItem intRow
                SetEnable
            End If
                
            If intRow < .rows - 1 Then
                .Row = intRow
            Else
                If .rows = 2 Then
                    .Row = 1
                Else
                    .Row = intRow - 1
                End If
            End If
            .Col = 0
            .ColSel = .Cols - 1
        End If
    End With
    
    mlastRow = 0
    Call vsfList_EnterCell
    staThis.Panels(2).Text = "��ǰ����" & intRecord & "�ŵ���"
    Exit Sub

errHandle:
    Exit Sub
End Sub

Private Sub mnuEditExeAmount_Click()
    '�޸�ִ������
    
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    With vsfList
        strNo = .TextMatrix(.Row, 0)
        frmDrugPlanCard.ShowCard Me, strNo, 6, blnSuccess, cboStock.ItemData(cboStock.ListIndex)
    
    End With
    If blnSuccess = True Then
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditExport_Click()
    gobjDrugPurchase.PurchasePlan gcnOracle ', UserInfo.�û�ID
End Sub

Private Sub mnuEditVerify_Click()
    '����
    
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    With vsfList
        strNo = .TextMatrix(.Row, 0)
        frmDrugPlanCard.ShowCard Me, strNo, 3, blnSuccess, cboStock.ItemData(cboStock.ListIndex)
    
    End With
    If blnSuccess = True Then
        mnuViewRefresh_Click
    End If
End Sub

Private Sub mnuEditDel_Click()
    'ɾ��
    Dim lngBillId As Long
    Dim intRow As Integer
    Dim strTitle As String
    Dim intReturn As Integer
    Dim intRecord As Integer
     
    With vsfList
        
        On Error GoTo errHandle
        intRow = .Row
        lngBillId = .TextMatrix(intRow, 1)
        intReturn = MsgBox("��ȷʵҪɾ�����ݺ�Ϊ��" & .TextMatrix(.Row, 0) & "���Ĳɹ��ƻ�����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
        intRecord = .rows - 1
        If intReturn = vbYes Then
            gstrSQL = "zl_ҩƷ�ƻ�����_DELETE('" & lngBillId & "')"
            
            If gstrSQL = "" Then Exit Sub
            Call zlDatabase.ExecuteProcedure(gstrSQL, MStrCaption)
            
            intRecord = intRecord - 1
            If .rows > 2 Then
                .RemoveItem intRow
            ElseIf .rows = 2 Then
                .rows = 3
                .RemoveItem intRow
                
                SetEnable
                
            End If
                
            If intRow < .rows - 1 Then
                .Row = intRow
            Else
                If .rows = 2 Then
                    .Row = 1
                Else
                    .Row = intRow - 1
                End If
            End If
            .Col = 0
            .ColSel = .Cols - 1
        End If
    End With
    
    mlastRow = 0
    Call vsfList_EnterCell
    staThis.Panels(2).Text = "��ǰ����" & intRecord & "�ŵ���"
    Exit Sub

errHandle:
    Exit Sub
End Sub

Private Sub mnuEditDisplay_Click()
    '�鿴����
    
    Dim strNo As String
    With vsfList
        strNo = .TextMatrix(.Row, 0)
        frmDrugPlanCard.ShowCard Me, strNo, 4, , cboStock.ItemData(cboStock.ListIndex)
        
    End With
    
End Sub

Private Sub mnuEditModify_Click()
    '�޸�
    Dim strNo As String
    Dim blnSuccess As Boolean
    
    blnSuccess = False
    With vsfList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        strNo = .TextMatrix(.Row, 0)
        frmDrugPlanCard.ShowCard Me, strNo, 2, blnSuccess, cboStock.ItemData(cboStock.ListIndex)
        If blnSuccess = True Then
            mnuViewRefresh_Click
        End If
    End With
End Sub

Private Sub mnuFileBillPreview_Click()
    With vsfList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        'If Val(vsfDetail.TextMatrix(1, 9)) = 0 Then
        If mint�۸���ʾ = 1 Then
            '���ۼۺ��ۼ۽����ʾ
            ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1330", "zl8_bill_1330"), Me, "���ݱ��=" & .TextMatrix(.Row, 0), 1, "ReportFormat=2"
        ElseIf mint�۸���ʾ = 0 Then
            '���ɱ��ۺͳɱ������ʾ
            ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1330", "zl8_bill_1330"), Me, "���ݱ��=" & .TextMatrix(.Row, 0), 1, "ReportFormat=1"
        Else
            ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1330", "zl8_bill_1330"), Me, "���ݱ��=" & .TextMatrix(.Row, 0), 1, "ReportFormat=3"
        End If
    End With
End Sub

Private Sub mnuFileBillPrint_Click()
    With vsfList
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        'If Val(vsfDetail.TextMatrix(1, 9)) = 0 Then
        If mint�۸���ʾ = 1 Then
            '���ۼۺ��ۼ۽����ʾ
            ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1330", "zl8_bill_1330"), Me, "���ݱ��=" & .TextMatrix(.Row, 0), 2, "ReportFormat=2"
        ElseIf mint�۸���ʾ = 0 Then
            '���ɱ��ۺͳɱ������ʾ
            ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1330", "zl8_bill_1330"), Me, "���ݱ��=" & .TextMatrix(.Row, 0), 2, "ReportFormat=1"
        Else
            ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1330", "zl8_bill_1330"), Me, "���ݱ��=" & .TextMatrix(.Row, 0), 2, "ReportFormat=3"
        End If
    End With
End Sub

Private Sub mnuFileExcel_Click()
    '�����Excel
    
    If Me.ActiveControl Is vsfList Then
        vsfList.Redraw = flexRDNone
        subPrint 3
        vsfList.Redraw = flexRDDirect
        vsfList.Col = 0
        vsfList.ColSel = vsfList.Cols - 1
    ElseIf Me.ActiveControl Is vsfDetail Then
        vsfDetail.Redraw = flexRDNone
        subPrint 3
        vsfDetail.Redraw = flexRDDirect
        vsfDetail.Col = 0
        vsfDetail.ColSel = vsfDetail.Cols - 1
    End If
    
End Sub

Private Sub mnufileexit_Click()
    '�˳�
    Unload Me
End Sub

Private Sub mnuFileParameter_Click()
    '��������
    Dim strDept As String
    Dim strTemp As String
    Dim i As Integer
    Dim dateCurrentDate As Date
    Dim int��ѯ���� As Integer
    
    frm��������.���ò��� Me, mstrPrivs, mlngMode, MStrCaption
    mint�۸���ʾ = Val(zlDatabase.GetPara("�۸���ʾ��ʽ", glngSys, ģ���.ҩƷ�ƻ�))
    mlastRow = 0
    
    mintPlanPoint = Val(zlDatabase.GetPara("ȫԺ�ƻ�����վ��", glngSys, mlngMode, 0))
    With cboStock
        If mintPlanPoint = 1 Or (gstrNodeNo = "-" Or gstrNodeNo = "0") Then
            strDept = ""
            For i = 0 To .ListCount - 1
                If .List(i) <> "ȫԺ" Then
                    strDept = strDept & .ItemData(i) & "," & .List(i) & "|"
                End If
            Next
            
            If strDept <> "" Then
                .Clear
                
                .AddItem "ȫԺ"
                .ItemData(.NewIndex) = 0
                
                For i = 0 To UBound(Split(strDept, "|")) - 1
                    .AddItem Mid(Split(strDept, "|")(i), InStr(1, Split(strDept, "|")(i), ",") + 1)
                    
                    .ItemData(.NewIndex) = Mid(Split(strDept, "|")(i), 1, InStr(1, Split(strDept, "|")(i), ",") - 1)
                    If Mid(Split(strDept, "|")(i), 1, InStr(1, Split(strDept, "|")(i), ",") - 1) = UserInfo.����ID Then
                        .ListIndex = .NewIndex
                        mlng�ⷿID = .ItemData(.NewIndex)
                    End If
                Next
            End If
            
            For i = 0 To .ListCount - 1
                If .ItemData(i) = mlng�ⷿID Then
                    .ListIndex = i
                End If
            Next
        Else
            For i = 0 To .ListCount - 1
                If .List(i) = "ȫԺ" Then
                    .RemoveItem i
                    .ListIndex = 0
                    Exit For
                End If
            Next
        End If
    End With
    
    Call GetDrugDigit(mlng�ⷿID, MStrCaption, mintUnit, mintShowCostDigit, mintShowPriceDigit, mintShowNumberDigit, mintShowMoneyDigit)
        
    '��֯��ʽ����
    mstrCostFormat = "'999999999990." & String(mintShowCostDigit, "0") & "'"
    mstrPriceFormat = "'999999999990." & String(mintShowPriceDigit, "0") & "'"
    mstrNumberFormat = "'999999999990." & String(mintShowNumberDigit, "0") & "'"
    mstrMoneyFormat = "'999999999990." & String(mintShowMoneyDigit, "0") & "'"
    
    dateCurrentDate = zlDatabase.Currentdate
    
    int��ѯ���� = Val(zlDatabase.GetPara("��ѯ����", glngSys, mlngMode, 7)) - 1
    strStart = Format(DateAdd("d", -int��ѯ����, dateCurrentDate), "yyyy-MM-dd")
    strEnd = Format(dateCurrentDate, "yyyy-MM-dd")
    
    SQLCondition.date����ʱ�俪ʼ = CDate(Format(strStart, "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date����ʱ����� = CDate(Format(strEnd, "yyyy-mm-dd") & " 23:59:59")
    
    Call GetList(mstrFind)
End Sub

Private Sub mnuFilePreView_Click()
    '��ӡԤ��
    vsfList.Redraw = flexRDNone
    subPrint 2
    vsfList.Redraw = flexRDDirect
    vsfList.Col = 0
    vsfList.ColSel = vsfList.Cols - 1
    
End Sub

Private Sub mnuFilePrint_Click()
    '��ӡ
    vsfList.Redraw = flexRDNone
    subPrint 1
    vsfList.Redraw = flexRDDirect
    vsfList.Col = 0
    vsfList.ColSel = vsfList.Cols - 1
End Sub

Private Sub mnuFilePrintSet_Click()
    '��ӡ����
    zlPrintSet
End Sub

Private Sub mnuHelpAbout_Click()
    '����
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
    '��������
'    ReportMan gcnOracle, Me
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub mnuHelpWebHome_Click()
    '������ҳ
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    '���ͷ���
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    'Ĭ�ϲ������ⷿ=�ⷿid����ʼʱ��=���ƿ�ʼʱ�䣬����ʱ��=���ƽ���ʱ�䣬NO=�ƻ���NO
    Dim str��ʼʱ�� As String
    Dim str����ʱ�� As String
    Dim strNo As String
    
    If vsfList.Row >= 1 And LTrim(vsfList.TextMatrix(vsfList.Row, 0)) <> "" Then
        strNo = vsfList.TextMatrix(vsfList.Row, 0)
    End If
    
    str��ʼʱ�� = IIf(Format(SQLCondition.date����ʱ�俪ʼ, "yyyy-mm-dd") = "1899-12-30", "", Format(SQLCondition.date����ʱ�俪ʼ, "yyyy-mm-dd"))
    str����ʱ�� = IIf(Format(SQLCondition.date����ʱ�����, "yyyy-mm-dd") = "1899-12-30", "", Format(SQLCondition.date����ʱ�����, "yyyy-mm-dd"))
    
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
        "�ⷿ=" & IIf(Val(cboStock.ItemData(cboStock.ListIndex)) = 0, "", Val(cboStock.ItemData(cboStock.ListIndex))), _
        "��ʼʱ��=" & str��ʼʱ��, _
        "����ʱ��=" & str����ʱ��, _
        "NO=" & strNo)
End Sub

Private Sub mnuViewColDefine_Click()
    Dim strColumn_All As String, strColumn_Select As String, strColumn_UnSelect As String
    Dim strѡ���� As String
    Dim str������ As String
    Dim strAllCol As String
    Dim arr����, arr������
    
    On Error Resume Next
    
    Select Case mlngMode
    Case ģ���.ҩƷ�ƻ�           'ҩƷ�ƻ�����
        strColumn_All = "ҩ��,0|��Ʒ��,0|ҩƷ��Դ,1|���,1|����,0|��λ,1|ҽ������,1|ǰ������,1|��������,1|�������,1|�������,1|" & _
                        "�������,0|�����,1|��������,1|��������,1|�ƻ�����,0|ִ������,0|�ͻ���λ,1|�ͻ�����,1|�ɱ���,0|�ɱ����,0|�ۼ�,0|�ۼ۽��,0|�ϴι�Ӧ��,1|˵��,1"
        strѡ���� = "ҩ��|��Ʒ��|ҩƷ��Դ|���|����|��λ|ҽ������|ǰ������|��������|�������|�������|�������|�����|��������|��������|�ƻ�����|ִ������|�ͻ���λ|�ͻ�����|�ɱ���|�ɱ����|�ۼ�|�ۼ۽��|�ϴι�Ӧ��|˵��"
        str������ = ""
    End Select
    
    'ȡ��ѡ���е���Ϣ
    strColumn_Select = zlDatabase.GetPara("ѡ����", glngSys, mlngMode, "")
    strColumn_UnSelect = zlDatabase.GetPara("������", glngSys, mlngMode, "")
    
    If strColumn_Select <> "" Then
        If strColumn_UnSelect <> "" Then
            strAllCol = strColumn_Select & "|" & strColumn_UnSelect
        Else
            strAllCol = strColumn_Select
        End If
        arr���� = Split(strColumn_All, "|")
        arr������ = Split(strAllCol, "|")
        If UBound(arr����) <> UBound(arr������) Then
            strColumn_Select = "ҩ��|��Ʒ��|ҩƷ��Դ|���|����|��λ|ҽ������|ǰ������|��������|�������|�������|�������|�����|��������|��������|�ƻ�����|ִ������|�ͻ���λ|�ͻ�����|�ɱ���|�ɱ����|�ۼ�|�ۼ۽��|�ϴι�Ӧ��|˵��"
            strColumn_UnSelect = ""
            zlDatabase.SetPara "ѡ����", strColumn_Select, glngSys, ģ���.ҩƷ�ƻ�
            zlDatabase.SetPara "������", strColumn_UnSelect, glngSys, ģ���.ҩƷ�ƻ�
        End If
    Else
        strColumn_Select = "ҩ��|��Ʒ��|ҩƷ��Դ|���|����|��λ|ҽ������|ǰ������|��������|�������|�������|�������|�����|��������|��������|�ƻ�����|ִ������|�ͻ���λ|�ͻ�����|�ɱ���|�ɱ����|�ۼ�|�ۼ۽��|�ϴι�Ӧ��|˵��"
        strColumn_UnSelect = ""
        zlDatabase.SetPara "ѡ����", strColumn_Select, glngSys, ģ���.ҩƷ�ƻ�
        zlDatabase.SetPara "������", strColumn_UnSelect, glngSys, ģ���.ҩƷ�ƻ�
    End If
    
    If Not frm������.ShowME(Me, strColumn_All, strColumn_Select) Then Exit Sub
    
    zlDatabase.SetPara "ѡ����", Split(strColumn_Select, "||")(0), glngSys, mlngMode
    zlDatabase.SetPara "������", Split(strColumn_Select, "||")(1), glngSys, mlngMode
End Sub

Private Sub mnuViewRefresh_Click()
    'ˢ��
    mlastRow = 0
    GetList mstrFind
End Sub

Private Sub mnuViewSearch_Click()
    '����
    Dim StrFind As String
    
    StrFind = FrmDrugPlanSearch.GetSearch(Me, strStart, strEnd, strVerifyStart, strVerifyEnd, _
                SQLCondition.strNO��ʼ, _
                SQLCondition.strNO����, _
                SQLCondition.date����ʱ�俪ʼ, _
                SQLCondition.date����ʱ�����, _
                SQLCondition.date���ʱ�俪ʼ, _
                SQLCondition.date���ʱ�����, _
                SQLCondition.date����ʱ�俪ʼ, _
                SQLCondition.date����ʱ�����, _
                SQLCondition.str������, _
                SQLCondition.str�����, _
                SQLCondition.str������, _
                SQLCondition.lng�ƻ�����, _
                SQLCondition.lng���Ʒ���, _
                SQLCondition.lngҩƷ)
    
    If StrFind <> "" Then
        mstrFind = StrFind
        mlastRow = 0
        GetList mstrFind
        If Format(strStart, "yyyy-mm-dd") = "1901-01-01" And Format(strVerifyStart, "yyyy-mm-dd") = "1901-01-01" Then
            lblRange.Visible = False
        ElseIf Format(strStart, "yyyy-mm-dd") <> "1901-01-01" And Format(strVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
            lblRange.Visible = True
            lblRange = "��ѯ��Χ:�������� " & Format(strStart, "yyyy��MM��dd��") & "��" & Format(strEnd, "yyyy��MM��dd��") & "  ������� " & Format(strVerifyStart, "yyyy��MM��dd��") & "��" & Format(strVerifyEnd, "yyyy��MM��dd��")
        ElseIf Format(strStart, "yyyy-mm-dd") <> "1901-01-01" Then
            lblRange.Visible = True
            lblRange = "��ѯ��Χ:�������� " & Format(strStart, "yyyy��MM��dd��") & "��" & Format(strEnd, "yyyy��MM��dd��")
        ElseIf Format(strVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
            lblRange.Visible = True
            lblRange = "��ѯ��Χ:������� " & Format(strVerifyStart, "yyyy��MM��dd��") & "��" & Format(strVerifyEnd, "yyyy��MM��dd��")
        End If
             
    End If
    
End Sub

Private Sub mnuViewStatus_Click()
    With mnuViewStatus
        .Checked = Not .Checked  ' Xor True
        staThis.Visible = .Checked
    End With
    
    Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    With mnuViewToolButton
        .Checked = Not .Checked   ' Xor True
        cbrTool.Bands(1).Visible = .Checked
        mnuViewToolText.Enabled = .Checked
    End With
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intCount As Integer      '����������
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked   ' Xor True
    With tlbTool.Buttons
        If mnuViewToolText.Checked = False Then
            'ȡ�����е��ı���ǩ��ʾ
            For intCount = 1 To .count
                .Item(intCount).Caption = ""
            Next
        Else
            '�����е��ı���ǩ��ʾ��˵����Tag�зŵ��ı���ǩ
            For intCount = 1 To .count
                .Item(intCount).Caption = .Item(intCount).Tag
            Next
        End If
    End With
    
    cbrTool.Bands(1).MinHeight = tlbTool.Height
    
    Form_Resize
End Sub

Private Sub vsfDetail_GotFocus()
    Call SetGridFocus(vsfDetail, True)
End Sub

Private Sub vsfDetail_LostFocus()
    Call SetGridFocus(vsfDetail, False)
End Sub

Private Sub vsfList_DblClick()
    If mnuEditModify.Visible = False Then Exit Sub
    If mnuEditModify.Enabled = False Then Exit Sub
    If vsfList.MouseRow = 0 Then Exit Sub
    mnuEditModify_Click
End Sub

Private Sub vsfList_EnterCell()
    Dim rsDetail As New Recordset
    Dim strSqlҩ�� As String
    Dim n As Integer
    Dim intCol As Integer
    Dim strUnit As String

    If mlastRow = vsfList.Row Then Exit Sub
    mlastRow = vsfList.Row
    
    On Error GoTo errHandle
    If vsfList.Row >= 1 And LTrim(vsfList.TextMatrix(vsfList.Row, 0)) <> "" Then
        vsfList.Col = 0
        vsfList.ColSel = vsfList.Cols - 1

        If gintҩƷ������ʾ = 0 Then
            strSqlҩ�� = ",('['||D.����||']'||D.ͨ����) AS ҩƷ��Ϣ"
        ElseIf gintҩƷ������ʾ = 1 Then
            strSqlҩ�� = ",('['||D.����||']'||NVL(D.��Ʒ��,D.ͨ����)) AS ҩƷ��Ϣ"
        Else
            strSqlҩ�� = ",('['||D.����||']'||D.ͨ����) AS ҩƷ��Ϣ,D.��Ʒ��"
        End If
        Select Case mintUnit '��λϵ����1-�ۼ�;2-����;3-סԺ;4-ҩ��
        Case 1
            gstrSQL = "decode(d.�ͻ���λ,null,'',d.�ͻ���λ|| '(1'||d.�ͻ���λ||'='||d.�ͻ���װ/1|| d.���㵥λ ||')') as �ͻ���λ,to_char(b.�ͻ�����,'999999999990.0') as �ͻ�����,"
            strUnit = "1,"
        Case 2
            gstrSQL = "decode(d.�ͻ���λ,null,'',d.�ͻ���λ|| '(1'||d.�ͻ���λ||'='||d.�ͻ���װ/d.�����װ|| d.���ﵥλ ||')') as �ͻ���λ,to_char(b.�ͻ�����,'999999999990.0') as �ͻ�����,"
            strUnit = "d.�����װ,"
        Case 3
            gstrSQL = "decode(d.�ͻ���λ,null,'',d.�ͻ���λ|| '(1'||d.�ͻ���λ||'='||d.�ͻ���װ/d.סԺ��װ|| d.סԺ��λ ||')') as �ͻ���λ,to_char(b.�ͻ�����,'999999999990.0') as �ͻ�����,"
            strUnit = "d.סԺ��װ,"
        Case Else
            gstrSQL = "decode(d.�ͻ���λ,null,'',d.�ͻ���λ|| '(1'||d.�ͻ���λ||'='||d.�ͻ���װ/d.ҩ���װ|| d.ҩ�ⵥλ ||')') as �ͻ���λ,to_char(b.�ͻ�����,'999999999990.0') as �ͻ�����,"
            strUnit = "d.ҩ���װ,"
        End Select
        
        gstrSQL = "SELECT B.���" & strSqlҩ�� & ",D.ҩƷ��Դ,D.���, Decode(" & mintUnit & ", 1, d.���㵥λ, 2, d.���ﵥλ, 3, d.סԺ��λ, d.ҩ�ⵥλ) As ��λ,d.ҽ������," & _
                " TRIM(TO_CHAR(B.ǰ������ / " & strUnit & mstrNumberFormat & ")) ǰ������," & _
                " TRIM(TO_CHAR(B.�������� / " & strUnit & mstrNumberFormat & ")) ��������," & _
                " TRIM(TO_CHAR(B.������� / " & strUnit & mstrNumberFormat & ")) �������," & _
                " Trim(To_Char(b.������� * b.�ۼ�, " & mstrMoneyFormat & ")) �����," & _
                " TRIM(TO_CHAR(B.�������� / " & strUnit & mstrNumberFormat & ")) ��������," & _
                " TRIM(TO_CHAR(B.�������� / " & strUnit & mstrNumberFormat & ")) ��������," & _
                " TRIM(TO_CHAR(B.�ƻ����� / " & strUnit & mstrNumberFormat & ")) �ƻ�����," & _
                gstrSQL & _
                " Trim(To_Char(B.���� * " & strUnit & mstrCostFormat & ")) �ɱ���," & _
                " Trim(To_Char(B.���, " & mstrMoneyFormat & ")) �ɱ����, " & _
                " Trim(To_Char(B.�ۼ� * " & strUnit & mstrPriceFormat & ")) �ۼ�, " & _
                " Trim(To_Char(B.�ۼ۽��, " & mstrMoneyFormat & ")) �ۼ۽��, " & _
                " B.�ϴι�Ӧ��,B.�ϴ�������,NVL(B.˵��,'') ˵��, " & _
                " TRIM(TO_CHAR(B.ִ������ / " & strUnit & mstrNumberFormat & ")) ִ������ " & _
                " FROM ҩƷ�ɹ��ƻ� A, ҩƷ�ƻ����� B,���ű� C," & _
                "     (SELECT DISTINCT A.ҩƷID, F.����,F.���� As ͨ����,B.���� As ��Ʒ��,f.�������� As ҽ������,A.ҩƷ��Դ,f.���㵥λ,A.סԺ��װ,A.�����װ,A.ҩ���װ," & _
                "      F.���, a.���ﵥλ,a.סԺ��λ,A.ҩ�ⵥλ,a.�ͻ���λ,a.�ͻ���װ " & _
                "     FROM ҩƷ��� A, �շ���Ŀ���� B, �շ���ĿĿ¼ F " & _
                "     WHERE A.ҩƷID = B.�շ�ϸĿID(+) AND B.����(+)=3 " & _
                "     AND A.ҩƷID = F.ID) D " & _
                " WHERE A.ID = B.�ƻ�ID AND NVL(A.�ⷿID,0)=C.ID(+) " & _
                " AND B.ҩƷID=D.ҩƷID AND B.�ƻ�ID = [1] " & IIf(SQLCondition.lngҩƷ > 0, " And B.ҩƷID=[2] ", "") & _
                " ORDER BY ���"
        Set rsDetail = zlDatabase.OpenSQLRecord(gstrSQL, MStrCaption, Val(vsfList.TextMatrix(vsfList.Row, 1)), SQLCondition.lngҩƷ)
        
        vsfDetail.Redraw = flexRDNone
        
        Set vsfDetail.DataSource = rsDetail
        rsDetail.Close
        
        With vsfDetail
            .Row = IIf(.rows > 1, 1, 0)
            .Col = 0
            .ColSel = .Cols - 1
            
            For n = 0 To .Cols - 1
                .ColKey(n) = .TextMatrix(0, n)
                .FixedAlignment(n) = flexAlignCenterCenter
            Next
            
            If Trim(vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("���Ʒ���"))) = "�Զ���������շ�" Then
                .TextMatrix(0, .ColIndex("ǰ������")) = "��������"
                .TextMatrix(0, .ColIndex("��������")) = "��������"
                .TextMatrix(0, .ColIndex("��������")) = "��������"
                .TextMatrix(0, .ColIndex("��������")) = "��������"
            End If
            
            .Redraw = flexRDDirect
        End With
    ElseIf LTrim(vsfList.TextMatrix(vsfList.Row, 0)) = "" Then
        With vsfDetail
            .Redraw = flexRDNone
            .Cols = IIf(gintҩƷ������ʾ = 2, 22, 21)
            .rows = 2
            .Clear
            
            intCol = 0
            
            .TextMatrix(0, intCol) = "���": intCol = intCol + 1
            .TextMatrix(0, intCol) = "ҩƷ��Ϣ": intCol = intCol + 1
            
            If gintҩƷ������ʾ = 2 Then
                .TextMatrix(0, intCol) = "��Ʒ��": intCol = intCol + 1
            End If
            
            .TextMatrix(0, intCol) = "ҩƷ��Դ": intCol = intCol + 1
            .TextMatrix(0, intCol) = "���": intCol = intCol + 1
            .TextMatrix(0, intCol) = "��λ": intCol = intCol + 1
            .TextMatrix(0, intCol) = "ҽ������": intCol = intCol + 1
            .TextMatrix(0, intCol) = "ǰ������": intCol = intCol + 1
            .TextMatrix(0, intCol) = "��������": intCol = intCol + 1
            .TextMatrix(0, intCol) = "�������": intCol = intCol + 1
            .TextMatrix(0, intCol) = "�����": intCol = intCol + 1
            .TextMatrix(0, intCol) = "��������": intCol = intCol + 1
            .TextMatrix(0, intCol) = "��������": intCol = intCol + 1
            .TextMatrix(0, intCol) = "�ƻ�����": intCol = intCol + 1
            .TextMatrix(0, intCol) = "�ɱ���": intCol = intCol + 1
            .TextMatrix(0, intCol) = "�ɱ����": intCol = intCol + 1
            .TextMatrix(0, intCol) = "�ۼ�": intCol = intCol + 1
            .TextMatrix(0, intCol) = "�ۼ۽��": intCol = intCol + 1
            .TextMatrix(0, intCol) = "�ϴι�Ӧ��": intCol = intCol + 1
            .TextMatrix(0, intCol) = "�ϴ�������": intCol = intCol + 1
            .TextMatrix(0, intCol) = "˵��": intCol = intCol + 1
            .TextMatrix(0, intCol) = "ִ������": intCol = intCol + 1

            .Row = 1
            .Col = 0
            .ColSel = .Cols - 1
            
            For n = 0 To .Cols - 1
                .ColKey(n) = .TextMatrix(0, n)
                .FixedAlignment(n) = flexAlignCenterCenter
            Next
            
            .Redraw = flexRDDirect
        End With
    End If
    SetDetailColWidth
    SetEnable
    
    With vsfDetail
        If .rows <= 1 Then Exit Sub
        If .TextMatrix(1, 0) <> "" Then
            If mint�۸���ʾ = 0 Then
                vsfDetail.ColWidth(.ColIndex("�ۼ�")) = 0
                vsfDetail.ColWidth(.ColIndex("�ۼ۽��")) = 0
            ElseIf mint�۸���ʾ = 1 Then
                vsfDetail.ColWidth(.ColIndex("�ɱ���")) = 0
                vsfDetail.ColWidth(.ColIndex("�ɱ����")) = 0
            End If
        End If
        If mblnViewCost = False Then
            .ColWidth(.ColIndex("�ɱ���")) = 0
            .ColWidth(.ColIndex("�ɱ����")) = 0
        End If
    End With
    vsfDetail.Row = 1
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetDetailColWidth()
    With vsfDetail
        .ColWidth(.ColIndex("���")) = 500
        .ColWidth(.ColIndex("ҩƷ��Ϣ")) = 1500
        .ColWidth(.ColIndex("ҩƷ��Դ")) = 1000
        .ColWidth(.ColIndex("���")) = 800
        .ColWidth(.ColIndex("��λ")) = 800
        .ColWidth(.ColIndex("ǰ������")) = 1200
        .ColWidth(.ColIndex("��������")) = 1200
        .ColWidth(.ColIndex("�������")) = 1200
        .ColWidth(.ColIndex("�����")) = 1200
        .ColWidth(.ColIndex("��������")) = 1200
        .ColWidth(.ColIndex("��������")) = 1200
        .ColWidth(.ColIndex("�ƻ�����")) = 1200
        .ColWidth(.ColIndex("�ɱ���")) = 1200
        .ColWidth(.ColIndex("�ɱ����")) = 1200
        .ColWidth(.ColIndex("�ۼ�")) = 1200
        .ColWidth(.ColIndex("�ۼ۽��")) = 1200
        .ColWidth(.ColIndex("�ϴι�Ӧ��")) = 1200
        .ColWidth(.ColIndex("�ϴ�������")) = 1200
        .ColWidth(.ColIndex("˵��")) = 1200
        .ColWidth(.ColIndex("ִ������")) = 1200
        .ColAlignment(.ColIndex("ǰ������")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("��������")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("�������")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("�����")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("��������")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("��������")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("�ƻ�����")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("�ɱ���")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("�ɱ����")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("�ۼ�")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("�ۼ۽��")) = flexAlignRightCenter
        .ColAlignment(.ColIndex("ִ������")) = flexAlignRightCenter
        If .TextMatrix(1, 0) <> "" Then
            '0:��ʾ�ɱ���;  1:��ʾ�ۼ�;  2:��ʾ�ɱ����ۼ�
            If mint�۸���ʾ = 0 Then
                .ColWidth(.ColIndex("�ۼ�")) = 0
                .ColWidth(.ColIndex("�ۼ۽��")) = 0
            ElseIf mint�۸���ʾ = 1 Then
                .ColWidth(.ColIndex("�ɱ���")) = 0
                .ColWidth(.ColIndex("�ɱ����")) = 0
            End If
        End If
        If mblnViewCost = False Then
            .ColWidth(.ColIndex("�ɱ���")) = 0
            .ColWidth(.ColIndex("�ɱ����")) = 0
        End If
    End With
End Sub

Private Sub vsfList_GotFocus()
    Call SetGridFocus(vsfList, True)
End Sub

Private Sub vsfList_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If mnuEditModify.Visible = False Then Exit Sub
        If mnuEditModify.Enabled = False Then Exit Sub
        mnuEditModify_Click
    End If
        
End Sub

Private Sub vsfList_LostFocus()
    Call SetGridFocus(vsfList, False)
End Sub

Private Sub vsfList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    If mnuEdit.Visible = False Then Exit Sub
    
    PopupMenu mnuEdit, 2
    
End Sub

Private Sub picSeparate_s_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '�ָ�������
    
    If Button <> 1 Then Exit Sub
    
    With picSeparate_s
        If .Top + y < 2000 Then Exit Sub
        If .Top + y > ScaleHeight - 2000 Then Exit Sub
        .Move .Left, .Top + y
    End With
    
    With vsfList
        .Height = picSeparate_s.Top - .Top
    End With
    
    With Cmd����
        .Top = vsfList.Top + vsfList.Height + 30
    End With
    
    With vsfDetail
        .Top = picSeparate_s.Top + picSeparate_s.Height + 100
        .Height = ScaleHeight - .Top - IIf(staThis.Visible, staThis.Height, 0)
    End With
End Sub

Private Sub tlbTool_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "PrintView"
            mnuFilePreView_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Add"
            mnuEditAddPlan_Click 0
        Case "Modify"
            mnuEditModify_Click
        Case "Delete"
            mnuEditDel_Click
        Case "Verify"
            mnuEditVerify_Click
        Case "Clear"
            mnuEditClear_Click
        Case "Cancel"
            mnuEditCancel_Click
        Case "Check"
            mnuEditCheck_Click
        Case "Search"
            mnuViewSearch_Click
        Case "Refresh"
            mnuViewRefresh_Click
        Case "Help"
            mnuHelpTitle_Click
        Case "Exit"
            mnufileexit_Click
        
    End Select
    
End Sub

'���ò˵��͹��߰�ť�Ŀ�������
Private Sub SetEnable()
    With vsfList
        .ToolTipText = ""
        If .TextMatrix(.Row, 0) = "" Or .Row = 0 Then         'û�е�
            mnuFilePreView.Enabled = False
            mnuFilePrint.Enabled = False
            mnuFileBillPreview.Enabled = False
            mnuFileBillPrint.Enabled = False
            mnuFileExcel.Enabled = False
            tlbTool.Buttons("Print").Enabled = False
            tlbTool.Buttons("PrintView").Enabled = False
            
            If mnuEditModify.Visible = True Then
                mnuEditModify.Enabled = False
                tlbTool.Buttons("Modify").Enabled = False
            End If
            If mnuEditDel.Visible = True Then
                mnuEditDel.Enabled = False
                tlbTool.Buttons("Delete").Enabled = False
            End If
            If mnuEditVerify.Visible = True Then
                mnuEditVerify.Enabled = False
                tlbTool.Buttons("Verify").Enabled = False
            End If
            
            If mnuEditClear.Visible = True Then
                mnuEditClear.Enabled = False
                tlbTool.Buttons("Clear").Enabled = False
            End If
            
            If mnuEditCancel.Visible = True Then
                mnuEditCancel.Enabled = False
                tlbTool.Buttons("Cancel").Enabled = False
            End If
            
            If mnuEditDisplay.Visible = True Then
                mnuEditDisplay.Enabled = False
            End If
            
            If mnuEditCheck.Visible = True Then
                mnuEditCheck.Enabled = False
                tlbTool.Buttons("Check").Enabled = False
            End If
            
            If mnuEditExeAmount.Visible = True Then
                mnuEditExeAmount.Enabled = False
            End If
        Else
            mnuFilePreView.Enabled = True
            mnuFilePrint.Enabled = True
            mnuFileBillPreview.Enabled = True
            mnuFileBillPrint.Enabled = True
            mnuFileExcel.Enabled = True
            tlbTool.Buttons("Print").Enabled = True
            tlbTool.Buttons("PrintView").Enabled = True
            
            If .TextMatrix(.Row, .ColIndex("�����")) = "" Then    'δ��˵�
                If mnuEditModify.Visible = True Then
                    mnuEditModify.Enabled = True
                    tlbTool.Buttons("Modify").Enabled = True
                End If
                If mnuEditDel.Visible = True Then
                    mnuEditDel.Enabled = True
                    tlbTool.Buttons("Delete").Enabled = True
                End If
                If mnuEditVerify.Visible = True Then
                    mnuEditVerify.Enabled = True
                    tlbTool.Buttons("Verify").Enabled = True
                End If
                
                If mnuEditClear.Visible = True Then
                    mnuEditClear.Enabled = False
                    tlbTool.Buttons("Clear").Enabled = False
                End If
                
                If mnuEditCancel.Visible = True Then
                    mnuEditCancel.Enabled = False
                    tlbTool.Buttons("Cancel").Enabled = False
                End If
            
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
                
                If mnuEditCheck.Visible = True Then
                    mnuEditCheck.Enabled = False
                    tlbTool.Buttons("Check").Enabled = False
                End If
                
                If mnuEditExeAmount.Visible = True Then
                    mnuEditExeAmount.Enabled = False
                End If
            Else    '��˵�
                If mnuEditModify.Visible = True Then
                    mnuEditModify.Enabled = False
                    tlbTool.Buttons("Modify").Enabled = False
                End If
                If mnuEditDel.Visible = True Then
                    mnuEditDel.Enabled = False
                    tlbTool.Buttons("Delete").Enabled = False
                End If
                If mnuEditVerify.Visible = True Then
                    mnuEditVerify.Enabled = False
                    tlbTool.Buttons("Verify").Enabled = False
                End If
                
                If mnuEditDisplay.Visible = True Then
                    mnuEditDisplay.Enabled = True
                End If
                
                If mnuEditClear.Visible = True Then
                    mnuEditClear.Enabled = True
                    tlbTool.Buttons("Clear").Enabled = True
                End If
                
                If mnuEditExeAmount.Visible = True Then
                    mnuEditExeAmount.Enabled = True
                End If
                
                If .TextMatrix(.Row, .ColIndex("������")) = "" Then    'δ���˵�
                    If mnuEditCancel.Visible = True And IsHavePrivs(mstrPrivs, "ȡ�����") Then
                        mnuEditCancel.Enabled = True
                        tlbTool.Buttons("Cancel").Enabled = True
                    End If
                    
                    If mnuEditCheck.Visible = True Then
                        mnuEditCheck.Enabled = True
                        tlbTool.Buttons("Check").Enabled = True
                    End If
                Else
                    '�Ѹ��˵�
                    If mnuEditCancel.Visible = True And IsHavePrivs(mstrPrivs, "ȡ������") Then
                        mnuEditCancel.Enabled = True
                        tlbTool.Buttons("Cancel").Enabled = True
                    End If
                    
                    If mnuEditCheck.Visible = True Then
                        mnuEditCheck.Enabled = False
                        tlbTool.Buttons("Check").Enabled = False
                    End If
                End If
            End If
        End If
        
    End With
End Sub

Private Sub subPrint(bytMode As Byte)
'����:���д�ӡ,Ԥ���������EXCEL
'����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
'    If gstrUserName = "" Then Call GetUserInfo
    Dim objPrint As Object
    Dim objRow As New zlTabAppRow
    Dim strRange As String
    
    If Format(strStart, "yyyy-mm-dd") = "1901-01-01" And Format(strVerifyStart, "yyyy-mm-dd") = "1901-01-01" Then
        strRange = "������� " & Format(strVerifyStart, "yyyy��MM��dd��") & "��" & Format(strVerifyEnd, "yyyy��MM��dd��")
    ElseIf Format(strVerifyStart, "yyyy-mm-dd") <> "1901-01-01" Then
        strRange = "�������� " & Format(strStart, "yyyy��MM��dd��") & "��" & Format(strEnd, "yyyy��MM��dd��") & "  ������� " & Format(strVerifyStart, "yyyy��MM��dd��") & "��" & Format(strVerifyEnd, "yyyy��MM��dd��")
    Else
        strRange = "�������� " & Format(strStart, "yyyy��MM��dd��") & "��" & Format(strEnd, "yyyy��MM��dd��")
    End If
    
    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "����_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
    objPrint.Title.Text = MStrCaption
        
    objRow.Add "ʱ�䣺" & strRange
    objRow.Add "���ţ�" & cboStock.Text
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow

    objRow.Add "��ӡ��:" & UserInfo.�û�����
    objRow.Add "��ӡ����:" & Format(zlDatabase.Currentdate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    
    If Me.ActiveControl Is vsfDetail Then
        Set objPrint.Body = vsfDetail
    Else
        Set objPrint.Body = vsfList
    End If
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub tlbTool_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "Payment"
'            mnuEditAddPayment_Click
        Case "Imprest"
'            mnuEditAddImprest_Click
        Case "next"
            mnuEditAddPlan_Click 0
        Case "period"
            mnuEditAddPlan_Click 1
    End Select
End Sub

Private Sub tlbTool_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool
    End If
End Sub

'Ѱ����ĳһ����ȵ���
Public Function FindRow(ByVal FlexTemp As MSHFlexGrid, ByVal intTemp As Variant, ByVal intCol As Integer) As Integer
    Dim i As Integer
    
    With FlexTemp
        For i = 1 To .rows - 1
            If IsDate(intTemp) Then
               If Format(.TextMatrix(i, intCol), "yyyy-mm-dd") = Format(intTemp, "yyyy-mm-dd") Then
                  FindRow = i
                  Exit Function
               End If
            Else
                If .TextMatrix(i, intCol) = intTemp Then
                  FindRow = i
                  Exit Function
                End If
            End If
        Next
    End With
    FindRow = 1
End Function


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

