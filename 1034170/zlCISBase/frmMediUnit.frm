VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMediUnit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�б굥λ"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   Icon            =   "frmMediUnit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin MSComctlLib.ImageList imgList 
      Left            =   5160
      Top             =   8520
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
            Picture         =   "frmMediUnit.frx":1CFA
            Key             =   "ItemUse"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediUnit.frx":2294
            Key             =   "ItemStop"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   2790
      Left            =   0
      TabIndex        =   0
      Top             =   8520
      Visible         =   0   'False
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   4921
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf�б굥λѡ�� 
      Height          =   2565
      Left            =   5400
      TabIndex        =   1
      Top             =   8520
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4524
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin TabDlg.SSTab sstCustom 
      Height          =   8295
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   14631
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "��ҩƷ�����б굥λ(&1)"
      TabPicture(0)   =   "frmMediUnit.frx":282E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "imgNote"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblSpec"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblnote"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblMedi"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "vsfͣ��(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "vsfUnit"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chk��ʾͣ��(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdStartAll(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdMedi"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdSave(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdRestore(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdStopAll(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdClose(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtMedi"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "���б굥λ����ҩƷ(&2)"
      TabPicture(1)   =   "frmMediUnit.frx":284A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdStartAll(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdSave(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdRestore(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdStopAll(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdClose(1)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmd�б굥λ"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txt�б굥λ"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "chk��ʾͣ��(1)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "vsf�б굥λ"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "vsfͣ��(1)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "imgNotes"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "lblnotes"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "lbl�б굥λ"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).ControlCount=   13
      Begin VB.CommandButton cmdStartAll 
         Caption         =   "ͣ��ȫ��(&B)"
         Height          =   350
         Index           =   1
         Left            =   -73560
         TabIndex        =   23
         Top             =   7740
         Width           =   1245
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "����(&S)"
         Height          =   350
         Index           =   1
         Left            =   -70080
         TabIndex        =   22
         Top             =   7740
         Width           =   1095
      End
      Begin VB.CommandButton cmdRestore 
         Caption         =   "ȫ���ָ�(&R)"
         Height          =   350
         Index           =   1
         Left            =   -74880
         TabIndex        =   21
         Top             =   7740
         Width           =   1245
      End
      Begin VB.CommandButton cmdStopAll 
         Caption         =   "ͣ��ȫѡ(&A)"
         Height          =   350
         Index           =   1
         Left            =   -72240
         TabIndex        =   20
         Top             =   7740
         Width           =   1245
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "�ر�(&C)"
         Height          =   350
         Index           =   1
         Left            =   -68880
         TabIndex        =   19
         Top             =   7740
         Width           =   1100
      End
      Begin VB.CommandButton cmd�б굥λ 
         Caption         =   "��"
         Height          =   285
         Left            =   -68160
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1185
         Width           =   285
      End
      Begin VB.TextBox txt�б굥λ 
         Height          =   300
         Left            =   -73740
         MaxLength       =   50
         TabIndex        =   17
         Top             =   1170
         Width           =   5580
      End
      Begin VB.CheckBox chk��ʾͣ�� 
         Caption         =   "��ʾͣ��"
         Height          =   255
         Index           =   1
         Left            =   -68880
         TabIndex        =   16
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtMedi 
         Height          =   300
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1170
         Width           =   5580
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "�ر�(&C)"
         Height          =   350
         Index           =   0
         Left            =   6120
         TabIndex        =   9
         Top             =   7740
         Width           =   1100
      End
      Begin VB.CommandButton cmdStopAll 
         Caption         =   "ͣ��ȫѡ(&A)"
         Height          =   350
         Index           =   0
         Left            =   2760
         TabIndex        =   8
         Top             =   7740
         Width           =   1245
      End
      Begin VB.CommandButton cmdRestore 
         Caption         =   "ȫ���ָ�(&R)"
         Height          =   350
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   7740
         Width           =   1245
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "����(&S)"
         Height          =   350
         Index           =   0
         Left            =   4920
         TabIndex        =   6
         Top             =   7740
         Width           =   1100
      End
      Begin VB.CommandButton cmdMedi 
         Caption         =   "��"
         Height          =   285
         Left            =   6840
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1185
         Width           =   285
      End
      Begin VB.CommandButton cmdStartAll 
         Caption         =   "ͣ��ȫ��(&B)"
         Height          =   350
         Index           =   0
         Left            =   1440
         TabIndex        =   4
         Top             =   7740
         Width           =   1245
      End
      Begin VB.CheckBox chk��ʾͣ�� 
         Caption         =   "��ʾͣ��"
         Height          =   255
         Index           =   0
         Left            =   6120
         TabIndex        =   3
         Top             =   1560
         Width           =   1095
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfUnit 
         Height          =   3255
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Width           =   7095
         _cx             =   12515
         _cy             =   5741
         Appearance      =   0
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   10329501
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmMediUnit.frx":2866
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
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
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
      Begin VSFlex8Ctl.VSFlexGrid vsfͣ�� 
         Height          =   2295
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   5280
         Visible         =   0   'False
         Width           =   7095
         _cx             =   12515
         _cy             =   4048
         Appearance      =   0
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
         BackColor       =   -2147483644
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483644
         GridColor       =   10329501
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmMediUnit.frx":2974
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
         ExplorerBar     =   0
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
         VirtualData     =   -1  'True
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
      Begin VSFlex8Ctl.VSFlexGrid vsf�б굥λ 
         Height          =   3255
         Left            =   -74880
         TabIndex        =   24
         Top             =   1920
         Width           =   7095
         _cx             =   12515
         _cy             =   5741
         Appearance      =   0
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   10329501
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmMediUnit.frx":2A0E
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
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
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
      Begin VSFlex8Ctl.VSFlexGrid vsfͣ�� 
         Height          =   2295
         Index           =   1
         Left            =   -74880
         TabIndex        =   25
         Top             =   5280
         Visible         =   0   'False
         Width           =   7095
         _cx             =   12515
         _cy             =   4048
         Appearance      =   0
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
         BackColor       =   -2147483644
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483644
         GridColor       =   10329501
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmMediUnit.frx":2B20
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
         ExplorerBar     =   0
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
         VirtualData     =   -1  'True
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
      Begin VB.Image imgNotes 
         Height          =   480
         Left            =   -74880
         Picture         =   "frmMediUnit.frx":2BBA
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblnotes 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "    ��ѡ���б굥λ��ָ�����е�ҩƷ���б�ҩƷ���ʱ���乩Ӧ�̱��������б굥λ"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   -74160
         TabIndex        =   27
         Top             =   540
         Width           =   5685
      End
      Begin VB.Label lbl�б굥λ 
         AutoSize        =   -1  'True
         Caption         =   "�б굥λ(&Z)"
         Height          =   180
         Left            =   -74880
         TabIndex        =   26
         Top             =   1230
         Width           =   990
      End
      Begin VB.Label lblMedi 
         AutoSize        =   -1  'True
         Caption         =   "ҩƷ���(&M)"
         Height          =   180
         Left            =   120
         TabIndex        =   15
         Top             =   1230
         Width           =   990
      End
      Begin VB.Label lblnote 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "    ��ѡ��ҩƷ��ָ�����б굥λ���б�ҩƷ���ʱ���乩Ӧ�̱��������б굥λ"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   840
         TabIndex        =   14
         Top             =   540
         Width           =   5685
      End
      Begin VB.Label lblSpec 
         AutoSize        =   -1  'True
         Caption         =   "���      ���ƣ�       ��λ��ƿ"
         Height          =   180
         Left            =   135
         TabIndex        =   13
         Top             =   1560
         Width           =   2970
      End
      Begin VB.Image imgNote 
         Height          =   480
         Left            =   120
         Picture         =   "frmMediUnit.frx":3484
         Top             =   360
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmMediUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lblTag As String
Public frmTag As String
Public strPrivs As String
Private mstrTemp As String
Private mobjItem As ListItem
Private mrsTemp As New ADODB.Recordset
Private mlngId As Long      '��¼ѡ��id
Private mstrԭֵ As String
Private mstrԭֵ�б굥λ As String
Private mblnSave As Boolean     '��¼�Ƿ񱣴��˽������޸ĵ�ֵ
Private mblnSave�б굥λ As Boolean
'��¼״̬����
Private Enum mStates
    ԭʼ = 0
    ���� = 1
    �޸� = 2
    ɾ�� = 3
End Enum
Private Const mcstDelColor = &HDBDBDB
Private Sub vsf_ResetSerial(Optional ByVal bln As Boolean = False)
    Dim i As Integer
    
    If bln Then
        With vsf�б굥λ
            For i = 1 To .Rows - 1
                .TextMatrix(i, 0) = i
            Next
        End With
    Else
        With vsfUnit
            For i = 1 To .Rows - 1
                .TextMatrix(i, 0) = i
            Next
        End With
    End If
End Sub

Private Sub chk��ʾͣ��_Click(Index As Integer)
    Call Resize(Index)
End Sub

Private Sub cmdClose_Click(Index As Integer)
    Dim intRow As Integer
    Dim intCol As Integer
    Dim strTemp As String
    Dim strԭֵ�б굥λ As String
    
    With vsfUnit
        For intRow = 1 To .Rows - 1
            For intCol = 1 To .Cols - 1
                If intCol = .ColIndex("ͣ��") Then
                    strTemp = strTemp & IIf(.TextMatrix(intRow, intCol) = "", 0, .TextMatrix(intRow, intCol)) & "|"
                ElseIf intCol <> .ColIndex("����") Then
                    strTemp = strTemp & .TextMatrix(intRow, intCol) & "|"
                End If
            Next
        Next
    End With

    With vsf�б굥λ
        For intRow = 1 To .Rows - 1
            For intCol = 1 To .Cols - 1
                If intCol = .ColIndex("ͣ��") Then
                    strԭֵ�б굥λ = strԭֵ�б굥λ & IIf(.TextMatrix(intRow, intCol) = "", 0, .TextMatrix(intRow, intCol)) & "|"
                ElseIf intCol <> .ColIndex("����") Then
                    strԭֵ�б굥λ = strԭֵ�б굥λ & .TextMatrix(intRow, intCol) & "|"
                End If
            Next
        Next
    End With
    
    If (strTemp <> mstrԭֵ And mblnSave = False) Or (strԭֵ�б굥λ <> mstrԭֵ�б굥λ And mblnSave�б굥λ = False) Then
        If MsgBox("��ǰ���ݱ��޸ĺ�δ���棬��ȷ��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            mblnSave = False
            mblnSave�б굥λ = False
            Unload Me
        End If
    Else
        mblnSave = False
        mblnSave�б굥λ = False
        Unload Me
    End If

End Sub

Private Sub Delete(Optional ByVal bln As Boolean = False)
    Dim i As Integer
    
    If bln Then
        With vsf�б굥λ
            If .Rows = 1 Then Exit Sub
            If Val(.TextMatrix(.Row, .ColIndex("ҩƷID"))) = 0 Then Exit Sub
            
            Select Case Val(.TextMatrix(.Row, .ColIndex("״̬")))
                Case mStates.����
                    If .Rows - 1 = 1 Then
                        For i = 1 To .Cols - 1
                            .TextMatrix(1, i) = ""
                        Next
                    Else
                        .RemoveItem .Row
                        Call vsf_ResetSerial(True)
                    End If
            End Select
        End With
    Else
        With vsfUnit
            If .Rows = 1 Then Exit Sub
            If Val(.TextMatrix(.Row, .ColIndex("��λID"))) = 0 Then Exit Sub
            
            Select Case Val(.TextMatrix(.Row, .ColIndex("״̬")))
                Case mStates.����
                    If .Rows - 1 = 1 Then
                        For i = 1 To .Cols - 1
                            .TextMatrix(1, i) = ""
                        Next
                    Else
                        .RemoveItem .Row
                        vsf_ResetSerial
                    End If
            End Select
        End With
    End If
End Sub

Private Sub cmdStartAll_Click(Index As Integer)
    Dim i As Integer
    
    If Index = 0 Then
        With vsfUnit
            If .Rows = 1 Then Exit Sub
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("ͣ��")) = 0
            Next
        End With
    ElseIf Index = 1 Then
        With vsf�б굥λ
            If .Rows = 1 Then Exit Sub
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("ͣ��")) = 0
            Next
        End With
    End If
End Sub

Private Sub cmdStopAll_Click(Index As Integer)
    Dim i As Integer
    
    If Index = 0 Then
        With vsfUnit
            If .Rows = 1 Then Exit Sub
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("ͣ��")) = 1
            Next
        End With
    ElseIf Index = 1 Then
        With vsf�б굥λ
            If .Rows = 1 Then Exit Sub
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("ͣ��")) = 1
            Next
        End With
    End If
End Sub
Private Sub cmdMedi_Click()
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select I.ID,I.����,I.����,I.���,I.����,I.���㵥λ as ��λ" & _
            " from �շ���ĿĿ¼ I" & _
            " where I.���=[1] " & _
            "       and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))"
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.Tag)
    
    With mrsTemp
        If .BOF Or .EOF = 1 Then
            MsgBox "��δ��������������ҩƷ��", vbExclamation, gstrSysName
            Me.lblMedi.Tag = 0: Me.txtMedi.Tag = "": Me.txtMedi.Text = Me.txtMedi.Tag: Me.txtMedi.SetFocus: Exit Sub
        End If
        If .RecordCount = 1 Then
            If Me.lblMedi.Tag <> !ID Then
                Me.lblMedi.Tag = !ID
                Me.txtMedi.Tag = "[" & !���� & "]" & !����
                Me.txtMedi.Text = Me.txtMedi.Tag
                If Me.Tag <> "7" Then
                    Me.lblSpec.Caption = "���" & IIf(IsNull(!���), "", !���) & _
                        "   ���ƣ�" & IIf(IsNull(!����), "", !����) & _
                        "   ��λ��" & IIf(IsNull(!��λ), "", !��λ)
                Else
                    Me.lblSpec.Caption = "���ƣ�" & IIf(IsNull(!����), "", !����) & "   ��λ��" & IIf(IsNull(!��λ), "", !��λ)
                End If
                Call ShowData
            End If
            Call OS.PressKey(vbKeyTab)
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set mobjItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !����)
            mobjItem.Icon = "ItemUse": mobjItem.SmallIcon = "ItemUse"
            mobjItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = !����
            mobjItem.SubItems(Me.lvwItems.ColumnHeaders("���").Index - 1) = IIf(IsNull(!���), "", !���)
            mobjItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = IIf(IsNull(!����), "", !����)
            mobjItem.SubItems(Me.lvwItems.ColumnHeaders("��λ").Index - 1) = IIf(IsNull(!��λ), "", !��λ)
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    With Me.lvwItems
        .Tag = Me.txtMedi.Name
        .Left = Me.txtMedi.Left + 120
        .Top = Me.txtMedi.Top + Me.txtMedi.Height + 120
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdRestore_Click(Index As Integer)
    Call ShowData(Index = 1)
End Sub

Private Sub cmdSave_Click(Index As Integer)
    Dim lngUnitId As Long
    Dim lngMediId As Long
    Dim str�б���� As String
    Dim strDelDate As String
    Dim i As Integer
    Dim strTemp As String
    Dim strContent As String
    Dim rsTemp As ADODB.Recordset
    Dim bln���� As Boolean
    
    On Error GoTo ErrHand
     
    If vsfUnit.Rows > 1 And Val(lblMedi.Tag) > 0 Then
        mblnSave = True
        lngMediId = Val(lblMedi.Tag)
        
        With vsfUnit
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, .ColIndex("��λID"))) > 0 And .TextMatrix(i, .ColIndex("��λ")) <> "" Then
                    strTemp = strTemp & "|" & .TextMatrix(i, .ColIndex("��λ"))
                    lngUnitId = Val(.TextMatrix(i, .ColIndex("��λID")))
                    str�б���� = .TextMatrix(i, .ColIndex("�б����"))
                    strDelDate = .TextMatrix(i, .ColIndex("����ʱ��"))
                    
                    If .TextMatrix(i, .ColIndex("ͣ��")) Like "*1" Then .TextMatrix(i, .ColIndex("����")) = mStates.ɾ��
                    
                    gstrSql = ""
                    Select Case Val(.TextMatrix(i, .ColIndex("����")))
                        Case mStates.����
                            gstrSql = "ZL_ҩƷ�б굥λ_INSERT(" & lngMediId & "," & lngUnitId & ", '" & str�б���� & "')"
                        Case mStates.�޸�
                            gstrSql = "Zl_ҩƷ�б굥λ_Update(" & lngMediId & "," & lngUnitId & ",to_date('" & strDelDate & "','YYYY-MM-DD HH24:MI:SS') , '" & str�б���� & "')"
                        Case mStates.ɾ��
                            gstrSql = "ZL_ҩƷ�б굥λ_DELETE(" & lngMediId & "," & lngUnitId & ",to_date('" & strDelDate & "','YYYY-MM-DD HH24:MI:SS'))"
                    End Select
                    
                    If gstrSql <> "" Then Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption): bln���� = True
                
                    'ͬ������ƽ̨ҩƷ��Ϣ
                    If Not gobjLogisticPlatform Is Nothing Then
                        If Val(.TextMatrix(i, .ColIndex("����"))) = mStates.ɾ�� Then
                            gobjLogisticPlatform.ClearDrugInfo lngMediId, lngUnitId
                        End If
                        gobjLogisticPlatform.UploadDrugInfo Me, gcnOracle, lngMediId
                    End If
                
                End If
            Next
        End With
    End If
    
    If vsf�б굥λ.Rows = 1 Or Val(lbl�б굥λ.Tag) = 0 Then
        If Index = 0 Then
            txtMedi.SetFocus
        Else
            txt�б굥λ.SetFocus
        End If
        Call ShowData
        Call ShowData(True)
        If bln���� Then MsgBox "����ɹ���", vbInformation, gstrSysName
        Exit Sub
    End If
    mblnSave�б굥λ = True
    lngUnitId = Val(lbl�б굥λ.Tag)
    
    With vsf�б굥λ
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("ҩƷID"))) > 0 And .TextMatrix(i, .ColIndex("ҩƷ���")) <> "" Then
                lngMediId = Val(.TextMatrix(i, .ColIndex("ҩƷID")))
                str�б���� = .TextMatrix(i, .ColIndex("�б����"))
                strDelDate = .TextMatrix(i, .ColIndex("����ʱ��"))
                
                If .TextMatrix(i, .ColIndex("ͣ��")) Like "*1" Then .TextMatrix(i, .ColIndex("����")) = mStates.ɾ��
                
                If Val(.TextMatrix(i, .ColIndex("����"))) = mStates.���� Then
                    gstrSql = "Select 1" & vbNewLine & _
                                    "From ҩƷ�б굥λ T" & vbNewLine & _
                                    "Where t.��λid =[1] And t.ҩƷid =[2]" & vbNewLine & _
                                    "And Sysdate Between t.����ʱ�� And Nvl(t.����ʱ��, To_Date('3000-1-1', 'YYYY-MM-DD'))"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngUnitId, lngMediId)
                    If rsTemp.RecordCount > 0 Then .TextMatrix(i, .ColIndex("����")) = mStates.ԭʼ
                End If
                
                gstrSql = ""
                Select Case Val(.TextMatrix(i, .ColIndex("����")))
                    Case mStates.����
                        gstrSql = "ZL_ҩƷ�б굥λ_INSERT(" & lngMediId & "," & lngUnitId & ", '" & str�б���� & "')"
                    Case mStates.�޸�
                        gstrSql = "Zl_ҩƷ�б굥λ_Update(" & lngMediId & "," & lngUnitId & ",to_date('" & strDelDate & "','YYYY-MM-DD HH24:MI:SS') , '" & str�б���� & "')"
                    Case mStates.ɾ��
                        gstrSql = "ZL_ҩƷ�б굥λ_DELETE(" & lngMediId & "," & lngUnitId & ",to_date('" & strDelDate & "','YYYY-MM-DD HH24:MI:SS'))"
                End Select
                
                If gstrSql <> "" Then Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption): bln���� = True
            
                'ͬ������ƽ̨ҩƷ��Ϣ
                If Not gobjLogisticPlatform Is Nothing Then
                    If Val(.TextMatrix(i, .ColIndex("����"))) = mStates.ɾ�� Then
                        gobjLogisticPlatform.ClearDrugInfo lngMediId, lngUnitId
                    End If
                    gobjLogisticPlatform.UploadDrugInfo Me, gcnOracle, lngMediId
                End If
            
            End If
        Next
    End With
    
    If Index = 0 Then
        txtMedi.SetFocus
    Else
        txt�б굥λ.SetFocus
    End If

    Call ShowData
    Call ShowData(True)
    
    If bln���� Then MsgBox "����ɹ���", vbInformation, gstrSysName
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmd�б굥λ_Click()
    Dim rsRecord As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    
    On Error GoTo ErrHand
    
    vRect = zlControl.GetControlRect(txt�б굥λ.hWnd) '��ȡλ��
    dblLeft = vRect.Left
    dblTop = vRect.Top

    gstrSql = "Select ID,����,����,���� From ��Ӧ�� Where ĩ��=1 And (instr(����,1,1)=1 Or Nvl(ĩ��,0)=0) And (����ʱ�� is null or ����ʱ��=to_date('3000-01-01','YYYY-MM-DD')) Order By ���� "
    
    Set rsRecord = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, "��Ӧ��", False, "", "", False, False, _
    True, dblLeft, dblTop, txt�б굥λ.Height, blnCancel, False, True)
    
    If rsRecord Is Nothing Then
        Exit Sub
    Else
        Me.lbl�б굥λ.Tag = rsRecord!ID
        Me.txt�б굥λ.Tag = "[" & rsRecord!���� & "]" & rsRecord!����
        Me.txt�б굥λ.Text = Me.txt�б굥λ.Tag
        Call ShowData(True)
    End If

    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        If Msf�б굥λѡ��.Visible Then
            Msf�б굥λѡ��.Visible = False
            Exit Sub
        End If
        If lvwItems.Visible Then
            lvwItems.Visible = False: txtMedi.SetFocus: Exit Sub
        End If
        Call cmdClose_Click(0)
    Case Else
    End Select
End Sub

Private Sub sstCustom_Click(PreviousTab As Integer)
    Call Resize(IIf(PreviousTab = 1, 0, 1))
End Sub

Private Sub txt�б굥λ_GotFocus()
    Me.txt�б굥λ.SelStart = 0: Me.txt�б굥λ.SelLength = 100
End Sub

Private Sub txt�б굥λ_KeyPress(KeyAscii As Integer)
    Dim rsRecord As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    
    On Error GoTo ErrHand
    
    If InStr("~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub
    
    vRect = zlControl.GetControlRect(txt�б굥λ.hWnd) '��ȡλ��
    dblLeft = vRect.Left
    dblTop = vRect.Top

    gstrSql = "Select ID,����,����,���� From ��Ӧ�� Where ĩ��=1 And (instr(����,1,1)=1 Or Nvl(ĩ��,0)=0) And (����ʱ�� is null or ����ʱ��=to_date('3000-01-01','YYYY-MM-DD')) and (���� like [1] or ���� like[1] or ���� like [1]) Order By ���� "
    
    Set rsRecord = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, "��Ӧ��", False, "", "", False, False, _
    True, dblLeft, dblTop, txt�б굥λ.Height, blnCancel, False, True, UCase(txt�б굥λ.Text) & "%")
    
    If blnCancel = True Then
        Me.txt�б굥λ.Text = Me.txt�б굥λ.Tag
        Exit Sub
    End If
    
    If rsRecord Is Nothing Then
        MsgBox "�޸��б굥λ�����������룡", vbInformation, gstrSysName
        Me.txt�б굥λ.Text = Me.txt�б굥λ.Tag
        Exit Sub
    Else
        Me.lbl�б굥λ.Tag = rsRecord!ID
        Me.txt�б굥λ.Tag = "[" & rsRecord!���� & "]" & rsRecord!����
        Me.txt�б굥λ.Text = Me.txt�б굥λ.Tag
        Call ShowData(True)
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    Dim i As Long
    
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select I.ID,I.����,I.����,I.���,I.����,I.���㵥λ as ��λ" & _
            " from �շ���ĿĿ¼ I" & _
            " where I.���=[1] and I.ID=[2] " & _
            "       and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))"
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.Tag, Val(Me.lblMedi.Tag))
    
    With mrsTemp
        If .BOF Or .EOF = 1 Then
            Me.lblMedi.Tag = 0: Me.txtMedi.Tag = "": Me.txtMedi.Text = Me.txtMedi.Tag
        Else
            Me.lblMedi.Tag = !ID
            Me.txtMedi.Tag = "[" & !���� & "]" & !����
            Me.txtMedi.Text = Me.txtMedi.Tag
            Me.lblSpec.Caption = "���" & IIf(IsNull(!���), "", !���) & _
                "   ���ƣ�" & IIf(IsNull(!����), "", !����) & _
                "   ��λ��" & IIf(IsNull(!��λ), "", !��λ)
            Call ShowData
        End If
    End With
    
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    '�в�ҩ�����������б굥λ
    On Error Resume Next
    
    Me.Tag = frmTag
    Me.lblMedi.Tag = lblTag
    frmMediUnit.Height = 6560
    sstCustom.Height = 5895
    cmdRestore(0).Top = 5340
    cmdStartAll(0).Top = 5340
    cmdStopAll(0).Top = 5340
    cmdSave(0).Top = 5340
    cmdClose(0).Top = 5340
    cmdRestore(1).Top = 5340
    cmdStartAll(1).Top = 5340
    cmdStopAll(1).Top = 5340
    cmdSave(1).Top = 5340
    cmdClose(1).Top = 5340
    
    If InStr(1, strPrivs, "�б굥λ") = 0 Then
        MsgBox "�㲻�߱������б굥λ��Ȩ�ޣ�", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    
    Me.lvwItems.ListItems.Clear
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "����", "����", 2000
        .Add , "����", "����", 1000
        .Add , "���", "���", 1200
        .Add , "����", IIf(Me.Tag = "7", "����", "����"), 1200
        .Add , "��λ", "��λ", 600
    End With
    With Me.lvwItems
        .ColumnHeaders("����").Position = 1
        .SortKey = .ColumnHeaders("����").Index - 1
        .SortOrder = lvwAscending
    End With
    With vsfUnit
        .ColComboList(.ColIndex("��λ")) = "|..."
        .Editable = flexEDKbdMouse
    End With
    
    With vsf�б굥λ
        .ColComboList(.ColIndex("ҩƷ���")) = "|..."
        .Editable = flexEDKbdMouse
    End With
    
    Call ShowData
    Call ShowData(True)
End Sub

Private Sub txt�б굥λ_LostFocus()
    Me.txt�б굥λ.Text = Me.txt�б굥λ.Tag
End Sub

Private Sub vsfUnit_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsRecord As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    
    vRect = zlControl.GetControlRect(vsfUnit.hWnd) '��ȡλ��
    dblLeft = vRect.Left + vsfUnit.CellLeft
    dblTop = vRect.Top + vsfUnit.CellTop + vsfUnit.CellHeight + 3200
    With vsfUnit
        If Col = .ColIndex("��λ") Then
            gstrSql = "Select ID,����,����,���� From ��Ӧ�� Where ĩ��=1 And (instr(����,1,1)=1 Or Nvl(ĩ��,0)=0) And (����ʱ�� is null or ����ʱ��=to_date('3000-01-01','YYYY-MM-DD')) Order By ���� "
            
            Set rsRecord = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, "��Ӧ��", False, "", "", False, False, _
            True, dblLeft, dblTop, .Height, blnCancel, False, True)
            
            If rsRecord Is Nothing Then
                Exit Sub
            Else
                mlngId = 0
                mlngId = rsRecord!ID
                If CheckDub = False Then
                    .TextMatrix(Row, 0) = Row
                    .TextMatrix(Row, .ColIndex("��λ")) = "[" & rsRecord!���� & "]" & rsRecord!����
                    .TextMatrix(Row, .ColIndex("״̬")) = mStates.����
                    .TextMatrix(Row, .ColIndex("����")) = mStates.����
                    .TextMatrix(Row, .ColIndex("��λID")) = rsRecord!ID
    '                .Cell(flexcpBackColor, Row, 0, Row, .Cols - 1) = vbWhite  'mcstInsertColor
                    .Cell(flexcpBackColor, Row, .ColIndex("����ʱ��"), Row, .ColIndex("����ʱ��")) = mcstDelColor
                    .Col = .ColIndex("�б����")
                Else
                    MsgBox "�Ѿ��и��б굥λ��", vbInformation, gstrSysName
                End If
            End If
        End If
    End With
End Sub

Private Sub txtMedi_GotFocus()
    Me.txtMedi.SelStart = 0: Me.txtMedi.SelLength = 100
End Sub

Private Sub txtMedi_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub
    mstrTemp = UCase(Trim(Me.txtMedi.Text))
    If mstrTemp = "" Then Me.lblMedi.Tag = 0: Me.txtMedi.Tag = "": Me.txtMedi.Text = "": Exit Sub
    
    If InStr(1, mstrTemp, "[") <> 0 And InStr(1, mstrTemp, "]") <> 0 Then mstrTemp = Mid(mstrTemp, 2, InStr(1, mstrTemp, "]") - 2)
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select distinct I.ID,I.����,I.����,I.���,I.����,I.���㵥λ as ��λ" & _
            " from �շ���ĿĿ¼ I,�շ���Ŀ���� N" & _
            " where I.ID=N.�շ�ϸĿID and I.���=[1] " & _
            "       and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & _
            "       and (I.���� like [2] or N.���� like [3] or N.���� like [3])"
    Set mrsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.Tag, mstrTemp & "%", gstrMatch & mstrTemp & "%")
    
    With mrsTemp
        If .BOF Or .EOF = 1 Then
            MsgBox "δ�ҵ�ָ������ҩƷ��������ָ����", vbExclamation, gstrSysName
            Me.txtMedi.Text = Me.txtMedi.Tag: Me.txtMedi.SetFocus
            Exit Sub
        End If
        If .RecordCount = 1 Then
            If Me.lblMedi.Tag <> !ID Then
                Me.lblMedi.Tag = !ID
                Me.txtMedi.Tag = "[" & !���� & "]" & !����
                Me.txtMedi.Text = Me.txtMedi.Tag
                If Me.Tag <> "7" Then
                    Me.lblSpec.Caption = "���" & IIf(IsNull(!���), "", !���) & _
                        "   ���ƣ�" & IIf(IsNull(!����), "", !����) & _
                        "   ��λ��" & IIf(IsNull(!��λ), "", !��λ)
                Else
                    Me.lblSpec.Caption = "���ƣ�" & IIf(IsNull(!����), "", !����) & "   ��λ��" & IIf(IsNull(!��λ), "", !��λ)
                End If
                Call ShowData
            End If
            Call OS.PressKey(vbKeyTab)
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set mobjItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !����)
            mobjItem.Icon = "ItemUse": mobjItem.SmallIcon = "ItemUse"
            mobjItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = !����
            mobjItem.SubItems(Me.lvwItems.ColumnHeaders("���").Index - 1) = IIf(IsNull(!���), "", !���)
            mobjItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = IIf(IsNull(!����), "", !����)
            mobjItem.SubItems(Me.lvwItems.ColumnHeaders("��λ").Index - 1) = IIf(IsNull(!��λ), "", !��λ)
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    With Me.lvwItems
        .Tag = Me.txtMedi.Name
        .Left = Me.txtMedi.Left + 120
        .Top = Me.txtMedi.Top + Me.txtMedi.Height + 120
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtMedi_LostFocus()
    Me.txtMedi.Text = Me.txtMedi.Tag
End Sub

Private Sub ShowData(Optional ByVal bln As Boolean = False)
    Dim intRow As Integer
    Dim intCol As Integer
    Dim rsTemps As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    If bln Then
        '��ʾ��ͣ�õ��б굥λ
        vsfͣ��(1).TextMatrix(1, 0) = "1"
        
        gstrSql = "Select a.Id, '[' || a.���� || ']' || a.���� || '(' || a.��� || ')'  as ҩƷ���, b.����ʱ��, b.�б����" & vbNewLine & _
                        "From �շ���ĿĿ¼ A, ҩƷ�б굥λ B, ��Ӧ�� C" & vbNewLine & _
                        "Where a.Id = b.ҩƷid And Instr(c.����, 1, 1) = 1 And b.��λid = c.Id And b.��λid =[1] And" & vbNewLine & _
                        "     Not (b.����ʱ�� Is Null Or b.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) " & vbNewLine & _
                        "Order By b.����ʱ��"
                    
        Set rsTemps = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(lbl�б굥λ.Tag))
        
        With rsTemps
            vsfͣ��(1).Rows = 1
            vsfͣ��(1).Rows = IIf(.RecordCount > 0, .RecordCount + 1, 2)
            Do While Not .EOF
                vsfͣ��(1).TextMatrix(.AbsolutePosition, 0) = .AbsolutePosition
                vsfͣ��(1).TextMatrix(.AbsolutePosition, vsfͣ��(1).ColIndex("ҩƷ���")) = !ҩƷ���
                vsfͣ��(1).TextMatrix(.AbsolutePosition, vsfͣ��(1).ColIndex("�б����")) = IIf(IsNull(!�б����), "", !�б����)
                vsfͣ��(1).TextMatrix(.AbsolutePosition, vsfͣ��(1).ColIndex("����ʱ��")) = Format(!����ʱ��, "YYYY-MM-DD HH:MM:SS")
                .MoveNext
            Loop
        End With
    
        '��ʾ�ѳ�ʼ�����б굥λ
        vsf�б굥λ.TextMatrix(1, 0) = "1"
        
        gstrSql = "Select a.Id, '[' || a.���� || ']' || a.���� || '(' || a.��� || ')'  as ҩƷ���, b.����ʱ��, b.�б����" & vbNewLine & _
                        "From �շ���ĿĿ¼ A, ҩƷ�б굥λ B, ��Ӧ�� C" & vbNewLine & _
                        "Where a.Id = b.ҩƷid And Instr(c.����, 1, 1) = 1 And b.��λid = c.Id And b.��λid =[1] And" & vbNewLine & _
                        "      (b.����ʱ�� Is Null Or b.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
                        "Order By b.����ʱ��"

        Set rsTemps = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(lbl�б굥λ.Tag))
        
        With rsTemps
            vsf�б굥λ.Rows = 1
            vsf�б굥λ.Rows = IIf(.RecordCount > 0, .RecordCount + 1, 2)
            Do While Not .EOF
                vsf�б굥λ.TextMatrix(.AbsolutePosition, 0) = .AbsolutePosition
                vsf�б굥λ.TextMatrix(.AbsolutePosition, vsf�б굥λ.ColIndex("ҩƷ���")) = !ҩƷ���
                vsf�б굥λ.TextMatrix(.AbsolutePosition, vsf�б굥λ.ColIndex("�б����")) = IIf(IsNull(!�б����), "", !�б����)
                vsf�б굥λ.TextMatrix(.AbsolutePosition, vsf�б굥λ.ColIndex("����ʱ��")) = Format(!����ʱ��, "YYYY-MM-DD HH:MM:SS")
                vsf�б굥λ.TextMatrix(.AbsolutePosition, vsf�б굥λ.ColIndex("״̬")) = 0
                vsf�б굥λ.TextMatrix(.AbsolutePosition, vsf�б굥λ.ColIndex("����")) = 0
                vsf�б굥λ.TextMatrix(.AbsolutePosition, vsf�б굥λ.ColIndex("ҩƷID")) = !ID
                .MoveNext
            Loop
        End With
        
        mstrԭֵ�б굥λ = ""
        With vsf�б굥λ
            .Cell(flexcpBackColor, 0, .ColIndex("����ʱ��"), .Rows - 1, .ColIndex("����ʱ��")) = mcstDelColor
            For intRow = 1 To .Rows - 1
                For intCol = 1 To .Cols - 1
                    If intCol = .ColIndex("ͣ��") Then
                        mstrԭֵ�б굥λ = mstrԭֵ�б굥λ & "0|"
                    ElseIf intCol <> .ColIndex("����") Then
                        mstrԭֵ�б굥λ = mstrԭֵ�б굥λ & .TextMatrix(intRow, intCol) & "|"
                    End If
                Next
            Next
        End With
    Else
        '��ʾ��ͣ�õ��б굥λ
        vsfͣ��(0).TextMatrix(1, 0) = "1"
        gstrSql = "Select C.ID,'['||C.����||']'||C.���� ��λ,B.����ʱ��,B.�б���� From ҩƷ��� A,ҩƷ�б굥λ B,��Ӧ�� C" & _
                " Where A.ҩƷID=B.ҩƷID And instr(C.����,1,1)=1 And B.��λID=C.ID And A.ҩƷID=[1] " & _
                " And Not (b.����ʱ�� Is Null Or b.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) " & _
                " Order by B.����ʱ��"
        Set rsTemps = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(lblMedi.Tag))
        
        With rsTemps
            vsfͣ��(0).Rows = 1
            vsfͣ��(0).Rows = IIf(.RecordCount > 0, .RecordCount + 1, 2)
            Do While Not .EOF
                vsfͣ��(0).TextMatrix(.AbsolutePosition, 0) = .AbsolutePosition
                vsfͣ��(0).TextMatrix(.AbsolutePosition, vsfͣ��(0).ColIndex("�б굥λ")) = !��λ
                vsfͣ��(0).TextMatrix(.AbsolutePosition, vsfͣ��(0).ColIndex("�б����")) = IIf(IsNull(!�б����), "", !�б����)
                vsfͣ��(0).TextMatrix(.AbsolutePosition, vsfͣ��(0).ColIndex("����ʱ��")) = Format(!����ʱ��, "YYYY-MM-DD HH:MM:SS")
                .MoveNext
            Loop
        End With
        
        '��ʾ�ѳ�ʼ�����б굥λ
        vsfUnit.TextMatrix(1, 0) = "1"
        
        gstrSql = "Select C.ID,'['||C.����||']'||C.���� ��λ,B.����ʱ��,B.�б���� From ҩƷ��� A,ҩƷ�б굥λ B,��Ӧ�� C" & _
                " Where A.ҩƷID=B.ҩƷID And instr(C.����,1,1)=1 And B.��λID=C.ID And A.ҩƷID=[1] " & _
                " And (B.����ʱ�� is null or B.����ʱ��=to_date('3000-01-01','YYYY-MM-DD')) " & _
                " Order by B.����ʱ��"
        Set rsTemps = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(lblMedi.Tag))
        
        With rsTemps
            vsfUnit.Rows = 1
            vsfUnit.Rows = IIf(.RecordCount > 0, .RecordCount + 1, 2)
            Do While Not .EOF
                vsfUnit.TextMatrix(.AbsolutePosition, 0) = .AbsolutePosition
                vsfUnit.TextMatrix(.AbsolutePosition, vsfUnit.ColIndex("��λ")) = !��λ
                vsfUnit.TextMatrix(.AbsolutePosition, vsfUnit.ColIndex("�б����")) = IIf(IsNull(!�б����), "", !�б����)
                vsfUnit.TextMatrix(.AbsolutePosition, vsfUnit.ColIndex("����ʱ��")) = Format(!����ʱ��, "YYYY-MM-DD HH:MM:SS")
                vsfUnit.TextMatrix(.AbsolutePosition, vsfUnit.ColIndex("״̬")) = 0
                vsfUnit.TextMatrix(.AbsolutePosition, vsfUnit.ColIndex("����")) = 0
                vsfUnit.TextMatrix(.AbsolutePosition, vsfUnit.ColIndex("��λID")) = !ID
                .MoveNext
            Loop
        End With
        
        mstrԭֵ = ""
        With vsfUnit
            .Cell(flexcpBackColor, 0, .ColIndex("����ʱ��"), .Rows - 1, .ColIndex("����ʱ��")) = mcstDelColor
            For intRow = 1 To .Rows - 1
                For intCol = 1 To .Cols - 1
                    If intCol = .ColIndex("ͣ��") Then
                        mstrԭֵ = mstrԭֵ & "0|"
                    ElseIf intCol <> .ColIndex("����") Then
                        mstrԭֵ = mstrԭֵ & .TextMatrix(intRow, intCol) & "|"
                    End If
                Next
            Next
        End With
    End If
        
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItems.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwItems.SortOrder = IIf(Me.lvwItems.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItems.SortKey = ColumnHeader.Index - 1
        Me.lvwItems.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwItems_DblClick()
    Dim intRow As Integer
    Dim intCol As Integer
    Dim strTemp As String
    
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    
    With vsfUnit
        For intRow = 1 To .Rows - 1
            For intCol = 1 To .Cols - 1
                If intCol = .ColIndex("ͣ��") Then
                    strTemp = strTemp & IIf(.TextMatrix(intRow, intCol) = "", 0, .TextMatrix(intRow, intCol)) & "|"
                ElseIf intCol <> .ColIndex("����") Then
                    strTemp = strTemp & .TextMatrix(intRow, intCol) & "|"
                End If
            Next
        Next
    End With
    
    If strTemp <> mstrԭֵ Then
        If mblnSave = False Then
            If MsgBox("��ǰ���ݱ��޸ĺ�δ���棬��ȷ��Ҫ������", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                lvwItems.Visible = False
                Exit Sub
            End If
        End If
    End If
    
    With Me.lvwItems
        If Me.lblMedi.Tag <> Mid(.SelectedItem.Key, 2) Then
            Me.lblMedi.Tag = Mid(.SelectedItem.Key, 2)
            Me.txtMedi.Tag = "[" & .SelectedItem.SubItems(.ColumnHeaders("����").Index - 1) & "]" & .SelectedItem.Text
            Me.txtMedi.Text = Me.txtMedi.Tag
            lblSpec.Caption = "���" & lvwItems.SelectedItem.SubItems(.ColumnHeaders("���").Index - 1) & _
                                "      ���ƣ�" & lvwItems.SelectedItem.SubItems(3) & _
                                "     ��λ��" & lvwItems.SelectedItem.SubItems(.ColumnHeaders("��λ").Index - 1)
            Call ShowData
        End If
        Me.txtMedi.SetFocus
        Call OS.PressKey(vbKeyTab)
    End With
End Sub

Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
        Call lvwItems_DblClick
    End Select
End Sub

Private Sub lvwItems_LostFocus()
    Me.lvwItems.Visible = False
End Sub

Private Sub msf�б굥λѡ��_LostFocus()
    With Msf�б굥λѡ��
        .ZOrder 1
        .Visible = False
    End With
End Sub

Private Sub vsfUnit_EnterCell()
    Dim rs�޸��б���� As ADODB.Recordset
    
    On Error GoTo ErrHand
    With vsfUnit
        If .Rows = 1 Then Exit Sub
        If .Row = 0 Then Exit Sub

        If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mcstDelColor Then
            .Editable = flexEDNone
        Else
            .Editable = flexEDKbdMouse
        End If
        If .Col = .ColIndex("��λ") Then
            .ColComboList(.ColIndex("��λ")) = ""

            gstrSql = "Select 1" & vbNewLine & _
                            "From ҩƷ�б굥λ T" & vbNewLine & _
                            "Where t.��λid =[1] And t.ҩƷid =[2]" & vbNewLine & _
                            "And Sysdate Between t.����ʱ�� And Nvl(t.����ʱ��, To_Date('3000-1-1', 'YYYY-MM-DD'))"
            Set rs�޸��б���� = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, .TextMatrix(.Row, .ColIndex("��λID")), Val(lblMedi.Tag))
            If rs�޸��б����.RecordCount > 0 Then .Editable = flexEDNone
        End If
    End With

    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfUnit_GotFocus()
    lvwItems.Visible = False
End Sub

Private Sub vsfUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsfUnit
        If KeyCode = vbKeyReturn Then
            If .Col <> .ColIndex("�б����") Then
                .Col = .Col + 1
            ElseIf .Row <> .Rows - 1 And .Col = .ColIndex("�б����") Then
                .Row = .Row + 1
                .Col = .ColIndex("��λ")
            ElseIf .Row = .Rows - 1 And .TextMatrix(.Row, 1) <> "" Then
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = .Rows - 1
                .Row = .Rows - 1
                .Col = .ColIndex("��λ")
            End If
        ElseIf KeyCode = vbKeyDelete Then
            Call Delete
        End If
        If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mcstDelColor Then
            KeyCode = 0
        End If
    End With
End Sub

Private Sub vsfUnit_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If Col = vsfUnit.ColIndex("��λ") Then
        vsfUnit.ColComboList(Col) = "|..."
    End If
End Sub

Private Sub vsfUnit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With vsfUnit
        If .Col = .ColIndex("��λ") Then
            .ColComboList(.ColIndex("��λ")) = "|..."
        Else
            .ColComboList(.ColIndex("��λ")) = ""
        End If
    End With
End Sub

Private Sub vsfUnit_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    With vsfUnit
        .EditSelStart = 0
        .EditSelLength = zlcommfun.ActualLen(.EditText)
    End With
End Sub

Private Sub vsfUnit_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfUnit
        If Col = .ColIndex("�б����") Then
            .EditMaxLength = 50
        Else
            .EditMaxLength = 50
        End If
    End With
End Sub

Private Sub vsfUnit_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsRecord As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim rs�޸��б���� As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    vRect = zlControl.GetControlRect(vsfUnit.hWnd) '��ȡλ��
    dblLeft = vRect.Left + vsfUnit.CellLeft
    dblTop = vRect.Top + vsfUnit.CellTop + vsfUnit.CellHeight + 3200
    With vsfUnit
        If Col = .ColIndex("��λ") And .EditText = "" Then Exit Sub
        If .Rows < 2 Then Exit Sub
        If Col = .ColIndex("��λ") And InStr(1, .EditText, "[") = 0 Then
            gstrSql = "Select ID,����,����,���� From ��Ӧ�� Where ĩ��=1 And (instr(����,1,1)=1 Or Nvl(ĩ��,0)=0) And (����ʱ�� is null or ����ʱ��=to_date('3000-01-01','YYYY-MM-DD')) and (���� like [1] or ���� like[1] or ���� like [1]) Order By ���� "
            
            Set rsRecord = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, "��Ӧ��", False, "", "", False, False, _
            True, dblLeft, dblTop, .Height, blnCancel, False, True, UCase(.EditText) & "%")

            If blnCancel = True Then
                Cancel = True
                Exit Sub
            End If
            
            If rsRecord Is Nothing Then
                MsgBox "�޸��б굥λ��", vbInformation, gstrSysName
                Cancel = True
                Exit Sub
            Else
                mlngId = 0
                mlngId = rsRecord!ID
                If CheckDub = False Then
                    .TextMatrix(Row, 0) = Row
                    .EditText = "[" & rsRecord!���� & "]" & rsRecord!����
                    .TextMatrix(Row, .ColIndex("��λ")) = .EditText
                    .TextMatrix(Row, .ColIndex("״̬")) = mStates.����
                    .TextMatrix(Row, .ColIndex("����")) = mStates.����
                    .TextMatrix(Row, .ColIndex("��λID")) = rsRecord!ID
                    .Cell(flexcpBackColor, Row, .ColIndex("����ʱ��"), Row, .ColIndex("����ʱ��")) = mcstDelColor
                    .Col = .ColIndex("�б����")
                Else
                    MsgBox "�Ѿ��и��б굥λ��", vbInformation, gstrSysName
                    Cancel = True
                End If
            End If
 
         ElseIf Col = .ColIndex("�б����") And .TextMatrix(.Row, .ColIndex("�б����")) <> .EditText Then
            gstrSql = "Select 1" & vbNewLine & _
                            "From ҩƷ�б굥λ T" & vbNewLine & _
                            "Where t.��λid =[1] And t.ҩƷid =[2]" & vbNewLine & _
                            "And Sysdate Between t.����ʱ�� And Nvl(t.����ʱ��, To_Date('3000-1-1', 'YYYY-MM-DD'))"
            Set rs�޸��б���� = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, .TextMatrix(Row, .ColIndex("��λID")), Val(lblMedi.Tag))
            If rs�޸��б����.RecordCount > 0 Then .TextMatrix(Row, .ColIndex("����")) = mStates.�޸�
            
        End If
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function CheckDub(Optional ByVal bln As Boolean = False) As Boolean
    '����Ƿ���ڸ��б굥λ���Ƿ����ҩƷ
    Dim i As Integer
    
    If bln Then
        With vsf�б굥λ
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("ҩƷid")) <> "" Then
                    If .TextMatrix(i, .ColIndex("ҩƷid")) = mlngId Then
                        CheckDub = True
                        Exit Function
                    End If
                Else
                    CheckDub = False
                End If
            Next
        End With
        CheckDub = False
    Else
        With vsfUnit
            For i = 1 To .Rows - 1
                If .TextMatrix(i, .ColIndex("��λid")) <> "" Then
                    If .TextMatrix(i, .ColIndex("��λid")) = mlngId Then
                        CheckDub = True
                        Exit Function
                    End If
                Else
                    CheckDub = False
                End If
            Next
        End With
        CheckDub = False
    End If
End Function

Private Sub vsf�б굥λ_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsRecord As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    
    vRect = zlControl.GetControlRect(vsf�б굥λ.hWnd) '��ȡλ��
    dblLeft = vRect.Left + vsf�б굥λ.CellLeft
    dblTop = vRect.Top + vsf�б굥λ.CellTop + vsf�б굥λ.CellHeight + 3300
    With vsf�б굥λ
        If Col = .ColIndex("ҩƷ���") Then
            gstrSql = "select I.ID,I.����,I.����,I.���,I.����" & _
                    " from �շ���ĿĿ¼ I" & _
                    " where I.��� in('5','6')" & _
                    "       and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))"
                    
            Set rsRecord = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, "ҩƷ���", False, "", "", False, False, _
            True, dblLeft, dblTop, .Height, blnCancel, False, True)

            If rsRecord Is Nothing Then
                Exit Sub
            Else
                mlngId = 0
                mlngId = rsRecord!ID
                If CheckDub(True) = False Then
                    .TextMatrix(Row, 0) = Row
                    .TextMatrix(Row, .ColIndex("ҩƷ���")) = "[" & rsRecord!���� & "]" & rsRecord!���� & "(" & rsRecord!��� & ")"
                    .TextMatrix(Row, .ColIndex("״̬")) = mStates.����
                    .TextMatrix(Row, .ColIndex("����")) = mStates.����
                    .TextMatrix(Row, .ColIndex("ҩƷID")) = rsRecord!ID
                    .Cell(flexcpBackColor, Row, .ColIndex("����ʱ��"), Row, .ColIndex("����ʱ��")) = mcstDelColor
                    .Col = .ColIndex("�б����")
                Else
                    MsgBox "�Ѿ��и�ҩƷ��", vbInformation, gstrSysName
                End If
            End If
        End If
    End With
End Sub

Private Sub vsf�б굥λ_EnterCell()
    Dim rs�޸��б���� As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    With vsf�б굥λ
        If .Rows = 1 Then Exit Sub
        If .Row = 0 Then Exit Sub

        If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mcstDelColor Then
            .Editable = flexEDNone
        Else
            .Editable = flexEDKbdMouse
        End If
        If .Col = .ColIndex("ҩƷ���") Then
            .ColComboList(.ColIndex("ҩƷ���")) = ""

            gstrSql = "Select 1" & vbNewLine & _
                            "From ҩƷ�б굥λ T" & vbNewLine & _
                            "Where t.��λid =[1] And t.ҩƷid =[2]" & vbNewLine & _
                            "And Sysdate Between t.����ʱ�� And Nvl(t.����ʱ��, To_Date('3000-1-1', 'YYYY-MM-DD'))"
            Set rs�޸��б���� = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(lbl�б굥λ.Tag), .TextMatrix(.Row, .ColIndex("ҩƷID")))
            If rs�޸��б����.RecordCount > 0 Then .Editable = flexEDNone
        End If
    End With

    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsf�б굥λ_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsf�б굥λ
        If KeyCode = vbKeyReturn Then
            If .Col <> .ColIndex("�б����") Then
                .Col = .Col + 1
            ElseIf .Row <> .Rows - 1 And .Col = .ColIndex("�б����") Then
                .Row = .Row + 1
                .Col = .ColIndex("ҩƷ���")
            ElseIf .Row = .Rows - 1 And .TextMatrix(.Row, 1) <> "" Then
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = .Rows - 1
                .Row = .Rows - 1
                .Col = .ColIndex("ҩƷ���")
            End If
        ElseIf KeyCode = vbKeyDelete Then
            Call Delete(True)
        End If
        If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mcstDelColor Then
            KeyCode = 0
        End If
    End With
End Sub

Private Sub vsf�б굥λ_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If Col = vsf�б굥λ.ColIndex("ҩƷ���") Then
        vsf�б굥λ.ColComboList(Col) = "|..."
    End If
End Sub

Private Sub vsf�б굥λ_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With vsf�б굥λ
        If .Col = .ColIndex("ҩƷ���") Then
            .ColComboList(.ColIndex("ҩƷ���")) = "|..."
        Else
            .ColComboList(.ColIndex("ҩƷ���")) = ""
        End If
    End With
End Sub

Private Sub vsf�б굥λ_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    With vsf�б굥λ
        .EditSelStart = 0
        .EditSelLength = zlcommfun.ActualLen(.EditText)
    End With
End Sub

Private Sub vsf�б굥λ_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsf�б굥λ
        If Col = .ColIndex("�б����") Then
            .EditMaxLength = 50
        Else
            .EditMaxLength = 50
        End If
    End With
End Sub

Private Sub vsf�б굥λ_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsRecord As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim rs�޸��б���� As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    vRect = zlControl.GetControlRect(vsf�б굥λ.hWnd) '��ȡλ��
    dblLeft = vRect.Left + vsf�б굥λ.CellLeft
    dblTop = vRect.Top + vsf�б굥λ.CellTop + vsf�б굥λ.CellHeight + 3300
    With vsf�б굥λ
        If Col = .ColIndex("ҩƷ���") And .EditText = "" Then Exit Sub
        If .Rows < 2 Then Exit Sub
        If Col = .ColIndex("ҩƷ���") And InStr(1, .EditText, "[") = 0 Then
        gstrSql = "select distinct I.ID,I.����,I.����,I.���,I.����" & _
                " from �շ���ĿĿ¼ I,�շ���Ŀ���� N" & _
                " where I.ID=N.�շ�ϸĿID and I.��� in('5','6') " & _
                "       and (I.����ʱ�� is null or I.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & _
                "       and (I.���� like [1] or N.���� like [2] or N.���� like [2])"
            
            Set rsRecord = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, "ҩƷ���", False, "", "", False, False, _
            True, dblLeft, dblTop, .Height, blnCancel, False, True, UCase(.EditText) & "%", gstrMatch & UCase(.EditText) & "%")
            
            If blnCancel = True Then
                Cancel = True
                Exit Sub
            End If
  
            If rsRecord Is Nothing Then
                MsgBox "û���ҵ���ҩƷ��", vbInformation, gstrSysName
                Cancel = True
                Exit Sub
            Else
                mlngId = 0
                mlngId = rsRecord!ID
                If CheckDub(True) = False Then
                    .TextMatrix(Row, 0) = Row
                    .EditText = "[" & rsRecord!���� & "]" & rsRecord!���� & "(" & rsRecord!��� & ")"
                    .TextMatrix(Row, .ColIndex("ҩƷ���")) = .EditText
                    .TextMatrix(Row, .ColIndex("״̬")) = mStates.����
                    .TextMatrix(Row, .ColIndex("����")) = mStates.����
                    .TextMatrix(Row, .ColIndex("ҩƷID")) = rsRecord!ID
                    .Cell(flexcpBackColor, Row, .ColIndex("����ʱ��"), Row, .ColIndex("����ʱ��")) = mcstDelColor
                    .Col = .ColIndex("�б����")
                Else
                    MsgBox "�Ѿ��и�ҩƷ��", vbInformation, gstrSysName
                    Cancel = True
                End If
            End If
            
         ElseIf Col = .ColIndex("�б����") And .TextMatrix(.Row, .ColIndex("�б����")) <> .EditText Then
            gstrSql = "Select 1" & vbNewLine & _
                            "From ҩƷ�б굥λ T" & vbNewLine & _
                            "Where t.��λid =[1] And t.ҩƷid =[2]" & vbNewLine & _
                            "And Sysdate Between t.����ʱ�� And Nvl(t.����ʱ��, To_Date('3000-1-1', 'YYYY-MM-DD'))"
            Set rs�޸��б���� = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(lbl�б굥λ.Tag), .TextMatrix(Row, .ColIndex("ҩƷID")))
            If rs�޸��б����.RecordCount > 0 Then .TextMatrix(Row, .ColIndex("����")) = mStates.�޸�
            
        End If
    End With
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Resize(ByVal Index As Integer)
    On Error Resume Next
 
    If chk��ʾͣ��(Index).Value = 1 Then
        vsfͣ��(Index).Visible = True
        frmMediUnit.Height = 9000
        sstCustom.Height = 8295
        cmdRestore(Index).Top = 7740
        cmdStartAll(Index).Top = 7740
        cmdStopAll(Index).Top = 7740
        cmdSave(Index).Top = 7740
        cmdClose(Index).Top = 7740
        vsfͣ��(Index).Top = 5280
        vsfͣ��(Index).Left = 120
    Else
        vsfͣ��(Index).Visible = False
        frmMediUnit.Height = 6560
        sstCustom.Height = 5895
        cmdRestore(Index).Top = 5340
        cmdStartAll(Index).Top = 5340
        cmdStopAll(Index).Top = 5340
        cmdSave(Index).Top = 5340
        cmdClose(Index).Top = 5340
    End If
End Sub


