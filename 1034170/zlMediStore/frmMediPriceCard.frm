VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmMediPriceCard 
   Caption         =   "ҩƷ���۵�"
   ClientHeight    =   10380
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14655
   Icon            =   "frmMediPriceCard.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10380
   ScaleWidth      =   14655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picSplit 
      BorderStyle     =   0  'None
      Height          =   100
      Left            =   240
      MousePointer    =   7  'Size N S
      ScaleHeight     =   105
      ScaleWidth      =   2775
      TabIndex        =   48
      Top             =   4200
      Width           =   2775
   End
   Begin VB.TextBox txtFind 
      Height          =   300
      Left            =   1080
      TabIndex        =   44
      Top             =   9505
      Width           =   1965
   End
   Begin VB.PictureBox picOtherSelect 
      Height          =   3135
      Left            =   3600
      ScaleHeight     =   3075
      ScaleWidth      =   4755
      TabIndex        =   28
      Top             =   1320
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton cmdFilterOk 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   2400
         Picture         =   "frmMediPriceCard.frx":6852
         TabIndex        =   41
         Top             =   2640
         Width           =   1100
      End
      Begin VB.CommandButton cmdFilterCan 
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   3480
         Picture         =   "frmMediPriceCard.frx":699C
         TabIndex        =   40
         Top             =   2640
         Width           =   1100
      End
      Begin VB.Frame fra����ѡ�� 
         Caption         =   "����ѡ��ɱ��۵�����أ�"
         Height          =   2535
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   4695
         Begin VB.CheckBox chk�ӳ��� 
            Caption         =   "ָ���ӳ���"
            Height          =   180
            Left            =   120
            TabIndex        =   35
            Top             =   1125
            Width           =   1215
         End
         Begin VB.CheckBox chk��Ӧ�� 
            Caption         =   "ָ����Ӧ��"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox chkӦ����¼ 
            Caption         =   "�����ɱ��۵��۴�����Ӧ����������¼"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   1920
            Width           =   3495
         End
         Begin VB.TextBox txt�ӳ��� 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   270
            Left            =   1440
            TabIndex        =   32
            Text            =   "15.0000"
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox txt��Ӧ�� 
            Enabled         =   0   'False
            Height          =   270
            Left            =   1440
            TabIndex        =   31
            Top             =   360
            Width           =   2655
         End
         Begin VB.CommandButton cmd��Ӧ�� 
            Caption         =   "��"
            Enabled         =   0   'False
            Height          =   270
            Left            =   4080
            TabIndex        =   30
            Top             =   350
            Width           =   375
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshProvider 
            Height          =   1695
            Left            =   120
            TabIndex        =   36
            Top             =   2280
            Visible         =   0   'False
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   2990
            _Version        =   393216
            FixedCols       =   0
            GridColor       =   32768
            FocusRect       =   0
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label lblComment�ӳ��� 
            Caption         =   "��ָ���ӳ��ʣ���ͳһĬ�ϰ��üӳ��ʼ���ɱ��ۣ���ָ������Ĭ����ʾʵ�ʼӳ��ʣ�"
            ForeColor       =   &H00FF0000&
            Height          =   540
            Left            =   240
            TabIndex        =   39
            Top             =   1440
            Width           =   4260
         End
         Begin VB.Label lblComment��Ӧ�� 
            AutoSize        =   -1  'True
            Caption         =   "��ָ����Ӧ�̣���ֻ�����ù�Ӧ�̵Ŀ��ҩƷ�ɱ��ۣ�"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   240
            TabIndex        =   38
            Top             =   720
            Width           =   4320
         End
         Begin VB.Label lblPercent 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   180
            Left            =   2415
            TabIndex        =   37
            Top             =   1125
            Width           =   90
         End
      End
   End
   Begin VB.PictureBox picInfo 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   10575
      TabIndex        =   23
      Top             =   8640
      Width           =   10575
      Begin VB.TextBox txtSummary 
         Height          =   300
         Left            =   4320
         MaxLength       =   100
         TabIndex        =   26
         Top             =   120
         Width           =   5565
      End
      Begin VB.TextBox txtValuer 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         Left            =   990
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   120
         Width           =   1965
      End
      Begin VB.Label lblSummary 
         AutoSize        =   -1  'True
         Caption         =   "����˵��"
         Height          =   180
         Left            =   3360
         TabIndex        =   27
         Top             =   180
         Width           =   720
      End
      Begin VB.Label lblValuer 
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   180
         Left            =   360
         TabIndex        =   25
         Top             =   180
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "���(&D)"
      Height          =   350
      Left            =   6960
      Picture         =   "frmMediPriceCard.frx":6AE6
      TabIndex        =   15
      Top             =   9480
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   12600
      Picture         =   "frmMediPriceCard.frx":6C30
      TabIndex        =   14
      Top             =   9480
      Width           =   1215
   End
   Begin VB.CommandButton cmdCanc 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   13920
      Picture         =   "frmMediPriceCard.frx":6D7A
      TabIndex        =   13
      Top             =   9480
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "��ӡ���䶯��(&P)��"
      Height          =   350
      Left            =   10200
      Picture         =   "frmMediPriceCard.frx":6EC4
      TabIndex        =   12
      Top             =   9480
      Width           =   1935
   End
   Begin VB.CommandButton cmdItem 
      Caption         =   "����ѡ����Ŀ(&I)"
      Height          =   350
      Left            =   8400
      Picture         =   "frmMediPriceCard.frx":700E
      TabIndex        =   11
      Top             =   9480
      Width           =   1695
   End
   Begin VB.Frame fraCondition 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   16335
      Begin VB.CheckBox chkAutoPay 
         Caption         =   "�Զ�����Ӧ����䶯��¼"
         Height          =   210
         Left            =   8160
         TabIndex        =   43
         Top             =   480
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CheckBox chkCostBatch 
         Caption         =   "�ɱ��۰��ⷿ���ε���"
         Height          =   210
         Left            =   2160
         TabIndex        =   42
         Top             =   480
         Width           =   2370
      End
      Begin VB.CheckBox chkAotuCost 
         Caption         =   "���ۼ�ʱ�Զ����ӳ��ʵ����ɱ���"
         Height          =   210
         Left            =   4680
         TabIndex        =   20
         Top             =   480
         Width           =   3015
      End
      Begin VB.CheckBox Chk���� 
         Caption         =   "ʱ��ҩƷ��Ϊ����"
         Height          =   210
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   1770
      End
      Begin VB.CommandButton cmdPriceMethod 
         Caption         =   "��"
         Height          =   300
         Left            =   3360
         TabIndex        =   17
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox cboPriceMethod 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   0
         Width           =   2415
      End
      Begin VB.CheckBox chk������ 
         Caption         =   "�ɱ��۰��ⷿ���ε���"
         Height          =   210
         Left            =   10560
         TabIndex        =   8
         Top             =   -225
         Width           =   2175
      End
      Begin VB.CheckBox chk�Զ�����Ӧ����䶯 
         Caption         =   "�Զ�����Ӧ����䶯"
         Height          =   210
         Left            =   12840
         TabIndex        =   7
         Top             =   -225
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.OptionButton optʱ�� 
         Caption         =   "����ִ��"
         Height          =   255
         Index           =   0
         Left            =   5040
         TabIndex        =   6
         Top             =   8
         Width           =   1095
      End
      Begin VB.OptionButton optʱ�� 
         Caption         =   "ָ������ִ��"
         Height          =   255
         Index           =   1
         Left            =   6240
         TabIndex        =   5
         Top             =   8
         Width           =   1455
      End
      Begin VB.ComboBox cbo�ۼۼ��㷽ʽ 
         Height          =   300
         Left            =   13080
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   0
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker dtpRunDate 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "yyyy-MM-dd"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
         Height          =   300
         Left            =   8040
         TabIndex        =   9
         Top             =   0
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
         Format          =   125698051
         CurrentDate     =   36846.5833333333
      End
      Begin VB.Label lbl���۷�ʽ 
         AutoSize        =   -1  'True
         Caption         =   "�ۼۼ��㷽ʽ"
         Height          =   180
         Left            =   11520
         TabIndex        =   22
         Top             =   60
         Width           =   1080
      End
      Begin VB.Label lblMethod 
         AutoSize        =   -1  'True
         Caption         =   "���۷�ʽ"
         Height          =   180
         Left            =   120
         TabIndex        =   21
         Top             =   60
         Width           =   720
      End
      Begin VB.Label lblִ��ʱ�� 
         Caption         =   "ִ��ʱ��"
         Height          =   180
         Left            =   4200
         TabIndex        =   10
         Top             =   45
         Width           =   855
      End
   End
   Begin VB.TextBox txtNO 
      Enabled         =   0   'False
      Height          =   300
      Left            =   13200
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin XtremeSuiteControls.TabControl TabCtlDetails 
      Height          =   975
      Left            =   240
      TabIndex        =   18
      Top             =   5040
      Width           =   1815
      _Version        =   589884
      _ExtentX        =   3201
      _ExtentY        =   1720
      _StockProps     =   64
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfStore 
      Height          =   975
      Left            =   2880
      TabIndex        =   46
      Top             =   4680
      Width           =   3495
      _cx             =   6165
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   10526880
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   Begin VSFlex8Ctl.VSFlexGrid vsfPay 
      Height          =   975
      Left            =   8040
      TabIndex        =   47
      Top             =   4680
      Width           =   3495
      _cx             =   6165
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   10526880
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   Begin VSFlex8Ctl.VSFlexGrid vsfPrice 
      Height          =   2295
      Left            =   480
      TabIndex        =   49
      Top             =   2040
      Width           =   11055
      _cx             =   19500
      _cy             =   4048
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   10526880
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   50
      Top             =   10020
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMediPriceCard.frx":7158
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20082
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1402
            MinWidth        =   1411
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
            Object.ToolTipText     =   "��ǰ���ּ�״̬"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
            Object.ToolTipText     =   "��ǰ��д��״̬"
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
   Begin VB.Label lblFind 
      Caption         =   "����"
      Height          =   255
      Left            =   480
      TabIndex        =   45
      Top             =   9528
      Width           =   495
   End
   Begin VB.Label lblNO 
      AutoSize        =   -1  'True
      Caption         =   "������ˮ��"
      Height          =   180
      Left            =   12120
      TabIndex        =   1
      Top             =   180
      Width           =   900
   End
   Begin VB.Label lblDrugName 
      AutoSize        =   -1  'True
      Caption         =   "ҩƷ���۵�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6120
      TabIndex        =   0
      Top             =   120
      Width           =   1875
   End
End
Attribute VB_Name = "frmMediPriceCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'����ȫ�ֱ���
Private Const mlngRowHeight As Long = 300 '����и����и�
Private mintUnit As Integer     '������¼���õ���ʲô��λ
Private mint���� As Integer     '0-���ۼ�;1-���ɱ���;2-���ۼۼ��ɱ���
Private mlng��Ӧ��ID As Long  '������¼��Ӧ��id
Private mdbl�ӳ��� As Double
Private mblnӦ����¼ As Boolean '��¼�Ƿ����Ӧ����¼
Private marrSql() As Variant     '��¼��delete��ɾ��ҩƷ�Ĵ洢���̵�����

Private mintCostDigit As Integer        '�ɱ���С��λ��
Private mintPriceDigit As Integer       '�ۼ�С��λ��
Private mintNumberDigit As Integer      '����С��λ��
Private mintMoneyDigit As Integer       '���С��λ��
Private mstrMoneyFormat As String
Private mintSalePriceDigit As Integer
'��ɫ����
Private Const mconlngColor As Long = &HFFFFFF        '�����޸�����ɫΪ��ɫ
Private Const mconlngCanColColor As Long = &HE7CFBA    '���޸�����ɫΪ����ɫ

Private mblnʱ��ҩƷ�����ε��� As Boolean 'ʱ��ҩƷ�������ε���
Private mbln�ּ���ʾ As Boolean         '�޼�ҩƷ��ʾ true-��ʾ false-����ʾ
Private mdbl�ֶμӳ��� As Double    '������¼�ֶμӳ���
Private mdbl�ɱ��� As Double            '��¼�޸�֮ǰ�ĳɱ���
Private mrs�ֶμӳ� As ADODB.Recordset  '��¼�ֶμӳ��ʼ���
Private mstrNo As String            '���۵�No
Private mintModal As Integer        '������ʲô״̬ 0-���� 1-�޸� 2-����
Private mintMethod As Integer   '���۷�ʽ 0-���ۼ�;1-���ɱ���;2-���ۼۼ��ɱ���
Private mstr���ۻ��ܺ� As String
Private mblnLoad As Boolean     '�Ƿ�������
Private mrsReturn As ADODB.Recordset '����ѡ�񷵻ص����ݼ�
Private mblnOK As Boolean
Private mrsFindName As ADODB.Recordset '��ѯ�����ݼ�
Private mBlnClick As Boolean
Private mblnUpdateAdd As Boolean    '�޸�����µ���������
Private mlngOldDrugID As Long '���ԭʼ���Ƿ���ҩƷ
Private mdblOldPrice As Double   'ԭ�ۼ�
Private mblnBatchItem As Boolean   '��¼�Ƿ���������ѡ��ť
Private mstrPrivs As String     '����ԱȨ��
Private Const MStrCaption As String = "ҩƷ���۵�"

Private Enum menuPriceCol
    ҩƷid = 0
    ԭ��id = 1
    Ʒ�� = 2
    ��� = 3
    �Ƿ���
    ����
    ��λ
    ��װϵ��
    �ӳ���
    ���������
    �Ƿ��п��
    ������ĿID
    ԭ�ɱ���
    �ֳɱ���
    ԭ���ۼ�
    �����ۼ�
    ԭ�ɹ��޼�
    �ֲɹ��޼�
    ԭָ���ۼ�
    ��ָ���ۼ�
    ������
End Enum
Private Enum menuStoreCol
    ҩƷid = 0
    �ⷿ = 1
    �ⷿid = 2
    ��Ӧ��
    ��Ӧ��id
    ҩƷ
    ���
    ����
    Ч��
    ����
    ����
    ���
    ����
    ��λ
    ��װϵ��
    ԭ���ۼ�
    �����ۼ�
    �������
    �ӳ���
    ԭ�ɹ���
    �ֲɹ���
    ��۲�
    ������
End Enum

Private Enum menuPayCol
    ҩƷid = 0
    Ʒ�� = 1
    ��Ʊ�� = 2
    ��Ʊ����
    ��Ʊ���
    ������
End Enum

Public Sub ShowME(ByVal frmParent As Form, ByVal intModal As Integer, ByVal str���ۻ��ܺ� As String, ByVal intMethod As Integer)
    mintModal = intModal
    mstr���ۻ��ܺ� = str���ۻ��ܺ�
    mintMethod = intMethod

    Me.Show vbModal, frmParent
End Sub

Private Sub cboPriceMethod_Click()
    Dim intCol As Integer
    Dim intTemp As Integer

    With cboPriceMethod
        If .Text = "�����ۼ�" Then
            intTemp = 0
            lbl���۷�ʽ.Visible = False
            cbo�ۼۼ��㷽ʽ.Visible = False
        ElseIf .Text = "�����ɱ���" Then
            intTemp = 1
            lbl���۷�ʽ.Visible = False
            cbo�ۼۼ��㷽ʽ.Visible = False
        Else
            intTemp = 2
            lbl���۷�ʽ.Visible = True
            cbo�ۼۼ��㷽ʽ.Visible = True
        End If
    End With


    If mblnLoad = True And intTemp <> Val(lblMethod.Tag) Then
        If vsfPrice.TextMatrix(1, menuPriceCol.ҩƷid) <> "" Then
            If MsgBox("���۷�ʽ�ı佫����б������ݣ��Ƿ������", vbYesNo, gstrSysName) = vbNo Then
                cboPriceMethod.ListIndex = mint����
                Exit Sub
            Else
                vsfPrice.rows = 2
                For intCol = 0 To vsfPrice.Cols - 1
                    vsfPrice.TextMatrix(1, intCol) = ""
                Next
                vsfStore.rows = 1
                vsfPay.rows = 1
            End If
        End If
    End If
    With cboPriceMethod
        If .Text = "�����ۼ�" Then
            mint���� = 0
            lblMethod.Tag = 0
            optʱ��(0).Value = False
            optʱ��(1).Value = True
            optʱ��(0).Enabled = True
            optʱ��(1).Enabled = True
            dtpRunDate.Enabled = True
            chkCostBatch.Visible = False
            chkCostBatch.Value = False
            chkAutoPay.Visible = False
            chkAutoPay.Value = 0
            chkAotuCost.Visible = False
            chkAotuCost.Value = False
        ElseIf .Text = "�����ɱ���" Then
            mint���� = 1
            lblMethod.Tag = 1
            optʱ��(0).Value = True
            optʱ��(0).Enabled = False
            optʱ��(1).Enabled = False
            dtpRunDate.Enabled = False
            chkCostBatch.Visible = True
            If mblnӦ����¼ = True Then
                chkAutoPay.Visible = True
                chkAutoPay.Value = 1
            End If
            chkAotuCost.Visible = False
            chkAotuCost.Value = False
        ElseIf .Text = "�ۼ۳ɱ���һ�����" Then
            mint���� = 2
            lblMethod.Tag = 2
            optʱ��(0).Value = False
            optʱ��(1).Value = True
            optʱ��(0).Enabled = True
            optʱ��(1).Enabled = True
            dtpRunDate.Enabled = True
            chkCostBatch.Visible = True
            If mblnӦ����¼ = True Then
                chkAutoPay.Visible = True
                chkAutoPay.Value = 1
            Else
                chkAutoPay.Visible = False
                chkAutoPay.Value = 0
            End If
            chkAotuCost.Visible = True
        End If
        If .Text = "�����ۼ�" Then
            cmdPriceMethod.Visible = False
            picOtherSelect.Visible = cmdPriceMethod.Visible
        Else
            cmdPriceMethod.Visible = True
        End If
    End With
    vsfStore.Cols = menuStoreCol.������
    vsfPay.Cols = menuPayCol.������
    vsfPrice.Cols = menuPriceCol.������
    Call setColEdit
    Call setColHiddenVsf
End Sub

Private Sub cboPriceMethod_DropDown()
    With cboPriceMethod
        If .Text = "�����ۼ�" Then
            mint���� = 0
        ElseIf .Text = "�����ɱ���" Then
            mint���� = 1
        ElseIf .Text = "�ۼ۳ɱ���һ�����" Then
            mint���� = 2
        End If
    End With
End Sub

Private Sub cbo�ۼۼ��㷽ʽ_Click()
    On Error GoTo errHandle
    Set mrs�ֶμӳ� = Nothing
    If cbo�ۼۼ��㷽ʽ.Text = "�ۼ۰��ֶμӳɼ���" Then
        gstrSQL = "select ���, ��ͼ�, ��߼�, �ӳ���, ��۶�, ˵��, ���� from ҩƷ�ӳɷ��� order by ���"
        Set mrs�ֶμӳ� = zlDataBase.OpenSQLRecord(gstrSQL, "ҩƷ�ӳɷ���")
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub chkAotuCost_Click()
    If chkAotuCost.Value = 1 Then
        cbo�ۼۼ��㷽ʽ.Visible = False
        cbo�ۼۼ��㷽ʽ.ListIndex = 0
        lbl���۷�ʽ.Visible = False
    Else
        cbo�ۼۼ��㷽ʽ.Visible = True
        lbl���۷�ʽ.Visible = True
    End If
End Sub


Private Sub Chk��Ӧ��_Click()
    If chk��Ӧ��.Value = 1 Then
        cmd��Ӧ��.Enabled = True
        txt��Ӧ��.Enabled = True
        chkӦ����¼.Enabled = True
    Else
        cmd��Ӧ��.Enabled = False
        txt��Ӧ��.Enabled = False
        chkӦ����¼.Enabled = False
        chkӦ����¼.Value = 0
    End If
End Sub

Private Sub chk�ӳ���_Click()
    If chk�ӳ���.Value = 1 Then
        txt�ӳ���.Enabled = True
    Else
        txt�ӳ���.Enabled = False
    End If
End Sub

Private Sub cmdCanc_Click()
    Call ReleaseSelectorRS 'ж�����ݼ�
    Unload Me
End Sub

Private Sub cmdClear_Click()
    Dim intCol As Integer

    If MsgBox("��ȷ��Ҫ����������ݣ�", vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        vsfPrice.rows = 2
        For intCol = 0 To vsfPrice.Cols - 1
            vsfPrice.TextMatrix(1, intCol) = ""
        Next
        vsfStore.rows = 1
        vsfPay.rows = 1
    End If
End Sub

Private Sub cmdFilterCan_Click()
    picOtherSelect.Visible = False
End Sub

Private Sub cmdFilterOk_Click()
    Dim i As Integer

    If chk��Ӧ��.Value = 1 Then
        If Val(Split(txt��Ӧ��.Tag, "|")(0)) = 0 Then
            MsgBox "��ѡ��Ӧ�̡�", vbInformation, gstrSysName
            txt��Ӧ��.SetFocus
            Exit Sub
        End If
    End If
    With vsfPrice
        If Val(.TextMatrix(1, menuPriceCol.ҩƷid)) <> 0 Then
            If MsgBox("����ձ���е����ݣ��Ƿ������", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            Else
                vsfPrice.rows = 2
                For i = 0 To vsfPrice.Cols - 1
                    .TextMatrix(1, i) = ""
                Next
                vsfStore.rows = 1
                vsfPay.rows = 1
            End If
        End If
    End With

    mlng��Ӧ��ID = IIf(chk��Ӧ��.Value = 1, Val(Split(txt��Ӧ��.Tag, "|")(0)), 0)
    mdbl�ӳ��� = IIf(chk�ӳ���.Value = 1, Val(Trim(txt�ӳ���.Text)), 0)
    mblnӦ����¼ = (chkӦ����¼.Enabled And chkӦ����¼.Value = 1)
    picOtherSelect.Visible = False
    If mblnӦ����¼ = True Then
        TabCtlDetails.Item(1).Visible = True
    Else
        TabCtlDetails.Item(1).Visible = False
    End If

    With cboPriceMethod
        If .Text = "�����ۼ�" Then
            mint���� = 0
            lblMethod.Tag = 0
            optʱ��(0).Value = False
            optʱ��(1).Value = True
            optʱ��(0).Enabled = True
            optʱ��(1).Enabled = True
            dtpRunDate.Enabled = True
            chkCostBatch.Visible = False
            chkCostBatch.Value = False
            chkAutoPay.Visible = False
            chkAutoPay.Value = 0
            chkAotuCost.Visible = False
            chkAotuCost.Value = False
        ElseIf .Text = "�����ɱ���" Then
            mint���� = 1
            lblMethod.Tag = 1
            optʱ��(0).Value = True
            optʱ��(0).Enabled = False
            optʱ��(1).Enabled = False
            dtpRunDate.Enabled = False
            chkCostBatch.Visible = True
            If mblnӦ����¼ = True Then
                chkAutoPay.Visible = True
                chkAutoPay.Value = 1
            Else
                chkAutoPay.Visible = False
                chkAutoPay.Value = 0
            End If
            chkAotuCost.Visible = False
            chkAotuCost.Value = False
        ElseIf .Text = "�ۼ۳ɱ���һ�����" Then
            mint���� = 2
            lblMethod.Tag = 2
            optʱ��(0).Value = False
            optʱ��(1).Value = True
            optʱ��(0).Enabled = True
            optʱ��(1).Enabled = True
            dtpRunDate.Enabled = True
            chkCostBatch.Visible = True
            If mblnӦ����¼ = True Then
                chkAutoPay.Visible = True
                chkAutoPay.Value = 1
            Else
                chkAutoPay.Visible = False
                chkAutoPay.Value = 0
            End If
            chkAotuCost.Visible = True
        End If
    End With

End Sub

Private Sub CmdHelp_Click()

End Sub

Private Sub cmdItem_Click()
    Dim intRow As Integer

    frmBatchSelect.ShowME Me, mrsReturn, mblnOK

    On Error GoTo errHandle
    If mblnOK = False Then Exit Sub
    If mrsReturn.RecordCount = 0 Then Exit Sub

    With vsfPrice
        If .TextMatrix(.rows - 1, menuPriceCol.ҩƷid) = "" Then
            intRow = .rows - 1
        Else
            .rows = .rows + 1
            intRow = .rows - 1
        End If
    End With
    mblnBatchItem = True

    Call GetDrugPirce(mrsReturn, intRow)
    mblnBatchItem = False
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub deleteNotExecutePirce()
    '���δִ�м۸�
    Dim intRow As Integer

    On Error GoTo errHandle
    With vsfPrice
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, menuPriceCol.ҩƷid) <> "" Then
                gstrSQL = "Zl_ɾ��δִ�м۸�_Delete(" & Val(.TextMatrix(intRow, menuPriceCol.ҩƷid)) & "," & 0 & ")"
                Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
            End If
        Next
    End With
    
    'ɾ��deleteɾ��������
    For intRow = 0 To UBound(marrSql)
        Call zlDataBase.ExecuteProcedure(CStr(marrSql(intRow)), Me.Caption)
    Next
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdOK_Click()
    Dim intRow As Integer
    Dim intCol As Integer
    Dim dtToday As Date
    Dim lngAdjId As Long
    Dim LngCurID As Long
    Dim strID As String
    Dim intCount As Integer
    Dim dbl��װ As Double
    Dim strTmp As String
    Dim lngCurrBatch As Long
    Dim str���μ۸� As String
    Dim blnPrint As Boolean '�Ƿ��ӡ����֪ͨ��
    Dim blnOne As Boolean   '����Ƿ��ǵ�һ��
    Dim n As Integer
    Dim intProc As Integer
    Dim blnIgnore As Boolean
    Dim blnPrice As Boolean '��¼�Ƿ��ۼ۵�����
    Dim blnCost As Boolean  '��¼�Ƿ�ɱ��۵�����
    Dim intUpdateModel As Integer '����ģʽ 0-�ۼ۵��� 1-�ɱ��۵��� 2-�ɱ����ۼ�һ�����
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim ArrayID
    Dim Array���μ۸�
    Dim strUpdate As String

    Dim lng�ⷿID As Long
    Dim lng��Ӧ��ID As Long
    Dim lngҩƷID As Long
    Dim lng����  As Long
    Dim str���� As String
    Dim strЧ�� As String
    Dim str���� As String
    Dim dblOldCost As Double
    Dim dblNewCost As Double
    Dim Str��Ʊ�� As String
    Dim str��Ʊ���� As String
    Dim dbl��Ʊ��� As Double
    Dim strInfo As String
    Dim strMsg As String '��¼��ʾ��Ϣ
    Dim intCount2 As Integer '��������
    Dim lngDouID As Long

    If vsfPrice.rows > 1 Then   'ֻ�������ݵ�����²��ܱ���
        If Val(vsfPrice.TextMatrix(1, menuPriceCol.ҩƷid)) = 0 Then Exit Sub
    End If
    If CheckPrice = False Then Exit Sub

    On Error GoTo ErrHand
    dtToday = zlDataBase.Currentdate()

    gstrSQL = "select �շѼ�Ŀ_ID.nextval from dual"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "ȡ�շѼ�Ŀ���")
    lngAdjId = rsTemp.Fields(0).Value

    gcnOracle.BeginTrans
    If mintModal = 1 Then '�޸� ���޸�ģʽ����ɾ��ԭ���ĵ�����Ϣ��Ȼ������µĵ�����Ϣ
        Call deleteNotExecutePirce
    End If

    '����Ƿ����δִ�еļ۸�
    If checkNotExecutePrice(, strInfo) = True Then
        MsgBox strInfo, vbInformation, gstrSysName
        Exit Sub
    End If
    '��ȡ����NO
    mstrNo = zlDataBase.GetNextNo(9)
    '��ȡ���ۻ���NO
    gstrSQL = "select nextno(135) as ��ˮ�� from dual"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "������ˮ��")
    If rsTemp.RecordCount = 0 Then
        MsgBox "������ˮ��δ�ܳ�ʼ���ɹ����������Ա��ϵ��", vbInformation, gstrSysName
        Exit Sub
    End If
    txtNO.Text = rsTemp!��ˮ��

    With Me.vsfPrice
        '�ۼ۵���
        strID = ""
        For intCount = 1 To IIf(Trim(.TextMatrix(.rows - 1, 0)) = "", .rows - 2, .rows - 1)
            If mint���� <> 1 Then

                LngCurID = zlDataBase.GetNextId("�շѼ�Ŀ")
                strID = strID & IIf(strID = "", "", ",") & LngCurID

                dbl��װ = Val(.TextMatrix(intCount, menuPriceCol.��װϵ��))

                If .TextMatrix(intCount, menuPriceCol.�Ƿ���) = "1" And mblnʱ��ҩƷ�����ε��� And mint���� <> 1 Then
                    strTmp = ""
                    lngCurrBatch = -1
                    For n = 1 To vsfStore.rows - 1
                        If Val(.TextMatrix(intCount, menuPriceCol.ҩƷid)) = Val(vsfStore.TextMatrix(n, menuStoreCol.ҩƷid)) Then
                            If InStr(1, "|" & strTmp, "|" & vsfStore.TextMatrix(n, menuStoreCol.����) & ",") = 0 Then
                                lngCurrBatch = vsfStore.TextMatrix(n, menuStoreCol.����)
                                strTmp = strTmp & IIf(strTmp = "", "", "|") & vsfStore.TextMatrix(n, menuStoreCol.����) & "," & vsfStore.TextMatrix(n, menuStoreCol.�����ۼ�) / dbl��װ
                            End If
                        End If
                    Next
                    str���μ۸� = str���μ۸� & strTmp
                End If
                str���μ۸� = str���μ۸� & ";"

                If CLng(.TextMatrix(intCount, menuPriceCol.ԭ��id)) <> 0 Then
                    '������һ�εļ۸��¼��ִֹ��
                    gstrSQL = "zl_�շѼ�Ŀ_stop(" & .TextMatrix(intCount, menuPriceCol.ҩƷid) & ","
                    If optʱ��(0).Value = True Then
                        gstrSQL = gstrSQL & "to_date('" & Format(DateAdd("s", -1, dtToday), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        gstrSQL = gstrSQL & "to_date('" & Format(DateAdd("s", -1, Me.dtpRunDate.Value), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    End If
                    gstrSQL = gstrSQL & ")"
                    Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)

                    '�����۸��¼
                    gstrSQL = "zl_�շѼ�Ŀ_Insert(" & LngCurID & "," & IIf(.TextMatrix(intCount, menuPriceCol.ԭ��id) = "", "NUll", Val(.TextMatrix(intCount, menuPriceCol.ԭ��id))) & _
                              "," & .TextMatrix(intCount, menuPriceCol.ҩƷid) & "," & Val(.TextMatrix(intCount, menuPriceCol.������ĿID)) & "," & _
                              Round(Val(.TextMatrix(intCount, menuPriceCol.ԭ���ۼ�)) / dbl��װ, gtype_UserDrugDigits.Digit_���ۼ�) & "," & _
                              Round(IIf(Val(.TextMatrix(intCount, menuPriceCol.�����ۼ�)) = Val(.TextMatrix(intCount, menuPriceCol.ԭ���ۼ�)), Val(.TextMatrix(intCount, menuPriceCol.�����ۼ�)) + 1, Val(.TextMatrix(intCount, menuPriceCol.�����ۼ�))) / dbl��װ, gtype_UserDrugDigits.Digit_���ۼ�) & _
                              ",NULL,NULL,'" & Me.txtSummary.Text & "'," & lngAdjId & ",'" & Trim(Me.txtValuer.Text) & "',"
                    If optʱ��(0).Value = True Then
                        gstrSQL = gstrSQL & "to_date('" & Format(dtToday, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    Else
                        gstrSQL = gstrSQL & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                    End If
                    gstrSQL = gstrSQL & ",0,'" & mstrNo & "'," & intCount & ",Null," & txtNO & ")"
                    Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
                    blnPrice = True
                    blnPrint = True
                End If
            End If
        Next
    End With

    '�ɱ��۵��۴���
    If mint���� = 1 Or mint���� = 2 Then
        If vsfStore.rows > 1 Then
            If vsfStore.TextMatrix(1, menuStoreCol.ҩƷid) <> "" Then
'                lngDouID = 0
'                For n = 1 To vsfStore.rows - 1
'                    If vsfStore.TextMatrix(n, menuStoreCol.ҩƷid) = "" Then Exit For
'
'                    '���δ��˵���
'                    If CheckUnVerify(Val(vsfStore.TextMatrix(n, menuStoreCol.ҩƷid))) = True And Val(vsfStore.TextMatrix(n, menuStoreCol.ҩƷid)) <> lngDouID Then
'                        lngDouID = Val(vsfStore.TextMatrix(n, menuStoreCol.ҩƷid))
'                        strMsg = vsfStore.TextMatrix(n, menuStoreCol.ҩƷ) & ","
'                        intCount2 = intCount2 + 1
'                        If intCount2 > 3 Then Exit For 'ֻ�ж�3��
'                    End If
'                Next
'
'                If strMsg <> "" Then
'                    If MsgBox(strMsg & "����δ��˵��ݣ������ɱ��ۿ��ܻ���ɲ����" & _
'                        vbCrLf & Space(4) & "�����ȴ���δ��˵��ݡ��Ƿ񻹼������ۣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
'                        gcnOracle.RollbackTrans
'                        Exit Sub
'                    End If
'                End If

                For n = 1 To vsfStore.rows - 1
                    For i = 1 To vsfPay.rows - 1
                        If vsfPay.TextMatrix(i, 0) = "" Then Exit For
                        If Val(vsfStore.TextMatrix(n, menuStoreCol.ҩƷid)) = Val(vsfPay.TextMatrix(i, menuPayCol.ҩƷid)) Then
                            lng�ⷿID = Val(vsfStore.TextMatrix(n, menuStoreCol.�ⷿid))
                            lng��Ӧ��ID = Val(vsfStore.TextMatrix(n, menuStoreCol.��Ӧ��id))
                            lngҩƷID = Val(vsfStore.TextMatrix(n, menuStoreCol.ҩƷid))
                            lng���� = Val(vsfStore.TextMatrix(n, menuStoreCol.����))
                            str���� = vsfStore.TextMatrix(n, menuStoreCol.����)
                            strЧ�� = IIf(Trim(vsfStore.TextMatrix(n, menuStoreCol.Ч��)) = "", "", vsfStore.TextMatrix(n, menuStoreCol.Ч��))
                            str���� = vsfStore.TextMatrix(n, menuStoreCol.����)
                            dblOldCost = GetFormat(Val(vsfStore.TextMatrix(n, menuStoreCol.ԭ�ɹ���)) / Val(vsfStore.TextMatrix(n, menuStoreCol.��װϵ��)), gtype_UserDrugDigits.Digit_�ɱ���)
                            dblNewCost = GetFormat(Val(vsfStore.TextMatrix(n, menuStoreCol.�ֲɹ���)) / Val(vsfStore.TextMatrix(n, menuStoreCol.��װϵ��)), gtype_UserDrugDigits.Digit_�ɱ���)
                            Str��Ʊ�� = vsfPay.TextMatrix(i, menuPayCol.��Ʊ��)
                            str��Ʊ���� = Format(vsfPay.TextMatrix(i, menuPayCol.��Ʊ����), "yyyy-mm-dd")
                            dbl��Ʊ��� = Val(vsfPay.TextMatrix(i, menuPayCol.��Ʊ���))

                            gstrSQL = "Zl_�ɱ��۵�����Ϣ_Insert(" & IIf(lng��Ӧ��ID = 0, "Null", lng��Ӧ��ID) & "," & lng�ⷿID & "," & lngҩƷID & "," & lng���� & ",'" & str���� & "'" & _
                                    "," & IIf(strЧ�� = "", "Null", "to_date('" & Format(strЧ��, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ",'" & str���� & "',Null," & dblOldCost & ", " & dblNewCost & "," & _
                                    IIf(Str��Ʊ�� <> "", "'" & Str��Ʊ�� & "'", "NULL") & "," & IIf(str��Ʊ���� = "", "Null", "to_date('" & Format(str��Ʊ����, "yyyy-mm-dd") & "','yyyy-mm-dd')") & ", " & dbl��Ʊ��� & "," & IIf(mblnӦ����¼ = True, 1, 0) & "," & txtNO.Text & ")"
                            Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
                            blnCost = True
                        End If
                    Next
                Next
            End If
        End If
    End If

    '�޿��ʱ�����ɱ���
    If mint���� = 1 Or mint���� = 2 Then
        With Me.vsfPrice
            For intCount = 1 To IIf(Trim(.TextMatrix(.rows - 1, 0)) = "", .rows - 2, .rows - 1)
                If .TextMatrix(intCount, menuPriceCol.�Ƿ��п��) = "0" And Val(.TextMatrix(intCount, menuPriceCol.ԭ�ɱ���)) <> Val(.TextMatrix(intCount, menuPriceCol.�ֳɱ���)) Then
                    dbl��װ = Val(.TextMatrix(intCount, menuPriceCol.��װϵ��))

                    lngҩƷID = Val(.TextMatrix(intCount, menuPriceCol.ҩƷid))
                    dblOldCost = Val(Round(Val(.TextMatrix(intCount, menuPriceCol.ԭ�ɱ���)) / dbl��װ, gtype_UserDrugDigits.Digit_�ɱ���))
                    dblNewCost = Val(Round(Val(.TextMatrix(intCount, menuPriceCol.�ֳɱ���)) / dbl��װ, gtype_UserDrugDigits.Digit_�ɱ���))

                    gstrSQL = "Zl_�ɱ��۵�����Ϣ_Insert(Null,Null," & lngҩƷID & ",0,Null,Null,Null,Null," & dblOldCost & ", " & dblNewCost & ",NULL,Null,0,0, " & txtNO.Text & ")"
                    Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
                    blnCost = True
                End If
            Next
        End With
    End If

    '����ִ��
    If mint���� = 1 Then
        '�����ɱ��۵���ʱ
        If optʱ��(0).Value = True Then
            With Me.vsfPrice
                For intCount = 1 To IIf(Trim(.TextMatrix(.rows - 1, 0)) = "", .rows - 2, .rows - 1)
                    gstrSQL = "zl_ҩƷ�շ���¼_Adjust(0,0,Null," & Val(.TextMatrix(intCount, menuPriceCol.ҩƷid)) & ")"
                    Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
                Next
            End With
        End If
    Else
        '���ۼ�
        ArrayID = Split(strID, ",")
        Array���μ۸� = Split(str���μ۸�, ";")
        For intCount = 0 To UBound(ArrayID)
            If optʱ��(0).Value = True Or vsfPrice.TextMatrix(intCount + 1, menuPriceCol.ԭ��id) = "" Then
                gstrSQL = "zl_ҩƷ�շ���¼_Adjust(" & ArrayID(intCount) & "," & Me.Chk����.Value & ",'" & Array���μ۸�(intCount) & "')"
                Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
            End If
        Next
    End If

    '����ָ���۸�
    With Me.vsfPrice
        For intCount = 1 To IIf(Trim(.TextMatrix(.rows - 1, 0)) = "", .rows - 2, .rows - 1)
            dbl��װ = Val(.TextMatrix(intCount, menuPriceCol.��װϵ��))

            '����ָ�����ۼ�
            If Val(.TextMatrix(intCount, menuPriceCol.ԭָ���ۼ�)) < Val(.TextMatrix(intCount, menuPriceCol.�����ۼ�)) And Val(.TextMatrix(intCount, menuPriceCol.ԭָ���ۼ�)) <> 0 Then
                strUpdate = Val(Round(Val(.TextMatrix(intCount, menuPriceCol.��ָ���ۼ�)) / dbl��װ, gtype_UserDrugDigits.Digit_���ۼ�))

                gstrSQL = "zl_ҩƷĿ¼_UpdateCustom(" & Val(.TextMatrix(intCount, menuPriceCol.ҩƷid)) & ",'ָ�����ۼ�=" & strUpdate & "')"
                Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
            End If

            '���²ɹ��޼�
            If Val(.TextMatrix(intCount, menuPriceCol.ԭ�ɹ��޼�)) < Val(.TextMatrix(intCount, menuPriceCol.�ֳɱ���)) And Val(.TextMatrix(intCount, menuPriceCol.ԭ�ɹ��޼�)) <> 0 Then
                strUpdate = Val(Round(Val(.TextMatrix(intCount, menuPriceCol.�ֲɹ��޼�)) / dbl��װ, gtype_UserDrugDigits.Digit_�ɱ���))

                gstrSQL = "zl_ҩƷĿ¼_UpdateCustom(" & Val(.TextMatrix(intCount, menuPriceCol.ҩƷid)) & ",'ָ��������=" & strUpdate & "')"
                Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)
            End If
        Next
    End With

    '�������ۻ��ܼ�¼
    If blnPrice = True And blnCost = True Then
        intUpdateModel = 2
    ElseIf blnPrice = True And blnCost = False Then
        intUpdateModel = 0
    ElseIf blnPrice = False And blnCost = True Then
        intUpdateModel = 1
    End If

    gstrSQL = "Zl_���ۻ��ܼ�¼_Insert(" & txtNO.Text & "," & intUpdateModel & ","
    If optʱ��(0).Value = True Then
        gstrSQL = gstrSQL & "sysdate" & ","
    Else
        gstrSQL = gstrSQL & "to_date('" & Format(Me.dtpRunDate.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
    End If
    gstrSQL = gstrSQL & IIf(txtSummary.Text = "", "Null", "'" & txtSummary.Text & "'") & ",0,'" & UserInfo.�û����� & "')"
    Call zlDataBase.ExecuteProcedure(gstrSQL, MStrCaption)

    gcnOracle.CommitTrans

    If blnPrint = True Then
        If MsgBox("����Ҫ��ӡ����֪ͨ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1333", Me, "NO=" & txtNO.Text, "��װ��λ=" & mintUnit, 2)
        End If
    End If

    '����б�������
    With vsfPrice
        .rows = 2
        For intCol = 0 To .Cols - 1
            .TextMatrix(1, intCol) = ""
        Next
    End With
    vsfStore.rows = 1
    vsfPay.rows = 1
    txtNO.Text = ""
    txtSummary.Text = ""

    Exit Sub

ErrHand:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Function CheckUnVerify(ByVal lngҩƷID As Long) As Boolean
    '���ҩƷ�Ƿ����δ��˵���
    Dim rsTemp As ADODB.Recordset

    On Error GoTo errHandle
    gstrSQL = "Select 1 From ҩƷ�շ���¼ Where ҩƷid = [1] And Rownum = 1 And ������� Is Null"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "���ҩƷ�Ƿ����δ��˵���", lngҩƷID)

    If rsTemp.RecordCount > 0 Then
        CheckUnVerify = True
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function checkNotExecutePrice(Optional ByVal lngDrugID As Long = 0, Optional ByRef strInfo As String) As Boolean
    '���� ������Ƿ����δִ�еļ۸�
    Dim RecCheck As New ADODB.Recordset
    Dim LngmediIDThis As Long, IntCheck As Integer

    Err = 0
    On Error GoTo ErrHand

    If lngDrugID = 0 Then
        'ѭ���ж�����ҩƷ
        For IntCheck = 1 To vsfPrice.rows - 1
            LngmediIDThis = Val(vsfPrice.TextMatrix(IntCheck, menuPriceCol.ҩƷid))
            If LngmediIDThis <> 0 Then
                If mint���� = 0 Or mint���� = 2 Then
                    '�ж��Ƿ���δִ�е���ʷ�۸�
                    gstrSQL = " Select Count(*) Records From �շѼ�Ŀ Where �䶯ԭ��=0 And ִ������ > Sysdate And �շ�ϸĿID=[1]"
                    Set RecCheck = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, LngmediIDThis)

                    With RecCheck
                        If Not .EOF Then
                            If Not IsNull(!Records) Then
                                If !Records <> 0 Then
                                    strInfo = "ҩƷ" & vsfPrice.TextMatrix(IntCheck, menuPriceCol.Ʒ��) & "����δִ�м۸�δִ��ҩƷ���ܵ��ۣ�"
                                    checkNotExecutePrice = True
                                    Exit Function
                                End If
                            End If
                        End If
                    End With
                End If

                If mint���� = 1 Or mint���� = 2 Then
                    '����Ƿ���δִ�еĳɱ��۵��ۼƻ�
                    gstrSQL = "Select 1 From �ɱ��۵�����Ϣ Where ҩƷid = [1] And ִ������ Is Null And Rownum = 1 "
                    Set RecCheck = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, LngmediIDThis)

                    If RecCheck.RecordCount > 0 Then
                        strInfo = "ҩƷ" & vsfPrice.TextMatrix(IntCheck, menuPriceCol.Ʒ��) & "����δִ�гɱ��ۣ�δִ��ҩƷ���ܵ��ۣ�"
                        checkNotExecutePrice = True
                        Exit Function
                    End If
                End If
            End If
        Next
    Else
        If mint���� = 0 Or mint���� = 2 Then
            '�ж��Ƿ���δִ�е���ʷ�۸�
            gstrSQL = " Select Count(*) Records From �շѼ�Ŀ Where �䶯ԭ��=0 And ִ������ > Sysdate And �շ�ϸĿID=[1]"
            Set RecCheck = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lngDrugID)

            With RecCheck
                If Not .EOF Then
                    If Not IsNull(!Records) Then
                        If !Records <> 0 Then
                            strInfo = "������δִ�е��ۼ۵��ۼ�¼��δִ��ҩƷ���ܵ��ۣ�"
                            checkNotExecutePrice = True
                            Exit Function
                        End If
                    End If
                End If
            End With
        End If

        If mint���� = 1 Or mint���� = 2 Then
            '����Ƿ���δִ�еĳɱ��۵��ۼƻ�
            gstrSQL = "Select 1 From �ɱ��۵�����Ϣ Where ҩƷid = [1] And ִ������ Is Null And Rownum = 1 "
            Set RecCheck = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lngDrugID)

            If RecCheck.RecordCount > 0 Then
                strInfo = "������δִ�еĳɱ��۵��ۣ�δִ��ҩƷ���ܵ��ۣ�"
                checkNotExecutePrice = True
                Exit Function
            End If
        End If
    End If


    checkNotExecutePrice = False
    Exit Function
ErrHand:
    Call ErrCenter
    Call SaveErrLog
    Me.vsfPrice.SetFocus

End Function

Private Function CheckPrice() As Boolean
    Dim IntCheck As Integer
    Dim n As Integer
    Dim strTmp As String
    Dim bln�޿�� As Boolean
    Dim dbl��װ As Double
    Dim bln���޿�� As Boolean
    Dim lngDouID As Long
    Dim strMsg As String '��¼��ʾ��Ϣ
    Dim intCount2 As Integer '��������
    
    '����ִ�м۸��Ƿ���ȷ
    '�Լ�������Ŀ��ͬ��������ּ��Ƿ���ԭ����ͬ
    CheckPrice = False
    With vsfPrice
        For IntCheck = 1 To .rows - 1
            If Val(.TextMatrix(IntCheck, menuPriceCol.ҩƷid)) <> 0 Then
                If Not IsNumeric(Trim(.TextMatrix(IntCheck, menuPriceCol.�����ۼ�))) Then
                    MsgBox "��" & IntCheck & "�е�ҩƷ�ۼ��к��зǷ��ַ���", vbInformation, gstrSysName
                    .Row = IntCheck
                    .Col = menuPriceCol.�����ۼ�
                    vsfPrice.SetFocus
                    .Select IntCheck, 0, IntCheck, .Cols - 1
                    .TopRow = IntCheck
                    Exit Function
                End If

                '���۸��Ƿ�Ϊ��
                If .TextMatrix(IntCheck, menuPriceCol.�����ۼ�) = "" Or .TextMatrix(IntCheck, menuPriceCol.ԭ���ۼ�) = "" Or .TextMatrix(IntCheck, menuPriceCol.�ֳɱ���) = "" Or .TextMatrix(IntCheck, menuPriceCol.ԭ�ɱ���) = "" Then
                    MsgBox "��" & IntCheck & "�е�ҩƷ�м۸�Ϊ�գ�����ִ�е��ۣ�", vbInformation, gstrSysName
                    .Row = IntCheck
                    vsfPrice.SetFocus
                    .Select IntCheck, 0, IntCheck, .Cols - 1
                    .TopRow = IntCheck
                    Exit Function
                End If
                For n = 1 To vsfStore.rows - 1
                    If Val(.TextMatrix(IntCheck, menuPriceCol.ҩƷid)) = Val(vsfStore.TextMatrix(n, menuStoreCol.ҩƷid)) Then
                        If vsfStore.TextMatrix(n, menuStoreCol.�����ۼ�) = "" Or vsfStore.TextMatrix(n, menuStoreCol.ԭ���ۼ�) = "" Or vsfStore.TextMatrix(n, menuStoreCol.�ֲɹ���) = "" Or vsfStore.TextMatrix(n, menuStoreCol.ԭ�ɹ���) = "" Then
                            MsgBox "��" & IntCheck & "�е�ҩƷ�м۸�Ϊ�գ�����ִ�е��ۣ�", vbInformation, gstrSysName
                            .Row = IntCheck
                            vsfPrice.SetFocus
                            .Select IntCheck, 0, IntCheck, .Cols - 1
                            .TopRow = IntCheck
                            Exit Function
                        End If
                    End If
                Next
                
                '����ۼ��Ƿ���ͬ
                If mint���� = 0 Or mint���� = 2 Then
                    strTmp = ""
                    bln���޿�� = False
                    dbl��װ = Val(.TextMatrix(IntCheck, menuPriceCol.��װϵ��))
                    If .TextMatrix(IntCheck, menuPriceCol.�Ƿ���) = "1" Then
                        For n = 1 To vsfStore.rows - 1
                            If Val(.TextMatrix(IntCheck, menuPriceCol.ҩƷid)) = Val(vsfStore.TextMatrix(n, menuStoreCol.ҩƷid)) Then
                                bln���޿�� = True
                                If InStr(1, "|" & strTmp, "|" & vsfStore.TextMatrix(n, menuStoreCol.����) & ",") = 0 And vsfStore.TextMatrix(n, menuStoreCol.�����ۼ�) <> vsfStore.TextMatrix(n, menuStoreCol.ԭ���ۼ�) Then
                                    strTmp = strTmp & IIf(strTmp = "", "", "|") & vsfStore.TextMatrix(n, menuStoreCol.����) & "," & vsfStore.TextMatrix(n, menuStoreCol.�����ۼ�) / dbl��װ
                                End If
                            End If
                        Next
                        If strTmp = "" And bln���޿�� = True Then
                            MsgBox "��" & IntCheck & "�е�ҩƷ�����ۼ���ԭ���ۼ���ͬ������ִ�е��ۣ�", vbInformation, gstrSysName
                            .Row = IntCheck
                            .Col = menuPriceCol.�����ۼ�
                            vsfPrice.SetFocus
                            .Select IntCheck, 0, IntCheck, .Cols - 1
                            .TopRow = IntCheck
                            Exit Function
                        End If
                        If bln���޿�� = False And .TextMatrix(IntCheck, menuPriceCol.�����ۼ�) = .TextMatrix(IntCheck, menuPriceCol.ԭ���ۼ�) Then
                            MsgBox "��" & IntCheck & "�е�ҩƷ�����ۼ���ԭ���ۼ���ͬ������ִ�е��ۣ�", vbInformation, gstrSysName
                            .Row = IntCheck
                            .Col = menuPriceCol.�����ۼ�
                            vsfPrice.SetFocus
                            .Select IntCheck, 0, IntCheck, .Cols - 1
                            .TopRow = IntCheck
                            Exit Function
                        End If
                    End If
                    If .TextMatrix(IntCheck, menuPriceCol.�Ƿ���) <> "1" And .TextMatrix(IntCheck, menuPriceCol.�����ۼ�) = .TextMatrix(IntCheck, menuPriceCol.ԭ���ۼ�) Then
                        MsgBox "��" & IntCheck & "�е�ҩƷ�����ۼ���ԭ���ۼ���ͬ������ִ�е��ۣ�", vbInformation, gstrSysName
                        .Row = IntCheck
                        .Col = menuPriceCol.�����ۼ�
                        vsfPrice.SetFocus
                        .Select IntCheck, 0, IntCheck, .Cols - 1
                        .TopRow = IntCheck
                        Exit Function
                    End If
                End If
                
                '���ɱ����Ƿ���ͬ
                If mint���� = 1 Or mint���� = 2 Then
                    bln���޿�� = False
                    strTmp = ""
                    For n = 1 To vsfStore.rows - 1
                        If Val(.TextMatrix(IntCheck, menuPriceCol.ҩƷid)) = Val(vsfStore.TextMatrix(n, menuStoreCol.ҩƷid)) Then
                            bln���޿�� = True
                            If vsfStore.TextMatrix(n, menuStoreCol.�ֲɹ���) <> vsfStore.TextMatrix(n, menuStoreCol.ԭ�ɹ���) Then
                                strTmp = "�����ɱ���"
                            End If
                        End If
                    Next
                    If bln���޿�� = True And strTmp = "" Then
                        MsgBox "��" & IntCheck & "�е�ҩƷ�ֲɹ�����ԭ�ɹ�����ͬ������ִ�е��ۣ�", vbInformation, gstrSysName
                        .Row = IntCheck
                        .Col = menuPriceCol.�ֳɱ���
                        vsfPrice.SetFocus
                        .Select IntCheck, 0, IntCheck, .Cols - 1
                        .TopRow = IntCheck
                        Exit Function
                    End If
                    If bln���޿�� = False And .TextMatrix(IntCheck, menuPriceCol.�ֳɱ���) = .TextMatrix(IntCheck, menuPriceCol.ԭ�ɱ���) Then
                        MsgBox "��" & IntCheck & "�е�ҩƷ�ֳɱ�����ԭ�ɱ�����ͬ������ִ�е��ۣ�", vbInformation, gstrSysName
                        .Row = IntCheck
                        .Col = menuPriceCol.�ֳɱ���
                        vsfPrice.SetFocus
                        .Select IntCheck, 0, IntCheck, .Cols - 1
                        .TopRow = IntCheck
                        Exit Function
                    End If
                End If

                If .TextMatrix(IntCheck, menuPriceCol.�Ƿ���) = "1" And optʱ��(0).Value <> True And mint���� <> 1 Then
                    MsgBox "��" & IntCheck & "��Ϊʱ��ҩƷ����������Ϊ����ִ�У�", vbInformation, gstrSysName
                    .Row = IntCheck
                    .Col = menuPriceCol.�����ۼ�
                    vsfPrice.SetFocus
                    .Select IntCheck, 0, IntCheck, .Cols - 1
                    .TopRow = IntCheck
                    Exit Function
                End If

            End If
        Next
    End With

    '���δ��˵���
    If vsfStore.rows > 1 And (mint���� = 1 Or mint���� = 2) Then
        If vsfStore.TextMatrix(1, menuStoreCol.ҩƷid) <> "" Then
            lngDouID = 0
            For n = 1 To vsfStore.rows - 1
                If vsfStore.TextMatrix(n, menuStoreCol.ҩƷid) = "" Then Exit For
    
                If CheckUnVerify(Val(vsfStore.TextMatrix(n, menuStoreCol.ҩƷid))) = True And Val(vsfStore.TextMatrix(n, menuStoreCol.ҩƷid)) <> lngDouID Then
                    lngDouID = Val(vsfStore.TextMatrix(n, menuStoreCol.ҩƷid))
                    strMsg = strMsg & vsfStore.TextMatrix(n, menuStoreCol.ҩƷ) & ","
                    intCount2 = intCount2 + 1
                    If intCount2 > 3 Then Exit For 'ֻ�ж�3��
                End If
            Next
    
            If strMsg <> "" Then
                If MsgBox(strMsg & "����δ��˵��ݣ������ɱ��ۿ��ܻ���ɲ����" & _
                    vbCrLf & Space(4) & "�����ȴ���δ��˵��ݡ��Ƿ񻹼������ۣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            End If
        End If
    End If
    
    CheckPrice = True
End Function


Private Sub cmdPriceMethod_Click()
    If txt��Ӧ��.Tag = "" Then
        Me.txt��Ӧ��.Tag = "0|"
    End If
    picOtherSelect.Visible = True
End Sub

Private Sub cmdPrint_Click()
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    If vsfStore.rows = 1 Then Exit Sub
    If Trim(Me.vsfStore.TextMatrix(1, menuStoreCol.�ⷿ)) = "" Then Exit Sub

    objPrint.Title.Text = "���ۿ��䶯��"

    Set objRow = New zlTabAppRow
    objRow.Add "����˵��:" & Me.txtSummary.Text
    objPrint.UnderAppRows.Add objRow

    Set objRow = New zlTabAppRow
    objRow.Add "ִ��ʱ��:" & Format(IIf(optʱ��(0).Value = True, zlDataBase.Currentdate, Me.dtpRunDate.Value), "yyyy��MM��DD�� HH:mm:ss")
    objRow.Add "������:" & Me.txtValuer.Text
    objPrint.UnderAppRows.Add objRow

    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & gstrUserName
    objRow.Add "��ӡʱ��:" & Format(zlDataBase.Currentdate, "yyyy��MM��DD�� HH:mm:ss")
    objPrint.BelowAppRows.Add objRow

    Set objPrint.Body = Me.vsfStore.Object
    objPrint.PageFooter = 2

    Select Case zlPrintAsk(objPrint)
    Case 1
         zlPrintOrView1Grd objPrint, 1
    Case 2
        zlPrintOrView1Grd objPrint, 2
    Case 3
        zlPrintOrView1Grd objPrint, 3
    End Select
    Set objPrint = Nothing
End Sub

Private Sub Cmd��Ӧ��_Click()
    Dim rsTemp As ADODB.Recordset

    On Error GoTo errHandle
    gstrSQL = "Select ����,����,����,id" & _
        " From ��Ӧ��" & _
        " where ĩ��=1 And substr(����,1,1) = '1' And (����ʱ�� is null or ����ʱ��=to_date('3000-01-01','YYYY-MM-DD')) " & _
        " Order By ���� "
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "ȡ��Ӧ����Ϣ")
    If rsTemp.EOF Then
        MsgBox "���ʼ����Ӧ�̣��ֵ������", vbInformation, gstrSysName
        Exit Sub
    End If

    With Me.mshProvider
        .Left = chk��Ӧ��.Left
        .Top = txt��Ӧ��.Top + txt��Ӧ��.Height
        .Clear
        Set .DataSource = rsTemp
        .ColWidth(0) = 800: .ColWidth(1) = 2500: .ColWidth(2) = 800: .ColWidth(3) = 0
        .Row = 1: .ColSel = .Cols - 1
        .ZOrder 0: .Visible = True: .SetFocus
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub Form_Activate()
    If mblnLoad = False Then
        vsfPrice.SetFocus
    End If
    If mBlnClick = False Then
        vsfPrice.Row = 1
        vsfPrice.Col = menuPriceCol.Ʒ��
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        picOtherSelect.Visible = False
    End If
End Sub

Private Sub Form_Load()
    Dim StrToday As String
    Dim intUnitTemp As Integer

    Me.Height = 768 * 15
    Me.Width = 1024 * 15
    '��ȡ���õĵ�λ
    mintUnit = Val(zlDataBase.GetPara("ҩƷ��λ", glngSys, 1333, "1"))
    mstrPrivs = GetPrivFunc(glngSys, 1333)
    Select Case mintUnit
        Case 0 'ҩ��
            intUnitTemp = 4
        Case 1 'סԺ
            intUnitTemp = 3
        Case 2 '����
            intUnitTemp = 2
        Case 3 '�ۼ�
            intUnitTemp = 1
    End Select
    '��ȡ������λ����
    mintCostDigit = GetDigitTiaoJia(1, 1, intUnitTemp)
    mintPriceDigit = GetDigitTiaoJia(1, 2, intUnitTemp)
    mintNumberDigit = GetDigitTiaoJia(1, 3, intUnitTemp)
    mintMoneyDigit = GetDigitTiaoJia(1, 4)
    mstrMoneyFormat = "0." & String(mintMoneyDigit, "0")
    mintSalePriceDigit = GetDigitTiaoJia(1, 2, 1)
    '��ʼ��ʱ��Ϊ��ǰʱ��+1��
    StrToday = Format(zlDataBase.Currentdate(), "yyyy-MM-dd hh:mm:ss")
    
    If mintModal = 0 Then '������ʱ����Сʱ������Ϊ��ǰʱ��+1��
        Me.dtpRunDate.MinDate = DateAdd("s", 1, CDate(StrToday))
    End If
    Me.dtpRunDate.Value = DateAdd("d", 1, CDate(StrToday))

    mblnʱ��ҩƷ�����ε��� = Val(zlDataBase.GetPara("ʱ��ҩƷ�����ε���", glngSys, 1333, 0))
    mbln�ּ���ʾ = Val(zlDataBase.GetPara("�޼���ʾ", glngSys, 1333, 1))
    
    marrSql = Array()
    
    txtValuer.Text = UserInfo.�û�����  'gstrUserName

    txtNO.Text = IIf(mintModal = 0, "", mstr���ۻ��ܺ�)
    If mintModal = 0 Then
        lblNO.Visible = False
        txtNO.Visible = False
    End If

    Call initComboBox '��ʼ�������ؼ�
    If mintModal = 1 Then '�޸�
        If (InStr(1, ";" & mstrPrivs & ";", ";�ɱ��۵���;") > 0 And InStr(1, ";" & mstrPrivs & ";", ";�ۼ۵���;") = 0) Or (InStr(1, ";" & mstrPrivs & ";", ";�ɱ��۵���;") = 0 And InStr(1, ";" & mstrPrivs & ";", ";�ۼ۵���;") > 0) Then
            cboPriceMethod.ListIndex = 0
        ElseIf (InStr(1, ";" & mstrPrivs & ";", ";�ɱ��۵���;") > 0 And InStr(1, ";" & mstrPrivs & ";", ";�ۼ۵���;") > 0) Then
            cboPriceMethod.ListIndex = mintMethod
        End If
    ElseIf mintModal = 2 Then '����
        cboPriceMethod.ListIndex = mintMethod
    End If

    Call InitTabControl
    Call InitVsfGridFlex

    Call RestoreWinState(Me, App.ProductName, MStrCaption)
    If mblnӦ����¼ = False Then
        TabCtlDetails.Item(1).Visible = False
    End If
    If mintModal <> 0 Then
        Call initGrid
    End If

    If mintModal = 2 Then '����
        cboPriceMethod.Enabled = False
        cmdPriceMethod.Enabled = False
        optʱ��(0).Enabled = False
        optʱ��(1).Enabled = False
        dtpRunDate.Enabled = False
        cbo�ۼۼ��㷽ʽ.Enabled = False
        Chk����.Enabled = False
        chkCostBatch.Enabled = False
        chkAotuCost.Enabled = False
        chkAutoPay.Enabled = False
        txtSummary.Enabled = False
        cmdClear.Visible = False
        cmdItem.Visible = False
        cmdOk.Visible = False
        vsfPrice.Cell(flexcpBackColor, 1, 0, vsfPrice.rows - 1, vsfPrice.Cols - 1) = mconlngColor
        If vsfStore.rows > 1 Then
            vsfStore.Cell(flexcpBackColor, 1, 0, vsfStore.rows - 1, vsfStore.Cols - 1) = mconlngColor
        End If
        If vsfPay.rows > 1 Then
            vsfPay.Cell(flexcpBackColor, 0, 0, vsfPay.rows - 1, vsfPay.Cols - 1) = mconlngColor
        End If
    End If
    mblnLoad = True
End Sub

Private Sub initComboBox()
    With cbo�ۼۼ��㷽ʽ
        .AddItem "�ۼ���ɱ��۲���������"
        .AddItem "�ۼ۰��̶���������"
        .AddItem "�ۼ۰��ֶμӳɼ���"
        .ListIndex = 0
    End With

    With cboPriceMethod
        If mintModal <> 2 Then  '�ǲ���
            If InStr(1, ";" & mstrPrivs & ";", ";�ɱ��۵���;") > 0 And InStr(1, ";" & mstrPrivs & ";", ";�ۼ۵���;") = 0 Then
                .AddItem "�����ɱ���"
                .ListIndex = 0
                lblMethod.Tag = 0
            ElseIf InStr(1, ";" & mstrPrivs & ";", ";�ɱ��۵���;") = 0 And InStr(1, ";" & mstrPrivs & ";", ";�ۼ۵���;") > 0 Then
                .AddItem "�����ۼ�"
                .ListIndex = 0
                lblMethod.Tag = 0
            ElseIf InStr(1, ";" & mstrPrivs & ";", ";�ɱ��۵���;") > 0 And InStr(1, ";" & mstrPrivs & ";", ";�ۼ۵���;") > 0 Then
                .AddItem "�����ۼ�"
                .AddItem "�����ɱ���"
                .AddItem "�ۼ۳ɱ���һ�����"
                .ListIndex = 0
                lblMethod.Tag = 0
            End If
        Else
            .AddItem "�����ۼ�"
            .AddItem "�����ɱ���"
            .AddItem "�ۼ۳ɱ���һ�����"
            .ListIndex = 0
            lblMethod.Tag = 0
        End If
    End With
End Sub

Private Sub InitTabControl()
    '��ʼ��TabControl�ؼ�
    Dim objtabctl As TabControlItem

    picSplit.Left = 0
    picSplit.Top = vsfPrice.Top + vsfPrice.Height + 5
    With TabCtlDetails
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        .InsertItem 0, "���䶯��", vsfStore.hWnd, 0
        .InsertItem 1, "Ӧ����䶯��", vsfPay.hWnd, 0
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - vsfPrice.Height - vsfPrice.Top - 20
        .Top = picSplit.Height + picSplit.Top + 20
        .Item(1).Selected = True
        .Item(0).Selected = True
    End With
End Sub

Private Sub Form_Resize()

    On Error Resume Next

    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.WindowState <> vbMaximized Then
        If Me.Height < 8145 Then
            Me.Height = 8145
        End If
    End If

    With fraCondition
        .Width = Me.ScaleWidth
    End With
    txtNO.Left = Me.ScaleWidth - txtNO.Width
    lblNO.Left = txtNO.Left - lblNO.Width - 200
    lblDrugName.Left = Me.ScaleWidth / 2 - lblDrugName.Width / 2
    vsfPrice.Move 20, fraCondition.Top + fraCondition.Height + 20, Me.ScaleWidth, 3000
    picSplit.Left = 50
    picSplit.Top = vsfPrice.Top + vsfPrice.Height + 5
    picSplit.Width = Me.ScaleWidth
    txtSummary.Width = Me.ScaleWidth - lblSummary.Left - lblSummary.Width - 300
    TabCtlDetails.Move 20, picSplit.Height + picSplit.Top, Me.ScaleWidth, Me.ScaleHeight - picSplit.Top - picSplit.Height - picInfo.Height - cmdClear.Height - 300 - stbThis.Height
    picInfo.Move 0, TabCtlDetails.Top + TabCtlDetails.Height, Me.ScaleWidth
    lblFind.Top = picInfo.Top + picInfo.Height + 180
    lblFind.Left = picInfo.Left + 380
    txtFind.Top = lblFind.Top - 50
    txtFind.Left = 985
    cmdClear.Top = txtFind.Top
    cmdItem.Top = txtFind.Top
    cmdPrint.Top = txtFind.Top
    cmdOk.Top = txtFind.Top
    cmdCanc.Top = txtFind.Top
    cmdCanc.Left = Me.ScaleWidth - cmdCanc.Width - 300
    cmdOk.Left = cmdCanc.Left - cmdOk.Width - 200
    cmdPrint.Left = cmdOk.Left - cmdPrint.Width - 500
    cmdItem.Left = cmdPrint.Left - cmdPrint.Width - 20
    cmdClear.Left = cmdItem.Left - cmdItem.Width - 20
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ReleaseSelectorRS
    Call SaveWinState(Me, App.ProductName, MStrCaption)
    mblnLoad = False
    mblnӦ����¼ = False
    mlng��Ӧ��ID = 0
    mblnUpdateAdd = False
End Sub

Private Sub mshProvider_DblClick()
    With Me.mshProvider
        Me.txt��Ӧ��.Text = .TextMatrix(.Row, 1)
        Me.txt��Ӧ��.Tag = .TextMatrix(.Row, 3) & "|" & .TextMatrix(.Row, 1)
        .Visible = False
    End With

    Me.txt��Ӧ��.SetFocus
End Sub

Private Sub optʱ��_Click(Index As Integer)
    If Index = 0 Then
        dtpRunDate.Enabled = False
    Else
        dtpRunDate.Enabled = True
    End If
End Sub

Private Sub InitVsfGridFlex()
    With vsfPrice

        .Cols = menuPriceCol.������
        .rows = 2
        .RowHeight(1) = mlngRowHeight
        .ColWidth(0) = 200
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = 50
        .RowHeight(0) = mlngRowHeight
        .AllowSelection = False '���ܶ�ѡ
'        .SelectionMode = flexSelectionByRow '����ѡ��
        .ExplorerBar = flexExMoveRows '�϶�
        .AllowUserResizing = flexResizeBoth  '���Ըı����п��
        .Editable = flexEDNone
'        .GridLineWidth = 2
'        .GridLines = flexGridInset
'        .GridColor = &H80000011
'        .GridColorFixed = &H80000011
'        .ForeColorFixed = &H80000012
'        .BackColorSel = &HF4F4EA

        .TextMatrix(0, menuPriceCol.ҩƷid) = "ҩƷID"
        .TextMatrix(0, menuPriceCol.ԭ��id) = "ԭ��id"
        .TextMatrix(0, menuPriceCol.Ʒ��) = "Ʒ��"
        .TextMatrix(0, menuPriceCol.���) = "���"
        .TextMatrix(0, menuPriceCol.�Ƿ���) = "�Ƿ���"
        .TextMatrix(0, menuPriceCol.����) = "����"
        .TextMatrix(0, menuPriceCol.��λ) = "��λ"
        .TextMatrix(0, menuPriceCol.��װϵ��) = "��װϵ��"
        .TextMatrix(0, menuPriceCol.�ӳ���) = "�ӳ���"
        .TextMatrix(0, menuPriceCol.���������) = "���������"
        .TextMatrix(0, menuPriceCol.�Ƿ��п��) = "�Ƿ��п��"
        .TextMatrix(0, menuPriceCol.������ĿID) = "������Ŀid"
        .TextMatrix(0, menuPriceCol.ԭ�ɱ���) = "ԭ�ɱ���"
        .TextMatrix(0, menuPriceCol.�ֳɱ���) = "�ֳɱ���"
        .TextMatrix(0, menuPriceCol.ԭ���ۼ�) = "ԭ���ۼ�"
        .TextMatrix(0, menuPriceCol.�����ۼ�) = "�����ۼ�"
        .TextMatrix(0, menuPriceCol.ԭ�ɹ��޼�) = "ԭ�ɹ��޼�"
        .TextMatrix(0, menuPriceCol.�ֲɹ��޼�) = "�ֲɹ��޼�"
        .TextMatrix(0, menuPriceCol.ԭָ���ۼ�) = "ԭָ���ۼ�"
        .TextMatrix(0, menuPriceCol.��ָ���ۼ�) = "��ָ���ۼ�"

        '�����п�
        .ColWidth(menuPriceCol.ҩƷid) = 0
        .ColWidth(menuPriceCol.ԭ��id) = 0
        .ColWidth(menuPriceCol.Ʒ��) = 3000
        .ColWidth(menuPriceCol.���) = 1500
        .ColWidth(menuPriceCol.�Ƿ���) = 0
        .ColWidth(menuPriceCol.����) = 2000
        .ColWidth(menuPriceCol.��λ) = 800
        .ColWidth(menuPriceCol.��װϵ��) = 0
        .ColWidth(menuPriceCol.�ӳ���) = 0
        .ColWidth(menuPriceCol.���������) = 0
        .ColWidth(menuPriceCol.�Ƿ��п��) = 0
        .ColWidth(menuPriceCol.������ĿID) = 0
        .ColWidth(menuPriceCol.ԭ�ɱ���) = 1000
        .ColWidth(menuPriceCol.�ֳɱ���) = 1000
        .ColWidth(menuPriceCol.ԭ���ۼ�) = 1000
        .ColWidth(menuPriceCol.�����ۼ�) = 1000
        .ColWidth(menuPriceCol.ԭ�ɹ��޼�) = 0
        .ColWidth(menuPriceCol.�ֲɹ��޼�) = 0
        .ColWidth(menuPriceCol.ԭָ���ۼ�) = 0
        .ColWidth(menuPriceCol.��ָ���ۼ�) = 0
        '���ö��뷽ʽ
        .ColAlignment(menuPriceCol.Ʒ��) = flexAlignLeftCenter
        .ColAlignment(menuPriceCol.���) = flexAlignLeftCenter
        .ColAlignment(menuPriceCol.����) = flexAlignLeftCenter
        .ColAlignment(menuPriceCol.��λ) = flexAlignCenterCenter
        .ColAlignment(menuPriceCol.ԭ�ɱ���) = flexAlignRightCenter
        .ColAlignment(menuPriceCol.�ֳɱ���) = flexAlignRightCenter
        .ColAlignment(menuPriceCol.ԭ���ۼ�) = flexAlignRightCenter
        .ColAlignment(menuPriceCol.�����ۼ�) = flexAlignRightCenter
        .ColAlignment(menuPriceCol.ԭ�ɹ��޼�) = flexAlignRightCenter
        .ColAlignment(menuPriceCol.ԭָ���ۼ�) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter '��ͷ���ж���
        .ColComboList(menuPriceCol.Ʒ��) = "|..."
    End With

    With vsfStore
        .Editable = flexEDNone
        .Cols = menuStoreCol.������
        .rows = 1
        .ColWidth(0) = 200
'        .RowHeight(1) = mlngRowHeight
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = 50
        .RowHeight(0) = mlngRowHeight
        .AllowSelection = False '���ܶ�ѡ
'        .SelectionMode = flexSelectionByRow '����ѡ��
        .ExplorerBar = flexExMoveRows '�϶�
        .AllowUserResizing = flexResizeBoth  '���Ըı����п��
        .GridLineWidth = 2
        .GridLines = flexGridInset
        .GridColor = &H0&

        '��������
        .TextMatrix(0, menuStoreCol.ҩƷid) = "ҩƷid"
        .TextMatrix(0, menuStoreCol.�ⷿ) = "�ⷿ"
        .TextMatrix(0, menuStoreCol.�ⷿid) = "�ⷿid"
        .TextMatrix(0, menuStoreCol.��Ӧ��) = "��Ӧ��"
        .TextMatrix(0, menuStoreCol.��Ӧ��id) = "��Ӧ��id"
        .TextMatrix(0, menuStoreCol.ҩƷ) = "ҩƷ"
        .TextMatrix(0, menuStoreCol.���) = "���"
        .TextMatrix(0, menuStoreCol.��λ) = "��λ"
        .TextMatrix(0, menuStoreCol.����) = "����"
        .TextMatrix(0, menuStoreCol.Ч��) = "Ч��"
        .TextMatrix(0, menuStoreCol.����) = "����"
        .TextMatrix(0, menuStoreCol.����) = "����"
        .TextMatrix(0, menuStoreCol.��װϵ��) = "��װϵ��"
        .TextMatrix(0, menuStoreCol.����) = "����"
        .TextMatrix(0, menuStoreCol.���) = "���"
        .TextMatrix(0, menuStoreCol.ԭ���ۼ�) = "ԭ���ۼ�"
        .TextMatrix(0, menuStoreCol.�����ۼ�) = "�����ۼ�"
        .TextMatrix(0, menuStoreCol.�������) = "�������"
        .TextMatrix(0, menuStoreCol.�ӳ���) = "�ӳ���"
        .TextMatrix(0, menuStoreCol.ԭ�ɹ���) = "ԭ�ɹ���"
        .TextMatrix(0, menuStoreCol.�ֲɹ���) = "�ֲɹ���"
        .TextMatrix(0, menuStoreCol.��۲�) = "��۲�"
        '�����п�
        .ColWidth(0) = 0
        .ColWidth(menuStoreCol.�ⷿ) = 1500
        .ColWidth(menuStoreCol.�ⷿid) = 0
        .ColWidth(menuStoreCol.��Ӧ��) = 2000
        .ColWidth(menuStoreCol.��Ӧ��id) = 0
        .ColWidth(menuStoreCol.ҩƷ) = 3000
        .ColWidth(menuStoreCol.���) = 1500
        .ColWidth(menuStoreCol.��λ) = 800
        .ColWidth(menuStoreCol.����) = 1500
        .ColWidth(menuStoreCol.Ч��) = 2000
        .ColWidth(menuStoreCol.����) = 1500
        .ColWidth(menuStoreCol.����) = 1500
        .ColWidth(menuStoreCol.��װϵ��) = 0
        .ColWidth(menuStoreCol.����) = 0
        .ColWidth(menuStoreCol.���) = 0
        .ColWidth(menuStoreCol.ԭ���ۼ�) = 1000
        .ColWidth(menuStoreCol.�����ۼ�) = 1000
        .ColWidth(menuStoreCol.�������) = 1000
        .ColWidth(menuStoreCol.�ӳ���) = 1000
        .ColWidth(menuStoreCol.ԭ�ɹ���) = 1000
        .ColWidth(menuStoreCol.�ֲɹ���) = 1000
        .ColWidth(menuStoreCol.��۲�) = 1000
        '���뷽ʽ
        .ColAlignment(menuStoreCol.�ⷿ) = flexAlignLeftCenter
        .ColAlignment(menuStoreCol.��Ӧ��) = flexAlignLeftCenter
        .ColAlignment(menuStoreCol.ҩƷ) = flexAlignLeftCenter
        .ColAlignment(menuStoreCol.���) = flexAlignLeftCenter
        .ColAlignment(menuStoreCol.��λ) = flexAlignCenterCenter
        .ColAlignment(menuStoreCol.����) = flexAlignLeftCenter
        .ColAlignment(menuStoreCol.Ч��) = flexAlignLeftCenter
        .ColAlignment(menuStoreCol.����) = flexAlignLeftCenter
        .ColAlignment(menuStoreCol.����) = flexAlignRightCenter
        .ColAlignment(menuStoreCol.ԭ���ۼ�) = flexAlignRightCenter
        .ColAlignment(menuStoreCol.�����ۼ�) = flexAlignRightCenter
        .ColAlignment(menuStoreCol.�������) = flexAlignRightCenter
        .ColAlignment(menuStoreCol.�ӳ���) = flexAlignRightCenter
        .ColAlignment(menuStoreCol.ԭ�ɹ���) = flexAlignRightCenter
        .ColAlignment(menuStoreCol.�ֲɹ���) = flexAlignRightCenter
        .ColAlignment(menuStoreCol.��۲�) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter '��ͷ���ж���
    End With

    With vsfPay
        .Editable = flexEDNone
        .Cols = menuPayCol.������
        .rows = 1
        .ColWidth(0) = 200
'        .RowHeight(1) = mlngRowHeight
        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = 50
        .RowHeight(0) = mlngRowHeight
        .AllowSelection = False '���ܶ�ѡ
'        .SelectionMode = flexSelectionByRow '����ѡ��
        .ExplorerBar = flexExMoveRows '�϶�
        .AllowUserResizing = flexResizeBoth  '���Ըı����п��
        .GridLineWidth = 2
        .GridLines = flexGridInset
        .GridColor = &H0&

        .TextMatrix(0, menuPayCol.ҩƷid) = "ҩƷid"
        .TextMatrix(0, menuPayCol.Ʒ��) = "Ʒ��"
        .TextMatrix(0, menuPayCol.��Ʊ��) = "��Ʊ��"
        .TextMatrix(0, menuPayCol.��Ʊ����) = "��Ʊ����"
        .TextMatrix(0, menuPayCol.��Ʊ���) = "��Ʊ���"
        '�����п�
        .ColWidth(menuPayCol.ҩƷid) = 0
        .ColWidth(menuPayCol.Ʒ��) = 2000
        .ColWidth(menuPayCol.��Ʊ��) = 1500
        .ColWidth(menuPayCol.��Ʊ����) = 2000
        .ColWidth(menuPayCol.��Ʊ���) = 1500
        '���뷽ʽ
        .ColAlignment(menuPayCol.Ʒ��) = flexAlignLeftCenter
        .ColAlignment(menuPayCol.��Ʊ��) = flexAlignLeftCenter
        .ColAlignment(menuPayCol.��Ʊ����) = flexAlignLeftCenter
        .ColAlignment(menuPayCol.��Ʊ���) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter '��ͷ���ж���
    End With
End Sub

Private Sub initGrid()
    '������޸Ļ��߲�������ȡ��Ӧ�ļ�¼����䵽�����
    Dim rsTemp As ADODB.Recordset
    Dim intRow As Long
    Dim i As Long
    Dim lngDrugID As Long
    Dim db��װϵ�� As Double
    Dim strUnit As String
    Dim StrToday As String
    Dim rs���� As ADODB.Recordset

    On Error GoTo errHandle
    '���۷�ʽ 0-���ۼ�;1-���ɱ���;2-���ۼۼ��ɱ���
    If mintMethod = 0 Then
        gstrSQL = "Select Distinct p.ԭ��id, i.�Ƿ���, Nvl(s.ָ��������, 0) As ָ������, Nvl(s.����, 0) As ����, Nvl(s.ָ�����ۼ�, 0) As ָ���ۼ�," & vbNewLine & _
            "                nvl(s.�ӳ���,0) / 100 As �ӳ���, i.����, b.���� As ��Ʒ��, i.���� As ͨ����, i.���, i.���� As ����, i.���㵥λ As ��λ," & vbNewLine & _
            "                s.���ﵥλ, s.�����װ, s.סԺ��λ, s.סԺ��װ, s.ҩ�ⵥλ, Nvl(s.ҩ���װ, 1) ҩ���װ, s.�ɱ��� As ԭ�ɱ���, s.�ɱ��� As �³ɱ���, p.ԭ��, p.�ּ�," & vbNewLine & _
            "                p.������Ŀid, p.������, p.����˵��, s.���������, To_Char(a.ִ������, 'YYYY-MM-DD HH24:MI:SS') As ִ������, i.Id ҩƷid," & vbNewLine & _
            "                Decode(k.ҩƷid, Null, 0, 1) �Ƿ��п��" & vbNewLine & _
            "From (Select ҩƷid From ҩƷ��� where ����=1) K, ���ۻ��ܼ�¼ A, �շ���Ŀ���� B, ҩƷ��� S, �շ���ĿĿ¼ I, �շѼ�Ŀ P" & vbNewLine & _
            "Where a.���ۺ� = p.���ۻ��ܺ� And b.�շ�ϸĿid(+) = s.ҩƷid And s.ҩƷid = i.Id And i.Id = k.ҩƷid(+) And i.Id = p.�շ�ϸĿid And" & vbNewLine & _
            "      p.���ۻ��ܺ� = [1] And a.���� = 0 And b.����(+) = 3 And a.���ۺ� = [1] " & vbNewLine & _
            IIf(mintModal = 2, "", "  And (i.����ʱ�� Is Null Or i.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))") & vbNewLine & _
            "Order By ҩƷid"
    ElseIf mintMethod = 1 Then
        gstrSQL = "Select Distinct i.�Ƿ���, Nvl(s.ָ��������, 0) As ָ������, Nvl(s.����, 0) As ����, Nvl(s.ָ�����ۼ�, 0) As ָ���ۼ�," & vbNewLine & _
            "                nvl(s.�ӳ���,0) / 100As �ӳ���, i.����, b.���� As ��Ʒ��, i.���� As ͨ����, i.���, i.���� As ����, i.���㵥λ As ��λ," & vbNewLine & _
            "                s.���ﵥλ, s.�����װ, s.סԺ��λ, s.סԺ��װ, s.ҩ�ⵥλ, Nvl(s.ҩ���װ, 1) ҩ���װ, m.ԭ�ɱ���, m.�³ɱ���, p.�ּ� as ԭ��, p.�ּ�, p.������Ŀid," & vbNewLine & _
            "                a.������ As ������, a.˵�� As ����˵��, s.���������, To_Char(m.ִ������, 'YYYY-MM-DD HH24:MI:SS') As ִ������, i.Id ҩƷid," & vbNewLine & _
            "                Decode(k.ҩƷid, Null, 0, 1) �Ƿ��п��" & vbNewLine & _
            "From (Select Min(ԭ�ɱ���) As ԭ�ɱ���, Min(�³ɱ���) As �³ɱ���, min(����) as ����,���ۻ��ܺ�,ҩƷid,min(ִ������) as ִ������ From �ɱ��۵�����Ϣ Where ���ۻ��ܺ� = [1] Group By ���ۻ��ܺ�,ҩƷid) M, (Select ҩƷid From ҩƷ��� where ����=1) K, ���ۻ��ܼ�¼ A, �շ���Ŀ���� B, ҩƷ��� S, �շ���ĿĿ¼ I, �շѼ�Ŀ P" & vbNewLine & _
            "Where m.���ۻ��ܺ�(+) = a.���ۺ� And b.�շ�ϸĿid(+) = s.ҩƷid And s.ҩƷid = i.Id And i.Id = k.ҩƷid(+) And m.ҩƷid = i.Id And" & vbNewLine & _
            "      i.Id = p.�շ�ϸĿid And Sysdate Between p.ִ������ And p.��ֹ���� And m.���ۻ��ܺ� = [1] And a.���� = 0 And b.����(+) = 3 And" & vbNewLine & _
            "      a.���ۺ� = [1] " & IIf(mintModal = 2, "", " And (i.����ʱ�� Is Null Or i.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))") & vbNewLine & _
            "Order By ҩƷid"
    ElseIf mintMethod = 2 Then
        gstrSQL = "Select distinct p.ԭ��id, i.�Ƿ���, Nvl(s.ָ��������, 0) As ָ������, Nvl(s.����, 0) As ����, Nvl(s.ָ�����ۼ�, 0) As ָ���ۼ�," & vbNewLine & _
            "       nvl(s.�ӳ���,0) / 100 As �ӳ���, i.����, b.���� As ��Ʒ��, i.���� As ͨ����, i.���, i.���� As ����, i.���㵥λ As ��λ, s.���ﵥλ," & vbNewLine & _
            "       s.�����װ, s.סԺ��λ, s.סԺ��װ, s.ҩ�ⵥλ, Nvl(s.ҩ���װ, 1) ҩ���װ, m.ԭ�ɱ���, m.�³ɱ���, p.ԭ��, p.�ּ�, p.������Ŀid, p.������, p.����˵��, s.���������," & vbNewLine & _
            "       To_Char(p.ִ������, 'YYYY-MM-DD HH24:MI:SS') As ִ������, i.Id ҩƷid, Decode(k.ҩƷid, Null, 0, 1) �Ƿ��п��" & vbNewLine & _
            "From (Select ҩƷid,Min(ԭ�ɱ���) As ԭ�ɱ���, Min(�³ɱ���) As �³ɱ���, min(����) as ����,���ۻ��ܺ� From �ɱ��۵�����Ϣ Where ���ۻ��ܺ� = [1] Group By ҩƷid,���ۻ��ܺ�) M, �շѼ�Ŀ P, ���ۻ��ܼ�¼ A, (Select ҩƷid From ҩƷ��� where ����=1) K, �շ���Ŀ���� B, ҩƷ��� S, �շ���ĿĿ¼ I" & vbNewLine & _
            "Where m.���ۻ��ܺ� = a.���ۺ� and m.ҩƷid=i.id And p.���ۻ��ܺ� = a.���ۺ� And p.�շ�ϸĿid = k.ҩƷid(+) And p.�շ�ϸĿid = b.�շ�ϸĿid(+) And p.�շ�ϸĿid = s.ҩƷid And" & vbNewLine & _
            "      s.ҩƷid = i.Id And a.���ۺ� =[1] And b.����(+) = 3 " & vbNewLine & _
            IIf(mintModal = 2, "", "  And (i.����ʱ�� Is Null Or i.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))") & "Order By ҩƷid "
    End If
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, mstr���ۻ��ܺ�)
    If rsTemp.RecordCount = 0 Then
        MsgBox "�õ��ۼ�¼�Ѿ���ɾ���ˣ�", vbInformation, gstrSysName
        Exit Sub
    End If

    With vsfPrice
        .rows = 2
        rsTemp.MoveFirst
        For i = 0 To rsTemp.RecordCount - 1
            If rsTemp!ҩƷid <> lngDrugID Then
                Select Case mintUnit
                    Case 0
                        db��װϵ�� = rsTemp!ҩ���װ
                        strUnit = rsTemp!ҩ�ⵥλ
                    Case 1
                        db��װϵ�� = rsTemp!סԺ��װ
                        strUnit = rsTemp!סԺ��λ
                    Case 2
                        db��װϵ�� = rsTemp!�����װ
                        strUnit = rsTemp!���ﵥλ
                    Case 3
                        db��װϵ�� = 1
                        strUnit = rsTemp!��λ
                End Select

                lngDrugID = rsTemp!ҩƷid
                If mintMethod = 0 Or mintMethod = 2 Then
                    .TextMatrix(.rows - 1, menuPriceCol.ԭ��id) = IIf(IsNull(rsTemp!ԭ��id), "", rsTemp!ԭ��id)
                End If
                .TextMatrix(.rows - 1, menuPriceCol.ҩƷid) = rsTemp!ҩƷid

                If gintҩƷ������ʾ = 1 Then
                    .TextMatrix(.rows - 1, menuPriceCol.Ʒ��) = "[" & rsTemp!���� & "]" & IIf(IsNull(rsTemp!��Ʒ��), rsTemp!ͨ����, rsTemp!��Ʒ��)
                Else
                    .TextMatrix(.rows - 1, menuPriceCol.Ʒ��) = "[" & rsTemp!���� & "]" & rsTemp!ͨ����
                End If
                .TextMatrix(.rows - 1, menuPriceCol.���) = IIf(IsNull(rsTemp!���), "", rsTemp!���)
                .TextMatrix(.rows - 1, menuPriceCol.�Ƿ���) = rsTemp!�Ƿ���
                
'                If mintMethod = 1 Or mintMethod = 2 Then
'                    gstrSQL = "select min(����) as ���� from �ɱ��۵�����Ϣ where ���ۻ��ܺ�=[1] and ҩƷid=[2]"
'                    Set rs���� = zldatabase.OpenSQLRecord(gstrSQL, "���ز�ѯ", mstr���ۻ��ܺ�, rsTemp!ҩƷID)
'                    If rs����.RecordCount > 0 Then
'                        .TextMatrix(.rows - 1, menuPriceCol.����) = IIf(IsNull(rs����!����), "", rs����!����)
'                    End If
'                Else
                    .TextMatrix(.rows - 1, menuPriceCol.����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
'                End If
                
                .TextMatrix(.rows - 1, menuPriceCol.��λ) = strUnit
                .TextMatrix(.rows - 1, menuPriceCol.��װϵ��) = db��װϵ��

                .TextMatrix(.rows - 1, menuPriceCol.�ӳ���) = rsTemp!�ӳ���
                .TextMatrix(.rows - 1, menuPriceCol.���������) = Nvl(rsTemp!���������, 100)
                .TextMatrix(.rows - 1, menuPriceCol.�Ƿ��п��) = rsTemp!�Ƿ��п��
                .TextMatrix(.rows - 1, menuPriceCol.������ĿID) = IIf(IsNull(rsTemp!������ĿID), "", rsTemp!������ĿID)
                .TextMatrix(.rows - 1, menuPriceCol.ԭ�ɱ���) = GetFormat(Nvl(rsTemp!ԭ�ɱ���, 0) * db��װϵ��, mintCostDigit)
                .TextMatrix(.rows - 1, menuPriceCol.�ֳɱ���) = GetFormat(rsTemp!�³ɱ��� * db��װϵ��, mintCostDigit)
                .TextMatrix(.rows - 1, menuPriceCol.ԭ���ۼ�) = GetFormat(IIf(IsNull(rsTemp!ԭ��), rsTemp!�ּ�, rsTemp!ԭ��) * db��װϵ��, mintPriceDigit)
                .TextMatrix(.rows - 1, menuPriceCol.�����ۼ�) = GetFormat(rsTemp!�ּ� * db��װϵ��, mintPriceDigit)
                .TextMatrix(.rows - 1, menuPriceCol.ԭ�ɹ��޼�) = GetFormat(rsTemp!ָ������ * db��װϵ��, mintCostDigit)
                .TextMatrix(.rows - 1, menuPriceCol.�ֲɹ��޼�) = GetFormat(rsTemp!ָ������ * db��װϵ��, mintCostDigit)
                .TextMatrix(.rows - 1, menuPriceCol.ԭָ���ۼ�) = GetFormat(rsTemp!ָ���ۼ� * db��װϵ��, mintPriceDigit)
                .TextMatrix(.rows - 1, menuPriceCol.��ָ���ۼ�) = GetFormat(rsTemp!ָ���ۼ� * db��װϵ��, mintPriceDigit)

                txtValuer.Text = IIf(IsNull(rsTemp!������), "", rsTemp!������)
                txtSummary.Text = IIf(IsNull(rsTemp!����˵��), "", rsTemp!����˵��)
                If mintModal = 1 Then
                    Me.dtpRunDate.MinDate = CDate(rsTemp!ִ������)
                End If
                If IsNull(rsTemp!ִ������) Then
                    StrToday = Format(zlDataBase.Currentdate(), "yyyy-MM-dd hh:mm:ss")
                Else
                    StrToday = Format(rsTemp!ִ������, "yyyy-MM-dd hh:mm:ss")
                End If
                Me.dtpRunDate.Value = CDate(StrToday)

                .rows = .rows + 1
                Call setColEdit
                .RowHeight(.rows - 1) = mlngRowHeight
            End If
            rsTemp.MoveNext
        Next
        Call GetDrugStore(Val(.TextMatrix(1, menuPriceCol.ҩƷid)), 1)
    End With

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FindGridRow(ByVal strInput As String)
    Dim n As Integer
    Dim lngFindRow As Long
    Dim strҩ�� As String
    Dim lngRow As Long

    '����ҩƷ
    On Error GoTo errHandle
    If strInput <> txtFind.Tag Then
        '��ʾ�µĲ���
        txtFind.Tag = strInput

        gstrSQL = "Select Distinct A.Id,'[' || A.���� || ']' As ҩƷ����, A.���� As ͨ����, B.���� As ��Ʒ�� " & _
                  "From �շ���ĿĿ¼ A,�շ���Ŀ���� B " & _
                  "Where (A.վ�� = [3] Or A.վ�� is Null) And A.Id =B.�շ�ϸĿid And A.��� In ('5','6','7') " & _
                  "  And (A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2] ) " & _
                  "Order By ҩƷ���� "
        Set mrsFindName = zlDataBase.OpenSQLRecord(gstrSQL, "ȡƥ���ҩƷID", strInput & "%", "%" & strInput & "%", gstrNodeNo)

        If mrsFindName.RecordCount = 0 Then Exit Sub
        mrsFindName.MoveFirst
    End If

    '��ʼ����
    If mrsFindName.State <> adStateOpen Then Exit Sub
    If mrsFindName.RecordCount = 0 Then Exit Sub

    For n = 1 To mrsFindName.RecordCount
        '��������ˣ��򷵻ص�1����¼
        If mrsFindName.EOF Then mrsFindName.MoveFirst

        If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
            strҩ�� = mrsFindName!ҩƷ���� & mrsFindName!ͨ����
        Else
            strҩ�� = mrsFindName!ҩƷ���� & IIf(IsNull(mrsFindName!��Ʒ��), mrsFindName!ͨ����, mrsFindName!��Ʒ��)
        End If

        For lngRow = 1 To vsfPrice.rows - 1
            lngFindRow = vsfPrice.FindRow(strҩ��, lngRow, CLng(menuPriceCol.Ʒ��), True, True)
            If lngFindRow > 0 Then
                vsfPrice.Select lngFindRow, 1, lngFindRow, vsfPrice.Cols - 1
                vsfPrice.TopRow = lngFindRow
                Exit For
            End If
        Next

        If lngFindRow > 0 Then  '��ѯ�����ݺ���ƶ�����һ�����˳����β�ѯ
            mrsFindName.MoveNext
            Exit For
        Else
            mrsFindName.MoveNext 'δ��ѯ���������ƶ�����һ�����ݼ�������ѯ
        End If
    Next
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    If vsfPrice.Height + y <= 800 Then Exit Sub
    If TabCtlDetails.Height - y <= 1000 Then Exit Sub
    picSplit.Move 0, picSplit.Top + y
    vsfPrice.Move 0, fraCondition.Top + fraCondition.Height + 20, Me.ScaleWidth, vsfPrice.Height + y

    With TabCtlDetails
        .Top = picSplit.Top + picSplit.Height + 5
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = TabCtlDetails.Height - y
    End With
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Trim(txtFind.Text) = "" Then Exit Sub

    Call FindGridRow(UCase(Trim(txtFind.Text)))
End Sub

Private Sub txtSummary_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If LenB(StrConv(txtSummary.Text, vbFromUnicode)) >= 100 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtSummary_Validate(Cancel As Boolean)
    If LenB(StrConv(txtSummary.Text, vbFromUnicode)) > 100 Then
        MsgBox "˵��̫����", vbInformation, gstrSysName
        txtSummary.SelStart = 0
        txtSummary.SelLength = LenB(StrConv(txtSummary.Text, vbFromUnicode))
        Cancel = True
    End If
End Sub

Private Sub txt��Ӧ��_GotFocus()
    Me.txt��Ӧ��.SelStart = 0: Me.txt��Ӧ��.SelLength = Len(Me.txt��Ӧ��.Text)
End Sub

Private Sub txt��Ӧ��_KeyPress(KeyAscii As Integer)
    Dim strTmp As String
    Dim rsTemp As ADODB.Recordset

    On Error GoTo errHandle
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub

    strTmp = UCase(Trim(Me.txt��Ӧ��.Text))

    If strTmp = "" Then
        Me.txt��Ӧ��.Tag = "|"
        Exit Sub
    ElseIf strTmp = Split(Me.txt��Ӧ��.Tag, "|")(1) Then
        Exit Sub
    End If

    gstrSQL = "Select ����,����,����,id" & _
            " From ��Ӧ��" & _
            " where (���� Like [1] " & _
            "       Or ���� Like [2] " & _
            "       Or ���� Like [2])" & _
            " And ĩ��=1 And substr(����,1,1) = '1' And (����ʱ�� is null or ����ʱ��=to_date('3000-01-01','YYYY-MM-DD')) " & _
            " Order By ���� "
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, strTmp & "%", IIf(gstrMatchMethod = "0", "%", "") & strTmp & "%")

    With rsTemp
        If .EOF Then
            MsgBox "û���ҵ�ƥ��Ĺ�Ӧ�̣����ڹ�Ӧ�̹��������ӹ�Ӧ�̣�", vbInformation, gstrSysName
            Me.txt��Ӧ��.Text = Split(Me.txt��Ӧ��.Tag, "|")(1)
            Me.txt��Ӧ��.SelStart = 0: Me.txt��Ӧ��.SelLength = Len(Me.txt��Ӧ��.Text)
            Exit Sub
        End If

        If .RecordCount = 1 Then
            Me.txt��Ӧ��.Text = Trim(rsTemp!����): Me.txt��Ӧ��.Tag = rsTemp!id & "|" & rsTemp!����
            Exit Sub
        Else
            With Me.mshProvider
                .Left = Me.chk��Ӧ��.Left
                .Top = Me.txt��Ӧ��.Top + Me.txt��Ӧ��.Height
                .Clear
                Set .DataSource = rsTemp
                .ColWidth(0) = 800: .ColWidth(1) = 2500: .ColWidth(2) = 800: .ColWidth(3) = 0
                .Row = 1: .ColSel = .Cols - 1
                .ZOrder 0: .Visible = True: .SetFocus
            End With
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub get�ֶμӳ��ۼ�(ByVal lngҩƷID As Long, ByVal lng����ϵ�� As Long, ByVal dbl�ɹ��� As Double, ByRef dbl�ۼ� As Double)
'���ܣ�ͨ���ɱ��۰��ֶμӳɷ�ʽ�����ۼ�
'�������ɱ���,�ۼ�
    Dim dbl��۶� As Double
    Dim blnData As Boolean
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle

    mdbl�ֶμӳ��� = 0
    dbl��۶� = 0
    
    gstrSQL = "select ��� from  �շ���ĿĿ¼ a where a.id=[1]"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "ȡ��ҩƷ���ʷ���", lngҩƷID)
    If rsTemp!��� = 7 Then
        mrs�ֶμӳ�.Filter = "����=1"
    Else
        mrs�ֶμӳ�.Filter = "����=0"
    End If
    
    If mrs�ֶμӳ�.RecordCount <> 0 Then
        mrs�ֶμӳ�.MoveFirst
        Do While Not mrs�ֶμӳ�.EOF
            With mrs�ֶμӳ�
                If dbl�ɹ��� > !��ͼ� And dbl�ɹ��� <= !��߼� Then
                    mdbl�ֶμӳ��� = IIf(IsNull(!�ӳ���), 0, !�ӳ���) / 100
                    dbl��۶� = IIf(IsNull(!��۶�), 0, !��۶�)
                    blnData = True
                    Exit Do
                End If
            End With
            mrs�ֶμӳ�.MoveNext
        Loop
    End If
    
    If blnData = False Then
        MsgBox "û�����ý���Ϊ��" & dbl�ɹ��� & "  �ķֶμӳ����ݣ�����ҩƷĿ¼�����ֶμӳ��ʣ������ã�", vbInformation, gstrSysName
        dbl�ۼ� = 0
        Exit Sub
    End If
    
    dbl�ۼ� = dbl�ɹ��� * (1 + mdbl�ֶμӳ���) + dbl��۶�
    
    Set rsTemp = Nothing
    gstrSQL = "Select ָ�����ۼ� From ҩƷ��� Where ҩƷID=[1] "
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ȡָ�����ۼ�]", lngҩƷID)
    If rsTemp!ָ�����ۼ� * lng����ϵ�� < dbl�ۼ� Then
        dbl�ۼ� = rsTemp!ָ�����ۼ� * lng����ϵ��
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub txt��Ӧ��_Validate(Cancel As Boolean)
    If Me.txt��Ӧ��.Text = "" Then
        Me.txt��Ӧ��.Tag = "|"
    ElseIf Me.txt��Ӧ��.Text <> Split(Me.txt��Ӧ��.Tag, "|")(1) Then
        txt��Ӧ��_KeyPress (vbKeyReturn)
    End If
End Sub


Private Sub vsfPay_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsfPay
        .Move 0, 360, TabCtlDetails.Width, TabCtlDetails.Height - 370
    End With
End Sub

Private Sub vsfPay_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfPay
        If .Cell(flexcpBackColor, Row, Col, Row, Col) = mconlngColor Then
            Cancel = True
            .Editable = flexEDNone
        Else
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub


Private Sub vsfPay_DblClick()
    With vsfPay
        If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mconlngCanColColor Then
            .EditCell
            .EditSelStart = 0
            .EditSelLength = Len(.EditText)
        End If
    End With
End Sub

Private Sub vsfPay_EnterCell()
    With vsfPay
        If .CellBackColor = mconlngColor Then
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
        End If
    End With
End Sub

Private Sub vsfPay_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsfPay
        If KeyCode = vbKeyReturn Then
            If .Col = menuPayCol.Ʒ�� Then
                .Col = menuPayCol.��Ʊ��
            ElseIf .Col = menuPayCol.��Ʊ�� Then
                .Col = menuPayCol.��Ʊ����
            ElseIf .Col = menuPayCol.��Ʊ���� Then
                .Col = menuPayCol.��Ʊ���
            ElseIf .Col = menuPayCol.��Ʊ��� And .Row <> .rows - 1 Then
                .Col = menuPayCol.Ʒ��
                .Row = .Row + 1
            End If
        End If
    End With
End Sub

Private Sub vsfPay_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        With vsfPay
            If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mconlngCanColColor Then
                .Editable = flexEDKbdMouse
            Else
                .Editable = flexEDNone
            End If
        End With
    End If
End Sub

Private Sub vsfPay_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strkey As String
    Dim intDigit As Integer
    
    If KeyAscii = vbKeyReturn Then Exit Sub
    If KeyAscii <> vbKeyBack Then
        With vsfPay
            If Col = menuPayCol.��Ʊ��� Then
                strkey = .EditText
                intDigit = mintMoneyDigit
                If KeyAscii = vbKeyDelete Then
                    If InStr(1, .EditText, ".") > 0 Then
                        KeyAscii = 0
                    End If
                ElseIf KeyAscii = Asc(".") Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
                    If .EditSelLength = Len(strkey) Then Exit Sub
                    If InStr(strkey, ".") <> 0 And Chr(KeyAscii) = "." Then   'ֻ�ܴ���һ��С����
                        KeyAscii = 0
                        Exit Sub
                    End If
                    If Len(Mid(strkey, InStr(1, strkey, ".") + 1)) >= intDigit And strkey Like "*.*" Then
                        KeyAscii = 0
                        Exit Sub
                    Else
                        Exit Sub
                    End If
                Else
                    KeyAscii = 0
                End If
            ElseIf Col = menuPayCol.��Ʊ�� Then
                If InStr("`~!@#$%^&*()_-+={[}]|\:;""'<,>.?/", Chr(KeyAscii)) > 0 Then
                    KeyAscii = 0
                End If
            End If
        End With
    End If
End Sub

Private Sub vsfPay_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strkey As String

    With vsfPay
        If Col = menuPayCol.��Ʊ���� Then
            strkey = .EditText
            If strkey <> "" Then
                If Len(strkey) = 8 And InStr(1, strkey, "-") = 0 Then
                    strkey = TranNumToDate(strkey)
                    If strkey = "" Then
                        MsgBox "�Բ��𣬷�Ʊ���ڱ���Ϊ������,��ʽ(20000101����2000-01-01)��", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                    .EditText = strkey
                    .TextMatrix(Row, menuPayCol.��Ʊ����) = .EditText
                End If
                
                If Not IsDate(strkey) Then
                    MsgBox "�Բ��𣬷�Ʊ���ڱ���Ϊ������(20000101����2000-01-01)��", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
            End If
        End If
    End With
End Sub

Private Sub vsfprice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsfPrice
        If Col = menuPriceCol.�ֳɱ��� Then
            If Val(.TextMatrix(Row, Col)) <> Val(.TextMatrix(Row, menuPriceCol.ԭ�ɱ���)) Then
                .Cell(flexcpFontBold, Row, Col, Row, Col) = 10
                .Cell(flexcpForeColor, Row, Col, Row, Col) = vbRed
            End If
        ElseIf Col = menuPriceCol.�����ۼ� Then
            If Val(.TextMatrix(Row, Col)) <> Val(.TextMatrix(Row, menuPriceCol.ԭ���ۼ�)) Then
                .Cell(flexcpFontBold, Row, Col, Row, Col) = 10
                .Cell(flexcpForeColor, Row, Col, Row, Col) = vbRed
            End If
        End If
    End With
End Sub

Private Sub vsfPrice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow = NewRow Then Exit Sub
'    Call SetRowHidden(Val(vsfPrice.TextMatrix(NewRow, menuPriceCol.ҩƷid)))
End Sub

Private Sub SetRowHidden(ByVal lngDrugID As Long)
    '���ܣ��е���ʾ������
    '������ҩƷid
    Dim intRow As Integer

    If lngDrugID = 0 Then Exit Sub
    With vsfStore
        For intRow = 1 To .rows - 1
            If Val(.TextMatrix(intRow, menuStoreCol.ҩƷid)) = lngDrugID Then
                .RowHidden(intRow) = False
            Else
                .RowHidden(intRow) = True
            End If
        Next
    End With

    With vsfPay
        For intRow = 1 To .rows - 1
            If Val(.TextMatrix(intRow, menuPayCol.ҩƷid)) = lngDrugID Then
                .RowHidden(intRow) = False
            Else
                .RowHidden(intRow) = True
            End If
        Next
    End With
End Sub

'Private Sub vsfPrice_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'    With vsfPrice
'        If .Cell(flexcpBackColor, Row, Col, Row, Col) = mconlngColor Then
'            Cancel = True
'            .Editable = flexEDNone
'        Else
'            .Editable = flexEDKbdMouse
'        End If
'    End With
'End Sub

Private Sub vsfPrice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim mrsReturn As Recordset
    Dim vRect As RECT
    Dim dblLeft As Double
    Dim dblTop As Double

    mBlnClick = True
    vRect = GetControlRect(vsfPrice.hWnd) '��ȡλ��
    dblLeft = vsfPrice.CellLeft
    dblTop = vRect.Top + vsfPrice.CellTop + vsfPrice.CellHeight


    On Error GoTo errHandle
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(1, "", 0, , , , , , , , , True)
    End If
    Set mrsReturn = frmSelector.ShowME(Me, 0, 1, , dblLeft, dblTop, , , , , , , , , False, mstrPrivs)

    If mrsReturn.RecordCount = 0 Then Exit Sub
    mblnUpdateAdd = True
    Call GetDrugPirce(mrsReturn, Row)
    mblnUpdateAdd = False
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetDrugPirce(ByVal rsReturn As ADODB.Recordset, ByVal Row As Integer)
    '������ȡҩƷ��Ϣ
    Dim rsTemp As Recordset
    Dim lngDrugID As Long
    Dim lngRow As Long
    Dim i As Long
    Dim intCurrentPrice As Integer '�Ƿ���ʱ��
    Dim strUnit As String
    Dim db��װϵ�� As Double
    Dim strInfo As String

    On Error GoTo errHandle

    mlngOldDrugID = Val(vsfPrice.TextMatrix(Row, menuPriceCol.ҩƷid))
    Set rsReturn = CheckDoubleDrug(rsReturn)
    If rsReturn.RecordCount = 0 Then Exit Sub

    rsReturn.MoveFirst
    For i = 0 To rsReturn.RecordCount - 1
        With vsfPrice
            lngDrugID = rsReturn!ҩƷid

            '����Ƿ����Ϊִ�еļ۸�
            If checkNotExecutePrice(lngDrugID, strInfo) = True Then
                MsgBox strInfo, vbInformation, gstrSysName
                Exit Sub
            End If

            Select Case mintUnit
                Case 0
                    db��װϵ�� = rsReturn!ҩ���װ
                    strUnit = rsReturn!ҩ�ⵥλ
                Case 1
                    db��װϵ�� = rsReturn!סԺ��װ
                    strUnit = rsReturn!סԺ��λ
                Case 2
                    db��װϵ�� = rsReturn!�����װ
                    strUnit = rsReturn!���ﵥλ
                Case 3
                    db��װϵ�� = 1
                    strUnit = rsReturn!�ۼ۵�λ
            End Select

            .TextMatrix(Row, menuPriceCol.ҩƷid) = lngDrugID

            If gintҩƷ������ʾ = 1 Then
                .TextMatrix(Row, menuPriceCol.Ʒ��) = "[" & rsReturn!ҩƷ���� & "]" & IIf(IsNull(rsReturn!��Ʒ��), rsReturn!ͨ����, rsReturn!��Ʒ��)
            Else
                .TextMatrix(Row, menuPriceCol.Ʒ��) = "[" & rsReturn!ҩƷ���� & "]" & rsReturn!ͨ����
            End If

            .TextMatrix(Row, menuPriceCol.���) = IIf(IsNull(rsReturn!���), "", rsReturn!���)
            .TextMatrix(Row, menuPriceCol.�Ƿ���) = rsReturn!ʱ��
            intCurrentPrice = rsReturn!ʱ��
            .TextMatrix(Row, menuPriceCol.����) = IIf(IsNull(rsReturn!����), "", rsReturn!����)
            .TextMatrix(Row, menuPriceCol.��λ) = strUnit
            .TextMatrix(Row, menuPriceCol.��װϵ��) = db��װϵ��
            gstrSQL = "select ҩƷid from ҩƷ��� where ҩƷid=[1] and ����=1 "
            Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "�����", lngDrugID)
            If rsTemp.RecordCount = 0 Then
                .TextMatrix(Row, menuPriceCol.�Ƿ��п��) = 0
            Else
                .TextMatrix(Row, menuPriceCol.�Ƿ��п��) = 1
            End If

            If intCurrentPrice = 0 Then '����ҩƷ
                '��ʾ����ҩƷ���ۣ��ɱ���ȡƽ���۸��ۼ�ȡ�շѼ�Ŀ�ּ�
                gstrSQL = "Select b.Id, Decode(Nvl(k.�������, 0), 0, a.�ɱ���, (k.����� - k.�����) / k.�������) As �ɱ���, a.ָ��������, a.ָ�����ۼ�, b.�ּ�, a.���������," & vbNewLine & _
                            "       nvl(a.�ӳ���,0) / 100 As �ӳ���, b.������Ŀid" & vbNewLine & _
                            "From ҩƷ��� A, �շѼ�Ŀ B," & vbNewLine & _
                            "     (Select Sum(ʵ�ʽ��) �����, Sum(ʵ�ʲ��) As �����, Sum(ʵ������) �������" & vbNewLine & _
                            "       From ҩƷ���" & vbNewLine & _
                            "       Where ���� = 1 And ҩƷid = [1] ) K" & vbNewLine & _
                            "Where a.ҩƷid = b.�շ�ϸĿid And a.ҩƷid = [1] And Sysdate Between ִ������ And ��ֹ����"
            Else 'ʱ��ҩƷ
                '��ʾʱ��ҩƷ���ۣ�ȡ�����/���������Ϊ��۸�
                gstrSQL = "select P.id,Decode(Nvl(K.�������,0),0,P.�ּ�,K.�����/Nvl(K.�������,1)) �ּ�,nvl(j.�ӳ���,0) / 100 as �ӳ���,decode(nvl(k.�������,0),0,j.�ɱ���,(k.�����-k.�����)/k.�������) as �ɱ���,j.ָ��������,j.ָ�����ۼ�,j.���������,p.������Ŀid,P.ִ������,P.������Ŀid,I.���� as ��������" & _
                        " from �շѼ�Ŀ P,������Ŀ I,ҩƷ��� J," & _
                        "   (Select Sum(ʵ�ʽ��) �����,Sum(ʵ�ʲ��) as �����,Sum(ʵ������) �������" & _
                        "    From ҩƷ��� Where ����=1 and ҩƷID=[1] ) K" & _
                        " where P.������Ŀid=I.id and p.�շ�ϸĿid=j.ҩƷid and P.�շ�ϸĿid=[1] " & _
                        "       and (P.��ֹ���� is null or SYSDATE BETWEEN P.ִ������ AND P.��ֹ����)"
            End If
            Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "��ѯҩƷ", lngDrugID)
            If rsTemp.RecordCount = 0 Then
                MsgBox "��ҩƷ�����ڣ������½�����ҩƷ��Ƭ��", vbInformation, gstrSysName
                Exit Sub
            End If
            .TextMatrix(Row, menuPriceCol.ԭ��id) = rsTemp!id
            .TextMatrix(Row, menuPriceCol.������ĿID) = IIf(IsNull(rsTemp!������ĿID), 0, rsTemp!������ĿID)
            .TextMatrix(Row, menuPriceCol.�ӳ���) = GetFormat(IIf(IsNull(rsTemp!�ӳ���), 0, rsTemp!�ӳ���), 5)
            .TextMatrix(Row, menuPriceCol.���������) = IIf(IsNull(rsTemp!���������), 100, rsTemp!���������)
            .TextMatrix(Row, menuPriceCol.ԭ�ɱ���) = GetFormat(IIf(IsNull(rsTemp!�ɱ���), 0, rsTemp!�ɱ���) * db��װϵ��, mintCostDigit)
            .TextMatrix(Row, menuPriceCol.�ֳɱ���) = GetFormat(IIf(IsNull(rsTemp!�ɱ���), 0, rsTemp!�ɱ���) * db��װϵ��, mintCostDigit)
            .TextMatrix(Row, menuPriceCol.ԭ���ۼ�) = GetFormat(IIf(IsNull(rsTemp!�ּ�), 0, rsTemp!�ּ�) * db��װϵ��, mintPriceDigit)
            .TextMatrix(Row, menuPriceCol.�����ۼ�) = GetFormat(IIf(IsNull(rsTemp!�ּ�), 0, rsTemp!�ּ�) * db��װϵ��, mintPriceDigit)
            .TextMatrix(Row, menuPriceCol.ԭ�ɹ��޼�) = GetFormat(IIf(IsNull(rsTemp!ָ��������), 0, rsTemp!ָ��������) * db��װϵ��, mintCostDigit)
            .TextMatrix(Row, menuPriceCol.�ֲɹ��޼�) = .TextMatrix(Row, menuPriceCol.ԭ�ɹ��޼�)
            .TextMatrix(Row, menuPriceCol.ԭָ���ۼ�) = GetFormat(IIf(IsNull(rsTemp!ָ�����ۼ�), 0, rsTemp!ָ�����ۼ�) * db��װϵ��, mintPriceDigit)
            .TextMatrix(Row, menuPriceCol.��ָ���ۼ�) = .TextMatrix(Row, menuPriceCol.ԭָ���ۼ�)

            Call GetDrugStore(lngDrugID, Row)
            If Row = .rows - 1 Then '���һ�в�������
                .rows = .rows + 1
                .RowHeight(.rows - 1) = mlngRowHeight
                Row = Row + 1
            End If
        End With
'        If mint���� = 0 And mblnʱ��ҩƷ�����ε��� = True Then '�ۼ۵���
'            Call GetDrugStore(lngDrugID, db��װϵ��)
'        ElseIf mint���� <> 0 Then

'        End If
'        Call SetRowHidden(lngDrugID)

        rsReturn.MoveNext
    Next
    Call setColEdit

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetDrugStore(ByVal lngDrugID As Long, ByVal intRow As Integer)
    Dim rsTemp As ADODB.Recordset
    Dim dblOldCost As Double
    Dim dblOldPrice As Double
    Dim dblNewCost As Double
    Dim dblNewPrice As Double
    Dim dbl�ӳ��� As Double
    Dim lngCurRow As Long     '��ǰ��
    Dim i As Long
    Dim dbl��Ʊ��� As Double
    Dim strҩƷ���� As String
    Dim str��Ʊ As String
    Dim str��Ʊ���� As String
    Dim rsPirce As ADODB.Recordset
    Dim rsCost As ADODB.Recordset
    Dim dbl��װ���� As Double
    Dim bln��ͬҩƷ As Boolean
    Dim lngҩƷID As Long
    Dim str��λ As String


    '���ܣ�Ϊ����б��������
    '������ҩƷid

    On Error GoTo errHandle
    '�ȼ���Ƿ����ظ������ݣ�����о���������ظ�������
    With vsfStore
        For i = .rows - 1 To 1 Step -1
            If Val(.TextMatrix(i, menuStoreCol.ҩƷid)) = mlngOldDrugID And mlngOldDrugID <> 0 Then
                .RemoveItem i
            End If
        Next
    End With

    With vsfPay
        For i = .rows - 1 To 1 Step -1
            If Val(.TextMatrix(i, menuPayCol.ҩƷid)) = mlngOldDrugID And mlngOldDrugID <> 0 Then
                .RemoveItem i
            End If
        Next
    End With

    If mintModal = 0 Or mblnUpdateAdd = True Or mblnBatchItem = True Then
        gstrSQL = "Select s.�ⷿid,s.ҩƷid, d.���� As �ⷿ, '[' || m.���� || ']' || m.���� As ҩƷ, m.���, m.����, m.���㵥λ �ۼ۵�λ," & vbNewLine & _
                        "       p.ҩ�ⵥλ, s.�ϴ����� As ����, nvl(s.ʵ������,0) As ����,s.����, Nvl(m.�Ƿ���, 0) ���, m.Id," & vbNewLine & _
                        "       Decode(Nvl(m.�Ƿ���, 0), 0, e.�ּ�, Decode(Nvl(s.���ۼ�, 0),0,Decode(Nvl(s.ʵ������, 0),0,e.�ּ�, s.ʵ�ʽ��/s.ʵ������),s.���ۼ�)) As ʱ���ۼ�," & vbNewLine & _
                        "       p.ָ������� As �����,nvl(p.�ӳ���,0) as �ӳ��� ,Decode(Nvl(s.ƽ���ɱ���, 0), 0, p.�ɱ���, s.ƽ���ɱ���) As �ɱ���, s.�ϴι�Ӧ��id, n.���� As ��Ӧ��, s.Ч��, s.�ϴβ��� As ����" & vbNewLine & _
                        "From ҩƷ��� S, ���ű� D, �շ���ĿĿ¼ M, ҩƷ��� P, ��Ӧ�� N, �շѼ�Ŀ E" & vbNewLine & _
                        "Where d.Id = s.�ⷿid And s.ҩƷid = m.Id And m.Id = p.ҩƷid And Nvl(s.�ϴι�Ӧ��id, 0) = n.Id(+) And m.Id = e.�շ�ϸĿid And" & vbNewLine & _
                        "      s.���� = 1 And s.ҩƷid = [1] And Sysdate Between e.ִ������ And e.��ֹ���� " & vbNewLine & _
                        "Order By �ⷿ, s.�ϴ�����"

        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lngDrugID)

        If mlng��Ӧ��ID > 0 Then
            rsTemp.Filter = "�ϴι�Ӧ��ID=" & mlng��Ӧ��ID
        End If
    Else '�޸ģ�����
        If mintModal = 2 Then   '����
            If cboPriceMethod.Text = "�����ɱ���" Or cboPriceMethod.Text = "�ۼ۳ɱ���һ�����" Then
                gstrSQL = "select (sysdate-ִ������ ) as �Ƿ�ִ�� from ���ۻ��ܼ�¼ where ���ۺ�=[1]"
                Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "�Ƿ�ִ��", txtNO.Text)
                If rsTemp!�Ƿ�ִ�� > 0 Then
                    gstrSQL = "Select Distinct a.�ⷿid, c.���� As �ⷿ, b.ҩƷid, b.��ҩ��λid As �ϴι�Ӧ��id, '[' || e.���� || ']' || e.���� As ҩƷ, e.���, d.���� As ��Ӧ��," & vbNewLine & _
                            "                b.�³ɱ���, b.ԭ�ɱ���, b.��Ʊ��, b.��Ʊ����, b.��Ʊ���, b.����, b.����, b.����, e.�Ƿ��� As ���, e.���㵥λ As �ۼ۵�λ, f.ҩ�ⵥλ," & vbNewLine & _
                            "                nvl(a.��д����,0) As ����, f.ָ������� As �����, nvl(f.�ӳ���,0) as �ӳ��� ,b.Ч��" & vbNewLine & _
                            "From ҩƷ�շ���¼ A,�ɱ��۵�����Ϣ B, ���ű� C, ��Ӧ�� D, �շ���ĿĿ¼ E, ҩƷ��� F" & vbNewLine & _
                            "Where a.id=b.�շ�id And a.�ⷿid = c.Id And b.��ҩ��λid = d.Id(+) And" & vbNewLine & _
                            "      a.ҩƷid = e.Id And e.Id = f.ҩƷid And b.���ۻ��ܺ� = [1] and a.���� = 5"
                Else
                    gstrSQL = "Select Distinct a.�ⷿid,c.���� as �ⷿ, b.ҩƷid,a.�ϴι�Ӧ��id, '[' || e.���� || ']' ||e.���� as ҩƷ,e.���,d.���� as ��Ӧ��, b.�³ɱ���, b.ԭ�ɱ���, b.��Ʊ��, b.��Ʊ����, b.��Ʊ���" & _
                            " ,a.�ϴβ��� as ����,a.����,a.�ϴ����� as ����,e.�Ƿ��� as ���,e.���㵥λ as �ۼ۵�λ,f.ҩ�ⵥλ,nvl(a.ʵ������,0) as ����,f.ָ������� as �����,nvl(f.�ӳ���,0) as �ӳ��� ,a.Ч��" & _
                            " From ҩƷ��� A,���ű� C,��Ӧ�� D,�շ���ĿĿ¼ E,ҩƷ��� F," & _
                                 " (Select Distinct ҩƷid, �ⷿid, ����, ����, Ч��, ����, ԭ�ɱ���, �³ɱ���, ��Ʊ��, ��Ʊ����, ��Ʊ���, Ӧ����䶯, ִ������" & _
                                   " From �ɱ��۵�����Ϣ" & _
                                   " Where ���ۻ��ܺ� = [1]) B" & _
                            " Where a.ҩƷid = b.ҩƷid And Decode(b.�ⷿid, Null, 1, a.�ⷿid) = Decode(b.�ⷿid, Null, 1, b.�ⷿid) " & _
                            " and Decode(b.�ⷿid, Null, 1, Nvl(a.����, 0)) = Decode(b.�ⷿid, Null, 1, Nvl(b.����, 0)) " & _
                            " and a.�ⷿid=c.id and a.�ϴι�Ӧ��id=d.id(+) and a.ҩƷid=e.id and e.id=f.ҩƷid and a.����=1 "
                End If
            ElseIf cboPriceMethod.Text = "�����ۼ�" Then
                gstrSQL = "select (sysdate-ִ������ ) as �Ƿ�ִ�� from ���ۻ��ܼ�¼ where ���ۺ�=[1]"
                Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "�Ƿ�ִ��", txtNO.Text)
                If rsTemp!�Ƿ�ִ�� > 0 Then
                    gstrSQL = "Select Distinct a.�ⷿid, c.���� As �ⷿ, b.�շ�ϸĿid As ҩƷid, a.��ҩ��λid As �ϴι�Ӧ��id, '[' || e.���� || ']' || e.���� As ҩƷ, e.���," & vbNewLine & _
                            "                d.���� As ��Ӧ��, f.�ɱ��� As �³ɱ���, f.�ɱ��� As ԭ�ɱ���, '' ��Ʊ��, '' ��Ʊ����, '' ��Ʊ���, a.����, a.����, a.����, e.�Ƿ��� As ���," & vbNewLine & _
                            "                e.���㵥λ As �ۼ۵�λ, f.ҩ�ⵥλ, nvl(a.��д����,0) As ����, f.ָ������� As �����, nvl(f.�ӳ���,0) as �ӳ��� ,a.Ч��" & vbNewLine & _
                            "From ҩƷ�շ���¼ A, �շѼ�Ŀ B, ���ű� C, ��Ӧ�� D, �շ���ĿĿ¼ E, ҩƷ��� F" & vbNewLine & _
                            "Where a.�۸�id = b.Id And a.�ⷿid = c.Id And a.��ҩ��λid = d.Id(+) And a.ҩƷid = e.Id And e.Id = f.ҩƷid And" & vbNewLine & _
                            "      b.���ۻ��ܺ� = [1] and a.����=13"
                Else
                    gstrSQL = "Select Distinct a.�ⷿid, c.���� As �ⷿ, b.�շ�ϸĿid As ҩƷid, a.�ϴι�Ӧ��id, '[' || e.���� || ']' || e.���� As ҩƷ, e.���, d.���� As ��Ӧ��," & _
                                            " nvl(a.ƽ���ɱ���,f.�ɱ���) As �³ɱ���, nvl(a.ƽ���ɱ���,f.�ɱ���) As ԭ�ɱ���, '' ��Ʊ��, '' ��Ʊ����, '' ��Ʊ���, a.�ϴβ��� As ����, a.����, a.�ϴ����� As ����," & _
                                            " e.�Ƿ��� As ���, e.���㵥λ As �ۼ۵�λ, f.ҩ�ⵥλ, nvl(a.ʵ������,0) As ����, f.ָ������� As �����, nvl(f.�ӳ���,0) as �ӳ��� ,a.Ч��" & _
                            " From ҩƷ��� A, �շѼ�Ŀ B, ���ű� C, ��Ӧ�� D, �շ���ĿĿ¼ E, ҩƷ��� F" & _
                            " Where a.ҩƷid = b.�շ�ϸĿid And a.�ⷿid = c.Id And a.�ϴι�Ӧ��id = d.Id(+) And a.ҩƷid = e.Id And e.Id = f.ҩƷid And a.���� = 1  And" & _
                                  " b.���ۻ��ܺ� = [1]"
                End If
            End If
        Else '�޸�
            If cboPriceMethod.Text = "�����ɱ���" Or cboPriceMethod.Text = "�ۼ۳ɱ���һ�����" Then
                gstrSQL = "Select Distinct a.�ⷿid,c.���� as �ⷿ, b.ҩƷid,a.�ϴι�Ӧ��id, '[' || e.���� || ']' ||e.���� as ҩƷ,e.���,d.���� as ��Ӧ��, b.�³ɱ���, b.ԭ�ɱ���, b.��Ʊ��, b.��Ʊ����, b.��Ʊ���" & _
                            " ,a.�ϴβ��� as ����,a.����,a.�ϴ����� as ����,e.�Ƿ��� as ���,e.���㵥λ as �ۼ۵�λ,f.ҩ�ⵥλ,nvl(a.ʵ������,0) as ����,f.ָ������� as �����,nvl(f.�ӳ���,0) as �ӳ��� ,a.Ч��" & _
                            " From ҩƷ��� A,���ű� C,��Ӧ�� D,�շ���ĿĿ¼ E,ҩƷ��� F," & _
                                 " (Select Distinct ҩƷid, �ⷿid, ����, ����, Ч��, ����, ԭ�ɱ���, �³ɱ���, ��Ʊ��, ��Ʊ����, ��Ʊ���, Ӧ����䶯, ִ������" & _
                                   " From �ɱ��۵�����Ϣ" & _
                                   " Where ���ۻ��ܺ� = [1]) B" & _
                            " Where a.ҩƷid = b.ҩƷid And Decode(b.�ⷿid, Null, 1, a.�ⷿid) = Decode(b.�ⷿid, Null, 1, b.�ⷿid) " & _
                            " and Decode(b.�ⷿid, Null, 1, Nvl(a.����, 0)) = Decode(b.�ⷿid, Null, 1, Nvl(b.����, 0)) " & _
                            " and a.�ⷿid=c.id and a.�ϴι�Ӧ��id=d.id(+) and a.ҩƷid=e.id and e.id=f.ҩƷid and a.����=1  "
            ElseIf cboPriceMethod.Text = "�����ۼ�" Then
                gstrSQL = "Select Distinct a.�ⷿid, c.���� As �ⷿ, b.�շ�ϸĿid As ҩƷid, a.�ϴι�Ӧ��id, '[' || e.���� || ']' || e.���� As ҩƷ, e.���, d.���� As ��Ӧ��," & _
                                            " nvl(a.ƽ���ɱ���,f.�ɱ���) As �³ɱ���, nvl(a.ƽ���ɱ���,f.�ɱ���) As ԭ�ɱ���, '' ��Ʊ��, '' ��Ʊ����, '' ��Ʊ���, a.�ϴβ��� As ����, a.����, a.�ϴ����� As ����," & _
                                            " e.�Ƿ��� As ���, e.���㵥λ As �ۼ۵�λ, f.ҩ�ⵥλ, nvl(a.ʵ������,0) As ����, f.ָ������� As �����, nvl(f.�ӳ���,0) as �ӳ��� ,a.Ч��" & _
                            " From ҩƷ��� A, �շѼ�Ŀ B, ���ű� C, ��Ӧ�� D, �շ���ĿĿ¼ E, ҩƷ��� F" & _
                            " Where a.ҩƷid = b.�շ�ϸĿid And a.�ⷿid = c.Id And a.�ϴι�Ӧ��id = d.Id(+) And a.ҩƷid = e.Id And e.Id = f.ҩƷid And a.���� = 1  " & _
                                  "And b.���ۻ��ܺ� = [1]"
            End If
        End If
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, txtNO.Text)
    End If
    
    With vsfStore
        Do While Not rsTemp.EOF
            dbl��װ���� = 0
            dbl��Ʊ��� = 0
            dblOldPrice = 0
            dblNewPrice = 0
            str��λ = ""
            For i = 0 To vsfPrice.rows - 1
                If rsTemp!ҩƷid = vsfPrice.TextMatrix(i, menuPriceCol.ҩƷid) Then
                    dbl��װ���� = vsfPrice.TextMatrix(i, menuPriceCol.��װϵ��)
                    dblOldPrice = Val(vsfPrice.TextMatrix(i, menuPriceCol.ԭ���ۼ�))
                    dblNewPrice = Val(vsfPrice.TextMatrix(i, menuPriceCol.�����ۼ�))
                    str��λ = vsfPrice.TextMatrix(i, menuPriceCol.��λ)
                    Exit For
                End If
            Next
            .rows = .rows + 1
            Call setColEdit
            .RowHeight(.rows - 1) = mlngRowHeight

            '�ӿհ��п�ʼ��������
            .TextMatrix(.rows - 1, menuStoreCol.ҩƷid) = rsTemp!ҩƷid
            .TextMatrix(.rows - 1, menuStoreCol.�ⷿ) = rsTemp!�ⷿ
            .TextMatrix(.rows - 1, menuStoreCol.�ⷿid) = rsTemp!�ⷿid
            .TextMatrix(.rows - 1, menuStoreCol.��Ӧ��) = Nvl(rsTemp!��Ӧ��, "")
            .TextMatrix(.rows - 1, menuStoreCol.��Ӧ��id) = IIf(mlng��Ӧ��ID > 0, mlng��Ӧ��ID, Nvl(rsTemp!�ϴι�Ӧ��ID))
            .TextMatrix(.rows - 1, menuStoreCol.ҩƷ) = rsTemp!ҩƷ
            strҩƷ���� = rsTemp!ҩƷ

            .TextMatrix(.rows - 1, menuStoreCol.���) = IIf(IsNull(rsTemp!���), "", rsTemp!���)
            .TextMatrix(.rows - 1, menuStoreCol.��λ) = str��λ
            .TextMatrix(.rows - 1, menuStoreCol.����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
            .TextMatrix(.rows - 1, menuStoreCol.Ч��) = Format(IIf(IsNull(rsTemp!Ч��), "", rsTemp!Ч��), "YYYY-MM-DD")
            .TextMatrix(.rows - 1, menuStoreCol.����) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
            .TextMatrix(.rows - 1, menuStoreCol.����) = GetFormat(rsTemp!���� / dbl��װ����, mintNumberDigit)
            .TextMatrix(.rows - 1, menuStoreCol.��װϵ��) = dbl��װ����
            .TextMatrix(.rows - 1, menuStoreCol.����) = Nvl(rsTemp!����, 0)
            .TextMatrix(.rows - 1, menuStoreCol.���) = rsTemp!���


            If mintModal = 0 Or mblnUpdateAdd = True Or mblnBatchItem = True Then
                dblOldCost = IIf(IsNull(rsTemp!�ɱ���), 0, rsTemp!�ɱ���) * dbl��װ����
                
                If mdbl�ӳ��� > 0 Then
                    dbl�ӳ��� = Round(mdbl�ӳ��� / 100, 7)
                ElseIf dblOldCost > 0 Then
                    dbl�ӳ��� = Round(IIf(rsTemp!��� = 1, rsTemp!ʱ���ۼ� * dbl��װ����, dblOldPrice) / dblOldCost - 1, 7)
                Else
                    dbl�ӳ��� = Nvl(rsTemp!�ӳ���, 0) / 100
                End If
                If 1 + dbl�ӳ��� = 0 Then
                    dblNewCost = 0
                Else
                    dblNewCost = rsTemp!ʱ���ۼ� * dbl��װ���� / (1 + dbl�ӳ���)
                End If
                If dbl�ӳ��� = -1 Then dbl�ӳ��� = 0

                .TextMatrix(.rows - 1, menuStoreCol.ԭ���ۼ�) = GetFormat(IIf(rsTemp!��� = 1, rsTemp!ʱ���ۼ� * dbl��װ����, dblOldPrice), mintPriceDigit)
                .TextMatrix(.rows - 1, menuStoreCol.�����ۼ�) = GetFormat(IIf(rsTemp!��� = 1, rsTemp!ʱ���ۼ� * dbl��װ����, dblOldPrice), mintPriceDigit)
                .TextMatrix(.rows - 1, menuStoreCol.�������) = Format(rsTemp!���� / dbl��װ���� * (Val(.TextMatrix(.rows - 1, menuStoreCol.�����ۼ�)) - Val(.TextMatrix(.rows - 1, menuStoreCol.ԭ���ۼ�))), mstrMoneyFormat)
                .TextMatrix(.rows - 1, menuStoreCol.�ӳ���) = GetFormat(GetFormat(dbl�ӳ���, 5) * 100, 5)
                .TextMatrix(.rows - 1, menuStoreCol.ԭ�ɹ���) = GetFormat(dblOldCost, mintCostDigit)
                .TextMatrix(.rows - 1, menuStoreCol.�ֲɹ���) = GetFormat(dblNewCost, mintCostDigit)
                .TextMatrix(.rows - 1, menuStoreCol.��۲�) = Format((Val(.TextMatrix(.rows - 1, menuStoreCol.�ֲɹ���)) - Val(.TextMatrix(.rows - 1, menuStoreCol.ԭ�ɹ���))) * Val(.TextMatrix(.rows - 1, menuStoreCol.����)), mstrMoneyFormat)
                dbl��Ʊ��� = dbl��Ʊ��� + (dblNewCost - dblOldCost) * Val(.TextMatrix(.rows - 1, menuStoreCol.����))
                
                'ΪӦ����¼��ֵ
                If mint���� = 1 Or mint���� = 2 Then
                    If vsfPay.rows > 1 Then
                        bln��ͬҩƷ = False
                        For i = 1 To vsfPay.rows - 1
                            If vsfPay.TextMatrix(i, menuPayCol.ҩƷid) = rsTemp!ҩƷid Then
                                bln��ͬҩƷ = True
                                Exit For
                            End If
                        Next
                        If bln��ͬҩƷ = True Then
                            vsfPay.TextMatrix(i, menuPayCol.��Ʊ���) = GetFormat(Val(vsfPay.TextMatrix(i, menuPayCol.��Ʊ���)) + dbl��Ʊ���, mintMoneyDigit)
                        Else
                            vsfPay.rows = vsfPay.rows + 1
                            vsfPay.RowHeight(vsfPay.rows - 1) = mlngRowHeight
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.ҩƷid) = rsTemp!ҩƷid
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.Ʒ��) = strҩƷ����
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.��Ʊ��) = str��Ʊ
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.��Ʊ����) = Format(str��Ʊ����, "yyyy-mm-dd")
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.��Ʊ���) = GetFormat(dbl��Ʊ���, mintMoneyDigit)
                        End If
                    Else
                        vsfPay.rows = vsfPay.rows + 1
                        vsfPay.RowHeight(vsfPay.rows - 1) = mlngRowHeight
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.ҩƷid) = rsTemp!ҩƷid
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.Ʒ��) = strҩƷ����
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.��Ʊ��) = str��Ʊ
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.��Ʊ����) = Format(str��Ʊ����, "yyyy-mm-dd")
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.��Ʊ���) = GetFormat(dbl��Ʊ���, mintMoneyDigit)
                    End If
                End If
            Else
                If mintModal = 2 And (cboPriceMethod.Text = "�����ۼ�" Or cboPriceMethod.Text = "�ۼ۳ɱ���һ�����") Then   '����
                    gstrSQL = "Select a.�ɱ��� As ԭ��, a.���ۼ� As �ּ�" & vbNewLine & _
                        "From ҩƷ�շ���¼ A, �շѼ�Ŀ B" & vbNewLine & _
                        "Where a.�۸�id = b.Id And b.���ۻ��ܺ� = [1] And a.�ⷿid = [2] And a.ҩƷid = [3] And Nvl(a.����, 0) = [4]"
                    Set rsPirce = zlDataBase.OpenSQLRecord(gstrSQL, "��ȡ�ۼ�", txtNO.Text, rsTemp!�ⷿid, rsTemp!ҩƷid, Nvl(rsTemp!����, 0))
                    
                    If Not rsPirce.EOF Then
                        .TextMatrix(.rows - 1, menuStoreCol.ԭ���ۼ�) = GetFormat(Val(rsPirce!ԭ��) * dbl��װ����, mintPriceDigit)
                        .TextMatrix(.rows - 1, menuStoreCol.�����ۼ�) = GetFormat(Val(rsPirce!�ּ�) * dbl��װ����, mintPriceDigit)
                        .TextMatrix(.rows - 1, menuStoreCol.�������) = Format(rsTemp!���� / dbl��װ���� * (Val(.TextMatrix(.rows - 1, menuStoreCol.�����ۼ�)) - Val(.TextMatrix(.rows - 1, menuStoreCol.ԭ���ۼ�))), mstrMoneyFormat)
                    Else
                        .TextMatrix(.rows - 1, menuStoreCol.ԭ���ۼ�) = GetFormat(dblOldPrice, mintPriceDigit)
                        .TextMatrix(.rows - 1, menuStoreCol.�����ۼ�) = GetFormat(dblNewPrice, mintPriceDigit)
                        .TextMatrix(.rows - 1, menuStoreCol.�������) = Format(rsTemp!���� / dbl��װ���� * (Val(.TextMatrix(.rows - 1, menuStoreCol.�����ۼ�)) - Val(.TextMatrix(.rows - 1, menuStoreCol.ԭ���ۼ�))), mstrMoneyFormat)
                    End If
                    If cboPriceMethod.Text = "�����ۼ�" Then
                        gstrSQL = "Select �ɱ���" & vbNewLine & _
                                    "      From (Select ƽ���ɱ��� As �ɱ���" & vbNewLine & _
                                    "             From ҩƷ���" & vbNewLine & _
                                    "             Where ����=1 And �ⷿid = [1] And ҩƷid = [2] And nvl(����,0) = [3]" & vbNewLine & _
                                    "             Union All" & vbNewLine & _
                                    "             Select �ɱ��� From ҩƷ��� Where ҩƷid = [2])" & vbNewLine & _
                                    "      Where Rownum <= 1"

                        Set rsCost = zlDataBase.OpenSQLRecord(gstrSQL, "��ȡ�ɱ���", rsTemp!�ⷿid, rsTemp!ҩƷid, Nvl(rsTemp!����, 0))
                        .TextMatrix(.rows - 1, menuStoreCol.ԭ�ɹ���) = GetFormat(rsCost!�ɱ��� * dbl��װ����, mintCostDigit)
                        .TextMatrix(.rows - 1, menuStoreCol.�ֲɹ���) = GetFormat(rsCost!�ɱ��� * dbl��װ����, mintCostDigit)
                        .TextMatrix(.rows - 1, menuStoreCol.��۲�) = Format(0, mstrMoneyFormat)
                    Else
                        .TextMatrix(.rows - 1, menuStoreCol.ԭ�ɹ���) = GetFormat(Nvl(rsTemp!ԭ�ɱ���, 0) * dbl��װ����, mintCostDigit)
                        .TextMatrix(.rows - 1, menuStoreCol.�ֲɹ���) = GetFormat(rsTemp!�³ɱ��� * dbl��װ����, mintCostDigit)
                        .TextMatrix(.rows - 1, menuStoreCol.��۲�) = Format((rsTemp!�³ɱ��� * dbl��װ���� - Nvl(rsTemp!ԭ�ɱ���, 0) * dbl��װ����) * Val(.TextMatrix(.rows - 1, menuStoreCol.����)), mstrMoneyFormat)
                    End If
                Else '�޸Ļ��߳ɱ��۵���
                    '����ֱ�Ӵ��շѼ�Ŀȡ�ּۣ�ʱ�����ȴӿ��ȡ�����û������շѼ�Ŀȡ
                    If Nvl(rsTemp!���, 0) = 1 Then
                        gstrSQL = "Select Nvl(s.���ۼ�, Decode(Nvl(s.ʵ������, 0), 0, 0, Nvl(s.ʵ�ʽ��, 0) / s.ʵ������)) ʱ���ۼ�" & vbNewLine & _
                        "From ҩƷ��� S" & vbNewLine & _
                        "Where s.����=1 And s.�ⷿid = [1] And s.ҩƷid = [2] And nvl(s.����,0) = [3]"
                        
                        Set rsPirce = zlDataBase.OpenSQLRecord(gstrSQL, "��ȡ�ۼ�", rsTemp!�ⷿid, rsTemp!ҩƷid, Nvl(rsTemp!����, 0))
                        If rsPirce.RecordCount > 0 Then
                            If rsPirce!ʱ���ۼ� > 0 Then
                                .TextMatrix(.rows - 1, menuStoreCol.ԭ���ۼ�) = GetFormat(rsPirce!ʱ���ۼ� * dbl��װ����, mintPriceDigit)
                                .TextMatrix(.rows - 1, menuStoreCol.�����ۼ�) = GetFormat(rsPirce!ʱ���ۼ� * dbl��װ����, mintPriceDigit)
                            Else
                                .TextMatrix(.rows - 1, menuStoreCol.ԭ���ۼ�) = GetFormat(dblOldPrice, mintPriceDigit)
                                .TextMatrix(.rows - 1, menuStoreCol.�����ۼ�) = GetFormat(dblNewPrice, mintPriceDigit)
                            End If
                        Else
                            .TextMatrix(.rows - 1, menuStoreCol.ԭ���ۼ�) = GetFormat(dblOldPrice, mintPriceDigit)
                            .TextMatrix(.rows - 1, menuStoreCol.�����ۼ�) = GetFormat(dblNewPrice, mintPriceDigit)
                        End If
                    Else
                        .TextMatrix(.rows - 1, menuStoreCol.ԭ���ۼ�) = GetFormat(dblOldPrice, mintPriceDigit)
                        .TextMatrix(.rows - 1, menuStoreCol.�����ۼ�) = GetFormat(dblNewPrice, mintPriceDigit)
                    End If
                    .TextMatrix(.rows - 1, menuStoreCol.�������) = Format(rsTemp!���� / dbl��װ���� * (Val(.TextMatrix(.rows - 1, menuStoreCol.�����ۼ�)) - Val(.TextMatrix(.rows - 1, menuStoreCol.ԭ���ۼ�))), mstrMoneyFormat)
                    .TextMatrix(.rows - 1, menuStoreCol.ԭ�ɹ���) = GetFormat(Nvl(rsTemp!ԭ�ɱ���, 0) * dbl��װ����, mintCostDigit)
                    .TextMatrix(.rows - 1, menuStoreCol.�ֲɹ���) = GetFormat(rsTemp!�³ɱ��� * dbl��װ����, mintCostDigit)
                    .TextMatrix(.rows - 1, menuStoreCol.��۲�) = Format((rsTemp!�³ɱ��� * dbl��װ���� - Nvl(rsTemp!ԭ�ɱ���, 0) * dbl��װ����) * Val(.TextMatrix(.rows - 1, menuStoreCol.����)), mstrMoneyFormat)
                End If
                 
                If cboPriceMethod.Text = "�����ɱ���" Or cboPriceMethod.Text = "�ۼ۳ɱ���һ�����" Then
                    If rsTemp!�³ɱ��� = 0 Then
                        dbl�ӳ��� = 0
                    Else
                        dbl�ӳ��� = Round(Val(.TextMatrix(.rows - 1, menuStoreCol.�����ۼ�)) / (rsTemp!�³ɱ��� * dbl��װ����) - 1, 7)
                    End If
                    .TextMatrix(.rows - 1, menuStoreCol.�ӳ���) = GetFormat(GetFormat(dbl�ӳ���, 5) * 100, 5)
                    .TextMatrix(.rows - 1, menuStoreCol.ԭ�ɹ���) = GetFormat(Nvl(rsTemp!ԭ�ɱ���, 0) * dbl��װ����, mintCostDigit)
                    .TextMatrix(.rows - 1, menuStoreCol.�ֲɹ���) = GetFormat(rsTemp!�³ɱ��� * dbl��װ����, mintCostDigit)
                    .TextMatrix(.rows - 1, menuStoreCol.��۲�) = Format((rsTemp!�³ɱ��� * dbl��װ���� - Nvl(rsTemp!ԭ�ɱ���, 0) * dbl��װ����) * Val(.TextMatrix(.rows - 1, menuStoreCol.����)), mstrMoneyFormat)
                    dbl��Ʊ��� = dbl��Ʊ��� + (rsTemp!�³ɱ��� * dbl��װ���� - Nvl(rsTemp!ԭ�ɱ���, 0) * dbl��װ����) * Val(.TextMatrix(.rows - 1, menuStoreCol.����))
                    str��Ʊ = IIf(IsNull(rsTemp!��Ʊ��), "", rsTemp!��Ʊ��)
                    str��Ʊ���� = IIf(IsNull(rsTemp!��Ʊ����), "", rsTemp!��Ʊ����)
                    
                    'Ϊ�����¼�б�ֵ
                    If vsfPay.rows > 1 Then
                        bln��ͬҩƷ = False
                        For i = 1 To vsfPay.rows - 1
                            If vsfPay.TextMatrix(i, menuPayCol.ҩƷid) = rsTemp!ҩƷid Then
                                bln��ͬҩƷ = True
                                Exit For
                            End If
                        Next
                        If bln��ͬҩƷ = True Then
                            vsfPay.TextMatrix(i, menuPayCol.��Ʊ���) = GetFormat(Val(vsfPay.TextMatrix(i, menuPayCol.��Ʊ���)) + dbl��Ʊ���, mintMoneyDigit)
                        Else
                            vsfPay.rows = vsfPay.rows + 1
                            vsfPay.RowHeight(vsfPay.rows - 1) = mlngRowHeight
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.ҩƷid) = rsTemp!ҩƷid
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.Ʒ��) = strҩƷ����
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.��Ʊ��) = str��Ʊ
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.��Ʊ����) = Format(str��Ʊ����, "yyyy-mm-dd")
                            vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.��Ʊ���) = GetFormat(dbl��Ʊ���, mintMoneyDigit)
                        End If
                    Else
                        vsfPay.rows = vsfPay.rows + 1
                        vsfPay.RowHeight(vsfPay.rows - 1) = mlngRowHeight
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.ҩƷid) = rsTemp!ҩƷid
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.Ʒ��) = strҩƷ����
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.��Ʊ��) = str��Ʊ
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.��Ʊ����) = Format(str��Ʊ����, "yyyy-mm-dd")
                        vsfPay.TextMatrix(vsfPay.rows - 1, menuPayCol.��Ʊ���) = GetFormat(dbl��Ʊ���, mintMoneyDigit)
                    End If
                End If
            End If
            rsTemp.MoveNext
        Loop
    End With
    '�޸ĺͲ���ʱ�������б�ƽ���ɱ��ۣ��ۼ�
    'mintModal 0-���� 1-�޸� 2-����
'    If mintModal = 1 Or mintModal = 2 Then
        With vsfStore
            For i = 1 To .rows - 1
                If lngҩƷID <> .TextMatrix(i, menuStoreCol.ҩƷid) Then
                    Call CaluateAverCost(Val(.TextMatrix(i, menuStoreCol.ҩƷid)))
                    Call CaluateAverOldCost(Val(.TextMatrix(i, menuStoreCol.ҩƷid)))
                    Call CaculateAverPirce(Val(.TextMatrix(i, menuStoreCol.ҩƷid)))
                    Call CaculateAverOldPirce(Val(.TextMatrix(i, menuStoreCol.ҩƷid)))
                    lngҩƷID = Val(.TextMatrix(i, menuStoreCol.ҩƷid))
                End If
            Next
        End With
'    End If

    If mint���� = 1 Or mint���� = 2 Then
        If rsTemp.RecordCount = 0 Then Exit Sub
        TabCtlDetails.Item(1).Visible = True
    End If

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfPrice_DblClick()
    With vsfPrice
        If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mconlngCanColColor Then
            .EditCell
            .EditSelStart = 0
            .EditSelLength = Len(.EditText)
        End If
    End With
End Sub

Private Sub vsfPrice_EnterCell()
    Dim i As Integer

    With vsfPrice
        .Editable = flexEDNone
        If .CellBackColor = mconlngColor Then
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
        End If

        If .Col = menuPriceCol.�����ۼ� Then
            mdblOldPrice = Val(vsfPrice.TextMatrix(.Row, menuPriceCol.�����ۼ�))
        ElseIf .Col = menuPriceCol.�ֳɱ��� Then
            mdblOldPrice = Val(vsfPrice.TextMatrix(.Row, menuPriceCol.�ֳɱ���))
        End If
    End With
    With vsfStore
        If Val(vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.ҩƷid)) = 0 Then Exit Sub

        If .rows > 1 Then
            For i = 1 To .rows - 1
                If Val(vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.ҩƷid)) = Val(.TextMatrix(i, menuStoreCol.ҩƷid)) Then
                    .Select i, 0, i, .Cols - 1
                    .TopRow = i
                End If
            Next
        End If
    End With
End Sub

Private Sub vsfPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intRow As Integer
    Dim intCol As Integer
    Dim lngDrugID As Long
    Dim strRow As String

    With vsfPrice
        If KeyCode = vbKeyReturn Then
            If .Col <> menuPriceCol.�����ۼ� Then '�ɱ��۵���
                If .Col = menuPriceCol.Ʒ�� And cboPriceMethod.Text = "�����ɱ���" Then
                    .Col = menuPriceCol.�ֳɱ���
'                    .EditCell
                ElseIf .Col = menuPriceCol.Ʒ�� And cboPriceMethod.Text = "�����ۼ�" Then
                    .Col = menuPriceCol.�����ۼ�
'                    .EditCell
                ElseIf .Col = menuPriceCol.�ֳɱ��� And cboPriceMethod.Text = "�����ɱ���" Then
                    If .Row = .rows - 1 And Val(.TextMatrix(.Row, menuPriceCol.ҩƷid)) <> 0 Then
                        .rows = .rows + 1
                        .Row = .Row + 1
                        .Col = menuPriceCol.Ʒ��
                        .RowHeight(.rows - 1) = mlngRowHeight
'                        .EditCell
                        Call setColEdit
                    ElseIf Val(.TextMatrix(.Row, menuPriceCol.ҩƷid)) <> 0 Then
                        .ColComboList(menuPriceCol.Ʒ��) = ""
                        .Row = .Row + 1
                        .Col = menuPriceCol.Ʒ��
                    End If
                ElseIf .Col = menuPriceCol.Ʒ�� And cboPriceMethod.Text = "�ۼ۳ɱ���һ�����" Then
                    .Col = menuPriceCol.�ֳɱ���
'                    .EditCell
                ElseIf .Col = menuPriceCol.�ֳɱ��� And cboPriceMethod.Text = "�ۼ۳ɱ���һ�����" Then
                    .Col = menuPriceCol.�����ۼ�
'                    .EditCell
                ElseIf .Col = menuPriceCol.�����ۼ� And cboPriceMethod.Text = "�ۼ۳ɱ���һ�����" Then
                    If .Row = .rows - 1 Then
                        .rows = .rows + 1
                        .Row = .Row + 1
                        .Col = menuPriceCol.Ʒ��
                        .RowHeight(.rows - 1) = mlngRowHeight
'                        .EditCell
                        Call setColEdit
                    ElseIf Val(.TextMatrix(.Row, menuPriceCol.ҩƷid)) <> 0 Then
                        .ColComboList(menuPriceCol.Ʒ��) = ""
                        .Row = .Row + 1
                        .Col = menuPriceCol.Ʒ��
'                        .EditCell
                    End If
                Else
                    .Col = .Col + 1
'                    .EditCell
                End If
            Else
                If Val(.TextMatrix(.Row, menuPriceCol.ҩƷid)) <> 0 And .Row = .rows - 1 Then
                    .ColComboList(menuPriceCol.Ʒ��) = ""
                    .rows = .rows + 1
                    .Row = .Row + 1
                    .Col = menuPriceCol.Ʒ��
                    .RowHeight(.rows - 1) = mlngRowHeight
'                    .EditCell
                    Call setColEdit
                ElseIf Val(.TextMatrix(.Row, menuPriceCol.ҩƷid)) <> 0 Then
                    .ColComboList(menuPriceCol.Ʒ��) = ""
                    .Row = .Row + 1
                    .Col = menuPriceCol.Ʒ��
'                    .EditCell
                End If
            End If
        ElseIf KeyCode = vbKeyDelete Then
            lngDrugID = Val(vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.ҩƷid))
            If lngDrugID = 0 Then Exit Sub
            If MsgBox("�Ƿ����ɾ�����ҩƷ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            
            '�޸�ģʽʱɾ��һ���۸�������ݣ������δִ�м۸�
            If mintModal = 1 Then
                gstrSQL = "Zl_ɾ��δִ�м۸�_Delete(" & lngDrugID & "," & 0 & ")"
                ReDim Preserve marrSql(UBound(marrSql) + 1)
                marrSql(UBound(marrSql)) = gstrSQL
            End If
            
            If .rows > 2 Then
                .RemoveItem .Row
            Else
                For intCol = 0 To .Cols - 1
                    .TextMatrix(.Row, intCol) = ""
                Next
            End If

            With vsfStore
                If lngDrugID = 0 Then Exit Sub
                For intRow = .rows - 1 To 1 Step -1
                    If Val(.TextMatrix(intRow, menuStoreCol.ҩƷid)) = lngDrugID Then
                        .RemoveItem intRow
                    End If
                Next
            End With

            With vsfPay
                If lngDrugID = 0 Then Exit Sub
                For intRow = .rows - 1 To 1 Step -1
                    If Val(.TextMatrix(intRow, menuPayCol.ҩƷid)) = lngDrugID Then
                        .RemoveItem intRow
                    End If
                Next
            End With
        End If
    End With
End Sub

Private Sub vsfPrice_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim mrsReturn As Recordset
    Dim rsTemp As Recordset
    Dim vRect As RECT
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim strkey As String
    Dim lngDrugID As Long
    Dim intCurrentPirce As Integer '�Ƿ���ʱ��

    On Error GoTo errHandle
    If KeyCode <> vbKeyReturn Then Exit Sub
    mBlnClick = True
    vRect = GetControlRect(vsfPrice.hWnd) '��ȡλ��
    dblLeft = vRect.Left + vsfPrice.CellLeft
    dblTop = vRect.Top + vsfPrice.CellTop + vsfPrice.CellHeight

    With vsfPrice
        strkey = .EditText
        Select Case Col
        Case menuPriceCol.Ʒ��
            If grsMaster.State = adStateClosed Then
                Call SetSelectorRS(1, "", 0, , , , , , , , , True)
            End If
            Set mrsReturn = frmSelector.ShowME(Me, 1, 1, strkey, dblLeft, dblTop, , , , , , , , , False, mstrPrivs)
            If mrsReturn.RecordCount = 0 Then Exit Sub
            mblnUpdateAdd = True
            Call GetDrugPirce(mrsReturn, Row)
            mblnUpdateAdd = False
        End Select
    End With

    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CheckDoubleDrug(ByVal rsTemp As ADODB.Recordset) As ADODB.Recordset
    '����Ƿ����ظ���ҩƷ
    'lngDrugId ҩƷid
    '����ֵ true-�����ظ�ֵ false-�������ظ�ֵ
    Dim i As Integer
    Dim j As Integer
    Dim strTemp As String
    Dim strName As String
    Dim intCount As Integer
    Dim intLength As Integer

    If rsTemp.RecordCount = 0 Then Exit Function
    rsTemp.MoveFirst
    With vsfPrice
        For i = 0 To rsTemp.RecordCount - 1
            For j = 1 To .rows - 1
                If Val(.TextMatrix(j, menuPriceCol.ҩƷid)) = rsTemp!ҩƷid Then
                    strTemp = strTemp & " ҩƷid <> " & rsTemp!ҩƷid & " and "
                    intCount = intCount + 1
                    If intCount < 5 Then
                        strName = strName & rsTemp!ͨ���� & " "
                    End If
                End If
            Next
            rsTemp.MoveNext
        Next
    End With

    If strTemp <> "" Then
        intLength = LenB(StrConv(strTemp, vbFromUnicode)) '�õ��ַ�������
        Do Until Mid(strTemp, intLength, 3) = "and" '�Ӻ���ǰ���ҵ�����һ��"and"
           intLength = intLength - 1
        Loop
        strTemp = Left(strTemp, intLength - 1) '������һ��"and"֮ǰ���ַ���

        rsTemp.Filter = strTemp
        MsgBox strName & "��" & intCount & "��ҩƷ���б����Ѿ����ڣ��Ѵ���ҩƷ������ӣ�", vbInformation, gstrSysName
    End If

    Set CheckDoubleDrug = rsTemp
End Function

Private Sub vsfPrice_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        With vsfPrice
            If .Col = menuPriceCol.Ʒ�� Then
                .Editable = flexEDKbdMouse
                Exit Sub
            End If
            If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mconlngCanColColor Then
                .Editable = flexEDKbdMouse
            Else
                .Editable = flexEDNone
            End If
        End With
    End If
End Sub

Private Sub vsfPrice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strkey As String
    Dim intDigit As Integer

    With vsfPrice
        strkey = .EditText
        If .Col = menuPriceCol.�ֳɱ��� Then
            mdbl�ɱ��� = Val(.TextMatrix(Row, Col))
        End If
    End With

    If Col = menuPriceCol.�ֳɱ��� Or Col = menuPriceCol.�����ۼ� Then
        If KeyAscii = vbKeyReturn Then Exit Sub
        If KeyAscii <> vbKeyBack Then
            Select Case Col
                Case menuPriceCol.�ֳɱ���
                    intDigit = mintCostDigit
                Case menuPriceCol.�����ۼ�
                    intDigit = mintPriceDigit
            End Select

            If KeyAscii = vbKeyDelete Then
                If InStr(1, strkey, ".") > 0 Then
                    KeyAscii = 0
                End If
            ElseIf KeyAscii = Asc(".") Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
                If vsfPrice.EditSelLength = Len(strkey) Then Exit Sub
                If InStr(strkey, ".") <> 0 And Chr(KeyAscii) = "." Then   'ֻ�ܴ���һ��С����
                    KeyAscii = 0
                    Exit Sub
                End If
                If Len(Mid(strkey, InStr(1, strkey, ".") + 1)) >= intDigit And strkey Like "*.*" Then
                    KeyAscii = 0
                    Exit Sub
                Else
                    Exit Sub
                End If
            Else
                KeyAscii = 0
            End If
        End If
    ElseIf Col = menuPriceCol.Ʒ�� Then
        If InStr("`~!@#$%^&*()_-+={[}]|\:;""'<,>.?/", Chr(KeyAscii)) > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub vsfPrice_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If Col = menuPriceCol.Ʒ�� Then
        vsfPrice.ColComboList(menuPriceCol.Ʒ��) = "|..."
    End If
End Sub

Private Sub setColEdit()
    '���ܣ��������Ƿ�����޸�
    '�����޸ĵ�����ɫΪ��ɫ�����޸ĵ�����ɫΪ��ɫ
    Dim intCol As Integer
    Dim intRow As Integer

    With vsfPrice
        .Cell(flexcpBackColor, 1, 1, .rows - 1, .Cols - 1) = mconlngColor
        If cboPriceMethod.Text = "�����ۼ�" Then
            .Cell(flexcpBackColor, 1, menuPriceCol.Ʒ��, .rows - 1, menuPriceCol.Ʒ��) = mconlngCanColColor
            .Cell(flexcpBackColor, 1, menuPriceCol.�����ۼ�, .rows - 1, menuPriceCol.�����ۼ�) = mconlngCanColColor
        ElseIf cboPriceMethod.Text = "�����ɱ���" Then
            .Cell(flexcpBackColor, 1, menuPriceCol.Ʒ��, .rows - 1, menuPriceCol.Ʒ��) = mconlngCanColColor
            .Cell(flexcpBackColor, 1, menuPriceCol.�ֳɱ���, .rows - 1, menuPriceCol.�ֳɱ���) = mconlngCanColColor
        Else
            .Cell(flexcpBackColor, 1, menuPriceCol.Ʒ��, .rows - 1, menuPriceCol.Ʒ��) = mconlngCanColColor
            .Cell(flexcpBackColor, 1, menuPriceCol.�ֳɱ���, .rows - 1, menuPriceCol.�ֳɱ���) = mconlngCanColColor
            .Cell(flexcpBackColor, 1, menuPriceCol.�����ۼ�, .rows - 1, menuPriceCol.�����ۼ�) = mconlngCanColColor
        End If

    End With

    With vsfStore
        If .rows = 1 Then Exit Sub
        .Cell(flexcpBackColor, 1, 0, .rows - 1, .Cols - 1) = mconlngColor
        If cboPriceMethod.Text = "�����ۼ�" Then
            .Cell(flexcpBackColor, 1, menuStoreCol.�����ۼ�, .rows - 1, menuStoreCol.�����ۼ�) = mconlngCanColColor
        ElseIf cboPriceMethod.Text = "�����ɱ���" Then
'            .Cell(flexcpBackColor, 1, menuStoreCol.�ӳ���, .rows - 1, menuStoreCol.�ӳ���) = mconlngCanColColor
            .Cell(flexcpBackColor, 1, menuStoreCol.�ֲɹ���, .rows - 1, menuStoreCol.�ֲɹ���) = mconlngCanColColor
        Else
            .Cell(flexcpBackColor, 1, menuStoreCol.�ӳ���, .rows - 1, menuStoreCol.�ӳ���) = mconlngCanColColor
            .Cell(flexcpBackColor, 1, menuStoreCol.�ֲɹ���, .rows - 1, menuStoreCol.�ֲɹ���) = mconlngCanColColor
            .Cell(flexcpBackColor, 1, menuStoreCol.�����ۼ�, .rows - 1, menuStoreCol.�����ۼ�) = mconlngCanColColor
        End If
        If .rows > 1 Then
            For intRow = 1 To .rows - 1
                If Val(.TextMatrix(intRow, menuStoreCol.���)) = 1 And mblnʱ��ҩƷ�����ε��� = True And mint���� <> 1 Then
                    .Cell(flexcpBackColor, intRow, menuStoreCol.�����ۼ�, intRow, menuStoreCol.�����ۼ�) = mconlngCanColColor
                Else
                    .Cell(flexcpBackColor, intRow, menuStoreCol.�����ۼ�, intRow, menuStoreCol.�����ۼ�) = mconlngColor
                End If
            Next
        End If
    End With

    With vsfPay
        If .rows = 1 Then Exit Sub
        .Cell(flexcpBackColor, 1, 0, .rows - 1, .Cols - 1) = mconlngColor
        .Cell(flexcpBackColor, 1, menuPayCol.��Ʊ��, .rows - 1, menuPayCol.��Ʊ��) = mconlngCanColColor
        .Cell(flexcpBackColor, 1, menuPayCol.��Ʊ����, .rows - 1, menuPayCol.��Ʊ����) = mconlngCanColColor
        .Cell(flexcpBackColor, 1, menuPayCol.��Ʊ���, .rows - 1, menuPayCol.��Ʊ���) = mconlngCanColColor
    End With
End Sub


Private Sub vsfPrice_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        vsfPrice.Editable = flexEDNone
        If vsfPrice.Col = menuPriceCol.Ʒ�� And mintModal <> 2 Then
            vsfPrice.ColComboList(menuPriceCol.Ʒ��) = "|..."
            vsfPrice.Editable = flexEDKbdMouse
        End If
    End If
End Sub

Private Sub vsfPrice_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngDrugID As Long
    Dim dblSalePrice As Double
    Dim intRow As Integer
    Dim dbl�ӳ��� As Double

    With vsfPrice
        If .EditText = "" Then Exit Sub
        lngDrugID = Val(.TextMatrix(Row, menuPriceCol.ҩƷid))
        If lngDrugID = 0 Then Exit Sub

        Select Case Col
            Case menuPriceCol.�ֳɱ���
                If Val(.EditText) < 0 Then
                    MsgBox "�ɱ��۲���Ϊ������", vbExclamation, gstrSysName
                    Cancel = True
                End If
                If Not IsNumeric(.EditText) Then
                    Cancel = True
                    Exit Sub
                End If
                If .EditText > 9999999 Then
                    MsgBox "�ɱ��۹������������룡", vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
                .EditText = GetFormat(.EditText, mintPriceDigit)
                If mbln�ּ���ʾ = True Then
                    If Val(.EditText) > Val(.TextMatrix(Row, menuPriceCol.ԭ�ɹ��޼�)) Then
                        If MsgBox("�ֳɱ��۸��ڲɹ����޼�" & Val(.TextMatrix(.Row, menuPriceCol.ԭ�ɹ��޼�)) & "��" & vbCrLf & "������", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                End If
                .TextMatrix(.Row, menuPriceCol.�ֲɹ��޼�) = GetFormat(.EditText, mintCostDigit)

                If cbo�ۼۼ��㷽ʽ.Text = "�ۼ۰��ֶμӳɼ���" And .TextMatrix(.Row, menuPriceCol.�Ƿ���) = "1" And mint���� = 2 Then
                    Call get�ֶμӳ��ۼ�(lngDrugID, Val(.TextMatrix(.Row, menuPriceCol.��װϵ��)), Val(.EditText), dblSalePrice)
                    If dblSalePrice = 0 Then
                        .EditText = mdbl�ɱ���
                        .TextMatrix(vsfPrice.Row, menuPriceCol.�ֳɱ���) = GetFormat(.EditText, mintCostDigit)
                        Exit Sub
                    End If
                    dblSalePrice = dblSalePrice + (Val(.TextMatrix(.Row, menuPriceCol.ԭָ���ۼ�)) - dblSalePrice) * (1 - Val(.TextMatrix(.Row, menuPriceCol.���������)) / 100)
                    .TextMatrix(.Row, menuPriceCol.�����ۼ�) = GetFormat(dblSalePrice, mintPriceDigit)
                    
                    '�����ۼ�Ӧ��ͬ�����¿���б�۸���Ϣ
                    If vsfStore.rows > 1 Then
                        For intRow = 1 To vsfStore.rows - 1
                            If vsfStore.TextMatrix(intRow, menuStoreCol.ҩƷid) = .TextMatrix(.Row, menuPriceCol.ҩƷid) Then
                                vsfStore.TextMatrix(intRow, menuStoreCol.�����ۼ�) = GetFormat(dblSalePrice, mintPriceDigit)
                                vsfStore.TextMatrix(intRow, menuStoreCol.�������) = Format(Val(vsfStore.TextMatrix(intRow, menuStoreCol.����)) * (Val(vsfStore.TextMatrix(intRow, menuStoreCol.�����ۼ�)) - Val(vsfStore.TextMatrix(intRow, menuStoreCol.ԭ���ۼ�))), mstrMoneyFormat)
                                
                                If Val(vsfStore.TextMatrix(intRow, menuStoreCol.�ֲɹ���)) <> 0 Then
                                    dbl�ӳ��� = GetFormat(GetFormat(((Val(vsfStore.TextMatrix(intRow, menuStoreCol.�����ۼ�))) / Val(vsfStore.TextMatrix(intRow, menuStoreCol.�ֲɹ���)) - 1), 5) * 100, 5)
                                Else
                                    dbl�ӳ��� = 0
                                End If
                                vsfStore.TextMatrix(intRow, menuStoreCol.�ӳ���) = dbl�ӳ���
                            End If
                        Next
                    End If
                ElseIf cbo�ۼۼ��㷽ʽ = "�ۼ۰��̶���������" And .TextMatrix(.Row, menuPriceCol.�Ƿ���) = "1" And mint���� = 2 Then
                    dblSalePrice = Val(.EditText) * (1 + Val(.TextMatrix(.Row, menuPriceCol.�ӳ���)))
                    If dblSalePrice > Val(.TextMatrix(.Row, menuPriceCol.ԭָ���ۼ�)) Then dblSalePrice = Val(.TextMatrix(.Row, menuPriceCol.ԭָ���ۼ�))
                    .TextMatrix(.Row, menuPriceCol.�����ۼ�) = GetFormat(dblSalePrice, mintPriceDigit)
                    
                    '�����ۼ�Ӧ��ͬ�����¿���б�۸���Ϣ
                    If vsfStore.rows > 1 Then
                        For intRow = 1 To vsfStore.rows - 1
                            If vsfStore.TextMatrix(intRow, menuStoreCol.ҩƷid) = .TextMatrix(.Row, menuPriceCol.ҩƷid) Then
                                vsfStore.TextMatrix(intRow, menuStoreCol.�����ۼ�) = GetFormat(dblSalePrice, mintPriceDigit)
                                vsfStore.TextMatrix(intRow, menuStoreCol.�������) = Format(Val(vsfStore.TextMatrix(intRow, menuStoreCol.����)) * (Val(vsfStore.TextMatrix(intRow, menuStoreCol.�����ۼ�)) - Val(vsfStore.TextMatrix(intRow, menuStoreCol.ԭ���ۼ�))), mstrMoneyFormat)
                                
                                If Val(vsfStore.TextMatrix(intRow, menuStoreCol.�ֲɹ���)) <> 0 Then
                                    dbl�ӳ��� = GetFormat(GetFormat(((Val(vsfStore.TextMatrix(intRow, menuStoreCol.�����ۼ�))) / Val(vsfStore.TextMatrix(intRow, menuStoreCol.�ֲɹ���)) - 1), 5) * 100, 5)
                                Else
                                    dbl�ӳ��� = 0
                                End If
                                vsfStore.TextMatrix(intRow, menuStoreCol.�ӳ���) = dbl�ӳ���
                            End If
                        Next
                    End If
                End If

                Call CaculateCost(lngDrugID, .EditText) '���¼���ɱ���
            Case menuPriceCol.�����ۼ�
                If Val(.EditText) < 0 Then
                    MsgBox "�ۼ۲���Ϊ������", vbExclamation, gstrSysName
                    Cancel = True
                End If
                If Not IsNumeric(.EditText) Then
                    Cancel = True
                    Exit Sub
                End If

                If .EditText > 9999999 Then
                    MsgBox "���ۼ۹������������룡", vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If

                .EditText = GetFormat(.EditText, mintPriceDigit)
'                If mdblOldPrice = .EditText Then 'δ���޸�ֱ���˳�
'                    Exit Sub
'                End If

                If mbln�ּ���ʾ = True Then
                    If Val(.EditText) > Val(.TextMatrix(Row, menuPriceCol.ԭָ���ۼ�)) Then
                        If MsgBox("�����ۼ۸���ָ���ۼ�" & Val(.TextMatrix(.Row, menuPriceCol.ԭָ���ۼ�)) & "��" & vbCrLf & "������", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                End If
                .TextMatrix(.Row, menuPriceCol.��ָ���ۼ�) = GetFormat(.EditText, mintPriceDigit)
                If chkAotuCost.Value = 1 Then '�޸��ۼۺ��Զ�����ɱ���
                    .TextMatrix(.Row, menuPriceCol.�ֳɱ���) = GetFormat(.EditText / (1 + Val(.TextMatrix(.Row, menuPriceCol.�ӳ���))), mintCostDigit)
                    
                    If vsfStore.rows > 1 Then
                        For intRow = 1 To vsfStore.rows - 1
                            If vsfStore.TextMatrix(intRow, menuStoreCol.ҩƷid) = .TextMatrix(.Row, menuPriceCol.ҩƷid) Then
                                vsfStore.TextMatrix(intRow, menuStoreCol.�ֲɹ���) = GetFormat(.TextMatrix(.Row, menuPriceCol.�ֳɱ���), mintCostDigit)
                                
                                If Val(vsfStore.TextMatrix(intRow, menuStoreCol.�ֲɹ���)) <> 0 Then
                                    dbl�ӳ��� = GetFormat((.EditText / Val(vsfStore.TextMatrix(intRow, menuStoreCol.�ֲɹ���)) - 1), 5)
                                Else
                                    dbl�ӳ��� = 0
                                End If
                                vsfStore.TextMatrix(intRow, menuStoreCol.�ӳ���) = GetFormat(dbl�ӳ��� * 100, 5)
                                vsfStore.TextMatrix(intRow, menuStoreCol.��۲�) = Format((Val(vsfStore.TextMatrix(intRow, menuStoreCol.�ֲɹ���)) - Val(vsfStore.TextMatrix(intRow, menuStoreCol.ԭ�ɹ���))) * Val(vsfStore.TextMatrix(intRow, menuStoreCol.����)), mstrMoneyFormat)
                            End If
                        Next
                    End If
                End If

                Call ChangeDrugStore(Row, lngDrugID, .EditText)
        End Select
    End With
End Sub

Private Sub ChangeDrugStore(ByVal intRow As Integer, ByVal lngDrugID As Long, ByVal dblNewPrice As Double)
    '���ܣ�ͨ���޸ļ۸���е����ۼ��޸Ŀ���б������Ӧ�����ۼ�
    Dim dblOldPrice As Double
    Dim dblOldCost As Double
    Dim dblNewCost As Double
    Dim dblNum As Double
    Dim dbl��װ As Double
    Dim n As Integer
    Dim dbl��Ʊ��� As Double
    Dim dbl�ӳ��� As Double

    If intRow = 0 Or mint���� = 1 Then Exit Sub

    dbl��װ = Val(vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.��װϵ��))

    With vsfStore
        For n = 1 To .rows - 1
            If .TextMatrix(n, 0) <> "" Then
                If Val(.TextMatrix(n, menuStoreCol.ҩƷid)) = lngDrugID Then
                    dblNum = Val(.TextMatrix(n, menuStoreCol.����))
                    dblOldPrice = Val(vsfStore.TextMatrix(n, menuStoreCol.ԭ���ۼ�))

                    .TextMatrix(n, menuStoreCol.�����ۼ�) = GetFormat(dblNewPrice, mintPriceDigit)
                    .TextMatrix(n, menuStoreCol.�������) = Format(Val(.TextMatrix(n, menuStoreCol.����)) * (dblNewPrice - dblOldPrice), mstrMoneyFormat)
                    
                    If Val(.TextMatrix(n, menuStoreCol.�ֲɹ���)) <> 0 Then
                        dbl�ӳ��� = GetFormat(((Val(.TextMatrix(n, menuStoreCol.�����ۼ�))) / Val(.TextMatrix(n, menuStoreCol.�ֲɹ���)) - 1), 5)
                    Else
                        dbl�ӳ��� = 0
                    End If
                    .TextMatrix(n, menuStoreCol.�ӳ���) = GetFormat(dbl�ӳ��� * 100, 5)
                
                    If mint���� = 2 And chkAotuCost.Value = 1 Then
                        dblOldCost = .TextMatrix(n, menuStoreCol.ԭ�ɹ���)
                        dblNewCost = dblNewPrice / (1 + Round(Val(.TextMatrix(n, menuStoreCol.�ӳ���)) / 100, 7))
                        .TextMatrix(n, menuStoreCol.�ֲɹ���) = GetFormat(dblNewCost, mintCostDigit)
                        .TextMatrix(n, menuStoreCol.��۲�) = Format((.TextMatrix(n, menuStoreCol.�ֲɹ���) - dblOldCost) * dblNum, mstrMoneyFormat)
                    End If
                    dbl��Ʊ��� = dbl��Ʊ��� + Val(.TextMatrix(n, menuStoreCol.��۲�))
                End If
            End If
        Next
    End With

    If chkAutoPay.Value = 1 Then
        With vsfPay
            For n = 1 To .rows - 1
                If .TextMatrix(1, 0) <> "" Then
                    If Val(.TextMatrix(n, menuPayCol.ҩƷid)) = lngDrugID Then
                        .TextMatrix(n, menuPayCol.��Ʊ���) = GetFormat(dbl��Ʊ���, mintMoneyDigit)
                    End If
                End If
            Next
        End With
    End If

    If mint���� = 2 Then
        CaluateAverCost lngDrugID
    End If
End Sub

Private Sub CaluateAverCost(ByVal lngҩƷID As Long)
    '����ƽ���ɱ���
    Dim i As Integer
    Dim dblSumCost As Double
    Dim dblSumNumber As Double

    With vsfStore
        For i = 1 To .rows - 1
            If .TextMatrix(i, menuStoreCol.ҩƷid) <> "" Then
                If Val(.TextMatrix(i, menuStoreCol.ҩƷid)) = lngҩƷID Then
                    dblSumCost = dblSumCost + Val(.TextMatrix(i, menuStoreCol.�ֲɹ���)) * Val(.TextMatrix(i, menuStoreCol.����))
                    dblSumNumber = dblSumNumber + Val(.TextMatrix(i, menuStoreCol.����))
                End If
            End If
        Next
    End With

    With vsfPrice
        If dblSumNumber > 0 Then
            For i = 1 To .rows - 1
                If .TextMatrix(i, menuPriceCol.ҩƷid) <> "" Then
                    If Val(.TextMatrix(i, menuPriceCol.ҩƷid)) = lngҩƷID Then
                        .TextMatrix(i, menuPriceCol.�ֳɱ���) = GetFormat(dblSumCost / dblSumNumber, mintCostDigit)
                        Exit For
                    End If
                End If
            Next
        End If
    End With
End Sub

Private Sub CaluateAverOldCost(ByVal lngҩƷID As Long)
    '����ԭʼƽ���ɱ���
    Dim i As Integer
    Dim dblSumCost As Double
    Dim dblSumNumber As Double

    With vsfStore
        For i = 1 To .rows - 1
            If .TextMatrix(i, menuStoreCol.ҩƷid) <> "" Then
                If Val(.TextMatrix(i, menuStoreCol.ҩƷid)) = lngҩƷID Then
                    dblSumCost = dblSumCost + Val(.TextMatrix(i, menuStoreCol.ԭ�ɹ���)) * Val(.TextMatrix(i, menuStoreCol.����))
                    dblSumNumber = dblSumNumber + Val(.TextMatrix(i, menuStoreCol.����))
                End If
            End If
        Next
    End With

    With vsfPrice
        If dblSumNumber > 0 Then
            For i = 1 To .rows - 1
                If .TextMatrix(i, menuPriceCol.ҩƷid) <> "" Then
                    If Val(.TextMatrix(i, menuPriceCol.ҩƷid)) = lngҩƷID Then
                        .TextMatrix(i, menuPriceCol.ԭ�ɱ���) = GetFormat(dblSumCost / dblSumNumber, mintCostDigit)
                        Exit For
                    End If
                End If
            Next
        End If
    End With
End Sub

Private Sub CaculateCost(ByVal lngҩƷID As Long, ByVal dbl�ֳɱ��� As Double)
    '���ܣ�ͨ���޸ļ۸���еĳɱ����޸Ŀ���б������Ӧ�ĳɱ���

    Dim n As Integer
    Dim dbl��Ʊ��� As Double

    With vsfStore
        For n = 1 To .rows - 1
            If .TextMatrix(n, menuStoreCol.ҩƷid) <> "" Then
                If Val(.TextMatrix(n, menuStoreCol.ҩƷid)) = lngҩƷID Then
                    .TextMatrix(n, menuStoreCol.�ֲɹ���) = GetFormat(dbl�ֳɱ���, mintCostDigit)
                    If (cbo�ۼۼ��㷽ʽ.Text = "�ۼ۰��ֶμӳɼ���" Or cbo�ۼۼ��㷽ʽ.Text = "�ۼ۰��̶���������") And vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.�Ƿ���) = "1" And mint���� = 2 Then
                        .TextMatrix(n, menuStoreCol.�����ۼ�) = vsfPrice.TextMatrix(vsfPrice.Row, menuPriceCol.�����ۼ�)
                    End If
                    If dbl�ֳɱ��� <> 0 Then
                        .TextMatrix(n, menuStoreCol.�ӳ���) = GetFormat(GetFormat((Val(.TextMatrix(n, menuStoreCol.�����ۼ�)) / dbl�ֳɱ��� - 1), 5) * 100, 5)
                    End If
                    If cbo�ۼۼ��㷽ʽ = "�ۼ۰��ֶμӳɼ���" Then
                        .TextMatrix(n, menuStoreCol.�ӳ���) = GetFormat(GetFormat(mdbl�ֶμӳ���, 5) * 100, 5)
                    End If
                    .TextMatrix(n, menuStoreCol.��۲�) = Format((dbl�ֳɱ��� - Val(.TextMatrix(n, menuStoreCol.ԭ�ɹ���))) * Val(.TextMatrix(n, menuStoreCol.����)), mstrMoneyFormat)

                    dbl��Ʊ��� = dbl��Ʊ��� + (dbl�ֳɱ��� - .TextMatrix(n, menuStoreCol.ԭ�ɹ���)) * Val(.TextMatrix(n, menuStoreCol.����))
                    .TextMatrix(n, menuStoreCol.�������) = (Val(.TextMatrix(n, menuStoreCol.�����ۼ�)) - Val(.TextMatrix(n, menuStoreCol.ԭ���ۼ�))) * Val(.TextMatrix(n, menuStoreCol.����))
                End If
            End If
        Next
    End With

    If chkAutoPay.Value = 1 Then
        For n = 1 To vsfPay.rows - 1
            If vsfPay.TextMatrix(1, 0) <> "" Then
                If Val(vsfPay.TextMatrix(n, menuPayCol.ҩƷid)) = lngҩƷID Then
                    vsfPay.TextMatrix(n, menuPayCol.��Ʊ���) = Format(dbl��Ʊ���, mstrMoneyFormat)
                End If
            End If
        Next
    End If
End Sub


Private Sub vsfStore_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsfStore
        .Move 0, 360, TabCtlDetails.Width, TabCtlDetails.Height - 370
    End With
End Sub

Private Sub vsfStore_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfStore
        If .Cell(flexcpBackColor, Row, Col, Row, Col) = mconlngColor Then
            Cancel = True
            .Editable = flexEDNone
        Else
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub

Private Sub setColHiddenVsf()
    '��ͬģʽ���棬����ʾ��һ��
    With vsfStore
        If cboPriceMethod.Text = "�����ۼ�" Then
            .ColHidden(menuStoreCol.����) = True
            .ColHidden(menuStoreCol.���) = True
            .ColHidden(menuStoreCol.�ӳ���) = True
            .ColHidden(menuStoreCol.ԭ�ɹ���) = True
            .ColHidden(menuStoreCol.�ֲɹ���) = True
            .ColHidden(menuStoreCol.��۲�) = True
            .ColHidden(menuStoreCol.ԭ���ۼ�) = False
            .ColHidden(menuStoreCol.�����ۼ�) = False
        ElseIf cboPriceMethod.Text = "�����ɱ���" Then
            .ColHidden(menuStoreCol.ԭ���ۼ�) = True
            .ColHidden(menuStoreCol.�����ۼ�) = True
            .ColHidden(menuStoreCol.�������) = True
            .ColHidden(menuStoreCol.�ӳ���) = False
            .ColHidden(menuStoreCol.ԭ�ɹ���) = False
            .ColHidden(menuStoreCol.�ֲɹ���) = False
            .ColHidden(menuStoreCol.��۲�) = False
        ElseIf cboPriceMethod.Text = "�ۼ۳ɱ���һ�����" Then
            .ColHidden(menuStoreCol.ԭ���ۼ�) = False
            .ColHidden(menuStoreCol.�����ۼ�) = False
            .ColHidden(menuStoreCol.�������) = False
            .ColHidden(menuStoreCol.�ӳ���) = False
            .ColHidden(menuStoreCol.ԭ�ɹ���) = False
            .ColHidden(menuStoreCol.�ֲɹ���) = False
            .ColHidden(menuStoreCol.��۲�) = False
        End If
    End With
End Sub

Private Sub vsfStore_Click()
    Dim i As Integer
    With vsfStore
        For i = 1 To vsfPrice.rows - 1
            If Val(.TextMatrix(.Row, menuStoreCol.ҩƷid)) = Val(vsfPrice.TextMatrix(i, menuPriceCol.ҩƷid)) Then
                vsfPrice.Tag = i
            End If
        Next
    End With
End Sub

Private Sub vsfStore_DblClick()
    With vsfStore
        If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mconlngCanColColor Then
            .EditCell
            .EditSelStart = 0
            .EditSelLength = Len(.EditText)
        End If
    End With
End Sub

Private Sub vsfStore_EnterCell()
    With vsfStore
        If .CellBackColor = mconlngColor Then
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
        End If
        If .Col = menuStoreCol.�ӳ��� Then
            mdblOldPrice = Val(.TextMatrix(.Row, menuStoreCol.�ӳ���))
        ElseIf .Col = menuStoreCol.�ֲɹ��� Then
            mdblOldPrice = Val(.TextMatrix(.Row, menuStoreCol.�ֲɹ���))
        ElseIf .Col = menuStoreCol.�����ۼ� Then
            mdblOldPrice = Val(.TextMatrix(.Row, menuStoreCol.�����ۼ�))
        End If
    End With
End Sub

Private Sub vsfStore_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsfStore
        If KeyCode = vbKeyReturn Then
            If .Col < vsfStore.Cols - 1 Then
                .Col = .Col + 1
            Else
                If .Row <> .rows - 1 Then
                    .Row = .Row + 1
                    .Col = menuStoreCol.���
                End If
            End If
        End If
    End With
End Sub

Private Sub vsfStore_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        With vsfStore
            If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mconlngCanColColor Then
                .Editable = flexEDKbdMouse
            Else
                .Editable = flexEDNone
            End If
        End With
    End If
End Sub

Private Sub vsfStore_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strkey As String
    Dim intDigit As Integer

    If KeyAscii = vbKeyReturn Then Exit Sub
    If KeyAscii <> vbKeyBack Then
        With vsfStore
            If Col = menuStoreCol.�ֲɹ��� Or Col = menuStoreCol.�����ۼ� Or Col = menuStoreCol.�ӳ��� Then
                strkey = .EditText
                Select Case Col
                    Case menuStoreCol.�ֲɹ���
                        intDigit = mintCostDigit
                    Case menuStoreCol.�����ۼ�
                        intDigit = mintPriceDigit
                    Case menuStoreCol.�ӳ���
                        intDigit = 5
                End Select
                If KeyAscii = vbKeyDelete Then
                    If InStr(1, .EditText, ".") > 0 Then
                        KeyAscii = 0
                    End If
                ElseIf KeyAscii = Asc(".") Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
                    If .EditSelLength = Len(strkey) Then Exit Sub
                    If InStr(strkey, ".") <> 0 And Chr(KeyAscii) = "." Then   'ֻ�ܴ���һ��С����
                        KeyAscii = 0
                        Exit Sub
                    End If
                    If Len(Mid(strkey, InStr(1, strkey, ".") + 1)) >= intDigit And strkey Like "*.*" Then
                        KeyAscii = 0
                        Exit Sub
                    Else
                        Exit Sub
                    End If
                Else
                    KeyAscii = 0
                End If
            End If
        End With
    End If
End Sub

Private Sub vsfStore_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strInput As String
    Dim n As Integer
    Dim intRow As Integer
    Dim dbl��Ʊ��� As Double
    Dim Dbl���� As Double
    Dim Dbl��� As Double
    Dim dbl�ֲɹ��� As Double
    Dim dblTempNum As Double
    Dim dbl�ɱ���� As Double

    With vsfStore
        If .EditText = "" Then Exit Sub
        intRow = .Row
        Select Case .Col
            Case menuStoreCol.�����ۼ�
                If Not IsNumeric(.EditText) Then
                    MsgBox "�������µ��ۼۡ�", vbInformation, gstrSysName
                    Exit Sub
                Else
                    .EditText = GetFormat(.EditText, mintPriceDigit)
                End If

                If .EditText > 9999999 Then
                    MsgBox "���ۼ۹������������룡", vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If

'                If mdblOldPrice = .EditText Then Exit Sub

                If chkAotuCost.Value = 1 Then '�޸��ۼۺ��Զ�����ɱ���
                    .TextMatrix(intRow, menuStoreCol.�ֲɹ���) = GetFormat(.EditText / (1 + Val(.TextMatrix(intRow, menuStoreCol.�ӳ���)) / 100), mintCostDigit)
                    .TextMatrix(intRow, menuStoreCol.��۲�) = Format((Val(vsfStore.TextMatrix(intRow, menuStoreCol.�ֲɹ���)) - Val(vsfStore.TextMatrix(intRow, menuStoreCol.ԭ�ɹ���))) * Val(vsfStore.TextMatrix(intRow, menuStoreCol.����)), mstrMoneyFormat)
                End If
                
                .TextMatrix(intRow, menuStoreCol.�������) = Format(Val(.TextMatrix(intRow, menuStoreCol.����)) * (Val(.EditText) - Val(.TextMatrix(intRow, menuStoreCol.ԭ���ۼ�))), mstrMoneyFormat)
                .TextMatrix(intRow, menuStoreCol.�����ۼ�) = GetFormat(Val(.EditText), mintPriceDigit)
'                .TextMatrix(intRow, menuStoreCol.�ֲɹ���) = GetFormat(Val(.TextMatrix(intRow, menuStoreCol.�����ۼ�)) / (1 + Val(.TextMatrix(intRow, menuStoreCol.�ӳ���)) / 100), mintCostDigit)
'                .TextMatrix(intRow, menuStoreCol.��۲�) = Format((Val(.TextMatrix(intRow, menuStoreCol.�ֲɹ���)) - Val(.TextMatrix(intRow, menuStoreCol.ԭ�ɹ���))) * Val(.TextMatrix(intRow, menuStoreCol.����)), mstrMoneyFormat)
                If chkAotuCost.Value <> 1 Then
                    If Val(.TextMatrix(intRow, menuStoreCol.�ֲɹ���)) <> 0 Then
                        .TextMatrix(intRow, menuStoreCol.�ӳ���) = GetFormat(GetFormat((Val(.TextMatrix(intRow, menuStoreCol.�����ۼ�)) / Val(.TextMatrix(intRow, menuStoreCol.�ֲɹ���)) - 1), 5) * 100, 5)
                    Else
                        .TextMatrix(intRow, menuStoreCol.�ӳ���) = GetFormat(0, 5)
                    End If
                End If
                
                For n = 1 To .rows - 1
                    If .TextMatrix(intRow, menuStoreCol.ҩƷid) = .TextMatrix(n, menuStoreCol.ҩƷid) Then
                        If Val(.TextMatrix(intRow, menuStoreCol.����)) <> 0 And Val(.TextMatrix(intRow, menuStoreCol.����)) = Val(.TextMatrix(n, menuStoreCol.����)) Then
                            .TextMatrix(n, menuStoreCol.�����ۼ�) = .TextMatrix(intRow, menuStoreCol.�����ۼ�)
                            .TextMatrix(n, menuStoreCol.�������) = Format(Val(.TextMatrix(n, menuStoreCol.����)) * (Val(.EditText) - Val(.TextMatrix(n, menuStoreCol.ԭ���ۼ�))), mstrMoneyFormat)
                            If chkAotuCost.Value <> 1 Then
                                If Val(.TextMatrix(n, menuStoreCol.�ֲɹ���)) <> 0 Then
                                    .TextMatrix(n, menuStoreCol.�ӳ���) = GetFormat(GetFormat((Val(.TextMatrix(n, menuStoreCol.�����ۼ�)) / Val(.TextMatrix(n, menuStoreCol.�ֲɹ���)) - 1), 5) * 100, 5)
                                Else
                                    .TextMatrix(n, menuStoreCol.�ӳ���) = GetFormat(0, 5)
                                End If
                            End If
                        End If
                        Dbl���� = Dbl���� + .TextMatrix(n, menuStoreCol.����)
                        Dbl��� = Dbl��� + .TextMatrix(n, menuStoreCol.����) * Val(.TextMatrix(n, menuStoreCol.�����ۼ�))
                        dbl�ɱ���� = dbl�ɱ���� + .TextMatrix(n, menuStoreCol.����) * Val(.TextMatrix(n, menuStoreCol.�ֲɹ���))
                    End If
                Next
                For n = 1 To vsfPrice.rows - 1
                    If .TextMatrix(intRow, menuStoreCol.ҩƷid) = vsfPrice.TextMatrix(n, menuPriceCol.ҩƷid) Then
                        If Dbl���� <> 0 Then
                            If chkAotuCost.Value = 1 Then
                                vsfPrice.TextMatrix(n, menuPriceCol.�ֳɱ���) = GetFormat(dbl�ɱ���� / Dbl����, mintPriceDigit)
                            End If
                            vsfPrice.TextMatrix(n, menuPriceCol.�����ۼ�) = GetFormat(Dbl��� / Dbl����, mintPriceDigit)
                        Else
                            If chkAotuCost.Value = 1 Then
                                vsfPrice.TextMatrix(n, menuPriceCol.�ֳɱ���) = vsfStore.TextMatrix(intRow, menuStoreCol.�ֲɹ���)
                            End If
                            vsfPrice.TextMatrix(n, menuPriceCol.�����ۼ�) = vsfStore.TextMatrix(intRow, menuStoreCol.�����ۼ�)
                        End If
                    End If
                Next

                If mint���� > 0 Then
                    For n = 1 To .rows - 1
                        If .TextMatrix(n, menuStoreCol.ҩƷid) <> "" Then
                            If Val(.TextMatrix(n, menuStoreCol.ҩƷid)) = Val(.TextMatrix(intRow, menuStoreCol.ҩƷid)) Then
                                dbl��Ʊ��� = dbl��Ʊ��� + (Val(.TextMatrix(n, menuStoreCol.�ֲɹ���)) - Val(.TextMatrix(n, menuStoreCol.ԭ�ɹ���))) * Val(.TextMatrix(n, menuStoreCol.����))
                            End If
                        End If
                    Next

                    If chkAutoPay.Value = 1 Then
                        For n = 1 To vsfPay.rows - 1
                            If vsfPay.TextMatrix(1, 0) <> "" Then
                                If Val(vsfPay.TextMatrix(n, menuPayCol.ҩƷid)) = Val(vsfStore.TextMatrix(intRow, menuStoreCol.ҩƷid)) Then
                                    vsfPay.TextMatrix(n, menuPayCol.��Ʊ���) = GetFormat(dbl��Ʊ���, mintMoneyDigit)
                                End If
                            End If
                        Next
                    End If
                End If
            Case menuStoreCol.�ӳ���
                If Val(.EditText) < 0 Then Exit Sub
                If Not IsNumeric(.EditText) Then
                    Cancel = True
                    Exit Sub
                End If
'                If mdblOldPrice = .EditText Then Exit Sub
                
                .EditText = GetFormat(.EditText, 5)
                .TextMatrix(intRow, menuStoreCol.�ӳ���) = GetFormat(Val(.EditText), 5)
                .TextMatrix(intRow, menuStoreCol.�����ۼ�) = GetFormat(Val(.TextMatrix(intRow, menuStoreCol.�ֲɹ���)) * (1 + Val(.TextMatrix(intRow, menuStoreCol.�ӳ���)) / 100), mintCostDigit)
                .TextMatrix(intRow, menuStoreCol.�������) = Format(Val(.TextMatrix(intRow, menuStoreCol.����)) * (Val(.TextMatrix(intRow, menuStoreCol.�����ۼ�)) - Val(.TextMatrix(intRow, menuStoreCol.ԭ���ۼ�))), mstrMoneyFormat)
                For n = 1 To .rows - 1
                    If vsfPrice.TextMatrix(Val(vsfPrice.Tag), menuPriceCol.ҩƷid) = .TextMatrix(n, menuStoreCol.ҩƷid) Then
                        If Val(.TextMatrix(intRow, menuStoreCol.���)) = 0 Or mblnʱ��ҩƷ�����ε��� = False Then
                            .TextMatrix(n, menuStoreCol.�ӳ���) = GetFormat(Val(.EditText), 5)
                            .TextMatrix(n, menuStoreCol.�����ۼ�) = GetFormat(Val(.TextMatrix(n, menuStoreCol.�ֲɹ���)) * (1 + GetFormat(Val(.EditText), 5) / 100), mintCostDigit)
                            .TextMatrix(n, menuStoreCol.�������) = Format(Val(.TextMatrix(n, menuStoreCol.����)) * (Val(.TextMatrix(n, menuStoreCol.�����ۼ�)) - Val(.TextMatrix(n, menuStoreCol.ԭ���ۼ�))), mstrMoneyFormat)
                        End If
                        Dbl���� = Dbl���� + .TextMatrix(n, menuStoreCol.����)
                        Dbl��� = Dbl��� + .TextMatrix(n, menuStoreCol.����) * Val(.TextMatrix(n, menuStoreCol.�����ۼ�))
                    End If
                Next
                If Dbl���� <> 0 Then
                    vsfPrice.TextMatrix(Val(vsfPrice.Tag), menuPriceCol.�����ۼ�) = GetFormat(Dbl��� / Dbl����, mintPriceDigit)
                Else
                    vsfPrice.TextMatrix(Val(vsfPrice.Tag), menuPriceCol.�����ۼ�) = .TextMatrix(intRow, menuStoreCol.�����ۼ�)
                End If
            Case menuStoreCol.�ֲɹ���
                If Val(.EditText) > Val(.TextMatrix(.Row, menuStoreCol.�����ۼ�)) Then
                    MsgBox "ע�⣬�³ɱ��۴��������ۼۣ�", vbExclamation, gstrSysName
                End If

                If Val(.EditText) < 0 Then
                    MsgBox "�ɱ��۲���Ϊ������", vbExclamation, gstrSysName
                    Cancel = True
                End If
                If .EditText > 9999999 Then
                    MsgBox "�ɹ��۹������������룡", vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
'                If mdblOldPrice = .EditText Then Exit Sub
                
                .EditText = GetFormat(.EditText, mintCostDigit)
                .TextMatrix(intRow, menuStoreCol.�ֲɹ���) = GetFormat(Val(.EditText), mintCostDigit)
'                If Val(.EditText) <> 0 Then
'                    .TextMatrix(intRow, menuStoreCol.�ӳ���) = GetFormat((Val(.TextMatrix(intRow, menuStoreCol.�����ۼ�)) / Val(.EditText) - 1) * 100, 5)
'                End If
                .TextMatrix(intRow, menuStoreCol.��۲�) = Format((Val(.EditText) - .TextMatrix(intRow, menuStoreCol.ԭ�ɹ���)) * Val(.TextMatrix(intRow, menuStoreCol.����)), mstrMoneyFormat)
                
                If Val(.TextMatrix(intRow, menuStoreCol.���)) = 1 And mblnʱ��ҩƷ�����ε��� = True And mint���� <> 1 Then
                    .TextMatrix(intRow, menuStoreCol.�����ۼ�) = GetFormat(GetFormat(Val(.EditText), mintCostDigit) * (1 + (Val(.TextMatrix(intRow, menuStoreCol.�ӳ���)) / 100)), mintPriceDigit)
                    .TextMatrix(intRow, menuStoreCol.�������) = Format(Val(.TextMatrix(intRow, menuStoreCol.����)) * (Val(.TextMatrix(intRow, menuStoreCol.�����ۼ�)) - Val(.TextMatrix(intRow, menuStoreCol.ԭ���ۼ�))), mstrMoneyFormat)
                End If
                
                dbl��Ʊ��� = (Val(.EditText) - .TextMatrix(intRow, menuStoreCol.ԭ�ɹ���)) * Val(.TextMatrix(intRow, menuStoreCol.����))

                For n = 1 To .rows - 1
                    If .TextMatrix(n, menuStoreCol.ҩƷid) <> "" Then
                        If Val(.TextMatrix(n, menuStoreCol.ҩƷid)) = Val(.TextMatrix(intRow, menuStoreCol.ҩƷid)) And n <> intRow Then
                            If chkCostBatch.Value = 0 Or (Val(.TextMatrix(intRow, menuStoreCol.����)) <> 0 And Val(.TextMatrix(intRow, menuStoreCol.����)) = Val(.TextMatrix(n, menuStoreCol.����))) Then
                                dbl�ֲɹ��� = Val(.EditText)
                                .TextMatrix(n, menuStoreCol.�ֲɹ���) = GetFormat(dbl�ֲɹ���, mintCostDigit)
'                                If dbl�ֲɹ��� <> 0 Then
'                                    .TextMatrix(n, menuStoreCol.�ӳ���) = GetFormat((Val(.TextMatrix(n, menuStoreCol.�����ۼ�)) / dbl�ֲɹ��� - 1) * 100, 5)
'                                End If
                                .TextMatrix(n, menuStoreCol.��۲�) = Format((dbl�ֲɹ��� - .TextMatrix(n, menuStoreCol.ԭ�ɹ���)) * Val(.TextMatrix(n, menuStoreCol.����)), mstrMoneyFormat)
                                
                                If Val(.TextMatrix(intRow, menuStoreCol.���)) = 1 And mblnʱ��ҩƷ�����ε��� = True And mint���� <> 1 Then
                                    .TextMatrix(n, menuStoreCol.�����ۼ�) = GetFormat(GetFormat(dbl�ֲɹ���, mintCostDigit) * (1 + (Val(.TextMatrix(n, menuStoreCol.�ӳ���)) / 100)), mintPriceDigit)
                                    .TextMatrix(n, menuStoreCol.�������) = Format(Val(.TextMatrix(n, menuStoreCol.����)) * (Val(.TextMatrix(n, menuStoreCol.�����ۼ�)) - Val(.TextMatrix(n, menuStoreCol.ԭ���ۼ�))), mstrMoneyFormat)
                                End If
                            Else
                                dbl�ֲɹ��� = Val(.TextMatrix(n, menuStoreCol.�ֲɹ���))
                            End If
                            dbl��Ʊ��� = dbl��Ʊ��� + (dbl�ֲɹ��� - .TextMatrix(n, menuStoreCol.ԭ�ɹ���)) * Val(.TextMatrix(n, menuStoreCol.����))
                        End If
                    End If
                Next

                If chkAutoPay.Value = 1 Then
                    For n = 1 To vsfPay.rows - 1
                        If vsfPay.TextMatrix(1, 0) <> "" Then
                            If Val(vsfPay.TextMatrix(n, menuPayCol.ҩƷid)) = Val(vsfStore.TextMatrix(intRow, menuStoreCol.ҩƷid)) Then
                                vsfPay.TextMatrix(n, menuPayCol.��Ʊ���) = Format(dbl��Ʊ���, mstrMoneyFormat)
                            End If
                        End If
                    Next
                End If

                If chkCostBatch.Value = 0 Then
                    For n = 1 To vsfPrice.rows - 1
                        If Val(.TextMatrix(intRow, menuStoreCol.ҩƷid)) = Val(vsfPrice.TextMatrix(n, menuPriceCol.ҩƷid)) Then
                            vsfPrice.TextMatrix(n, menuPriceCol.�ֳɱ���) = .TextMatrix(intRow, menuStoreCol.�ֲɹ���)
                            Exit For
                        End If
                    Next
                Else
                    CaluateAverCost Val(.TextMatrix(intRow, menuStoreCol.ҩƷid))
                End If
                Call CaculateAverPirce(Val(.TextMatrix(intRow, menuStoreCol.ҩƷid)))  '�۸�䶯������ƽ���ۼ�
        End Select
    End With
End Sub

Private Sub CaculateAverPirce(ByVal lngҩƷID As Long)
    '�Զ�����ƽ���ۼ�
    Dim i As Integer
    Dim dblSumPrice As Double
    Dim dblSumNumber As Double
    
    With vsfStore
        For i = 1 To .rows - 1
            If .TextMatrix(i, menuStoreCol.ҩƷid) <> "" Then
                If Val(.TextMatrix(i, menuStoreCol.ҩƷid)) = lngҩƷID Then
                    dblSumPrice = dblSumPrice + Val(.TextMatrix(i, menuStoreCol.�����ۼ�)) * Val(.TextMatrix(i, menuStoreCol.����))
                    dblSumNumber = dblSumNumber + Val(.TextMatrix(i, menuStoreCol.����))
                End If
            End If
        Next
    End With

    With vsfPrice
        If dblSumNumber > 0 Then
            For i = 1 To .rows - 1
                If .TextMatrix(i, menuPriceCol.ҩƷid) <> "" Then
                    If Val(.TextMatrix(i, menuPriceCol.ҩƷid)) = lngҩƷID Then
                        .TextMatrix(i, menuPriceCol.�����ۼ�) = GetFormat(dblSumPrice / dblSumNumber, mintPriceDigit)
                        Exit For
                    End If
                End If
            Next
        End If
    End With
End Sub

Private Sub CaculateAverOldPirce(ByVal lngҩƷID As Long)
    '�Զ�ԭʼ����ƽ���ۼ�
    Dim i As Integer
    Dim dblSumPrice As Double
    Dim dblSumNumber As Double
    
    With vsfStore
        For i = 1 To .rows - 1
            If .TextMatrix(i, menuStoreCol.ҩƷid) <> "" Then
                If Val(.TextMatrix(i, menuStoreCol.ҩƷid)) = lngҩƷID Then
                    dblSumPrice = dblSumPrice + Val(.TextMatrix(i, menuStoreCol.ԭ���ۼ�)) * Val(.TextMatrix(i, menuStoreCol.����))
                    dblSumNumber = dblSumNumber + Val(.TextMatrix(i, menuStoreCol.����))
                End If
            End If
        Next
    End With

    With vsfPrice
        If dblSumNumber > 0 Then
            For i = 1 To .rows - 1
                If .TextMatrix(i, menuPriceCol.ҩƷid) <> "" Then
                    If Val(.TextMatrix(i, menuPriceCol.ҩƷid)) = lngҩƷID Then
                        .TextMatrix(i, menuPriceCol.ԭ���ۼ�) = GetFormat(dblSumPrice / dblSumNumber, mintPriceDigit)
                        Exit For
                    End If
                End If
            Next
        End If
    End With
End Sub




