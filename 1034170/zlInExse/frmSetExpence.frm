VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSetExpence 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7485
   ControlBox      =   0   'False
   Icon            =   "frmSetExpence.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin TabDlg.SSTab stab 
      Height          =   6165
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   10874
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   3
      TabsPerRow      =   5
      TabHeight       =   520
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "���ʲ���"
      TabPicture(0)   =   "frmSetExpence.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtת��"
      Tab(0).Control(1)=   "txtOutDay0"
      Tab(0).Control(2)=   "fraDoctor"
      Tab(0).Control(3)=   "cboSendMateria"
      Tab(0).Control(4)=   "fra����ҩƷ��λ"
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(6)=   "lst�շ����"
      Tab(0).Control(7)=   "fraPrint"
      Tab(0).Control(8)=   "UDOutDay(0)"
      Tab(0).Control(9)=   "chkת��"
      Tab(0).Control(10)=   "fraҩ��"
      Tab(0).Control(11)=   "fra����"
      Tab(0).Control(12)=   "lblOutDate(0)"
      Tab(0).Control(13)=   "lbl��ҩ"
      Tab(0).Control(14)=   "Label1"
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "���ʲ���(&1)"
      TabPicture(1)   =   "frmSetExpence.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkRefundStyle"
      Tab(1).Control(1)=   "fraMzDepositDefaultUse"
      Tab(1).Control(2)=   "fra�ɿ����"
      Tab(1).Control(3)=   "fra��Ѫ���"
      Tab(1).Control(4)=   "chk��������"
      Tab(1).Control(5)=   "chk���ʲ���"
      Tab(1).Control(6)=   "chk(10)"
      Tab(1).Control(7)=   "cbo���տ���"
      Tab(1).Control(8)=   "chk(16)"
      Tab(1).Control(9)=   "chk(15)"
      Tab(1).Control(10)=   "chk(14)"
      Tab(1).Control(11)=   "fraFeeDate"
      Tab(1).Control(12)=   "chk(12)"
      Tab(1).Control(13)=   "UDOutDay(1)"
      Tab(1).Control(14)=   "txtOutDay1"
      Tab(1).Control(15)=   "chk(11)"
      Tab(1).Control(16)=   "Label2"
      Tab(1).Control(17)=   "lblOutDate(1)"
      Tab(1).ControlCount=   18
      TabCaption(2)   =   "���㷽ʽ˳������(&2)"
      TabPicture(2)   =   "frmSetExpence.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fra"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "����Ʊ�ݿ���(&3)"
      TabPicture(3)   =   "frmSetExpence.frx":0060
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "lbl�˿��վ�"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "lblListPrint"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "lblUnit"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "lblOutUse"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "lblInUse"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "fraƱ�ݸ�ʽ"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "cmd�˿��վ�"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "chk(13)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "cmdListPrintSet"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "cbo������ϸ"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "cmdPrintSetup"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "fraTitle"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "cbo�˿��վ�"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "cboʹ�����"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "cmdBillMZ"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "cboInvoiceKindMZ"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "cboInvoiceKindZY"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).Control(17)=   "cmdBillZY"
      Tab(3).Control(17).Enabled=   0   'False
      Tab(3).ControlCount=   18
      TabCaption(4)   =   "����Ʊ�ݿ���(&4)"
      TabPicture(4)   =   "frmSetExpence.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdRed"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "fraRed"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "fraDepositPrint"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "fraDeposit"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).ControlCount=   4
      Begin VB.CommandButton cmdBillZY 
         Caption         =   "����Ʊ������(&P)"
         Height          =   350
         Left            =   5085
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   2505
         Width           =   1560
      End
      Begin VB.ComboBox cboInvoiceKindZY 
         Height          =   300
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   107
         Top             =   2535
         Width           =   3270
      End
      Begin VB.ComboBox cboInvoiceKindMZ 
         Height          =   300
         Left            =   1740
         Style           =   2  'Dropdown List
         TabIndex        =   106
         Top             =   2160
         Width           =   3270
      End
      Begin VB.CommandButton cmdBillMZ 
         Caption         =   "����Ʊ������(&P)"
         Height          =   350
         Left            =   5085
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   2130
         Width           =   1560
      End
      Begin VB.CommandButton cmdRed 
         Caption         =   "���ʺ�Ʊ����(&P)"
         Height          =   350
         Left            =   -74355
         TabIndex        =   104
         Top             =   5640
         Width           =   1560
      End
      Begin VB.Frame fraRed 
         Caption         =   "���ʺ�Ʊ��ʽ"
         Height          =   1995
         Left            =   -74880
         TabIndex        =   102
         Top             =   3540
         Width           =   6870
         Begin VSFlex8Ctl.VSFlexGrid vsRedFormat 
            Height          =   1605
            Left            =   60
            TabIndex        =   103
            Top             =   285
            Width           =   6705
            _cx             =   11827
            _cy             =   2831
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
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   8421504
            GridColorFixed  =   8421504
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmSetExpence.frx":0098
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
            ExplorerBar     =   2
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
      End
      Begin VB.Frame fraDepositPrint 
         Caption         =   "Ԥ��Ʊ�ݴ�ӡ��ʽ"
         Height          =   765
         Left            =   -74880
         TabIndex        =   97
         Top             =   2640
         Width           =   6855
         Begin VB.CommandButton cmdDeposit 
            Caption         =   "Ԥ��Ʊ�ݴ�ӡ����"
            Height          =   350
            Left            =   4710
            TabIndex        =   101
            Top             =   282
            Width           =   1860
         End
         Begin VB.OptionButton optBalanceDepositPrint 
            Caption         =   "ѡ���Ƿ��ӡ"
            Height          =   255
            Index           =   2
            Left            =   2985
            TabIndex        =   100
            Top             =   330
            Width           =   1395
         End
         Begin VB.OptionButton optBalanceDepositPrint 
            Caption         =   "�Զ���ӡ"
            Height          =   255
            Index           =   1
            Left            =   1695
            TabIndex        =   99
            Top             =   330
            Width           =   1335
         End
         Begin VB.OptionButton optBalanceDepositPrint 
            Caption         =   "����ӡ"
            Height          =   255
            Index           =   0
            Left            =   495
            TabIndex        =   98
            Top             =   330
            Width           =   1110
         End
      End
      Begin VB.Frame fraDeposit 
         Caption         =   "���ع���Ԥ��Ʊ��"
         Height          =   1935
         Left            =   -74880
         TabIndex        =   95
         Top             =   555
         Width           =   6855
         Begin VSFlex8Ctl.VSFlexGrid vsDeposit 
            Height          =   1560
            Left            =   75
            TabIndex        =   96
            Top             =   255
            Width           =   6705
            _cx             =   11827
            _cy             =   2752
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
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   8421504
            GridColorFixed  =   8421504
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmSetExpence.frx":012A
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
            ExplorerBar     =   2
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
      End
      Begin VB.CheckBox chkRefundStyle 
         Caption         =   "�����˿�ȱʡ��Ԥ���ɿ���㷽ʽ�˿�"
         Height          =   255
         Left            =   -74700
         TabIndex        =   94
         Top             =   3255
         Width           =   3525
      End
      Begin VB.Frame fraMzDepositDefaultUse 
         Caption         =   "����Ԥ��ȱʡʹ�÷�ʽ"
         Height          =   900
         Left            =   -74748
         TabIndex        =   90
         Top             =   4980
         Width           =   6636
         Begin VB.OptionButton optMzDeposit 
            Caption         =   "ʹ��ʣ������Ԥ����"
            Height          =   564
            Index           =   2
            Left            =   4476
            TabIndex        =   93
            Top             =   288
            Value           =   -1  'True
            Width           =   2028
         End
         Begin VB.OptionButton optMzDeposit 
            Caption         =   "�����ʽ��ʹ��Ԥ��"
            Height          =   564
            Index           =   1
            Left            =   2136
            TabIndex        =   92
            Top             =   288
            Width           =   2256
         End
         Begin VB.OptionButton optMzDeposit 
            Caption         =   "��ʹ��Ԥ����"
            Height          =   300
            Index           =   0
            Left            =   252
            TabIndex        =   91
            Top             =   420
            Width           =   1524
         End
      End
      Begin VB.Frame fra�ɿ���� 
         Caption         =   "���ʽɿ����"
         Height          =   780
         Left            =   -74760
         TabIndex        =   87
         Top             =   4020
         Width           =   6645
         Begin VB.OptionButton opt�ɿ� 
            Caption         =   "������ȡ�ֽ�ʱ,��������ɿ�"
            Height          =   315
            Index           =   1
            Left            =   3000
            TabIndex        =   89
            Top             =   315
            Width           =   2835
         End
         Begin VB.OptionButton opt�ɿ� 
            Caption         =   "�����нɿ����"
            Height          =   315
            Index           =   0
            Left            =   885
            TabIndex        =   88
            Top             =   315
            Value           =   -1  'True
            Width           =   1770
         End
      End
      Begin VB.ComboBox cboʹ����� 
         Height          =   300
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   86
         Top             =   4995
         Width           =   2205
      End
      Begin VB.ComboBox cbo�˿��վ� 
         Height          =   300
         Left            =   4980
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   4650
         Width           =   2040
      End
      Begin VB.Frame fraTitle 
         Caption         =   "���ع����շ�Ʊ��"
         Height          =   1620
         Left            =   135
         TabIndex        =   83
         Top             =   480
         Width           =   6855
         Begin VSFlex8Ctl.VSFlexGrid vsBill 
            Height          =   1290
            Left            =   75
            TabIndex        =   84
            Top             =   255
            Width           =   6705
            _cx             =   11827
            _cy             =   2275
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
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   8421504
            GridColorFixed  =   8421504
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmSetExpence.frx":0208
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
            ExplorerBar     =   2
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
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "�ص�Ʊ�ݴ�ӡ����(&4)"
         Height          =   350
         Left            =   4695
         TabIndex        =   80
         Top             =   5700
         Width           =   1860
      End
      Begin VB.ComboBox cbo������ϸ 
         ForeColor       =   &H80000012&
         Height          =   300
         Left            =   1845
         Style           =   2  'Dropdown List
         TabIndex        =   79
         Top             =   5340
         Width           =   1980
      End
      Begin VB.CommandButton cmdListPrintSet 
         Caption         =   "��ӡ������ϸ����"
         Height          =   315
         Left            =   2475
         TabIndex        =   78
         Top             =   5715
         Width           =   1635
      End
      Begin VB.CheckBox chk 
         Caption         =   "��Լ��λ����ÿλ���˷ֱ��ӡƱ��"
         Height          =   225
         Index           =   13
         Left            =   165
         TabIndex        =   68
         Top             =   4695
         Width           =   3300
      End
      Begin VB.CommandButton cmd�˿��վ� 
         Caption         =   "�˿��վ�����(&S)"
         Height          =   350
         Left            =   240
         TabIndex        =   71
         Top             =   5700
         Width           =   1620
      End
      Begin VB.Frame fra��Ѫ��� 
         Caption         =   "����ʱ��Ѫ�Ѽ��"
         Height          =   1110
         Left            =   -70995
         TabIndex        =   72
         Top             =   2730
         Width           =   2880
         Begin VB.OptionButton opt��Ѫ 
            Caption         =   "��鲢��ʾ"
            Height          =   210
            Index           =   1
            Left            =   390
            TabIndex        =   76
            Top             =   705
            Width           =   1305
         End
         Begin VB.OptionButton opt��Ѫ 
            Caption         =   "�����"
            Height          =   210
            Index           =   0
            Left            =   405
            TabIndex        =   74
            Top             =   435
            Value           =   -1  'True
            Width           =   945
         End
      End
      Begin VB.Frame fra 
         Height          =   5430
         Left            =   -74490
         TabIndex        =   64
         Top             =   540
         Width           =   5880
         Begin VB.CommandButton cmdUp 
            Caption         =   "��"
            Height          =   510
            Left            =   5280
            TabIndex        =   66
            Top             =   1140
            Width           =   375
         End
         Begin VB.CommandButton cmdDown 
            Caption         =   "��"
            Height          =   510
            Left            =   5280
            TabIndex        =   65
            Top             =   1740
            Width           =   375
         End
         Begin VSFlex8Ctl.VSFlexGrid vsBalanceSort 
            Height          =   5055
            Left            =   135
            TabIndex        =   67
            Top             =   240
            Width           =   4995
            _cx             =   8811
            _cy             =   8916
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
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483634
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483632
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   2
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   7
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   300
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmSetExpence.frx":02E6
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
            ExplorerBar     =   8
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
      End
      Begin VB.CheckBox chk�������� 
         Caption         =   "���ʼ�鲡���������"
         Height          =   255
         Left            =   -74700
         TabIndex        =   39
         Top             =   2955
         Width           =   2190
      End
      Begin VB.CheckBox chk���ʲ��� 
         Caption         =   "���ʺ����������Ϣ"
         Height          =   225
         Left            =   -74700
         TabIndex        =   37
         Top             =   2655
         Width           =   2175
      End
      Begin VB.TextBox txtת�� 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   -70695
         MaxLength       =   2
         TabIndex        =   11
         Text            =   "3"
         Top             =   3090
         Width           =   255
      End
      Begin VB.CheckBox chk 
         Caption         =   "��;����ȱʡ��Ԥ����"
         Height          =   195
         Index           =   10
         Left            =   -74700
         TabIndex        =   33
         Top             =   1155
         Width           =   2160
      End
      Begin VB.ComboBox cbo���տ��� 
         ForeColor       =   &H80000012&
         Height          =   300
         Left            =   -72750
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   3600
         Width           =   1515
      End
      Begin VB.TextBox txtOutDay0 
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   -73965
         MaxLength       =   3
         TabIndex        =   9
         Text            =   "0"
         ToolTipText     =   "����Ϊ 0 ��ʾֻ��ѡ����Ժ����"
         Top             =   3090
         Width           =   450
      End
      Begin VB.CheckBox chk 
         Caption         =   "�ж��סԺ���õĲ����Զ�������������"
         Height          =   195
         Index           =   16
         Left            =   -74700
         TabIndex        =   38
         Top             =   2355
         Width           =   3720
      End
      Begin VB.CheckBox chk 
         Caption         =   "��ʹ��ָ��סԺ������Ԥ����"
         Height          =   195
         Index           =   15
         Left            =   -74700
         TabIndex        =   36
         Top             =   2070
         Width           =   2760
      End
      Begin VB.Frame fraDoctor 
         Caption         =   "��ʾ������"
         Height          =   1170
         Left            =   -71880
         TabIndex        =   60
         Top             =   480
         Width           =   1755
         Begin VB.OptionButton optDoctorKind 
            Caption         =   "������"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   210
            TabIndex        =   19
            Top             =   435
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton optDoctorKind 
            Caption         =   "������"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   210
            TabIndex        =   20
            Top             =   735
            Width           =   1020
         End
      End
      Begin VB.ComboBox cboSendMateria 
         ForeColor       =   &H80000012&
         Height          =   300
         Left            =   -73965
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3450
         Width           =   2010
      End
      Begin VB.CheckBox chk 
         Caption         =   "LED��ʾ��ӭ��Ϣ"
         Height          =   225
         Index           =   14
         Left            =   -70980
         TabIndex        =   40
         ToolTipText     =   "�շѴ������벡�˺�,�Ƿ���ʾ��ӭ��Ϣ������"
         Top             =   1095
         Value           =   1  'Checked
         Width           =   1770
      End
      Begin VB.Frame fraFeeDate 
         Caption         =   "���ʷ����ڼ�����"
         ForeColor       =   &H00000000&
         Height          =   1095
         Left            =   -70980
         TabIndex        =   56
         Top             =   1440
         Width           =   2865
         Begin VB.OptionButton optTime 
            Caption         =   "���Ǽ�ʱ��"
            Height          =   195
            Index           =   0
            Left            =   390
            TabIndex        =   41
            Top             =   360
            Value           =   -1  'True
            Width           =   1320
         End
         Begin VB.OptionButton optTime 
            Caption         =   "������ʱ��"
            Height          =   195
            Index           =   1
            Left            =   390
            TabIndex        =   42
            Top             =   720
            Width           =   1320
         End
      End
      Begin VB.Frame fra����ҩƷ��λ 
         Caption         =   " ҩƷ��λ "
         Height          =   1140
         Left            =   -71880
         TabIndex        =   27
         Top             =   1845
         Width           =   1785
         Begin VB.OptionButton opt����ҩƷ��λ 
            Caption         =   "סԺ��λ"
            Height          =   180
            Index           =   1
            Left            =   195
            TabIndex        =   22
            Top             =   705
            Width           =   1020
         End
         Begin VB.OptionButton opt����ҩƷ��λ 
            Caption         =   "�ۼ۵�λ"
            Height          =   180
            Index           =   0
            Left            =   195
            TabIndex        =   21
            Top             =   405
            Value           =   -1  'True
            Width           =   1020
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "�Բ��˵�����ý��н���"
         Height          =   195
         Index           =   12
         Left            =   -74700
         TabIndex        =   34
         Top             =   1755
         Width           =   2280
      End
      Begin VB.Frame Frame1 
         Height          =   2655
         Left            =   -74760
         TabIndex        =   1
         Top             =   345
         Width           =   2775
         Begin VB.CheckBox chk 
            Caption         =   "Ƿ��ʱ������Ϊ���۵�"
            Height          =   195
            Index           =   6
            Left            =   300
            TabIndex        =   8
            Top             =   2115
            Width           =   2400
         End
         Begin VB.CheckBox chk 
            Caption         =   "�����˶���������"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   300
            TabIndex        =   5
            Top             =   1275
            Width           =   1740
         End
         Begin VB.CheckBox chk 
            Caption         =   "��ҩ�������븶��"
            Height          =   195
            Index           =   0
            Left            =   300
            TabIndex        =   2
            Top             =   435
            Value           =   1  'Checked
            Width           =   1740
         End
         Begin VB.CheckBox chk 
            Caption         =   "�������а�����ʿ"
            Height          =   195
            Index           =   2
            Left            =   300
            TabIndex        =   4
            Top             =   990
            Width           =   1740
         End
         Begin VB.CheckBox chk 
            Caption         =   "���������������"
            Height          =   195
            Index           =   1
            Left            =   300
            TabIndex        =   3
            Top             =   720
            Width           =   1740
         End
         Begin VB.CheckBox chk 
            Caption         =   "�������۲��˼���"
            Height          =   195
            Index           =   4
            Left            =   300
            TabIndex        =   6
            Top             =   1560
            Width           =   1740
         End
         Begin VB.CheckBox chk 
            Caption         =   "סԺ���۲��˼���"
            Height          =   195
            Index           =   5
            Left            =   300
            TabIndex        =   7
            Top             =   1830
            Width           =   1740
         End
      End
      Begin MSComCtl2.UpDown UDOutDay 
         Height          =   270
         Index           =   1
         Left            =   -73410
         TabIndex        =   46
         Top             =   630
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   476
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtOutDay1"
         BuddyDispid     =   196648
         OrigLeft        =   1486
         OrigTop         =   3375
         OrigRight       =   1726
         OrigBottom      =   3645
         Max             =   999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtOutDay1 
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   -73860
         MaxLength       =   3
         TabIndex        =   32
         Text            =   "0"
         ToolTipText     =   "����Ϊ 0 ��ʾֻ��ѡ����Ժ����"
         Top             =   645
         Width           =   450
      End
      Begin VB.ListBox lst�շ���� 
         Height          =   2160
         Left            =   -69930
         Style           =   1  'Checkbox
         TabIndex        =   23
         ToolTipText     =   "�븴ѡ����ʹ�õ��շ����"
         Top             =   690
         Width           =   1875
      End
      Begin VB.Frame fraPrint 
         Caption         =   " ��ӡ����"
         Height          =   1515
         Left            =   -69870
         TabIndex        =   45
         Top             =   3945
         Width           =   1845
         Begin VB.CheckBox chkBillPrint 
            Caption         =   "���"
            Height          =   195
            Index           =   2
            Left            =   540
            TabIndex        =   26
            Top             =   1080
            Width           =   660
         End
         Begin VB.CheckBox chkBillPrint 
            Caption         =   "����"
            Height          =   195
            Index           =   1
            Left            =   540
            TabIndex        =   25
            Top             =   720
            Width           =   660
         End
         Begin VB.CheckBox chkBillPrint 
            Caption         =   "����"
            Height          =   195
            Index           =   0
            Left            =   540
            TabIndex        =   24
            Top             =   360
            Width           =   660
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "���˳�Ժ���ʺ��Զ���Ժ"
         Height          =   195
         Index           =   11
         Left            =   -74700
         TabIndex        =   35
         Top             =   1455
         Width           =   2280
      End
      Begin MSComCtl2.UpDown UDOutDay 
         Height          =   270
         Index           =   0
         Left            =   -73515
         TabIndex        =   61
         Top             =   3090
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   476
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtOutDay0"
         BuddyDispid     =   196639
         OrigLeft        =   1486
         OrigTop         =   2760
         OrigRight       =   1726
         OrigBottom      =   3030
         Max             =   999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.CheckBox chkת�� 
         Caption         =   "��ʾ���   ���ת������"
         Height          =   195
         Left            =   -71685
         TabIndex        =   10
         Top             =   3120
         Width           =   2370
      End
      Begin VB.Frame fraҩ�� 
         Caption         =   " ҩ���뷢�ϲ������� "
         Height          =   1515
         Left            =   -74760
         TabIndex        =   28
         Top             =   3945
         Width           =   4725
         Begin VB.ComboBox cbo���� 
            Height          =   300
            Left            =   2880
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   720
            Width           =   1305
         End
         Begin VB.CheckBox chk 
            Caption         =   "��ʾ����ҩ�����"
            Height          =   195
            Index           =   8
            Left            =   2400
            TabIndex        =   18
            Top             =   1150
            Width           =   1845
         End
         Begin VB.CheckBox chk 
            Caption         =   "��ʾ����ҩ����"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   17
            Top             =   1150
            Width           =   1850
         End
         Begin VB.ComboBox cbo��ҩ 
            Height          =   300
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   720
            Width           =   1305
         End
         Begin VB.ComboBox cbo��ҩ 
            Height          =   300
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   360
            Width           =   1305
         End
         Begin VB.ComboBox cbo��ҩ 
            Height          =   300
            Left            =   2880
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   360
            Width           =   1305
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���ϲ���"
            Height          =   180
            Left            =   2100
            TabIndex        =   57
            Top             =   780
            Width           =   720
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�в�ҩ"
            Height          =   180
            Left            =   120
            TabIndex        =   50
            Top             =   780
            Width           =   540
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ҩ"
            Height          =   180
            Left            =   120
            TabIndex        =   49
            Top             =   420
            Width           =   540
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�г�ҩ"
            Height          =   180
            Left            =   2280
            TabIndex        =   48
            Top             =   420
            Width           =   540
         End
      End
      Begin VB.Frame fra���� 
         Caption         =   " ���п�����ҩ�� "
         ForeColor       =   &H00C00000&
         Height          =   1515
         Left            =   -74760
         TabIndex        =   29
         Top             =   3960
         Visible         =   0   'False
         Width           =   4740
         Begin VB.ListBox lst��ҩ�� 
            Height          =   480
            Left            =   90
            Style           =   1  'Checkbox
            TabIndex        =   30
            Top             =   480
            Width           =   1350
         End
         Begin VB.ListBox lst��ҩ�� 
            Height          =   480
            Left            =   1485
            Style           =   1  'Checkbox
            TabIndex        =   31
            Top             =   480
            Width           =   1350
         End
         Begin VB.ListBox lst��ҩ�� 
            Height          =   480
            Left            =   2880
            Style           =   1  'Checkbox
            TabIndex        =   44
            Top             =   480
            Width           =   1350
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ҩ��"
            Height          =   180
            Left            =   90
            TabIndex        =   55
            Top             =   250
            Width           =   540
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ҩ��"
            Height          =   180
            Left            =   1485
            TabIndex        =   54
            Top             =   250
            Width           =   540
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ҩ��"
            Height          =   180
            Left            =   2880
            TabIndex        =   53
            Top             =   250
            Width           =   540
         End
      End
      Begin VB.Frame fraƱ�ݸ�ʽ 
         Caption         =   "�շ�Ʊ�ݸ�ʽ"
         Height          =   1725
         Left            =   135
         TabIndex        =   81
         Top             =   2880
         Width           =   6870
         Begin VSFlex8Ctl.VSFlexGrid vsBillFormat 
            Height          =   1365
            Left            =   60
            TabIndex        =   82
            Top             =   285
            Width           =   6705
            _cx             =   11827
            _cy             =   2408
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
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   8421504
            GridColorFixed  =   8421504
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmSetExpence.frx":03E0
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
            ExplorerBar     =   2
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
      End
      Begin VB.Label lblInUse 
         AutoSize        =   -1  'True
         Caption         =   "סԺ����Ʊ��ʹ��"
         Height          =   180
         Left            =   180
         TabIndex        =   110
         Top             =   2595
         Width           =   1440
      End
      Begin VB.Label lblOutUse 
         AutoSize        =   -1  'True
         Caption         =   "�������Ʊ��ʹ��"
         Height          =   180
         Left            =   180
         TabIndex        =   109
         Top             =   2220
         Width           =   1440
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         Caption         =   "��Լ��λ����ʹ��                         ��Ʊ��"
         Height          =   180
         Left            =   165
         TabIndex        =   85
         Top             =   5055
         Width           =   4230
      End
      Begin VB.Label lblListPrint 
         Caption         =   "���ʺ��ӡ������ϸ"
         Height          =   225
         Left            =   165
         TabIndex        =   77
         Top             =   5385
         Width           =   1665
      End
      Begin VB.Label lbl�˿��վ� 
         AutoSize        =   -1  'True
         Caption         =   "�����˿��վݴ�ӡ"
         Height          =   180
         Left            =   3495
         TabIndex        =   69
         Top             =   4710
         Width           =   1440
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "��Ժ���ʼ����տ���"
         Height          =   180
         Left            =   -74670
         TabIndex        =   63
         Top             =   3660
         Width           =   1800
      End
      Begin VB.Label lblOutDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ѡ��         ���ڳ�Ժ�Ĳ���"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   -74760
         TabIndex        =   62
         Top             =   3135
         Width           =   2790
      End
      Begin VB.Label lbl��ҩ 
         Caption         =   "����֮��"
         Height          =   255
         Left            =   -74760
         TabIndex        =   58
         Top             =   3495
         Width           =   735
      End
      Begin VB.Label lblOutDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ѡ��         ���ڳ�Ժ�Ĳ���"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   -74655
         TabIndex        =   52
         Top             =   705
         Width           =   2790
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������:"
         Height          =   180
         Left            =   -69960
         TabIndex        =   51
         Top             =   420
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdDeviceSetup 
      Caption         =   "�豸����(&S)"
      Height          =   350
      Left            =   1320
      TabIndex        =   59
      Top             =   6405
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4980
      TabIndex        =   73
      Top             =   6405
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6180
      TabIndex        =   75
      Top             =   6405
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   120
      TabIndex        =   47
      Top             =   6405
      Width           =   1100
   End
End
Attribute VB_Name = "frmSetExpence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Public mbytInFun As Byte '0=����,1=����
Public mbytUseType As Byte '0:��ͨ����,1-���ҷ�ɢ����,2-ҽ�����Ҽ���
Public mstrPrivs As String
Public mlngModul As Long
Public mblnOnlyDrugStock As Boolean  '����ʾҩ������
Private Enum chkBPS
    C0���� = 0
    C1���� = 1
    C2��� = 2
End Enum
Private Enum chks
    C00��ҩ�丶�� = 0
    C01�������� = 1
    C02�����˺���ʿ = 2
    C03�����˶����� = 3
    C04�������ۼ��� = 4
    C05סԺ���ۼ��� = 5
    C06Ƿ�Ѵ滮�۵� = 6
    C07����ҩ���� = 7
    C08����ҩ����� = 8
    C09ҽ�����ʲ��� = 9
    C10��;������Ԥ�� = 10
    C11�����Զ���Ժ = 11
    C12����ÿɽ��� = 12
    C13��Լ��λ�����˴�ӡ = 13
    C14LED��ӭ��Ϣ = 14
    C15����ָ��Ԥ���� = 15
    C16���סԺ������������ = 16
End Enum
Private Enum InvoiceKind
    C1�շ��վ� = 1
    C3�����վ� = 3
    C4�����վ� = 10
End Enum
Private Const CModule As Long = 1150    'סԺ���ʲ���
Private Sub zlOnlyDrugStrock()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ʾҩ�����������
    '����:���˺�
    '����:2010-01-25 15:24:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim ctl As Control
    Err = 0: On Error GoTo ErrHand:
    If mblnOnlyDrugStock And mbytInFun = 0 Then
        For Each ctl In Me.Controls
           Select Case UCase(TypeName(ctl))
           Case UCase("ImageList")
           Case UCase("sstab")
                ctl.Visible = True
           Case Else
                If ctl Is fra���� Or ctl Is fraҩ�� Or ctl.Container Is fraҩ�� Or ctl.Container Is fra���� Or ctl Is cmdOK Or ctl Is cmdCancel Then
                    ctl.Visible = True
                Else
                     ctl.Visible = False
                End If
           End Select
        Next
        fraҩ��.Top = Frame1.Top + 200
        fra����.Top = fraҩ��.Top
        
        Me.Height = 3525: Me.Width = 5470
        cmdCancel.Top = ScaleHeight - cmdCancel.Height - 100
        cmdCancel.Left = ScaleWidth - cmdCancel.Width - 100
        cmdOK.Top = cmdCancel.Top
        cmdOK.Left = cmdCancel.Left - cmdOK.Width - 100
        
        stab.Height = cmdOK.Top - stab.Top - 100
        stab.Width = ScaleWidth - stab.Left * 2
        stab.TabCaption(0) = "ҩ������"
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

'����:27380
Private Sub chkת��_Click()
    txtת��.Enabled = chkת��.Value = 1
    If txtת��.Visible And txtת��.Enabled Then txtת��.SetFocus
End Sub
Private Sub chkת��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdDeposit_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me)
End Sub

Private Sub cmdListPrintSet_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1137_3", Me)
End Sub

Private Sub cmdPrintSetup_Click()
     Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1137_4", Me)
End Sub
Private Sub cmd�˿��վ�_Click()
    '���˺� ����:27776 ����:2010-02-04 16:44:39
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1137_3", Me)
End Sub

Private Sub txtת��_GotFocus()
   zlControl.TxtSelAll txtת��
End Sub

Private Sub txtת��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub


Private Sub cboInvoiceKindZY_Click()
    Dim bytKind As Byte
    If Visible Then '����ʱǿ�Ƶ���
        If cboInvoiceKindZY.ListIndex = 0 And cboInvoiceKindMZ.ListIndex = 0 Then
            bytKind = InvoiceKind.C3�����վ�
        ElseIf cboInvoiceKindZY.ListIndex = 1 And cboInvoiceKindMZ.ListIndex = 1 Then
            bytKind = InvoiceKind.C1�շ��վ�
        Else
            bytKind = InvoiceKind.C4�����վ�
        End If
        Call InitShareInvoice(bytKind)
        Call InitDepositInvoice
    End If
End Sub

Private Sub cboInvoiceKindMZ_Click()
    Dim bytKind As Byte
    If Visible Then '����ʱǿ�Ƶ���
        If cboInvoiceKindZY.ListIndex = 0 And cboInvoiceKindMZ.ListIndex = 0 Then
            bytKind = InvoiceKind.C3�����վ�
        ElseIf cboInvoiceKindZY.ListIndex = 1 And cboInvoiceKindMZ.ListIndex = 1 Then
            bytKind = InvoiceKind.C1�շ��վ�
        Else
            bytKind = InvoiceKind.C4�����վ�
        End If
        Call InitShareInvoice(bytKind)
        Call InitDepositInvoice
    End If
End Sub

Private Sub cmdBillZY_Click()
    If gblnBillPrint Then
        Call gobjBillPrint.zlConfigure
    Else
        Call ReportPrintSet(gcnOracle, glngSys, IIf(cboInvoiceKindZY.ListIndex = 0, "ZL" & glngSys \ 100 & "_BILL_1137", "ZL" & glngSys \ 100 & "_BILL_1137_2"), Me)
    End If
End Sub

Private Sub cmdBillMZ_Click()
    If gblnBillPrint Then
        Call gobjBillPrint.zlConfigure
    Else
        Call ReportPrintSet(gcnOracle, glngSys, IIf(cboInvoiceKindMZ.ListIndex = 0, "ZL" & glngSys \ 100 & "_BILL_1137", "ZL" & glngSys \ 100 & "_BILL_1137_2"), Me)
    End If
End Sub

Private Sub cmdRed_Click()
    Call ReportPrintSet(gcnOracle, glngSys, IIf(cboInvoiceKindZY.ListIndex = 0, "ZL" & glngSys \ 100 & "_BILL_1137_5", "ZL" & glngSys \ 100 & "_BILL_1137_6"), Me)
End Sub

Private Sub cmdCancel_Click()
    mblnOnlyDrugStock = False
    Unload Me
End Sub

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, 100, 1137)
End Sub

Private Sub cmdHelp_Click()
    Select Case stab.Tab
        Case 0
            ShowHelp App.ProductName, Me.hWnd, "frmSetExpence1"
        Case 1
            ShowHelp App.ProductName, Me.hWnd, "frmSetExpence2"
    End Select
End Sub

Private Sub cmdOK_Click()
    Dim strValue As String, i As Long, lngShareID As Long
    Dim blnHavePrivs As Boolean, strTemp As String
    
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    
    If mbytInFun = 0 And cbo��ҩ.Visible Then
        If cbo��ҩ.ListIndex = -1 And cbo��ҩ.ListCount > 0 And cbo��ҩ.Enabled Then
            MsgBox "��ѡ����ҩ��.", vbInformation, gstrSysName
            stab.Tab = 0: cbo��ҩ.SetFocus: Exit Sub
        End If
        If cbo��ҩ.ListIndex = -1 And cbo��ҩ.ListCount > 0 And cbo��ҩ.Enabled Then
            MsgBox "��ѡ���ҩ��.", vbInformation, gstrSysName
            stab.Tab = 0: cbo��ҩ.SetFocus: Exit Sub
        End If
        If cbo��ҩ.ListIndex = -1 And cbo��ҩ.ListCount > 0 And cbo��ҩ.Enabled Then
            MsgBox "��ѡ����ҩ��.", vbInformation, gstrSysName
            stab.Tab = 0: cbo��ҩ.SetFocus: Exit Sub
        End If
        If cbo����.ListIndex = -1 And cbo����.ListCount > 0 And cbo����.Enabled Then
            MsgBox "��ѡ�����ķ��ϲ���.", vbInformation, gstrSysName
            stab.Tab = 0: cbo����.SetFocus: Exit Sub
        End If
    End If
    '�������ע����Ϣ
    '����ʹ���������ۼ���ʱ,����������ʾ��������Ƿ����������ü��ʿ���
    If mbytInFun = 0 And (mbytUseType = 0 Or mbytUseType = 1) And chk(chks.C04�������ۼ���).Value = 0 Then
        If Not CheckUnits Then
            MsgBox "����ʹ���������ۼ���ʱ,��û�п��Լ��ʵĿ���,�����޷������ã�", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    On Error Resume Next
    If mbytInFun = 0 Then
    
        'ҩ��
        zlDatabase.SetPara "ȱʡ��ҩ��", IIf(cbo��ҩ.ListIndex = 0, "0", cbo��ҩ.ItemData(cbo��ҩ.ListIndex)), glngSys, CModule, blnHavePrivs
        zlDatabase.SetPara "ȱʡ��ҩ��", IIf(cbo��ҩ.ListIndex = 0, "0", cbo��ҩ.ItemData(cbo��ҩ.ListIndex)), glngSys, CModule, blnHavePrivs
        zlDatabase.SetPara "ȱʡ��ҩ��", IIf(cbo��ҩ.ListIndex = 0, "0", cbo��ҩ.ItemData(cbo��ҩ.ListIndex)), glngSys, CModule, blnHavePrivs
        zlDatabase.SetPara "ȱʡ���ϲ���", IIf(cbo����.ListIndex = 0, "0", cbo����.ItemData(cbo����.ListIndex)), glngSys, CModule, blnHavePrivs
        zlDatabase.SetPara "��ʾ����ҩ�����", chk(chks.C08����ҩ�����).Value, glngSys, CModule, blnHavePrivs
        zlDatabase.SetPara "��ʾ����ҩ����", chk(chks.C07����ҩ����).Value, glngSys, CModule, blnHavePrivs
        '���뷢ҩʱ��ѡ��
        '--------------------------------------------------------------------------
        strValue = ""
        For i = 0 To lst��ҩ��.ListCount - 1
            If lst��ҩ��.Selected(i) Then
                strValue = strValue & "," & lst��ҩ��.ItemData(i)
            End If
        Next
        zlDatabase.SetPara "��ҩ��ѡ��", Mid(strValue, 2), glngSys, CModule
        strValue = ""
        For i = 0 To lst��ҩ��.ListCount - 1
            If lst��ҩ��.Selected(i) Then
                strValue = strValue & "," & lst��ҩ��.ItemData(i)
            End If
        Next
        zlDatabase.SetPara "��ҩ��ѡ��", Mid(strValue, 2), glngSys, CModule
        strValue = ""
        For i = 0 To lst��ҩ��.ListCount - 1
            If lst��ҩ��.Selected(i) Then
                strValue = strValue & "," & lst��ҩ��.ItemData(i)
            End If
        Next
        zlDatabase.SetPara "��ҩ��ѡ��", Mid(strValue, 2), glngSys, CModule, blnHavePrivs
        '--------------------------------------------------------------------------
        If mblnOnlyDrugStock Then GoTo GoOver:
        
        
        zlDatabase.SetPara "���ʴ�ӡ", chkBillPrint(chkBPS.C0����).Value, glngSys, mlngModul, blnHavePrivs  '����1150�Ĳ���
        
        '1150�Ĳ���
        '--------------------------------------------------------------------------------
        '�շ����
        For i = lst�շ����.ListCount - 1 To 0 Step -1
            If lst�շ����.Selected(i) Then strValue = strValue & "'" & Chr(lst�շ����.ItemData(i)) & "',"
        Next
        If strValue <> "" Then strValue = Left(strValue, Len(strValue) - 1)
        zlDatabase.SetPara "�շ����", strValue, glngSys, CModule, blnHavePrivs
    
           
        '���۲��˼���
        zlDatabase.SetPara "�������۲��˼���", chk(chks.C04�������ۼ���).Value, glngSys, CModule, blnHavePrivs
        zlDatabase.SetPara "סԺ���۲��˼���", chk(chks.C05סԺ���ۼ���).Value, glngSys, CModule, blnHavePrivs
        
        zlDatabase.SetPara "��Ժ��������", Val(txtOutDay0.Text), glngSys, CModule, blnHavePrivs
        zlDatabase.SetPara "��������ʾ��ʽ", IIf(optDoctorKind(0).Value, 1, 2), glngSys, CModule, blnHavePrivs
        zlDatabase.SetPara "����ҽ��", IIf(chk(chks.C03�����˶�����).Value = 1, 0, 1), glngSys, CModule, blnHavePrivs
        
        zlDatabase.SetPara "������Ϊ���۵�", chk(chks.C06Ƿ�Ѵ滮�۵�).Value, glngSys, CModule, blnHavePrivs
        zlDatabase.SetPara "��ҩ����", chk(chks.C00��ҩ�丶��).Value, glngSys, CModule, blnHavePrivs
        zlDatabase.SetPara "�������", chk(chks.C01��������).Value, glngSys, CModule, blnHavePrivs
        zlDatabase.SetPara "��ʾ��ʿ", chk(chks.C02�����˺���ʿ).Value, glngSys, CModule, blnHavePrivs
        zlDatabase.SetPara "����ҩƷ��λ", IIf(opt����ҩƷ��λ(0).Value, 0, 1), glngSys, CModule, blnHavePrivs
        If mbytUseType = 0 Then
            zlDatabase.SetPara "���۴�ӡ", chkBillPrint(chkBPS.C1����).Value, glngSys, CModule, blnHavePrivs
            zlDatabase.SetPara "��˴�ӡ", chkBillPrint(chkBPS.C2���).Value, glngSys, CModule, blnHavePrivs
            zlDatabase.SetPara "���ʺ�ҩ", cboSendMateria.ListIndex, glngSys, CModule, blnHavePrivs
        ElseIf mbytUseType = 1 Then
            '���˺� ����:27380 ����:2010-01-22 14:45:32
            zlDatabase.SetPara "���ת������", IIf(chkת��.Value = 1, "1", "0") & "|" & Val(txtת��.Text), glngSys, mlngModul, blnHavePrivs
        End If
    Else
        '���ع��ý���Ʊ��
        zlDatabase.SetPara "סԺ����Ʊ������", cboInvoiceKindZY.ListIndex, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "�������Ʊ������", cboInvoiceKindMZ.ListIndex, glngSys, mlngModul, blnHavePrivs
        Call SaveInvoice
        
'        lngShareID = 0
'        For i = 1 To lvwBill.ListItems.Count
'            If lvwBill.ListItems(i).Checked Then lngShareID = Val(Mid(lvwBill.ListItems(i).Key, 2))
'        Next
'        zlDatabase.SetPara "���ý���Ʊ������", lngShareID, glngSys, mlngModul, blnHavePrivs
        
        'LED�豸
        zlDatabase.SetPara "LED��ʾ��ӭ��Ϣ", chk(chks.C14LED��ӭ��Ϣ).Value, glngSys, mlngModul, blnHavePrivs
                
        zlDatabase.SetPara "���ʼ����տ���", cbo���տ���.ListIndex, glngSys, mlngModul, blnHavePrivs
        
        zlDatabase.SetPara "��Ժ��������", Val(txtOutDay1.Text), glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "����ָ��Ԥ����", chk(chks.C15����ָ��Ԥ����).Value, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "���סԺ������������", chk(chks.C16���סԺ������������).Value, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "��Ժ���˽��ʺ��Զ���Ժ", chk(chks.C11�����Զ���Ժ).Value, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "��;������Ԥ��", chk(chks.C10��;������Ԥ��).Value, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "���������", chk(chks.C12����ÿɽ���).Value, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "���ʷ���ʱ��", IIf(optTime(1).Value, 1, 0), glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "��Լ��λ�����˴�ӡ", chk(chks.C13��Լ��λ�����˴�ӡ).Value, glngSys, mlngModul, blnHavePrivs
        '���˺� ����:27776 ����:2010-02-04 16:44:39
        zlDatabase.SetPara "�˿��վݴ�ӡ", cbo�˿��վ�.ItemData(cbo�˿��վ�.ListIndex), glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "���ʺ������Ϣ", chk���ʲ���.Value, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "���ʼ�鲡������", chk��������.Value, glngSys, mlngModul, blnHavePrivs  '30036
        zlDatabase.SetPara "�����˿�ȱʡ��ʽ", chkRefundStyle.Value, glngSys, mlngModul, blnHavePrivs  '30036
        zlDatabase.SetPara "����ʱ��Ѫ�Ѽ��", IIf(opt��Ѫ(0).Value, 0, 1), glngSys, mlngModul, blnHavePrivs '34260
        zlDatabase.SetPara "������ϸ��ӡ", cbo������ϸ.ItemData(cbo������ϸ.ListIndex), glngSys, mlngModul, blnHavePrivs
        
        '65352
        zlDatabase.SetPara "����Ԥ��ȱʡʹ�÷�ʽ", IIf(optMzDeposit(2).Value, 2, IIf(optMzDeposit(1).Value, 1, 0)), glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "Ԥ��Ʊ�ݴ�ӡ��ʽ", IIf(optBalanceDepositPrint(2).Value, 2, IIf(optBalanceDepositPrint(1).Value, 1, 0)), glngSys, mlngModul, blnHavePrivs
        
        '����Ԥ��Ʊ��
        strValue = ""
        With vsDeposit
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, .ColIndex("ѡ��"))) <> 0 And Val(.RowData(i)) <> 0 Then
                    strValue = strValue & "|" & Val(.RowData(i)) & "," & Val(.Cell(flexcpData, i, .ColIndex("ʹ�����")))
                End If
            Next
        End With
        If strValue <> "" Then strValue = Mid(strValue, 2)
        zlDatabase.SetPara "����Ԥ��Ʊ������", strValue, glngSys, mlngModul, blnHavePrivs
        
        '43153
        zlDatabase.SetPara "���ʽɿ��������", IIf(opt�ɿ�(0).Value, 0, 1), glngSys, mlngModul, blnHavePrivs
        '32322
        With vsBalanceSort
            strTemp = ""
            For i = 1 To .Rows - 1
                strTemp = strTemp & ";" & Trim(.TextMatrix(i, .ColIndex("�������")))
            Next
            If strTemp <> "" Then strTemp = Mid(strTemp, 2)
            zlDatabase.SetPara "���㷽ʽ��ʾ˳��", strTemp, glngSys, mlngModul, blnHavePrivs  '30036
        End With
    
    End If
GoOver:
    If mblnOnlyDrugStock Then
        Call zlInitҩ��
    Else
        Call InitLocPar(mlngModul)
    End If
    gblnOK = True
    mblnOnlyDrugStock = False
    Unload Me
End Sub

Private Sub Form_Activate()
    If stab.TabVisible(0) Then
        If chk(chks.C00��ҩ�丶��).Visible And chk(chks.C00��ҩ�丶��).Enabled Then chk(chks.C00��ҩ�丶��).SetFocus
    Else
        If vsBill.Enabled And vsBill.Visible Then vsBill.SetFocus
    End If
End Sub


Private Sub Loadҩ��()
    Dim rsTmp As ADODB.Recordset
        
    On Error GoTo errH
    Set rsTmp = GetDepartments("'��ҩ��','��ҩ��','��ҩ��','���ϲ���'", "2,3")
        
    cbo��ҩ.AddItem "�˹�ѡ��"
    cbo��ҩ.AddItem "�˹�ѡ��"
    cbo��ҩ.AddItem "�˹�ѡ��"
    cbo����.AddItem "�˹�ѡ��"
    
    If Not rsTmp.EOF Then
        rsTmp.Filter = "��������='��ҩ��'"
        Do While Not rsTmp.EOF
            cbo��ҩ.AddItem rsTmp!����
            cbo��ҩ.ItemData(cbo��ҩ.ListCount - 1) = rsTmp!ID
            
            lst��ҩ��.AddItem rsTmp!����
            lst��ҩ��.ItemData(lst��ҩ��.ListCount - 1) = rsTmp!ID
            
            rsTmp.MoveNext
        Loop
        rsTmp.Filter = "��������='��ҩ��'"
        Do While Not rsTmp.EOF
            cbo��ҩ.AddItem rsTmp!����
            cbo��ҩ.ItemData(cbo��ҩ.ListCount - 1) = rsTmp!ID
            
            lst��ҩ��.AddItem rsTmp!����
            lst��ҩ��.ItemData(lst��ҩ��.ListCount - 1) = rsTmp!ID
            
            rsTmp.MoveNext
        Loop
        rsTmp.Filter = "��������='��ҩ��'"
        Do While Not rsTmp.EOF
            cbo��ҩ.AddItem rsTmp!����
            cbo��ҩ.ItemData(cbo��ҩ.ListCount - 1) = rsTmp!ID
            
            lst��ҩ��.AddItem rsTmp!����
            lst��ҩ��.ItemData(lst��ҩ��.ListCount - 1) = rsTmp!ID
            
            rsTmp.MoveNext
        Loop
        
        rsTmp.Filter = "��������='���ϲ���'"
        Do While Not rsTmp.EOF
            cbo����.AddItem rsTmp!����
            cbo����.ItemData(cbo����.ListCount - 1) = rsTmp!ID
                            
            rsTmp.MoveNext
        Loop
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset, strSql As String
    Dim i As Long, strValue As String, blnParSet As Boolean, blnBillOptSet As Boolean
    Dim strDefault As String
    Dim varData As Variant
    Dim bytKind As Byte
    
    gblnOK = False
    On Error GoTo errH
    blnParSet = InStr(1, mstrPrivs, ";��������;") > 0

    If mbytInFun = 0 Then
        blnBillOptSet = InStr(1, GetInsidePrivs(Enum_Inside_Program.p���ʲ���), "����ѡ������") > 0
        '����1150�Ĳ���
        '--------------------------------------------------------------------------------------
    
        '���ݴ�ӡ
        chkBillPrint(chkBPS.C0����).Value = IIf(zlDatabase.GetPara("���ʴ�ӡ", glngSys, mlngModul, , Array(chkBillPrint(chkBPS.C0����)), blnParSet) = "1", 1, 0)
        
        
        '1150�Ĳ���
        '------------------------------------------------------------------
        '�շ����(�Һų���)
        strSql = "Select ����,���� as ��� From �շ���Ŀ��� Where ����<>'1' Order by ���"
        Call zlDatabase.OpenRecordset(rsTmp, strSql, Me.Caption)
        Do While Not rsTmp.EOF
            lst�շ����.AddItem rsTmp!���
            lst�շ����.ItemData(lst�շ����.NewIndex) = Asc(rsTmp!����)
            rsTmp.MoveNext
        Loop
        strValue = zlDatabase.GetPara("�շ����", glngSys, CModule, , Array(lst�շ����), blnBillOptSet)
        If strValue = "" Then
            For i = 0 To lst�շ����.ListCount - 1
                lst�շ����.Selected(i) = True
            Next
        Else
            For i = 0 To lst�շ����.ListCount - 1
                If InStr(strValue, Chr(lst�շ����.ItemData(i))) Then lst�շ����.Selected(i) = True
            Next
        End If
        If lst�շ����.ListCount > 0 Then lst�շ����.TopIndex = 0: lst�շ����.ListIndex = 0
        
        '���۲��˼���
        chk(chks.C04�������ۼ���).Value = IIf(zlDatabase.GetPara("�������۲��˼���", glngSys, CModule, , Array(chk(chks.C04�������ۼ���)), blnBillOptSet) = "1", 1, 0)
        chk(chks.C05סԺ���ۼ���).Value = IIf(zlDatabase.GetPara("סԺ���۲��˼���", glngSys, CModule, , Array(chk(chks.C05סԺ���ۼ���)), blnBillOptSet) = "1", 1, 0)
                      
        txtOutDay0.Text = Val(zlDatabase.GetPara("��Ժ��������", glngSys, CModule, 0, Array(txtOutDay0, lblOutDate(0), UDOutDay(0)), blnBillOptSet))
        If Val(zlDatabase.GetPara("��������ʾ��ʽ", glngSys, CModule, 0, Array(optDoctorKind(0), optDoctorKind(1)), blnBillOptSet)) = 1 Then
            optDoctorKind(0).Value = True
        Else
            optDoctorKind(1).Value = True
        End If
        
        
        chk(chks.C00��ҩ�丶��).Value = IIf(zlDatabase.GetPara("��ҩ����", glngSys, CModule, , Array(chk(chks.C00��ҩ�丶��)), blnBillOptSet) = "1", 1, 0)
        chk(chks.C01��������).Value = IIf(zlDatabase.GetPara("�������", glngSys, CModule, , Array(chk(chks.C01��������)), blnBillOptSet) = "1", 1, 0)
        chk(chks.C02�����˺���ʿ).Value = IIf(zlDatabase.GetPara("��ʾ��ʿ", glngSys, CModule, , Array(chk(chks.C02�����˺���ʿ)), blnBillOptSet) = "1", 1, 0)
        
        chk(chks.C03�����˶�����).Value = IIf(zlDatabase.GetPara("����ҽ��", glngSys, CModule, , Array(chk(chks.C03�����˶�����)), blnBillOptSet) = "1", 0, 1)
        chk(chks.C06Ƿ�Ѵ滮�۵�).Value = IIf(zlDatabase.GetPara("������Ϊ���۵�", glngSys, CModule, , Array(chk(chks.C06Ƿ�Ѵ滮�۵�)), blnBillOptSet) = "1", 1, 0)
        
                
        i = Val(zlDatabase.GetPara("����ҩƷ��λ", glngSys, CModule, 0, Array(opt����ҩƷ��λ(0), opt����ҩƷ��λ(1)), blnBillOptSet))
        opt����ҩƷ��λ(IIf(i = 0, 0, 1)).Value = True
        
       
        '--------------------------
        Call Loadҩ��
        
        strValue = zlDatabase.GetPara("ȱʡ��ҩ��", glngSys, CModule, , Array(cbo��ҩ), blnBillOptSet)
        If IsNumeric(strValue) Then Call zlControl.CboLocate(cbo��ҩ, strValue, True)
        If cbo��ҩ.ListIndex = -1 And Val(strValue) = 0 Then cbo��ҩ.ListIndex = 0
        
        strValue = zlDatabase.GetPara("ȱʡ��ҩ��", glngSys, CModule, , Array(cbo��ҩ), blnBillOptSet)
        If IsNumeric(strValue) Then Call zlControl.CboLocate(cbo��ҩ, strValue, True)
        If cbo��ҩ.ListIndex = -1 And Val(strValue) = 0 Then cbo��ҩ.ListIndex = 0
        
        strValue = zlDatabase.GetPara("ȱʡ��ҩ��", glngSys, CModule, , Array(cbo��ҩ), blnBillOptSet)
        If IsNumeric(strValue) Then Call zlControl.CboLocate(cbo��ҩ, strValue, True)
        If cbo��ҩ.ListIndex = -1 And Val(strValue) = 0 Then cbo��ҩ.ListIndex = 0
        
        strValue = zlDatabase.GetPara("ȱʡ���ϲ���", glngSys, CModule, , Array(cbo����), blnBillOptSet)
        If IsNumeric(strValue) Then Call zlControl.CboLocate(cbo����, strValue, True)
        If cbo����.ListIndex = -1 And Val(strValue) = 0 Then cbo����.ListIndex = 0
        
        chk(chks.C08����ҩ�����).Value = IIf(zlDatabase.GetPara("��ʾ����ҩ�����", glngSys, CModule, , Array(chk(chks.C08����ҩ�����)), blnBillOptSet) = "1", 1, 0)
        chk(chks.C07����ҩ����).Value = IIf(zlDatabase.GetPara("��ʾ����ҩ����", glngSys, CModule, , Array(chk(chks.C07����ҩ����)), blnBillOptSet) = "1", 1, 0)
        
        '���뷢ҩʱ��ѡ��
        '------------------------------------------------------------------
        strValue = zlDatabase.GetPara("��ҩ��ѡ��", glngSys, CModule, , Array(lst��ҩ��), blnBillOptSet)
        For i = 0 To lst��ҩ��.ListCount - 1
            If InStr("," & strValue & ",", "," & lst��ҩ��.ItemData(i) & ",") > 0 Then
                lst��ҩ��.Selected(i) = True
            End If
        Next
        strValue = zlDatabase.GetPara("��ҩ��ѡ��", glngSys, CModule, , Array(lst��ҩ��), blnBillOptSet)
        For i = 0 To lst��ҩ��.ListCount - 1
            If InStr("," & strValue & ",", "," & lst��ҩ��.ItemData(i) & ",") > 0 Then
                lst��ҩ��.Selected(i) = True
            End If
        Next
        strValue = zlDatabase.GetPara("��ҩ��ѡ��", glngSys, CModule, , Array(lst��ҩ��), blnBillOptSet)
        For i = 0 To lst��ҩ��.ListCount - 1
            If InStr("," & strValue & ",", "," & lst��ҩ��.ItemData(i) & ",") > 0 Then
                lst��ҩ��.Selected(i) = True
            End If
        Next
        If lst��ҩ��.ListCount > 0 Then lst��ҩ��.ListIndex = 0
        If lst��ҩ��.ListCount > 0 Then lst��ҩ��.ListIndex = 0
        If lst��ҩ��.ListCount > 0 Then lst��ҩ��.ListIndex = 0
        '------------------------------------------------------------------
        chkת��.Visible = False: txtת��.Visible = False
        If mbytUseType = 0 Then
            chkBillPrint(chkBPS.C1����).Value = IIf(zlDatabase.GetPara("���۴�ӡ", glngSys, CModule, , Array(chkBillPrint(chkBPS.C1����)), blnBillOptSet) = "1", 1, 0)
            chkBillPrint(chkBPS.C2���).Value = IIf(zlDatabase.GetPara("��˴�ӡ", glngSys, CModule, , Array(chkBillPrint(chkBPS.C2���)), blnBillOptSet) = "1", 1, 0)
            
            cboSendMateria.AddItem "����ҩ"
            cboSendMateria.AddItem "�Զ���ҩ"
            cboSendMateria.AddItem "��ʾ��ҩ"
            i = Val(zlDatabase.GetPara("���ʺ�ҩ", glngSys, CModule, 0, Array(cboSendMateria), blnBillOptSet))
            If i > cboSendMateria.ListCount Then i = 0
            cboSendMateria.ListIndex = i
        ElseIf mbytUseType = 1 Then
            '���˺� ����:27380 ����:2010-01-22 14:45:32
            chkת��.Visible = True: txtת��.Visible = True
            Dim strת�� As String
            'CModule
            strת�� = zlDatabase.GetPara("���ת������", glngSys, mlngModul, "0|3", Array(chkת��, txtת��), InStr(1, mstrPrivs, ";��������;") > 0)
            txtת��.Text = Val(Split(strת�� & "|", "|")(1))
            chkת��.Value = IIf(Val(Split(strת�� & "|", "|")(0)) = 1, 1, 0)
        End If
        
    ElseIf mbytInFun = 1 Then
        '���˺� ����:27776 ����:2010-02-04 16:44:39
        i = Val(zlDatabase.GetPara("�˿��վݴ�ӡ", glngSys, mlngModul, , Array(lbl�˿��վ�, cbo�˿��վ�), blnParSet))
        With cbo�˿��վ�
            .AddItem "0-����ӡ": .ItemData(.NewIndex) = 0: If i = 0 Then .ListIndex = .NewIndex
            .AddItem "1-��ʾ��ӡ": .ItemData(.NewIndex) = 1: .ItemData(.NewIndex) = 1: If i = 1 Then .ListIndex = .NewIndex
            .AddItem "2-��ӡ,������ʾ": .ItemData(.NewIndex) = 2: .ItemData(.NewIndex) = 2: If i = 2 Then .ListIndex = .NewIndex
            If .ListIndex < 0 Then .ListIndex = 0
        End With
        '����:35511
        i = Val(zlDatabase.GetPara("������ϸ��ӡ", glngSys, mlngModul, , Array(lblListPrint, cbo������ϸ), blnParSet))
        With cbo������ϸ
            .AddItem "0-����ӡ": .ItemData(.NewIndex) = 0: If i = 0 Then .ListIndex = .NewIndex
            .AddItem "1-��ʾ��ӡ": .ItemData(.NewIndex) = 1: .ItemData(.NewIndex) = 1: If i = 1 Then .ListIndex = .NewIndex
            .AddItem "2-��ӡ,������ʾ": .ItemData(.NewIndex) = 2: .ItemData(.NewIndex) = 2: If i = 2 Then .ListIndex = .NewIndex
            If .ListIndex < 0 Then .ListIndex = 0
        End With
        chk���ʲ���.Value = IIf(Val(zlDatabase.GetPara("���ʺ������Ϣ", glngSys, mlngModul, , Array(chk���ʲ���), blnParSet)) = 1, 1, 0)
        chk��������.Value = IIf(Val(zlDatabase.GetPara("���ʼ�鲡������", glngSys, mlngModul, , Array(chk��������), blnParSet)) = 1, 1, 0) '30036
        chkRefundStyle.Value = IIf(Val(zlDatabase.GetPara("�����˿�ȱʡ��ʽ", glngSys, mlngModul, , Array(chkRefundStyle), blnParSet)) = 1, 1, 0)
       If Val(zlDatabase.GetPara("����ʱ��Ѫ�Ѽ��", glngSys, mlngModul, , Array(opt��Ѫ(0), opt��Ѫ(1), fra��Ѫ���), blnParSet)) = 1 Then '34260
            opt��Ѫ(1).Value = True
       Else
            opt��Ѫ(0).Value = True
       End If
       '43153
       If Val(zlDatabase.GetPara("���ʽɿ��������", glngSys, mlngModul, , Array(opt�ɿ�(0), opt�ɿ�(1), fra�ɿ����), blnParSet)) = 1 Then  '34260
            opt�ɿ�(1).Value = True
       Else
            opt�ɿ�(0).Value = True
       End If

        cboInvoiceKindZY.AddItem "סԺҽ�Ʒ��վ�"
        cboInvoiceKindZY.AddItem "����ҽ�Ʒ��վ�"
        i = Val(zlDatabase.GetPara("סԺ����Ʊ������", glngSys, mlngModul, 0, Array(cboInvoiceKindZY), blnParSet))
        If i <> 0 Then i = 1
        cboInvoiceKindZY.ListIndex = i
        
        cboInvoiceKindMZ.AddItem "סԺҽ�Ʒ��վ�"
        cboInvoiceKindMZ.AddItem "����ҽ�Ʒ��վ�"
        i = Val(zlDatabase.GetPara("�������Ʊ������", glngSys, mlngModul, 0, Array(cboInvoiceKindMZ), blnParSet))
        If i <> 0 Then i = 1
        cboInvoiceKindMZ.ListIndex = i
        
        If InStr(1, mstrPrivs, ";������ý���;") = 0 Then '�������������ý���ʱ,ֻ��ʹ��סԺҽ�Ʒ��վ�
            cboInvoiceKindZY.ListIndex = 0
            cboInvoiceKindZY.Enabled = False
            cboInvoiceKindMZ.Enabled = False
        End If
        
        If cboInvoiceKindZY.ListIndex = 0 And cboInvoiceKindMZ.ListIndex = 0 Then
            bytKind = InvoiceKind.C3�����վ�
        ElseIf cboInvoiceKindZY.ListIndex = 1 And cboInvoiceKindMZ.ListIndex = 1 Then
            bytKind = InvoiceKind.C1�շ��վ�
        Else
            bytKind = InvoiceKind.C4�����վ�
        End If
        Call InitShareInvoice(bytKind)
        Call InitDepositInvoice
        'Call SetShareInvoice(IIf(cboInvoiceKindZY.ListIndex = 0, InvoiceKind.C3�����վ�, InvoiceKind.C1�շ��վ�))
        '����:35142
        'Call SetFactBillFormat '������ͨ��ҽ�����˽��ʷ�Ʊ��ʽ
        
        'LED�豸
        chk(chks.C14LED��ӭ��Ϣ).Value = IIf(zlDatabase.GetPara("LED��ʾ��ӭ��Ϣ", glngSys, mlngModul, "1", Array(chk(chks.C14LED��ӭ��Ϣ)), blnParSet) = "1", 1, 0)
        
        cbo���տ���.AddItem "0-��ֹ"
        cbo���տ���.AddItem "1-����"
        cbo���տ���.ListIndex = IIf(zlDatabase.GetPara("���ʼ����տ���", glngSys, mlngModul, , Array(cbo���տ���), blnParSet) = "1", 1, 0)
        
        txtOutDay1.Text = Val(zlDatabase.GetPara("��Ժ��������", glngSys, mlngModul, 0, Array(txtOutDay1, lblOutDate(1), UDOutDay(1)), blnParSet))
        chk(chks.C13��Լ��λ�����˴�ӡ).Value = IIf(zlDatabase.GetPara("��Լ��λ�����˴�ӡ", glngSys, mlngModul, , Array(chk(chks.C13��Լ��λ�����˴�ӡ)), blnParSet) = "1", 1, 0)
        chk(chks.C15����ָ��Ԥ����).Value = IIf(zlDatabase.GetPara("����ָ��Ԥ����", glngSys, mlngModul, , Array(chk(chks.C15����ָ��Ԥ����)), blnParSet) = "1", 1, 0)
        chk(chks.C16���סԺ������������).Value = IIf(zlDatabase.GetPara("���סԺ������������", glngSys, mlngModul, , Array(chk(chks.C16���סԺ������������)), blnParSet) = "1", 1, 0)
        chk(chks.C10��;������Ԥ��).Value = IIf(zlDatabase.GetPara("��;������Ԥ��", glngSys, mlngModul, , Array(chk(chks.C10��;������Ԥ��)), blnParSet) = "1", 1, 0)
        chk(chks.C11�����Զ���Ժ).Value = IIf(zlDatabase.GetPara("��Ժ���˽��ʺ��Զ���Ժ", glngSys, mlngModul, , Array(chk(chks.C11�����Զ���Ժ)), blnParSet) = "1", 1, 0)
        chk(chks.C12����ÿɽ���).Value = IIf(zlDatabase.GetPara("���������", glngSys, mlngModul, , Array(chk(chks.C12����ÿɽ���)), blnParSet) = "1", 1, 0)
                
        i = Val(zlDatabase.GetPara("���ʷ���ʱ��", glngSys, mlngModul, 0, Array(optTime(0), optTime(1)), blnParSet))
        If i <> 0 Then i = 1
        optTime(i).Value = True
        
        '65352
        i = Val(zlDatabase.GetPara("����Ԥ��ȱʡʹ�÷�ʽ", glngSys, mlngModul, 2, Array(optMzDeposit(0), optMzDeposit(1), optMzDeposit(2), fraMzDepositDefaultUse), blnParSet))
        If i < 0 Or i > 2 Then i = 2
        optMzDeposit(i).Value = True
        
        i = Val(zlDatabase.GetPara("Ԥ��Ʊ�ݴ�ӡ��ʽ", glngSys, mlngModul, 2, Array(optBalanceDepositPrint(0), optBalanceDepositPrint(1), optBalanceDepositPrint(2), fraDepositPrint), blnParSet))
        If i < 0 Or i > 2 Then i = 2
        optBalanceDepositPrint(i).Value = True
        
        
        '32322
        strDefault = "��ҽ������-�н��;��ҽ������-�޽��;ҽ������-�н���������޸�;ҽ������-�޽���������޸�;ҽ������-�н���Ҳ������޸�;ҽ������-�޽���Ҳ������޸�"
        strValue = Trim(zlDatabase.GetPara("���㷽ʽ��ʾ˳��", glngSys, mlngModul, strDefault, Array(vsBalanceSort, cmdUp, cmdDown), blnParSet))
        varData = Split(strValue, ";")
        With vsBalanceSort
            .Clear 1
            .Rows = 2
            For i = 0 To UBound(varData)
                .TextMatrix(i + 1, .ColIndex("���")) = i + 1
                 .TextMatrix(i + 1, .ColIndex("�������")) = varData(i)
                 If i < UBound(varData) Then .Rows = .Rows + 1
            Next
        End With
    End If
    If mbytInFun = 0 Then
        cboSendMateria.Visible = (mbytInFun = 0 And mbytUseType = 0)
        lbl��ҩ.Visible = (mbytInFun = 0 And mbytUseType = 0)
    
        If gbln���뷢ҩ Then
            fraҩ��.Visible = False
            fra����.Visible = True
        End If
        stab.TabVisible(1) = False
        stab.TabVisible(2) = False
        stab.TabVisible(3) = False
        stab.TabVisible(4) = False
        If mbytUseType <> 0 Then
            chkBillPrint(1).Visible = False
            chkBillPrint(2).Visible = False
        End If
        
        '����:27380
        txtת��.Visible = mbytUseType = 1 '���ҷ�ɢ����
        chkת��.Visible = mbytUseType = 1 '���ҷ�ɢ����

    ElseIf mbytInFun = 1 Then
        If InStr(1, mstrPrivs, ";������ý���;") = 0 Then chk(chks.C13��Լ��λ�����˴�ӡ).Visible = False
        stab.TabVisible(0) = False
    End If
    Call zlOnlyDrugStrock
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'
'Private Sub SetShareInvoice(ByVal bytKind As Byte)
'    Dim rstmp As New ADODB.Recordset, strSQL As String
'    Dim i As Long, lngShareID As Long
'    Dim objItem As ListItem
'
'    '��ȡ���ù��ý�������
'    Set rstmp = GetShareInvoiceGroupID(bytKind)
'    lngShareID = Val(zlDatabase.GetPara("���ý���Ʊ������", glngSys, mlngModul, 0, Array(lvwBill), InStr(1, mstrPrivs, ";��������;") > 0))
'    lvwBill.ListItems.Clear
'    For i = 1 To rstmp.RecordCount
'        Set objItem = lvwBill.ListItems.Add(, "_" & rstmp!ID, rstmp!������, , 1)
'        objItem.SubItems(1) = Format(rstmp!�Ǽ�ʱ��, "yyyy-MM-dd")
'        objItem.SubItems(2) = rstmp!��ʼ���� & "," & rstmp!��ֹ����
'        objItem.SubItems(3) = rstmp!ʣ������
'        If rstmp!ID = lngShareID Then
'            objItem.Checked = True
'            objItem.Selected = True
'            lngShareID = 0
'        End If
'        rstmp.MoveNext
'    Next
'    If lngShareID <> 0 Then zlDatabase.SetPara "���ý���Ʊ������", 0, glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
'
'    Exit Sub
'errH:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbytInFun = 0
    mbytUseType = 0
End Sub

Private Sub lst�շ����_ItemCheck(Item As Integer)
    If lst�շ����.SelCount = 0 And Not lst�շ����.Selected(Item) Then
        lst�շ����.Selected(Item) = True
    End If
End Sub
'
'Private Sub lvwBill_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'    Dim i As Long
'    For i = 1 To lvwBill.ListItems.Count
'        If lvwBill.ListItems(i).Key <> Item.Key Then lvwBill.ListItems(i).Checked = False
'    Next
'    Item.Selected = True
'End Sub

Private Sub txtOutDay0_GotFocus()
    SelAll txtOutDay0
End Sub

Private Sub txtOutDay0_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtOutDay1_GotFocus()
    SelAll txtOutDay1
End Sub

Private Sub txtOutDay1_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
Private Function CheckUnits() As Boolean
'���ܣ���鰴��������֮��,�Ƿ��п��ü����ٴ�����
'˵��������ʹ���������ۼ���֮��,������ʾ�����ٴ�����
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, lng����ID As Long
    Dim strSql As String
    
    On Error GoTo errH
    
    '��Ȩ����ʾ����۲��Ҷ�Ӧ���ٴ�����,סԺ������סԺ��ͬ
    If InStr(mstrPrivs, ";�������ۼ���;") And (chk(chks.C04�������ۼ���).Value = 1) Then
        strSql = "1,2,3"
    Else
        strSql = "2,3"
    End If
    If InStr(";" & mstrPrivs, ";���в���;") > 0 Then
        strSql = _
             " Select Distinct A.ID,A.����,A.����" & _
             " From ���ű� A,��������˵�� B" & _
             " Where B.����ID = A.ID And B.������� IN(" & strSql & ") And B.�������� IN('�ٴ�','����')" & _
             " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
             " Order by A.����"
    Else
        '����Ȩ�޵Ŀ��ң��������ڿ���+�������������Ŀ���
        '#������Ա��������۲���ʱ����ʹû���������ۼ��ʵ�Ȩ��,Ҳ��ʾ��Ӧ�������ٴ�����,���޷�����
        strSql = _
            " Select A.ID,A.����,A.����,Nvl(C.ȱʡ,0) as ȱʡ" & _
            " From ���ű� A,��������˵�� B,������Ա C" & _
            " Where A.ID=B.����ID And A.ID=C.����ID And C.��ԱID=[1]" & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
            " And B.������� IN(" & strSql & ") And B.�������� IN('�ٴ�','����')" & _
            " Order by A.����"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)
    CheckUnits = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Sub vsBalanceSort_AfterMoveRow(ByVal Row As Long, Position As Long)
    '����˳��
    '32322
    Call RefreshNO
    Call SetDownAndUpEnable
End Sub
Private Sub RefreshNO()
    Dim lngRow As Long
    With vsBalanceSort
        For lngRow = 1 To .Rows - 1
            .TextMatrix(lngRow, .ColIndex("���")) = lngRow
        Next
    End With
End Sub
Private Sub cmdDown_Click()
    With vsBalanceSort
        If .Row >= .Rows - 1 Then Exit Sub
        .RowPosition(.Row) = .Row + 1
        .Row = .Row + 1
    End With
    Call RefreshNO
End Sub
Private Sub cmdUp_Click()
    With vsBalanceSort
        If .Row <= 1 Then Exit Sub
        .RowPosition(.Row) = .Row - 1
        .Row = .Row - 1
    End With
    Call RefreshNO
End Sub
Private Sub SetDownAndUpEnable()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������¿ؼ���Enable����
    '����:���˺�
    '����:2010-09-26 11:11:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    With vsBalanceSort
        cmdUp.Enabled = vsBalanceSort.Enabled And .Row > 1
        cmdDown.Enabled = vsBalanceSort.Enabled And (.Row < .Rows - 1)
    End With
ErrHand:
End Sub
Private Sub vsBalanceSort_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call SetDownAndUpEnable
End Sub
'
'Private Sub SetFactBillFormat()
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '����:���÷�Ʊ��ʽ
'    '����:���˺�
'    '����:2010-12-31 19:29:48
'    '����:35142
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim strRptName As String, rstmp As ADODB.Recordset, i As Long, blnParSet As Boolean, strSQL As String
'    blnParSet = zlCheckPrivs(mstrPrivs, ";��������;")
'    strRptName = IIf(cboInvoiceKindZY.ListIndex = 0, "ZL" & glngSys \ 100 & "_BILL_1137", "ZL" & glngSys \ 100 & "_BILL_1137_2")
'    cboFactNormal.Clear: cboFactMediCare.Clear
'
'    cboFactNormal.AddItem "ʹ�ñ���ȱʡ��ʽ"
'    cboFactMediCare.AddItem "ʹ�ñ���ȱʡ��ʽ"
'    '    Call ReportPrintSet(gcnOracle, glngSys, IIf(cboInvoiceKindZY.ListIndex = 0, "ZL" & glngSys \ 100 & "_BILL_1137", "ZL" & glngSys \ 100 & "_BILL_1137_2"), Me)
'    strSQL = "" & _
'    "   Select B.˵��,B.��� From zlReports A,zlRptFmts B" & _
'    "    Where A.ID=B.����ID And A.���=[1] " & _
'    "   Order by b.���"
'    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strRptName)
'    For i = 1 To rstmp.RecordCount
'        cboFactNormal.AddItem rstmp!˵��
'        cboFactNormal.ItemData(cboFactNormal.NewIndex) = rstmp!���
'        cboFactMediCare.AddItem rstmp!˵��
'        cboFactMediCare.ItemData(cboFactMediCare.NewIndex) = rstmp!���
'        rstmp.MoveNext
'    Next
'    cboFactNormal.ListIndex = 0: cboFactMediCare.ListIndex = 0
'    i = Val(zlDatabase.GetPara("��ͨ��Ʊ��ʽ", glngSys, mlngModul, , Array(lblFactNormal, cboFactNormal), blnParSet))
'    Call zlControl.CboLocate(cboFactNormal, i, True)
'    i = Val(zlDatabase.GetPara("ҽ����Ʊ��ʽ", glngSys, mlngModul, , Array(lblFactMediCare, cboFactMediCare), blnParSet))
'    Call zlControl.CboLocate(cboFactMediCare, i, True)
'End Sub

Private Sub vsBill_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        Dim i As Long
        With vsBill
            Select Case Col
            Case .ColIndex("ѡ��")
                If Val(.TextMatrix(Row, Col)) <> 0 And Val(.RowData(Row)) <> 0 Then
                    For i = 1 To .Rows - 1
                        If Trim(.TextMatrix(Row, .ColIndex("ʹ�����"))) = Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) _
                            And i <> Row Then
                            .TextMatrix(i, Col) = 0
                        End If
                    Next
                End If
            End Select
        End With
End Sub
 
Private Sub vsBill_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModul, vsBill, Me.Name, "����Ʊ��������", False, False
End Sub

Private Sub vsBill_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModul, vsBill, Me.Name, "����Ʊ��������", False, False
End Sub

Private Sub vsBill_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        With vsBill
            If Val(.Tag) = 1 Then
                If InStr(1, mstrPrivs, ";��������;") = 0 Then Cancel = True: Exit Sub
            End If
            Select Case Col
            Case .ColIndex("ѡ��")
                If Val(.RowData(Row)) = 0 Then Cancel = True
            Case Else
                Cancel = True
            End Select
        End With
End Sub

Private Sub vsBillFormat_AfterMoveColumn(ByVal Col As Long, Position As Long)
        zl_vsGrid_Para_Save mlngModul, vsBillFormat, Me.Name, "����Ʊ�ݸ�ʽ", False, False
End Sub
Private Sub vsBillFormat_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
        zl_vsGrid_Para_Save mlngModul, vsBillFormat, Me.Name, "����Ʊ�ݸ�ʽ", False, False
End Sub

Private Sub vsBillFormat_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsBillFormat
        Select Case Col
        Case .ColIndex("�������Ʊ�ݸ�ʽ"), .ColIndex("���ʺ��ӡ��ʽ"), .ColIndex("��Լ��λ����"), .ColIndex("סԺ����Ʊ�ݸ�ʽ")
            If Val(.ColData(Col)) = 1 And InStr(1, mstrPrivs, ";��������;") = 0 Then Cancel = True: Exit Sub
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub vsRedFormat_AfterMoveColumn(ByVal Col As Long, Position As Long)
        zl_vsGrid_Para_Save mlngModul, vsRedFormat, Me.Name, "���ʺ�Ʊ��ʽ", False, False
End Sub
Private Sub vsRedFormat_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
        zl_vsGrid_Para_Save mlngModul, vsRedFormat, Me.Name, "���ʺ�Ʊ��ʽ", False, False
End Sub

Private Sub vsRedFormat_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsRedFormat
        Select Case Col
        Case .ColIndex("Ʊ�ݸ�ʽ"), .ColIndex("���Ϻ��ӡ��ʽ"), .ColIndex("��Լ��λ����")
            If Val(.ColData(Col)) = 1 And InStr(1, mstrPrivs, ";��������;") = 0 Then Cancel = True: Exit Sub
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub SaveInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���淢Ʊ���Ʊ��
    '����:���˺�
    '����:2011-04-28 18:16:48
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHavePrivs As Boolean, strValue As String
    Dim i As Long
    Dim strPrintMode As String, str��Լ���� As String
    Dim strMZValue As String, strZYValue As String
    
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    '���湲��Ʊ��
    strValue = ""
    With vsBill
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("ѡ��"))) <> 0 And Val(.RowData(i)) <> 0 Then
                strValue = strValue & "|" & Val(.RowData(i)) & "," & Trim(.TextMatrix(i, .ColIndex("ʹ�����")))
            End If
        Next
    End With
    If strValue <> "" Then strValue = Mid(strValue, 2)
    zlDatabase.SetPara "���ý���Ʊ������", strValue, glngSys, mlngModul, blnHavePrivs
    
    '�����շѸ�ʽ
    strValue = "": strPrintMode = "": str��Լ���� = "��ͨ����"
    With vsBillFormat
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) <> "" Then
                strMZValue = strMZValue & "|" & Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) & "," & Val(.TextMatrix(i, .ColIndex("�������Ʊ�ݸ�ʽ")))
                strZYValue = strZYValue & "|" & Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) & "," & Val(.TextMatrix(i, .ColIndex("סԺ����Ʊ�ݸ�ʽ")))
                strPrintMode = strPrintMode & "|" & Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("���ʺ��ӡ��ʽ")), 1))
            End If
        Next
        str��Լ���� = Trim(cboʹ�����.Text)
        If strMZValue <> "" Then strMZValue = Mid(strMZValue, 2)
        If strZYValue <> "" Then strZYValue = Mid(strZYValue, 2)
        If strPrintMode <> "" Then strPrintMode = Mid(strPrintMode, 2)
        zlDatabase.SetPara "������ʷ�Ʊ��ʽ", strMZValue, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "סԺ���ʷ�Ʊ��ʽ", strZYValue, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "���˽��ʴ�ӡ", strPrintMode, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "��Լ��λ���ʴ�ӡ", str��Լ����, glngSys, mlngModul, blnHavePrivs
    End With
    '�����Ʊ��ʽ
    strValue = "": strPrintMode = ""
    With vsRedFormat
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) <> "" Then
                strValue = strValue & "|" & Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) & "," & Val(.TextMatrix(i, .ColIndex("Ʊ�ݸ�ʽ")))
                strPrintMode = strPrintMode & "|" & Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("���Ϻ��ӡ��ʽ")), 1))
            End If
        Next
        If strValue <> "" Then strValue = Mid(strValue, 2)
        If strPrintMode <> "" Then strPrintMode = Mid(strPrintMode, 2)
        zlDatabase.SetPara "���Ϸ�Ʊ��ʽ", strValue, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "���Ϸ�Ʊ��ӡ��ʽ", strPrintMode, glngSys, mlngModul, blnHavePrivs
    End With
End Sub


Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ч�Լ��
    '����:���Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-04-28 18:24:16
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, lngSelCount As Long, str��� As String
     If mbytInFun <> 0 Then isValied = True: Exit Function
     
    isValied = False
    On Error GoTo errHandle
    '���ÿ��ʹ����ʽֻ��һ��ѡ��
    With vsBill
        str��� = "-"
        For i = 1 To vsBill.Rows - 1
            If str��� <> Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) Then
               str��� = Trim(.TextMatrix(i, .ColIndex("ʹ�����")))
               lngSelCount = 0
                For j = 1 To vsBill.Rows - 1
                    If Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) = Trim(.TextMatrix(j, .ColIndex("ʹ�����"))) Then
                        If Val(.TextMatrix(j, .ColIndex("ѡ��"))) <> 0 Then
                            lngSelCount = lngSelCount + 1
                        End If
                    End If
                Next
                If lngSelCount > 1 Then
                    MsgBox "ע��:" & vbCrLf & "    ʹ�����Ϊ��" & str��� & "����ֻ��ѡ��һ��Ʊ��,����!", vbInformation + vbOKOnly
                    Exit Function
                End If
            End If
        Next
    End With
    If cboʹ�����.ListIndex < 0 Then
        MsgBox "ע��:" & vbCrLf & "    ��δѡ���Լ��λ����ʱ��ʹ�õĺ���Ʊ��!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub InitShareInvoice(ByVal intKind As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ù���Ʊ
    '����:���˺�
    '����:2011-04-28 15:09:10
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, lngRow As Long
    Dim strShareInvoice As String '����Ʊ������,��ʽ:����,����
    Dim varData As Variant, varTemp As Variant
    Dim VarType As Variant, varTemp1 As Variant
    Dim intType As Integer, intType1 As Integer, intType2 As Integer   '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    Dim lngTemp As Long, i As Long, strSql As String
    Dim strRptName As String, blnHavePrivs As Boolean
    Dim strPrintMode As String, varDataMZ As Variant
    Dim str��Լ��λ���� As String, strShareInvoiceMZ As String
    
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    
    On Error GoTo errHandle
    
    '�ָ��п��
    zl_vsGrid_Para_Restore mlngModul, vsBill, Me.Name, "����Ʊ��������", False, False
    zl_vsGrid_Para_Restore mlngModul, vsBillFormat, Me.Name, "����Ʊ�ݸ�ʽ", False, False
    zl_vsGrid_Para_Restore mlngModul, vsRedFormat, Me.Name, "���ʺ�Ʊ��ʽ", False, False
    strShareInvoice = zlDatabase.GetPara("���ý���Ʊ������", glngSys, mlngModul, , , True, intType)
    '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    vsBill.Tag = ""
    Select Case intType
    Case 1, 3, 5, 15
        vsBill.ForeColor = vbBlue: vsBill.ForeColorFixed = vbBlue
        fraTitle.ForeColor = vbBlue: vsBill.Tag = 1
        If intType = 5 Then vsBill.Tag = ""
    Case Else
        vsBill.ForeColor = &H80000008: vsBill.ForeColorFixed = &H80000008
        fraTitle.ForeColor = &H80000008
    End Select
    With vsBill
        .Editable = flexEDKbdMouse
        If Val(.Tag) = 1 And Not blnHavePrivs Then .Editable = flexEDNone
    End With
    
    
    '��ʽ:����ID1,ʹ�����1|����IDn,ʹ�����n|...
    varData = Split(strShareInvoice, "|")
    '1.���ù���Ʊ��
    Set rsTemp = GetShareInvoiceGroupID(intKind)
    With vsBill
        .Clear 1: .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFixedOnly
        .MergeCol(0) = True
        Do While Not rsTemp.EOF
            .RowData(lngRow) = Val(Nvl(rsTemp!ID))
            .TextMatrix(lngRow, .ColIndex("ʹ�����")) = Nvl(rsTemp!ʹ�����, " ")
            .TextMatrix(lngRow, .ColIndex("������")) = Nvl(rsTemp!������)
            .TextMatrix(lngRow, .ColIndex("��������")) = Format(rsTemp!�Ǽ�ʱ��, "yyyy-MM-dd")
            .TextMatrix(lngRow, .ColIndex("���뷶Χ")) = rsTemp!��ʼ���� & "," & rsTemp!��ֹ����
            .TextMatrix(lngRow, .ColIndex("ʣ��")) = Format(Val(Nvl(rsTemp!ʣ������)), "##0;-##0;;")
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & ",", ",")
                lngTemp = Val(varTemp(0))
                If Val(.RowData(lngRow)) = lngTemp _
                    And varTemp(1) = Trim(.TextMatrix(lngRow, .ColIndex("ʹ�����"))) Then
                    .TextMatrix(lngRow, .ColIndex("ѡ��")) = -1: Exit For
                End If
            Next
            .MergeRow(lngRow) = True
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    
    strRptName = IIf(cboInvoiceKindZY.ListIndex = 0, "ZL" & glngSys \ 100 & "_BILL_1137", "ZL" & glngSys \ 100 & "_BILL_1137_2")
    'סԺƱ�ݸ�ʽ����
    strSql = "" & _
    "   Select 'ʹ�ñ���ȱʡ��ʽ' as ˵��,0 as ���  From Dual Union ALL " & _
    "   Select B.˵��,B.���  " & _
    "   From zlReports A,zlRptFmts B" & _
    "   Where A.ID=B.����ID And A.���=[1]" & _
    "   Order by  ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strRptName)
    With vsBillFormat
        .Clear 1
        .ColComboList(.ColIndex("סԺ����Ʊ�ݸ�ʽ")) = .BuildComboList(rsTemp, "���,*˵��", "���")
    End With
    
    strRptName = IIf(cboInvoiceKindMZ.ListIndex = 0, "ZL" & glngSys \ 100 & "_BILL_1137", "ZL" & glngSys \ 100 & "_BILL_1137_2")
    '����Ʊ�ݸ�ʽ����
    strSql = "" & _
    "   Select 'ʹ�ñ���ȱʡ��ʽ' as ˵��,0 as ���  From Dual Union ALL " & _
    "   Select B.˵��,B.���  " & _
    "   From zlReports A,zlRptFmts B" & _
    "   Where A.ID=B.����ID And A.���=[1]" & _
    "   Order by  ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strRptName)
    With vsBillFormat
        .ColComboList(.ColIndex("�������Ʊ�ݸ�ʽ")) = .BuildComboList(rsTemp, "���,*˵��", "���")
    End With
    
    '��ȡ����ֵ
    strShareInvoice = zlDatabase.GetPara("סԺ���ʷ�Ʊ��ʽ", glngSys, mlngModul, , , True, intType)
    strShareInvoiceMZ = zlDatabase.GetPara("������ʷ�Ʊ��ʽ", glngSys, mlngModul, , , True, intType)
    strPrintMode = zlDatabase.GetPara("���˽��ʴ�ӡ", glngSys, mlngModul, , , True, intType1)
    str��Լ��λ���� = zlDatabase.GetPara("��Լ��λ���ʴ�ӡ", glngSys, mlngModul, "��ͨ����", Array(cboʹ�����, lblUnit), blnHavePrivs)
    '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    With vsBillFormat
         .ColData(.ColIndex("סԺ����Ʊ�ݸ�ʽ")) = "0"
         .ColData(.ColIndex("���ʺ��ӡ��ʽ")) = "0"
        .ForeColor = &H80000008:  .ForeColorFixed = &H80000008
        Select Case intType
        Case 1, 3, 5, 15
             .ColData(.ColIndex("סԺ����Ʊ�ݸ�ʽ")) = IIf(intType = 5, 0, 1)
        End Select
        Select Case intType1
        Case 1, 3, 5, 15
             .ColData(.ColIndex("���ʺ��ӡ��ʽ")) = IIf(intType1 = 5, 0, 1)
        End Select
        If (Val(.ColData(.ColIndex("סԺ����Ʊ�ݸ�ʽ"))) = 1 And _
            Val(.ColData(.ColIndex("���ʺ��ӡ��ʽ"))) = 1) Then
            .Editable = IIf(Not blnHavePrivs, flexEDNone, flexEDKbdMouse)
        Else
            .Editable = flexEDKbdMouse
        End If
        .ColComboList(.ColIndex("���ʺ��ӡ��ʽ")) = "0-����ӡƱ��|1-�Զ���ӡƱ��|2-ѡ���Ƿ��ӡƱ��"
    End With
    With vsBillFormat
         .ColData(.ColIndex("�������Ʊ�ݸ�ʽ")) = "0"
        .ForeColor = &H80000008:  .ForeColorFixed = &H80000008
        Select Case intType
        Case 1, 3, 5, 15
             .ColData(.ColIndex("�������Ʊ�ݸ�ʽ")) = IIf(intType = 5, 0, 1)
        End Select
        If Val(.ColData(.ColIndex("�������Ʊ�ݸ�ʽ"))) = 1 Then
            .Editable = IIf(Not blnHavePrivs, flexEDNone, flexEDKbdMouse)
        Else
            .Editable = flexEDKbdMouse
        End If
        '.ColComboList(.ColIndex("���ʺ��ӡ��ʽ")) = "0-����ӡƱ��|1-�Զ���ӡƱ��|2-ѡ���Ƿ��ӡƱ��"
    End With
    varData = Split(strShareInvoice, "|")
    varDataMZ = Split(strShareInvoiceMZ, "|")
    VarType = Split(strPrintMode, "|")
    strSql = "" & _
    "   Select ���� ,����" & _
    "   From  Ʊ��ʹ�����" & _
    "   order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    With vsBillFormat
        .Clear 1: cboʹ�����.Clear
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("ʹ�����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("���ʺ��ӡ��ʽ")) = "0-����ӡƱ��"
            .TextMatrix(lngRow, .ColIndex("סԺ����Ʊ�ݸ�ʽ")) = "0"
            .TextMatrix(lngRow, .ColIndex("�������Ʊ�ݸ�ʽ")) = "0"
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(Nvl(rsTemp!����)) Then
                    .TextMatrix(lngRow, .ColIndex("סԺ����Ʊ�ݸ�ʽ")) = Val(varTemp(1)): Exit For
                End If
            Next
            For i = 0 To UBound(varDataMZ)
                varTemp = Split(varDataMZ(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(Nvl(rsTemp!����)) Then
                    .TextMatrix(lngRow, .ColIndex("�������Ʊ�ݸ�ʽ")) = Val(varTemp(1)): Exit For
                End If
            Next
            For i = 0 To UBound(VarType)
                varTemp1 = Split(VarType(i) & "," & ",", ",")
                If Trim(varTemp1(0)) = Trim(Nvl(rsTemp!����)) Then
                    .TextMatrix(lngRow, .ColIndex("���ʺ��ӡ��ʽ")) = Decode(Val(varTemp1(1)), 0, "0-����ӡƱ��", 1, "1-�Զ���ӡƱ��", "2-ѡ���Ƿ��ӡƱ��")
                    Exit For
                End If
            Next
            cboʹ�����.AddItem Nvl(rsTemp!����)
            If Nvl(rsTemp!����) = str��Լ��λ���� Then
                cboʹ�����.ListIndex = cboʹ�����.NewIndex
            End If
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If Val(.ColData(.ColIndex("���ʺ��ӡ��ʽ"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("���ʺ��ӡ��ʽ"), .Rows - 1, .ColIndex("���ʺ��ӡ��ʽ")) = vbBlue
        End If
        If Val(.ColData(.ColIndex("סԺ����Ʊ�ݸ�ʽ"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("סԺ����Ʊ�ݸ�ʽ"), .Rows - 1, .ColIndex("סԺ����Ʊ�ݸ�ʽ")) = vbBlue
        End If
        If Val(.ColData(.ColIndex("�������Ʊ�ݸ�ʽ"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("�������Ʊ�ݸ�ʽ"), .Rows - 1, .ColIndex("�������Ʊ�ݸ�ʽ")) = vbBlue
        End If
    End With
    
    strRptName = IIf(cboInvoiceKindZY.ListIndex = 0, "ZL" & glngSys \ 100 & "_BILL_1137_5", "ZL" & glngSys \ 100 & "_BILL_1137_6")
    'Ʊ�ݸ�ʽ����
    strSql = "" & _
    "   Select 'ʹ�ñ���ȱʡ��ʽ' as ˵��,0 as ���  From Dual Union ALL " & _
    "   Select B.˵��,B.���  " & _
    "   From zlReports A,zlRptFmts B" & _
    "   Where A.ID=B.����ID And A.���=[1]" & _
    "   Order by  ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strRptName)
    
    With vsRedFormat
        .Clear 1
        .ColComboList(.ColIndex("Ʊ�ݸ�ʽ")) = .BuildComboList(rsTemp, "���,*˵��", "���")
    End With
    
    '��ȡ����ֵ
    strShareInvoice = zlDatabase.GetPara("���Ϸ�Ʊ��ʽ", glngSys, mlngModul, , , True, intType)
    strPrintMode = zlDatabase.GetPara("���Ϸ�Ʊ��ӡ��ʽ", glngSys, mlngModul, , , True, intType1)
    '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    With vsRedFormat
         .ColData(.ColIndex("Ʊ�ݸ�ʽ")) = "0"
         .ColData(.ColIndex("���Ϻ��ӡ��ʽ")) = "0"
        .ForeColor = &H80000008:  .ForeColorFixed = &H80000008
        Select Case intType
        Case 1, 3, 5, 15
             .ColData(.ColIndex("Ʊ�ݸ�ʽ")) = IIf(intType = 5, 0, 1)
        End Select
        Select Case intType1
        Case 1, 3, 5, 15
             .ColData(.ColIndex("���Ϻ��ӡ��ʽ")) = IIf(intType1 = 5, 0, 1)
        End Select
        If (Val(.ColData(.ColIndex("Ʊ�ݸ�ʽ"))) = 1 And _
            Val(.ColData(.ColIndex("���Ϻ��ӡ��ʽ"))) = 1) Then
            .Editable = IIf(Not blnHavePrivs, flexEDNone, flexEDKbdMouse)
        Else
            .Editable = flexEDKbdMouse
        End If
        .ColComboList(.ColIndex("���Ϻ��ӡ��ʽ")) = "0-����ӡƱ��|1-�Զ���ӡƱ��|2-ѡ���Ƿ��ӡƱ��"
    End With
    varData = Split(strShareInvoice, "|")
    VarType = Split(strPrintMode, "|")
    strSql = "" & _
    "   Select ���� ,����" & _
    "   From  Ʊ��ʹ�����" & _
    "   order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    With vsRedFormat
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("ʹ�����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("���Ϻ��ӡ��ʽ")) = "0-����ӡƱ��"
            .TextMatrix(lngRow, .ColIndex("Ʊ�ݸ�ʽ")) = "0"
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(Nvl(rsTemp!����)) Then
                    .TextMatrix(lngRow, .ColIndex("Ʊ�ݸ�ʽ")) = Val(varTemp(1)): Exit For
                End If
            Next
            For i = 0 To UBound(VarType)
                varTemp1 = Split(VarType(i) & "," & ",", ",")
                If Trim(varTemp1(0)) = Trim(Nvl(rsTemp!����)) Then
                    .TextMatrix(lngRow, .ColIndex("���Ϻ��ӡ��ʽ")) = Decode(Val(varTemp1(1)), 0, "0-����ӡƱ��", 1, "1-�Զ���ӡƱ��", "2-ѡ���Ƿ��ӡƱ��")
                    Exit For
                End If
            Next
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If Val(.ColData(.ColIndex("���Ϻ��ӡ��ʽ"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("���Ϻ��ӡ��ʽ"), .Rows - 1, .ColIndex("���Ϻ��ӡ��ʽ")) = vbBlue
        End If
        If Val(.ColData(.ColIndex("Ʊ�ݸ�ʽ"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("Ʊ�ݸ�ʽ"), .Rows - 1, .ColIndex("Ʊ�ݸ�ʽ")) = vbBlue
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub InitDepositInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ù���Ʊ
    '����:���˺�
    '����:2011-07-06 18:41:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, lngRow As Long
    Dim strShareInvoice As String '����Ʊ������,��ʽ:����,����
    Dim varData As Variant, varTemp As Variant, VarType As Variant, varTemp1 As Variant
    Dim intType As Integer, intType1 As Integer   '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    Dim lngTemp As Long, i As Long, strSql As String, rsҽ�ƿ���� As ADODB.Recordset
    Dim strPrintMode As String, blnHavePrivs As Boolean, lngCardTypeID As Long
    Dim strȱʡҽ�ƿ� As String, lngȱʡҽ�ƿ� As Long
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    
    On Error GoTo errHandle
    '����Ԥ��Ʊ������
    '�ָ��п��
    zl_vsGrid_Para_Restore mlngModul, vsDeposit, Me.Name, "����Ԥ��Ʊ���б�", False, False
    
    strShareInvoice = zlDatabase.GetPara("����Ԥ��Ʊ������", glngSys, mlngModul, , , True, intType)
    '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    Select Case intType
    Case 1, 3, 5, 15
        vsDeposit.ForeColor = vbBlue: vsDeposit.ForeColorFixed = vbBlue
        fraDepositPrint.ForeColor = vbBlue
    Case Else
        vsDeposit.ForeColor = &H80000008: vsDeposit.ForeColorFixed = &H80000008
        fraDepositPrint.ForeColor = &H80000008
    End Select
    With vsDeposit
        .Editable = flexEDKbdMouse
        If Val(.Tag) = 1 And InStr(1, mstrPrivs, ";��������;") = 0 Then .Editable = flexEDNone
    End With
    
    '��ʽ:����ID1,Ԥ�����ID1|����IDn,Ԥ�����IDn|...
    varData = Split(strShareInvoice, "|")
    '1.���ù���Ʊ��
    Set rsTemp = GetShareInvoiceGroupID(2)
    With vsDeposit
        .Clear 1: .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFixedOnly
        .MergeCol(0) = True
        Do While Not rsTemp.EOF
            .RowData(lngRow) = Val(Nvl(rsTemp!ID))
            '58071
            Select Case Val(Nvl(rsTemp!ʹ�����, ""))
            Case 0 '�����������סԺƱ��
                .TextMatrix(lngRow, .ColIndex("ʹ�����")) = ""
                .Cell(flexcpData, lngRow, .ColIndex("ʹ�����")) = 0
            Case 1  '����Ʊ��
                .TextMatrix(lngRow, .ColIndex("ʹ�����")) = "Ԥ������Ʊ��"
                .Cell(flexcpData, lngRow, .ColIndex("ʹ�����")) = 1
            Case Else   'סԺƱ��
                .TextMatrix(lngRow, .ColIndex("ʹ�����")) = "Ԥ��סԺƱ��"
                .Cell(flexcpData, lngRow, .ColIndex("ʹ�����")) = 2
            End Select
            
            .TextMatrix(lngRow, .ColIndex("������")) = Nvl(rsTemp!������)
            .TextMatrix(lngRow, .ColIndex("��������")) = Format(rsTemp!�Ǽ�ʱ��, "yyyy-MM-dd")
            .TextMatrix(lngRow, .ColIndex("���뷶Χ")) = rsTemp!��ʼ���� & "," & rsTemp!��ֹ����
            .TextMatrix(lngRow, .ColIndex("ʣ��")) = Format(Val(Nvl(rsTemp!ʣ������)), "##0;-##0;;")
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & ",", ",")
                lngTemp = Val(varTemp(0))
                If Val(.RowData(lngRow)) = lngTemp _
                    And varTemp(1) = Val(.Cell(flexcpData, lngRow, .ColIndex("ʹ�����"))) Then
                    .TextMatrix(lngRow, .ColIndex("ѡ��")) = -1: Exit For
                End If
            Next
            .MergeRow(lngRow) = True
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


