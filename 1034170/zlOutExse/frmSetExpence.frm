VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmSetExpence 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8430
   ControlBox      =   0   'False
   Icon            =   "frmSetExpence.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   6975
      TabIndex        =   79
      Top             =   6585
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6975
      TabIndex        =   78
      Top             =   840
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6975
      TabIndex        =   77
      Top             =   345
      Width           =   1100
   End
   Begin TabDlg.SSTab stab 
      Height          =   7110
      Left            =   90
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   6480
      _ExtentX        =   11430
      _ExtentY        =   12541
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   564
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "���ݿ���(&1)"
      TabPicture(0)   =   "frmSetExpence.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "chk��ֹȡ���Һŵ�"
      Tab(0).Control(1)=   "chk������"
      Tab(0).Control(2)=   "fraSetMoneyMode"
      Tab(0).Control(3)=   "fra��λ"
      Tab(0).Control(4)=   "chkҽ��������ȱʡ��λ"
      Tab(0).Control(5)=   "chkInsurePartFee"
      Tab(0).Control(6)=   "chkƤ��"
      Tab(0).Control(7)=   "fra�˷�ȱʡѡ��ʽ"
      Tab(0).Control(8)=   "fraDrugNotFee"
      Tab(0).Control(9)=   "chkPayKey"
      Tab(0).Control(10)=   "chk���������ɿ�"
      Tab(0).Control(11)=   "fra�����ʾ"
      Tab(0).Control(12)=   "chkAddedItem"
      Tab(0).Control(13)=   "txtAddedItem"
      Tab(0).Control(14)=   "cmdAddedItem"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "chkPrePayPriority"
      Tab(0).Control(16)=   "chkTime"
      Tab(0).Control(17)=   "chk��ʿ"
      Tab(0).Control(18)=   "chk�ۼ�"
      Tab(0).Control(19)=   "txtDay"
      Tab(0).Control(20)=   "txtMax"
      Tab(0).Control(21)=   "chkPay"
      Tab(0).Control(22)=   "fra���"
      Tab(0).Control(23)=   "cbo�ѱ�"
      Tab(0).Control(24)=   "cbo���㷽ʽ"
      Tab(0).Control(25)=   "udDay"
      Tab(0).Control(26)=   "fraPrintBill"
      Tab(0).Control(27)=   "chkסԺ�������շ�"
      Tab(0).Control(28)=   "fra����"
      Tab(0).Control(29)=   "lblDay"
      Tab(0).Control(30)=   "lblMax"
      Tab(0).Control(31)=   "lbl�ѱ�"
      Tab(0).Control(32)=   "lbl���㷽ʽ"
      Tab(0).ControlCount=   33
      TabCaption(1)   =   "�������(&2)"
      TabPicture(1)   =   "frmSetExpence.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "chk�շ�ִ�п���"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "chkSeekName"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fra����֪ͨ����ӡ"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "chkSeekBill"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdPrintSetup(3)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "fraInputItem"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "chkLed"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "chk�����俪����"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "chk��ȱʡ������"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "opt����(0)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "opt����(1)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "chkMulti"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "fra������ҽ��"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "fra����"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtSeekDays"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "fraLine"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "chkLedDispDetail"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "chkLedWelcome"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "fraDoctor"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "fraShortLine"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "txtNameDays"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "chkOnlyUnitPatient"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "cmdDeviceSetup"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "chkAutoSplitBill"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "chkȱʡ��������"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "chkUnPopPriceBill"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "cboAutoSplitBill"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "fraRegPrompt"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "chkMustRegevent"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "opt����(2)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "fra�ɿ����"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "txt�շ�ִ�п���"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "cmd�շ�ִ�п���"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).ControlCount=   33
      TabCaption(2)   =   "Ʊ�ݿ���(&3)"
      TabPicture(2)   =   "frmSetExpence.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdPrintSetup(7)"
      Tab(2).Control(1)=   "picDelBillFormat"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "picBillFormat"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "tbBillSet"
      Tab(2).Control(4)=   "fraFeeExe"
      Tab(2).Control(5)=   "cmdPrintSetup(6)"
      Tab(2).Control(6)=   "fraRefundReceipt"
      Tab(2).Control(7)=   "cmdPrintSetup(5)"
      Tab(2).Control(8)=   "fraFeeList"
      Tab(2).Control(9)=   "cmdPrintSetup(4)"
      Tab(2).Control(10)=   "updƱ������"
      Tab(2).Control(11)=   "txtƱ������"
      Tab(2).Control(12)=   "chkRegistInvoice"
      Tab(2).Control(13)=   "cmdPrintSetup(2)"
      Tab(2).Control(14)=   "cmdPrintSetup(1)"
      Tab(2).Control(15)=   "cmdPrintSetup(0)"
      Tab(2).Control(16)=   "fraTitle"
      Tab(2).Control(17)=   "chkƱ������"
      Tab(2).ControlCount=   18
      TabCaption(3)   =   "Ʊ�ŷ���(&4)"
      TabPicture(3)   =   "frmSetExpence.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "picBill(2)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "ҩ������(&5)"
      TabPicture(4)   =   "frmSetExpence.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cbo����"
      Tab(4).Control(1)=   "vsfDrugStore"
      Tab(4).Control(2)=   "lbl���ϲ���"
      Tab(4).ControlCount=   3
      Begin VB.CheckBox chk��ֹȡ���Һŵ� 
         Caption         =   "��ֹȡ���ҺŻ��۵�"
         Height          =   180
         Left            =   -74805
         TabIndex        =   34
         Top             =   5430
         Width           =   1920
      End
      Begin VB.CheckBox chk������ 
         Caption         =   "����¼������ʹ�õĿ�����"
         Height          =   180
         Left            =   -72450
         TabIndex        =   35
         Top             =   5430
         Width           =   2490
      End
      Begin VB.Frame fraSetMoneyMode 
         Caption         =   "�����շ�ˢ��ȱʡ����������������"
         Height          =   780
         Left            =   -74805
         TabIndex        =   43
         Top             =   6480
         Width           =   6075
         Begin VB.OptionButton optSetMoneyMode 
            Caption         =   "ȱʡˢ������ҽ���������"
            Height          =   210
            Index           =   1
            Left            =   210
            TabIndex        =   45
            Top             =   510
            Width           =   2670
         End
         Begin VB.OptionButton optSetMoneyMode 
            Caption         =   "ȱʡˢ������ҽ��������"
            Height          =   210
            Index           =   2
            Left            =   3150
            TabIndex        =   46
            Top             =   510
            Width           =   2820
         End
         Begin VB.OptionButton optSetMoneyMode 
            Caption         =   "��ȱʡˢ�����"
            Height          =   210
            Index           =   0
            Left            =   210
            TabIndex        =   44
            Top             =   255
            Value           =   -1  'True
            Width           =   1590
         End
      End
      Begin VB.Frame fra��λ 
         Caption         =   " ҩƷ��λ "
         Height          =   630
         Left            =   -74805
         TabIndex        =   21
         Top             =   2520
         Width           =   4455
         Begin VB.OptionButton opt��λ 
            Caption         =   "���ﵥλ"
            Height          =   180
            Index           =   1
            Left            =   2880
            TabIndex        =   23
            Top             =   285
            Width           =   1020
         End
         Begin VB.OptionButton opt��λ 
            Caption         =   "�ۼ۵�λ"
            Height          =   180
            Index           =   0
            Left            =   1590
            TabIndex        =   22
            Top             =   285
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.Label lbl��λ 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�շ�ʱ��"
            Height          =   180
            Left            =   600
            TabIndex        =   86
            Top             =   285
            Width           =   720
         End
      End
      Begin VB.CheckBox chkҽ��������ȱʡ��λ 
         Caption         =   "ҽ��������ȱʡ��λ����ҽ�����㡱��ť"
         Height          =   180
         Left            =   -72450
         TabIndex        =   31
         Top             =   4980
         Width           =   3735
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "�˷�Ʊ�ݴ�ӡ����(&2)"
         Height          =   350
         Index           =   7
         Left            =   -70725
         TabIndex        =   193
         Top             =   4020
         Width           =   1950
      End
      Begin VB.PictureBox picDelBillFormat 
         BorderStyle     =   0  'None
         Height          =   1260
         Left            =   -72780
         ScaleHeight     =   1260
         ScaleWidth      =   6015
         TabIndex        =   188
         TabStop         =   0   'False
         Top             =   2340
         Width           =   6015
         Begin VSFlex8Ctl.VSFlexGrid vsDelBillFormat 
            Height          =   1230
            Left            =   30
            TabIndex        =   189
            Top             =   30
            Width           =   5865
            _cx             =   10345
            _cy             =   2170
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
      Begin VB.PictureBox picBillFormat 
         BorderStyle     =   0  'None
         Height          =   1260
         Left            =   -74850
         ScaleHeight     =   1260
         ScaleWidth      =   6015
         TabIndex        =   190
         TabStop         =   0   'False
         Top             =   2100
         Width           =   6015
         Begin VB.CheckBox chkOnePatiPrint 
            Caption         =   "�����˲���Ʊ�ݲ����ݽ����������Ʊ"
            Height          =   180
            Left            =   30
            TabIndex        =   192
            Top             =   30
            Width           =   3540
         End
         Begin VSFlex8Ctl.VSFlexGrid vsBillFormat 
            Height          =   1020
            Left            =   30
            TabIndex        =   191
            Top             =   240
            Width           =   5865
            _cx             =   10345
            _cy             =   1799
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
            FormatString    =   $"frmSetExpence.frx":012E
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
      Begin XtremeSuiteControls.TabControl tbBillSet 
         Height          =   1650
         Left            =   -74760
         TabIndex        =   187
         Top             =   1980
         Width           =   5985
         _Version        =   589884
         _ExtentX        =   10557
         _ExtentY        =   2910
         _StockProps     =   64
      End
      Begin VB.CommandButton cmd�շ�ִ�п��� 
         Caption         =   "��"
         Height          =   280
         Left            =   5850
         TabIndex        =   185
         TabStop         =   0   'False
         Top             =   5830
         Width           =   280
      End
      Begin VB.TextBox txt�շ�ִ�п��� 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   184
         Top             =   5820
         Width           =   3735
      End
      Begin VB.CheckBox chkInsurePartFee 
         Caption         =   "�൥�ݷֵ��ݽ���ʱ��ֻ��ҽ������ɹ��ĵ����շ�"
         Height          =   195
         Left            =   -74805
         TabIndex        =   20
         Top             =   2430
         Width           =   4470
      End
      Begin VB.CheckBox chkƤ�� 
         Caption         =   "��ȡ���۵��շ�ʱ���Ƥ�Խ��"
         Height          =   195
         Left            =   -74805
         TabIndex        =   15
         Top             =   1700
         Width           =   2820
      End
      Begin VB.Frame fraFeeExe 
         Caption         =   "�շ�ִ�е�"
         Height          =   585
         Left            =   -74760
         TabIndex        =   180
         Top             =   5505
         Width           =   3870
         Begin VB.OptionButton optExe 
            Caption         =   "����ӡ"
            Height          =   180
            Index           =   0
            Left            =   225
            TabIndex        =   181
            Top             =   300
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton optExe 
            Caption         =   "ѡ���Ƿ��ӡ"
            Height          =   180
            Index           =   2
            Left            =   2385
            TabIndex        =   183
            Top             =   300
            Width           =   1455
         End
         Begin VB.OptionButton optExe 
            Caption         =   "�Զ���ӡ"
            Height          =   180
            Index           =   1
            Left            =   1215
            TabIndex        =   182
            Top             =   300
            Width           =   1065
         End
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "ִ���嵥��ӡ����(&7)"
         Height          =   350
         Index           =   6
         Left            =   -70725
         TabIndex        =   175
         Top             =   5775
         Width           =   1950
      End
      Begin VB.Frame fra�˷�ȱʡѡ��ʽ 
         Caption         =   "�˷�ȱʡѡ��ʽ"
         Height          =   570
         Left            =   -74805
         TabIndex        =   40
         Top             =   5880
         Width           =   6075
         Begin VB.OptionButton opt�˷�ȱʡѡ��ʽ 
            Caption         =   "ȱʡȫѡ���˷���Ŀ"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   3840
            TabIndex        =   42
            Top             =   300
            Width           =   2010
         End
         Begin VB.OptionButton opt�˷�ȱʡѡ��ʽ 
            Caption         =   "ȱʡ�����ݺŻ�Ʊ��ѡ���˷���Ŀ"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   41
            Top             =   300
            Value           =   -1  'True
            Width           =   3195
         End
      End
      Begin VB.Frame fraRefundReceipt 
         Caption         =   "�˷ѻص�����"
         Height          =   585
         Left            =   -74760
         TabIndex        =   176
         Top             =   4905
         Width           =   3870
         Begin VB.OptionButton optRefund 
            Caption         =   "�Զ���ӡ"
            Height          =   180
            Index           =   1
            Left            =   1215
            TabIndex        =   179
            Top             =   300
            Width           =   1065
         End
         Begin VB.OptionButton optRefund 
            Caption         =   "ѡ���Ƿ��ӡ"
            Height          =   180
            Index           =   2
            Left            =   2385
            TabIndex        =   178
            Top             =   300
            Width           =   1455
         End
         Begin VB.OptionButton optRefund 
            Caption         =   "����ӡ"
            Height          =   180
            Index           =   0
            Left            =   225
            TabIndex        =   177
            Top             =   300
            Value           =   -1  'True
            Width           =   1020
         End
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "�˷ѻص���ӡ����(&6)"
         Height          =   350
         Index           =   5
         Left            =   -70725
         TabIndex        =   174
         Top             =   5430
         Width           =   1950
      End
      Begin VB.Frame fraDrugNotFee 
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   -74805
         TabIndex        =   170
         Top             =   5625
         Width           =   4125
         Begin VB.OptionButton optDrug 
            Caption         =   "����"
            Height          =   180
            Index           =   2
            Left            =   3450
            TabIndex        =   39
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton optDrug 
            Caption         =   "��ֹ"
            Height          =   180
            Index           =   1
            Left            =   2670
            TabIndex        =   38
            Top             =   15
            Width           =   855
         End
         Begin VB.OptionButton optDrug 
            Caption         =   "�����"
            Height          =   180
            Index           =   0
            Left            =   1770
            TabIndex        =   37
            Top             =   15
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.Label lblDrugNotFee 
            AutoSize        =   -1  'True
            Caption         =   "ҩƷ��ҩ���˷ѷ�ʽ"
            Height          =   180
            Left            =   15
            TabIndex        =   36
            Top             =   15
            Width           =   1620
         End
      End
      Begin VB.CheckBox chkPayKey 
         Caption         =   "ʹ��С���̵ļӼ�(+-)���л�֧����ʽ"
         Height          =   180
         Left            =   -72450
         TabIndex        =   33
         Top             =   5205
         Width           =   3375
      End
      Begin VB.ComboBox cbo���� 
         Height          =   300
         Left            =   -73575
         Style           =   2  'Dropdown List
         TabIndex        =   132
         Top             =   5280
         Width           =   2355
      End
      Begin VB.Frame fra�ɿ���� 
         Caption         =   "�ɿ����������"
         Height          =   975
         Left            =   300
         TabIndex        =   114
         Top             =   4830
         Width           =   5865
         Begin VB.OptionButton opt�ɿ� 
            Caption         =   "�շ�ʱ���������ۼ�"
            Height          =   285
            Index           =   3
            Left            =   2985
            TabIndex        =   130
            Top             =   555
            Width           =   2715
         End
         Begin VB.OptionButton opt�ɿ� 
            Caption         =   $"frmSetExpence.frx":01FC
            Height          =   285
            Index           =   2
            Left            =   2985
            TabIndex        =   117
            Top             =   270
            Width           =   2655
         End
         Begin VB.OptionButton opt�ɿ� 
            Caption         =   "�շ�ʱ���ಡ���ۼ�"
            Height          =   285
            Index           =   1
            Left            =   225
            TabIndex        =   116
            Top             =   555
            Width           =   2715
         End
         Begin VB.OptionButton opt�ɿ� 
            Caption         =   $"frmSetExpence.frx":021A
            Height          =   285
            Index           =   0
            Left            =   225
            TabIndex        =   115
            Top             =   270
            Value           =   -1  'True
            Width           =   3780
         End
      End
      Begin VB.CheckBox chk���������ɿ� 
         Caption         =   "��ȡ���۵��������ɿ�"
         Height          =   180
         Left            =   -74805
         TabIndex        =   32
         Top             =   5220
         Width           =   2160
      End
      Begin VB.Frame fraFeeList 
         Caption         =   "�շѺ�����嵥"
         Height          =   585
         Left            =   -74760
         TabIndex        =   126
         Top             =   4290
         Width           =   3870
         Begin VB.OptionButton optPrint 
            Caption         =   "�Զ���ӡ"
            Height          =   180
            Index           =   1
            Left            =   1215
            TabIndex        =   128
            Top             =   300
            Width           =   1065
         End
         Begin VB.OptionButton optPrint 
            Caption         =   "ѡ���Ƿ��ӡ"
            Height          =   180
            Index           =   2
            Left            =   2385
            TabIndex        =   127
            Top             =   300
            Width           =   1455
         End
         Begin VB.OptionButton optPrint 
            Caption         =   "����ӡ"
            Height          =   180
            Index           =   0
            Left            =   225
            TabIndex        =   129
            Top             =   300
            Value           =   -1  'True
            Width           =   1020
         End
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "ҽ���ص���ӡ����(&5)"
         Height          =   350
         Index           =   4
         Left            =   -70725
         TabIndex        =   124
         Top             =   5070
         Width           =   1950
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "�����ݷ��������ʾ"
         Height          =   195
         Index           =   2
         Left            =   315
         TabIndex        =   63
         Top             =   3750
         Width           =   2280
      End
      Begin VB.Frame fra�����ʾ 
         Caption         =   "�����ʾ"
         Height          =   1155
         Left            =   -74820
         TabIndex        =   118
         Top             =   3240
         Width           =   4455
         Begin VB.OptionButton opt��� 
            Caption         =   "����ʾ����"
            Height          =   180
            Index           =   1
            Left            =   2835
            TabIndex        =   123
            Top             =   810
            Width           =   1215
         End
         Begin VB.OptionButton opt��� 
            Caption         =   "��ʾ�����"
            Height          =   180
            Index           =   0
            Left            =   1440
            TabIndex        =   121
            Top             =   810
            Width           =   1290
         End
         Begin VB.CheckBox chkҩ�� 
            Caption         =   "��ʾ����ҩ�����"
            Height          =   195
            Left            =   150
            TabIndex        =   120
            Top             =   375
            Width           =   1770
         End
         Begin VB.CheckBox chkҩ�� 
            Caption         =   "��ʾ����ҩ����"
            Height          =   195
            Left            =   2250
            TabIndex        =   119
            Top             =   390
            Width           =   1770
         End
         Begin VB.Line lnSplit 
            BorderColor     =   &H00FFFFFF&
            Index           =   0
            X1              =   15
            X2              =   4425
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Line lnSplit 
            BorderColor     =   &H80000000&
            Index           =   1
            X1              =   15
            X2              =   4440
            Y1              =   705
            Y2              =   705
         End
         Begin VB.Label lbl�����ʾ��ʽ 
            AutoSize        =   -1  'True
            Caption         =   "�����ʾ��ʽ"
            Height          =   180
            Left            =   300
            TabIndex        =   122
            Top             =   810
            Width           =   1080
         End
      End
      Begin MSComCtl2.UpDown updƱ������ 
         Height          =   300
         Left            =   -73275
         TabIndex        =   103
         Top             =   3960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Value           =   10
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtƱ������"
         BuddyDispid     =   196656
         OrigLeft        =   1500
         OrigTop         =   3285
         OrigRight       =   1755
         OrigBottom      =   3570
         Max             =   100
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtƱ������ 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   -73740
         TabIndex        =   102
         Text            =   "10"
         Top             =   3960
         Width           =   465
      End
      Begin VB.CheckBox chkMustRegevent 
         Caption         =   "�շ�ʱ��鲡�˹Һſ���"
         Height          =   195
         Left            =   315
         TabIndex        =   68
         ToolTipText     =   "Ҫ�������˸ÿ��ҵĺŲ��ܱ��濪������Ϊ�ÿ��ҵķ��õ���"
         Top             =   4560
         Width           =   2295
      End
      Begin VB.Frame fraRegPrompt 
         Caption         =   "δ�ҺŲ����շ�"
         Height          =   990
         Left            =   4560
         TabIndex        =   110
         Top             =   465
         Visible         =   0   'False
         Width           =   1620
         Begin VB.OptionButton optRegPrompt 
            Caption         =   "��ֹ"
            Height          =   180
            Index           =   2
            Left            =   210
            TabIndex        =   113
            Top             =   720
            Width           =   1020
         End
         Begin VB.OptionButton optRegPrompt 
            Caption         =   "����"
            Height          =   180
            Index           =   0
            Left            =   210
            TabIndex        =   112
            Top             =   270
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton optRegPrompt 
            Caption         =   "����"
            Height          =   180
            Index           =   1
            Left            =   210
            TabIndex        =   111
            Top             =   490
            Width           =   1020
         End
      End
      Begin VB.ComboBox cboAutoSplitBill 
         Height          =   300
         Left            =   1880
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   4800
         Width           =   1170
      End
      Begin VB.CheckBox chkUnPopPriceBill 
         Caption         =   "���������۵�ѡ�񴰿�"
         Height          =   195
         Left            =   3975
         TabIndex        =   67
         Top             =   4530
         Width           =   2160
      End
      Begin VB.CheckBox chkȱʡ�������� 
         Caption         =   "ȱʡ��������"
         Height          =   195
         Left            =   2160
         TabIndex        =   109
         Top             =   2760
         Width           =   1620
      End
      Begin VB.CheckBox chkAutoSplitBill 
         Caption         =   "�շ���ϸ�Զ���              ��ϵ���"
         Height          =   195
         Left            =   315
         TabIndex        =   69
         Top             =   4830
         Width           =   3960
      End
      Begin VB.CheckBox chkAddedItem 
         Caption         =   "δ�Һ�ʱ�Զ������շ���Ŀ"
         Height          =   195
         Left            =   -74805
         TabIndex        =   17
         Top             =   2200
         Width           =   2460
      End
      Begin VB.TextBox txtAddedItem 
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00C00000&
         Height          =   270
         Left            =   -72250
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   2162
         Width           =   1575
      End
      Begin VB.CommandButton cmdAddedItem 
         Caption         =   "��"
         Height          =   280
         Left            =   -70680
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2157
         Width           =   280
      End
      Begin VB.CommandButton cmdDeviceSetup 
         Caption         =   "�豸����(&S)"
         Height          =   350
         Left            =   4605
         TabIndex        =   108
         Top             =   4830
         Width           =   1500
      End
      Begin VB.CheckBox chkPrePayPriority 
         Caption         =   "����ʹ��Ԥ����ɷ�"
         Height          =   195
         Left            =   -74805
         TabIndex        =   16
         Top             =   1950
         Width           =   2340
      End
      Begin VB.CheckBox chkOnlyUnitPatient 
         Caption         =   "ֻ���Һ�Լ��λ����"
         Height          =   195
         Left            =   2160
         TabIndex        =   107
         Top             =   3000
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.TextBox txtNameDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   2925
         MaxLength       =   3
         TabIndex        =   105
         Text            =   "0"
         ToolTipText     =   "0��ʾ����ʱ������ʱ��"
         Top             =   2520
         Width           =   285
      End
      Begin VB.Frame fraShortLine 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   2925
         TabIndex        =   104
         Top             =   2700
         Width           =   285
      End
      Begin VB.Frame fraDoctor 
         Caption         =   "��ʾ������"
         Height          =   885
         Left            =   4560
         TabIndex        =   98
         Top             =   1530
         Width           =   1620
         Begin VB.OptionButton optDoctorKind 
            Caption         =   "������"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   210
            TabIndex        =   100
            Top             =   540
            Width           =   1020
         End
         Begin VB.OptionButton optDoctorKind 
            Caption         =   "������"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   210
            TabIndex        =   99
            Top             =   270
            Value           =   -1  'True
            Width           =   1020
         End
      End
      Begin VB.CheckBox chkLedWelcome 
         Caption         =   "LED��ʾ��ӭ��Ϣ"
         Height          =   225
         Left            =   3975
         TabIndex        =   93
         ToolTipText     =   "�շѴ������벡�˺�,�Ƿ���ʾ��ӭ��Ϣ������"
         Top             =   4050
         Value           =   1  'Checked
         Width           =   1770
      End
      Begin VB.CheckBox chkLedDispDetail 
         Caption         =   "LED��ʾ�շ���ϸ"
         Height          =   225
         Left            =   3975
         TabIndex        =   92
         ToolTipText     =   "�շѴ���,�����շ���Ŀ���Ƿ���ʾ��Ϣ"
         Top             =   3810
         Value           =   1  'Checked
         Width           =   1770
      End
      Begin VB.CheckBox chkRegistInvoice 
         Caption         =   "�Һ�ʱʹ�����շ���ͬ��Ʊ��"
         Height          =   195
         Left            =   -74760
         TabIndex        =   73
         Top             =   3705
         Width           =   2640
      End
      Begin VB.Frame fraLine 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   1695
         TabIndex        =   88
         Top             =   4500
         Width           =   405
      End
      Begin VB.TextBox txtSeekDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   180
         Left            =   1695
         MaxLength       =   4
         TabIndex        =   66
         Text            =   "1"
         Top             =   4305
         Width           =   435
      End
      Begin VB.Frame fra���� 
         Caption         =   "������Դ"
         Height          =   990
         Left            =   285
         TabIndex        =   87
         Top             =   465
         Width           =   1380
         Begin VB.OptionButton opt���� 
            Caption         =   "���ﲡ��"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   210
            TabIndex        =   47
            Top             =   330
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "סԺ����"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   210
            TabIndex        =   48
            Top             =   650
            Width           =   1020
         End
      End
      Begin VB.Frame fra������ҽ�� 
         Caption         =   "������ҽ��"
         Height          =   990
         Left            =   1800
         TabIndex        =   83
         Top             =   465
         Width           =   2685
         Begin VB.OptionButton optSelf 
            Caption         =   "���Һ�ҽ�������������"
            Height          =   195
            Left            =   180
            TabIndex        =   51
            Top             =   720
            Width           =   2280
         End
         Begin VB.OptionButton optDoctor 
            Caption         =   "ͨ������ҽ����ȷ������"
            Height          =   180
            Left            =   180
            TabIndex        =   50
            Top             =   490
            Width           =   2280
         End
         Begin VB.OptionButton optUnit 
            Caption         =   "ͨ�����������ȷ��ҽ��"
            Height          =   180
            Left            =   180
            TabIndex        =   49
            Top             =   270
            Value           =   -1  'True
            Width           =   2280
         End
      End
      Begin VB.CheckBox chkMulti 
         Caption         =   "�շ�ʱ����ͬʱ������ŵ���"
         Height          =   195
         Left            =   315
         TabIndex        =   64
         Top             =   4035
         Width           =   3000
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "��������Ŀ��ʾ����ϼ�"
         Height          =   195
         Index           =   1
         Left            =   315
         TabIndex        =   62
         Top             =   3530
         Width           =   2280
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "���վݷ�Ŀ��ʾ����ϼ�"
         Height          =   195
         Index           =   0
         Left            =   315
         TabIndex        =   61
         Top             =   3300
         Value           =   -1  'True
         Width           =   2280
      End
      Begin VB.CheckBox chk��ȱʡ������ 
         Caption         =   "��ʹ��ȱʡ������"
         Height          =   195
         Left            =   315
         TabIndex        =   59
         Top             =   2760
         Width           =   1740
      End
      Begin VB.CheckBox chk�����俪���� 
         Caption         =   "����Ҫ���뿪����"
         Height          =   195
         Left            =   315
         TabIndex        =   60
         Top             =   3000
         Width           =   1740
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "�շ��嵥��ӡ����(&4)"
         Height          =   350
         Index           =   2
         Left            =   -70725
         TabIndex        =   76
         Top             =   4725
         Width           =   1950
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "�վ�֤����ӡ����(&3)"
         Height          =   350
         Index           =   1
         Left            =   -70725
         TabIndex        =   75
         Top             =   4380
         Width           =   1950
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "�շ�Ʊ�ݴ�ӡ����(&1)"
         Height          =   350
         Index           =   0
         Left            =   -70725
         TabIndex        =   74
         Top             =   3675
         Width           =   1950
      End
      Begin VB.CheckBox chkTime 
         Caption         =   "�����������"
         Height          =   195
         Left            =   -72030
         TabIndex        =   5
         Top             =   810
         Width           =   1380
      End
      Begin VB.CheckBox chkLed 
         Caption         =   "�˹�����LED����"
         Height          =   225
         Left            =   3975
         TabIndex        =   72
         Top             =   3540
         Width           =   1650
      End
      Begin VB.CheckBox chk��ʿ 
         Caption         =   "�����˺���ʿ"
         Height          =   195
         Left            =   -72030
         TabIndex        =   6
         Top             =   1080
         Width           =   1380
      End
      Begin VB.CheckBox chk�ۼ� 
         Caption         =   "��ʾ�տ��ۼ�"
         Height          =   195
         Left            =   -72030
         TabIndex        =   7
         Top             =   1350
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Frame fraTitle 
         Caption         =   "���ع����շ�Ʊ��"
         Height          =   1500
         Left            =   -74775
         TabIndex        =   85
         Top             =   450
         Width           =   6000
         Begin VSFlex8Ctl.VSFlexGrid vsBill 
            Height          =   1155
            Left            =   75
            TabIndex        =   125
            Top             =   255
            Width           =   5790
            _cx             =   10213
            _cy             =   2037
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
            FormatString    =   $"frmSetExpence.frx":0238
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
      Begin VB.TextBox txtDay 
         ForeColor       =   &H80000012&
         Height          =   270
         Left            =   -73665
         MaxLength       =   3
         TabIndex        =   13
         Text            =   "0"
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtMax 
         ForeColor       =   &H80000012&
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   -73665
         MaxLength       =   12
         TabIndex        =   1
         Text            =   "0.00"
         Top             =   525
         Width           =   1335
      End
      Begin VB.CheckBox chkPay 
         Caption         =   "��ҩ���븶��"
         Height          =   195
         Left            =   -72030
         TabIndex        =   4
         Top             =   555
         Value           =   1  'Checked
         Width           =   1380
      End
      Begin VB.Frame fra��� 
         Caption         =   "�����շ����"
         Height          =   4020
         Left            =   -70275
         TabIndex        =   28
         Top             =   480
         Width           =   1485
         Begin VB.ListBox lst�շ���� 
            ForeColor       =   &H00C00000&
            Height          =   3630
            Left            =   105
            Style           =   1  'Checkbox
            TabIndex        =   29
            ToolTipText     =   "�븴ѡ����ʹ�õ��շ����"
            Top             =   255
            Width           =   1275
         End
      End
      Begin VB.ComboBox cbo�ѱ� 
         Height          =   300
         Left            =   -73665
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   915
         Width           =   1350
      End
      Begin VB.ComboBox cbo���㷽ʽ 
         Height          =   300
         Left            =   -73665
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1305
         Width           =   1350
      End
      Begin VB.Frame fraInputItem 
         Caption         =   "�շѻ򻮼�ʱҪ�������Ŀ"
         Height          =   885
         Left            =   285
         TabIndex        =   80
         Top             =   1530
         Width           =   4170
         Begin VB.CheckBox chkҽ�Ƹ��� 
            Caption         =   "ҽ�Ƹ��ʽ"
            Height          =   210
            Left            =   2520
            TabIndex        =   57
            Top             =   540
            Value           =   1  'Checked
            Width           =   1380
         End
         Begin VB.CheckBox chk�Ա� 
            Caption         =   "�Ա�"
            Height          =   210
            Left            =   165
            TabIndex        =   52
            Top             =   270
            Value           =   1  'Checked
            Width           =   660
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "����"
            Height          =   210
            Left            =   2520
            TabIndex        =   55
            Top             =   270
            Value           =   1  'Checked
            Width           =   660
         End
         Begin VB.CheckBox chk�ѱ� 
            Caption         =   "�ѱ�"
            Height          =   210
            Left            =   3240
            TabIndex        =   56
            Top             =   270
            Value           =   1  'Checked
            Width           =   660
         End
         Begin VB.CheckBox chk�Ƿ�Ӱ� 
            Caption         =   "�Ƿ�Ӱ�"
            Height          =   210
            Left            =   1350
            TabIndex        =   53
            Top             =   270
            Value           =   1  'Checked
            Width           =   1020
         End
         Begin VB.CheckBox chk�������� 
            Caption         =   "��������"
            Height          =   210
            Left            =   165
            TabIndex        =   54
            Top             =   540
            Value           =   1  'Checked
            Width           =   1020
         End
         Begin VB.CheckBox chk������ 
            Caption         =   "������"
            Height          =   210
            Left            =   1350
            TabIndex        =   58
            Top             =   540
            Value           =   1  'Checked
            Width           =   840
         End
      End
      Begin MSComCtl2.UpDown udDay 
         Height          =   270
         Left            =   -72570
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1665
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   476
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtDay"
         BuddyDispid     =   196694
         OrigLeft        =   3045
         OrigTop         =   615
         OrigRight       =   3285
         OrigBottom      =   885
         Max             =   32767
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "����֪ͨ����ӡ����(&1)"
         Height          =   350
         Index           =   3
         Left            =   3840
         TabIndex        =   71
         Top             =   2760
         Visible         =   0   'False
         Width           =   2325
      End
      Begin VB.CheckBox chkSeekBill 
         Caption         =   "�Զ���Ѱ����      ���ڵĻ��۵���"
         Height          =   195
         Left            =   315
         TabIndex        =   65
         Top             =   4305
         Width           =   3180
      End
      Begin VB.Frame fra����֪ͨ����ӡ 
         Caption         =   "����֪ͨ����ӡ"
         Height          =   1230
         Left            =   3840
         TabIndex        =   94
         Top             =   3240
         Visible         =   0   'False
         Width           =   2325
         Begin VB.OptionButton optPrintRequisition 
            Caption         =   "ѡ���Ƿ��ӡ"
            Height          =   180
            Index           =   2
            Left            =   135
            TabIndex        =   97
            Top             =   900
            Width           =   1500
         End
         Begin VB.OptionButton optPrintRequisition 
            Caption         =   "����ӡ"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   96
            Top             =   285
            Width           =   1020
         End
         Begin VB.OptionButton optPrintRequisition 
            Caption         =   "�Զ���ӡ"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   95
            Top             =   600
            Value           =   -1  'True
            Width           =   1260
         End
      End
      Begin VB.CheckBox chkSeekName 
         Caption         =   "����ͨ������������ģ������    ���ڵĲ�����Ϣ"
         Height          =   195
         Left            =   315
         TabIndex        =   106
         Top             =   2535
         Width           =   4260
      End
      Begin VB.CheckBox chkƱ������ 
         Caption         =   "Ʊ��ʣ��         ��ʱ��ʼ�����շ�Ա"
         Height          =   285
         Left            =   -74760
         TabIndex        =   101
         Top             =   3960
         Width           =   3450
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfDrugStore 
         Height          =   4695
         Left            =   -74775
         TabIndex        =   131
         Top             =   540
         Width           =   5655
         _cx             =   9975
         _cy             =   8281
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
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSetExpence.frx":0316
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
      Begin VB.Frame fraPrintBill 
         Caption         =   "��ӡ����"
         Height          =   1185
         Left            =   -72000
         TabIndex        =   8
         Top             =   600
         Visible         =   0   'False
         Width           =   1605
         Begin VB.CheckBox chk 
            Caption         =   "����ʱ"
            Height          =   195
            Index           =   0
            Left            =   285
            TabIndex        =   9
            Top             =   300
            Width           =   840
         End
         Begin VB.CheckBox chk 
            Caption         =   "����ʱ"
            Height          =   195
            Index           =   1
            Left            =   285
            TabIndex        =   10
            Top             =   555
            Width           =   840
         End
         Begin VB.CheckBox chk 
            Caption         =   "���ʱ"
            Height          =   195
            Index           =   2
            Left            =   285
            TabIndex        =   11
            Top             =   810
            Width           =   840
         End
      End
      Begin VB.CheckBox chk�շ�ִ�п��� 
         Caption         =   "�����շ�ִ�п���"
         Height          =   210
         Left            =   330
         TabIndex        =   186
         Top             =   5880
         Width           =   1770
      End
      Begin VB.PictureBox picBill 
         BorderStyle     =   0  'None
         Height          =   4635
         Index           =   2
         Left            =   -74850
         ScaleHeight     =   4635
         ScaleWidth      =   5985
         TabIndex        =   134
         Top             =   615
         Width           =   5985
         Begin VB.PictureBox picRuleBack 
            BorderStyle     =   0  'None
            Height          =   2790
            Index           =   0
            Left            =   -15
            ScaleHeight     =   2790
            ScaleWidth      =   6255
            TabIndex        =   143
            Top             =   495
            Visible         =   0   'False
            Width           =   6255
            Begin VB.CheckBox chk��찴���ݷֱ��ӡ 
               Caption         =   "��첡��ÿ�ŵ��ݷֱ��ӡ(�ò���ͬʱӰ�칤������������)"
               Height          =   195
               Left            =   630
               TabIndex        =   145
               Top             =   300
               Width           =   5160
            End
            Begin VB.CheckBox chkAutoAddBookFee 
               Caption         =   "�����շ�ʱ�Զ����չ�����"
               Height          =   195
               Left            =   345
               TabIndex        =   147
               Top             =   825
               Width           =   2460
            End
            Begin VB.CheckBox chkOlnyOneBill 
               Caption         =   "�շ�ÿ�δ�ӡֻ��һ��Ʊ��(�ò���ͬʱӰ�칤������������)"
               Height          =   195
               Left            =   345
               TabIndex        =   146
               Top             =   555
               Width           =   5160
            End
            Begin VB.Frame fraActuallyPrint 
               Height          =   1695
               Left            =   150
               TabIndex        =   148
               Top             =   825
               Width           =   5850
               Begin VB.CheckBox chkErrorItemNotBill 
                  Caption         =   "����ʹ��Ʊ��"
                  Height          =   195
                  Left            =   195
                  TabIndex        =   149
                  Top             =   885
                  Width           =   1740
               End
               Begin VB.OptionButton optBillMode 
                  Caption         =   "��ӡ�վݷ�Ŀ"
                  Height          =   255
                  Index           =   0
                  Left            =   2640
                  TabIndex        =   153
                  Top             =   1245
                  Value           =   -1  'True
                  Width           =   1575
               End
               Begin VB.OptionButton optBillMode 
                  Caption         =   "��ӡ�շ���Ŀ"
                  Height          =   255
                  Index           =   1
                  Left            =   4200
                  TabIndex        =   154
                  Top             =   1245
                  Width           =   1455
               End
               Begin VB.CheckBox chkExcuteDept 
                  Caption         =   "��ִ�п��ҷֱ��ӡ"
                  Height          =   195
                  Left            =   200
                  TabIndex        =   152
                  Top             =   1275
                  Width           =   1980
               End
               Begin MSComCtl2.UpDown updRows 
                  Height          =   300
                  Left            =   4320
                  TabIndex        =   151
                  TabStop         =   0   'False
                  Top             =   825
                  Width           =   255
                  _ExtentX        =   450
                  _ExtentY        =   529
                  _Version        =   393216
                  Value           =   3
                  BuddyControl    =   "txtRowsUD"
                  BuddyDispid     =   196729
                  OrigLeft        =   4440
                  OrigTop         =   825
                  OrigRight       =   4695
                  OrigBottom      =   1125
                  Max             =   100
                  SyncBuddy       =   -1  'True
                  BuddyProperty   =   65547
                  Enabled         =   -1  'True
               End
               Begin VB.TextBox txtRowsUD 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Left            =   4005
                  Locked          =   -1  'True
                  TabIndex        =   150
                  Text            =   "3"
                  Top             =   832
                  Width           =   330
               End
               Begin VB.Label lblRows 
                  AutoSize        =   -1  'True
                  Caption         =   "�շ��վ��д�"
                  Height          =   180
                  Left            =   2850
                  TabIndex        =   160
                  Top             =   892
                  Width           =   1080
               End
               Begin VB.Label lbl 
                  Caption         =   "������������Ʊ����������,Ʊ�����������¹������.��ʵ�ʴ�ӡ������Ʊ������Դ��Ʊ����ƾ���,������߲�һ��,��������������׼ȷ."
                  Height          =   495
                  Index           =   25
                  Left            =   120
                  TabIndex        =   155
                  Top             =   360
                  Width           =   5655
               End
            End
            Begin VB.CheckBox chkBillNO 
               Caption         =   "�����շ�ÿ�ŵ��ݷֱ��ӡ(�ò���ͬʱӰ�칤������������)"
               Height          =   195
               Left            =   345
               TabIndex        =   144
               Top             =   75
               Width           =   5160
            End
         End
         Begin VB.ComboBox cboBillRole 
            Height          =   300
            ItemData        =   "frmSetExpence.frx":03A4
            Left            =   1125
            List            =   "frmSetExpence.frx":03A6
            Style           =   2  'Dropdown List
            TabIndex        =   158
            Top             =   105
            Width           =   3015
         End
         Begin VB.PictureBox picRuleBack 
            BorderStyle     =   0  'None
            Height          =   1035
            Index           =   2
            Left            =   30
            ScaleHeight     =   1035
            ScaleWidth      =   6330
            TabIndex        =   156
            Top             =   435
            Visible         =   0   'False
            Width           =   6330
            Begin VB.Label lblCustomInfor 
               Caption         =   $"frmSetExpence.frx":03A8
               Height          =   570
               Left            =   165
               TabIndex        =   157
               Top             =   375
               Width           =   5670
            End
         End
         Begin VB.PictureBox picRuleBack 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3825
            Index           =   1
            Left            =   15
            ScaleHeight     =   3825
            ScaleWidth      =   6150
            TabIndex        =   135
            Top             =   435
            Visible         =   0   'False
            Width           =   6150
            Begin VB.Frame fraRuleSystem 
               Height          =   3525
               Left            =   0
               TabIndex        =   136
               Top             =   105
               Width           =   5955
               Begin VB.OptionButton optRuleTotal 
                  Caption         =   "��ִ�п��ҷ������"
                  Height          =   240
                  Index           =   2
                  Left            =   2985
                  TabIndex        =   173
                  Top             =   2250
                  Width           =   2025
               End
               Begin VB.OptionButton optRuleTotal 
                  Caption         =   "��ҳ��ӡ����"
                  Height          =   240
                  Index           =   1
                  Left            =   1425
                  TabIndex        =   172
                  Top             =   2250
                  Width           =   1440
               End
               Begin VB.OptionButton optRuleTotal 
                  Caption         =   "������"
                  Height          =   240
                  Index           =   0
                  Left            =   330
                  TabIndex        =   171
                  Top             =   2250
                  Value           =   -1  'True
                  Width           =   1005
               End
               Begin VB.TextBox txtBillRuleNum 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Index           =   2
                  Left            =   2430
                  Locked          =   -1  'True
                  TabIndex        =   167
                  Text            =   "3"
                  Top             =   1875
                  Width           =   330
               End
               Begin VB.TextBox txtBillRuleNum 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Index           =   1
                  Left            =   2430
                  Locked          =   -1  'True
                  TabIndex        =   164
                  Text            =   "3"
                  Top             =   1530
                  Width           =   330
               End
               Begin VB.TextBox txtBillRuleNum 
                  Alignment       =   1  'Right Justify
                  Height          =   300
                  Index           =   0
                  Left            =   2430
                  Locked          =   -1  'True
                  TabIndex        =   161
                  Text            =   "3"
                  Top             =   1155
                  Width           =   345
               End
               Begin VB.CheckBox chkBillRule 
                  Caption         =   "4.���շ�ϸĿ��ҳ"
                  Height          =   180
                  Index           =   3
                  Left            =   285
                  TabIndex        =   140
                  Top             =   1920
                  Width           =   1770
               End
               Begin VB.CheckBox chkBillRule 
                  Caption         =   "3.���վݷ�Ŀ��ҳ"
                  Height          =   180
                  Index           =   2
                  Left            =   270
                  TabIndex        =   139
                  Top             =   1575
                  Width           =   1770
               End
               Begin VB.CheckBox chkBillRule 
                  Caption         =   "2.��ִ�п��ҷ�ҳ"
                  Height          =   180
                  Index           =   1
                  Left            =   270
                  TabIndex        =   138
                  Top             =   1215
                  Width           =   1770
               End
               Begin VB.CheckBox chkBillRule 
                  Caption         =   "1.�����ݷ�ҳ"
                  Height          =   225
                  Index           =   0
                  Left            =   270
                  TabIndex        =   137
                  Top             =   915
                  Width           =   1635
               End
               Begin MSComCtl2.UpDown updBillRuleNum 
                  Height          =   300
                  Index           =   0
                  Left            =   2775
                  TabIndex        =   162
                  TabStop         =   0   'False
                  Top             =   1155
                  Width           =   255
                  _ExtentX        =   450
                  _ExtentY        =   529
                  _Version        =   393216
                  Value           =   1
                  AutoBuddy       =   -1  'True
                  BuddyControl    =   "txtBillRuleNum(0)"
                  BuddyDispid     =   196737
                  BuddyIndex      =   0
                  OrigLeft        =   4440
                  OrigTop         =   825
                  OrigRight       =   4695
                  OrigBottom      =   1125
                  Max             =   100
                  SyncBuddy       =   -1  'True
                  BuddyProperty   =   65547
                  Enabled         =   -1  'True
               End
               Begin MSComCtl2.UpDown updBillRuleNum 
                  Height          =   300
                  Index           =   1
                  Left            =   2760
                  TabIndex        =   165
                  TabStop         =   0   'False
                  Top             =   1530
                  Width           =   255
                  _ExtentX        =   450
                  _ExtentY        =   529
                  _Version        =   393216
                  Value           =   4
                  AutoBuddy       =   -1  'True
                  BuddyControl    =   "txtBillRuleNum(1)"
                  BuddyDispid     =   196737
                  BuddyIndex      =   1
                  OrigLeft        =   4440
                  OrigTop         =   825
                  OrigRight       =   4695
                  OrigBottom      =   1125
                  Max             =   100
                  SyncBuddy       =   -1  'True
                  BuddyProperty   =   65547
                  Enabled         =   -1  'True
               End
               Begin MSComCtl2.UpDown updBillRuleNum 
                  Height          =   300
                  Index           =   2
                  Left            =   2760
                  TabIndex        =   168
                  TabStop         =   0   'False
                  Top             =   1875
                  Width           =   255
                  _ExtentX        =   450
                  _ExtentY        =   529
                  _Version        =   393216
                  Value           =   20
                  AutoBuddy       =   -1  'True
                  BuddyControl    =   "txtBillRuleNum(2)"
                  BuddyDispid     =   196737
                  BuddyIndex      =   2
                  OrigLeft        =   4440
                  OrigTop         =   825
                  OrigRight       =   4695
                  OrigBottom      =   1125
                  Max             =   100
                  SyncBuddy       =   -1  'True
                  BuddyProperty   =   65547
                  Enabled         =   -1  'True
               End
               Begin VB.Label lblBillRuleNum 
                  AutoSize        =   -1  'True
                  Caption         =   "��ÿ����   ���շ�ϸĿ��һҳ"
                  Height          =   180
                  Index           =   2
                  Left            =   2025
                  TabIndex        =   169
                  Top             =   1935
                  Width           =   2430
               End
               Begin VB.Label lblBillRuleNum 
                  AutoSize        =   -1  'True
                  Caption         =   "��ÿ����   ���վݷ�Ŀ��һҳ"
                  Height          =   180
                  Index           =   1
                  Left            =   2025
                  TabIndex        =   166
                  Top             =   1590
                  Width           =   2430
               End
               Begin VB.Label lblBillRuleNum 
                  AutoSize        =   -1  'True
                  Caption         =   "��ÿ����   ��ִ�п��ҷ�һҳ"
                  Height          =   180
                  Index           =   0
                  Left            =   2025
                  TabIndex        =   163
                  Top             =   1215
                  Width           =   2430
               End
               Begin VB.Label lblInfor 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  BorderStyle     =   1  'Fixed Single
                  ForeColor       =   &H80000008&
                  Height          =   540
                  Left            =   90
                  TabIndex        =   142
                  Top             =   2880
                  Width           =   5760
               End
               Begin VB.Label lblRuleSystem 
                  Caption         =   "������������Ʊ����������,Ʊ�����������¹������.��ʵ�ʴ�ӡ�������շѻ��۵�����,����ֹ�¼����õ���,�����ѵ��������㽫��׼ȷ."
                  Height          =   585
                  Left            =   180
                  TabIndex        =   141
                  Top             =   330
                  Width           =   5730
               End
            End
         End
         Begin VB.Label lblBillRole 
            AutoSize        =   -1  'True
            Caption         =   "Ʊ�ݷ������"
            Height          =   180
            Left            =   15
            TabIndex        =   159
            Top             =   165
            Width           =   1080
         End
      End
      Begin VB.CheckBox chkסԺ�������շ� 
         Caption         =   "סԺ���˰������շ�"
         Height          =   180
         Left            =   -74805
         TabIndex        =   30
         Top             =   4980
         Width           =   2025
      End
      Begin VB.Frame fra���� 
         Caption         =   " ������ҩ������� "
         Height          =   1230
         Left            =   -74805
         TabIndex        =   24
         Top             =   3750
         Visible         =   0   'False
         Width           =   4440
         Begin VB.ListBox lst��ҩ�� 
            ForeColor       =   &H00C00000&
            Height          =   690
            Left            =   2955
            Style           =   1  'Checkbox
            TabIndex        =   27
            Top             =   465
            Width           =   1350
         End
         Begin VB.ListBox lst��ҩ�� 
            ForeColor       =   &H00C00000&
            Height          =   690
            Left            =   1560
            Style           =   1  'Checkbox
            TabIndex        =   26
            Top             =   465
            Width           =   1350
         End
         Begin VB.ListBox lst��ҩ�� 
            ForeColor       =   &H00C00000&
            Height          =   690
            Left            =   165
            Style           =   1  'Checkbox
            TabIndex        =   25
            Top             =   465
            Width           =   1350
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ҩ��"
            Height          =   180
            Left            =   2955
            TabIndex        =   91
            Top             =   255
            Width           =   540
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ҩ��"
            Height          =   180
            Left            =   1560
            TabIndex        =   90
            Top             =   255
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ҩ��"
            Height          =   180
            Left            =   165
            TabIndex        =   89
            Top             =   255
            Width           =   540
         End
      End
      Begin VB.Label lblDay 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ȡ�����۳���                 ��δ����Ļ��۵�"
         Height          =   180
         Left            =   -74805
         TabIndex        =   12
         Top             =   1680
         Width           =   4050
      End
      Begin VB.Label lbl���ϲ��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ȱʡ���ϲ���"
         Height          =   180
         Left            =   -74775
         TabIndex        =   133
         Top             =   5340
         Width           =   1080
      End
      Begin VB.Label lblMax 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���������"
         Height          =   180
         Left            =   -74820
         TabIndex        =   84
         Top             =   585
         Width           =   1080
      End
      Begin VB.Label lbl�ѱ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ȱʡ���˷ѱ�"
         Height          =   180
         Left            =   -74820
         TabIndex        =   82
         Top             =   975
         Width           =   1080
      End
      Begin VB.Label lbl���㷽ʽ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ȱʡ���㷽ʽ"
         Height          =   180
         Left            =   -74805
         TabIndex        =   81
         Top             =   1365
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmSetExpence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Public mbytInFun As Byte '0=�շ�,1=����,2=�������
Public mstrPrivs As String
Public mlngModul As Long
Public mblnSetDrugStore As Boolean
Private mblnNotClick As Boolean

Private Sub cboBillRole_Click()
     '56963
      Call SetBillNoRule
End Sub
Private Function GetPrintListHaveData() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡƱ�ݴ�ӡ��ϸ�Ƿ�������
    '����:�����ݷ���true,���򷵻�False
    '����:���˺�
    '����:2013-05-17 14:24:40
    '˵��:56963
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errHandle
    strSQL = "Select 1 From Ʊ�ݴ�ӡ��ϸ where Rownum<=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    GetPrintListHaveData = rsTemp.RecordCount >= 1
    rsTemp.Close: Set rsTemp = Nothing
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub ShowRuleInfor()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾƱ�ŵķ������
    '����:���˺�
    '����:2013-03-26 14:14:08
    '����:56963
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInfor As String, i As Integer
    Dim strName As String
    
    On Error GoTo errHandle
    strInfor = ""
    If chkBillRule(0).Value = 1 Then
            strInfor = strInfor & "+ NO"
    End If
    For i = 1 To 3
        If chkBillRule(i).Value = 1 Then
            strName = Switch(i = 1, "ִ�п���", i = 2, "�վݷ�Ŀ", True, "�վ�ϸĿ")
            strInfor = strInfor & "+" & strName & "(" & txtBillRuleNum(i - 1).Text & ")"
        End If
    Next
    If strInfor <> "" Then strInfor = Mid(strInfor, 2)
    lblInfor.Caption = strInfor
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub chkAddedItem_Click()
    Dim i As Long
    If chkAddedItem.Value = 1 Then
        If txtAddedItem.Text = "" And Me.Visible Then cmdAddedItem_Click
    
        '����������
        For i = 0 To lst�շ����.ListCount - 1
            If lst�շ����.ItemData(i) = Asc("Z") Then
                If lst�շ����.Selected(i) = False Then lst�շ����.Selected(i) = True
            End If
        Next
    Else
        txtAddedItem.Text = ""
    End If
End Sub

Private Sub chkAutoSplitBill_Click()
    cboAutoSplitBill.Enabled = chkAutoSplitBill.Value = 1 And cboAutoSplitBill.Tag = "1"
End Sub

Private Sub chkBillNO_Click()
    chk��찴���ݷֱ��ӡ.Enabled = (chkBillNO.Value = vbChecked)
End Sub

Private Sub chkBillRule_Click(Index As Integer)
    '56963
    If Index <> 0 And chkBillRule(Index).Value = 1 Then
        If Val(txtBillRuleNum(Index - 1).Text) = 0 Then
            updBillRuleNum(Index - 1).Value = Val(txtBillRuleNum(Index - 1).Tag)    '�ָ�ȱʡֵ
        End If
    End If
    Call SetBillRuleEnable
    Call ShowRuleInfor
    If Not optRuleTotal(2).Visible Then
         If optRuleTotal(2).Value Then optRuleTotal(0).Value = True
    End If
End Sub
Private Sub SetBillRuleEnable()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ʊ�ݷ������,������Ӧ�ؼ���Enabled����
    '����:���˺�
    '����:2013-03-26 17:55:47
    '����:56963
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intIndex As Integer, blnEnable As Boolean
    On Error GoTo errHandle
    '��������(0-������;1-��ҳ����(����1ҳ����),2-�������(ѡ����ϸʱ��Ч))
    '1.�������:�з�����ϸʱ,ͬʱ���ڹ�ѡִ�п��һ����վݷ�Ŀ�򰴵��ݲŻ���ڷ��������,�Ż���ڻ�����
    blnEnable = chkBillRule(3).Enabled And chkBillRule(3).Value = 1 And (chkBillRule(2).Value = 1 Or chkBillRule(1).Value = 1 Or chkBillRule(0).Value = 1)
    optRuleTotal(2).Visible = blnEnable
    optRuleTotal(2).Enabled = blnEnable
    '2.��ҳ����:���������óɻ��ܶ�
    optRuleTotal(1).Enabled = chkBillRule(3).Enabled
    optRuleTotal(0).Enabled = chkBillRule(3).Enabled
    
    '���÷����������
    If chkBillRule(0).Value = 1 Then
        optRuleTotal(2).Caption = "�����ݺŷ������"
    ElseIf chkBillRule(1).Value = 1 Then
        optRuleTotal(2).Caption = "��ִ�п��ҷ������"
    ElseIf chkBillRule(3).Value = 1 Then
        optRuleTotal(2).Caption = "���վݷ�Ŀ�������"
    ElseIf chkBillRule(3).Value = 1 Then
        optRuleTotal(2).Caption = "��������������"
    End If
    For intIndex = 1 To 3
        txtBillRuleNum(intIndex - 1).Enabled = chkBillRule(intIndex).Value = 1 And chkBillRule(intIndex).Enabled
        updBillRuleNum(intIndex - 1).Enabled = txtBillRuleNum(intIndex - 1).Enabled
        lblBillRuleNum(intIndex - 1).Enabled = txtBillRuleNum(intIndex - 1).Enabled
    Next
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 

End Sub

Private Sub chkOnePatiPrint_Click()
  With vsBillFormat
        .ColHidden(.ColIndex("�����˲���Ʊ�ݸ�ʽ")) = chkOnePatiPrint.Value <> 1
    End With

End Sub

Private Sub chkƱ������_Click()
    txtƱ������.Enabled = chkƱ������.Enabled And chkƱ������.Value = 1
    updƱ������.Enabled = txtƱ������.Enabled
End Sub

Private Sub chk�շ�ִ�п���_Click()
    If mblnNotClick Then Exit Sub
    If chk�շ�ִ�п���.Value = vbChecked Then
        cmd�շ�ִ�п���.Enabled = True
        Call cmd�շ�ִ�п���_Click
    Else
        txt�շ�ִ�п���.Text = ""
        txt�շ�ִ�п���.Tag = ""
        cmd�շ�ִ�п���.Enabled = False
    End If
End Sub

Private Sub cmdAddedItem_Click()
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    strSQL = "Select ID, ����, ����, ���㵥λ, ˵��" & vbNewLine & _
            "From �շ���ĿĿ¼" & vbNewLine & _
            "Where ��� = 'Z' And Nvl(�Ƿ���, 0) = 0 And ������� In(1,3)" & vbNewLine & _
            "Order By ����"
    On Error GoTo errH
    Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "(����)�շ���Ŀ")
    If Not rsTmp Is Nothing Then
        txtAddedItem.Text = rsTmp!����
        txtAddedItem.Tag = rsTmp!ID
        If chkAddedItem.Value = 0 Then chkAddedItem.Value = 1
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdDeviceSetup_Click()
    Dim lngModule As Long
    Select Case mbytInFun
    Case 0
        lngModule = 1121
    Case 1
        lngModule = 1120
    Case 2
        lngModule = 1122
    End Select
    Call zlCommFun.DeviceSetup(Me, 100, lngModule)
End Sub
Private Sub cbo�ѱ�_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then cbo�ѱ�.ListIndex = -1
End Sub

Private Sub chkMulti_Click()
    If chkMulti.Value = 0 Then
        chkSeekBill.Value = 0
        chkSeekBill.Enabled = False
        chkUnPopPriceBill.Value = 0
        chkUnPopPriceBill.Enabled = False
        
        chkAutoSplitBill.Value = 0
        chkAutoSplitBill.Enabled = False
    Else
        chkSeekBill.Enabled = True And chkSeekBill.Tag = "1"
        chkUnPopPriceBill.Enabled = chkSeekBill.Value = 1 And chkUnPopPriceBill.Tag = "1"
        chkAutoSplitBill.Enabled = True And chkAutoSplitBill.Tag = "1"
    End If
    cboAutoSplitBill.Enabled = chkAutoSplitBill.Enabled And cboAutoSplitBill.Tag = "1"
End Sub

Private Sub chkSeekBill_Click()
    txtSeekDays.Enabled = chkSeekBill.Value = 1 And txtSeekDays.Tag = "1"
    If Visible And txtSeekDays.Enabled And txtSeekDays.Visible Then
        txtSeekDays.SetFocus
    End If
    chkUnPopPriceBill.Enabled = chkSeekBill.Value = 1 And chkUnPopPriceBill.Tag = "1"
    If chkSeekBill.Value = 0 Then chkUnPopPriceBill.Value = 0
End Sub

Private Sub chkSeekName_Click()
    txtNameDays.Enabled = chkSeekName.Value = 1 And txtNameDays.Tag = "1"
    chkOnlyUnitPatient.Enabled = chkSeekName.Value = 1 And chkOnlyUnitPatient.Tag = "1"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name & "1"
End Sub
Private Sub SaveInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���淢Ʊ���Ʊ��
    '����:���˺�
    '����:2011-04-28 18:16:48
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHavePrivs As Boolean, strValue As String, strOnePatiPrintValue As String
    Dim i As Long
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
    zlDatabase.SetPara "�����շ�Ʊ������", strValue, glngSys, mlngModul, blnHavePrivs
    '�����շѸ�ʽ
    
    Dim strPrintMode As String
    '�����շѸ�ʽ
    strValue = "": strPrintMode = ""
    With vsBillFormat
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) <> "" Then
                strValue = strValue & "|" & Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) & "," & Val(.TextMatrix(i, .ColIndex("�շ�Ʊ�ݸ�ʽ")))
                strOnePatiPrintValue = strOnePatiPrintValue & "|" & Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) & "," & Val(.TextMatrix(i, .ColIndex("�����˲���Ʊ�ݸ�ʽ")))
                strPrintMode = strPrintMode & "|" & Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("�շѴ�ӡ��ʽ")), 1))
            End If
        Next
        If strValue <> "" Then strValue = Mid(strValue, 2)
        If strOnePatiPrintValue <> "" Then strOnePatiPrintValue = Mid(strOnePatiPrintValue, 2)
        If strPrintMode <> "" Then strPrintMode = Mid(strPrintMode, 2)
        zlDatabase.SetPara "�շѷ�Ʊ��ʽ", strValue, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "�����˲���Ʊ��ʽ", strOnePatiPrintValue, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "���ʷ�Ʊ��ʽ", strValue, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "�շѷ�Ʊ��ӡ��ʽ", strPrintMode, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "�����˲���Ʊ�����ֽ������", IIf(chkOnePatiPrint.Value = 1, 1, 0), glngSys, mlngModul, blnHavePrivs
    End With
    
    '�����˷Ѹ�ʽ
    strValue = "": strPrintMode = ""
    With vsDelBillFormat
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) <> "" Then
                strValue = strValue & "|" & Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) & "," & Val(.TextMatrix(i, .ColIndex("�˷�Ʊ�ݸ�ʽ")))
                strPrintMode = strPrintMode & "|" & Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("�˷Ѵ�ӡ��ʽ")), 1))
            End If
        Next
        If strValue <> "" Then strValue = Mid(strValue, 2)
        If strPrintMode <> "" Then strPrintMode = Mid(strPrintMode, 2)
        zlDatabase.SetPara "�˷ѷ�Ʊ��ʽ", strValue, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "�˷ѷ�Ʊ��ӡ��ʽ", strPrintMode, glngSys, mlngModul, blnHavePrivs
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
    If cboBillRole.ListIndex = 1 Then
        If chkBillRule(0).Value = 0 And chkBillRule(1).Value = 0 And chkBillRule(2).Value = 0 And chkBillRule(3).Value = 0 Then
            MsgBox "ע��:" & vbCrLf & "    Ʊ�ݺŷ�����򰴡�" & cboBillRole.Text & "���ı�������һ�ַ������,����!", vbInformation + vbOKOnly
            stab.Tab = 3
            If chkBillRule(0).Enabled And chkBillRule(0).Visible Then chkBillRule(0).SetFocus
            Exit Function
        End If
    End If
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub cmdOK_Click()
    Dim strValue As String, i As Long
    Dim str��ҩ������ As String, str��ҩ������ As String, str��ҩ������ As String
    Dim lngȱʡ��ҩ�� As Long, lngȱʡ��ҩ�� As Long, lngȱʡ��ҩ�� As Long, lngȱʡ���ϲ��� As Long
    
    'a.���ݼ��
    '--------------------------------------------------------------
    'b.����ע���洢��ģ�����
    '------------------------------------------------------------------------------------------------------------------
    Dim blnHavePrivs As Boolean
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    If isValied = False Then Exit Sub
     
    'c.���ݿ�洢��ģ�����
    '----------------------------------------------------------------------------------------
    If Not mblnSetDrugStore Then
        For i = lst�շ����.ListCount - 1 To 0 Step -1
            If lst�շ����.Selected(i) Then strValue = strValue & "'" & Chr(lst�շ����.ItemData(i)) & "',"
        Next
        If strValue <> "" Then strValue = Left(strValue, Len(strValue) - 1)
        zlDatabase.SetPara "�շ����", strValue, glngSys, mlngModul, blnHavePrivs
        
        If mbytInFun <> 2 Then
            zlDatabase.SetPara "ȱʡ�ѱ�", cbo�ѱ�.Text, glngSys, mlngModul, blnHavePrivs
        End If
        If mbytInFun = 0 Then
            Call SaveInvoice
            zlDatabase.SetPara "ȱʡ���㷽ʽ", cbo���㷽ʽ.Text, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "�ҺŹ����շ�Ʊ��", chkRegistInvoice.Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "�ֹ�����", chkLed.Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "LED��ʾ�շ���ϸ", chkLedDispDetail.Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "LED��ʾ��ӭ��Ϣ", chkLedWelcome.Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "ҽ��������ȱʡ��λ", chkҽ��������ȱʡ��λ.Value, glngSys, mlngModul, blnHavePrivs
        End If
        
        On Error Resume Next
        zlDatabase.SetPara "��ҩ����", chkPay.Value, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "�������", chkTime.Value, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "��ʾ��ʿ", chk��ʿ.Value, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "ҩƷ��λ", IIf(opt��λ(0).Value, 0, 1), glngSys, mlngModul, blnHavePrivs
    End If
    
    If gbln���뷢ҩ Then
        strValue = ""
        For i = 0 To lst��ҩ��.ListCount - 1
            If lst��ҩ��.Selected(i) Then
                strValue = strValue & "," & lst��ҩ��.ItemData(i)
            End If
        Next
        zlDatabase.SetPara "��ҩ��ѡ��", Mid(strValue, 2), glngSys, mlngModul, blnHavePrivs
        strValue = ""
        For i = 0 To lst��ҩ��.ListCount - 1
            If lst��ҩ��.Selected(i) Then
                strValue = strValue & "," & lst��ҩ��.ItemData(i)
            End If
        Next
        zlDatabase.SetPara "��ҩ��ѡ��", Mid(strValue, 2), glngSys, mlngModul, blnHavePrivs
        strValue = ""
        For i = 0 To lst��ҩ��.ListCount - 1
            If lst��ҩ��.Selected(i) Then
                strValue = strValue & "," & lst��ҩ��.ItemData(i)
            End If
        Next
        zlDatabase.SetPara "��ҩ��ѡ��", Mid(strValue, 2), glngSys, mlngModul, blnHavePrivs
    Else
        With vsfDrugStore
            For i = 1 To vsfDrugStore.Rows - 1
                If (mbytInFun = 0 Or mbytInFun = 1) And .TextMatrix(i, .ColIndex("����")) <> "�Զ�����" And .TextMatrix(i, .ColIndex("����")) <> "" Then
                    Select Case .TextMatrix(i, 0)
                        Case "��ҩ��"
                            str��ҩ������ = str��ҩ������ & "," & .RowData(i) & ":" & .TextMatrix(i, .ColIndex("����"))
                        Case "��ҩ��"
                            str��ҩ������ = str��ҩ������ & "," & .RowData(i) & ":" & .TextMatrix(i, .ColIndex("����"))
                        Case "��ҩ��"
                            str��ҩ������ = str��ҩ������ & "," & .RowData(i) & ":" & .TextMatrix(i, .ColIndex("����"))
                    End Select
                End If
                
                If Abs(Val(.TextMatrix(i, .ColIndex("ȱʡ")))) = 1 Then
                    Select Case .TextMatrix(i, .ColIndex("���"))
                        Case "��ҩ��"
                            lngȱʡ��ҩ�� = .RowData(i)
                        Case "��ҩ��"
                            lngȱʡ��ҩ�� = .RowData(i)
                        Case "��ҩ��"
                            lngȱʡ��ҩ�� = .RowData(i)
                    End Select
                End If
            Next
        End With
        If cbo����.ListIndex <> -1 Then
            lngȱʡ���ϲ��� = cbo����.ItemData(cbo����.ListIndex)
        End If
        
        
        If mbytInFun = 0 Or mbytInFun = 1 Then
            str��ҩ������ = Mid(str��ҩ������, 2)
            str��ҩ������ = Mid(str��ҩ������, 2)
            str��ҩ������ = Mid(str��ҩ������, 2)
            zlDatabase.SetPara "��ҩ������", str��ҩ������, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "��ҩ������", str��ҩ������, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "��ҩ������", str��ҩ������, glngSys, mlngModul, blnHavePrivs
        End If
        
        zlDatabase.SetPara "ȱʡ��ҩ��", lngȱʡ��ҩ��, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "ȱʡ��ҩ��", lngȱʡ��ҩ��, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "ȱʡ��ҩ��", lngȱʡ��ҩ��, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "ȱʡ���ϲ���", lngȱʡ���ϲ���, glngSys, mlngModul, blnHavePrivs
                    
                    
        zlDatabase.SetPara "��ʾ����ҩ�����", chkҩ��.Value, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "��ʾ����ҩ����", chkҩ��.Value, glngSys, mlngModul, blnHavePrivs
        If mbytInFun <> 0 Then
            zlDatabase.SetPara "�����ʾ��ʽ", IIf(opt���(0).Value, 0, 1), glngSys, mlngModul, blnHavePrivs
        End If
    End If
    
        
    If Not mblnSetDrugStore Then
        zlDatabase.SetPara "����ҽ��", IIf(optDoctor.Value, 0, IIf(optUnit.Value, 1, 2)), glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "��������ʾ��ʽ", IIf(optDoctorKind(0).Value, 1, 2), glngSys, mlngModul, blnHavePrivs
        
        zlDatabase.SetPara "����ģ������", chkSeekName.Value, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "������������", Val(txtNameDays.Text), glngSys, mlngModul, blnHavePrivs
        '92727
        zlDatabase.SetPara "����¼������ʹ�õĿ�����", IIf(chk������.Value = 1, "1", "0"), glngSys, mlngModul, blnHavePrivs
            
        If mbytInFun = 0 Or mbytInFun = 1 Then
            zlDatabase.SetPara "������Դ", IIf(opt����(0).Value, 1, 2), glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "�����", txtMax.Text, glngSys, mlngModul, blnHavePrivs
            
            zlDatabase.SetPara "�Ա�", chk�Ա�.Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "����", chk����.Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "�ѱ�", chk�ѱ�.Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "ҽ�Ƹ���", chkҽ�Ƹ���.Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "�Ӱ�", chk�Ƿ�Ӱ�.Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "��������", chk��������.Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "������", chk������.Value, glngSys, mlngModul, blnHavePrivs
                    
            zlDatabase.SetPara "����Ҫ���뿪����", chk�����俪����.Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "��ʹ��ȱʡ������", chk��ȱʡ������.Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "ȱʡ��������", chkȱʡ��������.Value, glngSys, mlngModul, blnHavePrivs
            
            zlDatabase.SetPara "����ϼƷ�ʽ", IIf(opt����(0).Value, 0, IIf(opt����(1).Value, 1, 2)), glngSys, mlngModul, blnHavePrivs '34179
            
            '���˺� ����:27663 ����:2010-01-27 11:17:48
            zlDatabase.SetPara "סԺ���˰������շ�", IIf(chkסԺ�������շ�.Value = 1, "1", "0"), glngSys, mlngModul, blnHavePrivs
            '���˺� ����:39253
            zlDatabase.SetPara "��ȡ���ۺ������ɿ�", IIf(chk���������ɿ�.Value = 1, "1", "0"), glngSys, mlngModul, blnHavePrivs
            '47457
            zlDatabase.SetPara "ʹ�üӼ��л�֧����ʽ", IIf(chkPayKey.Value = 1, "1", "0"), glngSys, mlngModul, blnHavePrivs
            '47400
            zlDatabase.SetPara "ҩƷ��ҩ�˷ѷ�ʽ", IIf(optDrug(0).Value, 0, IIf(optDrug(1).Value, "1", "2")), glngSys, mlngModul, blnHavePrivs
            '87489
            zlDatabase.SetPara "�˷�ȱʡѡ��ʽ", IIf(opt�˷�ȱʡѡ��ʽ(0).Value, 0, 1), glngSys, mlngModul, blnHavePrivs
            '86853
            zlDatabase.SetPara "ˢ��ȱʡ������", IIf(optSetMoneyMode(0).Value, 0, IIf(optSetMoneyMode(1).Value, 1, 2)), glngSys, 1151, blnHavePrivs
            
            If mbytInFun = 0 Then
                zlDatabase.SetPara "��ʾ�ۼ�", chk�ۼ�.Value, glngSys, mlngModul, blnHavePrivs
                zlDatabase.SetPara "���Ƥ�Խ��", chkƤ��.Value, glngSys, mlngModul, blnHavePrivs
                zlDatabase.SetPara "����ʹ��Ԥ����", chkPrePayPriority.Value, glngSys, mlngModul, blnHavePrivs
                '120836
                zlDatabase.SetPara "��ֹȡ���ҺŻ��۵�", IIf(chk��ֹȡ���Һŵ�.Value = 1, "1", "0"), glngSys, mlngModul, blnHavePrivs
                '���˺�:22343:51670
                zlDatabase.SetPara "�շѽɿ��������", IIf(opt�ɿ�(0).Value = True, 0, IIf(opt�ɿ�(1).Value = True, 1, IIf(opt�ɿ�(2).Value = True, 2, 3))), glngSys, mlngModul, blnHavePrivs
                '91665
                zlDatabase.SetPara "ֻ��ҽ������ɹ������շ�", chkInsurePartFee.Value, glngSys, mlngModul, blnHavePrivs
                '96357
                zlDatabase.SetPara "�����շ�ִ�п���", txt�շ�ִ�п���.Tag, glngSys, mlngModul, blnHavePrivs
                
                If chkAddedItem.Value = 1 And Val(txtAddedItem.Tag) <> 0 Then
                    zlDatabase.SetPara "�Զ����չҺŷ�", txtAddedItem.Tag & ";" & txtAddedItem.Text, glngSys, mlngModul, blnHavePrivs
                Else
                    zlDatabase.SetPara "�Զ����չҺŷ�", "", glngSys, mlngModul, blnHavePrivs
                End If
'                zlDatabase.SetPara "��ʾ������", chkShowError.Value, glngSys, mlngModul, blnHavePrivs
                zlDatabase.SetPara "�൥���շ�", chkMulti.Value, glngSys, mlngModul, blnHavePrivs
                
                zlDatabase.SetPara "��Ѱ���۵���", chkSeekBill.Value, glngSys, mlngModul, blnHavePrivs
                zlDatabase.SetPara "��Ѱ��������", Val(txtSeekDays.Text), glngSys, mlngModul, blnHavePrivs
                zlDatabase.SetPara "���������۵�ѡ��", chkUnPopPriceBill.Value, glngSys, mlngModul, blnHavePrivs
                    
                zlDatabase.SetPara "��鲡�˹Һſ���", chkMustRegevent.Value, glngSys, mlngModul, blnHavePrivs
                For i = 0 To optRegPrompt.UBound
                    If optRegPrompt(i).Value Then
                        zlDatabase.SetPara "δ�ҺŲ����շ�", i, glngSys, mlngModul, blnHavePrivs
                    End If
                Next
                zlDatabase.SetPara "�Զ���ϵ���", IIf(chkAutoSplitBill.Value = 1, cboAutoSplitBill.ListIndex + 1, 0), glngSys, mlngModul
               For i = 0 To optPrint.UBound
                    If optPrint(i).Value Then
                        zlDatabase.SetPara "�շ��嵥��ӡ��ʽ", i, glngSys, mlngModul, blnHavePrivs
                    End If
                Next
                For i = 0 To optRefund.UBound
                    If optRefund(i).Value Then
                        zlDatabase.SetPara "�˷ѻص���ӡ��ʽ", i, glngSys, mlngModul, blnHavePrivs
                    End If
                Next
                '62982:���ϴ�,2015/08/25,�շ�ִ�е�
                For i = 0 To optExe.UBound
                    If optExe(i).Value Then
                        zlDatabase.SetPara "�շ�ִ�е���ӡ��ʽ", i, glngSys, mlngModul, blnHavePrivs
                    End If
                Next
                
                '���˺� ����:26948 ����:2009-12-28 16:54:11
                zlDatabase.SetPara "Ʊ��ʣ��X��ʱ��ʼ�����շ�Ա", IIf(chkƱ������.Value = 1, "1", "0") & "|" & Val(txtƱ������.Text), glngSys, mlngModul, blnHavePrivs
            
            Else
                zlDatabase.SetPara "ȡ�����۵�", Val(txtDay.Text), glngSys, mlngModul, blnHavePrivs
                For i = 0 To optPrintRequisition.UBound
                    If optPrintRequisition(i).Value Then
                        zlDatabase.SetPara "����֪ͨ����ӡ��ʽ", i, glngSys, mlngModul, blnHavePrivs
                    End If
                Next
            End If
        ElseIf mbytInFun = 2 Then
            zlDatabase.SetPara "ֻ���Һ�Լ��λ����", chkOnlyUnitPatient.Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "���ʴ�ӡ", chk(0).Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "���۴�ӡ", chk(1).Value, glngSys, mlngModul, blnHavePrivs
            zlDatabase.SetPara "��˴�ӡ", chk(2).Value, glngSys, mlngModul, blnHavePrivs
        End If
    End If
    Call SaveBillRulePara '56963
    Call InitLocPar(Choose(mbytInFun + 1, 1121, 1120, 1122))     '��Ҫ��Ҫ�ض��浽����ע���Ĳ���,�������ݿ�Ĳ����ڱ���ʱ���ض�
    gblnOK = True
    Unload Me
End Sub

Private Sub cmdPrintSetup_Click(Index As Integer)
    Select Case Index
        Case 0 '����ҽ�Ʒ��շ�
            If gblnBillPrint Then
                Call gobjBillPrint.zlConfigure
            Else
                If glngSys Like "8??" Then
                    Call ReportPrintSet(gcnOracle, glngSys, "ZL8_BILL_1121_1", Me)
                Else
                    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_1", Me)
                End If
            End If
        Case 1 '�������֤��
            If glngSys Like "8??" Then
                Call ReportPrintSet(gcnOracle, glngSys, "ZL8_BILL_1121_2", Me)
            Else
                Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_2", Me)
            End If
        Case 2 '�����շ��嵥
            If glngSys Like "8??" Then
                Call ReportPrintSet(gcnOracle, glngSys, "ZL8_BILL_1121_3", Me)
            Else
                Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me)
            End If
        Case 3 '����֪ͨ��
            Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1120", Me)
        Case 4 'ҽ���ص�����
            Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_4", Me)
        Case 5  '�˷ѻص�����
            Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_5", Me)
        '62982:���ϴ�,2015/08/25,�շ�ִ�е�
        Case 6  '�շ�ִ�е�����
            Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_6", Me)
        Case 7  '����ҽ�Ʒ��շ�(��Ʊ)
            If gblnBillPrint Then
                Call gobjBillPrint.zlConfigure
            Else
                If glngSys Like "8??" Then
                    Call ReportPrintSet(gcnOracle, glngSys, "ZL8_BILL_1121_7", Me)
                Else
                    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_7", Me)
                End If
            End If
    End Select
End Sub

Private Sub SetStockCheck()
'����:���÷��뷢ҩģʽ�¼��ָ��ҩ���Ŀ��
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    Dim str��ҩID As String, str��ҩID As String, str��ҩID As String
    
    On Error GoTo errH
    
    For i = 0 To lst��ҩ��.ListCount - 1
        If lst��ҩ��.Selected(i) Then
            str��ҩID = str��ҩID & "," & lst��ҩ��.ItemData(i)
        End If
    Next
    For i = 0 To lst��ҩ��.ListCount - 1
        If lst��ҩ��.Selected(i) Then
            str��ҩID = str��ҩID & "," & lst��ҩ��.ItemData(i)
        End If
    Next
    For i = 0 To lst��ҩ��.ListCount - 1
        If lst��ҩ��.Selected(i) Then
            str��ҩID = str��ҩID & "," & lst��ҩ��.ItemData(i)
        End If
    Next
    If str��ҩID <> "" Then str��ҩID = str��ҩID & ","
    If str��ҩID <> "" Then str��ҩID = str��ҩID & ","
    If str��ҩID <> "" Then str��ҩID = str��ҩID & ","
    lst��ҩ��.Clear: lst��ҩ��.Clear: lst��ҩ��.Clear
    
    Set rsTmp = GetDepartments("'��ҩ��','��ҩ��','��ҩ��'", IIf(opt����(0).Value, 1, 2) & ",3")
    If Not rsTmp.EOF Then
        rsTmp.Filter = "��������='��ҩ��'"
        Do While Not rsTmp.EOF
            lst��ҩ��.AddItem rsTmp!����
            lst��ҩ��.ItemData(lst��ҩ��.ListCount - 1) = rsTmp!ID
            If InStr(str��ҩID, "," & rsTmp!ID & ",") > 0 Then lst��ҩ��.Selected(lst��ҩ��.NewIndex) = True
            rsTmp.MoveNext
        Loop
        
        rsTmp.Filter = "��������='��ҩ��'"
        Do While Not rsTmp.EOF
            lst��ҩ��.AddItem rsTmp!����
            lst��ҩ��.ItemData(lst��ҩ��.ListCount - 1) = rsTmp!ID
            If InStr(str��ҩID, "," & rsTmp!ID & ",") > 0 Then lst��ҩ��.Selected(lst��ҩ��.NewIndex) = True
            rsTmp.MoveNext
        Loop
        
        rsTmp.Filter = "��������='��ҩ��'"
        Do While Not rsTmp.EOF
            lst��ҩ��.AddItem rsTmp!����
            lst��ҩ��.ItemData(lst��ҩ��.ListCount - 1) = rsTmp!ID
            If InStr(str��ҩID, "," & rsTmp!ID & ",") > 0 Then lst��ҩ��.Selected(lst��ҩ��.NewIndex) = True
            rsTmp.MoveNext
        Loop
    End If
    
    If lst��ҩ��.ListCount > 0 Then lst��ҩ��.ListIndex = 0
    If lst��ҩ��.ListCount > 0 Then lst��ҩ��.ListIndex = 0
    If lst��ҩ��.ListCount > 0 Then lst��ҩ��.ListIndex = 0
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetDrugStore()
    Dim lngType As Long, strTmp As String, arrTmp As Variant
    Dim i As Long, j As Long, lngRow As Long
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    With vsfDrugStore
        strTmp = "'��ҩ��','��ҩ��','��ҩ��','���ϲ���'"
        
        If stab.TabVisible(1) = True Then
            lngType = IIf(opt����(0).Value, 1, 2)
        Else
            lngType = gint������Դ
        End If
        Set rsTmp = GetDepartments(strTmp, lngType & ",3")
        .Rows = 1
        If mbytInFun = 2 Then .ColHidden(3) = True '������ʲ��贰��
        
        If rsTmp.RecordCount > 0 Then
            rsTmp.Filter = "��������<>'���ϲ���'"
            .Rows = rsTmp.RecordCount + 1
            .MergeCells = flexMergeFixedOnly
            .MergeCol(0) = True
            
            strTmp = "'��ҩ��','��ҩ��','��ҩ��'"
            arrTmp = Split(strTmp, ",")
            lngRow = 1
            For j = 0 To UBound(arrTmp)
                rsTmp.Filter = "��������=" & arrTmp(j)
                If rsTmp.RecordCount > 0 Then
                    For i = 1 To rsTmp.RecordCount
                        .TextMatrix(lngRow, 0) = Replace(arrTmp(j), "'", "")
                        .TextMatrix(lngRow, 1) = 0
                        .TextMatrix(lngRow, 2) = rsTmp!����
                        If mbytInFun <> 2 Then .TextMatrix(lngRow, 3) = "�Զ�����"
                        .RowData(lngRow) = Val(rsTmp!ID)
                        lngRow = lngRow + 1
                        rsTmp.MoveNext
                    Next
                    
                    If lngRow < .Rows - 1 Then  '���ָ���
                        .Select lngRow, .FixedCols, lngRow, .COLS - 1
                        .CellBorder vbBlue, 0, 1, 0, 0, 0, 0
                    End If
                End If
            Next
            
            cbo����.AddItem "�˹�ѡ��"
            rsTmp.Filter = "��������='���ϲ���'"
            For j = 1 To rsTmp.RecordCount
                cbo����.AddItem rsTmp!����
                cbo����.ItemData(cbo����.NewIndex) = rsTmp!ID
                rsTmp.MoveNext
            Next
            cbo����.ListIndex = 0
        End If
    
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmd�շ�ִ�п���_Click()
    Dim rsDept As ADODB.Recordset
    Dim strSQL As String
    Dim vRect As RECT
    Dim blnCancel As Boolean
    Dim strTemp As String
    
    Err = 0: On Error GoTo errHandler
    '96357
    strSQL = "Select Distinct A.ID, A.����, A.����, A.����" & vbNewLine & _
            " From ���ű� A, ��������˵�� B" & vbNewLine & _
            " Where B.����ID=A.ID And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & vbNewLine & _
            "       And B.�������� In('��ҩ��', '��ҩ��', '��ҩ��', '���ϲ���')" & vbNewLine & _
            "       And B.������� In (1, 2, 3)" & vbNewLine & _
            "       And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & vbNewLine & _
            " Order by A.����"
    vRect = GetControlRect(txt�շ�ִ�п���.hWnd)
    Set rsDept = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "�����շ�ִ�п���", True, "", "", False, False, False, _
        vRect.Left, vRect.Top, txt�շ�ִ�п���.Height, blnCancel, False, True, "MultiCheckReturn=1")
    If blnCancel Then Exit Sub
    If rsDept Is Nothing Then Exit Sub
    
    txt�շ�ִ�п���.Text = ""
    txt�շ�ִ�п���.Tag = ""
    Do While Not rsDept.EOF
        txt�շ�ִ�п���.Text = txt�շ�ִ�п���.Text & ";" & Nvl(rsDept!����)
        strTemp = strTemp & "," & Nvl(rsDept!ID)
        rsDept.MoveNext
    Loop
    If txt�շ�ִ�п���.Text <> "" Then txt�շ�ִ�п���.Text = Mid(txt�շ�ִ�п���.Text, 2)
    If strTemp <> "" Then txt�շ�ִ�п���.Tag = Mid(strTemp, 2)
    
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
        
    Select Case mbytInFun
    Case 0
    Case 1
        Me.Height = 5325: stab.Height = 4695
        cmdHelp.Top = Me.ScaleHeight - cmdHelp.Height - 100
        cbo����.Top = stab.Height + stab.Top - cbo����.Height - 200
        lbl���ϲ���.Top = cbo����.Top + (cbo����.Height - lbl���ϲ���.Height) \ 2
        vsfDrugStore.Height = cbo����.Top - vsfDrugStore.Top - 50
    Case 2
        Me.Height = 6025 + IIf(chkסԺ�������շ�.Visible, chkסԺ�������շ�.Height + 20, 0)
        stab.Height = 6055 + IIf(chkסԺ�������շ�.Visible, chkסԺ�������շ�.Height + 20, 0)
        Me.cmdHelp.Top = 5095
        cbo����.Top = stab.Height + stab.Top - cbo����.Height - 200
        lbl���ϲ���.Top = cbo����.Top + (cbo����.Height - lbl���ϲ���.Height) \ 2
        vsfDrugStore.Height = cbo����.Top - vsfDrugStore.Top - 50
    End Select
End Sub
Private Sub MoveCtrol()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ���λ��
    '����:���˺�
    '����:2011-09-12 13:46:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    ''0=�շ�,1=����,2=�������

    Select Case mbytInFun
    Case 2   '�������
        '��ʾΪ�����һ��
        chkPay.Top = chkPay.Top
        chkPay.Left = fra��λ.Left
        chkTime.Left = chkPay.Left + chkPay.Width + 50
        chk��ʿ.Left = chkTime.Left + chkTime.Width + 50
        chkTime.Top = chkPay.Top
        chk��ʿ.Top = chkPay.Top
        '�Զ���ʾ,����Frame�еĿؼ�,��Ȼ���ڵ�ǰ��ҳ��
        fra������ҽ��.Left = fra��λ.Left
        fra������ҽ��.Top = fra��λ.Top - fra������ҽ��.Height - 100
        optUnit.Top = optUnit.Top + 30
        optDoctor.Top = optUnit.Top + optUnit.Height + 80
        optSelf.Top = optDoctor.Top + optDoctor.Height + 80
        
        fraPrintBill.Height = fraPrintBill.Height - 50
        fra������ҽ��.Height = fraPrintBill.Height
        
        fra��λ.Top = fra��λ.Top + 100
        fraPrintBill.Left = fra��λ.Left + fra��λ.Width - fraPrintBill.Width
        fraPrintBill.Top = fra������ҽ��.Top ' fra��λ.Top - fraPrintBill.Height - 100
        
        chkSeekName.Top = chkPay.Top + chkPay.Height + 80
        chkSeekName.Left = chkPay.Left
        txtNameDays.Top = chkSeekName.Top
        txtNameDays.Left = chkSeekName.Left + chkSeekName.Width * 14 / 23 + 50
        fraShortLine.Top = txtNameDays.Top + txtNameDays.Height
        fraShortLine.Left = txtNameDays.Left
        
        chkOnlyUnitPatient.Top = chkSeekName.Top + chkSeekName.Height + 50
        chkOnlyUnitPatient.Left = chkPay.Left
        
       ' lst�շ����.Height = lst�շ����.Height - 80
        cmdDeviceSetup.Left = fra���.Left
        cmdDeviceSetup.Visible = True
        cmdDeviceSetup.Top = cmdDeviceSetup.Top - 150
        'fraDoctor.Top = cmdDeviceSetup.Top - fraDoctor.Height - 100
        'fraDoctor.Left = fra���.Left
        
        fra�����ʾ.Top = fra�����ʾ.Top + 100
        fraDoctor.Left = fra�����ʾ.Left
        fraDoctor.Top = fra�����ʾ.Top + fra�����ʾ.Height + 50
        fraDoctor.Width = fra�����ʾ.Width
        optDoctorKind(0).Caption = "������+������ʾ"
        optDoctorKind(1).Caption = "������+������ʾ"
        
        optDoctorKind(0).Width = Me.TextWidth(optDoctorKind(0).Caption) + 400
        optDoctorKind(1).Width = Me.TextWidth(optDoctorKind(1).Caption) + 400
        
        
        optDoctorKind(0).Top = optDoctorKind(0).Top + 40
        optDoctorKind(0).Left = fraDoctor.Left + (fraDoctor.Width - optDoctorKind(0).Width - optDoctorKind(1).Width - 1000) \ 2
        
        optDoctorKind(1).Left = optDoctorKind(0).Left + optDoctorKind(0).Width + 50
        optDoctorKind(1).Top = optDoctorKind(0).Top
        
        fraDoctor.Height = fraDoctor.Height - 250
        chk������.Top = fraDoctor.Top + fraDoctor.Height + 50: chk������.Left = chk���������ɿ�.Left
        'fra���.Height = fraDoctor.Top - fra���.Top - 100
        fra���.Height = fra���.Height - 80
        lst�շ����.Height = fra���.Height - 300
    Case 1 '���۵�
        txtDay.Top = cbo���㷽ʽ.Top
        udDay.Top = txtDay.Top
        lblDay.Top = txtDay.Top + (txtDay.Height - lblDay.Height) \ 2
        fra��λ.Top = txtDay.Top + txtDay.Height + 200
        fra�����ʾ.Top = fra��λ.Top + fra��λ.Height + 200
        chk������.Top = fra�����ʾ.Top + fra�����ʾ.Height + 50: chk������.Left = chk���������ɿ�.Left
        
        fraDoctor.Top = fraRegPrompt.Top
        fraDoctor.Height = fraRegPrompt.Height
        fraDoctor.Left = fraRegPrompt.Left
        fraDoctor.Width = fraRegPrompt.Width
        optDoctorKind(0).Top = optDoctorKind(0).Top + 80
        optDoctorKind(1).Top = optDoctorKind(0).Top + optDoctorKind(0).Height + 50
        fraInputItem.Width = fraRegPrompt.Left + fraRegPrompt.Width - fraInputItem.Left
        chk�Ա�.Left = chk�Ա�.Left + 100
        chk��������.Left = chk�Ա�.Left
        
        chk�Ƿ�Ӱ�.Left = chk��������.Left + chk��������.Width + 800
        chk������.Left = chk�Ƿ�Ӱ�.Left
        
        chk����.Left = chk�Ƿ�Ӱ�.Left + chk�Ƿ�Ӱ�.Width + 800
        chkҽ�Ƹ���.Left = chk����.Left
        chk�ѱ�.Left = chkҽ�Ƹ���.Left + chkҽ�Ƹ���.Width - chk�ѱ�.Width
        fra����֪ͨ����ӡ.Top = cmdPrintSetup(3).Top
        cmdPrintSetup(3).Top = fra����֪ͨ����ӡ.Top + fra����֪ͨ����ӡ.Height + 100
        cmdDeviceSetup.Left = fraInputItem.Left
        cmdDeviceSetup.Top = cmdPrintSetup(3).Top
           
    Case Else   '�շ�
        fra���.Top = fra���.Top
        txtMax.Top = fra���.Top
        lblMax.Top = txtMax.Top + (txtMax.Height - lblMax.Height) \ 2
        chkPay.Top = lblMax.Top
        chkTime.Top = chkPay.Top + chkPay.Height + IIf(gbln���뷢ҩ, 50, 100)
        chk��ʿ.Top = chkTime.Top + chkTime.Height + IIf(gbln���뷢ҩ, 50, 100)
        chk�ۼ�.Top = chk��ʿ.Top + chk��ʿ.Height + IIf(gbln���뷢ҩ, 50, 100)
        
        cbo�ѱ�.Top = txtMax.Top + txtMax.Height + IIf(gbln���뷢ҩ, 50, 100)
        lbl�ѱ�.Top = cbo�ѱ�.Top + (cbo�ѱ�.Height - lbl�ѱ�.Height) \ 2
        
        
        cbo���㷽ʽ.Top = cbo�ѱ�.Top + cbo�ѱ�.Height + IIf(gbln���뷢ҩ, 50, 100)
        lbl���㷽ʽ.Top = cbo���㷽ʽ.Top + (cbo���㷽ʽ.Height - lbl���㷽ʽ.Height) \ 2
        
        chkƤ��.Top = cbo���㷽ʽ.Height + cbo���㷽ʽ.Top + IIf(gbln���뷢ҩ, 50, 100)
        chkPrePayPriority.Top = chkƤ��.Top + chkƤ��.Height + IIf(gbln���뷢ҩ, 50, 100)
        
        txtAddedItem.Top = chkPrePayPriority.Top + chkPrePayPriority.Height + IIf(gbln���뷢ҩ, 0, 100)
        cmdAddedItem.Top = txtAddedItem.Top
        chkAddedItem.Top = txtAddedItem.Top + (txtAddedItem.Height - chkAddedItem.Height) \ 2
        
        chkInsurePartFee.Top = txtAddedItem.Top + txtAddedItem.Height + IIf(gbln���뷢ҩ, 50, 100)
        
        fra��λ.Top = chkInsurePartFee.Top + chkInsurePartFee.Height + IIf(gbln���뷢ҩ, 0, 100)
        fra��λ.Height = fra��λ.Height + IIf(gbln���뷢ҩ, 0, 100)
        opt��λ(0).Top = opt��λ(0).Top + IIf(gbln���뷢ҩ, 0, 50)
        opt��λ(1).Top = opt��λ(0).Top
        lbl��λ.Top = opt��λ(0).Top
        fra�����ʾ.Height = 700
        fra�����ʾ.Top = fra��λ.Top + fra��λ.Height + IIf(gbln���뷢ҩ, 50, 100)
        fra����.Top = fra�����ʾ.Top
        
        If Not gbln���뷢ҩ Then
            chkסԺ�������շ�.Top = fra�����ʾ.Top + fra�����ʾ.Height + 50
        Else
            chkסԺ�������շ�.Top = fra����.Top + fra����.Height + 20
        End If
        chkҽ��������ȱʡ��λ.Top = chkסԺ�������շ�.Top
        chk���������ɿ�.Top = chkסԺ�������շ�.Top + chkסԺ�������շ�.Height + 50
        chkPayKey.Top = chk���������ɿ�.Top
        chk��ֹȡ���Һŵ�.Top = chk���������ɿ�.Top + chk���������ɿ�.Height + 50
        chk������.Top = chk��ֹȡ���Һŵ�.Top: chk������.Left = chkPayKey.Left
        fraDrugNotFee.Top = chk������.Top + chk������.Height + 50
        fra�˷�ȱʡѡ��ʽ.Top = fraDrugNotFee.Top + fraDrugNotFee.Height + 50
        fraSetMoneyMode.Top = fra�˷�ȱʡѡ��ʽ.Top + fra�˷�ȱʡѡ��ʽ.Height + 50
        
        '�ڶ�ҳ����
        chk�����俪����.Top = chkȱʡ��������.Top
        chk�����俪����.Left = chkLed.Left
        opt����(0).Top = chkOnlyUnitPatient.Top
        opt����(1).Top = opt����(0).Top + opt����(0).Height + 20
        opt����(2).Top = opt����(1).Top + opt����(1).Height + 20
        
        chkLed.Top = chk�����俪����.Top + chk�����俪����.Height + 20
        chkLedDispDetail.Top = chkLed.Top + chkLed.Height + 20
        chkLedWelcome.Top = chkLedDispDetail.Top + chkLedDispDetail.Height + 20
        chkUnPopPriceBill.Top = chkLedWelcome.Top + chkLedWelcome.Height + 20
        chkMustRegevent.Top = chkUnPopPriceBill.Top + chkUnPopPriceBill.Height + 20
        chkMustRegevent.Left = chkUnPopPriceBill.Left
        cmdDeviceSetup.Top = chkMustRegevent.Top + chkMustRegevent.Height + 100
        
'        chkShowError.Top = opt����(2).Top + opt����(2).Height + 20
'        chkMulti.Top = chkShowError.Top + chkShowError.Height + 20
        chkMulti.Top = opt����(2).Top + opt����(2).Height + 20
        chkSeekBill.Top = chkMulti.Top + chkMulti.Height + 20
        txtSeekDays.Top = chkSeekBill.Top
        fraLine.Top = txtSeekDays.Top + txtSeekDays.Height
        
        cboAutoSplitBill.Top = chkSeekBill.Top + chkSeekBill.Height + 20
        chkAutoSplitBill.Top = cboAutoSplitBill.Top + (cboAutoSplitBill.Height - chkAutoSplitBill.Height) \ 2
        fra�ɿ����.Top = IIf(cboAutoSplitBill.Top + cboAutoSplitBill.Height + 20 > cmdDeviceSetup.Top + cmdDeviceSetup.Height + 20, cboAutoSplitBill.Top + cboAutoSplitBill.Height + 20, cmdDeviceSetup.Top + cmdDeviceSetup.Height + 20)
        
        txt�շ�ִ�п���.Top = fra�ɿ����.Top + fra�ɿ����.Height + 100
        chk�շ�ִ�п���.Top = txt�շ�ִ�п���.Top + (txt�շ�ִ�п���.Height - chk�շ�ִ�п���.Height) / 2
        cmd�շ�ִ�п���.Top = txt�շ�ִ�п���.Top
        
        fra���.Height = fra���.Height - 50
        lst�շ����.Height = lst�շ����.Height + 100
    End Select
End Sub
Private Sub SetCtrlVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ�����������
    '����:���˺�
    '����:2011-09-12 14:55:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If gbln���뷢ҩ Then
        stab.TabVisible(4) = False: fra�����ʾ.Visible = False
        fra����.Visible = True
    End If
    '���˺� ����:27663 ����:2010-01-27 13:29:19
    chkסԺ�������շ�.Visible = mbytInFun = 0
    chk���������ɿ�.Visible = mbytInFun = 0
    '47457
    chkPayKey.Visible = mbytInFun = 0
    '87489
    fra�˷�ȱʡѡ��ʽ.Visible = mbytInFun = 0
    fraSetMoneyMode.Visible = mbytInFun = 0
End Sub

Private Sub Loadҩ��ParaValue()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҩ����ز���ֵ
    '����:���˺�
    '����:2011-12-07 15:05:10
    '����:43775
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String, blnParSet As Boolean
    Dim i As Long, k As Long, j As Long, intType As Integer
    Dim arrTmp  As Variant, arrWindow As Variant
    Dim str��ҩ������ As String, str��ҩ������ As String, str��ҩ������ As String
    Dim lngȱʡ��ҩ�� As Long, lngȱʡ��ҩ�� As Long, lngȱʡ��ҩ�� As Long, lngȱʡ���ϲ��� As Long
    blnParSet = InStr(1, mstrPrivs, ";��������;") > 0
    If gbln���뷢ҩ = True Then
        strTmp = zlDatabase.GetPara("��ҩ��ѡ��", glngSys, mlngModul, , Array(lst��ҩ��), blnParSet)
        For i = 0 To lst��ҩ��.ListCount - 1
            If InStr("," & strTmp & ",", "," & lst��ҩ��.ItemData(i) & ",") > 0 Then
                lst��ҩ��.Selected(i) = True
            End If
        Next
        strTmp = zlDatabase.GetPara("��ҩ��ѡ��", glngSys, mlngModul, , Array(lst��ҩ��), blnParSet)
        For i = 0 To lst��ҩ��.ListCount - 1
            If InStr("," & strTmp & ",", "," & lst��ҩ��.ItemData(i) & ",") > 0 Then
                lst��ҩ��.Selected(i) = True
            End If
        Next
        strTmp = zlDatabase.GetPara("��ҩ��ѡ��", glngSys, mlngModul, , Array(lst��ҩ��), blnParSet)
        For i = 0 To lst��ҩ��.ListCount - 1
            If InStr("," & strTmp & ",", "," & lst��ҩ��.ItemData(i) & ",") > 0 Then
                lst��ҩ��.Selected(i) = True
            End If
        Next
        If lst��ҩ��.ListCount > 0 Then lst��ҩ��.ListIndex = 0
        If lst��ҩ��.ListCount > 0 Then lst��ҩ��.ListIndex = 0
        If lst��ҩ��.ListCount > 0 Then lst��ҩ��.ListIndex = 0
        Exit Sub
    End If
    
    With vsfDrugStore
        arrTmp = Split("ȱʡ��ҩ��,ȱʡ��ҩ��,ȱʡ��ҩ��", ",")
        .Cell(flexcpData, 0, 0, .Rows - 1, .COLS - 1) = "0" '�洢�Ƿ��������.:0-������,1-����
        
        For j = 0 To UBound(arrTmp)
            '���˺�:���ڿ��ܲ���Ȩ�޷������,���,����ͳһ��������,��Ҫ����ĳһ����:
            '����:25132,intType-'���ز������ͣ�1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
            strTmp = zlDatabase.GetPara(CStr(arrTmp(j)), glngSys, mlngModul, "0", , blnParSet, intType)
            If Val(strTmp) > 0 Then
                Select Case arrTmp(j)
                    Case "ȱʡ��ҩ��"
                        lngȱʡ��ҩ�� = Val(strTmp)
                    Case "ȱʡ��ҩ��"
                        lngȱʡ��ҩ�� = Val(strTmp)
                    Case "ȱʡ��ҩ��"
                        lngȱʡ��ҩ�� = Val(strTmp)
                End Select
                Call SetDrugStockEdit(Replace(arrTmp(j), "ȱʡ", ""), intType, .ColIndex("ȱʡ"), Val(strTmp))
            Else
                Call SetDrugStockEdit(Replace(arrTmp(j), "ȱʡ", ""), intType, .ColIndex("ȱʡ"), "")
            End If
        Next
        
        strTmp = zlDatabase.GetPara("ȱʡ���ϲ���", glngSys, mlngModul, "0", Array(cbo����), blnParSet)
        zlControl.CboLocate cbo����, strTmp, True
        
        If mbytInFun <> 2 Then
                arrTmp = Split("��ҩ������,��ҩ������,��ҩ������", ",")
                For j = 0 To UBound(arrTmp)
                    strTmp = Trim(zlDatabase.GetPara(CStr(arrTmp(j)), glngSys, mlngModul, , , blnParSet, intType))
                    If strTmp <> "" Then
                        '����ɵ�����,���ڲ�����û�д洢ҩ��ID
                        If InStr(strTmp, ":") = 0 Then
                            Select Case arrTmp(j)
                                Case "��ҩ������"
                                    strTmp = lngȱʡ��ҩ�� & ":" & strTmp
                                Case "��ҩ������"
                                    strTmp = lngȱʡ��ҩ�� & ":" & strTmp
                                Case "��ҩ������"
                                    strTmp = lngȱʡ��ҩ�� & ":" & strTmp
                            End Select
                        End If
                        arrWindow = Split(strTmp, ",")
                        strTmp = Replace(arrTmp(j), "����", "")
                        For k = 0 To UBound(arrWindow)
                            Call SetDrugStockEdit(Replace(arrTmp(j), "����", ""), intType, .ColIndex("����"), Val(Split(arrWindow(k), ":")(0)), CStr(Split(arrWindow(k), ":")(1)))
                        Next
                    Else
                        Call SetDrugStockEdit(Replace(arrTmp(j), "����", ""), intType, .ColIndex("����"), "")
                    End If
                Next
            End If
        End With
        chkҩ��.Value = IIf(zlDatabase.GetPara("��ʾ����ҩ�����", glngSys, mlngModul, , Array(chkҩ��), blnParSet) = "1", 1, 0)
        chkҩ��.Value = IIf(zlDatabase.GetPara("��ʾ����ҩ����", glngSys, mlngModul, , Array(chkҩ��), blnParSet) = "1", 1, 0)
        If mbytInFun <> 0 Then
            If Val(Val(zlDatabase.GetPara("�����ʾ��ʽ", glngSys, mlngModul, , Array(opt���(0), opt���(1)), blnParSet))) = 0 Then
                opt���(0).Value = True
            Else
                opt���(1).Value = True
            End If
            If opt���(0).Enabled = False Then opt���(0).Tag = "1"
        End If
     '   Call chkҩ��_Click
End Sub
Private Sub LoadParaValue()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ز���ֵ
    '����:���˺�
    '����:2011-09-12 15:03:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String, blnParSet As Boolean, k As Long, rsTmp As ADODB.Recordset
    Dim i As Long, arrTmp As Variant, j As Long, intType As Integer, arrWindow As Variant
        
    
    blnParSet = InStr(1, mstrPrivs, ";��������;") > 0

    strTmp = zlDatabase.GetPara("�շ����", glngSys, mlngModul, , Array(lst�շ����), blnParSet)
    If strTmp = "" Then
        For i = 0 To lst�շ����.ListCount - 1
            lst�շ����.Selected(i) = True
        Next
    Else
        For i = 0 To lst�շ����.ListCount - 1
            If InStr(strTmp, Chr(lst�շ����.ItemData(i))) Then lst�շ����.Selected(i) = True
        Next
    End If
    If lst�շ����.ListCount > 0 Then lst�շ����.TopIndex = 0: lst�շ����.ListIndex = 0
    If mbytInFun <> 2 Then
        strTmp = zlDatabase.GetPara("ȱʡ�ѱ�", glngSys, mlngModul, , Array(cbo�ѱ�), blnParSet)
        zlControl.CboLocate cbo�ѱ�, strTmp
    End If
    chkPay.Value = IIf(zlDatabase.GetPara("��ҩ����", glngSys, mlngModul, , Array(chkPay), blnParSet) = "1", 1, 0)
    chkTime.Value = IIf(zlDatabase.GetPara("�������", glngSys, mlngModul, , Array(chkTime), blnParSet) = "1", 1, 0)
    chk��ʿ.Value = IIf(zlDatabase.GetPara("��ʾ��ʿ", glngSys, mlngModul, , Array(chk��ʿ), blnParSet) = "1", 1, 0)
    i = IIf(zlDatabase.GetPara("ҩƷ��λ", glngSys, mlngModul, , Array(opt��λ(0), opt��λ(1)), blnParSet) = "0", 0, 1)
    opt��λ(i).Value = True
    If mbytInFun = 0 Or mbytInFun = 1 Then
        i = IIf(zlDatabase.GetPara("������Դ", glngSys, mlngModul, , Array(opt����(0), opt����(1)), blnParSet) = "1", 0, 1)
        opt����(i).Value = True
    End If
    chk������.Value = IIf(Val(zlDatabase.GetPara("����¼������ʹ�õĿ�����", glngSys, mlngModul, "0", Array(chk������), blnParSet)) = 1, 1, 0)
    
    Call opt����_Click(IIf(opt����(0).Value, 0, 1)) '����ҩƷ�ⷿ�����ķ��ϲ���
    Call Loadҩ��ParaValue
    Select Case mbytInFun
    Case 2 '����
    Case 1 '����
    Case Else
        opt���(0).Visible = False: opt���(1).Visible = False: lbl�����ʾ��ʽ.Visible = False
        lnSplit(0).Visible = False: lnSplit(1).Visible = False
        chkRegistInvoice.Value = IIf(zlDatabase.GetPara("�ҺŹ����շ�Ʊ��", glngSys, mlngModul, 0, Array(chkRegistInvoice), blnParSet) = "1", 1, 0)
        chkLed.Value = IIf(zlDatabase.GetPara("�ֹ�����", glngSys, mlngModul, 0, Array(chkLed), blnParSet) = "1", 1, 0)
        chkLedDispDetail.Value = IIf(zlDatabase.GetPara("LED��ʾ�շ���ϸ", glngSys, mlngModul, 1, Array(chkLedDispDetail), blnParSet) = "1", 1, 0)
        chkLedWelcome.Value = IIf(zlDatabase.GetPara("LED��ʾ��ӭ��Ϣ", glngSys, mlngModul, 1, Array(chkLedWelcome), blnParSet) = "1", 1, 0)
        Set rsTmp = Get���㷽ʽ("�շ�", "1,2,7")
        For i = 1 To rsTmp.RecordCount
            cbo���㷽ʽ.AddItem rsTmp!����
            If rsTmp!ȱʡ = 1 Then cbo���㷽ʽ.ListIndex = cbo���㷽ʽ.NewIndex
            rsTmp.MoveNext
        Next
        '����:54923
        strTmp = zlDatabase.GetPara("ȱʡ���㷽ʽ", glngSys, mlngModul, , Array(cbo���㷽ʽ), blnParSet)
        For i = 0 To cbo���㷽ʽ.ListCount - 1
            If cbo���㷽ʽ.List(i) = strTmp Then cbo���㷽ʽ.ListIndex = i: Exit For
        Next
        '���ط�Ʊ���
        Call InitShareInvoice
         '39253
        chk���������ɿ�.Value = IIf(Val(zlDatabase.GetPara("��ȡ���ۺ������ɿ�", glngSys, mlngModul, "0", Array(chk���������ɿ�), blnParSet)) = 1, 1, 0)
        chkסԺ�������շ�.Value = IIf(Val(zlDatabase.GetPara("סԺ���˰������շ�", glngSys, mlngModul, "0", Array(chkסԺ�������շ�), blnParSet)) = 1, 1, 0)
        '120836
        chk��ֹȡ���Һŵ�.Value = IIf(Val(zlDatabase.GetPara("��ֹȡ���ҺŻ��۵�", glngSys, mlngModul, "0", Array(chk��ֹȡ���Һŵ�), blnParSet)) = 1, 1, 0)
        '47457
        chkPayKey.Value = IIf(Val(zlDatabase.GetPara("ʹ�üӼ��л�֧����ʽ", glngSys, mlngModul, "1", Array(chkPayKey), blnParSet)) = 1, 1, 0)
        '47400
        strTmp = zlDatabase.GetPara("ҩƷ��ҩ�˷ѷ�ʽ", glngSys, mlngModul, , Array(optDrug(0), optDrug(1), optDrug(2)), blnParSet)
        For i = 0 To 2
            If Val(strTmp) = i Then
                optDrug(i).Value = True: Exit For
            End If
        Next
        '87489
        strTmp = zlDatabase.GetPara("�˷�ȱʡѡ��ʽ", glngSys, mlngModul, "0", Array(opt�˷�ȱʡѡ��ʽ(0), opt�˷�ȱʡѡ��ʽ(1)), blnParSet)
        For i = 0 To 1
            If Val(strTmp) = i Then opt�˷�ȱʡѡ��ʽ(i).Value = True: Exit For
        Next
        chkҽ��������ȱʡ��λ.Value = IIf(zlDatabase.GetPara("ҽ��������ȱʡ��λ", glngSys, mlngModul, "0", Array(chkҽ��������ȱʡ��λ), blnParSet) = "1", 1, 0)
        '86853
        i = Val(zlDatabase.GetPara("ˢ��ȱʡ������", glngSys, 1151, "0", Array(optSetMoneyMode(0), optSetMoneyMode(1), optSetMoneyMode(2)), blnParSet))
        If i < 0 Or i > optSetMoneyMode.UBound Then i = 0
        optSetMoneyMode(i).Value = True
        
        '56963:Ʊ�ŷ������
        chkAutoAddBookFee.Value = IIf(Val(zlDatabase.GetPara("�վݼ��չ�����", glngSys, mlngModul, "0", Array(chkAutoAddBookFee), blnParSet)) = 1, 1, 0)
        chkErrorItemNotBill.Value = IIf(Val(zlDatabase.GetPara("����ʹ��Ʊ��", glngSys, mlngModul, "0", Array(chkErrorItemNotBill), blnParSet)) = 1, 1, 0)

         
         '56963:2.����Ԥ���������Ʊ��
         strTmp = Trim(zlDatabase.GetPara("Ʊ�ݷ������", glngSys, mlngModul, "0||0;0;0;0;0;0", _
         Array(cboBillRole, lblBillRole, chkBillRule(0), chkBillRule(1), chkBillRule(2), chkBillRule(3), optRuleTotal(0), optRuleTotal(1), optRuleTotal(2), _
         lblBillRuleNum(0), updBillRuleNum(0), txtBillRuleNum(0), lblBillRuleNum(1), updBillRuleNum(1), lblBillRuleNum(2), txtBillRuleNum(1), updBillRuleNum(2), txtBillRuleNum(2)), blnParSet))
         arrTmp = Split(strTmp & "||", "||")
         
         '���ⱻ����
         optRuleTotal(0).Tag = IIf(optRuleTotal(0).Enabled, 1, 0)
        With cboBillRole
            .Clear
            .AddItem "1-����ʵ�ʴ�ӡ����Ʊ��"
            If Val(arrTmp(0)) = 0 Then .ListIndex = .NewIndex
            .AddItem "2-����Ԥ���������Ʊ��"
            If Val(arrTmp(0)) = 1 Then .ListIndex = .NewIndex
            .AddItem "3-�����Զ���������Ʊ��"
            If Val(arrTmp(0)) = 2 Then .ListIndex = .NewIndex
            If .ListIndex < 0 Then .ListIndex = 0
            .Tag = .ListIndex   '��¼�޸�ǰ��ѡ��
            '56963:���ڴ�ӡ����ʱ,���������Ʊ�ŷ������
            .Enabled = .Enabled And Not GetPrintListHaveData
        End With
        '2.����Ԥ���������Ʊ��
        arrTmp = Split(arrTmp(1) & ";;;", ";")
        '�����ݷ�
        i = Val(arrTmp(0))
        chkBillRule(0).Value = IIf(i = 1, 1, 0)
        '��ִ�п��ҷ�
        i = Val(arrTmp(1))
        chkBillRule(1).Value = IIf(i >= 1, 1, 0)
        updBillRuleNum(0).Value = IIf(i < 0 Or i > 100, 0, i)
        txtBillRuleNum(0).Text = updBillRuleNum(0).Value
        txtBillRuleNum(0).Tag = IIf(updBillRuleNum(0).Value = 0, 1, updBillRuleNum(0).Value)
        
        '���վݷ�Ŀ
        i = Val(arrTmp(2))
        chkBillRule(2).Value = IIf(i >= 1, 1, 0)
        updBillRuleNum(1).Value = IIf(i < 0 Or i > 100, 0, i)
        txtBillRuleNum(1).Text = updBillRuleNum(1).Value
        txtBillRuleNum(1).Tag = IIf(updBillRuleNum(1).Value = 0, 1, updBillRuleNum(1).Value)
        '���շ�ϸĿ(�ȴ����շ�ϸĿ����Ȼ�ᴥ��Click�¼�������ҳ����ִ��Ϊ����
        i = Val(arrTmp(3))
        chkBillRule(3).Value = IIf(i >= 1, 1, 0)
        updBillRuleNum(2).Value = IIf(i < 0 Or i > 100, 0, i)
        txtBillRuleNum(2).Text = updBillRuleNum(2).Value
        txtBillRuleNum(2).Tag = IIf(updBillRuleNum(2).Value = 0, 20, updBillRuleNum(2).Value)
        
        '�������
        i = Val(arrTmp(4)): i = IIf(i > 3 Or i < 0, 0, i)
        optRuleTotal(i).Value = True
         
        '1.����ʵ�ʴ�ӡ����Ʊ��
        chkBillNO.Value = IIf(Val(zlDatabase.GetPara("���ŵ����շѷֱ��ӡ", glngSys, mlngModul, "0", Array(chkBillNO), blnParSet)) = 1, 1, 0)
        chk��찴���ݷֱ��ӡ.Value = IIf(Val(zlDatabase.GetPara("��첡�˷ֵ��ݴ�ӡ", glngSys, mlngModul, "0", Array(chk��찴���ݷֱ��ӡ), blnParSet)) = 1, 1, 0)
        chk��찴���ݷֱ��ӡ.Enabled = (chkBillNO.Value = vbChecked)
        chkOlnyOneBill.Value = IIf(Val(zlDatabase.GetPara("�շ�ÿ��ֻ��һ��Ʊ��", glngSys, mlngModul, "0", Array(chkOlnyOneBill), blnParSet)) = 1, 1, 0)
        i = Val(zlDatabase.GetPara("�շ��վ����д�", glngSys, mlngModul, "3", Array(lblRows, updRows, txtRowsUD), blnParSet))
        updRows.Value = IIf(i < 0 Or i > 100, 3, i)
        i = Val(zlDatabase.GetPara("�շ�Ʊ�����ɷ�ʽ", glngSys, mlngModul, "0", Array(optBillMode(0), optBillMode(1), chkExcuteDept), blnParSet))
        chkExcuteDept.Value = IIf(i >= 10, 1, 0)
        optBillMode(i Mod 10).Value = True
        '3-�����Զ���������Ʊ��
        Call SetBillRuleEnable
        Call ShowRuleInfor
    End Select
End Sub
Private Sub SaveBillRulePara()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ʊ�ݷ���������ز���
    '����:���˺�
    '����:2013-03-26 16:32:46
    '����:56963
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHavePrivs As Boolean, strTemp As String
    Dim intBillRull As Integer
    
    On Error GoTo errHandle
    If mbytInFun <> 0 Then Exit Sub
        
    If Val(cboBillRole.Tag) <> cboBillRole.ListIndex And cboBillRole.ListIndex <> 0 And Val(cboBillRole.Tag) <= 0 Then
       '�����ǰ�л�����ģʽ,��Ҫ��Ʊ�ݴ�ӡ��ʽ��¼����,�Ա����ش�򲿷��˷�ʱ���л�ǰ��Ʊ�ݸ�ʽ��ӡ
       Call zlDatabase.ExecuteProcedure("Zl_Update_Bill_Printformat(" & glngSys & ")", Me.Caption)
    End If
    'ֻ�ʺ��շ�
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    intBillRull = IIf(cboBillRole.ListIndex < 0, 0, cboBillRole.ListIndex)
    strTemp = intBillRull & "||"
    '�ֵ���
    strTemp = strTemp & IIf(chkBillRule(0).Value = 1, 1, 0)
    'ִ�п���
    strTemp = strTemp & ";" & IIf(chkBillRule(1).Value = 1, Val(txtBillRuleNum(0).Text), 0)
    '�վݷ�Ŀ
    strTemp = strTemp & ";" & IIf(chkBillRule(2).Value = 1, Val(txtBillRuleNum(1).Text), 0)
    '�շ�ϸĿ
    strTemp = strTemp & ";" & IIf(chkBillRule(3).Value = 1, Val(txtBillRuleNum(2).Text), 0)
    '��������
    strTemp = strTemp & ";" & IIf(optRuleTotal(0).Value, 0, IIf(optRuleTotal(1).Value, 1, 2))
    
    zlDatabase.SetPara "Ʊ�ݷ������", strTemp, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "�վݼ��չ�����", IIf(chkAutoAddBookFee.Value = 1, 1, 0), glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "����ʹ��Ʊ��", IIf(chkErrorItemNotBill.Value = 1, 1, 0), glngSys, mlngModul, blnHavePrivs
    If intBillRull = 0 Then
        '����ʵ�ʴ�ӡ����Ʊ��
        zlDatabase.SetPara "�շ�ÿ��ֻ��һ��Ʊ��", IIf(chkOlnyOneBill.Value = 1, 1, 0), glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "���ŵ����շѷֱ��ӡ", IIf(chkBillNO.Value = 1, 1, 0), glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "��첡�˷ֵ��ݴ�ӡ", IIf(chk��찴���ݷֱ��ӡ.Value = 1, 1, 0), glngSys, mlngModul, blnHavePrivs
        strTemp = CStr(IIf(optBillMode(1).Value, 1, 0) + Val(chkExcuteDept.Value) * 10)
        zlDatabase.SetPara "�շ�Ʊ�����ɷ�ʽ", strTemp, glngSys, mlngModul, blnHavePrivs
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Sub

Private Sub SetBillNoRule()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����Ʊ�ݷ�������λ��
    '����:���˺�
    '����:2013-03-26 15:43:12
    '����:56963
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, intIndex As Integer
    On Error GoTo errHandle
    intIndex = cboBillRole.ListIndex
    If intIndex < 0 Or intIndex > 2 Then intIndex = 0
    For i = 0 To 2
        picRuleBack(i).Visible = intIndex = i
    Next
    If intIndex = 0 Then
        '����ʵ�ʴ�ӡ����Ʊ��
        '��������
        Set chkAutoAddBookFee.Container = picRuleBack(0)
        Set chkErrorItemNotBill.Container = fraActuallyPrint
        '������
       chkAutoAddBookFee.Top = fraActuallyPrint.Top
       chkAutoAddBookFee.Left = chkOlnyOneBill.Left
       '�����
       chkErrorItemNotBill.Left = chkExcuteDept.Left
       chkErrorItemNotBill.Top = txtRowsUD.Top + (txtRowsUD.Height - chkErrorItemNotBill.Height) \ 2
    End If

    If intIndex = 1 Then
        '����Ԥ���������Ʊ��
        '��������
        Set chkAutoAddBookFee.Container = picRuleBack(1)
        Set chkErrorItemNotBill.Container = fraRuleSystem
        '������
       chkAutoAddBookFee.Top = fraRuleSystem.Top
       chkAutoAddBookFee.Left = fraRuleSystem.Left + 100
       '�����
       chkErrorItemNotBill.Left = optRuleTotal(0).Left
       chkErrorItemNotBill.Top = optRuleTotal(0).Top + optRuleTotal(0).Height + 50
       lblInfor.Top = chkErrorItemNotBill.Top + chkErrorItemNotBill.Height + 50
       
    End If
    If intIndex = 2 Then
        '�����û��Զ��������
        '��������
        Set chkAutoAddBookFee.Container = picRuleBack(2)
        Set chkErrorItemNotBill.Container = picRuleBack(2)
        '������
       chkAutoAddBookFee.Top = lblCustomInfor.Top - chkAutoAddBookFee.Height - 50
       chkAutoAddBookFee.Left = lblCustomInfor.Left
       '�����
       chkErrorItemNotBill.Left = chkAutoAddBookFee.Left + chkAutoAddBookFee.Width + 100
       chkErrorItemNotBill.Top = chkAutoAddBookFee.Top
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 

End Sub

Private Sub InitShareInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ù���Ʊ
    '����:���˺�
    '����:2011-04-28 15:09:10
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, lngRow As Long
    Dim strShareInvoice As String '����Ʊ������,��ʽ:����,����
    Dim varData As Variant, varTemp As Variant, VarType As Variant, varTemp1 As Variant
    Dim intType As Integer, intType1 As Integer   '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    Dim lngTemp As Long, i As Long, strSQL As String
    Dim strPrintMode As String, blnHavePrivs As Boolean
    Dim strOnePatiPrintShareInvoice As String, intOnePatiPrintType As Integer, varData1 As Variant
    
    On Error GoTo errHandle
    
    
      
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    
    chkOnePatiPrint.Value = IIf(Val(zlDatabase.GetPara("�����˲���Ʊ�����ֽ������", glngSys, mlngModul, "0", Array(chkOnePatiPrint), blnHavePrivs)) = 1, 1, 0)
    
    '�ָ��п��
    zl_vsGrid_Para_Restore mlngModul, vsBill, Me.Name, "����Ʊ��������", False, False
    zl_vsGrid_Para_Restore mlngModul, vsBillFormat, Me.Name, "�շ�Ʊ�ݸ�ʽ", False, False
    zl_vsGrid_Para_Restore mlngModul, vsDelBillFormat, Me.Name, "�˷�Ʊ�ݸ�ʽ", False, False
    
    strShareInvoice = zlDatabase.GetPara("�����շ�Ʊ������", glngSys, mlngModul, , , True, intType)
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
        If Val(.Tag) = 1 And InStr(1, mstrPrivs, ";��������;") = 0 Then .Editable = flexEDNone
    End With
    
    '��ʽ:����ID1,ʹ�����1|����IDn,ʹ�����n|...
    varData = Split(strShareInvoice, "|")
    '1.���ù���Ʊ��
    Set rsTemp = GetShareInvoiceGroupID(1)
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
    
    
    With vsBillFormat
        .ColHidden(.ColIndex("�����˲���Ʊ�ݸ�ʽ")) = chkOnePatiPrint.Value <> 1
    End With
    
    'Ʊ�ݸ�ʽ����
    strSQL = "" & _
    "   Select 'ʹ�ñ���ȱʡ��ʽ' as ˵��,0 as ���  From Dual Union ALL " & _
    "   Select B.˵��,B.���  " & _
    "   From zlReports A,zlRptFmts B" & _
    "   Where A.ID=B.����ID And A.���='ZL" & glngSys \ 100 & "_BILL_1121_1'  " & _
    "   Order by  ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With vsBillFormat
        .Clear 1
        .ColComboList(.ColIndex("�շ�Ʊ�ݸ�ʽ")) = .BuildComboList(rsTemp, "���,*˵��", "���")
        .ColComboList(.ColIndex("�����˲���Ʊ�ݸ�ʽ")) = .BuildComboList(rsTemp, "���,*˵��", "���")
        .ColComboList(.ColIndex("�շѴ�ӡ��ʽ")) = "0-����ӡƱ��|1-�Զ���ӡƱ��|2-ѡ���Ƿ��ӡƱ��"
    End With
    

    '��ȡ����ֵ
    strShareInvoice = zlDatabase.GetPara("�շѷ�Ʊ��ʽ", glngSys, mlngModul, , , True, intType)
    strOnePatiPrintShareInvoice = zlDatabase.GetPara("�����˲���Ʊ��ʽ", glngSys, mlngModul, , , True, intOnePatiPrintType)
    strPrintMode = zlDatabase.GetPara("�շѷ�Ʊ��ӡ��ʽ", glngSys, mlngModul, , , True, intType1)
    '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    With vsBillFormat
         .ColData(.ColIndex("�շ�Ʊ�ݸ�ʽ")) = "0"
         .ColData(.ColIndex("�����˲���Ʊ�ݸ�ʽ")) = "0"
         .ColData(.ColIndex("�շѴ�ӡ��ʽ")) = "0"
        .ForeColor = &H80000008:  .ForeColorFixed = &H80000008
        Select Case intType
        Case 1, 3, 5, 15
             .ColData(.ColIndex("�շ�Ʊ�ݸ�ʽ")) = IIf(intType = 5, 0, 1)
        End Select
        
        Select Case intOnePatiPrintType
        Case 1, 3, 5, 15
             .ColData(.ColIndex("�����˲���Ʊ�ݸ�ʽ")) = IIf(intOnePatiPrintType = 5, 0, 1)
        End Select
        
        Select Case intType1
        Case 1, 3, 5, 15
             .ColData(.ColIndex("�շѴ�ӡ��ʽ")) = IIf(intType1 = 5, 0, 1)
        End Select
        
        If (Val(.ColData(.ColIndex("�շ�Ʊ�ݸ�ʽ"))) = 1 Or _
            Val(.ColData(.ColIndex("�����˲���Ʊ�ݸ�ʽ"))) = 1 Or _
            Val(.ColData(.ColIndex("�շѴ�ӡ��ʽ"))) = 1) Then
            .Editable = IIf(Not blnHavePrivs, flexEDNone, flexEDKbdMouse)
        Else
            .Editable = flexEDKbdMouse
        End If
    End With
    
    vsBillFormat.Tag = ""
    varData = Split(strShareInvoice, "|")
    VarType = Split(strPrintMode, "|")
    varData1 = Split(strOnePatiPrintShareInvoice, "|")
    strSQL = "" & _
    "   Select ���� ,����" & _
    "   From  Ʊ��ʹ�����" & _
    "   order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With vsBillFormat
        .Clear 1
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("ʹ�����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("�շѴ�ӡ��ʽ")) = "0-����ӡƱ��"
            .TextMatrix(lngRow, .ColIndex("�շ�Ʊ�ݸ�ʽ")) = "0"
            .TextMatrix(lngRow, .ColIndex("�����˲���Ʊ�ݸ�ʽ")) = "0"
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(Nvl(rsTemp!����)) Then
                    .TextMatrix(lngRow, .ColIndex("�շ�Ʊ�ݸ�ʽ")) = Val(varTemp(1)): Exit For
                End If
            Next
            
            For i = 0 To UBound(varData1)
                varTemp = Split(varData1(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(Nvl(rsTemp!����)) Then
                    .TextMatrix(lngRow, .ColIndex("�����˲���Ʊ�ݸ�ʽ")) = Val(varTemp(1)): Exit For
                End If
            Next
            
            For i = 0 To UBound(VarType)
                varTemp1 = Split(VarType(i) & "," & ",", ",")
                If Trim(varTemp1(0)) = Trim(Nvl(rsTemp!����)) Then
                    .TextMatrix(lngRow, .ColIndex("�շѴ�ӡ��ʽ")) = Decode(Val(varTemp1(1)), 0, "0-����ӡƱ��", 1, "1-�Զ���ӡƱ��", "2-ѡ���Ƿ��ӡƱ��")
                    Exit For
                End If
            Next
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If Val(.ColData(.ColIndex("�շѴ�ӡ��ʽ"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("�շѴ�ӡ��ʽ"), .Rows - 1, .ColIndex("�շѴ�ӡ��ʽ")) = vbBlue
        End If
        
        If Val(.ColData(.ColIndex("�շ�Ʊ�ݸ�ʽ"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("�շ�Ʊ�ݸ�ʽ"), .Rows - 1, .ColIndex("�շ�Ʊ�ݸ�ʽ")) = vbBlue
        End If
        If Val(.ColData(.ColIndex("�����˲���Ʊ�ݸ�ʽ"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("�����˲���Ʊ�ݸ�ʽ"), .Rows - 1, .ColIndex("�����˲���Ʊ�ݸ�ʽ")) = vbBlue
        End If
    End With
    
    '�˷�Ʊ�ݸ�ʽ����
    strSQL = "" & _
    "   Select 'ʹ�ñ���ȱʡ��ʽ' as ˵��,0 as ���  From Dual Union ALL " & _
    "   Select B.˵��,B.���  " & _
    "   From zlReports A,zlRptFmts B" & _
    "   Where A.ID=B.����ID And A.���='ZL" & glngSys \ 100 & "_BILL_1121_7'  " & _
    "   Order by  ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With vsDelBillFormat
        .Clear 1
        .ColComboList(.ColIndex("�˷�Ʊ�ݸ�ʽ")) = .BuildComboList(rsTemp, "���,*˵��", "���")
        .ColComboList(.ColIndex("�˷Ѵ�ӡ��ʽ")) = "0-����ӡƱ��|1-�Զ���ӡƱ��|2-ѡ���Ƿ��ӡƱ��"
    End With

    '��ȡ����ֵ
    strShareInvoice = zlDatabase.GetPara("�˷ѷ�Ʊ��ʽ", glngSys, mlngModul, , , True, intType)
    strPrintMode = zlDatabase.GetPara("�˷ѷ�Ʊ��ӡ��ʽ", glngSys, mlngModul, , , True, intType1)
    '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    With vsDelBillFormat
        .ColData(.ColIndex("�˷�Ʊ�ݸ�ʽ")) = "0"
        .ColData(.ColIndex("�˷Ѵ�ӡ��ʽ")) = "0"
        .ForeColor = &H80000008:  .ForeColorFixed = &H80000008
        Select Case intType
        Case 1, 3, 5, 15
             .ColData(.ColIndex("�˷�Ʊ�ݸ�ʽ")) = IIf(intType = 5, 0, 1)
        End Select
        
        Select Case intType1
        Case 1, 3, 5, 15
             .ColData(.ColIndex("�˷Ѵ�ӡ��ʽ")) = IIf(intType1 = 5, 0, 1)
        End Select
        
        If (Val(.ColData(.ColIndex("�˷�Ʊ�ݸ�ʽ"))) = 1 Or _
            Val(.ColData(.ColIndex("�˷Ѵ�ӡ��ʽ"))) = 1) Then
            .Editable = IIf(Not blnHavePrivs, flexEDNone, flexEDKbdMouse)
        Else
            .Editable = flexEDKbdMouse
        End If
    End With
    
    vsDelBillFormat.Tag = ""
    varData = Split(strShareInvoice, "|")
    VarType = Split(strPrintMode, "|")
    strSQL = "" & _
    "   Select ���� ,����" & _
    "   From  Ʊ��ʹ�����" & _
    "   order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With vsDelBillFormat
        .Clear 1
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("ʹ�����")) = Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("�˷Ѵ�ӡ��ʽ")) = "0-����ӡƱ��"
            .TextMatrix(lngRow, .ColIndex("�˷�Ʊ�ݸ�ʽ")) = "0"
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(Nvl(rsTemp!����)) Then
                    .TextMatrix(lngRow, .ColIndex("�˷�Ʊ�ݸ�ʽ")) = Val(varTemp(1)): Exit For
                End If
            Next
            
            For i = 0 To UBound(VarType)
                varTemp1 = Split(VarType(i) & "," & ",", ",")
                If Trim(varTemp1(0)) = Trim(Nvl(rsTemp!����)) Then
                    .TextMatrix(lngRow, .ColIndex("�˷Ѵ�ӡ��ʽ")) = Decode(Val(varTemp1(1)), 0, "0-����ӡƱ��", 1, "1-�Զ���ӡƱ��", "2-ѡ���Ƿ��ӡƱ��")
                    Exit For
                End If
            Next
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If Val(.ColData(.ColIndex("�˷Ѵ�ӡ��ʽ"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("�˷Ѵ�ӡ��ʽ"), .Rows - 1, .ColIndex("�˷Ѵ�ӡ��ʽ")) = vbBlue
        End If
        
        If Val(.ColData(.ColIndex("�˷�Ʊ�ݸ�ʽ"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("�˷�Ʊ�ݸ�ʽ"), .Rows - 1, .ColIndex("�˷�Ʊ�ݸ�ʽ")) = vbBlue
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset, strSQL As String, objItem As ListItem, blnParSet As Boolean
    Dim strTmp As String, i As Integer, j As Long, k As Long, arrTmp As Variant, arrWindow As Variant, intType As Integer, blnSeted As Boolean '��������ȱʡֵ

    Dim str��ҩ������ As String, str��ҩ������ As String, str��ҩ������ As String
    Dim lngȱʡ��ҩ�� As Long, lngȱʡ��ҩ�� As Long, lngȱʡ��ҩ�� As Long, lngȱʡ���ϲ��� As Long
    
    gblnOK = False
    On Error GoTo errH
    Call InitTabControl
    blnParSet = InStr(1, mstrPrivs, "��������") > 0
    
    'a.��ʼ����
    '----------------------------------------------------------------------------------------
    '�շ����(�Һų���):���������
    strSQL = "Select ����,���� as ��� from �շ���Ŀ��� Where ����<>'1' Order by ���"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    Do While Not rsTmp.EOF
        lst�շ����.AddItem rsTmp!���
        lst�շ����.ItemData(lst�շ����.NewIndex) = Asc(rsTmp!����)
        rsTmp.MoveNext
    Loop
    If mbytInFun <> 2 Then
        strSQL = _
            " Select ����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From �ѱ�" & _
            " Where ����=1 And Nvl(���޳���,0)=0 And Nvl(�������,3) IN(1,3)" & _
            " Order by ����"
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        For i = 1 To rsTmp.RecordCount
            cbo�ѱ�.AddItem rsTmp!����
            If rsTmp!ȱʡ = 1 Then cbo�ѱ�.ListIndex = cbo�ѱ�.NewIndex
            rsTmp.MoveNext
        Next
    End If
    
    If mbytInFun = 0 Then
        lbl��λ.Caption = "�շ�ʱ����"
    ElseIf mbytInFun = 1 Then
        lbl��λ.Caption = "����ʱ����"
    ElseIf mbytInFun = 2 Then
        lbl��λ.Caption = "����ʱ����"
    End If
    
    'b.����ע���洢��ģ�����
    '----------------------------------------------------------------------------------------
    Call LoadParaValue
    'c.���ݿ�洢��ģ�����
    '----------------------------------------------------------------------------------------
    '--------------------------
    strTmp = zlDatabase.GetPara("����ҽ��", glngSys, mlngModul, , Array(optUnit, optDoctor, optSelf), blnParSet)
    If strTmp = "1" Then
        optUnit.Value = True
    ElseIf strTmp = "0" Then
        optDoctor.Value = True
    Else
        optSelf.Value = True
    End If
    
    i = IIf(zlDatabase.GetPara("��������ʾ��ʽ", glngSys, mlngModul, "1", Array(optDoctorKind(0), optDoctorKind(1)), blnParSet) = "1", 0, 1)
    optDoctorKind(i).Value = True
    
    txtNameDays.Enabled = True
    txtNameDays.Text = zlDatabase.GetPara("������������", glngSys, mlngModul, , Array(txtNameDays), blnParSet)
    txtNameDays.Tag = IIf(txtNameDays.Enabled, "1", "0")
    chkSeekName.Value = IIf(zlDatabase.GetPara("����ģ������", glngSys, mlngModul, , Array(chkSeekName), blnParSet) = "1", 1, 0)
    '����chkSeekName_click
    
    If mbytInFun = 0 Or mbytInFun = 1 Then
        txtMax.Text = Format(zlDatabase.GetPara("�����", glngSys, mlngModul, "0", Array(txtMax), blnParSet), "0.00")
        
        chk�Ա�.Value = IIf(zlDatabase.GetPara("�Ա�", glngSys, mlngModul, , Array(chk�Ա�), blnParSet) = "1", 1, 0)
        chk����.Value = IIf(zlDatabase.GetPara("����", glngSys, mlngModul, , Array(chk����), blnParSet) = "1", 1, 0)
        chk�ѱ�.Value = IIf(zlDatabase.GetPara("�ѱ�", glngSys, mlngModul, , Array(chk�ѱ�), blnParSet) = "1", 1, 0)
        chkҽ�Ƹ���.Value = IIf(zlDatabase.GetPara("ҽ�Ƹ���", glngSys, mlngModul, , Array(chkҽ�Ƹ���), blnParSet) = "1", 1, 0)
        chk�Ƿ�Ӱ�.Value = IIf(zlDatabase.GetPara("�Ӱ�", glngSys, mlngModul, , Array(chk�Ƿ�Ӱ�), blnParSet) = "1", 1, 0)
        chk��������.Value = IIf(zlDatabase.GetPara("��������", glngSys, mlngModul, , Array(chk��������), blnParSet) = "1", 1, 0)
        chk������.Value = IIf(zlDatabase.GetPara("������", glngSys, mlngModul, , Array(chk������), blnParSet) = "1", 1, 0)
                
        chk��ȱʡ������.Value = IIf(zlDatabase.GetPara("��ʹ��ȱʡ������", glngSys, mlngModul, , Array(chk��ȱʡ������), blnParSet) = "1", 1, 0)
        chk�����俪����.Value = IIf(zlDatabase.GetPara("����Ҫ���뿪����", glngSys, mlngModul, , Array(chk�����俪����), blnParSet) = "1", 1, 0)
        chkȱʡ��������.Value = IIf(zlDatabase.GetPara("ȱʡ��������", glngSys, mlngModul, , Array(chkȱʡ��������), blnParSet) = "1", 1, 0)
        chkȱʡ��������.Left = chk��ȱʡ������.Left
        Call optUnit_Click
        
        i = Val(zlDatabase.GetPara("����ϼƷ�ʽ", glngSys, mlngModul, , Array(opt����(0), opt����(1), opt����(2)), blnParSet))  '34179
        If i > 2 Or i < 0 Then i = 0
        opt����(i).Value = True
        
        If mbytInFun = 0 Then
            chk�ۼ�.Value = IIf(zlDatabase.GetPara("��ʾ�ۼ�", glngSys, mlngModul, , Array(chk�ۼ�), blnParSet) = "1", 1, 0)
            chkƤ��.Value = IIf(zlDatabase.GetPara("���Ƥ�Խ��", glngSys, mlngModul, , Array(chkƤ��), blnParSet) = "1", 1, 0)
            chkPrePayPriority.Value = IIf(zlDatabase.GetPara("����ʹ��Ԥ����", glngSys, mlngModul, , Array(chkPrePayPriority), blnParSet) = "1", 1, 0)
            '91665
            chkInsurePartFee.Value = IIf(zlDatabase.GetPara("ֻ��ҽ������ɹ������շ�", glngSys, mlngModul, , Array(chkInsurePartFee), blnParSet) = "1", 1, 0)
            '96357
            strTmp = zlDatabase.GetPara("�����շ�ִ�п���", glngSys, mlngModul, , Array(chk�շ�ִ�п���, txt�շ�ִ�п���, cmd�շ�ִ�п���), blnParSet)
            mblnNotClick = True
            chk�շ�ִ�п���.Value = IIf(strTmp <> "", vbChecked, vbUnchecked)
            mblnNotClick = False
            cmd�շ�ִ�п���.Enabled = chk�շ�ִ�п���.Value = vbChecked
            txt�շ�ִ�п���.Text = GetDeptNameStr(strTmp)
            txt�շ�ִ�п���.Tag = strTmp
            '���˺�:22343:51670
             i = Val(zlDatabase.GetPara("�շѽɿ��������", glngSys, mlngModul, , Array(opt�ɿ�(0), opt�ɿ�(1), opt�ɿ�(2), opt�ɿ�(3)), blnParSet))
             If i <= opt�ɿ�.UBound And i >= opt�ɿ�.LBound Then opt�ɿ�(i).Value = True
            strTmp = zlDatabase.GetPara("�Զ����չҺŷ�", glngSys, mlngModul, , Array(chkAddedItem, txtAddedItem, cmdAddedItem), blnParSet)
            If InStr(1, strTmp, ";") > 0 Then
                chkAddedItem.Value = 1  '�����click�¼�,���ȼ����շ����
                txtAddedItem.Tag = Split(strTmp, ";")(0)
                txtAddedItem.Text = Split(strTmp, ";")(1)
            End If
'            chkShowError.Value = IIf(zlDatabase.GetPara("��ʾ������", glngSys, mlngModul, , Array(chkShowError), blnParSet) = "1", 1, 0)
            chkMulti.Value = IIf(zlDatabase.GetPara("�൥���շ�", glngSys, mlngModul, , Array(chkMulti), blnParSet) = "1", 1, 0)
            
            chkSeekBill.Enabled = True
            chkSeekBill.Value = IIf(zlDatabase.GetPara("��Ѱ���۵���", glngSys, mlngModul, , Array(chkSeekBill), blnParSet) = "1", 1, 0)
            chkSeekBill.Tag = IIf(chkSeekBill.Enabled, "1", "0")
            txtSeekDays.Enabled = True
            txtSeekDays.Text = zlDatabase.GetPara("��Ѱ��������", glngSys, mlngModul, , Array(txtSeekDays), blnParSet)
            txtSeekDays.Tag = IIf(txtSeekDays.Enabled, "1", "0")
            chkUnPopPriceBill.Enabled = True
            chkUnPopPriceBill.Value = IIf(zlDatabase.GetPara("���������۵�ѡ��", glngSys, mlngModul, , Array(chkUnPopPriceBill), blnParSet) = "1", 1, 0)
            chkUnPopPriceBill.Tag = IIf(chkUnPopPriceBill.Enabled, "1", "0")
            
            chkMustRegevent.Value = IIf(zlDatabase.GetPara("��鲡�˹Һſ���", glngSys, mlngModul, , Array(chkMustRegevent), blnParSet) = "1", 1, 0)
            i = Val(zlDatabase.GetPara("δ�ҺŲ����շ�", glngSys, mlngModul, , Array(optRegPrompt(0), optRegPrompt(1), optRegPrompt(2)), blnParSet))
            If i <= optRegPrompt.UBound Then optRegPrompt(i).Value = True
            
            chkAutoSplitBill.Enabled = True
            cboAutoSplitBill.Enabled = True
            i = Val(zlDatabase.GetPara("�Զ���ϵ���", glngSys, mlngModul, , Array(chkAutoSplitBill, cboAutoSplitBill), blnParSet))
            chkAutoSplitBill.Tag = IIf(chkAutoSplitBill.Enabled, "1", "0")
            cboAutoSplitBill.Tag = IIf(cboAutoSplitBill.Enabled, "1", "0")
            chkAutoSplitBill.Value = IIf(i > 0, 1, 0)
            cboAutoSplitBill.AddItem "�շ����"
            cboAutoSplitBill.AddItem "ִ�п���"
            cboAutoSplitBill.ListIndex = IIf(i = 1 Or i = 2, i - 1, 0)
            If chkAutoSplitBill.Value = 0 Then cboAutoSplitBill.Enabled = False
            Call chkMulti_Click
                
            i = Val(zlDatabase.GetPara("�շ��嵥��ӡ��ʽ", glngSys, mlngModul, , Array(optPrint(0), optPrint(1), optPrint(2)), blnParSet))
            If i <= optPrint.UBound Then optPrint(i).Value = True
            i = Val(zlDatabase.GetPara("�˷ѻص���ӡ��ʽ", glngSys, mlngModul, , Array(optRefund(0), optRefund(1), optRefund(2)), blnParSet))
            If i <= optRefund.UBound Then optRefund(i).Value = True
            '62982:���ϴ�,2015/08/25,�շ�ִ�е�
            i = Val(zlDatabase.GetPara("�շ�ִ�е���ӡ��ʽ", glngSys, mlngModul, , Array(optExe(0), optExe(1), optExe(2)), blnParSet))
            If i <= optExe.UBound Then optExe(i).Value = True
            
            '���˺� ����:26948 ����:2009-12-28 16:54:11
            strTmp = zlDatabase.GetPara("Ʊ��ʣ��X��ʱ��ʼ�����շ�Ա", glngSys, mlngModul, "0|10", Array(txtƱ������, updƱ������, chkƱ������), blnParSet)
            
            updƱ������.Value = Val(Split(strTmp & "|", "|")(1))
            txtƱ������.Text = updƱ������.Value
            chkƱ������.Value = IIf(Val(Split(strTmp & "|", "|")(0)) = 1, 1, 0)
            txtƱ������.Enabled = chkƱ������.Enabled And chkƱ������.Value = 1
            updƱ������.Enabled = txtƱ������.Enabled
        Else
            txtDay.Text = zlDatabase.GetPara("ȡ�����۵�", glngSys, mlngModul, , Array(txtDay, lblDay, udDay), blnParSet)
            
            i = Val(zlDatabase.GetPara("����֪ͨ����ӡ��ʽ", glngSys, mlngModul, , Array(optPrintRequisition(0), optPrintRequisition(1), optPrintRequisition(2)), blnParSet))
            If i <= optPrintRequisition.UBound Then optPrintRequisition(i).Value = True
        End If
    ElseIf mbytInFun = 2 Then
        chkOnlyUnitPatient.Enabled = True
        chkOnlyUnitPatient.Value = IIf(zlDatabase.GetPara("ֻ���Һ�Լ��λ����", glngSys, mlngModul, , Array(chkOnlyUnitPatient), blnParSet) = "1", 1, 0)
        chkOnlyUnitPatient.Tag = IIf(chkOnlyUnitPatient.Enabled, "1", "0")
        chkOnlyUnitPatient.Enabled = chkSeekName.Value = 1 And chkOnlyUnitPatient.Tag = "1"
        
        Call chkSeekName_Click
        
        chk(0).Value = IIf(zlDatabase.GetPara("���ʴ�ӡ", glngSys, mlngModul, , Array(chk(0)), blnParSet) = "1", 1, 0)
        chk(1).Value = IIf(zlDatabase.GetPara("���۴�ӡ", glngSys, mlngModul, , Array(chk(1)), blnParSet) = "1", 1, 0)
        chk(2).Value = IIf(zlDatabase.GetPara("��˴�ӡ", glngSys, mlngModul, , Array(chk(2)), blnParSet) = "1", 1, 0)
    End If
    
    '���˺� ����:26948 ����:2009-12-28 16:54:11
    If mbytInFun <> 0 Then
        txtƱ������.Visible = False: updƱ������.Visible = False: chkƱ������.Visible = False
    End If
    
    Call SetCtrlVisible
    
    'd.Ȩ�޿���
    '----------------------------------------------------------------------------------------
    chkLed.Visible = mbytInFun = 0
    If InStr(mstrPrivs, "LED������") = 0 Then
        chkLed.Visible = False
        chkLed.Value = Unchecked
    End If
    chkLedDispDetail.Visible = mbytInFun = 0
    chkLedWelcome.Visible = mbytInFun = 0
    chkҽ��������ȱʡ��λ.Visible = mbytInFun = 0
    
    chkƤ��.Visible = mbytInFun = 0
    chkPrePayPriority.Visible = mbytInFun = 0
    lbl���㷽ʽ.Visible = mbytInFun = 0
    cbo���㷽ʽ.Visible = mbytInFun = 0
    
    chkAddedItem.Visible = mbytInFun = 0
    txtAddedItem.Visible = mbytInFun = 0
    cmdAddedItem.Visible = mbytInFun = 0
    
    chkInsurePartFee.Visible = mbytInFun = 0
    chk��ֹȡ���Һŵ�.Visible = mbytInFun = 0
    
    '���ݲ�����Ѱ���۵���
    chkSeekBill.Visible = mbytInFun = 0
    txtSeekDays.Visible = mbytInFun = 0
    fraLine.Visible = mbytInFun = 0
    chkUnPopPriceBill.Visible = mbytInFun = 0
    
    '֧�ֶ൥���շ�,��ʾ������
    fraRegPrompt.Visible = mbytInFun = 0
    chkMustRegevent.Visible = mbytInFun = 0
    chkMulti.Visible = mbytInFun = 0
    chkAutoSplitBill.Visible = mbytInFun = 0
    cboAutoSplitBill.Visible = mbytInFun = 0
'    chkShowError.Visible = mbytInFun = 0
    
    '���˺� �����:22343
    fra�ɿ����.Visible = mbytInFun = 0
    chk�ۼ�.Visible = mbytInFun = 0
    fra����֪ͨ����ӡ.Visible = mbytInFun = 1
    cmdPrintSetup(3).Visible = mbytInFun = 1
    lblDay.Visible = mbytInFun = 1
    txtDay.Visible = mbytInFun = 1
    udDay.Visible = mbytInFun = 1
    
    chkOnlyUnitPatient.Visible = mbytInFun = 2
    fraPrintBill.Visible = mbytInFun = 2
        
    lbl�ѱ�.Visible = mbytInFun <> 2
    cbo�ѱ�.Visible = mbytInFun <> 2
    lblMax.Visible = mbytInFun <> 2
    txtMax.Visible = mbytInFun <> 2
    
    stab.TabVisible(1) = mbytInFun <> 2
    stab.TabVisible(2) = mbytInFun = 0
    stab.TabVisible(3) = mbytInFun = 0  '56963
    
    txt�շ�ִ�п���.Visible = mbytInFun = 0
    chk�շ�ִ�п���.Visible = mbytInFun = 0
    cmd�շ�ִ�п���.Visible = mbytInFun = 0
    
    'f.λ�õ���
    '-------------------------------------------------------------
    Call MoveCtrol
    'ҩ������
    If glngSys Like "8??" Then
        fra����.Visible = False
                
        lbl�ѱ�.Caption = "ȱʡ��Ա�ȼ�"
                
        chkҩ��.Caption = "��ʾ����ҩ����"
        fra���.Visible = False '�̶�����ҩƷ���
        fra������ҽ��.Visible = False '�̶���������
        
        chk��ʿ.Visible = False
        chk��ʿ.Value = 0
    End If
    
    If mblnSetDrugStore Then
        '56963
        stab.TabCaption(4) = "ҩ������"
        stab.TabVisible(0) = False
        stab.TabVisible(1) = False
        stab.TabVisible(2) = False
    Else
        If mbytInFun = 1 Then
            stab.TabCaption(4) = "ҩ������(&3)"
        ElseIf mbytInFun = 2 Then
            stab.TabCaption(4) = "ҩ������(&2)"
        End If
    End If
    
    If stab.TabVisible(0) Then stab.Tab = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetDeptNameStr(ByVal strIDs As String) As String
    '������ID�ַ���װ���������ַ���
    '��Σ�
    '   strIDs ����ID����ʽ��ID1,ID2,ID3,...
    '���أ�
    '   ��������s����ʽ����������1;��������2;��������3;...
    Dim strSQL As String, rsTemp As Recordset
    Dim strTemp As String
    
    Err = 0: On Error GoTo errHandler
    If strIDs = "" Then Exit Function
    strSQL = "Select /*+cardinality(b,10) */a.����, a.����" & vbNewLine & _
            " From ���ű� A, Table(f_Str2list([1], ',')) B" & vbNewLine & _
            " Where a.Id = b.Column_Value"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "���ݿ���ID��ȡ��������", strIDs)
    If rsTemp Is Nothing Then Exit Function
    
    Do While Not rsTemp.EOF
        strTemp = strTemp & ";" & Nvl(rsTemp!����)
        rsTemp.MoveNext
    Loop
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    
    GetDeptNameStr = strTemp
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    mblnSetDrugStore = False
    If mbytInFun = 0 Then
        zl_vsGrid_Para_Save mlngModul, vsBill, Me.Name, "����Ʊ��������", False, False
        zl_vsGrid_Para_Save mlngModul, vsBillFormat, Me.Name, "�շ�Ʊ�ݸ�ʽ", False, False
        zl_vsGrid_Para_Save mlngModul, vsDelBillFormat, Me.Name, "�˷�Ʊ�ݸ�ʽ", False, False
    End If
End Sub

Private Sub lst�շ����_ItemCheck(Item As Integer)
    If lst�շ����.SelCount = 0 And Not lst�շ����.Selected(Item) Then
        lst�շ����.Selected(Item) = True
    End If
End Sub
Private Sub optDoctor_Click()
    Call optUnit_Click
End Sub

Private Sub optSelf_Click()
    Call optUnit_Click
End Sub

Private Sub optUnit_Click()
    chk��ȱʡ������.Visible = optUnit.Value
    chkȱʡ��������.Visible = optDoctor.Value
End Sub

Private Sub opt����_Click(Index As Integer)
    If gbln���뷢ҩ Then
        Call SetStockCheck
    Else
        Call SetDrugStore
    End If
    
    If opt����(0).Value Then
        opt��λ(1).Caption = "���ﵥλ"
    Else
        opt��λ(1).Caption = "סԺ��λ"
    End If
End Sub

Private Sub stab_Click(PreviousTab As Integer)
    Select Case stab.Tab
        Case 0
            If txtMax.Visible Then
                If txtMax.Enabled And txtMax.Visible Then txtMax.SetFocus
            Else
                If cbo�ѱ�.Enabled And cbo�ѱ�.Visible Then cbo�ѱ�.SetFocus
            End If
        Case 1
            If opt����(0).Enabled And opt����(0).Visible And opt����(0).Value Then
                opt����(0).SetFocus
            ElseIf opt����(1).Enabled And opt����(1).Visible And opt����(1).Value Then
                opt����(1).SetFocus
            End If
        Case 2
            If vsBill.Enabled And vsBill.Visible Then vsBill.SetFocus
        Case 3
            ' 56963
            If cboBillRole.Visible And cboBillRole.Enabled Then cboBillRole.SetFocus
        Case 4
            ' 56963
            If vsfDrugStore.Visible And vsfDrugStore.Enabled Then vsfDrugStore.SetFocus
    End Select
End Sub

Private Sub txtBillRuleNum_Change(Index As Integer)
    '56963
    Call ShowRuleInfor
End Sub

Private Sub txtDay_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtDay_LostFocus()
    If Not IsNumeric(txtDay.Text) Then txtDay.Text = 0
End Sub

Private Sub txtMax_GotFocus()
    SelAll txtMax
End Sub

Private Sub txtMax_KeyPress(KeyAscii As Integer)
    If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtMax_LostFocus()
    If IsNumeric(txtMax.Text) Then
        txtMax.Text = Format(txtMax.Text, "0.00")
    Else
        txtMax.Text = "0.00"
    End If
End Sub

Private Sub txtSeekDays_GotFocus()
    Call SelAll(txtSeekDays)
End Sub

Private Sub txtSeekDays_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtSeekDays_Validate(Cancel As Boolean)
    If Val(txtSeekDays.Text) < 1 Then
        txtSeekDays.Text = 1
    End If
End Sub

Private Sub txtNameDays_GotFocus()
    Call SelAll(txtNameDays)
End Sub

Private Sub txtNameDays_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNameDays_Validate(Cancel As Boolean)
    If Val(txtNameDays.Text) <= 0 Then
        txtNameDays.Text = 0
    ElseIf Val(txtNameDays.Text) > 999 Then
        txtNameDays.Text = 999
    End If
End Sub
Private Sub updBillRuleNum_Change(Index As Integer)
    If updBillRuleNum(Index).Value = 0 Then
        chkBillRule(Index + 1).Value = 0
    End If
End Sub

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
        zl_vsGrid_Para_Save mlngModul, vsBillFormat, Me.Name, "�շ�Ʊ�ݸ�ʽ", False, False
End Sub
Private Sub vsBillFormat_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
        zl_vsGrid_Para_Save mlngModul, vsBillFormat, Me.Name, "�շ�Ʊ�ݸ�ʽ", False, False
End Sub

Private Sub vsBillFormat_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsBillFormat
        Select Case Col
        Case .ColIndex("�շ�Ʊ�ݸ�ʽ"), .ColIndex("�����˲���Ʊ�ݸ�ʽ"), .ColIndex("�շѴ�ӡ��ʽ")
            If Val(.ColData(Col)) = 1 And InStr(1, mstrPrivs, ";��������;") = 0 Then Cancel = True: Exit Sub
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub vsDelBillFormat_AfterMoveColumn(ByVal Col As Long, Position As Long)
        zl_vsGrid_Para_Save mlngModul, vsDelBillFormat, Me.Name, "�˷�Ʊ�ݸ�ʽ", False, False
End Sub
Private Sub vsDelBillFormat_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
        zl_vsGrid_Para_Save mlngModul, vsDelBillFormat, Me.Name, "�˷�Ʊ�ݸ�ʽ", False, False
End Sub

Private Sub vsDelBillFormat_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsDelBillFormat
        Select Case Col
        Case .ColIndex("�˷�Ʊ�ݸ�ʽ"), .ColIndex("�˷Ѵ�ӡ��ʽ")
            If Val(.ColData(Col)) = 1 And InStr(1, mstrPrivs, ";��������;") = 0 Then Cancel = True: Exit Sub
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub vsfDrugStore_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsfDrugStore
        Select Case Col
        Case .ColIndex("ȱʡ")
           Call SetDrugStockDeFault(Row)
        Case Else
            Exit Sub
        End Select
    End With
End Sub

Private Sub vsfDrugStore_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfDrugStore
        Select Case Col
        Case .ColIndex("ȱʡ"), .ColIndex("����")
            Cancel = Val(.Cell(flexcpData, Row, Col)) = 1
        Case Else
            Cancel = True
            Exit Sub
        End Select
    End With
End Sub

Private Sub vsfDrugStore_DblClick()
    Dim strTmp As String, i As Long
    With vsfDrugStore
        If Not (.Row > 0 And .Col = 1) Then Exit Sub
        If .Cell(flexcpData, .Row, .ColIndex("ȱʡ")) = 1 Then Exit Sub
        
        .TextMatrix(.Row, .Col) = IIf(Val(.TextMatrix(.Row, .Col)) = 0, 1, 0)
        Call SetDrugStockDeFault(.Row)
    End With
End Sub
Private Sub SetDrugStockDeFault(ByVal lngRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҩ����ȱʡֵ
    '���:lngRow-ָ����
    '����:���˺�
    '����:2009-09-02 14:37:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, lngȱʡ As Long, strType As String
    With vsfDrugStore
        lngȱʡ = Abs(Val(.TextMatrix(lngRow, .ColIndex("ȱʡ"))))
        If lngȱʡ = 1 Then
            strType = .TextMatrix(lngRow, .ColIndex("���"))
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 0) = strType And i <> lngRow Then
                    .TextMatrix(i, .ColIndex("ȱʡ")) = 0
                End If
            Next
        End If
    End With
End Sub
Private Sub SetDrugStockEdit(ByVal strType As String, ByVal intType As Integer, ByVal lngEditCol As Long, Optional strMachValue As String = "", Optional strDefaultValue As String = "")
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҩ���ı༭����
    '���:strType-���
    '     intType-���ز������ͣ�1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    '     lngEditCol-���Ƶı༭��
    '����:
    '����:
    '����:���˺�
    '����:2009-09-02 14:53:10
    '����:25132
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, blnSetDefault As Boolean '������ȱʡֵ��,�����������ȱʡֵ
    Dim lngEditForColor As Long, blnAllowEdit As Boolean, bytLockEdit As Integer '1-����,0-������
    
    '���˺�:���ڿ��ܲ���Ȩ�޷������,���,����ͳһ��������,��Ҫ����ĳһ����:
    With vsfDrugStore
        blnSetDefault = False: blnAllowEdit = InStr(1, mstrPrivs, ";��������;") > 0
        bytLockEdit = 0
        If InStr(1, ",1,3,15,", "," & intType & ",") > 0 Then
            lngEditForColor = IIf(blnAllowEdit, vbBlue, &H8000000C)  '��Ȩ�޿���
            bytLockEdit = IIf(blnAllowEdit, 0, 1)
        ElseIf intType = 5 Then
            lngEditForColor = vbBlue    '����ģ��,������Ȩ�޿���
        Else
            lngEditForColor = &H80000008    '�����༭
        End If
        
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("���")) = strType Then
                If lngEditCol = .ColIndex("ȱʡ") Then
                    '����ҩ��
                    If Val(.RowData(i)) = Val(strMachValue) And strMachValue <> "" And Not blnSetDefault Then
                        .TextMatrix(i, .ColIndex("ȱʡ")) = IIf(Val(strMachValue) > 0, 1, 0)
                        blnSetDefault = True
                    End If
                     .Cell(flexcpForeColor, i, .ColIndex("ȱʡ")) = lngEditForColor
                     .Cell(flexcpForeColor, i, .ColIndex("ҩ��")) = lngEditForColor:
                Else
                    If Val(.RowData(i)) = Val(strMachValue) And strMachValue <> "" And Not blnSetDefault Then
                        .TextMatrix(i, lngEditCol) = strDefaultValue
                    End If
                    '���ô���
                     .Cell(flexcpForeColor, i, .ColIndex("����")) = lngEditForColor
                End If
                .Cell(flexcpData, i, lngEditCol) = bytLockEdit
            End If
        Next
    End With
End Sub

Private Sub vsfDrugStore_EnterCell()
    Dim rsTmp As ADODB.Recordset, strList As String
    With vsfDrugStore
        If .Row > 0 Then
            If .Col = .ColIndex("����") Then
                Set rsTmp = Read��ҩ����(.RowData(.Row))
                strList = "�Զ�����|" & .BuildComboList(rsTmp, "����")
                .ColComboList(.Col) = strList
            Else
                .ColComboList(.Col) = ""
              '  .Editable = flexEDNone
            End If
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub

Private Sub InitTabControl()
    With tbBillSet
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Position = xtpTabPositionTop
'        .PaintManager.StaticFrame = True
'        .PaintManager.ClientFrame = xtpTabFrameSingleLine
        .InsertItem 0, "�շ�Ʊ�ݸ�ʽ", picBillFormat.hWnd, 0
        .InsertItem 1, "�˷�Ʊ�ݸ�ʽ", picDelBillFormat.hWnd, 0
        .Item(0).Selected = True
    End With
End Sub
