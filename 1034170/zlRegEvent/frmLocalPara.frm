VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLocalPara 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������������"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   ControlBox      =   0   'False
   Icon            =   "frmLocalPara.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin TabDlg.SSTab Tabs1 
      Height          =   6585
      Left            =   120
      TabIndex        =   95
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   11615
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   617
      TabCaption(0)   =   "����(&0)"
      TabPicture(0)   =   "frmLocalPara.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblSortMode"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblColor"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lstDept"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraDefaultSet"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraClearMZInfor"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cboSortMode"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "pic��ǰ��ɫ"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "�������(&1)"
      TabPicture(1)   =   "frmLocalPara.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblGuardian"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Line1(6)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "chkSeekName"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "fraTitle"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "fraInput"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdDeviceSetup"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "chkPrintFree"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "chkTotal"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "chkDoctor"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "fraLine2"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "txtNameDays"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "fraInvoice"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "chkPrintCase"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "fraDeposit"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "chkAddressAssnInput"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "txtMustGuardianInfo"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "fraSlip"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "fra�˺Żص���ӡ��ʽ"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).ControlCount=   18
      TabCaption(2)   =   "ҽ�ƿ�(&2)"
      TabPicture(2)   =   "frmLocalPara.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblDefaultPayCard"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "chkAutoAddName"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "chkCardMoney"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "chkNewCardNoPop"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "chkRePrint"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cboType"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "fraCards"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "chkScanIDVisa"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "chkAlwaysSendCard"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "ChkMustBill"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).ControlCount=   10
      TabCaption(3)   =   "ԤԼ�Һ�(&3)"
      TabPicture(3)   =   "frmLocalPara.frx":0060
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "lblAvailabilityTimes"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label7"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label8"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "lblBreakAnAppointmentNums"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Line1(0)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Line1(2)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "lblCancelBespeak"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Line1(3)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Line1(4)"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "lblBespeakMinTime"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "lblBespeakDefaultDays"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "Line1(5)"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "Line1(1)"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "Line1(7)"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "Label1"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "chkDeptNums"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "txtAvailabilityTimes"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).Control(17)=   "chkDeptBespeakOneNum"
      Tab(3).Control(17).Enabled=   0   'False
      Tab(3).Control(18)=   "Frame3"
      Tab(3).Control(18).Enabled=   0   'False
      Tab(3).Control(19)=   "Frame5"
      Tab(3).Control(19).Enabled=   0   'False
      Tab(3).Control(20)=   "txtBreakAnAppointmentNums"
      Tab(3).Control(20).Enabled=   0   'False
      Tab(3).Control(21)=   "chkBespeakFee"
      Tab(3).Control(21).Enabled=   0   'False
      Tab(3).Control(22)=   "chkMzh"
      Tab(3).Control(22).Enabled=   0   'False
      Tab(3).Control(23)=   "chkBackNoToVerfy"
      Tab(3).Control(23).Enabled=   0   'False
      Tab(3).Control(24)=   "chkBreakAnAppointmentToRegist"
      Tab(3).Control(24).Enabled=   0   'False
      Tab(3).Control(25)=   "txtCancelBespeak"
      Tab(3).Control(25).Enabled=   0   'False
      Tab(3).Control(26)=   "txtBespeakMinTime"
      Tab(3).Control(26).Enabled=   0   'False
      Tab(3).Control(27)=   "txtBespeakDefaultDays"
      Tab(3).Control(27).Enabled=   0   'False
      Tab(3).Control(28)=   "fraBespeak"
      Tab(3).Control(28).Enabled=   0   'False
      Tab(3).Control(29)=   "fraReceiveMode"
      Tab(3).Control(29).Enabled=   0   'False
      Tab(3).Control(30)=   "txtDeptNums"
      Tab(3).Control(30).Enabled=   0   'False
      Tab(3).Control(31)=   "txtDeptBespeakOneNum"
      Tab(3).Control(31).Enabled=   0   'False
      Tab(3).Control(32)=   "cboDefaultStyle"
      Tab(3).Control(32).Enabled=   0   'False
      Tab(3).Control(33)=   "cboԤԼ��Чʱ��"
      Tab(3).Control(33).Enabled=   0   'False
      Tab(3).ControlCount=   34
      Begin VB.Frame fra�˺Żص���ӡ��ʽ 
         Caption         =   "�˺Żص���ӡ��ʽ"
         Height          =   645
         Left            =   -74760
         TabIndex        =   127
         Top             =   4470
         Width           =   6480
         Begin VB.CommandButton cmdPrintSet 
            Caption         =   "��ӡ����"
            Height          =   345
            Index           =   4
            Left            =   5160
            TabIndex        =   60
            Top             =   210
            Width           =   990
         End
         Begin VB.OptionButton optReceipt 
            Caption         =   "ѡ���Ƿ��ӡ"
            Height          =   315
            Index           =   2
            Left            =   2640
            TabIndex        =   59
            Top             =   240
            Width           =   1425
         End
         Begin VB.OptionButton optReceipt 
            Caption         =   "�Զ���ӡ"
            Height          =   225
            Index           =   1
            Left            =   1320
            TabIndex        =   58
            Top             =   300
            Width           =   1275
         End
         Begin VB.OptionButton optReceipt 
            Caption         =   "����ӡ"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   57
            Top             =   300
            Width           =   915
         End
      End
      Begin VB.PictureBox pic��ǰ��ɫ 
         BackColor       =   &H00000000&
         Height          =   270
         Left            =   -68415
         ScaleHeight     =   210
         ScaleWidth      =   210
         TabIndex        =   126
         Top             =   5820
         Width           =   270
      End
      Begin VB.ComboBox cboԤԼ��Чʱ�� 
         ForeColor       =   &H80000012&
         Height          =   300
         Left            =   3270
         Style           =   2  'Dropdown List
         TabIndex        =   83
         Top             =   4110
         Width           =   780
      End
      Begin VB.CheckBox ChkMustBill 
         Caption         =   "�ϸ�����¿���Ϊ��Ҳ��Ʊ��"
         Height          =   195
         Left            =   -74625
         TabIndex        =   68
         Top             =   2520
         Width           =   3345
      End
      Begin VB.ComboBox cboDefaultStyle 
         Height          =   300
         Left            =   1830
         Style           =   2  'Dropdown List
         TabIndex        =   82
         Top             =   3405
         Width           =   2160
      End
      Begin VB.ComboBox cboSortMode 
         ForeColor       =   &H80000012&
         Height          =   300
         Left            =   -69480
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   5445
         Width           =   1350
      End
      Begin VB.TextBox txtDeptBespeakOneNum 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   2460
         MaxLength       =   2
         TabIndex        =   74
         Text            =   "0"
         Top             =   1200
         Width           =   660
      End
      Begin VB.TextBox txtDeptNums 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   2670
         MaxLength       =   2
         TabIndex        =   72
         Text            =   "0"
         Top             =   885
         Width           =   660
      End
      Begin VB.Frame fraReceiveMode 
         Caption         =   "ԤԼ����ģʽ"
         Height          =   615
         Left            =   480
         TabIndex        =   123
         Top             =   5835
         Width           =   6210
         Begin VB.OptionButton optReceiveMode 
            Caption         =   "��ԤԼ����"
            Height          =   255
            Index           =   1
            Left            =   2865
            TabIndex        =   91
            Top             =   255
            Width           =   3165
         End
         Begin VB.OptionButton optReceiveMode 
            Caption         =   "ԤԼ���վ���"
            Height          =   255
            Index           =   0
            Left            =   675
            TabIndex        =   90
            Top             =   255
            Width           =   2130
         End
      End
      Begin VB.CheckBox chkAlwaysSendCard 
         Caption         =   "���ϸ���ƿ�ʱʼ��Ϊ����"
         Height          =   195
         Left            =   -74625
         TabIndex        =   67
         Top             =   2190
         Width           =   3345
      End
      Begin VB.Frame fraSlip 
         Caption         =   "�Һ�ƾ��"
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74760
         TabIndex        =   122
         Top             =   3180
         Width           =   6480
         Begin VB.OptionButton optSlipPrint 
            Caption         =   "ѡ���Ƿ��ӡ"
            Height          =   180
            Index           =   2
            Left            =   2640
            TabIndex        =   51
            Top             =   285
            Width           =   1380
         End
         Begin VB.OptionButton optSlipPrint 
            Caption         =   "�Զ���ӡ"
            Height          =   180
            Index           =   1
            Left            =   1305
            TabIndex        =   50
            Top             =   285
            Width           =   1020
         End
         Begin VB.OptionButton optSlipPrint 
            Caption         =   "����ӡ"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   49
            Top             =   285
            Value           =   -1  'True
            Width           =   900
         End
         Begin VB.CommandButton cmdPrintSet 
            Caption         =   "��ӡ����"
            Height          =   345
            Index           =   3
            Left            =   5160
            TabIndex        =   52
            Top             =   180
            Width           =   990
         End
      End
      Begin VB.TextBox txtMustGuardianInfo 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   -74760
         MaxLength       =   2
         TabIndex        =   35
         Text            =   "0"
         Top             =   2100
         Width           =   660
      End
      Begin VB.Frame fraBespeak 
         Caption         =   "ԤԼ�Һŵ�"
         Height          =   615
         Left            =   480
         TabIndex        =   120
         Top             =   5085
         Width           =   6210
         Begin VB.OptionButton optPrintBespeak 
            Caption         =   "�Զ���ӡ"
            Height          =   180
            Index           =   1
            Left            =   1755
            TabIndex        =   87
            Top             =   300
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton optPrintBespeak 
            Caption         =   "����ӡ"
            Height          =   180
            Index           =   0
            Left            =   675
            TabIndex        =   86
            Top             =   315
            Width           =   900
         End
         Begin VB.OptionButton optPrintBespeak 
            Caption         =   "ѡ���Ƿ��ӡ"
            Height          =   180
            Index           =   2
            Left            =   2850
            TabIndex        =   88
            Top             =   300
            Width           =   1380
         End
         Begin VB.CommandButton cmdPrintSet 
            Caption         =   "��ӡ����"
            Height          =   345
            Index           =   2
            Left            =   5100
            TabIndex        =   89
            Top             =   180
            Width           =   990
         End
      End
      Begin VB.CheckBox chkScanIDVisa 
         Caption         =   "ɨ�����֤ǩԼ"
         Height          =   195
         Left            =   -74625
         TabIndex        =   66
         Top             =   1875
         Width           =   3345
      End
      Begin VB.CheckBox chkAddressAssnInput 
         Caption         =   "��ͥ��ַ��������"
         Height          =   255
         Left            =   -74760
         TabIndex        =   34
         ToolTipText     =   "�Һ�ʱ��ͥ��ַ����ʱ�Ƿ�����"
         Top             =   1800
         Value           =   1  'Checked
         Width           =   1980
      End
      Begin VB.Frame fraDeposit 
         Caption         =   "���������ӡ"
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   -74760
         TabIndex        =   119
         Top             =   3825
         Width           =   6480
         Begin VB.CommandButton cmdPrintSet 
            Caption         =   "��ӡ����"
            Height          =   345
            Index           =   0
            Left            =   5160
            TabIndex        =   56
            Top             =   180
            Width           =   990
         End
         Begin VB.OptionButton optPrepayPrint 
            Caption         =   "����ӡ"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   53
            Top             =   285
            Value           =   -1  'True
            Width           =   900
         End
         Begin VB.OptionButton optPrepayPrint 
            Caption         =   "�Զ���ӡ"
            Height          =   180
            Index           =   1
            Left            =   1305
            TabIndex        =   54
            Top             =   285
            Width           =   1020
         End
         Begin VB.OptionButton optPrepayPrint 
            Caption         =   "ѡ���Ƿ��ӡ"
            Height          =   180
            Index           =   2
            Left            =   2640
            TabIndex        =   55
            Top             =   285
            Width           =   1380
         End
      End
      Begin VB.TextBox txtBespeakDefaultDays 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   81
         Text            =   "0"
         Top             =   3060
         Width           =   540
      End
      Begin VB.TextBox txtBespeakMinTime 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   1800
         MaxLength       =   4
         TabIndex        =   80
         Text            =   "0"
         Top             =   2640
         Width           =   540
      End
      Begin VB.TextBox txtCancelBespeak 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   77
         Text            =   "0"
         Top             =   1875
         Width           =   660
      End
      Begin VB.CheckBox chkBreakAnAppointmentToRegist 
         Caption         =   "ԤԼʧԼ���ڹҺ�"
         Height          =   210
         Left            =   3960
         TabIndex        =   78
         Top             =   1890
         Width           =   2055
      End
      Begin VB.CheckBox chkBackNoToVerfy 
         Caption         =   "�˺����:N����ȡ��ԤԼ��Ҫͨ�����"
         Height          =   210
         Left            =   720
         TabIndex        =   79
         Top             =   2280
         Width           =   3735
      End
      Begin VB.Frame fraCards 
         Caption         =   "���ع���ҽ�ƿ�"
         Height          =   2655
         Left            =   -74775
         TabIndex        =   114
         Top             =   2895
         Width           =   6660
         Begin VSFlex8Ctl.VSFlexGrid vsBill 
            Height          =   2220
            Left            =   195
            TabIndex        =   69
            Top             =   300
            Width           =   6405
            _cx             =   11298
            _cy             =   3916
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
            FormatString    =   $"frmLocalPara.frx":007C
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
      Begin VB.ComboBox cboType 
         Height          =   300
         Left            =   -73548
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   5640
         Width           =   2580
      End
      Begin VB.CheckBox chkMzh 
         Caption         =   "ԤԼʱ�����������"
         Height          =   285
         Left            =   3960
         TabIndex        =   76
         Top             =   1515
         Width           =   2655
      End
      Begin VB.Frame fraClearMZInfor 
         Caption         =   "�˺�����������Ϣ(�Һ���Ч�����ڵĲ���)"
         Height          =   645
         Left            =   -74865
         TabIndex        =   113
         Top             =   5805
         Width           =   4455
         Begin VB.OptionButton optClearInfor 
            Caption         =   "��ʾ���"
            Height          =   180
            Index           =   2
            Left            =   2880
            TabIndex        =   24
            Top             =   300
            Width           =   1110
         End
         Begin VB.OptionButton optClearInfor 
            Caption         =   "�Զ����"
            Height          =   180
            Index           =   1
            Left            =   1410
            TabIndex        =   23
            Top             =   300
            Width           =   1110
         End
         Begin VB.OptionButton optClearInfor 
            Caption         =   "�����"
            Height          =   180
            Index           =   0
            Left            =   210
            TabIndex        =   22
            Top             =   300
            Width           =   1110
         End
      End
      Begin VB.CheckBox chkBespeakFee 
         Caption         =   "�Һŷ�����ԤԼ����ʱ��Ϊ׼!"
         Height          =   210
         Left            =   720
         TabIndex        =   75
         Top             =   1560
         Width           =   3375
      End
      Begin VB.TextBox txtBreakAnAppointmentNums 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   1665
         MaxLength       =   4
         TabIndex        =   85
         Text            =   "0"
         Top             =   4650
         Width           =   660
      End
      Begin VB.Frame Frame5 
         Height          =   120
         Left            =   1425
         TabIndex        =   111
         Top             =   3750
         Width           =   4845
      End
      Begin VB.Frame Frame3 
         Height          =   120
         Left            =   1290
         TabIndex        =   109
         Top             =   540
         Width           =   5130
      End
      Begin VB.CheckBox chkDeptBespeakOneNum 
         Caption         =   "����ͬһ������Լ        ����"
         Height          =   210
         Left            =   720
         TabIndex        =   73
         Top             =   1215
         Width           =   3735
      End
      Begin VB.TextBox txtAvailabilityTimes 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   4035
         MaxLength       =   4
         TabIndex        =   84
         Text            =   "0"
         Top             =   4170
         Width           =   660
      End
      Begin VB.CheckBox chkPrintCase 
         Caption         =   "�Һź��ӡ������ǩ"
         Height          =   255
         Left            =   -74760
         TabIndex        =   33
         ToolTipText     =   "���������ӡ������ǩ"
         Top             =   1550
         Width           =   1980
      End
      Begin VB.CheckBox chkRePrint 
         Caption         =   "�˺Ų��˿�ʱ�ش�Ʊ��"
         Height          =   195
         Left            =   -74625
         TabIndex        =   65
         Top             =   1545
         Width           =   2400
      End
      Begin VB.Frame fraInvoice 
         Caption         =   "�Һ�Ʊ��"
         Height          =   615
         Left            =   -74760
         TabIndex        =   106
         Top             =   2520
         Width           =   6480
         Begin VB.CommandButton cmdPrintSet 
            Caption         =   "��ӡ����"
            Height          =   345
            Index           =   1
            Left            =   5175
            TabIndex        =   48
            Top             =   180
            Width           =   990
         End
         Begin VB.OptionButton optPrintFact 
            Caption         =   "ѡ���Ƿ��ӡ"
            Height          =   180
            Index           =   2
            Left            =   2640
            TabIndex        =   47
            Top             =   285
            Width           =   1380
         End
         Begin VB.OptionButton optPrintFact 
            Caption         =   "����ӡ"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   45
            Top             =   285
            Width           =   900
         End
         Begin VB.OptionButton optPrintFact 
            Caption         =   "�Զ���ӡ"
            Height          =   180
            Index           =   1
            Left            =   1320
            TabIndex        =   46
            Top             =   285
            Value           =   -1  'True
            Width           =   1020
         End
      End
      Begin VB.CheckBox chkNewCardNoPop 
         Caption         =   "����������������Ϣ�ǼǴ���"
         Height          =   195
         Left            =   -74640
         TabIndex        =   62
         Top             =   600
         Width           =   3345
      End
      Begin VB.TextBox txtNameDays 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Left            =   -72870
         MaxLength       =   3
         TabIndex        =   29
         Text            =   "0"
         ToolTipText     =   "0��ʾ����ʱ������ʱ��"
         Top             =   480
         Width           =   285
      End
      Begin VB.Frame fraLine2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   -72840
         TabIndex        =   105
         Top             =   675
         Width           =   285
      End
      Begin VB.CheckBox chkDoctor 
         Caption         =   "��δ����ҽ���ĺű�ʱ��������ҽ��"
         Height          =   195
         Left            =   -74760
         TabIndex        =   30
         ToolTipText     =   "�����ѡ�ű�û��ҽ��,������ѡ��ѱ��������ҵ�ҽ��,��������ѡ��."
         Top             =   750
         Width           =   3180
      End
      Begin VB.CheckBox chkTotal 
         Caption         =   "����ɿ���֮��Ž������ιҺ��շ�"
         Height          =   195
         Left            =   -74760
         TabIndex        =   31
         Top             =   1020
         Width           =   3360
      End
      Begin VB.CheckBox chkCardMoney 
         Caption         =   "������Һŷ�һ����(���򿨷Ѵ�Ϊ���۵�)"
         Height          =   195
         Left            =   -74640
         TabIndex        =   64
         Top             =   1200
         Width           =   3960
      End
      Begin VB.CheckBox chkAutoAddName 
         Caption         =   "�����²����Զ�������ʱ����"
         Height          =   195
         Left            =   -74640
         TabIndex        =   63
         Top             =   900
         Width           =   3345
      End
      Begin VB.CheckBox chkPrintFree 
         Caption         =   "�Һŷ���Ϊ��ʱҲ��ӡƱ��"
         Height          =   255
         Left            =   -74760
         TabIndex        =   32
         ToolTipText     =   "�ű�Ҫ�󽨲�����ʱ�������"
         Top             =   1260
         Width           =   2460
      End
      Begin VB.Frame fraDefaultSet 
         Caption         =   "ȱʡֵ"
         Height          =   1005
         Left            =   -74850
         TabIndex        =   17
         Top             =   4695
         Width           =   4440
         Begin VB.ComboBox cboDefaultSex 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   2895
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   615
            Width           =   1350
         End
         Begin VB.ComboBox cboDefaultPayMode 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   975
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   270
            Width           =   1350
         End
         Begin VB.ComboBox cboDefaultFeeType 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   2895
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   270
            Width           =   1350
         End
         Begin VB.ComboBox cboDefaultBalance 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   975
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   645
            Width           =   1350
         End
         Begin VB.Label lblDefaultSex 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�Ա�"
            Height          =   180
            Left            =   2460
            TabIndex        =   103
            Top             =   675
            Width           =   360
         End
         Begin VB.Label lblDefaultPayMode 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���ʽ"
            Height          =   180
            Left            =   180
            TabIndex        =   102
            Top             =   330
            Width           =   720
         End
         Begin VB.Label lblDefaultBalance 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���㷽ʽ"
            Height          =   180
            Left            =   180
            TabIndex        =   101
            Top             =   705
            Width           =   720
         End
         Begin VB.Label lblDefaultFeeType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ѱ�"
            Height          =   180
            Left            =   2460
            TabIndex        =   100
            Top             =   330
            Width           =   360
         End
      End
      Begin VB.CommandButton cmdDeviceSetup 
         Caption         =   "�豸����(&S)"
         Height          =   330
         Left            =   -69810
         TabIndex        =   44
         Top             =   2190
         Width           =   1425
      End
      Begin VB.ListBox lstDept 
         ForeColor       =   &H80000012&
         Height          =   4680
         Left            =   -70320
         Style           =   1  'Checkbox
         TabIndex        =   25
         ToolTipText     =   "Ctrl+Aȫѡ,Ctrl+Cȫ��,���һ����δѡ���ʾ�����ƿ���"
         Top             =   660
         Width           =   2175
      End
      Begin VB.Frame fraInput 
         Caption         =   "Ҫ��������"
         Height          =   1635
         Left            =   -70800
         TabIndex        =   98
         Top             =   480
         Width           =   2715
         Begin VB.CheckBox chkAllowPhoneInput 
            Caption         =   "��ϵ�绰"
            Height          =   195
            Left            =   1560
            TabIndex        =   43
            Top             =   1200
            Width           =   1020
         End
         Begin VB.CheckBox chkAllowPatientInput 
            Caption         =   "����"
            Height          =   195
            Left            =   240
            TabIndex        =   36
            ToolTipText     =   "�ű�Ҫ�󽨲�����ʱ�������"
            Top             =   315
            Width           =   660
         End
         Begin VB.CheckBox chkAllowSexInput 
            Caption         =   "�Ա�"
            Height          =   195
            Left            =   1560
            TabIndex        =   37
            ToolTipText     =   "�ű�Ҫ�󽨲�����ʱ�������"
            Top             =   315
            Width           =   660
         End
         Begin VB.CheckBox chkAllowAgeInput 
            Caption         =   "����"
            Height          =   195
            Left            =   240
            TabIndex        =   38
            ToolTipText     =   "�ű�Ҫ�󽨲�����ʱ�������"
            Top             =   610
            Width           =   660
         End
         Begin VB.CheckBox chkAllowFeeTypeInput 
            Caption         =   "�ѱ�"
            Height          =   195
            Left            =   1560
            TabIndex        =   39
            Top             =   610
            Width           =   660
         End
         Begin VB.CheckBox chkAllowBalanceInput 
            Caption         =   "���㷽ʽ"
            Height          =   195
            Left            =   1560
            TabIndex        =   41
            Top             =   915
            Width           =   1020
         End
         Begin VB.CheckBox chkAllowPayModeInput 
            Caption         =   "���ʽ"
            Height          =   195
            Left            =   240
            TabIndex        =   40
            Top             =   905
            Width           =   1020
         End
         Begin VB.CheckBox chkAllowAddressInput 
            Caption         =   "��ͥ��ַ"
            Height          =   195
            Left            =   240
            TabIndex        =   42
            Top             =   1200
            Width           =   1020
         End
      End
      Begin VB.Frame fraTitle 
         Caption         =   "���ùҺ�Ʊ��"
         Height          =   1155
         Left            =   -74760
         TabIndex        =   97
         Top             =   5160
         Width           =   6675
         Begin MSComctlLib.ListView lvwBill 
            Height          =   1455
            Left            =   150
            TabIndex        =   61
            Top             =   240
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   2566
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483630
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "������"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "��������"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "���뷶Χ"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "ʣ��"
               Object.Width           =   1499
            EndProperty
         End
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "���ܲ���"
         Height          =   4560
         Left            =   -74880
         TabIndex        =   96
         Top             =   480
         Width           =   4395
         Begin VB.CheckBox chkDeptRegistOneEmer 
            Caption         =   "����ͬ�ƹҺ��������ڼ���"
            Height          =   210
            Left            =   360
            TabIndex        =   14
            Top             =   3420
            Width           =   3735
         End
         Begin VB.TextBox txtDeptRegistOneNum 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   210
            Left            =   1845
            MaxLength       =   2
            TabIndex        =   13
            Text            =   "0"
            Top             =   3135
            Width           =   660
         End
         Begin VB.TextBox txtRegistDeptNums 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   2055
            MaxLength       =   2
            TabIndex        =   16
            Text            =   "0"
            Top             =   3675
            Width           =   660
         End
         Begin VB.CheckBox chkRegistDeptNums 
            Caption         =   "ͬһ��������ܹҺ�        ������"
            Height          =   210
            Left            =   120
            TabIndex        =   15
            Top             =   3690
            Width           =   3705
         End
         Begin VB.CheckBox chkDeptRegistOneNum 
            Caption         =   "����ͬһ�����޹�        ����"
            Height          =   210
            Left            =   120
            TabIndex        =   12
            Top             =   3120
            Width           =   3735
         End
         Begin VB.CheckBox chkNOValidityCheck 
            Caption         =   "�������Ч�Լ��"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   2835
            Width           =   2325
         End
         Begin VB.CheckBox chkDefaultMedBook 
            Caption         =   "�Һ�ʱĬ�Ϲ�ѡ������ѡ��"
            Height          =   300
            Left            =   120
            TabIndex        =   10
            Top             =   2550
            Width           =   2700
         End
         Begin VB.CheckBox chkReuseCanceledNO 
            Caption         =   "�����������Һ�"
            Height          =   300
            Left            =   120
            TabIndex        =   9
            Top             =   2280
            Width           =   2550
         End
         Begin VB.CheckBox chkTimeRangeRegist 
            Caption         =   "��ʱ�κű��ϸ�ʱ�ιҺ�"
            Height          =   300
            Left            =   120
            TabIndex        =   8
            ToolTipText     =   $"frmLocalPara.frx":015D
            Top             =   2010
            Width           =   2550
         End
         Begin VB.CheckBox chkRandSelectNum 
            Caption         =   "������ѡ��"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   1500
            Width           =   2325
         End
         Begin VB.CheckBox chkRigistHeadSort 
            Caption         =   "�ҺŰ��ű�����ͷ����"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   1755
            Width           =   2325
         End
         Begin VB.CheckBox chkAllowZyRigist 
            Caption         =   "����סԺ���˹Һ�"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   1215
            Width           =   1845
         End
         Begin VB.CheckBox chkPrePayPriority 
            Caption         =   "����ʹ��Ԥ����ɷ�"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   975
            Width           =   2340
         End
         Begin VB.CheckBox chkPrice 
            Caption         =   "�������˹ҺŴ�Ϊ���۵�    (��ģʽ�²��ܽ���ҽ���鿨)"
            Height          =   435
            Left            =   120
            TabIndex        =   3
            Top             =   510
            Width           =   2700
         End
         Begin VB.TextBox txtInterval 
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
            Left            =   825
            MaxLength       =   2
            TabIndex        =   1
            Text            =   "5"
            Top             =   30
            Width           =   285
         End
         Begin VB.Frame fraLine 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   15
            Left            =   820
            TabIndex        =   104
            Top             =   225
            Width           =   285
         End
         Begin VB.CheckBox chkAutoRefresh 
            Caption         =   "ÿ��     �����Զ�ˢ�¹ҺŰ��ű�"
            Height          =   195
            Left            =   120
            TabIndex        =   0
            Top             =   30
            Width           =   3480
         End
         Begin VB.CheckBox chkAutoGet 
            Caption         =   "�Һű��뽨����ʱ�Զ��������������"
            Height          =   195
            Left            =   120
            TabIndex        =   2
            Top             =   285
            Width           =   3360
         End
         Begin VB.Line Line1 
            Index           =   9
            X1              =   1815
            X2              =   2535
            Y1              =   3345
            Y2              =   3345
         End
         Begin VB.Line Line1 
            Index           =   8
            X1              =   2025
            X2              =   2745
            Y1              =   3915
            Y2              =   3915
         End
      End
      Begin VB.CheckBox chkSeekName 
         Caption         =   "����������ģ������    ���ڵĲ���"
         Height          =   195
         Left            =   -74760
         TabIndex        =   28
         Top             =   480
         Width           =   3300
      End
      Begin VB.CheckBox chkDeptNums 
         Caption         =   "ͬһ���������ԤԼ        ������"
         Height          =   210
         Left            =   720
         TabIndex        =   71
         Top             =   900
         Width           =   3705
      End
      Begin VB.Label lblColor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�°���ǰ�ҺŰ�����ɫ"
         Height          =   180
         Left            =   -70320
         TabIndex        =   27
         Top             =   5850
         Width           =   1800
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ȱʡԤԼ��ʽ"
         Height          =   180
         Left            =   720
         TabIndex        =   125
         Top             =   3465
         Width           =   1080
      End
      Begin VB.Label lblSortMode 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ʽ"
         Height          =   180
         Left            =   -70320
         TabIndex        =   124
         Top             =   5505
         Width           =   720
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   2430
         X2              =   3150
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   2640
         X2              =   3360
         Y1              =   1125
         Y2              =   1125
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   -74760
         X2              =   -74040
         Y1              =   2340
         Y2              =   2340
      End
      Begin VB.Label lblGuardian 
         AutoSize        =   -1  'True
         Caption         =   "�����±���¼��໤��"
         Height          =   180
         Left            =   -74040
         TabIndex        =   121
         Top             =   2100
         Width           =   1800
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   1800
         X2              =   2520
         Y1              =   3300
         Y2              =   3300
      End
      Begin VB.Label lblBespeakDefaultDays 
         AutoSize        =   -1  'True
         Caption         =   "ԤԼȱʡ����        ��"
         Height          =   180
         Left            =   720
         TabIndex        =   118
         Top             =   3045
         Width           =   1980
      End
      Begin VB.Label lblBespeakMinTime 
         AutoSize        =   -1  'True
         Caption         =   "ԤԼ����ʱ��        ���ӣ�ָԤԼʱ���������ʱ�̵���С���"
         Height          =   180
         Left            =   705
         TabIndex        =   117
         Top             =   2655
         Width           =   5220
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   1800
         X2              =   2520
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   1320
         X2              =   2040
         Y1              =   2115
         Y2              =   2115
      End
      Begin VB.Label lblCancelBespeak 
         AutoSize        =   -1  'True
         Caption         =   "ԤԼ��         ���ڲ���ȡ��ԤԼ"
         Height          =   180
         Left            =   735
         TabIndex        =   116
         Top             =   1890
         Width           =   2790
      End
      Begin VB.Label lblDefaultPayCard 
         Caption         =   "ȱʡ��������"
         Height          =   210
         Left            =   -74700
         TabIndex        =   115
         Top             =   5700
         Width           =   1290
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   4080
         X2              =   4800
         Y1              =   4410
         Y2              =   4410
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   1695
         X2              =   2415
         Y1              =   4890
         Y2              =   4890
      End
      Begin VB.Label lblBreakAnAppointmentNums 
         AutoSize        =   -1  'True
         Caption         =   "����ԤԼʧԼ         ���Զ����������"
         Height          =   180
         Left            =   540
         TabIndex        =   112
         Top             =   4650
         Width           =   3330
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "����������"
         Height          =   180
         Left            =   540
         TabIndex        =   110
         Top             =   3750
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "ԤԼ����"
         Height          =   180
         Left            =   525
         TabIndex        =   108
         Top             =   540
         Width           =   720
      End
      Begin VB.Label lblAvailabilityTimes 
         AutoSize        =   -1  'True
         Caption         =   "ԤԼ��Чʱ�䣺ԤԼ����ԤԼʱ��                 ����δ���յ�ΪʧԼ��"
         Height          =   180
         Left            =   540
         TabIndex        =   107
         Top             =   4170
         Width           =   6030
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Һſ���"
         Height          =   180
         Left            =   -70335
         TabIndex        =   99
         ToolTipText     =   "�趨�����ɹ���Щ���ҵĺ�"
         Top             =   450
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   330
      Left            =   120
      TabIndex        =   92
      Top             =   6825
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   330
      Left            =   6060
      TabIndex        =   94
      Top             =   6825
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   330
      Left            =   4680
      TabIndex        =   93
      Top             =   6825
      Width           =   1100
   End
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmLocalPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Public mstrPrivs As String
Public mlngModul As Long
Private mstrColor As String

'Private mblnOK As Boolean
Private Sub cboType_Click()
  chkScanIDVisa.Enabled = Not (cboType.Text = "�������֤")
  If cboType.Text = "�������֤" Then
        chkScanIDVisa.Value = 1
  End If
End Sub
Private Sub chkAutoAddName_Click()
    If chkAutoAddName.Value = 1 Then
        chkAllowPatientInput.Value = 0
        chkAllowSexInput.Value = 0
        chkAllowAgeInput.Value = 0
        chkAllowAddressInput.Value = 0
        chkAllowPayModeInput.Value = 0
        chkNewCardNoPop.Value = 0
    End If
End Sub

Private Sub chkAutoRefresh_Click()
    txtInterval.Enabled = chkAutoRefresh.Value = 1
    If txtInterval.Enabled And txtInterval.Visible Then
        txtInterval.SetFocus
    End If
End Sub

Private Sub chkCardMoney_Click()
    If chkCardMoney.Value = 0 Then
        chkNewCardNoPop.Value = 0
        
        chkRePrint.Value = 0
        chkRePrint.Enabled = False
    Else
        chkRePrint.Enabled = True
    End If
End Sub

Private Sub chkDeptBespeakOneNum_Click()
    If chkDeptBespeakOneNum.Value = 1 Then
        txtDeptBespeakOneNum.Enabled = True
    Else
        txtDeptBespeakOneNum.Enabled = False
    End If
End Sub

Private Sub chkDeptNums_Click()
    If chkDeptNums.Value = 1 Then
        txtDeptNums.Enabled = True
    Else
        txtDeptNums.Enabled = False
    End If
End Sub

Private Sub chkDeptRegistOneNum_Click()
    If chkDeptRegistOneNum.Value = 1 Then
        txtDeptRegistOneNum.Enabled = True
        chkDeptRegistOneEmer.Enabled = True
    Else
        txtDeptRegistOneNum.Enabled = False
        chkDeptRegistOneEmer.Enabled = False
    End If
End Sub

Private Sub chkNewCardNoPop_Click()
    If chkNewCardNoPop.Value = 1 Then
        chkAutoAddName.Value = 0
        chkCardMoney.Value = 1  '��������ʱ,���Ѳ����ȴ�Ϊ���۵�,��Ϊ��ʱδ���������ܽ���
    End If
End Sub


Private Sub chkRegistDeptNums_Click()
    If chkRegistDeptNums.Value = 1 Then
        txtRegistDeptNums.Enabled = True
    Else
        txtRegistDeptNums.Enabled = False
    End If
End Sub

Private Sub chkSeekName_Click()
    txtNameDays.Enabled = chkSeekName.Value = 1 And txtNameDays.Tag = "1"
End Sub

Private Sub chkAllowPatientInput_Click()
    If chkAllowPatientInput.Value = 0 Then
        chkAllowSexInput.Value = 0
        chkAllowAgeInput.Value = 0
        chkAllowAddressInput.Value = 0
        chkAllowPayModeInput.Value = 0
        chkAllowSexInput.Enabled = False
        chkAllowAgeInput.Enabled = False
        chkAllowAddressInput.Enabled = False
        chkAllowPayModeInput.Enabled = False
    Else
        chkAllowSexInput.Enabled = True And chkAllowSexInput.Tag = "1"
        chkAllowAgeInput.Enabled = True And chkAllowAgeInput.Tag = "1"
        chkAllowAddressInput.Enabled = True And chkAllowAddressInput.Tag = "1"
        chkAllowPayModeInput.Enabled = True And chkAllowPayModeInput.Tag = "1"
    End If
End Sub

 

Private Sub chkDeptBespeakOneNum_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, 100, 1111)
End Sub

Private Sub cmdHelp_Click()
ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub



Private Sub cmdOK_Click()
    Dim i As Integer
    Dim strTmp As String
    Dim blnHavePrivs As Boolean
    
    On Error GoTo Hd
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    '���ݿ�洢��ģ�����
    '-------------------------------------------------------------------------------------------
    zlDatabase.SetPara "�Զ�ˢ�¼��", IIf(chkAutoRefresh.Value = 1, Val(txtInterval.Text), 0), glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "�Զ������", chkAutoGet.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "��Ϊ���۵�", chkPrice.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "����ʹ��Ԥ����", chkPrePayPriority.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "ȱʡ���ʽ", cboDefaultPayMode.Text, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "ȱʡ�ѱ�", cboDefaultFeeType.Text, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "ȱʡ���㷽ʽ", cboDefaultBalance.Text, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "ȱʡ�Ա�", cboDefaultSex.Text, glngSys, mlngModul, blnHavePrivs
    '�����:53408
    zlDatabase.SetPara "ɨ�����֤ǩԼ", chkScanIDVisa.Value, glngSys, mlngModul, blnHavePrivs
    '69506
    zlDatabase.SetPara "�����������Һ�", chkReuseCanceledNO.Value, glngSys, mlngModul, blnHavePrivs
    '53045
    zlDatabase.SetPara "Ĭ�Ϲ�����", chkDefaultMedBook.Value, glngSys, mlngModul, blnHavePrivs
    '����:35176
    zlDatabase.SetPara "�˺����������Ϣ", IIf(optClearInfor(0).Value, 0, IIf(optClearInfor(1).Value, 1, 2)), glngSys, mlngModul, blnHavePrivs
    '����:31182
    If chkDeptBespeakOneNum.Value = 1 Then
        zlDatabase.SetPara "����ͬ����ԼN����", Val(txtDeptBespeakOneNum.Text), glngSys, mlngModul, blnHavePrivs
    Else
        zlDatabase.SetPara "����ͬ����ԼN����", 0, glngSys, mlngModul, blnHavePrivs
    End If
    
    If chkDeptRegistOneNum.Value = 1 Then
        zlDatabase.SetPara "����ͬ���޹�N����", Val(txtDeptRegistOneNum.Text) & "|" & IIf(chkDeptRegistOneEmer.Value = 1, 1, 0), glngSys, mlngModul, blnHavePrivs
    Else
        zlDatabase.SetPara "����ͬ���޹�N����", 0, glngSys, mlngModul, blnHavePrivs
    End If
    
    If chkDeptNums.Value = 1 Then
        zlDatabase.SetPara "����ԤԼ������", Val(txtDeptNums.Text), glngSys, mlngModul, blnHavePrivs
    Else
        zlDatabase.SetPara "����ԤԼ������", 0, glngSys, mlngModul, blnHavePrivs
    End If
    
    If chkRegistDeptNums.Value = 1 Then
        zlDatabase.SetPara "���˹Һſ�������", Val(txtRegistDeptNums.Text), glngSys, mlngModul, blnHavePrivs
    Else
        zlDatabase.SetPara "���˹Һſ�������", 0, glngSys, mlngModul, blnHavePrivs
    End If
    
    zlDatabase.SetPara "ԤԼ��Чʱ��", IIf(cboԤԼ��Чʱ��.ListIndex = 0, 1, -1) * Val(txtAvailabilityTimes.Text), glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "ԤԼ����ʱ��", Val(txtBespeakDefaultDays) & "|" & Val(txtBespeakMinTime.Text), glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "ԤԼʧԼ����", Val(txtBreakAnAppointmentNums.Text), glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "ԤԼ����ȷ���Һŷ�", chkBespeakFee.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "ʧԼ���ڹҺ�", chkBreakAnAppointmentToRegist.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "ԤԼ�����������", chkMzh.Value, glngSys, mlngModul, blnHavePrivs '36028
    '71651:������,2014-03-31,�������� �������Ч�Լ��
    zlDatabase.SetPara "�������Ч�Լ��", chkNOValidityCheck.Value, glngSys, mlngModul, blnHavePrivs
     '���� 43847
    zlDatabase.SetPara "������ͷ����", chkRigistHeadSort.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "ȱʡ����ʽ", IIf(cboSortMode.ListIndex = -1, 0, cboSortMode.ListIndex), glngSys, mlngModul, blnHavePrivs
    
    zlDatabase.SetPara "N���ڲ���ȡ��ԤԼ��", Val(txtCancelBespeak.Text), glngSys, mlngModul, blnHavePrivs
    
    zlDatabase.SetPara "N�����±���¼��໤��", Val(txtMustGuardianInfo.Text), glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "�˺����", chkBackNoToVerfy.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "������ѡ��", chkRandSelectNum.Value, glngSys, mlngModul, blnHavePrivs
     '62467
    zlDatabase.SetPara "�ϸ�ʱ�ιҺ�", chkTimeRangeRegist.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "ȱʡԤԼ��ʽ", NeedName(cboDefaultStyle.Text), glngSys, mlngModul, blnHavePrivs
    strTmp = ""
    If lstDept.ListCount <> lstDept.SelCount Then
        For i = 0 To lstDept.ListCount - 1
            If lstDept.Selected(i) = True Then
                strTmp = strTmp & "," & lstDept.ItemData(i)
            End If
        Next
        If strTmp <> "" Then strTmp = Mid(strTmp, 2)
    End If
    zlDatabase.SetPara "�Һſ���", strTmp, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "����ģ������", chkSeekName.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "������������", Val(txtNameDays.Text), glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "����ҽ��", chkDoctor.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "�ɿ�ҺŽ���", chkTotal.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "ԤԼ����ģʽ", IIf(optReceiveMode(0).Value, 0, 1), glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "����ô�ӡ", chkPrintFree.Value, glngSys, mlngModul, blnHavePrivs
    For i = 0 To optPrintFact.UBound
        If optPrintFact(i).Value Then
            zlDatabase.SetPara "�Һŷ�Ʊ��ӡ��ʽ", i, glngSys, mlngModul, blnHavePrivs
            Exit For
        End If
    Next
    For i = 0 To Me.optPrepayPrint.UBound
        If optPrepayPrint(i).Value Then
            zlDatabase.SetPara "���������ӡ��ʽ", i, glngSys, mlngModul, blnHavePrivs
            Exit For
        End If
    Next
    '56274
    For i = 0 To Me.optPrintBespeak.UBound
        If optPrintBespeak(i).Value Then
            zlDatabase.SetPara "ԤԼ�Һŵ���ӡ��ʽ", i, glngSys, mlngModul, blnHavePrivs
            Exit For
        End If
    Next
    '68408
    For i = 0 To Me.optSlipPrint.UBound
        If optSlipPrint(i).Value Then
            zlDatabase.SetPara "�Һ�ƾ����ӡ��ʽ", i, glngSys, mlngModul, blnHavePrivs
            Exit For
        End If
    Next
    
    For i = 0 To optReceipt.UBound
        If optReceipt(i).Value Then
            zlDatabase.SetPara "�˺Żص���ӡ��ʽ", i, glngSys, mlngModul, blnHavePrivs
            Exit For
        End If
    Next
    
    zlDatabase.SetPara "��ӡ������ǩ", chkPrintCase.Value, glngSys, mlngModul, blnHavePrivs
    
    '���ùҺ�Ʊ������
    strTmp = "0"
    For i = 1 To lvwBill.ListItems.Count
        If lvwBill.ListItems(i).Checked Then strTmp = Mid(lvwBill.ListItems(i).Key, 2)
    Next
    zlDatabase.SetPara "���ùҺ�Ʊ������", strTmp, glngSys, mlngModul, blnHavePrivs
    
    zlDatabase.SetPara "��������", chkAllowPatientInput.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "�����Ա�", chkAllowSexInput.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "��������", chkAllowAgeInput.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "�����ͥ��ַ", chkAllowAddressInput.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "���븶�ʽ", chkAllowPayModeInput.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "����ѱ�", chkAllowFeeTypeInput.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "������㷽ʽ", chkAllowBalanceInput.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "������ϵ�绰", chkAllowPhoneInput.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "����סԺ���˹Һ�", chkAllowZyRigist.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "������������", chkNewCardNoPop.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "�Զ���������", chkAutoAddName.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "��ȡ����", chkCardMoney.Value, glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "�˷��ش�", IIf(chkRePrint.Enabled, chkRePrint.Value, 0), glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "��ͥ��ַ���뷽ʽ", IIf(chkAddressAssnInput.Enabled, chkAddressAssnInput.Value, 1), glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "���ϸ����ʱʼ�շ���", IIf(chkAlwaysSendCard.Enabled, chkAlwaysSendCard.Value, 0), glngSys, mlngModul, blnHavePrivs
    '92468:���ϴ�,2016/1/25,�ϸ�����¿���Ϊ0Ҳ��Ʊ��
    zlDatabase.SetPara "�㿨����Ʊ��", IIf(ChkMustBill.Enabled, ChkMustBill.Value, 0), glngSys, mlngModul, blnHavePrivs
    zlDatabase.SetPara "��ǰ�Һ���ɫ", mstrColor, glngSys, mlngModul, blnHavePrivs
    
    Call SaveInvoice
    Call InitLocPar(mlngModul)
    gblnOk = True
    Unload Me
    Exit Sub
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Sub

Private Function LoadFactList(bytKind As Byte) As Boolean
'���ܣ���ȡ���ù��ùҺ�Ʊ�ݻ���￨����
'����:bytKind=4-�Һ�Ʊ��,5-���￨
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer, lngTmp As Long
    Dim ObjItem As ListItem
    Dim blnBill As Boolean
    
    On Error GoTo errH
    lngTmp = zlDatabase.GetPara("���ùҺ�Ʊ������", glngSys, mlngModul, 0, Array(lvwBill), InStr(mstrPrivs, "��������") > 0)
    Set rsTmp = GetShareInvoiceGroupID(bytKind)
    
    For i = 1 To rsTmp.RecordCount
        Set ObjItem = lvwBill.ListItems.Add(, "_" & rsTmp!ID, rsTmp!������)
        ObjItem.SubItems(1) = Format(rsTmp!�Ǽ�ʱ��, "yyyy-MM-dd")
        ObjItem.SubItems(2) = rsTmp!��ʼ���� & "," & rsTmp!��ֹ����
        ObjItem.SubItems(3) = rsTmp!ʣ������
        If rsTmp!ID = lngTmp Then
            ObjItem.Checked = True
            ObjItem.Selected = True
            blnBill = True
        End If
        rsTmp.MoveNext
    Next
    
    If Not blnBill Then
        zlDatabase.SetPara IIf(bytKind = 4, "���ùҺ�Ʊ������", "���þ��￨����"), "0", glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
    End If
    
    LoadFactList = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdPrintSet_Click(Index As Integer)
    On Error GoTo Hd
    Select Case Index
    '���������ӡ
    Case 0:
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111_2", Me)
    Case 1:
        '�Һ��շѴ�ӡ
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111", Me)
    Case 2:
        'ԤԼ�ҺŴ�ӡ   '56274
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111_1", Me)
    Case 3:
        '68408,������,2013-12-11,�Һ�ƾ����ӡ
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1111_3", Me)
    Case 4:
        '�˺Żص���ӡ
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111_3", Me)
    Case Else:
    End Select
    Exit Sub
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 Then
        Dim i As Integer
        If UCase(Chr(KeyCode)) = "A" Then
            For i = 0 To lstDept.ListCount - 1
                lstDept.Selected(i) = True
            Next
        ElseIf UCase(Chr(KeyCode)) = "C" Then
            For i = 0 To lstDept.ListCount - 1
                lstDept.Selected(i) = False
            Next
        End If
    End If
End Sub

Private Sub Load֧����ʽ()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ч��֧����ʽ
    '����:Ƚ����
    '����:2014-07-02
    '�����:74552
    '˵��:�ҺŹ���������Ĭ�Ͻ��㷽ʽʱ�����ѡ����㷽ʽ����Ϊ"7-һ��ͨ����"�Ľ��㷽ʽ
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long, strPayType As String, varData As Variant, varTemp As Variant, i As Long
    Dim rsTemp As ADODB.Recordset, blnFind As Boolean
    Dim strSQL As String

    strSQL = _
        " Select B.����,B.����,Nvl(B.ȱʡ��־,0) as ȱʡ,Nvl(B.����,1) as ����,Nvl(B.Ӧ����,0) as Ӧ����" & _
        " From ���㷽ʽӦ�� A,���㷽ʽ B" & _
        " Where A.Ӧ�ó���=[1] And B.����=A.���㷽ʽ" & _
        "   And(B.����<>7 Or B.����=7 And Exists(Select 1 From һ��ͨĿ¼ C Where C.���㷽ʽ=B.���� And C.����=1))" & _
        "   and B.����<>8 And Instr(',1,2,7,',','||B.����||',')>0" & _
        " Order by ����,lpad(����,3,' ')"
    Err = 0: On Error GoTo Errhand
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "�Һ�")
    
    '��ȡ�������Ľ��㷽ʽ
    '��|ȫ��|ˢ����־|�����ID(���ѿ����)|����|�Ƿ����ѿ�|���㷽ʽ|�Ƿ�����|�Ƿ����ƿ�;��
    If Not gobjSquare.objSquareCard Is Nothing Then
        strPayType = gobjSquare.objSquareCard.zlGetAvailabilityCardType
    End If
    
    varData = Split(strPayType, ";")
    With cboDefaultBalance
        .Clear
        Do While Not rsTemp.EOF
            blnFind = False
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "|||||", "|")
                If varTemp(6) = Nvl(rsTemp!����) Then
                    blnFind = True: Exit For
                End If
            Next
                         
            If Not blnFind Then
                .AddItem Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����)
                .ItemData(.NewIndex) = 1
                If Nvl(rsTemp!����) & "-" & Nvl(rsTemp!����) = gstr���㷽ʽ Then
                     .ItemData(.NewIndex) = 1
                     .ListIndex = .NewIndex
                End If
                If Val(Nvl(rsTemp!ȱʡ)) = 1 Then .ListIndex = .NewIndex
            End If
            rsTemp.MoveNext
        Loop
        
        '���ؽ��㷽ʽ����Ϊ��7-һ��ͨ���㡱��ҽ�ƿ����
        For i = 0 To UBound(varData)
            If InStr(1, varData(i), "|") <> 0 Then
                varTemp = Split(varData(i), "|")
                .AddItem varTemp(1): .ItemData(.NewIndex) = -1
            End If
        Next
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
    End With
    
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim rsTmp           As New ADODB.Recordset
    Dim strSQL          As String
    Dim i               As Integer
    Dim str����ID       As String
    Dim strTmp          As String
    Dim blnParSet       As Boolean
    Dim intIndex        As Integer
    Dim lngValue        As Long
    
    gblnOk = False
    
    blnParSet = InStr(mstrPrivs, "��������") > 0
    On Error GoTo errH
    'a.��ʼ����
    '----------------------------------------------------------------------------------------
    strSQL = "Select Distinct B.���� ||'-'|| B.���� as ����,B.ID From �ҺŰ��� A,���ű� B Where A.����ID=B.ID Order by ����"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    zlControl.CboAddData lstDept, rsTmp, True
    
    strSQL = "Select 'ҽ�Ƹ��ʽ' ����,����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From ҽ�Ƹ��ʽ" & _
            " Union All " & _
            " Select '�Ա�' ����,����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From �Ա�" & _
            " Union All " & _
            " Select '�ѱ�' ����,����,����,����,Nvl(ȱʡ��־,0) as ȱʡ From �ѱ�" & _
            " Where ����=1 And Nvl(���޳���,0)=0 And Nvl(�������,3) IN(1,3)" & _
            " Order by ����,����"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    'ȱʡҽ�Ƹ��ʽ
    rsTmp.Filter = "����='ҽ�Ƹ��ʽ'"
    For i = 1 To rsTmp.RecordCount
        cboDefaultPayMode.AddItem rsTmp!����
        If rsTmp!ȱʡ = 1 Then cboDefaultPayMode.ListIndex = cboDefaultPayMode.NewIndex
        rsTmp.MoveNext
    Next
     'ȱʡ�ѱ�    '���ǽ��޳������Ψһ����Ŀ(������ȱʡ�ѱ�),������Ч�ڼ估����
    rsTmp.Filter = "����='�ѱ�'"
    For i = 1 To rsTmp.RecordCount
        cboDefaultFeeType.AddItem rsTmp!����
        If rsTmp!ȱʡ = 1 Then cboDefaultFeeType.ListIndex = cboDefaultFeeType.NewIndex
        rsTmp.MoveNext
    Next
    'ȱʡ�Ա�
    rsTmp.Filter = "����='�Ա�'"
    For i = 1 To rsTmp.RecordCount
        cboDefaultSex.AddItem rsTmp!����
        If rsTmp!ȱʡ = 1 Then cboDefaultSex.ListIndex = cboDefaultSex.NewIndex
        rsTmp.MoveNext
    Next
    cboDefaultSex.AddItem "��"
    'ȱʡ���㷽ʽ
    Call Load֧����ʽ

    cboSortMode.Clear
    cboSortMode.AddItem "0.�ű�"
    cboSortMode.ItemData(cboSortMode.NewIndex) = 0
    cboSortMode.ListIndex = 0
    cboSortMode.AddItem "1.����-��Ŀ"
    cboSortMode.ItemData(cboSortMode.NewIndex) = 1
    cboSortMode.AddItem "2.����"
    cboSortMode.ItemData(cboSortMode.NewIndex) = 2
    
    With cboԤԼ��Чʱ��
        .Clear
        .AddItem "��ǰ"
        .AddItem "�Ӻ�"
    End With
                
    
    'c.���ݿ�洢��ģ�����
    '----------------------------------------------------------------------------------------
    strTmp = zlDatabase.GetPara("�Զ�ˢ�¼��", glngSys, mlngModul, , Array(chkAutoRefresh, txtInterval), blnParSet)
    chkAutoRefresh.Value = IIf(Val(strTmp) > 0, 1, 0)
    If chkAutoRefresh.Value = 1 Then txtInterval.Text = strTmp
    chkAutoGet.Value = IIf(zlDatabase.GetPara("�Զ������", glngSys, mlngModul, , Array(chkAutoGet), blnParSet) = "1", 1, 0)
    chkPrice.Value = IIf(zlDatabase.GetPara("��Ϊ���۵�", glngSys, mlngModul, , Array(chkPrice), blnParSet) = "1", 1, 0)
    chkPrePayPriority.Value = IIf(zlDatabase.GetPara("����ʹ��Ԥ����", glngSys, mlngModul, , Array(chkPrePayPriority), blnParSet) = "1", 1, 0)
    chkRigistHeadSort.Value = IIf(zlDatabase.GetPara("������ͷ����", glngSys, mlngModul, , Array(chkRigistHeadSort), blnParSet) = "1", 1, 0)
    chkRandSelectNum.Value = IIf(zlDatabase.GetPara("������ѡ��", glngSys, mlngModul, , Array(chkRandSelectNum), blnParSet) = "1", 1, 0)
    chkBreakAnAppointmentToRegist.Value = IIf(zlDatabase.GetPara("ʧԼ���ڹҺ�", glngSys, mlngModul, 0, Array(chkBreakAnAppointmentToRegist), blnParSet) = "1", 1, 0)
    chkBackNoToVerfy.Value = IIf(zlDatabase.GetPara("�˺����", glngSys, mlngModul, 0, Array(chkBackNoToVerfy), blnParSet) = "1", 1, 0)
    chkAddressAssnInput.Value = IIf(zlDatabase.GetPara("��ͥ��ַ���뷽ʽ", glngSys, mlngModul, 1, Array(chkAddressAssnInput), blnParSet) = "1", 1, 0)
    chkAlwaysSendCard.Value = IIf(zlDatabase.GetPara("���ϸ����ʱʼ�շ���", glngSys, mlngModul, 1, Array(chkAlwaysSendCard), blnParSet) = "1", 1, 0)
    '69506��������,2014-01-14,��������"�����������Һ�"
    chkReuseCanceledNO.Value = IIf(zlDatabase.GetPara("�����������Һ�", glngSys, mlngModul, 1, Array(chkReuseCanceledNO), blnParSet) = "1", 1, 0)
    chkNOValidityCheck.Value = IIf(zlDatabase.GetPara("�������Ч�Լ��", glngSys, mlngModul, 1, Array(chkNOValidityCheck), blnParSet) = "1", 1, 0)
    '53045��������,2014-02-13,Ĭ�Ϲ�ѡ������ѡ��
    chkDefaultMedBook.Value = IIf(zlDatabase.GetPara("Ĭ�Ϲ�����", glngSys, mlngModul, 0, Array(chkDefaultMedBook), blnParSet) = "1", 1, 0)
    strTmp = zlDatabase.GetPara("ȱʡ����ʽ", glngSys, glngModul, , Array(cboSortMode), blnParSet)
    cboSortMode.ListIndex = Val(strTmp)
    strTmp = zlDatabase.GetPara("ȱʡ���ʽ", glngSys, mlngModul, , Array(cboDefaultPayMode), blnParSet)
    zlControl.CboLocate cboDefaultPayMode, strTmp
    strTmp = zlDatabase.GetPara("ȱʡ�ѱ�", glngSys, mlngModul, , Array(cboDefaultFeeType), blnParSet)
    zlControl.CboLocate cboDefaultFeeType, strTmp
    strTmp = zlDatabase.GetPara("ȱʡ�Ա�", glngSys, mlngModul, , Array(cboDefaultSex), blnParSet)
    zlControl.CboLocate cboDefaultSex, strTmp
    If cboDefaultSex.ListIndex = -1 Or strTmp = "��" Then cboDefaultSex.ListIndex = cboDefaultSex.ListCount - 1
    strTmp = zlDatabase.GetPara("ȱʡ���㷽ʽ", glngSys, mlngModul, , Array(cboDefaultBalance), blnParSet)
    zlControl.CboLocate cboDefaultBalance, strTmp
    '�����:53408
    chkScanIDVisa.Value = IIf(zlDatabase.GetPara("ɨ�����֤ǩԼ", glngSys, mlngModul, 0, Array(chkScanIDVisa), blnParSet) = "1", 1, 0)
    '����:35176
    strTmp = zlDatabase.GetPara("�˺����������Ϣ", glngSys, mlngModul, , Array(fraClearMZInfor, optClearInfor(0), optClearInfor(1), optClearInfor(2)), blnParSet)
    If Val(strTmp) = 0 Then
        optClearInfor(0).Value = True
    ElseIf Val(strTmp) = 1 Then
        optClearInfor(1).Value = True
    Else
        optClearInfor(2).Value = True
    End If
    
    strTmp = zlDatabase.GetPara("ԤԼ����ģʽ", glngSys, mlngModul, , Array(fraReceiveMode, optReceiveMode(0), optReceiveMode(1)), blnParSet)
    If Val(strTmp) = 0 Then
        optReceiveMode(0).Value = True
    Else
        optReceiveMode(1).Value = True
    End If
    
    mstrColor = zlDatabase.GetPara("��ǰ�Һ���ɫ", glngSys, 1111, "", , blnParSet)
    If mstrColor = "" Then mstrColor = &H0&
    pic��ǰ��ɫ.BackColor = mstrColor
    
    strTmp = zlDatabase.GetPara("ȱʡԤԼ��ʽ", glngSys, 1115, "", Array(cboDefaultStyle), True)
    strSQL = "Select ����,����,ȱʡ��־ From ԤԼ��ʽ Order By ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    cboDefaultStyle.Clear
    Do While Not rsTmp.EOF
        cboDefaultStyle.AddItem rsTmp!���� & "-" & rsTmp!����
        If strTmp = Nvl(rsTmp!����) Then intIndex = cboDefaultStyle.NewIndex
        If Val(Nvl(rsTmp!ȱʡ��־)) = 1 Then cboDefaultStyle.ListIndex = cboDefaultStyle.NewIndex
        rsTmp.MoveNext
    Loop
    If cboDefaultStyle.ListCount <> 0 And intIndex <> 0 Then cboDefaultStyle.ListIndex = intIndex
    
    '62467
    chkTimeRangeRegist.Value = IIf(zlDatabase.GetPara("�ϸ�ʱ�ιҺ�", glngSys, mlngModul, 0, Array(chkTimeRangeRegist), blnParSet) = "1", 1, 0)
    
    '����:31182
    txtDeptBespeakOneNum.Text = Val(zlDatabase.GetPara("����ͬ����ԼN����", glngSys, mlngModul, 0, Array(chkDeptBespeakOneNum, txtDeptBespeakOneNum), blnParSet))
    If Val(txtDeptBespeakOneNum.Text) = 0 Then
        chkDeptBespeakOneNum.Value = 0
        txtDeptBespeakOneNum.Text = ""
        txtDeptBespeakOneNum.Enabled = False
    Else
        chkDeptBespeakOneNum.Value = 1
        txtDeptBespeakOneNum.Enabled = True
    End If
    
    txtDeptRegistOneNum.Text = Val(Split(zlDatabase.GetPara("����ͬ���޹�N����", glngSys, mlngModul, 0, Array(chkDeptRegistOneNum, txtDeptRegistOneNum), blnParSet) & "|", "|")(0))
    If Val(txtDeptRegistOneNum.Text) = 0 Then
        chkDeptRegistOneNum.Value = 0
        chkDeptRegistOneEmer.Value = 0
        chkDeptRegistOneEmer.Enabled = False
        txtDeptRegistOneNum.Text = ""
        txtDeptRegistOneNum.Enabled = False
    Else
        chkDeptRegistOneNum.Value = 1
        chkDeptRegistOneEmer.Value = Val(Split(zlDatabase.GetPara("����ͬ���޹�N����", glngSys, mlngModul, 0, Array(chkDeptRegistOneNum, txtDeptRegistOneNum), blnParSet) & "|", "|")(1))
        chkDeptRegistOneEmer.Enabled = True
        txtDeptRegistOneNum.Enabled = True
    End If
    
    txtDeptNums.Text = Val(zlDatabase.GetPara("����ԤԼ������", glngSys, mlngModul, 0, Array(txtDeptNums, chkDeptNums), blnParSet))
    If Val(txtDeptNums.Text) = 0 Then
        chkDeptNums.Value = 0
        txtDeptNums.Text = ""
        txtDeptNums.Enabled = False
    Else
        chkDeptNums.Value = 1
        txtDeptNums.Enabled = True
    End If
    txtRegistDeptNums.Text = Val(zlDatabase.GetPara("���˹Һſ�������", glngSys, mlngModul, 0, Array(chkRegistDeptNums, txtRegistDeptNums), blnParSet))
    If Val(txtRegistDeptNums.Text) = 0 Then
        chkRegistDeptNums.Value = 0
        txtRegistDeptNums.Text = ""
        txtRegistDeptNums.Enabled = False
    Else
        chkRegistDeptNums.Value = 1
        txtRegistDeptNums.Enabled = True
    End If
    
    lngValue = Val(zlDatabase.GetPara("ԤԼ��Чʱ��", glngSys, mlngModul, 0, Array(txtAvailabilityTimes, lblAvailabilityTimes), blnParSet))
    
    If lngValue >= 0 Then
        cboԤԼ��Чʱ��.ListIndex = 0
    Else
        cboԤԼ��Чʱ��.ListIndex = 1
    End If
    txtAvailabilityTimes.Text = Abs(lngValue)
    
    txtBreakAnAppointmentNums.Text = Val(zlDatabase.GetPara("ԤԼʧԼ����", glngSys, mlngModul, 0, Array(txtBreakAnAppointmentNums, lblBreakAnAppointmentNums), blnParSet))
    chkBespeakFee.Value = IIf(zlDatabase.GetPara("ԤԼ����ȷ���Һŷ�", glngSys, mlngModul, 0, Array(chkBespeakFee), blnParSet) = "1", 1, 0)
    chkBreakAnAppointmentToRegist.Value = IIf(zlDatabase.GetPara("ʧԼ���ڹҺ�", glngSys, mlngModul, 0, Array(chkBreakAnAppointmentToRegist), blnParSet) = "1", 1, 0)
    txtCancelBespeak.Text = Val(zlDatabase.GetPara("N���ڲ���ȡ��ԤԼ��", glngSys, mlngModul, 0, Array(txtCancelBespeak, lblCancelBespeak), blnParSet))
    Call txtCancelBespeak_Change
    txtMustGuardianInfo.Text = Val(zlDatabase.GetPara("N�����±���¼��໤��", glngSys, mlngModul, 0, Array(txtMustGuardianInfo, lblGuardian), blnParSet))
    
    strTmp = zlDatabase.GetPara("ԤԼ����ʱ��", glngSys, mlngModul, "1|60", Array(txtBespeakMinTime, lblBespeakMinTime, lblBespeakDefaultDays, txtBespeakDefaultDays), blnParSet)
    txtBespeakMinTime.Text = Val(Split(strTmp, "|")(1))
    txtBespeakDefaultDays.Text = Val(Split(strTmp, "|")(0))
    '��ȡ���õĹҺſ���
    str����ID = zlDatabase.GetPara("�Һſ���", glngSys, mlngModul, , Array(lstDept), blnParSet)
    If str����ID = "" Then
        For i = 0 To lstDept.ListCount - 1
            lstDept.Selected(i) = True
        Next
    Else
        For i = 0 To lstDept.ListCount - 1
            lstDept.Selected(i) = InStr(1, "," & str����ID & ",", "," & lstDept.ItemData(i) & ",") > 0
        Next
    End If
    If lstDept.ListCount > 0 Then lstDept.TopIndex = 0: lstDept.ListIndex = 0
    
    
    txtNameDays.Text = zlDatabase.GetPara("������������", glngSys, mlngModul, , Array(txtNameDays), blnParSet)
    txtNameDays.Tag = IIf(txtNameDays.Enabled, "1", "0")
    chkSeekName.Value = IIf(zlDatabase.GetPara("����ģ������", glngSys, mlngModul, , Array(chkSeekName), blnParSet) = "1", 1, 0)
    
    chkDoctor.Value = IIf(zlDatabase.GetPara("����ҽ��", glngSys, mlngModul, , Array(chkDoctor), blnParSet) = "1", 1, 0)
    chkTotal.Value = IIf(zlDatabase.GetPara("�ɿ�ҺŽ���", glngSys, mlngModul, , Array(chkTotal), blnParSet) = "1", 1, 0)
    chkPrintFree.Value = IIf(zlDatabase.GetPara("����ô�ӡ", glngSys, mlngModul, , Array(chkPrintFree), blnParSet) = "1", 1, 0)
    
    i = Val(zlDatabase.GetPara("�Һŷ�Ʊ��ӡ��ʽ", glngSys, mlngModul, 1, Array(optPrintFact(0), optPrintFact(1), optPrintFact(2)), blnParSet))
    If i <= optPrintFact.UBound Then optPrintFact(i).Value = True
    i = Val(zlDatabase.GetPara("���������ӡ��ʽ", glngSys, mlngModul, 1, Array(optPrepayPrint(0), optPrepayPrint(1), optPrepayPrint(2), cmdPrintSet(0)), blnParSet))
    If i <= optPrepayPrint.UBound Then optPrepayPrint(i).Value = True
    
    '����:56274
    i = Val(zlDatabase.GetPara("ԤԼ�Һŵ���ӡ��ʽ", glngSys, mlngModul, 1, Array(optPrintBespeak(0), optPrintBespeak(1), optPrintBespeak(2), cmdPrintSet(2)), blnParSet))
    If i <= optPrintBespeak.UBound Then optPrintBespeak(i).Value = True
    '68408
    i = Val(zlDatabase.GetPara("�Һ�ƾ����ӡ��ʽ", glngSys, mlngModul, 1, Array(optSlipPrint(0), optSlipPrint(1), optSlipPrint(2), cmdPrintSet(3)), blnParSet))
    If i <= optSlipPrint.UBound Then optSlipPrint(i).Value = True
    
    i = Val(zlDatabase.GetPara("�˺Żص���ӡ��ʽ", glngSys, mlngModul, 1, Array(optReceipt(0), optReceipt(1), optReceipt(2), cmdPrintSet(4)), blnParSet))
    If i <= optSlipPrint.UBound Then optReceipt(i).Value = True
    
    i = Val(zlDatabase.GetPara("��ӡ������ǩ", glngSys, mlngModul, 0, Array(chkPrintCase), blnParSet))
    chkPrintCase.Value = i
    chkAllowPatientInput.Value = IIf(zlDatabase.GetPara("��������", glngSys, mlngModul, , Array(chkAllowPatientInput), blnParSet) = "1", 1, 0)
    chkAllowSexInput.Enabled = True
    chkAllowSexInput.Value = IIf(zlDatabase.GetPara("�����Ա�", glngSys, mlngModul, , Array(chkAllowSexInput), blnParSet) = "1", 1, 0)
    chkAllowSexInput.Tag = IIf(chkAllowSexInput.Enabled, "1", "0")
    chkAllowAgeInput.Enabled = True
    chkAllowAgeInput.Value = IIf(zlDatabase.GetPara("��������", glngSys, mlngModul, , Array(chkAllowAgeInput), blnParSet) = "1", 1, 0)
    chkAllowAgeInput.Tag = IIf(chkAllowAgeInput.Enabled, "1", "0")
    chkAllowAddressInput.Enabled = True
    chkAllowAddressInput.Value = IIf(zlDatabase.GetPara("�����ͥ��ַ", glngSys, mlngModul, , Array(chkAllowAddressInput), blnParSet) = "1", 1, 0)
    chkAllowAddressInput.Tag = IIf(chkAllowAddressInput.Enabled, "1", "0")
    chkAllowPayModeInput.Enabled = True
    chkAllowPayModeInput.Value = IIf(zlDatabase.GetPara("���븶�ʽ", glngSys, mlngModul, , Array(chkAllowPayModeInput), blnParSet) = "1", 1, 0)
    chkAllowPayModeInput.Tag = IIf(chkAllowPayModeInput.Enabled, "1", "0")
    
    chkAllowFeeTypeInput.Value = IIf(zlDatabase.GetPara("����ѱ�", glngSys, mlngModul, , Array(chkAllowFeeTypeInput), blnParSet) = "1", 1, 0)
    '31724
    chkAllowZyRigist.Value = IIf(zlDatabase.GetPara("����סԺ���˹Һ�", glngSys, mlngModul, , Array(chkAllowZyRigist), blnParSet) = "1", 1, 0)
    chkAllowBalanceInput.Value = IIf(zlDatabase.GetPara("������㷽ʽ", glngSys, mlngModul, , Array(chkAllowBalanceInput), blnParSet) = "1", 1, 0)
    chkAllowPhoneInput.Value = IIf(zlDatabase.GetPara("������ϵ�绰", glngSys, mlngModul, , Array(chkAllowPhoneInput), blnParSet) = "1", 1, 0)
    Call chkAllowPatientInput_Click
    
    '��ȡ���ù��ùҺ�Ʊ������
    Call LoadFactList(4)
    
            
    chkNewCardNoPop.Value = IIf(zlDatabase.GetPara("������������", glngSys, mlngModul, , Array(chkNewCardNoPop), blnParSet) = "1", 1, 0)
    chkAutoAddName.Value = IIf(zlDatabase.GetPara("�Զ���������", glngSys, mlngModul, , Array(chkAutoAddName), blnParSet) = "1", 1, 0)
    
    
    chkRePrint.Value = IIf(zlDatabase.GetPara("�˷��ش�", glngSys, mlngModul, , Array(chkRePrint), blnParSet) = "1", 1, 0)
    
    chkCardMoney.Value = IIf(zlDatabase.GetPara("��ȡ����", glngSys, mlngModul, , Array(chkCardMoney), blnParSet) = "1", 1, 0)
    Call chkCardMoney_Click
    '36028
    chkMzh.Value = IIf(zlDatabase.GetPara("ԤԼ�����������", glngSys, mlngModul, , Array(chkMzh), blnParSet) = "1", 1, 0)
    '92468:���ϴ�,2016/1/25,�ϸ�����¿���Ϊ0Ҳ��Ʊ��
    ChkMustBill.Value = IIf(zlDatabase.GetPara("�㿨����Ʊ��", glngSys, mlngModul, 0, Array(ChkMustBill), blnParSet) = "1", 1, 0)
    
    
    '�Һ��б��������,�Ƿ��ѷ���ʱ��Ϊ׼
'    ��chkRegList.Value = IIf(zlDatabase.GetPara("������ʱ����ʾ��¼", glngSys, mlngModul, , Array(chkRegList), blnParSet) = "1", 1, 0)

    '��ȡ���õľ��￨����
     Call InitShareInvoice
    If Tabs1.TabVisible(0) Then Tabs1.Tab = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lvwBill_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim i As Integer
    For i = 1 To lvwBill.ListItems.Count
        If lvwBill.ListItems(i).Key <> Item.Key Then lvwBill.ListItems(i).Checked = False
    Next
    Item.Selected = True
End Sub

Private Sub pic��ǰ��ɫ_Click()
    dlgColor.ShowColor
    mstrColor = dlgColor.Color
    pic��ǰ��ɫ.BackColor = mstrColor
End Sub

Private Sub txtCancelBespeak_Change()
    '�����:56407
    If Val(txtCancelBespeak.Text) = 0 Then
        chkBackNoToVerfy.Caption = "������ȡ��ԤԼ��Ҫͨ�����"
    Else
        chkBackNoToVerfy.Caption = "��" & txtCancelBespeak.Text & "����ȡ��ԤԼ��Ҫͨ�����"
    End If
End Sub

Private Sub txtInterval_GotFocus()
    Call SelAll(txtInterval)
End Sub

Private Sub txtInterval_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtInterval_Validate(Cancel As Boolean)
    If Val(txtInterval.Text) < 1 Then
        txtInterval.Text = 1
    ElseIf Val(txtInterval.Text) > 99 Then
        txtInterval.Text = 99
    End If
End Sub

Private Sub txtMustGuardianInfo_GotFocus()
    zlControl.TxtSelAll txtMustGuardianInfo
End Sub

Private Sub txtMustGuardianInfo_KeyPress(KeyAscii As Integer)
     If KeyAscii = Asc("-") Then KeyAscii = 0: Exit Sub
    zlControl.TxtCheckKeyPress txtMustGuardianInfo, KeyAscii, m����ʽ
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

Private Sub txtCancelBespeak_GotFocus()
    zlControl.TxtSelAll txtCancelBespeak
End Sub

Private Sub txtDeptNums_GotFocus()
    zlControl.TxtSelAll txtDeptNums
End Sub

Private Sub txtCancelBespeak_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtCancelBespeak_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtCancelBespeak, KeyAscii, m����ʽ
End Sub

Private Sub txtDeptNums_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtDeptNums_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtDeptNums, KeyAscii, m����ʽ
End Sub

Private Sub txtBreakAnAppointmentNums_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtAvailabilityTimes_GotFocus()
    zlControl.TxtSelAll txtAvailabilityTimes
End Sub

Private Sub txtAvailabilityTimes_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtAvailabilityTimes_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtAvailabilityTimes, KeyAscii, m����ʽ
End Sub
Private Sub txtBreakAnAppointmentNums_GotFocus()
    zlControl.TxtSelAll txtBreakAnAppointmentNums
End Sub

Private Sub txtBreakAnAppointmentNums_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtBreakAnAppointmentNums, KeyAscii, m����ʽ
End Sub
Private Sub InitShareInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ù���Ʊ
    '����:���˺�
    '����:2011-07-06 18:41:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, lngRow As Long
    Dim strShareInvoice As String '����Ʊ������,��ʽ:����,����
    Dim varData As Variant, varTemp As Variant, VarType As Variant, varTemp1 As Variant
    Dim intType As Integer, intType1 As Integer   '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    Dim lngTemp As Long, i As Long, strSQL As String, rsҽ�ƿ���� As ADODB.Recordset
    Dim strPrintMode As String, blnHavePrivs As Boolean, lngCardTypeID As Long
    Dim strȱʡҽ�ƿ� As String, lngȱʡҽ�ƿ� As Long
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    '�ָ��п��
    lngCardTypeID = Val(zlDatabase.GetPara("ȱʡҽ�ƿ����", glngSys, mlngModul, , , True, intType))
    gstrSQL = "Select ID,����,����, nvl(�Ƿ�̶�,0) as �Ƿ�̶�  from ҽ�ƿ����  Where nvl(�Ƿ�����,0)=1"
    On Error GoTo Hd
    Set rsҽ�ƿ���� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    rsҽ�ƿ����.Filter = "����='���￨' and �Ƿ�̶�=1"
    If rsҽ�ƿ����.EOF = False Then
        strȱʡҽ�ƿ� = rsҽ�ƿ����!����: lngȱʡҽ�ƿ� = Val(rsҽ�ƿ����!ID)
    End If
    With rsҽ�ƿ����
        cboType.Clear
        rsҽ�ƿ����.Filter = 0
        If rsҽ�ƿ����.RecordCount <> 0 Then rsҽ�ƿ����.MoveFirst
        Do While Not .EOF
            cboType.AddItem Nvl(!����)
            cboType.ItemData(cboType.NewIndex) = Nvl(!ID)
            If Nvl(!����) = "���￨" And cboType.ListIndex < 0 Then cboType.ListIndex = cboType.NewIndex
            If lngCardTypeID = Val(Nvl(!ID)) Then
                cboType.ListIndex = cboType.NewIndex
            End If
            .MoveNext
        Loop
    End With
    zl_vsGrid_Para_Restore mlngModul, vsBill, Me.Name, "����ҽ��Ʊ���б�", False, False
    strShareInvoice = zlDatabase.GetPara("����ҽ�ƿ�����", glngSys, mlngModul, , , True)
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
    
    '��ʽ:����ID1,ҽ�ƿ����ID1|����IDn,ҽ�ƿ����IDn|...
    varData = Split(strShareInvoice, "|")
    '1.���ù���Ʊ��
    Set rsTemp = GetShareInvoiceGroupID(5)
    With vsBill
        .Clear 1: .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFixedOnly
        .MergeCol(0) = True
        Do While Not rsTemp.EOF
            .RowData(lngRow) = Val(Nvl(rsTemp!ID))
            If Val(Nvl(rsTemp!ʹ�����ID)) = 0 Then
                .TextMatrix(lngRow, .ColIndex("ҽ�ƿ����")) = strȱʡҽ�ƿ�
                .Cell(flexcpData, lngRow, .ColIndex("ҽ�ƿ����")) = lngȱʡҽ�ƿ�
            Else
                rsҽ�ƿ����.Filter = "ID=" & Val(Nvl(rsTemp!ʹ�����ID))
                If Not rsҽ�ƿ����.EOF Then
                    .TextMatrix(lngRow, .ColIndex("ҽ�ƿ����")) = Nvl(rsҽ�ƿ����!����)
                Else
                    .TextMatrix(lngRow, .ColIndex("ҽ�ƿ����")) = Nvl(rsTemp!ʹ�����)
                End If
                .Cell(flexcpData, lngRow, .ColIndex("ҽ�ƿ����")) = Val(Nvl(rsTemp!ʹ�����ID))
            End If
            .TextMatrix(lngRow, .ColIndex("������")) = Nvl(rsTemp!������)
            .TextMatrix(lngRow, .ColIndex("��������")) = Format(rsTemp!�Ǽ�ʱ��, "yyyy-MM-dd")
            .TextMatrix(lngRow, .ColIndex("���뷶Χ")) = rsTemp!��ʼ���� & "," & rsTemp!��ֹ����
            .TextMatrix(lngRow, .ColIndex("ʣ��")) = Format(Val(Nvl(rsTemp!ʣ������)), "##0;-##0;;")
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & ",", ",")
                lngTemp = Val(varTemp(0))
                If Val(.RowData(lngRow)) = lngTemp _
                    And Val(varTemp(1)) = Val(.Cell(flexcpData, lngRow, .ColIndex("ҽ�ƿ����"))) Then
                    .TextMatrix(lngRow, .ColIndex("ѡ��")) = -1: Exit For
                End If
            Next
            .MergeRow(lngRow) = True
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Sub

Private Sub SaveInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������Ʊ��
    '����:���˺�
    '����:2011-07-06 18:27:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHavePrivs As Boolean, strValue As String
    Dim i As Long, lng�����ID As Long
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    '���湲��Ʊ��
    strValue = ""
    With vsBill
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("ѡ��"))) <> 0 And Val(.RowData(i)) <> 0 Then
                strValue = strValue & "|" & Val(.RowData(i)) & "," & Trim(.Cell(flexcpData, i, .ColIndex("ҽ�ƿ����")))
            End If
        Next
    End With
    If strValue <> "" Then strValue = Mid(strValue, 2)
    zlDatabase.SetPara "����ҽ�ƿ�����", strValue, glngSys, mlngModul, blnHavePrivs
    If cboType.ListIndex >= 0 Then
        lng�����ID = cboType.ItemData(cboType.ListIndex)
    End If
    Call zlDatabase.SetPara("ȱʡҽ�ƿ����", lng�����ID, glngSys, mlngModul, blnHavePrivs)
End Sub

Private Sub txtBespeakMinTime_GotFocus()
    zlControl.TxtSelAll txtBespeakMinTime
End Sub
Private Sub txtBespeakMinTime_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtBespeakMinTime_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtBespeakMinTime, KeyAscii, m����ʽ
End Sub
 
Private Sub txtBespeakDefaultDays_GotFocus()
    zlControl.TxtSelAll txtBespeakDefaultDays
End Sub
Private Sub txtBespeakDefaultDays_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtBespeakDefaultDays_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtBespeakDefaultDays, KeyAscii, m����ʽ
End Sub

'Public Sub ShowParSet(ByVal frmMain As Object, ByRef blnCancel As Boolean)
'    '��ʾ��������
'    mblnOK = False
'    If frmMain Is Nothing Then
'        Me.Show 1
'    Else
'        Me.Show 1, frmMain
'    End If
'
'    blnCancel = Not mblnOK
'End Sub
