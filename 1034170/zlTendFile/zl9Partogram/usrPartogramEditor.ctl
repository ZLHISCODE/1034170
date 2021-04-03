VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.UserControl usrPartogramEditor 
   AutoRedraw      =   -1  'True
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8565
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   5550
   ScaleWidth      =   8565
   Begin VB.PictureBox picTmp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
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
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   840
      ScaleHeight     =   300
      ScaleWidth      =   1005
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   120
      Width           =   1000
      Begin VB.ComboBox cboBaby 
         Height          =   300
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   0
         Width           =   1005
      End
   End
   Begin MSComctlLib.ImageList imgRow 
      Left            =   7200
      Top             =   600
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
            Picture         =   "usrPartogramEditor.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrPartogramEditor.ctx":039A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picCloumn 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   3075
      Left            =   0
      ScaleHeight     =   3075
      ScaleWidth      =   5955
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   5955
      Begin VB.PictureBox picBiref 
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         Height          =   2145
         Left            =   240
         ScaleHeight     =   2145
         ScaleWidth      =   5085
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   5085
         Begin VB.CommandButton cmdOk 
            Height          =   315
            Left            =   3750
            Picture         =   "usrPartogramEditor.ctx":0734
            Style           =   1  'Graphical
            TabIndex        =   52
            ToolTipText     =   "ȷ��"
            Top             =   1680
            Width           =   450
         End
         Begin VB.CommandButton cmdCancel 
            Height          =   315
            Left            =   4290
            Picture         =   "usrPartogramEditor.ctx":0CBE
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "ȡ��"
            Top             =   1680
            Width           =   450
         End
         Begin VB.ComboBox cboС�� 
            Height          =   300
            Left            =   3480
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   180
            Width           =   1365
         End
         Begin VB.TextBox txtС������ 
            Height          =   300
            Left            =   930
            TabIndex        =   49
            Top             =   1260
            Width           =   3885
         End
         Begin VB.ComboBox cbo��ʶ 
            Height          =   300
            Left            =   930
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   900
            Width           =   3915
         End
         Begin VB.TextBox txt��ʼʱ�� 
            Enabled         =   0   'False
            Height          =   300
            Left            =   930
            MaxLength       =   5
            TabIndex        =   47
            Top             =   540
            Width           =   1365
         End
         Begin VB.TextBox txt����ʱ�� 
            Enabled         =   0   'False
            Height          =   300
            Left            =   3480
            MaxLength       =   5
            TabIndex        =   46
            Top             =   540
            Width           =   1365
         End
         Begin MSComCtl2.DTPicker DTPDate 
            Height          =   300
            Left            =   930
            TabIndex        =   45
            Top             =   180
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   133300227
            CurrentDate     =   40805
         End
         Begin VB.Label lblС�� 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "С��"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   3060
            TabIndex        =   59
            Top             =   240
            Width           =   360
         End
         Begin VB.Label lblС������ 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   510
            TabIndex        =   58
            Top             =   1320
            Width           =   360
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�������ݺ���ʾ��ȷ����"
            ForeColor       =   &H000000FF&
            Height          =   180
            Left            =   690
            TabIndex        =   57
            Top             =   1770
            Width           =   2010
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblʱ�� 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ʱ��"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   510
            TabIndex        =   56
            Top             =   240
            Width           =   360
         End
         Begin VB.Label lbl��ʶ 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��ʶ"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   510
            TabIndex        =   55
            Top             =   960
            Width           =   360
         End
         Begin VB.Label lbl��ʼʱ�� 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��ʼʱ��"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   150
            TabIndex        =   54
            Top             =   600
            Width           =   720
         End
         Begin VB.Label lbl����ʱ�� 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�� ����ʱ��"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   2430
            TabIndex        =   53
            Top             =   600
            Width           =   990
         End
      End
      Begin MSComctlLib.ListView lstColumnItems 
         Height          =   2490
         Left            =   180
         TabIndex        =   24
         Top             =   450
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   4392
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "��Ŀ���"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "��Ŀ����"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "��λ"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.CommandButton cmdFilterOK 
         Height          =   315
         Left            =   2460
         Picture         =   "usrPartogramEditor.ctx":1248
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "ȷ��"
         Top             =   2310
         Width           =   450
      End
      Begin VB.CommandButton cmdFilterCancel 
         Height          =   315
         Left            =   3000
         Picture         =   "usrPartogramEditor.ctx":17D2
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "ȡ��"
         Top             =   2310
         Width           =   450
      End
      Begin VB.CommandButton cmdColumn 
         Caption         =   "ѡ��(&S)"
         Height          =   300
         Index           =   0
         Left            =   2430
         TabIndex        =   25
         Top             =   1245
         Width           =   1100
      End
      Begin VB.CommandButton cmdColumn 
         Caption         =   "ɾ��(&E)"
         Height          =   300
         Index           =   1
         Left            =   2430
         TabIndex        =   26
         Top             =   1575
         Width           =   1100
      End
      Begin VB.TextBox txtColumnNo 
         Height          =   300
         Left            =   4545
         MaxLength       =   20
         TabIndex        =   30
         Top             =   120
         Width           =   1185
      End
      Begin MSComctlLib.ListView lstColumnUsed 
         Height          =   2490
         Left            =   3720
         TabIndex        =   31
         Top             =   450
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   4392
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "��Ŀ���"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "��Ŀ����"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "��λ"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label lblNote 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�ѷ������ݣ�������������á�"
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   2340
         TabIndex        =   32
         Top             =   690
         Width           =   1260
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblColumnItems 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ѡ�����¼��Ŀ:"
         Height          =   180
         Left            =   240
         TabIndex        =   23
         Top             =   180
         Width           =   1530
      End
      Begin VB.Label lblColumnNo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ͷ����:"
         Height          =   180
         Left            =   3735
         TabIndex        =   29
         Top             =   180
         Width           =   810
      End
   End
   Begin VB.TextBox txtLength 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   1005
      Left            =   2340
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   3090
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4515
      Left            =   60
      ScaleHeight     =   4515
      ScaleWidth      =   8385
      TabIndex        =   10
      Top             =   510
      Width           =   8385
      Begin VB.PictureBox picBaby 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   5040
         ScaleHeight     =   1695
         ScaleWidth      =   1965
         TabIndex        =   61
         Top             =   120
         Visible         =   0   'False
         Width           =   1965
         Begin VSFlex8Ctl.VSFlexGrid vsfBaby 
            Height          =   1300
            Left            =   0
            TabIndex        =   62
            Top             =   0
            Width           =   1935
            _cx             =   3413
            _cy             =   2311
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
            BackColor       =   -2147483624
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16764057
            ForeColorSel    =   0
            BackColorBkg    =   -2147483624
            BackColorAlternate=   -2147483624
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   0
            GridLinesFixed  =   0
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   1
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   300
            ColWidthMin     =   1900
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   0   'False
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
            Begin VB.CommandButton cmdDelBaby 
               Height          =   300
               Left            =   1600
               Picture         =   "usrPartogramEditor.ctx":1D5C
               Style           =   1  'Graphical
               TabIndex        =   63
               ToolTipText     =   "ɾ��"
               Top             =   300
               Width           =   300
            End
         End
         Begin VB.CommandButton cmdBabyCancle 
            Height          =   315
            Left            =   1440
            Picture         =   "usrPartogramEditor.ctx":22E6
            Style           =   1  'Graphical
            TabIndex        =   65
            ToolTipText     =   "ȡ��"
            Top             =   1320
            Width           =   450
         End
         Begin VB.CommandButton cmdAddBaby 
            Height          =   315
            Left            =   960
            Picture         =   "usrPartogramEditor.ctx":2870
            Style           =   1  'Graphical
            TabIndex        =   64
            ToolTipText     =   "���"
            Top             =   1320
            Width           =   450
         End
      End
      Begin VB.PictureBox picDoubleChoose 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   6330
         ScaleHeight     =   240
         ScaleWidth      =   900
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   3300
         Visible         =   0   'False
         Width           =   930
         Begin VB.PictureBox picChooseRight 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   540
            ScaleHeight     =   255
            ScaleWidth      =   375
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   0
            Width           =   375
            Begin VB.ComboBox cboChoose 
               BackColor       =   &H80000018&
               Height          =   300
               Index           =   1
               ItemData        =   "usrPartogramEditor.ctx":2DFA
               Left            =   -30
               List            =   "usrPartogramEditor.ctx":2E0A
               Style           =   2  'Dropdown List
               TabIndex        =   39
               Top             =   -30
               Width           =   1605
            End
         End
         Begin VB.PictureBox picChooseLeft 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   435
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   0
            Width           =   435
            Begin VB.ComboBox cboChoose 
               BackColor       =   &H80000018&
               Height          =   300
               Index           =   0
               ItemData        =   "usrPartogramEditor.ctx":2E1C
               Left            =   -30
               List            =   "usrPartogramEditor.ctx":2E2C
               Style           =   2  'Dropdown List
               TabIndex        =   37
               Top             =   -30
               Width           =   1605
            End
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "/"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   435
            TabIndex        =   40
            Top             =   30
            Width           =   105
         End
      End
      Begin VB.PictureBox picMutilInput 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   6060
         ScaleHeight     =   405
         ScaleWidth      =   1575
         TabIndex        =   8
         Top             =   3720
         Visible         =   0   'False
         Width           =   1600
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   0
            Left            =   810
            TabIndex        =   9
            Top             =   90
            Width           =   675
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "������¼"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   15
            TabIndex        =   13
            Top             =   112
            Width           =   720
         End
      End
      Begin VB.CheckBox chkSwitch 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ȫѡ"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   0
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   930
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.CommandButton cmdWord 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6120
         Picture         =   "usrPartogramEditor.ctx":2E3E
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1290
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.PictureBox picDouble 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   6330
         ScaleHeight     =   240
         ScaleWidth      =   900
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2910
         Visible         =   0   'False
         Width           =   930
         Begin VB.PictureBox picDnInput 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   540
            ScaleHeight     =   255
            ScaleWidth      =   375
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   0
            Width           =   375
            Begin VB.Label lblDnInput 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   60
               TabIndex        =   20
               Top             =   30
               Width           =   315
            End
         End
         Begin VB.PictureBox picUpInput 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   435
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   0
            Width           =   435
            Begin VB.Label lblUpInput 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H80000008&
               Height          =   165
               Left            =   60
               TabIndex        =   19
               Top             =   30
               Width           =   315
            End
         End
         Begin VB.TextBox txtDnInput 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   525
            MaxLength       =   12
            TabIndex        =   7
            Top             =   30
            Width           =   345
         End
         Begin VB.TextBox txtUpInput 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   30
            MaxLength       =   12
            TabIndex        =   6
            Top             =   30
            Width           =   375
         End
         Begin VB.Label lblSplit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "/"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   435
            TabIndex        =   16
            Top             =   30
            Width           =   105
         End
      End
      Begin VB.ListBox lstSelect 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   2550
         Index           =   1
         ItemData        =   "usrPartogramEditor.ctx":3180
         Left            =   6660
         List            =   "usrPartogramEditor.ctx":3196
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   1590
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.PictureBox picInput 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5790
         ScaleHeight     =   225
         ScaleWidth      =   585
         TabIndex        =   1
         Top             =   1290
         Visible         =   0   'False
         Width           =   615
         Begin VB.TextBox txtInput 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   2
            Top             =   30
            Width           =   315
         End
         Begin VB.Label lblInput 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Caption         =   "��"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   0
            TabIndex        =   14
            Top             =   30
            Width           =   315
         End
      End
      Begin VB.ListBox lstSelect 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   2550
         Index           =   0
         ItemData        =   "usrPartogramEditor.ctx":31CE
         Left            =   5790
         List            =   "usrPartogramEditor.ctx":31E4
         TabIndex        =   3
         Top             =   1590
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VSFlex8Ctl.VSFlexGrid VsfData 
         Height          =   2655
         Left            =   0
         TabIndex        =   0
         Top             =   930
         Width           =   4305
         _cx             =   7594
         _cy             =   4683
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
         BackColorSel    =   16764057
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   3
         FixedCols       =   1
         RowHeightMin    =   255
         RowHeightMax    =   5000
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"usrPartogramEditor.ctx":321C
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
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
         AutoSizeMouse   =   0   'False
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
      Begin VSFlex8Ctl.VSFlexGrid vsTest 
         Height          =   495
         Left            =   1920
         TabIndex        =   41
         Top             =   930
         Visible         =   0   'False
         Width           =   1845
         _cx             =   3254
         _cy             =   873
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
         BackColorSel    =   16764057
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   3
         FixedCols       =   1
         RowHeightMin    =   255
         RowHeightMax    =   5000
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"usrPartogramEditor.ctx":327E
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   1
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
         OwnerDraw       =   1
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
         AutoSizeMouse   =   0   'False
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
      Begin VB.Label lblSubEnd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ע:##"
         Height          =   180
         Left            =   1320
         TabIndex        =   60
         Top             =   480
         Width           =   630
      End
      Begin VB.Label lblCurPage 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "P333"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   7650
         TabIndex        =   34
         Top             =   240
         Width           =   555
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "һ�㻤���¼��"
         Height          =   180
         Left            =   3450
         TabIndex        =   12
         Top             =   30
         Width           =   1275
      End
      Begin VB.Label lblSubhead 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:##"
         Height          =   180
         Left            =   390
         TabIndex        =   11
         Top             =   540
         Width           =   720
         WordWrap        =   -1  'True
      End
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "usrPartogramEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'��������:
'1.�����¼ͬһʱ��ֻ���ܴ���һ����¼
'2.�����¼�в���Ҫ�����µ����� , ��¼�����Ƿ����, �ܲ������, �����˵����ݲż�¼
'3.¼�뻤���¼����ʱ,�����¼������ݴ�����������, ����ȡ����
'4.�����¼���в���Ҫ¼�������¼�������׾����ȷ��Ҫ��¼���ڻ���ժҪ�������͵�����
'#ʵ��ԭ��:
'1.�����û��޸Ĺ�������,�����ṩ�༭״̬ҳ���л��Ĺ���,���û��޸Ĺ���ҳ���ݽ�����ҳ����,���ٳ���ʵ���Ѷ�
'2.���Ӽ�¼����¼��Щҳ��Щ��Ԫ���û��޸Ĺ�
'3.�κα༭(ճ��,�������),����Ҫ���¼���ÿ�����ݵ�ռ����

Public mblnEditable As Boolean
'Public objFileSys As New FileSystemObject
'Public objStream As TextStream

Private mblnRestore As Boolean              '���¼������ݻ��ǻָ�ҳ������
Private mFrmParent As Object
Private mblnInit As Boolean
Private mblnShow As Boolean                 '�Ƿ���ʾ¼���
Private mblnVerify As Boolean               '�Ƿ���ǩģʽ(���޸�,����������и���ճ������Ȳ���,ֻ���޸�)
Private mstrVerify As String                '�ȴ���ǩ��ID��
Private mintVerify As Integer               '��ǰ����Ա����߼���
Private mintVerify_Last As Integer          '��ѡ��ǩ��¼����߼���
Private mblnBlowup As Boolean               '�Ŵ�񣿷Ŵ�1/3��������9�ŷŴ�Ϊ12��
Private mblnChange As Boolean               '�Ƿ��޸�����
Private mstrData As String                  '����༭״̬ǰ����֮ǰ������
Private mintPreDays As Long
Private mstrMaxDate As String

Private mint����ҳ As Integer
Private mintҳ�� As Integer
Private mlng�ļ�ID As Long
Private mlng��ʽID As Long
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mlng����id As Long
Private mlng����ID As Long
Private mintӤ�� As Integer
Private mlngFileIndex As Long
Private mstrPrivs As String

Private mintSymbol As Integer               '��ǰ�ؼ�����
Private mstrSymbol As String                '�����ַ�
Private mblnClear As Boolean                '���Ϊ��,���mrsDataMap��¼��;����ҳʱӦ����,�����û��޸ĵ������Ա���ʾ������ʹ��
Private mstrCollectItems As String          '������Ŀ����
Private mstrColCollect As String            '������Ŀ�м���:col;1|col;4,5
Private mstrCOLNothing As String            'δ�󶨵��м���+���Ŀ��(���ܻ��Ŀ���Ƿ��)
Private mstrCOLActive As String             '��м���
Private mstrCatercorner As String           '�жԽ��߼���
Private mblnEditAssistant As Boolean        '��ǰѡ�����Ŀ�Ƿ�������дʾ�ѡ��
Private mlngPageRows As Long                '���ļ���ʽһҳ����ʾ��������
Private mlngLitterRows() As String           '��¼��ҳ�����е�����
Private mlngOverrunRows As Long             '����������
Private mlngRowCount As Long                '��ǰ��¼������
Private mlngRowCurrent As Long              '��ǰ��¼�ڱ�ҳ��ʵ������
Private mlngDate As Long                    '����
Private mlngTime As Long                    'ʱ��
Private mlngSpread As Long                  '��������
Private mlngJust As Long                    '��¶����
Private mlngProduce As Long                 '����
Private mlngChoose As Long                  'ѡ����
Private mlngOperator As Long                '��ʿ
Private mlngSignLevel As Long               'ǩ������
Private mlngSigner As Long                  'ǩ����Ϣ
Private mlngSignName As Long                'ǩ����
Private mlngSignTime As Long                'ǩ��ʱ��
Private mlngRecord As Long                  '��¼ID
Private mlngNoEditor As Long                '��ֹ�༭��,���ڻ�ʿ�����Ի�ʿ��Ϊ׼,�����ڻ�ʿ������ǩ����Ϊ׼
Private mlngCollectType As Long             '�������
Private mlngCollectText As Long             '�����ı�
Private mlngCollectStyle As Long            '���ܱ��
Private mlngCollectDay As Long              '��������:0-����;1-����
Private mlngCollectStart As Long            '���ܿ�ʼʱ��
Private mlngCollectEnd As Long              '���ܽ���ʱ��
Private mlngDemo As Long                    '������
Private mlngActTime As Long                 '����ʱ��

Private mblnClass As Boolean                '�����־
Private mstrClassRow As String              '������ʼ��
Private mblnSign As Boolean                 '�Ƿ�ǩ��
Private mblnArchive As Boolean              '�Ƿ�鵵
Private mintType As Integer                 '��¼��ǰ�ı༭ģʽ
Private mintCollectDef As Integer           'ȱʡС���ʽ
Private mblnDateAd As Boolean               '������д?
Private mblnDate As Boolean                 '�Ƿ����������
Private mstr��ʼʱ�� As String              '��ǰ�ļ��Ŀ�ʼʱ��
Private mstr����ʱ�� As String              '��ǰ�ļ��Ľ���ʱ��
Private mstrBeginTime As String             '������ʼʱ��
Private CellRect As RECT

Private rsTemp As New ADODB.Recordset
Private mrsItems As New ADODB.Recordset             '���л����¼��Ŀ�嵥
Private mrsPartogram As New ADODB.Recordset         '����Ҫ�ؼ�¼�嵥
Private mrsSelItems As New ADODB.Recordset          '��ǰ¼��Ļ����¼��Ŀ�嵥
Private mrsDataMap As New ADODB.Recordset           '��ǰ����Ա¼������ݾ���,���¼����ʽһ��,���������ȫ�������Ա�Ѹ�ٻָ�
Private mrsCellMap As New ADODB.Recordset           '�༭�������ݾ���,�ֶ���:ҳ��,�к�,�к�,��¼ID,����,��λ,ɾ��
Private mrsCopyMap As New ADODB.Recordset           '����������

Private Enum ColIcon
    ǩ�� = 1
    ��ǩ = 2
End Enum
Private Enum SignLevel
    ���� = 1
    ���� = 2
    �м� = 3
    ʦ�� = 4
    Աʿ = 5
    δ���� = 9
End Enum

Private Const conMenu_Save = 1

Private Const WS_MAXIMIZE = &H1000000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_CAPTION = &HC00000
Private Const WS_SYSMENU = &H80000
Private Const WS_THICKFRAME = &H40000
Private Const WS_CHILD = &H40000000
Private Const WS_POPUP = &H80000000
Private Const SWP_NOZORDER = &H4
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

Private Const madLongVarCharDefault As Integer = 10          '�ַ����ֶ�ȱʡ����
Private Const madDoubleDefault As Integer = 18               '�������ֶ�ȱʡ����
Private Const madDbDateDefault As Integer = 20               '�������ֶ�ȱʡ����

Public Event AfterDataChanged(ByVal blnChange As Boolean)
Public Event AfterRefresh()
Public Event AfterRowColChange(ByVal strInfo As String, ByVal blnImportant As Boolean, ByVal blnSign As Boolean, ByVal blnArchive As Boolean)
Public Event AfterDataSave(ByVal blnSave As Boolean)
Public Event AfterFileIndex(ByVal lngFileIndex As Long)
Public Event AfterPartogramInfo(ByVal lngFlieId As Long, ByVal lngFileIndex As Long, ByVal lngFileFormatID As Long, ByVal rsPartogram As ADODB.Recordset)

Dim strFields As String
Dim strValues As String
Dim blnScroll As Boolean

'��¼�ϴ�ѡ����,����,�Ա�ˢ�º����¶�λ
Dim lngLastRow As Long
Dim lngLastTopRow As Long
Dim lngLastPatientID As Long

Private mstrTag As String           '�ݴ�

'�����ļ���ʽ�������
Private mintTabTiers As Integer     '��ͷ���
Private mintTagFormHour As Integer  '��ʼʱ������
Private mintTagToHour As Integer    '��ֹʱ������
Private mobjTagFont As New StdFont  '������ʽ����
Private mlngTagColor As Long        '������ʽ��ɫ
Private mstrPaperSet As String      '��ʽ
Private mstrPageHead As String      'ҳü
Private mstrPageFoot As String      'ҳ��
Private mblnChildForm As Boolean
Private mstrSubHead As String       '���ϱ�ǩ
Private mstrSubEnd As String        '���±�ǩ
Private mstrTabHead As String       '��ͷ��Ԫ
Private mstrColWidth As String      '�п����д�
Private mstrColumns As String       '��ǰ�����ļ����ж�Ӧ����Ŀ
Private lngCurColor As Long, strCurFont As String, objFont As StdFont
'����򿪻����¼�ļ���SQL���������ط�Ҳ��ʹ�ã������޸�
Private mstrSQL�� As String
Private mstrSQL�� As String
Private mstrSQL�� As String
Private mstrSQL���� As String
Private mstrSQL As String

'######################################################################################################################
'**********************************************************************************************************************
'��#�ָ��������ڵĴ��붼���ͼ���,û�±�
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long

Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Const WHITE_BRUSH = 0    '��ɫ����
Private Const cdblWidth As Double = 6          'һ��Ӣ���ַ��Ŀ��
Private Const cHideCols = 3         'ǰ׺������:����,ʱ��,ѡ��
Private Const cControlFields = 2    '��¼��������:ҳ��,�к�

Private Function GetRBGFromOLEColor(ByVal dwOleColour As Long) As Long
    '��VB����ɫת��ΪRGB��ʾ
    Dim clrref As Long
    Dim r As Long, g As Long, b As Long
    
    OleTranslateColor dwOleColour, 0, clrref
    
    b = (clrref \ 65536) And &HFF
    g = (clrref \ 256) And &HFF
    r = clrref And &HFF
    
    GetRBGFromOLEColor = RGB(r, g, b)
End Function

Private Function GetSymbolWidth(ByVal strPara As String) As Double
    'ȱʡ������9��,�������Сͬ�ȷŴ�
    Dim sinFontSize As Single
    Dim i As Integer, j As Integer
    
    j = Len(strPara)
    sinFontSize = VsfData.FontSize
    For i = 1 To j
        GetSymbolWidth = GetSymbolWidth + IIf(Asc(Mid(strPara, i, 1)) > 0, 1, 2) * cdblWidth * sinFontSize / 9
    Next
End Function

Private Sub DrawCell(ByVal hDC As Long, ByVal ROW As Long, ByVal COL As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim strText As String
    Dim strLeft As String
    Dim strRight As String
    Dim lngLeft As Long
    Dim lngRight As Long
    Dim dblWidth As Double
    Dim lngBackColor As Long
    Dim lngForeColor As Long
    Dim blnDraw As Boolean
    '��ͼ���
    Dim lngPen As Long
    Dim lngOldPen As Long
    Dim lngBrush As Long
    Dim lngOldBrush As Long
    Dim lpPoint As POINTAPI
    Dim T_ClientRect As RECT
    On Error GoTo errHand
    '******************************************
    '�ڴ��¼��в��ܶԵ�Ԫ����κ����Ը�ֵ,����Celldata,�����������¼�����ѭ��,���¹��������ʱ���޷�����������
    '******************************************
    'ʹ��ƥ��ı���ɫ��ǰ��ɫ����������ı������
    If Not mblnInit Then Exit Sub
    If VsfData.RowHidden(ROW) Then Exit Sub
    Done = False
    
    strText = VsfData.TextMatrix(ROW, COL)
    If IsDiagonal(COL) And InStr(1, strText, "/") <> 0 Then
        blnDraw = True
        '����ֵ
        strLeft = Split(strText, "/")(0)
        strRight = Mid(strText, InStr(1, strText, "/") + 1)
        lngLeft = LenB(StrConv(strLeft, vbFromUnicode))
        lngRight = LenB(StrConv(strRight, vbFromUnicode))
        'ȡ�ַ����
        dblWidth = GetSymbolWidth(strRight)
        '�趨�ͻ������С
        With T_ClientRect
            .Left = Left + 1
            .Top = Top + 1
            .Right = Right - 1
            .Bottom = Bottom - 1
        End With
        
        '1���������
        '�����뱳��ɫ��ͬ��ˢ��
        If ROW < VsfData.FixedRows Then
            lngBackColor = GetRBGFromOLEColor(VsfData.BackColorFixed)
            lngForeColor = GetRBGFromOLEColor(VsfData.ForeColorFixed)
        Else
            If ROW = VsfData.RowSel Then
                lngBackColor = GetRBGFromOLEColor(VsfData.BackColorSel)
                lngForeColor = RGB(0, 0, 0)
            Else
                lngBackColor = RGB(255, 255, 255)
                lngForeColor = GetRBGFromOLEColor(VsfData.Cell(flexcpForeColor, ROW, COL))
            End If

        End If
        lngBrush = CreateSolidBrush(lngBackColor)
        'ʹ�ø�ˢ����䱳��ɫ
        lngOldBrush = SelectObject(hDC, lngBrush)
        Call FillRect(hDC, T_ClientRect, lngBrush)
        '����������ʱʹ�õ�ˢ�Ӳ���ԭˢ��
        Call SelectObject(hDC, lngOldBrush)
        Call DeleteObject(lngBrush)
        
        '2��׼������
        '�����»���
        Call SetTextColor(hDC, lngForeColor)
        lngPen = CreatePen(0, 1, lngForeColor)
        lngOldPen = SelectObject(hDC, lngPen)
        '����
        Call MoveToEx(hDC, Left, Bottom - 2, lpPoint)
        Call LineTo(hDC, Right, Top)
        '����ı�
        Call TextOut(hDC, Left, Top, strLeft, lngLeft)
        Call TextOut(hDC, IIf(Right - dblWidth >= Left, Right - dblWidth, Left), Bottom - 16, strRight, lngRight)
        
        '��ԭ���ʲ�����
        Call SelectObject(hDC, lngOldPen)
        Call DeleteObject(lngPen)
        
        '�������ͼ
        Done = True
    End If
    
    '3������ǻ����У���������⴦��
    If Val(VsfData.TextMatrix(ROW, mlngCollectType)) < 0 And Val(VsfData.TextMatrix(ROW, mlngCollectStyle)) > 0 _
        And (COL >= mlngDate And COL < mlngNoEditor) Then
        Call DrawCollectCell(hDC, ROW, COL, Left, Top, Right, Bottom)
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub DrawCollectCell(ByVal hDC As Long, ByVal ROW As Long, ByVal COL As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    Dim lngPen As Long, lngOldPen As Long
    Dim lpPoint As POINTAPI
    
    '�����»���
    lngPen = CreatePen(0, 1, vbRed)
    lngOldPen = SelectObject(hDC, lngPen)
    
    If Val(VsfData.TextMatrix(ROW, mlngCollectStyle)) = 1 Then  '���»�����
        '����
        Call MoveToEx(hDC, Left, Top, lpPoint)
        Call LineTo(hDC, Right, Top)
        Call MoveToEx(hDC, Left, Bottom - 2, lpPoint)
        Call LineTo(hDC, Right, Bottom - 2)
    Else                                                        '��������˫����
        If InStr(1, "|" & mstrColCollect & ";", "|" & COL - (cHideCols + VsfData.FixedCols - 1) & ";") <> 0 Then 'And Val(VsfData.TextMatrix(ROW, COL)) <> 0 Then
            '����
            Call MoveToEx(hDC, Left, Bottom - 4, lpPoint)
            Call LineTo(hDC, Right, Bottom - 4)
            Call MoveToEx(hDC, Left, Bottom - 2, lpPoint)
            Call LineTo(hDC, Right, Bottom - 2)
        End If
    End If
    
    '��ԭ���ʲ�����
    Call SelectObject(hDC, lngOldPen)
    Call DeleteObject(lngPen)
End Sub

'######################################################################################################################
'**********************************************************************************************************************
'��#�ָ��������ڵĴ��붼��������,û�±�
Private Function GetData(ByVal strInput As String) As Variant
    Dim arrData
    Dim strData As String
    Dim strLine(256) As Byte
    Dim lngRow As Long, lngRows As Long, lngLen As Long
    
    GetData = ""
    lngRows = SendMessage(txtLength.Hwnd, EM_GETLINECOUNT, 0&, 0&)
    For lngRow = 1 To lngRows
        Call ClearArray(strLine)
        lngLen = SendMessage(txtLength.Hwnd, EM_GETLINE, lngRow - 1, strLine(0))
        Call ClearArray(strLine, lngLen)
        strData = StrConv(strLine, vbUnicode)
        strData = TruncZero(strData)
        GetData = GetData & IIf(GetData = "", "", "|LPF.ZLSOFT|") & strData & IIf(lngRow < lngRows, vbCrLf, "")
    Next
    GetData = Split(GetData, "|LPF.ZLSOFT|")
End Function

Private Sub ClearArray(strLine() As Byte, Optional ByVal lngPos As Long = 0)
    Dim intDo As Integer, intMax As Integer
    intMax = UBound(strLine)
    For intDo = lngPos To intMax
        strLine(intDo) = 0
        If lngPos > 0 Then Exit Sub     '��Ϊ��,��ʾ�������ַ���������
    Next
    strLine(1) = 1
End Sub

Private Function TrimStr(ByVal str As String) As String
'���ܣ�ȥ���ַ�����\0�Ժ���ַ�������ȥ�����˵Ŀո�

    If InStr(str, Chr(0)) > 0 Then
        TrimStr = Trim(Left(str, InStr(str, Chr(0)) - 1))
    Else
        TrimStr = Trim(str)
    End If
End Function

Private Function TruncZero(ByVal strInput As String) As String
'���ܣ�ȥ���ַ�����\0�Ժ���ַ�
    Dim lngPos As Long
    
    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZero = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZero = strInput
    End If
End Function

'**********************************************************************************************************************
'######################################################################################################################


Private Sub BoundItems(ByVal intCOl As Integer)
    Dim lstItem As ListItem
    Dim rsActive As New ADODB.Recordset
    On Error GoTo errHand
    'ֻ�ṩ������,ѡ����������Ļ��Ŀ
    '�󶨻��Ŀ(��һ����Ŀ������,��������Ŀʱ,��Ŀ���ͱ���=0����Ŀ��ʾֻ������ֵ,ѡ������,��������Ŀ��Ŀ��������Ŀ��ʾ��������һ��)
    
    gstrSQL = "" & _
        " SELECT A.��Ŀ���,A.��λ,A.��Ŀ����,B.��ͷ����,NVL(B.��־,0) AS ��־" & vbNewLine & _
        " FROM" & vbNewLine & _
        "     (SELECT A.��Ŀ���,B.��λ,B.��λ||A.��Ŀ���� AS ��Ŀ����" & vbNewLine & _
        "     FROM �����¼��Ŀ A,���²�λ B" & vbNewLine & _
        "     WHERE A.��Ŀ��� =B.��Ŀ���(+) AND A.��Ŀ����=2 And A.��Ŀ����=0 And A.��Ŀ��ʾ IN (0,4,5)) A," & vbNewLine & _
        "     (SELECT A.��ͷ����,A.��Ŀ���,A.��λ||B.��Ŀ���� AS ��Ŀ����,1 AS ��־" & vbNewLine & _
        "     FROM ���˻�����Ŀ A,�����¼��Ŀ B" & vbNewLine & _
        "     WHERE A.��Ŀ���=B.��Ŀ��� AND A.�ļ�ID=[1] AND A.ҳ��=[2] AND A.�к�=[3]) B" & vbNewLine & _
        " WHERE A.��Ŀ���=B.��Ŀ���(+) AND A.��Ŀ����=B.��Ŀ����(+)" & vbNewLine & _
        " ORDER BY A.��Ŀ���"
    Set rsActive = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡδ���õĻ��Ŀ", mlng�ļ�ID, mintҳ��, intCOl)
    If rsActive.RecordCount = 0 Then
        RaiseEvent AfterRowColChange("û�пɹ�ѡ��Ļ��Ŀ�����ڻ�����Ŀ����ģ���н������ã�", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    '������Ŀ
    lstColumnItems.ListItems.Clear
    lstColumnUsed.ListItems.Clear
    With rsActive
        Do While Not .EOF
            If !��־ = 1 Then
                txtColumnNo.Text = NVL(!��ͷ����)
                Set lstItem = lstColumnUsed.ListItems.Add(, Now & "_" & !��Ŀ��� & "_" & lstColumnUsed.ListItems.Count, !��Ŀ���)
                lstItem.SubItems(1) = !��Ŀ����
                lstItem.SubItems(2) = NVL(!��λ)
            Else
                Set lstItem = lstColumnItems.ListItems.Add(, Now & "_" & !��Ŀ��� & "_" & lstColumnItems.ListItems.Count + 100, !��Ŀ���)
                lstItem.SubItems(1) = !��Ŀ����
                lstItem.SubItems(2) = NVL(!��λ)
            End If
            .MoveNext
        Loop
    End With
    
    '���ÿؼ����꣨��߻��ұ߳�����Ļ��С���һ�����ʾ����������Ϊ������ʾ��
    With picCloumn
        .Left = VsfData.Left + VsfData.CellLeft + VsfData.CellWidth / 2 - .Width / 2
        .Top = picMain.Top + VsfData.Top + VsfData.CellTop
        If .Height + .Top + picMain.Top > ScaleHeight Then
            .Top = ScaleHeight - picMain.Top - .Height
        End If
        If .Left + .Width > ScaleWidth Then
            .Left = ScaleWidth - .Width
        End If
        If .Left < VsfData.Left Then
            .Left = VsfData.Left
        End If
        .Visible = True
    End With
    
    lblNote.Visible = ISColHaveData
    cmdColumn(0).Enabled = Not lblNote.Visible
    cmdColumn(1).Enabled = Not lblNote.Visible
    cmdFilterOK.Enabled = Not lblNote.Visible
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub LoadBabyNum()
    Dim i As Integer
    On Error GoTo errHand
    
    '���ÿؼ����꣨��߻��ұ߳�����Ļ��С���һ�����ʾ����������Ϊ������ʾ��
    With picBaby
        .Left = VsfData.Left
        .Top = 0
        If .Height + .Top + picMain.Top > ScaleHeight Then
            .Top = ScaleHeight - picMain.Top - .Height
        End If
        If .Left + .Width > ScaleWidth Then
            .Left = ScaleWidth - .Width
        End If
        If .Left < VsfData.Left Then
            .Left = VsfData.Left
        End If
        If cboBaby.ListCount > 0 Then
            .Visible = True
        Else
            .Visible = False
            RaiseEvent AfterRowColChange("���ٴ���һ��Ӥ�������뿪������ϵ��", True, mblnSign, mblnArchive)
        End If
    End With
    
    '����Ӥ��������Ϣ
    With vsfBaby
        .FixedCols = 0
        .FixedRows = 0
        .Rows = cboBaby.ListCount
        For i = 0 To cboBaby.ListCount - 1
            .RowData(i) = cboBaby.ItemData(i)
            .TextMatrix(i, 0) = "Ӥ��" & .RowData(i)
        Next i
        .FocusRect = flexFocusHeavy
        .COL = .FixedCols: .ROW = .Rows - 1
        Call vsfBaby_AfterRowColChange(.FixedRows, .FixedCols, .ROW, .COL)
    End With
     
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function GetPeriod() As String
    On Error GoTo errHand
    gstrSQL = " Select   ��Ժ���� AS ��ʼʱ�� From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��Ժ���ڻ��������", mlng����ID, mlng��ҳID)
    GetPeriod = Format(rsTemp!��ʼʱ��, "yyyy-MM-dd HH:mm:ss") & "��" & Format(mstr����ʱ��, "yyyy-MM-dd HH:mm:ss")
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ReadStruDef() As Boolean
    Dim lngCol As Long
    Dim strSQLHead As String, strSQLRow As String
    On Error GoTo errHand
    Debug.Print Now & "ReadStruDef"
    '��ȡ�ļ�����
    mblnDateAd = False
    mblnDate = False
    Call GetFileProperty
    
    '��ȡ���Ŀ�������ж���(��ʽ���к�;��ͷ����|��Ŀ���,��λ;��Ŀ���,��λ||�к�;��ͷ����...)
    mstrCOLActive = ""
    mstrCOLNothing = ""
    mstrCollectItems = ""
    mstrColCollect = ""
'   ����ͼû�л��Ŀ
'    gstrSQL = " Select   A.�к�,A.��ͷ����,A.���,A.��Ŀ���,A.��λ From ���˻�����Ŀ A " & _
'              " Where A.�ļ�ID=[1] And A.ҳ��=[2] " & _
'              " Order by A.�к�,A.���"
'    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�������Զ���Ļ��Ŀ", mlng�ļ�ID, mintҳ��)
'    If rsTemp.RecordCount <> 0 Then
'        Do While Not rsTemp.EOF
'            If lngCOL <> rsTemp!�к� Then
'                lngCOL = rsTemp!�к�
'                mstrCOLActive = mstrCOLActive & "||" & rsTemp!�к� & ";" & rsTemp!��ͷ���� & "|" & rsTemp!��Ŀ��� & "," & NVL(rsTemp!��λ)
'            Else
'                mstrCOLActive = mstrCOLActive & ";" & rsTemp!��Ŀ��� & "," & NVL(rsTemp!��λ)
'            End If
'            rsTemp.MoveNext
'        Loop
'    End If
    If mstrCOLActive <> "" Then mstrCOLActive = Mid(mstrCOLActive, 3)
    
    '��ȡ�����ļ���ʽ����
    gstrSQL = "Select   d.�������, d.�����ı�, d.Ҫ������" & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '�����ʽ'" & _
        " Order By d.�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ļ���ʽ����", mlng��ʽID)
    With rsTemp
        Do While Not .EOF
            Select Case "" & !Ҫ������
            Case "��ͷ����": mintTabTiers = Val("" & !�����ı�)
            Case "������":  VsfData.Cols = Val("" & !�����ı�)
            Case "��С�и�": VsfData.RowHeightMin = BlowUp(Val("" & !�����ı�))
            Case "�ı�����"
                strCurFont = "" & !�����ı�
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = BlowUp(Val(Split(strCurFont, ",")(1)))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                End With
                Set VsfData.Font = objFont
                Set lblSubhead.Font = VsfData.Font
                Set lblSubEnd.Font = VsfData.Font
                Set Font = lblSubhead.Font
                
            Case "�ı���ɫ": VsfData.ForeColor = Val("" & !�����ı�)
            Case "�����ɫ": VsfData.GridColor = Val("" & !�����ı�): VsfData.GridColorFixed = VsfData.GridColor
            
            Case "�����ı�"
                lblTitle.Caption = "" & !�����ı�
                lblTitle.AutoSize = True
            Case "��������"
                strCurFont = "" & !�����ı�
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = BlowUp(Val(Split(strCurFont, ",")(1)))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                End With
                Set lblTitle.Font = objFont
                lblTitle.AutoSize = False
            
            Case "��ʼʱ��": mintTagFormHour = Val("" & !�����ı�)
            Case "��ֹʱ��": mintTagToHour = Val("" & !�����ı�)
            Case "��������"
                strCurFont = "" & !�����ı�
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = BlowUp(Val(Split(strCurFont, ",")(1)))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                End With
                Set mobjTagFont = objFont
            Case "������ɫ": mlngTagColor = Val("" & !�����ı�)
            Case "��Ч������"
                mlngOverrunRows = 0
                mlngPageRows = Val(!�����ı�)
            End Select
            .MoveNext
        Loop
    End With
    
    gstrSQL = "Select   ��ʽ, ҳü, ҳ��,���� From ����ҳ���ʽ Where ���� = 3 And ��� In (Select ҳ�� From �����ļ��б� Where Id = [1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ҳ���ʽ", mlng��ʽID)
    If Not rsTemp.EOF Then
        mstrPaperSet = "" & rsTemp!��ʽ: mstrPageHead = "" & rsTemp!ҳü: mstrPageFoot = "" & rsTemp!ҳ��
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select   d.�������, d.�����ı�, d.Ҫ������, Nvl(d.�Ƿ���, 0) As �Ƿ���" & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '���ϱ�ǩ'" & _
        " Order By d.�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ϱ�ǩ����", mlng��ʽID)
    With rsTemp
        mstrSubHead = ""
        Do While Not .EOF
            mstrSubHead = mstrSubHead & "|" & IIf(!�Ƿ��� = 0, "", vbCrLf) & !�����ı� & "{" & !Ҫ������ & "}"
            .MoveNext
        Loop
        If mstrSubHead <> "" Then mstrSubHead = Replace(Mid(mstrSubHead, 2), Chr(1), " ")
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select   d.�������, d.�����ı�, d.Ҫ������, Nvl(d.�Ƿ���, 0) As �Ƿ���" & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '���±�ǩ'" & _
        " Order By d.�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ϱ�ǩ����", mlng��ʽID)
    With rsTemp
        mstrSubEnd = ""
        Do While Not .EOF
            mstrSubEnd = mstrSubEnd & "|" & IIf(!�Ƿ��� = 0, "", vbCrLf) & !�����ı� & "{" & !Ҫ������ & "}"
            .MoveNext
        Loop
        If mstrSubEnd <> "" Then mstrSubEnd = Replace(Mid(mstrSubEnd, 2), Chr(1), " ")
    End With
    '------------------------------------------------------------------------------------------------------------------
    '����Ƿ����������
    gstrSQL = "Select  d.�������, d.��������, d.�����д�, d.�����ı�, d.Ҫ������, d.Ҫ�ص�λ,d.Ҫ�ر�ʾ" & vbNewLine & _
            "        From �����ļ��ṹ d, �����ļ��ṹ p" & vbNewLine & _
            "        Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '���м���'" & vbNewLine & _
            "        And D.Ҫ������='����'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���м��϶���", mlng��ʽID)
    mblnDate = (rsTemp.RecordCount > 0)
    If mblnDate = False Then VsfData.Cols = VsfData.Cols + 1
    '------------------------------------------------------------------------------------------------------------------
    '������������Ĭ�����������
    If mblnDate = False Then
        gstrSQL = "SELECT �������, �����д�, �����ı�" & vbNewLine & _
                "FROM (SELECT 1 �������, 1 �����д�, '����' �����ı�" & vbNewLine & _
                "       FROM DUAL" & vbNewLine & _
                "       UNION ALL" & vbNewLine & _
                "       SELECT D.�������+1 �������, D.�����д�, D.�����ı�" & vbNewLine & _
                "       FROM �����ļ��ṹ D, �����ļ��ṹ P" & vbNewLine & _
                "       WHERE P.ID = D.��ID AND P.�ļ�ID = [1] AND P.�������� = 1 AND P.�����ı� = '��ͷ��Ԫ')" & vbNewLine & _
                "ORDER BY �������"
    Else
        gstrSQL = "Select   d.�������, d.�����д�, d.�����ı�" & _
            " From �����ļ��ṹ d, �����ļ��ṹ p" & _
            " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '��ͷ��Ԫ'" & _
            " Order By d.�������"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ͷ��Ԫ����", mlng��ʽID)
    With rsTemp
        mstrTabHead = ""
        Do While Not .EOF
            mstrTabHead = mstrTabHead & "|" & !�����д� - 1 & "," & !������� & "," & !�����ı�
            .MoveNext
        Loop
        If mstrTabHead <> "" Then mstrTabHead = Mid(mstrTabHead, 2)
    End With
    
    '��ѯ�����֯
    '------------------------------------------------------------------------------------------------------------------
    Dim strSql�� As String, str��ʽ As String, strSqlNull As String
    Dim bln���� As Boolean, blnʱ�� As Boolean, bln��ʿ As Boolean
    Dim blnǩ���� As Boolean, blnǩ��ʱ�� As Boolean, blnǩ������ As Boolean
    Dim bln�Խ��� As Boolean, blnѡ���� As Boolean          '�����һ���ǶԽ�����ѡ����,��ֱ����ȡ��������,ƴ��ͷʱ����ֵ�����/
    Dim lngColumn As Long, blnAddCollect As Boolean
      
    If mblnDate = False Then
        gstrSQL = "SELECT �������, ��������, �����д�, �����ı�, Ҫ������, Ҫ�ص�λ, Ҫ�ر�ʾ" & vbNewLine & _
            "FROM (SELECT 1 �������, '0`4' ��������, 1 �����д�, '' �����ı�, '����' Ҫ������, '' Ҫ�ص�λ, 0 Ҫ�ر�ʾ" & vbNewLine & _
            "       FROM DUAL" & vbNewLine & _
            "       UNION ALL" & vbNewLine & _
            "       SELECT D.�������+1 �������, D.��������, D.�����д�, D.�����ı�, D.Ҫ������, D.Ҫ�ص�λ, D.Ҫ�ر�ʾ" & vbNewLine & _
            "       FROM �����ļ��ṹ D, �����ļ��ṹ P" & vbNewLine & _
            "       WHERE P.ID = D.��ID AND P.�ļ�ID = [1] AND P.�������� = 1 AND P.�����ı� = '���м���')" & vbNewLine & _
            "ORDER BY �������, �����д�"
    Else
        gstrSQL = "Select   d.�������, d.��������, d.�����д�, d.�����ı�, d.Ҫ������, d.Ҫ�ص�λ,d.Ҫ�ر�ʾ " & _
            " From �����ļ��ṹ d, �����ļ��ṹ p" & _
            " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '���м���'" & _
            " Order By d.�������, d.�����д�"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���м��϶���", mlng��ʽID)
    With rsTemp
        lngColumn = 0: mstrColumns = "": mstrColWidth = "": mstrCatercorner = ""
        mstrSQL�� = "": mstrSQL�� = "": strSql�� = "": mstrSQL�� = "": mstrSQL���� = "": strSqlNull = ""
        bln���� = False: blnʱ�� = False: bln��ʿ = False
        blnǩ���� = False: blnǩ��ʱ�� = False: blnǩ������ = False
        Do While Not .EOF
            If lngColumn <> !������� Then
                blnAddCollect = False
                mstrColumns = mstrColumns & IIf(mstrColumns = "", "", "'1'" & str��ʽ) & "|" & !������� & "'" & !Ҫ������
                mstrColWidth = mstrColWidth & "," & !�������� & "`" & !������� & "`" & !Ҫ�ر�ʾ
                If !Ҫ�ر�ʾ = 1 Then mstrCatercorner = mstrCatercorner & "," & !�������
                str��ʽ = ""
                If !Ҫ������ <> "" Then
                    str��ʽ = "{" & NVL(!�����ı�) & "[" & !Ҫ������ & "]" & NVL(!Ҫ�ص�λ) & "}"
                    'mstrSQL�� = mstrSQL�� & "," & IIf(Mid(strSql��, 3) = "", "''", Mid(strSql��, 3)) & " As C" & Format(lngColumn, "00")
                    If Mid(strSqlNull, 3) = "" Then
                        strSqlNull = "''"
                    Else
                        strSqlNull = Mid(strSqlNull, 3)
                    End If
                    mstrSQL�� = mstrSQL�� & "," & IIf(Mid(strSql��, 3) = "", "''", "Decode(" & Mid(strSql��, 3) & "," & strSqlNull & ",''," & Mid(strSql��, 3) & ")") & " As C" & Format(lngColumn, "00")
                Else
                    If strSql�� <> "" Then
                        mstrSQL�� = mstrSQL�� & "," & Mid(strSql��, 3) & " As C" & Format(lngColumn, "00")
                    Else
                        mstrSQL�� = mstrSQL�� & ",'' As C" & Format(lngColumn, "00")
                    End If
                End If
                strSql�� = ""
                strSqlNull = ""
                lngColumn = !�������
                bln�Խ��� = (NVL(!Ҫ�ر�ʾ, 0) = 1)
                blnѡ���� = False
                mrsItems.Filter = "��Ŀ����='" & NVL(!Ҫ������) & "'"
                If mrsItems.RecordCount <> 0 Then
                    blnѡ���� = (mrsItems!��Ŀ��ʾ = 5)
                    If mrsItems!��Ŀ��ʾ = 4 Then   '������Ŀ
                        blnAddCollect = True
                        mstrCollectItems = mstrCollectItems & "," & mrsItems!��Ŀ���
                        mstrColCollect = mstrColCollect & "|" & !������� & ";" & mrsItems!��Ŀ���
                    End If
                End If
                mrsItems.Filter = 0
            Else
                mstrColumns = mstrColumns & "," & !Ҫ������
                str��ʽ = str��ʽ & "{" & NVL(!�����ı�) & "[" & !Ҫ������ & "]" & NVL(!Ҫ�ص�λ) & "}"
                mrsItems.Filter = "��Ŀ����='" & NVL(!Ҫ������) & "'"
                If mrsItems.RecordCount <> 0 Then
                    If mrsItems!��Ŀ��ʾ = 4 Then   '������Ŀ
                        mstrCollectItems = mstrCollectItems & "," & mrsItems!��Ŀ���
                        If blnAddCollect Then
                            mstrColCollect = mstrColCollect & "," & mrsItems!��Ŀ���
                        Else    '�п���һ�а�������Ŀ,��һ����Ŀ���ǻ�����Ŀ,�ڶ�����Ŀ���ǻ�����Ŀ,���,����Ĵ��뱣֤���������
                            mstrColCollect = mstrColCollect & "|" & !������� & ";" & mrsItems!��Ŀ���
                        End If
                    End If
                End If
                mrsItems.Filter = 0
            End If
            
            Select Case !Ҫ������
            Case "����"
                bln���� = True
                mblnDateAd = (NVL(!Ҫ�ر�ʾ, 0) = 1)
                mstrSQL�� = mstrSQL�� & ",����"
                mstrSQL�� = mstrSQL�� & ",To_Char(l.����ʱ��, " & IIf(mblnDateAd, "'dd/MM'", "'yyyy-mm-dd'") & ") As ����"
                strSql�� = strSql�� & "||" & !Ҫ������
            Case "ʱ��"
                blnʱ�� = True
                mstrSQL�� = mstrSQL�� & ",ʱ��"
                mstrSQL�� = mstrSQL�� & ",To_Char(l.����ʱ��, 'hh24:mi') As ʱ��"
                strSql�� = strSql�� & "||" & !Ҫ������
                
            Case "ǩ����"
                blnǩ���� = True
                mstrSQL�� = mstrSQL�� & ",ǩ����"
                mstrSQL�� = mstrSQL�� & ",l.ǩ����"
                strSql�� = strSql�� & "||" & !Ҫ������
                
            Case "ǩ��ʱ��"
                blnǩ��ʱ�� = True
                mstrSQL�� = mstrSQL�� & ",ǩ��ʱ��"
                mstrSQL�� = mstrSQL�� & ",l.ǩ��ʱ��"
                strSql�� = strSql�� & "||" & !Ҫ������
                
            Case "��ʿ"
                bln��ʿ = True
                mstrSQL�� = mstrSQL�� & ",��ʿ"
                mstrSQL�� = mstrSQL�� & ",l.������ As ��ʿ"
                strSql�� = strSql�� & "||" & !Ҫ������
            Case Else
                If !Ҫ������ <> "" Then
                    mstrSQL�� = mstrSQL�� & ",Max(""" & !Ҫ������ & """) As """ & !Ҫ������ & """"
                    'mstrSQL���� = mstrSQL���� & " Or """ & !Ҫ������ & """ Is Not Null"
                    
                    If bln�Խ��� And blnѡ���� Then
                        If strSql�� <> "" Then
                            '�ڶ���
                            strSql�� = strSql�� & "||'/'||""" & !Ҫ������ & """"
                        Else
                            '��һ��
                            strSql�� = strSql�� & "||""" & !Ҫ������ & """"
                        End If
                    Else
                        strSql�� = strSql�� & "||""" & !Ҫ������ & """"
                        strSqlNull = strSqlNull & "||" & "'" & !�����ı� & "'||'" & !Ҫ�ص�λ & "'"
                    End If
                    
                    If (Trim("" & !�����ı�) = "" And Trim("" & !Ҫ�ص�λ) = "") Or (bln�Խ��� And blnѡ����) Then
                        mstrSQL�� = mstrSQL�� & ", Decode(c.��Ŀ����, '" & !Ҫ������ & "', Nvl(c.δ��˵��,c.��¼����), '') As """ & !Ҫ������ & """"
                        mstrSQL���� = mstrSQL���� & " Or Decode(c.��Ŀ����, '" & !Ҫ������ & "', Nvl(c.δ��˵��,c.��¼����), '') Is Not Null"
                    Else
                        mstrSQL�� = mstrSQL�� & ", Decode(c.��Ŀ����, '" & !Ҫ������ & "', Nvl(c.δ��˵��,Decode(c.��¼����,Null,'" & !�����ı� & "'||'" & !Ҫ�ص�λ & "','" & !�����ı� & "'||c.��¼����||'" & !Ҫ�ص�λ & "')),  '" & !�����ı� & "'||'" & !Ҫ�ص�λ & "') As """ & !Ҫ������ & """"
                        mstrSQL���� = mstrSQL���� & " Or Decode(c.��Ŀ����, '" & !Ҫ������ & "', Nvl(c.δ��˵��,Decode(c.��¼����,Null,'" & !�����ı� & "'||'" & !Ҫ�ص�λ & "','" & !�����ı� & "'||c.��¼����||'" & !Ҫ�ص�λ & "')), '') Is Not Null"
                    End If
                Else
                    'Ϊ�ձ�ʾδ����,ǿ�Ƽ�,��������滻
                    mstrCOLNothing = mstrCOLNothing & "," & Val(Format(!�������, "00"))
                    mstrSQL�� = mstrSQL�� & ",Max(""" & "C" & Format(!�������, "00") & """) As C" & Format(!�������, "00")
                    mstrSQL���� = mstrSQL���� & " Or """ & "C" & Format(!�������, "00") & """ Is Not Null"
                    mstrSQL�� = mstrSQL�� & ", C" & Format(!�������, "00") & " AS C" & Format(!�������, "00")
                End If
            End Select
            .MoveNext
        Loop
        
        If mstrCollectItems <> "" Then
            mstrCollectItems = Mid(mstrCollectItems, 2)
            mstrColCollect = Mid(mstrColCollect, 2)
        End If
        mstrCOLNothing = Mid(mstrCOLNothing, 2)
        mstrCatercorner = Mid(mstrCatercorner, 2)
        mstrColWidth = Mid(mstrColWidth, 2)
        '�������һ�еĸ�ʽ
        mstrColumns = mstrColumns & IIf(mstrColumns = "", "", "'1'" & str��ʽ) '& "|" & !������� & "'" & !Ҫ������
        mstrColumns = Mid(mstrColumns, 2)     '��ʽ��:�к�;��Ŀ����1,��Ŀ����2|�к�...,ʵ��;1;����|2;����|3...
        If Mid(strSql��, 3) <> "" Then
            mstrSQL�� = mstrSQL�� & "," & Mid(strSql��, 3) & " As C" & Format(lngColumn, "00")
        Else
            mstrSQL�� = mstrSQL�� & ",'' As C" & Format(lngColumn, "00")
        End If
        
        If mstrSQL���� <> "" Then mstrSQL���� = "(" & Mid(mstrSQL����, 5) & ")"
        
        '���û�г������ڣ�ʱ�䣬��ʿ�����ڲ���Ҫ���䣬�Ա�֤�в�����������
        If bln���� = False Then mstrSQL�� = mstrSQL�� & ",To_Char(l.����ʱ��, 'yyyy-mm-dd') As ����"
        If blnʱ�� = False Then mstrSQL�� = mstrSQL�� & ",To_Char(l.����ʱ��, 'hh24:mi') As ʱ��"
        
        If blnǩ���� = False Then mstrSQL�� = mstrSQL�� & ",l.ǩ���� As ǩ����"
        If blnǩ��ʱ�� = False Then mstrSQL�� = mstrSQL�� & ",l.ǩ��ʱ��"
        
        If Mid(mstrSQL��, 2) = "" Then
            MsgBox "�Բ�����û�ж��嵱ǰ��������ʾ����Ϣ�����ڲ����ļ������ж��壡", vbInformation, gstrSysName
            Exit Function
        End If
        
        '�����ڲ��������ӹ̶���
        mstrSQL�� = UCase(mstrSQL�� & ",MAX(ǩ������) AS ǩ������,MAX(ǩ����Ϣ) AS ǩ����Ϣ,MAX(��¼ID) AS ��¼ID,MAX(����) AS ����,MAX(ʵ������) AS ʵ������,MAX(�������) AS �������,MAX(�����ı�) AS �����ı�,MAX(���ܱ��) AS ���ܱ��,MAX(��������) AS ��������,MAX(��ʼʱ��) AS ��ʼʱ��,MAX(����ʱ��) AS ����ʱ��")
        mstrSQL�� = UCase(mstrSQL�� & ",l.ǩ������,l.ǩ���� AS ǩ����Ϣ,C.��¼ID,P.����||'' AS ����,DECODE(SIGN(P.����ҳ��-P.��ʼҳ��),1,DECODE(SIGN([5]-P.��ʼҳ��),1, P.�����к�,P.����-P.�����к� ),P.����) AS ʵ������,0 AS �������,L.�����ı�,L.���ܱ��,to_char(L.����ʱ��,'yyyy-MM-dd hh24:mi:ss')||'' AS ��������,L.��ʼʱ��,L.����ʱ��")
        mstrSQL�� = UCase(mstrSQL�� & ",ǩ������,ǩ����Ϣ,��¼ID,����,ʵ������,�������,�����ı�,���ܱ��,��������,��ʼʱ��,����ʱ��")
        
        If bln��ʿ = False Then
            'ǿ����ӻ�ʿ��,Ϊ�˱����޸�����������(����¼�������,����������Ҳ������)
            mstrSQL�� = mstrSQL�� & ",��ʿ"
            mstrSQL�� = mstrSQL�� & ",l.������ As ��ʿ"
            mstrSQL�� = mstrSQL�� & ",��ʿ"
        End If
        
        '�����Ŀ���뵽SQL��
        Call PreActiveCOL
        Call SQLCombination
    End With
    
    ReadStruDef = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub PreActiveHead()
    Dim arrData
    Dim intCOl As Integer
    Dim strName As String
    Dim intDo As Integer, intCount As Integer
    '���±�ͷ
    
    arrData = Split(mstrCOLActive, "||")
    intCount = UBound(arrData)
    For intDo = 0 To intCount
        intCOl = Split(Split(arrData(intDo), "|")(0), ";")(0)
        strName = Split(Split(arrData(intDo), "|")(0), ";")(1)
        VsfData.TextMatrix(mintTabTiers - 1, intCOl + cHideCols + VsfData.FixedCols - 1) = strName
        If mintTabTiers = 3 And VsfData.TextMatrix(1, intCOl + cHideCols + VsfData.FixedCols - 1) = "" Then VsfData.TextMatrix(1, intCOl + cHideCols + VsfData.FixedCols - 1) = strName
        If mintTabTiers = 2 And VsfData.TextMatrix(0, intCOl + cHideCols + VsfData.FixedCols - 1) = "" Then VsfData.TextMatrix(0, intCOl + cHideCols + VsfData.FixedCols - 1) = strName
    Next
    
    With chkSwitch
        .Value = 0
        .Top = VsfData.Top + VsfData.Cell(flexcpTop, mintTabTiers - 1, mlngChoose) + VsfData.Cell(flexcpHeight, mintTabTiers - 1, mlngChoose) - .Height
        .Left = VsfData.Left + VsfData.Cell(flexcpLeft, mintTabTiers - 1, mlngChoose) + 50
        .Visible = mblnVerify
    End With
End Sub

Private Sub PreActiveCOL()
    Dim arrData
    Dim arrCol
    Dim intCOl As Integer, strName As String
    Dim strColFormat As String, strCOLNames As String, strCOLPart As String, strCOLCOND As String, strCOLDEF As String, strCOLMID As String, strCOLIN As String
    Dim intDo As Integer, intCount As Integer
    Dim intIn As Integer, intMax As Integer
    '�����Ŀ���뵽��ѯSQL�У���ʽ���к�;��ͷ����|��Ŀ���,��λ;��Ŀ���,��λ||�к�;��ͷ����...
    '�󶨶����Ŀ�����о��Զ�תΪ�Խ�����
    
    arrData = Split(mstrCOLActive, "||")
    intCount = UBound(arrData)
    For intDo = 0 To intCount
        intCOl = Split(Split(arrData(intDo), "|")(0), ";")(0)
        strName = Split(Split(arrData(intDo), "|")(0), ";")(1)
        arrCol = Split(Split(arrData(intDo), "|")(1), ";")
        
        '�����б�ʾ(ÿ������������Ŀ)
        strCOLPart = ""
        strCOLNames = ""
        strColFormat = ""
        strCOLCOND = ""
        strCOLMID = ""
        strCOLIN = ""
        strCOLDEF = ""
        intMax = UBound(arrCol)
        For intIn = 0 To intMax
            strCOLPart = Split(arrCol(intIn), ",")(1)
            mrsItems.Filter = "��Ŀ���=" & Val(Split(arrCol(intIn), ",")(0))
            strCOLNames = strCOLNames & "," & mrsItems!��Ŀ����
            strCOLCOND = strCOLCOND & " OR """ & strCOLPart & mrsItems!��Ŀ���� & """ IS NOT NULL"
            strCOLMID = strCOLMID & ",Max(""" & strCOLPart & mrsItems!��Ŀ���� & """) As """ & strCOLPart & mrsItems!��Ŀ���� & """"
            If intIn = 0 Then
                strCOLIN = strCOLIN & ", Decode(" & IIf(strCOLPart = "", "", "c.���²�λ||") & "c.��Ŀ����, '" & strCOLPart & mrsItems!��Ŀ���� & "', Nvl(c.δ��˵��,c.��¼����), '') As """ & strCOLPart & mrsItems!��Ŀ���� & """"
            Else
                strCOLIN = strCOLIN & ", Decode(" & IIf(strCOLPart = "", "", "c.���²�λ||") & "c.��Ŀ����, '" & strCOLPart & mrsItems!��Ŀ���� & "', Nvl(c.δ��˵��,Decode(c.��¼����,Null,'/','/'||c.��¼����||'')), '') As """ & strCOLPart & mrsItems!��Ŀ���� & """"
            End If
            If intIn = 0 Then
                If intMax = 0 Then
                    strCOLDEF = strCOLDEF & " """ & strCOLPart & mrsItems!��Ŀ���� & """ AS C" & intCOl
                Else
                    strCOLDEF = strCOLDEF & " """ & strCOLPart & mrsItems!��Ŀ���� & """||"
                End If
            Else
                strCOLDEF = strCOLDEF & "NVL(" & strCOLPart & mrsItems!��Ŀ���� & ",'/') AS C" & intCOl
            End If
            
            strColFormat = strColFormat & "{[" & strCOLPart & mrsItems!��Ŀ���� & "]" & IIf(intMax > 0 And intIn = 0, "/", "") & "}"
        Next
        If strCOLPart <> "" Then
            strCOLPart = Mid(strCOLPart, 2)
        End If
        strCOLNames = Mid(strCOLNames, 2)
        
        '�Խ���
        If intMax > 0 Then
            mstrCatercorner = mstrCatercorner & IIf(mstrCatercorner = "", "", ",") & intCOl
        End If
        '�и�ʽ:15'��ʿ'1'{[��ʿ]}
        mstrColumns = Replace(mstrColumns, intCOl & "''1'", intCOl & "'" & strCOLNames & "'1'" & strColFormat)
        '��
        mstrSQL�� = Replace(mstrSQL��, "'' AS C" & Format(intCOl, "00"), strCOLDEF)
        '����
        mstrSQL���� = Replace(UCase(mstrSQL����), " OR """ & "C" & Format(intCOl, "00") & """ IS NOT NULL", strCOLCOND)
        '��
        mstrSQL�� = Replace(mstrSQL��, ",MAX(""" & "C" & Format(intCOl, "00") & """) AS C" & Format(intCOl, "00"), strCOLMID)
        '��
        mstrSQL�� = Replace(mstrSQL��, ", C" & Format(intCOl, "00") & " AS C" & Format(intCOl, "00"), strCOLIN)
    Next
    mrsItems.Filter = 0
    
    '��δ�󶨵��е�SQL�������
    If mstrCOLNothing = "" Then Exit Sub
    arrData = Split(mstrCOLNothing, ",")
    intCount = UBound(arrData)
    For intDo = 0 To intCount
        '��(����Ҫ����)
'        mstrSQL�� = Replace(mstrSQL��, ",'' AS C" & arrData(intDo), "")
        '����
        mstrSQL���� = Replace(UCase(mstrSQL����), " OR """ & "C" & Format(arrData(intDo), "00") & """ IS NOT NULL", "")
        '��
        mstrSQL�� = Replace(mstrSQL��, ",MAX(""" & "C" & Format(arrData(intDo), "00") & """) AS C" & Format(arrData(intDo), "00"), "")
        '��
        mstrSQL�� = Replace(mstrSQL��, ", C" & Format(arrData(intDo), "00") & " AS C" & Format(arrData(intDo), "00"), "")
    Next
End Sub

Private Sub SQLCombination(Optional ByVal lng��¼ID As Long = 0)
    Dim str���� As String
    str���� = IIf(lng��¼ID = 0, "", " And ��¼ID=[7]")
    
    mstrSQL = "Select '' ����,to_char(����ʱ��,'yyyy-MM-dd hh24:mi:ss') AS ����ʱ��,'' AS ѡ��," & Mid(mstrSQL��, 12) & vbCrLf & _
                " From (Select ����ʱ��," & Mid(mstrSQL��, 2) & vbCrLf & _
                "        From (Select l.����ʱ��," & Mid(mstrSQL��, 2) & vbCrLf & _
                "               From ���˻������� l, ���˻�����ϸ c,���˻����ļ� f,���˻����ӡ p " & vbCrLf & _
                "               Where l.ID=p.��¼ID And l.Id = c.��¼id And l.�ļ�ID=f.ID And f.ID=p.�ļ�ID " & _
                "               And c.��ֹ�汾 Is Null And c.��¼����<>5  " & _
                "               And f.id=[1] And f.����id = [2] And f.��ҳid = [3] And Nvl(f.Ӥ��,0)=[4] And (P.��ʼҳ��=[5] Or P.����ҳ��=[5]) And l.�������=[6]" & IIf(mstrSQL���� <> "", " And (" & mstrSQL���� & ")", "") & ")" & vbCrLf & _
                IIf(str���� <> "", "Where " & str����, "") & _
                "       Group By ����, ʱ��, ����ʱ��,��ʿ,ǩ����,ǩ��ʱ��" & _
                                "       Order By ����ʱ��,��ʿ,ǩ����,ǩ��ʱ��)"
End Sub

Public Sub zlRefresh(Optional ByVal blnRefresh As Boolean = True)
'-----------------------------------------------------------------------
'���ܣ��������ˢ��,blnRefresh=false ��ʾֻˢ�±��Ϻͱ�ǩ������Ϣ
'-----------------------------------------------------------------------
    Dim aryRow() As String, aryItem() As String, arrItemEnd() As String
    Dim strPrefix As String, strItemName As String
    Dim lngRow As Long, lngCol As Long, lngCount As Long, strCell As String, i As Integer
    Dim strTmpSQL As String
    Dim aryPeriod() As String
    Dim strTmp As String, str��λ As String
    
    Err = 0: On Error GoTo errHand
    Debug.Print Now & "zlRefresh"
    Call InitCons
    '���ϱ�ǩ��ȡ
    lblSubhead.Caption = ""
    lblSubhead.Tag = ""
    lblSubEnd.Caption = ""
    lblSubEnd.Tag = ""
    aryPeriod = Split(GetPeriod, "��")
    gstrSQL = "Select [1] || Zl_Replace_Element_Value([2],[3],[4],2,NULL,[5]) as ��Ϣ From Dual"
    aryItem = Split(mstrSubHead, "|")
    arrItemEnd = Split(mstrSubEnd, "|")
    For i = 0 To 1
        For lngCount = 0 To IIf(i = 0, UBound(aryItem), UBound(arrItemEnd))
            If i = 0 Then
                strPrefix = Left(aryItem(lngCount), InStr(1, aryItem(lngCount), "{") - 1)
                strItemName = Mid(aryItem(lngCount), InStr(1, aryItem(lngCount), "{") + 1, InStr(1, aryItem(lngCount), "}") - InStr(1, aryItem(lngCount), "{") - 1)
            Else
                strPrefix = Left(arrItemEnd(lngCount), InStr(1, arrItemEnd(lngCount), "{") - 1)
                strItemName = Mid(arrItemEnd(lngCount), InStr(1, arrItemEnd(lngCount), "{") + 1, InStr(1, arrItemEnd(lngCount), "}") - InStr(1, arrItemEnd(lngCount), "{") - 1)
            End If
            mrsPartogram.Filter = 0
            mrsPartogram.Filter = "������='" & strItemName & "'"
            '�������Ҳ����������ֹ��޸�����
            If mrsPartogram.RecordCount = 0 Then GoTo ErrNext
            str��λ = Trim(NVL(mrsPartogram!��λ))
            If Val(NVL(mrsPartogram!�滻��)) = 1 Then
                '���̶̹�Ҫ����Ϣ
                strTmp = strPrefix
                Select Case strItemName
                Case "��ǰ����"
                
                    strTmpSQL = "Select   b.����" & vbNewLine & _
                                "From (Select ����id, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                                "            From ���˱䶯��¼" & vbNewLine & _
                                "            Where ����id = [1] And ��ҳid = [2] And ����id = [3]) a,���ű� b " & vbNewLine & _
                                "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [4] >= a.��ʼʱ��) And a.����id Is Not Null And b.ID=a.����id" & vbNewLine & _
                                "Order By a.��ʼʱ��"
                                
                    Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "��ǰ����", mlng����ID, mlng��ҳID, mlng����id, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
                    
                Case "��ǰ����"
                
                    strTmpSQL = "Select   a.����" & vbNewLine & _
                                "From (Select ����, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                                "            From ���˱䶯��¼" & vbNewLine & _
                                "            Where ����id = [1] And ��ҳid = [2] And ����id = [3]) a" & vbNewLine & _
                                "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [4] >= a.��ʼʱ��) And a.���� Is Not Null" & vbNewLine & _
                                "Order By a.��ʼʱ��"
        
                    Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "��ǰ����", mlng����ID, mlng��ҳID, mlng����id, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
                    If rsTemp.BOF = False Then rsTemp.MoveLast
                    
                Case "��ǰ����"
                
                    strTmpSQL = "Select   ���� From ���ű� a Where a.ID=[1]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "��ǰ����", mlng����id)
                    
                Case "סԺҽʦ"
                    strTmpSQL = "Select   a.����ҽʦ" & vbNewLine & _
                                "From (Select ����ҽʦ, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                                "            From ���˱䶯��¼" & vbNewLine & _
                                "            Where ����id = [1] And ��ҳid = [2] And ����id = [3]) a" & vbNewLine & _
                                "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [4] >= a.��ʼʱ��) And a.����ҽʦ Is Not Null" & vbNewLine & _
                                "Order By a.��ʼʱ��"
                    Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "סԺҽʦ", mlng����ID, mlng��ҳID, mlng����id, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
                    If rsTemp.BOF = False Then rsTemp.MoveLast
                Case "���λ�ʿ"
                
                    strTmpSQL = "Select   a.���λ�ʿ" & vbNewLine & _
                                "From (Select ���λ�ʿ, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                                "            From ���˱䶯��¼" & vbNewLine & _
                                "            Where ����id = [1] And ��ҳid = [2] And ����id = [3]) a" & vbNewLine & _
                                "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [4] >= a.��ʼʱ��) And a.���λ�ʿ Is Not Null" & vbNewLine & _
                                "Order By a.��ʼʱ��"
                    Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "���λ�ʿ", mlng����ID, mlng��ҳID, mlng����id, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
                    If rsTemp.BOF = False Then rsTemp.MoveLast
                    
                Case "����ȼ�"
                    strTmpSQL = "Select   b.����" & vbNewLine & _
                                "From (Select ����ȼ�ID, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                                "            From ���˱䶯��¼" & vbNewLine & _
                                "            Where ����id = [1] And ��ҳid = [2] And ����id = [3]) a,����ȼ� b" & vbNewLine & _
                                "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [4] >= a.��ʼʱ��) And a.����ȼ�ID Is Not Null And b.���=a.����ȼ�ID" & vbNewLine & _
                                "Order By a.��ʼʱ��"
                    Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "����ȼ�", mlng����ID, mlng��ҳID, mlng����id, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
                    If rsTemp.BOF = False Then rsTemp.MoveLast
                Case "������"
                    strTmp = ""
                    gstrSQL = "Select [1] || Zl_Replace_Element_Value([2],[3],[4],2,NULL,[5],[6]) as ��Ϣ From Dual"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡҪ��", strPrefix, strItemName, mlng����ID, mlng��ҳID, mintӤ��, CDate(aryPeriod(0)))
                Case Else
                    strTmp = ""
                    gstrSQL = "Select [1] || Zl_Replace_Element_Value([2],[3],[4],2,NULL,[5]) as ��Ϣ From Dual"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡҪ��", strPrefix, strItemName, mlng����ID, mlng��ҳID, mintӤ��)
                End Select
            Else
                '����¼��Ҫ����Ϣ
                strTmp = strPrefix
                gstrSQL = "SELECT ���� From ����Ҫ������" & _
                    "   Where �ļ�ID = [1] And Ӥ�� = [2] And ���� =[3]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȥҪ��", mlng�ļ�ID, mlngFileIndex, strItemName)
            End If
            If rsTemp.BOF = False Then
                If i = 0 Then
                    If strTmp <> "" Then
                        lblSubhead.Tag = lblSubhead.Tag & " " & strTmp & rsTemp.Fields(0).Value & str��λ
                    Else
                        lblSubhead.Tag = lblSubhead.Tag & " " & rsTemp.Fields(0).Value & str��λ
                    End If
                Else
                    If strTmp <> "" Then
                        lblSubEnd.Tag = lblSubEnd.Tag & " " & strTmp & rsTemp.Fields(0).Value & str��λ
                    Else
                        lblSubEnd.Tag = lblSubEnd.Tag & " " & rsTemp.Fields(0).Value & str��λ
                    End If
                End If
            Else
            If i = 0 Then
                If strTmp <> "" Then
                        lblSubhead.Tag = lblSubhead.Tag & " " & strTmp
                    Else
                        lblSubhead.Tag = lblSubhead.Tag & " "
                    End If
                Else
                    If strTmp <> "" Then
                        lblSubEnd.Tag = lblSubEnd.Tag & " " & strTmp
                    Else
                        lblSubEnd.Tag = lblSubEnd.Tag & " "
                    End If
                End If
            End If
ErrNext:
        Next
    Next i
    lblSubhead.Tag = Trim(lblSubhead.Tag)
    lblSubEnd.Tag = Trim(lblSubEnd.Tag)
    '���ϱ�ǩ��ɢ����
    Call zlLableBruit
    
    If Not blnRefresh Then Exit Sub
    '�����м�¼��
    Call InitRecords
    
    'װ������
    Call SQLCombination
    gstrSQL = mstrSQL
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������", mlng�ļ�ID, mlng����ID, mlng��ҳID, mintӤ��, mintҳ��, cboBaby.ItemData(cboBaby.ListIndex))
    '�����������¼���ṹ
    Call DataMap_Init(rsTemp)
    '�����ݲ����û����¼���ĸ�ʽ,ͬʱʵ��һ�����ݷ�����ʾ�Ĺ���
    Call PreTendFormat(rsTemp)
    
    lblCurPage.Caption = "P" & mintҳ��
    Exit Sub

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub DataMap_Init(ByVal rsSource As ADODB.Recordset)
    '��ʼ���ڴ����ݼ�
    
    If Not mblnClear Then Exit Sub
    
    '���ݼ�¼��,���ڿ��ٻָ�
    Set mrsDataMap = CopyNewRec(rsSource)
    mrsDataMap.Sort = "ҳ��,�к�"
    '�޸ĵ�Ԫ���¼,���ڱ���
    Call Record_Init(mrsCellMap, "ID," & adLongVarChar & ",50|ҳ��," & adDouble & ",18|�к�," & adDouble & ",18|" & _
            "�к�," & adDouble & ",18|��¼ID," & adDouble & ",18|����," & adLongVarChar & ",4000|��λ," & adLongVarChar & ",100|" & _
            "����," & adDouble & ",1|ɾ��," & adDouble & ",1")
    mrsCellMap.Sort = "ҳ��,�к�,�к�"
    '���Ƽ�¼��
    Set mrsCopyMap = New ADODB.Recordset
    Set mrsCopyMap = CopyNewRec(mrsDataMap, False)
    
    'Ϊ�˲�Ӱ��֮��Ļ�ҳ,���˲�������Ϊ��
    mblnClear = False
End Sub

Private Function DataMap_Save() As Boolean
    '����ǰҳ�����û��༭�������ݱ�������,ҳ���л��򱣴�ǰ����
    Dim blnExit As Boolean
    Dim lngRow As Long, lngCol As Long, lngRows As Long, lngCols As Long
    Dim strDate As String, strTime As String, strDatetime As String, strCurrDate As String
    On Error GoTo errHand
    
    '�����Ƿ�༭��������
'    '�����ǰҳδ�༭��,�򲻱ر���
'    mrsCellMap.Filter = "ҳ��=" & mintҳ��
'    blnExit = (mrsCellMap.RecordCount = 0)
'    If blnExit Then
'        mrsCellMap.Filter = 0
'        DataMap_Save = True
'        Exit Function
'    End If
'    mrsCellMap.Filter = 0
    If Not CheckFlip Then Exit Function
    
    '��ɾ��ָ��ҳ�ŵ�����������
    mrsDataMap.Filter = "ҳ��=" & mintҳ��
    Do While True
        If mrsDataMap.RecordCount = 0 Then Exit Do
        mrsDataMap.Delete
        mrsDataMap.MoveNext
    Loop
    mrsDataMap.Filter = 0
    
    strCurrDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm")
    '����ָ��ҳ�ŵ�����������
    lngRows = VsfData.Rows - 1
    lngCols = VsfData.Cols - 1
    For lngRow = VsfData.FixedRows To lngRows
        mrsDataMap.AddNew
        mrsDataMap!ҳ�� = mintҳ��
        mrsDataMap!�к� = lngRow
        mrsDataMap!ɾ�� = IIf(VsfData.RowHidden(lngRow), 1, 0)
        strDatetime = ""
        For lngCol = 0 To lngCols - VsfData.FixedCols
            If lngCol + VsfData.FixedCols = mlngChoose Then
                mrsDataMap.Fields(cControlFields + lngCol).Value = VsfData.Cell(flexcpChecked, lngRow, mlngChoose)
            ElseIf InStr(1, "," & mlngCollectType & "," & mlngRecord & ",", "," & lngCol + VsfData.FixedCols & ",") <> 0 Then
                mrsDataMap.Fields(cControlFields + lngCol).Value = Val(VsfData.TextMatrix(lngRow, lngCol + VsfData.FixedCols))
            Else
                mrsDataMap.Fields(cControlFields + lngCol).Value = IIf(VsfData.TextMatrix(lngRow, lngCol + VsfData.FixedCols) = "", Null, VsfData.TextMatrix(lngRow, lngCol + VsfData.FixedCols))
            End If
            
            If lngCol + VsfData.FixedCols = mlngDate Then
                  strDate = Trim(VsfData.TextMatrix(lngRow, lngCol + VsfData.FixedCols))
                  If strDate <> "" Then
                      If mblnDateAd Then
                          strDate = Mid(strCurrDate, 1, 5) & ToStandDate(strDate)
                      Else
                          strDate = Format(strDate, "yyyy-MM-dd")
                      End If
                  End If
            ElseIf lngCol + VsfData.FixedCols = mlngTime Then
                  strTime = Trim(VsfData.TextMatrix(lngRow, lngCol + VsfData.FixedCols))
            End If
        Next
        If strDate <> "" And strTime <> "" Then
            strDatetime = strDate & " " & strTime & ":00"
            If mblnDateAd Then
                strDatetime = GetDateAdCurrDate(strDatetime)
            End If
            strDatetime = Format(strDatetime, "YYYY-MM-DD HH:mm:ss")
        End If
        mrsDataMap!�������� = Format(strDatetime, "YYYY-MM-DD HH:mm:ss")
        mrsDataMap.Update
    Next
    
    DataMap_Save = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function DataMap_Restore(ByVal rsTemp As ADODB.Recordset) As Boolean
    '��ָ��ҳ������ݻָ��������
    Dim lngRow As Long, lngCol As Long, lngRows As Long, lngCols As Long
    On Error GoTo errHand
    
    mblnRestore = False
    VsfData.Cell(flexcpChecked, VsfData.FixedRows, mlngChoose, VsfData.Rows - 1, mlngChoose) = flexTSUnchecked
    '����ָ��ҳ�ŵ����������е������
    mrsDataMap.Filter = "ҳ��=" & mintҳ��
    lngRows = mrsDataMap.RecordCount
    
    If lngRows = 0 Then
        'û���޸Ĺ���������󶨶�ȡ�ļ�¼��
        mrsDataMap.Filter = 0
        Set VsfData.DataSource = rsTemp
        DataMap_Restore = True
        Exit Function
    Else
        Set VsfData.DataSource = rsTemp
        mblnRestore = True
    End If
    
    mrsDataMap.MoveFirst
    lngCols = VsfData.Cols - 1
    For lngRow = 0 To lngRows - 1
        If lngRow > VsfData.Rows - VsfData.FixedRows - 1 Then VsfData.Rows = VsfData.Rows + 1
        For lngCol = 0 To lngCols - VsfData.FixedCols
            If lngCol + VsfData.FixedCols = mlngChoose Then
                If InStr(1, "3,4", NVL(mrsDataMap.Fields(cControlFields + lngCol).Value, 0)) <> 0 Then
                    VsfData.Cell(flexcpChecked, VsfData.FixedRows + lngRow, lngCol + VsfData.FixedCols) = NVL(mrsDataMap.Fields(cControlFields + lngCol).Value)
                End If
            Else
                VsfData.TextMatrix(VsfData.FixedRows + lngRow, lngCol + VsfData.FixedCols) = NVL(mrsDataMap.Fields(cControlFields + lngCol).Value)
            End If
        Next
        If mrsDataMap!ɾ�� = 1 Then VsfData.RowHidden(VsfData.FixedRows + lngRow) = True
        
        mrsDataMap.MoveNext
    Next
    
    mrsDataMap.Filter = 0
    DataMap_Restore = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub CellMap_Update(ByVal lngStart As Long, ByVal lngDeff As Long)
    Dim lngPos As Long
    Dim intCOl As Integer
    
    '���µ�ǰҳ�����д�����ʼ�е��к�����
    With mrsCellMap
        If .RecordCount <> 0 Then .MoveLast
        If .BOF Then Exit Sub
        Do While Not .BOF
            If !ҳ�� = mintҳ�� And !�к� > lngStart Then
                intCOl = !�к�
                lngPos = .AbsolutePosition
                !�к� = !�к� + lngDeff
                !ID = mintҳ�� & "," & !�к� & "," & !�к�
                .Update
                .MoveFirst
                .Move lngPos - 2
            Else
                .MovePrevious
            End If
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
End Sub

Public Function CopyNewRec(ByVal rsSource As ADODB.Recordset, Optional ByVal blnAddPage As Boolean = True) As ADODB.Recordset
    'ֻ������¼���Ľṹ,ͬʱ����ҳ��,�к��ֶ�
    Dim rsTarget As New ADODB.Recordset
    Dim intFields As Integer
    
    Set rsTarget = New ADODB.Recordset
    With rsTarget
        If blnAddPage Then
            .Fields.Append "ҳ��", adDouble, 18
            .Fields.Append "�к�", adDouble, 18
        End If
        For intFields = 0 To rsSource.Fields.Count - 1
            If rsSource.Fields(intFields).Name = "��������" Then
                .Fields.Append rsSource.Fields(intFields).Name, adLongVarChar, 50, adFldIsNullable      '0:��ʾ����
            ElseIf rsSource.Fields(intFields).Type = 200 Then       '�����ʹ���Ϊ�ַ���
                .Fields.Append rsSource.Fields(intFields).Name, adLongVarChar, rsSource.Fields(intFields).DefinedSize, adFldIsNullable     '0:��ʾ����
            Else
                .Fields.Append rsSource.Fields(intFields).Name, IIf(rsSource.Fields(intFields).Type = adNumeric, adDouble, rsSource.Fields(intFields).Type), rsSource.Fields(intFields).DefinedSize, adFldIsNullable    '0:��ʾ����
            End If
        Next
        If blnAddPage Then
            .Fields.Append "ɾ��", adDouble, 1
            .Fields.Append "��������", adLongVarChar, 20, adFldIsNullable
        End If
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set CopyNewRec = rsTarget
End Function

Private Sub PreTendMutilRows()
    Dim lngRowCount As Long, lngRowCurrent As Long  '��ǰ��¼������,��ǰ��¼�ڱ�ҳ��ʵ������
    Dim lngCol As Long, lngMax As Long
    Dim lngRow As Long
    Dim str����ʱ�� As String, str����ʱ��_L As String, lngLastRow As Long
    Dim lngLiterrRow As Long
    Dim lngTestRow As Long
    Dim strDate As String
    Dim intCOl As Integer, intCols As Integer
    Dim rsData As New ADODB.Recordset
    On Error GoTo errHand
    
    Dim arrData
    Dim intData As Integer, intDatas As Integer
    '���һ����ʾ�����������ʾ(���ݵ�ǰ����ռ����������ӿհ��в�����������,Ȼ�������δ���ǰ�е�����)
    'ÿҳֻ��ʾʵ�ʵ�������,��'@��ȡ��ע�ͼ���
    lngRow = VsfData.FixedRows
    Do While True
        If lngRow > VsfData.Rows - 1 Then Exit Do
        If lngRow >= mlngPageRows + mlngOverrunRows + VsfData.FixedRows Then Exit Do
        If InStr(1, VsfData.TextMatrix(lngRow, mlngRowCount), "|") <> 0 Then Exit Do
        
        lngRowCount = Val(VsfData.TextMatrix(lngRow, mlngRowCount))
        '@ʵ��������
'        lngRowCurrent = Val(VsfData.TextMatrix(lngRow, mlngRowCurrent))
        
        str����ʱ�� = VsfData.TextMatrix(lngRow, mlngActTime)
        If Val(VsfData.TextMatrix(lngRow, mlngCollectType)) < 0 Then str����ʱ��_L = "" ': str����ʱ�� = ""
        If str����ʱ��_L <> "" And Mid(str����ʱ��_L, 1, 16) = Mid(str����ʱ��, 1, 16) Then
            '������ͬ��������ͬ���Ҳ��ǻ��������У���˵����Щ������һ�飬����lngDemo��
            VsfData.TextMatrix(lngRow, mlngDate) = ""
            VsfData.TextMatrix(lngRow, mlngTime) = ""
            VsfData.TextMatrix(lngRow, mlngDemo) = lngRow - lngLastRow + 1
            If lngRow - lngLastRow = 1 Then
                VsfData.TextMatrix(lngLastRow, mlngDemo) = 1
            End If
        Else
            lngLastRow = lngRow
        End If
        
        If lngRowCount > 1 Then
            '�����ӿ���
            VsfData.Rows = VsfData.Rows + lngRowCount - 1
            '�ӵ�ǰ�е���һ�п�ʼ��ÿ�е�λ��+�����ӵĿհ���������֤�����Ŀհ��дӵ�ǰ�е���һ�п�ʼ
            For intData = VsfData.Rows - lngRowCount To lngRow + 1 Step -1
                VsfData.RowPosition(intData) = intData + lngRowCount - 1
            Next
            
            'ѭ������ǰ������
            For lngCol = 0 To VsfData.Cols - 1
                If VsfData.ColHidden(lngCol) And lngCol <> mlngRowCount Then
                    'ѭ����ֵ
                    For intData = 2 To lngRowCount
                        VsfData.TextMatrix(lngRow + intData - 1, lngCol) = VsfData.TextMatrix(lngRow, lngCol)
                    Next
                ElseIf (lngCol < mlngNoEditor And lngCol <> mlngDate And lngCol <> mlngTime) Then
                    '׼����ֵ
                    With txtLength
                        .Width = VsfData.ColWidth(lngCol)
                        .Text = VsfData.TextMatrix(lngRow, lngCol)
                        .FontName = VsfData.CellFontName
                        .FontSize = VsfData.CellFontSize
                        .FontBold = VsfData.CellFontBold
                        .FontItalic = VsfData.CellFontItalic
                    End With
                    arrData = GetData(txtLength.Text)
                    intDatas = UBound(arrData)
                    
                    If intDatas > 0 Then
                        'ѭ����ֵ
                        If intDatas + 1 > lngRowCount Then intDatas = lngRowCount - 1
                        For intData = 0 To intDatas
                            If VsfData.Rows <= lngRow + intData Then VsfData.Rows = VsfData.Rows + 1
                            VsfData.TextMatrix(lngRow + intData, lngCol) = arrData(intData)
                        Next
                    End If
                ElseIf lngCol = mlngNoEditor Then
                    '����ֵ��Ϊ��1��ʼ,������4������,����4|1
                    For intData = 1 To lngRowCount
                        VsfData.TextMatrix(lngRow + intData - 1, mlngRowCount) = lngRowCount & "|" & intData
                    Next
                    '���һ����Ҫ��д���ǩ��
                    If mlngSignName > 0 Then VsfData.TextMatrix(lngRow + lngRowCount - 1, mlngSignName) = VsfData.TextMatrix(lngRow, mlngSignName)
                    If mlngSignTime > 0 Then VsfData.TextMatrix(lngRow + lngRowCount - 1, mlngSignTime) = VsfData.TextMatrix(lngRow, mlngSignTime)
                Else
                    
                End If
            Next
            '@ʵ��������
'            '�����ҳ��һ�е����ݲ�ȫ,���Ƚ��ü�¼��һ�е�������(����,ʱ��,ǩ��)��Ϣ���Ƶ�
'            If lngRow = VsfData.FixedRows And lngRowCount <> lngRowCurrent Then
'                '�̶�������ʾ����ʱ����ǩ����
'                lngMax = lngRowCount - lngRowCurrent
'                If mlngDate > -1 Then VsfData.TextMatrix(lngRow + lngMax, mlngDate) = VsfData.TextMatrix(lngRow, mlngDate)
'                If mlngTime > -1 Then VsfData.TextMatrix(lngRow + lngMax, mlngTime) = VsfData.TextMatrix(lngRow, mlngTime)
'                if mlngOperator <>-1 then VsfData.TextMatrix(lngRow + lngMax, mlngOperator) = VsfData.TextMatrix(lngRow, mlngOperator)
'                if mlngSignName <>-1 then VsfData.TextMatrix(lngRow + lngMax, mlngsignname) = VsfData.TextMatrix(lngRow, mlngsignname)
'                'ɾ���������
'                For lngCol = 1 To lngMax
'                    VsfData.RemoveItem lngRow
'                Next
'            End If
'            lngRow = lngRow + lngRowCurrent - 1 '���ϸü�¼�ڱ�ҳʵ�ʵ�����
            '@ʵ��������Ҫע���������д���
            lngRow = lngRow + lngRowCount - 1
        Else
            VsfData.TextMatrix(lngRow, mlngRowCount) = "1|1"
        End If
        lngRow = lngRow + 1
        str����ʱ��_L = str����ʱ��
    Loop
    If mblnRestore Then Exit Sub
    
    'Modified by zyb 2011-09-15
    '��������,ֻ������һҳ������ʾ���༭,��һҳ����ʾ��ҳ�ķ�������
    '��鵱ǰҳ���Ƿ�Ϊ��������,����ɾ����Щ��������(���������)
    '������������Ƿ�Ϊ��������,�����ȡ��һҳ,����ҳ�Ĳ��ַ���������װ��һ��
    'If Val(VsfData.TextMatrix(VsfData.Rows - 1, mlngDemo)) > 0 And VsfData.Rows - VsfData.FixedRows >= mlngPageRows Then
    lngLiterrRow = 0
    mlngLitterRows(mintҳ��) = 0
    If VsfData.Rows - VsfData.FixedRows >= mlngPageRows And Val(VsfData.TextMatrix(VsfData.Rows - 1, mlngRowCount)) = 1 Then
        
        intCols = VsfData.Cols - 1
        lngTestRow = VsfData.FixedRows
        strDate = Mid(VsfData.TextMatrix(VsfData.Rows - 1, mlngActTime), 1, 16)
        Call SQLCombination
        gstrSQL = mstrSQL
        Call SQLDIY(gstrSQL)
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������", mlng�ļ�ID, mlng����ID, mlng��ҳID, mintӤ��, mintҳ�� + 1, cboBaby.ItemData(cboBaby.ListIndex))
        If rsData.RecordCount > 0 Then
            Set vsTest.DataSource = rsData
            Do While True
                If lngTestRow > vsTest.Rows - 1 Then Exit Do
                If strDate <> Mid(vsTest.TextMatrix(lngTestRow, mlngActTime), 1, 16) Then Exit Do
                VsfData.Rows = VsfData.Rows + 1
                lngLiterrRow = lngLiterrRow + 1
                For intCOl = 0 To intCols
                    VsfData.TextMatrix(lngRow, intCOl) = vsTest.TextMatrix(lngTestRow, intCOl)
                Next
                VsfData.TextMatrix(lngRow, mlngDate) = ""
                VsfData.TextMatrix(lngRow, mlngTime) = ""
                VsfData.TextMatrix(lngRow, mlngRowCount) = "1|1"
                VsfData.TextMatrix(lngRow, mlngDemo) = lngRow - lngLastRow + 1
                
                lngRow = lngRow + 1
                lngTestRow = lngTestRow + 1
            Loop
        End If
        
        Set vsTest.DataSource = Nothing
        vsTest.Clear
        rsData.Close
        Set rsData = Nothing
    End If
    
    If VsfData.Rows > VsfData.FixedRows Then
        If Val(VsfData.TextMatrix(VsfData.FixedRows, mlngRowCount)) = 1 Then
        'If Val(VsfData.TextMatrix(VsfData.FixedRows, mlngDemo)) > 0 And mintҳ�� > 1 Then
            If mintҳ�� > 1 Then
                Call SQLCombination
                gstrSQL = mstrSQL
                Call SQLDIY(gstrSQL)
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������", mlng�ļ�ID, mlng����ID, mlng��ҳID, mintӤ��, mintҳ�� - 1, cboBaby.ItemData(cboBaby.ListIndex))
                If rsData.RecordCount > 0 Then
                    Set vsTest.DataSource = rsData
                    VsfData.Rows = VsfData.Rows + 1
                    If Mid(VsfData.TextMatrix(VsfData.FixedRows, mlngActTime), 1, 16) = Mid(vsTest.TextMatrix(vsTest.Rows - 1, mlngActTime), 1, 16) Then
                        VsfData.RemoveItem VsfData.FixedRows
                        mlngLitterRows(mintҳ��) = 1
                        Do While True
                            If Val(VsfData.TextMatrix(VsfData.FixedRows, mlngDemo)) <= 1 Then Exit Do
                            VsfData.RemoveItem VsfData.FixedRows
                            mlngLitterRows(mintҳ��) = Val(CStr(mlngLitterRows(mintҳ��))) + 1
                        Loop
                        If VsfData.Rows - 1 > VsfData.FixedRows Then
                            VsfData.RemoveItem VsfData.Rows - 1
                        End If
                    End If
                End If
                Set vsTest.DataSource = Nothing
                vsTest.Clear
                rsData.Close
                Set rsData = Nothing
            End If
        End If
    End If
    If Val(CStr(mlngLitterRows(mintҳ��))) - lngLiterrRow <= 0 Then
        mlngLitterRows(mintҳ��) = 0
    Else
        mlngLitterRows(mintҳ��) = Val(CStr(mlngLitterRows(mintҳ��))) - lngLiterrRow
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub PreTendFormat(ByVal rsTemp As ADODB.Recordset)
    Dim aryItem() As String
    Dim lngRow As Long, lngCol As Long, lngCount As Long, strCell As String
    On Error GoTo errHand
    
    '���û����¼���ĸ�ʽ
    With VsfData
        .Redraw = flexRDNone
        .Clear
        Call DataMap_Restore(rsTemp)
        
        '��ͷ��д
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeRow(2) = True
        
        '�����ڲ�����������
        .ColHidden(1) = True
        .ColHidden(2) = True
        .ColHidden(mlngChoose) = Not mblnVerify
        .ColHidden(mlngRowCurrent) = True
        .ColHidden(mlngRowCount) = True
        .ColHidden(mlngRecord) = True
        .ColHidden(mlngSigner) = True
        .ColHidden(mlngSignLevel) = True
        .ColHidden(mlngCollectStyle) = True
        .ColHidden(mlngCollectText) = True
        .ColHidden(mlngCollectType) = True
        .ColHidden(mlngCollectDay) = True
        .ColHidden(mlngCollectStart) = True
        .ColHidden(mlngCollectEnd) = True
        If mlngOperator = -1 Then .ColHidden(.Cols - 1) = True
        .ColWidth(0) = 250
        .ColWidth(mlngChoose) = 250      'ѡ����
        
        .FrozenCols = mlngTime
        .SheetBorder = &H40C0&
        
        '������ͷ
        aryItem = Split(mstrTabHead, "|")
        For lngCount = 0 To UBound(aryItem)
            strCell = aryItem(lngCount)
            lngRow = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            lngCol = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            .TextMatrix(lngRow, lngCol + cHideCols + .FixedCols - 1) = strCell
        Next
        
        '���ù̶��м�ѡ����
        .TextMatrix(0, 0) = " "
        .TextMatrix(1, 0) = " "
        .TextMatrix(2, 0) = " "
        .TextMatrix(0, mlngChoose) = " "
        .TextMatrix(1, mlngChoose) = " "
        .TextMatrix(2, mlngChoose) = " "
        Call PreActiveHead
        
        '�п�����
        Dim blnAlign As Boolean
        
        aryItem = Split(mstrColWidth, ",")
        For lngCount = cHideCols + .FixedCols To .Cols - 1
            If Not .ColHidden(lngCount) Then
                .ColWidth(lngCount) = BlowUp(Val(Split(aryItem(lngCount - cHideCols - .FixedCols), "`")(0)))
                If InStr(1, aryItem(lngCount - cHideCols - .FixedCols), "`") <> 0 Then
                    blnAlign = True
                    .ColAlignment(lngCount) = Val(Split(aryItem(lngCount - cHideCols - .FixedCols), "`")(1))
                End If
            End If
        Next
        
        '�̶��и�ʽΪ����
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        '�ٰ��кϲ�
        For lngCount = 0 To .Cols - 1
            .MergeCol(lngCount) = True
        Next
        
        If blnAlign = False Then
            '��Ϊ�����û���������ʾ�ж��뷽ʽ
            If .FixedRows < .Rows Then .Cell(flexcpAlignment, .FixedRows, 0, .Rows - 1, .Cols - 1) = flexAlignGeneralCenter
        End If
        For lngCount = 0 To .Rows - 1
            If .RowHeight(lngCount) <> .RowHeightMin Then .RowHeight(lngCount) = .RowHeightMin
        Next
        Select Case mintTabTiers
        Case 1
            .RowHidden(0) = False
            .RowHidden(1) = True
            .RowHidden(2) = True
        Case 2
            .RowHidden(0) = False
            .RowHidden(1) = False
            .RowHidden(2) = True
        Case 3
            .RowHidden(0) = False
            .RowHidden(1) = False
            .RowHidden(2) = False
        End Select
        For lngCount = 0 To .Cols - 1
            .MergeCol(lngCount) = True
        Next
        
        If .Rows = .FixedRows Then
            mlngOverrunRows = 0
        Else
            '�õ���һ�еĳ�����
            mlngOverrunRows = Val(.TextMatrix(3, mlngRowCount)) - Val(.TextMatrix(3, mlngRowCurrent))
            '�������һ�еĳ�����
            If .Rows - 1 <> 3 Then
                mlngOverrunRows = mlngOverrunRows + Val(.TextMatrix(.Rows - 1, mlngRowCount)) - Val(.TextMatrix(.Rows - 1, mlngRowCurrent))
            End If
        End If
        Call PreTendMutilRows
        Call FillPage
        
        Call WriteColor
        
        '���̶ܹ��е��и߲���ȷ��Ҫ�Զ�������
        .AutoResize = True
        .WordWrap = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 0, .Cols - 1
        .AutoResize = False
        '���ǹ̶��е��и�����Ϊ��С�и�
        For lngCount = .FixedRows To .Rows - 1
            .RowHeight(lngCount) = .RowHeightMin
        Next
        
        .Redraw = flexRDDirect
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub WriteColor()
    Dim blnTag As Boolean
    Dim lngCount As Long, lngCol As Long
    '����Ժ�ɫ��ʾ��ͬʱ������ʼ������ΪNoCheckBox������ͼ��
    With VsfData
        For lngCount = .FixedRows To .Rows - 1
            If .TextMatrix(lngCount, 2) <> "" And Val(.TextMatrix(lngCount, mlngCollectType)) = 0 Then
                '����Ժ�ɫ��ʾ
                blnTag = False
                If mintTagFormHour < mintTagToHour Then
                    blnTag = (Hour(.TextMatrix(lngCount, 2)) >= mintTagFormHour And Hour(.TextMatrix(lngCount, 2)) < mintTagToHour)
                Else
                    blnTag = (Hour(.TextMatrix(lngCount, 2)) >= mintTagFormHour Or Hour(.TextMatrix(lngCount, 2)) < mintTagToHour)
                End If
                If blnTag Then
                    Set .Cell(flexcpFont, lngCount, 0, lngCount, .Cols - 1) = mobjTagFont
                    .Cell(flexcpForeColor, lngCount, 0, lngCount, .Cols - 1) = mlngTagColor
                End If
            End If
            
            '���������,���Ϊ����ʾΪ��
            If Val(.TextMatrix(lngCount, mlngCollectType)) < 0 Then
                For lngCol = mlngTime + 1 To .Cols - 1
                    If .TextMatrix(lngCount, lngCol) = "0" Then .TextMatrix(lngCount, lngCol) = ""
                    If IsNumeric(.TextMatrix(lngCount, lngCol)) Then .TextMatrix(lngCount, lngCol) = IIf(lngCol Mod 2 = 1, " ", "") & .TextMatrix(lngCount, lngCol)
                Next
                .MergeRow(lngCount) = True
            Else
                .MergeRow(lngCount) = False
            End If
            
            '������ʼ������ΪNoCheckBox
            If Not VsfData.RowHidden(lngCount) Then
                If Not VsfData.TextMatrix(lngCount, mlngRowCount) Like "*|1" Then
                    VsfData.Cell(flexcpChecked, lngCount, mlngChoose) = flexNoCheckbox
                Else
                    If VsfData.Cell(flexcpChecked, lngCount, mlngChoose) <> flexTSChecked Then
                        VsfData.Cell(flexcpChecked, lngCount, mlngChoose) = flexTSUnchecked
                    End If
                    
                    '����ͼ��
                    If VsfData.TextMatrix(lngCount, mlngSigner) = "" Then
                        VsfData.Cell(flexcpPicture, lngCount, 0) = Nothing
                    Else
                        If InStr(1, VsfData.TextMatrix(lngCount, mlngSigner), "/") <> 0 Then
                            VsfData.Cell(flexcpPicture, lngCount, 0) = imgRow.ListImages(��ǩ).Picture
                        Else
                            VsfData.Cell(flexcpPicture, lngCount, 0) = imgRow.ListImages(ǩ��).Picture
                        End If
                    End If
                
                    '����С�����ʾ
                    If Val(VsfData.TextMatrix(lngCount, mlngCollectType)) <> 0 Then
                        VsfData.TextMatrix(lngCount, mlngDate) = VsfData.TextMatrix(lngCount, mlngCollectText)
                        VsfData.TextMatrix(lngCount, mlngTime) = VsfData.TextMatrix(lngCount, mlngCollectText)
                    End If
                End If
            End If
        Next
    End With
End Sub

Private Sub zlLableBruit()
    Dim lngScaleLeft As Long, lngScaleTop As Long, lngScaleRight As Long, lngScaleBottom As Long
    
    lblSubhead.Caption = lblSubhead.Tag
    lblSubEnd.Caption = lblSubEnd.Tag
    lblSubhead.Top = lblTitle.Top + lblTitle.Height + 120
    Call cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    VsfData.Move lngScaleLeft + 210, lblSubhead.Top + lblSubhead.Height + 45, lngScaleRight - lngScaleLeft - 210 * 2
    VsfData.Height = picMain.Height - VsfData.Top - lblSubEnd.Height - 45
    lblSubEnd.Move lblSubhead.Left, VsfData.Top + VsfData.Height + 45
End Sub

Private Sub GetFileProperty()
    '��ȡ�ļ�����
    Dim strEnd As String
    On Error GoTo errHand
    
    gstrSQL = " Select   ��ʼʱ��,����ʱ��,��ʽID,����ID,�鵵�� From ���˻����ļ� " & _
              " Where ����ID=[1] And ��ҳID=[2] And Ӥ��=[3] And ID=[4] And Rownum<2"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ļ�����", mlng����ID, mlng��ҳID, mintӤ��, mlng�ļ�ID)
    If rsTemp.RecordCount <> 0 Then
        mlng��ʽID = rsTemp!��ʽID
        mlng����id = rsTemp!����ID
        mblnArchive = (NVL(rsTemp!�鵵��) <> "")
        mstr��ʼʱ�� = Format(rsTemp!��ʼʱ��, "yyyy-MM-dd HH:mm")
        mstr����ʱ�� = Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm")
        mstrBeginTime = mstr��ʼʱ��
        strEnd = DateAdd("n", -1, CDate(Format(CDate(mstr��ʼʱ��) + 1, "yyyy-MM-dd HH:mm:ss")))
        If mstr����ʱ�� = "" Then
            mstr����ʱ�� = Format(strEnd, "YYYY-MM-DD HH:mm:ss")
        Else
            If (mstr����ʱ�� <> "" And CDate(mstr����ʱ��) > CDate(strEnd)) Then mstr����ʱ�� = Format(strEnd, "yyyy-MM-dd HH:mm:ss")
        End If
    End If
    
    '�ڶ����ļ�������ȡ��ʼʱ��
    If mlngFileIndex > 1 Then
        gstrSQL = "SELECT Max(B.����ʱ��) ����ʱ��" & vbNewLine & _
            "FROM ���˻����ļ� A,���˻������� B,���˻�����ϸ C,�����¼��Ŀ D" & vbNewLine & _
            "WHERE A.ID=B.�ļ�ID AND B.ID=C.��¼ID AND A.ID=[1] And ����ID=[2] And ��ҳID=[3] And Ӥ��=[4] AND B.�������<[5] AND C.��Ŀ���=D.��Ŀ���" & vbNewLine & _
            "AND NVL(D.��Ŀ����,'')='����' AND NVL(D.������Ŀ,1)=1 ORDER BY B.����ʱ��"
        Call SQLDIY(gstrSQL)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ļ�����", mlng�ļ�ID, mlng����ID, mlng��ҳID, mintӤ��, mlngFileIndex)
        If rsTemp.RecordCount <> 0 Then
            mstr��ʼʱ�� = DateAdd("n", 1, CDate(Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm")))
        End If
    End If
    '���ҳ��=-1,˵��ȱʡ��ʾ���һҳ
    If mintҳ�� = -1 Then
        gstrSQL = "Select MAX(����ҳ��) AS ҳ�� from ���˻����ӡ A,���˻������� B" & vbNewLine & _
            "WHERE A.�ļ�ID=B.�ļ�ID And A.��¼ID=B.ID" & vbNewLine & _
            "And B.�������=[1] And A.�ļ�ID=[2]"
        Call SQLDIY(gstrSQL)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡָ��ҳ������ݷ���ʱ�䷶Χ", mlngFileIndex, mlng�ļ�ID)
        mintҳ�� = NVL(rsTemp!ҳ��, 1)
        mint����ҳ = mintҳ��
    End If
    If mblnClear = True Then
        ReDim mlngLitterRows(0 To mint����ҳ + 1)
    Else
        ReDim Preserve mlngLitterRows(0 To mint����ҳ + 1)
    End If
    RaiseEvent AfterRowColChange("", False, mblnSign, mblnArchive)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitEnv()
    Dim rs As New ADODB.Recordset
    On Error GoTo errHand
    
    glngHours = Val(zlDatabase.GetPara("���ݲ�¼ʱ��", glngSys))
    mintCollectDef = Val(zlDatabase.GetPara("С��ȱʡ��ʽ", glngSys, 1255))
    '���ִ��ڵ����л����¼��Ŀ
    gstrSQL = " Select   ��Ŀ���,��Ŀ����,��Ŀ����,��Ŀ����,��Ŀ����,��ĿС��,��Ŀ��ʾ,��Ŀ��λ,��Ŀֵ��,����ȼ�,Ӧ�÷�ʽ" & _
              " From �����¼��Ŀ B" & _
              " Order by ��Ŀ���"
    Set mrsItems = zlDatabase.OpenSQLRecord(gstrSQL, "���ִ��ڵ����л����¼��Ŀ")
    '��ȡ���в���Ҫ����Ϣ
    gstrSQL = "Select ������,�滻��,����,����,С��,��λ,��ʾ��,��ֵ��,����" & vbNewLine & _
        "From (Select i.����id, i.����, i.������, nvl(i.�滻��,0) �滻��,i.����,i.����,i.С��,i.��λ,i.��ʾ��,i.��ֵ��,i.����" & vbNewLine & _
        "       From ����������Ŀ I, ������������ K" & vbNewLine & _
        "       Where k.Id = i.����id And k.���� In ('02', '03', '05', '06') And i.�滻�� = 1 And k.���� = 1" & vbNewLine & _
        "       Union" & vbNewLine & _
        "       Select i.����id, i.����, i.������, nvl(i.�滻��,0) �滻��,i.����,i.����,i.С��,i.��λ,i.��ʾ��,i.��ֵ��,i.����" & vbNewLine & _
        "       From ����������Ŀ I, ������������ K" & vbNewLine & _
        "       Where k.Id = i.����id And k.���� In ('04', '05') And k.���� = 2)" & vbNewLine & _
        "Order By ����id, ����, �滻��"

    Set mrsPartogram = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����Ҫ����Ϣ")
    'ȡ��ǰ����Ա�ļ���
    mintVerify = δ����
    mintVerify_Last = δ����
    gstrSQL = "select  Ƹ�μ���ְ�� from ��Ա�� p where ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mRichEPR", glngUserId)
    If Not rs.EOF Then
        mintVerify = NVL(rs("Ƹ�μ���ְ��"), δ����)
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitRecords()
    Dim i As Integer, j As Integer, k As Integer, l As Integer
    Dim lngCol As Long, lngOrder As Long, strName As String, intImmovable As Integer, strFormat As String
    Dim arrColumn, arrItem, strColumns As String
    Dim blnSet As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    strColumns = mstrColumns
    If Not mblnInit Then
        '��ʼ���ڴ��¼��(δ��Ӧ��Ŀ����Ϊ���Ŀ,�����о�Ϊ�̶���)
        strFields = "��," & adDouble & ",18|���," & adDouble & ",2|��Ŀ���," & adDouble & ",18|��Ŀ����," & adLongVarChar & ",20|�̶�," & adDouble & ",2|��ʽ," & adLongVarChar & ",2000"
        Call Record_Init(mrsSelItems, strFields)
        strFields = "��|���|��Ŀ���|��Ŀ����|�̶�|��ʽ"
    End If
    
    '�����ж���
    If Not mblnInit Then
        arrColumn = Split(strColumns, "|")
        j = UBound(arrColumn)
        For i = 0 To j
            lngCol = Split(arrColumn(i), "'")(0)
            arrItem = Split(Split(arrColumn(i), "'")(1), ",")
            blnSet = False   '����������Դ���ֵΪ׼'�����Ҳ�����Ŀ���ǻ��Ŀ
            If UBound(Split(arrColumn(i), "'")) > 1 Then
                blnSet = True
                intImmovable = Split(arrColumn(i), "'")(2)
            End If
            If UBound(Split(arrColumn(i), "'")) > 2 Then
                strFormat = Split(arrColumn(i), "'")(3)
            End If
            
            k = UBound(arrItem)
            For l = 0 To k
                strName = arrItem(l)
                mrsItems.Filter = "��Ŀ����='" & strName & "'"
                If mrsItems.RecordCount <> 0 Then
                    lngOrder = mrsItems!��Ŀ���
                    If Not blnSet Then intImmovable = 1   '�̶��������޸�
                    Select Case strName
                        Case "��������"
                            mlngSpread = i + cHideCols + VsfData.FixedCols
                        Case "��¶�ߵ�"
                            mlngJust = i + cHideCols + VsfData.FixedCols
                        Case "����"
                            mlngProduce = i + cHideCols + VsfData.FixedCols
                    End Select
                Else
                    lngOrder = 0
                    If Not blnSet Then intImmovable = 0
                    
                    '��¼������
                    Select Case strName
                    Case "����"
                        mlngDate = i + cHideCols + VsfData.FixedCols
                    Case "ʱ��"
                        mlngTime = i + cHideCols + VsfData.FixedCols
                    Case "��ʿ"
                        mlngOperator = i + cHideCols + VsfData.FixedCols
                    Case "ǩ����"
                        mlngSignName = i + cHideCols + VsfData.FixedCols
                    Case "ǩ��ʱ��"
                        mlngSignTime = i + cHideCols + VsfData.FixedCols
                    End Select
                End If
                strValues = lngCol & "|" & l + 1 & "|" & lngOrder & "|" & strName & "|" & intImmovable & "|" & strFormat
                Call Record_Add(mrsSelItems, strFields, strValues)
            Next
        Next
        
        'Call OutputRsData(mrsSelItems)
        
        '��������ڲ�������(�����ڶ�ȡ���ݺ��ʱ���ӵ�,��ʱֻ��Ԥ������)
        mlngDemo = 1
        mlngActTime = 2
        mlngChoose = 2 + VsfData.FixedCols
        mlngSignLevel = VsfData.Cols + cHideCols + VsfData.FixedCols '����������
        mlngSigner = mlngSignLevel + 1
        mlngRecord = mlngSigner + 1
        mlngRowCount = mlngRecord + 1
        mlngRowCurrent = mlngRowCount + 1
        mlngCollectType = mlngRowCurrent + 1
        mlngCollectText = mlngCollectType + 1
        mlngCollectStyle = mlngCollectText + 1
        mlngCollectDay = mlngCollectStyle + 1
        mlngCollectStart = mlngCollectDay + 1
        mlngCollectEnd = mlngCollectStart + 1
        
        If mlngOperator <> -1 And mlngSignName <> -1 Then
            mlngNoEditor = IIf(mlngOperator < mlngSignName, mlngOperator, mlngSignName)
        Else
            mlngNoEditor = IIf(mlngOperator <> -1, mlngOperator, mlngSignName)
        End If
    End If
    
    mrsItems.Filter = 0
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub SetArchiveValue(ByVal lngPatiID As Long, ByVal lngPageId As Long, ByVal intBaby As Integer)
    mlng����ID = lngPatiID
    mlng��ҳID = lngPageId
    mintӤ�� = intBaby
End Sub

Public Sub ArchiveMe()
    On Error GoTo errHand
    
    If mlng����ID = 0 Or gblnMoved Then Exit Sub
    If MsgBox("��Ҫ���ò��˱���סԺ���л����ļ��鵵��", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbYes Then
        Dim strNow As String

        strNow = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        gstrSQL = "ZL_���˻����ļ�_ARCHIVE(" & mlng����ID & "," & mlng��ҳID & "," & mintӤ�� & ",1)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�鵵")

        mblnArchive = True
        RaiseEvent AfterRowColChange("", False, mblnSign, mblnArchive)
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Sub UnArchiveMe()
    On Error GoTo errHand
    
    If mlng����ID = 0 Or gblnMoved Then Exit Sub
    If MsgBox("��Ҫȡ���ò��˵Ĺ鵵״̬��", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbYes Then

        gstrSQL = "ZL_���˻����ļ�_ARCHIVE(" & mlng����ID & "," & mlng��ҳID & "," & mintӤ�� & ",0)"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "�����鵵")
        
        mblnArchive = False
        RaiseEvent AfterRowColChange("", False, mblnSign, mblnArchive)
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function SignMe(Optional ByVal bln��ǩ As Boolean = False) As Boolean
    Dim blnSign As Boolean          '�Ƿ�ǩ���ɹ�
    Dim blnRefresh As Boolean
    Dim strSignTime As String       '��֤����ǩ����ǩ��ʱ��һ��,����ȡ��ǩ��ʱ��ǩ��ʱ��ͳһȡ��
    Dim str״̬ As String           '����ǩ��ѡ��,����ѭ��ǩ��ʱ��ͣ�ĵ���ǩ������
    Dim str�д��� As String
    Dim str���� As String
    Dim intRow As Integer, intRows As Integer
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '������ʱ��ѭ��������δǩ�����ݽ���ǩ��
    
    If mlng����ID = 0 Then Exit Function
    
    '��ǩ:������δǩ�������ݽ���ǩ��
    '��ǩ:��������ǩ�������ݽ�����ǩ
    If bln��ǩ Then
        If Not mblnVerify Then
            '��������ҲҪǩ��,���ȥ������: And B.�������=0
            gstrSQL = " Select  distinct B.����ʱ�� " & vbNewLine & _
                      " From ���˻�����ϸ A,���˻������� B,���˻����ļ� C" & vbNewLine & _
                      " Where A.��¼ID=B.ID And B.�ļ�ID=C.ID And A.������Դ=0 And MOD(A.��¼����,10)=5 AND A.��ֹ�汾 Is NULL And C.ID=[1] "
            Call SQLDIY(gstrSQL)
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ����δǩ��������", mlng�ļ�ID)
            If rsTemp.RecordCount = 0 Then
                RaiseEvent AfterRowColChange("��������ǩ�������ݣ�", True, mblnSign, mblnArchive)
                Exit Function
            End If
        
            '������ǩģʽ,���޸�����,�ɹ�ѡ����
            mblnVerify = True
            chkSwitch.Visible = mblnVerify
            VsfData.ColHidden(mlngChoose) = Not mblnVerify
            VsfData.Cell(flexcpChecked, VsfData.FixedRows, mlngChoose, VsfData.Rows - 1, mlngChoose) = flexTSUnchecked
            Call WriteColor
            RaiseEvent AfterDataChanged(mblnChange Or mblnVerify)
            Exit Function
        Else
            '��ȡ����ǩ������
            '��������ҲҪǩ��,���ȥ������: And B.�������=0
            gstrSQL = " Select /*+ RULE */ distinct B.����ʱ�� " & vbNewLine & _
                      " From ���˻�����ϸ A,���˻������� B,���˻����ļ� C,(SELECT COLUMN_VALUE FROM TABLE(CAST(F_NUM2LIST([2]) AS ZLTOOLS.T_NUMLIST))) G" & vbNewLine & _
                      " Where A.��¼ID=B.ID And B.ID=G.COLUMN_VALUE And B.�ļ�ID=C.ID And MOD(A.��¼����,10)=5 AND A.��ֹ�汾 Is NULL And C.ID=[1] "
            Call SQLDIY(gstrSQL)
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ����δǩ��������", mlng�ļ�ID, mstrVerify)
        End If
    Else
        '���Ա����޸ĵ����ݽ���ǩ��(��ȡδǩ������-��ǩ������)
        '��������ҲҪǩ��,���ȥ������: And B.�������=0
        gstrSQL = "" & _
                "SELECT  DISTINCT B.����ʱ��" & vbNewLine & _
                "FROM ���˻�����ϸ A,���˻������� B" & vbNewLine & _
                "WHERE A.��¼ID=B.ID And A.������Դ=0 AND A.��ֹ�汾 IS NULL AND A.��¼���� =1 AND instr(NVL(B.ǩ����,'QMR'),'/',1)=0 AND A.��¼��=[2] AND B.�ļ�ID=[1]" & vbNewLine & _
                "MINUS" & vbNewLine & _
                "SELECT DISTINCT B.����ʱ��" & vbNewLine & _
                "FROM ���˻�����ϸ A,���˻������� B" & vbNewLine & _
                "WHERE A.��¼ID=B.ID And A.������Դ=0 AND A.��ֹ�汾 IS NULL AND A.��¼���� =5 AND B.�ļ�ID=[1]"
        Call SQLDIY(gstrSQL)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ����δǩ��������", mlng�ļ�ID, gstrUserName)
        If rsTemp.RecordCount = 0 Then
            RaiseEvent AfterRowColChange("û���ҵ���Ҫǩ�������ݣ�ֻ�ܶ��Լ��Ǽǻ��޸ĵ����ݽ���ǩ������", True, mblnSign, mblnArchive)
            Exit Function
        End If
    End If
    
    '׼��ǩ��
    strSignTime = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    str�д��� = ""
    With rsTemp
        Do While Not .EOF
            blnSign = SignName(Format(!����ʱ��, "yyyy-MM-dd HH:mm:ss"), strSignTime, bln��ǩ, str״̬, str�д���)
            If Not blnSign Then Exit Do
            If Not blnRefresh Then blnRefresh = blnSign
'            If str�д��� <> "" Then
'                str���� = str���� & vbCrLf & "����ʱ��=[" & Format(!����ʱ��, "yyyy-MM-dd HH:mm:ss") & "]" & str�д���
'            End If
            .MoveNext
        Loop
    End With
    
    If blnRefresh And Not mblnVerify Then Call ShowMe(mFrmParent, mlng�ļ�ID, mlng����ID, mlng��ҳID, mlng����ID, mintӤ��, mstrPrivs, mblnEditable, mintҳ��)
    If str���� <> "" Then MsgBox "ǩ��ʱ�������´���" & str����, vbInformation, gstrSysName
    SignMe = blnRefresh
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub UnSignMe(Optional ByVal bln��ǩ As Boolean = False)
    Dim intPos As Integer
    Dim lngStart As Long                '��ʼ��
    Dim lngRecord As Long
    Dim blnOK As Boolean
    Dim strSignTime As String           'ǩ��ʱ��
    Dim blnClear As Boolean             'ȡ��ǩ��ʱ�Ƿ�����ð汾�����ݻ��˵��ϴ�ǩ�����״̬
    Dim blnTrans As Boolean
    Dim strSignName As String
    Dim clsSign As Object
    Dim rsTemp As New ADODB.Recordset
    Dim rsSign As New ADODB.Recordset
    Dim lngRecordCount As Long
    On Error GoTo errHand
    '�������һ���Ǳ��˵�ǩ�������ݵ�ǰѡ�����ݵ�ǩ��ʱ�䣬����ȡ��ǩ��
    
    If mlng����ID = 0 Then Exit Sub
    
    '��Ҫ�Լ��
    '��ǰ��¼���¼�¼���˳�
    If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "" Then Exit Sub
    lngStart = GetStartRow(VsfData.ROW)
    lngRecord = Val(VsfData.TextMatrix(lngStart, mlngRecord))
    If lngRecord = 0 Then
        RaiseEvent AfterRowColChange("������¼������ȡ��ǩ����", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    '��ǰ��¼δǩ�����˳�
    If VsfData.TextMatrix(lngStart, mlngSigner) = "" Then
        RaiseEvent AfterRowColChange("��ǰ��¼��δǩ����", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    '��ǩ����ǰ��¼δ��ǩ���˳���ƽǩ����ǰ��¼����ǩ���˳�
    intPos = InStr(1, VsfData.TextMatrix(lngStart, mlngSigner), "/")
    If bln��ǩ Then
        If intPos = 0 Then
            RaiseEvent AfterRowColChange("��ǰ��¼δ��ǩ���޷�ִ��ȡ����ǩ������", True, mblnSign, mblnArchive)
            Exit Sub
        End If
    Else
        If intPos <> 0 Then
            RaiseEvent AfterRowColChange("��ǰ��¼����ǩ����ȡ����ǩ���ٲ�����", True, mblnSign, mblnArchive)
            Exit Sub
        End If
    End If
    'ȡ��ǩ��ʱ�����Ի����Լ���ǩ�������Ի��˴�ǩ������
    '��������ҲҪǩ��,���ȥ������: And B.�������=0
    gstrSQL = "" & _
              " SELECT  A.��¼��,A.��¼ʱ��,A.��Ŀ���� AS ǩ��ʱ��,nvl(A.��ʼ�汾,1) ��ʼ�汾" & vbNewLine & _
              " FROM ���˻�����ϸ A,���˻������� B" & vbNewLine & _
              " WHERE A.��¼ID=B.ID AND B.�ļ�ID=[1] AND A.��¼ID=[2] AND A.��¼����=" & IIf(bln��ǩ, 15, 5) & vbNewLine & _
              " ORDER BY A.��Ŀ���� DESC"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ǰ��¼�����ǩ���˲��Ǳ������˳�", mlng�ļ�ID, lngRecord)
    If NVL(rsTemp!��¼��) <> gstrUserName Then
        strSignName = NVL(rsTemp!��¼��)
        '������Ǳ���ǩ��������Ƿ��Ǵ�ǩ
        gstrSQL = "" & _
              " SELECT A.��¼��" & vbNewLine & _
              " FROM ���˻�����ϸ A,���˻������� B" & vbNewLine & _
              " WHERE A.��¼ID=B.ID AND A.��¼����=1 AND B.�ļ�ID=[1] AND A.��¼ID=[2] And nvl(A.��ʼ�汾,1)=[3]"
        Call SQLDIY(gstrSQL)
        Set rsSign = zlDatabase.OpenSQLRecord(gstrSQL, "��ǰ��¼�����ǩ���˲��Ǳ������˳�", mlng�ļ�ID, lngRecord, Val(NVL(rsTemp!��ʼ�汾, 1)))
        lngRecordCount = rsSign.RecordCount
        rsSign.Filter = "��¼��='" & gstrUserName & "'"
        If rsSign.RecordCount = 0 And lngRecordCount > 0 Then
            RaiseEvent AfterRowColChange("���������ǩ����[" & strSignName & "]���ǩ��[" & gstrUserName & "]������ִ�б�������", True, mblnSign, mblnArchive)
            Exit Sub
        End If
    Else
        strSignName = gstrUserName
    End If
    
    '��ȡ��������׼��ȡ��ǩ������ǩ(��¼ʱ��<>""˵�����°�ǩ��)
    '��������ҲҪǩ��,���ȥ������: And B.�������=0
    If Not IsNull(rsTemp!��¼ʱ��) Then
        gstrSQL = "" & _
                  " SELECT  A.��ĿID AS ֤��ID,B.����ʱ��" & vbNewLine & _
                  " FROM ���˻�����ϸ A,���˻������� B" & vbNewLine & _
                  " WHERE A.��¼ID=B.ID AND B.�ļ�ID=[1] And A.��¼��=[2] And A.��¼ʱ��=[3] " & _
                  " AND A.��¼����=" & IIf(bln��ǩ, 15, 5)
        Call SQLDIY(gstrSQL)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������׼��ȡ��ǩ������ǩ", mlng�ļ�ID, strSignName, CDate(rsTemp!��¼ʱ��))
    Else
        gstrSQL = "" & _
                  " SELECT  A.��ĿID AS ֤��ID,B.����ʱ��" & vbNewLine & _
                  " FROM ���˻�����ϸ A,���˻������� B" & vbNewLine & _
                  " WHERE A.��¼ID=B.ID AND B.�ļ�ID=[1] And A.��¼��=[2] And A.��Ŀ����=[3] " & _
                  " AND A.��¼����=" & IIf(bln��ǩ, 15, 5) & vbNewLine & _
                  " ORDER BY A.��Ŀ���� DESC"
        Call SQLDIY(gstrSQL)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������׼��ȡ��ǩ������ǩ", mlng�ļ�ID, strSignName, CStr(rsTemp!ǩ��ʱ��))
    End If
    
    'ǩ���������޸ģ������޸ı������ǩ�������ȡ����ǩʱ��������ʾ�Ƿ�������ݵ����⣬��ǩ�Զ����ˣ�����ȡ����ʾ ѯ���Ƿ���Ҫ�������
'    If Not bln��ǩ Then
'        blnClear = (MsgBox("ȡ��ǩ��ʱ�Ƿ�ð汾�����ݻ��˵��ϴ�ǩ�����״̬��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
'    End If
    blnClear = True
    gcnOracle.BeginTrans
    blnTrans = True
    Do While Not rsTemp.EOF
        If NVL(rsTemp!֤��ID, 0) > 0 Then
            '����ǩ����֤��ֻ��֤һ��
            Err.Clear
            On Error Resume Next
            If clsSign Is Nothing Then
                Set clsSign = CreateObject("zl9ESign.clsESign")
                If Err <> 0 Then Err = 0
                
                If Not clsSign Is Nothing Then
                    If clsSign.Initialize(gcnOracle, glngSys) Then
                        If Not clsSign.CheckCertificate(gstrDBUser) Then
                            gcnOracle.RollbackTrans
                            Exit Sub
                        End If
                    Else
                        gcnOracle.RollbackTrans
                        RaiseEvent AfterRowColChange("ȡ��ǩ��ʱ��Ҫ�ٴ���֤����ϵͳû������ǩ����֤���ģ�����ȡ����", True, mblnSign, mblnArchive)
                        Exit Sub
                    End If
                Else
                    gcnOracle.RollbackTrans
                    RaiseEvent AfterRowColChange("ǩ��������ʼ��ʧ�ܣ�", True, mblnSign, mblnArchive)
                    Exit Sub
                End If
            End If
        End If
        
        'ȡ��ǩ��
        gstrSQL = "ZL_���˻�������_UNSIGNNAME("
        gstrSQL = gstrSQL & mlng�ļ�ID & ","
        gstrSQL = gstrSQL & "To_Date('" & Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss')," & IIf(blnClear, "1", "0") & ")"
        
        Debug.Print gstrSQL
        Call zlDatabase.ExecuteProcedure(gstrSQL, "ִ��ȡ��ǩ��")
        
        rsTemp.MoveNext
    Loop
    gcnOracle.CommitTrans
    blnTrans = False
    
    'ˢ������
    Call ShowMe(mFrmParent, mlng�ļ�ID, mlng����ID, mlng��ҳID, mlng����ID, mintӤ��, mstrPrivs, mblnEditable, mintҳ��)
    Exit Sub
errHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function SignName(ByVal strStart As String, ByVal strSignTime As String, ByVal bln��ǩ As Boolean, _
    str״̬ As String, Optional str���� As String) As Boolean
    '******************************************************************************************************************
    '����:
    '
    '
    '******************************************************************************************************************
    Dim oSign As cPartogramSign
    Dim strSource As String             '��ǩԴ���ݴ�
    Dim lngLoop As Long
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    '��ʼ����
    '------------------------------------------------------------------------------------------------------------------
    strSource = ""
    
    '��ȡҪǩ��������(��������ҲҪǩ��,���ȥ������: And B.�������=0)
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = " Select  a.��¼����,a.��Ŀ����,a.��Ŀ���,a.��Ŀ����,a.��Ŀ����,a.��¼����,a.��Ŀ��λ,a.��¼���,a.���²�λ,a.��¼���,a.���Ժϸ�,a.δ��˵��,a.��¼��,a.��¼ʱ��  " & _
              " From ���˻�����ϸ a,���˻������� b,���˻����ļ� c " & _
              " Where a.��¼id=b.ID And b.�ļ�ID=c.ID AND MOD(A.��¼����,10)<>5 And a.��ֹ�汾 Is Null And C.ID=[1] And b.����ʱ��=[2]" & _
              " Order by a.��Ŀ���"
    Call SQLDIY(gstrSQL)
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҪǩ��������", mlng�ļ�ID, CDate(strStart))
    If rs.BOF = False Then
        Do While Not rs.EOF
            For lngLoop = 0 To rs.Fields.Count - 1
                strSource = strSource & CStr(zlCommFun.NVL(rs.Fields(lngLoop).Value, ""))
            Next
            rs.MoveNext
        Loop
    End If
    Debug.Print "��ʼǩ����" & Now & vbCrLf & strSource
    If strSource = "" Then
        RaiseEvent AfterRowColChange("��ǰû����Ҫǩ������Ϣ��", True, mblnSign, mblnArchive)
        Exit Function
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    Err = 0
    '76223:������,2014-08-05,����ǩ������ʱ�����Ϣ
    Set oSign = frmPartogramSign.ShowMe(Me, mstrPrivs, mlng�ļ�ID, mlng����ID, mintVerify_Last, strSource, bln��ǩ, str״̬, str����)
    On Error GoTo errHand
    
    If Not oSign Is Nothing Then
        gstrSQL = "ZL_���˻�������_SIGNNAME("
        gstrSQL = gstrSQL & mlng�ļ�ID & ","
        gstrSQL = gstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss')," & IIf(bln��ǩ, 1, 0) & ","
        gstrSQL = gstrSQL & "'" & oSign.���� & "',"
        gstrSQL = gstrSQL & "'" & oSign.ǩ����Ϣ & "'," & oSign.ǩ������ & ","
        gstrSQL = gstrSQL & oSign.֤��ID & ","
        gstrSQL = gstrSQL & oSign.ǩ����ʽ & ",'" & oSign.ʱ��� & "',0,'" & oSign.ʱ�����Ϣ & "')"
        
        Debug.Print gstrSQL
        Call zlDatabase.ExecuteProcedure(gstrSQL, "ִ��ǩ��")
        SignName = True
    End If
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function CancelMe() As Boolean
    CancelMe = True
    mblnVerify = False
    mblnChange = False
    Call ShowMe(mFrmParent, mlng�ļ�ID, mlng����ID, mlng��ҳID, mlng����ID, mintӤ��, mstrPrivs, mblnEditable, mintҳ��)
End Function

Public Function SaveME() As Boolean
    If Not CheckData Then Exit Function
    If Not SaveData Then Exit Function
    
    mblnShow = False
    Call InitCons
    SaveME = True
    
    Call ShowMe(mFrmParent, mlng�ļ�ID, mlng����ID, mlng��ҳID, mlng����ID, mintӤ��, mstrPrivs, mblnEditable, mintҳ��)
End Function

Public Function ShowMe(ByVal frmParent As Form, ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, _
    ByVal lngDeptID As Long, ByVal intBaby As Integer, Optional ByVal strPrivs As String, Optional ByVal blnEditable As Boolean = True, _
    Optional ByVal intҳ�� As Integer = -1, Optional ByVal blnClear As Boolean = True, Optional ByVal bytSize As Byte = 0) As Boolean
    '******************************************************************************************************************
    '���ܣ� ��ʾ�����¼�ļ�����
    '������ frmParent           �ϼ��������
    '       lngPatiID           ����id
    '       lngPageID           ��ҳid
    '       lngDeptID           Ҫ��ʾ�����¼�Ŀ���
    '       intBaby             Ӥ����־
    '       blnEditable         ���Ϊ��,˵������Ϊ��ѯ�Ӵ�����ʹ��,ȡ����༭��صĹ���
    '       blnClear            ���Ϊ��,���mrsDataMap��¼��;����ҳʱӦ����,�����û��޸ĵ������Ա���ʾ������ʹ��
    '���أ� ��
    '******************************************************************************************************************
    Dim lngRow As Long
    Dim lngCount As Long, i As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    Err = 0
    
    mblnInit = False
    mblnChange = False
    mblnClass = False
    mstrClassRow = ""
    mblnClear = blnClear
    mintҳ�� = intҳ��
    mlng�ļ�ID = lngFileID
    mlng����ID = lngPatiID
    mlng��ҳID = lngPageId
    mlng����ID = lngDeptID
    mintӤ�� = intBaby
    mlngFileIndex = frmParent.FileNumIndex
    mstrPrivs = strPrivs
    mblnBlowup = (bytSize = 1) '(zlDatabase.GetPara("�����ļ���ʾģʽ", glngSys, 1255, 0) = 1)
    Set mFrmParent = frmParent
    
    mintPreDays = Val(zlDatabase.GetPara("����¼�뻤����������", glngSys, 1255, "1"))
    mstrMaxDate = Format(DateAdd("d", mintPreDays, zlDatabase.Currentdate), "yyyy-MM-dd")
    
    If mrsItems.State = 0 Then
        Call InitMenuBar
        Call InitEnv            '��ʼ������
        cbsThis.ActiveMenuBar.Visible = False
        cbsThis.RecalcLayout
    End If
    Call ReSetFontSize
    
    '��ȡ�����ļ�����
    lngCount = GetFileCount(mlng�ļ�ID, mlng����ID, mlng��ҳID)
    If mlngFileIndex < 1 Or mlngFileIndex > lngCount Then mlngFileIndex = 1
    With cboBaby
        .Tag = ""
        .Clear
        For i = 1 To lngCount
            .AddItem i: .ItemData(.NewIndex) = i
            If i = mlngFileIndex Then
                .ListIndex = i - 1
            End If
        Next i
        If .ListIndex = -1 Then .ListIndex = 0
        mlngFileIndex = .ItemData(.ListIndex)
    End With
    If mlngFileIndex <= 0 Then Exit Function
    
    mblnEditable = blnEditable And Not gblnMoved And Not mblnArchive
    
    RaiseEvent AfterDataChanged(mblnChange Or mblnVerify)
    RaiseEvent AfterRefresh
    
'    Call OutputRsData(mrsSelItems)
    ShowMe = True
    Exit Function
    
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:bytSize��0-С(ȱʡ)��1-��
    '����:������
    '����:2012-06-20 15:15:00
    '����:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtrl As Control
    Dim CtlFont As StdFont
    Dim bytFontSize As Byte
    bytFontSize = BlowUp(9)
    
    UserControl.FontSize = bytFontSize
    UserControl.FontName = "����"
    
    Set CtlFont = cbsThis.Options.Font
    If CtlFont Is Nothing Then
        Set CtlFont = UserControl.Font
    End If
    CtlFont.Size = bytFontSize
    Set cbsThis.Options.Font = CtlFont
    
    cboBaby.FontSize = bytFontSize
    picTmp.Height = cboBaby.Height
    vsfBaby.FontSize = bytFontSize
End Sub

Private Function CheckFlip() As Boolean
    Dim blnExit As Boolean
    Dim lngRow As Long, lngCol As Long, lngRows As Long, lngCols As Long
    'ҳ���л�ǰ��飺����ʱ����ȷ����������������ڱ���ʱ�Ͳ����ټ������ҳ��������ˣ�����������¼��ʱ�Ѿ������˼�飬�˴��Թ���
    
    lngRows = VsfData.Rows - 1
    lngCols = VsfData.Cols - 1
    For lngRow = VsfData.FixedRows To lngRows
        mrsCellMap.Filter = "ҳ��=" & mintҳ�� & " And �к�=" & lngRow & " And �к�>" & mlngTime
        If mrsCellMap.RecordCount <> 0 Then
            If Not VsfData.RowHidden(lngRow) And VsfData.TextMatrix(lngRow, mlngDemo) = "" Then
                blnExit = (VsfData.TextMatrix(lngRow, mlngDate) = "" Or VsfData.TextMatrix(lngRow, mlngTime) = "")
                If blnExit Then
                    mrsCellMap.Filter = 0
                    VsfData.ROW = lngRow
                    If Not VsfData.RowIsVisible(lngRow) Then VsfData.TopRow = lngRow
                    If mblnDate = False Then
                        lngCol = mlngTime
                    Else
                        lngCol = IIf(VsfData.TextMatrix(lngRow, mlngDate) = "", mlngDate, mlngTime)
                    End If
                    VsfData.COL = lngCol
                    If Not VsfData.ColIsVisible(lngCol) Then VsfData.LeftCol = lngCol
                    RaiseEvent AfterRowColChange("�벹������ʱ�䣡", True, mblnSign, mblnArchive)
                    CheckFlip = False
                    Exit Function
                End If
            End If
        End If
    Next
    
    mrsCellMap.Filter = 0
    CheckFlip = True
End Function

Private Function CheckProduce() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim blnExit As Boolean, blnProduce As Boolean
    Dim lngRow As Long, lngCol As Long
    Dim lngPage As Long, lngCount As Long
    Dim strDatetime As String
    
    '����ʱ���,1��Ӥ������ʱ������ڶ�Ӧ�Ĺ�����̥ͷ����,2.�������Ӥ���ļ���>1������ȡ����һ�ļ�Ӥ��������־
    '˵�����������DataMap_Save֮��ʹ��
    
    On Error GoTo errHand
    
    '--1��Ӥ������ʱ������ڶ�Ӧ�Ĺ�����̥ͷ����
    mrsDataMap.Filter = ""
    mrsDataMap.Filter = "ɾ��=0 And ��������<>''"
    mrsDataMap.Sort = "��������"
    'Call OutputRsData(mrsDataMap)
    Do While Not mrsDataMap.EOF
        If blnExit = False Then blnExit = (NVL(mrsDataMap.Fields(cControlFields + mlngSpread - VsfData.FixedCols).Value) <> "") And (NVL(mrsDataMap.Fields(cControlFields + mlngJust - VsfData.FixedCols).Value) <> "")
        If Mid(NVL(mrsDataMap.Fields(cControlFields + mlngProduce - VsfData.FixedCols).Value, 0), 1, 1) = "��" Then
            blnProduce = True
            strDatetime = Format(mrsDataMap!��������, "YYYY-MM-DD HH:mm:ss")
            lngPage = mrsDataMap!ҳ��
            lngRow = mrsDataMap!�к�
            If NVL(mrsDataMap.Fields(cControlFields + mlngSpread - VsfData.FixedCols).Value) = "" Then
                lngCol = mlngSpread
            Else
                lngCol = mlngJust
            End If
            GoTo ErrProduce
        End If
    mrsDataMap.MoveNext
    Loop
ErrProduce:
    If blnExit = False And blnProduce = True Then
        If mint����ҳ <> 1 Then '���ֻ��һҳ �����¼�����Ѿ����й����
            '��¼����û���ڴ����ݿ�����ȡ
            mstrSQL = "SELECT 1 FROM" & vbNewLine & _
                "���˻����ļ� A,���˻������� B,���˻�����ϸ C,�����¼��Ŀ D" & vbNewLine & _
                "WHERE A.ID=B.�ļ�ID AND B.ID=C.��¼ID AND A.ID=[1] AND A.����ID = [2] AND A.��ҳID = [3] AND B.������� = [4]" & vbNewLine & _
                "AND B.����ʱ�� BETWEEN [5] AND [6] AND INSTR('����������¶�ߵ�',D.��Ŀ����,1)<>0 AND D.������Ŀ=1" & vbNewLine & _
                "AND C.��Ŀ���=D.��Ŀ��� AND C.��¼���� IS NOT NULL " & vbNewLine & _
                "GROUP BY B.����ʱ�� HAVING COUNT(*)=2"
            Call SQLDIY(mstrSQL)
            Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "����ͼ���", mlng�ļ�ID, mlng����ID, mlng��ҳID, mlngFileIndex, CDate(mstr��ʼʱ��), CDate(strDatetime))
            lngCount = rsTemp.RecordCount
        Else
            lngCount = 0
        End If
        If lngCount = 0 Then
            lngRow = IIf(lngPage = mintҳ��, lngRow, VsfData.FixedRows)
            VsfData.ROW = lngRow
            VsfData.COL = lngCol
            If Not VsfData.RowIsVisible(lngRow) Then VsfData.TopRow = lngRow
            If Not VsfData.ColIsVisible(mlngSpread) Then VsfData.LeftCol = mlngSpread
            RaiseEvent AfterRowColChange("Ӥ����������ͬʱ���ڹ����������¶�½����ݣ����飡", True, mblnSign, mblnArchive)
            CheckProduce = False
            Exit Function
        End If
    End If
    
    '2.�������Ӥ���ļ���>1������ȡ����һ�ļ�Ӥ��������־
    If cboBaby.ListCount > 1 And cboBaby.ListIndex < cboBaby.ListCount - 1 Then
        If blnProduce = False Then '��ʾ������Ӥ���������
            RaiseEvent AfterRowColChange("��ǰӤ���ļ�С�����ļ���ʱ�����������Ӧ��������־�����飡", True, mblnSign, mblnArchive)
            CheckProduce = False
            Exit Function
        End If
    End If
    CheckProduce = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckData() As Boolean
    Dim intLevel As Integer
    Dim lngPage As Long
    On Error GoTo errHand
    '�������
    
    '����޸������ݶ�����ʱ�䲻ȫ����ʾ�����ݺϷ�����¼��ʱ�Ѿ���飩
    If Not DataMap_Save Then Exit Function
'    Call OutputRsData(mrsCellMap)
     'Call OutputRsData(mrsDataMap)

    If Not CheckProduce Then Exit Function
    
    '�������ǩģʽ,������ѡ�����Ƿ���ڲ�����ǩ�����
    If mblnVerify Then
        mstrVerify = ""
        mintVerify_Last = δ����
        '��ǩ��������������
        For lngPage = 1 To mint����ҳ
            mrsDataMap.Filter = "ҳ��=" & lngPage
            Do While Not mrsDataMap.EOF
                If NVL(mrsDataMap!ѡ��, 0) = flexTSChecked Then
                    mstrVerify = mstrVerify & "," & mrsDataMap!��¼ID
                    
                    If IsNull(mrsDataMap!ǩ������) Then
                        intLevel = NVL(mrsDataMap!ǩ������, δ����)
                    Else
                        intLevel = Val(mrsDataMap!ǩ������) + 1
                    End If
                    If mintVerify < intLevel Then mintVerify_Last = intLevel
                End If
                mrsDataMap.MoveNext
            Loop
        Next
        mrsDataMap.Filter = 0
        
        If mstrVerify = "" Then
            RaiseEvent AfterRowColChange("����Ҫѡ��һ�����ݲ��������ǩ������", True, mblnSign, mblnArchive)
            Exit Function
        End If
        mstrVerify = Mid(mstrVerify, 2)
    End If
    
    CheckData = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function SaveData() As Boolean
    Dim arrValue, arrOrder, arrPart, arrCollect
    Dim strSQL() As String
    Dim intAllow As Integer
    Dim lngRecord As Long
    Dim blnTrans As Boolean, blnSaved As Boolean, blnDel As Boolean
    Dim intPos As Integer, intMax As Integer, intPage As Integer, intRow As Integer, intUsedRows As Integer
    Dim strReturn As String, strCellData As String, strPart As String
    Dim strMonth As String, strDay As String
    Dim strDate As String, strTime As String, strTemp As String
    Dim strDatetime As String, strCurrDate As String, strDays As String, strLastDate As String
    
    ReDim Preserve strSQL(1 To 1)
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    'ͬ�ж���ѭ�����ã�ZL_���˻�������_UPDATE
    '��һ��ǰ���ã�
    '   1��ZL_���˻�������_SYNCHRO��ͬ�����ݵ����µ��뻤���¼���У���Ҫ��¼ɾ������ϸID��
    '   2��ZL_���˻����ӡ_UPDATE����ɴ�ӡ���ݽ���
    'ɾ����Ŀ���¼��ɾ����Ҳ��Ҫ��¼
    '�޸����ݵ�ͬ���ͽ��������ݶ�Ӧ��������ʱ�䱣�浽mrsCellMap��
    
'    objStream.WriteLine (Now & "��������SQL")
    intAllow = IIf(InStr(mstrPrivs, "���˻����¼") > 0, 1, 0)
    strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    With mrsCellMap
        '����Ч���ݹ��˳���:��¼ID>0����ʷ����+��������Ч����
        .Filter = "��¼ID>0 or (��¼ID=0 And ɾ��=0)"
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If intRow <> !�к� Then
endWork:
                If intRow > 0 Then
                    mrsDataMap.Filter = "ҳ��=" & intPage & " And �к�=" & intRow
                    If mrsDataMap.RecordCount <> 0 Then
                        blnDel = (mrsDataMap!ɾ�� = 1)
                        intUsedRows = Val(Split(NVL(mrsDataMap!���� & "|"), "|")(0))
                    Else
                        mrsDataMap.Filter = 0
                        intUsedRows = 1
                        RaiseEvent AfterRowColChange("��" & intRow & "�е������ڲ��������¼���β������貢������Ȼ������¼�����ݣ�лл��", True, mblnSign, mblnArchive)
                        Exit Function
                    End If
                    mrsDataMap.Filter = 0
                End If

                If blnSaved Then
                    '��ɴ�ӡ���ݽ���
'                    �ļ�ID_IN IN ���˻����ӡ.�ļ�ID%TYPE,
'                    ����ʱ��_IN IN ���˻����ӡ.����ʱ��%TYPE,
'                    ����_IN IN ���˻����ӡ.����%TYPE,
'                    ɾ��_IN Number:=0
'                    Ӥ��_IN IN ���˻�������.�������%TYPE
                    gstrSQL = "ZL_����ͼ���ݴ�ӡ_Update(" & mlng�ļ�ID & ",to_date('" & strDatetime & "','yyyy-MM-dd hh24:mi:ss')," & intUsedRows & "," & IIf(blnDel, "1", "0") & "," & mlngFileIndex & ")"
                    strSQL(ReDimArray(strSQL)) = gstrSQL
                    
                    blnSaved = False
                    If .EOF Then Exit Do
                End If
                
                '����ֵ
                intPage = !ҳ��
                intRow = !�к�
                strDate = ""
                strDatetime = ""
                lngRecord = NVL(!��¼ID, 0)
            End If
            
            If !�к� = mlngDate Then
                strDate = NVL(!����)
                If strDate <> "" Then
                    If mblnDateAd Then
                        strDate = Mid(strCurrDate, 1, 5) & ToStandDate(strDate)
                    Else
                        strDate = Format(strDate, "yyyy-MM-dd")
                    End If
                End If
            ElseIf !�к� = mlngTime Then
                strTime = NVL(!����)
                If strDate = "" Then strDate = Mid(strCurrDate, 1, 10)
                strDatetime = strDate & " " & strTime & ":00"
                If mblnDateAd Then
                    strDatetime = GetDateAdCurrDate(strDatetime)
                End If
                '����������ݣ�����ʱ����ͨ����������ֻ������+
'                If Val(NVL(!��λ)) >= 1 Then
'                    strDatetime = Mid(strDatetime, 1, 17) & String(2 - Len(!��λ), "0") & Val(!��λ) - 1
'                End If
                
                If lngRecord <> 0 Then
                    '���·���ʱ��
                    gstrSQL = "ZL_����ͼ����_����ʱ��(" & lngRecord & ",to_date('" & strDatetime & "','yyyy-MM-dd hh24:mi:ss')," & mlngFileIndex & ")"
                    strSQL(ReDimArray(strSQL)) = gstrSQL
                    blnSaved = True
                End If
            Else
                If !�к� > mlngTime Then
                    'ȡָ����Ԫ�������
                    strCellData = NVL(!����)
                    strPart = NVL(!��λ)
                    strReturn = ShowInput(!�к�, strCellData, True)
                    'strOrders��ʽ����Ŀ���,��Ŀ���...
                    'strValues��ʽ��ֵ'ֵ'ֵ...
                    arrOrder = Split(Split(strReturn, "||")(0), ",")
                    arrValue = Split(Split(strReturn, "||")(1) & "'", "'")
                    arrPart = Split(strPart & "/////", "/")
                    
                    intMax = UBound(arrOrder)
                    For intPos = 0 To intMax
    '                    �ļ�ID_IN IN ���˻�������.�ļ�ID%TYPE,
    '                    ����ʱ��_IN IN ���˻�������.����ʱ��%TYPE,
    '                    ��¼����_IN IN ���˻�����ϸ.��¼����%TYPE,          --������Ŀ=1���ϱ�˵��=2�������ձ��=4��ǩ����¼=5���±�˵��=6�����������=9
    '                    ��Ŀ���_IN IN ���˻�����ϸ.��Ŀ���%TYPE,          --������Ŀ����ţ��ǻ�����Ŀ�̶�Ϊ0
    '                    ��¼����_IN IN ���˻�����ϸ.��¼����%TYPE := NULL,  --��¼���ݣ��������Ϊ�գ��������ǰ�����ݣ�37��38/37
    '                    ���²�λ_IN IN ���˻�����ϸ.���²�λ%TYPE := NULL,
    '                    ��¼���_IN IN ���˻�������.�������%TYPE := 1,             --����Ǹ�Ӥ��������
    '                    ���˼�¼_IN IN NUMBER := 1,
                        gstrSQL = "ZL_����ͼ����_UPDATE(" & mlng�ļ�ID & ",to_date('" & strDatetime & "','yyyy-MM-dd hh24:mi:ss'),1," & _
                                arrOrder(intPos) & ",'" & arrValue(intPos) & "','" & arrPart(intPos) & "'," & mlngFileIndex & "," & intAllow & ",0," & IIf(mblnVerify, 1, 0) & ")"
                        strSQL(ReDimArray(strSQL)) = gstrSQL
                        blnSaved = True
                    Next
                    mrsItems.Filter = 0
                End If
            End If
            
            .MoveNext
        Loop
        
        If blnSaved Then GoTo endWork
        mrsDataMap.Filter = 0
    End With

    'ѭ��ִ��SQL��������
    'On Error Resume Next
    intMax = UBound(strSQL)
    
    gcnOracle.BeginTrans
    blnTrans = True
    
    On Error GoTo errHand
    If intMax > 0 Then
'        objStream.WriteLine (Now & "׼����������")
        For intPos = 1 To intMax
            If strSQL(intPos) <> "" Then
                Debug.Print strSQL(intPos)
    '            objStream.WriteLine (Now & "��SQL��" & strSQL(intPos))
                Call zlDatabase.ExecuteProcedure(strSQL(intPos), "���滤���¼������")
            End If
        Next
    '    objStream.WriteLine (Now & "�����������")
    End If
    If mblnVerify Then
        If Not SignMe(mblnVerify) Then
            gcnOracle.RollbackTrans
            Exit Function
        End If
    End If
    
    gcnOracle.CommitTrans
    SaveData = True
    blnTrans = False
    mblnChange = False
    mblnVerify = False
    
    RaiseEvent AfterDataChanged(mblnChange Or mblnVerify)
    RaiseEvent AfterRefresh
    RaiseEvent AfterDataSave(True)
    Exit Function
errHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cboBaby_Click()
    Dim i As Integer
    If Val(cboBaby.Tag) = cboBaby.ItemData(cboBaby.ListIndex) Then Exit Sub
    mblnInit = False
    If mblnChange Then
        If MsgBox("��ǰ���˵����ݻ�δ���棬�㡰�ǡ��ֹ����б��棬�㡰�񡱽����������޸ģ�", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            For i = 0 To cboBaby.ListCount - 1
                If cboBaby.ItemData(i) = mlngFileIndex Then
                    Call zlControl.CboSetIndex(cboBaby.Hwnd, i)
                    Exit For
                End If
            Next i
            Call InitCons
            Exit Sub
        Else
            mblnChange = False
        End If
    End If
    mintҳ�� = -1
    mblnClear = True
    mblnClass = False
    mstrClassRow = ""
    mlngFileIndex = cboBaby.ItemData(cboBaby.ListIndex)
    cboBaby.Tag = mlngFileIndex
    Debug.Print Now & "Begin"
    Call InitVariable
    Call InitCons
    If Not ReadStruDef Then Exit Sub
    Call zlRefresh
    Debug.Print Now & "Over"
    mblnInit = True
    RaiseEvent AfterFileIndex(mlngFileIndex)
    RaiseEvent AfterDataChanged(mblnChange)
End Sub

Private Sub cboChoose_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Index = 0 Then
            cboChoose(1).SetFocus
        Else
            Call MoveNextCell
        End If
    End If
End Sub

Private Sub cboС��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txt��ʼʱ��.Enabled Then
            txt��ʼʱ��.SetFocus
        Else
            txtС������.SetFocus
        End If
    End If
End Sub


Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strDate As String, strTime As String
    Dim strLockItem As String                   'ͬ������������,�������޸Ļ�ɾ��
    Dim lngTop As Long, lngHeight As Long
    Dim intMax As Integer                       'ͬ������������ռ�õ��������
    Dim intNULL As Integer, lngStartRow As Long
    Dim lngRow As Long, lngCol As Long, lngRows As Long, lngCols As Long
    Dim strKey As String, strField As String, strValue As String
    On Error GoTo err_exit
    
    mblnClass = False
    Select Case Control.ID
    'ճ��,���ʱ��Ҫͬ��mrsCellMap����
    Case conMenu_Edit_FileMan
        '�ļ����
        Call LoadBabyNum
    Case conMenu_Edit_Copy
        '����ָ�������е�����
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then Exit Sub
        lngRow = GetStartRow(VsfData.ROW)
        If Val(VsfData.TextMatrix(lngRow, mlngCollectType)) <> 0 Then Exit Sub
        
        '���Ƽ�¼��
        Set mrsCopyMap = New ADODB.Recordset
        Set mrsCopyMap = CopyNewRec(mrsDataMap, False)
        
        '�õ�ָ�������е���ʼ��,������
        lngCols = VsfData.Cols - 1
        lngRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
        lngRows = lngRow + lngRows - 1
        For lngRow = lngRow To lngRows
            mrsCopyMap.AddNew
            mrsCopyMap!ҳ�� = mintҳ��
            mrsCopyMap!�к� = lngRow
            For lngCol = 0 To lngCols - VsfData.FixedCols    '����һ���̶���
                mrsCopyMap.Fields(cControlFields + lngCol).Value = IIf(VsfData.TextMatrix(lngRow, lngCol + VsfData.FixedCols) = "", Null, VsfData.TextMatrix(lngRow, lngCol + VsfData.FixedCols))
            Next
            mrsCopyMap.Update
        Next
    Case conMenu_Edit_PASTE
        'ճ��ʱ����Ŀ�������帲�ǣ�ͬ�������������У���г���
        '���Ŀ���ܲ�ͬҳ����Ŀ��ͬ����λ��ͬ�����Բ����ǻ��Ŀ
        'ͬ������ռ�õ��������䣬�粻������ӿհ��У�����ճ��
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub
        If VsfData.TextMatrix(VsfData.ROW, mlngDemo) <> "" Then Exit Sub
        If mrsCopyMap.RecordCount = 0 Then Exit Sub
        
        '��ҳ�����в���������н���ճ��,ɾ��,ֻ�ܱ༭�����Ŀ�����
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
            lngStartRow = VsfData.ROW
        Else
            lngStartRow = GetStartRow(VsfData.ROW)
        End If
        If InStr(1, VsfData.TextMatrix(lngStartRow, mlngRowCount), "|") <> 0 And lngStartRow = 3 Then
            If Val(Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0)) <> Val(VsfData.TextMatrix(lngStartRow, mlngRowCurrent)) Then
                RaiseEvent AfterRowColChange("��ҳ�����в�����ճ�������л�����һҳ���в�����", True, mblnSign, mblnArchive)
                Exit Sub
            End If
        End If
        
        '���Ŀ���������Ƿ����ͬ������������,�����������ͬ���ļ�¼
        strLockItem = GetSynItems(2, intMax)        '1.������Ŀ���;2.�����к�
        
        '�õ�Ŀ�������е���ʼ��,������
        strField = "ID|ҳ��|�к�|�к�|��¼ID|����|ɾ��"
        lngCols = VsfData.Cols - 1
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
            lngRow = VsfData.ROW
            If Val(VsfData.TextMatrix(lngRow, mlngCollectType)) <> 0 Then Exit Sub
            lngStartRow = lngRow
            If mlngDate > -1 Then strDate = VsfData.TextMatrix(lngRow, mlngDate)
            strTime = VsfData.TextMatrix(lngRow, mlngTime)
        Else
            'ɾ�������������,����һ��
            lngRow = GetStartRow(VsfData.ROW)
            If Val(VsfData.TextMatrix(lngRow, mlngCollectType)) <> 0 Then Exit Sub
            lngStartRow = lngRow
            strDate = VsfData.TextMatrix(lngRow, mlngDate)
            strTime = VsfData.TextMatrix(lngRow, mlngTime)
            lngRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0)) - 1
            For intNULL = 1 To lngRows
                VsfData.RemoveItem lngRow + 1
            Next
        End If
        
        '������������,�������������������������ӵ�����
        intNULL = mrsCopyMap.RecordCount - 1
        For lngRow = 1 To mrsCopyMap.RecordCount - 1
            '��֤��ǰ�����������һҳ����ʾȫ
            If lngRow + VsfData.ROW > VsfData.Rows - 1 Then Exit For
            
            If Val(VsfData.TextMatrix(lngRow + VsfData.ROW, mlngRecord)) = 0 And VsfData.TextMatrix(lngRow + VsfData.ROW, mlngRowCount) = "" Then
                intNULL = intNULL - 1
            Else
                Exit For
            End If
        Next
        '�����ӿ���
        If intNULL > 0 Then
            VsfData.Rows = VsfData.Rows + intNULL
            '�ӵ�ǰ�м�¼�Ŀհ��п�ʼ��ÿ�е�λ��+�����ӵĿհ�����
            For lngRow = 1 To intNULL
                VsfData.RowPosition(VsfData.Rows - 1) = lngStartRow + 1
            Next
        End If
        
        '��ԭ���ڣ�ʱ�䣬ǿ�Ʋ������޸�
        VsfData.TextMatrix(lngStartRow, mlngDate) = strDate
        VsfData.TextMatrix(lngStartRow, mlngTime) = strTime
        '��¼�û��޸Ĺ��ĵ�Ԫ��
        If mlngDate <> -1 Then
            strKey = mintҳ�� & "," & lngStartRow & "," & mlngDate
            strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & mlngDate & "|" & _
                Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & VsfData.TextMatrix(lngStartRow, mlngDate) & "|0"
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        End If
        '2\ʱ��
        strKey = mintҳ�� & "," & lngStartRow & "," & mlngTime
        strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & mlngTime & "|" & _
            Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & VsfData.TextMatrix(lngStartRow, mlngTime) & "|0"
        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        '�����������
        With mrsCopyMap
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                For lngCol = 0 To lngCols - VsfData.FixedCols
                    Select Case lngCol + VsfData.FixedCols
                    Case 1, mlngDate, mlngTime, mlngOperator, mlngSigner, mlngSignTime, mlngRecord
                    Case Else
                        If InStr(1, "," & strLockItem & ",", "," & lngCol - (cHideCols - 1) & ",") = 0 And InStr(1, "," & mstrCOLNothing & ",", "," & lngCol - (cHideCols - 1) & ",") = 0 Then
                            VsfData.TextMatrix(lngStartRow + .AbsolutePosition - 1, lngCol + VsfData.FixedCols) = NVL(.Fields(cControlFields + lngCol).Value)
                            
                            '�޸ı�־
                            If .AbsolutePosition = 1 Then
                                strKey = mintҳ�� & "," & lngStartRow & "," & lngCol + VsfData.FixedCols
                                strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & lngCol + VsfData.FixedCols & "|" & _
                                    Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & GetMutilData(lngStartRow, lngCol + VsfData.FixedCols, lngTop, lngHeight) & "|0"
                                Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                            End If
                        End If
                    End Select
                Next
                .MoveNext
            Loop
        End With
        '�����ɫ
        'Call WriteColor
        mblnChange = True
        RaiseEvent AfterDataChanged(mblnChange Or mblnVerify)
    
    Case conMenu_Edit_Clear
        Dim lngRowCount As Long
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then Exit Sub
        If VsfData.TextMatrix(VsfData.ROW, mlngSigner) <> "" Then
            RaiseEvent AfterRowColChange("��ǩ�������ݲ�����ɾ����", True, mblnSign, mblnArchive)
            Exit Sub
        End If
        
        '��ҳ�����в���������н���ճ��,ɾ��,ֻ�ܱ༭�����Ŀ�����
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
            lngStartRow = VsfData.ROW
        Else
            lngStartRow = GetStartRow(VsfData.ROW)
        End If
        If InStr(1, VsfData.TextMatrix(lngStartRow, mlngRowCount), "|") <> 0 And lngStartRow = 3 Then
            If Val(Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0)) <> Val(VsfData.TextMatrix(lngStartRow, mlngRowCurrent)) Then
                RaiseEvent AfterRowColChange("��ҳ�����в�����ɾ�������л�����һҳ���в�����", True, mblnSign, mblnArchive)
                Exit Sub
            End If
        End If
        
        lngRowCount = Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0)
        '���Ŀ���������Ƿ����ͬ������������,�����������ͬ���ļ�¼
        strLockItem = GetSynItems(2, intMax)        '1.������Ŀ���;2.�����к�
        
        '׼��ɾ��
        strField = "ID|ҳ��|�к�|�к�|��¼ID|����|����|ɾ��"
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
            lngRow = VsfData.ROW
        Else
            lngRow = GetStartRow(VsfData.ROW)
            lngStartRow = lngRow
            If VsfData.TextMatrix(lngStartRow, mlngSigner) <> "" Then
                RaiseEvent AfterRowColChange("��ǩ�������ݲ�����ɾ����", True, mblnSign, mblnArchive)
                Exit Sub
            End If
            
            'ɾ������������
            lngRows = Val(Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0))
            For intNULL = 2 To lngRows
                VsfData.RowHidden(lngRow + intNULL - 1) = True
            Next
        End If
        
        '��¼�û��޸Ĺ��ĵ�Ԫ��
        If Val(VsfData.TextMatrix(lngStartRow, mlngCollectType)) = 0 Then
            If Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) > 0 Then
                If Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) > 0 Then
                    '�������޸ĲŽ�������̣�ȡ������¼��ʵ��ʱ��
                    strDate = Mid(VsfData.TextMatrix(lngStartRow, mlngActTime), 1, 10)
                    strTime = Mid(VsfData.TextMatrix(lngStartRow, mlngActTime), 12, 5)
                Else
                    '����ʱ���������
                    strDate = VsfData.TextMatrix(lngStartRow - Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) + 1, mlngDate)
                    strTime = VsfData.TextMatrix(lngStartRow - Val(VsfData.TextMatrix(lngStartRow, mlngDemo)) + 1, mlngTime)
                End If
            Else
                '��ͨ����
                strDate = VsfData.TextMatrix(lngStartRow, mlngDate)
                strTime = VsfData.TextMatrix(lngStartRow, mlngTime)
            End If
            If mblnDateAd Then
                If InStr(1, strDate, "/") <> 0 Then
                    strDate = Mid(zlDatabase.Currentdate, 1, 5) & Split(strDate, "/")(1) & "-" & Split(strDate, "/")(0)
                End If
                strDate = Mid(strDate, 9, 2) & "/" & Mid(strDate, 6, 2)
            End If
            
            strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|����|ɾ��"
            strKey = mintҳ�� & "," & lngStartRow & "," & mlngDate
            strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & mlngDate & "|" & _
                Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & strDate & "|" & VsfData.TextMatrix(lngStartRow, mlngDemo) & "|0|1"
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
            '2\ʱ��
            strKey = mintҳ�� & "," & lngStartRow & "," & mlngTime
            strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & mlngTime & "|" & _
                Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & strTime & "|" & VsfData.TextMatrix(lngStartRow, mlngDemo) & "|0|1"
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
            
            strField = "ID|ҳ��|�к�|�к�|��¼ID|����|����|ɾ��"
        Else
            '1\����
            strKey = mintҳ�� & "," & lngStartRow & "," & mlngDate
            strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & mlngDate & "|" & Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & _
                    VsfData.TextMatrix(lngStartRow, mlngCollectText) & ";" & Val(VsfData.TextMatrix(lngStartRow, mlngCollectType)) & ";" & _
                    Val(VsfData.TextMatrix(lngStartRow, mlngCollectStyle)) & ";" & VsfData.TextMatrix(lngStartRow, mlngCollectDay) & ";" & _
                    VsfData.TextMatrix(lngStartRow, mlngCollectStart) & ";" & VsfData.TextMatrix(lngStartRow, mlngCollectEnd) & "|1|1"
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        End If
        'ɾ����ʼ���з�ͬ��������
        If strLockItem = "" Then
            VsfData.RowHidden(lngRow) = True
            If Val(VsfData.TextMatrix(lngStartRow, mlngCollectType)) = 0 Then
                '��д�޸ı�־
                For lngCol = mlngTime + 1 To mlngNoEditor - 1
                    strKey = mintҳ�� & "," & lngStartRow & "," & lngCol
                    strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & lngCol & "|" & _
                        Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "||0|1"
                    Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                Next
            End If
        Else
            '��д�޸ı�־(����ͬ������,������ʱ���в��������)``
            For lngCol = mlngTime + 1 To mlngNoEditor - 1
                If InStr(1, "," & strLockItem & ",", "," & lngCol - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 And lngCol <> mlngDate And lngCol <> mlngTime Then
                    VsfData.TextMatrix(lngStartRow, lngCol) = ""
                    
                    strKey = mintҳ�� & "," & lngStartRow & "," & lngCol
                    strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & lngCol & "|" & _
                        Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "||0|1"
                    Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                End If
            Next
            VsfData.TextMatrix(lngStartRow, mlngRowCount) = "1|1"
        End If
            
        Call FillPage(False)
        If lngStartRow + lngRowCount < VsfData.Rows - 1 Then
            VsfData.ROW = lngStartRow + lngRowCount
        End If
        mblnChange = True
        RaiseEvent AfterDataChanged(mblnChange Or mblnVerify)
        
    Case conMenu_Edit_SPECIALCHAR
        
        '��鵱ǰ¼��ؼ�
        On Error Resume Next
        Dim objTXT As TextBox
        Dim strText As String
        Dim intPos As Integer, intLen As Integer
        
        mstrSymbol = frmInsSymbol.ShowMe(False, 0)
        If mintSymbol = -1 Then
            Set objTXT = txtInput
        Else
            Set objTXT = txt(mintSymbol)
        End If
        objTXT.SetFocus
        strText = objTXT.Text
        intPos = objTXT.SelStart
        intLen = Len(objTXT)
        objTXT.Text = Mid(strText, 1, intPos) & mstrSymbol & Mid(strText, intPos + 1)
    
        If mintSymbol = -1 Then
            Call txtInput_KeyDown(vbKeyReturn, 0)
        Else
            Call txt_KeyDown(Val(txt(mintSymbol)), vbKeyReturn, 0)
        End If
    
    Case conMenu_Edit_Append '����Ҫ��
        RaiseEvent AfterPartogramInfo(mlng�ļ�ID, mlngFileIndex, mlng��ʽID, mrsPartogram)
        'Call BoundItems(VsfData.COL - (cHideCols + VsfData.FixedCols - 1))
    Case conMenu_Edit_PrevPage
        If mintҳ�� > 1 Then
            If Not DataMap_Save Then Exit Sub
            mintҳ�� = mintҳ�� - 1
            '���²�ѯSQL
            '������ȡ����
            mblnInit = False
            Call InitVariable
            Call InitCons
            Call ReadStruDef
            Call zlRefresh
            mblnInit = True
            cbsThis.RecalcLayout
        End If
    Case conMenu_Edit_NextPage
        If mintҳ�� < mint����ҳ + 1 Then
            If Not DataMap_Save Then Exit Sub
            mintҳ�� = mintҳ�� + 1
            '���²�ѯSQL
            '������ȡ����
            mblnInit = False
            Call InitVariable
            Call InitCons
            Call ReadStruDef
            Call zlRefresh
            mblnInit = True
            cbsThis.RecalcLayout
        End If
    Case conMenu_Edit_Word
        Call cmdWord_Click
    End Select
    
err_exit:
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim arrData
    Dim blnFind As Boolean
    Dim strItem As String
    Dim intDo  As Integer, intCount As Integer
    On Error GoTo errHand
    
    If Not mblnInit Then Exit Sub
    Select Case Control.ID
    Case conMenu_Edit_FileMan '���Ӥ��
        Control.Enabled = Not mblnArchive And mblnEditable And Not mblnVerify And Not mblnChange
        If picBaby.Visible = True Then
            picBaby.Visible = Control.Enabled
        End If
    Case conMenu_Edit_Copy
        Control.Enabled = Not mblnShow And Not mblnArchive And Not mblnVerify And mblnEditable
    Case conMenu_Edit_PASTE
        Control.Enabled = False
        If mrsCopyMap.State = 0 Then Exit Sub
        'ǩ�����ݲ�����ճ��
        If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) <> "" Then
            intDo = GetStartRow(VsfData.ROW)
        Else
            intDo = VsfData.ROW
        End If
        If VsfData.TextMatrix(intDo, mlngSigner) <> "" Then Exit Sub
        If Val(VsfData.TextMatrix(intDo, mlngCollectType)) <> 0 Then Exit Sub
        
        Control.Enabled = Not mblnShow And Not mblnArchive And mblnEditable And mrsCopyMap.RecordCount
    Case conMenu_Edit_Clear
        Control.Enabled = False
        If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) <> "" Then
            intDo = GetStartRow(VsfData.ROW)
        Else
            intDo = VsfData.ROW
        End If
        If VsfData.TextMatrix(intDo, mlngSigner) <> "" Then Exit Sub
        
        Control.Enabled = Not mblnArchive And Not mblnVerify And mblnEditable
    Case conMenu_Edit_SPECIALCHAR
        Control.Enabled = mblnShow And Not mblnArchive And mblnEditable And (mintType = 0 Or mintType = 6)
    Case conMenu_Edit_Append '����Ҫ��
        'Control.Enabled = (InStr(1, "," & mstrCOLNothing & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0) And Not mblnArchive And mblnEditable
        Control.Enabled = Not mblnArchive And mblnEditable
    Case conMenu_Edit_PrevPage
        Control.Enabled = (mintҳ�� > 1)
    Case conMenu_Edit_NextPage
        Control.Enabled = (mintҳ�� < mint����ҳ + 1)
    Case conMenu_Edit_Word
        Control.Enabled = mblnEditAssistant And mblnShow And Not mblnArchive And mblnEditable
    End Select
errHand:
End Sub

Private Sub chkSwitch_Click()
    Dim blnSel As Boolean            '�Ƿ�ȫ��ѡ��
    Dim blnUpdate As Boolean
    Dim intLevel As Integer
    Dim lngRow As Long, lngRows As Long
    Dim strKey As String, strField As String, strValue As String
    '��������ȫ��ѡ�л�ȡ��ѡ�У����������
    
    If Not mblnInit Then Exit Sub
    lngRows = VsfData.Rows - 1
    strField = "ID|ҳ��|�к�|�к�|��¼ID|����|ɾ��"
    
    blnSel = chkSwitch.Value
    For lngRow = VsfData.FixedRows To lngRows
        If Not VsfData.RowHidden(lngRow) Then
            If VsfData.TextMatrix(lngRow, mlngRowCount) Like "*|1" Then
                '��������ҲҪǩ��,���ע��
                'If Val(VsfData.TextMatrix(lngRow, mlngCollectType)) = 0 Then    '�����в�����༭
                    blnUpdate = False
                    If blnSel Then
                        '���,ǩ�����ļ�¼,�ҵ�ǰ����Ա������ϴ�ǩ�������
                        If VsfData.TextMatrix(lngRow, mlngSignLevel) = "" Then
                            intLevel = δ����
                        Else
                            intLevel = Val(VsfData.TextMatrix(lngRow, mlngSignLevel)) + 1
                        End If
                        If mintVerify < intLevel And intLevel <> δ���� Then
                            blnUpdate = (VsfData.Cell(flexcpChecked, lngRow, mlngChoose) <> flexTSChecked)
                            VsfData.Cell(flexcpChecked, lngRow, mlngChoose) = flexTSChecked
                        End If
                    Else
                        blnUpdate = (VsfData.Cell(flexcpChecked, lngRow, mlngChoose) <> flexTSUnchecked)
                        VsfData.Cell(flexcpChecked, lngRow, mlngChoose) = flexTSUnchecked
                    End If
                    
                    If blnUpdate Then
                        '�����޸ļ�¼�Ա�ͬ��
                        strKey = mintҳ�� & "," & lngRow & "," & mlngChoose
                        strValue = strKey & "|" & mintҳ�� & "|" & lngRow & "|" & mlngChoose & "|" & _
                            Val(VsfData.TextMatrix(lngRow, mlngRecord)) & "|" & VsfData.Cell(flexcpChecked, lngRow, mlngChoose) & "|1"
                        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                    End If
                'End If
            End If
        End If
    Next
End Sub

Private Sub cmdAddBaby_Click()
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim lngGroupID As Long
    If picBaby.Visible = False Then Exit Sub
    
    On Error GoTo errHand
    
    lngGroupID = Val(cboBaby.ItemData(cboBaby.ListCount - 1))
    '��ӹ���Ϊ��һ��Ӥ���������������һ��Ӥ��
    strSQL = " SELECT 1" & vbNewLine & _
            " FROM ���˻����ļ� A, ���˻������� B, ���˻�����ϸ C,�����¼��Ŀ D" & vbNewLine & _
            " WHERE A.ID = B.�ļ�ID AND B.ID = C.��¼ID AND A.ID = [1] AND A.����ID = [2] AND A.��ҳID = [3] AND B.������� = [4]" & vbNewLine & _
            " AND substr(nvl(C.��¼����,''),1,1)='��' AND C.��Ŀ���=D.��Ŀ��� AND D.��Ŀ����='����' AND NVL(D.������Ŀ,0)=1"
    Call SQLDIY(strSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�ļ�ɾ�����", mlng�ļ�ID, mlng����ID, mlng��ҳID, lngGroupID)
    If rsTemp.RecordCount = 0 Then
        RaiseEvent AfterRowColChange("��ӹ���Ϊ��һӤ���ļ���Ӥ���Ѿ����������������һ����", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    cboBaby.AddItem lngGroupID + 1
    cboBaby.ItemData(cboBaby.NewIndex) = lngGroupID + 1
    cboBaby.ListIndex = cboBaby.ListCount - 1
    cboBaby.Refresh
    
    picBaby.Visible = False
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdBabyCancle_Click()
    picBaby.Visible = False
End Sub


Private Sub cmdCancel_Click()
    picBiref.Visible = False
End Sub

Private Sub cmdColumn_Click(Index As Integer)
    Dim lstItem As ListItem
    
    If cmdColumn(Index).Enabled = False Then Exit Sub
    If Index = 0 Then
        'add
        If Not lstColumnItems.SelectedItem Is Nothing Then
            Set lstItem = lstColumnUsed.ListItems.Add(, lstColumnItems.SelectedItem.Key, lstColumnItems.SelectedItem.Text)
            lstItem.SubItems(1) = lstColumnItems.SelectedItem.SubItems(1)
            lstItem.SubItems(2) = lstColumnItems.SelectedItem.SubItems(2)
            lstColumnItems.ListItems.Remove lstColumnItems.SelectedItem.Index
        End If
        If txtColumnNo.Text = "" Then
            txtColumnNo.Text = Replace(lstItem.SubItems(1), lstItem.SubItems(2), "")
        End If
    Else
        'del
        If Not lstColumnUsed.SelectedItem Is Nothing Then
            Set lstItem = lstColumnItems.ListItems.Add(, lstColumnUsed.SelectedItem.Key, lstColumnUsed.SelectedItem.Text)
            lstItem.SubItems(1) = lstColumnUsed.SelectedItem.SubItems(1)
            lstItem.SubItems(2) = lstColumnUsed.SelectedItem.SubItems(2)
            lstColumnUsed.ListItems.Remove lstColumnUsed.SelectedItem.Index
            If lstColumnUsed.ListItems.Count = 0 Then txtColumnNo.Text = ""
        End If
    End If
End Sub

Private Sub cmdDelBaby_Click()
    Dim intRow As Integer
    Dim lngFileIndex As Long, lngFileOldIndex As Long
    If picBaby.Visible = False Or Val(vsfBaby.RowData(vsfBaby.ROW)) < 2 Then Exit Sub
    
    lngFileIndex = Val(vsfBaby.RowData(vsfBaby.ROW))
    'Ϊ�˱�֤ɾ��ֻ�ܴӺ���ǰ���˴��ٴν����ж�
    For intRow = vsfBaby.FixedRows To vsfBaby.Rows - 1
        If lngFileOldIndex < Val(vsfBaby.RowData(intRow)) Then
            lngFileOldIndex = Val(vsfBaby.RowData(intRow))
        End If
    Next intRow
    
    If lngFileIndex < lngFileOldIndex Then
       RaiseEvent AfterRowColChange("ɾ��ֻ�ܴ����һ��Ӥ����ʼ�����飡", True, mblnSign, mblnArchive)
       Exit Sub
    End If
    
    If MsgBox("�˲�����ɾ�����Ӥ����ص�����������Ϣ���������Ƿ�Ҫɾ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    '���Ӥ����Ӧ�ļ����ݵ�ɾ��
'    zl_������������_DelBaby
    mstrSQL = "zl_������������_DelBaby("
'    �ļ�ID_IN ���˻�������.�ļ�ID%TYPE,
    mstrSQL = mstrSQL & mlng�ļ�ID & ","
'    Ӥ��_IN   ���˻�������.�������%TYPE
    mstrSQL = mstrSQL & lngFileIndex & ")"
    Call zlDatabase.ExecuteProcedure(mstrSQL, "zl_������������_DelBaby")
    '�������ˢ��
    mblnVerify = False
    mblnChange = False
    lngFileIndex = lngFileIndex - 1
    If lngFileIndex < 1 Then lngFileIndex = 1
    RaiseEvent AfterFileIndex(mlngFileIndex)
    RaiseEvent AfterDataSave(True)
    Call ShowMe(mFrmParent, mlng�ļ�ID, mlng����ID, mlng��ҳID, mlng����ID, mintӤ��, mstrPrivs, mblnEditable)
End Sub

Private Sub cmdFilterCancel_Click()
    picCloumn.Visible = False
End Sub

Private Sub cmdFilterOK_Click()
    Dim strPara As String
    Dim strTest As String
    Dim lngCol As Long, lngRow As Long
    Dim intDo As Integer, intCount As Integer, intFace As Integer
    On Error GoTo errHand
    
    If lstColumnUsed.ListItems.Count > 0 Then
        If Trim(txtColumnNo.Text) = "" Then
            RaiseEvent AfterRowColChange("��ͷ���Ʋ���Ϊ�գ�", True, mblnSign, mblnArchive)
            txtColumnNo.SetFocus
            Exit Sub
        End If
        If LenB(StrConv(txtColumnNo.Text, vbFromUnicode)) > 20 Then
            RaiseEvent AfterRowColChange("��ͷ���Ʋ��ܳ���10�����ֻ�20���ַ���", True, mblnSign, mblnArchive)
            txtColumnNo.SetFocus
            Exit Sub
        End If
    End If
    
    'ƴ������ʽ����ͷ����|��Ŀ���,��λ;��Ŀ���,��λ
    strPara = Trim(txtColumnNo.Text) & "|"
    intCount = lstColumnUsed.ListItems.Count
    If intCount > 2 Then
        RaiseEvent AfterRowColChange("ÿ�а󶨵���Ŀ�����ܳ���2����", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    '��Ŀ��ʾ����һ��
    For intDo = 1 To intCount
        mrsItems.Filter = "��Ŀ���=" & Val(lstColumnUsed.ListItems(intDo).Text)
        If intDo = 1 Then
            intFace = mrsItems!��Ŀ��ʾ
        Else
            If intFace <> mrsItems!��Ŀ��ʾ Then
                RaiseEvent AfterRowColChange("�󶨵�������Ŀ�ı�ʾ��������һ�£���Ҫô����ѡ���Ҫô������ֵ¼���", True, mblnSign, mblnArchive)
                mrsItems.Filter = 0
                Exit Sub
            End If
        End If
        
        'ƴ��
        strTest = lstColumnUsed.ListItems(intDo).Text
        If lstColumnUsed.ListItems(intDo).SubItems(2) <> "" Then
            strTest = strTest & "," & lstColumnUsed.ListItems(intDo).SubItems(2)
        End If
        If ISActiveUsed(strTest) Then Exit Sub
        
        strPara = strPara & IIf(intDo > 1, ";", "") & strTest
        mrsItems.Filter = 0
    Next
    
    '��������
    gstrSQL = "ZL_���˻���ҳ��_UPDATE(" & mlng�ļ�ID & "," & mintҳ�� & "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",'" & strPara & "','" & gstrUserName & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "������Ŀ������")
    picCloumn.Visible = False
    lngCol = VsfData.COL
    lngRow = VsfData.ROW
    
    '���²�ѯSQL
    '������ȡ����
    mblnInit = False
    Call InitVariable
    Call InitCons
    Call ReadStruDef
    Call zlRefresh
    mblnInit = True
    
    VsfData.ROW = lngRow
    VsfData.COL = lngCol
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function ISActiveUsed(ByVal strTest As String) As Boolean
    Dim arrData, arrCol
    Dim lngCol As Long
    Dim intDo As Integer, intCount As Integer
    Dim intIn As Integer, intMax As Integer
    '���ĳ�����Ŀ�Ƿ��ѱ������а�
    ISActiveUsed = True
    
    arrData = Split(mstrCOLActive, "||")
    intCount = UBound(arrData)
    For intDo = 0 To intCount
        arrCol = Split(Split(arrData(intDo), "|")(1), ";")
        lngCol = Split(Split(arrData(intDo), "|")(0), ";")(0)
        intMax = UBound(arrCol)
        For intIn = 0 To intMax
            If strTest = arrCol(intIn) And VsfData.COL - (cHideCols + VsfData.FixedCols - 1) <> lngCol Then
                RaiseEvent AfterRowColChange(Split(strTest, ",")(1) & mrsItems!��Ŀ���� & " �Ѿ����󶨵�" & lngCol & "�У��������ظ��󶨣�", True, mblnSign, mblnArchive)
                Exit Function
            End If
        Next
    Next
    ISActiveUsed = False
End Function

Private Function GetActivePart(ByVal intFindCol As Integer, ByVal intItem As Integer) As String
    '��ȡָ���еĻ��Ŀ
    Dim arrData
    Dim arrCol
    Dim intCOl As Integer, strPart As String
    Dim intDo As Integer, intCount As Integer
    Dim intIn As Integer, intMax As Integer
    '�����Ŀ���뵽��ѯSQL�У���ʽ���к�;��ͷ����|��Ŀ���,��λ;��Ŀ���,��λ||�к�;��ͷ����...
    '�󶨶����Ŀ�����о��Զ�תΪ�Խ�����
    
    arrData = Split(mstrCOLActive, "||")
    intCount = UBound(arrData)
    For intDo = 0 To intCount
        intCOl = Split(Split(arrData(intDo), "|")(0), ";")(0)
        If intCOl = intFindCol - cHideCols Then
            arrCol = Split(Split(arrData(intDo), "|")(1), ";")
            strPart = Split(arrCol(intItem), ",")(1)
            Exit For
        End If
    Next
    GetActivePart = strPart
End Function

Private Function CalcCollect(ByVal lngItem As Long, ByVal strStart As String, ByVal strEnd As String) As Double
    Dim dblCollect As Double
    On Error GoTo errHand
    
    gstrSQL = " SELECT  NVL(SUM(NVL(��¼����,0)),0) AS ����" & _
              " From ���˻�����ϸ A,���˻������� B," & vbNewLine & _
              "      (Select ��� From ���������Ŀ Start With ���=[2] Connect By Prior ���=�����) C" & vbNewLine & _
              " Where A.��¼ID=B.ID And A.��ֹ�汾 Is NULL And A.��¼����=1 AND B.�������=0 And A.��Ŀ���=C.���" & vbNewLine & _
              " And B.�ļ�ID=[1] And B.����ʱ�� Between [3] And [4]"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��������", mlng�ļ�ID, lngItem, CDate(strStart), CDate(strEnd))
    dblCollect = rsTemp!����
    
    CalcCollect = dblCollect
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cmdOK_Click()
    Dim ArrTime
    Dim arrItem
    Dim intType As Integer      'С�����
    Dim arrValue() As Double
    Dim bln���� As Boolean, blnExit As Boolean
    Dim lngStart As Long
    Dim lngCol As Long, lngCount As Long, lngRow As Long, lngRows As Long
    Dim strToday As String, str����ʱ�� As String
    Dim strStartDate As String, strEndDate As String
    Dim strStartTime As String, strEndTime As String
    Dim strKey As String, strField As String, strValue As String
    On Error GoTo errHand
    '����һ���µĻ��ܼ�¼
    
    If cboС��.Text = "��ʱ" And Val(txtС������.Tag) = 0 Then
        RaiseEvent AfterRowColChange("��ʼʱ������ʱ���ʽ����ȷ��", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    If InStr(1, txtС������.Text, ";") <> 0 Then
        RaiseEvent AfterRowColChange("С�������в��ܺ��зֺţ�", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    If InStr(1, txtС������.Text, "'") <> 0 Then
        RaiseEvent AfterRowColChange("С�������в��ܺ��е����ţ�", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    If LenB(StrConv(txtС������.Text, vbFromUnicode)) > 50 Then
        RaiseEvent AfterRowColChange("С�����Ʋ��ܳ���50���ַ���25�����֣�", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    '���ʱ�䷶Χ�Ƿ����
    '��ָ��������Ϊ׼
    '    �� ����
    '    ҹ ���� - ����
    '    ȫ ���� - ����
    strToday = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    ArrTime = Split(cboС��.Tag, ";")   '��ʽ:��ʼʱ��,����ʱ��;��ʼʱ��,����ʱ��
    strStartTime = txt��ʼʱ��.Text
    strEndTime = txt����ʱ��.Text
    If strEndTime < strStartTime Then bln���� = True
    If bln���� = True Then
        strStartDate = strToday & " " & strStartTime & ":00"
        strEndDate = DateAdd("d", 1, CDate(strToday)) & " " & strEndTime & IIf(cboС��.Text <> "��ʱ", ":59", ":00")
    Else
        strStartDate = strToday & " " & strStartTime & ":00"
        strEndDate = strToday & " " & strEndTime & IIf(cboС��.Text <> "��ʱ", ":59", ":00")
    End If

    strStartDate = Format(DateAdd("d", -1 * DateDiff("d", CDate(DTPDate.Value), CDate(strToday)), CDate(strStartDate)), "yyyy-MM-dd HH:mm:ss")
    strEndDate = Format(DateAdd("d", -1 * DateDiff("d", CDate(DTPDate.Value), CDate(strToday)), CDate(strEndDate)), "yyyy-MM-dd HH:mm:ss")
    
    If cboС��.Text <> "��ʱ" Then
        intType = -1 * cboС��.ItemData(cboС��.ListIndex)
        str����ʱ�� = DateAdd("s", 1 + cboС��.ItemData(cboС��.ListIndex), strEndDate) '�װ�С��+1��,ҹ��С��+2��,ȫ��С��+3��
    Else
        intType = -1 * (5 + Val(txtС������.Tag))
        str����ʱ�� = DateAdd("s", (5 + Val(txtС������.Tag)), strEndDate) '��ʱС�������+4+С��Сʱ
    End If
    
    
    '����Ƿ��Ѿ����ڸ�����
    blnExit = False
    mrsDataMap.Filter = "ɾ��=0 And �������=" & intType & " And ��������='" & str����ʱ�� & "'"    '��¼ID>0������,���ǵ��������
    blnExit = (mrsDataMap.RecordCount)
    mrsDataMap.Filter = 0
    
    If blnExit Then
        RaiseEvent AfterRowColChange("��Ҫ��ӵ�С�������Ѵ��ڣ�", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    If CDate(Format(strEndDate, "YYYY-MM-DD HH:mm:ss")) < CDate(Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm:ss")) Then
       RaiseEvent AfterRowColChange("С��Ľ���ʱ�䲻��С�ڹ�����ʼʱ��:[" & CDate(Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm:ss")) & "]", True, mblnSign, mblnArchive)
       Exit Sub
    End If
    
    '���ҿհ���
    lngRows = VsfData.Rows - 1
    For lngRow = VsfData.FixedRows To lngRows
        If Val(VsfData.TextMatrix(lngRow, mlngRecord)) = 0 And VsfData.TextMatrix(lngRow, mlngRowCount) = "" Then
            lngStart = lngRow
            Exit For
        End If
    Next
    If lngStart = 0 Then
        '˵��û���ҵ��հ���
        VsfData.Rows = VsfData.Rows + 1
        lngStart = VsfData.Rows - 1
    End If
    
    'ͳ�ƻ�������(�����ݿ��л���,��ǰ����ֻ��¼���Ƿ��޸�,����֪��ԭֵ�Ƕ���,���Ե�ǰδ��������ݲ�����)
    '������Ŀ����
    '������Ŀ�м���:col;1|col;4,5
    arrItem = Split(mstrCollectItems, ",")
    lngRows = UBound(arrItem)
    ReDim Preserve arrValue(lngRows) As Double
    For lngRow = 0 To lngRows
        arrValue(lngRow) = CalcCollect(arrItem(lngRow), strStartDate, strEndDate)
    Next
    
    'ͨ�ò���
    VsfData.TextMatrix(lngStart, mlngDate) = txtС������.Text
    VsfData.TextMatrix(lngStart, mlngTime) = txtС������.Text
    VsfData.TextMatrix(lngStart, mlngRowCount) = "1|1"                          'Ϊ�˱�֤ʱ�䲻�ظ�,��ȡ����ʱ��+��ķ�ʽ
    VsfData.TextMatrix(lngStart, mlngRowCurrent) = "1"
    VsfData.TextMatrix(lngStart, mlngCollectText) = txtС������.Text
    VsfData.TextMatrix(lngStart, mlngCollectType) = intType                     '��ʾС��;-1�װ�;-2ҹ��;3-ȫ��
    VsfData.TextMatrix(lngStart, mlngCollectStyle) = cbo��ʶ.ListIndex         '����24Сʱ,���»�����
    VsfData.TextMatrix(lngStart, mlngCollectDay) = str����ʱ��
    VsfData.TextMatrix(lngStart, mlngCollectStart) = txt��ʼʱ��.Text
    VsfData.TextMatrix(lngStart, mlngCollectEnd) = txt����ʱ��.Text
    
    'ͬ������������ʱ���е�����
    strField = "ID|ҳ��|�к�|�к�|��¼ID|����|����|ɾ��"
    '1\����
    strKey = mintҳ�� & "," & lngStart & "," & mlngDate
    strValue = strKey & "|" & mintҳ�� & "|" & lngStart & "|" & mlngDate & "|" & Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & _
            txtС������.Text & ";" & intType & ";" & cbo��ʶ.ListIndex & ";" & str����ʱ�� & ";" & txt��ʼʱ��.Text & ";" & txt����ʱ��.Text & "|1|0"
    Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
    
    'չ��
    arrItem = Split(mstrColCollect, "|")
    lngCount = 0
    lngRows = UBound(arrItem)
    For lngRow = 0 To lngRows
        lngCol = Split(arrItem(lngRow), ";")(0)
        If UBound(Split(Split(arrItem(lngRow), ";")(1), ",")) = 1 Then
            strValue = arrValue(lngCount) & "/" & arrValue(lngCount + 1)
            lngCount = lngCount + 2
        Else
            strValue = arrValue(lngCount)
            lngCount = lngCount + 1
        End If
        
        VsfData.TextMatrix(lngStart, lngCol + cHideCols) = strValue
        strKey = mintҳ�� & "," & lngStart & "," & lngCol + cHideCols
        strValue = strKey & "|" & mintҳ�� & "|" & lngStart & "|" & lngCol + cHideCols & "|" & _
            Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & VsfData.TextMatrix(lngStart, lngCol + cHideCols) & "|1|0"
        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
    Next
    
'    '�ϲ���Ԫ��
'    lngRows = Split(Split(mstrColCollect, "|")(0), ";")(0) + cHideCols - 1
'    For lngRow = mlngTime + 1 To lngRows
'        VsfData.TextMatrix(lngStart, lngRow) = txtС������.Text
'    Next
'    VsfData.MergeCells = flexMergeRestrictRows          '���ᵥԪ��Ȼ�ǵ����ϲ�,�ϲ�����������ϲ���Ԫ��
'    VsfData.MergeRow(lngStart) = True
    
    mblnChange = True
    picBiref.Visible = False
    
    RaiseEvent AfterDataChanged(mblnChange)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdWord_Click()
    Dim strInput As String
    '�����ʾ�ѡ����
    
    If cmdWord.Tag = -1 Then
        strInput = txtInput.Text
    Else
        strInput = txt(Val(cmdWord.Tag)).Text
    End If
    strInput = frmEditAssistant.ShowMe(Me, mlng����ID, mlng��ҳID, mintӤ��, strInput)
    
    If cmdWord.Tag = -1 Then
        txtInput.Text = strInput
        Call txtInput_KeyDown(vbKeyReturn, 0)
    Else
        txt(Val(cmdWord.Tag)).Text = strInput
        Call txt_KeyDown(Val(cmdWord.Tag), vbKeyReturn, 0)
    End If
End Sub

Private Sub ShowBrief()
    Dim strStart As String, strEnd As String
    Dim strHave As String, strDate As String
    Dim strTag As String    'cboС���tag�б���ʱ��Σ���ʽ����ʼ,����;��ʼ,����
    Dim rsData As New ADODB.Recordset
    Dim i As Integer
    Dim strCurDate As String
    Dim intStart As Integer, intEnd As Integer, intCur As Integer, intIndex As Integer, intCount As Integer
    On Error GoTo errHand
    '��ʾС�ᴰ��
    
    If Not DataMap_Save Then Exit Sub       '��������,�Ա�ѡ��С���ʱ��������ݼ��
    
    '����¼���Ƿ���ڻ�����Ŀ�У�������������˳�
    If mstrCollectItems = "" Then
        RaiseEvent AfterRowColChange("��ǰ�ļ���δʹ�û�����Ŀ��", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    '��ȡ����ʱ��(���=3Ϊȫ��С��)
    strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    gstrSQL = "Select ����,���,����,��ʼ,���� From �������ʱ�� Order by ��� "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡС��")
    rsTemp.Filter = "����=2"
    If rsTemp.RecordCount = 0 Then
        rsTemp.Filter = 0
        RaiseEvent AfterRowColChange("��δ���ü�¼��С��,�����ڻ�����Ŀ����ģ��Ļ�����Ŀ�����ã�", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    rsTemp.Filter = "����=1 And ���=3"
    If rsTemp.RecordCount = 0 Then
        rsTemp.Filter = 0
        RaiseEvent AfterRowColChange("ȫ�����ʱ��δ����,�����ڻ�����Ŀ����ģ��Ļ�����Ŀ�����ã�", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    strStart = NVL(rsTemp!��ʼ)
    strEnd = NVL(rsTemp!����)
    rsTemp.Filter = 0
    
    strCurDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD")

    With DTPDate
        .Value = Format(DateAdd("D", -1, CDate(strCurDate)), "YYYY-MM-DD")
        .MinDate = Format(DateAdd("D", -30, CDate(strCurDate)), "YYYY-MM-DD")
        If CDate(.MinDate) < CDate(Format(mstr��ʼʱ��, "YYYY-MM-DD")) Then .MinDate = Format(mstr��ʼʱ��, "YYYY-MM-DD")
        .MaxDate = Format(strCurDate, "YYYY-MM-DD")
    End With
    
    '���ػ������(��¼��С�����3��)
    intIndex = 0
    intCount = 0
    intCur = Format(Now, "HH")
    cboС��.Clear
    If strStart <> "" Or strEnd <> "" Then
        cboС��.AddItem "ȫ��С��"
        cboС��.ItemData(cboС��.NewIndex) = 4
        strTag = strTag & ";" & strStart & "," & strEnd
    End If
    intCount = intCount + 1
    
    With rsTemp
        rsTemp.Filter = "���� = 2"
        Do While Not .EOF
            If Not (NVL(!��ʼ) = "" Or NVL(!����) = "") Then
                cboС��.AddItem !����
                cboС��.ItemData(cboС��.NewIndex) = !���
                strTag = strTag & ";" & !��ʼ & "," & !����
                
                '��λ��ǰʱ���Ӧ��С��
                intStart = Val(!��ʼ)
                intEnd = Val(!����)
                If intStart <= intEnd Then
                    If intCur >= intStart And intCur <= intEnd Then intIndex = intCount
                Else
                    If intCur >= intStart Then intIndex = intCount
                End If
            End If
            
            intCount = intCount + 1
            .MoveNext
        Loop
        If strTag <> "" Then
            cboС��.Tag = Mid(strTag, 2)
            cboС��.ListIndex = intIndex
        Else
            rsTemp.Filter = 0
            RaiseEvent AfterRowColChange("����Ļ�����ȫ����ӣ�", True, mblnSign, mblnArchive)
            Exit Sub
        End If
        cboС��.AddItem "��ʱ"
    End With
    rsTemp.Filter = 0
    
    With cbo��ʶ
        .Clear
        .AddItem "������"
        .AddItem "���º��߱�ʶ"
        .AddItem "���ܽ����˫���߱�ʶ"
        .ListIndex = mintCollectDef
    End With
    
    '��������
    With picBiref
        .Top = VsfData.Top + 200
        .Left = (ScaleWidth - .Width) / 2
        .Visible = True
    End With
    
    On Error Resume Next
    cboС��.SetFocus
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub DTPDate_Change()
    cboС��_Click
End Sub

Private Sub DTPDate_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then cboС��.SetFocus
End Sub

Private Sub vsfBaby_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim CellRect As RECT
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHand
    
    If picBaby.Visible = False Or Val(vsfBaby.RowData(NewRow)) <= 0 Then Exit Sub
    With vsfBaby
        CellRect.Left = .CellLeft + .Left
        CellRect.Top = .CellTop + .Top
        CellRect.Bottom = .CellHeight + CellRect.Top
        CellRect.Right = .CellWidth + CellRect.Left
        cmdDelBaby.Top = CellRect.Top
        cmdDelBaby.Left = CellRect.Right - cmdDelBaby.Width
        cmdDelBaby.Height = CellRect.Bottom - CellRect.Top
        cmdDelBaby.Visible = True
        cmdDelBaby.Enabled = True
        '��һ���ļ�����ɾ��
        If .RowData(NewRow) = 1 Then cmdDelBaby.Visible = False: cmdDelBaby.Enabled = False: Exit Sub
        '�ļ�ֻ�ܴӺ���ǰɾ��
        strSQL = " SELECT 1" & vbNewLine & _
            " FROM ���˻����ļ� A, ���˻������� B, ���˻�����ϸ C" & vbNewLine & _
            " WHERE A.ID = B.�ļ�ID AND B.ID = C.��¼ID AND A.ID = [1] AND A.����ID = [2] AND A.��ҳID = [3] AND B.������� > [4]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�ļ�ɾ�����", mlng�ļ�ID, mlng����ID, mlng��ҳID, Val(.RowData(NewRow)))
        If rsTemp.RecordCount > 0 Then cmdDelBaby.Visible = False: cmdDelBaby.Enabled = False
    End With
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub VsfData_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Dim lngRow As Long, lngCol As Long
    Dim dblHeight As Double, dblWidth As Double
    
    If Not mblnInit Then Exit Sub
    Call InitCons
    
'    '����̶��еĸ߶�
'    For lngRow = 0 To 2
'        If Not VsfData.RowHidden(lngRow) Then dblHeight = dblHeight + VsfData.ROWHEIGHT(lngRow)
'    Next
'    '�ӿɼ��п�ʼ���²������һ���ɼ���
'    For lngRow = NewTopRow To VsfData.Rows - 1
'        If Not VsfData.RowIsVisible(lngRow) Then
'            lngRow = lngRow - 1
'            Exit For
'        End If
'    Next
'    '�ӿɼ��п�ʼ�������һ���ɼ���
'    For lngCol = NewLeftCol To VsfData.Cols - 1
'        If Not VsfData.ColIsVisible(lngCol) Then
'            lngCol = lngCol - 1
'            Exit For
'        Else
'            dblWidth = dblWidth + VsfData.ColWidth(lngCol)
'        End If
'    Next
'
'    If Not VsfData.RowIsVisible(VsfData.Row) Then
'        VsfData.Row = VsfData.Row + IIf(OldTopRow < NewTopRow, 1, -1)
'    Else
'        '��ǰ�����еĸ߶�+�̶��еĸ߶�������ڱ��ؼ��ĸ߶�,˵����ǰѡ��������д�����ס���ֵ����
'        If VsfData.Row >= lngRow - 1 And CellRect.Bottom * (lngRow - NewTopRow + 1) + dblHeight >= VsfData.ClientHeight Then
'            '��ס���ֵ������
'            VsfData.Row = VsfData.Row + IIf(OldTopRow < NewTopRow, 1, -1)
'        End If
'    End If
'
'    If Not VsfData.ColIsVisible(VsfData.Col) Then
'        VsfData.Col = VsfData.Col + IIf(OldLeftCol < NewLeftCol, 1, -1)
'    Else
'        '��ǰ�����еĸ߶�+�̶��еĸ߶�������ڱ��ؼ��ĸ߶�,˵����ǰѡ��������д�����ס���ֵ����
'        If VsfData.Col = lngCol And dblWidth >= VsfData.ClientWidth Then
'            '��ס���ֵ������
'            VsfData.Col = VsfData.Col + IIf(OldLeftCol < NewLeftCol, 1, -1)
'        End If
'    End If
'
'    Call VsfData_EnterCell
End Sub

Private Sub VsfData_DblClick()
    Call vsfdata_KeyDown(Asc("A"), 0)
End Sub

Private Sub VsfData_EnterCell()
    Dim strCols As String
    Dim strName As String
    Dim intMax As Integer
    Dim lngStart As Long
    On Error Resume Next
    
    '��������ʾ��¼��ؼ�
    Select Case mintType
    Case 0, 3
        picInput.Visible = False
    Case 1, 2
        lstSelect(mintType - 1).Visible = False
    Case 4, 5
        picDouble.Visible = False
    Case 6
        picMutilInput.Visible = False
    Case 7
        picDoubleChoose.Visible = False
    End Select
    cmdWord.Visible = False
    
    'δ������в�����¼������
    mintType = -1
    If InStr(1, mstrPrivs, "����ͼ��ͼ") = 0 Then Exit Sub
    
    If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "" Then
        lngStart = VsfData.ROW
    Else
        lngStart = GetStartRow(VsfData.ROW)
    End If
    If Val(VsfData.TextMatrix(lngStart, mlngCollectType)) < 0 And _
        (VsfData.COL <= mlngTime Or _
        InStr(1, "|" & mstrColCollect, "|" & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ";") <> 0) Then
        Exit Sub '�����в������޸�����ʱ��,�Լ������е�����
    End If
    If mblnVerify Then  '�������mblnShow�ж���������
        If VsfData.COL = mlngChoose Then Call vsfdata_KeyDown(vbKeySpace, 0): Exit Sub
        If VsfData.COL = mlngDate Or VsfData.COL = mlngTime Then Exit Sub
        If Val(VsfData.TextMatrix(lngStart, mlngRecord)) = 0 Then Exit Sub
        If VsfData.TextMatrix(lngStart, mlngSigner) = "" Then Exit Sub
    Else
        '��ǩ��������ֻ������ǩ״̬���޸�
        If InStr(1, VsfData.TextMatrix(lngStart, mlngSigner), "/") <> 0 Then
            RaiseEvent AfterRowColChange("����ǩ�����ݲ�����༭��", True, mblnSign, mblnArchive)
            Exit Sub
        End If
        
        '--ֻҪ��ǩ�������ݾͲ������޸�
        If VsfData.TextMatrix(lngStart, mlngSigner) <> "" Then
            RaiseEvent AfterRowColChange("��ǩ�������ݲ������ٴα༭����ȡ��ǩ�������ԣ�", True, mblnSign, mblnArchive)
            Exit Sub
        End If
        '�����ǰ����Ա�ļ������ǩ������Ա�ļ����,��������༭����
'        If VsfData.TextMatrix(lngStart, mlngSigner) <> "" Then
'            If mintVerify > Val(VsfData.TextMatrix(lngStart, mlngSignLevel)) + 1 Then
'                RaiseEvent AfterRowColChange("��ǰ����Ա�ļ������ǩ������Ա�ļ����,������༭���ݣ�", True, mblnSign, mblnArchive)
'                Exit Sub
'            End If
'        End If
        'Ĭ��ǩ�����뱣������ͬ,�������޸����˻����¼Ȩ�޵Ĳ���Ա,�������޸����˵�����
        strName = VsfData.TextMatrix(lngStart, IIf(mlngOperator = -1, VsfData.Cols - 1, mlngOperator))
        If strName <> "" Then
            If strName <> gstrUserName And _
                InStr(1, mstrPrivs, "���˻����¼") = 0 Then
                RaiseEvent AfterRowColChange("��û���޸����˻����¼���ݵ�Ȩ�ޣ�ԭ����Ա:" & strName, True, mblnSign, mblnArchive)
                Exit Sub
            End If
        End If
    End If
    If mblnArchive Then Exit Sub
    If Not mblnShow Or Not mblnEditable Then Exit Sub
    If VsfData.TextMatrix(VsfData.ROW, mlngDemo) <> "" Then
        'ֻ��������δ��������ݣ��������޸�������ʱ��
        If (VsfData.COL = mlngDate Or VsfData.COL = mlngTime) Then
            If Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) > 1 Then
                Exit Sub
            Else
                If Val(VsfData.TextMatrix(VsfData.ROW, mlngRecord)) > 0 Then Exit Sub
            End If
        End If
    End If
    
    '��ҳ�����в���������н���ճ��,ɾ��,ֻ�ܱ༭�����Ŀ�����
    If InStr(1, VsfData.TextMatrix(lngStart, mlngRowCount), "|") <> 0 And lngStart = 3 Then
        If Val(Split(VsfData.TextMatrix(lngStart, mlngRowCount), "|")(0)) <> Val(VsfData.TextMatrix(lngStart, mlngRowCurrent)) Then
            If InStr(1, "," & mstrCOLNothing & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0 Then
                RaiseEvent AfterRowColChange("�������޸Ŀ�ҳ�����еĻ��Ŀ���ݣ�", True, mblnSign, mblnArchive)
                Exit Sub
            End If
        End If
    End If
    'ͬ�������в�����༭
    strCols = GetSynItems(2, intMax)
    If strCols <> "" Then
        '����ͬ�����ݵ���,������ʱ���ǲ������޸ĵ�
        If VsfData.COL = mlngDate Or VsfData.COL = mlngTime Then Exit Sub
        strCols = "," & strCols & ","
        If InStr(1, strCols, "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0 Then Exit Sub
    End If
    
    If VsfData.COL <= mlngNoEditor - 1 Then Call ShowInput
    'Debug.Print txtInput.Text
    '�ÿؼ���ý���
    Select Case mintType
    Case 0, 3
        picInput.SetFocus
    Case 1, 2
        lstSelect(mintType - 1).SetFocus
    Case 4, 5
        picDouble.SetFocus
    Case 6
        picMutilInput.SetFocus
    End Select
    Debug.Print txtInput.Text
End Sub

Private Sub VsfData_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strInfo As String
    Dim strCols As String
    Dim intMax As Integer
    If mblnInit = False Then Exit Sub
    If mblnEditable = False Then Exit Sub
    If OldRow = NewRow And OldCol = NewCol Then Exit Sub
    On Error GoTo errHand
    
    'ѡ����,ͬ��������ֱ���˳�,����˴������ʾ��Ϣ
    If NewCol = mlngChoose Then Exit Sub
    strCols = GetSynItems(2, intMax)
    If strCols <> "" Then
        strCols = "," & strCols & ","
        If InStr(1, strCols, "," & NewCol - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0 Then Exit Sub
    End If
    
    '��ʾ��ǰ��Ŀ�������Ϣ
    mrsSelItems.Filter = "��=" & NewCol - (cHideCols + VsfData.FixedCols - 1)
    If mrsSelItems.RecordCount <> 0 Then
        mrsItems.Filter = "��Ŀ���=" & mrsSelItems!��Ŀ���
        If mrsItems.RecordCount <> 0 Then
            If NVL(mrsItems!��Ŀֵ��) <> "" Then
                If mrsItems!��Ŀ���� = 0 Then
                    strInfo = "��Ч��Χ:" & Split(mrsItems!��Ŀֵ��, ";")(0) & "��" & Split(mrsItems!��Ŀֵ��, ";")(1)
                Else
                    strInfo = "��Ч��Χ:" & mrsItems!��Ŀֵ��
                End If
            Else
                strInfo = ""
            End If
        End If
    End If
    mrsSelItems.Filter = 0
    mrsItems.Filter = 0
    
    '����Ƿ���ǩ��
    If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "" Then
        intMax = VsfData.ROW
    Else
        intMax = GetStartRow(VsfData.ROW)
    End If
    mblnSign = (VsfData.TextMatrix(intMax, mlngSigner) <> "")
    
    RaiseEvent AfterRowColChange(strInfo, False, mblnSign, mblnArchive)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsfData_DrawCell(ByVal hDC As Long, ByVal ROW As Long, ByVal COL As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Call DrawCell(hDC, ROW, COL, Left, Top, Right, Bottom, Done)
End Sub

Private Sub vsfdata_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngStart As Long
    Dim intLevel As Integer
    Dim strField As String, strKey As String, strValue As String
    On Error GoTo errHand
    
    If KeyCode = vbKeyReturn Then
        If Not mblnShow And VsfData.COL = mlngDate Then
            mblnShow = True
            Call VsfData_EnterCell
        Else
            Call MoveNextCell
        End If
    ElseIf KeyCode = vbKeySpace And mblnVerify Then
        'ֻ��ѡ��ʼ��
        lngStart = GetStartRow(VsfData.ROW)
        If VsfData.TextMatrix(lngStart, mlngTime) = "" Then Exit Sub
        
        '��ǩʱ,��ǰ��¼��ǩ��,�Ҳ���Ա��ǩ��������ϴ�ǩ������߲�����
        If VsfData.TextMatrix(lngStart, mlngSignLevel) = "" Then
            RaiseEvent AfterRowColChange("�����ݻ�δǩ�������ܽ�����ǩ��", True, mblnSign, mblnArchive)
            Exit Sub
        Else
            intLevel = Val(VsfData.TextMatrix(lngStart, mlngSignLevel)) + 1
        End If
        If mintVerify >= intLevel Then
            RaiseEvent AfterRowColChange("���ļ���Ҫ���ϴ���ǩ�˵ļ���߲��ܹ�ѡ�ü�¼��", True, mblnSign, mblnArchive)
            Exit Sub
        End If
        
        VsfData.Cell(flexcpChecked, lngStart, mlngChoose) = IIf(VsfData.Cell(flexcpChecked, lngStart, mlngChoose) = flexTSChecked, flexTSUnchecked, flexTSChecked)
        '�����޸ļ�¼�Ա�ͬ��
        strField = "ID|ҳ��|�к�|�к�|��¼ID|����|ɾ��"
        strKey = mintҳ�� & "," & lngStart & "," & mlngChoose
        strValue = strKey & "|" & mintҳ�� & "|" & lngStart & "|" & mlngChoose & "|" & Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & VsfData.Cell(flexcpChecked, lngStart, mlngChoose) & "|1"
        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
    Else
        If Not (KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or Shift <> 0) Then
            mblnShow = True
            Call VsfData_EnterCell
        End If
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitVariable()
    '�������
    mlngDate = -1
    mlngTime = -1
    mlngOperator = -1
    mlngSigner = -1
    mlngSignTime = -1
    mlngSignName = -1
    mlngRecord = -1
    mlngNoEditor = -1
    
    mblnChange = False
    mblnShow = False
    mblnSign = False
    mblnArchive = False
    mblnEditAssistant = False
    mblnVerify = False
    
End Sub

Private Sub InitCons()
    '��������ؼ�
    picInput.Visible = False
    lstSelect(0).Visible = False
    lstSelect(1).Visible = False
    picDouble.Visible = False
    picMutilInput.Visible = False
    cmdWord.Visible = False
    
    picBiref.Visible = False
    picCloumn.Visible = False
    picBaby.Visible = False
End Sub

Private Function InitMenuBar() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim cbrPop As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim rs As ADODB.Recordset
    Dim objExtendedBar As CommandBar
    
    On Error GoTo errHand
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsThis.ActiveMenuBar.Title = "�˵���"
    cbsThis.ActiveMenuBar.Visible = False
    
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    With cbsThis.Options
        .AlwaysShowFullMenus = False
        .ShowExpandButtonAlways = False
        .UseDisabledIcons = True
        .SetIconSize True, 16, 16
        .LargeIcons = True
    End With

    '------------------------------------------------------------------------------------------------------------------
    '�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    
        '------------------------------------------------------------------------------------------------------------------
        '����������
        Set cbrToolBar = cbsThis.Add("��׼", xtpBarTop)
        cbrToolBar.ShowTextBelowIcons = False
        cbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
        cbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
        With cbrToolBar.Controls
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_FileMan, "Ӥ��"): cbrControl.ToolTipText = "�ļ�����"
            Set cbrCustom = .Add(xtpControlCustom, conMenu_View_Option, "")
            cbrCustom.Flags = xtpFlagAlignLeft
            picTmp.Visible = True
            cbrCustom.Handle = picTmp.Hwnd
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Copy, "����"): cbrControl.ToolTipText = "����(Ctrl+C)": cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PASTE, "ճ��"):  cbrControl.ToolTipText = "ճ��(Ctrl+V)"
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Clear, "���"):   cbrControl.ToolTipText = "���"
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SPECIALCHAR, "�������"):  cbrControl.ToolTipText = "�����������(Ctrl+D)": cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Word, "�ʾ�ѡ��"):  cbrControl.ToolTipText = "�ʾ�ѡ��(Ctrl+W)"
            'Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Brief, "С��"): cbrControl.ToolTipText = "С��"
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Append, "������Ϣ"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "������Ϣ"
        
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PrevPage, "��ҳ"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "��ҳ"
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NextPage, "��ҳ"):   cbrControl.ToolTipText = "��ҳ"
        End With
    
        For Each cbrControl In cbrToolBar.Controls
            If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
                cbrControl.Style = xtpButtonIconAndCaption
            End If
        Next
    
         '�����
        With cbsThis.KeyBindings
            .Add FCONTROL, Asc("C"), conMenu_Edit_Copy
            .Add FCONTROL, Asc("V"), conMenu_Edit_PASTE
            .Add FCONTROL, Asc("D"), conMenu_Edit_SPECIALCHAR
            .Add FCONTROL, Asc("W"), conMenu_Edit_Word
            .Add FCONTROL, Asc("S"), conMenu_Save
        End With
    
    InitMenuBar = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckTxtTime(objText As TextBox) As String
    Dim strInput As String
    Dim strHour As String, strMin As String
    '���ʱ��¼���Ƿ�Ϸ�������������
    
    If Trim(objText.Text) <> "" Then
        strInput = Trim(objText.Text)
        If InStr(1, strInput, ":") > 0 Then
            strHour = Split(strInput, ":")(0)
            strMin = Split(strInput, ":")(1)
        ElseIf InStr(1, strInput, "��") > 0 Then
            strHour = Split(strInput, "��")(0)
            strMin = Split(strInput, "��")(1)
        Else
            strHour = strInput
            strMin = "00"
        End If
        strHour = Format(strHour, "00")
        strMin = Format(strMin, "00")
        If Not IsNumeric(strHour) Then
            RaiseEvent AfterRowColChange("��ʼʱ���к��зǷ��ַ���", True, mblnSign, mblnArchive)
            Exit Function
        End If
        If Val(strHour) < 0 Or Val(strHour) > 23 Then
            RaiseEvent AfterRowColChange("��ʼʱ�㲻��ȷ��СʱֵӦ��>0��С��24��", True, mblnSign, mblnArchive)
            Exit Function
        End If
        If Not IsNumeric(strMin) Then
            RaiseEvent AfterRowColChange("��ʼʱ���к��зǷ��ַ���", True, mblnSign, mblnArchive)
            Exit Function
        End If
        If Val(strMin) < 0 Or Val(strMin) > 59 Then
            RaiseEvent AfterRowColChange("��ʼʱ�㲻��ȷ������ֵӦ��>0��С��60��", True, mblnSign, mblnArchive)
            Exit Function
        End If
        strInput = strHour & ":" & strMin
    End If
    CheckTxtTime = strInput
End Function

Private Function CheckTime(ByVal lngRow As Long, ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
    ByVal strTime As String, ByVal strCurTime As String, ByRef strMsg As String) As Boolean
    Dim blnMsg As Boolean
    Dim blnExist As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '���ݷ���ʱ������ڵ�ǰ���ҵ���Чʱ�䷶Χ��
    
    blnMsg = (strMsg <> "")
    
    '����ļ���ʼ,����ʱ��
    If strTime < Format(mstr��ʼʱ��, "yyyy-MM-dd HH:mm") Or strTime > Format(mstr����ʱ��, "yyyy-MM-dd HH:mm") Then
        strMsg = "����ʱ��[" & strTime & "]���ڿ�ʼʱ��[" & Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm") & "]������ʱ��[" & Format(mstr����ʱ��, "YYYY-MM-DD HH:mm") & "]֮��"
        GoTo exitHand
    End If
    
    '������ڶ���ļ�,��һ�ļ���ʱ�䲻�ܴ�����һ�ļ���ʼʱ��
    If cboBaby.ListCount > 1 And cboBaby.ListIndex < cboBaby.ListCount - 1 Then
        gstrSQL = "Select 1 From ���˻����ļ� A,���˻������� B" & _
            "   Where A.ID=B.�ļ�ID And A.ID=[1] And A.����ID=[2] And A.��ҳID=[3] And A.Ӥ��=[4] AND B.�������=[5] And B.����ʱ��<=[6]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���ݼ��", mlng�ļ�ID, mlng����ID, mlng��ҳID, mintӤ��, cboBaby.ItemData(cboBaby.ListIndex + 1), CDate(strTime))
        If rsTemp.RecordCount > 0 Then
            strMsg = "��" & lngRow & "�еķ���ʱ��" & Format(strTime, "YYYY-MM-DD HH:mm") & "���󣬲��ܴ�����һӤ���ļ��Ŀ�ʼʱ�䣡"
            GoTo exitHand
        End If
    End If
    
    '���ݲ��˱䶯��¼���м��
    gstrSQL = " Select   ��ʼԭ��,����ID,to_char(��ʼʱ��,'yyyy-MM-dd hh24:mi') AS ��ʼʱ��,to_char(NVL(��ֹʱ��,sysDate+" & mintPreDays & "),'yyyy-MM-dd hh24:mi') AS ��ֹʱ�� " & _
              " From ���˱䶯��¼ " & _
              " Where ����ID=[1] And ��ҳID=[2]" & _
              " Order by ��ʼʱ��,��ʼԭ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��ǰ������Чʱ�䷶Χ", lng����ID, lng��ҳID)
    With rsTemp
        .Filter = "����ID=" & mlng����ID
        Do While Not .EOF
            If strTime >= !��ʼʱ�� And strTime <= !��ֹʱ�� Then
                blnExist = True
                Exit Do
            End If
            .MoveNext
        Loop
        .Filter = 0
        '�ҵ��˾��˳�
        If blnExist Then
            If Not IsAllowInput(lng����ID, lng��ҳID, strTime, strCurTime) Then
                strMsg = "��" & lngRow & "�еķ���ʱ��" & strTime & "����[�������ݲ�¼����Чʱ��:" & glngHours & "Сʱ]"
                GoTo exitHand
            End If
            
            CheckTime = True
            Exit Function
        End If
        
        'û�ҵ�,������ԭ�����׼ȷ����ʾ
'        .Filter = "��ʼԭ��=1"
'        If .RecordCount <> 0 Then
'            If !��ʼԭ�� = 1 And strTime < !��ʼʱ�� Then
'                strMsg = "��" & lngRow & "�еķ���ʱ��" & strTime & "����[����ʱ�䲻��С�ڲ�����Ժʱ��:" & !��ʼʱ�� & "]"
'                GoTo exitHand
'            End If
'        End If
'        .Filter = "��ʼԭ��=2"
'        If .RecordCount <> 0 Then
'            If !��ʼԭ�� = 2 And strTime < !��ʼʱ�� Then
'                strMsg = "��" & lngRow & "�еķ���ʱ��" & strTime & "����[����ʱ�䲻��С�ڲ������ʱ��:" & !��ʼʱ�� & "]"
'                GoTo exitHand
'            End If
'        End If
        .Filter = "��ʼԭ��=10"
        If .RecordCount <> 0 Then
            If !��ʼԭ�� = 10 And strTime > !��ֹʱ�� Then
                strMsg = "��" & lngRow & "�еķ���ʱ��" & strTime & "����[����ʱ�䲻�ܴ��ڳ�Ժʱ��:" & !��ֹʱ�� & "]"
                GoTo exitHand
            End If
        End If
'        .Filter = 0
'        '�������˵��
'        strMsg = "��" & lngRow & "�еķ���ʱ��" & strTime & "����[���ڵ�ǰ��������Чʱ�䷶Χ��]"
'        GoTo exitHand
    End With
    
    CheckTime = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
exitHand:
    rsTemp.Filter = 0
    If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
End Function

Private Function CheckInput(strReturn As String, strInfo As String) As Boolean
    Dim i As Integer, j As Integer
    Dim strOrders As String, strText As String
    '���¼�����ݵĺϷ���(����Ҳ��Ϊ��һ���ַ�,���ǵ�������Ŀ�ȴ��ڲ���\�������Ϣ)
    '���ص�����,���һ�а󶨶����Ŀ,�Ե�������Ϊ�ָ���
    
    'mintType:0=�ı���¼��;1=��ѡ;2=��ѡ;3=ѡ��;4-Ѫѹ��һ�а���������Ŀ,���ʽ����Ѫѹ��������Ŀ;5=һ�а���������Ŀ�Ҿ���ѡ����Ŀ;
    '6=һ�а�N����Ŀ,�ֹ�¼��
    Select Case mintType
    Case 0
        strText = txtInput.Text
        strOrders = txtInput.Tag
    Case 1, 2   '���
        If mintType = 1 Then
            If InStr(1, lstSelect(mintType - 1).Text, "-") <> 0 Then
                strText = Split(lstSelect(mintType - 1).Text, "-")(1)
            Else
                strText = ""
            End If
        Else
            j = lstSelect(mintType - 1).ListCount
            For i = 1 To j
                If lstSelect(mintType - 1).Selected(i - 1) Then
                    strText = strText & "," & Split(lstSelect(mintType - 1).List(i - 1), "-")(1)
                End If
            Next
            If strText <> "" Then strText = Mid(strText, 2)
        End If
        strOrders = lstSelect(mintType - 1).Tag
    Case 4
        strText = txtUpInput.Text & "'" & txtDnInput.Text
        strOrders = txtUpInput.Tag & "'" & txtDnInput.Tag
    Case 6
        j = txt.Count
        For i = 1 To j
            strText = strText & "'" & txt(i - 1).Text
            strOrders = strOrders & "'" & txt(i - 1).Tag
        Next
        If strText <> "" Then
            strText = Mid(strText, 2)
            strOrders = Mid(strOrders, 2)
        End If
    Case 3      '���
        strText = lblInput.Caption
    Case 5      '���
        strText = lblUpInput.Caption & "/" & lblDnInput.Caption
    Case 7
        strText = cboChoose(0).Text & "/" & cboChoose(1).Text
    End Select
    If Val(strOrders) <> 0 Then
        If Not CheckValid(strText, strOrders, strInfo) Then Exit Function
    ElseIf VsfData.COL = mlngDate Or VsfData.COL = mlngTime Then
        If Not CheckDateTime(strText, strInfo) Then Exit Function
    End If
    
    strReturn = strText
    CheckInput = True
End Function

Private Function CheckDateTime(strText As String, strInfo As String) As Boolean
    Dim blnCheck As Boolean, blnExist As Boolean
    Dim strCurrDate As String
    Dim strDate As String
    Dim rsCheck As New ADODB.Recordset
    
    On Error GoTo errHand
    
    If VsfData.COL = mlngDate Then
        If mblnDateAd Then
            If Trim(strText) = "" Then
                strInfo = "���ڲ���Ϊ�գ�"
                Exit Function
            End If
            If InStr(1, strText, "/") = 0 Then
                strInfo = "���ڸ�ʽ������1��12�գ�12/01"
                Exit Function
            End If
            
            strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
            strDate = Mid(strCurrDate, 1, 5) & ToStandDate(strText)
            If Not IsDate(strDate) Then
                strInfo = "¼������ݲ��ǺϷ������ڣ���1��12�գ�12/01"
                Exit Function
            End If
        Else
            If Trim(strText) = "" Then
                strInfo = "���ڲ���Ϊ�գ�"
                Exit Function
            End If
            If Not IsDate(strText) Then
                strInfo = "¼������ݲ��ǺϷ������ڣ���1��12�գ�2011-01-12"
                Exit Function
            End If
            strDate = Format(strText, "yyyy-MM-dd")
        End If
        If strDate > mstrMaxDate Then
            strInfo = "¼��������ѳ�������[����¼��������" & mintPreDays & "��]��ָ���ķ�Χ��"
            Exit Function
        End If
        
        If VsfData.TextMatrix(VsfData.ROW, mlngTime) <> "" Then
            blnCheck = True
            strDate = strDate & " " & VsfData.TextMatrix(VsfData.ROW, mlngTime)
        End If
    Else
        If Trim(strText) = "" Then
            strInfo = "ʱ�䲻��Ϊ�գ�"
            Exit Function
        End If
        If Len(strText) <= 2 Then
            strText = String(2 - Len(strText), "0") & strText
            strText = strText & ":00"
        End If
        If Val(Mid(strText, 1, 2)) < 0 Or Val(Mid(strText, 1, 2)) > 23 Then
            strInfo = "¼���ʱ����Ч��СʱӦ����0-23֮�䣡"
            Exit Function
        End If
        If Len(strText) < 5 And InStr(1, strText, ":") > 0 Then strText = String(5 - Len(strText), "0") & strText
        If Mid(strText, 3, 1) <> ":" Then
            strInfo = "¼���ʱ���ʽ����[09:00]��"
            Exit Function
        End If
        If Len(strText) < 5 Then strText = strText & String(5 - Len(strText), "0")
        If Not (Val(Mid(strText, 4, 2)) >= 0 And Val(Mid(strText, 4, 2)) <= 59) Then
            strInfo = "¼���ʱ����Ч������Ӧ����0-59֮�䣡"
            Exit Function
        End If
        If Len(strText) > 5 Then
            strInfo = "¼���ʱ���ʽ����[09:00]��"
            Exit Function
        End If
        
        'û������Ĭ�Ͻ������ڼ���
        If mblnDate = False Then
            If Format(strText, "HH:mm") >= Format(mstr��ʼʱ��, "HH:mm") Then
                strDate = Format(mstr��ʼʱ��, "YYYY-MM-DD")
            Else
                strDate = Format(CDate(mstr��ʼʱ��) + 1, "YYYY-MM-DD")
            End If
            VsfData.TextMatrix(VsfData.ROW, mlngDate) = strDate
        End If
        '���кϷ��Լ��
        If VsfData.TextMatrix(VsfData.ROW, mlngDate) <> "" Then
            strCurrDate = GetDateAdCurrDate(strText)
            strDate = VsfData.TextMatrix(VsfData.ROW, mlngDate)
            If mblnDateAd Then
                strDate = Mid(strCurrDate, 1, 5) & ToStandDate(strDate)
            Else
                strDate = Format(VsfData.TextMatrix(VsfData.ROW, mlngDate), "yyyy-MM-dd")
            End If
            strDate = Format(strDate & " " & strText, "YYYY-MM-DD HH:mm")
            
             '��������¼�뻹���޸ĵ����� ���������ʷ���ݶ��������޸�
            gstrSQL = " Select 1 From ���˻������� Where �ļ�ID=[1] And ����ʱ��=[2] And ([3]=0 OR ID<>[3])"
            Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "��鷢��ʱ��", mlng�ļ�ID, CDate(strDate), Val(VsfData.TextMatrix(VsfData.ROW, mlngRecord)))
            If rsCheck.RecordCount > 0 Then
                strInfo = "��¼���ʱ���Ѿ�������ʷ���ݣ�"
                Exit Function
            End If
            blnCheck = True
        End If
    End If
    
    If blnCheck Then
        '���ݷ���ʱ�䲻���ڵ�ǰ����Ա�������ҵ���Чʱ����ǰ
        If Not CheckTime(VsfData.ROW, mlng����ID, mlng��ҳID, strDate, strCurrDate, strInfo) Then
            Exit Function
        End If
    End If
    
    CheckDateTime = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetDateAdCurrDate(ByVal strTime As String) As String
'���ڴ��ڶԽ��ߣ���ȡ��ǰʱ��
    Dim strDate As String
    If Format(strTime, "HH:mm") >= Format(mstr��ʼʱ��, "HH:mm") Then
        strDate = Format(mstr��ʼʱ��, "YYYY-MM-DD")
    Else
        strDate = Format(CDate(mstr��ʼʱ��) + 1, "YYYY-MM-DD")
    End If
    GetDateAdCurrDate = Format(strDate & " " & Format(strTime, "HH:mm") & ":00", "yyyy-MM-dd HH:mm")
End Function

Private Function CheckValid(strReturn As String, ByVal strOrders As String, strInfo As String) As Boolean
    Dim arrData, arrOrder
    Dim blnCheck As Boolean
    Dim i As Integer, j As Integer
    Dim dblMin As Double, dblMax As Double
    Dim strText As String, strName As String, strFormat As String, strFormat1 As String
    
    '���и�ʽ��װ����
    mrsSelItems.Filter = "��=" & VsfData.COL - (cHideCols + VsfData.FixedCols - 1)
    If mrsSelItems.RecordCount <> 0 Then
        '�д��е�δ���ж���
        strFormat = NVL(mrsSelItems!��ʽ)   '{P[����]C}{...}
        strFormat1 = strFormat
    End If
    mrsSelItems.Filter = 0
    
    '�������
    arrData = Split(strReturn, "'")
    arrOrder = Split(strOrders, "'")
    j = UBound(arrData)
    For i = 0 To j
        strText = arrData(i)
        strOrders = arrOrder(i)
        If Val(strOrders) <> 0 Then
            mrsItems.Filter = "��Ŀ���=" & strOrders
            strName = GetActivePart(VsfData.COL, i) & mrsItems!��Ŀ����
            If strText <> "" Then
                blnCheck = True
                '�����������Ŀ,�������Ĳ����������򲻼��
                If mrsItems!��Ŀ��� >= 1 And mrsItems!��Ŀ��� <= 3 Then
                    If Not IsNumeric(Trim(strText)) Then
                        blnCheck = False
                    End If
                End If
                
                If blnCheck Then
                    If mrsItems!��Ŀ���� = 0 And InStr(1, "0,4", mrsItems!��Ŀ��ʾ) <> 0 Then
                        strText = Val(strText)
                        If NVL(mrsItems!��ĿС��, 0) <> 0 Then   '��������ͨ���ؼ���MaxLength�����Ƶ�
                            If InStr(1, strText, ".") <> 0 Then strText = Mid(strText, 1, InStr(1, strText, ".") - 1)
                            If Len(strText) > mrsItems!��Ŀ���� Then
                                mrsItems.Filter = 0
                                strInfo = "[" & strName & "]¼������ݳ����˺Ϸ����ȣ�"
                                Exit Function
                            End If
                            
                            strText = Val(arrData(i))
                            If InStr(1, strText, ".") <> 0 Then
                                strText = Mid(strText, InStr(1, strText, ".") + 1)
                                If Len(strText) > mrsItems!��ĿС�� Then
                                    mrsItems.Filter = 0
                                    strInfo = "[" & strName & "]¼���С�����ֳ����˺Ϸ����ȣ�"
                                    Exit Function
                                End If
                            End If
                            strText = Val(arrData(i))
                        End If
                        If mrsItems!��Ŀ��ʾ = 0 Then
                            If Not IsNull(mrsItems!��Ŀֵ��) Then
                                dblMin = Val(Split(mrsItems!��Ŀֵ��, ";")(0))
                                dblMax = Val(Split(mrsItems!��Ŀֵ��, ";")(1))
                                If Not (Val(strText) >= dblMin And Val(strText) <= dblMax) Then
                                    mrsItems.Filter = 0
                                    strInfo = "[" & strName & "]¼������ݲ���" & Format(dblMin, "#0.00") & "��" & Format(dblMax, "#0.00") & "����Ч��Χ��"
                                    Exit Function
                                End If
                            End If
                        End If
                    Else
                        If LenB(StrConv(strText, vbFromUnicode)) > mrsItems!��Ŀ���� Then
                            strInfo = "[" & strName & "]¼������ݳ�������󳤶ȣ�" & mrsItems!��Ŀ���� & "��"
                            mrsItems.Filter = 0
                            Exit Function
                        End If
                    End If
                End If
                strFormat = Replace(strFormat, "[" & strName & "]", strText)
            Else
                'ɾ������Ŀ
                If InStr(1, "," & mstrCOLNothing & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
                    Call SubstrPro(strFormat, strName)
                Else
                    '����Ŀ������ʱ,�����ǰ�о��жԽ�������,�����
                    strFormat = Replace(strFormat, "[" & strName & "]", strText)
                End If
            End If
        Else
            strFormat = strReturn
        End If
    Next
    If j = -1 Then
        strOrders = arrOrder(i)
        If Val(strOrders) <> 0 Then
            mrsItems.Filter = "��Ŀ���=" & strOrders
            strName = mrsItems!��Ŀ����
            strFormat = Replace(strFormat, "[" & strName & "]", strText)
        End If
    End If
    mrsItems.Filter = 0
    
    strFormat = Replace(strFormat, "{", "")
    strFormat = Replace(strFormat, "}", "")
    If InStr(1, "," & mstrCOLNothing & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
        If strFormat = SubstrFormat(strFormat1, arrOrder) Then strFormat = ""
    End If
    
    If IsNumeric(strFormat) Then
        If Val(strFormat) < 1 And Val(strFormat) > 0 Then strFormat = "0" & strFormat
    End If
    strReturn = strFormat
    
    CheckValid = True
End Function

Public Function SubstrFormat(ByVal strData As String, ByVal arrOrder As Variant) As String
    '��ȡ����Ŀ��ǰ��׺����
    Dim i As Integer
    Dim strOrders As String, strName As String
    For i = 0 To UBound(arrOrder)
        strOrders = CStr(arrOrder(i))
        If Val(strOrders) <> 0 Then
            mrsItems.Filter = "��Ŀ���=" & strOrders
            strName = GetActivePart(VsfData.COL, i) & mrsItems!��Ŀ����
        End If
        strData = Replace(strData, "[" & strName & "]", "")
    Next i
    strData = Replace(strData, "{", "")
    strData = Replace(strData, "}", "")
    
    SubstrFormat = strData
End Function

Public Function SubstrVal(ByVal strData As String, ByVal strFormat As String, ByVal strName As String, intPos As Integer) As String
    Dim i As Integer, j As Integer, l As Integer, r As Integer
    Dim strQZ As String, strHZ As String
    '����ǰһ����Ŀ�ĺ�׺����+��ǰ��Ŀ��ǰ׺���ŵ�λ��
    
    If strData = "" Then Exit Function
    strData = strData
    j = Len(strFormat)
    l = InStr(1, strFormat, "[" & strName & "]")
    If l = 0 Then Exit Function
    '�õ�ǰ׺
    For i = l To 1 Step -1
        If Mid(strFormat, i, 1) = "{" Then Exit For
    Next
    strQZ = Mid(strFormat, i + 1, l - i - 1)
    '�ҵ�����Ŀ��ʽ���еĽ�������
    i = l + Len(strName) + 2
    For r = i To j
        If Mid(strFormat, r, 1) = "}" Then Exit For
    Next
    '�õ���׺
    strHZ = Mid(strFormat, i, r - i)
    '�����׺Ϊ��,�������Ѱ����һ����Ŀ��ǰ׺����
    If strHZ = "" And r < j Then
        For r = r + 1 To j
            If Mid(strFormat, r, 1) = "[" Then Exit For
        Next
        strHZ = Mid(strFormat, InStr(i, strFormat, "{") + 1, r - InStr(i, strFormat, "{") - 1)
    End If
    'ȡ��ָ����Ŀ���������ݴ�
    If strHZ <> "" Then
        j = InStr(intPos, strData, strHZ) '��Ϊ������ȡ��,���ǵ��ָ���������ͬ�����,��¼��һ�ε����λ��,�´δ����λ������ȡ����
        If j = 0 Then
            '�п����м���ڻس����з�
            j = InStr(intPos, Replace(strData, vbCrLf, ""), strHZ)
            If j = 0 Then Exit Function
        End If
    End If
    strData = Mid(strData, intPos)
    'ǰ׺Ϊ��,������ǰѰ����һ����Ŀ�ĺ�׺����
'    If strQZ = "" And i > 1 And intPos > 1 Then
'        For i = i - 1 To 1 Step -1
'            If Mid(strFormat, i, 1) = "]" Then Exit For
'        Next
'        strQZ = Mid(strFormat, i + 1, InStr(i, strFormat, "}") - i - 1)
'    End If
    
    SubstrVal = SubstrAnaly(strData, strHZ, strQZ)
    intPos = intPos + Len(strQZ & SubstrVal & strHZ)
    '�������������ȥ���س����з�����,������ַ�����ԭ������
'    If strHZ <> "" Then
'
'        strData = Mid(strData, 1, InStr(1, Replace(strData, vbCrLf, ""), strHZ) - 1) '��������Ŀ�������
'        intPOS = i + Len(strHZ)
'    End If
'    If strQZ <> "" Then strData = Mid(strData, InStr(1, strData, strQZ) + Len(strQZ)) '��������Ŀ�������
'    SubstrVal = strData ' Replace(strData, vbCrLf, "")
End Function

Private Function SubstrAnaly(ByVal strData As String, ByVal strHZ As String, ByVal strQZ As String) As String
    Dim strText As String
    Dim strCompare As String           '�Աȴ�
    Dim intLen As Integer, intActLen As Integer           'ǰ׺/��׺�ĳ���
    Dim intPos As Integer, intEnd As Integer
    Dim lngASC As Long
    Dim blnFind As Boolean
    '�����س����з�����,�ո����±ȶ�
    
    strText = strData
    If strHZ <> "" Then
        '�Ѻ�׺ȥ��
        strHZ = Replace(strHZ, vbCrLf, "")
        intEnd = Len(strText)
        intLen = Len(strHZ)
        For intPos = 1 To intEnd
            lngASC = Asc(Mid(strText, intPos, 1))
            intActLen = intActLen + 1
            If Not (lngASC = 13 Or lngASC = 10) Then
                If lngASC = 32 Then
                    strCompare = ""
                    intActLen = 0
                Else
                    strCompare = strCompare & Mid(strText, intPos, 1)
                End If
                If Len(strCompare) = intLen Then
                    If strCompare = strHZ Then
                        blnFind = True
                        intPos = intPos - intActLen
                    Else
                        strCompare = ""
                        intPos = intPos - intActLen + 1
                        intActLen = 0
                    End If
                End If
            End If
            If blnFind Then Exit For
        Next
        '�϶���
        strText = Mid(strText, 1, intPos)
    End If
    
    '��ȥ��ǰ׺
    If strQZ <> "" Then
        If InStr(1, strText, strQZ) = 0 Then strText = strQZ & strText
        strQZ = Replace(strQZ, vbCrLf, "")
        intEnd = Len(strText)
        intLen = Len(strQZ)
        strCompare = ""
        intActLen = 0
        blnFind = False
        For intPos = 1 To intEnd
            lngASC = Asc(Mid(strText, intPos, 1))
            intActLen = intActLen + 1
            If Not (lngASC = 13 Or lngASC = 10) Then
                If lngASC = 32 Then
                    strCompare = ""
                    intActLen = 0
                Else
                    strCompare = strCompare & Mid(strText, intPos, 1)
                End If
                If Len(strCompare) = intLen Then
                    If strCompare = strQZ Then
                        blnFind = True
                        intPos = intPos + 1
                    Else
                        strCompare = ""
                        intPos = intPos + 1
                        intActLen = 0
                    End If
                End If
            End If
            If blnFind Then Exit For
        Next
        strText = Mid(strText, intPos)
    End If
    
    If IsNumeric(Replace(strText, vbCrLf, "")) Then
        SubstrAnaly = Replace(strText, vbCrLf, "")
    Else
        SubstrAnaly = strText
    End If
End Function

Public Sub SubstrPro(strFormat As String, ByVal strName As String, Optional ByVal intType As Integer = 0)
    Dim i As Integer, j As Integer, l As Integer, r As Integer, strHZ As String, strQZ As String
    'intType=0-ɾ��ָ����ʽ��;1-�õ�ָ����ʽ��
    j = Len(strFormat)
    i = InStr(1, strFormat, "[" & strName & "]")
    If i = 0 Then Exit Sub
    
    For l = i To 1 Step -1
        If Mid(strFormat, l, 1) = "{" Then Exit For
    Next
    strQZ = Mid(strFormat, l + 1, i - l - 1)
    For r = i To j
        If Mid(strFormat, r, 1) = "}" Then Exit For
    Next
    strHZ = Mid(strFormat, i + Len(strName) + 2, r - i - Len(strName) - 2)
    If intType = 0 Then
        'strFormat = Mid(strFormat, 1, l - 1) & strQZ & strHZ & Mid(strFormat, r + 1)
        If Mid(strFormat, 1, l - 1) = "" And Mid(strFormat, r + 1) = "" Then
            strFormat = ""
        Else
            strFormat = Mid(strFormat, 1, l - 1) & strQZ & strHZ & Mid(strFormat, r + 1)
        End If
    Else
        strFormat = Mid(strFormat, l, r - l + 1)
    End If
End Sub

Private Sub MoveNextCell(Optional ByVal blnNext As Boolean = True)
    Dim arrData
    Dim blnNULL As Boolean                      '�Ƿ�Ϊ����
    Dim blnGroup As Boolean                     '������
    Dim strDate As String, strTime As String    '�����׼�¼��������ʱ��
    Dim strReturn As String, strMsg As String, strPart As String
    Dim lngStart As Long, lngMutilRows As Long, lngDeff As Long
    Dim intRow As Integer, intCount As Integer, intNULL As Integer  '����ж��ٿ���
    Dim blnTrue As Boolean
    '��ֵȻ���ƶ�����һ����Ч��Ԫ��
    Dim strKey As String, strField As String, strValue As String, strAppend As String
    Dim blnCallback As Boolean
    
    On Error GoTo errHand
    
    '�������,���ϸ���ٴε���Ҫ��¼��
    If mintType >= 0 Then
        If Not CheckInput(strReturn, strMsg) Then
            RaiseEvent AfterRowColChange(strMsg, True, mblnSign, mblnArchive)
            Exit Sub
        End If
        
        '��ǵ�ǰ��Ϊ������
        If mstrClassRow <> "" Then
            VsfData.TextMatrix(VsfData.ROW, mlngDemo) = VsfData.ROW - Val(mstrClassRow) + 1
            blnGroup = True
        Else
            
            blnGroup = ((VsfData.TextMatrix(VsfData.ROW, mlngDemo) = "1") And mblnEditAssistant) Or (Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) > 1) '����ı��в��Զ��ֽ�
        End If
        
        blnGroup = False
        blnTrue = False
        If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "" Then VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "1|1"
        If Not blnGroup Then
            If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") <> 0 Then
                lngMutilRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
            End If
            lngStart = GetStartRow(VsfData.ROW)
        Else
            lngMutilRows = 1
            If VsfData.TextMatrix(VsfData.ROW, mlngDemo) = 1 Then
                For intCount = VsfData.ROW + 1 To VsfData.Rows - 1
                    If Val(VsfData.TextMatrix(intCount, mlngDemo)) <= 1 Then Exit For      '����������·�����˳�
                    lngMutilRows = lngMutilRows + 1
                Next
            Else
                lngMutilRows = 1
            End If
            lngStart = VsfData.ROW
        End If
        
        
        '׼����ֵ
        With txtLength
            '������ʱ���еĿ�Ȳ���,Ϊ�˱��ⷵ�ض���,ǿ������Ϊ5000
            .Width = IIf(VsfData.COL = mlngDate Or VsfData.COL = mlngTime, 5000, VsfData.CellWidth)
            .Text = strReturn
            .FontName = VsfData.CellFontName
            .FontSize = VsfData.CellFontSize
            .FontBold = VsfData.CellFontBold
            .FontItalic = VsfData.CellFontItalic
        End With
        arrData = GetData(txtLength.Text)
        intCount = UBound(arrData)
        
        If intCount > lngMutilRows - 1 Then
            If blnGroup = True And lngMutilRows = 1 Then
                strMsg = "�Ƿ�����ʼ�У��޸ĵķ��������в��ܴ���ԭ��������,�����ʼ���޸�."
                RaiseEvent AfterRowColChange(strMsg, True, mblnSign, mblnArchive)
                strMsg = ""
                Exit Sub
            End If
            '������������,�������������������������ӵ�����
            '20110830���������ͬһ�����У��������ı��ֽ⵽���У�������ı�����ͳһ�������һ����;�ڷ����а��س�,ֻ���������ݽ����޸�,�����з����仯
            intNULL = intCount - (lngMutilRows - 1)
            For intRow = lngMutilRows To intCount
                '��֤��ǰ�����������һҳ����ʾȫ
                If intRow + lngStart > VsfData.Rows - 1 Then Exit For
                
                If Val(VsfData.TextMatrix(intRow + lngStart, mlngRecord)) = 0 And VsfData.TextMatrix(intRow + lngStart, mlngRowCount) = "" Then
                    intNULL = intNULL - 1
                Else
                    Exit For
                End If
            Next
            '�����ӿ���
            If intNULL > 0 Then
                lngDeff = intNULL
                VsfData.Rows = VsfData.Rows + intNULL
                '�ӵ�ǰ�м�¼�Ŀհ��п�ʼ��ÿ�е�λ��+�����ӵĿհ�����
                For intRow = VsfData.Rows - intNULL - 1 To lngStart + intCount - intNULL + 1 Step -1
                    VsfData.RowPosition(intRow) = intRow + intNULL
                Next
            End If
            'ѭ����ֵ
            intCount = UBound(arrData)
            For intRow = 0 To intCount
                VsfData.TextMatrix(lngStart + intRow, VsfData.COL) = Replace(Replace(arrData(intRow), Chr(10), ""), Chr(13), "")
                If Not blnGroup Then
                    VsfData.TextMatrix(lngStart + intRow, mlngRowCount) = intCount + 1 & "|" & intRow + 1
                    VsfData.TextMatrix(lngStart + intRow, mlngRowCurrent) = intCount + 1
                Else
                    '�����е����⴦��,�����ڲ���¼���Ĵ���϶�
                    '##########################################
                    If lngMutilRows > 1 Then
                        '����������Զ��ֽ�ʱ�����ÿհ��е�demo��;���򲻴���(�Զ�����µķ�����ϸ��)
                        VsfData.TextMatrix(lngStart + intRow, mlngDemo) = intRow + 1
                        VsfData.TextMatrix(lngStart + intRow, mlngRowCount) = "1|1"
                    End If
                    
                    '��������
                    If Val(VsfData.TextMatrix(lngStart + intRow, mlngDemo)) > 0 Then
                        If Val(VsfData.TextMatrix(lngStart + intRow, mlngRecord)) > 0 Then
                            '�������޸ĲŽ�������̣�ȡ������¼��ʵ��ʱ��
                            If mblnDateAd Then
                                strDate = Format(VsfData.TextMatrix(lngStart + intRow, mlngActTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStart + intRow, mlngActTime), "MM")
                            Else
                                strDate = Mid(VsfData.TextMatrix(lngStart + intRow, mlngActTime), 1, 10)
                            End If
                            strTime = Mid(VsfData.TextMatrix(lngStart + intRow, mlngActTime), 12, 5)
                        Else
                            '����ʱ���������
                            strDate = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngDate)
                            strTime = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngTime)
                        End If
                    Else
                        '��ͨ����
                        strDate = VsfData.TextMatrix(lngStart + intRow, mlngDate)
                        strTime = VsfData.TextMatrix(lngStart + intRow, mlngTime)
                    End If
                    
                    '1\����
                    strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|ɾ��"
                    If mlngDate <> -1 Then
                        strKey = mintҳ�� & "," & lngStart + intRow & "," & mlngDate
                        strValue = strKey & "|" & mintҳ�� & "|" & lngStart + intRow & "|" & mlngDate & "|" & _
                            Val(VsfData.TextMatrix(lngStart + intRow, mlngRecord)) & "|" & strDate & "|" & VsfData.TextMatrix(lngStart + intRow, mlngDemo) & "|0"
                        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                    End If
                    '2\ʱ��
                    strKey = mintҳ�� & "," & lngStart + intRow & "," & mlngTime
                    strValue = strKey & "|" & mintҳ�� & "|" & lngStart + intRow & "|" & mlngTime & "|" & _
                        Val(VsfData.TextMatrix(lngStart + intRow, mlngRecord)) & "|" & strTime & "|" & _
                        VsfData.TextMatrix(lngStart + intRow, mlngDemo) & "|0"
                    Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                    
                    strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|ɾ��"
                    strKey = mintҳ�� & "," & lngStart + intRow & "," & VsfData.COL
                    strValue = strKey & "|" & mintҳ�� & "|" & lngStart + intRow & "|" & VsfData.COL & "|" & _
                        Val(VsfData.TextMatrix(lngStart + intRow, mlngRecord)) & "|" & Replace(Replace(arrData(intRow), Chr(10), ""), Chr(13), "") & "||0"
                    Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                    '##########################################
                End If
            Next
            '���������н��и�ֵ
            lngMutilRows = lngStart + intCount
            For intRow = lngStart + 1 To lngMutilRows
                For intCount = 0 To VsfData.Cols - 1
                    VsfData.Cell(flexcpForeColor, intRow, intCount) = VsfData.Cell(flexcpForeColor, lngStart, intCount)
                    If VsfData.ColHidden(intCount) And InStr(1, "," & mlngRowCount & "," & mlngRowCurrent & ",", "," & intCount & ",") = 0 Then
                        If blnGroup And InStr(1, "," & mlngDemo & "," & mlngRecord & ",", "," & intCount & ",") = 0 Then
                            VsfData.TextMatrix(intRow, intCount) = VsfData.TextMatrix(lngStart, intCount)
                        End If
                    End If
                Next
            Next
        Else
            '�Ը������¸�ֵ����ֻ����һ������ʱ����֪Ϊ�λ�����ַ�ASCII��Ϊ1�ķ��ţ�
            For intRow = 0 To intCount
                VsfData.TextMatrix(lngStart + intRow, VsfData.COL) = Replace(Replace(Replace(arrData(intRow), Chr(10), ""), Chr(13), ""), Chr(1), "")
                If blnGroup Then
                    '�����е����⴦��,�����ڲ���¼���Ĵ���϶�
                    '##########################################
                    '��������
                    If Val(VsfData.TextMatrix(lngStart + intRow, mlngDemo)) > 0 Then
                        If Val(VsfData.TextMatrix(lngStart + intRow, mlngRecord)) > 0 Then
                            '�������޸ĲŽ�������̣�ȡ������¼��ʵ��ʱ��
                            If mblnDateAd Then
                                strDate = Format(VsfData.TextMatrix(lngStart + intRow, mlngActTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStart + intRow, mlngActTime), "MM")
                            Else
                                strDate = Mid(VsfData.TextMatrix(lngStart + intRow, mlngActTime), 1, 10)
                            End If
                            strTime = Mid(VsfData.TextMatrix(lngStart + intRow, mlngActTime), 12, 5)
                        Else
                            '����ʱ���������
                            strDate = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngDate)
                            strTime = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngTime)
                        End If
                    Else
                        '��ͨ����
                        strDate = VsfData.TextMatrix(lngStart + intRow, mlngDate)
                        strTime = VsfData.TextMatrix(lngStart + intRow, mlngTime)
                    End If
                    
                    '1\����
                    strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|ɾ��"
                    If mlngDate <> -1 Then
                        strKey = mintҳ�� & "," & lngStart + intRow & "," & mlngDate
                        strValue = strKey & "|" & mintҳ�� & "|" & lngStart + intRow & "|" & mlngDate & "|" & _
                            Val(VsfData.TextMatrix(lngStart + intRow, mlngRecord)) & "|" & strDate & "|" & VsfData.TextMatrix(lngStart + intRow, mlngDemo) & "|0"
                        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                    End If
                    '2\ʱ��
                    strKey = mintҳ�� & "," & lngStart + intRow & "," & mlngTime
                    strValue = strKey & "|" & mintҳ�� & "|" & lngStart + intRow & "|" & mlngTime & "|" & _
                        Val(VsfData.TextMatrix(lngStart + intRow, mlngRecord)) & "|" & strTime & "|" & _
                        VsfData.TextMatrix(lngStart + intRow, mlngDemo) & "|0"
                    Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                    
                    strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|ɾ��"
                    strKey = mintҳ�� & "," & lngStart + intRow & "," & VsfData.COL
                    strValue = strKey & "|" & mintҳ�� & "|" & lngStart + intRow & "|" & VsfData.COL & "|" & _
                        Val(VsfData.TextMatrix(lngStart + intRow, mlngRecord)) & "|" & Replace(Replace(Replace(arrData(intRow), Chr(10), ""), Chr(13), ""), Chr(1), "") & "||0"
                    Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                    '##########################################
                End If
            Next
            For intRow = intCount + 1 To lngMutilRows - 1
                VsfData.TextMatrix(lngStart + intRow, VsfData.COL) = ""
                '�����е����⴦��,�����ڲ���¼���Ĵ���϶�
                '##########################################
                '��������
                If blnGroup Then
                    If Val(VsfData.TextMatrix(lngStart + intRow, mlngDemo)) > 0 Then
                        If Val(VsfData.TextMatrix(lngStart + intRow, mlngRecord)) > 0 Then
                            '�������޸ĲŽ�������̣�ȡ������¼��ʵ��ʱ��
                            If mblnDateAd Then
                                strDate = Format(VsfData.TextMatrix(lngStart + intRow, mlngActTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStart + intRow, mlngActTime), "MM")
                            Else
                                strDate = Mid(VsfData.TextMatrix(lngStart + intRow, mlngActTime), 1, 10)
                            End If
                            strTime = Mid(VsfData.TextMatrix(lngStart + intRow, mlngActTime), 12, 5)
                        Else
                            '����ʱ���������
                            strDate = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngDate)
                            strTime = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngTime)
                        End If
                    Else
                        '��ͨ����
                        strDate = VsfData.TextMatrix(lngStart + intRow, mlngDate)
                        strTime = VsfData.TextMatrix(lngStart + intRow, mlngTime)
                    End If
                    
                    '1\����
                    strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|ɾ��"
                    If mlngDate <> -1 Then
                        strKey = mintҳ�� & "," & lngStart + intRow & "," & mlngDate
                        strValue = strKey & "|" & mintҳ�� & "|" & lngStart + intRow & "|" & mlngDate & "|" & _
                            Val(VsfData.TextMatrix(lngStart + intRow, mlngRecord)) & "|" & strDate & "|" & VsfData.TextMatrix(lngStart + intRow, mlngDemo) & "|0"
                        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                    End If
                    '2\ʱ��
                    strKey = mintҳ�� & "," & lngStart + intRow & "," & mlngTime
                    strValue = strKey & "|" & mintҳ�� & "|" & lngStart + intRow & "|" & mlngTime & "|" & _
                        Val(VsfData.TextMatrix(lngStart + intRow, mlngRecord)) & "|" & strTime & "|" & _
                        VsfData.TextMatrix(lngStart + intRow, mlngDemo) & "|0"
                    Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                    
                    strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|ɾ��"
                    strKey = mintҳ�� & "," & lngStart + intRow & "," & VsfData.COL
                    strValue = strKey & "|" & mintҳ�� & "|" & lngStart + intRow & "|" & VsfData.COL & "|" & _
                        Val(VsfData.TextMatrix(lngStart + intRow, mlngRecord)) & "|" & "" & "||1"
                    Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                End If
                '##########################################
            Next
            
            '����������������д������,intNULL��¼���һ����Ϊ���е��к�
            intNULL = lngStart + lngMutilRows - 1
            For intRow = lngMutilRows To 1 Step -1
                blnNULL = True
                For intCount = 0 To VsfData.Cols - 1
                    If Not VsfData.ColHidden(intCount) Then
                        If VsfData.TextMatrix(intRow + lngStart - 1, intCount) <> "" Then
                            blnNULL = False
                            Exit For
                        End If
                    End If
                Next
                
                If Not blnNULL Then Exit For
                intNULL = intNULL - 1
            Next
            '������д�����
            If Not blnGroup Then
                For intRow = lngStart To intNULL
                    VsfData.TextMatrix(intRow, mlngRowCount) = (intNULL - lngStart + 1) & "|" & intRow - lngStart + 1
                    VsfData.TextMatrix(intRow, mlngRowCurrent) = (intNULL - lngStart + 1)
                Next
            Else '�������Ա��������ɾ��ʱ��������к�
                For intRow = intNULL + 1 To lngStart + lngMutilRows - 1
                    If Val(VsfData.TextMatrix(intRow, mlngRecord)) > 0 Then intNULL = intNULL + 1
                Next intRow
            End If
            For intRow = intNULL + 1 To lngStart + lngMutilRows - 1
                VsfData.TextMatrix(intRow, mlngRowCount) = ""
                VsfData.TextMatrix(intRow, mlngRowCurrent) = ""
                VsfData.TextMatrix(intRow, mlngRecord) = ""
            Next
        End If
        
        '���кŷ����仯����ͬ������mrsCellMap�д��ڸ��кŵ��к�����
        If lngDeff <> 0 Then
            If Not blnGroup Then
                Call CellMap_Update(lngStart, lngDeff)
            Else
                Call CellMap_Update(lngMutilRows, lngDeff)  '���������ݴ����һ����ϸ��֮��ʼ����
            End If
        End If
        
        If mstrData <> strReturn Then
            mblnChange = True
            
            'ͬ������������ʱ���е�����
            If Val(VsfData.TextMatrix(lngStart, mlngCollectType)) >= 0 Then
                If Val(VsfData.TextMatrix(lngStart, mlngDemo)) > 0 Then
                    If Val(VsfData.TextMatrix(lngStart, mlngRecord)) > 0 Then
                        '�������޸ĲŽ�������̣�ȡ������¼��ʵ��ʱ��
                        If mblnDateAd Then
                            strDate = Format(VsfData.TextMatrix(lngStart, mlngActTime), "DD") & "/" & Format(VsfData.TextMatrix(lngStart, mlngActTime), "MM")
                        Else
                            strDate = Mid(VsfData.TextMatrix(lngStart, mlngActTime), 1, 10)
                        End If
                        strTime = Mid(VsfData.TextMatrix(lngStart, mlngActTime), 12, 5)
                    Else
                        '����ʱ���������
                        strDate = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngDate)
                        strTime = VsfData.TextMatrix(VsfData.ROW - Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) + 1, mlngTime)
                    End If
                Else
                    '��ͨ����
                    strDate = VsfData.TextMatrix(lngStart, mlngDate)
                    strTime = VsfData.TextMatrix(lngStart, mlngTime)
                End If
                
                '1\����
                strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|ɾ��"
                If mlngDate <> -1 Then
                    strKey = mintҳ�� & "," & lngStart & "," & mlngDate
                    strValue = strKey & "|" & mintҳ�� & "|" & lngStart & "|" & mlngDate & "|" & _
                        Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & strDate & "|" & VsfData.TextMatrix(lngStart, mlngDemo) & "|0"
                    Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                End If
                '2\ʱ��
                strKey = mintҳ�� & "," & lngStart & "," & mlngTime
                strValue = strKey & "|" & mintҳ�� & "|" & lngStart & "|" & mlngTime & "|" & _
                    Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & strTime & "|" & _
                    VsfData.TextMatrix(lngStart, mlngDemo) & "|0"
                Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
            Else
                strField = "ID|ҳ��|�к�|�к�|��¼ID|����|����|ɾ��"
                strKey = mintҳ�� & "," & lngStart & "," & mlngDate
                strValue = strKey & "|" & mintҳ�� & "|" & lngStart & "|" & mlngDate & "|" & Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & _
                        VsfData.TextMatrix(lngStart, mlngCollectText) & ";" & VsfData.TextMatrix(lngStart, mlngCollectType) & ";" & _
                        VsfData.TextMatrix(lngStart, mlngCollectStyle) & ";" & VsfData.TextMatrix(lngStart, mlngCollectDay) & ";" & _
                    VsfData.TextMatrix(lngStart, mlngCollectStart) & ";" & VsfData.TextMatrix(lngStart, mlngCollectEnd) & "|1|0"
                Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
            End If
            
            If Not blnGroup Then
                '��¼�û��޸Ĺ��ĵ�Ԫ��
                If InStr(1, "," & mstrCatercorner & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
                    strPart = GetActivePart(VsfData.COL, 0)
                Else
                    strPart = GetActivePart(VsfData.COL, 0) & "/" & GetActivePart(VsfData.COL, 1)
                End If
                
                strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|ɾ��"
                strKey = mintҳ�� & "," & lngStart & "," & VsfData.COL
                strValue = strKey & "|" & mintҳ�� & "|" & lngStart & "|" & VsfData.COL & "|" & _
                    Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & strReturn & "|" & strPart & "|" & IIf(strReturn = "", "1", "0")
                Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
            End If
        End If
        RaiseEvent AfterDataChanged(mblnChange Or mblnVerify)
    End If
    
    If blnNext Then
toMoveNextCol:
        If VsfData.COL < mlngNoEditor - 1 Then       '�����¼���϶��л�ʿǩ����
            VsfData.COL = VsfData.COL + 1
            If VsfData.ColWidth(VsfData.COL) = 0 Or VsfData.ColHidden(VsfData.COL) Then GoTo toMoveNextCol
        Else
toMoveNextRow:
            '������һ��
            If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") <> 0 Then
                intRow = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
                intRow = intRow - Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(1)) + 1
            Else
                intRow = 1
            End If
            If VsfData.ROW + intRow < VsfData.Rows Then
                VsfData.ROW = VsfData.ROW + intRow
            End If
            If VsfData.RowHidden(VsfData.ROW) Then GoTo toMoveNextRow
            VsfData.COL = IIf(mlngDate > 0 And mblnDate = True, mlngDate, mlngTime)
            If mstrClassRow <> "" Then VsfData.COL = VsfData.COL + 2
        End If
    Else
toMovePrevCol:
        If VsfData.COL > mlngDate Then      '�����¼���϶��л�ʿǩ����
            VsfData.COL = VsfData.COL - 1
            If VsfData.ColWidth(VsfData.COL) = 0 Or VsfData.ColHidden(VsfData.COL) Then GoTo toMovePrevCol
        Else
toMovePrevRow:
'            '������һ��
'            intRow = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
'            intRow = intRow - Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(1)) + 1
'            If VsfData.ROW + intRow < VsfData.Rows Then
'                VsfData.ROW = VsfData.ROW + intRow
'            End If
'            If VsfData.RowHidden(VsfData.ROW) Then GoTo toMovePrevRow
'            VsfData.COL = IIf(mlngDate > 0, mlngDate, mlngTime)
        End If
    End If
    If VsfData.ColIsVisible(VsfData.COL) = False Then
        VsfData.LeftCol = VsfData.COL
    End If
    If VsfData.RowIsVisible(VsfData.ROW) = False Then
        VsfData.TopRow = VsfData.ROW
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function GetStartRow(ByVal lngRow As Long) As Long
    Dim lngStart As Long
    Dim lngCurRows As Long, lngRows As Long
    '��ȡ������ʼ��,������ҳ�򷵻�0
    '�����ҳδ��ʾȫ,��˵��������ҳ,Ҳ����0
    '���������������������в�������
    
    lngRows = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0))    '������
    lngCurRows = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(1)) '��ǰ��
    If lngCurRows = 1 Then
        GetStartRow = lngRow
        Exit Function
    End If
    
    'Ѱ����ʼ��
    For lngRow = lngRow To 3 Step -1
        If VsfData.TextMatrix(lngRow, mlngRowCount) = lngRows & "|1" Then
            lngStart = lngRow
            Exit For
        End If
    Next
    
    GetStartRow = lngStart
End Function

Private Function GetMutilData(ByVal lngRow As Long, ByVal lngCol As Long, dblTop As Long, dblHeight As Long) As String
    Dim lngCurRow As Long
    Dim lngCount As Long
    Dim lngStart As Long    '��ʼ��
    Dim lngRecordId As Long
    Dim strReturn As String
    Dim blnAdjust As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '���ص�һ�е�����
    '������ֱ��ȡ������ʱ��������ҳ��ʾȫ��ƴ�ӣ�����ӿ��ж�ȡ
    
    If VsfData.TextMatrix(lngRow, mlngRowCount) = "" Then
        GetMutilData = VsfData.TextMatrix(lngRow, lngCol)
        Exit Function
    End If
    lngRecordId = Val(VsfData.TextMatrix(lngRow, mlngRecord))
    lngCount = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0))
    lngCurRow = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(1))
    
    If lngCount > 1 Then
        lngStart = GetStartRow(lngRow)
    Else
        lngStart = lngRow
    End If
    If lngRecordId <> 0 And (lngStart = 0 Or lngStart + lngCount > VsfData.Rows) Then   'ҳ��Ч��=�̶�������+��ͷ
        '�����ݿ�����ȡ
        Call SQLCombination(lngRecordId)
        gstrSQL = mstrSQL
        Call SQLDIY(gstrSQL)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������", mlng�ļ�ID, mlng����ID, mlng��ҳID, mintӤ��, mintҳ��, cboBaby.ItemData(cboBaby.ListIndex), lngRecordId)
        strReturn = NVL(rsTemp.Fields(lngCol).Value)
        If lngStart = 0 Then lngStart = 3       '���δ�ҵ���ʼ�����趨Ϊ��1��
        blnAdjust = True
    Else
        For lngRow = lngStart To lngStart + lngCount - 1
            strReturn = strReturn & VsfData.TextMatrix(lngRow, lngCol)
        Next
    End If
    
'    'У���и�(�п���ʵ������ռ5�ж���ǰҳ��ֻ��ʾ��3��,����3����ʾ�������Բ�ȫ,���Ի�����ԭ�����и���ʾ����,���´�������)
'    If blnAdjust Then
'        If lngStart = 3 Then
'            lngCurRow = Val(Split(VsfData.TextMatrix(lngStart, mlngRowCount), "|")(1))
'            lngCount = lngCount - lngCurRow + 1
'        Else
'            lngCount = mlngPageRows +mlngOverrunRows + VsfData.FixedRows - lngStart
'        End If
'    End If
    'ȡ�и�
    VsfData.ROW = lngStart
    dblHeight = lngCount * VsfData.RowHeightMin + 20
    dblTop = VsfData.Top + VsfData.CellTop
    
    GetMutilData = strReturn
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ShowInput(Optional ByVal intCOl As Integer = -1, Optional ByVal strCellData As String = "", Optional ByVal blnAnalyse As Boolean = False) As String
    Dim arrData, arrValue
    Dim lngOrder As Long
    Dim i As Integer, j As Integer, intPos As Integer, intIndex As Integer
    Dim strFormat As String, strText As String, strValue As String  '��ʽ��,���ݴ�,��ֵ��
    Dim strOrders As String, strTypes As String, strBounds As String, strLen As String, strName As String
    Dim strCurDate As String
    Const txtHeight = 300
    On Error GoTo errHand
    
    '�����ļ��������ģ����Ҫ����:
    '1��һ�а�һ����Ŀ�Ĳ��ù�
    '2��һ�а�������Ŀ�ģ�Ѫѹ����ɶԣ�Ҫô����¼�룬Ҫô����ѡ�񣬲���������֣�Ҳ��������ֵ�ѡ����ѡ
    '3��һ�а󶨶����Ŀ�ģ�ֻ����¼����Ŀ
    '���������������ƣ�ֻȡ��һ����Ŀ�����ʼ���
    
    '����Ǳ��洦�����������´���
    If intCOl = -1 Then intCOl = VsfData.COL
    If blnAnalyse Then
        strText = strCellData
    Else
        'ȡ��ǰ��Ԫ�������
        CellRect.Left = VsfData.CellLeft + VsfData.Left
        CellRect.Top = VsfData.CellTop + VsfData.Top
        CellRect.Bottom = VsfData.CellHeight + 20
        CellRect.Right = VsfData.CellWidth + 20
        strText = GetMutilData(VsfData.ROW, intCOl, CellRect.Top, CellRect.Bottom)
    End If
    mstrData = strText
    mintType = 0
    intIndex = 0
    
    'ȡ��ǰ�еİ���Ŀ
    intPos = 1
    mrsSelItems.Filter = "��=" & intCOl - cHideCols
    Do While Not mrsSelItems.EOF
        lngOrder = mrsSelItems!��Ŀ���
        If lngOrder = 0 Then
            strLen = 0
            strValue = strText
            Exit Do
        End If
        
        '��Ŀ��ʾ:2��ѡ;3-��ѡ;4-����;5-ѡ��
        '��Ŀֵ��:��Ŀ��ʾΪ0-��ʾ��Сֵ;���ֵ;��Ŀ��ʾΪ2,3-��ʾ��ĿA;��ĿB,ǰ�й��ı�ʾȱʡ��
        strFormat = NVL(mrsSelItems!��ʽ)
        strOrders = strOrders & "," & lngOrder
        If lngOrder <> 0 Then
            mrsItems.Filter = "��Ŀ���=" & lngOrder
            strName = strName & "," & mrsItems!��Ŀ����
            strLen = strLen & "," & mrsItems!��Ŀ���� & ";" & NVL(mrsItems!��ĿС��)
            strTypes = strTypes & "," & mrsItems!��Ŀ��ʾ
            strBounds = strBounds & "," & mrsItems!��Ŀֵ��
            strValue = strValue & "'" & SubstrVal(strText, strFormat, GetActivePart(intCOl, intIndex) & mrsItems!��Ŀ����, intPos)
            
            Select Case mrsItems!��Ŀ��ʾ
            Case 0  '�ı�¼����
                If mrsSelItems.RecordCount = 2 Then
                    mintType = 4
                ElseIf mrsSelItems.RecordCount > 2 Then
                    mintType = 6
                End If
            Case 2  '��ѡ
                If mrsSelItems.RecordCount = 1 Then
                    mintType = 1
                ElseIf mrsSelItems.RecordCount = 2 Then
                    mintType = 7
                End If
            Case 3  '��ѡ
                mintType = 2
            Case 4  '����
                If mrsSelItems.RecordCount = 2 Then
                    mintType = 4
                ElseIf mrsSelItems.RecordCount > 2 Then
                    mintType = 6
                End If
            Case 5  'ѡ��
                If mrsSelItems.RecordCount = 1 Then
                    mintType = 3
                Else
                    mintType = 5
                End If
            End Select
        Else
            strTypes = strTypes & ","
            strBounds = strBounds & ","
            strLen = strLen & ","
            strName = strName & ","
        End If
        
        intIndex = intIndex + 1
        mrsSelItems.MoveNext
    Loop
    If strOrders <> "" Then
        strOrders = Mid(strOrders, 2)
        strName = Mid(strName, 2)
        strLen = Mid(strLen, 2)
        strTypes = Mid(strTypes, 2)
        strBounds = Mid(strBounds, 2)
        strValue = Mid(strValue, 2)
    End If
    mrsSelItems.Filter = 0
    mrsItems.Filter = 0
    
    If blnAnalyse Then
        ShowInput = strOrders & "||" & strValue
        Exit Function
    End If
    
    '���4����У��,�����ͷ�ı�����/����Ϊ6
    If mintType = 4 Then
        If Not IsDiagonal(intCOl) Then
            mintType = 6
        End If
    End If
    
    '�жϵ�ǰ�е�����
    'mintType:0=�ı���¼��;1=��ѡ;2=��ѡ;3=ѡ��;4-Ѫѹ��һ�а���������Ŀ,���ʽ����Ѫѹ��������Ŀ;5=һ�а���������Ŀ�Ҿ���ѡ����Ŀ;
    '6=һ�а�2����������Ŀ,�ֹ�¼��,7=һ�а���������ѡ��Ŀ
    arrValue = Split(strValue & "'", "'")
    Select Case mintType
    Case 0, 3
        With picInput
            .Left = CellRect.Left
            .Top = CellRect.Top
            .Width = CellRect.Right
            .Height = CellRect.Bottom
            If .Height + .Top + picMain.Top > ScaleHeight Then
                .Top = ScaleHeight - picMain.Top - .Height
            End If
            .Visible = True
        End With
        If mintType = 0 Then
            txtInput.Visible = True
            If Val(strLen) <> 0 And Val(strOrders) <> 10 Then
                txtInput.MaxLength = Val(Split(strLen, ";")(0)) + IIf(Val(Split(strLen, ";")(1)) = 0, 0, 1) 'С��λ��Ҫ����С����
            Else
                txtInput.MaxLength = 0
            End If
            txtInput.Tag = lngOrder
        Else
            txtInput.Visible = False
        End If
        With txtInput
            .Top = 0
            .Text = strValue
            .Width = CellRect.Right
            .Height = CellRect.Bottom
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Width = .Width - (180 + IIf(mblnBlowup, 180 * 1 / 3, 0)) / 2 '����9��ʱ��ȥ90,����Խ��۳��ı߾�ԽС,�Ա�֤�ı��������ʵ��һ��
        End With
        With lblInput
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Height = CellRect.Bottom
            .Width = CellRect.Right
            .Top = 50
            .Tag = lngOrder
            .Caption = strValue
            .Visible = (mintType = 3)
        End With
        
        '��������ڻ�ʱ���У��趨�̶�ֵ
        strCurDate = zlDatabase.Currentdate
        If mintType = 0 And txtInput.Text = "" Then
            If intCOl = mlngDate Or intCOl = mlngTime Then
                txtInput.Text = Format(strCurDate, "YYYY-MM-DD HH:mm")
                If Format(strCurDate, "YYYY-MM-DD HH:mm") >= Format(mstr����ʱ��, "YYYY-MM-DD HH:mm") Then
                    txtInput.Text = Format(mstr����ʱ��, "YYYY-MM-DD HH:mm")
                End If
                If Format(strCurDate, "YYYY-MM-DD HH:mm") <= Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm") Then
                    txtInput.Text = Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm")
                End If
            End If
            If intCOl = mlngDate Then
                If mblnDateAd Then
                    txtInput.Text = Format(txtInput.Text, "d-M")
                    txtInput.Text = Replace(txtInput.Text, "-", "/")
                Else
                    txtInput.Text = Format(txtInput.Text, "yyyy-MM-dd")
                End If
            ElseIf intCOl = mlngTime Then
                txtInput.Text = Format(txtInput.Text, "HH:mm")
            End If
        End If
    Case 1, 2
        '��������
        lstSelect(mintType - 1).Clear
        If mintType = 1 Then
            lstSelect(mintType - 1).AddItem "0-"
            If mlngProduce = intCOl Then lstSelect(mintType - 1).ListIndex = 0
        End If
        arrData = Split(strBounds, ";")
        j = UBound(arrData)
        For i = 0 To j
            If arrData(i) <> "" Then
                If Mid(arrData(i), 1, 1) = "��" Then
                    lstSelect(mintType - 1).AddItem lstSelect(mintType - 1).NewIndex + 1 & "-" & Mid(arrData(i), 2)
                    If strText = "" And lstSelect(mintType - 1).ListIndex = -1 Then lstSelect(mintType - 1).ListIndex = lstSelect(mintType - 1).NewIndex
                Else
                    lstSelect(mintType - 1).AddItem lstSelect(mintType - 1).NewIndex + 1 & "-" & arrData(i)
                End If
            End If
        Next
        '��ѡ����¼�����ݵ������
        If strValue <> "" Then
            strValue = Replace(strValue, vbCrLf, "")
            j = lstSelect(mintType - 1).ListCount - 1
            For i = 0 To j
                If InStr(1, "," & strValue & ",", "," & Split(lstSelect(mintType - 1).List(i), "-")(1) & ",") <> 0 Then
                    lstSelect(mintType - 1).Selected(i) = True
                End If
            Next
        End If
        '��ʾ
        With lstSelect(mintType - 1)
            .Left = CellRect.Left
            .Top = CellRect.Top
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Height = .ListCount * 300
            If .Height < CellRect.Bottom Then .Height = CellRect.Bottom
            .Width = LenB(StrConv(.List(.ListCount \ 2), vbFromUnicode)) * 120 + 500    '���м���ĳ���Ϊ����
            If .Width < CellRect.Right Then .Width = CellRect.Right
            If .Height + .Top + picMain.Top > ScaleHeight Then
                .Top = ScaleHeight - picMain.Top - .Height
            End If
            .Tag = lngOrder
            .Visible = True
        End With
    Case 4, 5
        With picDouble
            .Left = CellRect.Left
            .Top = CellRect.Top
            .Height = CellRect.Bottom
            If .Height < 280 Then .Height = 280
            .Width = CellRect.Right
            If .Width < 820 Then .Width = 820
            If .Height + .Top + picMain.Top > ScaleHeight Then
                .Top = ScaleHeight - picMain.Top - .Height
            End If
            .Visible = True
        End With
        lblSplit.FontName = VsfData.FontName
        lblSplit.FontSize = VsfData.FontSize
        lblSplit.Left = (picDouble.Width - lblSplit.Width) / 2
        If mblnBlowup Then
            lblSplit.Width = 150
        Else
            lblSplit.Width = 105
        End If
        
        With txtUpInput
            .Text = arrValue(0)
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Width = (picDouble.Width - lblSplit.Width) * 0.4
            .ZOrder IIf(mintType = 4, 0, 1)
            .Locked = Not (mintType = 4)
            .Tag = Split(strOrders, ",")(0)
        End With
        With picUpInput
            .Left = txtUpInput.Left
            .Width = txtUpInput.Width
            .Height = CellRect.Bottom
            .ZOrder IIf(mintType = 5, 0, 1)
            .Tag = Split(strOrders, ",")(0)
        End With
        With lblUpInput
            .Alignment = 2
            .Caption = arrValue(0)
            .Left = 0
            .Top = 50
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Width = txtUpInput.Width
            .Height = CellRect.Bottom
            .Tag = Split(strOrders, ",")(0)
        End With
        With txtDnInput
            .Text = arrValue(1)
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Left = lblSplit.Left + lblSplit.Width
            .Width = picDouble.Width - .Left
            .ZOrder IIf(mintType = 4, 0, 1)
            .Locked = Not (mintType = 4)
            .Tag = Split(strOrders, ",")(1)
        End With
        With picDnInput
            .Left = txtDnInput.Left
            .Height = CellRect.Bottom
            .Width = txtDnInput.Width
            .ZOrder IIf(mintType = 5, 0, 1)
            .Tag = Split(strOrders, ",")(1)
        End With
        With lblDnInput
            .Alignment = 2
            .Caption = arrValue(1)
            .Left = 0
            .Top = 50
            .Height = CellRect.Bottom
            .Width = txtDnInput.Width
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Tag = Split(strOrders, ",")(1)
        End With
        
        If mintType = 4 Then
            If strLen <> "" Then txtUpInput.MaxLength = Val(Split(Split(strLen, ",")(0), ";")(0)) + IIf(Val(Split(Split(strLen, ",")(0), ";")(1)) = 0, 0, 1) 'С��λ��Ҫ����С����
            If strLen <> "" Then txtDnInput.MaxLength = Val(Split(Split(strLen, ",")(1), ";")(0)) + IIf(Val(Split(Split(strLen, ",")(1), ";")(1)) = 0, 0, 1) 'С��λ��Ҫ����С����
        End If
    Case 6
        '��ɾ����ǰ�Ŀؼ�
        j = txt.Count - 1
        For i = 1 To j
            Unload lbl(i)
            Unload txt(i)
        Next
        '�趨����
        With picMutilInput
            .Left = CellRect.Left
            .Top = CellRect.Top
            .Width = IIf(CellRect.Right < 1600, 1600, CellRect.Right)
        End With
        '��ȱʡ�ؼ���ֵ
        arrData = Split(strOrders, ",")
        j = UBound(arrData)
        lbl(0).Top = 130
        lbl(0).Caption = Split(strName, ",")(0)
        lbl(0).FontName = VsfData.FontName
        lbl(0).FontSize = VsfData.FontSize
        txt(0).Tag = arrData(0)
        txt(0).FontName = VsfData.FontName
        txt(0).FontSize = VsfData.FontSize
        txt(0).Width = picMutilInput.Width - txt(0).Left - 100
        txt(0).MaxLength = Val(Split(Split(strLen, ",")(0), ";")(0)) + IIf(Val(Split(Split(strLen, ",")(0), ";")(1)) = 0, 0, 1)  'С��λ��Ҫ����С����
        txt(0).Text = arrValue(0)
        If Not mblnBlowup Then
            txt(0).Height = 225
        End If
        
        '���ؿؼ�
        For i = 1 To j
            Load lbl(i)
            With lbl(i)
                .Caption = Split(strName, ",")(i)
                .Left = lbl(0).Left + lbl(0).Width - .Width
                .Top = lbl(i - 1).Top + txtHeight + IIf(mblnBlowup, txtHeight * 1 / 3, 0)
                .Visible = True
            End With
            Load txt(i)
            With txt(i)
                .TabIndex = txt(i - 1).TabIndex + 1
                .Left = txt(0).Left
                .Top = txt(i - 1).Top + txtHeight + IIf(mblnBlowup, txtHeight * 1 / 3, 0)
                .Tag = arrData(i)
                If strLen <> "" Then
                    .MaxLength = Val(Split(Split(strLen, ",")(i), ";")(0)) + IIf(Val(Split(Split(strLen, ",")(i), ";")(1)) = 0, 0, 1) 'С��λ��Ҫ����С����
                End If
                .Text = arrValue(i)
                .Visible = True
            End With
        Next
        
        With picMutilInput
            .Height = txt(j).Top + txt(j).Height + 120
            If .Height < CellRect.Bottom Then .Height = CellRect.Bottom
            If .Height + .Top + picMain.Top > ScaleHeight Then
                .Top = ScaleHeight - picMain.Top - .Height
            End If
            .Visible = True
        End With
    Case 7
        cboChoose(0).Clear
        arrData = Split(Split(strBounds, ",")(0), ";")
        j = UBound(arrData)
        For i = 0 To j
            If Mid(arrData(i), 1, 1) = "��" Then
                cboChoose(0).AddItem Mid(arrData(i), 2)
                If strValue = "" Then
                    cboChoose(0).ListIndex = i
                Else
                    If Mid(arrData(i), 2) = Split(strValue, "'")(0) Then
                        cboChoose(0).ListIndex = i
                    End If
                End If
            Else
                cboChoose(0).AddItem arrData(i)
                If strValue <> "" Then
                    If arrData(i) = Split(strValue, "'")(0) Then
                        cboChoose(0).ListIndex = i
                    End If
                End If
            End If
        Next
        cboChoose(1).Clear
        arrData = Split(Split(strBounds, ",")(1), ";")
        j = UBound(arrData)
        For i = 0 To j
            If Mid(arrData(i), 1, 1) = "��" Then
                cboChoose(1).AddItem Mid(arrData(i), 2)
                If strValue = "" Then
                    cboChoose(1).ListIndex = i
                Else
                    If Mid(arrData(i), 2) = Split(strValue, "'")(1) Then
                        cboChoose(1).ListIndex = i
                    End If
                End If
            Else
                cboChoose(1).AddItem arrData(i)
                If strValue <> "" Then
                    If arrData(i) = Split(strValue, "'")(1) Then
                        cboChoose(1).ListIndex = i
                    End If
                End If
            End If
        Next
        With picDoubleChoose
            .Left = CellRect.Left
            .Top = CellRect.Top
            .Height = CellRect.Bottom
            If .Height < 280 Then .Height = 280
            .Width = CellRect.Right
            If .Width < 820 Then .Width = 820
            If .Height + .Top + picMain.Top > ScaleHeight Then
                .Top = ScaleHeight - picMain.Top - .Height
            End If
            .Visible = True
        End With
        lblSplit.FontName = VsfData.FontName
        lblSplit.FontSize = VsfData.FontSize
        lblSplit.Left = (picDoubleChoose.Width - lblSplit.Width) / 2
        If mblnBlowup Then
            lblSplit.Width = 150
        Else
            lblSplit.Width = 105
        End If
        
        cboChoose(0).SetFocus
    End Select
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub CheckFormat(ByVal strNames As String, ByVal strFormat As String)
    '�����ʽ��Ѫѹ�ķ�ʽ��ͬ,����ʽ����Ϊ6
    
    'ȥ��ǰ׺����жԱ�
    strFormat = Mid(strFormat, InStr(1, strFormat, "["))
    strFormat = Replace(strFormat, "[", "")
    strFormat = Replace(strFormat, "]", "")
    If Not (strFormat Like Split(strNames, ",")(0) & "/}*" Or strFormat Like "{/*" & Split(strNames, ",")(1)) Then
        mintType = 6
    End If
End Sub

Private Function IsDiagonal(ByVal intCOl As Integer) As Boolean
    Dim arrCol, arrData
    Dim intDo As Integer, intCount As Integer
    '�ж�ָ�����Ƿ��������жԽ��ߣ�mstrColWidth�ĸ�ʽ��765`11`1`1,765`11`2`1,...����������`�������`�жԽ��ߣ�
    
    IsDiagonal = (InStr(1, "," & mstrCatercorner & ",", "," & intCOl - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0)
End Function

Private Sub ISAssistant(ByVal lngOrder As Long, ByVal objTXT As TextBox)
    Dim intIndex As Integer
    Dim objParent As Object
    Dim intRow As Integer, intCount As Integer, i As Integer
    Dim strText As String
    Dim arrData
    '������Ŀ�ĳ��Ⱦ����Ƿ�������дʾ�ѡ��
    mblnEditAssistant = False
    cmdWord.Visible = mblnEditAssistant
    
    mrsItems.Filter = "��Ŀ���=" & lngOrder
    If mrsItems.RecordCount = 0 Then
        mrsItems.Filter = 0
        Exit Sub
    End If
    mblnEditAssistant = (mrsItems!��Ŀ���� = 1 And mrsItems!��Ŀ���� >= 100) And Val(VsfData.TextMatrix(VsfData.ROW, mlngDemo)) <= "1"
    mrsItems.Filter = 0
    
    '�������ʾ�ѡ��,��ʾ����λ
    If mblnEditAssistant Then
        If UCase(objTXT.Name) = "TXTINPUT" Then
            intIndex = -1 '��ʾtxtInput
            Set objParent = picInput
        Else
            intIndex = objTXT.Index
            Set objParent = picMutilInput
        End If
        With cmdWord
            .Tag = intIndex
            .Top = objParent.Top + objTXT.Top + 25
            .Left = objParent.Left + objTXT.Left + objTXT.Width - .Width + 25
            .Visible = True
        End With
        strText = ""
        intCount = 0
        'Ϊ������ʱ��ѡ��������ʼ�У��༭������ʾ���д��ı���
        If VsfData.TextMatrix(VsfData.ROW, mlngDemo) = "1" Then
            For intRow = VsfData.ROW To VsfData.Rows - 1
                If Val(VsfData.TextMatrix(intRow, mlngDemo)) <= 1 And intCount = 1 Then Exit For
                intCount = 1
                strText = strText & VsfData.TextMatrix(intRow, VsfData.COL)
                'strText = strText & IIf(intRow > VsfData.ROW And strText <> "", vbCrLf, "") & Replace(VsfData.TextMatrix(intRow, VsfData.COL), vbCrLf, "")
            Next intRow
            '׼����ֵ
            With txtLength
                '������ʱ���еĿ�Ȳ���,Ϊ�˱��ⷵ�ض���,ǿ������Ϊ5000
                .Width = VsfData.CellWidth
                .Text = strText
                .FontName = VsfData.CellFontName
                .FontSize = VsfData.CellFontSize
                .FontBold = VsfData.CellFontBold
                .FontItalic = VsfData.CellFontItalic
            End With
            arrData = GetData(txtLength.Text)
            intCount = UBound(arrData)
            strText = ""
            For i = 0 To intCount
                strText = strText & CStr(arrData(i))
            Next i
            intRow = intRow - VsfData.ROW
            picInput.Height = intRow * VsfData.RowHeightMin + 20
            If picInput.Height + picInput.Top + picMain.Top > ScaleHeight Then
                picInput.Top = ScaleHeight - picMain.Top - picInput.Height
            End If
            txtInput.Height = picInput.Height
            txtInput.Text = strText
            mstrData = strText
            txtInput.SelStart = 0
            txtInput.SelLength = 100
            lblInput.Height = picInput.Height
        End If
    End If
End Sub

Private Sub FillPage(Optional ByVal blnLocate As Boolean = True)
    Dim lngRow As Long, lngRows As Long, lngCount As Long, lngData As Long
    '��֤ÿҳ��Ч������
    
    lngRows = VsfData.Rows - 1
    For lngRow = VsfData.FixedRows To lngRows
        If VsfData.TextMatrix(lngRow, mlngRowCount) <> "" Then
            lngData = lngData + 1
        End If
        If Not VsfData.RowHidden(lngRow) Then
            lngCount = lngCount + 1
        End If
    Next
    
    If lngCount < mlngPageRows + mlngOverrunRows Then
        VsfData.Rows = VsfData.Rows + (mlngPageRows + mlngOverrunRows - lngCount)
'    Else
'        VsfData.Rows = VsfData.FixedRows + (mlngPageRows + mlngOverrunRows)
    End If
    If mintҳ�� <= mint����ҳ And mintҳ�� > 0 Then VsfData.Rows = VsfData.Rows - Val(CStr(mlngLitterRows(mintҳ��)))
    
    On Error Resume Next
    If Not blnLocate Then Exit Sub
    If lngData + VsfData.FixedRows < VsfData.Rows - 1 Then
        VsfData.ROW = lngData + VsfData.FixedRows
    Else
        VsfData.ROW = VsfData.FixedRows
    End If
    If Not VsfData.RowIsVisible(VsfData.ROW) Then VsfData.TopRow = VsfData.ROW
    '������һ���ǿ���,�����༭״̬
    If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "" Then
        VsfData.SetFocus
        VsfData.COL = mlngDate
    End If
End Sub

Public Function GetSynItems(ByVal intType As Integer, ByRef intMax As Integer) As String
    Dim arrCols
    Dim strItems As String
    Dim strCols As String
    Dim strNames As String
    Dim lngRecord As Long, lngStartRow As Long, lngEndRow As Long
    Dim intIn As Integer, intOut As Integer, intInMAX As Integer, intOutMax As Integer, intCount As Integer
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    'intType��ȷ����������ֵ��1)������Ŀ���;2)�����к�
    'intMAX������ͬ����������ռ�õ��и�
    '����ͬ��������(һ���ļ��в����ܳ����ظ�����Ŀ,����,�ж�ʱ���ؼ���к�)
    
    lngRecord = Val(VsfData.TextMatrix(VsfData.ROW, mlngRecord))
    If lngRecord = 0 Then Exit Function
    
    gstrSQL = "" & _
        " SELECT  B.��Ŀ���,B.��Ŀ����,A.������� AS �к�" & vbNewLine & _
        " FROM �����ļ��ṹ A,���˻�����ϸ B" & vbNewLine & _
        " WHERE A.Ҫ������=B.��Ŀ���� AND A.��ID=" & vbNewLine & _
        "      (SELECT A.ID FROM �����ļ��ṹ A,���˻����ļ� B " & vbNewLine & _
        "       WHERE B.ID=[2] And A.�ļ�ID=B.��ʽID AND A.�������=4 AND A.��ID IS NULL)" & vbNewLine & _
        " AND B.������Դ>0 AND B.��¼ID=[1]"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ͬ��������", lngRecord, mlng�ļ�ID)
    If rsTemp.RecordCount = 0 Then Exit Function
    
    '��ȡͬ�������Ϣ
    Do While Not rsTemp.EOF
        If InStr(1, "," & strCols & ",", "," & rsTemp!�к� & ",") = 0 Then strCols = strCols & "," & rsTemp!�к�
        strItems = strItems & "," & rsTemp!��Ŀ���
        strNames = strNames & "," & rsTemp!��Ŀ����
        rsTemp.MoveNext
    Loop
    strCols = Mid(strCols, 2)
    strItems = Mid(strItems, 2)
    strNames = Mid(strNames, 2)
    
    '������ѭ�����������ռ�и�
    If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
        lngStartRow = VsfData.ROW
        lngEndRow = VsfData.ROW
        intInMAX = 1
    Else
        lngStartRow = GetStartRow(VsfData.ROW)
        intInMAX = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
        lngEndRow = lngStartRow + intInMAX - 1
    End If
    
    intCount = 1    'ͬ����ֻ������������Ŀ������ռ����ֻ������1�У��������ݲ�����Ҫ���
'    '����ռ�ó���1�вż��
'    If intInMAX > 1 Then
'        arrCols = Split(strCols, ",")
'        intOutMax = UBound(arrCols)
'        For intOut = 0 To intOutMax
'            For intIn = 2 To intInMAX
'                If VsfData.TextMatrix(intIn + lngStartRow - 1, arrCols(intOut) + 1) <> "" Then
'                    If intIn > intCount Then intCount = intIn
'                End If
'            Next
'        Next
'    End If
    
    intMax = intCount
    GetSynItems = IIf(intType = 1, strItems, strCols)
    If strNames <> "" Then
        RaiseEvent AfterRowColChange("������,ʱ����,�Լ� " & strNames & " ��ͬ�����������ݣ��������޸Ļ�ɾ����", True, mblnSign, mblnArchive)
    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ISColHaveData() As Boolean
    Dim arrData
    Dim arrCol
    Dim intCOl As Integer
    Dim intDo As Integer, intCount As Integer
    Dim intIn As Integer, intMax As Integer
    Dim strCond As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '�����ݿ�����ȡ���ݣ������ǰ���Ŀ�д�������������������Ŀ����
    
    '�����Ŀ���뵽��ѯSQL�У���ʽ���к�;��ͷ����|��Ŀ���,��λ;��Ŀ���,��λ||�к�;��ͷ����...
    '�󶨶����Ŀ�����о��Զ�תΪ�Խ�����
    arrData = Split(mstrCOLActive, "||")
    intCount = UBound(arrData)
    For intDo = 0 To intCount
        intCOl = Split(Split(arrData(intDo), "|")(0), ";")(0)
        If intCOl = VsfData.COL - cHideCols - VsfData.FixedCols + 1 Then
            arrCol = Split(Split(arrData(intDo), "|")(1), ";")
            intMax = UBound(arrCol)
            For intIn = 0 To intMax
                strCond = strCond & " OR (��Ŀ���=" & Split(arrCol(intIn), ",")(0)
                If Split(arrCol(intIn), ",")(1) = "" Then
                    strCond = strCond & ")"
                Else
                    strCond = strCond & " AND NVL(���²�λ,'TWBW')='" & Split(arrCol(intIn), ",")(1) & "')"
                End If
            Next
            
            Exit For
        End If
    Next
    
    If strCond <> "" Then
        strCond = " AND (" & Mid(strCond, 4) & ")"
        '��ѯ���ݿ�
        gstrSQL = " SELECT  1 FROM ���˻�����ϸ A,���˻������� B,���˻����ӡ C" & vbNewLine & _
                  " Where A.��¼ID=B.ID And B.�������=0 And B.ID=C.��¼ID And C.�ļ�ID=B.�ļ�ID " & vbNewLine & _
                  " And C.�ļ�ID=[1] And (C.����ҳ��=[2] OR C.��ʼҳ��=[2])" & strCond & " AND ROWNUM<2"
        Call SQLDIY(gstrSQL)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ѯ���ݿ⵱ǰҳ��ָ������Ƿ���ڻ��Ŀ", mlng�ļ�ID, mintҳ��)
        ISColHaveData = rsTemp.RecordCount
    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


'######################################################################################################################
'**********************************************************************************************************************
'�����ǻ������������

Private Sub txt����ʱ��_GotFocus()
    Call zlControl.TxtSelAll(txt����ʱ��)
End Sub

Private Sub txt����ʱ��_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnCancel As Boolean
    If KeyCode = vbKeyReturn Then
        Call txt����ʱ��_Validate(blnCancel)
        txtС������.SetFocus
    End If
End Sub

Private Sub txt����ʱ��_Validate(Cancel As Boolean)
    Dim strFormat As String
    Dim intDef As Integer   'ʱ��+1
    '��鿪ʼʱ��,����ʱ��Ϸ���
    
    strFormat = CheckTxtTime(txt��ʼʱ��)
    If strFormat = "" Then Exit Sub
    txt��ʼʱ��.Text = strFormat
    strFormat = CheckTxtTime(txt����ʱ��)
    If strFormat = "" Then Exit Sub
    txt����ʱ��.Text = strFormat
    
    '����С������
    If txt����ʱ��.Text > txt��ʼʱ��.Text Then
        intDef = Val(txt����ʱ��.Text) - Val(txt��ʼʱ��.Text)
    Else
        intDef = Val(txt����ʱ��.Text) + 24 - Val(txt��ʼʱ��.Text)
    End If
    '�����������59�����1Сʱ
    If Split(txt����ʱ��.Text, ":")(1) = "59" Then intDef = intDef + 1
    txtС������.Tag = intDef
    txtС������.Text = intDef & "СʱС��"
End Sub

Private Sub txtС������_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmdOk.SetFocus
End Sub

Private Sub txt��ʼʱ��_GotFocus()
    Call zlControl.TxtSelAll(txt��ʼʱ��)
End Sub

Private Sub txt��ʼʱ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txt����ʱ��.SetFocus
End Sub

Private Sub lblDnInput_Click()
    txtDnInput.SetFocus
End Sub

Private Sub lblUpInput_Click()
    txtUpInput.SetFocus
End Sub

Private Sub lstColumnItems_DblClick()
    Call cmdColumn_Click(0)
End Sub

Private Sub lstColumnItems_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call lstColumnItems_DblClick
End Sub

Private Sub lstColumnUsed_DblClick()
    Call cmdColumn_Click(1)
End Sub

Private Sub lstColumnUsed_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call lstColumnUsed_DblClick
End Sub

Private Sub lstSelect_DblClick(Index As Integer)
    Call lstSelect_KeyDown(Index, vbKeyReturn, 0)
End Sub

Private Sub lstSelect_GotFocus(Index As Integer)
    mblnEditAssistant = False
End Sub

Private Sub txtColumnNo_GotFocus()
    txtColumnNo.SelStart = 0
    txtColumnNo.SelLength = 100
End Sub

Private Sub txtColumnNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" &[]{}+'""|", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtDnInput_GotFocus()
    txtDnInput.SelStart = 0
    txtDnInput.SelLength = 100
    Call ISAssistant(Val(txtDnInput.Tag), txtDnInput)
End Sub

Private Sub txtInput_GotFocus()
    txtInput.SelStart = 0
    txtInput.SelLength = 100
    mintSymbol = -1
    Call ISAssistant(Val(txtInput.Tag), txtInput)
End Sub

Private Sub txtUpInput_GotFocus()
    txtUpInput.SelStart = 0
    txtUpInput.SelLength = 100
    Call ISAssistant(Val(txtUpInput.Tag), txtUpInput)
End Sub

Private Sub txt_GotFocus(Index As Integer)
    txt(Index).SelStart = 0
    txt(Index).SelLength = 100
    mintSymbol = Index
    Call ISAssistant(Val(txt(Index).Tag), txt(Index))
End Sub

Private Sub lblUpInput_DblClick()
    lblUpInput.Caption = IIf(lblUpInput.Caption = "", "��", "")
    txtUpInput.SetFocus
End Sub

Private Sub lblDnInput_DblClick()
    lblDnInput.Caption = IIf(lblDnInput.Caption = "", "��", "")
    txtDnInput.SetFocus
End Sub

Private Sub lblInput_DblClick()
    lblInput.Caption = IIf(lblInput.Caption = "", "��", "")
End Sub

Private Sub txtUpInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtDnInput.SetFocus
    ElseIf KeyCode = vbKeyRight Then
        If txtUpInput.SelStart = Len(txtUpInput.Text) Then txtDnInput.SetFocus
    ElseIf KeyCode = vbKeyLeft And txtUpInput.SelStart = 1 Then
        Call MoveNextCell(False)
    ElseIf KeyCode = vbKeySpace And txtUpInput.Locked Then
        lblUpInput.Caption = IIf(lblUpInput.Caption = "", "��", "")
    End If
End Sub

Private Sub txtDnInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyRight And txtDnInput.SelStart = Len(txtDnInput.Text)) Then
        Call picDouble_KeyDown(KeyCode, Shift)
    ElseIf KeyCode = vbKeyLeft Then
        If txtDnInput.SelStart = 0 Then txtUpInput.SetFocus
    ElseIf KeyCode = vbKeySpace And txtDnInput.Locked Then
        lblDnInput.Caption = IIf(lblDnInput.Caption = "", "��", "")
    End If
End Sub

Private Sub picMutilInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call MoveNextCell
End Sub

Private Sub picDouble_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Then
        Call MoveNextCell(Not (KeyCode = vbKeyLeft))
    End If
End Sub

Private Sub picInput_GotFocus()
    If txtInput.Visible Then
        txtInput.SetFocus
    End If
End Sub

Private Sub picInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not txtInput.Visible Then
        If KeyCode = vbKeySpace Then
            Call lblInput_DblClick
        End If
    End If
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Then
        '�ƶ�����һ����Ԫ��
        Call MoveNextCell(Not (KeyCode = vbKeyLeft))
    End If
End Sub

Private Sub lstSelect_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call MoveNextCell
End Sub

Private Sub picMutilInput_GotFocus()
    On Error Resume Next
    txt(0).SetFocus
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Index < txt.Count - 1 Then
            txt(Index + 1).SetFocus
        Else
            Call picMutilInput_KeyDown(KeyCode, Shift)
        End If
    End If
End Sub

Private Sub picDouble_GotFocus()
    If txtUpInput.Visible Then
        txtUpInput.SetFocus
    End If
End Sub

Private Sub picMain_Resize()
    picMain.Left = 0
    VsfData.Width = picMain.Width
    VsfData.Height = IIf(picMain.Height - VsfData.Top < 0, 0, picMain.Height - VsfData.Top)
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Shift = vbCtrlMask Then Exit Sub
    
    If KeyCode = vbKeyReturn Or _
        (KeyCode = vbKeyRight And txtInput.SelStart = Len(txtInput.Text)) Or _
        (KeyCode = vbKeyLeft And txtInput.SelStart = 1) Then
        Call picInput_KeyDown(KeyCode, Shift)
    End If
End Sub

Private Sub txtUpInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("/") Then
        KeyAscii = 0
        txtDnInput.SetFocus
    End If
End Sub
 

Private Sub txtС������_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = Asc(";") Then KeyAscii = 0
End Sub

Private Sub cboС��_Click()
    If cboС��.Tag = "" Then Exit Sub
    
    txt��ʼʱ��.Enabled = (cboС��.Text = "��ʱ")
    txt����ʱ��.Enabled = txt��ʼʱ��.Enabled
    If cboС��.Text <> "��ʱ" Then
        txt��ʼʱ��.Text = Split(Split(cboС��.Tag, ";")(cboС��.ListIndex), ",")(0)
        txt����ʱ��.Text = Split(Split(cboС��.Tag, ";")(cboС��.ListIndex), ",")(1)
        'txtС������.Text = Format(DateAdd("d", -1 * cboС�᷶Χ.ListIndex, zldatabase.Currentdate), "MM-DD") & " " & cboС��.Text
        txtС������.Text = Format(DTPDate.Value, "MM-DD") & " " & cboС��.Text
        txtС������.Tag = 0
    Else
        txtС������.Text = ""
        txt��ʼʱ��.Text = ""
        txt����ʱ��.Text = ""
        txtС������.Tag = 0
    End If
End Sub


Private Sub UserControl_GotFocus()
    On Error Resume Next
    VsfData.SetFocus
End Sub

Private Sub UserControl_Initialize()
    mblnShow = False
    mblnChange = False
    mblnInit = False
    
'    Set objStream = objFileSys.OpenTextFile("C:\WORKLOG.txt", ForAppending, True)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    '�����ַ���Ϊ���ݷָ�������¼�¼���ķָ�������˲�����¼��
    If KeyAscii = 39 Or KeyAscii = 13 Or KeyAscii = Asc("|") Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyEscape And mblnShow Then
        mblnShow = False
        Call InitCons
    End If
End Sub

Private Sub cbsThis_Resize()
    Dim lngScaleLeft As Long, lngScaleTop As Long, lngScaleRight As Long, lngScaleBottom As Long
    Call cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    
    Err = 0: On Error Resume Next
    lblTitle.Move lngScaleLeft, lngScaleTop + 120, lngScaleRight - lngScaleLeft
    With lblSubhead
        .Left = lngScaleLeft + 210: .Width = lngScaleRight - lngScaleLeft - 210 * 2
        .Top = lblTitle.Top + lblTitle.Height + 120
    End With
    picMain.Move lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom - lngScaleTop
    VsfData.Move lngScaleLeft + 210, lblSubhead.Top + lblSubhead.Height + 45, lngScaleRight - lngScaleLeft - 210 * 2
    VsfData.Height = picMain.Height - VsfData.Top - lblSubEnd.Height - 45
    lblSubEnd.Move lblSubhead.Left, VsfData.Top + VsfData.Height + 45
    
    lblCurPage.Top = picMain.Top
    lblCurPage.Left = picMain.Width - lblCurPage.Width
    
    '���ϱ�ǩ��ɢ����
    Call zlLableBruit
End Sub

Private Sub UserControl_Terminate()
'    objStream.Close
End Sub

Private Sub SetDockRight(BarToDock As CommandBar, BarOnLeft As CommandBar)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    cbsThis.RecalcLayout
    BarOnLeft.GetWindowRect Left, Top, Right, Bottom
    cbsThis.DockToolBar BarToDock, Right, (Bottom + Top) / 2, BarOnLeft.Position
End Sub

Public Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '��Ӽ�¼
    'strFields:�ֶ���|�ֶ���
    'strValues:ֵ|ֵ
    
    '���ӣ�
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|��ĿID|ժҪ"
    'strValues = "5188|6666|��Ŀ����"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Public Sub Record_Update(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False)
    Dim arrFields, arrValues, intField As Integer
    '���¼�¼,���������,������
    'strPrimary:�ֶ���|ֵ
    'strFields:�ֶ���|�ֶ���
    'strValues:ֵ|ֵ
    
    '���ӣ�
    'Dim strFields As String, strValues As String, strPrimary As String
    'strFields = "RecordID|��ĿID|ժҪ"
    'strValues = "5188|6666|��Ŀ����"
    'strPrimary = "RecordID|5188"
    'Call Record_Update(rsVoucher, strFields, strValues, strPrimary, True)

    If strValues = "" Then strValues = " "
    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField < 0 Then Exit Sub

    With rsObj
        If Record_Locate(rsObj, strPrimary, blnDelete) = False Then .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Public Function Record_Locate(ByRef rsObj As ADODB.Recordset, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False) As Boolean
    Dim arrTmp
    '��λ��ָ����¼
    'strPrimary:����,ֵ
    'blnDelete=True,��ü�¼������"ɾ��"�ֶ�
    Record_Locate = False
    
    arrTmp = Split(strPrimary, "|")
    With rsObj
        If .RecordCount = 0 Then Exit Function
        .MoveFirst
        .Find arrTmp(0) & "='" & arrTmp(1) & "'"
        If .EOF Then Exit Function
        If blnDelete Then
            Do While Not .EOF
                If !ɾ�� = 0 Then Record_Locate = True: Exit Do
                .MoveNext
            Loop
        Else
            Record_Locate = True
        End If
    End With
End Function

Public Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intType As Integer, lngLength As Long
    '��ʼ��ӳ���¼��
    'strFields:�ֶ���,����,����|�ֶ���,����,����    �������Ϊ��,��ȡĬ�ϳ���
    '�ַ���:adLongVarChar;������:adDouble;������:adDBDate
    
    '���ӣ�
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|��ĿID," & adDouble & ",18|ժҪ, " & adLongVarChar & ",50|" & _
    '"ɾ��," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '��ȡ�ֶ�ȱʡ����
            If lngLength = 0 Then
                Select Case intType
                Case adDouble
                    lngLength = madDoubleDefault
                Case adVarChar
                    lngLength = madLongVarCharDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub OutputRsData(ByVal rsObj As ADODB.Recordset)
    Dim intCOl As Integer, intCols As Integer
    Dim strValues As String
    With rsObj
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strValues = ""
            intCols = .Fields.Count - 1
            For intCOl = 0 To intCols
                strValues = strValues & "," & .Fields(intCOl).Name & ":" & .Fields(intCOl).Value
            Next
            Debug.Print Mid(strValues, 2)
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
End Sub

Private Function BlowUp(ByRef dblChange As Double) As Double
    '�Ŵ����壬��Ԫ����
    BlowUp = dblChange
    If Not mblnBlowup Then Exit Function
    BlowUp = dblChange + (dblChange * 1 / 3)
End Function

