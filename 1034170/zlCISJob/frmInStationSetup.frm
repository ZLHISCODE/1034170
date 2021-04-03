VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmInStationSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   Icon            =   "frmInStationSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   7740
      TabIndex        =   47
      Top             =   8445
      Width           =   7740
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   6480
         TabIndex        =   49
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   5280
         TabIndex        =   48
         Top             =   120
         Width           =   1100
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         X1              =   0
         X2              =   8040
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   8040
         Y1              =   15
         Y2              =   0
      End
   End
   Begin TabDlg.SSTab sstInfo 
      Height          =   8475
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7770
      _ExtentX        =   13705
      _ExtentY        =   14949
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "����"
      TabPicture(0)   =   "frmInStationSetup.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraAdvice"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraEPR"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra��ҳ����"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraMedRec"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkDeptView"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraҩ��"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "��ҳ������Ŀ"
      TabPicture(1)   =   "frmInStationSetup.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraGroup"
      Tab(1).ControlCount=   1
      Begin VB.Frame fraҩ�� 
         Caption         =   "סԺҽ���༭"
         Height          =   615
         Left            =   120
         TabIndex        =   90
         Top             =   7320
         Width           =   7575
         Begin VB.CheckBox chkȱʡҩ�� 
            Caption         =   "סԺҽ���´�ǿ��ȱʡҩ��"
            Height          =   240
            Left            =   120
            TabIndex        =   91
            Top             =   240
            Width           =   2580
         End
      End
      Begin VB.CheckBox chkDeptView 
         Caption         =   "ӵ��ȫԺ����Ȩ�޲����߲���ʾ�޴�λ�Ĳ�������"
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   8040
         Width           =   4335
      End
      Begin VB.Frame fraMedRec 
         Caption         =   "������鷴������"
         Height          =   615
         Left            =   120
         TabIndex        =   32
         Top             =   6690
         Width           =   7560
         Begin VB.TextBox txtMedRec 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            IMEMode         =   3  'DISABLE
            Left            =   1040
            MaxLength       =   3
            TabIndex        =   34
            Text            =   "1"
            Top             =   240
            Width           =   300
         End
         Begin VB.Frame Frame4 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   15
            Left            =   1025
            TabIndex        =   33
            Top             =   420
            Width           =   300
         End
         Begin VB.Label lblMedRec 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ʾ    ���ڵĲ�����鷴����"
            Height          =   180
            Left            =   645
            TabIndex        =   35
            Top             =   255
            Width           =   2520
         End
      End
      Begin VB.Frame fra��ҳ���� 
         Caption         =   " ��ҳ�������� "
         Height          =   4575
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   7560
         Begin VB.CheckBox chkZLZD 
            Caption         =   "��Ժ���������д�ֻ��̶Ⱥ�����������"
            Height          =   255
            Left            =   2280
            TabIndex        =   89
            Top             =   4200
            Width           =   4095
         End
         Begin VB.CheckBox Chk���� 
            Caption         =   "��ICD-10¼��ʱ���������ֻ����¼��M��ͷ��������̬ѧ����"
            Height          =   255
            Left            =   240
            TabIndex        =   77
            Top             =   1560
            Width           =   5295
         End
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   255
            Index           =   3
            Left            =   3480
            TabIndex        =   72
            Top             =   1875
            Width           =   3615
            Begin VB.OptionButton optICD���� 
               Caption         =   "������д"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   75
               Top             =   0
               Width           =   1095
            End
            Begin VB.OptionButton optICD���� 
               Caption         =   "��ʾ�Ƿ���д"
               Height          =   255
               Index           =   2
               Left            =   1200
               TabIndex        =   74
               Top             =   0
               Width           =   1455
            End
            Begin VB.OptionButton optICD���� 
               Caption         =   "�����"
               Height          =   255
               Index           =   0
               Left            =   2640
               TabIndex        =   73
               Top             =   0
               Value           =   -1  'True
               Width           =   855
            End
         End
         Begin VB.CheckBox chkSeparEdit 
            Caption         =   "ҽ���ͻ�ʿ�ֱ���д������ҳ"
            Height          =   255
            Left            =   240
            TabIndex        =   50
            Top             =   3420
            Width           =   4095
         End
         Begin VB.CheckBox chk��ҽ 
            Caption         =   $"frmInStationSetup.frx":0044
            Height          =   255
            Left            =   240
            TabIndex        =   37
            Top             =   2475
            Width           =   4095
         End
         Begin VB.CheckBox chkʹ����������ʱ�� 
            Caption         =   "ʹ����������ʱ��"
            Height          =   255
            Left            =   240
            TabIndex        =   36
            Top             =   600
            Width           =   2055
         End
         Begin VB.ComboBox cboType 
            Height          =   300
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   240
            Width           =   2535
         End
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   255
            Index           =   2
            Left            =   1440
            TabIndex        =   25
            Top             =   2200
            Width           =   3735
            Begin VB.OptionButton opt���� 
               Caption         =   "������д"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   28
               Top             =   0
               Width           =   1095
            End
            Begin VB.OptionButton opt���� 
               Caption         =   "��ʾ�Ƿ���д"
               Height          =   255
               Index           =   2
               Left            =   1200
               TabIndex        =   27
               Top             =   0
               Width           =   1455
            End
            Begin VB.OptionButton opt���� 
               Caption         =   "�����"
               Height          =   255
               Index           =   0
               Left            =   2640
               TabIndex        =   26
               Top             =   0
               Value           =   -1  'True
               Width           =   855
            End
         End
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   255
            Index           =   0
            Left            =   3840
            TabIndex        =   11
            Top             =   970
            Width           =   3495
            Begin VB.OptionButton opt�����ж� 
               Caption         =   "������д"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   14
               ToolTipText     =   "�����Ժ��ϲ���S��T�࣬���ֹ��д�����ж���"
               Top             =   0
               Width           =   1095
            End
            Begin VB.OptionButton opt�����ж� 
               Caption         =   "��ʾ�Ƿ���д"
               Height          =   255
               Index           =   2
               Left            =   1200
               TabIndex        =   13
               Top             =   0
               Width           =   1455
            End
            Begin VB.OptionButton opt�����ж� 
               Caption         =   "�����"
               Height          =   255
               Index           =   0
               Left            =   2640
               TabIndex        =   12
               Top             =   0
               Value           =   -1  'True
               Width           =   855
            End
         End
         Begin VB.Frame Frame1 
            BorderStyle     =   0  'None
            Caption         =   "Frame1"
            Height          =   255
            Index           =   1
            Left            =   3600
            TabIndex        =   7
            Top             =   1300
            Width           =   3495
            Begin VB.OptionButton opt������� 
               Caption         =   "�����"
               Height          =   255
               Index           =   0
               Left            =   2640
               TabIndex        =   10
               Top             =   0
               Value           =   -1  'True
               Width           =   855
            End
            Begin VB.OptionButton opt������� 
               Caption         =   "��ʾ�Ƿ���д"
               Height          =   255
               Index           =   2
               Left            =   1200
               TabIndex        =   9
               Top             =   0
               Width           =   1455
            End
            Begin VB.OptionButton opt������� 
               Caption         =   "������д"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   8
               ToolTipText     =   "�����Ժ��ϲ���C00��D48ʱ�����ֹ��д������ϡ�"
               Top             =   0
               Width           =   1095
            End
         End
         Begin VB.CheckBox chkGet���� 
            Caption         =   "����Զ���ȡ����"
            Height          =   255
            Left            =   240
            TabIndex        =   84
            Top             =   4200
            Width           =   1815
         End
         Begin VB.Label lblICD���� 
            Caption         =   "��Ҫ��Ժ��ϱ���ΪC00��D48ʱ,ICD���룺"
            Height          =   210
            Left            =   240
            TabIndex        =   76
            Top             =   1920
            Width           =   3495
         End
         Begin VB.Label lblSeparEdit 
            Caption         =   "�����¼�����Ŀ����Һ��Ӧ������ҩ��ٴ����֡�סԺ�ڼ�����Լ������Ժʱ͸��(Ѫ͸����͸)���ص�ֵ����Ϣ�����øò���ʱֻ���ɻ�ʿ��д"
            Height          =   360
            Left            =   480
            TabIndex        =   51
            Top             =   3720
            Width           =   6615
         End
         Begin VB.Label lbl��ҽ 
            Caption         =   $"frmInStationSetup.frx":0070
            Height          =   585
            Left            =   480
            TabIndex        =   38
            Top             =   2790
            Width           =   6735
         End
         Begin VB.Label lbl��ҳ��׼ 
            Caption         =   "������ҳ��׼"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   285
            Width           =   1215
         End
         Begin VB.Label lbl���� 
            Caption         =   "����ʱ�����"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   2200
            Width           =   1335
         End
         Begin VB.Label lbl�����ж� 
            Caption         =   "��Ҫ��Ժ��ϱ���ΪS��T��ʱ,�����ж���ϣ�"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   970
            Width           =   3735
         End
         Begin VB.Label lbl������� 
            Caption         =   "��Ҫ��Ժ��ϱ���ΪC00��D48ʱ,������ϣ�"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   1320
            Width           =   3375
         End
      End
      Begin VB.Frame fraGroup 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6800
         Left            =   -74880
         TabIndex        =   1
         Top             =   360
         Width           =   7395
         Begin VB.CommandButton cmdAdd 
            Caption         =   "����(&A)"
            Height          =   350
            Left            =   3120
            TabIndex        =   4
            Top             =   30
            Width           =   1100
         End
         Begin VB.CommandButton cmdModify 
            Caption         =   "�޸�(&M)"
            Enabled         =   0   'False
            Height          =   350
            Left            =   4320
            TabIndex        =   3
            Top             =   30
            Width           =   1100
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "ɾ��(&D)"
            Enabled         =   0   'False
            Height          =   350
            Left            =   5520
            TabIndex        =   2
            Top             =   30
            Width           =   1100
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfMain 
            Height          =   6300
            Left            =   315
            TabIndex        =   5
            Top             =   480
            Width           =   6840
            _cx             =   12065
            _cy             =   11112
            Appearance      =   3
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
            BackColorSel    =   16574424
            ForeColorSel    =   -2147483642
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483632
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   2
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   250
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
      End
      Begin VB.Frame fraEPR 
         Caption         =   "��������"
         Height          =   1545
         Left            =   120
         TabIndex        =   39
         Top             =   5040
         Width           =   7560
         Begin VB.CheckBox chkWarn 
            Caption         =   "��Ѫ��Ӧ"
            Height          =   195
            Index           =   26
            Left            =   6480
            TabIndex        =   88
            Top             =   1200
            Width           =   1035
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "��Ѫ���"
            Height          =   195
            Index           =   25
            Left            =   5400
            TabIndex        =   83
            Top             =   1185
            Width           =   1035
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "У������"
            Height          =   195
            Index           =   24
            Left            =   4320
            TabIndex        =   78
            Top             =   1200
            Width           =   1035
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "��Ѫ���"
            Height          =   195
            Index           =   23
            Left            =   3255
            TabIndex        =   55
            Top             =   1200
            Width           =   1035
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "�����ʿ�"
            Height          =   195
            Index           =   22
            Left            =   2235
            TabIndex        =   71
            Top             =   1185
            Width           =   1035
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "�������"
            Height          =   195
            Index           =   20
            Left            =   6480
            TabIndex        =   70
            Top             =   885
            Width           =   1035
         End
         Begin VB.CheckBox chkSoundYS 
            Caption         =   "����������ʾ"
            Height          =   195
            Left            =   4335
            TabIndex        =   69
            Top             =   345
            Width           =   1470
         End
         Begin VB.CommandButton cmdSoundYSSet 
            Caption         =   "��������(&S)"
            Height          =   350
            Left            =   5775
            TabIndex        =   68
            Top             =   270
            Width           =   1410
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "��Ⱦ��"
            Height          =   195
            Index           =   21
            Left            =   1200
            TabIndex        =   65
            Top             =   1185
            Width           =   885
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "ҽ�����"
            Height          =   195
            Index           =   19
            Left            =   5400
            TabIndex        =   63
            Top             =   885
            Width           =   1035
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "���泷��"
            Height          =   195
            Index           =   18
            Left            =   4320
            TabIndex        =   56
            Top             =   885
            Width           =   1035
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "Σ��ֵ"
            Height          =   195
            Index           =   17
            Left            =   3255
            TabIndex        =   54
            Top             =   885
            Width           =   885
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "ҽ������"
            Height          =   195
            Index           =   16
            Left            =   2235
            TabIndex        =   53
            Top             =   885
            Width           =   1020
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "��������"
            Height          =   195
            Index           =   15
            Left            =   1200
            TabIndex        =   52
            Top             =   885
            Width           =   1065
         End
         Begin VB.TextBox txtNotifyEPRDay 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            IMEMode         =   3  'DISABLE
            Left            =   600
            MaxLength       =   2
            TabIndex        =   43
            Text            =   "1"
            Top             =   600
            Width           =   300
         End
         Begin VB.Frame fraNotifyEPRDay 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   15
            Left            =   585
            TabIndex        =   42
            Top             =   780
            Width           =   300
         End
         Begin VB.Frame fraNotifyEPR 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   15
            Left            =   585
            TabIndex        =   41
            Top             =   510
            Width           =   300
         End
         Begin VB.TextBox txtNotifyEPR 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   180
            IMEMode         =   3  'DISABLE
            Left            =   600
            MaxLength       =   3
            TabIndex        =   40
            Text            =   "10"
            Top             =   330
            Width           =   300
         End
         Begin VB.CheckBox chkNotifyEPR 
            Caption         =   "ÿ    �����Զ�ˢ�����������е�����"
            Height          =   195
            Left            =   105
            TabIndex        =   44
            Top             =   345
            Width           =   3900
         End
         Begin VB.Label lblNotifyEPRDay 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��    ������ɵ�������ʾ����������"
            Height          =   180
            Left            =   375
            TabIndex        =   46
            Top             =   615
            Width           =   3060
         End
         Begin VB.Label lblArea 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������:"
            Height          =   180
            Left            =   360
            TabIndex        =   45
            Top             =   885
            Width           =   810
         End
      End
      Begin VB.Frame fraAdvice 
         Caption         =   "�������� "
         Height          =   1590
         Left            =   120
         TabIndex        =   17
         Top             =   5040
         Width           =   7560
         Begin VB.CheckBox chkWarn 
            Caption         =   "Ѫ������"
            Height          =   195
            Index           =   12
            Left            =   6240
            TabIndex        =   87
            Top             =   1245
            Width           =   1020
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "��Ѫ���"
            Height          =   195
            Index           =   11
            Left            =   5160
            TabIndex        =   86
            Top             =   1245
            Width           =   1020
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "RISԤԼ׼��"
            Height          =   195
            Index           =   8
            Left            =   2640
            TabIndex        =   79
            Top             =   1245
            Width           =   1335
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "ȡѪ֪ͨ"
            Height          =   195
            Index           =   9
            Left            =   3960
            TabIndex        =   82
            Top             =   1245
            Width           =   1095
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "��������"
            Height          =   195
            Index           =   6
            Left            =   600
            TabIndex        =   81
            Top             =   1230
            Width           =   1035
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "RISԤԼ"
            Height          =   195
            Index           =   7
            Left            =   1680
            TabIndex        =   80
            Top             =   1245
            Width           =   1005
         End
         Begin VB.CommandButton cmdSoundHSSet 
            Caption         =   "��������(&S)"
            Height          =   350
            Left            =   6000
            TabIndex        =   67
            Top             =   240
            Width           =   1410
         End
         Begin VB.CheckBox chkSoundHS 
            Caption         =   "����������ʾ"
            Height          =   195
            Left            =   4440
            TabIndex        =   66
            Top             =   360
            Width           =   1470
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "��Һ�ܾ�"
            Height          =   195
            Index           =   5
            Left            =   5520
            TabIndex        =   62
            Top             =   915
            Width           =   1035
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "Σ��ֵ"
            Height          =   195
            Index           =   4
            Left            =   4560
            TabIndex        =   61
            Top             =   915
            Width           =   870
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "����"
            Height          =   195
            Index           =   3
            Left            =   3840
            TabIndex        =   60
            Top             =   915
            Width           =   675
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "�·�"
            Height          =   195
            Index           =   2
            Left            =   3060
            TabIndex        =   59
            Top             =   915
            Width           =   660
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "��ͣ"
            Height          =   195
            Index           =   1
            Left            =   2280
            TabIndex        =   58
            Top             =   915
            Width           =   675
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "�¿�"
            Height          =   195
            Index           =   0
            Left            =   1500
            TabIndex        =   57
            Top             =   915
            Width           =   675
         End
         Begin VB.TextBox txtNotifyAdviceDay 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Height          =   180
            IMEMode         =   3  'DISABLE
            Left            =   795
            MaxLength       =   2
            TabIndex        =   21
            Text            =   "1"
            Top             =   585
            Width           =   300
         End
         Begin VB.Frame fraNotifyAdviceDay 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   15
            Left            =   780
            TabIndex        =   20
            Top             =   765
            Width           =   300
         End
         Begin VB.Frame fraNotifyAdvice 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   15
            Left            =   780
            TabIndex        =   19
            Top             =   495
            Width           =   300
         End
         Begin VB.TextBox txtNotifyAdvice 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   180
            IMEMode         =   3  'DISABLE
            Left            =   795
            MaxLength       =   3
            TabIndex        =   18
            Text            =   "10"
            Top             =   315
            Width           =   300
         End
         Begin VB.CheckBox chkNotifyAdvice 
            Caption         =   "ÿ    �����Զ�ˢ��ҽ�����������е�����"
            Height          =   195
            Left            =   300
            TabIndex        =   22
            Top             =   330
            Width           =   3900
         End
         Begin VB.CheckBox chkWarn 
            Caption         =   "�걾���գ���δ���ã�"
            Enabled         =   0   'False
            Height          =   195
            Index           =   10
            Left            =   4560
            TabIndex        =   85
            Top             =   600
            Visible         =   0   'False
            Width           =   2115
         End
         Begin VB.Label lblNotifyAdviceDay 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��    ���ڴ����ҽ��������ʾ����������"
            Height          =   180
            Left            =   570
            TabIndex        =   24
            Top             =   600
            Width           =   3420
         End
         Begin VB.Label lbl�������� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������:"
            Height          =   180
            Left            =   555
            TabIndex        =   23
            Top             =   915
            Width           =   810
         End
      End
   End
End
Attribute VB_Name = "frmInStationSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mbln��ʿվ As Boolean
Public mstrPrivs As String
Private mlngModual As Long
Private lng�����ж� As Long
Private lng������� As Long
Private lngICD���� As Long
Private lng���� As Long

Private Enum Enum_chkWarn
    '��ʿվ���Ѳ���
    chkN�¿� = 0
    chkN��ͣ = 1
    chkN�·� = 2
    chkN���� = 3
    chkNΣ��ֵ = 4
    chkN��Һ�ܾ� = 5
    chkN�������� = 6
    chkNRISԤԼ = 7
    chkNRISԤԼ׼�� = 8
    chkȡѪ֪ͨ = 9
    chk�걾���� = 10        '��δ���ã������˴�һ��Ϊ�˰汾֮��ļ����ԣ�һ��Ϊ�˺�������˹���ʱ����ɻ��ҡ�
    chk��Ѫ��� = 11
    chkѪ������ = 12
    'ҽ��վ���Ѳ���
    chkD�������� = 15
    chkDҽ������ = 16
    chkDΣ��ֵ = 17
    chkD���泷�� = 18
    chkDҽ����� = 19
    chkD������� = 20
    chkD��Ⱦ�� = 21
    chkD�����ʿ� = 22
    chkD��Ѫ��� = 23
    chkDУ������ = 24
    chkD��Ѫ��� = 25
    chkD��Ѫ��Ӧ = 26
End Enum

Public Sub ShowMe()
    '���°�סԺ��ʿ����վ���ã���ʾ��ע��ť
    Me.Show vbModal
End Sub

Private Sub cboType_Click()
    If cboType.ListIndex = 0 Or cboType.ListIndex = 3 Then
        chkʹ����������ʱ��.Visible = True
    Else
        chkʹ����������ʱ��.Visible = False
    End If
End Sub

Private Sub chkNotifyAdvice_Click()
    txtNotifyAdvice.Enabled = chkNotifyAdvice.Value = 1
    If Visible And txtNotifyAdvice.Enabled Then txtNotifyAdvice.SetFocus
End Sub

Private Sub chkNotifyEPR_Click()
    txtNotifyEPR.Enabled = chkNotifyEPR.Value = 1
    If Visible And txtNotifyEPR.Enabled Then txtNotifyEPR.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdModify_Click()
    If vsfMain.Row > 0 Then
        If CheckVsf = False Then vsfMain.SetFocus: Exit Sub
        frmInMedSetup.ShowMe vsfMain.TextMatrix(vsfMain.Row, 0), vsfMain.TextMatrix(vsfMain.Row, 1), vsfMain.TextMatrix(vsfMain.Row, 2), "�޸�", Me
        vsfMain.SetFocus
    End If
End Sub

Private Sub cmdDelete_Click()
    Dim strSQL As String
    
    If vsfMain.Row > 0 Then
        If CheckVsf = False Then vsfMain.SetFocus: Exit Sub
        If MsgBox("ȷ��Ҫɾ��[" & vsfMain.TextMatrix(vsfMain.Row, 1) & "]��?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        strSQL = "zl_������Ŀ_edit(null,null,null,'" & vsfMain.TextMatrix(vsfMain.Row, 0) & "',2)"
        On Error GoTo errHandle
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        vsfMain.RemoveItem vsfMain.Row
        If vsfMain.Rows = 1 Then
            cmdModify.Enabled = False
            cmdDelete.Enabled = False
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdAdd_Click()
    frmInMedSetup.ShowMe "", "", "", "����", Me
End Sub

Private Sub cmdOK_Click()
    Dim curDate As Date
    Dim strTmp As String
    Dim i As Integer
    Dim blnSetup As Boolean
    
    If mbln��ʿվ Then
        If chkNotifyAdvice.Value = 1 And Val(txtNotifyAdvice.Text) = 0 Then
            If txtNotifyAdvice.Text = "" Then
                MsgBox "������ҽ�����ѵ��Զ�ˢ�¼����", vbInformation, gstrSysName
            Else
                MsgBox "ҽ�����ѵ��Զ�ˢ�¼������ӦΪ1���ӡ�", vbInformation, gstrSysName
            End If
            txtNotifyAdvice.SetFocus: Exit Sub
        End If
        If Val(txtNotifyAdviceDay.Text) = 0 Then
            If txtNotifyAdviceDay.Text = "" Then
                MsgBox "������Ҫ���ѵ�ҽ��������", vbInformation, gstrSysName
            Else
                MsgBox "Ҫ���ѵ�ҽ����������ӦΪ1�졣", vbInformation, gstrSysName
            End If
            txtNotifyAdviceDay.SetFocus: Exit Sub
        End If
    Else
        If chkNotifyEPR.Value = 1 And Val(txtNotifyEPR.Text) = 0 Then
            If txtNotifyEPR.Text = "" Then
                MsgBox "�����ò����������ѵ��Զ�ˢ�¼����", vbInformation, gstrSysName
            Else
                MsgBox "�����������ѵ��Զ�ˢ�¼������ӦΪ1���ӡ�", vbInformation, gstrSysName
            End If
            txtNotifyEPR.SetFocus: Exit Sub
        End If
        
        If Val(txtNotifyEPRDay.Text) = 0 Then
            If txtNotifyEPRDay.Text = "" Then
                MsgBox "������Ҫ�������ĵĲ������������", vbInformation, gstrSysName
            Else
                MsgBox "Ҫ�������ĵĲ��������������ӦΪ1�졣", vbInformation, gstrSysName
            End If
            txtNotifyEPRDay.SetFocus: Exit Sub
        End If
    End If
    
    If txtMedRec.Text = "" Then
        MsgBox "�����ò�����鷴�����ѵ�������", vbInformation, gstrSysName
        txtMedRec.SetFocus: Exit Sub
    End If
    blnSetup = InStr(";" & mstrPrivs & ";", ";��������;") > 0
    Call zlDatabase.SetPara("������鷴������", txtMedRec.Text, glngSys, mlngModual, blnSetup)
        
    '�Զ�ˢ��ҽ������
    If mbln��ʿվ Then
        Call zlDatabase.SetPara("�Զ�ˢ��ҽ�����", IIf(chkNotifyAdvice.Value = 1, Val(txtNotifyAdvice.Text), ""), glngSys, pסԺ��ʿվ, blnSetup)
        Call zlDatabase.SetPara("�Զ�ˢ��ҽ������", Val(txtNotifyAdviceDay.Text), glngSys, pסԺ��ʿվ, blnSetup)
        strTmp = ""
        For i = chkN�¿� To chkѪ������
            strTmp = strTmp & chkWarn(i).Value
        Next
        Call zlDatabase.SetPara("�Զ�ˢ��ҽ������", strTmp, glngSys, pסԺ��ʿվ, blnSetup)
        Call zlDatabase.SetPara("����������ʾ", chkSoundHS.Value, glngSys, pסԺ��ʿվ, blnSetup)
    Else
        Call zlDatabase.SetPara("�Զ�ˢ�²������ļ��", IIf(chkNotifyEPR.Value = 1, Val(txtNotifyEPR.Text), ""), glngSys, pסԺҽ��վ, blnSetup)
        Call zlDatabase.SetPara("�Զ�ˢ�²�����������", Val(txtNotifyEPRDay.Text), glngSys, pסԺҽ��վ, blnSetup)
        Call zlDatabase.SetPara("�����ж����", lng�����ж�, glngSys, pסԺҽ��վ, blnSetup)
        Call zlDatabase.SetPara("������ϼ��", lng�������, glngSys, pסԺҽ��վ, blnSetup)
        Call zlDatabase.SetPara("ICD������", lngICD����, glngSys, pסԺҽ��վ, blnSetup)
        Call zlDatabase.SetPara("������", lng����, glngSys, pסԺҽ��վ, blnSetup)
        Call zlDatabase.SetPara("�������ֻ����¼��������̬ѧ����", Chk����.Value, glngSys, pסԺҽ��վ, blnSetup)
        
        strTmp = ""
        For i = chkD�������� To chkD��Ѫ��Ӧ
            strTmp = strTmp & chkWarn(i).Value
        Next
        Call zlDatabase.SetPara("�Զ�ˢ������", strTmp, glngSys, pסԺҽ��վ, blnSetup)
        
        Call zlDatabase.SetPara("������ҳ��׼", cboType.ListIndex, glngSys, pסԺҽ��վ, blnSetup)
        Call zlDatabase.SetPara("ʹ����������ʱ��", chkʹ����������ʱ��, glngSys, pסԺҽ��վ, blnSetup)
        Call zlDatabase.SetPara("��ҽ���Ҳ�ʹ����ҽ������ҳ��Ŀ", chk��ҽ.Value, glngSys, pסԺҽ��վ, blnSetup)
        Call zlDatabase.SetPara("ҽ���ͻ�ʿ�ֱ���д������ҳ", chkSeparEdit.Value, glngSys, pסԺҽ��վ, blnSetup)
        Call zlDatabase.SetPara("����ʾ�޴�λ�Ĳ�������", chkDeptView.Value, glngSys, pסԺҽ��վ, blnSetup)
        Call zlDatabase.SetPara("����������ʾ", chkSoundYS.Value, glngSys, pסԺҽ��վ, blnSetup)
        Call zlDatabase.SetPara("����Զ���ȡ����", chkGet����.Value, glngSys, 0, blnSetup)
        Call zlDatabase.SetPara("��Ժ���������д�ֻ��̶Ⱥ�����������", chkZLZD.Value, glngSys, 0, blnSetup)
    End If
    
    Call zlDatabase.SetPara("סԺҽ���´�ǿ��ȱʡҩ��", chkȱʡҩ��.Value, glngSys, pסԺҽ���´�, blnSetup)
    
    gblnOK = True
    Unload Me
End Sub

Private Sub cmdSoundHSSet_Click()
    Call frmMsgCallSetup.ShowMe(Me, 2)
End Sub

Private Sub cmdSoundYSSet_Click()
    Call frmMsgCallSetup.ShowMe(Me, 1)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strPar As String, i As Long
    Dim curDate As Date, intDay As Integer
    Dim intType As Integer
    Dim strNotify As String
    Dim varTmp As Variant
    
    gblnOK = False
    lng������� = 0
    lng�����ж� = 0
    lngICD���� = 0
    mlngModual = IIf(mbln��ʿվ, pסԺ��ʿվ, pסԺҽ��վ)
    If mbln��ʿվ Then
        fraAdvice.Visible = True
        fraEPR.Visible = False
        sstInfo.TabVisible(1) = False
        fra��ҳ����.Visible = False
        chkDeptView.Visible = False
        fraAdvice.Top = fra��ҳ����.Top
        fraEPR.Top = fra��ҳ����.Top
        fraMedRec.Top = fraAdvice.Top + fraAdvice.Height + 50
        fraҩ��.Top = fraMedRec.Top + fraMedRec.Height + 50
        i = fra��ҳ����.Height + chkDeptView.Height + 100
    Else
        fraAdvice.Visible = False
        fraEPR.Visible = True
        chkDeptView.Visible = True
        fraMedRec.Top = fraEPR.Top + fraEPR.Height + 50
        chkDeptView.Top = fraҩ��.Top + fraҩ��.Height + 50
        i = fraAdvice.Height - fraEPR.Height
    End If
    Me.Height = Me.Height - i
    sstInfo.Height = sstInfo.Height - i
    chkWarn(chkȡѪ֪ͨ).Visible = gblnѪ��ϵͳ
    chkWarn(chk��Ѫ���).Visible = gblnѪ��ϵͳ
    chkWarn(chkѪ������).Visible = gblnѪ��ϵͳ
    
    
    'סԺҽ���´�ǿ��ȱʡҩ��
    chkȱʡҩ��.Value = Val(zlDatabase.GetPara("סԺҽ���´�ǿ��ȱʡҩ��", glngSys, pסԺҽ���´�, "1", Array(chkȱʡҩ��), intType))
    
    '�Զ�ˢ��ҽ������
    If mbln��ʿվ Then
        strPar = zlDatabase.GetPara("�Զ�ˢ��ҽ�����", glngSys, mlngModual, , Array(chkNotifyAdvice), InStr(mstrPrivs, "��������") > 0, intType)
        If Val(strPar) > 0 Then
            chkNotifyAdvice.Value = 1: txtNotifyAdvice.Text = Val(strPar)
        End If
        'ǰ���¼��л��Զ����ã���˺���ǿ������
        If (intType = 3 Or intType = 15) And InStr(mstrPrivs, "��������") = 0 Then
            txtNotifyAdvice.Enabled = False
        End If
        
        strPar = zlDatabase.GetPara("�Զ�ˢ��ҽ������", glngSys, mlngModual, 1, Array(lblNotifyAdviceDay, txtNotifyAdviceDay), InStr(mstrPrivs, "��������") > 0)
        txtNotifyAdviceDay.Text = Val(strPar)
        
        strPar = zlDatabase.GetPara("�Զ�ˢ��ҽ������", glngSys, mlngModual, "000000000000", Array(lbl��������, chkWarn(0), chkWarn(1), chkWarn(2), chkWarn(3), chkWarn(4), chkWarn(5), chkWarn(6), chkWarn(7), chkWarn(8), chkWarn(9), chkWarn(10), chkWarn(11), chkWarn(12)), InStr(mstrPrivs, "��������") > 0)
        For i = 1 To Len(strPar)
            chkWarn(i - 1).Value = IIf(Val(Mid(strPar, i, 1)) = 1, 1, 0)
        Next
        txtMedRec.Text = zlDatabase.GetPara("������鷴������", glngSys, mlngModual, "3", Array(lblMedRec, txtMedRec), InStr(mstrPrivs, "��������") > 0)
        
        chkSoundHS.Value = Val(zlDatabase.GetPara("����������ʾ", glngSys, mlngModual, , Array(chkSoundHS, cmdSoundHSSet), InStr(mstrPrivs, "��������") > 0, intType))
        
    Else
        strPar = zlDatabase.GetPara("�Զ�ˢ�²������ļ��", glngSys, mlngModual, , Array(chkNotifyEPR), InStr(mstrPrivs, "��������") > 0, intType)
        If Val(strPar) > 0 Then
            chkNotifyEPR.Value = 1: txtNotifyEPR.Text = Val(strPar)
        End If
        'ǰ���¼��л��Զ����ã���˺���ǿ������
        If (intType = 3 Or intType = 15) And InStr(mstrPrivs, "��������") = 0 Then
            txtNotifyEPR.Enabled = False
        End If
        
        strPar = zlDatabase.GetPara("�Զ�ˢ�²�����������", glngSys, mlngModual, 1, Array(lblNotifyEPRDay, txtNotifyEPRDay), InStr(mstrPrivs, "��������") > 0)
        txtNotifyEPRDay.Text = Val(strPar)
        opt�����ж�(Val(zlDatabase.GetPara("�����ж����", glngSys, pסԺҽ��վ, 0, Array(opt�����ж�(0), opt�����ж�(1), opt�����ж�(2), lbl�����ж�), InStr(mstrPrivs, "��������") > 0) & "")).Value = True
        opt�������(Val(zlDatabase.GetPara("������ϼ��", glngSys, pסԺҽ��վ, 0, Array(opt�������(0), opt�������(1), opt�������(2), lbl�������), InStr(mstrPrivs, "��������") > 0) & "")).Value = True
        optICD����(Val(zlDatabase.GetPara("ICD������", glngSys, pסԺҽ��վ, 0, Array(optICD����(0), optICD����(1), optICD����(2), lblICD����), InStr(mstrPrivs, "��������") > 0) & "")).Value = True
        opt����(Val(zlDatabase.GetPara("������", glngSys, pסԺҽ��վ, 1, Array(opt����(0), lbl����, opt����(1), opt����(2)), InStr(mstrPrivs, "��������") > 0) & "")).Value = True
        With vsfMain
            vsfMain.Rows = 1
            vsfMain.Cols = 3
            .TextMatrix(0, 0) = "����"
            .TextMatrix(0, 1) = "����"
            .TextMatrix(0, 2) = "����"
            .ColWidth(0) = 1400
            .ColWidth(1) = 2500
            .ColWidth(2) = 2500
            .Cell(flexcpAlignment, 0, 0, 0, 2) = 4
        End With
        
        varTmp = Array(chkWarn(chkD��������), chkWarn(chkDҽ������), chkWarn(chkDΣ��ֵ), chkWarn(chkD���泷��), chkWarn(chkDҽ�����), chkWarn(chkD�������), chkWarn(chkD��Ⱦ��), chkWarn(chkD�����ʿ�), chkWarn(chkD��Ѫ���), chkWarn(chkDУ������), chkWarn(chkD��Ѫ���), chkWarn(chkD��Ѫ��Ӧ), lblArea)
        strNotify = zlDatabase.GetPara("�Զ�ˢ������", glngSys, pסԺҽ��վ, , varTmp, InStr(mstrPrivs, "��������") > 0)
            
        chkWarn(chkD��������).Value = Val(Mid(strNotify, 1, 1))
        chkWarn(chkDҽ������).Value = Val(Mid(strNotify, 2, 1))
        chkWarn(chkDΣ��ֵ).Value = Val(Mid(strNotify, 3, 1))
        chkWarn(chkD���泷��).Value = Val(Mid(strNotify, 4, 1))
        chkWarn(chkDҽ�����).Value = Val(Mid(strNotify, 5, 1))
        chkWarn(chkD�������).Value = Val(Mid(strNotify, 6, 1))
        chkWarn(chkD��Ⱦ��).Value = Val(Mid(strNotify, 7, 1))
        chkWarn(chkD�����ʿ�).Value = Val(Mid(strNotify, 8, 1))
        chkWarn(chkD��Ѫ���).Value = Val(Mid(strNotify, 9, 1))
        chkWarn(chkD��Ѫ���).Visible = gblnѪ��ϵͳ
        chkWarn(chkDУ������).Value = Val(Mid(strNotify, 10, 1))
        chkWarn(chkD��Ѫ���).Value = Val(Mid(strNotify, 11, 1))
        chkWarn(chkD��Ѫ���).Visible = gblnѪ��ϵͳ
        chkWarn(chkD��Ѫ��Ӧ).Value = Val(Mid(strNotify, 12, 1))
        chkWarn(chkD��Ѫ��Ӧ).Visible = gblnѪ��ϵͳ
        
        Call Get������Ŀ
        cboType.Clear
        cboType.AddItem "0-��������׼"
        cboType.AddItem "1-�Ĵ�ʡ��׼"
        cboType.AddItem "2-����ʡ��׼"
        cboType.AddItem "3-����ʡ��׼"
        Call zlControl.CboSetIndex(cboType.hwnd, Val(zlDatabase.GetPara("������ҳ��׼", glngSys, pסԺҽ��վ, "0", Array(cboType, lbl��ҳ��׼), InStr(mstrPrivs, "��������") > 0)))
        Call cboType_Click
        If InStr(mstrPrivs, "��������") = 0 Then
        
            chkWarn(chkD��������).Enabled = False
            chkWarn(chkDҽ������).Enabled = False
            chkWarn(chkDΣ��ֵ).Enabled = False
            chkWarn(chkD���泷��).Enabled = False
            chkWarn(chkDҽ�����).Enabled = False
            chkWarn(chkD�������).Enabled = False
            chkWarn(chkD��Ⱦ��).Enabled = False
            chkWarn(chkD�����ʿ�).Enabled = False
            chkWarn(chkD��Ѫ���).Enabled = False
            chkWarn(chkDУ������).Enabled = False
            chkWarn(chkD��Ѫ���).Enabled = False
            chkWarn(chkD��Ѫ��Ӧ).Enabled = False
            cmdAdd.Enabled = False
            cmdModify.Enabled = False
            cmdDelete.Enabled = False
        End If
        Chk����.Value = Val(zlDatabase.GetPara("�������ֻ����¼��������̬ѧ����", glngSys, mlngModual, 0, Array(Chk����), InStr(mstrPrivs, "��������") > 0))
        txtMedRec.Text = zlDatabase.GetPara("������鷴������", glngSys, mlngModual, "3", Array(lblMedRec, txtMedRec), InStr(mstrPrivs, "��������") > 0)
        chkʹ����������ʱ��.Value = Val(zlDatabase.GetPara("ʹ����������ʱ��", glngSys, mlngModual, 0, Array(chkʹ����������ʱ��), InStr(mstrPrivs, "��������") > 0))
        chk��ҽ.Value = Val(zlDatabase.GetPara("��ҽ���Ҳ�ʹ����ҽ������ҳ��Ŀ", glngSys, mlngModual, 0, Array(chk��ҽ), InStr(mstrPrivs, "��������") > 0))
        chkSeparEdit.Value = Val(zlDatabase.GetPara("ҽ���ͻ�ʿ�ֱ���д������ҳ", glngSys, mlngModual, 0, Array(chkSeparEdit), InStr(mstrPrivs, "��������") > 0))
        chkDeptView.Value = Val(zlDatabase.GetPara("����ʾ�޴�λ�Ĳ�������", glngSys, mlngModual, 0, Array(chkDeptView), InStr(mstrPrivs, "��������") > 0))
        chkSoundYS.Value = Val(zlDatabase.GetPara("����������ʾ", glngSys, mlngModual, , Array(chkSoundYS, cmdSoundYSSet), InStr(mstrPrivs, "��������") > 0))
        chkGet����.Value = Val(zlDatabase.GetPara("����Զ���ȡ����", glngSys, 0, 0, Array(chkGet����), InStr(mstrPrivs, "��������") > 0))
        chkZLZD.Value = Val(zlDatabase.GetPara("��Ժ���������д�ֻ��̶Ⱥ�����������", glngSys, 0, 0, Array(chkZLZD), InStr(mstrPrivs, "��������") > 0))
    End If
End Sub

Private Sub Get������Ŀ()
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long, strSQL As String
    
    strSQL = "select ����,����,���� from ������Ŀ order by ����"
On Error GoTo errHandle
    Call zlDatabase.OpenRecordset(rsTemp, strSQL, Me.Caption)
    If rsTemp.RecordCount = 0 Then cmdModify.Enabled = False: cmdDelete.Enabled = False: Exit Sub
    lngRow = 1
    While Not rsTemp.EOF
        With vsfMain
            .Rows = lngRow + 1
            .TextMatrix(lngRow, 0) = rsTemp!���� & ""
            .TextMatrix(lngRow, 1) = rsTemp!���� & ""
            .TextMatrix(lngRow, 2) = rsTemp!���� & ""
        End With
        lngRow = lngRow + 1
        rsTemp.MoveNext
    Wend
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbln��ʿվ = False
End Sub

Private Sub optICD����_Click(Index As Integer)
    lngICD���� = Index
End Sub

Private Sub opt�������_Click(Index As Integer)
    lng������� = Index
End Sub

Private Sub opt����_Click(Index As Integer)
    lng���� = Index
End Sub

Private Sub opt�����ж�_Click(Index As Integer)
    lng�����ж� = Index
End Sub

Private Sub sstInfo_Click(PreviousTab As Integer)
    If sstInfo.Tab = 1 Then
        vsfMain.SetFocus
        If vsfMain.Rows > 1 Then vsfMain.Row = 1
        cmdModify.Enabled = True
        cmdDelete.Enabled = True
    End If
End Sub

Private Sub txtMedRec_GotFocus()
    Call zlControl.TxtSelAll(txtMedRec)
End Sub

Private Sub txtMedRec_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNotifyAdvice_GotFocus()
    Call zlControl.TxtSelAll(txtNotifyAdvice)
End Sub

Private Sub txtNotifyAdvice_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNotifyEPR_GotFocus()
    Call zlControl.TxtSelAll(txtNotifyEPR)
End Sub

Private Sub txtNotifyEPR_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNotifyAdviceDay_GotFocus()
    Call zlControl.TxtSelAll(txtNotifyAdviceDay)
End Sub

Private Sub txtNotifyAdviceDay_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtNotifyEPRDay_GotFocus()
    Call zlControl.TxtSelAll(txtNotifyEPRDay)
End Sub

Private Sub txtNotifyEPRDay_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub vsfMain_DblClick()
    Call cmdModify_Click
End Sub

Private Sub vsfMain_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If vsfMain.Row > 0 Then
            Call cmdDelete_Click
        End If
    End If
End Sub

Private Function CheckVsf() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "select count(��Ϣ��) as ���� from ������ҳ�ӱ� where ��Ϣ��='" & vsfMain.TextMatrix(vsfMain.Row, 1) & "'"
    
    err = 0: On Error GoTo errHandle
    Call zlDatabase.OpenRecordset(rsTemp, strSQL, Me.Caption)
    If rsTemp!���� > 0 Then
        MsgBox "����Ŀ�Ѿ�ʹ��,���ܽ����޸Ļ�ɾ��!"
        CheckVsf = False
        Exit Function
    End If
    CheckVsf = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
