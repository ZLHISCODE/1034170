VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm�������� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   Icon            =   "frm��������.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin TabDlg.SSTab tabMain 
      Height          =   7215
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   12726
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "����(&0)"
      TabPicture(0)   =   "frm��������.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl����"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl��ѯ����"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraͬ��������"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fra�ⷿѡ��"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fra�ɱ���"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "chkALLPlanPoint"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "fra�ƿ����̿���"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "fra�̵�ʱ�䷶Χ"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "fra�ϴβɹ���Ϣ"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "fra�ɹ��ƻ�"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txt��ѯ����"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "����У��(&1)"
      TabPicture(1)   =   "frm��������.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraCheck"
      Tab(1).Control(1)=   "vsfCheck"
      Tab(1).Control(2)=   "lblComment"
      Tab(1).ControlCount=   3
      Begin VB.TextBox txt��ѯ���� 
         Height          =   300
         Left            =   3240
         TabIndex        =   66
         Text            =   "7"
         Top             =   6000
         Width           =   540
      End
      Begin VB.Frame fra�ɹ��ƻ� 
         Caption         =   "��������"
         Height          =   1485
         Left            =   240
         TabIndex        =   4
         Top             =   5160
         Width           =   7485
         Begin VB.ComboBox cbo��Ӧ��ѡ�� 
            Height          =   300
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   300
            Width           =   2655
         End
         Begin VB.ComboBox cbo��Ӧ�̷�Χ 
            Height          =   300
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   660
            Width           =   2655
         End
         Begin VB.Label lbl��Ӧ��ѡ�� 
            AutoSize        =   -1  'True
            Caption         =   "��Ӧ��Ĭ��ѡ��"
            Height          =   180
            Left            =   360
            TabIndex        =   9
            Top             =   360
            Width           =   1260
         End
         Begin VB.Label lbl��Ӧ�̷�Χ 
            AutoSize        =   -1  'True
            Caption         =   "��Ӧ��ѡ��Χ"
            Height          =   180
            Left            =   360
            TabIndex        =   8
            Top             =   720
            Width           =   1260
         End
         Begin VB.Label Label3 
            Caption         =   "    ҩƷ�ɹ��ƻ��༭������ҩƷ��Ӧ�̵�Ĭ�ϴ����Լ��ֹ�ѡ��Ӧ��ʱ�Ŀ�ѡ��Χ��"
            Height          =   855
            Left            =   4680
            TabIndex        =   7
            Top             =   360
            Width           =   2085
         End
      End
      Begin VB.Frame fra�ϴβɹ���Ϣ 
         Caption         =   "�ϴβɹ���Ϣ��Դ��ʽ"
         ForeColor       =   &H00800000&
         Height          =   1005
         Left            =   240
         TabIndex        =   17
         Top             =   5280
         Visible         =   0   'False
         Width           =   7485
         Begin VB.OptionButton optȡ�ɱ��۷�ʽ 
            Caption         =   "���ȴ���һ�����ҵ����ȡ�ɱ��۵���Ϣ"
            Height          =   180
            Index           =   1
            Left            =   240
            TabIndex        =   19
            Top             =   720
            Width           =   6615
         End
         Begin VB.OptionButton optȡ�ɱ��۷�ʽ 
            Caption         =   "���ȴӵ�ǰ�ⷿ�Ŀ�����������ȡ�ɱ��۵���Ϣ"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   18
            Top             =   360
            Value           =   -1  'True
            Width           =   6615
         End
      End
      Begin VB.Frame fra�̵�ʱ�䷶Χ 
         Caption         =   "�̵�ʱ�䷶Χ����"
         Height          =   735
         Left            =   240
         TabIndex        =   58
         Top             =   5280
         Visible         =   0   'False
         Width           =   3675
         Begin VB.TextBox txt�̵�ʱ�� 
            Height          =   300
            Left            =   1560
            TabIndex        =   60
            Top             =   240
            Width           =   705
         End
         Begin MSComCtl2.UpDown UpD�̵�ʱ�� 
            Height          =   300
            Left            =   2266
            TabIndex        =   59
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            BuddyControl    =   "txt�̵�ʱ��"
            BuddyDispid     =   196620
            OrigLeft        =   1800
            OrigTop         =   360
            OrigRight       =   2055
            OrigBottom      =   735
            Max             =   90
            Enabled         =   -1  'True
         End
         Begin VB.Label lblday 
            Caption         =   "��"
            Height          =   195
            Left            =   2880
            TabIndex        =   61
            Top             =   300
            Width           =   255
         End
      End
      Begin VB.Frame fra�ƿ����̿��� 
         Caption         =   "�ƿ����̿���"
         Height          =   1485
         Left            =   240
         TabIndex        =   10
         Top             =   5280
         Width           =   7485
         Begin VB.CheckBox chk�ƿ����̿��� 
            Caption         =   "�ƿ�ʱ��Ҫ��ҩ�����͡�������һ���̡�"
            Height          =   180
            Left            =   180
            TabIndex        =   12
            Top             =   270
            Value           =   1  'Checked
            Width           =   6945
         End
         Begin VB.CheckBox chkRequestStrike 
            Caption         =   "�ƿ����ʱ������ⷿ��Ҫ���������"
            Height          =   180
            Left            =   180
            TabIndex        =   11
            Top             =   1080
            Width           =   5895
         End
         Begin VB.Label Label1 
            Caption         =   "ע�⣺�������ѡ����ô����д�ƿⵥ������һ����˲�������˺��Զ���ɱ�ҩ�����͡�������һ���̡����ǰ�����޸ĵ��ݡ�"
            Height          =   375
            Left            =   450
            TabIndex        =   13
            Top             =   540
            Width           =   6945
         End
      End
      Begin VB.CheckBox chkALLPlanPoint 
         Caption         =   "ȫԺ�ƻ�����վ��"
         Height          =   255
         Left            =   240
         TabIndex        =   65
         Top             =   6000
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Frame fra�ɱ��� 
         Caption         =   "�ɱ�����Դ��ʽ"
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   240
         TabIndex        =   54
         Top             =   5280
         Width           =   7665
         Begin VB.OptionButton opt�ɱ���Դ 
            Caption         =   $"frm��������.frx":0044
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   1
            Left            =   3360
            TabIndex        =   56
            Top             =   277
            Width           =   3735
         End
         Begin VB.OptionButton opt�ɱ���Դ 
            Caption         =   "����ԭ��ҩƷ�ĳɱ��ۼ���"
            ForeColor       =   &H00800000&
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   55
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame fraCheck 
         Caption         =   "ѡ��У�鷽ʽ"
         Height          =   615
         Left            =   -74760
         TabIndex        =   47
         Top             =   5160
         Width           =   7350
         Begin VB.OptionButton optCheck 
            Caption         =   "У��δͨ��ʱ��ֹ����"
            Height          =   180
            Index           =   0
            Left            =   360
            TabIndex        =   49
            Top             =   280
            Width           =   2175
         End
         Begin VB.OptionButton optCheck 
            Caption         =   "У��δͨ��ʱ����"
            Height          =   180
            Index           =   1
            Left            =   3360
            TabIndex        =   48
            Top             =   280
            Width           =   2175
         End
      End
      Begin VB.Frame fra�ⷿѡ�� 
         Caption         =   "�ⷿѡ��"
         Height          =   1665
         Left            =   240
         TabIndex        =   43
         Top             =   480
         Width           =   3675
         Begin VB.CheckBox chkStock 
            Caption         =   "����ѡ��ⷿ"
            Height          =   375
            Left            =   210
            TabIndex        =   44
            Top             =   240
            Width           =   2805
         End
         Begin VB.Label Label4 
            Caption         =   "    ���ѡ��ⷿ�����ڵ�������'���пⷿ'Ȩ���˾Ϳ���ѡ��ͬ�ⷿ�����򣬲���ѡ��ⷿ��"
            Height          =   615
            Left            =   120
            TabIndex        =   45
            Top             =   720
            Width           =   3285
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "ҩƷ��λ"
         Enabled         =   0   'False
         Height          =   1665
         Left            =   3960
         TabIndex        =   37
         Top             =   480
         Width           =   3675
         Begin VB.ComboBox cboUnit 
            Height          =   300
            Left            =   870
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   390
            Width           =   2655
         End
         Begin VB.ComboBox CboUnit1 
            Height          =   300
            Left            =   870
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   780
            Width           =   2655
         End
         Begin VB.Label lblUnitComment 
            Caption         =   "    ��ѡ��һ��ҩƷ��λ���ڵ��������У�����ҩƷ�������ֵ�λ��"
            Height          =   405
            Left            =   120
            TabIndex        =   42
            Top             =   1200
            Width           =   3315
         End
         Begin VB.Label lbl�̵�� 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "���װ"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   210
            TabIndex        =   41
            Top             =   450
            Width           =   540
         End
         Begin VB.Label lbl�̵㵥 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "С��װ"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   210
            TabIndex        =   40
            Top             =   840
            Width           =   540
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "����"
         Height          =   2835
         Left            =   3960
         TabIndex        =   24
         Top             =   2160
         Width           =   3675
         Begin VB.Frame fra����������ʱ��ʽ 
            Caption         =   "����������ʱ��ʽ"
            Height          =   735
            Left            =   120
            TabIndex        =   72
            Top             =   1680
            Visible         =   0   'False
            Width           =   3450
            Begin VB.OptionButton Opt����ʵ������ 
               Caption         =   "����ʵ������"
               Height          =   180
               Left            =   120
               TabIndex        =   74
               Top             =   375
               Width           =   1380
            End
            Begin VB.OptionButton Opt���ÿ������� 
               Caption         =   "���ÿ�������"
               Height          =   180
               Left            =   1680
               TabIndex        =   73
               Top             =   375
               Width           =   1440
            End
         End
         Begin VB.CheckBox chk�ֶμӳ���� 
            Caption         =   "ʱ��ҩƷ���ֶμӳ����"
            Height          =   255
            Left            =   165
            TabIndex        =   71
            Top             =   2040
            Visible         =   0   'False
            Width           =   2880
         End
         Begin VB.CheckBox chkȡ�ϴ��ۼ� 
            Caption         =   "ʱ��ҩƷ���ʱȡ�ϴ��ۼ�"
            Height          =   255
            Left            =   165
            TabIndex        =   70
            Top             =   1800
            Visible         =   0   'False
            Width           =   2880
         End
         Begin VB.CheckBox chk�ӳ���� 
            Caption         =   "ʱ��ҩƷ�Լӳ������"
            Height          =   255
            Left            =   165
            TabIndex        =   69
            Top             =   1560
            Visible         =   0   'False
            Width           =   2880
         End
         Begin VB.CheckBox chk�б�ҩƷ 
            Caption         =   "�б�ҩƷ��ѡ����б굥λ���"
            Height          =   255
            Left            =   165
            TabIndex        =   30
            Top             =   1335
            Visible         =   0   'False
            Width           =   2880
         End
         Begin VB.CheckBox chk������� 
            Caption         =   "����ҩƷ�������"
            Height          =   255
            Left            =   165
            TabIndex        =   53
            Top             =   840
            Visible         =   0   'False
            Width           =   1920
         End
         Begin VB.CheckBox chkʱ�۵��� 
            Caption         =   "ʱ��ҩƷ�����ε���"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   165
            TabIndex        =   63
            Top             =   960
            Visible         =   0   'False
            Width           =   2010
         End
         Begin VB.CheckBox chk�޼���ʾ 
            Caption         =   "�³ɱ��ۡ����ۼ۳����޼�ʱ��ʾ"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   240
            TabIndex        =   62
            Top             =   1320
            Visible         =   0   'False
            Width           =   3090
         End
         Begin VB.CheckBox chkStopDrug 
            Caption         =   "�̵�ͣ��ҩƷ"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   165
            TabIndex        =   57
            Top             =   1100
            Visible         =   0   'False
            Width           =   1920
         End
         Begin VB.CheckBox chkVerifyPrint 
            Caption         =   "��˺��ӡ"
            Height          =   255
            Left            =   2010
            TabIndex        =   35
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox chkFixPrice 
            Caption         =   "���۲ɹ�"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   165
            TabIndex        =   34
            Top             =   495
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.CheckBox Chk�����޸������� 
            Caption         =   "�޸Ĳɹ��޼�"
            Height          =   255
            Left            =   2010
            TabIndex        =   33
            Top             =   513
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.CommandButton cmd��ӡ���� 
            Caption         =   "��ӡ����(&P)"
            Height          =   315
            Left            =   270
            TabIndex        =   32
            Top             =   2400
            Width           =   3135
         End
         Begin VB.CheckBox chk�⹺NO 
            Caption         =   "�޸ĵ��ݺ�"
            Height          =   255
            Left            =   165
            TabIndex        =   29
            Top             =   780
            Visible         =   0   'False
            Width           =   1305
         End
         Begin VB.CheckBox chkSavePrint 
            Caption         =   "���̺��ӡ"
            Height          =   255
            Left            =   165
            TabIndex        =   36
            Top             =   240
            Width           =   1275
         End
         Begin VB.CheckBox Chk�洢�ⷿ 
            Caption         =   "�����̵�û�����ô洢�ⷿ��ҩƷ"
            Height          =   255
            Left            =   165
            TabIndex        =   46
            Top             =   540
            Visible         =   0   'False
            Width           =   3360
         End
         Begin VB.CheckBox chkSendPrint 
            Caption         =   "���ͺ��ӡ"
            Height          =   255
            Left            =   165
            TabIndex        =   52
            Top             =   495
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Frame Frame�۸���ʾ 
            Caption         =   "�۸���ʾ��ʽ"
            Height          =   735
            Left            =   120
            TabIndex        =   25
            Top             =   705
            Visible         =   0   'False
            Width           =   3450
            Begin VB.OptionButton Opt��� 
               Caption         =   "�ɱ��ۺ��ۼ�"
               Height          =   180
               Left            =   1950
               TabIndex        =   27
               Top             =   375
               Width           =   1400
            End
            Begin VB.OptionButton Opt�ɱ��� 
               Caption         =   "�ɱ���"
               Height          =   180
               Left            =   45
               TabIndex        =   28
               Top             =   375
               Width           =   900
            End
            Begin VB.OptionButton Opt�ۼ� 
               Caption         =   "�ۼ�"
               Height          =   180
               Left            =   1100
               TabIndex        =   26
               Top             =   375
               Width           =   720
            End
         End
         Begin VB.CheckBox chk�������� 
            Caption         =   "������������"
            Height          =   255
            Left            =   165
            TabIndex        =   64
            Top             =   1680
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CheckBox chk�˻���Ʊ��� 
            Caption         =   "�˻�ʱ��Ʊ��������۽��Ϊ׼"
            Height          =   255
            Left            =   165
            TabIndex        =   31
            Top             =   1059
            Visible         =   0   'False
            Width           =   3015
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "����ʽ"
         Height          =   2835
         Left            =   240
         TabIndex        =   20
         Top             =   2160
         Width           =   3675
         Begin VB.ComboBox Cbo���� 
            Height          =   300
            ItemData        =   "frm��������.frx":0066
            Left            =   120
            List            =   "frm��������.frx":0068
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   390
            Width           =   2415
         End
         Begin VB.ComboBox Cbo���� 
            Height          =   300
            Left            =   2700
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   390
            Width           =   885
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "    �����������ã���Ӱ�����б༭�����е��ݵ���ʾ���ݵ�����ʽ��ȱʡ�����û������˳����ʾ�����ݵ�����"
            ForeColor       =   &H80000008&
            Height          =   825
            Left            =   180
            TabIndex        =   23
            Top             =   1080
            Width           =   3345
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfCheck 
         Height          =   4125
         Left            =   -74760
         TabIndex        =   50
         Top             =   960
         Width           =   7095
         _cx             =   12515
         _cy             =   7276
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
         BackColorSel    =   16711680
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   3
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   13
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frm��������.frx":006A
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
      Begin VB.Frame fraͬ�������� 
         Caption         =   "ͬ�����������"
         ForeColor       =   &H00800000&
         Height          =   1485
         Left            =   240
         TabIndex        =   14
         Top             =   5280
         Visible         =   0   'False
         Width           =   7485
         Begin VB.CheckBox chk������ 
            Caption         =   "ҩƷ�����������ʱͬ�����ٿ��"
            Height          =   180
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   3495
         End
         Begin VB.Label Label2 
            Caption         =   "˵���������ѡ��ѡ��൱������˺��Զ�����������������Ҫʵ�ָù��ܣ�����ȷ���������������������������"
            ForeColor       =   &H00800000&
            Height          =   540
            Left            =   360
            TabIndex        =   16
            Top             =   840
            Width           =   7020
         End
      End
      Begin VB.Label lbl��ѯ���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��ѯ����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2400
         TabIndex        =   67
         Top             =   6060
         Width           =   720
      End
      Begin VB.Label lbl���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3840
         TabIndex        =   68
         Top             =   6060
         Width           =   180
      End
      Begin VB.Label lblComment 
         Caption         =   "    ˵����ҩƷ�⹺���༭����ʱ�Ƿ�У��Ӧ�̵���Ϣ�Ƿ��������������Ƿ���ڡ���ѡ����Ҫ����У�����Ŀ����˫����У�顱�д򹴡�"
         Height          =   540
         Left            =   -74760
         TabIndex        =   51
         Top             =   480
         Width           =   7140
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6930
      TabIndex        =   2
      Top             =   7560
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   5730
      TabIndex        =   1
      Top             =   7560
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   240
      TabIndex        =   0
      Top             =   7560
      Width           =   1100
   End
End
Attribute VB_Name = "frm��������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFunction As String
Dim mstrPrivs As String
Dim mlngModul As Long
Dim mblnSetPara As Boolean      '�Ƿ���в�������Ȩ��
Private mint�̵�ʱ�� As Integer  '������¼���õ��̵�ʱ�䷶Χ

Private Sub Cbo����_Click()
    If Cbo����.ListCount < 1 Then Exit Sub
    Cbo����.Enabled = Not (Cbo����.ListIndex = 0)
    If Not Cbo����.Enabled Then Cbo����.ListIndex = 0
End Sub

Private Sub chkRequestStrike_Click()
    '����Ϊ����Ҫ����ʱ��Ҫ����Ƿ���δ��˵ĳ������뵥����������ܸı�
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    If chkRequestStrike.Value = 0 Then
        If MsgBox("��������Ƿ����δ��˵ĳ������뵥��������Ҫ�ϳ�ʱ�䣬�Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            '�ù�����10.20�汾����������һ�������������ڷ�Χ������ȫ��ɨ��
            gstrSQL = "Select 1 From ҩƷ�շ���¼ Where ���� = 6 And Mod(��¼״̬, 3) = 2 And ������� Is Null " & _
                " And �������� Between To_Date('2008/3/6 00:00:00', 'yyyy-mm-dd hh24:mi:ss') And Sysdate And Rownum = 1"
            
            DoEvents
            zlCommFun.ShowFlash "���ڲ�������,���Ժ�...", Me
            
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ���δ��˵ĳ������뵥")
            
            DoEvents
            zlCommFun.StopFlash
            
            If rsTemp.RecordCount > 0 Then
                MsgBox "����δ��˵ĳ������뵥�����ܸı�˲�����", vbInformation, gstrSysName
                chkRequestStrike.Value = 1
            End If
        Else
            chkRequestStrike.Value = 1
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub chk�ֶμӳ����_Click()
    If chk�ֶμӳ����.Value = 1 Then
        chk�ӳ����.Value = 0
        chkȡ�ϴ��ۼ�.Value = 0
    End If
End Sub

Private Sub chk�ӳ����_Click()
    If chk�ӳ����.Value = 1 Then
        chkȡ�ϴ��ۼ�.Value = 0
        chk�ֶμӳ����.Value = 0
    End If
End Sub

Private Sub chk��������_Click()
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errh
    
    If chk��������.Value = 0 Then
        gstrSQL = "Select �ڼ� From ҩƷ���� Where Length(�ڼ�) > 4"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        If rsTemp.RecordCount > 0 Then
            MsgBox "��������ģʽ���Ѿ��������ݣ������޸ģ�", vbInformation, gstrSysName
            chk��������.Value = 1
        End If
    End If
    Exit Sub
errh:
If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub chkȡ�ϴ��ۼ�_Click()
    If chkȡ�ϴ��ۼ�.Value = 1 Then
        chk�ӳ����.Value = 0
        chk�ֶμӳ����.Value = 0
    End If
End Sub

Private Sub chk�ƿ����̿���_Click()
    If chk�ƿ����̿���.Value = 1 Then
        chkSendPrint.Visible = True
    Else
        chkSendPrint.Visible = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cmdOk_Click()
    On Error Resume Next
    
    If ISValid = False Then Exit Sub
    
    Select Case mlngModul
        Case 1300   'ҩƷ�⹺������
            zlDatabase.SetPara "�Ƿ�ѡ��ⷿ", IIf(chkStock.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "����", CStr(Cbo����.ListIndex) & CStr(Cbo����.ListIndex), glngSys, mlngModul
            zlDatabase.SetPara "���̴�ӡ", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "��˴�ӡ", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "ҩƷ��λ", cboUnit.ListIndex, glngSys, mlngModul
            
            zlDatabase.SetPara "���۲ɹ�", chkFixPrice.Value, glngSys, mlngModul
            zlDatabase.SetPara "�޸��⹺���ݺ�", chk�⹺NO.Value, glngSys, mlngModul
            zlDatabase.SetPara "�޸Ĳɹ��޼�", Chk�����޸�������.Value, glngSys, mlngModul
            zlDatabase.SetPara "�б�ҩƷ��ѡ����б굥λ���", chk�б�ҩƷ.Value, glngSys, mlngModul
            zlDatabase.SetPara "�˻���Ʊ���", chk�˻���Ʊ���.Value, glngSys, mlngModul
            zlDatabase.SetPara "ȡ�ϴβɹ��۷�ʽ", IIf(optȡ�ɱ��۷�ʽ(0).Value, "0", "1"), glngSys, mlngModul
            zlDatabase.SetPara "��ѯ����", Val(txt��ѯ����.Text), glngSys, mlngModul
            
            zlDatabase.SetPara "ʱ��ҩƷ�ԼӼ������", chk�ӳ����.Value, glngSys, mlngModul
            zlDatabase.SetPara "ʱ��ҩƷ���ʱȡ�ϴ��ۼ�", chkȡ�ϴ��ۼ�.Value, glngSys, mlngModul
            zlDatabase.SetPara "ʱ��ҩƷ�����÷ֶμӳ�", chk�ֶμӳ����.Value, glngSys, mlngModul
            
            Save����У��
        Case 1301   'ҩƷ����������
            zlDatabase.SetPara "�Ƿ�ѡ��ⷿ", IIf(chkStock.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "����", CStr(Cbo����.ListIndex) & CStr(Cbo����.ListIndex), glngSys, mlngModul
            zlDatabase.SetPara "���̴�ӡ", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "��˴�ӡ", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "ҩƷ��λ", cboUnit.ListIndex, glngSys, mlngModul
            zlDatabase.SetPara "ҩƷ�������ɱ��ۼ��㷽ʽ", IIf(opt�ɱ���Դ(0).Value = True, "0", "1"), glngSys, mlngModul
            zlDatabase.SetPara "��ѯ����", Val(txt��ѯ����.Text), glngSys, mlngModul
        Case 1302   'ҩƷ����������
            zlDatabase.SetPara "�Ƿ�ѡ��ⷿ", IIf(chkStock.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "����", CStr(Cbo����.ListIndex) & CStr(Cbo����.ListIndex), glngSys, mlngModul
            zlDatabase.SetPara "���̴�ӡ", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "��˴�ӡ", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "ҩƷ��λ", cboUnit.ListIndex, glngSys, mlngModul
            zlDatabase.SetPara "��ѯ����", Val(txt��ѯ����.Text), glngSys, mlngModul
            
            zlDatabase.SetPara "ʱ��ҩƷ�ԼӼ������", chk�ӳ����.Value, glngSys, mlngModul
            zlDatabase.SetPara "ʱ��ҩƷ���ʱȡ�ϴ��ۼ�", chkȡ�ϴ��ۼ�.Value, glngSys, mlngModul
            zlDatabase.SetPara "ʱ��ҩƷ�����÷ֶμӳ�", chk�ֶμӳ����.Value, glngSys, mlngModul
        Case 1303   'ҩƷ����۵�������
            zlDatabase.SetPara "����", CStr(Cbo����.ListIndex) & CStr(Cbo����.ListIndex), glngSys, mlngModul
            zlDatabase.SetPara "���̴�ӡ", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "��˴�ӡ", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "ҩƷ��λ", cboUnit.ListIndex, glngSys, mlngModul
            zlDatabase.SetPara "��ѯ����", Val(txt��ѯ����.Text), glngSys, mlngModul
        Case 1304   'ҩƷ�ƿ����
            zlDatabase.SetPara "�Ƿ�ѡ��ⷿ", IIf(chkStock.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "ҩƷ��λ", cboUnit.ListIndex, glngSys, mlngModul
            zlDatabase.SetPara "����", CStr(Cbo����.ListIndex) & CStr(Cbo����.ListIndex), glngSys, mlngModul
            zlDatabase.SetPara "���̴�ӡ", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "��˴�ӡ", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "���ʹ�ӡ", IIf(chkSendPrint.Value = 1, "1", "0"), glngSys, mlngModul
            
            zlDatabase.SetPara "�ƿ�����", chk�ƿ����̿���.Value, glngSys, mlngModul
            zlDatabase.SetPara "��������", chkRequestStrike.Value, glngSys, mlngModul
            zlDatabase.SetPara "��ѯ����", Val(txt��ѯ����.Text), glngSys, mlngModul
        Case 1305   'ҩƷ���ù���
            zlDatabase.SetPara "�Ƿ�ѡ��ⷿ", IIf(chkStock.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "ҩƷ��λ", cboUnit.ListIndex, glngSys, mlngModul
            zlDatabase.SetPara "����", CStr(Cbo����.ListIndex) & CStr(Cbo����.ListIndex), glngSys, mlngModul
            zlDatabase.SetPara "���̴�ӡ", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "��˴�ӡ", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "������������", IIf(chk��������.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "��ѯ����", Val(txt��ѯ����.Text), glngSys, mlngModul
        Case 1306   'ҩƷ�����������
            zlDatabase.SetPara "�Ƿ�ѡ��ⷿ", IIf(chkStock.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "����", CStr(Cbo����.ListIndex) & CStr(Cbo����.ListIndex), glngSys, mlngModul
            zlDatabase.SetPara "���̴�ӡ", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "��˴�ӡ", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "ҩƷ��λ", cboUnit.ListIndex, glngSys, mlngModul
            zlDatabase.SetPara "��ѯ����", Val(txt��ѯ����.Text), glngSys, mlngModul
        Case 1307   'ҩƷ�̵����
            zlDatabase.SetPara "����", CStr(Cbo����.ListIndex) & CStr(Cbo����.ListIndex), glngSys, mlngModul
            zlDatabase.SetPara "С��װ��λ", CboUnit1.ListIndex, glngSys, mlngModul
            zlDatabase.SetPara "���̴�ӡ", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "��˴�ӡ", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngModul
                       
            zlDatabase.SetPara "�洢�ⷿ", IIf(Chk�洢�ⷿ.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "����ҩƷ�������", IIf(chk�������.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "����ͣ�õ�ҩƷ", IIf(chkStopDrug.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "�̵�ʱ�䷶Χ����", txt�̵�ʱ��.Text, glngSys, mlngModul
            zlDatabase.SetPara "��ѯ����", Val(txt��ѯ����.Text), glngSys, mlngModul
        Case 1330   'ҩƷ�ƻ�����
            zlDatabase.SetPara "�Ƿ�ѡ��ⷿ", IIf(chkStock.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "����", CStr(Cbo����.ListIndex) & CStr(Cbo����.ListIndex), glngSys, mlngModul
            zlDatabase.SetPara "�۸���ʾ��ʽ", IIf(Opt�ɱ���.Value = True, "0", IIf(Opt�ۼ�.Value = True, "1", "2")), glngSys, mlngModul
            zlDatabase.SetPara "���̴�ӡ", IIf(chkSavePrint.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "��˴�ӡ", IIf(chkVerifyPrint.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "��Ӧ��Ĭ��ѡ��", cbo��Ӧ��ѡ��.ListIndex, glngSys, mlngModul
            zlDatabase.SetPara "��Ӧ��ѡ��Χ", cbo��Ӧ�̷�Χ.ListIndex, glngSys, mlngModul
            zlDatabase.SetPara "ҩƷ��λ", cboUnit.ListIndex, glngSys, mlngModul
            zlDatabase.SetPara "ȫԺ�ƻ�����վ��", IIf(chkALLPlanPoint.Value = 1, "1", "0"), glngSys, mlngModul
            zlDatabase.SetPara "��ѯ����", Val(txt��ѯ����.Text), glngSys, mlngModul
            zlDatabase.SetPara "����������ʱ��ʽ", IIf(Opt����ʵ������.Value = True, "0", "1"), glngSys, mlngModul
        Case 1331   'ҩƷ��������
            zlDatabase.SetPara "���ʱ���ٿ��", chk������.Value, glngSys, mlngModul
            zlDatabase.SetPara "��ѯ����", Val(txt��ѯ����.Text), glngSys, mlngModul
        Case 1333 'ҩƷ���۹���
            zlDatabase.SetPara "����", CStr(Cbo����.ListIndex) & CStr(Cbo����.ListIndex), glngSys, mlngModul
            zlDatabase.SetPara "ʱ��ҩƷ�����ε���", chkʱ�۵���.Value, glngSys, mlngModul
            zlDatabase.SetPara "�޼���ʾ", chk�޼���ʾ.Value, glngSys, mlngModul
            zlDatabase.SetPara "ҩƷ��λ", cboUnit.ListIndex, glngSys, mlngModul
    End Select
           
    Unload Me
End Sub

Private Function ISValid() As Boolean
    Dim i As Integer
    Dim blnAllUnCheck As Boolean
    
    '����У��
    If tabMain.TabVisible(1) = True Then
        blnAllUnCheck = True
        With vsfCheck
            For i = 1 To .rows - 1
                If .TextMatrix(i, .ColIndex("У��")) <> "" Then
                    blnAllUnCheck = False
                    Exit For
                End If
            Next
        End With
        
        '���ѡ����У����Ŀ�������ѡ��У�鷽ʽ
        If blnAllUnCheck = False And optCheck(0).Value = 0 And optCheck(1).Value = 0 Then
            MsgBox "��ѡ������У�鷽ʽ��", vbExclamation, gstrSysName
            tabMain.Tab = 1
            If vsfCheck.Enabled Then vsfCheck.SetFocus
            Exit Function
        End If
    End If
    
    If Val(txt��ѯ����.Text) > 7 Then
        If MsgBox("��ѯʱ�����7����ܻᵼ�²�ѯ�������Ƿ������", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            txt��ѯ����.SetFocus
            zlControl.TxtSelAll txt��ѯ����
            Exit Function
        End If
    End If
    If Val(txt��ѯ����.Text) = 0 Then
        MsgBox "��ѯʱ��������0�����������룡", vbInformation, gstrSysName
        txt��ѯ����.SetFocus
        zlControl.TxtSelAll txt��ѯ����
        Exit Function
    End If
    
    ISValid = True
End Function

Private Sub Save����У��()
    Dim i As Integer
    Dim strCheck As String
    Dim blnAllUnCheck As Boolean
    
    If mlngModul <> 1300 Then Exit Sub
    
    blnAllUnCheck = True
    
    '��������У����Ŀ�ͷ�ʽ����ʽ��У�鷽ʽ|���1,��Ŀ1,�Ƿ�У��;���1,��Ŀ2,�Ƿ�У��;���2,��Ŀ1,�Ƿ�У��;���2,��Ŀ2....
    With vsfCheck
        For i = 1 To .rows - 1
            strCheck = IIf(strCheck = "", "", strCheck & ";") & .TextMatrix(i, .ColIndex("���")) & "," & .TextMatrix(i, .ColIndex("У����Ŀ")) & "," & _
                IIf(.TextMatrix(i, .ColIndex("У��")) = "", 0, 1)
                
            If .TextMatrix(i, .ColIndex("У��")) <> "" Then blnAllUnCheck = False
        Next
    End With
    
    If blnAllUnCheck = True Then
        strCheck = "0|" & strCheck
    ElseIf optCheck(0).Value = True Then
        strCheck = "2|" & strCheck
    Else
        strCheck = "1|" & strCheck
    End If
        
    Call zlDatabase.SetPara("����У��", strCheck, glngSys, mlngModul)
End Sub
Public Sub ���ò���(frmParent As Object, ByVal strPrivs As String, ByVal lngModual As Long, Optional ByVal strFunction As String = "")
    mstrFunction = strFunction
    mstrPrivs = strPrivs
    mlngModul = lngModual
    Dim str���ݴ�ӡ As String
    Dim int��ѯ���� As Integer
    
    'ͨ�ã�˽��ģ�飩
    Dim int�Ƿ�ѡ��ⷿ As Integer
    Dim str���� As String
    Dim int���̴�ӡ As Integer
    Dim int��˴�ӡ As Integer
        
    '������Ҫ��ͨģ�飨˽��ģ�飩
    Dim intҩƷ��λ As Integer
    Dim int�ɱ�����Դ As Integer
        
    '�����̵㣨˽��ģ�飩
    Dim intС��װ��λ As Integer
        
    '����ҩƷ�ƻ���˽��ģ�飩
    Dim int�۸���ʾ��ʽ As Integer
    Dim int��Ӧ��ѡ�� As Integer
    Dim int��Ӧ�̷�Χ As Integer
    Dim intPlanPoint As Integer
    Dim int����������ʱ��ʽ As Integer
    
    '�����⹺��⣨����ģ�飩
    Dim int���۲ɹ� As Integer
    Dim int�޸��⹺���ݺ� As Integer
    Dim int�޸������� As Integer
    Dim int�б�ҩƷ As Integer
    Dim int�˻���Ʊ��� As Integer
    Dim intȡ�ϴβɹ��۷�ʽ As Integer
    Dim int�ӳ������ As Integer    '��������Ҳ���������
    Dim intȡ�ϴ��ۼ� As Integer    '��������Ҳ���������
    Dim int�ֶμӳ���� As Integer  '��������Ҳ���������
    
    '�����ƿ⣨����ģ�飩
    Dim int�ƿ����� As Integer
    Dim int�������� As Integer
    
    '(˽��)
    Dim int���ʹ�ӡ As Integer
    
    '�����̵㣨����ģ�飩
    Dim int�洢�ⷿ As Integer
    Dim int���������� As Integer
    Dim int������� As Integer
    Dim int�̵�ͣ�� As Integer
    
    '����������������ģ�飩
    Dim int������ As Integer
    
    '��������
    Dim int�������� As Integer
    
    On Error Resume Next
    
    mblnSetPara = IsHavePrivs(mstrPrivs, "��������")
    
    'ȡ����ֵ
    Select Case mlngModul
        Case 1300   'ҩƷ�⹺������
            int�Ƿ�ѡ��ⷿ = Val(zlDatabase.GetPara("�Ƿ�ѡ��ⷿ", glngSys, mlngModul, 0, Array(fra�ⷿѡ��, chkStock, Label4), mblnSetPara))
            str���� = zlDatabase.GetPara("����", glngSys, mlngModul, "00", Array(Frame5, Cbo����, Cbo����, Label5), mblnSetPara)
            int���̴�ӡ = Val(zlDatabase.GetPara("���̴�ӡ", glngSys, mlngModul, 0, Array(chkSavePrint), mblnSetPara))
            int��˴�ӡ = Val(zlDatabase.GetPara("��˴�ӡ", glngSys, mlngModul, 0, Array(chkVerifyPrint), mblnSetPara))
            intҩƷ��λ = Val(zlDatabase.GetPara("ҩƷ��λ", glngSys, mlngModul, 0, Array(lbl�̵��, cboUnit), mblnSetPara))
            
            int���۲ɹ� = Val(zlDatabase.GetPara("���۲ɹ�", glngSys, mlngModul, 0, Array(chkFixPrice), mblnSetPara))
            int�޸��⹺���ݺ� = Val(zlDatabase.GetPara("�޸��⹺���ݺ�", glngSys, mlngModul, 0, Array(chk�⹺NO), mblnSetPara))
            int�޸������� = Val(zlDatabase.GetPara("�޸Ĳɹ��޼�", glngSys, mlngModul, 0, Array(Chk�����޸�������), mblnSetPara))
            int�б�ҩƷ = Val(zlDatabase.GetPara("�б�ҩƷ��ѡ����б굥λ���", glngSys, mlngModul, 0, Array(chk�б�ҩƷ), mblnSetPara))
            int�˻���Ʊ��� = Val(zlDatabase.GetPara("�˻���Ʊ���", glngSys, mlngModul, 1, Array(chk�˻���Ʊ���), mblnSetPara))
            intȡ�ϴβɹ��۷�ʽ = Val(zlDatabase.GetPara("ȡ�ϴβɹ��۷�ʽ", glngSys, mlngModul, 0, Array(fra�ϴβɹ���Ϣ, optȡ�ɱ��۷�ʽ(0), optȡ�ɱ��۷�ʽ(1)), mblnSetPara))
            int��ѯ���� = Val(zlDatabase.GetPara("��ѯ����", glngSys, mlngModul, 7))
            
            int�ӳ������ = Val(zlDatabase.GetPara("ʱ��ҩƷ�ԼӼ������", glngSys, mlngModul, 1, Array(chk�ӳ����), mblnSetPara))
            intȡ�ϴ��ۼ� = Val(zlDatabase.GetPara("ʱ��ҩƷ���ʱȡ�ϴ��ۼ�", glngSys, mlngModul, 0, Array(chkȡ�ϴ��ۼ�), mblnSetPara))
            int�ֶμӳ���� = Val(zlDatabase.GetPara("ʱ��ҩƷ�����÷ֶμӳ�", glngSys, mlngModul, 0, Array(chk�ֶμӳ����), mblnSetPara))
            
            '����������
            If int�ӳ������ = 1 Then
                intȡ�ϴ��ۼ� = 0
                int�ֶμӳ���� = 0
            ElseIf intȡ�ϴ��ۼ� = 1 Then
                int�ӳ������ = 0
                int�ֶμӳ���� = 0
            ElseIf int�ֶμӳ���� = 1 Then
                int�ӳ������ = 0
                intȡ�ϴ��ۼ� = 0
            End If
        Case 1301   'ҩƷ����������
            int�Ƿ�ѡ��ⷿ = Val(zlDatabase.GetPara("�Ƿ�ѡ��ⷿ", glngSys, mlngModul, 0, Array(fra�ⷿѡ��, chkStock, Label4), mblnSetPara))
            str���� = zlDatabase.GetPara("����", glngSys, mlngModul, "00", Array(Frame5, Cbo����, Cbo����, Label5), mblnSetPara)
            int���̴�ӡ = Val(zlDatabase.GetPara("���̴�ӡ", glngSys, mlngModul, 0, Array(chkSavePrint), mblnSetPara))
            int��˴�ӡ = Val(zlDatabase.GetPara("��˴�ӡ", glngSys, mlngModul, 0, Array(chkVerifyPrint), mblnSetPara))
            intҩƷ��λ = Val(zlDatabase.GetPara("ҩƷ��λ", glngSys, mlngModul, 0, Array(lbl�̵��, cboUnit), mblnSetPara))
            int�ɱ�����Դ = Val(zlDatabase.GetPara("ҩƷ�������ɱ��ۼ��㷽ʽ", glngSys, mlngModul, 0, Array(fra�ɱ���, opt�ɱ���Դ(0), opt�ɱ���Դ(1)), mblnSetPara))
            int��ѯ���� = Val(zlDatabase.GetPara("��ѯ����", glngSys, mlngModul, 7))
        Case 1302   'ҩƷ����������
            int�Ƿ�ѡ��ⷿ = Val(zlDatabase.GetPara("�Ƿ�ѡ��ⷿ", glngSys, mlngModul, 0, Array(fra�ⷿѡ��, chkStock, Label4), mblnSetPara))
            str���� = zlDatabase.GetPara("����", glngSys, mlngModul, "00", Array(Frame5, Cbo����, Cbo����, Label5), mblnSetPara)
            int���̴�ӡ = Val(zlDatabase.GetPara("���̴�ӡ", glngSys, mlngModul, 0, Array(chkSavePrint), mblnSetPara))
            int��˴�ӡ = Val(zlDatabase.GetPara("��˴�ӡ", glngSys, mlngModul, 0, Array(chkVerifyPrint), mblnSetPara))
            intҩƷ��λ = Val(zlDatabase.GetPara("ҩƷ��λ", glngSys, mlngModul, 0, Array(lbl�̵��, cboUnit), mblnSetPara))
            int��ѯ���� = Val(zlDatabase.GetPara("��ѯ����", glngSys, mlngModul, 7))
            
            int�ӳ������ = Val(zlDatabase.GetPara("ʱ��ҩƷ�ԼӼ������", glngSys, mlngModul, 1, Array(chk�ӳ����), mblnSetPara))
            intȡ�ϴ��ۼ� = Val(zlDatabase.GetPara("ʱ��ҩƷ���ʱȡ�ϴ��ۼ�", glngSys, mlngModul, 0, Array(chkȡ�ϴ��ۼ�), mblnSetPara))
            int�ֶμӳ���� = Val(zlDatabase.GetPara("ʱ��ҩƷ�����÷ֶμӳ�", glngSys, mlngModul, 0, Array(chk�ֶμӳ����), mblnSetPara))
            
            '����������
            If int�ӳ������ = 1 Then
                intȡ�ϴ��ۼ� = 0
                int�ֶμӳ���� = 0
            ElseIf intȡ�ϴ��ۼ� = 1 Then
                int�ӳ������ = 0
                int�ֶμӳ���� = 0
            ElseIf int�ֶμӳ���� = 1 Then
                int�ӳ������ = 0
                intȡ�ϴ��ۼ� = 0
            End If
        Case 1303   'ҩƷ����۵�������
            str���� = zlDatabase.GetPara("����", glngSys, mlngModul, "00", Array(Frame5, Cbo����, Cbo����, Label5), mblnSetPara)
            int���̴�ӡ = Val(zlDatabase.GetPara("���̴�ӡ", glngSys, mlngModul, 0, Array(chkSavePrint), mblnSetPara))
            int��˴�ӡ = Val(zlDatabase.GetPara("��˴�ӡ", glngSys, mlngModul, 0, Array(chkVerifyPrint), mblnSetPara))
            intҩƷ��λ = Val(zlDatabase.GetPara("ҩƷ��λ", glngSys, mlngModul, 0, Array(lbl�̵��, cboUnit), mblnSetPara))
            int��ѯ���� = Val(zlDatabase.GetPara("��ѯ����", glngSys, mlngModul, 7))
        Case 1304   'ҩƷ�ƿ����
            int�Ƿ�ѡ��ⷿ = Val(zlDatabase.GetPara("�Ƿ�ѡ��ⷿ", glngSys, mlngModul, 0, Array(fra�ⷿѡ��, chkStock, Label4), mblnSetPara))
            intҩƷ��λ = Val(zlDatabase.GetPara("ҩƷ��λ", glngSys, mlngModul, 0, Array(lbl�̵��, cboUnit), mblnSetPara))
            str���� = zlDatabase.GetPara("����", glngSys, mlngModul, "00", Array(Frame5, Cbo����, Cbo����, Label5), mblnSetPara)
            int���̴�ӡ = Val(zlDatabase.GetPara("���̴�ӡ", glngSys, mlngModul, 0, Array(chkSavePrint), mblnSetPara))
            int��˴�ӡ = Val(zlDatabase.GetPara("��˴�ӡ", glngSys, mlngModul, 0, Array(chkVerifyPrint), mblnSetPara))
            int���ʹ�ӡ = Val(zlDatabase.GetPara("���ʹ�ӡ", glngSys, mlngModul, 0, Array(chkSendPrint), mblnSetPara))
            
            int�ƿ����� = Val(zlDatabase.GetPara("�ƿ�����", glngSys, mlngModul, 1, Array(chk�ƿ����̿���, Label1), mblnSetPara))
            int�������� = Val(zlDatabase.GetPara("��������", glngSys, mlngModul, 0, Array(chkRequestStrike), mblnSetPara))
            int��ѯ���� = Val(zlDatabase.GetPara("��ѯ����", glngSys, mlngModul, 7))
        Case 1305   'ҩƷ���ù���
            int�Ƿ�ѡ��ⷿ = Val(zlDatabase.GetPara("�Ƿ�ѡ��ⷿ", glngSys, mlngModul, 0, Array(fra�ⷿѡ��, chkStock, Label4), mblnSetPara))
            intҩƷ��λ = Val(zlDatabase.GetPara("ҩƷ��λ", glngSys, mlngModul, 0, Array(lbl�̵��, cboUnit), mblnSetPara))
            str���� = zlDatabase.GetPara("����", glngSys, mlngModul, "00", Array(Frame5, Cbo����, Cbo����, Label5), mblnSetPara)
            int���̴�ӡ = Val(zlDatabase.GetPara("���̴�ӡ", glngSys, mlngModul, 0, Array(chkSavePrint), mblnSetPara))
            int��˴�ӡ = Val(zlDatabase.GetPara("��˴�ӡ", glngSys, mlngModul, 0, Array(chkVerifyPrint), mblnSetPara))
            int�������� = Val(zlDatabase.GetPara("������������", glngSys, mlngModul, 0, Array(chk��������), mblnSetPara))
            int��ѯ���� = Val(zlDatabase.GetPara("��ѯ����", glngSys, mlngModul, 7))
        Case 1306   'ҩƷ�����������
            int�Ƿ�ѡ��ⷿ = Val(zlDatabase.GetPara("�Ƿ�ѡ��ⷿ", glngSys, mlngModul, 0, Array(fra�ⷿѡ��, chkStock, Label4), mblnSetPara))
            str���� = zlDatabase.GetPara("����", glngSys, mlngModul, "00", Array(Frame5, Cbo����, Cbo����, Label5), mblnSetPara)
            int���̴�ӡ = Val(zlDatabase.GetPara("���̴�ӡ", glngSys, mlngModul, 0, Array(chkSavePrint), mblnSetPara))
            int��˴�ӡ = Val(zlDatabase.GetPara("��˴�ӡ", glngSys, mlngModul, 0, Array(chkVerifyPrint), mblnSetPara))
            intҩƷ��λ = Val(zlDatabase.GetPara("ҩƷ��λ", glngSys, mlngModul, 0, Array(lbl�̵��, cboUnit), mblnSetPara))
            int��ѯ���� = Val(zlDatabase.GetPara("��ѯ����", glngSys, mlngModul, 7))
        Case 1307   'ҩƷ�̵����
            str���� = zlDatabase.GetPara("����", glngSys, mlngModul, "00", Array(Frame5, Cbo����, Cbo����, Label5), mblnSetPara)
            intС��װ��λ = Val(zlDatabase.GetPara("С��װ��λ", glngSys, mlngModul, 0, Array(lbl�̵㵥, CboUnit1), mblnSetPara))
            int���̴�ӡ = Val(zlDatabase.GetPara("���̴�ӡ", glngSys, mlngModul, 0, Array(chkSavePrint), mblnSetPara))
            int��˴�ӡ = Val(zlDatabase.GetPara("��˴�ӡ", glngSys, mlngModul, 0, Array(chkVerifyPrint), mblnSetPara))
                        
            int�洢�ⷿ = Val(zlDatabase.GetPara("�洢�ⷿ", glngSys, mlngModul, 0, Array(Chk�洢�ⷿ), mblnSetPara))
            int������� = Val(zlDatabase.GetPara("����ҩƷ�������", glngSys, mlngModul, 0, Array(chk�������), mblnSetPara))
            int�̵�ͣ�� = Val(zlDatabase.GetPara("����ͣ�õ�ҩƷ", glngSys, mlngModul, 0, Array(chkStopDrug), mblnSetPara))
            mint�̵�ʱ�� = Val(zlDatabase.GetPara("�̵�ʱ�䷶Χ����", glngSys, mlngModul, 30))
            txt�̵�ʱ��.Text = mint�̵�ʱ��
            UpD�̵�ʱ��.Value = mint�̵�ʱ��
            int��ѯ���� = Val(zlDatabase.GetPara("��ѯ����", glngSys, mlngModul, 7))
        Case 1330   'ҩƷ�ƻ�����
            int�Ƿ�ѡ��ⷿ = Val(zlDatabase.GetPara("�Ƿ�ѡ��ⷿ", glngSys, mlngModul, 0, Array(fra�ⷿѡ��, chkStock, Label4), mblnSetPara))
            str���� = zlDatabase.GetPara("����", glngSys, mlngModul, "00", Array(Frame5, Cbo����, Cbo����, Label5), mblnSetPara)
            int�۸���ʾ��ʽ = Val(zlDatabase.GetPara("�۸���ʾ��ʽ", glngSys, mlngModul, 1, Array(Frame�۸���ʾ, Opt�ɱ���, Opt�ۼ�, Opt���), mblnSetPara))
            int���̴�ӡ = Val(zlDatabase.GetPara("���̴�ӡ", glngSys, mlngModul, 0, Array(chkSavePrint), mblnSetPara))
            int��˴�ӡ = Val(zlDatabase.GetPara("��˴�ӡ", glngSys, mlngModul, 0, Array(chkVerifyPrint), mblnSetPara))
            int��Ӧ��ѡ�� = Val(zlDatabase.GetPara("��Ӧ��Ĭ��ѡ��", glngSys, mlngModul, 0, Array(cbo��Ӧ��ѡ��), mblnSetPara))
            int��Ӧ�̷�Χ = Val(zlDatabase.GetPara("��Ӧ��ѡ��Χ", glngSys, mlngModul, 0, Array(cbo��Ӧ�̷�Χ), mblnSetPara))
            intҩƷ��λ = Val(zlDatabase.GetPara("ҩƷ��λ", glngSys, mlngModul, 0, Array(lbl�̵��, cboUnit), mblnSetPara))
            intPlanPoint = Val(zlDatabase.GetPara("ȫԺ�ƻ�����վ��", glngSys, mlngModul, 0, Array(chkALLPlanPoint), mblnSetPara))
            chkALLPlanPoint.Value = intPlanPoint
            int��ѯ���� = Val(zlDatabase.GetPara("��ѯ����", glngSys, mlngModul, 7))
            int����������ʱ��ʽ = Val(zlDatabase.GetPara("����������ʱ��ʽ", glngSys, mlngModul, 0, Array(fra����������ʱ��ʽ, Opt����ʵ������, Opt���ÿ�������), mblnSetPara))
        Case 1331  'ҩƷ��������
            int������ = Val(zlDatabase.GetPara("���ʱ���ٿ��", glngSys, mlngModul))
            int��ѯ���� = Val(zlDatabase.GetPara("��ѯ����", glngSys, mlngModul, 7))
        Case 1333 'ҩƷ���۹���
            str���� = zlDatabase.GetPara("����", glngSys, mlngModul, "00", Array(Frame5, Cbo����, Cbo����, Label5), mblnSetPara)
            chkʱ�۵���.Value = Val(zlDatabase.GetPara("ʱ��ҩƷ�����ε���", glngSys, 1333, 0, Array(Frame3, chkʱ�۵���), mblnSetPara))
            chk�޼���ʾ.Value = Val(zlDatabase.GetPara("�޼���ʾ", glngSys, 1333, 1, Array(Frame3, chk�޼���ʾ), mblnSetPara))
            intҩƷ��λ = Val(zlDatabase.GetPara("ҩƷ��λ", glngSys, mlngModul, 0, Array(lbl�̵��, cboUnit), mblnSetPara))
            int��ѯ���� = Val(zlDatabase.GetPara("��ѯ����", glngSys, mlngModul, 7))
    End Select
    
    txt��ѯ����.Text = int��ѯ����
    If strFunction = "ҩƷ�ƻ�����" Then
        str���ݴ�ӡ = "�ɹ��ƻ���ӡ"
    Else
        str���ݴ�ӡ = "���ݴ�ӡ"
    End If
    
    'װ��ȱʡ����
    With Cbo����
        .Clear
        .AddItem "����˳��"
        .ItemData(.NewIndex) = 0
        .AddItem "����"
        .ItemData(.NewIndex) = 1
        .AddItem "ҩƷ����"
        .ItemData(.NewIndex) = 2
        
        If InStr("ҩƷ�̵����/ҩƷ�ƿ����/ҩƷ���ù���/ҩƷ�����������", strFunction) > 0 Then
            .AddItem "�ⷿ��λ"
            .ItemData(.NewIndex) = 3
        End If
     
        .ListIndex = 0
    End With
    With Cbo����
        .Clear
        .AddItem "����"
        .ItemData(.NewIndex) = 0
        .AddItem "����"
        .ItemData(.NewIndex) = 1
        .ListIndex = 0
    End With
    
    'ȡ�����ֶμ��������Ϊȱʡ������cbo����.Enabled=False
    Cbo����.ListIndex = Mid(str����, 1, 1)
    Cbo����.ListIndex = Right(str����, 1)
    Cbo����.Enabled = Not (Cbo����.ListIndex = 0)
    
    If int���̴�ӡ = 0 Then
        chkSavePrint.Value = 0
    Else
        chkSavePrint.Value = 1
    End If
    
    If int��˴�ӡ = 0 Then
        chkVerifyPrint.Value = 0
    Else
        chkVerifyPrint.Value = 1
    End If


    If int�Ƿ�ѡ��ⷿ = 0 Then
        chkStock.Value = 0
    Else
        chkStock.Value = 1
    End If
    
    If int�洢�ⷿ = 0 Then
        Chk�洢�ⷿ.Value = 0
    Else
        Chk�洢�ⷿ.Value = 1
    End If
    If int�ɱ�����Դ = 0 Then
        opt�ɱ���Դ(0).Value = 1
        opt�ɱ���Դ(1).Value = 0
    Else
        opt�ɱ���Դ(0).Value = 0
        opt�ɱ���Դ(1).Value = 1
    End If
    
    If int�������� = 0 Then
        chk��������.Value = 0
    Else
        chk��������.Value = 1
    End If
    
    chk�������.Value = IIf(int������� = 1, 1, 0)
    chkStopDrug.Value = IIf(int�̵�ͣ�� = 1, 1, 0)
    
    fra�ƿ����̿���.Visible = False
    fra�ϴβɹ���Ϣ.Visible = False
    fra�ɹ��ƻ�.Visible = False
    chkALLPlanPoint.Visible = False
    fra����������ʱ��ʽ.Visible = False
    
    If mstrFunction = "ҩƷ�̵����" Then
        If glngSys \ 100 = 8 Then
            With CboUnit1
                .AddItem "�ɹ���λ"
                .AddItem "�ۼ۵�λ"
            End With
        Else
            With CboUnit1
                .AddItem "�ʹ��װ��ͬ"
                .AddItem "ҩ�ⵥλ"
                .AddItem "���ﵥλ"
                .AddItem "סԺ��λ"
                .AddItem "�ۼ۵�λ"
            End With
        End If
        CboUnit1.ListIndex = intС��װ��λ
        lblUnitComment.Caption = "    ��ѡ���̵�ʱ�Ĵ�С��װ���̵㵥���̵��༭ʱ����ѡ��װ�����̵㡣"
    Else
        CboUnit1.Visible = False
        lbl�̵��.Visible = False
        lbl�̵㵥.Visible = False
        cboUnit.Left = lbl�̵��.Left
        cboUnit.Width = Frame2.Width - cboUnit.Left - 250
        lblUnitComment.Top = lbl�̵㵥.Top
    End If
    
    With cboUnit
        .Clear
        If glngSys \ 100 = 8 Then
            .AddItem "ȱʡ����ǰ�ⷿ��Ӧ�ĵ�λ��"
            .AddItem "�ɹ���λ"
            .AddItem "�ۼ۵�λ"
        Else
            If mlngModul <> 1333 Then   '���۲���Ҫ�ⷿ
                .AddItem "ȱʡ����ǰ�ⷿ��Ӧ�ĵ�λ��"
            End If
            .AddItem "ҩ�ⵥλ"
            .AddItem "���ﵥλ"
            .AddItem "סԺ��λ"
            .AddItem "�ۼ۵�λ"
        End If
        .ListIndex = intҩƷ��λ
    End With
    
    If strFunction = "ҩƷ�⹺������" Then
        chkFixPrice.Visible = True
        chk�⹺NO.Visible = True
        Chk�����޸�������.Visible = True
        chk�˻���Ʊ���.Visible = True
        chk�б�ҩƷ.Visible = True
        chkFixPrice.Value = int���۲ɹ�
        chk�⹺NO.Value = int�޸��⹺���ݺ�
        Chk�����޸�������.Value = int�޸�������
        chk�˻���Ʊ���.Value = int�˻���Ʊ���
        chk�б�ҩƷ.Value = int�б�ҩƷ
        
        fra�ϴβɹ���Ϣ.Visible = True
        If intȡ�ϴβɹ��۷�ʽ = 1 Then
            optȡ�ɱ��۷�ʽ(1).Value = True
        Else
            optȡ�ɱ��۷�ʽ(0).Value = True
        End If
        
        chk�ӳ����.Visible = True
        chkȡ�ϴ��ۼ�.Visible = True
        chk�ֶμӳ����.Visible = True
        
        chk�ӳ����.Value = int�ӳ������
        chkȡ�ϴ��ۼ�.Value = intȡ�ϴ��ۼ�
        chk�ֶμӳ����.Value = int�ֶμӳ����
        
        lbl��ѯ����.Move fra�ϴβɹ���Ϣ.Left, fra�ϴβɹ���Ϣ.Top + fra�ϴβɹ���Ϣ.Height + 200
        txt��ѯ����.Move lbl��ѯ����.Left + lbl��ѯ����.Width + 100, lbl��ѯ����.Top - 50
        lbl����.Move txt��ѯ����.Left + txt��ѯ����.Width + 50, lbl��ѯ����.Top
    End If
    
    If strFunction = "ҩƷ����������" Then
        chk�ӳ����.Visible = True
        chkȡ�ϴ��ۼ�.Visible = True
        chk�ֶμӳ����.Visible = True
        
        chk�ӳ����.Value = int�ӳ������
        chkȡ�ϴ��ۼ�.Value = intȡ�ϴ��ۼ�
        chk�ֶμӳ����.Value = int�ֶμӳ����
    End If
    
'    Frame2.Enabled = (strFunction = "ҩƷ�������" Or strFunction = "ҩƷ�ƿ����" Or strFunction = "ҩƷ���ù���")
    If strFunction <> "ҩƷ��������" Then
        Frame2.Enabled = True
    End If
    
    chkStopDrug.Visible = False
    If strFunction = "ҩƷ�̵����" Then
        Frame2.Enabled = True
        cboUnit.Enabled = False
        Chk�洢�ⷿ.Visible = True
        chk�������.Visible = True
        chkStopDrug.Visible = True
    End If
    
    fra�ⷿѡ��.Enabled = (InStr(1, "ҩƷ�̵����,����۵�������", strFunction) = 0)
    If fra�ⷿѡ��.Enabled = False Then
        chkStock.Enabled = False
    End If
    
    If strFunction = "ҩƷ�ƿ����" Then
        chk�ƿ����̿���.Value = int�ƿ�����
        chkRequestStrike.Value = int��������
        fra�ƿ����̿���.Visible = True
        
        chkSendPrint.Value = IIf(int���ʹ�ӡ = 1, 1, 0)
        chkSendPrint.Visible = (chk�ƿ����̿���.Value = 1)
        
        lbl��ѯ����.Move fra�ƿ����̿���.Left, fra�ƿ����̿���.Top + fra�ƿ����̿���.Height + 150
        txt��ѯ����.Move lbl��ѯ����.Left + lbl��ѯ����.Width + 100, lbl��ѯ����.Top - 50
        lbl����.Move txt��ѯ����.Left + txt��ѯ����.Width + 50, lbl��ѯ����.Top
    End If
    
    If strFunction = "ҩƷ�ƻ�����" Then
        cbo��Ӧ��ѡ��.Clear
        cbo��Ӧ��ѡ��.AddItem "1-ȡ�ϴ���⹩Ӧ��"
        cbo��Ӧ��ѡ��.AddItem "2-ȡ��ͬ��λ"
        cbo��Ӧ��ѡ��.ListIndex = IIf(int��Ӧ��ѡ�� < 0 Or int��Ӧ��ѡ�� > 1, 0, int��Ӧ��ѡ��)
        
        cbo��Ӧ�̷�Χ.Clear
        cbo��Ӧ�̷�Χ.AddItem "1-���й�Ӧ��"
        cbo��Ӧ�̷�Χ.AddItem "2-�б굥λ"
        cbo��Ӧ�̷�Χ.ListIndex = IIf(int��Ӧ�̷�Χ < 0 Or int��Ӧ�̷�Χ > 1, 0, int��Ӧ�̷�Χ)
        
        fra�ɹ��ƻ�.Visible = True
        chkALLPlanPoint.Visible = True
        fra����������ʱ��ʽ.Visible = True
        chkALLPlanPoint.Top = fra�ɹ��ƻ�.Top + fra�ɹ��ƻ�.Height + 113
        chkALLPlanPoint.Left = fra�ɹ��ƻ�.Left
        
        fra����������ʱ��ʽ.Top = Frame�۸���ʾ.Top + Frame�۸���ʾ.Height + 150
        fra����������ʱ��ʽ.Left = Frame�۸���ʾ.Left
        
        lbl��ѯ����.Move chkALLPlanPoint.Left + chkALLPlanPoint.Width + 150, fra�ɹ��ƻ�.Top + fra�ɹ��ƻ�.Height + 150
        txt��ѯ����.Move lbl��ѯ����.Left + lbl��ѯ����.Width + 100, lbl��ѯ����.Top - 50
        lbl����.Move txt��ѯ����.Left + txt��ѯ����.Width + 50, lbl��ѯ����.Top
    End If
    
    If Frame3.Enabled = True Then
        If strFunction = "ҩƷ�ƻ�����" Then
            Frame�۸���ʾ.Visible = True
            If int�۸���ʾ��ʽ = 0 Then
                Opt�ɱ���.Value = True
            ElseIf int�۸���ʾ��ʽ = 1 Then
                Opt�ۼ�.Value = True
            Else
                Opt���.Value = True
            End If
            
            If int����������ʱ��ʽ = 0 Then
                Opt����ʵ������.Value = True
            Else
                Opt���ÿ�������.Value = True
            End If
        End If
    End If
    
    If mlngModul = 1331 Then    'ҩƷ��������
        chk������.Value = int������
    End If
    
    '�������
    If strFunction <> "ҩƷ�ƿ����" And strFunction <> "ҩƷ�⹺������" And strFunction <> "ҩƷ�ƻ�����" Then
        fra�ƿ����̿���.Visible = False
        
        tabMain.Height = tabMain.Height - fra�ƿ����̿���.Height
        
        cmdHelp.Top = cmdHelp.Top - fra�ƿ����̿���.Height
        cmdOK.Top = cmdHelp.Top
        cmdCancel.Top = cmdHelp.Top
        
        Me.Height = Me.Height - fra�ƿ����̿���.Height
    End If
    
    If strFunction = "ҩƷ��������" Then
        fra�ⷿѡ��.Visible = False
        Frame2.Visible = False
        Frame5.Visible = False
        Frame3.Visible = False
        fra�ƿ����̿���.Visible = False
        fra�ϴβɹ���Ϣ.Visible = False
        
        fraͬ��������.Visible = True
        fraͬ��������.Top = 580
        
        lbl��ѯ����.Move fraͬ��������.Left, fraͬ��������.Top + fraͬ��������.Height + 200
        txt��ѯ����.Move lbl��ѯ����.Left + lbl��ѯ����.Width + 100, lbl��ѯ����.Top - 50
        lbl����.Move txt��ѯ����.Left + txt��ѯ����.Width + 50, lbl��ѯ����.Top
        
        tabMain.Height = tabMain.Height - fraͬ��������.Height
        
        cmdHelp.Top = cmdHelp.Top - fraͬ��������.Height
        cmdOK.Top = cmdHelp.Top
        cmdCancel.Top = cmdHelp.Top
        
        Me.Height = Me.Height - fraͬ��������.Height
    End If
    If strFunction = "ҩƷ�̵����" Then
        tabMain.Height = fra�̵�ʱ�䷶Χ.Height + fra�̵�ʱ�䷶Χ.Top + 200
        Me.Height = fra�̵�ʱ�䷶Χ.Top + fra�̵�ʱ�䷶Χ.Height + 1300
        fra�̵�ʱ�䷶Χ.Visible = True
        cmdHelp.Top = tabMain.Height + 250
        cmdOK.Top = cmdHelp.Top
        cmdCancel.Top = cmdHelp.Top
        
        lbl��ѯ����.Move Frame3.Left, Frame3.Top + Frame3.Height + 500
        txt��ѯ����.Move lbl��ѯ����.Left + lbl��ѯ����.Width + 100, lbl��ѯ����.Top - 50
        lbl����.Move txt��ѯ����.Left + txt��ѯ����.Width + 50, lbl��ѯ����.Top
    End If
    
    '����У��ҳ��
    tabMain.TabVisible(1) = strFunction = "ҩƷ�⹺������"
    If tabMain.TabVisible(1) = True Then
        With vsfCheck
            .MergeCol(0) = True
            .MergeCells = flexMergeRestrictColumns
        End With
        fraCheck.Top = tabMain.Height - fraCheck.Height - 100
        vsfCheck.Height = fraCheck.Top - vsfCheck.Top - 100
        
        Load����У��
    End If
    
    If strFunction = "ҩƷ����������" Then
        fra�ɱ���.Visible = True
        tabMain.Height = tabMain.Height + fra�ɱ���.Height + 100
        frm��������.Height = tabMain.Height + cmdOK.Height + 800
        cmdHelp.Top = frm��������.Height - 900
        cmdOK.Top = cmdHelp.Top
        cmdCancel.Top = cmdHelp.Top
        
        lbl��ѯ����.Move fra�ɱ���.Left, fra�ɱ���.Top + fra�ɱ���.Height + 200
        txt��ѯ����.Move lbl��ѯ����.Left + lbl��ѯ����.Width + 100, lbl��ѯ����.Top - 50
        lbl����.Move txt��ѯ����.Left + txt��ѯ����.Width + 50, lbl��ѯ����.Top
    Else
        fra�ɱ���.Visible = False
    End If
    
    If strFunction = "ҩƷ���۹���" Then
        fra�ⷿѡ��.Visible = False
        Frame5.Visible = False
        Frame2.Left = fra�ⷿѡ��.Left
        Frame2.Top = fra�ⷿѡ��.Top
        
        Frame3.Top = Frame2.Top
        
        Frame2.Height = Frame3.Height
        fraͬ��������.Visible = False
        chkSavePrint.Visible = False
        chkVerifyPrint.Visible = False
        Frame5.Enabled = False
        chkʱ�۵���.Visible = True
        chk�޼���ʾ.Visible = True
        chkʱ�۵���.Move chkSavePrint.Left, chkSavePrint.Top
        chk�޼���ʾ.Move chkʱ�۵���.Left, chkʱ�۵���.Top + chkʱ�۵���.Height + 100
        
        tabMain.Height = Frame2.Height + cmdOK.Height + 500
        cmdHelp.Top = tabMain.Height + tabMain.Top + 100
        cmdOK.Top = cmdHelp.Top
        cmdCancel.Top = cmdHelp.Top
        
        Me.Height = tabMain.Height + tabMain.Top + 1000
        
    End If
    If strFunction = "ҩƷ���ù���" Then
        chk��������.Visible = True
        chk��������.Left = 165
        chk��������.Top = chkSavePrint.Top + chkSavePrint.Height + 50
        lbl��ѯ����.Move Frame5.Left, Frame5.Top + Frame5.Height + 200
        txt��ѯ����.Move lbl��ѯ����.Left + lbl��ѯ����.Width + 100, lbl��ѯ����.Top - 50
        lbl����.Move txt��ѯ����.Left + txt��ѯ����.Width + 50, lbl��ѯ����.Top
    End If
    
    If mlngModul = 1302 Or mlngModul = 1303 Or mlngModul = 1306 Then
        '1302 :�������;1303:�����; 1306����������
        lbl��ѯ����.Move Frame5.Left, Frame5.Top + Frame5.Height + 200
        txt��ѯ����.Move lbl��ѯ����.Left + lbl��ѯ����.Width + 100, lbl��ѯ����.Top - 50
        lbl����.Move txt��ѯ����.Left + txt��ѯ����.Width + 50, lbl��ѯ����.Top
    End If
    
    frm��������.Show vbModal, frmParent
End Sub

Private Sub Load����У��()
    Dim i As Integer
    Dim n As Integer
    Dim strCheck As String
    Dim intCheckType As Integer
    Dim arrColumn
    
    On Error Resume Next
    
    '����У����Ŀ�ͷ�ʽ�ı����ʽ��У�鷽ʽ|���1,��Ŀ1,�Ƿ�У��;���1,��Ŀ2,�Ƿ�У��;���2,��Ŀ1,�Ƿ�У��;���2,��Ŀ2....
    strCheck = zlDatabase.GetPara("����У��", glngSys, mlngModul, "", Array(vsfCheck, fraCheck), mblnSetPara)
    
    If strCheck <> "" Then
        If InStr(1, strCheck, "|") > 0 Then
            'У�鷽ʽ��0-����飻1�����ѣ�2����ֹ
            intCheckType = Val(Mid(strCheck, 1, InStr(1, strCheck, "|") - 1))
            If intCheckType = 2 Then
                optCheck(0).Value = True
            ElseIf intCheckType = 1 Then
                optCheck(1).Value = True
            End If
            
            strCheck = Mid(strCheck, InStr(1, strCheck, "|") + 1)
             
            If strCheck <> "" Then
                strCheck = strCheck & ";"
                arrColumn = Split(strCheck, ";")
                For n = 0 To UBound(arrColumn)
                    If arrColumn(n) <> "" Then
                        With vsfCheck
                            For i = 1 To .rows - 1
                                If Split(arrColumn(n), ",")(0) = .TextMatrix(i, .ColIndex("���")) And Split(arrColumn(n), ",")(1) = .TextMatrix(i, .ColIndex("У����Ŀ")) Then
                                    If Val(Split(arrColumn(n), ",")(2)) = 1 Then
                                        .TextMatrix(i, .ColIndex("У��")) = "��"
                                    End If
                                End If
                            Next
                        End With
                    End If
                Next
            End If
        End If
    End If
End Sub

Private Sub cmd��ӡ����_Click()
    Dim strBill As String
    Select Case mstrFunction
    Case "ҩƷ�⹺������"
        strBill = "ZL1_BILL_1300"
    Case "ҩƷ����������"
        strBill = "ZL1_BILL_1302"
    Case "ҩƷ����������"
        strBill = "ZL1_BILL_1301"
    Case "����۵�������"
        strBill = "ZL1_BILL_1303"
    Case "ҩƷ�ƿ����"
        strBill = "ZL1_BILL_1304"
    Case "ҩƷ���ù���"
        strBill = "ZL1_BILL_1305"
    Case "ҩƷ�����������"
        strBill = "ZL1_BILL_1306"
    Case "ҩƷ�̵����"
        strBill = "ZL1_BILL_1307"
    Case "ҩƷ�ƻ�����"
        strBill = "zl1_bill_1330"
    Case "ҩƷ���۹���"
        strBill = "ZL1_BILL_1333"
    End Select
    Call ReportPrintSet(gcnOracle, glngSys, strBill, Me)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Me.cmd��ӡ����.Caption = "Ʊ�ݡ�" & Mid(mstrFunction, 1, Len(mstrFunction) - 2) & "������ӡ����"
End Sub

Private Sub tabMain_Click(PreviousTab As Integer)
    If PreviousTab = 0 Then
        If vsfCheck.Enabled = True Then vsfCheck.SetFocus
    End If
End Sub

Private Sub txt��ѯ����_KeyPress(KeyAscii As Integer)
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Then Exit Sub
    KeyAscii = 0
End Sub

Private Sub txt��ѯ����_Validate(Cancel As Boolean)
    If Val(txt��ѯ����.Text) = 0 Then
        MsgBox "��ѯʱ��������0�����������룡", vbInformation, gstrSysName
        Cancel = False
        txt��ѯ����.SetFocus
        zlControl.TxtSelAll txt��ѯ����
    End If
End Sub

Private Sub txt�̵�ʱ��_Change()
    txt�̵�ʱ��.Text = IIf(txt�̵�ʱ��.Text = "", "0", txt�̵�ʱ��.Text) '��ֹ�ı�Ϊ��
    UpD�̵�ʱ��.Value = IIf(Val(txt�̵�ʱ��.Text) > 90, Val(Mid(txt�̵�ʱ��.Text, 1, Len(txt�̵�ʱ��.Text) - 1)), Val(txt�̵�ʱ��.Text))
End Sub

Private Sub txt�̵�ʱ��_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyBack Then
        If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txt�̵�ʱ��_Validate(Cancel As Boolean)
    If Val(txt�̵�ʱ��.Text) > 90 Then
        MsgBox "�̵�ʱ�䷶Χ���ܴ���3���£�", vbInformation, gstrSysName
        Cancel = True
    End If
End Sub

Private Sub UpD�̵�ʱ��_Change()
    txt�̵�ʱ��.Text = UpD�̵�ʱ��.Value
    txt�̵�ʱ��.SelStart = Len(txt�̵�ʱ��.Text) '��λ���ı�ĩβ
End Sub

Private Sub vsfCheck_DblClick()
    With vsfCheck
        If .Row = 0 Then Exit Sub
        If .Col <> .ColIndex("У��") Then Exit Sub
        If .MouseRow <> .Row Or .MouseCol <> .Col Then Exit Sub
        
        If .TextMatrix(.Row, .Col) = "��" Then
            .TextMatrix(.Row, .Col) = ""
        Else
            .TextMatrix(.Row, .Col) = "��"
        End If
    End With
End Sub


