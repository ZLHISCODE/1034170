VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ϵͳѡ��"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6300
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Caption         =   "1"
      Height          =   6375
      Index           =   3
      Left            =   165
      TabIndex        =   35
      Top             =   540
      Width           =   5805
      Begin VB.Frame fraSplit 
         Height          =   120
         Index           =   3
         Left            =   1800
         TabIndex        =   40
         Top             =   3720
         Width           =   3615
      End
      Begin VB.TextBox txtRemote 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         Left            =   2640
         MaxLength       =   5
         TabIndex        =   76
         ToolTipText     =   "���ڽ��ܷ�������Ϣ�����磺Զ�����ӡ�������Ϣ�����������Ϣ�ȡ�������ʱ���������������������������ȣ��޷�Զ�����ӡ�"
         Top             =   6100
         Width           =   675
      End
      Begin VB.TextBox txtAutoLock 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   2160
         MaxLength       =   3
         TabIndex        =   75
         Text            =   "5"
         ToolTipText     =   "�������������0-999"
         Top             =   5768
         Width           =   300
      End
      Begin VB.Frame fraSplit 
         Height          =   120
         Index           =   8
         Left            =   1215
         TabIndex        =   73
         Top             =   5190
         Width           =   4200
      End
      Begin VB.CheckBox chkLanJoin 
         Caption         =   "����Ͽ����ָ����Զ��������ӷ�����"
         Height          =   210
         Left            =   1500
         TabIndex        =   71
         Top             =   5460
         Width           =   3465
      End
      Begin VB.Frame fraSplit 
         Height          =   120
         Index           =   5
         Left            =   1560
         TabIndex        =   58
         Top             =   3030
         Width           =   3825
      End
      Begin VB.CheckBox chkAutoHide 
         Caption         =   "�������������ṩ�Զ����ع���(&I)"
         Height          =   195
         Left            =   1500
         TabIndex        =   18
         Top             =   3390
         Width           =   3090
      End
      Begin VB.PictureBox pic3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   480
         Index           =   5
         Left            =   600
         Picture         =   "frmOptions.frx":000C
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   3255
         Width           =   480
      End
      Begin VB.PictureBox picƥ�䷽ʽ 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   1455
         ScaleHeight     =   315
         ScaleWidth      =   3495
         TabIndex        =   55
         Top             =   870
         Width           =   3495
         Begin VB.OptionButton opt 
            Caption         =   "˫��ƥ��(&D)"
            Height          =   210
            Index           =   0
            Left            =   45
            TabIndex        =   12
            Top             =   60
            Value           =   -1  'True
            Width           =   1320
         End
         Begin VB.OptionButton opt 
            Caption         =   "����ƥ��(&L)"
            Height          =   210
            Index           =   1
            Left            =   1815
            TabIndex        =   13
            Top             =   60
            Width           =   1320
         End
      End
      Begin VB.PictureBox pic3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   480
         Index           =   4
         Left            =   585
         Picture         =   "frmOptions.frx":08D6
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   4665
         Width           =   480
      End
      Begin VB.PictureBox pic3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   480
         Index           =   3
         Left            =   600
         Picture         =   "frmOptions.frx":32C8
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   3900
         Width           =   480
      End
      Begin VB.PictureBox pic3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   480
         Index           =   2
         Left            =   615
         Picture         =   "frmOptions.frx":4C4A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   2415
         Width           =   480
      End
      Begin VB.PictureBox pic3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   480
         Index           =   1
         Left            =   645
         Picture         =   "frmOptions.frx":65CC
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   1575
         Width           =   480
      End
      Begin VB.PictureBox pic3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   480
         Index           =   0
         Left            =   645
         Picture         =   "frmOptions.frx":7F4E
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   825
         Width           =   480
      End
      Begin VB.PictureBox pic2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   480
         Left            =   255
         Picture         =   "frmOptions.frx":98D0
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   0
         Width           =   480
      End
      Begin VB.TextBox txtTime 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   1695
         MaxLength       =   4
         TabIndex        =   21
         Text            =   "60"
         ToolTipText     =   "�����������Ϊ0��յ�ʱ�򣬱�ʾ����顣����������÷�Χ��10-300��"
         Top             =   4830
         Width           =   540
      End
      Begin VB.Frame fraSplit 
         Height          =   120
         Index           =   4
         Left            =   1935
         TabIndex        =   41
         Top             =   4455
         Width           =   3480
      End
      Begin VB.CheckBox chkShutDown 
         Caption         =   "�˳�����ʱ�Զ��ر� Windows (&S)"
         Height          =   210
         Left            =   1500
         TabIndex        =   20
         Top             =   4200
         Width           =   3045
      End
      Begin VB.CheckBox chkAutoStart 
         Caption         =   "�� Windows ����ʱ�Զ�����(&A)"
         Height          =   210
         Left            =   1500
         TabIndex        =   19
         Top             =   3990
         Width           =   2865
      End
      Begin VB.Frame fraSplit 
         Height          =   120
         Index           =   2
         Left            =   1560
         TabIndex        =   39
         Top             =   2115
         Width           =   3825
      End
      Begin VB.Frame fraSplit 
         Height          =   120
         Index           =   1
         Left            =   1365
         TabIndex        =   38
         Top             =   1335
         Width           =   4020
      End
      Begin VB.ComboBox cmbIME 
         Height          =   300
         Left            =   1530
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1710
         Width           =   3135
      End
      Begin VB.Frame fraSplit 
         Height          =   120
         Index           =   0
         Left            =   1950
         TabIndex        =   37
         Top             =   570
         Width           =   3435
      End
      Begin VB.PictureBox pic���� 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   870
         Left            =   1485
         ScaleHeight     =   870
         ScaleWidth      =   3345
         TabIndex        =   36
         Top             =   2220
         Width           =   3345
         Begin VB.CheckBox chkIMETurn 
            Caption         =   "�����ڴ��ڽ���Ĺ������л����뷽ʽ "
            Height          =   255
            Left            =   15
            TabIndex        =   17
            Top             =   560
            Value           =   1  'Checked
            Width           =   3360
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "��ʣ�ȡÿ�ֵ�����ĸ���ɼ���(&W)"
            Height          =   210
            Index           =   1
            Left            =   15
            TabIndex        =   16
            Top             =   330
            Width           =   3150
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "ƴ����ȡÿ�ֵ�����ĸ���ɼ���(&P)"
            Height          =   210
            Index           =   0
            Left            =   15
            TabIndex        =   15
            Top             =   90
            Value           =   -1  'True
            Width           =   3150
         End
      End
      Begin VB.CheckBox chkAutoLock 
         Caption         =   "�ȴ�     �����޲������Զ�����ϵͳ"
         Height          =   210
         Left            =   1500
         TabIndex        =   74
         Top             =   5760
         Width           =   3345
      End
      Begin VB.Label lblRemote 
         AutoSize        =   -1  'True
         Caption         =   "���ؼ����˿�________��ֵΪ-1ʱ������"
         Height          =   180
         Left            =   1500
         TabIndex        =   77
         Top             =   6120
         Width           =   3240
      End
      Begin VB.Line Line1 
         X1              =   2160
         X2              =   2600
         Y1              =   5985
         Y2              =   5985
      End
      Begin VB.Label lblSplit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Index           =   8
         Left            =   330
         TabIndex        =   72
         Top             =   5175
         Width           =   720
      End
      Begin VB.Label lblSplit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������������"
         Height          =   180
         Index           =   5
         Left            =   300
         TabIndex        =   59
         Top             =   3045
         Width           =   1080
      End
      Begin VB.Line lineTime 
         X1              =   1680
         X2              =   2295
         Y1              =   5025
         Y2              =   5025
      End
      Begin VB.Label lblSplit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ϣ֪ͨ�������"
         Height          =   180
         Index           =   4
         Left            =   300
         TabIndex        =   48
         Top             =   4455
         Width           =   1440
      End
      Begin VB.Label lblSplit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Windows�Ľ��"
         Height          =   180
         Index           =   3
         Left            =   300
         TabIndex        =   47
         Top             =   3720
         Width           =   1350
      End
      Begin VB.Label lblSplit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���봦����ʽ"
         Height          =   180
         Index           =   2
         Left            =   300
         TabIndex        =   46
         Top             =   2130
         Width           =   1080
      End
      Begin VB.Label lblSplit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������뷨"
         Height          =   180
         Index           =   1
         Left            =   300
         TabIndex        =   45
         Top             =   1350
         Width           =   900
      End
      Begin VB.Label lblSplit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ŀ����ƥ�䷽ʽ"
         Height          =   180
         Index           =   0
         Left            =   300
         TabIndex        =   44
         Top             =   570
         Width           =   1440
      End
      Begin VB.Label Label1 
         Caption         =   "    �û����Ը���������ϰ����ѡ�������ƥ�䷽ʽ���������͡����뷨�ȣ�����߹���Ч��"
         Height          =   480
         Left            =   870
         TabIndex        =   43
         Top             =   120
         Width           =   4575
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image2 
         Height          =   510
         Left            =   240
         Picture         =   "frmOptions.frx":A59A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   465
      End
      Begin VB.Label lblTime 
         Caption         =   "ÿ       ������Ϣ֪ͨ"
         Height          =   255
         Left            =   1485
         TabIndex        =   42
         Top             =   4830
         Width           =   2145
      End
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   6015
      Index           =   4
      Left            =   240
      TabIndex        =   60
      Top             =   540
      Width           =   5685
      Begin VB.PictureBox pic3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   480
         Index           =   8
         Left            =   390
         Picture         =   "frmOptions.frx":ABAF
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   240
         Width           =   480
      End
      Begin VB.ComboBox cbo����ҩƷ��ʾ 
         Height          =   300
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   3120
         Width           =   3015
      End
      Begin VB.Frame fraSplit 
         Height          =   120
         Index           =   7
         Left            =   1200
         TabIndex        =   67
         Top             =   2400
         Width           =   4380
      End
      Begin VB.PictureBox pic3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   480
         Index           =   7
         Left            =   360
         Picture         =   "frmOptions.frx":B879
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   3000
         Width           =   480
      End
      Begin VB.ComboBox cboҩƷ������ʾ 
         Height          =   300
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   1800
         Width           =   3015
      End
      Begin VB.Frame fraSplit 
         Height          =   120
         Index           =   6
         Left            =   1200
         TabIndex        =   62
         Top             =   960
         Width           =   4260
      End
      Begin VB.PictureBox pic3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   480
         Index           =   6
         Left            =   390
         Picture         =   "frmOptions.frx":C543
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   1680
         Width           =   480
      End
      Begin VB.Label lblSplit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ҩƷ��ʾ��ͨ��������뷽ʽ����ѡ����ʱҩƷ���Ƶ���ʾ��"
         Height          =   180
         Index           =   7
         Left            =   120
         TabIndex        =   69
         Top             =   2640
         Width           =   5220
      End
      Begin VB.Label lblMedi 
         Caption         =   "    �û����Ը���������ϰ����ѡ��ҩƷ���Ƶ���ʾ��ʽ��֧����ʾͨ��������Ʒ����"
         Height          =   585
         Left            =   1080
         TabIndex        =   65
         Top             =   240
         Width           =   4455
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblSplit 
         BackStyle       =   0  'Transparent
         Caption         =   "ҩƷ������ʾ�������浥����ϸ������������桢ֱ�ӽ����ҩƷѡ����ʱ��ҩƷ������ʾ��"
         Height          =   420
         Index           =   6
         Left            =   120
         TabIndex        =   64
         Top             =   1200
         Width           =   5220
      End
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   6015
      Index           =   2
      Left            =   180
      TabIndex        =   28
      Top             =   540
      Width           =   5805
      Begin VB.CommandButton cmdFavorite 
         Height          =   345
         Index           =   3
         Left            =   5250
         Picture         =   "frmOptions.frx":D20D
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "����ƶ�"
         Top             =   5520
         Width           =   345
      End
      Begin VB.CommandButton cmdFavorite 
         Height          =   345
         Index           =   2
         Left            =   5250
         Picture         =   "frmOptions.frx":D35A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "��ǰ�ƶ�"
         Top             =   5040
         Width           =   345
      End
      Begin VB.CommandButton cmdFavorite 
         Height          =   345
         Index           =   1
         Left            =   5250
         Picture         =   "frmOptions.frx":D4A7
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "ɾ������"
         Top             =   3210
         Width           =   345
      End
      Begin VB.CommandButton cmdFavorite 
         Height          =   345
         Index           =   0
         Left            =   5250
         Picture         =   "frmOptions.frx":D54D
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "��������"
         Top             =   2730
         Width           =   345
      End
      Begin MSComctlLib.ListView lvwFavorite 
         Height          =   3165
         Left            =   3000
         TabIndex        =   7
         Top             =   2700
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   5583
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ils16"
         SmallIcons      =   "ils16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "����"
            Object.Width           =   3175
         EndProperty
      End
      Begin VB.Frame Frame1 
         Caption         =   "���ɿ�ݷ�ʽ"
         Height          =   1335
         Left            =   3030
         TabIndex        =   29
         Top             =   750
         Width           =   2625
         Begin VB.CommandButton cmdStartup 
            Caption         =   "�������˵�(&S)"
            Height          =   350
            Left            =   630
            TabIndex        =   6
            Top             =   840
            Width           =   1725
         End
         Begin VB.CommandButton cmdDesktop 
            Caption         =   "������(&D)"
            Height          =   350
            Left            =   630
            TabIndex        =   5
            Top             =   360
            Width           =   1725
         End
      End
      Begin VB.ComboBox cboGroup 
         Height          =   300
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   750
         Width           =   1620
      End
      Begin MSComctlLib.ImageList ils16 
         Left            =   2160
         Top             =   2910
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin MSComctlLib.TreeView tvwMain 
         Height          =   4755
         Left            =   300
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1140
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   8387
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "ils16"
         Appearance      =   1
      End
      Begin VB.Label lblFavorite 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ó����б�(&R)"
         Height          =   180
         Left            =   3030
         TabIndex        =   33
         Top             =   2430
         Width           =   1350
      End
      Begin VB.Label lblNote 
         Caption         =   $"frmOptions.frx":D5FA
         Height          =   570
         Left            =   900
         TabIndex        =   31
         Top             =   30
         Width           =   4740
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "�˵���ϵ"
         Height          =   180
         Left            =   360
         TabIndex        =   30
         Top             =   810
         Width           =   720
      End
      Begin VB.Image imgNote 
         Height          =   510
         Left            =   240
         Picture         =   "frmOptions.frx":D682
         Stretch         =   -1  'True
         Top             =   0
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   390
      TabIndex        =   24
      Top             =   7110
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4860
      TabIndex        =   23
      Top             =   7110
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   3540
      TabIndex        =   22
      Top             =   7110
      Width           =   1100
   End
   Begin MSComctlLib.TabStrip tabMain 
      Height          =   6975
      Left            =   90
      TabIndex        =   25
      Top             =   90
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   12303
      TabWidthStyle   =   2
      TabMinWidth     =   989
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�������"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�˵�ѡ��"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ʹ��ϰ��"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ҩƷ����"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   6015
      Index           =   1
      Left            =   180
      TabIndex        =   27
      Top             =   540
      Width           =   5805
      Begin VB.PictureBox pic������� 
         BorderStyle     =   0  'None
         Height          =   420
         Left            =   615
         ScaleHeight     =   420
         ScaleWidth      =   4410
         TabIndex        =   56
         Top             =   960
         Width           =   4410
         Begin VB.OptionButton optStyle 
            Caption         =   "Windows���"
            Height          =   195
            Index           =   1
            Left            =   1455
            TabIndex        =   1
            Tag             =   "zlwin"
            Top             =   60
            Width           =   1428
         End
         Begin VB.OptionButton optStyle 
            Caption         =   "��ͳ���"
            Height          =   195
            Index           =   0
            Left            =   15
            TabIndex        =   0
            Tag             =   "zlBrw"
            Top             =   60
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton optStyle 
            Caption         =   "MDI���"
            Height          =   195
            Index           =   2
            Left            =   3270
            TabIndex        =   2
            Tag             =   "zlmdi"
            Top             =   60
            Width           =   945
         End
      End
      Begin VB.PictureBox picPreview 
         Height          =   3375
         Left            =   630
         ScaleHeight     =   221
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   269
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1710
         Width           =   4095
      End
      Begin VB.PictureBox picBorder 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3765
         Left            =   420
         ScaleHeight     =   251
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   301
         TabIndex        =   34
         Top             =   1500
         Width           =   4515
      End
      Begin VB.Image Image1 
         Height          =   510
         Left            =   240
         Picture         =   "frmOptions.frx":D88C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   465
      End
      Begin VB.Label Label3 
         Caption         =   "    ���ո��˵���Ȥ����ѡ���Լ�ϲ���ĵ������ʹ���������졢���������ɡ�"
         Height          =   480
         Left            =   870
         TabIndex        =   26
         Top             =   120
         Width           =   4140
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum constOpt
    opt_˫��ƥ�� = 0
    opt_����ƥ�� = 1
End Enum

Private Declare Function OSfCreateShellLink Lib "vb6stkit.dll" Alias "fCreateShellLink" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArguments As String, ByVal fPrivate As Long, ByVal sParent As String) As Long
Private Declare Function OSfRemoveShellLink Lib "vb6stkit.dll" Alias "fRemoveShellLink" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String) As Long
Dim mintIndex As Integer

Private Sub chkAutoLock_Click()
        txtAutoLock.Enabled = chkAutoLock.Value = 1
End Sub



Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDesktop_Click()
    Dim strPath As String
    Dim StrName As String
    
    '������ݷ�ʽ
    strPath = GetSetting("ZLSOFT", "����ȫ��", "����·��", "C:\AppSoft\ZLHIS+.exe")
    StrName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "")
    
'    '�������
    If Not OSfCreateShellLink("..\DeskTop", StrName & "����̨(" & cboGroup.Text & ")", strPath, cboGroup.Text, True, "$(Start Menu)") Then
        If Not OSfCreateShellLink("..\����", StrName & "����̨(" & cboGroup.Text & ")", strPath, cboGroup.Text, True, "$(Start Menu)") Then
            'Win7��·��
            If Not OSfCreateShellLink("..\..\..\..\..\DeskTop", StrName & "����̨(" & cboGroup.Text & ")", strPath, cboGroup.Text, True, "$(Start Menu)") Then
                Call OSfCreateShellLink("..\..\..\..\..\����", StrName & "����̨(" & cboGroup.Text & ")", strPath, cboGroup.Text, True, "$(Start Menu)")
            End If
        End If
    End If
End Sub

Private Sub cmdStartup_Click()
    Dim strPath As String
    Dim StrName As String
    
    '������ݷ�ʽ
    strPath = GetSetting("ZLSOFT", "����ȫ��", "����·��", "C:\AppSoft\ZLHIS+.exe")
    StrName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "")

    '�������˵�
    If Not OSfCreateShellLink("\Startup", StrName & "����̨(" & cboGroup.Text & ")", strPath, cboGroup.Text, True, "$(Programs)") Then
        Call OSfCreateShellLink("\����", StrName & "����̨(" & cboGroup.Text & ")", strPath, cboGroup.Text, True, "$(Programs)")
    End If
End Sub

Private Sub cmdFavorite_Click(Index As Integer)
    Dim lst As ListItem, lngIndex As Long
    
    If Index <> 0 And lvwFavorite.SelectedItem Is Nothing Then Exit Sub
    
    Select Case Index
        Case 0 '����
            If tvwMain.SelectedItem Is Nothing Then Exit Sub
            
            With tvwMain.SelectedItem
                If lvwFavorite.ListItems.Count >= 10 Then
                    MsgBox "���ֻ������10������ģ�飡", vbInformation, gstrSysName
                    Exit Sub
                End If
                For Each lst In lvwFavorite.ListItems
                    If lst.Tag = .Tag Then
                        MsgBox "��" & .Text & "���Ѿ����ڡ�", vbInformation, gstrSysName
                        Exit Sub
                    End If
                Next
                If Right(.Tag, 1) = "_" Then Exit Sub
                
                Set lst = lvwFavorite.ListItems.Add(, , .Text, .Image, .Image)
                lst.Tag = .Tag
                lst.Selected = True
                lst.EnsureVisible
            End With
        Case 1 'ɾ��
            lngIndex = lvwFavorite.SelectedItem.Index
            lvwFavorite.ListItems.Remove lngIndex
            
            If lvwFavorite.ListItems.Count = 0 Then Exit Sub
            If lngIndex > lvwFavorite.ListItems.Count Then
                lvwFavorite.ListItems.item(lngIndex - 1).Selected = True
            Else
                lvwFavorite.ListItems.item(lngIndex).Selected = True
            End If
        Case 2 'ǰ��
            With lvwFavorite.SelectedItem
                If .Index = 1 Then Exit Sub
                Set lst = lvwFavorite.ListItems.Add(.Index - 1, , .Text, .Icon, .SmallIcon)
                lst.Tag = .Tag
                
                lngIndex = .Index
                lst.Selected = True
                lvwFavorite.ListItems.Remove lngIndex
                lst.EnsureVisible
            End With
        Case 3 '����
            With lvwFavorite.SelectedItem
                If .Index = lvwFavorite.ListItems.Count Then Exit Sub
                Set lst = lvwFavorite.ListItems.Add(.Index + 2, , .Text, .Icon, .SmallIcon)
                lst.Tag = .Tag
                
                lngIndex = .Index
                lst.Selected = True
                lvwFavorite.ListItems.Remove lngIndex
                lst.EnsureVisible
            End With
    End Select
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, "ZL9AppTool\" & Me.Name, 0)
End Sub

Private Sub cmdOK_Click()
    Dim lst As ListItem, i As Integer
    Dim strϵͳ As String, str��� As String, strͼ�� As String, str���� As String
    Dim strOldStyle As String
    strOldStyle = zlDatabase.GetPara("����̨")
    '���浼�����
    For i = optStyle.LBound To optStyle.UBound
        If optStyle(i).Value = True Then
            Call zlDatabase.SetPara("����̨", optStyle(i).Tag)
            SaveSetting "ZLSOFT", "˽��ȫ��\" & gstrDbUser, "����̨", optStyle(i).Tag
            Exit For
        End If
    Next
    
    '���泣��ģ��
    For Each lst In lvwFavorite.ListItems
        strϵͳ = strϵͳ & "," & Val(Mid(lst.Tag, 1, InStr(lst.Tag, "_") - 1))
        str��� = str��� & "," & Mid(lst.Tag, InStr(lst.Tag, "_") + 1)
        strͼ�� = strͼ�� & "," & Mid(lst.Icon, 2)
        str���� = str���� & "," & lst.Text
    Next
    If strϵͳ <> "" Then
        strϵͳ = Mid(strϵͳ, 2)
        str��� = Mid(str���, 2)
        strͼ�� = Mid(strͼ��, 2)
        str���� = Mid(str����, 2)
    End If
    Call zlDatabase.SetPara("���ù���ģ��", strϵͳ & "|" & str��� & "|" & strͼ�� & "|" & str����)
    
    Call SaveRegister
    
    '����ҩƷ������ʾ��ʽ
    Call zlDatabase.SetPara("ҩƷ������ʾ", cboҩƷ������ʾ.ListIndex)
    Call zlDatabase.SetPara("����ҩƷ��ʾ", cbo����ҩƷ��ʾ.ListIndex)
    
    'д��zl9Comlib.glngAutoConnect
    zl9ComLib.gblnAutoConnect = IIf(chkLanJoin.Value = 1, True, False)
    If UCase(strOldStyle) <> UCase(zlDatabase.GetPara("����̨")) Then
        If gclsAppTool Is Nothing Then
            MsgBox "�µķ����Ҫ������������̨����Ч�����Ҫ���������·������������̨��", vbInformation, gstrSysName
        Else
            If MsgBox("�µķ����Ҫ������������̨����Ч���Ƿ�������������̨��", vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                gclsAppTool.IsRestart = True
            End If
        End If
    Else
        MsgBox "�޸���ɣ������޸���Ҫ�������ͻ��˺���Ч��"
    End If

    Unload Me
End Sub

Private Sub Form_Activate()
    Call tabMain_Click
End Sub

Private Sub Form_Load()
    Dim intIndex As Integer
    Dim strStyle As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHandle
    
    Call InitIcon
    Call FillCommon
    
    gstrSQL = "select ��� from ZLMENUS group by  ��� "
    
    rsTemp.CursorLocation = adUseClient
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    
    Do Until rsTemp.EOF
        cboGroup.AddItem rsTemp("���")
        If rsTemp("���") = gstrMenuSys Then intIndex = cboGroup.NewIndex
        rsTemp.MoveNext
    Loop
    mintIndex = -3
    cboGroup.ListIndex = intIndex
    
    strStyle = UCase(zlDatabase.GetPara("����̨"))
    
    For intIndex = optStyle.LBound To optStyle.UBound
        If UCase(optStyle(intIndex).Tag) = strStyle Then
            optStyle(intIndex).Value = True
            Exit For
        End If
    Next
    Call optStyle_Click(0)

    Call LoadCustom
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub InitIcon()
'���ܣ���ͼ��װ�뵽�ؼ���
    Dim i As Long
    ils16.ListImages.Clear
    ils16.ImageWidth = 16
    ils16.ImageHeight = 16
    
    With ils16.ListImages
        For i = glngLBound To glngUBound
            .Add , "K" & i, LoadResPicture(i, vbResIcon)
        Next
    End With
End Sub

Private Sub FillTree(ByVal str��� As String)
    Dim strSQL As String
    Dim strTemp As String
    Dim rsMenus As New ADODB.Recordset
    
    On Error GoTo ErrH
    gstrSQL = "select * " & _
            " from zlMenus" & _
            " start with �ϼ�ID is null and ���=[1] " & _
            " connect by prior ID =�ϼ�ID  and ���=[1] order by level,ID"
    rsMenus.CursorLocation = adUseClient
    Set rsMenus = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str���)
    With rsMenus
        tvwMain.Nodes.Clear
        Do While Not .EOF
            If IsNull(rsMenus("ͼ��")) Or rsMenus("ͼ��") = 0 Then
                strTemp = IIf(IsNull(.Fields("ģ��").Value), "K99", "K100")
            Else
                strTemp = "K" & rsMenus("ͼ��")
            End If
            If IsNull(.Fields("�ϼ�ID")) Then
                tvwMain.Nodes.Add , tvwChild, "C" & .Fields("ID").Value, .Fields("����").Value, strTemp, strTemp
            Else
               tvwMain.Nodes.Add "C" & .Fields("�ϼ�ID").Value, tvwChild, "C" & .Fields("ID").Value, .Fields("����").Value, strTemp, strTemp
            End If
            tvwMain.Nodes("C" & .Fields("ID").Value).Tag = IIf(IsNull(.Fields("ϵͳ")), "", .Fields("ϵͳ")) & "_" & IIf(IsNull(.Fields("ģ��")), "", .Fields("ģ��"))
            .MoveNext
        Loop
    End With
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cboGroup_Click()
    If mintIndex <> cboGroup.ListIndex Then FillTree cboGroup.Text
    mintIndex = cboGroup.ListIndex
End Sub


Private Sub optStyle_Click(Index As Integer)
    Dim i As Integer
    
    For i = optStyle.LBound To optStyle.UBound
        If optStyle(i).Value = True Then
            picPreview.Picture = LoadResPicture(optStyle(i).Tag, vbResBitmap)
            Exit For
        End If
    Next
End Sub

Private Sub picBorder_Paint()
    Dim rc As RECT
    
    With picBorder
        rc.Left = .ScaleLeft + 1
        rc.Right = .ScaleWidth - 2
        rc.Top = .ScaleTop + 1
        rc.Bottom = .ScaleHeight - 2
    End With
    DrawEdge picBorder.hDC, rc, EDGE_RAISED, BF_RECT
End Sub

Private Sub tabMain_Click()
    fra(1).Visible = False
    fra(2).Visible = False
    fra(3).Visible = False
    fra(4).Visible = False
    fra(tabMain.SelectedItem.Index).Visible = True
    fra(tabMain.SelectedItem.Index).ZOrder 0
End Sub

Private Sub FillCommon()
'���ܣ�װ�볣�õĳ���
    Dim varϵͳ As Variant, var��� As Variant, varͼ�� As Variant, var���� As Variant
    Dim lngMax As Long, lngCount As Long, lst As ListItem, strValue As String
    
    strValue = zlDatabase.GetPara("���ù���ģ��")
    If UBound(Split(strValue, "|")) < 3 Then Exit Sub
    varϵͳ = Split(Split(strValue, "|")(0), ",")
    var��� = Split(Split(strValue, "|")(1), ",")
    varͼ�� = Split(Split(strValue, "|")(2), ",")
    var���� = Split(Split(strValue, "|")(3), ",")
    
    lngMax = IIf(UBound(varϵͳ) > UBound(var���), UBound(varϵͳ), UBound(var���))
    lngMax = IIf(lngMax > UBound(varͼ��), lngMax, UBound(varͼ��))
    lngMax = IIf(lngMax > UBound(var����), lngMax, UBound(var����))
    If lngMax = -1 Then Exit Sub
    
    For lngCount = 0 To lngMax
        Set lst = lvwFavorite.ListItems.Add(, , var����(lngCount), "K" & varͼ��(lngCount), "K" & varͼ��(lngCount))
        lst.Tag = varϵͳ(lngCount) & "_" & var���(lngCount)
    Next
End Sub

Private Sub tvwMain_DblClick()
    Call cmdFavorite_Click(0)
End Sub

Private Sub LoadCustom()
'����û�ϰ�ߵĳ�ʼ������
    Dim lng���� As Long
    Dim strPath As String
    Dim intҩƷ������ʾ As Integer
    Dim int����ҩƷ��ʾ As Integer
    
    '����ƥ��
    If Val(zlDatabase.GetPara("����ƥ��")) = 0 Then
        opt(opt_˫��ƥ��).Value = True
        opt(opt_����ƥ��).Value = False
    Else
        opt(opt_˫��ƥ��).Value = False
        opt(opt_����ƥ��).Value = True
    End If
    
    '�������뷨
    Call ChooseIME(cmbIME)
    
    '����ƥ�䷽ʽ�л�
    chkIMETurn.Value = IIf(Val(zlDatabase.GetPara("����ƥ�䷽ʽ�л�", , , 1)) = 1, 1, 0)
    
    '�������ɷ�ʽ
    lng���� = Val(zlDatabase.GetPara("���뷽ʽ"))
    If lng���� = 0 Then
        opt����(0).Value = True
    Else
        opt����(1).Value = True
    End If
    
    '�Զ����б�����
    Call zlCommFun.GetRegValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "zlExplorer", strPath)
    chkAutoStart.Value = IIf(Trim(strPath) <> "", 1, 0)
    
    '�Զ��ر�Windows
    chkShutDown.Value = IIf(Val(zlDatabase.GetPara("�ر�Windows")) = 1, 1, 0)
    
    '������������
    chkAutoHide.Value = IIf(Val(zlDatabase.GetPara("������������")) = 1, 1, 0)
    
    '֪ͨ�������
    txtTime.Text = zlDatabase.GetPara("�ʼ���Ϣ�������")
    If (Val(txtTime.Text) < 10 Or Val(txtTime.Text) > 300) And Val(txtTime.Text) <> 0 Then txtTime.Text = 60
    
    '��������Զ�����
    chkLanJoin.Value = IIf(Val(zlDatabase.GetPara("��������Զ�����")) = 1, 1, 0)
    '�Զ�����
    txtAutoLock.Text = Val(zlDatabase.GetPara("�Զ�����"))
    If Val(txtAutoLock.Text) = 0 Then
        chkAutoLock.Value = 0
        txtAutoLock.Text = "5"
    Else
        chkAutoLock.Value = 1
    End If
    
    'Զ������
    txtRemote.Text = Val(zlDatabase.GetPara("����Զ�̿���"))
    If Val(txtRemote.Text) = 0 Then
        txtRemote.Text = "1001"
    End If
    
    'ҩƷ����
    intҩƷ������ʾ = Val(zlDatabase.GetPara("ҩƷ������ʾ", , , 2))
    int����ҩƷ��ʾ = Val(zlDatabase.GetPara("����ҩƷ��ʾ"))
    
    If intҩƷ������ʾ < 0 Or intҩƷ������ʾ > 2 Then intҩƷ������ʾ = 2
    If int����ҩƷ��ʾ < 0 Or int����ҩƷ��ʾ > 1 Then int����ҩƷ��ʾ = 0
    
    cboҩƷ������ʾ.Clear
    cboҩƷ������ʾ.AddItem "0-��ʾͨ����"
    cboҩƷ������ʾ.AddItem "1-��ʾ��Ʒ��"
    cboҩƷ������ʾ.AddItem "2-ͬʱ��ʾͨ��������Ʒ��"
    cboҩƷ������ʾ.ListIndex = intҩƷ������ʾ
    
    cbo����ҩƷ��ʾ.Clear
    cbo����ҩƷ��ʾ.AddItem "0-������ƥ����ʾ"
    cbo����ҩƷ��ʾ.AddItem "1-�̶���ʾͨ��������Ʒ��"
    cbo����ҩƷ��ʾ.ListIndex = int����ҩƷ��ʾ
End Sub

Private Sub SaveRegister()
'���浽ע����е���Ϣ
    Dim lng���� As Long
    Dim strExeName As String
        
    '��ΪAppTools�ǹ���ʹ�ã�Ϊ����9�汾ϵͳ�������ڸ���ģ���п���ʹ�õĲ�����ͬʱ������ע�����
    Call zlDatabase.SetPara("����ƥ��", IIf(opt(opt_˫��ƥ��).Value = True, "0", "1"))
    SaveSetting "ZLSOFT", "����ģ��\����", "����ƥ��", IIf(opt(opt_˫��ƥ��).Value = True, "0", "1")
    
    Call zlDatabase.SetPara("���뷨", IIf(cmbIME.Text = "���Զ�����", "", cmbIME.Text))
    SaveSetting "ZLSOFT", "˽��ȫ��\" & gstrDbUser, "���뷨", IIf(cmbIME.Text = "���Զ�����", "", cmbIME.Text)
    
    Call zlDatabase.SetPara("����ƥ�䷽ʽ�л�", chkIMETurn.Value)
    
    For lng���� = opt����.LBound To opt����.UBound
        If opt����(lng����).Value = True Then
            Call zlDatabase.SetPara("���뷽ʽ", lng����)
            SaveSetting "ZLSOFT", "˽��ȫ��\" & gstrDbUser, "��������", lng����
            Exit For
        End If
    Next

    '�Զ����б�����
    If chkAutoStart.Value = 1 Then
        strExeName = GetSetting("ZLSOFT", "����ȫ��", "����·��", "C:\AppSoft\ZLHIS+.exe")
        strExeName = Replace(strExeName, ":\\", ":\")
        Call zlCommFun.SetRegValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "zlExplorer", strExeName)
    Else
        Call zlCommFun.DeleteRegValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", "zlExplorer")
    End If
    
    '������������
    Call zlDatabase.SetPara("������������", chkAutoHide.Value)
    
    '�Զ��ر�Windows
    Call zlDatabase.SetPara("�ر�Windows", chkShutDown.Value)
    
    '��Ϣ֪ͨ�������
    If (Val(txtTime.Text) < 10 Or Val(txtTime.Text) > 300) And Val(txtTime.Text) <> 0 Then txtTime.Text = 60
    Call zlDatabase.SetPara("�ʼ���Ϣ�������", Val(txtTime.Text))
    
    '��������Զ�����
    Call zlDatabase.SetPara("��������Զ�����", chkLanJoin.Value)
    '�Զ�����
    Call zlDatabase.SetPara("�Զ�����", IIf(chkAutoLock.Value = 0, "", Val(txtAutoLock.Text)))
    '����Զ�̿��ƶ˿�
    Call zlDatabase.SetPara("����Զ�̿���", IIf(Val(txtRemote.Text) = 0, "-1", Val(txtRemote.Text)))
End Sub

Private Sub txtAutoLock_GotFocus()
    Call zlControl.TxtSelAll(txtAutoLock)
End Sub

Private Sub txtAutoLock_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtAutoLock_Validate(Cancel As Boolean)
    If Val(txtAutoLock.Text) < 0 Or Val(txtAutoLock.Text) > 999 Then txtAutoLock.Text = 5
End Sub

Private Sub txtTime_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtTime_Validate(Cancel As Boolean)
    If (Val(txtTime.Text) < 10 Or Val(txtTime.Text) > 300) And Val(txtTime.Text) <> 0 Then txtTime.Text = 60
End Sub

Private Sub txtTime_GotFocus()
    Call zlControl.TxtSelAll(txtTime)
End Sub

Private Sub txtRemote_GotFocus()
    Call zlControl.TxtSelAll(txtRemote)
End Sub

Private Sub txtRemote_KeyPress(KeyAscii As Integer)
    If InStr("0123456789-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
