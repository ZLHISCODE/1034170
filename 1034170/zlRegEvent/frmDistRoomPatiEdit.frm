VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{D01C2596-4FE0-4EA9-9EE8-D97BE62A1165}#2.1#0"; "ZlPatiAddress.ocx"
Begin VB.Form frmDistRoomPatiEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������Ϣ�༭"
   ClientHeight    =   8955
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7725
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "����"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&R)"
      Height          =   420
      Left            =   3900
      TabIndex        =   31
      Top             =   8505
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   420
      Left            =   5490
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   8505
      Width           =   1500
   End
   Begin VB.PictureBox picAddInfo 
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   8430
      ScaleHeight     =   825
      ScaleWidth      =   1755
      TabIndex        =   59
      Top             =   3630
      Visible         =   0   'False
      Width           =   1755
      Begin XtremeSuiteControls.TaskPanel wndTaskPanel 
         Height          =   435
         Left            =   330
         TabIndex        =   60
         Top             =   150
         Width           =   855
         _Version        =   589884
         _ExtentX        =   1508
         _ExtentY        =   767
         _StockProps     =   64
         VisualTheme     =   7
         ItemLayout      =   2
         HotTrackStyle   =   1
      End
   End
   Begin XtremeSuiteControls.TabControl tabPage 
      Height          =   1905
      Left            =   8340
      TabIndex        =   58
      Top             =   540
      Visible         =   0   'False
      Width           =   1755
      _Version        =   589884
      _ExtentX        =   3096
      _ExtentY        =   3360
      _StockProps     =   64
   End
   Begin VB.PictureBox picPatiInfo 
      BorderStyle     =   0  'None
      Height          =   8415
      Left            =   0
      ScaleHeight     =   8415
      ScaleWidth      =   7725
      TabIndex        =   34
      Top             =   0
      Width           =   7725
      Begin ZlPatiAddress.PatiAddress padd��ͥ��ַ 
         Height          =   360
         Left            =   1140
         TabIndex        =   9
         Tag             =   "��ͥ��ַ"
         Top             =   1005
         Visible         =   0   'False
         Width           =   6270
         _ExtentX        =   11060
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   100
      End
      Begin VB.TextBox txt����� 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   5205
         TabIndex        =   3
         Top             =   120
         Width           =   2205
      End
      Begin VB.TextBox txt���� 
         Height          =   360
         IMEMode         =   2  'OFF
         Left            =   5220
         TabIndex        =   6
         Top             =   600
         Width           =   1575
      End
      Begin VB.ComboBox cbo�Ա� 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   3330
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   120
         Width           =   880
      End
      Begin VB.ComboBox cbo�ѱ� 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   5955
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1845
         Width           =   1455
      End
      Begin VB.ComboBox cbo���� 
         Height          =   360
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2505
         Width           =   2310
      End
      Begin VB.ComboBox cbo���� 
         Height          =   360
         Left            =   5235
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   2505
         Width           =   2175
      End
      Begin VB.ComboBox cbo���� 
         Height          =   360
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2955
         Width           =   2340
      End
      Begin VB.ComboBox cboְҵ 
         Height          =   360
         Left            =   5235
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2955
         Width           =   2175
      End
      Begin VB.TextBox txt���֤�� 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1140
         MaxLength       =   18
         TabIndex        =   19
         Top             =   3450
         Width           =   2325
      End
      Begin VB.TextBox txt��λ���� 
         Height          =   360
         Left            =   4560
         MaxLength       =   100
         TabIndex        =   20
         Top             =   3450
         Width           =   2450
      End
      Begin VB.TextBox txt��λ�绰 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1140
         MaxLength       =   20
         TabIndex        =   22
         Top             =   3885
         Width           =   2325
      End
      Begin VB.TextBox txt��λ�ʱ� 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   5925
         MaxLength       =   6
         TabIndex        =   23
         Top             =   3885
         Width           =   1455
      End
      Begin VB.TextBox txt��ͥ�绰 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1140
         MaxLength       =   20
         TabIndex        =   24
         Top             =   4320
         Width           =   2325
      End
      Begin VB.TextBox txt��ͥ�ʱ� 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   5925
         MaxLength       =   6
         TabIndex        =   25
         Top             =   4320
         Width           =   1455
      End
      Begin VB.CommandButton cmd��ͥ��ַ 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6990
         TabIndex        =   0
         ToolTipText     =   "�ȼ�F3"
         Top             =   1005
         Width           =   360
      End
      Begin VB.CommandButton cmd��λ���� 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7005
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   3450
         Width           =   360
      End
      Begin VB.ComboBox cbo���ʽ 
         Height          =   360
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1845
         Width           =   1455
      End
      Begin VB.CommandButton cmd���� 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6345
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "�ȼ�:F3"
         Top             =   6600
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.TextBox txt���� 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   29
         Top             =   6480
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.ComboBox cbo���䵥λ 
         Height          =   360
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   600
         Width           =   585
      End
      Begin VB.TextBox txtPatient 
         Height          =   360
         Left            =   1140
         MaxLength       =   100
         TabIndex        =   1
         Top             =   120
         Width           =   1515
      End
      Begin VB.TextBox txtPatiMCNO 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   1140
         MaxLength       =   30
         TabIndex        =   10
         Top             =   1395
         Width           =   2205
      End
      Begin VB.TextBox txtPatiMCNO 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   5205
         MaxLength       =   30
         TabIndex        =   11
         Top             =   1395
         Width           =   2205
      End
      Begin VB.ComboBox cboҽ����� 
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   3735
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1845
         Width           =   1455
      End
      Begin VB.ComboBox cbo��ͥ��ַ 
         Height          =   360
         Left            =   1140
         TabIndex        =   8
         Top             =   1005
         Width           =   5835
      End
      Begin VB.TextBox txtEdit 
         Height          =   360
         Index           =   0
         Left            =   1140
         MaxLength       =   64
         TabIndex        =   26
         Top             =   4800
         Width           =   2325
      End
      Begin MSComctlLib.ListView lvwItems 
         Height          =   1515
         Left            =   2010
         TabIndex        =   33
         Top             =   6510
         Visible         =   0   'False
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   2672
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         PictureAlignment=   1
         _Version        =   393217
         Icons           =   "imgList"
         SmallIcons      =   "imgList"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh���� 
         Height          =   1440
         Left            =   15
         TabIndex        =   28
         ToolTipText     =   "F2:�޸�,F3:ѡ��"
         Top             =   6480
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   2540
         _Version        =   393216
         Rows            =   4
         Cols            =   1
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         AllowBigSelection=   0   'False
         HighLight       =   0
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
         FormatString    =   "<����ҩ��                                               "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   1
      End
      Begin MSComctlLib.ImageList imgList 
         Left            =   2820
         Top             =   7320
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
               Picture         =   "frmDistRoomPatiEdit.frx":0000
               Key             =   "Itemps"
               Object.Tag             =   "Itemgm"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDistRoomPatiEdit.frx":059A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox txt����ʱ�� 
         Height          =   360
         Left            =   2625
         TabIndex        =   5
         Top             =   600
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt�������� 
         Height          =   360
         Left            =   1140
         TabIndex        =   4
         Top             =   600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   635
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "YYYY-MM-DD"
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin zl9RegEvent.UCPatiVitalSigns UCPatiVitalSigns 
         Height          =   945
         Left            =   540
         TabIndex        =   27
         Top             =   5400
         Width           =   7080
         _ExtentX        =   12488
         _ExtentY        =   1667
         TextBackColor   =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   0
         XDis            =   100
         YDis            =   120
         LabToTxt        =   120
      End
      Begin VB.Line Line9 
         BorderColor     =   &H80000014&
         X1              =   -60
         X2              =   7515
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         X1              =   -60
         X2              =   7395
         Y1              =   2340
         Y2              =   2340
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   240
         Left            =   4440
         TabIndex        =   57
         Top             =   180
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   240
         Left            =   540
         TabIndex        =   56
         Top             =   180
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   240
         Left            =   2820
         TabIndex        =   55
         Top             =   180
         Width           =   480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   240
         Left            =   4680
         TabIndex        =   54
         Top             =   660
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����״��"
         Height          =   240
         Left            =   60
         TabIndex        =   53
         Top             =   3015
         Width           =   960
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ְҵ"
         Height          =   240
         Left            =   4665
         TabIndex        =   52
         Top             =   3015
         Width           =   480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   240
         Left            =   4680
         TabIndex        =   51
         Top             =   2565
         Width           =   480
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   240
         Left            =   540
         TabIndex        =   50
         Top             =   2565
         Width           =   480
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���֤��"
         Height          =   240
         Left            =   60
         TabIndex        =   49
         Top             =   3510
         Width           =   960
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ����"
         Height          =   240
         Left            =   3555
         TabIndex        =   48
         Top             =   3510
         Width           =   960
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ�绰"
         Height          =   240
         Left            =   60
         TabIndex        =   47
         Top             =   3945
         Width           =   960
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ�ʱ�"
         Height          =   240
         Left            =   4890
         TabIndex        =   46
         Top             =   3945
         Width           =   960
      End
      Begin VB.Label lbl��ͥ��ַ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ͥ��ַ"
         Height          =   240
         Left            =   60
         TabIndex        =   45
         Top             =   1065
         Width           =   960
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ͥ�绰"
         Height          =   240
         Left            =   60
         TabIndex        =   44
         Top             =   4380
         Width           =   960
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ͥ�ʱ�"
         Height          =   240
         Left            =   4890
         TabIndex        =   43
         Top             =   4380
         Width           =   960
      End
      Begin VB.Label lbl�ѱ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ѱ�"
         Height          =   240
         Left            =   5415
         TabIndex        =   42
         Top             =   1905
         Width           =   480
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         X1              =   -60
         X2              =   7740
         Y1              =   2325
         Y2              =   2325
      End
      Begin VB.Label lbl���ʽ 
         BackStyle       =   0  'Transparent
         Caption         =   "���ʽ"
         Height          =   240
         Left            =   45
         TabIndex        =   41
         Top             =   1905
         Width           =   975
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         X1              =   -60
         X2              =   7740
         Y1              =   6360
         Y2              =   6360
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000014&
         X1              =   -60
         X2              =   7260
         Y1              =   7935
         Y2              =   7935
      End
      Begin VB.Label Label18 
         Caption         =   "��λ������ҩ�ﴦ,F2�޸�,F3ѡ��.�����ǰ��������,���������ֿ��޸Ĺ���ҩ������,��������������Ϊ�ؼ��ְ����롢���ơ�������ҹ���ҩ��."
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   315
         TabIndex        =   40
         Top             =   7995
         Width           =   6975
      End
      Begin VB.Label lbl�������� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   240
         Left            =   45
         TabIndex        =   39
         Top             =   660
         Width           =   975
      End
      Begin VB.Label lblPatiMCNO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ����"
         Height          =   240
         Index           =   0
         Left            =   300
         TabIndex        =   38
         Top             =   1455
         Width           =   720
      End
      Begin VB.Label lblPatiMCNO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��֤ҽ����"
         Height          =   240
         Index           =   1
         Left            =   3960
         TabIndex        =   37
         Top             =   1455
         Width           =   1200
      End
      Begin VB.Label lblҽ����� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ�����"
         Height          =   240
         Left            =   2700
         TabIndex        =   36
         Top             =   1905
         Width           =   960
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�໤��"
         Height          =   300
         Index           =   22
         Left            =   300
         TabIndex        =   35
         Top             =   4830
         Width           =   780
      End
      Begin VB.Line Line7 
         BorderColor     =   &H80000014&
         X1              =   -60
         X2              =   7515
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Line Line8 
         BorderColor     =   &H80000010&
         X1              =   -60
         X2              =   7740
         Y1              =   5280
         Y2              =   5280
      End
   End
End
Attribute VB_Name = "frmDistRoomPatiEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Public mstrNO As String '�Һŵ���
Public mlng�Һ�ID As Long  '�Һ�ID
Public mlng����ID As Long
Public mlng���� As Long
Public mstrҽ���� As String  '��������ʱ,�����ز���ҽ����,��Ϣ�ᵼ��,�޸Ĳ�����ҽ���Ŷ�ʧ
Public m���￨�� As String
Public m��֤���� As String
Public mbytType As Byte         '���ݸ��洢����,��ȷ����������
Public mlngOutModeMC As Long '����ҽ�����õ����ʽҽ������
Public mstrPrivs As String
Public mblnChange As Boolean
Public mstr���� As String
Public mstr�Ա� As String
Public mstr���� As String
Public mstr����_���� As String '������Ϣ�е�����
Public mstr����_�Ա� As String '������Ϣ�е��Ա�
Public mstr����_���� As String '������Ϣ�е�����
Public mstr�������� As String '��������
Public mblnҽ��ҵ�� As Boolean  '�Ƿ�����ҽ��ҽҵ��
Public mblnStructAdress As Boolean  '���˵�ַ�ṹ��¼��
Public mblnShowTown As Boolean      '�����ַ�ṹ��¼��

Private mrs��ͥ��ַ As ADODB.Recordset  '�����ͥ��ַ,��ʼʱ��ȡ������
Private mstrSQL As String
Private mDateSys As Date

Private Enum mIndex
    idx_�໤�� = 0
    idx_��� = 1
    idx_���� = 2
    idx_���� = 3
End Enum

Private mobjPlugIn As Object '73935,Ƚ����,20114-7-3,���������ƵĽ���Ƕ�뵽������Ϣ�༭��
Private mlngPlugInHwnd As Long
Private Enum mPageIndex
    ������Ϣ = 1
    ������Ϣ = 2
End Enum
Private mobjPubPatient As Object
Private mblnGetBirth As Boolean '�ж��Ƿ�ͨ�������������

Private Sub cbo�ѱ�_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = 13 And cbo�ѱ�.ListIndex <> -1 Then Call zlCommFun.PressKey(vbKeyTab)
    
    If SendMessage(cbo�ѱ�.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo�ѱ�.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo�ѱ�.ListIndex = lngIdx
    If cbo�ѱ�.ListIndex = -1 And cbo�ѱ�.ListCount > 0 Then cbo�ѱ�.ListIndex = 0
End Sub

Private Sub cbo���ʽ_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo���ʽ.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo���ʽ.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo���ʽ.ListIndex = lngIdx
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo����.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo����.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo����.ListIndex = lngIdx
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo����.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo����.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo����.ListIndex = lngIdx
End Sub
Private Sub Load��ͥ��ַ()
    Dim strSQL As String, strFile As String
    Dim fld As Field
    Dim fso As Scripting.FileSystemObject
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    strFile = App.Path & "\ZLAddressForRegEvent.Adtg"
    
    Set mrs��ͥ��ַ = New ADODB.Recordset
    
    On Error Resume Next
    If fso.FileExists(strFile) Then
        mrs��ͥ��ַ.Open strFile, "Provider=MSPersist", adOpenKeyset, adLockOptimistic, adCmdFile   '��Updateʱ������
    End If
    Err.Clear
    On Error GoTo errH
    
    If mrs��ͥ��ַ.State = 0 Then
        strSQL = "Select 'ϵͳ' as ���,����,����,1 as ���� From ����"
        Call zlDatabase.OpenRecordset(mrs��ͥ��ַ, strSQL, Me.Caption)            '������adUseClient���ܽ�����
        
        If Not mrs��ͥ��ַ.EOF Then
            '��������:����,����
            Set fld = mrs��ͥ��ַ.Fields(1)
            fld.Properties("Optimize") = True
            Set fld = mrs��ͥ��ַ.Fields(2)
            fld.Properties("Optimize") = True
            
            If fso.FileExists(strFile) Then
                Kill strFile
            End If
            mrs��ͥ��ַ.Save strFile, adPersistADTG
        End If
        mrs��ͥ��ַ.Close
        mrs��ͥ��ַ.Open strFile, "Provider=MSPersist", adOpenKeyset, adLockOptimistic, adCmdFile   '��Updateʱ������
        
    End If
    
    lbl��ͥ��ַ.ToolTipText = "�붨�ڱ��ݱ���[��ͥ��ַ]�����ļ�:" & strFile
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub cbo��ͥ��ַ_GotFocus()
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub cbo��ͥ��ַ_LostFocus()
    Call zlCommFun.OpenIme
End Sub

Private Sub cbo��ͥ��ַ_KeyDown(KeyCode As Integer, Shift As Integer)
    '�˹��̴������������ݵ�ɾ��,�Լ���������ʱ���������б�
    '�����б���ʱ,�������ɾ����ʱ,��ɾ�������¼
    
    Dim str��ͥ��ַ As String
    
    If KeyCode = vbKeyDelete Then
        str��ͥ��ַ = cbo��ͥ��ַ.Text
        If Not mrs��ͥ��ַ Is Nothing Then
            If mrs��ͥ��ַ.State = 1 And str��ͥ��ַ <> "" Then
                If cbo��ͥ��ַ.SelText = str��ͥ��ַ And SendMessage(cbo��ͥ��ַ.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = True Then
                    mrs��ͥ��ַ.Filter = "����='" & str��ͥ��ַ & "'"
                    If Not mrs��ͥ��ַ.EOF Then
                        mrs��ͥ��ַ.Delete adAffectCurrent
                        mrs��ͥ��ַ.Update
                    End If
                End If
            End If
        End If
    ElseIf KeyCode = vbKeyDown And cbo��ͥ��ַ.Text <> "" Then
        If SendMessage(cbo��ͥ��ַ.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 Then Call SendMessage(cbo��ͥ��ַ.Hwnd, CB_SHOWDROPDOWN, True, 0&)
    ElseIf KeyCode = vbKeyF3 Then
        cmd��ͥ��ַ.SetFocus
        Call cmd��ͥ��ַ_Click
    End If
End Sub

Private Sub cbo��ͥ��ַ_KeyUp(KeyCode As Integer, Shift As Integer)
    '��ʱtext���ѽ����������Ϣ
    '���¼�����ɾ�����˸��,ɾ������������Ŀ��,�����б�����������Ӧ������ɸѡ
    '���ȫ�����ֶ�ɾ����,����������б�����
        
    Dim str��ͥ��ַ As String, i As Long
    Dim lngλ�� As Long
    
    If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Then
        If mrs��ͥ��ַ Is Nothing Then Exit Sub
        
        str��ͥ��ַ = cbo��ͥ��ַ.Text                      '��ʱ,���ѡ���˲�������,��ѡ��������Ѿ���ɾ��
        lngλ�� = cbo��ͥ��ַ.SelStart
        
        If mrs��ͥ��ַ.State = 1 And Len(str��ͥ��ַ) > 1 Then
            If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Left(str��ͥ��ַ, 1))) > 0 Then
                mrs��ͥ��ַ.Filter = "���� like '" & gstrLike & UCase(str��ͥ��ַ) & "*'"
            Else
                mrs��ͥ��ַ.Filter = "���� Like '" & gstrLike & str��ͥ��ַ & "*'"
            End If
            
            If Not mrs��ͥ��ַ.EOF Then
                
                If mrs��ͥ��ַ.RecordCount <> cbo��ͥ��ַ.ListCount Then
                    Call SendMessage(cbo��ͥ��ַ.Hwnd, CB_RESETCONTENT, 0, 0)
                    mrs��ͥ��ַ.Sort = "���� Desc,����"
                    For i = 1 To mrs��ͥ��ַ.RecordCount
                        AddComboItem cbo��ͥ��ַ.Hwnd, CB_ADDSTRING, 0, mrs��ͥ��ַ!����
                        mrs��ͥ��ַ.MoveNext
                    Next
                    If SendMessage(cbo��ͥ��ַ.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 Then Call SendMessage(cbo��ͥ��ַ.Hwnd, CB_SHOWDROPDOWN, True, 0&)
                                        
                    cbo��ͥ��ַ.Text = str��ͥ��ַ
                    cbo��ͥ��ַ.SelStart = lngλ��
                End If
            Else
                Call SendMessage(cbo��ͥ��ַ.Hwnd, CB_SHOWDROPDOWN, False, 0&)
            End If
        ElseIf str��ͥ��ַ = "" Then
            cbo��ͥ��ַ.Clear
            Call SendMessage(cbo��ͥ��ַ.Hwnd, CB_SHOWDROPDOWN, False, 0&)
        End If
    End If
End Sub

Private Sub cbo��ͥ��ַ_KeyPress(KeyAscii As Integer)
    Dim i As Long
    Dim str���� As String
    Dim str��ͥ��ַ As String
    Dim lng�м������ As Long
    
    If mrs��ͥ��ַ Is Nothing Then Exit Sub
    
    '�ñ��ػ���ƥ������
    If KeyAscii <> 13 And KeyAscii <> vbKeyF4 And KeyAscii <> vbKeyEscape And _
        KeyAscii <> vbKeyBack And KeyAscii <> 26 And KeyAscii <> 3 And KeyAscii <> 22 Then   '26��ʾctrl+z,3-ctrl+c,22-ctrl+v
            
        If mrs��ͥ��ַ.State = 0 Or cbo��ͥ��ַ.Text = "" Then  '���һ����ʱ��ƥ��
            Exit Sub
        End If
       
        'ѡ���м䲿���ı�����������
        If cbo��ͥ��ַ.SelText <> "" And (cbo��ͥ��ַ.SelStart + cbo��ͥ��ַ.SelLength) <> Len(cbo��ͥ��ַ.Text) Then
            lng�м������ = cbo��ͥ��ַ.SelStart + 1
            cbo��ͥ��ַ.Text = Mid(cbo��ͥ��ַ.Text, 1, cbo��ͥ��ַ.SelStart) & Chr(KeyAscii) & Mid(cbo��ͥ��ַ.Text, cbo��ͥ��ַ.SelStart + cbo��ͥ��ַ.SelLength + 1)
            cbo��ͥ��ַ.SelText = ""
            str��ͥ��ַ = cbo��ͥ��ַ.Text
        Else
            '�������β��,�����м�ʱ,�������ѡ��
            If cbo��ͥ��ַ.SelStart = Len(cbo��ͥ��ַ.Text) Or (cbo��ͥ��ַ.SelStart + cbo��ͥ��ַ.SelLength) = Len(cbo��ͥ��ַ.Text) Then
                str��ͥ��ַ = Mid(cbo��ͥ��ַ.Text, 1, cbo��ͥ��ַ.SelStart) & Chr(KeyAscii)
            Else
                str��ͥ��ַ = Mid(cbo��ͥ��ַ.Text, 1, cbo��ͥ��ַ.SelStart) & Chr(KeyAscii) & Mid(cbo��ͥ��ַ.Text, cbo��ͥ��ַ.SelStart + 1)
                lng�м������ = cbo��ͥ��ַ.SelStart + 1
            End If
        End If
        
        If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Left(str��ͥ��ַ, 1))) > 0 Then
            mrs��ͥ��ַ.Filter = "���� like '" & gstrLike & UCase(str��ͥ��ַ) & "*'"
        Else
            mrs��ͥ��ַ.Filter = "���� Like '" & gstrLike & str��ͥ��ַ & "*'"
        End If
        
        If Not mrs��ͥ��ַ.EOF Then
            If mrs��ͥ��ַ.RecordCount <> cbo��ͥ��ַ.ListCount Then
                Call SendMessage(cbo��ͥ��ַ.Hwnd, CB_RESETCONTENT, 0, 0)
                mrs��ͥ��ַ.Sort = "���� Desc,����"
                For i = 1 To mrs��ͥ��ַ.RecordCount
                    AddComboItem cbo��ͥ��ַ.Hwnd, CB_ADDSTRING, 0, mrs��ͥ��ַ!����
                    mrs��ͥ��ַ.MoveNext
                Next
                If SendMessage(cbo��ͥ��ַ.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 Then Call SendMessage(cbo��ͥ��ַ.Hwnd, CB_SHOWDROPDOWN, True, 0&)
            End If
            
            i = KeyAscii    '���������ж��Ƿ��ǰ��˸�ɾ����
            KeyAscii = 0
            cbo��ͥ��ַ.Text = str��ͥ��ַ
            cbo��ͥ��ַ.SelStart = Len(cbo��ͥ��ַ.Text)

            mrs��ͥ��ַ.MoveFirst   '�����������ļ���,��ͬ��ȡ��һ�������
            If mrs��ͥ��ַ!���� = str��ͥ��ַ And i <> vbKeyBack Then
                mrs��ͥ��ַ.MoveNext
            End If
            If Not mrs��ͥ��ַ.EOF Then
                If InStr(1, mrs��ͥ��ַ!����, str��ͥ��ַ) > 0 Or mrs��ͥ��ַ!���� = UCase(str��ͥ��ַ) Then    '�������������������ݵ�һ����,��ѡ�л����������
                    i = Len(cbo��ͥ��ַ.Text)
                    cbo��ͥ��ַ.Text = mrs��ͥ��ַ!����
                    cbo��ͥ��ַ.SelStart = i
                    cbo��ͥ��ַ.SelLength = Len(cbo��ͥ��ַ.Text) - cbo��ͥ��ַ.SelStart
                    
                    If mrs��ͥ��ַ.RecordCount = 1 Then Exit Sub
                End If
            End If
            
        'û���ҵ�ƥ��Ļ�������ʱ,����������б�����
        Else
            Call SendMessage(cbo��ͥ��ַ.Hwnd, CB_RESETCONTENT, 0, 0)
            If SendMessage(cbo��ͥ��ַ.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 1 Then Call SendMessage(cbo��ͥ��ַ.Hwnd, CB_SHOWDROPDOWN, False, 0&)
            KeyAscii = 0
            cbo��ͥ��ַ.Text = str��ͥ��ַ
            cbo��ͥ��ַ.SelStart = Len(cbo��ͥ��ַ.Text)
        End If
        
        If lng�м������ > 0 Then cbo��ͥ��ַ.SelStart = lng�м������: cbo��ͥ��ַ.SelText = ""
        
    ElseIf KeyAscii = 13 Then
        'a.��û��ѡ���κ�����,����������Ϊ��,���Ϊ��ĩ��ʱ,ȷ������,��������Ϣ�����ػ���
        Call SendMessage(cbo��ͥ��ַ.Hwnd, CB_SHOWDROPDOWN, False, 0&)
        
        If cbo��ͥ��ַ.Text = "" Then
            If gbln��ͥ��ַ And txtPatient.Text <> "" Then
                Exit Sub
            Else
                Call zlCommFun.PressKey(vbKeyTab): Exit Sub
            End If
        End If
        
        '�����б���ʱ���س�,��λ��ĩβ
        If cbo��ͥ��ַ.SelText = cbo��ͥ��ַ.Text Then cbo��ͥ��ַ.SelStart = Len(cbo��ͥ��ַ.Text): Exit Sub
        
        If mrs��ͥ��ַ.State = 0 Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
       
        'a.������״̬�°��س�,û��ѡ���ı�
        If cbo��ͥ��ַ.SelText = "" Then
            str��ͥ��ַ = cbo��ͥ��ַ.Text
            mrs��ͥ��ַ.Filter = "����='" & str��ͥ��ַ & "'"
            If mrs��ͥ��ַ.EOF Then
                str���� = Mid(zlCommFun.zlGetSymbol(str��ͥ��ַ), 1, 10)
                If str���� <> UCase(str��ͥ��ַ) Then
                    With mrs��ͥ��ַ
                        .AddNew
                        !��� = "�û�"
                        !���� = str��ͥ��ַ
                        !���� = str����
                        !���� = 1
                        .Update                 '�ڴ���Unload��save
                    End With
                End If
            Else
                mrs��ͥ��ַ!���� = mrs��ͥ��ַ!���� + 1
                mrs��ͥ��ַ.Update
                
                If zlCommFun.IsCharAlpha(str��ͥ��ַ) Then
                    If mrs��ͥ��ַ.RecordCount = 1 Then
                        cbo��ͥ��ַ.Text = mrs��ͥ��ַ!����
                    Else
                        Call SendMessage(cbo��ͥ��ַ.Hwnd, CB_SHOWDROPDOWN, True, 0&)
                        Exit Sub
                    End If
                End If
            End If
            
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo����.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo����.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo����.ListIndex = lngIdx
End Sub

Private Sub cbo���䵥λ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cbo���䵥λ_LostFocus()
    If cbo���䵥λ.Tag <> cbo���䵥λ.Text Then
        mblnChange = False
        If mblnGetBirth Then
            txt��������.Text = ReCalcBirth(Trim(txt����.Text), IIf(cbo���䵥λ.Visible, cbo���䵥λ.Text, ""))
        End If
        mblnChange = True
    End If
    '69026,Ƚ����,2014-8-8,�����������
    '76703,Ƚ����,2014-8-15
    If mobjPubPatient Is Nothing Then Exit Sub
    If mobjPubPatient.CheckPatiAge(Trim(txt����.Text) & cbo���䵥λ.Text, _
            IIf(txt��������.Text = "____-__-__", "", txt��������.Text) & _
            IIf(txt����ʱ��.Text = "__:__", "", " " & txt����ʱ��.Text)) = False Then
        If txt����.Visible And txt����.Enabled Then txt����.SetFocus
    End If
End Sub

Private Sub cbo�Ա�_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo�Ա�.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo�Ա�.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo�Ա�.ListIndex = lngIdx
    If cbo�Ա�.ListIndex = -1 And cbo�Ա�.ListCount > 0 Then cbo�Ա�.ListIndex = 0
    
End Sub

Private Sub cboҽ�����_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cboҽ�����.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cboҽ�����.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cboҽ�����.ListIndex = lngIdx
End Sub

Private Sub cboְҵ_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cboְҵ.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cboְҵ.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cboְҵ.ListIndex = lngIdx
End Sub

Private Sub cmdCancel_Click()
    gblnOk = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Long, blnTrans As Boolean, lngTmp As Long, lngLen As Long
    Dim strDate As String, str���� As String, strMCAccount As String, strTmp As String
    Dim str���� As String, str�������� As String
    Dim cllPro As Collection
    Dim blnPlugInCheck As Boolean
    On Error GoTo errH
    
    txtPatient.Text = Trim(txtPatient.Text)
    txt����.Text = Trim(txt����.Text)
    
    If CheckValied = False Then Exit Sub

    strMCAccount = Trim(txtPatiMCNO(0).Text)
    If mlngOutModeMC = 920 And strMCAccount <> txtPatiMCNO(0).Tag And strMCAccount <> "" Then
        strMCAccount = UCase(strMCAccount)
        If CheckExistsMCNO(strMCAccount) Then
            If txtPatiMCNO(0).Visible And txtPatiMCNO(0).Enabled Then txtPatiMCNO(0).SetFocus
            Exit Sub
        End If
    End If
    If mlng���� > 0 And strMCAccount = "" Then
        strMCAccount = mstrҽ����
    End If
    
    If txt����ʱ�� = "__:__" Then
        str�������� = IIf(IsDate(txt��������.Text), "TO_Date('" & txt��������.Text & "','YYYY-MM-DD')", "NULL")
    Else
        str�������� = IIf(IsDate(txt��������.Text), "TO_Date('" & txt��������.Text & " " & txt����ʱ��.Text & "','YYYY-MM-DD HH24:MI:SS')", "NULL")
    End If

    str���� = Trim(IIf(IsNumeric(txt����.Text), txt����.Text & cbo���䵥λ.Text, txt����.Text))
    
    If Me.Caption Like "�����Һ�*" Then
        '����ԭ�Һŵ����ϵĲ���IDΪ�µ�ID
        mlng����ID = zlDatabase.GetNextNo(1)
        mstr����_���� = txtPatient.Text
        mstr����_�Ա� = NeedName(cbo�Ա�.Text)
        mstr����_���� = txt����.Text & IIf(cbo���䵥λ.Visible, cbo���䵥λ, "")
    End If
    
    '���²�����Ϣ����ֵ
    If Not mblnҽ��ҵ�� Then
        mstr����_���� = txtPatient.Text
        mstr����_�Ա� = NeedName(cbo�Ա�.Text)
        mstr����_���� = txt����.Text & IIf(cbo���䵥λ.Visible, cbo���䵥λ, "")
    End If
    '73935,Ƚ����,20114-7-3,���������ƵĽ���Ƕ�뵽������Ϣ�༭��
    If Not mobjPlugIn Is Nothing And mlngPlugInHwnd <> 0 Then '������������Ϣǰ��������Ч�Լ��
        On Error Resume Next
        blnPlugInCheck = mobjPlugIn.PatiInfoSaveBefore(mlng����ID)
        Call zlPlugInErrH(Err, "PatiInfoSaveBefore")
        If Err = 0 And blnPlugInCheck = False Then
            Exit Sub '���δͨ����ֹ����
        End If
        Err.Clear: On Error GoTo errH
    End If
    
    strDate = "To_Date('" & Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS") & "','YYYY-MM-DD HH24:MI:SS')"
    Set cllPro = New Collection
    '-----------------------------
    mstrSQL = "zl_�ҺŲ��˲���_INSERT(" & mbytType & "," & mlng����ID & "," & txt�����.Text & "," & _
              "'" & m���￨�� & "','" & m��֤���� & "'," & _
              "'" & mstr����_���� & "','" & mstr����_�Ա� & "','" & mstr����_���� & "'," & _
              "'" & NeedName(cbo�ѱ�.Text) & "','" & NeedName(cbo���ʽ.Text) & "'," & _
              "'" & NeedName(cbo����.Text) & "','" & NeedName(cbo����.Text) & "','" & NeedName(cbo����.Text) & "'," & _
              "'" & NeedName(cboְҵ.Text, True) & "','" & txt���֤��.Text & "','" & txt��λ����.Text & "'," & _
              Val(txt��λ����.Tag) & ",'" & txt��λ�绰.Text & "','" & txt��λ�ʱ�.Text & "'," & _
              "'" & IIf(mblnStructAdress, padd��ͥ��ַ.Value, cbo��ͥ��ַ.Text) & "'," & _
              "'" & txt��ͥ�绰.Text & "','" & txt��ͥ�ʱ�.Text & "'," & strDate & _
              ",'" & mstrNO & "'," & str�������� & ",'" & strMCAccount & "',Null," & IIf(mlng���� = 0, "NULL", mlng����) & ","
    '            ����_In           ������Ϣ.����%Type := Null,
    mstrSQL = mstrSQL & "Null,"
    '  ���ڵ�ַ_In       ������Ϣ.���ڵ�ַ%Type := Null,
    mstrSQL = mstrSQL & "Null,"
    '  ���ڵ�ַ�ʱ�_In   ������Ϣ.���ڵ�ַ�ʱ�%Type := Null,
    mstrSQL = mstrSQL & "Null,"
    '  ��ϵ�����֤��_In In ������Ϣ.��ϵ�����֤��%Type := Null,
    mstrSQL = mstrSQL & "Null,"
    '  ��ϵ������_In     In ������Ϣ.��ϵ������%Type := Null,
    mstrSQL = mstrSQL & "Null,"
    '  ��ϵ�˵绰_In     In ������Ϣ.��ϵ�˵绰%Type := Null,
    mstrSQL = mstrSQL & "Null,"
    '  ��ϵ�˹�ϵ_In     In ������Ϣ.��ϵ�˹�ϵ%Type := Null,
    mstrSQL = mstrSQL & "Null,"
    '  �໤��_In         In ������Ϣ.�໤��%Type := Null
    mstrSQL = mstrSQL & "'" & Trim(txtEdit(idx_�໤��).Text) & "'" & ")"
    zlAddArray cllPro, mstrSQL
    
    '89242:���ϴ�,2015/12/7,���²��˵�ַ��Ϣ
    If mblnStructAdress Then
        If padd��ͥ��ַ.Value <> "" Then
           mstrSQL = "zl_���˵�ַ��Ϣ_update(1," & mlng����ID & ",NULL,3,'" & padd��ͥ��ַ.valueʡ & "','" & _
               padd��ͥ��ַ.value�� & "','" & padd��ͥ��ַ.value���� & "','" & padd��ͥ��ַ.value���� & "','" & _
               padd��ͥ��ַ.value��ϸ��ַ & "','" & padd��ͥ��ַ.Code & "')"
        Else
           mstrSQL = "zl_���˵�ַ��Ϣ_update(2," & mlng����ID & ",NULL,3)"
        End If
        zlAddArray cllPro, mstrSQL
    End If
    
    mstrSQL = "ZL_�Һŷ�����Ϣ_Update('" & mstrNO & "'," & mlng����ID & "," & txt�����.Text & "," & _
              "'" & txtPatient.Text & "','" & NeedName(cbo�Ա�.Text) & "','" & str���� & "'," & _
              "'" & NeedName(cbo�ѱ�.Text) & "')"
    zlAddArray cllPro, mstrSQL
    If mlngOutModeMC > 0 And cboҽ�����.ListIndex > 0 Then
        If IsDate(cboҽ�����.Tag) Then strDate = "To_Date('" & cboҽ�����.Tag & "','YYYY-MM-DD HH24:MI:SS')"
        str���� = cboҽ�����.Text
        str���� = Mid(str����, 1, InStr(1, str����, "-") - 1)
        mstrSQL = "zl_����ǼǼ�¼_UPDATE(" & mlngOutModeMC & "," & mlng����ID & ",0," & strDate & ",0,'" & str���� & "')"
        zlAddArray cllPro, mstrSQL
    End If
    '67070:������,2013-11-04,��ȡд�벡��������¼��SQL
    mstrSQL = UCPatiVitalSigns.GetSaveSQL(mlng����ID, mlng�Һ�ID)
    If mstrSQL <> "" Then zlAddArray cllPro, mstrSQL
    
    'ִ�д洢����
    On Error GoTo errSQL
    zlExecuteProcedureArrAy cllPro, Me.Caption
    
    '73935,Ƚ����,20114-7-3,���������ƵĽ���Ƕ�뵽������Ϣ�༭��
    If Not mobjPlugIn Is Nothing And mlngPlugInHwnd <> 0 Then  '������������Ϣ
        On Error Resume Next
        Call mobjPlugIn.PatiInfoSaveAfter(mlng����ID)
        Call zlPlugInErrH(Err, "PatiInfoSaveAfter")
        Err.Clear: On Error GoTo 0
    End If
    
    'ֻ����ȷ������ˢ��
    gblnOk = True
    Unload Me
    Exit Sub
errSQL:
    gcnOracle.RollbackTrans
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
 
 
Private Sub cmd��λ����_Click()
    Call SearchUnit("", txt��λ����)
End Sub


Private Sub Form_Activate()
    '78408:���ϴ�,2014/10/9,�����ת
    If Me.ActiveControl Is msh���� Then Exit Sub
    If txtPatient.Enabled Then txtPatient.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If txt����.Visible Then
            txt����.Visible = False
            msh����_EnterCell
            msh����.SetFocus
        ElseIf lvwItems.Visible Then
            lvwItems.Visible = False
            txt����.Visible = True
            txt����.SetFocus
        Else
            cmdCancel_Click
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        '89242:���ϴ�,2015/12/7,PatiAddress�ؼ��ڲ���������ת���ⲿ���ٴ���
        If UCase(TypeName(Me.ActiveControl)) = UCase("PatiAddress") Then Exit Sub
        If InStr(1, "lvwItems,txt����,cbo���䵥λ,txt��������,msh����,txt����,txtPatiMCNO", Me.ActiveControl.Name) <= 0 Then
            KeyAscii = 0
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub Form_Resize()
    If tabPage.Visible Then
        tabPage.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - 500
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '73935,Ƚ����,20114-7-3,���������ƵĽ���Ƕ�뵽������Ϣ�༭��
    If Not mobjPlugIn Is Nothing Then Set mobjPlugIn = Nothing
    mlngPlugInHwnd = 0
    
    If Not mrs��ͥ��ַ Is Nothing Then
        If mrs��ͥ��ַ.State = 1 Then
            On Error Resume Next
            Kill App.Path & "\ZLAddressForRegEvent.Adtg"
            Err.Clear
            mrs��ͥ��ַ.Filter = ""
            mrs��ͥ��ַ.Save App.Path & "\ZLAddressForRegEvent.Adtg"
        End If
    End If
    Set mrs��ͥ��ַ = Nothing
    
    mlng����ID = 0
    mstrNO = ""
    mlng���� = 0
    mstrҽ���� = ""
    m���￨�� = ""
    m��֤���� = ""
    mbytType = 0
    mblnChange = False
    mblnGetBirth = False
End Sub


Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItems.SortKey = ColumnHeader.index - 1 Then
        Me.lvwItems.SortOrder = IIf(Me.lvwItems.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItems.SortKey = ColumnHeader.index - 1
        Me.lvwItems.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwItems_DblClick()
    Dim i As Integer
    
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    
    For i = 1 To msh����.Rows - 1
        If i <> msh����.Row Then
            If msh����.RowData(i) = Replace(lvwItems.SelectedItem.Key, "_", "") Then
                MsgBox "�� " & i & " �е�ҩ���Ѿ�����ѡ���ҩ��������ͬ,������ѡ��", vbInformation, gstrSysName
                lvwItems.SetFocus
                Exit Sub
            End If
        End If
    Next
    lvwItems.Visible = False
    msh����.RowData(msh����.Row) = Replace(lvwItems.SelectedItem.Key, "_", "")
    msh����.Text = Trim(lvwItems.SelectedItem.Text)
    msh����.SetFocus
    msh����_EnterCell
    
End Sub

Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
        Call lvwItems_DblClick
    Case vbKeyEscape
        lvwItems.Visible = False
        txt����.Visible = True
        txt����.SetFocus
    End Select
End Sub

Private Sub lvwItems_LostFocus()
    Me.lvwItems.Visible = False
End Sub

Private Sub msh����_Click()
    msh����_EnterCell
End Sub

Private Sub msh����_GotFocus()
    msh����_EnterCell
End Sub

Private Sub msh����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    
    If KeyCode = vbKeyF2 Then msh����_DblClick
    If KeyCode = vbKeyF3 And cmd����.Visible Then cmd����_Click
    If KeyCode = vbKeyDelete Then
        msh����.Text = ""
        msh����.RowData(msh����.Row) = 0
        For i = msh����.Row + 1 To msh����.Rows - 1
            msh����.TextMatrix(i - 1, 0) = msh����.TextMatrix(i, 0)
            msh����.RowData(i - 1) = msh����.RowData(i)
            msh����.TextMatrix(i, 0) = ""
            msh����.RowData(i) = 0
        Next
        msh����_EnterCell
    End If
End Sub

Private Sub msh����_DblClick()
    cmd����.Visible = False
    txt����.Visible = False
    
    'If msh����.Row > 1 And msh����.TextMatrix(msh����.Row - 1, 0) = "" Or msh����.RowData(msh����.Row) = 0 Then Exit Sub
    
    txt����.Top = msh����.CellTop + msh����.Top + (msh����.CellHeight - txt����.Height) / 2 - 15
    txt����.Left = msh����.Left + msh����.CellLeft + 30
    txt����.Width = msh����.CellWidth - 60
    
    txt����.Text = msh����.Text
    txt����.ZOrder
    SelAll txt����
    txt����.Visible = True
    If txt����.Visible And txt����.Enabled Then txt����.SetFocus
End Sub

Private Sub msh����_EnterCell()
    cmd����.Visible = False
    txt����.Visible = False

    If msh����.Row > 1 And msh����.TextMatrix(msh����.Row - 1, 0) = "" Then Exit Sub
    
    cmd����.Top = msh����.CellTop + msh����.Top - 15
    If msh����.Rows < 5 Then
        cmd����.Left = msh����.Left + msh����.CellWidth - cmd����.Width + 45
    Else
        cmd����.Left = msh����.Left + msh����.CellWidth - cmd����.Width + 45
    End If
    
    cmd����.ZOrder
    cmd����.Visible = True
End Sub

Private Sub msh����_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then
        'If msh����.Row > 1 And msh����.TextMatrix(msh����.Row - 1, 0) = "" Or msh����.RowData(msh����.Row) = 0 Then Exit Sub
        msh����_DblClick
        txt����.Text = Chr(KeyAscii)
        txt����.SelStart = Len(txt����.Text)
    Else
        If msh����.Row = msh����.Rows - 1 And msh����.TextMatrix(msh����.Row, msh����.Col) <> "" Then
            msh����.Rows = msh����.Rows + 1
            msh����.Row = msh����.Rows - 1
            
            msh����_EnterCell
        ElseIf msh����.TextMatrix(msh����.Row, msh����.Col) <> "" Then
            msh����.Row = msh����.Row + 1
            msh����_EnterCell
        Else
            cmdOK.SetFocus
        End If
    End If
End Sub

Private Sub cmd����_Click()
On Error GoTo errH
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    Dim i As Integer
    
    strSQL = _
        " Select -1 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'����ҩ' as ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as ��ҩ,NULL as Ƥ�� From Dual Union ALL" & _
        " Select -2 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'�г�ҩ' as ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as ��ҩ,NULL as Ƥ�� From Dual Union ALL" & _
        " Select -3 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'�в�ҩ' as ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as ��ҩ,NULL as Ƥ�� From Dual Union ALL" & _
        " Select ID,nvl(�ϼ�ID,-����) as �ϼ�ID,0 as ĩ��,NULL as ����,����," & _
        " NULL as ��λ,NULL as ����,NULL as �������,NULL as ��ҩ,NULL as Ƥ��" & _
        " From ���Ʒ���Ŀ¼ Where ���� IN (1,2,3) And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null)" & _
        " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID" & _
        " Union All" & _
        " Select Distinct A.ID,A.����ID as �ϼ�ID,1 as ĩ��,A.����," & _
        " A.����,A.���㵥λ as ��λ,B.ҩƷ���� as ����,B.�������," & _
        " Decode(B.�Ƿ���ҩ,1,'��','') as ��ҩ,Decode(B.�Ƿ�Ƥ��,1,'��','') as Ƥ��" & _
        " From ������ĿĿ¼ A,ҩƷ���� B" & _
        " Where A.��� IN('5','6','7') And A.ID=B.ҩ��ID" & _
        " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)"
    
    Set rsTmp = frmPubSel.ShowSelect(Me, strSQL, 2, "����ҩ��", , msh����.Text, "��������ҩƷ��ѡ��һ����Ϊ���˹���ҩ�")
    If Not rsTmp Is Nothing Then
        For i = 1 To msh����.Rows - 1
            If i <> msh����.Row Then
                If msh����.RowData(i) = rsTmp!ID Then
                    MsgBox "�� " & i & " �е�ҩ���Ѿ�����ѡ���ҩ��������ͬ,������ѡ��", vbInformation, gstrSysName
                    msh����.SetFocus
                    msh����_EnterCell
                    Exit Sub
                End If
            End If
        Next
        msh����.RowData(msh����.Row) = rsTmp!ID
        msh����.Text = Trim(rsTmp!����)
    End If
    msh����.SetFocus
    msh����_EnterCell
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmd��ͥ��ַ_Click()
On Error GoTo errH
    Dim rsTmp As ADODB.Recordset
    Set rsTmp = frmPubSel.ShowSelect(Me, _
            " Select Distinct Substr(����,1,2) as ID,NULL as �ϼ�ID,0 as ĩ��,NULL as ����," & _
            " Substr(����,1,2) as ���� From ����" & _
            " Union All" & _
            " Select ���� as ID,Substr(����,1,2) as �ϼ�ID,1 as ĩ��,����,���� " & _
            " From ���� Order by ����", 2, "����", , cbo��ͥ��ַ.Text)
    If Not rsTmp Is Nothing Then
        cbo��ͥ��ַ.Text = rsTmp!����
        cbo��ͥ��ַ.SelStart = Len(cbo��ͥ��ַ.Text)
    End If
    cbo��ͥ��ַ.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function InitData() As Boolean
On Error GoTo errH
'���ܣ���ʼ����Ҫ����
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer
    Dim objCtl As Control
    Dim strSQL As String, lngTmp As Long
    
    If mlng����ID = 0 Then
        Me.Caption = "�����ҺŲ�����Ϣ"
    End If
    Me.txt�����.Enabled = False
    
    '�ѱ�
    strSQL = "Select ����,����,����,Nvl(ȱʡ��־,0) as ȱʡ" & vbCrLf & _
        " From �ѱ� Where ����=1 And Nvl(�������,3) IN(1,3)" & vbCrLf & _
        " Order by ����"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        For i = 1 To rsTmp.RecordCount
            cbo�ѱ�.AddItem rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ȱʡ = 1 Then
                cbo�ѱ�.ItemData(cbo�ѱ�.NewIndex) = 1
                cbo�ѱ�.ListIndex = cbo�ѱ�.NewIndex
            End If
            rsTmp.MoveNext
        Next
        cbo�ѱ�.Enabled = False
    End If
    
    
    If mlngOutModeMC > 0 Then
        Set rsTmp = GetDictData("ҽ�����")
        cboҽ�����.Clear
        cboҽ�����.AddItem " "
        For i = 1 To rsTmp.RecordCount
            cboҽ�����.AddItem rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ȱʡ = 1 Then
                cboҽ�����.ItemData(cboҽ�����.NewIndex) = 1
            End If
            rsTmp.MoveNext
        Next
        cboҽ�����.ListIndex = 0
        Call zlControl.CboSetWidth(cbo���ʽ.Hwnd, txtPatiMCNO(0).Width)
        
    Else
        lblPatiMCNO(0).Visible = False: lblPatiMCNO(1).Visible = False
        txtPatiMCNO(0).Visible = False: txtPatiMCNO(1).Visible = False
        lblҽ�����.Visible = False: cboҽ�����.Visible = False
        
        
        lngTmp = txtPatiMCNO(0).Height / 2 + 100
        
        cbo���ʽ.Width = txtPatiMCNO(0).Width
        lbl���ʽ.Top = 1910 - lngTmp
        cbo���ʽ.Top = 1850 - lngTmp
        lbl�ѱ�.Top = 1910 - lngTmp
        cbo�ѱ�.Top = 1850 - lngTmp
        Line1.Y1 = 2340 - lngTmp + 100
        Line1.Y2 = 2340 - lngTmp + 100
        Line2.Y1 = 2325 - lngTmp + 100
        Line2.Y2 = 2325 - lngTmp + 100
    End If
    
    '�Ա�
    Set rsTmp = GetDictData("�Ա�")
    cbo�Ա�.Clear
    If Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
            cbo�Ա�.AddItem rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ȱʡ = 1 Then
                cbo�Ա�.ItemData(cbo�Ա�.NewIndex) = 1
                cbo�Ա�.ListIndex = cbo�Ա�.NewIndex
            End If
            rsTmp.MoveNext
        Next
    End If
    
    '���䵥λ
    cbo���䵥λ.Clear
    cbo���䵥λ.AddItem "��"
    cbo���䵥λ.AddItem "��"
    cbo���䵥λ.AddItem "��"
    cbo���䵥λ.ListIndex = 0
    mDateSys = zlDatabase.Currentdate
    
    If Not mblnStructAdress Then Call Load��ͥ��ַ

    'ҽ�Ƹ��ʽ
    Set rsTmp = GetDictData("ҽ�Ƹ��ʽ")
    cbo���ʽ.Clear
    If Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
            cbo���ʽ.AddItem rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ȱʡ = 1 Then
                cbo���ʽ.ItemData(cbo���ʽ.NewIndex) = 1
                cbo���ʽ.ListIndex = cbo���ʽ.NewIndex
            End If
            rsTmp.MoveNext
        Next
    End If
    cbo���ʽ.Enabled = False

    '����
    Set rsTmp = GetDictData("����")
    cbo����.Clear
    If Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
            cbo����.AddItem rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ȱʡ = 1 Then
                cbo����.ItemData(cbo����.NewIndex) = 1
                cbo����.ListIndex = cbo����.NewIndex
            End If
            rsTmp.MoveNext
        Next
    End If

    '����
    Set rsTmp = GetDictData("����")
    cbo����.Clear
    If Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
            cbo����.AddItem rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ȱʡ = 1 Then
                cbo����.ItemData(cbo����.NewIndex) = 1
                cbo����.ListIndex = cbo����.NewIndex
            End If
            rsTmp.MoveNext
        Next
    End If

    '����״��
    Set rsTmp = GetDictData("����״��")
    cbo����.Clear
    If Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
            cbo����.AddItem rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ȱʡ = 1 Then
                cbo����.ItemData(cbo����.NewIndex) = 1
                cbo����.ListIndex = cbo����.NewIndex
            End If
            rsTmp.MoveNext
        Next
    End If

    'ְҵ
    Set rsTmp = GetDictData("ְҵ")
    cboְҵ.Clear
    If Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
            cboְҵ.AddItem rsTmp!���� & "-" & rsTmp!����
            If rsTmp!ȱʡ = 1 Then
                cboְҵ.ItemData(cboְҵ.NewIndex) = 1
                cboְҵ.ListIndex = cboְҵ.NewIndex
            End If
            rsTmp.MoveNext
        Next
    End If
     
    If mlng����ID = 0 Then
        Me.cbo���ʽ.Enabled = True
        Me.cbo�ѱ�.Enabled = True
    End If
    Call SetPatiBaseInforEnabled
    
    InitData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Sub SetPatiBaseInforEnabled()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ò��˵Ļ�����Ϣ(����,�Ա�,����,��������)��Eanbeld
    '���:
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-11-11 10:40:42
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnEdit As Boolean
    Dim lngColor As Long
    
    blnEdit = mlng����ID = 0
    If mlng�Һ�ID <> 0 Then
        '������ҽ��ҵ������,���ܵ������˻�����Ϣ
        blnEdit = Not mblnҽ��ҵ��
        'Not zlExistOperationData(mlng����ID, mstrNO, mlng�Һ�ID)
    End If
    lngColor = IIf(blnEdit = False, Me.BackColor, Me.txt�����.BackColor)
    
    txtPatient.Enabled = blnEdit
    cbo�Ա�.Enabled = blnEdit
    txt����.Enabled = blnEdit
    txt��������.Enabled = blnEdit
    txt����ʱ��.Enabled = blnEdit
    txtPatient.BackColor = lngColor
    cbo�Ա�.BackColor = lngColor
    txt����.BackColor = lngColor
    txt��������.BackColor = lngColor
    txt����ʱ��.BackColor = lngColor
End Sub



Private Sub msh����_Scroll()
    cmd����.Visible = False
End Sub

 

Private Sub picAddInfo_Resize()
    wndTaskPanel.Move 0, 0, picAddInfo.Width, picAddInfo.Height
End Sub

Private Sub txtEdit_GotFocus(index As Integer)
    Call zlControl.TxtSelAll(txtEdit(index))
End Sub

Private Sub txtEdit_KeyPress(index As Integer, KeyAscii As Integer)
        Dim strMask As String
        If KeyAscii = 8 Then Exit Sub
        If KeyAscii = 13 Then zlCommFun.PressKey vbKeyTab: Exit Sub
        Select Case index
            Case idx_���, idx_����, idx_����
                strMask = "1234567890."
            Case Else
                strMask = ""
        End Select
        If strMask <> "" Then
            If InStr(strMask, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0: Exit Sub
            End If
        End If
End Sub

Private Sub txtPatiMCNO_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtPatiMCNO_Validate(index As Integer, Cancel As Boolean)
    
    txtPatiMCNO(index).Text = Trim(txtPatiMCNO(index).Text)
    If cbo���ʽ.ListCount > 0 Then cbo���ʽ.ListIndex = 0

    If index = 1 Then
        If txtPatiMCNO(1).Text <> txtPatiMCNO(0).Text Then
            MsgBox "����,���������ҽ���Ų�һ�£�", vbInformation, gstrSysName
            Cancel = True
            Exit Sub
        End If
    End If
    
    If mlngOutModeMC = 920 And txtPatiMCNO(0).Text <> txtPatiMCNO(0).Tag And txtPatiMCNO(0).Text <> "" Then
        If CheckExistsMCNO(txtPatiMCNO(0).Text) Then
            Cancel = True
        End If
    End If
End Sub

Private Sub txt����ʱ��_Change()
    Dim str����ʱ�� As String
    '76669�����ϴ�,2014-8-18,�����������
    If IsDate(txt��������.Text) Then
        str����ʱ�� = txt��������.Text & IIf(IsDate(txt����ʱ��.Text), " " & txt����ʱ��.Text, "")
        txt����.Text = ReCalcOld(CDate(str����ʱ��), cbo���䵥λ, mlng����ID)
    End If
End Sub

Private Sub txt����ʱ��_GotFocus()
    zlControl.TxtSelAll txt����ʱ��
End Sub

Private Sub txt����ʱ��_KeyPress(KeyAscii As Integer)
    If Not IsDate(txt��������.Text) Then
        KeyAscii = 0
        txt����ʱ��.Text = "__:__"
    End If
End Sub


Private Sub txt����ʱ��_Validate(Cancel As Boolean)
    If txt����ʱ��.Text <> "__:__" And Not IsDate(txt����ʱ��.Text) Then
        txt����ʱ��.SetFocus
        Cancel = True
    End If
End Sub

Private Sub txt��������_Change()
    Dim str����ʱ�� As String
    If IsDate(txt��������.Text) And mblnChange Then
        mblnChange = False
        txt��������.Text = Format(CDate(txt��������.Text), "yyyy-mm-dd") '0002-02-02�Զ�ת��Ϊ2002-02-02,����,��������2002,ʵ��ֵȴ��0002
        mblnChange = True
        
        str����ʱ�� = txt��������.Text & IIf(IsDate(txt����ʱ��.Text), " " & txt����ʱ��.Text, "")
        txt����.Text = ReCalcOld(CDate(str����ʱ��), cbo���䵥λ, mlng����ID)
    End If
End Sub
Private Sub txt��������_GotFocus()
    zlControl.TxtSelAll txt��������
End Sub
Private Sub txt��������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
       KeyAscii = 0
       If txt��������.Text = "____-__-__" Then
              zlCommFun.PressKey (vbKeyTab) '����ʱ��
              zlCommFun.PressKey (vbKeyTab)
          Else
              zlCommFun.PressKey (vbKeyTab)
          End If
    End If
End Sub

Private Sub txt��������_LostFocus()
    If txt��������.Text <> "____-__-__" And Not IsDate(txt��������.Text) Then
        txt��������.SetFocus
    End If
End Sub

Private Sub txt��λ�绰_GotFocus()
    SelAll txt��λ�绰
End Sub

Private Sub txt��λ�绰_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckLen txt��λ�绰, KeyAscii
End Sub

Private Sub txt��λ����_Change()
    txt��λ����.Tag = ""
End Sub

Private Sub txt��λ����_GotFocus()
    SelAll txt��λ����
    OpenIme gstrIme
End Sub

Private Sub txt��λ����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 And cmd��λ����.Enabled And cmd��λ����.Visible Then cmd��λ����_Click
    '73022,Ƚ����,2014-5-20,�ڵ�λ���ơ������ص㡢���ڵ�ַ����ģ�����ҹ���
    If KeyCode = vbKeyReturn And Trim(txt��λ����.Text) <> "" Then
        Call SearchUnit(Trim(txt��λ����.Text), txt��λ����)
    End If
End Sub

Private Sub SearchUnit(ByVal strInput As String, txtInput As Object)
    '--------------------------------------------------------------
    '����:ģ�����ң�������Լ��λѡ���б�
    '����:Ƚ����
    '����:2014-5-23
    '����:
    '   strInput:�����ı�����Ϊ�ձ�ʾ�����ť����
    '   txtInput:�ı������
    '--------------------------------------------------------------
    Dim strSQL As String, strWhere As String
    Dim strKey As String, blnCancel As Boolean
    Dim rsTemp As ADODB.Recordset, vRect As RECT
    
    On Error GoTo Errhand
    If strInput <> "" And txtInput.Tag <> "" Then Exit Sub
    vRect = GetControlRect(txtInput.Hwnd)
    If strInput = "" Then '�����ť
        strSQL = "" & _
        "       Select ID,�ϼ�ID,ĩ��,����,����,��ַ,�绰,��������,�ʺ�,��ϵ�� From  ��Լ��λ" & _
        "       Where ����ʱ�� Is Null Or ����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD')" & _
        "       Start With �ϼ�ID is NULL" & _
        "       Connect by Prior ID=�ϼ�ID"
        '75888,Ƚ����,2014-7-28
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "��λ", False, _
                       "", "", False, True, False, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, False)
    Else
        'ȥ��"'"
        strInput = Replace(strInput, "'", " ")
        strKey = GetMatchingSting(strInput, False)
        If strInput <> "" Then
            If IsNumeric(strInput) Then '����ȫ������ʱֻƥ�����
                strWhere = " Where ���� Like Upper([1])"
            ElseIf zlCommFun.IsCharAlpha(strInput) Then '����ȫ����ĸʱֻƥ�����
                strWhere = " Where ���� Like Upper([1])"
            Else
                strWhere = " Where ���� Like Upper([1]) Or ���� Like [1] Or ���� Like Upper([1])"
            End If
        End If
        
        strSQL = "" & _
        "       Select ID,�ϼ�ID,ĩ��,����,����,��ַ,�绰,��������,�ʺ�,��ϵ�� From  ��Լ��λ" & strWhere & _
        "       And (����ʱ�� Is Null Or ����ʱ�� >= To_Date('3000-01-01', 'YYYY-MM-DD'))"
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "��λ", False, _
                       "", "", False, False, True, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, False, strKey)
    End If
    If blnCancel Then txtInput.SetFocus: Exit Sub

    If rsTemp Is Nothing Then txtInput.SetFocus: Exit Sub
    If rsTemp.State <> 1 Then txtInput.SetFocus: Exit Sub
    
    txtInput.Text = Nvl(rsTemp!����)
    txtInput.Tag = Nvl(rsTemp!ID)
    txtInput.SelStart = Len(Nvl(txtInput.Text))
    txtInput.SetFocus
    
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt��λ����_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckLen txt��λ����, KeyAscii
End Sub

Private Sub txt��λ����_LostFocus()
    If gstrIme <> "���Զ�����" Then Call OpenIme
End Sub

Private Sub txt��λ�ʱ�_GotFocus()
    SelAll txt��λ�ʱ�
End Sub

Private Sub txt��λ�ʱ�_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
    CheckLen txt��λ�ʱ�, KeyAscii
End Sub
Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim ObjItem As ListItem
    Dim strSQL As String
            
    If KeyAscii = 13 Then
        KeyAscii = 0
        
        If msh����.Text = "" Then
            '1.������������,��ʾ���ҽ���������б�
            strSQL = " Select Distinct A.ID,A.����," & _
            " A.����,A.���㵥λ as ��λ,B.ҩƷ���� as ����,B.�������," & _
            " Decode(B.�Ƿ���ҩ,1,'��','') as ��ҩ,Decode(B.�Ƿ�Ƥ��,1,'��','') as Ƥ��" & _
            " From ������ĿĿ¼ A,ҩƷ���� B,������Ŀ���� C" & _
            " Where A.��� IN('5','6','7') And A.ID=B.ҩ��ID And A.Id=C.������Ŀid" & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
            " And (C.���� like [1] OR A.���� like [1] OR C.���� like [1])"
            
            On Error GoTo errH
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, gstrLike & UCase(txt����.Text) & "%")
            
            With rsTmp
                If .BOF Or .EOF Then Exit Sub
                Me.lvwItems.ListItems.Clear
                Do While Not .EOF
                    Set ObjItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !����, , IIf(!Ƥ�� <> "", 1, 2))
                    ObjItem.SubItems(Me.lvwItems.ColumnHeaders("����").index - 1) = !����
                    ObjItem.SubItems(Me.lvwItems.ColumnHeaders("��λ").index - 1) = IIf(IsNull(!��λ), "", !��λ)
                    ObjItem.SubItems(Me.lvwItems.ColumnHeaders("����").index - 1) = IIf(IsNull(!����), "", !����)
                    ObjItem.SubItems(Me.lvwItems.ColumnHeaders("�������").index - 1) = IIf(IsNull(!�������), "", !�������)
                    ObjItem.SubItems(Me.lvwItems.ColumnHeaders("��ҩ").index - 1) = IIf(IsNull(!��ҩ), "", !��ҩ)
                    ObjItem.SubItems(Me.lvwItems.ColumnHeaders("Ƥ��").index - 1) = IIf(IsNull(!Ƥ��), "", !Ƥ��)
                    .MoveNext
                Loop
                Me.lvwItems.ListItems(1).Selected = True
            End With
            
            With Me.lvwItems
                .Left = msh����.Left
                .Width = msh����.Width
                .Height = msh����.Height + 300
                If msh����.Rows < 5 Then
                    .Top = msh����.Top + msh����.RowHeight(msh����.Row) * (msh����.Row) - .Height
                Else
                    .Top = msh����.Top + msh����.RowHeight(4) * (3) - .Height
                End If
                .ZOrder 0: .Visible = True
                .SetFocus
            End With
        Else
            '2.�����ǰ��Ԫ����������,����Ϊ�༭����
                    msh����.Text = txt����.Text
            If msh����.Row + 1 <= msh����.Rows - 1 Then msh����.Row = msh����.Row + 1
            msh����.SetFocus
            msh����_EnterCell
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
        
    mblnChange = True
    txtPatient.MaxLength = zlGetPatiInforMaxLen.intPatiName
    
    '����ҩ����ҽ���б��ʼ��
    Me.lvwItems.ListItems.Clear
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "����", "����", 1400, 0
        .Add , "����", "����", 900
        .Add , "��λ", "��λ", 600
        .Add , "����", "����", 600
        .Add , "�������", "�������", 900
        .Add , "��ҩ", "��ҩ", 600
        .Add , "Ƥ��", "Ƥ��", 600
    End With
    
    With Me.lvwItems
        .ColumnHeaders("����").Position = 1
        .SortKey = .ColumnHeaders("����").index - 1
        .SortOrder = lvwAscending
        .Visible = False
    End With
    
    '73935,Ƚ����,20114-7-3,���������ƵĽ���Ƕ�뵽������Ϣ�༭��
    Call CreateObjectPlugIn
    If Not mobjPlugIn Is Nothing Then
        On Error Resume Next
        mlngPlugInHwnd = mobjPlugIn.GetFormHwnd
        Call zlPlugInErrH(Err, "GetFormHwnd")
        Err.Clear: On Error GoTo 0
        If mlngPlugInHwnd <> 0 Then
            tabPage.Visible = True: Me.Height = Me.Height + 350
            cmdOK.Top = cmdOK.Top + 330: cmdCancel.Top = cmdOK.Top
            Call InitTagPage
            Call InitTaskPanel
        End If
    End If
    
    '����������Ϣ��������
    '69026,Ƚ����,2014-8-8,�����������
    Call CreatePublicPatient
    
    '��ʼ����ַ�ؼ�
    If Not mblnStructAdress Then Exit Sub
    padd��ͥ��ַ.Visible = mblnStructAdress
    padd��ͥ��ַ.ShowTown = mblnShowTown
    cbo��ͥ��ַ.Visible = False: cmd��ͥ��ַ.Visible = False
    padd��ͥ��ַ.Top = cbo��ͥ��ַ.Top: padd��ͥ��ַ.Left = cbo��ͥ��ַ.Left
End Sub


Private Sub txt��ͥ�绰_GotFocus()
    SelAll txt��ͥ�绰
End Sub

Private Sub txt��ͥ�绰_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckLen txt��ͥ�绰, KeyAscii
End Sub

Private Sub txt��ͥ�ʱ�_GotFocus()
    SelAll txt��ͥ�ʱ�
End Sub

Private Sub txt��ͥ�ʱ�_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
    CheckLen txt��ͥ�ʱ�, KeyAscii
End Sub

Private Sub txt�����_GotFocus()
    SelAll txt�����
End Sub

Private Sub txt�����_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
    CheckLen txt�����, KeyAscii
End Sub


Private Sub txt����_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txt����.Hwnd, GWL_WNDPROC)
        Call SetWindowLong(txt����.Hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt����_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txt����.Hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt����_GotFocus()
    Call zlCommFun.OpenIme
    SelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Dim blnTab As Boolean
    
    If KeyAscii = vbKeyReturn Then
        If cbo���䵥λ.Visible = False And IsNumeric(txt����.Text) Then
            Call txt����_Validate(False)
            Call cbo���䵥λ.SetFocus
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
        If Not IsNumeric(txt����.Text) And cbo���䵥λ.Visible Then Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr(KeyAscii))) > 0 Then KeyAscii = 0
    End If
End Sub



Private Sub txt����_Validate(Cancel As Boolean)
    txt����.Text = Trim(txt����.Text)
    
    If Not IsNumeric(txt����.Text) And Trim(txt����.Text) <> "" Then
        cbo���䵥λ.ListIndex = -1: cbo���䵥λ.Visible = False
    ElseIf cbo���䵥λ.Visible = False Then
        cbo���䵥λ.ListIndex = 0: cbo���䵥λ.Visible = True
    End If
    If Not IsDate(txt��������.Text) Then mblnGetBirth = True
    mblnChange = False
    If mblnGetBirth Then
        txt��������.Text = ReCalcBirth(Trim(txt����.Text), "")
    End If
    mblnChange = True
    '69026,Ƚ����,2014-8-8,�����������
    '76703,Ƚ����,2014-8-15
    If cbo���䵥λ.Visible Then Exit Sub
    If mobjPubPatient Is Nothing Then Exit Sub
    If mobjPubPatient.CheckPatiAge(Trim(txt����.Text), _
            IIf(txt��������.Text = "____-__-__", "", txt��������.Text) & _
            IIf(txt����ʱ��.Text = "__:__", "", " " & txt����ʱ��.Text)) = False Then
        Cancel = True
    End If
End Sub

Private Sub txt���֤��_GotFocus()
    SelAll txt���֤��
End Sub

Private Sub txt���֤��_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckLen txt���֤��, KeyAscii
End Sub

Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtPatient.Hwnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.Hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.Hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txtPatient_GotFocus()
    SelAll txtPatient
    OpenIme gstrIme
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    CheckLen txtPatient, KeyAscii
End Sub

Public Sub ClearFace()
    Dim i As Integer
    
    txt�����.Text = ""
    SetCboDefault cbo�ѱ�
    SetCboDefault cbo�Ա�
    If mlngOutModeMC > 0 Then
        txtPatiMCNO(0).Text = ""
        txtPatiMCNO(0).Tag = "" '�����޸�ʱ�ж��Ƿ��Ѵ���
        txtPatiMCNO(1).Text = ""
        If cboҽ�����.ListIndex >= 0 Then cboҽ�����.ListIndex = 0
    End If
    
    txtPatient.Text = ""
    txt����.Text = ""
    Call zlControl.CboLocate(cbo���䵥λ, "��")
    
    SetCboDefault cbo���ʽ
    SetCboDefault cbo����
    SetCboDefault cbo����
    SetCboDefault cbo����
    SetCboDefault cboְҵ
    
    txt���֤��.Text = ""
    
    txt��λ����.Text = ""
    txt��λ����.Tag = ""
    txt��λ�绰.Text = ""
    txt��λ�ʱ�.Text = ""
    
    cbo��ͥ��ַ.Text = ""
    padd��ͥ��ַ.Value = ""
    txt��ͥ�ʱ�.Text = ""
    txt��ͥ�绰.Text = ""
    For i = 1 To msh����.Rows - 1
        msh����.TextMatrix(i, 0) = ""
        msh����.RowData(i) = 0
    Next
End Sub

Private Sub txtPatient_LostFocus()
    If gstrIme <> "���Զ�����" Then Call OpenIme
End Sub

Private Function GetDictData(strDict As String) As ADODB.Recordset
'���ܣ���ָ�����ֵ��ж�ȡ����
'������strDict=�ֵ��Ӧ�ı���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
        
    strSQL = "Select ����,����,Nvl(ȱʡ��־,0) as ȱʡ From " & strDict & " Order by ����"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    Set GetDictData = rsTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub Setҽ������()
    Dim i As Integer
    For i = 0 To cbo���ʽ.ListCount - 1
        If Left(cbo���ʽ.List(i), InStr(cbo���ʽ.List(i), "-") - 1) = "1" Then
            cbo���ʽ.ListIndex = i: Exit Sub
        End If
    Next
End Sub

Public Function GetRegBillID() As Boolean
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    
     On Error GoTo Errhand
     If mstrNO = "" Then Exit Function
        '������Ϣ
    strSQL = "Select A.����ID,B.ID as �Һ�ID,B.ժҪ,B.����,a.����,A.��������," & _
        " Nvl(Nvl(B.�������ID,Decode(B.ת��״̬,1,B.ת�����ID,NULL)),B.ִ�в���ID) as ����ID," & _
        " B.��Ⱦ���ϴ�,B.����ʱ��,A.����,A.�����,A.����,A.�Ա�,A.����,A.��������,A.ҽ�Ƹ��ʽ," & _
        " A.����,A.����,A.����״��,A.ְҵ,A.���֤��,A.�����ص�,A.�໤��,A.��ͥ��ַ,A.��ͥ�绰," & _
        " A.����,A.��ͥ��ַ�ʱ�,A.������λ,A.��ͬ��λid,A.��λ�绰,A.��λ�ʱ�,B.����,C.������,A.����֤��,A.���ڵ�ַ,a.���ڵ�ַ�ʱ�" & _
        " From ������Ϣ A,���˹Һż�¼ B,����������Ϣ C" & _
        " Where A.����ID=B.����ID And B.����ID=C.����ID(+) And B.����=C.����(+) And B.NO=[1] And B.��¼����=1 And B.��¼״̬=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrNO)
    If rsTmp.EOF Then Exit Function
    mlng�Һ�ID = Val(Nvl(rsTmp!�Һ�ID))
    '74428�����ϴ���2014-7-8������������ʾ��ɫ����
    Call SetPatiColor(txtPatient, Nvl(rsTmp!��������), IIf(IsNull(rsTmp!����), Me.ForeColor, vbRed))
    Set rsTmp = Nothing
    GetRegBillID = True
Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Sub txt���֤��_Validate(Cancel As Boolean)
    '65663:������,2014-02-20,�������֤�ż����������
    If IsDate(zlCommFun.GetIDCardDate(txt���֤��.Text)) = False Then Exit Sub
    If Format(zlCommFun.GetIDCardDate(txt���֤��.Text), "yyyy-mm-dd") <> Format(txt��������.Text, "yyyy-mm-dd") Then
        MsgBox "��������֤��������ĳ������ڲ�һ�£���ʹ�����֤�Ż�ȡ�������滻��", vbInformation, gstrSysName
        txt��������.Text = zlCommFun.GetIDCardDate(txt���֤��.Text)
    End If
End Sub

Private Function CreateObjectPlugIn() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������������Ϣ���
    '����:�����ɹ�,����True,���򷵻�False
    '�����:73935
    '����:Ƚ����
    '����:2014-07-3
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjPlugIn Is Nothing Then
        On Error Resume Next
        Set mobjPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        Err.Clear: On Error GoTo 0
    End If
    
    If Not mobjPlugIn Is Nothing Then
        On Error Resume Next
        Call mobjPlugIn.Initialize(gcnOracle, glngSys, 1113)
        Call zlPlugInErrH(Err, "Initialize")
        Err.Clear: On Error GoTo 0
    End If
    CreateObjectPlugIn = True
End Function

Public Sub HideFormCaption(ByVal lngHwnd As Long, Optional ByVal blnBorder As Boolean = True)
'���ܣ�����һ������ı�����
'������blnBorder=���ر�������ʱ��,�Ƿ�Ҳ���ش���߿�
    Dim vRect As RECT, lngStyle As Long
    
    Call GetWindowRect(lngHwnd, vRect)
    lngStyle = GetWindowLong(lngHwnd, GWL_STYLE)

    If blnBorder Then
        lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX)
    Else
        lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_THICKFRAME)
    End If
    SetWindowLong lngHwnd, GWL_STYLE, lngStyle
    SetWindowPos lngHwnd, 0, 0, 0, vRect.Right - vRect.Left, vRect.Bottom - vRect.Top, SWP_NOREPOSITION Or SWP_FRAMECHANGED Or SWP_NOZORDER
End Sub

Private Sub InitTagPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ҳ�ؼ�
    '�����:73935
    '����:Ƚ����
    '����:2014-07-4
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, ObjItem As TabControlItem, objForm As Object
    
    Err = 0: On Error GoTo Errhand:

    Set ObjItem = tabPage.InsertItem(mPageIndex.������Ϣ, "������Ϣ", picPatiInfo.Hwnd, 0)
    ObjItem.Tag = mPageIndex.������Ϣ

    If Not mobjPlugIn Is Nothing Then
        If mlngPlugInHwnd <> 0 Then
            picAddInfo.Visible = True
            Set ObjItem = tabPage.InsertItem(mPageIndex.������Ϣ, "������Ϣ", picAddInfo.Hwnd, 0)
            ObjItem.Tag = mPageIndex.������Ϣ
        End If
    End If
        
    With tabPage
        tabPage.Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        Set .PaintManager.Font = lbl��������.Font
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = False
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then Resume
End Sub

Private Function InitTaskPanel() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ظ�����Ϣҳ��
    '����:
    '�����:73935
    '����:Ƚ����
    '����:2014-07-3
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim tkpGroup As TaskPanelGroup, Item As TaskPanelGroupItem
    
    Err = 0: On Error GoTo Errhand
    If Not mobjPlugIn Is Nothing Then
        If mlngPlugInHwnd <> 0 Then
            With wndTaskPanel
                Call .SetGroupInnerMargins(0, 0, 0, 0)
                Call .SetGroupOuterMargins(-1, -24, -1, -1)
                
                Set tkpGroup = .Groups.Add(1, "������Ϣ")
                tkpGroup.CaptionVisible = False
                tkpGroup.Expandable = False
                tkpGroup.Expanded = True
                
                Set Item = tkpGroup.Items.Add(1, "", xtpTaskItemTypeControl)
                Call HideFormCaption(mlngPlugInHwnd, False) '���ش���߿�
                Item.Handle = mlngPlugInHwnd
                
                .HotTrackStyle = xtpTaskPanelHighlightItem
                .Reposition
                .DrawFocusRect = True
            End With
        End If
    End If

    InitTaskPanel = True
    
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
End Function

Private Sub zlPlugInErrH(ByVal objErr As Object, ByVal strFunName As String)
'���ܣ���Ҳ���������
'������objErr ������� strFunName �ӿڷ�������
'˵���������������ڣ������438��ʱ����ʾ���������󵯳���ʾ��
    If InStr(",438,0,", "," & objErr.Number & ",") = 0 Then
        MsgBox "zlPlugIn ��Ҳ���ִ�� " & strFunName & " ʱ����" & vbCrLf & objErr.Number & vbCrLf & objErr.Description, vbInformation, gstrSysName
    End If
End Sub

Private Function CreatePublicPatient() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����zlPublicPatient����
    '����:�����ɹ�,����True,���򷵻�False
    '����:Ƚ����
    '����:2014-08-08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mobjPubPatient Is Nothing Then
        On Error Resume Next
        Set mobjPubPatient = CreateObject("zlPublicPatient.clsPublicPatient")
        Err.Clear: On Error GoTo 0
    End If
    If mobjPubPatient Is Nothing Then
        MsgBox "������Ϣ����������zlPublicPatient������ʧ�ܣ�", vbInformation, gstrSysName
        Exit Function
    Else
        If mobjPubPatient.zlInitCommon(gcnOracle, glngSys, gstrDBUser) = False Then
            MsgBox "������Ϣ����������zlPublicPatient����ʼ��ʧ�ܣ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    CreatePublicPatient = True
End Function

Private Function CheckValied() As Boolean
'---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鲡����Ϣ�Ƿ�Ϸ�
    '����:������Ϣ�Ϸ�,����True,���򷵻�False
    '����:����
    '����:2017-07-03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSimilar  As String
    Dim i As Long
    Dim str�������� As String
    On Error GoTo errH
    
    
    If CheckTextLength("����", txtPatient) = False Then Exit Function
    If CheckTextLength("����", txt����) = False Then Exit Function
    If mblnStructAdress Then
        If Not CheckStructAddr(padd��ͥ��ַ, padd��ͥ��ַ.MaxLength) Then Exit Function
    End If

    If Trim(txtPatient.Text) = "" Then
        MsgBox "�������벡�����������飡", vbInformation, gstrSysName
        Call zlControl.ControlSetFocus(txtPatient): Exit Function
    End If
    If Trim(txtPatient.Text) <> mstr���� Then
        If MsgBox("�������Ѳ������֡�" & mstr���� & "���޸�Ϊ��" & Trim(txtPatient.Text) & "��,�Ƿ����?", _
            vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Call zlControl.ControlSetFocus(txtPatient): Exit Function
        End If
    End If
    
    If mbytType = 1 Then
        '������Ʋ�����Ϣ(����֮ǰ���,����������ظ���Ϣ������)
        If Trim(txt���֤��.Text) <> "" Then
            strSimilar = SimilarIDs(Trim(txt���֤��.Text))
            If strSimilar <> "" Then
                i = UBound(Split(strSimilar, "|")) + 1
                strSimilar = Replace(strSimilar, "|", vbCrLf)
                If i > 20 Then strSimilar = Mid(strSimilar, 1, 200) & "..."

                If MsgBox("�����еĲ�����Ϣ�з��� " & i & " ����Ϣ���ƵĲ���(���֤����ͬ): " & vbCrLf & vbCrLf & _
                          strSimilar & vbCrLf & vbCrLf & "ȷʵҪ�Ǽ�Ϊ�²�����", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            End If
        End If
    End If
    
    If txtPatiMCNO(0).Text <> "" Or txtPatiMCNO(1).Text <> "" Then
        If txtPatiMCNO(0).Text <> txtPatiMCNO(1).Text Then
            MsgBox "����,���������ҽ���Ų�һ�£�", vbInformation, gstrSysName
            If txtPatiMCNO(0).Visible Then txtPatiMCNO(0).SetFocus
            Exit Function
        End If
        If zlCommFun.ActualLen(txtPatiMCNO(0).Text) > txtPatiMCNO(0).MaxLength Then
            MsgBox "����,ҽ������󳤶Ȳ��ܳ���" & txtPatiMCNO(0).MaxLength & "���ַ���", vbInformation, gstrSysName
            If txtPatiMCNO(0).Visible Then txtPatiMCNO(0).SetFocus
            Exit Function
        End If
        If cboҽ�����.ListIndex <= 0 Then
            MsgBox "��ȷ��ҽ�����˵�ҽ�����", vbInformation, gstrSysName
            If cboҽ�����.Visible Then cboҽ�����.SetFocus
            Exit Function
        End If
    Else
        If cboҽ�����.ListIndex > 0 Then
            MsgBox "ѡ��ҽ�����ʱ��������ҽ���ţ�", vbInformation, gstrSysName
            If txtPatiMCNO(0).Visible Then txtPatiMCNO(0).SetFocus
            Exit Function
        End If
    End If
    
    '69026,Ƚ����,2014-8-11,������Ч�Լ��
    '76703,Ƚ����,2014-8-15
    If txt����.Enabled And txt����.Visible Then
        If mobjPubPatient Is Nothing Then Exit Function
        If mobjPubPatient.CheckPatiAge(Trim(txt����.Text) & IIf(cbo���䵥λ.Visible, cbo���䵥λ.Text, ""), _
                IIf(txt��������.Text = "____-__-__", "", txt��������.Text) & _
                IIf(txt����ʱ��.Text = "__:__", "", " " & txt����ʱ��.Text)) = False Then
            txt����.SetFocus: Exit Function
        End If
    End If
    
    If IsDate(txt��������.Text) Then
        '76669�����ϴ�,2014-8-15,������������ڼ��
        str�������� = txt��������.Text & IIf(IsDate(txt����ʱ��.Text), " " & txt����ʱ��.Text, "")
        If CDate(str��������) > zlDatabase.Currentdate Then
            If MsgBox("����ʱ�䣺" & str�������� & " �����˵�ǰϵͳʱ�䡣" & _
                vbCrLf & vbCrLf & "���������������ڵ���ȷ�� ��Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                If txt��������.Enabled And txt��������.Visible Then txt��������.SetFocus
                Exit Function
            End If
        End If
    End If
    
    If Not Me.Caption Like "�����Һ�*" Then
        '75909
        If mlng�Һ�ID <> 0 And mblnҽ��ҵ�� = True Then
            If mstr���� <> txtPatient.Text _
                Or mstr�Ա� <> NeedName(cbo�Ա�.Text) _
                Or mstr���� <> txt����.Text & cbo���䵥λ _
                Or mstr�������� <> txt��������.Text Then
                MsgBox "�ò����Ѿ�������ҽ������,������������˵Ļ�����Ϣ(����,�Ա�,�����),���ڡ�������Ϣ�����н��е�����", vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    If mlng�Һ�ID = 0 Then
        If Not GetRegBillID() Then
            MsgBox "�޷���ȡ�Һ�ID", vbInformation, gstrSysName
            Exit Function
        End If
    End If

    CheckValied = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
