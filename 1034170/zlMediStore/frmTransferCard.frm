VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmTransferCard 
   Caption         =   "ҩƷ�ƿⵥ"
   ClientHeight    =   6975
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10770
   Icon            =   "frmTransferCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   10770
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmd�޿������ɸѡ 
      Caption         =   "�޿������ɸѡ"
      Height          =   350
      Left            =   3240
      TabIndex        =   36
      Top             =   5520
      Visible         =   0   'False
      Width           =   1515
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh������Ϣ 
      Height          =   2175
      Left            =   5880
      TabIndex        =   33
      Top             =   1095
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3836
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
      GridColor       =   32768
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh���� 
      Height          =   2175
      Left            =   2310
      TabIndex        =   32
      Top             =   1485
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3836
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
   Begin VB.CommandButton cmdExpend 
      Caption         =   "�Զ��ֽ�(&A)"
      Height          =   350
      Left            =   4950
      TabIndex        =   7
      Top             =   5490
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmdAllSel 
      Caption         =   "ȫ��(&A)"
      Height          =   350
      Left            =   6180
      TabIndex        =   31
      Top             =   5490
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdAllCls 
      Caption         =   "ȫ��(&L)"
      Height          =   350
      Left            =   7500
      TabIndex        =   30
      Top             =   5490
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3720
      TabIndex        =   12
      Top             =   5137
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "����(&F)"
      Height          =   350
      Left            =   2040
      TabIndex        =   11
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   10
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6180
      TabIndex        =   8
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   7560
      TabIndex        =   9
      Top             =   5040
      Width           =   1100
   End
   Begin VB.PictureBox Pic���� 
      BackColor       =   &H80000004&
      Height          =   4965
      Left            =   0
      ScaleHeight     =   4905
      ScaleWidth      =   11655
      TabIndex        =   13
      Top             =   0
      Width           =   11715
      Begin VB.CheckBox chkIn 
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "������ʵ�:F3"
         Top             =   150
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtIn 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   135
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.ComboBox cboEnterStock 
         Height          =   300
         Left            =   9240
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   2115
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   2805
         Left            =   195
         TabIndex        =   4
         Top             =   960
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   4948
         Appearance      =   0
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Active          =   -1  'True
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483634
      End
      Begin VB.TextBox txtժҪ 
         Height          =   300
         Left            =   900
         MaxLength       =   40
         TabIndex        =   6
         Top             =   4080
         Width           =   10410
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   1515
      End
      Begin VB.Label lblDifference 
         AutoSize        =   -1  'True
         Caption         =   "��ۺϼ�:"
         Height          =   180
         Left            =   4920
         TabIndex        =   28
         Top             =   3840
         Width           =   810
      End
      Begin VB.Label lblSalePrice 
         AutoSize        =   -1  'True
         Caption         =   "�ۼ۽��ϼ�:"
         Height          =   180
         Left            =   2040
         TabIndex        =   27
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "�ɱ����ϼ�:"
         Height          =   180
         Left            =   240
         TabIndex        =   26
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label Txt����� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7950
         TabIndex        =   24
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label Txt������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10050
         TabIndex        =   23
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt�������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   22
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt������ 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   21
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label txtNo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   9960
         TabIndex        =   20
         Top             =   158
         Width           =   1425
      End
      Begin VB.Label LblNo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NO."
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   9480
         TabIndex        =   19
         Top             =   195
         Width           =   480
      End
      Begin VB.Label lblժҪ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ժҪ(&M)"
         Height          =   180
         Left            =   240
         TabIndex        =   5
         Top             =   4155
         Width           =   650
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ҩƷ�ƿⵥ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   30
         TabIndex        =   18
         Top             =   120
         Width           =   11535
      End
      Begin VB.Label LblStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ƴ��ⷿ(&S)"
         Height          =   180
         Left            =   240
         TabIndex        =   0
         Top             =   660
         Width           =   990
      End
      Begin VB.Label Lbl������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   300
         TabIndex        =   17
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Lbl�������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   2160
         TabIndex        =   16
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label Lbl����� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   180
         Left            =   7365
         TabIndex        =   15
         Top             =   4500
         Width           =   540
      End
      Begin VB.Label Lbl������� 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         Height          =   180
         Left            =   9240
         TabIndex        =   14
         Top             =   4500
         Width           =   720
      End
      Begin VB.Label LblEnterStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ⷿ(&I)"
         Height          =   180
         Left            =   8040
         TabIndex        =   2
         Top             =   660
         Width           =   990
      End
   End
   Begin MSComctlLib.ImageList imghot 
      Left            =   840
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":014A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":0364
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":057E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":0798
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":09B2
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":0BCC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":0DE6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":1000
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgcold 
      Left            =   120
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":121A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":1434
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":164E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":1868
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":1A82
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":1C9C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":1EB6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransferCard.frx":20D0
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   29
      Top             =   6615
      Width           =   10770
      _ExtentX        =   18997
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmTransferCard.frx":22EA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12647
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmTransferCard.frx":2B7E
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmTransferCard.frx":3080
            Key             =   "WB"
            Object.ToolTipText     =   "���(F7)"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
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
   Begin VB.Label lblCode 
      Caption         =   "����"
      Height          =   255
      Left            =   3240
      TabIndex        =   25
      Top             =   5160
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Menu mnuCol 
      Caption         =   "����"
      Visible         =   0   'False
      Begin VB.Menu mnuColDrug 
         Caption         =   "ҩ��(���������)"
         Index           =   0
      End
      Begin VB.Menu mnuColDrug 
         Caption         =   "ҩ��(������)"
         Index           =   1
      End
      Begin VB.Menu mnuColDrug 
         Caption         =   "ҩ��(������)"
         Index           =   2
      End
   End
   Begin VB.Menu mnuFilter 
      Caption         =   "�޿������ɸѡ"
      Visible         =   0   'False
      Begin VB.Menu mnuFilterDrug 
         Caption         =   "�޿���������"
         Index           =   0
      End
      Begin VB.Menu mnuFilterDrug 
         Caption         =   "ɾ���޿������"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmTransferCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintSelectStock As Integer           '�Ƿ��ѡ�ⷿ
Private mint�༭״̬ As Integer             '1.������2���޸ģ�3�����գ�4���鿴��5��6-������10-����,11-����ⵥ��ȡ����
Private mstr���ݺ� As String                '����ĵ��ݺ�;
Private mint��¼״̬ As Integer             '1:������¼;2-������¼;3-�Ѿ�������ԭ��¼
Private mblnSuccess As Boolean              'ֻҪ��һ�ųɹ�����ΪTrue������ΪFalse
Private mblnSave As Boolean                 '�Ƿ���̺����   TURE���ɹ���
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '�Ƿ�����޸�
Private mblnChange As Boolean               '�Ƿ���й��༭
Private mbln���쵥 As Boolean               '�Ƿ������쵥�������������ִ���Զ��ֽ�Ĺ���
Private mintApplyType As Integer            '���췽ʽ��0-�ֹ�����;1-����������;2-��������;3-��������;4-����������;5-�������쵥δ����;6-������������;7-������������
Private mstrEndTime As String               '���Զ����췽ʽΪ7ʱ������ʱ�䷶Χ�еĽ���ʱ��
Private mbln��ȷ���� As Boolean             '�Ƿ���ȷ���Σ��������쵥��Ч
Private mblnEnterCell As Boolean            '�Ƿ�������ENTERCELL�����¼���ȱʡΪ��
Private mlng����ⷿ As Long
Private mlng����ⷿ As Long                '����������ⵥ�ƿ�
Private mstr��ⵥ�� As String              '����������ⵥ�ƿ�
Private mint����� As Integer             '��ʾҩƷ����ʱ�Ƿ���п���飺0-�����;1-��飬�������ѣ�2-��飬�����ֹ
Private mint��������ⷿ As Integer     '�����ڳ���ʱ��ԭ���ⷿ�Ƿ���п���飺0-�����;1-��飬�������ѣ�2-��飬�����ֹ
Private mcolUsedCount As Collection         '��ʹ�õ���������
Private mstrPrivs As String                 'Ȩ��
Private mblnRS As Boolean                   '������¼���ݼ���״̬
Private mblnViewCost As Boolean             '�鿴�ɱ��� true-���Բ鿴�ɱ��� false-�����Բ鿴�ɱ���

Private mintParallelRecord As Integer       '���������󵥾ݲ���ִ�еĴ��� 1���������������2���Ѿ�ɾ���ļ�¼��3���Ѿ���˵ļ�¼
Private mintBatchNoLen As Integer           '���ݿ������Ŷ��峤��
Private rsDepend As New ADODB.Recordset
Private recSort As ADODB.Recordset          '��ҩƷID�����ר�ü�¼��
Private mrsMyAppend As New ADODB.Recordset  '������̬��¼��

Private Const MStrCaption As String = "ҩƷ�ƿ����"

Private mobjPlugIn As Object '��Ҳ���

Private Const mlng��ɫ As Long = &HC000C0

Private mblnUpdate As Boolean               '��ʾ�Ƿ��Ѹ������¼۸���µ�������

Private mintUnit As Integer             '��λϵ����1-�ۼ�;2-����;3-סԺ;4-ҩ��

Private mintDrugNameShow As Integer         'ҩƷ��ʾ��0����ʾ��������ƣ�1������ʾ���룻2������ʾ����

'�Ӳ�������ȡҩƷ�۸����������С��λ�������㾫�ȣ�
Private mintCostDigit As Integer        '�ɱ���С��λ��
Private mintPriceDigit As Integer       '�ۼ�С��λ��
Private mintNumberDigit As Integer      '����С��λ��
Private mintMoneyDigit As Integer       '���С��λ��

Private Const mconint�ۼ۵�λ As Integer = 1
Private Const mconint���ﵥλ As Integer = 2
Private Const mconintסԺ��λ As Integer = 3
Private Const mconintҩ�ⵥλ As Integer = 4

Private mstrTime_Start As String                      '���뵥�ݱ༭����ʱ�����༭���ݵ�����޸�ʱ��
Private mstrTime_End As String                        '�˿̸ñ༭���ݵ�����޸�ʱ��

Private mint�ƿ⴦������ As Integer                   '1-��Ҫ��ҩ�����͡�������һ����  0-����Ҫ��һ����
Private mint����ʽ As Integer                       '����ʱ��0������������1�������������뵥�ݣ�2������Ѳ����ĳ������뵥��
Private mbln�Զ��ֽ�δ��� As Boolean                 '��Ҫ�Զ��ֽⲢ���Զ��ֽ�δ���
Private mbln�¿������� As Boolean                     '�Ƿ��·�ҩҩ���Ŀ�������

'=========================================================================================
Private Const mconIntCol��� As Integer = 1
Private Const mconIntCol�к� As Integer = 2
Private Const mconIntColҩ�� As Integer = 3
Private Const mconIntCol��Ʒ�� As Integer = 4
Private Const mconIntCol��Դ As Integer = 5
Private Const mconIntCol����ҩ�� As Integer = 6
Private Const mconIntCol��� As Integer = 7
Private Const mconIntCol�������� As Integer = 8
Private Const mconIntCol���Ч�� As Integer = 9
Private Const mconIntCol�������� As Integer = 10
Private Const mconIntColָ������� As Integer = 11
Private Const mconIntColʵ�ʽ�� As Integer = 12
Private Const mconIntColʵ�ʲ�� As Integer = 13
Private Const mconIntCol����ϵ�� As Integer = 14
Private Const mconIntCol���� As Integer = 15
Private Const mconIntCol���� As Integer = 16
Private Const mconIntCol��λ As Integer = 17
Private Const mconIntCol�ͻ���λ As Integer = 18
Private Const mconIntCol���� As Integer = 19
Private Const mconIntColЧ�� As Integer = 20
Private Const mconIntCol��׼�ĺ� As Integer = 21
Private Const mconIntCol�ⷿ��� As Integer = 22
Private Const mconIntCol�Է���� As Integer = 23
Private Const mconIntCol��д���� As Integer = 24
Private Const mconIntColʵ������ As Integer = 25
Private Const mconIntCol�ɹ��� As Integer = 26
Private Const mconIntCol�ɹ���� As Integer = 27
Private Const mconIntCol�ۼ� As Integer = 28
Private Const mconIntCol�ۼ۽�� As Integer = 29
Private Const mconintCol��� As Integer = 30
Private Const mconIntCol�ϴι�Ӧ��ID As Integer = 31
Private Const mconintCol��ʵ���� As Integer = 32
Private Const mconIntColҩƷ��������� = 33
Private Const mconIntColҩƷ���� = 34
Private Const mconIntColҩƷ���� = 35
Private Const mconIntCol�������� = 36
Private Const mconIntColS  As Integer = 37             '������
'=========================================================================================

Private Sub RefreshBill()
    '�����¼۸����µ���������ݣ����ڵ������ʱ
    Dim lngRow As Long, lngRows As Long, lngҩƷID As Long
    Dim Dbl���� As Double, dbl�ɱ��� As Double, dbl�ɱ���� As Double, dbl���ۼ� As Double, dbl���۽�� As Double, dbl��� As Double
    Dim rsPrice As New ADODB.Recordset
    Dim rsStock As ADODB.Recordset
    Dim blnAdj As Boolean
    Dim intCostDigit As Integer
    Dim intPriceDigit As Integer
        
    On Error GoTo errHandle
    intPriceDigit = GetDigit(0, 1, 2, 1)
    intCostDigit = GetDigit(0, 1, 1, 1)
        
    gstrSQL = " Select '�ۼ�' As ����, a.���, a.ҩƷid ,nvl(a.����,0) as ����, 0 ԭ��, b.�ּ�" & _
            " From ҩƷ�շ���¼ A," & _
                 " (Select �շ�ϸĿid, Nvl(�ּ�, 0) �ּ�, ִ������" & _
                   " From �շѼ�Ŀ" & _
                   " Where (��ֹ���� Is Null Or Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'yyyy-MM-dd')))) B, �շ���ĿĿ¼ C" & _
            " Where a.���� = 6 And a.No = [1] And a.ҩƷid = b.�շ�ϸĿid And c.Id = b.�շ�ϸĿid And Round(a.���ۼ�," & intPriceDigit & ") <> Round(b.�ּ�, " & intPriceDigit & ") And" & _
              "    NVL(c.�Ƿ���, 0) = 0  and b.ִ������>a.��������" & _
            " Union All" & _
            " Select '�ۼ�' As ����, a.���, a.ҩƷid ,nvl(a.����,0) as ����, 0 ԭ��, decode(nvl(b.����,0),0,b.ʵ�ʽ�� / b.ʵ������,b.���ۼ�) As �ּ�" & _
            " From ҩƷ�շ���¼ A, ҩƷ��� B, �շ���ĿĿ¼ C" & _
            " Where a.���� = 6 And a.No = [1] And c.Id = a.ҩƷid And Round(a.���ۼ�," & intPriceDigit & ") <> Round(decode(nvl(b.����,0),0,b.ʵ�ʽ�� / b.ʵ������,b.���ۼ�), " & intPriceDigit & ") And Nvl(c.�Ƿ���, 0) = 1 And" & _
                  " b.���� = 1 And b.�ⷿid = a.�ⷿid And b.ҩƷid = a.ҩƷid And NVL(b.����, 0) = NVL(a.����, 0) And NVL(b.ʵ������, 0) <> 0 And a.���ϵ�� = -1" & _
            " Union All" & _
            " Select '�ɱ���' As ����, a.���, a.ҩƷid ,nvl(a.����,0) as ����, 0 ԭ��, b.ƽ���ɱ��� As �ּ�" & _
            " From ҩƷ�շ���¼ A, ҩƷ��� B" & _
            " Where a.���� = 6 And a.No = [1] And a.ҩƷid = b.ҩƷid And Nvl(a.����, 0) = Nvl(b.����, 0) and round(a.�ɱ���," & intCostDigit & ")<>round(b.ƽ���ɱ���," & intCostDigit & ") And a.�ⷿid = b.�ⷿid and a.���ϵ��=-1  and b.����=1" & _
            " Order By ����, ҩƷid, ���"
    Set rsPrice = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[ȡ��ǰ�۸�]", CStr(Me.txtNo.Tag))
    
    If rsPrice.EOF Then Exit Sub
    
    lngRows = mshBill.rows - 1
    For lngRow = 1 To lngRows
        blnAdj = False
        lngҩƷID = Val(mshBill.TextMatrix(lngRow, 0))
        Dbl���� = Val(mshBill.TextMatrix(lngRow, mconIntColʵ������))
        dbl�ɱ��� = Val(mshBill.TextMatrix(lngRow, mconIntCol�ɹ���))
        dbl���ۼ� = Val(mshBill.TextMatrix(lngRow, mconIntCol�ۼ�))
        dbl�ɱ���� = dbl�ɱ��� * Dbl����
        dbl���۽�� = dbl���ۼ� * Dbl����
        dbl��� = dbl���۽�� - dbl�ɱ����
                
        If lngҩƷID <> 0 Then
            rsPrice.Filter = "����='�ۼ�' And ҩƷID=" & lngҩƷID & " And ����=" & Val(mshBill.TextMatrix(lngRow, mconIntCol����))
            If rsPrice.RecordCount > 0 Then
                blnAdj = True
                dbl���ۼ� = Val(GetFormat(rsPrice!�ּ� * Val(mshBill.TextMatrix(lngRow, mconIntCol����ϵ��)), intPriceDigit))
                dbl���۽�� = Val(GetFormat(Val(FormatEx(dbl���ۼ�, intPriceDigit)) * Dbl����, mintMoneyDigit))
                dbl��� = Val(GetFormat(dbl���۽�� - dbl�ɱ����, mintMoneyDigit))
            End If
            
            rsPrice.Filter = "����='�ɱ���' And ҩƷID=" & lngҩƷID & " And ����=" & Val(mshBill.TextMatrix(lngRow, mconIntCol����))
            If rsPrice.RecordCount > 0 Then
                blnAdj = True
                dbl���۽�� = Val(GetFormat(Val(FormatEx(dbl���ۼ�, intPriceDigit)) * Dbl����, mintMoneyDigit))
                dbl�ɱ��� = Val(GetFormat(rsPrice!�ּ� * Val(mshBill.TextMatrix(lngRow, mconIntCol����ϵ��)), intCostDigit))
                dbl�ɱ���� = Val(GetFormat(dbl�ɱ��� * Dbl����, mintMoneyDigit))
                dbl��� = Val(GetFormat(dbl���۽�� - dbl�ɱ����, mintMoneyDigit))
            End If
            
            If blnAdj = True Then
                '�Ե�ǰ���¼۸����µ���������ݣ��ۼۡ��ɱ��ۡ����۽��ɱ�����ۣ�
                mshBill.TextMatrix(lngRow, mconIntCol�ۼ�) = GetFormat(dbl���ۼ�, intPriceDigit)
                mshBill.TextMatrix(lngRow, mconIntCol�ۼ۽��) = GetFormat(dbl���۽��, mintMoneyDigit)
                mshBill.TextMatrix(lngRow, mconIntCol�ɹ���) = GetFormat(dbl�ɱ���, intCostDigit)
                mshBill.TextMatrix(lngRow, mconIntCol�ɹ����) = GetFormat(dbl�ɱ����, mintMoneyDigit)
                mshBill.TextMatrix(lngRow, mconintCol���) = GetFormat(dbl���, mintMoneyDigit)
            End If
        End If
    Next
    rsPrice.Filter = 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetSortRecord()
    Dim n As Integer
    
    If mshBill.rows < 2 Then Exit Sub
    If mshBill.TextMatrix(1, 0) = "" Then Exit Sub
    
    Set recSort = New ADODB.Recordset
    With recSort
        If .State = 1 Then .Close
        .Fields.Append "�к�", adDouble, 18, adFldIsNullable
        .Fields.Append "���", adDouble, 18, adFldIsNullable
        .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
        
        For n = 1 To mshBill.rows - 1
            If mshBill.TextMatrix(n, 0) <> "" Then
                .AddNew
                !�к� = n
                !��� = IIf(Val(mshBill.TextMatrix(n, mconIntCol���)) = 0, n, Val(mshBill.TextMatrix(n, mconIntCol���)))
                !ҩƷid = Val(mshBill.TextMatrix(n, 0))
                !���� = Val(mshBill.TextMatrix(n, mconIntCol����))
                
                .Update
            End If
        Next
        
    End With
End Sub
Public Function CheckBill(ByVal strNo As String) As String
    Dim rs As New ADODB.Recordset
    
    CheckBill = ""
    On Error GoTo errHandle
    gstrSQL = "Select �������,��ҩ����,��ҩ�� From ҩƷ�շ���¼ " & _
              "Where ����=6 And NO=[1] And ��¼״̬=1 And RowNum=1 "
    Set rs = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��鵥��]", strNo)
    
    With rs
        '���ؿգ���ʾ�Ѿ�ɾ��
        If .EOF Then
            CheckBill = "�õ����Ѿ�����������Աɾ����"
        End If
        If mint�༭״̬ = 3 Then
            If Not IsNull(!�������) Then
                CheckBill = "�õ����Ѿ�����������Ա��ˣ�"
            End If
            Exit Function
        End If
        
        If mint�༭״̬ = 10 Then
            If Not IsNull(!��ҩ����) Then
                CheckBill = "�õ����Ѿ�����������Ա���ͣ�"
            End If
            Exit Function
        End If
                    
        If Not IsNull(!��ҩ��) Then
            CheckBill = "�õ����Ѿ�����������Ա��ҩ��"
        End If
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function
Private Function Auto�����ƿ�����(Optional ByVal blnǿ�Ʊ��� As Boolean = False) As Boolean
    Dim blnTrans As Boolean
        
    '�Զ������ƿ����� 1����ҩ 2������ 3������
    
    On Error GoTo errHandle
    
    If Not ��鵥��(6, txtNo, False) And Not mblnUpdate Then
        '�����µļ۸���µ����壬�˳���Ŀ�������û���һ�����յĵ���
        MsgBox "�м�¼δʹ�������ۼۣ������Զ���ɸ��£��ۼۡ��ۼ۽���ۣ������º����飡", vbInformation, gstrSysName
        Call RefreshBill
        mblnUpdate = True
        mblnChange = True
        Exit Function
    End If
    
    If Not ҩƷ�������(Txt������.Caption) Then Exit Function
    
    If Not blnǿ�Ʊ��� Then
        blnTrans = True
        gcnOracle.BeginTrans
    End If
    
    '1-
    gstrSQL = "zl_ҩƷ�ƿ�_PREPARE('" & txtNo.Tag & "','" & UserInfo.�û����� & "')"
    Call zlDataBase.ExecuteProcedure(gstrSQL, "��ҩ")
    
    '2-
    If Not ValidData Then
        If blnTrans Then
            gcnOracle.RollbackTrans
        End If
        Exit Function
    End If
    
    '��������¿�������Ϊ�٣���Ҫ�����ÿ���Ƿ��㹻
    If mbln�¿������� = False Then
        If Not CheckStock Then
            If blnTrans Then
                gcnOracle.RollbackTrans
            End If
            Exit Function
        End If
    End If
    
    '��ɾ�����쵥�������ݵ�ǰ���ݲ����ƿⵥ������Ǵ����ת���ƿ�ĵ��ݣ���ִ��
    If mint�༭״̬ <> 11 And mblnChange = True Then
        If Not SaveCard(True) Then
            If blnTrans Then
                gcnOracle.RollbackTrans
            End If
            Exit Function
        End If
    End If
    
    '��ҩ
    gstrSQL = "zl_ҩƷ�ƿ�_Prepare('" & txtNo.Tag & "','" & UserInfo.�û����� & "')"
    Call zlDataBase.ExecuteProcedure(gstrSQL, "��ҩ")
    '���ͣ��³���ⷿ��ҩƷ���ÿ�棩
    gstrSQL = "zl_ҩƷ�ƿ�_Prepare('" & txtNo.Tag & "')"
    Call zlDataBase.ExecuteProcedure(gstrSQL, "����")
       
   
    '3-
    If SaveCheck(True) = True Then
        If Val(zlDataBase.GetPara("��˴�ӡ", glngSys, ģ���.ҩƷ�ƿ�)) = 1 Then
            '��ӡ
            If IsHavePrivs(mstrPrivs, "���ݴ�ӡ") Then
                printbill
            End If
        End If
        Unload Me
    Else
        GoTo errHandle
    End If
    
    If Not blnǿ�Ʊ��� Then
        blnTrans = True
        gcnOracle.CommitTrans
    End If
    
    Auto�����ƿ����� = True
    
    Exit Function
    
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    Auto�����ƿ����� = False
End Function

'�������������
Private Function GetDepend() As Boolean
    Dim strMsg As String
    GetDepend = False
    On Error GoTo errHandle
    
    '���ҩƷ�������Ƿ�����
    strMsg = "û������ҩƷ�ƿ����⼰�����������ҩƷ������࣡"
    gstrSQL = "SELECT B.Id,B.ϵ�� " _
            & "FROM ҩƷ�������� A, ҩƷ������ B " _
            & "Where A.���id = B.ID AND A.���� = 6 "
    
    If rsDepend.State = 1 Then rsDepend.Close
    Set rsDepend = zlDataBase.OpenSQLRecord(gstrSQL, "ҩƷ�ƿ����")
    
    With rsDepend
        If .RecordCount = 0 Then Exit Function
        .Filter = "ϵ��=1"
        If .RecordCount = 0 Then
            .Filter = 0
            strMsg = "û������ҩƷ�ƿ������������ҩƷ������࣡"
            Exit Function
        End If
        .Filter = "ϵ��=-1"
        If .RecordCount = 0 Then
            .Filter = 0
            strMsg = "û������ҩƷ�ƿ�ĳ����������ҩƷ������࣡"
            Exit Function
        End If
        .Filter = 0
        .Close
    End With
    
    GetDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Sub ShowCard(FrmMain As Form, ByVal str���ݺ� As String, ByVal int�༭״̬ As Integer, _
    Optional int��¼״̬ As Integer = 1, Optional BlnSuccess As Boolean = False, Optional int����ʽ As Integer = 0)
    mblnSave = False
    mblnSuccess = False
    If int�༭״̬ = 11 Then
        mstr��ⵥ�� = str���ݺ�
        mstr���ݺ� = ""
    Else
        mstr���ݺ� = str���ݺ�
    End If
    mint�༭״̬ = int�༭״̬
    mint��¼״̬ = int��¼״̬
    mblnSuccess = BlnSuccess
    mblnChange = False
    mintParallelRecord = 1
    mint����ʽ = int����ʽ
    mstrPrivs = GetPrivFunc(glngSys, 1304)
    
    mint�ƿ⴦������ = Val(zlDataBase.GetPara("�ƿ�����", glngSys, ģ���.ҩƷ�ƿ�))
    
    Set mfrmMain = FrmMain
    If Not GetDepend Then Exit Sub
    
    mblnEdit = False
    If mint�༭״̬ = 1 Then
        mblnEdit = True
    ElseIf mint�༭״̬ = 2 Then
        mblnEdit = True
    ElseIf mint�༭״̬ = 3 Then
        mblnEdit = False
        CmdSave.Caption = "���(&V)"
    ElseIf mint�༭״̬ = 4 Then
        mblnEdit = False
        CmdSave.Caption = "��ӡ(&P)"
        If Not IsHavePrivs(mstrPrivs, "���ݴ�ӡ") Then
            CmdSave.Visible = False
        Else
            CmdSave.Visible = True
        End If
    ElseIf mint�༭״̬ = 6 Then
        mblnEdit = False
        
        cmdAllSel.Visible = True
        cmdAllCls.Visible = True
        
        If mint����ʽ = 1 Then
            CmdSave.Caption = "�������(&O)"
            CmdSave.Width = CmdSave.Width + 200
        ElseIf mint����ʽ = 2 Then
            CmdSave.Caption = "��˳���(&V)"
            CmdSave.Width = CmdSave.Width + 200
            
            cmdAllSel.Visible = False
            cmdAllCls.Visible = False
        Else
            CmdSave.Caption = "����(&O)"
            CmdSave.Width = CmdCancel.Width
        End If
    ElseIf mint�༭״̬ = 11 Then
        mblnEdit = True
        
        '�����û��������Ȩ�޲��Ҳ���Ҫ��ҩ���͹���ʱ������ֱ�����
        If IsHavePrivs(mstrPrivs, "���") And mint�ƿ⴦������ = 0 Then
            CmdSave.Caption = "���(&V)"
        End If
    End If
    
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    
    Me.Show vbModal, FrmMain
    BlnSuccess = mblnSuccess
    str���ݺ� = mstr���ݺ�
End Sub

Private Sub cboEnterStock_Click()
    If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
        If mblnRS Then
            Call SetSelectorRS(2, MStrCaption, cboStock.ItemData(cboStock.ListIndex), cboEnterStock.ItemData(cboEnterStock.ListIndex), mlng����ⷿ)
        End If
        mblnRS = True
    End If
End Sub

Private Sub cboEnterStock_Validate(Cancel As Boolean)
    Dim i As Integer
    
    With cboEnterStock
        If .ListCount = 0 Then Exit Sub
        If .ListIndex <> Val(.Tag) Then
            For i = 1 To mshBill.rows - 1
                If mshBill.TextMatrix(i, 0) <> "" Then
                    Exit For
                End If
            Next
            If i <> mshBill.rows Then
                If MsgBox("����ı�����ⷿ���п���Ҫ�ı���ӦҩƷ�ĵ�λ����������Ҫ������е������ݣ����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    '����ҩƷ��λ�ı�
                    cboEnterStock.Tag = .ListIndex
                    mshBill.ClearBill
                Else
                    .ListIndex = Val(.Tag)
                End If
            Else
                .Tag = .ListIndex
            End If
        End If
        
    End With
End Sub

Private Sub cboStock_Change()
    mblnChange = True
End Sub

Private Sub cboStock_Click()
    Dim rsStock As New ADODB.Recordset
    Dim lngEnterStockIndex As Long
    Dim blnHaveIndex As Boolean
    
    '��鲢װ������ⷿ
    On Error Resume Next
    
    lngEnterStockIndex = 0
    blnHaveIndex = False
    
    Set rsStock = ReturnSQL(cboStock.ItemData(cboStock.ListIndex), MStrCaption, True, 1304)
    
    With rsStock
         cboEnterStock.Clear
         Do While Not .EOF
             cboEnterStock.AddItem !����
             cboEnterStock.ItemData(cboEnterStock.NewIndex) = !id
             If Not blnHaveIndex And mint�༭״̬ = 11 Then
                 If .Fields(0) = mlng����ⷿ Then
                     lngEnterStockIndex = .AbsolutePosition - 1
                     blnHaveIndex = True
                 End If
             End If
             .MoveNext
         Loop
         cboEnterStock.ListIndex = 0
         
         If cboEnterStock.ListCount > 0 Then
            If cboEnterStock.ListCount > Val(cboEnterStock.Tag) Or (lngEnterStockIndex <> 0 And cboEnterStock.ListCount > lngEnterStockIndex) Then
                cboEnterStock.ListIndex = IIf(lngEnterStockIndex = 0, Val(cboEnterStock.Tag), lngEnterStockIndex)
                cboEnterStock.Tag = cboEnterStock.ListIndex
            End If
         End If
             
    End With
    
    mlng����ⷿ = cboStock.ItemData(cboStock.ListIndex)
    mint����� = MediWork_GetCheckStockRule(cboStock.ItemData(cboStock.ListIndex))
    mint��������ⷿ = MediWork_GetCheckStockRule(cboEnterStock.ItemData(cboEnterStock.ListIndex))
    
    If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
        Call SetSelectorRS(2, MStrCaption, cboStock.ItemData(cboStock.ListIndex), cboEnterStock.ItemData(cboEnterStock.ListIndex), mlng����ⷿ)
    End If
End Sub

Private Sub cboStock_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim i As Integer
        Dim blnReturn As Boolean
        
        blnReturn = False
        
        cboStock_Validate blnReturn
        If blnReturn = True Then Exit Sub
        
        zlCommFun.PressKey (vbKeyTab)
    End If
    
End Sub

Private Sub cboEnterStock_KeyPress(KeyAscii As Integer)
    Dim blnReturn As Boolean

    If KeyAscii <> 13 Then Exit Sub
    blnReturn = False
    cboEnterStock_Validate blnReturn
    If blnReturn = True Then Exit Sub

    With mshBill
        .SetFocus
        .Row = 1
        .Col = mconIntColҩ��
    End With
        
End Sub

Private Sub cboStock_Validate(Cancel As Boolean)
    Dim i As Integer
    
    With cboStock
        If .ListIndex <> mintcboIndex Then
            For i = 1 To mshBill.rows - 1
                If mshBill.TextMatrix(i, 0) <> "" Then
                    Exit For
                End If
            Next
            If i <> mshBill.rows Then
                If MsgBox("����ı��Ƴ��ⷿ���п���Ҫ�ı���ӦҩƷ�ĵ�λ����Ҫ������е������ݣ����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    '����ҩƷ��λ�ı�
                    mintcboIndex = .ListIndex
                    mshBill.ClearBill
                    
                    mlng����ⷿ = Me.cboStock.ItemData(Me.cboStock.ListIndex)
                    Call GetDrugDigit(mlng����ⷿ, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
                Else
                    .ListIndex = mintcboIndex
                End If
            Else
                mintcboIndex = .ListIndex
                
                mlng����ⷿ = Me.cboStock.ItemData(Me.cboStock.ListIndex)
                Call GetDrugDigit(mlng����ⷿ, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
            End If
        End If
        
        
    End With
End Sub

Private Sub chkIn_Click()
    txtIn.Enabled = chkIn.Value
    If chkIn.Value Then
        txtIn.SetFocus
    Else
        txtIn.Text = ""
    End If
End Sub


Private Sub cmdAllCls_Click()
    Dim intRow As Integer
    
    With mshBill
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                .TextMatrix(intRow, mconIntColʵ������) = GetFormat(0, mintNumberDigit)
                .TextMatrix(intRow, mconIntCol�ɹ����) = GetFormat(0, mintMoneyDigit)
                .TextMatrix(intRow, mconIntCol�ۼ۽��) = GetFormat(0, mintMoneyDigit)
                .TextMatrix(intRow, mconintCol���) = GetFormat(0, mintMoneyDigit)
            End If
        Next
    End With
    Call ��ʾ�ϼƽ��
    mblnChange = False
End Sub

Private Sub cmdAllSel_Click()
    Dim intRow As Integer
    
    With mshBill
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                .TextMatrix(intRow, mconIntColʵ������) = .TextMatrix(intRow, mconIntCol��д����)
                .TextMatrix(intRow, mconIntCol�ɹ����) = GetFormat(.TextMatrix(intRow, mconIntCol��д����) * .TextMatrix(intRow, mconIntCol�ɹ���), mintMoneyDigit)
                .TextMatrix(intRow, mconIntCol�ۼ۽��) = GetFormat(.TextMatrix(intRow, mconIntCol��д����) * .TextMatrix(intRow, mconIntCol�ۼ�), mintMoneyDigit)
                .TextMatrix(intRow, mconintCol���) = GetFormat(.TextMatrix(intRow, mconIntCol�ۼ۽��) - .TextMatrix(intRow, mconIntCol�ɹ����), mintMoneyDigit)
            End If
        Next
    End With
    Call ��ʾ�ϼƽ��
    mblnChange = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdExpend_Click()
    Call AutoExpend
End Sub

'����
Private Sub cmdFind_Click()
    If lblCode.Visible = False Then
        lblCode.Visible = True
        txtCode.Visible = True
        txtCode.SetFocus
    Else
        FindRow mshBill, mconIntColҩƷ���������, txtCode.Text, True
        lblCode.Visible = False
        txtCode.Visible = False
    End If
End Sub

Private Sub SetDrugName(ByVal intType As Integer)
    'ҩƷ������ʾ��
    'intType��0����ʾ��������ƣ�1������ʾ���룻2������ʾ����
    Dim lngRow As Long
    
    If intType = mintDrugNameShow Then Exit Sub
    
    mintDrugNameShow = intType
    
    With mshBill
        For lngRow = 1 To .rows - 1
            If .TextMatrix(lngRow, mconIntColҩ��) <> "" Then
                If mintDrugNameShow = 1 Then
                    .TextMatrix(lngRow, mconIntColҩ��) = .TextMatrix(lngRow, mconIntColҩƷ����)
                ElseIf mintDrugNameShow = 2 Then
                    .TextMatrix(lngRow, mconIntColҩ��) = .TextMatrix(lngRow, mconIntColҩƷ����)
                Else
                    .TextMatrix(lngRow, mconIntColҩ��) = .TextMatrix(lngRow, mconIntColҩƷ���������)
                End If
            End If
        Next
    End With
End Sub
Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub CmdSave_Click()
    Dim BlnSuccess As Boolean
    Dim blnTrans As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim bln�ⷿ As Boolean
    Dim bln���� As Boolean
    Dim intRow As Integer
    Dim lngҩƷID As Long
    
    On Error GoTo ErrHand
    '���͵ĳ��������̣��Զ��ֽ⡢����桢ɾ��ԭ���ݡ����������ݲ����µ��ƿⵥ�����±�ҩ������
    '��˵ĳ��������̣���˵��ݣ����ʵ����������д��������������������ⷿ�Ŀ������������³���ⷿ��ʵ�������������ⷿ�Ŀ�����ʵ��������
    
    '�����������ݼ�
    Call SetSortRecord
   
    If mint�༭״̬ = 4 Then    '�鿴
        '��ӡ
        printbill
        '�˳�
        Unload Me
        Exit Sub
    End If
    
    If mint�༭״̬ = 10 Then        '����
        '����������ֽ⣬����������ˣ���˴˴�����飬ǿ���û��ֹ�����ֽ⹦��
        'If Not AutoExpend(True) Then Exit Sub
        
        If mbln�Զ��ֽ�δ��� = True Then
            MsgBox "��ҩƷδ�����Զ��ֽ⣬����ִ���Զ��ֽ⣡", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If Not ValidData Then Exit Sub
        
        '��������¿�������Ϊ�٣���Ҫ�����ÿ���Ƿ��㹻
        If mbln�¿������� = False Then
            If Not CheckStock Then Exit Sub
        End If
 
        '����Ƿ��ѱ�ҩ
        gstrSQL = "Select 1 From ҩƷ�շ���¼ Where ����=6 And NO=[1] And ��ҩ�� Is Not NULL"
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[����Ƿ�ҩ]", txtNo.Tag)
        
        If rsTemp.RecordCount = 0 Then
            MsgBox "�õ����ѱ���������Աȡ����ҩ����ǰ������ֹ��", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '����Ƿ��ѷ���
        gstrSQL = "Select 1 From ҩƷ�շ���¼ Where ����=6 And NO=[1] And ��ҩ���� Is Not NULL"
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[����Ƿ���]", txtNo.Tag)
        
        If rsTemp.RecordCount <> 0 Then
            MsgBox "�õ����ѱ���������Ա���ͣ���ǰ������ֹ��", vbInformation, gstrSysName
            Exit Sub
        End If
        
        blnTrans = True
        gcnOracle.BeginTrans
        
        '��ɾ�����쵥�������ݵ�ǰ���ݲ����ƿⵥ
        If Not SaveCard(True) Then
            gcnOracle.RollbackTrans: Exit Sub
        End If
        '��ҩ
        gstrSQL = "zl_ҩƷ�ƿ�_Prepare('" & txtNo.Tag & "','" & Txt�����.Caption & "')"
        Call zlDataBase.ExecuteProcedure(gstrSQL, "��ҩ")
        '���ͣ��³���ⷿ��ҩƷ���ÿ�棩
        gstrSQL = "zl_ҩƷ�ƿ�_Prepare('" & txtNo.Tag & "')"
        Call zlDataBase.ExecuteProcedure(gstrSQL, "����")
        
        gcnOracle.CommitTrans
        blnTrans = True
        
        If Val(zlDataBase.GetPara("���ʹ�ӡ", glngSys, ģ���.ҩƷ�ƿ�)) = 1 Then
            '��ӡ
            If IsHavePrivs(mstrPrivs, "���ݴ�ӡ") Then
                printbill
            End If
        End If
        
        Unload Me
        Exit Sub
    End If
    
    If mint�༭״̬ = 3 Then       '���
        '�ƿ����ʱ��Ҫ���ж��Ƿ�������û�����εļ�¼
        If cmdExpend.Visible = True And mbln�Զ��ֽ�δ��� = True Then '�Զ��ֽⲻ�ɼ���ʾ��ҩʱ�Ѿ��Զ��ֽ⣻mbln�Զ��ֽ�δ��ɼ�¼���ʱ�Ѿ��Զ��ֽ�
            If cmdExpend.Enabled = True Then
                bln�ⷿ = CheckStockProperty(cboStock.ItemData(cboStock.ListIndex))
                With mshBill
                    For intRow = 1 To .rows - 1
                        lngҩƷID = Val(.TextMatrix(intRow, 0))
                        If lngҩƷID <> 0 Then
                            gstrSQL = " Select Nvl(A.ҩ�����,0) ҩ�����,Nvl(A.ҩ������,0) ҩ������" & _
                                              " From ҩƷ��� A" & _
                                              " Where A.ҩƷID =[1] "
                            Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[ȡ��������]", lngҩƷID)
                            bln���� = IIf(bln�ⷿ, (rsTemp!ҩ����� = 1), (rsTemp!ҩ������ = 1))
                            If bln���� = True And Val(.TextMatrix(intRow, mconIntCol����)) = 0 Then
                                MsgBox .TextMatrix(intRow, mconIntColҩƷ����) & "�ǲ��������ƿ�ҩƷ�������Զ��ֽ������ˣ�", vbInformation, gstrSysName
                                Exit Sub
                            End If
                        End If
                    Next
                End With
            End If
        End If
        
        If Not ��鵥��(6, txtNo, False) And Not mblnUpdate Then
            '�����µļ۸���µ����壬�˳���Ŀ�������û���һ�����յĵ���
            MsgBox "�м�¼δʹ�����¼۸񣬳����Զ���ɸ��£��ۼۡ��ɱ��ۡ��ۼ۽��ɱ�����ۣ������º����飡", vbInformation, gstrSysName
            Call RefreshBill
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
        
        '���������Ժ����Ź�ϵ������ҩƷ�ƿ����¼�����źͲ���
        If CheckBatchNum = False Then
            Exit Sub
        End If
        
        
        '�ж��Ƿ��Զ�ִ���ƿ����̣�����Ǿ��Զ���ɱ�ҩ�����͡����չ���
        If mint�ƿ⴦������ = 0 Then
            BlnSuccess = Auto�����ƿ�����
            Exit Sub
        End If

        
        'ִ�г�����˲���
        If Not SendPhysic Then Exit Sub
        If Not ҩƷ�������(Txt������.Caption) Then Exit Sub
        
        blnTrans = True
        gcnOracle.BeginTrans
        
        '��������¿�������Ϊ�٣���Ҫ�����ÿ���Ƿ��㹻
        If mbln�¿������� = False Then
            If Not CheckStock Then
                If blnTrans Then
                    gcnOracle.RollbackTrans
                End If
                Exit Sub
            End If
        End If
        
        '������ʱ�޸��˵��ݣ����������ɵ��ݱ���
        If mblnChange Then
            If Not SaveCard(True) Then
                gcnOracle.RollbackTrans: Exit Sub
            End If
            
            '��ҩ
            gstrSQL = "zl_ҩƷ�ƿ�_Prepare('" & txtNo.Tag & "','" & UserInfo.�û����� & "')"
            Call zlDataBase.ExecuteProcedure(gstrSQL, "��ҩ")
            '���ͣ��³���ⷿ��ҩƷ���ÿ�棩
            gstrSQL = "zl_ҩƷ�ƿ�_Prepare('" & txtNo.Tag & "')"
            Call zlDataBase.ExecuteProcedure(gstrSQL, "����")
        End If
        
        If Not SaveCheck(True) Then
            gcnOracle.RollbackTrans: Exit Sub
        End If

        gcnOracle.CommitTrans
        
        If Val(zlDataBase.GetPara("��˴�ӡ", glngSys, ģ���.ҩƷ�ƿ�)) = 1 Then
            '��ӡ
            If IsHavePrivs(mstrPrivs, "���ݴ�ӡ") Then
                printbill
            End If
        End If
        
        Unload Me
        Exit Sub
    End If
    
    If mint�༭״̬ = 6 Then '����
        If mblnChange = False And mint����ʽ <> 2 Then
            MsgBox "��¼�����������", vbInformation + vbOKOnly, gstrSysName
            Exit Sub
        End If
        
        If MsgBox("��ȷʵҪ����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            If SaveStrike = True Then
                If Val(zlDataBase.GetPara("��˴�ӡ", glngSys, ģ���.ҩƷ�ƿ�)) = 1 And mint����ʽ = 2 Then
                    '��ӡ
                    If IsHavePrivs(mstrPrivs, "���ݴ�ӡ") Then
                        printbill
                    End If
                End If
                Unload Me
            End If
        End If
        Exit Sub
    End If
    
    '�޸�״̬Ҫ����µ���
    If mint�༭״̬ = 2 Then
        If Not ��鵥��(6, txtNo, False) And Not mblnUpdate Then
            '�����µļ۸���µ����壬�˳���Ŀ�������û���һ�����յĵ���
            MsgBox "�м�¼δʹ�����¼۸񣬳����Զ���ɸ��£��ۼۡ��ɱ��ۡ��ۼ۽��ɱ�����ۣ������º����飡", vbInformation, gstrSysName
            Call RefreshBill
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
    End If
    
    If ValidData = False Then Exit Sub
    
    If mint�༭״̬ = 1 Then '��������ʱ���жϼ۸��Ƿ��Ѿ�����
        If ���۸� Then
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
    End If
    
    
    '�����ת���ƿ�����ĵ��ݣ�����������Ȩ�ޣ��򱣴浥�ݺ��Զ����
    If mint�༭״̬ = 11 And CmdSave.Caption = "���(&V)" Then
        blnTrans = True
        gcnOracle.BeginTrans
        
        '���浥��
        If Not SaveCard(True) Then
            gcnOracle.RollbackTrans
            Exit Sub
        End If
        
        mstr���ݺ� = txtNo.Tag
        
        'ִ��ִ���Զ���˲���
        If Not Auto�����ƿ�����(True) Then
            gcnOracle.RollbackTrans
            Exit Sub
        End If
        
        gcnOracle.CommitTrans
        blnTrans = True
        Unload Me
        Exit Sub
    End If
    
    BlnSuccess = SaveCard
    
    If BlnSuccess = True Then
        If Val(zlDataBase.GetPara("���̴�ӡ", glngSys, ģ���.ҩƷ�ƿ�)) = 1 Then
            '��ӡ
            If IsHavePrivs(mstrPrivs, "���ݴ�ӡ") Then
                printbill
            End If
        End If
        If mint�༭״̬ = 2 Then   '�޸�
            Unload Me
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
    mblnSave = False
    mblnEdit = True
    mshBill.ClearBill
    Call RefreshRowNO(mshBill, mconIntCol�к�, 1)
    
    txtժҪ.Text = ""
    cboEnterStock.SetFocus
    mblnChange = False
    If txtNo.Tag <> "" Then Me.staThis.Panels(2).Text = "��һ�ŵ��ݵ�NO�ţ�" & txtNo.Tag
    
    If mint�༭״̬ = 11 Then
        Unload Me
    End If
    
    Exit Sub
ErrHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmd�޿������ɸѡ_Click()
    PopupMenu mnuFilter, 2
End Sub

Private Sub Form_Activate()
    Debug.Print "����װ�أ�" & Now
    mblnChange = False
    Select Case mintParallelRecord
        Case 1
            '����
        Case 2
            If mint�༭״̬ = 6 Then
                MsgBox "�õ�����û�п��Գ�����ҩƷ�����飡", vbOKOnly, gstrSysName
            Else
                '�����ѱ�ɾ��
                MsgBox "�õ����ѱ�ɾ�������飡", vbOKOnly, gstrSysName
            End If
            Unload Me
            Exit Sub
        Case 3
            '�޸ĵĵ����ѱ����
            MsgBox "�õ����ѱ���������ˣ����飡", vbOKOnly, gstrSysName
            Unload Me
            Exit Sub
        Case 4
            '�������������
            MsgBox "�ÿⷿδ����ҩƷ������ƣ�", vbOKOnly, gstrSysName
            Unload Me
            Exit Sub
    End Select
    '��ʼ�����뷽ʽ
    If (mint�༭״̬ = 1 Or mint�༭״̬ = 2) And gbytSimpleCodeTrans = 1 Then
        staThis.Panels("PY").Visible = True
        staThis.Panels("WB").Visible = True
        gint���뷽ʽ = Val(zlDataBase.GetPara("���뷽ʽ", , , 0))    'Ĭ��ƴ������
        Logogram staThis, gint���뷽ʽ
    Else
        staThis.Panels("PY").Visible = False
        staThis.Panels("WB").Visible = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 70 Or KeyCode = 102 Then
        If Shift = vbCtrlMask Then   'Ctrl+F
            cmdFind_Click
        End If
    ElseIf KeyCode = vbKeyF3 Then
        FindRow mshBill, mconIntColҩ��, txtCode.Text, False
    ElseIf KeyCode = vbKeyF7 Then
        If staThis.Panels("PY").Bevel = sbrRaised Then
            Logogram staThis, 0
        Else
            Logogram staThis, 1
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim strStock As String
    Dim rsPara As New ADODB.Recordset
    
    mblnUpdate = False
    mblnEnterCell = False
    mintBatchNoLen = GetBatchNoLen()
    mintSelectStock = Val(zlDataBase.GetPara("�Ƿ�ѡ��ⷿ", glngSys, ģ���.ҩƷ�ƿ�))
    mintApplyType = -1
    mstrEndTime = ""
    mintDrugNameShow = Int(Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ�ƿ����", "ҩƷ������ʾ��ʽ", 0)))
    If mintDrugNameShow > 2 Or mintDrugNameShow < 0 Then mintDrugNameShow = 0
    mnuColDrug.Item(mintDrugNameShow).Checked = True
    mblnViewCost = IsHavePrivs(mstrPrivs, "�鿴�ɱ���")
    
    txtNo = mstr���ݺ�
    txtNo.Tag = mstr���ݺ�

    If mint�༭״̬ = 11 Then
        mlng����ⷿ = mfrmMain.cboEnterStock.ItemData(mfrmMain.cboEnterStock.ListIndex)
    End If
    
    'ȡϵͳ��������ȷ����ҩƷ���Ρ�
    mbln��ȷ���� = (gtype_UserSysParms.P174_ҩƷ�ƿ���ȷ���� = 1)
    
    mbln�¿������� = (gtype_UserSysParms.P96_ҩƷ��¿��ÿ�� = 1)
    
    '����ⷿȱʡΪ�����浱ǰѡ��Ŀⷿ������������Ч
    On Error Resume Next
    mlng����ⷿ = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
        
    Call GetDrugDigit(mlng����ⷿ, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    
    Call initCard
        
    mstrTime_Start = GetBillInfo(6, mstr���ݺ�)
    RestoreWinState Me, App.ProductName, MStrCaption
    
    '����ϵͳ��������ҩ����Ա�鿴����ʱ���Ƿ���ʾ�ɱ���
    mshBill.ColWidth(mconIntCol�ɹ���) = IIf(mblnViewCost, 1000, 0)
    mshBill.ColWidth(mconIntCol�ɹ����) = IIf(mblnViewCost, 900, 0)
    mshBill.ColWidth(mconintCol���) = IIf(mblnViewCost, 900, 0)
    mshBill.ColWidth(mconIntCol�ϴι�Ӧ��ID) = 0
    mshBill.ColWidth(mconintCol��ʵ����) = 0
    
    '��Ʒ���д���
    If gintҩƷ������ʾ = 2 Then
        '��ʾ��Ʒ����
        mshBill.ColWidth(mconIntCol��Ʒ��) = IIf(mshBill.ColWidth(mconIntCol��Ʒ��) = 0, 2000, mshBill.ColWidth(mconIntCol��Ʒ��))
    Else
        '��������ʾ��Ʒ����
        mshBill.ColWidth(mconIntCol��Ʒ��) = 0
    End If
    
    mint����� = MediWork_GetCheckStockRule(cboStock.ItemData(cboStock.ListIndex))
    mint��������ⷿ = MediWork_GetCheckStockRule(cboEnterStock.ItemData(cboEnterStock.ListIndex))
    
    mshBill.MsfObj.FixedCols = 4
    mshBill.CmdVisible = False
    mblnEnterCell = True
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim rsInitCard As New Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim strUnitQuantity As String
    Dim str���� As String
    Dim strArray As String
    Dim intRow As Integer
    Dim numUseAbleCount As Double
    Dim vardrug As Variant
    Dim strOrder As String, strCompare As String
    Dim lng���ⷿ As Long, lng����ⷿ As Long
    Dim intCostDigit As Integer        '�ɱ���С��λ��
    Dim intPriceDigit As Integer       '�ۼ�С��λ��
    Dim intNumberDigit As Integer      '����С��λ��
    Dim intMoneyDigit As Integer       '���С��λ��
    Dim strҩ�� As String
    Dim strSqlOrder As String
    Dim rsPrice As ADODB.Recordset
    
    '�ⷿ
    mbln���쵥 = False
    strOrder = zlDataBase.GetPara("����", glngSys, ģ���.ҩƷ�ƿ�)
    strCompare = Mid(strOrder, 1, 1)
    
    strSqlOrder = "���"
    
    If strCompare = "0" Then
        strSqlOrder = "���"
    ElseIf strCompare = "1" Then
        strSqlOrder = "ҩƷ����"
    ElseIf strCompare = "2" Then
        If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
            strSqlOrder = "ͨ����"
        Else
            strSqlOrder = "Nvl(��Ʒ��, ͨ����)"
        End If
    End If
    
    strSqlOrder = strSqlOrder & IIf(Right(strOrder, 1) = "0", " ASC", " DESC")
    
    On Error GoTo ErrHand
   
    'ȡָ�����ݵĳ���ⷿ�����ⷿ
    gstrSQL = " Select �ⷿID,�Է�����ID From ҩƷ�շ���¼" & _
              " Where NO=[1] And ����=6 And ���ϵ��=-1 And Rownum<2"
    Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[ȡָ�����ݵĳ���ⷿ�����ⷿ]", mstr���ݺ�)
              
    If rsInitCard.RecordCount <> 0 Then
        lng����ⷿ = rsInitCard!�ⷿid
        lng���ⷿ = rsInitCard!�Է�����id
        
        If lng����ⷿ > 0 Then
            mlng����ⷿ = lng����ⷿ
                
            Call GetDrugDigit(mlng����ⷿ, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
        End If
    Else
        lng����ⷿ = mlng����ⷿ
    End If
    
    intCostDigit = mintCostDigit
    intPriceDigit = mintPriceDigit
    intNumberDigit = mintNumberDigit
    intMoneyDigit = mintMoneyDigit
    
    If mint�༭״̬ <> 4 Then
        With mfrmMain.cboStock
            cboStock.Clear
            For i = 0 To .ListCount - 1
                cboStock.AddItem .List(i)
                cboStock.ItemData(cboStock.NewIndex) = .ItemData(i)
                If .ItemData(i) = lng����ⷿ Then cboStock.ListIndex = cboStock.ListCount - 1
            Next
            mintcboIndex = cboStock.ListIndex
            '���û��ָ����ҩ�����������
            If mintcboIndex = -1 Then
                gstrSQL = "Select ID,���� From ���ű� Where ID=[1] "
                Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[���û��ָ����ҩ�����������]", lng����ⷿ)
                
                cboStock.AddItem rsInitCard!����
                cboStock.ItemData(cboStock.NewIndex) = rsInitCard!id
                cboStock.ListIndex = cboStock.ListCount - 1
            End If
            mintcboIndex = cboStock.ListIndex
            cboStock.Enabled = .Enabled
        End With
    End If
    
    Select Case mint�༭״̬
        Case 1
            Txt������ = UserInfo.�û�����
            Txt�������� = Format(zlDataBase.Currentdate, "yyyy-mm-dd hh:mm:ss")
            initGrid
            
            If cboEnterStock.ListCount <> 0 Then
                If cboStock.ItemData(cboStock.ListIndex) = cboEnterStock.ItemData(cboEnterStock.ListIndex) Then
                    If cboEnterStock.ListCount > 1 Then
                        cboEnterStock.ListIndex = cboEnterStock.ListIndex + 1
                    End If
                End If
            Else
                mintParallelRecord = 4
                Exit Sub
            End If
        Case 2, 3, 4, 6, 10, 11
            initGrid
            '���õ����Ƿ������쵥��
            gstrSQL = " Select Nvl(��ҩ��ʽ,0) ���� From ҩƷ�շ���¼ " & _
                      " Where ����=6 And NO=[1] And ���ϵ�� = -1 and rownum = 1"
            Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[���õ����Ƿ������쵥��]", mstr���ݺ�)
                      
            If Not rsTemp.EOF Then
                mbln���쵥 = (rsTemp!���� = 1)
                If mbln���쵥 Then LblTitle.Caption = "ҩƷ���쵥"
            End If
            
            If mint�༭״̬ = 4 Then
                gstrSQL = "select distinct b.id,b.���� from ҩƷ�շ���¼ a,���ű� b where a.�ⷿid=b.id and A.���� = 6 and a.no=[1] and a.���ϵ��=-1"
                Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, mstr���ݺ�)
                
                If rsInitCard.EOF Then
                    mintParallelRecord = 2
                    Exit Sub
                End If
                
                With cboStock
                    .AddItem rsInitCard!����
                    .ItemData(.NewIndex) = rsInitCard!id
                    .ListIndex = 0
                End With
                rsInitCard.Close
            End If
            
            Select Case mintUnit
                Case mconint�ۼ۵�λ
                    strUnitQuantity = "B.�ͻ���λ,B.�ͻ���װ,C.���㵥λ AS ��λ, A.��д����,a.ʵ������,a.�ɱ���,a.���ۼ�,'1' as ����ϵ��,"
                Case mconint���ﵥλ
                    strUnitQuantity = "B.�ͻ���λ,B.�ͻ���װ,B.���ﵥλ AS ��λ,(A.��д���� / B.�����װ) AS ��д����,(A.ʵ������ / B.�����װ) AS ʵ������,a.�ɱ���*B.�����װ as �ɱ���,a.���ۼ�*B.�����װ as ���ۼ�,B.�����װ as ����ϵ��,"
                Case mconintסԺ��λ
                    strUnitQuantity = "B.�ͻ���λ,B.�ͻ���װ,B.סԺ��λ AS ��λ,(A.��д���� / B.סԺ��װ) AS ��д����,(A.ʵ������ / B.סԺ��װ) AS ʵ������,a.�ɱ���*B.סԺ��װ as �ɱ���,a.���ۼ�*B.סԺ��װ as ���ۼ�,B.סԺ��װ as ����ϵ��,"
                Case mconintҩ�ⵥλ
                    strUnitQuantity = "B.�ͻ���λ,B.�ͻ���װ,B.ҩ�ⵥλ AS ��λ,(A.��д���� / B.ҩ���װ) AS ��д����,(A.ʵ������ / B.ҩ���װ) AS ʵ������,a.�ɱ���*B.ҩ���װ as �ɱ���,a.���ۼ�*B.ҩ���װ as ���ۼ�,B.ҩ���װ as ����ϵ��,"
            End Select
            
            Select Case mint�༭״̬
            Case 6
                '��������
                If mint����ʽ <> 2 Then
                    gstrSQL = "SELECT W.*,Z.��������/W.����ϵ�� AS  ��������,Z.ʵ�ʽ��,Z.ʵ�ʲ�� " & _
                        " FROM " & _
                        "     (SELECT DISTINCT A.ҩƷID,A.���,'[' || C.���� || ']' As ҩƷ����, C.���� As ͨ����, E.���� As ��Ʒ��," & _
                        "     B.ҩƷ��Դ,B.����ҩ��,C.���,C.���� AS ԭ����,A.����, A.����,A.����,B.ָ�������,B.ҩ����� AS ��������," & _
                        "     B.���Ч��,A.Ч��," & strUnitQuantity & _
                        "     A.�ɱ����,0 ���۽��, 0 ���,D.ժҪ,A.�ⷿID,A.�Է�����ID,C.�Ƿ���,B.ҩ������ AS ҩ����������,A.�ϴι�Ӧ��ID,A.��׼�ĺ�,A.��д���� ��ʵ���� " & _
                        "     FROM " & _
                        "         (SELECT MIN(ID) AS ID, SUM(ʵ������) AS ��д����,0 ʵ������,SUM(�ɱ����) AS �ɱ����,ҩƷID,���,����, ����,Ч��,NVL(����,0) ����,����,�ɱ���,���ۼ�,�ⷿID,�Է�����ID,������ID,NVL(��ҩ��λID,0) �ϴι�Ӧ��ID,��׼�ĺ�" & _
                        "          FROM ҩƷ�շ���¼ X " & _
                        "          WHERE NO=[1] AND ����=6 AND ���ϵ��=-1 " & _
                        "          GROUP BY ҩƷID,���,����,����,Ч��,NVL(����,0),����,�ɱ���,���ۼ�,�ⷿID,�Է�����ID,������ID,NVL(��ҩ��λID,0),��׼�ĺ�" & _
                        "          HAVING SUM(ʵ������)<>0 ) A," & _
                        "     ҩƷ��� B,�շ���ĿĿ¼ C,�շ���Ŀ���� E, " & _
                        " (Select ���, ժҪ From ҩƷ�շ���¼ " & _
                        "  Where ���� = 6 And NO = [1] And ���ϵ�� = -1 And (��¼״̬ = 1 Or Mod(��¼״̬, 3) = 0)) D " & _
                        "     WHERE A.ҩƷID = B.ҩƷID AND B.ҩƷID=E.�շ�ϸĿID(+) AND E.����(+)=3 AND B.ҩƷID=C.ID And A.��� = D.���) W," & _
                        "     (SELECT  ҩƷID,NVL(����,0) ����,��������,ʵ�ʽ��,ʵ�ʲ�� " & _
                        "     FROM ҩƷ��� WHERE �ⷿID=[2] AND ����=1) Z " & _
                        " WHERE W.ҩƷID=Z.ҩƷID(+) AND NVL(W.����,0)=Nvl(Z.����(+),0) " & _
                        " ORDER BY " & strSqlOrder
                Else
                    '������˳���ʱ����ʾδ��˵������������
                    gstrSQL = "SELECT W.*,Z.��������/W.����ϵ�� AS  ��������,Z.ʵ�ʽ��,Z.ʵ�ʲ�� " & _
                        " FROM " & _
                        "     (SELECT DISTINCT A.ҩƷID,A.���,'[' || C.���� || ']' As ҩƷ����, C.���� As ͨ����, E.���� As ��Ʒ��," & _
                        "     B.ҩƷ��Դ,B.����ҩ��,C.���,C.���� AS ԭ����,A.����, A.����,A.����,B.ָ�������,B.ҩ����� AS ��������," & _
                        "     B.���Ч��,A.Ч��," & strUnitQuantity & "A.�ɱ����,A.���۽��, A.���,A.��ҩ��, " & _
                        "     A.ժҪ,������,��������,�����,�������,A.�ⷿID,A.�Է�����ID,C.�Ƿ���,B.ҩ������ AS ҩ����������,NVL(A.��ҩ��λID,0) �ϴι�Ӧ��ID,A.��׼�ĺ�,A.��д���� ��ʵ���� " & _
                        "     FROM ҩƷ�շ���¼ A, ҩƷ��� B,�շ���ĿĿ¼ C,�շ���Ŀ���� E " & _
                        "     WHERE A.ҩƷID = B.ҩƷID AND B.ҩƷID=C.ID AND B.ҩƷID=E.�շ�ϸĿID(+) AND E.����(+)=3 " & _
                        "     AND A.��¼״̬ =[3] " & _
                        "     AND A.���� = 6 AND A.���ϵ��=-1 AND A.NO =[1] ) W," & _
                        "     (SELECT  ҩƷID,NVL(����,0) ����,��������,ʵ�ʽ��,ʵ�ʲ�� " & _
                        "     FROM ҩƷ��� WHERE �ⷿID=[2] AND ����=1) Z " & _
                        " WHERE W.ҩƷID=Z.ҩƷID(+) AND NVL(W.����,0)=Nvl(Z.����(+),0) " & _
                        " ORDER BY " & strSqlOrder
                End If
            Case 11
                gstrSQL = "SELECT W.*,Z.��������/W.����ϵ�� AS  ��������,Z.ʵ�ʽ��,Z.ʵ�ʲ��,z.ƽ���ɱ��� * w.����ϵ�� As �ɱ���," & _
                    " z.���ۼ�*w.����ϵ�� as ���ۼ�, w.ʵ������ * z.ƽ���ɱ��� * w.����ϵ�� As �ɱ����,z.ʵ������/w.����ϵ�� as ������� " & _
                    " FROM " & _
                    "     (SELECT DISTINCT A.ҩƷID,A.���,'[' || C.���� || ']' As ҩƷ����, C.���� As ͨ����, E.���� As ��Ʒ��," & _
                    "     B.ҩƷ��Դ,B.����ҩ��,C.���,C.���� AS ԭ����,A.����, A.����,A.����,B.ָ�������,B.ҩ����� AS ��������," & _
                    "     B.���Ч��,A.Ч��," & strUnitQuantity & "A.�ɱ����,A.���۽��, A.���,A.��ҩ��,A.��д���� ��ʵ����, " & _
                    "     A.ժҪ,������,��������,�����,�������,A.�ⷿID," & cboEnterStock.ItemData(cboEnterStock.ListIndex) & " �Է�����ID,C.�Ƿ���,B.ҩ������ AS ҩ����������,NVL(A.��ҩ��λID,0) �ϴι�Ӧ��ID,A.��׼�ĺ� " & _
                    "     FROM ҩƷ�շ���¼ A, ҩƷ��� B,�շ���ĿĿ¼ C,�շ���Ŀ���� E " & _
                    "     WHERE A.ҩƷID = B.ҩƷID AND B.ҩƷID=C.ID AND B.ҩƷID=E.�շ�ϸĿID(+) AND E.����(+)=3 " & _
                    "     AND A.��¼״̬ =[3] " & _
                    "     AND A.���� = 1 AND A.NO = [1] And A.����� Is Not Null) W," & _
                    "     (SELECT  ҩƷID,NVL(����,0) ����,��������,ʵ������,ʵ�ʽ��,ʵ�ʲ��,ƽ���ɱ���,nvl(���ۼ�,0) as ���ۼ�  " & _
                    "     FROM ҩƷ��� WHERE �ⷿID=[2] AND ����=1) Z, " & _
                    "     (Select Distinct �շ�ϸĿid From �շ�ִ�п��� f Where ִ�п���ID=[4] ) Y " & _
                    " WHERE W.ҩƷID=Z.ҩƷID(+) AND W.ҩƷid=Y.�շ�ϸĿid AND NVL(W.����,0)=Nvl(Z.����(+),0) " & _
                    " ORDER BY " & strSqlOrder
            Case Else
                gstrSQL = "SELECT W.*,Z.��������/W.����ϵ�� AS  ��������,Z.ʵ�ʽ��,Z.ʵ�ʲ�� " & _
                    " FROM " & _
                    "     (SELECT DISTINCT A.ҩƷID,A.���,'[' || C.���� || ']' As ҩƷ����, C.���� As ͨ����, E.���� As ��Ʒ��," & _
                    "     B.ҩƷ��Դ,B.����ҩ��,C.���,C.���� AS ԭ����,A.����, A.����,A.����,B.ָ�������,B.ҩ����� AS ��������," & _
                    "     B.���Ч��,A.Ч��," & strUnitQuantity & "A.�ɱ����,A.���۽��, A.���,A.��ҩ��,Nvl(A.����,-1) As ���췽ʽ,A.Ƶ�� As ����ʱ��, " & _
                    "     A.ժҪ,������,��������,�����,�������,A.�ⷿID,A.�Է�����ID,C.�Ƿ���,B.ҩ������ AS ҩ����������,NVL(A.��ҩ��λID,0) �ϴι�Ӧ��ID,A.��׼�ĺ�,A.��д���� ��ʵ���� " & _
                    "     FROM ҩƷ�շ���¼ A, ҩƷ��� B,�շ���ĿĿ¼ C,�շ���Ŀ���� E " & _
                    "     WHERE A.ҩƷID = B.ҩƷID AND B.ҩƷID=C.ID AND B.ҩƷID=E.�շ�ϸĿID(+) AND E.����(+)=3 " & _
                    "     AND A.��¼״̬ =[3] " & _
                    "     AND A.���� = 6 AND A.���ϵ��=-1 AND A.NO =[1] ) W," & _
                    "     (SELECT  ҩƷID,NVL(����,0) ����,��������,ʵ�ʽ��,ʵ�ʲ�� " & _
                    "     FROM ҩƷ��� WHERE �ⷿID=[2] AND ����=1) Z " & _
                    " WHERE W.ҩƷID=Z.ҩƷID(+) AND NVL(W.����,0)=Nvl(Z.����(+),0) " & _
                    " ORDER BY " & strSqlOrder
            End Select

            Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, IIf(mint�༭״̬ = 11, mstr��ⵥ��, mstr���ݺ�), cboStock.ItemData(cboStock.ListIndex), mint��¼״̬, cboEnterStock.ItemData(cboEnterStock.ListIndex))
                        
            If rsInitCard.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            
            Select Case mint�༭״̬
            Case 2, 6, 10, 11
                Txt������ = UserInfo.�û�����
                Txt�������� = Format(zlDataBase.Currentdate, "yyyy-mm-dd hh:mm:ss")
                If mint�༭״̬ = 6 Or mint�༭״̬ = 10 Then
                    Txt����� = UserInfo.�û�����
                    Txt������� = Format(zlDataBase.Currentdate, "yyyy-mm-dd hh:mm:ss")
                End If
                If mint�༭״̬ = 10 Then
                    Txt����� = Nvl(rsInitCard!��ҩ��)
                    Txt������ = rsInitCard!������
                    Txt�������� = Format(rsInitCard!��������, "yyyy-mm-dd hh:mm:ss")
                    Lbl�����.Caption = "��ҩ��"
                    Lbl�������.Caption = "��������"
                End If
            Case Else
                Txt������ = rsInitCard!������
                Txt�������� = Format(rsInitCard!��������, "yyyy-mm-dd hh:mm:ss")
                Txt����� = IIf(IsNull(rsInitCard!�����), "", rsInitCard!�����)
                Txt������� = IIf(IsNull(rsInitCard!�������), "", Format(rsInitCard!�������, "yyyy-mm-dd hh:mm:ss"))
            End Select
            txtժҪ.Text = IIf(IsNull(rsInitCard!ժҪ), "", rsInitCard!ժҪ)
            
            If (mint�༭״̬ = 2 Or mint�༭״̬ = 3) And Txt����� <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            If mint�༭״̬ = 2 Or mint�༭״̬ = 3 Or mint�༭״̬ = 10 Then
                mintApplyType = rsInitCard!���췽ʽ
                mstrEndTime = Nvl(rsInitCard!����ʱ��)
            End If
            
            Dim intCount As Integer
            With cboEnterStock
                For intCount = 0 To .ListCount - 1
                    If .ItemData(intCount) = lng���ⷿ Then
                        .ListIndex = intCount
                        .Tag = intCount
                        Exit For
                    End If
                Next
            End With
            
            If mint�༭״̬ = 2 Or mint�༭״̬ = 3 Or mint�༭״̬ = 10 Then
                Set mcolUsedCount = New Collection
            End If
            
            With mshBill
                Do While Not rsInitCard.EOF
                    intRow = rsInitCard.AbsolutePosition
                    .rows = intRow + 1
                    .TextMatrix(intRow, 0) = rsInitCard.Fields(0)
                    
                    If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
                        strҩ�� = rsInitCard!ͨ����
                    Else
                        strҩ�� = IIf(IsNull(rsInitCard!��Ʒ��), rsInitCard!ͨ����, rsInitCard!��Ʒ��)
                    End If
                    
                    .TextMatrix(intRow, mconIntColҩƷ���������) = rsInitCard!ҩƷ���� & strҩ��
                    .TextMatrix(intRow, mconIntColҩƷ����) = rsInitCard!ҩƷ����
                    .TextMatrix(intRow, mconIntColҩƷ����) = strҩ��
                    
                    If mintDrugNameShow = 1 Then
                        .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ����)
                    ElseIf mintDrugNameShow = 2 Then
                        .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ����)
                    Else
                        .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ���������)
                    End If
                    
                    .TextMatrix(intRow, mconIntCol��Ʒ��) = IIf(IsNull(rsInitCard!��Ʒ��), "", rsInitCard!��Ʒ��)

                    .TextMatrix(intRow, mconIntCol��Դ) = Nvl(rsInitCard!ҩƷ��Դ)
                    .TextMatrix(intRow, mconIntCol����ҩ��) = Nvl(rsInitCard!����ҩ��)
                    If mint�༭״̬ <> 11 Then .TextMatrix(intRow, mconIntCol���) = rsInitCard!���
                    .TextMatrix(intRow, mconIntCol���) = IIf(IsNull(rsInitCard!���), "", rsInitCard!���)
                    .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsInitCard!����), "", rsInitCard!����)
                    .TextMatrix(intRow, mconIntCol��λ) = rsInitCard!��λ
                    If IIf(IsNull(rsInitCard!����ϵ��), 0, rsInitCard!����ϵ��) = 0 Or Nvl(rsInitCard!�ͻ���װ) = "" Or Nvl(rsInitCard!�ͻ���λ) = "" Then
                        .TextMatrix(intRow, mconIntCol�ͻ���λ) = ""
                    Else
                        .TextMatrix(intRow, mconIntCol�ͻ���λ) = rsInitCard!�ͻ���λ & "(1" & rsInitCard!�ͻ���λ & "=" & zlStr.FormatEx(rsInitCard!�ͻ���װ / rsInitCard!����ϵ��, 1, , True) & rsInitCard!��λ & ")"
                    End If
                    .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsInitCard!����), "", rsInitCard!����)
                    .TextMatrix(intRow, mconIntColЧ��) = IIf(IsNull(rsInitCard!Ч��), "", Format(rsInitCard!Ч��, "yyyy-mm-dd"))
                    If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And .TextMatrix(intRow, mconIntColЧ��) <> "" Then
                        '����Ϊ��Ч��
                        .TextMatrix(intRow, mconIntColЧ��) = Format(DateAdd("D", -1, .TextMatrix(intRow, mconIntColЧ��)), "yyyy-mm-dd")
                    End If
                    
                    .TextMatrix(intRow, mconIntCol��д����) = GetFormat(IIf(mint�༭״̬ = 6 And mint����ʽ = 2, -1, 1) * rsInitCard!��д����, intNumberDigit)
                    .TextMatrix(intRow, mconIntColʵ������) = GetFormat(IIf(mint�༭״̬ = 6 And mint����ʽ = 2, -1, 1) * rsInitCard!ʵ������, intNumberDigit)
                    
                    If mint�༭״̬ = 6 Then
                        .TextMatrix(intRow, mconIntCol�ɹ���) = GetFormat(rsInitCard!�ɱ���, intCostDigit)
                        If Val(rsInitCard!��д����) <> 0 And Val(.TextMatrix(intRow, mconIntCol�ɹ���)) = 0 Then
                            .TextMatrix(intRow, mconIntCol�ɹ���) = GetFormat((rsInitCard!���۽�� - rsInitCard!���) / Val(rsInitCard!��д����), intCostDigit)
                        End If
                    Else
                        .TextMatrix(intRow, mconIntCol�ɹ���) = GetFormat(rsInitCard!�ɱ���, intCostDigit)
                        If Val(rsInitCard!ʵ������) <> 0 And Val(.TextMatrix(intRow, mconIntCol�ɹ���)) = 0 Then
                            .TextMatrix(intRow, mconIntCol�ɹ���) = GetFormat((rsInitCard!���۽�� - rsInitCard!���) / Val(rsInitCard!ʵ������), intCostDigit)
                        End If
                    End If
                    .TextMatrix(intRow, mconIntCol�ɹ����) = GetFormat(IIf(mint�༭״̬ = 6 And mint����ʽ <> 2, 0, IIf(mint�༭״̬ = 6 And mint����ʽ = 2, -1, 1) * rsInitCard!�ɱ����), intMoneyDigit)
                    
                    If mint�༭״̬ = 11 Then
                        If rsInitCard!�Ƿ��� = 0 Then
                            gstrSQL = "Select �ּ� From �շѼ�Ŀ Where �շ�ϸĿid = [1] And Sysdate Between ִ������ And ��ֹ����"
                            
                            Set rsPrice = zlDataBase.OpenSQLRecord(gstrSQL, "��ѯ�۸�", rsInitCard!ҩƷid)
                            .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(rsPrice!�ּ� * rsInitCard!����ϵ��, intPriceDigit, , True)
                            .TextMatrix(intRow, mconIntCol�ۼ۽��) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol�ۼ�)) * rsInitCard!ʵ������, intMoneyDigit, , True)
                        Else
                            'ʱ��
                            If rsInitCard!���ۼ� = 0 Then
                                If rsInitCard!������� <> 0 Then
                                    .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(rsInitCard!ʵ�ʽ�� / rsInitCard!�������, intPriceDigit, , True)
                                Else
                                    gstrSQL = "Select �ּ� From �շѼ�Ŀ Where �շ�ϸĿid = [1] And Sysdate Between ִ������ And ��ֹ����"
                                    
                                    Set rsPrice = zlDataBase.OpenSQLRecord(gstrSQL, "��ѯ�۸�", rsInitCard!ҩƷid)
                                    .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(rsPrice!�ּ� * rsInitCard!����ϵ��, intPriceDigit, , True)
                                End If
                            Else
                                .TextMatrix(intRow, mconIntCol�ۼ�) = zlStr.FormatEx(rsInitCard!���ۼ�, intPriceDigit, , True)
                            End If
                            .TextMatrix(intRow, mconIntCol�ۼ۽��) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol�ۼ�)) * rsInitCard!ʵ������, intMoneyDigit, , True)
                        End If
                        .TextMatrix(intRow, mconintCol���) = zlStr.FormatEx(Val(.TextMatrix(intRow, mconIntCol�ۼ۽��)) - Val(.TextMatrix(intRow, mconIntCol�ɹ����)), intMoneyDigit, , True)
                    Else
                        .TextMatrix(intRow, mconIntCol�ۼ�) = GetFormat(rsInitCard!���ۼ�, intPriceDigit)
                        .TextMatrix(intRow, mconIntCol�ۼ۽��) = GetFormat(IIf(mint�༭״̬ = 6 And mint����ʽ = 2, -1, 1) * rsInitCard!���۽��, intMoneyDigit)
                        .TextMatrix(intRow, mconintCol���) = GetFormat(IIf(mint�༭״̬ = 6 And mint����ʽ = 2, -1, 1) * rsInitCard!���, intMoneyDigit)
                    End If
                    
                    .TextMatrix(intRow, mconIntCol���Ч��) = IIf(IsNull(rsInitCard!���Ч��), "0", rsInitCard!���Ч��) & "||" & rsInitCard!�Ƿ��� & "||" & rsInitCard!ҩ����������
                    .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsInitCard!����), "0", rsInitCard!����)
                    .TextMatrix(intRow, mconIntCol����ϵ��) = rsInitCard!����ϵ��
                    .TextMatrix(intRow, mconIntColָ�������) = rsInitCard!ָ�������
                    .TextMatrix(intRow, mconIntCol��������) = IIf(IsNull(rsInitCard!��������), "0", rsInitCard!��������)
                    .TextMatrix(intRow, mconIntCol��������) = GetFormat(IIf(IsNull(rsInitCard!��������), "0", rsInitCard!��������), intNumberDigit)
                    .TextMatrix(intRow, mconIntColʵ�ʲ��) = IIf(IsNull(rsInitCard!ʵ�ʲ��), "0", rsInitCard!ʵ�ʲ��)
                    .TextMatrix(intRow, mconIntColʵ�ʽ��) = IIf(IsNull(rsInitCard!ʵ�ʽ��), "0", rsInitCard!ʵ�ʽ��)
                    .TextMatrix(intRow, mconIntCol�ϴι�Ӧ��ID) = rsInitCard!�ϴι�Ӧ��ID
                    .TextMatrix(intRow, mconIntCol��׼�ĺ�) = IIf(IsNull(rsInitCard!��׼�ĺ�), "", rsInitCard!��׼�ĺ�)
                    .TextMatrix(intRow, mconIntCol��������) = IIf(GetҩƷ��������(Val(.TextMatrix(intRow, 0)), cboEnterStock.ItemData(cboEnterStock.ListIndex)) = True, "1", 0)
                
'                    If (mint�༭״̬ = 3 Or mint�༭״̬ = 10) And Nvl(rsInitCard!��������, 0) = 1 And Nvl(rsInitCard!����, 0) = 0 And mbln�Զ��ֽ�δ��� = False Then
'                        mbln�Զ��ֽ�δ��� = True
'                    End If

                    If (mint�༭״̬ = 3 Or mint�༭״̬ = 10) And mbln�Զ��ֽ�δ��� = False Then
                        If Get��������(Val(cboStock.ItemData(cboStock.ListIndex)), Val(.TextMatrix(intRow, 0))) = 1 And Nvl(rsInitCard!����, 0) = 0 Then
                            mbln�Զ��ֽ�δ��� = True
                        End If
                    End If
                    
                    Call ��ʾ�����(intRow)
                                        
                    If mint�༭״̬ = 2 Or mint�༭״̬ = 6 Or mint�༭״̬ = 3 Or mint�༭״̬ = 10 Or mint�༭״̬ = 11 Then
                        .TextMatrix(intRow, mconintCol��ʵ����) = IIf(mint�༭״̬ = 6 And mint����ʽ = 2, -1, 1) * rsInitCard!��ʵ����
                    End If
                    If mint�༭״̬ = 2 Or mint�༭״̬ = 3 Or mint�༭״̬ = 10 Then
                        numUseAbleCount = 0
                        For Each vardrug In mcolUsedCount
                            If vardrug(0) = CStr(rsInitCard!ҩƷid & IIf(IsNull(rsInitCard!����), "0", rsInitCard!����)) Then
                                numUseAbleCount = vardrug(1)
                                mcolUsedCount.Remove vardrug(0)
                                Exit For
                            End If
                        Next
                        str���� = rsInitCard!ҩƷid & IIf(IsNull(rsInitCard!����), "0", rsInitCard!����)
                        If mint�༭״̬ = 2 Then
                            strArray = numUseAbleCount + IIf(IsNull(rsInitCard!��д����), "0", rsInitCard!��д����)
                        Else
                            strArray = numUseAbleCount + IIf(IsNull(rsInitCard!ʵ������), "0", rsInitCard!ʵ������)
                        End If
                        mcolUsedCount.Add Array(str����, strArray), str����
                    End If
                    rsInitCard.MoveNext
                Loop
                .rows = intRow + 2
            End With
            rsInitCard.Close
    End Select
    Call RefreshRowNO(mshBill, mconIntCol�к�, 1)
    
    SetEdit         '���ñ༭����
    '�޸Ļ����ʱ���Զ��޿����ɸѡ
    If (mint�༭״̬ = 2 Or mint�༭״̬ = 3 Or mint�༭״̬ = 10) Then
        If mint�ƿ⴦������ = 1 And mint�༭״̬ = 3 Then
            cmd�޿������ɸѡ.Visible = False
        Else
            cmd�޿������ɸѡ.Visible = True
        End If
        
    End If
    '���ġ��޸Ļ����ʱ�����ݿ��������������ʾ����
    If (mint�༭״̬ = 2 Or mint�༭״̬ = 3 Or mint�༭״̬ = 4 Or mint�༭״̬ = 10) Then
'        If mbln���쵥 Then
        Call ShowColor
        Select Case mint�༭״̬
        Case 10 '2, 10 �޸�ʱ�Զ��ֽⲻ�ɼ�
            cmdExpend.Visible = True
        End Select
    End If
    If mint�ƿ⴦������ = 0 And mint�༭״̬ = 3 Then
        cmdExpend.Visible = True
    End If
    Call ��ʾ�ϼƽ��
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetEdit()
    Dim intCol As Integer
    
    With mshBill
        If mblnEdit = False Then
            For intCol = 0 To .Cols - 1
                .ColData(intCol) = 0
            Next
            cboStock.Enabled = False
            cboEnterStock.Enabled = False
            txtժҪ.Enabled = (mint�༭״̬ = 6)
            
            If mint�༭״̬ = 10 Or (mint�༭״̬ = 6 And mint����ʽ <> 2) Then
                .ColData(mconIntColʵ������) = 4
            End If
        Else
            .ColData(0) = 5
            .ColData(mconIntColҩ��) = 1
            .ColData(mconIntCol���) = 5
            .ColData(mconIntCol���) = 5
            .ColData(mconIntCol����) = 5
            .ColData(mconIntCol��λ) = 5
            .ColData(mconIntCol�ͻ���λ) = 5
            .ColData(mconIntCol����) = 5
            .ColData(mconIntColЧ��) = 5
            .ColData(mconIntCol��׼�ĺ�) = 5
            If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
                .ColData(mconIntCol��д����) = 4
                .ColData(mconIntColʵ������) = 5
            ElseIf mint�༭״̬ = 3 Then
                .ColData(mconIntCol��д����) = 5
                .ColData(mconIntColʵ������) = 4
            ElseIf mint�༭״̬ = 11 Then
                If mint�ƿ⴦������ = 1 Then
                    .ColData(mconIntCol��д����) = 4
                    .ColData(mconIntColʵ������) = 5
                Else
                    .ColData(mconIntCol��д����) = 5
                    .ColData(mconIntColʵ������) = 4
                End If
            End If
            .ColData(mconIntCol�ɹ���) = 5
            .ColData(mconIntCol�ɹ����) = 5
            .ColData(mconIntCol�ۼ�) = 5
            .ColData(mconIntCol�ۼ۽��) = 5
            .ColData(mconintCol���) = 5
            
            .ColData(mconIntCol��������) = 5
            .ColData(mconIntCol��������) = 5
            .ColData(mconIntCol���Ч��) = 5
            
            .ColData(mconIntColָ�������) = 5
            .ColData(mconIntColʵ�ʽ��) = 5
            .ColData(mconIntColʵ�ʲ��) = 5
            .ColData(mconIntCol����ϵ��) = 5
            .ColData(mconIntCol����) = 5
                     
            .ColAlignment(mconIntColҩ��) = flexAlignLeftCenter
            .ColAlignment(mconIntCol���) = flexAlignLeftCenter
            .ColAlignment(mconIntCol����) = flexAlignLeftCenter
            .ColAlignment(mconIntCol��λ) = flexAlignCenterCenter
            .ColAlignment(mconIntCol�ͻ���λ) = flexAlignCenterCenter
            .ColAlignment(mconIntCol����) = flexAlignLeftCenter
            .ColAlignment(mconIntColЧ��) = flexAlignLeftCenter
            .ColAlignment(mconIntCol��д����) = flexAlignRightCenter
            .ColAlignment(mconIntColʵ������) = flexAlignRightCenter
            
            .ColAlignment(mconIntCol�ɹ���) = flexAlignRightCenter
            .ColAlignment(mconIntCol�ɹ����) = flexAlignRightCenter
            .ColAlignment(mconIntCol�ۼ�) = flexAlignRightCenter
            .ColAlignment(mconIntCol�ۼ۽��) = flexAlignRightCenter
            .ColAlignment(mconintCol���) = flexAlignRightCenter
            
            If mintSelectStock = 0 Then
                cboStock.Enabled = False
            Else
                cboStock.Enabled = True
            End If
            
            cboEnterStock.Enabled = True
            txtժҪ.Enabled = True
        End If
    End With
End Sub


Private Sub initGrid()
    With mshBill
        .Active = True
        .Cols = mconIntColS
        .MsfObj.FixedCols = 4
        
        .TextMatrix(0, mconIntCol�к�) = ""
        .TextMatrix(0, mconIntColҩ��) = "ҩƷ���������"
        .TextMatrix(0, mconIntCol��Ʒ��) = "��Ʒ��"
        .TextMatrix(0, mconIntCol��Դ) = "ҩƷ��Դ"
        .TextMatrix(0, mconIntCol����ҩ��) = "����ҩ��"
        .TextMatrix(0, mconIntCol���) = "���"
        .TextMatrix(0, mconIntCol���) = "���"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntCol��λ) = "��λ"
        .TextMatrix(0, mconIntCol�ͻ���λ) = "��λ"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntColЧ��) = IIf(gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1, "��Ч����", "ʧЧ��")
        .TextMatrix(0, mconIntCol��׼�ĺ�) = "��׼�ĺ�"
        .TextMatrix(0, mconIntCol�ⷿ���) = "�ⷿ���"
        .TextMatrix(0, mconIntCol�Է����) = "�Է����"
        .TextMatrix(0, mconIntCol��д����) = IIf(mint�༭״̬ = 6, "����", "��д����")
        .TextMatrix(0, mconIntColʵ������) = IIf(mint�༭״̬ = 6, "��������", "ʵ������")
        
        .TextMatrix(0, mconIntCol�ɹ���) = "�ɱ���"
        .TextMatrix(0, mconIntCol�ɹ����) = "�ɱ����"
        .TextMatrix(0, mconIntCol�ۼ�) = "�ۼ�"
        .TextMatrix(0, mconIntCol�ۼ۽��) = "�ۼ۽��"
        .TextMatrix(0, mconintCol���) = "���"
        
        .TextMatrix(0, mconIntCol��������) = "��������"
        .TextMatrix(0, mconIntCol��������) = "��������"
        .TextMatrix(0, mconIntCol���Ч��) = "���Ч��"
        .TextMatrix(0, mconIntColʵ�ʲ��) = "ʵ�ʲ��"
        .TextMatrix(0, mconIntColʵ�ʽ��) = "ʵ�ʽ��"
        .TextMatrix(0, mconIntColָ�������) = "ָ�������"
        .TextMatrix(0, mconIntCol����ϵ��) = "����ϵ��"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntCol�ϴι�Ӧ��ID) = "�ϴι�Ӧ��ID"
        .TextMatrix(0, mconintCol��ʵ����) = "��ʵ����"
        .TextMatrix(0, mconIntColҩƷ���������) = "ҩƷ���������"
        .TextMatrix(0, mconIntColҩƷ����) = "ҩƷ����"
        .TextMatrix(0, mconIntColҩƷ����) = "ҩƷ����"
        .TextMatrix(0, mconIntCol��������) = "��������"
        
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mconIntCol�к�) = "1"
        
        .ColWidth(0) = 0
        .ColWidth(mconIntCol�к�) = 300
        .ColWidth(mconIntColҩ��) = 2000
        .ColWidth(mconIntCol��Ʒ��) = 2000
        .ColWidth(mconIntCol��Դ) = 900
        .ColWidth(mconIntCol����ҩ��) = 900
        .ColWidth(mconIntCol���) = 0
        .ColWidth(mconIntCol���) = 900
        .ColWidth(mconIntCol����) = 800
        .ColWidth(mconIntCol��λ) = 500
        .ColWidth(mconIntCol�ͻ���λ) = 2000
        .ColWidth(mconIntCol����) = 800
        .ColWidth(mconIntColЧ��) = 1000
        .ColWidth(mconIntCol��д����) = 1100
        .ColWidth(mconIntColʵ������) = 1100
        .ColWidth(mconIntCol�ɹ���) = 1000
        .ColWidth(mconIntCol�ɹ����) = 900
        .ColWidth(mconIntCol�ۼ�) = 1000
        .ColWidth(mconIntCol�ۼ۽��) = 900
        .ColWidth(mconintCol���) = 800
        
        .ColWidth(mconIntCol��������) = 0
        .ColWidth(mconIntCol��������) = 0
        .ColWidth(mconIntCol���Ч��) = 0
        .ColWidth(mconIntColʵ�ʲ��) = 0
        .ColWidth(mconIntColʵ�ʽ��) = 0
        .ColWidth(mconIntColָ�������) = 0
        .ColWidth(mconIntCol����ϵ��) = 0
        .ColWidth(mconIntCol����) = 0
        .ColWidth(mconIntCol�ϴι�Ӧ��ID) = 0
        .ColWidth(mconintCol��ʵ����) = 0
        
        .ColWidth(mconIntColҩƷ���������) = 0
        .ColWidth(mconIntColҩƷ����) = 0
        .ColWidth(mconIntColҩƷ����) = 0
        .ColWidth(mconIntCol��������) = 0
        
        '-1����ʾ���п���ѡ���ǲ����ͣ�"��"��" "��
        ' 0����ʾ���п���ѡ�񣬵������޸�
        ' 1����ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
        ' 2����ʾ�����������У��ⲿ��ʾΪ��ťѡ�񣬵���������ѡ���
        ' 3����ʾ������ѡ���У��ⲿ��ʾΪ������ѡ��
        '4:  ��ʾ����Ϊ�������ı����û�����
        '5:  ��ʾ���в�����ѡ��

        .ColData(0) = 5
        .ColData(mconIntCol��Ʒ��) = 5
        .ColData(mconIntCol��Դ) = 5
        .ColData(mconIntCol����ҩ��) = 5
        .ColData(mconIntCol�к�) = 5
        .ColData(mconIntCol���) = 5
        .ColData(mconIntCol���) = 5
        .ColData(mconIntCol����) = 5
        .ColData(mconIntCol��λ) = 5
        .ColData(mconIntCol�ͻ���λ) = 5
        .ColData(mconIntCol����) = 5
        .ColData(mconIntColЧ��) = 5
        .ColData(mconIntCol��׼�ĺ�) = 5
        .ColData(mconIntCol�ⷿ���) = 5
        .ColData(mconIntCol�Է����) = 5
        .ColData(mconIntColҩƷ���������) = 5
        .ColData(mconIntColҩƷ����) = 5
        .ColData(mconIntColҩƷ����) = 5
        .ColData(mconIntCol��������) = 5
        
        If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Then
            cboEnterStock.Enabled = True
            txtժҪ.Enabled = True
            
            If mintSelectStock = 0 Then
                cboStock.Enabled = False
            Else
                cboStock.Enabled = True
            End If
            
            .ColData(mconIntColҩ��) = 1
            .ColData(mconIntCol��д����) = 4
            .ColData(mconIntColʵ������) = 5
        ElseIf mint�༭״̬ = 3 Then
            cboStock.Enabled = False
            cboEnterStock.Enabled = False
            txtժҪ.Enabled = False
            
            .ColData(mconIntColҩ��) = 5
            .ColData(mconIntCol��д����) = 5
            .ColData(mconIntColʵ������) = 4
        ElseIf mint�༭״̬ = 6 Then
            cboStock.Enabled = False
            cboEnterStock.Enabled = False
            txtժҪ.Enabled = True
            
            .ColData(mconIntColҩ��) = 5
            .ColData(mconIntCol��д����) = 5
            .ColData(mconIntColʵ������) = 5
                
            If mint����ʽ <> 2 Then
                .ColData(mconIntColʵ������) = 4
            End If
        ElseIf mint�༭״̬ = 4 Then
            cboStock.Enabled = False
            cboEnterStock.Enabled = False
            txtժҪ.Enabled = False
            
            .ColData(mconIntCol��д����) = 5
            .ColData(mconIntColʵ������) = 5
            .ColData(mconIntColҩ��) = 5
        ElseIf mint�༭״̬ = 11 Then
            cboStock.Enabled = False
            cboEnterStock.Enabled = True
            txtժҪ.Enabled = True
            
            If mint�ƿ⴦������ = 1 Then
                .ColData(mconIntCol��д����) = 4
                .ColData(mconIntColʵ������) = 5
            Else
                .ColData(mconIntCol��д����) = 5
                .ColData(mconIntColʵ������) = 4
            End If
            .ColData(mconIntColҩ��) = 5
        End If
        
        .ColData(mconIntCol�ɹ���) = 5
        .ColData(mconIntCol�ɹ����) = 5
        .ColData(mconIntCol�ۼ�) = 5
        .ColData(mconIntCol�ۼ۽��) = 5
        .ColData(mconintCol���) = 5
        
        .ColData(mconIntCol��������) = 5
        .ColData(mconIntCol��������) = 5
        .ColData(mconIntCol���Ч��) = 5
        .ColData(mconIntColʵ�ʲ��) = 5
        .ColData(mconIntColʵ�ʽ��) = 5
        .ColData(mconIntColָ�������) = 5
        .ColData(mconIntCol����ϵ��) = 5
        .ColData(mconIntCol����) = 5
        .ColData(mconIntCol�ϴι�Ӧ��ID) = 5
        .ColData(mconintCol��ʵ����) = 5
        
        .ColAlignment(mconIntColҩ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��Ʒ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol���) = flexAlignLeftCenter
        .ColAlignment(mconIntCol����) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��λ) = flexAlignCenterCenter
        .ColAlignment(mconIntCol�ͻ���λ) = flexAlignCenterCenter
        .ColAlignment(mconIntCol����) = flexAlignLeftCenter
        .ColAlignment(mconIntColЧ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��׼�ĺ�) = flexAlignLeftCenter
        .ColAlignment(mconIntCol�ⷿ���) = flexAlignRightCenter
        .ColAlignment(mconIntCol�Է����) = flexAlignRightCenter
        .ColAlignment(mconIntCol��д����) = flexAlignRightCenter
        .ColAlignment(mconIntColʵ������) = flexAlignRightCenter
        .ColAlignment(mconintCol��ʵ����) = flexAlignRightCenter
        
        .ColAlignment(mconIntCol�ɹ���) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ɹ����) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ۼ�) = flexAlignRightCenter
        .ColAlignment(mconIntCol�ۼ۽��) = flexAlignRightCenter
        .ColAlignment(mconintCol���) = flexAlignRightCenter
        
        
        .PrimaryCol = mconIntColҩ��
        .LocateCol = mconIntColҩ��
        If InStr(1, "346", mint�༭״̬) <> 0 Then .ColData(mconIntColҩ��) = 0
    End With
    txtժҪ.MaxLength = GetLength("ҩƷ�շ���¼", "ժҪ")
    chkIn.Visible = (mint�༭״̬ = 1)
    txtIn.Visible = (mint�༭״̬ = 1)
End Sub

Private Function CheckBatchNum() As Boolean
    '���ܣ�����������ҩƷ�����Ƿ�Ϊ��
    '����ֵ��true-����ҩƷ�������Σ�false-����ҩƷ��������Ϊ�����
    Dim intRow As Integer
    
    With mshBill
        If .rows > 1 Then
            For intRow = 1 To .rows - 1
                If .TextMatrix(intRow, mconIntCol��������) = "1" And .TextMatrix(intRow, mconIntCol����) = "" And .TextMatrix(intRow, 0) <> "" Then
                    CheckBatchNum = False
                    MsgBox "��" & intRow & "�У����ⷿ�Ƿ�����������¼�����ţ�", vbInformation, gstrSysName
                    .SetFocus
                    .Row = intRow
                    Exit Function
                End If
            Next
            CheckBatchNum = True
        Else
            CheckBatchNum = True
        End If
    End With
End Function

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.WindowState = vbMinimized Then Exit Sub
   
    With Pic����
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - IIf(staThis.Visible, staThis.Height, 0) - .Top - 100 - CmdCancel.Height - 200
    End With
    
    With LblTitle
        .Left = 0
        .Top = 150
        .Width = Pic����.Width
    End With
    
    
    With mshBill
        .Left = 200
        .Width = Pic����.Width - .Left * 2
    End With
    With txtNo
        .Left = mshBill.Left + mshBill.Width - .Width
        LblNo.Left = .Left - LblNo.Width - 100
        .Top = LblTitle.Top
        LblNo.Top = .Top
    End With
    
    
    LblStock.Left = mshBill.Left
    cboStock.Left = LblStock.Left + LblStock.Width + 100
    
    cboEnterStock.Left = mshBill.Left + mshBill.Width - cboEnterStock.Width
    
    LblEnterStock.Left = cboEnterStock.Left - LblEnterStock.Width - 100
    
    
    With Lbl������
        .Top = Pic����.Height - 200 - .Height
        .Left = mshBill.Left + 100
    End With
    
    With Txt������
        .Top = Lbl������.Top - 80
        .Left = Lbl������.Left + Lbl������.Width + 100
    End With
    
    With Lbl��������
        .Top = Lbl������.Top
        .Left = Txt������.Left + Txt������.Width + 250
    End With
    
    With Txt��������
        .Top = Lbl��������.Top - 80
        .Left = Lbl��������.Left + Lbl��������.Width + 100
    End With
    
    With Txt�������
        .Top = Lbl������.Top - 80
        .Left = mshBill.Left + mshBill.Width - .Width
    End With
    
    With Lbl�������
        .Top = Lbl������.Top
        .Left = Txt�������.Left - 100 - .Width
    End With
    
    With Txt�����
        .Top = Lbl������.Top - 80
        .Left = Lbl�������.Left - 200 - .Width
    End With
    
    With Lbl�����
        .Top = Lbl������.Top
        .Left = Txt�����.Left - 100 - .Width
    End With
    
    With txtժҪ
        .Top = Lbl������.Top - 140 - .Height
        .Left = Txt������.Left
        .Width = mshBill.Left + mshBill.Width - .Left
    End With
    
    With lblժҪ
        .Top = txtժҪ.Top + 50
        .Left = txtժҪ.Left - .Width - 100
        '.Width = .Left - .Left
        Debug.Print .Width
    End With
    
    
    With lblPurchasePrice
        .Left = mshBill.Left
        .Top = txtժҪ.Top - 60 - .Height
        .Width = mshBill.Width
        lblSalePrice.Top = .Top
        lblDifference.Top = .Top
    End With
    
    With lblSalePrice
        .Left = lblPurchasePrice.Left + mshBill.Width / 3
    End With
    With lblDifference
        .Left = lblPurchasePrice.Left + mshBill.Width / 3 * 2
    End With
    If mblnViewCost = False Then
        lblPurchasePrice.Visible = False
        lblDifference.Visible = False
    End If
    
    With mshBill
        .Height = lblPurchasePrice.Top - .Top - 60
    End With
    
    With CmdCancel
        .Left = Pic����.Left + mshBill.Left + mshBill.Width - .Width
        .Top = Pic����.Top + Pic����.Height + 100
    End With
    
    With CmdSave
        .Left = CmdCancel.Left - .Width - 100
        .Top = CmdCancel.Top
    End With
    
    With cmdAllCls
        .Left = CmdSave.Left - .Width - 500
        .Top = CmdCancel.Top
    End With
    
    With cmdAllSel
        .Left = cmdAllCls.Left - .Width - 100
        .Top = CmdCancel.Top
    End With
    
    With cmdHelp
        .Left = Pic����.Left + mshBill.Left
        .Top = CmdCancel.Top
    End With
        
    With cmdFind
        .Top = CmdCancel.Top
    End With
    
    With cmdExpend
        .Top = CmdSave.Top
        .Left = CmdSave.Left - 150 - .Width
    End With

    With cmd�޿������ɸѡ
        .Top = CmdSave.Top
        .Left = CmdSave.Left - 150 - .Width - cmdExpend.Width - 100
    End With
    
    With lblCode
        .Top = CmdCancel.Top + 50
    End With
    With txtCode
        .Top = CmdCancel.Top + 30
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\ҩƷ�ƿ����", "ҩƷ������ʾ��ʽ", mintDrugNameShow)
    
    If mblnChange = False Or mint�༭״̬ = 4 Or mint�༭״̬ = 3 Then
        SaveWinState Me, App.ProductName, MStrCaption
        Call ReleaseSelectorRS  'ж�����ݼ�
        mblnRS = False
        zlPlugIn_Unload mobjPlugIn
        Exit Sub
    End If
    If MsgBox("���ݿ����Ѹı䣬��δ���̣���Ҫ�˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
        Exit Sub
    Else
        SaveWinState Me, App.ProductName, MStrCaption
    End If
    Call ReleaseSelectorRS  'ж�����ݼ�
    mblnRS = False
    zlPlugIn_Unload mobjPlugIn
End Sub

Private Function SaveCheck(Optional ByVal blnǿ�Ʊ��� As Boolean = False) As Boolean
    Dim rs��� As New Recordset
    Dim intRow As Integer
    
    Dim strNo As String
    Dim lng�ⷿID As Long
    Dim lng�Է�����id As Long
    Dim str����� As String
    
    Dim lngҩƷID As Long
    Dim str���� As String
    Dim lng������ As Long
    Dim num��д���� As Double
    Dim numʵ������ As Double
    Dim num�ɱ��� As Double
    Dim num�ɱ���� As Double
    Dim dbl�ۼ� As Double
    Dim num���۽�� As Double
    Dim num��� As Double
    Dim lng�����id As Long
    Dim lng�����id As Long
    Dim str���� As String
    Dim datЧ�� As String
    Dim dat������� As String
    Dim int���к� As Integer
    Dim lng�ϴι�Ӧ��ID As Long
    Dim str��׼�ĺ� As String
        
    Dim arrSql As Variant
    Dim n As Integer
    
    arrSql = Array()
    mblnSave = False
    SaveCheck = False
    
    '���õ����Ƿ��ڽ���༭����󣬱���������Ա�޸ģ���������ת���ƿⵥ�ݣ��򲻼��
    If mint�༭״̬ <> 11 Then
        mstrTime_End = GetBillInfo(6, mstr���ݺ�)
        If mstrTime_End = "" Then
            MsgBox "�õ����Ѿ�����������Աɾ����", vbInformation, gstrSysName
            Exit Function
        End If
   
        If mint�ƿ⴦������ <> 0 Then
            If mstrTime_End > mstrTime_Start Then
                MsgBox "�õ����Ѿ�����������Ա�༭�����˳������ԣ�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    lng�ⷿID = cboStock.ItemData(cboStock.ListIndex)
    lng�Է�����id = cboEnterStock.ItemData(cboEnterStock.ListIndex)
    str����� = UserInfo.�û�����
    strNo = txtNo.Tag
    
    gstrSQL = "SELECT b.ϵ��,b.id AS ���id " _
            & "FROM ҩƷ�������� a, ҩƷ������ b " _
            & "Where a.���id = b.ID AND a.���� = 6 "
    Set rs��� = zlDataBase.OpenSQLRecord(gstrSQL, "ҩƷ�ƿ����")
    
    If rs���.EOF Then
        MsgBox "�Բ���ҩƷ������಻ȫ������!", vbExclamation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If rs���.RecordCount < 2 Then
        MsgBox "�Բ���ҩƷ������಻ȫ������!", vbExclamation + vbOKOnly, gstrSysName
        Exit Function
    End If
    rs���.MoveFirst
    Do While Not rs���.EOF
        If rs���!ϵ�� = 1 Then
            lng�����id = rs���!���id
        Else
            lng�����id = rs���!���id
        End If
        rs���.MoveNext
    Loop
    rs���.Close
    
    dat������� = Format(zlDataBase.Currentdate, "yyyy-mm-dd hh:mm:ss")
    With mshBill
        On Error GoTo errHandle
        
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
            If .TextMatrix(intRow, 0) <> "" Then
                lngҩƷID = .TextMatrix(intRow, 0)
                str���� = .TextMatrix(intRow, mconIntCol����)
                lng������ = .TextMatrix(intRow, mconIntCol����)
                
                If .TextMatrix(intRow, mconIntCol��д����) = .TextMatrix(intRow, mconIntColʵ������) Then
                    num��д���� = .TextMatrix(intRow, mconintCol��ʵ����)
                    numʵ������ = .TextMatrix(intRow, mconintCol��ʵ����)
                Else
                    num��д���� = GetFormat(Val(.TextMatrix(intRow, mconIntCol��д����)) * .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_����)
                    numʵ������ = GetFormat(Val(.TextMatrix(intRow, mconIntColʵ������)) * .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_����)
                End If
                
'                num�ɱ��� = GetFormat(Val(.TextMatrix(intRow, mconIntCol�ɹ���)) / .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_�ɱ���)
                num�ɱ��� = Get�ɱ���(lngҩƷID, lng�ⷿID, lng������)
                num�ɱ���� = Val(.TextMatrix(intRow, mconIntCol�ɹ����))
'                dbl�ۼ� = GetFormat(Val(.TextMatrix(intRow, mconIntCol�ۼ�)) / .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_���ۼ�)
                dbl�ۼ� = Get�ۼ�(Split(.TextMatrix(intRow, mconIntCol���Ч��), "||")(1) = 1, lngҩƷID, lng�ⷿID, lng������)
                num���۽�� = Val(.TextMatrix(intRow, mconIntCol�ۼ۽��))
                num��� = Val(.TextMatrix(intRow, mconintCol���))
                str���� = .TextMatrix(intRow, mconIntCol����)
                datЧ�� = IIf(.TextMatrix(intRow, mconIntColЧ��) = "", "", .TextMatrix(intRow, mconIntColЧ��))
                If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And datЧ�� <> "" Then
                    '����ΪʧЧ��������
                    datЧ�� = Format(DateAdd("D", 1, datЧ��), "yyyy-mm-dd")
                End If
                
                If mint�༭״̬ = 11 And CmdSave.Caption = "���(&V)" Then
                    '������ֱ�������ˣ��������ʱ�������ʵ��Ӧ���ǵ������������Ǳʵ���ţ�
                    int���к� = 2 * intRow - 1 '2 * Val(.TextMatrix(intRow, mconIntCol���)) - 1
                Else
                    int���к� = Val(.TextMatrix(intRow, mconIntCol���))
                End If
                
                lng�ϴι�Ӧ��ID = .TextMatrix(intRow, mconIntCol�ϴι�Ӧ��ID)
                str��׼�ĺ� = IIf(Trim(.TextMatrix(intRow, mconIntCol��׼�ĺ�)) = "", "", .TextMatrix(intRow, mconIntCol��׼�ĺ�))
                
                gstrSQL = "zl_ҩƷ�ƿ�_Verify("
                '���
                gstrSQL = gstrSQL & int���к�
                '�ⷿID
                gstrSQL = gstrSQL & "," & lng�ⷿID
                '�Է�����ID
                gstrSQL = gstrSQL & "," & lng�Է�����id
                'ҩƷID
                gstrSQL = gstrSQL & "," & lngҩƷID
                '����
                gstrSQL = gstrSQL & ",'" & str���� & "'"
                '������
                gstrSQL = gstrSQL & "," & lng������
                'ʵ������
                gstrSQL = gstrSQL & "," & numʵ������
                '�ɱ���
                gstrSQL = gstrSQL & "," & num�ɱ���
                '�ɱ����
                gstrSQL = gstrSQL & "," & num�ɱ����
                '���۽��
                gstrSQL = gstrSQL & "," & num���۽��
                '���
                gstrSQL = gstrSQL & "," & num���
                'NO
                gstrSQL = gstrSQL & ",'" & strNo & "'"
                '�����
                gstrSQL = gstrSQL & ",'" & str����� & "'"
                '����
                gstrSQL = gstrSQL & ",'" & str���� & "'"
                'Ч��
                gstrSQL = gstrSQL & "," & IIf(datЧ�� = "", "Null", "to_date('" & Format(datЧ��, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                '�������
                gstrSQL = gstrSQL & ",to_date('" & dat������� & "','yyyy-mm-dd HH24:MI:SS')"
                '��ҩ��λID
                gstrSQL = gstrSQL & "," & IIf(lng�ϴι�Ӧ��ID = 0, "NULL", lng�ϴι�Ӧ��ID)
                '��׼�ĺ�
                gstrSQL = gstrSQL & ",'" & str��׼�ĺ� & "'"
                '�ۼ�
                gstrSQL = gstrSQL & "," & dbl�ۼ�
                gstrSQL = gstrSQL & ")"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = CStr(lngҩƷID) & ";" & gstrSQL
            End If
            recSort.MoveNext
        Next
    End With
    
    If Not ExecuteSql(arrSql, MStrCaption, False, Not blnǿ�Ʊ���) Then Exit Function

    If Not blnǿ�Ʊ��� Then gcnOracle.CommitTrans
    
    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    
    '��ҹ���
    Call zlPlugIn_Ini(glngSys, glngModul, mobjPlugIn)
    Call CallPlugInDrugStuffWork(mobjPlugIn, 3, lng�ⷿID, strNo, ���ݺ�.ҩƷ�ƿ�)
    
    Exit Function
errHandle:
    If Not blnǿ�Ʊ��� Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function

Private Sub mnuColDrug_Click(Index As Integer)
    Dim n As Integer
    
    With mnuColDrug
        For n = 0 To .count - 1
            .Item(n).Checked = False
        Next
        
        .Item(Index).Checked = True
        
        Call SetDrugName(Index)
    End With
End Sub

Private Sub mnuFilterDrug_Click(Index As Integer)
    
    If Index = 1 Then
        If MsgBox("��ȷʵҪɾ��ʵ������Ϊ0��ҩƷ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    Call MyAppend
    Call AddAppend(Index)
    With mrsMyAppend
        mshBill.ClearBill
        mshBill.rows = 2
        
        If Not .EOF Then .MoveFirst
        Do While Not .EOF
            mshBill.TextMatrix(mshBill.rows - 1, 0) = .Fields!ҩƷid
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol�к�) = mshBill.rows - 1
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol���) = (mshBill.rows - 2) * 2 + 1
            mshBill.TextMatrix(mshBill.rows - 1, mconIntColҩ��) = .Fields!ҩ��
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol��Ʒ��) = .Fields!��Ʒ��
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol��Դ) = .Fields!��Դ
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol����ҩ��) = .Fields!����ҩ��
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol���) = .Fields!���
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol��������) = .Fields!��������
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol���Ч��) = .Fields!���Ч��
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol��������) = zlStr.FormatEx(.Fields!��������, mintNumberDigit, , True)
            mshBill.TextMatrix(mshBill.rows - 1, mconIntColָ�������) = .Fields!ָ�������
            mshBill.TextMatrix(mshBill.rows - 1, mconIntColʵ�ʽ��) = .Fields!ʵ�ʽ��
            mshBill.TextMatrix(mshBill.rows - 1, mconIntColʵ�ʲ��) = .Fields!ʵ�ʲ��
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol����ϵ��) = .Fields!����ϵ��
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol����) = .Fields!����
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol����) = .Fields!����
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol��λ) = .Fields!��λ
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol�ͻ���λ) = .Fields!�ͻ���λ
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol����) = .Fields!����
            mshBill.TextMatrix(mshBill.rows - 1, mconIntColЧ��) = .Fields!Ч��
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol��׼�ĺ�) = .Fields!��׼�ĺ�
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol�ⷿ���) = zlStr.FormatEx(.Fields!�ⷿ���, mintNumberDigit, , True)
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol�Է����) = zlStr.FormatEx(.Fields!�Է����, mintNumberDigit, , True)
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol��д����) = zlStr.FormatEx(.Fields!��д����, mintNumberDigit, , True)
            mshBill.TextMatrix(mshBill.rows - 1, mconIntColʵ������) = zlStr.FormatEx(.Fields!ʵ������, mintNumberDigit, , True)
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol�ɹ���) = zlStr.FormatEx(.Fields!�ɹ���, mintCostDigit, , True)
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol�ɹ����) = zlStr.FormatEx(.Fields!�ɹ����, mintMoneyDigit, , True)
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol�ۼ�) = zlStr.FormatEx(.Fields!�ۼ�, mintPriceDigit, , True)
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol�ۼ۽��) = zlStr.FormatEx(.Fields!�ۼ۽��, mintMoneyDigit, , True)
            mshBill.TextMatrix(mshBill.rows - 1, mconintCol���) = zlStr.FormatEx(.Fields!���, mintMoneyDigit, , True)
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol�ϴι�Ӧ��ID) = .Fields!�ϴι�Ӧ��ID
            mshBill.TextMatrix(mshBill.rows - 1, mconintCol��ʵ����) = zlStr.FormatEx(.Fields!��ʵ����, mintNumberDigit, , True)
            mshBill.TextMatrix(mshBill.rows - 1, mconIntColҩƷ���������) = .Fields!ҩƷ���������
            mshBill.TextMatrix(mshBill.rows - 1, mconIntColҩƷ����) = .Fields!ҩƷ����
            mshBill.TextMatrix(mshBill.rows - 1, mconIntColҩƷ����) = .Fields!ҩƷ����
            mshBill.TextMatrix(mshBill.rows - 1, mconIntCol��������) = .Fields!��������
            
            mshBill.rows = mshBill.rows + 1
            .MoveNext
        Loop
        
        mshBill.Row = mshBill.rows - 1
    End With
    
    Call ShowColor
End Sub

Private Sub MyAppend()
    '������̬��¼��
    Set mrsMyAppend = New ADODB.Recordset
    With mrsMyAppend
        If .State = 1 Then .Close
        .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
        .Fields.Append "ҩ��", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "��Ʒ��", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "��Դ", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "����ҩ��", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "���", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "��������", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "���Ч��", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "��������", adDouble, 18, adFldIsNullable
        .Fields.Append "ָ�������", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "ʵ�ʽ��", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "ʵ�ʲ��", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "����ϵ��", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "��λ", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "�ͻ���λ", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "����", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "Ч��", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "��׼�ĺ�", adLongVarChar, 40, adFldIsNullable
        .Fields.Append "�ⷿ���", adDouble, 18, adFldIsNullable
        .Fields.Append "�Է����", adDouble, 18, adFldIsNullable
        .Fields.Append "��д����", adDouble, 18, adFldIsNullable
        .Fields.Append "ʵ������", adDouble, 18, adFldIsNullable
        .Fields.Append "�ɹ���", adDouble, 18, adFldIsNullable
        .Fields.Append "�ɹ����", adDouble, 18, adFldIsNullable
        .Fields.Append "�ۼ�", adDouble, 18, adFldIsNullable
        .Fields.Append "�ۼ۽��", adDouble, 18, adFldIsNullable
        .Fields.Append "���", adDouble, 18, adFldIsNullable
        .Fields.Append "�ϴι�Ӧ��ID", adDouble, 18, adFldIsNullable
        .Fields.Append "��ʵ����", adDouble, 18, adFldIsNullable
        .Fields.Append "ҩƷ���������", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "ҩƷ����", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "ҩƷ����", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "��������", adLongVarChar, 40, adFldIsNullable
    
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub AddAppend(ByVal Index As Integer)
    '����̬��¼������ֵ
    Dim i As Integer
    On Error GoTo ErrHand

    With mrsMyAppend
        For i = 1 To mshBill.rows - 1
            If mshBill.TextMatrix(i, 0) <> "" And Val(mshBill.TextMatrix(i, mconIntColʵ������)) <> 0 Then
                .AddNew
                .Fields!ҩƷid = mshBill.TextMatrix(i, 0)
                .Fields!ҩ�� = mshBill.TextMatrix(i, mconIntColҩ��)
                .Fields!��Ʒ�� = mshBill.TextMatrix(i, mconIntCol��Ʒ��)
                .Fields!��Դ = mshBill.TextMatrix(i, mconIntCol��Դ)
                .Fields!����ҩ�� = mshBill.TextMatrix(i, mconIntCol����ҩ��)
                .Fields!��� = mshBill.TextMatrix(i, mconIntCol���)
                .Fields!�������� = mshBill.TextMatrix(i, mconIntCol��������)
                .Fields!���Ч�� = mshBill.TextMatrix(i, mconIntCol���Ч��)
                .Fields!�������� = mshBill.TextMatrix(i, mconIntCol��������)
                .Fields!ָ������� = mshBill.TextMatrix(i, mconIntColָ�������)
                .Fields!ʵ�ʽ�� = mshBill.TextMatrix(i, mconIntColʵ�ʽ��)
                .Fields!ʵ�ʲ�� = mshBill.TextMatrix(i, mconIntColʵ�ʲ��)
                .Fields!����ϵ�� = mshBill.TextMatrix(i, mconIntCol����ϵ��)
                .Fields!���� = mshBill.TextMatrix(i, mconIntCol����)
                .Fields!���� = mshBill.TextMatrix(i, mconIntCol����)
                .Fields!��λ = mshBill.TextMatrix(i, mconIntCol��λ)
                .Fields!�ͻ���λ = mshBill.TextMatrix(i, mconIntCol�ͻ���λ)
                .Fields!���� = mshBill.TextMatrix(i, mconIntCol����)
                .Fields!Ч�� = mshBill.TextMatrix(i, mconIntColЧ��)
                .Fields!��׼�ĺ� = mshBill.TextMatrix(i, mconIntCol��׼�ĺ�)
                .Fields!�ⷿ��� = mshBill.TextMatrix(i, mconIntCol�ⷿ���)
                .Fields!�Է���� = mshBill.TextMatrix(i, mconIntCol�Է����)
                .Fields!��д���� = IIf(mshBill.TextMatrix(i, mconIntCol��д����) = "", 0, mshBill.TextMatrix(i, mconIntCol��д����))
                .Fields!ʵ������ = IIf(mshBill.TextMatrix(i, mconIntColʵ������) = "", 0, mshBill.TextMatrix(i, mconIntColʵ������))
                .Fields!�ɹ��� = mshBill.TextMatrix(i, mconIntCol�ɹ���)
                .Fields!�ɹ���� = IIf(mshBill.TextMatrix(i, mconIntCol�ɹ����) = "", 0, mshBill.TextMatrix(i, mconIntCol�ɹ����))
                .Fields!�ۼ� = mshBill.TextMatrix(i, mconIntCol�ۼ�)
                .Fields!�ۼ۽�� = IIf(mshBill.TextMatrix(i, mconIntCol�ۼ۽��) = "", 0, mshBill.TextMatrix(i, mconIntCol�ۼ۽��))
                .Fields!��� = IIf(mshBill.TextMatrix(i, mconintCol���) = "", 0, mshBill.TextMatrix(i, mconintCol���))
                .Fields!�ϴι�Ӧ��ID = mshBill.TextMatrix(i, mconIntCol�ϴι�Ӧ��ID)
                .Fields!��ʵ���� = IIf(mshBill.TextMatrix(i, mconintCol��ʵ����) = "", 0, mshBill.TextMatrix(i, mconintCol��ʵ����))
                .Fields!ҩƷ��������� = mshBill.TextMatrix(i, mconIntColҩƷ���������)
                .Fields!ҩƷ���� = mshBill.TextMatrix(i, mconIntColҩƷ����)
                .Fields!ҩƷ���� = mshBill.TextMatrix(i, mconIntColҩƷ����)
                .Fields!�������� = mshBill.TextMatrix(i, mconIntCol��������)
                .Update
            End If
        Next
    
        For i = 1 To mshBill.rows - 1
            If mshBill.TextMatrix(i, 0) <> "" And Index = 0 And Val(mshBill.TextMatrix(i, mconIntColʵ������)) = 0 Then
                .AddNew
                .Fields!ҩƷid = mshBill.TextMatrix(i, 0)
                .Fields!ҩ�� = mshBill.TextMatrix(i, mconIntColҩ��)
                .Fields!��Ʒ�� = mshBill.TextMatrix(i, mconIntCol��Ʒ��)
                .Fields!��Դ = mshBill.TextMatrix(i, mconIntCol��Դ)
                .Fields!����ҩ�� = mshBill.TextMatrix(i, mconIntCol����ҩ��)
                .Fields!��� = mshBill.TextMatrix(i, mconIntCol���)
                .Fields!�������� = mshBill.TextMatrix(i, mconIntCol��������)
                .Fields!���Ч�� = mshBill.TextMatrix(i, mconIntCol���Ч��)
                .Fields!�������� = mshBill.TextMatrix(i, mconIntCol��������)
                .Fields!ָ������� = mshBill.TextMatrix(i, mconIntColָ�������)
                .Fields!ʵ�ʽ�� = mshBill.TextMatrix(i, mconIntColʵ�ʽ��)
                .Fields!ʵ�ʲ�� = mshBill.TextMatrix(i, mconIntColʵ�ʲ��)
                .Fields!����ϵ�� = mshBill.TextMatrix(i, mconIntCol����ϵ��)
                .Fields!���� = mshBill.TextMatrix(i, mconIntCol����)
                .Fields!���� = mshBill.TextMatrix(i, mconIntCol����)
                .Fields!��λ = mshBill.TextMatrix(i, mconIntCol��λ)
                .Fields!�ͻ���λ = mshBill.TextMatrix(i, mconIntCol�ͻ���λ)
                .Fields!���� = mshBill.TextMatrix(i, mconIntCol����)
                .Fields!Ч�� = mshBill.TextMatrix(i, mconIntColЧ��)
                .Fields!��׼�ĺ� = mshBill.TextMatrix(i, mconIntCol��׼�ĺ�)
                .Fields!�ⷿ��� = mshBill.TextMatrix(i, mconIntCol�ⷿ���)
                .Fields!�Է���� = mshBill.TextMatrix(i, mconIntCol�Է����)
                .Fields!��д���� = IIf(mshBill.TextMatrix(i, mconIntCol��д����) = "", 0, mshBill.TextMatrix(i, mconIntCol��д����))
                .Fields!ʵ������ = IIf(mshBill.TextMatrix(i, mconIntColʵ������) = "", 0, mshBill.TextMatrix(i, mconIntColʵ������))
                .Fields!�ɹ��� = mshBill.TextMatrix(i, mconIntCol�ɹ���)
                .Fields!�ɹ���� = IIf(mshBill.TextMatrix(i, mconIntCol�ɹ����) = "", 0, mshBill.TextMatrix(i, mconIntCol�ɹ����))
                .Fields!�ۼ� = mshBill.TextMatrix(i, mconIntCol�ۼ�)
                .Fields!�ۼ۽�� = IIf(mshBill.TextMatrix(i, mconIntCol�ۼ۽��) = "", 0, mshBill.TextMatrix(i, mconIntCol�ۼ۽��))
                .Fields!��� = IIf(mshBill.TextMatrix(i, mconintCol���) = "", 0, mshBill.TextMatrix(i, mconintCol���))
                .Fields!�ϴι�Ӧ��ID = mshBill.TextMatrix(i, mconIntCol�ϴι�Ӧ��ID)
                .Fields!��ʵ���� = IIf(mshBill.TextMatrix(i, mconintCol��ʵ����) = "", 0, mshBill.TextMatrix(i, mconintCol��ʵ����))
                .Fields!ҩƷ��������� = mshBill.TextMatrix(i, mconIntColҩƷ���������)
                .Fields!ҩƷ���� = mshBill.TextMatrix(i, mconIntColҩƷ����)
                .Fields!ҩƷ���� = mshBill.TextMatrix(i, mconIntColҩƷ����)
                .Fields!�������� = mshBill.TextMatrix(i, mconIntCol��������)
                .Update
            End If
        Next
    End With
       
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mshBill_AfterAddRow(Row As Long)
    Call RefreshRowNO(mshBill, mconIntCol�к�, Row)
End Sub

Private Sub mshBill_AfterDeleteRow()
    Call ��ʾ�ϼƽ��
    Call RefreshRowNO(mshBill, mconIntCol�к�, mshBill.Row)
End Sub

Private Sub mshBill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    If mint�༭״̬ = 10 Then
        Cancel = True
        Exit Sub
    End If
    If InStr(1, "34", mint�༭״̬) <> 0 Then
        If mint�༭״̬ = 3 And mbln���쵥 Then Exit Sub
        Cancel = True
        Exit Sub
    End If
    With mshBill
        If .TextMatrix(.Row, 0) <> "" Then
            If MsgBox("��ȷʵҪɾ������ҩƷ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True
            End If
        End If
    End With
End Sub

Private Sub mshbill_CommandClick()
    Dim i As Integer
    Dim intRowvalid As Integer  '��¼��Ч������
    Dim RecReturn As Recordset
    Dim rsMaterial As New ADODB.Recordset
    Dim intCheckAll As Integer
    Dim blnReturn As Boolean    '�����жϽ�������Ƿ��Ƕ�ѡ����
    Dim intRow As Integer       '��ǰ��
    Dim strҩƷID As String     '����Щ���ظ���ҩƷid
    Dim rsTemp As ADODB.Recordset '��ʱ��¼�����ظ�ֵ������ݼ�
    Dim lngҩƷID As Long
    Dim strTemp As String
    Dim intOldRow As Integer
    
    On Error GoTo errHandle
    If cboEnterStock.ListCount = 0 Then Exit Sub
    intOldRow = mshBill.Row
    intRow = mshBill.Row
    Select Case mshBill.Col
    Case mconIntColҩ��
        mshBill.CmdEnable = False
        
        If grsMaster.State = adStateClosed Then
            Call SetSelectorRS(2, MStrCaption, cboStock.ItemData(cboStock.ListIndex), cboEnterStock.ItemData(cboEnterStock.ListIndex), mlng����ⷿ)
        End If
        If Not mbln���쵥 Then
'            Set RecReturn = FrmҩƷѡ����.ShowME(Me, 2, cboStock.ItemData(cboStock.ListIndex), cboEnterStock.ItemData(cboEnterStock.ListIndex), mlng����ⷿ, _
'                True, True, False, False, True)
            Set RecReturn = frmSelector.ShowME(Me, 0, 2, , , , cboStock.ItemData(cboStock.ListIndex), cboEnterStock.ItemData(cboEnterStock.ListIndex), mlng����ⷿ, , True, True, True, , , mstrPrivs)
        Else    '���쵥
'            Set RecReturn = FrmҩƷѡ����.ShowME(Me, 2, cboStock.ItemData(cboStock.ListIndex), cboEnterStock.ItemData(cboEnterStock.ListIndex), mlng����ⷿ, _
'                mbln��ȷ����, Not mbln��ȷ����, False, False, True)
            Set RecReturn = frmSelector.ShowME(Me, 0, 2, , , , cboStock.ItemData(cboStock.ListIndex), cboEnterStock.ItemData(cboEnterStock.ListIndex), mlng����ⷿ, , True, True, True, , , mstrPrivs)
        End If
        mshBill.CmdEnable = True
        
        If RecReturn.RecordCount > 0 Then
            Set RecReturn = CheckData(RecReturn) '����ظ���¼��ʱ���޿��ļ�¼���������������ļ�¼���˵�
        End If
        
        If RecReturn.RecordCount > 0 Then
            With mshBill
                Dim lngCurRow As Long
                
                RecReturn.MoveFirst
                For i = 1 To RecReturn.RecordCount
                    intRow = .Row
                    If IsSelf_Command(RecReturn!ҩƷid) Then
                        '��ȡ������ҩƷ�����ҩƷ�������β�������
                        Set rsMaterial = GetMaterial(RecReturn!ҩƷid)
                        
                        If rsMaterial.RecordCount > 0 Then
                            Set rsMaterial = CheckData(rsMaterial) '����ظ���¼��ʱ���޿��ļ�¼���������������ļ�¼���˵�
                        End If
                        
                        If rsMaterial.RecordCount <> 0 Then '��������ݣ��������ƶ�����һ����¼
                            rsMaterial.MoveFirst
                        End If
                        lngCurRow = mshBill.Row
                        mshBill.rows = mshBill.rows + rsMaterial.RecordCount
                        mshBill.Row = lngCurRow
                        With rsMaterial
                            Do While Not .EOF
                                mshBill.TextMatrix(mshBill.Row, mconIntCol�к�) = mshBill.Row
                                SetColValue mshBill.Row, !ҩƷid, "[" & !ҩƷ���� & "]", !ͨ����, IIf(IsNull(!��Ʒ��), "", !��Ʒ��), _
                                    Nvl(!ҩƷ��Դ), "" & !����ҩ��, _
                                    IIf(IsNull(!���), "", !���), IIf(IsNull(!����), "", !����), _
                                    Choose(mintUnit, !�ۼ۵�λ, !���ﵥλ, !סԺ��λ, !ҩ�ⵥλ), _
                                    !�ۼ�, IIf(IsNull(!����), "", !����), _
                                    IIf(IsNull(!Ч��), "", Format(!Ч��, "yyyy-MM-dd")), _
                                    IIf(IsNull(!���Ч��), "0", !���Ч��), _
                                    !ҩ�����, _
                                    IIf(IsNull(!��������), "0", !��������), _
                                    IIf(IsNull(!ʵ�ʽ��), "0", !ʵ�ʽ��), _
                                    IIf(IsNull(!ʵ�ʲ��), "0", !ʵ�ʲ��), _
                                    IIf(IsNull(!ָ�������), "0", !ָ�������), _
                                    Choose(mintUnit, 1, !�����װ, !סԺ��װ, !ҩ���װ), _
                                    IIf(IsNull(!����), 0, !����), !ʱ��, !ҩ������, !�ϴι�Ӧ��ID, _
                                    IIf(IsNull(!��׼�ĺ�), "", !��׼�ĺ�)
                                .MoveNext
                                mshBill.Row = mshBill.Row + 1
                            Loop
                        End With
'                        mshBill.Row = lngCurRow
                    Else
                        mshBill.TextMatrix(mshBill.Row, mconIntCol�к�) = .Row
                        SetColValue .Row, RecReturn!ҩƷid, "[" & RecReturn!ҩƷ���� & "]", RecReturn!ͨ����, IIf(IsNull(RecReturn!��Ʒ��), "", RecReturn!��Ʒ��), _
                            Nvl(RecReturn!ҩƷ��Դ), "" & RecReturn!����ҩ��, _
                            IIf(IsNull(RecReturn!���), "", RecReturn!���), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                            Choose(mintUnit, RecReturn!�ۼ۵�λ, RecReturn!���ﵥλ, RecReturn!סԺ��λ, RecReturn!ҩ�ⵥλ), _
                            RecReturn!�ۼ�, IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                            IIf(IsNull(RecReturn!Ч��), "", Format(RecReturn!Ч��, "yyyy-MM-dd")), _
                            IIf(IsNull(RecReturn!���Ч��), "0", RecReturn!���Ч��), _
                            RecReturn!ҩ�����, _
                            IIf(IsNull(RecReturn!��������), "0", RecReturn!��������), _
                            IIf(IsNull(RecReturn!ʵ�ʽ��), "0", RecReturn!ʵ�ʽ��), _
                            IIf(IsNull(RecReturn!ʵ�ʲ��), "0", RecReturn!ʵ�ʲ��), _
                            IIf(IsNull(RecReturn!ָ�������), "0", RecReturn!ָ�������), _
                            Choose(mintUnit, 1, RecReturn!�����װ, RecReturn!סԺ��װ, RecReturn!ҩ���װ), _
                            IIf(IsNull(RecReturn!����), 0, RecReturn!����), RecReturn!ʱ��, RecReturn!ҩ������, _
                            RecReturn!�ϴι�Ӧ��ID, IIf(IsNull(RecReturn!��׼�ĺ�), "", RecReturn!��׼�ĺ�)
                    End If
                    
                    .Col = mconIntCol��д����

                    If (.TextMatrix(intRow, 0) = "" Or intRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, 0) <> "" Then
                        .rows = .rows + 1
                    End If

                    .Row = .rows - 1
                    RecReturn.MoveNext
                Next
                .Row = intOldRow
            End With
            RecReturn.Close
        Else
            mshBill.Row = intOldRow
        End If
    Case mconIntCol����
        gstrSQL = "Select Distinct �ϴ�����,�ϴβ���,��׼�ĺ�,�ϴι�Ӧ��ID From ҩƷ��� Where ����=1 And �ⷿid=[1] And ҩƷid=[2] "
        Set RecReturn = zlDataBase.OpenSQLRecord(gstrSQL, "ȡ������Ϣ", cboEnterStock.ItemData(cboEnterStock.ListIndex), mshBill.TextMatrix(mshBill.Row, 0))
        If RecReturn.RecordCount = 0 Then
            MsgBox "û���ҵ���ҩƷ��������Ϣ�����ֹ��������š�"
            Exit Sub
        End If
        
        Set msh������Ϣ.Recordset = RecReturn
        With msh������Ϣ
            .Redraw = False
            .Left = Pic����.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
            .Top = Pic����.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight
            .Visible = True
            .SetFocus
            .ColWidth(0) = 800
            .ColWidth(1) = 1000
            .ColWidth(2) = 1000
            .ColWidth(3) = 0
            .Row = 1
            .Col = 0
            .TopRow = 1
            .ColSel = .Cols - 1
            .Redraw = True
            Exit Sub
        End With
    Case mconIntCol����
        Dim rsProvider As New Recordset
        
        gstrSQL = "Select ����,����,���� From ҩƷ������ Where վ�� = [1] Or վ�� is Null Order By ����"
        Set rsProvider = zlDataBase.OpenSQLRecord(gstrSQL, "ҩƷ������", gstrNodeNo)
        
        If rsProvider.EOF Then
            rsProvider.Close
            Exit Sub
        End If
        With FrmSelect
            Set .TreeRec = rsProvider
            .StrNode = "����ҩƷ������"
            .lngMode = 1
            .Show 1, Me
            If .BlnSuccess = True Then
                mshBill.TextMatrix(mshBill.Row, mconIntCol����) = .CurrentName
                mshBill.Col = mconIntCol����
            End If
        End With
        Unload FrmSelect
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mshbill_EditChange(curText As String)
    mblnChange = True
End Sub

Private Sub mshBill_EditKeyPress(KeyAscii As Integer)
    Dim strkey As String
    Dim intDigit As Integer
    
    With mshBill
        strkey = .Text
'        If strkey = "" Then
'            strkey = .TextMatrix(.Row, .Col)
'        End If
        Select Case .Col
            Case mconIntCol��д����, mconIntColʵ������
                intDigit = mintNumberDigit
        End Select
        
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then
            If .SelLength = Len(strkey) Then Exit Sub
            If Len(Mid(strkey, InStr(1, strkey, ".") + 1)) >= intDigit And strkey Like "*.*" Then
                KeyAscii = 0
                Exit Sub
            Else
                Exit Sub
            End If
        End If
    End With
End Sub

Private Sub mshbill_EnterCell(Row As Long, Col As Long)
    If Not mblnEnterCell Then Exit Sub
    
    With mshBill
        If Row > 0 Then
            .SetRowColor CLng(Row), &HFFCECE, True
        End If
        
        If .Row <> .LastRow Then
            SetInputFormat .Row
        End If
        
        Select Case .Col
            Case mconIntColҩ��
                .txtCheck = False
                .MaxLength = 40
                'ֻ��ҩ���в���ʾ�ϼ���Ϣ�Ϳ����
                Call ��ʾ�ϼƽ��
                
            Case mconIntCol����
                .txtCheck = False
'                .TextMask = "1234567890"
                .MaxLength = mintBatchNoLen
                If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Or mint�༭״̬ = 3 Or mint�༭״̬ = 10 Then
                    .ColData(mconIntCol����) = 4
                End If
            Case mconIntColЧ��
                .txtCheck = True
                .TextMask = "1234567890-"
                .MaxLength = 10
                If .TextMatrix(.Row, mconIntColЧ��) = "" And .TextMatrix(.Row, mconIntCol����) <> "" And .ColData(.Col) = 2 Then
                    Dim strxq As String
                    
                    If IsNumeric(.TextMatrix(.Row, mconIntCol����)) And .TextMatrix(.Row, mconIntCol���Ч��) <> "" Then
                        If Split(.TextMatrix(.Row, mconIntCol���Ч��), "||")(0) <> 0 Then
                            strxq = UCase(.TextMatrix(.Row, mconIntCol����))
                            If Not (InStr(1, strxq, "D") <> 0 Or InStr(1, strxq, "E") <> 0) Then
                                strxq = TranNumToDate(strxq)
                                If strxq = "" Then Exit Sub
                                
                                .TextMatrix(.Row, mconIntColЧ��) = Format(DateAdd("M", Split(.TextMatrix(.Row, mconIntCol���Ч��), "||")(0), strxq), "yyyy-mm-dd")
                                If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 Then
                                    '����Ϊ��Ч��
                                    .TextMatrix(.Row, mconIntColЧ��) = Format(DateAdd("D", -1, .TextMatrix(.Row, mconIntColЧ��)), "yyyy-mm-dd")
                                End If
                            End If
                        End If
                    End If
                End If
            Case mconIntCol��д����, mconIntColʵ������
                .txtCheck = True
                .MaxLength = 16
                .TextMask = "-.1234567890"
            Case mconIntCol����
                If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Or mint�༭״̬ = 3 Or mint�༭״̬ = 10 Then
                    .ColData(mconIntCol����) = 1
                End If
                
                OpenIme GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDbUser, "���뷨", "")
                .txtCheck = False
                .MaxLength = 30
                .TxtSetFocus
            Case mconIntCol��׼�ĺ�
                .txtCheck = False
                .MaxLength = 40
        End Select
        
    End With
End Sub

Private Sub mshBill_GotFocus()
    If mintParallelRecord <> 1 Then Exit Sub
    If cboEnterStock.ListCount = 0 Then Exit Sub
    If cboStock.ItemData(cboStock.ListIndex) = cboEnterStock.ItemData(cboEnterStock.ListIndex) Then
        MsgBox "�Բ�������ⷿ���Ƴ��ⷿ��ͬ�ˣ����������ѡ��", vbOKOnly + vbExclamation, gstrSysName
        cboEnterStock.SetFocus
        Exit Sub
    End If
    
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strkey As String
    Dim rsDrug As New Recordset
    Dim strUnitQuantity As String
    Dim intRow As Integer
    Dim strҩƷID As String
    Dim i As Integer
    Dim intOldRow As Integer
    
    On Error GoTo errHandle
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cboEnterStock.ListCount = 0 Then Exit Sub
    
    With mshBill
        .Text = UCase(Trim(.Text))
        strkey = UCase(Trim(.Text))
        
        intOldRow = .Row
        intRow = .Row
        If Mid(strkey, 1, 1) = "[" Then
            If InStr(2, strkey, "]") <> 0 Then
                strkey = Mid(strkey, 2, InStr(2, strkey, "]") - 2)
            Else
                strkey = Mid(strkey, 2)
            End If
        End If
        Select Case .Col
            
            Case mconIntColҩ��
                If strkey <> "" Then
                    Dim lngCurRow As Long
                    Dim rsMaterial As New ADODB.Recordset

                    Dim RecReturn As Recordset
                    Dim sngLeft As Single
                    Dim sngTop As Single
                    
                    sngLeft = Me.Left + Pic����.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                    sngTop = Me.Top + Me.Height - Me.ScaleHeight + Pic����.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                    If sngTop + 3630 > Screen.Height Then
                        sngTop = sngTop - mshBill.MsfObj.CellHeight - 4530
                    End If
                    
                    If grsMaster.State = adStateClosed Then
                        Call SetSelectorRS(2, MStrCaption, cboStock.ItemData(cboStock.ListIndex), cboEnterStock.ItemData(cboEnterStock.ListIndex), mlng����ⷿ)
                    End If
'                    If Not mbln���쵥 Then
'                        Set RecReturn = FrmҩƷ��ѡѡ����.ShowME(Me, 2, cboStock.ItemData(cboStock.ListIndex), cboEnterStock.ItemData(cboEnterStock.ListIndex), mlng����ⷿ, _
'                            strkey, sngLeft, sngTop, True, True, False, False, True)
'                    Else
'                        Set RecReturn = FrmҩƷ��ѡѡ����.ShowME(Me, 2, cboStock.ItemData(cboStock.ListIndex), cboEnterStock.ItemData(cboEnterStock.ListIndex), mlng����ⷿ, _
'                            strkey, sngLeft, sngTop, mbln��ȷ����, Not mbln��ȷ����, False, False, True)
'                    End If
                    
                    If mbln���쵥 Then
                        Set RecReturn = frmSelector.ShowME(Me, 1, 2, strkey, sngLeft, sngTop, cboStock.ItemData(cboStock.ListIndex), cboEnterStock.ItemData(cboEnterStock.ListIndex), mlng����ⷿ, , True, True, True, , , mstrPrivs)
                    Else
                        Set RecReturn = frmSelector.ShowME(Me, 1, 2, strkey, sngLeft, sngTop, cboStock.ItemData(cboStock.ListIndex), cboEnterStock.ItemData(cboEnterStock.ListIndex), mlng����ⷿ, , True, True, True, , , mstrPrivs)
                    End If
                    
                    If RecReturn.RecordCount > 0 Then
                        Set RecReturn = CheckData(RecReturn) '����ظ���¼��ʱ���޿��ļ�¼���������������ļ�¼���˵�
                    End If
                    
                    If RecReturn.RecordCount > 0 Then
                        RecReturn.MoveFirst
                        
                        For i = 1 To RecReturn.RecordCount
                            intRow = .Row
                            If IsSelf_Command(RecReturn!ҩƷid) Then
                                '��ȡ������ҩƷ�����ҩƷ�������β�������
                                Set rsMaterial = GetMaterial(RecReturn!ҩƷid)
                                
                                If rsMaterial.RecordCount > 0 Then
                                    Set rsMaterial = CheckData(rsMaterial) '����ظ���¼��ʱ���޿��ļ�¼���������������ļ�¼���˵�
                                End If
                                
                                If rsMaterial.RecordCount <> 0 Then '��������ݣ��������ƶ�����һ����¼
                                    rsMaterial.MoveFirst
                                End If
                                
                                lngCurRow = mshBill.Row
                                mshBill.rows = mshBill.rows + rsMaterial.RecordCount
                                mshBill.Row = lngCurRow
                                With rsMaterial
                                    Do While Not .EOF
                                        mshBill.TextMatrix(mshBill.Row, mconIntCol�к�) = mshBill.Row
                                        SetColValue mshBill.Row, !ҩƷid, "[" & !ҩƷ���� & "]", !ͨ����, IIf(IsNull(!��Ʒ��), "", !��Ʒ��), _
                                            Nvl(!ҩƷ��Դ), "" & !����ҩ��, _
                                            IIf(IsNull(!���), "", !���), IIf(IsNull(!����), "", !����), _
                                            Choose(mintUnit, !�ۼ۵�λ, !���ﵥλ, !סԺ��λ, !ҩ�ⵥλ), _
                                            !�ۼ�, IIf(IsNull(!����), "", !����), _
                                            IIf(IsNull(!Ч��), "", Format(!Ч��, "yyyy-MM-dd")), _
                                            IIf(IsNull(!���Ч��), "0", !���Ч��), _
                                            !ҩ�����, _
                                            IIf(IsNull(!��������), "0", !��������), _
                                            IIf(IsNull(!ʵ�ʽ��), "0", !ʵ�ʽ��), _
                                            IIf(IsNull(!ʵ�ʲ��), "0", !ʵ�ʲ��), _
                                            IIf(IsNull(!ָ�������), "0", !ָ�������), _
                                            Choose(mintUnit, 1, !�����װ, !סԺ��װ, !ҩ���װ), _
                                            IIf(IsNull(!����), 0, !����), !ʱ��, !ҩ������, !�ϴι�Ӧ��ID, _
                                            IIf(IsNull(!��׼�ĺ�), "", !��׼�ĺ�)
                                        .MoveNext
                                        mshBill.Row = mshBill.Row + 1
                                    Loop
                                End With
'                                mshBill.Row = lngCurRow
                            Else
                                mshBill.TextMatrix(mshBill.Row, mconIntCol�к�) = .Row
                                If SetColValue(.Row, RecReturn!ҩƷid, "[" & RecReturn!ҩƷ���� & "]", RecReturn!ͨ����, IIf(IsNull(RecReturn!��Ʒ��), "", RecReturn!��Ʒ��), _
                                        Nvl(RecReturn!ҩƷ��Դ), "" & RecReturn!����ҩ��, _
                                        IIf(IsNull(RecReturn!���), "", RecReturn!���), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                                        Choose(mintUnit, RecReturn!�ۼ۵�λ, RecReturn!���ﵥλ, RecReturn!סԺ��λ, RecReturn!ҩ�ⵥλ), _
                                        IIf(IsNull(RecReturn!�ۼ�), 0, RecReturn!�ۼ�), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                                        IIf(IsNull(RecReturn!Ч��), "", Format(RecReturn!Ч��, "yyyy-MM-dd")), _
                                        IIf(IsNull(RecReturn!���Ч��), "0", RecReturn!���Ч��), _
                                        RecReturn!ҩ�����, _
                                        IIf(IsNull(RecReturn!��������), "0", RecReturn!��������), _
                                        IIf(IsNull(RecReturn!ʵ�ʽ��), "0", RecReturn!ʵ�ʽ��), _
                                        IIf(IsNull(RecReturn!ʵ�ʲ��), "0", RecReturn!ʵ�ʲ��), _
                                        IIf(IsNull(RecReturn!ָ�������), "0", RecReturn!ָ�������), _
                                        Choose(mintUnit, 1, RecReturn!�����װ, RecReturn!סԺ��װ, RecReturn!ҩ���װ), _
                                        IIf(IsNull(RecReturn!����), 0, RecReturn!����), RecReturn!ʱ��, RecReturn!ҩ������, _
                                        RecReturn!�ϴι�Ӧ��ID, IIf(IsNull(RecReturn!��׼�ĺ�), "", RecReturn!��׼�ĺ�)) = False Then
                                    Cancel = True
                                    Exit Sub
                                End If
                                
                                .Text = .TextMatrix(.Row, .Col)
                            End If
                            
                            If (.TextMatrix(intRow, 0) = "" Or intRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, 0) <> "" Then
                                .rows = .rows + 1
                            End If
                            .Row = .rows - 1
                            RecReturn.MoveNext
                        Next
                        .Row = intOldRow
                    Else
                        .Text = .TextMatrix(.Row, .Col)
                        Cancel = True
                    End If
                End If
            Case mconIntCol����
                '�޴���
                If strkey = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mconIntCol����) = ""
                    End If
                    If .ColData(mconIntColЧ��) = 2 Then
                        .Col = mconIntColЧ��
                    Else
                        .Col = mconIntCol��д����
                    End If
                    Cancel = True
                    Exit Sub
                Else
                    gstrSQL = "Select Distinct �ϴ�����,�ϴβ���,��׼�ĺ�,�ϴι�Ӧ��ID From ҩƷ��� Where ����=1 And �ⷿid=[1] And ҩƷid=[2] And �ϴ����� like [3] "
                    Set RecReturn = zlDataBase.OpenSQLRecord(gstrSQL, "ȡ������Ϣ", cboEnterStock.ItemData(cboEnterStock.ListIndex), mshBill.TextMatrix(mshBill.Row, 0), IIf(gstrMatchMethod = "0", "%", "") & strkey & "%")
                    If RecReturn.RecordCount = 0 Then
                        If .ColData(mconIntColЧ��) = 2 Then
                            .Col = mconIntColЧ��
                        Else
                            .Col = mconIntCol��д����
                        End If
                        .TextMatrix(.Row, mconIntCol����) = strkey
                        Cancel = True
                        Exit Sub
                    ElseIf RecReturn.RecordCount = 1 Then
                        .TextMatrix(.Row, mconIntCol����) = Nvl(RecReturn.Fields("�ϴ�����"), "")
                        .Text = Nvl(RecReturn.Fields("�ϴ�����"), "")
                        .TextMatrix(.Row, mconIntCol����) = Nvl(RecReturn.Fields("�ϴβ���"), "")
                        .TextMatrix(.Row, mconIntCol��׼�ĺ�) = Nvl(RecReturn.Fields("��׼�ĺ�"), "")
                        .TextMatrix(.Row, mconIntCol�ϴι�Ӧ��ID) = Nvl(RecReturn.Fields("�ϴι�Ӧ��ID"), 0)
                    Else
                        Set msh������Ϣ.Recordset = RecReturn
                        With msh������Ϣ
                            .Redraw = False
                            .Left = Pic����.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                            .Top = Pic����.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight
                            .Visible = True
                            .SetFocus
                            .ColWidth(0) = 800
                            .ColWidth(1) = 1000
                            .ColWidth(2) = 1000
                            .ColWidth(3) = 0
                            .Row = 1
                            .Col = 0
                            .TopRow = 1
                            .ColSel = .Cols - 1
                            .Redraw = True
                            Cancel = True
                            Exit Sub
                        End With
                    End If
                End If
            Case mconIntColЧ��
                '�д���
                If strkey <> "" Then
                    If Len(strkey) = 8 And InStr(1, strkey, "-") = 0 Then
                        strkey = TranNumToDate(strkey)
                        If strkey = "" Then
                            MsgBox "�Բ���ʧЧ�ڱ���Ϊ�����ͣ�", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            Exit Sub
                        End If
                        .Text = strkey
                        Exit Sub
                    End If
                    If Not IsDate(strkey) Then
                        MsgBox "�Բ���ʧЧ�ڱ���Ϊ��������(2000-10-10) ��20001010��,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        Exit Sub
                    End If
                ElseIf strkey = "" And strkey <> .TextMatrix(.Row, mconIntColЧ��) Then
                
                    If .TxtVisible = True Then
                        .Text = " "
                        Exit Sub
                    End If
                    
                    Exit Sub
                End If
            
            Case mconIntCol��д����, mconIntColʵ������
                If .TextMatrix(.Row, 0) = "" Then
                    .Text = ""
                    Exit Sub
                End If
                
                If .TextMatrix(.Row, .Col) = "" And strkey = "" Then
                    MsgBox "�Բ��������������룡", vbOKOnly + vbInformation, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
'                If .TextMatrix(.Row, .Col) <> "" And strkey = "" Then
'                    MsgBox "�Բ��������������룡", vbOKOnly + vbInformation, gstrSysName
'                    Cancel = True
'                    .TxtSetFocus
'                    Exit Sub
'                End If
                If Not IsNumeric(strkey) And strkey <> "" Then
                    MsgBox "�Բ�����������Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strkey <> "" Then
                    If Val(strkey) <= 0 And Not (mint�༭״̬ = 3 Or mint�༭״̬ = 6 Or mint�༭״̬ = 10) Then
                        MsgBox "�Բ����������������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If mint�༭״̬ = 6 Then
                        If Not ��ͬ����(Val(strkey), Val(.TextMatrix(.Row, mconIntCol��д����))) Then
                            MsgBox "�Բ��𣬳��������ķ���Ӧ����ԭ������һ�£�", vbInformation, gstrSysName
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                        
                        If Val(strkey) >= 0 Then
                            If Val(strkey) > Val(.TextMatrix(.Row, mconIntCol��д����)) Then
                                MsgBox "�Բ��𣬳����������ܴ���ԭ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                                Cancel = True
                                .TxtSetFocus
                                Exit Sub
                            End If
                        Else
                            If Val(strkey) < Val(.TextMatrix(.Row, mconIntCol��д����)) Then
                                MsgBox "�Բ��𣬳����������ܴ���ԭ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                                Cancel = True
                                .TxtSetFocus
                                Exit Sub
                            End If
                        End If
                    End If
                    
                    If gtype_UserSysParms.P174_ҩƷ�ƿ���ȷ���� = 1 Then
                        If Not CompareUsableQuantity(.Row, strkey) Then
                            Cancel = True
                            .TxtSetFocus
                            Exit Sub
                        End If
                    End If
                    
                    strkey = GetFormat(strkey, mintNumberDigit)
                    .Text = strkey
                    .TextMatrix(.Row, .Col) = .Text
                    
                    If .TextMatrix(.Row, mconIntCol�ۼ�) <> "" Then
                        .TextMatrix(.Row, mconIntCol�ۼ۽��) = GetFormat(.TextMatrix(.Row, mconIntCol�ۼ�) * strkey, mintMoneyDigit)
                    End If
                    
'                    .TextMatrix(.Row, mconintCol���) = FormatEx(Get������(Val(cboStock.ItemData(cboStock.ListIndex)), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol����)), Val(.TextMatrix(.Row, mconIntColʵ�ʽ��)), Val(.TextMatrix(.Row, mconIntColʵ�ʲ��)), Val(.TextMatrix(.Row, mconIntCol�ۼ۽��)), Val(strkey) * Val(mshBill.TextMatrix(.Row, mconIntCol����ϵ��))), mintMoneyDigit)
                        
                    If strkey <> 0 And (mint�༭״̬ = 1 Or mint�༭״̬ = 2 Or mint�༭״̬ = 3) Then
                        .TextMatrix(.Row, mconIntCol�ɹ���) = GetFormat(Get�ɱ���(Val(.TextMatrix(.Row, 0)), cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, mconIntCol����))) * Val(mshBill.TextMatrix(.Row, mconIntCol����ϵ��)), mintCostDigit)
                    End If
                    .TextMatrix(.Row, mconIntCol�ɹ����) = GetFormat(Val(.TextMatrix(.Row, mconIntCol�ɹ���)) * strkey, mintMoneyDigit)
                    .TextMatrix(.Row, mconintCol���) = GetFormat(Val(.TextMatrix(.Row, mconIntCol�ۼ۽��)) - .TextMatrix(.Row, mconIntCol�ɹ����), mintMoneyDigit)
                    
                    If .Col = mconIntCol��д���� Then
                        .TextMatrix(.Row, mconIntColʵ������) = strkey
                    End If
                End If
                ��ʾ�ϼƽ��
                If mbln���쵥 Then Call ShowColor(mshBill.Row)
            Case mconIntCol����
                '�޴���
                If .Text = "" Then
                    If .TxtVisible = True Then
                        .TextMatrix(.Row, mconIntCol����) = ""
                    End If
                    If .ColData(mconIntCol����) = 5 Then
                        .Col = mconIntCol��д����
                    Else
                        .Col = mconIntCol����
                    End If
                    Cancel = True
                            
                    Exit Sub
                Else
                    Dim rs���� As New Recordset
                    
                    gstrSQL = "Select ����,����,���� From ҩƷ������ " _
                            & "Where (վ�� = [4] Or վ�� is Null) And (upper(����) like [1] or Upper(����) like [2] or Upper(����) like [3]) Order By ����"
                    Set rs���� = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, _
                        IIf(gstrMatchMethod = "0", "%", "") & strkey & "%", _
                        IIf(gstrMatchMethod = "0", "%", "") & strkey & "%", _
                        IIf(gstrMatchMethod = "0", "%", "") & strkey & "%", _
                        gstrNodeNo)
                    
                    If rs����.RecordCount = 0 Then
                        Exit Sub
                    ElseIf rs����.RecordCount = 1 Then
                        .TextMatrix(.Row, mconIntCol����) = rs����.Fields("����")
                        .Text = rs����.Fields("����")
                    Else
                        Set msh����.Recordset = rs����
                        With msh����
                            .Redraw = False
                            .Left = Pic����.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                            .Top = Pic����.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight
                            .Visible = True
                            .SetFocus
                            .ColWidth(0) = 800
                            .ColWidth(1) = 800
                            .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                            .Row = 1
                            .Col = 0
                            .TopRow = 1
                            .ColSel = .Cols - 1
                            .Redraw = True
                            Cancel = True
                            Exit Sub
                        End With
                    End If
                End If
                OpenIme
        End Select
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'��ҩƷĿ¼��ȡֵ��������Ӧ����
Private Function SetColValue(ByVal intRow As Integer, ByVal lngҩƷID As Long, _
    ByVal strҩƷ���� As String, ByVal strͨ���� As String, ByVal str��Ʒ�� As String, ByVal strҩƷ��Դ As String, ByVal str����ҩ�� As String, _
    ByVal str��� As String, ByVal str���� As String, ByVal str��λ As String, ByVal num�ۼ� As Double, ByVal str���� As String, _
    ByVal strЧ�� As String, ByVal int���Ч�� As Integer, ByVal int�������� As Integer, _
    ByVal num�������� As Double, ByVal numʵ�ʽ�� As Double, ByVal numʵ�ʲ�� As Double, _
    ByVal numָ������� As Double, ByVal num����ϵ�� As Double, ByVal lng���� As Long, _
    ByVal int�Ƿ��� As Integer, ByVal intҩ������ As Integer, ByVal lng�ϴι�Ӧ��ID As Long, ByVal str��׼�ĺ� As String) As Boolean
    
    Dim intCount As Integer
    Dim intCol As Integer
    Dim dbltotal As Double
    Dim dblPrice As Double
    Dim intLop As Integer
    Dim dblCost As Double
    Dim strҩ�� As String
    
    Dim rsTemp As New ADODB.Recordset
    Dim rsRecord As ADODB.Recordset
    Dim dbl���ۼ� As Double, dblָ�����ۼ� As Double, dbl��������� As Double, dbl�ӳ��� As Double, dbl�ɱ��� As Double
    
    SetColValue = False
    
    On Error GoTo errHandle
    
    With mshBill
        For intCol = 0 To .Cols - 1
            If intCol <> mconIntCol�к� Then .TextMatrix(intRow, intCol) = ""
        Next
        
        .TextMatrix(intRow, 0) = lngҩƷID
        
        If gintҩƷ������ʾ = 0 Or gintҩƷ������ʾ = 2 Then
            strҩ�� = strͨ����
        Else
            strҩ�� = IIf(str��Ʒ�� <> "", str��Ʒ��, strͨ����)
        End If
        
        .TextMatrix(intRow, mconIntColҩƷ���������) = strҩƷ���� & strҩ��
        .TextMatrix(intRow, mconIntColҩƷ����) = strҩƷ����
        .TextMatrix(intRow, mconIntColҩƷ����) = strҩ��
        
        If mintDrugNameShow = 1 Then
            .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ����)
        ElseIf mintDrugNameShow = 2 Then
            .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ����)
        Else
            .TextMatrix(intRow, mconIntColҩ��) = .TextMatrix(intRow, mconIntColҩƷ���������)
        End If
        
        .TextMatrix(intRow, mconIntCol��Ʒ��) = str��Ʒ��
        
        .TextMatrix(intRow, mconIntCol��Դ) = strҩƷ��Դ
        .TextMatrix(intRow, mconIntCol����ҩ��) = str����ҩ��
        .TextMatrix(intRow, mconIntCol���) = str���
        .TextMatrix(intRow, mconIntCol����) = str����
        .TextMatrix(intRow, mconIntCol��λ) = str��λ
        
        '���ش��װ�����Ϣ
        gstrSQL = "select a.�ͻ���λ,a.�ͻ���װ from ҩƷ��� a where a.ҩƷid=[1]"
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "�����Ϣ", lngҩƷID)
        If num����ϵ�� = 0 Or Nvl(rsTemp!�ͻ���װ) = "" Or Nvl(rsTemp!�ͻ���λ) = "" Then
            .TextMatrix(intRow, mconIntCol�ͻ���λ) = ""
        Else
            .TextMatrix(intRow, mconIntCol�ͻ���λ) = rsTemp!�ͻ���λ & "(1" & rsTemp!�ͻ���λ & "=" & zlStr.FormatEx(rsTemp!�ͻ���װ / num����ϵ��, 1, , True) & str��λ & ")"
        End If
        
        If gtype_UserSysParms.P174_ҩƷ�ƿ���ȷ���� = 0 Then
            .TextMatrix(intRow, mconIntCol����) = 0
        Else
            .TextMatrix(intRow, mconIntCol����) = lng����
        End If
        .TextMatrix(intRow, mconIntCol����) = str����
        .TextMatrix(intRow, mconIntColЧ��) = Format(strЧ��, "yyyy-mm-dd")
        .TextMatrix(intRow, mconIntCol��׼�ĺ�) = str��׼�ĺ�
        .TextMatrix(intRow, mconIntCol�ۼ�) = GetFormat(num�ۼ� * num����ϵ��, mintPriceDigit)
        .TextMatrix(intRow, mconIntCol��������) = int��������
        .TextMatrix(intRow, mconIntCol��������) = GetFormat(num��������, mintNumberDigit)
        .TextMatrix(intRow, mconIntCol���Ч��) = int���Ч�� & "||" & int�Ƿ��� & "||" & intҩ������
        .TextMatrix(intRow, mconIntColʵ�ʲ��) = numʵ�ʲ��
        .TextMatrix(intRow, mconIntColʵ�ʽ��) = numʵ�ʽ��
        .TextMatrix(intRow, mconIntColָ�������) = numָ�������
        .TextMatrix(intRow, mconIntCol����ϵ��) = num����ϵ��
        
        .TextMatrix(intRow, mconIntCol�ϴι�Ӧ��ID) = lng�ϴι�Ӧ��ID
        .TextMatrix(intRow, mconIntCol��������) = IIf(GetҩƷ��������(lngҩƷID, cboEnterStock.ItemData(cboEnterStock.ListIndex)) = True, "1", 0)
                
        If int�Ƿ��� = 1 Then
            .TextMatrix(intRow, mconIntCol�ۼ�) = GetFormat(Get�ۼ�(True, lngҩƷID, cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(intRow, mconIntCol����))) * num����ϵ��, mintPriceDigit)
'            If mbln��ȷ���� = True Then
'                gstrSQL = "Select Decode(Nvl(����, 0), 0, ʵ�ʽ�� / ʵ������, Nvl(���ۼ�, ʵ�ʽ�� / ʵ������))*" & num����ϵ�� & " as  �ۼ� " _
'                    & "  from ҩƷ��� " _
'                    & " where �ⷿid=[1] " _
'                    & " and ҩƷid=[2] " _
'                    & " and ����=1 and ʵ������>0 and " _
'                    & " nvl(����,0)=[3] "
'                Set rsRecord = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, cboStock.ItemData(cboStock.ListIndex), lngҩƷID, lng����)
'
'                If Not rsRecord.EOF Then
'                    dblPrice = rsRecord.Fields(0).Value
'                Else
'                    dblPrice = 0
'                End If
'                .TextMatrix(intRow, mconIntCol�ۼ�) = GetFormat(dblPrice, mintPriceDigit)
'            End If
        End If
        
        If IsLowerLimit(mlng����ⷿ, lngҩƷID) Then Call SetForeColor_ROW(mlng��ɫ)
        Call CheckLapse(strЧ��)
        SetInputFormat intRow
        
        Call ��ʾ�����(intRow)
    End With
    SetColValue = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Function

Private Sub SetInputFormat(ByVal intRow As Integer)
    Dim bln�ⷿ As Boolean
    Dim blnҩ����� As Boolean, blnҩ������ As Boolean
    Dim rsTemp As ADODB.Recordset
    
    '˵����ֻ�������ⷿ�����ж�
    '   1�����ⷿ��ҩ�����������������������Ϣ
    '   2�����ҩ����ҩ������������������������Ϣ
    
    On Error GoTo errHandle
    If mblnEdit = False Then Exit Sub
    If cboEnterStock.ListCount = 0 Then Exit Sub
    If Val(mshBill.TextMatrix(mshBill.Row, 0)) = 0 Then Exit Sub
    blnҩ����� = (mshBill.TextMatrix(mshBill.Row, mconIntCol��������) = 1)
    blnҩ������ = (Split(mshBill.TextMatrix(mshBill.Row, mconIntCol���Ч��), "||")(2) = 1)
    bln�ⷿ = CheckStockProperty(cboEnterStock.ItemData(cboEnterStock.ListIndex))
    
    With mshBill
        If .TextMatrix(intRow, mconIntCol����) = "" And _
            ((bln�ⷿ And blnҩ�����) Or (Not bln�ⷿ And blnҩ������)) Then
            .ColData(mconIntCol����) = 4              '���ı�����
            .ColData(mconIntCol����) = 1               '��ť
            If .TextMatrix(intRow, mconIntCol���Ч��) <> "" Then
                If Split(.TextMatrix(intRow, mconIntCol���Ч��), "||")(0) <> 0 Then
                    .ColData(mconIntColЧ��) = 2          '���������
                Else
                    .ColData(mconIntColЧ��) = 5
                End If
            Else
                .ColData(mconIntColЧ��) = 5
            End If
        ElseIf bln�ⷿ And blnҩ����� And Not blnҩ������ Then 'ҩ����ҩ���ƿ⣬ҩ����������ҩ�����
            gstrSQL = "Select �ⷿid From ҩƷ��� Where �ⷿid=[1] And ҩƷid=[2] And Rownum=1 "
            Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "�жϿⷿ���޿��", cboEnterStock.ItemData(cboEnterStock.ListIndex), mshBill.TextMatrix(mshBill.Row, 0))
            
            If rsTemp.RecordCount = 0 Then
                .ColData(mconIntCol����) = 4
                .ColData(mconIntCol����) = 1
                .ColData(mconIntCol��׼�ĺ�) = 4
            Else
                .ColData(mconIntCol����) = 1
                If .TextMatrix(intRow, mconIntCol���Ч��) <> "" Then
                    If Split(.TextMatrix(intRow, mconIntCol���Ч��), "||")(0) <> 0 Then
                        .ColData(mconIntColЧ��) = 2          '���������
                    Else
                        .ColData(mconIntColЧ��) = 5
                    End If
                Else
                    .ColData(mconIntColЧ��) = 5
                End If
            End If
        Else
            .ColData(mconIntCol����) = 5              '��ֹ
            .ColData(mconIntColЧ��) = 5
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mshBill_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        With mshBill
           If .Col = mconIntColҩ�� Then
                PopupMenu mnuCol, 2
            End If
        End With
    End If
End Sub

Private Sub msh����_DblClick()
    msh����_KeyDown vbKeyReturn, 0
End Sub


Private Sub msh����_KeyDown(KeyCode As Integer, Shift As Integer)
    With mshBill
    
        If KeyCode = vbKeyEscape Then
            msh����.Visible = False
            .SetFocus
        End If
        
        If KeyCode = vbKeyReturn Then
            .TextMatrix(.Row, .Col) = msh����.TextMatrix(msh����.Row, 2)
            msh����.Visible = False
            .Col = mconIntCol����
            .SetFocus
        End If
    
    End With
End Sub


Private Sub msh����_LostFocus()
    If msh����.Visible Then
        msh����.Visible = False
    End If
End Sub


Private Sub msh������Ϣ_DblClick()
    msh������Ϣ_KeyDown vbKeyReturn, 0
End Sub


Private Sub msh������Ϣ_KeyDown(KeyCode As Integer, Shift As Integer)
    With mshBill
        If KeyCode = vbKeyEscape Then
            msh������Ϣ.Visible = False
            .SetFocus
        End If
        
        If KeyCode = vbKeyReturn Then
            .TextMatrix(.Row, mconIntCol����) = msh������Ϣ.TextMatrix(msh������Ϣ.Row, 0)
            .TextMatrix(.Row, mconIntCol����) = msh������Ϣ.TextMatrix(msh������Ϣ.Row, 1)
            .TextMatrix(.Row, mconIntCol��׼�ĺ�) = msh������Ϣ.TextMatrix(msh������Ϣ.Row, 2)
            .TextMatrix(.Row, mconIntCol�ϴι�Ӧ��ID) = Val(msh������Ϣ.TextMatrix(msh������Ϣ.Row, 3))
            msh������Ϣ.Visible = False
            .Col = mconIntCol��д����
            .SetFocus
        End If
    
    End With
End Sub


Private Sub msh������Ϣ_LostFocus()
    If msh������Ϣ.Visible Then
        msh������Ϣ.Visible = False
    End If
End Sub


Private Sub staThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If Panel.Key = "PY" And staThis.Tag <> "PY" Then
        Logogram staThis, 0
        staThis.Tag = Panel.Key
    ElseIf Panel.Key = "WB" And staThis.Tag <> "WB" Then
        Logogram staThis, 1
        staThis.Tag = Panel.Key
    End If
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 97 And KeyAscii <= 122 Then
        KeyAscii = KeyAscii - 32
    End If
    
    If KeyAscii = 13 Then
        cmdFind_Click
    End If
End Sub

Private Function ValidData() As Boolean
    Dim bln���ⷿ As Boolean, bln����ⷿ As Boolean
    Dim blnҩ����� As Boolean, blnҩ������ As Boolean
    ValidData = False
    bln���ⷿ = CheckStockProperty(cboEnterStock.ItemData(cboEnterStock.ListIndex))
    bln����ⷿ = CheckStockProperty(cboStock.ItemData(cboStock.ListIndex))
    Dim intLop As Integer
    
    With mshBill
        If .TextMatrix(1, 0) <> "" Then         '�����з�����
            If cboEnterStock.ListCount = 0 Then
                MsgBox "��������������Ĳ��ţ�[������������]�е�ҩƷ����", vbInformation, gstrSysName
                Exit Function
            End If
            If cboStock.ItemData(cboStock.ListIndex) = cboEnterStock.ItemData(cboEnterStock.ListIndex) Then
                MsgBox "�Բ�������ⷿ���Ƴ��ⷿ��ͬ�ˣ�������ѡ��", vbInformation, gstrSysName
                cboEnterStock.SetFocus
                Exit Function
            End If
            
            
            If LenB(StrConv(txtժҪ.Text, vbFromUnicode)) > txtժҪ.MaxLength Then
                MsgBox "ժҪ����,���������" & CInt(txtժҪ.MaxLength / 2) & "�����ֻ�" & txtժҪ.MaxLength & "���ַ�!", vbInformation + vbOKOnly, gstrSysName
                txtժҪ.SetFocus
                Exit Function
            End If
        
            For intLop = 1 To .rows - 1
                If Trim(.TextMatrix(intLop, mconIntColҩ��)) <> "" Then
                    If Trim(Trim(.TextMatrix(intLop, mconIntCol��д����))) = "" Then
                        MsgBox "��" & intLop & "��ҩƷ������Ϊ���ˣ����飡", vbInformation, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol��д����
                        Exit Function
                    End If
                    
                    If LenB(StrConv(Trim(Trim(.TextMatrix(intLop, mconIntCol����))), vbFromUnicode)) > mintBatchNoLen Then
                        MsgBox "��" & intLop & "��ҩƷ�����ų���,���������" & Int(mintBatchNoLen / 2) & "�����ֻ�" & mintBatchNoLen & "���ַ�!", vbInformation + vbOKOnly, gstrSysName
                        .SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol����
                        Exit Function
                    End If
                    
                    
'                    '˵����ֻ�������ⷿ�����ж�
'                    '   1�����ⷿ��ҩ�����������������������Ϣ
'                    '   2�����ҩ����ҩ������������������������Ϣ
'                    blnҩ����� = (mshBill.TextMatrix(intLop, mconIntCol��������) = 1)
'                    blnҩ������ = (Split(mshBill.TextMatrix(intLop, mconIntCol���Ч��), "||")(2) = 1)
'                    If ((bln���ⷿ And blnҩ�����) Or (Not bln���ⷿ And blnҩ������)) And Val(mshBill.TextMatrix(intLop, mconIntColʵ������)) <> 0 Then
'                        If Split(.TextMatrix(intLop, mconIntCol���Ч��), "||")(0) <> 0 Then
'                            If .TextMatrix(intLop, mconIntCol����) = "" Or .TextMatrix(intLop, mconIntColЧ��) = "" Then
'                                MsgBox "��" & intLop & "�е�ҩƷ��Ч��ҩƷ,����������ż�ʧЧ���������뵥���У�", vbInformation, gstrSysName
'                                mshBill.SetFocus
'                                .Row = intLop
'                                .MsfObj.TopRow = intLop
'                                If .TextMatrix(intLop, mconIntCol����) = "" Then
'                                    .Col = mconIntCol����
'                                Else
'                                    .Col = mconIntColЧ��
'                                End If
'                                Exit Function
'                            End If
'                        End If
'                    End If
                    'ֻ������ſ��ܲ�����˼�¼
                    '   3������ⷿ��ҩ����������ҩ����ҩ���������������С�ڵ����㣬˵��������ҩƷ�޿�棬�������ͣ������棩
                    If mint�༭״̬ <> 2 Then
                        If ((bln����ⷿ And blnҩ�����) Or (Not bln����ⷿ And blnҩ������)) Then
                            If Val(.TextMatrix(intLop, mconIntCol����)) = 0 And Val(.TextMatrix(intLop, mconIntColʵ������)) <> 0 Then
                                MsgBox "��" & intLop & "�е�ҩƷ������ҩƷ���޿�棬�������ͣ�", vbInformation, gstrSysName
                                mshBill.SetFocus
                                .Row = intLop
                                .Col = mconIntColʵ������
                                .MsfObj.TopRow = intLop
                                Exit Function
                            End If
                        End If
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntCol��д����)) > 9999999999# Then
                        MsgBox "��" & intLop & "��ҩƷ����д�������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntCol��д����
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntColʵ������)) > 9999999999# Then
                        MsgBox "��" & intLop & "��ҩƷ��ʵ���������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = mconIntColʵ������
                        Exit Function
                    End If
                    
                    If Val(.TextMatrix(intLop, mconIntCol�ɹ����)) > 9999999999999# Then
                        MsgBox "��" & intLop & "��ҩƷ�ĳɱ������������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = IIf(.ColData(mconIntCol��д����) = 4, mconIntCol��д����, mconIntColʵ������)
                        Exit Function
                    End If
                    If Val(.TextMatrix(intLop, mconIntCol�ۼ۽��)) > 9999999999999# Then
                        MsgBox "��" & intLop & "��ҩƷ���ۼ۽����������ݿ��ܹ������" & vbCrLf & "���Χ9999999999999�����飡", vbInformation + vbOKOnly, gstrSysName
                        mshBill.SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        .Col = IIf(.ColData(mconIntCol��д����) = 4, mconIntCol��д����, mconIntColʵ������)
                        Exit Function
                    End If
                    
                    If .TextMatrix(intLop, mconIntCol��������) = "1" And (.TextMatrix(intLop, mconIntCol����) = "" Or .TextMatrix(intLop, mconIntCol����) = "") And .TextMatrix(intLop, 0) <> "" Then
                        MsgBox "��" & intLop & "�У����ⷿ�Ƿ�����������¼�����źͲ��أ�", vbInformation, gstrSysName
                        .SetFocus
                        .Row = intLop
                        .MsfObj.TopRow = intLop
                        If .TextMatrix(intLop, mconIntCol����) = "" Then
                            .Col = mconIntCol����
                        Else
                            .Col = mconIntCol����
                        End If
                        Exit Function
                    End If
                End If
            Next
        Else
            Exit Function
        End If
    End With
    
    ValidData = True
End Function



Private Function SaveCard(Optional ByVal blnǿ�Ʊ��� As Boolean = False) As Boolean
    Dim chrNo As Variant
    Dim lngSerial As Long
    Dim lngStockID As Long
    Dim lngEnterStockID As Long
    Dim lngDrugID As Long
    Dim strBatchNo As String
    Dim lngBatchID As Long
    Dim strProducingArea As String
    Dim datTimeLimit As String
    Dim dblQuantity As Double
    Dim dblRealNum As Double
    Dim dblPurchasePrice As Double
    Dim dblPurchaseMoney As Double
    Dim dblSalePrice As Double
    Dim dblSaleMoney As Double
    Dim dblMistakePrice As Double
    Dim strBrief As String
    Dim strBooker As String
    Dim datBookDate As String
    Dim strAssessor As String
    Dim datAssessDate As String
    Dim arrSql As Variant
    Dim intRow As Integer
    Dim blnTran As Boolean
    Dim lng�ϴι�Ӧ��ID As Long
    Dim strCheckString As String
    Dim str��׼�ĺ� As String
    Dim n As Integer
    
    arrSql = Array()
    SaveCard = False
    
    '���õ����Ƿ��ڽ���༭����󣬱���������Ա�޸ģ������ת���ƿ�ĵ��ݲ�����
    If mint�༭״̬ = 2 Or (blnǿ�Ʊ��� And mint�༭״̬ <> 11) Then        '�޸�
        mstrTime_End = GetBillInfo(6, mstr���ݺ�)
        If mstrTime_End = "" Then
            MsgBox "�õ����Ѿ�����������Աɾ����", vbInformation, gstrSysName
            Exit Function
        End If
        strCheckString = CheckBill(mstr���ݺ�)
        If strCheckString <> "" Then
            MsgBox strCheckString, vbInformation, gstrSysName
            Exit Function
        End If
        If mstrTime_End > mstrTime_Start Then
            MsgBox "�õ����Ѿ�����������Ա�༭�����˳������ԣ�", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    On Error GoTo errHandle
     
    With mshBill
        chrNo = Trim(txtNo)
        If chrNo = "" Then chrNo = zlDataBase.GetNextNo(26, Me.cboStock.ItemData(Me.cboStock.ListIndex))
        If IsNull(chrNo) Then Exit Function
        txtNo.Tag = chrNo
        lngStockID = cboStock.ItemData(cboStock.ListIndex)
        lngEnterStockID = cboEnterStock.ItemData(cboEnterStock.ListIndex)
        strBrief = Trim(txtժҪ.Text)
        strBooker = Txt������
        If Txt��������.Caption = "" Or Not IsDate(Txt��������.Caption) Then
            datBookDate = Format(zlDataBase.Currentdate, "yyyy-mm-dd hh:mm:ss")
        Else
            datBookDate = Format(Txt��������.Caption, "yyyy-mm-dd hh:mm:ss")
        End If
        strAssessor = Txt�����
        
        'Modified by ZYB 2004-05-16 ������������ǿ�Ʊ�����ʼ����
        If blnǿ�Ʊ��� Then blnTran = True
        
        '�����ת���ƿ�ĵ��ݲ�����
        If mint�༭״̬ = 2 Or (blnǿ�Ʊ��� And mint�༭״̬ <> 11) Then        '�޸�
            If Not mbln���쵥 Then
                gstrSQL = "zl_ҩƷ�ƿ�_Delete('" & mstr���ݺ� & "')"
            Else
                gstrSQL = "zl_ҩƷ����_Delete('" & mstr���ݺ� & "')"
            End If
            
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = "0;" & gstrSQL
        End If
        
        
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
            If .TextMatrix(intRow, 0) <> "" Then
                lngDrugID = .TextMatrix(intRow, 0)
                strProducingArea = .TextMatrix(intRow, mconIntCol����)
                strBatchNo = .TextMatrix(intRow, mconIntCol����)
                lngBatchID = .TextMatrix(intRow, mconIntCol����)
                datTimeLimit = IIf(Trim(.TextMatrix(intRow, mconIntColЧ��)) = "", "", .TextMatrix(intRow, mconIntColЧ��))
                
                If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And datTimeLimit <> "" Then
                    '����ΪʧЧ��������
                    datTimeLimit = Format(DateAdd("D", 1, datTimeLimit), "yyyy-mm-dd")
                End If
                
                dblQuantity = GetFormat(Val(.TextMatrix(intRow, mconIntCol��д����)) * .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_����)
                dblRealNum = GetFormat(Val(.TextMatrix(intRow, mconIntColʵ������)) * .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_����)
                
                If Val(.TextMatrix(intRow, mconintCol��ʵ����)) <> 0 Then
                    If Val(GetFormat(Val(.TextMatrix(intRow, mconintCol��ʵ����)) / Val(.TextMatrix(intRow, mconIntCol����ϵ��)), mintNumberDigit)) = Val(.TextMatrix(intRow, mconIntCol��д����)) Then
                        If dblQuantity = dblRealNum Then
                            dblQuantity = Val(.TextMatrix(intRow, mconintCol��ʵ����))
                            dblRealNum = Val(.TextMatrix(intRow, mconintCol��ʵ����))
                        End If
                    End If
                End If
                
'                dblPurchasePrice = FormatEx(Val(.TextMatrix(intRow, mconIntCol�ɹ���)) / .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_�ɱ���)
                dblPurchasePrice = Get�ɱ���(lngDrugID, lngStockID, lngBatchID)
                dblPurchaseMoney = Val(.TextMatrix(intRow, mconIntCol�ɹ����))
'                dblSalePrice = FormatEx(.TextMatrix(intRow, mconIntCol�ۼ�) / .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_���ۼ�)
                dblSalePrice = Get�ۼ�(Split(.TextMatrix(intRow, mconIntCol���Ч��), "||")(1) = 1, lngDrugID, lngStockID, lngBatchID)
                dblSaleMoney = Val(.TextMatrix(intRow, mconIntCol�ۼ۽��))
                dblMistakePrice = Val(.TextMatrix(intRow, mconintCol���))
                lng�ϴι�Ӧ��ID = .TextMatrix(intRow, mconIntCol�ϴι�Ӧ��ID)
'                If Val(.TextMatrix(intRow, mconIntCol���)) = 0 Then
'                    lngSerial = 2 * intRow - 1
'                Else
'                    lngSerial = Val(.TextMatrix(intRow, mconIntCol���))
'                End If
                lngSerial = 2 * intRow - 1
                .TextMatrix(intRow, mconIntCol���) = lngSerial
                
                str��׼�ĺ� = IIf(Trim(.TextMatrix(intRow, mconIntCol��׼�ĺ�)) = "", "", .TextMatrix(intRow, mconIntCol��׼�ĺ�))

                If Not mbln���쵥 Or blnǿ�Ʊ��� Then
                    gstrSQL = "zl_ҩƷ�ƿ�_INSERT("
                Else
                    gstrSQL = "zl_ҩƷ����_INSERT("
                End If
                
                'NO
                gstrSQL = gstrSQL & "'" & chrNo & "'"
                '���
                gstrSQL = gstrSQL & "," & lngSerial
                '�ⷿID
                gstrSQL = gstrSQL & "," & lngStockID
                '�Է�����ID
                gstrSQL = gstrSQL & "," & lngEnterStockID
                'ҩƷID
                gstrSQL = gstrSQL & "," & lngDrugID
                '����
                gstrSQL = gstrSQL & "," & lngBatchID
                '��д����
                gstrSQL = gstrSQL & "," & dblQuantity
                'ʵ������
                gstrSQL = gstrSQL & "," & dblRealNum
                '�ɱ���
                gstrSQL = gstrSQL & "," & dblPurchasePrice
                '�ɱ����
                gstrSQL = gstrSQL & "," & dblPurchaseMoney
                '���ۼ�
                gstrSQL = gstrSQL & "," & dblSalePrice
                '���۽��
                gstrSQL = gstrSQL & "," & dblSaleMoney
                '���
                gstrSQL = gstrSQL & "," & dblMistakePrice
                '������
                gstrSQL = gstrSQL & ",'" & strBooker & "'"
                '����
                gstrSQL = gstrSQL & ",'" & strProducingArea & "'"
                '����
                gstrSQL = gstrSQL & ",'" & strBatchNo & "'"
                'Ч��
                gstrSQL = gstrSQL & "," & IIf(datTimeLimit = "", "Null", "to_date('" & Format(datTimeLimit, "yyyy-MM-dd") & "','yyyy-mm-dd')")
                'ժҪ
                gstrSQL = gstrSQL & ",'" & strBrief & "'"
                '��������
                gstrSQL = gstrSQL & ",to_date('" & datBookDate & "','yyyy-mm-dd HH24:MI:SS')"
                '��ҩ��λID
                gstrSQL = gstrSQL & "," & IIf(lng�ϴι�Ӧ��ID = 0, "NULL", lng�ϴι�Ӧ��ID)
                '��׼�ĺ�
                gstrSQL = gstrSQL & ",'" & str��׼�ĺ� & "'"
                '���췽ʽ
                gstrSQL = gstrSQL & "," & IIf(mintApplyType = -1, "Null", mintApplyType)
                '����ʱ��
                gstrSQL = gstrSQL & ",'" & mstrEndTime & "'"
                gstrSQL = gstrSQL & ")"
                
                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = CStr(lngDrugID) & ";" & gstrSQL
            End If
            recSort.MoveNext
        Next
        
        If Not ExecuteSql(arrSql, MStrCaption, False, Not blnǿ�Ʊ���) Then Exit Function
        
        If Not blnǿ�Ʊ��� Then gcnOracle.CommitTrans: blnTran = False
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveCard = True
    Exit Function
errHandle:
    If Not blnǿ�Ʊ��� Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetҩƷ��������(ByVal lngҩƷID As Long, ByVal lng�ⷿID As Long) As Boolean
'���ܣ����ݴ�������ҩƷid�Ϳⷿid�жϸ�ҩƷ���ڿⷿ�Ƿ����
'���� lngҩƷid����Ҫ����ҩƷ
'���� lng�ⷿid����Ҫ���Ŀⷿ
'����ֵ�� true-�����ҩƷ���ڿⷿ������false-�����ҩƷ���ڿⷿ������
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim int�������� As Integer      '0-������;1-����
    Dim intҩ����� As Integer      '0-������;1-����
    Dim intҩ������ As Integer      '0-������;1-����
    Dim bln�Ƿ����ҩ������ As Boolean  'True-����ҩ������;False-������ҩ������
    
    If lngҩƷID = 0 Then Exit Function
    On Error GoTo errHandle
    strSQL = "SELECT NVL(ҩ�����, 0) ҩ�����,NVL(ҩ������, 0) ҩ������ " & _
            " From ҩƷ��� WHERE ҩƷID = [1] "
    Set rsTemp = zlDataBase.OpenSQLRecord(strSQL, "ȡҩƷ�ⷿ��������", lngҩƷID)
    
    If rsTemp.RecordCount > 0 Then
        intҩ����� = rsTemp!ҩ�����
        intҩ������ = rsTemp!ҩ������
    End If
    
    If intҩ������ = 1 Then     '���ҩ�����������������Ϊ1
        GetҩƷ�������� = True
    Else
        If intҩ����� = 1 Then
            strSQL = "SELECT ����ID From ��������˵�� " & _
                    " WHERE ((�������� LIKE '%ҩ��') OR (�������� LIKE '�Ƽ���')) AND ����ID = [1] "
            Set rsTemp = zlDataBase.OpenSQLRecord(strSQL, "ȡ��������", lng�ⷿID)
            
            bln�Ƿ����ҩ������ = (rsTemp.RecordCount > 0)
                    
            If bln�Ƿ����ҩ������ Then
                GetҩƷ�������� = False
            Else
                GetҩƷ�������� = True
            End If
        End If
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveStrike() As Boolean
    '���ʳ��� Write by zyb, ##20021016##
    Dim �д�_IN As Integer
    Dim ԭ��¼״̬_IN As Integer
    Dim NO_IN As String
    Dim ���_IN As Integer
    Dim ҩƷID_IN As Long
    Dim ��������_IN As Double
    Dim ������_IN As String
    Dim ��������_IN  As String
    Dim intRow As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim n As Integer
    Dim ժҪ_IN As String
    Dim strҩƷID As String
    Dim i As Integer
    Dim arrSql As Variant
    Dim j As Integer
    Dim strҩƷ As String
    
    SaveStrike = False
    arrSql = Array()
    
    With mshBill
        For intRow = 1 To .rows - 1
            '����������������С����
            If Val(.TextMatrix(intRow, mconIntColʵ������)) <> 0 Then
                If Not ��ͬ����(Val(.TextMatrix(intRow, mconIntCol��д����)), Val(.TextMatrix(intRow, mconIntColʵ������))) Then
                    MsgBox "������Ϸ��ĳ�����������" & intRow & "�У���", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            
            '�����������Ƿ��㹻����������Ϊ�������ʱ������
            '����ҩƷ�����μ�飬������ҩƷ�����б�������������飻����ֻ�жϳ����ⷿ��������
            If mint��������ⷿ <> 0 And .TextMatrix(intRow, 0) <> "" Then
                If mbln�¿������� = True And mint����ʽ = 2 Then
                    '�����
                Else
                    If .TextMatrix(intRow, mconIntColʵ������) = .TextMatrix(intRow, mconIntCol��д����) Then
                        ��������_IN = .TextMatrix(intRow, mconintCol��ʵ����)
                    Else
                        ��������_IN = GetFormat(.TextMatrix(intRow, mconIntColʵ������) * .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_����)
                    End If
                    
                    '����ҩƷ�����μ�飬������ҩƷ�����б�������������飻����ֻ�жϳ����ⷿ��������
                    If GetҩƷ��������(.TextMatrix(intRow, 0), Val(cboEnterStock.ItemData(cboEnterStock.ListIndex))) = False Then
                        For j = 1 To .rows - 1
                            If intRow <> j Then
                                If .TextMatrix(intRow, 0) = .TextMatrix(j, 0) And .TextMatrix(intRow, 0) <> "" And .TextMatrix(j, 0) <> "" Then
                                    If .TextMatrix(j, mconIntColʵ������) = .TextMatrix(j, mconIntCol��д����) Then
                                        ��������_IN = ��������_IN + .TextMatrix(j, mconintCol��ʵ����)
                                    Else
                                        ��������_IN = ��������_IN + GetFormat(.TextMatrix(j, mconIntColʵ������) * .TextMatrix(j, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_����)
                                    End If
                                End If
                            End If
                        Next
                    End If
                    If CheckStrickUsable(6, Val(cboEnterStock.ItemData(cboEnterStock.ListIndex)), Val(.TextMatrix(intRow, 0)), .TextMatrix(intRow, mconIntColҩ��), _
                        Val(.TextMatrix(intRow, mconIntCol����)), Val(��������_IN), mint��������ⷿ, Trim(txtNo.Tag), Val(.TextMatrix(intRow, mconIntCol���)) + 1) = False Then
                        .Row = intRow
                        .MsfObj.TopRow = intRow
                        Exit Function
                    End If
                    ��������_IN = 0
                End If
            End If
        Next
        
        '��ͨ����˳������ʵ������
        If mint�༭״̬ = 6 And mint����ʽ <> 1 Then
            strҩƷ = CheckNumStock(mshBill, Val(cboEnterStock.ItemData(cboEnterStock.ListIndex)), 0, mconIntCol����, mconIntColʵ������, mconIntCol����ϵ��, 2, 0, mintNumberDigit)
            If strҩƷ <> "" Then
                If mint��������ⷿ = 1 Then '��������
                    If MsgBox(strҩƷ & " ҩƷ��ʵ�ʿ�桱���㣬�Ƿ������", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                ElseIf mint��������ⷿ = 2 Then '�����ֹ
                    MsgBox strҩƷ & " ҩƷ��ʵ�ʿ�桱���㣬���ܳ�����", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
        
        NO_IN = Trim(txtNo.Tag)
        ������_IN = UserInfo.�û�����
        ��������_IN = Format(zlDataBase.Currentdate, "yyyy-mm-dd hh:mm:ss")
        ԭ��¼״̬_IN = mint��¼״̬
        ժҪ_IN = Trim(txtժҪ.Text)
        
        On Error GoTo errHandle
        
        �д�_IN = 0
        
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
            If .TextMatrix(intRow, 0) <> "" And Val(.TextMatrix(intRow, mconIntColʵ������)) <> 0 Then
                �д�_IN = �д�_IN + 1
                
                ҩƷID_IN = .TextMatrix(intRow, 0)
                strҩƷID = IIf(strҩƷID = "", "", strҩƷID & ",") & ҩƷID_IN
                If .TextMatrix(intRow, mconIntColʵ������) = .TextMatrix(intRow, mconIntCol��д����) Then
                    ��������_IN = .TextMatrix(intRow, mconintCol��ʵ����)
                Else
                    ��������_IN = GetFormat(.TextMatrix(intRow, mconIntColʵ������) * .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_����)
                End If
                
                ��������_IN = IIf(mint�༭״̬ = 6 And mint����ʽ = 2, -1, 1) * ��������_IN
                
                ���_IN = .TextMatrix(intRow, mconIntCol���)
                
                gstrSQL = "ZL_ҩƷ�ƿ�_STRIKE("
                '�д�
                gstrSQL = gstrSQL & �д�_IN
                'ԭ��¼״̬
                gstrSQL = gstrSQL & "," & ԭ��¼״̬_IN
                'NO
                gstrSQL = gstrSQL & ",'" & NO_IN & "'"
                '���
                gstrSQL = gstrSQL & "," & ���_IN
                'ҩƷID
                gstrSQL = gstrSQL & "," & ҩƷID_IN
                '��������
                gstrSQL = gstrSQL & "," & ��������_IN
                '������
                gstrSQL = gstrSQL & ",'" & ������_IN & "'"
                '��������
                gstrSQL = gstrSQL & ",to_date('" & Format(��������_IN, "yyyy-mm-dd hh:mm:ss") & "','yyyy-mm-dd HH24:MI:SS')"
                'ժҪ
                gstrSQL = gstrSQL & "," & IIf(ժҪ_IN = "", "Null", "'" & ժҪ_IN & "'")
                '������ʽ
                gstrSQL = gstrSQL & "," & mint����ʽ
                gstrSQL = gstrSQL & ")"

                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
            End If
            recSort.MoveNext
        Next
        gcnOracle.BeginTrans
        For i = 0 To UBound(arrSql)
            Call zlDataBase.ExecuteProcedure(CStr(arrSql(i)), "SaveCard")
        Next
        gcnOracle.CommitTrans
        
        If �д�_IN = 0 Then
            MsgBox "û��ѡ��һ��ҩƷ����������¼�����������", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        
        '��ʾͣ��ҩƷ
        If strҩƷID <> "" And mint����ʽ <> 1 Then
            Call CheckStopMedi(strҩƷID)
        End If
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveStrike = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    'MsgBox "����ʧ�ܣ����飡", vbInformation + vbOKOnly, gstrSysName
    Call SaveErrLog
End Function

Private Sub ��ʾ�ϼƽ��()
    Dim curTotal As Double, Cur���ʽ�� As Double, Cur���ʲ�� As Double
    Dim intLop As Integer
    
    curTotal = 0: Cur���ʽ�� = 0: Cur���ʲ�� = 0:
    
    With mshBill
        For intLop = 1 To .rows - 1
            curTotal = curTotal + Val(.TextMatrix(intLop, mconIntCol�ɹ����))
            Cur���ʽ�� = Cur���ʽ�� + Val(.TextMatrix(intLop, mconIntCol�ۼ۽��))
        Next
    End With
    
    Cur���ʲ�� = Cur���ʽ�� - curTotal
    lblPurchasePrice.Caption = "�ɱ����ϼƣ�" & GetFormat(curTotal, mintMoneyDigit)
    lblSalePrice.Caption = "�ۼ۽��ϼƣ�" & GetFormat(Cur���ʽ��, mintMoneyDigit)
    lblDifference.Caption = "��ۺϼƣ�" & GetFormat(Cur���ʲ��, mintMoneyDigit)
End Sub

Private Sub ��ʾ�����(ByVal intRow As Integer)
    Dim strNote As String, strUnit As String
    Dim rsUseCount As New Recordset
    Dim int�Է��ⷿ�������� As Integer
    Dim int�������� As Integer
    Dim str��ǰ��� As String
    Dim str�Է���� As String
    Dim dbl�Է���� As Double
    Dim dbl��ǰ��� As Double
    Dim strTemp As String
    
    On Error GoTo errHandle
    With mshBill
        If .TextMatrix(intRow, 0) = "" Then Exit Sub
        
        int�������� = 0
        '�Է��ⷿ��������
        gstrSQL = "Select a.ҩ�����,a.ҩ������,b.�������� " & _
            " From ҩƷ��� a,��������˵�� b " & _
            " Where a.ҩƷid =[1] And b.����id =[2] "
        Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[���տⷿ��������]", Val(.TextMatrix(intRow, 0)), cboEnterStock.ItemData(cboEnterStock.ListIndex))
        Do While Not rsUseCount.EOF
            If int�������� <> 1 Then
                If InStr(rsUseCount!��������, "ҩ��") > 0 Or rsUseCount!�������� = "�Ƽ���" Then
                    int�������� = 1
                    int�Է��ⷿ�������� = rsUseCount!ҩ������
                ElseIf InStr(rsUseCount!��������, "ҩ��") > 0 Then
                    int�������� = 2
                    int�Է��ⷿ�������� = rsUseCount!ҩ�����
                End If
            End If
            rsUseCount.MoveNext
        Loop
        
        strUnit = .TextMatrix(intRow, mconIntCol��λ)
        
        If mint�༭״̬ <> 10 Then
            '�Է��ⷿ���
            gstrSQL = "select sum(��������/" & .TextMatrix(intRow, mconIntCol����ϵ��) & ") as  �������� from ҩƷ��� where �ⷿid=[1] " _
                & " and ҩƷid=[2] " _
                & " and ����=1 "
            If int�Է��ⷿ�������� = 1 And Val(.TextMatrix(intRow, mconIntCol����)) <> 0 And gtype_UserSysParms.P174_ҩƷ�ƿ���ȷ���� = 1 Then
                gstrSQL = gstrSQL & " and nvl(����,0)=[3] "
            End If
            Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ʾ�����]", cboEnterStock.ItemData(cboEnterStock.ListIndex), Val(.TextMatrix(intRow, 0)), Val(.TextMatrix(intRow, mconIntCol����)))
            
            If Not rsUseCount.EOF Then
                dbl�Է���� = GetFormat(Nvl(rsUseCount!��������, 0), mintNumberDigit)
            End If
            rsUseCount.Close
            
            '��ǰ�ⷿ���
            gstrSQL = "select Sum(Nvl(��������,0))/" & .TextMatrix(intRow, mconIntCol����ϵ��) & " as  �������� from ҩƷ��� where �ⷿid=[1] " _
                & " and ҩƷid=[2] " _
                & " and ����=1  "
            
            If Not (mbln���쵥 = True And mbln��ȷ���� = False) Or Val(.TextMatrix(intRow, mconIntCol����)) > 0 Then
                If gtype_UserSysParms.P174_ҩƷ�ƿ���ȷ���� = 1 Then    '���ڿ��Ĳ�ѯ��������û����ȷҩƷ�ƿ�ʱ��ȷ���ε������
                    gstrSQL = gstrSQL & " and nvl(����,0)=[3] "
                End If
            End If
            
            Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ʾ�����]", cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(intRow, 0)), Val(.TextMatrix(intRow, mconIntCol����)))
                
            If rsUseCount.EOF Then
                .TextMatrix(intRow, mconIntCol��������) = 0
            Else
                .TextMatrix(intRow, mconIntCol��������) = GetFormat(IIf(IsNull(rsUseCount.Fields(0)), 0, rsUseCount.Fields(0)), mintNumberDigit)
            End If
            rsUseCount.Close
            
            dbl��ǰ��� = GetFormat(.TextMatrix(intRow, mconIntCol��������), mintNumberDigit)
        Else
            '���ڷ���ʱ����ʾ��ҩƷ�����пⷿ�Ŀ�棬�Ա��ڿⷿ��Ա����ʵ�ʵķ�������
            If gtype_UserSysParms.P174_ҩƷ�ƿ���ȷ���� = 1 And Val(.TextMatrix(intRow, mconIntCol����)) > 0 Then  '���ڿ��Ĳ�ѯ��������û����ȷҩƷ�ƿ�ʱ��ȷ���ε������
                strTemp = " and nvl(a.����,0)=[2] "
            End If
            
            gstrSQL = "Select B.���� AS �ⷿ,Sum(Nvl(A.��������,0))/" & .TextMatrix(intRow, mconIntCol����ϵ��) & " as �������� " _
                & " From ҩƷ��� A,���ű� B" _
                & " Where A.�ⷿID=B.ID And A.����=1 And A.ҩƷid=[1] " & strTemp _
                & " Group By B.����, ҩƷid "
                
            Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ʾ�����]", Val(.TextMatrix(intRow, 0)), Val(.TextMatrix(intRow, mconIntCol����)))
            With rsUseCount
                Do While Not .EOF
                    strNote = strNote & "," & !�ⷿ & ":" & GetFormat(Nvl(!��������, 0), mintNumberDigit) & strUnit
                    If cboStock.Text = !�ⷿ Then
                        dbl��ǰ��� = GetFormat(Nvl(!��������, 0), mintNumberDigit)
                    End If
                    
                    If cboEnterStock.Text = !�ⷿ Then
                        dbl�Է���� = GetFormat(Nvl(!��������, 0), mintNumberDigit)
                    End If
                    .MoveNext
                Loop
            End With
            str��ǰ��� = Mid(strNote, 2)
        End If
        .TextMatrix(intRow, mconIntCol�ⷿ���) = GetFormat(dbl��ǰ���, mintNumberDigit)
        .TextMatrix(intRow, mconIntCol�Է����) = GetFormat(dbl�Է����, mintNumberDigit)
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtIn_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim IntCheck As Integer
    Dim intRow As Integer
    Dim blnEXIST As Boolean
    Dim intIndex As Integer, intCount As Integer
    Dim rsBill As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim lng�ⷿID As Long
    Dim intNO As Integer, strNo As String
    Dim int��װϵ�� As Integer
    Dim lngҩƷID As Long
    Dim blnInput As Boolean
    
    On Error GoTo ErrHand
    '��ʼ׼��
    intNO = 28
    lng�ⷿID = Me.cboStock.ItemData(Me.cboStock.ListIndex)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txtIn.Text) = "" Then Exit Sub
    
    If Len(txtIn.Text) < 8 Then
        txtIn.Text = zlCommFun.GetFullNO(txtIn.Text, intNO, lng�ⷿID)
    End If
    
    '��ҪҪ������е�������
    For IntCheck = 1 To mshBill.rows - 1
        If mshBill.TextMatrix(IntCheck, 0) <> "" Then
            Exit For
        End If
    Next
    If IntCheck <> mshBill.rows Then
        If MsgBox("��ҪҪ������е������ݣ����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        '����ҩƷ��λ�ı�
        mshBill.ClearBill
    End If
    
    'ȡ����������
    IntCheck = 0
    gstrSQL = "Select Nvl(��鷽ʽ,0) ����� From ҩƷ������ Where �ⷿID=[1] "
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[ȡ����������]", Me.cboStock.ItemData(Me.cboStock.ListIndex))

    If rsTemp.RecordCount <> 0 Then IntCheck = rsTemp!�����
    
    gstrSQL = "select �շ�ϸĿid,ִ�п���id from �շ�ִ�п���"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "��ѯ�洢�ⷿ")
    
    '��ȡ�õ��ݲ���ձ��ֻ������ȡ�������ݣ��ҷ��˻�����
    gstrSQL = "SELECT A.ҩƷID,'['||C.����||']' As ����,'['||C.����||']'|| Nvl(F.����,C.����) As ҩƷ����, C.���� As ͨ����,F.���� As ��Ʒ��,C.���,a.����," & _
             "        C.���㵥λ AS ���۵�λ,1 AS ����ϵ��,B.���ﵥλ,B.�����װ,B.סԺ��λ,B.סԺ��װ,B.ҩ�ⵥλ,B.ҩ���װ, " & _
             "        NVL(A.����,0) AS ����,Nvl(C.�Ƿ���,0) AS ʱ��,Nvl(B.ҩ������,0) AS ҩ������,Nvl(B.ҩ�����,0) AS ҩ�����,b.���Ч��,A.����,A.Ч��," & _
             "        B.����ѱ���,B.ָ�������,A.ʵ������,D.��������,D.ʵ�ʽ��,D.ʵ�ʲ��,E.�ּ�,A.��׼�ĺ�,B.ҩƷ��Դ,B.����ҩ��,nvl(d.ƽ���ɱ���,0) as ƽ���ɱ���,a.��ҩ��λid " & _
             " FROM ҩƷ�շ���¼ A,ҩƷ��� B,�շ���ĿĿ¼ C,ҩƷ��� D,�շѼ�Ŀ E,�շ���Ŀ���� F " & _
             " WHERE A.ҩƷID=B.ҩƷID AND B.ҩƷID=C.ID AND B.ҩƷID=D.ҩƷID(+) " & _
             " AND B.ҩƷID=F.�շ�ϸĿID(+) AND F.����(+)=3 AND F.����(+)=1" & _
             " AND B.ҩƷID=E.�շ�ϸĿID(+) AND SYSDATE >=E.ִ������(+)  AND sysdate<=NVL(E.��ֹ����(+),SYSDATE)" & _
             " AND D.�ⷿID(+)=[2] AND D.����(+)=1 AND Nvl(A.����,0)=Nvl(D.����,0)" & _
             " AND A.����=1 AND A.��¼״̬=1 AND NVL(A.��ҩ��ʽ,0)=0 AND A.������� Is Not NULL" & _
             " AND A.NO=[1] And A.�ⷿID+0=[2] " & _
             " ORDER BY A.���"
    Set rsBill = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ȡ�⹺��ⵥ]", txtIn.Text, Me.cboStock.ItemData(Me.cboStock.ListIndex))
             
    If rsBill.RecordCount = 0 Then
        MsgBox "û���ҵ����⹺��ⵥ�ݣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With rsBill
        intRow = 1
        Do While Not .EOF
            lngҩƷID = !ҩƷid
            rsTemp.Filter = " �շ�ϸĿid=" & lngҩƷID & " and ִ�п���id=" & lng�ⷿID
            If rsTemp.RecordCount = 0 Then
                MsgBox "ҩƷ[" & !ҩƷ���� & "]δ��" & cboStock.Text & "�����ô洢���ԣ��������ƿ⣡"
                blnInput = True
            End If
            rsTemp.Filter = ""
            rsTemp.Filter = " �շ�ϸĿid=" & lngҩƷID & " and ִ�п���id=" & cboEnterStock.ItemData(cboEnterStock.ListIndex)
            If rsTemp.RecordCount = 0 Then
                MsgBox "ҩƷ[" & !ҩƷ���� & "]δ��" & cboEnterStock.Text & "�����ô洢���ԣ��������ƿ⣡"
                blnInput = True
            End If
            
            If blnInput = False Then
                '����ƻ����൱�ڶ��ǰ������ƿ⣬��Ҫ��װ������ǰ���ȼ����
                If !ʵ������ > !�������� Then
                    '���λ�ʱ��ҩƷ�����������
                    If !���� <> 0 Or !ʱ�� <> 0 Then
                        MsgBox !ҩƷ���� & "��治�㣬��������⣡��ʱ�ۻ����ҩƷ��", vbInformation, gstrSysName
                        blnInput = True
                    End If
                    'ֻ��ʾһ��
                    If blnInput = False Then
                        Select Case IntCheck
                        Case 1
                            If MsgBox(!ҩƷ���� & "��治�㣬�Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                blnInput = True
                            End If
                        Case 2
                            MsgBox !ҩƷ���� & "��治�㣬�������ƿ⣡", vbInformation, gstrSysName
                            blnInput = True
                        End Select
                    End If
                End If
            End If
            
            'װ������(SetColValue)
            If blnInput = False Then
                int��װϵ�� = Choose(mintUnit, 1, !�����װ, !סԺ��װ, !ҩ���װ)
                If Not SetColValue(intRow, !ҩƷid, !����, !ͨ����, IIf(IsNull(!��Ʒ��), "", !��Ʒ��), _
                    Nvl(!ҩƷ��Դ), Nvl(!����ҩ��), Nvl(!���), Nvl(!����), _
                    Choose(mintUnit, !���۵�λ, !���ﵥλ, !סԺ��λ, !ҩ�ⵥλ), Nvl(!�ּ�, 0), _
                    Nvl(!����), Nvl(!Ч��), Nvl(!���Ч��, 24), !ҩ�����, Nvl(!��������, 0), Nvl(!ʵ�ʽ��, 0), Nvl(!ʵ�ʲ��, 0), _
                    Nvl(!ָ�������, 0), int��װϵ��, Nvl(!����, 0), !ʱ��, _
                    !ҩ������, !��ҩ��λID, IIf(IsNull(!��׼�ĺ�), "", !��׼�ĺ�)) Then
                    mshBill.ClearBill
                    Exit Sub
                End If
    
                '��д�������ɹ��ۡ��ۼ۵���
                mshBill.TextMatrix(intRow, mconIntCol�к�) = intRow
                mshBill.TextMatrix(intRow, mconIntColʵ������) = GetFormat(!ʵ������ / int��װϵ��, mintNumberDigit)
                mshBill.TextMatrix(intRow, mconIntCol��д����) = GetFormat(!ʵ������ / int��װϵ��, mintNumberDigit)
                mshBill.TextMatrix(intRow, mconIntCol�ɹ���) = GetFormat(!ƽ���ɱ��� * int��װϵ��, mintCostDigit)
                mshBill.TextMatrix(intRow, mconIntCol�ɹ����) = GetFormat(Val(mshBill.TextMatrix(intRow, mconIntCol�ɹ���)) * Val(mshBill.TextMatrix(intRow, mconIntColʵ������)), mintMoneyDigit)
                mshBill.TextMatrix(intRow, mconIntCol�ۼ۽��) = GetFormat(Val(mshBill.TextMatrix(intRow, mconIntCol�ۼ�)) * Val(mshBill.TextMatrix(intRow, mconIntColʵ������)), mintMoneyDigit)
                mshBill.TextMatrix(intRow, mconintCol���) = GetFormat(Val(mshBill.TextMatrix(intRow, mconIntCol�ۼ۽��)) - mshBill.TextMatrix(intRow, mconIntCol�ɹ����), mintMoneyDigit)
    
                intRow = intRow + 1
                mshBill.rows = mshBill.rows + 1
            End If
            blnInput = False
            .MoveNext
        Loop
    End With
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    mshBill.ClearBill
End Sub

Private Sub txtժҪ_Change()
   
    mblnChange = True
End Sub

Private Sub txtժҪ_GotFocus()
    
    OpenIme GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDbUser, "���뷨", "")
    With txtժҪ
        .SelStart = 0
        .SelLength = Len(txtժҪ.Text)
    End With
End Sub

Private Sub txtժҪ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
        KeyCode = 0
    End If
End Sub

Private Sub txtժҪ_LostFocus()
    OpenIme
End Sub

'������������бȽ�
Private Function CompareUsableQuantity(ByVal intRow As Integer, ByVal dbl��д���� As Double) As Boolean
    Dim dblUsableQuantity As Double      'ʵ��������Ӧ���������
    Dim numUsedCount As Double
    Dim vardrug As Variant
    Dim dbltotal As Double              'ĳ��ҩƷ�������������
    Dim intLop As Integer
    Dim dblԭ��д���� As Double
    
    'mint�����: 0-�����;1-��飬�������ѣ�2-��飬�����ֹ
    
    CompareUsableQuantity = False
'    If (mbln���쵥 And Not mbln��ȷ����) Then CompareUsableQuantity = True: Exit Function
    
    With mshBill
        If .TextMatrix(intRow, 0) = "" Then Exit Function
        dblUsableQuantity = GetFormat(.TextMatrix(intRow, mconIntCol��������), mintNumberDigit)
        If .TextMatrix(intRow, mconIntCol����) > 0 Or Split(.TextMatrix(intRow, mconIntCol���Ч��), "||")(1) = 1 Then     '�Է�������ʱ��ҩƷ�����
            If mint�༭״̬ = 1 Then
                dbltotal = 0
                For intLop = 1 To .rows - 1
                    If .TextMatrix(intLop, 0) <> "" Then
                        If intLop <> intRow And .TextMatrix(intLop, 0) = .TextMatrix(intRow, 0) And .TextMatrix(intRow, mconIntCol����) = .TextMatrix(intLop, mconIntCol����) Then
                            dbltotal = dbltotal + .TextMatrix(intLop, mconIntCol��д����)
                        End If
                    End If
                Next
                
                If dbl��д���� + dbltotal > dblUsableQuantity Then
                    MsgBox "�Բ����������������" & dbl��д���� & "�������˸�ҩƷ�Ŀ��ÿ��������" & dblUsableQuantity - dbltotal & "���������䣡", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
                
            ElseIf mint�༭״̬ = 2 Or mint�༭״̬ = 3 Or mint�༭״̬ = 10 Then
                numUsedCount = 0
                For Each vardrug In mcolUsedCount
                    If vardrug(0) = .TextMatrix(.Row, 0) & .TextMatrix(.Row, mconIntCol����) Then
                        numUsedCount = vardrug(1)
                        Exit For
                    End If
                Next
                
                dbltotal = 0
                For intLop = 1 To .rows - 1
                    If .TextMatrix(intLop, 0) <> "" Then
                        If intLop <> intRow And .TextMatrix(intLop, 0) = .TextMatrix(intRow, 0) And .TextMatrix(intRow, mconIntCol����) = .TextMatrix(intLop, mconIntCol����) Then
                            dbltotal = dbltotal + .TextMatrix(intLop, mconIntColʵ������)
                        End If
                    End If
                Next
                
                dblԭ��д���� = IIf(mbln�¿�������, numUsedCount, 0)
                
                If dbl��д���� + dbltotal > dblUsableQuantity + dblԭ��д���� Then
                    MsgBox "�Բ����������������" & dbl��д���� & "�������˸�ҩƷ�Ŀ��ÿ��������" & dblUsableQuantity + dblԭ��д���� - dbltotal & "���������䣡", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
                
            End If
            CompareUsableQuantity = True
            Exit Function
        End If
        
        ' ���Ƴ��ⷿ�ǿⷿ��ҩƷ�Ƿ��������ҩƷ������ж�
        
        If mint����� = 0 Then
            '0-�����
        ElseIf mint����� = 1 Then
            '1-��飬��������
            If mint�༭״̬ = 1 Then
                If dbl��д���� > dblUsableQuantity Then
                    If MsgBox("�Բ����������������" & dbl��д���� & "�������˸�ҩƷ�Ŀ��ÿ��������" & dblUsableQuantity & "�����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                End If
            ElseIf mint�༭״̬ = 2 Or mint�༭״̬ = 3 Or mint�༭״̬ = 10 Then
                numUsedCount = 0
                For Each vardrug In mcolUsedCount
                    If vardrug(0) = .TextMatrix(.Row, 0) & .TextMatrix(.Row, mconIntCol����) Then
                        numUsedCount = vardrug(1)
                        Exit For
                    End If
                Next
                
                dblԭ��д���� = IIf(mbln�¿�������, numUsedCount, 0)
                
                If dbl��д���� > dblUsableQuantity + dblԭ��д���� Then
                    If MsgBox("�Բ����������������" & dbl��д���� & "�������˸�ҩƷ�Ŀ��ÿ��������" & dblUsableQuantity + dblԭ��д���� & "�����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                End If
            End If
            
        ElseIf mint����� = 2 Then
            '2-��飬�����ֹ
            If mint�༭״̬ = 1 Then
                dbltotal = 0
                For intLop = 1 To .rows - 1
                    If .TextMatrix(intLop, 0) <> "" Then
                        If intLop <> intRow And .TextMatrix(intLop, 0) = .TextMatrix(intRow, 0) And IIf(.TextMatrix(intLop, mconIntCol����) = "", "0", .TextMatrix(intLop, mconIntCol����)) = "0" Then
                            dbltotal = dbltotal + .TextMatrix(intLop, mconIntCol��д����)
                        End If
                    End If
                Next
                
                If dbl��д���� + dbltotal > dblUsableQuantity Then
                    MsgBox "�Բ����������������" & dbl��д���� & "�������˸�ҩƷ�Ŀ��ÿ��������" & dblUsableQuantity - dbltotal & "���������䣡", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            ElseIf mint�༭״̬ = 2 Or mint�༭״̬ = 3 Or mint�༭״̬ = 10 Then
                numUsedCount = 0
                For Each vardrug In mcolUsedCount
                    If vardrug(0) = .TextMatrix(.Row, 0) & .TextMatrix(.Row, mconIntCol����) Then
                        numUsedCount = vardrug(1)
                        Exit For
                    End If
                Next
                
                dbltotal = 0
                For intLop = 1 To .rows - 1
                    If .TextMatrix(intLop, 0) <> "" Then
                        If intLop <> intRow And .TextMatrix(intLop, 0) = .TextMatrix(intRow, 0) And IIf(.TextMatrix(intLop, mconIntCol����) = "", "0", .TextMatrix(intLop, mconIntCol����)) = "0" Then
                            dbltotal = dbltotal + .TextMatrix(intLop, mconIntColʵ������)
                        End If
                    End If
                Next
                
                dblԭ��д���� = IIf(mbln�¿�������, numUsedCount, 0)
                
                If dbl��д���� + dbltotal > dblUsableQuantity + dblԭ��д���� Then
                    MsgBox "�Բ����������������" & dbl��д���� & "�������˸�ҩƷ�Ŀ��ÿ��������" & dblUsableQuantity + dblԭ��д���� - dbltotal & "���������䣡", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            End If
        End If
            
    End With
    
    CompareUsableQuantity = True
    
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

'��ӡ����
Private Sub printbill()
    Dim int��λϵ�� As Integer
    Dim strNo As String
    
    Select Case mintUnit
        Case mconint�ۼ۵�λ
            int��λϵ�� = 4
        Case mconint���ﵥλ
            int��λϵ�� = 2
        Case mconintסԺ��λ
            int��λϵ�� = 1
        Case mconintҩ�ⵥλ
            int��λϵ�� = 3
    End Select
    strNo = txtNo.Tag
    DoEvents
    FrmBillPrint.ShowME Me, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1304", "zl8_bill_1304"), mint��¼״̬, int��λϵ��, 1304, "ҩƷ����������", strNo
End Sub

'ȡ���ݿ������ŵĳ��ȣ������������е����ų��������ݿ��б���һ����
Private Function GetBatchNoLen() As Integer
    Dim rsBatchNolen As New Recordset
    On Error GoTo errHandle
    
    gstrSQL = "select ���� from ҩƷ�շ���¼ where rownum<1 "
    Call zlDataBase.OpenRecordset(rsBatchNolen, gstrSQL, "ȡ�ֶγ���")
    GetBatchNoLen = rsBatchNolen.Fields(0).DefinedSize
    rsBatchNolen.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function AutoExpend(Optional blnCheck As Boolean = False) As Boolean
    Dim lng�ⷿID As Long, lngҩƷID As Long, lngҩƷID_Last As Long, lng���� As Long
    Dim bln�ⷿ As Boolean, bln���� As Boolean, blnʱ�� As Boolean, blnAddRow As Boolean
    Dim dbl��д���� As Double, dbl�������� As Double, Dbl���� As Double, dbl����ϵ�� As Double
    Dim dbl�ּ� As Currency, dbl�ּ�_ʱ�� As Double, dbl�ɱ��� As Double
    Dim lngCol As Long, lngCols As Long, lngRow As Long, lngRows As Long
    Dim rsTemp As New ADODB.Recordset
    Dim rsCheck As New ADODB.Recordset
    Dim dblʵ������ As Double
    Dim intCount As Integer
            
    '��ҩƷ��¼�����Զ��ֽ⣬����������ҩƷ
    On Error GoTo ErrHand
    Debug.Print "��ʼ�ֽ⣺" & Now
    Screen.MousePointer = 11
    lngRow = 1: lngCols = mshBill.Cols - 1
    lng�ⷿID = Val(cboStock.ItemData(cboStock.ListIndex))
    bln�ⷿ = CheckStockProperty(lng�ⷿID)
    
    Do While True
        lngҩƷID = Val(mshBill.TextMatrix(lngRow, 0))
        dbl�������� = Val(mshBill.TextMatrix(lngRow, mconIntCol��д����))
'        dbl��д���� = Val(mshBill.TextMatrix(lngRow, mconIntColʵ������))
        dbl��д���� = dbl��������
        dbl����ϵ�� = Val(mshBill.TextMatrix(lngRow, mconIntCol����ϵ��))
        lng���� = Val(mshBill.TextMatrix(lngRow, mconIntCol����))
        
        If lngҩƷID = 0 Then Exit Do
        
        '��ȡ��ҩƷ���ڳ���ⷿ�Ƿ������ʱ�۵�����
        If lngҩƷID <> lngҩƷID_Last Then
            lngҩƷID_Last = lngҩƷID
            gstrSQL = " Select Nvl(A.ҩ�����,0) ҩ�����,Nvl(A.ҩ������,0) ҩ������," & _
                      " Nvl(B.�Ƿ���,0) ʱ��,Nvl(P.�ּ�,0) �ּ�,Nvl(A.�ɱ���,0) �ɱ���" & _
                      " From ҩƷ��� A,�շ���ĿĿ¼ B,�շѼ�Ŀ P" & _
                      " Where A.ҩƷID = B.ID And B.ID=P.�շ�ϸĿID And A.ҩƷID =[1] " & _
                      " And Sysdate between P.ִ������ And Nvl(P.��ֹ����,Sysdate)"
            Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ȡ��ҩƷ���ڳ���ⷿ�Ƿ������ʱ�۵�����]", lngҩƷID)
            
            blnʱ�� = (rsTemp!ʱ�� = 1)
            dbl�ּ� = rsTemp!�ּ� * dbl����ϵ��
            bln���� = IIf(bln�ⷿ, (rsTemp!ҩ����� = 1), (rsTemp!ҩ������ = 1))
        End If
        
        '�����ҩƷ�Ƿ���ҩƷ��������Ϊ�㣬��˵����Ҫ�Զ��ֽ�
        blnAddRow = False
        If bln���� = True And lng���� = 0 Then
'            If blnCheck Then
'                If dbl��д���� > Val(mshBill.TextMatrix(lngRow, mconIntCol��������)) Then
'                    MsgBox "��" & lngRow & "�е�ҩƷ�����λ�ʱ��ҩƷ������ҩƷ��ǰ��治�㣬���ܼ�����", vbInformation, gstrSysName
'                    Screen.MousePointer = 0: Exit Function
'                End If
'            End If
            gstrSQL = " Select Nvl(��������,0)/" & dbl����ϵ�� & " As ��������,Nvl(ʵ������,0)/" & dbl����ϵ�� & " As ʵ������," & _
                      " Nvl(ʵ�ʽ��,0) ʵ�ʽ��,Nvl(ʵ�ʲ��,0) ʵ�ʲ��, nvl(ƽ���ɱ���,0) ƽ���ɱ���," & _
                      " Nvl(����,0) ����,�ϴ����� ����,to_char(Ч��,'yyyy-MM-dd') Ч��,�ϴβ��� ����,NVL(�ϴι�Ӧ��ID,0) �ϴι�Ӧ��ID,��׼�ĺ�,nvl(���ۼ�,0)*" & dbl����ϵ�� & " As ���ۼ� " & _
                      " From ҩƷ��� Where �ⷿID=[1] And ҩƷID=[2] And ����=1 And Nvl(��������,0)>0 "
        
            If gtype_UserSysParms.P150_ҩƷ���������㷨 = 0 Then
                gstrSQL = gstrSQL & " Order by Nvl(����,0)"
            ElseIf gtype_UserSysParms.P150_ҩƷ���������㷨 = 1 Then
                gstrSQL = gstrSQL & " Order by Ч��,Nvl(����,0)"
            ElseIf gtype_UserSysParms.P150_ҩƷ���������㷨 = 2 Then
                gstrSQL = gstrSQL & " Order by �ϴ�����,Nvl(����,0)"
            Else
                gstrSQL = gstrSQL & " Order by Nvl(����,0)"
            End If

            Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ȡ��ҩƷ��ָ���������п���¼]", lng�ⷿID, lngҩƷID)
            With rsCheck
                intCount = 0
                Do While Not .EOF
                    intCount = intCount + 1
                    mshBill.Redraw = False
                    '����д��¼
                    blnAddRow = False
                    If .AbsolutePosition <> 1 Then
                        mshBill.MsfObj.AddItem "", lngRow
                        For lngCol = 0 To lngCols
                            mshBill.TextMatrix(lngRow, lngCol) = mshBill.TextMatrix(lngRow - 1, lngCol)
                        Next
                        mshBill.TextMatrix(lngRow, mconIntCol��д����) = "0"
                        mshBill.RowData(lngRow) = mshBill.RowData(lngRow - 1)
                    End If
                    
                    If intCount = 1 Then
                        dblʵ������ = Val(mshBill.TextMatrix(lngRow, mconIntColʵ������))
                    End If
                    
                    '��д���������Ϣ
                    mshBill.TextMatrix(lngRow, mconIntCol�к�) = lngRow
                    mshBill.TextMatrix(lngRow, mconIntCol���) = (lngRow - 1) * 2 + 1
                    mshBill.TextMatrix(lngRow, mconIntCol����) = rsCheck!����
                    mshBill.TextMatrix(lngRow, mconIntCol����) = IIf(IsNull(rsCheck!����), "", rsCheck!����)
                    mshBill.TextMatrix(lngRow, mconIntCol����) = IIf(IsNull(rsCheck!����), "", rsCheck!����)
                    mshBill.TextMatrix(lngRow, mconIntColЧ��) = IIf(IsNull(rsCheck!Ч��), "", rsCheck!Ч��)
                    If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And mshBill.TextMatrix(lngRow, mconIntColЧ��) <> "" Then
                        '����Ϊ��Ч��
                        mshBill.TextMatrix(lngRow, mconIntColЧ��) = Format(DateAdd("D", -1, mshBill.TextMatrix(lngRow, mconIntColЧ��)), "yyyy-mm-dd")
                    End If
                    
                    mshBill.TextMatrix(lngRow, mconIntCol�ϴι�Ӧ��ID) = rsCheck!�ϴι�Ӧ��ID
                    mshBill.TextMatrix(lngRow, mconIntCol��׼�ĺ�) = IIf(IsNull(rsCheck!��׼�ĺ�), "", rsCheck!��׼�ĺ�)
                    
                    '���¼���۸������Ϣ
                    If rsCheck!ʵ������ > 0 Then
                        If Val(mshBill.TextMatrix(lngRow, mconIntCol����)) > 0 Then
                            dbl�ּ�_ʱ�� = IIf(rsCheck!���ۼ� > 0, rsCheck!���ۼ�, rsCheck!ʵ�ʽ�� / rsCheck!ʵ������)
                        Else
                            dbl�ּ�_ʱ�� = rsCheck!ʵ�ʽ�� / rsCheck!ʵ������
                        End If
                    End If
                    
                    If dbl��д���� <= rsCheck!�������� Then
                        Dbl���� = dbl��д����
                    Else
                        Dbl���� = rsCheck!��������
                    End If
                    If Dbl���� > dbl��д���� Then Dbl���� = dbl��д����
                    
                    mshBill.TextMatrix(lngRow, mconIntCol��д����) = GetFormat(Dbl����, mintNumberDigit)
                    mshBill.TextMatrix(lngRow, mconIntColʵ������) = GetFormat(Dbl����, mintNumberDigit)
                    
                    '���⴦����������û�п��ʱ��Ҫ�����ź��ϴβ����Զ����ϣ��޿��������Ϣ��Ӱ�죩���������Ա����
                    If Val(mshBill.TextMatrix(lngRow, mconIntCol��������)) = 1 And Dbl���� = 0 Then
                        gstrSQL = "select �ϴβ���,�ϴ����� from ҩƷ��� where ҩƷid=[1]"
                        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "����������Ϣ", lngҩƷID)
                        mshBill.TextMatrix(lngRow, mconIntCol����) = IIf(IsNull(rsTemp!�ϴβ���), "", rsTemp!�ϴβ���)
                        mshBill.TextMatrix(lngRow, mconIntCol����) = IIf(IsNull(rsTemp!�ϴ�����), "", rsTemp!�ϴ�����)
                    End If
                    
                    If dblʵ������ <> mshBill.TextMatrix(lngRow, mconIntColʵ������) Then
                        mshBill.TextMatrix(lngRow, mconintCol��ʵ����) = GetFormat(Val(mshBill.TextMatrix(lngRow, mconIntColʵ������)) * Val(mshBill.TextMatrix(lngRow, mconIntCol����ϵ��)), mintNumberDigit)
                    End If
                    
                    If Trim(mshBill.TextMatrix(lngRow, mconIntColʵ������)) = "" Then mshBill.TextMatrix(lngRow, mconIntColʵ������) = GetFormat(0, mintNumberDigit)
                    
                    mshBill.TextMatrix(lngRow, mconIntColʵ�ʲ��) = GetFormat(rsCheck!ʵ�ʲ��, mintMoneyDigit)
                    mshBill.TextMatrix(lngRow, mconIntColʵ�ʽ��) = GetFormat(rsCheck!ʵ�ʽ��, mintMoneyDigit)
                    mshBill.TextMatrix(lngRow, mconIntCol��������) = GetFormat(rsCheck!��������, mintMoneyDigit)
                    mshBill.TextMatrix(lngRow, mconIntCol�ۼ�) = GetFormat(IIf(blnʱ��, dbl�ּ�_ʱ��, dbl�ּ�), mintPriceDigit)
                    mshBill.TextMatrix(lngRow, mconIntCol�ۼ۽��) = GetFormat(Val(mshBill.TextMatrix(lngRow, mconIntCol�ۼ�)) * Dbl����, mintMoneyDigit)
                    
                    If Dbl���� <> 0 Then
                        mshBill.TextMatrix(lngRow, mconIntCol�ɹ���) = GetFormat(rsCheck!ƽ���ɱ��� * Val(mshBill.TextMatrix(lngRow, mconIntCol����ϵ��)), mintCostDigit)
                    End If
                    mshBill.TextMatrix(lngRow, mconIntCol�ɹ����) = GetFormat(mshBill.TextMatrix(lngRow, mconIntCol�ɹ���) * Dbl����, mintMoneyDigit)
                    mshBill.TextMatrix(lngRow, mconintCol���) = GetFormat(Val(mshBill.TextMatrix(lngRow, mconIntCol�ۼ۽��)) - Val(mshBill.TextMatrix(lngRow, mconIntCol�ɹ����)), mintMoneyDigit)
                    
                    dbl��д���� = dbl��д���� - Dbl����
                    dbl�������� = dbl�������� - Dbl����
                    If dbl��д���� = 0 Then Exit Do
                    lngRow = lngRow + 1
                    blnAddRow = True
                    .MoveNext
                Loop
                If dbl�������� <> 0 And rsCheck.RecordCount <> 0 Then
                    If blnAddRow Then
                        mshBill.TextMatrix(lngRow - 1, mconIntCol��д����) = GetFormat(dbl�������� + Dbl����, mintNumberDigit)
                    Else
                        mshBill.TextMatrix(lngRow, mconIntCol��д����) = GetFormat(dbl�������� + Dbl����, mintNumberDigit)
                    End If
                End If
            End With
            
            '�������¼Ϊ�㣬��˵��δ���зֽ⣬��Ҫ������������ʵ��������Ϊ��
            If dbl��д���� <> 0 And rsCheck.RecordCount = 0 Then
                mshBill.TextMatrix(lngRow, mconIntCol�к�) = lngRow
                mshBill.TextMatrix(lngRow, mconIntCol���) = (lngRow - 1) * 2 + 1
                mshBill.TextMatrix(lngRow, mconIntColʵ������) = GetFormat(0, mintNumberDigit)
                mshBill.TextMatrix(lngRow, mconIntCol�ۼ۽��) = ""
                mshBill.TextMatrix(lngRow, mconIntCol�ɹ����) = ""
                mshBill.TextMatrix(lngRow, mconintCol���) = ""
                
                '���⴦����������û�п��ʱ��Ҫ�����ź��ϴβ����Զ����ϣ��޿��������Ϣ��Ӱ�죩���������Ա����
                If Val(mshBill.TextMatrix(lngRow, mconIntCol��������)) = 1 Then
                    gstrSQL = "select �ϴβ���,�ϴ����� from ҩƷ��� where ҩƷid=[1]"
                    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "����������Ϣ", lngҩƷID)
                    mshBill.TextMatrix(lngRow, mconIntCol����) = IIf(IsNull(rsTemp!�ϴβ���), "", rsTemp!�ϴβ���)
                    mshBill.TextMatrix(lngRow, mconIntCol����) = IIf(IsNull(rsTemp!�ϴ�����), "", rsTemp!�ϴ�����)
                End If
            End If
        Else
            mshBill.TextMatrix(lngRow, mconIntCol�к�) = lngRow
            mshBill.TextMatrix(lngRow, mconIntCol���) = (lngRow - 1) * 2 + 1
        End If
        If blnAddRow = False Then lngRow = lngRow + 1
    Loop
    
    mblnChange = True
    AutoExpend = True
    mshBill.Redraw = True
    Call ShowColor
    Screen.MousePointer = 0
    Debug.Print "�����ֽ⣺" & Now
    
    If mbln�Զ��ֽ�δ��� = True Then mbln�Զ��ֽ�δ��� = False
    
    Exit Function
ErrHand:
    Screen.MousePointer = 0
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckStockProperty(ByVal lng�ⷿID As Long) As Boolean
    Dim rsCheck As New ADODB.Recordset
    On Error GoTo errHandle
    
    '���ָ���ⷿ��ҩ�⡢ҩ�������Ƽ���(����Ŀⷿ�϶���ҩ�⡢ҩ�����Ƽ����е�һ��)
    gstrSQL = " Select ����ID From ��������˵�� " & _
              " Where (�������� like '%ҩ��' Or �������� like '%�Ƽ���') And ����id=[1] "
    Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[�ж��ǲ���ҩ�����Ƽ���]", lng�ⷿID)
              
    If rsCheck.EOF Then
        CheckStockProperty = True
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InsertRow(ByVal lngRow As Long)
    Dim lngReserve As Long, lngRows As Long
    Dim lngCol As Long, lngCols As Long
    Debug.Print Now
    lngReserve = lngRow
    lngRows = mshBill.rows - 1
    lngCols = mshBill.Cols - 1
    mshBill.rows = mshBill.rows + 1
    
    '����ǰ�м�������ȫ������
    For lngRow = lngRows To lngReserve Step -1
        For lngCol = 0 To lngCols
            mshBill.TextMatrix(lngRow + 1, lngCol) = mshBill.TextMatrix(lngRow, lngCol)
        Next
        mshBill.RowData(lngRow + 1) = mshBill.RowData(lngRow)
        'У���к�
        mshBill.TextMatrix(lngRow + 1, mconIntCol�к�) = lngRow + 1
    Next
    Debug.Print Now
End Sub

Private Sub ShowColor(Optional ByVal lngCurRow As Long = 0)
    '�ڲ��Ļ����ʱ������治��ļ�¼�԰���ɫ��ʾ����
    Dim lngSelect_Row  As Long, lngSelect_Col As Long
    Dim lngҩƷID As Long
    Dim lngColor As Long, lngNewColor As Long '������ڵ���ɫ��Ҫ�ϵ���ɫһ�����򲻴���
    Dim dbl��д���� As Double, dbl�������� As Double
    Dim lngRow As Long, BlnDO As Boolean
    Dim i As Long, j As Long
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHand
    mshBill.Redraw = False
    mblnEnterCell = False
    lngSelect_Row = mshBill.Row: lngSelect_Col = mshBill.Col
    lngRow = IIf(lngCurRow > 0, lngCurRow, 1)
    
    Do While True
        If lngRow > mshBill.rows - 1 Then Exit Do
        mshBill.Row = lngRow: mshBill.Col = mconIntColҩ��
        lngColor = mshBill.MsfObj.CellForeColor
        
        lngҩƷID = Val(mshBill.TextMatrix(lngRow, 0))
        dbl��д���� = Val(mshBill.TextMatrix(lngRow, mconIntCol��д����))
        dbl�������� = Val(mshBill.TextMatrix(lngRow, mconIntCol��������))
        If lngҩƷID = 0 Then Exit Do
        
        gstrSQL = "select decode(ҩ�����,Null,0,ҩ�����) ҩ�����,decode(ҩ������,Null,0,ҩ������) ҩ������ from ҩƷ��� where ҩƷid=[1]"
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "��ѯ����", lngҩƷID)
        
        If rsTemp Is Nothing Then
            Exit Sub
        Else
            If rsTemp!ҩ����� = 1 Or rsTemp!ҩ������ = 1 Then
                '��治���ҩƷ������ɫ
                BlnDO = False
                If dbl�������� < dbl��д���� Then BlnDO = True
                lngNewColor = IIf(BlnDO, &HC0, &H0)
                If lngColor <> lngNewColor Then
                    'ֻ��ҩ���н�����ɫ����
                    j = mshBill.ColData(mconIntColҩ��)
                    If j = 5 Then mshBill.ColData(mconIntColҩ��) = 0
                    mshBill.Col = mconIntColҩ��
                    mshBill.MsfObj.CellForeColor = lngNewColor
                    mshBill.ColData(mconIntColҩ��) = j
                End If
            End If
            If lngCurRow > 0 Then Exit Do
            lngRow = lngRow + 1
        End If
    Loop
    mshBill.Row = lngSelect_Row: mshBill.Col = lngSelect_Col
    mshBill.Redraw = True
    mblnEnterCell = True
    Exit Sub
ErrHand:
    mshBill.Redraw = True
    mblnEnterCell = True
    If ErrCenter = 1 Then Resume
End Sub

Private Function CheckStock() As Boolean
    Dim dbl����ϵ�� As Double, dblʵ������ As Double, dbl��д���� As Double
    Dim lngRow As Long, lngRows As Long, int����� As Integer
    Dim lngҩƷID As Long, lng�ⷿID As Long, lng���� As Long
    Dim bln�ⷿ As Boolean, bln��ҩ As Boolean
    Dim strҩƷID As String, strMsg As String
    Dim rsTemp As ADODB.Recordset
    Dim rsProperty As ADODB.Recordset           'ҩƷ���
    Dim rsCheck As ADODB.Recordset              'ҩƷ���
    Dim arrDrugID As Variant
    Dim i As Integer
    
    On Error GoTo errHandle
    Set rsProperty = New ADODB.Recordset
    With rsProperty
        If .State = 1 Then .Close
        .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
        .Fields.Append "ͨ����", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "ҩ�����", adDouble, 18, adFldIsNullable
        .Fields.Append "ҩ������", adDouble, 18, adFldIsNullable
        .Fields.Append "�Ƿ���", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set rsCheck = New ADODB.Recordset
    With rsCheck
        If .State = 1 Then .Close
        .Fields.Append "ҩƷID", adDouble, 18, adFldIsNullable
        .Fields.Append "����", adDouble, 18, adFldIsNullable
        .Fields.Append "��������", adDouble, 18, adFldIsNullable
        .Fields.Append "ʵ������", adDouble, 18, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    arrDrugID = Array()
    
    '��鵥���и�ҩƷ�Ŀ��
    'mint�����:0-�����;1-��飬�������ѣ�2-��飬�����ֹ
    '������ʱ��ҩƷ���ܴ���
    Debug.Print "��ʼ�����:" & Now
    lngRows = mshBill.rows - 1
    lng�ⷿID = Val(cboStock.ItemData(cboStock.ListIndex))
    bln�ⷿ = CheckStockProperty(lng�ⷿID)
    For lngRow = 1 To lngRows
        lngҩƷID = Val(mshBill.TextMatrix(lngRow, 0))
        If lngҩƷID <> 0 Then
            If InStr(1, "," & strҩƷID & ",", "," & lngҩƷID & ",") = 0 Then
                If Len(IIf(strҩƷID = "", "", strҩƷID & ",") & lngҩƷID) > 4000 Then
                    ReDim Preserve arrDrugID(UBound(arrDrugID) + 1)
                    arrDrugID(UBound(arrDrugID)) = strҩƷID
                    strҩƷID = lngҩƷID
                Else
                    strҩƷID = IIf(strҩƷID = "", "", strҩƷID & ",") & lngҩƷID
                End If
            End If
        End If
    Next
    
    If strҩƷID = "" And UBound(arrDrugID) < 0 Then
        CheckStock = True
        Exit Function
    ElseIf strҩƷID <> "" Then
        ReDim Preserve arrDrugID(UBound(arrDrugID) + 1)
        arrDrugID(UBound(arrDrugID)) = strҩƷID
    End If
    
    '��ȡ������������ҩƷ������
    gstrSQL = " Select  A.ҩƷID,'['||B.����||']'||B.���� ͨ����,A.ҩ�����,A.ҩ������,B.�Ƿ���" & _
              " From ҩƷ��� A,�շ���ĿĿ¼ B" & _
              " Where A.ҩƷID=B.ID And A.ҩƷID in (select * from Table(Cast(f_Num2List([1]) As zlTools.t_NumList))) "
    
    For i = 0 To UBound(arrDrugID)
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "��ȡ������������ҩƷ������", CStr(arrDrugID(i)))
        
        If Not rsTemp.EOF Then
            Do While Not rsTemp.EOF
                With rsProperty
                    .AddNew
                    !ҩƷid = rsTemp!ҩƷid
                    !ͨ���� = rsTemp!ͨ����
                    !ҩ����� = rsTemp!ҩ�����
                    !ҩ������ = rsTemp!ҩ������
                    !�Ƿ��� = rsTemp!�Ƿ���
                    
                    .Update
                End With
                rsTemp.MoveNext
            Loop
        End If
    Next
    
    gstrSQL = "Select a.ҩƷid, Nvl(a.����, 0) As ����, Sum(Nvl(��������, 0)) As ��������,Sum(Nvl(ʵ������, 0)) As ʵ������ " & _
        " From ҩƷ��� A, ҩƷ��� C" & _
        " Where a.�ⷿid = [1] And a.ҩƷid = c.ҩƷid And a.���� = 1 And c.ҩƷid in (select * from Table(Cast(f_Num2list([2]) As Zltools.t_Numlist))) " & _
        " Group By a.ҩƷid, Nvl(a.����, 0) "
    For i = 0 To UBound(arrDrugID)
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ȡ������������ҩƷ�ĵ�ǰ���]", lng�ⷿID, CStr(arrDrugID(i)))
    
        If Not rsTemp.EOF Then
            Do While Not rsTemp.EOF
                With rsCheck
                    .AddNew
                    !ҩƷid = rsTemp!ҩƷid
                    !���� = rsTemp!����
                    !�������� = rsTemp!��������
                    !ʵ������ = rsTemp!ʵ������
                    
                    .Update
                End With
                rsTemp.MoveNext
            Loop
        End If
    Next
    
    '���ÿ��ҩƷ
    For lngRow = 1 To lngRows
        lngҩƷID = Val(mshBill.TextMatrix(lngRow, 0))
        If lngҩƷID <> 0 Then
            lng���� = Val(mshBill.TextMatrix(lngRow, mconIntCol����))
            dbl����ϵ�� = Val(mshBill.TextMatrix(lngRow, mconIntCol����ϵ��))
            dbl��д���� = Val(mshBill.TextMatrix(lngRow, mconIntColʵ������))
            
            dblʵ������ = 0
            '���Ҹ�ҩƷ�Ŀ���¼
            rsCheck.Filter = "ҩƷID=" & lngҩƷID & " And ����=" & lng����
            
            If rsCheck.RecordCount <> 0 Then
                If mint�༭״̬ = 10 Then   '�����ÿ��������ж�
                    dblʵ������ = Val(GetFormat(Nvl(rsCheck!��������, 0) / dbl����ϵ��, mintNumberDigit))
                Else    '�����ʵ�������ж�
                    dblʵ������ = Val(GetFormat(Nvl(rsCheck!ʵ������, 0) / dbl����ϵ��, mintNumberDigit))
                End If
            End If
            
            '�������ʵ����������
            If Not (dblʵ������ >= dbl��д����) Then
                int����� = mint�����
                '�����ҩƷ��ʱ�ۻ��������治�㲻������⣬�൱�ڽ�ֹ����
                rsProperty.Filter = "ҩƷID=" & lngҩƷID
                bln��ҩ = (IIf(bln�ⷿ, (rsProperty!ҩ����� = 1), (rsProperty!ҩ������ = 1)) Or (rsProperty!�Ƿ��� = 1))
                strMsg = ""
                If bln��ҩ Then
                    int����� = 2
                    '���������ҩƷ��������С�ڵ����㣬˵��δִ�зֽ⹦��
                    If lng���� <= 0 And IIf(bln�ⷿ, (rsProperty!ҩ����� = 1), (rsProperty!ҩ������ = 1)) Then
                        strMsg = "������ִ�зֽ⹦����ȷ����ҩƷ�ĳ������Σ�"
                    End If
                End If
                
                '��λ��������
                mshBill.Row = lngRow
                mshBill.MsfObj.TopRow = lngRow
                '���������̽�����ʾ���ֹ
                Select Case int�����
                Case 1  '����ʾ
                    Debug.Print "�޿���˳�:" & Now
                    If MsgBox(rsProperty!ͨ���� & "�Ŀ�治�㣬�Ƿ������" & strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                Case 2
                    Debug.Print "�޿���˳�:" & Now
                    MsgBox rsProperty!ͨ���� & "�Ŀ�治�㣡" & strMsg, vbInformation, gstrSysName
                    Exit Function
                End Select
            End If
        End If
    Next
    
    rsCheck.Filter = 0
    rsProperty.Filter = 0
    CheckStock = True
    Debug.Print "��ɼ����:" & Now
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SendPhysic() As Boolean
    Dim rsTemp As New ADODB.Recordset
    '��鵱ǰ�����Ƿ��ѷ���
    On Error GoTo ErrHand

    gstrSQL = "Select ��ҩ���� From ҩƷ�շ���¼ " & _
              "Where ����=6 And NO=[1] And Rownum<2"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��鵱ǰ�����Ƿ��ѷ���]", Me.txtNo.Tag)
              
    If (Nvl(rsTemp!��ҩ����) = "") Then
        MsgBox "�õ����ѱ���������Աȡ�����ͣ���������գ�", vbInformation, gstrSysName
        Exit Function
    End If
    SendPhysic = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub SetForeColor_ROW(ByVal lngColor As Long)
    Dim i As Integer, j As Integer
    Dim intCol As Integer
    '����ĳ�е���ɫ
    With mshBill
        intCol = .Col
        mblnEnterCell = False
        For i = mconIntColҩ�� To .Cols - 1
            j = .ColData(i)
            If .ColData(i) = 5 Then .ColData(i) = 0
            .Col = i
            .MsfObj.CellForeColor = lngColor
            .ColData(i) = j
        Next
        .Col = intCol
        mblnEnterCell = True
    End With
End Sub

Private Function IsSelf_Command(ByVal lngҩƷID As Long) As Boolean
    '�ж��Ƿ�Ϊ����ҩƷ��������ⷿ���Ƽ��ң������Ƽ��ҵ����ԣ�
    Dim bln����ҩƷ As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    '�������ⷿ
    gstrSQL = "Select 1 From ��������˵�� Where ����ID=[1] And ��������='�Ƽ���'"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[�������ⷿ]", cboEnterStock.ItemData(cboEnterStock.ListIndex))
    
    If rsTemp.RecordCount = 0 Then Exit Function
    
    '����Ƿ�������ҩƷ
    gstrSQL = "Select Nvl(����ҩƷ,0) As ����ҩƷ From ҩƷ��� Where ҩƷID=[1] "
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[����Ƿ�����ҩƷ]", lngҩƷID)
    
    bln����ҩƷ = (rsTemp!����ҩƷ = 1)
    '��ȡ�������ҩƷ
    If bln����ҩƷ Then
        gstrSQL = "Select ԭ��ҩƷID,����,��ĸ From ����ҩƷ���� Where ����ҩƷID=[1] "
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ȡ�������ҩƷ]", lngҩƷID)
        bln����ҩƷ = (rsTemp.RecordCount <> 0)
    End If
    
    IsSelf_Command = bln����ҩƷ
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetMaterial(ByVal lngҩƷID As Long) As ADODB.Recordset
    '��ȡ����ҩƷ��ԭ��ҩƷ��Ϣ
    Dim rsMaterial As New ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "" & _
        " Select B.ҩƷID,Nvl(B.ҩ�����,0) As ҩ�����,Nvl(B.ҩ������,0) As ҩ������,C.���� AS ҩƷ����,D.���� As ��Ʒ��,C.���� As ͨ����," & _
        "        B.ҩƷ��Դ,B.����ҩ��,C.���,C.����,C.���㵥λ AS �ۼ۵�λ,B.���ﵥλ,B.�����װ,B.סԺ��λ,B.סԺ��װ,B.ҩ�ⵥλ,B.ҩ���װ,Nvl(C.�Ƿ���,0) As ʱ��," & _
        "        E.�ּ� AS �ۼ�,Nvl(F.����,0) As ����,F.�ϴ����� As ����,F.Ч�� As Ч��,Nvl(B.���Ч��,0) As ���Ч��,Nvl(F.��������,0) As ��������," & _
        "        Nvl(F.ʵ�ʽ��,0) As ʵ�ʽ��,Nvl(F.ʵ�ʲ��,0) As ʵ�ʲ��,Nvl(B.ָ�������,0) As ָ�������,Nvl(F.�ϴι�Ӧ��ID,0) �ϴι�Ӧ��ID,F.��׼�ĺ� " & _
        " From ����ҩƷ���� A,ҩƷ��� B,�շ���ĿĿ¼ C,�շ���Ŀ���� D,�շѼ�Ŀ E," & _
        "      (Select �ⷿID,ҩƷID,����,�ϴ�����,Ч��,��������,ʵ�ʽ��,ʵ�ʲ��,�ϴι�Ӧ��ID,��׼�ĺ� From ҩƷ���" & _
        "      Where (�ⷿID,ҩƷID,Nvl(����,0)) In" & _
        "           (Select A.�ⷿID,A.ҩƷID,Min(Nvl(A.����,0)) From ҩƷ��� A,����ҩƷ���� B" & _
        "            Where A.�ⷿID = [1] And A.ҩƷID = B.ԭ��ҩƷID And A.���� = 1 And B.����ҩƷID =[2] " & _
        "            Group By A.�ⷿID,A.ҩƷID)) F" & _
        " Where A.����ҩƷID = [2] And A.ԭ��ҩƷID = B.ҩƷID And B.ҩƷID = C.Id" & _
        " And B.ҩƷID=D.�շ�ϸĿId(+) And D.����(+)=3 And D.����(+)=1" & _
        " And B.ҩƷID=E.�շ�ϸĿID And ((Sysdate Between ִ������ And ��ֹ����) Or ��ֹ���� Is Null )" & _
        " And B.ҩƷID=F.ҩƷID"
    Set rsMaterial = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ȡ����ҩƷ��ԭ��ҩƷ��Ϣ]", cboStock.ItemData(cboStock.ListIndex), lngҩƷID)
    Set GetMaterial = rsMaterial
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckRedo(ByVal rsTemp As ADODB.Recordset) As String
    '���ܣ���������б�������ҩƷ����ѡ���ҩƷ�Ƿ��ظ����Դ����ж���Ҫ����������
    Dim i As Integer
    Dim strTemp As String
    Dim str���� As String
    
    rsTemp.MoveFirst
    str���� = ""
    Do While Not rsTemp.EOF
        If gtype_UserSysParms.P174_ҩƷ�ƿ���ȷ���� = 0 Then
            str���� = "0"
        Else
            str���� = IIf(IsNull(rsTemp!����), "0", rsTemp!����)
        End If
        If InStr(1, strTemp, rsTemp!ҩƷid & "," & str����) = 0 Then
            strTemp = strTemp & rsTemp!ҩƷid & "," & str���� & "|"
        End If
        rsTemp.MoveNext
    Loop
    
    With mshBill
        For i = 1 To .rows - 2
            If InStr(1, strTemp, .TextMatrix(i, 0) & "," & .TextMatrix(i, mconIntCol����)) > 0 Then
                CheckRedo = CheckRedo & .TextMatrix(i, 0) & ","
            End If
        Next
    End With
End Function

'Private Function GetRs(ByVal strҩƷid As String, ByVal rsTemp As ADODB.Recordset) As ADODB.Recordset
'    '���ܣ������Ƿ����ظ���¼�������ظ��ļ�¼���˵�
'    '��ͬʱѡ���˶����¼ �����ж����¼��֮��ͬʱ����ֻ��ʾһ��
'
'    Dim strTemp As String
'    Dim i As Integer
'
'    If strҩƷid <> "" Then
'        strTemp = ""
'        For i = 0 To UBound(Split(strҩƷid, ",")) - 1
'            strTemp = strTemp & "ҩƷid<>" & Split(strҩƷid, ",")(i) & " and "
'        Next
'
'        If strTemp <> "" Then
'            strTemp = Mid(strTemp, 1, Len(strTemp) - 4)
'        End If
'        rsTemp.Filter = strTemp
'    End If
'    If strҩƷid <> "" And mbln��ʾ = False Then
'        MsgBox "�Բ������и�ҩƷ���ҩƷ����ͬ���Σ��ظ���¼������ӣ�", vbInformation, gstrSysName
'        mbln��ʾ = True
'    End If
'    Set GetRs = rsTemp
'End Function

Private Function CheckData(ByVal rsTemp As ADODB.Recordset) As ADODB.Recordset
    '���ܣ���������б�������ҩƷ����ѡ���ҩƷ�Ƿ��ظ���ʱ��ҩƷ�Ƿ��п��

    Dim i As Integer
    Dim strTemp As String
    Dim str���� As String
    Dim strInfo As String
    Dim rsPrice As ADODB.Recordset
    Dim str��� As String
    Dim strSQL As String
    Dim strDub As String    '�ظ�ҩƷ
    Dim strNotNum As String  '�޿��ҩƷ
    Dim str�ظ�ҩ�� As String   '������¼�ظ�ѡ���˵�ҩƷ����
    Dim strNotҩ�� As String    '������¼��ЩҩƷ��ʱ�۵��޿��
    Dim rsRe As ADODB.Recordset
    
    On Error GoTo errHandle
    rsTemp.MoveFirst
    str���� = ""
    strTemp = ""
    Do While Not rsTemp.EOF
        If gtype_UserSysParms.P174_ҩƷ�ƿ���ȷ���� = 0 Then
            str���� = "0"
        Else
            str���� = IIf(IsNull(rsTemp!����), "0", rsTemp!����)
        End If
        If InStr(1, strTemp, rsTemp!ҩƷid & "," & str����) = 0 Then
            strTemp = strTemp & rsTemp!ҩƷid & "," & str���� & "," & rsTemp!ͨ���� & "|"
        End If
        
        If rsTemp!ʱ�� = 1 Then 'ʱ���޿���ҩƷ
            If Not mbln���쵥 Or (mbln���쵥 And mbln��ȷ����) Then
                gstrSQL = "Select Decode(Nvl(����, 0), 0, ʵ�ʽ�� / ʵ������, Nvl(���ۼ�, ʵ�ʽ�� / ʵ������))*" & Choose(mintUnit, 1, rsTemp!�����װ, rsTemp!סԺ��װ, rsTemp!ҩ���װ) & " as  �ۼ� " _
                    & "  from ҩƷ��� " _
                    & " where �ⷿid=[1] " _
                    & " and ҩƷid=[2] " _
                    & " and ����=1 and ʵ������>0 and " _
                    & " nvl(����,0)=[3] "
                Set rsRe = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, cboStock.ItemData(cboStock.ListIndex), rsTemp!ҩƷid, IIf(IsNull(rsTemp!����), 0, rsTemp!����))
                
                If rsRe.EOF Then
                    str��� = str��� & rsTemp!ҩƷid & "," & rsTemp!ͨ���� & "|"
                End If
            End If
        End If
        rsTemp.MoveNext
    Loop
        
    With mshBill    '���ظ��Ĳ�ѯ����
        For i = 1 To .rows - 2
            If InStr(1, strTemp, .TextMatrix(i, 0) & "," & .TextMatrix(i, mconIntCol����)) > 0 Then
                strInfo = strInfo & .TextMatrix(i, 0) & "," & .TextMatrix(i, mconIntColҩ��) & "|"
            End If
        Next
        
        If strInfo <> "" Then   'Ϊ��������ƴ��sql
            strDub = ""
            For i = 0 To UBound(Split(strInfo, "|")) - 1
                strDub = strDub & "ҩƷid<>" & Split(Split(strInfo, "|")(i), ",")(0) & " and "
                If UBound(Split(str�ظ�ҩ��, ",")) <= 2 Then
                    str�ظ�ҩ�� = str�ظ�ҩ�� & Split(Split(strInfo, "|")(i), ",")(1) & ","
                End If
            Next
            If strDub <> "" Then
                strDub = Mid(strDub, 1, Len(strDub) - 4)
            End If
        End If
        If str��� <> "" Then
            strNotNum = ""
            For i = 0 To UBound(Split(str���, "|")) - 1
                strNotNum = strNotNum & "ҩƷid<>" & Split(Split(str���, "|")(i), ",")(0) & " and "
                If UBound(Split(strNotҩ��, ",")) <= 2 Then
                    strNotҩ�� = strNotҩ�� & Split(Split(str���, "|")(i), ",")(1) & ","
                End If
            Next
            If strNotNum <> "" Then
                strNotNum = Mid(strNotNum, 1, Len(strNotNum) - 4)
            End If
        End If
        '�ж���ʲô��ʽƴ��sql
        
        If str�ظ�ҩ�� <> "" And strNotҩ�� <> "" Then
            MsgBox str�ظ�ҩ�� & "�б����Ѿ������ˣ�" & vbCrLf & strNotҩ�� & "��ʱ��ҩƷ��û�п�治������⣡" & vbCrLf & "����ҩƷ������ӣ�", vbInformation, gstrSysName
            strSQL = strDub & " and " & strNotNum
        End If
        If str�ظ�ҩ�� <> "" And strNotҩ�� = "" Then
            MsgBox str�ظ�ҩ�� & "�б����Ѿ������ˣ�" & vbCrLf & "����ҩƷ������ӣ�", vbInformation, gstrSysName
            strSQL = strDub
        End If
        If str�ظ�ҩ�� = "" And strNotҩ�� <> "" Then
            MsgBox strNotҩ�� & "��ʱ��ҩƷ��û�п�治������⣡" & vbCrLf & "����ҩƷ������ӣ�", vbInformation, gstrSysName
            strSQL = strNotNum
        End If
        If strSQL <> "" Then
            rsTemp.Filter = strSQL
        End If
        
        Set CheckData = rsTemp
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ���۸�() As Boolean
    '���ܣ�����ʱ���ж�ҩƷ�Ƿ������¼۸񣬲������޸ĺ���ʾ
    Dim strMsg As String '������ʾ��Ϣ
    Dim i As Integer, intSum As Integer, intPriceDigit As Integer
    Dim rsPrice As New ADODB.Recordset
    Dim Dbl���� As Double, dbl�ɱ��� As Double, dbl�ɱ���� As Double, dbl���ۼ� As Double, dbl���۽�� As Double, dbl��� As Double
    Dim bln�Ƿ�ʱ�� As Boolean
    
    On Error GoTo errHandle
    
    ���۸� = False
    
    With mshBill
        For i = 1 To .rows - 1
            If mshBill.TextMatrix(i, 0) <> "" Then
            
                bln�Ƿ�ʱ�� = Val(Split(.TextMatrix(i, mconIntCol���Ч��), "||")(1)) = 1
                Dbl���� = Val(.TextMatrix(i, mconIntColʵ������))
                
                '���ɱ���
                dbl�ɱ��� = zlStr.FormatEx(Get�ɱ���(Val(.TextMatrix(i, 0)), cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(i, mconIntCol����))) * Val(.TextMatrix(i, mconIntCol����ϵ��)), mintCostDigit)
                If .TextMatrix(i, mconIntCol�ɹ���) <> dbl�ɱ��� Then
                    intSum = intSum + 1
                    .TextMatrix(i, mconIntCol�ɹ���) = zlStr.FormatEx(dbl�ɱ���, mintCostDigit, , True)
                    .TextMatrix(i, mconIntCol�ɹ����) = zlStr.FormatEx(.TextMatrix(i, mconIntCol�ɹ���) * Dbl����, mintMoneyDigit, , True)
                End If
                
                '����ۼ�
                dbl���ۼ� = zlStr.FormatEx(Get�ۼ�(bln�Ƿ�ʱ��, Val(.TextMatrix(i, 0)), cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(i, mconIntCol����))) * Val(.TextMatrix(i, mconIntCol����ϵ��)), mintPriceDigit)
                If .TextMatrix(i, mconIntCol�ۼ�) <> dbl���ۼ� Then
                    intSum = intSum + 1
                    .TextMatrix(i, mconIntCol�ۼ�) = zlStr.FormatEx(dbl���ۼ�, mintPriceDigit, , True)
                    .TextMatrix(i, mconIntCol�ۼ۽��) = zlStr.FormatEx(.TextMatrix(i, mconIntCol�ۼ�) * Dbl����, mintMoneyDigit, , True)
                End If
                
                .TextMatrix(i, mconintCol���) = zlStr.FormatEx(Val(.TextMatrix(i, mconIntCol�ۼ۽��)) - Val(.TextMatrix(i, mconIntCol�ɹ����)), mintMoneyDigit, , True)
                
            End If
        Next
        
        If intSum > 0 Then
            MsgBox "�м�¼δʹ�����¼۸񣬳������Զ���ɸ��£��ɱ��ۡ��ɱ����ۼۡ��ۼ۽���ۣ������º����飡", vbInformation, gstrSysName
            ���۸� = True
        End If
    End With
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
