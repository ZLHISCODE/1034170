VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmRequestDrugCard 
   Caption         =   "ҩƷ���쵥"
   ClientHeight    =   6345
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10785
   Icon            =   "frmRequestDrugCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6345
   ScaleWidth      =   10785
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdȫ������ 
      Caption         =   "ȫ������"
      Height          =   350
      Left            =   9360
      TabIndex        =   31
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdȫ�� 
      Caption         =   "ȫ�����"
      Height          =   350
      Left            =   8040
      TabIndex        =   30
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CheckBox chkExportPlan 
      Caption         =   "����ʱֻͬ�������ǳ���ҩƷ�ļƻ�����"
      Height          =   380
      Left            =   5160
      TabIndex        =   29
      Top             =   5160
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdAllCls 
      Caption         =   "ȫ��(&L)"
      Height          =   350
      Left            =   6480
      TabIndex        =   28
      Top             =   5640
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.CommandButton cmdAllSel 
      Caption         =   "ȫ��(&A)"
      Height          =   350
      Left            =   5160
      TabIndex        =   27
      Top             =   5640
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.TextBox txtCode 
      Height          =   300
      Left            =   3240
      TabIndex        =   9
      Top             =   5137
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "����(&F)"
      Height          =   350
      Left            =   1560
      TabIndex        =   8
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   7
      Top             =   5040
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   8040
      TabIndex        =   5
      Top             =   5520
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   9360
      TabIndex        =   6
      Top             =   5520
      Width           =   1100
   End
   Begin VB.PictureBox Pic���� 
      BackColor       =   &H80000004&
      Height          =   4965
      Left            =   0
      ScaleHeight     =   4905
      ScaleWidth      =   11655
      TabIndex        =   10
      Top             =   0
      Width           =   11715
      Begin VB.TextBox txtժҪ 
         Height          =   300
         Left            =   900
         MaxLength       =   40
         TabIndex        =   4
         Top             =   4080
         Width           =   10410
      End
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   557
         Width           =   1515
      End
      Begin ZL9BillEdit.BillEdit mshBill 
         Height          =   2805
         Left            =   195
         TabIndex        =   2
         Top             =   950
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
      Begin VB.Label lblDifference 
         AutoSize        =   -1  'True
         Caption         =   "��ۺϼ�:"
         Height          =   180
         Left            =   4920
         TabIndex        =   25
         Top             =   3840
         Width           =   810
      End
      Begin VB.Label lblSalePrice 
         AutoSize        =   -1  'True
         Caption         =   "�ۼ۽��ϼ�:"
         Height          =   180
         Left            =   2040
         TabIndex        =   24
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label lblPurchasePrice 
         AutoSize        =   -1  'True
         Caption         =   "�ɱ����ϼ�:"
         Height          =   180
         Left            =   240
         TabIndex        =   23
         Top             =   3840
         Width           =   1170
      End
      Begin VB.Label Txt����� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   7950
         TabIndex        =   21
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label Txt������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   10050
         TabIndex        =   20
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt�������� 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   2940
         TabIndex        =   19
         Top             =   4440
         Width           =   1875
      End
      Begin VB.Label Txt������ 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   900
         TabIndex        =   18
         Top             =   4440
         Width           =   915
      End
      Begin VB.Label txtNo 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   9960
         TabIndex        =   17
         Top             =   550
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
         TabIndex        =   16
         Top             =   587
         Width           =   480
      End
      Begin VB.Label lblժҪ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ժҪ(&M)"
         Height          =   180
         Left            =   240
         TabIndex        =   3
         Top             =   4155
         Width           =   650
      End
      Begin VB.Label LblTitle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ҩƷ���쵥"
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
         TabIndex        =   15
         Top             =   120
         Width           =   11535
      End
      Begin VB.Label LblStock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ҩ�ⷿ(&S)"
         Height          =   180
         Left            =   240
         TabIndex        =   0
         Top             =   617
         Width           =   990
      End
      Begin VB.Label Lbl������ 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   300
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   4500
         Width           =   720
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
            Picture         =   "frmRequestDrugCard.frx":014A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":0364
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":057E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":0798
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":09B2
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":0BCC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":0DE6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":1000
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
            Picture         =   "frmRequestDrugCard.frx":121A
            Key             =   "PreView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":1434
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":164E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":1868
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":1A82
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":1C9C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":1EB6
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestDrugCard.frx":20D0
            Key             =   "Find"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar staThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   26
      Top             =   5985
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmRequestDrugCard.frx":22EA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12674
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmRequestDrugCard.frx":2B7E
            Key             =   "PY"
            Object.ToolTipText     =   "ƴ��(F7)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   617
            MinWidth        =   617
            Picture         =   "frmRequestDrugCard.frx":3080
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
      Left            =   2760
      TabIndex        =   22
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
End
Attribute VB_Name = "frmRequestDrugCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint�༭״̬ As Integer             '1.������2���޸ģ�3�����գ�4���鿴��5��ͨ����������6�����ܣ����պ��¼���յǼ��ˣ�����ȡ������Ľ��գ���7������
Private mstr���ݺ� As String                '����ĵ��ݺ�;
Private mint��¼״̬ As Integer             '1:������¼;2-������¼;3-�Ѿ�������ԭ��¼
Private mblnSuccess As Boolean              'ֻҪ��һ�ųɹ�����ΪTrue������ΪFalse
Private mblnFirst As Boolean
Private mblnSave As Boolean                 '�Ƿ���̺����   TURE���ɹ���
Private mfrmMain As Form
Private mintcboIndex As Integer
Private mblnEdit As Boolean                 '�Ƿ�����޸�
Private mblnChange As Boolean               '�Ƿ���й��༭
Private mbln����״̬ As Boolean
Private mstr�ⷿ As String                  '��¼�Ѿ�����˵Ŀⷿ

Private mint��ȷ���� As Integer             '��ʾ����д���쵥ʱ���Ƿ���ȷҩƷ������
Private mint����� As Integer             '��ʾҩƷ����ʱ�Ƿ���п���飺0-�����;1-��飬�������ѣ�2-��飬�����ֹ
Private mint��������ⷿ As Integer     '�����ڳ���ʱ��ԭ���ⷿ�Ƿ���п���飺0-�����;1-��飬�������ѣ�2-��飬�����ֹ
Private mcolUsedCount As Collection         '��ʹ�õ���������
Private mstrPrivs As String                     'Ȩ��
Private mlngStockID As Long                 '��ǰ�û���ѡ��ҩ��ID
Private mintApplyType As Integer            '���췽ʽ��0-�ֹ�����;1-����������;2-��������;3-��������;4-����������;5-�������쵥δ����;6-������������;7-������������
Private mstrEndTime As String               '���Զ����췽ʽΪ7ʱ������ʱ�䷶Χ�еĽ���ʱ��
Private rsDepend As New ADODB.Recordset

Private mintParallelRecord As Integer       '���������󵥾ݲ���ִ�еĴ��� 1���������������2���Ѿ�ɾ���ļ�¼��3���Ѿ���˵ļ�¼
Private mstrTime_Start As String                        '���뵥�ݱ༭����ʱ�����༭���ݵ�����޸�ʱ��
Private mstrTime_End As String                        '�˿̸ñ༭���ݵ�����޸�ʱ��

Private mintUnit As Integer                 '��λϵ����1-�ۼ�;2-����;3-סԺ;4-ҩ��
Private mintDrugNameShow As Integer         'ҩƷ��ʾ��0����ʾ��������ƣ�1������ʾ���룻2������ʾ����
Private mblnUpdate As Boolean               '������¼������˺��Ƿ�������¼۸�
Private mblnViewCost As Boolean             '�鿴�ɱ��� true-���Բ鿴�ɱ��� false-�����Բ鿴�ɱ���
Private Const MStrCaption As String = "ҩƷ�������"

Private mobjPlugIn As Object '��Ҳ���

'�Ӳ�������ȡҩƷ�۸����������С��λ�������㾫�ȣ�
Private mintCostDigit As Integer            '�ɱ���С��λ��
Private mintPriceDigit As Integer           '�ۼ�С��λ��
Private mintNumberDigit As Integer          '����С��λ��
Private mintMoneyDigit As Integer           '���С��λ��

Private Const mconint�ۼ۵�λ As Integer = 1
Private Const mconint���ﵥλ As Integer = 2
Private Const mconintסԺ��λ As Integer = 3
Private Const mconintҩ�ⵥλ As Integer = 4


Private mbln�¿������� As Boolean           '�Ƿ��·�ҩҩ���Ŀ�������
Private recSort As ADODB.Recordset          '��ҩƷID�����ר�ü�¼��

Private mint����ʽ As Integer             '����ʱ��0������������1�������������뵥��

Private mbln����� As Boolean

'=========================================================================================
Private Const mconIntCol�к� As Integer = 1
Private Const mconIntColҩ�� As Integer = 2
Private Const mconIntCol��Ʒ�� As Integer = 3
Private Const mconIntCol��Դ As Integer = 4
Private Const mconIntCol����ҩ�� As Integer = 5
Private Const mconIntCol��� As Integer = 6
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
Private Const mconintcol��ǰ��� As Integer = 22
Private Const mconintcol�Է���� As Integer = 23
Private Const mconIntCol�������� As Integer = 24
Private Const mconIntCol��д���� As Integer = 25
Private Const mconIntColʵ������ As Integer = 26
Private Const mconIntCol�ɹ��� As Integer = 27
Private Const mconIntCol�ɹ���� As Integer = 28
Private Const mconIntCol�ۼ� As Integer = 29
Private Const mconIntCol�ۼ۽�� As Integer = 30
Private Const mconintCol��� As Integer = 31
Private Const mconIntCol�ϴι�Ӧ��ID As Integer = 32
Private Const mconintCol��ʵ���� As Integer = 33
Private Const mconIntColҩƷ��������� As Integer = 34
Private Const mconIntColҩƷ���� As Integer = 35
Private Const mconIntColҩƷ���� As Integer = 36
Private Const mconIntCol����ҩƷ As Integer = 37
Private Const mconIntColԭʼ���� As Integer = 38
Private Const mconIntColS  As Integer = 39             '������
'=========================================================================================

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
    gstrSQL = " Select �������,��ҩ����,��ҩ�� From ҩƷ�շ���¼ " & _
            " Where ����=6 And NO=[1] And ��¼״̬=1 And RowNum=1 "
    Set rs = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��鵥��]", strNo)
    
    With rs
        '���ؿգ���ʾ�Ѿ�ɾ��
        If .EOF Then
            CheckBill = "�õ����Ѿ�����������Աɾ����"
        ElseIf Not IsNull(!�������) Then
            CheckBill = "�õ����Ѿ�����������Ա��ˣ�"
        ElseIf Not IsNull(!��ҩ����) Then
            CheckBill = "�õ����Ѿ�����������Ա���ͣ�"
        ElseIf Not IsNull(!��ҩ��) Then
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
'�������������
Private Function GetDepend() As Boolean
    Dim strMsg As String

    GetDepend = False
    On Error GoTo ErrHand

    '���ҩƷ�������Ƿ�����
    strMsg = "û������ҩƷ�ƿ����⼰�����������ҩƷ������࣡"
    gstrSQL = "SELECT B.Id,B.ϵ�� " _
            & "FROM ҩƷ�������� A, ҩƷ������ B " _
            & "Where A.���id = B.ID AND A.���� = 6"
    Set rsDepend = zlDataBase.OpenSQLRecord(gstrSQL, "ҩƷ�ƿ����")

    With rsDepend
        If .RecordCount = 0 Then GoTo ErrHand
        .Filter = "ϵ��=1"
        If .RecordCount = 0 Then
            .Filter = 0
            strMsg = "û������ҩƷ�ƿ������������ҩƷ������࣡"
            GoTo ErrHand
        End If
        .Filter = "ϵ��=-1"
        If .RecordCount = 0 Then
            .Filter = 0
            strMsg = "û������ҩƷ�ƿ�ĳ����������ҩƷ������࣡"
            GoTo ErrHand
        End If
        .Filter = 0
        
        'gstrSQL = ReturnSQL(mlngStockID, False)
    End With
    'Set rsDepend = zldatabase.OpenSQLRecord(gstrSQL, "ҩƷ�������", mlngStockID)
    Set rsDepend = ReturnSQL(mlngStockID, "ҩƷ�������", False, 1343)

    strMsg = "û���κοⷿ�������죬����[������������]��ҩƷ���������ã�"
    If rsDepend.RecordCount = 0 Then
        MsgBox strMsg, vbInformation, gstrSysName
        rsDepend.Close
        Exit Function
    End If
    
    GetDepend = True
    Exit Function
ErrHand:
    MsgBox strMsg, vbInformation, gstrSysName
    rsDepend.Close
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub ShowCard(FrmMain As Form, ByVal str���ݺ� As String, ByVal int�༭״̬ As Integer, Optional int��¼״̬ As Integer = 1, Optional BlnSuccess As Boolean = False, Optional lngStockid As Long = 0, Optional int����ʽ As Integer = 0, Optional intApplyType As Integer = 0)
    Dim strSQL As String
    Dim rsPara As New ADODB.Recordset
    
    mblnSave = False
    mblnSuccess = False
    mstr���ݺ� = str���ݺ�
    mint�༭״̬ = int�༭״̬
    mint��¼״̬ = int��¼״̬
    mint����ʽ = int����ʽ
    mintApplyType = intApplyType
    mblnSuccess = BlnSuccess
    mblnChange = False
    mintParallelRecord = 1
    mstrPrivs = GetPrivFunc(glngSys, 1343)
    mlngStockID = IIf(lngStockid = 0, glngDeptId, lngStockid)
    
    Set mfrmMain = FrmMain
    If Not GetDepend Then Exit Sub
    mint����� = MediWork_GetCheckStockRule(cboStock.ItemData(cboStock.ListIndex))
    mint��������ⷿ = MediWork_GetCheckStockRule(mlngStockID)
    
    mint��ȷ���� = gtype_UserSysParms.P73_��ȷ����ҩƷ����
    
    If mint�༭״̬ <> 5 Then
        Me.cmdȫ������.Visible = False
        Me.cmdȫ��.Visible = False
    End If
    
    If mint��ȷ���� = 0 Then
        mint����� = 0
    Else
        mbln����� = True
    End If
    
    mbln�¿������� = (gtype_UserSysParms.P96_ҩƷ��¿��ÿ�� = 1)
    
    mblnEdit = False
         
    If mint�༭״̬ = 5 Then
        Me.Height = Me.Height + Me.cmdȫ��.Height
    End If
         
    If mint�༭״̬ = 1 Or mint�༭״̬ = 5 Then
        mblnEdit = True
        mblnFirst = True
        
        chkExportPlan.Visible = True
    
    ElseIf mint�༭״̬ = 2 Then
        mblnEdit = True
        mblnFirst = True
    ElseIf mint�༭״̬ = 4 Then
        mblnFirst = True
        mblnEdit = False
        CmdSave.Caption = "��ӡ(&P)"
        If Not IsHavePrivs(mstrPrivs, "���ݴ�ӡ") Then
            CmdSave.Visible = False
        Else
            CmdSave.Visible = True
        End If
    ElseIf mint�༭״̬ = 7 Then
        mblnEdit = False
        mblnFirst = True
        cmdAllSel.Visible = True
        cmdAllCls.Visible = True
        
        If mint����ʽ = 1 Then
            CmdSave.Caption = "�������(&O)"
            CmdSave.Width = CmdSave.Width + 200
        Else
            CmdSave.Caption = "����(&O)"
            CmdSave.Width = CmdCancel.Width
        End If
    End If
    
    LblTitle.Caption = GetUnitName & LblTitle.Caption
    Me.Show vbModal, FrmMain
    BlnSuccess = mblnSuccess
    str���ݺ� = mstr���ݺ�
    
End Sub

Private Sub cboStock_Click()
    If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Or mint�༭״̬ = 5 Then
        Call SetSelectorRS(2, "ҩƷ�������", cboStock.ItemData(cboStock.ListIndex), mlngStockID, mlngStockID, , , , , , mbln�����)
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

Private Sub cboStock_Change()
    mblnChange = True
End Sub

Private Sub cboStock_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboStock_Validate False
        zlCommFun.PressKey (vbKeyTab)
    End If
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
                If MsgBox("����ı�ⷿ���п���Ҫ�ı���ӦҩƷ�ĵ�λ����Ҫ������е������ݣ����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    '����ҩƷ��λ�ı�
                    mintcboIndex = .ListIndex
                    mshBill.ClearBill
                    
                    If Me.mshBill.ColWidth(mconIntCol��������) > 0 Then
                        Me.mshBill.ColWidth(mconIntCol��������) = 0
                        Me.cmdȫ������.Visible = False
                        Me.cmdȫ��.Visible = False
                        Call Form_Resize
                    End If
                    
                Else
                    .ListIndex = mintcboIndex
                End If
            Else
                mintcboIndex = .ListIndex
            End If
        End If
        mint����� = MediWork_GetCheckStockRule(cboStock.ItemData(cboStock.ListIndex))
        If mint��ȷ���� = 0 Then
            mint����� = 0
        Else
            mbln����� = True
        End If
        
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'����
Private Sub cmdFind_Click()
    
    If lblCode.Visible = False Then
        lblCode.Visible = True
        txtCode.Visible = True
        txtCode.SetFocus
    Else
        FindRow mshBill, mconIntColҩ��, txtCode.Text, True
        lblCode.Visible = False
        txtCode.Visible = False
    End If
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub cmdȫ������_Click()
    Dim Row As Integer
    Dim count As Integer
    
    For Row = 1 To Me.mshBill.rows - 1
        If Val(Me.mshBill.TextMatrix(Row, 0)) <> 0 Then
            count = count + 1
            Exit For
        End If
    Next
    
    If count = 0 Then
        Exit Sub
    End If
    
    If MsgBox("��ȷ��Ҫ������������ֵ��Ϊ��д������ʵ��������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        For Row = 1 To Me.mshBill.rows - 2
            Me.mshBill.TextMatrix(Row, mconIntCol��д����) = Me.mshBill.TextMatrix(Row, mconIntCol��������)
            Me.mshBill.TextMatrix(Row, mconIntColʵ������) = Me.mshBill.TextMatrix(Row, mconIntCol��������)
            If Val(Me.mshBill.TextMatrix(Row, mconIntCol��д����)) <> 0 Then
                Call GetPrice(Row)
            Else
                With Me.mshBill
                    .TextMatrix(Row, mconIntCol�ۼ۽��) = 0
                    .TextMatrix(Row, mconintCol���) = 0
                    .TextMatrix(Row, mconIntCol�ɹ���) = 0
                    .TextMatrix(Row, mconIntCol�ɹ����) = 0
                End With
            End If
        Next
        Call ��ʾ�ϼƽ��
    End If
End Sub

Private Sub cmdȫ��_Click()
    Dim Row As Integer
    Dim count As Integer
    
    For Row = 1 To Me.mshBill.rows - 1
        If Val(Me.mshBill.TextMatrix(Row, 0)) <> 0 Then
            count = count + 1
            Exit For
        End If
    Next
    
    If count = 0 Then
        Exit Sub
    End If
    
    If MsgBox("��ȷ��Ҫ����д������ʵ��������Ϊ0��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
        For Row = 1 To Me.mshBill.rows - 2
            Me.mshBill.TextMatrix(Row, mconIntCol��д����) = 0
            Me.mshBill.TextMatrix(Row, mconIntColʵ������) = 0
            With Me.mshBill
                .TextMatrix(Row, mconIntCol�ۼ۽��) = 0
                .TextMatrix(Row, mconintCol���) = 0
                .TextMatrix(Row, mconIntCol�ɹ���) = 0
                .TextMatrix(Row, mconIntCol�ɹ����) = 0
            End With
        Next
        Call ��ʾ�ϼƽ��
    End If
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then
        If mshBill.rows > 50 Then
            Call AviShow(Me) '��ʾ�û����ڲ�ѯ����
        End If
        Call get�������    'Ϊ��ǰ��������ͶԷ���������и�ֵ
        If mshBill.rows > 50 Then
            Call AviShow(Me, False)
        End If
        Exit Sub
    End If
    
    mblnFirst = False
    If mint�༭״̬ = 5 Then
        If Not frmRequestNavigation.ShowNavigation(Me, mlngStockID, mintApplyType, mstrEndTime, mbln����״̬) = True Then
            Unload Me
            Exit Sub
        End If
        mshBill.SetFocus
        If mintApplyType = 7 And Not IsHavePrivs(mstrPrivs, "�Զ�����ʱ�޸�ҩƷ����") Then
            mshBill.Active = False
        End If
    End If
    If mbln����״̬ = True Then
        Call Form_Resize
    End If
    
    mblnChange = False
    Select Case mintParallelRecord
        Case 1
            '����
        Case 2
            If mint�༭״̬ = 7 Then
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


Private Sub CmdSave_Click()
    Dim BlnSuccess As Boolean
    Dim Row As Integer
    Dim count As Integer
    Dim intRows As Integer
    
    '�����������ݼ�
    Call SetSortRecord
        
    If mint�༭״̬ = 4 Then    '�鿴
        '��ӡ
        printbill
        '�˳�
        Unload Me
        Exit Sub
    End If
    
    If Me.mshBill.TextMatrix(Me.mshBill.rows - 1, 0) <> "" Then
        intRows = Me.mshBill.rows - 1
    Else
        intRows = Me.mshBill.rows - 2
    End If
    
    For Row = 1 To intRows
        If Val(Me.mshBill.TextMatrix(Row, mconIntCol��д����)) = 0 Then
            count = count + 1
            If count = intRows Then
                MsgBox "�����쵥�ϵ�����ҩƷ����д������Ϊ0�����ܼ���������", vbInformation + vbOKOnly, gstrSysName
                Exit Sub
            End If
        End If
    Next

    For Row = 1 To Me.mshBill.rows - 2
        If NVL(Me.mshBill.TextMatrix(Row, mconIntCol��д����), 0) = 0 Then
            If MsgBox("�����쵥������д����Ϊ0��ҩƷ��" & vbCrLf & "��д����Ϊ0��ҩƷ�����ܱ���Ϊ���쵥��" & vbCrLf & "�Ƿ����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
            Exit For
        End If
    Next
    
    If mint�༭״̬ = 6 Then       '���
        If Not ҩƷ�������(Txt������.Caption) Then Exit Sub
        
        '��������¿�������Ϊ�٣���Ҫ�����ÿ���Ƿ��㹻
        If mbln�¿������� = False Then
            If Not CheckStock Then Exit Sub
        End If
        
        If SaveCheck() = True Then
            If Val(zlDataBase.GetPara("��˴�ӡ", glngSys, 1343)) = 1 Then
                '��ӡ
                If IsHavePrivs(mstrPrivs, "���ݴ�ӡ") Then
                    printbill
                End If
            End If
            Unload Me
        End If
        Exit Sub
    End If
    
    If mint�༭״̬ = 7 Then '����
        If mblnChange = False Then
            MsgBox "��¼�����������", vbInformation + vbOKOnly, gstrSysName
            Exit Sub
        End If
        
        If MsgBox("��ȷʵҪ����������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            If SaveStrike = True Then
                Unload Me
            End If
        End If
        Exit Sub
    End If
    
    If mint�༭״̬ = 2 And mblnUpdate = False Then
        If Not ��鵥��(6, txtNo.Tag, True, True) Then
            MsgBox "�м�¼δʹ�����¼۸񣬳����Զ���ɸ��£��ۼۡ��ɱ��ۡ��ۼ۽��ɱ�����ۣ������º����飡", vbInformation, gstrSysName
            Call RefreshBill
            mblnUpdate = True
            Exit Sub
        End If
    End If
        
    If ValidData = False Then Exit Sub
    
    If mint�༭״̬ = 1 Or mint�༭״̬ = 5 Then '��������ʱ���жϼ۸��Ƿ��Ѿ�����
        If ���۸� Then
            mblnUpdate = True
            mblnChange = True
            Exit Sub
        End If
    End If

    BlnSuccess = SaveCard
        
    If BlnSuccess = True Then
            
        If Val(zlDataBase.GetPara("���̴�ӡ", glngSys, 1343)) = 1 Then
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
    cboStock.SetFocus
    mblnChange = False

    If txtNo.Tag <> "" Then Me.staThis.Panels(2).Text = "��һ�ŵ��ݵ�NO�ţ�" & txtNo.Tag
End Sub

Private Sub Form_Load()
    Dim strStock As String
    Dim rsStock As New Recordset
    Dim intStock As Integer
    
    txtNo = mstr���ݺ�
    txtNo.Tag = txtNo
    mblnViewCost = IsHavePrivs(mstrPrivs, "�鿴�ɱ���")
    chkExportPlan.Value = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ�������", "ͬ�����ɼƻ���", 0))
    
    mintDrugNameShow = Int(Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ�������", "ҩƷ������ʾ��ʽ", 0)))
    If mintDrugNameShow > 2 Or mintDrugNameShow < 0 Then mintDrugNameShow = 0
    mnuColDrug.Item(mintDrugNameShow).Checked = True
    
    intStock = -1
    With cboStock
        .Clear
        mstr�ⷿ = ""
        Do While Not rsDepend.EOF
            If InStr(1, mstr�ⷿ, "|" & rsDepend!Id & "|") = 0 Then
                .AddItem rsDepend!����
                .ItemData(.NewIndex) = rsDepend!Id
                mstr�ⷿ = mstr�ⷿ & "|" & rsDepend!Id & "|"
                
                If rsDepend!ҩ������ = 1 And intStock = -1 Then
                    intStock = .NewIndex
                End If
            End If
            
            rsDepend.MoveNext
        Loop
        .ListIndex = IIf(intStock = -1, 0, intStock)
    End With
    
    If mlngStockID = 0 Then
        mlngStockID = mfrmMain.cboStock.ItemData(Me.cboStock.ListIndex)
    End If
    Call GetDrugDigit(mlngStockID, MStrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    
    Call initCard
    
    mstrTime_Start = GetBillInfo(6, mstr���ݺ�)
    RestoreWinState Me, App.ProductName, MStrCaption
    
    
    '����ϵͳ��������ҩ����Ա�鿴����ʱ���Ƿ���ʾ�ɱ���
    mshBill.ColWidth(mconIntCol�ɹ���) = IIf(mblnViewCost, 1000, 0)
    mshBill.ColWidth(mconIntCol�ɹ����) = IIf(mblnViewCost, 900, 0)
    mshBill.ColWidth(mconintCol���) = IIf(mblnViewCost, 900, 0)
    mshBill.ColWidth(mconIntCol�ϴι�Ӧ��ID) = 0
    mshBill.ColWidth(mconintCol��ʵ����) = 0
    mshBill.ColWidth(mconIntCol��������) = 0
    
    If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Or mint�༭״̬ = 5 Then
        mshBill.ColWidth(mconintcol��ǰ���) = 1100
        mshBill.ColWidth(mconintcol�Է����) = 1100
    Else
        mshBill.ColWidth(mconintcol��ǰ���) = 0
        mshBill.ColWidth(mconintcol�Է����) = 0
    End If
    
    '��Ʒ���д���
    If gintҩƷ������ʾ = 2 Then
        '��ʾ��Ʒ����
        mshBill.ColWidth(mconIntCol��Ʒ��) = IIf(mshBill.ColWidth(mconIntCol��Ʒ��) = 0, 2000, mshBill.ColWidth(mconIntCol��Ʒ��))
    Else
        '��������ʾ��Ʒ����
        mshBill.ColWidth(mconIntCol��Ʒ��) = 0
    End If
    
    mbln����� = (Val(zlDataBase.GetPara("��ʾ�޿��ҩƷ", glngSys, 1343, 0)) = 0)
End Sub

Private Sub initCard()
    Dim i As Integer
    Dim lngStockid As Long
    Dim rsInitCard As New Recordset
    Dim strUnitQuantity As String
    Dim strUnitQuantity_Stock As String
    Dim intRow As Integer
    Dim vardrug As Variant
    Dim numUseAbleCount As Double
    Dim dateCurDate As Date
    Dim strOrder As String, strCompare As String
    Dim IntCount As Integer
    Dim intCostDigit As Integer        '�ɱ���С��λ��
    Dim intPricedigit As Integer       '�ۼ�С��λ��
    Dim intNumberDigit As Integer      '����С��λ��
    Dim intMoneyDigit As Integer       '���С��λ��
    Dim strҩ�� As String
    Dim strSqlOrder As String
    
    On Error GoTo errHandle
    strOrder = zlDataBase.GetPara("����", glngSys, 1343)
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
    
    intCostDigit = mintCostDigit
    intPricedigit = mintPriceDigit
    intNumberDigit = mintNumberDigit
    intMoneyDigit = mintMoneyDigit
        
    If mint�༭״̬ = 4 Then
        With cboStock
            'ȡָ�����ݵĳ���ⷿ�����ⷿ
            gstrSQL = " Select �ⷿID,�Է�����ID From ҩƷ�շ���¼" & _
                      " Where NO=[1] And ����=6 And ���ϵ��=-1 And Rownum<2"
            Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[ȡָ�����ݵĳ���ⷿ�����ⷿ]", mstr���ݺ�)
            
            If rsInitCard.RecordCount <> 0 Then
                lngStockid = rsInitCard!�ⷿid
            End If
            
            For IntCount = 0 To .ListCount - 1
                If .ItemData(IntCount) = lngStockid Then
                    .ListIndex = IntCount: Exit For
                End If
            Next
        End With
    Else
        With cboStock
            If Not (mint�༭״̬ = 1 Or mint�༭״̬ = 5) Then
                'ȡָ�����ݵĳ���ⷿ�����ⷿ
                gstrSQL = " Select �ⷿID,�Է�����ID From ҩƷ�շ���¼" & _
                          " Where NO=[1] And ����=6 And ���ϵ��=-1 And Rownum<2"
                Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[ȡָ�����ݵĳ���ⷿ�����ⷿ]", mstr���ݺ�)
                
                If rsInitCard.RecordCount <> 0 Then
                    lngStockid = rsInitCard!�ⷿid
                End If
            End If
            For IntCount = 0 To .ListCount - 1
                If .ItemData(IntCount) = lngStockid Then
                    .ListIndex = IntCount: Exit For
                End If
            Next
            mintcboIndex = .ListIndex
        End With
    End If
    
    If mint�༭״̬ = 7 Then
       lngStockid = mlngStockID
    End If
    
    dateCurDate = zlDataBase.Currentdate()
    
    Select Case mint�༭״̬
        Case 1, 5
            Txt������ = gstrUserName
            Txt�������� = Format(dateCurDate, "yyyy-mm-dd hh:mm:ss")
            initGrid
        Case 2, 4, 6, 7
            initGrid
            
            Select Case mintUnit
                Case mconint�ۼ۵�λ
                    strUnitQuantity = "B.�ͻ���λ,B.�ͻ���װ,D.���㵥λ AS ��λ, A.��д����,a.ʵ������,a.�ɱ���,a.���ۼ�,'1' as ����ϵ��,"
                    strUnitQuantity_Stock = "Z.��������,Z.ʵ�ʽ��,Z.ʵ�ʲ��"
                Case mconint���ﵥλ
                    strUnitQuantity = "B.�ͻ���λ,B.�ͻ���װ,B.���ﵥλ AS ��λ,(A.��д���� / B.�����װ) AS ��д����,(A.ʵ������ / B.�����װ) AS ʵ������,a.�ɱ���*B.�����װ as �ɱ���,a.���ۼ�*B.�����װ as ���ۼ�,B.�����װ as ����ϵ��,"
                    strUnitQuantity_Stock = "Z.��������/B.�����װ As ��������,Z.ʵ�ʽ��,Z.ʵ�ʲ��"
                Case mconintסԺ��λ
                    strUnitQuantity = "B.�ͻ���λ,B.�ͻ���װ,B.סԺ��λ AS ��λ,(A.��д���� / B.סԺ��װ) AS ��д����,(A.ʵ������ / B.סԺ��װ) AS ʵ������,a.�ɱ���*B.סԺ��װ as �ɱ���,a.���ۼ�*B.סԺ��װ as ���ۼ�,B.סԺ��װ as ����ϵ��,"
                    strUnitQuantity_Stock = "Z.��������/B.סԺ��װ As ��������,Z.ʵ�ʽ��,Z.ʵ�ʲ��"
                Case mconintҩ�ⵥλ
                    strUnitQuantity = "B.�ͻ���λ,B.�ͻ���װ,B.ҩ�ⵥλ AS ��λ,(A.��д���� / B.ҩ���װ) AS ��д����,(A.ʵ������ / B.ҩ���װ) AS ʵ������,a.�ɱ���*B.ҩ���װ as �ɱ���,a.���ۼ�*B.ҩ���װ as ���ۼ�,B.ҩ���װ as ����ϵ��,"
                    strUnitQuantity_Stock = "Z.��������/B.ҩ���װ As ��������,Z.ʵ�ʽ��,Z.ʵ�ʲ��"
            End Select
            
            If mint�༭״̬ = 7 Then
                gstrSQL = "SELECT W.*,Z.��������/W.����ϵ�� AS  ��������,Z.ʵ�ʽ��,Z.ʵ�ʲ�� " & _
                    " FROM " & _
                    "     (SELECT DISTINCT A.ҩƷID,A.���,'[' || D.���� || ']' As ҩƷ����, D.���� As ͨ����, E.���� As ��Ʒ��," & _
                    "     B.ҩƷ��Դ,B.����ҩ��,D.���,D.���� AS ԭ����,A.����, A.����,A.����,B.ָ�������,B.ҩ����� AS ��������," & _
                    "     B.���Ч��,A.Ч��," & strUnitQuantity & _
                    "     A.�ɱ����,0 ���۽��, 0 ���,D.ժҪ,A.�ⷿID,A.�Է�����ID,D.�Ƿ���,B.ҩ������ AS ҩ����������,A.�ϴι�Ӧ��ID,A.��׼�ĺ�,A.��д���� ��ʵ���� " & _
                    "     FROM " & _
                    "         (SELECT MIN(ID) AS ID, SUM(ʵ������) AS ��д����,0 ʵ������,SUM(�ɱ����) AS �ɱ����," & _
                    "          ҩƷID,���,����, ����,Ч��,NVL(����,0) ����,����,�ɱ���,���ۼ�,�ⷿID,�Է�����ID,������ID,NVL(��ҩ��λID,0) �ϴι�Ӧ��ID,��׼�ĺ� " & _
                    "          FROM ҩƷ�շ���¼ X " & _
                    "          WHERE NO=[1] AND ����=6 AND ���ϵ��=-1 " & _
                    "          GROUP BY ҩƷID,���,����,����,Ч��,NVL(����,0),����,�ɱ���,���ۼ�,�ⷿID,�Է�����ID,������ID,NVL(��ҩ��λID,0),��׼�ĺ�" & _
                    "          HAVING SUM(ʵ������)<>0 ) A," & _
                    "     ҩƷ��� B,�շ���ĿĿ¼ D,�շ���Ŀ���� E, " & _
                    " (Select ���, ժҪ From ҩƷ�շ���¼ " & _
                    "  Where ���� = 6 And NO = [1] And ���ϵ�� = -1 And (��¼״̬ = 1 Or Mod(��¼״̬, 3) = 0)) D " & _
                    "     WHERE A.ҩƷID = B.ҩƷID AND B.ҩƷID=E.�շ�ϸĿID(+) AND E.����(+)=3 AND B.ҩƷID=D.ID And A.��� = D.���) W," & _
                    "     (SELECT  ҩƷID,NVL(����,0) ����,��������,ʵ�ʽ��,ʵ�ʲ�� " & _
                    "     FROM ҩƷ��� WHERE �ⷿID=[2] AND ����=1) Z " & _
                    " WHERE W.ҩƷID=Z.ҩƷID(+) AND NVL(W.����,0)=Z.����(+) " & _
                     " ORDER BY " & strSqlOrder
            Else
                gstrSQL = "SELECT DISTINCT A.ҩƷID,A.���,'[' || D.���� || ']' As ҩƷ����, D.���� As ͨ����, E.���� As ��Ʒ��," & _
                    " B.ҩƷ��Դ,B.����ҩ��,D.���,D.���� AS ԭ����,A.����,A.����,A.����,B.ָ�������,B.ҩ����� AS ��������,A.��д���� as ԭʼ����, " & _
                    " B.���Ч��,A.Ч��," & strUnitQuantity & _
                    " A.�ɱ����,A.���۽��, A.���, " & strUnitQuantity_Stock & _
                    " ,A.ժҪ,������,��������,�����,�������,A.�ⷿID,A.�Է�����ID,D.�Ƿ���,B.ҩ������ AS ҩ����������,NVL(A.��ҩ��λID,0) �ϴι�Ӧ��ID,A.��׼�ĺ�,nvl(A.����,0) As ���췽ʽ  " & _
                    " FROM ҩƷ�շ���¼ A, ҩƷ��� B,�շ���Ŀ���� E ,�շ���ĿĿ¼ D, " & _
                    "     (SELECT ҩƷID,NVL(����,0) ����,��������,ʵ�ʽ��,ʵ�ʲ�� " & _
                    "     FROM ҩƷ��� WHERE �ⷿID=[2] AND ����=1) Z " & _
                    " WHERE A.ҩƷID = B.ҩƷID AND b.ҩƷID=D.ID " & _
                    " AND B.ҩƷID=E.�շ�ϸĿID(+) AND E.����(+)=3 " & _
                    " AND A.���� = 6 AND A.���ϵ��=-1 AND A.NO = [1] AND A.��¼״̬ =[3] " & _
                    " AND A.ҩƷID=Z.ҩƷID(+) AND NVL(A.����,0)=Z.����(+) " & _
                    " ORDER BY " & strSqlOrder
            End If
            
            Set rsInitCard = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, mstr���ݺ�, lngStockid, mint��¼״̬)
        
            If rsInitCard.EOF Then
                mintParallelRecord = 2
                Exit Sub
            End If
            
            If mint�༭״̬ = 2 Or mint�༭״̬ = 4 Or mint�༭״̬ = 6 Then
                mintApplyType = rsInitCard!���췽ʽ
            End If
            mshBill.Active = IIf(mintApplyType = 0, True, IsHavePrivs(mstrPrivs, "�Զ�����ʱ�޸�ҩƷ����"))
            
            If mint�༭״̬ = 7 Then
                Txt������ = gstrUserName
                Txt�������� = Format(dateCurDate, "yyyy-mm-dd hh:mm:ss")
                Txt����� = gstrUserName
                Txt������� = Format(dateCurDate, "yyyy-mm-dd hh:mm:ss")
            Else
                Txt������ = rsInitCard!������
                If mint�༭״̬ = 2 Then
                    Txt������ = gstrUserName
                End If
                Txt�������� = Format(rsInitCard!��������, "yyyy-mm-dd hh:mm:ss")
                Txt����� = IIf(IsNull(rsInitCard!�����), "", rsInitCard!�����)
                Txt������� = IIf(IsNull(rsInitCard!�������), "", Format(rsInitCard!�������, "yyyy-mm-dd hh:mm:ss"))
            End If
            txtժҪ.Text = IIf(IsNull(rsInitCard!ժҪ), "", rsInitCard!ժҪ)
            
            If (mint�༭״̬ = 2 Or mint�༭״̬ = 3) And Txt����� <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            If mint�༭״̬ = 2 Then
                Set mcolUsedCount = New Collection
            End If
            
            With mshBill
                Do While Not rsInitCard.EOF
                    intRow = rsInitCard.AbsolutePosition
                    'IntRow = rsInitCard!���
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
                    .TextMatrix(intRow, mconIntCol��Դ) = NVL(rsInitCard!ҩƷ��Դ)
                    .TextMatrix(intRow, mconIntCol����ҩ��) = NVL(rsInitCard!����ҩ��)
                    .TextMatrix(intRow, mconIntCol���) = rsInitCard!���
                    .TextMatrix(intRow, mconIntCol���) = IIf(IsNull(rsInitCard!���), "", rsInitCard!���)
                    .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsInitCard!����), "", rsInitCard!����)
                    .TextMatrix(intRow, mconIntCol��λ) = rsInitCard!��λ
                    If IIf(IsNull(rsInitCard!����ϵ��), 0, rsInitCard!����ϵ��) = 0 Or NVL(rsInitCard!�ͻ���װ) = "" Or NVL(rsInitCard!�ͻ���λ) = "" Then
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
                    
                    .TextMatrix(intRow, mconIntCol��׼�ĺ�) = IIf(IsNull(rsInitCard!��׼�ĺ�), "", rsInitCard!��׼�ĺ�)
                                
                    .TextMatrix(intRow, mconIntCol��д����) = GetFormat(rsInitCard!��д����, intNumberDigit)
                    .TextMatrix(intRow, mconIntColʵ������) = GetFormat(rsInitCard!ʵ������, intNumberDigit)
                    
                    If mint�༭״̬ = 6 Then
                        .TextMatrix(intRow, mconIntColԭʼ����) = GetFormat(rsInitCard!ԭʼ����, intNumberDigit)
                    End If
                    
                    .TextMatrix(intRow, mconIntCol�ɹ���) = GetFormat(rsInitCard!�ɱ���, intCostDigit)
                    
                    .TextMatrix(intRow, mconIntCol�ɹ����) = GetFormat(IIf(mint�༭״̬ = 7, 0, rsInitCard!�ɱ����), intMoneyDigit)
                    .TextMatrix(intRow, mconIntCol�ۼ�) = GetFormat(rsInitCard!���ۼ�, intPricedigit)
                    .TextMatrix(intRow, mconIntCol�ۼ۽��) = GetFormat(rsInitCard!���۽��, intMoneyDigit)
                    .TextMatrix(intRow, mconintCol���) = GetFormat(rsInitCard!���, intMoneyDigit)
                    
                    .TextMatrix(intRow, mconIntCol���Ч��) = IIf(IsNull(rsInitCard!���Ч��), "0", rsInitCard!���Ч��) & "||" & rsInitCard!�Ƿ��� & "||" & rsInitCard!ҩ����������
                    .TextMatrix(intRow, mconIntCol����) = IIf(IsNull(rsInitCard!����), "0", rsInitCard!����)
                    .TextMatrix(intRow, mconIntCol����ϵ��) = rsInitCard!����ϵ��
                    .TextMatrix(intRow, mconIntColָ�������) = rsInitCard!ָ�������
                    .TextMatrix(intRow, mconIntCol��������) = IIf(IsNull(rsInitCard!��������), "0", rsInitCard!��������)
                    .TextMatrix(intRow, mconIntCol��������) = IIf(IsNull(rsInitCard!��������), "0", rsInitCard!��������)
                    .TextMatrix(intRow, mconIntColʵ�ʲ��) = IIf(IsNull(rsInitCard!ʵ�ʲ��), "0", rsInitCard!ʵ�ʲ��)
                    .TextMatrix(intRow, mconIntColʵ�ʽ��) = IIf(IsNull(rsInitCard!ʵ�ʽ��), "0", rsInitCard!ʵ�ʽ��)
                    .TextMatrix(intRow, mconIntCol�ϴι�Ӧ��ID) = rsInitCard!�ϴι�Ӧ��ID
                                        
                    If mint�༭״̬ = 7 Then
                        .TextMatrix(intRow, mconintCol��ʵ����) = rsInitCard!��ʵ����
                    End If
                        
                    
                    If mint�༭״̬ = 2 Then
                        numUseAbleCount = 0
                        For Each vardrug In mcolUsedCount
                            If vardrug(0) = CStr(rsInitCard!ҩƷid & IIf(IsNull(rsInitCard!����), "0", rsInitCard!����)) Then
                                numUseAbleCount = vardrug(1)
                                mcolUsedCount.Remove vardrug(0)
                                Exit For
                            End If
                        Next
                        mcolUsedCount.Add Array(CStr(rsInitCard!ҩƷid & IIf(IsNull(rsInitCard!����), "0", rsInitCard!����)), CStr(numUseAbleCount + IIf(IsNull(rsInitCard!��д����), "0", rsInitCard!��д����))), CStr(rsInitCard!ҩƷid) & CStr(IIf(IsNull(rsInitCard!����), "0", rsInitCard!����))
                        
                    End If
                    
                    rsInitCard.MoveNext
                Loop
            End With
            rsInitCard.Close
    End Select
    
    Call get�������
    Call RefreshRowNO(mshBill, mconIntCol�к�, 1)
    Call ��ʾ�ϼƽ��
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub initGrid()
    With mshBill
        .Active = True
        .Cols = mconIntColS
        
        .MsfObj.FixedCols = 1
        
        .TextMatrix(0, mconIntCol�к�) = ""
        .TextMatrix(0, mconIntColҩ��) = "ҩƷ���������"
        .TextMatrix(0, mconIntCol��Ʒ��) = "��Ʒ��"
        .TextMatrix(0, mconIntCol��Դ) = "ҩƷ��Դ"
        .TextMatrix(0, mconIntCol����ҩ��) = "����ҩ��"
        .TextMatrix(0, mconIntCol���) = "���"
        .TextMatrix(0, mconIntCol���) = "���"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntCol��λ) = "��λ"
        .TextMatrix(0, mconIntCol�ͻ���λ) = "�ͻ���λ"
        .TextMatrix(0, mconIntCol����) = "����"
        .TextMatrix(0, mconIntColЧ��) = IIf(gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1, "��Ч����", "ʧЧ��")
        .TextMatrix(0, mconIntCol��׼�ĺ�) = "��׼�ĺ�"
        .TextMatrix(0, mconintcol��ǰ���) = "��ǰ���"
        .TextMatrix(0, mconintcol�Է����) = "�Է����"
        .TextMatrix(0, mconIntCol��������) = "��������"
        .TextMatrix(0, mconIntCol��д����) = IIf(mint�༭״̬ = 7, "����", "��д����")
        .TextMatrix(0, mconIntColʵ������) = IIf(mint�༭״̬ = 7, "��������", "ʵ������")
        
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
        .TextMatrix(0, mconIntCol����ҩƷ) = "����ҩƷ"
        .TextMatrix(0, mconIntColԭʼ����) = "ԭʼ����"
        
        .TextMatrix(1, 0) = ""
        .TextMatrix(1, mconIntCol�к�) = "1"
        
        .ColWidth(0) = 0
        .ColWidth(mconIntCol���) = 0
        .ColWidth(mconIntCol�к�) = 300
        .ColWidth(mconIntColҩ��) = 2200
        .ColWidth(mconIntCol��Ʒ��) = 2000
        .ColWidth(mconIntCol��Դ) = 900
        .ColWidth(mconIntCol����ҩ��) = 900
        .ColWidth(mconIntCol���) = 900
        .ColWidth(mconIntCol����) = 800
        .ColWidth(mconIntCol��λ) = 400
        .ColWidth(mconIntCol�ͻ���λ) = 2000
        .ColWidth(mconIntCol����) = 800
        .ColWidth(mconIntColЧ��) = 1000
        .ColWidth(mconIntCol��׼�ĺ�) = 1000
        .ColWidth(mconIntCol��������) = 0
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
        .ColWidth(mconIntCol����ҩƷ) = 0
        .ColWidth(mconIntColԭʼ����) = 0
        
        
        '-1����ʾ���п���ѡ���ǲ����ͣ�"��"��" "��
        ' 0����ʾ���п���ѡ�񣬵������޸�
        ' 1����ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
        ' 2����ʾ�����������У��ⲿ��ʾΪ��ťѡ�񣬵���������ѡ���
        ' 3����ʾ������ѡ���У��ⲿ��ʾΪ������ѡ��
        '4:  ��ʾ����Ϊ�������ı����û�����
        '5:  ��ʾ���в�����ѡ��

        .ColData(0) = 5
        .ColData(mconIntCol��������) = 0
        .ColData(mconIntCol��Ʒ��) = 5
        .ColData(mconIntCol�к�) = 5
        .ColData(mconIntCol��Դ) = 5
        .ColData(mconIntCol����ҩ��) = 5
        .ColData(mconIntCol���) = 5
        .ColData(mconIntCol���) = 5
        .ColData(mconIntCol����) = 5
        .ColData(mconIntCol��λ) = 5
        .ColData(mconIntCol�ͻ���λ) = 5
        .ColData(mconIntCol����) = 5
        .ColData(mconIntColЧ��) = 5
        .ColData(mconIntCol��׼�ĺ�) = 5
        .ColData(mconIntColҩƷ���������) = 5
        .ColData(mconIntColҩƷ����) = 5
        .ColData(mconIntColҩƷ����) = 5
        .ColData(mconIntCol����ҩƷ) = 5
        .ColData(mconIntColԭʼ����) = 5
        
        '��״̬Ϊ���ܱ༭
        .ColData(mconintcol��ǰ���) = 5
        .ColData(mconintcol�Է����) = 5
        
        '��������Ϊ�༭״̬���������޸ģ�ʱ�ɼ�
        If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Or mint�༭״̬ = 5 Then
            .ColWidth(mconintcol��ǰ���) = 1100
            '��û����ʾ�Է����Ȩ�޵�ʱ������ʾ�Է����
            If IsHavePrivs(mstrPrivs, "��ʾ�Է����") Then
                .ColWidth(mconintcol�Է����) = 1100
            Else
                .ColWidth(mconintcol�Է����) = 0
            End If
        Else
            .ColWidth(mconintcol��ǰ���) = 0
            .ColWidth(mconintcol�Է����) = 0
        End If
        
        If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Or mint�༭״̬ = 5 Then
            
            cboStock.Enabled = True
            txtժҪ.Enabled = True
            
            .ColData(mconIntColҩ��) = 1
            .ColData(mconIntCol��д����) = 4
            .ColData(mconIntColʵ������) = 5
        ElseIf mint�༭״̬ = 4 Or mint�༭״̬ = 6 Then
            cboStock.Enabled = False
            txtժҪ.Enabled = False
            
            .ColData(mconIntCol��д����) = 5
            .ColData(mconIntColʵ������) = IIf(mint�༭״̬ <> 6, 4, 5)
            .ColData(mconIntColҩ��) = 0
        ElseIf mint�༭״̬ = 7 Then
            cboStock.Enabled = False
            txtժҪ.Enabled = True
            
            .ColData(mconIntCol��д����) = 5
            .ColData(mconIntColʵ������) = 4
            .ColData(mconIntColҩ��) = 0
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
        .ColData(mconintCol��ʵ����) = 5
        .ColData(mconIntCol�ϴι�Ӧ��ID) = 5
        
        .ColAlignment(mconIntColҩ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��Ʒ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol���) = flexAlignLeftCenter
        .ColAlignment(mconIntCol����) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��λ) = flexAlignCenterCenter
        .ColAlignment(mconIntCol�ͻ���λ) = flexAlignCenterCenter
        .ColAlignment(mconIntCol����) = flexAlignLeftCenter
        .ColAlignment(mconIntColЧ��) = flexAlignLeftCenter
        .ColAlignment(mconIntCol��׼�ĺ�) = flexAlignLeftCenter
        .ColAlignment(mconintcol��ǰ���) = flexAlignRightCenter
        .ColAlignment(mconintcol�Է����) = flexAlignRightCenter
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
        If InStr(1, "34", mint�༭״̬) <> 0 Then .ColData(mconIntColҩ��) = 0
    End With
    txtժҪ.MaxLength = GetLength("ҩƷ�շ���¼", "ժҪ")
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    
    With Pic����
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - IIf(staThis.Visible, staThis.Height, 0) - IIf(Me.cmdȫ������.Visible, 350, 0) - .Top - 100 - CmdCancel.Height - 200
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
    End With
    
    
    LblStock.Left = mshBill.Left
    cboStock.Left = LblStock.Left + LblStock.Width + 100
    
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
    

    
    With cmdHelp
        .Left = Pic����.Left + mshBill.Left
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
        
    With cmdFind
        .Top = CmdCancel.Top
    End With
    
    With lblCode
        .Top = CmdCancel.Top + 50
    End With
    With txtCode
        .Top = CmdCancel.Top + 30
    End With
    
    With chkExportPlan
        .Top = lblCode.Top
    End With
    
    With cmdȫ��
        If .Visible = True Then
            .Left = Me.CmdSave.Left
            .Top = Me.CmdSave.Top
        End If
    End With
    
    With cmdȫ������
        If .Visible = True Then
            .Left = Me.CmdCancel.Left
            .Top = Me.CmdCancel.Top
        End If
    End With
    
    If mint�༭״̬ = 5 And Me.cmdȫ��.Visible = True Then
        With Me.CmdSave
            .Left = Me.CmdSave.Left
            .Top = Me.CmdSave.Height + Me.CmdSave.Top + 100
        End With
    
        With Me.CmdCancel
            .Left = Me.CmdCancel.Left
            .Top = Me.CmdCancel.Height + Me.CmdCancel.Top + 100
        End With
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mintApplyType = 0
    mstrEndTime = ""
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\ҩƷ�������", "ͬ�����ɼƻ���", Me.chkExportPlan.Value)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\ҩƷ�������", "ҩƷ������ʾ��ʽ", mintDrugNameShow)
    
    If mblnChange = False Or mint�༭״̬ = 4 Then
        SaveWinState Me, App.ProductName, MStrCaption
        Call ReleaseSelectorRS
        mblnUpdate = False
        zlPlugIn_Unload mobjPlugIn
        Exit Sub
    End If
    If MsgBox("���ݿ����Ѹı䣬��δ���̣���Ҫ�˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
        Exit Sub
    Else
        SaveWinState Me, App.ProductName, MStrCaption
    End If
    Call ReleaseSelectorRS
    mblnUpdate = False
    zlPlugIn_Unload mobjPlugIn
End Sub

Private Sub mnuColDrug_Click(index As Integer)
    Dim n As Integer
    
    With mnuColDrug
        For n = 0 To .count - 1
            .Item(n).Checked = False
        Next
        
        .Item(index).Checked = True
        
        Call SetDrugName(index)
    End With
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
Private Sub mshBill_AfterAddRow(Row As Long)
    Call RefreshRowNO(mshBill, mconIntCol�к�, Row)
End Sub

Private Sub mshBill_AfterDeleteRow()
    Call ��ʾ�ϼƽ��
    Call RefreshRowNO(mshBill, mconIntCol�к�, mshBill.Row)
End Sub

Private Sub mshBill_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    If InStr(1, "3467", mint�༭״̬) <> 0 Then
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
    Dim RecReturn As Recordset
    Dim i As Integer
    Dim intCurRow As Integer
    Dim intOldRow As Integer
    
    intOldRow = mshBill.Row
    mshBill.CmdEnable = False
'    Set RecReturn = FrmҩƷѡ����.ShowME(Me, 2, cboStock.ItemData(cboStock.ListIndex), _
'        mlngStockID, mlngStockID, mbln�����, IIf(mint��ȷ���� = 0, False, True), _
'        False, False, IsHavePrivs(mstrPrivs, "��ʾ�Է����"))
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(2, "ҩƷ�������", cboStock.ItemData(cboStock.ListIndex), mlngStockID, mlngStockID, , , , , , mbln�����)
    End If
    Set RecReturn = frmSelector.ShowMe(Me, 0, 2, , , , cboStock.ItemData(cboStock.ListIndex), mlngStockID, mlngStockID, , mbln�����, IIf(mint��ȷ���� = 0, False, True), IsHavePrivs(mstrPrivs, "��ʾ�Է����"), False, , mstrPrivs)
    If RecReturn.RecordCount > 0 Then
        Set RecReturn = CheckData(RecReturn)  '����ظ���¼ �����ظ���¼��ҩƷid���ػ���
    End If
    mshBill.CmdEnable = True
    
    If RecReturn.RecordCount > 0 Then
        RecReturn.MoveFirst
        For i = 1 To RecReturn.RecordCount
            intCurRow = mshBill.Row
            With mshBill
                .TextMatrix(intCurRow, mconIntCol�к�) = .Row
                SetColValue .Row, RecReturn!ҩƷid, "[" & RecReturn!ҩƷ���� & "]", RecReturn!ͨ����, IIf(IsNull(RecReturn!��Ʒ��), "", RecReturn!��Ʒ��), _
                    NVL(RecReturn!ҩƷ��Դ), NVL(RecReturn!����ҩ��), _
                    IIf(IsNull(RecReturn!���), "", RecReturn!���), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                    Choose(mintUnit, RecReturn!�ۼ۵�λ, RecReturn!���ﵥλ, RecReturn!סԺ��λ, RecReturn!ҩ�ⵥλ), _
                    IIf(IsNull(RecReturn!�ۼ�), 0, RecReturn!�ۼ�), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                    IIf(IsNull(RecReturn!Ч��), "", RecReturn!Ч��), _
                    IIf(IsNull(RecReturn!���Ч��), "0", RecReturn!���Ч��), _
                    RecReturn!ҩ�����, _
                    IIf(IsNull(RecReturn!��������), "0", RecReturn!��������), _
                    IIf(IsNull(RecReturn!ʵ�ʽ��), "0", RecReturn!ʵ�ʽ��), _
                    IIf(IsNull(RecReturn!ʵ�ʲ��), "0", RecReturn!ʵ�ʲ��), _
                    IIf(IsNull(RecReturn!ָ�������), "0", RecReturn!ָ�������), _
                    Choose(mintUnit, 1, RecReturn!�����װ, RecReturn!סԺ��װ, RecReturn!ҩ���װ), _
                    IIf(IsNull(RecReturn!����), 0, RecReturn!����), RecReturn!ʱ��, RecReturn!ҩ������, _
                    RecReturn!�ϴι�Ӧ��ID, IIf(IsNull(RecReturn!��׼�ĺ�), "", RecReturn!��׼�ĺ�)
                .Col = mconIntCol��д����
'                .TextMatrix(.Row, mconIntCol����ҩƷ) = True
                
                If (.TextMatrix(intCurRow, 0) = "" Or intCurRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, 0) <> "" Then
                    .rows = .rows + 1
                End If
                .Row = .rows - 1
                RecReturn.MoveNext
            End With
        Next
        mshBill.Row = intOldRow
        RecReturn.Close
    End If
End Sub

Private Sub mshBill_DblClick(Cancel As Boolean)
    If Me.mshBill.Row <> Me.mshBill.rows - 1 Then
        If Me.mshBill.Col = mconIntCol�������� And Me.mshBill.Row <> 0 Then
            If Val(Me.mshBill.TextMatrix(Me.mshBill.Row, mconIntCol��д����)) = 0 Then
                Me.mshBill.TextMatrix(Me.mshBill.Row, mconIntCol��д����) = Me.mshBill.TextMatrix(Me.mshBill.Row, mconIntCol��������)
                Me.mshBill.TextMatrix(Me.mshBill.Row, mconIntColʵ������) = Me.mshBill.TextMatrix(Me.mshBill.Row, mconIntCol��������)
            Else
                Me.mshBill.TextMatrix(Me.mshBill.Row, mconIntCol��д����) = 0
                Me.mshBill.TextMatrix(Me.mshBill.Row, mconIntColʵ������) = 0
            End If
        End If
        
        If Val(Me.mshBill.TextMatrix(Me.mshBill.Row, mconIntCol��д����)) <> 0 Then
            Call GetPrice(Me.mshBill.Row)
        Else
             With Me.mshBill
                .TextMatrix(Me.mshBill.Row, mconIntCol�ۼ۽��) = 0
                .TextMatrix(Me.mshBill.Row, mconintCol���) = 0
                .TextMatrix(Me.mshBill.Row, mconIntCol�ɹ���) = 0
                .TextMatrix(Me.mshBill.Row, mconIntCol�ɹ����) = 0
            End With
        End If
        
        Call ��ʾ�ϼƽ��
    End If
End Sub

Private Sub mshbill_EditChange(curText As String)
    With mshBill
        If .Col <> mconIntCol���� Then
            mshBill.Text = UCase(curText)
            mshBill.SelStart = Len(mshBill.Text)
        End If
    End With
    mblnChange = True
End Sub

Private Sub mshBill_EditKeyPress(KeyAscii As Integer)
    Dim strkey As String
    Dim intDigit As Integer
    
    With mshBill
        If .Col = mconIntCol��д���� Or .Col = mconIntColʵ������ Then
            strkey = .Text
            If strkey = "" Then
                strkey = .TextMatrix(.Row, .Col)
            End If
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
        End If
    End With
End Sub

Private Sub mshbill_EnterCell(Row As Long, Col As Long)
    With mshBill
        If Row > 0 Then
            .SetRowColor CLng(Row), &HFFCECE, True
        End If
        Select Case .Col
            Case mconIntColҩ��
                .TxtCheck = False
                .MaxLength = 40
                'ֻ��ҩ���в���ʾ�ϼ���Ϣ�Ϳ����
                Call ��ʾ�ϼƽ��
                Call ��ʾ�����
                
            Case mconIntCol����
                .TxtCheck = True
                .TextMask = "1234567890"
                .MaxLength = 8
            
            Case mconIntColЧ��
                .TxtCheck = True
                .TextMask = "1234567890-"
                .MaxLength = 10
                If .TextMatrix(.Row, mconIntCol����) <> "" And .ColData(.Col) = 2 Then
                    Dim strxq As String
                    
                    If IsNumeric(.TextMatrix(.Row, mconIntCol����)) And .TextMatrix(.Row, mconIntCol���Ч��) <> "" Then
                        If Split(.TextMatrix(.Row, mconIntCol���Ч��), "||")(0) <> 0 Then
                            strxq = .TextMatrix(.Row, mconIntCol����)
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
            Case mconIntCol��д����, mconIntColʵ������
                .TxtCheck = True
                .MaxLength = 16
                .TextMask = ".1234567890"
                Call ��ʾ�����
                
        End Select
    End With
End Sub

Private Sub mshbill_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strkey As String
    Dim rsDrug As New Recordset
    Dim strUnitQuantity As String
    Dim i As Integer
    Dim intCurRow As Integer
    Dim intOldRow As Integer
    
    intOldRow = mshBill.Row
    If KeyCode <> vbKeyReturn Then Exit Sub
    With mshBill
        strkey = UCase(Trim(.Text))
        
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
                    Dim RecReturn As Recordset
                    Dim sngLeft As Single
                    Dim sngTop As Single
                    
                    sngLeft = Me.Left + Pic����.Left + mshBill.Left + mshBill.MsfObj.CellLeft + Screen.TwipsPerPixelX
                    sngTop = Me.Top + Me.Height - Me.ScaleHeight + Pic����.Top + mshBill.Top + mshBill.MsfObj.CellTop + mshBill.MsfObj.CellHeight  '  50
                    If sngTop + 3630 > Screen.Height Then
                        sngTop = sngTop - mshBill.MsfObj.CellHeight - 3630
                    End If
                    
'                    Set RecReturn = FrmҩƷ��ѡѡ����.ShowME(Me, 2, cboStock.ItemData(cboStock.ListIndex), _
'                        mlngStockID, mlngStockID, strkey, sngLeft, sngTop, mbln�����, _
'                        IIf(mint��ȷ���� = 0, False, True), False, False, IsHavePrivs(mstrPrivs, "��ʾ�Է����"))

                    If grsMaster.State = adStateClosed Then
                        Call SetSelectorRS(2, "ҩƷ�������", cboStock.ItemData(cboStock.ListIndex), mlngStockID, mlngStockID, , , , , , mbln�����)
                    End If
                    Set RecReturn = frmSelector.ShowMe(Me, 1, 2, strkey, sngLeft, sngTop, cboStock.ItemData(cboStock.ListIndex), mlngStockID, mlngStockID, , mbln�����, IIf(mint��ȷ���� = 0, False, True), IsHavePrivs(mstrPrivs, "��ʾ�Է����"), False, , mstrPrivs)
                    If RecReturn.RecordCount > 0 Then
                        Set RecReturn = CheckData(RecReturn)  '����ظ���¼ �����ظ���¼��ҩƷid���ػ���
                    End If
                    If RecReturn.RecordCount > 0 Then
                        RecReturn.MoveFirst
                        For i = 1 To RecReturn.RecordCount
                            intCurRow = .Row
                            .TextMatrix(intCurRow, mconIntCol�к�) = .Row
                            If SetColValue(.Row, RecReturn!ҩƷid, "[" & RecReturn!ҩƷ���� & "]", RecReturn!ͨ����, IIf(IsNull(RecReturn!��Ʒ��), "", RecReturn!��Ʒ��), _
                                    NVL(RecReturn!ҩƷ��Դ), NVL(RecReturn!����ҩ��), _
                                    IIf(IsNull(RecReturn!���), "", RecReturn!���), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                                    Choose(mintUnit, RecReturn!�ۼ۵�λ, RecReturn!���ﵥλ, RecReturn!סԺ��λ, RecReturn!ҩ�ⵥλ), _
                                    IIf(IsNull(RecReturn!�ۼ�), 0, RecReturn!�ۼ�), IIf(IsNull(RecReturn!����), "", RecReturn!����), _
                                    IIf(IsNull(RecReturn!Ч��), "", RecReturn!Ч��), _
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
                            If (.TextMatrix(intCurRow, 0) = "" Or intCurRow = 1 Or .Row = .rows - 1) And .TextMatrix(.rows - 1, 0) <> "" Then
                                .rows = .rows + 1
                            End If
                            .Row = .rows - 1
                            RecReturn.MoveNext
                        Next
                        .Row = intOldRow
                        .Col = mconIntCol��д����
                    Else
                        .TextMatrix(.Row, mconIntCol����ҩƷ) = True
                        Cancel = True
                    End If
                    Call ��ʾ�����
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
                End If
                
                If Len(strkey) < 8 Then
                    MsgBox "�Բ������ų��Ȳ���������Ϊ8λ,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
            Case mconIntColЧ��
                '�д���
                If strkey <> "" Then
                    If Len(strkey) = 8 And InStr(1, strkey, "-") = 0 Then
                        strkey = TranNumToDate(strkey)
                        If strkey = "" Then
                            MsgBox "�Բ���Ч�ڱ���Ϊ�����ͣ�", vbInformation + vbOKOnly, gstrSysName
                            Cancel = True
                            Exit Sub
                        End If
                        .Text = strkey
                        Exit Sub
                    End If
                    If Not IsDate(strkey) Then
                        MsgBox "�Բ���Ч�ڱ���Ϊ��������(2000-10-10) ��20001010��,�����䣡", vbInformation + vbOKOnly, gstrSysName
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
                If Not IsNumeric(strkey) And strkey <> "" Then
                    MsgBox "�Բ�����������Ϊ������,�����䣡", vbInformation + vbOKOnly, gstrSysName
                    Cancel = True
                    .TxtSetFocus
                    Exit Sub
                End If
                
                If strkey <> "" Then
'                    If Val(strkey) = 0 Then
'                        MsgBox "�Բ�����������Ϊ��,�����䣡", vbInformation + vbOKOnly, gstrSysName
'                        Cancel = True
'                        .TxtSetFocus
'                        Exit Sub
'                    End If
'
'                    If Val(strkey) < 0.00001 Then
'                        MsgBox "�Բ�����������Ϊ�������Ҳ���С��0.00001,�����䣡", vbInformation + vbOKOnly, gstrSysName
'                        Cancel = True
'                        .TxtSetFocus
'                        Exit Sub
'                    End If
                    
                    If Val(strkey) >= 10 ^ 11 - 1 Then
                        MsgBox "��������С��" & (10 ^ 11 - 1), vbInformation + vbOKOnly, gstrSysName
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    If Not CompareUsableQuantity(.Row, strkey) Then
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                    
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    '�ɱ��۵Ĺ�ʽ��     ������=����*�ۼ�
                    '                  ������=������*��ʵ�ʲ��/ʵ�ʽ�
                    '                  if ʵ�ʽ��=0 then  ������=������*ָ�������
                    '                  ���ۣ��ɱ��ۣ�=ֱ�Ӵӿ�����ȡƽ���ɱ���
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    'ʵ�ʽ��=0������£����ο��Ǵӡ�����¼���ϴβɹ��ۡ�����ҩƷ���ĳɱ��ۡ�����ָ������ʡ�ȡֵ
                    
                    strkey = GetFormat(strkey, mintNumberDigit)
                    .Text = strkey
                    
                    If .TextMatrix(.Row, mconIntCol�ۼ�) <> "" Then
                        .TextMatrix(.Row, mconIntCol�ۼ۽��) = GetFormat(.TextMatrix(.Row, mconIntCol�ۼ�) * strkey, mintMoneyDigit)
                    End If
                    
                    If strkey <> 0 Then
'                        .TextMatrix(.Row, mconIntCol�ɹ���) = FormatEx((Val(.TextMatrix(.Row, mconIntCol�ۼ۽��)) - .TextMatrix(.Row, mconintCol���)) / strkey, mintCostDigit)
                        If mint�༭״̬ <> 7 Then .TextMatrix(.Row, mconIntCol�ɹ���) = GetFormat(Get�ɱ���(Val(.TextMatrix(.Row, 0)), cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, mconIntCol����))) * Val(.TextMatrix(.Row, mconIntCol����ϵ��)), mintCostDigit)
                    End If
                    
                    .TextMatrix(.Row, mconIntCol�ɹ����) = GetFormat(Val(.TextMatrix(.Row, mconIntCol�ɹ���)) * strkey, mintMoneyDigit)
                    
'                    If mint�༭״̬ = 7 Then
                        .TextMatrix(.Row, mconintCol���) = GetFormat(Val(.TextMatrix(.Row, mconIntCol�ۼ۽��)) - Val(.TextMatrix(.Row, mconIntCol�ɹ����)), mintMoneyDigit)
'                    Else
'                        .TextMatrix(.Row, mconintCol���) = GetFormat(Get������(Val(cboStock.ItemData(cboStock.ListIndex)), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol����)), Val(.TextMatrix(.Row, mconIntColʵ�ʽ��)), Val(.TextMatrix(.Row, mconIntColʵ�ʲ��)), Val(.TextMatrix(.Row, mconIntCol�ۼ۽��)), Val(strkey) * Val(mshBill.TextMatrix(.Row, mconIntCol����ϵ��))), mintMoneyDigit)
'                    End If
                    
                    If .Col = mconIntCol��д���� Then
                        .TextMatrix(.Row, mconIntColʵ������) = strkey
                    End If
                    
                    
                End If
                
                ��ʾ�ϼƽ��
            
        End Select
    End With
End Sub

Private Sub GetPrice(ByVal intRow As Integer)
    With Me.mshBill
        .TextMatrix(intRow, mconIntCol�ۼ۽��) = GetFormat(.TextMatrix(intRow, mconIntCol�ۼ�) * Me.mshBill.TextMatrix(intRow, mconIntCol��д����), mintMoneyDigit)
        .TextMatrix(intRow, mconIntCol�ɹ���) = GetFormat(Get�ɱ���(Val(.TextMatrix(intRow, 0)), cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(intRow, mconIntCol����))) * Val(.TextMatrix(intRow, mconIntCol����ϵ��)), mintCostDigit)
        .TextMatrix(intRow, mconIntCol�ɹ����) = GetFormat(Val(.TextMatrix(intRow, mconIntCol�ɹ���)) * Val(Me.mshBill.TextMatrix(intRow, mconIntCol��д����)), mintMoneyDigit)
        .TextMatrix(intRow, mconintCol���) = GetFormat(Val(.TextMatrix(intRow, mconIntCol�ۼ۽��)) - .TextMatrix(intRow, mconIntCol�ɹ����), mintMoneyDigit)
    End With
End Sub

'��ҩƷĿ¼��ȡֵ��������Ӧ����
Private Function SetColValue(ByVal intRow As Integer, ByVal lngҩƷID As Long, _
    ByVal strҩƷ���� As String, ByVal strͨ���� As String, ByVal str��Ʒ�� As String, ByVal strҩƷ��Դ As String, _
    ByVal str����ҩ�� As String, ByVal str��� As String, ByVal str���� As String, _
    ByVal str��λ As String, ByVal num�ۼ� As Double, ByVal str���� As String, _
    ByVal strЧ�� As String, ByVal int���Ч�� As Integer, ByVal int�������� As Integer, _
    ByVal num�������� As Double, ByVal numʵ�ʽ�� As Double, ByVal numʵ�ʲ�� As Double, _
    ByVal numָ������� As Double, ByVal num����ϵ�� As Double, ByVal lng���� As Long, _
    ByVal int�Ƿ��� As Integer, ByVal intҩ������ As Integer, ByVal lng�ϴι�Ӧ��ID As Long, ByVal str��׼�ĺ� As String) As Boolean
    
    Dim IntCount As Integer
    Dim intCol As Integer
    Dim dblPrice As Double
    Dim rsPrice As New Recordset
    Dim strҩ�� As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    SetColValue = False
    
    '����Ƿ��ظ�
'    If Not CheckRepeatMedicine(mshBill, lngҩƷID & "," & "0" & "|" & IIf(mint��ȷ���� = 1, lng����, 0) & "," & mconIntCol����, intRow) Then
'        Exit Function
'    End If
    
    With mshBill
'        If int�Ƿ��� = 1 Then
'            gstrSQL = "select Decode(Nvl(����, 0), 0, ʵ�ʽ�� / ʵ������, Nvl(���ۼ�, ʵ�ʽ�� / ʵ������))*" & num����ϵ�� & " as  �ۼ� " _
'                & "  from ҩƷ��� " _
'                & " where �ⷿid=[1] " _
'                & " and ҩƷid=[2] " _
'                & " and ����=1 and ʵ������>0 and " _
'                & " nvl(����,0)=[3]"
'            Set rsPrice = zldatabase.OpenSQLRecord(gstrSQL, MStrCaption, cboStock.ItemData(cboStock.ListIndex), lngҩƷID, lng����)
'
'            If rsPrice.EOF Then
'                If mint��ȷ���� = 1 Then
'                    MsgBox "ʱ��ҩƷû�п�棬���ܳ��⣬���飡", vbOKOnly, gstrSysName
'                    Exit Function
'                End If
'            Else
'                dblPrice = rsPrice.Fields(0).Value
'            End If
'        End If
        
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
        If num����ϵ�� = 0 Or NVL(rsTemp!�ͻ���װ) = "" Or NVL(rsTemp!�ͻ���λ) = "" Then
            .TextMatrix(intRow, mconIntCol�ͻ���λ) = ""
        Else
            .TextMatrix(intRow, mconIntCol�ͻ���λ) = rsTemp!�ͻ���λ & "(1" & rsTemp!�ͻ���λ & "=" & zlStr.FormatEx(rsTemp!�ͻ���װ / num����ϵ��, 1, , True) & str��λ & ")"
        End If
        
        .TextMatrix(intRow, mconIntCol�ۼ�) = GetFormat(num�ۼ� * num����ϵ��, mintPriceDigit)
        .TextMatrix(intRow, mconIntCol��������) = int��������
        .TextMatrix(intRow, mconIntCol��������) = GetFormat(num��������, mintNumberDigit)
        .TextMatrix(intRow, mconIntCol���Ч��) = int���Ч�� & "||" & int�Ƿ��� & "||" & intҩ������
        .TextMatrix(intRow, mconIntColʵ�ʲ��) = numʵ�ʲ��
        .TextMatrix(intRow, mconIntColʵ�ʽ��) = numʵ�ʽ��
        .TextMatrix(intRow, mconIntColָ�������) = numָ�������
        .TextMatrix(intRow, mconIntCol����ϵ��) = num����ϵ��
        .TextMatrix(intRow, mconIntCol�ϴι�Ӧ��ID) = lng�ϴι�Ӧ��ID

        If mint��ȷ���� = 1 Then
            .TextMatrix(intRow, mconIntCol����) = lng����
            .TextMatrix(intRow, mconIntCol����) = str����
            .TextMatrix(intRow, mconIntColЧ��) = Format(strЧ��, "yyyy-mm-dd")
            .TextMatrix(intRow, mconIntCol��׼�ĺ�) = str��׼�ĺ�
        Else
            .TextMatrix(intRow, mconIntCol����) = 0
            .TextMatrix(intRow, mconIntCol����) = ""
            .TextMatrix(intRow, mconIntColЧ��) = ""
            .TextMatrix(intRow, mconIntCol��׼�ĺ�) = ""
        End If
        If int�Ƿ��� = 1 Then
            dblPrice = Get�ۼ�(True, lngҩƷID, cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(intRow, mconIntCol����))) 'Get�۸�(lngҩƷid, Val(.TextMatrix(intRow, mconIntCol����)), num����ϵ��)
            .TextMatrix(intRow, mconIntCol�ۼ�) = GetFormat(dblPrice * num����ϵ��, mintPriceDigit)
        End If
        Call CheckLapse(strЧ��)
        
        '�Ƿ񳣱�ҩƷ
        Dim rsTmp As ADODB.Recordset
        gstrSQL = "select nvl(�Ƿ񳣱�,0) �Ƿ񳣱� from ҩƷ��� where ҩƷid=[1]"
        Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, lngҩƷID)
        .TextMatrix(intRow, mconIntCol����ҩƷ) = IIf(rsTmp!�Ƿ񳣱� = 1, False, True)
    End With
    
    Call get�������
    SetColValue = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub mshBill_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        With mshBill
           If .Col = mconIntColҩ�� Then
                PopupMenu mnuCol, 2
            End If
        End With
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
    ValidData = False
    Dim intLop As Integer
    
    With mshBill
        If .TextMatrix(1, 0) <> "" Then         '�����з�����
            
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
                                       
                End If
            Next
        Else
            Exit Function
        End If
    End With
    
    ValidData = True
End Function

Private Function SaveCard() As Boolean
    Dim chrNo As Variant
    Dim lngSerial As Long
    Dim lngStockid As Long
    Dim lngEnterStockID As Long
    Dim lngDrugID As Long
    Dim strBatchNo As String
    Dim lngBatchID As Long
    Dim strProducingArea As String
    Dim datTimeLimit As String
    Dim dblQuantity As Double
    Dim dblRealQuantity As Double
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
    Dim lng�ϴι�Ӧ��ID As Long
    Dim str��׼�ĺ� As String
    Dim int��� As Integer
    
    Dim intRow As Integer
    Dim arrSql As Variant
    'ҩƷ�ɹ��ƻ�
    Dim strSQLDrugPlan As String
    Dim arrSQLDrugPlanDetail As Variant
    Dim ID_IN As Long
    Dim NO_IN As Variant
    Dim arrSum As Variant
    
    '�Զ��ֽ������¼ʱʹ��
    Dim blnAuto As Boolean              '�Ƿ���Ҫ�Զ��ֽ�
    Dim rsStock As New ADODB.Recordset
    
    Dim strCheckString As String
    Dim n As Integer, intPlanSN As Integer
    Dim rsSpec As ADODB.Recordset   '������ݼ�
    Dim dbl�ͻ����� As Double
    
    SaveCard = False
    arrSql = Array()
    arrSQLDrugPlanDetail = Array()
    arrSum = Array()
    
    On Error GoTo errHandle
    
    '���¿��ÿ����������Ϊ��ʱ������ʱҪ�����
    If mbln�¿������� Then
        For n = 1 To mshBill.rows - 1
            If Val(mshBill.TextMatrix(n, 0)) <> 0 Then
                If Not CompareUsableQuantity(n, mshBill.TextMatrix(n, mconIntColʵ������)) Then
                    Exit Function
                End If
            End If
        Next
    End If
    
    With mshBill
        chrNo = Trim(txtNo)
        lngStockid = cboStock.ItemData(cboStock.ListIndex)
        If chrNo = "" Then chrNo = zlDataBase.GetNextNo(26, lngStockid)
        If IsNull(chrNo) Then Exit Function
        txtNo.Tag = chrNo
        lngEnterStockID = mlngStockID
        strBrief = Trim(txtժҪ.Text)
        strBooker = Txt������
        datBookDate = Format(zlDataBase.Currentdate, "yyyy-mm-dd hh:mm:ss")
        strAssessor = Txt�����
        
        ID_IN = zlDataBase.GetNextId("ҩƷ�ɹ��ƻ�")
        NO_IN = zlDataBase.GetNextNo(32, mlngStockID)
        
        If mint�༭״̬ = 2 Then        '�޸�
            strCheckString = CheckBill(chrNo)
            If strCheckString <> "" Then
                MsgBox strCheckString, vbInformation, gstrSysName
                Exit Function
            End If
        
            gstrSQL = "zl_ҩƷ�ƿ�_Delete('" & mstr���ݺ� & "')"
            ReDim Preserve arrSql(UBound(arrSql) + 1)
            arrSql(UBound(arrSql)) = "0;" & gstrSQL
        End If
        
        '��ҩƷID˳���������
        recSort.Sort = "ҩƷid,����,���"
        recSort.MoveFirst
        
        For n = 1 To recSort.RecordCount
            intRow = recSort!�к�
            If NVL(.TextMatrix(intRow, mconIntCol��д����), 0) <> 0 Then
                int��� = intRow 'int��� + 1
                If .TextMatrix(intRow, 0) <> "" Then
                    '�����ǰ����ҩƷ�������Զ�ȡ�������ε�ҩƷ��������������¼
                    lngDrugID = .TextMatrix(intRow, 0)
                    strProducingArea = .TextMatrix(intRow, mconIntCol����)
                    strBatchNo = .TextMatrix(intRow, mconIntCol����)
                    lngBatchID = .TextMatrix(intRow, mconIntCol����)
                    datTimeLimit = IIf(.TextMatrix(intRow, mconIntColЧ��) = "", "", .TextMatrix(intRow, mconIntColЧ��))
                    If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And datTimeLimit <> "" Then
                        '����ΪʧЧ��������
                        datTimeLimit = Format(DateAdd("D", 1, datTimeLimit), "yyyy-mm-dd")
                    End If
                    
                    dblQuantity = .TextMatrix(intRow, mconIntCol��д����) * .TextMatrix(intRow, mconIntCol����ϵ��)
                    dblRealQuantity = .TextMatrix(intRow, mconIntColʵ������) * .TextMatrix(intRow, mconIntCol����ϵ��)
'                    dblPurchasePrice = FormatEx(.TextMatrix(intRow, mconIntCol�ɹ���) / .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_�ɱ���)
                    dblPurchasePrice = Get�ɱ���(lngDrugID, lngStockid, lngBatchID)
                                        
                    dblPurchaseMoney = .TextMatrix(intRow, mconIntCol�ɹ����)
'                    dblSalePrice = FormatEx(.TextMatrix(intRow, mconIntCol�ۼ�) / .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_���ۼ�)
                    dblSalePrice = Get���ۼ�(Split(.TextMatrix(intRow, mconIntCol���Ч��), "||")(1) = 1, lngDrugID, lngStockid, lngBatchID)
                    
                    dblSaleMoney = .TextMatrix(intRow, mconIntCol�ۼ۽��)
                    dblMistakePrice = Val(.TextMatrix(intRow, mconintCol���))
                    lng�ϴι�Ӧ��ID = .TextMatrix(intRow, mconIntCol�ϴι�Ӧ��ID)
                    str��׼�ĺ� = IIf(Trim(.TextMatrix(intRow, mconIntCol��׼�ĺ�)) = "", "", .TextMatrix(intRow, mconIntCol��׼�ĺ�))
                    
'                    If Val(.TextMatrix(intRow, mconIntCol���)) = 0 Then
'                        lngSerial = 2 * int��� - 1  '����������ʽΪ��2n-1;�������Ϊż��
'                    Else
'                        lngSerial = Val(.TextMatrix(intRow, mconIntCol���))
'                    End If
                    lngSerial = 2 * int��� - 1
                    
                    gstrSQL = "zl_ҩƷ����_INSERT("
                    'NO
                    gstrSQL = gstrSQL & "'" & chrNo & "'"
                    '���
                    gstrSQL = gstrSQL & "," & lngSerial
                    '�ⷿID
                    gstrSQL = gstrSQL & "," & lngStockid
                    '�Է�����ID
                    gstrSQL = gstrSQL & "," & lngEnterStockID
                    'ҩƷID
                    gstrSQL = gstrSQL & "," & lngDrugID
                    '����
                    gstrSQL = gstrSQL & "," & lngBatchID
                    '��д����
                    gstrSQL = gstrSQL & "," & dblQuantity
                    'ʵ������
                    gstrSQL = gstrSQL & "," & dblRealQuantity
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
                    gstrSQL = gstrSQL & "," & IIf(datTimeLimit = "", "Null", "to_date('" & datTimeLimit & "','yyyy-mm-dd')")
                    'ժҪ
                    gstrSQL = gstrSQL & ",'" & strBrief & "'"
                    '��������
                    gstrSQL = gstrSQL & ",to_date('" & datBookDate & "','yyyy-mm-dd HH24:MI:SS')"
                    '��Ӧ��ID
                    gstrSQL = gstrSQL & "," & IIf(lng�ϴι�Ӧ��ID = 0, "NULL", lng�ϴι�Ӧ��ID)
                    '��׼�ĺ�
                    gstrSQL = gstrSQL & ",'" & str��׼�ĺ� & "'"
                    '���췽ʽ
                    gstrSQL = gstrSQL & "," & mintApplyType
                    '����ʱ��
                    gstrSQL = gstrSQL & ",'" & mstrEndTime & "'"
                    gstrSQL = gstrSQL & ")"
    
                    ReDim Preserve arrSql(UBound(arrSql) + 1)
                    arrSql(UBound(arrSql)) = CStr(lngDrugID) & ";" & gstrSQL
                    
                    'ҩƷ�ɹ��ƻ�����
                    If chkExportPlan.Value = 1 And chkExportPlan.Visible Then
                        If .TextMatrix(intRow, mconIntCol����ҩƷ) = "" Then .TextMatrix(intRow, mconIntCol����ҩƷ) = True
                        If .TextMatrix(intRow, mconIntCol����ҩƷ) = False Then
                            gstrSQL = "Select �ͻ���λ,�ͻ���װ From ҩƷ��� Where ҩƷid = [1]"
                            Set rsSpec = zlDataBase.OpenSQLRecord(gstrSQL, "��ѯ�ͻ���λ", lngDrugID)
                            If IsNull(rsSpec!�ͻ���λ) = False Then
                                dbl�ͻ����� = GetFormat(dblRealQuantity / rsSpec!�ͻ���װ, 1)
                            End If
                            '��������ͬҩƷID���ϲ�����
                            If CheckRepeatDrugID(recSort, n, lngDrugID) Then
                                '�ϲ�����
                                SumQuantity arrSum, lngDrugID, dblQuantity
                            Else
                                intPlanSN = intPlanSN + 1
                                gstrSQL = "zl_ҩƷ�ƻ�����α�_INSERT(" & _
                                          ID_IN & "," & _
                                          lngDrugID & "," & _
                                          intPlanSN & "," & _
                                          GetQuantity(arrSum, lngDrugID, dblQuantity) & "," & _
                                          dblPurchasePrice & "," & _
                                          dblPurchaseMoney & "," & _
                                          "null,null,0," & _
                                          IIf(lng�ϴι�Ӧ��ID <= 0, "null", "'" & GetProvider(lng�ϴι�Ӧ��ID) & "'") & "," & _
                                          IIf(strProducingArea = "", "null", "'" & strProducingArea & "'") & "," & _
                                          "null," & _
                                          dblSalePrice & "," & _
                                          dblSaleMoney & "," & _
                                          "null,null," & _
                                          dbl�ͻ����� & ")"
                                
                                ReDim Preserve arrSQLDrugPlanDetail(UBound(arrSQLDrugPlanDetail) + 1)
                                arrSQLDrugPlanDetail(UBound(arrSQLDrugPlanDetail)) = gstrSQL & ";"
                            End If
                        End If
                    End If
                End If
            End If
            recSort.MoveNext
        Next
        
        'ҩƷ�ɹ��ƻ�
        If chkExportPlan.Value = 1 And chkExportPlan.Visible Then
            strSQLDrugPlan = "zl_ҩƷ�ƻ���������_INSERT(" & _
                             ID_IN & ",'" & _
                             NO_IN & "'," & _
                             "0," & _
                             "null," & _
                             lngStockid & "," & _
                             lngEnterStockID & "," & _
                             "0,'" & _
                             strBooker & "'," & _
                             "to_date('" & datBookDate & "','yyyy-mm-dd HH24:MI:SS'),'" & _
                             "��ҩƷ�깺�����Զ����ɡ�')"
        End If
         
        If Not ExecuteSql(arrSql, strSQLDrugPlan, arrSQLDrugPlanDetail, MStrCaption) Then Exit Function
        
        mblnSave = True
        mblnSuccess = True
        mblnChange = False
    End With
    SaveCard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function SaveCheck() As Boolean
    Dim rsTemp As New Recordset
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
    On Error GoTo errHandle
    '���õ����Ƿ��ڽ���༭����󣬱���������Ա�޸�
    mstrTime_End = GetBillInfo(6, mstr���ݺ�)
    If mstrTime_End = "" Then
        MsgBox "�õ����Ѿ�����������Աɾ����", vbInformation, gstrSysName
        Exit Function
    End If
    If mstrTime_End > mstrTime_Start Then
        MsgBox "�õ����Ѿ�����������Ա�༭�����˳������ԣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    '���õ����Ƿ���������
    gstrSQL = " Select ��ҩ���� From ҩƷ�շ���¼ " & _
            " Where ����=6 And NO=[1] And Rownum<2"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[���õ����Ƿ���������]", Me.txtNo.Tag)
    
    If IsNull(rsTemp!��ҩ����) Then
        MsgBox "�õ��ݱ���������Աȡ�����ͣ���������գ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    lng�ⷿID = cboStock.ItemData(cboStock.ListIndex)
    lng�Է�����id = mlngStockID
    str����� = gstrUserName
    strNo = txtNo.Tag
    
    gstrSQL = "SELECT b.ϵ��,b.id AS ���id " _
            & " FROM ҩƷ�������� a, ҩƷ������ b " _
            & "Where a.���id = b.ID " _
            & "  AND a.���� = 6 "
    
    Call SQLTest(App.Title, "ҩƷ�ƿ����", gstrSQL)
    If rsTemp.State = 1 Then rsTemp.Close
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "SaveCheck")
    Call SQLTest
    
    If rsTemp.EOF Then
        MsgBox "�Բ���ҩƷ������಻ȫ������!", vbExclamation + vbOKOnly, gstrSysName
        Exit Function
    End If
    If rsTemp.RecordCount < 2 Then
        MsgBox "�Բ���ҩƷ������಻ȫ������!", vbExclamation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    rsTemp.MoveFirst
    Do While Not rsTemp.EOF
        If rsTemp!ϵ�� = 1 Then
            lng�����id = rsTemp!���id
        Else
            lng�����id = rsTemp!���id
        End If
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    
    If mblnUpdate = False Then
        If Not ��鵥��(6, txtNo.Tag, True, True) Then
            Call RefreshBill
            mblnUpdate = True
            Exit Function
        End If
    End If
    
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
                
                If Val(.TextMatrix(intRow, mconIntCol��д����)) = Val(.TextMatrix(intRow, mconIntColʵ������)) Then
                    num��д���� = Val(.TextMatrix(intRow, mconIntColԭʼ����))
                    numʵ������ = Val(.TextMatrix(intRow, mconIntColԭʼ����))
                Else
                    num��д���� = GetFormat(Val(.TextMatrix(intRow, mconIntCol��д����)) * .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_����)
                    numʵ������ = GetFormat(Val(.TextMatrix(intRow, mconIntColʵ������)) * .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_����)
                End If
                
'                num�ɱ��� = GetFormat(Val(.TextMatrix(intRow, mconIntCol�ɹ���)) / .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_�ɱ���)
                num�ɱ��� = Get�ɱ���(lngҩƷID, lng�ⷿID, lng������)
                num�ɱ���� = Val(.TextMatrix(intRow, mconIntCol�ɹ����))
'                dbl�ۼ� = GetFormat(Val(.TextMatrix(intRow, mconIntCol�ۼ�)) / .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_���ۼ�)
                dbl�ۼ� = Get���ۼ�(Split(.TextMatrix(intRow, mconIntCol���Ч��), "||")(1) = 1, lngҩƷID, lng�ⷿID, lng������)
                num���۽�� = Val(.TextMatrix(intRow, mconIntCol�ۼ۽��))
                num��� = Val(.TextMatrix(intRow, mconintCol���))
                str���� = .TextMatrix(intRow, mconIntCol����)
                datЧ�� = IIf(.TextMatrix(intRow, mconIntColЧ��) = "", "", .TextMatrix(intRow, mconIntColЧ��))
                If gtype_UserSysParms.P149_Ч����ʾ��ʽ = 1 And datЧ�� <> "" Then
                    '����ΪʧЧ��������
                    datЧ�� = Format(DateAdd("D", 1, datЧ��), "yyyy-mm-dd")
                End If
                                
                int���к� = Val(.TextMatrix(intRow, mconIntCol���))
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
                '��Ӧ��ID
                gstrSQL = gstrSQL & "," & IIf(lng�ϴι�Ӧ��ID = 0, "NULL", lng�ϴι�Ӧ��ID)
                '��׼�ĺ�
                gstrSQL = gstrSQL & ",'" & str��׼�ĺ� & "'"
                '���ۼ�
                gstrSQL = gstrSQL & "," & dbl�ۼ�
                gstrSQL = gstrSQL & ")"

                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = CStr(lngҩƷID) & ";" & gstrSQL
            End If
            recSort.MoveNext
        Next
    End With
    
'    gcnOracle.BeginTrans
    If Not ExecuteSql(arrSql, "", "", MStrCaption) Then
'        gcnOracle.RollbackTrans
        Exit Function
    End If
'    gcnOracle.CommitTrans
    
    SaveCheck = True
    mblnSave = True
    mblnSuccess = True
    mblnChange = False
    
    '��ҹ���
    Call zlPlugIn_Ini(glngSys, glngModul, mobjPlugIn)
    Call CallPlugInDrugStuffWork(mobjPlugIn, 3, lng�ⷿID, strNo, ���ݺ�.ҩƷ�ƿ�)
    
    Exit Function
errHandle:
'    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub RefreshBill()
    '�����¼۸����µ���������ݣ����ڵ������ʱ
    Dim lngRow As Long, lngRows As Long, lngҩƷID As Long
    Dim Dbl���� As Double, dbl�ɱ��� As Double, dbl�ɱ���� As Double, dbl���ۼ� As Double, dbl���۽�� As Double, dbl��� As Double
    Dim rsPrice As New ADODB.Recordset
    Dim rsStock As ADODB.Recordset
    Dim blnAdj As Boolean
    Dim intCostDigit As Integer
    Dim intPricedigit As Integer
            
    On Error GoTo errHandle
    intPricedigit = GetDigit(0, 1, 2, 1)
    intCostDigit = GetDigit(0, 1, 1, 1)
        
    gstrSQL = " Select '�ۼ�' As ����, a.���, a.ҩƷid ,nvl(a.����,0) as ����, 0 ԭ��, b.�ּ�" & _
            " From ҩƷ�շ���¼ A," & _
                 " (Select �շ�ϸĿid, Nvl(�ּ�, 0) �ּ�, ִ������" & _
                   " From �շѼ�Ŀ" & _
                   " Where (��ֹ���� Is Null Or Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'yyyy-MM-dd')))) B, �շ���ĿĿ¼ C" & _
            " Where a.���� = 6 And a.No = [1] And a.ҩƷid = b.�շ�ϸĿid And c.Id = b.�շ�ϸĿid And Round(a.���ۼ�," & intPricedigit & ") <> Round(b.�ּ�, " & intPricedigit & ") And" & _
              "    NVL(c.�Ƿ���, 0) = 0  and b.ִ������>a.��������" & _
            " Union All" & _
            " Select '�ۼ�' As ����, a.���, a.ҩƷid ,nvl(a.����,0) as ����, 0 ԭ��, decode(nvl(b.����,0),0,b.ʵ�ʽ�� / b.ʵ������,b.���ۼ�) As �ּ�" & _
            " From ҩƷ�շ���¼ A, ҩƷ��� B, �շ���ĿĿ¼ C" & _
            " Where a.���� = 6 And a.No = [1] And c.Id = a.ҩƷid And Round(a.���ۼ�," & intPricedigit & ") <> Round(decode(nvl(b.����,0),0,b.ʵ�ʽ�� / b.ʵ������,b.���ۼ�), " & intPricedigit & ") And Nvl(c.�Ƿ���, 0) = 1 And" & _
                  " b.���� = 1 And b.�ⷿid = a.�ⷿid And b.ҩƷid = a.ҩƷid And NVL(b.����, 0) = NVL(a.����, 0) And NVL(b.ʵ������, 0) <> 0 And a.���ϵ�� = -1" & _
            " Union All" & _
            " Select '�ɱ���' As ����, a.���, a.ҩƷid ,nvl(a.����,0) as ����, 0 ԭ��, b.ƽ���ɱ��� As �ּ�" & _
            " From ҩƷ�շ���¼ A, ҩƷ��� B" & _
            " Where a.���� = 6 And a.No = [1] And a.ҩƷid = b.ҩƷid And Nvl(a.����, 0) = Nvl(b.����, 0) and round(a.�ɱ���," & intCostDigit & ")<>round(b.ƽ���ɱ���," & intCostDigit & ") And a.�ⷿid = b.�ⷿid and a.���ϵ��=-1 and b.����=1" & _
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
                dbl���ۼ� = Val(GetFormat(rsPrice!�ּ� * Val(mshBill.TextMatrix(lngRow, mconIntCol����ϵ��)), intPricedigit))
                dbl���۽�� = Val(GetFormat(dbl���ۼ� * Dbl����, mintMoneyDigit))
                dbl��� = Val(GetFormat(dbl���۽�� - dbl�ɱ����, mintMoneyDigit))
            End If
            
            rsPrice.Filter = "����='�ɱ���' And ҩƷID=" & lngҩƷID & " And ����=" & Val(mshBill.TextMatrix(lngRow, mconIntCol����))
            If rsPrice.RecordCount > 0 Then
                blnAdj = True
                dbl���۽�� = Val(GetFormat(dbl���ۼ� * Dbl����, mintMoneyDigit))
                dbl�ɱ��� = Val(GetFormat(rsPrice!�ּ� * Val(mshBill.TextMatrix(lngRow, mconIntCol����ϵ��)), intCostDigit))
                dbl�ɱ���� = Val(GetFormat(dbl�ɱ��� * Dbl����, mintMoneyDigit))
                dbl��� = Val(GetFormat(dbl���۽�� - dbl�ɱ����, mintMoneyDigit))
            End If
            
            If blnAdj = True Then
                '�Ե�ǰ���¼۸����µ���������ݣ��ۼۡ��ɱ��ۡ����۽��ɱ�����ۣ�
                mshBill.TextMatrix(lngRow, mconIntCol�ۼ�) = GetFormat(dbl���ۼ�, intPricedigit)
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
    Dim strҩƷ As String
    
    arrSql = Array()
    SaveStrike = False
    
    With mshBill
        '����������������С����
        For intRow = 1 To .rows - 1
            If Val(.TextMatrix(intRow, mconIntColʵ������)) <> 0 Then
                If Not ��ͬ����(Val(.TextMatrix(intRow, mconIntCol��д����)), Val(.TextMatrix(intRow, mconIntColʵ������))) Then
                    MsgBox "������Ϸ��ĳ�����������" & intRow & "�У���", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            
            '�����������Ƿ��㹻����������Ϊ�������ʱ������
            If mint��������ⷿ <> 0 And .TextMatrix(intRow, 0) <> "" Then
                If .TextMatrix(intRow, mconIntColʵ������) = .TextMatrix(intRow, mconIntCol��д����) Then
                    ��������_IN = .TextMatrix(intRow, mconintCol��ʵ����)
                Else
                    ��������_IN = GetFormat(.TextMatrix(intRow, mconIntColʵ������) * .TextMatrix(intRow, mconIntCol����ϵ��), gtype_UserDrugDigits.Digit_����)
                End If
                
                If CheckStrickUsable(6, mlngStockID, Val(.TextMatrix(intRow, 0)), .TextMatrix(intRow, mconIntColҩ��), _
                    Val(.TextMatrix(intRow, mconIntCol����)), Val(��������_IN), mint��������ⷿ, Trim(txtNo.Tag), Val(.TextMatrix(intRow, mconIntCol���)) + 1) = False Then
                    .Row = intRow
                    .MsfObj.TopRow = intRow
                    Exit Function
                End If
            End If
        Next
        
        '��ͨ�������ʵ������
        If mint�༭״̬ = 7 And mint����ʽ = 0 Then
            strҩƷ = CheckNumStock(mshBill, mlngStockID, 0, mconIntCol����, mconIntColʵ������, mconIntCol����ϵ��, 2, 0, mconintCol��ʵ����)
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
        ������_IN = gstrUserName
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
        If strҩƷID <> "" And mint����ʽ = 0 Then
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

Private Sub ��ʾ�����()
    Dim rsUseCount As New Recordset
    Dim dblStock As Double
    Dim int�����ⷿ�������� As Integer
    Dim int���տⷿ�������� As Integer
    Dim int�������� As Integer
    
    On Error GoTo errHandle
    If mint�༭״̬ = 1 Or mint�༭״̬ = 2 Or mint�༭״̬ = 5 Then
        Exit Sub
    Else
        With mshBill
            int�������� = 0
        
            'ȡ�����ͽ��տⷿ�ķ�������
            '�����ⷿ
            gstrSQL = "Select a.ҩ�����,a.ҩ������,b.�������� " & _
                " From ҩƷ��� a,��������˵�� b " & _
                " Where a.ҩƷid = [2] And b.����id = [1] "
            Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[�����ⷿ��������]", cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, 0)))
            
            Do While Not rsUseCount.EOF
                If int�������� <> 1 Then
                    If InStr(rsUseCount!��������, "ҩ��") > 0 Or rsUseCount!�������� = "�Ƽ���" Then
                        int�������� = 1
                        int�����ⷿ�������� = rsUseCount!ҩ������
                    ElseIf InStr(rsUseCount!��������, "ҩ��") > 0 Then
                        int�������� = 2
                        int�����ⷿ�������� = rsUseCount!ҩ�����
                    End If
                End If
                rsUseCount.MoveNext
            Loop
            
            int�������� = 0
            '���տⷿ
            gstrSQL = "Select a.ҩ�����,a.ҩ������,b.�������� " & _
                " From ҩƷ��� a,��������˵�� b " & _
                " Where a.ҩƷid = [2] And b.����id = [1]"
            Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[���տⷿ��������]", mlngStockID, Val(.TextMatrix(.Row, 0)))
            
            Do While Not rsUseCount.EOF
                If int�������� <> 1 Then
                    If InStr(rsUseCount!��������, "ҩ��") > 0 Or rsUseCount!�������� = "�Ƽ���" Then
                        int�������� = 1
                        int���տⷿ�������� = rsUseCount!ҩ������
                    ElseIf InStr(rsUseCount!��������, "ҩ��") > 0 Then
                        int�������� = 2
                        int���տⷿ�������� = rsUseCount!ҩ�����
                    End If
                End If
                rsUseCount.MoveNext
            Loop
        
            
            If .TextMatrix(.Row, mconIntColҩ��) = "" Then
                staThis.Panels(2).Text = ""
                Exit Sub
            End If
            If .TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
            '�������ĵ�ǰҩƷ�Ŀ���������ͳ�����������������ܵ�����
            If mint��ȷ���� = 1 And int�����ⷿ�������� = 1 Then
                gstrSQL = " Select ��������/" & .TextMatrix(.Row, mconIntCol����ϵ��) & " as �������� from ҩƷ��� " & _
                          " Where �ⷿid=[1] " & _
                          " And ҩƷid=[2] And ����=1 " & _
                          " And Nvl(����,0)=[3]"
            Else
                gstrSQL = " Select Sum(��������)/" & .TextMatrix(.Row, mconIntCol����ϵ��) & " as �������� from ҩƷ��� " & _
                          " Where �ⷿid=[1] " & _
                          " And ҩƷid=[2] And ����=1 "
            End If
            Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[�����ⷿ��������]", cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(.Row, 0)), Val(.TextMatrix(.Row, mconIntCol����)))
            
            If rsUseCount.EOF Then
                .TextMatrix(.Row, mconIntCol��������) = 0
            Else
                .TextMatrix(.Row, mconIntCol��������) = IIf(IsNull(rsUseCount.Fields(0)), 0, rsUseCount.Fields(0))
            End If
            rsUseCount.Close
            
            '��ǰҩ���Ŀ�����������ͳ����������
            gstrSQL = " Select Sum(��������/" & .TextMatrix(.Row, mconIntCol����ϵ��) & ") as �������� from ҩƷ��� where �ⷿid=[1] " & _
                      " And ҩƷid=[2] And ����=1 "
            Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ǰҩ����������]", mlngStockID, Val(.TextMatrix(.Row, 0)))
            
            If rsUseCount.EOF Then
                dblStock = 0
            Else
                dblStock = IIf(IsNull(rsUseCount.Fields(0)), 0, rsUseCount.Fields(0))
            End If
            
            Dim blnIs��ʾ�Է���� As Boolean
            Dim str�Է������ As String
            
            blnIs��ʾ�Է���� = IsHavePrivs(mstrPrivs, "��ʾ�Է����")
            str�Է������ = "��" & Me.cboStock.Text & "�����Ϊ[" & GetFormat(.TextMatrix(.Row, mconIntCol��������), mintNumberDigit) & "]" & .TextMatrix(.Row, mconIntCol��λ)
            
            staThis.Panels(2).Text = "��ҩƷ" & frmRequestDrugList.cboStock.Text & "�����Ϊ[" & GetFormat(dblStock, mintNumberDigit) & "]" & .TextMatrix(.Row, mconIntCol��λ) _
                & IIf(blnIs��ʾ�Է����, str�Է������, "")
        End With
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
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

'ת����ֵΪ����
Private Function TranNumToDate(ByVal strNum As Long) As String
    Dim strYear As String
    Dim strMonth As String
    Dim strDay As String
    Dim strDate As String
    
    TranNumToDate = ""
    strYear = Mid(strNum, 1, 4)
    strMonth = Mid(strNum, 5, 2)
    strDay = Mid(strNum, 7, 2)
        
    If strYear < 2000 Or strYear > 5000 Then Exit Function
    If strMonth = "" Then strMonth = "01"
    If strDay = "" Then strDay = "01"
    
    If strMonth > 12 Or strMonth < 1 Then Exit Function
    strDate = strYear & "-" & strMonth & "-" & strDay
        
    If Not IsDate(strDate) Then Exit Function
    
    strDate = Format(strDate, "yyyy-mm-dd")
    TranNumToDate = strDate
    
End Function

'������������бȽ�
Private Function CompareUsableQuantity(ByVal intRow As Integer, ByVal dbl��д���� As Double) As Boolean
    Dim dblUsableQuantity As Double      'ʵ��������Ӧ���������
    Dim numUsedCount As Double
    Dim vardrug As Variant
    Dim lng�ⷿID As Long
    Dim lngҩƷID As Long
    Dim lng���� As Long
    Dim dblԭ��д���� As Double
    Dim rsCheck As New ADODB.Recordset
    
    On Error GoTo errHandle
    lng�ⷿID = cboStock.ItemData(cboStock.ListIndex)
        
    'mint�����: 0-�����;1-��飬�������ѣ�2-��飬�����ֹ
    'ֻҪ�Ƿ���ҩƷ����������ȵ�ǰ���δ�������������Զ��ֽ⣬��������ʱ��ҩƷ���ԵĲ�����
    CompareUsableQuantity = False
    If mint��ȷ���� = 0 Then CompareUsableQuantity = True: Exit Function
    
    With mshBill
        If .TextMatrix(intRow, 0) = "" Then Exit Function
        
        lngҩƷID = .TextMatrix(intRow, 0)
        lng���� = .TextMatrix(intRow, mconIntCol����)
        
        gstrSQL = "Select Nvl(Sum(��������),0) �������� From ҩƷ��� Where �ⷿID=[1] And Nvl(����,0)=[3] And ����=1 And ҩƷID=[2] "
        Set rsCheck = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[������Ƿ��㹻]", lng�ⷿID, lngҩƷID, lng����)
                
        dblUsableQuantity = rsCheck!�������� / .TextMatrix(intRow, mconIntCol����ϵ��)
        
        If .TextMatrix(intRow, mconIntCol����) > 0 Or Split(.TextMatrix(intRow, mconIntCol���Ч��), "||")(1) = 1 Then     '���Ƴ��ⷿ�ǿⷿ��ҩƷ�Ƿ��������ҩƷ��ʱ��ҩƷ���ж�
            If mint�༭״̬ = 1 Or mint�༭״̬ = 5 Then
                If dbl��д���� > dblUsableQuantity Then
                    MsgBox "��" & intRow & "��ҩƷ" & .TextMatrix(intRow, mconIntColҩ��) & "��" & vbCrLf & "�Բ����������������" & dbl��д���� & "�������˸�ҩƷ�Ŀ��ÿ��������" & dblUsableQuantity & "���������䣡", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            ElseIf mint�༭״̬ = 2 Then
                numUsedCount = 0
                For Each vardrug In mcolUsedCount
                    If vardrug(0) = .TextMatrix(intRow, 0) & .TextMatrix(intRow, mconIntCol����) Then
                        numUsedCount = vardrug(1)
                        Exit For
                    End If
                Next
                
                dblԭ��д���� = IIf(mbln�¿�������, numUsedCount, 0)
                
                If dbl��д���� > dblUsableQuantity + dblԭ��д���� Then
                    MsgBox "��" & intRow & "��ҩƷ" & .TextMatrix(intRow, mconIntColҩ��) & "��" & vbCrLf & "�Բ����������������" & dbl��д���� & "�������˸�ҩƷ�Ŀ��ÿ��������" & dblUsableQuantity + dblԭ��д���� & "���������䣡", vbExclamation + vbOKOnly, gstrSysName
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
                    If MsgBox("��" & intRow & "��ҩƷ" & .TextMatrix(intRow, mconIntColҩ��) & "��" & vbCrLf & "�Բ����������������" & dbl��д���� & "�������˸�ҩƷ�Ŀ��ÿ��������" & dblUsableQuantity & "�����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                End If
            ElseIf mint�༭״̬ = 2 Then
                numUsedCount = 0
                For Each vardrug In mcolUsedCount
                    If vardrug(0) = .TextMatrix(intRow, 0) & .TextMatrix(intRow, mconIntCol����) Then
                        numUsedCount = vardrug(1)
                        Exit For
                    End If
                Next
                
                dblԭ��д���� = IIf(mbln�¿�������, numUsedCount, 0)
                
                If dbl��д���� > dblUsableQuantity + dblԭ��д���� Then
                    If MsgBox("��" & intRow & "��ҩƷ" & .TextMatrix(intRow, mconIntColҩ��) & "��" & vbCrLf & "�Բ����������������" & dbl��д���� & "�������˸�ҩƷ�Ŀ��ÿ��������" & dblUsableQuantity + dblԭ��д���� & "�����Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                End If
            End If
            
        ElseIf mint����� = 2 Then
            '2-��飬�����ֹ
            If mint�༭״̬ = 1 Then
                If dbl��д���� > dblUsableQuantity Then
                    MsgBox "��" & intRow & "��ҩƷ" & .TextMatrix(intRow, mconIntColҩ��) & ":" & vbCrLf & "�Բ����������������" & dbl��д���� & "�������˸�ҩƷ�Ŀ��ÿ��������" & dblUsableQuantity & "���������䣡", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            ElseIf mint�༭״̬ = 2 Then
                numUsedCount = 0
                For Each vardrug In mcolUsedCount
                    If vardrug(0) = .TextMatrix(intRow, 0) & .TextMatrix(intRow, mconIntCol����) Then
                        numUsedCount = vardrug(1)
                        Exit For
                    End If
                Next
                
                If dbl��д���� > dblUsableQuantity + numUsedCount Then
                    MsgBox "��" & intRow & "��ҩƷ" & .TextMatrix(intRow, mconIntColҩ��) & "��" & vbCrLf & "�Բ����������������" & dbl��д���� & "�������˸�ҩƷ�Ŀ��ÿ��������" & dblUsableQuantity + numUsedCount & "���������䣡", vbExclamation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            End If
        End If
            
    End With
    
    CompareUsableQuantity = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Function ExecuteSql(ByRef arrSql As Variant, ByVal strSQLDrugPlan As String _
    , ByRef arrSQLDrugPlanDetail As Variant, strTitle As String, Optional ByVal blnǿ�Ʊ��� As Boolean = False) As Boolean
    Dim strTmp As Variant
    Dim i As Integer, j As Integer

    ExecuteSql = False
    If UBound(arrSql) >= 0 Then
        '��SQL���а�ҩƷID��������
        For i = 0 To UBound(arrSql) - 1
            For j = i + 1 To UBound(arrSql)
                If CLng(Split(arrSql(j), ";")(0)) < CLng(Split(arrSql(i), ";")(0)) Then
                    strTmp = CStr(arrSql(j))
                    arrSql(j) = arrSql(i)
                    arrSql(i) = strTmp
                End If
            Next
        Next
        
        'ִ��SQL���
        On Error GoTo errH
        If Not blnǿ�Ʊ��� Then gcnOracle.BeginTrans
        For i = 0 To UBound(arrSql)
            Call zlDataBase.ExecuteProcedure(CStr(Mid(arrSql(i), InStr(arrSql(i), ";") + 1)), strTitle)
        Next
        'ҩƷ�ɹ��ƻ�
        If Trim(strSQLDrugPlan) <> "" Then
            If UBound(arrSQLDrugPlanDetail) >= 0 Then
                Call zlDataBase.ExecuteProcedure(strSQLDrugPlan, strTitle & "-�ɹ��ƻ�")
                For i = 0 To UBound(arrSQLDrugPlanDetail)
                    Call zlDataBase.ExecuteProcedure(CStr(Split(arrSQLDrugPlanDetail(i), ";")(0)), strTitle & "-�ɹ��ƻ�����")
                Next
            End If
        End If
        
        If Not blnǿ�Ʊ��� Then gcnOracle.CommitTrans
        ExecuteSql = True
    End If
    Exit Function
errH:
    If Not blnǿ�Ʊ��� Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'��ӡ����
Private Sub printbill()
    Dim int��λϵ�� As Integer
    
    With mshBill
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
        FrmBillPrint.ShowMe Me, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1304", "zl8_bill_1304"), mint��¼״̬, int��λϵ��, 1304, "ҩƷ���쵥", txtNo.Tag
    End With
End Sub


Private Sub get�������()
'''''''''''''''''''''''''''''''''''''
'��ȡ��������ķ���
'''''''''''''''''''''''''''''''''''''
    Dim rsUseCount As New Recordset
    Dim dblStock As Double
    Dim int�����ⷿ�������� As Integer
    Dim int���տⷿ�������� As Integer
    Dim int�������� As Integer '��ȡ�ⷿ�Ĺ������ʣ���ҩ�⻹��ҩ��
    Dim blnIs��ʾ�Է���� As Boolean
    Dim str�Է������ As String
    Dim i As Integer
    
    On Error GoTo errHandle
    With mshBill
        For i = 1 To .rows - 1
            If .TextMatrix(i, 0) = "" Then Exit Sub
            int�������� = 0
        
            'ȡ�����ͽ��տⷿ�ķ�������
            '�����ⷿ
            gstrSQL = "Select a.ҩ�����,a.ҩ������,b.�������� " & _
                " From ҩƷ��� a,��������˵�� b " & _
                " Where a.ҩƷid = [2] And b.����id = [1] "
            Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[�����ⷿ��������]", cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(i, 0)))
            
            Do While Not rsUseCount.EOF
                If int�������� <> 1 Then
                    If InStr(rsUseCount!��������, "ҩ��") > 0 Or rsUseCount!�������� = "�Ƽ���" Then
                        int�������� = 1
                        int�����ⷿ�������� = rsUseCount!ҩ������
                    ElseIf InStr(rsUseCount!��������, "ҩ��") > 0 Then
                        int�������� = 2
                        int�����ⷿ�������� = rsUseCount!ҩ�����
                    End If
                End If
                rsUseCount.MoveNext
            Loop
            
            int�������� = 0
            '���տⷿ
            gstrSQL = "Select a.ҩ�����,a.ҩ������,b.�������� " & _
                " From ҩƷ��� a,��������˵�� b " & _
                " Where a.ҩƷid = [2] And b.����id = [1]"
            Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[���տⷿ��������]", mlngStockID, Val(.TextMatrix(i, 0)))
            
            Do While Not rsUseCount.EOF
                If int�������� <> 1 Then
                    If InStr(rsUseCount!��������, "ҩ��") > 0 Or rsUseCount!�������� = "�Ƽ���" Then
                        int�������� = 1
                        int���տⷿ�������� = rsUseCount!ҩ������
                    ElseIf InStr(rsUseCount!��������, "ҩ��") > 0 Then
                        int�������� = 2
                        int���տⷿ�������� = rsUseCount!ҩ�����
                    End If
                End If
                rsUseCount.MoveNext
            Loop
            
            If .TextMatrix(mshBill.Row, 0) = "" Then Exit Sub
            blnIs��ʾ�Է���� = IsHavePrivs(mstrPrivs, "��ʾ�Է����")
            
            If blnIs��ʾ�Է���� Then
                '�������ĵ�ǰҩƷ�Ŀ���������ͳ�����������������ܵ�����
                If mint��ȷ���� = 1 And int�����ⷿ�������� = 1 Then
                    gstrSQL = " Select ��������/" & .TextMatrix(i, mconIntCol����ϵ��) & " as �������� from ҩƷ��� " & _
                              " Where �ⷿid=[1] " & _
                              " And ҩƷid=[2] And ����=1 " & _
                              " And Nvl(����,0)=[3]"
                Else
                    gstrSQL = " Select Sum(��������)/" & .TextMatrix(i, mconIntCol����ϵ��) & " as �������� from ҩƷ��� " & _
                              " Where �ⷿid=[1] " & _
                              " And ҩƷid=[2] And ����=1 "
                End If
                Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[�����ⷿ��������]", cboStock.ItemData(cboStock.ListIndex), Val(.TextMatrix(i, 0)), Val(.TextMatrix(i, mconIntCol����)))
                
                If rsUseCount.EOF Then
                    .TextMatrix(i, mconIntCol��������) = 0
                Else
                    .TextMatrix(i, mconIntCol��������) = IIf(IsNull(rsUseCount.Fields(0)), 0, rsUseCount.Fields(0))
                End If
                .TextMatrix(i, mconintcol�Է����) = GetFormat(.TextMatrix(i, mconIntCol��������), mintNumberDigit)
                rsUseCount.Close
            End If
                
            '��ǰҩ���Ŀ�����������ͳ����������
            gstrSQL = " Select Sum(��������/" & .TextMatrix(i, mconIntCol����ϵ��) & ") as �������� from ҩƷ��� where �ⷿid=[1] " & _
                      " And ҩƷid=[2] And ����=1 "
            Set rsUseCount = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption & "[��ǰҩ����������]", mlngStockID, Val(.TextMatrix(i, 0)))
            
            If rsUseCount.EOF Then
                dblStock = 0
            Else
                dblStock = IIf(IsNull(rsUseCount.Fields(0)), 0, rsUseCount.Fields(0))
            End If
            .TextMatrix(i, mconintcol��ǰ���) = GetFormat(dblStock, mintNumberDigit)
       Next
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetProvider(ByVal lngProviderID As Long) As String
    Dim rsTemp As ADODB.Recordset
    
    If lngProviderID <= 0 Then Exit Function
    On Error GoTo errHandle
    gstrSQL = "select ���� from ��Ӧ�� where id=[1]"
    Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, "��Ӧ������", lngProviderID)
    If Not rsTemp.EOF Then
        GetProvider = NVL(rsTemp!����)
    End If
    rsTemp.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

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
    
    On Error GoTo errHandle
    rsTemp.MoveFirst
    str���� = ""
    strTemp = ""
    Do While Not rsTemp.EOF
        If mint��ȷ���� = 1 Then
            str���� = IIf(IsNull(rsTemp!����), "0", rsTemp!����)
        Else
            str���� = "0"
        End If
        If InStr(1, strTemp, rsTemp!ҩƷid & "," & str����) = 0 Then
            strTemp = strTemp & rsTemp!ҩƷid & "," & str���� & "," & rsTemp!ͨ���� & "|"
        End If
        
        If rsTemp!ʱ�� = 1 Then '��ʱ���޿��ļ�¼�ҳ���
            gstrSQL = "select Decode(Nvl(����,0),0,ʵ�ʽ��/ʵ������,Nvl(���ۼ�,ʵ�ʽ��/ʵ������))*" & Choose(mintUnit, 1, rsTemp!�����װ, rsTemp!סԺ��װ, rsTemp!ҩ���װ) & " as  �ۼ� " _
                & "  from ҩƷ��� " _
                & " where �ⷿid=[1] " _
                & " and ҩƷid=[2] " _
                & " and ����=1 and ʵ������>0 and " _
                & " nvl(����,0)=[3]"
            Set rsPrice = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, cboStock.ItemData(cboStock.ListIndex), rsTemp!ҩƷid, IIf(IsNull(rsTemp!����), 0, rsTemp!����))
            If rsPrice.EOF Then
                If mint��ȷ���� = 1 Then
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

Private Function Get�۸�(ByVal lngҩƷID As Long, ByVal lng���� As Long, ByVal dbl����ϵ�� As Double) As Double
    Dim rsPrice As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "select Decode(Nvl(����,0),0,ʵ�ʽ��/ʵ������,Nvl(���ۼ�,ʵ�ʽ��/ʵ������))*" & dbl����ϵ�� & " as  �ۼ� " _
        & "  from ҩƷ��� " _
        & " where �ⷿid=[1] " _
        & " and ҩƷid=[2] " _
        & " and ����=1 and ʵ������>0 and " _
        & " nvl(����,0)=[3]"
    Set rsPrice = zlDataBase.OpenSQLRecord(gstrSQL, MStrCaption, cboStock.ItemData(cboStock.ListIndex), lngҩƷID, lng����)

    If rsPrice.EOF Then
        If mint��ȷ���� = 1 Then
            Get�۸� = 0
            Exit Function
        End If
    Else
        Get�۸� = rsPrice.Fields(0).Value
    End If
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckRepeatDrugID(ByVal rsTemp As ADODB.Recordset, ByVal intRecEnd As Integer, ByVal lngDrugID As Long) As Boolean
'----------------------
'���ܣ����¼���ظ�ҩƷ
'----------------------
    Dim i As Integer
    Dim rsClone As ADODB.Recordset
    
    CheckRepeatDrugID = False
    Set rsClone = rsTemp.Clone
    With rsClone
        .Sort = "ҩƷid,����,���"
        .MoveFirst
        For i = 1 To .RecordCount
            If i > intRecEnd Then
                If lngDrugID = !ҩƷid Then
                    CheckRepeatDrugID = True
                    Exit Function
                End If
            End If
            .MoveNext
        Next
    End With

End Function

Private Sub SumQuantity(ByRef arrVal As Variant, ByVal lngDrugID As Long, ByVal dblQTY As Double)
'------------------------
'���ܣ�����ͬҩƷID������
'------------------------
    Dim i As Integer
    Dim blnFind As Boolean
    
    If UBound(arrVal) > 0 Then
        For i = 0 To UBound(arrVal, 2) - 1
            If arrVal(0, i) = lngDrugID Then
                arrVal(1, i) = arrVal(1, i) + dblQTY
                blnFind = True
                Exit For
            End If
        Next
    Else
        ReDim arrVal(2, 1)
        arrVal(0, 0) = lngDrugID
        arrVal(1, 0) = dblQTY
        blnFind = True
    End If
    If blnFind = False Then
        ReDim Preserve arrVal(2, UBound(arrVal) + 1)
        arrVal(0, UBound(arrVal)) = lngDrugID
        arrVal(1, UBound(arrVal)) = dblQTY
    End If
End Sub

Private Function GetQuantity(ByVal arrVal As Variant, ByVal lngDrugID As Long, ByVal dblQTY As Double) As Double
'----------------------------
'���ܣ���ȡ������ҩƷID������
'----------------------------
    If UBound(arrVal) > 0 Then
        Dim i As Integer
        For i = 0 To UBound(arrVal, 2) - 1
            If arrVal(0, i) = lngDrugID Then
                GetQuantity = arrVal(1, i) + dblQTY
                Exit Function
            End If
        Next
    End If
    GetQuantity = dblQTY
End Function


Private Function ���۸�() As Boolean
    '���ܣ�����ʱ���ж�ҩƷ�Ƿ������¼۸񣬲������޸ĺ���ʾ
    Dim strMsg As String '������ʾ��Ϣ
    Dim i As Integer, intSum As Integer, intPricedigit As Integer
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
        
        If intSum > 0 Then '����0��ʾ�м۸����
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
                If mint�༭״̬ = 6 Then   '�����ʵ�������ж�
                    dblʵ������ = Val(GetFormat(NVL(rsCheck!ʵ������, 0) / dbl����ϵ��, mintNumberDigit))
                Else    '�����ʵ�������ж�
                    dblʵ������ = Val(GetFormat(NVL(rsCheck!��������, 0) / dbl����ϵ��, mintNumberDigit))
                End If
            End If
            
            '�������ʵ����������
            If Not (dblʵ������ >= dbl��д����) Then
                int����� = mint�����
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
                    If MsgBox(rsProperty!ͨ���� & "�Ŀ�治�㣬�Ƿ������" & strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                Case 2
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

