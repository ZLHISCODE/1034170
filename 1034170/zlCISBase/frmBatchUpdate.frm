VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBatchUpdate 
   Caption         =   "�����޸Ĺ��"
   ClientHeight    =   5550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6885
   Icon            =   "frmBatchUpdate.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5550
   ScaleWidth      =   6885
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtFind 
      Height          =   300
      Left            =   3000
      TabIndex        =   10
      Top             =   4680
      Width           =   1455
   End
   Begin VB.PictureBox picList 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   2520
      ScaleHeight     =   2415
      ScaleWidth      =   3855
      TabIndex        =   5
      Top             =   2160
      Width           =   3855
      Begin VB.Frame fraSplit 
         Height          =   50
         Left            =   -120
         TabIndex        =   7
         Top             =   1440
         Width           =   3855
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfOtherName 
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   3615
         _cx             =   6376
         _cy             =   873
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
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
      Begin VSFlex8Ctl.VSFlexGrid vsfDetails 
         Height          =   975
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   3375
         _cx             =   5953
         _cy             =   1720
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
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
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
   Begin VB.PictureBox picDetails 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   2520
      ScaleHeight     =   1815
      ScaleWidth      =   3495
      TabIndex        =   2
      Top             =   120
      Width           =   3495
      Begin XtremeSuiteControls.TabControl tbcDetails 
         Height          =   975
         Left            =   720
         TabIndex        =   3
         Top             =   360
         Width           =   1935
         _Version        =   589884
         _ExtentX        =   3413
         _ExtentY        =   1720
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picClass 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4215
      Left            =   120
      ScaleHeight     =   4215
      ScaleWidth      =   2175
      TabIndex        =   0
      Top             =   840
      Width           =   2175
      Begin VB.CheckBox chkAllDetails 
         Caption         =   "��ʾ�����¼�ҩƷ"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1815
      End
      Begin MSComctlLib.TreeView tvwDetails 
         Height          =   4800
         Left            =   0
         TabIndex        =   9
         Tag             =   "1000"
         Top             =   600
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   8467
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "ImgTvw"
         Appearance      =   0
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   5190
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmBatchUpdate.frx":6852
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7064
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin MSComctlLib.ImageList ImgTvw 
      Left            =   1680
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchUpdate.frx":70E6
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchUpdate.frx":7680
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBatchUpdate.frx":DEE2
            Key             =   "���U"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgTool 
      Bindings        =   "frmBatchUpdate.frx":E47C
      Left            =   1320
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmBatchUpdate.frx":E490
   End
   Begin XtremeDockingPane.DockingPane dkpPanel 
      Left            =   720
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmBatchUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint״̬ As Integer         '��¼��Ʒ���޸Ļ��ǹ���޸� 1-Ʒ�� 2-���
Private mint���� As Integer         '��¼�ǲ����״μ��� 1-�״� 2-����
Private mblnData As Boolean  '�����ж��Ƿ��ڴ������ʱ��������ֵ
Private mstr�ϴνڵ� As String  '���������ϴ���ѡ�еĽڵ�
Private mintRow As Integer        '������¼�ϴ���ѡ�е��к�
Private mintRow�ϴ� As Integer
Private mintCol�ϴ� As Integer
Private mbln��� As Boolean        '������¼�Ƿ��п�� true-�п�� flase-�޿��
Private mblnҩ����� As Boolean    'ҩ����� true-���� false-������
Private mblnҩ������ As Boolean    'ҩ������ true-���� false-������
Private mint�Ƿ��� As Integer     '���ۻ���ʱ�� 0-���� 1-ʱ��
Private mstr��� As String         '������¼��ʲô���� �в�ҩ������ҩ���г�ҩ
Private mstrNode As String         '��¼������Ľڵ��ֵ
Private mstrPrivs As String        '��¼�û�����ЩȨ��
Private mrsRecord As ADODB.Recordset '������¼ѡ�нڵ��ѯ���������ݣ�Ϊ�Ժ�ָ�������׼��
Private mstrOtherName As String    '��¼����
Private mintOtherRow As Integer
Private mintExit As Integer         '������¼�˳�ʱ�Ƿ����˱��水ť 1
Private mintLen As Integer          '��¼סԺ��λ�ĳ���

'�Ӳ�������ȡҩƷ�۸�С��λ��
Private mintCostDigit As Integer        '�ɱ���С��λ��
Private mintPriceDigit As Integer       '�ۼ�С��λ��
Private mintSaleCostDigit As Integer
Private mintSalePriceDigit As Integer

Private mstrFind As String           '������¼ҪҪ��ѯ��ֵ
Private mlngFind As Long
Private mlngFindFirst As Long
Private mrsFindName As ADODB.Recordset
Private mstrValue As String         '������¼���ҿ��е�ֵ

Private mstrMatch As String         'ƥ�䷽ʽ
Private mstrOldValue As String      '��¼ԭ���ĵ�Ԫ���е�ֵ
Private mblnClick As Boolean
Private mblnSetKey As Boolean       '�ж��Ƿ�������
Private mint��ǰ��λ As Integer      '����ϵͳ���������õ���ʾ��λ

Private Const mconӦ���ڱ��� As Integer = 101
Private Const mconĬ��ֵ As Integer = 102
Private Const mcon���� As Integer = 103
Private Const mcon���� As Integer = 104
Private Const mcon�˳� As Integer = 105
Private Const mcon���� As Integer = 106
Private Const mconFind As Integer = 107

Private Const cstcolor_backcolor = &H80000005   '��ɫ
Private Const CSTCOLOR_UNMODIFY = &HC0C0FF       '�ۺ� ѡ��ҳ��ɫ
Private Const CSTCOLOR_NORECORDS = &HFFFFFF   '
Private Const mlngColor As Long = &H8000000F        '�����޸ĵ��н�������ɫ�ĳɻ�ɫ
Private Const mlngApplyColor As Long = &HB18383          '����ɫ

Private mobjPopup As CommandBar
Private mobjControl As CommandBarControl
Private mcbrToolBar As CommandBar


'Ʒ�����
Private Enum mVariList
    ������Ϣ = 0
    Ʒ������ = 1
    �ٴ�Ӧ�� = 2
End Enum
'Ʒ����
Private Enum mVaricolumn
    Ʒ��_��� = 0
    Ʒ��_id = 1
    Ʒ��_����id = 2
    Ʒ��_ҩƷ����
    Ʒ��_ҩƷ����
    Ʒ��_ͨ������
    Ʒ��_Ӣ������
    Ʒ��_ƴ����
    Ʒ��_�����
    'Ʒ������
    Ʒ��_�������
    Ʒ��_��ֵ����
    Ʒ��_��Դ���
    Ʒ��_��ҩ�ݴ�
    Ʒ��_ҩƷ����
    Ʒ��_����
    Ʒ��_ԭ��ҩ
    Ʒ��_ר��ҩ
    Ʒ��_��������
    Ʒ��_����ҩ
    Ʒ��_��ҩ
    Ʒ��_����ҩ
    Ʒ��_��ý
    Ʒ��_������ҩ
    Ʒ��_ԭ��ҩ
    Ʒ��_��ζʹ��
    '�ٴ�Ӧ��
    Ʒ��_�ο���Ŀ
    Ʒ��_����ְ��
    Ʒ��_ҽ��ְ��
    Ʒ��_��������
    Ʒ��_�����Ա�
    Ʒ��_������λ
    Ʒ��_Ƥ��
    Ʒ��_������
    Ʒ��_ATCCODE
    Ʒ��_Ʒ���³���ҽ��
    Ʒ��_�ο���ĿID
    Ʒ��_Count
End Enum

'������
Private Enum mSpecList
    ������Ϣ = 0
    ��Ʒ��Ϣ = 1
    ��װ��λ = 2
    �۸���Ϣ = 3
    ҩ������ = 4
    �������� = 5
    �ٴ�Ӧ�� = 6
    ��ҩ���� = 7
End Enum

'�����
Private Enum mSpecColumn
    ���_��� = 0
    ���_id = 1
    ���_ҩ��id = 2
'    ���_ҩƷ���� = 3
    ���_ͨ������
    ���_������
    ���_ҩƷ���
    ���_��λ��
    ���_������
    ���_��ʶ��
    ���_��ѡ��
    ���_����
    ���_��Ʒ����
    ���_��������
    ���_��Դ����
    ���_ƴ����
    ���_�����
    ���_��ͬ��λ
    ���_��׼�ĺ�
    ���_ע���̱�
    ���_GMP��֤
    ���_�ǳ���ҩ
    ���_�ۼ۵�λ
    ���_����ϵ��
    ���_������λ
    ���_סԺ��λ
    ���_סԺϵ��
    ���_���ﵥλ
    ���_����ϵ��
    ���_ҩ�ⵥλ
    ���_ҩ��ϵ��
    ���_�ͻ���λ
    ���_�ͻ���װ
    ���_���쵥λ
    ���_���췧ֵ
    ���_��ҩ��̬
    ���_ҩ������
    ���_�ɹ��޼�
    ���_�ɹ�����
    ���_�����
    ���_ָ���ۼ�
    ���_ָ������
    ���_�ӳ���
    ���_�������
    ���_�ɱ��۸�
    ���_��ǰ�ۼ�
    ���_������Ŀ
    ���_������Ŀ
    ���_����ѱ���
    ���_ҩ�ۼ���
    ���_���ηѱ�
    ���_��ֵ˰��
    ���_ҽ������
    ���_ҩ�����
    ���_ҩ������
    ���_������
    ���_��ʶ˵��
    ���_��ҩ����
    ���_վ����
    ���_DDDֵ
    ���_�������
    ���_סԺ����ʹ��
    ���_סԺ��̬����
    ���_�������ʹ��
    ���_��ΣҩƷ
    ���_����ҩ��
    ���_�洢�¶�
    ���_�洢����
    ���_��ҩ����
    ���_��Һע������
    ���_�������
    ���_�б�ҩƷ
    ���_��ͬ��λid
    ���_������Ŀid
    ���_ԭҩ�����
    ���_ԭҩ������
    ���_count = 75
End Enum

Private Sub CheckValue(ByVal intRow As Integer, ByVal lngҩƷID As Long)
    Dim rsTemp As ADODB.Recordset
    Dim dblTemp As Double
    
    gstrSql = ""
    On Error GoTo ErrHandle
    With vsfDetails
        If .TextMatrix(intRow, mSpecColumn.���_ҩ�����) = "0" Then
            .Cell(flexcpBackColor, intRow, mSpecColumn.���_ҩ������, intRow) = mlngColor: .TextMatrix(intRow, mSpecColumn.���_ҩ������) = 0
            .Cell(flexcpBackColor, intRow, mSpecColumn.���_������, intRow) = mlngColor: .TextMatrix(intRow, mSpecColumn.���_������) = 0
        Else
            If Val(.TextMatrix(intRow, mSpecColumn.���_������)) = 0 Then
                .Cell(flexcpBackColor, intRow, mSpecColumn.���_������, intRow) = mlngColor
            End If
        End If
        
        '��ȡ��ʾ��ǰ�ۼ�
        If Mid(.TextMatrix(intRow, mSpecColumn.���_ҩ������), 1, 1) <> 0 Then
            'ʱ��ҩƷ��ȡ�����/���������Ϊ��۸��޿��ʱȡ�۱��� ��ʱ��ҩƷ���ۣ�ȡ��۸��¼�еļ۸�
            gstrSql = "select Decode(K.�������,0,P.�ּ�,K.�����/Nvl(K.�������,1)) as �ּ�,P.������Ŀid" & _
                    " from �շѼ�Ŀ P," & _
                    "     (Select nvl(Sum(ʵ�ʽ��),0) as �����,nvl(Sum(ʵ������),0) as �������" & _
                    "      From ҩƷ��� Where ҩƷID=[1]) K" & _
                    " where P.�շ�ϸĿid=[1] and (P.��ֹ���� is null or Sysdate Between P.ִ������ And P.��ֹ����)"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngҩƷID)
        End If
    
        If gstrSql <> "" Then
            If rsTemp.RecordCount > 0 Then
                If Val(mint��ǰ��λ) <> 0 Then
                    .TextMatrix(intRow, mSpecColumn.���_��ǰ�ۼ�) = FormatEx(rsTemp!�ּ� * Val(.TextMatrix(intRow, mSpecColumn.���_ҩ��ϵ��)), mintPriceDigit)
                Else
                    .TextMatrix(intRow, mSpecColumn.���_��ǰ�ۼ�) = FormatEx(rsTemp!�ּ�, mintPriceDigit)
                End If
                .TextMatrix(intRow, mSpecColumn.���_������Ŀid) = rsTemp!������Ŀid
            End If
        End If

        '�����Ƿ��з�����ȷ����ҩ�����ԡ��ɱ��۸����ۼ۸���޸ķ�
        gstrSql = " Select nvl(Count(*),0) From ҩƷ�շ���¼ Where ҩƷID=[1] And rownum<2"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngҩƷID)
        
        If rsTemp.Fields(0).Value > 0 Then
            If Mid(.TextMatrix(intRow, mSpecColumn.���_ҩ������), 1, 1) <> 0 Then .Cell(flexcpBackColor, intRow, mSpecColumn.���_ҩ������, intRow) = mlngColor
            .Cell(flexcpBackColor, intRow, mSpecColumn.���_�ɱ��۸�, intRow) = mlngColor
            .Cell(flexcpBackColor, intRow, mSpecColumn.���_��ǰ�ۼ�, intRow) = mlngColor
            .Cell(flexcpBackColor, intRow, mSpecColumn.���_������Ŀ, intRow) = mlngColor
            .Cell(flexcpBackColor, intRow, mSpecColumn.���_סԺϵ��, intRow) = mlngColor
            .Cell(flexcpBackColor, intRow, mSpecColumn.���_����ϵ��, intRow) = mlngColor
            .Cell(flexcpBackColor, intRow, mSpecColumn.���_ҩ��ϵ��, intRow) = mlngColor
        End If
        
        '����Ʒ���Ƿ��ǿ�����ȷ�����DDDֵ�Ƿ��ܹ��޸�
        gstrSql = " Select Nvl(b.������, 0) As ������ From ҩƷ��� A, ҩƷ���� B Where a.ҩ��id = b.ҩ��id And a.ҩƷid = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngҩƷID)
        If rsTemp!������ = 0 Then
            .Cell(flexcpBackColor, intRow, mSpecColumn.���_DDDֵ, intRow) = mlngColor
        End If
        
        '�����Ƿ����ҽ����¼��ȷ������ϵ���Ƿ��ܹ��޸�
        gstrSql = "Select 1 From ����ҽ����¼ Where �շ�ϸĿID=[1] And Rownum=1"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngҩƷID)
        If rsTemp.RecordCount > 0 Then
            .Cell(flexcpBackColor, intRow, mSpecColumn.���_����ϵ��, intRow) = mlngColor
        End If
        
        '�����Ƿ��п�棬ȷ�����������Կ��޸ķ�
        gstrSql = " Select nvl(Count(*),0) From ҩƷ��� A,��������˵�� B" & _
                 " Where A.ҩƷID=[1] And A.�ⷿID=B.����ID And B.�������� Like '%ҩ��'"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngҩƷID)
        
        If rsTemp.Fields(0).Value > 0 Then
            .Cell(flexcpBackColor, intRow, mSpecColumn.���_ҩ�����, intRow) = mlngColor
            .Cell(flexcpBackColor, intRow, mSpecColumn.���_������, intRow) = mlngColor
        End If
        If .TextMatrix(intRow, mSpecColumn.���_ҩ�����) <> "0" Then
            gstrSql = " Select nvl(Count(*),0) From ҩƷ��� A,��������˵�� B" & _
                     " Where A.ҩƷID=[1] And A.�ⷿID=B.����ID And (B.�������� Like '%ҩ��' Or B.�������� Like '%�Ƽ���')"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngҩƷID)
            
            If rsTemp.Fields(0).Value > 0 Then
                .Cell(flexcpBackColor, intRow, mSpecColumn.���_ҩ������, intRow) = mlngColor
                If .Cell(flexcpBackColor, intRow, mSpecColumn.���_ҩ�����) <> mlngColor Then
                    .Cell(flexcpBackColor, intRow, mSpecColumn.���_ҩ�����, intRow) = IIf(.TextMatrix(intRow, mSpecColumn.���_ҩ������) = "0", cstcolor_backcolor, mlngColor)
                End If
            End If
        End If
            .Cell(flexcpBackColor, intRow, mSpecColumn.���_�����, intRow) = mlngColor
            If Val(Mid(.TextMatrix(intRow, mSpecColumn.���_סԺ����ʹ��), 1, 1)) = 0 Then
                .Cell(flexcpBackColor, intRow, mSpecColumn.���_סԺ��̬����, intRow) = mlngColor
            End If
            If .TextMatrix(intRow, mSpecColumn.���_��ҩ��̬) = "ɢװ" And mstrNode Like "�в�ҩ*" Then
                .Cell(flexcpBackColor, intRow, mSpecColumn.���_סԺ����ʹ��, intRow) = mlngColor
                .Cell(flexcpBackColor, intRow, mSpecColumn.���_�������ʹ��, intRow) = mlngColor
            End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub ShowMe(ByVal int״̬ As Integer, ByVal strPrivs As String)
    '�ṩ����������ʱ�����Ĺ��÷���
    mint״̬ = int״̬
    mstrPrivs = strPrivs
    
    Me.Show vbModal, frmMediLists
End Sub

Private Sub InitTreeView()
    With tvwDetails
        .LabelEdit = 1  '����treeviewΪ���ɱ༭״̬
    End With
End Sub

Private Sub InitComandBars()
    '��ʼ���������������˵���
    Dim cbrControlMain As CommandBarControl
    Dim ctrCustom As CommandBarControlCustom
    Dim intCount As Integer
    
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    Me.cbsMain.VisualTheme = xtpThemeOffice2003 + xtpThemeOfficeXP

    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16

    End With
    
    Me.cbsMain.EnableCustomization False
    Me.cbsMain.Icons = imgTool.Icons
    
    '����������
    Set mcbrToolBar = Me.cbsMain.Add("������", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.EnableDocking xtpFlagStretched Or xtpFlagAlignAny
    
    With mcbrToolBar.Controls
        Set cbrControlMain = .Add(xtpControlButton, mconӦ���ڱ���, "Ӧ���ڱ���")
        cbrControlMain.Style = xtpButtonIconAndCaption  'ͬʱ��ʾͼ�������
        Set cbrControlMain = .Add(xtpControlButton, mconĬ��ֵ, "�ָ�Ĭ��ֵ")
        cbrControlMain.Style = xtpButtonIconAndCaption  'ͬʱ��ʾͼ�������
        
        Set cbrControlMain = .Add(xtpControlButton, mcon����, "����")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Style = xtpButtonIconAndCaption  'ͬʱ��ʾͼ�������
        
'        Set cbrControlMain = .Add(xtpControlButton, mcon����, "����")
'        cbrControlMain.Style = xtpButtonIconAndCaption  'ͬʱ��ʾͼ�������
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, mcon�˳�, "�˳�")
        cbrControlMain.Style = xtpButtonIconAndCaption  'ͬʱ��ʾͼ�������
        
        Set cbrControlMain = .Add(xtpControlLabel, mcon����, "����")
        cbrControlMain.Flags = xtpFlagRightAlign    '���Ҷ���

        Set ctrCustom = mcbrToolBar.Controls.Add(xtpControlCustom, mconFind, "��ѯ")
        ctrCustom.Handle = txtFind.hWnd
        ctrCustom.Flags = xtpFlagRightAlign
    End With
    
    cbsMain.Item(1).Delete
    
    '�Ҽ��˵�
    Set mobjPopup = cbsMain.Add("Popup", xtpBarPopup)
    With mobjPopup.Controls
        Set mobjControl = .Add(xtpControlButton, mconӦ���ڱ���, "Ӧ���ڱ���")
        Set mobjControl = .Add(xtpControlButton, mconĬ��ֵ, "�ָ�Ĭ��ֵ")
    End With
    
    '�����
    With Me.cbsMain.KeyBindings
        .Add 0, VK_F3, mconFind
    End With
End Sub

Private Sub initPanel()
    '��ʼ�������ؼ�
    'DockingPane
    '-----------------------------------------------------
    Dim objPaneCon As Pane
    Dim objPaneDetail As Pane
    
    Me.dkpPanel.SetCommandBars Me.cbsMain
    Me.dkpPanel.Options.UseSplitterTracker = False 'ʵʱ�϶�
    Me.dkpPanel.Options.ThemedFloatingFrames = True
    Me.dkpPanel.Options.AlphaDockingContext = True
    
    Set objPaneCon = Me.dkpPanel.CreatePane(1, 200, 0, DockLeftOf, Nothing)
    objPaneCon.Options = PaneNoCloseable Or PaneNoFloatable
    objPaneCon.Title = "����"
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strTemp As String
    
    Select Case Control.ID
        Case mconӦ���ڱ���
            Call SetBatch
        Case mconĬ��ֵ
'            mrsRecord.MoveFirst
'            Call showColumn(mrsRecord, mstrNode)
            Call tvwDetails_NodeClick(tvwDetails.Nodes(tvwDetails.SelectedItem.Index))
        Case mcon����
            Call Save
        Case mconFind
'            If TypeName(Control) = "ICommandBarButton" Then
'                Call FindGridRow(mstrValue)
'            Else
'                strTemp = Trim(UCase(Control.Text))
'                mstrValue = strTemp
'            End If
'            Call FindGridRow(strTemp)
        Case mcon�˳�
            Call ExitFrom
    End Select
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    On Error Resume Next
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    Me.picDetails.Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop - stbThis.Height
    
    Call initControl
End Sub

Private Sub chkAllDetails_Click()
    If mint״̬ = 1 Then
        With vsfDetails
            If chkAllDetails.Value = 1 Then
                .ColWidth(mVaricolumn.Ʒ��_ҩƷ����) = 2000
                .ColHidden(mVaricolumn.Ʒ��_ҩƷ����) = False
            Else
                .ColHidden(mVaricolumn.Ʒ��_ҩƷ����) = True
            End If
        End With
    End If
    Call tvwDetails_NodeClick(tvwDetails.SelectedItem)
End Sub

Private Sub dkpPanel_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
        Case 1
            Item.Handle = picClass.hWnd '���ؼ����뵽dockingpanel�ؼ���
    End Select
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim j As Integer
    
    Me.Width = 14000    '��һ�μ���ʱ�������С
    Me.Height = 9000
    
    Call RestoreWinState(Me, App.ProductName, Me.Caption)
    If GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser, "ʹ�ø��Ի����", "1") = "1" Then
        chkAllDetails = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name, "�Ƿ���ʾ�¼�", 0)
    End If
    
    mint���� = 1
    Call InitTreeView   '��ʼ����
    Call InitComandBars '��ʼ���˵��͹�����
    Call initPanel  '��ʼ�����
    Call InitTabControl '��TabControl�ؼ��м��봰��
    Call initControl    '��ʼ���ؼ�
    
    If mint״̬ = 1 Then
        Call initColumn_Ʒ����Ϣ    '��ʼ��Ʒ����
        mint���� = 2
    ElseIf mint״̬ = 2 Then
        Call initColumn_�����Ϣ
        mint���� = 2
    End If
    
'    mstrNode = "����ҩ"
    mblnData = ReadAndSendDataToTvw(mint״̬)     '���������ֵ
    Call setColumn(0)    '��ʼ��vsflexgrid�ؼ���
    Call SetȨ���ж� 'Ȩ���ж�
    
    mstrMatch = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = 0, "%", "")  'ƥ�䷽ʽ
    mint��ǰ��λ = Val(GetSysPara(29))  '��¼��ǰ���õ���ʾ��λ
        
    mintCostDigit = GetDigit(1, 1, IIf(mint��ǰ��λ = 0, 1, 4))
    mintPriceDigit = GetDigit(1, 2, IIf(mint��ǰ��λ = 0, 1, 4))
    
    mintSaleCostDigit = GetDigit(1, 1, 1)
    mintSalePriceDigit = GetDigit(1, 2, 1)
    
    If tvwDetails.Nodes.Count > 0 Then
        If chkAllDetails = 1 And Not tvwDetails.Nodes(tvwDetails.SelectedItem.Index) Is Nothing Then
            Call tvwDetails_NodeClick(tvwDetails.Nodes(tvwDetails.SelectedItem.Index))
        End If
    End If
End Sub

Private Sub initControl()
    '���²��ֿؼ�λ��
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    On Error Resume Next
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    chkAllDetails.Move 0, 0, picClass.Width
    tvwDetails.Move 0, chkAllDetails.Height + chkAllDetails.Top, picClass.ScaleWidth, lngBottom - lngTop - stbThis.Height - chkAllDetails.Height - 300
    tbcDetails.Move 0, 0, picDetails.ScaleWidth, picDetails.ScaleHeight
     
    If mint״̬ = 1 Then    'Ʒ�ֲ��б���
        frmBatchUpdate.Caption = "Ʒ�������޸�"
        vsfDetails.Move 0, 0, picList.ScaleWidth, picList.ScaleHeight
        fraSplit.Visible = False
        vsfOtherName.Visible = False
    Else    '����ޱ���
        frmBatchUpdate.Caption = "��������޸�"
        
        vsfDetails.Move 0, 0, picList.ScaleWidth, picList.ScaleHeight
        fraSplit.Visible = False
        vsfOtherName.Visible = False
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Call initControl
End Sub

Private Sub InitTabControl()
    '��ʼ��Tabcontrol�ؼ�
    With Me.tbcDetails
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        
        If mint״̬ = 1 Then    'Ʒ��
            .InsertItem(mVariList.������Ϣ, "������Ϣ", picList.hWnd, 0).Tag = "������Ϣ_"
            .InsertItem(mVariList.Ʒ������, "Ʒ������", picList.hWnd, 0).Tag = "Ʒ������_"
            .InsertItem(mVariList.�ٴ�Ӧ��, "�ٴ�Ӧ��", picList.hWnd, 0).Tag = "�ٴ�Ӧ��_"
            
            .Item(mVariList.Ʒ������).Selected = True
            .Item(mVariList.������Ϣ).Selected = True
            
        Else    '���
            .InsertItem(mSpecList.������Ϣ, "������Ϣ", picList.hWnd, 0).Tag = "������Ϣ_"
            .InsertItem(mSpecList.��Ʒ��Ϣ, "��Ʒ��Ϣ", picList.hWnd, 0).Tag = "��Ʒ��Ϣ_"
            .InsertItem(mSpecList.��װ��λ, "��װ��λ", picList.hWnd, 0).Tag = "��װ��λ_"
            .InsertItem(mSpecList.�۸���Ϣ, "�۸���Ϣ", picList.hWnd, 0).Tag = "�۸���Ϣ_"
            .InsertItem(mSpecList.ҩ������, "ҩ������", picList.hWnd, 0).Tag = "ҩ������_"
            .InsertItem(mSpecList.��������, "��������", picList.hWnd, 0).Tag = "��������_"
            .InsertItem(mSpecList.�ٴ�Ӧ��, "�ٴ�Ӧ��", picList.hWnd, 0).Tag = "�ٴ�Ӧ��_"
            .InsertItem(mSpecList.��ҩ����, "��ҩ����", picList.hWnd, 0).Tag = "��ҩ����_"
            
            .Item(mSpecList.��Ʒ��Ϣ).Selected = True
            .Item(mSpecList.������Ϣ).Selected = True
        End If
    End With
    Call setTabControlColor(tbcDetails)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call Recover
    mblnSetKey = False
    mintExit = 0
    Call SaveWinState(Me, App.ProductName, Me.Caption)
    
    If GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser, "ʹ�ø��Ի����", "1") = "1" Then
        Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\��������\" & App.ProductName & "\" & Me.Name, "�Ƿ���ʾ�¼�", chkAllDetails.Value)
    End If
    Unload Me
End Sub
Private Sub fraSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 And vsfDetails.Height + y > 100 And fraSplit.Height + fraSplit.Top + y < stbThis.Top - 1000 Then
        vsfDetails.Move 0, 0, picList.ScaleWidth, vsfDetails.Height + y
        fraSplit.Move 0, fraSplit.Top + y, picList.ScaleWidth, 50
        vsfOtherName.Move 0, fraSplit.Top + fraSplit.Height, picList.ScaleWidth, vsfOtherName.Height - y
    End If
End Sub

Private Sub tbcDetails_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'ֻ����ҩƷƷ�ֻ���ҳ��Ż��б������û��޸�

    If mint״̬ = 1 Then    'Ʒ��
        fraSplit.Visible = False
        vsfOtherName.Visible = False
        vsfDetails.Move 0, 0, picList.ScaleWidth, picList.ScaleHeight
    Else
        vsfDetails.Move 0, 0, picList.ScaleWidth, picList.ScaleHeight
    End If
    
    Call setTabControlColor(tbcDetails)

    If mint���� = 2 Then    'ֻ�����г�ʼ������ܽ���������
        Call setColumn(Item.Index)  '��������ʾ����
    End If
End Sub

Private Sub setTabControlColor(ByVal objtbc As TabControl)
    '��Tabcontrol�ؼ�������ɫ�ж�
    Dim i As Integer
    
    With objtbc
        For i = 0 To .ItemCount - 1
            If .Item(i).Selected = True Then
                .Item(i).Color = CSTCOLOR_UNMODIFY
            Else
                .Item(i).Color = CSTCOLOR_NORECORDS
            End If
        Next
    End With
End Sub

Private Sub setColumn(ByVal intPageItem As Integer)
    '����ʾ����������
    With vsfDetails
        .Editable = flexEDKbdMouse
        .MergeCells = flexMergeRestrictColumns
        .GridLineWidth = 2
        .GridLines = flexGridInset
        .GridColor = &H0&
        .ExplorerBar = flexExSortShowAndMove
        .AllowUserResizing = flexResizeColumns
        .AllowSelection = False '���ܶ�ѡ��Ԫ��
    End With
    
    With vsfDetails
        If mint״̬ = 1 Then 'Ʒ��
            vsfDetails.MergeCol(mVaricolumn.Ʒ��_ҩƷ����) = True   '�������.MergeCells���Խ��ʹ�ò�ͬ��ͬ��������ͬ�ĺϲ�
            '������Ϣ
            .ColWidth(mVaricolumn.Ʒ��_���) = 600
            .ColHidden(mVaricolumn.Ʒ��_id) = True
            .ColHidden(mVaricolumn.Ʒ��_����id) = True
            .ColHidden(mVaricolumn.Ʒ��_�ο���ĿID) = True
            
            .ColWidth(mVaricolumn.Ʒ��_ͨ������) = 2000 '�����ظ���
            .ColHidden(mVaricolumn.Ʒ��_ҩƷ����) = IIf(intPageItem = mVariList.������Ϣ, False, True)
            .ColHidden(mVaricolumn.Ʒ��_Ӣ������) = IIf(intPageItem = mVariList.������Ϣ, False, True)
            .ColHidden(mVaricolumn.Ʒ��_ƴ����) = IIf(intPageItem = mVariList.������Ϣ, False, True)
            .ColHidden(mVaricolumn.Ʒ��_�����) = IIf(intPageItem = mVariList.������Ϣ, False, True)
            
            'Ʒ������
            .ColHidden(mVaricolumn.Ʒ��_�������) = IIf(intPageItem = mVariList.Ʒ������, False, True)
            .ColHidden(mVaricolumn.Ʒ��_��ֵ����) = IIf(intPageItem = mVariList.Ʒ������, False, True)
            .ColHidden(mVaricolumn.Ʒ��_��Դ���) = IIf(intPageItem = mVariList.Ʒ������, False, True)
            .ColHidden(mVaricolumn.Ʒ��_��ҩ�ݴ�) = IIf(intPageItem = mVariList.Ʒ������, False, True)
            .ColHidden(mVaricolumn.Ʒ��_ҩƷ����) = IIf(intPageItem = mVariList.Ʒ������, False, True)
            .ColHidden(mVaricolumn.Ʒ��_����) = IIf(intPageItem = mVariList.Ʒ������, False, True)
            .ColHidden(mVaricolumn.Ʒ��_ԭ��ҩ) = IIf(intPageItem = mVariList.Ʒ������, False, True)
            .ColHidden(mVaricolumn.Ʒ��_ר��ҩ) = IIf(intPageItem = mVariList.Ʒ������, False, True)
            .ColHidden(mVaricolumn.Ʒ��_��������) = IIf(intPageItem = mVariList.Ʒ������, False, True)
            .ColHidden(mVaricolumn.Ʒ��_����ҩ) = IIf(intPageItem = mVariList.Ʒ������, False, True)
            .ColHidden(mVaricolumn.Ʒ��_��ҩ) = IIf(intPageItem = mVariList.Ʒ������, False, True)
            .ColHidden(mVaricolumn.Ʒ��_ԭ��ҩ) = IIf(intPageItem = mVariList.Ʒ������, False, True)
            .ColHidden(mVaricolumn.Ʒ��_��ζʹ��) = IIf(intPageItem = mVariList.Ʒ������, False, True)
            .ColHidden(mVaricolumn.Ʒ��_������ҩ) = IIf(intPageItem = mVariList.Ʒ������, False, True)
            .ColHidden(mVaricolumn.Ʒ��_����ҩ) = IIf(intPageItem = mVariList.Ʒ������, False, True)
            .ColHidden(mVaricolumn.Ʒ��_��ý) = IIf(intPageItem = mVariList.Ʒ������, False, True)
            
            '�ٴ�Ӧ��
            .ColHidden(mVaricolumn.Ʒ��_�ο���Ŀ) = IIf(intPageItem = mVariList.�ٴ�Ӧ��, False, True)
            .ColHidden(mVaricolumn.Ʒ��_����ְ��) = IIf(intPageItem = mVariList.�ٴ�Ӧ��, False, True)
            .ColHidden(mVaricolumn.Ʒ��_ҽ��ְ��) = IIf(intPageItem = mVariList.�ٴ�Ӧ��, False, True)
            .ColHidden(mVaricolumn.Ʒ��_��������) = IIf(intPageItem = mVariList.�ٴ�Ӧ��, False, True)
            .ColHidden(mVaricolumn.Ʒ��_�����Ա�) = IIf(intPageItem = mVariList.�ٴ�Ӧ��, False, True)
            .ColHidden(mVaricolumn.Ʒ��_������λ) = IIf(intPageItem = mVariList.�ٴ�Ӧ��, False, True)
            .ColHidden(mVaricolumn.Ʒ��_Ƥ��) = IIf(intPageItem = mVariList.�ٴ�Ӧ��, False, True)
            .ColHidden(mVaricolumn.Ʒ��_������) = IIf(intPageItem = mVariList.�ٴ�Ӧ��, False, True)
            .ColHidden(mVaricolumn.Ʒ��_ATCCODE) = IIf(intPageItem = mVariList.�ٴ�Ӧ��, False, True)
            .ColHidden(mVaricolumn.Ʒ��_Ʒ���³���ҽ��) = IIf(intPageItem = mVariList.�ٴ�Ӧ��, False, True)
            
            If mstrNode Like "�в�ҩ*" And intPageItem = mVariList.�ٴ�Ӧ�� Then
                .ColHidden(mVaricolumn.Ʒ��_Ƥ��) = True
                .ColHidden(mVaricolumn.Ʒ��_������) = True
                .ColHidden(mVaricolumn.Ʒ��_ATCCODE) = True
                .ColHidden(mVaricolumn.Ʒ��_Ʒ���³���ҽ��) = True
            Else
                If intPageItem = mVariList.�ٴ�Ӧ�� Then
                    .ColHidden(mVaricolumn.Ʒ��_Ƥ��) = False
                    .ColHidden(mVaricolumn.Ʒ��_������) = False
                    .ColHidden(mVaricolumn.Ʒ��_ATCCODE) = False
                    .ColHidden(mVaricolumn.Ʒ��_Ʒ���³���ҽ��) = False
                End If
            End If
            
            If mstrNode Like "�в�ҩ*" Then
                If intPageItem = mVariList.Ʒ������ Then
                    .ColHidden(mVaricolumn.Ʒ��_��ζʹ��) = False
                    .ColHidden(mVaricolumn.Ʒ��_ԭ��ҩ) = False
                End If
                .ColHidden(mVaricolumn.Ʒ��_����) = True
                .ColHidden(mVaricolumn.Ʒ��_ԭ��ҩ) = True
                .ColHidden(mVaricolumn.Ʒ��_ר��ҩ) = True
                .ColHidden(mVaricolumn.Ʒ��_��������) = True
                .ColHidden(mVaricolumn.Ʒ��_����ҩ) = True
                .ColHidden(mVaricolumn.Ʒ��_��ҩ) = True
                .ColHidden(mVaricolumn.Ʒ��_����ҩ) = True
                .ColHidden(mVaricolumn.Ʒ��_��ý) = True
            Else
                .ColHidden(mVaricolumn.Ʒ��_��ζʹ��) = True
                If intPageItem = mVariList.Ʒ������ Then
                    .ColHidden(mVaricolumn.Ʒ��_����) = False
                    .ColHidden(mVaricolumn.Ʒ��_ԭ��ҩ) = False
                    .ColHidden(mVaricolumn.Ʒ��_ר��ҩ) = False
                    .ColHidden(mVaricolumn.Ʒ��_��������) = False
                    .ColHidden(mVaricolumn.Ʒ��_����ҩ) = False
                    .ColHidden(mVaricolumn.Ʒ��_��ҩ) = False
                    .ColHidden(mVaricolumn.Ʒ��_ԭ��ҩ) = False
                    .ColHidden(mVaricolumn.Ʒ��_����ҩ) = False
                    .ColHidden(mVaricolumn.Ʒ��_��ý) = False
                End If
            End If
            
            If chkAllDetails.Value = 1 Then
                .ColHidden(mVaricolumn.Ʒ��_ҩƷ����) = False
            Else
                .ColHidden(mVaricolumn.Ʒ��_ҩƷ����) = True
            End If
        Else    '���
            vsfDetails.MergeCol(mSpecColumn.���_ͨ������) = True    '���úϲ�
            
            .ColWidth(mSpecColumn.���_���) = 600
'            .ColWidth(mSpecColumn.���_ҩƷ����) = 1000
            .ColWidth(mSpecColumn.���_ͨ������) = 1800
            .ColWidth(mSpecColumn.���_ҩƷ���) = 1500
            .ColHidden(mSpecColumn.���_id) = True
            .ColHidden(mSpecColumn.���_ҩ��id) = True
            .ColHidden(mSpecColumn.���_�б�ҩƷ) = True
            .ColHidden(mSpecColumn.���_��ͬ��λid) = True
            .ColHidden(mSpecColumn.���_������Ŀid) = True
            .ColHidden(mSpecColumn.���_ԭҩ�����) = True
            .ColHidden(mSpecColumn.���_ԭҩ������) = True
            '������Ϣ
            .ColHidden(mSpecColumn.���_������) = IIf(intPageItem = mSpecList.������Ϣ, False, True)
            .ColHidden(mSpecColumn.���_��λ��) = IIf(intPageItem = mSpecList.������Ϣ, False, True)
            .ColHidden(mSpecColumn.���_������) = IIf(intPageItem = mSpecList.������Ϣ, False, True)
            .ColHidden(mSpecColumn.���_��ʶ��) = IIf(intPageItem = mSpecList.������Ϣ, False, True)
            .ColHidden(mSpecColumn.���_��ѡ��) = IIf(intPageItem = mSpecList.������Ϣ, False, True)
            
            If mstrNode Like "�в�ҩ*" Then
                .ColHidden(mSpecColumn.���_����) = True
            Else
                .ColHidden(mSpecColumn.���_����) = IIf(intPageItem = mSpecList.������Ϣ, False, True)
            End If
            '��Ʒ��Ϣ
            .ColHidden(mSpecColumn.���_��Ʒ����) = IIf(intPageItem = mSpecList.��Ʒ��Ϣ, False, True)
            .ColHidden(mSpecColumn.���_��������) = IIf(intPageItem = mSpecList.��Ʒ��Ϣ, False, True)
            .ColHidden(mSpecColumn.���_��Դ����) = IIf(intPageItem = mSpecList.��Ʒ��Ϣ, False, True)
            .ColHidden(mSpecColumn.���_��ͬ��λ) = IIf(intPageItem = mSpecList.��Ʒ��Ϣ, False, True)
            .ColHidden(mSpecColumn.���_��׼�ĺ�) = IIf(intPageItem = mSpecList.��Ʒ��Ϣ, False, True)
            .ColHidden(mSpecColumn.���_ע���̱�) = IIf(intPageItem = mSpecList.��Ʒ��Ϣ, False, True)
            .ColHidden(mSpecColumn.���_ƴ����) = IIf(intPageItem = mSpecList.��Ʒ��Ϣ, False, True)
            .ColHidden(mSpecColumn.���_�����) = IIf(intPageItem = mSpecList.��Ʒ��Ϣ, False, True)
            .ColHidden(mSpecColumn.���_GMP��֤) = IIf(intPageItem = mSpecList.��Ʒ��Ϣ, False, True)
            .ColHidden(mSpecColumn.���_�ǳ���ҩ) = IIf(intPageItem = mSpecList.��Ʒ��Ϣ, False, True)
            If mstrNode Like "�в�ҩ*" Then
                .ColHidden(mSpecColumn.���_ƴ����) = True
                .ColHidden(mSpecColumn.���_�����) = True
                .ColHidden(mSpecColumn.���_GMP��֤) = True
                .ColHidden(mSpecColumn.���_��Ʒ����) = True
            End If
            
            '��װ��λ
            .ColHidden(mSpecColumn.���_�ۼ۵�λ) = IIf(intPageItem = mSpecList.��װ��λ, False, True)
            .ColHidden(mSpecColumn.���_����ϵ��) = IIf(intPageItem = mSpecList.��װ��λ, False, True)
            .ColHidden(mSpecColumn.���_������λ) = IIf(intPageItem = mSpecList.��װ��λ, False, True)
            .ColHidden(mSpecColumn.���_סԺ��λ) = IIf(intPageItem = mSpecList.��װ��λ, False, True)
            .ColHidden(mSpecColumn.���_סԺϵ��) = IIf(intPageItem = mSpecList.��װ��λ, False, True)
            .ColHidden(mSpecColumn.���_���ﵥλ) = IIf(intPageItem = mSpecList.��װ��λ, False, True)
            .ColHidden(mSpecColumn.���_����ϵ��) = IIf(intPageItem = mSpecList.��װ��λ, False, True)
            .ColHidden(mSpecColumn.���_ҩ�ⵥλ) = IIf(intPageItem = mSpecList.��װ��λ, False, True)
            .ColHidden(mSpecColumn.���_ҩ��ϵ��) = IIf(intPageItem = mSpecList.��װ��λ, False, True)
            .ColHidden(mSpecColumn.���_�ͻ���λ) = IIf(intPageItem = mSpecList.��װ��λ, False, True)
            .ColHidden(mSpecColumn.���_�ͻ���װ) = IIf(intPageItem = mSpecList.��װ��λ, False, True)
            .ColHidden(mSpecColumn.���_���쵥λ) = IIf(intPageItem = mSpecList.��װ��λ, False, True)
            .ColHidden(mSpecColumn.���_���췧ֵ) = IIf(intPageItem = mSpecList.��װ��λ, False, True)
            .ColHidden(mSpecColumn.���_��ҩ��̬) = IIf(intPageItem = mSpecList.��װ��λ, False, True)
            
            If mstrNode Like "�в�ҩ*" Then
                If intPageItem = mSpecList.��װ��λ Then
                    .ColHidden(mSpecColumn.���_��ҩ��̬) = False
                    .ColHidden(mSpecColumn.���_���ﵥλ) = True
                    .ColHidden(mSpecColumn.���_����ϵ��) = True
                    VsfGridColFormat vsfDetails, mSpecColumn.���_סԺ��λ, "ҩ����λ", 1000, flexAlignLeftCenter, "ҩ����λ"
                    VsfGridColFormat vsfDetails, mSpecColumn.���_סԺϵ��, "ҩ��ϵ��", 1000, flexAlignRightCenter, "ҩ��ϵ��"
                End If
            Else
                VsfGridColFormat vsfDetails, mSpecColumn.���_סԺ��λ, "סԺ��λ", 1000, flexAlignLeftCenter, "סԺ��λ"
                VsfGridColFormat vsfDetails, mSpecColumn.���_סԺϵ��, "סԺϵ��", 1000, flexAlignRightCenter, "סԺϵ��"
                .ColHidden(mSpecColumn.���_��ҩ��̬) = True
            End If
            '�۸���Ϣ
            .ColHidden(mSpecColumn.���_ҩ������) = IIf(intPageItem = mSpecList.�۸���Ϣ, False, True)
            .ColHidden(mSpecColumn.���_�ɹ��޼�) = IIf(intPageItem = mSpecList.�۸���Ϣ, False, True)
            .ColHidden(mSpecColumn.���_�ɹ�����) = IIf(intPageItem = mSpecList.�۸���Ϣ, False, True)
            .ColHidden(mSpecColumn.���_�����) = IIf(intPageItem = mSpecList.�۸���Ϣ, False, True)
            .ColHidden(mSpecColumn.���_ָ���ۼ�) = IIf(intPageItem = mSpecList.�۸���Ϣ, False, True)
            .ColHidden(mSpecColumn.���_ָ������) = IIf(intPageItem = mSpecList.�۸���Ϣ, False, True)
            .ColHidden(mSpecColumn.���_�ӳ���) = IIf(intPageItem = mSpecList.�۸���Ϣ, False, True)
            .ColHidden(mSpecColumn.���_�������) = IIf(intPageItem = mSpecList.�۸���Ϣ, False, True)
            If mstrNode Like "�в�ҩ*" Then
                .ColHidden(mSpecColumn.���_�������) = True
            Else
                If intPageItem = mSpecList.�۸���Ϣ Then
                    .ColHidden(mSpecColumn.���_�������) = False
                Else
                    .ColHidden(mSpecColumn.���_�������) = True
                End If
            End If
            .ColHidden(mSpecColumn.���_�ɱ��۸�) = IIf(intPageItem = mSpecList.�۸���Ϣ, False, True)
            .ColHidden(mSpecColumn.���_��ǰ�ۼ�) = IIf(intPageItem = mSpecList.�۸���Ϣ, False, True)
            'ҩ������
            .ColHidden(mSpecColumn.���_������Ŀ) = IIf(intPageItem = mSpecList.ҩ������, False, True)
            .ColHidden(mSpecColumn.���_������Ŀ) = IIf(intPageItem = mSpecList.ҩ������, False, True)
            .ColHidden(mSpecColumn.���_����ѱ���) = IIf(intPageItem = mSpecList.ҩ������, False, True)
            .ColHidden(mSpecColumn.���_ҩ�ۼ���) = IIf(intPageItem = mSpecList.ҩ������, False, True)
            .ColHidden(mSpecColumn.���_���ηѱ�) = IIf(intPageItem = mSpecList.ҩ������, False, True)
            .ColHidden(mSpecColumn.���_��ֵ˰��) = IIf(intPageItem = mSpecList.ҩ������, False, True)
            .ColHidden(mSpecColumn.���_ҽ������) = IIf(intPageItem = mSpecList.ҩ������, False, True)
            '��������
            .ColHidden(mSpecColumn.���_ҩ�����) = IIf(intPageItem = mSpecList.��������, False, True)
            .ColHidden(mSpecColumn.���_ҩ������) = IIf(intPageItem = mSpecList.��������, False, True)
            .ColHidden(mSpecColumn.���_������) = IIf(intPageItem = mSpecList.��������, False, True)
            
            If mstrNode Like "�в�ҩ*" Then
                If intPageItem = mSpecList.�������� Then
                    .ColHidden(mSpecColumn.���_������) = True
                End If
            Else
                If intPageItem = mSpecList.�������� Then
                    .ColHidden(mSpecColumn.���_������) = False
                End If
            End If
            
            '�ٴ�Ӧ��
            .ColHidden(mSpecColumn.���_��ʶ˵��) = IIf(intPageItem = mSpecList.�ٴ�Ӧ��, False, True)
            .ColHidden(mSpecColumn.���_��ҩ����) = IIf(intPageItem = mSpecList.�ٴ�Ӧ��, False, True)
            .ColHidden(mSpecColumn.���_վ����) = IIf(intPageItem = mSpecList.�ٴ�Ӧ��, False, True)
            .ColHidden(mSpecColumn.���_DDDֵ) = IIf(intPageItem = mSpecList.�ٴ�Ӧ��, False, True)
            .ColHidden(mSpecColumn.���_�������) = IIf(intPageItem = mSpecList.�ٴ�Ӧ��, False, True)
            .ColHidden(mSpecColumn.���_סԺ����ʹ��) = IIf(intPageItem = mSpecList.�ٴ�Ӧ��, False, True)
            .ColHidden(mSpecColumn.���_�������ʹ��) = IIf(intPageItem = mSpecList.�ٴ�Ӧ��, False, True)
            .ColHidden(mSpecColumn.���_����ҩ��) = IIf(intPageItem = mSpecList.�ٴ�Ӧ��, False, True)
            .ColHidden(mSpecColumn.���_סԺ��̬����) = IIf(intPageItem = mSpecList.�ٴ�Ӧ��, False, True)
            .ColHidden(mSpecColumn.���_��ΣҩƷ) = IIf(intPageItem = mSpecList.�ٴ�Ӧ��, False, True)
            If mstrNode Like "�в�ҩ*" Then
                .ColHidden(mSpecColumn.���_����ҩ��) = True
                .ColHidden(mSpecColumn.���_סԺ��̬����) = True
                .ColHidden(mSpecColumn.���_��ΣҩƷ) = True
            Else
                If intPageItem = mSpecList.�ٴ�Ӧ�� Then
                    .ColHidden(mSpecColumn.���_����ҩ��) = False
                    .ColHidden(mSpecColumn.���_סԺ��̬����) = False
                    .ColHidden(mSpecColumn.���_��ΣҩƷ) = False
                End If
            End If
            
            '��ҩ����
            .ColHidden(mSpecColumn.���_�洢�¶�) = IIf(intPageItem = mSpecList.��ҩ����, False, True)
            .ColHidden(mSpecColumn.���_�洢����) = IIf(intPageItem = mSpecList.��ҩ����, False, True)
            .ColHidden(mSpecColumn.���_��ҩ����) = IIf(intPageItem = mSpecList.��ҩ����, False, True)
            .ColHidden(mSpecColumn.���_�������) = IIf(intPageItem = mSpecList.��ҩ����, False, True)
            .ColHidden(mSpecColumn.���_��Һע������) = IIf(intPageItem = mSpecList.��ҩ����, False, True)
            
            If mstrNode Like "�в�ҩ*" Then
                If intPageItem = mSpecList.��ҩ���� Then
                    tbcDetails.Item(mSpecList.������Ϣ).Selected = True
                End If
                tbcDetails.Item(mSpecList.��ҩ����).Visible = False
                .ColHidden(mSpecColumn.���_�洢�¶�) = True
                .ColHidden(mSpecColumn.���_�洢����) = True
                .ColHidden(mSpecColumn.���_��ҩ����) = True
                .ColHidden(mSpecColumn.���_�������) = True
                .ColHidden(mSpecColumn.���_��Һע������) = True
            Else
                tbcDetails.Item(mSpecList.��ҩ����).Visible = True
                If intPageItem = mSpecList.��ҩ���� Then
                    .ColHidden(mSpecColumn.���_�洢�¶�) = False
                    .ColHidden(mSpecColumn.���_�洢����) = False
                    .ColHidden(mSpecColumn.���_��ҩ����) = False
                    .ColHidden(mSpecColumn.���_�������) = False
                    .ColHidden(mSpecColumn.���_��Һע������) = False
                Else
                    .ColHidden(mSpecColumn.���_�洢�¶�) = True
                    .ColHidden(mSpecColumn.���_�洢����) = True
                    .ColHidden(mSpecColumn.���_��ҩ����) = True
                    .ColHidden(mSpecColumn.���_�������) = True
                    .ColHidden(mSpecColumn.���_��Һע������) = True
                End If
            End If
        End If
    End With
End Sub

Private Sub initColumn_Ʒ����Ϣ()
    '��ʼ��������Ϣҳ��
    Dim rsRecord As ADODB.Recordset
    Dim strTemp As String
    Dim i As Integer
    
    With vsfDetails
        .Cols = mVaricolumn.Ʒ��_Count
        .Rows = 1
        '������Ϣ
        VsfGridColFormat vsfDetails, mVaricolumn.Ʒ��_���, "���", 600, flexAlignCenterCenter, "���"
        VsfGridColFormat vsfDetails, mVaricolumn.Ʒ��_id, "id", 300, flexAlignCenterCenter, "id"
        VsfGridColFormat vsfDetails, mVaricolumn.Ʒ��_����id, "����id", 300, flexAlignCenterCenter, "����id"
        VsfGridColFormat vsfDetails, mVaricolumn.Ʒ��_ҩƷ����, "ҩƷ����", 2000, flexAlignLeftCenter, "ҩƷ����"
        VsfGridColFormat vsfDetails, mVaricolumn.Ʒ��_ҩƷ����, "ҩƷ����", 1000, flexAlignLeftCenter, "ҩƷ����"
        VsfGridColFormat vsfDetails, mVaricolumn.Ʒ��_ͨ������, "ͨ������", 1000, flexAlignLeftCenter, "ͨ������"
        VsfGridColFormat vsfDetails, mVaricolumn.Ʒ��_Ӣ������, "Ӣ������", 1000, flexAlignLeftCenter, "Ӣ������"
        VsfGridColFormat vsfDetails, mVaricolumn.Ʒ��_ƴ����, "ƴ����", 1000, flexAlignLeftCenter, "ƴ����"
        VsfGridColFormat vsfDetails, mVaricolumn.Ʒ��_�����, "�����", 1000, flexAlignLeftCenter, "�����"
        'Ʒ������
        VsfGridColFormat vsfDetails, mVaricolumn.Ʒ��_�������, "�������", 1000, flexAlignLeftCenter, "�������"
        VsfGridColFormat vsfDetails, mVaricolumn.Ʒ��_��ֵ����, "��ֵ����", 1000, flexAlignLeftCenter, "��ֵ����"
        VsfGridColFormat vsfDetails, mVaricolumn.Ʒ��_��Դ���, "��Դ���", 1000, flexAlignLeftCenter, "��Դ���"
        VsfGridColFormat vsfDetails, mVaricolumn.Ʒ��_��ҩ�ݴ�, "��ҩ�ݴ�", 1000, flexAlignLeftCenter, "��ҩ�ݴ�"
        VsfGridColFormat vsfDetails, mVaricolumn.Ʒ��_ҩƷ����, "ҩƷ����", 1000, flexAlignLeftCenter, "ҩƷ����"
        VsfGridColFormat vsfDetails, mVaricolumn.Ʒ��_����, "����", 2000, flexAlignLeftCenter, "����"
        VsfGridColFormat vsfDetails, mVaricolumn.Ʒ��_ԭ��ҩ, "ԭ��ҩ", 800, flexAlignCenterCenter, "ԭ��ҩ"
        VsfGridColFormat vsfDetails, mVaricolumn.Ʒ��_ר��ҩ, "ר��ҩ", 800, flexAlignCenterCenter, "ר��ҩ"
        VsfGridColFormat vsfDetails, mVaricolumn.Ʒ��_��������, "��������", 1000, flexAlignCenterCenter, "��������"
        VsfGridColFormat vsfDetails, mVaricolumn.Ʒ��_����ҩ, "����ҩ", 800, flexAlignCenterCenter, "����ҩ"
        VsfGridColFormat vsfDetails, mVaricolumn.Ʒ��_��ҩ, "��ҩ", 800, flexAlignRightCenter, "��ҩ"
        VsfGridColFormat vsfDetails, mVaricolumn.Ʒ��_ԭ��ҩ, "ԭ��ҩ", 1000, flexAlignLeftCenter, "ԭ��ҩ"
        VsfGridColFormat vsfDetails, mVaricolumn.Ʒ��_��ζʹ��, "��ζʹ��", 1000, flexAlignLeftCenter, "��ζʹ��"
        VsfGridColFormat vsfDetails, mVaricolumn.Ʒ��_������ҩ, "������ҩ", 1000, flexAlignCenterCenter, "������ҩ"
        VsfGridColFormat vsfDetails, mVaricolumn.Ʒ��_����ҩ, "����ҩ", 1000, flexAlignLeftCenter, "����ҩ"
        VsfGridColFormat vsfDetails, mVaricolumn.Ʒ��_��ý, "��ý", 1000, flexAlignCenterCenter, "��ý"
        '�ٴ�Ӧ��
        VsfGridColFormat vsfDetails, mVaricolumn.Ʒ��_�ο���Ŀ, "�ο���Ŀ", 1000, flexAlignLeftCenter, "�ο���Ŀ"
        VsfGridColFormat vsfDetails, mVaricolumn.Ʒ��_����ְ��, "����ְ��", 1000, flexAlignLeftCenter, "����ְ��"
        VsfGridColFormat vsfDetails, mVaricolumn.Ʒ��_ҽ��ְ��, "ҽ��ְ��", 1000, flexAlignLeftCenter, "ҽ��ְ��"
        VsfGridColFormat vsfDetails, mVaricolumn.Ʒ��_��������, "��������", 1000, flexAlignRightCenter, "��������"
        VsfGridColFormat vsfDetails, mVaricolumn.Ʒ��_�����Ա�, "ʹ���Ա�", 1500, flexAlignLeftCenter, "ʹ���Ա�"
        VsfGridColFormat vsfDetails, mVaricolumn.Ʒ��_������λ, "������λ", 1000, flexAlignLeftCenter, "������λ"
        VsfGridColFormat vsfDetails, mVaricolumn.Ʒ��_Ƥ��, "Ƥ��", 800, flexAlignCenterCenter, "Ƥ��"
        VsfGridColFormat vsfDetails, mVaricolumn.Ʒ��_������, "������", 1500, flexAlignLeftCenter, "������"
        VsfGridColFormat vsfDetails, mVaricolumn.Ʒ��_ATCCODE, "ATCCODE", 1500, flexAlignLeftCenter, "ATCCODE"
        VsfGridColFormat vsfDetails, mVaricolumn.Ʒ��_Ʒ���³���ҽ��, "Ʒ���³���ҽ��", 1500, flexAlignLeftCenter, "Ʒ���³���ҽ��"
        VsfGridColFormat vsfDetails, mVaricolumn.Ʒ��_�ο���ĿID, "�ο���Ŀid", 10, flexAlignLeftCenter, "�ο���Ŀid"
        
        If chkAllDetails.Value = 1 Then
            .ColWidth(mVaricolumn.Ʒ��_ҩƷ����) = 2000
        Else
            .ColHidden(mVaricolumn.Ʒ��_ҩƷ����) = True
        End If
    End With
    
    With vsfDetails
        'ԭ��ҩ
        .ColDataType(mVaricolumn.Ʒ��_ԭ��ҩ) = flexDTBoolean
        'ר��ҩ
        .ColDataType(mVaricolumn.Ʒ��_ר��ҩ) = flexDTBoolean
        '��������
        .ColDataType(mVaricolumn.Ʒ��_��������) = flexDTBoolean
        '����ҩ
        .ColDataType(mVaricolumn.Ʒ��_����ҩ) = flexDTBoolean
        '��ҩ
        .ColDataType(mVaricolumn.Ʒ��_��ҩ) = flexDTBoolean
        '������ҩ
        .ColDataType(mVaricolumn.Ʒ��_������ҩ) = flexDTBoolean
        'ԭ��ҩ
        .ColDataType(mVaricolumn.Ʒ��_ԭ��ҩ) = flexDTBoolean
        '����ҩ
        .ColDataType(mVaricolumn.Ʒ��_����ҩ) = flexDTBoolean
        '��ý
        .ColDataType(mVaricolumn.Ʒ��_��ý) = flexDTBoolean
        'Ƥ��
        .ColDataType(mVaricolumn.Ʒ��_Ƥ��) = flexDTBoolean
        'ҩƷ��Ʒ���³���ҽ��
        .ColDataType(mVaricolumn.Ʒ��_Ʒ���³���ҽ��) = flexDTBoolean
        '������
        .ColComboList(mVaricolumn.Ʒ��_������) = "0-�ǿ�����|1-������ʹ��|2-����ʹ��|3-����ʹ��"
        '��ζʹ��
        .ColDataType(mVaricolumn.Ʒ��_��ζʹ��) = flexDTBoolean
    
        '������λ
        gstrSql = "select distinct ���㵥λ from ������ĿĿ¼ where ���  in ('5','6','7')"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_Ʒ����Ϣ")
        If Not rsRecord.EOF Then
            For i = 1 To rsRecord.RecordCount
                strTemp = strTemp & "|" & rsRecord!���㵥λ
                rsRecord.MoveNext
            Next
        End If
        .ColComboList(mVaricolumn.Ʒ��_������λ) = strTemp
        
        '����
        gstrSql = "select ����||'-'|| ���� as ���� from ҩƷ����"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_Ʒ����Ϣ")
        .ColComboList(mVaricolumn.Ʒ��_����) = vsfDetails.BuildComboList(rsRecord, "����")
        '�ο���Ŀ
        .ColComboList(mVaricolumn.Ʒ��_�ο���Ŀ) = "|..."
        '�������
        gstrSql = "select ����||'-'|| ���� as ������� from ҩƷ�������"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_Ʒ����Ϣ")
        .ColComboList(mVaricolumn.Ʒ��_�������) = vsfDetails.BuildComboList(rsRecord, "�������")
        '��ֵ����
        gstrSql = "select ����||'-'|| ���� as ��ֵ���� from ҩƷ��ֵ����"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_Ʒ����Ϣ")
        .ColComboList(mVaricolumn.Ʒ��_��ֵ����) = vsfDetails.BuildComboList(rsRecord, "��ֵ����")
        '��Դ���
        gstrSql = "select ����||'-'|| ���� as ��Դ��� from ҩƷ��Դ���"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_Ʒ����Ϣ")
        .ColComboList(mVaricolumn.Ʒ��_��Դ���) = vsfDetails.BuildComboList(rsRecord, "��Դ���")
        '��ҩ�ݴ�
        gstrSql = "select ����||'-'|| ���� as ��ҩ�ݴ� from ҩƷ��ҩ�ݴ�"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_Ʒ����Ϣ")
        .ColComboList(mVaricolumn.Ʒ��_��ҩ�ݴ�) = vsfDetails.BuildComboList(rsRecord, "��ҩ�ݴ�")
        'ҩƷ����
        .ColComboList(mVaricolumn.Ʒ��_ҩƷ����) = "0-δ�趨|1-����ҩ|2-����Ǵ���ҩ|3-����Ǵ���ҩ|4-�Ǵ���ҩ|5-������ҩ"
        '����ְ��
        .ColComboList(mVaricolumn.Ʒ��_����ְ��) = "0-����|1-����|2-����|3-�м�|4-����/ʦ��|5-Ա/ʿ|9-��Ƹ"
        'ҽ��ְ��
        .ColComboList(mVaricolumn.Ʒ��_ҽ��ְ��) = "0-����|1-����|2-����|3-�м�|4-����/ʦ��|5-Ա/ʿ|9-��Ƹ"
        '�����Ա�
        .ColComboList(mVaricolumn.Ʒ��_�����Ա�) = "0-���Ա�����|1-����|2-Ů��"
        
    End With
End Sub

Private Sub initColumn_�����Ϣ()
    Dim rsRecord As ADODB.Recordset
    
    '��ʼ�������
    On Error GoTo ErrHandle
    With vsfDetails
        .Cols = mSpecColumn.���_count
        .Rows = 1
        '������Ϣ
        VsfGridColFormat vsfDetails, mSpecColumn.���_���, "���", 600, flexAlignCenterCenter, "���"
        VsfGridColFormat vsfDetails, mSpecColumn.���_id, "id", 300, flexAlignLeftCenter, "id"
        VsfGridColFormat vsfDetails, mSpecColumn.���_ҩ��id, "ҩ��id", 600, flexAlignCenterCenter, "ҩ��id"
'        VsfGridColFormat vsfDetails, mSpecColumn.���_ҩƷ����, "ҩƷ����", 300, flexAlignLeftCenter, "ҩƷ����"
        VsfGridColFormat vsfDetails, mSpecColumn.���_ͨ������, "ͨ������", 1000, flexAlignLeftCenter, "ͨ������"
        VsfGridColFormat vsfDetails, mSpecColumn.���_������, "������", 1000, flexAlignLeftCenter, "������"
        VsfGridColFormat vsfDetails, mSpecColumn.���_ҩƷ���, "ҩƷ���", 1500, flexAlignLeftCenter, "ҩƷ���"
        VsfGridColFormat vsfDetails, mSpecColumn.���_��λ��, "��λ��", 2500, flexAlignLeftCenter, "��λ��"
        VsfGridColFormat vsfDetails, mSpecColumn.���_������, "������", 1000, flexAlignLeftCenter, "������"
        VsfGridColFormat vsfDetails, mSpecColumn.���_��ʶ��, "��ʶ��", 1000, flexAlignLeftCenter, "��ʶ��"
        VsfGridColFormat vsfDetails, mSpecColumn.���_��ѡ��, "��ѡ��", 1000, flexAlignLeftCenter, "��ѡ��"
        VsfGridColFormat vsfDetails, mSpecColumn.���_����, "����", 800, flexAlignRightCenter, "����"
        '��Ʒ��Ϣ
        VsfGridColFormat vsfDetails, mSpecColumn.���_��Ʒ����, "��Ʒ����", 1500, flexAlignLeftCenter, "��Ʒ����"
        VsfGridColFormat vsfDetails, mSpecColumn.���_��������, "��������", 1500, flexAlignLeftCenter, "��������"
        VsfGridColFormat vsfDetails, mSpecColumn.���_��Դ����, "��Դ����", 1000, flexAlignLeftCenter, "��Դ����"
        VsfGridColFormat vsfDetails, mSpecColumn.���_ƴ����, "ƴ����", 1000, flexAlignLeftCenter, "ƴ����"
        VsfGridColFormat vsfDetails, mSpecColumn.���_�����, "�����", 1000, flexAlignLeftCenter, "�����"
        VsfGridColFormat vsfDetails, mSpecColumn.���_��ͬ��λ, "��ͬ��λ", 1000, flexAlignLeftCenter, "��ͬ��λ"
        VsfGridColFormat vsfDetails, mSpecColumn.���_��׼�ĺ�, "��׼�ĺ�", 1000, flexAlignLeftCenter, "��׼�ĺ�"
        VsfGridColFormat vsfDetails, mSpecColumn.���_ע���̱�, "ע���̱�", 1000, flexAlignLeftCenter, "ע���̱�"
        VsfGridColFormat vsfDetails, mSpecColumn.���_GMP��֤, "GMP��֤", 800, flexAlignCenterCenter, "GMP��֤"
        VsfGridColFormat vsfDetails, mSpecColumn.���_�ǳ���ҩ, "�ǳ���ҩ", 800, flexAlignCenterCenter, "�ǳ���ҩ"
        '��װ��λ
        VsfGridColFormat vsfDetails, mSpecColumn.���_�ۼ۵�λ, "�ۼ۵�λ", 1000, flexAlignLeftCenter, "�ۼ۵�λ"
        VsfGridColFormat vsfDetails, mSpecColumn.���_����ϵ��, "����ϵ��", 1000, flexAlignRightCenter, "����ϵ��"
        VsfGridColFormat vsfDetails, mSpecColumn.���_������λ, "������λ", 1000, flexAlignRightCenter, "������λ"
        VsfGridColFormat vsfDetails, mSpecColumn.���_סԺ��λ, "סԺ��λ", 1000, flexAlignLeftCenter, "סԺ��λ"
        VsfGridColFormat vsfDetails, mSpecColumn.���_סԺϵ��, "סԺϵ��", 1000, flexAlignRightCenter, "סԺϵ��"
        VsfGridColFormat vsfDetails, mSpecColumn.���_���ﵥλ, "���ﵥλ", 1000, flexAlignLeftCenter, "���ﵥλ"
        VsfGridColFormat vsfDetails, mSpecColumn.���_����ϵ��, "����ϵ��", 1000, flexAlignRightCenter, "����ϵ��"
        VsfGridColFormat vsfDetails, mSpecColumn.���_ҩ�ⵥλ, "ҩ�ⵥλ", 1000, flexAlignLeftCenter, "ҩ�ⵥλ"
        VsfGridColFormat vsfDetails, mSpecColumn.���_ҩ��ϵ��, "ҩ��ϵ��", 1000, flexAlignRightCenter, "ҩ��ϵ��"
        VsfGridColFormat vsfDetails, mSpecColumn.���_�ͻ���λ, "�ͻ���λ", 1000, flexAlignRightCenter, "�ͻ���λ"
        VsfGridColFormat vsfDetails, mSpecColumn.���_�ͻ���װ, "�ͻ���װ", 1000, flexAlignRightCenter, "�ͻ���װ"
        VsfGridColFormat vsfDetails, mSpecColumn.���_���쵥λ, "���쵥λ", 1000, flexAlignLeftCenter, "���쵥λ"
        VsfGridColFormat vsfDetails, mSpecColumn.���_���췧ֵ, "���췧ֵ", 1000, flexAlignRightCenter, "���췧ֵ"
        VsfGridColFormat vsfDetails, mSpecColumn.���_��ҩ��̬, "��ҩ��̬", 1000, flexAlignRightCenter, "��ҩ��̬"
        '�۸���Ϣ
        VsfGridColFormat vsfDetails, mSpecColumn.���_ҩ������, "ҩ������", 900, flexAlignLeftCenter, "ҩ������"
        VsfGridColFormat vsfDetails, mSpecColumn.���_�ɹ��޼�, "�ɹ��޼�", 1000, flexAlignRightCenter, "�ɹ��޼�"
        VsfGridColFormat vsfDetails, mSpecColumn.���_�ɹ�����, "�ɹ�����", 1000, flexAlignRightCenter, "�ɹ�����"
        VsfGridColFormat vsfDetails, mSpecColumn.���_�����, "�����", 1000, flexAlignRightCenter, "�����"
        VsfGridColFormat vsfDetails, mSpecColumn.���_ָ���ۼ�, "ָ���ۼ�", 1000, flexAlignRightCenter, "ָ���ۼ�"
        VsfGridColFormat vsfDetails, mSpecColumn.���_ָ������, "ָ������", 0, flexAlignRightCenter, "ָ������"
        VsfGridColFormat vsfDetails, mSpecColumn.���_�ӳ���, "�ӳ���", 1000, flexAlignRightCenter, "�ӳ���"
        VsfGridColFormat vsfDetails, mSpecColumn.���_�������, "�������", 1000, flexAlignRightCenter, "�������"
        VsfGridColFormat vsfDetails, mSpecColumn.���_�ɱ��۸�, "�ɱ��۸�", 1000, flexAlignRightCenter, "�ɱ��۸�"
        VsfGridColFormat vsfDetails, mSpecColumn.���_��ǰ�ۼ�, "��ǰ�ۼ�", 1000, flexAlignRightCenter, "��ǰ�ۼ�"
        'ҩ������
        VsfGridColFormat vsfDetails, mSpecColumn.���_������Ŀ, "������Ŀ", 1500, flexAlignLeftCenter, "������Ŀ"
        VsfGridColFormat vsfDetails, mSpecColumn.���_������Ŀ, "������Ŀ", 1000, flexAlignLeftCenter, "������Ŀ"
        VsfGridColFormat vsfDetails, mSpecColumn.���_����ѱ���, "����ѱ���", 1000, flexAlignRightCenter, "����ѱ���"
        VsfGridColFormat vsfDetails, mSpecColumn.���_ҩ�ۼ���, "ҩ�ۼ���", 1000, flexAlignLeftCenter, "ҩ�ۼ���"
        VsfGridColFormat vsfDetails, mSpecColumn.���_���ηѱ�, "���ηѱ�", 900, flexAlignLeftCenter, "���ηѱ�"
        VsfGridColFormat vsfDetails, mSpecColumn.���_��ֵ˰��, "��ֵ˰��", 1000, flexAlignRightCenter, "��ֵ˰��"
        VsfGridColFormat vsfDetails, mSpecColumn.���_ҽ������, "ҽ������", 1000, flexAlignLeftCenter, "ҽ������"
        '��������
        VsfGridColFormat vsfDetails, mSpecColumn.���_ҩ�����, "ҩ�����", 800, flexAlignCenterCenter, "ҩ�����"
        VsfGridColFormat vsfDetails, mSpecColumn.���_ҩ������, "ҩ������", 800, flexAlignCenterCenter, "ҩ������"
        VsfGridColFormat vsfDetails, mSpecColumn.���_ԭҩ�����, "ԭҩ�����", 800, flexAlignCenterCenter, "ԭҩ�����"
        VsfGridColFormat vsfDetails, mSpecColumn.���_ԭҩ������, "ԭҩ������", 800, flexAlignCenterCenter, "ԭҩ������"
        VsfGridColFormat vsfDetails, mSpecColumn.���_������, "������(��)", 1000, flexAlignRightCenter, "������(��)"
        '�ٴ�Ӧ��
        VsfGridColFormat vsfDetails, mSpecColumn.���_��ʶ˵��, "��ʶ˵��", 1000, flexAlignLeftCenter, "��ʶ˵��"
        VsfGridColFormat vsfDetails, mSpecColumn.���_��ҩ����, "��ҩ����", 900, flexAlignLeftCenter, "��ҩ����"
        VsfGridColFormat vsfDetails, mSpecColumn.���_վ����, "վ����", 900, flexAlignLeftCenter, "վ����"
        VsfGridColFormat vsfDetails, mSpecColumn.���_DDDֵ, "DDDֵ", 900, flexAlignLeftCenter, "DDDֵ"
        VsfGridColFormat vsfDetails, mSpecColumn.���_�������, "�������", 1500, flexAlignLeftCenter, "�������"
        VsfGridColFormat vsfDetails, mSpecColumn.���_סԺ����ʹ��, "סԺ����ʹ��", 1300, flexAlignLeftCenter, "סԺ����ʹ��"
        VsfGridColFormat vsfDetails, mSpecColumn.���_�������ʹ��, "�������ʹ��", 1300, flexAlignLeftCenter, "�������ʹ��"
        VsfGridColFormat vsfDetails, mSpecColumn.���_סԺ��̬����, "סԺ��̬����", 1300, flexAlignLeftCenter, "סԺ��̬����"
        VsfGridColFormat vsfDetails, mSpecColumn.���_����ҩ��, "����ҩ��", 1000, flexAlignLeftCenter, "����ҩ��"
        VsfGridColFormat vsfDetails, mSpecColumn.���_��ΣҩƷ, "��ΣҩƷ", 1000, flexAlignLeftCenter, "��ΣҩƷ"
        '��ҩ����
        VsfGridColFormat vsfDetails, mSpecColumn.���_�洢�¶�, "�洢�¶�", 1500, flexAlignLeftCenter, "�洢�¶�"
        VsfGridColFormat vsfDetails, mSpecColumn.���_�洢����, "�洢����", 1000, flexAlignLeftCenter, "�洢����"
        VsfGridColFormat vsfDetails, mSpecColumn.���_��ҩ����, "��ҩ����", 1500, flexAlignLeftCenter, "��ҩ����"
        VsfGridColFormat vsfDetails, mSpecColumn.���_�������, "�������", 1000, flexAlignLeftCenter, "�������"
        VsfGridColFormat vsfDetails, mSpecColumn.���_�б�ҩƷ, "�б�ҩƷ", 1000, flexAlignLeftCenter, "�б�ҩƷ"
        VsfGridColFormat vsfDetails, mSpecColumn.���_��ͬ��λid, "��ͬ��λid", 1000, flexAlignLeftCenter, "��ͬ��λid"
        VsfGridColFormat vsfDetails, mSpecColumn.���_������Ŀid, "������Ŀid", 1000, flexAlignLeftCenter, "������Ŀid"
        VsfGridColFormat vsfDetails, mSpecColumn.���_��Һע������, "��Һע������", 5000, flexAlignLeftCenter, "��Һע������"
    End With
    
    With vsfDetails
        '���ηѱ�
        .ColDataType(mSpecColumn.���_���ηѱ�) = flexDTBoolean
        'סԺ/���ﶯ̬����
        .ColDataType(mSpecColumn.���_סԺ��̬����) = flexDTBoolean
        'GMP��֤
        .ColDataType(mSpecColumn.���_GMP��֤) = flexDTBoolean
        '�ǳ���ҩ
        .ColDataType(mSpecColumn.���_�ǳ���ҩ) = flexDTBoolean
        'ҩ�����
        .ColDataType(mSpecColumn.���_ҩ�����) = flexDTBoolean
        'ҩ������
        .ColDataType(mSpecColumn.���_ҩ������) = flexDTBoolean
        '�洢����
        .ColDataType(mSpecColumn.���_�洢����) = flexDTBoolean
        '�������
        .ColDataType(mSpecColumn.���_�������) = flexDTBoolean
        
        '��������
        .ColComboList(mSpecColumn.���_��������) = "|..."
        '��Դ����
        gstrSql = "select ����||'-'|| ���� as ��Դ���� from ҩƷ��Դ����"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_Ʒ����Ϣ")
        .ColComboList(mSpecColumn.���_��Դ����) = vsfDetails.BuildComboList(rsRecord, "��Դ����")
        '��ͬ��λ
        .ColComboList(mSpecColumn.���_��ͬ��λ) = "|..."
        '��ҩ����
        gstrSql = "select ���� as ��ҩ���� from ��ҩ����"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_Ʒ����Ϣ")
        .ColComboList(mSpecColumn.���_��ҩ����) = vsfDetails.BuildComboList(rsRecord, "��ҩ����")
        'վ����
        gstrSql = "select ���||'-'||���� as վ���� from zlnodelist"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_Ʒ����Ϣ")
        .ColComboList(mSpecColumn.���_վ����) = vsfDetails.BuildComboList(rsRecord, "վ����")
        '���쵥λ
        .ColComboList(mSpecColumn.���_���쵥λ) = "�ۼ۵�λ|סԺ��λ|���ﵥλ|ҩ�ⵥλ"
        'ҩ������
        .ColComboList(mSpecColumn.���_ҩ������) = "0-����|1-ʱ��"
        '����ҩ��
        gstrSql = "Select ���� as ����ҩ��  From ����ҩ��˵��  Order By ����"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_Ʒ����Ϣ")
        .ColComboList(mSpecColumn.���_����ҩ��) = vsfDetails.BuildComboList(rsRecord, "����ҩ��")
        '������Ŀ
        gstrSql = "Select ID, '[' || ���� || ']' || ���� As ������Ŀ" & _
                  "  From ������Ŀ" & _
                  "  Where ĩ�� = 1 And (����ʱ�� Is Null Or ����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))" & _
                  "  Order By ����"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_Ʒ����Ϣ")
        .ColComboList(mSpecColumn.���_������Ŀ) = vsfDetails.BuildComboList(rsRecord, "������Ŀ")
        '������Ŀ
        .ColComboList(mSpecColumn.���_������Ŀ) = "..."
        'ҩ�۹�����
        gstrSql = "select ����||'-'||���� as ������ from ҩ�۹�����"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_Ʒ����Ϣ")
        .ColComboList(mSpecColumn.���_ҩ�ۼ���) = vsfDetails.BuildComboList(rsRecord, "������")
        'ҽ������
        gstrSql = "Select ����||'-'||���� as ҽ������ From �������� where ����=1 Order By ����"
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_Ʒ����Ϣ")
        .ColComboList(mSpecColumn.���_ҽ������) = vsfDetails.BuildComboList(rsRecord, "ҽ������")
        '�������
        .ColComboList(mSpecColumn.���_�������) = "0-��Ӧ���ڲ���|1-����|2-סԺ|3-�����סԺ"
        'סԺ/�������ʹ��
        .ColComboList(mSpecColumn.���_סԺ����ʹ��) = "0-���Է���|1-���ɷ���|2-һ����ʹ��|3-�����һ������Ч|4-�������������Ч|5-�������������Ч"
        .ColComboList(mSpecColumn.���_�������ʹ��) = "0-���Է���|1-���ɷ���|2-һ����ʹ��|3-�����һ������Ч|4-�������������Ч|5-�������������Ч"
        '�洢�¶�
        .ColComboList(mSpecColumn.���_�洢�¶�) = " |1-����(0-30��)|2-����(20������)|3-���(2-8��)"
        '��ҩ����
        gstrSql = "Select ����||'-'|| ���� as ��ҩ���� From ��Һ��ҩ���� "
        Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "initColumn_�����Ϣ")
        .ColComboList(mSpecColumn.���_��ҩ����) = " |" & vsfDetails.BuildComboList(rsRecord, "��ҩ����")
        '��ҩ��̬
        .ColComboList(mSpecColumn.���_��ҩ��̬) = "ɢװ|��ҩ��Ƭ|����"
        '��ΣҩƷ
        .ColComboList(mSpecColumn.���_��ΣҩƷ) = " |1-A��|2-B��|3-C��"
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function ReadAndSendDataToTvw(ByVal int״̬ As Integer) As Boolean
'���ܣ��������������ڵ�
'���� int״̬ �����жϽ������ʱ��Ʒ���޸Ļ��ǹ���޸�

    Dim NodeThis As Node
    Dim Intĩ�� As Integer
    Dim lng�ⷿID As Long
    Dim rs���ʷ��� As ADODB.Recordset
    Dim recdata As ADODB.Recordset
    
    'ҩƷ��;�����Ƿ�������
    ReadAndSendDataToTvw = False
    On Error GoTo ErrHandle
    gstrSql = " Select ����,���� From ������Ŀ��� " & _
              " Where Instr([1],����,1) > 0 " & _
              " Order by ����"
    Set rs���ʷ��� = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, "567")
    
    If rs���ʷ��� Is Nothing Then
        Exit Function
    End If
    
'    Set rs���ʷ��� = GetFilter����(rs���ʷ���)
    With tvwDetails
        .Nodes.Clear
        Do While Not rs���ʷ���.EOF
            .Nodes.Add , , "Root" & rs���ʷ���!����, rs���ʷ���!����, 1, 1
            .Nodes("Root" & rs���ʷ���!����).Tag = rs���ʷ���!����
            rs���ʷ���.MoveNext
        Loop
    End With
    
'    gstrSql = "Select Rownum As ID, ID As ��Ŀid, �ϼ�id, ����, ����, ����, ���" & _
'               " From (Select ID, �ϼ�id, ����, ����, Decode(����, 1, '����ҩ', 2, '�г�ҩ', 3, '�в�ҩ') ����, '����' As ���" & _
'               " From ���Ʒ���Ŀ¼" & _
'               " Where ���� In ('1', '2', '3') And Nvl(To_Char(����ʱ��, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01'" & _
'               " Start With �ϼ�id Is Null" & _
'               "  Connect By Prior ID = �ϼ�id " & _
'               " Union All " & _
'               " Select ID, ����id, ����, ����, Decode(���, 5, '����ҩ', 6, '�г�ҩ', 7, '�в�ҩ') ����, 'Ʒ��' As ���" & _
'               " From ������ĿĿ¼" & _
'               " Where ����id In (Select ID " & _
'               "               From ���Ʒ���Ŀ¼" & _
'               "               Where ���� In ('1', '2', '3') And Nvl(To_Char(����ʱ��, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01'" & _
'               "               Start With �ϼ�id Is Null " & _
'               "               Connect By Prior ID = �ϼ�id)) " & _
'        " Start  With �ϼ�id Is Null" & _
'        " Connect By Prior ID = �ϼ�id order by id,��Ŀid"

    gstrSql = "Select ID, �ϼ�id, ����, ����, Decode(����, 1, '����ҩ', 2, '�г�ҩ', 3, '�в�ҩ') ����, '����' As ���" & _
            " From ���Ʒ���Ŀ¼" & _
            " Where ���� In ('1', '2', '3') And Nvl(To_Char(����ʱ��, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01' " & _
            " Start With �ϼ�id Is Null" & _
            " Connect By Prior ID = �ϼ�id"

    Set recdata = zlDatabase.OpenSQLRecord(gstrSql, "ReadAndSendDataToTvw")
    
    If recdata.EOF Then
        MsgBox "���ʼ��ҩƷ��;���ࣨҩƷ��;���ࣩ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    With recdata
        Do While Not .EOF
            If IsNull(!�ϼ�ID) Then
                Set NodeThis = tvwDetails.Nodes.Add("Root" & !����, 4, "K_" & !ID, !����, 1, 1)
            Else
                Set NodeThis = tvwDetails.Nodes.Add("K_" & !�ϼ�ID, 4, "K_" & !ID, !����, 1, 1)
            End If
            NodeThis.Tag = !���� & "-" & !���  '��ŷ�������:1-����ҩ,2-�г�ҩ,3-�в�ҩ
            .MoveNext
        Loop
    End With
    
    If int״̬ <> 1 Then 'Ʒ���޸�
        gstrSql = "Select ID, ����id, ����, ����, Decode(���, 5, '����ҩ', 6, '�г�ҩ', 7, '�в�ҩ') ����, 'Ʒ��' As ���" & _
                  "  From ������ĿĿ¼" & _
                  "  Where ����id In (Select ID" & _
                                   " From ���Ʒ���Ŀ¼" & _
                                   " Where ���� In ('1', '2', '3') And Nvl(To_Char(����ʱ��, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01'" & _
                                   " Start With �ϼ�id Is Null" & _
                                   " Connect By Prior ID = �ϼ�id)"
        Set recdata = zlDatabase.OpenSQLRecord(gstrSql, "Ʒ��")
        
        With recdata
            Do While Not .EOF
                Set NodeThis = tvwDetails.Nodes.Add("K_" & !����id, 4, !��� & "K_" & !ID, !����, 1, 1)
                NodeThis.Tag = !���� & "-" & !���  '��ŷ�������:1-����ҩ,2-�г�ҩ,3-�в�ҩ
                .MoveNext
            Loop
        End With
    End If
        
    Call GetFilterȨ��  '�����û������е�Ȩ������������
    
    With tvwDetails
        If .Nodes.Count <> 0 Then
            .Nodes(1).Selected = True
            If .Nodes(1).Children <> 0 Then
                Intĩ�� = 1
                .Nodes(Intĩ��).Child.Selected = True
                .SelectedItem.Selected = True
            ElseIf .Nodes(2).Children <> 0 Then
                Intĩ�� = 2
                .Nodes(Intĩ��).Child.Selected = True
                .SelectedItem.Selected = True
            ElseIf .Nodes(3).Children <> 0 Then
                Intĩ�� = 3
                .Nodes(Intĩ��).Child.Selected = True
                .SelectedItem.Selected = True
            Else
                Intĩ�� = 0
                .Nodes(1).Selected = True
                .SelectedItem.Selected = True
            End If
            If Intĩ�� <> 0 Then .Nodes(Intĩ��).Expanded = True
        End If
    End With
    
    ReadAndSendDataToTvw = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub GetFilterȨ��()
    Dim strTemp As String
    
    With tvwDetails
        If mint״̬ = 1 Then
            If InStr(1, mstrPrivs, "��������ҩƷ��") = 0 Then
                .Nodes.Remove (.Nodes("Root����ҩ").Index)
            End If
            If InStr(1, mstrPrivs, "�����г�ҩƷ��") = 0 Then
                .Nodes.Remove (.Nodes("Root�г�ҩ").Index)
            End If
            If InStr(1, mstrPrivs, "�����в�ҩƷ��") = 0 Then
                .Nodes.Remove (.Nodes("Root�в�ҩ").Index)
            End If
        Else
            If InStr(1, mstrPrivs, "��������ҩ���") = 0 Then
                .Nodes.Remove (.Nodes("Root����ҩ").Index)
            End If
            If InStr(1, mstrPrivs, "�����г�ҩ���") = 0 Then
                .Nodes.Remove (.Nodes("Root�г�ҩ").Index)
            End If
            If InStr(1, mstrPrivs, "�����в�ҩ���") = 0 Then
                .Nodes.Remove (.Nodes("Root�в�ҩ").Index)
            End If
        End If
    End With
End Sub

Private Sub tvwDetails_NodeClick(ByVal Node As MSComctlLib.Node)
    '�ڵ����¼�
    Dim rsRecord As ADODB.Recordset
    Dim lngkey As Long  '����������ѡ�е�keyֵ
    Dim str���� As String   'ҩƷ����޸��������ж�ѡ�еĽڵ���Ʒ�ֻ��Ƿ���
    Dim intupdate As Integer
    Dim i As Integer
    Dim j As Integer
    Dim bln�޸� As Boolean  '������¼�Ƿ���ֵ���޸���
    
    If Node Is Nothing Then
        Exit Sub
    End If
    mstrNode = Node.Tag '��¼�ڵ��е�ֵ
    mblnClick = False
    
    On Error GoTo ErrHandle
    If Node.Tag Like "�в�ҩ*" And mint״̬ = 2 Then
        vsfDetails.ColComboList(mSpecColumn.���_סԺ����ʹ��) = "0-���Է���|1-���ɷ���"
        vsfDetails.ColComboList(mSpecColumn.���_�������ʹ��) = "0-���Է���|1-���ɷ���"
    ElseIf mint״̬ = 2 Then
        vsfDetails.ColComboList(mSpecColumn.���_סԺ����ʹ��) = "0-���Է���|1-���ɷ���|2-һ����ʹ��|3-�����һ������Ч|4-�������������Ч|5-�������������Ч"
        vsfDetails.ColComboList(mSpecColumn.���_�������ʹ��) = "0-���Է���|1-���ɷ���|2-һ����ʹ��|3-�����һ������Ч|4-�������������Ч|5-�������������Ч"
    End If
    If Node.Key Like "Root*" Then Exit Sub  '���ѡ��Ľڵ�ʱ����ڵ����˳�
'    If Node.Key = mstr�ϴνڵ� Then
'        Exit Sub
'    Else
'        mstr�ϴνڵ� = Node.Key
'    End If
    
    '�жϽ������Ƿ���ֵ�ձ��޸���
    bln�޸� = Check�޸�
    
    If bln�޸� = True Then
        intupdate = MsgBox("�������ݱ��޸��ˣ��Ƿ������", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName)
        If intupdate = vbNo Then Exit Sub
    End If
    
    If mint״̬ = 1 Then    'Ʒ��
        gstrSql = "Select Distinct a.id,c.id as ���id,a.�ο�Ŀ¼ID, '['||c.����||']'||c.���� as ����, a.����, a.���� As ͨ������, d.Ӣ����, d.ƴ����, d.�����, e.�������, e.��ֵ����, e.��Դ���, e.��ҩ�ݴ�, nvl(e.ҩƷ����,0) as ҩƷ����, e.ҩƷ����, nvl(e.����ҩ��,0)  as ����ҩ��," & _
                        "  e.�Ƿ�����ҩ as ����ҩ, e.��ý, e.ATCCODE, e.�Ƿ�ԭ��ҩ, e.�Ƿ�ר��ҩ, e.�Ƿ񵥶�����, nvl(e.�Ƿ���ҩ,0) as �Ƿ���ҩ, nvl(e.�Ƿ�ԭ��,0) as �Ƿ�ԭ��, f.���� As �ο���Ŀ, nvl(e.����ְ��,'00') as ����ְ��, nvl(e.��������,0) as ��������, Nvl(a.�����Ա�,0) AS �����Ա�, a.���㵥λ As ������λ, nvl(e.�Ƿ�Ƥ��,0) as �Ƿ�Ƥ��, nvl(e.������,0) as ������, nvl(e.Ʒ��ҽ��,0) as Ʒ��ҽ��,a.����Ӧ�� as ��ζʹ��,e.�Ƿ�����ҩ as ������ҩ" & _
                        "  From ������ĿĿ¼ A, ������Ŀ���� B, ���Ʒ���Ŀ¼ C," & _
                    " (Select n.������Ŀid, n.����, n.ƴ����, m.�����, p.Ӣ����" & _
                    "  From (Select ������Ŀid, ����, ���� As ƴ���� From ������Ŀ���� Where ���� = 1 And ���� = 1) N," & _
                    "       (Select ������Ŀid, ����, ���� As ����� From ������Ŀ���� Where ���� = 1 And ���� = 2) M," & _
                    "       (Select ������Ŀid, ���� As Ӣ���� From ������Ŀ���� Where ���� = 2) P" & _
                    "  Where n.������Ŀid = m.������Ŀid And n.������Ŀid = p.������Ŀid) D, ҩƷ���� E, ���Ʋο�Ŀ¼ F " & _
                    "   Where a.Id = b.������Ŀid(+) And a.����id = c.Id And a.Id = d.������Ŀid(+) And a.Id = e.ҩ��id And a.�ο�Ŀ¼id = f.Id(+) And " & _
                    " a.����ʱ�� = To_Date('3000-1-1', 'yyyy-MM-DD') And " & _
                    " a.����id "
        
        If chkAllDetails.Value = 1 Then '��ѡ������ʾ���нڵ��е�����ʱ
            gstrSql = gstrSql & " in (Select ID From ���Ʒ���Ŀ¼ Where ���� In (1, 2, 3) Start With ID = [1] Connect By Prior ID = �ϼ�id) order by id"
        Else
            gstrSql = gstrSql & " =[1] order by id"
        End If
    Else    '���
        str���� = Node.Tag
        If str���� Like "*Ʒ��" Then 'ѡ�е���Ʒ�ֽڵ�
            gstrSql = "Select a.Id, c.ҩ��id, a.���� As ������, a.���, j.���� As Ʒ�ֱ���, j.���� As ͨ������, m.������, c.��ʶ��, a.��ѡ��," & _
                              " Decode(n.��Ʒ��, Null, p.��Ʒ��, n.��Ʒ��) ��Ʒ��, a.���� As ��������, n.ƴ����, p.�����, c.ҩƷ��Դ As ��Դ����, d.���� As ��ͬ��λ, c.��׼�ĺ�, c.ע���̱�," & _
                              " c.Gmp��֤, c.�Ƿ񳣱�, a.���㵥λ As �ۼ۵�λ, c.����ϵ�� As �ۼ�ϵ��,j.���㵥λ, c.סԺ��λ, c.סԺ��װ, c.���ﵥλ, c.�����װ, c.ҩ�ⵥλ, c.ҩ���װ, c.���쵥λ, c.���췧ֵ," & _
                              " c.��ҩ��̬, a.�Ƿ��� As ҩ������, c.ָ�������� As �ɹ��޼�, c.���� As �ɹ�����, c.ָ�����ۼ� As ָ���ۼ�, c.ָ������� As ָ������, c.��������� As �������, c.�ɱ���," & _
                              " e.�ּ� As ��ǰ�ۼ�, f.���� As ������Ŀ,a.������Ŀ, c.����ѱ���,c.����, c.ҩ�ۼ���, a.���ηѱ�, c.��ֵ˰��, a.�������� As ҽ������, c.ҩ�����, c.ҩ������, c.�б�ҩƷ, c.��ͬ��λid," & _
                              " e.������Ŀid, c.���Ч�� As ������, a.˵�� As ��ʶ˵��, c.��ҩ����, a.�������, c.סԺ�ɷ����, c.��̬���� as סԺ��̬����,c.����ɷ����, c.����ҩ��, a.վ�� As վ����,C.dddֵ, i.�洢�¶�, i.�洢����," & _
                              " i.��ҩ����, i.�Ƿ������� As �������,C.��λ��,c.��ΣҩƷ, c.�ͻ���λ, c.�ͻ���װ,i.��Һע������ " & _
                       " From �շ���ĿĿ¼ A, (Select �շ�ϸĿid, ���� As ������ From �շ���Ŀ���� Where ���� = 3 And ���� = 1) M," & _
                            " (Select �շ�ϸĿid, ���� As ƴ����, ���� As ��Ʒ�� From �շ���Ŀ���� Where ���� = 1 And ���� = 3) N," & _
                            " (Select �շ�ϸĿid, ���� As �����, ���� As ��Ʒ�� From �շ���Ŀ���� Where ���� = 2 And ���� = 3) P, ҩƷ��� C, ������ĿĿ¼ J, ��Ӧ�� D, �շѼ�Ŀ E," & _
                            " ������Ŀ F, ��ҺҩƷ���� I" & _
                       " Where c.ҩ��id = j.Id And j.Id = [1] And a.����ʱ�� = To_Date('3000-1-1', 'yyyy-MM-DD') And a.Id = c.ҩƷid And" & _
                             " c.��ͬ��λid = d.Id(+) And e.�շ�ϸĿid = a.Id And e.������Ŀid = f.Id And a.Id = i.ҩƷid(+) And a.Id = m.�շ�ϸĿid(+) And" & _
                             " a.Id = n.�շ�ϸĿid(+) And a.Id = p.�շ�ϸĿid(+)  and (e.��ֹ���� is null or Sysdate Between e.ִ������ And e.��ֹ����)" & _
                       " Order By a.Id"

        Else    'ѡ�е��Ƿ���ڵ�
            gstrSql = " Select a.Id, c.ҩ��id, a.���� As ������, a.���, j.���� As Ʒ�ֱ���, j.���� As ͨ������, m.������, c.��ʶ��, a.��ѡ��," & _
                              " Decode(n.��Ʒ��, Null, p.��Ʒ��, n.��Ʒ��) ��Ʒ��, a.���� As ��������, n.ƴ����, p.�����, c.ҩƷ��Դ As ��Դ����, d.���� As ��ͬ��λ, c.��׼�ĺ�, c.ע���̱�, " & _
                              " c.Gmp��֤, c.�Ƿ񳣱�, a.���㵥λ As �ۼ۵�λ, c.����ϵ�� As �ۼ�ϵ��,j.���㵥λ, c.סԺ��λ, c.סԺ��װ, c.���ﵥλ, c.�����װ, c.ҩ�ⵥλ, c.ҩ���װ, c.���쵥λ, c.���췧ֵ," & _
                              " c.��ҩ��̬, a.�Ƿ��� As ҩ������, c.ָ�������� As �ɹ��޼�, c.���� As �ɹ�����, c.ָ�����ۼ� As ָ���ۼ�, c.ָ������� As ָ������, c.��������� As �������, c.�ɱ���," & _
                              " e.�ּ� As ��ǰ�ۼ�, f.���� As ������Ŀ,a.������Ŀ, c.����ѱ���, c.����,c.ҩ�ۼ���, a.���ηѱ�, c.��ֵ˰��, a.�������� As ҽ������, c.ҩ�����, c.ҩ������, c.�б�ҩƷ, ��ͬ��λid," & _
                              " e.������Ŀid, c.���Ч�� As ������, a.˵�� As ��ʶ˵��, c.��ҩ����, a.�������, c.סԺ�ɷ����, c.��̬���� as סԺ��̬����,c.����ɷ����,c.����ҩ��, a.վ�� As վ����,c.DDDֵ, i.�洢�¶�, i.�洢����," & _
                              " i.��ҩ����, i.�Ƿ������� As �������,C.��λ��,c.��ΣҩƷ, c.�ͻ���λ, c.�ͻ���װ,i.��Һע������" & _
                       " From �շ���ĿĿ¼ A, (Select �շ�ϸĿid, ���� As ������ From �շ���Ŀ���� Where ���� = 3 And ���� = 1) M," & _
                            " (Select �շ�ϸĿid, ���� As ƴ����, ���� As ��Ʒ�� From �շ���Ŀ���� Where ���� = 1 And ���� = 3) N," & _
                            " (Select �շ�ϸĿid, ���� As �����, ���� As ��Ʒ�� From �շ���Ŀ���� Where ���� = 2 And ���� = 3) P, ҩƷ��� C, ��Ӧ�� D, �շѼ�Ŀ E, ������Ŀ F," & _
                            " ��ҺҩƷ���� I, ������ĿĿ¼ J" & _
                       " Where a.Id In" & _
                            "  (Select ҩƷid" & _
                              " From ҩƷ���" & _
                              " Where ҩ��id In " & _
                                    " (Select ID " & _
                                    "  From ������ĿĿ¼ " & _
                                     " Where ����id In " & _
                                          "  (Select ID From ���Ʒ���Ŀ¼ Where ���� In (1, 2, 3) Start With ID = [1] Connect By Prior ID = �ϼ�id))) And" & _
                             " a.����ʱ�� = To_Date('3000-1-1', 'yyyy-MM-DD')" & _
                             " And a.Id = c.ҩƷid And c.��ͬ��λid = d.Id(+) And e.�շ�ϸĿid = a.Id And e.������Ŀid = f.Id And a.Id = i.ҩƷid(+) And" & _
                             " c.ҩ��id = j.Id And a.Id = m.�շ�ϸĿid(+) And a.Id = n.�շ�ϸĿid(+) And a.Id = p.�շ�ϸĿid(+) and (e.��ֹ���� is null or Sysdate Between e.ִ������ And e.��ֹ����)" & _
                       " Order By j.����,a.Id"

        End If
        Call setColumn(tbcDetails.Selected.Index)
        If chkAllDetails.Value = 0 Then '���ܻ�ȡ���¼��ڵ�
            If Node.Tag Like "*����" Then
                vsfDetails.Rows = 1
                Exit Sub
            End If
        End If
    End If

    If mint״̬ = 2 Then '���
        If Node.Tag Like "�в�ҩ*" Then  '�Ƿ���ʾ��ҩ����
            tbcDetails.Item(mSpecList.��ҩ����).Visible = False
            
            With vsfDetails
                .ColHidden(mSpecColumn.���_�洢�¶�) = True
                .ColHidden(mSpecColumn.���_�洢����) = True
                .ColHidden(mSpecColumn.���_��ҩ����) = True
                .ColHidden(mSpecColumn.���_�������) = True
                .ColHidden(mSpecColumn.���_��Һע������) = True
                If tbcDetails.Selected.Index = tbcDetails.ItemCount - 1 Then
                    tbcDetails.Item(mSpecList.������Ϣ).Selected = True
                End If
            End With
        Else
            tbcDetails.Item(mSpecList.��ҩ����).Visible = True
            With vsfDetails
                If tbcDetails.Item(mSpecList.��ҩ����).Selected = True Then
                    .ColHidden(mSpecColumn.���_�洢�¶�) = False
                    .ColHidden(mSpecColumn.���_�洢����) = False
                    .ColHidden(mSpecColumn.���_��ҩ����) = False
                    .ColHidden(mSpecColumn.���_�������) = False
                    .ColHidden(mSpecColumn.���_��Һע������) = False
                Else
                    .ColHidden(mSpecColumn.���_�洢�¶�) = True
                    .ColHidden(mSpecColumn.���_�洢����) = True
                    .ColHidden(mSpecColumn.���_��ҩ����) = True
                    .ColHidden(mSpecColumn.���_�������) = True
                    .ColHidden(mSpecColumn.���_��Һע������) = True
                End If
            End With
        End If
    End If
    '��ȡkeyֵ
    lngkey = Mid(Node.Key, InStr(1, Node.Key, "_") + 1, Len(Node.Key) - InStr(1, Node.Key, "_"))
    Set rsRecord = zlDatabase.OpenSQLRecord(gstrSql, "�ڵ���", lngkey)
    
    vsfDetails.Rows = 1
    If rsRecord.EOF Then
        Call setColumn(tbcDetails.Selected.Index)
        Exit Sub
    End If
    Set mrsRecord = rsRecord.Clone  '��¡
    
    Call showColumn(rsRecord, Node.Tag)   '��ֵ�󶨵�vsflexgrid�ؼ���
    Call setColumn(tbcDetails.Selected.Index)
    Call GetDefineSize(rsRecord)
    With vsfDetails
        If .Rows > 1 Then
            .Row = 1
            .Col = mVaricolumn.Ʒ��_ͨ������
        End If
    End With
    Call SetȨ���ж�
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub showColumn(ByVal rsRecord As ADODB.Recordset, ByVal str���� As String)
    '��������ڵ�ʱ����ֵ�󶨵�vsflexgrid�ؼ���
    Dim i As Integer
    Dim j As Integer
    Dim strTemp As String
    Dim intTemp As Integer
    Dim bln����ϵ�� As Boolean

    vsfDetails.Rows = rsRecord.RecordCount + 1 '���ݲ�ѯ������ֵ��������ȷ���б�����
    
    vsfDetails.Select 1, 1
    If mint״̬ = 1 Then    'Ʒ��
        For i = 1 To rsRecord.RecordCount
            With vsfDetails
                .TextMatrix(i, mVaricolumn.Ʒ��_���) = i
                .TextMatrix(i, mVaricolumn.Ʒ��_id) = IIf(IsNull(rsRecord!ID), "", rsRecord!ID)
                .TextMatrix(i, mVaricolumn.Ʒ��_����id) = IIf(IsNull(rsRecord!���id), "", rsRecord!���id)
                .TextMatrix(i, mVaricolumn.Ʒ��_ҩƷ����) = IIf(IsNull(rsRecord!����), "", rsRecord!����)
                .TextMatrix(i, mVaricolumn.Ʒ��_ҩƷ����) = IIf(IsNull(rsRecord!����), "", rsRecord!����)
                .TextMatrix(i, mVaricolumn.Ʒ��_ͨ������) = IIf(IsNull(rsRecord!ͨ������), "", rsRecord!ͨ������)
                .TextMatrix(i, mVaricolumn.Ʒ��_Ӣ������) = IIf(IsNull(rsRecord!Ӣ����), "", rsRecord!Ӣ����)
                .TextMatrix(i, mVaricolumn.Ʒ��_ƴ����) = IIf(IsNull(rsRecord!ƴ����), "", rsRecord!ƴ����)
                .TextMatrix(i, mVaricolumn.Ʒ��_�����) = IIf(IsNull(rsRecord!�����), "", rsRecord!�����)
                
                If .TextMatrix(i, mVaricolumn.Ʒ��_ƴ����) = "" Then
                    .TextMatrix(i, mVaricolumn.Ʒ��_ƴ����) = zlGetSymbol(.TextMatrix(i, mVaricolumn.Ʒ��_ͨ������), 0, 30)
                End If
                
                If .TextMatrix(i, mVaricolumn.Ʒ��_�����) = "" Then
                    .TextMatrix(i, mVaricolumn.Ʒ��_�����) = zlGetSymbol(.TextMatrix(i, mVaricolumn.Ʒ��_ͨ������), 1, 30)
                End If
                
                .TextMatrix(i, mVaricolumn.Ʒ��_�������) = ShowValue(.ColComboList(mVaricolumn.Ʒ��_�������), IIf(IsNull(rsRecord!�������), "", rsRecord!�������), "-")
                .TextMatrix(i, mVaricolumn.Ʒ��_��ֵ����) = ShowValue(.ColComboList(mVaricolumn.Ʒ��_��ֵ����), IIf(IsNull(rsRecord!��ֵ����), "", rsRecord!��ֵ����), "-")
                .TextMatrix(i, mVaricolumn.Ʒ��_��Դ���) = ShowValue(.ColComboList(mVaricolumn.Ʒ��_��Դ���), IIf(IsNull(rsRecord!��Դ���), "", rsRecord!��Դ���), "-")
                .TextMatrix(i, mVaricolumn.Ʒ��_��ҩ�ݴ�) = ShowValue(.ColComboList(mVaricolumn.Ʒ��_��ҩ�ݴ�), IIf(IsNull(rsRecord!��ҩ�ݴ�), "", rsRecord!��ҩ�ݴ�), "-")
                .TextMatrix(i, mVaricolumn.Ʒ��_ҩƷ����) = ShowValue(.ColComboList(mVaricolumn.Ʒ��_ҩƷ����), IIf(IsNull(rsRecord!ҩƷ����), "", rsRecord!ҩƷ����), "-", True)
                .TextMatrix(i, mVaricolumn.Ʒ��_����) = ShowValue(.ColComboList(mVaricolumn.Ʒ��_����), IIf(IsNull(rsRecord!ҩƷ����), "", rsRecord!ҩƷ����), "-")
                .TextMatrix(i, mVaricolumn.Ʒ��_ԭ��ҩ) = IIf(IsNull(rsRecord!�Ƿ�ԭ��ҩ), "", rsRecord!�Ƿ�ԭ��ҩ)
                .TextMatrix(i, mVaricolumn.Ʒ��_ר��ҩ) = IIf(IsNull(rsRecord!�Ƿ�ר��ҩ), "", rsRecord!�Ƿ�ר��ҩ)
                .TextMatrix(i, mVaricolumn.Ʒ��_��������) = IIf(IsNull(rsRecord!�Ƿ񵥶�����), "", rsRecord!�Ƿ񵥶�����)
                .TextMatrix(i, mVaricolumn.Ʒ��_����ҩ) = IIf(IsNull(rsRecord!����ҩ��), "", rsRecord!����ҩ��)
                .TextMatrix(i, mVaricolumn.Ʒ��_��ҩ) = IIf(IsNull(rsRecord!�Ƿ���ҩ), "", rsRecord!�Ƿ���ҩ)
                .TextMatrix(i, mVaricolumn.Ʒ��_ԭ��ҩ) = IIf(IsNull(rsRecord!�Ƿ�ԭ��), "", rsRecord!�Ƿ�ԭ��)
                .TextMatrix(i, mVaricolumn.Ʒ��_��ζʹ��) = IIf(IsNull(rsRecord!��ζʹ��), "", rsRecord!��ζʹ��)
                .TextMatrix(i, mVaricolumn.Ʒ��_������ҩ) = IIf(IsNull(rsRecord!������ҩ), "", rsRecord!������ҩ)
                .TextMatrix(i, mVaricolumn.Ʒ��_����ҩ) = IIf(IsNull(rsRecord!����ҩ), "", rsRecord!����ҩ)
                .TextMatrix(i, mVaricolumn.Ʒ��_��ý) = IIf(IsNull(rsRecord!��ý), "", rsRecord!��ý)
                .TextMatrix(i, mVaricolumn.Ʒ��_ATCCODE) = IIf(IsNull(rsRecord!ATCCODE), "", rsRecord!ATCCODE)
                .TextMatrix(i, mVaricolumn.Ʒ��_�ο���Ŀ) = IIf(IsNull(rsRecord!�ο���Ŀ), "", rsRecord!�ο���Ŀ)
                .TextMatrix(i, mVaricolumn.Ʒ��_����ְ��) = ShowValue(.ColComboList(mVaricolumn.Ʒ��_����ְ��), IIf(IsNull(Mid(rsRecord!����ְ��, 1, 1)), "", Mid(rsRecord!����ְ��, 1, 1)), "-", True)
                .TextMatrix(i, mVaricolumn.Ʒ��_ҽ��ְ��) = ShowValue(.ColComboList(mVaricolumn.Ʒ��_ҽ��ְ��), IIf(IsNull(Mid(rsRecord!����ְ��, 2, 1)), "", Mid(rsRecord!����ְ��, 2, 1)), "-", True)
                .TextMatrix(i, mVaricolumn.Ʒ��_��������) = IIf(IsNull(rsRecord!��������), "", rsRecord!��������)
                .TextMatrix(i, mVaricolumn.Ʒ��_�����Ա�) = ShowValue(.ColComboList(mVaricolumn.Ʒ��_�����Ա�), IIf(IsNull(rsRecord!�����Ա�), "0", rsRecord!�����Ա�), "-", True)
                .TextMatrix(i, mVaricolumn.Ʒ��_������λ) = IIf(IsNull(rsRecord!������λ), "", rsRecord!������λ)
                .TextMatrix(i, mVaricolumn.Ʒ��_Ƥ��) = IIf(IsNull(rsRecord!�Ƿ�Ƥ��), "", rsRecord!�Ƿ�Ƥ��)
                .TextMatrix(i, mVaricolumn.Ʒ��_������) = ShowValue(.ColComboList(mVaricolumn.Ʒ��_������), IIf(IsNull(rsRecord!������), "", rsRecord!������), "-", True)
                .TextMatrix(i, mVaricolumn.Ʒ��_Ʒ���³���ҽ��) = IIf(IsNull(rsRecord!Ʒ��ҽ��), "0", rsRecord!Ʒ��ҽ��)
                .TextMatrix(i, mVaricolumn.Ʒ��_�ο���ĿID) = IIf(IsNull(rsRecord!�ο�Ŀ¼ID), "", rsRecord!�ο�Ŀ¼ID)
                
                If rsRecord!������ = 0 Then
                    .Cell(flexcpBackColor, i, mVaricolumn.Ʒ��_ATCCODE, i) = mlngColor
                End If
            End With
'            Call ShowPercent(rsRecord.AbsolutePosition / rsRecord.RecordCount) '������
            
            rsRecord.MoveNext
        Next
        vsfDetails.Cell(flexcpBackColor, 1, mVaricolumn.Ʒ��_ҩƷ����, vsfDetails.Rows - 1) = mlngColor    '���ò��ɱ༭�еı�����ɫΪ��ɫ
        vsfDetails.Cell(flexcpBackColor, 1, mVaricolumn.Ʒ��_ҩƷ����, vsfDetails.Rows - 1) = mlngColor     '���ò��ɱ༭�еı�����ɫΪ��ɫ
        
'        With vsfDetails
'            If str���� Like "�в�ҩ*" Then
'                .ColHidden(mVaricolumn.Ʒ��_��ζʹ��) = False
'                .ColHidden(mVaricolumn.Ʒ��_����) = True
'                .ColHidden(mVaricolumn.Ʒ��_����ҩ) = True
'                .ColHidden(mVaricolumn.Ʒ��_��ҩ) = True
'                .ColHidden(mVaricolumn.Ʒ��_Ƥ��) = True
'                .ColHidden(mVaricolumn.Ʒ��_������) = True
'                .ColHidden(mVaricolumn.Ʒ��_Ʒ���³���ҽ��) = True
'            Else
'                .ColHidden(mVaricolumn.Ʒ��_��ζʹ��) = True
'                .ColHidden(mVaricolumn.Ʒ��_����) = False
'                .ColHidden(mVaricolumn.Ʒ��_����ҩ) = False
'                .ColHidden(mVaricolumn.Ʒ��_��ҩ) = False
'                .ColHidden(mVaricolumn.Ʒ��_Ƥ��) = False
'                .ColHidden(mVaricolumn.Ʒ��_������) = False
'                .ColHidden(mVaricolumn.Ʒ��_Ʒ���³���ҽ��) = False
'            End If
'        End With
        
        vsfDetails.MergeCol(mVaricolumn.Ʒ��_ҩƷ����) = True  '��ͬ���� ҩƷ������ͬ�ϲ�
    Else    '���
        For i = 1 To rsRecord.RecordCount
            With vsfDetails
                .TextMatrix(i, mSpecColumn.���_���) = i
                .TextMatrix(i, mSpecColumn.���_id) = IIf(IsNull(rsRecord!ID), "", rsRecord!ID)
                .TextMatrix(i, mSpecColumn.���_ҩ��id) = IIf(IsNull(rsRecord!ҩ��ID), "", rsRecord!ҩ��ID)
'                .TextMatrix(i, mSpecColumn.���_ҩƷ����) = IIf(IsNull(rsRecord!Ʒ�ֱ���), "", rsRecord!Ʒ�ֱ���)
                .TextMatrix(i, mSpecColumn.���_ͨ������) = IIf(IsNull(rsRecord!ͨ������), "", rsRecord!ͨ������)
                .TextMatrix(i, mSpecColumn.���_������) = IIf(IsNull(rsRecord!������), "", rsRecord!������)
                .TextMatrix(i, mSpecColumn.���_ҩƷ���) = IIf(IsNull(rsRecord!���), "", rsRecord!���)
                .TextMatrix(i, mSpecColumn.���_��λ��) = IIf(IsNull(rsRecord!��λ��), "", rsRecord!��λ��)
                .TextMatrix(i, mSpecColumn.���_������) = IIf(IsNull(rsRecord!������), "", rsRecord!������)
                
                If .TextMatrix(i, mSpecColumn.���_������) = "" And .TextMatrix(i, mSpecColumn.���_ҩƷ���) <> "" Then
                    .TextMatrix(i, mSpecColumn.���_������) = zlGetDigitSign(rsRecord!ҩ��ID, rsRecord!���)
                End If
                
                .TextMatrix(i, mSpecColumn.���_��ʶ��) = IIf(IsNull(rsRecord!��ʶ��), "", rsRecord!��ʶ��)
                .TextMatrix(i, mSpecColumn.���_��ѡ��) = IIf(IsNull(rsRecord!��ѡ��), "", rsRecord!��ѡ��)
                .TextMatrix(i, mSpecColumn.���_����) = IIf(IsNull(rsRecord!����), "", rsRecord!����)
                .TextMatrix(i, mSpecColumn.���_��Ʒ����) = IIf(IsNull(rsRecord!��Ʒ��), "", rsRecord!��Ʒ��)
                .TextMatrix(i, mSpecColumn.���_��������) = IIf(IsNull(rsRecord!��������), "", rsRecord!��������)
                .TextMatrix(i, mSpecColumn.���_��Դ����) = ShowValue(.ColComboList(mSpecColumn.���_��Դ����), IIf(IsNull(rsRecord!��Դ����), "", rsRecord!��Դ����), "-")
                .TextMatrix(i, mSpecColumn.���_ƴ����) = IIf(IsNull(rsRecord!ƴ����), "", rsRecord!ƴ����)
                .TextMatrix(i, mSpecColumn.���_�����) = IIf(IsNull(rsRecord!�����), "", rsRecord!�����)
                
                If .TextMatrix(i, mSpecColumn.���_��Ʒ����) <> "" And .TextMatrix(i, mSpecColumn.���_ƴ����) = "" Then
                    .TextMatrix(i, mSpecColumn.���_ƴ����) = zlGetSymbol(.TextMatrix(i, mSpecColumn.���_ͨ������), 0, 30)
                End If
                
                If .TextMatrix(i, mSpecColumn.���_��Ʒ����) <> "" And .TextMatrix(i, mSpecColumn.���_ƴ����) = "" Then
                    .TextMatrix(i, mSpecColumn.���_�����) = zlGetSymbol(.TextMatrix(i, mSpecColumn.���_ͨ������), 1, 30)
                End If
                
                .TextMatrix(i, mSpecColumn.���_��ͬ��λ) = IIf(IsNull(rsRecord!��ͬ��λ), "", rsRecord!��ͬ��λ)
                .TextMatrix(i, mSpecColumn.���_��׼�ĺ�) = IIf(IsNull(rsRecord!��׼�ĺ�), "", rsRecord!��׼�ĺ�)
                
                .TextMatrix(i, mSpecColumn.���_ע���̱�) = IIf(IsNull(rsRecord!ע���̱�), "", rsRecord!ע���̱�)
                .TextMatrix(i, mSpecColumn.���_GMP��֤) = IIf(IsNull(rsRecord!GMP��֤), "", rsRecord!GMP��֤)
                .TextMatrix(i, mSpecColumn.���_�ǳ���ҩ) = IIf(IsNull(rsRecord!�Ƿ񳣱�), "", rsRecord!�Ƿ񳣱�)
                .TextMatrix(i, mSpecColumn.���_�ۼ۵�λ) = IIf(IsNull(rsRecord!�ۼ۵�λ), "", rsRecord!�ۼ۵�λ)
                .TextMatrix(i, mSpecColumn.���_����ϵ��) = IIf(IsNull(rsRecord!�ۼ�ϵ��), "", rsRecord!�ۼ�ϵ��)
                .TextMatrix(i, mSpecColumn.���_������λ) = IIf(IsNull(rsRecord!���㵥λ), "", rsRecord!���㵥λ)
                .TextMatrix(i, mSpecColumn.���_סԺ��λ) = IIf(IsNull(rsRecord!סԺ��λ), "", rsRecord!סԺ��λ)
                .TextMatrix(i, mSpecColumn.���_סԺϵ��) = IIf(IsNull(rsRecord!סԺ��װ), "", rsRecord!סԺ��װ)
                .TextMatrix(i, mSpecColumn.���_���ﵥλ) = IIf(IsNull(rsRecord!���ﵥλ), "", rsRecord!���ﵥλ)
                .TextMatrix(i, mSpecColumn.���_����ϵ��) = IIf(IsNull(rsRecord!�����װ), "", rsRecord!�����װ)
                .TextMatrix(i, mSpecColumn.���_ҩ�ⵥλ) = IIf(IsNull(rsRecord!ҩ�ⵥλ), "", rsRecord!ҩ�ⵥλ)
                
                .TextMatrix(i, mSpecColumn.���_ҩ������) = ShowValue(.ColComboList(mSpecColumn.���_ҩ������), IIf(IsNull(rsRecord!ҩ������), "", rsRecord!ҩ������), "-", True)
                .TextMatrix(i, mSpecColumn.���_ҩ��ϵ��) = IIf(IsNull(rsRecord!ҩ���װ), "", rsRecord!ҩ���װ)
                .TextMatrix(i, mSpecColumn.���_�ͻ���λ) = IIf(IsNull(rsRecord!�ͻ���λ), "", rsRecord!�ͻ���λ)
                .TextMatrix(i, mSpecColumn.���_�ͻ���װ) = IIf(IsNull(rsRecord!�ͻ���װ), "", rsRecord!�ͻ���װ)
                Select Case rsRecord!��ҩ��̬
                    Case "0"
                        strTemp = "ɢװ"
                    Case "1"
                        strTemp = "��ҩ��Ƭ"
                    Case Else
                        strTemp = "����"
                End Select
                
                .TextMatrix(i, mSpecColumn.���_��ҩ��̬) = strTemp
                
                Select Case rsRecord!���쵥λ
                    Case "1"
                        strTemp = "�ۼ۵�λ"
                    Case "2"
                        strTemp = "סԺ��λ"
                    Case "3"
                        strTemp = "���ﵥλ"
                    Case "4"
                        strTemp = "ҩ�ⵥλ"
                    Case Else
                        strTemp = "�ۼ۵�λ"
                End Select
                .TextMatrix(i, mSpecColumn.���_���쵥λ) = strTemp
                
                Select Case Nvl(rsRecord!���쵥λ, 1)
                    Case 1 '����
                        .TextMatrix(i, mSpecColumn.���_���췧ֵ) = Format(Nvl(rsRecord!���췧ֵ, 0), "#0.00;-#0.00; ;")
                    Case 2 'סԺ
                        .TextMatrix(i, mSpecColumn.���_���췧ֵ) = Format(Nvl(rsRecord!���췧ֵ, 0) / Nvl(rsRecord!סԺ��װ, 1), "#0.00;-#0.00; ;")
                    Case 3 '����
                        .TextMatrix(i, mSpecColumn.���_���췧ֵ) = Format(Nvl(rsRecord!���췧ֵ, 0) / Nvl(rsRecord!�����װ, 1), "#0.00;-#0.00; ;")
                    Case 4 'ҩ��
                        .TextMatrix(i, mSpecColumn.���_���췧ֵ) = Format(Nvl(rsRecord!���췧ֵ, 0) / Nvl(rsRecord!ҩ���װ, 1), "#0.00;-#0.00; ;")
                End Select
                
                If mint��ǰ��λ <> 0 Then
                    .TextMatrix(i, mSpecColumn.���_�ɹ��޼�) = FormatEx(IIf(IsNull(rsRecord!�ɹ��޼�), 0, rsRecord!�ɹ��޼�) * .TextMatrix(i, mSpecColumn.���_ҩ��ϵ��), mintCostDigit)
                    .TextMatrix(i, mSpecColumn.���_ָ���ۼ�) = FormatEx(IIf(IsNull(rsRecord!ָ���ۼ�), 0, rsRecord!ָ���ۼ�) * .TextMatrix(i, mSpecColumn.���_ҩ��ϵ��), mintPriceDigit)
                    .TextMatrix(i, mSpecColumn.���_�ɱ��۸�) = FormatEx(IIf(IsNull(rsRecord!�ɱ���), "", rsRecord!�ɱ���) * .TextMatrix(i, mSpecColumn.���_ҩ��ϵ��), mintCostDigit)
                Else
                    .TextMatrix(i, mSpecColumn.���_�ɹ��޼�) = FormatEx(IIf(IsNull(rsRecord!�ɹ��޼�), 0, rsRecord!�ɹ��޼�), mintCostDigit)
                    .TextMatrix(i, mSpecColumn.���_ָ���ۼ�) = FormatEx(IIf(IsNull(rsRecord!ָ���ۼ�), 0, rsRecord!ָ���ۼ�), mintPriceDigit)
                    .TextMatrix(i, mSpecColumn.���_�ɱ��۸�) = FormatEx(IIf(IsNull(rsRecord!�ɱ���), "", rsRecord!�ɱ���), mintCostDigit)
                End If
                
                .TextMatrix(i, mSpecColumn.���_�ɹ�����) = IIf(IsNull(rsRecord!�ɹ�����), "", rsRecord!�ɹ�����)
                .TextMatrix(i, mSpecColumn.���_�����) = FormatEx(.TextMatrix(i, mSpecColumn.���_�ɹ��޼�) * (.TextMatrix(i, mSpecColumn.���_�ɹ�����) / 100), mintCostDigit)
                .TextMatrix(i, mSpecColumn.���_ָ������) = Format(IIf(IsNull(rsRecord!ָ������), "", rsRecord!ָ������), "0.00000")
                .TextMatrix(i, mSpecColumn.���_�ӳ���) = Format((1 / (1 - .TextMatrix(i, mSpecColumn.���_ָ������) / 100) - 1) * 100, "0.00")
                .TextMatrix(i, mSpecColumn.���_�������) = Format(IIf(IsNull(rsRecord!�������), "", rsRecord!�������), "0.00")
                
                If mint��ǰ��λ <> 0 Then
                    .TextMatrix(i, mSpecColumn.���_��ǰ�ۼ�) = FormatEx(IIf(IsNull(rsRecord!��ǰ�ۼ�), 0, rsRecord!��ǰ�ۼ�) * .TextMatrix(i, mSpecColumn.���_ҩ��ϵ��), mintPriceDigit)
                Else
                    .TextMatrix(i, mSpecColumn.���_��ǰ�ۼ�) = FormatEx(IIf(IsNull(rsRecord!��ǰ�ۼ�), 0, rsRecord!��ǰ�ۼ�), mintPriceDigit)
                End If
                .TextMatrix(i, mSpecColumn.���_������Ŀ) = ShowValue(.ColComboList(mSpecColumn.���_������Ŀ), rsRecord!������Ŀ, "]")
                .TextMatrix(i, mSpecColumn.���_������Ŀ) = IIf(IsNull(rsRecord!������Ŀ), "", rsRecord!������Ŀ)
                .TextMatrix(i, mSpecColumn.���_����ѱ���) = Format(IIf(IsNull(rsRecord!����ѱ���), "", rsRecord!����ѱ���), "0.00")
                .TextMatrix(i, mSpecColumn.���_ҩ�ۼ���) = ShowValue(.ColComboList(mSpecColumn.���_ҩ�ۼ���), IIf(IsNull(rsRecord!ҩ�ۼ���), "", rsRecord!ҩ�ۼ���), "-")
                .TextMatrix(i, mSpecColumn.���_���ηѱ�) = IIf(IsNull(rsRecord!���ηѱ�), "", rsRecord!���ηѱ�)
                .TextMatrix(i, mSpecColumn.���_��ֵ˰��) = Format(IIf(IsNull(rsRecord!��ֵ˰��), "", rsRecord!��ֵ˰��), "0.00")
                .TextMatrix(i, mSpecColumn.���_ҽ������) = ShowValue(.ColComboList(mSpecColumn.���_ҽ������), IIf(IsNull(rsRecord!ҽ������), "", rsRecord!ҽ������), "-")
                .TextMatrix(i, mSpecColumn.���_ҩ�����) = IIf(IsNull(rsRecord!ҩ�����), "", rsRecord!ҩ�����)
                .TextMatrix(i, mSpecColumn.���_ҩ������) = IIf(IsNull(rsRecord!ҩ������), "", rsRecord!ҩ������)
                .TextMatrix(i, mSpecColumn.���_ԭҩ�����) = IIf(IsNull(rsRecord!ҩ�����), "", rsRecord!ҩ�����)
                .TextMatrix(i, mSpecColumn.���_ԭҩ������) = IIf(IsNull(rsRecord!ҩ������), "", rsRecord!ҩ������)
                
                .TextMatrix(i, mSpecColumn.���_������) = FormatEx(IIf(Nvl(rsRecord!������, 0) = 0, 0, rsRecord!������), 5)
                .TextMatrix(i, mSpecColumn.���_��ʶ˵��) = IIf(IsNull(rsRecord!��ʶ˵��), "", rsRecord!��ʶ˵��)
                .TextMatrix(i, mSpecColumn.���_��ҩ����) = ShowValue(.ColComboList(mSpecColumn.���_��ҩ����), IIf(IsNull(rsRecord!��ҩ����), "", rsRecord!��ҩ����))
                .TextMatrix(i, mSpecColumn.���_վ����) = ShowValue(.ColComboList(mSpecColumn.���_վ����), IIf(IsNull(rsRecord!վ����), "", rsRecord!վ����), "-", True)
                .TextMatrix(i, mSpecColumn.���_DDDֵ) = IIf(IsNull(rsRecord!dddֵ), "", rsRecord!dddֵ)
                .TextMatrix(i, mSpecColumn.���_�������) = ShowValue(.ColComboList(mSpecColumn.���_�������), IIf(IsNull(rsRecord!�������), "", rsRecord!�������), "-", True)
                .TextMatrix(i, mSpecColumn.���_��ΣҩƷ) = ShowValue(.ColComboList(mSpecColumn.���_��ΣҩƷ), IIf(IsNull(rsRecord!��ΣҩƷ), "", rsRecord!��ΣҩƷ), "-", True)

                If str���� Like "�в�ҩ*" Then
                    If IsNull(rsRecord!סԺ�ɷ����) Or rsRecord!סԺ�ɷ���� = 0 Then
                        .TextMatrix(i, mSpecColumn.���_סԺ����ʹ��) = "0-���Է���"
                    Else
                        .TextMatrix(i, mSpecColumn.���_סԺ����ʹ��) = "1-���ɷ���"
                    End If
                    If IsNull(rsRecord!����ɷ����) Or rsRecord!����ɷ���� = 0 Then
                        .TextMatrix(i, mSpecColumn.���_�������ʹ��) = "0-���Է���"
                    Else
                        .TextMatrix(i, mSpecColumn.���_�������ʹ��) = "1-���ɷ���"
                    End If
                    
                    If .TextMatrix(i, mSpecColumn.���_��ҩ��̬) = "ɢװ" Then
                        .TextMatrix(i, mSpecColumn.���_סԺ����ʹ��) = "0-���Է���"
                        .Cell(flexcpBackColor, i, mSpecColumn.���_סԺ����ʹ��) = mlngColor
                        .TextMatrix(i, mSpecColumn.���_�������ʹ��) = "0-���Է���"
                        .Cell(flexcpBackColor, i, mSpecColumn.���_�������ʹ��) = mlngColor
                    Else
                        .Cell(flexcpBackColor, i, mSpecColumn.���_סԺ����ʹ��) = mlngApplyColor
                        .Cell(flexcpBackColor, i, mSpecColumn.���_�������ʹ��) = mlngApplyColor
                    End If
                Else
                    If IsNull(rsRecord!סԺ�ɷ����) Or rsRecord!סԺ�ɷ���� = 0 Then
                        intTemp = 0
                    ElseIf rsRecord!סԺ�ɷ���� = 1 Then
                        intTemp = 1
                    ElseIf rsRecord!סԺ�ɷ���� = 2 Then
                        intTemp = 2
                    ElseIf rsRecord!סԺ�ɷ���� = -1 Then
                        intTemp = 3
                    ElseIf rsRecord!סԺ�ɷ���� = -2 Then
                        intTemp = 4
                    ElseIf rsRecord!סԺ�ɷ���� = -3 Then
                        intTemp = 5
                    End If
                    .TextMatrix(i, mSpecColumn.���_סԺ����ʹ��) = ShowValue(.ColComboList(mSpecColumn.���_סԺ����ʹ��), IIf(IsNull(rsRecord!����ɷ����), "", intTemp), "-", True)
                    
                    If IsNull(rsRecord!����ɷ����) Or rsRecord!����ɷ���� = 0 Then
                        intTemp = 0
                    ElseIf rsRecord!����ɷ���� = 1 Then
                        intTemp = 1
                    ElseIf rsRecord!����ɷ���� = 2 Then
                        intTemp = 2
                    ElseIf rsRecord!����ɷ���� = -1 Then
                        intTemp = 3
                    ElseIf rsRecord!����ɷ���� = -2 Then
                        intTemp = 4
                    ElseIf rsRecord!����ɷ���� = -3 Then
                        intTemp = 5
                    End If
                    .TextMatrix(i, mSpecColumn.���_�������ʹ��) = ShowValue(.ColComboList(mSpecColumn.���_�������ʹ��), IIf(IsNull(rsRecord!����ɷ����), "", intTemp), "-", True)
                End If
                .TextMatrix(i, mSpecColumn.���_����ҩ��) = ShowValue(.ColComboList(mSpecColumn.���_����ҩ��), IIf(IsNull(rsRecord!����ҩ��), "", rsRecord!����ҩ��))
                .TextMatrix(i, mSpecColumn.���_סԺ��̬����) = IIf(IsNull(rsRecord!סԺ��̬����), "", rsRecord!סԺ��̬����)
                .TextMatrix(i, mSpecColumn.���_�洢�¶�) = ShowValue(.ColComboList(mSpecColumn.���_�洢�¶�), IIf(IsNull(rsRecord!�洢�¶�), "", rsRecord!�洢�¶�), "-", True)
                .TextMatrix(i, mSpecColumn.���_�洢����) = IIf(IsNull(rsRecord!�洢����), "", rsRecord!�洢����)
                .TextMatrix(i, mSpecColumn.���_��Һע������) = IIf(IsNull(rsRecord!��Һע������), "", rsRecord!��Һע������)
                .TextMatrix(i, mSpecColumn.���_��ҩ����) = ShowValue(.ColComboList(mSpecColumn.���_��ҩ����), IIf(IsNull(rsRecord!��ҩ����), "", rsRecord!��ҩ����))
                .TextMatrix(i, mSpecColumn.���_�������) = IIf(IsNull(rsRecord!�������), "", rsRecord!�������)
                .TextMatrix(i, mSpecColumn.���_�б�ҩƷ) = IIf(IsNull(rsRecord!�б�ҩƷ), 0, rsRecord!�б�ҩƷ)
                .TextMatrix(i, mSpecColumn.���_��ͬ��λid) = IIf(IsNull(rsRecord!��ͬ��λid), "", rsRecord!��ͬ��λid)
                .TextMatrix(i, mSpecColumn.���_������Ŀid) = IIf(IsNull(rsRecord!������Ŀid), "", rsRecord!������Ŀid)
                
                Call CheckValue(i, rsRecord!ID)
            End With
            rsRecord.MoveNext
        Next
        vsfDetails.MergeCol(mSpecColumn.���_ͨ������) = True   '�ϲ�ͨ������
        With vsfDetails
'            .Cell(flexcpBackColor, 1, mSpecColumn.���_ҩƷ����, .Rows - 1) = mlngColor
            .Cell(flexcpBackColor, 1, mSpecColumn.���_������, .Rows - 1) = mlngColor
            .Cell(flexcpBackColor, 1, mSpecColumn.���_ͨ������, .Rows - 1) = mlngColor
            .Cell(flexcpBackColor, 1, mSpecColumn.���_������λ, .Rows - 1) = mlngColor
            
'            If str���� Like "�в�ҩ*" Then
'                .ColHidden(mSpecColumn.���_��Ʒ����) = True
'                .ColHidden(mSpecColumn.���_ƴ����) = True
'                .ColHidden(mSpecColumn.���_�����) = True
'                .ColHidden(mSpecColumn.���_GMP��֤) = True
'                .ColHidden(mSpecColumn.���_סԺ��λ) = True
'                .ColHidden(mSpecColumn.���_סԺϵ��) = True
'                .ColHidden(mSpecColumn.���_��ҩ��̬) = False
'                .ColHidden(mSpecColumn.���_�������) = True
'                .ColHidden(mSpecColumn.���_������) = True
'                .ColHidden(mSpecColumn.���_����ҩ��) = True
'                .ColHidden(mSpecColumn.���_סԺ��̬����) = True
'            Else
'                .ColHidden(mSpecColumn.���_��Ʒ����) = False
'                .ColHidden(mSpecColumn.���_ƴ����) = False
'                .ColHidden(mSpecColumn.���_�����) = False
'                .ColHidden(mSpecColumn.���_GMP��֤) = False
'                .ColHidden(mSpecColumn.���_סԺ��λ) = False
'                .ColHidden(mSpecColumn.���_סԺϵ��) = False
'                .ColHidden(mSpecColumn.���_��ҩ��̬) = True
'                .ColHidden(mSpecColumn.���_�������) = False
'                .ColHidden(mSpecColumn.���_������) = False
'                .ColHidden(mSpecColumn.���_����ҩ��) = False
'                .ColHidden(mSpecColumn.���_סԺ��̬����) = False
'            End If
        End With
    End If
    
    Call Recover    '���޸��˵���ɫ�ı����
    
    '�����и�
    With vsfDetails
        For i = 1 To .Rows - 1
            .RowHeight(i) = 350
        Next
    End With
End Sub

Private Function ShowValue(ByVal strValue As String, ByVal strBiJiao As String, Optional str�ֽ�� As String, Optional bln��ƥ�� As Boolean) As String
    '���� ��ͨ�������ֵ�ȽϷ�������ȡ��ֵ
    '���� strvalue ԭ�ַ���
    'strBiJiao ��Ҫ�Ƚϵ��ַ���
    'str�ֽ�� �����ַ����еķָ����ţ����ڽ��жԱȷֽ���ұ߻���ߵ��ַ��� "0-�ɷ���"
    'bln��ƥ��=true �Աȷֽ����ߵ��ַ���/=false �Աȷֽ���ұߵ��ַ���
    Dim arr As Variant
    Dim i As Integer

    If strValue = "" Then Exit Function
    ReDim arr(UBound(Split(strValue, "|"))) As String   '���¶������鳤��

    '��ֵ�ֽ⿪�����浽������
    For i = 0 To UBound(Split(strValue, "|"))
        arr(i) = Split(strValue, "|")(i)
    Next
    If strBiJiao = "" Then
        ShowValue = ""
        Exit Function
    End If

    'ѭ���Ƚ�
    For i = 0 To UBound(Split(strValue, "|"))
        If Trim(arr(i)) <> "" Then
            If str�ֽ�� = "" Then '�ֽ��Ϊ��ֱ�ӽ��жԱ�
                If arr(i) = strBiJiao Then
                    ShowValue = arr(i)
                    Exit Function
                End If
            Else
                If bln��ƥ�� Then '�Աȷֽ����ߵ��ַ���
                    If Mid(arr(i), 1, InStr(1, arr(i), str�ֽ��) - 1) = strBiJiao Then
                        ShowValue = arr(i)
                        Exit Function
                    End If
                Else                    '�Աȷֽ���ұߵ��ַ���
                    If Mid(arr(i), InStr(1, arr(i), str�ֽ��) + 1) = strBiJiao Then
                        ShowValue = arr(i)
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
End Function

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call FindGridRow(UCase(txtFind))
        txtFind.SetFocus
    End If
End Sub

Private Sub vsfDetails_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If mint���� = 2 Then
        With vsfDetails
            If mint״̬ <> 1 Then   '���
                Select Case Col
                    Case mSpecColumn.���_ҩ�����
                        If .TextMatrix(Row, Col) <> "0" Then
                            .Cell(flexcpBackColor, Row, mSpecColumn.���_ҩ������) = cstcolor_backcolor
                            .Cell(flexcpBackColor, Row, mSpecColumn.���_������) = cstcolor_backcolor
                            If .TextMatrix(Row, mSpecColumn.���_������) = 0 Then
                                .TextMatrix(Row, mSpecColumn.���_������) = 24
                                .Cell(flexcpForeColor, Row, mSpecColumn.���_������) = mlngApplyColor
                                .Cell(flexcpFontBold, Row, mSpecColumn.���_������) = 10
                            End If
                        Else
                            .Cell(flexcpBackColor, Row, mSpecColumn.���_ҩ������) = mlngColor
                            .Cell(flexcpBackColor, Row, mSpecColumn.���_������) = mlngColor
                            .TextMatrix(Row, mSpecColumn.���_ҩ������) = 0
                            .TextMatrix(Row, mSpecColumn.���_������) = 0
                            .Cell(flexcpForeColor, Row, mSpecColumn.���_������) = mlngApplyColor
                            .Cell(flexcpFontBold, Row, mSpecColumn.���_������) = 10
                        End If
                        If .TextMatrix(Row, Col) <> mstrOldValue Then
                            .Cell(flexcpBackColor, Row, Col) = mlngApplyColor
                        End If
                    Case mSpecColumn.���_���ηѱ�, mSpecColumn.���_סԺ��̬����, mSpecColumn.���_GMP��֤, mSpecColumn.���_�ǳ���ҩ, mSpecColumn.���_ҩ������, mSpecColumn.���_�洢����, mSpecColumn.���_�������
                        If .TextMatrix(Row, Col) <> mstrOldValue Then
                            .Cell(flexcpBackColor, Row, Col) = mlngApplyColor
                        End If
                End Select
            Else    'Ʒ��
                Select Case Col
                    Case mVaricolumn.Ʒ��_ԭ��ҩ, mVaricolumn.Ʒ��_ר��ҩ, mVaricolumn.Ʒ��_��������, mVaricolumn.Ʒ��_����ҩ, mVaricolumn.Ʒ��_��ҩ, mVaricolumn.Ʒ��_������ҩ, mVaricolumn.Ʒ��_ԭ��ҩ, mVaricolumn.Ʒ��_����ҩ, mVaricolumn.Ʒ��_��ý, mVaricolumn.Ʒ��_Ʒ���³���ҽ��, mVaricolumn.Ʒ��_Ƥ��, mVaricolumn.Ʒ��_��ζʹ��
                        If .TextMatrix(Row, Col) <> mstrOldValue Then
                            .Cell(flexcpBackColor, Row, Col) = mlngApplyColor
                        End If
                End Select
            End If
        End With
    End If
End Sub

Private Sub vsfDetails_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim i As Integer
    Dim j As Integer
    Dim rsRecord As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim strPro As String
    Dim strSql As String
    Dim intAttr As Integer
    Dim strSQLItem As String
    Dim intupdate As Integer
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim strTemp As String

    On Error GoTo ErrHandle
    With vsfDetails
        If .Cell(flexcpBackColor, NewRow, NewCol) = mlngColor Then
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
        End If
        
        If .Row < OldRow Then
            OldRow = 1
        End If
        
        If .Rows = 1 Then
           OldRow = 0
        End If
    End With
    
    '���Ʋ˵���Ӧ������������ʾ���
    With vsfDetails
        If mint״̬ = 1 Then 'Ʒ��
            Select Case NewCol
                Case mVaricolumn.Ʒ��_ͨ������
                    mcbrToolBar.Controls(1).Visible = False
                    mobjPopup.Controls(1).Visible = False
                Case mVaricolumn.Ʒ��_Ӣ������
                    mcbrToolBar.Controls(1).Visible = False
                    mobjPopup.Controls(1).Visible = False
                Case mVaricolumn.Ʒ��_ƴ����
                    mcbrToolBar.Controls(1).Visible = False
                    mobjPopup.Controls(1).Visible = False
                Case mVaricolumn.Ʒ��_�����
                    mcbrToolBar.Controls(1).Visible = False
                    mobjPopup.Controls(1).Visible = False
                Case Else
                    If .Cell(flexcpBackColor, NewRow, NewCol, NewRow, NewCol) = mlngColor Then
                        mcbrToolBar.Controls(1).Visible = False
                        mobjPopup.Controls(1).Visible = False
                    Else
                        mcbrToolBar.Controls(1).Visible = True
                        mobjPopup.Controls(1).Visible = True
                    End If
            End Select
            
            Select Case OldCol
                Case mVaricolumn.Ʒ��_ͨ������, mVaricolumn.Ʒ��_������λ
                    If vsfDetails.TextMatrix(OldRow, OldCol) = "" Then
                        MsgBox "�õ�Ԫ�����ݲ���Ϊ�գ������룡", vbInformation, gstrSysName
                        vsfDetails.Select OldRow, OldCol
                    End If
            End Select
        Else    '���
            Select Case NewCol
                Case mSpecColumn.���_��Ʒ����
                    mcbrToolBar.Controls(1).Visible = False
                    mobjPopup.Controls(1).Visible = False
                Case mSpecColumn.���_ƴ����
                    mcbrToolBar.Controls(1).Visible = False
                    mobjPopup.Controls(1).Visible = False
                Case mSpecColumn.���_�����
                    mcbrToolBar.Controls(1).Visible = False
                    mobjPopup.Controls(1).Visible = False
                Case Else
                    If .Cell(flexcpBackColor, NewRow, NewCol, NewRow, NewCol) = mlngColor Then
                        mcbrToolBar.Controls(1).Visible = False
                        mobjPopup.Controls(1).Visible = False
                    Else
                        mcbrToolBar.Controls(1).Visible = True
                        mobjPopup.Controls(1).Visible = True
                    End If
            End Select
            
            Select Case OldCol
                Case mSpecColumn.���_ָ������
                    If Val(.TextMatrix(OldRow, OldCol)) < 100 Then
                       .TextMatrix(OldRow, mSpecColumn.���_�ӳ���) = Format((1 / (1 - Val(.TextMatrix(OldRow, OldCol)) / 100) - 1) * 100, "0.00")
                       
                       If mstrOldValue <> .TextMatrix(OldRow, OldCol) Then
                            .Cell(flexcpForeColor, OldRow, mSpecColumn.���_�ӳ���) = mlngApplyColor: .Cell(flexcpFontSize, OldRow, mSpecColumn.���_�ӳ���) = 10: .Cell(flexcpFontBold, OldRow, mSpecColumn.���_�ӳ���) = True
                            .Cell(flexcpForeColor, OldRow, mSpecColumn.���_ָ������) = mlngApplyColor: .Cell(flexcpFontSize, OldRow, mSpecColumn.���_ָ������) = 10: .Cell(flexcpFontBold, OldRow, mSpecColumn.���_ָ������) = True
                       End If
                    Else
                        '���������ָ������ʴ��ڵ���100������������Ҫ�Ӽӳ��ʷ������
                        .TextMatrix(OldRow, OldCol) = Format((1 - (1 / (1 + Val(.TextMatrix(OldRow, mSpecColumn.���_�ӳ���)) / 100))) * 100, "0.00000")
                        If mstrOldValue <> .TextMatrix(OldRow, OldCol) Then
                            .Cell(flexcpForeColor, OldRow, mSpecColumn.���_ָ������) = mlngApplyColor: .Cell(flexcpFontSize, OldRow, mSpecColumn.���_ָ������) = 10: .Cell(flexcpFontBold, OldRow, mSpecColumn.���_ָ������) = True
                            .Cell(flexcpForeColor, OldRow, mSpecColumn.���_�ӳ���) = mlngApplyColor: .Cell(flexcpFontSize, OldRow, mSpecColumn.���_�ӳ���) = 10: .Cell(flexcpFontBold, OldRow, mSpecColumn.���_�ӳ���) = True
                        End If
                    End If
                Case mSpecColumn.���_�ӳ���
                    If .TextMatrix(OldRow, OldCol) <> "" Then
                       .TextMatrix(OldRow, mSpecColumn.���_ָ������) = Format((1 - (1 / (1 + Val(.TextMatrix(OldRow, mSpecColumn.���_�ӳ���)) / 100))) * 100, "0.00000")
                    End If
                    If mstrOldValue <> .TextMatrix(OldRow, OldCol) Then
                        .Cell(flexcpForeColor, OldRow, mSpecColumn.���_ָ������) = mlngApplyColor: .Cell(flexcpFontSize, OldRow, mSpecColumn.���_ָ������) = 10: .Cell(flexcpFontBold, OldRow, mSpecColumn.���_ָ������) = True
                        .Cell(flexcpForeColor, OldRow, mSpecColumn.���_�ӳ���) = mlngApplyColor: .Cell(flexcpFontSize, OldRow, mSpecColumn.���_�ӳ���) = 10: .Cell(flexcpFontBold, OldRow, mSpecColumn.���_�ӳ���) = True
                    End If
                Case mSpecColumn.���_����ϵ��, mSpecColumn.���_סԺϵ��, mSpecColumn.���_����ϵ��, mSpecColumn.���_ҩ��ϵ��, mSpecColumn.���_�ͻ���װ, mSpecColumn.���_���췧ֵ, mSpecColumn.���_�ɹ��޼�, mSpecColumn.���_�ɹ�����, mSpecColumn.���_�����, mSpecColumn.���_ָ���ۼ�, mSpecColumn.���_ָ������, mSpecColumn.���_�ӳ���, mSpecColumn.���_����ϵ��, mSpecColumn.���_�������, mSpecColumn.���_�ɱ��۸�, mSpecColumn.���_��ǰ�ۼ�, mSpecColumn.���_����ѱ���, mSpecColumn.���_��ֵ˰��
                    If .TextMatrix(OldRow, OldCol) <> "" Then
                        If Mid(.TextMatrix(OldRow, OldCol), 1, 1) = "." Then
                            .TextMatrix(OldRow, OldCol) = "0" & .TextMatrix(OldRow, OldCol)
                        End If
                        
                        If Mid(.TextMatrix(OldRow, OldCol), Len(.TextMatrix(OldRow, OldCol)), 1) = "." Then
                            .TextMatrix(OldRow, OldCol) = Mid(.TextMatrix(OldRow, OldCol), 1, Len(.TextMatrix(OldRow, OldCol)) - 1)
                        End If
                    End If
                Case mSpecColumn.���_������Ŀ
                    If OldRow <> 0 Then
                        If .TextMatrix(OldRow, OldCol) <> "" Then
                            strTemp = Mid(.TextMatrix(OldRow, OldCol), 2, InStr(1, .TextMatrix(OldRow, OldCol), "]") - 2)
                        End If
                        gstrSql = "Select ID" & _
                                  "  From ������Ŀ" & _
                                  "  Where ����=[1] and ĩ�� = 1 And (����ʱ�� Is Null Or ����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD'))"
            
                        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, "������Ŀ��ѯ", strTemp)
                        If rsTmp.RecordCount > 0 Then
                            .TextMatrix(OldRow, mSpecColumn.���_������Ŀid) = rsTmp!ID
                        End If
                    End If
            End Select
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfDetails_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '������Щ�п��Ա༭����Щ�в����Ա༭,��������ɫΪ��ɫ���ж��������޸�
    With vsfDetails
        mstrOldValue = vsfDetails.TextMatrix(Row, Col)
        If .Cell(flexcpBackColor, Row, Col) = mlngColor Then
            Cancel = True
        End If
    End With
End Sub

Private Sub vsfDetails_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsRecord As ADODB.Recordset
    Dim rsTmp As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim strPro As String
    Dim strSql As String
    Dim intAttr As Integer
    Dim strSQLItem As String
    Dim dblLeft As Double
    Dim dblTop As Double
    
    vRect = GetControlRect(vsfDetails.hWnd) '��ȡλ��
    dblLeft = vRect.Left + vsfDetails.CellLeft
    dblTop = vRect.Top + vsfDetails.CellTop + vsfDetails.CellHeight + 3200
    On Error GoTo ErrHandle
    With vsfDetails
        If mint״̬ = 1 Then    'Ʒ��
            If Col = mVaricolumn.Ʒ��_�ο���Ŀ Then
                strSql = "Select ���� From ���Ʒ���Ŀ¼ Where ID=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, .TextMatrix(.Row, mVaricolumn.Ʒ��_����id))
                
                If rsTmp.EOF Then
                    intAttr = -1
                Else
                    intAttr = rsTmp!����
                End If
                
                strSql = " Select ID,����ID,����,����,˵�� From ���Ʋο�Ŀ¼ a Where ����=[1] Order By ����"
                
                Set rsRecord = zlDatabase.ShowSQLSelect(Me, strSql, 0, "������Ŀ", False, "", "", False, False, _
                True, dblLeft, dblTop, vsfDetails.Height, blnCancel, False, True, intAttr)

                If rsRecord Is Nothing Then
                    Exit Sub
                End If
                .TextMatrix(.Row, mVaricolumn.Ʒ��_�ο���Ŀ) = rsRecord!����
                .TextMatrix(.Row, mVaricolumn.Ʒ��_�ο���ĿID) = rsRecord!ID
            End If
        Else    '���
            Select Case Col
                Case mSpecColumn.���_��������
                    strSql = "Select ���� as id,����,���� From ҩƷ������ Order By ���� "
                    
                    Set rsRecord = zlDatabase.ShowSQLSelect(Me, strSql, 0, "������Ŀ", False, "", "", False, False, _
                        True, dblLeft, dblTop, vsfDetails.Height, blnCancel, False, True)
                    
                    If rsRecord Is Nothing Then
                        Exit Sub
                    Else
                        .TextMatrix(.Row, mSpecColumn.���_��������) = rsRecord!����
                    End If
                Case mSpecColumn.���_��ͬ��λ
                    strSql = "Select id,����,����,����" & _
                                " From ��Ӧ��" & _
                                " where ĩ��=1 And substr(����,1,1) = '1' And (����ʱ�� is null or ����ʱ��=to_date('3000-01-01','YYYY-MM-DD')) " & _
                                " Order By ���� "
                    Set rsRecord = zlDatabase.ShowSQLSelect(Me, strSql, 0, "������Ŀ", False, "", "", False, False, _
                        True, dblLeft, dblTop, vsfDetails.Height, blnCancel, False, True)
                    
                    If rsRecord Is Nothing Then
                        Exit Sub
                    Else
                        .TextMatrix(.Row, mSpecColumn.���_��ͬ��λ) = rsRecord!����
                        .TextMatrix(.Row, mSpecColumn.���_��ͬ��λid) = rsRecord!ID
                    End If
                Case mSpecColumn.���_������Ŀ
                    Dim blnRe As Boolean
                    Dim str���� As String
                    Dim strID As String
                    
                    gstrSql = "Select ���� as id,�ϼ� as �ϼ�id, ����, ����, ĩ�� From ������Ŀ Start With �ϼ� Is Null Connect By Prior ���� = �ϼ�"
                    blnRe = frmTreeLeafSel.ShowTree(gstrSql, strID, str����, "������Ŀ")
                    '�ɹ�����
                    If blnRe Then
                        .TextMatrix(.Row, mSpecColumn.���_������Ŀ) = str����
                    End If
            End Select
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfDetails_CellChanged(ByVal Row As Long, ByVal Col As Long)
    With vsfDetails
        If mstrOldValue <> .TextMatrix(Row, Col) And .CellBackColor <> mlngColor And mblnClick = True And Row = mintRow And .Rows <> 1 Then
            .Cell(flexcpForeColor, Row, Col) = mlngApplyColor: .Cell(flexcpFontSize, Row, Col) = 10: .Cell(flexcpFontBold, Row, Col) = True
        End If
    End With
End Sub

Private Sub vsfDetails_ChangeEdit()
    Dim lngId As Long
    Dim strTemp As String
    
    With vsfDetails
        If mint״̬ = 1 Then 'Ʒ��
            Select Case .Col
                Case mVaricolumn.Ʒ��_ͨ������
                    .TextMatrix(.Row, mVaricolumn.Ʒ��_ƴ����) = zlGetSymbol(.EditText, 0, 30)
                    .TextMatrix(.Row, mVaricolumn.Ʒ��_�����) = zlGetSymbol(.EditText, 1, 30)
                    .Cell(flexcpForeColor, .Row, mVaricolumn.Ʒ��_ƴ����) = mlngApplyColor: .Cell(flexcpFontSize, .Row, mVaricolumn.Ʒ��_ƴ����) = 10: .Cell(flexcpFontBold, .Row, mVaricolumn.Ʒ��_ƴ����) = True
                    .Cell(flexcpForeColor, .Row, mVaricolumn.Ʒ��_�����) = mlngApplyColor: .Cell(flexcpFontSize, .Row, mVaricolumn.Ʒ��_�����) = 10: .Cell(flexcpFontBold, .Row, mVaricolumn.Ʒ��_�����) = True
                 Case mVaricolumn.Ʒ��_������
                    If Mid(.EditText, 1, InStr(1, .EditText, "-") - 1) = 0 Then
                        .Cell(flexcpBackColor, .Row, mVaricolumn.Ʒ��_ATCCODE, .Row) = mlngColor
                        .TextMatrix(.Row, mVaricolumn.Ʒ��_ATCCODE) = ""
                    Else
                        .Cell(flexcpBackColor, .Row, mVaricolumn.Ʒ��_ATCCODE, .Row) = cstcolor_backcolor
                    End If
            End Select
        Else    '���
            Select Case .Col
                Case mSpecColumn.���_��Ʒ����
                    .TextMatrix(.Row, mSpecColumn.���_ƴ����) = zlGetSymbol(.EditText, 0, 30)
                    .TextMatrix(.Row, mSpecColumn.���_�����) = zlGetSymbol(.EditText, 1, 30)
                    .Cell(flexcpForeColor, .Row, mSpecColumn.���_ƴ����) = mlngApplyColor: .Cell(flexcpFontSize, .Row, mSpecColumn.���_ƴ����) = 10: .Cell(flexcpFontBold, .Row, mSpecColumn.���_ƴ����) = True
                    .Cell(flexcpForeColor, .Row, mSpecColumn.���_�����) = mlngApplyColor: .Cell(flexcpFontSize, .Row, mSpecColumn.���_�����) = 10: .Cell(flexcpFontBold, .Row, mSpecColumn.���_�����) = True
                Case mSpecColumn.���_ҩƷ���
                    lngId = .TextMatrix(.Row, mSpecColumn.���_id)
                    .TextMatrix(.Row, mSpecColumn.���_������) = zlGetDigitSign(lngId, .EditText)
                    If mstrOldValue <> .EditText Then
                        .Cell(flexcpForeColor, .Row, mSpecColumn.���_������) = mlngApplyColor: .Cell(flexcpFontSize, .Row, mSpecColumn.���_������) = 10: .Cell(flexcpFontBold, .Row, mSpecColumn.���_������) = True
                    End If
                Case mSpecColumn.���_�ɱ��۸�
                    .TextMatrix(.Row, mSpecColumn.���_�����) = FormatEx(Val(.EditText) * (Val(.TextMatrix(.Row, mSpecColumn.���_�ɹ�����)) / 100), mintPriceDigit)
                    .Cell(flexcpForeColor, .Row, mSpecColumn.���_�����) = mlngApplyColor: .Cell(flexcpFontSize, .Row, mSpecColumn.���_�����) = 10: .Cell(flexcpFontBold, .Row, mSpecColumn.���_�����) = True
                    .TextMatrix(.Row, mSpecColumn.���_�ɹ��޼�) = FormatEx(Val(.EditText), mintPriceDigit)
                    .Cell(flexcpForeColor, .Row, mSpecColumn.���_�ɹ��޼�) = mlngApplyColor: .Cell(flexcpFontSize, .Row, mSpecColumn.���_�ɹ��޼�) = 10: .Cell(flexcpFontBold, .Row, mSpecColumn.���_�ɹ��޼�) = True
                Case mSpecColumn.���_�ɹ��޼�
                    .TextMatrix(.Row, mSpecColumn.���_�����) = FormatEx(Val(.EditText) * (Val(.TextMatrix(.Row, mSpecColumn.���_�ɹ�����)) / 100), mintPriceDigit)
                    .Cell(flexcpForeColor, .Row, mSpecColumn.���_�����) = mlngApplyColor: .Cell(flexcpFontSize, .Row, mSpecColumn.���_�����) = 10: .Cell(flexcpFontBold, .Row, mSpecColumn.���_�����) = True
                Case mSpecColumn.���_�ɹ�����
                    .TextMatrix(.Row, mSpecColumn.���_�����) = FormatEx(Val(.EditText) / 100 * Val(.TextMatrix(.Row, mSpecColumn.���_�ɹ��޼�)), mintPriceDigit)
                    .Cell(flexcpForeColor, .Row, mSpecColumn.���_�����) = mlngApplyColor: .Cell(flexcpFontSize, .Row, mSpecColumn.���_�����) = 10: .Cell(flexcpFontBold, .Row, mSpecColumn.���_�����) = True
                Case mSpecColumn.���_סԺ����ʹ��
                    If Val(Mid(.EditText, 1, 1)) = 0 Then
                        .Cell(flexcpBackColor, .Row, mSpecColumn.���_סԺ��̬����) = mlngColor
                    Else
                        .Cell(flexcpBackColor, .Row, mSpecColumn.���_סԺ��̬����) = cstcolor_backcolor
                    End If
                Case mSpecColumn.���_��ҩ��̬
                    If .EditText <> "ɢװ" Then
                        .Cell(flexcpForeColor, .Row, mSpecColumn.���_��ҩ��̬) = mlngApplyColor
                        .Cell(flexcpFontBold, .Row, mSpecColumn.���_��ҩ��̬) = True
                        .Cell(flexcpFontSize, .Row, mSpecColumn.���_��ҩ��̬) = 10
                        MsgBox "���޸��ˡ���ҩ��̬����ϵͳ��ǿ���趨���ٴ�Ӧ�á�ҳ�з���ʹ��Ϊ�����ɷ��㡱��", vbInformation, gstrSysName
                        .Cell(flexcpBackColor, .Row, mSpecColumn.���_סԺ����ʹ��) = cstcolor_backcolor
                        .TextMatrix(.Row, mSpecColumn.���_סԺ����ʹ��) = "1-���ɷ���"
                        .Cell(flexcpBackColor, .Row, mSpecColumn.���_�������ʹ��) = cstcolor_backcolor
                        .TextMatrix(.Row, mSpecColumn.���_�������ʹ��) = "1-���ɷ���"
                    Else
                        .Cell(flexcpBackColor, .Row, mSpecColumn.���_סԺ����ʹ��) = mlngColor
                        .Cell(flexcpBackColor, .Row, mSpecColumn.���_�������ʹ��) = mlngColor
                    End If
            End Select
        End If
    End With
End Sub

Private Sub vsfDetails_Click()
    mblnClick = True
End Sub

Private Sub vsfDetails_EnterCell()
    Dim rsRecord As ADODB.Recordset
    Dim strKey As String
    Dim i As Integer
    Dim j As Integer
    
    If mintRow�ϴ� > vsfDetails.Rows - 1 Then
        mintRow�ϴ� = 1
    End If
    
    If vsfDetails.Rows <> 1 Then
        vsfDetails.Cell(flexcpPicture, mintRow�ϴ�, 0, mintRow�ϴ�, 0) = Nothing    '����ͼƬ
        For i = 1 To vsfDetails.Rows - 1    '�����л�ѡ��ҳ+����ʱ����ֶ��ͼƬ ������������Ƚ������һ�������
            If Not vsfDetails.Cell(flexcpPicture, i, 0, i, 0) Is Nothing Then
                vsfDetails.Cell(flexcpPicture, i, 0, i, 0) = Nothing
                Exit For
            End If
        Next
        vsfDetails.Cell(flexcpPicture, vsfDetails.Row, 0, vsfDetails.Row, 0) = Me.ImgTvw.ListImages(2).Picture
    End If
    
    With vsfDetails
        If .Row = mintRow Then Exit Sub
        mintRow = .Row '��¼��ǰ��
        strKey = .TextMatrix(.Row, mVaricolumn.Ʒ��_id)
    End With
End Sub

Private Sub vsfDetails_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call MoveRowCol
    End If
End Sub

Private Sub vsfDetails_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strSql As String, strSQLItem As String
    Dim rsRecord As ADODB.Recordset
    Dim iAttr As Integer
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim intAllCol As Integer
    
    On Error GoTo ErrHandle
    If KeyAscii = vbKeyReturn Then
        If vsfDetails.EditText = "" Then
            Call MoveRowCol
            Exit Sub
        End If
        
        If mint״̬ = 1 Then 'Ʒ��
            vRect = GetControlRect(vsfDetails.hWnd) '��ȡλ��
            dblLeft = vRect.Left + vsfDetails.CellLeft
            dblTop = vRect.Top + vsfDetails.CellTop + vsfDetails.CellHeight + 3200
            With vsfDetails
                If .Col = mVaricolumn.Ʒ��_�ο���Ŀ Then
                    strSql = "Select ���� From ���Ʒ���Ŀ¼ Where ID=[1]"
                    Set rsRecord = zlDatabase.OpenSQLRecord(strSql, Me.Caption, .TextMatrix(.Row, mVaricolumn.Ʒ��_����id))
                    
                    If rsRecord.EOF Then
                        iAttr = -1
                    Else
                        iAttr = rsRecord(0)
                    End If
                    If .EditText = "" Then
                        strSql = " Select ID,����ID,����,����,˵�� From ���Ʋο�Ŀ¼ a Where ����=" & iAttr & " Order By ����"
                    Else
                        strSQLItem = " From ���Ʋο�Ŀ¼ A,���Ʋο����� B" & _
                            " Where A.ID=B.�ο�Ŀ¼ID And A.����=[1]" & _
                            " And (Upper(A.����) Like [2] " & _
                            " Or Upper(A.����) Like [3] " & _
                            " Or Upper(B.����) Like [3] " & _
                            " Or Upper(B.����) Like [3] " & ")"
                
                        strSql = " Select DISTINCT A.ID,A.����ID,A.����,A.����,A.˵�� " & strSQLItem & " Order By ����"
                        Set rsRecord = zlDatabase.ShowSQLSelect(Me, strSql, 0, "������Ŀ", False, "", "", False, False, _
                        True, dblLeft, dblTop, vsfDetails.Height, blnCancel, False, True, iAttr, UCase(.EditText) & "%", mstrMatch & UCase(.EditText) & "%")
        
                        If rsRecord Is Nothing Then
                            .EditText = ""
                            .TextMatrix(.Row, mVaricolumn.Ʒ��_�ο���Ŀ) = ""
                            .TextMatrix(.Row, mVaricolumn.Ʒ��_�ο���ĿID) = ""
                            Exit Sub
                        End If
                        .EditText = rsRecord!����
                        .TextMatrix(.Row, mVaricolumn.Ʒ��_�ο���Ŀ) = rsRecord!����
                        .TextMatrix(.Row, mVaricolumn.Ʒ��_�ο���ĿID) = rsRecord!ID
                        End If
                End If
            End With
        Else    '���
            Dim str As String
            vRect = GetControlRect(vsfDetails.hWnd) '��ȡλ��
            dblLeft = vRect.Left + vsfDetails.CellLeft
            dblTop = vRect.Top + vsfDetails.CellTop + vsfDetails.CellHeight + 3200
            With vsfDetails
                If .EditText = "" Then Exit Sub
                Select Case Col
                    Case mSpecColumn.���_��������
                        str = UCase(.EditText)
                        If .Col = mSpecColumn.���_�������� Then
                            strSql = "Select ���� as id,����,����" & _
                                        " From ҩƷ������" & _
                                        " where ���� Like [1] " & _
                                        "       Or ���� Like [2] " & _
                                        "       Or ���� Like [2] Order By ���� "
                            Set rsRecord = zlDatabase.ShowSQLSelect(Me, strSql, 0, "������Ŀ", False, "", "", False, False, _
                                True, dblLeft, dblTop, vsfDetails.Height, blnCancel, False, True, str & "%", mstrMatch & str & "%")
                            If rsRecord Is Nothing Then
                                .EditText = ""
                                Exit Sub
                            Else
                                .EditText = rsRecord!����
                                .TextMatrix(.Row, mSpecColumn.���_��������) = rsRecord!����
                            End If
                        End If
                    Case mSpecColumn.���_��ͬ��λ
                        strSql = "Select ����,����,����,id" & _
                                    " From ��Ӧ��" & _
                                    " where (���� Like [1] " & _
                                    "       Or ���� Like [2] " & _
                                    "       Or ���� Like [2])" & _
                                    " And ĩ��=1 And substr(����,1,1) = '1' And (����ʱ�� is null or ����ʱ��=to_date('3000-01-01','YYYY-MM-DD')) " & _
                                    " Order By ���� "
                        Set rsRecord = zlDatabase.ShowSQLSelect(Me, strSql, 0, "������Ŀ", False, "", "", False, False, _
                            True, dblLeft, dblTop, vsfDetails.Height, blnCancel, False, True, UCase(.EditText) & "%", mstrMatch & UCase(.EditText) & "%")
                        
                        If rsRecord Is Nothing Then
                            MsgBox "û���ҵ�ƥ��Ĺ�Ӧ�̣����ڹ�Ӧ�̹��������ӹ�Ӧ�̣�", vbInformation, gstrSysName
                            .TextMatrix(.Row, mSpecColumn.���_��ͬ��λ) = ""
                            .TextMatrix(.Row, mSpecColumn.���_��ͬ��λid) = ""
                            Exit Sub
                        Else
                            .EditText = rsRecord!����
                            .TextMatrix(.Row, mSpecColumn.���_��ͬ��λ) = rsRecord!����
                            .TextMatrix(.Row, mSpecColumn.���_��ͬ��λid) = rsRecord!ID
                        End If
                End Select
            End With
        End If
        
        Call MoveRowCol
    End If
    
    If KeyAscii <> vbKeyBack Then
        With vsfDetails
            If mint״̬ = 1 Then    'Ʒ��
                Select Case Col
                    Case mVaricolumn.Ʒ��_ͨ������
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mVaricolumn.Ʒ��_ͨ������)) Then
                            KeyAscii = 0
                        End If
                    Case mVaricolumn.Ʒ��_Ӣ������
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mVaricolumn.Ʒ��_Ӣ������)) Then
                            KeyAscii = 0
                        End If
                    Case mVaricolumn.Ʒ��_ƴ����
                        If Not (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Or KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mVaricolumn.Ʒ��_ƴ����)) Then
                            KeyAscii = 0
                        End If
                    Case mVaricolumn.Ʒ��_�����
                        If Not (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Or KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mVaricolumn.Ʒ��_�����)) Then
                            KeyAscii = 0
                        End If
                    Case mVaricolumn.Ʒ��_��������
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mVaricolumn.Ʒ��_��������)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mVaricolumn.Ʒ��_������λ
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mVaricolumn.Ʒ��_��������)) Then
                            KeyAscii = 0
                        End If
                    Case mVaricolumn.Ʒ��_ATCCODE
                        If KeyAscii <> vbKeyDelete Then
                            If Not (KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mVaricolumn.Ʒ��_ATCCODE)) Then
                                KeyAscii = 0
                            Else
                                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                            End If
                        End If
                End Select
            Else    '���
                Select Case Col
                    Case mSpecColumn.���_ҩƷ���
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.���_ҩƷ���)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.���_��λ��
                        If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= 20 Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.���_������
                        If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= 7 Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.���_��ʶ��
                        If Not (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Or KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.���_��ʶ��)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.���_��ѡ��
                        If Not (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Or KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.���_��ѡ��)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.���_����
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.���_��Ʒ����
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.���_��Ʒ����)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.���_��������
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.���_��������)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.���_ƴ����
                        If Not (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Or KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.���_ƴ����)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.���_�����
                        If Not (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Or KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Or KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.���_�����)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.���_��ͬ��λ
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.���_��ͬ��λ)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.���_��׼�ĺ�
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.���_��׼�ĺ�)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.���_ע���̱�
                    
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.���_ע���̱�)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.���_�ۼ۵�λ
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.���_�ۼ۵�λ)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.���_����ϵ��
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.���_����ϵ��)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.���_סԺ��λ
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= mintLen Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.���_סԺϵ��
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.���_סԺϵ��)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.���_���ﵥλ
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.���_���ﵥλ)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.���_����ϵ��
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.���_����ϵ��)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.���_ҩ�ⵥλ
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.���_ҩ�ⵥλ)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.���_ҩ��ϵ��
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.���_ҩ��ϵ��)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.���_�ͻ���λ
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.���_�ͻ���λ)) Or InStr(" ^&`'""", Chr(KeyAscii)) > 0 Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.���_�ͻ���װ
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.���_�ͻ���װ)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.���_���췧ֵ
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.���_���췧ֵ)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.���_�ɹ��޼�
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.���_�ɹ��޼�)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.���_�ɹ�����
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.���_�ɹ�����)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.���_ָ���ۼ�
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.���_ָ���ۼ�)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.���_ָ������
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.���_ָ������)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.���_�ӳ���
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= 19 Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.���_�������
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.���_�������)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.���_�ɱ��۸�
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.���_�ɱ��۸�)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.���_��ǰ�ۼ�
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.���_��ǰ�ۼ�)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.���_����ѱ���
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.���_����ѱ���)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.���_��ֵ˰��
                        If KeyAscii = vbKeyDelete Then
                            If InStr(1, .EditText, ".") > 0 Then
                                KeyAscii = 0
                            End If
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.���_��ֵ˰��)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.���_������
                        If KeyAscii = vbKeyDelete Then
                            KeyAscii = 0
                        Else
                            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.���_������)) Then
                                KeyAscii = 0
                            End If
                        End If
                    Case mSpecColumn.���_��ʶ˵��
                        If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.���_��ʶ˵��)) Then
                            KeyAscii = 0
                        End If
                    Case mSpecColumn.���_��Һע������
                      If LenB(StrConv(.EditText, vbFromUnicode)) >= Val(.ColKey(mSpecColumn.���_��Һע������)) Or InStr(" ^&`'""", Chr(KeyAscii)) > 0 Then
                          KeyAscii = 0
                      End If
                End Select
            End If
        End With
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfDetails_LeaveCell()
    mintRow�ϴ� = vsfDetails.Row
    mintCol�ϴ� = vsfDetails.Col
End Sub


Private Sub vsfDetails_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        mobjPopup.ShowPopup
    End If
End Sub

Private Sub SetȨ���ж�()
'Ȩ���жϹ���
    With vsfDetails
        If .Rows > 1 Then
            If mint״̬ = 1 Then    'Ʒ��
                If InStr(1, mstrPrivs, "ҽ����ҩĿ¼") = 0 Then
                    .Cell(flexcpBackColor, 1, mVaricolumn.Ʒ��_ҽ��ְ��, .Rows - 1, mVaricolumn.Ʒ��_ҽ��ְ��) = mlngColor
                End If
            Else    '���
                If InStr(1, mstrPrivs, "ҽ����ҩĿ¼") = 0 Then
                    .Cell(flexcpBackColor, 1, mSpecColumn.���_ҽ������, .Rows - 1, mSpecColumn.���_ҽ������) = mlngColor
                End If
                If InStr(1, mstrPrivs, "�������") = 0 Then
                    .Cell(flexcpBackColor, 1, mSpecColumn.���_�ɹ�����, .Rows - 1, mSpecColumn.���_�ɹ�����) = mlngColor
                End If
                If InStr(1, mstrPrivs, "ָ���۸����") = 0 Then
                    .Cell(flexcpBackColor, 1, mSpecColumn.���_ָ������, .Rows - 1, mSpecColumn.���_ָ������) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.���_�ӳ���, .Rows - 1, mSpecColumn.���_�ӳ���) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.���_�ɹ��޼�, .Rows - 1, mSpecColumn.���_�ɹ��޼�) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.���_ָ���ۼ�, .Rows - 1, mSpecColumn.���_ָ���ۼ�) = mlngColor
                End If
                If InStr(1, mstrPrivs, "�ۼ۹���") = 0 Then
                    .Cell(flexcpBackColor, 1, mSpecColumn.���_ҩ������, .Rows - 1, mSpecColumn.���_ҩ������) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.���_������Ŀ, .Rows - 1, mSpecColumn.���_������Ŀ) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.���_�������, .Rows - 1, mSpecColumn.���_�������) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.���_��ǰ�ۼ�, .Rows - 1, mSpecColumn.���_��ǰ�ۼ�) = mlngColor
                End If
                If InStr(1, mstrPrivs, "ҩ�ۼ���") = 0 Then
                    .Cell(flexcpBackColor, 1, mSpecColumn.���_ҩ�ۼ���, .Rows - 1, mSpecColumn.���_ҩ�ۼ���) = mlngColor
                End If
                If InStr(1, mstrPrivs, "�ɱ��۹���") = 0 Then
                    .Cell(flexcpBackColor, 1, mSpecColumn.���_�ɱ��۸�, .Rows - 1, mSpecColumn.���_�ɱ��۸�) = mlngColor
                End If
                If InStr(1, mstrPrivs, "�����������") = 0 Then
                    .Cell(flexcpBackColor, 1, mSpecColumn.���_�������, .Rows - 1, mSpecColumn.���_�������) = mlngColor
                End If
                If InStr(1, mstrPrivs, "ҩƷ��λ����") = 0 Then
                    .Cell(flexcpBackColor, 1, mSpecColumn.���_�ۼ۵�λ, .Rows - 1, mSpecColumn.���_�ۼ۵�λ) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.���_סԺ��λ, .Rows - 1, mSpecColumn.���_סԺ��λ) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.���_���ﵥλ, .Rows - 1, mSpecColumn.���_���ﵥλ) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.���_ҩ�ⵥλ, .Rows - 1, mSpecColumn.���_ҩ�ⵥλ) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.���_����ϵ��, .Rows - 1, mSpecColumn.���_����ϵ��) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.���_סԺϵ��, .Rows - 1, mSpecColumn.���_סԺϵ��) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.���_����ϵ��, .Rows - 1, mSpecColumn.���_����ϵ��) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.���_ҩ��ϵ��, .Rows - 1, mSpecColumn.���_ҩ��ϵ��) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.���_�ͻ���λ, .Rows - 1, mSpecColumn.���_�ͻ���λ) = mlngColor
                    .Cell(flexcpBackColor, 1, mSpecColumn.���_�ͻ���װ, .Rows - 1, mSpecColumn.���_�ͻ���װ) = mlngColor
                End If
                
                If mstrNode Like "�в�ҩ*" Then
                    If InStr(1, mstrPrivs, "��ҩ�ְ�����") = 0 Then
                        .Cell(flexcpBackColor, 1, mSpecColumn.���_��ҩ��̬, .Rows - 1, mSpecColumn.���_��ҩ��̬) = mlngColor
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub Save()
    '���ݱ��淽��
    Dim i As Integer
    Dim strTemp As String
    Dim j As Integer
    Dim m As Integer
    Dim n As Integer
    Dim intupdate As Integer
    Dim rsRecord As ADODB.Recordset
    Dim str���� As String
    Dim intCount As Integer
    Dim bln�޸� As Boolean
    Dim lng���� As Long
    Dim lngSave As Long
    Dim intTemp As Integer
    
    bln�޸� = Check�޸�
    On Error GoTo ErrHandle
    If bln�޸� = False Then 'û���޸ĵĻ�ֱ���˳������б���
        Exit Sub
    End If
    
    If mintExit <> 2 Then
        lngSave = MsgBox("ȷ�����棿", vbInformation + vbYesNo, gstrSysName)
        If lngSave = vbNo Then
            Exit Sub
        End If
        mintExit = 0
    End If
    With vsfDetails
        If mint״̬ = 1 Then    'Ʒ��
            If .TextMatrix(1, mVaricolumn.Ʒ��_id) = "" Then Exit Sub
            '������ݵĺϷ���
            If CheckData = False Then Exit Sub
            
            If mstrNode Like "�в�ҩ*" Then '�в�ҩ
                For i = 1 To .Rows - 1
                    gstrSql = ""
                    strTemp = ""
                    gstrSql = "Zl_��ҩƷ��_Update (" & .TextMatrix(i, mVaricolumn.Ʒ��_����id) & ","
                    gstrSql = gstrSql + .TextMatrix(i, mVaricolumn.Ʒ��_id) & ","
                    gstrSql = gstrSql + "'" + .TextMatrix(i, mVaricolumn.Ʒ��_ҩƷ����) + "'" & ","
                    gstrSql = gstrSql + "'" + .TextMatrix(i, mVaricolumn.Ʒ��_ͨ������) + "'" & ","
                    gstrSql = gstrSql + "'" + .TextMatrix(i, mVaricolumn.Ʒ��_ƴ����) + "'" & ","
                    gstrSql = gstrSql + "'" + .TextMatrix(i, mVaricolumn.Ʒ��_�����) + "'" & ","
                    gstrSql = gstrSql + "'" + .TextMatrix(i, mVaricolumn.Ʒ��_Ӣ������) + "'" & ","
                    gstrSql = gstrSql + "'" + .TextMatrix(i, mVaricolumn.Ʒ��_������λ) + "'" & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.Ʒ��_�������), InStr(1, .TextMatrix(i, mVaricolumn.Ʒ��_�������), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.Ʒ��_��ֵ����), InStr(1, .TextMatrix(i, mVaricolumn.Ʒ��_��ֵ����), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.Ʒ��_��Դ���), InStr(1, .TextMatrix(i, mVaricolumn.Ʒ��_��Դ���), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.Ʒ��_��ҩ�ݴ�), InStr(1, .TextMatrix(i, mVaricolumn.Ʒ��_��ҩ�ݴ�), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.Ʒ��_ҩƷ����), 1, 1)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.Ʒ��_����ְ��), 1, 1) + Mid(.TextMatrix(i, mVaricolumn.Ʒ��_ҽ��ְ��), 1, 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    gstrSql = gstrSql + .TextMatrix(i, mVaricolumn.Ʒ��_��������) & ","
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.Ʒ��_��ζʹ��) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.Ʒ��_ԭ��ҩ) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.Ʒ��_�����Ա�), 1, 1)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.Ʒ��_�ο���Ŀ) = "", "Null", IIf(.TextMatrix(i, mVaricolumn.Ʒ��_�ο���ĿID) = "", "Null", .TextMatrix(i, mVaricolumn.Ʒ��_�ο���ĿID)))
                    gstrSql = gstrSql + strTemp & ","
                    
                    str���� = "Select distinct n.���� as ҩƷ����, p.���� As ƴ��, w.���� As ���" & _
                              "  From (Select Distinct ������Ŀid,���� From ������Ŀ���� Where  ���� = 9) N," & _
                                    " (Select ����, ���� From ������Ŀ���� Where  ���� = 9 And ���� = 1) P," & _
                                    " (Select ����, ���� From ������Ŀ���� Where  ���� = 9 And ���� = 2) W" & _
                               " Where n.���� = p.����(+) And n.���� = w.����(+) and n.������Ŀid = [1]"
                    Set rsRecord = zlDatabase.OpenSQLRecord(str����, "Ʒ�ֱ���", .TextMatrix(i, mVaricolumn.Ʒ��_id))
                    
                    strTemp = ""
                    If Not rsRecord.EOF Then
                        Do While Not rsRecord.EOF
                            strTemp = strTemp & "|" & rsRecord!ҩƷ���� & "^" & rsRecord!ƴ�� & "^" & rsRecord!���
                            rsRecord.MoveNext
                        Loop
                    End If

                    If strTemp <> "" Then
                        strTemp = Mid(strTemp, 2)
                        gstrSql = gstrSql + "'" + strTemp + "'" & ")"
                    Else
                        strTemp = "Null"
                        gstrSql = gstrSql + strTemp
                    End If

                    strTemp = ",NULL," & IIf(.TextMatrix(i, mVaricolumn.Ʒ��_������ҩ) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ")"
                    
                    zlDatabase.ExecuteProcedure gstrSql, "����"
                Next
                '��������
            Else    '����ҩ���г�ҩ
                For i = 1 To vsfDetails.Rows - 1
                    gstrSql = ""
                    strTemp = ""
                    
                    gstrSql = "Zl_��ҩƷ��_Update (" & .TextMatrix(i, mVaricolumn.Ʒ��_����id) & ","
                    gstrSql = gstrSql + .TextMatrix(i, mVaricolumn.Ʒ��_id) & ","
                    gstrSql = gstrSql + "'" + .TextMatrix(i, mVaricolumn.Ʒ��_ҩƷ����) + "'" & ","
                    gstrSql = gstrSql + "'" + .TextMatrix(i, mVaricolumn.Ʒ��_ͨ������) + "'" & ","
                    gstrSql = gstrSql + "'" + .TextMatrix(i, mVaricolumn.Ʒ��_ƴ����) + "'" & ","
                    gstrSql = gstrSql + "'" + .TextMatrix(i, mVaricolumn.Ʒ��_�����) + "'" & ","
                    gstrSql = gstrSql + "'" + .TextMatrix(i, mVaricolumn.Ʒ��_Ӣ������) + "'" & ","
                    gstrSql = gstrSql + "'" + .TextMatrix(i, mVaricolumn.Ʒ��_������λ) + "'" & ","
                    
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.Ʒ��_����), InStr(1, .TextMatrix(i, mVaricolumn.Ʒ��_����), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.Ʒ��_�������), InStr(1, .TextMatrix(i, mVaricolumn.Ʒ��_�������), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.Ʒ��_��ֵ����), InStr(1, .TextMatrix(i, mVaricolumn.Ʒ��_��ֵ����), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.Ʒ��_��Դ���), InStr(1, .TextMatrix(i, mVaricolumn.Ʒ��_��Դ���), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.Ʒ��_��ҩ�ݴ�), InStr(1, .TextMatrix(i, mVaricolumn.Ʒ��_��ҩ�ݴ�), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.Ʒ��_ҩƷ����), 1, 1)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.Ʒ��_����ְ��), 1, 1) + Mid(.TextMatrix(i, mVaricolumn.Ʒ��_ҽ��ְ��), 1, 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    gstrSql = gstrSql + .TextMatrix(i, mVaricolumn.Ʒ��_��������) & ","
                    
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.Ʒ��_����ҩ) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.Ʒ��_��ҩ) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.Ʒ��_ԭ��ҩ) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.Ʒ��_Ƥ��) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.Ʒ��_������), 1, 1)
                    gstrSql = gstrSql + strTemp & ","
                    
                    '�ο�Ŀ¼id
                    '''''''''''''''''''''
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.Ʒ��_�ο���Ŀ) = "", "Null", IIf(.TextMatrix(i, mVaricolumn.Ʒ��_�ο���ĿID) = "", "Null", .TextMatrix(i, mVaricolumn.Ʒ��_�ο���ĿID)))
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.Ʒ��_Ʒ���³���ҽ��) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = Mid(.TextMatrix(i, mVaricolumn.Ʒ��_�����Ա�), 1, 1)
                    gstrSql = gstrSql + strTemp & ","
                    
                    '����
                    str���� = "Select distinct n.���� as ҩƷ����, p.���� As ƴ��, w.���� As ���" & _
                              "  From (Select Distinct ������Ŀid,���� From ������Ŀ���� Where  ���� = 9) N," & _
                                    " (Select ����, ���� From ������Ŀ���� Where  ���� = 9 And ���� = 1) P," & _
                                    " (Select ����, ���� From ������Ŀ���� Where  ���� = 9 And ���� = 2) W" & _
                               " Where n.���� = p.����(+) And n.���� = w.����(+) and n.������Ŀid = [1]"
                    Set rsRecord = zlDatabase.OpenSQLRecord(str����, "Ʒ�ֱ���", .TextMatrix(i, mVaricolumn.Ʒ��_id))
                    
                    strTemp = ""
                    If Not rsRecord.EOF Then
                        Do While Not rsRecord.EOF
                            strTemp = strTemp & "|" & rsRecord!ҩƷ���� & "^" & rsRecord!ƴ�� & "^" & rsRecord!���
                            rsRecord.MoveNext
                        Loop
                    End If
                    
                    If strTemp <> "" Then
                        strTemp = Mid(strTemp, 2)
                        gstrSql = gstrSql + "'" & strTemp & "',"
                    Else
                        strTemp = "Null"
                        gstrSql = gstrSql + strTemp & ","
                    End If
                    gstrSql = gstrSql + "Null,"
                    gstrSql = gstrSql + "'" & .TextMatrix(i, mVaricolumn.Ʒ��_ATCCODE) & "',"
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.Ʒ��_����ҩ) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.Ʒ��_��ý) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.Ʒ��_ԭ��ҩ) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.Ʒ��_ר��ҩ) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.Ʒ��_��������) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mVaricolumn.Ʒ��_������ҩ) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ")"
                    
                    zlDatabase.ExecuteProcedure gstrSql, "����"
                Next
            End If
        Else    '���
            If .TextMatrix(1, mSpecColumn.���_id) = "" Then Exit Sub
            '������ݵĺϷ���
            If CheckData = False Then Exit Sub
            
            For i = 1 To vsfDetails.Rows - 1
                If .TextMatrix(i, mSpecColumn.���_ҩƷ���) = "" Then
                    MsgBox "��" & i & "��ҩƷ���Ϊ�գ�������ҩƷ���", vbExclamation, gstrSysName
                    Exit Sub
                End If
            Next
            
            If mstrNode Like "�в�ҩ*" Then '�в�ҩ
                For i = 1 To vsfDetails.Rows - 1
                    gstrSql = "zl_��ҩ���_Update(" & .TextMatrix(i, mSpecColumn.���_id) & ","
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.���_������) & "',"
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.���_ҩƷ���) & "',"
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.���_��������) & "',"
                    
                    If .TextMatrix(i, mSpecColumn.���_��Ʒ����) <> "" Then
                        strTemp = "'" & .TextMatrix(i, mSpecColumn.���_��Ʒ����) & "'"
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    
                    strTemp = "Null"    'ƴ����
                    gstrSql = gstrSql & strTemp & ","
                    
                    strTemp = "Null"    '�����
                    gstrSql = gstrSql & strTemp & ","
                    
                    If .TextMatrix(i, mSpecColumn.���_������) <> "" Then
                        strTemp = "'" & .TextMatrix(i, mSpecColumn.���_������) & "'"
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    
                    If .TextMatrix(i, mSpecColumn.���_��ʶ��) <> "" Then
                        strTemp = "'" & .TextMatrix(i, mSpecColumn.���_��ʶ��) & "'"
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    
                    If .TextMatrix(i, mSpecColumn.���_��Դ����) <> "" Then
                        strTemp = "'" & Mid(.TextMatrix(i, mSpecColumn.���_��Դ����), InStr(1, .TextMatrix(i, mSpecColumn.���_��Դ����), "-") + 1) & "'"
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql + strTemp + ","
                    
                    If .TextMatrix(i, mSpecColumn.���_��׼�ĺ�) <> "" Then
                        strTemp = "'" & .TextMatrix(i, mSpecColumn.���_��׼�ĺ�) & "'"
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    
                    If .TextMatrix(i, mSpecColumn.���_ע���̱�) <> "" Then
                        strTemp = "'" & .TextMatrix(i, mSpecColumn.���_ע���̱�) & "'"
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.���_�ۼ۵�λ) & "',"
                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.���_����ϵ��) & ","
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.���_���ﵥλ) & "',"
                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.���_����ϵ��) & ","
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.���_ҩ�ⵥλ) & "',"
                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.���_ҩ��ϵ��) & ","
                    
                    Select Case .TextMatrix(i, mSpecColumn.���_���쵥λ)
                        Case "�ۼ۵�λ"
                            strTemp = 1
                        Case "סԺ��λ"
                            strTemp = 2
                        Case "���ﵥλ"
                            strTemp = 3
                        Case "ҩ�ⵥλ"
                            strTemp = 4
                    End Select
                    gstrSql = gstrSql & strTemp & ","
                    
                    If Trim(.TextMatrix(i, mSpecColumn.���_���췧ֵ)) <> "" Then
                        strTemp = .TextMatrix(i, mSpecColumn.���_���췧ֵ)
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    
                    strTemp = Mid(.TextMatrix(i, mSpecColumn.���_ҩ������), 1, 1)
                    gstrSql = gstrSql & strTemp & ","

                    If mint��ǰ��λ <> 0 Then
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.���_�ɹ��޼�) / Nvl(.TextMatrix(i, mSpecColumn.���_ҩ��ϵ��), 1) & ","
                    Else
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.���_�ɹ��޼�) & ","
                    End If
                    
                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.���_�ɹ�����) & ","
                    
                    If mint��ǰ��λ <> 0 Then
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.���_ָ���ۼ�) / Nvl(.TextMatrix(i, mSpecColumn.���_ҩ��ϵ��), 1) & ","
                    Else
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.���_ָ���ۼ�) & ","
                    End If
                    
                    gstrSql = gstrSql & Val(.TextMatrix(i, mSpecColumn.���_�ӳ���)) & ","
                    
                    If .TextMatrix(i, mSpecColumn.���_����ѱ���) <> "" Then
                        strTemp = .TextMatrix(i, mSpecColumn.���_����ѱ���)
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql & strTemp & ","
                        
                    strTemp = Mid(.TextMatrix(i, mSpecColumn.���_ҩ�ۼ���), InStr(1, .TextMatrix(i, mSpecColumn.���_ҩ�ۼ���), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    
                    strTemp = Mid(.TextMatrix(i, mSpecColumn.���_ҽ������), InStr(1, .TextMatrix(i, mSpecColumn.���_ҽ������), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    
                    strTemp = Mid(.TextMatrix(i, mSpecColumn.���_�������), 1, 1)
                    gstrSql = gstrSql & strTemp & ","
                    
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.���_GMP��֤) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    
                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.���_�б�ҩƷ) & ","
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.���_���ηѱ�) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = Mid(.TextMatrix(i, mSpecColumn.���_סԺ����ʹ��), 1, 1)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.���_ҩ�����) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.���_ҩ������) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.���_������) & ","
                    
                    If .TextMatrix(i, mSpecColumn.���_�������) <> "" Then
                        strTemp = .TextMatrix(i, mSpecColumn.���_�������)
                    Else
                       strTemp = "Null"
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    
                    If mint��ǰ��λ <> 0 Then
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.���_�ɱ��۸�) / Nvl(.TextMatrix(i, mSpecColumn.���_ҩ��ϵ��), 1) & ","
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.���_��ǰ�ۼ�) / Nvl(.TextMatrix(i, mSpecColumn.���_ҩ��ϵ��), 1) & ","
                    Else
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.���_�ɱ��۸�) & ","
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.���_��ǰ�ۼ�) & ","
                    End If
                    
                    
                    strTemp = .TextMatrix(i, mSpecColumn.���_������Ŀid)
                    gstrSql = gstrSql & strTemp & ","
                    
                    gstrSql = gstrSql & IIf(.TextMatrix(i, mSpecColumn.���_��ͬ��λid) = "", "Null", .TextMatrix(i, mSpecColumn.���_��ͬ��λid)) & ","
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.���_��ʶ˵��) & "',"
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.���_סԺ��̬����) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.���_��ҩ����) & "',"
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.���_��ѡ��) & "',"
                    
                    If .TextMatrix(i, mSpecColumn.���_��ֵ˰��) <> "" Then
                        strTemp = Mid(.TextMatrix(i, mSpecColumn.���_��ֵ˰��), 1, 1)
                    Else
                       strTemp = "Null"
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    
                    If .TextMatrix(i, mSpecColumn.���_����ҩ��) <> "" Then
                        strTemp = "'" & Mid(.TextMatrix(i, mSpecColumn.���_����ҩ��), 1, 1) & "'"
                    Else
                       strTemp = "Null"
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    
                    Select Case .TextMatrix(i, mSpecColumn.���_��ҩ��̬)
                        Case "ɢװ"
                            strTemp = 0
                        Case "��ҩ��Ƭ"
                            strTemp = 1
                        Case "����"
                            strTemp = 2
                    End Select
                    gstrSql = gstrSql + strTemp & ","
                        
                    If .TextMatrix(i, mSpecColumn.���_վ����) <> "" Then
                        strTemp = Mid(.TextMatrix(i, mSpecColumn.���_վ����), 1, 1)
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.���_�ǳ���ҩ) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    
                    If .TextMatrix(i, mSpecColumn.���_������Ŀ) <> "" Then
                        strTemp = "'" & .TextMatrix(i, mSpecColumn.���_������Ŀ) & "'"
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql + strTemp & ","
                    
                    strTemp = Mid(.TextMatrix(i, mSpecColumn.���_�������ʹ��), 1, 1)
                    gstrSql = gstrSql + strTemp & ","
                    
                    gstrSql = gstrSql + "'" & Trim(.TextMatrix(i, mSpecColumn.���_�ͻ���λ)) & "',"
                    strTemp = IIf(Trim(.TextMatrix(i, mSpecColumn.���_�ͻ���װ)) = "", "Null", Trim(.TextMatrix(i, mSpecColumn.���_�ͻ���װ)))
                    gstrSql = gstrSql + strTemp & ","
                    
                    strTemp = "'" & .TextMatrix(i, mSpecColumn.���_��λ��) & "'"
                    gstrSql = gstrSql + strTemp & ")"
                    
                                                            
                    zlDatabase.ExecuteProcedure gstrSql, "��ҩ��񱣴�"
                Next
            Else    '����ҩ���г�ҩ
                For i = 1 To vsfDetails.Rows - 1
                    gstrSql = ""
                    gstrSql = "zl_��ҩ���_Update(" & .TextMatrix(i, mSpecColumn.���_id) & ","
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.���_������) & "',"
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.���_ҩƷ���) & "',"
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.���_��������) & "',"
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.���_��Ʒ����) & "',"
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.���_ƴ����) & "',"
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.���_�����) & "',"
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.���_������) & "',"
                    
                    If Trim(.TextMatrix(i, mSpecColumn.���_��ʶ��)) = "" Then
                        strTemp = "Null"
                    Else
                        strTemp = "'" & .TextMatrix(i, mSpecColumn.���_��ʶ��) & "'"
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    
                    strTemp = Mid(.TextMatrix(i, mSpecColumn.���_��Դ����), InStr(1, .TextMatrix(i, mSpecColumn.���_��Դ����), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.���_��׼�ĺ�) & "',"
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.���_ע���̱�) & "',"
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.���_�ۼ۵�λ) & "',"
                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.���_����ϵ��) & ","
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.���_���ﵥλ) & "',"
                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.���_����ϵ��) & ","
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.���_סԺ��λ) & "',"
                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.���_סԺϵ��) & ","
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.���_ҩ�ⵥλ) & "',"
                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.���_ҩ��ϵ��) & ","
                    
                    Select Case .TextMatrix(i, mSpecColumn.���_���쵥λ)
                        Case "�ۼ۵�λ"
                            strTemp = 1
                        Case "סԺ��λ"
                            strTemp = 2
                        Case "���ﵥλ"
                            strTemp = 3
                        Case "ҩ�ⵥλ"
                            strTemp = 4
                    End Select
                    gstrSql = gstrSql & strTemp & ","
                    
                    If Trim(.TextMatrix(i, mSpecColumn.���_���췧ֵ)) = "" Then
                        strTemp = "Null"
                    Else
                        strTemp = .TextMatrix(i, mSpecColumn.���_���췧ֵ)
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    
                    strTemp = Mid(.TextMatrix(i, mSpecColumn.���_ҩ������), 1, 1)
                    gstrSql = gstrSql & strTemp & ","
                    
                    If mint��ǰ��λ <> 0 Then
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.���_�ɹ��޼�) / Nvl(.TextMatrix(i, mSpecColumn.���_ҩ��ϵ��), 1) & ","
                    Else
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.���_�ɹ��޼�) & ","
                    End If
                    
                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.���_�ɹ�����) & ","
                    
                    If mint��ǰ��λ <> 0 Then
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.���_ָ���ۼ�) / Nvl(.TextMatrix(i, mSpecColumn.���_ҩ��ϵ��), 1) & ","
                    Else
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.���_ָ���ۼ�) & ","
                    End If
                    
                    gstrSql = gstrSql & Val(.TextMatrix(i, mSpecColumn.���_�ӳ���)) & ","
                    
                    If .TextMatrix(i, mSpecColumn.���_����ѱ���) = "" Then
                        strTemp = "Null"
                    Else
                        strTemp = .TextMatrix(i, mSpecColumn.���_����ѱ���)
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    
                    strTemp = Mid(.TextMatrix(i, mSpecColumn.���_ҩ�ۼ���), InStr(1, .TextMatrix(i, mSpecColumn.���_ҩ�ۼ���), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    
                    strTemp = Mid(.TextMatrix(i, mSpecColumn.���_ҽ������), InStr(1, .TextMatrix(i, mSpecColumn.���_ҽ������), "-") + 1)
                    gstrSql = gstrSql + "'" + strTemp + "'" & ","
                    
                    strTemp = Mid(.TextMatrix(i, mSpecColumn.���_�������), 1, 1)
                    gstrSql = gstrSql & strTemp & ","
                    
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.���_GMP��֤) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    
                    gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.���_�б�ҩƷ) & ","
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.���_���ηѱ�) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    
                    If .TextMatrix(i, mSpecColumn.���_סԺ����ʹ��) <> "" Then
                        strTemp = Mid(.TextMatrix(i, mSpecColumn.���_סԺ����ʹ��), 1, 1)
                        If strTemp = 0 Then
                            strTemp = "0"
                        ElseIf strTemp = 1 Then
                            strTemp = "1"
                        ElseIf strTemp = 2 Then
                            strTemp = "2"
                        ElseIf strTemp = 3 Then
                            strTemp = "-1"
                        ElseIf strTemp = 4 Then
                            strTemp = "-2"
                        ElseIf strTemp = 5 Then
                            strTemp = "-3"
                        End If
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql + strTemp & ","
                    
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.���_ҩ�����) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.���_ҩ������) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    gstrSql = gstrSql & Val(.TextMatrix(i, mSpecColumn.���_������)) & ","
                    gstrSql = gstrSql & IIf(.TextMatrix(i, mSpecColumn.���_�������) = "", "null", .TextMatrix(i, mSpecColumn.���_�������)) & ","
                    
                    If mint��ǰ��λ <> 0 Then
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.���_�ɱ��۸�) / Nvl(.TextMatrix(i, mSpecColumn.���_ҩ��ϵ��), 1) & ","
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.���_��ǰ�ۼ�) / Nvl(.TextMatrix(i, mSpecColumn.���_ҩ��ϵ��), 1) & ","
                    Else
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.���_�ɱ��۸�) & ","
                        gstrSql = gstrSql & .TextMatrix(i, mSpecColumn.���_��ǰ�ۼ�) & ","
                    End If
                    
                    strTemp = .TextMatrix(i, mSpecColumn.���_������Ŀid)
                    gstrSql = gstrSql & strTemp & ","
                    
                    gstrSql = gstrSql & IIf(.TextMatrix(i, mSpecColumn.���_��ͬ��λid) = "", "Null", .TextMatrix(i, mSpecColumn.���_��ͬ��λid)) & ","
                    gstrSql = gstrSql & "'" & .TextMatrix(i, mSpecColumn.���_��ʶ˵��) & "',"
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.���_סԺ��̬����) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    
                    If .TextMatrix(i, mSpecColumn.���_��ҩ����) = "" Then
                        strTemp = "Null"
                    Else
                        strTemp = "'" & .TextMatrix(i, mSpecColumn.���_��ҩ����) & "'"
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    
                    If .TextMatrix(i, mSpecColumn.���_��ѡ��) = "" Then
                        strTemp = "Null"
                    Else
                        strTemp = "'" & .TextMatrix(i, mSpecColumn.���_��ѡ��) & "'"
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    
                    If .TextMatrix(i, mSpecColumn.���_��ֵ˰��) = "" Then
                        strTemp = "Null"
                    Else
                        strTemp = .TextMatrix(i, mSpecColumn.���_��ֵ˰��)
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    
                    If .TextMatrix(i, mSpecColumn.���_����ҩ��) <> "" Then
                        strTemp = "'" & .TextMatrix(i, mSpecColumn.���_����ҩ��) & "'"
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    
                    If .TextMatrix(i, mSpecColumn.���_վ����) <> "" Then
                        strTemp = Mid(.TextMatrix(i, mSpecColumn.���_վ����), 1, 1)
                    Else
                       strTemp = "Null"
                    End If
                    gstrSql = gstrSql & strTemp & ","
                    
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.���_�ǳ���ҩ) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    
                    If Trim(.TextMatrix(i, mSpecColumn.���_�洢�¶�)) <> "" Then
                        strTemp = Mid(.TextMatrix(i, mSpecColumn.���_�洢�¶�), 1, 1)
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql + strTemp & ","
                    
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.���_�洢����) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    
                    If Trim(.TextMatrix(i, mSpecColumn.���_��ҩ����)) <> "" Then
                        strTemp = "'" & .TextMatrix(i, mSpecColumn.���_��ҩ����) & "'"
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql + strTemp & ","
                    
                    strTemp = IIf(.TextMatrix(i, mSpecColumn.���_�������) Like "*1", 1, 0)
                    gstrSql = gstrSql + strTemp & ","
                    
                    If Trim(.TextMatrix(i, mSpecColumn.���_����)) <> "" Then
                        strTemp = .TextMatrix(i, mSpecColumn.���_����)
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql + strTemp & ","
                    
                    If .TextMatrix(i, mSpecColumn.���_������Ŀ) <> "" Then
                        strTemp = "'" & .TextMatrix(i, mSpecColumn.���_������Ŀ) & "'"
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql + strTemp & ","
                    
                    If .TextMatrix(i, mSpecColumn.���_�������ʹ��) <> "" Then
                        strTemp = Mid(.TextMatrix(i, mSpecColumn.���_�������ʹ��), 1, 1)
                        If strTemp = 0 Then
                            strTemp = "0"
                        ElseIf strTemp = 1 Then
                            strTemp = "1"
                        ElseIf strTemp = 2 Then
                            strTemp = "2"
                        ElseIf strTemp = 3 Then
                            strTemp = "-1"
                        ElseIf strTemp = 4 Then
                            strTemp = "-2"
                        ElseIf strTemp = 5 Then
                            strTemp = "-3"
                        End If
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql + strTemp & ","
                    
                    If Trim(.TextMatrix(i, mSpecColumn.���_DDDֵ)) <> "" Then
                        strTemp = .TextMatrix(i, mSpecColumn.���_DDDֵ)
                    Else
                        strTemp = "Null"
                    End If
                    gstrSql = gstrSql + strTemp & ","
                    
                    strTemp = IIf(Trim(.TextMatrix(i, mSpecColumn.���_��ΣҩƷ)) = "", 0, Mid(Trim(.TextMatrix(i, mSpecColumn.���_��ΣҩƷ)), 1, 1))
                    gstrSql = gstrSql + strTemp & ","
                    
                    gstrSql = gstrSql + "'" & Trim(.TextMatrix(i, mSpecColumn.���_�ͻ���λ)) & "',"
                    strTemp = IIf(Trim(.TextMatrix(i, mSpecColumn.���_�ͻ���װ)) = "", "Null", Trim(.TextMatrix(i, mSpecColumn.���_�ͻ���װ)))
                    gstrSql = gstrSql + strTemp & ","
                    gstrSql = gstrSql + "'" & Trim(.TextMatrix(i, mSpecColumn.���_��Һע������)) & "',"
                    
                    strTemp = "'" & .TextMatrix(i, mSpecColumn.���_��λ��) & "'"
                    gstrSql = gstrSql + strTemp & ")"
                    
                    zlDatabase.ExecuteProcedure gstrSql, "��񱣴�"
                Next
            End If
        End If
    End With
    Call Recover    '�����ˢ�½���
    Call tvwDetails_NodeClick(tvwDetails.SelectedItem)
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Recover()
    'ʹ�����иı����ɫ�������廹ԭ
    Dim i As Integer
    Dim j As Integer
    
    With vsfDetails
        For i = 1 To .Rows - 1
            For j = 1 To .Cols - 1
               If .Cell(flexcpBackColor, i, j) <> mlngColor Then
                    .Cell(flexcpBackColor, i, j) = cstcolor_backcolor
                    .Cell(flexcpForeColor, i, j) = vbBlack
                    .Cell(flexcpFontSize, i, j) = 9
                    .Cell(flexcpFontBold, i, j) = False
                End If
                If j = mSpecColumn.���_������ And mint״̬ = 2 Then
                    .Cell(flexcpForeColor, i, j) = vbBlack
                    .Cell(flexcpFontSize, i, j) = 9
                    .Cell(flexcpFontBold, i, j) = False
                End If
                If .Cell(flexcpForeColor, i, j) = mlngApplyColor Then
                    .Cell(flexcpForeColor, i, j) = vbBlack
                    .Cell(flexcpFontSize, i, j) = 9
                    .Cell(flexcpFontBold, i, j) = False
                End If
            Next
        Next
    End With
End Sub

Private Sub SetBatch()
    '��������ÿһ�е�ֵ
    Dim i As Integer
    
    With vsfDetails
        For i = 1 To .Rows - 1
            If .Cell(flexcpBackColor, i) <> mlngColor Then 'ֻ���ڱ�����ɫ���ǻ�ɫ������²��ܽ�������
                .TextMatrix(i, .Col) = .TextMatrix(.Row, .Col)
                If mint״̬ = 1 Then   'Ʒ��
                    If .Col = mVaricolumn.Ʒ��_�ο���Ŀ Then
                        .TextMatrix(i, mVaricolumn.Ʒ��_�ο���ĿID) = .TextMatrix(.Row, mVaricolumn.Ʒ��_�ο���ĿID)
                    End If
                End If
                If mint״̬ <> 1 Then   '���
                    If .Col = mSpecColumn.���_������Ŀ Then
                        .TextMatrix(i, mSpecColumn.���_������Ŀid) = .TextMatrix(.Row, mSpecColumn.���_������Ŀid)
                    End If
                End If
                .Cell(flexcpForeColor, i, .Col) = mlngApplyColor
                .Cell(flexcpFontSize, i, .Col) = 10
                .Cell(flexcpFontBold, i, .Col) = True
                
                
                If .ColDataType(.Col) = 11 Then '��ѡ��
                    .Cell(flexcpBackColor, i, .Col) = mlngApplyColor
                    If mint״̬ <> 1 Then   '���
                        If .Col = mSpecColumn.���_ҩ����� Then
                            If .TextMatrix(.Row, .Col) = "0" Then
                                .Cell(flexcpBackColor, i, mSpecColumn.���_ҩ������) = mlngColor
                                .Cell(flexcpBackColor, i, mSpecColumn.���_������) = mlngColor
                                .TextMatrix(i, mSpecColumn.���_������) = 0
                                .Cell(flexcpForeColor, i, mSpecColumn.���_������) = mlngApplyColor
                                .Cell(flexcpFontBold, i, mSpecColumn.���_������) = 10
                            Else
                                .Cell(flexcpBackColor, i, mSpecColumn.���_ҩ������) = cstcolor_backcolor
                                .Cell(flexcpBackColor, i, mSpecColumn.���_������) = cstcolor_backcolor
                                .TextMatrix(i, mSpecColumn.���_������) = 24
                                .Cell(flexcpForeColor, i, mSpecColumn.���_������) = mlngApplyColor
                                .Cell(flexcpFontBold, i, mSpecColumn.���_������) = 10
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End With
End Sub

Public Sub VsfGridColFormat(ByVal objGrid As VSFlexGrid, ByVal intCol As Integer, ByVal strColName As String, _
    ByVal lngColWidth As Long, ByVal intColAlignment As Integer, _
    Optional ByVal strColKey As String = "", Optional ByVal intFixedColAlignment As Integer = 4)
    'vsf�����ã��������п��ж��뷽ʽ���̶��ж��뷽ʽ��Ĭ��Ϊ���ж��룩
    
    With objGrid
        .TextMatrix(0, intCol) = strColName
        .ColWidth(intCol) = lngColWidth
        .ColAlignment(intCol) = intColAlignment
        .ColKey(intCol) = strColKey
        .FixedAlignment(intCol) = intFixedColAlignment
    End With
End Sub

Private Sub GetDefineSize(ByVal rsRecord As ADODB.Recordset)
    '���ܣ��õ����ݿ�ı��ֶεĳ���
    If mblnSetKey = False Then
        mblnSetKey = True
        With vsfDetails
            If mint״̬ = 1 Then
                .ColKey(mVaricolumn.Ʒ��_ͨ������) = rsRecord.Fields("ͨ������").DefinedSize
                .ColKey(mVaricolumn.Ʒ��_Ӣ������) = rsRecord.Fields("Ӣ����").DefinedSize
                .ColKey(mVaricolumn.Ʒ��_ƴ����) = rsRecord.Fields("ƴ����").DefinedSize
                .ColKey(mVaricolumn.Ʒ��_�����) = rsRecord.Fields("�����").DefinedSize
                .ColKey(mVaricolumn.Ʒ��_��������) = rsRecord.Fields("��������").DefinedSize
                .ColKey(mVaricolumn.Ʒ��_������λ) = rsRecord.Fields("������λ").DefinedSize
                .ColKey(mVaricolumn.Ʒ��_ATCCODE) = rsRecord.Fields("ATCCODE").DefinedSize
            Else
                .ColKey(mSpecColumn.���_ҩƷ���) = rsRecord.Fields("���").DefinedSize
                .ColKey(mSpecColumn.���_��λ��) = rsRecord.Fields("��λ��").DefinedSize
                .ColKey(mSpecColumn.���_������) = rsRecord.Fields("������").DefinedSize
                .ColKey(mSpecColumn.���_��ʶ��) = rsRecord.Fields("��ʶ��").DefinedSize
                .ColKey(mSpecColumn.���_��ѡ��) = rsRecord.Fields("��ѡ��").DefinedSize
                .ColKey(mSpecColumn.���_����) = rsRecord.Fields("����").DefinedSize
                .ColKey(mSpecColumn.���_��Ʒ����) = rsRecord.Fields("��Ʒ��").DefinedSize
                .ColKey(mSpecColumn.���_��������) = rsRecord.Fields("��������").DefinedSize
                .ColKey(mSpecColumn.���_ƴ����) = rsRecord.Fields("ƴ����").DefinedSize
                .ColKey(mSpecColumn.���_�����) = rsRecord.Fields("�����").DefinedSize
                .ColKey(mSpecColumn.���_��ͬ��λ) = rsRecord.Fields("��ͬ��λ").DefinedSize
                .ColKey(mSpecColumn.���_��׼�ĺ�) = rsRecord.Fields("��׼�ĺ�").DefinedSize
                .ColKey(mSpecColumn.���_ע���̱�) = rsRecord.Fields("ע���̱�").DefinedSize
                .ColKey(mSpecColumn.���_�ۼ۵�λ) = rsRecord.Fields("�ۼ۵�λ").DefinedSize
                .ColKey(mSpecColumn.���_����ϵ��) = rsRecord.Fields("�ۼ�ϵ��").DefinedSize
                .ColKey(mSpecColumn.���_סԺ��λ) = rsRecord.Fields("סԺ��λ").DefinedSize
                mintLen = Val(rsRecord.Fields("סԺ��λ").DefinedSize)
                .ColKey(mSpecColumn.���_סԺϵ��) = rsRecord.Fields("סԺ��װ").DefinedSize
                .ColKey(mSpecColumn.���_���ﵥλ) = rsRecord.Fields("���ﵥλ").DefinedSize
                .ColKey(mSpecColumn.���_����ϵ��) = rsRecord.Fields("�����װ").DefinedSize
                .ColKey(mSpecColumn.���_ҩ�ⵥλ) = rsRecord.Fields("ҩ�ⵥλ").DefinedSize
                .ColKey(mSpecColumn.���_ҩ��ϵ��) = rsRecord.Fields("ҩ���װ").DefinedSize
                .ColKey(mSpecColumn.���_�ͻ���λ) = rsRecord.Fields("�ͻ���λ").DefinedSize
                .ColKey(mSpecColumn.���_�ͻ���װ) = rsRecord.Fields("�ͻ���װ").DefinedSize
                .ColKey(mSpecColumn.���_���췧ֵ) = rsRecord.Fields("���췧ֵ").DefinedSize
                .ColKey(mSpecColumn.���_�ɹ��޼�) = rsRecord.Fields("�ɹ��޼�").DefinedSize
                .ColKey(mSpecColumn.���_�ɹ�����) = rsRecord.Fields("�ɹ�����").DefinedSize
                .ColKey(mSpecColumn.���_ָ���ۼ�) = rsRecord.Fields("ָ���ۼ�").DefinedSize
                .ColKey(mSpecColumn.���_ָ������) = rsRecord.Fields("ָ������").DefinedSize
                .ColKey(mSpecColumn.���_�������) = rsRecord.Fields("�������").DefinedSize
                .ColKey(mSpecColumn.���_�ɱ��۸�) = rsRecord.Fields("�ɱ���").DefinedSize
                .ColKey(mSpecColumn.���_��ǰ�ۼ�) = rsRecord.Fields("��ǰ�ۼ�").DefinedSize
                .ColKey(mSpecColumn.���_����ѱ���) = rsRecord.Fields("����ѱ���").DefinedSize
                .ColKey(mSpecColumn.���_��ֵ˰��) = rsRecord.Fields("��ֵ˰��").DefinedSize
                .ColKey(mSpecColumn.���_������) = rsRecord.Fields("������").DefinedSize
                .ColKey(mSpecColumn.���_��ʶ˵��) = rsRecord.Fields("��ʶ˵��").DefinedSize
                .ColKey(mSpecColumn.���_��Һע������) = rsRecord.Fields("��Һע������").DefinedSize
            End If
        End With
   End If
End Sub

Public Function zlGetSymbol(strInput As String, Optional bytIsWB As Byte, Optional intOutNum As Integer = 10) As String
    '----------------------------------
    '���ܣ������ַ����ļ���
    '��Σ�strInput-�����ַ�����bytIsWB-�Ƿ����(����Ϊƴ��)
    '���Σ���ȷ�����ַ��������󷵻�"-"
    '----------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    If bytIsWB Then
        strSql = "select zlWBcode('" & strInput & "'," & intOutNum & ") from dual"
    Else
        strSql = "select zlSpellcode('" & strInput & "'," & intOutNum & ") from dual"
    End If
    On Error GoTo ErrHand
    With rsTmp
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, "mdlCISBase", strSql)
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "zlGetSymbol")
'        Call SQLTest
        zlGetSymbol = IIf(IsNull(rsTmp.Fields(0).Value), "", rsTmp.Fields(0).Value)
    End With
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlGetSymbol = "-"
End Function

Private Sub FindGridRow(ByVal strInput As String)
    '�ڿؼ��в�ѯָ����Ʒ�ֺ͹��
    
    Dim lngStart As Long, lngRows As Long
    Dim str���� As String, str���� As String, str���� As String
    Dim str�������� As String
    Dim n As Integer
    Dim blnEnd As Boolean
    Dim lngFindRow As Long
    Dim strFindStyle As String
    Dim strTmp As String
    
    If strInput = "" Then Exit Sub
    '����ҩƷ
    If strInput = mstrFind Then
        '��ʾ������һ����¼
        If mlngFind >= vsfDetails.Rows - 1 Then
            lngStart = 0
        Else
            lngStart = mlngFind
        End If
    Else
        '��ʾ�µĲ���
        lngStart = 0
        mlngFindFirst = 0
        mstrFind = strInput
        
        strFindStyle = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = "0", "%", "")
        
        Set mrsFindName = New ADODB.Recordset
        
        If mint״̬ = 1 Then    'Ʒ��
            gstrSql = "Select Distinct a.Id, a.����" & _
                      "  From ������ĿĿ¼ A, ������Ŀ���� B " & _
                      " Where a.Id = b.������Ŀid And a.��� = [1] "
        Else    '���
            gstrSql = "Select Distinct A.Id,A.���� From �շ���ĿĿ¼ A,�շ���Ŀ���� B" & _
                 " Where A.Id =B.�շ�ϸĿid And A.���=[1] "
        End If

        If IsNumeric(Replace(strInput, "-", "")) Then       '����ȫ�����֣������һ��"-"��ʱֻƥ�����
            gstrSql = gstrSql & " And A.���� Like [2] Or B.���� Like [2] And B.����=3 "
        ElseIf zlCommFun.IsCharAlpha(strInput) Then         '����ȫ����ĸʱֻƥ�����
            gstrSql = gstrSql & " And B.���� Like [3] "
        ElseIf zlCommFun.IsCharChinese(strInput) Then       '����ȫ�Ǻ���ʱֻƥ������
            gstrSql = gstrSql & " And B.���� Like [3] "
        Else
            gstrSql = gstrSql & " And (A.���� Like [2] Or B.���� Like [3] Or B.���� Like [3] )"
        End If
        
        gstrSql = gstrSql & " Order By A.���� "
        
        If mstrNode Like "����ҩ*" Then
            strTmp = "5"
        ElseIf mstrNode Like "�г�ҩ*" Then
            strTmp = "6"
        Else
            strTmp = "7"
        End If
                 
        Set mrsFindName = zlDatabase.OpenSQLRecord(gstrSql, "ȡƥ���ҩƷID", strTmp, strInput & "%", strFindStyle & strInput & "%")
        
        If mrsFindName.RecordCount = 0 Then Exit Sub
    End If
    
    '��ʼ����
    If mrsFindName.State <> adStateOpen Then Exit Sub
    If mrsFindName.RecordCount = 0 Then Exit Sub

    lngStart = lngStart + 1
    lngRows = vsfDetails.Rows - 1
    
    With mrsFindName
        If .EOF Then .MoveFirst
        
        Do While Not .EOF
            If mint״̬ = 1 Then    'Ʒ��
                lngFindRow = vsfDetails.FindRow(!����, lngStart, mVaricolumn.Ʒ��_ҩƷ����, True, True)
            Else    '���
                lngFindRow = vsfDetails.FindRow(!����, lngStart, mSpecColumn.���_������, True, True)
            End If
        
            If lngFindRow > 0 Then
                vsfDetails.SetFocus
                vsfDetails.TopRow = lngFindRow
                vsfDetails.Row = lngFindRow
                
                mlngFind = lngFindRow
                
                '��¼�ҵ��ĵ�1����¼
                If mlngFindFirst = 0 Then mlngFindFirst = mlngFind
                
                mrsFindName.MoveNext
                Exit Do
            End If
            mrsFindName.MoveNext
    
            '��������ˣ��򷵻ص�1����¼
            If .EOF And lngFindRow = -1 Then
                mlngFind = mlngFindFirst
                If vsfDetails.Rows > 1 Then
                    vsfDetails.Row = 1
                End If
            End If
        Loop
    End With
End Sub

Public Function zlGetDigitSign(ByVal lngMediId As Long, ByVal strSpec As String) As String
    '-------------------------------------------------------------
    '���ܣ�����ҩƷͨ�����ơ����͵����ֱ����͹��ǰ��λ��ֵ����������ҩƷ��λ��
    '��Σ�strSpellcode-ͨ�����Ƶ�ƴ���룻strDoseCode:���͵����ֱ����, strSpec�������ֵ
    '���أ�ҩƷ����
    '-------------------------------------------------------------
    Dim rsThis As New ADODB.Recordset
    Dim strSpellcode As String, strDoseCode As String
    Dim strChange As String
    Dim intLocate As Integer
    Dim strTemp As String
    Dim intCount As Integer
    
    gstrSql = "Select ���� From ������Ŀ���� where ������Ŀid=[1] and ����=1 and ����=1"
    Set rsThis = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngMediId)
    
    If rsThis.RecordCount > 0 Then
        strSpellcode = IIf(IsNull(rsThis!����), "", rsThis!����)
    Else
        strSpellcode = ""
    End If
    
    gstrSql = "select P.����� from ҩƷ���� T,ҩƷ���� P where T.ҩƷ����=P.����(+) and ҩ��id=[1]"
    Set rsThis = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngMediId)
    
    If rsThis.RecordCount > 0 Then
        strDoseCode = IIf(IsNull(rsThis!�����), "", rsThis!�����)
    Else
        strDoseCode = ""
    End If

    strChange = "AOEYUVBP MF DT NL GKHJQXZCSRW "
    
    strTemp = ""
    strSpellcode = Mid(strSpellcode, 1, 3)
    For intCount = 1 To Len(strSpellcode)
        intLocate = InStr(1, strChange, Mid(strSpellcode, intCount, 1))
        If intLocate Mod 3 = 0 Then
            intLocate = (intLocate \ 3) - 1
        Else
            intLocate = intLocate \ 3
        End If
        If intLocate <> -1 Then strTemp = strTemp & CStr(intLocate)
    Next
    strTemp = strTemp & strDoseCode & Format(Val(Mid(strSpec, 1, 3)), "000")
    zlGetDigitSign = strTemp
End Function

Private Sub ExitFrom()
    '�˳�ʱ����
    '�жϽ������Ƿ���ֵ�ձ��޸���
    Dim i As Integer
    Dim j As Integer
    Dim intupdate As Integer
    Dim bln�޸� As Boolean
    
    bln�޸� = Check�޸�
    mintExit = 0
    
    If bln�޸� = True Then
        intupdate = MsgBox("�������ݱ��޸��ˣ��˳�֮ǰ�Ƿ񱣴棿", vbQuestion + vbDefaultButton2 + vbYesNo, gstrSysName)
        If intupdate = vbYes Then
            mintExit = 2
            Call Save
            Unload Me
        Else
            Unload Me
        End If
    Else
        Unload Me
    End If
End Sub

Private Function CheckData() As Boolean
    '������ݵĺϷ��Ժ�������
    Dim i As Integer
    Dim j As Integer
    Dim n As Integer
    Dim intupdate As Integer
    Dim rsTmp As ADODB.Recordset
    Dim strMsg As String
    
    On Error GoTo ErrHandle
    With vsfDetails
        If mint״̬ = 1 Then 'Ʒ��
            For i = 1 To .Rows - 1
                If .TextMatrix(i, mVaricolumn.Ʒ��_ͨ������) = "" Then
                    MsgBox "������Ϣҳ��" & i & "��ͨ�����Ʋ���Ϊ�գ������룡", vbInformation, gstrSysName
                    tbcDetails.Item(mVariList.������Ϣ).Selected = True
                    .Select i, mVaricolumn.Ʒ��_ͨ������
                    Exit Function
                End If
                If .TextMatrix(i, mVaricolumn.Ʒ��_������λ) = "" Then
                    MsgBox "�ٴ�Ӧ��ҳ��" & i & "�м�����λ����Ϊ�գ������룡", vbInformation, gstrSysName
                    tbcDetails.Item(mVariList.�ٴ�Ӧ��).Selected = True
                    .Select i, mVaricolumn.Ʒ��_������λ
                    Exit Function
                End If
                For j = 2 To .Rows - 1
                    If .TextMatrix(i, mVaricolumn.Ʒ��_ͨ������) = .TextMatrix(j, mVaricolumn.Ʒ��_ͨ������) And i <> j Then
                        MsgBox "������Ϣҳ��" & i & "��ͨ���������" & j & "��ͨ��������ͬ�ˣ�", vbExclamation, gstrSysName
                        tbcDetails.Item(mVariList.������Ϣ).Selected = True
                        .Select i, mVaricolumn.Ʒ��_ͨ������
                        Exit Function
                    End If
                Next
            Next
        Else    '���
            For i = 1 To .Rows - 1
                If .TextMatrix(i, mSpecColumn.���_ҩƷ���) = "" Then
                    MsgBox "������Ϣҳ��" & i & "��ҩƷ�����Ϊ�գ������룡", vbInformation, gstrSysName
                    tbcDetails.Item(0).Selected = True
                    .Select i, mSpecColumn.���_ҩƷ���
                    Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.���_����)) >= 100000000000# Then
                    MsgBox "������Ϣҳ��" & i & "�������������������룡", vbInformation, gstrSysName
                    tbcDetails.Item(0).Selected = True
                    .Select i, mSpecColumn.���_����
                    Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.���_����ϵ��)) >= 100000 Then
                    MsgBox "��װ��λҳ��" & i & "���ۼ�ϵ���������������룡", vbInformation, gstrSysName
                    tbcDetails.Item(2).Selected = True
                    .Select i, mSpecColumn.���_����ϵ��
                    Exit Function
                End If
                If .TextMatrix(i, mSpecColumn.���_����ϵ��) = "" Then
                    MsgBox "��װ��λҳ��" & i & "�м���ϵ������Ϊ�գ������룡", vbInformation, gstrSysName
                    tbcDetails.Item(2).Selected = True
                    .Select i, mSpecColumn.���_����ϵ��
                    Exit Function
                End If
                If Trim(.TextMatrix(i, mSpecColumn.���_�ۼ۵�λ)) = "" Then
                    MsgBox "��װ��λҳ��" & i & "���ۼ۵�λ����Ϊ�գ������룡", vbInformation, gstrSysName
                    tbcDetails.Item(2).Selected = True
                    .Select i, mSpecColumn.���_�ۼ۵�λ
                    Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.���_����ϵ��)) >= 100000 Then
                    MsgBox "��װ��λҳ��" & i & "������ϵ���������������룡", vbInformation, gstrSysName
                    tbcDetails.Item(2).Selected = True
                    .Select i, mSpecColumn.���_����ϵ��
                    Exit Function
                End If
                If .TextMatrix(i, mSpecColumn.���_����ϵ��) = "" Then
                    MsgBox "��װ��λҳ��" & i & "������ϵ������Ϊ�գ������룡", vbInformation, gstrSysName
                    tbcDetails.Item(2).Selected = True
                    .Select i, mSpecColumn.���_����ϵ��
                    Exit Function
                End If
                 If Trim(.TextMatrix(i, mSpecColumn.���_���ﵥλ)) = "" Then
                    MsgBox "��װ��λҳ��" & i & "�����ﵥλ����Ϊ�գ������룡", vbInformation, gstrSysName
                    tbcDetails.Item(2).Selected = True
                    .Select i, mSpecColumn.���_���ﵥλ
                    Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.���_סԺϵ��)) >= 100000 Then
                    MsgBox "��װ��λҳ��" & i & "��" & IIf(mstrNode Like "�в�ҩ*", "ҩ��ϵ��", "סԺϵ��") & "�������������룡", vbInformation, gstrSysName
                    tbcDetails.Item(2).Selected = True
                    .Select i, mSpecColumn.���_סԺϵ��
                    Exit Function
                End If
                If .TextMatrix(i, mSpecColumn.���_סԺϵ��) = "" Then
                    MsgBox "��װ��λҳ��" & i & "��" & IIf(mstrNode Like "�в�ҩ*", "ҩ��ϵ��", "סԺϵ��") & "����Ϊ�գ������룡", vbInformation, gstrSysName
                    tbcDetails.Item(2).Selected = True
                    .Select i, mSpecColumn.���_סԺϵ��
                    Exit Function
                End If
                If Trim(.TextMatrix(i, mSpecColumn.���_סԺ��λ)) = "" Then
                    MsgBox "��װ��λҳ��" & i & "��" & IIf(mstrNode Like "�в�ҩ*", "ҩ����λ", "סԺ��λ") & "����Ϊ�գ������룡", vbInformation, gstrSysName
                    tbcDetails.Item(2).Selected = True
                    .Select i, mSpecColumn.���_סԺ��λ
                    Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.���_ҩ��ϵ��)) >= 100000 Then
                    MsgBox "��װ��λҳ��" & i & "��ҩ��ϵ���������������룡", vbInformation, gstrSysName
                    tbcDetails.Item(2).Selected = True
                    .Select i, mSpecColumn.���_ҩ��ϵ��
                    Exit Function
                End If
                If .TextMatrix(i, mSpecColumn.���_ҩ��ϵ��) = "" Then
                    MsgBox "��װ��λҳ��" & i & "��ҩ��ϵ������Ϊ�գ������룡", vbInformation, gstrSysName
                    tbcDetails.Item(2).Selected = True
                    .Select i, mSpecColumn.���_ҩ��ϵ��
                    Exit Function
                End If
                If Trim(.TextMatrix(i, mSpecColumn.���_ҩ�ⵥλ)) = "" Then
                    MsgBox "��װ��λҳ��" & i & "��ҩ�ⵥλ����Ϊ�գ������룡", vbInformation, gstrSysName
                    tbcDetails.Item(2).Selected = True
                    .Select i, mSpecColumn.���_ҩ�ⵥλ
                Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.���_���췧ֵ)) >= 100000 Then
                    MsgBox "��װ��λҳ��" & i & "�����췧ֵ�������������룡", vbInformation, gstrSysName
                    tbcDetails.Item(2).Selected = True
                    .Select i, mSpecColumn.���_���췧ֵ
                    Exit Function
                End If
                If .TextMatrix(i, mSpecColumn.���_�ͻ���װ) = "" And Trim(.TextMatrix(i, mSpecColumn.���_�ͻ���λ)) <> "" Then
                    MsgBox "��װ��λҳ��" & i & "�����ͻ���λ����£��ͻ���װ����Ϊ�գ������룡", vbInformation, gstrSysName
                    tbcDetails.Item(2).Selected = True
                    .Select i, mSpecColumn.���_�ͻ���װ
                    Exit Function
                End If
                If .TextMatrix(i, mSpecColumn.���_�ͻ���װ) <> "" And Trim(.TextMatrix(i, mSpecColumn.���_�ͻ���λ)) = "" Then
                    MsgBox "��װ��λҳ��" & i & "�����ͻ���װ����£��ͻ���λ����Ϊ�գ������룡", vbInformation, gstrSysName
                    tbcDetails.Item(2).Selected = True
                    .Select i, mSpecColumn.���_�ͻ���λ
                    Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.���_�ɹ��޼�)) > 1000000 Then
                    MsgBox "�۸���Ϣҳ��" & i & "�вɹ��޼۹������������룡", vbInformation, gstrSysName
                    tbcDetails.Item(3).Selected = True
                    .Select i, mSpecColumn.���_�ɹ��޼�
                    Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.���_ָ���ۼ�)) > 1000000 Then
                    MsgBox "�۸���Ϣҳ��" & i & "��ָ���ۼ۹������������룡", vbInformation, gstrSysName
                    tbcDetails.Item(3).Selected = True
                    .Select i, mSpecColumn.���_ָ���ۼ�
                    Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.���_�ɱ��۸�)) > 1000000 Then
                    MsgBox "�۸���Ϣҳ��" & i & "�гɱ��۸�������������룡", vbInformation, gstrSysName
                    tbcDetails.Item(3).Selected = True
                    .Select i, mSpecColumn.���_�ɱ��۸�
                    Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.���_��ǰ�ۼ�)) > 1000000 Then
                    MsgBox "�۸���Ϣҳ��" & i & "�е�ǰ�ۼ۹������������룡", vbInformation, gstrSysName
                    tbcDetails.Item(3).Selected = True
                    .Select i, mSpecColumn.���_��ǰ�ۼ�
                    Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.���_�ӳ���)) > 1000000 Then
                    MsgBox "�۸���Ϣҳ��" & i & "�мӳ��ʳ��������ֵ�����������룡", vbInformation, gstrSysName
                    tbcDetails.Item(3).Selected = True
                    .Select i, mSpecColumn.���_�ӳ���
                    Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.���_�ɹ�����)) > 100 Then
                    MsgBox "�۸���Ϣҳ��" & i & "�вɹ����ʳ��������ֵ�����������룡", vbInformation, gstrSysName
                    tbcDetails.Item(3).Selected = True
                    .Select i, mSpecColumn.���_�ɹ�����
                    Exit Function
                End If
                If .TextMatrix(i, mSpecColumn.���_�ɹ��޼�) = "" Then
                    MsgBox "�۸���Ϣҳ��" & i & "�вɹ��޼۲���Ϊ�գ������룡", vbInformation, gstrSysName
                    tbcDetails.Item(3).Selected = True
                    .Select i, mSpecColumn.���_�ɹ��޼�
                    Exit Function
                End If
                If .TextMatrix(i, mSpecColumn.���_�ɱ��۸�) = "" Then
                    MsgBox "�۸���Ϣҳ��" & i & "�гɱ��۸���Ϊ�գ������룡", vbInformation, gstrSysName
                    tbcDetails.Item(3).Selected = True
                    .Select i, mSpecColumn.���_�ɱ��۸�
                    Exit Function
                End If
                If .TextMatrix(i, mSpecColumn.���_��ǰ�ۼ�) = "" Then
                    MsgBox "�۸���Ϣҳ��" & i & "�е�ǰ�ۼ۲���Ϊ�գ������룡", vbInformation, gstrSysName
                    tbcDetails.Item(3).Selected = True
                    .Select i, mSpecColumn.���_��ǰ�ۼ�
                    Exit Function
                End If
                If .TextMatrix(i, mSpecColumn.���_ָ���ۼ�) = "" Then
                    MsgBox "�۸���Ϣҳ��" & i & "��ָ���ۼ۲���Ϊ�գ������룡", vbInformation, gstrSysName
                    tbcDetails.Item(3).Selected = True
                    .Select i, mSpecColumn.���_ָ���ۼ�
                    Exit Function
                End If
                If .TextMatrix(i, mSpecColumn.���_�ɹ�����) = "" Then
                    MsgBox "�۸���Ϣҳ��" & i & "�вɹ����ʲ���Ϊ�գ������룡", vbInformation, gstrSysName
                    tbcDetails.Item(3).Selected = True
                    .Select i, mSpecColumn.���_�ɹ�����
                    Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.���_ָ������)) > 100 Then
                    MsgBox "�۸���Ϣҳ��" & i & "��ָ�����ʳ��������ֵ�����������룡", vbInformation, gstrSysName
                    tbcDetails.Item(3).Selected = True
                    .Select i, mSpecColumn.���_ָ������
                    Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.���_����ѱ���)) > 100 Then
                    MsgBox "ҩ������ҳ��" & i & "�й���ѱ������������ֵ�����������룡", vbInformation, gstrSysName
                    tbcDetails.Item(4).Selected = True
                    .Select i, mSpecColumn.���_����ѱ���
                    Exit Function
                End If
                If Val(.TextMatrix(i, mSpecColumn.���_��ֵ˰��)) > 100 Then
                    MsgBox "ҩ������ҳ��" & i & "����ֵ˰�ʳ��������ֵ�����������룡", vbInformation, gstrSysName
                    tbcDetails.Item(4).Selected = True
                    .Select i, mSpecColumn.���_��ֵ˰��
                    Exit Function
                End If
                If CheckUnit(i) = False Then
                    Exit Function
                End If
                If CheckBatches(.TextMatrix(i, mSpecColumn.���_ҩ�����) Like "*1", .TextMatrix(i, mSpecColumn.���_ҩ������) Like "*1") = True Then
                    If Not (.TextMatrix(i, mSpecColumn.���_ԭҩ�����) Like "*1" And .TextMatrix(i, mSpecColumn.���_ԭҩ������) = "0") Then
                        MsgBox "��������ҳ��" & i & "�е�ǰ�в��ŵĹ�������ͬʱ������ҩ��ҩ������ͬʱ����ҩ��ҩ�������򲻷�����", vbInformation, gstrSysName
                        tbcDetails.Item(5).Selected = True
                        .Select i, mSpecColumn.���_ҩ�����
                        Exit Function
                    Else
                        n = n + 1
                        If n < 4 Then
                            strMsg = IIf(strMsg = "", "", strMsg & vbCrLf) & "[" & .TextMatrix(i, mSpecColumn.���_������) & "]" & _
                                                    .TextMatrix(i, mSpecColumn.���_ͨ������) & "(" & .TextMatrix(i, mSpecColumn.���_ҩƷ���) & ")" & "��"
                        End If
                    End If
                End If
                
            Next
        End If
    End With
    
    If strMsg <> "" Then
        MsgBox strMsg & vbCrLf & IIf(n > 3, "�����ϳ���3��", "����" & n & "��") & "ҩƷ������ҩ�����ҩ�������������в��ŵĹ�������ͬʱ������ҩ���ҩ������ע��鿴��", vbExclamation, gstrSysName
    End If
    CheckData = True
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckUnit(ByVal intRow As Integer) As Boolean
    Dim intOut As Integer, intIN As Integer
    Dim arr��λ, arrϵ��
    Dim str��λ As String, strϵ�� As String
    Dim str��λ_Tmp As String, strϵ��_Tmp As String
    Dim intλ�� As Integer
    Dim strTemp As String
    
    With vsfDetails
        '����Ƿ���ڵ�λ����һ������ϵ����һ�µ����
        '����Ƿ����ϵ��һ��������λ���Ʋ�һ�������
        If mstrNode Like "�в�ҩ*" Then
            str��λ = .TextMatrix(intRow, mSpecColumn.���_�ۼ۵�λ) & "|" & .TextMatrix(intRow, mSpecColumn.���_סԺ��λ) & "|" & .TextMatrix(intRow, mSpecColumn.���_ҩ�ⵥλ)
            strϵ�� = .TextMatrix(intRow, mSpecColumn.���_����ϵ��) & "|" & .TextMatrix(intRow, mSpecColumn.���_סԺϵ��) & "|" & .TextMatrix(intRow, mSpecColumn.���_ҩ��ϵ��)
        Else
            str��λ = .TextMatrix(intRow, mSpecColumn.���_�ۼ۵�λ) & "|" & .TextMatrix(intRow, mSpecColumn.���_סԺ��λ) & "|" & .TextMatrix(intRow, mSpecColumn.���_���ﵥλ) & "|" & .TextMatrix(intRow, mSpecColumn.���_ҩ�ⵥλ)
            strϵ�� = .TextMatrix(intRow, mSpecColumn.���_����ϵ��) & "|" & .TextMatrix(intRow, mSpecColumn.���_סԺϵ��) & "|" & .TextMatrix(intRow, mSpecColumn.���_����ϵ��) & "|" & .TextMatrix(intRow, mSpecColumn.���_ҩ��ϵ��)
        End If
                
        '���ǵ�������λ�������ۼ۵�λһ�£���ϵ���϶���һ�£����Ա���ֿ��ж�
        '���ۼ۵�λ��ļ��
        For intOut = 2 To IIf(mstrNode Like "�в�ҩ*" = True, 3, 4)
            If mstrNode Like "�в�ҩ*" Then
                str��λ_Tmp = IIf(intOut = 2, .TextMatrix(intRow, mSpecColumn.���_סԺ��λ), .TextMatrix(intRow, mSpecColumn.���_ҩ�ⵥλ))
                strϵ��_Tmp = Val(IIf(intOut = 2, .TextMatrix(intRow, mSpecColumn.���_סԺϵ��), .TextMatrix(intRow, mSpecColumn.���_ҩ��ϵ��)))
            Else
                str��λ_Tmp = IIf(intOut = 1, .TextMatrix(intRow, mSpecColumn.���_�ۼ۵�λ), IIf(intOut = 2, .TextMatrix(intRow, mSpecColumn.���_סԺ��λ), IIf(intOut = 3, .TextMatrix(intRow, mSpecColumn.���_���ﵥλ), .TextMatrix(intRow, mSpecColumn.���_ҩ�ⵥλ))))
                strϵ��_Tmp = Val(IIf(intOut = 1, .TextMatrix(intRow, mSpecColumn.���_����ϵ��), IIf(intOut = 2, .TextMatrix(intRow, mSpecColumn.���_סԺϵ��), IIf(intOut = 3, .TextMatrix(intRow, mSpecColumn.���_����ϵ��), .TextMatrix(intRow, mSpecColumn.���_ҩ��ϵ��)))))
            End If
            arr��λ = Split(str��λ, "|")
            arrϵ�� = Split(strϵ��, "|")
            For intIN = 2 To IIf(mstrNode Like "�в�ҩ*" = True, 3, 4)
                If intIN <> intOut Then
                    '��λ��ͬϵ����ͬ
                    If str��λ_Tmp = arr��λ(intIN - 1) And (Val(strϵ��_Tmp) <> Val(arrϵ��(intIN - 1))) Then
                        If mstrNode Like "�в�ҩ*" Then
                            strTemp = IIf(intOut = 2, "ҩ��", "ҩ��") & "��λ��" & IIf(intIN = 2, "ҩ��", "ҩ��") & "��λһ�£�����ϵ��ȴ����ͬ�����飡"
                        Else
                            strTemp = IIf(intOut = 2, "סԺ", IIf(intOut = 3, "����", "ҩ��")) & "��λ��" & IIf(intIN = 2, "סԺ", IIf(intIN = 3, "����", "ҩ��")) & "��λһ�£�����ϵ��ȴ����ͬ�����飡"
                        End If
                        
                        MsgBox strTemp, vbInformation, gstrSysName
                        tbcDetails.Item(2).Selected = True
                        If InStr(1, strTemp, "��λ��סԺ") > 0 Then
                            intλ�� = mSpecColumn.���_סԺ��λ
                        ElseIf InStr(1, strTemp, "��λ������") > 0 Then
                            intλ�� = mSpecColumn.���_���ﵥλ
                        ElseIf InStr(1, strTemp, "��λ��ҩ��") > 0 Then
                            intλ�� = mSpecColumn.���_ҩ�ⵥλ
                        ElseIf InStr(1, strTemp, "ҩ����λһ��") > 0 Then
                            intλ�� = mSpecColumn.���_סԺ��λ
                        ElseIf InStr(1, strTemp, "ҩ�ⵥλһ��") > 0 Then
                            intλ�� = mSpecColumn.���_ҩ�ⵥλ
                        End If
                        
                        .Select intRow, intλ��
                        Exit Function
                    End If
                    If str��λ_Tmp <> arr��λ(intIN - 1) And (Val(strϵ��_Tmp) = Val(arrϵ��(intIN - 1))) Then
                        If mstrNode Like "�в�ҩ*" Then
                            strTemp = IIf(intOut = 2, "ҩ��", "ҩ��") & "��װ��" & IIf(intIN = 2, "ҩ��", "ҩ��") & "��װһ�£����䵥λȴ����ͬ�����飡"
                        Else
                            strTemp = IIf(intOut = 2, "סԺ", IIf(intOut = 3, "����", "ҩ��")) & "��װ��" & IIf(intIN = 2, "סԺ", IIf(intIN = 3, "����", "ҩ��")) & "��װһ�£����䵥λȴ����ͬ�����飡"
                        End If
                        
                        MsgBox strTemp, vbInformation, gstrSysName
                        tbcDetails.Item(2).Selected = True
                        
                        If InStr(1, strTemp, "��װ��סԺ") > 0 Then
                            intλ�� = mSpecColumn.���_סԺ��λ
                        ElseIf InStr(1, strTemp, "��װ������") > 0 Then
                            intλ�� = mSpecColumn.���_���ﵥλ
                        ElseIf InStr(1, strTemp, "��װ��ҩ��") > 0 Then
                            intλ�� = mSpecColumn.���_ҩ�ⵥλ
                        ElseIf InStr(1, strTemp, "ҩ����װһ��") > 0 Then
                            intλ�� = mSpecColumn.���_סԺ��λ
                        ElseIf InStr(1, strTemp, "ҩ���װһ��") > 0 Then
                            intλ�� = mSpecColumn.���_ҩ�ⵥλ
                        End If
                        .Select intRow, intλ��
                        Exit Function
                    End If
                End If
            Next
        Next
        
        '����������λ���ۼ۵�λ��ͬ����ϵ����Ϊ1�����
        '����λ���ۼ۵�λ���м��
        For intOut = 2 To IIf(mstrNode Like "�в�ҩ*" = True, 3, 4)
            If mstrNode Like "�в�ҩ*" Then
                str��λ_Tmp = IIf(intOut = 2, .TextMatrix(intRow, mSpecColumn.���_סԺ��λ), .TextMatrix(intRow, mSpecColumn.���_ҩ�ⵥλ))
                strϵ��_Tmp = Val(IIf(intOut = 2, .TextMatrix(intRow, mSpecColumn.���_סԺϵ��), .TextMatrix(intRow, mSpecColumn.���_ҩ��ϵ��)))
            Else
                str��λ_Tmp = IIf(intOut = 1, .TextMatrix(intRow, mSpecColumn.���_�ۼ۵�λ), IIf(intOut = 2, .TextMatrix(intRow, mSpecColumn.���_סԺ��λ), IIf(intOut = 3, .TextMatrix(intRow, mSpecColumn.���_���ﵥλ), .TextMatrix(intRow, mSpecColumn.���_ҩ�ⵥλ))))
                strϵ��_Tmp = Val(IIf(intOut = 1, .TextMatrix(intRow, mSpecColumn.���_����ϵ��), IIf(intOut = 2, .TextMatrix(intRow, mSpecColumn.���_סԺϵ��), IIf(intOut = 3, .TextMatrix(intRow, mSpecColumn.���_����ϵ��), .TextMatrix(intRow, mSpecColumn.���_ҩ��ϵ��)))))
            End If
            
            If str��λ_Tmp = .TextMatrix(intRow, mSpecColumn.���_�ۼ۵�λ) And Val(strϵ��_Tmp) <> 1 Then
                If mstrNode Like "�в�ҩ*" Then
                    strTemp = IIf(intOut = 2, "ҩ��", "ҩ��") & "��λ���ۼ۵�λһ�£�" & IIf(intOut = 2, "ҩ��", "ҩ��") & "ϵ��Ӧ��Ϊ1"
                Else
                    strTemp = IIf(intOut = 2, "סԺ", IIf(intOut = 3, "����", "ҩ��")) & "��λ���ۼ۵�λһ�£�" & IIf(intOut = 2, "סԺ", IIf(intOut = 3, "����", "ҩ��")) & "ϵ��Ӧ��Ϊ1"
                End If
                MsgBox strTemp, vbInformation, gstrSysName
                tbcDetails.Item(2).Selected = True
                
                If InStr(1, strTemp, "סԺϵ��") > 0 Then
                    intλ�� = mSpecColumn.���_סԺ��λ
                ElseIf InStr(1, strTemp, "����ϵ��") > 0 Then
                    intλ�� = mSpecColumn.���_���ﵥλ
                ElseIf InStr(1, strTemp, "ҩ��ϵ��") > 0 Then
                    intλ�� = mSpecColumn.���_ҩ�ⵥλ
                ElseIf InStr(1, strTemp, "ҩ��ϵ��") > 0 Then
                    intλ�� = mSpecColumn.���_סԺ��λ
                ElseIf InStr(1, strTemp, "ҩ��ϵ��") > 0 Then
                    intλ�� = mSpecColumn.���_ҩ�ⵥλ
                End If
                .Select intRow, intλ��
                Exit Function
            End If
        Next
        
    End With
    CheckUnit = True
End Function

'Private Sub ShowPercent(sngPercent As Single)
''����:��״̬���ϸ��ݰٷֱ���ʾ��ǰ�������(��)
'    Dim intAll As Integer
'    intAll = stbThis.Panels(2).Width / TextWidth("��") - 4
'    stbThis.Panels(2).Text = Format(sngPercent, "0% ") & String(intAll * sngPercent, "��")
'End Sub

Private Function Check�޸�() As Boolean
    '�жϽ������Ƿ���ֵ�ձ��޸���
    '����ֵΪtrue �Ѿ��޸��� ����δ�޸�
    Dim i As Integer
    Dim j As Integer
    
    With vsfDetails
        Check�޸� = False
        For i = 1 To .Rows - 1
            For j = 1 To vsfDetails.Cols - 1
                If .Cell(flexcpForeColor, i, j) = mlngApplyColor Or .Cell(flexcpFontSize, i, j) = 10 Or .Cell(flexcpFontBold, i, j) = True Or .Cell(flexcpBackColor, i, j) = mlngApplyColor Then
                    Check�޸� = True
                    Exit Function
                End If
            Next
        Next
    End With
End Function

Private Sub MoveRowCol()
    '�����ƶ�����
    With vsfDetails
        If mint״̬ = 1 Then    'Ʒ��
            If mstrNode Like "�в�ҩ*" Then
                If tbcDetails.Selected.Index = mVariList.������Ϣ Then    '����ҳ��
                    If .Col = mVaricolumn.Ʒ��_����� Then
                        tbcDetails.Item(mVariList.Ʒ������).Selected = True
                        .SetFocus
                        .Col = mVaricolumn.Ʒ��_�������
                    Else
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mVariList.Ʒ������ Then    'Ʒ������
                    If .Col = mVaricolumn.Ʒ��_��ζʹ�� Then
                        tbcDetails.Item(mVariList.�ٴ�Ӧ��).Selected = True
                        .SetFocus
                        .Col = mVaricolumn.Ʒ��_�ο���Ŀ
                    Else
                        If .Col = mVaricolumn.Ʒ��_ͨ������ Then
                            .Col = mVaricolumn.Ʒ��_�������
                        ElseIf .Col = mVaricolumn.Ʒ��_ҩƷ���� Then
                            .Col = mVaricolumn.Ʒ��_������ҩ
                        Else
                            .Col = .Col + 1
                        End If
                    End If
                ElseIf tbcDetails.Selected.Index = mVariList.�ٴ�Ӧ�� Then    '�ٴ�Ӧ��
                    If .Col = mVaricolumn.Ʒ��_������λ And .Row <> .Rows - 1 Then
                        tbcDetails.Item(mVariList.������Ϣ).Selected = True
                        .SetFocus
                        .Row = .Row + 1
                        .Col = mVaricolumn.Ʒ��_ͨ������
                    Else
                        If .Col = mVaricolumn.Ʒ��_ͨ������ Then
                            .Col = mVaricolumn.Ʒ��_�ο���Ŀ
                        Else
                            If .Col <> mVaricolumn.Ʒ��_������λ Then
                                .Col = .Col + 1
                            End If
                        End If
                    End If
                End If
            Else    '����ҩ���г�ҩ
                If tbcDetails.Selected.Index = mVariList.������Ϣ Then    '����ҳ��
                    If .Col = mVaricolumn.Ʒ��_����� Then
                        tbcDetails.Item(mVariList.Ʒ������).Selected = True
                        .SetFocus
                        .Col = mVaricolumn.Ʒ��_�������
                    Else
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mVariList.Ʒ������ Then    'Ʒ������
                    If .Col = mVaricolumn.Ʒ��_ԭ��ҩ Then
                        tbcDetails.Item(mVariList.�ٴ�Ӧ��).Selected = True
                        .SetFocus
                        .Col = mVaricolumn.Ʒ��_�ο���Ŀ
                    Else
                        If .Col = mVaricolumn.Ʒ��_ͨ������ Then
                            .Col = mVaricolumn.Ʒ��_�������
                        Else
                            .Col = .Col + 1
                        End If
                    End If
                ElseIf tbcDetails.Selected.Index = mVariList.�ٴ�Ӧ�� Then    '�ٴ�Ӧ��
                    If .Col = mVaricolumn.Ʒ��_Ʒ���³���ҽ�� And .Row <> .Rows - 1 Then
                        tbcDetails.Item(mVariList.������Ϣ).Selected = True
                        .SetFocus
                        .Row = .Row + 1
                        .Col = 2
                    Else
                        If .Col = mVaricolumn.Ʒ��_ͨ������ Then
                            .Col = mVaricolumn.Ʒ��_�ο���Ŀ
                        Else
                            If .Col <> mVaricolumn.Ʒ��_Ʒ���³���ҽ�� Then
                                .Col = .Col + 1
                            End If
                        End If
                    End If
                End If
            End If
        Else    '���
            If mstrNode Like "�в�ҩ*" Then '�в�ҩ
                If tbcDetails.Selected.Index = mSpecList.������Ϣ Then
                    If .Col = mSpecColumn.���_��ѡ�� Then
                        tbcDetails.Item(mSpecList.��Ʒ��Ϣ).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.���_��Ʒ����
                    Else
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.��Ʒ��Ϣ Then
                    If .Col = mSpecColumn.���_ע���̱� Then
                        tbcDetails.Item(mSpecList.��װ��λ).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.���_�ۼ۵�λ
                    Else
                        If .Col = mSpecColumn.���_ҩƷ��� Then
                            .Col = mSpecColumn.���_��Ʒ����
                            Exit Sub
                        End If
                        If .Col = mSpecColumn.���_��Դ���� Then
                            .Col = .Col + 3
                        Else
                            .Col = .Col + 1
                        End If
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.��װ��λ Then
                    If .Col = mSpecColumn.���_��ҩ��̬ Then
                        tbcDetails.Item(mSpecList.�۸���Ϣ).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.���_ҩ������
                    Else
                        If .Col = mSpecColumn.���_ҩƷ��� Then
                            .Col = mSpecColumn.���_�ۼ۵�λ
                            Exit Sub
                        ElseIf .Col = mSpecColumn.���_סԺϵ�� Then
                            .Col = mSpecColumn.���_ҩ�ⵥλ
                            Exit Sub
                        End If
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.�۸���Ϣ Then
                    If .Col = mSpecColumn.���_��ǰ�ۼ� Then
                        tbcDetails.Item(mSpecList.ҩ������).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.���_������Ŀ
                    Else
                        If .Col = mSpecColumn.���_ҩƷ��� Then
                            .Col = mSpecColumn.���_ҩ������
                            Exit Sub
                        End If
                        If .Col = mSpecColumn.���_ָ���ۼ� Then
                            .Col = mSpecColumn.���_�ӳ���
                            Exit Sub
                        End If
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.ҩ������ Then
                    If .Col = mSpecColumn.���_ҽ������ Then
                        tbcDetails.Item(mSpecList.��������).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.���_ҩ�����
                    Else
                        If .Col = mSpecColumn.���_ҩƷ��� Then
                            .Col = mSpecColumn.���_������Ŀ
                            Exit Sub
                        End If
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.�������� Then
                    If .Col = mSpecColumn.���_������ Then
                        tbcDetails.Item(mSpecList.�ٴ�Ӧ��).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.���_��ʶ˵��
                    Else
                        If .Col = mSpecColumn.���_ҩƷ��� Then
                            .Col = mSpecColumn.���_ҩ�����
                            Exit Sub
                        End If
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.�ٴ�Ӧ�� Then
                    If .Col <> mSpecColumn.���_����ҩ�� Then
                        If .Col = mSpecColumn.���_ҩƷ��� Then
                            .Col = mSpecColumn.���_��ʶ˵��
                            Exit Sub
                        End If
                        .Col = .Col + 1
                    Else
                        If .Row <> .Rows - 1 Then
                            tbcDetails.Item(mSpecList.������Ϣ).Selected = True
                            .SetFocus
                            .Row = .Row + 1
                            .Col = mSpecColumn.���_ҩƷ���
                        End If
                    End If
                End If
            Else    '����ҩ���г�ҩ
                If tbcDetails.Selected.Index = mSpecList.������Ϣ Then
                    If .Col = mSpecColumn.���_���� Then
                        tbcDetails.Item(mSpecList.��Ʒ��Ϣ).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.���_��Ʒ����
                    Else
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.��Ʒ��Ϣ Then
                    If .Col = mSpecColumn.���_�ǳ���ҩ Then
                        tbcDetails.Item(mSpecList.��װ��λ).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.���_�ۼ۵�λ
                    Else
                        If .Col = mSpecColumn.���_ҩƷ��� Then
                            .Col = mSpecColumn.���_��Ʒ����
                            Exit Sub
                        End If
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.��װ��λ Then
                    If .Col = mSpecColumn.���_���췧ֵ Then
                        tbcDetails.Item(mSpecList.�۸���Ϣ).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.���_ҩ������
                    Else
                        If .Col = mSpecColumn.���_ҩƷ��� Then
                            .Col = mSpecColumn.���_�ۼ۵�λ
                            Exit Sub
                        End If
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.�۸���Ϣ Then
                    If .Col = mSpecColumn.���_��ǰ�ۼ� Then
                        tbcDetails.Item(mSpecList.ҩ������).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.���_������Ŀ
                    Else
                        If .Col = mSpecColumn.���_ҩƷ��� Then
                            .Col = mSpecColumn.���_ҩ������
                            Exit Sub
                        End If
                        If .Col = mSpecColumn.���_ָ���ۼ� Then
                            .Col = mSpecColumn.���_�ӳ���
                            Exit Sub
                        End If
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.ҩ������ Then
                    If .Col = mSpecColumn.���_ҽ������ Then
                        tbcDetails.Item(mSpecList.��������).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.���_ҩ�����
                    Else
                        If .Col = mSpecColumn.���_ҩƷ��� Then
                            .Col = mSpecColumn.���_������Ŀ
                            Exit Sub
                        End If
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.�������� Then
                    If .Col = mSpecColumn.���_������ Then
                        tbcDetails.Item(mSpecList.�ٴ�Ӧ��).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.���_��ʶ˵��
                    Else
                        If .Col = mSpecColumn.���_ҩƷ��� Then
                            .Col = mSpecColumn.���_ҩ�����
                            Exit Sub
                        End If
                        .Col = .Col + 1
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.�ٴ�Ӧ�� Then
                    If .Col = mSpecColumn.���_ҩƷ��� Then
                        .Col = mSpecColumn.���_��ʶ˵��
                        Exit Sub
                    End If
                    If .Col <> mSpecColumn.���_����ҩ�� Then
                        .Col = .Col + 1
                    Else
                        tbcDetails.Item(mSpecList.��ҩ����).Selected = True
                        .SetFocus
                        .Col = mSpecColumn.���_�洢�¶�
                    End If
                ElseIf tbcDetails.Selected.Index = mSpecList.��ҩ���� Then
                    If .Col = mSpecColumn.���_ҩƷ��� Then
                        .Col = mSpecColumn.���_�洢�¶�
                        Exit Sub
                    End If
                    If .Col <> mSpecColumn.���_������� Then
                        .Col = .Col + 1
                    Else
                        If .Row <> .Rows - 1 Then
                            tbcDetails.Item(mSpecList.������Ϣ).Selected = True
                            .SetFocus
                            .Row = .Row + 1
                            .Col = mSpecColumn.���_ҩƷ���
                        End If
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub vsfDetails_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfDetails
        If mint״̬ = 1 Then 'Ʒ��
            If .Col = mVaricolumn.Ʒ��_�ο���Ŀ Then
                If .TextMatrix(Row, mVaricolumn.Ʒ��_�ο���ĿID) = "" Then
                    .TextMatrix(Row, mVaricolumn.Ʒ��_�ο���Ŀ) = ""
                    .EditText = ""
                End If
            End If
        End If
    End With
End Sub
