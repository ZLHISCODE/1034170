VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm������ҩ�б� 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame fraColSel 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4200
      TabIndex        =   0
      Top             =   300
      Width           =   195
      Begin VB.Image imgColSel 
         Height          =   195
         Left            =   0
         Picture         =   "frm������ҩ�б�.frx":0000
         ToolTipText     =   "ѡ����Ҫ��ʾ����(ALT+C)"
         Top             =   0
         Width           =   195
      End
   End
   Begin MSComctlLib.ImageList imgCheck 
      Left            =   4680
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":054E
            Key             =   ""
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":0AE8
            Key             =   ""
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":1082
            Key             =   ""
            Object.Tag             =   "3"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfColSel 
      Height          =   1095
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1470
      _cx             =   2593
      _cy             =   1931
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
      BackColorFixed  =   8421504
      ForeColorFixed  =   16777215
      BackColorSel    =   14737632
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frm������ҩ�б�.frx":11DC
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
      Editable        =   2
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
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   960
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   1800
      _cx             =   3175
      _cy             =   1693
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
      BackColorSel    =   16769992
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   10329501
      GridColorFixed  =   10329501
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm������ҩ�б�.frx":122A
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   5280
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   42
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":129F
            Key             =   "��ӡ11"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":1639
            Key             =   "��ǰ"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":7E9B
            Key             =   "ָʾ��"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":E6FD
            Key             =   "����"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":EC97
            Key             =   "����"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":F031
            Key             =   "��־"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":F3CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":F765
            Key             =   "ͼ��"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":FAFF
            Key             =   "ѡ��"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":10511
            Key             =   "Person"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":16D73
            Key             =   "δ��"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":1D5D5
            Key             =   "�ڼ�"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":23E37
            Key             =   "�Ѽ�"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":2A699
            Key             =   "Family"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":30EFB
            Key             =   "����"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":31295
            Key             =   "����_ѡ��"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":3162F
            Key             =   "�ײ�"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":37E91
            Key             =   "����"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":3E6F3
            Key             =   "��Ƭ"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":44F55
            Key             =   "����"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":4B7B7
            Key             =   "ָ��"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":52019
            Key             =   "���"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":5887B
            Key             =   "������ʽ"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":5F0DD
            Key             =   "�����ļ�"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":6593F
            Key             =   "����"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":6C1A1
            Key             =   "�շ�"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":6CBB3
            Key             =   "���"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":73415
            Key             =   "����"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":79C77
            Key             =   "ȷ��"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":804D9
            Key             =   "��ʼ"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":86D3B
            Key             =   "����"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":8D59D
            Key             =   "����"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":8D937
            Key             =   "ȫ��"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":8DCD1
            Key             =   "�����ܼ�"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":8E06B
            Key             =   "ȫ���ܼ�"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":8E405
            Key             =   "�ܼ�"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":8E79F
            Key             =   "��ӡ"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":8F1B1
            Key             =   "�Ѿ���ӡ"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":8FBC3
            Key             =   "����"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":9015D
            Key             =   "δȡҩ"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":906F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm������ҩ�б�.frx":96F59
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm������ҩ�б�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOutPut As Boolean

'�б���ʾ����
Private Type Type_ShowListCondition
    int�б����� As Integer                          '0-��ҩȷ��,1-����ҩ;2-����ҩ;3-����ҩ;4-��ҩ
    bln����ģʽ As Boolean
    blnȡ����ҩ As Boolean
    bln�Ƿ���� As Boolean
    bln�Ƿ�ǩ��ȷ�� As Boolean
    blnȡҩȷ�� As Boolean                          'ȡҩȷ��Ȩ��
    bln������� As Boolean
End Type
Private mcondition As Type_ShowListCondition

Private mstrUnallowSetColHide As String             '�������������ص���
Private mstrUnallowShow As String                   '��������ʾ����

Private mrsList As ADODB.Recordset
Private mIntOldRow As Integer

Private mintLocate As Integer
Private mintFindType As Integer
Private mstrFind As String
Private mblnSortByName As Boolean                   '�ж��Ƿ���������
Public mstrLastName As String                       '�ϴη�ҩ�Ĳ�������
Private mstrLastNo As String                        '�ϴ�ѡ���NO
Private mblnFreshList As Boolean
Private mblnNoRefreshDetail As Boolean
Private mblnFindOver As Boolean

Private mintMoneyDigit As Integer           '���С��λ��

Private Type FindProcess
    FindType As Integer
    FindContent As String
    StartRow As Integer
End Type
Private mFindProcess As FindProcess

'�������ͣ���ͨ�����ơ������������һ������
Private Enum ��������
    ��ͨ = 0
    ���� = 1
    ���� = 2
    ���� = 3
    ��һ = 4
    ���� = 5
End Enum

'�û�����Ĵ�����ɫ����ע���ȡ���ַ�������;�ָ�
Private mstrUserRecipeColor As String

Private mint�����ʾ As Integer     '�����ʾ��ʽ��0-��ʾӦ�ս��,1-��ʾʵ�ս��,2-��ʾӦ�պ�ʵ�ս��
Private mblnȡҩȷ�� As Boolean       '�Ƿ����ò���ʵ��ȡҩȷ��ģʽ��0-�����ã�1-����
Private mintShowBill��ҩ As Integer '0-��ʾ������ҩ��,1-ֻ��ʾδ��ӡ�Ĵ���ҩ����,2-ֻ��ʾ�Ѵ�ӡ�Ĵ���ҩ����


'�б�����
Private Enum mListType
    ��ҩȷ�� = 0
    ����ҩ = 1
    ����ҩ = 2
    ����ҩ = 3
    ��ʱδ�� = 4
    ��ҩ = 5
End Enum

Private Enum mFindType
    ���ݺ� = 1
    ����� = 2
    ���� = 3
    ���֤ = 4
    IC�� = 5
    ҽ���� = 6
    סԺ�� = 7
End Enum

Private Const mconIntCol���� = 35
Private mIntCol��ǰ�� As Integer
Private mintcolѡ�� As Integer
Private mIntCol��� As Integer
Private mIntCol���� As Integer
Private mIntCol��ɫ As Integer
Private mIntCol�������� As Integer
Private mIntCol��־ As Integer
Private mIntCol���� As Integer
Private mIntCol���� As Integer
Private mIntCol�շ� As Integer
Private mIntCol��ҩ�� As Integer
Private mIntColNO As Integer
Private mIntCol���� As Integer
Private mIntCol��� As Integer
Private mIntCol���� As Integer
Private mIntColǩ������ As Integer
Private mIntCol�ɲ��� As Integer
Private mIntCol˵�� As Integer
Private mIntCol���￨�� As Integer
Private mIntCol����� As Integer
Private mIntCol���֤ As Integer
Private mIntColIC�� As Integer
Private mIntCol����ID As Integer
Private mIntColҽ���� As Integer
Private mIntColסԺ�� As Integer
Private mIntColʵ�ս��  As Integer
Private mIntCol�����־  As Integer
Private mIntCol��¼����  As Integer
Private mIntCol�շ����  As Integer
Private mIntColƴ������ As Integer
Private mIntCol��ʼ��� As Integer
Private mIntCol�Ŷ�״̬ As Integer
Private mIntCol��ҩ���� As Integer
Private mIntColδȡҩ As Integer
Private mIntCol����� As Integer


Public Function GetPrintObject(ByVal blnOutPut As Boolean) As Object
    mblnOutPut = blnOutPut
    If vsfList.rows = 1 Then
        Set GetPrintObject = Nothing
    Else
        Set GetPrintObject = vsfList
    End If
End Function

Public Sub SetFontSize(ByVal intFont As Integer)
    With vsfList
        .Font.Size = intFont
        Me.Font.Size = .Font.Size
        .Cell(flexcpFontSize, 0, 0, .rows - 1, .Cols - 1) = .Font.Size
        
        .RowHeightMin = TextHeight("��") + 100
        .RowHeightMax = TextHeight("��") + 100
        .Refresh
    End With
End Sub
Public Function FindSpecialRow(ByVal strFindType As String, ByVal strFindContent As String, Optional strNos As String, Optional ByRef objSquareCard As Object) As Boolean
    '��������п�����strFindContent��ʽΪ����ID|����
    'strNOs���ص�ǰ�ҵ����еĲ������б��е����д����ţ�����,NO|����,NO
    Dim intCol As Integer
    Dim intFindRow As Integer
    Dim strNo As String
    Dim lng����ID As Long
    Dim strName As String
    Dim IntCount As Integer
    
    mblnFindOver = True
    
    With mFindProcess
'        .FindType = intFindType
        .FindContent = UCase(strFindContent)
        .StartRow = 1
    End With
    
    With vsfList
        Select Case strFindType
            Case "����"
                If Mid(mFindProcess.FindContent, 1, 1) = "-" And Val(Mid(mFindProcess.FindContent, 2)) <> 0 Then        '�жϵ�ǰ�����Ƿ񰴡�-����ID����ʽ����
                    'ͨ����-����ID����ʽ����
                    intCol = mIntCol����ID                                                  '���¶��������
                    mFindProcess.FindContent = Val(Mid(mFindProcess.FindContent, 2))        '��д��������
                Else
                    '��������������
                    intCol = mIntCol����
                    
                    If zlCommFun.IsCharAlpha(mFindProcess.FindContent) Then
                        'ȫ��ĸʱƥ�����
                        If zlDatabase.GetPara("���뷽ʽ") = 0 Then
                            intCol = mIntColƴ������
                        Else
                            intCol = mIntCol��ʼ���
                        End If
                    End If
                End If
            Case "���ݺ�"
                intCol = mIntColNO
            Case "�����"
                intCol = mIntCol�����
            Case "���֤"
                intCol = mIntCol���֤
            Case "IC��"
                intCol = mIntCol����ID
            Case "ҽ����"
                intCol = mIntColҽ����
            Case "סԺ��"
                intCol = mIntColסԺ��
            Case Else
                '����Ϊ���ѿ���������ID����
                intCol = mIntCol����ID
                mFindProcess.FindContent = zlfuncCard_GetPatiID(objSquareCard, Val(Split(strFindContent, "|")(0)), Split(strFindContent, "|")(1))
        End Select
        
        mFindProcess.StartRow = .FindRow(mFindProcess.FindContent, mFindProcess.StartRow, intCol)
        
        If mFindProcess.StartRow > 0 Then
            If .Row <> mFindProcess.StartRow Then
                mblnFindOver = False
            End If
            .Row = mFindProcess.StartRow
            mblnFindOver = True
            .TopRow = .Row
            FindSpecialRow = True
            strNo = .TextMatrix(.Row, mIntColNO)
            lng����ID = Val(.TextMatrix(.Row, mIntCol����ID))
            strName = .TextMatrix(.Row, mIntCol����)
            
            If mFindProcess.StartRow + 1 >= .rows Then
                mFindProcess.StartRow = 1
            Else
                mFindProcess.StartRow = mFindProcess.StartRow + 1
            End If
        Else
            mFindProcess.StartRow = 1
        End If
                
        If strNo <> "" Then
            For IntCount = 1 To .rows - 1
                If lng����ID > 0 Then
                    If Val(.TextMatrix(IntCount, mIntCol����ID)) = lng����ID Then
                        strNos = IIf(strNos = "", "", strNos & "|") & .TextMatrix(IntCount, mIntCol����) & "," & .TextMatrix(IntCount, mIntColNO)
                    End If
                Else
                    If .TextMatrix(IntCount, mIntCol����) = strName Then
                        strNos = IIf(strNos = "", "", strNos & "|") & .TextMatrix(IntCount, mIntCol����) & "," & .TextMatrix(IntCount, mIntColNO)
                    End If
                End If
            Next
        End If
    End With
    
    mblnFindOver = False
End Function

Private Sub FindNextPati(ByVal blnFirst As Boolean)
'    Static intStar As Integer
'    Dim n As Integer
'    Dim strFind As String
'    Dim blnDo As Boolean
'
'    If BlnFirst Then intStar = 1
'
'    If Trim(txtFind.Text) = "" Then Exit Sub
'
'    strFind = Trim(txtFind.Text)
'
'    With Msf�б�
'        If .Rows < 2 Then Exit Sub
'
'        For n = intStar To .Rows - 1
'            Select Case lblFind.Tag
'                Case FindType.���￨
'                    If Trim(.TextMatrix(n, ��������.���￨��)) = strFind Then blnDo = True
'                Case FindType.�����
'                    If Trim(.TextMatrix(n, ��������.�����)) = strFind Then blnDo = True
'                Case FindType.���ݺ�
'                    If Trim(.TextMatrix(n, ��������.NO)) = strFind Then blnDo = True
'                Case FindType.����
'                    If mblnCard = True Then
'                        If Trim(.TextMatrix(n, ��������.���￨��)) = strFind Then blnDo = True
'                    Else
'                        If gbytCode = 1 Then
'                            If Trim(.TextMatrix(n, ��������.����)) Like "*" & strFind & "*" Or mWBX(Trim(.TextMatrix(n, ��������.����)), 1) Like "*" & UCase(strFind) & "*" Then blnDo = True
'                        Else
'                            If Trim(.TextMatrix(n, ��������.����)) Like "*" & strFind & "*" Or mPinYin(Trim(.TextMatrix(n, ��������.����))) Like "*" & UCase(strFind) & "*" Then blnDo = True
'                        End If
'                    End If
'                Case FindType.���֤
'                    If Trim(.TextMatrix(n, ��������.���֤)) = strFind Then blnDo = True
'                Case FindType.IC��
'                    If Trim(.TextMatrix(n, ��������.IC��)) = strFind Then blnDo = True
'            End Select
'
'            If blnDo Then
'                txtFind.Tag = txtFind.Text
'                .Row = n
'                Call Msf�б�_EnterCell
'                .TopRow = n
'                intStar = n + 1
'                If intStar > .Rows - 1 Then intStar = .Rows - 1
'                Exit Sub
'            End If
'        Next
'    End With
'    intStar = 1
'    txtFind.SetFocus
    
End Sub
Public Function GetCurrentRecipe() As String
    'ȡ��ǰ����
    '���أ�0����|1NO|2����|3����ID|4��¼����|5�����־|6��������|7�շ�����|8��������|9��ҩ����|10����|11δȡҩ|12�к�
    
    With vsfList
        If .Row = 0 Then Exit Function
        If Val(.TextMatrix(.Row, mIntCol����)) = 0 Then Exit Function

        GetCurrentRecipe = .TextMatrix(.Row, mIntCol����) & "|" & _
            .TextMatrix(.Row, mIntColNO) & "|" & _
            .TextMatrix(.Row, mIntCol����) & "|" & _
            .TextMatrix(.Row, mIntCol����ID) & "|" & _
            .TextMatrix(.Row, mIntCol��¼����) & "|" & _
            .TextMatrix(.Row, mIntCol�����־) & "|" & _
            .TextMatrix(.Row, mIntCol��������) & "|" & _
            .TextMatrix(.Row, mIntCol�շ����) & "|" & _
            .TextMatrix(.Row, mIntCol����) & "|" & _
            .TextMatrix(.Row, mIntCol��ҩ����) & "|" & _
            .TextMatrix(.Row, mIntCol����) & "|" & _
            .TextMatrix(.Row, mIntColδȡҩ) & "|" & _
            .Row
    End With
End Function

Public Function GetCurrentBatchRecipe() As String
    '��ҩʱ��ȡ��ǰ��ѡ����
    '���أ�����,NO,����ID,���,δ���,��¼����,�����־|����,NO,����ID,ʵ�ս��,δ���,��¼����,�����־
    Dim i As Integer
    Dim strRecipe As String
    
    If mblnFreshList = True Then Exit Function
    
    With vsfList
        If mcondition.bln����ģʽ = False Then
            If .TextMatrix(.Row, mIntColNO) <> "" Then
                strRecipe = .TextMatrix(.Row, mIntCol����) & "," & _
                            .TextMatrix(.Row, mIntColNO) & "," & _
                            .TextMatrix(.Row, mIntCol����ID) & "," & _
                            .TextMatrix(.Row, mIntColʵ�ս��) & "," & _
                            .TextMatrix(.Row, mIntCol�շ�) & "," & _
                            .TextMatrix(.Row, mIntCol��¼����) & "," & _
                            .TextMatrix(.Row, mIntCol�����־) & "," & _
                            .TextMatrix(.Row, mIntCol�շ����)
            End If
        Else
            For i = 1 To .rows - 1
                If .TextMatrix(i, mIntColNO) <> "" And Val(.TextMatrix(i, mIntCol��־)) = 1 Then
                    strRecipe = IIf(strRecipe = "", "", strRecipe & "|") & _
                        .TextMatrix(i, mIntCol����) & "," & _
                        .TextMatrix(i, mIntColNO) & "," & _
                        .TextMatrix(i, mIntCol����ID) & "," & _
                        .TextMatrix(i, mIntColʵ�ս��) & "," & _
                        .TextMatrix(i, mIntCol�շ�) & "," & _
                        .TextMatrix(i, mIntCol��¼����) & "," & _
                        .TextMatrix(i, mIntCol�����־) & "," & _
                        .TextMatrix(.Row, mIntCol�շ����)
                End If
            Next
        End If
        GetCurrentBatchRecipe = strRecipe
    End With
End Function
Sub InitList(ByVal intType As Integer)
    Dim i As Integer
    Dim n As Integer
    Dim str������ As String
    Dim arr������
    Dim bln��������Ч As Boolean
    
    '''��ʼ����˳��
    'Ĭ����˳��
    mIntCol��ǰ�� = 0
    mintcolѡ�� = 1
    mIntCol��� = 2
    mIntCol���� = 3
    mIntCol��ɫ = 4
    mIntCol�������� = 5
    mIntCol���� = 6
    mIntColNO = 7
    mIntCol���� = 8
    mIntCol��� = 9
    mIntColʵ�ս�� = 10
    mIntCol���� = 11
    mIntColǩ������ = 12
    mIntCol�ɲ��� = 13
    mIntCol˵�� = 14
    mIntCol���￨�� = 15
    mIntCol����� = 16
    mIntCol���֤ = 17
    mIntColIC�� = 18
    mIntCol����ID = 19
    mIntColҽ���� = 20
    mIntColסԺ�� = 21
    mIntCol��־ = 22
    mIntCol���� = 23
    mIntCol�շ� = 24
    mIntCol��ҩ�� = 25
    mIntCol�����־ = 26
    mIntCol��¼���� = 27
    mIntCol�շ���� = 28
    mIntColƴ������ = 29
    mIntCol��ʼ��� = 30
    mIntCol�Ŷ�״̬ = 31
    mIntCol��ҩ���� = 32
    mIntColδȡҩ = 33
    mIntCol����� = 34
    
    '�ָ��û��Զ�����˳��
    str������ = LoadListColState
    If str������ <> "" Then
        arr������ = Split(str������, "|")
        If UBound(arr������) + 1 <> mconIntCol���� Then
            str������ = ""
        Else
            For n = 0 To UBound(arr������)
                If Split(arr������(n), ",")(0) = "" Then
                    bln��������Ч = True
                    Exit For
                End If
            Next
            
            If bln��������Ч = True Then
                str������ = ""
            Else
                For n = 0 To UBound(arr������)
                    SetColumnValue Split(arr������(n), ",")(0), n
                Next
            End If
        End If
    End If
     
    '��ʼ��δ��ҩ�嵥
    With vsfList
        .Redraw = flexRDNone
        
        .rows = 1
        .rows = 2
        .Cols = mconIntCol����
        .ExplorerBar = IIf(intType = mListType.����ҩ And mcondition.bln����ģʽ = True, flexExNone, flexExSortShowAndMove)
        
        .Cell(flexcpPicture, 1, mIntCol��ǰ��, 1, mIntCol��ǰ��) = Me.imgList.ListImages(2).Picture
        .Cell(flexcpPictureAlignment, 1, mIntCol��ǰ��, .rows - 1, mIntCol��ǰ��) = flexPicAlignRightCenter
        
        VsfGridColFormat vsfList, mIntCol��ǰ��, "", 250, flexAlignCenterCenter, "��ǰ��"
        
        VsfGridColFormat vsfList, mintcolѡ��, "", IIf(intType = mListType.����ҩ And mcondition.bln����ģʽ = True, 300, 0), flexAlignCenterCenter, "ѡ��"
        VsfGridColFormat vsfList, mIntCol���, "��", IIf((intType = mListType.����ҩ Or intType = mListType.����ҩ) And mcondition.bln������� = True, 300, 0), flexAlignCenterCenter, "���"
        VsfGridColFormat vsfList, mIntCol����, "����", 500, flexAlignCenterCenter, "����"
        VsfGridColFormat vsfList, mIntCol��ɫ, "����", 500, flexAlignCenterCenter, "����"
        VsfGridColFormat vsfList, mIntCol��������, "��������", 0, flexAlignCenterCenter, "��������"
        VsfGridColFormat vsfList, mIntCol��־, "1", 0, flexAlignCenterCenter, "��־"
        VsfGridColFormat vsfList, mIntCol����, "���", 1000, flexAlignLeftCenter, "���"
        VsfGridColFormat vsfList, mIntCol����, "����", 0, flexAlignCenterCenter, "����"
        VsfGridColFormat vsfList, mIntCol�շ�, "�շ�", 0, flexAlignCenterCenter, "�շ�"
        VsfGridColFormat vsfList, mIntCol��ҩ��, "��ҩ��", 0, flexAlignCenterCenter, "��ҩ��"
        
        If mblnȡҩȷ�� = True Or intType = mListType.����ҩ Then
            VsfGridColFormat vsfList, mIntColNO, "NO", 1100, flexAlignRightCenter, "NO"
        Else
            VsfGridColFormat vsfList, mIntColNO, "NO", 800, flexAlignLeftCenter, "NO"
        End If
        
        VsfGridColFormat vsfList, mIntCol����, "����", 800, flexAlignLeftCenter, "����"
        
        VsfGridColFormat vsfList, mIntCol���, "Ӧ�ս��", IIf(mint�����ʾ = 1, 0, 1000), flexAlignRightCenter, "Ӧ�ս��"
        VsfGridColFormat vsfList, mIntColʵ�ս��, "ʵ�ս��", IIf(mint�����ʾ = 0, 0, 1000), flexAlignRightCenter, "ʵ�ս��"
        VsfGridColFormat vsfList, mIntCol����, "����", 1500, flexAlignLeftCenter, "����"
        VsfGridColFormat vsfList, mIntColǩ������, "ǩ������", 1500, flexAlignLeftCenter, "ǩ������"
        VsfGridColFormat vsfList, mIntCol�ɲ���, "�ɲ���", 0, flexAlignCenterCenter, "�ɲ���"
        VsfGridColFormat vsfList, mIntCol˵��, "˵��", 1500, flexAlignLeftCenter, "˵��"
        VsfGridColFormat vsfList, mIntCol���￨��, "���￨��", 1000, flexAlignLeftCenter, "���￨��"
        VsfGridColFormat vsfList, mIntCol�����, "�����", 1000, flexAlignLeftCenter, "�����"
        VsfGridColFormat vsfList, mIntCol���֤, "���֤", 1600, flexAlignLeftCenter, "���֤"
        VsfGridColFormat vsfList, mIntColIC��, "IC��", 1600, flexAlignLeftCenter, "IC��"
        VsfGridColFormat vsfList, mIntCol����ID, "����ID", 0, flexAlignCenterCenter, "����ID"
        VsfGridColFormat vsfList, mIntColҽ����, "ҽ����", 1500, flexAlignLeftCenter, "ҽ����"
        VsfGridColFormat vsfList, mIntColסԺ��, "סԺ��", 1000, flexAlignLeftCenter, "סԺ��"
        
        VsfGridColFormat vsfList, mIntCol�����־, "�����־", 0, flexAlignCenterCenter, "�����־"
        VsfGridColFormat vsfList, mIntCol��¼����, "��¼����", 0, flexAlignCenterCenter, "��¼����"
        VsfGridColFormat vsfList, mIntCol�շ����, "�շ�����", 0, flexAlignCenterCenter, "�շ�����"
        VsfGridColFormat vsfList, mIntColƴ������, "ƴ������", 0, flexAlignCenterCenter, "ƴ������"
        VsfGridColFormat vsfList, mIntCol��ʼ���, "��ʼ���", 0, flexAlignCenterCenter, "��ʼ���"
        VsfGridColFormat vsfList, mIntCol�Ŷ�״̬, "�Ŷ�״̬", 0, flexAlignCenterCenter, "�Ŷ�״̬"
        VsfGridColFormat vsfList, mIntCol��ҩ����, "��ҩ����", 0, flexAlignCenterCenter, "��ҩ����"
        VsfGridColFormat vsfList, mIntColδȡҩ, "δȡҩ", 0, flexAlignCenterCenter, "δȡҩ"
        VsfGridColFormat vsfList, mIntCol�����, "�����", 0, flexAlignCenterCenter, "�����"
        
        mstrUnallowSetColHide = "NO"
        mstrUnallowShow = "��ǰ��;��������;��־;����;�շ�;��ҩ��;�ɲ���;����ID;δ���;�����־;��¼����;�շ�����;ƴ������;��ʼ���;�Ŷ�״̬;��ҩ����;δȡҩ"
        If mint�����ʾ = 0 Then mstrUnallowShow = mstrUnallowShow & ";ʵ�ս��"
        If mint�����ʾ = 1 Then mstrUnallowShow = mstrUnallowShow & ";Ӧ�ս��"
        If mcondition.bln����ģʽ = False Then mstrUnallowShow = mstrUnallowShow & ";" & "ѡ��"
        If mcondition.int�б����� = mListType.��ҩ Or Not mcondition.bln�Ƿ�ǩ��ȷ�� Then mstrUnallowShow = mstrUnallowShow & ";" & "ǩ������"
        If mcondition.int�б����� <> mListType.����ҩ Or Not mcondition.bln�Ƿ���� Then mstrUnallowShow = mstrUnallowShow & ";" & "����"
        
        '�ָ��Զ����п���������������ʾ���У�
        If str������ <> "" Then
            arr������ = Split(str������, "|")
            For n = 0 To UBound(arr������)
                If IsInString(mstrUnallowShow, Split(arr������(n), ",")(0), ";") = False Then
                    For i = 0 To vsfList.Cols - 1
                        If Split(arr������(n), ",")(0) = vsfList.ColKey(i) Then
                            vsfList.ColWidth(i) = Val(Split(arr������(n), ",")(1))
                        End If
                    Next
                End If
            Next
        End If
        
        If .ColWidth(mIntCol��ɫ) = 0 Then .ColWidth(mIntCol��ɫ) = 500
        
        If mcondition.int�б����� = mListType.����ҩ And mcondition.bln�Ƿ���� Then
            .ColHidden(mIntCol����) = False
        Else
            .ColHidden(mIntCol����) = True
        End If
        
        If mcondition.int�б����� <> mListType.��ҩ And mcondition.bln�Ƿ�ǩ��ȷ�� Then
            .ColHidden(mIntColǩ������) = False
        Else
            .ColHidden(mIntColǩ������) = True
        End If
                
        .RowSel = 1
        
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub SaveListColState(ByVal int���� As Integer)
    Dim str������ As String
    Dim i As Integer
    Dim strType As String
    
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then Exit Sub
    
    Select Case int����
        Case mListType.��ҩȷ��
            strType = "��ҩȷ��"
        Case mListType.����ҩ
            strType = "����ҩ"
        Case mListType.����ҩ
            strType = "����ҩ"
        Case mListType.����ҩ
            strType = "����ҩ"
        Case mListType.��ʱδ��
            strType = "��ʱδ��"
        Case mListType.��ҩ
            strType = "��ҩ"
    End Select
    
    With vsfList
        For i = 0 To .Cols - 1
            str������ = IIf(str������ = "", "", str������ & "|") & vsfList.ColKey(i) & "," & .ColWidth(i)
        Next
    End With
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(vsfList), strType, str������)
End Sub

Private Function LoadListColState() As String
    Dim str������ As String
    Dim i As Integer
    Dim strType As String
    
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then Exit Function
    
    Select Case mcondition.int�б�����
        Case mListType.��ҩȷ��
            strType = "��ҩȷ��"
        Case mListType.����ҩ
            strType = "����ҩ"
        Case mListType.����ҩ
            strType = "����ҩ"
        Case mListType.����ҩ
            strType = "����ҩ"
        Case mListType.��ʱδ��
            strType = "��ʱδ��"
        Case mListType.��ҩ
            strType = "��ҩ"
    End Select
    
    LoadListColState = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\��������\" & App.ProductName & "\" & Me.Name & "\" & TypeName(vsfList), strType, "")
End Function

Private Sub SetMainComandBars(ByVal intListType As Integer, ByVal lngRow As Long)
    '���ݵ�ǰ��¼�嵥���ͼ���ǰ��¼������������Ĳ˵�״̬
    Dim cbrControl As CommandBarControl
    Dim cbrMenu As CommandBarControl
    Dim bln����ȡ�� As Boolean
    Dim int�����־ As Integer
    Dim int��¼���� As Integer
    Dim Int���� As Integer
    Dim strNo As String
    Dim blnAddSign As Boolean
    Dim blnVeirfySign As Boolean
    Dim intδȡҩ As Integer
    Dim dateNow As Date
    Dim lngTime As Double
    Dim int��˽�� As Integer
     
    If lngRow = 0 Then Exit Sub
    
    int�����־ = Val(vsfList.TextMatrix(lngRow, mIntCol�����־))
    int��¼���� = Val(vsfList.TextMatrix(lngRow, mIntCol��¼����))
    Int���� = Val(vsfList.TextMatrix(lngRow, mIntCol����))
    strNo = vsfList.TextMatrix(lngRow, mIntColNO)
    intδȡҩ = Val(vsfList.TextMatrix(lngRow, mIntColδȡҩ))
    int��˽�� = Val(vsfList.TextMatrix(lngRow, mIntCol�����))
    
    
    '��ҩ״̬ʱȡ��ǰ���ݵļ�¼���ʺ������־�����ж��Ƿ����ȡ����ҩ
    If intListType = mListType.��ҩ And mcondition.blnȡ����ҩ = True Then
        '����ȡ����ҩģʽ�Ƿ����
        If (((int�����־ = 1 Or int�����־ = 4) And gtype_UserSysParms.P15_�����շ��뷢ҩ���� = 1) Or _
            (int�����־ = 2 And gtype_UserSysParms.P16_סԺ�����뷢ҩ���� = 1)) And CheckIsSended(Int����, strNo) = False Then
            bln����ȡ�� = True
        End If
        
        Set cbrMenu = frmҩƷ������ҩNew.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Cancel, , True)
        Set cbrControl = frmҩƷ������ҩNew.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Cancel, , True)
        
        If Not cbrMenu Is Nothing Then cbrMenu.Enabled = bln����ȡ��
        If Not cbrControl Is Nothing Then cbrControl.Enabled = bln����ȡ��
    End If
    
    
    'δ��˴������ܽк�
    If mcondition.bln�Ƿ���� And intListType = mListType.����ҩ Then
        Set cbrMenu = frmҩƷ������ҩNew.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Call, , True)
        Set cbrControl = frmҩƷ������ҩNew.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Call, , True)

        If Not cbrMenu Is Nothing Then
            If mcondition.bln������� Then
                cbrMenu.Enabled = (int��˽�� = 1)
            Else
                cbrMenu.Enabled = True
            End If
        End If
        If Not cbrControl Is Nothing Then
            If mcondition.bln������� Then
                cbrControl.Enabled = (int��˽�� = 1)
            Else
                cbrControl.Enabled = True
            End If
        End If
    End If
    
    
    '��ҩ״̬ʱ���������Ѿ�����3�죬�򲻿��Խ��к��в���
    If intListType = mListType.����ҩ Then
        dateNow = zlDatabase.Currentdate
        lngTime = dateNow - CDate(vsfList.TextMatrix(lngRow, mIntCol����))
        Set cbrMenu = frmҩƷ������ҩNew.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Call, , True)
        Set cbrControl = frmҩƷ������ҩNew.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Call, , True)
        If lngTime > 3 Then
            If Not cbrMenu Is Nothing Then cbrMenu.Enabled = False
            If Not cbrControl Is Nothing Then cbrControl.Enabled = False
        Else
            If Not cbrMenu Is Nothing Then cbrMenu.Enabled = True
            If Not cbrControl Is Nothing Then cbrControl.Enabled = True
        End If
    End If
    
    '��ҩʱ��ȡ����ҩ������֤����ǩ��
    If intListType = mListType.��ҩ Or intListType = mListType.����ҩ Then
        If gblnESign������ҩ = True Then
            blnAddSign = RecipeSendWork_JudgeSign(Int����, strNo, Val(vsfList.TextMatrix(vsfList.Row, mIntCol�ɲ���)), 0, CDate(vsfList.TextMatrix(vsfList.Row, mIntCol����)))
            
            Set cbrMenu = frmҩƷ������ҩNew.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_VerifySign, , True)
            Set cbrControl = frmҩƷ������ҩNew.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_VerifySign, , True)

            If Not cbrMenu Is Nothing Then cbrMenu.Enabled = blnAddSign
            If Not cbrControl Is Nothing Then cbrControl.Enabled = blnAddSign
        End If
    End If
    
    '����ʵ��ȡҩȷ��
    If intListType = mListType.��ҩ Then
        Set cbrMenu = frmҩƷ������ҩNew.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_TakeDrug, , True)
        Set cbrControl = frmҩƷ������ҩNew.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_TakeDrug, , True)

        If Not cbrMenu Is Nothing Then
            If mblnȡҩȷ�� = True And mcondition.blnȡҩȷ�� = True Then
                cbrMenu.Enabled = (intδȡҩ = 1)
            Else
                cbrMenu.Visible = False
            End If
        End If
        If Not cbrControl Is Nothing Then
            If mblnȡҩȷ�� = True And mcondition.blnȡҩȷ�� = True Then
                cbrControl.Enabled = (intδȡҩ = 1)
            Else
                cbrControl.Visible = False
            End If
        End If
    End If
    
    
End Sub

Private Function CheckIsSended(ByVal Int���� As Integer, ByVal strNo As String) As Boolean
    '����Ƿ�����ҩ
    Dim rstemp As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select 1 From ҩƷ�շ���¼ Where ���� = [1] And NO = [2] And ��¼״̬ <> 1 And ������� Is Not Null And Rownum =1 "
    Set rstemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ�����ȡ����ҩ", Int����, strNo)
    
    CheckIsSended = (rstemp.RecordCount > 0)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Sub SetParams()
    mstrUserRecipeColor = zlDatabase.GetPara("������ɫ", glngSys, 1341)
    If mstrUserRecipeColor = "" Then mstrUserRecipeColor = GetDefaultRecipeColor
    
    mcondition.blnȡ����ҩ = IsInString(gstrprivs, "ȡ����ҩ", ";")
    mcondition.blnȡҩȷ�� = IsInString(gstrprivs, "ȡҩȷ��", ";")
       
    mint�����ʾ = Val(zlDatabase.GetPara("�����ʾ��ʽ", glngSys, 1341, 0))
    mblnȡҩȷ�� = (Val(zlDatabase.GetPara("���ò���ʵ��ȡҩȷ��ģʽ", glngSys, 1341, 0)) = 1)
    mintShowBill��ҩ = Val(zlDatabase.GetPara("����ҩ���ݴ�ӡ��ʾ��ʽ", glngSys, 1341, 0))
End Sub

Public Sub SetPrintFlag(ByVal lngRow As Long)
    '����������ô�ӡ��ҩ�������ô���ҩ�б��еĴ�ӡͼ��
    If mcondition.int�б����� <> mListType.����ҩ Then Exit Sub
    If lngRow <= 0 Or lngRow > vsfList.rows - 1 Then Exit Sub
    
    vsfList.Redraw = flexRDNone
    
    If mintShowBill��ҩ = 1 Then
        vsfList.RemoveItem lngRow
        
        If lngRow <= vsfList.rows - 1 Then
            vsfList.Row = lngRow
        Else
            vsfList.Row = vsfList.rows - 1
        End If
        
        Call vsfList_EnterCell
    Else
        vsfList.Cell(flexcpPicture, lngRow, mIntColNO) = Me.imgList.ListImages("��ӡ").Picture
        vsfList.Cell(flexcpPictureAlignment, lngRow, mIntColNO) = flexPicAlignLeftCenter
    End If
    
    vsfList.Redraw = flexRDDirect
End Sub

Public Sub ShowList(ByVal intType As Integer, ByVal bln����ģʽ As Boolean, ByVal bln�Ƿ���� As Boolean, ByVal bln�Ƿ�ǩ��ȷ�� As Boolean, ByVal bln������� As Boolean, Optional ByVal intFindType As Integer = 0, Optional ByVal StrFind As String = "")
    vsfColSel.Visible = False
    
    If mcondition.int�б����� <> intType Then
        mintLocate = 1
        mcondition.int�б����� = intType
        mcondition.bln����ģʽ = bln����ģʽ
    End If
    
    mcondition.bln�Ƿ���� = bln�Ƿ����
    mcondition.bln�Ƿ�ǩ��ȷ�� = bln�Ƿ�ǩ��ȷ��
    mintFindType = intFindType
    mstrFind = StrFind
    mcondition.bln������� = bln�������
    
    Call InitList(mcondition.int�б�����)
    
    Call InitColSelList(mcondition.int�б�����)
End Sub
Private Sub Form_Load()
    'ȡ���λ��
    mintMoneyDigit = GetDigit(0, 1, 4)
    
    Call SetParams
End Sub

Private Sub Form_Resize()
    vsfList.Move 0, 0, Me.Width, Me.Height
    
    fraColSel.Left = vsfList.ColWidth(0) - fraColSel.Width - 50
    fraColSel.Top = (vsfList.RowHeight(0) - fraColSel.Height) / 2 + 30
    fraColSel.ZOrder
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrLastNo = ""
    
    Call SaveListColState(mcondition.int�б�����)
    
    'û�����ø��Ի�����ʱɾ���û���������
    On Error Resume Next
    If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then
        DeleteSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ������ҩ", "�����嵥����" & mListType.��ҩȷ��
        DeleteSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ������ҩ", "�����嵥����" & mListType.����ҩ
        DeleteSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ������ҩ", "�����嵥����" & mListType.����ҩ
        DeleteSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ������ҩ", "�����嵥����" & mListType.����ҩ
        DeleteSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ������ҩ", "�����嵥����" & mListType.��ʱδ��
        DeleteSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ������ҩ", "�����嵥����" & mListType.��ҩ
    End If
End Sub
Private Sub imgColSel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    
    If Button = 1 Then '��ѡ����
        '���ݵ�ǰ״ֱ̬��ȷ����ѡ״̬
        With vsfColSel
            If .Visible Then
                .Visible = False
                vsfList.SetFocus
            Else
                For i = .FixedRows To .rows - 1
                    If vsfList.ColHidden(.RowData(i)) Or vsfList.ColWidth(.RowData(i)) = 0 Then
                        .TextMatrix(i, 0) = 0
                    Else
                        .TextMatrix(i, 0) = 1
                    End If
                Next
                
                .Height = .RowHeightMin * .rows + 150
                .Top = fraColSel.Top + fraColSel.Height
                If .Top + .Height > Me.ScaleHeight - vsfList.Top Then
                    .Height = Me.ScaleHeight - .Top - vsfList.Top
                    .Width = 1750
                Else
                    .Width = 1470
                End If
                
                .Left = fraColSel.Left
                .ZOrder
                .Visible = True
                .SetFocus
            End If
        End With
    End If
End Sub


Private Sub vsfColSel_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngCol As Long
    
    If Col = 0 Then
        lngCol = vsfColSel.RowData(Row)
        If Val(vsfColSel.TextMatrix(Row, 0)) <> 0 Then
            vsfList.ColWidth(lngCol) = vsfList.ColData(lngCol)
            vsfList.ColHidden(lngCol) = False
        Else
            vsfList.ColWidth(lngCol) = 0
            vsfList.ColHidden(lngCol) = True
        End If
    End If
    
    Call SaveListColState(mcondition.int�б�����)
End Sub

Private Sub vsfColSel_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsfColSel
        If NewRow >= .FixedRows - 1 And NewCol >= .FixedCols - 1 Then
            .ForeColorSel = .Cell(flexcpForeColor, NewRow, 1)
            .Col = 0
        End If
    End With
End Sub


Private Sub vsfColSel_LostFocus()
    vsfColSel.Visible = False
End Sub

Private Sub vsfColSel_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Or vsfColSel.Cell(flexcpForeColor, Row, 1) = vsfColSel.BackColorFixed Then Cancel = True
End Sub


Private Sub vsfList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    Dim i As Integer
    
    '������ѡ���б�
    Call InitColSelList(mcondition.int�б�����)
    
    '������˳���
    For i = 0 To vsfList.Cols - 1
        Call SetColumnValue(vsfList.TextMatrix(0, i), i)
    Next
    
    '�����б��״̬
    Call SaveListColState(mcondition.int�б�����)
End Sub

Private Sub vsfList_AfterSort(ByVal Col As Long, Order As Integer)
    If Col = mIntCol���� Then
        mblnSortByName = True
    Else
        mblnSortByName = False
    End If
    
    Call vsfList_EnterCell

    '���洦���嵥���û��������
    '�������
    '����б�����
    'ֵ=�к�|��/����
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ������ҩ", "�����嵥����" & mcondition.int�б�����, Col & "|" & Order)
End Sub


Private Sub vsfList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    '�����б��״̬
    Call SaveListColState(mcondition.int�б�����)
End Sub
Private Sub vsfList_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    '���ò����ƶ�����
    Select Case mcondition.int�б�����
        Case mListType.����ҩ, mListType.����ҩ, mListType.��ʱδ��, mListType.��ҩȷ��
            If Col = mIntCol��ɫ Then
                Position = mIntCol��ɫ
            End If

            If Col <> mIntCol��ɫ And Position = mIntCol��ɫ Then
                Position = Col
            End If
        Case mListType.����ҩ
            If Col = mIntCol��ɫ Then
                Position = mIntCol��ɫ
            End If
            
            If Col = mintcolѡ�� Then
                Position = mintcolѡ��
            End If
            
            If Col = mIntCol���� Then
                Position = mIntCol����
            End If
            
            If (Col <> mIntCol��ɫ And Position = mIntCol��ɫ) Or (Col <> mintcolѡ�� And Position = mintcolѡ��) Or (Col <> mIntCol���� And Position = mIntCol����) Then
                Position = Col
            End If
    End Select
End Sub

Private Sub vsfList_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '���ò��ܵ����п����
    Select Case mcondition.int�б�����
        Case mListType.����ҩ, mListType.����ҩ, mListType.��ʱδ��
            If Col = mIntCol��ǰ�� Or Col = mIntCol��ɫ Then Cancel = True
        Case mListType.����ҩ
            If Col = mIntCol��ǰ�� Or Col = mIntCol��ɫ Or Col = mintcolѡ�� Or Col = mIntCol���� Then Cancel = True
        Case Else
            If Col = 0 Then Cancel = True
    End Select
End Sub

Private Sub InitColSelList(ByVal intListType As Integer)
    Dim i As Integer
    
    With vsfColSel
        .Tag = intListType
        
        .rows = .FixedRows
        For i = 1 To vsfList.Cols - 1
            '���ڲ�������ʾ�б���в��ܼ�����ѡ���б�
            If IsInString(mstrUnallowShow, vsfList.ColKey(i), ";") = False Then
                .rows = .rows + 1
                .TextMatrix(.rows - 1, 1) = vsfList.TextMatrix(0, i)
                .RowData(.rows - 1) = i
                
                '�п�Ϊ�ջ������ص�������Ϊ����ѡ
                If Not (vsfList.ColWidth(i) = 0 Or vsfList.ColHidden(i)) Then
                    .TextMatrix(.rows - 1, 0) = 0
                End If
                
                'ָ����������Ϊ������������
                If IsInString(mstrUnallowSetColHide, vsfList.ColKey(i), ";") = True Then
                    .Cell(flexcpForeColor, .rows - 1, 1) = .BackColorFixed
                End If
            End If
        Next
    End With
End Sub
Public Sub RefreshList(ByVal intType As Integer, ByVal rsData As ADODB.Recordset, Optional ByVal strNo As String, Optional ByVal blnNoRefreshDetail As Boolean)
    '���������б����ݼ��ĸ���
    Set mrsList = rsData
    Dim intRow As Integer
    Dim lngColor As Long
    Dim strSort As String
    Dim lngFindRow As Long
    Dim StrFind As String
    Dim intFindCol As Integer
    Dim lngTime As Long
    Dim dateNow As Date
    
    mblnFreshList = True
    
    mblnNoRefreshDetail = blnNoRefreshDetail
    
    mcondition.bln����ģʽ = (Val(GetSetting("ZLSOFT", "����ģ��\����\" & App.ProductName & "\" & "ҩƷ������ҩ", "���涨λ", 0)) = 1)
    
    Call InitList(intType)
    
    With vsfList
        .Redraw = flexRDNone
        .MergeCells = flexMergeNever
        .rows = 1
        
        If mrsList.EOF Then
            .rows = 2
            .Cell(flexcpText, 1, mIntCol����, 1, .Cols - 1) = "û���ҵ����������ļ�¼......"
            .MergeCells = flexMergeRestrictRows
            .MergeRow(1) = True
            frmҩƷ������ҩNew.ClearForm_Detail
            frmҩƷ������ҩNew.ClearForm_Recipe
        Else
            Do While Not mrsList.EOF
                If intRow <> 0 And (.TextMatrix(intRow, mIntColNO) = mrsList!NO And .TextMatrix(intRow, mIntCol����) = mrsList!����) And mcondition.bln������� And ((mcondition.int�б����� = mListType.����ҩ Or mcondition.int�б����� = mListType.����ҩ)) Then
                    If Val(NVL(mrsList!���id)) <> 0 And .TextMatrix(intRow, mIntCol�����) <> mrsList!����� Then
                        .Cell(flexcpPicture, intRow, mIntCol���, intRow, mIntCol���) = Me.imgList.ListImages(42).Picture
                        .TextMatrix(intRow, mIntCol�����) = 2
                    End If
                Else
                    intRow = intRow + 1
                    .rows = intRow + 1
            
                    .TextMatrix(intRow, mintcolѡ��) = ""
    
                    If mcondition.int�б����� = mListType.����ҩ Then
                        If NVL(mrsList!����ʱ��) <> "" Then
    '                        .Cell(flexcpPicture, intRow, mIntCol����, intRow, mIntCol����) = Me.imgList.ListImages(39).Picture
    '                        .Cell(flexcpPictureAlignment, intRow, mIntCol����, intRow, mIntCol����) = flexPicAlignLeftCenter
                            .Cell(flexcpFontBold, intRow, mIntCol����, intRow, mIntCol����) = True
                            
                            dateNow = zlDatabase.Currentdate
                            lngTime = DateDiff("n", mrsList!����ʱ��, dateNow)
                            If lngTime > 60 Then
                                .TextMatrix(intRow, mIntCol����) = ">60"
                            Else
                                .TextMatrix(intRow, mIntCol����) = IIf(lngTime < 0, 0, lngTime)
                            End If
                        End If
                    End If
                    
                    If (mcondition.int�б����� = mListType.����ҩ Or mcondition.int�б����� = mListType.����ҩ) And mcondition.bln������� Then
                        If Val(NVL(mrsList!���id)) <> 0 Then
                            If mrsList!����� = 1 Then
                                .Cell(flexcpPicture, intRow, mIntCol���, intRow, mIntCol���) = Me.imgList.ListImages(41).Picture
                                 .TextMatrix(intRow, mIntCol�����) = 1
                            Else
                                .Cell(flexcpPicture, intRow, mIntCol���, intRow, mIntCol���) = Me.imgList.ListImages(42).Picture
                                .TextMatrix(intRow, mIntCol�����) = 2
                            End If
                        Else
                            .Cell(flexcpPicture, intRow, mIntCol���, intRow, mIntCol���) = Me.imgList.ListImages(41).Picture
                            .TextMatrix(intRow, mIntCol�����) = 1
                        End If
                    End If
                    
                    
                    If mcondition.int�б����� <> mListType.��ҩ Then
                        .TextMatrix(intRow, mIntColǩ������) = NVL(mrsList!ǩ��ʱ��)
                    End If
                    
                    If mrsList!�������� = 1 Then
                        .TextMatrix(intRow, mIntCol��ɫ) = "����"
                    ElseIf mrsList!�������� = 2 Then
                        .TextMatrix(intRow, mIntCol��ɫ) = "����"
                    ElseIf mrsList!�������� = 3 Then
                        .TextMatrix(intRow, mIntCol��ɫ) = "����"
                    ElseIf mrsList!�������� = 4 Then
                        .TextMatrix(intRow, mIntCol��ɫ) = "��һ"
                    ElseIf mrsList!�������� = 5 Then
                        .TextMatrix(intRow, mIntCol��ɫ) = "����"
                    Else
                        .TextMatrix(intRow, mIntCol��ɫ) = "��ͨ"
                    End If
    
                    .TextMatrix(intRow, mIntCol��������) = IIf(IsNull(mrsList!��������), "", mrsList!��������)
                    .TextMatrix(intRow, mIntCol��־) = mrsList!��־
                    .TextMatrix(intRow, mIntCol����) = IIf(IsNull(mrsList!����), "", mrsList!����)
                    
                    .TextMatrix(intRow, mIntCol����) = mrsList!����
                    .TextMatrix(intRow, mIntCol�շ�) = mrsList!���շ�
                    .TextMatrix(intRow, mIntCol��ҩ��) = IIf(IsNull(mrsList!��ҩ��), "", mrsList!��ҩ��)
                    .TextMatrix(intRow, mIntColNO) = mrsList!NO
    
                    Select Case intType
                        Case mListType.����ҩ
                            If mrsList!��ӡ״̬ = 1 Then    'ֻ�д���ҩ���ڼ�¼�����С���ӡ״̬��
                                .Cell(flexcpPicture, intRow, mIntColNO) = Me.imgList.ListImages("��ӡ").Picture
                                .Cell(flexcpPictureAlignment, intRow, mIntColNO) = flexPicAlignLeftCenter
                            End If
                        
                        Case mListType.��ҩ
                            If mblnȡҩȷ�� = True Then
                                .TextMatrix(intRow, mIntColδȡҩ) = NVL(mrsList!δȡҩ, 0)
                                If Val(.TextMatrix(intRow, mIntColδȡҩ)) = 1 Then
                                    .Cell(flexcpPicture, intRow, mIntColNO) = Me.imgList.ListImages("δȡҩ").Picture
                                    .Cell(flexcpPictureAlignment, intRow, mIntColNO) = flexPicAlignRightCenter
                                End If
                            End If
                    End Select
                    
                    .TextMatrix(intRow, mIntCol����) = IIf(IsNull(mrsList!����), "", mrsList!����)
                    
                    .TextMatrix(intRow, mIntCol���) = GetFormat(Val(mrsList!���), mintMoneyDigit)
                    .TextMatrix(intRow, mIntColʵ�ս��) = GetFormat(Val(mrsList!ʵ�ս��), mintMoneyDigit)
                    .TextMatrix(intRow, mIntCol����) = mrsList!����
                    .TextMatrix(intRow, mIntCol�ɲ���) = mrsList!�ɲ���
                    .TextMatrix(intRow, mIntCol˵��) = IIf(IsNull(mrsList!˵��), "", mrsList!˵��)
                    .TextMatrix(intRow, mIntCol���￨��) = IIf(IsNull(mrsList!���￨��), "", mrsList!���￨��)
                    
                    .TextMatrix(intRow, mIntCol�����) = IIf(IsNull(mrsList!�����), "", mrsList!�����)
                    .TextMatrix(intRow, mIntCol���֤) = IIf(IsNull(mrsList!���֤��), "", mrsList!���֤��)
                    .TextMatrix(intRow, mIntColIC��) = IIf(IsNull(mrsList!IC����), "", mrsList!IC����)
                    .TextMatrix(intRow, mIntCol����ID) = IIf(IsNull(mrsList!����ID), "", mrsList!����ID)
                    .TextMatrix(intRow, mIntColҽ����) = IIf(IsNull(mrsList!ҽ����), "", mrsList!ҽ����)
                    .TextMatrix(intRow, mIntColסԺ��) = IIf(IsNull(mrsList!סԺ��), "", mrsList!סԺ��)
                    
                    .TextMatrix(intRow, mIntCol�����־) = mrsList!�����־
                    .TextMatrix(intRow, mIntCol��¼����) = mrsList!��¼����
                    .TextMatrix(intRow, mIntCol�շ����) = mrsList!�շ����
                    
                    .TextMatrix(intRow, mIntColƴ������) = mPinYin(IIf(IsNull(mrsList!����), "", mrsList!����))
                    .TextMatrix(intRow, mIntCol��ʼ���) = mWBX(IIf(IsNull(mrsList!����), "", mrsList!����), 1)
                    
                    If intType = mListType.��ҩȷ�� Then
                        .TextMatrix(intRow, mIntCol�Ŷ�״̬) = NVL(mrsList!�Ŷ�״̬)
                    End If
                    
                    If intType <> mListType.��ҩ Then
                        .TextMatrix(intRow, mIntCol��ҩ����) = NVL(mrsList!��ҩ����)
                    End If
                    
                    .Cell(flexcpBackColor, intRow, mIntCol��ɫ, intRow, mIntCol��ɫ) = Val(Split(mstrUserRecipeColor, ";")(Val(mrsList!��������)))
                    
                    '������ɫ
                    lngColor = IIf(mcondition.int�б����� <> mListType.��ҩ Or mrsList!�ɲ��� = 0, &H80000008, IIf(mrsList!�ɲ��� = 1, glng����, IIf(mrsList!�ɲ��� = 2, glng��ҩ, glng��ҩ)))
                    .Cell(flexcpForeColor, intRow, 1, intRow, .Cols - 1) = lngColor
                    .Cell(flexcpForeColor, intRow, mIntCol����, intRow, mIntCol����) = vbRed
                    
                    '���������ò�ͬǰ��ɫ������Ӵ�����
                    .Cell(flexcpForeColor, intRow, mIntCol����, intRow, mIntCol����) = zlDatabase.GetPatiColor(IIf(IsNull(mrsList!��������), "", mrsList!��������))
                End If
                mrsList.MoveNext
            Loop
            
            If mcondition.bln����ģʽ = True Then
                .Cell(flexcpPicture, 0, mintcolѡ��, .rows - 1, mintcolѡ��) = LoadResPicture("checked", vbResBitmap)
                .Cell(flexcpPictureAlignment, 0, mintcolѡ��, .rows - 1, mintcolѡ��) = flexAlignCenterCenter
                .Cell(flexcpText, 0, mIntCol��־, .rows - 1, mIntCol��־) = 1
            End If
            
            If mintLocate = 0 Or mintLocate > intRow Then mintLocate = 1
            If mcondition.bln����ģʽ Then
                .Row = 1
                .TopRow = .Row
            End If
        End If
        
        '����״̬�¶�λ
        If mcondition.bln����ģʽ = False And mstrFind <> "" Then
            mFindProcess.StartRow = 1
'            FindSpecialRow mintFindType, mstrFind
        End If
        
        mblnSortByName = False
        
        '�ָ��û��������
        '����б�����
        'ֵ=�к�|��/����
        strSort = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & "ҩƷ������ҩ", "�����嵥����" & mcondition.int�б�����, "")
        If strSort <> "" And InStr(1, strSort, "|") > 0 Then    'ֵ����Ϊ�գ�����Ҫ�зָ���
            If Val(Split(strSort, "|")(0)) > 0 And Val(Split(strSort, "|")(0)) < .Cols - 1 Then     '���ص������Ҫ���嵥�з�Χ��
                If .ColHidden(Val(Split(strSort, "|")(0))) = False Then      '���ص��б��벻������
                    .ColSort(Val(Split(strSort, "|")(0))) = IIf(Val(Split(strSort, "|")(1)) = 2, 2, 1)
                    .Col = Val(Split(strSort, "|")(0))
                    .Sort = flexSortUseColSort
                    
                    If Val(Split(strSort, "|")(0)) = mIntCol���� Then
                        mblnSortByName = True
                    End If
                End If
            End If
        End If
        
        If mcondition.bln����ģʽ = False Then
            If mintLocate = 0 Or mintLocate > intRow Then mintLocate = 1
            
            If mblnSortByName = True And mstrLastName <> "" Then
                '����������ʱ�����ϴη�ҩ���˵����ŵ���
                StrFind = mstrLastName
                intFindCol = mIntCol����
            ElseIf strNo <> "" Then
                StrFind = strNo
                intFindCol = mIntColNO
                mintLocate = 1
            Else
                '���ϴ�ѡ���NO����
                StrFind = mstrLastNo
                intFindCol = mIntColNO
                mintLocate = 1
            End If
            
            If StrFind <> "" Then
                lngFindRow = .FindRow(StrFind, mintLocate, intFindCol)
                If lngFindRow > 0 Then
                    .Row = lngFindRow
                Else
                    lngFindRow = .FindRow(StrFind, 1, intFindCol)
                    If lngFindRow > 0 Then
                        .Row = lngFindRow
                    Else
                        .Row = mintLocate
                    End If
                End If
            Else
                If .rows > 1 Then .Row = 1
            End If
            .TopRow = .Row
        End If
        
        .Redraw = flexRDDirect
    End With
    
    mblnFreshList = False
End Sub

Private Sub SetColumnValue(ByVal str���� As String, ByVal intValue As Integer)
    Select Case str����
        Case "����"
            mIntCol���� = intValue
        Case "����"
            mIntCol��ɫ = intValue
        Case "ѡ��"
            mintcolѡ�� = intValue
        Case "���"
            mIntCol���� = intValue
        Case "NO"
            mIntColNO = intValue
        Case "����"
            mIntCol���� = intValue
        Case "���", "Ӧ�ս��"
            mIntCol��� = intValue
        Case "ʵ�ս��"
            mIntColʵ�ս�� = intValue
        Case "����"
            mIntCol���� = intValue
        Case "ǩ������"
            mIntColǩ������ = intValue
        Case "˵��"
            mIntCol˵�� = intValue
        Case "���￨��"
            mIntCol���￨�� = intValue
        Case "�����"
            mIntCol����� = intValue
    
        Case "���֤"
            mIntCol���֤ = intValue
        Case "IC��"
            mIntColIC�� = intValue
        Case "����ID"
            mIntCol����ID = intValue
        Case "�ɲ���"
            mIntCol�ɲ��� = intValue
        Case "ҽ����"
            mIntColҽ���� = intValue
        Case "סԺ��"
            mIntColסԺ�� = intValue
        Case "��������"
            mIntCol�������� = intValue
    End Select
End Sub

Private Sub vsfList_Click()
    Dim intCheck As Integer
    Dim strCheck As String
    
    With vsfList
        If mcondition.bln����ģʽ = False Then Exit Sub
        If .MouseRow < 0 Then Exit Sub
        If .MouseCol <> mintcolѡ�� Then Exit Sub
        
        If IsNumeric(.TextMatrix(.rows - 1, mIntCol����)) Then
            intCheck = Abs(.Cell(flexcpText, .MouseRow, mIntCol��־, .MouseRow, mIntCol��־) - 1)
        Else
            intCheck = Abs(.Cell(flexcpText, 0, mIntCol��־, 0, mIntCol��־) - 1)
            .TextMatrix(0, mIntCol��־) = intCheck
        End If
        strCheck = IIf(intCheck = 1, "checked", "unchecked")

        If .MouseRow = 0 Then
            If IsNumeric(.TextMatrix(.rows - 1, mIntCol����)) Then .Cell(flexcpText, 0, mIntCol��־, .rows - 1, mIntCol��־) = intCheck
            .Cell(flexcpPicture, 0, mintcolѡ��, .rows - 1, mintcolѡ��) = LoadResPicture(strCheck, vbResBitmap)
            .Cell(flexcpPictureAlignment, 0, mintcolѡ��, .rows - 1, mintcolѡ��) = flexAlignCenterCenter
        Else
            If IsNumeric(.TextMatrix(.rows - 1, mIntCol����)) Then
                .Cell(flexcpText, .MouseRow, mIntCol��־, .MouseRow, mIntCol��־) = intCheck
            Else
                .Cell(flexcpPicture, 0, mintcolѡ��, 0, mintcolѡ��) = LoadResPicture(strCheck, vbResBitmap)
                .Cell(flexcpPictureAlignment, 0, mintcolѡ��, 0, mintcolѡ��) = flexAlignCenterCenter
            End If
            .Cell(flexcpPicture, .MouseRow, mintcolѡ��, .MouseRow, mintcolѡ��) = LoadResPicture(strCheck, vbResBitmap)
            .Cell(flexcpPictureAlignment, .MouseRow, mintcolѡ��, .MouseRow, mintcolѡ��) = flexAlignCenterCenter
        End If
    End With
End Sub

Private Sub vsfList_EnterCell()
    Dim lngColor As Long
    Dim lngNameColor As Long
    
    
    If mblnOutPut = True Then Exit Sub
    If mblnNoRefreshDetail = True Then Exit Sub
    If mblnFindOver = True Then Exit Sub
    
    With vsfList
        If Not gobjPass Is Nothing Then Call gobjPass.zlPassClearLight_YF
        If .Row = 0 Then Exit Sub
        .Cell(flexcpPicture, 1, 0, .rows - 1, 0) = Nothing
        .Cell(flexcpPicture, .Row, 0, .Row, 0) = Me.imgList.ListImages(2).Picture
        
        lngColor = IIf(mcondition.int�б����� <> mListType.��ҩ Or Val(.TextMatrix(.Row, mIntCol�ɲ���)) <= 1, glng����, IIf(Val(.TextMatrix(.Row, mIntCol�ɲ���)) = 2, glng��ҩ, glng��ҩ))
        lngNameColor = .Cell(flexcpForeColor, .Row, mIntCol����, .Row, mIntCol����)
        
        'ѡ����ʱ��ǰ��ɫ�ò���������ɫ��ʶ
        .ForeColorSel = lngNameColor
        
        If Val(.TextMatrix(.Row, mIntCol����)) = 0 Then Exit Sub
        
        If mblnFreshList = False Then mstrLastNo = .TextMatrix(.Row, mIntColNO)
        
        mintLocate = .Row
        
        SetMainComandBars mcondition.int�б�����, .Row
        
        If mcondition.int�б����� = mListType.��ҩ Then
            If Trim(.TextMatrix(.Row, mIntCol˵��)) = "" Then
                frmҩƷ������ҩNew.RefreshDetail_Return Val(.TextMatrix(.Row, mIntCol����)), .TextMatrix(.Row, mIntColNO), .TextMatrix(.Row, mIntCol����), Val(.TextMatrix(.Row, mIntCol�ɲ���)), Val(.TextMatrix(.Row, mIntCol�����־)), Val(.TextMatrix(.Row, mIntCol��¼����))
            Else
                frmҩƷ������ҩNew.RefreshDetail_Return Val(.TextMatrix(.Row, mIntCol����)), .TextMatrix(.Row, mIntColNO), .TextMatrix(.Row, mIntCol����), Val(.TextMatrix(.Row, mIntCol�ɲ���)), Val(.TextMatrix(.Row, mIntCol�����־)), Val(.TextMatrix(.Row, mIntCol��¼����)), False, Val(Mid(.TextMatrix(.Row, mIntCol˵��), InStr(1, .TextMatrix(.Row, mIntCol˵��), "��") + 1, InStr(1, .TextMatrix(.Row, mIntCol˵��), "��") - InStr(1, .TextMatrix(.Row, mIntCol˵��), "��") - 1))
            End If
        Else
            frmҩƷ������ҩNew.RefreshDetail_Send 0, Val(.TextMatrix(.Row, mIntCol����)), .TextMatrix(.Row, mIntColNO), Val(.TextMatrix(.Row, mIntCol�����־)), Val(.TextMatrix(.Row, mIntCol��¼����)), Val(.TextMatrix(.Row, mIntCol�Ŷ�״̬)), Val(.TextMatrix(.Row, mIntCol�����))
        End If
        
        
    End With
End Sub

Private Sub vsfList_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> mintcolѡ�� Then Cancel = True
End Sub

Public Sub SetCalling()
    With Me.vsfList
        If mIntOldRow <> 0 And mIntOldRow <= .rows - 1 Then
            If .Cell(flexcpText, mIntOldRow, mIntCol����, mIntOldRow, mIntCol����) = "" Then
                .Cell(flexcpText, mIntOldRow, mIntCol����, mIntOldRow, mIntCol����) = "0"
                .Cell(flexcpPicture, mIntOldRow, mIntCol����, mIntOldRow, mIntCol����) = Nothing
            End If
        End If
        mIntOldRow = .Row
        
        .Cell(flexcpText, .Row, mIntCol����, .Row, mIntCol����) = ""
        .Cell(flexcpPicture, .Row, mIntCol����, .Row, mIntCol����) = Me.imgList.ListImages(39).Picture
        .Cell(flexcpPictureAlignment, .Row, mIntCol����, .Row, mIntCol����) = flexPicAlignCenterCenter
    End With
End Sub

Public Sub SetSign(ByRef strNo As String)
    Dim i As Integer
    
    With Me.vsfList
        For i = 1 To .rows - 1
            If InStr(1, strNo, .TextMatrix(i, mIntColNO)) <> 0 Then
                strNo = Replace(strNo, .TextMatrix(.Row, mIntColNO), "")
                .Row = i
                Exit Sub
            End If
        Next
    End With
End Sub
