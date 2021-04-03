VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEPRCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'################################################################################################################
'##ģ �� ����cEPRCell.cls
'##�� �� �ˣ�����ΰ
'##��    �ڣ�2005��5��1��
'##�� �� �ˣ�
'##��    �ڣ�
'##��    ��������Ԫ��
'##��    ����
'################################################################################################################
Option Explicit

'################################################################################################################
'## ö�ٳ���
'################################################################################################################

'��Ԫ������
Public Enum CellTypeEnum
    cprCTEText = 0          '��ͨ�ı�
    cprCTEElement = 1       '������Ҫ��
End Enum

'################################################################################################################
'## �ڲ�����
'################################################################################################################

Private mvarKey As Long                         '������
Private mvarID As Double                          'ID����Ԫ/�ı�������Ҫ�ص�ID��
Private mvar�ļ�ID As Long                      '�ļ�ID
Private mvar��ID As Double                        '��ID������ID��
Private mvar������� As Double                    '�������
Private mvar�������� As Boolean                 '�����������
Private mvar��ʼ�� As Long              '��ʼ�棬ר���ڡ����Ӳ�����¼��
Private mvar��ֹ�� As Long              '��ֹ�棬ר���ڡ����Ӳ�����¼��
Private mvar�����ı� As String                  '�����ı�

Private mvarRow As Long                         '�У���1��ʼ���
Private mvarCol As Long                         '�У���1��ʼ���
Private mvarWidth As Long                       '��Ԫ�����
Private mvarHeight As Long                      '��Ԫ��߶�
Private mvarMergeNo As String                   '�ϲ���
Private mvarCellType As CellTypeEnum            '�ڵ�����
Private mvarElementKey As Long                 '����Ҫ�ص�Key
Private mvarPictureKey As Long                 '����Ҫ�ص�Key

'������
Private mvarMargin As Long              '��Ԫ��߾࣬Ĭ��Ϊ30
Private mvarSingleLine As Boolean       '�Ƿ�����ʾ��Ĭ��ΪFalse
Private mvarForeColor As OLE_COLOR      'ǰ��ɫ��Ĭ��Ϊ��ɫ
Private mvarBackColor As OLE_COLOR      '����ɫ��Ĭ��Ϊ��ɫ
Private mvarGridLineColor As OLE_COLOR  '�߿���ɫ��Ĭ��Ϊ��ɫ
Private mvarGridLineWidth As Long       '�߿���ȣ�Ĭ��Ϊ1
Private mvarFixedWidth As Boolean       '�Ƿ��������ȣ�Ĭ��ΪFalse
Private mvarAutoHeight As Boolean       '�Զ��߶ȣ�Ĭ��ΪTrue
Private mvarFontName As String          '��������
Private mvarFontSize As Long            '�����С����λ����ֵ
Private mvarFontBold As Boolean         '����
Private mvarFontItalic As Boolean       'б��
Private mvarFontStrikeout As Boolean    'ɾ����
Private mvarFontUnderline As Boolean    '�»���
Private mvarFontWeight As Long          '��������
Private mvarFormatString As String      '�ı���ʽ
Private mvarIndent As Long              '����ֵ
Private mvarHAlignment As HAlignEnum    'ˮƽ���뷽ʽ
Private mvarVAlignment As VAlignEnum    '��ֱ���뷽ʽ
Private mvarProtected As Boolean        '�Ƿ񱣻�
Private mvarToolTipText As String       '��ʾ�ı�

Public Property Let Margin(ByVal vData As Long)
    mvarMargin = vData
End Property

Public Property Get Margin() As Long
    Margin = mvarMargin
End Property

Public Property Let SingleLine(ByVal vData As Boolean)
    mvarSingleLine = vData
End Property

Public Property Get SingleLine() As Boolean
    SingleLine = mvarSingleLine
End Property

Public Property Let ForeColor(ByVal vData As OLE_COLOR)
    mvarForeColor = vData
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = mvarForeColor
End Property

Public Property Let BackColor(ByVal vData As OLE_COLOR)
    mvarBackColor = vData
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = mvarBackColor
End Property

Public Property Let GridLineColor(ByVal vData As OLE_COLOR)
    mvarGridLineColor = vData
End Property

Public Property Get GridLineColor() As OLE_COLOR
    GridLineColor = mvarGridLineColor
End Property

Public Property Let GridLineWidth(ByVal vData As Long)
    mvarGridLineWidth = vData
End Property

Public Property Get GridLineWidth() As Long
    GridLineWidth = mvarGridLineWidth
End Property

Public Property Let FixedWidth(ByVal vData As Boolean)
    mvarFixedWidth = vData
End Property

Public Property Get FixedWidth() As Boolean
    FixedWidth = mvarFixedWidth
End Property

Public Property Let AutoHeight(ByVal vData As Boolean)
    mvarAutoHeight = vData
End Property

Public Property Get AutoHeight() As Boolean
    AutoHeight = mvarAutoHeight
End Property

Public Property Let FontName(ByVal vData As String)
    mvarFontName = vData
End Property

Public Property Get FontName() As String
    FontName = mvarFontName
End Property

Public Property Let FontSize(ByVal vData As Long)
    mvarFontSize = vData
End Property

Public Property Get FontSize() As Long
    FontSize = mvarFontSize
End Property

Public Property Let FontBold(ByVal vData As Boolean)
    mvarFontBold = vData
End Property

Public Property Get FontBold() As Boolean
    FontBold = mvarFontBold
End Property

Public Property Let FontItalic(ByVal vData As Boolean)
    mvarFontItalic = vData
End Property

Public Property Get FontItalic() As Boolean
    FontItalic = mvarFontItalic
End Property

Public Property Let FontStrikeout(ByVal vData As Boolean)
    mvarFontStrikeout = vData
End Property

Public Property Get FontStrikeout() As Boolean
    FontStrikeout = mvarFontStrikeout
End Property

Public Property Let FontUnderline(ByVal vData As Boolean)
    mvarFontUnderline = vData
End Property

Public Property Get FontUnderline() As Boolean
    FontUnderline = mvarFontUnderline
End Property

Public Property Let FontWeight(ByVal vData As Long)
    mvarFontWeight = vData
End Property

Public Property Get FontWeight() As Long
    FontWeight = mvarFontWeight
End Property

Public Property Let FormatString(ByVal vData As String)
    mvarFormatString = vData
End Property

Public Property Get FormatString() As String
    FormatString = mvarFormatString
End Property

Public Property Let Indent(ByVal vData As Long)
    mvarIndent = vData
End Property

Public Property Get Indent() As Long
    Indent = mvarIndent
End Property

Public Property Let HAlignment(ByVal vData As HAlignEnum)
    mvarHAlignment = vData
End Property

Public Property Get HAlignment() As HAlignEnum
    HAlignment = mvarHAlignment
End Property

Public Property Let VAlignment(ByVal vData As VAlignEnum)
    mvarVAlignment = vData
End Property

Public Property Get VAlignment() As VAlignEnum
    VAlignment = mvarVAlignment
End Property

Public Property Let Protected(ByVal vData As Boolean)
    mvarProtected = vData
End Property

Public Property Get Protected() As Boolean
    Protected = mvarProtected
End Property

Public Property Let ToolTipText(ByVal vData As String)
    mvarToolTipText = vData
End Property

Public Property Get ToolTipText() As String
    ToolTipText = mvarToolTipText
End Property

Public Property Let Key(ByVal vData As Long)
    mvarKey = vData
End Property

Public Property Get Key() As Long
    Key = mvarKey
End Property

Public Property Let ID(ByVal vData As Double)
    mvarID = vData
End Property

Public Property Get ID() As Double
    ID = mvarID
End Property

Public Property Let �ļ�ID(ByVal vData As Long)
    mvar�ļ�ID = vData
End Property

Public Property Get �ļ�ID() As Long
    �ļ�ID = mvar�ļ�ID
End Property

Public Property Let ��ID(ByVal vData As Double)
    mvar��ID = vData
End Property

Public Property Get ��ID() As Double
    ��ID = mvar��ID
End Property

Public Property Let �������(ByVal vData As Double)
    mvar������� = vData
End Property

Public Property Get �������() As Double
    ������� = mvar�������
End Property

Public Property Let ��������(ByVal vData As Boolean)
    mvar�������� = vData
End Property

Public Property Get ��������() As Boolean
    �������� = mvar��������
End Property

Public Property Let ��ʼ��(ByVal vData As Long)
    mvar��ʼ�� = vData
End Property

Public Property Get ��ʼ��() As Long
    ��ʼ�� = IIf(mvar��ʼ�� = 0, 1, mvar��ʼ��)
End Property

Public Property Let ��ֹ��(ByVal vData As Long)
    mvar��ֹ�� = vData
End Property

Public Property Get ��ֹ��() As Long
'    ��ֹ�� = IIf(mvar��ֹ�� = 0, 0, mvar��ֹ��)
    ��ֹ�� = 0
End Property

Public Property Let �����ı�(ByVal vData As String)
    mvar�����ı� = vData
End Property

Public Property Get �����ı�() As String
    �����ı� = mvar�����ı�
End Property

Public Property Let Row(ByVal vData As Long)
    mvarRow = vData
End Property

Public Property Get Row() As Long
    Row = mvarRow
End Property

Public Property Let Col(ByVal vData As Long)
    mvarCol = vData
End Property

Public Property Get Col() As Long
    Col = mvarCol
End Property

Public Property Let Width(ByVal vData As Long)
    mvarWidth = vData
End Property

Public Property Get Width() As Long
    Width = mvarWidth
End Property

Public Property Let Height(ByVal vData As Long)
    mvarHeight = vData
End Property

Public Property Get Height() As Long
    Height = mvarHeight
End Property

Public Property Let MergeNo(ByVal vData As String)
    mvarMergeNo = vData
End Property

Public Property Get MergeNo() As String
    MergeNo = mvarMergeNo
End Property

Public Property Let CellType(ByVal vData As CellTypeEnum)
    mvarCellType = vData
End Property

Public Property Get CellType() As CellTypeEnum
    CellType = mvarCellType
End Property

Public Property Let ElementKey(ByVal vData As Long)
    mvarElementKey = vData
End Property

Public Property Get ElementKey() As Long
    ElementKey = mvarElementKey
End Property

Public Property Let PictureKey(ByVal vData As Long)
    mvarPictureKey = vData
End Property

Public Property Get PictureKey() As Long
    PictureKey = mvarPictureKey
End Property

'################################################################################################################
'## ���ܣ�  ������ٶ���
'##
'## ������  blnWithKey  :�Ƿ���ͬKeyֵһ���ƣ�
'################################################################################################################
Public Function Clone(Optional blnWithKey As Boolean = False) As cEPRCell
    Dim cNode As New cEPRCell
    If blnWithKey Then cNode.Key = Me.Key
    cNode.ID = Me.ID
    cNode.�ļ�ID = Me.�ļ�ID
    cNode.��ID = Me.��ID
    cNode.������� = Me.�������
    cNode.�����ı� = Me.�����ı�
    cNode.Row = Me.Row
    cNode.Col = Me.Col
    cNode.Width = Me.Width
    cNode.Height = Me.Height
    cNode.MergeNo = Me.MergeNo
    cNode.ElementKey = Me.ElementKey
    cNode.PictureKey = Me.PictureKey
    cNode.Margin = Me.Margin
    cNode.SingleLine = Me.SingleLine
    cNode.ForeColor = Me.ForeColor
    cNode.BackColor = Me.BackColor
    cNode.GridLineColor = Me.GridLineColor
    cNode.GridLineWidth = Me.GridLineWidth
    cNode.FixedWidth = Me.FixedWidth
    cNode.AutoHeight = Me.AutoHeight
    cNode.FontName = Me.FontName
    cNode.FontSize = Me.FontSize
    cNode.FontBold = Me.FontBold
    cNode.FontItalic = Me.FontItalic
    cNode.FontStrikeout = Me.FontStrikeout
    cNode.FontUnderline = Me.FontUnderline
    cNode.FontWeight = Me.FontWeight
    cNode.FormatString = Me.FormatString
    cNode.Indent = Me.Indent
    cNode.HAlignment = Me.HAlignment
    cNode.VAlignment = Me.VAlignment
    cNode.Protected = Me.Protected
    cNode.ToolTipText = Me.ToolTipText
    Set Clone = cNode
End Function

'################################################################################################################
'## ���ܣ�  ��ȡ��Ԫ�������ַ���
'################################################################################################################
Public Property Get ��������() As String
'    '��Format�ó���ʽ�ַ����������ݿ�洢��
'    Dim strR As String
'    strR = Me.����Ҫ��Key & "|"
'    strR = strR & Me.Row & "|"
'    strR = strR & Me.Col & "|"
'    strR = strR & Me.MergeNo & "|"
'    strR = strR & Me.Width & "|"
'    strR = strR & Me.Height & "|"
'    strR = strR + IIf(CellFormat.FontBold, "1", "0") & ";"
'    strR = strR + IIf(CellFormat.FontItalic, "1", "0") & ";"
'    strR = strR + IIf(CellFormat.FontStrikeout, "1", "0") & ";"
'    strR = strR + IIf(CellFormat.FontUnderline, "1", "0") & ";"
'    strR = strR + IIf(CellFormat.MergeCells, "1", "0") & ";"
'    strR = strR + IIf(CellFormat.ProtectionHidden, "1", "0") & ";"
'    strR = strR + IIf(CellFormat.ProtectionLocked, "1", "0") & ";"
'    strR = strR + IIf(CellFormat.WordWrap, "1", "0") & "|"
'    strR = strR & CellFormat.PatternStyle & "|"
'    strR = strR & CellFormat.FontSize & "|"
'    strR = strR & CellFormat.AlignHorizontal & "|"
'    strR = strR & CellFormat.AlignVertical & "|"
'    strR = strR & CellFormat.FontCharSet & "|"
'    strR = strR & CellFormat.FontColor & "|"
'    strR = strR & CellFormat.PatternBG & "|"
'    strR = strR & CellFormat.PatternFG & "|"
'    strR = strR & CellFormat.BorderColor(F1HInsideBorder) & ";"
'    strR = strR & CellFormat.BorderColor(F1VInsideBorder) & ";"
'    strR = strR & CellFormat.BorderColor(F1TopBorder) & ";"
'    strR = strR & CellFormat.BorderColor(F1LeftBorder) & ";"
'    strR = strR & CellFormat.BorderColor(F1BottomBorder) & ";"
'    strR = strR & CellFormat.BorderColor(F1RightBorder) & "|"
'    strR = strR & CellFormat.BorderStyle(F1HInsideBorder) & ";"
'    strR = strR & CellFormat.BorderStyle(F1VInsideBorder) & ";"
'    strR = strR & CellFormat.BorderStyle(F1TopBorder) & ";"
'    strR = strR & CellFormat.BorderStyle(F1LeftBorder) & ";"
'    strR = strR & CellFormat.BorderStyle(F1BottomBorder) & ";"
'    strR = strR & CellFormat.BorderStyle(F1RightBorder) & "|"
'    strR = strR & CellFormat.ValidationRule & "|"
'    strR = strR & CellFormat.ValidationRuleLocal & "|"
'    strR = strR & CellFormat.ValidationText & "|"
'    strR = strR & CellFormat.NumberFormat & "|"
'    strR = strR & CellFormat.NumberFormatLocal & "|"
'    strR = strR & Trim(CellFormat.FontName)
'    �������� = strR
    Dim strR As String
    strR = mvarElementKey & "|"
    strR = strR & mvarPictureKey & "|"
    strR = strR & Me.Row & "|"
    strR = strR & Me.Col & "|"
    strR = strR & Me.MergeNo & "|"
    strR = strR & Me.Width & "|"
    strR = strR & Me.Height & "|"
    strR = strR & Me.Margin & "|"
    strR = strR & Me.SingleLine & "|"
    strR = strR & Me.ForeColor & "|"
    strR = strR & Me.BackColor & "|"
    strR = strR & Me.GridLineColor & "|"
    strR = strR & Me.GridLineWidth & "|"
    strR = strR & Me.FixedWidth & "|"
    strR = strR & Me.AutoHeight & "|"
    strR = strR & Me.FontName & "|"
    strR = strR & Me.FontSize & "|"
    strR = strR & Me.FontBold & "|"
    strR = strR & Me.FontItalic & "|"
    strR = strR & Me.FontStrikeout & "|"
    strR = strR & Me.FontUnderline & "|"
    strR = strR & Me.FontWeight & "|"
    strR = strR & Me.FormatString & "|"
    strR = strR & Me.HAlignment & "|"
    strR = strR & Me.VAlignment & "|"
    strR = strR & Me.Protected & "|"
    �������� = strR
End Property

'################################################################################################################
'## ���ܣ�  ͨ��һ���ַ��������õ�Ԫ������
'################################################################################################################
Public Property Let ��������(ByRef vData As String)
'    Dim T As Variant, i As Long, T2 As Variant
'    T = Split(vData, "|")
'    If UBound(T) > 0 Then
'        Me.����Ҫ��Key = T(0)
'        Me.Row = T(1)
'        Me.Col = T(2)
'        Me.MergeNo = T(3)
'        Me.Width = T(4)
'        Me.Height = T(5)
'
'        T2 = Split(T(6), ";")
'        Me.CellFormat.FontBold = T2(0)
'        Me.CellFormat.FontItalic = T2(1)
'        Me.CellFormat.FontStrikeout = T2(2)
'        Me.CellFormat.FontUnderline = T2(3)
'        Me.CellFormat.MergeCells = T2(4)
'        Me.CellFormat.ProtectionHidden = T2(5)
'        Me.CellFormat.ProtectionLocked = T2(6)
'        Me.CellFormat.WordWrap = T2(7)
'
'        Me.CellFormat.PatternStyle = T(7)
'        Me.CellFormat.FontSize = T(8)
'        Me.CellFormat.AlignHorizontal = T(9)
'        Me.CellFormat.AlignVertical = T(10)
'        Me.CellFormat.FontCharSet = T(11)
'        Me.CellFormat.FontColor = T(12)
'        Me.CellFormat.PatternBG = T(13)
'        Me.CellFormat.PatternFG = T(14)
'
'        T2 = Split(T(15), ";")
'        Me.CellFormat.BorderColor(F1HInsideBorder) = T2(0)
'        Me.CellFormat.BorderColor(F1VInsideBorder) = T2(1)
'        Me.CellFormat.BorderColor(F1TopBorder) = T2(2)
'        Me.CellFormat.BorderColor(F1LeftBorder) = T2(3)
'        Me.CellFormat.BorderColor(F1BottomBorder) = T2(4)
'        Me.CellFormat.BorderColor(F1RightBorder) = T2(5)
'
'        T2 = Split(T(16), ";")
'        Me.CellFormat.BorderStyle(F1HInsideBorder) = T2(0)
'        Me.CellFormat.BorderStyle(F1VInsideBorder) = T2(1)
'        Me.CellFormat.BorderStyle(F1TopBorder) = T2(2)
'        Me.CellFormat.BorderStyle(F1LeftBorder) = T2(3)
'        Me.CellFormat.BorderStyle(F1BottomBorder) = T2(4)
'        Me.CellFormat.BorderStyle(F1RightBorder) = T2(5)
'
'        Me.CellFormat.ValidationRule = T(17)
'        Me.CellFormat.ValidationRuleLocal = T(18)
'        Me.CellFormat.ValidationText = T(19)
'        Me.CellFormat.NumberFormat = T(20)
'        Me.CellFormat.NumberFormatLocal = T(21)
'        Me.CellFormat.FontName = T(22)
'    End If
    On Error Resume Next
    Dim T As Variant, i As Long, T2 As Variant
    T = Split(vData, "|")
    If UBound(T) = 22 Then
        Me.ElementKey = T(0)
        Me.Row = T(1)
        Me.Col = T(2)
        Me.MergeNo = T(3)
        Me.Width = T(4)
        Me.Height = T(5)

        T2 = Split(T(6), ";")
        Me.FontBold = T2(0)
        Me.FontItalic = T2(1)
        Me.FontStrikeout = T2(2)
        Me.FontUnderline = T2(3)
        Me.Protected = (Me.ElementKey > 0)
        Me.SingleLine = False
        Me.GridLineColor = vbBlack
        Me.GridLineWidth = 1

        Me.FontSize = T(8)
        Me.ForeColor = T(12)
        Me.BackColor = vbWhite
        Me.FixedWidth = False
        Me.AutoHeight = True
        Me.FontName = T(22)
    ElseIf UBound(T) = 26 Then
        Me.ElementKey = T(0)
        Me.PictureKey = T(1)
        Me.Row = T(2)
        Me.Col = T(3)
        Me.MergeNo = T(4)
        Me.Width = T(5)
        Me.Height = T(6)
        Me.Margin = T(7)
        Me.SingleLine = T(8)
        Me.ForeColor = T(9)
        Me.BackColor = T(10)
        Me.GridLineColor = T(11)
        Me.GridLineWidth = T(12)
        Me.FixedWidth = T(13)
        Me.AutoHeight = T(14)
        Me.FontName = T(15)
        Me.FontSize = T(16)
        Me.FontBold = T(17)
        Me.FontItalic = T(18)
        Me.FontStrikeout = T(19)
        Me.FontUnderline = T(20)
        Me.FontWeight = T(21)
        Me.FormatString = T(22)
        Me.HAlignment = T(23)
        Me.VAlignment = T(24)
        Me.Protected = T(25)
    End If
End Property

'################################################################################################################
'## ���ܣ�  �������ID�Ƿ���Ч
'##
'## ������  eEditType :��ǰ�ı༭ģʽ
'##
'## ˵����  ���û��ID�������ݿ�����ȡһ��ΨһID�š�
'################################################################################################################
Public Sub CheckValidID(eEdtType As EditTypeEnum)
    Select Case eEdtType
    Case cprET_�����ļ�����
        If Me.ID = 0 Then Me.ID = zlDatabase.GetNextId("�����ļ��ṹ")
    Case cprET_ȫ��ʾ���༭
        If Me.ID = 0 Then Me.ID = zlDatabase.GetNextId("������������")
    Case cprET_�������༭, cprET_���������
        If Me.ID = 0 Then Me.ID = GetEPRContentNextId
    End Select
End Sub
