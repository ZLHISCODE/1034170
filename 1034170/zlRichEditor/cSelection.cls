VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'######################################################################################
'##ģ �� ����cSelection.cls
'##�� �� �ˣ�����ΰ
'##��    �ڣ�2005��5��1��
'##�� �� �ˣ�
'##��    �ڣ�
'##��    ������ǰѡ�����򡣶�Ӧ��TOM�ӿ��е�Selection��
'##��    ����
'######################################################################################

Option Explicit

Public Enum SelectionTypeEnum
    cprSTNone = tomNoSelection              '0   No selection and no insertion point.
    cprSTInsertPoint = tomSelectionIP       '1   Insertion point.
    cprSTNormal = tomSelectionNormal        '2   Single nondegenerate range.
    cprSTFrame = tomSelectionFrame          '3   Frame.
    cprSTColumn = tomSelectionColumn        '4   Table column.
    cprSTRow = tomSelectionRow              '5   Table rows.
    cprSTBlock = tomSelectionBlock          '6   Block selection.
    cprSTPicture = tomSelectionInlineShape  '7   Picture.
    cprSTShape = tomSelectionShape          '8   Shape.
End Enum

Public Enum HUnitEnum
    cprHUCharacter = tomCharacter
    cprHUWord = tomWord
End Enum

Public Enum VUnitEnum
    cprVULine = tomLine             '����һ�У�Ĭ��ֵ��
    cprVUParagraph = tomParagraph   '����һ�Σ������ס�
    cprVUScreen = tomScreen         '�ƶ�����Ļ����ĩһ���ַ���
    cprVUWidow = tomWindow          '�ƶ�����������ĩһ���ַ���
End Enum

Private m_CanCopy As Boolean
Private m_CanPaste As Boolean
Private m_ReadOnly As Boolean   '�Ƿ�ֻ��

Private TOM As cTextDocument     'TOM����

Public Property Let Text(ByVal vData As String)
    TOM.TextDocument.Selection.Text = vData
End Property

Public Property Get Text() As String
    Text = TOM.TextDocument.Selection.Text
End Property

Public Property Get Para() As cPara
    Dim cP As New cPara
    cP.Init TOM, TOM.TextDocument.Selection.Start, TOM.TextDocument.Selection.End, m_ReadOnly, True
    Set Para = cP
End Property

Public Property Get Font() As cFont
    Dim cF As New cFont
    cF.Init TOM, StartPos, EndPos, m_ReadOnly, True
    Set Font = cF
End Property

Public Sub Init(ByVal cTOM As cTextDocument, Optional bReadOnly As Boolean = False)
    Set TOM = cTOM
    m_ReadOnly = bReadOnly
End Sub

Public Property Get EndPos() As Long
    EndPos = TOM.TextDocument.Selection.End
End Property

Public Property Get StartPos() As Long
    StartPos = TOM.TextDocument.Selection.Start
End Property

Public Property Get CanPaste() As Boolean
    CanPaste = SendMessageLong(TOM.hWnd, EM_CANPASTE, 0, 0)
End Property

Public Property Get CanCopy() As Boolean
    If TOM.TextDocument.Selection.End > TOM.TextDocument.Selection.Start Then
        CanCopy = True
    Else
        CanCopy = False
    End If
End Property

Public Sub ChangeCase(Optional lStyle As CaseTypeEnum)
'��;���ı��Сд��
    TOM.TextDocument.Selection.ChangeCase lStyle
End Sub

Public Sub Copy()
    TOM.TextDocument.Selection.Copy 0
End Sub

Public Sub Cut()
    TOM.TextDocument.Selection.Cut 0
End Sub

Public Sub Paste()
    TOM.TextDocument.Selection.Paste 0, 0
End Sub

Public Sub Delete()
'��;: ɾ��
    Dim lS As Long, lE As Long
    lS = TOM.TextDocument.Selection.Start
    lE = TOM.TextDocument.Selection.End
    If lS = lE And TOM.TextDocument.Range(lS, lS + 2) = vbCrLf And TOM.TextDocument.Range(lS, lS + 2).Font.Protected = False Then
        TOM.TextDocument.Selection.Delete tomCharacter, 2
    Else
        TOM.TextDocument.Selection.Delete tomCharacter, IIf(lS = lE, 1, 0)
    End If
End Sub

Public Sub Expand()
'��;: չ�����  ���ڴ��ģʽ
    SendMessageLong TOM.hWnd, EM_OUTLINE, EMO_EXPAND, 1
End Sub

Public Sub Collapse()
'��;: �۵����?���ڴ��ģʽ?
    SendMessageLong TOM.hWnd, EM_OUTLINE, EMO_EXPAND, -1
End Sub

Public Sub ScrollIntoView(Start As StartPosEnum)
'��;: ��ѡ���������������ʾ��λ��?bStart����ȡֵΪ: tomEnd��tomStart?
    TOM.TextDocument.Selection.ScrollIntoView Start
End Sub

Public Sub Selected()
'��;: ������������Ϊ��ǰ���򲢸�����ʾ?
    TOM.TextDocument.Selection.Select
End Sub

Public Sub GetPoint(lType As GetPointTypeEnum, pX As Long, pY As Long)
'��;: ��ȡ��ǰλ������
    Dim lX As Long, lY As Long
    TOM.TextDocument.Selection.GetPoint lType, lX, lY
    Dim pt1 As POINTAPI
    pt1.x = 0
    pt1.Y = 0
    ClientToScreen TOM.hWnd, pt1
    
    pX = (lX - pt1.x) * Screen.TwipsPerPixelX + IIf(True, Screen.TwipsPerPixelX, 0)
    pY = (lY - pt1.Y) * Screen.TwipsPerPixelY + IIf(True, Screen.TwipsPerPixelX, 0)
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'�����Range�����������ĺ�����
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function GetType() As SelectionTypeEnum
    GetType = TOM.TextDocument.Selection.Type
End Function

Public Sub MoveLeft(Optional Unit As HUnitEnum = cprHUCharacter, Optional Count As Long = 1, Optional Extend As Boolean = False)
' ��;��ѡȡ�����ƶ���
' Unit��ȡtomCharacter��tomWord��Ĭ��ΪtomCharacter��
' Count��ʾ�ƶ�����Ŀ��Ĭ��Ϊ1��
' Extend��ʾ�Ƿ�����չ��ǰѡ������ֻ�ı����㣬True��ʾ��չ��False��ʾ�ı����㣬Ĭ��ΪFalse��
    TOM.TextDocument.Selection.MoveLeft Unit, Count, Extend
End Sub

Public Sub MoveRight(Optional Unit As HUnitEnum = cprHUCharacter, Optional Count As Long = 1, Optional Extend As Boolean = False)
'��;��ѡȡ�����ƶ���
    TOM.TextDocument.Selection.MoveRight Unit, Count, Extend
End Sub

Public Sub MoveDown(Optional Unit As VUnitEnum = cprVULine, Optional Count As Long = 1, Optional Extend As Boolean = False)
'��;��ѡȡ�����ƶ���
'Unit��ȡtomLine��tomParagraph��tomScreen��tomWindow��Ĭ��ΪtomLine��
'Count��ʾ�ƶ�����Ŀ��Ĭ��Ϊ1��
'Extend��ʾ�Ƿ�����չ��ǰѡ������ֻ�ı����㣬True��ʾ��չ��False��ʾ�ı����㣬Ĭ��ΪFalse��
    TOM.TextDocument.Selection.MoveDown Unit, Count, Extend
End Sub

Public Sub MoveUp(Optional Unit As VUnitEnum = cprVULine, Optional Count As Long = 1, Optional Extend As Boolean = False)
'��;��ѡȡ�����ƶ���
    TOM.TextDocument.Selection.MoveUp Unit, Count, Extend
End Sub





Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()
    Set TOM = Nothing
End Sub