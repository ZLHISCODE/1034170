VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'######################################################################################
'##ģ �� ����cRange.cls
'##�� �� �ˣ�����ΰ
'##��    �ڣ�2005��5��1��
'##�� �� �ˣ�
'##��    �ڣ�
'##��    ����ָ�����򡣶�Ӧ��TOM�ӿ��е�Range()��
'##��    ����
'######################################################################################

Option Explicit

Public Enum CaseTypeEnum
    cprCTLowerCase = tomLowerCase         'Сд.
    cprCTUpperCase = tomUpperCase         '��д.
    cprCTTitleCase = tomTitleCase         '���״�д.
    cprCTSentenceCase = tomSentenceCase   '���״�д.
    cprCTToggleCase = tomToggleCase       'ȡ�෴ֵ.
End Enum

Public Enum StartPosEnum
    cprSPStart = tomStart           '��ʼλ��
    cprSPEnd = tomEnd               '����λ��
End Enum

Public Enum GetPointTypeEnum
    cprGPStart = tomStart           '��ʼλ��
    cprGPEnd = tomEnd               '����λ��
    cprGPTop = TA_TOP               '�߽���εĶ�����Ե
    cprGPBaseLine = TA_BASELINE     '�ı�����
    cprGPBottom = TA_BOTTOM         '�߽���εĵײ���Ե
    cprGPLeft = TA_LEFT             '�߽���ε����Ե
    cprGPCenter = TA_CENTER         '�߽���ε�����λ��
    cprGPRight = TA_RIGHT           '�߽���ε��ұ�Ե
End Enum

Private m_CanCopy As Boolean
Private m_CanPaste As Boolean

Private m_Start As Long         'Range����ʼλ��
Private m_End As Long           'Range����ֹλ��
Private m_ReadOnly As Boolean   '�Ƿ�ֻ��

Private TOM As cTextDocument     'TOM����

Public Property Let Text(ByVal vData As String)
    TOM.TextDocument.Range(m_Start, m_End).Text = vData
End Property

Public Property Get Text() As String
Attribute Text.VB_UserMemId = 0
    Text = TOM.TextDocument.Range(m_Start, m_End).Text
End Property

Public Property Get Para() As cPara
    Dim cP As New cPara
    cP.Init TOM, m_Start, m_End, m_ReadOnly
    Set Para = cP
End Property

Public Property Get Font() As cFont
    Dim cF As New cFont
    cF.Init TOM, m_Start, m_End, m_ReadOnly
    Set Font = cF
End Property

Public Sub Init(ByVal cTOM As cTextDocument, lStart As Long, lEnd As Long, Optional bReadOnly As Boolean = False)
    Set TOM = cTOM
    m_Start = lStart
    m_End = lEnd
    m_ReadOnly = bReadOnly
End Sub

Public Property Let EndPos(vData As Long)
    m_End = vData
End Property

Public Property Get EndPos() As Long
Attribute EndPos.VB_Description = "��ȡ��ǰ��Χ����λ�á�ֻ����"
    EndPos = m_End
End Property

Public Property Let StartPos(vData As Long)
    m_Start = vData
End Property

Public Property Get StartPos() As Long
Attribute StartPos.VB_Description = "��ȡ��ǰ��Χ��ʼλ�á�ֻ����"
    StartPos = m_Start
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

Public Sub Copy()
    TOM.TextDocument.Range(m_Start, m_End).Copy 0
End Sub

Public Sub Cut()
    TOM.TextDocument.Range(m_Start, m_End).Cut 0
End Sub

Public Sub Paste()
    TOM.TextDocument.Range(m_Start, m_End).Paste 0, 0
End Sub

Public Sub ChangeCase(Optional lStyle As CaseTypeEnum)
'��;���ı��Сд��
    TOM.TextDocument.Range(m_Start, m_End).ChangeCase lStyle
End Sub

Public Sub Delete()
'��;: ɾ��
    If m_Start = m_End And TOM.TextDocument.Range(m_Start, m_Start + 2) = vbCrLf And TOM.TextDocument.Range(m_Start, m_Start + 2).Font.Protected = False Then
        TOM.TextDocument.Range(m_Start, m_End).Delete tomCharacter, 2
    Else
        TOM.TextDocument.Range(m_Start, m_End).Delete tomCharacter, IIf(m_Start = m_End, 1, 0)
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
    TOM.TextDocument.Range(m_Start, m_End).ScrollIntoView Start
End Sub

Public Sub Selected()
'��;: ������������Ϊ��ǰ���򲢸�����ʾ?
    TOM.TextDocument.Range(m_Start, m_End).Select
End Sub

Public Sub GetPoint(lType As GetPointTypeEnum, pX As Long, pY As Long)
'��;: ��ȡ��ǰλ������
    Dim lX As Long, lY As Long
    TOM.TextDocument.Range(m_Start, m_End).GetPoint lType, lX, lY
    Dim pt1 As POINTAPI
    pt1.x = 0
    pt1.Y = 0
    ClientToScreen TOM.hWnd, pt1
    pX = (lX - pt1.x) * Screen.TwipsPerPixelX + IIf(True, Screen.TwipsPerPixelX, 0)
    pY = (lY - pt1.Y) * Screen.TwipsPerPixelY + IIf(True, Screen.TwipsPerPixelX, 0)
End Sub


Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()
    Set TOM = Nothing
End Sub