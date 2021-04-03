VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRTFTexts"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cRTFText"
Attribute VB_Ext_KEY = "Member0" ,"cRTFText"
Option Explicit

Private mCol As Collection
Private mvarIsInTable As Boolean
Private mvarSpaceBefore As Long
Private mvarSpaceAfter As Long
Private mvarSpaceLine As Long
Private mvarSpaceFirstLine As Long
Private mvarFontSizeDefault As Long

Public Property Let FontSizeDefault(ByVal vData As Long)
Attribute FontSizeDefault.VB_Description = "Ĭ������ߴ硣"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.FontSizeDefault = 5
    mvarFontSizeDefault = vData
End Property

Public Property Get FontSizeDefault() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.FontSizeDefault
    FontSizeDefault = mvarFontSizeDefault
End Property

Public Property Let SpaceFirstLine(ByVal vData As Long)
Attribute SpaceFirstLine.VB_Description = "��������ֵ��"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.SpaceFirstLine = 5
    Dim i As Long
    For i = 1 To mCol.Count
        mCol(i).SpaceFirstLine = vData
    Next
    
    mvarSpaceFirstLine = vData
End Property

Public Property Get SpaceFirstLine() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.SpaceFirstLine
    SpaceFirstLine = mvarSpaceFirstLine
End Property

Public Property Let SpaceLine(ByVal vData As Long)
Attribute SpaceLine.VB_Description = "�м�ࡣ"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.SpaceLine = 5
    Dim i As Long
    For i = 1 To mCol.Count
        mCol(i).SpaceLine = vData
    Next
    mvarSpaceLine = vData
End Property

Public Property Get SpaceLine() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.SpaceLine
    SpaceLine = mvarSpaceLine
End Property

Public Property Let SpaceAfter(ByVal vData As Long)
Attribute SpaceAfter.VB_Description = "��ĩ��ࡣ"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.SpaceAfter = 5
    Dim i As Long
    For i = 1 To mCol.Count
        mCol(i).SpaceAfter = vData
    Next
    mvarSpaceAfter = vData
End Property

Public Property Get SpaceAfter() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.SpaceAfter
    SpaceAfter = mvarSpaceAfter
End Property

Public Property Let SpaceBefore(ByVal vData As Long)
Attribute SpaceBefore.VB_Description = "��ǰ��ࡣ"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.SpaceBefore = 5
    Dim i As Long
    For i = 1 To mCol.Count
        mCol(i).SpaceBefore = vData
    Next
    mvarSpaceBefore = vData
End Property

Public Property Get SpaceBefore() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.SpaceBefore
    SpaceBefore = mvarSpaceBefore
End Property

Public Property Let IsInTable(ByVal vData As Boolean)
Attribute IsInTable.VB_Description = "�Ƿ��ڱ����С�Ĭ�ϣ���"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.IsInTable = 5
    Dim i As Long
    For i = 1 To mCol.Count
        mCol(i).IsInTable = True
    Next
    mvarIsInTable = vData
End Property

Public Function GetTextRTF() As String
Attribute GetTextRTF.VB_Description = "��ȡ�ı����ϵ�RTF�ַ�����"
    Dim i As Long, strR As String
    For i = 1 To mCol.Count
        strR = strR & mCol(i).GetTextRTF
    Next
    GetTextRTF = strR
End Function

Public Property Get IsInTable() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.IsInTable
    IsInTable = mvarIsInTable
End Property

Public Function LastText() As cRTFText
Attribute LastText.VB_Description = "��ȡ������ӵ��ı�����"
    Set LastText = mCol(mCol.Count)
End Function

Public Function Add(Optional Text As String, Optional sKey As String, _
Optional Alignment As AlignTypeEnum = alm0_Ĭ�϶���, _
Optional Bold As Boolean = False, _
Optional Italic As Boolean = False, Optional ForeColor As ColorEnum = clr00_Ĭ��, _
Optional FontSize As FontSizeEnum = fs00_Ĭ��, Optional FontStyle As FontStyleEnum = ft0_����, _
Optional CharStyle As CharStyleEnum = csĬ��, Optional AnimType As AnimTypeEnum = ani0_��, _
Optional ScaleX As Long = 100, Optional Embo As EmboEnum = emb0_��, _
Optional Expand As Long = 0, Optional OutLine As Boolean = False, _
Optional Shadow As Boolean = False, Optional Strike As Boolean = False, _
Optional StrikeD As Boolean = False, Optional SubChar As Boolean = False, _
Optional SuperChar As Boolean = False, Optional UnderLineType As UnderLineTypeEnum = udl00_��, _
Optional UnderLineColor As ColorEnum = clr00_Ĭ��, Optional Hided As Boolean = False, _
Optional WebHidden As Boolean = False, Optional CharBorder As Boolean = False, _
Optional CharShadow As Long = 100, Optional CharPatForeColor As ColorEnum = clr00_Ĭ��, _
Optional CharPatBackColor As ColorEnum = clr00_Ĭ��, Optional PatType As PatTypeEnum = ptt00_��, _
Optional Highlight As Boolean = False, Optional HighlightColor As ColorEnum = clr00_Ĭ��, _
Optional ListType As ListTypeEnum = ltp00_��, _
Optional IsInTable As Boolean = False) As Long
Attribute Add.VB_Description = "����һ���ı�����"
    '�����¶���
    Dim objNewMember As cRTFText
    Set objNewMember = New cRTFText


    '���ô��뷽��������
    objNewMember.Text = Text
    objNewMember.Alignment = Alignment
    objNewMember.Bold = Bold
    objNewMember.Italic = Italic
    objNewMember.ForeColor = ForeColor
    objNewMember.FontSizeFixup = FontSize
    objNewMember.FontStyle = FontStyle
    objNewMember.CharStyle = CharStyle
    objNewMember.AnimType = AnimType
    objNewMember.ScaleX = ScaleX
    objNewMember.Embo = Embo
    objNewMember.Expand = Expand
    objNewMember.OutLine = OutLine
    objNewMember.Shadow = Shadow
    objNewMember.Strike = Strike
    objNewMember.StrikeD = StrikeD
    objNewMember.SubChar = SubChar
    objNewMember.SuperChar = SuperChar
    objNewMember.UnderLineType = UnderLineType
    objNewMember.UnderLineColor = UnderLineColor
    objNewMember.Hided = Hided
    objNewMember.WebHidden = WebHidden
    objNewMember.CharBorder = CharBorder
    objNewMember.CharShadow = CharShadow
    objNewMember.CharPatForeColor = CharPatForeColor
    objNewMember.CharPatBackColor = CharPatBackColor
    objNewMember.PatType = PatType
    objNewMember.Highlight = Highlight
    objNewMember.HighlightColor = HighlightColor
    objNewMember.ListType = ListType
    objNewMember.IsInTable = IsInTable
    
'    Set objNewMember.cRTFTexts = cRTFTexts
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    '�����Ѵ����Ķ���
    Add = mCol.Count
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As cRTFText
Attribute Item.VB_UserMemId = 0
    '���ü����е�һ��Ԫ��ʱʹ�á�
    'vntIndexKey �������ϵ�������ؼ��֣�
    '����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
Attribute Count.VB_Description = "�ı�����������"
    '���������е�Ԫ����ʱʹ�á��﷨��Debug.Print x.Count
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
Attribute Remove.VB_Description = "ɾ��ָ�����ı�����"
    'ɾ�������е�Ԫ��ʱʹ�á�
    'vntIndexKey ����������ؼ��֣�����ΪʲôҪ����Ϊ Variant ��ԭ��
    '�﷨��x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    '������������ For...Each �﷨ö�ٸü��ϡ�
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    '������󴴽�����
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    '����ֹ���ƻ�����
    Set mCol = Nothing
End Sub
