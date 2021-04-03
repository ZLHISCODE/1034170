VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'######################################################################################
'##ģ �� ����cFont.cls
'##�� �� �ˣ�����ΰ
'##��    �ڣ�2005��5��1��
'##�� �� �ˣ�
'##��    �ڣ�
'##��    �������������ࡣ��Ӧ��TOM�ӿ��е�Font��
'##��    ����
'######################################################################################

Option Explicit

Public Enum UnderLineEnum
    cprNone = 0
    cprDotted = 4
    cprDash = 5
    cprDashDot = 6
    cprDashDotDot = 7
    cprWave = 8
    cprThick = 9
    cprHair = 10
End Enum

Public Enum UnderlineColorEnum
    cprUnderLineColor_Black = &H0
    cprUnderLineColor_Blue = &H10
    cprUnderLineColor_Cyan = &H20
    cprUnderLineColor_LimeGreen = &H30
    cprUnderLineColor_Magenta = &H40
    cprUnderLineColor_Red = &H50
    cprUnderLineColor_Yellow = &H60
    cprUnderLineColor_White = &H70
    cprUnderLineColor_DarkBlue = &H80
    cprUnderLineColor_DarkCyan = &H90
    cprUnderLineColor_Green = &HA0
    cprUnderLineColor_DarkMagenta = &HB0
    cprUnderLineColor_Brown = &HC0
    cprUnderLineColor_OliveGreen = &HD0
    cprUnderLineColor_DarkGray = &HE0
    cprUnderLineColor_Gray = &HF0
End Enum

Private m_Start As Long         'Range����ʼλ��
Private m_End As Long           'Range����ֹλ��
Private m_ReadOnly As Boolean   '�Ƿ�ֻ��
Private m_bSelection As Boolean '�Ƿ���Selection����

Private TOM As cTextDocument     'TOM����

Public Property Let UnderlineColor(vData As UnderlineColorEnum)
    'ע�⣺�»�����ɫ��һ�����־÷�����ֻ���ڱ༭������������Ч��һ������ΪRTF������ɫ���»��߾���ʧ��
    Dim fmt As CHARFORMAT2
    Dim udl As UnderLineEnum
    
    udl = IIf(m_bSelection, TOM.TextDocument.Selection.Font.Underline, TOM.TextDocument.Range(m_Start, m_End).Font.Underline)
    If udl = tomUndefined Then GoTo LL
    fmt.cbSize = Len(fmt)
    fmt.dwMask = CFM_UNDERLINETYPE
    fmt.bUnderlineType = CByte(CByte(udl) Or CByte(vData))
    
    SendMessage TOM.hWnd, EM_SETCHARFORMAT, SCF_SELECTION, fmt
LL:
End Property

Public Property Get UnderlineColor() As UnderlineColorEnum
    'ע�⣺�»�����ɫ��һ�����־÷�����ֻ���ڱ༭������������Ч��һ������ΪRTF������ɫ���»��߾���ʧ��
    Dim fmt As CHARFORMAT2
    
    fmt.cbSize = Len(fmt)
    SendMessage TOM.hWnd, EM_GETCHARFORMAT, SCF_SELECTION, fmt
    If (fmt.dwMask And CFM_UNDERLINETYPE) = 0 Then UnderlineColor = cprUnderLineColor_Black
    UnderlineColor = CByte(fmt.bUnderlineType And &HF0)
End Property

Public Function GetFontFmt() As cFontFormat
    '��ȡ��������
    Dim R As New cFontFormat
    R.AllCaps = Me.AllCaps
    R.BackColor = Me.BackColor
    R.Bold = Me.Bold
    R.ForeColor = Me.ForeColor
    R.Hidden = Me.Hidden
    R.Italic = Me.Italic
    R.Link = Me.Link
    R.Name = Me.Name
    R.Protected = Me.Protected
    R.Size = Me.Size
    R.Strikethrough = Me.Strikethrough
    R.Subscript = Me.Subscript
    R.Superscript = Me.Superscript
    R.Underline = Me.Underline
    Set GetFontFmt = R
End Function

Public Function SetFontFmt(ByRef vData As cFontFormat)
    '������������
    '������ blnExtend ��չģʽ����ʾ�Ƿ������������ã�Ĭ��ΪFalse����ʾ��
    '       ����������ɫΪ SPECIALBACKCOLOR
    '       �������»���Ϊ cprWave ������
    '       ������ǰ��ɫΪ REVISION_COLOR
    '˵����
    'һ�ɲ��������ñ��������ء�ɾ����
    
'    Me.AllCaps = vData.AllCaps
'    Me.Hidden = vData.Hidden
'    Me.Link = vData.Link
'    Me.Protected = vData.Protected
'    Me.Strikethrough = vData.Strikethrough
'    If blnExtend Then
'        '��չģʽ
'        Me.BackColor = vData.BackColor
'        Me.ForeColor = vData.ForeColor
'        Me.Underline = vData.Underline
'    Else
'        Me.BackColor = IIf(vData.BackColor = SPECIALBACKCOLOR, &H99FFFF, vData.BackColor)   '�滻Ϊǳ��ɫ
'        Me.ForeColor = IIf(vData.ForeColor = REVISION_COLOR, vbRed, vData.ForeColor)        '�滻Ϊ��ɫ
'        Me.Underline = IIf(vData.Underline = cprWave, cprHair, vData.Underline)             '�滻Ϊ���»���
'    End If
    
    Me.Underline = vData.Underline
    Me.Bold = vData.Bold
    Me.Italic = vData.Italic
    Me.Name = vData.Name
    Me.Size = vData.Size
    Me.Subscript = vData.Subscript
    Me.Superscript = vData.Superscript
End Function

Public Sub Init(ByVal cTOM As cTextDocument, lStart As Long, lEnd As Long, Optional bReadOnly As Boolean = False, Optional bSelection As Boolean = False)
    Set TOM = cTOM
    m_Start = lStart
    m_End = lEnd
    m_ReadOnly = bReadOnly
    m_bSelection = bSelection
End Sub

Public Sub Reset()
'������������ΪϵͳĬ��ֵ��
    If m_bSelection Then
        TOM.TextDocument.Selection.Font.Reset (tomDefault)
    Else
        TOM.TextDocument.Range(m_Start, m_End).Font.Reset (tomDefault)
    End If
    TOM.TextDocument.DefaultTabStop = 21
End Sub

Public Property Let Underline(ByVal vData As UnderLineEnum)
Attribute Underline.VB_Description = "�»�����ʽ����ȡ����ֵ��\r\n'    tomNone = 0\r\n'    tomDotted = 4\r\n'    tomDash = 5\r\n'    tomDashDot = 6\r\n'    tomDashDotDot = 7\r\n'    tomWave = 8\r\n'    tomThick = 9\r\n'    tomHair = 10"
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        TOM.TextDocument.Selection.Font.Underline = vData
    Else
        TOM.TextDocument.Range(m_Start, m_End).Font.Underline = vData
    End If
End Property

Public Property Get Underline() As UnderLineEnum
    Dim l As Long
    If m_bSelection Then
        l = TOM.TextDocument.Selection.Font.Underline
    Else
        l = TOM.TextDocument.Range(m_Start, m_End).Font.Underline
    End If
    Select Case l
    Case 0, 4, 5, 6, 7, 8, 9, 10
    Case 1
        l = cprHair
    Case Else
        l = 0
    End Select
    Underline = l
End Property

Public Property Let Link(ByVal vData As Boolean)
Attribute Link.VB_Description = "�Ƿ��ǳ�����"
    If m_ReadOnly Then Exit Property
    Dim tCF2 As CHARFORMAT2
    Dim lR As Long
    tCF2.dwMask = CFM_LINK
    If (vData) Then
       tCF2.dwEffects = CFE_LINK
    End If
    tCF2.cbSize = Len(tCF2)
    lR = SendMessage(TOM.hWnd, EM_SETCHARFORMAT, SCF_SELECTION, tCF2)
End Property

Public Property Get Link() As Boolean
    Dim tCF2 As CHARFORMAT2
    Dim lR As Long
    tCF2.dwMask = CFM_LINK
    tCF2.cbSize = Len(tCF2)
    lR = SendMessage(TOM.hWnd, EM_GETCHARFORMAT, SCF_SELECTION, tCF2)
    Link = ((tCF2.dwEffects And CFE_LINK) = CFE_LINK)
End Property

Public Property Let Superscript(ByVal vData As Boolean)
Attribute Superscript.VB_Description = "�ϱ�"
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        TOM.TextDocument.Selection.Font.Superscript = vData
    Else
        TOM.TextDocument.Range(m_Start, m_End).Font.Superscript = vData
    End If
End Property

Public Property Get Superscript() As Boolean
    Superscript = IIf(m_bSelection, TOM.TextDocument.Selection.Font.Superscript, TOM.TextDocument.Range(m_Start, m_End).Font.Superscript)
End Property

Public Property Let Subscript(ByVal vData As Boolean)
Attribute Subscript.VB_Description = "�±�"
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        TOM.TextDocument.Selection.Font.Subscript = vData
    Else
        TOM.TextDocument.Range(m_Start, m_End).Font.Subscript = vData
    End If
End Property

Public Property Get Subscript() As Boolean
    Subscript = IIf(m_bSelection, TOM.TextDocument.Selection.Font.Subscript, TOM.TextDocument.Range(m_Start, m_End).Font.Subscript)
End Property

Public Property Let Strikethrough(ByVal vData As Boolean)
Attribute Strikethrough.VB_Description = "ɾ����"
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        TOM.TextDocument.Selection.Font.Strikethrough = vData
    Else
        TOM.TextDocument.Range(m_Start, m_End).Font.Strikethrough = vData
    End If
End Property

Public Property Get Strikethrough() As Boolean
    Strikethrough = IIf(m_bSelection, TOM.TextDocument.Selection.Font.Strikethrough, TOM.TextDocument.Range(m_Start, m_End).Font.Strikethrough)
End Property

Public Property Let Size(ByVal vData As Single)
Attribute Size.VB_Description = "����ߴ�"
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        TOM.TextDocument.Selection.Font.Size = vData
    Else
        TOM.TextDocument.Range(m_Start, m_End).Font.Size = vData
    End If
End Property

Public Property Get Size() As Single
    Size = IIf(m_bSelection, TOM.TextDocument.Selection.Font.Size, TOM.TextDocument.Range(m_Start, m_End).Font.Size)
End Property

Public Property Let Protected(ByVal vData As Boolean)
Attribute Protected.VB_Description = "�Ƿ񱣻�"
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        TOM.TextDocument.Selection.Font.Protected = vData
    Else
        TOM.TextDocument.Range(m_Start, m_End).Font.Protected = vData
    End If
End Property

Public Property Get Protected() As Boolean
    Protected = IIf(m_bSelection, TOM.TextDocument.Selection.Font.Protected, TOM.TextDocument.Range(m_Start, m_End).Font.Protected)
End Property

Public Property Let Name(ByVal vData As String)
Attribute Name.VB_Description = "��������"
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        TOM.TextDocument.Selection.Font.Name = vData
    Else
        TOM.TextDocument.Range(m_Start, m_End).Font.Name = vData
    End If
End Property

Public Property Get Name() As String
    Name = IIf(m_bSelection, TOM.TextDocument.Selection.Font.Name, TOM.TextDocument.Range(m_Start, m_End).Font.Name)
End Property

Public Property Let Italic(ByVal vData As Boolean)
Attribute Italic.VB_Description = "б��"
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        TOM.TextDocument.Selection.Font.Italic = vData
    Else
        TOM.TextDocument.Range(m_Start, m_End).Font.Italic = vData
    End If
End Property

Public Property Get Italic() As Boolean
    Italic = IIf(m_bSelection, TOM.TextDocument.Selection.Font.Italic, TOM.TextDocument.Range(m_Start, m_End).Font.Italic)
End Property

Public Property Let Hidden(ByVal vData As Boolean)
Attribute Hidden.VB_Description = "�Ƿ�����"
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        TOM.TextDocument.Selection.Font.Hidden = vData
    Else
        TOM.TextDocument.Range(m_Start, m_End).Font.Hidden = vData
    End If
End Property

Public Property Get Hidden() As Boolean
    Hidden = IIf(m_bSelection, TOM.TextDocument.Selection.Font.Hidden, TOM.TextDocument.Range(m_Start, m_End).Font.Hidden)
End Property

Public Property Let ForeColor(ByVal vData As OLE_COLOR)
Attribute ForeColor.VB_Description = "ǰ��ɫ"
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        TOM.TextDocument.Selection.Font.ForeColor = vData
    Else
        TOM.TextDocument.Range(m_Start, m_End).Font.ForeColor = vData
    End If
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = IIf(m_bSelection, TOM.TextDocument.Selection.Font.ForeColor, TOM.TextDocument.Range(m_Start, m_End).Font.ForeColor)
End Property

Public Sub SetForeColorRGB(ByVal R As Long, ByVal G As Long, ByVal B As Long)
    If m_ReadOnly Then Exit Sub
    If m_bSelection Then
        TOM.TextDocument.Selection.Font.ForeColor = RGB(R, G, B)
    Else
        TOM.TextDocument.Range(m_Start, m_End).Font.ForeColor = RGB(R, G, B)
    End If
End Sub

Public Sub GetForeColorRGB(ByRef R As Long, ByRef G As Long, ByRef B As Long)
    Dim mColor As OLE_COLOR
    mColor = ForeColor
    R = RGBRed(mColor)
    G = RGBGreen(mColor)
    B = RGBBlue(mColor)
End Sub

Public Property Get ForeColorR() As Long
    Dim mColor As OLE_COLOR
    mColor = ForeColor
    ForeColorR = RGBRed(mColor)
End Property

Public Property Get ForeColorG() As Long
    Dim mColor As OLE_COLOR
    mColor = ForeColor
    ForeColorG = RGBGreen(mColor)
End Property

Public Property Get ForeColorB() As Long
    Dim mColor As OLE_COLOR
    mColor = ForeColor
    ForeColorB = RGBBlue(mColor)
End Property

Public Property Let Bold(ByVal vData As Boolean)
Attribute Bold.VB_Description = "����"
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        TOM.TextDocument.Selection.Font.Bold = vData
    Else
        TOM.TextDocument.Range(m_Start, m_End).Font.Bold = vData
    End If
End Property

Public Property Get Bold() As Boolean
    Bold = IIf(m_bSelection, TOM.TextDocument.Selection.Font.Bold, TOM.TextDocument.Range(m_Start, m_End).Font.Bold)
End Property

Public Property Let BackColor(ByVal vData As OLE_COLOR)
Attribute BackColor.VB_Description = "����ɫ"
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        TOM.TextDocument.Selection.Font.BackColor = vData
    Else
        TOM.TextDocument.Range(m_Start, m_End).Font.BackColor = vData
    End If
End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = IIf(m_bSelection, TOM.TextDocument.Selection.Font.BackColor, TOM.TextDocument.Range(m_Start, m_End).Font.BackColor)
End Property

Public Property Let AllCaps(ByVal vData As Boolean)
Attribute AllCaps.VB_Description = "ȫ����д"
    If m_ReadOnly Then Exit Property
    If m_bSelection Then
        TOM.TextDocument.Selection.Font.AllCaps = vData
    Else
        TOM.TextDocument.Range(m_Start, m_End).Font.AllCaps = vData
    End If
End Property

Public Property Get AllCaps() As Boolean
    AllCaps = IIf(m_bSelection, TOM.TextDocument.Selection.Font.AllCaps, TOM.TextDocument.Range(m_Start, m_End).Font.AllCaps)
End Property

Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()
    Set TOM = Nothing
End Sub