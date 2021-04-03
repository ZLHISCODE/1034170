VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommonLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'����
'######################################################################################################################

Public Enum USERCOLOR
    ��ɫ = 16711680
    ��ɫ = 16777215
    ��ɫ = 0
    ����ɫ = 10191921
    ������ɫ = 5880731
    ���񱳾�ɫ = 15790320

    ��Ŀ���ⱳ��ɫ = -2147483636
    
End Enum

Public Type USERRECT
    X0 As Single
    Y0 As Single
    X1 As Single
    Y1 As Single
    R0 As Single                          '�м��
    B0 As Single                          '�߽�
    Page As Integer                     'ҳ��
End Type

Public Type USERFONT
    Name As String
    Size As Single
    Bold As Boolean
    Italic As Boolean
    Underline As Boolean
    Strikethrough As Boolean
    ForeColor As Long
    BackColor As Long
    LineWidth As Integer
    LineStyle As Integer
    HAlignment As Byte                              '�������
    VAlignment As Byte                              '�������
    WordWarp As Boolean                             '�Զ�����
    MuliLine As Boolean                             '�����ı�
    AutoFit As Boolean                              '�Զ���Ӧ
    Rows As Integer                                 '����
End Type

Public Type USERPAPER
    Width As Single
    Height As Single
    BorderLeft As Single
    BorderRight As Single
    BorderTop As Single
    BorderBottom As Single
    PageHead As Single
    PageFoot As Single
    SpaceTop As Single
    SpaceBottom As Single
    PaperSize As Integer
    Orientation As Byte
End Type

Public Event FormatData()

'######################################################################################################################
Public Function InitOracle(cnOracle As ADODB.Connection)
    
    Set gcnOracle = cnOracle
    Call InitCommon(gcnOracle)
    
End Function


Public Function ReadLobV2(ByVal strTable As String, _
                          ByVal strLOBField As String, _
                          ByVal strWhere As String, _
                          ByVal strFile As String, _
                          ByVal lngPara As Long, Optional ByVal lngPara1 As Long) As String
    ReadLobV2 = sys.ReadLobV2(strTable, strLOBField, strWhere, strFile, lngPara, lngPara1)
    
End Function

Public Property Get Rect() As USERRECT
    Rect = gobjRect
End Property

Public Property Let Rect(vData As USERRECT)
    gobjRect = vData
End Property

Public Property Get Font() As USERFONT
    Font = gobjFont
End Property

Public Property Let Font(vData As USERFONT)
    gobjFont = vData
    
    gobjDraw.FontName = gobjFont.Name
    gobjDraw.FontSize = gobjFont.Size
    gobjDraw.FontBold = gobjFont.Bold
    gobjDraw.FontItalic = gobjFont.Italic
    
End Property

Public Property Get Paper() As USERPAPER
    Paper = gobjPaper
End Property

Public Property Let Paper(vData As USERPAPER)
    gobjPaper = vData
End Property

Public Property Get Draw() As Object
    Set Draw = gobjDraw
End Property

Public Property Get DataSource(vntIndexKey As String) As clsDataSource
    Set DataSource = gclsDataSources.Item(vntIndexKey)
End Property

Public Property Get RegInfo(ByVal strItem As String) As String
    RegInfo = zlRegInfo(strItem)
End Property

Public Function GetTextWidth(ByVal strText As String) As Single
    GetTextWidth = gobjDraw.TextWidth(strText)
End Function

Public Function GetTextHeight(ByVal strText As String) As Single
    GetTextHeight = gobjDraw.TextHeight(strText)
End Function

Public Function GetTwipsX(ByVal sglNumber As Single) As Single
    GetTwipsX = gobjDraw.ScaleX(sglNumber, vbCentimeters, vbTwips)
End Function

Public Function GetTwipsY(ByVal sglNumber As Single) As Single
    GetTwipsY = gobjDraw.ScaleY(sglNumber, vbCentimeters, vbTwips)
End Function

Public Function GetCentimetersX(ByVal sglNumber As Single) As Single
    GetCentimetersX = gobjDraw.ScaleX(sglNumber, vbTwips, vbCentimeters)
End Function

Public Function GetCentimetersY(ByVal sglNumber As Single) As Single
    GetCentimetersY = gobjDraw.ScaleY(sglNumber, vbTwips, vbCentimeters)
End Function

Public Function GetRows(ByVal objDraw As Object, ByVal strText As String, ByVal lngCX As Long) As Integer
    '******************************************************************************************************************
    '���ܣ���ȡ��Ҫ����������Ϊ�п���Ҫ����
    '������
    '���أ�
    '******************************************************************************************************************
    GetRows = GetLines(objDraw, strText, lngCX)
End Function

Public Function ShowError() As Byte
    
    ShowError = ErrCenter
    
End Function

Public Sub SaveError()
    Call SaveErrLog
End Sub

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
    NVL = zlCommFun.NVL(varValue, DefaultValue)
End Function

Public Function PrintSetForm() As Object
    Set PrintSetForm = frmChildPrintSet
End Function

Public Function CreateParameter() As ADODB.Recordset
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Fields.Append "��������", adVarChar, 30
        .Fields.Append "�������", adVarChar, 4000
        .Fields.Append "��������", adVarChar, 30
        .Open
    End With
    
    Set CreateParameter = rs
    
End Function

Public Function SetParameter(ByRef rs As ADODB.Recordset, ByVal strParameterName As String, ByVal strParameterValue As String, Optional ByVal strParameterType As String = "�ı�") As Boolean
    
    rs.Filter = ""
    rs.Filter = "��������='" & strParameterName & "'"
    If rs.RecordCount = 0 Then rs.AddNew
    rs("��������").value = strParameterName
    rs("�������").value = strParameterValue
    rs("��������").value = strParameterType
    
    SetParameter = True
    
End Function

Public Function GetParameter(ByRef rs As ADODB.Recordset, ByVal strParameterName As String) As String
    
    rs.Filter = ""
    rs.Filter = "��������='" & strParameterName & "'"
    If rs.RecordCount > 0 Then
        GetParameter = CStr(rs("�������").value)
    End If
    
End Function

Public Function ActualLen(ByVal strAsk As String) As Long
    '--------------------------------------------------------------
    '���ܣ���ȡָ���ַ�����ʵ�ʳ��ȣ������ж�ʵ�ʰ���˫�ֽ��ַ�����
    '       ʵ�����ݴ洢����
    '������
    '       strAsk
    '���أ�
    '-------------------------------------------------------------
    ActualLen = LenB(StrConv(strAsk, vbFromUnicode))
End Function

Public Function GetListRecord() As ADODB.Recordset
    If grsList.RecordCount > 0 Then grsList.MoveFirst
    Set GetListRecord = grsList
End Function

Public Function DatabaseObject() As clsDatabase
    
    Set DatabaseObject = zlDatabase
    
End Function

Public Function OpenSQLRecord(ByVal strSQL As String, ByVal strTitle As String, ParamArray arrInput() As Variant) As ADODB.Recordset
    
    Set OpenSQLRecord = zlDatabase.OpenSQLRecord(strSQL, strTitle, arrInput())
    
End Function

Public Function FixText(ByVal strPreText As String, ByVal strLstText As String, ByVal lngCX As Single, ByVal strRepaceChar As String) As String
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strReplace As String
    Dim strText As String
    
    strText = strPreText & strLstText

    gobjDraw.FontName = gobjFont.Name
    gobjDraw.FontSize = gobjFont.Size
    gobjDraw.FontBold = gobjFont.Bold
    gobjDraw.FontItalic = gobjFont.Italic
        
    Do While (True)
        strReplace = strReplace & strRepaceChar
        If gobjDraw.TextWidth(strText & strReplace & strRepaceChar) > lngCX Then
            Exit Do
        End If
    Loop
    
    FixText = strPreText & strReplace & strLstText
    
End Function

Public Function ReadPicture(rsTable As ADODB.Recordset, strField As String, Optional strFile As String) As String
    ReadPicture = zlDatabase.ReadPicture(rsTable, strField, strFile)
End Function

Public Function GetRowText(ByVal objDraw As Object, ByVal strText As String, ByVal intRow As Integer, ByVal lngCX As Long) As String
    '******************************************************************************************************************
    '���ܣ���ȡָ���е����ݣ�������������������������ٸ��ַ���Ȼ����ù��̡�Get_LineText�����ָ��������
    '������
    '���أ�
    '******************************************************************************************************************

    GetRowText = GetLineText2(objDraw, strText, intRow, lngCX)
End Function

Private Sub Class_Initialize()
    gobjFont.Name = IIf(gobjFont.Name = "", "����", gobjFont.Name)
    gobjFont.Size = IIf(gobjFont.Size = 0, 12, gobjFont.Size)
    gobjFont.Bold = False
    gobjFont.Underline = False
    gobjFont.Strikethrough = False
    gobjFont.Italic = False
End Sub