VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "zlPrintFlxDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Option Explicit

Private mvarAppColor As Long '���ظ���
Private mvarAppFont As StdFont '���ظ���
Private mvarBelowAppRows As zlTabAppRows '���ظ���
Private mvarEmptyDown As Single '���ظ���
Private mvarEmptyLeft As Single '���ظ���
Private mvarEmptyUp As Single '���ظ���
Private mvarTitle As zlTabTitle '���ظ���
Private mvarUnderAppRows As zlTabAppRows '���ظ���
Private mvarUseGrdColor As Boolean '���ظ���
Private mvarBodyHead As Object  '���ظ���
Private mvarBodyGrid As Object  '���ظ���
Private mvarDataSource As Object  '���ظ���
Private mvarFixRow As Integer '�ֲ�����
Private mvarFixCol As Integer '�ֲ�����
Private mvarEmptyRight As Single '�ֲ�����
Private mvarPageHeader As Single '�ֲ�����
Private mvarPageFooter As Single '�ֲ�����
Private mvarHeader As String '�ֲ�����
Private mvarFooter As String '�ֲ�����
Private mvarPrivileged As Boolean '�ֲ�����
Public Property Get Privileged() As Boolean
'����������ֵʱ�ڲ����ұ�ʹ�á�
'Syntax: Debug.Print X.Privileged
    Privileged = mvarPrivileged
End Property

Public Property Let Privileged(ByVal vData As Boolean)
'�������Ը�ֵʱ�ڲ������ʹ�á�
'Syntax: X.Privileged = True
    mvarPrivileged = vData
End Property

Public Property Let Footer(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Footer = 5
    mvarFooter = vData
End Property


Public Property Get Footer() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Footer
    Footer = mvarFooter
End Property



Public Property Let Header(ByVal vData As String)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Header = 5
    mvarHeader = vData
End Property


Public Property Get Header() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Header
    Header = mvarHeader
End Property



Public Property Let PageFooter(ByVal vData As Single)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.PageFooter = 5
    mvarPageFooter = vData
End Property


Public Property Get PageFooter() As Single
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.PageFooter
    PageFooter = mvarPageFooter
End Property



Public Property Let PageHeader(ByVal vData As Single)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.PageHeader = 5
    mvarPageHeader = vData
End Property


Public Property Get PageHeader() As Single
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.PageHeader
    PageHeader = mvarPageHeader
End Property



Public Property Let EmptyRight(ByVal vData As Single)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.EmptyRight = 5
    mvarEmptyRight = vData
End Property


Public Property Get EmptyRight() As Single
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.EmptyRight
    EmptyRight = mvarEmptyRight
End Property



Public Property Let FixCol(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.FixCol = 5
    mvarFixCol = vData
End Property


Public Property Get FixCol() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.FixCol
    FixCol = mvarFixCol
End Property



Public Property Let FixRow(ByVal vData As Integer)
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.FixRow = 5
    mvarFixRow = vData
End Property


Public Property Get FixRow() As Integer
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.FixRow
    FixRow = mvarFixRow
End Property




Public Property Set DataSource(ByVal vData As Object)
'���Ѷ���ֵ������ʱ�� Set ������ʹ�á�
'Syntax: Set x.DataSource = Form1
    Set mvarDataSource = vData
End Property


Public Property Get DataSource() As Object
'����������ֵʱ�ڲ����ұ�ʹ�á�
'Syntax: Debug.Print X.DataSource
    Set DataSource = mvarDataSource
End Property

Public Property Set BodyGrid(ByVal vData As Object)
'���Ѷ���ֵ������ʱ�� Set ������ʹ�á�
'Syntax: Set x.BodyGrid = Form1
    Set mvarBodyGrid = vData
End Property


Public Property Get BodyGrid() As Object
'����������ֵʱ�ڲ����ұ�ʹ�á�
'Syntax: Debug.Print X.BodyGrid
    Set BodyGrid = mvarBodyGrid
End Property



Public Property Set BodyHead(ByVal vData As Object)
'���Ѷ���ֵ������ʱ�� Set ������ʹ�á�
'Syntax: Set x.BodyHead = Form1
    Set mvarBodyHead = vData
End Property


Public Property Get BodyHead() As Object
'����������ֵʱ�ڲ����ұ�ʹ�á�
'Syntax: Debug.Print X.BodyHead
    Set BodyHead = mvarBodyHead
End Property



Public Property Let UseGrdColor(ByVal vData As Boolean)
'�������Ը�ֵʱ�ڲ������ʹ�á�
'Syntax: X.UseGrdColor = 5
    mvarUseGrdColor = vData
End Property


Public Property Get UseGrdColor() As Boolean
'����������ֵʱ�ڲ����ұ�ʹ�á�
'Syntax: Debug.Print X.UseGrdColor
    UseGrdColor = mvarUseGrdColor
End Property



Public Property Set UnderAppRows(ByVal vData As Object)
'���Ѷ���ֵ������ʱ�� Set ������ʹ�á�
'Syntax: Set x.UnderAppRows = Form1
    Set mvarUnderAppRows = vData
End Property


Public Property Get UnderAppRows() As zlTabAppRows
'����������ֵʱ�ڲ����ұ�ʹ�á�
'Syntax: Debug.Print X.UnderAppRows
    Set UnderAppRows = mvarUnderAppRows
End Property



Public Property Set Title(ByVal vData As Object)
'���Ѷ���ֵ������ʱ�� Set ������ʹ�á�
'Syntax: Set x.Title = Form1
    Set mvarTitle = vData
End Property


Public Property Get Title() As zlTabTitle
'����������ֵʱ�ڲ����ұ�ʹ�á�
'Syntax: Debug.Print X.Title
    Set Title = mvarTitle
End Property

Public Property Let EmptyUp(ByVal vData As Single)
'�������Ը�ֵʱ�ڲ������ʹ�á�
'Syntax: X.EmptyUp = 5
    mvarEmptyUp = vData
End Property


Public Property Get EmptyUp() As Single
'����������ֵʱ�ڲ����ұ�ʹ�á�
'Syntax: Debug.Print X.EmptyUp
    EmptyUp = mvarEmptyUp
End Property



Public Property Let EmptyLeft(ByVal vData As Single)
'�������Ը�ֵʱ�ڲ������ʹ�á�
'Syntax: X.EmptyLeft = 5
    mvarEmptyLeft = vData
End Property


Public Property Get EmptyLeft() As Single
'����������ֵʱ�ڲ����ұ�ʹ�á�
'Syntax: Debug.Print X.EmptyLeft
    EmptyLeft = mvarEmptyLeft
End Property



Public Property Let EmptyDown(ByVal vData As Single)
'�������Ը�ֵʱ�ڲ������ʹ�á�
'Syntax: X.EmptyDown = 5
    mvarEmptyDown = vData
End Property


Public Property Get EmptyDown() As Single
'����������ֵʱ�ڲ����ұ�ʹ�á�
'Syntax: Debug.Print X.EmptyDown
    EmptyDown = mvarEmptyDown
End Property



Public Property Set BelowAppRows(ByVal vData As Object)
'���Ѷ���ֵ������ʱ�� Set ������ʹ�á�
'Syntax: Set x.BelowAppRows = Form1
    Set mvarBelowAppRows = vData
End Property


Public Property Get BelowAppRows() As zlTabAppRows
'����������ֵʱ�ڲ����ұ�ʹ�á�
'Syntax: Debug.Print X.BelowAppRows
    Set BelowAppRows = mvarBelowAppRows
End Property



Public Property Set AppFont(ByVal vData As Object)
'���Ѷ���ֵ������ʱ�� Set ������ʹ�á�
'Syntax: Set x.AppFont = Form1
    Set mvarAppFont = vData
End Property


Public Property Get AppFont() As StdFont
'����������ֵʱ�ڲ����ұ�ʹ�á�
'Syntax: Debug.Print X.AppFont
    Set AppFont = mvarAppFont
End Property



Public Property Let AppColor(ByVal vData As Long)
'�������Ը�ֵʱ�ڲ������ʹ�á�
'Syntax: X.AppColor = 5
    mvarAppColor = vData
End Property


Public Property Get AppColor() As Long
'����������ֵʱ�ڲ����ұ�ʹ�á�
'Syntax: Debug.Print X.AppColor
    AppColor = mvarAppColor
End Property



Private Sub Class_Initialize()
    '������ zlPrint1Grd ��ʱ������ mTitle ����
    Set mvarTitle = New zlTabTitle
    '������ zlPrint1Grd ��ʱ������ mAppFont����
    Set mvarAppFont = New StdFont
    mvarAppFont.Name = "����"
    mvarAppFont.Size = 9
    mvarAppFont.Bold = False
    mvarAppFont.Italic = False
    '������ zlPrint1Grd ��ʱ������ mUnderAppRows����
    Set mvarUnderAppRows = New zlTabAppRows
    '������ zlPrint1Grd ��ʱ������ mBelowAppRows����
    Set mvarBelowAppRows = New zlTabAppRows
    '������ zlPrint1Grd ��ʱ�������������
    mvarAppColor = 0
    mvarEmptyUp = 25
    mvarEmptyDown = 25
    mvarEmptyLeft = 19
    mvarEmptyRight = 19
    mvarPageHeader = 13
    mvarPageFooter = 13
    mvarUseGrdColor = True
    mvarFooter = ";;"
    mvarHeader = ";;"
    Set mvarDataSource = New ADODB.Recordset
    mvarPrivileged = False
End Sub


Private Sub Class_Terminate()
    Set mvarBodyHead = Nothing
    Set mvarBodyGrid = Nothing
    Set mvarDataSource = Nothing
    Set mvarBelowAppRows = Nothing
    Set mvarUnderAppRows = Nothing
    Set mvarTitle = Nothing
    Set mvarAppFont = Nothing
End Sub
