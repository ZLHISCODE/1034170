VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRTFDocument"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Type ImageSizeInf
   Width As Long
   Height As Long
End Type

Private Nodes As cRTFNodes              '��¼�ڵ���������ID
Public Texts As cRTFTexts               '�ı�����
Public Pictures As cRTFPictures         'ͼƬ����
Public Tables As cRTFTables             '���񼯺�
Public Paragraphs As cRTFParagraphs     '���伯��

'��������ֵ�ľֲ�����
Private mvarFileName As String
Private mvarTextRTF As String
Private mRTFHead As String
Private mRTFBody As String
Private mvarPageHead As String
Private mvarPageFooter As String
Private mvarPageNumber As Boolean
Private mvarPageNumberAlign As AlignTypeEnum
Private mvarPageHeadAlign As AlignTypeEnum
Private mvarPageFooterAlign As AlignTypeEnum
Private mvarSpaceBefore As Long
Private mvarSpaceAfter As Long
Private mvarSpaceLine As Long
Private mvarSpaceFirstLine As Long
Private mvarFontSizeDefault As Long
Private mvarDrawingDirection As DrawingDirectionEnum
Private mvarCaption As String
Private mvarDefaultTabWidth As Long
Private mvarViewKind As ViewKindEnum
Private mvarPaperWidth As Long
Private mvarPaperHeight As Long
Private mvarMarginLeft As Long
Private mvarMarginRight As Long
Private mvarMarginTop As Long
Private mvarMarginBottom As Long
Private mvarMarginGutter As Long
Private mvarPageBorderWidth As Long
Private mvarPageBorderStyle As BorderStyleEnum
Private mvarPageBorderColor As ColorEnum
Private mvarPageBorderArt As Long
Private mvarPaperSize As PaperSizeEnum
Private mvarViewScale As Long
Private mvarPageBorderSpace As Long

Public Property Let PageBorderSpace(ByVal vData As Long)
Attribute PageBorderSpace.VB_Description = "ҳ��߿��ı��ľ��롣"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.PageBorderSpace = 5
    mvarPageBorderSpace = vData
End Property


Public Property Get PageBorderSpace() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.PageBorderSpace
    PageBorderSpace = mvarPageBorderSpace
End Property




Public Property Let PageBorderColor(ByVal vData As ColorEnum)
Attribute PageBorderColor.VB_Description = "ҳ��߿���ɫ������"
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.PageBorderColor = Form1
    mvarPageBorderColor = vData
End Property


Public Property Get PageBorderColor() As ColorEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.PageBorderColor
    PageBorderColor = mvarPageBorderColor
End Property



Public Property Let PageBorderStyle(ByVal vData As BorderStyleEnum)
Attribute PageBorderStyle.VB_Description = "ҳ��߿���ʽ������"
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.PageBorderStyle = Form1
    mvarPageBorderStyle = vData
End Property


Public Property Get PageBorderStyle() As BorderStyleEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.PageBorderStyle
    PageBorderStyle = mvarPageBorderStyle
End Property



Public Property Let ViewScale(ByVal vData As Long)
Attribute ViewScale.VB_Description = "�ĵ����ű�����Ĭ��100��"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ViewScale = 5
    mvarViewScale = vData
End Property


Public Property Get ViewScale() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ViewScale
    ViewScale = mvarViewScale
End Property




Public Property Let PaperSize(ByVal vData As PaperSizeEnum)
Attribute PaperSize.VB_Description = "ҳ��ߴ糣����1��8�����Ϊ0����ҳ��ߴ���PaperWidth��PaperHeight������"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.PaperSize = 5
    mvarPaperSize = vData
End Property


Public Property Get PaperSize() As PaperSizeEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.PaperSize
    PaperSize = mvarPaperSize
End Property




Public Property Let PageBorderArt(ByVal vData As Long)
Attribute PageBorderArt.VB_Description = "�����Ա߿�����͡�1��165��"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.PageBorderArt = 5
    mvarPageBorderArt = vData
End Property


Public Property Get PageBorderArt() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.PageBorderArt
    PageBorderArt = mvarPageBorderArt
End Property


Public Property Let PageBorderWidth(ByVal vData As Long)
Attribute PageBorderWidth.VB_Description = "ҳ��߿���ȡ�0��75��"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.PageBorderWidth = 5
    mvarPageBorderWidth = vData
End Property


Public Property Get PageBorderWidth() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.PageBorderWidth
    PageBorderWidth = mvarPageBorderWidth
End Property



Public Property Let MarginGutter(ByVal vData As Long)
Attribute MarginGutter.VB_Description = "װ����λ�á���λ��羡�"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.MarginGutter = 5
    mvarMarginGutter = vData
End Property


Public Property Get MarginGutter() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.MarginGutter
    MarginGutter = mvarMarginGutter
End Property



Public Property Let MarginBottom(ByVal vData As Long)
Attribute MarginBottom.VB_Description = "�±߾ࡣ��λ��羡�"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.MarginBottom = 5
    mvarMarginBottom = vData
End Property


Public Property Get MarginBottom() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.MarginBottom
    MarginBottom = mvarMarginBottom
End Property



Public Property Let MarginTop(ByVal vData As Long)
Attribute MarginTop.VB_Description = "�ϱ߾ࡣ��λ��羡�"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.MarginTop = 5
    mvarMarginTop = vData
End Property


Public Property Get MarginTop() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.MarginTop
    MarginTop = mvarMarginTop
End Property



Public Property Let MarginRight(ByVal vData As Long)
Attribute MarginRight.VB_Description = "�ұ߾ࡣ��λ��羡�"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.MarginRight = 5
    mvarMarginRight = vData
End Property


Public Property Get MarginRight() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.MarginRight
    MarginRight = mvarMarginRight
End Property



Public Property Let MarginLeft(ByVal vData As Long)
Attribute MarginLeft.VB_Description = "��߾ࡣ��λ��羡�"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.MarginLeft = 5
    mvarMarginLeft = vData
End Property


Public Property Get MarginLeft() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.MarginLeft
    MarginLeft = mvarMarginLeft
End Property



Public Property Let PaperHeight(ByVal vData As Long)
Attribute PaperHeight.VB_Description = "ҳ�ߡ���λ��羡�"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.PaperHeight = 5
    mvarPaperHeight = vData
End Property


Public Property Get PaperHeight() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.PaperHeight
    PaperHeight = mvarPaperHeight
End Property



Public Property Let PaperWidth(ByVal vData As Long)
Attribute PaperWidth.VB_Description = "ҳ������λ��羡�"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.PaperWidth = 5
    mvarPaperWidth = vData
End Property


Public Property Get PaperWidth() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.PaperWidth
    PaperWidth = mvarPaperWidth
End Property



Public Property Let ViewKind(ByVal vData As ViewKindEnum)
Attribute ViewKind.VB_Description = "��ͼģʽ��0��5��"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.ViewKind = 5
    mvarViewKind = vData
End Property


Public Property Get ViewKind() As ViewKindEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.ViewKind
    ViewKind = mvarViewKind
End Property



Public Property Let DefaultTabWidth(ByVal vData As Long)
Attribute DefaultTabWidth.VB_Description = "Ĭ��Tab���ȡ�"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.DefaultTabWidth = 5
    mvarDefaultTabWidth = vData
End Property


Public Property Get DefaultTabWidth() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.DefaultTabWidth
    DefaultTabWidth = mvarDefaultTabWidth
End Property



Public Property Let Caption(ByVal vData As String)
Attribute Caption.VB_Description = "�ĵ�������⡣"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.Caption = 5
    mvarCaption = vData
End Property


Public Property Get Caption() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.Caption
    Caption = mvarCaption
End Property



Public Property Let DrawingDirection(ByVal vData As DrawingDirectionEnum)
Attribute DrawingDirection.VB_Description = "��ͼ����Ĭ��Ϊˮƽ��ͼ��"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.DrawingDirection = 5
    mvarDrawingDirection = vData
End Property


Public Property Get DrawingDirection() As DrawingDirectionEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.DrawingDirection
    DrawingDirection = mvarDrawingDirection
End Property




Public Property Let FontSizeDefault(ByVal vData As Long)
Attribute FontSizeDefault.VB_Description = "Ĭ������ߴ硣һ��Ϊ9��С�壩��"
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
Attribute SpaceFirstLine.VB_Description = "���������ַ�����Ĭ��Ϊ0��"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.SpaceFirstLine = 5
    mvarSpaceFirstLine = vData
End Property


Public Property Get SpaceFirstLine() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.SpaceFirstLine
    SpaceFirstLine = mvarSpaceFirstLine
End Property



Public Property Let SpaceLine(ByVal vData As Long)
Attribute SpaceLine.VB_Description = "�м�ࡣĬ��Ϊ30�����Ϊ���������ʾ��ȷ�м�ࡣ"
Attribute SpaceLine.VB_UserMemId = 0
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.SpaceLine = 5
    mvarSpaceLine = vData
End Property


Public Property Get SpaceLine() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.SpaceLine
    SpaceLine = mvarSpaceLine
End Property



Public Property Let SpaceAfter(ByVal vData As Long)
Attribute SpaceAfter.VB_Description = "�κ��ࡣĬ��Ϊ0�����Ϊ���������ʾ��ȷ��ࡣ"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.SpaceAfter = 5
    mvarSpaceAfter = vData
End Property


Public Property Get SpaceAfter() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.SpaceAfter
    SpaceAfter = mvarSpaceAfter
End Property



Public Property Let SpaceBefore(ByVal vData As Long)
Attribute SpaceBefore.VB_Description = "��ǰ��ࡣĬ��Ϊ0�����Ϊ���������ʾ��ȷ��ࡣ"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.SpaceBefore = 5
    mvarSpaceBefore = vData
End Property


Public Property Get SpaceBefore() As Long
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.SpaceBefore
    SpaceBefore = mvarSpaceBefore
End Property




Public Property Let PageFooterAlign(ByVal vData As AlignTypeEnum)
Attribute PageFooterAlign.VB_Description = "ҳ���ı����뷽ʽ��"
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.PageFooterAlign = Form1
    mvarPageFooterAlign = vData
End Property
Public Property Get PageFooterAlign() As AlignTypeEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.PageFooterAlign
    PageFooterAlign = mvarPageFooterAlign
End Property

Public Property Let PageHeadAlign(ByVal vData As AlignTypeEnum)
Attribute PageHeadAlign.VB_Description = "ҳü�ı����뷽ʽ��"
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.PageHeadAlign = Form1
    mvarPageHeadAlign = vData
End Property
Public Property Get PageHeadAlign() As AlignTypeEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.PageHeadAlign
    PageHeadAlign = mvarPageHeadAlign
End Property

Public Property Let PageNumberAlign(ByVal vData As AlignTypeEnum)
Attribute PageNumberAlign.VB_Description = "ҳ���ı����뷽ʽ��"
'������ָ�ɶ���ʱʹ�ã�λ�� Set ������ߡ�
'Syntax: Set x.PageNumberAlign = Form1
    mvarPageNumberAlign = vData
End Property


Public Property Get PageNumberAlign() As AlignTypeEnum
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.PageNumberAlign
    PageNumberAlign = mvarPageNumberAlign
End Property



Public Property Let PageNumber(ByVal vData As Boolean)
Attribute PageNumber.VB_Description = "�Ƿ���ʾҳ�롣Ĭ��Ϊ��"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.PageNumber = 5
    mvarPageNumber = vData
End Property


Public Property Get PageNumber() As Boolean
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.PageNumber
    PageNumber = mvarPageNumber
End Property



Public Property Let PageFooter(ByVal vData As String)
Attribute PageFooter.VB_Description = "ҳ���ı���"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.PageFooter = 5
    mvarPageFooter = vData
End Property


Public Property Get PageFooter() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.PageFooter
    PageFooter = mvarPageFooter
End Property



Public Property Let PageHead(ByVal vData As String)
Attribute PageHead.VB_Description = "ҳü�ı���"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
'Syntax: X.PageHead = 5
    mvarPageHead = vData
End Property


Public Property Get PageHead() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
'Syntax: Debug.Print X.PageHead
    PageHead = mvarPageHead
End Property




Public Function LastTable() As cRTFTable
Attribute LastTable.VB_Description = "��ȡ�ĵ���������ӵı������"
    Set LastTable = Tables(Tables.Count)
End Function

Public Function LastText() As cRTFText
Attribute LastText.VB_Description = "��ȡ�ĵ���������ӵ��ı�����"
    Set LastText = Texts(Texts.Count)
End Function

Public Function LastPicture() As cRTFPicture
Attribute LastPicture.VB_Description = "��ȡ�ĵ���������ӵ�ͼƬ����"
    Set LastPicture = Pictures(Pictures.Count)
End Function

Public Function LastParagraph() As cRTFParagraph
    Set LastParagraph = Paragraphs(Paragraphs.Count)
End Function

Private Function GetImageSize(sFileName As String) As ImageSizeInf
   On Error Resume Next
   Dim bTemp(3) As Byte, lPos As Long, lFlen As Long
   Open sFileName For Binary As #1
       lFlen = LOF(1)
       Get #1, 1, bTemp()
       
       If bTemp(0) = &H89 And bTemp(1) = &H50 And bTemp(2) = &H4E And bTemp(3) = &H47 Or bTemp(0) = &H42 And bTemp(1) = &H4D Then
'       Debug.Print "\PNG OR BMP\"
           Get #1, 19, bTemp
           GetImageSize.Width = Byte2Long(bTemp(0), bTemp(1))
           Get #1, 23, bTemp
           GetImageSize.Height = Byte2Long(bTemp(0), bTemp(1))
       End If
       
       'JPG
       If bTemp(0) = &HFF And bTemp(1) = &HD8 And bTemp(2) = &HFF Then
'           Debug.Print "\JPEG\"
           lPos = 4
           Do
               Do
                   Get #1, lPos, bTemp
                   lPos = lPos + 1
               Loop Until (bTemp(0) = &HFF And bTemp(1) <> &HFF) Or lPos > lFlen
           
               Get #1, lPos, bTemp
                   
               If bTemp(0) >= &HC0 And bTemp(0) <= &HC3 Then
                   Get #1, lPos + 4, bTemp
                   Exit Do
               Else
                   lPos = lPos + (Byte2Long(bTemp(2), bTemp(1))) + 1
               End If
           Loop While lPos < lFlen
           GetImageSize.Width = Byte2Long(bTemp(3), bTemp(2))
           GetImageSize.Height = Byte2Long(bTemp(1), bTemp(0))
       End If

       'GIF file
       If bTemp(0) = &H47 And bTemp(1) = &H49 And bTemp(2) = &H46 And bTemp(3) = &H38 Then
'           Debug.Print "\GIF\"
           Get #1, 7, bTemp
           GetImageSize.Width = Byte2Long(bTemp(0), bTemp(1))
           GetImageSize.Height = Byte2Long(bTemp(2), bTemp(3))
       End If
           
       'PSD
       If bTemp(0) = &H38 And bTemp(1) = &H42 And bTemp(2) = &H50 And bTemp(3) = &H53 Then
'           Debug.Print "\PSD\"
           Get #1, 17, bTemp
           GetImageSize.Width = Byte2Long(bTemp(1), bTemp(0))
           Get #1, 21, bTemp
           GetImageSize.Height = Byte2Long(bTemp(1), bTemp(0))
       End If
       
       'TIF
       If bTemp(0) = &H4D And bTemp(1) = &H4D And bTemp(2) = &H0 And bTemp(3) = &H2A Then
'           Debug.Print "\TIF1\"
           Get #1, 31, bTemp
           GetImageSize.Width = Byte2Long(bTemp(1), bTemp(0))
           Get #1, 43, bTemp
           GetImageSize.Height = Byte2Long(bTemp(1), bTemp(0))
       End If
       
       If bTemp(0) = &H49 And bTemp(1) = &H49 And bTemp(2) = &H2A And bTemp(3) = &H0 Then
           Get #1, 5, bTemp
           If bTemp(0) = &H8 And bTemp(1) = &H0 And bTemp(2) = &H0 And bTemp(3) = &H0 Then
               'TIF
'               Debug.Print "\TIF2-1\"
               Get #1, 31, bTemp
               GetImageSize.Width = Byte2Long(bTemp(0), bTemp(1))
               Get #1, 43, bTemp
               GetImageSize.Height = Byte2Long(bTemp(0), bTemp(1))
           Else
               'TIF
'               Debug.Print "\TIF2-2\"
               lPos = Byte2Long(bTemp(0), bTemp(1)) + Byte2Long(bTemp(2), bTemp(3)) * 65536 + 11
               Get #1, lPos, bTemp
               GetImageSize.Width = Byte2Long(bTemp(0), bTemp(1))
               Get #1, lPos + 12, bTemp
               GetImageSize.Height = Byte2Long(bTemp(0), bTemp(1))
           End If
       End If

   Close #1
End Function

Private Function Byte2Long(ByVal lsb As Long, ByVal msb As Long) As Long
   Byte2Long = lsb + (msb * 256)
End Function


Public Property Let TextRTF(ByVal vData As String)
Attribute TextRTF.VB_Description = "�ĵ���RTF�ı��ַ�����"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
    mvarTextRTF = vData
End Property

Public Property Get TextRTF() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
    TextRTF = mvarTextRTF
End Property

Public Property Let FileName(ByVal vData As String)
Attribute FileName.VB_Description = "�ĵ�Ŀ���ļ�����"
'������ָ��ֵʱʹ�ã�λ�ڸ�ֵ������ߡ�
    mvarFileName = vData
End Property

Public Property Get FileName() As String
'��������ֵʱʹ�ã�λ�ڸ�ֵ�����ұߡ�
    FileName = mvarFileName
End Property

Private Sub Class_Initialize()
'���ʼ������
    'Ĭ��ģ�����
    mRTFHead = LoadResString(101)
    PageNumber = True
    PageNumberAlign = alm3_���Ҷ���
    PageHeadAlign = alm2_���ж���
    PageFooterAlign = alm2_���ж���
    
    Set Nodes = New cRTFNodes
    Set Texts = New cRTFTexts
    Set Pictures = New cRTFPictures
    Set Tables = New cRTFTables
    Set Paragraphs = New cRTFParagraphs
    
End Sub

Private Sub Class_Terminate()
'���봦�������¼�
    On Error Resume Next
    Set Texts = Nothing
    Set Pictures = Nothing
    Set Texts = Nothing
    Set Paragraphs = Nothing
End Sub

Public Function Save(Optional strFileName As String) As Boolean
Attribute Save.VB_Description = "�����ĵ�Ϊ�����ļ���"
On Error GoTo LL
'��֯����
    If Trim(strFileName) <> "" Then mvarFileName = strFileName
    Dim i As Long, j As Long
    For i = 1 To Nodes.Count
        Select Case Nodes(i).NodeType
        Case "�ı�"
            '�Զ��������������д���
            If mvarSpaceAfter <> 0 Then
                If Texts(Nodes(i).SubID).SpaceAfter = 0 Then Texts(Nodes(i).SubID).SpaceAfter = mvarSpaceAfter
            End If
            If mvarSpaceBefore <> 0 Then
                If Texts(Nodes(i).SubID).SpaceBefore = 0 Then Texts(Nodes(i).SubID).SpaceBefore = mvarSpaceBefore
            End If
            If mvarSpaceLine <> 0 Then
                If Texts(Nodes(i).SubID).SpaceLine = 0 Then Texts(Nodes(i).SubID).SpaceLine = mvarSpaceLine
            End If
            If mvarSpaceFirstLine <> 0 Then
                If Texts(Nodes(i).SubID).SpaceFirstLine = 0 Then Texts(Nodes(i).SubID).SpaceFirstLine = mvarSpaceFirstLine
            End If
            If mvarFontSizeDefault <> 0 Then
                If Texts(Nodes(i).SubID).FontSizeDefault = 0 Then Texts(Nodes(i).SubID).FontSizeDefault = mvarFontSizeDefault
            End If
            mRTFBody = mRTFBody & Texts(Nodes(i).SubID).GetTextRTF
        Case "ͼƬ"
            mRTFBody = mRTFBody & Pictures(Nodes(i).SubID).GetTextRTF
        Case "����"
            mRTFBody = mRTFBody & Tables(Nodes(i).SubID).GetTextRTF
        Case "����"
            '�Զ��������������д���
            If mvarSpaceAfter <> 0 Then
                If Paragraphs(Nodes(i).SubID).SpaceAfter = 0 Then Paragraphs(Nodes(i).SubID).SpaceAfter = mvarSpaceAfter
            End If
            If mvarSpaceBefore <> 0 Then
                If Paragraphs(Nodes(i).SubID).SpaceBefore = 0 Then Paragraphs(Nodes(i).SubID).SpaceBefore = mvarSpaceBefore
            End If
            If mvarSpaceLine <> 0 Then
                If Paragraphs(Nodes(i).SubID).SpaceLine = 0 Then Paragraphs(Nodes(i).SubID).SpaceLine = mvarSpaceLine
            End If
            If mvarSpaceFirstLine <> 0 Then
                If Paragraphs(Nodes(i).SubID).SpaceFirstLine = 0 Then Paragraphs(Nodes(i).SubID).SpaceFirstLine = mvarSpaceFirstLine
            End If
            If mvarFontSizeDefault <> 0 Then
                If Paragraphs(Nodes(i).SubID).FontSizeDefault = 0 Then Paragraphs(Nodes(i).SubID).FontSizeDefault = mvarFontSizeDefault
            End If
            mRTFBody = mRTFBody & Paragraphs(Nodes(i).SubID).GetTextRTF
        Case Else
            '
        End Select
    Next
    'ҳüҳ���滻
    Dim strPgHd As String
    Dim strAlignment As String
    Select Case mvarPageNumberAlign
    Case alm1_��˶���
        strAlignment = "\posxl"
    Case alm2_���ж���
        strAlignment = "\posxc"
    Case Else
        strAlignment = "\posxr"
    End Select
    
    Dim strAlignTMP As String
    'ҳü���뷽ʽ��
    Select Case mvarPageHeadAlign
    Case alm1_��˶���
        strAlignTMP = "\ql"
    Case alm3_���Ҷ���
        strAlignTMP = "\qr"
    Case alm4_���˶���
        strAlignTMP = "\qj"
    Case alm5_��ɢ����
        strAlignTMP = "\qd"
    Case Else   'Ĭ�Ͼ���
        strAlignTMP = "\qc"
    End Select

    If mvarPageHead <> "" Then
        strPgHd = "{\header \pard\plain \s15" & strAlignTMP & " \li0\ri0\nowidctlpar\brdrb\brdrs\brdrw15\brsp20 " & _
        "\tqc\tx4153\tqr\tx8306\aspalpha\aspnum\faauto\nosnaplinegrid\adjustright\rin0\lin0\itap0 \fs18\lang1033\langfe2052\kerning2\loch\af0\hich\af0\dbch\af13\cgrid\langnp1033\langfenp2052 {\insrsid13047022 \loch\af0\hich\af0\dbch\f13 " & StrToASC(mvarPageHead) & "}{" & _
        "\insrsid13047022\par}}"
    End If
    If mvarPageFooter <> "" Then
        strPgHd = strPgHd & "{\footer \pard\plain \s16\ql \li0\ri0\nowidctlpar\tqc\tx4153\tqr\tx8306\pvpara\phmrg" & strAlignment & "\posy0\aspalpha\aspnum\faauto\nosnaplinegrid\adjustright\rin0\lin0\itap0"
        If mvarPageNumber Then
            strPgHd = strPgHd & "\pararsid2622169 \fs18\lang1033\langfe2052\kerning2\loch\af0\hich\af0\dbch\af13\cgrid\langnp1033\langfenp2052 -\chpgn -{\cs17\insrsid13047022 \par }"
        End If
        'ҳ�Ŷ��뷽ʽ��
        Select Case mvarPageFooterAlign
        Case alm1_��˶���
            strAlignTMP = "\ql"
        Case alm3_���Ҷ���
            strAlignTMP = "\qr"
        Case alm4_���˶���
            strAlignTMP = "\qj"
        Case alm5_��ɢ����
            strAlignTMP = "\qd"
        Case Else   'Ĭ�Ͼ���
            strAlignTMP = "\qc"
        End Select
        strPgHd = strPgHd & "\pard \s16" & strAlignTMP & " \li0\ri360\nowidctlpar\tqc\tx4153\tqr\tx8306\aspalpha\aspnum\faauto\nosnaplinegrid\adjustright\rin360\lin0\itap0\pararsid13047022 {\insrsid13047022 \loch\af0\hich\af0\dbch\f13 " & StrToASC(mvarPageFooter) & "}{\insrsid13047022 \par }}"
    End If
    
    'ҳüҳ�Ŵ���
    mRTFHead = Replace(mRTFHead, "<ҳüҳ��>", strPgHd)
    '�ĵ����Դ���
    Dim strDoc As String
    If mvarPaperWidth <= 0 And mvarPaperHeight <= 0 Then
        'û�ж���ҳ�������߶ȣ�����PaperSize�Ƿ���ڣ�����ΪĬ��ֵ��
        '�߼���һӢ��ԼΪ 1440 羣�һ����Ϊ 567 � ( ����ʾʱ����Ļ��Ŀ�ĳ�������Ӣ������ײ����� )��
        Select Case mvarPaperSize
        Case 1
            strDoc = "\paperw" & CLng(21.59 * 567) & "\paperh" & CLng(27.94 * 567)
        Case 2
            strDoc = "\paperw" & CLng(27.94 * 567) & "\paperh" & CLng(43.17 * 567)
        Case 3
            strDoc = "\paperw" & CLng(21.59 * 567) & "\paperh" & CLng(35.56 * 567)
        Case 4
            strDoc = "\paperw" & CLng(29.7 * 567) & "\paperh" & CLng(42 * 567)
        Case 5
            strDoc = "\paperw" & CLng(21 * 567) & "\paperh" & CLng(29.7 * 567)
        Case 6
            strDoc = "\paperw" & CLng(14.8 * 567) & "\paperh" & CLng(21 * 567)
        Case 7
            strDoc = "\paperw" & CLng(25.7 * 567) & "\paperh" & CLng(36.4 * 567)
        Case 8
            strDoc = "\paperw" & CLng(18.2 * 567) & "\paperh" & CLng(25.7 * 567)
        Case Else
            strDoc = "\paperw11906\paperh16838" 'Ĭ��ֵ
        End Select
    Else
        If mvarPaperWidth > 0 Then
            strDoc = "\paperw" & mvarPaperWidth
        Else    'Ĭ�Ͽ���11906
            strDoc = "\paperw11906"
        End If
        If mvarPaperHeight > 0 Then
            strDoc = strDoc & "\paperh" & mvarPaperHeight
        Else    'Ĭ�Ͽ���16838
            strDoc = strDoc & "\paperh16838"
        End If
    End If
    
    If mvarMarginLeft > 0 Then
        strDoc = strDoc & "\margl" & mvarMarginLeft
    Else    'Ĭ��1800
        strDoc = strDoc & "\margl1800"
    End If
    If mvarMarginRight > 0 Then
        strDoc = strDoc & "\margr" & mvarMarginRight
    Else    'Ĭ��1800
        strDoc = strDoc & "\margr1800"
    End If
    If mvarMarginTop > 0 Then
        strDoc = strDoc & "\margt" & mvarMarginTop
    Else    'Ĭ��1440
        strDoc = strDoc & "\margt1440"
    End If
    If mvarMarginBottom > 0 Then
        strDoc = strDoc & "\margb" & mvarMarginBottom
    Else    'Ĭ��1440
        strDoc = strDoc & "\margb1440"
    End If
    If mvarMarginGutter > 0 Then
        strDoc = strDoc & "\gutter" & mvarMarginGutter
    Else    'Ĭ��0
        strDoc = strDoc & "\gutter0"
    End If
    '�������
    If mvarCaption <> "" Then strDoc = strDoc & "{\windowcaption" & StrToASC(mvarCaption) & "}"
    
    If mvarDefaultTabWidth > 0 Then
        strDoc = strDoc & "\deftab" & mvarDefaultTabWidth
    Else
        strDoc = strDoc & "\deftab420"
    End If
    strDoc = strDoc & "\ftnbj\aenddoc\hyphcaps0\formshade"
    
    If mvarDrawingDirection = dwd1_��ֱ��ͼ Then
        strDoc = strDoc & "\vertsect"
    Else
        strDoc = strDoc & "\horzsect"
    End If
    strDoc = strDoc & "\dgmargin\dghspace180\dgvspace156\dghorigin1797\dgvorigin1440\dghshow1\dgvshow2\jcompress\lnongrid"
    
    If mvarViewKind > 0 Then
        strDoc = strDoc & "\viewkind" & mvarViewKind
    Else
        strDoc = strDoc & "\viewkind1"
    End If
    If mvarViewScale > 0 Then
        strDoc = strDoc & "\viewscale" & mvarViewScale
    Else
        strDoc = strDoc & "\viewscale100"
    End If
    strDoc = strDoc & "\pgbrdrhead\pgbrdrfoot\pgbrdrsnap\splytwnine\ftnlytwnine\htmautsp\useltbaln\alntblind\lytcalctblwd\lyttblrtgr\lnbrkrule\nobrkwrptbl\snaptogridincell\allowfieldendsel\wrppunct\asianbrkrule\rsidroot10699452\newtblstyruls\nogrowautofit "
    
    '�滻�ĵ������ı�
    mRTFHead = Replace(mRTFHead, "<�ĵ�����>", strDoc)
    
    '�ڸ�ʽ�������߿�����
    Dim strSect As String
    Dim strBorderStyle As String, strBorderWidth As String, strBorderColor As String
    If mvarPageBorderArt > 0 Then   '�����߿�1��165
        If mvarPageBorderArt <= 165 Then
            strBorderStyle = "\brdrart" & mvarPageBorderArt
        Else
            strBorderStyle = "\brdrnone"
        End If
    Else
        Select Case mvarPageBorderStyle
        Case cbs01_�ޱ߿�
            strBorderStyle = "\brdrnone"
        Case cbs02_������ȱ߿�
            strBorderStyle = "\brdrs"
        Case cbs03_˫����ȱ߿�
            strBorderStyle = "\brdrth"
        Case cbs04_��Ӱ�߿�
            strBorderStyle = "\brdrsh"
        Case cbs05_˫�߿�
            strBorderStyle = "\brdrdb"
        Case cbs06_���߱߿�
            strBorderStyle = "\brdrdot"
        Case cbs07_���߱߿�
            strBorderStyle = "\brdrdash"
        Case cbs08_ϸ�߱߿�
            strBorderStyle = "\brdrhair"
        Case cbs09_����ʽ�߿�
            strBorderStyle = "\brdrinset"
        Case cbs10_����ϸ�߿�
            strBorderStyle = "\brdrdashsm"
        Case cbs11_�㻮�߱߿�
            strBorderStyle = "\brdrdashd"
        Case cbs12_��㻮�߱߿�
            strBorderStyle = "\brdrdashdd"
        Case cbs13_��ʼ�߿�
            strBorderStyle = "\brdroutset"
        Case cbs14_���ر߿�
            strBorderStyle = "\brdrtriple"
        Case cbs15_�񱡵���ϸ�߿�
            strBorderStyle = "\brdrtnthsg"
        Case cbs16_�������ϸ�߿�
            strBorderStyle = "\brdrthtnsg"
        Case cbs17_���񱡵���ϸ�߿�
            strBorderStyle = "\brdrtnthtnsg"
        Case cbs18_�񱡵����б߿�
            strBorderStyle = "\brdrtnthmg"
        Case cbs19_��������б߿�
            strBorderStyle = "\brdrthtnmg"
        Case cbs20_���񱡵����б߿�
            strBorderStyle = "\brdrtnthtnmg"
        Case cbs21_�񱡵��Ӵֱ߿�
            strBorderStyle = "\brdrtnthlg"
        Case cbs22_������Ӵֱ߿�
            strBorderStyle = "\brdrthtnlg"
        Case cbs23_���񱡵��Ӵֱ߿�
            strBorderStyle = "\brdrtnthtnlg"
        Case cbs24_�����߱߿�
            strBorderStyle = "\brdrwavy"
        Case cbs25_˫�����߱߿�
            strBorderStyle = "\brdrwavydb"
        Case cbs26_���Ʊ߿�
            strBorderStyle = "\brdrdashdotstr"
        Case cbs27_����߿�
            strBorderStyle = "\brdremboss"
        Case cbs28_��̱߿�
            strBorderStyle = "\brdrengrave"
        Case Else   'Ĭ�ϵ��߿�
            strBorderStyle = "\brdrs"
        End Select
    End If
        
    '�߿����
    If mvarPageBorderWidth <> 0 Then
        strBorderWidth = "\brdrw" & mvarPageBorderWidth
    Else    'Ĭ�Ͽ���15
        strBorderWidth = "\brdrw15"
    End If
    
    If mvarPageBorderColor <> clr00_Ĭ�� Then
        strBorderColor = "\brdrcf" & mvarPageBorderColor
    End If
    
    
    strSect = "\fet0\sectd \linex0\headery851\footery992\colsx425\endnhere"
    
    Dim strPgBrdSpc As String
    If mvarPageBorderSpace > 0 Then
        strPgBrdSpc = "\brsp" & mvarPageBorderSpace
    Else
        strPgBrdSpc = "\brsp480"
    End If
    '���ɵ�Ԫ��߿������
    If mvarPageBorderArt > 0 Then   '�����ͱ߿�Ŀ���Ĭ��Ϊ50�����Ϊ480����ɫΪ��
        strSect = strSect & "\pgbrdrt" & strBorderStyle & "\brdrw50" & strPgBrdSpc & vbCrLf
        strSect = strSect & "\pgbrdrb" & strBorderStyle & "\brdrw50" & strPgBrdSpc & vbCrLf
        strSect = strSect & "\pgbrdrl" & strBorderStyle & "\brdrw50" & strPgBrdSpc & vbCrLf
        strSect = strSect & "\pgbrdrr" & strBorderStyle & "\brdrw50" & strPgBrdSpc & vbCrLf
    ElseIf mvarPageBorderStyle <> cbs00_Ĭ�ϱ߿���ʽ Then
        strSect = strSect & "\pgbrdrt" & strBorderStyle & strBorderWidth & strBorderColor & strPgBrdSpc & vbCrLf
        strSect = strSect & "\pgbrdrb" & strBorderStyle & strBorderWidth & strBorderColor & strPgBrdSpc & vbCrLf
        strSect = strSect & "\pgbrdrl" & strBorderStyle & strBorderWidth & strBorderColor & strPgBrdSpc & vbCrLf
        strSect = strSect & "\pgbrdrr" & strBorderStyle & strBorderWidth & strBorderColor & strPgBrdSpc & vbCrLf
    End If
    strSect = strSect & "\sectlinegrid312\sectspecifyl\sectrsid11756820\sftnbj "
    '�滻�ڸ�ʽ�����ı�
    mRTFHead = Replace(mRTFHead, "<������>", strSect)
    
    
    TextRTF = LinkRTF(mRTFHead, mRTFBody, "}")
    
    If Me.FileName = "" Then
        Save = False
        Exit Function
    End If
    Const ForReading = 1, ForWriting = 2, ForAppending = 3
    Dim fs As FileSystemObject, f As TextStream
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.OpenTextFile(FileName, ForWriting, TristateUseDefault)
    f.Write Me.TextRTF
    f.Close
    Save = True
    Exit Function
LL:
    Save = False
End Function

Public Function AddText(Optional Text As String, Optional sKey As String, _
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
Attribute AddText.VB_Description = "���ĵ�������һ���ı���"
    Dim i As Long

    i = Texts.Add(Text, sKey, Alignment, Bold, Italic, ForeColor, FontSize, FontStyle, CharStyle, _
        AnimType, ScaleX, Embo, Expand, OutLine, Shadow, Strike, StrikeD, _
        SubChar, SuperChar, UnderLineType, UnderLineColor, Hided, WebHidden, CharBorder, _
        CharShadow, CharPatForeColor, CharPatBackColor, PatType, Highlight, HighlightColor, ListType, IsInTable)
    AddText = i
    Nodes.Add i, "�ı�"
End Function

Public Function AddParagraph(Optional sKey As String, Optional IsInTable As Boolean = False, _
    Optional SpaceBefore As Long = 0, Optional SpaceAfter As Long = 0, _
    Optional SpaceLine As Long = 0, Optional SpaceFirstLine As Long = 0, _
    Optional FontSizeDefault As Long = 0) As Long
    Dim i As Long

    i = Paragraphs.Add(sKey, IsInTable, SpaceBefore, SpaceAfter, SpaceLine, SpaceFirstLine, FontSizeDefault)
    
    AddParagraph = i
    Nodes.Add i, "����"
End Function


Public Function AddPicture(Optional FileName As String, _
Optional Width As Long = 0, _
Optional Height As Long = 0, _
Optional Alignment As AlignTypeEnum = alm0_Ĭ�϶���, _
Optional sKey As String, _
Optional ScaleX As Long = 100, _
Optional ScaleY As Long = 100, _
Optional CropLeft As Long = 0, _
Optional CropRight As Long = 0, _
Optional CropTop As Long = 0, _
Optional CropBottom As Long = 0, _
Optional WidthOld As Long = 100, _
Optional HeightOld As Long = 100, _
Optional PicType As PicTypeEnum = pic_JPG, _
Optional AddCRLF As Boolean = True) As Long
Attribute AddPicture.VB_Description = "���ĵ�������һ��ͼƬ��"
    Dim i As Long
    Dim picSize As ImageSizeInf
    Dim lngWidth As Long, lngHeight As Long
    picSize = GetImageSize(FileName)
    lngWidth = picSize.Width * Screen.TwipsPerPixelX
    lngHeight = picSize.Height * Screen.TwipsPerPixelY
    If Width = 0 Then Width = lngWidth
    If Height = 0 Then Height = lngHeight
    
    i = Pictures.Add(FileName, Width, Height, Alignment, sKey, ScaleX, ScaleY, CropLeft, CropRight, _
        CropTop, CropBottom, lngWidth, lngHeight, PicType, AddCRLF)
    AddPicture = i
    Nodes.Add i, "ͼƬ"
End Function


Public Function AddTable(Optional RowCount As Long = 1, Optional ColCount As Long = 1, _
Optional Alignment As RowAlignTypeEnum = ral0_Ĭ�϶���, _
Optional sKey As String) As Long
Attribute AddTable.VB_Description = "���ĵ�������һ������"
    Dim i As Long

    i = Tables.Add(RowCount, ColCount, Alignment, sKey)

    AddTable = i
    Nodes.Add i, "����"
End Function
