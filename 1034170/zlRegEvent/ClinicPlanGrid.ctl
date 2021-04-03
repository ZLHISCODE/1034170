VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.UserControl ClinicPlanGrid 
   BackColor       =   &H80000005&
   ClientHeight    =   6885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11445
   ScaleHeight     =   6885
   ScaleWidth      =   11445
   Begin VSFlex8Ctl.VSFlexGrid vsfRegistPlan 
      Height          =   4305
      Left            =   630
      TabIndex        =   0
      Top             =   240
      Width           =   6555
      _cx             =   11562
      _cy             =   7594
      Appearance      =   0
      BorderStyle     =   0
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483638
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"ClinicPlanGrid.ctx":0000
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
      Begin VB.PictureBox Picture1 
         Height          =   105
         Left            =   5400
         ScaleHeight     =   105
         ScaleWidth      =   30
         TabIndex        =   5
         Top             =   2160
         Width           =   30
      End
      Begin VB.PictureBox picImgPlan 
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   30
         Picture         =   "ClinicPlanGrid.ctx":0075
         ScaleHeight     =   135
         ScaleWidth      =   150
         TabIndex        =   4
         Top             =   90
         Width           =   150
      End
      Begin VB.PictureBox picToolTip 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   570
         Left            =   630
         ScaleHeight     =   540
         ScaleWidth      =   1365
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1110
         Visible         =   0   'False
         Width           =   1395
         Begin VB.Label lblToolTip1 
            BackStyle       =   0  'Transparent
            Caption         =   "�ѹ�����120"
            Height          =   180
            Left            =   90
            TabIndex        =   3
            Top             =   60
            Width           =   1290
         End
         Begin VB.Label lblToolTip2 
            BackStyle       =   0  'Transparent
            Caption         =   "�޺�����220"
            Height          =   180
            Left            =   90
            TabIndex        =   2
            Top             =   300
            Width           =   1290
         End
      End
   End
   Begin VB.Image imgReplace 
      Height          =   120
      Left            =   9270
      Picture         =   "ClinicPlanGrid.ctx":016B
      Top             =   2340
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H8000000C&
      Height          =   6435
      Left            =   480
      Top             =   90
      Width           =   10425
   End
   Begin VB.Image imgLock 
      Height          =   210
      Left            =   8880
      Picture         =   "ClinicPlanGrid.ctx":0615
      Top             =   2310
      Visible         =   0   'False
      Width           =   150
   End
End
Attribute VB_Name = "ClinicPlanGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowRgn Lib "user32" (ByVal Hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

Private Enum mPlanGridFixedColIndex '����̶�������
    COl_���� = 0
    COL_��ԴID
    COL_����ID
    COL_����
    col_����
    COL_��Ŀ
    COL_����
    Col_ҽ��
    
    COL_��ʼʱ��
    COL_��ֹʱ��
    
    COL_�Ƿ񽨲���
    COL_�Ƿ���ſ���
    COL_�Ƿ��ʱ��
    COL_ԤԼ����
    COL_����Ƶ��
    COL_���տ���״̬
    COL_�Ű෽ʽ
    COL_���﷽ʽ
End Enum

Public Enum gDataStyle
    Data_Templet = 0
    Data_FixedRule = 1
    Data_Plan = 2
End Enum

Private Const m_def_DataStyle = 0

Private m_DataStyle As gDataStyle
Private m_MinDate As Date
Private m_MaxDate As Date

'�¼�����
Public Event AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private mrsData As ADODB.Recordset

Public Property Get vsfGrid() As VSFlexGrid
    Set vsfGrid = vsfRegistPlan
End Property

Public Property Get Hwnd() As Long
    Hwnd = UserControl.Hwnd
End Property

Public Property Get DataStyle() As gDataStyle
Attribute DataStyle.VB_Description = "����/���ñ�����͡�"
    DataStyle = m_DataStyle
End Property

Public Property Let DataStyle(ByVal New_DataStyle As gDataStyle)
    m_DataStyle = New_DataStyle
    PropertyChanged "DataStyle"
    Call InitPlanGrid
End Property

Private Sub UserControl_Initialize()
    Call InitPlanGrid
    Call SetPicToolTipEffect
End Sub

Private Sub UserControl_InitProperties()
    m_DataStyle = m_def_DataStyle
    m_MinDate = Format(Now, "yyyy-mm-dd")
    m_MaxDate = Format(Now, "yyyy-mm-dd")
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_DataStyle = PropBag.ReadProperty("DataStyle", m_def_DataStyle)
    m_MinDate = PropBag.ReadProperty("MinDate", Format(Now, "yyyy-mm-dd"))
    m_MaxDate = PropBag.ReadProperty("MaxDate", Format(Now, "yyyy-mm-dd"))
    Call InitPlanGrid
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    shpBorder.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    vsfRegistPlan.Move 10, 10, UserControl.ScaleWidth - 20, UserControl.ScaleHeight - 20
End Sub

Private Sub UserControl_Terminate()
    Set mrsData = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("DataStyle", m_DataStyle, m_def_DataStyle)
    Call PropBag.WriteProperty("MinDate", m_MinDate, Format(Now, "yyyy-mm-dd"))
    Call PropBag.WriteProperty("MaxDate", m_MaxDate, Format(Now, "yyyy-mm-dd"))
End Sub

'��ʼʱ��
Public Property Get MinDate() As Date
    MinDate = m_MinDate
End Property

Public Property Let MinDate(ByVal vNewValue As Date)
    m_MinDate = Format(vNewValue, "yyyy-mm-dd")
    If m_MinDate > m_MaxDate Then m_MaxDate = m_MinDate
    PropertyChanged "MinDate"
    Call InitPlanGrid
End Property

'����ʱ��
Public Property Get MaxDate() As Date
    MaxDate = m_MaxDate
End Property

Public Property Let MaxDate(ByVal vNewValue As Date)
    Dim dtCur As Date
    dtCur = Format(vNewValue, "yyyy-mm-dd")
    If m_MinDate > dtCur Then Exit Property
    m_MaxDate = dtCur
    PropertyChanged "MaxDate"
    Call InitPlanGrid
End Property

'��ǰѡ��������
'------------------------------------------
Public Property Get ������Ŀ() As String
    Dim lngCol As Long
    
    If IsSelectedNotNull() = False Then Exit Property
    With vsfRegistPlan
        lngCol = GetItemNameCol(.Col, .FixedCols) 'ʱ����
        ������Ŀ = .Cell(flexcpData, 0, lngCol)
    End With
End Property

Public Property Get ʱ���() As String
    Dim lngCol As Long
    
    If IsSelectedNotNull() = False Then Exit Property
    With vsfRegistPlan
        lngCol = GetItemNameCol(.Col, .FixedCols) 'ʱ����
        ʱ��� = .TextMatrix(.Row, lngCol)
    End With
End Property

Public Property Get ��ԴID() As Long
    With vsfRegistPlan
        If .Visible = False Then Exit Property
        If .Row < .FixedRows Or .Col < .FixedCols Then Exit Property
        If .RowData(.Row) = -1 Then Exit Property '����
        ��ԴID = Val(.TextMatrix(.Row, COL_��ԴID))
    End With
End Property

Public Property Get ����ID() As Long
    With vsfRegistPlan
        If .Visible = False Then Exit Property
        If .Row < .FixedRows Or .Col < .FixedCols Then Exit Property
        If .RowData(.Row) = -1 Then Exit Property '����
        ����ID = Val(.TextMatrix(.Row, COL_����ID))
    End With
End Property

Public Property Get ��¼ID() As Long
    Dim lngCol As Long
    
    If IsSelectedNotNull() = False Then Exit Property
    With vsfRegistPlan
        lngCol = GetItemNameCol(.Col, .FixedCols) 'ʱ����
        ��¼ID = Val(.Cell(flexcpData, .Row, lngCol))
    End With
End Property

Public Property Get Is����() As Boolean
    Dim lngCol As Long
    
    If IsSelectedNotNull() = False Then Exit Property
    With vsfRegistPlan
        lngCol = GetItemNameCol(.Col, .FixedCols) 'ʱ����
        Is���� = Not .Cell(flexcpPicture, .Row, lngCol) Is Nothing
    End With
End Property

Public Property Get Isͣ��() As Boolean
    Dim lngCol As Long
    
    If IsSelectedNotNull() = False Then Exit Property
    With vsfRegistPlan
        lngCol = GetItemNameCol(.Col, .FixedCols) 'ʱ����
        Isͣ�� = .Cell(flexcpForeColor, .Row, lngCol) = vbRed
    End With
End Property

Public Property Get Is����() As Boolean
    Dim lngCol As Long
    
    If IsSelectedNotNull() = False Then Exit Property
    With vsfRegistPlan
        lngCol = GetItemNameCol(.Col, .FixedCols) 'ʱ����
        Is���� = .Cell(flexcpForeColor, .Row, lngCol) = vbBlue
    End With
End Property

Public Property Get IsSelectedNotNull() As Boolean
    Dim lngCol As Long
    
    '�жϵ�ǰѡ���Ƿ�Ϊ��
    Err = 0: On Error GoTo errHandler
    With vsfRegistPlan
        If .Visible = False Then Exit Property
        If .Row < .FixedRows Or .Col < .FixedCols Then Exit Property
        lngCol = GetItemNameCol(.Col, .FixedCols) 'ʱ����
        If .RowData(.Row) = -1 Or Trim(.TextMatrix(.Row, lngCol)) = "" Then Exit Property '���л���ʱ����
    End With
    IsSelectedNotNull = True
    Exit Property
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Property
'------------------------------------------

Private Sub SetPicToolTipEffect()
    '���ܣ�������ʾ��PicToolTip�ı߿�
    Dim lngR As Long
    
    '�߿�API=RoundRect
    picToolTip.Line (Screen.TwipsPerPixelX, 0)-(picToolTip.Width - Screen.TwipsPerPixelX, 0), RGB(118, 118, 118)
    picToolTip.Line (Screen.TwipsPerPixelX, picToolTip.Height - Screen.TwipsPerPixelY)-(picToolTip.Width - Screen.TwipsPerPixelX, picToolTip.Height - Screen.TwipsPerPixelY), RGB(118, 118, 118)
    picToolTip.Line (0, Screen.TwipsPerPixelY)-(0, picToolTip.Height - Screen.TwipsPerPixelY), RGB(118, 118, 118)
    picToolTip.Line (picToolTip.Width - Screen.TwipsPerPixelX, Screen.TwipsPerPixelY)-(picToolTip.Width - Screen.TwipsPerPixelX, picToolTip.Height - Screen.TwipsPerPixelY), RGB(118, 118, 118)
    picToolTip.PSet (Screen.TwipsPerPixelX, Screen.TwipsPerPixelY), RGB(186, 186, 186)
    picToolTip.PSet (picToolTip.Width - Screen.TwipsPerPixelX * 2, Screen.TwipsPerPixelY), RGB(186, 186, 186)
    picToolTip.PSet (Screen.TwipsPerPixelX, picToolTip.Height - Screen.TwipsPerPixelY * 2), RGB(186, 186, 186)
    picToolTip.PSet (picToolTip.Width - Screen.TwipsPerPixelX * 2, picToolTip.Height - Screen.TwipsPerPixelY * 2), RGB(186, 186, 186)
    
    '��״
    lngR = CreateRoundRectRgn(0, 0, picToolTip.ScaleX(picToolTip.Width, picToolTip.ScaleMode, vbPixels) + 1, picToolTip.ScaleY(picToolTip.Height, picToolTip.ScaleMode, vbPixels) + 1, 2, 2)
    Call SetWindowRgn(picToolTip.Hwnd, lngR, False)
End Sub

Private Sub InitPlanGrid()
    '���ܣ���ʼ���������ݱ��
    '   vsfGrid - VSF���
    Dim strHead As String, varData As Variant
    Dim strHeadSub As String, varDataSub As Variant
    Dim i As Long, lngCol As Long
    Dim arrDate As Variant
    Dim dtCurDate As Date, dtMaxDate As Date, intDays As Integer

    Err = 0: On Error GoTo errHandler
    With vsfRegistPlan
        .Redraw = False
        .Rows = 2
        
        '�̶���
        strHead = " ,4,200|��ԴID,4,0|����ID,4,0|����,4,500|����,4,500|��Ŀ,1,1000|����,1,1000|ҽ��,1,650"
        strHeadSub = " ,��ԴID,����ID,����,����,��Ŀ,����,ҽ��"
        If m_DataStyle = Data_FixedRule Then
            strHead = strHead & "|��ʼʱ��,4,1900|��ֹʱ��,4,1900"
            strHeadSub = strHeadSub & ",��ʼʱ��,��ֹʱ��"
        End If
        varData = Split(strHead, "|")
        varDataSub = Split(strHeadSub, ",")
        .Cols = UBound(varData) + 1
        For i = 0 To UBound(varData)
            .TextMatrix(0, i) = Split(varData(i), ",")(0): .TextMatrix(1, i) = varDataSub(i)
            .ColAlignment(i) = Split(varData(i), ",")(1)
            .ColWidth(i) = Split(varData(i), ",")(2)
            .ColKey(i) = Split(varData(i), ",")(0)
        Next
        .FixedCols = .Cols: .FixedRows = 2
        '��̬��
        Select Case m_DataStyle
        Case Data_Templet, Data_FixedRule 'ģ��,�̶�����
            strHead = "��һ,1,450|��һ,4,550|��һ,4,550|�ܶ�,1,450|�ܶ�,4,450|�ܶ�,4,550|����,1,450|����,4,550|����,4,550|" & _
                    "����,1,450|����,4,550|����,4,550|����,1,450|����,4,550|����,4,550|����,1,450|����,4,550|����,4,550|" & _
                    "����,1,450|����,4,550|����,4,550"
            strHeadSub = "ʱ��,�޺�,��Լ,ʱ��,�޺�,��Լ,ʱ��,�޺�,��Լ," & _
                    "ʱ��,�޺�,��Լ,ʱ��,�޺�,��Լ,ʱ��,�޺�,��Լ," & _
                    "ʱ��,�޺�,��Լ"
            If m_DataStyle = Data_Templet Then
                strHead = strHead & "|��������,1,1150|��������,1,450|��������,4,550|��������,4,550"
                strHeadSub = strHeadSub & ",������Ŀ,ʱ��,�޺�,��Լ"
            End If
            varData = Split(strHead, "|")
            varDataSub = Split(strHeadSub, ",")
            lngCol = .Cols
            .Cols = .Cols + UBound(varData) + 1
            For i = 0 To UBound(varData)
                .TextMatrix(0, lngCol) = Split(varData(i), ",")(0): .TextMatrix(1, lngCol) = varDataSub(i)
                .Cell(flexcpData, 0, lngCol) = CStr(Split(varData(i), ",")(0))
                .ColAlignment(lngCol) = Split(varData(i), ",")(1)
                .ColWidth(lngCol) = Split(varData(i), ",")(2)
'                .ColKey(i) = Split(varData(i), ",")(0) & "-" & varDataSub(i)
                If i Mod 3 <> 0 Then .FixedAlignment(i) = flexAlignCenterCenter
                lngCol = lngCol + 1
            Next
        Case Data_Plan '���ż�¼
            intDays = DateDiff("d", m_MinDate, m_MaxDate) + 1 '����
            dtCurDate = m_MinDate
            lngCol = .Cols
            .Cols = .Cols + intDays * 3
            For i = 1 To intDays
                .Cell(flexcpText, 0, lngCol, 0, lngCol + 2) = Format(dtCurDate, "mm��dd��") & Chr(10) & _
                    Choose(Weekday(dtCurDate, vbMonday), "��һ", "�ܶ�", "����", "����", "����", "����", "����")
                .Cell(flexcpData, 0, lngCol) = dtCurDate '����
                .Cell(flexcpText, 1, lngCol, 1, lngCol + 2) = "ʱ��" & vbTab & "�޺�" & vbTab & "��Լ"
                .ColAlignment(lngCol) = 1: .ColAlignment(lngCol + 1) = 4: .ColAlignment(lngCol + 2) = 4
                .ColWidth(lngCol) = 580: .ColWidth(lngCol + 1) = 650: .ColWidth(lngCol + 2) = 650
'                .ColKey(lngCol) = i & "-ʱ��": .ColKey(lngCol + 1) = i & "-�޺�": .ColKey(lngCol + 2) = i & "-��Լ"
                .FixedAlignment(lngCol + 1) = flexAlignCenterCenter: .FixedAlignment(lngCol + 2) = flexAlignCenterCenter
                dtCurDate = DateAdd("d", 1, dtCurDate)
                lngCol = lngCol + 3
            Next
            .RowHeight(0) = 500
        End Select
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .HighLight = flexHighlightNever
        .FocusRect = flexFocusNone
        .SelectionMode = flexSelectionFree
        .AllowUserResizing = flexResizeColumns
        .GridLines = flexGridFlat
'        .WordWrap = True '�����Զ�����
        .RowHeightMin = 450
        
        '����������,�����û�ѡ����ʾ��
        For i = 0 To .Cols - 1
            'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)|������(0-��������,1-��ֹ����,2-��������,�����س���������)
            Select Case i
            Case COl_����, COL_��ԴID, COL_����ID
                 .ColData(i) = "-1|1"
            Case col_����
                .ColData(i) = "1|0"
            End Select
        Next

        '�ϲ�����
        .MergeCellsFixed = flexMergeRestrictColumns
        .MergeRow(0) = True: .MergeCol(-1) = True
        .Redraw = True
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function LoadDataByRecordset(ByVal rsData As ADODB.Recordset) As Boolean
    '���ܣ�����Recordset�����������
    '˵�������ݱ����ǰ�"����,����,����,��Ŀ,ҽ��"���������˵ģ����������ʾ����ȷ
    Dim i As Long, j As Long, lngCurRow As Long, lngCurCol As Long
    Dim strGroupKey As String '��������"����,����,����,��Ŀ,ҽ��"����
    Dim lngBackColor As Long '����������Ľ���ɫ
    Dim strTemp As String, blnAddRow As Boolean
    Dim lngRowStart As Long, lngRowEnd As Long
    Dim lngOldRow As Long, lngoldCol As Long, blnFind As Boolean
    
    Err = 0: On Error GoTo errHandle
    vsfGrid.Rows = 2
    Set mrsData = Nothing
    If rsData Is Nothing Then Exit Function
    If rsData.RecordCount = 0 Then Exit Function
    Set mrsData = rsData
    
    With vsfRegistPlan
        lngOldRow = .Row
        lngoldCol = .Col
        lngCurRow = 2
        strGroupKey = ""
        lngBackColor = G_AlternateColor
        Do While Not rsData.EOF
            blnFind = False
            '1.�������
            strTemp = Nvl(rsData!����) & "," & Nvl(rsData!����) & "," & Nvl(rsData!����) & "," & _
                    Nvl(rsData!�շ���Ŀ) & "," & Nvl(rsData!ҽ������)
            If m_DataStyle = Data_FixedRule Then strTemp = strTemp & "," & Nvl(rsData!��ʼʱ��) & "," & Nvl(rsData!��ֹʱ��)
            If strGroupKey <> strTemp Then
                strGroupKey = strTemp
                lngCurCol = -1  '�����ж��Ƿ�ȷ������
                lngBackColor = IIf(lngBackColor = vbWindowBackground, G_AlternateColor, vbWindowBackground)
                
                .Rows = .Rows + 1: lngCurRow = .Rows - 1
                .RowHidden(lngCurRow) = True '��һ�п��У�������
                .RowData(lngCurRow) = -1 '��ǣ������ж��Ƿ�Ϊ���ؿ���
                
                lngCurRow = lngCurRow + 1
            End If
            '2.�������
            '2.1ȷ����ǰ��
            Select Case m_DataStyle
            Case Data_Templet 'ģ��
                If Nvl(rsData!�Ű����) <> 1 Then
                    '�Ű����:1-�����Ű�;2-�����Ű�;3-˫���Ű�;4-������ѭ;5-��ѭ������;6-�ض�����
                    lngCurCol = .Cols - 4: blnFind = True
                Else
                    strTemp = Nvl(rsData!������Ŀ, "��")
                    For i = .FixedCols To .Cols - 1 Step 3
                        If strTemp = .Cell(flexcpData, 0, i) Then
                            lngCurCol = i: blnFind = True
                            Exit For
                        End If
                    Next
                End If
            Case Data_FixedRule '�̶�����
                strTemp = Nvl(rsData!������Ŀ, "��")
                For i = .FixedCols To .Cols - 1 Step 3
                    If strTemp = .Cell(flexcpData, 0, i) Then
                        lngCurCol = i: blnFind = True
                        Exit For
                    End If
                Next
            Case Else '���ż�¼
                strTemp = Format(Nvl(rsData!��������), "yyyy-mm-dd")
                For i = .FixedCols To .Cols - 1 Step 3
                    If DateDiff("d", strTemp, .Cell(flexcpData, 0, i)) = 0 Then
                        lngCurCol = i: blnFind = True
                        Exit For
                    End If
                Next
            End Select
            If blnFind Then
                '2.2ȷ����ǰ��
                If m_DataStyle = Data_Templet Then
                    '�Ű����:1-�����Ű�;2-�����Ű�;3-˫���Ű�;4-������ѭ;5-��ѭ������;6-�ض�����
                    If Nvl(rsData!�Ű����) = 6 Then
                        Call GetGroupRange(IIf(.Rows - 1 > lngCurRow, lngCurRow, .Rows - 1), lngRowStart, lngRowEnd)
                        lngCurRow = lngRowEnd + 1
                        For i = lngRowStart To lngRowEnd
                            If .TextMatrix(i, lngCurCol + 1) = "" Then lngCurRow = i: Exit For
                            If .TextMatrix(i, lngCurCol + 1) = Nvl(rsData!�ϰ�ʱ��) Then
                                lngCurRow = i: Exit For
                            End If
                        Next
                    Else
                        For i = IIf(.Rows - 1 > lngCurRow, lngCurRow, .Rows - 1) To 2 Step -1
                            If .RowData(i) = -1 Or .TextMatrix(i, lngCurCol) <> "" Then  '�����ؿ��л�����������
                                lngCurRow = i + 1: Exit For
                            End If
                        Next
                    End If
                Else
                    For i = IIf(.Rows - 1 > lngCurRow, lngCurRow, .Rows - 1) To 2 Step -1
                        If .RowData(i) = -1 Or .TextMatrix(i, lngCurCol) <> "" Then  '�����ؿ��л�����������
                            lngCurRow = i + 1: Exit For
                        End If
                    Next
                End If
            End If
            '3.��������
            blnAddRow = False
            If .Rows - 1 < lngCurRow Then
                '�����в�������1��
                .Rows = .Rows + 1: lngCurRow = .Rows - 1
                .RowData(lngCurRow) = lngBackColor '�������ý���ɫ
                
                .TextMatrix(lngCurRow, COL_��ԴID) = Nvl(rsData!��ԴID)
                .TextMatrix(lngCurRow, COL_����ID) = Nvl(rsData!����ID)
                .TextMatrix(lngCurRow, COL_����) = Nvl(rsData!����)
                .TextMatrix(lngCurRow, col_����) = Nvl(rsData!����)
                .TextMatrix(lngCurRow, COL_����) = Nvl(rsData!����)
                .TextMatrix(lngCurRow, COL_��Ŀ) = Nvl(rsData!�շ���Ŀ)
                .TextMatrix(lngCurRow, Col_ҽ��) = Nvl(rsData!ҽ������)
                If m_DataStyle = Data_FixedRule Then
                    .TextMatrix(lngCurRow, COL_��ʼʱ��) = Format(Nvl(rsData!��ʼʱ��), "yyyy-mm-dd hh:mm:ss")
                    .TextMatrix(lngCurRow, COL_��ֹʱ��) = Format(Nvl(rsData!��ֹʱ��), "yyyy-mm-dd hh:mm:ss")
                End If
                blnAddRow = True
            End If
            
            If blnFind Then
                '�Ű����:1-�����Ű�;2-�����Ű�;3-˫���Ű�;4-������ѭ;5-��ѭ������;6-�ض�����
                'ԤԼ���ƣ�0-����ԤԼ����;1-�ú����ֹԤԼ;2-����ֹ��������ƽ̨��ԤԼ
                If Nvl(rsData!�ϰ�ʱ��) <> "" Then
                    Select Case m_DataStyle
                    Case Data_Templet
                        If Nvl(rsData!�Ű����) = 1 Then
                            .TextMatrix(lngCurRow, lngCurCol) = Nvl(rsData!�ϰ�ʱ��)
                            .TextMatrix(lngCurRow, lngCurCol + 1) = IIf(Val(Nvl(rsData!�޺���)) = 0, "��", Nvl(rsData!�޺���))
                            .TextMatrix(lngCurRow, lngCurCol + 2) = IIf(Nvl(rsData!ԤԼ����) = 1, "-", _
                                                                        IIf(Val(Nvl(rsData!��Լ��)) = 0, "��", Nvl(rsData!��Լ��)))
                        ElseIf Nvl(rsData!�Ű����) = 6 Then
                            If .TextMatrix(lngCurRow, lngCurCol) <> "" Then
                                .TextMatrix(lngCurRow, lngCurCol) = .TextMatrix(lngCurRow, lngCurCol) & "," & Val(Nvl(rsData!������Ŀ))
                            Else
                                .TextMatrix(lngCurRow, lngCurCol) = Val(Nvl(rsData!������Ŀ))
                                .TextMatrix(lngCurRow, lngCurCol + 1) = Nvl(rsData!�ϰ�ʱ��)
                                .TextMatrix(lngCurRow, lngCurCol + 2) = IIf(Val(Nvl(rsData!�޺���)) = 0, "��", Nvl(rsData!�޺���))
                                .TextMatrix(lngCurRow, lngCurCol + 3) = IIf(Nvl(rsData!ԤԼ����) = 1, "-", _
                                                                            IIf(Val(Nvl(rsData!��Լ��)) = 0, "��", Nvl(rsData!��Լ��)))
                            End If
                        Else
                            .TextMatrix(lngCurRow, lngCurCol) = Nvl(rsData!������Ŀ)
                            .TextMatrix(lngCurRow, lngCurCol + 1) = Nvl(rsData!�ϰ�ʱ��)
                            .TextMatrix(lngCurRow, lngCurCol + 2) = IIf(Val(Nvl(rsData!�޺���)) = 0, "��", Nvl(rsData!�޺���))
                            .TextMatrix(lngCurRow, lngCurCol + 3) = IIf(Nvl(rsData!ԤԼ����) = 1, "-", _
                                                                        IIf(Val(Nvl(rsData!��Լ��)) = 0, "��", Nvl(rsData!��Լ��)))
                        End If
                    Case Data_FixedRule
                        .TextMatrix(lngCurRow, lngCurCol) = Nvl(rsData!�ϰ�ʱ��)
                        .TextMatrix(lngCurRow, lngCurCol + 1) = IIf(Val(Nvl(rsData!�޺���)) = 0, "��", Nvl(rsData!�޺���))
                        .TextMatrix(lngCurRow, lngCurCol + 2) = IIf(Nvl(rsData!ԤԼ����) = 1, "-", _
                                                                    IIf(Val(Nvl(rsData!��Լ��)) = 0, "��", Nvl(rsData!��Լ��)))
                    Case Data_Plan
                        .TextMatrix(lngCurRow, lngCurCol) = Nvl(rsData!�ϰ�ʱ��)
                        .Cell(flexcpData, lngCurRow, lngCurCol) = Nvl(rsData!��¼ID)
                        If Val(Nvl(rsData!�Ƿ���ʱ����)) = 1 Then '��ʱ���������ɫ������ʾ
                            .Cell(flexcpForeColor, lngCurRow, lngCurCol, lngCurRow, lngCurCol + 2) = vbBlue
                        End If
                        If Val(Nvl(rsData!�Ƿ�����)) = 1 Then '��������ʾ����ͼ��
                            .Cell(flexcpPicture, lngCurRow, lngCurCol) = imgLock.Picture
                            .Cell(flexcpPictureAlignment, lngCurRow, lngCurCol) = flexAlignRightBottom
                        End If
                        If Nvl(rsData!ͣ�￪ʼʱ��) <> "" Then 'ͣ����ú�ɫ������ʾ
                            .Cell(flexcpForeColor, lngCurRow, lngCurCol, lngCurRow, lngCurCol + 2) = vbRed
                        End If
                        If Nvl(rsData!����ҽ������) <> "" Then '���������ɫ������ʾ����ʾ����ҽ��
                            .Cell(flexcpForeColor, lngCurRow, lngCurCol, lngCurRow, lngCurCol + 2) = vbBlue
'                            .Cell(flexcpPicture, lngCurRow, lngCurCol) = imgReplace.Picture
'                            .Cell(flexcpPictureAlignment, lngCurRow, lngCurCol) = flexAlignRightBottom
                            .TextMatrix(lngCurRow, lngCurCol) = .TextMatrix(lngCurRow, lngCurCol) & Chr(10) & Space(4 - Len(Nvl(rsData!����ҽ������))) & Nvl(rsData!����ҽ������)
                        End If
                        .TextMatrix(lngCurRow, lngCurCol + 1) = Nvl(rsData!�ѹ���) & "/" & IIf(Nvl(rsData!�޺���) = "", "��", Nvl(rsData!�޺���))
                        .TextMatrix(lngCurRow, lngCurCol + 2) = IIf(Nvl(rsData!ԤԼ����) = 1, "-", _
                                                                    Nvl(rsData!��Լ��) & "/" & IIf(Nvl(rsData!��Լ��) = "", "��", Nvl(rsData!��Լ��)))
                    End Select
                End If
                    
                If blnAddRow Then lngCurRow = lngCurRow + 1
            End If
            rsData.MoveNext
        Loop
        If .Rows > 1 Then 'ȱʡ��λ��
            .Row = -1 '��֤��ѡ���в���������Ҳ����RowColChange�¼�
            .Row = IIf(lngOldRow = 0 Or lngOldRow > .Rows - 1, .FixedRows, lngOldRow)
            .Col = IIf(lngoldCol = 0 Or lngoldCol > .Cols - 1, .FixedCols, lngoldCol)
        End If
        Call SetGridFormat
    End With
    LoadDataByRecordset = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SetGridFormat()
    Dim i As Long, j As Long
    Dim lngCurRowInGroup As Long, strSpace As String
    
    With vsfRegistPlan
        '���⴦���Ա��ܹ��ϲ�����
        lngCurRowInGroup = 0 '���������к�
        For i = 2 To .Rows - 1
            If .RowData(i) = -1 Then lngCurRowInGroup = 0
            For j = 0 To .Cols - 1
                If .RowData(i) = -1 Then Exit For
                If .TextMatrix(i, j) = "" Then .TextMatrix(i, j) = " " '��ֹ����Ϊ�ղ��ܺϲ�
                If .RowData(i - 1) <> -1 And j >= .FixedCols Then '�Ƿ�Ϊ��������
                    If (j - .FixedCols) Mod 3 = 0 Then '"ʱ���"��
                        If .TextMatrix(i, j) = " " Then '�ϲ�����Ŀ���
                            .TextMatrix(i, j) = .TextMatrix(i - 1, j)
                            .TextMatrix(i, j + 1) = .TextMatrix(i - 1, j + 1)
                            .TextMatrix(i, j + 2) = .TextMatrix(i - 1, j + 2)
                        Else
                            strSpace = Space(lngCurRowInGroup Mod 2) '���ո񣬷�ֹ������ͬ�ϲ�
                            .TextMatrix(i, j + 1) = strSpace & .TextMatrix(i, j + 1) & strSpace
                            .TextMatrix(i, j + 2) = strSpace & .TextMatrix(i, j + 2) & strSpace
                        End If
                        'ģ��������Ű�
                        If m_DataStyle = Data_Templet And j = .Cols - 4 Then
                            If .TextMatrix(i, j) = " " Then '�ϲ�����Ŀ���
                                .TextMatrix(i, j + 3) = .TextMatrix(i - 1, j + 3)
                            Else
                                strSpace = Space(lngCurRowInGroup Mod 2) '���ո񣬷�ֹ������ͬ�ϲ�
                                .TextMatrix(i, j + 3) = strSpace & .TextMatrix(i, j + 3) & strSpace
                            End If
                            j = j + 1
                        End If
                        j = j + 2
                    End If
                End If
            Next
            If .RowData(i) <> -1 Then lngCurRowInGroup = lngCurRowInGroup + 1
        Next
        
        '�б���ɫ
        For i = 2 To .Rows - 1
            .Cell(flexcpBackColor, i, .FixedCols, i, .Cols - 1) = .RowData(i)
        Next
        
        '������ɫ
        For i = 2 To .Rows - 1 '�ް��ŵĺ�Դ�û�ɫ��ʾ
            If Val(.TextMatrix(i, COL_����ID)) = 0 Then
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbGrayText
            End If
        Next
        If m_DataStyle = Data_Plan Then
            For j = .FixedCols To .Cols - 1 Step 3 'ʧЧ�İ����û�ɫ��ʾ
                If CDate(.Cell(flexcpData, 0, j)) < Format(Now, "yyyy-mm-dd") Then
                    .Cell(flexcpForeColor, 0, j, .Rows - 1, j + 2) = vbGrayText
                End If
            Next
        End If
            
        
        If .Row < .FixedRows And .Rows > .FixedRows Then .Row = .FixedRows
        If .RowData(.Row) = -1 And .Rows > .Row + 1 Then .Row = .Row + 1
        If .Col < .FixedCols And .Cols > .FixedCols Then .Col = .FixedCols
        Call SetVSFSelectedRangeColor(.Row, .Col, False)
    End With
End Sub

Private Sub SetVSFSelectedRangeColor(ByVal Row As Long, ByVal Col As Long, ByVal blnOld As Boolean)
    '���ܣ�����ѡ������ɫ,.RowData�д�����ɫֵ
    Dim i As Long
    Dim lngRowStart As Long, lngRowEnd As Long '��ʼ�к���ֹ��
    Dim lngColStart As Long, lngColEnd As Long '��ʼ�к���ֹ��

    On Error Resume Next
    With vsfRegistPlan
        If Not .Visible Then Exit Sub
        If Row > .FixedRows - 1 And Col > .FixedCols - 1 Then
'            lngRowStart = .FixedRows
'            For i = Row To .FixedRows Step -1
'                If .RowData(i) = -1 Then lngRowStart = i + 1: Exit For
'            Next
'            lngRowEnd = .Rows - 1
'            For i = Row + 1 To .Rows - 1
'                If .RowData(i) = -1 Then lngRowEnd = i - 1: Exit For
'            Next
            lngRowStart = Row
            lngRowEnd = Row
            lngColStart = GetItemNameCol(Col, .FixedCols) 'ȷ��"ʱ���"��
            If m_DataStyle = Data_Templet And lngColStart = .Cols - 4 Then lngColStart = lngColStart + 1
            lngColEnd = lngColStart + 2
            .Cell(flexcpBackColor, lngRowStart, lngColStart, lngRowEnd, lngColEnd) = IIf(blnOld, .RowData(lngRowStart), .BackColorSel)
        End If
    End With
End Sub

Private Function GetItemNameCol(ByVal lngCurCol As Long, ByVal lngFixedCols As Long) As Long
    'ȷ��"ʱ��"�е�������
    GetItemNameCol = lngCurCol - Choose(((lngCurCol - lngFixedCols) Mod 3) + 1, 0, 1, 2)
End Function

Private Sub GetGroupRange(ByVal lngCurRow As Long, ByRef lngRowStart As Long, ByRef lngRowEnd As Long)
    '��ǰ�����������Χ
    Dim i As Integer
    
    With vsfRegistPlan
        lngRowStart = .FixedRows
        For i = lngCurRow To .FixedRows Step -1
            If .RowData(i) = -1 Then lngRowStart = i + 1: Exit For
        Next
        lngRowEnd = .Rows - 1
        For i = lngCurRow + 1 To .Rows - 1
            If .RowData(i) = -1 And i <> .Rows - 1 Then lngRowEnd = i - 1: Exit For
        Next
    End With
End Sub

Private Sub vsfRegistPlan_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call SetVSFSelectedRangeColor(OldRow, OldCol, True)
    Call SetVSFSelectedRangeColor(NewRow, NewCol, False)
    RaiseEvent AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsfRegistPlan_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = COl_���� Then Cancel = True
End Sub

Private Sub vsfRegistPlan_DblClick()
    Dim i As Long, strSort As String
    Dim lngCol As Long, lngRow As Long
    
    RaiseEvent DblClick
    lngCol = vsfRegistPlan.MouseCol
    lngRow = vsfRegistPlan.MouseRow
    If (lngRow = 0 Or lngRow = 1) And Not mrsData Is Nothing Then
        '��������
        Select Case lngCol
        Case COL_����
            strSort = SortCircle(lngCol, "����") & "����,"
        Case col_����
            strSort = SortCircle(lngCol, "����")
        Case COL_����
            strSort = SortCircle(lngCol, "����") & "����,"
        Case COL_��Ŀ
            strSort = SortCircle(lngCol, "�շ���Ŀ") & "����,"
        Case Col_ҽ��
            strSort = SortCircle(lngCol, "ҽ������") & "����,"
        End Select
        If strSort <> "" Then
            Select Case m_DataStyle
            Case Data_FixedRule
                strSort = "����ID," & strSort & "��ʼʱ��,��ֹʱ��,������Ŀ,�ϰ�ʱ��"
            Case Data_Templet
                strSort = "����ID," & strSort & "������Ŀ,�ϰ�ʱ��"
            Case Data_Plan
                strSort = strSort & "��������,�ϰ�ʱ��"
            End Select
            mrsData.Sort = strSort
            Call LoadDataByRecordset(mrsData)
        End If
    End If
End Sub

Private Function SortCircle(ByVal lngCol As Long, ByVal strColName As String) As String
    'Cell(flexcpData, 1, lngCol)��¼�˵�ǰ����ʽ��ע�������¼�������ʱ���
    Select Case vsfRegistPlan.Cell(flexcpData, 1, lngCol)
    Case ""
        If lngCol = col_���� Then '�����г�ʼʱ�������������е�
            vsfRegistPlan.Cell(flexcpData, 1, lngCol) = "DESC"
            SortCircle = strColName & " DESC,"
        Else
            vsfRegistPlan.Cell(flexcpData, 1, lngCol) = "ASC"
            SortCircle = strColName & " Asc,"
        End If
    Case "ASC" '����
        vsfRegistPlan.Cell(flexcpData, 1, lngCol) = "DESC"
        SortCircle = strColName & " Desc,"
    Case "DESC" '����
        If lngCol = col_���� Then '������Ҫô����Ҫô����
            vsfRegistPlan.Cell(flexcpData, 1, lngCol) = "ASC"
            SortCircle = strColName & " Asc,"
        Else
            vsfRegistPlan.Cell(flexcpData, 1, lngCol) = "-"
            SortCircle = ""
        End If
    Case "-" '������
        vsfRegistPlan.Cell(flexcpData, 1, lngCol) = "ASC"
        SortCircle = strColName & " Asc,"
    End Select
End Function

Private Sub vsfRegistPlan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub vsfRegistPlan_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim varData As Variant, strTemp As String
    Dim lngRow As Long, lngCol As Long
    Dim sngLeft As Single, sngTop As Single

    On Error Resume Next
    If m_DataStyle <> Data_Plan Then Exit Sub
    With vsfRegistPlan
        If Not .Visible Then Exit Sub
        
        lngRow = .MouseRow: lngCol = .MouseCol
        If .Tag = lngRow & "," & lngCol Then Exit Sub
        .Tag = lngRow & "," & lngCol
        picToolTip.Visible = False
        
        If lngRow < .FixedRows Or lngCol < .FixedCols Then Exit Sub
        If (lngCol - .FixedCols) Mod 3 = 0 Then Exit Sub '"ʱ��"���˳�
        strTemp = .TextMatrix(lngRow, lngCol)
        If (strTemp = "" Or InStr(strTemp, "/") = 0) And strTemp <> "-" Then Exit Sub
        
        '2.��ʾ����
        If strTemp = "-" Then
            lblToolTip1.Caption = "��ֹԤԼ��"
            lblToolTip2.Visible = False
        Else
            lblToolTip2.Visible = True
            varData = Split(strTemp, "/")
            If (lngCol - .FixedCols) Mod 3 = 1 Then
                lblToolTip1.Caption = "�ѹ�����" & Trim(varData(0))
                lblToolTip2.Caption = "�޺�����" & IIf(Trim(varData(1)) = "��", "������", Trim(varData(1)))
            ElseIf (lngCol - .FixedCols) Mod 3 = 2 Then
                lblToolTip1.Caption = "��Լ����" & Trim(varData(0))
                lblToolTip2.Caption = "��Լ����" & IIf(Trim(varData(1)) = "��", "������", Trim(varData(1)))
            End If
        End If
        
        '3.��ʾ���λ��
        sngLeft = .Cell(flexcpLeft, lngRow, lngCol) + .ColWidth(lngCol) - 50
        sngTop = .Cell(flexcpTop, lngRow, lngCol) + .RowHeight(lngRow) - 10
        If sngLeft + picToolTip.Width > .Width Then
            sngLeft = .Cell(flexcpLeft, lngRow, lngCol) - picToolTip.Width + 50
        End If
        If sngTop + picToolTip.Height > .Height Then
            sngTop = .Cell(flexcpTop, lngRow, lngCol) - picToolTip.Height + 20
        End If
        picToolTip.Move sngLeft, sngTop
        picToolTip.Visible = True
    End With
End Sub

Private Sub picImgPlan_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    
    vRect = GetControlRect(picImgPlan.Hwnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImgPlan.Height
    Call frmVsColSel.ShowColSet(UserControl.Parent, UserControl.Parent.Caption, vsfRegistPlan, lngLeft, lngTop, picImgPlan.Height)
End Sub
