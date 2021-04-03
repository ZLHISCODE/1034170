VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.UserControl DateSelecter 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VSFlex8Ctl.VSFlexGrid vsfSelectDate 
      Height          =   2085
      Left            =   450
      TabIndex        =   0
      Top             =   720
      Width           =   3975
      _cx             =   7011
      _cy             =   3678
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "微软雅黑"
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
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   0
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   7
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"DateSelecter.ctx":0000
      ScrollTrack     =   0   'False
      ScrollBars      =   0
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
   Begin VB.Shape shpLine 
      BorderColor     =   &H80000003&
      Height          =   510
      Left            =   0
      Top             =   0
      Width           =   525
   End
   Begin VB.Image imgSelected 
      Height          =   120
      Left            =   750
      Picture         =   "DateSelecter.ctx":00F4
      Top             =   2820
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "DateSelecter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum g_Date_ShowStye
    Date_Show_Week = 0
    Date_Show_Month = 1
End Enum
'缺省属性值:
Const m_def_Enabled = 0
Const m_def_MultiSelectionMode As Boolean = False
Const m_def_Date_ShowStye = 0
Const m_def_BorderStyle = True
'属性变量:
Dim m_Enabled As Boolean
Dim m_MultiSelectionMode As Boolean
Dim m_Date_ShowStye As g_Date_ShowStye
Dim m_BorderStyle As Boolean

'事件声明:
Event Click(idxDay As Integer, Value As Boolean, Text As String)

Private Sub UserControl_Initialize()
    Call InitFace
    Call ClearAll
End Sub

Private Sub InitFace()
    Dim i As Integer, j As Integer, intCount As Integer
    Dim blnExitFor As Boolean
    
    With vsfSelectDate
        .Clear
        If m_Date_ShowStye = Date_Show_Week Then
            .Rows = 5: .Cols = 7 '5行7列
            .RowHeightMin = 350
        ElseIf m_Date_ShowStye = Date_Show_Month Then
            .Rows = 2: .Cols = 16 '2行16列
            .RowHeightMin = 100
        End If
        intCount = 1
        For i = 0 To .Rows - 1
            If blnExitFor Then Exit For
            For j = 0 To .Cols - 1
                If intCount > 31 Then blnExitFor = True: Exit For
                .TextMatrix(i, j) = intCount
                intCount = intCount + 1
            Next
        Next
        Call ClearAll
    End With
End Sub

Public Sub ClearAll()
    With vsfSelectDate
        .Cell(flexcpPictureAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignLeftTop
        .Cell(flexcpBackColor, 0, 0, .Rows - 1, .Cols - 1) = .BackColor
        .Cell(flexcpPicture, 0, 0, .Rows - 1, .Cols - 1) = Nothing
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
    End With
End Sub

Private Sub UserControl_Resize()
    On Error Resume Next
    With shpLine
        .Left = ScaleLeft
        .Width = ScaleWidth
        .Top = ScaleTop
        .Height = ScaleHeight
    End With
    With vsfSelectDate
        .Left = 10
        .Top = 10
        .Height = ScaleHeight - 20
        .Width = ScaleWidth - 20
'        .RowHeightMin = .Height / 5
        .ColWidth(-1) = .Width / vsfSelectDate.Cols
    End With
End Sub

Private Sub vsfSelectDate_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If vsfSelectDate.TextMatrix(NewRow, NewCol) = "" Then Cancel = True
End Sub

Private Sub vsfSelectDate_Click()
    Dim i As Integer, j As Integer
    Dim lngCol As Long, lngRow As Long
    Dim intIndex As Integer, blnValue As Boolean, strText As String
    Dim blnExitFor As Boolean
    
    With vsfSelectDate
        lngCol = .Col
        lngRow = .Row
        
        If lngCol < 0 Or lngCol > .Cols - 1 Then Exit Sub
        If lngRow < 0 Or lngRow > .Rows - 1 Then Exit Sub
        If .TextMatrix(lngRow, lngCol) = "" Then Exit Sub
        
        If .Cell(flexcpPicture, lngRow, lngCol) Is Nothing Then
            If m_MultiSelectionMode = False Then '单选模式
                .Cell(flexcpPicture, 0, 0, .Rows - 1, .Cols - 1) = Nothing
                .Cell(flexcpBackColor, 0, 0, .Rows - 1, .Cols - 1) = .BackColor
            End If
            .Cell(flexcpPicture, lngRow, lngCol) = imgSelected.Picture
            .Cell(flexcpBackColor, lngRow, lngCol) = 16772055
        Else
            .Cell(flexcpPicture, lngRow, lngCol) = Nothing
            .Cell(flexcpBackColor, lngRow, lngCol) = .BackColor
        End If
        
        intIndex = 1
        For i = 0 To .Rows - 1
            If blnExitFor Then Exit For
            For j = 0 To .Cols - 1
                If i = lngRow And j = lngCol Then
                    blnValue = Not (.Cell(flexcpPicture, lngRow, lngCol) Is Nothing)
                    strText = .TextMatrix(lngRow, lngCol)
                    blnExitFor = True: Exit For
                End If
                intIndex = intIndex + 1
            Next
        Next
    End With
    RaiseEvent Click(intIndex, blnValue, strText)
End Sub

Public Property Get GetSelectedItems() As String
    Dim i As Integer, j As Integer
    Dim strSeled As String
    
    For i = 0 To vsfSelectDate.Rows - 1
        For j = 0 To vsfSelectDate.Cols - 1
            If Not vsfSelectDate.Cell(flexcpPicture, i, j) Is Nothing Then
                strSeled = strSeled & "," & vsfSelectDate.TextMatrix(i, j)
            End If
        Next
    Next
    strSeled = Mid(strSeled, 2)
    GetSelectedItems = strSeled
End Property

Public Property Get ItemValue(ByVal intIndex As Integer) As Boolean
    If intIndex < 1 Or intIndex > 31 Then Exit Property
    
    Dim i As Integer, j As Integer, intCount As Integer
    Dim blnExitFor As Boolean
    intCount = 1
    For i = 0 To vsfSelectDate.Rows - 1
        If blnExitFor Then Exit For
        For j = 0 To vsfSelectDate.Cols - 1
            If intCount = intIndex Then
                ItemValue = Not (vsfSelectDate.Cell(flexcpPicture, i, j) Is Nothing)
                blnExitFor = True: Exit For
            End If
            intCount = intCount + 1
        Next
    Next
End Property

Public Property Let ItemValue(ByVal intIndex As Integer, ByVal New_ItemValue As Boolean)
    If intIndex < 1 Or intIndex > 31 Then Exit Property
    
    Dim i As Integer, j As Integer, intCount As Integer
    Dim blnExitFor As Boolean
    intCount = 1
    For i = 0 To vsfSelectDate.Rows - 1
        If blnExitFor Then Exit For
        For j = 0 To vsfSelectDate.Cols - 1
            If intCount = intIndex Then
                vsfSelectDate.Cell(flexcpPicture, i, j) = IIf(New_ItemValue, imgSelected.Picture, Nothing)
                vsfSelectDate.Cell(flexcpBackColor, i, j) = IIf(New_ItemValue, 16772055, vsfSelectDate.BackColor)
                blnExitFor = True: Exit For
            End If
            intCount = intCount + 1
        Next
    Next
End Property

Public Property Get Count() As Integer
    Count = 31
End Property

Public Property Get MultiSelectionMode() As Boolean
    MultiSelectionMode = m_MultiSelectionMode
End Property

Public Property Let MultiSelectionMode(ByVal NewValue As Boolean)
    m_MultiSelectionMode = NewValue
    PropertyChanged "MultiSelectionMode"
End Property

'注意！不要删除或修改下列被注释的行！
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get Date_ShowStye() As g_Date_ShowStye
    Date_ShowStye = m_Date_ShowStye
End Property

Public Property Let Date_ShowStye(ByVal New_Date_ShowStye As g_Date_ShowStye)
    m_Date_ShowStye = New_Date_ShowStye
    PropertyChanged "Date_ShowStye"
    Call InitFace
    UserControl_Resize
End Property

'为用户控件初始化属性
Private Sub UserControl_InitProperties()
    m_Enabled = m_def_Enabled
    m_MultiSelectionMode = m_def_MultiSelectionMode
    m_Date_ShowStye = m_def_Date_ShowStye
    m_BorderStyle = m_def_BorderStyle
End Sub

'从存贮器中加载属性值
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_MultiSelectionMode = PropBag.ReadProperty("MultiSelectionMode", m_def_MultiSelectionMode)
    m_Date_ShowStye = PropBag.ReadProperty("Date_ShowStye", m_def_Date_ShowStye)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    
    shpLine.Visible = m_BorderStyle
    Call InitFace
    UserControl_Resize
End Sub

'将属性值写到存储器
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("MultiSelectionMode", m_MultiSelectionMode, m_def_MultiSelectionMode)
    Call PropBag.WriteProperty("Date_ShowStye", m_Date_ShowStye, m_def_Date_ShowStye)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
End Sub


'注意！不要删除或修改下列被注释的行！
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As Boolean
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Boolean)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
    shpLine.Visible = m_BorderStyle
End Property
