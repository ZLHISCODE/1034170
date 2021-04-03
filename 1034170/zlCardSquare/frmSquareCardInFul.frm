VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmSquareCardInFul 
   BorderStyle     =   0  'None
   Caption         =   "充值记录"
   ClientHeight    =   5565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VSFlex8Ctl.VSFlexGrid vsGrid 
      Height          =   2145
      Left            =   870
      TabIndex        =   0
      Top             =   1740
      Width           =   2895
      _cx             =   5106
      _cy             =   3784
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
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
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   9
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmSquareCardInFul.frx":0000
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
      ExplorerBar     =   7
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
      Begin VB.PictureBox picImg 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   45
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   1
         Top             =   60
         Width           =   210
         Begin VB.Image imgCol 
            Height          =   195
            Left            =   0
            Picture         =   "frmSquareCardInFul.frx":018E
            ToolTipText     =   "选择需要显示的列(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
   End
End
Attribute VB_Name = "frmSquareCardInFul"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String, mlngModule As Long, mblnHaveData As Boolean
Private mlng消费卡ID As Long, mlng接口编号 As Long
'一些公共事件
Public Event zlPopupMenus(ByVal vsGrid As VSFlexGrid) '弹出菜单操作
Public Event AfterRowChange(ByVal vsGrid As VSFlexGrid) '弹出菜单操作

Public Function zlReLoadData(ByVal lng接口编号 As Long, ByVal lng消费卡ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新加载数据
    '入参:mcllFilter-过滤条件(目前无)
    '出参:
    '返回:加载成功,返回true,否则返回Flase
    '编制:刘兴洪
    '日期:2009-11-20 16:00:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlng消费卡ID = lng消费卡ID: mlng接口编号 = lng接口编号
    Err = 0: On Error GoTo ErrHand:
    Call LoadDataToRpt
    zlReLoadData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitVsGrid()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化网格数据
    '编制:刘兴洪
    '日期:2009-11-20 16:05:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsGrid
        'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
        .ColData(.ColIndex("记录状态")) = "-1|1"
        .ColData(.ColIndex("ID")) = "-1|1"
        .ColHidden(.ColIndex("ID")) = True
        .ColHidden(.ColIndex("记录状态")) = True
        
        .ColData(.ColIndex("序号")) = "1|1"
        .ColData(.ColIndex("充值金额")) = "1|0"
    End With
End Sub

Private Sub LoadDataToRpt()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载数据给网格
    '编制:刘兴洪
    '日期:2009-09-07 11:53:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFilter As String, rsTemp As New ADODB.Recordset, lngRow As Long
    Dim blnHistory As Boolean, strStartDate As String
    
    mblnHaveData = False
    Err = 0: On Error GoTo ErrHand:
    gstrSQL = "" & _
    "   Select ID, 消费卡id, 序号, 记录状态, 充值金额, 充值折扣, 缴款金额, to_char(充值时间,'yyyy-mm-dd hh24:mi:ss') as 充值时间, 操作员姓名, 缴款人,备注 " & _
    "  From 消费卡充值记录 A " & _
    "  Where a.消费卡id=[1]" & _
    "  Order by 序号 desc,充值时间 desc "
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng消费卡ID)
    
    With Me.vsGrid
        .Clear 1
        .Rows = 2: lngRow = 1
        .Redraw = flexRDNone
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("ID")) = Nvl(rsTemp!ID)
            .TextMatrix(lngRow, .ColIndex("序号")) = Nvl(rsTemp!序号)
            .Cell(flexcpData, lngRow, .ColIndex("序号")) = Nvl(rsTemp!ID)
            .TextMatrix(lngRow, .ColIndex("记录状态")) = Nvl(rsTemp!记录状态)
            .TextMatrix(lngRow, .ColIndex("充值金额")) = Format(Val(Nvl(rsTemp!充值金额)), "####0.00;-####0.00;;")
            .TextMatrix(lngRow, .ColIndex("充值折扣")) = Format(Val(Nvl(rsTemp!充值折扣)), "####0.00;-####0.00;;")
            .TextMatrix(lngRow, .ColIndex("缴款金额")) = Format(Val(Nvl(rsTemp!缴款金额)), "####0.00;-####0.00;;")
            .TextMatrix(lngRow, .ColIndex("充值时间")) = Nvl(rsTemp!充值时间)
            .TextMatrix(lngRow, .ColIndex("充值说明")) = Nvl(rsTemp!备注)
            .TextMatrix(lngRow, .ColIndex("操作员姓名")) = Nvl(rsTemp!操作员姓名)
            .TextMatrix(lngRow, .ColIndex("缴款人")) = Nvl(rsTemp!缴款人)
        
            Select Case Val(Nvl(rsTemp!记录状态))
            Case 2
                .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = vbRed
            Case 3
                .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = &H8000000D
            Case Else
                .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = &H80000012
            End Select
            lngRow = lngRow + 1
           rsTemp.MoveNext
        Loop
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        .Redraw = flexRDBuffered
        mblnHaveData = rsTemp.RecordCount > 0
        '恢复列设置
        zl_vsGrid_Para_Restore mlngModule, vsGrid, Me.Caption, "充值列表", True
        .ColWidth(.ColIndex("标志")) = 285
        If .Rows >= 1 Then .Row = 1
    End With
    Call vsGrid_AfterRowColChange(0, 0, 1, 1)
   Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    zlCommFun.StopFlash
End Sub
Private Sub Form_Load()
    mlngModule = glngModul: mstrPrivs = gstrPrivs

    Call InitVsGrid
    Call vsGrid_GotFocus
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With vsGrid
        .Left = ScaleLeft: .Top = ScaleTop
        .Width = ScaleWidth: .Height = ScaleHeight
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "充值列表", True
End Sub
Private Sub imgCol_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = GetControlRect(picImg.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImg.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsGrid, lngLeft, lngTop, imgCol.Height)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "充值列表", True
End Sub

Private Sub picImg_Click()
    Call imgCol_Click
End Sub
 
Public Sub zlRptPrint(ByVal bytFunc As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:进行打印,预览和输出到EXCEL
    '入参:bytFunc=1 打印;2 预览;3 输出到EXCEL
    '编制:刘兴洪
    '日期:2009-11-20 16:36:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intCol As Long, objPrint As Object, objRow As New zlTabAppRow, bytPrn As Byte
    Dim rsTemp As New ADODB.Recordset
    
    
    Dim vsGrid As VSFlexGrid
    Err = 0: On Error GoTo ErrH:
    
    gstrSQL = "Select A.卡类型, A.卡号, to_char(A.发卡时间,'yyyy-mm-dd hh24:mi:ss') From 消费卡目录 A where ID=[1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng消费卡ID)
    If rsTemp.EOF = True Then Exit Sub '无卡信息，退出
    
    

    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
        
    objPrint.Title.Text = gstrUnitName & "消费卡充值情况"
    
    objRow.Add "卡类型：" & Nvl(rsTemp!卡类型)
    objRow.Add "卡号：" & Nvl(rsTemp!卡号)
    objRow.Add "发卡时间：" & Nvl(rsTemp!发卡时间)
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & UserInfo.姓名
    objRow.Add "打印日期:" & Format(zldatabase.Currentdate, "yyyy年MM月dd日")
    objPrint.BelowAppRows.Add objRow
    
    Err = 0: On Error GoTo ErrHand:
    With vsGrid
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .Cell(flexcpData, 0, intCol) = .ColWidth(intCol)
            If .ColHidden(intCol) Or intCol = .ColIndex("标志") Then .ColWidth(intCol) = 0
        Next
    End With
    Set objPrint.Body = vsGrid
    If bytFunc = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytPrn
    End If
    
    With vsGrid
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .ColWidth(intCol) = Val(.Cell(flexcpData, 0, intCol))
        Next
        .Redraw = flexRDBuffered
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    With vsGrid
        .Redraw = flexRDNone
        For intCol = 0 To .Cols - 1
            .ColWidth(intCol) = Val(.Cell(flexcpData, 0, intCol))
        Next
        .Redraw = flexRDBuffered
    End With
    Exit Sub
ErrH:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub vsGrid_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsGrid, OldRow, NewRow, OldCol, NewCol, gSysColor.lngGridColorSel
    If OldRow <> NewRow Then
        RaiseEvent AfterRowChange(vsGrid)
    End If
End Sub

 Private Sub vsGrid_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsGrid
        If Col = .ColIndex("标志") Then Cancel = True
    End With
End Sub
Private Sub vsGrid_GotFocus()
    zl_VsGridGotFocus vsGrid, gSysColor.lngGridColorSel
End Sub
Private Sub vsGrid_LostFocus()
    zl_VsGridLOSTFOCUS vsGrid, gSysColor.lngGridColorLost
End Sub

Private Sub vsGrid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button <> vbRightButton Then Exit Sub
    RaiseEvent zlPopupMenus(vsGrid)
End Sub
'------------------------------------------------------------------------------------------------------------------
'设置相关属性
Public Property Get zlIsHaveData() As Boolean
    zlIsHaveData = mblnHaveData
End Property

Public Property Get zl允许回退() As Boolean
    With vsGrid
        If .Row < 1 Then zl允许回退 = False: Exit Property
        zl允许回退 = Val(.TextMatrix(.Row, .ColIndex("记录状态"))) = 1
    End With
End Property
Public Property Get zlGet充值ID() As Long
    With vsGrid
        If .Row < 1 Then zlGet充值ID = 0: Exit Property
        zlGet充值ID = Val(.TextMatrix(.Row, .ColIndex("ID")))
    End With
End Property
Public Function zl充值取消() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:取消充值
    '编制:刘兴洪
    '日期:2009-12-22 17:36:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long
    Err = 0: On Error GoTo ErrHand:
    With vsGrid
        If MsgBox("你真要取消当前（序号为" & .TextMatrix(.Row, .ColIndex("序号")) & "  充值金额为:" & .TextMatrix(.Row, .ColIndex("充值金额")) & ")的充值记录吗？", vbQuestion + vbDefaultButton2 + vbYesNo) = vbNo Then Exit Function
        ' Zl_消费卡充值记录_Cancel
        gstrSQL = "Zl_消费卡充值记录_Cancel("
        '  Id_In     In 消费卡充值记录.ID%Type,
        gstrSQL = gstrSQL & "" & Val(.TextMatrix(.Row, .ColIndex("ID"))) & ","
        '  取消人_In In 消费卡充值记录.操作员姓名%Type
         gstrSQL = gstrSQL & "'" & UserInfo.姓名 & "')"
         zldatabase.ExecuteProcedure gstrSQL, Me.Caption
    End With
    Call LoadDataToRpt
    zl充值取消 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume

End Function




