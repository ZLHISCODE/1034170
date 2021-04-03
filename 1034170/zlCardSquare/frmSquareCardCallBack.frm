VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmSquareCardCallBack 
   BorderStyle     =   0  'None
   Caption         =   "卡回收记录"
   ClientHeight    =   8805
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   10935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VSFlex8Ctl.VSFlexGrid vsGrid 
      Height          =   2250
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   3930
      _cx             =   6932
      _cy             =   3969
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   9
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   300
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmSquareCardCallBack.frx":0000
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
         Left            =   60
         ScaleHeight     =   225
         ScaleWidth      =   210
         TabIndex        =   1
         Top             =   45
         Width           =   210
         Begin VB.Image imgCol 
            Height          =   195
            Left            =   0
            Picture         =   "frmSquareCardCallBack.frx":0066
            ToolTipText     =   "选择需要显示的列(ALT+C)"
            Top             =   0
            Width           =   195
         End
      End
   End
End
Attribute VB_Name = "frmSquareCardCallBack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String, mlngModule As Long, mblnHaveData As Boolean
Private mlng消费卡ID As Long, mlng接口编号 As Long
'一些公共事件
Public Event zlPopupMenus(ByVal vsGrid As VSFlexGrid) '弹出菜单操作
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
    Dim i As Long
    
    With vsGrid
        'ColData(i):列设置属性(1-固定,-1-不能选,0-可选)||列设置(0-允许移入,1-禁止移入,2-允许移入,但按回车后不能移入)
        .ColData(.ColIndex("ID")) = "-1|1"
        .ColData(.ColIndex("卡号")) = "1|1"
        .ColData(.ColIndex("领卡部门ID")) = "-1|1"
    End With
End Sub

Private Sub LoadDataToRpt()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载数据给网格
    '编制:刘兴洪
    '日期:2009-09-07 11:53:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strFilter As String, rsTemp As New ADODB.Recordset, lngRow As Long
    Dim blnHistory As Boolean, strStartDate As String, i As Long
    
    mblnHaveData = False
    Err = 0: On Error GoTo ErrHand:
    gstrSQL = "" & _
    "   Select a.Id,a.卡类型,a.卡号,a.序号,decode(a.可否充值,1,'√','') as 可否充值,decode(to_char(a.有效期,'yyyy-mm-dd'),'3000-01-01','',NULL,'',to_char(a.有效期,'yyyy-mm-dd hh24:mi:ss')) as 有效期,a.发卡原因, " & _
    "          a.发卡人,a.领卡人, to_char(a.发卡时间,'yyyy-mm-dd hh24:mi:ss') as 发卡时间, " & _
    "          a.回收人,to_char(a.回收时间,'yyyy-mm-dd hh24:mi:ss') as 回收时间 , " & _
    "          decode(mod(a.当前状态,10),2,'回收',3,'退卡','回收') as 当前状态,a.备注, " & _
    "          ltrim(to_char(a.卡面金额," & gOraFmtString.FM_金额 & ")) as 卡面金额 ," & _
    "          ltrim(to_char(a.销售金额," & gOraFmtString.FM_金额 & ")) as 销售金额 ," & _
    "          ltrim(to_char(a.充值折扣率," & gOraFmtString.FM_折扣率 & ")) as 充值折扣率 ," & _
    "          ltrim(to_char(a.余额," & gOraFmtString.FM_金额 & ")) as 余额 ," & _
    "          a.停用人,to_char(case when a.停用日期>=to_date('3000-01-01','yyyy-mm-dd') then NULL else a.停用日期 end  ,'yyyy-mm-dd hh24:mi:ss') as 停用日期," & _
    "          a.领卡部门ID,b.编码||'-'||b.名称 AS 领卡部门,a.限制类别 " & _
    "   From 消费卡目录 A,部门表 B,(Select 卡号,decode(当前状态,2,1,3,1,0)+序号 as 序号 From 消费卡目录 Where Id =[1]) C " & _
    "   Where  A.卡号=C.卡号 and A.序号<C.序号 and a.领卡部门id=b.Id(+)  and A.接口编号=[2] " & _
    "   Order by 序号"
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng消费卡ID, mlng接口编号)
    With Me.vsGrid
        .Redraw = flexRDNone
        Set .DataSource = rsTemp
        If .Rows <= 1 Then .Rows = 2
        For i = 1 To .Cols - 1
            .ColKey(i) = Trim(.TextMatrix(0, i))
            .FixedAlignment(i) = flexAlignCenterCenter
            If .ColKey(i) Like "*ID" Then
                .ColAlignment(i) = flexAlignCenterCenter
                .ColHidden(i) = True
            ElseIf .ColKey(i) Like "*额" Or .ColKey(i) Like "*率" Then
                .ColAlignment(i) = flexAlignRightCenter
            ElseIf .ColKey(i) Like "*时间" Or .ColKey(i) Like "*日期" Or .ColKey(i) Like "*状态" Or .ColKey(i) = "可否充值" Then
                .ColAlignment(i) = flexAlignCenterCenter
            Else
                .ColAlignment(i) = flexAlignLeftCenter
            End If
        Next
        Call InitVsGrid
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        '恢复列设置
        zl_vsGrid_Para_Restore mlngModule, vsGrid, Me.Caption, "回收列表", True
        .ColWidth(.ColIndex("标志")) = 285
        .ColAlignment(.ColIndex("标志")) = flexAlignCenterCenter
        .Redraw = flexRDBuffered
    End With
    mblnHaveData = rsTemp.RecordCount > 0
   Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
     Me.vsGrid.Redraw = flexRDBuffered
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
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "回收列表", True
End Sub
Private Sub imgCol_Click()
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    vRect = GetControlRect(picImg.hWnd)
    lngLeft = vRect.Left
    lngTop = vRect.Top + picImg.Height
    Call frmVsColSel.ShowColSet(Me, Me.Caption, vsGrid, lngLeft, lngTop, imgCol.Height)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "回收列表", True
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
    
    Err = 0: On Error GoTo errH:
    gstrSQL = "Select A.卡类型, A.卡号, to_char(A.发卡时间,'yyyy-mm-dd hh24:mi:ss') From 消费卡目录 A where ID=[1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng消费卡ID)
    If rsTemp.EOF = True Then Exit Sub '无卡信息，退出
    
    

    Set objPrint = New zlPrint1Grd
    objPrint.Title.Font.Name = "楷体_GB2312"
    objPrint.Title.Font.Size = 18
    objPrint.Title.Font.Bold = True
    
        
    objPrint.Title.Text = gstrUnitName & "消费卡回收情况"
    
    objRow.Add "卡类型：" & Nvl(rsTemp!卡类型)
    objRow.Add "卡号：" & Nvl(rsTemp!卡号)
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
errH:
    If ErrCenter = 1 Then Resume
    
End Sub

Private Sub vsGrid_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "回收列表", True
End Sub

Private Sub vsGrid_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    zl_VsGridRowChange vsGrid, OldRow, NewRow, OldCol, NewCol, gSysColor.lngGridColorSel
End Sub

Private Sub vsGrid_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModule, vsGrid, Me.Caption, "回收列表", True
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


