VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmReused 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   7860
   ClientLeft      =   15
   ClientTop       =   -255
   ClientWidth     =   17970
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   17970
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   17970
      TabIndex        =   11
      Top             =   0
      Width           =   17970
      Begin VB.Frame fraTopCmd 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFF0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   12720
         TabIndex        =   17
         Top             =   120
         Width           =   5175
         Begin VB.CommandButton cmdResizeAll 
            Caption         =   "收缩全部数据文件"
            Height          =   375
            Left            =   0
            TabIndex        =   20
            Top             =   120
            Width           =   1695
         End
         Begin VB.CommandButton cmdResizeTemp 
            Caption         =   "收缩临时数据文件"
            Height          =   375
            Left            =   1695
            TabIndex        =   19
            Top             =   120
            Width           =   1695
         End
         Begin VB.CommandButton cmdResizeUndo 
            Caption         =   "收缩Undo表空间"
            Height          =   375
            Left            =   3390
            TabIndex        =   18
            Top             =   120
            Width           =   1695
         End
      End
      Begin VB.TextBox txtPrompt 
         Appearance      =   0  'Flat
         BackColor       =   &H00EFF0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   1200
         MultiLine       =   -1  'True
         TabIndex        =   12
         Text            =   "frmReused.frx":0000
         Top             =   120
         Width           =   9255
      End
      Begin VB.Line lineTop 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   10000
         Y1              =   830
         Y2              =   830
      End
      Begin VB.Line lineTop 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   10000
         Y1              =   845
         Y2              =   845
      End
      Begin VB.Image imgTop 
         Height          =   720
         Left            =   120
         Picture         =   "frmReused.frx":0139
         Top             =   75
         Width           =   720
      End
   End
   Begin VB.TextBox txtFind 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00808080&
      Height          =   350
      Left            =   16100
      TabIndex        =   9
      Tag             =   "请输入名称后按回车"
      Text            =   "请输入名称后按回车"
      ToolTipText     =   "首尾输入*号可进行模糊查找"
      Top             =   870
      Width           =   1785
   End
   Begin VB.CommandButton cmdLOBGO 
      Caption         =   "定位到LOB"
      Height          =   350
      Left            =   13800
      TabIndex        =   8
      Top             =   890
      Width           =   1215
   End
   Begin VB.CheckBox chkFree 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "只显示空块"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11640
      TabIndex        =   7
      Top             =   930
      Width           =   1275
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   650
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   17970
      TabIndex        =   5
      Top             =   7215
      Width           =   17970
      Begin VB.CommandButton cmdShrinkAll 
         Caption         =   "回收当前表空间所有对象"
         Height          =   375
         Left            =   8850
         TabIndex        =   16
         ToolTipText     =   "对当前表空间的所有文件中的对象执行Shrink操作"
         Top             =   120
         Width           =   2175
      End
      Begin VB.CommandButton cmdMoveAll 
         Caption         =   "重整当前表空间所有对象"
         Height          =   375
         Left            =   4530
         TabIndex        =   15
         ToolTipText     =   "对当前表空间的所有文件中的对象执行Move操作"
         Top             =   120
         Width           =   2175
      End
      Begin VB.CommandButton cmdShrink 
         Caption         =   "回收(Shrink)当前对象"
         Height          =   375
         Left            =   6795
         TabIndex        =   14
         ToolTipText     =   "可在线执行而不影响业务的正常使用，一般用于大量删除数据后降低高水标记以收回空间"
         Top             =   120
         Width           =   2055
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "重整(Move)当前对象"
         Height          =   375
         Left            =   2600
         TabIndex        =   13
         ToolTipText     =   "通过Move命令(对于索引，是Rebuild)移动当前对象所在块的物理位置以便收缩文件，需要获得对象的排它独占锁，执行期间会影响业务正常使用"
         Top             =   120
         Width           =   1935
      End
      Begin VB.CommandButton cmdResize 
         Caption         =   "收缩(Resize)当前数据文件"
         Height          =   375
         Left            =   60
         TabIndex        =   6
         ToolTipText     =   "收缩当前表空间中当前数据文件的大小，如果文件尾部存在对象，则无法收缩"
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label lblPrompt 
         BackColor       =   &H00EFF0E0&
         ForeColor       =   &H00C00000&
         Height          =   350
         Left            =   11115
         TabIndex        =   21
         ToolTipText     =   "每个单元格包含8个数据块。"
         Top             =   170
         Width           =   6735
      End
      Begin VB.Line lineBottom 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   10000
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line lineBottom 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   10000
         Y1              =   15
         Y2              =   15
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfExtents 
      Height          =   5850
      Left            =   2640
      TabIndex        =   4
      Top             =   1280
      Width           =   15255
      _cx             =   26908
      _cy             =   10319
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
      ForeColorSel    =   12582912
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   32768
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
      OwnerDraw       =   1
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
   Begin VSFlex8Ctl.VSFlexGrid vsfTbs 
      Height          =   5850
      Left            =   60
      TabIndex        =   3
      Top             =   1280
      Width           =   2535
      _cx             =   4471
      _cy             =   10319
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
      GridColor       =   32768
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   200
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
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
      OwnerDraw       =   1
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
   Begin VB.ComboBox cboFiles 
      Height          =   300
      ItemData        =   "frmReused.frx":698B
      Left            =   3390
      List            =   "frmReused.frx":698D
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   900
      Width           =   8175
   End
   Begin VB.Label lblFind 
      BackColor       =   &H00FFFFFF&
      Caption         =   "对象查找"
      Height          =   255
      Left            =   15195
      TabIndex        =   10
      Top             =   930
      Width           =   735
   End
   Begin VB.Label lblFiles 
      BackColor       =   &H00FFFFFF&
      Caption         =   "数据文件"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblTableSpaces 
      BackColor       =   &H00FFFFFF&
      Caption         =   "表空间列表"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
   Begin VB.Menu mnuResize 
      Caption         =   "收缩选项"
      NegotiatePosition=   2  'Middle
      Visible         =   0   'False
      Begin VB.Menu mnuResizeAll 
         Caption         =   "收缩全部数据文件"
      End
      Begin VB.Menu mnuResizeTemp 
         Caption         =   "收缩临时数据文件"
      End
      Begin VB.Menu mnuResizeUndo 
         Caption         =   "收缩Undo表空间"
      End
   End
End
Attribute VB_Name = "frmReused"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const CONCOLS As Long = 50
Private Const CONBLOCKS As Long = 8
Private mrsExtents As ADODB.Recordset
Private mrsLobs As ADODB.Recordset
Private mcolCells As Collection
Private mlngRowPre As Long, mlngColPre As Long

Private Enum opt
    P1回收 = 1
    P2重整
    P3收缩
End Enum

Private Enum rowColor
    OFF_颜色 = &HB3DEF5
End Enum

Public Sub ShowMe()
    Me.Show
End Sub

Private Sub cboFiles_Click()
    
    '由于循环中使用了doevents，所以需禁用任何可操作的功能
    Call SetCommandEnable(0)
    
    Call LoadExtents(vsfTbs.TextMatrix(vsfTbs.Row, vsfTbs.ColIndex("名称")), Val(cboFiles.ItemData(cboFiles.ListIndex)))
    cboFiles.ToolTipText = cboFiles.List(cboFiles.ListIndex)
    Call SetCommandEnable(1)
    
    vsfExtents.SetFocus
    vsfExtents.Select vsfExtents.Rows - 1, vsfExtents.Cols - 1
    vsfExtents.TopRow = vsfExtents.Row
End Sub

Private Sub SetCommandEnable(bytEnable As Byte)
'功能：设置命令按钮的可用性
    fraTopCmd.Enabled = bytEnable = 1
    picBottom.Enabled = bytEnable = 1
    
    chkFree.Enabled = bytEnable = 1
    txtFind.Enabled = bytEnable = 1
    
    If cmdLOBGO.Visible Then cmdLOBGO.Enabled = bytEnable = 1
        
    vsfTbs.Enabled = cmdShrink.Enabled
    cboFiles.Enabled = cmdShrink.Enabled
End Sub

Private Sub chkFree_Click()
    If cboFiles.ListIndex >= 0 Then Call cboFiles_Click
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Function ResizeTBS(ByVal strTbs As String, Optional ByVal lngFile As Long) As Boolean
'功能：收缩表空间
'参数：strTBS-表空间名称
'      blnPrompt-数据文件号,不传入时，在不提示的情况下收缩当前表空间的所有数据文件至最小尺寸
    Dim strSql As String, dblMax As Double, dblFileSize As Double, dblLimit As Double, dblBlockSize As Double
    Dim i As Long, blnTry As Boolean
    Dim rstmp As ADODB.Recordset
           
    On Error GoTo errH
    
    dblBlockSize = Val(vsfTbs.RowData(vsfTbs.Row))
    If dblBlockSize = 0 Then dblBlockSize = 8192
        
    If lngFile <> 0 Then
        dblLimit = CDbl(1024) * 1024 * 2
        
        strSql = "Select a.File_Id, a.Last_Block, b.Bytes" & vbNewLine & _
            "From (Select a.File_Id, Max(a.Block_Id + a.Blocks - 1) Last_Block" & vbNewLine & _
            "       From Dba_Extents A" & vbNewLine & _
            "       Where a.Tablespace_Name = [1] And File_Id = [2]" & vbNewLine & _
            "       Group By a.File_Id) A, Dba_Data_Files B" & vbNewLine & _
            "Where a.File_Id = b.File_Id"
        Set rstmp = OpenSQLRecord(strSql, Me.Caption, strTbs, lngFile)
        
        If rstmp.RecordCount = 0 Then
            MsgBox "在Dba_Extents中没有找到当前表空间及数据文件的记录", vbInformation, "错误"
            Exit Function
        End If
    
        dblMax = rstmp!Last_Block * dblBlockSize
        dblFileSize = rstmp!Bytes
        If dblFileSize - dblMax < dblLimit Then '小于2M，不收缩
            If MsgBox("可收缩的空间(" & Round((dblFileSize - dblMax) / 1024) & "KB)小于2M,是否确实要收缩该文件？", vbYesNo + vbDefaultButton2, "提醒") = vbNo Then
                Exit Function
            End If
            dblMax = Round(dblMax / 1024 / 1024) + 1 '取整加1，单位M
        Else
            dblMax = Round(dblMax / 1024 / 1024) + 1 '取整加1，单位M
            If MsgBox("你确定要将当前文件收缩到" & dblMax & "M吗?", vbQuestion + vbOKCancel + vbDefaultButton1, Me.Caption) = vbCancel Then
                Exit Function
            End If
        End If
        
        If dblMax >= Round(rstmp!Bytes / 1024 / 1024) Then
            MsgBox "数据文件已达到最大尺寸，无法更改！", vbInformation
        Else
            Err.Clear
            On Error Resume Next
retry1:     strSql = "Alter Database Datafile " & lngFile & " Resize " & CStr(dblMax) & "M"
            gcnOracle.Execute strSql
            
            If Err.Number <> 0 Then
                If MsgBox("收缩数据文件失败，可能是删除对象后未清空回收站引起的，是否清空后重试？", vbYesNo + vbDefaultButton1, Me.Caption) = vbYes Then
                    Err.Clear
                    strSql = "purge tablespace " & strTbs
                    gcnOracle.Execute strSql
                    GoTo retry1
                Else
                    GoTo errH
                End If
            End If
            ResizeTBS = True
        End If
        
    Else
        dblLimit = CDbl(1024) * 1024 * 10
        
        '可收缩空间小于10M，不收缩，避免在循环中频繁执行收缩
        strSql = "Select a.File_Id, a.Last_Block * " & dblBlockSize & " as MaxBytes, b.Bytes" & vbNewLine & _
                "From (Select a.File_Id, Max(a.Block_Id + a.Blocks - 1) Last_Block" & vbNewLine & _
                "       From Dba_Extents A" & vbNewLine & _
                "       Where a.Tablespace_Name = [1]" & vbNewLine & _
                "       Group By a.File_Id) A, Dba_Data_Files B" & vbNewLine & _
                "Where a.File_Id = b.File_Id And (b.Bytes - a.Last_Block * " & dblBlockSize & ") > " & dblLimit
    
        Set rstmp = OpenSQLRecord(strSql, Me.Caption, strTbs)
        
        On Error Resume Next
        For i = 1 To rstmp.RecordCount
            dblMax = Round(rstmp!MaxBytes / 1024 / 1024) + 1 '取整加1，单位M
            If dblMax < Round(rstmp!Bytes / 1024 / 1024) Then
                lblPrompt.Caption = "收缩" & rstmp!File_Id & "号数据文件至" & CStr(dblMax) & "M"
                lblPrompt.Refresh
                blnTry = False
retry2:         strSql = "Alter Database Datafile " & rstmp!File_Id & " Resize " & CStr(dblMax) & "M"
                gcnOracle.Execute strSql
                If Err.Number <> 0 And blnTry = False Then
                    Err.Clear
                    strSql = "purge tablespace " & strTbs
                    gcnOracle.Execute strSql
                    blnTry = True
                    GoTo retry2
                Else
                    Err.Clear   '重试一次后跳过
                End If
                
                ResizeTBS = True
            End If
            
            rstmp.MoveNext
        Next
    End If
    
    Exit Function
errH:
    Call ErrCenter(strSql)
    Call SetCommandEnable(1)
End Function

Private Function CheckSelFile() As Boolean
'功能：检查当前是否选择了文件

    If cboFiles.ListCount <= 0 Then
        MsgBox "请选择一个有数据文件的表空间！", vbInformation, "提醒"
        If cboFiles.Enabled Then cboFiles.SetFocus
    End If
    CheckSelFile = cboFiles.ListCount > 0
End Function

Private Sub cmdResize_Click()
'功能：执行数据文件收缩
    If cboFiles.ListIndex >= 0 Then
        Call SetCommandEnable(0)
        
        If ResizeTBS(vsfTbs.TextMatrix(vsfTbs.Row, vsfTbs.ColIndex("名称")), Val(cboFiles.ItemData(cboFiles.ListIndex))) Then
            lblPrompt.Caption = "已完成文件收缩，正在刷新。"
            lblPrompt.Refresh
            
            Call RefreshData
            
            lblPrompt.Caption = "已完成操作。"
        End If
        
        Call SetCommandEnable(1)
    Else
        MsgBox "请选择一个数据文件！", vbInformation, "提醒"
    End If
End Sub

Private Sub RefreshData()
'功能：刷新当前表空间的当前数据文件的段的数据信息

    Dim i As Long, strTbs As String
    Dim lngFile As Long
    
    strTbs = vsfTbs.TextMatrix(vsfTbs.Row, vsfTbs.ColIndex("名称"))
    lngFile = cboFiles.ListIndex
    Call LoadTablespaces
    
    vsfTbs.Redraw = flexRDNone
    i = vsfTbs.FindRow(strTbs, , vsfTbs.ColIndex("名称"))
    If i <> -1 Then vsfTbs.Row = i: vsfTbs.TopRow = i
    vsfTbs.Redraw = flexRDDirect
    
    Call LoadFiles(vsfTbs.TextMatrix(vsfTbs.Row, vsfTbs.ColIndex("名称")))
    If lngFile <= cboFiles.ListCount Then
        cboFiles.ListIndex = lngFile
    Else
        cboFiles.ListIndex = 0
    End If
End Sub

Private Function CheckUnSuportObject(strSegment As String, strOpt As String, Optional blnIndex As Boolean) As Boolean
'功能：检查指定的表是否存在Move或Shrink不支持的对象:LONG,LONG RAW
    Dim strSql As String
    Dim rstmp As ADODB.Recordset
    
    'All_Tab_Cols比All_Tab_Columns多一些隐藏字段
    strSql = "Select 1" & vbNewLine & _
            "From All_Tab_Cols" & vbNewLine & _
            "Where Table_Name = " & IIf(blnIndex, "(Select max(Table_name) From All_indexes Where index_Name = [2] And Owner = [1])", "[2]") & " And Owner = [1] And Data_Type In ('LONG','LONG RAW')"
    Set rstmp = OpenSQLRecord(strSql, Me.Caption, Split(strSegment, ".")(0), Split(strSegment, ".")(1))
    If rstmp.RecordCount > 0 Then
        lblPrompt.Caption = strSegment & "含有LONG,LONG RAW类型字段，不能进行" & strOpt & "操作."
        CheckUnSuportObject = True
    Else
        CheckUnSuportObject = False
    End If
End Function

Private Function CheckFuncIDX(strSegment As String) As Boolean
'功能：检查指定的表是否存在函数索引（不支持shrink表）
    Dim strSql As String
    Dim rstmp As ADODB.Recordset
    
    strSql = "Select 1 From All_Indexes Where Owner = [1] And Table_Name = [2] And Index_Type = 'FUNCTION-BASED NORMAL' And rownum<2"
    Set rstmp = OpenSQLRecord(strSql, Me.Caption, Split(strSegment, ".")(0), Split(strSegment, ".")(1))
    CheckFuncIDX = rstmp.RecordCount > 0
End Function

Private Function CheckIOT(strSegment As String) As Boolean
'功能：检查指定的索引是否为索引组织表的索引（不支持重建）
    Dim strSql As String
    Dim rstmp As ADODB.Recordset
    
    strSql = "Select 1 From All_Indexes Where Owner = [1] And Index_Name = [2] And Index_Type = 'IOT - TOP'"
    Set rstmp = OpenSQLRecord(strSql, Me.Caption, Split(strSegment, ".")(0), Split(strSegment, ".")(1))
    CheckIOT = rstmp.RecordCount > 0
End Function

Private Function CheckIOTTab(strSegment As String) As Boolean
'功能：检查指定的表是否为索引组织表（不支持并行重建）
    Dim strSql As String
    Dim rstmp As ADODB.Recordset
    
    strSql = "Select 1 From All_Tables Where Owner = [1] And Table_Name = [2] And Iot_Name Is Not Null"
    Set rstmp = OpenSQLRecord(strSql, Me.Caption, Split(strSegment, ".")(0), Split(strSegment, ".")(1))
    CheckIOTTab = rstmp.RecordCount > 0
End Function

Private Function GetIOTName(strSegment As String) As String
'功能：根据索引组织表的索引名返回索引组织表名(含所有者前缀)
    Dim strSql As String
    Dim rstmp As ADODB.Recordset
    
    strSql = "Select Table_Owner||'.'||Table_Name as Tab_Name From All_Indexes Where Owner = [1] And Index_Name = [2] And Index_Type = 'IOT - TOP'"
    Set rstmp = OpenSQLRecord(strSql, Me.Caption, Split(strSegment, ".")(0), Split(strSegment, ".")(1))
    If rstmp.RecordCount > 0 Then
        GetIOTName = rstmp!Tab_Name
    End If
End Function


Private Function CheckLOBIndex(strSegment As String) As Boolean
'功能：检查指定的索引是否为LOB的索引（不支持重建）
    Dim strSql As String
    Dim rstmp As ADODB.Recordset
    
    strSql = "Select 1 From All_Lobs Where Owner = [1] And Index_Name = [2]"
    Set rstmp = OpenSQLRecord(strSql, Me.Caption, Split(strSegment, ".")(0), Split(strSegment, ".")(1))
    CheckLOBIndex = rstmp.RecordCount > 0
End Function

Private Function GetLOBNameByIndex(strSegment As String) As String
'功能：检查指定的索引是否为LOB的索引（不支持重建）
    Dim strSql As String
    Dim rstmp As ADODB.Recordset
    
    strSql = "Select Segment_Name From All_Lobs Where Owner = [1] And Index_Name = [2]"
    Set rstmp = OpenSQLRecord(strSql, Me.Caption, Split(strSegment, ".")(0), Split(strSegment, ".")(1))
    If rstmp.RecordCount > 0 Then
        GetLOBNameByIndex = Split(strSegment, ".")(0) & "." & rstmp!Segment_Name
    End If
End Function

Private Sub ReBuildIndex(ByVal strOwner As String, ByVal strTable As String, ByVal strParallel As String)
'功能：重建某张表上失效的索引
'参数：strOwner=所有者,strTable=表名
'      strParallel=" Parallel X",并行度
    Dim rstmp As ADODB.Recordset, rsIndex As ADODB.Recordset
    Dim strSql As String
    
    lblPrompt.Caption = "正在重建[" & strOwner & "." & strTable & "]上失效的索引"
    lblPrompt.Refresh
    On Error GoTo errH
    
    '重建失效的索引
    strSql = "Select Index_Name From DBA_Indexes Where Status='UNUSABLE' And Owner = [1] And Table_Name = [2]"
    Set rstmp = OpenSQLRecord(strSql, Me.Caption, strOwner, strTable)
    
    Do While Not rstmp.EOF
        '如果是分区索引，则要单独处理
        strSql = "Select Partition_Name From Dba_Ind_Partitions Where Index_Owner = [1] And Index_Name = [2]"
        Set rsIndex = OpenSQLRecord(strSql, Me.Caption, strOwner, rstmp!Index_Name)
        If rsIndex.RecordCount > 0 Then
            Do While Not rsIndex.EOF
                strSql = "Alter Index " & strOwner & "." & rstmp!Index_Name & " Rebuild Partition " & rsIndex!Partition_Name & " Nologging" & strParallel
                gcnOracle.Execute strSql
                rsIndex.MoveNext
            Loop
        Else
            strSql = "Alter Index " & strOwner & "." & rstmp!Index_Name & " Rebuild Nologging" & strParallel
            gcnOracle.Execute strSql
        End If
        
        If strParallel <> "" Then
            strSql = "Alter Index " & strOwner & "." & rstmp!Index_Name & " NOParallel"
            gcnOracle.Execute strSql
        End If
        
        rstmp.MoveNext
    Loop
    Exit Sub
errH:
    Call ErrCenter(strSql)
End Sub

Private Sub AdjustDBParameter()
'功能：调整重建索引的性能相关参数,缺省情况，每个会话一般只能用到最大200M的PGA，通过以下调整一般可以加快40%以上的创建速度
    Dim strSql(6) As String
    Dim i As Long
    
    On Error Resume Next
    
    strSql(0) = "alter session set workarea_size_policy=MANUAL"
    
    '直接路径IO的大小
    strSql(1) = "alter session set events '10351 trace name context forever, level 128'"
    strSql(2) = "alter session SET db_file_multiblock_read_count=128"
    strSql(3) = "alter session set ""_sort_multiblock_read_count""=128"
    
    '10g目前存在一个bug，sort_area_size需要手动设置2次才能生效！
    strSql(4) = "alter session SET sort_area_size=2000000000"
    strSql(5) = "alter session SET sort_area_size=2000000000"
            
    For i = 0 To UBound(strSql) - 1
        gcnOracle.Execute strSql(i)
        
        If Err.Number <> 0 Then
            MsgBox "索引重建优化的参数调整失败，已忽略。" & vbCrLf & strSql(i) & vbCrLf & Err.Description, vbInformation
            Err.Clear
        End If
    Next
End Sub


Private Sub cmdMoveAll_Click()
'功能：重整当前表空间的所有文件中的对象
'      从文件的尾部开始，直到最前面一个空块的位置
    Dim rstmp As ADODB.Recordset
    Dim i As Long
    Dim bytMode As Byte, blnRemove As Boolean
    Dim datBegin As Date, strTime As String
    
    Dim lngErrCount As Long, strPrompt As String, strOnline As String, strParallel As String
    
    Dim strSegmentAll As String, strSegment As String, strSegment_Type As String, strSegmentPre As String
    Dim strTbsTemp As String, strTbsOriginal As String, strSql As String
        
    Dim strRemoveTable As String, strRemovePARTable As String, strRemoveLob As String, strRemoveIndex As String, strRemovePARIndex As String, strRemovePARLOB As String
    
    
    If CheckSelFile = False Then Exit Sub

    strTbsOriginal = vsfTbs.TextMatrix(vsfTbs.Row, vsfTbs.ColIndex("名称"))
    strPrompt = "重整操作需要在对象上持有独占排它锁，可能影响相关业务运行，并且需要大量空闲空间，请慎重执行。" & _
            "Move表之后将会自动重建失效的索引，可能耗时较长。" & vbCrLf & vbCrLf & _
            "你确定要对表空间" & strTbsOriginal & "的" & cboFiles.ListCount & "个文件中的所有对象进行重整操作吗？"
            
    If MsgBox(strPrompt, vbOKCancel + vbDefaultButton1, Me.Caption) = vbCancel Then
        Exit Sub
    End If
    
    If frmSelTablespace.ShowMe(Me, bytMode, strTbsTemp, strParallel, strOnline) = False Then
        Exit Sub
    End If
    If strTbsTemp = strTbsOriginal Then strTbsTemp = ""
    
        
    Me.Refresh  '避免表格上的残影
    datBegin = GetCurrentdate
    Call SetCommandEnable(0)
    
    For i = 0 To cboFiles.ListCount - 1
        blnRemove = False
       
        strRemoveTable = "": strRemovePARTable = "": strRemoveLob = "": strRemoveIndex = "": strRemovePARIndex = "": strRemovePARLOB = ""
        RefreshInfo "正在查询[" & cboFiles.List(i) & "]文件中的所有对象"
        Set rstmp = GetExtents(strTbsOriginal, Val(cboFiles.ItemData(i)))
        
        '如果是移到其他表空间再移回来，由于移回来之前会先收缩原表空间的数据文件，所以要从末尾开始重整
        Do While Not IIf(bytMode = 2, rstmp.BOF, rstmp.EOF)
            strSegment = rstmp!Full_Segment_Name
            strSegment_Type = rstmp!Segment_Type
                            
            If strSegment & "|" & strSegment_Type <> strSegmentPre Then
                If InStr(strSegmentAll & ",", "," & strSegment & "|" & strSegment_Type & ",") = 0 Then
                    ExecMove strSegment, strSegment_Type, strSegmentAll, lngErrCount, strTbsOriginal, strTbsTemp, strOnline, strParallel, _
                        strRemoveTable, strRemovePARTable, strRemoveLob, strRemoveIndex, strRemovePARIndex, strRemovePARLOB
                
                    strSegmentAll = strSegmentAll & "," & strSegment & "|" & strSegment_Type
                End If
                strSegmentPre = strSegment & "|" & strSegment_Type
            End If
            
            If bytMode = 2 Then
                rstmp.MovePrevious
            Else
                rstmp.MoveNext
            End If
        Loop
        
        If bytMode = 2 And strTbsTemp <> "" Then
reMove:     blnRemove = True
            ExecMoveOriginal strTbsTemp, strOnline, strParallel, _
                    strRemoveTable, strRemovePARTable, strRemoveLob, strRemoveIndex, strRemovePARIndex, strRemovePARLOB
        End If
    Next
    
    
    '刷新数据
    Call RefreshData
    Call ShowSpendTime(datBegin, "重整", lngErrCount)
            
    If strParallel <> "" Then
        Call SetNOParallel(gcnOracle, 0)
        Call SetNOParallel(gcnOracle, 1)
    End If
    
    Call SetCommandEnable(1)
End Sub

Private Sub ExecMove(strSegment As String, ByVal strSegment_Type As String, strSegmentAll As String, lngErrCount As Long, ByVal strTbsOriginal As String, ByVal strTbsTemp As String, ByVal strOnline As String, ByVal strParallel As String, _
    strRemoveTable As String, strRemovePARTable As String, strRemoveLob As String, strRemoveIndex As String, strRemovePARIndex As String, strRemovePARLOB As String _
    )
'功能：执行一个对象的Move或Rebuild操作
'参数：strSegment可能被改变，因为LOBINDEX类型时需重新获取原始的段名
    Dim strObjName As String, strPartition As String
    Dim rstmp As ADODB.Recordset, strSql As String
    Dim strTbsLob As String, strColumn As String, strTableName As String
        
    On Error GoTo errH
    
    '1.普通表
    If strSegment_Type = "TABLE" Then
        'mdsys用户下存在类似GridFile1044_TAB这种含有小写字母的表段，但在dba_tables中却查不到记录
        If CheckSuportTable(strSegment) Then
            If Not CheckUnSuportObject(strSegment, "重整(Move)") Then
                RefreshInfo "正在重整：" & strSegment
                DoEvents
                If strTbsTemp = "" Then
                    strSql = "Alter Table " & strSegment & " Move Nologging" & strParallel
                    gcnOracle.Execute strSql
                    Call ReBuildIndex(Split(strSegment, ".")(0), Split(strSegment, ".")(1), strParallel)
                Else
                    strSql = "Alter Table " & strSegment & " Move TableSpace " & strTbsTemp & " Nologging" & strParallel
                    gcnOracle.Execute strSql
                    strRemoveTable = strRemoveTable & "," & strSegment & "||" & strTbsOriginal
                End If
                
                If strParallel <> "" Then
                    strSql = "Alter Table " & strSegment & " NOParallel"
                    gcnOracle.Execute strSql
                End If
            Else
                TraceFileWrite "未重整含有Long或Long Raw字段的表:" & strSegment
            End If
        Else
            TraceFileWrite "未重整IOT索引的溢出表或含有自定义字段的表:" & strSegment
        End If

    '2.分区表(不含LOB分区表)
    ElseIf strSegment_Type = "TABLE PARTITION" Then
        If Not CheckUnSuportObject(strSegment, "重整(Move)") Then
            
            strSql = "Select Partition_Name From Dba_Tab_Partitions Where Table_Owner = [1] And Table_Name = [2]"
            Set rstmp = OpenSQLRecord(strSql, Me.Caption, Split(strSegment, ".")(0), Split(strSegment, ".")(1))
            Do While Not rstmp.EOF
                strPartition = rstmp!Partition_Name
                
                RefreshInfo "正在重整：" & strSegment & "(" & strPartition & ")"
                DoEvents
                '未加级联更新索引update indexes，在后面调用ReBuildIndex来恢复，因为可能移两次
                If strTbsTemp = "" Then
                    strSql = "Alter Table " & strSegment & " Move Partition " & strPartition & " Nologging" & strParallel
                    gcnOracle.Execute strSql
                    
                    Call ReBuildIndex(Split(strSegment, ".")(0), Split(strSegment, ".")(1), strParallel)
                Else
                    strSql = "Alter Table " & strSegment & " Move Partition " & strPartition & " TableSpace " & strTbsTemp & " Nologging" & strParallel
                    gcnOracle.Execute strSql
                    strRemovePARTable = strRemovePARTable & "," & strSegment & "||" & strPartition & "||" & strTbsOriginal
                End If
                rstmp.MoveNext
            Loop
            
            If strParallel <> "" Then
                strSql = "Alter Table " & strSegment & " NOParallel"
                gcnOracle.Execute strSql
            End If
        Else
            TraceFileWrite "未重整含有Long或Long Raw字段的分区表:" & strSegment
        End If
        
    '3.LOB段（不含LOB分区索引和LOB分区表）
    ElseIf strSegment_Type = "LOBSEGMENT" Or strSegment_Type = "LOBINDEX" Then
        If strSegment_Type = "LOBINDEX" Then
            strSql = "Select Owner ||'.'|| Segment_Name as Segment_Name From Dba_Lobs Where Owner = [1] And Index_Name = [2]"
            Set rstmp = OpenSQLRecord(strSql, Me.Caption, Split(strSegment, ".")(0), Split(strSegment, ".")(1))
            If rstmp.RecordCount > 0 Then
                If InStr(strSegmentAll & ",", "," & rstmp!Segment_Name & "|" & strSegment_Type & ",") = 0 Then
                    strSegment = rstmp!Segment_Name
                Else
                    Exit Sub '如果LOB已重整过，则跳过
                End If
            End If
        End If
    
        mrsLobs.Filter = "Owner='" & Split(strSegment, ".")(0) & "' And Segment_Name='" & Split(strSegment, ".")(1) & "'"   '为了取表名和列名
        If mrsLobs.RecordCount > 0 Then
            'mdsys用户下存在类似GridFile1044_TAB这种含有小写字母的表段，但在dba_tables中却查不到记录
            If CheckSuportTable(mrsLobs!Owner & "." & mrsLobs!Table_name) Then
                If Not CheckUnSuportObject(mrsLobs!Owner & "." & mrsLobs!Table_name, "重整(Move)") Then
                    strTableName = mrsLobs!Owner & "." & mrsLobs!Table_name
                    strTbsLob = mrsLobs!Tablespace_Name
                    strColumn = mrsLobs!Column_Name
                    
                    RefreshInfo "正在重整：" & strTableName & "(" & strColumn & ")"
                    DoEvents
                    If strTbsTemp = "" Then
                        strSql = "ALTER TABLE " & strTableName & " Move LOB (" & strColumn & ") Store as(Tablespace " & strTbsLob & ") Nologging" & strParallel
                        gcnOracle.Execute strSql
                        
                    Else
                        strSql = "ALTER TABLE " & strTableName & " Move LOB (" & strColumn & ") Store as(Tablespace " & strTbsTemp & ") Nologging" & strParallel
                        gcnOracle.Execute strSql
                        strRemoveLob = strRemoveLob & "," & strTableName & "||" & strColumn & "||" & strTbsLob
                    End If
                Else
                     TraceFileWrite "未重整含有Long或Long Raw字段的表::" & mrsLobs!Table_name
                End If
                'LOB并行执行不会导致表及索引的degree属性被设置，所以不必执行noparallel
            Else
                TraceFileWrite "未重整IOT索引的溢出表或含有自定义字段的表:" & mrsLobs!Table_name
            End If
        Else
            RefreshInfo "在视图Dba_Lobs中未找到LOB对象" & strSegment & "。"
        End If
        
    '4.普通索引
    ElseIf strSegment_Type = "INDEX" Then
        If CheckIOT(strSegment) = False Then    'IOT索引只能通过move原表重建
            RefreshInfo "正在重建：" & strSegment
            DoEvents
            If strTbsTemp = "" Then
                strSql = "Alter Index " & strSegment & " Rebuild " & strOnline & " Nologging" & strParallel
                gcnOracle.Execute strSql
            Else
                strSql = "Alter Index " & strSegment & " Rebuild " & strOnline & " TableSpace " & strTbsTemp & " Nologging" & strParallel
                gcnOracle.Execute strSql
                strRemoveIndex = strRemoveIndex & "," & strSegment & "||" & strTbsOriginal
            End If
                                        
            If strParallel <> "" Then
                strSql = "Alter Index " & strSegment & " NOParallel"
                gcnOracle.Execute strSql
            End If
            
        Else 'IOT索引组织表（不支持并行执行）
            strObjName = GetIOTName(strSegment)
            RefreshInfo "正在重整：" & strObjName
            DoEvents
            '部分索引组织表(SYSMAN.AQ$_MGMT_NOTIFY_QTABLE_I)move时会报错：ORA-08108: 可能没有建立或重建该类型的索引联机
            '忽略错误
            On Error Resume Next
            If strTbsTemp = "" Then
                strSql = "Alter Table " & strObjName & " Move Nologging"
                gcnOracle.Execute strSql
            Else
                strSql = "Alter Table " & strObjName & " Move TableSpace " & strTbsTemp & " Nologging"
                gcnOracle.Execute strSql
                strRemoveTable = strRemoveTable & "," & strObjName & "||" & strTbsOriginal
            End If
            
            If Err.Number <> 0 Then
                TraceFileWrite "IOT表的Move失败:" & strObjName & "," & Err.Description
                Err.Clear
            End If
        End If
        
    '5.分区索引
    ElseIf strSegment_Type = "INDEX PARTITION" Then
        If CheckLOBIndex(strSegment) Then
            'LOB分区索引跟LOB分区表一起Move
            
        ElseIf CheckIOT(strSegment) = False Then
            
            strSql = "Select Partition_Name From Dba_Ind_Partitions Where Index_Owner = [1] And Index_Name = [2]"
            Set rstmp = OpenSQLRecord(strSql, Me.Caption, Split(strSegment, ".")(0), Split(strSegment, ".")(1))
            Do While Not rstmp.EOF
                strPartition = rstmp!Partition_Name
            
                RefreshInfo "正在重建：" & strSegment & "(" & strPartition & ")"
                DoEvents
                If strTbsTemp = "" Then
                    strSql = "Alter Index " & strSegment & " Rebuild Partition " & strPartition & " Nologging" & strParallel & " " & strOnline
                    gcnOracle.Execute strSql
                Else
                    strSql = "Alter Index " & strSegment & " Rebuild Partition " & strPartition & " TableSpace " & strTbsTemp & " Nologging" & strParallel & " " & strOnline
                    gcnOracle.Execute strSql
                    strRemovePARIndex = strRemovePARIndex & "," & strSegment & "||" & strPartition & "||" & strTbsOriginal
                End If
                rstmp.MoveNext
            Loop
                                        
            If strParallel <> "" Then
                strSql = "Alter Index " & strSegment & " NOParallel"
                gcnOracle.Execute strSql
            End If
        Else
            TraceFileWrite "未重整索引组织表（IOT）的分区索引:" & strSegment
        End If
        
    '6.LOB分区表
    ElseIf strSegment_Type = "LOB PARTITION" Then
        If Not CheckUnSuportObject(strSegment, "重整(Move)") Then
                                                
            mrsLobs.Filter = "Owner='" & Split(strSegment, ".")(0) & "' And Segment_Name='" & Split(strSegment, ".")(1) & "'"   '为了取表空间名
            If mrsLobs.RecordCount > 0 Then
                strTableName = mrsLobs!Owner & "." & mrsLobs!Table_name
                strTbsLob = mrsLobs!Tablespace_Name
                strColumn = mrsLobs!Column_Name
                
                strSql = "Select Partition_Name From Dba_Lob_Partitions Where Table_Owner = [1] And Lob_Name = [2]"
                Set rstmp = OpenSQLRecord(strSql, Me.Caption, Split(strSegment, ".")(0), Split(strSegment, ".")(1))
                Do While Not rstmp.EOF
                    strPartition = rstmp!Partition_Name
                    
                    RefreshInfo "正在重整：" & strTableName & "(" & strPartition & ")"
                    DoEvents
                    If strTbsTemp = "" Then
                        strSql = "Alter Table " & strTableName & " Move Partition " & strPartition & " Lob(" & strColumn & ") Store as (Tablespace " & strTbsLob & ") Nologging" & strParallel
                        gcnOracle.Execute strSql
                    Else
                        strSql = "Alter Table " & strTableName & " Move Partition " & strPartition & " Lob(" & strColumn & ") Store as (Tablespace " & strTbsTemp & ") Nologging" & strParallel
                        gcnOracle.Execute strSql
                        strRemovePARLOB = strRemovePARLOB & "," & strTableName & "||" & strPartition & "||" & strColumn & "||" & strTbsLob
                    End If
                    rstmp.MoveNext
                Loop
                
                'LOB分区并行执行不会导致表及索引的degree属性被设置，所以不必执行noparallel
                If strTbsTemp = "" Then Call ReBuildIndex(Split(strTableName, ".")(0), Split(strTableName, ".")(1), strParallel)
            Else
                TraceFileWrite "在视图Dba_Lobs中未找到LOB对象" & strSegment & "。"
            End If
        Else
            TraceFileWrite "未重整含有Long或Long Raw字段的分区表:" & strSegment
        End If
    
    ElseIf strSegment_Type <> " " Then
        TraceFileWrite strSegment & ",不支持的对象类型：" & strSegment_Type
    End If

    Exit Sub
errH:
    Call TraceFileWrite(Err.Description & vbCrLf & strSql, LTT_OnlyTime)
    lngErrCount = lngErrCount + 1
End Sub

Private Sub cmdMove_Click()
'功能：执行表或索引的空闲空间重整(Move)
    Dim strSegment_Type As String, strSegmentAll As String, strSegment As String, strSegmentPre As String
    Dim strTbsTemp As String, strTbsOriginal As String
    
    Dim datBegin As Date, strTime As String, strSql As String
    Dim r1 As Long, c1 As Long, r2 As Long, c2 As Long, c As Long, r As Long
    Dim blnRemove As Boolean, bytMode As Byte
    Dim lngErrCount As Long, strOnline As String, strParallel As String
    
    '移到临时存储的表空间，最后才移回来，记录原始表空间
    Dim strRemoveTable As String, strRemovePARTable As String, strRemoveLob As String, strRemoveIndex As String, strRemovePARIndex As String, strRemovePARLOB As String

        
    If CheckExtent(P2重整) = False Then Exit Sub
    
    strSegment_Type = Trim(lblPrompt.Tag)
    If strSegment_Type <> "" And InStr(",TABLE,TABLE PARTITION,INDEX,INDEX PARTITION,LOBSEGMENT,LOBINDEX,LOB PARTITION,", "," & strSegment_Type & ",") = 0 Then '对LOBINDEX对象，则重整其LOBSEGMENT
        Call MsgBox("仅支持对表或索引进行空闲空间收回，不支持的数据类型：" & strSegment_Type, vbInformation, Me.Caption)
        Exit Sub
    End If
    Call SetCommandEnable(0)
    
    '在windows环境下验证，没有效果，可能受限于内存大小和磁盘性能
    '住院费用记录_IX_结帐ID，这类索引，差别不大，而“电子病历内容_PK”这类索引，设置了“workarea_size_policy=MANUAL”后，反而慢几倍
'    If cmdMove.Tag = "" Then
'        Call AdjustDBParameter
'        cmdMove.Tag = "已优化"
'    End If
    
    
    If frmSelTablespace.ShowMe(Me, bytMode, strTbsTemp, strParallel, strOnline) = False Then
        Exit Sub
    End If
    
    strTbsOriginal = vsfTbs.TextMatrix(vsfTbs.Row, vsfTbs.ColIndex("名称"))
    If strTbsTemp = strTbsOriginal Then strTbsTemp = ""
    
    
    Me.Refresh  '避免表格上的残影
    datBegin = GetCurrentdate
    
    
    '处理一次选择多行多列的情况
    With vsfExtents
        .GetSelection r1, c1, r2, c2
                
        For r = r2 To r1 Step -1
            For c = c2 To c1 Step -1
                strSegment = mcolCells("_" & r & "_" & c)     '含所有者
                strSegment_Type = CStr(.Cell(flexcpData, r, c))
                
                If strSegment & "|" & strSegment_Type <> strSegmentPre Then
                    If InStr(strSegmentAll & ",", "," & strSegment & "|" & strSegment_Type & ",") = 0 Then
                    
                        mrsExtents.Filter = "Row=" & r & " And Col=" & c
                        If mrsExtents.RecordCount > 0 Then
                            ExecMove strSegment, strSegment_Type, strSegmentAll, lngErrCount, strTbsOriginal, strTbsTemp, strOnline, strParallel, _
                                strRemoveTable, strRemovePARTable, strRemoveLob, strRemoveIndex, strRemovePARIndex, strRemovePARLOB
                        End If
                        strSegmentAll = strSegmentAll & "," & strSegment & "|" & strSegment_Type
                    End If
                    strSegmentPre = strSegment & "|" & strSegment_Type
                End If
            Next
        Next
    End With
    
    '对移到临时存储的表空间的对象，移回原表空间。
    If bytMode = 2 And strTbsTemp <> "" Then
        blnRemove = True
        ExecMoveOriginal strTbsTemp, strOnline, strParallel, _
                strRemoveTable, strRemovePARTable, strRemoveLob, strRemoveIndex, strRemovePARIndex, strRemovePARLOB
    End If
    
    '刷新数据
    Call RefreshData
    
    If strSegment <> "" Then
        mrsExtents.Filter = "Segment_Name='" & Split(strSegment, ".")(1) & "' And Owner='" & Split(strSegment, ".")(0) & "'"
        If mrsExtents.RecordCount > 0 Then
            vsfExtents.SetFocus
            vsfExtents.Select mrsExtents!Row, mrsExtents!Col
            vsfExtents.TopRow = vsfExtents.Row
        End If
    End If
    
    Call ShowSpendTime(datBegin, "重整", lngErrCount)
        
    If strParallel <> "" Then
        Call SetNOParallel(gcnOracle, 0)
        Call SetNOParallel(gcnOracle, 1)
    End If
    
    Call SetCommandEnable(1)
End Sub

Private Sub ExecMoveOriginal(ByVal strTbsTemp As String, ByVal strOnline As String, ByVal strParallel As String, _
        strRemoveTable As String, strRemovePARTable As String, strRemoveLob As String, strRemoveIndex As String, strRemovePARIndex As String, strRemovePARLOB As String)
'功能：将对象移回到原表空间，并重建失效索引
    Dim arrTmp As Variant, strSql As String
    Dim r As Long
    Dim strSegment As String, strTbsOriginal As String, strPartition As String
    Dim strColumn As String
    
    On Error GoTo errH
     '1.表
    If strRemoveTable <> "" Then
         arrTmp = Split(Mid(strRemoveTable, 2), ",")
         For r = 0 To UBound(arrTmp)
             strSegment = Split(arrTmp(r), "||")(0)
             strTbsOriginal = Split(arrTmp(r), "||")(1)
             
             DoEvents
             If r = 0 Then
                 RefreshInfo "正在收缩" & strTbsOriginal & "的空间"
                 Call ResizeTBS(strTbsOriginal)
             End If
             
             RefreshInfo "正在将[" & strSegment & "]移回原表空间"
             
             '索引组织表不支持并行执行
             If CheckIOTTab(strSegment) Then
                strSql = "Alter Table " & strSegment & " Move TableSpace " & strTbsOriginal & " Nologging"
                gcnOracle.Execute strSql
             Else
                strSql = "Alter Table " & strSegment & " Move TableSpace " & strTbsOriginal & " Nologging" & strParallel
                gcnOracle.Execute strSql
                            
                If strParallel <> "" Then
                    strSql = "Alter Table " & strSegment & " NOParallel"
                    gcnOracle.Execute strSql
                End If
                        
                Call ReBuildIndex(Split(strSegment, ".")(0), Split(strSegment, ".")(1), strParallel)
             End If
         Next
         If strRemovePARTable & strRemoveLob & strRemoveIndex & strRemovePARIndex & strRemovePARLOB = "" Then GoTo ResizeLine
     End If
     
     '2.分区表
    If strRemovePARTable <> "" Then
         arrTmp = Split(Mid(strRemovePARTable, 2), ",")
         For r = 0 To UBound(arrTmp)
             strSegment = Split(arrTmp(r), "||")(0)
             strPartition = Split(arrTmp(r), "||")(1)
             strTbsOriginal = Split(arrTmp(r), "||")(2)
             
             DoEvents
             If r = 0 Then
                 RefreshInfo "正在收缩" & strTbsOriginal & "的空间"
                 Call ResizeTBS(strTbsOriginal)
             End If
             
             RefreshInfo "正在将[" & strSegment & "(" & strPartition & ")]移回原表空间"
             
             strSql = "Alter Table " & strSegment & " Move Partition " & strPartition & " TableSpace " & strTbsOriginal & " Nologging" & strParallel
             gcnOracle.Execute strSql
             
             
             If strParallel <> "" Then
                 strSql = "Alter Table " & strSegment & " NOParallel"
                 gcnOracle.Execute strSql
                 
             End If
                  
             '移回最后一个分区后重建表上失效的索引
             If r = UBound(arrTmp) Then
                 Call ReBuildIndex(Split(strSegment, ".")(0), Split(strSegment, ".")(1), strParallel)
             End If
         Next
         If strRemoveLob & strRemoveIndex & strRemovePARIndex & strRemovePARLOB = "" Then GoTo ResizeLine
     End If
     
     '3.LOB段
     If strRemoveLob <> "" Then
         arrTmp = Split(Mid(strRemoveLob, 2), ",")
         For r = 0 To UBound(arrTmp)
             strSegment = Split(arrTmp(r), "||")(0)
             strColumn = Split(arrTmp(r), "||")(1)
             strTbsOriginal = Split(arrTmp(r), "||")(2)
             
             DoEvents
             If r = 0 Then
                 RefreshInfo "正在收缩" & strTbsOriginal & "的空间"
                 Call ResizeTBS(strTbsOriginal)
             End If
                         
             RefreshInfo "正在将[" & strSegment & "(" & strColumn & ")]移回原表空间"
             
             strSql = "ALTER TABLE " & strSegment & " Move LOB (" & strColumn & ") Store as(Tablespace " & strTbsOriginal & ") Nologging" & strParallel
             gcnOracle.Execute strSql
             
             
             'LOB并行执行不会导致表及索引的degree属性被设置，所以不必执行noparallel
         Next
         If strRemoveIndex & strRemovePARIndex & strRemovePARLOB = "" Then GoTo ResizeLine
     End If
     
     '4.索引
     If strRemoveIndex <> "" Then
         arrTmp = Split(Mid(strRemoveIndex, 2), ",")
         For r = 0 To UBound(arrTmp)
             strSegment = Split(arrTmp(r), "||")(0)
             strTbsOriginal = Split(arrTmp(r), "||")(1)
             
             DoEvents
             If r = 0 Then
                 RefreshInfo "正在收缩" & strTbsOriginal & "的空间"
                 Call ResizeTBS(strTbsOriginal)
             End If
                         
             RefreshInfo "正在将[" & strSegment & "]移回原表空间"
             
             strSql = "Alter Index " & strSegment & " Rebuild " & strOnline & " TableSpace " & strTbsOriginal & " Nologging" & strParallel
             gcnOracle.Execute strSql
                                             
             If strParallel <> "" Then
                 strSql = "Alter Index " & strSegment & " NOParallel"
                 gcnOracle.Execute strSql
             End If
         Next
         
         If strRemovePARIndex & strRemovePARLOB = "" Then GoTo ResizeLine
     End If
     
     '5.分区索引
     If strRemovePARIndex <> "" Then
         arrTmp = Split(Mid(strRemovePARIndex, 2), ",")
         For r = 0 To UBound(arrTmp)
             strSegment = Split(arrTmp(r), "||")(0)
             strPartition = Split(arrTmp(r), "||")(1)
             strTbsOriginal = Split(arrTmp(r), "||")(2)
             
             DoEvents
             If r = 0 Then
                 RefreshInfo "正在收缩" & strTbsOriginal & "的空间"
                 Call ResizeTBS(strTbsOriginal)
             End If
                         
             RefreshInfo "正在将[" & strSegment & "(" & strPartition & ")]移回原表空间"
             
             strSql = "Alter Index " & strSegment & " Rebuild Partition " & strPartition & " TableSpace " & strTbsOriginal & " Nologging" & strParallel & " " & strOnline
             gcnOracle.Execute strSql
             
                                
             If strParallel <> "" Then
                 strSql = "Alter Index " & strSegment & " NOParallel"
                 gcnOracle.Execute strSql
                 
             End If
         Next
         If strRemovePARLOB = "" Then GoTo ResizeLine
     End If
     
      '6.LOB分区
     If strRemovePARLOB <> "" Then
         arrTmp = Split(Mid(strRemovePARLOB, 2), ",")
         For r = 0 To UBound(arrTmp)
             strSegment = Split(arrTmp(r), "||")(0)
             strPartition = Split(arrTmp(r), "||")(1)
             strColumn = Split(arrTmp(r), "||")(2)
             strTbsOriginal = Split(arrTmp(r), "||")(3)
             
             DoEvents
             If r = 0 Then
                 RefreshInfo "正在收缩" & strTbsOriginal & "的空间"
                 Call ResizeTBS(strTbsOriginal)
             End If
                         
             RefreshInfo "正在将[" & strSegment & "(" & strPartition & ")]移回原表空间"
             
             strSql = "Alter Table " & strSegment & " Move Partition " & strPartition & " Lob(" & strColumn & ") Store as (Tablespace " & strTbsOriginal & ") Nologging" & strParallel
             gcnOracle.Execute strSql
             
             '移回最后一个分区后重建表上失效的索引
             If r = UBound(arrTmp) Then
                 Call ReBuildIndex(Split(strSegment, ".")(0), Split(strSegment, ".")(1), strParallel)
             End If
             
             'LOB分区并行执行不会导致表及索引的degree属性被设置，所以不必执行noparallel
         Next
     End If
              

ResizeLine:
    If strRemoveTable & strRemovePARTable & strRemoveLob & strRemoveIndex & strRemovePARIndex & strRemovePARLOB <> "" Then
        RefreshInfo "正在收缩" & strTbsOriginal & "的空间"
        Call ResizeTBS(strTbsOriginal)  '移回原表空间，可能分配了较多空闲空间，所以要收缩
    End If
    
    RefreshInfo "正在收缩" & strTbsTemp & "的空间"
    Call ResizeTBS(strTbsTemp)
        
    
    Exit Sub
errH:
    Call ErrCenter(strSql)
End Sub

Private Sub RefreshInfo(strInfo As String)
'功能：显示并刷新当前操作信息
    lblPrompt.Caption = strInfo
    lblPrompt.Refresh
End Sub

Private Function CheckSuportTable(ByVal strSegment As String)
'功能：检查表是否存在(mdsys用户下存在类似GridFile1044_TAB这种含有小写字母的表段，但在dba_tables中却查不到记录)
    Dim rstmp As ADODB.Recordset, strSql As String
 
    'iot_name不为空的，是IOT索引的溢出表
    'mdsys有一张表SDO_3DTXFMS_TABLE，存在SDO_NUMBER_ARRAY数据类型，导致不能Move
    'Data_Type_Owner为Public的是XMLTYPE
    strSql = "Select 1" & vbNewLine & _
            "From Dba_Tables A" & vbNewLine & _
            "Where Owner = [1] And Table_Name = [2] And Iot_Name Is Null And Not Exists" & vbNewLine & _
            " (Select 1" & vbNewLine & _
            "       From Dba_Tab_Cols B" & vbNewLine & _
            "       Where a.Owner = b.Owner And a.Table_Name = b.Table_Name And Nvl(b.Data_Type_Owner,'PUBLIC') <> 'PUBLIC' And b.Data_Type<> 'XMLTYPE')"

    On Error GoTo errH
    Set rstmp = OpenSQLRecord(strSql, Me.Caption, Split(strSegment, ".")(0), Split(strSegment, ".")(1))

    CheckSuportTable = rstmp.RecordCount > 0
    
    Exit Function
errH:
    Call ErrCenter(strSql)
End Function


Private Function CheckExtent(ByVal bytOpt As opt) As Boolean
    Dim strSegment As String, strPrompt As String
    Dim r1&, c1&, r2&, c2&, r&, c&
    
    If vsfExtents.Row = -1 Or vsfExtents.Col = -1 Then
        MsgBox "请先选中一个单元格再执行本操作", vbInformation, Me.Caption
        Exit Function
    End If
    If mcolCells Is Nothing Then
        MsgBox "请先刷新数据并加载一个存储了数据的单元格再执行本操作", vbInformation, Me.Caption
        Exit Function
    End If
    
    With vsfExtents
        .GetSelection r1, c1, r2, c2
        If r1 = r2 And c1 = c2 Then '仅选中一个单元格时才检查
            strSegment = mcolCells("_" & .Row & "_" & .Col)
            If strSegment = "" Or strSegment = "sys.free" Or cboFiles.ListIndex = -1 Then
                MsgBox "请先选中一个存储了数据的单元格再执行本操作", vbInformation, Me.Caption
                Exit Function
            End If
            mrsLobs.Filter = "Owner='" & Split(strSegment, ".")(0) & "' And Segment_Name='" & Split(strSegment, ".")(1) & "'"
            If mrsLobs.RecordCount > 0 Then
                strSegment = strSegment & "(" & mrsLobs!Table_name & "." & mrsLobs!Column_Name & ")"
            End If
        Else
            strSegment = mcolCells("_" & .Row & "_" & .Col) & "等"
        End If
    End With
        
    If bytOpt = P1回收 Then
        strSegment = "回收(Shrink)一般用于删除大量数据后降低高水标记，以便进行文件收缩操作，操作过程不影响业务,你确定要对" & vbCrLf & vbTab & strSegment & vbCrLf & "进行回收操作吗？"
        
    ElseIf bytOpt = P2重整 Then
        strSegment = "重整(Move Or Rebuild)一般用于移动块的物理位置，操作过程会锁表，并且需要与该对象等量的空闲空间，可能影响业务，请慎重。" & vbCrLf & _
                "Move表之后，相关索引会失效，本操作将会自动重建，可能耗时较长，你确定要对" & vbCrLf & vbTab & strSegment & vbCrLf & "进行重整操作吗？"
    End If
    If MsgBox(strSegment, vbOKCancel + vbDefaultButton1, Me.Caption) = vbCancel Then
        Exit Function
    End If
        
    CheckExtent = True
End Function


Private Sub ExecShrink(ByVal strSegment As String, ByVal strSegment_Type As String, lngErrCount As Long)
'功能：执行回收操作
    Dim strSql As String, strObjName As String
    Dim rstmp As ADODB.Recordset
    Dim blnRow_Movement As Boolean
    
    On Error GoTo errH
    
    If strSegment_Type = "TABLE" Then
        If CheckSuportTable(strSegment) Then
            If Not CheckUnSuportObject(strSegment, "收回(Shrink Space)") And Not CheckFuncIDX(strSegment) Then
                strSql = "Select Row_Movement From All_Tables Where Table_Name = [1] And Owner = [2]"
                Set rstmp = OpenSQLRecord(strSql, Me.Caption, Split(strSegment, ".")(1), Split(strSegment, ".")(0))
                If rstmp.RecordCount = 0 Then
                    Call MsgBox("从视图All_Tables中未找到对象" & strSegment, vbInformation, Me.Caption)
                    Exit Sub
                End If
                If rstmp!Row_Movement = "DISABLED" Then 'enable row movement语句会造成引用表XXX的对象(如存储过程、包、视图等)变为无效
                    strSql = "Alter Table " & strSegment & " Enable Row Movement"
                    gcnOracle.Execute strSql
                    blnRow_Movement = True
                End If
                       
                RefreshInfo "正在对[" & strSegment & "]进行空间收回"
                DoEvents
                                                
                strSql = "Alter Table " & strSegment & " Shrink Space"
                gcnOracle.Execute strSql
                
                If blnRow_Movement Then
                    strSql = "Alter Table " & strSegment & " Disable Row Movement"
                    gcnOracle.Execute strSql
                End If
            Else
                TraceFileWrite "未重整含有Long或Long Raw字段的表:" & strSegment
            End If
        Else
            TraceFileWrite "未重整IOT索引的溢出表或含有自定义字段的表:" & strSegment
        End If
    ElseIf strSegment_Type = "LOBSEGMENT" Then
        mrsLobs.Filter = "Owner='" & Split(strSegment, ".")(0) & "' And Segment_Name='" & Split(strSegment, ".")(1) & "'"
        If mrsLobs.RecordCount > 0 Then
            If Not CheckUnSuportObject(mrsLobs!Owner & "." & mrsLobs!Table_name, "收回(Shrink Space)") Then
                RefreshInfo "正在对[" & mrsLobs!Table_name & "." & mrsLobs!Column_Name & "]进行空间回收"
                DoEvents
                strSql = "ALTER TABLE " & mrsLobs!Owner & "." & mrsLobs!Table_name & " MODIFY LOB (" & mrsLobs!Column_Name & ") (SHRINK SPACE)"
                gcnOracle.Execute strSql
            Else
                TraceFileWrite "未重整含有Long或Long Raw字段的表:" & strSegment
            End If
        Else
            RefreshInfo "在视图Dba_Lobs中未找到LOB对象" & strSegment & "。"
        End If
        
    ElseIf strSegment_Type = "INDEX" Then
        
        If Not CheckIOT(strSegment) Then
            If Not CheckUnSuportObject(strSegment, "收回(Shrink Space)", True) Then
                RefreshInfo "正在对[" & strSegment & "]进行空间收回"
                DoEvents
                strSql = "Alter Index " & strSegment & " Shrink Space"
                gcnOracle.Execute strSql
            Else
                TraceFileWrite "未重整含有Long或Long Raw字段的表:" & strSegment
            End If
        Else
            strObjName = GetIOTName(strSegment)
            strSql = "Alter Table " & strObjName & " Shrink Space"
            gcnOracle.Execute strSql
        End If
    ElseIf strSegment_Type <> " " Then
        RefreshInfo strSegment & ",不支持的对象类型：" & strSegment_Type
    End If
    
    Exit Sub
errH:
    If InStr(Err.Description, "ORA-10638:") > 0 Then    'Index status is invalid
        gcnOracle.Execute Replace(strSql, "Shrink Space", "Rebuild nologging")
        Resume
    End If
    
    Call TraceFileWrite(Err.Description & vbCrLf & strSql, LTT_OnlyTime)
    lngErrCount = lngErrCount + 1
    
    If blnRow_Movement Then
        On Error Resume Next    '避免并发环境可能无法锁定资源
        strSql = "Alter Table " & strSegment & " Disable Row Movement"
        gcnOracle.Execute strSql
        Err.Clear
    End If
End Sub

Private Sub cmdShrink_Click()
'功能：执行表或索引的空闲空间收回(Shrink Space)
    Dim strSegment_Type As String, strSegment As String, strSegmentPre As String, strSegmentAll As String

    Dim r1 As Long, c1 As Long, r2 As Long, c2 As Long, r As Long, c As Long
    Dim datBegin As Date, strTime As String, lngErrCount As Long
        
    If CheckExtent(P1回收) = False Then Exit Sub
        
    strSegment_Type = Trim(lblPrompt.Tag)
    If strSegment_Type <> "" And InStr(",TABLE,INDEX,LOBSEGMENT,", "," & strSegment_Type & ",") = 0 Then
        Call MsgBox("仅支持对表或索引进行空闲空间收回，不支持的数据类型：" & strSegment_Type, vbInformation, Me.Caption)
        Exit Sub
    End If
    
    datBegin = GetCurrentdate
    
    Call SetCommandEnable(0)
    vsfExtents.GetSelection r1, c1, r2, c2
    For r = r2 To r1 Step -1
        For c = c2 To c1 Step -1
            strSegment = mcolCells("_" & r & "_" & c)     '含所有者
            strSegment_Type = CStr(vsfExtents.Cell(flexcpData, r, c))
            
            If strSegment & "|" & strSegment_Type <> strSegmentPre Then
                If InStr(strSegmentAll & ",", "," & strSegment & "|" & strSegment_Type & ",") = 0 Then
                    mrsExtents.Filter = "Row=" & r & " And Col=" & c
                    If mrsExtents.RecordCount > 0 Then
                        ExecShrink strSegment, mrsExtents!Segment_Type, lngErrCount
                    End If
                    strSegmentAll = strSegmentAll & "," & strSegment & "|" & strSegment_Type
                End If
                strSegmentPre = strSegment & "|" & strSegment_Type
            End If
        Next
    Next
    
    '未改变数据文件大小，不用刷新表空间及数据文件列表
    Call LoadExtents(vsfTbs.TextMatrix(vsfTbs.Row, vsfTbs.ColIndex("名称")), Val(cboFiles.ItemData(cboFiles.ListIndex)))
    
    If strSegment <> "" Then
        mrsExtents.Filter = "Segment_Name='" & Split(strSegment, ".")(1) & "' And Owner='" & Split(strSegment, ".")(0) & "'"
        If mrsExtents.RecordCount > 0 Then
            vsfExtents.SetFocus
            vsfExtents.Select mrsExtents!Row, mrsExtents!Col
            vsfExtents.TopRow = vsfExtents.Row
        End If
    End If
        
    Call ShowSpendTime(datBegin, "回收", lngErrCount)
    
    Call SetCommandEnable(1)
End Sub

Private Sub ShowSpendTime(ByVal datBegin As Date, ByVal strOpt As String, ByVal lngErrCount As Long)
'功能：显示耗时信息
    Dim strTime As String
    
    strTime = GetTimeString(datBegin, GetCurrentdate)
    strTime = strOpt & "操作完成，耗时：" & strTime & "。"
    
    TraceFileClose
    
    If lngErrCount <> 0 Then
        strTime = strTime & vbCrLf & "共发生" & lngErrCount & "个错误，详情请查看日志文件" & gstrTraceFile & "。"
        MsgBox strTime, vbInformation, gstrSysName
        On Error Resume Next
        Shell "notepad " & gstrTraceFile
        Err.Clear
    Else
        lblPrompt.Caption = strTime
    End If

End Sub

Private Sub cmdShrinkAll_Click()
'功能：对当前表空间的所有文件中的对象执行空间回收
    Dim rstmp As ADODB.Recordset
    
    Dim strSegment_Type As String, strSegment As String, strSegmentPre As String, strSegmentAll As String
    Dim strTbsOriginal As String, strPrompt As String
    Dim datBegin As Date, strTime As String, lngErrCount As Long
    Dim i As Long
    
    
    If CheckSelFile = False Then Exit Sub
        
    strTbsOriginal = vsfTbs.TextMatrix(vsfTbs.Row, vsfTbs.ColIndex("名称"))
    
    strPrompt = "回收(Shrink)操作的执行不影响业务运行，一般用于删除大量数据后降低高水位。" & _
            "相对于Move操作，它不需要额外的空闲空间，但耗时较长，并且仍然可能在文件末尾存储数据，导致无法收缩文件。" & vbCrLf & vbCrLf & _
            "你确定要对表空间" & strTbsOriginal & "的" & cboFiles.ListCount & "个文件中的所有对象进行回收操作吗？"
            
    If MsgBox(strPrompt, vbOKCancel + vbDefaultButton1, Me.Caption) = vbCancel Then
        Exit Sub
    End If
        
    Me.Refresh  '避免表格上的残影
    datBegin = GetCurrentdate
    
    Call SetCommandEnable(0)
        
    For i = 0 To cboFiles.ListCount - 1
        
        RefreshInfo "正在查询[" & cboFiles.List(i) & "]文件中的所有对象"
        Set rstmp = GetExtents(strTbsOriginal, Val(cboFiles.ItemData(i)))
        
        Do While Not rstmp.EOF
            strSegment = rstmp!Full_Segment_Name
            strSegment_Type = rstmp!Segment_Type
                            
            If strSegment & "|" & strSegment_Type <> strSegmentPre Then
                If InStr(strSegmentAll & ",", "," & strSegment & "|" & strSegment_Type & ",") = 0 Then
                    ExecShrink strSegment, strSegment_Type, lngErrCount
                
                    strSegmentAll = strSegmentAll & "," & strSegment & "|" & strSegment_Type
                End If
                strSegmentPre = strSegment & "|" & strSegment_Type
            End If
            rstmp.MoveNext
        Loop
    Next
        
    '刷新数据
    Call RefreshData
        
    Call ShowSpendTime(datBegin, "回收", lngErrCount)
    
    Call SetCommandEnable(1)
End Sub

Private Sub Form_load()
    Dim strCol As String, i As Long
    
    strCol = "行,300,1;状态;名称,1200,1;大小,650,1"
    Call InitTable(vsfTbs, strCol)
    vsfTbs.FixedCols = 1
    
    strCol = ""
    For i = 0 To CONCOLS
        If strCol = "" Then
            strCol = i & ",550,1"
        Else
            strCol = strCol & ";" & i & ",280,4"
        End If
    Next

    Call InitTable(vsfExtents, strCol)
    vsfExtents.FixedCols = 1
    vsfExtents.Rows = vsfExtents.FixedRows + 1
    vsfExtents.TextMatrix(0, 0) = "行\列"
    
    
    Call LoadTablespaces
    
    vsfTbs.Editable = flexEDNone
    vsfExtents.Editable = flexEDNone
    cmdLOBGO.Visible = False
        
End Sub


Private Sub cmdResizeAll_Click()
    Call ResizeAll
End Sub

Private Sub cmdResizeTemp_Click()
'收缩临时表空间
    Call ResizeTemp
End Sub

Private Sub cmdResizeUndo_Click()
'收缩Undo表空间
    Call frmResizeUndo.ShowMe(Me)
End Sub


Private Sub LoadTablespaces()
    Dim rstmp As ADODB.Recordset, strSql  As String
    Dim i As Long, lngStart As Long
    
    strSql = "Select a.Status, a.Tablespace_Name, a.Block_Size, Round(Sum(b.Bytes) / 1024 / 1024, 2) Tsize , Max(Decode(b.autoextensible,'YES',0,1)) as autoextensible" & vbNewLine & _
            "From Dba_Tablespaces A, Dba_Data_Files B" & vbNewLine & _
            "Where a.Contents = 'PERMANENT' And a.Tablespace_Name = b.Tablespace_Name And b.Online_status in('ONLINE','SYSTEM')" & vbNewLine & _
            "Group By a.Tablespace_Name, a.Status, a.Block_Size" & vbNewLine & _
            "Order By 4 Desc"

    On Error GoTo errH
    Set rstmp = OpenSQLRecord(strSql, Me.Caption)
    
    With vsfTbs
        .Redraw = flexRDNone
        lngStart = .FixedRows
        .Rows = lngStart
        .Rows = lngStart + rstmp.RecordCount
        For i = lngStart To rstmp.RecordCount
            If rstmp!autoextensible = 1 Then
                .Cell(flexcpBackColor, i, .ColIndex("行"), i, .ColIndex("大小")) = OFF_颜色
                .Cell(flexcpData, i, .ColIndex("大小")) = "NO"
            End If
            .TextMatrix(i, .ColIndex("行")) = i
            .TextMatrix(i, .ColIndex("状态")) = rstmp!Status
            .TextMatrix(i, .ColIndex("名称")) = rstmp!Tablespace_Name
            
            If Val("" & rstmp!Tsize) > 1024 Then
                .TextMatrix(i, .ColIndex("大小")) = Round(rstmp!Tsize / 1024, 2) & "G"
            Else
                .TextMatrix(i, .ColIndex("大小")) = rstmp!Tsize & "M"
            End If
            
            .RowData(i) = Val(rstmp!Block_Size)
            rstmp.MoveNext
        Next
        .Redraw = flexRDDirect
    End With

    Exit Sub
errH:
    Call ErrCenter(strSql)
End Sub




Private Sub Form_Resize()
    
    On Error Resume Next
    
    txtFind.Left = Me.ScaleWidth - txtFind.Width - 60
    lblFind.Left = txtFind.Left - lblFind.Width - 60
    cmdLOBGO.Left = lblFind.Left - cmdLOBGO.Width - 60
    chkFree.Left = cmdLOBGO.Left - chkFree.Width - 60
    cboFiles.Width = chkFree.Left - cboFiles.Left - 60
        
    vsfExtents.Width = Me.ScaleWidth - vsfExtents.Left - 60
    vsfTbs.Height = Me.ScaleHeight - vsfTbs.Top - picBottom.Height - 60
    vsfExtents.Height = vsfTbs.Height
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsExtents = Nothing
    Set mcolCells = Nothing
    Set mrsLobs = Nothing
End Sub

Private Sub LoadLobs(ByVal strTbs As String)
'功能：读取当前表空间的Lob段信息
    Dim strSql As String
 
    strSql = "Select Table_Name, TableSpace_Name, Column_Name, Owner, Segment_Name, Index_Name From Dba_Lobs Where Tablespace_Name = [1]"
    On Error GoTo errH
    Set mrsLobs = OpenSQLRecord(strSql, Me.Caption, strTbs)

    Exit Sub
errH:
    Call ErrCenter(strSql)
End Sub

Private Function GetExtents(ByVal strTbs As String, ByVal lngFile As Long, Optional ByVal blnOnlyFree As Boolean) As ADODB.Recordset
'功能：获取指定表空间和文件id的块集合
    Dim strSql  As String
    
    If blnOnlyFree Then
        strSql = "Select File_Id,-1 as Extent_ID, Block_Id as First_Block, Block_Id + Blocks - 1 as Last_Block,Blocks, 'free' as Segment_Name, 'sys.free' as Full_Segment_Name, ' ' as Segment_Type,' ' as Owner" & vbNewLine & _
            "From Dba_Free_Space A" & vbNewLine & _
            "Where Tablespace_Name = [1] And a.File_Id = [2]" & vbNewLine & _
            "Order By First_Block"
    Else
        strSql = "Select a.File_Id,a.Extent_ID, a.Block_Id First_Block, a.Block_Id + a.Blocks - 1 Last_Block,a.Blocks, a.Segment_Name, a.Owner || '.' || a.Segment_Name as Full_Segment_Name, b.Segment_Type, a.Owner" & vbNewLine & _
            "From Dba_Extents A, Dba_Segments B" & vbNewLine & _
            "Where a.Tablespace_Name = [1] And a.File_Id = [2] And a.Segment_Name = b.Segment_Name And a.Owner = b.Owner" & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select File_Id,-1, Block_Id, Block_Id + Blocks - 1,Blocks, 'free', 'sys.free' as Full_Segment_Name, ' ',' '" & vbNewLine & _
            "From Dba_Free_Space A" & vbNewLine & _
            "Where Tablespace_Name = [1] And a.File_Id = [2]" & vbNewLine & _
            "Order By First_Block"
    End If
    On Error GoTo errH
    Set GetExtents = OpenSQLRecord(strSql, Me.Caption, strTbs, lngFile)
    Exit Function
    
errH:
    Call ErrCenter(strSql)
End Function

Private Sub LoadExtents(ByVal strTbs As String, ByVal lngFile As Long)
'功能：加载Extents到单元格
    Dim rstmp As ADODB.Recordset, strSegment As String, strPreSegment As String, strFullSegment As String
    Dim i As Long, j As Long, n As Long, lngStart As Long, lngRows As Long
    Dim lngCells As Long, lngFixedCols As Long
    Dim blnFree As Boolean, blnSameCell As Boolean, strFirst As String
    
    lblPrompt.Caption = "正在读取数据块信息......"
    lblPrompt.Refresh
    
    Set rstmp = GetExtents(strTbs, lngFile, chkFree.Value = 1)
            
    Call InitmrsExtents
    
    If rstmp.RecordCount = 0 Then
        lblPrompt.Caption = lblPrompt.ToolTipText
        lblPrompt.Refresh
        vsfExtents.Rows = vsfExtents.FixedRows
        Exit Sub
    End If
    
    lblPrompt.Caption = "正在加载数据块信息......"
    lblPrompt.Refresh
    lngFixedCols = vsfExtents.FixedCols
    lngStart = vsfExtents.FixedRows
    
    
    '先计算出行数,用于显示进度
    j = lngFixedCols
    lngRows = lngStart
    Do While Not rstmp.EOF
        lngCells = rstmp!blocks \ CONBLOCKS '取整
        If rstmp!blocks <> lngCells * CONBLOCKS Then lngCells = lngCells + 1
        
        For n = 1 To lngCells
            j = j + 1
            If j > CONCOLS Then '换行
                lngRows = lngRows + 1
                j = lngFixedCols
            End If
        Next
        rstmp.MoveNext
    Loop
    rstmp.MoveFirst
    
    vsfExtents.Redraw = flexRDNone  '避免触发事件vsfExtents_AfterRowColChange
    vsfExtents.Rows = lngStart
    vsfExtents.Redraw = flexRDDirect
    vsfExtents.ToolTipText = ""
    vsfExtents.Refresh
    
    vsfExtents.Redraw = flexRDNone
    vsfExtents.Rows = lngStart + lngRows
    vsfExtents.Redraw = flexRDDirect
    
        
    With vsfExtents
        .Redraw = flexRDNone
                
        i = lngStart
        j = .FixedCols
        If i > 0 Then .TextMatrix(1, 0) = 1
        
        Do While Not rstmp.EOF
            strSegment = rstmp!Segment_Name
            blnFree = (strSegment = "free")
            strFullSegment = rstmp!Full_Segment_Name
                                    
            strFirst = Mid$(strSegment, 1, 1)
            If strPreSegment <> strSegment & "|" & rstmp!Segment_Type Then
                blnSameCell = Mid$(strPreSegment, 1, 1) = strFirst
            Else
                blnSameCell = False
            End If
            
            lngCells = rstmp!blocks \ CONBLOCKS '取整
            If rstmp!blocks <> lngCells * CONBLOCKS Then lngCells = lngCells + 1
           
            For n = 1 To lngCells
                If blnFree Then
                    .Cell(flexcpBackColor, i, j) = &HCCEDC7 '空闲空间
                    If n = 1 Then .TextMatrix(i, j) = "B"
                Else
                    .TextMatrix(i, j) = strFirst
                    .Cell(flexcpData, i, j) = CStr(rstmp!Segment_Type)
                End If
                mcolCells.Add strFullSegment, "_" & i & "_" & j
               
                '第一个字相同，但对象不同，用加粗来区别
                If blnSameCell Then .Cell(flexcpFontItalic, i, j) = True
                
                mrsExtents.AddNew Array("Row", "Col", "Segment_Name", "Extent_ID", "First_Block", "Blocks", "Last_Block", "Segment_Type", "Owner"), _
                            Array(i, j, strSegment, rstmp!Extent_ID, rstmp!First_Block, rstmp!blocks, rstmp!Last_Block, rstmp!Segment_Type, rstmp!Owner)
                               
                j = j + 1
                If j > CONCOLS Then '换行
                    j = lngFixedCols
                    
                    i = i + 1
                   .TextMatrix(i, 0) = i   '行号
                   
                   If i Mod 100 = 0 Then
                     DoEvents
                     lblPrompt.Caption = "正在加载信息(" & i & "/" & lngRows & ")"
                   End If
                End If
           Next
           strPreSegment = strSegment & "|" & rstmp!Segment_Type
           rstmp.MoveNext
        Loop
        
        '剩余的空单元格加上空值以避免从集合取值时出错
        For n = j To CONCOLS
            mcolCells.Add "", "_" & i & "_" & n
        Next
        
        .Redraw = flexRDDirect
    End With
    lblPrompt.Caption = lblPrompt.ToolTipText
        
End Sub


Private Sub picBottom_Resize()
    On Error Resume Next
    lineBottom(0).X2 = picBottom.ScaleWidth
    lineBottom(1).X2 = lineBottom(0).X2
        
    lblPrompt.Width = picBottom.ScaleWidth - lblPrompt.Left - 60
End Sub

Private Sub picTop_Resize()
    On Error Resume Next
    lineTop(0).X2 = picTop.ScaleWidth
    lineTop(1).X2 = lineTop(0).X2
    
    fraTopCmd.Left = picTop.ScaleWidth - fraTopCmd.Width - 60
End Sub

Private Sub txtFind_LostFocus()
    
    If Trim(txtFind.Text) = "" Then
        txtFind.Text = txtFind.Tag
        txtFind.ForeColor = &H808080
    End If
End Sub


Private Sub txtFind_GotFocus()
    If txtFind.Text = txtFind.Tag Then
        txtFind.Text = ""
        txtFind.ForeColor = vbBlack
    Else
        txtFind.SelStart = 0
        txtFind.SelLength = Len(txtFind.Text)
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Not mrsExtents Is Nothing Then
            If InStr(txtFind.Text, "*") > 0 Then
                mrsExtents.Filter = "Segment_Name Like '" & UCase(Trim(txtFind.Text)) & "'"
            Else
                mrsExtents.Filter = "Segment_Name='" & UCase(Trim(txtFind.Text)) & "'"
            End If
            If mrsExtents.RecordCount > 0 Then
                vsfExtents.SetFocus
                vsfExtents.Select mrsExtents!Row, mrsExtents!Col
                vsfExtents.TopRow = vsfExtents.Row
            Else
                lblPrompt.Caption = "没有找到匹配的表或索引。"
                txtFind.SetFocus
                txtFind_GotFocus
            End If
        Else
            lblPrompt.Caption = "没有找到匹配的表或索引。"
            txtFind.SetFocus
            txtFind_GotFocus
        End If
    ElseIf KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub



Private Sub cmdLOBGO_Click()
'功能：根据LOB索引或分区索引 定位到LOB对象
    Dim strObjName As String, strSegment As String, strSegment_Type As String
    Dim i As Long, j As Long
    
    With vsfExtents
        If .Row < 0 Or .Col < 0 Then Exit Sub
        
        strSegment_Type = .Cell(flexcpData, .Row, .Col)
        strSegment = mcolCells("_" & .Row & "_" & .Col)
        If strSegment = "" Then Exit Sub
        
        If strSegment_Type = "LOBINDEX" Or strSegment_Type = "INDEX PARTITION" Then
            strObjName = GetLOBNameByIndex(strSegment)
        End If
        
        If strObjName <> "" Then
            For i = .FixedRows To .Rows - 1
                For j = .FixedCols To .Cols - 1
                    If strObjName = mcolCells("_" & i & "_" & j) Then
                        .Select i, j
                        .TopRow = i
                        .SetFocus
                        strObjName = ""
                        Exit Sub
                    End If
                Next
            Next
            If strObjName <> "" Then Call MsgBox("未找到" & strObjName, vbInformation)
        End If
    End With
End Sub

Private Sub vsfExtents_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsfExtents
        Dim strSegment As String, i As Long, lngBlockSize As Long, strSegment_Type As String
        
        If Me.Visible = False Or .Redraw = flexRDNone Or mcolCells Is Nothing Or vsfTbs.Enabled = False Then Exit Sub
        
        .Redraw = flexRDNone
        '先去掉之前选中的段的背景色
        If OldRow > 0 And OldCol > 0 Then
            strSegment = mcolCells("_" & OldRow & "_" & OldCol)
            If strSegment <> "" Then
                strSegment_Type = .Cell(flexcpData, OldRow, OldCol)
                mrsExtents.Filter = "Segment_Name='" & Split(strSegment, ".")(1) & "' And Owner='" & Split(strSegment, ".")(0) & "' And Segment_Type='" & strSegment_Type & "'"
                For i = 1 To mrsExtents.RecordCount
                    If mrsExtents!Segment_Name = "free" Then
                        .Cell(flexcpBackColor, mrsExtents!Row, mrsExtents!Col) = &HCCEDC7 '空闲空间
                    Else
                        .Cell(flexcpBackColor, mrsExtents!Row, mrsExtents!Col) = &H80000005 '白色
                    End If
                    .Cell(flexcpForeColor, mrsExtents!Row, mrsExtents!Col) = vbBlack
                    mrsExtents.MoveNext
                Next
            End If
        End If
                
        .Redraw = flexRDDirect
        
        
        '再设置当前选中段的背景色
        .Redraw = flexRDNone
        cmdLOBGO.Visible = False
        
        strSegment = mcolCells("_" & NewRow & "_" & NewCol)
        lblPrompt.Tag = ""
        If strSegment <> "" Then
            strSegment_Type = .Cell(flexcpData, NewRow, NewCol)
            
            If strSegment_Type = "LOBINDEX" Then
                cmdLOBGO.Visible = True
                cmdLOBGO.Caption = "定位到LOB"
            ElseIf strSegment_Type = "INDEX PARTITION" Then
                If CheckLOBIndex(strSegment) Then
                    cmdLOBGO.Visible = True
                    cmdLOBGO.Caption = "定位到LOB"
                End If
            End If
            
            mrsExtents.Filter = "Row=" & NewRow & " And Col=" & NewCol
            If mrsExtents.RecordCount > 0 Then
                If mrsExtents!Segment_Type = "LOBSEGMENT" Then
                    mrsLobs.Filter = "Owner='" & mrsExtents!Owner & "' And Segment_Name='" & mrsExtents!Segment_Name & "'"
                    If mrsLobs.RecordCount > 0 Then .ToolTipText = mrsLobs!Table_name & "." & mrsLobs!Column_Name
                
                ElseIf mrsExtents!Segment_Type = "LOBINDEX" Then
                    mrsLobs.Filter = "Owner='" & mrsExtents!Owner & "' And Index_Name='" & mrsExtents!Segment_Name & "'"
                    If mrsLobs.RecordCount > 0 Then .ToolTipText = mrsLobs!Table_name & "." & mrsLobs!Column_Name & "(Index)"
                Else
                    .ToolTipText = strSegment & "(一个单元格包含" & CONBLOCKS & "个块)"
                End If
                
                lngBlockSize = Val(vsfTbs.RowData(vsfTbs.Row))
                If lngBlockSize = 0 Then lngBlockSize = 8192
                
                If strSegment = "sys.free" Then
                    lblPrompt.Caption = "已格式化的空闲空间，" & mrsExtents!blocks & "块：从" & Round(mrsExtents!First_Block * 8192 / 1024 / 1024, 2) & _
                                        "M到" & Round(mrsExtents!Last_Block * lngBlockSize / 1024 / 1024, 2) & "M"
                Else
                    lblPrompt.Caption = mrsExtents!Segment_Type & "：" & strSegment & "，Extent_ID：" & mrsExtents!Extent_ID & "(" & mrsExtents!blocks & "块，从" & _
                                        Round(mrsExtents!First_Block * lngBlockSize / 1024 / 1024, 2) & "M到" & Round(mrsExtents!Last_Block * lngBlockSize / 1024 / 1024, 2) & "M)"
                    lblPrompt.Tag = mrsExtents!Segment_Type
                End If
            End If
            
            mrsExtents.Filter = "Segment_Name='" & Split(strSegment, ".")(1) & "' And Owner='" & Split(strSegment, ".")(0) & "' And Segment_Type='" & strSegment_Type & "'"
            For i = 1 To mrsExtents.RecordCount
                .Cell(flexcpBackColor, mrsExtents!Row, mrsExtents!Col) = &H8000000D     '蓝色
                .Cell(flexcpForeColor, mrsExtents!Row, mrsExtents!Col) = &H80000005
                mrsExtents.MoveNext
            Next
        Else
            lblPrompt.Caption = lblPrompt.ToolTipText
        End If
        .Redraw = flexRDDirect
    End With
End Sub

Private Sub vsfExtents_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngRow As Long
    Dim lngCol As Long
    
    If Me.Visible = False Or vsfExtents.Redraw = flexRDNone Or vsfTbs.Enabled = False Then Exit Sub
    
    lngRow = vsfExtents.MouseRow
    lngCol = vsfExtents.MouseCol
    If lngRow > 0 And lngCol > 0 And Not mcolCells Is Nothing Then
       If (lngRow <> mlngRowPre Or lngCol <> mlngColPre) And lngRow <> vsfExtents.Row And lngCol <> vsfExtents.Col Then
           vsfExtents.ToolTipText = mcolCells("_" & lngRow & "_" & lngCol) & "(一个单元格包含" & CONBLOCKS & "个块)"
           mlngRowPre = lngRow
           mlngColPre = lngCol
       End If
    Else
        vsfExtents.ToolTipText = ""
    End If
End Sub

Private Sub vsfTbs_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    If Me.Visible And NewRowSel <> OldRowSel And vsfTbs.Redraw <> flexRDNone Then
        vsfTbs.Refresh
        Call LoadFiles(vsfTbs.TextMatrix(NewRowSel, vsfTbs.ColIndex("名称")))

        If cboFiles.ListCount < 2 Then
            cboFiles.ListIndex = 0
        Else
            vsfExtents.Redraw = flexRDNone '避免触发事件vsfExtents_AfterRowColChange
            vsfExtents.Rows = vsfExtents.FixedRows
            vsfExtents.Redraw = flexRDDirect
            vsfExtents.ToolTipText = ""
            vsfExtents.Refresh
        End If
        
        Call LoadLobs(vsfTbs.TextMatrix(NewRowSel, vsfTbs.ColIndex("名称")))
        
        If vsfTbs.Cell(flexcpData, NewRowSel, vsfTbs.ColIndex("大小")) = "NO" Then
            lblPrompt.Caption = "所选表空间中存在自增长属性为NO的数据文件。"
        End If
    End If
End Sub

Private Sub LoadFiles(strTbs As String)
    Dim rstmp As ADODB.Recordset, strSql  As String
    Dim i As Long, lngStart As Long
    
    strSql = "Select a.File_Name, a.File_Id, Round(a.Bytes / 1024 / 1024) As Fsize, Round(Nvl(Sum(b.Bytes),0) / 1024 / 1024) As Free_Size , a.autoextensible " & vbNewLine & _
            "From Dba_Data_Files A, Dba_Free_Space B" & vbNewLine & _
            "Where a.Tablespace_Name = [1] And a.File_Id = b.File_Id(+) And a.Tablespace_Name = b.Tablespace_Name(+) And a.Online_status in('ONLINE','SYSTEM')" & vbNewLine & _
            "Group By a.File_Name, a.File_Id, a.Bytes,a.autoextensible" & vbNewLine & _
            "Order By a.File_Id"

    On Error GoTo errH
    Set rstmp = OpenSQLRecord(strSql, Me.Caption, strTbs)
    
    cboFiles.Clear
    cboFiles.Tag = ""
    For i = 1 To rstmp.RecordCount
        cboFiles.AddItem rstmp!File_Name & "(占用" & rstmp!fsize & "M,空闲" & rstmp!Free_Size & "M" & IIf(rstmp!autoextensible & "" <> "YES", ",不自动扩展", "") & ")"
        cboFiles.ItemData(cboFiles.NewIndex) = Val(rstmp!File_Id)
        rstmp.MoveNext
    Next
    
    Exit Sub
errH:
    Call ErrCenter(strSql)
End Sub


Private Sub InitmrsExtents()
    
    Set mcolCells = New Collection
    
    Set mrsExtents = New ADODB.Recordset
    mrsExtents.Fields.Append "Row", adBigInt
    mrsExtents.Fields.Append "Col", adBigInt
    mrsExtents.Fields.Append "Owner", adVarChar, 20
    mrsExtents.Fields.Append "Segment_Name", adVarChar, 100
    mrsExtents.Fields.Append "Segment_Type", adVarChar, 20
    
    mrsExtents.Fields.Append "Extent_ID", adBigInt
    mrsExtents.Fields.Append "Blocks", adBigInt
    mrsExtents.Fields.Append "First_Block", adBigInt
    mrsExtents.Fields.Append "Last_Block", adBigInt
    
    mrsExtents.CursorLocation = adUseClient
    mrsExtents.LockType = adLockOptimistic
    mrsExtents.CursorType = adOpenStatic
    mrsExtents.Open
End Sub


Private Sub ResizeAll()
'功能：收缩所有数据文件
    Dim strErr As String
    Dim rstmp As ADODB.Recordset, rsSize As ADODB.Recordset
    Dim lngBlockSize As Long, lngSumSize As Long
    
    If MsgBox("你确定要收缩所有数据文件吗？" & vbCrLf & vbCrLf & "此操作比较耗时，业务运行期间请谨慎执行此操作！", vbYesNo + vbQuestion + vbDefaultButton2, "确认收缩") = vbNo Then
        lblPrompt.Caption = "操作被取消。"
        Call SetCommandEnable(1)
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    
    Call SetCommandEnable(0)
    '获取Block_size大小
    gstrSQL = "select value from v$parameter where name = 'db_block_size'"
    Set rstmp = OpenSQLRecord(gstrSQL, Me.Caption)
    lngBlockSize = Val("" & rstmp!Value)
    
    '记录执行操作语句
    lblPrompt.Caption = "正在查询待收缩的数据文件。"
    lblPrompt.Refresh
    gstrSQL = "Select File_Name,'alter database datafile ''' || Trim(File_Name) || ''' resize ' || Ceil((Nvl(Hwm, 1) * " & lngBlockSize & ") / 1024 / 1024+10) || 'm' Cmd" & vbNewLine & _
            "From Dba_Data_Files A, (Select File_Id, Max(Block_Id + Blocks ) Hwm From Dba_Extents Group By File_Id) B" & vbNewLine & _
            "Where a.File_Id = b.File_Id(+) And Exists(Select 1 From Dba_Tablespaces C Where a.Tablespace_Name = c.Tablespace_Name And c.Status = 'ONLINE' And Contents != 'UNDO')" & vbNewLine & _
            "      And Ceil(Blocks * " & lngBlockSize & " / 1024 / 1024) - Ceil((Nvl(Hwm, 1) * " & lngBlockSize & ") / 1024 / 1024) > 10"
    Set rstmp = OpenSQLRecord(gstrSQL, Me.Caption)
    If rstmp.RecordCount = 0 Then
        Call MsgBox("没有要收缩数据文件！", vbInformation, "收缩数据文件")
        lblPrompt.Caption = lblPrompt.ToolTipText
        Call SetCommandEnable(1)
        Screen.MousePointer = vbDefault
        Exit Sub
    Else
    
        If MsgBox("共有" & rstmp.RecordCount & "个待收缩的数据文件，你确定要收缩这些数据文件吗？", vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
            
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    lblPrompt.Caption = "开始进行收缩操作。"
    lblPrompt.Refresh
    
    '执行操作
    '1.记录收缩前的大小
    gstrSQL = "Select Trunc(Sum(Bytes) / 1024 / 1024) Mb_Size From Dba_Data_Files"
    Set rsSize = OpenSQLRecord(gstrSQL, Me.Caption)
    lngSumSize = rsSize!Mb_Size
    
    On Error Resume Next
    strErr = ""
    While Not rstmp.EOF
        lblPrompt.Caption = "正在收缩：" & rstmp!File_Name
        lblPrompt.Refresh
        gstrSQL = rstmp!cmd
        gcnOracle.Execute gstrSQL
        
        If Err.Number <> 0 Then
            strErr = strErr & vbCrLf & rstmp!cmd & "，错误：" & Err.Description
            Err.Clear
        End If
        
        rstmp.MoveNext
    Wend
    
    Call RefreshData
    
    '2.记录收缩后的总大小
    gstrSQL = "Select Trunc(Sum(Bytes) / 1024 / 1024) Mb_Size From Dba_Data_Files"
    Set rsSize = OpenSQLRecord(gstrSQL, Me.Caption)
    lngSumSize = lngSumSize - rsSize!Mb_Size
        
    If strErr <> "" Then
        MsgBox "错误信息：" & strErr, vbExclamation
        lblPrompt.Caption = lblPrompt.ToolTipText
    Else
        lblPrompt.Caption = "操作完成，共收缩了" & lngSumSize & "M的空间。"
    End If
    
    Screen.MousePointer = vbDefault
    
    Call SetCommandEnable(1)
End Sub

Private Sub ResizeTemp()
    Dim strError As String, strVersion As String, strTbsInfo As String
    Dim rstmp As ADODB.Recordset, strFile As String
    Dim strSize As String, lngMax As Long, lngPos As Long
    
    strVersion = getVersion
    If strVersion = "" Then
        Exit Sub
    End If
    
    Call SetCommandEnable(0)

    On Error GoTo errH
    gstrSQL = "Select Tablespace_Name, File_Name, Trunc(Bytes / 1024 / 1024) Siz" & vbNewLine & _
            "From Dba_Temp_Files" & vbNewLine & _
            "Where Bytes / 1024 / 1024 > 10" & vbNewLine & _
            "Order By Tablespace_Name, File_Name"
    Set rstmp = OpenSQLRecord(gstrSQL, Me.Caption)
    
    If rstmp.RecordCount <> 0 Then
        While Not rstmp.EOF
            lngPos = InStrRev(rstmp!File_Name, "\")
            If lngPos = 0 Then  '兼容LINUX系统
                lngPos = InStrRev(rstmp!File_Name, "/")
            End If
            strFile = Mid(rstmp!File_Name, lngPos + 1)
            strTbsInfo = strTbsInfo & strFile & "：" & rstmp!Siz & "M" & vbCrLf
            If rstmp!Siz > lngMax Then lngMax = rstmp!Siz
            rstmp.MoveNext
        Wend
        strTbsInfo = "当前临时表空间文件：" & vbCrLf & vbCrLf & strTbsInfo
        
        '获取重置后的大小
input_line:
        strSize = Trim(InputBox(strTbsInfo & vbCrLf & vbCrLf & "请输入收缩后的数据文件大小(单位M)，小于等于指定值的不收缩", "收缩临时表空间"))
        If strSize = "" Then
            Call SetCommandEnable(1)
            Exit Sub
        Else
            strError = ""
            If Not IsNumeric(strSize) Then
                strError = "请重新输入数字"
            ElseIf Val(strSize) <= 0 Then
                strError = "请重新输入大于零的数字"
            ElseIf Val(strSize) >= lngMax Then
                strError = "请重新输入小于" & lngMax & "的数字。"
            ElseIf InStr(strSize, ".") > 0 Then
                strError = "请重新输入不含小数的数字"
            End If
            
            If strError <> "" Then
                MsgBox strError, vbInformation, gstrSysName
                GoTo input_line
            End If
        End If
        
        On Error Resume Next
        strError = ""
        strTbsInfo = ""
        lblPrompt.Caption = lblPrompt.ToolTipText
        rstmp.MoveFirst
        rstmp.Filter = "Siz>" & strSize
        While Not rstmp.EOF
            lblPrompt.Caption = "正在收缩临时表空间 " & rstmp!Tablespace_Name & "。"
            lblPrompt.Refresh
            If strVersion = 11 Then
                '一个表空间有多个数据文件，11GR1是按表空间来收缩的
                '也可以按数据文件逐个收缩: alter tablespace temp shrink tempfile '/u01/app/oracle/oradata/anqing/temp01.dbf' keep 300M;
                If rstmp!Tablespace_Name <> strTbsInfo Then
                    strTbsInfo = rstmp!Tablespace_Name
                    gstrSQL = "alter tablespace " & strTbsInfo & "  shrink space keep " & Val(strSize) & "M"
                    gcnOracle.Execute gstrSQL
                End If
            Else
                gstrSQL = "alter database tempfile '" & rstmp!File_Name & "'  resize " & Val(strSize) & "M"
                gcnOracle.Execute gstrSQL
            End If
            
            If Err <> 0 Then
                strError = strError & vbCrLf & rstmp!File_Name & vbCrLf & Err.Description
                Err.Clear
            End If
            rstmp.MoveNext
        Wend
        
        If strError <> "" Then
            MsgBox "收缩表空间出错 " & vbCrLf & strError & vbCrLf & "请重新指定保留文件的大小，或者重启系统后执行收缩。", vbInformation, gstrSysName
        Else
            lblPrompt.Caption = "临时表空间收缩完毕！"
        End If
    Else
        MsgBox "当前没有大于10M的临时数据文件，不需要收缩。"
    End If
    
    Call SetCommandEnable(1)
    Exit Sub
errH:
    Call ErrCenter(gstrSQL)
    Call SetCommandEnable(1)
End Sub


