VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmAppChkRpt 
   Caption         =   "对象检查结果"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12840
   Icon            =   "frmAppChkRpt.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   12840
   StartUpPosition =   2  '屏幕中心
   Tag             =   "17500"
   Begin VB.CommandButton cmdSQL 
      Caption         =   "复制相关SQL"
      Height          =   350
      Left            =   8400
      TabIndex        =   11
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "修正(&M)"
      Height          =   350
      Left            =   9960
      TabIndex        =   6
      Top             =   7320
      Width           =   1100
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   11760
      TabIndex        =   5
      Top             =   7200
      Width           =   1100
   End
   Begin VB.CommandButton cmdOut 
      Caption         =   "输出到Excel"
      Height          =   350
      Left            =   6840
      TabIndex        =   3
      Top             =   7440
      Width           =   1335
   End
   Begin VB.ComboBox cboFilter 
      Height          =   300
      Index           =   0
      Left            =   6240
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   30
      Width           =   2205
   End
   Begin VB.ComboBox cboFilter 
      Height          =   300
      Index           =   1
      Left            =   9000
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   0
      Width           =   1500
   End
   Begin VB.ComboBox cboFilter 
      Height          =   300
      Index           =   2
      Left            =   12240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   30
      Width           =   1200
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfResult 
      Height          =   6255
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   10695
      _cx             =   18865
      _cy             =   11033
      Appearance      =   3
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
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
      Editable        =   2
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
   Begin VB.Label lblFilter 
      AutoSize        =   -1  'True
      Caption         =   "系统"
      Height          =   180
      Index           =   0
      Left            =   5760
      TabIndex        =   10
      Top             =   75
      Width           =   360
   End
   Begin VB.Label lblFilter 
      AutoSize        =   -1  'True
      Caption         =   "类型"
      Height          =   180
      Index           =   1
      Left            =   8520
      TabIndex        =   9
      Top             =   75
      Width           =   360
   End
   Begin VB.Label lblFilter 
      AutoSize        =   -1  'True
      Caption         =   "严重程度"
      Height          =   180
      Index           =   2
      Left            =   11400
      TabIndex        =   8
      Top             =   75
      Width           =   720
   End
   Begin VB.Label lblRsFilter 
      Caption         =   "Label1"
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   7200
      Width           =   5535
   End
End
Attribute VB_Name = "frmAppChkRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MSTR_COL = ",300,4;序号,500,4;系统,2000,1;类型,1500,1;对象名,2450,1;问题描述,6300,1;修正说明,3000,1;严重程度,930,4;修正SQL,0,4;约束字段,0,4"
Private Const MSTR_ProCOL = "序号,800,4;过程名称,2450,1;问题描述,6300,1"
Private mrsProData As New ADODB.Recordset
Private mrsDataFromFile As New ADODB.Recordset
Private mstrSysModul As String
Private Enum enuResult
    Col_选择 = 0
    Col_序号
    Col_系统
    Col_类型
    Col_对象名
    Col_问题描述
    Col_修正说明
    Col_严重程度
    Col_修正SQL
    Col_约束字段
End Enum

Private Enum enuPro
    Procol_序号 = 0
    Procol_过程名称 = 1
    Procol_问题描述 = 2
End Enum

Private mblnFirst As Boolean
Private mstrPath As String
Private mbytType As Byte   '1-对象检查结果，其它-过程检查结果
Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal DirPath As String) As Long '多级目录不存在，也可创建指定目录

Private Sub cboFilter_Click(Index As Integer)
    Dim strFilter As String
    
    If mblnFirst = False Then Exit Sub
    
    If cboFilter(0).Text = "所有系统" Then
        strFilter = ""
    Else
        strFilter = "系统名称='" & cboFilter(0).Text & "'"
    End If
    
    If cboFilter(1).Text = "所有类型" Then
        strFilter = strFilter
    Else
        strFilter = IIf(strFilter = "", "类别='" & cboFilter(1).Text & "'", strFilter & " and 类别='" & cboFilter(1).Text & "'")
    End If
    
    If cboFilter(2).Text = "所有程度" Then
        strFilter = strFilter
    Else
        strFilter = IIf(strFilter = "", "严重程度='" & cboFilter(2).Text & "'", strFilter & " and 严重程度='" & cboFilter(2).Text & "'")
    End If
    
    Call AddvsfData(strFilter)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSQL_Click()
    Dim strText As String
    Dim strTableName As String
    Dim strFild As String
    Dim varTemp As Variant
    Dim i As Long
    
    If cmdSQL.Visible = False Then Exit Sub
    Clipboard.Clear
    strText = vsfResult.Cell(flexcpText, vsfResult.Row, Col_约束字段)
    strTableName = vsfResult.Cell(flexcpText, vsfResult.Row, Col_对象名)
    strTableName = Mid(strTableName, 1, InStr(strTableName, "_") - 1)
    varTemp = Split(strText, ",")
    For i = 0 To UBound(varTemp)
        strFild = IIf(strFild = "", "", strFild & " And ") & "a." & varTemp(i) & "=b." & varTemp(i)
    Next
    strText = "Delete " & strTableName & " Where Rowid In (Select a.Rowid From " & strTableName & " a,(Select " & strText & ", Max(Rowid) Rid From " & _
           strTableName & " Group By " & strText & ") b Where " & strFild & " And a.Rowid <> b.Rid)"
    Clipboard.SetText strText
End Sub

Private Sub Form_Load()
    
    If mbytType = 1 Then
        Me.Caption = "对象检查结果"
        mblnFirst = False
        Call InitTable(vsfResult, MSTR_COL)
        Call InivsfData
        mblnFirst = True
        Me.Tag = 17400
        cmdClose.Caption = "取消(&C)"
        Call vsfResult_Click
    Else
        Me.Caption = "过程检查结果"
        cmdClose.Caption = "退出(&E)"
        cmdModify.Caption = "查看过程"
        Call InitTable(vsfResult, MSTR_ProCOL)
        With vsfResult
            .Rows = .Rows - 1
            .rowHeight(0) = 500
            mrsProData.Sort = "问题描述"
            mrsProData.MoveFirst
            Do While Not mrsProData.EOF
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, Procol_序号) = .Rows - 1
                .TextMatrix(.Rows - 1, Procol_过程名称) = mrsProData!过程名称
                .TextMatrix(.Rows - 1, Procol_问题描述) = mrsProData!问题描述
                .rowHeight(.Rows - 1) = 500
                mrsProData.MoveNext
            Loop
            If .Rows > 1 Then
                .Row = 1
                Call .ShowCell(1, 1)
            End If
        End With
        lblRsFilter.Caption = "检查结果:共" & mrsProData.RecordCount & "个过程问题"
    End If
End Sub

Public Function ShowMe(ByVal bytType As Byte, ByVal rsProData As ADODB.Recordset, Optional ByVal strPath As String, Optional ByVal rsDataFromFile As ADODB.Recordset) As Boolean
    'bytType,1-脚本检查修复，0-存储过程检查
    
    mbytType = bytType
    Set mrsProData = rsProData
    If bytType = 1 Then
        Set mrsDataFromFile = rsDataFromFile
        mstrPath = strPath & "\Log\日志跟踪\zlObjCheck_" & Replace(Format(Now, "yyyy-mm-dd"), "-", "") & ".Log"
    End If
    SetVisible
    Me.Show 1
End Function

Private Sub SetVisible()
    '设置控件的可见性
    Dim i As Long
    
    For i = 0 To lblFilter.UBound
        lblFilter(i).Visible = IIf(mbytType = 1, True, False)
    Next
    For i = 0 To lblFilter.UBound
        cboFilter(i).Visible = IIf(mbytType = 1, True, False)
    Next
    cmdSQL.Visible = False
End Sub

Private Sub InivsfData()
'功能：对象检查第首次数据显示
    Dim i As Long
    Dim strSys As String
    Dim strType2 As String
    Dim strSer As String
    
    With vsfResult
        strSys = "所有系统"
        strType2 = "所有类型"
        strSer = "所有程度"
        cboFilter(0).AddItem "所有系统"
        cboFilter(1).AddItem "所有类型"
        cboFilter(2).AddItem "所有程度"
        cboFilter(2).AddItem "严重"
        cboFilter(2).AddItem "较重"
        cboFilter(2).AddItem "轻微"
        .Cell(flexcpChecked, 0, Col_选择) = flexUnchecked
        .Rows = .Rows - 1
        Call AddvsfData
        
        For i = 1 To .Rows - 1
            If InStr("|" & strSys & "|", "|" & .TextMatrix(i, Col_系统) & "|") = 0 Then
                strSys = strSys & "|" & .TextMatrix(i, Col_系统)
                cboFilter(0).AddItem .TextMatrix(i, Col_系统)
            End If
            
            If InStr("|" & strType2 & "|", "|" & .TextMatrix(i, Col_类型) & "|") = 0 Then
                strType2 = strType2 & "|" & .TextMatrix(i, Col_类型)
                cboFilter(1).AddItem .TextMatrix(i, Col_类型)
            End If
        Next
    End With
    
    cboFilter(0).ListIndex = 0
    cboFilter(1).ListIndex = 0
    cboFilter(2).ListIndex = 0
End Sub

Private Sub AddvsfData(Optional ByVal strFilter As String)
'功能：添加问题错误到表格中
    Dim i As Long
    
    With vsfResult
        .Rows = 1
        .Redraw = flexRDNone
        .ColHidden(Col_修正SQL) = True
        mrsProData.Filter = strFilter
        .Rows = mrsProData.RecordCount + 1
        i = 0
        Do While Not mrsProData.EOF
            i = i + 1
            .TextMatrix(i, Col_序号) = i
            .TextMatrix(i, Col_系统) = mrsProData!系统名称
            .TextMatrix(i, Col_类型) = mrsProData!类别
            .TextMatrix(i, Col_对象名) = mrsProData!对象名
            .TextMatrix(i, Col_问题描述) = mrsProData!问题描述
            .TextMatrix(i, Col_修正说明) = mrsProData!修正说明
            .TextMatrix(i, Col_严重程度) = mrsProData!严重程度
            .TextMatrix(i, Col_修正SQL) = mrsProData!修正SQL
            .TextMatrix(i, Col_约束字段) = "" & mrsProData!约束字段
            If .TextMatrix(i, Col_严重程度) = "轻微" Then
                .Cell(flexcpBackColor, i, Col_严重程度) = RGB(238, 230, 133)
            ElseIf .TextMatrix(i, Col_严重程度) = "较重" Then
                .Cell(flexcpBackColor, i, Col_严重程度) = RGB(238, 201, 0)
            ElseIf .TextMatrix(i, Col_严重程度) = "严重" Then
                .Cell(flexcpBackColor, i, Col_严重程度) = RGB(238, 154, 0)
            End If
            If InStr(.TextMatrix(i, Col_修正说明), "人工") > 0 Then
                .TextMatrix(i, Col_选择) = ""
            Else
                .Cell(flexcpChecked, i, Col_选择) = flexUnchecked
            End If
            mrsProData.MoveNext
        Loop
        .Cell(flexcpAlignment, 0, 0, .Rows - 1) = 4
        .Redraw = flexRDDirect
        If .Rows > 1 Then
            .Row = 1
            Call .ShowCell(1, 1)
        End If
    End With
    lblRsFilter.Caption = "检查结果：共" & mrsProData.RecordCount & "个问题。"
End Sub

Private Sub cmdModify_Click()
'功能：修正勾选的对象修正
    Dim i As Long
    Dim j As Long
    Dim lngLine As Long
    Dim varTemp As Variant
    Dim strErr As String
    Dim strTemp As String
    Dim strSQL As String
    Dim blnModify As Boolean
    Dim blnFalse As Boolean
    Dim cnChoose As ADODB.Connection
    Dim rsTemp As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
            
    If mbytType = 1 Then
        With vsfResult
            For i = 1 To .Rows - 1
                If .Cell(flexcpChecked, i, 0) = flexChecked Then
                    Call ShowFlash("正在进行对象或数据的修正，请稍候！")
                    If .TextMatrix(i, Col_系统) = "服务器管理工具" Then
                        If gcnTools Is Nothing Then
                            Set gcnTools = GetConnection("ZLTOOLS")
                        End If
                        Set cnChoose = gcnTools
                    Else
                        Set cnChoose = gcnOracle
                    End If
                    blnFalse = True
                    varTemp = Split(UCase(.TextMatrix(i, Col_修正SQL)), "{JM|SQL分隔符}" & vbNewLine)
                    For j = 0 To UBound(varTemp)
                        strSQL = varTemp(j)
                        If strSQL <> "" Then
                            On Error Resume Next
                            cnChoose.Execute strSQL
                            If err.Number <> 0 Then
                                If strSQL Like "INSERT INTO ZLPARAMETERS*" Then
                                    strTemp = Replace(Mid(strSQL, InStr(strSQL, "(") + 1, InStr(strSQL, ")") - InStr(strSQL, "(") - 1), " ", "")
                                    Set rsTemp = SetSelectRecordset(strSQL, strTemp, Split(strTemp, ","), "ZLPARAMETERS")
                                    If InStr(rsTemp!模块, "NULL") = 0 And InStr(rsTemp!系统, "NULL") = 0 Then
                                        If InStr(mstrSysModul, rsTemp!系统 & "&" & rsTemp!模块) = 0 Then
                                            mrsDataFromFile.Filter = "类别='参数'"
                                            Set rsData = CopyNewRec(mrsDataFromFile)
                                            mstrSysModul = mstrSysModul & "|" & rsTemp!系统 & "&" & rsTemp!模块
                                            strSQL = "Update Zlparameters Set 参数号 = -1 * 参数号 Where 系统 =" & rsTemp!系统 & " And 模块 = " & rsTemp!模块
                                            cnChoose.Execute strSQL
                                            rsData.Filter = "类别='参数' and 对象=" & rsTemp!模块 & " and 系统编号=" & rsTemp!系统
                                            Do While Not rsData.EOF
                                                mrsDataFromFile.Filter = "类别='参数' and 对象=" & rsTemp!模块 & " and 系统编号=" & rsTemp!系统 & " and 参数名='" & rsData!参数名 & "'"
                                                If mrsDataFromFile.RecordCount > 0 Then
                                                    strSQL = "Update Zlparameters Set 参数号 = " & rsTemp!参数号 & " Where 系统 =" & rsTemp!系统 & " And 模块 = " & rsTemp!模块 & " and 参数名='" & rsData!参数名 & "'"
                                                    cnChoose.Execute strSQL
                                                End If
                                                rsData.MoveNext
                                            Loop
                                            cnChoose.Execute varTemp(j)
    '                                        strSQL = "Update Zlparameters Set 参数号 = -1 * 参数号 Where 系统 =" & rsTemp!系统 & " And 模块 = " & rsTemp!模块
    '                                        cnChoose.Execute strSQL
                                        End If
                                    Else
                                        blnFalse = False
                                        strErr = IIf(strErr = "", "修正失败的SQL；" & vbCrLf & varTemp(j) & ";" & vbCrLf & "原因:" & err.Description & vbCrLf, strErr & vbCrLf & varTemp(j) & ";" & vbCrLf & "原因:" & err.Description & vbCrLf)
                                    End If
                                Else
                                    '删除时不存在则表示删除成功
                                    If UCase(err.Description) Like "ORA-01418*" Then
                                    Else
                                        blnFalse = False
                                        strErr = IIf(strErr = "", "修正失败的SQL；" & vbCrLf & varTemp(j) & ";" & vbCrLf & "原因:" & err.Description & vbCrLf, strErr & vbCrLf & varTemp(j) & ";" & vbCrLf & "原因:" & err.Description & vbCrLf)
                                    End If
                                End If
                            Else
                                '删除表或过程/函数时，需删除对应的公共同义词
                                If .TextMatrix(i, Col_类型) = "ZLTOOL对象" Then
                                    If strSQL Like "DROP TABLE*" Or strSQL Like "DROP PROCEDURE*" Or strSQL Like "DROP FUNCTION" Then
                                        gstrSQL = "Select 'Drop Public SYNONYM ' || Synonym_Name 执行SQL" & vbNewLine & _
                                                    "From All_Synonyms a" & vbNewLine & _
                                                    "Where Table_Owner=[1] And Owner = 'PUBLIC' And Not Exists" & vbNewLine & _
                                                    " (Select 1 From All_Objects b Where a.Table_Name = b.Object_Name And a.Table_Owner = b.Owner) And" & vbNewLine & _
                                                    "      a.Synonym_Name =[2]"
                                        Set rsTemp = gclsBase.OpenSQLRecord(cnChoose, gstrSQL, Me.Caption, UCase(.TextMatrix(i, Col_系统)), UCase(.TextMatrix(i, Col_对象名)))
                                        Do While Not rsTemp.EOF
                                            cnChoose.Execute rsTemp!执行SQL
                                            rsTemp.MoveNext
                                        Loop
                                    End If
                                End If
                            End If
                        End If
                    Next
                    blnModify = True
                    If blnFalse Then
                        .Cell(flexcpData, i, 0) = 1
                    Else
                        .Cell(flexcpData, i, 0) = 0
                    End If
                End If
            Next
            If blnModify = False Then
                MsgBox "未勾选可自动修正的数据！"
                Exit Sub
            End If
            Call ShowFlash("")
            If strErr <> "" Then
                On Error Resume Next
                Call WriteErrorLog(strErr)
                If err.Number = 0 Then
                    MsgBox "修正完成！有部分数据未成功修正，详情见：" & mstrPath
                Else
                    MsgBox "修正完成！错误日志记录失败，可能是该日志文件(" & mstrPath & ")已打开，请检查！"
                End If
                err.Clear: On Error GoTo 0
            Else
                MsgBox "修正完成！"
            End If
        End With
        Call AfterModify
    Else
        With vsfResult
            If .Rows = 1 Then Exit Sub
            mrsProData.Filter = "过程名称='" & .TextMatrix(.Row, Procol_过程名称) & "'"
            strTemp = "Create or Replace " & mrsProData!原始SQL
            If InStr(.TextMatrix(.Row, Procol_问题描述), "Commit") > 0 Then
                lngLine = GetFirstLine(mrsProData!原始SQL, "COMMIT")
            ElseIf InStr(.TextMatrix(.Row, Procol_问题描述), "绑定变量") > 0 Then
                lngLine = GetFirstLine(mrsProData!原始SQL, "EXECUTE IMMEDIATE")
            End If
            Call frmProcEditCommon.ShowMe(0, .TextMatrix(.Row, Procol_过程名称), strTemp, "", "", "", 1, lngLine)
        End With
    End If
End Sub

Private Function GetFirstLine(ByVal strSQL As String, ByVal strKey As String) As Long
'获取指定关键字首次出现的行数
    Dim i As Long
    Dim varTemp As Variant
    
    varTemp = Split(UCase(strSQL), vbLf)
    For i = 0 To UBound(varTemp)
        If InStr(varTemp(i), strKey) > 0 Then
            GetFirstLine = i + 1
            Exit Function
        End If
    Next
End Function

Private Sub AfterModify()
'修正完成后重新刷新界面数据
    Dim i As Long
    Dim strFilter As String
    Dim lngSelRow As Long
    
    lblRsFilter.Caption = "正在重新刷新界面......"
    With vsfResult
        lngSelRow = .Row
        For i = 1 To .Rows - 1
            If .Cell(flexcpData, i, 0) = 1 Then
                strFilter = "问题描述='" & .TextMatrix(i, Col_问题描述) & "' and 对象名='" & .TextMatrix(i, Col_对象名) & "' and 类别='" & .TextMatrix(i, Col_类型) & "'"
                Call RecDelete(mrsProData, strFilter)
            End If
        Next
        Call cboFilter_Click(0)
        If .Rows > 1 Then
            If .Rows > lngSelRow Then
                .Row = lngSelRow
                Call .ShowCell(lngSelRow, 1)
            Else
                .Row = .Rows - 1
                Call .ShowCell(.Rows - 1, 1)
            End If
        End If
    End With
    Call vsfResult_AfterEdit(1, 0)
End Sub

Private Sub WriteErrorLog(ByVal strErr As String)
    Dim objFile As Object
    Dim objStream As TextStream
        
    Call MakeSureDirectoryPathExists(mstrPath)
    Set objFile = CreateObject("Scripting.FileSystemObject")
    If objFile.FileExists(mstrPath) = False Then objFile.CreateTextFile mstrPath
    Set objStream = objFile.OpenTextFile(mstrPath)

    Open mstrPath For Append Shared As #1
    Print #1, strErr
    Close #1
End Sub

Private Sub cmdOut_Click()
    
    Call OutExcel
End Sub

Private Sub OutExcel()
'功能：将vsf表格结果输出到Excel中
    Dim spShell, spFolder, spFolderItem, spPath As String
    Const WINDOW_HANDLE = 0
    Const NO_OPTIONS = 0

    On Error GoTo errH
    If IsInstallExcel Then
        With vsfResult
            If .Rows < 2 Then
                MsgBox "表格中没有数据，无法输出数据，请检查！"
                Exit Sub
            Else
                Set spShell = CreateObject("Shell.Application")
                Set spFolder = spShell.BrowseForFolder(WINDOW_HANDLE, "选择目录:", NO_OPTIONS)
                If spFolder Is Nothing Then
                    Exit Sub
                Else
                    Set spFolderItem = spFolder.Self
                    spPath = spFolderItem.Path
                    .SaveGrid Replace(spPath & "\zlObjectCheck_" & Replace(Format(Now, "yyyy-mm-dd"), "-", "") & ".xls", "\\", "\"), flexFileExcel, True
                    .BackColorSel = &H8000000D
                     MsgBox "保存成功！检查结果已保存至" & Replace(spPath & "\zlObjectCheck_" & Replace(Format(Now, "yyyy-mm-dd"), "-", "") & ".xls", "\\", "\")
                     Exit Sub
                End If
            End If
        End With
    End If
    Exit Sub
errH:
    MsgBox "所选路径的该文件处于打开状态或所选路径错误！"
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    If ScaleHeight < 2000 Then Exit Sub
        
    With vsfResult
        .Left = ScaleLeft
        .Width = ScaleWidth
        If mbytType = 1 Then
            .Top = ScaleTop + 600
            .Height = ScaleHeight - cmdModify.Height - 900
            .ColWidth(Col_类型) = 1500 + 0.05 * (Me.Width - Me.Tag)
            .ColWidth(Col_对象名) = 2450 + 0.25 * (Me.Width - Me.Tag)
            .ColWidth(Col_问题描述) = 6300 + 0.3 * (Me.Width - Me.Tag)
            .ColWidth(Col_修正说明) = 3000 + 0.4 * (Me.Width - Me.Tag)
        Else
            .Top = ScaleTop
            .Height = ScaleHeight - cmdModify.Height - 300
            .ColWidth(Procol_过程名称) = (.Width - 800) * 0.3
            .ColWidth(Procol_问题描述) = (.Width - 800) * 0.7
        End If
    End With
    cmdClose.Top = vsfResult.Top + vsfResult.Height + 150
    cmdClose.Left = ScaleWidth - cmdClose.Width - 300
    
    cmdModify.Top = cmdClose.Top
    cmdModify.Left = cmdClose.Left - cmdModify.Width - 500
    
    cmdSQL.Top = cmdClose.Top
    cmdSQL.Left = cmdModify.Left - cmdSQL.Width - 500
    
    cmdOut.Top = cmdClose.Top
    cmdOut.Left = cmdSQL.Left - cmdOut.Width - 500
    
    lblRsFilter.Top = cmdOut.Top + 150
    lblRsFilter.Left = 300
    
    cboFilter(2).Top = 200
    cboFilter(2).Left = ScaleWidth - cboFilter(2).Width - 300
    lblFilter(2).Top = 250
    lblFilter(2).Left = cboFilter(2).Left - lblFilter(2).Width - 150
    
    cboFilter(1).Top = 200
    cboFilter(1).Left = lblFilter(2).Left - cboFilter(1).Width - 300
    lblFilter(1).Top = 250
    lblFilter(1).Left = cboFilter(1).Left - lblFilter(1).Width - 150
    
    cboFilter(0).Top = 200
    cboFilter(0).Left = lblFilter(1).Left - cboFilter(0).Width - 300
    lblFilter(0).Top = 250
    lblFilter(0).Left = cboFilter(0).Left - lblFilter(0).Width - 150

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mbytType = 1 Then
        Call ReleaseMe
    End If
    Set mrsProData = Nothing
    Set mrsDataFromFile = Nothing
End Sub

Private Sub vsfResult_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    
    If mbytType = 1 Then
        With vsfResult
            If Col = Col_选择 Then
                If Row = 0 Then
                    If .Cell(flexcpChecked, 0, Col_选择) = flexChecked Then
                        .Cell(flexcpChecked, 0, Col_选择) = flexChecked
                        For i = 1 To .Rows - 1
                            If .Cell(flexcpChecked, i, Col_选择) = flexUnchecked Then
                                .Cell(flexcpChecked, i, Col_选择) = flexChecked
                            End If
                        Next
                    Else
                        .Cell(flexcpChecked, 0, Col_选择) = flexUnchecked
                        For i = 1 To .Rows - 1
                            If .Cell(flexcpChecked, i, Col_选择) = flexChecked Then
                                .Cell(flexcpChecked, i, Col_选择) = flexUnchecked
                            End If
                        Next
                    End If
                Else
                    If .Cell(flexcpChecked, 0, Col_选择) = flexChecked Then
                        .Cell(flexcpChecked, 0, Col_选择) = flexUnchecked
                    End If
                    For i = 1 To .Rows - 1
                        If .Cell(flexcpChecked, i, Col_选择) = flexUnchecked Then
                            Exit For
                        Else
                            If i = .Rows - 1 Then
                                .Cell(flexcpChecked, 0, Col_选择) = flexChecked
                            End If
                        End If
                    Next
                End If
            End If
        End With
    End If
End Sub

Private Sub vsfResult_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    If mbytType = 1 Then
        If Col <> 0 Then Cancel = True
    End If
End Sub

Private Sub vsfResult_Click()
    Dim strTemp As String
    
    If mbytType = 0 Then Exit Sub
    '对象检查修复时才进行复制相关SQL的可见性调整
    strTemp = vsfResult.Cell(flexcpText, vsfResult.Row, Col_对象名)
    If (strTemp Like "*_PK" Or strTemp Like "*_UQ_*") And Mid(vsfResult.Cell(flexcpText, vsfResult.Row, Col_问题描述), 1, 6) <> "数据库中存在" Then
        cmdSQL.Visible = True
    Else
        cmdSQL.Visible = False
    End If
End Sub

Private Sub vsfResult_DblClick()
    If mbytType = 1 Then Exit Sub
    '过程检查结果查看过程才执行该操作
    Call cmdModify_Click
End Sub

Private Sub vsfResult_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim strTip As String
    
    If mbytType = 1 Then
        With vsfResult
            If .MouseRow <> -1 And .MouseRow <> 0 And .MouseCol = Col_修正说明 Then
                If .TextMatrix(.MouseRow, Col_修正SQL) <> "" Then
                    strTip = "修正SQL:" & vbNewLine & Replace(.TextMatrix(.MouseRow, Col_修正SQL), "{JM|SQL分隔符}", "")
                    Call ShowTipInfo(.hwnd, strTip, True)
                Else
                    Call ShowTipInfo(.hwnd, "")
                End If
            Else
                Call ShowTipInfo(.hwnd, "")
            End If
        End With
    End If
End Sub


