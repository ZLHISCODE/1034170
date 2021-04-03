VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmAutoJobs 
   BackColor       =   &H80000005&
   Caption         =   "后台作业管理"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6465
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmAutoJobs.frx":0000
   ScaleHeight     =   6450
   ScaleWidth      =   6465
   WindowState     =   2  'Maximized
   Begin VB.Frame fraComment 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1125
      Left            =   885
      TabIndex        =   5
      Top             =   4095
      Width           =   4920
      Begin VB.Label lblPara 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "参数："
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   825
         Width           =   540
      End
      Begin VB.Label lblComment 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "说明："
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   195
         Width           =   540
      End
      Begin VB.Label lbl说明 
         BackStyle       =   0  'Transparent
         Height          =   525
         Left            =   600
         TabIndex        =   6
         Top             =   210
         Width           =   1965
      End
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "测试执行(&T)…"
      Height          =   350
      Left            =   885
      TabIndex        =   4
      Top             =   5685
      Width           =   1395
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "运行设置(&T)…"
      Height          =   350
      Left            =   2280
      TabIndex        =   3
      Top             =   5685
      Width           =   1395
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "增加(&A)"
      Height          =   350
      Left            =   3945
      TabIndex        =   2
      Top             =   5685
      Width           =   945
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "删除(&D)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   4845
      TabIndex        =   1
      Top             =   5685
      Width           =   945
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfJobs 
      Height          =   2175
      Left            =   885
      TabIndex        =   0
      Top             =   1455
      Width           =   5415
      _cx             =   9551
      _cy             =   3836
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
      BackColorBkg    =   16777215
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
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   3
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   0
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
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
   Begin VB.Image imgMain 
      Height          =   480
      Left            =   285
      Picture         =   "frmAutoJobs.frx":04F9
      Top             =   660
      Width           =   480
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "后台作业管理"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   1440
   End
   Begin VB.Label lblMain 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Left            =   1380
      TabIndex        =   10
      Top             =   6255
      Width           =   4890
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl附加 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   885
      TabIndex        =   9
      Top             =   5745
      Width           =   4890
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAutoJobs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MSTR_COL = "系统,2000,1;类别,1200,1;序号,450,1;说明,0,1;参数,0,1;名称,2000,1;调用过程,3000,1;作业号,0,1;自动执行,800,4;状态,600,4;开始执行时间,1900,1;间隔时间,900,1;系统编号,0,1;所有者,0,1"
Private Enum vsfCol
    Col_系统 = 0
    Col_类别 = 1
    Col_序号 = 2
    Col_说明 = 3
    Col_参数 = 4
    Col_名称 = 5
    Col_调用过程 = 6
    Col_作业号 = 7
    Col_自动执行 = 8
    Col_状态 = 9
    Col_开始执行时间 = 10
    Col_间隔时间 = 11
    Col_系统编号 = 12
    Col_所有者 = 13
End Enum
Private mlngMaxJobs As Long '本数据库最多可设的作业数
Private mstrPro As String '当前的所有调用过程字符串

Private Sub cmdAdd_Click()
    Dim lngSelectRow As Long

    lngSelectRow = vsfJobs.Row
    Call frmAutoJobset.Add(mstrPro)
    Call LoadData(lngSelectRow)
End Sub

Private Sub cmdDel_Click()
    Dim lngSystem As Long
    Dim cnTools As ADODB.Connection
    Dim strTemp As String

    With vsfJobs
        strTemp = "你确认要删除（" & .TextMatrix(.Row, Col_名称) & "）后台作业吗？"
        If MsgBox(strTemp, vbExclamation + vbDefaultButton1 + vbYesNo) = vbNo Then
            Exit Sub
        End If
        lngSystem = .TextMatrix(.Row, Col_系统编号)
        strTemp = UCase(.TextMatrix(.Row, Col_调用过程))
        If Val(.TextMatrix(.Row, Col_作业号)) <> 0 Then
            If lngSystem = 0 Then
                Set cnTools = GetConnection("ZLTOOLS")
                If cnTools Is Nothing Then Exit Sub
            Else
                Set cnTools = gcnOracle
            End If
            gstrSQL = "zl_JobRemove(" & IIf(lngSystem = 0, "Null", lngSystem) & ",3," & .TextMatrix(.Row, Col_序号) & ")"
            err = 0
            cnTools.Execute gstrSQL, , adCmdStoredProc
            If err <> 0 Then
                MsgBox "作业删除失败！", vbExclamation, gstrSysName
                Exit Sub
            End If
        End If
        gstrSQL = "delete zlAutoJobs" & _
                " where Nvl(系统,0)=" & lngSystem & " and 类型=3" & _
                " and 序号=" & .TextMatrix(.Row, Col_序号)
        err = 0
        On Error Resume Next
        gcnOracle.Execute gstrSQL
        If err <> 0 Then
            MsgBox "作业删除失败！", vbExclamation, gstrSysName
            Exit Sub
        Else
            mstrPro = Replace(mstrPro, strTemp & ",", "")
        End If
    End With
    Call LoadData(vsfJobs.Row)
End Sub

Private Sub cmdSet_Click()
    Dim strParas As String
    Dim aryPara() As String
    Dim intCount As Integer
    
    If Val(vsfJobs.TextMatrix(vsfJobs.Row, Col_序号)) = 0 Then Exit Sub
    Call frmAutoJobset.RunSet(vsfJobs)
    Call LoadData(vsfJobs.Row)
End Sub

Private Sub cmdTest_Click()
    Dim lngSystem As Long
    Dim cnTools As ADODB.Connection
    Dim lngType As Long
    
    With vsfJobs
        If gblnDBA Then
            'DBA用户不需要判断
        ElseIf .TextMatrix(.Row, Col_系统) = "服务器管理工具" Then
            '因为管理工具所有者为空，这里需要先判断是否为管理工具
        ElseIf .TextMatrix(.Row, Col_所有者) <> gstrUserName Then
            MsgBox "当前用户不是该系统的所有者，无法进行该操作！"
            Exit Sub
        End If
        lngSystem = .TextMatrix(.Row, Col_系统编号)
        If Val(.TextMatrix(.Row, Col_作业号)) <> 0 Then
            If lngSystem = 0 Then
                Set cnTools = GetConnection("ZLTOOLS")
                If cnTools Is Nothing Then Exit Sub
            Else
                Set cnTools = gcnOracle
            End If
            If .TextMatrix(.Row, Col_类别) = "系统设定" Then
                lngType = 1
            ElseIf .TextMatrix(.Row, Col_类别) = "数据转移" Then
                lngType = 2
            Else
                lngType = 3
            End If
            gstrSQL = "zl_JobRun(" & IIf(lngSystem = 0, "Null", lngSystem) & "," & lngType & "," & .TextMatrix(.Row, Col_序号) & ")"
            err = 0
            On Error Resume Next
            cnTools.Execute gstrSQL, , adCmdStoredProc
            If err <> 0 Then
                MsgBox "测试过程发生意外错误！" & vbNewLine & err.Description, vbExclamation, gstrSysName
                Exit Sub
            End If
            MsgBox "测试执行完成，如果该作业状态变为“有效”，说明执行成功。", vbInformation, gstrSysName
        End If
    End With
    Call LoadData(vsfJobs.Row)
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    Dim strSame As String
    Dim i As Integer
    
    On Error GoTo errHandle
    '转入不存在的数据转移记录作为作业记录
    gstrSQL = "INSERT INTO zlAutoJobs (系统,类型,序号,名称,说明,内容,参数,执行时间,间隔时间)" & _
            " SELECT 系统,2,组号,组名,说明,'zl'||floor(系统/100)||'_DataMoveOut'||组号,日期字段||','||转出描述,to_date('2000-01-01 01:00:00','YYYY-MM-DD HH24:MI:SS'),30" & _
            " FROM zlDataMove" & _
            " WHERE (系统,组号) not in( select 系统,序号 from zlAutoJobs where 类型=2)"
    gcnOracle.Execute gstrSQL
    
    lblMain.Caption = "本功设置数据库后台自动作业，用于定期批量进行的数据计算和数据修改等事务。" & _
        vbCrLf & vbCrLf & "建议设置在系统比较空闲的时间执行，以减少和其他任务的资源竞争，保证前台事务的运行速度。"
    
    gstrSQL = "select value" & _
            " from v$parameter" & _
            " where name='job_queue_processes'"
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open gstrSQL, gcnOracle, adOpenKeyset
    mlngMaxJobs = 0
    If Not rsTemp.EOF Then
        mlngMaxJobs = rsTemp.Fields(0).value
        If mlngMaxJobs > 0 Then
            lbl附加.Caption = "根据数据库参数job_queue_processes设置，目前最多可设置" & mlngMaxJobs & "个自动作业"
        Else
            lbl附加.Caption = "当前不能设置自动作业，如有必要，请修改数据库参数job_queue_processes"
        End If
    End If
    If mlngMaxJobs = 0 Then
        cmdTest.Enabled = False
        cmdSet.Enabled = False
        cmdAdd.Enabled = False
        cmdDel.Enabled = False
    End If
    Call InitTable(vsfJobs, MSTR_COL)
    Call LoadData(1)
    mstrPro = ""
    With vsfJobs
        For i = 1 To .Rows - 1
            If InStr(mstrPro, UCase(.TextMatrix(i, Col_调用过程)) & ",") = 0 Then
                mstrPro = mstrPro & UCase(.TextMatrix(i, Col_调用过程)) & ","
            End If
        Next
    End With
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, Me.Caption
 
End Sub

Private Sub LoadData(ByVal lngRow As Long)
'功能：加载界面时界面显示
    Dim rsTemp As New ADODB.Recordset
    Dim i As Long
    Dim j As Long
    Dim lngColor As Long
    Dim strTemp As String
    Dim strPro As String
    Dim varTemp As Variant
                
    With rsTemp
        gstrSQL = "select Nvl(C.名称,'服务器管理工具') 系统,decode(A.类型,1,'系统设定',2,'数据转移',3,'用户自定义') as 类别 ,A.序号,A.说明," & _
                "A.参数,A.内容 调用过程,A.作业号,A.名称,decode(A.作业号,0,'否',null,'否','是') as 自动执行," & _
                "decode(B.BROKEN,null,'缺失','Y','无效','有效') as 状态,A.执行时间 开始执行时间,A.间隔时间||Nvl(A.时间单位,'天') as 间隔时间,Nvl(A.系统,0) 系统编号,C.所有者 " & _
                "From zlAutoJobs A," & IIf(gblnDBA, "dba_jobs", "user_jobs") & " B,zlsystems C " & _
                "where A.作业号=B.JOB(+) and A.系统=C.编号(+) " & IIf(gblnOwner, " And c.所有者=user", "") & " order by 系统编号,类别,序号"
        If rsTemp.State = adStateOpen Then rsTemp.Close
        rsTemp.Open gstrSQL, gcnOracle, adOpenKeyset
    End With

    With vsfJobs
        .Rows = 1
        .rowHeight(0) = 300
        .MergeCells = flexMergeRestrictRows
        .MergeCol(Col_系统) = True
        .MergeCol(Col_类别) = True
        Do While Not rsTemp.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, Col_系统) = rsTemp!系统 & ""
            .TextMatrix(.Rows - 1, Col_类别) = rsTemp!类别 & ""
            .TextMatrix(.Rows - 1, Col_序号) = rsTemp!序号 & ""
            .TextMatrix(.Rows - 1, Col_说明) = rsTemp!说明 & ""
            .TextMatrix(.Rows - 1, Col_参数) = rsTemp!参数 & ""
            .TextMatrix(.Rows - 1, Col_名称) = rsTemp!名称 & ""
            .TextMatrix(.Rows - 1, Col_调用过程) = rsTemp!调用过程 & ""
            .TextMatrix(.Rows - 1, Col_作业号) = rsTemp!作业号 & ""
            .TextMatrix(.Rows - 1, Col_自动执行) = rsTemp!自动执行 & ""
            .TextMatrix(.Rows - 1, Col_状态) = rsTemp!状态 & ""
            .TextMatrix(.Rows - 1, Col_开始执行时间) = rsTemp!开始执行时间 & ""
            .TextMatrix(.Rows - 1, Col_间隔时间) = rsTemp!间隔时间 & ""
            .TextMatrix(.Rows - 1, Col_系统编号) = rsTemp!系统编号 & ""
            .TextMatrix(.Rows - 1, Col_所有者) = rsTemp!所有者 & ""
            rsTemp.MoveNext
        Loop
        For i = 1 To .Rows - 1
            .rowHeight(i) = 300
            strPro = UCase(.TextMatrix(i, Col_调用过程))
            If InStr(strTemp, strPro & ",") > 0 Then
                varTemp = Split(strTemp, ",")
                For j = 0 To UBound(varTemp)
                    If varTemp(j) = strPro Then
                        .Cell(flexcpBackColor, j + 1, 2, j + 1, .Cols - 1) = RGB(238, 230, 133 + lngColor * 10)
                    End If
                Next
                .Cell(flexcpBackColor, i, 2, i, .Cols - 1) = RGB(238, 230, 133 + lngColor * 10)
                lngColor = lngColor + 1
            End If
            strTemp = strTemp & strPro & ","
        Next
        If .Rows > 1 Then
            If lngRow > .Rows Then lngRow = .Rows
            .Row = lngRow
            Call .ShowCell(lngRow, 1)
            Call vsfJobs_RowColChange
        End If
    End With
End Sub

Private Sub Form_Resize()
    Dim sngBottom As Single
    
    On Error Resume Next
    
    With imgMain
        .Top = 700
        .Left = ScaleLeft + 200
    End With


    With vsfJobs
        .Top = imgMain.Top + 50
        .Left = imgMain.Left + imgMain.Width + 50
        .Width = ScaleWidth - .Left - 200
        sngBottom = ScaleHeight - lblMain.Height - 420 - cmdTest.Height - fraComment.Height - lbl附加.Height
        .Height = IIf(sngBottom - .Top > 2500, sngBottom - .Top, 2500)
    End With
    
    With lblMain
        .Left = vsfJobs.Left
        .Width = vsfJobs.Width

        lbl附加.Left = .Left
        lbl附加.Width = .Width
    End With
    
    fraComment.Width = vsfJobs.Width
    fraComment.Left = vsfJobs.Left
    fraComment.Top = vsfJobs.Top + vsfJobs.Height
    lbl说明.Width = fraComment.Width - lbl说明.Left - 300

    cmdDel.Left = vsfJobs.Left + vsfJobs.Width - cmdDel.Width
    cmdAdd.Left = cmdDel.Left - cmdAdd.Width
    cmdTest.Top = fraComment.Top + fraComment.Height + 60
    cmdTest.Left = vsfJobs.Left
    cmdSet.Top = cmdTest.Top
    cmdSet.Left = cmdTest.Left + cmdTest.Width
    cmdAdd.Top = cmdTest.Top
    cmdDel.Top = cmdTest.Top

    lblMain.Top = cmdTest.Top + cmdTest.Height + 200
    lbl附加.Top = lblMain.Top + lblMain.Height + 60
    
End Sub

Public Function SupportPrint() As Boolean
'返回本窗口是否支持打印，供主窗口调用
    SupportPrint = True
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
'供主窗口调用，实现具体的打印工作
'如果没有可打印的，就留下一个空的接口

'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow

    On Error GoTo errHandle
    objPrint.Title.Text = "后台作业"

    Set objRow = New zlTabAppRow
    objRow.Add "打印时间：" & Format(CurrentDate, "yyyy年MM月dd日")
    Set objPrint.Body = vsfJobs
    objPrint.BelowAppRows.Add objRow
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
    Exit Sub
errHandle:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Sub vsfJobs_RowColChange()

    With vsfJobs
        lbl说明.Caption = .TextMatrix(.Row, Col_说明)
        lblPara.Caption = "参数：" & .TextMatrix(.Row, Col_参数)
        If .TextMatrix(.Row, Col_自动执行) = "是" Then
            cmdTest.Enabled = True
        Else
            cmdTest.Enabled = False
        End If
        If .TextMatrix(.Row, Col_类别) = "用户自定义" Then
            cmdDel.Enabled = True
        Else
            cmdDel.Enabled = False
        End If
    End With
End Sub


