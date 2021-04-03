VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBillUses 
   Caption         =   "票据明细"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11805
   Icon            =   "frmBillUses.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   11805
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picFilter 
      BorderStyle     =   0  'None
      Height          =   390
      Left            =   105
      ScaleHeight     =   390
      ScaleWidth      =   8055
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   135
      Width           =   8055
      Begin VB.PictureBox picTimeRange 
         BorderStyle     =   0  'None
         Height          =   390
         Left            =   2145
         ScaleHeight     =   390
         ScaleWidth      =   6285
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   -30
         Visible         =   0   'False
         Width           =   6285
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "获取数据(&R)"
            Height          =   350
            Left            =   4680
            TabIndex        =   4
            Top             =   45
            Width           =   1200
         End
         Begin MSComCtl2.DTPicker dtpStartDate 
            Height          =   300
            Left            =   0
            TabIndex        =   2
            Top             =   60
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   127401987
            CurrentDate     =   41520
         End
         Begin MSComCtl2.DTPicker dtpEndDate 
            Height          =   300
            Left            =   2550
            TabIndex        =   3
            Top             =   60
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   127401987
            CurrentDate     =   41520
         End
         Begin VB.Label lblEndDate 
            AutoSize        =   -1  'True
            Caption         =   "～"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2265
            TabIndex        =   21
            Top             =   120
            Width           =   225
         End
      End
      Begin VB.ComboBox cbo使用日期 
         Height          =   300
         Left            =   780
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   30
         Width           =   1350
      End
      Begin VB.Label lbl使用日期 
         AutoSize        =   -1  'True
         Caption         =   "使用时间"
         Height          =   180
         Left            =   0
         TabIndex        =   0
         Top             =   90
         Width           =   720
      End
   End
   Begin VB.Frame fraCMD 
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   10320
      TabIndex        =   18
      Top             =   240
      Width           =   1455
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1200
      End
      Begin VB.CommandButton cmdDistant 
         Caption         =   "定位断号(&H)"
         CausesValidation=   0   'False
         Height          =   350
         Left            =   120
         TabIndex        =   10
         Top             =   1920
         Width           =   1200
      End
      Begin VB.TextBox txt号码 
         Height          =   300
         Left            =   150
         TabIndex        =   12
         Top             =   2700
         Width           =   1100
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "定位票据(&F)"
         CausesValidation=   0   'False
         Height          =   350
         Left            =   120
         TabIndex        =   13
         Top             =   3060
         Width           =   1200
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1200
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "保存(&S)"
         Height          =   350
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   1200
      End
      Begin VB.CommandButton cmdAllDO 
         Caption         =   "全部核对(&A)"
         Height          =   350
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   3720
         Width           =   1200
      End
      Begin VB.CommandButton cmdAllDO 
         Caption         =   "全部取消(&R)"
         Height          =   350
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   4080
         Width           =   1200
      End
      Begin VB.Label lbl提示 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "号码(&N)"
         Height          =   180
         Index           =   1
         Left            =   150
         TabIndex        =   11
         Top             =   2430
         Width           =   630
      End
      Begin VB.Line linBlack 
         BorderColor     =   &H00000000&
         X1              =   120
         X2              =   1300
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line linWhite 
         BorderColor     =   &H00FFFFFF&
         X1              =   150
         X2              =   1300
         Y1              =   1815
         Y2              =   1815
      End
   End
   Begin VB.TextBox txtInput 
      Height          =   270
      Left            =   7520
      MaxLength       =   200
      TabIndex        =   17
      Top             =   740
      Visible         =   0   'False
      Width           =   1450
   End
   Begin VB.ComboBox cboResult 
      Height          =   300
      Left            =   6300
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   6735
      Left            =   120
      TabIndex        =   5
      Top             =   540
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   11880
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      BackColorSel    =   12320767
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorUnpopulated=   -2147483644
      GridColor       =   8421504
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   "  号码  |   使用时间   |使用人|    使用情况    |    核对时间    |核对人|   核对结果  |      备注     |ID"
      MouseIcon       =   "frmBillUses.frx":0442
      _NumberOfBands  =   1
      _Band(0).Cols   =   9
   End
   Begin VB.Label lbl提示 
      Caption         =   "使用明细"
      Height          =   180
      Index           =   0
      Left            =   8160
      TabIndex        =   6
      Top             =   255
      Width           =   6210
   End
End
Attribute VB_Name = "frmBillUses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytInFun As Byte    '0-查看票据使用明细,1-核对票据明细
Private mblnViewCheck As Boolean '当mbytInFun=0时,是否显示核对相关字段
Private mlng票种 As Long
Private mlng领用ID As Long
Private mdblGiveCount As Double   '该批次票据总张数
Private mstr前缀文本 As String
Private mblnUnClick As Boolean

Private Enum Col
    C0号码 = 0
    C1使用时间 = 1
    C2使用人 = 2
    C3使用情况 = 3
    C4核对时间 = 4
    C5核对人 = 5
    C6核对结果 = 6
    C7备注 = 7
    C8ID = 8
End Enum
Private Sub SetUnChecked(ByVal lngRow As Long)
    With mshDetail
        .TextMatrix(lngRow, Col.C4核对时间) = ""
        .TextMatrix(lngRow, Col.C5核对人) = ""
        .TextMatrix(lngRow, Col.C6核对结果) = ""
        .TextMatrix(lngRow, Col.C7备注) = ""
        
        .RowData(lngRow) = 1  '用于保存时判断
        If Not cmdSave.Enabled Then cmdSave.Enabled = True
    End With
End Sub

Private Sub SetChecked(ByVal lngRow As Long, ByVal lngCol As Long, ByVal strContent As String, Optional ByVal strDate As String)
    With mshDetail
        If lngCol = Col.C6核对结果 Then
            .TextMatrix(lngRow, Col.C4核对时间) = strDate
            .TextMatrix(lngRow, Col.C5核对人) = UserInfo.姓名
            .TextMatrix(lngRow, lngCol) = strContent
        ElseIf lngCol = Col.C7备注 Then
            .TextMatrix(lngRow, lngCol) = strContent
        End If
        
        .RowData(lngRow) = 1 '用于保存时判断
        If Not cmdSave.Enabled Then cmdSave.Enabled = True
    End With
End Sub

Private Sub cboResult_LostFocus()
    If cboResult.Visible Then cboResult.Visible = False
End Sub
Private Sub RefreshCustomTime()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:按指定时间过滤数据
    '编制:刘兴洪
    '日期:2013-11-01 10:44:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, lngCount As Long, dtDate As Date
    
    On Error GoTo errHandle
    If cbo使用日期.Text <> "时间范围" Then Exit Sub
    lngCount = 0
    With mshDetail
        For i = 1 To .Rows - 1
             '指定时间段
             If IsDate(.TextMatrix(i, Col.C1使用时间)) Then
                 dtDate = CDate(.TextMatrix(i, Col.C1使用时间))
                 If dtDate >= dtpStartDate.Value And dtDate <= dtpEndDate.Value Then
                     .RowHeight(i) = .RowHeight(0)
                 Else
                     .RowHeight(i) = 0
                 End If
             Else
                 .RowHeight(i) = 0
             End If
            If .RowHeight(i) <> 0 Then
                lngCount = lngCount + 1
            End If
        Next
    End With
    lbl提示(0).Caption = lbl提示(0).Tag & "其中当前选中" & lngCount & "张票据"
     Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub cbo使用日期_Click()
    Dim i As Long, lngCount As Long, dtDate As Date
    If mblnUnClick = True Then Exit Sub
    On Error GoTo errHandle
    '问题:29885
    picTimeRange.Visible = False
    With mshDetail
        For i = 1 To .Rows - 1
            If cbo使用日期.Text = "所有" Then
                .RowHeight(i) = .RowHeight(0)
            ElseIf cbo使用日期.Text = "时间范围" Then
                picTimeRange.Visible = True
                If dtpStartDate.Visible And dtpStartDate.Enabled Then dtpStartDate.SetFocus
                Call Form_Resize
                Exit Sub
            ElseIf InStr(1, .TextMatrix(i, Col.C1使用时间), cbo使用日期.Text) > 0 Then
                .RowHeight(i) = .RowHeight(0)
            Else
                .RowHeight(i) = 0
            End If
            If .RowHeight(i) <> 0 Then
                lngCount = lngCount + 1
            End If
        Next
    End With
    If cbo使用日期.Text <> "所有" Then
        lbl提示(0).Caption = lbl提示(0).Tag & "其中当前选中" & lngCount & "张票据"
    Else
        lbl提示(0).Caption = lbl提示(0).Tag
    End If
    Call Form_Resize
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cbo使用日期_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdAllDO_Click(Index As Integer)
    Dim i As Long, strDate As String
    Dim blnSel As Boolean '是否存在多行选择
    Dim lngRows As Long
    Dim lngStart As Long
    
    With mshDetail
        blnSel = .Row <> .RowSel And .RowSel > .Row
        
        If Index = 0 Then
            strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
        End If
        lngStart = IIf(blnSel, .Row, 1)
        lngRows = IIf(blnSel, .RowSel, .Rows - 1)
        For i = lngStart To lngRows
            
            If .RowHeight(i) <> 0 Then
                If Index = 0 Then
                   '即使已核对的也重新核对,填写新的核对人和核对时间,不填备注,以前填了的也不用清除
                   Call SetChecked(i, Col.C6核对结果, .TextMatrix(i, Col.C3使用情况), strDate)
                Else
                    '没有核对过的,不必取消核对
                    If Trim(.TextMatrix(i, Col.C6核对结果)) <> "" Then Call SetUnChecked(i)
                End If
            End If
            
        Next
    End With
End Sub

Private Sub cboResult_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = vbKeyReturn Then
        With mshDetail
            If cboResult.ListIndex <= 0 Then
                Call SetUnChecked(.Row)
            Else
                Call SetChecked(.Row, Col.C6核对结果, Trim(cboResult.Text), Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss"))
            End If
            .SetFocus    '调用lostfocus
            .Col = .Col + 1
        End With
    ElseIf KeyAscii >= 32 Then
        If Chr(KeyAscii) > 5 Or Chr(KeyAscii) < 0 Then Exit Sub
        lngIdx = zlControl.CboMatchIndex(cboResult.hwnd, KeyAscii, 0.008)
        If lngIdx = -1 And cboResult.ListCount > 0 And cboResult.ListIndex = -1 Then lngIdx = 0
        cboResult.ListIndex = lngIdx
    End If
End Sub

Private Function SaveData() As Boolean
    Dim i As Long, arrSQL As Variant, blnTrans As Boolean, bytAllChecked As Byte, bytAllCheckOK As Byte
    Dim strDate As String, lngGiveCount As Long
    Dim strSQL As String, rsTmp As ADODB.Recordset
    
    With mshDetail
        arrSQL = Array()
        For i = 1 To .Rows - 1
            If .RowData(i) = 1 Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                strDate = Trim(.TextMatrix(i, Col.C4核对时间))
                If strDate = "" Then
                    strDate = "Null"
                Else
                    strDate = "To_Date('" & .TextMatrix(i, Col.C4核对时间) & "','YYYY-MM-DD HH24:MI:SS')"
                End If
                arrSQL(UBound(arrSQL)) = "zl_票据使用明细_check(" & .TextMatrix(i, Col.C8ID) & "," & ZVal(Val(.TextMatrix(i, Col.C6核对结果))) & _
                                        ",'" & .TextMatrix(i, Col.C5核对人) & "','" & .TextMatrix(i, Col.C7备注) & "'," & strDate & ")"
            End If
        Next
    End With
    
    On Error GoTo errH
    If UBound(arrSQL) >= 0 Then
        gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
        Next
        
        '检查是否需要填写整体核对记录
        strSQL = "Select Nvl(Sum(Decode(核对结果, Null, 1, 0)), 0) As 未核对数, Count(Distinct 号码) As 已使用数" & vbNewLine & _
                "From 票据使用明细" & vbNewLine & _
                "Where 领用id = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng领用ID)
        If rsTmp!未核对数 = 0 And rsTmp!已使用数 = mdblGiveCount Then
            bytAllChecked = 1
            strSQL = "Select Count(ID) 不相符数 From 票据使用明细 Where 领用id = [1] And 核对结果 <> 原因"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng领用ID)
            If rsTmp!不相符数 = 0 Then bytAllCheckOK = 1
        End If
                
        If bytAllChecked = 1 Then
            strSQL = "zl_票据领用记录_check(" & mlng领用ID & "," & bytAllCheckOK & ",'" & UserInfo.姓名 & "',Null,1)"
        Else
            '取消整体核对
            strSQL = "zl_票据领用记录_check(" & mlng领用ID & ",Null,Null,Null,Null)"
        End If
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            
       gcnOracle.CommitTrans: blnTrans = False
    End If
    SaveData = True
    
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans: blnTrans = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
      Call RefreshCustomTime
End Sub

Private Sub cmdSave_Click()
    Dim i As Long
    
    If SaveData Then
        With mshDetail
            For i = 1 To .Rows - 1
                If .RowData(i) = 1 Then .RowData(i) = 0
            Next
        End With
        cmdSave.Enabled = False
    End If
End Sub

Private Sub dtpEndDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub dtpStartDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub dtpEndDate_Change()
    If dtpEndDate.Value < dtpStartDate.Value Then dtpStartDate.Value = dtpEndDate.Value
End Sub
Private Sub dtpStartDate_Change()
    If dtpStartDate.Value > dtpEndDate.Value Then dtpEndDate.Value = dtpStartDate.Value
End Sub

Private Sub Form_Activate()
    If mshDetail.Rows > 1 Then Call SetRow(1)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

 

Private Sub mshDetail_Click()
    If mbytInFun = 0 Or mshDetail.Row = 0 Then Exit Sub
    
    With mshDetail
        Select Case .Col
            Case Col.C6核对结果
                If .TextMatrix(.Row, .Col) <> "" Then
                    Call zlControl.CboLocate(cboResult, zlCommFun.GetNeedName(.TextMatrix(.Row, .Col)))
                Else
                    Call zlControl.CboLocate(cboResult, zlCommFun.GetNeedName(.TextMatrix(.Row, Col.C3使用情况)))
                End If
                Call SetCboResult
            Case Else
        End Select
    End With
End Sub

Private Sub mshDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    If mbytInFun = 0 Or mshDetail.Row = 0 Then Exit Sub
    
    If KeyCode = vbKeyDelete Then
        With mshDetail
            Select Case .Col
                Case Col.C6核对结果
                    Call SetUnChecked(.Row)
                Case Col.C7备注
                    Call SetChecked(.Row, Col.C7备注, "")
                Case Else
                
            End Select
        End With
    End If
End Sub

Private Sub mshDetail_KeyPress(KeyAscii As Integer)
    If mbytInFun = 0 Or mshDetail.Row = 0 Then Exit Sub
    
    With mshDetail
        If KeyAscii = vbKeyReturn Then
            KeyAscii = 0
            If .Row = .Rows - 1 And (.Col = Col.C7备注 Or .Col = Col.C6核对结果 And .TextMatrix(.Row, .Col) = "") Then
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                If .Col = Col.C7备注 Then
                    .Row = .Row + 1
                    .Col = Col.C6核对结果
                Else
                    .Col = .Col + 1
                End If
            End If
        Else
            Select Case .Col
                Case Col.C6核对结果
                    Call SetCboResult
                    Call cboResult_KeyPress(KeyAscii)
                Case Col.C7备注
                    If .TextMatrix(.Row, Col.C6核对结果) <> "" Then
                        txtInput.Text = Chr(KeyAscii)
                        txtInput.SelStart = 2
                        Call SetTxtInput
                    End If
                Case Else
                
            End Select
        End If
    End With
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If zlCommFun.ActualLen(txtInput.Text) > txtInput.MaxLength Then
            MsgBox "最多只允许输入" & txtInput.MaxLength & "个字符!", vbInformation, gstrSysName
            Exit Sub
        End If
        If InStr(1, txtInput.Text, "'") > 0 Then
            'MsgBox "注意:单引号是系统禁止输入的特殊字符!", vbInformation, gstrSysName
            Beep
            Beep
            Exit Sub
        End If
        
        With mshDetail
            Call SetChecked(.Row, Col.C7备注, Trim(txtInput.Text))
            txtInput.Visible = False
            .SetFocus  '调用lostfocus
            If .Row = .Rows - 1 Then
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                .Row = .Row + 1
                .Col = Col.C6核对结果
            End If
        End With
    End If
End Sub

Private Sub cmdDistant_Click()
    Dim lngRow As Long, bln提醒 As Boolean
    Dim lng前缀 As Long
    
    MousePointer = vbHourglass
    lng前缀 = Len(mstr前缀文本) + 1
    With mshDetail
        lngRow = .Row + 1
        
        While True
            If lngRow > .Rows - 1 Then
                '最后一行
                If bln提醒 = False Then
                    If .Row = 1 Then
                        MsgBox "往下未发现断号情况。", vbInformation, gstrSysName
                        MousePointer = vbDefault
                        Exit Sub
                    Else
                        If MsgBox("往下未发现断号的情况，是否从头开始？", vbQuestion Or vbYesNo, gstrSysName) = vbNo Then
                            MousePointer = vbDefault
                            Exit Sub
                        End If
                    End If
                    bln提醒 = True
                    lngRow = 1
                Else
                    MsgBox "往下未发现断号情况。", vbInformation, gstrSysName
                    MousePointer = vbDefault
                    Exit Sub
                End If
            End If
            
            If lngRow > 1 Then
                If Val(Mid(.TextMatrix(lngRow - 1, 0), lng前缀)) < Val(Mid(.TextMatrix(lngRow, 0), lng前缀)) - 1 Then
                    '出现断号
                    If .RowHeight(lngRow) = 0 Then
                        If MsgBox("注意:" & vbCrLf & "   已经查找到了断号,但不在当前时间范围内,是否进行定位?", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                             If cbo使用日期.Visible Then cbo使用日期.ListIndex = 0:
                        Else
                            Exit Sub
                        End If
                    End If
                    Call SetRow(lngRow)
                    MousePointer = vbDefault
                    Exit Sub
                End If
            End If
            lngRow = lngRow + 1
        Wend
     End With
End Sub

Private Sub cmdFind_Click()
'查找指定号码
    Dim strFind As String
    Dim lngRow As Long
    
    If txt号码.Text = "" Then Exit Sub
    If Len(txt号码.Text) > Len(mshDetail.TextMatrix(1, 0)) Then Exit Sub
    
    '把长度补齐
    strFind = UCase(Mid(mshDetail.TextMatrix(1, 0), 1, Len(mshDetail.TextMatrix(1, 0)) - Len(txt号码.Text)) & txt号码.Text)
    With mshDetail
        For lngRow = 1 To mshDetail.Rows - 1
            If mshDetail.TextMatrix(lngRow, 0) = strFind Then
                If .RowHeight(lngRow) = 0 Then
                    If MsgBox("注意:" & vbCrLf & "   你所查找的号码不在当前时间范围内,是否还要进行定位!", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                         If cbo使用日期.Visible Then cbo使用日期.ListIndex = 0:
                    Else
                        Exit Sub
                    End If
                End If
                Call SetRow(lngRow)
                Exit Sub
            End If
        Next
    End With
    MsgBox "未找到号码为 " & strFind & " 的使用记录。", vbInformation, gstrSysName
End Sub

Private Sub cmdOK_Click()
    If mbytInFun = 1 Then
        Call SaveData
    End If
    Unload Me
End Sub

Private Sub SetHeader()
    Dim strHead As String, arrTmp As Variant, i As Long
    
    With mshDetail
        If mbytInFun = 0 Then
            .SelectionMode = flexSelectionByRow
            .Row = 0: .Col = 0: .RowSel = 0: .ColSel = .Cols - 1
        Else
            .SelectionMode = flexSelectionFree
            .BackColorSel = &HE7CFBA
        End If
                
        If mbytInFun = 0 And Not mblnViewCheck Then
            strHead = "号码,1,1000|使用时间,1,1800|使用人,4,800|使用情况,1,1000"
        Else
            strHead = "号码,1,1000|使用时间,1,1800|使用人,4,800|使用情况,1,1000|核对时间,1,1800|核对人,4,800|核对结果,1,1000|备注,1,2000|ID,1,0"
        End If
        arrTmp = Split(strHead, "|")
        
        .Cols = UBound(arrTmp) + 1
        For i = 0 To UBound(arrTmp)
            .TextMatrix(0, i) = Split(arrTmp(i), ",")(0)
            .ColAlignment(i) = Split(arrTmp(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(arrTmp(i), ",")(2)
            .ColAlignmentFixed(i) = 4
        Next
    End With
End Sub

Private Sub Form_Load()

    If mbytInFun = 0 And Not mblnViewCheck Then Me.Width = 7000
    RestoreWinState Me, App.ProductName
    
    Me.Caption = IIf(mbytInFun = 0, "票据明细清单", "核对票据明细")
    Call SetHeader
    Call RestoreFlexState(mshDetail, Me.Caption)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
    Call SaveFlexState(mshDetail, Me.Caption)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picFilter.Width = cbo使用日期.Left + cbo使用日期.Width + IIf(picTimeRange.Visible, picTimeRange.Width, 50) + 50
    lbl提示(0).Top = IIf(picTimeRange.Visible, picFilter.Height + picFilter.Top + 50, picFilter.Top + (picFilter.Height - lbl提示(0).Height) \ 2)
    lbl提示(0).Left = IIf(picTimeRange.Visible, picFilter.Left, picFilter.Left + picFilter.Width + 50)
    If picTimeRange.Visible Then
        mshDetail.Top = lbl提示(0).Height + lbl提示(0).Top + 50
    Else
        mshDetail.Top = picFilter.Height + picFilter.Top + 50
    End If
    mshDetail.Height = Me.ScaleHeight - mshDetail.Top - 120
    If Me.ScaleWidth > 3000 Then
        fraCMD.Left = Me.ScaleWidth - fraCMD.Width - 120
        mshDetail.Width = fraCMD.Left - mshDetail.Left - 120
    End If
End Sub


Public Sub ShowMe(ByVal frmOwner As Form, ByVal bytInFun As Byte, ByVal blnViewCheck As Boolean, ByVal blnNOMoved As Boolean, _
    ByVal lng票种 As Long, ByVal lng领用ID As Long, ByVal str前缀 As String, _
    Optional strCondition As String, Optional lng原因 As Long, Optional lng性质 As Long, Optional str使用人 As String, Optional str提示 As String)
    '参数:bytInFun:0-查看票据明细,1-核对票据明细
    '   blnViewCheck:当bytInFun=0时,是否显示核对相关字段
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim i As Long, strNOs As String, strResult As String, arrTmp As Variant
    Dim strMinDate As String, strMaxDate As String
    mbytInFun = bytInFun
    mblnViewCheck = blnViewCheck
    mlng票种 = lng票种
    mlng领用ID = lng领用ID
    mstr前缀文本 = str前缀
    
    strSQL = "Select 号码, To_Char(使用时间, 'yyyy-mm-dd hh24:mi:ss') As 使用时间, 使用人," & vbNewLine & _
            "       Decode(原因, 1, '1-正常使用', 2, '2-作废收回', 3, '3-重打发出', 4, '4-重打收回', '5-报损') As 使用情况," & vbNewLine & _
            "       To_Char(核对时间, 'yyyy-mm-dd hh24:mi:ss') As 核对时间, 核对人, Decode(核对结果, 1, '1-正常使用', 2, '2-作废收回', 3, '3-重打发出', 4, '4-重打收回', 5,'5-报损','') as 核对结果, 备注, ID" & vbNewLine & _
            "From " & IIf(blnNOMoved, "H", "") & "票据使用明细" & vbNewLine & _
            "Where 领用id = [1] " & strCondition & vbNewLine & _
            "Order By 号码"
    If mbytInFun = 0 And Not mblnViewCheck Then
        strSQL = "Select 号码, To_Char(使用时间, 'yyyy-mm-dd hh24:mi:ss') As 使用时间, 使用人," & vbNewLine & _
            "       Decode(原因, 1, '1-正常使用', 2, '2-作废收回', 3, '3-重打发出', 4, '4-重打收回', '5-报损') As 使用情况" & vbNewLine & _
            "From " & IIf(blnNOMoved, "H", "") & "票据使用明细" & vbNewLine & _
            "Where 领用id = [1] " & strCondition & vbNewLine & _
            "Order By 号码"
    End If
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng领用ID, lng原因, lng性质, str使用人)
    
    Dim strTemp As String, str使用日期 As String
    
    '其实,如果没有使用明细,菜单已禁用,不会调用此过程
    If rsTmp.RecordCount > 0 Then
        For i = 1 To rsTmp.RecordCount
            If InStr(1, strNOs & ",", "," & rsTmp!号码 & ",") = 0 Then strNOs = strNOs & "," & rsTmp!号码
            strTemp = "|" & Format(rsTmp!使用时间, "yyyy-MM-DD")
            If InStr(1, str使用日期 & "|", strTemp & "|") = 0 Then str使用日期 = str使用日期 & strTemp
            rsTmp.MoveNext
        Next
        i = 0
        If strNOs <> "" Then
            strNOs = Mid(strNOs, 2)
            i = UBound(Split(strNOs, ",")) + 1
        End If
        lbl提示(0).Caption = str提示 & IIf(str提示 = "", "", ",") & "共计" & i & "张票据."
        lbl提示(0).Tag = lbl提示(0).Caption
    End If
    Set mshDetail.DataSource = rsTmp
    
    Dim varData As Variant
    Dim j As Long
    If str使用日期 <> "" Then str使用日期 = Mid(str使用日期, 2)
    varData = Split(str使用日期, "|")
    
    mblnUnClick = True
    '按日期重小到大排序
    cbo使用日期.AddItem "所有": cbo使用日期.ListIndex = cbo使用日期.NewIndex
    cbo使用日期.AddItem "时间范围"
    mblnUnClick = False
    For i = 0 To UBound(varData)
        If varData(i) <> "" Then
            For j = i + 1 To UBound(varData)
                If varData(j) < varData(i) Then
                    strTemp = varData(i)
                    varData(i) = varData(j)
                     varData(j) = strTemp
                End If
            Next
            If varData(i) < strMinDate Or strMinDate = "" Then strMinDate = varData(i)
            If varData(i) > strMaxDate Then strMaxDate = varData(i)
            cbo使用日期.AddItem varData(i)
        End If
    Next
    dtpStartDate.MaxDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59")
    dtpEndDate.MaxDate = dtpStartDate.MaxDate
    If strMinDate <> "" And IsDate(strMinDate) Then
        dtpStartDate.MinDate = Format(CDate(strMinDate), "yyyy-mm-dd 00:00:00")
        dtpStartDate.Value = dtpStartDate.MinDate
        dtpEndDate.MinDate = dtpStartDate.MinDate
        If IsDate(strMaxDate) Then
            dtpEndDate.Value = Format(CDate(strMaxDate), "yyyy-mm-dd 23:59:59")
        Else
            dtpEndDate.Value = Format(dtpStartDate.MinDate, "yyyy-mm-dd 23:59:59")
        End If
    End If
    cboResult.Visible = False
    txtInput.Visible = False
    If mbytInFun = 0 Then
        cmdOK.Caption = "退出(&X)"
        cmdOK.Cancel = True
        cmdCancel.Visible = False
        cmdSave.Visible = False
        cmdAllDO(0).Visible = False
        cmdAllDO(1).Visible = False
        picFilter.Visible = False
        lbl提示(0).Left = picFilter.Left
    Else
        strResult = " ,1-正常使用,2-作废收回,3-重打发出,4-重打收回,5-报损"
        arrTmp = Split(strResult, ",")
        For i = 0 To UBound(arrTmp)
            cboResult.AddItem arrTmp(i)
        Next
        Call zlControl.CboSetWidth(cboResult.hwnd, 800)
        
        mdblGiveCount = 0
        strSQL = "Select To_Number(Replace(终止号码, 前缀文本)) - To_Number(Replace(开始号码, 前缀文本))+1 张数" & vbNewLine & _
                "From 票据领用记录" & vbNewLine & _
                "Where ID = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng领用ID)
        If rsTmp.RecordCount > 0 Then mdblGiveCount = rsTmp!张数
        picFilter.Visible = True
    End If
    frmBillUses.Show vbModal, frmOwner
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
 
Private Sub mshDetail_DblClick()
    Dim strReportNO As String, strInvoiceNO As String
    
    With mshDetail
        Select Case .Col
            Case Col.C7备注
                If mbytInFun = 0 Or mshDetail.Row = 0 Then Exit Sub
                If .TextMatrix(.Row, Col.C6核对结果) = "" Then Exit Sub
                
                Call SetTxtInput
                txtInput.Text = .TextMatrix(.Row, .Col)
                Call SelAll(txtInput)
            Case Else
                strReportNO = "ZL" & glngSys \ 100 & "_INSIDE_1501"
                strInvoiceNO = .TextMatrix(.Row, Col.C0号码)
                Call ReportOpen(gcnOracle, glngSys, strReportNO, Me, "票据号=" & strInvoiceNO & "", "票种=" & mlng票种, "ReportFormat=" & mlng票种, 1)
        End Select
    End With
End Sub

Private Sub SetCboResult()
    With mshDetail
        cboResult.Left = .Left + .CellLeft - 15
        cboResult.Top = .Top + .CellTop - 15
        cboResult.Width = .CellWidth + 15
        cboResult.Visible = True
        cboResult.SetFocus
    End With
End Sub

Private Sub SetTxtInput()
    With mshDetail
        txtInput.Left = .Left + .CellLeft - 15
        txtInput.Top = .Top + .CellTop - 15
        txtInput.Width = .CellWidth + 15
        txtInput.Height = .CellHeight
        txtInput.Visible = True
        txtInput.SetFocus
    End With
End Sub

Private Sub mshDetail_LeaveCell()
    If mbytInFun = 0 Or mshDetail.Row = 0 Then Exit Sub
    
    With mshDetail
        If cboResult.Visible Then
            If .TextMatrix(.Row, .Col) <> Trim(cboResult.Text) Then
                If cboResult.ListIndex <= 0 Then
                    Call SetUnChecked(.Row)
                Else
                    Call SetChecked(.Row, Col.C6核对结果, Trim(cboResult.Text), Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss"))
                End If
            End If
        ElseIf txtInput.Visible Then
            If .TextMatrix(.Row, .Col) <> Trim(txtInput.Text) Then
                If zlCommFun.ActualLen(txtInput.Text) > txtInput.MaxLength Then
                    MsgBox "最多只允许输入" & txtInput.MaxLength & "个字符!", vbInformation, gstrSysName
                    Exit Sub
                End If
                If InStr(1, txtInput.Text, "'") > 0 Then
                    'MsgBox "注意:单引号是系统禁止输入的特殊字符!", vbInformation, gstrSysName
                    Beep
                    Beep
                    Exit Sub
                End If
                Call SetChecked(.Row, Col.C7备注, Trim(txtInput.Text))
            End If
        End If
    End With
End Sub

Private Sub mshDetail_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mshDetail.MouseRow = 0 Then
        mshDetail.MousePointer = 99
    Else
        mshDetail.MousePointer = 0
    End If
End Sub

Private Sub mshDetail_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim lngCol As Long
    
    With mshDetail
        If Button = 1 And .MousePointer = 99 Then
            lngCol = .MouseCol
            If .TextMatrix(0, lngCol) = "" Then Exit Sub
            
            .ColData(lngCol) = (.ColData(lngCol) + 1) Mod 2
            
            .Redraw = False
            .Col = lngCol: .ColSel = lngCol   '排序依据
            .Sort = IIf(.ColData(lngCol) = 1, 6, 5)
            .Col = 0
            .ColSel = .Cols - 1
            .Redraw = True
        End If
    End With
End Sub


Private Sub txtInput_LostFocus()
    If txtInput.Visible Then txtInput.Visible = False
End Sub

Private Sub txt号码_GotFocus()
    Call zlControl.TxtSelAll(txt号码)
End Sub

Private Sub txt号码_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call cmdFind_Click
        SelAll txt号码
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub SetRow(ByVal lngRow As Long)
    Dim lngTop As Long
    With mshDetail
        .Row = lngRow
        lngTop = lngRow - 1
        If lngTop < 1 Then lngTop = 1
        If .RowIsVisible(lngTop) = False Then
            .TopRow = lngTop
        End If
        If mbytInFun = 0 Then
            .Col = 0
            .ColSel = .Cols - 1
        Else
            .Col = Col.C6核对结果
        End If
    End With
End Sub


