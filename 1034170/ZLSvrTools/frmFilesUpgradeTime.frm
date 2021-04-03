VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmFilesUpgradeTime 
   Caption         =   "预升级时间设置"
   ClientHeight    =   3660
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7800
   Icon            =   "frmFilesUpgradeTime.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   7800
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Cmd删除 
      Caption         =   "删除(&D)"
      Height          =   350
      Left            =   1155
      TabIndex        =   3
      Top             =   3270
      Width           =   1100
   End
   Begin VB.CommandButton Cmd添加 
      Caption         =   "添加(&A)"
      Height          =   350
      Left            =   45
      TabIndex        =   2
      Top             =   3270
      Width           =   1100
   End
   Begin VB.CommandButton cmd保存 
      Caption         =   "保存(&S)"
      Height          =   350
      Left            =   5520
      TabIndex        =   1
      Top             =   3270
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "关闭(&C)"
      Height          =   350
      Left            =   6675
      TabIndex        =   0
      Top             =   3270
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfShift 
      Height          =   2145
      Left            =   675
      TabIndex        =   4
      Top             =   375
      Width           =   3990
      _cx             =   7038
      _cy             =   3784
      Appearance      =   1
      BorderStyle     =   0
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
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483626
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
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   330
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
End
Attribute VB_Name = "frmFilesUpgradeTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mblnOK As Boolean
Private WithEvents mclsVsfShift As clsVsf
Attribute mclsVsfShift.VB_VarHelpID = -1
Private mstrOldTimes As String
'关闭
Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmd保存_Click()
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lngLoop As Long
    Dim strTemp As String
    On Error GoTo errHand
    
    For lngLoop = 1 To vsfShift.Rows - 1
        If Len(strTemp) = 0 Then
            If vsfShift.TextMatrix(lngLoop, 1) <> "" Then
                strTemp = vsfShift.TextMatrix(lngLoop, 1)
            End If
        Else
            If vsfShift.TextMatrix(lngLoop, 1) <> "" Then
                strTemp = strTemp & "," & Format(vsfShift.TextMatrix(lngLoop, 1), "HH:mm")
            End If
        End If
    Next
    
    If strTemp <> mstrOldTimes Then
        Set rsTmp = New ADODB.Recordset
        gstrSQL = "Select 项目,内容 From zlRegInfo where 项目 = '客户端预升级时间点'"
        Call OpenRecordset(rsTmp, gstrSQL, Me.Caption)
        
        If rsTmp.EOF = False Then
            strSQL = "Update zlRegInfo Set 内容='" & strTemp & "' Where 项目='客户端预升级时间点'"
            gcnOracle.Execute strSQL
        Else
            strSQL = "INSERT INTO zlRegInfo (项目,行号,内容) VALUES ('客户端预升级时间点',Null,'" & strTemp & "')"
            gcnOracle.Execute strSQL
        End If
        
        If MsgBox("是否重新对所有站点进行预升级?" & vbNewLine & "是:所有站点预升级的完成状态将被清除。" & vbNewLine & "否:可以手工进行预升级完成状态的调整。", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
            strSQL = "Zl_Zlclients_Control(3,Null,Null,Null,Null,Null,0)"
            Call ExecuteProcedure(strSQL, Me.Caption)
        End If
        
        mblnOK = True
    Else
        mblnOK = False
    End If
    
    Unload Me
  Exit Sub
errHand:
    MsgBox err.Description, vbInformation + vbDefaultButton1, gstrSysName
End Sub

Private Sub Cmd删除_Click()
    If vsfShift.Rows > 1 Then
        If vsfShift.Row > 1 Then
            Call mclsVsfShift.DeleteRow(vsfShift.Row)
        Else
            Call mclsVsfShift.DeleteRow(vsfShift.Rows - 1)
        End If
    End If
End Sub

Private Sub Cmd添加_Click()
    If vsfShift.Rows = 2 And vsfShift.TextMatrix(1, 1) = "" Then
        vsfShift.TextMatrix(1, 0) = "1"
        vsfShift.TextMatrix(1, 1) = Format("12:30", "HH:mm")
    Else
        Call mclsVsfShift.AutoAddRow(vsfShift.MouseRow, vsfShift.MouseCol)
    End If
End Sub

Private Sub Form_Load()
    Call InitVSF
    Call LoadVSF
End Sub

Private Sub Form_Resize()
    With vsfShift
        .Top = 15
        .Left = 15
        .Width = ScaleWidth - 30
        .Height = ScaleHeight - 450
    End With
    mclsVsfShift.AppendRows = True
    
    With Cmd添加
        .Top = ScaleHeight - .Height - 30
        .Left = 15
    End With
    
    With Cmd删除
        .Top = Cmd添加.Top
        .Left = Cmd添加.Left + Cmd添加.Width + 30
    End With
    
    With cmdCancel
        .Top = Cmd添加.Top
        .Left = ScaleWidth - .Width - 30
    End With
    
    With cmd保存
        .Top = Cmd添加.Top
        .Left = cmdCancel.Left - .Width - 30
    End With
End Sub

Private Sub InitVSF()
     
     Set mclsVsfShift = New clsVsf
     
    With mclsVsfShift
        Call .Initialize(Me.Controls, vsfShift, True, False)
        Call .ClearColumn
        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, , "", False)
        Call .AppendColumn("预设客户端后台升级时间", 1670, flexAlignCenterCenter, flexDTString, "HH:mm", , True)
        
        .AppendRows = True
        .IndicatorMode = 2
        Call .InitializeEdit(True, True, True)
        Call .InitializeEditColumn(.ColIndex("预设客户端后台升级时间"), True, vbVsfEditDate, , , , "99:99")
        
    End With
End Sub

Private Sub LoadVSF()
    Dim rsTmp As ADODB.Recordset
    Dim strTemp() As String
    Dim i As Integer
    Set rsTmp = New ADODB.Recordset
    
    mstrOldTimes = ""
    gstrSQL = "Select 项目,内容 From zlRegInfo where 项目 like '客户端预升级时间点'"
    Call OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    
    With vsfShift
'        .Redraw = flexRDNone
        If rsTmp.RecordCount = 1 Then
            If Nvl(rsTmp!内容) <> "" Then
                strTemp = Split(Nvl(rsTmp!内容), ",")
                .Rows = UBound(strTemp) + 2
                For i = 0 To UBound(strTemp)
                    
                    .TextMatrix(i + 1, 0) = i + 1
                    .TextMatrix(i + 1, 1) = strTemp(i)
                Next
            Else
                For i = 0 To .Cols - 1
                    .TextMatrix(1, i) = ""
                Next
'                .Redraw = flexRDBuffered
                Exit Sub
            End If
        Else
            For i = 0 To .Cols - 1
                .TextMatrix(1, i) = ""
            Next
'            .Redraw = flexRDBuffered
            Exit Sub
        End If
        For i = 1 To vsfShift.Rows - 1
            If Len(mstrOldTimes) = 0 Then
                If vsfShift.TextMatrix(i, 1) <> "" Then
                    mstrOldTimes = vsfShift.TextMatrix(i, 1)
                End If
            Else
                If vsfShift.TextMatrix(i, 1) <> "" Then
                    mstrOldTimes = mstrOldTimes & "," & Format(vsfShift.TextMatrix(i, 1), "HH:mm")
                End If
            End If
        Next
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mclsVsfShift = Nothing
End Sub


Private Sub mclsVsfShift_AfterDeleteCell(ByVal Row As Long, ByVal Col As Long)
    mclsVsfShift.AppendRows = True
End Sub

Private Sub mclsVsfShift_AfterDeleteRow(ByVal Row As Long, ByVal Col As Long)
    mclsVsfShift.AppendRows = True
End Sub

Private Sub mclsVsfShift_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    With vsfShift
            Cancel = (.TextMatrix(Row, .ColIndex("预设客户端后台升级时间")) = "" Or .TextMatrix(Row, .ColIndex("预设客户端后台升级时间")) = "    -  -     :  " Or .TextMatrix(Row, .ColIndex("预设客户端后台升级时间")) = "__:__")
    End With
    mclsVsfShift.AppendRows = True
End Sub

Private Sub vsfShift_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Call mclsVsfShift.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
    Dim lngCol As Long
    lngCol = vsfShift.ColIndex("预设客户端后台升级时间")
   
    With vsfShift
        Select Case NewCol
        Case lngCol
            If mclsVsfShift.AllowColEdit(NewCol) = False Or mclsVsfShift.AllowEdit = False Then Exit Sub
            If IsDate(.TextMatrix(NewRow, NewCol)) = False Then
                
                If NewRow > 1 Then
                    If IsDate(.TextMatrix(NewRow - 1, NewCol)) Then
                        .TextMatrix(NewRow, NewCol) = GetUpgradeTime(.TextMatrix(NewRow - 1, NewCol)) & ":30"
                    Else
                        .TextMatrix(NewRow, NewCol) = Format(CurrentDate, "HH:mm")
                    End If
                Else
                    .TextMatrix(NewRow, NewCol) = Format(CurrentDate, "HH:mm")
                End If
            End If
        End Select
    End With
    mclsVsfShift.AppendRows = True
End Sub

Private Sub vsfShift_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsfShift.AppendRows = True
End Sub

Private Sub vsfShift_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    mclsVsfShift.AppendRows = True
End Sub

Private Sub vsfShift_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    With vsfShift
    
        If mclsVsfShift.CellButtonClick(Row, Col) Then
            Call mclsVsfShift.SetFocus(, , True)
        End If
    End With
End Sub

Private Sub vsfShift_KeyDown(KeyCode As Integer, Shift As Integer)
    Call mclsVsfShift.KeyDown(KeyCode, Shift)
End Sub

Private Sub vsfShift_KeyPress(KeyAscii As Integer)
    '编辑处理
    Call mclsVsfShift.KeyPress(KeyAscii)
End Sub

Private Sub vsfShift_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    '编辑处理
    Call mclsVsfShift.KeyPressEdit(KeyAscii)
End Sub

Private Sub vsfShift_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case 1
        Call mclsVsfShift.AutoAddRow(vsfShift.MouseRow, vsfShift.MouseCol)
    End Select
End Sub

Private Sub vsfShift_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    '编辑处理
    Call mclsVsfShift.EditSelAll
End Sub

Private Sub vsfShift_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '编辑处理
    Call mclsVsfShift.BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsfShift_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '编辑处理
    Call mclsVsfShift.ValidateEdit(Col, Cancel)
End Sub

Private Sub vsfShift_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case Col
    Case vsfShift.ColIndex("预设客户端后台升级时间")
        '只有服务器列和升级列才能更改
    Case Else
        '其他列不能更改
        Cancel = True
    End Select
End Sub

Private Function GetUpgradeTime(ByVal strTemp As String) As String
    Dim i As Integer
    Dim strTime As String
    If strTemp = "" Then
        GetUpgradeTime = Format(CurrentDate, "HH:mm")
        Exit Function
    End If
    
    i = InStrRev(strTemp, ":")
    If i > 0 Then
        strTime = Left(strTemp, i)
        strTime = Val(strTime) + 1
        If Val(strTime) >= 24 Then
            strTime = "00"
        End If
        
        GetUpgradeTime = strTime
        Exit Function
    End If
    GetUpgradeTime = Format(CurrentDate, "HH:mm")
End Function
