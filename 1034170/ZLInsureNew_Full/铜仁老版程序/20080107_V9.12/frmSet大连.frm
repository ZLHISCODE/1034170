VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.1#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmSet大连 
   AutoRedraw      =   -1  'True
   Caption         =   "设置"
   ClientHeight    =   5055
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7650
   Icon            =   "frmSet大连.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   7650
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   30
      ScaleHeight     =   375
      ScaleWidth      =   7575
      TabIndex        =   7
      Top             =   3960
      Width           =   7575
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1095
         MaxLength       =   2
         TabIndex        =   11
         Text            =   "1"
         Top             =   30
         Width           =   360
      End
      Begin VB.CheckBox chk开发区 
         Caption         =   "开发区(&K)"
         Height          =   255
         Left            =   2220
         TabIndex        =   10
         Top             =   60
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.CheckBox chk实时 
         Caption         =   "门诊明细时实上传(&M)"
         Height          =   285
         Index           =   0
         Left            =   3450
         TabIndex        =   9
         Top             =   45
         Width           =   2085
      End
      Begin VB.CheckBox chk实时 
         Caption         =   "住院明细时实上传(&Z)"
         Height          =   285
         Index           =   1
         Left            =   5520
         TabIndex        =   8
         Top             =   45
         Width           =   2085
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "当前串口(&D)"
         Height          =   180
         Index           =   3
         Left            =   0
         TabIndex        =   13
         Top             =   90
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "号串口"
         Height          =   180
         Index           =   4
         Left            =   1500
         TabIndex        =   12
         Top             =   90
         Width           =   540
      End
   End
   Begin ZL9BillEdit.BillEdit mshBill 
      Height          =   2685
      Left            =   120
      TabIndex        =   2
      Top             =   1170
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   4736
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Cols            =   2
      RowHeight0      =   315
      RowHeightMin    =   315
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Top             =   4380
      Width           =   7755
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6360
      TabIndex        =   4
      Top             =   4590
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5100
      TabIndex        =   3
      Top             =   4590
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   0
      Left            =   30
      TabIndex        =   5
      Top             =   690
      Width           =   7665
   End
   Begin MSComctlLib.TabStrip tabSel 
      Height          =   3105
      Left            =   60
      TabIndex        =   1
      Top             =   810
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   5477
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   60
      Picture         =   "frmSet大连.frx":000C
      Top             =   150
      Width           =   480
   End
   Begin VB.Label lbl 
      Caption         =   "设置设备的串口号及设置窗口是否默认为开发区,并将其收费类型与相关的医保项目相对应"
      Height          =   315
      Left            =   540
      TabIndex        =   0
      Top             =   390
      Width           =   7125
   End
End
Attribute VB_Name = "frmSet大连"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnReturn As Boolean
Private mlng医保中心 As Long
Private mlng险类 As Long
Private Enum mColHead
    收费类别 = 0
    保费项目
    分类项目
End Enum
Private Sub chk开发区_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{Tab}", 1
    End If
End Sub


Private Sub chk实时_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{Tab}", 1
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnReturn = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngRow As Long
    
    If Trim(txtEdit) = "" Then Exit Sub
    SaveRegInFor g公共模块, "操作", "端口号", Me.txtEdit
    SaveRegInFor g公共模块, "操作", "开发区", Me.chk开发区.Value
    If Val(txtEdit) = 0 Then
        gintComPort_大连 = 1
    Else
        gintComPort_大连 = Val(txtEdit)
    End If
    gblnKFQCom_大连 = IIf(chk开发区.Value = 1, True, False)
    gintComPort = txtEdit.Text
        
    '删除已经数据
    gcnOracle.BeginTrans
    On Error GoTo ErrHand
    
    gstrSQL = "zl_保险参数_Delete(" & mlng险类 & ",NUll)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    With MshBill
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, mColHead.收费类别) <> "" Then
                '新增参数数据
                gstrSQL = "zl_保险参数_Insert(" & mlng险类 & ",null,'" & .TextMatrix(lngRow, mColHead.收费类别) & "' ,'" & .TextMatrix(lngRow, mColHead.保费项目) & ";" & .TextMatrix(lngRow, mColHead.分类项目) & "'," & lngRow + 2 & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            End If
        Next
    End With
    '保存
    
    ' gstrSQL = "zl_保险参数_Insert(" & mlng险类 & ",null,'门诊明细时实上传' ,'" & IIf(chk实时(0).Value = 1, "1", "0") & "'," & 1 & ")"
     
     gstrSQL = "zl_保险参数_Update(" & mlng险类 & ",NULL,'门诊明细时实上传' ,'" & IIf(chk实时(0).Value = 1, "1", "0") & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
'     gstrSQL = "zl_保险参数_Insert(" & mlng险类 & ",null,'住院明细时实上传' ,'" & IIf(chk实时(1).Value = 1, "1", "0") & "'," & 2 & ")"
    gstrSQL = "zl_保险参数_Update(" & mlng险类 & ",NULL,'住院明细时实上传' ,'" & IIf(chk实时(1).Value = 1, "1", "0") & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gcnOracle.CommitTrans
    mblnReturn = True
    Unload Me
    Exit Sub
ErrHand:
    gcnOracle.RollbackTrans
    Resume
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{Tab}", 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    Dim strReg As String
    mblnReturn = False
    
    Call GetRegInFor(g公共模块, "操作", "端口号", strReg)
    
    If Val(strReg) = 0 Then
        txtEdit.Text = 1
    Else
        txtEdit.Text = Val(strReg)
    End If
    
    Call GetRegInFor(g公共模块, "操作", "开发区", strReg)
    
    If Val(strReg) = 1 Then
        Me.chk开发区.Value = 1
    Else
        Me.chk开发区.Value = 0
    End If
    RestoreWinState Me, App.ProductName
    
    '初始数据
    Call iniData
End Sub

Public Function ShowME(ByVal lng险类 As Long, ByVal lng医保中心 As Long) As Boolean
    mlng医保中心 = lng医保中心
    mlng险类 = lng险类
    
    Me.Show 1
    ShowME = mblnReturn
End Function

Private Sub Form_Resize()
   Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    
    sngTop = 0
    sngBottom = 0
    
    fra(0).Width = ScaleWidth + 50
    With cmdCancel
        .Top = ScaleHeight - .Height - 100
        .Left = ScaleWidth - .Width - 50
    End With
    With cmdOK
        .Top = cmdCancel.Top
        .Left = cmdCancel.Left - 50 - .Width
    End With
    
    fra(1).Width = fra(0).Width
    fra(1).Top = cmdOK.Top - fra(1).Height - 50
    
    With pic
        .Top = fra(1).Top - .Height - 50
        .Width = ScaleWidth - 50
    End With
    With tabSel
        .Width = ScaleWidth - 50
        .Height = pic.Top - .Top - 20
    End With
    With MshBill
        .Top = tabSel.Top + tabSel.Tabs(1).Height + 100
        .Left = tabSel.Left + 100
        .Height = tabSel.Height - tabSel.Tabs(1).Height - 200
        .Width = tabSel.Width - 200
    End With
    MshBill.ZOrder
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub mshBill_EnterCell(Row As Long, Col As Long)
    With MshBill
        Select Case Col
            Case mColHead.保费项目
                MshBill.Clear
                MshBill.AddItem "诊察费"
                MshBill.AddItem "草药费"
                MshBill.AddItem "成药费"
                MshBill.AddItem "西药费"
                MshBill.AddItem "检查费"
                MshBill.AddItem "大检费"
                MshBill.AddItem "治疗费"
                MshBill.AddItem "特殊治疗费"
                If mlng险类 = TYPE_大连开发区 Then
                    'mshBill.AddItem "其它费"
                Else
                   ' mshBill.AddItem "其它费"
                End If
            Case mColHead.分类项目
                MshBill.Clear
                MshBill.AddItem "A中草药费"
                MshBill.AddItem "B中成药费"
                MshBill.AddItem "C西药费"
                MshBill.AddItem "D检查费"
                MshBill.AddItem "E输氧费"
                MshBill.AddItem "F放射费"
                MshBill.AddItem "G手术费"
                MshBill.AddItem "H化验费"
                MshBill.AddItem "I诊疗费"
                MshBill.AddItem "J麻醉费"
                MshBill.AddItem "K床位费"
                MshBill.AddItem "L护理费"
                MshBill.AddItem "M其它费用"
        End Select
    End With
End Sub
Private Sub pic_Resize()
    Err = 0
    On Error Resume Next
    With chk实时(1)
        .Left = pic.ScaleWidth - .Width - 50
        chk实时(0).Left = .Left - chk实时(1).Width - 50
    End With
End Sub

Private Sub tabSel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
            SendKeys "{Tab}", 1
    End If
End Sub
Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys "{Tab}", 1
    End If
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
        zlControl.TxtCheckKeyPress txtEdit, KeyAscii, m数字式
End Sub
Private Function iniData() As Boolean
    '初始数据
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    Dim lngRow As Long
    Dim strTmp As String
    
    '设置页头
    Err = 0
    On Error Resume Next
    strSql = "Select * from 保险中心目录 where 险类=" & mlng险类
    zlDatabase.OpenRecordset rsTmp, strSql, Me.Caption
    If rsTmp.EOF Then
        tabSel.Tabs(1).Caption = "无"
    Else
        tabSel.Tabs(1).Caption = NVL(rsTmp!名称)
    End If
    rsTmp.Close
  
    If mlng险类 = TYPE_大连开发区 Then
        Me.chk开发区.Value = 1
    Else
        Me.chk开发区.Value = 0
    End If
    
    '设置表列头
    Call initGrid
    strSql = "" & _
        "   Select A.类别,b.参数值 From 收费类别 a,(Select * From 保险参数 where 险类=" & mlng险类 & ") b " & _
        "   Where A.类别=b.参数名(+) " & _
        "   order by A.编码 "
    zlDatabase.OpenRecordset rsTmp, strSql, Me.Caption
    With MshBill
        .ClearBill
        If rsTmp.RecordCount = 0 Then
            .Rows = 2
        Else
            .Rows = rsTmp.RecordCount + 1
        End If
        lngRow = 1
        Do While Not rsTmp.EOF
            .TextMatrix(lngRow, mColHead.收费类别) = NVL(rsTmp!类别)
            strTmp = NVL(rsTmp!参数值)
            If InStr(1, strTmp, ";") <> 0 Then
                .TextMatrix(lngRow, mColHead.保费项目) = Split(strTmp, ";")(0)
                .TextMatrix(lngRow, mColHead.分类项目) = Split(strTmp, ";")(1)
            End If
            lngRow = lngRow + 1
            rsTmp.MoveNext
        Loop
        
        strSql = "Select * From 保险参数 where 参数名 in('门诊明细时实上传','住院明细时实上传') and 险类=" & mlng险类
        zlDatabase.OpenRecordset rsTmp, strSql, Me.Caption
        chk实时(0).Value = 1
        chk实时(1).Value = 1
        Do While Not rsTmp.EOF
            Select Case NVL(rsTmp!参数名)
            Case "门诊明细时实上传"
                chk实时(0).Value = IIf(Val(NVL(rsTmp!参数值)) = 1, 1, 0)
            Case "住院明细时实上传"
                chk实时(1).Value = IIf(Val(NVL(rsTmp!参数值)) = 1, 1, 0)
            End Select
            rsTmp.MoveNext
        Loop
        
    End With
    
End Function
Private Sub initGrid()
    With MshBill
        .Active = True
        .Cols = 3
        
        .msfObj.FixedCols = 1
        .AllowAddRow = False
        
        .TextMatrix(0, mColHead.收费类别) = "收费类别"
        .TextMatrix(0, mColHead.保费项目) = "保费项目"
        .TextMatrix(0, mColHead.分类项目) = "分类项目"
        
        
        .ColWidth(mColHead.收费类别) = 1500
        .ColWidth(mColHead.保费项目) = 2000
        .ColWidth(mColHead.分类项目) = 2000
        
        '-1：表示该列可以选择，是布尔型［"√"，" "］
        ' 0：表示该列可以选择，但不能修改
        ' 1：表示该列可以输入，外部显示为按钮选择
        ' 2：表示该列是日期列，外部显示为按钮选择，弹出是日期选择框
        ' 3：表示该列是选择列，外部显示为下拉框选择
        '4:  表示该列为单纯的文本框供用户输入
        '5:  表示该列不允许选择

        .ColData(mColHead.收费类别) = 5
        .ColData(mColHead.保费项目) = 3
        .ColData(mColHead.分类项目) = 3
        
        .ColAlignment(mColHead.收费类别) = flexAlignLeftCenter
        .ColAlignment(mColHead.保费项目) = flexAlignLeftCenter
        .ColAlignment(mColHead.分类项目) = flexAlignLeftCenter
        .PrimaryCol = mColHead.保费项目
        .LocateCol = mColHead.保费项目
    End With
End Sub



