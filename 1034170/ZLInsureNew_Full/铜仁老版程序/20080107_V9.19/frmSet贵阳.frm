VERSION 5.00
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.1#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmSet贵阳 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "保险参数设置"
   ClientHeight    =   4260
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5835
   Icon            =   "frmSet贵阳.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fra 
      Caption         =   "保险参数"
      Height          =   2415
      Left            =   150
      TabIndex        =   1
      Top             =   150
      Width           =   4365
      Begin VB.TextBox txt项目数 
         Height          =   300
         Left            =   2220
         MaxLength       =   40
         TabIndex        =   5
         Top             =   1410
         Width           =   1995
      End
      Begin VB.CheckBox chk收费 
         Caption         =   "门诊采用连续收费(&L)"
         Height          =   255
         Left            =   990
         TabIndex        =   0
         Top             =   330
         Width           =   2055
      End
      Begin VB.CheckBox chk参保前在院 
         Caption         =   "入院时选择参保前在院(&T)"
         Height          =   255
         Left            =   990
         TabIndex        =   3
         Top             =   1050
         Width           =   2385
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "…"
         Height          =   240
         Left            =   3930
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1845
         Width           =   255
      End
      Begin VB.CheckBox chk病种 
         Caption         =   "支持特殊门诊(&T)"
         Height          =   255
         Left            =   990
         TabIndex        =   2
         Top             =   690
         Width           =   1695
      End
      Begin VB.TextBox txtServer 
         Height          =   300
         Left            =   2220
         MaxLength       =   40
         TabIndex        =   7
         Top             =   1815
         Width           =   1995
      End
      Begin VB.Label lbl项目数 
         AutoSize        =   -1  'True
         Caption         =   "门诊最大项目数(&N)"
         Height          =   180
         Left            =   630
         TabIndex        =   4
         Top             =   1485
         Width           =   1530
      End
      Begin VB.Label lbl服务器 
         AutoSize        =   -1  'True
         Caption         =   "医保服务器(&S)"
         Height          =   180
         Left            =   990
         TabIndex        =   6
         Top             =   1890
         Width           =   1170
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   180
         Picture         =   "frmSet贵阳.frx":000C
         Top             =   300
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4650
      TabIndex        =   9
      Top             =   390
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4650
      TabIndex        =   10
      Top             =   870
      Width           =   1100
   End
   Begin ZL9BillEdit.BillEdit Bill 
      Height          =   1485
      Left            =   120
      TabIndex        =   8
      Top             =   2670
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   2619
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
End
Attribute VB_Name = "frmSet贵阳"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'只有所有者才能设置收费类别与发票归属科目编码的对应关系
Dim mlng险类 As Long
Dim mblnOK As Boolean
Dim mblnChange As Boolean     '是否改变了

Private Sub Bill_cboClick(ListIndex As Long)
    If Bill.Active = False Then Exit Sub
    Bill.TextMatrix(Bill.Row, 2) = Bill.CboText
End Sub

Private Sub cmdSelect_Click()
    Dim strServer As String
    
    strServer = GetComputer(Me, "选择医保服务器")
    If strServer <> "" Then
        txtServer.Text = strServer
        mblnChange = True
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Me.ActiveControl.Name <> "Bill" Then SendKeys "{Tab}", 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim colPara As New Collection
    Dim lngCount As Long
    
    If txtServer.Text = "" Then
        MsgBox "医保服务器名不能为空。", vbInformation, gstrSysName
        txtServer.SetFocus
        Exit Sub
    End If
    If IsNumeric(txt项目数.Text) = False Then
        MsgBox "请输入正确的项目数。", vbInformation, gstrSysName
        txt项目数.SetFocus
        Exit Sub
    End If
    If zlCommFun.StrIsValid(txtServer.Text, txtServer.MaxLength) = False Then
        txtServer.SetFocus
        Exit Sub
    End If
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    '删除已经数据
    gstrSQL = "zl_保险参数_Delete(" & mlng险类 & ",null)"
    Call ExecuteProcedure(Me.Caption)
    
    '新增参数数据
    '这一部分参数不区分中心
    colPara.Add "null,'门诊连续收费','" & chk收费.Value
    colPara.Add "null,'支持特殊门诊','" & chk病种.Value
    colPara.Add "null,'入院时选择参保前在院','" & chk参保前在院.Value
    colPara.Add "null,'医保服务器','" & txtServer.Text
    colPara.Add "null,'门诊最大项目数','" & txt项目数.Text
    
    For lngCount = 1 To colPara.Count
        gstrSQL = "zl_保险参数_Insert(" & mlng险类 & "," & colPara(lngCount) & "'," & lngCount & ")"
        Call ExecuteProcedure(Me.Caption)
    Next
    
    '归属科目编码的保存
    For lngCount = 1 To Bill.Rows - 1
        gstrSQL = "zl_保险参数_Insert(" & mlng险类 & ",null,'" & Bill.TextMatrix(lngCount, 0) & "','" & Mid(Bill.TextMatrix(lngCount, 2), 1, 2) & "'," & lngCount + 5 & ")"
        Call ExecuteProcedure(Me.Caption)
    Next
    
    gcnOracle.CommitTrans
    mblnChange = False
    mblnOK = True
    Unload Me
    Exit Sub

errHandle:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
    Call SaveErrLog
End Sub

Private Sub txtServer_Change()
    mblnChange = True
End Sub

Private Sub txtServer_GotFocus()
    zlControl.TxtSelAll txtServer
End Sub

Public Function 参数设置(ByVal lng险类 As Long) As Boolean
'功能：设置我们中联医保所需要的参数
    Dim rsTemp As New ADODB.Recordset
    Dim str参数值 As String
    
    mblnOK = False
    mlng险类 = lng险类
    
    rsTemp.CursorLocation = adUseClient
    gstrSQL = "select 参数名,参数值 from 保险参数 where 险类=" & lng险类 & " and 中心 is null"
    Call OpenRecordset(rsTemp, Me.Caption)
    
    Do Until rsTemp.EOF
        Select Case rsTemp("参数名")
            Case "门诊连续收费"
                chk收费.Value = IIf(rsTemp("参数值") = 1, 1, 0)
            Case "支持特殊门诊"
                chk病种.Value = IIf(rsTemp("参数值") = 1, 1, 0)
            Case "入院时选择参保前在院"
                chk参保前在院.Value = IIf(rsTemp("参数值") = 1, 1, 0)
            Case "医保服务器"
                txtServer.Text = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "门诊最大项目数"
                txt项目数.Text = IIf(IsNull(rsTemp("参数值")), "7", rsTemp("参数值"))
        End Select
        
        rsTemp.MoveNext
    Loop
    
    '读取出已设置好的归属科目编码对应关系，仅所有都允许修改
    '发票归属科目编码
    '01：西药；02：中成药；03：中草药；04：床位费；05：诊查费；06：检查费；
    '07：治疗费；08：护理费；09：手术费；10：化验费；11：其他
    gstrSQL = "Select 编码,类别,'11-其他' 归属科目编码  " & _
             " From 收费类别 " & _
             " Where 编码 Not IN  " & _
             "     (Select 参数名 From 保险参数 Where 险类=" & lng险类 & " And 序号>=6) " & _
             " union   " & _
             " Select B.编码,B.类别,decode(A.参数值,'01','01-西药','02','02-中成药', " & _
             " '03','03-中草药','04','04-床位费','05','05-诊查费','06','06-检查费','07','07-治疗费', " & _
             " '08','08-护理费','09','09-手术费','10','10-化验费','11-其他') 归属科目编码   " & _
             " From 保险参数 A,收费类别 B " & _
             " Where A.序号>=6 And A.险类=" & lng险类 & " And A.参数名=B.编码"
    Call OpenRecordset(rsTemp, "获取收费类别")
    '初始化单据控件
    With Bill
        .Rows = 2
        .Cols = 3
        .TextMatrix(0, 0) = "编码"
        .TextMatrix(0, 1) = "收费类别"
        .TextMatrix(0, 2) = "归属科目编码"
        .ColWidth(0) = 500
        .ColWidth(1) = 1000
        .ColWidth(2) = 1800
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColData(0) = 0
        .ColData(2) = 3

        .AddItem "01-西药"
        .AddItem "02-中成药"
        .AddItem "03-中草药"
        .AddItem "04-床位费"
        .AddItem "05-诊查费"
        .AddItem "06-检查费"
        .AddItem "07-治疗费"
        .AddItem "08-护理费"
        .AddItem "09-手术费"
        .AddItem "10-化验费"
        .AddItem "11-其他"
        .ListIndex = 10
        
        .PrimaryCol = 0
        .LocateCol = 2
    End With
    
    With rsTemp
        Do While Not .EOF
            Bill.TextMatrix(.AbsolutePosition, 0) = !编码
            Bill.TextMatrix(.AbsolutePosition, 1) = !类别
            Bill.TextMatrix(.AbsolutePosition, 2) = !归属科目编码
            .MoveNext
            Bill.Rows = Bill.Rows + 1
        Loop
        If .RecordCount <> 0 Then Bill.Rows = Bill.Rows - 1
        Bill.Row = 1
    End With
    
    Bill.AllowAddRow = False
    Bill.Active = OwnerUser(gstrDbUser)
    
    mblnChange = False
    frmSet贵阳.Show vbModal, frm医保类别
    参数设置 = mblnOK
End Function

