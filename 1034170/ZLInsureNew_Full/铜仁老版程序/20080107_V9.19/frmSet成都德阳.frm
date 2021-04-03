VERSION 5.00
Begin VB.Form frmSet成都德阳 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chk拨号 
      Caption         =   "验号拨号连接(&L)"
      Height          =   270
      Left            =   1200
      TabIndex        =   8
      Top             =   3330
      Width           =   3120
   End
   Begin VB.ComboBox cbo社保机构 
      Height          =   300
      Left            =   1230
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2955
      Width           =   3315
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   1
      Left            =   -90
      TabIndex        =   16
      Top             =   3720
      Width           =   7665
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   0
      Left            =   0
      TabIndex        =   14
      Top             =   660
      Width           =   7665
   End
   Begin VB.Frame fra医保服务器 
      Caption         =   "医院前置医保服务器"
      Height          =   1980
      Left            =   435
      TabIndex        =   12
      Top             =   825
      Width           =   4155
      Begin VB.CommandButton cmdTest 
         Caption         =   "测试(&T)"
         Height          =   1095
         Left            =   2955
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   555
         Width           =   1005
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   5
         Top             =   1335
         Width           =   1635
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1200
         MaxLength       =   40
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   945
         Width           =   1635
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   1
         Top             =   555
         Width           =   1635
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "服务器(&S)"
         Height          =   180
         Index           =   2
         Left            =   330
         TabIndex        =   4
         Top             =   1395
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "密码(&P)"
         Height          =   180
         Index           =   1
         Left            =   510
         TabIndex        =   2
         Top             =   1005
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "用户名(&U)"
         Height          =   180
         Index           =   0
         Left            =   330
         TabIndex        =   0
         Top             =   615
         Width           =   810
      End
   End
   Begin VB.CommandButton cmd社保机构 
      Caption         =   "下载社保机构(&D)"
      Height          =   350
      Left            =   135
      TabIndex        =   11
      Top             =   3870
      Width           =   1845
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2385
      TabIndex        =   9
      Top             =   3870
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3570
      TabIndex        =   10
      Top             =   3870
      Width           =   1100
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   150
      Picture         =   "frmSet成都德阳.frx":0000
      Top             =   150
      Width           =   480
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "社保机构"
      Height          =   180
      Index           =   1
      Left            =   390
      TabIndex        =   6
      Top             =   3030
      Width           =   720
   End
   Begin VB.Label lbl 
      Caption         =   "配置相关的参数."
      Height          =   195
      Index           =   0
      Left            =   720
      TabIndex        =   15
      Top             =   360
      Width           =   7125
   End
End
Attribute VB_Name = "frmSet成都德阳"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mcnTest As New ADODB.Connection
Private mblnChange As Boolean
Private mblnFirst As Boolean
Private Enum enum文本
    Text医保用户 = 0
    Text医保密码 = 1
    Text医保服务器 = 2
End Enum



Public Function 参数设置() As Boolean
    mblnChange = False
    Dim rsTemp As New ADODB.Recordset
    frmSet成都德阳.Show vbModal, frm医保类别
    参数设置 = mblnOK
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cbo社保机构_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub chk拨号_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdTest_Click()
    Dim rsTemp As New ADODB.Recordset
    If mcnTest.State = adStateOpen Then mcnTest.Close
    
    If OraDataOpen(mcnTest, txtEdit(Text医保服务器).Text, txtEdit(Text医保用户).Text, txtEdit(Text医保密码).Tag) = False Then
        Exit Sub
    End If
    
    MsgBox "连接成功！", vbInformation, gstrSysName
End Sub
Private Sub cmd社保机构_Click()
    Dim strOutPut As String, strInPut As String
    Dim strArr, strArr1
    Dim i As Long
    If mcnTest Is Nothing Then
        MsgBox "请先测试中间库是否成常!"
        Exit Sub
    End If
    If mcnTest.State <> 1 Then
        MsgBox "请先测试中间库是否成常!"
        Exit Sub
    End If
    If 医保初始化_成都德阳 = False Then Exit Sub
    
    zlCommFun.ShowFlash "正在下载社保机构,请稍后..."
    If 业务请求_成都德阳(获得社保机构, strInPut, strOutPut) = False Then
        zlCommFun.StopFlash
        Exit Sub
    End If
    If strOutPut = "" Then
        zlCommFun.StopFlash
        Exit Sub
    End If
    strArr = Split(strOutPut, "@$")
    For i = 0 To UBound(strArr)
        strArr1 = Split(strArr(i), "||")
        '更新社保机构
        gstrSQL = "ZL_社保机构目录_UPDATE("
        gstrSQL = gstrSQL & "'" & strArr1(0) & "',"
        gstrSQL = gstrSQL & "'" & strArr1(1) & "')"
        gcnOracle_成都德阳.Execute gstrSQL, , adCmdStoredProc
    Next
    '重新加载数据
    Call LoadCbo
    zlCommFun.StopFlash
End Sub

Private Sub Form_Activate()
    Dim rsTemp As New ADODB.Recordset
    If mblnFirst = False Then Exit Sub
    
    mblnFirst = False
    Call LoadCbo
    gstrSQL = "Select * From 保险参数 where 险类=" & TYPE_成都德阳
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    With rsTemp
        Do While Not .EOF
            Select Case Nvl(!参数名)
            Case "医保用户名"
                  txtEdit(Text医保用户).Text = Nvl(!参数值)
            Case "医保用户密码"
                  txtEdit(Text医保密码).Text = Nvl(!参数值)
            Case "医保服务器"
                  txtEdit(Text医保服务器).Text = Nvl(!参数值)
            Case "检查拨号连接"
                    chk拨号.Value = IIf(Nvl(!参数值, 1) = 1, 1, 0)
            End Select
            .MoveNext
        Loop
    End With
 End Sub

Private Sub Form_Load()
    mblnFirst = True
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = Text医保密码 Then
        txtEdit(Index).Tag = txtEdit(Index).Text
    End If
    
    If Index = Text医保服务器 Or Index = Text医保密码 Or Index = Text医保用户 Then
        '关闭对医保服务器的连接，因为在参数设置完成时需要重新打开
        If mcnTest.State = adStateOpen Then mcnTest.Close
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub cmdOK_Click()
    If IsValid = False Then Exit Sub
    If SaveData = False Then Exit Sub
    
    mblnOK = True
    mblnChange = False
    Unload Me
End Sub

Private Function IsValid() As Boolean
    Dim lngCount As Long
    Dim strTitle As String
    Dim rsTemp As New ADODB.Recordset
    
    
    For lngCount = txtEdit.LBound To txtEdit.UBound
        If zlCommFun.StrIsValid(txtEdit(lngCount).Text, txtEdit(lngCount).MaxLength) = False Then
            zlControl.TxtSelAll txtEdit(lngCount)
            txtEdit(lngCount).SetFocus
            Exit Function
        End If
    Next
    
    If mcnTest.State = adStateClosed Then
        If OraDataOpen(mcnTest, txtEdit(Text医保服务器).Text, txtEdit(Text医保用户).Text, txtEdit(Text医保密码).Tag, False) = False Then
            If MsgBox("医保服务器不能正常连接，是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
        
    IsValid = True
End Function

Private Function SaveData() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim lst As ListItem
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    '删除已经数据
    gstrSQL = "zl_保险参数_Delete(" & TYPE_成都德阳 & ",null)"
    Call ExecuteProcedure(Me.Caption)
    
    '新增参数数据
    gstrSQL = "zl_保险参数_Insert(" & TYPE_成都德阳 & ",null,'医保用户名','" & txtEdit(Text医保用户).Text & "',1)"
    Call ExecuteProcedure(Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_成都德阳 & ",null,'医保用户密码','" & txtEdit(Text医保密码).Tag & "',2)"
    Call ExecuteProcedure(Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_成都德阳 & ",null,'医保服务器','" & txtEdit(Text医保服务器).Text & "',3)"
    Call ExecuteProcedure(Me.Caption)
    
    gstrSQL = "zl_保险参数_Insert(" & TYPE_成都德阳 & ",null,'检查拨号连接','" & IIf(chk拨号.Value = 1, 1, 0) & "',4)"
    Call ExecuteProcedure(Me.Caption)
    
    gcnOracle.CommitTrans
    If cbo社保机构.ListIndex >= 0 Then
        SaveRegInFor g公共模块, "医保", "社保机构代码", Split(cbo社保机构.Text, "--")(0)
    End If
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
End Function
Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub
Private Sub LoadCbo()
        '加载Grid数据
        Err = 0
        On Error GoTo ErrHand:
        Dim rsTemp As New ADODB.Recordset
        Dim i As Long
        gstrSQL = "Select * From 社保机构目录"
        zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取社保机构目录" '  rsTemp.Open gstrSQL, gcnOracle_成都德阳
        With rsTemp
            i = 1
            Me.cbo社保机构.Clear
            Do While Not .EOF
                cbo社保机构.AddItem Nvl(!编码) & "--" & Nvl(!名称)
                .MoveNext
            Loop
        End With
        SetDefaultSel
        Exit Sub
ErrHand:
        If ErrCenter = 1 Then Resume
End Sub
Private Function SetDefaultSel() As Boolean
    Dim strReg As String
    Dim i As Integer
    
    SetDefaultSel = False
    Err = 0: On Error GoTo ErrHand:
    Call GetRegInFor(g公共模块, "医保", "社保机构代码", strReg)
    If cbo社保机构.ListCount = 0 Then Exit Function
    For i = 0 To cbo社保机构.ListCount
        If Split(cbo社保机构.List(i), "--")(0) = strReg Then
            cbo社保机构.ListIndex = i
            Exit For
        End If
    Next
    If cbo社保机构.ListIndex < 0 Then
        cbo社保机构.ListIndex = 0
    End If
    SetDefaultSel = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
