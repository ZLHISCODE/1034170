VERSION 5.00
Begin VB.Form frmSet北京尚洋 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "运行参数设置"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtNumber 
      Height          =   300
      Left            =   1380
      TabIndex        =   5
      Top             =   2175
      Width           =   3345
   End
   Begin VB.Frame fra医保服务器 
      Caption         =   "医院前置医保服务器"
      Height          =   1545
      Left            =   60
      TabIndex        =   8
      Top             =   90
      Width           =   4695
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1215
         MaxLength       =   40
         TabIndex        =   2
         Top             =   1110
         Width           =   2145
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1215
         MaxLength       =   40
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   720
         Width           =   2145
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1215
         MaxLength       =   40
         TabIndex        =   0
         Top             =   330
         Width           =   2145
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "测试(&T)"
         Height          =   1110
         Left            =   3450
         TabIndex        =   3
         Top             =   330
         Width           =   1110
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "数据源(&S)"
         Height          =   180
         Index           =   2
         Left            =   300
         TabIndex        =   11
         Top             =   1170
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "密码(&P)"
         Height          =   180
         Index           =   1
         Left            =   480
         TabIndex        =   10
         Top             =   780
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "用户名(&U)"
         Height          =   180
         Index           =   0
         Left            =   270
         TabIndex        =   9
         Top             =   390
         Width           =   840
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4890
      TabIndex        =   7
      Top             =   765
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4890
      TabIndex        =   6
      Top             =   285
      Width           =   1100
   End
   Begin VB.ComboBox cbo适用地区 
      Height          =   300
      Left            =   1380
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1755
      Width           =   3345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "统筹区号(&N)"
      Height          =   180
      Left            =   255
      TabIndex        =   13
      Top             =   2250
      Width           =   990
   End
   Begin VB.Label lbl适用地区 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "适用地区(&Q)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   255
      TabIndex        =   12
      Top             =   1815
      Width           =   990
   End
End
Attribute VB_Name = "frmSet北京尚洋"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mblnChange As Boolean
Private mblnChangePassword As Boolean  '密码被修改过
 
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If IsValid = False Then Exit Sub
    If SaveData = False Then Exit Sub
    
    mblnOK = True
    mblnChange = False
    Unload Me
End Sub

Private Sub cmdTest_Click()
    If gcn尚洋.State = adStateOpen Then gcn尚洋.Close
    On Error Resume Next
    If cbo适用地区.ListIndex = 1 Then
'        gcn尚洋.ConnectionString = "Provider=Sybase.ASEOLEDBProvider.2;口令=" & Trim(txtEdit(1).Tag) & ";持续安全性信息=True;用户 ID=" & Trim(txtEdit(0).Text) & ";数据源=" & Trim(txtEdit(2).Text)
        gcn尚洋.ConnectionString = "Provider=MSDASQL.1;Password=" & Trim(txtEdit(1).Tag) & ";Persist Security Info=True;User ID=" & Trim(txtEdit(0).Text) & ";Data Source=" & Trim(txtEdit(2).Text)
    Else
        gcn尚洋.ConnectionString = "Provider=MSDAORA.1;Password=" & Trim(txtEdit(1).Tag) & ";User ID=" & Trim(txtEdit(0).Text) & ";Data Source=" & Trim(txtEdit(2).Text) & ";Persist Security Info=True"
    End If
    gcn尚洋.CursorLocation = adUseClient
    gcn尚洋.Open
    If Err <> 0 Then
        MsgBox "医保前置服务器连接失败！", vbInformation, gstrSysName
        Exit Sub
    End If
    MsgBox "医保前置服务器连接成功", vbInformation, gstrSysName
End Sub

Private Function IsValid() As Boolean
    Dim lngCount As Long
    Dim strTitle As String
    
    '逐步判断字符的合法性
    For lngCount = txtEdit.LBound To txtEdit.UBound
        If zlCommFun.StrIsValid(txtEdit(lngCount).Text, txtEdit(lngCount).MaxLength) = False Then
            zlControl.TxtSelAll txtEdit(lngCount)
            txtEdit(lngCount).SetFocus
            Exit Function
        End If
    Next
    
    '对连接进行测试
    If gcn尚洋.State = adStateClosed Then
        On Error Resume Next
        If gcn尚洋.State = adStateOpen Then gcn尚洋.Close
        If cbo适用地区.ListIndex = 1 Then
'            gcn尚洋.ConnectionString = "Provider=Sybase.ASEOLEDBProvider.2;口令=" & Trim(txtEdit(1).Tag) & ";持续安全性信息=True;用户 ID=" & Trim(txtEdit(0).Text) & ";数据源=" & Trim(txtEdit(2).Text)
            gcn尚洋.ConnectionString = "Provider=MSDASQL.1;Password=" & Trim(txtEdit(1).Tag) & ";Persist Security Info=True;User ID=" & Trim(txtEdit(0).Text) & ";Data Source=" & Trim(txtEdit(2).Text)
        Else
            gcn尚洋.ConnectionString = "Provider=MSDAORA.1;Password=" & Trim(txtEdit(1).Tag) & ";User ID=" & Trim(txtEdit(0).Text) & ";Data Source=" & Trim(txtEdit(2).Text) & ";Persist Security Info=True"
        End If
        gcn尚洋.CursorLocation = adUseClient
        gcn尚洋.Open
        
        If Err <> 0 Then
            If MsgBox("医保服务器不能正常连接，是否继续？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    On Error Resume Next
    IsValid = True
End Function

Public Function 参数设置() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim str参数值 As String
    Dim int适用地区 As Integer
    
    mblnOK = False
    On Error GoTo errHandle
    
    
    gstrSQL = "select 参数名,参数值 from 保险参数 where 险类=" & TYPE_北京尚洋
    Call OpenRecordset(rsTemp, Me.Caption)
    
    int适用地区 = 0
    Do Until rsTemp.EOF
        str参数值 = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
        Select Case rsTemp("参数名")
            Case "用户名"
                txtEdit(0).Text = str参数值
            Case "服务器"
                txtEdit(2).Text = str参数值
            Case "用户密码"
                txtEdit(1).Text = "        "    '假密码
                txtEdit(1).Tag = str参数值
            Case "适用地区"
                int适用地区 = Val(str参数值)
            Case "统筹区号"
                txtNumber = str参数值
        End Select
        rsTemp.MoveNext
    Loop
'    If txtEdit(4).Text = "" Then txtEdit(4).Enabled = True
    On Error Resume Next
    With cbo适用地区
        .Clear
        .AddItem "测试使用(ORACLE环境)"
        .AddItem "慧丰职工医院(SYBASE环境)"
        .ListIndex = int适用地区
    End With
    
    mblnChange = False
    mblnChangePassword = False
    frmSet北京尚洋.Show vbModal, frm医保类别
    参数设置 = mblnOK
    Exit Function

errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function SaveData() As Boolean
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    '删除已经数据
    gstrSQL = "zl_保险参数_Delete(" & TYPE_北京尚洋 & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '新增参数数据
    gstrSQL = "zl_保险参数_Insert(" & TYPE_北京尚洋 & ",null,'用户名','" & txtEdit(0).Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_北京尚洋 & ",null,'用户密码','" & txtEdit(1).Tag & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_北京尚洋 & ",null,'服务器','" & txtEdit(2).Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_北京尚洋 & ",null,'适用地区','" & cbo适用地区.ListIndex & "',4)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_insert(" & TYPE_北京尚洋 & ",null,'统筹区号','" & txtNumber.Text & "',5)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    gcnOracle.CommitTrans
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
End Function

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = 1 Then
        txtEdit(1).Tag = txtEdit(1).Text
        mblnChangePassword = True
    End If
    
    '关闭对医保服务器的连接，因为在参数设置完成时需要重新打开
    If gcn尚洋.State = adStateOpen Then gcn尚洋.Close
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub
