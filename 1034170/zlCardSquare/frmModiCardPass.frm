VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmModiCardPass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "密码修改"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7425
   Icon            =   "frmModiCardPass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6120
      TabIndex        =   5
      Top             =   360
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6120
      TabIndex        =   6
      Top             =   870
      Width           =   1200
   End
   Begin VB.PictureBox picPass 
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   285
      ScaleHeight     =   3735
      ScaleWidth      =   5625
      TabIndex        =   0
      Top             =   285
      Width           =   5625
      Begin VB.TextBox txtAudi 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         IMEMode         =   3  'DISABLE
         Left            =   1095
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   3030
         Width           =   4245
      End
      Begin VB.TextBox txtPass 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         IMEMode         =   3  'DISABLE
         Left            =   1095
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   2520
         Width           =   4245
      End
      Begin VB.TextBox txtOldPass 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         IMEMode         =   3  'DISABLE
         Left            =   1095
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   2010
         Width           =   4245
      End
      Begin VB.TextBox txt卡号 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         IMEMode         =   3  'DISABLE
         Left            =   1095
         PasswordChar    =   "*"
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1470
         Width           =   4245
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   120
         X2              =   5520
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label lblNotes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请将[XX]从刷卡器上划过，  然后连续两次输入相同的密码！"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         TabIndex        =   11
         Top             =   180
         Width           =   8550
      End
      Begin VB.Image imgFlag 
         Height          =   720
         Left            =   120
         Picture         =   "frmModiCardPass.frx":058A
         Top             =   180
         Width           =   720
      End
      Begin VB.Label LabelVeriPass 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "验证"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   390
         TabIndex        =   10
         Top             =   3120
         Width           =   630
      End
      Begin VB.Label lblNewPass 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "新密码"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   75
         TabIndex        =   9
         Top             =   2580
         Width           =   945
      End
      Begin VB.Label lblOldPass 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "原密码"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   75
         TabIndex        =   8
         Top             =   2040
         Width           =   945
      End
      Begin VB.Label lbl卡号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "卡号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   390
         TabIndex        =   7
         Top             =   1500
         Width           =   630
      End
   End
   Begin XtremeSuiteControls.TaskPanel wndTaskPanel 
      Height          =   4065
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   5955
      _Version        =   589884
      _ExtentX        =   10504
      _ExtentY        =   7170
      _StockProps     =   64
      VisualTheme     =   6
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
End
Attribute VB_Name = "frmModiCardPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngModule As Long, mlngCardTypeID As Long
Private mblnOk As Boolean, mblnCheckOldPass As Boolean
Private mobjKeyboard As Object, mblnTest As Boolean
Private mblnFirst As Boolean
Private WithEvents mobjCommEvents As zl9CommEvents.clsCommEvents
Attribute mobjCommEvents.VB_VarHelpID = -1
Private mobjSquare As Object
Private mrsInfo As New ADODB.Recordset, mrsCardType As New ADODB.Recordset

Public Function zlModifyPass(ByVal frmMain As Object, _
    ByVal lngModule As Long, ByVal lngCardTypeID As Long, _
    Optional blnCheckOldPass As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调整密码入口参数
    '入参:frmMain-调用的主窗体
    '     lngModule -模块号
    '     lngCardTypeID-消费卡接口编号
    '返回:修改成功,返回true,否则返回false
    '编制:刘尔旋
    '日期:2013-10-21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule
    mblnOk = False: mlngCardTypeID = lngCardTypeID
    mblnCheckOldPass = blnCheckOldPass
    mblnTest = Val(GetSetting("ZLSOFT", "公共全局\zlSquareCard", "TestCardNO", 0)) = 1
    mblnTest = IsDesinMode Or mblnTest
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
    zlModifyPass = mblnOk
End Function

Private Function InitCardInfor() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化卡片信息
    '返回:初始化成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-29 14:25:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Set mrsCardType = zlGet消费卡接口
    mrsCardType.Filter = "编号=" & mlngCardTypeID
    If mrsCardType Is Nothing Then Exit Function
    
    lbl卡号.BorderStyle = Val(Nvl(mrsCardType!是否刷卡)): lbl卡号.Tag = Nvl(mrsCardType!是否刷卡)
    
    If Val(Nvl(mrsCardType!密码长度)) <> 0 Then
        txtOldPass.MaxLength = Val(Nvl(mrsCardType!密码长度))
        txtPass.MaxLength = Val(Nvl(mrsCardType!密码长度))
        txtAudi.MaxLength = Val(Nvl(mrsCardType!密码长度))
    Else
        txtOldPass.MaxLength = 10
        txtPass.MaxLength = 10
        txtAudi.MaxLength = 10
    End If
    lblNotes.Caption = Replace(lblNotes.Caption, "[XX]", "[" & Nvl(mrsCardType!名称) & "]")
    InitCardInfor = True
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Private Function CreateObjectKeyboard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建密码键盘
    '返回:创建成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-24 23:59:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    Set mobjKeyboard = CreateObject("zl9Keyboard.clsKeyboard")
    If Err <> 0 Then Exit Function
    Err = 0
    CreateObjectKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function OpenPassKeyboard(ctlText As Control, Optional bln确认密码 As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开密码键盘输入
    '返回:打成成功,返回true,否则False
    '编制:刘兴洪
    '日期:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.OpenPassKeyoardInput(Me, ctlText, bln确认密码) = False Then Exit Function
    OpenPassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Function ClosePassKeyboard(ctlText As Control) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开密码键盘输入
    '返回:打成成功,返回true,否则False
    '编制:刘兴洪
    '日期:2011-07-25 00:04:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mobjKeyboard Is Nothing Then Exit Function
    If mobjKeyboard.ColsePassKeyoardInput(Me, ctlText) = False Then Exit Function
    ClosePassKeyboard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查输入的数据是否有效
    '返回:数据有效,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-29 11:15:42
    '---------------------------------------------------------------------------------------------------------------------------------------------\
    On Error GoTo errHandle
    If CheckCard(mlngCardTypeID, txt卡号.Text) = False Then Exit Function
    If mrsInfo Is Nothing Then
        MsgBox "不能读取卡信息，请确定是否正确刷卡！", vbInformation, gstrSysName
        Call ClearFace: txt卡号.SetFocus: Exit Function
        Exit Function
    End If
    If mrsInfo.State <> 1 Then
        MsgBox "不能读取卡信息，请确定是否正确刷卡！", vbInformation, gstrSysName
        Call ClearFace: txt卡号.SetFocus: Exit Function
    End If
    If txtPass.Text <> txtAudi.Text Then
        MsgBox "两次输入的密码不一致，请重新输入！", vbInformation, gstrSysName
        txtPass.Text = "": txtAudi.Text = ""
        If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
        Exit Function
    End If
    If txtPass.Text = "" Then
        If MsgBox("当前设置的密码为空，确实要这样设置吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            If txtPass.Enabled And txtPass.Visible Then txtPass.SetFocus
            Exit Function
        End If
    End If
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function ModifCardPass() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:修改卡片的密码
    '返回:修改成功,返回true,否则返回False
    '编制:刘尔旋
    '日期:2013-10-21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strPassWord As String, intForce As Integer
    On Error GoTo errHandle
    If mrsInfo Is Nothing Then Exit Function
    strPassWord = zlCommFun.zlStringEncode(txtOldPass.Text)     '密码加密
    
    If strPassWord <> Nvl(mrsInfo!密码) And mblnCheckOldPass = True Then
        MsgBox "卡片原密码输入错误,请重新输入密码!", vbInformation, gstrSysName
        txtOldPass.SetFocus
        ModifCardPass = False
        Exit Function
    End If
    
    If mblnCheckOldPass = True Then
        intForce = 0
    Else
        intForce = 1
    End If
    
     'Zl_消费卡密码_Update
    strSQL = "Zl_消费卡密码_Update('" & Nvl(mrsInfo!卡号) & "'," & mlngCardTypeID & "," & _
             Val(Nvl(mrsInfo!序号)) & ",'" & strPassWord & "','" & zlCommFun.zlStringEncode(txtPass.Text) & "'," & intForce & ")"
    
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    ModifCardPass = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub cmdOK_Click()
    If isValied = False Then Exit Sub
    If ModifCardPass = False Then Exit Sub
    MsgBox "密码修改成功!", vbOKOnly + vbInformation, gstrSysName
    mblnOk = True
    Unload Me
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If InitCardInfor = False Then Unload Me: Exit Sub
    Call ClearFace
    Call txt卡号_Change
    txt卡号.SetFocus
End Sub

Private Sub Form_Load()
    mblnFirst = True
    Set mrsInfo = Nothing
    Call CreateObjectKeyboard

    Set mobjCommEvents = New zl9CommEvents.clsCommEvents
    
    If mblnCheckOldPass = False Then
        lblNotes.Top = 180
        lblNotes.Caption = "请将[XX]从刷卡器上划过后，" & vbCrLf & "连续两次输入相同的新密码！"
        txtOldPass.Enabled = False
        txtOldPass.BackColor = &H8000000F
    Else
        lblNotes.Top = 180
        lblNotes.Caption = "请将[XX]从刷卡器上划过后，" & vbCrLf & "输入旧密码与两次相同的新密码！"
    End If
    HookDefend txtOldPass.hWnd
    HookDefend txtPass.hWnd
    HookDefend txtAudi.hWnd
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
    txt卡号.Text = ""
    Set mobjCommEvents = Nothing
End Sub

Private Sub mobjCommEvents_ShowCardInfor(ByVal strCardType As String, ByVal strCardNo As String, ByVal strXmlCardInfor As String, strExpended As String, blnCancel As Boolean)
    If strCardNo = "" Then Exit Sub
    If Not GetCardPass(strCardType, strCardNo) Then
        Call ClearFace: If txt卡号.Enabled And txt卡号.Visible Then txt卡号.SetFocus
        Exit Sub
    End If
End Sub

Private Sub txtAudi_GotFocus()
    zlControl.TxtSelAll txtAudi
    OpenPassKeyboard txtAudi, True
End Sub

Private Sub txtAudi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: cmdOK_Click
    Else
        If InStr("';" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtAudi_LostFocus()
   ClosePassKeyboard txtPass
End Sub

Private Sub txtOldPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtPass_GotFocus()
    zlControl.TxtSelAll txtPass
    OpenPassKeyboard txtPass, False
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtPass.Text = "" And txtAudi.Text = "" Then
            cmdOk.SetFocus
        Else
            txtAudi.SetFocus
        End If
    Else
        If InStr("';" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub
Private Sub txtPass_LostFocus()
    ClosePassKeyboard txtPass
End Sub

Private Sub txtOldPass_GotFocus()
    zlControl.TxtSelAll txtOldPass
End Sub

Private Function CheckCard(ByVal lngCardTypeID As Long, ByVal strCardNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存前检查卡号有效性
    '入参:lngCardTypeID-消费卡接口编号
    '     strCardNO-卡号
    '返回:成功返回True,失败返回False
    '编制:刘尔旋
    '日期:2013-10-21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    Set mrsInfo = Nothing
    On Error GoTo errH
    
    strSQL = "" & _
    "   Select a.Id,a.卡类型,a.卡号,a.序号,a.可否充值,a.接口编号,to_char(a.有效期,'yyyy-mm-dd hh24:mi:ss') as 有效期,  a.密码," & _
    "          to_char(a.回收时间,'yyyy-mm-dd hh24:mi:ss') as 回收时间 , " & _
    "          decode(a.当前状态,2,'回收',3,'退卡','回收') as 当前状态, " & _
    "          to_char(a.停用日期,'yyyy-mm-dd hh24:mi:ss') as 停用日期," & _
    "          a.限制类别 " & _
    "   From 消费卡目录 A  " & _
    "   Where A.卡号 = [1] and A.接口编号=[2] And 序号 = (Select Max(序号) From 消费卡目录 B Where 卡号 = A.卡号 and 接口编号=A.接口编号)  " & _
    "   Order by a.序号"
    
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strCardNo, lngCardTypeID)
    If mrsInfo.EOF Then
        ShowMsgbox "未找到相关的" & Nvl(mrsCardType!名称) & "信息,请检查!"
        Exit Function
    End If
    
    '检查当前刷卡的合法性
    '是否回收
    If Nvl(mrsInfo!回收时间, "3000-01-01 00:00:00") < "3000-01-01 00:00:00" Then
        ShowMsgbox "卡号为" & strCardNo & "的" & Nvl(mrsCardType!名称) & "已经被" & Nvl(mrsInfo!当前状态) & ",不能再刷卡"
        Exit Function
    End If
    
    '是否停用
    If Nvl(mrsInfo!停用日期, "3000-01-01 00:00:00") < "3000-01-01 00:00:00" Then
        ShowMsgbox "卡号为" & strCardNo & "的" & Nvl(mrsCardType!名称) & "已经被停止使用,不能再刷卡"
        Exit Function
    End If
    '是否停用
    If Nvl(mrsInfo!停用日期, "3000-01-01 00:00:00") < "3000-01-01 00:00:00" Then
        ShowMsgbox "卡号为" & strCardNo & "的" & Nvl(mrsCardType!名称) & "已经被停止使用,不能再刷卡"
        Exit Function
    End If
    
    CheckCard = True
    Exit Function
errH:
    CheckCard = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Set mrsInfo = Nothing
End Function

Private Function GetCardPass(ByVal lngCardTypeID As Long, ByVal strCardNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取卡密码
    '入参:lngCardTypeID-消费卡接口编号
    '     strCardNO-卡号
    '返回:成功返回True,失败返回False
    '编制:刘尔旋
    '日期:2013-10-21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    Set mrsInfo = Nothing
    On Error GoTo errH
    
    txtPass.Text = "": txtAudi.Text = "": txtOldPass.Text = ""
    
    strSQL = "" & _
    "   Select a.Id,a.卡类型,a.卡号,a.序号,a.可否充值,a.接口编号,to_char(a.有效期,'yyyy-mm-dd hh24:mi:ss') as 有效期,  a.密码," & _
    "          to_char(a.回收时间,'yyyy-mm-dd hh24:mi:ss') as 回收时间 , " & _
    "          decode(a.当前状态,2,'回收',3,'退卡','回收') as 当前状态, " & _
    "          to_char(a.停用日期,'yyyy-mm-dd hh24:mi:ss') as 停用日期," & _
    "          a.限制类别 " & _
    "   From 消费卡目录 A  " & _
    "   Where A.卡号 = [1] and A.接口编号=[2] And 序号 = (Select Max(序号) From 消费卡目录 B Where 卡号 = A.卡号 and 接口编号=A.接口编号)  " & _
    "   Order by a.序号"
    
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strCardNo, lngCardTypeID)
    If mrsInfo.EOF Then
        ShowMsgbox "未找到相关的" & Nvl(mrsCardType!名称) & "信息,请检查!"
        Exit Function
    End If
    
    '检查当前刷卡的合法性
    '是否回收
    If Nvl(mrsInfo!回收时间, "3000-01-01 00:00:00") < "3000-01-01 00:00:00" Then
        ShowMsgbox "卡号为" & strCardNo & "的" & Nvl(mrsCardType!名称) & "已经被" & Nvl(mrsInfo!当前状态) & ",不能再刷卡"
        Exit Function
    End If
    
    '是否停用
    If Nvl(mrsInfo!停用日期, "3000-01-01 00:00:00") < "3000-01-01 00:00:00" Then
        ShowMsgbox "卡号为" & strCardNo & "的" & Nvl(mrsCardType!名称) & "已经被停止使用,不能再刷卡"
        Exit Function
    End If
    '是否停用
    If Nvl(mrsInfo!停用日期, "3000-01-01 00:00:00") < "3000-01-01 00:00:00" Then
        ShowMsgbox "卡号为" & strCardNo & "的" & Nvl(mrsCardType!名称) & "已经被停止使用,不能再刷卡"
        Exit Function
    End If
    
    GetCardPass = True
    Exit Function
errH:
    GetCardPass = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Set mrsInfo = Nothing
    Exit Function
End Function

Private Sub ClearFace()
    txt卡号.PasswordChar = IIf(Val(Nvl(mrsCardType!是否密文)) <> 0, "*", "")
    txt卡号.Text = ""
    txtPass.Text = "": txtAudi.Text = ""
End Sub

Private Sub txt卡号_Change()
    If mblnCheckOldPass = True Then txtOldPass.Enabled = Trim(txt卡号.Text) <> ""
    txtPass.Enabled = Trim(txt卡号.Text) <> ""
    txtAudi.Enabled = Trim(txt卡号.Text) <> ""
End Sub

Private Sub txt卡号_GotFocus()
    zlControl.TxtSelAll txt卡号
    txt卡号.PasswordChar = IIf(Val(Nvl(mrsCardType!是否密文)) <> 0, "*", "")

    If mobjSquare Is Nothing Then Set mobjSquare = CreateObject("zl9CardSquare.clsCardSquare")
    '初始化射频卡对象
    mobjSquare.zlInitEvents Me.hWnd, mobjCommEvents
    mobjSquare.SetEnabled True
End Sub

Private Sub txt卡号_KeyPress(KeyAscii As Integer)
    Static sngBegin As Single
    Dim sngNow As Single
    Dim blnCard As Boolean

    '是否刷卡完成
    blnCard = KeyAscii <> 8 And Len(txt卡号.Text) = Val(Nvl(mrsCardType!卡号长度)) - 1 And txt卡号.SelLength <> Len(txt卡号.Text)
    If blnCard Or KeyAscii = 13 Then
        If KeyAscii <> 13 Then
            txt卡号.Text = txt卡号.Text & Chr(KeyAscii)
            txt卡号.SelStart = Len(txt卡号.Text)
        End If
        KeyAscii = 0
        If GetCardPass(mlngCardTypeID, Trim(txt卡号.Text)) = False Then
            If txt卡号.Enabled Then txt卡号.SetFocus
            zlControl.TxtSelAll txt卡号
            Exit Sub
        End If
        If mblnCheckOldPass Then
            If txtOldPass.Enabled Then txtOldPass.SetFocus
        Else
            If txtPass.Enabled Then txtPass.SetFocus: Exit Sub
        End If
    Else
        If InStr(":：;；?？" & Chr(22), Chr(KeyAscii)) > 0 Then
            KeyAscii = 0 '去除特殊符号，并且不允许粘贴
        Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        End If
        
        If mblnTest Then Exit Sub
        '安全刷卡检测
        If KeyAscii <> 0 And KeyAscii > 32 Then
            sngNow = timer
            If txt卡号.Text = "" Then
                sngBegin = sngNow
            ElseIf Format((sngNow - sngBegin) / (Len(txt卡号.Text) + 1), "0.000") >= 0.04 Then '>0.007>=0.01
                txt卡号.Text = Chr(KeyAscii)
                txt卡号.SelStart = 1
                KeyAscii = 0
                sngBegin = sngNow
            End If
        End If
    End If
End Sub

Private Sub txt卡号_LostFocus()
   If Not mobjSquare Is Nothing Then mobjSquare.SetEnabled False
End Sub

