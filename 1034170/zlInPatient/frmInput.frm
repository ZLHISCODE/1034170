VERSION 5.00
Begin VB.Form frmInput 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "中联软件"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5130
   ControlBox      =   0   'False
   Icon            =   "frmInput.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdSelectNO 
      Caption         =   "…"
      Height          =   300
      Left            =   4000
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "热键:F8 缺号选择"
      Top             =   795
      Width           =   330
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3480
      TabIndex        =   2
      Top             =   1530
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2220
      TabIndex        =   1
      Top             =   1530
      Width           =   1100
   End
   Begin VB.TextBox txtInput 
      Height          =   300
      Left            =   1980
      MaxLength       =   18
      TabIndex        =   0
      Top             =   795
      Width           =   2025
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   6000
      Y1              =   1335
      Y2              =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   6000
      Y1              =   1350
      Y2              =   1350
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   240
      Picture         =   "frmInput.frx":000C
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lblNote 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   0  'Transparent
      Caption         =   "在留观病人 钟无艳 转为住院病人之前，请先为该病人确定一个住院号。"
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   975
      TabIndex        =   4
      Top             =   165
      Width           =   3825
   End
   Begin VB.Label lblInput 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "住院号"
      Height          =   180
      Left            =   1380
      TabIndex        =   3
      Top             =   855
      Width           =   540
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mblnIme As Boolean
Private mbytType As Byte
Private mblnAllowNull As Boolean
Private mblnUcase As Boolean

Private mstrInput As String
Private mblnOk As Boolean
Private mblnPassInput As Boolean
Private mblnAffirmPass As Boolean
Private mobjKeyboard As Object

Public Function InputVal(ByVal frmParent As Object, ByVal strItem As String, _
    ByVal strNote As String, ByRef strInput As String, ByVal bytType As Byte, _
    Optional ByVal intMax As Integer, Optional ByVal blnAllowNull As Boolean, Optional ByVal blnAllowInput As Boolean = True, _
    Optional ByVal blnUCase As Boolean, Optional ByVal blnIme As Boolean, Optional blnPassInput As Boolean = False, _
    Optional blnAffirmPass As Boolean = False) As Boolean
'功能：显示一个输入框,类似VB的InputBox函数
'参数：frmParent=父窗体
'      strItem=要输入的项目名称
'      strNote=输入框中的提示。
'      strInput=入/出参数:初始显示及返回的值
'      bytType=数据类型:0-字符串,1-数字(住院号),2-日期
'      intMax=输入长度限制
'      blnAllowNull=是否允许输入空
'      blnAllowInput=是否允许输入
'      blnUCase=输入是否全部大写
'      blnIme=是否自动打开输入法
'      blnPassInput-是否密码输入
'      blnAffirmPass-是否输入的确认密码
'返回：输入确定返回True,取消返回Fasle
    mblnPassInput = blnPassInput: mblnAffirmPass = blnAffirmPass
    Load Me
    Me.Caption = gstrSysName
    Me.lblNote.Caption = strNote
    Me.lblInput.Caption = strItem
    Me.txtInput.Text = strInput
    Me.txtInput.MaxLength = intMax
    '87794
    Me.txtInput.Enabled = blnAllowInput
    If Me.txtInput.Enabled = True Then
        Me.txtInput.BackColor = &H80000005
    Else
        Me.txtInput.BackColor = &H80000004
    End If
    Me.cmdSelectNO.Visible = ((bytType = 1) And blnAllowInput)
    Me.cmdSelectNO.Enabled = ((bytType = 1) And blnAllowInput)
    
    mblnIme = blnIme
    mbytType = bytType
    mblnUcase = blnUCase
    mblnAllowNull = blnAllowNull
    
    Me.Show 1, frmParent
    
    strInput = mstrInput
    InputVal = mblnOk
End Function

Private Sub cmdCancel_Click()
    mstrInput = ""
    mblnOk = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strNO As String
    If txtInput.Text = "" And Not mblnAllowNull Then
        MsgBox "必须输入" & lblInput.Caption & "！", vbInformation, gstrSysName
        txtInput.SetFocus: Exit Sub
    End If
    If txtInput.MaxLength <> 0 Then
        If zlCommFun.ActualLen(txtInput.Text) > txtInput.MaxLength Then
            MsgBox "最多允许输入 " & txtInput.MaxLength & " 个字符或 " & txtInput.MaxLength \ 2 & " 个汉字！", vbInformation, gstrSysName
            txtInput.SetFocus: Exit Sub
        End If
    End If
    If mbytType = 1 Then
        If Not IsNumeric(txtInput.Text) Then
            MsgBox "输入内容不是合法的数字！", vbInformation, gstrSysName
            txtInput.SetFocus: Exit Sub
        End If
    ElseIf mbytType = 2 Then
        If Not IsNumeric(txtInput.Text) Then
            MsgBox "输入内容不是合法的日期！", vbInformation, gstrSysName
            txtInput.SetFocus: Exit Sub
        End If
    End If
    
    '目前此窗体还没有其它用途仅仅是留观病人转住院病人时使用，frmManageCourse，所以未区分功能
    If mbytType = 1 Then
        If ExistInPatiNO(txtInput.Text) Then
            strNo = zlDatabase.GetNextNo(2)
            If Val(txtInput.Text) = Val(strNo) Then
                MsgBox "当前住院号和自动获取的新住院号重复,请手工修改住院号！", vbInformation, gstrSysName
                txtInput.Enabled = True: Me.cmdSelectNO.Visible = True: Me.cmdSelectNO.Enabled = True
            Else
                MsgBox "当前住院号已被使用,将自动获取一个新的住院号,请确认！", vbInformation, gstrSysName
                txtInput.Text = Val(strNo)
            End If
            txtInput.SetFocus: Exit Sub
        End If
    End If
    
    mstrInput = txtInput.Text
    mblnOk = True
    Unload Me
End Sub

Private Sub cmdSelectNO_Click()
    Dim strNO As String
    
    Call frmNOSelect.ShowMe(Me, strNO)
    txtInput.Text = strNO
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF8
            If cmdSelectNO.Enabled And cmdSelectNO.Visible Then cmdSelectNO_Click
    End Select
End Sub

Private Sub Form_Load()
    If mblnPassInput Then CreateObjectKeyboard
End Sub

Private Sub txtInput_GotFocus()
    SelAll txtInput
    If mblnIme Then Call OpenIme(gstrIme)
    If Not mblnPassInput Then Exit Sub
    OpenPassKeyboard txtInput, mblnAffirmPass
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cmdOK_Click
    Else
        If mbytType = 1 Then '数字
            If InStr("0123456789" & Chr(27), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
        ElseIf mbytType = 2 Then '日期
            If InStr("0123456789:-" & Chr(27), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
        End If
        If mblnUcase Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtInput_LostFocus()
    If mblnIme Then Call OpenIme
    If Not mblnPassInput Then Exit Sub
    ClosePassKeyboard txtInput
End Sub

Private Function OpenPassKeyboard(ctlText As Control, Optional bln确认密码 As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开密码键盘输入
    '返回:打成成功,返回true,否者False
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
    '返回:打成成功,返回true,否者False
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

Private Function CreateObjectKeyboard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建密码创建
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

