VERSION 5.00
Begin VB.Form frmAgentInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "代办人信息"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3840
   Icon            =   "frmAgentInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancle 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2640
      TabIndex        =   10
      Top             =   2280
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1440
      TabIndex        =   9
      Top             =   2280
      Width           =   1100
   End
   Begin VB.Frame fraAgent 
      Caption         =   "代办人"
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   3615
      Begin VB.TextBox txtAgentName 
         Height          =   300
         Left            =   1080
         MaxLength       =   18
         TabIndex        =   1
         Top             =   240
         Width           =   2130
      End
      Begin VB.TextBox txtAgentIDNO 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1080
         MaxLength       =   18
         TabIndex        =   2
         Top             =   600
         Width           =   2130
      End
      Begin VB.Label lblAgentName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         Height          =   180
         Left            =   600
         TabIndex        =   8
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lblAgentIDNO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份证号"
         Height          =   180
         Left            =   240
         TabIndex        =   7
         Top             =   660
         Width           =   720
      End
   End
   Begin VB.Frame fraPati 
      Caption         =   "病人"
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3615
      Begin VB.TextBox txtPatiIDNO 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1080
         MaxLength       =   18
         TabIndex        =   0
         Top             =   600
         Width           =   2130
      End
      Begin VB.Label lblPatiName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名  李建飞"
         Height          =   180
         Left            =   600
         TabIndex        =   5
         Top             =   300
         Width           =   1080
      End
      Begin VB.Label lblPatiIDNO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份证号"
         Height          =   180
         Left            =   240
         TabIndex        =   4
         Top             =   660
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmAgentInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOK As Boolean
Private mlng病人ID As Long
Private mlng就诊ID As Long
Private mstr姓名 As String
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1


Public Function ShowMe(ByVal frmParent As Object, ByVal lng病人ID As Long, ByVal lng就诊ID As Long, ByVal str病人姓名 As String, _
                ByVal str病人身份证号 As String, ByVal str代办人姓名 As String, ByVal str代办人身份证号 As String) As Boolean
    Screen.MousePointer = 0
    mblnOK = False
    mlng病人ID = lng病人ID
    mlng就诊ID = lng就诊ID
    mstr姓名 = str病人姓名
    
    lblPatiName.Caption = "姓名  " & str病人姓名
    txtPatiIDNO.Text = str病人身份证号
    txtAgentName.Text = str代办人姓名
    txtAgentIDNO.Text = str代办人身份证号
    
    Me.Show 1, frmParent
    ShowMe = mblnOK
End Function

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHandle
    
    If Trim(txtPatiIDNO.Text) = "" Then
        MsgBox "请输入病人身份证号！", vbInformation, gstrSysName
        txtPatiIDNO.SetFocus: Exit Sub
    End If
    
    If Len(txtPatiIDNO.Text) <> 15 And Len(txtPatiIDNO.Text) <> 18 Then
        MsgBox "身份证号长度不正确，请输入15或18位身份证号！", vbInformation, gstrSysName
        txtPatiIDNO.SetFocus: Exit Sub
    End If
'    If Trim(txtAgentName.Text) = "" Then
'        MsgBox "请输入代办人姓名！", vbInformation, gstrSysName
'        txtAgentName.SetFocus: Exit Sub
'    End If
'
'    If Trim(txtAgentIDNO.Text) = "" Then
'        MsgBox "请输入代办人身份证号！", vbInformation, gstrSysName
'        txtAgentIDNO.SetFocus: Exit Sub
'    End If

    If txtAgentIDNO.Text <> "" And Len(txtAgentIDNO.Text) <> 15 And Len(txtAgentIDNO.Text) <> 18 Then
        MsgBox "身份证号长度不正确，请输入15或18位身份证号！", vbInformation, gstrSysName
        txtAgentIDNO.SetFocus: Exit Sub
    End If
    
    If Trim(txtAgentIDNO.Text) = Trim(txtPatiIDNO.Text) Then
        MsgBox "代办人身份证号与病人身份证号相同，请重新输入！", vbInformation, gstrSysName
        txtAgentIDNO.SetFocus: Exit Sub
    End If
    
    Screen.MousePointer = 11
    gstrSQL = "Zl_代办人信息_Insert(" & mlng病人ID & ",'" & Trim(txtPatiIDNO.Text) & "'," & IIF(Trim(txtAgentName.Text) = "", "Null", "'" & Trim(txtAgentName.Text) & "'") & ",'" & _
                Trim(txtAgentIDNO.Text) & "'," & mlng就诊ID & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Screen.MousePointer = 0
    mblnOK = True
    Unload Me
    Exit Sub
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If txtPatiIDNO.Text <> "" Then
        txtAgentName.SetFocus
    End If
End Sub

Private Sub Form_Load()
    If InStr(GetInsidePrivs(p门诊医生站), "代办人信息允许自由录入") = 0 Then
        txtPatiIDNO.Locked = True
        txtAgentName.Locked = True
        txtAgentIDNO.Locked = True
    End If
    Me.Caption = Me.Caption & "  (身份证刷卡录入)"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    Screen.MousePointer = 11
End Sub

Private Sub txtAgentIDNO_Change()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (Me.ActiveControl Is txtAgentIDNO)
End Sub


Private Sub txtAgentIDNO_GotFocus()
    zlControl.TxtSelAll txtAgentIDNO
    Call OpenIDCard(True)
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (True)
End Sub

Private Sub txtAgentIDNO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf InStr("0123456789X" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtAgentIDNO_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
    txtAgentIDNO.Text = Trim(txtAgentIDNO.Text)
End Sub

Private Sub txtAgentName_Change()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (Me.ActiveControl Is txtAgentName)
End Sub


Private Sub txtAgentName_GotFocus()
    zlControl.TxtSelAll txtAgentName
    Call OpenIDCard(True)
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (True)
End Sub


Private Sub txtAgentName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtAgentName_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
    txtAgentName.Text = Trim(txtAgentName.Text)
End Sub


Private Sub txtPatiIDNO_Change()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (Me.ActiveControl Is txtPatiIDNO)
End Sub


Private Sub txtPatiIDNO_GotFocus()
    zlControl.TxtSelAll txtPatiIDNO
    Call OpenIDCard(True)
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (True)
End Sub


Private Sub txtPatiIDNO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf InStr("0123456789X" & Chr(8), UCase(Chr(KeyAscii))) = 0 Then
        KeyAscii = 0
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtPatiIDNO_LostFocus()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (False)
    txtPatiIDNO.Text = Trim(txtPatiIDNO.Text)
End Sub



Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthDay As Date, ByVal strAddress As String)
On Error GoTo errH
    If Me.ActiveControl Is txtPatiIDNO Then
        If mstr姓名 = strName Then
            txtPatiIDNO.Text = strID
        Else
            MsgBox "身份信息录入失败,请使用当前病人的身份证刷卡。", vbInformation, gstrSysName
        End If
    ElseIf Me.ActiveControl Is txtAgentName Or Me.ActiveControl Is txtAgentIDNO Then
        txtAgentName.Text = strName
        txtAgentIDNO.Text = strID
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetOldAcademic(ByVal DateBir As Date, ByVal str年龄单位 As String) As Long
'功能：根据当前的出生日期和年龄单位，计算理论上的年龄值
'返回：年龄
    Dim datCur As Date, lngOld As Long, strInterval As String
    If DateBir = CDate(0) Or InStr(" 岁月天", str年龄单位) < 2 Then Exit Function
    
    datCur = zlDatabase.Currentdate
    
    strInterval = Switch(str年龄单位 = "岁", "yyyy", str年龄单位 = "月", "m", str年龄单位 = "天", "d")
    lngOld = DateDiff(strInterval, DateBir, datCur)
    If DateAdd(strInterval, lngOld, DateBir) > datCur Then
        lngOld = lngOld - 1
    End If
    GetOldAcademic = lngOld
End Function


Private Sub OpenIDCard(ByVal blnEnabled As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打开身份证读卡器
    '编制:王吉
    '日期:2012-08-31 16:28:23
    '问题号:53408
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '初始化对卡对象
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.hWnd)
    End If
    '打开读卡器
    mobjIDCard.SetEnabled (blnEnabled)
End Sub

