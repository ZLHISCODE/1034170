VERSION 5.00
Begin VB.Form frmLock 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "解除窗口锁定"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4230
   Icon            =   "frmLock.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdShutDown 
      Caption         =   "关闭导航台(&C)"
      Height          =   350
      Left            =   960
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.TextBox txtPwd 
      Height          =   350
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   870
      Width           =   2940
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "解锁(&U)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2800
      TabIndex        =   1
      Top             =   1320
      Width           =   1100
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      Caption         =   "当前操作员："
      Height          =   180
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   1080
   End
   Begin VB.Image imgFlag 
      Height          =   720
      Left            =   120
      Picture         =   "frmLock.frx":6852
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "请输入登录密码解锁"
      Height          =   180
      Left            =   960
      TabIndex        =   2
      Top             =   600
      Width           =   1620
   End
End
Attribute VB_Name = "frmLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Const MF_BYPOSITION = &H400&
Private Const MF_REMOVE = &H1000&

Private Sub cmdOK_Click()
    Dim strError As String, strPassWord As String
    If txtPwd.Text = "" Then
        MsgBox "请输入登录密码！", vbInformation, gstrSysName
        Exit Sub
    End If
    On Error Resume Next
    strPassWord = UCase(Trim(txtPwd.Text))
    strPassWord = IIf(gobjRelogin.IsTransPwd, TranZLHISPasswd(strPassWord), strPassWord)
    If LoginValidate(gobjRelogin.ServerName, gobjRelogin.InputUser, strPassWord, strError) Then
        '隐藏界面
        Call LockProg(False)
        Unload Me
    Else
        MsgBox "解锁失败！信息：" & strError, vbInformation, gstrSysName
        txtPwd.Text = ""
        Call txtPwd.SetFocus
    End If
End Sub

Private Sub cmdShutDown_Click()
    Call LockProg(False)
    Unload Me
    Unload frmBrower
End Sub

Private Sub Form_Activate()
'    Call SetActiveWindow(Me.hwnd)
    Call txtPwd.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Me.Caption = gstrUserName & "-解除锁定"
    lblUser.Caption = "当前操作员：" & gstrUserName
    
    If gblnShutDown Then
        Me.Width = 4320
        txtPwd.Width = 2940
        cmdOK.Left = 2800
        cmdShutDown.Visible = True
    Else
        Me.Width = 3816
        txtPwd.Width = 2412
        cmdOK.Left = 2280
        cmdShutDown.Visible = False
    End If
    Call DisableX
    Call SetActiveWindow(Me.hwnd)
'    Call txtPwd.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = gblnLock
End Sub

Private Sub txtPwd_GotFocus()
    zlControl.TxtSelAll txtPwd
End Sub

Private Sub DisableX()
     Dim hMenu As Long
     Dim nCount As Long
     hMenu = GetSystemMenu(Me.hwnd, 0)
     nCount = GetMenuItemCount(hMenu)
     Call RemoveMenu(hMenu, nCount - 1, MF_REMOVE Or MF_BYPOSITION)
     Call RemoveMenu(hMenu, nCount - 2, MF_REMOVE Or MF_BYPOSITION)
End Sub

Private Function LoginValidate(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String, Optional ByRef strError As String) As Boolean
    '------------------------------------------------
    '功能： 验证用户能否登录数据库
    '参数：
    '   strServerName：主机字符串
    '   strUserName：用户名
    '   strUserPwd：密码
    '返回： 数据库打开成功，返回true；失败，返回false
    '------------------------------------------------
    Dim strSQL As String
    Dim cnOracle As New ADODB.Connection
    
    On Error Resume Next
    Err = 0
    DoEvents
    With cnOracle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
        If Err <> 0 Then
            strError = Err.Description
            Err.Clear
            LoginValidate = False
            Exit Function
        End If
        .Close
    End With
    LoginValidate = True
    Exit Function
    
ErrHand:
    strError = Err.Description
    LoginValidate = False
    Err.Clear
End Function

Public Function TranZLHISPasswd(strOld As String) As String
    '------------------------------------------------
    '功能： 密码转换函数
    '参数：
    '   strOld：原密码
    '返回： 加密生成的密码
    '------------------------------------------------
    Dim iBit As Integer, strBit As String
    Dim strNew As String
    If Len(Trim(strOld)) = 0 Then TranZLHISPasswd = "": Exit Function
    strNew = ""
    For iBit = 1 To Len(Trim(strOld))
        strBit = UCase(Mid(Trim(strOld), iBit, 1))
        Select Case (iBit Mod 3)
        Case 1
            strNew = strNew & _
                Switch(strBit = "0", "W", strBit = "1", "I", strBit = "2", "N", strBit = "3", "T", strBit = "4", "E", strBit = "5", "R", strBit = "6", "P", strBit = "7", "L", strBit = "8", "U", strBit = "9", "M", _
                   strBit = "A", "H", strBit = "B", "T", strBit = "C", "I", strBit = "D", "O", strBit = "E", "K", strBit = "F", "V", strBit = "G", "A", strBit = "H", "N", strBit = "I", "F", strBit = "J", "J", _
                   strBit = "K", "B", strBit = "L", "U", strBit = "M", "Y", strBit = "N", "G", strBit = "O", "P", strBit = "P", "W", strBit = "Q", "R", strBit = "R", "M", strBit = "S", "E", strBit = "T", "S", _
                   strBit = "U", "T", strBit = "V", "Q", strBit = "W", "L", strBit = "X", "Z", strBit = "Y", "C", strBit = "Z", "X", True, strBit)
        Case 2
            strNew = strNew & _
                Switch(strBit = "0", "7", strBit = "1", "M", strBit = "2", "3", strBit = "3", "A", strBit = "4", "N", strBit = "5", "F", strBit = "6", "O", strBit = "7", "4", strBit = "8", "K", strBit = "9", "Y", _
                   strBit = "A", "6", strBit = "B", "J", strBit = "C", "H", strBit = "D", "9", strBit = "E", "G", strBit = "F", "E", strBit = "G", "Q", strBit = "H", "1", strBit = "I", "T", strBit = "J", "C", _
                   strBit = "K", "U", strBit = "L", "P", strBit = "M", "B", strBit = "N", "Z", strBit = "O", "0", strBit = "P", "V", strBit = "Q", "I", strBit = "R", "W", strBit = "S", "X", strBit = "T", "L", _
                   strBit = "U", "5", strBit = "V", "R", strBit = "W", "D", strBit = "X", "2", strBit = "Y", "S", strBit = "Z", "8", True, strBit)
        Case 0
            strNew = strNew & _
                Switch(strBit = "0", "6", strBit = "1", "J", strBit = "2", "H", strBit = "3", "9", strBit = "4", "G", strBit = "5", "E", strBit = "6", "Q", strBit = "7", "1", strBit = "8", "X", strBit = "9", "L", _
                   strBit = "A", "S", strBit = "B", "8", strBit = "C", "5", strBit = "D", "R", strBit = "E", "7", strBit = "F", "M", strBit = "G", "3", strBit = "H", "A", strBit = "I", "N", strBit = "J", "F", _
                   strBit = "K", "O", strBit = "L", "4", strBit = "M", "K", strBit = "N", "Y", strBit = "O", "D", strBit = "P", "2", strBit = "Q", "T", strBit = "R", "C", strBit = "S", "U", strBit = "T", "P", _
                   strBit = "U", "B", strBit = "V", "Z", strBit = "W", "0", strBit = "X", "V", strBit = "Y", "I", strBit = "Z", "W", True, strBit)
        End Select
    Next
    TranZLHISPasswd = strNew
End Function

