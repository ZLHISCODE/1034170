VERSION 5.00
Begin VB.Form frmUserLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "操作员登录"
   ClientHeight    =   2205
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4170
   Icon            =   "frmUserLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   4170
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   -360
      TabIndex        =   8
      Top             =   1455
      Width           =   5025
   End
   Begin VB.CommandButton CMD放弃 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2745
      TabIndex        =   4
      Top             =   1710
      Width           =   1100
   End
   Begin VB.CommandButton CDM确认 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1455
      TabIndex        =   3
      Top             =   1710
      Width           =   1100
   End
   Begin VB.TextBox TXT密码 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1950
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   630
      Width           =   1920
   End
   Begin VB.TextBox txt数据库 
      Height          =   300
      Left            =   1950
      TabIndex        =   2
      Top             =   1050
      Width           =   1920
   End
   Begin VB.TextBox txt用户 
      Height          =   300
      Left            =   1950
      TabIndex        =   0
      Top             =   195
      Width           =   1920
   End
   Begin VB.Image imgFlag 
      Height          =   720
      Left            =   315
      Picture         =   "frmUserLogin.frx":0442
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Lbl服务器 
      AutoSize        =   -1  'True
      Caption         =   "服务器"
      Height          =   180
      Left            =   1320
      TabIndex        =   7
      Top             =   1110
      Width           =   540
   End
   Begin VB.Label Lbl口令 
      AutoSize        =   -1  'True
      Caption         =   "口令"
      Height          =   180
      Left            =   1500
      TabIndex        =   6
      Top             =   690
      Width           =   360
   End
   Begin VB.Label Lbl用户名 
      AutoSize        =   -1  'True
      Caption         =   "所有者"
      Height          =   180
      Left            =   1320
      TabIndex        =   5
      Top             =   255
      Width           =   540
   End
End
Attribute VB_Name = "frmUserLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim intTimes As Integer
Dim strNote As String
Dim strUserName As String
Dim strServerName As String
Dim strPassword As String

Private DblHeight As Double
Private DblWidth As Double
Private Const DrawCount As Integer = 100
Private PDC As Long
Private BlnFirst As Boolean

Private Sub CDM确认_Click()
    SetConState False
    intTimes = intTimes + 1
    
    '------检验用户是否oracle合法用户----------------
    strUserName = Trim(txt用户.Text)
    strServerName = Trim(txt数据库.Text)
    strPassword = Trim(TXT密码.Text)
    '有效字符串效验
    If Len(Trim(txt用户)) = 0 Then
        strNote = "请输入用户名"
        txt用户.SetFocus
        GoTo InputError
    End If
    
    If Len(strUserName) <> 1 Then
        If Mid(strUserName, 1, 1) = "/" Or Mid(strUserName, 1, 1) = "@" Or Mid(strUserName, Len(strUserName) - 1, 1) = "/" Or Mid(strUserName, Len(strUserName) - 1, 1) = "@" Then
            txt用户.SetFocus
            strNote = "用户名错误"
            SetConState
            Exit Sub
        End If
    End If
    If Trim(strPassword) <> "" And Len(strPassword) <> 1 Then
        If Mid(strPassword, Len(strPassword) - 1, 1) = "/" Or Mid(strPassword, Len(strPassword) - 1, 1) = "@" Or Mid(strPassword, 1, 1) = "/" Or Mid(strPassword, 1, 1) = "@" Then
            If TXT密码.Enabled Then TXT密码.SetFocus
            strNote = "口令错误"
            GoTo InputError
        End If
    End If
    If Trim(strServerName) <> "" Then
        If Mid(strServerName, Len(strServerName) - 1, 1) = "/" Or Mid(strServerName, Len(strServerName) - 1, 1) = "@" Or Mid(strServerName, 1, 1) = "/" Or Mid(strServerName, 1, 1) = "@" Then
            strNote = "主机连接串错误"
            txt数据库.SetFocus
            GoTo InputError
        End If
    End If
    
    '分离字符串
    Dim intPos As Integer
    intPos = InStr(1, strUserName, "@", vbTextCompare)
    If intPos > 0 Then
        strServerName = Mid(strUserName, intPos + 1)
        strUserName = Mid(strUserName, 1, intPos - 1)
    End If
    
    intPos = InStr(1, strUserName, "/", vbTextCompare)
    If intPos > 0 Then
        strPassword = Mid(strUserName, intPos + 1)
        strUserName = Mid(strUserName, 1, intPos - 1)
    End If
    
    intPos = InStr(1, strPassword, "@", vbTextCompare)
    If intPos > 0 Then
        strServerName = Mid(strPassword, intPos + 1)
        strPassword = Mid(strPassword, 1, intPos - 1)
    End If
    
    
    If Len(Trim(strPassword)) = 0 Then
        strNote = "请输入密码"
        GoTo InputError
    End If
    
    If Not OraDataOpen(strServerName, strUserName, IIf(UCase(strUserName) = "SYS" Or UCase(strUserName) = "SYSTEM", strPassword, TranPasswd(strPassword))) Then
        TXT密码.Text = ""
        If TXT密码.Enabled Then TXT密码.SetFocus
        SetConState
        Exit Sub
    End If
    
    
    '修改注册表
    SaveSetting "ZLSOFT", "公共", "USER", strUserName
    SaveSetting "ZLSOFT", "公共", "SERVER", strServerName
    
    '创建快捷方式用
    SaveSetting "ZLSOFT", "公共", "程序路径", App.Path & "\" & App.EXEName & ".exe"
    
    Unload Me
    Exit Sub
InputError:
    If intTimes > 3 Then
        MsgBox "超过三次登录失败，系统将自动退出", vbExclamation, gstrSysName
        CMD放弃_Click
    Else
        If strNote <> "" Then
            MsgBox strNote, vbExclamation, gstrSysName
        End If
        SetConState
        Exit Sub
    End If

End Sub

Private Sub CMD放弃_Click()
    Set gcnOracle = Nothing
    Unload Me
End Sub

Private Sub Form_Activate()
    Dim LngStyle As Long
    If BlnFirst = False Then
        LngStyle = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
        LngStyle = LngStyle Or WinStyle
        Call SetWindowLong(Me.hwnd, GWL_EXSTYLE, LngStyle)
        
        ShowWindow Me.hwnd, 0 '先隐藏
        ShowWindow Me.hwnd, 1 '再显示
'
'        Call SetWindowPos(Me.hwnd, HWND_TOPMOST, Me.Left / 15, Me.Top / 15, Me.Height / 15, Me.Width / 15, SWP_NOSIZE + SWP_SHOWWINDOW)
    End If
    
    If TXT密码.Enabled Then
        TXT密码.SetFocus
    Else
        CDM确认.SetFocus
    End If
    
    If txt用户.Text <> "" And TXT密码.Text <> "" Then Call CDM确认_Click
    BlnFirst = True
End Sub

Private Sub Form_Load()
    Dim strUser As String, strPass As String, strServer As String, strType As String
    Dim arrCmd
    
    BlnFirst = False
    intTimes = 1
    txt用户.Text = GetSetting(appName:="ZLSOFT", Section:="公共", Key:="USER", Default:="")
    txt数据库.Text = GetSetting(appName:="ZLSOFT", Section:="公共", Key:="SERVER", Default:="")
    
    '解析命令行
    arrCmd = Split(Command(), " ")
    If UBound(arrCmd) >= 0 Then
        strUser = Trim(arrCmd(0))
        If UBound(arrCmd) >= 1 Then
            strPass = Trim(arrCmd(1))
            If UBound(arrCmd) >= 2 Then
                strServer = Trim(arrCmd(2))
                If UBound(arrCmd) = 3 Then
                    strType = Val(arrCmd(3))
                End If
            End If
        End If
    End If
    If strUser <> "" Then txt用户.Text = strUser
    If strPass <> "" Then TXT密码.Text = strPass
    If strServer <> "" Then txt数据库.Text = strServer
    gintType = Val(strType)
End Sub

Private Sub GetFocus(ByVal TxtBox As TextBox)
    With TxtBox
        .SelStart = 0
        .SelLength = LenB(StrConv(.Text, vbFromUnicode))
    End With
End Sub

Private Sub txt用户_GotFocus()
    GetFocus txt用户
End Sub

Private Sub TXT密码_GotFocus()
    GetFocus TXT密码
End Sub

Private Sub txt数据库_GotFocus()
    GetFocus txt数据库
End Sub

Private Sub SetConState(Optional ByVal BlnState As Boolean = True)
    CMD放弃.Enabled = BlnState
    CDM确认.Enabled = BlnState
End Sub
