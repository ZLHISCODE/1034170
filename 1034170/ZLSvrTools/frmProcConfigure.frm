VERSION 5.00
Begin VB.Form frmProcConfigure 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "配置"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4995
   Icon            =   "frmProcConfigure.frx":0000
   LinkTopic       =   "连接配置"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   4995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtPort 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2190
      MaxLength       =   20
      TabIndex        =   4
      Top             =   2175
      Width           =   2625
   End
   Begin VB.TextBox txtServerIP 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2190
      MaxLength       =   20
      TabIndex        =   2
      Top             =   1410
      Width           =   2625
   End
   Begin VB.TextBox txtSID 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2190
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1800
      Width           =   2625
   End
   Begin VB.TextBox txtUserName 
      Height          =   300
      Left            =   2190
      TabIndex        =   0
      Top             =   615
      Width           =   2625
   End
   Begin VB.TextBox txtUserPwd 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   2190
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1005
      Width           =   2625
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2145
      TabIndex        =   8
      Top             =   2880
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3330
      TabIndex        =   7
      Top             =   2880
      Width           =   1100
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "测试连接(&T)"
      Height          =   350
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "单击此处测试连接"
      Top             =   2865
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   -75
      TabIndex        =   5
      Top             =   2610
      Width           =   5310
   End
   Begin VB.Image imgFlag 
      Height          =   720
      Left            =   225
      Picture         =   "frmProcConfigure.frx":000C
      Top             =   615
      Width           =   720
   End
   Begin VB.Label lblPort 
      AutoSize        =   -1  'True
      Caption         =   "端口号"
      Height          =   180
      Left            =   1545
      TabIndex        =   14
      Top             =   2205
      Width           =   540
   End
   Begin VB.Label LblIP 
      AutoSize        =   -1  'True
      Caption         =   "数据服务器IP"
      Height          =   180
      Left            =   1005
      TabIndex        =   13
      Top             =   1470
      Width           =   1080
   End
   Begin VB.Label lblSID 
      AutoSize        =   -1  'True
      Caption         =   "数据库实例"
      Height          =   180
      Left            =   1185
      TabIndex        =   12
      Top             =   1860
      Width           =   900
   End
   Begin VB.Label lblMarks 
      BackStyle       =   0  'Transparent
      Caption         =   "配置页面展示所连接的数据库，请填写对应数据库的用户名、密码、IP以及实例名"
      Height          =   390
      Left            =   240
      TabIndex        =   11
      Top             =   150
      Width           =   4590
   End
   Begin VB.Label lblUserName 
      AutoSize        =   -1  'True
      Caption         =   "连接用户"
      Height          =   180
      Left            =   1365
      TabIndex        =   10
      Top             =   675
      Width           =   720
   End
   Begin VB.Label lblPwd 
      AutoSize        =   -1  'True
      Caption         =   "用户密码"
      Height          =   180
      Left            =   1365
      TabIndex        =   9
      Top             =   1065
      Width           =   720
   End
End
Attribute VB_Name = "frmProcConfigure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrUserName As String
Private mstrUserPwd As String
Private mstrServerIP As String
Private mstrSID As String
Private mstrPort As String
Private mstrConnection As String
Private mobjMain As Object
Private mblnOK As Boolean
Private mblnLocal As Boolean
Private mclsCiph As clsCipher
Private mcnThis As ADODB.Connection

Public Event AfterConn(ByVal cnOracle As ADODB.Connection)

Public Function ShowConfigure(ByVal objMain As Object) As Boolean
    Set mobjMain = objMain
    Me.Show 1, mobjMain
    ShowConfigure = mblnOK
End Function

Private Function OraDataOpen(ByVal strServerIP As String, ByVal strSID As String, ByVal strUserName As String, ByVal strUserPwd As String, ByVal strPort As String) As Boolean
    '------------------------------------------------
    '功能： 打开指定的数据库
    '参数：
    '   strServerName：主机字符串
    '   strUserName：用户名
    '   strUserPwd：密码
    '返回： 数据库打开成功，返回true；失败，返回false
    '------------------------------------------------
    Dim strSQL As String
    Dim strError As String
    Dim cnOracle As New ADODB.Connection
    Dim strServer As String
    
    On Error Resume Next
    err = 0
    DoEvents
    strServer = "(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=" & strServerIP & ")(PORT = " & strPort & ")))(CONNECT_DATA=(SERVICE_NAME=" & strSID & ")))"
    With cnOracle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServer, strUserName, strUserPwd
        If err <> 0 Then
            '保存错误信息
            strError = err.Description
            If InStr(strError, "自动化错误") > 0 Then
                MsgBox "连接串无法创建，请检查数据访问部件是否正常安装。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "无法分析服务器名，" & vbCrLf & "请检查在Oracle配置中是否存在该本地网络服务名（主机字符串）。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "无法连接，请检查服务器上的Oracle监听器服务是否启动。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE正在初始化或在关闭，请稍候再试。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBox "ORACLE不可用，请检查服务或数据库实例是否启动。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBox "用户" & UCase(strUserName) & "已经登录，不允许重复登录(已达到系统所允许的最大登录数)。", vbExclamation, gstrSysName
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBox "由于用户、口令或服务器指定错误，无法登录。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-28000") > 0 Then
                MsgBox "由于用户已经被禁用，无法登录。", vbInformation, gstrSysName
            Else
                MsgBox strError, vbInformation, gstrSysName
            End If
            
            OraDataOpen = False
            Exit Function
        End If
    End With
    Set mcnThis = cnOracle
    
    err = 0
    On Error GoTo errHand
    
    OraDataOpen = True
    Exit Function
    
errHand:
    MsgBox err.Description, vbCritical, Me.Caption
    OraDataOpen = False
    err = 0
End Function

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()

    Dim rsSQL As ADODB.Recordset
    Dim clsCiph As New clsCipher
    Dim lngRow As Long

    On Error GoTo errHand
    
    If TestConnect = False Then Exit Sub
    
    Call gclsBase.SQLRecord(rsSQL)

    mstrConnection = mstrUserName & "/" & mstrUserPwd & "/" & mstrSID & "/" & mstrServerIP & "/" & mstrPort
    mstrConnection = clsCiph.Cipher("zLw09OewKKO1`;owEWO-=,./w[]wwqq3##=``44314325", mstrConnection)
    
    '对字符中&符号的处理
    mstrConnection = Replace(mstrConnection, "&", "' || chr(38) || '")
    gstrSQL = "Zl_Zlprocedureconnect_Update('连接配置','" & mstrConnection & "')"
    Call gclsBase.SQLRecordAdd(rsSQL, gstrSQL)
    
    If SQLRecordExecute(rsSQL) Then
        mblnOK = True
        RaiseEvent AfterConn(mcnThis)
        Unload Me
    End If
    Set clsCiph = Nothing
    Exit Sub
errHand:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Function TestConnect() As Boolean

    Dim strUserName As String
    Dim strServerIP As String
    Dim strPassword As String
    Dim strSID As String
    Dim strPort As String
    Dim strNote As String
    Dim strPwdTmp As String
    
    On Error GoTo InputError
    
    '------检验用户是否oracle合法用户----------------
    strUserName = Trim(txtUserName.Text)
    strPassword = Trim(txtUserPwd.Text)
    strServerIP = Trim(txtServerIP.Text)
    strSID = Trim(txtSID.Text)
    strPort = Trim(txtPort.Text)
    
    '有效字符串效验
    If Len(Trim(txtUserName.Text)) = 0 Then
        strNote = "请输入用户名"
        txtUserName.SetFocus
        GoTo InputError
    End If
    
    If Len(Trim(txtUserPwd.Text)) = 0 Then
        strNote = "请输入密码"
        txtUserName.SetFocus
        GoTo InputError
    End If
    
    If Len(Trim(txtSID.Text)) = 0 Then
        strNote = "请输入数据库实例"
        txtSID.SetFocus
        GoTo InputError
    End If
    
    If Len(Trim(txtServerIP.Text)) = 0 Then
        strNote = "请输入数据库IP"
        txtServerIP.SetFocus
        GoTo InputError
    End If
    
    If Len(Trim(txtPort.Text)) = 0 Then
        strNote = "请输入端口号"
        txtServerIP.SetFocus
        GoTo InputError
    End If
    
    If Len(strUserName) <> 1 Then
        If Mid(strUserName, 1, 1) = "/" Or Mid(strUserName, 1, 1) = "@" Or Mid(strUserName, Len(strUserName) - 1, 1) = "/" Or Mid(strUserName, Len(strUserName) - 1, 1) = "@" Then
            txtUserName.SetFocus
            strNote = "用户名错误"
            GoTo InputError
        End If
    End If
    If Trim(strPassword) <> "" And Len(strPassword) <> 1 Then
        If Mid(strPassword, Len(strPassword) - 1, 1) = "/" Or Mid(strPassword, Len(strPassword) - 1, 1) = "@" Or Mid(strPassword, 1, 1) = "/" Or Mid(strPassword, 1, 1) = "@" Then
            txtUserPwd.SetFocus
            strNote = "密码错误"
            GoTo InputError
        End If
    End If
    
    strPwdTmp = strPassword
    strUserName = UCase(strUserName)
    If UCase(strUserName) <> "SYSTEM" And UCase(strUserName) <> "SYS" Then
        strPassword = TranPasswd(strPassword)
    End If

    If OraDataOpen(strServerIP, strSID, strUserName, strPassword, strPort) Then
        mstrUserName = strUserName
        mstrUserPwd = strPwdTmp
        mstrServerIP = strServerIP
        mstrSID = strSID
        mstrPort = strPort
                
        TestConnect = True
    End If
    
    Exit Function
    
InputError:
    If strNote <> "" Then
        MsgBox strNote, vbExclamation, gstrSysName
    End If
    Exit Function
    
End Function

Private Sub cmdTest_Click()
    If TestConnect = True Then
        MsgBox "测试连接成功，可以正常访问该数据库。", vbInformation + vbOKOnly, gstrSysName
    End If
    
'    If OraDataOpen(strServerIP, strSID, strUserName, strPassword, strPort) Then
'        mstrUserName = strUserName
'        mstrUserPwd = strPassword
'        mstrServerIP = strServerIP
'        mstrSID = strSID
'        mstrPort = strPort
'        cmdOK.Enabled = True
'        MsgBox "测试连接成功，可以正常访问该数据库。", vbInformation + vbOKOnly, gstrSysName
'        Exit Sub
'    Else
'        cmdOK.Enabled = False
'    End If
End Sub

Private Sub Form_Activate()

    On Error GoTo errHand
    Dim strConnection As String
    Dim strCon() As String
    
    Set mclsCiph = New clsCipher
    strConnection = gclsBase.GetOraConn("连接配置")
    If strConnection <> "" Then
        strConnection = mclsCiph.Decipher("zLw09OewKKO1`;owEWO-=,./w[]wwqq3##=``44314325", strConnection)
        strCon = Split(strConnection, "/")
        txtUserName.Text = strCon(0)
        txtUserPwd.Text = strCon(1)
        txtSID.Text = strCon(2)
        txtServerIP.Text = strCon(3)
        txtPort.Text = strCon(4)
    End If
    Set mclsCiph = Nothing
    txtUserPwd.SetFocus
    Exit Sub
errHand:
    MsgBox err.Description, vbCritical, Me.Caption
End Sub

Private Sub Form_Load()
    txtUserName.Text = GetSetting(appName:="ZLSOFT", Section:="注册信息\登陆信息", Key:="USER", Default:="")
    txtPort.Text = "1521"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not (mclsCiph Is Nothing) Then
        Set mclsCiph = Nothing
    End If
End Sub

Private Sub txtPort_GotFocus()
    gclsBase.TxtSelAll txtPort
End Sub

Private Sub txtServerIP_GotFocus()
    gclsBase.TxtSelAll txtServerIP
End Sub

Private Sub txtSID_GotFocus()
    gclsBase.TxtSelAll txtSID
End Sub

Private Sub txtUserName_GotFocus()
    gclsBase.TxtSelAll txtUserName
End Sub

Private Sub txtUserPwd_GotFocus()
    gclsBase.TxtSelAll txtUserPwd
End Sub

