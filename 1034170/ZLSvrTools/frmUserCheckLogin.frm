VERSION 5.00
Begin VB.Form frmUserCheckLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "用户验证"
   ClientHeight    =   2700
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   4875
   Icon            =   "frmUserCheckLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   4875
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdReloadSvr 
      Caption         =   "刷新服务器(&R)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   216
      TabIndex        =   10
      Top             =   2256
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.ComboBox cboServer 
      Height          =   276
      Left            =   1716
      Sorted          =   -1  'True
      TabIndex        =   5
      Top             =   1716
      Width           =   2592
   End
   Begin VB.Frame fraSplit 
      Height          =   120
      Left            =   0
      TabIndex        =   8
      Top             =   1992
      Width           =   4965
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2268
      TabIndex        =   6
      Top             =   2256
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3528
      TabIndex        =   7
      Top             =   2256
      Width           =   1100
   End
   Begin VB.TextBox txtUser 
      Height          =   300
      Left            =   1716
      MaxLength       =   30
      TabIndex        =   1
      Top             =   900
      Width           =   2592
   End
   Begin VB.TextBox txtPWD 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1716
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1308
      Width           =   2592
   End
   Begin VB.Image imgFlag 
      Height          =   720
      Left            =   210
      Picture         =   "frmUserCheckLogin.frx":1CFA
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lblDataBase 
      AutoSize        =   -1  'True
      Caption         =   "服务器"
      Height          =   180
      Left            =   1092
      TabIndex        =   4
      Top             =   1776
      Width           =   540
   End
   Begin VB.Label lblUser 
      AutoSize        =   -1  'True
      Caption         =   "用户名"
      Height          =   180
      Left            =   1092
      TabIndex        =   0
      Top             =   960
      Width           =   540
   End
   Begin VB.Label lblPWD 
      AutoSize        =   -1  'True
      Caption         =   "口令"
      Height          =   180
      Left            =   1272
      TabIndex        =   2
      Top             =   1368
      Width           =   360
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "    请输入""Rac1(testbase)""的服务器，并进行用户验证"
      Height          =   360
      Left            =   1140
      TabIndex        =   9
      Top             =   240
      Width           =   3552
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmUserCheckLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrUser As String
Private mstrServer As String
Private mcnOracle As ADODB.Connection '验证用户的连接
Private muctCurType As Integer
Private mstrSystems As String
Private mblnFirst As Boolean  '为True表示已经正常显示出
Private mintTimes As Integer  '登录重试次数
Private mcolServer As New Collection
Private mblnOK As Boolean
Private mstrRacInfo As String 'RAC的信息
Public Function ShowLogin(Optional ByVal uctType As UserCheckType, Optional ByRef cnOracle As ADODB.Connection, Optional ByRef strUser As String, Optional ByVal strServer As String, Optional ByVal strSystems As String, Optional ByVal strRacInfo As String) As Boolean
'功能：验证用户登录
'参数：
'          cnOracle=返回的连接
'          strUser=验证的用户
'          strSystems=普通用户验证（ uctType=UCT_NormalUser）时限定用户所属系统。
'          strRacInfo=Rac验证时，RAC标识信息（ uctType=UCT_RACInsUser）,格式为：INST_ID,DBID,Instance_Name(DBname)
'说明：普通用户登录时以系统所有者用户连接数据库时的验证，用户输入的密码不是数据库密码
    muctCurType = uctType
    mstrUser = Decode(uctType, UCT_ZLTOOLS, "ZLTOOLS", strUser)
    mstrServer = IIf(strServer = "" And uctType <> UCT_RACInsUser, gstrServer, strServer)
    mstrRacInfo = strRacInfo
    mstrSystems = strSystems
    Me.Show 1
    Set cnOracle = mcnOracle
    If uctType = UCT_NormalUser Or uctType = UCT_SysOwner Then
        strUser = mstrUser
    End If
    ShowLogin = mblnOK
End Function

Private Sub cmdCancel_Click()
    mblnOK = False
    Set mcnOracle = Nothing
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strNote As String
    Dim strUser As String, strPwd As String, strServer As String
    Dim intPos As Integer
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim arrTmp As Variant
    
    SetConState False
    mintTimes = mintTimes + 1
    '------检验用户是否oracle合法用户----------------
    strUser = Trim(txtUser.Text)
    strPwd = Trim(txtPWD.Text)
    strServer = Trim(cboServer.Text)
    
    '有效字符串效验
    If Len(Trim(txtUser.Text)) = 0 Then
        strNote = "请输入用户名。"
        txtUser.SetFocus
        GoTo InputError
    End If
    
    If Len(strUser) <> 1 Then
        If Mid(strUser, 1, 1) = "/" Or Mid(strUser, 1, 1) = "@" Or Mid(strUser, Len(strUser) - 1, 1) = "/" Or Mid(strUser, Len(strUser) - 1, 1) = "@" Then
            txtUser.SetFocus
            strNote = "用户名错误。"
            Exit Sub
        End If
    End If
    If Trim(strPwd) <> "" And Len(strPwd) <> 1 Then
        If Mid(strPwd, Len(strPwd) - 1, 1) = "/" Or Mid(strPwd, Len(strPwd) - 1, 1) = "@" Or Mid(strPwd, 1, 1) = "/" Or Mid(strPwd, 1, 1) = "@" Then
            txtPWD.SetFocus
            strNote = "口令错误。"
            GoTo InputError
        End If
    End If
    If Trim(strServer) <> "" Then
        If Mid(strServer, Len(strServer) - 1, 1) = "/" Or Mid(strServer, Len(strServer) - 1, 1) = "@" Or Mid(strServer, 1, 1) = "/" Or Mid(strServer, 1, 1) = "@" Then
            strNote = "主机连接串错误。"
            cboServer.SetFocus
            GoTo InputError
        End If
    End If
    
    '分离字符串
    intPos = InStr(1, strUser, "@", vbTextCompare)
    If intPos > 0 Then
        strServer = Mid(strUser, intPos + 1)
        strUser = Mid(strUser, 1, intPos - 1)
    End If
    
    intPos = InStr(1, strUser, "/", vbTextCompare)
    If intPos > 0 Then
        strPwd = Mid(strUser, intPos + 1)
        strUser = Mid(strUser, 1, intPos - 1)
    End If
    
    intPos = InStr(1, strPwd, "@", vbTextCompare)
    If intPos > 0 Then
        strServer = Mid(strPwd, intPos + 1)
        strPwd = Mid(strPwd, 1, intPos - 1)
    End If
    
    If Len(Trim(strPwd)) = 0 Then
        strNote = "请输入密码"
        txtPWD.SetFocus
        GoTo InputError
    End If
    
    strUser = UCase(strUser)
    
    If Not OracleOpen(strServer, strUser, strPwd, strNote) Then
        txtPWD.Text = ""
        If txtPWD.Enabled Then txtPWD.SetFocus
        SetConState
        If strNote <> "" Then GoTo InputError
        Exit Sub
    End If
    
    Select Case muctCurType
        Case UCT_ZLTOOLS
            gstrToolsPwd = strPwd
            Set gcnTools = mcnOracle
        Case UCT_CurZLBAK
        Case UCT_DBAUser
            strSQL = "SELECT 1 FROM SESSION_ROLES WHERE ROLE='DBA'"
            Set rsTmp = gclsBase.OpenSQLRecord(mcnOracle, strSQL, "DBA判断")
            If rsTmp.EOF Then
                MsgBox "当前用户不是不具有DBA角色，请使用其他用户验证！", vbInformation, gstrSysName
                txtUser.SetFocus
                Exit Sub
            End If
            gstrSysUser = strUser
            gstrSysPwd = strPwd
            Set gcnSystem = mcnOracle
        Case UCT_NormalUser
            mstrUser = strUser
        Case UCT_SysOwner
            strSQL = "Select 1 存在  From Session_Roles Where Role = 'DBA'" & vbNewLine & _
                            "Union All" & vbNewLine & _
                            "Select 1 存在 From Zltools.Zlsystems Where Upper(所有者) = User"
            Set rsTmp = gclsBase.OpenSQLRecord(mcnOracle, strSQL, "管理工具管理员判断")
            If rsTmp.EOF Then
                MsgBox "当前用户不具备管理工具管理员权限！", vbInformation, gstrSysName
                txtUser.SetFocus
                Exit Sub
            Else
                mstrUser = strUser
                Set gcnOracle = mcnOracle
            End If
        '需要检查是否是指定数据库的指定实例
        Case UCT_RACInsUser
            arrTmp = Split(mstrRacInfo, ",")
            strSQL = "select 1" & vbNewLine & _
                    "  from v$database a" & vbNewLine & _
                    " where a.DBID = " & arrTmp(1) & vbNewLine & _
                    "   and userenv('instance') = " & arrTmp(0)
            Set rsTmp = gclsBase.OpenSQLRecord(mcnOracle, strSQL, "指定实例判断")
            If rsTmp.EOF Then
                MsgBox "该服务器不是需要验证的实例！", vbInformation, gstrSysName
                cboServer.SetFocus
                Exit Sub
            End If
    End Select
    
    mblnOK = True
    Unload Me
    Exit Sub
InputError:
    If mintTimes > 3 Then
        MsgBox "超过三次登录失败，系统将自动退出。", vbExclamation, gstrSysName
        cmdCancel_Click
    Else
        If strNote <> "" Then
            MsgBox strNote, vbExclamation, gstrSysName
        End If
        SetConState
        Exit Sub
    End If
End Sub

Private Sub cmdReloadSvr_Click()
    Dim strFileInfo As String
    Dim varItem As Variant
    Dim strServer As String
    
    strServer = cboServer.Text
    cboServer.Clear
    Set mcolServer = LoadServer(strFileInfo)
    For Each varItem In mcolServer
        cboServer.AddItem varItem(0)
    Next
    cboServer.ToolTipText = strFileInfo
    cboServer.Text = strServer
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then
        If Trim(txtUser.Text) = "" Then
            cmdOK.Default = False
            If txtUser.Enabled Then txtUser.SetFocus
        Else
            If txtPWD.Enabled Then
                txtPWD.SetFocus
            Else
                cmdOK.SetFocus
            End If
        End If
        mblnFirst = True
        If Trim(txtUser.Text) <> "" And Trim(txtPWD.Text) <> "" Then Call cmdOK_Click
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl.Name = "txtPWD" Then
            Call cmdOK_Click
        Else
            SendKeys "{Tab}"
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub cboServer_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        '回车键另行处理
        If KeyAscii <> vbKeyBack Then
            Call AppendText(KeyAscii)
        End If
    End If
End Sub

Private Sub txtUser_GotFocus()
    If Me.ActiveControl Is txtUser Then
        SelAll txtUser
        OpenIme False
    End If
End Sub

Private Sub txtPwd_GotFocus()
    SelAll txtPWD
End Sub

Private Sub cboServer_GotFocus()
    If Me.ActiveControl Is cboServer Then
        SelAll cboServer
        OpenIme False
    End If
End Sub

Private Sub Form_Load()
    Dim strFileInfo As String
    Dim varItem As Variant

    HookDefend txtPWD.hwnd
    mblnFirst = False
    mintTimes = 1
    If muctCurType = UCT_RACInsUser Then
        cmdReloadSvr.Enabled = True
        cmdReloadSvr.Visible = True
    End If
    '普通用户登录验证
    If muctCurType = UCT_NormalUser Then
        txtUser.Text = GetSetting(appName:="ZLSOFT", Section:="注册信息\登陆信息", Key:="USER", Default:="")
    Else
        txtUser.Text = mstrUser
        txtUser.Enabled = mstrUser = ""
    End If
    
    If mstrServer <> "" Then
        cboServer.Locked = True
        cboServer.Text = mstrServer
    Else
        Set mcolServer = LoadServer(strFileInfo)
        For Each varItem In mcolServer
            cboServer.AddItem varItem(0)
        Next
        cboServer.ToolTipText = strFileInfo
    End If
    
    Call ApplyOEM_Picture(Me, "Icon")
    Select Case muctCurType
        Case UCT_ZLTOOLS
            lblNote.Caption = "    请输入ZLTOOLS的密码。"
        Case UCT_CurZLBAK
            lblNote.Caption = "    请输入该历史库的密码。"
        Case UCT_DBAUser
            lblNote.Caption = "    请输入具有数据库DBA角色的用户。"
        Case UCT_NormalUser
            lblNote.Caption = "    请输入系统的授权用户进行验证。"
        Case UCT_SysOwner
            lblNote.Caption = "    请输入应用系统的所有者用户进行验证。"
        Case UCT_RACInsUser
            lblNote.Caption = "    请输入""" & Split(mstrRacInfo, ",")(2) & """的服务器，并进行用户验证"
    End Select
End Sub

Private Sub AppendText(KeyAscii As Integer)
'功能：向TextBox控件的Text追加内容，并根据当前Text的值在列表中检索可用的完整项目
'参数：KeyAscii    当前的按键
    Dim strTemp As String
    Dim strInput As String
    Dim lngIndex As Long, lngStart As Long
    Dim varItem As Variant
    
    '首先当前用户输入的字符
    If KeyAscii < 0 Or InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.", UCase(Chr(KeyAscii))) > 0 Then
        '输入字符只能是数字、英文和汉字
        strInput = Chr(KeyAscii)
        KeyAscii = 0
    End If
    
    With cboServer
        '记录上次的插入点位置
        lngStart = .SelStart + IIf(strInput <> "", 1, 0)
        '接着得到用户击键完成后文本框中出现的内容
        strInput = Mid(.Text, 1, .SelStart) & strInput & Mid(.Text, .SelStart + .SelLength + 1)
    End With
    '根据假想的内容得到可能的列表项
    strTemp = ""
    For Each varItem In mcolServer
        If UCase(varItem(0)) Like UCase(strInput & "*") Then
            strTemp = varItem(0)
        End If
    Next
    If strTemp <> "" Then
        cboServer.Text = strTemp
        cboServer.SelStart = Len(strInput)
        cboServer.SelLength = 100
    Else
        cboServer.Text = strInput
        cboServer.SelStart = lngStart
    End If
End Sub

Private Sub SetConState(Optional ByVal BlnState As Boolean = True)
    cmdOK.Enabled = BlnState
    cmdCancel.Enabled = BlnState
End Sub

Private Function OracleOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strPassword As String, Optional ByRef strErr As String) As Boolean
'功能： 打开指定的数据库
    Dim blnOwner As Boolean, blnTransPassword As Boolean
    Dim ctTmp As enuProvider
    strErr = ""
    If muctCurType <> UCT_RACInsUser Then
        blnTransPassword = muctCurType = UCT_NormalUser Or muctCurType = UCT_SysOwner
    Else
        blnTransPassword = Not (strUserName = "SYS" Or strUserName = "SYSTEM" Or strUserName = "ZLTOOLS")
    End If
    '特殊用户连接的获取，采用ODBC连接，因为不会用于一般的查询，或者执行过程，只会进行数据库的管理操作或者结构调整
    If Not blnTransPassword Then
        ctTmp = MSODBC
    Else
        ctTmp = OraOLEDB
    End If
    Set mcnOracle = gobjRegister.GetConnection(strServerName, strUserName, strPassword, blnTransPassword, ctTmp, strErr, muctCurType = UCT_SysOwner)
    If mcnOracle.State = adStateClosed Then
         OracleOpen = False
        Set mcnOracle = Nothing
        If muctCurType = UCT_NormalUser Or muctCurType = UCT_SysOwner Then
            Exit Function
        End If
    End If

    On Error GoTo errHand
    mstrUser = strUserName
    If muctCurType = UCT_NormalUser Then
        OracleOpen = zlGetUserInfo(mstrSystems, blnOwner)
        If Not blnOwner And Not OracleOpen Then
            MsgBox "请使用应用系统的授权用户进行验证！", vbOKOnly, gstrSysName
        End If
        mcnOracle.Close
        Set mcnOracle = Nothing
    Else
        OracleOpen = Not mcnOracle Is Nothing
    End If
    Exit Function
errHand:
    MsgBox "注意:" & vbCrLf & "    登陆失败,详细错误信息为:" & vbCrLf & _
           "错误信息:" & err.Number & "-" & err.Description, vbOKOnly, gstrSysName
    OracleOpen = False
    err = 0
End Function

Private Function zlGetUserInfo(ByVal strSystems As String, Optional ByRef blnOwner As Boolean) As Boolean
    Dim rsTmp As New ADODB.Recordset, rsUser As New ADODB.Recordset
    Dim strSQL As String, i As Integer
    
    '读用户信息赋予公共，便于其他程序使用
    zlGetUserInfo = False
    blnOwner = False
    With rsTmp
        If .State = adStateOpen Then .Close
        strSQL = "Select S.所有者" & _
                " From zlSystems S,(Select Distinct owner From All_Tables Where Table_Name='部门表') D" & _
                " Where Upper(S.所有者)=D.Owner And S.编号 In (" & strSystems & ") Order by S.编号"
        .Open strSQL, mcnOracle, adOpenKeyset
        If Not .EOF Then
            '因为可能该用户具有多个系统的身份，所以循环取身份
            If mstrUser = Nvl(!所有者) Then
                  MsgBox "注意:" & vbCrLf & "   不能以所有者身份登陆,请用另外的身份进行登陆!", vbOKOnly, gstrSysName
                  blnOwner = True
                  Exit Function
            End If

            For i = 1 To .RecordCount
                strSQL = "Select R.缺省,D.编码 as 部门编码,D.名称 as 部门名称,P.编号,P.姓名,P.简码" & _
                        " From " & !所有者 & ".上机人员表 U," & !所有者 & ".人员表 P," & !所有者 & ".部门表 D," & !所有者 & ".部门人员 R" & _
                        " Where U.人员ID = P.ID And R.部门ID = D.ID And P.ID=R.人员ID and U.用户名=USER And (P.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or P.撤档时间 Is Null) and R.缺省=1"
                Set rsUser = New ADODB.Recordset
                rsUser.CursorLocation = adUseClient
                rsUser.Open strSQL, mcnOracle, adOpenKeyset
                If Not rsUser.EOF Then
                    zlGetUserInfo = True
                    Exit For
                End If
                .MoveNext
            Next
        End If
        .Close
    End With
End Function

