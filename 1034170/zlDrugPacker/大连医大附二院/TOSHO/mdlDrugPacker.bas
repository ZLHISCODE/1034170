Attribute VB_Name = "mdlDrugPacker"
Option Explicit

Public Const GSTR_MESSAGE = "提示信息"

Public gstrUser As String, gstrUserNameNew As String
Public gintUserID As Integer, gintDeptID As Integer
Public gbyt效期 As Byte

Public gobjComLib As Object                         'zl9Comlib部件
Public gcnOutside As New ADODB.Connection           '外部数据库连接

Public Const GSTR_SYSNAME = "自动分包机接口"
Public Const GSTR_REGEDIT_PATH = "公共模块\DrugPackerDBServer"
Public Const MSTR_SERVER = "localhost"
Public Const MSTR_DBNAME = "atf"
Public Const MSTR_USER = "sa"
Public Const MSTR_PASSWORD = ""


Public Function MSSQLServerOpen(ByVal strServerName As String, ByVal strDBName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '------------------------------------------------
    '功能： 打开指定的MS SQL Server 数据库
    '参数：
    '   strServerName：主机字符串
    '   strUserName：用户名
    '   strUserPwd：密码
    '返回： 数据库打开成功，返回true；失败，返回false
    '------------------------------------------------
    Dim strSQL As String
    Dim strError As String
    
    If Len(Trim(strUserName)) = 0 Then
        MSSQLServerOpen = False
        MsgBox "请设置外联数据库信息！", vbInformation, GSTR_MESSAGE
        Exit Function
    End If
    
    On Error Resume Next
    Err = 0
    DoEvents
    With gcnOutside
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .ConnectionTimeout = 5
        .Open "Driver={SQL Server};Server=" & strServerName & ";Database=" & strDBName, strUserName, strUserPwd
        If Err <> 0 Then
            '保存错误信息
            strError = Err.Description
            If InStr(strError, "自动化错误") > 0 Then
                MsgBox "连接串无法创建，请检查数据访问部件是否正常安装。", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "无法分析服务器名，" & vbCrLf & "请检查在Oracle配置中是否存在该本地网络服务名（主机字符串）。", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "无法连接，请检查服务器上的Oracle监听器服务是否启动。", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE正在初始化或在关闭，请稍候再试。", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBox "ORACLE不可用，请检查服务或数据库实例是否启动。", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBox "用户" & UCase(strUserName) & "已经登录，不允许重复登录(已达到系统所允许的最大登录数)。", vbExclamation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBox "由于用户、口令或服务器指定错误，无法登录。", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-28000") > 0 Then
                MsgBox "由于用户已经被禁用，无法登录。", vbInformation, GSTR_SYSNAME
            ElseIf Err.Number = -2147217843 Or Err.Number = -2147467259 Then
                MsgBox "药品分包机数据库连接失败！", vbInformation, GSTR_SYSNAME
            Else
                MsgBox strError, vbInformation, GSTR_SYSNAME
            End If
            
            MSSQLServerOpen = False
            Exit Function
        End If
    End With
    
    Err = 0
    On Error GoTo errHand
    
    'gstrDbUser = UCase(strUserName)
    'SetDbUser gstrDbUser
    
    MSSQLServerOpen = True
    Exit Function
    
errHand:
    If gobjComLib.ErrCenter() = 1 Then Resume
    MSSQLServerOpen = False
    Err = 0
End Function


Public Function OraDataOpenTest(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
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
    
    On Error Resume Next
    Err = 0
    DoEvents
    With gcnOutside
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
        If Err <> 0 Then
            '保存错误信息
            strError = Err.Description
            If InStr(strError, "自动化错误") > 0 Then
                MsgBox "连接串无法创建，请检查数据访问部件是否正常安装。", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "无法分析服务器名，" & vbCrLf & "请检查在Oracle配置中是否存在该本地网络服务名（主机字符串）。", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "无法连接，请检查服务器上的Oracle监听器服务是否启动。", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE正在初始化或在关闭，请稍候再试。", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-01034") > 0 Then
                MsgBox "ORACLE不可用，请检查服务或数据库实例是否启动。", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-02391") > 0 Then
                MsgBox "用户" & UCase(strUserName) & "已经登录，不允许重复登录(已达到系统所允许的最大登录数)。", vbExclamation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-01017") > 0 Then
                MsgBox "由于用户、口令或服务器指定错误，无法登录。", vbInformation, GSTR_SYSNAME
            ElseIf InStr(strError, "ORA-28000") > 0 Then
                MsgBox "由于用户已经被禁用，无法登录。", vbInformation, GSTR_SYSNAME
            ElseIf Err.Number = -2147217843 Then
                MsgBox Mid(strError, InStr(1, strError, "[SQL Server]"), Len(strError)), vbInformation, GSTR_SYSNAME
            Else
                MsgBox strError, vbInformation, GSTR_SYSNAME
            End If
            
            OraDataOpenTest = False
            Exit Function
        End If
    End With
    
    Err = 0
    On Error GoTo errHand
    
    'gstrDbUser = UCase(strUserName)
    'SetDbUser gstrDbUser
    
    OraDataOpenTest = True
    Exit Function
    
errHand:
    If gobjComLib.ErrCenter() = 1 Then Resume
    OraDataOpenTest = False
    Err = 0
End Function

Public Function StringEnDeCodecn(strSource As String, MA) As String
'该函数只对中西文起到加密作用
'参数为：源文件，密码
    On Error GoTo ErrEnDeCode
    Dim X As Single, i As Integer
    Dim CHARNUM As Long, RANDOMINTEGER As Integer
    Dim SINGLECHAR As String * 1
    Dim strTmp As String
    
    If MA < 0 Then
        MA = MA * (-1)
    End If
    
    X = Rnd(-MA)
    For i = 1 To Len(strSource) Step 1                 '取单字节内容
        SINGLECHAR = Mid(strSource, i, 1)
        CHARNUM = Asc(SINGLECHAR)
g:
        RANDOMINTEGER = Int(127 * Rnd)
        If RANDOMINTEGER < 30 Or RANDOMINTEGER > 100 Then GoTo g
        CHARNUM = CHARNUM Xor RANDOMINTEGER
        strTmp = strTmp & Chr(CHARNUM)
    Next i
    StringEnDeCodecn = strTmp
    Exit Function

ErrEnDeCode:
    StringEnDeCodecn = ""
    MsgBox Err.Number & "\" & Err.Description
End Function

Public Function GetUserNameInfo() As Boolean
'获取用户信息
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    Set rsTmp = gobjComLib.GetUserInfo
    
    With rsTmp
        If Not .EOF Then
            gintUserID = IIf(IsNull(!Id), 0, !Id)
            gintDeptID = IIf(IsNull(!部门id), 0, !部门id)
            gstrUserNameNew = IIf(IsNull(!姓名), "", !姓名) '当前用户姓名
            GetUserNameInfo = True
        Else
            gintUserID = 0
            gintDeptID = 0
            gstrUserNameNew = "" '当前用户姓名
        End If
    End With
    rsTmp.Close

    strSQL = "Select 参数号, 参数值, 缺省值 From Zlparameters Where 系统 = [1] And Nvl(私有, 0) = 0 And 模块 Is Null and 参数号=[2] "
    Set rsTmp = gobjComLib.OpenSQLRecord(strSQL, "取系统参数", 100, 149)
    With rsTmp
        If Not .EOF Then
            gbyt效期 = IIf(IsNull(rsTmp!参数值), rsTmp!缺省值, rsTmp!参数值)
        Else
            gbyt效期 = 0
        End If
    End With
    
End Function
'
'Public Function CheckProvider(ByVal intProvider As Integer) As String
''审核供应商ID
'    Dim rsTmp As New ADODB.Recordset
'    Set rsTmp = zlDatabase.OpenSQLRecord("select 名称 from 供应商 where id=[1]", "审核供应商ID", intProvider)
'    If rsTmp.RecordCount = 1 Then
'        CheckProvider = rsTmp!名称
'    End If
'    rsTmp.Close
'End Function

Public Sub SelText(ByVal ctlVal As Control)
    If TypeOf ctlVal Is TextBox Then
        ctlVal.SelStart = 0
        ctlVal.SelLength = Len(ctlVal.Text)
    End If
End Sub



