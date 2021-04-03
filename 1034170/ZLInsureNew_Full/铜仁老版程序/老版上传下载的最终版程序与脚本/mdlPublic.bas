Attribute VB_Name = "mdlPublic"
Option Explicit
Public gintType As Integer
Public Enum 医保Enum
    TYPE_铜仁市 = 81
End Enum

'中软公司实现的DLL
Public Declare Function FTPUpLoad Lib "FTP_Trans.dll" (ByVal aHost As String, ByVal aPort As String, ByVal aUserID As String, ByVal aPassWord As String, ByVal aLocalFile As String, ByVal aRemoteDir As String, ByVal aRemoteFileName As String) As Long
Public Declare Function FTPDownLoad Lib "FTP_Trans.dll" (ByVal aHost As String, ByVal aPort As String, ByVal aUserID As String, ByVal aPassWord As String, ByVal aRemoteDir As String, ByVal aRemoteFileName As String, ByVal aLocalFile As String) As Long

Public Declare Function EncryptStr Lib "FTP_Trans.dll" (ByVal SourceStr As String, ByVal Key As String, ByVal IsEncrypt As Boolean) As String
Public Declare Function EncryptFiles Lib "FTP_Trans.dll" (ByVal INFName As String, ByVal OutFName As String) As Long
Public Declare Function DecryptFiles Lib "FTP_Trans.dll" (ByVal INFName As String, ByVal OutFName As String) As Long

Public Const GWL_EXSTYLE = (-20)
Public Const WinStyle = &H40000
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public gcnOracle As New ADODB.Connection
Public gcn医保 As New ADODB.Connection
Public gstrSysName As String
Public gstrOwner As String
Public gstrSQL As String

Public Sub Main()
    Dim lngReturn As Long
    Dim strCode As String, IntCount As Integer, StrStyle As String
    Dim rsMenu As ADODB.Recordset, StrHaveSys As String
    
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "提示", "")
    If gstrSysName = "" Then gstrSysName = "中联软件"
    
    '用户注册
    frmUserLogin.Show 1
    If gcnOracle.State <> adStateOpen Then
        Unload frmUserLogin
        Exit Sub
    End If
    
    If 检查医保服务器 = False Then
        Exit Sub
    End If
    If 检查医保数据表 = False Then
        Exit Sub
    End If
    
    frm上传下载.Show
    
    
End Sub

Public Function OraDataOpen(ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
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
    Dim rsTemp As New ADODB.Recordset

    On Error Resume Next
    DoEvents
    With gcnOracle
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
        If Err <> 0 Then
            '保存错误信息
            strError = Err.Description
            If InStr(strError, "自动化错误") > 0 Then
                MsgBox "连接串无法创建，请检查数据访问部件是否正常安装。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12154") > 0 Then
                MsgBox "无法分析服务器名，" & vbCrLf & "请检查在Oracle配置中是否存在该本地网络服务名（主机字符串）。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-12541") > 0 Then
                MsgBox "无法连接，请检查服务器上的Oracle监听器服务是否启动。", vbInformation, gstrSysName
            ElseIf InStr(strError, "ORA-01033") > 0 Then
                MsgBox "ORACLE正在初始化或在关闭，请稍候再试。", vbInformation, gstrSysName
            Else
                MsgBox "由于用户、口令或服务器指定错误，无法注册。", vbInformation, gstrSysName
            End If
            
            OraDataOpen = False
            Exit Function
        End If
        
        gstrOwner = UCase(strUserName)
        gstrSQL = "Select 编号 From zlsystems where 所有者='" & gstrOwner & "' and trunc(编号/100) in (1,8)"
        rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
        If rsTemp.RecordCount = 0 Then
            MsgBox "登录用户必须是系统所有者。", vbInformation, gstrSysName
            Exit Function
        End If
        
        .Execute "select * from 病人费用记录 where rownum<1"
        If Err <> 0 Then
            MsgBox "你不具有访问HIS数据表的权限。", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    OraDataOpen = True
End Function

Private Function 检查医保服务器() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strUser As String, strPass As String, strServer As String
    
    '读出连接医保服务器的配置
    gstrSQL = "select 参数名,参数值 from 保险参数 where 参数名 like '医保%' and 险类=" & TYPE_铜仁市
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    Do Until rsTemp.EOF
        Select Case rsTemp("参数名")
            Case "医保用户名"
                strUser = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "医保服务器"
                strServer = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
            Case "医保用户密码"
                strPass = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
                '解密
                If strPass <> "" Then strPass = EncryptStr(strPass, 256, False)
        End Select
        rsTemp.MoveNext
    Loop
    
    On Error Resume Next
    gcn医保.Provider = "MSDataShape"
    gcn医保.Open "Driver={Microsoft ODBC for Oracle};Server=" & strServer, strUser, strPass
    
    If Err <> 0 Then
        MsgBox "医保前置服务器连接失败。" & vbCrLf & _
               "请注意，保险类别参数设置中的服务器名在所有机器上应该相同。", vbInformation, gstrSysName
        Exit Function
    End If
    
    检查医保服务器 = True
End Function

Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '功能： 密码转换函数
    '参数：
    '   strOld：原密码
    '返回： 加密生成的密码
    '------------------------------------------------
    Dim iBit As Integer, strBit As String
    Dim strNew As String
    If Len(Trim(strOld)) = 0 Then TranPasswd = "": Exit Function
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
    TranPasswd = strNew

End Function

Public Function Currentdate() As Date
'功能：获得当前日期
    Dim rsTmp As New ADODB.Recordset
    On Error Resume Next
    rsTmp.Open "SELECT SYSDATE FROM DUAL", gcnOracle, adOpenKeyset
    Currentdate = rsTmp.Fields(0).Value
    If Err <> 0 Then
        '用当前机器时间
        Currentdate = date
        Err.Clear
    End If
End Function

Public Function MidUni(ByVal strTemp As String, ByVal Start As Long, ByVal Length As Long) As String
'功能：按数据库规则得到字符串的子集，也就是汉字按两个字符算，而字母仍是一个
    MidUni = StrConv(MidB(StrConv(strTemp, vbFromUnicode), Start, Length), vbUnicode)
End Function


Public Function AddDate(ByVal strOrin As String) As String
'功能：为不全的日期信息补充完整
    Dim strTemp As String
    Dim intPos As Integer
    
    strTemp = Trim(strOrin)
    
    If strTemp = "" Then
        AddDate = ""
        Exit Function
    End If
    
    intPos = InStr(strTemp, "-")
    If intPos = 0 Then
        intPos = InStr(strTemp, ".")
        If intPos <> 0 Then
            '使用 . 隔
            strTemp = Replace(strTemp, ".", "-")
        End If
    End If
    
    If intPos = 0 Then
        '没有"-",手工加上
        intPos = Len(strTemp)
        If intPos <= 8 Then
            If intPos = 8 Then
                strTemp = Mid(strTemp, 1, 4) & "-" & Mid(strTemp, 5, 2) & "-" & Mid(strTemp, 7, 2)
            ElseIf intPos > 4 Then
                strTemp = Left(strTemp, intPos - 4) & "-" & Mid(Right(strTemp, 4), 1, 2) & "-" & Right(strTemp, 2)
            ElseIf intPos > 2 Then
                strTemp = Format(date, "yyyy") & "-" & Left(strTemp, intPos - 2) & "-" & Right(strTemp, 2)
            Else
                strTemp = Format(date, "yyyy") & "-" & Format(date, "MM") & "-" & strTemp
            End If
        End If
    Else
        If IsDate(strTemp) Then
            strTemp = Format(CDate(strTemp), "yyyy-MM-dd")
        End If
    End If
    
    AddDate = strTemp
End Function

Private Function 检查医保数据表() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim strUser As String, strPass As String, strServer As String
    
    '只检查上传下载表
    On Error Resume Next
    gstrSQL = "select * from 上传下载 where rownum<1"
    rsTemp.Open gstrSQL, gcn医保, adOpenStatic, adLockReadOnly
    
    If Err <> 0 Then
        MsgBox "保险类别中设置的医保用户并不具备与医保有关的数据表，请运行安装脚本。", vbInformation, gstrSysName
        Exit Function
    End If
    
    '检查数据
    On Error GoTo errHandle
    
    
    
    检查医保数据表 = True
    Exit Function
errHandle:
    If frmErr.ShowErr(Err.Description) = vbYes Then
        Resume
    End If
End Function

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0, Optional ByVal hwnd As Long = 0, Optional str项目 As String) As Boolean
'检查字符串是否含有非法字符；如果提供长度，对长度的合法性也作检测。
    If str项目 = "" Then str项目 = "所输入内容"
    
    If InStr(strInput, "'") > 0 Or InStr(strInput, "|") > 0 Then
        MsgBox str项目 & "含有非法字符。", vbExclamation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '设置焦点
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox str项目 & "不能超过" & Int(intMax / 2) & "个汉字" & "或" & intMax & "个字母。", vbExclamation, gstrSysName
            If hwnd <> 0 Then SetFocusHwnd hwnd              '设置焦点
            Exit Function
        End If
    End If
    
    StrIsValid = True
End Function


Public Sub OpenRecordset(rsTemp As ADODB.Recordset, _
        Optional CursorType As CursorTypeEnum = adOpenStatic, Optional LockType As LockTypeEnum = adLockReadOnly)
'功能：打开记录。同时保存SQL语句
    If rsTemp.State = adStateOpen Then rsTemp.Close
    
    rsTemp.Open gstrSQL, gcn医保, CursorType, LockType
End Sub

Public Sub ExecuteProcedure()
'功能：执行过程式的SQL语句
    gcn医保.Execute gstrSQL, , adCmdStoredProc
End Sub

Public Function GetMax(ByVal strTable As String, ByVal strField As String, ByVal intLength As Integer, Optional ByVal strWhere As String) As String
'功能：读取指定表的本级编码的最大值
'参数：strTable  表名;
'      strField  字段名;
'      intLength 字段长度
'返回：成功返回 下级最大编码; 否者返回 0
    Dim rsTemp As New ADODB.Recordset
    Dim varTemp As Variant, strSQL As String
    Dim lngLengh As Long
    
    On Error GoTo ErrHand
    With rsTemp
        strSQL = "SELECT MAX(LPAD(" & strField & "," & intLength & ",' ')) as ""最大值"",max(length(" & _
             strField & ")) as ""最长值"" FROM " & strTable & strWhere
        rsTemp.Open strSQL, gcn医保, adOpenStatic, adLockReadOnly
        
        If rsTemp.EOF Then
            GetMax = Format(1, String(intLength, "0"))
            Exit Function
        End If
        varTemp = IIf(IsNull(.Fields("最大值").Value), "0", .Fields("最大值").Value)
        lngLengh = IIf(IsNull(.Fields("最长值").Value), intLength, .Fields("最长值").Value)
        If IsNumeric(varTemp) Then
            GetMax = CStr(Val(varTemp) + 1)
            GetMax = Format(GetMax, String(lngLengh, "0"))
        Else
            GetMax = Mid(varTemp, 1, Len(varTemp) - 1) & Chr(Asc(Right(varTemp, 1)) + 1)
            GetMax = Trim(GetMax)
        End If
        .Close
    End With
    Exit Function
    
ErrHand:
    If frmErr.ShowErr(Err.Description) = vbYes Then Resume
End Function


