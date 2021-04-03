Attribute VB_Name = "mdlPublic"
Option Explicit

Public gstrSysName As String
Public gcnOracle As New ADODB.Connection    'Oracle公共数据库连接
Public gcnYB As New ADODB.Connection    'Sybase公共数据库连接
Public gintInsure As Integer
Public gstr医院编码 As String
Public gstrSQL As String

Public Type TYPE_USER_INFO
    ID As Long
    部门ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
    部门 As String
    站点 As String
End Type
Public UserInfo As TYPE_USER_INFO
'----------------------------------------------------------------
Public glngOld As Long, glngFormW As Long, glngFormH As Long
Public Const WM_GETMINMAXINFO = &H24
Public Const GWL_WNDPROC = -4
Type POINTAPI
     X As Long
     Y As Long
End Type
Public Type MINMAXINFO
        ptReserved As POINTAPI
        ptMaxSize As POINTAPI
        ptMaxPosition As POINTAPI
        ptMinTrackSize As POINTAPI
        ptMaxTrackSize As POINTAPI
End Type
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    

Public Sub Main()
'功能：启动函数
    gstrSysName = GetSetting("ZLSOFT", "注册信息", UCase("gstrSysName"), "")
    
    '用户注册
    frmUserLogin.Show 1
    If gcnOracle.State <> adStateOpen Then
        Unload frmUserLogin
        Exit Sub
    End If
    
    If Not UserIsOwner Then
        MsgBox "对不起，只有HIS所有者才能运行本程序！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If Not GetUserInfo Then Exit Sub
    
    frm消息转发.Show
End Sub

Public Function GetUserInfo() As Boolean
 '功能：获取登陆用户信息
    Dim rsUser As ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo errH
    
    strSql = _
        "Select B.*,C.编码 as 部门编码,C.名称 as 部门名称,D.部门ID,A.用户名 " & _
        " From 上机人员表 A,人员表 B,部门表 C,部门人员 D " & _
        " Where A.人员ID=B.ID And B.ID=D.人员ID And D.缺省=1" & _
        " And D.部门ID=C.ID And A.用户名=USER"
    Set rsUser = New ADODB.Recordset
    rsUser.CursorLocation = adUseClient
    rsUser.Open strSql, gcnOracle, adOpenKeyset
    If Not rsUser.EOF Then
        UserInfo.ID = rsUser!ID
        UserInfo.编号 = rsUser!编号
        UserInfo.部门ID = IIf(IsNull(rsUser!部门ID), 0, rsUser!部门ID)
        UserInfo.简码 = IIf(IsNull(rsUser!简码), "", rsUser!简码)
        UserInfo.姓名 = IIf(IsNull(rsUser!姓名), "", rsUser!姓名)
        UserInfo.部门 = rsUser!部门名称
        UserInfo.用户名 = rsUser!用户名
        UserInfo.站点 = rsUser!用户名
    End If
    GetUserInfo = True
    Exit Function
errH:
    MsgBox "获取操作员信息时出错!", vbInformation, gstrSysName
    Resume
End Function

Public Function Currentdate() As Date
'功能：获取当前系统时间
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
    
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open "Select Sysdate From Dual", gcnOracle, adOpenKeyset
    Currentdate = rsTmp.fields(0).Value
    Exit Function
errH:
    MsgBox "获取系统日期时发生错误!", vbInformation, gstrSysName
End Function

Public Function UserIsOwner() As Boolean
'功能：判断用户是否所有者
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
        
    strSql = "Select * From zlSystems Where 编号 Like '1__' And Upper(所有者)=USER"
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, gcnOracle, adOpenKeyset
    UserIsOwner = Not rsTmp.EOF
End Function

Public Function OraDataOpen(ByVal gcnDatabase As ADODB.Connection, ByVal strServerName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
'功能：连接Oracle
    Dim rsTmp As New ADODB.Recordset
    Dim rsUser As ADODB.Recordset
    Dim strSql As String, i As Integer
    
    On Error Resume Next
    DoEvents
    With gcnDatabase
        If .State = adStateOpen Then .Close
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
        If Err <> 0 Then
            MsgBox "连接失败！（请确保用户名与密码输入无误）", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    OraDataOpen = True
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

Public Function InDesign() As Boolean
    On Error Resume Next
    Debug.Print 1 / 0
    If Err.Number <> 0 Then Err.Clear: InDesign = True
End Function

Public Function Custom_WndMessage(ByVal hwnd As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
'功能：自定义消息函数处理窗体尺寸调整限制
    If msg = WM_GETMINMAXINFO Then
        Dim MinMax As MINMAXINFO
        CopyMemory MinMax, ByVal lp, Len(MinMax)
        MinMax.ptMinTrackSize.X = glngFormW \ Screen.TwipsPerPixelX
        MinMax.ptMinTrackSize.Y = glngFormH \ Screen.TwipsPerPixelY
        MinMax.ptMaxTrackSize.X = glngFormW \ Screen.TwipsPerPixelX
        MinMax.ptMaxTrackSize.Y = glngFormH \ Screen.TwipsPerPixelY
        CopyMemory ByVal lp, MinMax, Len(MinMax)
        Custom_WndMessage = 1
        Exit Function
    End If
    Custom_WndMessage = CallWindowProc(glngOld, hwnd, msg, wp, lp)
End Function

Public Function Nvl(ByVal varValue As Variant, Optional varDefalut As Variant = "") As Variant
'功能：模仿Oracle的函数
    Nvl = IIf(IsNull(varValue) = True, varDefalut, varValue)
End Function

Public Sub OpenRecordset(rsTemp As ADODB.Recordset, ByVal strCaption As String, Optional strSql As String = "")
'功能：打开记录集
    If rsTemp.State = adStateOpen Then rsTemp.Close
    
    rsTemp.Open IIf(strSql = "", gstrSQL, strSql), gcnOracle, adOpenStatic, adLockReadOnly
End Sub

Public Function ToVarchar(ByVal varText As Variant, ByVal lngLength As Long) As String
'功能：将文本按Varchar2的长度计算方法进行截断
    Dim strText As String
    
    strText = IIf(IsNull(varText), "", varText)
    ToVarchar = StrConv(LeftB(StrConv(strText, vbFromUnicode), lngLength), vbUnicode)
    '去掉可能出现的半个字符
    ToVarchar = Replace(ToVarchar, Chr(0), "")
End Function

Public Sub ExecuteProcedure(ByVal strCaption As String)
'功能：执行SQL语句
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
End Sub

Public Function TruncZero(ByVal strInput As String) As String
'功能：去掉字符串中\0以后的字符
    Dim lngPos As Long
    
    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZero = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZero = strInput
    End If
End Function

Public Function OpenDatabase() As Boolean
    Dim strServer As String, strUser As String, strPass As String, strtemp As String
    Dim rsTemp As New ADODB.Recordset
    '首先读出参数，打开连接
    gstrSQL = "Select 参数名,参数值 From 保险参数 Where 险类=" & gintInsure
    Call OpenRecordset(rsTemp, "获取保险参数")
    Do Until rsTemp.EOF
        strtemp = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
        Select Case rsTemp("参数名")
            Case "医保服务器"
                strServer = strtemp
            Case "医保用户名"
                strUser = strtemp
            Case "医保用户密码"
                strPass = strtemp
        End Select
        rsTemp.MoveNext
    Loop
    If OraDataOpen(gcnYB, strServer, strUser, strPass) = False Then
        Exit Function
    End If
    OpenDatabase = True
End Function
