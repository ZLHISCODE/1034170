Attribute VB_Name = "mdlMain"
Option Explicit
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Type BrowseInfo
   hwndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type
Public Type POINTAPI
        x As Long
        y As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260
Public Const SWP_NOSIZE = &H1
Public Const SWP_SHOWWINDOW = &H40
Public Const HWND_TOPMOST = -1
Public Const GWL_EXSTYLE = (-20)
Public Const WinStyle = &H40000
Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2 '浅凹下
Public Const BDR_RAISEDINNER = &H4 '浅凸起
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER) '深凸起
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER) '深凹下
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER) 'Frame边线样式
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER) '反Frame边线样式
'---------------------------------------------------------------
'- 注册表 Api 常数...
'---------------------------------------------------------------
' Reg Data Types...
Const REG_SZ = 1                         ' Unicode空终结字符串
Const REG_EXPAND_SZ = 2                  ' Unicode空终结字符串
Const REG_DWORD = 4                      ' 32-bit 数字

' 注册表创建类型值...
Const REG_OPTION_NON_VOLATILE = 0       ' 当系统重新启动时，关键字被保留

' 注册表关键字安全选项...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Const KEY_EXECUTE = KEY_READ
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' 注册表关键字根类型...
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004

' 返回值...
Const ERROR_NONE = 0
Const ERROR_BADKEY = 2
Const ERROR_ACCESS_DENIED = 8
Const ERROR_SUCCESS = 0

'---------------------------------------------------------------
'- 注册表安全属性类型...
'---------------------------------------------------------------
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function Htmlhelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Any) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Public gcnOracle As New ADODB.Connection    '公共数据库连接
Public gcnAcc As New ADODB.Connection

Public gobjFunction As Object
Public gobjReport As Object

Public gstrProductTitle As String
Public gstrProductName As String
Public gstrDevelopers As String
Public gstrSustainer As String
Public gstrWebSustainer As String
Public gstrWebURL As String
Public gstrWebEmail As String
Public gstrSysName As String                '系统名称
Public gstrUserName As String               '用户名
Public gstrPassword As String               '用户口令
Public gstrToolsPwd As String               '管理工具的密码
Public gstrServer As String                 '服务器名
Public gstrSQL    As String                 '通用的SQL语句变量
Public gblnDBA As Boolean                   '是否DBA
Public gblnOwner As Boolean                 '是否所有者
Public gdtStart As Long
Public gstrDBUser As String

Public Function ShowHelp(SHwnd As Long, ByVal htmName As String) As Boolean
    '显示帮助窗体
    'SHwnd:传入窗口句柄(作为宿主窗口)
    'htmName:射映在CHM中的htm文件名称

    Dim Path As String
    Dim strSave As String
    On Error GoTo ShowHelpErr
    
    ShowHelp = False
    strSave = String(200, Chr$(0))
    Path = Left$(strSave, GetWindowsDirectory(strSave, Len(strSave))) + "\help\"
    If CBool(PathIsDirectory(Path)) = False Then GoTo ShowHelpErr
    strSave = "zlSDK.CHM"
    Path = Trim(Path & strSave)
    If Trim(Dir(Path)) = "" Then GoTo ShowHelpErr
    Call Htmlhelp(SHwnd, Path, &H0, htmName & ".htm")
    ShowHelp = True
    Exit Function
ShowHelpErr:
    Err.Clear
End Function

Public Sub Main()
    Dim StrUnitName As String, intCount As Integer
    '为实现XP风格，在显示窗体前必须执行该函数
    Call InitCommonControls

    '为了实现注销功能，对全局变量进行初始化
    gblnDBA = False
    gblnOwner = False
    Set gobjFunction = Nothing
    Set gobjReport = Nothing
    
    
    Load frmSplash
    '由注册表中获取用户注册相关信息,如果用户单位名称不为空,则显示闪现窗体
    StrUnitName = GetSetting("ZLSOFT", "注册信息", "单位名称", "")
    gstrSysName = GetSetting("ZLSOFT", "注册信息", "提示", "")
    If StrUnitName <> "" And StrUnitName <> "-" Then
        gdtStart = Timer
        With frmSplash
            '有两处需要处理
            Call ApplyOEM_Picture(.ImgIndicate, "Picture")
            Call ApplyOEM_Picture(.imgPic, "PictureB")
            .Show
            .lblGrant = StrUnitName
            StrUnitName = GetSetting("ZLSOFT", "注册信息", "开发商", "")
            If Trim(StrUnitName) = "" Then
                .Label3.Visible = False
                .lbl开发商.Visible = False
            Else
                .lbl开发商.Caption = ""
                For intCount = 0 To UBound(Split(StrUnitName, ";"))
                    .lbl开发商.Caption = .lbl开发商.Caption & Split(StrUnitName, ";")(intCount) & vbCrLf
                Next
            End If
            .LblProductName = GetSetting("ZLSOFT", "注册信息", "产品全称", "")
            .lbl技术支持商 = GetSetting("ZLSOFT", "注册信息", "技术支持商", "")
        End With
        Do
            If (Timer - gdtStart) > 3 Then Exit Do
            DoEvents
        Loop
        
        DoEvents
    End If
    
    
    frmUserLogin.Show 1
    Unload frmSplash
    If gcnOracle.State = adStateOpen Then
        Call InitCommon(gcnOracle)
        frmMDIMain.Show
    End If
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
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim strError As String
    
    On Error Resume Next
    If gcnOracle.State = adStateOpen Then gcnOracle.Close
    Set gcnOracle = Nothing
    With gcnOracle
        .Provider = "MSDataShape"
        .Open "Driver={Microsoft ODBC for Oracle};Server=" & strServerName, strUserName, strUserPwd
    End With
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
        ElseIf InStr(strError, "ORA-01034") > 0 Then
            MsgBox "ORACLE不可用，请检查服务或数据库实例是否启动。", vbInformation, gstrSysName
        ElseIf InStr(strError, "ORA-02391") > 0 Then
            MsgBox "用户" & UCase(strUserName) & "已经登录，不允许重复登录(已达到系统所允许的最大登录数)。", vbExclamation, gstrSysName
        ElseIf InStr(strError, "ORA-01017") > 0 Then
            MsgBox "由于用户、口令或服务器指定错误，无法注册。", vbInformation, gstrSysName
        Else
            MsgBox strError, vbInformation, gstrSysName
        End If
        
        OraDataOpen = False
        Exit Function
    End If
    
    Err = 0
    With rsTemp
        strSQL = "SELECT 1 FROM ZLSYSTEMS WHERE upper(所有者)=USER"
        If .State = adStateOpen Then .Close
        .Open strSQL, gcnOracle, adOpenKeyset
        If Err <> 0 Then
            gblnOwner = False
        Else
            gblnOwner = Not (.EOF Or .BOF)
        End If
        strSQL = "SELECT 1 FROM SESSION_ROLES WHERE ROLE='DBA'"
        If .State = adStateOpen Then .Close
        .Open strSQL, gcnOracle, adOpenKeyset
        gblnDBA = Not (.EOF Or .BOF)
    End With
    
    If Not (gblnDBA) And Not (gblnOwner) Then
        OraDataOpen = False
        MsgBox "不是数据库DBA或应用系统的所有者，不能使用该工具。", vbExclamation, gstrSysName
        Exit Function
    End If
    
    OraDataOpen = True
    gstrUserName = strUserName
    gstrPassword = strUserPwd
    gstrDBUser = UCase(strUserName)
    gstrServer = Trim(strServerName)
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


Public Sub SelAll(objTxt As Control)
'功能：对文本框的的文本选中
    If TypeName(objTxt) = "TextBox" Then
        objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
    ElseIf TypeName(objTxt) = "MaskEdBox" Then
        If Not IsDate(objTxt.Text) Then
            objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
        Else
            objTxt.SelStart = 0: objTxt.SelLength = 10
        End If
    End If
End Sub

Public Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0) As Boolean
'检查字符串是否含有非法字符；如果提供长度，对长度的合法性也作检测。
    If InStr(strInput, "'") > 0 Or InStr(strInput, """") > 0 Then
        MsgBox "所输入内容含有非法字符。", vbExclamation, gstrSysName
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox "所输入内容不能超过" & Int(intMax / 2) & "个汉字" & "或" & intMax & "个字母。", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    StrIsValid = True
End Function

Public Function Currentdate() As Date
    '-------------------------------------------------------------
    '功能：提取服务器上当前日期
    '参数：
    '返回：由于Oracle日期格式的问题，所以
    '-------------------------------------------------------------
    
    Dim rsTemp As New ADODB.Recordset
    
    Err = 0
    On Error GoTo ErrHand
    With rsTemp
        .Open "SELECT SYSDATE FROM DUAL", gcnOracle, adOpenKeyset
    End With
    Currentdate = rsTemp.Fields(0).Value
    rsTemp.Close
    Exit Function
    
ErrHand:
    Currentdate = Date
    Err = 0
End Function


'将PictureBox模拟成3D平面按钮
'intStyle=0=平面,-1=凹下,1=凸起,-2=深凹下,2=深凸起
Public Sub RaisEffect(picBox As PictureBox, Optional IntStyle As Integer, Optional strName As String = "")
    Dim PicRect As RECT
    Dim lngTmp As Long
    With picBox
        lngTmp = .ScaleMode
        .ScaleMode = 3
        .Cls
        .BorderStyle = 0
        
        If IntStyle <> 0 Then
            PicRect.Left = .ScaleLeft
            PicRect.Top = .ScaleTop
            PicRect.Right = .ScaleWidth
            PicRect.Bottom = .ScaleHeight
            
            Select Case IntStyle
                Case 1
                    DrawEdge .hDC, PicRect, CLng(BDR_RAISEDINNER), BF_RECT
                Case 2
                    DrawEdge .hDC, PicRect, CLng(EDGE_RAISED), BF_RECT
                Case -1
                    DrawEdge .hDC, PicRect, CLng(BDR_SUNKENOUTER), BF_RECT
                Case -2
                    DrawEdge .hDC, PicRect, CLng(EDGE_SUNKEN), BF_RECT
            End Select
        End If
        .ScaleMode = lngTmp
        If strName <> "" Then
            .CurrentX = (.ScaleWidth - .TextWidth(strName)) / 2
            .CurrentY = (.ScaleHeight - .TextHeight(strName)) / 2
            picBox.Print strName
        End If
    End With
End Sub



Public Function GetOwnerName(lngSys As Long, cnLink As ADODB.Connection) As String
    Dim rstmp As New ADODB.Recordset
    
    Set rstmp = New ADODB.Recordset
    rstmp.CursorLocation = adUseClient
    rstmp.Open "Select 所有者 From zlSystems Where 编号=" & lngSys, cnLink, adOpenKeyset
    If Not rstmp.EOF Then GetOwnerName = rstmp!所有者
End Function

Public Function GetMaxID(strTable As String, cnLink As ADODB.Connection) As Long
    Dim rstmp As New ADODB.Recordset
    
    Set rstmp = New ADODB.Recordset
    rstmp.CursorLocation = adUseClient
    rstmp.Open "Select Nvl(Max(ID),0) as ID From " & strTable, cnLink, adOpenKeyset
    If Not rstmp.EOF Then GetMaxID = rstmp!Id
End Function


Public Function LoadCustomPicture(strID As String) As StdPicture
'功能:将资源文件中的指定资源生成磁盘文件
'参数:ID=资源号,strExt=要生成文件的扩展名(如BMP)
'返回:生成文件名
    Dim arrData() As Byte
    Dim intFile As Integer
    Dim strFile As String * 255, strR As String
    
    arrData = LoadResData(strID, "CUSTOM")
    intFile = FreeFile
    
    GetTempPath 255, strFile
    strR = Trim(Left(strFile, InStr(strFile, Chr(0)) - 1)) & CLng(Timer * 100) & ".pic"

    Open strR For Binary As intFile
    Put intFile, , arrData()
    Close intFile
    Set LoadCustomPicture = VB.LoadPicture(strR)
    Kill strR
End Function

Public Function GetOEM(ByVal strAsk As String, Optional ByVal blnCorp As Boolean = True) As String
    '-------------------------------------------------------------
    '功能：返回每个字线的ASCII码
    '参数：
    '返回：
    '-------------------------------------------------------------
    Dim intBit As Integer, iCount As Integer, blnCan As Boolean
    Dim strCode As String
    
    'OEM图片有两种类型 ，一是指公司徽标，另一个是产品标识
    strCode = IIf(blnCorp = True, "OEM_", "PIC_")
    For intBit = 1 To Len(strAsk)
        '取每个字的ASCII码
        strCode = strCode & Hex(Asc(Mid(strAsk, intBit, 1)))
    Next
    GetOEM = strCode
End Function

Public Function OpenDire(odtvOwner As Form, Optional odtvTitle As String) As String
   Dim lpIDList As Long
   Dim sBuffer As String
   Dim szTitle As String
   Dim tBrowseInfo As BrowseInfo
   szTitle = odtvTitle
   With tBrowseInfo
      .hwndOwner = odtvOwner.hwnd
      .lpszTitle = lstrcat(szTitle, "")
      .ulFlags = BIF_RETURNONLYFSDIRS ' + BIF_DONTGOBELOWDOMAIN
   End With
   lpIDList = SHBrowseForFolder(tBrowseInfo)
   If (lpIDList) Then
      sBuffer = Space(MAX_PATH)
      SHGetPathFromIDList lpIDList, sBuffer
      sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
      OpenDire = sBuffer
   End If
End Function

Public Sub ReCompileProcedure(ByVal cnOwner As ADODB.Connection)
    '对本用户下所有已经失效的过程进行重新编译
    Dim rsTemp As New ADODB.Recordset
    Dim lngTime As Long
    
    For lngTime = 1 To 3
        '最多调用三次，因为有些过程是相互调用，一次编译不能解决问题
        '为了快速得到列表，不利用对象之间的引用关系
        If rsTemp.State = adStateOpen Then rsTemp.Close
        
        gstrSQL = "select OBJECT_NAME from user_objects where object_type='PROCEDURE' and STATUS='INVALID'"
        rsTemp.Open gstrSQL, cnOwner, adOpenStatic, adLockReadOnly
        
        On Error Resume Next
        If rsTemp.RecordCount = 0 Then
            '没有过程失效，直接退出
            Exit Sub
        Else
            Do Until rsTemp.EOF
                '有可能出错
                gstrSQL = "alter procedure " & rsTemp("OBJECT_NAME") & " compile"
                cnOwner.Execute gstrSQL
                rsTemp.MoveNext
            Loop
        End If
    Next
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String) As String
'功能：读注册表
    Dim i As Long                                           ' 循环计数器
    Dim rc As Long                                          ' 返回代码
    Dim hKey As Long                                        ' 处理打开的注册表关键字
    Dim hDepth As Long                                      '
    Dim sKeyVal As String
    Dim lKeyValType As Long                                 ' 注册表关键字数据类型
    Dim tmpVal As String                                    ' 注册表关键字的临时存储器
    Dim KeyValSize As Long                                  ' 注册表关键字变量尺寸
    
    ' 在 KeyRoot {HKEY_LOCAL_MACHINE...} 下打开注册表关键字
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' 打开注册表关键字
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 处理错误...
    
    tmpVal = String$(1024, 0)                             ' 分配变量空间
    KeyValSize = 1024                                       ' 标记变量尺寸
    
    '------------------------------------------------------------
    ' 检索注册表关键字的值...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         lKeyValType, tmpVal, KeyValSize)    ' 获得/创建关键字的值
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' 错误处理
      
    tmpVal = Left$(tmpVal, InStr(tmpVal, Chr(0)) - 1)

    '------------------------------------------------------------
    ' 决定关键字值的转换类型...
    '------------------------------------------------------------
    Select Case lKeyValType                                  ' 搜索数据类型...
    Case REG_SZ, REG_EXPAND_SZ                              ' 字符串注册表关键字数据类型
        sKeyVal = tmpVal                                     ' 复制字符串的值
    Case REG_DWORD                                          ' 四字节注册表关键字数据类型
        For i = Len(tmpVal) To 1 Step -1                    ' 转换每一位
            sKeyVal = sKeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' 一个字符一个字符地生成值。
        Next
        sKeyVal = Format$("&h" + sKeyVal)                     ' 转换四字节为字符串
    End Select
    
    GetKeyValue = sKeyVal                                   ' 返回值
    rc = RegCloseKey(hKey)                                  ' 关闭注册表关键字
    Exit Function                                           ' 退出
    
GetKeyError:    ' 错误发生过后进行清除...
    GetKeyValue = vbNullString                              ' 设置返回值为空字符串
    rc = RegCloseKey(hKey)                                  ' 关闭注册表关键字
End Function

Public Function CheckSpaceIsUse(ByVal strType As String, ByVal strName As String, ByVal strOwner As String) As Boolean
'功能：检查表空间或数据文件是否由其它用户使用
'参数：strType    表空间 数据文件
'      strName          表空间或数据文件的名字
'      strOwner         以区别其它用户的所有者名
    Dim rsTemp As New ADODB.Recordset
    
    If strType = "表空间" Then
        gstrSQL = "select owner from all_tables where tablespace_name='" & UCase(strName) & "' and owner<>'" & UCase(strOwner) & "' AND ROWNUM<2"
    Else
        gstrSQL = "select O.owner  from V$TABLESPACE T,V$DATAFILE F,all_tables O " & _
                  "Where T.TS# = F.TS# And T.name = O.TABLESPACE_NAME " & _
                  "    and F.name='" & UCase(strName) & "' and O.owner<>'" & UCase(strOwner) & "' AND ROWNUM<2 "
    End If
    
    On Error Resume Next
    rsTemp.Open gstrSQL, gcnOracle, , adLockReadOnly
    
    If rsTemp.RecordCount = 0 Then
        '没有其他用户使用，可以删除
        Exit Function
    End If
    '有用户使用
    CheckSpaceIsUse = True
End Function

Public Function GetVerDouble(ByVal varVer As Variant) As Double
'功能：根据版本字符串，产生数字化的版本
'参数：varVer   版本字符串，如9.5.0
    Dim varArray As Variant
    
    varVer = IIf(IsNull(varVer), "", varVer)
    varArray = Split(varVer, ".")
    If UBound(varArray) < 2 Then Exit Function
    
    GetVerDouble = Val(varArray(0)) * 10 ^ 8 + Val(varArray(1)) * 10 ^ 4 + Val(varArray(2))
End Function

Public Function GetVerString(ByVal dblVer As Double) As String
'功能：根据数字化的版本，产生版本字符串
'参数：dblVer   版本字符串，如900050000
    
    GetVerString = dblVer \ 10 ^ 8 & "." & (dblVer Mod 10 ^ 8) \ 10 ^ 4 & "." & dblVer Mod 10 ^ 4
End Function

Private Function JudgeManagerVer() As Double
'功能：判断管理工具的版本
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select 编号 from zlSvrTools where 编号='0502'"
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    If rsTemp.EOF = True Then
        '那是最早的，版本为9.0.0
        JudgeManagerVer = 9 * 10 ^ 8
        Exit Function
    End If
    rsTemp.Close
    
    gstrSQL = "SELECT CONSTRAINT_NAME FROM All_Constraints C WHERE C.CONSTRAINT_NAME='ZLOPTIONS_PK' AND C.OWNER='ZLTOOLS' AND C.TABLE_NAME='ZLOPTIONS'"
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    If rsTemp.EOF = True Then
        '如果不存在ZLOPTIONS_PK约束，说明没有执行第二个升级脚本，版本为9.1.0
        JudgeManagerVer = 9 * 10 ^ 8 + 1 * 10 ^ 4
        Exit Function
    End If
    rsTemp.Close
    
    gstrSQL = "SELECT CONSTRAINT_NAME FROM All_Constraints C WHERE C.CONSTRAINT_NAME='ZLXLSVERIFY_FK_报表号' AND C.OWNER='ZLTOOLS' AND C.TABLE_NAME='ZLXLSVERIFY'"
    rsTemp.Open gstrSQL, gcnOracle, adOpenStatic, adLockReadOnly
    If rsTemp.EOF = False Then
        '如果存在ZLXLSVERIFY_FK_报表号约束，说明没有执行第三个升级脚本，版本为9.2.0
        JudgeManagerVer = 9 * 10 ^ 8 + 2 * 10 ^ 4
        Exit Function
    End If
    
    JudgeManagerVer = 9 * 10 ^ 8 + 3 * 10 ^ 4
End Function

Public Function LvwSelectColumns(objSet As Object, ByVal strColumn As String, Optional ByVal blnInit As Boolean = False) As Boolean
'功能:对列表控件的列进行设置
'参数:
'   objSet：要设置的对象,目前只支持ListView，以后再加上FlexGrid,DataGrid。
'   strColumn；列串。格式是"列名,列宽,对齐数值,列特性;列名,列宽,对齐数值,列特性"    注意列之间是用分号
'      比如 "名称,2000,0,1;编码,800,0,0;简码,1440,0,0"
'      对ListView而言：列特性为1表示该列不可删除，列特性为0表示该列可以删除
'      对FlexGridView而言：列特性还要表示是否属于固定列，以便不能和其它列进行顺序调整
'   blnInit：True,不显示选择窗口，直接初始化
    Dim varColumns As Variant, varColumn As Variant
    Dim lngCol As Long

    If blnInit Then
        varColumns = Split(strColumn, ";")
        Select Case TypeName(objSet)
            Case "ListView"
                With objSet.ColumnHeaders
                    .Clear
                    For lngCol = LBound(varColumns) To UBound(varColumns)
                        varColumn = Split(varColumns(lngCol), ",")
                        .Add , "_" & varColumn(0), varColumn(0), varColumn(1), varColumn(2)
                    Next
                End With
            Case "MSHFlexGrid"
            Case "DataGrid"
        End Select
    End If
End Function

Public Sub NextLvwPos(lvwObj As Object, ByVal vIndex As Long)
        
    If lvwObj.ListItems.Count > 0 Then
        vIndex = IIf(lvwObj.ListItems.Count > vIndex, vIndex, lvwObj.ListItems.Count)
        lvwObj.ListItems(vIndex).Selected = True
        lvwObj.ListItems(vIndex).EnsureVisible
    End If
End Sub

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function
Public Sub OpenFolder(ByVal frmodtvOwner As Form, ByRef strFolderName As String, Optional strTitle As String)
    '----------------------------------------------------------------------------------------------------
    '功能:选择文件夹
    '参数:frmodtvOwner-选择文件夹的父窗体
    '     strFolderName-指定的文件夹
    '     strTitle-标题
    '返回:strFolderName-返回选择的文件夹
    '----------------------------------------------------------------------------------------------------
    
   Dim lpIDList As Long
   Dim sBuffer As String
   Dim szTitle As String
   Dim tBrowseInfo As BrowseInfo
   szTitle = strTitle
   With tBrowseInfo
      .hwndOwner = frmodtvOwner.hwnd
      .lpszTitle = lstrcat(szTitle, "")
      .ulFlags = BIF_RETURNONLYFSDIRS ' + BIF_DONTGOBELOWDOMAIN
   End With
   lpIDList = SHBrowseForFolder(tBrowseInfo)
   If (lpIDList) Then
      sBuffer = Space(MAX_PATH)
      SHGetPathFromIDList lpIDList, sBuffer
      sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
      strFolderName = sBuffer
   End If
End Sub

Public Sub OpenAccessRecordset(rsTemp As ADODB.Recordset, strSQL As String, ByVal strFormCaption As String, _
        Optional CursorType As CursorTypeEnum = adOpenStatic, Optional LockType As LockTypeEnum = adLockReadOnly)
    '功能：打开记录。同时保存SQL语句
    
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open strSQL, gcnAcc, CursorType, LockType
End Sub




Public Function Lpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:按指定长度填制空格
    '--入参数:
    '--出参数:
    '--返  回:返回字串
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp < lngLen Then
        strTmp = String(lngLen - lngTmp, strChar) & strTmp
    ElseIf lngTmp > lngLen Then  '大于长度时,自动载断
        strTmp = strCode
    End If
    Lpad = Replace(strTmp, Chr(0), strChar)
End Function
Public Function Rpad(ByVal strCode As String, lngLen As Long, Optional strChar As String = " ") As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:按指定长度填制空格
    '--入参数:
    '--出参数:
    '--返  回:返回字串
    '-----------------------------------------------------------------------------------------------------------
    Dim lngTmp As Long
    Dim strTmp As String
    strTmp = strCode
    lngTmp = LenB(StrConv(strCode, vbFromUnicode))
    If lngTmp < lngLen Then
        strTmp = strTmp & String(lngLen - lngTmp, strChar)
    Else
        '主要有空格引起的
        strTmp = strCode
    End If
    '取掉最后半个字符
    Rpad = Replace(strTmp, Chr(0), strChar)
End Function


Public Function AccDataOpen(ByVal strDatabaseName As String, ByVal strUserName As String, ByVal strUserPwd As String) As Boolean
    '------------------------------------------------
    '功能： 打开指定Access数据库
    '参数：
    '   strDataBaseName：数据库
    '   strUserName：用户名
    '   strUserPwd：密码
    '返回： 数据库打开成功，返回true；失败，返回false
    '------------------------------------------------
    Dim strSQL As String
    Dim sConnect As String
    Err = 0
    On Error GoTo ErrHand
    Set gcnAcc = New ADODB.Connection
    With gcnAcc
        If .State = adStateOpen Then .Close
        .Provider = "=Microsoft.Jet.OLEDB.4.0"
        sConnect = "Driver={Microsoft Access Driver (*.mdb)};Dbq="
        .Open sConnect & strDatabaseName, strUserName, strUserPwd
    End With
    AccDataOpen = True
    Exit Function
ErrHand:
    MsgBox "数据库打开失败", vbInformation, ""
    AccDataOpen = False
    Err = 0
End Function


Public Function Substr(ByVal strInfor As String, ByVal lngStart As Long, ByVal lngLen As Long) As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:读取指定字串的值,字串中可以包含汉字
    '--入参数:strInfor-原串
    '         lngStart-直始位置
    '         lngLen-长度
    '--出参数:
    '--返  回:子串
    '-----------------------------------------------------------------------------------------------------------
    Dim strTmp As String, i As Long
    
    Err = 0
    On Error GoTo ErrHand:

    Substr = StrConv(MidB(StrConv(strInfor, vbFromUnicode), lngStart, lngLen), vbUnicode)
    Substr = Replace(Substr, Chr(0), " ")
'    strTmp = Right(Substr, 1)
'    If zlCommFun.ActualLen(strTmp) = 1 Then
'        If asc(strTmp) < 32 Or asc(strTmp) > 126 Then
'            Substr = Left(Substr, Len(Substr) - 1)
'        End If
'    End If
    Exit Function
ErrHand:
    Substr = ""
End Function



