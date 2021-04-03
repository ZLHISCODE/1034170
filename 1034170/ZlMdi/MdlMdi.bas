Attribute VB_Name = "MdlMdi"
Option Explicit
'--菜单函数--
Public Declare Function CreateMenu Lib "user32" () As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long

'返回窗体的菜单句柄
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
'返回指定位置的弹出菜单的句柄
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
'获取指定菜单的句柄(弹出菜单返回-1;分隔菜单返回0)
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
'获取指定菜单的菜单项数
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
'取指定菜单项的字串
Public Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
'参数 类型及说明
'hMenu Long，                   菜单的句柄
'nPosition Long，               定义了新条目插入点的一个现有菜单条目的标志符。如果在wFlags中指定了MF_BYCOMMAND标志，
'                           这个参数就代表欲改变的菜单条目的命令ID。如设置的是MF_BYPOSITION标志，这个参数就代表
'                           菜单条目在菜单中的位置，第一个条目的位置为零
'wFlags Long，                  一系列常数标志的组合。参考ModifyMenu
'wIDNewItem Long，              指定菜单条目的新菜单ID。如果在wFlags中指定了MF_POPUP标志，就应该指定弹出式菜单的一个句柄
'lpNewItem                      如果在wFlags参数中设置了MF_STRING标志，就代表要设置到菜单中的字串（String）。
'                           如设置的是MF_BITMAP标志，就代表一个Long型变量，其中包含了一个位图句柄
'常数列表
Public Const MF_BYPOSITION = &H400&
Public Const MF_STRING = &H0&               '在指定的条目处放置一个字串。不与vb的caption属性兼容
Public Const MF_POPUP = &H10&               '将一个弹出式菜单置于指定的条目.可用于创建子菜单及弹出式菜单
Public Const MF_SEPARATOR = &H800&          '在指定的条目处显示一条分隔线
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
'为指定窗体设置新的菜单
Public Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const GWL_WNDPROC = (-4)
Public Const WM_COMMAND = &H111

Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CascadeWindows% Lib "user32" (ByVal hwndParent As Long, ByVal wHow As Long, lpRect As RECT, ByVal cKids As Long, lpKids As Long)
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Type POINTAPI
     x As Long
     y As Long
End Type
Public Type MINMAXINFO
        ptReserved As POINTAPI
        ptMaxSize As POINTAPI
        ptMaxPosition As POINTAPI
        ptMinTrackSize As POINTAPI
        ptMaxTrackSize As POINTAPI
End Type
Public Const WM_GETMINMAXINFO = &H24

'公用
Public LngAddFunc As Long
Public CollMenu As New Collection                       '菜单集合
Public CollOpenWindowHdl As New Collection              '已运行的窗体句柄
Public Const Menu_Hdl As Integer = 0                    '菜单句柄
Public Const Menu_Code As Integer = 1                   '菜单编号
Public Const Menu_Modul As Integer = 2                  '菜单模块
Public Const Menu_Component As Integer = 3              '对应部件名称
Public Const Menu_UpperHdl As Integer = 4               '其上级菜单句柄
Public Const Menu_Caption As Integer = 5                '标题及快捷键
Public Const Menu_ID As Integer = 6                     '菜单ID
Public Const Menu_Sys As Integer = 7                    '系统编号

Public gLngMinH As Double
Public gLngMinW As Double
Public gLngMaxH As Double
Public gLngMaxW As Double

'Command命令相关
Public Const INFINITE As Long = &HFFFF&
Private Const SW_HIDE As Integer = 0 '隐藏窗口，激活另一个窗口
Private Const NORMAL_PRIORITY_CLASS As Long = &H20&
Public Const STARTF_USESTDHANDLES = &H100&
Public Const STARTF_USESHOWWINDOW = &H1
Private Const ERROR_EXTENDED_ERROR          As Long = 1208
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200

Public Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadID As Long
End Type
Public Type STARTUPINFO
    Cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
'注册表安全属性类型
Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

Public Declare Function CreatePipe Lib "kernel32" (phReadPipe As Long, phWritePipe As Long, lpPipeAttributes As SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
Public Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function WNetGetLastError Lib "mpr.dll" Alias "WNetGetLastErrorA" (lpError As Long, ByVal lpErrorBuf As String, ByVal nErrorBufSize As Long, ByVal lpNameBuf As String, ByVal nNameBufSize As Long) As Long


Public Function MenuProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim LngFind As Long, BlnFind As Boolean, BlnRun As Boolean
    Dim StrComponent As String, lngModul As Long, StrCaption As String, lngSys As Long
    Dim LngTargetHdl As Long
    
    '处理菜单事件
    If uMsg = WM_COMMAND Then
        '查找对应的集合
        If wParam >= 菜单基准.其它功能菜单 - 1 Then
            Select Case wParam
            Case 99999901   '启动自定义报表
                Call ExecuteFunc(0, "ZL9REPORT", 99999901)
            End Select
        ElseIf wParam >= 菜单基准.窗口菜单 - 1 Then '窗口列表菜单
            For LngFind = 0 To CollOpenWindowHdl.Count - 1
                If wParam = CollOpenWindowHdl("K_" & LngFind)(2) Then
                    LngTargetHdl = CollOpenWindowHdl("K_" & LngFind)(0)
                    Exit For
                End If
            Next
            
            If IsIconic(LngTargetHdl) Then
                Call ShowWindow(LngTargetHdl, 9)    '还原指定窗体为原大小
            End If
            Call SetActiveWindow(LngTargetHdl)
            MenuProc = 1
        ElseIf wParam > 菜单基准.功能菜单 - 1 Then  '程序模块列表菜单
            BlnFind = False
            For LngFind = 0 To CollMenu.Count - 1
                If CollMenu("K_" & LngFind)(Menu_ID) = wParam Then
                    BlnFind = True
                    StrCaption = CollMenu("K_" & LngFind)(Menu_Caption)
                    If InStr(1, StrCaption, "(") <> 0 And InStr(1, StrCaption, ")") <> 0 Then StrCaption = Mid(StrCaption, 1, InStr(1, StrCaption, "("))
                    StrComponent = CollMenu("K_" & LngFind)(Menu_Component)
                    lngSys = CollMenu("K_" & LngFind)(Menu_Sys)
                    lngModul = CollMenu("K_" & LngFind)(Menu_Modul)
                    Exit For
                End If
            Next
            '找到则执行
            If BlnFind Then
                Call AddHistory(lngSys & "," & lngModul)
                Call frmMdi.LoadHistory
                
                '查找该模块是否已运行,是则设为活动窗体
                BlnRun = False
                For LngFind = 0 To CollOpenWindowHdl.Count - 1
                    If StrCaption = CollOpenWindowHdl("K_" & LngFind)(1) Then
                        BlnRun = True
                        LngTargetHdl = CollOpenWindowHdl("K_" & LngFind)(0)
                        Exit For
                    End If
                Next
                If BlnRun Then
                    If IsIconic(LngTargetHdl) Then
                        Call ShowWindow(LngTargetHdl, 9)            '还原指定窗体为原大小
                    End If
                    Call SetActiveWindow(LngTargetHdl)
                Else
                    ExecuteFunc lngSys, StrComponent, lngModul
                End If
                MenuProc = 1
            Else
                MenuProc = CallWindowProc(LngAddFunc, FrmMainface.hwnd, uMsg, wParam, lParam)
            End If
        Else
            MenuProc = CallWindowProc(LngAddFunc, FrmMainface.hwnd, uMsg, wParam, lParam)
        End If
    ElseIf uMsg = WM_GETMINMAXINFO Then
        Dim MinMax As MINMAXINFO
        CopyMemory MinMax, ByVal lParam, Len(MinMax)
        MinMax.ptMinTrackSize.x = gLngMinW \ 15
        MinMax.ptMinTrackSize.y = gLngMinH \ 15
        MinMax.ptMaxTrackSize.x = gLngMaxW \ 15
        MinMax.ptMaxTrackSize.y = gLngMaxH \ 15
        CopyMemory ByVal lParam, MinMax, Len(MinMax)
        MenuProc = 1
    Else
        MenuProc = CallWindowProc(LngAddFunc, FrmMainface.hwnd, uMsg, wParam, lParam)
    End If
End Function


Public Function RunCommand(ByVal strCommand As String, Optional ByRef strErr As String, Optional ByVal blnCiper As Boolean, Optional ByVal lngWait As Long = INFINITE) As String
'功能：执行命令行，并获取命令行输出
    Dim piProc          As PROCESS_INFORMATION '进程信息
    Dim stStart         As STARTUPINFO '启动信息
    Dim saSecAttr       As SECURITY_ATTRIBUTES '安全属性
    Dim lnghReadPipe    As Long '读取管道句柄
    Dim lnghWritePipe   As Long '写入管道句柄
    Dim lngBytesRead    As Long '读出数据的字节数
    Dim strBuffer       As String * 256 '读取管道的字符串buffer
    Dim lngRet          As Long 'API函数返回值
    Dim lngRetPro       As Long
    Dim strlpOutputs    As String '读出的最终结果
    
    DoEvents
    On Error Resume Next
    '设置安全属性
    With saSecAttr
        .nLength = LenB(saSecAttr)
        .bInheritHandle = True
        .lpSecurityDescriptor = 0
    End With
    
    '创建管道
    lngRet = CreatePipe(lnghReadPipe, lnghWritePipe, saSecAttr, 0)
    If lngRet = 0 Then
        strErr = "无法创建管道。" & GetLastDllErr()
        Exit Function
    End If
    '设置进程启动前的信息
    With stStart
        .Cb = LenB(stStart)
        .dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
        .wShowWindow = SW_HIDE
        .hStdOutput = lnghWritePipe '设置输出管道
        .hStdError = lnghWritePipe '设置错误管道
    End With
    '启动进程
    'Command = "c:\windows\system32\ipconfig.exe /all" 'DOS进程以ipconfig.exe为例
    lngRetPro = CreateProcess(vbNullString, strCommand & vbNullChar, saSecAttr, saSecAttr, 1&, NORMAL_PRIORITY_CLASS, ByVal 0&, vbNullString, stStart, piProc)
    If lngRetPro = 0 Then
        strErr = "无法启动进程。" & GetLastDllErr()
        lngRet = CloseHandle(lnghWritePipe)
        lngRet = CloseHandle(lnghReadPipe)
        Exit Function
    Else
        '因为无需写入数据，所以先关闭写入管道。而且这里必须关闭此管道，否则将无法读取数据
        lngRet = CloseHandle(lnghWritePipe)
        WaitForSingleObject piProc.hProcess, lngWait
        Do
            lngRet = ReadFile(lnghReadPipe, strBuffer, 256, lngBytesRead, ByVal 0)
            If lngRet <> 0 Then
                strlpOutputs = strlpOutputs & Left(strBuffer, lngBytesRead)
            Else
                strlpOutputs = strlpOutputs & Left(strBuffer, lngBytesRead)
            End If
            DoEvents
        Loop While (lngRet <> 0) '当ret=0时说明ReadFile执行失败，已经没有数据可读了
        '读取操作完成，关闭各句柄
        lngRet = CloseHandle(lngRetPro)
        lngRet = CloseHandle(piProc.hProcess)
        lngRet = CloseHandle(piProc.hThread)
        lngRet = CloseHandle(lnghReadPipe)
    End If
    RunCommand = Replace(strlpOutputs, vbNullChar, "")
End Function

Public Function GetLastDllErr(Optional ByVal LngErr As Long) As String
    Dim strReturn As String
    If LngErr = 0 Then
        LngErr = GetLastError
    End If
    If LngErr = ERROR_EXTENDED_ERROR Then
        GetLastDllErr = GetWNetErr(LngErr)
    Else
        strReturn = String$(256, 32)
        FormatMessage FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0&, LngErr, 0&, strReturn, Len(strReturn), ByVal 0
        strReturn = Trim(strReturn)
        GetLastDllErr = Replace(Replace(strReturn, Chr(10), ""), Chr(13), "")
    End If
End Function

Private Function GetWNetErr(ByVal LngErr As Long) As String
    Dim strErr As String * 256
    Dim strName As String * 256
    Dim lngRet As Long
    lngRet = WNetGetLastError(LngErr, strErr, Len(strErr), strName, Len(strName))
    GetWNetErr = Replace(Replace("[" & TruncZero(strName) & "]" & TruncZero(strErr), Chr(10), ""), Chr(13), "")
End Function

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





