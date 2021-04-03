Attribute VB_Name = "MdlMdi"
Option Explicit
'--�˵�����--
Public Declare Function CreateMenu Lib "user32" () As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long

'���ش���Ĳ˵����
Public Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
'����ָ��λ�õĵ����˵��ľ��
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
'��ȡָ���˵��ľ��(�����˵�����-1;�ָ��˵�����0)
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
'��ȡָ���˵��Ĳ˵�����
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
'ȡָ���˵�����ִ�
Public Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
'���� ���ͼ�˵��
'hMenu Long��                   �˵��ľ��
'nPosition Long��               ����������Ŀ������һ�����в˵���Ŀ�ı�־���������wFlags��ָ����MF_BYCOMMAND��־��
'                           ��������ʹ������ı�Ĳ˵���Ŀ������ID�������õ���MF_BYPOSITION��־����������ʹ���
'                           �˵���Ŀ�ڲ˵��е�λ�ã���һ����Ŀ��λ��Ϊ��
'wFlags Long��                  һϵ�г�����־����ϡ��ο�ModifyMenu
'wIDNewItem Long��              ָ���˵���Ŀ���²˵�ID�������wFlags��ָ����MF_POPUP��־����Ӧ��ָ������ʽ�˵���һ�����
'lpNewItem                      �����wFlags������������MF_STRING��־���ʹ���Ҫ���õ��˵��е��ִ���String����
'                           �����õ���MF_BITMAP��־���ʹ���һ��Long�ͱ��������а�����һ��λͼ���
'�����б�
Public Const MF_BYPOSITION = &H400&
Public Const MF_STRING = &H0&               '��ָ������Ŀ������һ���ִ�������vb��caption���Լ���
Public Const MF_POPUP = &H10&               '��һ������ʽ�˵�����ָ������Ŀ.�����ڴ����Ӳ˵�������ʽ�˵�
Public Const MF_SEPARATOR = &H800&          '��ָ������Ŀ����ʾһ���ָ���
Public Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
'Ϊָ�����������µĲ˵�
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

'����
Public LngAddFunc As Long
Public CollMenu As New Collection                       '�˵�����
Public CollOpenWindowHdl As New Collection              '�����еĴ�����
Public Const Menu_Hdl As Integer = 0                    '�˵����
Public Const Menu_Code As Integer = 1                   '�˵����
Public Const Menu_Modul As Integer = 2                  '�˵�ģ��
Public Const Menu_Component As Integer = 3              '��Ӧ��������
Public Const Menu_UpperHdl As Integer = 4               '���ϼ��˵����
Public Const Menu_Caption As Integer = 5                '���⼰��ݼ�
Public Const Menu_ID As Integer = 6                     '�˵�ID
Public Const Menu_Sys As Integer = 7                    'ϵͳ���

Public gLngMinH As Double
Public gLngMinW As Double
Public gLngMaxH As Double
Public gLngMaxW As Double

'Command�������
Public Const INFINITE As Long = &HFFFF&
Private Const SW_HIDE As Integer = 0 '���ش��ڣ�������һ������
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
'ע���ȫ��������
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
    
    '����˵��¼�
    If uMsg = WM_COMMAND Then
        '���Ҷ�Ӧ�ļ���
        If wParam >= �˵���׼.�������ܲ˵� - 1 Then
            Select Case wParam
            Case 99999901   '�����Զ��屨��
                Call ExecuteFunc(0, "ZL9REPORT", 99999901)
            End Select
        ElseIf wParam >= �˵���׼.���ڲ˵� - 1 Then '�����б�˵�
            For LngFind = 0 To CollOpenWindowHdl.Count - 1
                If wParam = CollOpenWindowHdl("K_" & LngFind)(2) Then
                    LngTargetHdl = CollOpenWindowHdl("K_" & LngFind)(0)
                    Exit For
                End If
            Next
            
            If IsIconic(LngTargetHdl) Then
                Call ShowWindow(LngTargetHdl, 9)    '��ԭָ������Ϊԭ��С
            End If
            Call SetActiveWindow(LngTargetHdl)
            MenuProc = 1
        ElseIf wParam > �˵���׼.���ܲ˵� - 1 Then  '����ģ���б�˵�
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
            '�ҵ���ִ��
            If BlnFind Then
                Call AddHistory(lngSys & "," & lngModul)
                Call frmMdi.LoadHistory
                
                '���Ҹ�ģ���Ƿ�������,������Ϊ�����
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
                        Call ShowWindow(LngTargetHdl, 9)            '��ԭָ������Ϊԭ��С
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
'���ܣ�ִ�������У�����ȡ���������
    Dim piProc          As PROCESS_INFORMATION '������Ϣ
    Dim stStart         As STARTUPINFO '������Ϣ
    Dim saSecAttr       As SECURITY_ATTRIBUTES '��ȫ����
    Dim lnghReadPipe    As Long '��ȡ�ܵ����
    Dim lnghWritePipe   As Long 'д��ܵ����
    Dim lngBytesRead    As Long '�������ݵ��ֽ���
    Dim strBuffer       As String * 256 '��ȡ�ܵ����ַ���buffer
    Dim lngRet          As Long 'API��������ֵ
    Dim lngRetPro       As Long
    Dim strlpOutputs    As String '���������ս��
    
    DoEvents
    On Error Resume Next
    '���ð�ȫ����
    With saSecAttr
        .nLength = LenB(saSecAttr)
        .bInheritHandle = True
        .lpSecurityDescriptor = 0
    End With
    
    '�����ܵ�
    lngRet = CreatePipe(lnghReadPipe, lnghWritePipe, saSecAttr, 0)
    If lngRet = 0 Then
        strErr = "�޷������ܵ���" & GetLastDllErr()
        Exit Function
    End If
    '���ý�������ǰ����Ϣ
    With stStart
        .Cb = LenB(stStart)
        .dwFlags = STARTF_USESHOWWINDOW Or STARTF_USESTDHANDLES
        .wShowWindow = SW_HIDE
        .hStdOutput = lnghWritePipe '��������ܵ�
        .hStdError = lnghWritePipe '���ô���ܵ�
    End With
    '��������
    'Command = "c:\windows\system32\ipconfig.exe /all" 'DOS������ipconfig.exeΪ��
    lngRetPro = CreateProcess(vbNullString, strCommand & vbNullChar, saSecAttr, saSecAttr, 1&, NORMAL_PRIORITY_CLASS, ByVal 0&, vbNullString, stStart, piProc)
    If lngRetPro = 0 Then
        strErr = "�޷��������̡�" & GetLastDllErr()
        lngRet = CloseHandle(lnghWritePipe)
        lngRet = CloseHandle(lnghReadPipe)
        Exit Function
    Else
        '��Ϊ����д�����ݣ������ȹر�д��ܵ��������������رմ˹ܵ��������޷���ȡ����
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
        Loop While (lngRet <> 0) '��ret=0ʱ˵��ReadFileִ��ʧ�ܣ��Ѿ�û�����ݿɶ���
        '��ȡ������ɣ��رո����
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
'���ܣ�ȥ���ַ�����\0�Ժ���ַ�
    Dim lngPos As Long
    
    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZero = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZero = strInput
    End If
End Function





