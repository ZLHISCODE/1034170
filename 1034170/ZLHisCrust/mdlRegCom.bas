Attribute VB_Name = "mdlRegCom"
'**************************************
'ģ����: ActiveX Dll ע��/��ע��
'����:�ڳ�����ע��ͷ�ע�ᣬ��regsvr32���Լ�����
'����Inputs:�ļ���
'����:7 ����־�����忴����ע��
'��д����:ף��
'**************************************

Option Explicit


Private Declare Function LoadLibraryRegister Lib "KERNEL32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibraryRegister Lib "KERNEL32" Alias "FreeLibrary" (ByVal hLibModule As Long) As Long
Private Declare Function CloseHandle Lib "KERNEL32" (ByVal hObject As Long) As Long
Private Declare Function GetProcAddressRegister Lib "KERNEL32" Alias "GetProcAddress" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function CreateThreadForRegister Lib "KERNEL32" Alias "CreateThread" (lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpparameter As Long, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function WaitForSingleObject Lib "KERNEL32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeThread Lib "KERNEL32" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Sub ExitThread Lib "KERNEL32" (ByVal dwExitCode As Long)

Private Const STATUS_WAIT_0 = &H0
Private Const WAIT_OBJECT_0 = ((STATUS_WAIT_0) + 0)
Private Const NOERRORS As Long = 0

Private Enum stRegisterStatus
    stFileCouldNotBeLoadedIntoMemorySpace = 1       '�����ڴ�ʧ��
    stNotAValidActiveXComponent = 2                 '����Active���
    stActiveXComponentRegistrationFailed = 3        'ע��ʧ��
    stActiveXComponentRegistrationSuccessful = 4    'ע��ɹ�
    stActiveXComponentUnRegisterSuccessful = 5      '��ע��ɹ�
    stActiveXComponentUnRegistrationFailed = 6      '��ע��ʧ��
    stNoFileProvided = 7                            'û���ҵ�ָ���ļ�
End Enum


Public Function Register(ByVal p_sFileName As String) As Variant
    Dim lLib As Long
    Dim lProcAddress As Long
    Dim lThreadID As Long
    Dim lSuccess As Long
    Dim lExitCode As Long
    Dim lThreadHandle As Long
    Dim lRet As Long
    On Error GoTo ErrorHandler


    If lRet = NOERRORS Then


        If p_sFileName = "" Then
            lRet = stNoFileProvided
        End If
    End If


    If lRet = NOERRORS Then
        lLib = LoadLibraryRegister(p_sFileName)


        If lLib = 0 Then
            lRet = stFileCouldNotBeLoadedIntoMemorySpace
        End If
    End If


    If lRet = NOERRORS Then
        lProcAddress = GetProcAddressRegister(lLib, "DllRegisterServer")


        If lProcAddress = 0 Then
            lRet = stNotAValidActiveXComponent
        Else
            lThreadHandle = CreateThreadForRegister(0, 0, lProcAddress, 0, 0, lThreadID)


            If lThreadHandle <> 0 Then
                lSuccess = (WaitForSingleObject(lThreadHandle, 10000) = WAIT_OBJECT_0)


                If lSuccess = 0 Then
                    Call GetExitCodeThread(lThreadHandle, lExitCode)
                    Call ExitThread(lExitCode)
                    lRet = stActiveXComponentRegistrationFailed
                Else
                    lRet = stActiveXComponentRegistrationSuccessful
                End If
            End If
        End If
    End If
ExitRoutine:
    Register = lRet


    If lThreadHandle <> 0 Then
        Call CloseHandle(lThreadHandle)
    End If


    If lLib <> 0 Then
        If lRet <> 2 Then
            Call FreeLibraryRegister(lLib)
        End If
    End If
    Exit Function
ErrorHandler:
    lRet = Err.Number
    GoTo ExitRoutine
End Function

Public Function UnRegister(ByVal p_sFileName As String) As Variant
    Dim lLib As Long
    Dim lProcAddress As Long
    Dim lThreadID As Long
    Dim lSuccess As Long
    Dim lExitCode As Long
    Dim lThreadHandle As Long
    Dim lRet As Long
    On Error GoTo ErrorHandler


    If lRet = NOERRORS Then


        If p_sFileName = "" Then
            lRet = stNoFileProvided
        End If
    End If


    If lRet = NOERRORS Then
        lLib = LoadLibraryRegister(p_sFileName)


        If lLib = 0 Then
            lRet = stFileCouldNotBeLoadedIntoMemorySpace
        End If
    End If


    If lRet = NOERRORS Then
        lProcAddress = GetProcAddressRegister(lLib, "DllUnregisterServer")


        If lProcAddress = 0 Then
            lRet = stNotAValidActiveXComponent
        Else
            lThreadHandle = CreateThreadForRegister(0, 0, lProcAddress, 0, 0, lThreadID)


            If lThreadHandle <> 0 Then
                lSuccess = (WaitForSingleObject(lThreadHandle, 10000) = WAIT_OBJECT_0)


                If lSuccess = 0 Then
                    Call GetExitCodeThread(lThreadHandle, lExitCode)
                    Call ExitThread(lExitCode)
                    lRet = stActiveXComponentUnRegistrationFailed
                Else
                    lRet = stActiveXComponentUnRegisterSuccessful
                End If
            End If
        End If
    End If
ExitRoutine:
    UnRegister = lRet


    If lThreadHandle <> 0 Then
        Call CloseHandle(lThreadHandle)
    End If


    If lLib <> 0 Then
        If lRet <> 2 Then
            Call FreeLibraryRegister(lLib)
        End If
    End If
    Exit Function
ErrorHandler:
    lRet = Err.Number
    GoTo ExitRoutine
End Function

'''Public Function RegSvr32(ByVal FileName As String, bUnReg As Boolean) As Boolean
'''    Dim lLib     As Long
'''    Dim lProcAddress     As Long
'''    Dim lThreadID     As Long
'''    Dim lSuccess     As Long
'''    Dim lExitCode     As Long
'''    Dim lThread     As Long
'''    Dim bAns     As Boolean
'''    Dim sPurpose     As String
'''
'''    sPurpose = IIf(bUnReg, "DllUnregisterServer ", _
'''        "DllRegisterServer ")
'''
'''    If Dir(FileName) = " " Then Exit Function
'''
'''    lLib = LoadLibraryRegister(FileName)
'''    'could   load   file
'''    If lLib = 0 Then Exit Function
'''
'''    lProcAddress = GetProcAddressRegister(lLib, sPurpose)
'''
'''    If lProcAddress = 0 Then
'''        'Not   an   ActiveX   Component
'''          FreeLibraryRegister lLib
'''          Exit Function
'''    Else
'''          lThread = CreateThreadForRegister(ByVal 0&, 0&, ByVal lProcAddress, ByVal 0&, 0&, lThread)
'''          If lThread Then
'''                    lSuccess = (WaitForSingleObject(lThread, 10000) = 0)
'''                    If Not lSuccess Then
'''                          Call GetExitCodeThread(lThread, lExitCode)
'''                          Call ExitThread(lExitCode)
'''                          bAns = False
'''                          Exit Function
'''                    Else
'''                          bAns = True
'''                    End If
'''                    CloseHandle lThread
'''                    FreeLibraryRegister lLib
'''          End If
'''    End If
'''            RegSvr32 = bAns
'''End Function


'��ע��ACTIVEX EXE
Public Function UnRegServer(ByVal p_sFileName As String)
    Dim iTask As Long
    Dim pHandle As Long
    Dim ret As Long
    
    iTask = Shell(p_sFileName & " /UnRegServer", vbNormalFocus)
    pHandle = OpenProcess(SYNCHRONIZE, False, iTask)
    ret = WaitForSingleObject(pHandle, INFINITE)
    ret = CloseHandle(pHandle)
    
End Function

'ע��ACTIVEX EXE
Public Function RegServer(ByVal p_sFileName As String)
    Dim iTask As Long
    Dim pHandle As Long
    Dim ret As Long
    iTask = Shell(p_sFileName & " /RegServer", vbNormalFocus)
    pHandle = OpenProcess(SYNCHRONIZE, False, iTask)
    ret = WaitForSingleObject(pHandle, INFINITE)
    ret = CloseHandle(pHandle)
End Function
