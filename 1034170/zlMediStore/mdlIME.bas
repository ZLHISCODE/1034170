Attribute VB_Name = "mdlIME"
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
'切换到指定的输入法。
Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long

'返回系统中可用的输入法个数及各输入法所在Layout,包括英文输入法。
Public Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long

'获取某个输入法的名称
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long

'判断某个输入法是否中文输入法
Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long

'''''''''''''''''''''
'获取指定输入法所在Layout,参数为0时表示当前输入法。
Public Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
'获取当前输入法所在Layout名
Public Declare Function GetKeyboardLayoutName Lib "user32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
'根据输入法Layout名将该输入法切换到输入法切换顺序的最前头(重新启动后无效),flags参数=KLF_REORDER
Public Declare Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal flags As Long) As Long

Public Const KLF_REORDER = &H8

Public Function SystemImes() As Variant
'功能：将系统中文输入法名称返回到一个字符串数组中
'返回：如果不存在中文输入法,则返回空串
    Dim arrIme(99) As Long, arrName() As String
    Dim lngLen As Long, strName As String * 255
    Dim lngCount As Long, i As Integer, j As Integer

    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    For i = 0 To lngCount - 1
        If ImmIsIME(arrIme(i)) = 1 Then '为1表示中文输入法
            ReDim Preserve arrName(j)
            lngLen = ImmGetDescription(arrIme(i), strName, Len(strName))
            arrName(j) = Mid(strName, 1, InStr(1, strName, Chr(0)) - 1)
            j = j + 1
        End If
    Next
    SystemImes = IIf(j > 0, arrName, vbNullString)
End Function

Public Function OpenIme(Optional strIme As String) As Boolean
    '功能:按名称打开中文输入法,不指定名称时关闭中文输入法。支持部分名称。
    Dim arrIme(99) As Long, lngCount As Long, strName As String * 255
    
    If GetSetting("ZLSOFT", "私有全局\" & gstrDbUser, "输入法", "") = "" Then Exit Function

    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    Do
        lngCount = lngCount - 1
        If ImmIsIME(arrIme(lngCount)) = 1 Then
            ImmGetDescription arrIme(lngCount), strName, Len(strName)
            If InStr(1, Mid(strName, 1, InStr(1, strName, Chr(0)) - 1), strIme) > 0 And strIme <> "" Then
                If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True
                Exit Function
            End If
        ElseIf strIme = "" Then
            If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True
            Exit Function
        End If
    Loop Until lngCount = 0
End Function

Public Function ChooseIME(cmbIME As Object) As Boolean
    Dim varIME As Variant
    Dim i As Integer
    Dim strIme As String
    
    varIME = SystemImes
     If Not IsArray(varIME) Then
        MsgBox "你还没安装任何汉字输入法，不能使用本功能。" & vbCrLf & _
               "输入法的安装可在控制面板中完成。", vbInformation, gstrSysName
        Exit Function
    End If
    cmbIME.Clear
    cmbIME.AddItem "不自动开启"
    strIme = GetSetting("ZLSOFT", "私有全局\" & gstrDbUser, "输入法", "")
    For i = LBound(varIME) To UBound(varIME)
        cmbIME.AddItem varIME(i)
        If strIme = varIME(i) Then cmbIME.Text = strIme
    Next
    If cmbIME.ListIndex < 0 Then cmbIME.ListIndex = 0
    ChooseIME = True
End Function


