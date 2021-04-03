Attribute VB_Name = "mdlIME"
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
'�л���ָ�������뷨��
Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long

'����ϵͳ�п��õ����뷨�����������뷨����Layout,����Ӣ�����뷨��
Public Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long

'��ȡĳ�����뷨������
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long

'�ж�ĳ�����뷨�Ƿ��������뷨
Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long

'''''''''''''''''''''
'��ȡָ�����뷨����Layout,����Ϊ0ʱ��ʾ��ǰ���뷨��
Public Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
'��ȡ��ǰ���뷨����Layout��
Public Declare Function GetKeyboardLayoutName Lib "user32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
'�������뷨Layout���������뷨�л������뷨�л�˳�����ǰͷ(������������Ч),flags����=KLF_REORDER
Public Declare Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal flags As Long) As Long

Public Const KLF_REORDER = &H8

Public Function SystemImes() As Variant
'���ܣ���ϵͳ�������뷨���Ʒ��ص�һ���ַ���������
'���أ�����������������뷨,�򷵻ؿմ�
    Dim arrIme(99) As Long, arrName() As String
    Dim lngLen As Long, strName As String * 255
    Dim lngCount As Long, i As Integer, j As Integer

    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    For i = 0 To lngCount - 1
        If ImmIsIME(arrIme(i)) = 1 Then 'Ϊ1��ʾ�������뷨
            ReDim Preserve arrName(j)
            lngLen = ImmGetDescription(arrIme(i), strName, Len(strName))
            arrName(j) = Mid(strName, 1, InStr(1, strName, Chr(0)) - 1)
            j = j + 1
        End If
    Next
    SystemImes = IIf(j > 0, arrName, vbNullString)
End Function

Public Function OpenIme(Optional strIme As String) As Boolean
    '����:�����ƴ��������뷨,��ָ������ʱ�ر��������뷨��֧�ֲ������ơ�
    Dim arrIme(99) As Long, lngCount As Long, strName As String * 255
    
    If GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDbUser, "���뷨", "") = "" Then Exit Function

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
        MsgBox "�㻹û��װ�κκ������뷨������ʹ�ñ����ܡ�" & vbCrLf & _
               "���뷨�İ�װ���ڿ����������ɡ�", vbInformation, gstrSysName
        Exit Function
    End If
    cmbIME.Clear
    cmbIME.AddItem "���Զ�����"
    strIme = GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDbUser, "���뷨", "")
    For i = LBound(varIME) To UBound(varIME)
        cmbIME.AddItem varIME(i)
        If strIme = varIME(i) Then cmbIME.Text = strIme
    Next
    If cmbIME.ListIndex < 0 Then cmbIME.ListIndex = 0
    ChooseIME = True
End Function


