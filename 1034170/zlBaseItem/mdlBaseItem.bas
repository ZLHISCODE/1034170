Attribute VB_Name = "mdlBaseItem"
Option Explicit
Public gblnʹ����ҽ As Boolean
Public gbln������ҽ As Boolean
Public gstrҽ�۽ӿڱ�� As String
Public gbln����ҽ���շ���Ŀ As Boolean
Public gbln��������ۿ�  As Boolean
'��ҹ���
Public gobjPlugIn As Object
Public gblnMyStyle As Boolean
Public gstrMatchMode As String
Public gbytCode As Byte
Public Enum gRegType
    gע����Ϣ = 0
    g����ȫ�� = 1
    g����ģ�� = 2
    g˽��ȫ�� = 3
    g˽��ģ�� = 4
End Enum
'-------------------------------------------------------------------------------------------------------------------------------------------------
'--����ϵͳ����
'����:27990
Private Type Ty_System_Para
     bytҩƷ������ʾ As Byte   'ҩƷ������ʾ�������浥����ϸ������������桢ֱ�ӽ����ҩƷѡ����ʱ��ҩƷ������ʾ����0-��ʾͨ������1-��ʾ��Ʒ����2-ͬʱ��ʾͨ��������Ʒ��
     byt����ҩƷ��ʾ As Byte  '����ҩƷ��ʾ��ͨ��������뷽ʽ����ѡ����ʱҩƷ���Ƶ���ʾ����0-������ƥ����ʾ��1-�̶���ʾͨ��������Ʒ��
End Type
Public gTy_System_Para As Ty_System_Para
Public gblnFeeKindCode As Boolean
'Windows���----------------------------------
Public Const GWL_STYLE = (-16)
Public Const WS_MAXIMIZE = &H1000000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const WS_CAPTION = &HC00000
Public Const WS_SYSMENU = &H80000
Public Const WS_THICKFRAME = &H40000
Public Const WS_CHILD = &H40000000
Public Const WS_POPUP = &H80000000
Public Const SWP_NOZORDER = &H4
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long



Public Type BITMAPINFOHEADER '40 bytes
    biSize            As Long
    biWidth           As Long
    biHeight          As Long
    biPlanes          As Integer
    biBitCount        As Integer
    biCompression     As Long
    biSizeImage       As Long
    biXPelsPerMeter   As Long
    biYPelsPerMeter   As Long
    biClrUsed         As Long
    biClrImportant    As Long
End Type
  
Public Type BITMAPFILEHEADER
    bfType            As Integer
    bfSize            As Long
    bfReserved1       As Integer
    bfReserved2       As Integer
    bfOhFileBits         As Long
End Type

Public Type POINTAPI
        X As Long
        Y As Long
End Type
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Public gstrLike As String
Public Const BF_BOTTOM = &H8
Public Const BF_LEFT = &H1
Public Const BF_RIGHT = &H4
Public Const BF_TOP = &H2
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const CB_FINDSTRING = &H14C
Private Const CB_GETCURSEL = &H147
'ϵͳ��������----------------------------------
Public Const SM_CXVSCROLL = 2
Public Const SM_CXHSCROLL = 21
Public Const SM_CYFULLSCREEN = 17
Public Const SM_CXBORDER = 5
Public Const SM_CXFRAME = 32
Public Const SM_CYCAPTION = 4 'Normal Caption
Public Const SM_CYBORDER = 6
Public Const SM_CYFRAME = 33
Public Const SM_CYSMCAPTION = 51 'Small Caption
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Public Const GCST_INVALIDCHAR = " '"    '�����������Ч�ַ�

Public gobjCustAcc As Object

Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Public Enum EditMode 'medit��ʽ  ȡֵΪ��0��������1���޸ģ�2�����ۣ�3��ִ�п��ҡ�4��������Ŀ��5�������޸�ִ�п���
    EditNew = 0
    EditModify = 1
    EditRaise = 2
    EditDept = 3
    EditSlave = 4
    EditCopy = 5
End Enum

Public gobjRIS As Object                    '����RIS�ӿڶ���
Public Enum RISBaseItemOper                 '����RIS�������ݲ������ͣ�1-������2-�޸ģ�3-ɾ��
    AddNew = 1
    Modify = 2
    Delete = 3
End Enum
Public Enum RISBaseItemType                 '����RIS�����������ͣ�3���û�(��Ա��
    Personnel = 3
End Enum
'������־ģ��
Private mobjFso As New FileSystemObject '�ļ�����
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

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

'����������ڼ���Ƿ�Ϸ�����
Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Sub SetFormVisible(ByVal new_Hwnd As Long)
'���ܣ����ش��������С��ť
    SetWindowLong new_Hwnd, GWL_STYLE, GetWindowLong(new_Hwnd, GWL_STYLE) And Not &HCC0000 Or WS_SYSMENU Or &H20000
End Sub

Public Sub IniRIS(Optional ByVal blnMsg As Boolean)
'���ܣ���ʼ�������ӿڲ���
'������blnMsg������ʧ��ʱ�Ƿ���ʾ
    If gobjRIS Is Nothing Then
        On Error Resume Next
        Set gobjRIS = CreateObject("zl9XWInterface.clsHISInner")
        Err.Clear: On Error GoTo 0
    End If
    If gobjRIS Is Nothing Then
        If blnMsg Then
            MsgBox "RIS�ӿڲ���(zl9XWInterface)δ�����ɹ���", vbInformation, gstrSysName
        End If
    End If
End Sub
Public Function MouseInRect(ByVal lngHwnd As Long) As Boolean
'���ܣ��жϵ�ǰ��Ļ����Ƿ���ָ�����ڵ���ʾ������
    Dim vRect As RECT, vPos As POINTAPI
    
    GetCursorPos vPos
    GetWindowRect lngHwnd, vRect
    
    If vPos.X >= vRect.Left And vPos.X <= vRect.Right _
        And vPos.Y >= vRect.Top And vPos.Y <= vRect.Bottom Then
        MouseInRect = True
    End If
End Function

Public Function FormatEx(ByVal vNumber As Variant, ByVal intBit As Integer, _
    Optional blnShowZero As Boolean = False) As String
'���ܣ��������뷽ʽ��ʽ����ʾ����,��֤С������󲻳���0,С����ǰҪ��0
'������vNumber=Single,Double,Currency���͵�����,intBit=���С��λ��
    Dim strNumber As String
            
    If TypeName(vNumber) = "String" Then
        If vNumber = "" Then Exit Function
        If Not IsNumeric(vNumber) Then Exit Function
        vNumber = Val(vNumber)
    End If
    If vNumber = 0 Then
        strNumber = IIF(blnShowZero, 0, "")
    ElseIf Int(vNumber) = vNumber Then
        strNumber = vNumber
    Else
        strNumber = Format(vNumber, "0." & String(intBit, "0"))
        If Left(strNumber, 1) = "." Then strNumber = "0" & strNumber
        If InStr(strNumber, ".") > 0 Then
            Do While Right(strNumber, 1) = "0"
                strNumber = Left(strNumber, Len(strNumber) - 1)
            Loop
        End If
    End If
    FormatEx = strNumber
End Function

Public Function MoveSpecialChar(ByVal strInputString As String) As String
    '1 ȥ��һ���ַ�: " '_%?"����_%?ת��Ϊ��Ӧ��ȫ���ַ�
    '2 ȥ�������ַ�:�˸��Ʊ����С��س�
    Dim n As Integer
    Dim intStrLen As Integer
    Dim intASC As Integer
    Dim strText As String
    Dim strTmp As String
    Const CST_SPECIALCHAR = "_%?"      '����ת�����ַ�
    
    strText = Trim(strInputString)
    
    If strText = "" Then
        MoveSpecialChar = ""
        Exit Function
    End If
    
    intStrLen = Len(strText)
    
    For n = 1 To intStrLen
        If InStr(GCST_INVALIDCHAR & CST_SPECIALCHAR, Mid(strText, n, 1)) = 0 Then
            strTmp = strTmp & Mid(strText, n, 1)
        Else
            Select Case Mid(strText, n, 1)
                Case "?"
                    strTmp = strTmp & "��"
                Case "%"
                    strTmp = strTmp & "��"
                Case "_"
                    strTmp = strTmp & "��"
            End Select
        End If
    Next
    
    strText = strTmp
    strTmp = ""
    
    intStrLen = Len(strText)
    
    If intStrLen = 0 Then
        MoveSpecialChar = ""
        Exit Function
    End If
        
    For n = 1 To intStrLen
        intASC = Asc(Mid(strText, n, 1))
        Select Case intASC
            Case 8, 9, 10, 13, 32
            Case Else
                strTmp = strTmp & Mid(strText, n, 1)
        End Select
    Next
    
    MoveSpecialChar = strTmp
    
End Function
Public Function MatchIndex(ByVal lngHwnd As Long, ByRef KeyAscii As Integer, Optional sngInterval As Single = 1) As Long
'���ܣ�����������ַ����Զ�ƥ��ComboBox��ѡ����,���Զ�ʶ��������
'������lngHwnd=ComboBox��Hwnd����,KeyAscii=ComboBox��KeyPress�¼��е�KeyAscii����,sngInterval=ָ��������
'���أ�-2=δ�Ӵ���,����=ƥ�������(����ƥ�������)
'˵�����뽫�ú�����KeyPress�¼��е��á�

    Static lngPreTime As Single, lngPreHwnd As Long
    Static strFind As String
    
    Dim sngTime As Single, lngR As Long
    Dim lngIndex As Long
    
    If lngPreHwnd <> lngHwnd Then lngPreTime = Empty: strFind = Empty
    lngPreHwnd = lngHwnd
    
    If KeyAscii <> 13 Then
        '����Ѿ�û��ѡ�����ô�������¿�ʼ
        lngIndex = SendMessageLong(lngHwnd, CB_GETCURSEL, 0, 0)
        If lngIndex < 0 Then lngPreTime = 0
        
        sngTime = Timer
        If Abs(sngTime - lngPreTime) > sngInterval Then '������(ȱʡΪ0.5��)
            strFind = ""
        End If
        If KeyAscii = vbKeyEscape Then
            lngPreTime = 0
        Else
            lngPreTime = Timer
        End If
        strFind = strFind & Chr(KeyAscii)
        
        KeyAscii = 0 'ʹComboBox����ĵ���ƥ�书��ʧЧ
        MatchIndex = SendMessage(lngHwnd, CB_FINDSTRING, -1, ByVal strFind)
        If MatchIndex = -1 Then Beep
    Else
        MatchIndex = -2 '������Իس���������
    End If
End Function

Public Sub �ı����(nodParent As Node, int��ȥ���� As Integer, str�������� As String)
'����:�ı������б���ڵ�ı����б����ֵ
'����:nodParent         Ҫ�ı�������ʼ�ڵ�
'     int��ȥ����       ��������ȥ����
'     str��������       ��������������

    Dim nod As Node
    '�����¼�ҲҪ�ı����
    If nodParent.Children > 0 Then
        Set nod = nodParent.Child
        Do While Not (nod Is Nothing)
            nod.Text = "��" & str�������� & Mid(nod.Text, int��ȥ���� + 2)
            �ı���� nod, int��ȥ����, str��������
            Set nod = nod.Next
        Loop
    End If
End Sub

Public Function GetRoot(ByVal nod As Node) As Node
'���ܣ���������ڵ�ĸ��ڵ�
    Dim nodTemp As Node
    
    If nod Is Nothing Then Exit Function
    Set nodTemp = nod
    Do Until nodTemp.Parent Is Nothing
        Set nodTemp = nodTemp.Parent
    Loop
    Set GetRoot = nodTemp
End Function

Public Function GetTextFromCombo(cmbTemp As ComboBox, ByVal blnAfter As Boolean, Optional strSplit As String = "-") As String
'������cmbTemp  ׼����ȡ���ݵ�ComboBox�ؼ�
'      blnAfter ��ʾ��.֮ǰ��֮��ȡֵ
    Dim lngPos As Long
    
    lngPos = InStr(cmbTemp.Text, strSplit)
    If lngPos = 0 Then
        'ֱ�ӷ��������ַ���
        GetTextFromCombo = "'" & cmbTemp.Text & "'"
    Else
        If blnAfter = False Then
            'Բ��֮ǰ
            GetTextFromCombo = "'" & Mid(cmbTemp.Text, 1, lngPos - 1) & "'"
        Else
            GetTextFromCombo = "'" & Mid(cmbTemp.Text, lngPos + 1) & "'"
        End If
    End If
End Function

Public Sub SetComboByText(cmbTemp As ComboBox, ByVal strText As String, ByVal blnAfter As Boolean, Optional strSplit As String = "-")
'������cmbTemp  ׼�����õ�ComboBox�ؼ�
'      blnAfter ��ʾ��.֮ǰ��֮��ȡֵ
    Dim lngPos As Long
    Dim lngCount As Long
    Dim strTemp As String
    Dim blnMatch As Boolean
    
    For lngCount = 0 To cmbTemp.ListCount - 1
        strTemp = cmbTemp.List(lngCount)
        
        lngPos = InStr(strTemp, strSplit)
        If lngPos = 0 Then
            'ֱ�ӷ��������ַ���
            If strText = cmbTemp.Text Then
                blnMatch = True
                Exit For
            End If
        Else
            If blnAfter = False Then
                'Բ��֮ǰ
                If strText = Mid(strTemp, 1, lngPos - 1) Then
                    blnMatch = True
                    Exit For
                End If
            Else
                If strText = Mid(strTemp, lngPos + 1) Then
                    blnMatch = True
                    Exit For
                End If
            End If
        End If
    Next
    If blnMatch = True Then
        '�Ѿ��ҵ�
        cmbTemp.ListIndex = lngCount
    Else
        cmbTemp.ListIndex = -1
        If blnAfter = True Then
            cmbTemp.AddItem strText
        End If
    End If
End Sub

Public Sub ���鱨��(frmParent As Form)
    MsgBox "�����в���ϵͳ����Ա����", vbInformation, gstrSysName
End Sub


Public Function GetPictureInfo(picTemp As StdPicture, Optional strBitmap As String = "") As String
'���һ��ͼƬ����Ϣ
    Dim hFile As Integer
    Dim FileHeader As BITMAPFILEHEADER
    Dim InfoHeader As BITMAPINFOHEADER
    
    If picTemp.Handle = 0 Then
        GetPictureInfo = "��ͼƬ"
        Exit Function
    End If
    
    Dim strFile As String, strPath As String
    Dim intFileNum As Integer
    
    If strBitmap = "" Then
        '������ʱ�ļ�
        strPath = Space(256): strFile = Space(256)
        GetTempPath 256, strPath
        strPath = Left$(strPath, InStr(strPath, Chr(0)) - 1)
        
        GetTempFileName strPath, "pic", 0, strFile
        strFile = Left$(strFile, InStr(strFile, Chr(0)) - 1)
    
        SavePicture picTemp, strFile
    Else
        'ֱ��ʹ�������ļ�
        strFile = strBitmap
    End If
    hFile = FreeFile
    Open strFile For Binary Access Read As #hFile
      Get #hFile, , FileHeader
      Get #hFile, , InfoHeader
    Close #hFile
    
    If strBitmap = "" Then
        'ɾ����ʱ�ļ�
        Kill strFile
    End If
    
    If InfoHeader.biBitCount > 8 Then
         GetPictureInfo = InfoHeader.biWidth & "��" & InfoHeader.biHeight & " " & InfoHeader.biBitCount & "λɫ"
    Else
         GetPictureInfo = InfoHeader.biWidth & "��" & InfoHeader.biHeight & " " & 2 ^ InfoHeader.biBitCount & "ɫ"
    End If
End Function

Public Sub PressShiftTab(bytKey As Byte)
    '���ܣ�����̷���һ����,����SendKey,��������Shift
    '������bytKey=VirtualKey Codes��1-254��������vbKeyTab,vbKeyReturn,vbKeyF4

    Call keybd_event(vbKeyShift, 0, KEYEVENTF_EXTENDEDKEY, 0)
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY, 0)
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
    Call keybd_event(vbKeyShift, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
End Sub

Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '���ܣ� ����ת������
    '������
    '   strOld��ԭ����
    '���أ� �������ɵ�����
    '------------------------------------------------
    Dim intDo As Integer
    Dim strPass As String, strReturn As String, strSource As String, strTarget As String
    
    strPass = "WriteByZybZL"
    strReturn = ""
    
    For intDo = 1 To 12
        strSource = Mid(strOld, intDo, 1)
        strTarget = Mid(strPass, intDo, 1)
        strReturn = strReturn & Chr(Asc(strSource) Xor Asc(strTarget))
    Next
    TranPasswd = strReturn
End Function

Public Function CheckValid() As Boolean
    Dim intAtom As Integer
    Dim blnValid As Boolean
    Dim strSource As String
    Dim strCurrent As String
    Dim strBuffer As String * 256
    CheckValid = False
    
    '��ȡע������������
    strCurrent = Format(Now, "yyyyMMddHHmm")
    intAtom = GetSetting("ZLSOFT", "����ȫ��", "����", 0)
    Call SaveSetting("ZLSOFT", "����ȫ��", "����", 0)
    blnValid = (intAtom <> 0)
    
    '������ڣ���Դ����н���
    If blnValid Then
        Call GlobalGetAtomName(intAtom, strBuffer, 255)
        strSource = Trim(Replace(strBuffer, Chr(0), ""))
        '���Ϊ�գ����ʾ�Ƿ�
        If strSource <> "" Then
            If Left(strSource, 1) <> "#" Then
                strSource = TranPasswd(Mid(strSource, 1, 12))
                If strSource <> strCurrent Then '�ж�ʱ�����Ƿ����1
                    If CStr(Mid(strSource, 11, 2) + 1) = CStr(Mid(strCurrent, 11, 2) + 0) Then
                        '�����ȣ���ͨ��
                    Else
                        '���ȣ���ʾ���ڽ�λ�����Ӧ��Ϊ��
                        If Not (Mid(strCurrent, 11, 2) = "00" And Mid(strSource, 11, 2) = "59") Then blnValid = False
                    End If
                End If
            Else
                blnValid = False
            End If
        Else
            blnValid = False
        End If
    End If
    
    If Not blnValid Then
        MsgBox "The component is lapse��", vbInformation, gstrSysName
        Exit Function
    End If
    CheckValid = True
End Function

Public Function zlGetSymbol(strInput As String, Optional bytIsWB As Byte = 0, Optional intOutNum As Integer = 10) As String
    '----------------------------------
    '���ܣ������ַ����ļ���
    '��Σ�strInput-�����ַ�����bytIsWB-�Ƿ����(����Ϊƴ��)
    '���Σ���ȷ�����ַ��������󷵻�"-"
    '----------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    If bytIsWB Then
        strSQL = "select zlWBcode('" & strInput & "'," & intOutNum & ") from dual"
    Else
        strSQL = "select zlSpellcode('" & strInput & "'," & intOutNum & ") from dual"
    End If
    On Error GoTo ErrHand

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlGetSymbol")
    zlGetSymbol = IIF(IsNull(rsTmp.Fields(0).Value), "", rsTmp.Fields(0).Value)

    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlGetSymbol = "-"
End Function

Public Function OpenIme(Optional StrIme As String) As Boolean
'����:�����ƴ��������뷨,��ָ������ʱ�ر��������뷨��֧�ֲ������ơ�
Dim arrIme(99) As Long, lngCount As Long, strName As String * 255

    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    Do
        lngCount = lngCount - 1
        If ImmIsIME(arrIme(lngCount)) = 1 Then
            ImmGetDescription arrIme(lngCount), strName, Len(strName)
            If InStr(1, Mid(strName, 1, InStr(1, strName, Chr(0)) - 1), StrIme) > 0 And StrIme <> "" Then
                If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True
                Exit Function
            End If
        ElseIf StrIme = "" Then
            If ActivateKeyboardLayout(arrIme(lngCount), 0) <> 0 Then OpenIme = True
            Exit Function
        End If
    Loop Until lngCount = 0
End Function

Public Sub GetFocus(ByVal TxtBox As TextBox)
    With TxtBox
        .SelStart = 0
        .SelLength = LenB(StrConv(.Text, vbFromUnicode))
    End With
End Sub

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIF(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'���ܣ�ģ��Oracle��Decode����
    Dim varValue As Variant, i As Integer
    
    i = 1
    varValue = arrPar(0)
    Do While i <= UBound(arrPar)
        If i = UBound(arrPar) Then
            Decode = arrPar(i): Exit Function
        ElseIf varValue = arrPar(i) Then
            Decode = arrPar(i + 1): Exit Function
        Else
            i = i + 2
        End If
    Loop
End Function

Public Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal lngY As Long) As POINTAPI
'���ܣ��ÿؼ���ָ����������Ļ�е�λ��(Twip)
    Dim vPoint As POINTAPI
    vPoint.X = lngX / Screen.TwipsPerPixelX: vPoint.Y = lngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.X = vPoint.X * Screen.TwipsPerPixelX: vPoint.Y = vPoint.Y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function

Public Function PreFixNO(Optional curDate As Date = #1/1/1900#) As String
'���ܣ����ش�д�ĵ��ݺ���ǰ׺
    If curDate = #1/1/1900# Then
        PreFixNO = CStr(CInt(Format(zlDatabase.Currentdate, "YYYY")) - 1990)
    Else
        PreFixNO = CStr(CInt(Format(curDate, "YYYY")) - 1990)
    End If
    PreFixNO = IIF(CInt(PreFixNO) < 10, PreFixNO, Chr(55 + CInt(PreFixNO)))
End Function

Public Function CloneRecord(rsSource As ADODB.Recordset) As ADODB.Recordset
'���ܣ�Clone����һ�����ؼ�¼��
'������rsSource=���ػ����ݿ��¼��
'˵����1.��Ϊ��¼�������Clone���ܶ��ڼ�¼������ͬ����
'      2.������С��������¼��
    Dim rsClone As New ADODB.Recordset
    Dim i As Long
    
    With rsSource
        For i = 0 To .Fields.Count - 1
            rsClone.Fields.Append .Fields(i).Name, .Fields(i).Type, .Fields(i).DefinedSize, adFldIsNullable
        Next

        rsClone.CursorLocation = adUseClient
        rsClone.LockType = adLockOptimistic
        rsClone.CursorType = adOpenStatic
        rsClone.Open
        
        .Filter = 0
        Do While Not .EOF
            rsClone.AddNew
            For i = 0 To .Fields.Count - 1
                rsClone.Fields(i).Value = .Fields(i).Value
            Next
            rsClone.Update
            .MoveNext
        Loop
    End With
    Set CloneRecord = rsClone
End Function
Public Function zlGetControlRect(ByVal lngHwnd As Long) As RECT
'���ܣ���ȡָ���ؼ�����Ļ�е�λ��(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    zlGetControlRect = vRect
End Function
Public Sub InitSystemPara()
    '����ȫ�ֲ���
    '-------------------------------------------------------------------------------------------------
    gbytCode = Val(zlDatabase.GetPara("���뷽ʽ"))
    '�շ���Ŀ�������ƥ�䷽ʽ:10.����ȫ������ʱֻƥ�����  01.����ȫ����ĸʱֻƥ�����,11���߾�Ҫ��
    gstrMatchMode = zlDatabase.GetPara(44, glngSys, , "00")
    '���������ʱ,���������Ŀʱ,��λ����������
    gblnFeeKindCode = zlDatabase.GetPara(144, glngSys) = "1"
    gstrLike = IIF(Val(zlDatabase.GetPara("����ƥ��")) = 0, "%", "")
    gbln��������ۿ� = zlDatabase.GetPara(93, glngSys) = "1"
    '����:27990
    With gTy_System_Para
        .byt����ҩƷ��ʾ = Val(zlDatabase.GetPara("����ҩƷ��ʾ")) '0-������ƥ����ʾ��1-�̶���ʾͨ��������Ʒ��
        .bytҩƷ������ʾ = Val(zlDatabase.GetPara("ҩƷ������ʾ"))  '��0-��ʾͨ������1-��ʾ��Ʒ����2-ͬʱ��ʾͨ��������Ʒ��
    End With
End Sub
Public Function GetFeeKind() As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select ����, ����, ���� From �շ���Ŀ���"
    Set GetFeeKind = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�շ����")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub FormSetCaption(ByVal objForm As Object, ByVal blnCaption As Boolean, Optional ByVal blnBorder As Boolean = True)
'���ܣ���ʾ������һ������ı�����
'������blnBorder=���ر�������ʱ��,�Ƿ�Ҳ���ش���߿�
    Dim vRect As RECT, lngStyle As Long
    
    Call GetWindowRect(objForm.hwnd, vRect)
    lngStyle = GetWindowLong(objForm.hwnd, GWL_STYLE)
    If blnCaption Then
        lngStyle = lngStyle Or WS_CAPTION Or WS_THICKFRAME
        If objForm.ControlBox Then lngStyle = lngStyle Or WS_SYSMENU
        If objForm.MaxButton Then lngStyle = lngStyle Or WS_MAXIMIZEBOX
        If objForm.MinButton Then lngStyle = lngStyle Or WS_MINIMIZEBOX
    Else
        If blnBorder Then
            lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX)
        Else
            lngStyle = lngStyle And Not (WS_SYSMENU Or WS_CAPTION Or WS_MAXIMIZEBOX Or WS_MINIMIZEBOX Or WS_THICKFRAME)
        End If
    End If
    SetWindowLong objForm.hwnd, GWL_STYLE, lngStyle
    SetWindowPos objForm.hwnd, 0, vRect.Left, vRect.Top, vRect.Right - vRect.Left, vRect.Bottom - vRect.Top, SWP_NOREPOSITION Or SWP_FRAMECHANGED Or SWP_NOZORDER
End Sub
Public Function NeedName(strList As String) As String
    NeedName = Mid(strList, InStr(strList, "-") + 1)
End Function

Public Sub ShowMsgBox(ByVal strMsgInfor As String, Optional blnYesNo As Boolean = False, Optional ByRef blnYes As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ��Ϣ��
    '���:strMsgInfor-��ʾ��Ϣ
    '        blnYesNo-�Ƿ��ṩYES��NO��ť
    '����:
    '����:blnYes-����ṩYESNO��ť,�򷵻�YES(True)��NO(False)
    '����:���˺�
    '����:2010-08-27 16:28:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If blnYesNo = False Then
        MsgBox strMsgInfor, vbInformation + vbDefaultButton1, gstrSysName
    Else
        blnYes = MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
    End If
End Sub


Public Sub SaveRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByVal strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '����:  ��ָ������Ϣ������ע�����
    '����:  RegType-ע������
    '       strSection-ע���Ŀ¼
    '       StrKey-����
    '       strKeyValue-��ֵ
    '����:
    '--------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    Select Case RegType
        Case gע����Ϣ
            SaveSetting "ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue
        Case g����ȫ��
            SaveSetting "ZLSOFT", "����ȫ��\" & strSection, strKey, strKeyValue
        Case g����ģ��
            SaveSetting "ZLSOFT", "����ģ��" & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
        Case g˽��ȫ��
            SaveSetting "ZLSOFT", "˽��ȫ��\" & gstrDbUser & "\" & strSection, strKey, strKeyValue
        Case g˽��ģ��
            SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
    End Select
ErrHand:
End Sub
Public Sub GetRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByRef strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '����:  ��ָ����ע����Ϣ��ȡ����
    '�����:  RegType-ע������
    '       strSection-ע���Ŀ¼
    '       StrKey-����
    '������:
    '       strKeyValue-���صļ�ֵ
    '����:
    '--------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Err = 0
    On Error GoTo ErrHand:
    Select Case RegType
        Case gע����Ϣ
            SaveSetting "ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue
            strKeyValue = GetSetting("ZLSOFT", "ע����Ϣ\" & strSection, strKey, "")
        Case g����ȫ��
            strKeyValue = GetSetting("ZLSOFT", "����ȫ��\" & strSection, strKey, "")
        Case g����ģ��
            strKeyValue = GetSetting("ZLSOFT", "����ģ��" & "\" & App.ProductName & "\" & strSection, strKey, "")
        Case g˽��ȫ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDbUser & "\" & strSection, strKey, "")
        Case g˽��ģ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & strSection, strKey, "")
    End Select
ErrHand:
End Sub

Public Sub zlAddArray(ByRef cllData As Collection, ByVal strSQL As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ָ���ļ����в�������
    '���:cllData-ָ����SQL��
    '     strSql-ָ����SQL���
    '����:���˺�
    '����:2010-08-30 15:51:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    i = cllData.Count + 1
    cllData.Add strSQL, "K" & i
End Sub
Public Sub zlExecuteProcedureArrAy(ByVal cllProcs As Variant, ByVal strCaption As String, Optional blnNoCommit As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ִ����ص�Oracle���̼�
    '���:cllProcs-oracle���̼�
    '     strCaption -ִ�й��̵ĸ����ڱ���
    '     blnNOCommit-ִ������̺�,���ύ����
    '����:
    '����:
    '����:���˺�
    '����:2010-08-30 15:52:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String
    gcnOracle.BeginTrans
    For i = 1 To cllProcs.Count
        strSQL = cllProcs(i)
        Call zlDatabase.ExecuteProcedure(strSQL, strCaption)
    Next
    If blnNoCommit = False Then
        gcnOracle.CommitTrans
    End If
End Sub

Public Function SeekCboIndex(objCbo As Object, lngData As Long) As Long
'���ܣ���ItemData����ComboBox������ֵ
    Dim i As Integer
    
    SeekCboIndex = -1
    If lngData <> 0 Then
        For i = 0 To objCbo.ListCount - 1
            If objCbo.ItemData(i) = lngData Then
                SeekCboIndex = i: Exit Function
            End If
        Next
    End If
End Function
 
Public Sub CalcPosition(ByRef X As Single, ByRef Y As Single, ByVal objBill As Object, Optional blnNoBill As Boolean = False)
    '----------------------------------------------------------------------
    '���ܣ� ����X,Y��ʵ�����꣬��������Ļ���������
    '������ X---���غ��������
    '       Y---�������������
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.hwnd, objPoint)
    If blnNoBill Then
        X = objPoint.X * 15 'objBill.Left +
        Y = objPoint.Y * 15 + objBill.Height '+ objBill.Top
    Else
        X = objPoint.X * 15 + objBill.CellLeft
        Y = objPoint.Y * 15 + objBill.CellTop + objBill.CellHeight
    End If
End Sub

Public Function MedicalTeamPatients(ByVal lngTeamID As Long, ByVal lngMemberID As Long) As String
'----------------------------------------------------------------------
'���ܣ� �г�ҽ��С��ҽ���Ĳ�����Ϣ
'������ lngTeamID: ҽ��С��ID
'       lngMemberID: ҽ��ID
'���أ� ������Ϣ�ַ���
'----------------------------------------------------------------------
    Dim strMess As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo errHandle
    gstrSQL = "Select a.����id, a.סԺ��, a.��Ժ����, b.����" & vbNewLine & _
              "From ������ҳ a, ������Ϣ b " & vbNewLine & _
              "Where a.סԺҽʦ = (Select ����" & vbNewLine & _
              "              From ��Ա��" & vbNewLine & _
              "              Where ID = [2] And (����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or ����ʱ�� Is Null)) And" & vbNewLine & _
              "      a.ҽ��С��id = [1] and a.����id=b.����id and a.��ҳid=b.��ҳid and b.��Ժ=1 "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "ҽ��С��ҽ��������Ϣ", lngTeamID, lngMemberID)
    With rsTmp
        For i = 1 To .RecordCount
            strMess = strMess & "������" & !���� & "��" & vbTab & _
                      "סԺ�ţ�" & IIF(IsNull(!סԺ��), "", !סԺ��) & "��" & vbTab & _
                      "���ţ�" & IIF(IsNull(!��Ժ����), "", !��Ժ����) & vbTab & vbNewLine
            .MoveNext
        Next
    End With
    MedicalTeamPatients = strMess
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Function CheckDeptPermission(ByVal lngOperationID As Long, Optional ByVal lngDeptID As Long) As Boolean
'����: ��鲿��Ȩ��
'lngOperationID: Ҫ��������ԱID
'lngDeptID: Ҫ������Ա�Ĳ���ID
'����: True��Ȩ��, False��Ȩ��
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errHandle
    If lngDeptID = 0 Then
        gstrSQL = "Select Count(*) Rec From ������Ա " & _
                  "Where ��Աid = [2] And [3] In (Select ����id From ������Ա Where ��Աid = [1])"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "�����Ա�Ĳ���Ȩ��", glngUserId, lngOperationID, lngDeptID)
    Else
        gstrSQL = "Select ID " & _
                  "From ���ű� " & _
                  "  Start With ID In (Select ����id From ������Ա Where ��Աid = [1]) " & _
                  "  Connect By Prior ID = �ϼ�id"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "�����Ա�Ĳ���Ȩ��", glngUserId)
        Do While Not rsTmp.EOF
            If rsTmp!ID = lngDeptID Then
                CheckDeptPermission = True
                Exit Function
            End If
            rsTmp.MoveNext
        Loop
        rsTmp.Close
    End If
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub WriteLog(ByVal strLogTxt As String)
    'дһ����־������������лس�,���з����滻Ϊ<CR><LF>
    '��־�����ڵ�ǰĿ¼�µ�[Ӧ�ó�������]LogĿ¼�£��ļ���Ϊ����.txt,Ĭ�ϱ���7�����־��

    Dim strLogPath As String, strLogFile  As String, strLogIni As String    '��־·�����ļ����������ļ���
    Dim strLogSaveDays As String '��־��������
    Dim dblFreeSpace As Double   'ʣ��ռ�
    Dim strDelOldFile As String  '�����ļ�
    Dim objFile As File

    If Val(OS.IniRead("LOG", "OPENLOG", App.Path & "\CONFIG.INI")) = 0 Then Exit Sub
    'ʼ�ձ�����־
    '2�����������־
    strLogSaveDays = "7"  '����7�����־
    strLogPath = App.Path
    
    strDelOldFile = Dir(strLogPath & "\��־*.log")
    Do While strDelOldFile <> ""
        Set objFile = mobjFso.GetFile(strLogPath & "\" & strDelOldFile)
        If DateDiff("d", objFile.DateLastModified, Now) > Val(strLogSaveDays) Then
            mobjFso.DeleteFile strLogPath & "\" & strDelOldFile, True
        End If
        strDelOldFile = Dir
    Loop
    '3���ռ��Ƿ��㹻
    dblFreeSpace = GetFreeSpace(strLogPath)
    If dblFreeSpace >= 1024 And dblFreeSpace <= 10240 Then
        '�ռ䲻�㣬��д��־,����һ�������ļ�
        If Not mobjFso.FileExists(strLogPath & "\�ռ䲻��.txt") Then Call mobjFso.CreateTextFile(strLogPath & "\�ռ䲻��.txt", True)
        Exit Sub
    Else
        '��������ļ�
        If mobjFso.FileExists(strLogPath & "\�ռ䲻��.txt") Then Call mobjFso.DeleteFile(strLogPath & "\�ռ䲻��.txt", True)
    End If
    '4��д����־��
    strLogFile = strLogPath & "\��־" & Format(Now, "yyyyMMdd") & ".log"

    Call SaveLog(strLogFile, strLogTxt)

End Sub

Public Sub SaveLog(ByVal strFileName As String, ByVal strInput As String, Optional ByVal strDate As String)
 
    Dim objStream As TextStream
    Dim strWritLing As String
    
    strWritLing = Replace$(strInput, Chr(&HD), "<CR>")
    strWritLing = Replace$(strInput, Chr(&HA), "<LF>")

    If strInput <> "" Then
        If Not mobjFso.FileExists(strFileName) Then Call mobjFso.CreateTextFile(strFileName)
        Set objStream = mobjFso.OpenTextFile(strFileName, ForAppending)
        If strDate = "" Then
            strDate = Format(Now(), "yyyy-MM-dd HH:mm:ss")
            objStream.WriteLine (strDate & Chr(&H9) & strInput)
        Else
            objStream.WriteLine (strInput)
        End If
        objStream.Close
        Set objStream = Nothing
    End If
    
End Sub

Private Function GetFreeSpace(ByVal strPath As String) As Double
    '��ȡʣ��ռ�
    Dim strDriv As String, Drv As Drive
    Dim strDir As String
    
    If mobjFso.FolderExists(strPath) Then
        strDriv = mobjFso.GetDriveName(mobjFso.GetAbsolutePathName(strPath))
        Set Drv = mobjFso.GetDrive(strDriv)
        If Drv.IsReady Then
            GetFreeSpace = Drv.FreeSpace
        End If
        Set Drv = Nothing
    End If
End Function

Public Function FuncGetStr(ByVal strVal As String) As String
    strVal = Replace(strVal, vbTab, "")
    strVal = Replace(strVal, vbCrLf, "")
    strVal = Replace(strVal, Chr(10), "")
    strVal = Replace(strVal, "'", "''")
    strVal = Replace(strVal, " ", "")
    FuncGetStr = Trim(strVal)
End Function

