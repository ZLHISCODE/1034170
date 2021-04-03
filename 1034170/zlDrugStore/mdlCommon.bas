Attribute VB_Name = "mdlCommon"
Option Explicit
Private mobjVoice As Object

'--�����С--
Public glngMinW As Double
Public glngMinH As Double
Public glngMaxW As Double
Public glngMaxH As Double
Public glngOld As Long

Private Type MousePoint
    CurX As Single
    CurY As Single
End Type
Public CurMousePoint As MousePoint          '���λ��

Public Const CALLSOUND_SYSTEM = 0
Public Const CALLSOUND_MS = 1

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

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

'����������ڼ���Ƿ�Ϸ�����
Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Public Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long


Public Const KLF_REORDER = &H8
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2 'ǳ����
Public Const BDR_RAISEDINNER = &H4 'ǳ͹��
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER) '��͹��
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER) '���
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER) 'Frame������ʽ
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER) '��Frame������ʽ
Public Const SRCCOPY = &HCC0020
Public Const WH_KEYBOARD = 2
Public Const HC_ACTION = 0
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Const BF_SOFT = &H1000
Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const WM_CLOSE = &H10
Public Const CB_FINDSTRING = &H14C
Public Const GWL_HWNDPARENT = (-8)
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
Public Const GWL_WNDPROC = -4
Public Const WM_GETMINMAXINFO = &H24

Public Const OPAQUE = 2
Public Const ETO_OPAQUE = 2
Public Const DT_CENTER = &H1
Public Const DT_VCENTER = &H4
Public Const DT_SINGLELINE = &H20

Public Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long

Public Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long

'�������ŵĺ���
Public Declare Function StartTextPlay Lib "StrSound.dll" (ByVal PlayText As String, ByVal intxx As Integer) As Long
Public Declare Function StopPlayStr Lib "StrSound" () As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function InitHtTextSound Lib "StrSound.dll" () As Boolean
Public Sub zlCall_MsSoundPlay(ByVal strCall As String, ByVal intVoiceSpeed As Integer)
    Dim Token As Object
    
    If mobjVoice Is Nothing Then
        Set mobjVoice = CreateObject("SAPI.SpVoice")
'        Set mobjVoice.Voice = mobjVoice.GetVoices("Name=Microsoft Lili").Item(0)
    End If
    
'    For Each Token In objVoice.GetVoices
'        Debug.Print Token.GetDescription()
'    Next
    
'    Microsoft Lili - Chinese(China)
'    Microsoft Anna - English (United States)
'    Microsoft Simplified Chinese
    
    '��������
'    Set objVoice.Voice = objVoice.GetVoices("Name=Microsoft Simplified Chinese").Item(0)
'    Set objVoice.Voice = objVoice.GetVoices("Name=Microsoft Sam").Item(0)
'    Set objVoice.Voice = objVoice.GetVoices("Name=Girl XiaoKun").Item(0)
    
    
    If intVoiceSpeed > 10 Or intVoiceSpeed < -10 Then
        intVoiceSpeed = -4
    End If
    
    mobjVoice.Rate = intVoiceSpeed   '�ٶ�:-10,10  0
    mobjVoice.Volume = 100           '����:0,100 100
    
'    objVoice.Speak "�롢" & "������" & "������" & "����һ�Ŵ���"
'    objVoice.Speak "�롢" & "�����¡�" & "�����¡�" & "����һ�Ŵ���"
'    objVoice.Speak "�롢" & "�� ����" & "�� ����" & "����ҩ������"
    
    mobjVoice.speak strCall, 1
'    Set objVoice = Nothing
End Sub

Public Sub zlCall_SystemSoundPlay(ByVal strCall As String, ByVal intVoiceSpeed As Integer)
'    Call StartTextPlay("�롢" & "�����¡�" & "�����¡�" & "����һ�Ŵ���", 60)
    
    If intVoiceSpeed > 100 Or intVoiceSpeed < 0 Then
        intVoiceSpeed = 65
    End If
    
    Call StartTextPlay(strCall, intVoiceSpeed)
End Sub



Public Function GetArrayByStr(ByVal strInput As String, ByVal lngLength As Long, ByVal strSplitChar As String) As Variant
    '���ݴ�����ַ������зֽ⣬����ָ���ַ����Ⱦ���Ҫ���зֽ⣬������浽������
    '��Σ�strInput-������ַ�����strSplitChar-�ַ��������ݵķָ���
    '���أ����飬���������Ա���ַ����Ȳ�����ָ������
    Dim strArray As Variant
    Dim ArrTmp As Variant
    Dim strTmp As String
    Dim lngCount As Long
    Dim i As Long
    
    strArray = Array()
   
    '����ָ���ַ�ʱ����Ҫ�ֽ�
    If Len(strInput) > lngLength Then
        If strSplitChar = "" Then
            '�޷ָ���ʱ
            strTmp = strInput
            Do While Len(strTmp) > lngLength
                ReDim Preserve strArray(UBound(strArray) + 1)
                strArray(UBound(strArray)) = Mid(strTmp, 1, lngLength)
                strTmp = Mid(strTmp, lngLength + 1)
            Loop
            
            If strTmp <> "" Then
                ReDim Preserve strArray(UBound(strArray) + 1)
                strArray(UBound(strArray)) = strTmp
            End If
        Else
            '�зָ���ʱ
            ArrTmp = Split(strInput & strSplitChar, strSplitChar)
            lngCount = UBound(ArrTmp)
        
            For i = 0 To lngCount
                If ArrTmp(i) <> "" Then
                    '�зָ�������Ҫ���ַָ���֮���ַ��������ԣ����ܰѷָ���֮����ַ���
                    If Len(IIf(strTmp = "", "", strTmp & strSplitChar) & ArrTmp(i)) > lngLength Then
                        ReDim Preserve strArray(UBound(strArray) + 1)
                        strArray(UBound(strArray)) = strTmp
                        strTmp = ArrTmp(i)
                    Else
                        strTmp = IIf(strTmp = "", "", strTmp & strSplitChar) & ArrTmp(i)
                    End If
                End If
                       
                If i = lngCount Then
                    ReDim Preserve strArray(UBound(strArray) + 1)
                    strArray(UBound(strArray)) = strTmp
                End If
            Next
        End If
    Else
        ReDim Preserve strArray(UBound(strArray) + 1)
        strArray(UBound(strArray)) = strInput
    End If
    
    GetArrayByStr = strArray
End Function


Public Function SysColor2RGB(ByVal lngColor As Long) As Long
'���ܣ���VB��ϵͳ��ɫת��ΪRGBɫ
    If lngColor < 0 Then
        Call OleTranslateColor(lngColor, 0, lngColor)
    End If
    SysColor2RGB = lngColor
End Function
Public Function AviShow(FrmMain As Form, Optional ByVal BlnShow As Boolean = True)
    '����Flash����
    DoEvents
    
    If BlnShow Then
        zlCommFun.ShowFlash "���ڲ�������,���Ժ�...", FrmMain
    Else
        zlCommFun.StopFlash
    End If
    
    DoEvents
End Function

Public Function AnalyseComputer() As String
    Dim strComputer As String * 256
    Call GetComputerName(strComputer, 255)
    AnalyseComputer = strComputer
    AnalyseComputer = Replace(AnalyseComputer, Chr(0), "")
End Function

Public Function CheckValid() As Boolean
    Dim intAtom As Integer
    Dim blnValid As Boolean
    Dim strSource As String
    Dim strCurrent As String
    Dim strBuffer As String * 256
    CheckValid = False
    
    '��ȡע�������������
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

Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '���ܣ� ����ת������
    '������
    '   strOld��ԭ����
    '���أ� �������ɵ�����
    '------------------------------------------------
    Dim intDO As Integer
    Dim strPass As String, strReturn As String, strSource As String, strTarget As String
    
    strPass = "WriteByZybZL"
    strReturn = ""
    
    For intDO = 1 To 12
        strSource = Mid(strOld, intDO, 1)
        strTarget = Mid(strPass, intDO, 1)
        strReturn = strReturn & Chr(Asc(strSource) Xor Asc(strTarget))
    Next
    TranPasswd = strReturn
End Function

Public Function ��ͬ����(ByVal sinFirst As Single, ByVal sinSecond As Single) As Boolean
    Dim blnFirst_���� As Boolean, blnSecond_���� As Boolean
    ��ͬ���� = False
    
    blnFirst_���� = (sinFirst <= 0)
    blnSecond_���� = (sinSecond <= 0)
    
    ��ͬ���� = (blnFirst_���� = blnSecond_����)
End Function

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

Public Function InDesign() As Boolean
    'InDesign = False: Exit Function
    
    On Error Resume Next
    Debug.Print 1 / 0
    If err.Number <> 0 Then err.Clear: InDesign = True
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

Public Function MatchIndex(ByVal lngHwnd As Long, ByRef KeyAscii As Integer, Optional sngInterval As Single = 1) As Long
'���ܣ�����������ַ����Զ�ƥ��ComboBox��ѡ����,���Զ�ʶ��������
'������lngHwnd=ComboBox��Hwnd����,KeyAscii=ComboBox��KeyPress�¼��е�KeyAscii����,sngInterval=ָ��������
'���أ�-2=δ�Ӵ���,����=ƥ�������(����ƥ�������)
'˵�����뽫�ú�����KeyPress�¼��е��á�

    Static lngPreTime As Single, lngPreHwnd As Long
    Static strFind As String
    Dim sngTime As Single, lngR As Long
    
    If lngPreHwnd <> lngHwnd Then lngPreTime = Empty: strFind = Empty
    lngPreHwnd = lngHwnd
    
    If KeyAscii <> 13 Then
        sngTime = Timer
        If Abs(sngTime - lngPreTime) > sngInterval Then '������(ȱʡΪ0.5��)
            strFind = ""
        End If
        strFind = strFind & Chr(KeyAscii)
        lngPreTime = Timer
        KeyAscii = 0 'ʹComboBox�����ĵ���ƥ�书��ʧЧ
        MatchIndex = SendMessage(lngHwnd, CB_FINDSTRING, -1, ByVal strFind)
        If MatchIndex = 0 Then Beep
    Else
        MatchIndex = -2 '������Իس���������
    End If
    If MatchIndex = -1 Then MatchIndex = 1
End Function

Public Sub RaisEffect(picBox As PictureBox, Optional IntStyle As Integer, Optional strName As String = "")
    '��PictureBoxģ���3Dƽ�水ť
    'intStyle=0=ƽ��,-1=����,1=͹��,-2=���,2=��͹��

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

Public Function Custom_WndMessage(ByVal hWnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
'���ܣ��Զ�����Ϣ������������ߴ��������
    If Msg = WM_GETMINMAXINFO Then
        Dim MinMax As MINMAXINFO
        CopyMemory MinMax, ByVal lp, Len(MinMax)
        MinMax.ptMinTrackSize.X = glngMinW \ 15
        MinMax.ptMinTrackSize.Y = glngMinH \ 15
        MinMax.ptMaxTrackSize.X = glngMaxW \ 15
        MinMax.ptMaxTrackSize.Y = glngMaxH \ 15
        CopyMemory ByVal lp, MinMax, Len(MinMax)
        Custom_WndMessage = 1
        Exit Function
    End If
    Custom_WndMessage = CallWindowProc(glngOld, hWnd, Msg, wp, lp)
End Function

Private Function mGet����By����(ByVal str����� As String, ByVal str���� As String, ByVal lngLen As Long) As String
'���ܣ����ݺ��ֵõ������
    Dim lngStart As Long, lngEnd As Long
    Dim str���� As String
    
    lngStart = InStr(str�����, str����)
    If lngStart = 0 Then
        'δ�ڱ�����ҵ����ֱ���
        mGet����By���� = "Z"
        Exit Function
    End If
    
    lngEnd = InStr(lngStart, str�����, "|")
    str���� = Mid(str�����, lngStart, lngEnd - lngStart)
    mGet����By���� = Mid(Split(str����, " ")(1), 1, lngLen)
End Function

Public Function mWBX(ByVal strAsk As String, ByVal lng��ʽ As Long) As String
'���ܣ�����ָ���ַ���������ͼ���
'������strAsk  ���������ַ���
'      lng��ʽ 1-ȡ����ĸ��2-����ʹ���
    Static blnNotFound As Boolean
    Dim lngFile As Long, strFile As String, strReturn As String
    Dim str����� As String, str���� As String, blnǰ��ĸ As Boolean, str���� As String
    Dim intBit As Integer, StrBit As String
    
    If blnNotFound = True Then
        'wbx.txt�ļ�δ�ҵ������ܽ��б����ѯ
        Exit Function
    End If
    
    '���ļ�
    strFile = gstrAviPath
    If Right(strFile, 1) <> "\" Then strFile = strFile & "\"
'    strFile = "C:\AppSoft\"
    strFile = strFile & "wbx.txt"
    
    On Error Resume Next
    lngFile = FreeFile
    Open strFile For Input Access Read As lngFile
    If err <> 0 Then
        blnNotFound = True
'        MsgBox "δ����" & strFile & "�ļ���", vbInformation, gstrSysName
        Exit Function
    End If
    
    '�ҵ�ÿһ���ֶ�Ӧ����
    Do Until EOF(lngFile)
        Line Input #lngFile, strReturn
        If InStr(strAsk, Left(strReturn, 1)) > 0 Then
            '������жϷ����ڲ�����Ҫ��Ϊ�˼ӿ��ٶȣ���Ϊֻ���־�����һ���ж�
            If InStr(strReturn, " ") > 0 Then
                str����� = str����� & strReturn & "|"
            End If
        End If
    Loop
    Close #lngFile
    str����� = UCase(str�����)
    
    '�õ��ַ������к���
    strAsk = StrConv(Trim(strAsk), vbNarrow + vbUpperCase)         '��ȫ��ת��Ϊ��ǣ����ַ�������ת��Сд
    If lng��ʽ = 1 Then
        '������ĸ
        For intBit = 1 To Len(strAsk)
            StrBit = Mid(strAsk, intBit, 1)
            If LenB(StrConv(StrBit, vbFromUnicode)) = 2 Then
                '����
                str���� = str���� & mGet����By����(str�����, StrBit, 1)
                blnǰ��ĸ = False
            ElseIf InStr(" ,.;:", StrBit) > 0 Then
                '�ո�
                blnǰ��ĸ = False
            Else
                If blnǰ��ĸ = False And StrBit >= "A" And StrBit <= "Z" Then
                    'ֻȡһ���ַ���������ĸ
                    str���� = str���� & StrBit
                End If
                blnǰ��ĸ = True
            End If
        Next
    Else
        '����ʹ���
        For intBit = 1 To Len(strAsk)
            StrBit = Mid(strAsk, intBit, 1)
            If LenB(StrConv(StrBit, vbFromUnicode)) = 2 Then
                '����
                str���� = str���� & StrBit
            End If
        Next
        
        Select Case Len(str����)
            Case 0
            Case 1
               str���� = mGet����By����(str�����, str����, 4)
            Case 2
               str���� = mGet����By����(str�����, Mid(str����, 1, 1), 2) & mGet����By����(str�����, Mid(str����, 2, 1), 2)
            Case 3
               str���� = mGet����By����(str�����, Mid(str����, 1, 1), 1) & mGet����By����(str�����, Mid(str����, 2, 1), 1) & mGet����By����(str�����, Mid(str����, 3, 1), 2)
            Case Else
               str���� = mGet����By����(str�����, Mid(str����, 1, 1), 1) & mGet����By����(str�����, Mid(str����, 2, 1), 1) & _
                         mGet����By����(str�����, Mid(str����, 3, 1), 1) & mGet����By����(str�����, Right(str����, 1), 1)
        End Select
    End If
    
    mWBX = str����
End Function

Public Function mPinYin(ByVal strAsk As String) As String
'���ܣ�����ָ���ַ�����ƴ������
'������strAsk  ���������ַ���

    Dim aryStard As Variant
    Dim intBit As Integer, iCount As Integer
    Dim StrCode As String, StrBit As String

'    aryStard = Split("��;��;��;��;��;��;��;��;;��;��;��;��;��;ž;��;��;��;��;��;;��;��;Ѿ;��", ";")
    aryStard = Split("��;��;��;��;��;�;��;��;;��;��;��;��;��;ž;��;��;��;��;��;;;��;Ѿ;��", ";")
    strAsk = StrConv(Trim(strAsk), vbNarrow + vbUpperCase)         '��ȫ��ת��Ϊ��ǣ�Сдת��Ϊ��д
    
    StrCode = ""
    For intBit = 1 To Len(strAsk)
        StrBit = Mid(strAsk, intBit, 1)
        If InStr(1, "��������������������¦���ſ������Ϧϫ�������������������������������", StrBit) > 0 Then
            '�����ֵĴ���
            StrCode = StrCode & Switch(StrBit = "��", "1", StrBit = "��", "2", StrBit = "��", "3", StrBit = "��", "4", StrBit = "��", "5" _
                            , StrBit = "��", "6", StrBit = "��", "7", StrBit = "��", "8", StrBit = "��", "9" _
                            , StrBit = "��", "A", StrBit = "��", "B", StrBit = "��", "G" _
                            , StrBit = "��", "N", StrBit = "ſ", "P", StrBit = "��", "S", StrBit = "��", "W" _
                            , StrBit = "��", "W", StrBit = "Ϧ", "X", StrBit = "ϫ", "X", StrBit = "��", "S" _
                            , StrBit = "��", "X", StrBit = "��", "P", StrBit = "��", "C", StrBit = "�", "X" _
                            , StrBit = "�", "C", StrBit = "��", "D", StrBit = "��", "C", StrBit = "�", "Q" _
                            , StrBit = "��", "T", StrBit = "��", "N", StrBit = "��", "H", StrBit = "��", "D" _
                            , StrBit = "��", "P", StrBit = "��", "Q", StrBit = "��", "Q", StrBit = "��", "T")
        ElseIf Asc(StrBit) < 0 Then
            For iCount = 0 To UBound(aryStard)
                If Len(aryStard(iCount)) <> 0 Then
                    If StrComp(StrBit, aryStard(iCount), vbTextCompare) = -1 Then
                        StrCode = StrCode & Chr(65 + iCount)
                        Exit For
                    ElseIf iCount = UBound(aryStard) Then
                        StrCode = StrCode & "Z"
                    End If
                End If
            Next
        Else
            If StrBit >= "A" And StrBit <= "Z" Then
                StrCode = StrCode & StrBit
            End If
        End If
        If Len(StrCode) >= 10 Then Exit For
    Next
    mPinYin = StrCode

End Function

Public Function ExchangeOrder(ByVal strTmp As String) As String
    Dim IntLocate As Integer
    'Add By Zyb 2002-11-27
    'ת�����򴮣�Asc->Desc��
    ExchangeOrder = strTmp
    IntLocate = InStr(1, ExchangeOrder, strAsc)
    If IntLocate = 0 Then
        ExchangeOrder = Replace(ExchangeOrder, strDesc, strAsc)
    Else
        ExchangeOrder = Replace(ExchangeOrder, strAsc, strDesc)
    End If
End Function

Public Function GetOrder(ByVal strTmp As String) As String
    'Add By Zyb 2002-11-27
    '����Ӧ�����򴮺Ϸ���
    GetOrder = strTmp
    GetOrder = Replace(GetOrder, strAsc, " ASC")
    GetOrder = Replace(GetOrder, strDesc, " DESC")
End Function

Public Function SelAll(txtObj As TextBox)
    With txtObj
        .SelStart = 0
        .SelLength = LenB(StrConv(.Text, vbFromUnicode))
    End With
End Function

Public Function FormatEx(ByVal vNumber As Variant, ByVal intBit As Integer) As String
'���ܣ��������뷽ʽ��ʽ����ʾ����,��֤С������󲻳���0,С����ǰҪ��0
'������vNumber=Single,Double,Currency���͵�����,intBit=���С��λ��
    Dim strNumber As String
            
    If TypeName(vNumber) = "String" Then
        If vNumber = "" Then Exit Function
        If Not IsNumeric(vNumber) Then Exit Function
        vNumber = Val(vNumber)
    End If
            
    If vNumber = 0 Then
        strNumber = 0
    ElseIf Int(vNumber) = vNumber Then
        strNumber = vNumber
    Else
        strNumber = Format(vNumber, "0." & String(intBit, "0"))
        If Left(strNumber, 1) = "." Then strNumber = "0" & strNumber
        If InStr(strNumber, ".") > 0 Then
            Do While Right(strNumber, 1) = "0"
                strNumber = Left(strNumber, Len(strNumber) - 1)
            Loop
            If Right(strNumber, 1) = "." Then strNumber = Left(strNumber, Len(strNumber) - 1)
        End If
    End If
    FormatEx = strNumber
End Function

Public Function NVL(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    NVL = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function GetMoneyFormat() As String
    Dim intDigit As Integer
    Dim strDigit As String
    Const strFormat As String = "#####0."
    
    intDigit = GetDigit(0, 1, 4)
    strDigit = String(intDigit, "0")
    GetMoneyFormat = strFormat & strDigit & ";" & _
                "-" & strFormat & strDigit & "; ;"
End Function

Public Sub GetFocus(ByVal TxtBox As TextBox)
    With TxtBox
        .SelStart = 0
        .SelLength = LenB(StrConv(.Text, vbFromUnicode))
    End With
End Sub


Public Function GetFormat(ByVal dblInput As Double, ByVal intDotBit As Integer) As String
    'ȡ��ֵ��С��λ��
    GetFormat = Format(dblInput, "#0." & String(intDotBit, "0"))
End Function

Public Function GetParentWindow(ByVal hwndFrm As Long) As Long
    On Error Resume Next
    '��ȡָ������ĸ�����
    GetParentWindow = GetWindowLong(hwndFrm, GWL_HWNDPARENT)
End Function

Public Function GetCol(mshFlex As Object, ByVal ColName As String) As Integer
    'ȡָ����ͷ����λ��
    
    Dim i As Integer
    
    On Error GoTo errH
    
    GetCol = -1
    
    If TypeName(mshFlex) = "MSHFlexGrid" Then
        With mshFlex
            For i = 0 To .Cols - 1
                If .TextMatrix(0, i) = ColName Then
                    GetCol = i
                    Exit Function
                End If
            Next
            
        End With
    ElseIf TypeName(mshFlex) = "VSFlexGrid" Then
        With mshFlex
            For i = 0 To .Cols - 1
                If .TextMatrix(0, i) = ColName Then
                    GetCol = i
                    Exit Function
                End If
            Next
            
        End With
    End If
    Exit Function
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetControlRect(ByVal lngHwnd As Long) As RECT
'���ܣ���ȡָ���ؼ�����Ļ�е�λ��(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function

Public Function GetText(ByVal hwndFrm As Long) As String
    Dim strCaption As String * 256
    On Error Resume Next
    '��ȡָ������ı���
    Call GetWindowText(hwndFrm, strCaption, 255)
    GetText = zlCommFun.TruncZero(strCaption)
End Function

Public Sub RefreshRowNO(ByRef mshBill As Object, ByVal lng����� As Long, Optional ByVal lngRow As Long = 1)
    Dim lngRows As Long
    '��ָ���п�ʼ�������
    
    With mshBill
        lngRows = .rows - 1
        For lngRow = lngRow To lngRows
            .TextMatrix(lngRow, lng�����) = lngRow
        Next
    End With
End Sub