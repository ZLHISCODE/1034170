Attribute VB_Name = "mdlCommon"
Option Explicit
Public glngTXTProc As Long '����Ĭ�ϵ���Ϣ�����ĵ�ַ
Public glngOld As Long, glngFormW As Long, glngFormH As Long
Public Const LONG_MAX = 2147483647 'Long�����ֵ
Public Type PointAPI
        X As Long
        Y As Long
End Type
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Type MINMAXINFO
        ptReserved As PointAPI
        ptMaxSize As PointAPI
        ptMaxPosition As PointAPI
        ptMinTrackSize As PointAPI
        ptMaxTrackSize As PointAPI
End Type
Public Const CB_FINDSTRING = &H14C
Public Const CB_GETDROPPEDSTATE = &H157
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const SW_RESTORE = 9
Public Const GWL_WNDPROC = -4
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
Public Const WM_CONTEXTMENU = &H7B ' ���һ��ı���ʱ������������Ϣ
Public Const WM_GETMINMAXINFO = &H24
Public Const SM_CXVSCROLL = 2
Public Const SM_CXHSCROLL = 21
Public Const SM_CYFULLSCREEN = 17
Public Const SM_CXBORDER = 5
Public Const SM_CXFRAME = 32
Public Const SM_CYCAPTION = 4 'Normal Caption
Public Const SM_CYBORDER = 6
Public Const SM_CYFRAME = 33
Public Const SM_CYSMCAPTION = 51 'Small Caption
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Public Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As PointAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As PointAPI) As Long
Public Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWndChild As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Const MK_LBUTTON = &H1 '��ȡ������״̬
Public Const CB_SHOWDROPDOWN = &H14F
Public Const CB_RESETCONTENT = &H14B
Public gstrMatchMethod As String

'��������ComboBox����
Public Const CB_ADDSTRING = &H143
Public Const CB_SETITEMDATA = &H151
Public Declare Function AddComboItem Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetComboData Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function FindComboStr Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

'����API
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

'��ͼAPI
Public Const PS_SOLID = 0
Public Const BS_SOLID = 0
Public Const BS_NULL = 1
Public Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Long) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long


Public Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type
Public Const VER_PLATFORM_WIN32s = 0 'Win32s on Windows 3.1.
Public Const VER_PLATFORM_WIN32_WINDOWS = 1 'Windows 95, Windows 98, or Windows Me.
Public Const VER_PLATFORM_WIN32_NT = 2 'Windows NT 3.51, Windows NT 4.0, Windows 2000, Windows XP, or Windows .NET Server.
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Public Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const WM_PAINT = &HF

Public Const GWL_EXSTYLE = (-20)
Public Const WS_EX_LAYERED As Long = &H80000
Public Const LWA_COLORKEY = &H1
Public Const LWA_ALPHA = &H2

Public Const HWND_TOPMOST = -1
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_HIDEWINDOW = &H80

Declare Function GetScrollRange Lib "user32" (ByVal hWnd As Long, ByVal nBar As Long, lpMinPos As Long, lpMaxPos As Long) As Long
    Public Const SB_HORZ = &H0
    Public Const SB_VERT = &H1

Public Const COLEditBackColor = &HE1FFE1    'ǳ��
Public Const COLSelBackColor = &HFAEADA          'ǳ��

Public Function VsScroll(vsf As VSFlexGrid) As Boolean '�ж�ˮƽ�������Ŀɼ���
    Dim i As Long, lpMinPos As Long, lpMaxPos As Long
    
    VsScroll = False
    i = GetScrollRange(vsf.hWnd, SB_HORZ, lpMinPos, lpMaxPos)
    If lpMaxPos <> lpMinPos Then VsScroll = True
    End Function
    
Public Function HeScroll(vsf As VSFlexGrid) As Boolean '�жϴ�ֱ�������Ŀɼ���
    Dim i As Long, lpMinPos As Long, lpMaxPos As Long
    
    HeScroll = False
    i = GetScrollRange(vsf.hWnd, SB_VERT, lpMinPos, lpMaxPos)
    If lpMaxPos <> lpMinPos Then HeScroll = True
End Function



Public Function MousePressButton(lngTbr As Long, objButton As Button) As Boolean
'���ܣ��жϵ�ǰ��Ļ����Ƿ���ָ�����߰�ť��ʾ�����ڰ���
    Dim vRect As RECT, vPos As PointAPI
        
    '���жϵ�ǰ�Ƿ��ڰ���״̬
    If (GetKeyState(MK_LBUTTON) And &H80) <> 0 Then
        '���жϵ�ǰ�����������Χ
        GetCursorPos vPos
        
        GetWindowRect lngTbr, vRect
        With objButton
            vRect.Left = vRect.Left + .Left / Screen.TwipsPerPixelX
            vRect.Top = vRect.Top + .Top / Screen.TwipsPerPixelY
            vRect.Right = vRect.Left + .Width / Screen.TwipsPerPixelX
            vRect.Bottom = vRect.Top + .Height / Screen.TwipsPerPixelY
        End With
        
        If vPos.X >= vRect.Left And vPos.X <= vRect.Right _
            And vPos.Y >= vRect.Top And vPos.Y <= vRect.Bottom Then
            MousePressButton = True
        End If
    End If
End Function

Public Function MouseInRect(ByVal lngHwnd As Long) As Boolean
'���ܣ��жϵ�ǰ��Ļ����Ƿ���ָ�����ڵ���ʾ������
    Dim vRect As RECT, vPos As PointAPI
    
    GetCursorPos vPos
    GetWindowRect lngHwnd, vRect
    
    If vPos.X >= vRect.Left And vPos.X <= vRect.Right _
        And vPos.Y >= vRect.Top And vPos.Y <= vRect.Bottom Then
        MouseInRect = True
    End If
End Function

Public Sub FormSetCaption(ByVal objForm As Object, ByVal blnCaption As Boolean, Optional ByVal blnBorder As Boolean = True)
'���ܣ���ʾ������һ������ı�����
'������blnBorder=���ر�������ʱ��,�Ƿ�Ҳ���ش���߿�
    Dim vRect As RECT, lngStyle As Long
    
    Call GetWindowRect(objForm.hWnd, vRect)
    lngStyle = GetWindowLong(objForm.hWnd, GWL_STYLE)
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
    SetWindowLong objForm.hWnd, GWL_STYLE, lngStyle
    SetWindowPos objForm.hWnd, 0, vRect.Left, vRect.Top, vRect.Right - vRect.Left, vRect.Bottom - vRect.Top, SWP_NOREPOSITION Or SWP_FRAMECHANGED Or SWP_NOZORDER
End Sub

Public Function MoveObj(lngHwnd As Long) As RECT
'���ܣ��ڶ����MouseDown�¼��е���,����������Hwnd����
'���أ������Ļ������ֵ
    Dim vPos As RECT
    ReleaseCapture
    SendMessage lngHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    GetWindowRect lngHwnd, vPos
    MoveObj = vPos
End Function

Public Sub PopupButtonMenu(ToolBar As Object, Button As Object, objMenu As Object)
'���ܣ�������ʽ���߰�ť�е���һ���˵�
    Dim vRect As RECT, vDot1 As PointAPI, vDot2 As PointAPI
    
    Call GetWindowRect(ToolBar.hWnd, vRect)
    vDot1.X = vRect.Left: vDot1.Y = vRect.Top
    vDot2.X = vRect.Right: vDot2.Y = vRect.Bottom
    
    Call ScreenToClient(ToolBar.Parent.hWnd, vDot1)
    Call ScreenToClient(ToolBar.Parent.hWnd, vDot2)
    
    vDot1.X = vDot1.X * 15: vDot1.Y = vDot1.Y * 15
    vDot2.X = vDot2.X * 15: vDot2.Y = vDot2.Y * 15
    ToolBar.Parent.PopupMenu objMenu, 2, vDot1.X + Button.Left, vDot2.Y
End Sub

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
        If MatchIndex = -1 Then Beep
    Else
        MatchIndex = -2 '������Իس���������
    End If
End Function

Public Function GetMatchingSting(ByVal strString As String, Optional blnUpper As Boolean = True) As String
    '--------------------------------------------------------------------------------------------------------------------------------------
    '����:����ƥ�䴮%
    '����:strString ��ƥ����ִ�
    '     blnUpper-�Ƿ�ת���ڴ�д
    '����:���ؼ�ƥ�䴮%dd%
    '--------------------------------------------------------------------------------------------------------------------------------------
    Dim strLeft As String
    Dim strRight As String
    If gstrMatchMethod = "" Then
        gstrMatchMethod = Val(gobjDatabase.GetPara("����ƥ��"))
    End If
    
    If gstrMatchMethod = "0" Then
        strLeft = "%"
        strRight = "%"
    Else
        strLeft = ""
        strRight = "%"
    End If
    If blnUpper Then
        GetMatchingSting = strLeft & UCase(strString) & strRight
    Else
        GetMatchingSting = strLeft & strString & strRight
    End If
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

Public Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal lngY As Long) As PointAPI
'���ܣ��ÿؼ���ָ����������Ļ�е�λ��(Twip)
    Dim vPoint As PointAPI
    vPoint.X = lngX / Screen.TwipsPerPixelX: vPoint.Y = lngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.X = vPoint.X * Screen.TwipsPerPixelX: vPoint.Y = vPoint.Y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function

Public Function SysColor2RGB(ByVal lngColor As Long) As Long
'���ܣ���VB��ϵͳ��ɫת��ΪRGBɫ
    If lngColor < 0 Then
        Call OleTranslateColor(lngColor, 0, lngColor)
    End If
    SysColor2RGB = lngColor
End Function

'ȥ��TextBox��Ĭ���Ҽ��˵�
Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal Msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' �����Ϣ����WM_CONTEXTMENU���͵���Ĭ�ϵĴ��ں�������
    If Msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(glngTXTProc, hWnd, Msg, wp, lp)
End Function



Public Sub FindCboIndex(objCbo As Object, lngData As Long, Optional Keep As Boolean)
'���ܣ�����Ŀֵ����ComboBox����Ŀ����
'������Keep=���δƥ�䣬�Ƿ񱣳�ԭ����
    Dim i As Integer
    
    If lngData <> 0 Then
        For i = 0 To objCbo.ListCount - 1
            If objCbo.ItemData(i) = lngData Then
                objCbo.ListIndex = i: Exit Sub
            End If
        Next
    End If
    If Not Keep Then objCbo.ListIndex = -1
End Sub

Public Sub GetCboIndex(objCbo As Object, strFind As String, Optional Keep As Boolean)
'���ܣ����ַ�����ComboBox�в�������
'������Keep=���δƥ�䣬�Ƿ񱣳�ԭ����
    Dim i As Integer
    
    '�Ⱦ�ȷ����
    For i = 0 To objCbo.ListCount - 1
        If objCbo.List(i) = strFind Then
            objCbo.ListIndex = i: Exit Sub
        ElseIf NeedName(objCbo.List(i)) = strFind And strFind <> "" Then
            objCbo.ListIndex = i: Exit Sub
        End If
    Next
    
    '���ģ������
    If strFind <> "" Then
        For i = 0 To objCbo.ListCount - 1
            If InStr(objCbo.List(i), strFind) > 0 And strFind <> "" Then
                objCbo.ListIndex = i: Exit Sub
            End If
        Next
    End If
    If Not Keep Then objCbo.ListIndex = -1
End Sub

Public Function SeekCboIndex(objCbo As Object, varData As Variant) As Long
'���ܣ���ItemData��Text����ComboBox������ֵ
    Dim strType As String, i As Integer
    
    SeekCboIndex = -1
    
    strType = TypeName(varData)
    If strType = "Field" Then
        If IsType(varData.Type, adVarChar) Then strType = "String"
    End If
    
    If strType = "String" Then
        If varData <> "" Then
            '�Ⱦ�ȷ����
            For i = 0 To objCbo.ListCount - 1
                If objCbo.List(i) = varData Then
                    SeekCboIndex = i: Exit Function
                ElseIf NeedName(objCbo.List(i)) = varData And varData <> "" Then
                    SeekCboIndex = i: Exit Function
                End If
            Next
            '��ģ������
            For i = 0 To objCbo.ListCount - 1
                If InStr(objCbo.List(i), varData) > 0 And varData <> "" Then
                    SeekCboIndex = i: Exit Function
                End If
            Next
        End If
    Else
        If varData <> 0 Then
            For i = 0 To objCbo.ListCount - 1
                If objCbo.ItemData(i) = varData Then
                    SeekCboIndex = i: Exit Function
                End If
            Next
        End If
    End If
End Function
Public Function IsType(ByVal varType As DataTypeEnum, ByVal varBase As DataTypeEnum) As Boolean
'���ܣ��ж�ĳ��ADO�ֶ����������Ƿ���ָ���ֶ�������ͬһ��(������,����,�ַ�,������)
    Dim intA As Integer, intB As Integer
    
    Select Case varBase
        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
            intA = -1
        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
            intA = -2
        Case adDBTimeStamp, adDBTime, adDBDate, adDate
            intA = -3
        Case adBinary, adVarBinary, adLongVarBinary
            intA = -4
        Case Else
            intA = varBase
    End Select
    Select Case varType
        Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
            intB = -1
        Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
            intB = -2
        Case adDBTimeStamp, adDBTime, adDBDate, adDate
            intB = -3
        Case adBinary, adVarBinary, adLongVarBinary
            intB = -4
        Case Else
            intB = varType
    End Select
    
    IsType = intA = intB
End Function
Public Function InDesign() As Boolean
    On Error Resume Next
    Debug.Print 1 / 0
    If Err.Number <> 0 Then Err.Clear: InDesign = True
End Function

Public Function Custom_WndMessage(ByVal hWnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
'���ܣ��Զ�����Ϣ������������ߴ��������
    If Msg = WM_GETMINMAXINFO Then
        Dim MinMax As MINMAXINFO
        CopyMemory MinMax, ByVal lp, Len(MinMax)
        MinMax.ptMinTrackSize.X = glngFormW \ Screen.TwipsPerPixelX
        MinMax.ptMinTrackSize.Y = glngFormH \ Screen.TwipsPerPixelY
        MinMax.ptMaxTrackSize.X = 1600
        MinMax.ptMaxTrackSize.Y = 1200
        CopyMemory ByVal lp, MinMax, Len(MinMax)
        Custom_WndMessage = 1
        Exit Function
    End If
    Custom_WndMessage = CallWindowProc(glngOld, hWnd, Msg, wp, lp)
End Function
Public Sub zlAddArray(ByRef cllData As Collection, ByVal strSQL As String)
    '---------------------------------------------------------------------------------------------
    '����:��ָ���ļ����в�������
    '����:cllData-ָ����SQL��
    '     strSql-ָ����SQL���
    '����:���˺�
    '����:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    i = cllData.Count + 1
    cllData.Add strSQL, "K" & i
End Sub
Public Sub zlExecuteProcedureArrAy(ByVal cllProcs As Variant, ByVal strCaption As String, Optional blnTrans As Boolean = True, Optional blnCommit As Boolean = True)
    '-------------------------------------------------------------------------------------------------------------------------
    '����:ִ����ص�Oracle���̼�
    '����:cllProcs-oracle���̼�
    '     strCaption -ִ�й��̵ĸ����ڱ���
    '     blnTrans-�Ƿ��������
    '     blnCommit-ִ������̺�,�ύ����(ǰ��:blnTrans=true)
    '����:���˺�
    '����:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String
    If blnTrans Then gcnOracle.BeginTrans
    For i = 1 To cllProcs.Count
        strSQL = cllProcs(i)
        Call gobjDatabase.ExecuteProcedure(strSQL, strCaption)
    Next
    If blnCommit And blnTrans Then
        gcnOracle.CommitTrans
    End If
End Sub
Public Function zlCboFindItem(ByVal cboObj As Object, ByVal lngFindID As Long, _
    Optional strItem As String = "", Optional blnOnlyFind As Boolean = True, Optional blnFindLocal As Boolean = False) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���Combox��ItemData���ݽ��ж�λ
    '��Σ�cboObj-Combox����
    '         lngFindID-��Ҫ���ҵ�ID
    '         strItem-��Ҫ���ҵĻ����ӵ�����(��blnOnlyFind=false)ʱ
    '         blnOnlyFind-�Ƿ����.
    '        blnFindLocal-�ҵ���,��λ��
    '���Σ�
    '���أ��ҵ�,����true,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-04-06 17:28:17
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim lngLocate As Long
    zlCboFindItem = False
    For lngLocate = 0 To cboObj.ListCount - 1
        If cboObj.ItemData(lngLocate) = lngFindID Then
            If blnFindLocal Then cboObj.ListIndex = lngLocate
            zlCboFindItem = True
            Exit Function
        End If
    Next
    If blnOnlyFind Then Exit Function
    cboObj.AddItem strItem
    cboObj.ItemData(cboObj.NewIndex) = lngFindID
    If blnFindLocal Then cboObj.ListIndex = cboObj.NewIndex
    zlCboFindItem = True
End Function
Public Function zlCheckPrivs(ByVal strPrivs As String, ByVal strMyPriv As String) As Boolean
    '---------------------------------------------------------------------------------------------
    '����:���ָ����Ȩ���Ƿ����
    '����:strPrivs-Ȩ�޴�
    '     strMyPriv-����Ȩ��
    '����,����Ȩ��,����true,���򷵻�False
    '����:���˺�
    '����:2008/01/09
    '---------------------------------------------------------------------------------------------
    zlCheckPrivs = InStr(";" & strPrivs & ";", ";" & strMyPriv & ";") > 0
End Function


Public Function zlSaveDockPanceToReg(ByVal frmMain As Form, ByVal objPance As DockingPane, _
                ByVal strKey As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����DockPane�ؼ��ľ���λ��
    '���:frmMain-������
    '     objPance:DockinPane�ؼ�
    '      StrKey-����
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-02-10 14:24:04
    '-----------------------------------------------------------------------------------------------------------
    Dim blnAutoHide As Boolean
    If Val(gobjDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then
        zlSaveDockPanceToReg = True: Exit Function
    End If
    Err = 0: On Error GoTo Errhand:
    objPance.SaveState "VB and VBA Program Settings\ZLSOFT\˽��ģ��\" & gstrDBUser & "\" & App.ProductName, frmMain.Name, "����"
    zlSaveDockPanceToReg = True
Errhand:
End Function

Public Function zlRestoreDockPanceToReg(ByVal frmMain As Form, ByVal objPance As DockingPane, _
                ByVal strKey As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����DockPane�ؼ��ľ���λ��
    '���:frmMain-������
    '     objPance:DockinPane�ؼ�
    '      StrKey-����
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-02-10 14:24:04
    '-----------------------------------------------------------------------------------------------------------
    Dim blnAutoHide As Boolean
    If Val(gobjDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then
        zlRestoreDockPanceToReg = True: Exit Function
    End If
    'blnAutoHide = Val(gobjDataBase.GetPara("������������", , , True)) = 1
    Err = 0: On Error GoTo Errhand:
    objPance.LoadState "VB and VBA Program Settings\ZLSOFT\˽��ģ��\" & gstrDBUser & "\" & App.ProductName, frmMain.Name, "����"
    zlRestoreDockPanceToReg = True
Errhand:
End Function
Public Function Between(X, a, B) As Boolean
'���ܣ��ж�x�Ƿ���a��b֮��
    If a < B Then
        Between = X >= a And X <= B
    Else
        Between = X >= B And X <= a
    End If
End Function
Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '���ܣ� ����ת������
    '������
    '   strOld��ԭ����
    '���أ� �������ɵ�����
    '------------------------------------------------
    Dim intDo As Integer
    Dim StrPass As String, strReturn As String, strSource As String, strTarget As String
    
    StrPass = "WriteByZybZL"
    strReturn = ""
    
    For intDo = 1 To 12
        strSource = Mid(strOld, intDo, 1)
        strTarget = Mid(StrPass, intDo, 1)
        strReturn = strReturn & Chr(Asc(strSource) Xor Asc(strTarget))
    Next
    TranPasswd = strReturn
End Function

Public Function zlInitMEPIPati(ByRef rsPati As ADODB.Recordset) As Boolean
    Set rsPati = New ADODB.Recordset
    With rsPati
        If .State = adStateOpen Then .Close
        With .Fields
            .Append "����ID", adBigInt, , adFldIsNullable
            .Append "��ҳID", adBigInt, , adFldIsNullable
            .Append "�Һ�ID", adBigInt, , adFldIsNullable
            .Append "�����", adVarChar, 18, adFldIsNullable
            .Append "סԺ��", adVarChar, 18, adFldIsNullable
            .Append "ҽ����", adVarChar, 30, adFldIsNullable
            .Append "����֤��", adVarChar, 18, adFldIsNullable
            .Append "����֤��", adVarChar, 20, adFldIsNullable
            .Append "����", adVarChar, 100, adFldIsNullable
            .Append "�Ա�", adVarChar, 4, adFldIsNullable
            .Append "��������", adVarChar, 20, adFldIsNullable
            .Append "�����ص�", adVarChar, 100, adFldIsNullable
            .Append "����", adVarChar, 30, adFldIsNullable
            .Append "����", adVarChar, 20, adFldIsNullable
            .Append "ѧ��", adVarChar, 10, adFldIsNullable
            .Append "ְҵ", adVarChar, 80, adFldIsNullable
            .Append "������λ", adVarChar, 100, adFldIsNullable
            .Append "����", adVarChar, 30, adFldIsNullable
            .Append "����״��", adVarChar, 4, adFldIsNullable
            .Append "��ͥ�绰", adVarChar, 20, adFldIsNullable
            .Append "��ϵ�˵绰", adVarChar, 20, adFldIsNullable
            .Append "��λ�绰", adVarChar, 20, adFldIsNullable
            .Append "��ͥ��ַ", adVarChar, 100, adFldIsNullable
            .Append "��ͥ��ַ�ʱ�", adVarChar, 6, adFldIsNullable
            .Append "���ڵ�ַ", adVarChar, 100, adFldIsNullable
            .Append "���ڵ�ַ�ʱ�", adVarChar, 6, adFldIsNullable
            .Append "��λ�ʱ�", adVarChar, 6, adFldIsNullable
            .Append "��ϵ�˵�ַ", adVarChar, 100, adFldIsNullable
            .Append "��ϵ�˹�ϵ", adVarChar, 30, adFldIsNullable
            .Append "��ϵ������", adVarChar, 64, adFldIsNullable
        End With
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
    zlInitMEPIPati = True
End Function