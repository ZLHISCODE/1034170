Attribute VB_Name = "mdlAPI"
Option Explicit
'--------------------------------------------------------
'��  �ܣ���ģ�����ڴ洢API���õĸ��ֺ���
'�����ˣ���ͮ��
'�������ڣ�2004.6
'���̺����嵥��
'       ShowTitle() ���ô����Ƿ���ʾ������
'       RaisEffect() ��PictureBoxģ���3Dƽ�水ť
'�޸ļ�¼��
'
'-------------------------------------------------------
Public frmMain As frmViewer
''������Ŀ¼
Public Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''����������
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const WM_MOUSEWHEEL = &H20A
Public Const GWL_WNDPROC = (-4)
Public Type POINTL
    x As Long
    y As Long
End Type
Public preWinProc As Long
Public plngFilmPreWndProc As Long       'Film����ԭ������Ϣ�������
Public plngFilmViewPreWndProc As Long       'Film����ԭ������Ϣ�������

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''��Pic����͹ʹ��
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Type RECT
        left As Long
        top As Long
        Right As Long
        Bottom As Long
End Type


Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2 'ǳ����
Public Const BDR_RAISEDINNER = &H4 'ǳ͹��
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER) '��͹��
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER) '���
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER) 'frame������ʽ
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER) '��frame������ʽ
Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_SOFT = &H1000

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Const GWL_STYLE = (-16)
Private Const WS_CAPTION = &HC00000
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOSIZE = &H1

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''[�Ŵ�ʹ��]''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const SM_CYMENU = 15
Public Const SM_CYCAPTION = 4

'�ж������Ƿ�Ϊ��
Public Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'ʹ��API�����޸�MsgBox��ʹ������ڵ��õ�ʱ��ָ��������
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Const MB_ABORTRETRYIGNORE = &H2&
Public Const MB_APPLMODAL = &H0&
Public Const MB_COMPOSITE = &H2         '  use composite chars
Public Const MB_DEFAULT_DESKTOP_ONLY = &H20000
Public Const MB_DEFBUTTON1 = &H0&
Public Const MB_DEFBUTTON2 = &H100&
Public Const MB_DEFBUTTON3 = &H200&
Public Const MB_DEFMASK = &HF00&
Public Const MB_ICONASTERISK = &H40&
Public Const MB_ICONEXCLAMATION = &H30&
Public Const MB_ICONHAND = &H10&
Public Const MB_ICONINFORMATION = MB_ICONASTERISK
Public Const MB_ICONMASK = &HF0&
Public Const MB_ICONQUESTION = &H20&
Public Const MB_ICONSTOP = MB_ICONHAND
Public Const MB_MISCMASK = &HC000&
Public Const MB_MODEMASK = &H3000&
Public Const MB_NOFOCUS = &H8000&
Public Const MB_OK = &H0&
Public Const MB_OKCANCEL = &H1&
Public Const MB_PRECOMPOSED = &H1         '  use precomposed chars
Public Const MB_RETRYCANCEL = &H5&
Public Const MB_SETFOREGROUND = &H10000
Public Const MB_SYSTEMMODAL = &H1000&
Public Const MB_TASKMODAL = &H2000&
Public Const MB_TYPEMASK = &HF&
Public Const MB_USEGLYPHCHARS = &H4         '  use glyph chars, not ctrl chars
Public Const MB_YESNO = &H4&
Public Const MB_YESNOCANCEL = &H3&

Public Const WS_THICKFRAME = &H40000
Public Const WS_SYSMENU = &H80000
Public Const WS_MAXIMIZEBOX = &H10000
Public Const WS_MINIMIZEBOX = &H20000
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

'ʹ�����岥������
Public Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Public Const BEEP_Do0 = 264
Public Const BEEP_Re = 297
Public Const BEEP_Mi = 330
Public Const BEEP_Fa = 352
Public Const BEEP_Sol = 396
Public Const BEEP_la = 440
Public Const BEEP_Ti = 495
Public Const BEEP_Do1 = 528


Public Sub ToggleTitleBar(f As Form, ShowTitle As Boolean)
'------------------------------------------------
'���ܣ� ���ô����Ƿ���ʾ������
'������ f������Ҫ����Ĵ��壻ShowTitle�����Ƿ���ʾ���壺True��ʾ���⣻Fasle����ʾ����
'���أ��ޣ�ֱ���޸Ĵ���f����ʾЧ��
'�����ˣ���ͮ��
'------------------------------------------------
    Dim style As Long
    style = GetWindowLong(f.hwnd, GWL_STYLE)
    If ShowTitle Then
        style = style Or WS_CAPTION
    Else
        style = style And Not WS_CAPTION
    End If
    SetWindowLong f.hwnd, GWL_STYLE, style
    SetWindowPos f.hwnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE
End Sub


Public Sub RaisEffect(picBox As PictureBox, Optional intStyle As Integer, Optional strname As String = "")
'------------------------------------------------
'���ܣ� ��PictureBoxģ���3Dƽ�水ť
'������ picBox������Ҫ�������PictureBox��intStyle����0=ƽ��,-1=����,1=͹��strname������λX,Y�����õ��ַ�
'���أ��ޣ�ֱ���޸�picBox����ʾЧ��
'�����ˣ���ͮ��
'------------------------------------------------

    Dim PicRect As RECT
    Dim lngTmp As Long
    With picBox
        lngTmp = .ScaleMode
        .ScaleMode = 3
        .BorderStyle = 0
        If intStyle <> 0 Then
            PicRect.left = .ScaleLeft
            PicRect.top = .ScaleTop
            PicRect.Right = .ScaleWidth * 2
            PicRect.Bottom = .ScaleHeight * 2
            DrawEdge .hDC, PicRect, CLng(IIf(intStyle = 1, BDR_RAISEDINNER Or BF_SOFT, BDR_SUNKENOUTER Or BF_SOFT)), BF_RECT
        End If
        .ScaleMode = lngTmp
        If strname <> "" Then
            .CurrentX = (.ScaleWidth - .TextWidth(strname)) / 2
            .CurrentY = (.ScaleHeight - .TextHeight(strname)) / 2
        End If
    End With
End Sub

Public Function HIWORD(LongIn As Long) As Integer
    ' ȡ��32λֵ�ĸ�16λ
    HIWORD = (LongIn And &HFFFF0000) \ &H10000
End Function

Public Function LOWORD(LongIn As Long) As Integer
    ' ȡ��32λֵ�ĵ�16λ
    If (LongIn And &HFFFF&) > &H7FFF Then
        LOWORD = (LongIn And &HFFFF&) - &H10000
    Else
        LOWORD = LongIn And &HFFFF&
    End If
End Function
Public Function Wndproc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim pt As POINTL
    Dim wzDelta, wKeys As Integer
    On Error Resume Next
    wzDelta = HIWORD(wParam)
    wKeys = LOWORD(wParam)
    Select Case Msg
        Case WM_MOUSEWHEEL
            If Sgn(wzDelta) = 1 Then    '����Ϲ�
                Call frmMain.MouseWheel(1)
            Else                        '����¹�
                Call frmMain.MouseWheel(0)
            End If
    End Select
    Wndproc = CallWindowProc(preWinProc, hwnd, Msg, wParam, lParam)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''
''Ϊ�˴���˫��ʱ�Ի������ȷ��ʾλ�ã���API������д��һ��MsgBox����
'''''''''''''''''''''''''''''''''''''''''''''''''
Public Function MsgBox(ByVal Prompt As String, Optional ByVal Buttons As VbMsgBoxStyle = MB_OK, _
    Optional Title As String = "", Optional frmParent As Object = Nothing) As Long
    If Not frmParent Is Nothing Then
        MsgBox = MessageBox(frmParent.hwnd, Prompt, Title, Buttons)
    Else
        MsgBox = MessageBox(frmMain.hwnd, Prompt, Title, Buttons)
    End If

End Function


Public Sub FormSetCaption(ByVal objForm As Object, ByVal blnCaption As Boolean, Optional ByVal blnBorder As Boolean = True)
'���ܣ���ʾ������һ������ı�����
'������blnBorder=���ر�������ʱ��,�Ƿ�Ҳ���ش���߿�
    Dim vRect As RECT, lngStyle As Long
    Dim lngOldStyle As Long
    
    Call GetWindowRect(objForm.hwnd, vRect)
    lngStyle = GetWindowLong(objForm.hwnd, GWL_STYLE)
    lngOldStyle = lngStyle
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
    '״̬�����ı䣬�������ı䴰��
    If lngOldStyle <> lngStyle Then
        SetWindowLong objForm.hwnd, GWL_STYLE, lngStyle
        SetWindowPos objForm.hwnd, 0, vRect.left, vRect.top, vRect.Right - vRect.left, vRect.Bottom - vRect.top, SWP_NOREPOSITION Or SWP_FRAMECHANGED Or SWP_NOZORDER
    End If
End Sub

Public Function FilmHook(ByVal hwnd As Long) As Long
    '���ز�����ԭ��Ĭ�ϵĴ��ڹ���ָ��
    If App.LogMode <> 0 Then
        FilmHook = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf FilmWindowProc)
    End If
End Function

Public Sub FilmUnhook(ByVal hwnd As Long, ByVal lpWndProc As Long)
  Dim temp As Long
  
    If App.LogMode <> 0 Then
        temp = SetWindowLong(hwnd, GWL_WNDPROC, lpWndProc)
    End If
End Sub

Function FilmWindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'------------------------------------------------
'���ܣ���Ƭ��ӡ���ڵ�windows��Ϣ�������ר�Ŵ��������� ��Ϣ
'������
'���أ�
'------------------------------------------------
    Dim pt As POINTL
    Dim wzDelta, wKeys As Integer

    wzDelta = HIWORD(wParam)
    wKeys = LOWORD(wParam)

    If uMsg = WM_MOUSEWHEEL Then
        If Not frmMain.mfrmFilm Is Nothing Then
            If Sgn(wzDelta) = 1 Then    '����Ϲ�
                Call frmMain.mfrmFilm.MouseWheel(1)
            Else                        '����¹�
                Call frmMain.mfrmFilm.MouseWheel(0)
            End If
        End If
    End If
  
    '����ԭ���Ĵ��ڹ���
    FilmWindowProc = CallWindowProc(plngFilmPreWndProc, hw, uMsg, wParam, lParam)
End Function

Public Function FilmViewHook(ByVal hwnd As Long) As Long
    '���ز�����ԭ��Ĭ�ϵĴ��ڹ���ָ��
    If App.LogMode <> 0 Then
        FilmViewHook = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf FilmViewWindowProc)
    End If
End Function

Public Sub FilmViewUnhook(ByVal hwnd As Long, ByVal lpWndProc As Long)
  Dim temp As Long
  
    If App.LogMode <> 0 Then
        temp = SetWindowLong(hwnd, GWL_WNDPROC, lpWndProc)
    End If
End Sub

Function FilmViewWindowProc(ByVal hw As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'------------------------------------------------
'���ܣ���Ƭ��ӡ���ڵ�windows��Ϣ�������ר�Ŵ��������� ��Ϣ
'������
'���أ�
'------------------------------------------------
    Dim pt As POINTL
    Dim wzDelta, wKeys As Integer

    wzDelta = HIWORD(wParam)
    wKeys = LOWORD(wParam)

    If uMsg = WM_MOUSEWHEEL Then
        If Not frmMain.mfrmFilm Is Nothing Then
            If Not frmMain.mfrmFilm.mfrmFilmView Is Nothing Then
                If Sgn(wzDelta) = 1 Then    '����Ϲ�
                    Call frmMain.mfrmFilm.mfrmFilmView.MouseWheel(1)
                Else                        '����¹�
                    Call frmMain.mfrmFilm.mfrmFilmView.MouseWheel(0)
                End If
            End If
        End If
    End If
  
    '����ԭ���Ĵ��ڹ���
    FilmViewWindowProc = CallWindowProc(plngFilmViewPreWndProc, hw, uMsg, wParam, lParam)
End Function
