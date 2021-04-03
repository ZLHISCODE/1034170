Attribute VB_Name = "mdlAPI"
Option Explicit
'--------------------------------------------------------
'功  能：本模块用于存储API调用的各种函数
'编制人：赵彤宇
'编制日期：2004.6
'过程函数清单：
'       ShowTitle() 设置窗体是否显示标题栏
'       RaisEffect() 将PictureBox模拟成3D平面按钮
'修改记录：
'
'-------------------------------------------------------
Public frmMain As frmViewer
''建立子目录
Public Declare Function CreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''处理鼠标滚轮
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
Public plngFilmPreWndProc As Long       'Film窗体原来的消息处理程序
Public plngFilmViewPreWndProc As Long       'Film窗体原来的消息处理程序

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''给Pic画凹凸使用
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
Public Const BDR_SUNKENOUTER = &H2 '浅凹下
Public Const BDR_RAISEDINNER = &H4 '浅凸起
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER) '深凸起
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER) '深凹下
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER) 'frame边线样式
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER) '反frame边线样式
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
''''''''''''''''''''''''''''[放大镜使用]''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const SM_CYMENU = 15
Public Const SM_CYCAPTION = 4

'判断数组是否为空
Public Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'使用API函数修改MsgBox，使其可以在调用的时候，指定父窗体
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

'使用主板播放声音
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
'功能： 设置窗体是否显示标题栏
'参数： f－－需要处理的窗体；ShowTitle－－是否显示窗体：True显示标题；Fasle不显示标题
'返回：无，直接修改窗体f的显示效果
'编制人：赵彤宇
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
'功能： 将PictureBox模拟成3D平面按钮
'参数： picBox－－需要做处理的PictureBox；intStyle－－0=平面,-1=凹下,1=凸起；strname－－定位X,Y坐标用的字符
'返回：无，直接修改picBox的显示效果
'编制人：赵彤宇
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
    ' 取出32位值的高16位
    HIWORD = (LongIn And &HFFFF0000) \ &H10000
End Function

Public Function LOWORD(LongIn As Long) As Integer
    ' 取出32位值的低16位
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
            If Sgn(wzDelta) = 1 Then    '鼠标上滚
                Call frmMain.MouseWheel(1)
            Else                        '鼠标下滚
                Call frmMain.MouseWheel(0)
            End If
    End Select
    Wndproc = CallWindowProc(preWinProc, hwnd, Msg, wParam, lParam)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''
''为了处理双屏时对话框的正确显示位置，用API函数改写了一下MsgBox函数
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
'功能：显示或隐藏一个窗体的标题栏
'参数：blnBorder=隐藏标题栏的时候,是否也隐藏窗体边框
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
    '状态发生改变，才真正改变窗体
    If lngOldStyle <> lngStyle Then
        SetWindowLong objForm.hwnd, GWL_STYLE, lngStyle
        SetWindowPos objForm.hwnd, 0, vRect.left, vRect.top, vRect.Right - vRect.left, vRect.Bottom - vRect.top, SWP_NOREPOSITION Or SWP_FRAMECHANGED Or SWP_NOZORDER
    End If
End Sub

Public Function FilmHook(ByVal hwnd As Long) As Long
    '返回并保存原来默认的窗口过程指针
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
'功能：胶片打印窗口的windows消息处理程序，专门处理鼠标滚轮 消息
'参数：
'返回：
'------------------------------------------------
    Dim pt As POINTL
    Dim wzDelta, wKeys As Integer

    wzDelta = HIWORD(wParam)
    wKeys = LOWORD(wParam)

    If uMsg = WM_MOUSEWHEEL Then
        If Not frmMain.mfrmFilm Is Nothing Then
            If Sgn(wzDelta) = 1 Then    '鼠标上滚
                Call frmMain.mfrmFilm.MouseWheel(1)
            Else                        '鼠标下滚
                Call frmMain.mfrmFilm.MouseWheel(0)
            End If
        End If
    End If
  
    '调用原来的窗口过程
    FilmWindowProc = CallWindowProc(plngFilmPreWndProc, hw, uMsg, wParam, lParam)
End Function

Public Function FilmViewHook(ByVal hwnd As Long) As Long
    '返回并保存原来默认的窗口过程指针
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
'功能：胶片打印窗口的windows消息处理程序，专门处理鼠标滚轮 消息
'参数：
'返回：
'------------------------------------------------
    Dim pt As POINTL
    Dim wzDelta, wKeys As Integer

    wzDelta = HIWORD(wParam)
    wKeys = LOWORD(wParam)

    If uMsg = WM_MOUSEWHEEL Then
        If Not frmMain.mfrmFilm Is Nothing Then
            If Not frmMain.mfrmFilm.mfrmFilmView Is Nothing Then
                If Sgn(wzDelta) = 1 Then    '鼠标上滚
                    Call frmMain.mfrmFilm.mfrmFilmView.MouseWheel(1)
                Else                        '鼠标下滚
                    Call frmMain.mfrmFilm.mfrmFilmView.MouseWheel(0)
                End If
            End If
        End If
    End If
  
    '调用原来的窗口过程
    FilmViewWindowProc = CallWindowProc(plngFilmViewPreWndProc, hw, uMsg, wParam, lParam)
End Function
