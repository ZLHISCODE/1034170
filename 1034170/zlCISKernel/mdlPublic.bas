Attribute VB_Name = "mdlPublic"
Option Explicit
Public glngTXTProc As Long '保存默认的消息函数的地址
Public glngPreHWnd As Long '用于支持鼠标滚轮功能
Public glngOld As Long, glngFormW As Long, glngFormH As Long
Public Const LONG_MAX = 2147483647 'Long型最大值
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
Public Const WM_CONTEXTMENU = &H7B ' 当右击文本框时，产生这条消息
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
Public Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Public Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As PointAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As PointAPI) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWndChild As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As PointAPI) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Const MK_LBUTTON = &H1 '获取鼠标左键状态
Public Const WM_MOUSEWHEEL = &H20A

'Shell------------------------------------------------
Public Const SW_SHOWNORMAL = 1
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


'快速设置ComboBox数据
Public Const CB_ADDSTRING = &H143
Public Const CB_SETITEMDATA = &H151
Public Declare Function AddComboItem Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SetComboData Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function FindComboStr Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long

'文字API
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

'作图API
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

Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
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

Declare Function GetScrollRange Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, lpMinPos As Long, lpMaxPos As Long) As Long
    Public Const SB_HORZ = &H0
    Public Const SB_VERT = &H1

Public Const COLEditBackColor = &HE1FFE1    '浅绿
Public Const COLSelBackColor = &HFAEADA          '浅蓝

Public Function VsScroll(vsf As VSFlexGrid) As Boolean '判断水平滚动条的可见性
    Dim i As Long, lpMinPos As Long, lpMaxPos As Long
    
    VsScroll = False
    i = GetScrollRange(vsf.hwnd, SB_HORZ, lpMinPos, lpMaxPos)
    If lpMaxPos <> lpMinPos Then VsScroll = True
    End Function
    
Public Function HeScroll(vsf As VSFlexGrid) As Boolean '判断垂直滚动条的可见性
    Dim i As Long, lpMinPos As Long, lpMaxPos As Long
    
    HeScroll = False
    i = GetScrollRange(vsf.hwnd, SB_VERT, lpMinPos, lpMaxPos)
    If lpMaxPos <> lpMinPos Then HeScroll = True
End Function



Public Function MousePressButton(lngTbr As Long, objButton As Button) As Boolean
'功能：判断当前屏幕鼠标是否在指定工具按钮显示区域内按下
    Dim vRect As RECT, vPos As PointAPI
        
    '先判断当前是否处于按下状态
    If (GetKeyState(MK_LBUTTON) And &H80) <> 0 Then
        '再判断当前鼠标光标所处范围
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
'功能：判断当前屏幕鼠标是否在指定窗口的显示区域内
    Dim vRect As RECT, vPos As PointAPI
    
    GetCursorPos vPos
    GetWindowRect lngHwnd, vRect
    
    If vPos.X >= vRect.Left And vPos.X <= vRect.Right _
        And vPos.Y >= vRect.Top And vPos.Y <= vRect.Bottom Then
        MouseInRect = True
    End If
End Function

Public Sub FormSetCaption(ByVal objForm As Object, ByVal blnCaption As Boolean, Optional ByVal blnBorder As Boolean = True)
'功能：显示或隐藏一个窗体的标题栏
'参数：blnBorder=隐藏标题栏的时候,是否也隐藏窗体边框
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

Public Function MoveObj(lngHwnd As Long) As RECT
'功能：在对象的MouseDown事件中调用,对象必须具有Hwnd属性
'返回：相对屏幕的像素值
    Dim vPos As RECT
    ReleaseCapture
    SendMessage lngHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    GetWindowRect lngHwnd, vPos
    MoveObj = vPos
End Function

Public Sub PopupButtonMenu(ToolBar As Object, Button As Object, objMenu As Object)
'功能：在下拉式工具按钮中弹出一个菜单
    Dim vRect As RECT, vDot1 As PointAPI, vDot2 As PointAPI
    
    Call GetWindowRect(ToolBar.hwnd, vRect)
    vDot1.X = vRect.Left: vDot1.Y = vRect.Top
    vDot2.X = vRect.Right: vDot2.Y = vRect.Bottom
    
    Call ScreenToClient(ToolBar.Parent.hwnd, vDot1)
    Call ScreenToClient(ToolBar.Parent.hwnd, vDot2)
    
    vDot1.X = vDot1.X * 15: vDot1.Y = vDot1.Y * 15
    vDot2.X = vDot2.X * 15: vDot2.Y = vDot2.Y * 15
    ToolBar.Parent.PopupMenu objMenu, 2, vDot1.X + Button.Left, vDot2.Y
End Sub

Public Function GetControlRect(ByVal lngHwnd As Long) As RECT
'功能：获取指定控件在屏幕中的位置(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function

Public Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal lngY As Long) As PointAPI
'功能：得控件中指定坐标在屏幕中的位置(Twip)
    Dim vPoint As PointAPI
    vPoint.X = lngX / Screen.TwipsPerPixelX: vPoint.Y = lngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.X = vPoint.X * Screen.TwipsPerPixelX: vPoint.Y = vPoint.Y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function

'去掉TextBox的默认右键菜单
Public Function WndMessage(ByVal hwnd As OLE_HANDLE, ByVal Msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' 如果消息不是WM_CONTEXTMENU，就调用默认的窗口函数处理
    If Msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(glngTXTProc, hwnd, Msg, wp, lp)
End Function

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIF(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function ZVal(ByVal varValue As Variant, Optional ByVal blnForceNum As Boolean) As String
'功能：将0零转换为"NULL"串,在生成SQL语句时用
'参数：blnForceNum=当为Null时，是否强制表示为数字型
    ZVal = IIF(Val(varValue) = 0, IIF(blnForceNum, "-NULL", "NULL"), Val(varValue))
End Function

Public Function FormatEx(ByVal vNumber As Variant, ByVal intBit As Integer, _
    Optional blnShowZero As Boolean = True) As String
'功能：四舍五入方式格式化显示数字,保证小数点最后不出现0,小数点前要有0
'参数：vNumber=Single,Double,Currency类型的数字,intBit=最大小数位数
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
            If Right(strNumber, 1) = "." Then strNumber = Left(strNumber, Len(strNumber) - 1)
        End If
    End If
    FormatEx = strNumber
End Function

Public Function Decode(ParamArray arrPar() As Variant) As Variant
'功能：模拟Oracle的Decode函数
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

Public Function GetFullDate(ByVal strText As String, Optional blnTime As Boolean = True) As String
'功能：根据输入的日期简串,返回完整的日期串(yyyy-MM-dd[ HH:mm])
'参数：blnTime=是否处理时间部份
    Dim curDate As Date, strTmp As String
    
    If strText = "" Then Exit Function
    curDate = zlDatabase.Currentdate
    strTmp = strText
    
    If InStr(strTmp, "-") > 0 Or InStr(strTmp, "/") Or InStr(strTmp, ":") > 0 Then
        '输入串中包含日期分隔符
        If IsDate(strTmp) Then
            strTmp = Format(strTmp, "yyyy-MM-dd HH:mm")
            If Right(strTmp, 5) = "00:00" And InStr(strText, ":") = 0 Then
                '只输入了日期部份
                strTmp = Mid(strTmp, 1, 11) & Format(curDate, "HH:mm")
            ElseIf Left(strTmp, 10) = "1899-12-30" Then
                '只输入了时间部份
                strTmp = Format(curDate, "yyyy-MM-dd") & Right(strTmp, 6)
            End If
        Else
            '输入非法日期,返回原内容
            strTmp = strText
        End If
    Else
        '不包含日期分隔符
        If Len(strTmp) <= 2 Then
            '当作输入dd
            strTmp = Format(strTmp, "00")
            strTmp = Format(curDate, "yyyy-MM") & "-" & strTmp & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 4 Then
            '当作输入MMdd
            strTmp = Format(strTmp, "0000")
            strTmp = Format(curDate, "yyyy") & "-" & Left(strTmp, 2) & "-" & Right(strTmp, 2) & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 6 Then
            '当作输入yyMMdd
            strTmp = Format(strTmp, "000000")
            strTmp = Format(Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & "-" & Right(strTmp, 2), "yyyy-MM-dd") & " " & Format(curDate, "HH:mm")
        ElseIf Len(strTmp) <= 8 Then
            '当作输入MMddHHmm
            strTmp = Format(strTmp, "00000000")
            strTmp = Format(curDate, "yyyy") & "-" & Left(strTmp, 2) & "-" & Mid(strTmp, 3, 2) & " " & Mid(strTmp, 5, 2) & ":" & Right(strTmp, 2)
            If Not IsDate(strTmp) Then
                '当作输入yyyyMMdd
                strTmp = Format(strText, "00000000")
                strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Right(strTmp, 2) & " " & Format(curDate, "HH:mm")
            End If
        Else
            '当作输入yyyyMMddHHmm
            strTmp = Format(strTmp, "000000000000")
            strTmp = Left(strTmp, 4) & "-" & Mid(strTmp, 5, 2) & "-" & Mid(strTmp, 7, 2) & " " & Mid(strTmp, 9, 2) & ":" & Right(strTmp, 2)
        End If
    End If
    
    If IsDate(strTmp) And Not blnTime Then
        strTmp = Format(strTmp, "yyyy-MM-dd")
    End If
    GetFullDate = strTmp
End Function

Public Function NeedName(strList As String) As String
'说明:1-strList以()或[]分割编码与名称时，必须以[编码]或(编码)开头,编码必须为数字或字母
'     2-分隔符有优先级：回车符(Chr(13)）> - > [] > ()

    '优先判断以回车符分割
    If InStr(strList, Chr(13)) > 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, Chr(13)) + 1))
        Exit Function
    End If
    '以[]分割
    If InStr(strList, "]") > 0 And InStr(strList, "-") = 0 And Left(LTrim(strList), 1) = "[" Then
        If zlCommFun.IsNumOrChar(Mid(strList, 2, InStr(strList, "]") - 2)) Then
            NeedName = LTrim(Mid(strList, InStr(strList, "]") + 1))
            Exit Function
        End If
    End If
    '以()分割
    If InStr(strList, ")") > 0 And InStr(strList, "-") = 0 And Left(LTrim(strList), 1) = "(" Then
        If zlCommFun.IsNumOrChar(Mid(strList, 2, InStr(strList, ")") - 2)) Then
            NeedName = LTrim(Mid(strList, InStr(strList, ")") + 1))
            Exit Function
        End If
    End If
    '以-分割
    NeedName = LTrim(Mid(strList, InStr(strList, "-") + 1))
    
End Function

Public Sub SeekIndex(objCbo As Object, ByVal strText As String, Optional blnEvent As Boolean)
'功能：在ComboBox中查找并定位
'参数：blnEvent=定位时是否触发Click事件
'说明：未能定位时,设置ListIndex=-1
    Dim i As Long
    
    For i = 0 To objCbo.ListCount - 1
        If NeedName(objCbo.List(i)) = strText Then
            If blnEvent Then
                objCbo.ListIndex = i
            Else
                Call zlControl.CboSetIndex(objCbo.hwnd, i)
            End If
            Exit Sub
        End If
    Next
    If blnEvent Then
        objCbo.ListIndex = -1
    Else
        Call zlControl.CboSetIndex(objCbo.hwnd, -1)
    End If
End Sub

Public Function GetColFormat(vsGrid As Object) As String
'功能：获取指定表格的列格式属性串
    Dim strTmp As String, i As Long
    
    For i = 0 To vsGrid.Cols - 1
        '列宽,列可见,列对齐,列数据
        strTmp = strTmp & ";" & vsGrid.ColWidth(i) & "," & IIF(vsGrid.ColHidden(i), 0, 1) & "," & vsGrid.ColAlignment(i) & "," & Val(vsGrid.ColData(i))
    Next
    GetColFormat = Mid(strTmp, 2)
End Function

Public Sub SetColFormat(vsGrid As Object, ByVal strFormat As String)
'功能：恢复指定表格的列格式
    Dim arrCol As Variant, i As Long
    If strFormat = "" Then Exit Sub
    
    arrCol = Split(strFormat, ";")
    For i = 0 To UBound(arrCol)
        '列宽,列可见,列对齐,列数据
        vsGrid.ColWidth(i) = Val(Split(arrCol(i), ",")(0))
        vsGrid.ColHidden(i) = Val(Split(arrCol(i), ",")(1)) = 0
        vsGrid.ColAlignment(i) = Val(Split(arrCol(i), ",")(2))
        vsGrid.Cell(2, vsGrid.FixedRows, i, vsGrid.Rows - 1, i) = Val(Split(arrCol(i), ",")(2))
        
        vsGrid.ColData(i) = Val(Split(arrCol(i), ",")(3))
    Next
    vsGrid.Cell(2, 0, 0, vsGrid.FixedRows - 1, vsGrid.Cols - 1) = 4
End Sub

Public Function Between(X, a, B) As Boolean
'功能：判断x是否在a和b之间
    If a < B Then
        Between = X >= a And X <= B
    Else
        Between = X >= B And X <= a
    End If
End Function

Public Function CheckScope(varL As Double, varR As Double, varI As Double) As String
'功能：判断输入金额是否在原价和现从限定的范围内
'参数：varL=原价,varR=现价,varI=输入金额
'返回：如果不在范围内,则为提示信息,否则为空串
    If (varL >= 0 And varR >= 0) Or (varL <= 0 And varR <= 0) Then
        '如果数值符号相同,则用绝对值判断
        If Abs(varI) < Abs(varL) Or Abs(varI) > Abs(varR) Then
            CheckScope = "输入的价格绝对值不在范围(" & FormatEx(Abs(varL), 5) & "-" & FormatEx(Abs(varR), 5) & ")内."
        End If
    Else
        '如果符号不相同,则用原始范围判断
        If varI < varL Or varI > varR Then
            CheckScope = "输入的价格值不在范围(" & FormatEx(varL, 5) & "-" & FormatEx(varR, 5) & ")内."
        End If
    End If
End Function

Public Sub FindCboIndex(objCbo As Object, lngData As Long, Optional Keep As Boolean)
'功能：由项目值查找ComboBox的项目索引
'参数：Keep=如果未匹配，是否保持原索引
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
'功能：由字符串在ComboBox中查找索引
'参数：Keep=如果未匹配，是否保持原索引
    Dim i As Integer
    
    '先精确查找
    For i = 0 To objCbo.ListCount - 1
        If objCbo.List(i) = strFind Then
            objCbo.ListIndex = i: Exit Sub
        ElseIf NeedName(objCbo.List(i)) = strFind And strFind <> "" Then
            objCbo.ListIndex = i: Exit Sub
        End If
    Next
    
    '最后模糊查找
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
'功能：由ItemData或Text查找ComboBox的索引值
    Dim strType As String, i As Integer
    
    SeekCboIndex = -1
    
    strType = TypeName(varData)
    If strType = "Field" Then
        If IsType(varData.Type, adVarChar) Then strType = "String"
    End If
    
    If strType = "String" Then
        If varData <> "" Then
            '先精确查找
            For i = 0 To objCbo.ListCount - 1
                If objCbo.List(i) = varData Then
                    SeekCboIndex = i: Exit Function
                ElseIf NeedName(objCbo.List(i)) = varData And varData <> "" Then
                    SeekCboIndex = i: Exit Function
                End If
            Next
            '再模糊查找
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

Public Function GetNextControl(ByVal intTab As Integer, ByVal frmForm As Object, Optional ByVal strSkip As String) As Object
'功能：获取下一个光标顺序的控件
'参数：strSkip=要跳过光标的控件名,以";"间隔
    Dim objNext As Object, i As Long
    
    '先找比当前控件TabIndex大的
    For i = 0 To frmForm.Controls.Count - 1
        If InStr("TextBox,ComboBox,VSFlexGrid", TypeName(frmForm.Controls(i))) > 0 Then
            If frmForm.Controls(i).Enabled And frmForm.Controls(i).Visible And frmForm.Controls(i).TabStop _
                And (InStr(";" & strSkip & ";", ";" & frmForm.Controls(i).Name & ";") = 0 Or strSkip = "") Then
                If frmForm.Controls(i).TabIndex > intTab Then
                    If objNext Is Nothing Then
                        Set objNext = frmForm.Controls(i)
                    ElseIf frmForm.Controls(i).TabIndex < objNext.TabIndex Then
                        Set objNext = frmForm.Controls(i)
                    End If
                End If
            End If
        End If
    Next
    If objNext Is Nothing Then
        '没有则找比当前控件TabIndex小的
        For i = 0 To frmForm.Controls.Count - 1
            If InStr("TextBox,ComboBox,VSFlexGrid", TypeName(frmForm.Controls(i))) > 0 Then
                If frmForm.Controls(i).Enabled And frmForm.Controls(i).Visible And frmForm.Controls(i).TabStop _
                    And (InStr(";" & strSkip & ";", ";" & frmForm.Controls(i).Name & ";") = 0 Or strSkip = "") Then
                    If frmForm.Controls(i).TabIndex < intTab Then
                        If objNext Is Nothing Then
                            Set objNext = frmForm.Controls(i)
                        ElseIf frmForm.Controls(i).TabIndex < objNext.TabIndex Then
                            Set objNext = frmForm.Controls(i)
                        End If
                    End If
                End If
            End If
        Next
    End If
    Set GetNextControl = objNext
End Function

Public Function PreFixNO(Optional curDate As Date = #1/1/1900#) As String
'功能：返回大写的单据号年前缀
    If curDate = #1/1/1900# Then
        PreFixNO = CStr(CInt(Format(zlDatabase.Currentdate, "YYYY")) - 1990)
    Else
        PreFixNO = CStr(CInt(Format(curDate, "YYYY")) - 1990)
    End If
    PreFixNO = IIF(CInt(PreFixNO) < 10, PreFixNO, Chr(55 + CInt(PreFixNO)))
End Function

Public Function InDesign() As Boolean
    On Error Resume Next
    Debug.Print 1 / 0
    If err.Number <> 0 Then err.Clear: InDesign = True
End Function

Public Function Custom_WndMessage(ByVal hwnd As Long, ByVal Msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
'功能：自定义消息函数处理窗体尺寸调整限制
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
    Custom_WndMessage = CallWindowProc(glngOld, hwnd, Msg, wp, lp)
End Function

Public Function CheckAdviceWindow(ByVal strTitle As String) As Boolean
'功能：检查医嘱编辑窗口是否已经打开
    Dim lngHwnd As Long
    
    '其它窗口打开了
    lngHwnd = FindWindow("ThunderFormDC", strTitle)
    If lngHwnd = 0 Then
        lngHwnd = FindWindow("ThunderRT6FormDC", strTitle)
    End If
    If lngHwnd <> 0 Then
        MsgBox "医嘱编辑窗口已经打开，请先完成当前操作后再执行。", vbInformation, gstrSysName
        Call ShowWindow(lngHwnd, SW_RESTORE)
        Call BringWindowToTop(lngHwnd)
        Exit Function
    End If
    CheckAdviceWindow = True
End Function

Public Function IsType(ByVal varType As DataTypeEnum, ByVal varBase As DataTypeEnum) As Boolean
'功能：判断某个ADO字段数据类型是否与指定字段类型是同一类(如数字,日期,字符,二进制)
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

Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '功能： 密码转换函数
    '参数：
    '   strOld：原密码
    '返回： 加密生成的密码
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

Public Function AnalyseComputer() As String
    Dim strComputer As String * 256
    Call GetComputerName(strComputer, 255)
    AnalyseComputer = strComputer
    AnalyseComputer = Replace(AnalyseComputer, Chr(0), "")
End Function

Public Function SetBit(ByVal strBit As String, ByVal intBit As Integer, Optional ByVal intVal As Integer = -1) As String
'功能：将指定位字符串strBit中的第intBit位设置为0或1
'参数：intVal=设置值,0或1,不传表示反转
    If intVal = -1 Then intVal = IIF(Val(Mid(strBit, intBit, 1)) = 0, 1, 0)
    SetBit = Left(strBit, intBit - 1) & intVal & Mid(strBit, intBit + 1)
End Function

Public Function BeSureMode(ByVal strMode1 As String, strMode2 As String, strMode3 As String) As Integer
'功能：识别三种报警模式的优先级
    If strMode1 = "-" Then BeSureMode = 1: Exit Function
    If strMode2 = "-" Then BeSureMode = 2: Exit Function
    If strMode3 = "-" Then BeSureMode = 3: Exit Function
    If strMode1 <> "" And strMode2 = "" And strMode3 = "" Then BeSureMode = 1: Exit Function
    If strMode1 = "" And strMode2 <> "" And strMode3 = "" Then BeSureMode = 2: Exit Function
    If strMode1 = "" And strMode2 = "" And strMode3 <> "" Then BeSureMode = 3: Exit Function
    If strMode3 <> "" Then BeSureMode = 3: Exit Function
    If strMode2 <> "" Then BeSureMode = 2: Exit Function
    If strMode1 <> "" Then BeSureMode = 1: Exit Function
End Function

Public Function RowIsLastVisible(ByRef vsf As VSFlexGrid, ByVal lngRow As Long) As Boolean
'功能：判断指定行是否最后一可见行
    Dim i As Long
    
    With vsf
        For i = .Rows - 1 To .FixedRows Step -1
            If Not .RowHidden(i) Then Exit For
        Next
        If i >= .FixedRows Then
            RowIsLastVisible = lngRow = i
        End If
    End With
End Function

Public Function VsfOnlyOneRow(ByRef vsf As VSFlexGrid) As Boolean
'功能：判断是否仅有一行可见行
    Dim i As Long, k As Long
    
    With vsf
        For i = .Rows - 1 To .FixedRows Step -1
            If Not .RowHidden(i) Then
                k = k + 1
                If k > 1 Then Exit For
            End If
        Next
        VsfOnlyOneRow = Not (k > 1)
    End With
End Function
Public Function zlCheckPrivs(ByVal strPrivs As String, ByVal strMyPriv As String) As Boolean
    '---------------------------------------------------------------------------------------------
    '功能:检查指定的权限是否存在
    '参数:strPrivs-权限串
    '     strMyPriv-具体权限
    '返回,存在权限,返回true,否则返回False
    '编制:刘兴宏
    '日期:2008/01/09
    '---------------------------------------------------------------------------------------------
    zlCheckPrivs = InStr(";" & strPrivs & ";", ";" & strMyPriv & ";") > 0
End Function

Public Function GetPriceType(ByVal lng病人ID As Long, ByVal lng收费细目ID As Long, ByVal int险类 As Integer, ByVal bln门诊 As Boolean) As String
'功能：获取对应医保费用类型
    Dim str类型 As String
    
    '取医保的费用类型
    If int险类 <> 0 Then
        On Error Resume Next
        str类型 = gclsInsure.GetItemInsure(lng病人ID, lng收费细目ID, 0, bln门诊, int险类)
        If str类型 <> "" Then
            If UBound(Split(str类型, ";")) >= 5 Then
                str类型 = Split(str类型, ";")(5)
            Else
                str类型 = ""
            End If
        End If
    End If
    GetPriceType = str类型
End Function

Public Sub SetCboFromList(ByRef objControl As Object, ByVal arrList As Variant, ByVal arrCboIdx As Variant, Optional ByVal intDefault As Integer = -1)
'功能：将指定数据装入指定ComboBox
'参数：arrList=List String数组
'      arrCboIdx=ComboBox索引数组,多个ComboBox时,装入数据相同
'      intDefaut=缺省索引
    Dim i As Long, j As Long
    
    For i = 0 To UBound(arrCboIdx)
        objControl(arrCboIdx(i)).Clear
        For j = 0 To UBound(arrList)
            objControl(arrCboIdx(i)).AddItem arrList(j)
        Next
        objControl(arrCboIdx(i)).ListIndex = intDefault '缺省为未选中
    Next
End Sub

Public Function DynamicCreate(ByVal strclass As String, ByVal strCaption As String, Optional ByVal blnMsg As Boolean) As Object
'动态创建对象
    On Error Resume Next
    Set DynamicCreate = CreateObject(strclass)
    
    If err <> 0 Then
        If blnMsg Then MsgBox strCaption & "组件创建失败，请联系管理员检查是否正确安装!", vbInformation, gstrSysName
        Set DynamicCreate = Nothing
    End If
    err.Clear
End Function



