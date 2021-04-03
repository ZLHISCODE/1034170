Attribute VB_Name = "mdlCustomCommon"
Option Explicit
Public Enum gRegType
    g注册信息 = 0
    g公共全局 = 1
    g公共模块 = 2
    g私有全局 = 3
    g私有模块 = 4
    g本机公共模块 = 5
    g本机私有模块 = 6
End Enum
Public gstrMatchMethod As String
Public Const SPI_GETWORKAREA = 48
'系统方案设置----------------------------------
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
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT) As Long
Public Const GCTRL_SELBACK_COLOR = &H8000000D

Private Const GWL_STYLE As Long = (-16&)
Private Const GWL_EXSTYLE As Long = (-20&)
Private Const WS_THICKFRAME = &H40000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_MAXIMIZEBOX = &H10000

Private Declare Function GetCursorPos& Lib "user32" (lpPoint As POINTAPI)
Private Declare Function SetCursorPos& Lib "user32" (ByVal X&, ByVal Y&)
Private Declare Function ClientToScreen& Lib "user32" (ByVal Hwnd&, lpPoint As POINTAPI)
Private Declare Function GetSystemMenu& Lib "user32" (ByVal Hwnd&, ByVal bRevert&)
Private Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd&, ByVal nIndex&)
Private Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd&, ByVal nIndex&, ByVal dwNewLong&)

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

Public Sub zlSetWindowsBroldStyle(ByVal frmMain As Form)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:改变可调窗体不不可调窗体（即设置只有关闭按钮窗口,如果窗体本身只有关闭，只会自动加上最大化、最小化等按钮)
    '入参:frmMain.hwnd-窗体的句柄
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-09-29 16:10:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim pt_SavePoint As POINTAPI, pt_MovePoint As POINTAPI
    Err = 0: On Error GoTo Errhand:
    With pt_MovePoint
      .X = (-1): .Y = 10
    End With
    '设置窗体的broldStyle
    Call SetWindowLong(frmMain.Hwnd, GWL_STYLE, GetWindowLong(frmMain.Hwnd, GWL_STYLE) Xor _
                              (WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX))
    Call GetSystemMenu(frmMain.Hwnd, 1&)
    '必需重画数据
    With frmMain
        .Move .Left, .Top, .Width - 15, .Height - 15
        .Move .Left, .Top, .Width + 15, .Height + 15
    End With
    Call GetCursorPos(pt_SavePoint)
    Call ClientToScreen(frmMain.Hwnd, pt_MovePoint)
    Call SetCursorPos(pt_MovePoint.X, pt_MovePoint.Y)
    Call SetCursorPos(pt_SavePoint.X, pt_SavePoint.Y)
Errhand:
End Sub
Public Sub zlSetCrlEnbled(ByVal objCrl As Object, blnEnabled As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置指定控件的Nabled属性,如果为False,同时需要设置相关的背景色
    '入参:objCrl-转入的指定控件
    '     blnEnabled-相关属性
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-09-08 14:44:25
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Select Case UCase(TypeName(objCrl))
    Case UCase("TextBox"), UCase("COMBOBOX")
        objCrl.Enabled = blnEnabled
        zlSetCtrolBackColor objCrl
    Case UCase("dtpicker"), UCase("frame"), UCase("CHECKBOX"), UCase("LABEL"), UCase("COMMANDBUTTON")
        objCrl.Enabled = blnEnabled
    Case Else
       ' objCrl.Enabled = blnEnabled
    End Select
End Sub
Public Sub zlSetCtrolBackColor(ByVal objCtl As Object)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件背景色的颜色
    '入参:objCtl-转入的控件
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-09-08 14:43:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If objCtl.Enabled = False Then
        objCtl.BackColor = &H8000000F
    Else
        objCtl.BackColor = vbWhite
    End If
End Sub
Public Sub zlCtlSetFocus(ByVal objCtl As Object, Optional blnDoEvnts As Boolean = False)
    '功能:将集点移动控件中:2008-07-08 16:48:35
    Err = 0: On Error Resume Next
    If blnDoEvnts Then DoEvents
    If objCtl.Visible And objCtl.Enabled = True Then: objCtl.SetFocus
End Sub
Public Function zlSaveDockPanceToReg(ByVal frmMain As Form, ByVal objPance As DockingPane, _
                ByVal strKey As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:保存DockPane控件的具体位置
    '入参:frmMain-窗体名
    '     objPance:DockinPane控件
    '      StrKey-键名
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-02-10 14:24:04
    '-----------------------------------------------------------------------------------------------------------
    Dim blnAutoHide As Boolean
    If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then
        zlSaveDockPanceToReg = True: Exit Function
    End If
    Err = 0: On Error GoTo Errhand:
    objPance.SaveState "VB and VBA Program Settings\ZLSOFT\私有模块\" & gstrDBUser & "\" & App.ProductName, frmMain.Name, "区域"
    zlSaveDockPanceToReg = True
Errhand:
End Function

Public Function zlRestoreDockPanceToReg(ByVal frmMain As Form, ByVal objPance As DockingPane, _
                ByVal strKey As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:保存DockPane控件的具体位置
    '入参:frmMain-窗体名
    '     objPance:DockinPane控件
    '      StrKey-键名
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2009-02-10 14:24:04
    '-----------------------------------------------------------------------------------------------------------
    Dim blnAutoHide As Boolean
    If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then
        zlRestoreDockPanceToReg = True: Exit Function
    End If
    'blnAutoHide = Val(zlDatabase.GetPara("界面区域隐藏", , , True)) = 1
    Err = 0: On Error GoTo Errhand:
    objPance.LoadState "VB and VBA Program Settings\ZLSOFT\私有模块\" & gstrDBUser & "\" & App.ProductName, frmMain.Name, "区域"
    zlRestoreDockPanceToReg = True
Errhand:
End Function
Public Sub ShowMsgbox(ByVal strMsgInfor As String, Optional blnYesNo As Boolean = False, Optional ByRef blnYes As Boolean)
    '----------------------------------------------------------------------------------------------------------------
    '功能：提示消息框
    '参数：strMsgInfor-提示信息
    '     blnYesNo-是否提供YES或NO按钮
    '返回：blnYes-如果提供YESNO按钮,则返回YES(True)或NO(False)
    '----------------------------------------------------------------------------------------------------------------
        
    If blnYesNo = False Then
        MsgBox strMsgInfor, vbInformation + vbDefaultButton1, gstrSysName
    Else
        blnYes = MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
    End If
End Sub
Public Function IsHavePrivs(ByVal strPrivs As String, ByVal strMyPriv As String) As Boolean
    '---------------------------------------------------------------------------------------------
    '功能:检查指定的权限是否存在
    '参数:strPrivs-权限串
    '     strMyPriv-具体权限
    '返回,存在权限,返回true,否则返回False
    '编制:刘兴宏
    '日期:2008/01/09
    '---------------------------------------------------------------------------------------------
    IsHavePrivs = InStr(";" & strPrivs & ";", ";" & strMyPriv & ";") > 0
End Function
Public Sub SaveRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByVal strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '功能:  将指定的信息保存在注册表中
    '参数:  RegType-注册类型
    '       strSection-注册表目录
    '       StrKey-键名
    '       strKeyValue-键值
    '返回:
    '--------------------------------------------------------------------------------------------------------------
    Err = 0
    On Error GoTo Errhand:
    Select Case RegType
        Case g注册信息
            SaveSetting "ZLSOFT", "注册信息\" & strSection, strKey, strKeyValue
        Case g公共全局
            SaveSetting "ZLSOFT", "公共全局\" & strSection, strKey, strKeyValue
        Case g公共模块
            SaveSetting "ZLSOFT", "公共模块" & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
        Case g私有全局
            SaveSetting "ZLSOFT", "私有全局\" & gstrDBUser & "\" & strSection, strKey, strKeyValue
        Case g私有模块
            SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
    End Select
Errhand:
End Sub
Public Sub GetRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByRef strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '功能:  将指定的注册信息读取出来
    '入参数:  RegType-注册类型
    '       strSection-注册表目录
    '       StrKey-键名
    '出参数:
    '       strKeyValue-返回的键值
    '返回:
    '--------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Err = 0
    On Error GoTo Errhand:
    Select Case RegType
        Case g注册信息
            SaveSetting "ZLSOFT", "注册信息\" & strSection, strKey, strKeyValue
            strKeyValue = GetSetting("ZLSOFT", "注册信息\" & strSection, strKey, "")
        Case g公共全局
            strKeyValue = GetSetting("ZLSOFT", "公共全局\" & strSection, strKey, "")
        Case g公共模块
            strKeyValue = GetSetting("ZLSOFT", "公共模块" & "\" & App.ProductName & "\" & strSection, strKey, "")
        Case g私有全局
            strKeyValue = GetSetting("ZLSOFT", "私有全局\" & gstrDBUser & "\" & strSection, strKey, "")
        Case g私有模块
            strKeyValue = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, "")
    End Select
Errhand:
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
Public Function GetMatchingSting(ByVal strString As String, Optional blnUpper As Boolean = True) As String
    '--------------------------------------------------------------------------------------------------------------------------------------
    '功能:加入匹配串%
    '参数:strString 需匹配的字串
    '     blnUpper-是否转换在大写
    '返回:返回加匹配串%dd%
    '--------------------------------------------------------------------------------------------------------------------------------------
    Dim strLeft As String
    Dim strRight As String
    If gstrMatchMethod = "" Then
        gstrMatchMethod = Val(zlDatabase.GetPara("输入匹配"))
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

Public Function GetTaskbarHeight() As Integer
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取任务栏高度
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-08-28 18:38:30
    '-----------------------------------------------------------------------------------------------------------
    Dim lRes As Long
    Dim vRect As RECT
    Err = 0: On Error GoTo Errhand:
    lRes = SystemParametersInfo(SPI_GETWORKAREA, 0, vRect, 0)
    GetTaskbarHeight = ((Screen.Height / Screen.TwipsPerPixelX) - vRect.Bottom) * Screen.TwipsPerPixelX
Errhand:
End Function
Public Function zlGetNextWeekDate(Optional strDate As String = "") As Date
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取下周一函数
    '返回:
    '编制:刘兴洪
    '日期:2009-09-21 11:19:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dtDate As Date
    If strDate <> "" Then
        dtDate = DateAdd("d", 7, CDate(strDate))
    Else
        dtDate = DateAdd("d", 7, zlDatabase.Currentdate)
    End If
    dtDate = DateAdd("d", -1 * (Weekday(dtDate, vbMonday) - 1), dtDate)
    zlGetNextWeekDate = CDate(Format(dtDate, "yyyy-mm-dd"))
End Function


