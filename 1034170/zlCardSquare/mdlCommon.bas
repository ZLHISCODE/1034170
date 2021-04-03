Attribute VB_Name = "mdlCommon"
Option Explicit
Public gcnOracle As ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例
Public gstrPrivs As String                   '当前用户具有的当前模块的功能
Public gstrSysName As String                '系统名称
Public glngModul As Long, glngSys As Long
Public gstrAviPath As String, gstrVersion As String
Public gstrMatchMethod As String
Public gstrProductName As String
Public gstrDBUser As String   '当前数据库用户
Public gstrUnitName As String '用户单位名称
Public gfrmMain As Object
Public gstrSQL As String
Public gblnTestCardNo As Boolean  '测试
Public gintDebug As Integer
Private Type gPrecision
      ty_小数 As Integer
      ty_Fmt_Vb As String
      ty_Fmt_Ora As String
End Type
Private Type FeePrecision   '费用相关精度
        ty_单价 As gPrecision
        ty_金额 As gPrecision
End Type
Public glngOld As Long
Private Type TY_WindowsRect
    MaxW As Long
    MaxH As Long
    MinW  As Long
    MinH As Long
End Type
Public gWinRect As TY_WindowsRect

Private Type SystemParameter
    int简码方式 As Integer
    bln个性化风格 As Boolean               '使用个性化风格
    bln全数字按编码查 As Boolean
    bln全字母按简码查 As Boolean
    bln存在站点 As Boolean      '是否存在站点管理
    ty_费用精度 As FeePrecision    '费用精度
End Type
Public gSystemPara As SystemParameter
Public Enum mAlignment
    mLeftAgnmt = 0
    mCenterAgnmt
    mRightAgnmt
End Enum

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type



Public Type Ty_UserInfor
    ID As Long
    部门ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
    部门名称 As String
    
End Type
Public UserInfo As Ty_UserInfor
Public Enum gRegType
    g注册信息 = 0
    g公共全局 = 1
    g公共模块 = 2
    g私有全局 = 3
    g私有模块 = 4
End Enum

Public Type Ty_Color
     lngGridColorSel As OLE_COLOR     '选择颜色
     lngGridColorLost As OLE_COLOR   '离开颜色
End Type
Public gSysColor As Ty_Color


Public glngHook As Long
Public gdtBegin As Date


'以下为卡对象
Public gblnRunLog As Boolean '是否记录使用日志
Public gblnErrLog As Boolean '是否记录运行错误
Public grsParas As ADODB.Recordset '系统参数表缓存
Public grsUserParas As ADODB.Recordset '系统参数表缓存
Public grsDeptParas As ADODB.Recordset    '系统参数部门缓存
Public gstrComputerName As String '计算机名称
Public glngInstanceCount As Long '当前实例个数
Public gcolPrivs As Collection '权限对象

Public Sub UnHookKBD()
    If glngHook <> 0 Then
    UnhookWindowsHookEx glngHook
    glngHook = 0
    End If
End Sub

Public Function EnableKBDHook()
    If glngHook <> 0 Then
        gdtBegin = Time
        Exit Function
    End If
    gdtBegin = Time
    glngHook = SetWindowsHookEx(WH_KEYBOARD, AddressOf MyKBHFunc, App.hInstance, App.ThreadID)
End Function

Public Function MyKBHFunc(ByVal iCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If (Time - gdtBegin) * 60 * 60 * 24 < 0.3 Then
        MyKBHFunc = 1 '表示要处理这个讯息If wParam = vbKeySnapshot Then '侦测 有没有按到PrintScreen键MyKBHFunc = 1 '在这个Hook便吃掉这个讯息End If
    Else
        MyKBHFunc = 0
    End If
    Call CallNextHookEx(glngHook, iCode, wParam, lParam) '传给下一个HookEnd Function
End Function


Public Function SetWindowResizeWndMessage(ByVal hWnd As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
'功能：自定义消息函数处理窗体尺寸调整限制
    If msg = WM_GETMINMAXINFO Then
        Dim MinMax As MINMAXINFO
        CopyMemory MinMax, ByVal lp, Len(MinMax)
        MinMax.ptMinTrackSize.X = gWinRect.MinW \ Screen.TwipsPerPixelX
        MinMax.ptMinTrackSize.Y = gWinRect.MinH \ Screen.TwipsPerPixelY
        MinMax.ptMaxTrackSize.X = gWinRect.MaxW \ Screen.TwipsPerPixelX
        MinMax.ptMaxTrackSize.Y = gWinRect.MaxH \ Screen.TwipsPerPixelY
        CopyMemory ByVal lp, MinMax, Len(MinMax)
        SetWindowResizeWndMessage = 1
        Exit Function
    End If
    SetWindowResizeWndMessage = CallWindowProc(glngOld, hWnd, msg, wp, lp)
End Function


'去掉TextBox的默认右键菜单
Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' 如果消息不是WM_CONTEXTMENU，就调用默认的窗口函数处理
    If msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(glngTXTProc, hWnd, msg, wp, lp)
End Function

Public Function GetParentWindow(ByVal hwndFrm As Long) As Long
    On Error Resume Next
    '获取指定窗体的父窗体
    GetParentWindow = GetWindowLong(hwndFrm, GWL_HWNDPARENT)
End Function


Public Function GetText(ByVal hwndFrm As Long) As String
    Dim strCaption As String * 256
    On Error Resume Next
    '获取指定窗体的标题
    Call GetWindowText(hwndFrm, strCaption, 255)
    GetText = zlCommFun.TruncZero(strCaption)
End Function



Public Sub FormSetCaption(ByVal objForm As Object, ByVal blnCaption As Boolean, Optional ByVal blnBorder As Boolean = True)
'功能：显示或隐藏一个窗体的标题栏
'参数：blnBorder=隐藏标题栏的时候,是否也隐藏窗体边框
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



Public Sub zlSetWindowsBroldStyle(ByVal frmMain As Form)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:改变可调窗体不不可调窗体（即设置只有关闭按钮窗口,如果窗体本身只有关闭，只会自动加上最大化、最小化等按钮)
    '入参:frmMain.hwnd-窗体的句柄
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-12-10 14:58:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim pt_SavePoint As POINTAPI, pt_MovePoint As POINTAPI
    Err = 0: On Error GoTo Errhand:
    With pt_MovePoint
      .X = (-1): .Y = 10
    End With
    '设置窗体的broldStyle
    Call SetWindowLong(frmMain.hWnd, GWL_STYLE, GetWindowLong(frmMain.hWnd, GWL_STYLE) Xor _
                              (WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX))
    Call GetSystemMenu(frmMain.hWnd, 1&)
    '必需重画数据
    With frmMain
        .Move .Left, .Top, .Width - 15, .Height - 15
        .Move .Left, .Top, .Width + 15, .Height + 15
    End With
    Call GetCursorPos(pt_SavePoint)
    Call ClientToScreen(frmMain.hWnd, pt_MovePoint)
    Call SetCursorPos(pt_MovePoint.X, pt_MovePoint.Y)
    Call SetCursorPos(pt_SavePoint.X, pt_SavePoint.Y)
Errhand:
End Sub

Public Sub zlInitColorSet()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始系统颜色
    '编制:刘兴洪
    '日期:2009-11-27 17:12:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '        Public Const G_Row_COLORSEL = &H8000000D
    '        Public Const G_Row_COLORLost = &HE0E0E0
    With gSysColor
        .lngGridColorLost = &HE0E0E0   '离开颜色
        .lngGridColorSel = &HFFEBD7       '选择颜色
    End With
End Sub
Public Function zl_GetUserInfo(Optional cnOracle As ADODB.Connection) As Boolean
'功能：获取登陆用户信息
    Dim rsTmp As ADODB.Recordset
    Dim objDatabase As clsDataBase
    
    
    If Not cnOracle Is Nothing Then
        Set objDatabase = New clsDataBase
        Call objDatabase.InitCommon(cnOracle)
        Set rsTmp = objDatabase.GetUserInfo
        Set objDatabase = Nothing
    Else
        Set rsTmp = zlDatabase.GetUserInfo
    End If
    
    UserInfo.用户名 = gstrDBUser
    UserInfo.姓名 = gstrDBUser
    If Not rsTmp.EOF Then
        UserInfo.ID = rsTmp!ID
        UserInfo.编号 = rsTmp!编号
        UserInfo.部门ID = IIf(IsNull(rsTmp!部门ID), 0, rsTmp!部门ID)
        UserInfo.部门名称 = "" & rsTmp!部门名
        UserInfo.简码 = "" & rsTmp!简码
        UserInfo.姓名 = "" & rsTmp!姓名
        zl_GetUserInfo = True
    End If
    Exit Function
Errhand:
    If Not objDatabase Is Nothing Then
        If objDatabase.ErrCenter() = 1 Then Resume
        Call objDatabase.SaveErrLog
    Else
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End If
End Function
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
Public Sub CalcPosition(ByRef X As Single, ByRef Y As Single, ByVal objBill As Object, Optional blnNoBill As Boolean = False)
    '----------------------------------------------------------------------
    '功能： 计算X,Y的实际坐标，并考虑屏幕超界的问题
    '参数： X---返回横坐标参数
    '       Y---返回纵坐标参数
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.hWnd, objPoint)
    If blnNoBill Then
        X = objPoint.X * 15 'objBill.Left +
        Y = objPoint.Y * 15 + objBill.Height '+ objBill.Top
    Else
        X = objPoint.X * 15 + objBill.CellLeft
        Y = objPoint.Y * 15 + objBill.CellTop + objBill.CellHeight
    End If
End Sub
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
Public Function GetMatchingSting(ByVal strString As String, Optional blnUpper As Boolean = True) As String
    '--------------------------------------------------------------------------------------------------------------------------------------
    '功能:加入匹配串%
    '参数:strString 需匹配的字串
    '返回:返回加匹配串%dd%,并且是大写
    '--------------------------------------------------------------------------------------------------------------------------------------
    Dim strLeft As String
    Dim strRight As String
    
    If gstrMatchMethod = "0" Then
        strLeft = "%"
        strRight = "%"
    Else
        strLeft = ""
        strRight = "%"
    End If
    If blnUpper = False Then
        GetMatchingSting = strLeft & strString & strRight
    Else
        GetMatchingSting = strLeft & UCase(strString) & strRight
    End If
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
Public Function CheckIsDate(ByVal strKey As String, ByVal strTittle As String) As String
    '------------------------------------------------------------------------------
    '功能:检查是否合法的日期型,可以为:20070101或2007-01-01
    '参数:strKey-需要检查的关建字
    '返回:合法的日期,返回标准格式(yyyy-mm-dd),否则返回""
    '编制:刘兴宏
    '日期:2008/01/24
    '------------------------------------------------------------------------------
    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
        strKey = TranNumToDate(strKey)
        If strKey = "" Then
            ShowMsgbox strTittle & "必须为日期型,请检查！"
            Exit Function
        End If
    End If
    If Not IsDate(strKey) Then
        ShowMsgbox strTittle & "必须为日期型如(2000-10-10) 或（20001010）,请检查！"
        Exit Function
    End If
    CheckIsDate = strKey
End Function


Public Sub SetTxtGotFocus(ByVal objTxt As Object, Optional blnOpenIme As Boolean = False)
    '--------------------------------------------------------------------------------------------------------
    '功能：对文本框的的文本选中或进入进打开输入法
    '参数:blnOpenIme-是否打开输入法
    '返回:
    '--------------------------------------------------------------------------------------------------------
    If TypeName(objTxt) = "TextBox" Then
        objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text) ' Len(objTxt.Text)
    ElseIf TypeName(objTxt) = "MaskEdBox" Then
        If Not IsDate(objTxt.Text) Then
            objTxt.SelStart = 0: objTxt.SelLength = Len(objTxt.Text)
        Else
            objTxt.SelStart = 0: objTxt.SelLength = 10
        End If
    End If
    
    If blnOpenIme Then
        zlCommFun.OpenIme (True)
    Else
        zlCommFun.OpenIme (False)
    End If
End Sub

Public Function Nvl(rsObj As Field, Optional ByVal varValue As Variant = "") As Variant
    '-----------------------------------------------------------------------------------
    '功能:取某字段的值
    '参数:rsObj          被检查的字段
    '     varValue       当rsObj为NULL值时的取新值
    '返回:如果不为空值,返回原来的值,如果为空值,则返回指定的varValue值
    '-----------------------------------------------------------------------------------
    If IsNull(rsObj) Then
        Nvl = varValue
    Else
        Nvl = rsObj
    End If
End Function
Public Function TranNumToDate(ByVal strNum As Long) As String
    Dim strYear As String
    Dim strMonth As String
    Dim strDay As String
    Dim strDate As String
    Err = 0
    On Error GoTo Errhand:
    TranNumToDate = ""
    strYear = Mid(strNum, 1, 4)
    strMonth = Mid(strNum, 5, 2)
    strDay = Mid(strNum, 7, 2)
        
    If strYear < 1000 Or strYear > 5000 Then Exit Function
    
    If strMonth > 12 Or strMonth < 1 Then Exit Function
    strDate = strYear & "-" & strMonth & "-" & strDay
        
    If Not IsDate(strDate) Then Exit Function
    
    strDate = Format(strDate, "yyyy-mm-dd")
    TranNumToDate = strDate
    Exit Function
Errhand:
    TranNumToDate = ""
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


Public Sub RaisEffect(picBox As PictureBox, Optional IntStyle As Integer, Optional strName As String = "", Optional TxtAlignment As mAlignment = 1, Optional blnFontBold As Boolean = False)
    '功能：将PictureBox模拟成3D平面按钮
    '参数：intStyle:0=平面,-1=凹下,1=凸起,2-深凸起
    Dim picRect As RECT
    Dim lngTmp As Long
    With picBox
        .Cls
        lngTmp = .ScaleMode
        .ScaleMode = 3
        .BorderStyle = 0
        If IntStyle <> 0 Then
            picRect.Left = .ScaleLeft
            picRect.Top = .ScaleTop
            picRect.Right = .ScaleWidth
            picRect.Bottom = .ScaleHeight
            Select Case IntStyle
            Case 1
                DrawEdge .hDC, picRect, BDR_RAISEDINNER Or BF_SOFT, BF_RECT
            Case 2
                DrawEdge .hDC, picRect, EDGE_RAISED, BF_RECT
            Case -1
                DrawEdge .hDC, picRect, BDR_SUNKENOUTER Or BF_SOFT, BF_RECT
            End Select
        End If
        .ScaleMode = lngTmp
        If strName <> "" Then
            .CurrentY = (.ScaleHeight - .TextHeight(strName)) / 2
            If TxtAlignment = mCenterAgnmt Then
                .CurrentX = (.ScaleWidth - .TextWidth(strName)) / 2
            ElseIf TxtAlignment = mLeftAgnmt Then
                .CurrentX = .ScaleLeft
            Else
                .CurrentX = (.ScaleWidth - .TextWidth(strName)) - 10
            End If
            .FontBold = blnFontBold
            picBox.Print strName
        End If
    End With
End Sub

Public Function zl_GetFieldLens(ByVal strTableName As String, ByVal strFields As String) As Collection
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取字段的实际长度
    '入参:strTableName-表名称
    '     strFields-字段数(字段名要唯一，否则报错),如:编码,名称,简码
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-11-17 16:39:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New Recordset, cllFields As New Collection
    Dim varFields As Variant, i As Long
    
    On Error GoTo errHandle
    
    gstrSQL = "Select " & strFields & " From " & strTableName & " where rownum<1 "
    zlDatabase.OpenRecordset rsTemp, gstrSQL, "取字段长度"
    
    varFields = Split(strFields, ",")
    With rsTemp
        For i = 0 To UBound(varFields)
            Select Case .Fields(varFields(i)).type
            Case 222
            Case Else
                cllFields.Add .Fields(varFields(i)).DefinedSize, varFields(i)
            End Select
        Next
    End With
    Set zl_GetFieldLens = cllFields
    rsTemp.Close
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Sub Init站点信息()
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始化站点的相关信息
    '编制:刘兴洪
    '日期:2009-03-02 17:23:24
    '-----------------------------------------------------------------------------------------------------------
    gbln存在站点控制 = gstrNodeNo <> "-"
 End Sub
Public Sub zl_加载站点信息(ByVal objcbo As ComboBox)
    '-----------------------------------------------------------------------------------------------------------
    '功能:加载站点信息值
    '编制:刘兴洪
    '日期:2009-03-03 12:09:01
    '-----------------------------------------------------------------------------------------------------------
    With objcbo
        .Clear
        .AddItem ""
        .AddItem "0"
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
        .AddItem "6"
        .AddItem "7"
        .AddItem "8"
        .AddItem "9"
        .ListIndex = 0
    End With
End Sub
 
Public Function zl_获取站点限制(Optional ByVal blnAnd As Boolean = True, _
    Optional ByVal str别名 As String = "") As String
    '功能:获取站点条件限制:2008-09-02 14:30:17
    Dim strWhere As String
    Dim strAlia As String
    strAlia = IIf(str别名 = "", "", str别名 & ".") & "站点"
    strWhere = IIf(blnAnd, " And ", "") & " (" & strAlia & "='" & gstrNodeNo & "' Or " & strAlia & " is Null)"
    zl_获取站点限制 = strWhere
End Function


Public Function IsCtrlSetFocus(ByVal objCtl As Object) As Boolean
    '------------------------------------------------------------------------------
    '功能:判断控件是否可
    '返回:初如成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2008/01/24
    '------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Err = 0: On Error GoTo Errhand:
    IsCtrlSetFocus = objCtl.Enabled And objCtl.Visible
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
Public Sub zlCtlSetFocus(ByVal objCtl As Object, Optional blnDoEvnts As Boolean = False)
    '功能:将集点移动控件中:2008-07-08 16:48:35
    Err = 0: On Error Resume Next
    If blnDoEvnts Then DoEvents
    If IsCtrlSetFocus(objCtl) = True Then: objCtl.SetFocus
End Sub


'*********************************************************************************************************************
Public Sub AddArray(ByRef cllData As Collection, ByVal strSQL As String)
    '---------------------------------------------------------------------------------------------
    '功能:向指定的集合中插入数据
    '参数:cllData-指定的SQL集
    '     strSql-指定的SQL语句
    '编制:刘兴宏
    '日期:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    i = cllData.count + 1
    cllData.Add strSQL, "K" & i
End Sub
Public Sub ExecuteProcedureArrAy(ByVal cllProcs As Variant, ByVal strCaption As String, Optional blnNoCommit As Boolean = False)
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:执行相关的Oracle过程集
    '参数:cllProcs-oracle过程集
    '     strCaption -执行过程的父窗口标题
    '     blnNOCommit-执行完过程后,不提交数据
    '编制:刘兴宏
    '日期:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String
    gcnOracle.BeginTrans
    For i = 1 To cllProcs.count
        strSQL = cllProcs(i)
        Call zlDatabase.ExecuteProcedure(strSQL, strCaption)
    Next
    If blnNoCommit = False Then
        gcnOracle.CommitTrans
    End If
End Sub
Public Function zlComboxLoadFromRecodeset(ByVal strFromCaption As String, ByVal rsSource As ADODB.Recordset, cboControls As Variant, Optional ByVal blnID As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:本函数的功能是从本地记录何时中装到下拉框中
    '入参:cboControls-控件数组
    '     rsSource:源记录(编码,名称,缺省标志)
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-12-09 14:54:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim intCount As Long
    Dim cboArrays As Variant
    On Error GoTo errHandle
    
    Set rsTemp = rsSource
    '下拉框数组
    If IsArray(cboControls) Then
        cboArrays = cboControls
    Else
        '强行组成一个数组
        cboArrays = Array(cboControls)
    End If
    For intCount = LBound(cboArrays) To UBound(cboArrays)
        cboArrays(intCount).Clear
        Do Until rsTemp.EOF
            If IsNull(rsTemp("编码")) Then
                cboArrays(intCount).AddItem rsTemp.AbsolutePosition & "." & rsTemp("名称")
            Else
                cboArrays(intCount).AddItem rsTemp("编码") & "." & rsTemp("名称")
            End If
            If blnID = True Then cboArrays(intCount).ItemData(cboArrays(intCount).NewIndex) = rsTemp("ID")
            If rsTemp("缺省标志") = 1 Then
                cboArrays(intCount).ListIndex = cboArrays(intCount).NewIndex
            End If
            rsTemp.MoveNext
        Loop
        If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
        If blnID = True And cboArrays(intCount).ListIndex < 0 Then cboArrays(intCount).ListIndex = 0
    Next
    zlComboxLoadFromRecodeset = True
    Exit Function
errHandle:
    zlComboxLoadFromRecodeset = False
End Function

Public Function zlComboxLoadFromArray(ByVal varArray As Variant, cboControls As Variant, Optional blnSaveItemData As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:本函数的功能是数组中读出列表值装到下拉框中
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-12-09 14:53:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cboArrays As Variant
    Dim intArray As Long
    Dim intCount As Long
    
    On Error GoTo errHandle
    
    If IsArray(cboControls) Then
        cboArrays = cboControls
    Else
        '强行组成一个数组
        cboArrays = Array(cboControls)
    End If
    
    For intCount = LBound(cboArrays) To UBound(cboArrays)
        cboArrays(intCount).Clear
        For intArray = LBound(varArray) To UBound(varArray)
            cboArrays(intCount).AddItem varArray(intArray)
        Next
        cboArrays(intCount).ListIndex = 0
    Next
    
    zlComboxLoadFromArray = True
    Exit Function
errHandle:
    zlComboxLoadFromArray = False
End Function

Public Function zlDblIsValid(ByVal strInput As String, ByVal intMax As Integer, Optional blnNegative As Boolean = True, Optional blnZero As Boolean = True, _
        Optional ByVal hWnd As Long = 0, Optional str项目 As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:检查字符串是否合法的金额
    '入参:strInput        输入的字符串
    '     intMax          整数的位数
    '     blnNegative     是否进行负数检查
    '     blnZero         是否进行零的检查
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-10-20 15:16:08
    '-----------------------------------------------------------------------------------------------------------
    Dim dblValue As Double
    If blnZero = True Then
        If strInput = "" Then
            ShowMsgbox str项目 & "未输入，请检查!"
            If hWnd <> 0 Then SetFocusHwnd hWnd
            Exit Function
        End If
    End If
    If strInput = "" Then zlDblIsValid = True: Exit Function
    If IsNumeric(strInput) = False Then
        MsgBox str项目 & "不是有效的数字格式。", vbInformation, gstrSysName
        If hWnd <> 0 Then SetFocusHwnd hWnd              '设置焦点
        Exit Function
    End If
    
    dblValue = Val(strInput)
    If dblValue >= 10 ^ intMax - 1 Then
        MsgBox str项目 & "数值过大，不能超过" & 10 ^ intMax - 1 & "。", vbInformation, gstrSysName
        If hWnd <> 0 Then SetFocusHwnd hWnd              '设置焦点
        Exit Function
    End If
    If blnNegative = True And dblValue < 0 Then
        MsgBox str项目 & "不能输入负数。", vbInformation, gstrSysName
        If hWnd <> 0 Then SetFocusHwnd hWnd              '设置焦点
        Exit Function
    End If
    
    If Abs(dblValue) >= 10 ^ intMax And dblValue < 0 Then
        MsgBox str项目 & "数值过小，不能小于-" & 10 ^ intMax - 1 & "位。", vbInformation, gstrSysName
        If hWnd <> 0 Then SetFocusHwnd hWnd              '设置焦点
        Exit Function
    End If
    
    
    If blnZero = True And dblValue = 0 Then
        MsgBox str项目 & "不能输入零。", vbInformation, gstrSysName
        If hWnd <> 0 Then SetFocusHwnd hWnd              '设置焦点
        Exit Function
    End If
    zlDblIsValid = True
End Function
Public Function zl_FromComboxGetData(cboControl As ComboBox, Optional ByVal blnID As Boolean = False, Optional strSplit As String = ".") As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:从Combox中获取数据
    '入参:blnID-是否读取ComboxData数据
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2009-12-11 15:22:18
    '---------------------------------------------------------------------------------------------------------------------------------------------

    If cboControl.ListIndex < 0 Then zl_FromComboxGetData = "NULL"
    If blnID = False Then
        If cboControl.Text = "" Or cboControl.Enabled = False Then
            zl_FromComboxGetData = "NULL"
        Else
            zl_FromComboxGetData = "'" & Mid(cboControl.Text, InStr(cboControl.Text, strSplit) + 1) & "'"
        End If
    Else
        zl_FromComboxGetData = cboControl.ItemData(cboControl.ListIndex)
    End If
End Function
 Public Function IsDesinMode() As Boolean
      '刘兴洪 确定当前模式为设计模式
     Err = 0: On Error Resume Next
     Debug.Print 1 / 0
     If Err <> 0 Then
        IsDesinMode = True
     Else
        IsDesinMode = False
     End If
     Err.Clear: Err = 0
 End Function
  

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

Public Function zlGetReDawImge(ByVal frmMain As Form, ByVal lngColor As Long, _
    ByVal strCaption As String, sngWidth As Single, sngHeight As Single, _
    Optional sngFontSize As Single = 9, _
    Optional blnFontBold As Boolean = True) As StdPicture
    Dim objPicture As PictureBox
    Set objPicture = frmMain.Controls.Add("VB.PictureBox", "objPictemp")
    With objPicture
        .Cls
        .AutoRedraw = True
        .FontSize = 9
        .Width = sngWidth: .Height = sngHeight
        objPicture.Line (20, 20)-(sngWidth, sngHeight), lngColor, BF              '一个矩形(填充)
        .ForeColor = &H80000016
        .CurrentY = 20
        .FontBold = blnFontBold
        .FontSize = sngFontSize
        If strCaption <> "" Then
            .CurrentY = (.ScaleHeight - .TextHeight("刘")) \ 2
            .CurrentX = (.ScaleWidth - .TextWidth(strCaption)) \ 2
            objPicture.Print strCaption
        End If
    End With
    Set zlGetReDawImge = objPicture.Image
    frmMain.Controls.Remove ("objPictemp")
    Set objPicture = Nothing
End Function
Public Sub zlSetStatusPanelCololor(ByVal frmMain As Form, ByVal objStatus As Object, _
    ByVal intPancelIdex As Integer, strCaption As String, _
    ByVal lngColor As Long, Optional blnTextCenter As Boolean = True)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：设置单元格的颜色
    '入参：blnTextCenter-文本居中
    '出参：
    '返回：
    '编制：刘兴洪
    '日期：2010-03-23 15:22:18
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim sngWidth As Single, sngHeight As Single
    With objStatus
        sngWidth = frmMain.TextWidth(strCaption) + 60
        sngHeight = frmMain.TextHeight("刘") + 60
        .Panels(intPancelIdex).Width = sngWidth
        If blnTextCenter = False Then
            .Panels(intPancelIdex).Width = sngWidth + 300
            .Panels(intPancelIdex).Text = strCaption
            .Panels(intPancelIdex).Picture = zlGetReDawImge(frmMain, lngColor, "", 300, sngHeight, 7, True)
        Else
            .Panels(intPancelIdex).Picture = zlGetReDawImge(frmMain, lngColor, strCaption, sngWidth, sngHeight, 7, True)
        End If
    End With
End Sub
Public Sub zlClearAllObjectValue()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除所有对象的值
    '编制:刘兴洪
    '日期:2011-05-23 15:06:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    If Not grsStatic.rs收费类别 Is Nothing Then
        If grsStatic.rs收费类别.State = 1 Then grsStatic.rs收费类别.Close
    End If
    If Not grsStatic.rs消费卡接口 Is Nothing Then
        If grsStatic.rs消费卡接口.State = 1 Then grsStatic.rs消费卡接口.Close
    End If
    If Not grs医疗卡类别 Is Nothing Then
        If grs医疗卡类别.State = 1 Then grs医疗卡类别.Close
    End If
    Set grsStatic.rs收费类别 = Nothing
    Set grsStatic.rs消费卡接口 = Nothing
    Set gobjComLib = Nothing
    Set gobjCommFun = Nothing
    Set gobjControl = Nothing
    Set gobjDatabase = Nothing
    Set grs医疗卡类别 = Nothing
End Sub

Public Sub zlDebugTool(ByVal strInfo As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:跟踪调试信息
    '入参:strInfo-调试信息
    '编制:刘兴洪
    '日期:2011-05-27 11:36:33
    '说明:
    '     gintDebug:1-表示提未调试信息,2-将调式信息写入文本；其它情况不输出调试信息
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objFile As FileSystemObject, objText As TextStream, strFile As String
    If gintDebug = -1 Then gintDebug = Val(GetSetting("ZLSOFT", "卡结算部件", "调试", 0))
    '先判断是否存在该文件，不存在则创建（调试=0，直接退出；其他情况都输出调试信息）
    If gintDebug <= 0 Or gintDebug > 2 Then Exit Sub
    If gintDebug = 2 Then
        '写入文件中
        Set objFile = New FileSystemObject
        strFile = App.Path & "\Square" & Format(Now, "yyyy_MM_DD") & ".Log"
        If Not Dir(strFile) <> "" Then objFile.CreateTextFile strFile
        Set objText = objFile.OpenTextFile(strFile, ForAppending)
        objText.WriteLine strInfo: objText.Close
    End If
    MsgBox strInfo, vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
End Sub
Public Sub zlAddArray(ByRef cllData As Collection, ByVal strSQL As String)
    '---------------------------------------------------------------------------------------------
    '功能:向指定的集合中插入数据
    '参数:cllData-指定的SQL集
    '     strSql-指定的SQL语句
    '编制:刘兴宏
    '日期:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    i = cllData.count + 1
    cllData.Add strSQL, "K" & i
End Sub
Public Sub zlExecuteProcedureArrAy(ByVal cllProcs As Variant, ByVal strCaption As String, _
    Optional blnNoCommit As Boolean = False, _
    Optional blnNoBeginTrans As Boolean = False)
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:执行相关的Oracle过程集
    '参数:cllProcs-oracle过程集
    '     strCaption -执行过程的父窗口标题
    '     blnNOCommit-执行完过程后,不提交数据
    '     blnNoBeginTrans:没有事务开始
    '编制:刘兴宏
    '日期:2008/01/09
    '---------------------------------------------------------------------------------------------
    Dim i As Long
    Dim strSQL As String
    If blnNoBeginTrans = False Then gcnOracle.BeginTrans
    For i = 1 To cllProcs.count
        strSQL = cllProcs(i)
        Call zlDatabase.ExecuteProcedure(strSQL, strCaption)
    Next
    If blnNoCommit = False Then gcnOracle.CommitTrans
End Sub

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
        strNumber = IIf(blnShowZero, 0, "")
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

Public Function zlAuditingWarn(ByVal strPrivs As String, _
    ByVal strNOs As String, ByVal lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:审核划价单时，对费用进行报警
    '入参:str序号=指定单据中要审核的行号,为空表示所有行
    '返回:
    '编制:刘兴洪
    '日期:2011-06-23 10:29:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsWarn As ADODB.Recordset, rsTmp As ADODB.Recordset
    Dim strSQL As String, j As Long, str类别s As String
    Dim cur当日额 As Currency, cur金额 As Currency, cur余额 As Currency
    Dim strWarn As String, intWarn As Integer
    Dim bln报警包含划价费用  As Boolean
    '记帐报警包含划价费用
    bln报警包含划价费用 = zlDatabase.GetPara(98, glngSys) = "1"
    
    strSQL = "" & _
    " Select /*+ rule */ A.门诊标志, A.姓名, A.病人id , E.预交余额 - E.费用余额 As 余额, B.担保额, C.编码 As 付款码," & vbNewLine & _
    "        A.收费类别, D.名称 As 类别名称, Sum(A.实收金额) As 金额, Zl_Patiwarnscheme(A.病人id) As 适用病人" & vbNewLine & _
    " From 门诊费用记录 A, 病人信息 B, Table(f_Str2list([1])) J," & _
    "           医疗付款方式 C, 收费项目类别 D," & _
    "           (   Select 病人ID,Sum(Nvl(预交余额,0)) as 预交余额,Sum(nvl(费用余额,0))  费用余额" & _
    "               From  病人余额 " & vbNewLine & _
    "               Where   病人ID=[2]  and 性质=1 And nvl(类型,2)=1 Group by 病人ID)  E" & vbNewLine & _
    " Where A.记录性质 = 2 And A.病人ID+0=[2] And A.记录状态 = 0 " & _
    "           And A.NO = J.Column_value " & vbNewLine & _
    "           And A.收费类别 = D.编码 And A.病人id = E.病人id(+) " & vbNewLine & _
    "           And A.病人id = B.病人id And B.医疗付款方式 = C.名称(+)" & vbNewLine & _
    " Group By Nvl(A.价格父号, A.序号), A.门诊标志, A.姓名, A.病人id,  B.担保额, E.预交余额, E.费用余额, C.编码," & vbNewLine & _
    "         A.收费类别, D.名称"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlOutExse", strNOs, lng病人ID)
    If rsTmp.RecordCount > 0 Then
        Do While Not rsTmp.EOF
            If InStr(str类别s, rsTmp!收费类别 & rsTmp!类别名称) = 0 Then
                str类别s = str类别s & "," & rsTmp!收费类别 & rsTmp!类别名称
            End If
            cur金额 = cur金额 + rsTmp!金额
            rsTmp.MoveNext
        Loop
        rsTmp.MoveFirst
        str类别s = Mid(str类别s, 2)
        If cur金额 > 0 Then
            Set rsWarn = zlGetUnitWarn(rsTmp!适用病人, "0")
            cur当日额 = GetPatiDayMoney(rsTmp!病人ID)
            cur余额 = Nvl(rsTmp!余额, 0)
            If bln报警包含划价费用 Then cur余额 = cur余额 - GetPriceMoneyTotal(rsTmp!病人ID) + cur金额
            '分类报警
            For j = 0 To UBound(Split(str类别s, ","))
                intWarn = zlBillingWarn(strPrivs, rsTmp!姓名, rsTmp!适用病人, rsWarn, _
                    cur余额, cur当日额, cur金额, Nvl(rsTmp!担保额, 0), _
                    Left(Split(str类别s, ",")(j), 1), Mid(Split(str类别s, ",")(j), 2), strWarn)
                If intWarn = 2 Or intWarn = 3 Then Exit Function
            Next
        End If
    End If
    zlAuditingWarn = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function zlGetUnitWarn(Optional ByVal str适用病人 As String, Optional ByVal str病区ID As String) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:返回病区记帐报警记录集
    '入参:str适用病人-适用的病人
    '        str病区IＤ－病区ID集
    '出参:
    '返回:病区报警集
    '编制:刘兴洪
    '日期:2011-06-24 14:59:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errH
    strSQL = "Select Nvl(病区ID,0) 病区ID,适用病人,Nvl(报警方法,1) as 报警方法," & _
            " 报警值,报警标志1,报警标志2,报警标志3" & _
            " From 记帐报警线 Where 1=1" & _
            IIf(str适用病人 = "", "", " And 适用病人 = [1]") & _
            IIf(str病区ID = "", "", " And Nvl(病区ID,0) = [2]")
    Set zlGetUnitWarn = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, str适用病人, str病区ID)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function zlBillingWarn(strPrivs As String, str姓名 As String, str适用病人 As String, _
    rsWarn As ADODB.Recordset, cur余额 As Currency, cur当日额 As Currency, _
    cur单据金额 As Currency, cur担保 As Currency, str类别 As String, _
    ByVal str类别名 As String, ByRef str已报类别 As String, Optional bln多病人 As Boolean, Optional strMoneyFMT As String = "") As Integer
'功能:对病人记帐进行报警提示
'参数:
'     str姓名=病人姓名,用于提示
'     str适用病人=根据病人身份返回的记帐报警适用方案
'     rsWarn=当前病区记帐报警设置记录
'     cur余额=病人余额,用于累计报警
'     cur当日额=病人当日发生的费用额,用于每日报警
'     cur单据金额=病人单据中输入的费用
'     cur担保=病人担保费用额,用于累计报警
'     str类别=当前要检查的类别,用于分类报警
'     str类别名=类别名称,用于提示
'     strMoneyFMT-格式精度
'返回:0;没有报警,继续
'     1:报警提示后用户选择继续
'     2:报警提示后用户选择中断
'     3:报警提示必须中断
'     4:强制记帐报警,继续
'     str报警类别="CDE":具体在本次报警的一组类别,"-"为所有类别。该返回用于处理重复报警
    Dim i As Integer, byt标志 As Byte
    Dim bln已报警 As Boolean
    Dim byt方式 As Byte, byt已报方式 As Byte
    Dim arrTmp As Variant
    
    On Error GoTo errH
    If strMoneyFMT = "" Then
        strMoneyFMT = "0." & String(Val(zlDatabase.GetPara(9, glngSys, , 2)), "0")
    End If
    '报警参数检查
    rsWarn.Filter = "病区ID=0 And 适用病人='" & str适用病人 & "'"
    If rsWarn.EOF Then Exit Function
    If IsNull(rsWarn!报警值) Then Exit Function
    
    '对应类别定位有效报警设置
    If Not IsNull(rsWarn!报警标志1) Then
        If rsWarn!报警标志1 = "-" Or InStr(rsWarn!报警标志1, str类别) > 0 Then byt标志 = 1
        If rsWarn!报警标志1 = "-" Then str类别名 = "" '所有类别时,不必提示具体的类别
    End If
    If byt标志 = 0 And Not IsNull(rsWarn!报警标志2) Then
        If rsWarn!报警标志2 = "-" Or InStr(rsWarn!报警标志2, str类别) > 0 Then byt标志 = 2
        If rsWarn!报警标志2 = "-" Then str类别名 = "" '所有类别时,不必提示具体的类别
    End If
    If byt标志 = 0 And Not IsNull(rsWarn!报警标志3) Then
        If rsWarn!报警标志3 = "-" Or InStr(rsWarn!报警标志3, str类别) > 0 Then byt标志 = 3
        If rsWarn!报警标志3 = "-" Then str类别名 = "" '所有类别时,不必提示具体的类别
    End If
    If byt标志 = 0 Then Exit Function '无有效设置
    
    '报警标志2实际上是两种判断①②,其它只有一种判断①
    '这种处理的前提是一种类别只能属于一种报警方式(报警参数设置时)
    If bln多病人 Then
        '示例：",周:-,张:DEF,李:567,张567"
        '报警标志2示例：",周:-①,张:DEF①,李:567①,张567②"
        bln已报警 = str已报类别 & "," Like "*," & str姓名 & ":-*,*" _
            Or str已报类别 & "," Like "*," & str姓名 & ":*" & str类别 & "*,*"
    Else
        '示例："-" 或 ",ABC,567,DEF"
        '报警标志2示例："-①" 或 ",ABC②,567①,DEF①"
        bln已报警 = InStr(str已报类别, str类别) > 0 Or str已报类别 Like "-*"
    End If
    
    If bln已报警 Then
        If byt标志 = 2 Then
            If bln多病人 Then
                arrTmp = Split(str已报类别, ",")
                For i = 0 To UBound(arrTmp)
                    If "," & arrTmp(i) & "," Like "*," & str姓名 & ":-*,*" _
                        Or "," & arrTmp(i) & "," Like "*," & str姓名 & ":*" & str类别 & "*,*" Then
                        byt已报方式 = IIf(Right(arrTmp(i), 1) = "②", 2, 1)
                        'Exit For  '说明见住院模块
                    End If
                Next
            Else
                If str已报类别 Like "-*" Then
                    byt已报方式 = IIf(Right(str已报类别, 1) = "②", 2, 1)
                Else
                    arrTmp = Split(str已报类别, ",")
                    For i = 0 To UBound(arrTmp)
                        If InStr(arrTmp(i), str类别) > 0 Then
                            byt已报方式 = IIf(Right(arrTmp(i), 1) = "②", 2, 1)
                            'Exit For '说明见住院模块
                        End If
                    Next
                End If
            End If
        Else
            Exit Function
        End If
    End If
    
    If str类别名 <> "" Then str类别名 = """" & str类别名 & """费用"
        
    '---------------------------------------------------------------------
    If rsWarn!报警方法 = 1 Then  '累计费用报警(低于)
        Select Case byt标志
            Case 1 '低于报警值(包括预交款耗尽)提示询问记帐
                If cur余额 + cur担保 - cur单据金额 < rsWarn!报警值 Then
                    If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                        '先只有两种:1.强制记帐,无权限时,禁止记帐
                        Call MsgBox(str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & ",低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",禁止记帐", vbInformation + vbOKOnly, gstrSysName)
                        zlBillingWarn = 3
                    Else
                        MsgBox "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & " 低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", vbInformation, gstrSysName
                        zlBillingWarn = 4
                    End If
                End If
            Case 2 '低于报警值提示询问记帐,预交款耗尽时禁止记帐
                If Not bln已报警 Then
                    If cur余额 + cur担保 - cur单据金额 < 0 Then
                        byt方式 = 2
                        If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                            MsgBox str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & ")已经耗尽," & str类别名 & "禁止记帐。", vbInformation, gstrSysName
                            zlBillingWarn = 3
                        Else
                            MsgBox str类别名 & "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & ")已经耗尽。", vbInformation, gstrSysName
                            zlBillingWarn = 4
                        End If
                    ElseIf cur余额 + cur担保 - cur单据金额 < rsWarn!报警值 Then
                        byt方式 = 1
                        If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                            If MsgBox(str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & ",低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",继续记帐吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                                zlBillingWarn = 2
                            Else
                                zlBillingWarn = 1
                            End If
                        Else
                            MsgBox "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & ",低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", vbInformation, gstrSysName
                            zlBillingWarn = 4
                        End If
                    End If
                Else
                    '上次已报警并选择继续或强制继续
                    If byt已报方式 = 1 Then
                        '上次低于报警值并选择继续或强制继续,不再处理低于的情况,但还需要判断预交款是否耗尽
                        If cur余额 + cur担保 - cur单据金额 < 0 Then
                            byt方式 = 2
                            If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                                MsgBox str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & ")已经耗尽," & str类别名 & "禁止记帐。", vbInformation, gstrSysName
                                zlBillingWarn = 3
                            Else
                                MsgBox str类别名 & "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & ")已经耗尽。", vbInformation, gstrSysName
                                zlBillingWarn = 4
                            End If
                        End If
                    ElseIf byt已报方式 = 2 Then
                        '上次预交款已经耗尽并强制继续,不再处理
                        Exit Function
                    End If
                End If
            Case 3 '低于报警值禁止记帐
                If cur余额 + cur担保 - cur单据金额 < rsWarn!报警值 Then
                    If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                        MsgBox str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & ",低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",禁止记帐。", vbInformation, gstrSysName
                        zlBillingWarn = 3
                    Else
                        MsgBox "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当前剩余款(含担保额:" & Format(cur担保, "0.00") & "):" & Format(cur余额 + cur担保 - cur单据金额, "0.00") & ",低于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", vbInformation, gstrSysName
                        zlBillingWarn = 4
                    End If
                End If
        End Select
    ElseIf rsWarn!报警方法 = 2 Then  '每日费用报警(高于)
        Select Case byt标志
            Case 1 '高于报警值提示询问记帐
                If cur当日额 + cur单据金额 > rsWarn!报警值 Then
                    If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                        Call MsgBox(str姓名 & " 当日费用:" & Format(cur当日额 + cur单据金额, strMoneyFMT) & ",高于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",禁止记帐.", vbOKOnly + vbInformation, gstrSysName)
                        zlBillingWarn = 3
                    Else
                        MsgBox "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当日费用:" & Format(cur当日额 + cur单据金额, strMoneyFMT) & ",高于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", vbInformation, gstrSysName
                        zlBillingWarn = 4
                    End If
                End If
            Case 3 '高于报警值禁止记帐
                If cur当日额 + cur单据金额 > rsWarn!报警值 Then
                    If InStr(";" & strPrivs & ";", ";欠费强制记帐;") = 0 Then
                        MsgBox str姓名 & " 当日费用:" & Format(cur当日额 + cur单据金额, strMoneyFMT) & ",高于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & ",禁止记帐。", vbInformation, gstrSysName
                        zlBillingWarn = 3
                    Else
                        MsgBox "强制记帐提醒:" & vbCrLf & vbCrLf & str姓名 & " 当日费用:" & Format(cur当日额 + cur单据金额, strMoneyFMT) & ",高于" & str类别名 & "报警值:" & Format(rsWarn!报警值, "0.00") & "。", vbInformation, gstrSysName
                        zlBillingWarn = 4
                    End If
                End If
        End Select
    End If
    
    '对于继续类的操作,返回已报警类别
    If zlBillingWarn = 1 Or zlBillingWarn = 4 Then
        If byt标志 = 1 Then
            If rsWarn!报警标志1 = "-" Then
                str已报类别 = IIf(bln多病人, str已报类别 & "," & str姓名 & ":", "") & "-"
            Else
                str已报类别 = str已报类别 & IIf(bln多病人, "," & str姓名 & ":", ",") & rsWarn!报警标志1
            End If
        ElseIf byt标志 = 2 Then
            If rsWarn!报警标志2 = "-" Then
                str已报类别 = IIf(bln多病人, str已报类别 & "," & str姓名 & ":", "") & "-"
            Else
                str已报类别 = str已报类别 & IIf(bln多病人, "," & str姓名 & ":", ",") & rsWarn!报警标志2
            End If
            '附加标注以判断已报警的具体方式
            str已报类别 = str已报类别 & IIf(byt方式 = 2, "②", "①")
        ElseIf byt标志 = 3 Then
            If rsWarn!报警标志3 = "-" Then
                str已报类别 = IIf(bln多病人, str已报类别 & "," & str姓名 & ":", "") & "-"
            Else
                str已报类别 = str已报类别 & IIf(bln多病人, "," & str姓名 & ":", ",") & rsWarn!报警标志3
            End If
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function NeedName(strList As String, Optional strSplit As String = "-") As String
    NeedName = Mid(strList, InStr(strList, strSplit) + 1)
End Function

 
Public Function GetCboIndex(cbo As ComboBox, strFind As String, _
    Optional blnKeep As Boolean, _
    Optional blnLike As Boolean, Optional strSplit As String = "-") As Long
'功能：由字符串在ComboBox中查找索引
    Dim i As Long
    If strFind = "" Then GetCboIndex = -1: Exit Function
    '先精确查找
    For i = 0 To cbo.ListCount - 1
        If InStr(cbo.List(i), strSplit) > 0 Then
            If NeedName(cbo.List(i)) = strFind Then GetCboIndex = i: Exit Function
        Else
            If cbo.List(i) = strFind Then GetCboIndex = i: Exit Function
        End If
    Next
    '最后模糊查找
    If blnLike Then
        For i = 0 To cbo.ListCount - 1
            If InStr(cbo.List(i), strFind) > 0 Then GetCboIndex = i: Exit Function
        Next
    End If
    If Not blnKeep Then GetCboIndex = -1
End Function

Public Function GetWidth(cboHwnd As Long) As Long
'功能： 取得 Combo 下拉的宽度,单位为 pixels
    Dim lRetVal As Long
    lRetVal = SendMessage(cboHwnd, &H15F, 0, 0)
    If lRetVal <> -1 Then
        GetWidth = lRetVal
    Else
        GetWidth = 0
    End If
End Function


Public Function GetUnitID(bytFlag As Byte, lngID As Long) As Long
'功能：返回收费特定项目的执行科室
'参数：bytFlag=执行科室标志,lngID=收费细目ID
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    Select Case bytFlag
        Case 0 '无明确科室
            GetUnitID = UserInfo.部门ID '取操作员所在科室
        Case 4 '指定科室
            strSQL = "Select B.执行科室ID From 收费项目目录 A,收费执行科室 B Where B.收费细目ID=A.ID And A.ID=[1]"

            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", lngID)
            If rsTmp.RecordCount <> 0 Then
                GetUnitID = rsTmp!执行科室ID '默认取第一个(如有多个)
            Else
                GetUnitID = UserInfo.部门ID '如没有指定，则取操作员所在科室
            End If
        Case 1, 2, 3 '病人科室,操作员科室
            GetUnitID = UserInfo.部门ID '都取操作员科室
    End Select
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPatiDayMoney(lng病人ID As Long) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定病人当天发生的费用总额
    '返回:获取病人当天发生的费用总额
    '编制:刘兴洪
    '日期:2011-06-23 10:40:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errH
    strSQL = "Select zl_PatiDayCharge([1]) as 金额 From Dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng病人ID)
    If Not rsTmp.EOF Then
        GetPatiDayMoney = Val("" & rsTmp!金额)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetPriceMoneyTotal(lng病人ID As Long) As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定病人的划价单金额合计
    '返回:指定病人的划价费用合计
    '编制:刘兴洪
    '日期:2011-06-23 10:44:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errH
    strSQL = "" & _
    "   Select Nvl(Sum(实收金额),0) As 划价费用合计  " & _
    "   From 门诊费用记录 " & _
    "   Where 记录状态=0 And 记帐费用=1 And 病人ID=[1]  " & _
    "   Union ALL  " & _
    "   Select Nvl(Sum(实收金额),0) As 划价费用合计 " & _
    "   From 住院费用记录  " & _
    "   Where 记录状态=0 And 记帐费用=1 And 病人ID=[1] "
    strSQL = "Select Sum(nvl(划价费用合计,0)) as 划价费用合计 From ( " & strSQL & ")"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, lng病人ID)
    If Not rsTmp.EOF Then GetPriceMoneyTotal = Val(Nvl(rsTmp!划价费用合计))
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function



Public Function IncStr(ByVal strVal As String) As String
'功能：对一个字符串自动加1。
'说明：每一位进位时,如果是数字,则按十进制处理,否则按26进制处理
    Dim i As Integer, strTmp As String, bytUp As Byte, bytAdd As Byte
    
    For i = Len(strVal) To 1 Step -1
        If i = Len(strVal) Then
            bytAdd = 1
        Else
            bytAdd = 0
        End If
        If IsNumeric(Mid(strVal, i, 1)) Then
            If CByte(Mid(strVal, i, 1)) + bytAdd + bytUp < 10 Then
                strVal = Left(strVal, i - 1) & CByte(Mid(strVal, i, 1)) + bytAdd + bytUp & Mid(strVal, i + 1)
                bytUp = 0
            Else
                strVal = Left(strVal, i - 1) & "0" & Mid(strVal, i + 1)
                bytUp = 1
            End If
        Else
            If Asc(Mid(strVal, i, 1)) + bytAdd + bytUp <= Asc("Z") Then
                strVal = Left(strVal, i - 1) & Chr(Asc(Mid(strVal, i, 1)) + bytAdd + bytUp) & Mid(strVal, i + 1)
                bytUp = 0
            Else
                strVal = Left(strVal, i - 1) & "0" & Mid(strVal, i + 1)
                bytUp = 1
            End If
        End If
        If bytUp = 0 Then Exit For
    Next
    IncStr = strVal
End Function


Public Function PreFixNO(Optional Curdate As Date = #1/1/1900#) As String
'功能：返回大写的单据号年前缀
    If Curdate = #1/1/1900# Then
        PreFixNO = CStr(CInt(Format(zlDatabase.Currentdate, "YYYY")) - 1990)
    Else
        PreFixNO = CStr(CInt(Format(Curdate, "YYYY")) - 1990)
    End If
    PreFixNO = IIf(CInt(PreFixNO) < 10, PreFixNO, Chr(55 + CInt(PreFixNO)))
End Function

Public Function CheckValid() As Boolean
    Dim intAtom As Integer
    Dim blnValid As Boolean
    Dim strSource As String
    Dim strCurrent As String
    Dim strBuffer As String * 256
    CheckValid = False
    
    If gfrmMain Is Nothing Then CheckValid = True: Exit Function
    '获取注册表后，马上清零
    strCurrent = Format(Now, "yyyyMMddHHmm")
    intAtom = GetSetting("ZLSOFT", "公共全局", "公共", 0)
    Call SaveSetting("ZLSOFT", "公共全局", "公共", 0)
    blnValid = (intAtom <> 0)
    
    '如果存在，则对串进行解析
    If blnValid Then
        Call GlobalGetAtomName(intAtom, strBuffer, 255)
        strSource = Trim(Replace(strBuffer, Chr(0), ""))
        '如果为空，则表示非法
        If strSource <> "" Then
            If Left(strSource, 1) <> "#" Then
                strSource = TranPasswd(Mid(strSource, 1, 12))
                If strSource <> strCurrent Then '判断时间间隔是否大于1
                    If CStr(Mid(strSource, 11, 2) + 1) = CStr(Mid(strCurrent, 11, 2) + 0) Then
                        '如果相等，则通过
                    Else
                        '不等，表示存在进位，则分应该为零
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
        MsgBox "The component is lapse！", vbInformation, gstrSysName
        Exit Function
    End If
    CheckValid = True
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
Public Function SetCboDefault(cbo As ComboBox) As Integer
    Dim i As Integer
    For i = 0 To cbo.ListCount - 1
        If cbo.ItemData(i) = 1 Then
            cbo.ListIndex = i
            SetCboDefault = i: Exit Function
        End If
    Next
End Function
Public Function StrToNum(ByVal strNumber As String) As Double
    '功能:将字符串转换成数据
    Dim strTmp As String
    strTmp = Replace(strNumber, ",", "")
    StrToNum = Val(strTmp)
End Function


Public Function ExistFeeInsurePatient(lng病人ID As Long) As Boolean
'功能：判断医保病人是否存在未结费用
'返回：
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo errH
        
    strSQL = "Select Nvl(sum(B.费用余额),0) 费用余额 From 病人信息 A,病人余额 B Where A.病人ID=B.病人ID And Nvl(A.险类,0)<>0 And A.病人ID=[1]"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlPatient", lng病人ID)
    
    If Not rsTmp.EOF Then ExistFeeInsurePatient = (rsTmp!费用余额 <> 0)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetArea(frmParent As Object, txtInput As TextBox, Optional blnShowAll As Boolean) As ADODB.Recordset
'功能：获取地区列表或选择的地区
'参数：
    Dim strSQL As String, blnCancel As Boolean
    Dim vRect As RECT
    
    On Error GoTo errH
    vRect = GetControlRect(txtInput.hWnd)
    If Not blnShowAll Then
        strSQL = " Select 编码 as ID,编码,名称,简码 From 区域" & _
                 " Where Nvl(级数,0)<3 And (编码 Like [1] Or upper(简码) Like '" & gstrLike & "'||[1]||'%' Or 名称 Like '" & gstrLike & "'||[1]||'%')"
        Set GetArea = zlDatabase.ShowSQLSelect(frmParent, strSQL, 0, "区域", True, txtInput.Text, "", True, True, True, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, True, gstrLike & txtInput.Text & "%")
    Else
        strSQL = "Select 编码 as ID,编码,名称,简码 From 区域 Where Nvl(级数,0) < 3 "
        Set GetArea = zlDatabase.ShowSQLSelect(frmParent, strSQL, 0, "区域", True, txtInput.Text, "", True, True, True, vRect.Left, vRect.Top, txtInput.Height, blnCancel, True, True, gstrLike & txtInput.Text & "%")
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Sub CheckInputLen(txt As Object, KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: Exit Sub
    If KeyAscii < 32 And KeyAscii >= 0 Then Exit Sub
    If txt.MaxLength = 0 Then Exit Sub
    If zlCommFun.ActualLen(txt.Text & Chr(KeyAscii)) > txt.MaxLength Then KeyAscii = 0
End Sub
Public Function AnalyseComputer() As String
    Dim strComputer As String * 256
    Call GetComputerName(strComputer, 255)
    AnalyseComputer = strComputer
    AnalyseComputer = Replace(AnalyseComputer, Chr(0), "")
End Function

Public Function SetPatiColor(ByVal objPatiControl As Object, ByVal str病人类型 As String, _
    Optional ByVal lngDefaultColor As Long = vbBlack) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人类型,设置不同病人类型的显示颜色
    '入参:objPatiControl-病人控件(文本框,标签)
    '    str病人类型-病人类型
    '    lngDefaultColor-缺省病人的显示颜色
    '返回:True-设置颜色成功，False-失败
    '编制:李南春
    '日期:2014-07-08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngColor As Long
    
    lngColor = lngDefaultColor
    If str病人类型 <> "" Then
        lngColor = zlDatabase.GetPatiColor(str病人类型)
    End If
    objPatiControl.ForeColor = lngColor
    SetPatiColor = True
End Function

Public Function FromatSQL(ByVal strText As String, Optional ByVal blnCrlf As Boolean) As String
'功能：去掉TAB字符，两边空格，回车，最后只由单空格分隔。
'参数：strText=处理字符
'         blnCrlf=是否去掉换行符
    Dim i As Long
    
    If blnCrlf Then
        strText = Replace(strText, vbCrLf, " ")
        strText = Replace(strText, vbCr, " ")
        strText = Replace(strText, vbLf, " ")
    End If
    strText = Trim(Replace(strText, vbTab, " "))
    
    i = 5
    Do While i > 1
        strText = Replace(strText, String(i, " "), " ")
        If InStr(strText, String(i, " ")) = 0 Then i = i - 1
    Loop
    FromatSQL = strText
End Function

Public Function SQLAdjust(Str As String) As String
'功能：将含有"'"符号的字符串调整为Oracle所能识别的字符常量
'说明：自动(必须)在两边加"'"界定符。

    Dim i As Long, strTmp As String
    
    If InStr(1, Str, "'") = 0 Then SQLAdjust = "'" & Str & "'": Exit Function
    
    For i = 1 To Len(Str)
        If Mid(Str, i, 1) = "'" Then
            If i = 1 Then
                strTmp = "CHR(39)||'"
            ElseIf i = Len(Str) Then
                strTmp = strTmp & "'||CHR(39)"
            Else
                strTmp = strTmp & "'||CHR(39)||'"
            End If
        Else
            If i = 1 Then
                strTmp = "'" & Mid(Str, i, 1)
            ElseIf i = Len(Str) Then
                strTmp = strTmp & Mid(Str, i, 1) & "'"
            Else
                strTmp = strTmp & Mid(Str, i, 1)
            End If
        End If
    Next
    SQLAdjust = strTmp
End Function

Public Function RoundEx(ByVal dblNumber As Double, ByVal intBit As Integer) As Double
'功能：四舍五入方式格式化数字
'参数：intBit=最大小数位数
'问题号：94552
'说明：VB自带的Round是银行家舍入法,与实际不一致。如Round(57.575,2)=57.58,Round(57.565,2)=57.56
    If intBit > 0 Then
        RoundEx = Val(Format(dblNumber, "0." & String(intBit, "0")))
    Else
        RoundEx = dblNumber
    End If
End Function
