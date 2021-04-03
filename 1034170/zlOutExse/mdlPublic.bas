Attribute VB_Name = "mdlPublic"
Option Explicit '要求变量声明
'系统公用临时变量
Public glngSys As Long
Public glngModul As Long
Public gstrPrivs As String                   '当前用户具有的当前模块的功能

Public gstrSQL As String
Public gblnOK As Boolean

Public gstrSysName As String                '系统名称
Public gstrUnitName As String '用户单位名称
Public gstrDBUser As String '当前数据库用户名
Public gfrmMain As Object
Public glngMain As Long
Public gcnOracle As ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例

'-----------------------------------------
Public Type TYPE_USER_INFO
    ID As Long
    部门ID As Long
    部门名称 As String
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
End Type
Public UserInfo As TYPE_USER_INFO


Public Enum gRegType
    g注册信息 = 0
    g公共全局 = 1
    g公共模块 = 2
    g私有全局 = 3
    g私有模块 = 4
End Enum

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
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
'----------------------------------------
Public Const LONG_MAX = 2147483647 'Long型最大值
Public glngTXTProc As Long '保存默认的消息函数的地址
Public glngOld As Long, glngFormW As Long, glngFormH As Long

'下列语句用于检测是否合法调用
Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Public Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'Windows风格----------------------------------
Public Const ETO_OPAQUE = 2
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
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
'---------------------------------------------
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Public Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const GWL_WNDPROC = -4
Public Const WM_GETMINMAXINFO = &H24
Public Const WM_CONTEXTMENU = &H7B ' 当右击文本框时，产生这条消息
Public Declare Sub GetKeyboardStateByString Lib "user32" Alias "GetKeyboardState" (ByVal pbKeyState As String)
Public Declare Sub SetKeyboardStateByString Lib "user32" Alias "SetKeyboardState" (ByVal lppbKeyState As String)
Public Const VK_CAPITAL = &H14
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const CB_FINDSTRING = &H14C
Public Const CB_GETDROPPEDSTATE = &H157
Public Const CB_SHOWDROPDOWN = &H14F
Public Const LVSCW_AUTOSIZE = -1
Public Const LVSCW_AUTOSIZE_USEHEADER = -2
Public Const LVM_SETCOLUMNWIDTH = &H101E
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
'输入法控制API----------------------------------------------------------------------------------------------
Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long
Public Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
Public Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
Public Declare Function GetKeyboardLayoutName Lib "user32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long
Public Declare Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal flags As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Boolean
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const KLF_REORDER = &H8
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2 '浅凹下
Public Const BDR_RAISEDINNER = &H4 '浅凸起
Public Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER) '深凸起
Public Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER) '深凹下
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER) 'Frame边线样式
Public Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER) '反Frame边线样式
Public Const BF_LEFT = &H1
Public Const BF_TOP = &H2
Public Const BF_RIGHT = &H4
Public Const BF_BOTTOM = &H8
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_SOFT = &H1000
Private Type TY_WindowsRect
    MaxW As Long
    MaxH As Long
    MinW  As Long
    MinH As Long
End Type
Public gWinRect As TY_WindowsRect

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
Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function ZVal(ByVal varValue As Variant, Optional ByVal varDefault As Variant = 0) As String
'功能：将0零转换为"NULL"串,在生成SQL语句时用
    Dim varTmp As Variant
    varTmp = IIf(Val(varValue) = 0, varDefault, varValue)
    ZVal = IIf(Val(varTmp) = 0, "NULL", varTmp)
End Function

Public Sub RaisEffect(picBox As PictureBox, Optional intStyle As Integer, Optional strName As String = "")
'功能：将PictureBox模拟成3D平面按钮
'参数：intStyle:0=平面,-1=凹下,1=凸起
    
    Dim PicRect As RECT
    Dim lngTmp As Long
    With picBox
        .Cls
        lngTmp = .ScaleMode
        .ScaleMode = 3
        .BorderStyle = 0
        If intStyle <> 0 Then
            PicRect.Left = .ScaleLeft
            PicRect.Top = .ScaleTop
            PicRect.Right = .ScaleWidth
            PicRect.Bottom = .ScaleHeight
            DrawEdge .hDC, PicRect, CLng(IIf(intStyle = 1, BDR_RAISEDINNER Or BF_SOFT, BDR_SUNKENOUTER Or BF_SOFT)), BF_RECT
        End If
        .ScaleMode = lngTmp
        If strName <> "" Then
            .CurrentX = (.ScaleWidth - .TextWidth(strName)) / 2
            .CurrentY = (.ScaleHeight - .TextHeight(strName)) / 2
            picBox.Print strName
        End If
    End With
End Sub

Public Function SystemImes() As Variant
'功能：将系统中文输入法名称返回到一个字符串数组中
'返回：如果不存在中文输入法,则返回空串
    Dim arrIme(99) As Long, arrName() As String
    Dim lngLen As Long, strName As String * 255
    Dim lngCount As Long, i As Integer, j As Integer
    
    lngCount = GetKeyboardLayoutList(UBound(arrIme) + 1, arrIme(0))
    For i = 0 To lngCount - 1
        If ImmIsIME(arrIme(i)) = 1 Then
            ReDim Preserve arrName(j)
            lngLen = ImmGetDescription(arrIme(i), strName, Len(strName))
            arrName(j) = Mid(strName, 1, InStr(strName, Chr(0)) - 1)
            j = j + 1
        End If
    Next
    SystemImes = IIf(j > 0, arrName, vbNullString)
End Function
 

Public Sub AutoSizeCol(lvw As Object)
'功能：根据自动ListView当前内容自动调整各列宽度
'参数：blnByHead=是否按列头文本调整,Col=指定列还是所有列(1-N)
    Dim i As Integer, lngW As Long
    For i = 1 To lvw.ColumnHeaders.Count
        SendMessage lvw.hWnd, LVM_SETCOLUMNWIDTH, i - 1, LVSCW_AUTOSIZE
        If lvw.ColumnHeaders(i).Width < 200 Then lvw.ColumnHeaders(i).Width = 0
        If lvw.ColumnHeaders(i).Width < (zlCommFun.ActualLen(lvw.ColumnHeaders(i).Text) + 2) * 90 And lvw.ColumnHeaders(i).Width <> 0 Then lvw.ColumnHeaders(i).Width = (zlCommFun.ActualLen(lvw.ColumnHeaders(i).Text) + 2) * 90
    Next
End Sub

Public Function GetCboIndex(cbo As ComboBox, strFind As String, Optional blnKeep As Boolean, Optional blnLike As Boolean) As Long
'功能：由字符串在ComboBox中查找索引
    Dim i As Long
    If strFind = "" Then GetCboIndex = -1: Exit Function
    '先精确查找
    For i = 0 To cbo.ListCount - 1
        If InStr(cbo.List(i), "-") > 0 Then
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

Public Function FindCboIndex(cbo As ComboBox, ByVal lngID As Long) As Long
'功能：由项目数据查找ComboBox的索引值
'参数：lngID=ComboBox的项目值
    Dim i As Integer
    If lngID = 0 Then FindCboIndex = -1: Exit Function
    For i = 0 To cbo.ListCount - 1
        If cbo.ItemData(i) = lngID Then
            FindCboIndex = i
            Exit Function
        End If
    Next
    FindCboIndex = -1
End Function

Public Function SetWidth(cboHwnd As Long, NewWidthPixel As Long) As Boolean
'功能：设置 Combo 下拉的宽度,单位为 pixels
    Dim lRetVal As Long
    lRetVal = SendMessage(cboHwnd, &H160, NewWidthPixel, 0)
    If lRetVal <> -1 Then
        SetWidth = True
    Else
        SetWidth = False
    End If
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

Public Function PreFixNO(Optional Curdate As Date = #1/1/1900#) As String
'功能：返回大写的单据号年前缀
    If Curdate = #1/1/1900# Then
        PreFixNO = CStr(CInt(Format(zlDatabase.Currentdate, "YYYY")) - 1990)
    Else
        PreFixNO = CStr(CInt(Format(Curdate, "YYYY")) - 1990)
    End If
    PreFixNO = IIf(CInt(PreFixNO) < 10, PreFixNO, Chr(55 + CInt(PreFixNO)))
End Function

Public Sub SelAll(objTXT As Control, Optional blnIme As Boolean = False)
'功能：对文本框的的文本选中
    If TypeName(objTXT) = "TextBox" Then
        objTXT.SelStart = 0: objTXT.SelLength = Len(objTXT.Text)
    ElseIf TypeName(objTXT) = "MaskEdBox" Then
        If Not IsDate(objTXT.Text) Then
            objTXT.SelStart = 0: objTXT.SelLength = Len(objTXT.Text)
        Else
            objTXT.SelStart = 0: objTXT.SelLength = 10
        End If
    End If
End Sub

Public Function HaveExist(cbo As ComboBox, str As String, lng As Long) As Boolean
'功能：判断指定项目在列表中是否已经存在
'说明：相同项目指Text及ItemData都相同
    Dim i As Long
    For i = 0 To cbo.ListCount
        If cbo.List(i) = str And cbo.ItemData(i) = lng Then
            HaveExist = True: Exit For
        End If
    Next
End Function

Public Function NeedCode(strList As String) As String
    '从字符串中分离出编码
    If InStr(strList, "-") > 0 Then
        NeedCode = LTrim(Left(strList, InStr(strList, "-") - 1))
    End If
End Function

Public Function NeedName(strList As String) As String
    '从字符串中分离出名称
    If InStr(strList, "]") > 0 And InStr(strList, "-") = 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, "]") + 1))
    ElseIf InStr(strList, ")") > 0 And InStr(strList, "-") = 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, ")") + 1))
    Else
        NeedName = LTrim(Mid(strList, InStr(strList, "-") + 1))
    End If
End Function

Public Function GetFirstRow(objBill As ExpenseBill, intPage As Long, Optional ByVal strClass As String) As Integer
'功能：获取当前单据中第一个为药品的行号
'参数：strClass=是否只取指定类别药品行
'返回：0=没有药品收费行
    Dim i As Integer
    
    For i = 1 To objBill.Pages(intPage).Details.Count
        If strClass = "" Then
            If InStr(",5,6,7,", objBill.Pages(intPage).Details(i).收费类别) > 0 Then
                GetFirstRow = i: Exit Function
            End If
        Else
            If objBill.Pages(intPage).Details(i).收费类别 = strClass Then
                GetFirstRow = i: Exit Function
            End If
        End If
    Next
End Function

Public Function CentMoney(ByVal curMoney As Currency) As Currency
'功能：对指定金额按分币处理规则进行处理,返回处理后的金额
'参数：curMoney=要进行分币处理的金额(为应缴金额,2位小数)
'      gBytMoney=
'         0.不处理
'         1.采取四舍五入法,eg:0.51=0.50;0.56=0.60
'         2.补整收法,eg:0.51=0.60,0.56=0.60
'         3.舍分收法,eg:0.51=0.50,0.56=0.50
'         4.四舍六入五成双,eg:0.14=0.10,0.16=0.20,0.151=0.20,0.15=0.20,0.25=0.20
'           四舍六入五成双,详见我国科学技术委员会正式颁布的《数字修约规则》,但根据vb的Round函数,若被舍弃的数字包括几位数字时，不对该数字进行连续修约
'           即银行家舍入法:四舍六入五考虑，五后非零就进一，五后皆零看奇偶，五前为偶应舍去，五前为奇要进一
'         5.三七作五、二舍八入,对角进行处理，不需要先对分币进行舍入,即0.29(含)以下都舍掉角，0.80(含)以上都进角，0.3-0.79处理为0.5。
'         6.五舍六入:eg:0.15=0.10:0.16=0.2:刘兴洪:34519
'91385,调整“5.三七作五、二舍八入”规则：先对分币进行四舍五入，即0.24(含)以下都舍掉角，0.75(含)以上都进角，0.25-0.74都处理为0.5
'       分币先四舍五入，那么0.00～0.24=0，0.25～0.5=0.50, 0.50～0.74=0.50，0.75～1.00=1，这样舍和入各占50%的比例
    
    Dim intSign As Integer, curTmp As Currency

    If gBytMoney = 0 Then
        CentMoney = Format(curMoney, "0.00")
    ElseIf gBytMoney = 1 Then
        curMoney = Format(curMoney, "0.00")    '先取两位金额,再处理分币,如:0.248 得0.3
        CentMoney = Format(curMoney, "0.0")
    ElseIf gBytMoney = 2 Then
        intSign = Sgn(curMoney)
        curMoney = Abs(curMoney)
        If Int(curMoney * 10) / 10 = curMoney Then
            CentMoney = intSign * curMoney
        Else
            CentMoney = intSign * Int(curMoney * 10 + 1) / 10
        End If
    ElseIf gBytMoney = 3 Then
        intSign = Sgn(curMoney)
        curMoney = Abs(curMoney)
        curMoney = Int(curMoney * 10) / 10
        CentMoney = intSign * curMoney
    ElseIf gBytMoney = 4 Then
        CentMoney = Format(Round(curMoney, 1), "0.00")
    ElseIf gBytMoney = 5 Then
        intSign = Sgn(curMoney)
        curMoney = Abs(curMoney)
        curTmp = Format(curMoney - Int(curMoney), "0.0")
        If curTmp >= 0.8 Then
            curTmp = 1
        ElseIf curTmp < 0.3 Then
            curTmp = 0
        Else
            curTmp = 0.5
        End If
        CentMoney = intSign * Format(Int(curMoney) + curTmp, "0.00")
    ElseIf gBytMoney = 6 Then
         '刘兴洪 问题:34519 五舍六入:eg:0.15=0.10:0.16=0.2:    日期:2010-12-06 09:58:02
          CentMoney = Format(Format(curMoney - 0.01, "0.0"), "0.00")
    End If
End Function

Public Function Custom_WndMessage(ByVal hWnd As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
'功能：自定义消息函数处理窗体尺寸调整限制
    If msg = WM_GETMINMAXINFO Then
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
    Custom_WndMessage = CallWindowProc(glngOld, hWnd, msg, wp, lp)
End Function

Public Function InDesign() As Boolean
    On Error Resume Next
    Debug.Print 1 / 0
    If Err.Number <> 0 Then Err.Clear: InDesign = True
End Function

Public Function strPad(ByVal strPre As String, ByVal intLen As Integer, ByVal strFill As String, ByVal bytAlign As Byte, Optional ByVal blnTrim As Boolean) As String
'功能：填充字符串
'参数：
'     strPre=要填充的字符串
'     intLen=填充后的长度
'     strFill=要填充的字符
'     bytAlign=1,2/左,右对齐，左对齐时，在原字符串右边填充
'     blnTrim=当字符串超长时，是否强行按指定长度截取。
'返回：处理后的字符串
'说明：一个汉字当作两个字符长度处理
    Dim i As Integer
    
    If LenB(StrConv(strPre, vbFromUnicode)) >= intLen Then
        If blnTrim Then
            For i = 1 To Len(strPre)
                strPad = strPad & Mid(strPre, i, 1)
                If LenB(StrConv(strPad, vbFromUnicode)) >= intLen Then Exit For
            Next
        Else
            strPad = strPre
        End If
    Else
        If Len(strFill) > 1 Then strFill = Left(strFill, 1)
        If bytAlign = 1 Then
            strPad = strPre
            For i = 1 To intLen - LenB(StrConv(strPre, vbFromUnicode))
                strPad = strPad & strFill
            Next
        ElseIf bytAlign = 2 Then
            For i = 1 To intLen - LenB(StrConv(strPre, vbFromUnicode))
                strPad = strPad & strFill
            Next
            strPad = strPad & strPre
        End If
    End If
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

Public Sub ExChangeLocate(objA As Object, objB As Object)
'功能：交换医生和开单科室的输入位置
    Dim x1 As Long, y1 As Long, w1 As Long, t1 As Integer
    Dim x2 As Long, y2 As Long, w2 As Long, t2 As Integer
    Dim obj1 As Object, obj2 As Object
    
    x1 = objA.Left
    y1 = objA.Top
    w1 = objA.Width
    t1 = objA.TabIndex
    Set obj1 = objA.Container

    x2 = objB.Left
    y2 = objB.Top
    w2 = objB.Width
    t2 = objB.TabIndex
    Set obj2 = objB.Container
    
    Set objB.Container = obj1
    objB.Left = x1
    objB.Top = y1
    objB.Width = w1
    objB.TabIndex = t1
    
    Set objA.Container = obj2
    objA.Left = x2
    objA.Top = y2
    objA.Width = w2
    objA.TabIndex = t2
End Sub

'去掉TextBox的默认右键菜单
Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' 如果消息不是WM_CONTEXTMENU，就调用默认的窗口函数处理
    If msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(glngTXTProc, hWnd, msg, wp, lp)
End Function

Public Function MakeBillRecord(objBill As ExpenseBill, ByVal bln急诊 As Boolean, ByVal intPage As Integer, _
    ByVal strDate As String, ByVal str费别 As String, ByVal strInvoice As String) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据单据对象内容创建一个记录信息(以售价单位)
    '入参:intPage=多单据收费模式时，指定的单据,如果为零,表示全部数据
    '        strDate=结算时间,
    '        strInvoice=票据号
    '出参:
    '返回:医保相关数据的数据集(单据序号(1--n),病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保)
    '编制:刘兴洪
    '日期:2011-08-15 16:40:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer, intStartPage As Integer, intPages As Integer
    Dim p As Integer, strSQL As String
    Dim dbl单价 As Double, cur实收 As Currency, cur统筹 As Currency
    Dim rsTmp As New ADODB.Recordset, rsNo As ADODB.Recordset
    
    Err = 0: On Error GoTo ErrHand:
    rsTmp.Fields.Append "单据序号", adBigInt, 50, adFldIsNullable
    rsTmp.Fields.Append "费别", adVarChar, 50, adFldIsNullable
    rsTmp.Fields.Append "NO", adVarChar, 8, adFldIsNullable
    rsTmp.Fields.Append "序号", adBigInt, , adFldIsNullable '问题:42961
    rsTmp.Fields.Append "实际票号", adVarChar, 20, adFldIsNullable
    rsTmp.Fields.Append "结算时间", adDBTimeStamp, , adFldIsNullable
    rsTmp.Fields.Append "病人ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "收费类别", adVarChar, 2, adFldIsNullable
    rsTmp.Fields.Append "收据费目", adVarChar, 20, adFldIsNullable
    rsTmp.Fields.Append "计算单位", adVarChar, 50, adFldIsNullable
    '69788:李南春,2014-6-5,调整开单人字段大小，由20改为100
    rsTmp.Fields.Append "开单人", adVarChar, 100, adFldIsNullable
    rsTmp.Fields.Append "收费细目ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "数量", adDouble, , adFldIsNullable
    rsTmp.Fields.Append "单价", adDouble, , adFldIsNullable
    rsTmp.Fields.Append "实收金额", adCurrency, , adFldIsNullable
    rsTmp.Fields.Append "统筹金额", adCurrency, , adFldIsNullable
    rsTmp.Fields.Append "保险支付大类ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "是否医保", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "保险编码", adVarChar, 50, adFldIsNullable
    rsTmp.Fields.Append "摘要", adVarChar, 2000, adFldIsNullable
    rsTmp.Fields.Append "是否急诊", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "开单部门ID", adBigInt, , adFldIsNullable
    rsTmp.Fields.Append "执行部门ID", adBigInt, , adFldIsNullable
    rsTmp.CursorLocation = adUseClient
    rsTmp.LockType = adLockOptimistic
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    intStartPage = IIf(intPage <= 0, 1, intPage)
    intPages = IIf(intPage <= 0, objBill.Pages.Count, intPage)
    For p = intStartPage To intPages
         If objBill.Pages(p).NO <> "" Then       '提取的是划价单
                '提取的划价单(售价单位)
                strSQL = _
                "Select '" & strInvoice & "' as 实际票号,NO,Nvl( 价格父号, 序号) as 序号,To_Date('" & strDate & "','YYYY-MM-DD HH24:MI:SS') as 结算时间," & _
                        objBill.病人ID & " As 病人ID,'" & str费别 & "' As 费别,收费类别,收据费目,计算单位,开单人," & _
                "       收费细目ID,保险大类ID As 保险支付大类ID,Nvl(保险项目否,0) As 是否医保,保险编码," & _
                "       Avg(Nvl(付数,0)*数次) As 数量,Avg(标准单价) As 单价," & _
                "       Sum(实收金额) As 实收金额,Sum(统筹金额) As 统筹金额,摘要," & _
                        IIf(bln急诊, "1", "0") & " as 是否急诊,开单部门ID,执行部门ID From 门诊费用记录" & _
                " Where 记录性质=1 And 记录状态=0 And NO=[1]" & _
                " Group By Nvl(价格父号,序号),收费类别,收据费目,计算单位,开单人," & _
                "       收费细目ID,保险大类ID,Nvl(保险项目否,0),保险编码,摘要,开单部门ID,执行部门ID,NO" & _
                " Order by  序号 "
                Set rsNo = zlDatabase.OpenSQLRecord(strSQL, "获取划价单数据-医保", objBill.Pages(p).NO)
                If rsNo.RecordCount <> 0 Then rsNo.MoveFirst
                Do While Not rsNo.EOF
                    rsTmp.AddNew
                    rsTmp!单据序号 = p
                    rsTmp!费别 = str费别
                    rsTmp!NO = Nvl(rsNo!NO)   '仅提取划价单时才有值
                    rsTmp!序号 = Val(Nvl(rsNo!序号))   '仅提取划价单时才有值
                    rsTmp!实际票号 = strInvoice
                    rsTmp!结算时间 = CDate(strDate)
                    rsTmp!病人ID = IIf(objBill.病人ID = 0, Null, objBill.病人ID)
                    rsTmp!收费类别 = Nvl(rsNo!收费类别)
                    rsTmp!收据费目 = Nvl(rsNo!收据费目)
                    rsTmp!开单人 = Nvl(rsNo!开单人)
                    rsTmp!收费细目ID = Val(Nvl(rsNo!收费细目ID))
                    rsTmp!计算单位 = Nvl(rsNo!计算单位)
                    rsTmp!数量 = Val(Nvl(rsNo!数量))
                    rsTmp!单价 = Val(Nvl(rsNo!单价))
                    rsTmp!实收金额 = Val(Nvl(rsNo!实收金额))
                    rsTmp!统筹金额 = Val(Nvl(rsNo!统筹金额))
                    rsTmp!保险支付大类ID = IIf(Val(Nvl(rsNo!保险支付大类ID)) = 0, Null, Val(Nvl(rsNo!保险支付大类ID)))
                    rsTmp!是否医保 = Val(Nvl(rsNo!是否医保))
                    rsTmp!保险编码 = Nvl(rsNo!保险编码)
                    rsTmp!摘要 = Nvl(rsNo!摘要)
                    rsTmp!是否急诊 = IIf(bln急诊, 1, 0)
                    rsTmp!开单部门ID = Val(Nvl(rsNo!开单部门ID))
                    rsTmp!执行部门ID = Val(Nvl(rsNo!执行部门ID))
                    rsTmp.Update
                    rsNo.MoveNext
                Loop
         Else
            For i = 1 To objBill.Pages(p).Details.Count
                dbl单价 = 0: cur实收 = 0: cur统筹 = 0
                With objBill.Pages(p).Details(i)
                    For j = 1 To .InComes.Count
                        dbl单价 = dbl单价 + .InComes(j).标准单价
                        cur实收 = cur实收 + .InComes(j).实收金额
                        cur统筹 = cur统筹 + .InComes(j).统筹金额
                    Next
                    rsTmp.AddNew
                    rsTmp!单据序号 = p
                    rsTmp!费别 = str费别
                    rsTmp!NO = ""   '仅提取划价单时才有值
                    rsTmp!序号 = i
                    rsTmp!实际票号 = strInvoice
                    rsTmp!结算时间 = CDate(strDate)
                    rsTmp!病人ID = IIf(objBill.病人ID = 0, Null, objBill.病人ID)
                    rsTmp!收费类别 = .收费类别
                    If .InComes.Count > 0 Then
                        rsTmp!收据费目 = .InComes(1).收据费目
                    Else
                        rsTmp!收据费目 = Null
                    End If
                    rsTmp!开单人 = objBill.Pages(p).开单人
                    
                    rsTmp!收费细目ID = .收费细目ID
                    
                    rsTmp!计算单位 = .计算单位
                    If InStr(",5,6,7,", .收费类别) > 0 And gbln药房单位 Then
                        '从药房单位转换为售价单位
                        rsTmp!数量 = IIf(.付数 = 0, 1, .付数) * .数次 * .Detail.药房包装
                        rsTmp!单价 = Format(dbl单价 / .Detail.药房包装, gstrFeePrecisionFmt)
                    Else
                        rsTmp!数量 = IIf(.付数 = 0, 1, .付数) * .数次
                        rsTmp!单价 = Format(dbl单价, gstrFeePrecisionFmt)
                    End If
                    rsTmp!实收金额 = Format(cur实收, gstrDec)
                    rsTmp!统筹金额 = Format(cur统筹, gstrDec)
                    rsTmp!保险支付大类ID = IIf(.保险大类ID = 0, Null, .保险大类ID)
                    rsTmp!是否医保 = IIf(.保险项目否, 1, 0)
                    rsTmp!保险编码 = .保险编码
                    rsTmp!摘要 = .摘要
                    rsTmp!是否急诊 = IIf(bln急诊, 1, 0)
                    rsTmp!开单部门ID = objBill.Pages(p).开单部门ID
                    rsTmp!执行部门ID = .执行部门ID
                    rsTmp.Update
                End With
            Next
        End If
    Next
    If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    Set MakeBillRecord = rsTmp
    Exit Function
ErrHand::
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function zlCreateFeeListStruc(ByRef rsFeelists As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建本地的费用记录集结构
    '入参:
    '出参:rsFeelists-返回本地记录集结构,同时打开了记录集的
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-01-05 16:18:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    Set rsFeelists = New ADODB.Recordset
    
    rsFeelists.Fields.Append "单据序号", adBigInt, , adFldIsNullable
    rsFeelists.Fields.Append "费别", adVarChar, 50, adFldIsNullable
    rsFeelists.Fields.Append "NO", adVarChar, 8, adFldIsNullable
    rsFeelists.Fields.Append "实际票号", adVarChar, 20, adFldIsNullable
    rsFeelists.Fields.Append "结算时间", adDBTimeStamp, , adFldIsNullable
    rsFeelists.Fields.Append "病人ID", adBigInt, , adFldIsNullable
    rsFeelists.Fields.Append "收费类别", adVarChar, 2, adFldIsNullable
    rsFeelists.Fields.Append "收据费目", adVarChar, 20, adFldIsNullable
    rsFeelists.Fields.Append "计算单位", adVarChar, 50, adFldIsNullable
    '69788:李南春,2014-6-5,调整开单人字段大小，由20改为100
    rsFeelists.Fields.Append "开单人", adVarChar, 100, adFldIsNullable
    rsFeelists.Fields.Append "收费细目ID", adBigInt, , adFldIsNullable
    rsFeelists.Fields.Append "数量", adDouble, , adFldIsNullable
    rsFeelists.Fields.Append "单价", adDouble, , adFldIsNullable
    rsFeelists.Fields.Append "实收金额", adCurrency, , adFldIsNullable
    rsFeelists.Fields.Append "统筹金额", adCurrency, , adFldIsNullable
    rsFeelists.Fields.Append "保险支付大类ID", adBigInt, , adFldIsNullable
    rsFeelists.Fields.Append "是否医保", adBigInt, , adFldIsNullable
    rsFeelists.Fields.Append "保险编码", adVarChar, 50, adFldIsNullable
    rsFeelists.Fields.Append "摘要", adVarChar, 2000, adFldIsNullable
    rsFeelists.Fields.Append "是否急诊", adBigInt, , adFldIsNullable
    rsFeelists.Fields.Append "开单部门ID", adBigInt, , adFldIsNullable
    rsFeelists.Fields.Append "执行部门ID", adBigInt, , adFldIsNullable
    rsFeelists.Fields.Append "本次结算", adDouble, , adFldIsNullable
    rsFeelists.CursorLocation = adUseClient
    rsFeelists.LockType = adLockOptimistic
    rsFeelists.CursorType = adOpenStatic
    rsFeelists.Open
    zlCreateFeeListStruc = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function zlBuldingFeeListdata(objBill As ExpenseBill, ByVal bln急诊 As Boolean, ByVal intPage As Integer, _
    ByVal strDate As String, ByVal str费别 As String, ByVal strInvoice As String, ByRef rsFeelists As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据单据对象内容创建一个记录信息(以售价单位)
    '入参:intPage=多单据收费模式时，指定的单据
    '     strDate=结算时间,
    '     strInvoice=票据号
    '出参:rsFeeLists-返回费用记录集( 单据序号(以单据为准),病人ID,收费细目ID,数量,单价,实收金额,统筹金额,保险支付大类ID,是否医保,本次结算(返回))
    '返回:
    '编制:刘兴洪
    '日期:2010-01-05 16:11:44
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim dbl单价 As Double, cur实收 As Currency, cur统筹 As Currency
    Err = 0: On Error GoTo ErrHand:
    For i = 1 To objBill.Pages(intPage).Details.Count
        dbl单价 = 0: cur实收 = 0: cur统筹 = 0
        With objBill.Pages(intPage).Details(i)
            For j = 1 To .InComes.Count
                dbl单价 = dbl单价 + .InComes(j).标准单价
                cur实收 = cur实收 + .InComes(j).实收金额
                cur统筹 = cur统筹 + .InComes(j).统筹金额
            Next
            rsFeelists.AddNew
            rsFeelists!单据序号 = intPage
            rsFeelists!费别 = str费别
            rsFeelists!NO = ""   '仅提取划价单时才有值
            rsFeelists!实际票号 = strInvoice
            rsFeelists!结算时间 = CDate(strDate)
            rsFeelists!病人ID = IIf(objBill.病人ID = 0, Null, objBill.病人ID)
            rsFeelists!收费类别 = .收费类别
            
            If .InComes.Count > 0 Then
                rsFeelists!收据费目 = .InComes(1).收据费目
            Else
                rsFeelists!收据费目 = Null
            End If
            rsFeelists!开单人 = objBill.Pages(intPage).开单人
            
            rsFeelists!收费细目ID = .收费细目ID
            
            rsFeelists!计算单位 = .计算单位
            If InStr(",5,6,7,", .收费类别) > 0 And gbln药房单位 Then
                '从药房单位转换为售价单位
                rsFeelists!数量 = IIf(.付数 = 0, 1, .付数) * .数次 * .Detail.药房包装
                rsFeelists!单价 = Format(dbl单价 / .Detail.药房包装, gstrFeePrecisionFmt)
            Else
                rsFeelists!数量 = IIf(.付数 = 0, 1, .付数) * .数次
                rsFeelists!单价 = Format(dbl单价, gstrFeePrecisionFmt)
            End If
            rsFeelists!实收金额 = Format(cur实收, gstrDec)
            rsFeelists!统筹金额 = Format(cur统筹, gstrDec)
            rsFeelists!保险支付大类ID = IIf(.保险大类ID = 0, Null, .保险大类ID)
            rsFeelists!是否医保 = IIf(.保险项目否, 1, 0)
            rsFeelists!保险编码 = .保险编码
            rsFeelists!摘要 = .摘要
            rsFeelists!是否急诊 = IIf(bln急诊, 1, 0)
            rsFeelists!开单部门ID = objBill.Pages(intPage).开单部门ID
            rsFeelists!执行部门ID = .执行部门ID
            rsFeelists!本次结算 = 0
            rsFeelists.Update
        End With
    Next
    If rsFeelists.RecordCount > 0 Then rsFeelists.MoveFirst
    zlBuldingFeeListdata = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function


Public Function TrimChar(ByVal str As String, ByVal Chr As String) As String
'功能：去除str两边的chr,功能类似Trim
    Dim i As Integer, intB As Integer, intE As Integer
    
    If str = "" Or Chr = "" Then TrimChar = str: Exit Function
        
    intB = 1
    For i = 1 To Len(str)
        If Mid(str, i, 1) <> Chr Then intB = i: Exit For
    Next
    intE = Len(str)
    For i = Len(str) To 1 Step -1
        If Mid(str, i, 1) <> Chr Then intE = i: Exit For
    Next
    TrimChar = Mid(str, intB, intE - intB + 1)
End Function

Public Function GetBill费别(objBill As ExpenseBill) As String
'功能：如果单据中所有行费别一致，则返回该费别,否则返回空
    Dim i As Integer, p As Integer, strTmp As String
    
    For p = 1 To objBill.Pages.Count
        For i = 1 To objBill.Pages(p).Details.Count
            If i = 1 Then
                strTmp = objBill.Pages(p).Details(i).费别
            ElseIf objBill.Pages(p).Details(i).费别 <> strTmp Then
                Exit Function
            End If
        Next
    Next
    GetBill费别 = strTmp
End Function

Public Function GetDrugTotal(ByVal objBill As ExpenseBill, ByVal lng药品ID As Long, ByVal lng药房ID As Long, Optional ByVal intPage As Integer) As Double
'功能：获取单据中指定药品在同一药房多行的数量合
'参数： lng药房ID-0表示分离发药时,不限定药房检查
    Dim i As Integer, p As Integer, dblCount As Double
    
    For p = 1 To objBill.Pages.Count
        If intPage = 0 Or p = intPage Then
            For i = 1 To objBill.Pages(p).Details.Count
                If objBill.Pages(p).Details(i).收费细目ID = lng药品ID And _
                    IIf(lng药房ID <> 0, objBill.Pages(p).Details(i).执行部门ID = lng药房ID, 1 = 1) Then
                    dblCount = dblCount + objBill.Pages(p).Details(i).付数 * objBill.Pages(p).Details(i).数次
                End If
            Next
        End If
    Next
    GetDrugTotal = dblCount
End Function
Public Function FormatEx(ByVal vNumber As Variant, ByVal intBit As Integer, _
    Optional blnShowZero As Boolean = True, _
    Optional intMinBit As Integer) As String
'功能：四舍五入方式格式化显示数字,保证小数点最后不出现0,小数点前要有0
'参数：vNumber=Single,Double,Currency类型的数字,intBit=最大小数位数
'       intMinBit=最少保留小数位数
    Dim strNumber As String
    Dim intDot As Integer
            
    If TypeName(vNumber) = "String" Then
        If vNumber = "" Then Exit Function
        If Not IsNumeric(vNumber) Then Exit Function
        vNumber = Val(vNumber)
    End If
    
    If intBit < intMinBit Then intMinBit = intBit
            
    If vNumber = 0 Then
        If blnShowZero Then
            If intMinBit > 0 Then
                strNumber = "0." & String(intMinBit, "0")
            Else
                strNumber = "0"
            End If
        Else
            strNumber = ""
        End If
    ElseIf Int(vNumber) = vNumber Then
        If intMinBit > 0 Then
            strNumber = vNumber & "." & String(intMinBit, "0")
        Else
            strNumber = vNumber
        End If
    Else
        intDot = intBit - intMinBit
        strNumber = Format(vNumber, "0." & String(intBit, "0"))
        If Left(strNumber, 1) = "." Then strNumber = "0" & strNumber
        If InStr(strNumber, ".") > 0 Then
            Do While Right(strNumber, 1) = "0" And intDot > 0
                strNumber = Left(strNumber, Len(strNumber) - 1)
                intDot = intDot - 1
            Loop
            If Right(strNumber, 1) = "." Then strNumber = Left(strNumber, Len(strNumber) - 1)
        End If
    End If
    FormatEx = strNumber
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

Public Function IntEx(vNumber As Variant) As Variant
'功能：取大于指定数值的最小整数
    IntEx = -1 * Int(-1 * Val(vNumber))
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

Public Sub SaveRegisterItem(ByVal RegType As gRegType, ByVal strSection As String, _
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
    On Error GoTo ErrHand:
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
ErrHand:
End Sub
Public Sub GetRegisterItem(ByVal RegType As gRegType, ByVal strSection As String, _
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
    On Error GoTo ErrHand:
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
ErrHand:
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

Public Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal lngY As Long) As POINTAPI
'功能：得控件中指定坐标在屏幕中的位置(Twip)
    Dim vPoint As POINTAPI
    vPoint.X = lngX / Screen.TwipsPerPixelX: vPoint.Y = lngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.X = vPoint.X * Screen.TwipsPerPixelX: vPoint.Y = vPoint.Y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
End Function

Public Function SysColor2RGB(ByVal lngColor As Long) As Long
'功能：将VB的系统颜色转换为RGB色
    If lngColor < 0 Then
        Call OleTranslateColor(lngColor, 0, lngColor)
    End If
    SysColor2RGB = lngColor
End Function

Public Function Between(X, a, B) As Boolean
'功能：判断x是否在a和b之间
    If a < B Then
        Between = X >= a And X <= B
    Else
        Between = X >= B And X <= a
    End If
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

Public Function MoneyOverFlow(objBill As ExpenseBill) As Boolean
'功能：检查单据合计金额是否溢出
'说明：以Currency上限922337203685477为准
    Dim dbl应收 As Double, dbl实收 As Double
    Dim i As Integer, j As Integer, k As Integer
    
    '要用VAL转为Double进行运算
    For i = 1 To objBill.Pages.Count
        For j = 1 To objBill.Pages(i).Details.Count
            For k = 1 To objBill.Pages(i).Details(j).InComes.Count
                If Abs(dbl应收 + Val(objBill.Pages(i).Details(j).InComes(k).应收金额)) > 922337203685477# Then
                    MoneyOverFlow = True: Exit Function
                End If
                If Abs(dbl实收 + Val(objBill.Pages(i).Details(j).InComes(k).实收金额)) > 922337203685477# Then
                    MoneyOverFlow = True: Exit Function
                End If
                dbl应收 = dbl应收 + Val(objBill.Pages(i).Details(j).InComes(k).应收金额)
                dbl实收 = dbl实收 + Val(objBill.Pages(i).Details(j).InComes(k).实收金额)
            Next
        Next
    Next
End Function

Public Function AnalyseComputer() As String
    Dim strComputer As String * 256
    Call GetComputerName(strComputer, 255)
    AnalyseComputer = strComputer
    AnalyseComputer = Replace(AnalyseComputer, Chr(0), "")
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

Public Function CheckValid() As Boolean
    Dim intAtom As Integer
    Dim blnValid As Boolean
    Dim strSource As String
    Dim strCurrent As String
    Dim strBuffer As String * 256
    
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

Public Function CollectionExitsValue(ByVal coll As Collection, _
    ByVal strKey As String) As Boolean
    '根据关键字判断元素是否存在于集合中
    Dim blnExits As Boolean
    
    If coll Is Nothing Then Exit Function
    CollectionExitsValue = True
    Err = 0: On Error Resume Next
    blnExits = IsObject(coll(strKey))
    If Err <> 0 Then Err = 0: CollectionExitsValue = False
End Function

Public Function GetModuleType() As Byte
    '99993:李南春,2016/8/26,BH调用刷新问题
    If gfrmMain Is Nothing And glngMain = 0 Then
        GetModuleType = 0
    Else
        GetModuleType = 1
    End If
End Function
