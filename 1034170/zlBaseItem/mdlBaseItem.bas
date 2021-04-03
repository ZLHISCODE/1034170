Attribute VB_Name = "mdlBaseItem"
Option Explicit
Public gbln使用中医 As Boolean
Public gbln购买中医 As Boolean
Public gstr医价接口编号 As String
Public gbln允许医价收费项目 As Boolean
Public gbln从项汇总折扣  As Boolean
'外挂功能
Public gobjPlugIn As Object
Public gblnMyStyle As Boolean
Public gstrMatchMode As String
Public gbytCode As Byte
Public Enum gRegType
    g注册信息 = 0
    g公共全局 = 1
    g公共模块 = 2
    g私有全局 = 3
    g私有模块 = 4
End Enum
'-------------------------------------------------------------------------------------------------------------------------------------------------
'--定义系统参数
'问题:27990
Private Type Ty_System_Para
     byt药品名称显示 As Byte   '药品名称显示（主界面单据明细、单据输入界面、直接进入的药品选择器时的药品名称显示）：0-显示通用名，1-显示商品名，2-同时显示通用名和商品名
     byt输入药品显示 As Byte  '输入药品显示（通过输入简码方式进入选择器时药品名称的显示）：0-按输入匹配显示，1-固定显示通用名和商品名
End Type
Public gTy_System_Para As Ty_System_Para
Public gblnFeeKindCode As Boolean
'Windows风格----------------------------------
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

Public Const GCST_INVALIDCHAR = " '"    '对于输入的无效字符

Public gobjCustAcc As Object

Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Public Enum EditMode 'medit方式  取值为：0、新增；1、修改；2、调价；3、执行科室、4、从属项目、5、批量修改执行科室
    EditNew = 0
    EditModify = 1
    EditRaise = 2
    EditDept = 3
    EditSlave = 4
    EditCopy = 5
End Enum

Public gobjRIS As Object                    '新网RIS接口对象
Public Enum RISBaseItemOper                 '新网RIS基础数据操作类型：1-新增；2-修改；3-删除
    AddNew = 1
    Modify = 2
    Delete = 3
End Enum
Public Enum RISBaseItemType                 '新网RIS基础数据类型：3：用户(人员）
    Personnel = 3
End Enum
'本地日志模块
Private mobjFso As New FileSystemObject '文件对象
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'切换到指定的输入法。
Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long

'返回系统中可用的输入法个数及各输入法所在Layout,包括英文输入法。
Public Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long

'获取某个输入法的名称
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long

'判断某个输入法是否中文输入法
Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long

'''''''''''''''''''''
'获取指定输入法所在Layout,参数为0时表示当前输入法。
Public Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
'获取当前输入法所在Layout名
Public Declare Function GetKeyboardLayoutName Lib "user32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
'根据输入法Layout名将该输入法切换到输入法切换顺序的最前头(重新启动后无效),flags参数=KLF_REORDER
Public Declare Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal flags As Long) As Long

'下列语句用于检测是否合法调用
Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Sub SetFormVisible(ByVal new_Hwnd As Long)
'功能：隐藏窗体最大最小按钮
    SetWindowLong new_Hwnd, GWL_STYLE, GetWindowLong(new_Hwnd, GWL_STYLE) And Not &HCC0000 Or WS_SYSMENU Or &H20000
End Sub

Public Sub IniRIS(Optional ByVal blnMsg As Boolean)
'功能：初始化新网接口部件
'参数：blnMsg－创建失败时是否提示
    If gobjRIS Is Nothing Then
        On Error Resume Next
        Set gobjRIS = CreateObject("zl9XWInterface.clsHISInner")
        Err.Clear: On Error GoTo 0
    End If
    If gobjRIS Is Nothing Then
        If blnMsg Then
            MsgBox "RIS接口部件(zl9XWInterface)未创建成功！", vbInformation, gstrSysName
        End If
    End If
End Sub
Public Function MouseInRect(ByVal lngHwnd As Long) As Boolean
'功能：判断当前屏幕鼠标是否在指定窗口的显示区域内
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
        End If
    End If
    FormatEx = strNumber
End Function

Public Function MoveSpecialChar(ByVal strInputString As String) As String
    '1 去除一般字符: " '_%?"，把_%?转换为对应的全角字符
    '2 去除特殊字符:退格、制表、换行、回车
    Dim n As Integer
    Dim intStrLen As Integer
    Dim intASC As Integer
    Dim strText As String
    Dim strTmp As String
    Const CST_SPECIALCHAR = "_%?"      '允许转换的字符
    
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
                    strTmp = strTmp & "？"
                Case "%"
                    strTmp = strTmp & "％"
                Case "_"
                    strTmp = strTmp & "＿"
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
'功能：根据输入的字符串自动匹配ComboBox的选中项,并自动识别输入间隔
'参数：lngHwnd=ComboBox的Hwnd属性,KeyAscii=ComboBox的KeyPress事件中的KeyAscii参数,sngInterval=指定输入间隔
'返回：-2=未加处理,其它=匹配的索引(含不匹配的索引)
'说明：请将该函数在KeyPress事件中调用。

    Static lngPreTime As Single, lngPreHwnd As Long
    Static strFind As String
    
    Dim sngTime As Single, lngR As Long
    Dim lngIndex As Long
    
    If lngPreHwnd <> lngHwnd Then lngPreTime = Empty: strFind = Empty
    lngPreHwnd = lngHwnd
    
    If KeyAscii <> 13 Then
        '如果已经没有选中项，那么马上重新开始
        lngIndex = SendMessageLong(lngHwnd, CB_GETCURSEL, 0, 0)
        If lngIndex < 0 Then lngPreTime = 0
        
        sngTime = Timer
        If Abs(sngTime - lngPreTime) > sngInterval Then '输入间隔(缺省为0.5秒)
            strFind = ""
        End If
        If KeyAscii = vbKeyEscape Then
            lngPreTime = 0
        Else
            lngPreTime = Timer
        End If
        strFind = strFind & Chr(KeyAscii)
        
        KeyAscii = 0 '使ComboBox本身的单字匹配功能失效
        MatchIndex = SendMessage(lngHwnd, CB_FINDSTRING, -1, ByVal strFind)
        If MatchIndex = -1 Then Beep
    Else
        MatchIndex = -2 '在这里对回车不作处理
    End If
End Function

Public Sub 改变编码(nodParent As Node, int舍去长度 As Integer, str新增长度 As String)
'功能:改变树形列表各节点的标题中编码的值
'参数:nodParent         要改变编码的起始节点
'     int舍去长度       编码中舍去长度
'     str新增长度       编码中新增部分

    Dim nod As Node
    '它是下级也要改变编码
    If nodParent.Children > 0 Then
        Set nod = nodParent.Child
        Do While Not (nod Is Nothing)
            nod.Text = "【" & str新增长度 & Mid(nod.Text, int舍去长度 + 2)
            改变编码 nod, int舍去长度, str新增长度
            Set nod = nod.Next
        Loop
    End If
End Sub

Public Function GetRoot(ByVal nod As Node) As Node
'功能：读出任意节点的根节点
    Dim nodTemp As Node
    
    If nod Is Nothing Then Exit Function
    Set nodTemp = nod
    Do Until nodTemp.Parent Is Nothing
        Set nodTemp = nodTemp.Parent
    Loop
    Set GetRoot = nodTemp
End Function

Public Function GetTextFromCombo(cmbTemp As ComboBox, ByVal blnAfter As Boolean, Optional strSplit As String = "-") As String
'参数：cmbTemp  准备获取数据的ComboBox控件
'      blnAfter 表示在.之前或之后取值
    Dim lngPos As Long
    
    lngPos = InStr(cmbTemp.Text, strSplit)
    If lngPos = 0 Then
        '直接返回整个字符串
        GetTextFromCombo = "'" & cmbTemp.Text & "'"
    Else
        If blnAfter = False Then
            '圆点之前
            GetTextFromCombo = "'" & Mid(cmbTemp.Text, 1, lngPos - 1) & "'"
        Else
            GetTextFromCombo = "'" & Mid(cmbTemp.Text, lngPos + 1) & "'"
        End If
    End If
End Function

Public Sub SetComboByText(cmbTemp As ComboBox, ByVal strText As String, ByVal blnAfter As Boolean, Optional strSplit As String = "-")
'参数：cmbTemp  准备设置的ComboBox控件
'      blnAfter 表示在.之前或之后取值
    Dim lngPos As Long
    Dim lngCount As Long
    Dim strTemp As String
    Dim blnMatch As Boolean
    
    For lngCount = 0 To cmbTemp.ListCount - 1
        strTemp = cmbTemp.List(lngCount)
        
        lngPos = InStr(strTemp, strSplit)
        If lngPos = 0 Then
            '直接返回整个字符串
            If strText = cmbTemp.Text Then
                blnMatch = True
                Exit For
            End If
        Else
            If blnAfter = False Then
                '圆点之前
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
        '已经找到
        cmbTemp.ListIndex = lngCount
    Else
        cmbTemp.ListIndex = -1
        If blnAfter = True Then
            cmbTemp.AddItem strText
        End If
    End If
End Sub

Public Sub 调查报盘(frmParent As Form)
    MsgBox "请运行病案系统的人员管理。", vbInformation, gstrSysName
End Sub


Public Function GetPictureInfo(picTemp As StdPicture, Optional strBitmap As String = "") As String
'获得一张图片的信息
    Dim hFile As Integer
    Dim FileHeader As BITMAPFILEHEADER
    Dim InfoHeader As BITMAPINFOHEADER
    
    If picTemp.Handle = 0 Then
        GetPictureInfo = "无图片"
        Exit Function
    End If
    
    Dim strFile As String, strPath As String
    Dim intFileNum As Integer
    
    If strBitmap = "" Then
        '产生临时文件
        strPath = Space(256): strFile = Space(256)
        GetTempPath 256, strPath
        strPath = Left$(strPath, InStr(strPath, Chr(0)) - 1)
        
        GetTempFileName strPath, "pic", 0, strFile
        strFile = Left$(strFile, InStr(strFile, Chr(0)) - 1)
    
        SavePicture picTemp, strFile
    Else
        '直接使用现在文件
        strFile = strBitmap
    End If
    hFile = FreeFile
    Open strFile For Binary Access Read As #hFile
      Get #hFile, , FileHeader
      Get #hFile, , InfoHeader
    Close #hFile
    
    If strBitmap = "" Then
        '删除临时文件
        Kill strFile
    End If
    
    If InfoHeader.biBitCount > 8 Then
         GetPictureInfo = InfoHeader.biWidth & "×" & InfoHeader.biHeight & " " & InfoHeader.biBitCount & "位色"
    Else
         GetPictureInfo = InfoHeader.biWidth & "×" & InfoHeader.biHeight & " " & 2 ^ InfoHeader.biBitCount & "色"
    End If
End Function

Public Sub PressShiftTab(bytKey As Byte)
    '功能：向键盘发送一个键,类似SendKey,不过加了Shift
    '参数：bytKey=VirtualKey Codes，1-254，可以用vbKeyTab,vbKeyReturn,vbKeyF4

    Call keybd_event(vbKeyShift, 0, KEYEVENTF_EXTENDEDKEY, 0)
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY, 0)
    Call keybd_event(bytKey, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
    Call keybd_event(vbKeyShift, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
End Sub

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
    CheckValid = False
    
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

Public Function zlGetSymbol(strInput As String, Optional bytIsWB As Byte = 0, Optional intOutNum As Integer = 10) As String
    '----------------------------------
    '功能：生成字符串的简码
    '入参：strInput-输入字符串；bytIsWB-是否五笔(否则为拼音)
    '出参：正确返回字符串；错误返回"-"
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
'功能:按名称打开中文输入法,不指定名称时关闭中文输入法。支持部分名称。
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
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIF(IsNull(varValue), DefaultValue, varValue)
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

Public Function GetCoordPos(ByVal lngHwnd As Long, ByVal lngX As Long, ByVal lngY As Long) As POINTAPI
'功能：得控件中指定坐标在屏幕中的位置(Twip)
    Dim vPoint As POINTAPI
    vPoint.X = lngX / Screen.TwipsPerPixelX: vPoint.Y = lngY / Screen.TwipsPerPixelY
    Call ClientToScreen(lngHwnd, vPoint)
    vPoint.X = vPoint.X * Screen.TwipsPerPixelX: vPoint.Y = vPoint.Y * Screen.TwipsPerPixelY
    GetCoordPos = vPoint
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

Public Function CloneRecord(rsSource As ADODB.Recordset) As ADODB.Recordset
'功能：Clone产生一个本地记录集
'参数：rsSource=本地或数据库记录集
'说明：1.因为记录集本身的Clone功能对于记录增减是同步的
'      2.适用于小数据量记录集
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
'功能：获取指定控件在屏幕中的位置(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    zlGetControlRect = vRect
End Function
Public Sub InitSystemPara()
    '个人全局参数
    '-------------------------------------------------------------------------------------------------
    gbytCode = Val(zlDatabase.GetPara("简码方式"))
    '收费项目输入简码匹配方式:10.输入全是数字时只匹配编码  01.输入全是字母时只匹配简码,11两者均要求
    gstrMatchMode = zlDatabase.GetPara(44, glngSys, , "00")
    '当不输类别时,输入费用项目时,首位当作类别简码
    gblnFeeKindCode = zlDatabase.GetPara(144, glngSys) = "1"
    gstrLike = IIF(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "")
    gbln从项汇总折扣 = zlDatabase.GetPara(93, glngSys) = "1"
    '问题:27990
    With gTy_System_Para
        .byt输入药品显示 = Val(zlDatabase.GetPara("输入药品显示")) '0-按输入匹配显示，1-固定显示通用名和商品名
        .byt药品名称显示 = Val(zlDatabase.GetPara("药品名称显示"))  '：0-显示通用名，1-显示商品名，2-同时显示通用名和商品名
    End With
End Sub
Public Function GetFeeKind() As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select 编码, 名称, 简码 From 收费项目类别"
    Set GetFeeKind = zlDatabase.OpenSQLRecord(strSQL, "获取收费类别")
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
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
Public Function NeedName(strList As String) As String
    NeedName = Mid(strList, InStr(strList, "-") + 1)
End Function

Public Sub ShowMsgBox(ByVal strMsgInfor As String, Optional blnYesNo As Boolean = False, Optional ByRef blnYes As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:提示消息框
    '入参:strMsgInfor-提示信息
    '        blnYesNo-是否提供YES或NO按钮
    '出参:
    '返回:blnYes-如果提供YESNO按钮,则返回YES(True)或NO(False)
    '编制:刘兴洪
    '日期:2010-08-27 16:28:02
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
    '功能:  将指定的信息保存在注册表中
    '参数:  RegType-注册类型
    '       strSection-注册表目录
    '       StrKey-键名
    '       strKeyValue-键值
    '返回:
    '--------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    Select Case RegType
        Case g注册信息
            SaveSetting "ZLSOFT", "注册信息\" & strSection, strKey, strKeyValue
        Case g公共全局
            SaveSetting "ZLSOFT", "公共全局\" & strSection, strKey, strKeyValue
        Case g公共模块
            SaveSetting "ZLSOFT", "公共模块" & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
        Case g私有全局
            SaveSetting "ZLSOFT", "私有全局\" & gstrDbUser & "\" & strSection, strKey, strKeyValue
        Case g私有模块
            SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
    End Select
ErrHand:
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
            strKeyValue = GetSetting("ZLSOFT", "私有全局\" & gstrDbUser & "\" & strSection, strKey, "")
        Case g私有模块
            strKeyValue = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & strSection, strKey, "")
    End Select
ErrHand:
End Sub

Public Sub zlAddArray(ByRef cllData As Collection, ByVal strSQL As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:向指定的集合中插入数据
    '入参:cllData-指定的SQL集
    '     strSql-指定的SQL语句
    '编制:刘兴洪
    '日期:2010-08-30 15:51:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    i = cllData.Count + 1
    cllData.Add strSQL, "K" & i
End Sub
Public Sub zlExecuteProcedureArrAy(ByVal cllProcs As Variant, ByVal strCaption As String, Optional blnNoCommit As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行相关的Oracle过程集
    '入参:cllProcs-oracle过程集
    '     strCaption -执行过程的父窗口标题
    '     blnNOCommit-执行完过程后,不提交数据
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2010-08-30 15:52:25
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
'功能：由ItemData查找ComboBox的索引值
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
    '功能： 计算X,Y的实际坐标，并考虑屏幕超界的问题
    '参数： X---返回横坐标参数
    '       Y---返回纵坐标参数
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
'功能： 列出医疗小组医生的病人信息
'参数： lngTeamID: 医疗小组ID
'       lngMemberID: 医生ID
'返回： 病人信息字符串
'----------------------------------------------------------------------
    Dim strMess As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo errHandle
    gstrSQL = "Select a.病人id, a.住院号, a.出院病床, b.姓名" & vbNewLine & _
              "From 病案主页 a, 病人信息 b " & vbNewLine & _
              "Where a.住院医师 = (Select 姓名" & vbNewLine & _
              "              From 人员表" & vbNewLine & _
              "              Where ID = [2] And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null)) And" & vbNewLine & _
              "      a.医疗小组id = [1] and a.病人id=b.病人id and a.主页id=b.主页id and b.在院=1 "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "医疗小组医生病人信息", lngTeamID, lngMemberID)
    With rsTmp
        For i = 1 To .RecordCount
            strMess = strMess & "姓名：" & !姓名 & "；" & vbTab & _
                      "住院号：" & IIF(IsNull(!住院号), "", !住院号) & "；" & vbTab & _
                      "床号：" & IIF(IsNull(!出院病床), "", !出院病床) & vbTab & vbNewLine
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
'功能: 检查部门权限
'lngOperationID: 要操作的人员ID
'lngDeptID: 要操作人员的部门ID
'返回: True有权限, False无权限
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errHandle
    If lngDeptID = 0 Then
        gstrSQL = "Select Count(*) Rec From 部门人员 " & _
                  "Where 人员id = [2] And [3] In (Select 部门id From 部门人员 Where 人员id = [1])"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "检查人员的部门权限", glngUserId, lngOperationID, lngDeptID)
    Else
        gstrSQL = "Select ID " & _
                  "From 部门表 " & _
                  "  Start With ID In (Select 部门id From 部门人员 Where 人员id = [1]) " & _
                  "  Connect By Prior ID = 上级id"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "检查人员的部门权限", glngUserId)
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
    '写一行日志，如果内容中有回车,换行符，替换为<CR><LF>
    '日志保存在当前目录下的[应用程序名称]Log目录下，文件名为日期.txt,默认保存7天的日志。

    Dim strLogPath As String, strLogFile  As String, strLogIni As String    '日志路径，文件名，配置文件名
    Dim strLogSaveDays As String '日志保留天数
    Dim dblFreeSpace As Double   '剩余空间
    Dim strDelOldFile As String  '过期文件
    Dim objFile As File

    If Val(OS.IniRead("LOG", "OPENLOG", App.Path & "\CONFIG.INI")) = 0 Then Exit Sub
    '始终保存日志
    '2、清除过期日志
    strLogSaveDays = "7"  '保留7天的日志
    strLogPath = App.Path
    
    strDelOldFile = Dir(strLogPath & "\日志*.log")
    Do While strDelOldFile <> ""
        Set objFile = mobjFso.GetFile(strLogPath & "\" & strDelOldFile)
        If DateDiff("d", objFile.DateLastModified, Now) > Val(strLogSaveDays) Then
            mobjFso.DeleteFile strLogPath & "\" & strDelOldFile, True
        End If
        strDelOldFile = Dir
    Loop
    '3、空间是否足够
    dblFreeSpace = GetFreeSpace(strLogPath)
    If dblFreeSpace >= 1024 And dblFreeSpace <= 10240 Then
        '空间不足，不写日志,产生一个警告文件
        If Not mobjFso.FileExists(strLogPath & "\空间不足.txt") Then Call mobjFso.CreateTextFile(strLogPath & "\空间不足.txt", True)
        Exit Sub
    Else
        '清除警告文件
        If mobjFso.FileExists(strLogPath & "\空间不足.txt") Then Call mobjFso.DeleteFile(strLogPath & "\空间不足.txt", True)
    End If
    '4、写入日志行
    strLogFile = strLogPath & "\日志" & Format(Now, "yyyyMMdd") & ".log"

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
    '获取剩余空间
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

