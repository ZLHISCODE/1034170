Attribute VB_Name = "mdlCISBase"
Option Explicit
Public gcnOracle As ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例
Public gstrPrivs As String                   '当前用户具有的当前模块的功能
Public gstrProductName As String
Public gstrSysName As String                '系统名称
Public glngModul As Long
Public glngSys As Long
Public gblnCancel As Boolean                '记录界面中的取消按钮是否被点击了

Public gstrDBOwner As String                '当前系统所有者
Public gstrDBUser As String                 '当前数据库用户
Public glngUserId As Long                   '当前用户id
Public gstrUserCode As String               '当前用户编码
Public gstrUserName As String               '当前用户姓名
Public gstrUserAbbr As String               '当前用户简码

Public glngDeptId As Long                   '当前用户部门id
Public gstrDeptCode As String               '当前用户部门编码
Public gstrDeptName As String               '当前用户部门名称
Public gstrItemName As String

Public gstrUnitName As String               '用户单位名称
Public gfrmMain As Object


Public glngPreHWnd As Long '用于支持鼠标滚轮功能

Public gstrSql As String
Public gstrMatch As String                  '根据本地参数“匹配模式”确定的左匹配符号
Public gblnOK As Boolean

Public gobjKernel As New clsCISKernel       '临床核心部件
Public gobjLogisticPlatform As Object       '物流平台接口

Public gobjRIS As Object                    '新网RIS接口对象
Public Enum RISBaseItemOper                 '新网RIS基础数据操作类型：1-新增；2-修改；3-删除
    AddNew = 1
    Modify = 2
    Delete = 3
End Enum
Public Enum RISBaseItemType                 '新网RIS基础数据类型：1：诊疗项目目录，2：诊疗项目部位
    ClinicItem = 1
    ClinicItemPart = 2
End Enum

Public gblnKSSStrict As Boolean             '是否启用抗菌药物严格控制
Public gblnIncomeItem As Boolean            '记录收入项目是否设置

Public Type type_user_Digits
    dig_成本价 As Double
    dig_零售价 As Double
    dig_数量 As Double
    dig_金额 As Double
End Type
Public gtype_MaxDigits As type_user_Digits  '用来记录最大精度

Public Type TYPE_USER_INFO
    ID As Long
    部门ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
    用药级别 As Long
End Type
Public UserInfo As TYPE_USER_INFO
Public Const gstrLisHelp As String = "zl9LisWork"               'LIS调用帮助时使用的部件名
Public glngTXTProc As Long '保存默认的消息函数的地址
Public Const WM_CONTEXTMENU = &H7B ' 当右击文本框时，产生这条消息
Public Const GCST_INVALIDCHAR = "'"             '对于输入的无效字符

'支持滑轮的常量
Public Const WM_MOUSEWHEEL = &H20A


Public Const GWL_STYLE = (-16)
Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetClientRect Lib "User32" (ByVal hWnd As Long, lpRect As RECT) As Long

'私有、公共模块参数
Public Enum 参数_药品目录管理_公共
    P1_西成药收入项目 = 1
    P2_中成药收入项目 = 2
    P3_中草药收入项目 = 3
    P4_应用范围 = 4
    P5_时价药品按批次调价 = 5
End Enum

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

Public Function GetFormat(ByVal dblInput As Double, ByVal intDotBit As Integer) As String
    '取数值的小数位数
    GetFormat = Format(dblInput, "#0." & String(intDotBit, "0"))
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

Public Sub GetMaxDigit()
    '用来取药品的各种最大精度
    Dim rsTemp As ADODB.Recordset
    On Error GoTo ErrHand
    
    gstrSql = "Select 零售金额, 成本价, 零售价, 实际数量 From 药品收发记录 Where Rownum < 1"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "最大精度")
    If rsTemp.RecordCount = 0 Then
        gtype_MaxDigits.dig_成本价 = 7
        gtype_MaxDigits.dig_金额 = 2
        gtype_MaxDigits.dig_零售价 = 7
        gtype_MaxDigits.dig_数量 = 7
    Else
        gtype_MaxDigits.dig_成本价 = rsTemp.Fields(1).NumericScale
        gtype_MaxDigits.dig_金额 = rsTemp.Fields(0).NumericScale
        gtype_MaxDigits.dig_零售价 = rsTemp.Fields(2).NumericScale
        gtype_MaxDigits.dig_数量 = rsTemp.Fields(3).NumericScale
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

'取药品金额、价格和数量的小数位数
Public Function GetDigit(ByVal int类别 As Integer, ByVal int内容 As Integer, Optional ByVal int单位 As Integer) As Integer
    'int类别：1-药品;2-卫材
    'int内容：1-成本价;2-零售价;3-数量;4-金额
    'int单位：如果是取金额位数，可以不输入该参数
    '         药品单位:1-售价;2-门诊;3-住院;4-药库;
    '         卫材单位:1-散装;2-包装
    '返回：最小2，最大为数据库最大小数位数
    
    Dim rsTmp As ADODB.Recordset
    Dim intMax金额 As Integer
    Dim intMax成本价 As Integer
    Dim intMax零售价 As Integer
    Dim intMax数量 As Integer
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    gstrSql = "Select 零售金额, 成本价, 零售价, 实际数量 From 药品收发记录 Where Rownum < 1"
    Set rs = zlDatabase.OpenSQLRecord(gstrSql, "取药品精度")
    
    intMax金额 = rs.Fields(0).NumericScale
    intMax成本价 = rs.Fields(1).NumericScale
    intMax零售价 = rs.Fields(2).NumericScale
    intMax数量 = rs.Fields(3).NumericScale
    
    gstrSql = "Select Nvl(精度, 0) 精度 From 药品卫材精度 Where 类别 = [1] And 内容 = [2] And 单位 = [3] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, "取药品" & Choose(int内容, "成本价", "零售价", "数量") & "小数位数", int类别, int内容, int单位)
    
    If rsTmp.RecordCount > 0 Then
        GetDigit = rsTmp!精度
    End If
    
    If GetDigit = 0 Then
        '如果没有设置精度，则取数据库允许的最大位数
        GetDigit = Choose(int内容, intMax成本价, intMax零售价, intMax数量)
    End If
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    GetDigit = Choose(int内容, intMax成本价, intMax零售价, intMax数量, intMax金额)
End Function


Public Function GetUserInfo() As Boolean
    '功能：获取登陆用户信息
    Dim rsTmp As New ADODB.Recordset
    
    Set rsTmp = zlDatabase.GetUserInfo
    
    UserInfo.用户名 = gstrDBUser
    UserInfo.姓名 = gstrDBUser
    If Not rsTmp.EOF Then
        UserInfo.ID = rsTmp!ID
        UserInfo.编号 = rsTmp!编号
        UserInfo.部门ID = IIf(IsNull(rsTmp!部门ID), 0, rsTmp!部门ID)
        UserInfo.简码 = IIf(IsNull(rsTmp!简码), "", rsTmp!简码)
        UserInfo.姓名 = IIf(IsNull(rsTmp!姓名), "", rsTmp!姓名)
        gstrUserName = UserInfo.姓名
        GetUserInfo = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function MoveSpecialChar(ByVal strInputString As String, Optional ByVal blnMoveSpace As Boolean = True) As String
    '1 去除一般字符: " '_%?"，把_%?转换为对应的全角字符
    '2 去除特殊字符:退格、制表、换行、回车
    '3 blnMoveSpace，是否去掉字符中的空格，Ture-去掉空格；注意头尾空格默认去掉
    Dim n As Integer
    Dim intStrLen As Integer
    Dim intAsc As Integer
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
        intAsc = Asc(Mid(strText, n, 1))
        Select Case intAsc
            Case 8, 9, 10, 13
            Case 32
                '空格处理
                If blnMoveSpace = False Then
                    strTmp = strTmp & Mid(strText, n, 1)
                End If
            Case Else
                strTmp = strTmp & Mid(strText, n, 1)
        End Select
    Next
    
    MoveSpecialChar = strTmp
    
End Function

Public Function zlGetSymbol(strInput As String, Optional bytIsWB As Byte, Optional intOutNum As Integer = 10) As String
    '----------------------------------
    '功能：生成字符串的简码
    '入参：strInput-输入字符串；bytIsWB-是否五笔(否则为拼音)
    '出参：正确返回字符串；错误返回"-"
    '----------------------------------
    Dim strReturn As String
    
    strReturn = zlCommFun.zlGetSymbol(strInput, bytIsWB)
    
    zlGetSymbol = Mid(strReturn, 1, intOutNum)
End Function

Public Function zlClinicCodeRepeat(strInputCode As String, Optional lngSelfID As Long) As Boolean
    '----------------------------------
    '功能：检查诊疗项目编码的是否与现有编码重复，重复则给出提示
    '入参：strInputCode-输入的编码；lngSelfID-自己的ID号，当修改时，需要将自身除开才能判断
    '出参：重复返回True；否则反馈Flase
    '----------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    strSql = "select K.名称||' ['||I.编码||']'||I.名称 as 名称" & _
            " from 诊疗项目目录 I,诊疗项目类别 K" & _
            " where I.类别=K.编码 and I.编码=[1] " & _
            "       and I.ID<>[2]"
    Err = 0: On Error GoTo ErrHand
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISBase", strInputCode, lngSelfID)
        
    With rsTmp
        If .RecordCount <> 0 Then
            MsgBox "该项目与“" & !名称 & "”编码重复！", vbExclamation, gstrSysName
            zlClinicCodeRepeat = True
        Else
            zlClinicCodeRepeat = False
        End If
    End With
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlClinicCodeRepeat = True
End Function

Public Function zlExseCodeRepeat(strInputCode As String, Optional lngSelfID As Long) As Boolean
    '----------------------------------
    '功能：检查收费项目编码的是否与现有编码重复，重复则给出提示
    '入参：strInputCode-输入的编码；lngSelfID-自己的ID号，当修改时，需要将自身除开才能判断
    '出参：重复返回True；否则反馈Flase
    '----------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    strSql = "select K.名称||' ['||I.编码||']'||I.名称 as 名称" & _
            " from 收费项目目录 I,收费项目类别 K" & _
            " where I.类别=K.编码 and I.编码=[1] " & _
            "       and I.ID<>[2]"
    Err = 0: On Error GoTo ErrHand
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISBase", strInputCode, lngSelfID)
    
    With rsTmp
        If .RecordCount <> 0 Then
            MsgBox "该项目与“" & !名称 & "”编码重复！", vbExclamation, gstrSysName
            zlExseCodeRepeat = True
        Else
            zlExseCodeRepeat = False
        End If
    End With
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlExseCodeRepeat = True
End Function


Public Function zlExistItem(ByVal strTbleName As String, ByVal strField As String, ByVal varValues As Variant, _
                            ByVal strItemName As String) As Boolean
    
    '----------------------------------
    '功能：检查项目是否存在,用于并发操作时的检查
    '入参：strTableName 表名 ,strField 字段名 , ,lngItemID,字段的值,strItemName 提示时显示的项目名称
    '出参：存在返回True；否则反馈Flase
    '----------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    Err = 0: On Error GoTo ErrHand
    strSql = "Select " & strField & " From " & strTbleName & " Where " & strField & "=[1]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISBase", varValues)
    If rsTmp.RecordCount > 0 Then
        zlExistItem = True
    Else
         MsgBox "“" & strItemName & "”已经被其他操作员删除！", vbExclamation, gstrSysName
        zlExistItem = False
    End If
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlExistItem = False
End Function

Public Function CheckIsInclude(strSource As String, strTarge As String) As Boolean
'检查strSource中的每一个字符是否在strTarge中
    Dim i As Long
    CheckIsInclude = False
    
    Select Case strTarge
    Case "日期"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "时间"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"";|\=+-_)(*&^%$#@!`~"
    Case "日期时间"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"";|\=+_)(*&^%$#@!`~"
    Case "整数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "小数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "正整数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+-_)(*&^%$#@!`~"
    Case "正小数"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+-_)(*&^%$#@!`~"
    Case "可打印字符"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/.'"":;|\=+-_)(*&^%$#@!`~0123456789"
    End Select
    For i = 1 To Len(strSource)
        If InStr(strTarge, Mid(strSource, i, 1)) <= 0 Then Exit Function
    Next
    CheckIsInclude = True
End Function

Public Function Between(x, a, b) As Boolean
'功能：判断x是否在a和b之间
    If a < b Then
        Between = x >= a And x <= b
    Else
        Between = x >= b And x <= a
    End If
End Function

Public Function FormatEx(ByVal vNumber As Variant, ByVal intBit As Integer) As String
'功能：四舍五入方式格式化显示数字,保证小数点最后不出现0,小数点前要有0
'参数：vNumber=Single,Double,Currency类型的数字,intBit=最大小数位数
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

Public Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'功能：相当于Oracle的NVL，将Null值改成另外一个预设值
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function ZVal(ByVal varValue As Variant) As String
'功能：将0零转换为"NULL"串,在生成SQL语句时用
    ZVal = IIf(Val(varValue) = 0, "NULL", Val(varValue))
End Function

Public Function OpenRecord(rsTmp As ADODB.Recordset, strSql As String, ByVal strTitle As String, _
    Optional CursorType As CursorTypeEnum = adOpenKeyset, Optional LockType As LockTypeEnum = adLockReadOnly) As ADODB.Recordset
    
    On Error GoTo ErrHandle
'    If rsTmp.State = 1 Then rsTmp.Close
'    rsTmp.CursorLocation = adUseClient
'    Call SQLTest(App.ProductName, strTitle, strSql)
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "OpenRecord")
'    Call SQLTest
    Set OpenRecord = rsTmp
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub ClearSpecRowCol(obj As Object, ByVal intRow As Integer, Optional intCol As Variant)
'功能: 清除指定网格的指定行指定列的数据
'参数: obj=要操作的网格控件
'      intRow=要清除的行号
'      intCol=要清除的列号列表如Array(1,2,3),若所有列则可以表示为Array()
    Dim i As Long
    If UBound(intCol) = -1 Then
        For i = 0 To obj.Cols - 1
            obj.TextMatrix(intRow, i) = ""
        Next
    Else
        For i = 0 To UBound(intCol)
            obj.TextMatrix(intRow, intCol(i)) = ""
        Next
    End If
    obj.RowData(intRow) = 0
End Sub

Public Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0) As Boolean
'检查字符串是否含有非法字符；如果提供长度，对长度的合法性也作检测。
    If InStr(strInput, "'") > 0 Then
        MsgBox "所输入内容含有非法字符。", vbExclamation, gstrSysName
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox "所输入内容不能超过" & Int(intMax / 2) & "个汉字" & "或" & intMax & "个字符！", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    StrIsValid = True
End Function

Public Function AppendFields(rsTmp As ADODB.Recordset, varField As Variant, varType As Variant, varLength As Variant) As ADODB.Recordset
    Dim i As Long
    For i = 0 To UBound(varField)
        rsTmp.Fields.Append varField(i), varType(i), varLength(i)
    Next
End Function

Public Sub OpenRecordset(rsTemp As ADODB.Recordset, ByVal strCaption As String, Optional strSql As String = "")
'功能：打开记录集
'    If rsTemp.State = adStateOpen Then rsTemp.Close
    
    On Error GoTo ErrHandle
'    Call SQLTest(App.ProductName, strCaption, IIf(strSql = "", gstrSql, strSql))
    Set rsTemp = zlDatabase.OpenSQLRecord(IIf(strSql = "", gstrSql, strSql), "cmd产地_Click")
'    Call SQLTest
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'----------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------------
Public Sub NewColumn(msf As Object, ByVal vText As String, Optional ByVal vWidth As Single = 1200, Optional ByVal vAlignment As Byte = 9, Optional ByVal vFormat As String, Optional ByVal vEditMask As String)
    Dim i As Long
    
    msf.Cols = msf.Cols + 1
    i = msf.Cols - 1
    
    msf.TextMatrix(0, i) = vText
    msf.ColWidth(i) = vWidth
    msf.ColAlignment(i) = vAlignment
    
    
    On Error Resume Next
    
    msf.ColFormat(i) = vFormat
    msf.ColEditMask(i) = vEditMask
        
    msf.ColAlignmentFixed(i) = vAlignment
    
End Sub

Public Sub CalcPosition(ByRef x As Single, ByRef y As Single, ByVal objBill As Object)
    '----------------------------------------------------------------------
    '功能： 计算X,Y的实际坐标，并考虑屏幕超界的问题
    '参数： X---返回横坐标参数
    '       Y---返回纵坐标参数
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.hWnd, objPoint)
    
    x = objPoint.x * 15 + objBill.CellLeft
    y = objPoint.y * 15 + objBill.CellTop + objBill.CellHeight
End Sub

Public Function FillGrid(ByRef objMsf As Object, ByVal rsData As ADODB.Recordset, Optional ByVal MaskArray As Variant, Optional ByVal blnClear As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------
    '功能:填充数据到网格
    '参数:
    '返回:
    '---------------------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim strMask As String
    Dim lngRow As Long
    
    If blnClear Then
        objMsf.Rows = 2
        objMsf.RowData(1) = 0
        For lngLoop = 0 To objMsf.Cols - 1
            objMsf.TextMatrix(1, lngLoop) = ""
        Next
    End If
    
    lngRow = 0
    Do While Not rsData.EOF
        
        lngRow = lngRow + 1
        If objMsf.Rows < lngRow + 1 Then objMsf.Rows = lngRow + 1
        
        On Error Resume Next
        objMsf.RowData(lngRow) = CStr(zlCommFun.Nvl(rsData("ID")))
        
        On Error GoTo ErrHand
        For lngLoop = 0 To objMsf.Cols - 1
        
            On Error Resume Next
            strMask = ""
            strMask = MaskArray(lngLoop)
                                    
            On Error GoTo ErrHand
            If strMask <> "" Then
                objMsf.TextMatrix(lngRow, lngLoop) = Format(zlCommFun.Nvl(rsData(objMsf.TextMatrix(0, lngLoop))), strMask)
            Else
                objMsf.TextMatrix(lngRow, lngLoop) = zlCommFun.Nvl(rsData(objMsf.TextMatrix(0, lngLoop)))
            End If
            
        Next
        
        rsData.MoveNext
    Loop
    
    FillGrid = True
    
    Exit Function
    
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Public Function FillListData(ByRef objLvw As Object, ByVal rs As ADODB.Recordset) As Boolean
    '-------------------------------------------------------------------------------------------------------------
    '功能:
    '参数:
    '返回:
    '-------------------------------------------------------------------------------------------------------------
    Dim objItem As ListItem
    Dim lngLoop As Long
    
    On Error GoTo ErrHand
    
    LockWindowUpdate objLvw.hWnd
    
    Do While Not rs.EOF
        Set objItem = objLvw.ListItems.Add(, "K" & rs("ID").Value, rs("名称").Value, _
                      IIf(rs("项目类别") = "微生物" Or rs("组合") = "√", "ItemGroup", "Item"), _
                      IIf(rs("项目类别") = "微生物" Or rs("组合") = "√", "ItemGroup", "Item"))
                      
        For lngLoop = 2 To objLvw.ColumnHeaders.Count
            objItem.SubItems(lngLoop - 1) = zlCommFun.Nvl(rs(objLvw.ColumnHeaders(lngLoop).Text).Value)
        Next
                        
        rs.MoveNext
    Loop
    
    LockWindowUpdate 0
    
    FillListData = True
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function


Public Sub NextLvwPos(lvwObj As Object, ByVal vIndex As Long)
        
    If lvwObj.ListItems.Count > 0 Then
        vIndex = IIf(lvwObj.ListItems.Count > vIndex, vIndex, lvwObj.ListItems.Count)
        lvwObj.ListItems(vIndex).Selected = True
        lvwObj.ListItems(vIndex).EnsureVisible
    End If
End Sub

Public Function FilterKeyAscii(ByVal KeyAscii As Long, ByVal bytMode As Byte, Optional ByVal KeyCustom As String) As Long
            
    FilterKeyAscii = KeyAscii
    
    If KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or KeyAscii = vbKeyBack Then
        Exit Function
    End If
    
    Select Case bytMode
    Case 1      '纯数字
        If InStr("0123456789", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 99
        If InStr(KeyCustom, Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    End Select
    
End Function

Public Function CheckStrType(ByVal Text As String, ByVal bytMode As Byte, Optional ByVal KeyCustom As String) As Boolean
    Dim lngLoop As Long
    
    Select Case bytMode
    Case 1
        If Trim(Text) <> "" Then
            If InStr(Text, ".") = 0 And InStr(Text, "-") = 0 Then
                If IsNumeric(Text) Then
                    CheckStrType = True
                End If
            End If
        End If
    Case 99
        For lngLoop = 1 To Len(Text)
            If InStr(KeyCustom, Mid(Text, lngLoop, 1)) = 0 Then
                CheckStrType = False
                Exit Function
            End If
        Next
        CheckStrType = True
    End Select
End Function

Public Function GetMaxLength(ByVal strTable As String, ByVal strField As String, Optional ByVal bytMode As Byte = 1) As Long
    
    Dim rs As New ADODB.Recordset
    
    On Error Resume Next
    
    zlDatabase.OpenRecordset rs, "SELECT " & strField & " FROM " & strTable & " WHERE ROWNUM<1", "mdlCISBase"
    If bytMode = 1 Then
        GetMaxLength = rs.Fields(0).DefinedSize
    Else
        GetMaxLength = rs.Fields(0).NumericScale
    End If
    
End Function

Public Sub AddComboData(objSource As Object, ByVal rsTemp1 As ADODB.Recordset, Optional ByVal blnClear As Boolean = True)
'功能: 装载数据入指定的组合下拉框或网格中的下拉框中
    If blnClear = True Then objSource.Clear
    
    If rsTemp1.BOF = False Then
        rsTemp1.MoveFirst
        While Not rsTemp1.EOF
            objSource.AddItem rsTemp1.Fields(0).Value
            objSource.ItemData(objSource.NewIndex) = Val(rsTemp1.Fields(1).Value)
            rsTemp1.MoveNext
        Wend
        rsTemp1.MoveFirst
    End If
End Sub

Public Sub LocationObj(ByRef objTxt As Object)
    On Error Resume Next
    
    zlControl.TxtSelAll objTxt
    objTxt.SetFocus
End Sub

Public Sub LocationVsf(objVsf As Object, ByVal lngRow As Long, ByVal lngCol As Long)
    
    On Error Resume Next
    
    objVsf.Row = lngRow
    objVsf.Col = lngCol
    objVsf.ShowCell objVsf.Row, objVsf.Col
    objVsf.SetFocus
End Sub

Public Sub ShowSimpleMsg(ByVal strInfo As String)
    '------------------------------------------------------------------------------------------------------
    '功能：
    '------------------------------------------------------------------------------------------------------
    MsgBox strInfo, vbInformation, gstrSysName
    
End Sub

Public Sub ClearGrid(vsf As Object, Optional ByVal Row As Long = 1)
    '--------------------------------------------------------------------------------------------------------
    '功能:清除表格数据
    '--------------------------------------------------------------------------------------------------------
    vsf.Rows = Row + 1
    vsf.RowData(Row) = 0
    vsf.Cell(flexcpText, Row, 0, Row, vsf.Cols - 1) = ""
    
End Sub

Public Function CheckNumeric(ByVal strText As String, ByVal lngLength As Long, Optional ByVal lngDecLength As Long = 0, Optional ByVal bytMode As Byte = 1) As String
    '--------------------------------------------------------------------------------------------------------
    '功能:检测字符串的数值有效性
    '--------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    
    Dim str整数部份 As String
    Dim str小数部份 As String
    
    If lngDecLength = 0 Then
        '整数
        Select Case bytMode
        Case 1      '正整数
            str整数部份 = strText
        Case 2      '负整数
            If Left(strText, 1) <> "-" And strText <> "0" Then
                CheckNumeric = "应为负数或者零！"
                Exit Function
            End If
            str整数部份 = Mid(strText, 2)
            
        Case 3      '正负整数
            If Left(strText, 1) = "-" Then str整数部份 = Mid(strText, 2)
        End Select
    Else
        '小数
        Select Case bytMode
        Case 1      '正小数
            If Len(strText) > lngLength + 1 Then
                CheckNumeric = "长度超过了" & lngLength & "位！"
                Exit Function
            End If
            
            If InStr(strText, ".") > 0 Then
                '有小数部份
                str整数部份 = Left(strText, InStr(strText, ".") - 1)
                str小数部份 = Mid(strText, InStr(strText, ".") + 1)
            Else
                str整数部份 = strText
            End If
            
        Case 2      '负小数
            If Len(strText) > lngLength + 2 Then
                CheckNumeric = "长度超过了" & lngLength & "位！"
                Exit Function
            End If
            
            If Left(strText, 1) <> "-" Then
                CheckNumeric = "不是负数！"
                Exit Function
            End If
            
            If InStr(strText, ".") > 0 Then
                '有小数部份
                str整数部份 = Mid(strText, 2, InStr(strText, ".") - 2)
                str小数部份 = Mid(strText, InStr(strText, ".") + 1)
            Else
                str整数部份 = Mid(strText, 2)
            End If
            
        Case 3      '正负小数
            If Left(strText, 1) = "-" Then
                If Len(strText) > lngLength + 2 Then
                    CheckNumeric = "长度超过了" & lngLength & "位！"
                    Exit Function
                End If
                If InStr(strText, ".") > 0 Then
                    '有小数部份
                    str整数部份 = Mid(strText, 2, InStr(strText, ".") - 2)
                    str小数部份 = Mid(strText, InStr(strText, ".") + 1)
                Else
                    str整数部份 = Mid(strText, 2)
                End If
            Else
                If Len(strText) > lngLength + 1 Then
                    CheckNumeric = "长度超过了" & lngLength & "位！"
                    Exit Function
                End If
                If InStr(strText, ".") > 0 Then
                    '有小数部份
                    str整数部份 = Mid(strText, 1, InStr(strText, ".") - 1)
                    str小数部份 = Mid(strText, InStr(strText, ".") + 1)
                Else
                    str整数部份 = strText
                End If
                
            End If
        End Select
    End If
    
    If Len(str整数部份) > (lngLength - lngDecLength) Then
        If lngDecLength = 0 Then
            CheckNumeric = "长度超过了" & (lngLength - lngDecLength) & "位！"
        Else
            CheckNumeric = "整数部份长度超过了" & (lngLength - lngDecLength) & "位！"
        End If
        Exit Function
    End If
    
    If Len(str小数部份) > lngDecLength Then
        CheckNumeric = "小数部份长度超过了" & lngDecLength & "位！"
        Exit Function
    End If
    
    For lngLoop = 1 To Len(str整数部份)
        If Mid(str整数部份, lngLoop, 1) < "0" Or Mid(str整数部份, lngLoop, 1) > "9" Then
            CheckNumeric = "应为数字型！"
            Exit Function
        End If
    Next
    
    For lngLoop = 1 To Len(str小数部份)
        If Mid(str小数部份, lngLoop, 1) < "0" Or Mid(str小数部份, lngLoop, 1) > "9" Then
            CheckNumeric = "应为数字型！"
            Exit Function
        End If
    Next
    
    
    CheckNumeric = ""
End Function

'去掉TextBox的默认右键菜单
Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal Msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' 如果消息不是WM_CONTEXTMENU，就调用默认的窗口函数处理
    If Msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(glngTXTProc, hWnd, Msg, wp, lp)
End Function

Public Function GetSysPara(ByVal int序号 As Integer) As String
    Dim rsTemp As New ADODB.Recordset
    '获取系统参数
    On Error GoTo ErrHandle
    gstrSql = "Select Nvl(参数值,缺省值) From Zlparameters Where 系统 = [1] And Nvl(私有, 0) = 0 And 模块 Is Null And 参数号=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "获取系统参数值", glngSys, int序号)
    
    If rsTemp.RecordCount <> 0 Then
        GetSysPara = rsTemp.Fields(0).Value
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

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

Private Function PreFixNO(Optional curDate As Date = #1/1/1900#) As String
'功能：返回大写的单据号年前缀
    If curDate = #1/1/1900# Then
        PreFixNO = CStr(CInt(Format(zlDatabase.Currentdate, "YYYY")) - 1990)
    Else
        PreFixNO = CStr(CInt(Format(curDate, "YYYY")) - 1990)
    End If
    PreFixNO = IIf(CInt(PreFixNO) < 10, PreFixNO, Chr(55 + CInt(PreFixNO)))
End Function

Public Function GetFullNO(ByVal strNo As String, ByVal intNum As Integer) As String
'功能：由用户输入的部份单号，返回全部的单号。
'参数：intNum=项目序号,为0时固定按年产生
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, intType As Integer
    Dim curDate As Date
    
    On Error GoTo errH
    If Len(strNo) >= 8 Then
        GetFullNO = Right(strNo, 8)
        Exit Function
    ElseIf Len(strNo) = 7 Then
        GetFullNO = PreFixNO & strNo
        Exit Function
    ElseIf intNum = 0 Then
        GetFullNO = PreFixNO & Format(Right(strNo, 7), "0000000")
        Exit Function
    End If
    GetFullNO = strNo
    
    strSql = "Select 编号规则,Sysdate as 日期 From 号码控制表 Where 项目序号=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, App.ProductName, intNum)
    If Not rsTmp.EOF Then
        intType = Val("" & rsTmp!编号规则)
        curDate = rsTmp!日期
    End If

    If intType = 1 Then
        '按日编号
        strSql = Format(CDate(Format(rsTmp!日期, "YYYY-MM-dd")) - CDate(Format(rsTmp!日期, "YYYY") & "-01-01") + 1, "000")
        GetFullNO = PreFixNO & strSql & Format(Right(strNo, 4), "0000")
    Else
        '按年编号
        GetFullNO = PreFixNO & Format(Right(strNo, 7), "0000000")
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'################################################################################################################
'## 功能：  将数据从一个XtremeReportControl控件复制到VSFlexGrid，以便进行打印
'################################################################################################################
Public Function zlReportToVSFlexGrid(vfgList As VSFlexGrid, rptList As ReportControl) As Boolean
    '-------------------------------------------------
    '将全部组强制展开,复制数据表格
    Dim rptCol As ReportColumn
    Dim rptRcd As ReportRecord
    Dim rptItem As ReportRecordItem
    Dim rptRow As ReportRow
    
    Dim lngCol As Long, lngRow As Long
    
    On Error GoTo ErrHand
    For Each rptRow In rptList.Rows
        If rptRow.Childs.Count > 0 Then rptRow.Expanded = True
    Next
    If rptList.Rows.Count < 1 Then zlReportToVSFlexGrid = False: Exit Function
        
    With vfgList
        .Clear
        .Rows = 1: .FixedRows = 1: .RowHeight(.Rows - 1) = 280
        .Cols = 0
        .MergeCells = flexMergeFree
        
        '标题行复制
        For Each rptCol In rptList.Columns
            If rptCol.Visible Then
                .Cols = .Cols + 1
                .TextMatrix(0, .Cols - 1) = rptCol.Caption
                .ColData(.Cols - 1) = rptCol.ItemIndex
                Select Case rptCol.Alignment
                Case xtpAlignmentLeft: .ColAlignment(.Cols - 1) = flexAlignLeftCenter
                Case xtpAlignmentCenter: .ColAlignment(.Cols - 1) = flexAlignCenterCenter
                Case xtpAlignmentRight: .ColAlignment(.Cols - 1) = flexAlignRightCenter
                End Select
                .Cell(flexcpAlignment, 0, .Cols - 1, .FixedRows - 1) = flexAlignCenterCenter
                If rptCol.Width < 20 * IIf(rptList.GroupsOrder.Count = 0, 1, rptList.GroupsOrder.Count) Then
                    .ColWidth(.Cols - 1) = 0
                Else
                    .ColWidth(.Cols - 1) = rptCol.Width * Screen.TwipsPerPixelX
                End If
            End If
        Next
        
        '数据行复制
        Dim intTiers As Integer, rptParent As ReportRow, rptChild As ReportRow
        For Each rptRow In rptList.Rows
            .Rows = .Rows + 1: .RowHeight(.Rows - 1) = 280
            If rptRow.GroupRow Then
                intTiers = 0
                Set rptParent = rptRow
                Do While Not (rptParent.ParentRow Is Nothing)
                    intTiers = intTiers + 1
                    Set rptParent = rptParent.ParentRow
                Loop
                Set rptChild = rptRow.Childs(0)
                Do While rptChild.GroupRow
                    Set rptChild = rptChild.Childs(0)
                Loop
                .MergeRow(.Rows - 1) = True
                For lngCol = 0 To .Cols - 1
                    .TextMatrix(.Rows - 1, lngCol) = String(intTiers, "　") & rptList.GroupsOrder(intTiers).Caption & ": "
                    .TextMatrix(.Rows - 1, lngCol) = .TextMatrix(.Rows - 1, lngCol) & rptChild.Record(rptList.GroupsOrder(intTiers).ItemIndex).Value
                Next
            Else
                For lngCol = 0 To .Cols - 1
                    If rptList.Columns(.ColData(lngCol)).TreeColumn Then
                        intTiers = 0
                        Set rptParent = rptRow
                        Do While Not (rptParent.ParentRow Is Nothing)
                            intTiers = intTiers + 1
                            Set rptParent = rptParent.ParentRow
                        Loop
                        .TextMatrix(.Rows - 1, lngCol) = String(intTiers, "　") & rptRow.Record(.ColData(lngCol)).Value
                    Else
                        .TextMatrix(.Rows - 1, lngCol) = rptRow.Record(.ColData(lngCol)).Value
                    End If
                    .Cell(flexcpAlignment, .Rows - 1, lngCol, .Rows - 1) = .ColAlignment(lngCol)
                Next
            End If
        Next
    End With
    zlReportToVSFlexGrid = True
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    zlReportToVSFlexGrid = False
End Function

Public Function DelInvalidChar(ByVal strchar As String, Optional ByVal strInvalidChar As String) As String
    '删除非法字符
    'strChar: 要处理的字符
    'strInvalidChar：非法字符串，如果为空，则为~!@#$%^&*()_+|=-`;'"":/.,<>?{}[]\<>,否则按传入的字符处理
    Dim strBit As String, i As Integer, strWord As String
    strWord = "~!@#$%^&*()_+|=-`;'"":/.,<>?{}[]\<>"
    If strInvalidChar <> "" Then strWord = strInvalidChar
    If Len(strchar) > 0 Then
        For i = 1 To Len(strchar)
            strBit = Mid$(strchar, i, 1)
            If InStr(strWord, strBit) <= 0 Then
                DelInvalidChar = DelInvalidChar & strBit
            End If
        Next
    End If
End Function

Public Function CheckKSSPrivilege() As Boolean
'功能：检查系统是否存在抗菌药物授权的人员，并且设置当前操作员的用药级别到UserInfo对象
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    UserInfo.用药级别 = 0
    
    On Error GoTo errH
    strSql = "Select 级别 From 人员抗菌药物权限 Where 记录状态=1 and 人员ID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISKernel", UserInfo.ID)
    If rsTmp.RecordCount > 0 Then
        UserInfo.用药级别 = Val("" & rsTmp!级别)
        CheckKSSPrivilege = True
    Else
        strSql = "Select 1 From 人员抗菌药物权限 Where 记录状态=1 and Rownum<2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISKernel")
        CheckKSSPrivilege = rsTmp.RecordCount > 0
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function IntEx(vNumber As Variant) As Variant
'功能：取大于指定数值的最小整数
    IntEx = -1 * Int(-1 * Val(vNumber))
End Function



Public Function FmgFlexScroll(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'支持frmDoctorManage窗体滚轮的滚动
    On Error GoTo errH
    Select Case wMsg
    Case WM_MOUSEWHEEL
        Select Case wParam
            Case -7864320  '向下滚
                If frmDoctorManage.vscBar.Value <> frmDoctorManage.vscBar.Max Then
                    frmDoctorManage.vscBar.SetFocus
                    zlCommFun.PressKey vbKeyPageDown
                End If
            Case 7864320   '向上滚
                If frmDoctorManage.vscBar.Value <> 0 Then
                    frmDoctorManage.vscBar.SetFocus
                    zlCommFun.PressKey vbKeyPageUp
                End If
        End Select
    End Select
    FmgFlexScroll = CallWindowProc(glngPreHWnd, hWnd, wMsg, wParam, lParam)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ShowSpecChar(frmParent As Object) As String
'功能：以模态窗体运行特殊字符程序
'参数：frmParent=调用父窗体
'返回：选择的特殊字符串；取消操作返回空
    Dim frmNew As frmSpecChar
    Set frmNew = New frmSpecChar
    frmNew.Show 1, frmParent
    If gblnOK Then ShowSpecChar = frmNew.mstrChar
End Function

Public Sub ArrayIcons(objLvw As ListView, Optional intBegin As Integer = 1, Optional blnShow As Boolean)
'功能：根据第一个图标的位置重新排列所有图标
    Dim i As Integer, t As Long
    Dim r As RECT

    Call GetClientRect(objLvw.hWnd, r)
    
    If blnShow Then
        If objLvw.ListItems(intBegin).Top < 30 Then
           objLvw.ListItems(intBegin).Top = 30
        ElseIf objLvw.ListItems(intBegin).Top + objLvw.ListItems(intBegin).Height > (r.Bottom - r.Top) * Screen.TwipsPerPixelY Then
            objLvw.ListItems(intBegin).Top = (r.Bottom - r.Top) * Screen.TwipsPerPixelY - objLvw.ListItems(intBegin).Height
        End If
    End If
    
    '下面的图标
    t = objLvw.ListItems(intBegin).Top
    For i = intBegin To objLvw.ListItems.Count
        With objLvw.ListItems(i)
            'Item的Width包含文字部分,Left仅指图标
            .Left = ((r.Right - r.Left - objLvw.Icons.ImageWidth) * Screen.TwipsPerPixelX) / 2
            .Top = t
            t = t + .Height
        End With
    Next
    
    '上面的图标
    t = objLvw.ListItems(intBegin).Top
    For i = intBegin To 1 Step -1
        With objLvw.ListItems(i)
            'Item的Width包含文字部分,Left仅指图标
            .Left = ((r.Right - r.Left - objLvw.Icons.ImageWidth) * Screen.TwipsPerPixelX) / 2
            .Top = t
            t = t - .Height
        End With
    Next
End Sub

Public Function GetArrayByStr(ByVal strInput As String, ByVal lngLength As Long, ByVal strSplitChar As String) As Variant
    '根据传入的字符串进行分解，大于指定字符长度就需要进行分解，结果保存到数组中
    '入参：strInput-输入的字符串；strSplitChar-字符串中内容的分隔符
    '返回：数组，其中数组成员的字符长度不超过指定长度
    Dim strArray As Variant
    Dim arrTmp As Variant
    Dim strTmp As String
    Dim lngCount As Long
    Dim i As Long
    
    strArray = Array()
   
    '大于指定字符时就需要分解
    If Len(strInput) > lngLength Then
        If strSplitChar = "" Then
            '无分隔符时
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
            '有分隔符时
            arrTmp = Split(strInput & strSplitChar, strSplitChar)
            lngCount = UBound(arrTmp)
        
            For i = 0 To lngCount
                If arrTmp(i) <> "" Then
                    '有分隔符的需要保持分隔符之间字符的完整性，不能把分隔符之间的字符拆开
                    If Len(IIf(strTmp = "", "", strTmp & strSplitChar) & arrTmp(i)) > lngLength Then
                        ReDim Preserve strArray(UBound(strArray) + 1)
                        strArray(UBound(strArray)) = strTmp
                        strTmp = arrTmp(i)
                    Else
                        strTmp = IIf(strTmp = "", "", strTmp & strSplitChar) & arrTmp(i)
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

Public Function CheckBatches(ByVal bln药库分批 As Boolean, ByVal bln药房分批 As Boolean) As Boolean
    '功能：检查药库分批药房不分批时，部门性质是否同时有设置了药库药房
    
    Dim rs部门性质 As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    If bln药库分批 = True And bln药房分批 = False Then
        gstrSql = "Select 1" & vbNewLine & _
                        "From 部门性质说明 T" & vbNewLine & _
                        "Where t.部门id In" & vbNewLine & _
                        "      (Select Distinct t.部门id From 部门性质说明 T Where t.工作性质 Like '%药库')" & vbNewLine & _
                        "      And (t.工作性质 Like '%药房' or t.工作性质 Like '%制剂室')"
                        
        Set rs部门性质 = zlDatabase.OpenSQLRecord(gstrSql, "是否有部门性质同时设置了药库药房")
        If rs部门性质.RecordCount > 0 Then
            CheckBatches = True
        End If
    End If
    
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function




