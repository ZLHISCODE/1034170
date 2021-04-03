Attribute VB_Name = "mdlStuff"
 Option Explicit

Public gcnOracle As ADODB.Connection        '公共数据库连接，特别注意：不能设置为新的实例
Public gstrPrivs As String                   '当前用户具有的当前模块的功能
Public gstrSysName As String                '系统名称
Public glngModul As Long
Public glngSys As Long
Public gstrAviPath As String
Public gstrVersion As String
Public gstrMatchMethod As String
Public gbytSimpleCodeTrans As Byte          '卡片界面是否允许简码切换控制

Public gstrDBUser As String                 '当前数据库用户
Public glngUserId As Long                   '当前用户id
Public gstrUserCode As String               '当前用户编码
Public gstrUserName As String               '当前用户姓名
Public gstrUserAbbr As String               '当前用户简码

Public glngDeptId As Long                   '当前用户部门id
Public gstrDeptCode As String               '当前用户部门编码
Public gstrDeptName As String               '当前用户部门名称

Public gstrUnitName As String '用户单位名称
Public gfrmMain As Object

Public gstrSQL As String
Public gblnOK As Boolean
Public gstrIme As String

Public gobjSquareCard As Object             '一卡通接口
Public gstrCardType As String           '银行卡类别，格式：短名|全名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密);…
Public gintCardCount As Integer  '卡数量
Public gblnIncomeItem As Boolean            '记录卫材目录管理中是否设置了收入项目

Public gobjPlugIn As Object             '外挂接口

Public Const glngRowByFocus = &HFFE3C8
Public Const glngRowByNotFocus = &HF4F4EA
Public Const glngFixedForeColorByFocus = &HFF0000
Public Const glngFixedForeColorNotFocus = &H80000012

'药品金额、价格、数量最大精度
Public Type Type_Digits
    Digit_金额 As Integer
    Digit_成本价 As Integer
    Digit_零售价 As Integer
    Digit_数量 As Integer
End Type
Public gtype_UserDrugDigits As Type_Digits

'消费卡格式
Public Enum gCardFormat
    短名 = 0
    全名 = 1
    刷卡标志 = 2
    卡类别ID = 3
    卡号长度 = 4
    缺省标志 = 5
    是否存在帐户 = 6
    卡号密文 = 7
End Enum

Public Type TYPE_USER_INFO
    Id As Long
    部门ID As Long
    编号 As String
    姓名 As String
    简码 As String
    用户名 As String
End Type
Public gOraFmt_Max As g_FmtString


Public UserInfo As TYPE_USER_INFO
Public Declare Function SetFocusHwnd Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

'切换到指定的输入法。
Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long
'返回系统中可用的输入法个数及各输入法所在Layout,包括英文输入法。
Public Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
'获取某个输入法的名称
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
'判断某个输入法是否中文输入法
Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long
'获取指定输入法所在Layout,参数为0时表示当前输入法。
Public Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
'获取当前输入法所在Layout名
Public Declare Function GetKeyboardLayoutName Lib "user32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
'根据输入法Layout名将该输入法切换到输入法切换顺序的最前头(重新启动后无效),flags参数=KLF_REORDER
Public Declare Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (ByVal pwszKLID As String, ByVal flags As Long) As Long

Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
        x As Long
        y As Long
End Type
'开始日期的标志
Public Enum StartDayFlag
    FirstDayOfWeek = 0
    FirstDayOfMonth = 1
    FirstDayOfQuarter = 2
    FirstDayOfHalfYear = 3
    FirstDayOfyear = 4
End Enum
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Public Sub zlPlugIn_Ini(ByVal lngSys As Long, ByVal lngModul As Long, objPlugIn As Object)
    '外挂扩展接口初始化
    If objPlugIn Is Nothing Then
        On Error Resume Next
        Set objPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        If Not objPlugIn Is Nothing Then
            Call objPlugIn.Initialize(gcnOracle, lngSys, lngModul)
            If InStr(",438,0,", "," & err.Number & ",") = 0 Then
                MsgBox "zlPlugIn 外挂部件执行 Initialize 时出错：" & vbCrLf & err.Number & vbCrLf & err.Description, vbInformation, gstrSysName
            End If
        End If
        err.Clear: On Error GoTo 0
    End If
End Sub

Public Sub zlPlugIn_Unload(objPlugIn As Object)
    '卸载外挂接口
    Set objPlugIn = Nothing
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
Public Function GetUserInfo() As Boolean
'功能：获取登陆用户信息
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = zlDatabase.GetUserInfo
    
    UserInfo.用户名 = gstrDBUser
    UserInfo.姓名 = gstrDBUser
    If Not rsTmp.EOF Then
        UserInfo.Id = rsTmp!Id
        UserInfo.编号 = rsTmp!编号
        UserInfo.部门ID = IIf(IsNull(rsTmp!部门ID), 0, rsTmp!部门ID)
        UserInfo.简码 = IIf(IsNull(rsTmp!简码), "", rsTmp!简码)
        UserInfo.姓名 = IIf(IsNull(rsTmp!姓名), "", rsTmp!姓名)
        UserInfo.用户名 = UserInfo.姓名
        gstrUserName = IIf(IsNull(rsTmp!姓名), "", rsTmp!姓名)
        GetUserInfo = True
    End If
    Exit Function
errH:
    Call ErrCenter
    Call SaveErrLog
End Function

Public Function NeedName(strList As String) As String
    If InStr(strList, Chr(13)) > 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, Chr(13)) + 1))
    ElseIf InStr(strList, "]") > 0 And InStr(strList, "-") = 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, "]") + 1))
    ElseIf InStr(strList, ")") > 0 And InStr(strList, "-") = 0 Then
        NeedName = LTrim(Mid(strList, InStr(strList, ")") + 1))
    Else
        NeedName = LTrim(Mid(strList, InStr(strList, "-") + 1))
    End If
End Function

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

Private Function IsType(ByVal varType As DataTypeEnum, ByVal varBase As DataTypeEnum) As Boolean
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
Public Function GetDownCodeLength(ByVal strID As String, ByVal strTableName As String, Optional ByVal strWhere As String) As Long
    '功能描述：读取指定表的本级编码的最大长度
    '输入参数：本级ID，表名
    '输出参数：成功返回 下级最大编码; 否者返回 0
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    err = 0
    On Error GoTo Error_Handle
    If strID = "" Then
        strSQL = "select nvl(max(Vsize(编码)),0) as LenCode from " & strTableName & " start with 上级ID is null " & strWhere & " connect by prior id=上级id"
    Else
        strSQL = "select nvl(max(Vsize(编码)),0) as LenCode from " & strTableName & " start with ID=" & strID & strWhere & " connect by prior id=上级id"
    End If
    
    Call zlDatabase.OpenRecordset(rsTemp, strSQL, "读取指定表的本级编码的最大长度")
    
    If rsTemp.EOF Then
        GetDownCodeLength = 0
    Else
        GetDownCodeLength = rsTemp.Fields("LenCode").Value
    End If
    Exit Function
Error_Handle:
    Call ErrCenter
    Call SaveErrLog
    GetDownCodeLength = 0
End Function

Public Function GetLocalCodeLength(ByVal str上级ID As String, ByVal strTableName As String, Optional ByVal strWhere As String) As Long
    '功能描述：读取指定表的本级编码的最大长度
    '输入参数：上级ID，表名
    '输出参数：成功返回 最大编码; 否者返回 0
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    err = 0
    On Error GoTo Error_Handle
    If str上级ID = "" Then
        strSQL = "select nvl(max(Vsize(编码)),0) as LenCode from " & strTableName & " where 上级ID is null" & strWhere
    Else
        strSQL = "select nvl(max(Vsize(编码)),0) as LenCode from " & strTableName & " where 上级ID=" & str上级ID & strWhere
    End If
    
    Call zlDatabase.OpenRecordset(rsTemp, strSQL, "mdlCureBase")
    
    If rsTemp.EOF Then
        GetLocalCodeLength = 0
    Else
        GetLocalCodeLength = rsTemp.Fields("LenCode").Value
    End If
    Exit Function
Error_Handle:
    Call ErrCenter
    Call SaveErrLog
    GetLocalCodeLength = 0
End Function

Public Function GetParentCode(ByVal str上级ID As String, ByVal strTableName As String) As String
    '功能描述：读取上级编码
    '输入参数：上级ID,表名
    '输出参数：成功返回 上级编码; 否者返回 空
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    err = 0
    On Error GoTo Error_Handle
    If str上级ID = "" Then
        GetParentCode = ""
        Exit Function
    End If
    
    strSQL = "select 编码 from " & strTableName & " where ID=[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取上级编码", str上级ID)
    
    If rsTemp.EOF Then
        GetParentCode = ""
    Else
        GetParentCode = rsTemp.Fields("编码").Value
    End If
    Exit Function
Error_Handle:
    Call ErrCenter
    Call SaveErrLog
    GetParentCode = ""
End Function

Public Function GetMaxLocalCode(ByVal str上级ID As String, ByVal strTableName As String, Optional ByVal strWhere As String) As String
    '功能描述：根据指定表的上级ID 读取本级的最大编码
    '输入参数：上级ID,表名
    '输出参数：成功返回 最大编码; 否者返回 空
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim intCode As Integer, strCode As String, strAllCode As String
    Dim intLength   As Integer
    err = 0
    On Error GoTo Error_Handle
    If str上级ID = "" Then
        strSQL = "select max(to_number(编码))+1 as MaxCode from " & strTableName & " where 上级ID is null" & strWhere
    Else
        strSQL = "select nvl(max(to_number(编码)),0)+1 as MaxCode from " & strTableName & " where 上级ID=" & str上级ID & strWhere
    End If
    intCode = GetLocalCodeLength(str上级ID, strTableName, strWhere)
    
    Call zlDatabase.OpenRecordset(rsTemp, strSQL, "根据指定表的上级ID 读取本级的最大编码")
    If rsTemp.EOF Then
        GetMaxLocalCode = ""
        Exit Function
    End If
    intLength = intCode - Len(IIf(IsNull(rsTemp.Fields("MaxCode").Value), 0, rsTemp.Fields("MaxCode").Value))
    strAllCode = String(IIf(intLength < 0, 0, intLength), "0") & rsTemp.Fields("MaxCode").Value
    GetMaxLocalCode = Mid(strAllCode, Len(GetParentCode(str上级ID, strTableName)) + 1)
    Exit Function
Error_Handle:
    Call ErrCenter
    Call SaveErrLog
    GetMaxLocalCode = ""
End Function

Public Sub SetGridFocus(ByVal objGrid As VSFlexGrid, ByVal blnGetFoucs As Boolean)
    With objGrid
        If blnGetFoucs Then
            .GridColorFixed = &H80000008
            .GridColor = &H80000008
            .ForeColorFixed = glngFixedForeColorByFocus
            .BackColorSel = glngRowByFocus
        Else
            .GridColorFixed = &H80000011
            .GridColor = &H80000011
            .ForeColorFixed = glngFixedForeColorNotFocus
            .BackColorSel = glngRowByNotFocus
        End If
    End With
End Sub
 
Public Function GetFormat(ByVal dblInput As Double, ByVal intDotBit As Integer) As String
    GetFormat = Format(dblInput, "#0." & String(intDotBit, "0"))
End Function

Public Function BinTOHex(sString As String) As String
    Dim lngLoop As Integer, lngTemp As Long, lngJLoop As Integer, lngTmp As Long
    lngTemp = 0
    For lngLoop = 1 To Len(sString)
        If Mid(sString, lngLoop, 1) = "1" Then
            lngTmp = 1
            For lngJLoop = 0 To lngLoop - 2
                lngTmp = lngTmp * 2
            Next
        Else
            lngTmp = 0
        End If
        lngTemp = lngTemp + lngTmp
    Next
    BinTOHex = CStr(lngTemp)
End Function
Public Sub ShowMsgBox(ByVal strMsgInfor As String, Optional blnYesNo As Boolean = False, Optional ByRef blnYes As Boolean)
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

Public Sub zlChangeCode(ByVal strTableName As String, _
    ByVal lng上级id As Long, _
    ByVal txtUpCode As TextBox, _
    ByVal txtCode As TextBox, _
    Optional ByVal chkChangeCode As CheckBox = Nothing, _
    Optional ByVal strCaption As String = "")
    '------------------------------------------------------------------------------------
    '功能：根据选择的上级确定当前的编码，并在上级及本级中显示出来
    '参数：strTableName-存在分类的表名
    '      lng上级ID-选择的上级
    '      TxtUpCode-显示的上级文本框
    '      TxtUpCode-显示的本级文本框
    '      chkChangeCode-设置是否改变原有数据库中的历史编码选择控件
    '      strCaption-调用窗体的Capiton
    '注意：表中必需有ID,上级id,编码
    '------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim intMaxCodeLen As Integer  '确定编码的实际长度
    err = 0: On Error GoTo ErrHand
    
   chkChangeCode.Value = 0
   chkChangeCode.Enabled = True
   
    If lng上级id = 0 Then
        txtUpCode.Text = ""
        gstrSQL = "select max(编码) as 编码 From " & strTableName & " Where 上级ID is null "
        zlDatabase.OpenRecordset rsTemp, gstrSQL, strCaption
            
        With rsTemp
            intMaxCodeLen = .Fields("编码").DefinedSize
            If IsNull(!编码) Then
                txtCode.Text = "01"
                txtCode.MaxLength = intMaxCodeLen
                txtCode.Tag = txtCode.MaxLength
                chkChangeCode.Value = 1
                chkChangeCode.Enabled = False
            Else
                txtCode.MaxLength = Len(Trim(!编码))
                txtCode.Tag = txtCode.MaxLength
                If !编码 = String(txtCode.MaxLength, "9") Then
                    If txtCode.MaxLength >= intMaxCodeLen Then
                        ShowMsgBox "最大编码和编码长度已经达到最大限制，无法递增编码"
                        txtCode.Text = Space(txtCode.MaxLength)
                       chkChangeCode.Value = 0
                       chkChangeCode.Enabled = False
                    Else
                        ShowMsgBox "最大编码已经达到本级限制，你可以扩充编码长度以满足需要"
                        txtCode.Text = "1" & String(txtCode.MaxLength, "0")
                        txtCode.MaxLength = txtCode.MaxLength + 1
                        txtCode.Tag = txtCode.MaxLength
                       chkChangeCode.Value = 1
                    End If
                Else
                    txtCode.Text = Format(Mid(!编码, Len(txtUpCode.Text) + 1) + 1, String(txtCode.MaxLength, "0"))
                End If
            End If
        End With
        Exit Sub
   End If
   '确定上级编码
   
    gstrSQL = "Select 编码 From " & strTableName & " where id=[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, strCaption, lng上级id)
    
    If Not rsTemp.EOF Then
        txtUpCode.Text = zlCommFun.NVL(rsTemp!编码)
    End If
    
    '先确定是否有下级
    gstrSQL = "select nvl(max(编码),'') as 编码  From " & strTableName & " Where  上级ID =[1] "
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, strCaption, lng上级id)
    
    intMaxCodeLen = rsTemp.Fields("编码").DefinedSize

    If zlCommFun.NVL(rsTemp!编码) = "" Then
        '不存在下级
        '根据上级ID取上级编码
'        gstrSQL = "Select 编码 From " & strTableName & " where id=" & lng上级id
'        zlDatabase.OpenRecordset rsTemp, gstrSQL, strCaption
'        txtUpCode.Text = zlCommFun.Nvl(rsTemp!编码)
        txtCode.MaxLength = intMaxCodeLen - Len(txtUpCode.Text)
        txtCode.Tag = txtCode.MaxLength
        If txtCode.MaxLength > 1 Then
            txtCode.Text = "01"
        Else
            txtCode.Text = "1"
        End If
        chkChangeCode.Value = 1
        chkChangeCode.Enabled = False
        Exit Sub
    End If
    
    With rsTemp
        txtCode.MaxLength = Len(!编码) - Len(txtUpCode.Text)
        txtCode.Tag = txtCode.MaxLength
        If Mid(!编码, Len(txtUpCode.Text) + 1) = String(txtCode.MaxLength, "9") Then
            If Len(txtUpCode.Text) + txtCode.MaxLength >= intMaxCodeLen Then
                ShowMsgBox "该分类下级最大编码和编码长度已经达到最大限制，无法递增编码"
                txtCode.Text = Space(txtCode.MaxLength)
               chkChangeCode.Value = 0
               chkChangeCode.Enabled = False
            Else
                ShowMsgBox "该分类下级最大编码已经达到本级限制，你可以扩充编码长度以满足需要"
                txtCode.Text = "1" & String(txtCode.MaxLength, "0")
                txtCode.MaxLength = txtCode.MaxLength + 1
                txtCode.Tag = txtCode.MaxLength
               chkChangeCode.Value = 1
            End If
        Else
            txtCode.Text = Format(Mid(!编码, Len(txtUpCode.Text) + 1) + 1, String(txtCode.MaxLength, "0"))
        End If
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub ImeLanguage(ByVal blnOpen As Boolean)
    '-----------------------------------------------------------------------------------
    '功能: 打开/关闭输入法
    '参数: blnOpen-是打开还是关闭(true为打开,false为关闭)
    '返回：
    '-----------------------------------------------------------------------------------
    If blnOpen Then
        zlCommFun.OpenIme (True)
    Else
        zlCommFun.OpenIme False
    End If
End Sub


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
        zlCommFun.OpenIme False
    End If
End Sub

Public Function DepotProperty(ByVal lng人员id As Long) As Boolean
    Dim rsCheck As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    '返回指定人员是否具有药库性质
    gstrSQL = "Select Distinct 工作性质 From 部门人员 B,部门性质说明 A " & _
             " Where A.工作性质 = '卫材库' And " & _
             " A.部门id = B.部门id And B.人员id = [1]"
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "取部门性质", lng人员id)
    If rsCheck.RecordCount <> 0 Then
        DepotProperty = True
        Exit Function
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ShowCostPrice() As Boolean
    Dim blnCostPrice As Boolean
    
    On Error GoTo ErrHandle
    '是否允许发料药房人员查看单据的成本价
    blnCostPrice = Val(zlDatabase.GetPara(190, 100, , 0))
    
    '药库人员不管，只管药房人员，以参数控制为准
    If DepotProperty(UserInfo.Id) Then
        ShowCostPrice = True
    Else
        ShowCostPrice = blnCostPrice
    End If
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function NVL(rsObj As Field, Optional ByVal varValue As Variant = "") As Variant
    '-----------------------------------------------------------------------------------
    '功能:取某字段的值
    '参数:rsObj          被检查的字段
    '     varValue       当rsObj为NULL值时的取新值
    '返回:如果不为空值,返回原来的值,如果为空值,则返回指定的varValue值
    '-----------------------------------------------------------------------------------
    If IsNull(rsObj) Then
        NVL = varValue
    Else
        NVL = rsObj
    End If
End Function
Public Function Dec2Bin(bDec As Byte) As String

    '功能：十进制转为二进制函数
    '用法：String  Dec2Bin(Bdec as Byte)
    '返回：  十进制的二进制 字符串(String)
    '错误：  返回"0"
    Dim strBin As String

    On Error GoTo err
    If bDec > 255 Then
        Dec2Bin = "-1"
        Exit Function
    End If
    strBin = ""
    '转为字符串
    While bDec > 0
        strBin = bDec Mod 2 & strBin
        bDec = Fix(bDec / 2)
    Wend
    '补零足8位
    If Len(strBin) < 9 Then
        While Len(strBin) < 8
            strBin = "0" & strBin
        Wend
    End If
    Dec2Bin = strBin
    Exit Function
err:
   Dec2Bin = "0"
End Function

Public Function Bin2Dec(strBin As String) As Long
    '功能：二进制转为十进制函数
    '用法：Long  bin2dec(strBin as String)
    '返回：  二进制的十进制 长整数（Long）
    '错误：  返回-1
    Dim lDec As Long
    Dim lCount As Long
    Dim i As Long
    
    On Error GoTo ErrHand
    lDec = 0
    If strBin = "" Then strBin = "0"
    lCount = Len(strBin)
    For i = 1 To lCount
        lDec = lDec + CInt(Left(strBin, 1)) * 2 ^ (Len(strBin) - 1)
        strBin = Right(strBin, Len(strBin) - 1)
        DoEvents
    Next
    Bin2Dec = lDec
    Exit Function
ErrHand:
    Bin2Dec = -1
End Function

Public Sub SetColumnSort(ByVal mshFilter As MSHFlexGrid, ByRef intPreCol As Integer, ByRef intPreSort As Integer, Optional blnNum As Boolean = False)
    '----------------------------------------------------------------------------------------------------------------
    '功能描述：对指定的列进行排序
    '输入参数：mshFilter-指定的网格
    '          intPreCol-上次列
    '           intPreSort-上次排序
    '           blnNum-是否为数字列
    '输出参数：
    '返回：
    '----------------------------------------------------------------------------------------------------------------
    
    Dim intCol As Integer
    Dim intRow As Integer
    Dim strTemp As String
    
    With mshFilter
        If .Rows > 1 Then
            .Redraw = False
            intCol = .MouseCol
            .Col = intCol
            .ColSel = intCol
            strTemp = .TextMatrix(.Row, 0)
            If blnNum Then
                If intCol = intPreCol And intPreSort = flexSortNumericDescending Then
                   .Sort = flexSortNumericAscending
                   intPreSort = flexSortNumericAscending
                Else
                   .Sort = flexSortNumericDescending
                   intPreSort = flexSortNumericDescending
                End If
            Else
                    If intCol = intPreCol And intPreSort = flexSortStringNoCaseDescending Then
                       .Sort = flexSortStringNoCaseAscending
                       intPreSort = flexSortStringNoCaseAscending
                    Else
                       .Sort = flexSortStringNoCaseDescending
                       intPreSort = flexSortStringNoCaseDescending
                    End If
            End If
            
            intPreCol = intCol
            .Row = FindRow(mshFilter, strTemp, 0)
            If .RowPos(.Row) + .RowHeight(.Row) > .Height Then
                .TopRow = .Row
            Else
                .TopRow = 1
            End If
            .Col = 0
            .ColSel = .Cols - 1
            .Redraw = True
            .SetFocus
        Else
            .ColSel = 0
        End If
    End With
End Sub

Public Function FindRow(ByVal mshgrd As MSHFlexGrid, ByVal varTemp As Variant, ByVal intCol As Integer) As Integer
    '----------------------------------------------------------------------------------------------------------------
    '功能描述：查找符合条件的行
    '输入参数：varTemp-指定的值
    '           mshGrd-指定网络
    '           intCol-指定的列
    '输出参数：
    '返回：成功返回找到的行
    '----------------------------------------------------------------------------------------------------------------
    
    Dim intTmp As Integer
    
    With mshgrd
        For intTmp = 1 To .Rows - 1
            If IsDate(varTemp) Then
               If Format(.TextMatrix(intTmp, intCol), "yyyy-mm-dd") = Format(varTemp, "yyyy-mm-dd") Then
                  FindRow = intTmp
                  Exit Function
               End If
            Else
                If .TextMatrix(intTmp, intCol) = varTemp Then
                  FindRow = intTmp
                  Exit Function
                End If
            End If
        Next
    End With
    FindRow = 1
End Function

'按编码，名称，别名查找某一列
Public Function FindRownew(ByVal mshBill As BillEdit, ByVal int比较列 As Integer, _
    ByVal str比较值 As String, ByVal blnFirst As Boolean) As Boolean
    Dim intStartRow As Integer
    Dim intRow As Integer
    Dim strSpell As String
    Dim strCode As String
    Dim rsCode As New Recordset
    
    On Error GoTo ErrHandle
    FindRownew = True
    With mshBill
        If .Rows = 2 Then Exit Function
        If str比较值 = "" Then Exit Function
        
        If blnFirst = True Then
            intStartRow = 0
        Else
            intStartRow = .Row
        End If
        If intStartRow = .Rows - 1 Then
            intStartRow = 1
        Else
            intStartRow = intStartRow + 1
        End If
        
        For intRow = intStartRow To .Rows - 1
            If .TextMatrix(intRow, int比较列) <> "" Then
                strCode = .TextMatrix(intRow, int比较列)
                If InStr(1, UCase(strCode), UCase(str比较值)) <> 0 Then
                    .SetFocus
                    .Row = intRow
                    .Col = int比较列
                    .MsfObj.TopRow = .Row
                    .SetRowColor CLng(.Row), &HFFCECE, True
                    Exit Function
                End If
            End If
        Next
        gstrSQL = "" & _
        " SELECT DISTINCT b.编码 " & _
        " FROM (SELECT DISTINCT A.收费细目id " & _
        "       FROM 收费项目别名 A" & _
        "       Where A.简码 LIKE upper([1]) " & _
        "      ) A, 收费项目目录 B " & _
        " Where a.收费细目id = b.ID And (b.站点=[2] or b.站点 is null) "
        
        Set rsCode = zlDatabase.OpenSQLRecord(gstrSQL, "查找指定卫生材料", GetMatchingSting(str比较值, False), gstrNodeNo)
        If rsCode.EOF Then
            FindRownew = False
            Exit Function
        End If
        
        For intRow = intStartRow To .Rows - 1
            If .TextMatrix(intRow, int比较列) <> "" Then
                strCode = .TextMatrix(intRow, int比较列)
                rsCode.MoveFirst
                Do While Not rsCode.EOF
                    If InStr(1, UCase(strCode), UCase(rsCode!编码)) <> 0 Then
                        .SetFocus
                        .Row = intRow
                        .Col = int比较列
                        .MsfObj.TopRow = .Row
                        .SetRowColor CLng(.Row), &HFFCECE, True
                        rsCode.Close
                        Exit Function
                    End If
                    rsCode.MoveNext
                Loop
            End If
        Next
        rsCode.Close
    End With
    FindRownew = False
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
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
    err = 0: On Error GoTo ErrHand:
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
    err = 0
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
Public Function 判断只具备发料部门(ByVal lng部门ID As Long) As Boolean
    '判断只具备发料备性质的:即除取卫材库和制剂室性制的所有具备发料部门性质的部门
    'lng部门id-部门id
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    判断只具备发料部门 = False
    gstrSQL = "select 工作性质, 部门id, 服务对象 from 部门性质说明 where 部门id =[1] And 工作性质='发料部门'"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取发料部门的工作性质", lng部门ID)
    
    
    If rsTemp.RecordCount = 0 Then
        Exit Function
    End If
    gstrSQL = "select 工作性质, 部门id, 服务对象 from 部门性质说明 where 部门id =[1] And 工作性质 in( '卫材库','制剂室','虚拟库房')"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取发料部门的工作性质", lng部门ID)
    
    If rsTemp.RecordCount <> 0 Then
        Exit Function
    End If
    判断只具备发料部门 = True
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function CheckNOExists(ByVal int单据 As Integer, ByVal strNo As String) As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "Select 1 From 药品收发记录 Where NO=[2] And 单据=[1] And Rownum<2"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断是否存在该单据", int单据, strNo)
    If rsTemp.RecordCount = 0 Then Exit Function
    ShowMsgBox "已经存在该单据号(" & strNo & ")"
    CheckNOExists = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetFullNO(ByVal strNo As String, ByVal intNum As Integer, Optional ByVal lng科室ID As Long) As String
'功能：由用户输入的部份单号，返回全部的单号。
'参数：intNum=项目序号,为0时固定按年产生
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, intType As Integer
    Dim curDate As Date
    Dim intYear As Integer
    Dim PreFixNO As String  '年份前缀
    Dim strPre As String    '最大号码表中前2位
    Dim str编号 As String
    Dim dateCurDate As Date
    Dim intMonth As Integer
    Dim strMonth As String
    
    On Error GoTo errH
    
    dateCurDate = zlDatabase.Currentdate
    intYear = Format(dateCurDate, "YYYY") - 1990
    PreFixNO = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
    intMonth = Month(dateCurDate)
    strMonth = intMonth
    strMonth = String(2 - Len(strMonth), "0") & strMonth
    
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
    
    strSQL = "Select 编号规则,最大号码,Sysdate as 日期 From 号码控制表 Where 项目序号=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetFullNO", intNum)
        
    If Not rsTmp.EOF Then
        intType = NVL(rsTmp!编号规则, 0)
        curDate = rsTmp!日期
        strPre = Left(NVL(rsTmp!最大号码, PreFixNO & "0"), 2)
    End If
    
    If intType = 0 Then
        '按年编号
        GetFullNO = strPre & Format(Right(strNo, 6), "000000")
    ElseIf intType = 1 Then
        '按日编号
        strSQL = Format(CDate("1992-" & Format(rsTmp!日期, "MM-dd")) - CDate("1992-01-01") + 1, "000")
        GetFullNO = PreFixNO & strSQL & Format(Right(strNo, 4), "0000")
    ElseIf intType = 2 Then
        '按科室分月编码
        gstrSQL = "Select 编号 From 科室号码表 Where 项目序号=[1] And Nvl(科室ID,0)=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "GetFullNO", intNum, lng科室ID)
        
        If rsTmp.RecordCount = 0 Then
            MsgBox "还未设置科室编号，无法产生号码！", vbInformation, gstrSysName
            Exit Function
        End If
        If NVL(rsTmp!编号) = "" Then
            MsgBox "还未设置科室编号，无法产生号码！", vbInformation, gstrSysName
            Exit Function
        End If
        str编号 = NVL(rsTmp!编号)
        
        '小于四位，按本月产生号码
        '五位或六位，则认为是指定月份的号码
        '七位，则认为是产生本年指定科室、月份的号码
        '大于等于八位，不处理
        If Len(strNo) <= 4 Then
            GetFullNO = PreFixNO & str编号 & strMonth & String(4 - Len(strNo), "0") & strNo
        ElseIf Len(strNo) <= 6 Then
            GetFullNO = String(6 - Len(strNo), "0") & GetFullNO
            GetFullNO = PreFixNO & str编号 & GetFullNO
        ElseIf Len(strNo) = 7 Then
            GetFullNO = PreFixNO & GetFullNO
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function IsHavePrivs(ByVal strPrivs As String, ByVal strMyPriv As String) As Boolean
    IsHavePrivs = InStr(";" & strPrivs & ";", ";" & strMyPriv & ";") > 0
End Function

Public Sub ExecuteProcedureArrAy(ByVal cllProcs As Variant, ByVal strCaption As String, Optional blnNoCommit As Boolean = False)
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:执行相关的Oracle过程集
    '参数:cllProcs-oracle过程集
    '     strCaption -执行过程的父窗口标题
    '     blnNOCommit-执行完过程后,不提交数据
    '-------------------------------------------------------------------------------------------------------------------------
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

Public Sub AddArray(ByRef cllData As Collection, ByVal strSQL As String)
    Dim i As Long
    i = cllData.Count + 1
    cllData.Add strSQL, "K" & i
End Sub

Public Function Check负出库按最后进价计算() As Boolean
    '功能:确定系统参数在负数情况下的成本计算方式
    Check负出库按最后进价计算 = Val(zlDatabase.GetPara(120, glngSys, 0)) = 1
End Function
Public Function 验证出库差价计算(ByVal lng库房ID As Long, ByVal lng材料ID As Long, ByVal lng批次 As Long, ByVal lng比例系数 As Long, _
                    ByVal dbl库存差价 As Double, ByVal dbl库存金额 As Double, _
                    ByVal dbl指导差价率 As Double, ByVal dbl数量 As Double, ByVal dbl零售金额 As Double, _
                    ByRef dblOut差价 As Double, ByRef dblOut购价 As Double, ByRef dblOut成本金额 As Double) As Boolean
    '------------------------------------------------------------------------------------------------------------
    ' 功能:获取本次的成本价和差价
    ' 计算公式:
    '       1.库存金额<=0：
    '         1) 库存金额-实际差价<=0 Or dbl库存数量 < 0
    '               a.卫材负数出库计算方式=1:
    '                      a)最后进价＝0：
    '                           差价=零售金额*指导差价率
    '                           成本价=（出库金额-出库差价）/数量
    '                      b)最后进价>0
    '                           成本价=最后进价
    '                           差价＝零售金额-数量*成本价
    '               b.卫材负数出库计算方式<>1
    '                           差价=零售金额*指导差价率
    '                           成本价=（出库金额-出库差价）/数量
    '          2)库存金额-实际差价>0
    '                成本价= (库存金额-实际差价)/库存数量
    '                差价＝零售金额-数量*成本价
    '        2.库存金额>0
    '                   出库差价=出库金额*（实际差价/实际金额）
    '                  成本价=（出库金额-出库差价）/数量
    '------------------------------------------------------------------------------------------------------------
    Dim dbl差价 As Double, dbl购价 As Double, dbl库存数量 As Double
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    If dbl数量 = 0 Then Exit Function
    dbl购价 = Get成本价(lng材料ID, lng库房ID, lng批次) * lng比例系数
    dbl差价 = dbl零售金额 - dbl购价 * dbl数量
    
'    If dbl库存金额 <= 0 Then
'        If dbl库存金额 - dbl库存差价 > 0 Then
'            gstrSQL = "Select (实际金额-实际差价)/实际数量 as 成本价 From 药品库存 where 库房id=[1] and 药品id=[2] and nvl(批次,0)=[3] and nvl(实际数量,0)>0"
'            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取成本价", lng库房ID, lng材料ID, lng批次)
'            If rsTemp.EOF = False Then
'                dbl购价 = Val(NVL(rsTemp!成本价)) * lng比例系数
'            End If
'        End If
'
'        If dbl库存金额 - dbl库存差价 <= 0 Or dbl购价 <= 0 Then
'            If Check负出库按最后进价计算 = True Then
'                dbl购价 = Get最后进价(lng材料ID) * lng比例系数
'                If dbl购价 = 0 Then
'                    dbl差价 = dbl零售金额 * dbl指导差价率
'                    dbl购价 = (dbl零售金额 - dbl差价) / Dbl数量
'                Else
'                    dbl差价 = dbl零售金额 - Dbl数量 * dbl购价
'                End If
'            Else
'                    dbl差价 = dbl零售金额 * dbl指导差价率
'                    dbl购价 = (dbl零售金额 - dbl差价) / Dbl数量
'            End If
'        Else
'            'dbl库存金额 - dbl库存差价>0
'            dbl差价 = dbl零售金额 - dbl购价 * Dbl数量
'        End If
'    Else
'                dbl差价 = dbl零售金额 * (dbl库存差价 / dbl库存金额)
'                dbl购价 = (dbl零售金额 - dbl差价) / Dbl数量
'    End If
    
    dblOut成本金额 = Round(dbl数量 * dbl购价, 7)
    dblOut差价 = Round(dbl差价, 7)
    dblOut购价 = Round(dbl购价, 7)
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get最后进价(ByVal lng材料ID As Long) As Double
    '功能:获取最后进价
    '参数:lng材料ID
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "Select 成本价 From 材料特性 where 材料id=[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取成本价", lng材料ID)
    
    If rsTemp.EOF Then
        Get最后进价 = 0
    Else
        Get最后进价 = Val(NVL(rsTemp!成本价))
    End If
    rsTemp.Close
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ISCHECK不强制控制指导价格() As Boolean
    '功能:判断是否不强制要求控制批价及售价
     ISCHECK不强制控制指导价格 = Val(zlDatabase.GetPara(123, glngSys, 0)) = 1
End Function

Public Function ISCHECK外购扣前销售() As Boolean
    '功能:判断是否不强制要求控制批价及售价
    ISCHECK外购扣前销售 = Val(zlDatabase.GetPara(127, glngSys, 0)) = 1
End Function
 
Public Function Check普通科室() As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------
    '功能：验证当前人员是普通科室的相关人员
    '返回:是返回true,否则返回false
    '------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, bln向发料部门领用 As Boolean, strStock As String
    
    On Error GoTo ErrHandle
    bln向发料部门领用 = Val(zlDatabase.GetPara(132, glngSys, 0)) = 1

    If bln向发料部门领用 = False Then
        strStock = "K,V,12"
    Else
        strStock = "K,V,W,12"
    End If
    
    Check普通科室 = False
    gstrSQL = "" & _
        "   SELECT /*+ Rule*/ DISTINCT a.id, a.名称 " & _
        "   FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " & _
        "       , Table(cast(f_Str2List([3]) as zlTools.t_StrList)) D " & _
        "   Where c.工作性质 = b.名称 And (a.站点=[2] or a.站点 is null) " & _
        "       And b.编码=D.Column_value " & _
        "       AND a.id = c.部门id " & _
        "       AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01' " & _
        "       And a.ID IN (Select 部门ID From 部门人员 Where 人员ID=[1]) "
        
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取人员库房性质", UserInfo.Id, gstrNodeNo, strStock)
    If rsTemp.EOF Then
        Check普通科室 = True
    Else
        Check普通科室 = False
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get成本价(ByVal lng材料ID As Long, ByVal lng库房ID As Long, ByVal lng批次 As Long) As Double
'功能：获取当前药品的成本价格
'参数：药品id,库房id,批次
'返回值： 成本价格
    Dim rsData As ADODB.Recordset
    Dim blnNullPrice As Boolean
    
    On Error GoTo ErrHandle
    
    gstrSQL = "select 平均成本价 from 药品库存 where 药品id=[1] and 库房id=[2] and nvl(批次,0)=[3] and 性质=1"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "成本价", lng材料ID, lng库房ID, lng批次)
    
    If rsData.EOF Then
        blnNullPrice = True
    ElseIf IsNull(rsData!平均成本价) = True Then
        blnNullPrice = True
    ElseIf Val(rsData!平均成本价) < 0 Then
        blnNullPrice = True
    End If
    
    If Not blnNullPrice Then
        Get成本价 = rsData!平均成本价
    Else
        '如果无法从库存中取成本价，则从材料特性中取
        gstrSQL = "select 成本价 from 材料特性 where 材料id=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "成本价", lng材料ID)
        If Not rsData.EOF Then
            If Val(NVL(rsData!成本价, 0)) > 0 Then
                Get成本价 = rsData!成本价
            End If
        End If
    End If
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get售价(ByVal bln是否时价 As Boolean, lng材料ID As Long, ByVal lng库房ID As Long, ByVal lng批次 As Long) As Double
    '功能：获取原始的售价单位售价，主要用于出库
    '参数: bln是否时价:false-定价,true-时价
    '返回值：最小单位的价格
    Dim rsData As ADODB.Recordset
    Dim dbl零售价 As Double, dbl指导零售价 As Double, dbl差价让利比 As Double, dbl加成率 As Double
    Dim dbl成本价 As Double
    
    On Error GoTo ErrHandle

    '取定价药品售价
    If bln是否时价 = False Then
        gstrSQL = "Select a.现价 " & _
            " From 收费价目 A " & _
            " Where A.收费细目ID=[1] And Sysdate Between A.执行日期 And Nvl(A.终止日期,Sysdate) "
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "Get售价-取定价药品售价", lng材料ID)
        
        If Not rsData.EOF Then
            Get售价 = rsData!现价
        End If
    Else
        '取时价药品售价
        gstrSQL = "select Decode(零售价, Null, Decode(Nvl(实际数量, 0), 0, 0, 实际金额 / 实际数量), 零售价) as 零售价 " & _
            " from 药品库存 where 性质=1 and  药品id=[1] and 库房id=[2] and nvl(批次,0)=[3]"
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "GetOriPrice-零售价", lng材料ID, lng库房ID, lng批次)
        
        If rsData.EOF Then
            
            '无数据，从规格中取最近一次价格
            gstrSQL = "Select 上次售价,指导零售价,nvl(指导差价率,0) as 指导差价率,nvl(加成率,0) as 加成率,Nvl(差价让利比,100) 差价让利比 From 材料特性 Where 材料ID=[1]"
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "零售价", lng材料ID)
            
            If IsNull(rsData!上次售价) Then
                dbl指导零售价 = rsData!指导零售价
                dbl差价让利比 = rsData!差价让利比
                
                Get售价 = 0
                dbl成本价 = Get成本价(lng材料ID, lng库房ID, lng批次)
                dbl加成率 = rsData!加成率 / 100
                dbl零售价 = dbl成本价 * (1 + dbl加成率)
                dbl零售价 = dbl零售价 + (dbl指导零售价 - dbl零售价) * (1 - dbl差价让利比 / 100)
                Get售价 = IIf(dbl零售价 > dbl指导零售价, dbl指导零售价, dbl零售价)
            Else
                Get售价 = rsData!上次售价
            End If
        Else '库存有数据
            If rsData!零售价 < 0 Then
                gstrSQL = "Select 上次售价,指导零售价,nvl(指导差价率,0) as 指导差价率,nvl(加成率,0) as 加成率,Nvl(差价让利比,100) 差价让利比 From 材料特性 Where 材料ID=[1]"
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "零售价", lng材料ID)
                
                If IsNull(rsData!上次售价) Then
                    dbl指导零售价 = rsData!指导零售价
                    dbl差价让利比 = rsData!差价让利比
                    
                    Get售价 = 0
                    dbl成本价 = Get成本价(lng材料ID, lng库房ID, lng批次)
                    dbl加成率 = rsData!加成率 / 100
                    dbl零售价 = dbl成本价 * (1 + dbl加成率)
                    dbl零售价 = dbl零售价 + (dbl指导零售价 - dbl零售价) * (1 - dbl差价让利比 / 100)
                    Get售价 = IIf(dbl零售价 > dbl指导零售价, dbl指导零售价, dbl零售价)
                Else
                    Get售价 = rsData!上次售价
                End If
            Else
                Get售价 = rsData!零售价
            End If
        End If
    End If
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get零售价(ByVal lng材料ID As Long, ByVal lng库房ID As Long, ByVal lng批次 As Long, ByVal dbl比例系数 As Double) As Double
    '功能：获取时价药品当前药品的零售价
    '参数:药品id,库房id,批次
    '返回值：零售价
    Dim rsData As ADODB.Recordset
    Dim dbl零售价 As Double, dbl指导零售价 As Double, dbl差价让利比 As Double, dbl加成率 As Double
    Dim dbl成本价 As Double
    
    On Error GoTo ErrHandle
    If lng批次 <> 0 Then
        gstrSQL = "select 零售价 from 药品库存 where 药品id=[1] and 库房id=[2] and nvl(批次,0)=[3] and 性质=1"
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "零售价", lng材料ID, lng库房ID, lng批次)
    Else
        gstrSQL = "Select 实际金额 / 实际数量 As 零售价" & vbNewLine & _
                "   From 药品库存" & vbNewLine & _
                "   Where 库房id = [2] And 药品id = [1] And 性质 = 1 And 实际数量 > 0"
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "零售价", lng材料ID, lng库房ID)
    End If
    
    If rsData.EOF Or IsNull(rsData!零售价) = True Then
        '时价药品零售价计算公式:采购价*(1+加成率)
        '改为:采购价*(1+加成率)+(指导零售价-采购价*(1+加成率))*(1-差价让利比)
        '由于差价让利比的存在,以前所有按指导差价率计算的地方,均需要将差价率转换成加成率进行计算,此函数用于返回本次公式增加的部分金额：(指导零售价-采购价*(1+加成率))*(1-差价让利比)
        gstrSQL = "Select 上次售价,指导零售价,nvl(指导差价率,0) as 指导差价率,nvl(加成率,0) as 加成率,Nvl(差价让利比,100) 差价让利比 From 材料特性 Where 材料ID=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "零售价", lng材料ID)
        
        If IsNull(rsData!上次售价) Then
            dbl指导零售价 = rsData!指导零售价
            dbl差价让利比 = rsData!差价让利比
            
            Get零售价 = 0
            dbl成本价 = Get成本价(lng材料ID, lng库房ID, lng批次)
            dbl加成率 = rsData!加成率 / 100
            dbl零售价 = dbl成本价 * (1 + dbl加成率)
            dbl零售价 = dbl零售价 + (dbl指导零售价 - dbl零售价) * (1 - dbl差价让利比 / 100)
            Get零售价 = IIf(dbl零售价 > dbl指导零售价, dbl指导零售价, dbl零售价) * dbl比例系数
        Else
            Get零售价 = rsData!上次售价 * dbl比例系数
        End If
    Else
        Get零售价 = rsData!零售价 * dbl比例系数
    End If
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function ZVal(ByVal varValue As Variant, Optional ByVal blnForceNum As Boolean) As String
'功能：将0零转换为"NULL"串,在生成SQL语句时用
'参数：blnForceNum=当为Null时，是否强制表示为数字型
    ZVal = IIf(Val(varValue) = 0, IIf(blnForceNum, "-NULL", "NULL"), Val(varValue))
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
 
Public Function CheckPrint(ByVal strNo As String, ByVal int单据 As Integer, lng库房ID As Long, Optional lng对方部门ID As Long = 0, Optional bln对方库房 As Boolean = False) As Boolean
    '检查是否已经打印，并返回是否继续打印，继续则返回true，否则返回false
    Dim rsTemp As ADODB.Recordset
    
    If bln对方库房 = False Then
        gstrSQL = "Select 1 From 药品收发主表 Where NO = [1] And 单据 = [2] And 库房id = [3]"
    Else
        gstrSQL = "Select 1 From 药品收发主表 Where NO = [1] And 单据 = [2] And 对方部门id = [4]"
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查是否已经打印单据", strNo, int单据, lng库房ID, lng对方部门ID)
    If rsTemp.RecordCount > 0 Then
        If MsgBox("此单据已经打印过，是否继续打印？", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            CheckPrint = False
            Exit Function
        Else
            CheckPrint = True
        End If
    Else
        gstrSQL = "Zl_药品收发主表_Insert("
        gstrSQL = gstrSQL & "'" & strNo & "'"
        gstrSQL = gstrSQL & "," & int单据
        gstrSQL = gstrSQL & "," & lng库房ID
        gstrSQL = gstrSQL & "," & lng对方部门ID & ")"
        zlDatabase.ExecuteProcedure gstrSQL, "单据打印"
        CheckPrint = True
    End If
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
 
Public Sub CalcPosition(ByRef x As Single, ByRef y As Single, ByVal objBill As Object, Optional blnNoBill As Boolean = False)
    '----------------------------------------------------------------------
    '功能： 计算X,Y的实际坐标，并考虑屏幕超界的问题
    '参数： X---返回横坐标参数
    '       Y---返回纵坐标参数
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.hwnd, objPoint)
    If blnNoBill Then
        x = objPoint.x * 15 'objBill.Left +
        y = objPoint.y * 15 + objBill.Height '+ objBill.Top
    Else
        x = objPoint.x * 15 + objBill.CellLeft
        y = objPoint.y * 15 + objBill.CellTop + objBill.CellHeight
    End If
End Sub


Public Function ReturnParaData(ByVal lngSys As Long, ByVal str参数号IN As String) As ADODB.Recordset
    '-------------------------------------------------------------------------------------------
    '功能:获取指定的参数值,返回一个记录集
    '参数:lngSys-系统
    '     str参数号IN-参数号In,以逗号分离
    '
    '返回:参数记录集
    '编制:刘兴宏
    '日期:2007/12/17
    '-------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    strSQL = "" & _
        "   Select  /*+ Rule*/ 参数号,nvl(参数值,缺省值) as 参数值,参数说明 " & _
        "   From zlParameters A,Table(Cast(f_Num2List([2]) As zlTools.t_NumList)) B" & _
        "   where A.参数号 = B.Column_Value and a.系统=[1] and nvl(A.私有,0)=0 and nvl(a.模块,0)=0  " & _
        "   order by 参数号"
        
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取参数值", lngSys, str参数号IN)
    
    Set ReturnParaData = rsTemp
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
'取周，月，季，半年，年的第一天
Public Function GetFirstDate(ByVal intInteval As Integer, ByVal datCurrent As Date) As Date
    Dim datReturn As Date
    
    Select Case intInteval
        Case FirstDayOfWeek       '当前周的第一天
            datReturn = DateAdd("d", -Weekday(datCurrent) + 1, Now)
        Case FirstDayOfMonth       '当前月的第一天
            datReturn = DateAdd("d", -Day(datCurrent) + 1, datCurrent)
        Case FirstDayOfQuarter       '当前季的第一天
            Select Case DatePart("q", datCurrent)
                Case 1
                    datReturn = DateSerial(Year(datCurrent), 1, 1)
                    
                Case 2
                    datReturn = DateSerial(Year(datCurrent), 4, 1)
                Case 3
                    datReturn = DateSerial(Year(datCurrent), 7, 1)
                Case 4
                    datReturn = DateSerial(Year(datCurrent), 10, 1)
            End Select
        Case FirstDayOfHalfYear       '当前半年的第一天
            If Month(datCurrent) > 6 Then
                datReturn = DateSerial(Year(datCurrent), 7, 1)
            Else
                datReturn = DateSerial(Year(datCurrent), 1, 1)
            End If
        Case FirstDayOfyear       '当前年的第一天
            datReturn = DateSerial(Year(datCurrent), 1, 1)
    End Select
    GetFirstDate = datReturn
End Function



Public Function Check可用数量(ByVal lng库房ID As Long, ByVal lng材料ID As Long, ByVal lng批次 As Long, _
    ByVal dbl冲销数量 As Double, ByVal int库存检查 As Integer, Optional ByVal intType As Integer = 0) As Boolean
    '------------------------------------------------------------------------------
    '功能:检查入库冲销时的可用数量是否足够
    '返回:允足返回返回true,否则返回False
    '参数:
    '    int库存检查:0-不检查;1-检查，不足提醒,2-检查，不足禁止
    '编制:刘兴宏
    '日期:2008/02/15
    '------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, dbl数量 As Double
    
    err = 0: On Error GoTo ErrHand:
    '0-不检查
    If int库存检查 = 0 Then Check可用数量 = True: Exit Function
    
    gstrSQL = "Select A.可用数量,A.实际数量,B.编码,B.名称 From 药品库存 A,收费项目目录 B where A.药品id=B.id And A.药品id=[1] and A.库房id=[2] and nvl(A.批次,0)=[3] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查可用可存", lng材料ID, lng库房ID, lng批次)
    
    If rsTemp.EOF Then
        dbl数量 = 0
        gstrSQL = "Select 0 as 可用数量,B.编码,B.名称 From 收费项目目录 B where B.id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查可用可存", lng材料ID, lng库房ID, lng批次)
        If rsTemp.EOF Then ShowMsgBox "指定的卫生材料不存在,请检查!": Exit Function
    Else
        If intType = 0 Then
            dbl数量 = Round(Val(zlStr.NVL(rsTemp!可用数量, 0)), g_小数位数.obj_最大小数.数量小数)
        Else
            dbl数量 = Round(Val(zlStr.NVL(rsTemp!实际数量, 0)), g_小数位数.obj_最大小数.数量小数)
        End If
    End If
    
    If dbl数量 < Round(dbl冲销数量, g_小数位数.obj_最大小数.数量小数) Then
        If intType = 0 Then
            If int库存检查 = 1 Then
                '1-检查，不足提醒
                If MsgBox("“[" & rsTemp!编码 & "]" & zlStr.NVL(rsTemp!名称) & "”的可用库存不足，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            Else
                '2-检查，不足禁止
                ShowMsgBox "“[" & rsTemp!编码 & "]" & zlStr.NVL(rsTemp!名称) & "”的可用库存不足，不能继续！"
                Exit Function
            End If
        Else
            If int库存检查 = 1 Then
                '1-检查，不足提醒
                If MsgBox("“[" & rsTemp!编码 & "]" & zlStr.NVL(rsTemp!名称) & "”的实际库存不足，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            Else
                '2-检查，不足禁止
                ShowMsgBox "“[" & rsTemp!编码 & "]" & zlStr.NVL(rsTemp!名称) & "”的实际库存不足，不能继续！"
                Exit Function
            End If
        End If
    End If
    
    Check可用数量 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
Public Function 取单据批次(ByVal int单据 As Integer, _
    ByVal strNo As String, _
    lng材料ID As Long, int序号 As Integer, Optional lng入出系数 As Long = 1) As Long
    '------------------------------------------------------------------------------
    '功能:获取单据批次
    '返回:返回指定行的批次
    '编制:刘兴宏
    '日期:2008/02/15
    '------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    err = 0: On Error GoTo ErrHand:
    gstrSQL = "Select Nvl(批次, 0) 批次 From 药品收发记录 Where 单据 = [1] And NO = [2] And 序号 = [3] And 药品id = [4] And 入出系数 = [5]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取入库批次", int单据, strNo, int序号, lng材料ID, lng入出系数)
    If rsTemp.EOF Then
        取单据批次 = 0
    Else
        取单据批次 = Val(NVL(rsTemp!批次))
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog

End Function

 
  
Public Function SelectItem(ByVal frmMain As Form, ByVal objCtl As Control, ByVal strKey As String, _
    ByVal strTable As String, ByVal strTittle As String, Optional ByVal blnNotMsg As Boolean = False) As Boolean
    '------------------------------------------------------------------------------
    '功能:多功能选择器
    '参数:objCtl-文本框控件
    '     strKey-要搜索的值
    '     strTable-表名
    '     strTittle-选择器名称
    '     blnNotMsg-不提示.
    '返回:
    '编制:刘兴宏
    '日期:2008/02/18
    '------------------------------------------------------------------------------
    Dim blnCancel As Boolean, lngH As Long
    Dim vRect As RECT, sngX As Single, sngY As Single
    Dim rsTemp  As ADODB.Recordset
    'zlDatabase.ShowSelect
    '功能：多功能选择器
    '参数：
    '     frmParent=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    
    gstrSQL = "Select rownum as ID,a.* From " & strTable & " a"
    
    If strKey <> "" Then
        gstrSQL = gstrSQL & _
        "   Where ((名称) like [1] or  编码  like [1] or  简码  like  upper([1]))  " & _
        "    "
    End If
    gstrSQL = gstrSQL & _
    "   order by 编码"
    strKey = GetMatchingSting(strKey, False)
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Then
        Call CalcPosition(sngX, sngY, objCtl)
        lngH = objCtl.CellHeight
        sngY = sngY - lngH
    Else
        vRect = GetControlRect(objCtl.hwnd)
        lngH = objCtl.Height
        sngX = vRect.Left - 15
        sngY = vRect.Top
    End If
    Set rsTemp = zlDatabase.ShowSQLSelect(frmMain, gstrSQL, 0, strTittle, False, "", "", False, False, True, sngX, sngY, lngH, blnCancel, False, False, strKey)
    If blnCancel = True Then
        Call zlCtlSetFocus(objCtl, True)
        Exit Function
    End If
    
    If rsTemp Is Nothing Then
        If blnNotMsg = False Then
            ShowMsgBox "没有找到满足条件的内容,请检查!"
        End If
        Call zlCtlSetFocus(objCtl, True)
        Exit Function
    End If
    
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Then
        With objCtl
            .TextMatrix(.Row, .Col) = NVL(rsTemp!编码) & "-" & NVL(rsTemp!名称)
            .Cell(flexcpData, .Row, .Col) = NVL(rsTemp!名称)
        End With
    Else
        Call zlCtlSetFocus(objCtl, True)
        objCtl.Text = NVL(rsTemp!名称)
        objCtl.Tag = NVL(rsTemp!名称)
        zlCommFun.PressKey vbKeyTab
    End If
    SelectItem = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub zlCtlSetFocus(ByVal objCtl As Object, Optional blnDoEvnts As Boolean = False)
    '功能:将集点移动控件中:2008-07-08 16:48:35
    err = 0: On Error Resume Next
    If blnDoEvnts Then DoEvents
    If IsCtrlSetFocus(objCtl) = True Then: objCtl.SetFocus
End Sub

Public Function Select部门选择器(ByVal frmMain As Form, ByVal objCtl As Control, ByVal strSearch As String, _
    Optional str工作性质 As String = "", _
    Optional bln操作员 As Boolean = False, _
    Optional strSQL As String = "") As Boolean
    '------------------------------------------------------------------------------
    '功能:部门选择器
    '参数:objCtl-指定控件
    '     strSearch-要搜索的条件
    '     str工作性质-工作性质:如"V,W,K"
    '     bln操作员-是否加操作员限制
    '     strSQL-直接根据SQL获取数据(但部门表的别名一定要是A)
    '返回:成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2008/02/18
    '------------------------------------------------------------------------------
    Dim i As Long
    Dim blnCancel As Boolean, strKey As String, strTittle As String, lngH As Long, strFind As String
    Dim vRect As RECT
    Dim rsTemp  As ADODB.Recordset
    'zlDatabase.ShowSelect
    '功能：多功能选择器
    '参数：
    '     frmParent=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    
    strTittle = "部门选择器"
    vRect = GetControlRect(objCtl.hwnd)
    lngH = objCtl.Height
    
    strKey = GetMatchingSting(strSearch, False)
    
    If strSQL <> "" Then
    
        gstrSQL = strSQL
    Else
'        gstrSQL = "" & _
'        "   Select /*+ Rule*/ distinct a.Id,a.上级id,a.编码,a.名称,a.简码,a.位置 ,To_Char(a.建档时间, 'yyyy-mm-dd') As 建档时间, " & _
'        "          decode(To_Char(a.撤档时间, 'yyyy-mm-dd'),'3000-01-01','',To_Char(a.撤档时间, 'yyyy-mm-dd')) 撤档时间"
    
        If str工作性质 = "" And bln操作员 = False Then
            gstrSQL = "" & _
            "   Select a.Id,a.上级id,a.编码,a.名称,a.简码,a.位置 ,To_Char(a.建档时间, 'yyyy-mm-dd') As 建档时间, " & _
            "          decode(To_Char(a.撤档时间, 'yyyy-mm-dd'),'3000-01-01','',To_Char(a.撤档时间, 'yyyy-mm-dd')) 撤档时间"
        
            gstrSQL = gstrSQL & vbCrLf & _
            "   From 部门表 a" & _
            "   Where 1=1"
        Else
            gstrSQL = "" & _
            "   Select /*+ Rule*/ distinct a.Id,a.上级id,a.编码,a.名称,a.简码,a.位置 ,To_Char(a.建档时间, 'yyyy-mm-dd') As 建档时间, " & _
            "          decode(To_Char(a.撤档时间, 'yyyy-mm-dd'),'3000-01-01','',To_Char(a.撤档时间, 'yyyy-mm-dd')) 撤档时间"
        
            gstrSQL = gstrSQL & vbCrLf & _
            "   From 部门表 a, 部门性质分类 b,部门性质说明 c," & _
            IIf(str工作性质 = "", "", "       (Select Column_Value From Table(Cast(f_Str2list([2]) As zlTools.t_Strlist))) J") & _
            "   Where c.工作性质 = b.名称" & IIf(str工作性质 = "", "(+)", " and B.编码=J.column_value ") & _
            "         AND a.id = c.部门id " & _
            IIf(bln操作员 = False, "", " And a.ID IN (Select 部门ID From 部门人员 Where 人员ID=[1])")
        End If
        gstrSQL = gstrSQL & vbCrLf & _
            "   and  (a.撤档时间>=to_date('3000-01-01','yyyy-mm-dd') or a.撤档时间 is null ) And (a.站点=[4] or a.站点 is null) "
    End If
    
    strFind = ""
    If strSearch <> "" Then
        strFind = "   and  (a.编码 like upper([3]) or a.简码 like upper([3]) or a.名称 like [3] )"
        If IsNumeric(strSearch) Then                         '如果是数字,则只取编码
            If Mid(gSystem_Para.Para_输入方式, 1, 1) = "1" Then strFind = " And (A.编码 Like Upper([3]))"
        ElseIf zlCommFun.IsCharAlpha(strSearch) Then           '01,11.输入全是字母时只匹配简码
            '0-拼音码,1-五笔码,2-两者
            '.int简码方式 = Val(zlDatabase.GetPara("简码方式" ))
            If Mid(gSystem_Para.Para_输入方式, 2, 1) = "1" Then strFind = " And  (a.简码 Like Upper([3]))"
        ElseIf zlCommFun.IsCharChinese(strSearch) Then  '全汉字
            strFind = " And a.名称 Like [3] "
        End If
    End If
    
    If strSearch = "" And str工作性质 = "" And bln操作员 = False And strSQL = "" Then
        gstrSQL = gstrSQL & _
        "   Start With A.上级id Is Null Connect By Prior A.ID = A.上级id "
    Else
        gstrSQL = gstrSQL & vbCrLf & strFind & vbCrLf & " Order by A.编码"
    End If
    
    If strSearch = "" And str工作性质 = "" And bln操作员 = False And strSQL = "" Then
        '分上下级
        Set rsTemp = zlDatabase.ShowSQLSelect(frmMain, gstrSQL, 1, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, UserInfo.Id, str工作性质, strKey, gstrNodeNo)
    Else
        Set rsTemp = zlDatabase.ShowSQLSelect(frmMain, gstrSQL, 0, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, UserInfo.Id, str工作性质, strKey, gstrNodeNo)
    End If
    If blnCancel = True Then
        Call zlCtlSetFocus(objCtl, True)
        Exit Function
    End If
    If rsTemp Is Nothing Then
        ShowMsgBox "没有满足条件的部门,请检查!"
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    Call zlCtlSetFocus(objCtl, True)
    If UCase(TypeName(objCtl)) = UCase("ComboBox") Then
        blnCancel = True
        For i = 0 To objCtl.ListCount - 1
            If objCtl.ItemData(i) = Val(rsTemp!Id) Then
                objCtl.Text = objCtl.List(i)
                objCtl.ListIndex = i
                blnCancel = False
                Exit For
            End If
        Next
        If blnCancel Then
            ShowMsgBox "你选择的部门在下拉列表中不存在,请检查!"
            If objCtl.Enabled Then objCtl.SetFocus
            Exit Function
        End If
    Else
        objCtl.Text = NVL(rsTemp!编码) & "-" & NVL(rsTemp!名称)
        objCtl.Tag = Val(rsTemp!Id)
    End If
    zlCommFun.PressKey vbKeyTab
    Select部门选择器 = True
End Function
  
Public Function zlDblIsValid(ByVal StrInput As String, ByVal intMax As Integer, Optional bln负数检查 As Boolean = True, Optional bln零检查 As Boolean = True, _
        Optional ByVal hwnd As Long = 0, Optional str项目 As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:检查字符串是否合法的金额
    '入参:strInput        输入的字符串
    '     intMax          整数的位数
    '     bln负数检查     是否进行负数检查
    '     bln零检查         是否进行零的检查
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-10-20 15:16:08
    '-----------------------------------------------------------------------------------------------------------
   
    Dim dblValue As Double
    If bln零检查 = True Then
        If StrInput = "" Then
            ShowMsgBox str项目 & "未输入，请检查!"
            If hwnd <> 0 Then SetFocusHwnd hwnd
            Exit Function
        End If
    End If
    If StrInput = "" Then zlDblIsValid = True: Exit Function
    
    If IsNumeric(StrInput) = False Then
        MsgBox str项目 & "不是有效的数字格式。", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '设置焦点
        Exit Function
    End If
    
    dblValue = Val(StrInput)
    If dblValue >= 10 ^ intMax - 1 Then
        MsgBox str项目 & "数值过大，不能超过" & 10 ^ intMax - 1 & "。", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '设置焦点
        Exit Function
    End If
    If bln负数检查 = True And dblValue < 0 Then
        MsgBox str项目 & "不能输入负数。", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '设置焦点
        Exit Function
    End If
    
    If Abs(dblValue) >= 10 ^ intMax And dblValue < 0 Then
        MsgBox str项目 & "数值过小，不能小于-" & 10 ^ intMax - 1 & "位。", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '设置焦点
        Exit Function
    End If
    
    
    If bln零检查 = True And dblValue = 0 Then
        MsgBox str项目 & "不能输入零。", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '设置焦点
        Exit Function
    End If
    zlDblIsValid = True
End Function

Public Function zlCheckIsDate(ByVal strKey As String, ByVal strTittle As String) As String
    '------------------------------------------------------------------------------
    '功能:检查是否合法的日期型,可以为:(20070101或2007-01-01)或则(01-01或0101)或则(01<01-31>)
    '参数:strKey-需要检查的关建字
    '返回:合法的日期,返回标准格式(yyyy-mm-dd),否则返回""
    '编制:刘兴宏
    '日期:2008/01/24
    '------------------------------------------------------------------------------
    If Len(strKey) = 4 And InStr(1, strKey, "-") = 0 Then
        '0101,需要再前面加年
        strKey = Year(Now) & strKey
    ElseIf Len(Replace(strKey, "-", "")) = 4 And InStr(1, strKey, "-") > 0 Then
        '01-01形式,需要补零
        strKey = Year(Now) & Replace(strKey, "-", "")
    ElseIf Len(strKey) <= 2 And IsNumeric(strKey) Then
        '指是日
        strKey = Format(Now, "YYYYMM") & IIf(Len(strKey) = 2, strKey, "0" & strKey)
    End If
    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
        strKey = TranNumToDate(strKey)
        If strKey = "" Then
            ShowMsgBox strTittle & "必须为日期型,请检查！"
            Exit Function
        End If
    End If
    If Not IsDate(strKey) Then
        ShowMsgBox strTittle & "必须为日期型如(2000-10-10) 或（20001010）,请检查！"
        Exit Function
    End If
    zlCheckIsDate = strKey
End Function

Public Function zl存在未审核单据(ByVal lng材料ID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:检查是否存在未审核的单据
    '入参:
    '出参:
    '返回:存在,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-11-07 15:33:14
    '-----------------------------------------------------------------------------------------------------------

    '检查药品是否存在未审核单据
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "Select 1 From 药品收发记录 Where 药品id = [1] And Rownum = 1 And 审核日期 Is Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查卫生材料是否存在未审核单据", lng材料ID)
    zl存在未审核单据 = rsTemp.RecordCount <> 0
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function Select供应商(ByVal frmMain As Form, ByVal objCtl As Control, Optional ByVal strSearch As String = "") As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:供应商选择
    '入参:frmMain-调用的主窗体
    '    objCtl-调用的控件
    '    strSearch-搜索条件(""表示所有选择)
    '出参:
    '返回:选择成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-11-10 10:38:26
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As Recordset, strKey As String
    Dim blnCancel As Boolean, lngH As Long
    Dim vRect As RECT, sngX As Single, sngY As Single
    Dim bytStyle As Byte, bln末级 As Boolean
    
    
    strKey = GetMatchingSting(strSearch, False)
    
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Then
        Call CalcPosition(sngX, sngY, objCtl)
        lngH = objCtl.CellHeight
        sngY = sngY - lngH
    Else
        vRect = GetControlRect(objCtl.hwnd)
        lngH = objCtl.Height
        sngX = vRect.Left - 15
        sngY = vRect.Top
    End If
    
    
 
    '     frmParent=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    gstrSQL = "" & _
        "   Select id,上级ID,编码, 名称, 简码, 末级, 许可证号, 许可证效期, 执照号, 执照效期, 税务登记号, 地址, 电话, 开户银行," & _
        "           帐号, 联系人, 类型, 信用期, 信用额, 销售委托人, to_char(销售委托日期,'yyyy-mm-dd') as 销售委托日期, 质量认证号, to_char(质量认证日期,'yyyy-mm-dd') as 质量认证日期," & _
        "           药监局备案号, to_char(药监局备案日期,'yyyy-mm-dd') as 药监局备案日期, 授权号, 授权期, 站点," & _
        "           to_char(建档时间,'yyyy-mm-dd') as 建档时间, decode(To_Char(撤档时间,'yyyy-MM-dd'),'3000-01-01','', to_char(撤档时间,'yyyy-mm-dd')) as 撤档时间" & _
        "   From 供应商 " & _
        "   Where  (To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' or 撤档时间 is null)  "
    If strSearch = "" Then
        gstrSQL = gstrSQL & _
            "           And (substr(类型,5,1)=1 And (站点=[2] or 站点 is null) Or Nvl(末级,0)=0) " & _
            "   Start with 上级ID is null connect by prior ID =上级ID " & _
            "   Order by level,ID"
        bln末级 = True
        bytStyle = 2
    Else
        gstrSQL = gstrSQL & _
            "    And (站点=[2] or 站点 is null) And 末级=1 And substr(类型,5,1)=1 " & _
            "    And (简码 like upper([1]) Or 编码 like [1] or 名称 like [1]) "
        bytStyle = 0
        bln末级 = False
    End If
    Set rsTemp = zlDatabase.ShowSQLSelect(frmMain, gstrSQL, bytStyle, "供应商选择器", Not bln末级, "", "请选择符合卫生材料的供应商", False, True, Not bln末级, sngX, sngY, lngH, blnCancel, False, False, strKey, gstrNodeNo)
        
    If blnCancel Then
        Call zlCtlSetFocus(objCtl, True)
        Exit Function
    End If
    If rsTemp Is Nothing Then
        ShowMsgBox "没有找到满足条件的供应商,请检查!"
        Call zlCtlSetFocus(objCtl, True)
        Exit Function
    End If
    
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Then
        With objCtl
            .TextMatrix(.Row, .Col) = NVL(rsTemp!编码) & "-" & NVL(rsTemp!名称)
            .Cell(flexcpData, .Row, .Col) = NVL(rsTemp!Id)
        End With
    Else
        Call zlCtlSetFocus(objCtl, True)
        objCtl.Text = NVL(rsTemp!名称)
        objCtl.Tag = NVL(rsTemp!Id)
        zlCommFun.PressKey vbKeyTab
    End If
    Select供应商 = True
End Function

'按编码，名称，别名查找某一列
Public Function FindVsRowNew(ByVal vsBill As VSFlexGrid, ByVal int比较列 As Integer, _
    ByVal str比较值 As String, ByVal blnFirst As Boolean) As Boolean
    Dim intStartRow As Integer
    Dim intRow As Integer
    Dim strSpell As String
    Dim strCode As String
    Dim rsCode As New Recordset
    
    On Error GoTo ErrHandle
    FindVsRowNew = True
    With vsBill
        If .Rows = 2 Then Exit Function
        If str比较值 = "" Then Exit Function
        If blnFirst = True Then
            intStartRow = 0
        Else
            intStartRow = .Row
        End If
        If intStartRow = .Rows - 1 Then
            intStartRow = 1
        Else
            intStartRow = intStartRow + 1
        End If
        For intRow = intStartRow To .Rows - 1
            If .TextMatrix(intRow, int比较列) <> "" Then
                strCode = .TextMatrix(intRow, int比较列)
                If InStr(1, UCase(strCode), UCase(str比较值)) <> 0 Then
                    .SetFocus
                    .Row = intRow
                    .Col = int比较列
                    .TopRow = .Row
                    Exit Function
                End If
            End If
        Next
        
        gstrSQL = "" & _
        " SELECT DISTINCT b.编码 " & _
        " FROM (    SELECT DISTINCT A.收费细目id " & _
        "           FROM 收费项目别名 A" & _
        "           Where A.简码 LIKE upper([1]) " & _
        "       ) a, 收费项目目录 B " & _
        " Where a.收费细目id = b.ID And (b.站点=[2] or b.站点 is null) "
        
        Set rsCode = zlDatabase.OpenSQLRecord(gstrSQL, "查找指定卫生材料", GetMatchingSting(str比较值, False), gstrNodeNo)
        If rsCode.EOF Then
            FindVsRowNew = False
            Exit Function
        End If
        
        For intRow = intStartRow To .Rows - 1
            If .TextMatrix(intRow, int比较列) <> "" Then
                strCode = .TextMatrix(intRow, int比较列)
                rsCode.MoveFirst
                Do While Not rsCode.EOF
                    If InStr(1, UCase(strCode), UCase(rsCode!编码)) <> 0 Then
                        .SetFocus
                        .Row = intRow
                        .Col = int比较列
                        .TopRow = .Row
                        rsCode.Close
                        Exit Function
                    End If
                    rsCode.MoveNext
                Loop
            End If
        Next
        rsCode.Close
    End With
    FindVsRowNew = False
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

 
Public Function SelectAndNotAddItem(ByVal frmMain As Form, ByVal objCtl As Control, ByVal strKey As String, _
    ByVal strTable As String, ByVal strTittle As String, Optional blnOnlyName As Boolean = False, _
    Optional bln未找到增加 As Boolean = False, Optional strOra过程 As String, Optional strWhere As String) As Boolean
    '------------------------------------------------------------------------------
    '功能:多功能选择器
    '参数:objCtl-文本框控件
    '     strKey-要搜索的值
    '     strTable-表名
    '     strTittle-选择器名称
    '返回:
    '编制:刘兴宏
    '日期:2008/02/18
    '------------------------------------------------------------------------------
    Dim blnCancel As Boolean, lngH As Long, str编码 As String, str名称 As String
    Dim vRect As RECT, sngX As Single, sngY As Single
    Dim rsTemp  As ADODB.Recordset
    'zlDatabase.ShowSelect
    '功能：多功能选择器
    '参数：
    '     frmParent=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    str名称 = strKey
    
    gstrSQL = "Select rownum as ID,a.* From " & strTable & " a where 1=1 "
    
    If strKey <> "" Then
        gstrSQL = gstrSQL & _
        "   And ((名称) like [1] or  编码  like [1] or  简码  like  upper([1]))  " & _
        "    "
    End If
    gstrSQL = gstrSQL & strWhere & _
    "   order by 编码"
    strKey = GetMatchingSting(strKey, False)
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Or UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
        If UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
            Call CalcPosition(sngX, sngY, objCtl.MsfObj)
            lngH = objCtl.MsfObj.CellHeight
        Else
            Call CalcPosition(sngX, sngY, objCtl)
            lngH = objCtl.CellHeight
        End If
        sngY = sngY - lngH
    Else
        vRect = GetControlRect(objCtl.hwnd)
        lngH = objCtl.Height
        sngX = vRect.Left - 15
        sngY = vRect.Top
    End If
    
    Set rsTemp = zlDatabase.ShowSQLSelect(frmMain, gstrSQL, 0, strTittle, False, "", "", False, False, True, sngX, sngY, lngH, blnCancel, False, False, strKey)
    If blnCancel = True Then
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    
    If rsTemp Is Nothing Then
        
        If bln未找到增加 Then
            If zlCommFun.IsCharChinese(str名称) = False Then GoTo NOAdd::
            If MsgBox("注意:" & vbCrLf & _
                   "     未找到相关的" & strTable & ",是否增加“" & str名称 & "”？", vbQuestion + vbYesNo + vbDefaultButton2, strTable) = vbNo Then
                If objCtl.Enabled Then objCtl.SetFocus
                Exit Function
            End If
            
            If AutoAddBaseItem(strTable, str编码, str名称, strTable & "增加", False) = False Then
                If objCtl.Enabled Then objCtl.SetFocus
                Exit Function
            End If
            
            If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Or UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
                With objCtl
                    .TextMatrix(.Row, .Col) = IIf(blnOnlyName, str名称, str编码 & "-" & str名称)
                    If Not (UCase(TypeName(objCtl)) = UCase("BILLEDIT")) Then
                        .Cell(flexcpData, .Row, .Col) = IIf(blnOnlyName, str名称, str编码 & "-" & str名称)
                    End If
                End With
            Else
                If IsCtrlSetFocus(objCtl) Then objCtl.SetFocus
                objCtl.Text = IIf(blnOnlyName, str名称, str编码 & "-" & str名称)
                objCtl.Tag = str名称
                zlCommFun.PressKey vbKeyTab
            End If
            SelectAndNotAddItem = True
            Exit Function
        Else
NOAdd:
            ShowMsgBox "没有找到满足条件的" & strTable & ",请检查!"
            If objCtl.Enabled Then objCtl.SetFocus
            Exit Function
        End If
    End If
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Or UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
        With objCtl
            .TextMatrix(.Row, .Col) = IIf(blnOnlyName, NVL(rsTemp!名称), NVL(rsTemp!编码) & "-" & NVL(rsTemp!名称))
            If Not (UCase(TypeName(objCtl)) = UCase("BILLEDIT")) Then
                .EditText = NVL(rsTemp!名称)
                .Cell(flexcpData, .Row, .Col) = NVL(rsTemp!名称)
            End If
        End With
    Else
        If IsCtrlSetFocus(objCtl) Then objCtl.SetFocus
        objCtl.Text = NVL(rsTemp!名称)
        objCtl.Tag = NVL(rsTemp!名称)
        zlCommFun.PressKey vbKeyTab
    End If
    SelectAndNotAddItem = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function




Public Function AutoAddBaseItem(ByVal strTable As String, str编码 As String, str名称 As String, _
    Optional strTittle As String = "增加项目", Optional blnMsg As Boolean = False) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:自动增加项目信息(只针对有编码,名称的信息增加(只增加：编码和名称,简码)
    '--入参数:
    '--出参数:
    '--返  回:增加成功,返回true,否则返回false
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New Recordset
    Dim int编码 As Integer, strCode As String, strSpecify As String
    AutoAddBaseItem = False
    If blnMsg = True Then
        If MsgBox("没有找到你输入的" & strTable & "，你要把它加入" & strTable & "中吗？", vbYesNo + vbQuestion, strTittle) = vbNo Then
            Exit Function
        End If
    End If
    
    err = 0: On Error GoTo ErrHand:
    
    gstrSQL = "SELECT Nvl(MAX(LENGTH(编码)), 2) As Length FROM  " & strTable
    zlDatabase.OpenRecordset rsTemp, gstrSQL, strTittle
    
    int编码 = rsTemp!Length
    
    gstrSQL = "SELECT Nvl(MAX(LPAD(编码," & int编码 & ",'0')),'00') As Code FROM  " & strTable
    zlDatabase.OpenRecordset rsTemp, gstrSQL, strTittle
    strCode = rsTemp!Code
    
    int编码 = Len(strCode)
    strCode = strCode + 1
    
    If int编码 >= Len(strCode) Then
    strCode = String(int编码 - Len(strCode), "0") & strCode
    End If
    strSpecify = zlCommFun.SpellCode(str名称)
    
    
    gstrSQL = "ZL_" & strTable & "_INSERT('" & strCode & "','" & str名称 & "','" & strSpecify & "')"
    zlDatabase.ExecuteProcedure gstrSQL, strTittle
    str编码 = strCode
    AutoAddBaseItem = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub Logogram(ByVal staVal As StatusBar, ByVal bytType As Byte)
'简码方式
'staVal: StartusBar控件
'bytType: 0=拼音; 1=五笔;  当前简码状态
    Dim i As Integer
    For i = 1 To staVal.Panels.Count
        If staVal.Panels(i).Key = "PY" And staVal.Panels(i).Visible = True Then
            If bytType = 0 Then
                staVal.Panels(i).Bevel = sbrInset
                zlDatabase.SetPara "简码方式", 0
                gSystem_Para.int简码方式 = 0
            ElseIf bytType = 1 Then
                staVal.Panels(i).Bevel = sbrRaised
            End If
        ElseIf staVal.Panels(i).Key = "WB" And staVal.Panels(i).Visible = True Then
            If bytType = 0 Then
                staVal.Panels(i).Bevel = sbrRaised
            ElseIf bytType = 1 Then
                staVal.Panels(i).Bevel = sbrInset
                zlDatabase.SetPara "简码方式", 1
                gSystem_Para.int简码方式 = 1
            End If
        End If
    Next
End Sub


Public Function CheckQualifications(ByVal lngModule As Long, ByVal intType As Integer, ByVal StrInput As String) As Boolean
    '校验卫材，生产商，供应商信息和资质效期
    'intType：0－卫材；1－生产商；2－供应商
    'strInput：字符串时为名称；数字时为ID
    Dim rsTmp As ADODB.Recordset
    Dim strMsgInfo As String
    Dim strMsgDate As String
    Dim dateCurrent As Date
    Dim strMsg As String
    
    Dim intCheckType As Integer
    Dim arrColumn
    Dim strCheck As String
    Dim strCheck_卫材 As String
    Dim strCheck_生产商 As String
    Dim strCheck_供应商 As String
    Dim n As Integer
    Dim strTmp As String
    
    On Error GoTo ErrHandle
    If StrInput = "" Then
        CheckQualifications = True
        Exit Function
    End If
        
    '资质校验项目和方式的保存格式：校验方式|类别1,项目1,是否校验;类别1,项目2,是否校验;类别2,项目1,是否校验;类别2,项目2....
    strCheck = zlDatabase.GetPara("资质校验", glngSys, lngModule, "")
    
    '保存的参数格式不正确时退出
    If InStr(1, strCheck, "|") = 0 Then
        CheckQualifications = True
        Exit Function
    End If

    '取校验方式：0-不检查；1－提醒；2－禁止
    intCheckType = Val(Mid(strCheck, 1, InStr(1, strCheck, "|") - 1))
    
    '不检查时退出
    If intCheckType = 0 Then
        CheckQualifications = True
        Exit Function
    End If

    '取校验内容：
    strCheck = Mid(strCheck, InStr(1, strCheck, "|") + 1)

    If strCheck = "" Then
        CheckQualifications = True
        Exit Function
    End If

    '分别取卫材，生产商，供应商需要校验的内容
    strCheck = strCheck & ";"
    arrColumn = Split(strCheck, ";")
    For n = 0 To UBound(arrColumn)
        If arrColumn(n) <> "" Then
            If Split(arrColumn(n), ",")(0) = "卫材" And Split(arrColumn(n), ",")(2) = 1 Then
                strCheck_卫材 = IIf(strCheck_卫材 = "", "", strCheck_卫材 & ";") & Split(arrColumn(n), ",")(1)
            End If

            If Split(arrColumn(n), ",")(0) = "卫材生产商" And Split(arrColumn(n), ",")(2) = 1 Then
                strCheck_生产商 = IIf(strCheck_生产商 = "", "", strCheck_生产商 & ";") & Split(arrColumn(n), ",")(1)
            End If

            If Split(arrColumn(n), ",")(0) = "卫材供应商" And Split(arrColumn(n), ",")(2) = 1 Then
                strCheck_供应商 = IIf(strCheck_供应商 = "", "", strCheck_供应商 & ";") & Split(arrColumn(n), ",")(1)
            End If
        End If
    Next
    
    '无校验内容时退出
    If (intType = 0 And strCheck_卫材 = "") Or (intType = 1 And strCheck_生产商 = "") Or (intType = 2 And strCheck_供应商 = "") Then
        CheckQualifications = True
        Exit Function
    End If
    
    dateCurrent = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd"))
    
    '卫材
    If intType = 0 Then
        gstrSQL = "Select ('[' || B.编码 || ']' || B.名称) AS 卫材信息, A.许可证号, A.许可证有效期,注册证号,注册证有效期 " & _
            " From 收费项目目录 B,材料特性 A " & _
            " Where B.ID = A.材料ID And A.材料ID = [1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "校验卫材资质", Val(StrInput))
        
        If Not rsTmp.EOF Then
            If NVL(rsTmp!许可证号) = "" And InStr(strCheck_卫材, "许可证号") > 0 Then
                strTmp = rsTmp!卫材信息 & "：" & "无许可证号"
            End If
            
            If NVL(rsTmp!许可证有效期) <> "" Then
                If DateDiff("d", rsTmp!许可证有效期, dateCurrent) > 0 And InStr(strCheck_卫材, "许可证有效期") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!卫材信息 & "：", strTmp & ",") & "许可证已过期"
                End If
            End If
        End If
        If NVL(rsTmp!注册证号) = "" And InStr(strCheck_卫材, "注册证号") > 0 Then
            strTmp = rsTmp!卫材信息 & "：" & "无注册证号"
        End If
        
        If NVL(rsTmp!注册证有效期) <> "" Then
            If DateDiff("d", rsTmp!注册证有效期, dateCurrent) > 0 And InStr(strCheck_卫材, "注册证有效期") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!卫材信息 & "：", strTmp & ",") & "注册证已过期"
            End If
        End If
    End If
    
    '生产商
    If intType = 1 Then
        gstrSQL = "Select ('[' || A.编码 || ']' || A.名称) AS 生产商, A.生产企业许可证, A.生产企业许可证效期,a.经营许可证,a.经营许可证效期,a.企业法人执照,a.企业法人执照效期 " & _
                        " From 材料生产商 A " & _
                        " Where A.名称 = [1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "校验卫材资质", StrInput)
        
        If Not rsTmp.EOF Then
            If NVL(rsTmp!生产企业许可证) = "" And InStr(strCheck_生产商 & ";", "生产企业许可证" & ";") > 0 Then
                strTmp = rsTmp!生产商 & "：" & "无生产企业许可证"
            End If
            
            If NVL(rsTmp!生产企业许可证效期) <> "" Then
                If DateDiff("d", rsTmp!生产企业许可证效期, dateCurrent) > 0 And InStr(strCheck_生产商 & ";", "生产企业许可证效期" & ";") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!生产商 & "：", strTmp & ",") & "生产企业许可证已过期"
                End If
            End If
        End If
        If Not rsTmp.EOF Then
            If NVL(rsTmp!经营许可证) = "" And InStr(strCheck_生产商 & ";", "经营许可证" & ";") > 0 Then
                strTmp = rsTmp!生产商 & "：" & "无经营许可证"
            End If
            
            If NVL(rsTmp!经营许可证效期) <> "" Then
                If DateDiff("d", rsTmp!生产企业许可证效期, dateCurrent) > 0 And InStr(strCheck_生产商 & ";", "经营许可证效期" & ";") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!生产商 & "：", strTmp & ",") & "经营许可证已过期"
                End If
            End If
        End If
        If Not rsTmp.EOF Then
            If NVL(rsTmp!企业法人执照) = "" And InStr(strCheck_生产商 & ";", "企业法人执照" & ";") > 0 Then
                strTmp = rsTmp!生产商 & "：" & "无企业法人执照"
            End If
            
            If NVL(rsTmp!企业法人执照效期) <> "" Then
                If DateDiff("d", rsTmp!生产企业许可证效期, dateCurrent) > 0 And InStr(strCheck_生产商 & ";", "企业法人执照效期" & ";") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!生产商 & "：", strTmp & ",") & "企业法人执照已过期"
                End If
            End If
        End If
    End If
    
    '供应商
    If intType = 2 Then
        gstrSQL = "Select ('[' || 编码 || ']' || 名称) AS 供应商, 税务登记号, 许可证号, 执照号, 授权号, 质量认证号, 质量认证日期, 药监局备案号, 药监局备案日期, 许可证效期, 执照效期, 授权期 " & _
            " From 供应商 " & _
            " Where (撤档时间 Is Null Or 撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And ID = [1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "供应商信息", Val(StrInput))
        
        strTmp = ""
        
        If Not rsTmp.EOF Then
            If NVL(rsTmp!税务登记号) = "" And InStr(strCheck_供应商, "税务登记号") > 0 Then
                strTmp = rsTmp!供应商 & "：" & "无税务登记号"
            End If
            
            If NVL(rsTmp!许可证号) = "" And InStr(strCheck_供应商, "许可证号") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!供应商 & "：", strTmp & ",") & "无许可证号"
            End If
            
            If NVL(rsTmp!执照号) = "" And InStr(strCheck_供应商, "执照号") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!供应商 & "：", strTmp & ",") & "无执照号"
            End If
            
            If NVL(rsTmp!授权号) = "" And InStr(strCheck_供应商, "授权号") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!供应商 & "：", strTmp & ",") & "无授权号"
            End If
            
            If NVL(rsTmp!质量认证号) = "" And InStr(strCheck_供应商, "质量认证号") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!供应商 & "：", strTmp & ",") & "无质量认证号"
            End If
            
            If NVL(rsTmp!质量认证日期) <> "" Then
                If DateDiff("d", rsTmp!质量认证日期, dateCurrent) > 0 And InStr(strCheck_供应商, "质量认证日期") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!供应商 & "：", strTmp & ",") & "质量认证号已过期"
                End If
            End If
            
            If NVL(rsTmp!药监局备案号) = "" And InStr(strCheck_供应商, "药监局备案号") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!供应商 & "：", strTmp & ",") & "无药监局备案号"
            End If
            
            If NVL(rsTmp!药监局备案日期) <> "" Then
                If DateDiff("d", rsTmp!药监局备案日期, dateCurrent) > 0 And InStr(strCheck_供应商, "药监局备案日期") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!供应商 & "：", strTmp & ",") & "药监局备案号已过期"
                End If
            End If
            
            If NVL(rsTmp!许可证效期) <> "" Then
                If DateDiff("d", rsTmp!许可证效期, dateCurrent) > 0 And InStr(strCheck_供应商, "许可证效期") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!供应商 & "：", strTmp & ",") & "许可证已过期"
                End If
            End If
            
            If NVL(rsTmp!执照效期) <> "" Then
                If DateDiff("d", rsTmp!执照效期, dateCurrent) > 0 And InStr(strCheck_供应商, "执照效期") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!供应商 & "：", strTmp & ",") & "执照已过期"
                End If
            End If
            
            If NVL(rsTmp!授权期) <> "" Then
                If DateDiff("d", rsTmp!执照效期, dateCurrent) > 0 And InStr(strCheck_供应商, "授权期") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!供应商 & "：", strTmp & ",") & "授权已过期"
                End If
            End If
        End If
    End If
    
    '提示或禁止
    If strTmp <> "" Then
        If intCheckType = 1 Then
            If MsgBox("未通过资质校验，是否继续？" & vbCrLf & strTmp, vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
                CheckQualifications = True
                Exit Function
            Else
                Exit Function
            End If
        ElseIf intCheckType = 2 Then
            MsgBox "未通过资质校验，不能入库！" & vbCrLf & strTmp, vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
    CheckQualifications = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub VsfGridColFormat(ByVal objGrid As VSFlexGrid, ByVal intCol As Integer, ByVal strColName As String, _
    ByVal lngColWidth As Long, ByVal intColAlignment As Integer, _
    Optional ByVal strColKey As String = "", Optional ByVal intFixedColAlignment As Integer = 4)
    'vsf列设置：列名，列宽，列对齐方式，固定列对齐方式（默认为居中对齐）
    
    With objGrid
        .TextMatrix(0, intCol) = strColName
        .ColWidth(intCol) = lngColWidth
        .ColAlignment(intCol) = intColAlignment
'        .ColData(intCol) = lngColWidth
        
        .ColKey(intCol) = strColKey
        .FixedAlignment(intCol) = intFixedColAlignment
    End With
End Sub


'取系统参数值
Public Sub GetSysParms()
    Dim rs As New ADODB.Recordset
    Dim n As Integer
    Dim strMsg As String
    
    On Error GoTo errH
    
    '取卫材最大允许精度
    gstrSQL = "Select 零售金额, 成本价, 零售价, 实际数量 From 药品收发记录 Where Rownum < 1"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "取药品精度")
    gtype_UserDrugDigits.Digit_金额 = rs.Fields(0).NumericScale
    gtype_UserDrugDigits.Digit_成本价 = rs.Fields(1).NumericScale
    gtype_UserDrugDigits.Digit_零售价 = rs.Fields(2).NumericScale
    gtype_UserDrugDigits.Digit_数量 = rs.Fields(3).NumericScale
    
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Public Function StuffWork_GetCheckStockRule(ByVal lng库房ID As Long) As Integer
    '取出库检查规则
    Dim rsData As ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = " Select Nvl(检查方式,0) 库存检查 From 材料出库检查 Where 库房ID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "取出库检查规则", lng库房ID)

    If Not rsData.EOF Then
        StuffWork_GetCheckStockRule = rsData!库存检查
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function Get分批属性(ByVal lng库房ID As Long, ByVal lng材料ID As Long) As Integer
    '返回指定库房，指定材料的分批属性
    '返回：0-不分批，1-分批
    Dim rsCheck As New ADODB.Recordset
    Dim int分批 As Integer
    Dim bln发料部门 As Boolean
    Dim strSQL As String
        
    On Error GoTo ErrHandle
    
    '判断是否是发料部门
    strSQL = "select 部门ID from 部门性质说明 where (工作性质 =  '发料部门' or 工作性质 =  '制剂室') And 部门id=[1]"
    Set rsCheck = zlDatabase.OpenSQLRecord(strSQL, "Get分批属性", lng库房ID)

    bln发料部门 = (Not rsCheck.EOF)
        
    '判断对应的药品目录中的分批属性
    strSQL = " Select Nvl(库房分批,0) As 库房分批,nvl(在用分批,0) As 在用分批 " & _
              " From 材料特性 Where 材料ID=[1]"
    Set rsCheck = zlDatabase.OpenSQLRecord(strSQL, "Get分批属性", lng材料ID)
              
    If bln发料部门 Then
        int分批 = rsCheck!在用分批
    Else
        int分批 = rsCheck!库房分批
    End If
    
    Get分批属性 = int分批
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function CheckNoStock(ByVal lng库房ID As Long, ByVal lng材料ID As Long, Optional ByVal lng批次 As Long = -1) As Boolean
    '检查是否无库存
    '返回：true-无库存;false-有库存
    Dim rsData As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    gstrSQL = "Select 1 From 药品库存 " & _
        " Where 性质 = 1 And 库房id = [1] And 药品id = [2] And (Nvl(实际数量, 0) <> 0 Or Nvl(实际金额, 0) <> 0 Or Nvl(实际差价, 0) <> 0) "
    
    If lng批次 <> -1 Then
        gstrSQL = gstrSQL & " And Nvl(批次,0) = [3] "
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "CheckNoStock", lng库房ID, lng材料ID, lng批次)
    
    CheckNoStock = rsData.EOF
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckNumStock(ByVal objVSF As Object, ByVal lng库房ID As Long, ByVal lntCol材料id As Integer, _
    ByVal intCol批次 As Integer, ByVal intCol数量 As Integer, ByVal intCol比例系数 As Integer, _
    ByVal intMethod As Integer, Optional int入出业务 As Integer, Optional ByVal int精度 As Integer, _
    Optional ByVal intType As Integer = 0) As String
    '功能：审核出库类单据时，整张单据检查库存表实际(可用)数量是否足够
    '参数：objVSF-需要检查的表格;lng库房id；intcol批次-批次所在列；intCol数量-数量所在列；intCol比例系数-比例系数所在列
    '参数：intMethod，1-正常审核，2-冲销，3-退库审核
    '参数：int入出业务，0-入库；1-出库
    '参数：intType，0-实际数量；1-可用数量
    '返回值：哪行具体的材料名称，为空-检查通过，数量充足；不为空-检查未通过，数量不充足
    Dim objCol As Collection         '已使用的数量集合
    Dim dblNum As Double
    Dim varNum As Variant
    Dim varTemp As Variant
    Dim strTemp As String
    Dim lng材料ID As Long
    Dim lng批次 As Long
    Dim rsData As ADODB.Recordset
    Dim strKey As String
    Dim vardrug As Variant
    Dim lngRow As Long
    Dim strArray As String
    Dim dbl比例系数 As Double
    Dim intSum As Integer '计数
    
    On Error GoTo ErrHandle
    
    '先组合表格中数量，组合数量主要是考虑不分批的情况
    Set objCol = New Collection
    With objVSF
        If .Rows < 2 Then Exit Function
        For lngRow = 1 To .Rows - 1
            dblNum = 0
            If .TextMatrix(lngRow, lntCol材料id) <> "" Then
                For Each vardrug In objCol
                    If vardrug(0) = .TextMatrix(lngRow, lntCol材料id) & "," & Val(.TextMatrix(lngRow, intCol批次)) & "," & Val(.TextMatrix(lngRow, intCol比例系数)) Then
                        dblNum = vardrug(1)
                        objCol.Remove vardrug(0)
                        Exit For
                    End If
                Next
                strKey = .TextMatrix(lngRow, lntCol材料id) & "," & Val(.TextMatrix(lngRow, intCol批次)) & "," & Val(.TextMatrix(lngRow, intCol比例系数))
                '以最小单位保存数量，方便审核时数量与库存数据比较
                strArray = dblNum + (Val(.TextMatrix(lngRow, intCol数量)))
                objCol.Add Array(strKey, strArray), strKey
            End If
        Next
    End With
    
    For Each varNum In objCol
        strTemp = varNum(0)  '格式是材料id,批次,比例系数
        dblNum = varNum(1)
        varTemp = Split(strTemp, ",")
        If int入出业务 = 0 Then '入库
            If intMethod = 1 Then '正常审核
                If dblNum < 0 Then
                    '负数入库，需要减库存，所以需要判断库存是否充足
                    dblNum = Abs(dblNum)
                Else
                    '正数入库，不见库存，所以不检查
                    dblNum = 0
                End If
            ElseIf intMethod = 2 Then
                '冲销
                If dblNum < 0 Then
                    dblNum = 0
                Else
                    dblNum = dblNum
                End If
            ElseIf intMethod = 3 Then
                '退库审核，退库必须录入正数
                dblNum = dblNum
            End If
        Else    '出库
            If intMethod = 1 Then '正常审核
                If dblNum < 0 Then
                    '负数出库，不减库存，所以不检查
                    dblNum = 0
                Else
                    '正常入库，需要减库存，所以需要判断库存是否充足
                    dblNum = dblNum
                End If
            ElseIf intMethod = 2 Then
                '冲销
                If dblNum < 0 Then
                    dblNum = Abs(dblNum)
                Else
                    dblNum = 0
                End If
            End If
        End If
        
        '只有有数量才判断
        If dblNum > 0 Then
            lng材料ID = varTemp(0)
            lng批次 = varTemp(1)
            dbl比例系数 = varTemp(2)

            If Get分批属性(lng库房ID, lng材料ID) = 0 Then
                lng批次 = 0
            End If
            
            gstrSQL = "Select a.可用数量,a.实际数量, '[' || b.编码 || ']' || b.名称 名称" & vbNewLine & _
                        "From 药品库存 A, 收费项目目录 B" & vbNewLine & _
                        "Where a.药品id = b.Id And a.药品id = [2] And a.库房id = [3] And Nvl(a.批次, 0) = [4] And b.类别 = '4' And a.性质 = 1"
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "库存检查", dblNum, lng材料ID, lng库房ID, lng批次)
            If rsData.RecordCount = 0 Then '无库存记录
                gstrSQL = "select '[' || 编码 || ']' || 名称 名称 from 收费项目目录 where id=[1]"
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "库存检查", lng材料ID)
                
                intSum = intSum + 1
                If intSum <= 3 Then CheckNumStock = IIf(CheckNumStock = "", "", CheckNumStock & "、" & vbCrLf) & rsData!名称
            Else '有库存记录
                If intType = 0 Then '比较实际数量
                    If zlStr.FormatEx(rsData!实际数量 / dbl比例系数, int精度, , False) >= dblNum Then
                    Else
                        intSum = intSum + 1
                        If intSum <= 3 Then CheckNumStock = IIf(CheckNumStock = "", "", CheckNumStock & "、" & vbCrLf) & rsData!名称
                    End If
                Else '比较可用数量
                    If zlStr.FormatEx(rsData!可用数量 / dbl比例系数, int精度, , False) >= dblNum Then
                    Else
                        intSum = intSum + 1
                        If intSum <= 3 Then CheckNumStock = IIf(CheckNumStock = "", "", CheckNumStock & "、" & vbCrLf) & rsData!名称
                    End If
                End If
            End If
'            Next
        End If
    Next
    CheckNumStock = CheckNumStock & IIf(intSum > 3, "等" & intSum & "个", "")
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

