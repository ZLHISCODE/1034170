Attribute VB_Name = "mdlPlugIn"
Option Explicit

Public gcnOracle As ADODB.Connection


'以下的所有代码是用于支持多插件同时挂接
'CkeckUseable 方法用于限制使用，编写扩展插件时直接传当前 单位名称即可。使用该方法时要引用 zl9ComLib.dll
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' 注册表关键字安全选项...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_READ = KEY_QUERY_VALUE + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + READ_CONTROL
Const KEY_WRITE = KEY_SET_VALUE + KEY_CREATE_SUB_KEY + READ_CONTROL
Const KEY_EXECUTE = KEY_READ
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
Const ERROR_NO_MORE_ITEMS = 259&
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_LOCAL_MACHINE = &H80000002

' 注册表数据类型...
Public Enum ValueType
    REG_SZ = 1                         ' 字符串值
    REG_EXPAND_SZ = 2                  ' 可扩充字符串值
    REG_BINARY = 3                     ' 二进制值
    REG_DWORD = 4                      ' DWORD值
    REG_MULTI_SZ = 7                   ' 多字符串值
End Enum

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Private marrName As Variant
Private mcolPlugIn As Collection

Private Sub GetPathNames()
'功能：获取注册表CLSID下级目录

    Dim hKey As Long, Cnt As Long, sName As String, sData As String, Ret As Long, RetData As Long
    Const BUFFER_SIZE As Long = 255
    marrName = Array()
    Ret = BUFFER_SIZE
    If RegOpenKey(HKEY_CLASSES_ROOT, "CLSID", hKey) = 0 Then
        sName = Space(BUFFER_SIZE)
        While RegEnumKeyEx(hKey, Cnt, sName, Ret, ByVal 0&, vbNullString, ByVal 0&, ByVal 0&) <> ERROR_NO_MORE_ITEMS
            ReDim Preserve marrName(UBound(marrName) + 1)
            marrName(UBound(marrName)) = "CLSID\" & Left$(sName, Ret)
            Cnt = Cnt + 1
            sName = Space(BUFFER_SIZE)
            Ret = BUFFER_SIZE
        Wend
        RegCloseKey hKey
    End If
    Cnt = 0
End Sub

Private Function GetKeyValue(KeyRoot As Long, KeyName As String, ValueName As String, Optional ValueType As Long) As String
'功能：获得已存在的注册表关键字的值
'参数：ValueName="" 则返回 KeyName 项的默认值
'      如果指定的注册表关键字不存在, 则返回空串
'      KeyRoot--根类型, KeyName--子项名称, ValueName--值项名称, ValueType--值项类型
    Dim i As Integer
    Dim hKey As Long
    Dim TempValue As String                             ' 注册表关键字的临时值
    Dim Value As String                                 ' 注册表关键字的值
    Dim ValueSize As Long                               ' 注册表关键字的值的实际长度
    TempValue = Space(1024)                             ' 存储注册表关键字的临时值的缓冲区
    ValueSize = 1024                                    ' 设置注册表关键字的值的默认长度
    
    ' 打开一个已存在的注册表关键字...
    RegOpenKeyEx KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey
    
    ' 获得已打开的注册表关键字的值...
    RegQueryValueEx hKey, ValueName, 0, ValueType, ByVal TempValue, ValueSize
    
    ' 返回注册表关键字的的值...
    Select Case ValueType                                                        ' 通过判断关键字的类型, 进行处理
        Case REG_SZ, REG_MULTI_SZ, REG_EXPAND_SZ
            TempValue = Left$(TempValue, ValueSize - 1)                          ' 去掉TempValue尾部空格
            Value = TempValue
        Case REG_DWORD
            ReDim dValue(3) As Byte
            RegQueryValueEx hKey, ValueName, 0, REG_DWORD, dValue(0), ValueSize
            For i = 3 To 0 Step -1
                Value = Value + String(2 - Len(Hex(dValue(i))), "0") + Hex(dValue(i))   ' 生成长度为8的十六进制字符串
            Next i
            If CDbl("&H" & Value) < 0 Then                                              ' 将十六进制的 Value 转换为十进制
                Value = 2 ^ 32 + CDbl("&H" & Value)
            Else
                Value = CDbl("&H" & Value)
            End If
        Case REG_BINARY
            If ValueSize > 0 Then
                ReDim bValue(ValueSize - 1) As Byte                                     ' 存储 REG_BINARY 值的临时数组
                RegQueryValueEx hKey, ValueName, 0, REG_BINARY, bValue(0), ValueSize
                For i = 0 To ValueSize - 1
                    Value = Value + String(2 - Len(Hex(bValue(i))), "0") + Hex(bValue(i)) + " "  ' 将数组转换成字符串
                Next i
            End If
    End Select
    
    ' 关闭注册表关键字...
    RegCloseKey hKey
    GetKeyValue = Trim(Value)                                                    ' 返回函数值
End Function

Private Function GetAllPlugIns() As String
'功能：获取扩展插件的部件名称，逗号割。
    Dim strTmp As String
    Dim strName As String
    Dim strResult As String
    Dim i As Integer
    
    Call GetPathNames
    
    For i = 1 To UBound(marrName)
        strResult = GetKeyValue(HKEY_CLASSES_ROOT, CStr(marrName(i)), strTmp, REG_SZ)
        '以ZLPLUGIN开头
        If UCase(Left(strResult, 8)) = "ZLPLUGIN" Then
            If InStr(strResult, ".") > 0 Then
                If Len(Split(strResult, ".")(0)) > 8 And InStr(strName, Split(strResult, ".")(0)) = 0 Then
                    strName = IIf(strName = "", "", strName & ",") & Split(strResult, ".")(0)
                End If
            End If
        End If
    Next
    GetAllPlugIns = strName
End Function

Public Function HandlePlugIn(ByVal bytType As Byte, ByVal lngSys As Long, ByVal lngModual As Long, Optional ByVal cnOracle As ADODB.Connection, _
        Optional ByVal int场合 As Integer = -1, Optional ByVal strReserve As String, Optional ByRef strFuncName As String, Optional ByVal lngPatiID As Long, _
        Optional ByVal varRecId As Variant, Optional ByVal varKeyId As Variant)
'功能：扩展插件功能支持相关处理
'参数：bytType 操作类型 1=初始化，2=获取功能名，3=执行功能，4=终止。当bytType=2时 strFunName作为出参
'      cnOracle=活动连接
'      lngSys,lngModual=当前调用接口的上级系统号及模块号
'      int场合  调用场合:0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS)
'      strReserve=保留参数,用于扩展使用
'      strFunName 出参和入参 当bytType=2时出参，当bytType=3时入参
'      lngPatiID=当前病人ID
'      varRecId=数字或者字符串；对门诊病人，为当前挂号单号或者挂号ID；对住院病人，为当前住院主页ID
'      varKeyId=数字或者字符串；当前的关键业务数据唯一标识ID，如医嘱ID
    Dim strTmp As String
    Dim strFuncNameTmp As String
    Dim strUserName As String
    Dim objTmp As Object
    Dim varArr As Variant
    Dim i As Integer
    
    On Error Resume Next
    
    If bytType = 1 Then
        strTmp = GetAllPlugIns
        If strTmp = "" Then Exit Function
        varArr = Split(strTmp, ",")
        Set mcolPlugIn = New Collection
        For i = 0 To UBound(varArr)
            Set objTmp = CreateObject(varArr(i) & ".clsPlugIn")
            If Not objTmp Is Nothing Then
                Call objTmp.Initialize(cnOracle, lngSys, lngModual, int场合)
                '部件使用限制，用户名空时表示不限制
                strUserName = objTmp.GetUserName '医院用户--单位名称
                
                If strUserName <> "" Then
                    If CkeckUseable(strUserName) Then
                        mcolPlugIn.Add objTmp, "_" & varArr(i)
                    End If
                Else
                    mcolPlugIn.Add objTmp, "_" & varArr(i)
                End If
            End If
            Set objTmp = Nothing
        Next i
    End If
    
    If mcolPlugIn Is Nothing Then Exit Function
    
    If bytType = 2 Then
        For i = 1 To mcolPlugIn.Count
            Set objTmp = mcolPlugIn.Item(i)
            strTmp = ""
            strTmp = objTmp.GetFuncNames(lngSys, lngModual, int场合, strReserve)
            strFuncNameTmp = IIf(strFuncNameTmp = "", "", strFuncNameTmp & ",") & strTmp
        Next i
        strFuncName = strFuncNameTmp
    ElseIf bytType = 3 Then
        For i = 1 To mcolPlugIn.Count
            Set objTmp = mcolPlugIn.Item(i)
            Call objTmp.ExecuteFunc(lngSys, lngModual, strFuncName, lngPatiID, varRecId, varKeyId, strReserve, int场合)
        Next i
    ElseIf bytType = 4 Then
        For i = 1 To mcolPlugIn.Count
            Set objTmp = mcolPlugIn.Item(i)
            Call objTmp.Terminate(lngSys, lngModual, int场合)
        Next i
    End If
    Err.Clear: On Error GoTo 0
End Function

Private Function CkeckUseable(ByVal str单位名称 As String) As Boolean
'功能：扩展插件使用限制示例代码
'参数：使用单位的全名
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errH
    strSQL = "Select Item,Text From Table(Cast(zltools.f_Reg_Info([1]) As zlTools.t_Reg_Rowset))"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ZlplugIn", 0)
    If rsTmp.EOF Then Exit Function
    
    rsTmp.Filter = "Item='单位名称'"
    If rsTmp.RecordCount <> 1 Then Exit Function
    
    If InStr("," & rsTmp!Text & ",", "," & str单位名称 & ",") > 0 Then CkeckUseable = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
