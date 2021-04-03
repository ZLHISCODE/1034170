Attribute VB_Name = "mdlPlugIn"
Option Explicit

Public gcnOracle As ADODB.Connection


'���µ����д���������֧�ֶ���ͬʱ�ҽ�
'CkeckUseable ������������ʹ�ã���д��չ���ʱֱ�Ӵ���ǰ ��λ���Ƽ��ɡ�ʹ�ø÷���ʱҪ���� zl9ComLib.dll
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' ע���ؼ��ְ�ȫѡ��...
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

' ע�����������...
Public Enum ValueType
    REG_SZ = 1                         ' �ַ���ֵ
    REG_EXPAND_SZ = 2                  ' �������ַ���ֵ
    REG_BINARY = 3                     ' ������ֵ
    REG_DWORD = 4                      ' DWORDֵ
    REG_MULTI_SZ = 7                   ' ���ַ���ֵ
End Enum

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Private marrName As Variant
Private mcolPlugIn As Collection

Private Sub GetPathNames()
'���ܣ���ȡע���CLSID�¼�Ŀ¼

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
'���ܣ�����Ѵ��ڵ�ע���ؼ��ֵ�ֵ
'������ValueName="" �򷵻� KeyName ���Ĭ��ֵ
'      ���ָ����ע���ؼ��ֲ�����, �򷵻ؿմ�
'      KeyRoot--������, KeyName--��������, ValueName--ֵ������, ValueType--ֵ������
    Dim i As Integer
    Dim hKey As Long
    Dim TempValue As String                             ' ע���ؼ��ֵ���ʱֵ
    Dim Value As String                                 ' ע���ؼ��ֵ�ֵ
    Dim ValueSize As Long                               ' ע���ؼ��ֵ�ֵ��ʵ�ʳ���
    TempValue = Space(1024)                             ' �洢ע���ؼ��ֵ���ʱֵ�Ļ�����
    ValueSize = 1024                                    ' ����ע���ؼ��ֵ�ֵ��Ĭ�ϳ���
    
    ' ��һ���Ѵ��ڵ�ע���ؼ���...
    RegOpenKeyEx KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey
    
    ' ����Ѵ򿪵�ע���ؼ��ֵ�ֵ...
    RegQueryValueEx hKey, ValueName, 0, ValueType, ByVal TempValue, ValueSize
    
    ' ����ע���ؼ��ֵĵ�ֵ...
    Select Case ValueType                                                        ' ͨ���жϹؼ��ֵ�����, ���д���
        Case REG_SZ, REG_MULTI_SZ, REG_EXPAND_SZ
            TempValue = Left$(TempValue, ValueSize - 1)                          ' ȥ��TempValueβ���ո�
            Value = TempValue
        Case REG_DWORD
            ReDim dValue(3) As Byte
            RegQueryValueEx hKey, ValueName, 0, REG_DWORD, dValue(0), ValueSize
            For i = 3 To 0 Step -1
                Value = Value + String(2 - Len(Hex(dValue(i))), "0") + Hex(dValue(i))   ' ���ɳ���Ϊ8��ʮ�������ַ���
            Next i
            If CDbl("&H" & Value) < 0 Then                                              ' ��ʮ�����Ƶ� Value ת��Ϊʮ����
                Value = 2 ^ 32 + CDbl("&H" & Value)
            Else
                Value = CDbl("&H" & Value)
            End If
        Case REG_BINARY
            If ValueSize > 0 Then
                ReDim bValue(ValueSize - 1) As Byte                                     ' �洢 REG_BINARY ֵ����ʱ����
                RegQueryValueEx hKey, ValueName, 0, REG_BINARY, bValue(0), ValueSize
                For i = 0 To ValueSize - 1
                    Value = Value + String(2 - Len(Hex(bValue(i))), "0") + Hex(bValue(i)) + " "  ' ������ת�����ַ���
                Next i
            End If
    End Select
    
    ' �ر�ע���ؼ���...
    RegCloseKey hKey
    GetKeyValue = Trim(Value)                                                    ' ���غ���ֵ
End Function

Private Function GetAllPlugIns() As String
'���ܣ���ȡ��չ����Ĳ������ƣ����Ÿ
    Dim strTmp As String
    Dim strName As String
    Dim strResult As String
    Dim i As Integer
    
    Call GetPathNames
    
    For i = 1 To UBound(marrName)
        strResult = GetKeyValue(HKEY_CLASSES_ROOT, CStr(marrName(i)), strTmp, REG_SZ)
        '��ZLPLUGIN��ͷ
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
        Optional ByVal int���� As Integer = -1, Optional ByVal strReserve As String, Optional ByRef strFuncName As String, Optional ByVal lngPatiID As Long, _
        Optional ByVal varRecId As Variant, Optional ByVal varKeyId As Variant)
'���ܣ���չ�������֧����ش���
'������bytType �������� 1=��ʼ����2=��ȡ��������3=ִ�й��ܣ�4=��ֹ����bytType=2ʱ strFunName��Ϊ����
'      cnOracle=�����
'      lngSys,lngModual=��ǰ���ýӿڵ��ϼ�ϵͳ�ż�ģ���
'      int����  ���ó���:0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����(PACS/LIS)
'      strReserve=��������,������չʹ��
'      strFunName ���κ���� ��bytType=2ʱ���Σ���bytType=3ʱ���
'      lngPatiID=��ǰ����ID
'      varRecId=���ֻ����ַ����������ﲡ�ˣ�Ϊ��ǰ�Һŵ��Ż��߹Һ�ID����סԺ���ˣ�Ϊ��ǰסԺ��ҳID
'      varKeyId=���ֻ����ַ�������ǰ�Ĺؼ�ҵ������Ψһ��ʶID����ҽ��ID
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
                Call objTmp.Initialize(cnOracle, lngSys, lngModual, int����)
                '����ʹ�����ƣ��û�����ʱ��ʾ������
                strUserName = objTmp.GetUserName 'ҽԺ�û�--��λ����
                
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
            strTmp = objTmp.GetFuncNames(lngSys, lngModual, int����, strReserve)
            strFuncNameTmp = IIf(strFuncNameTmp = "", "", strFuncNameTmp & ",") & strTmp
        Next i
        strFuncName = strFuncNameTmp
    ElseIf bytType = 3 Then
        For i = 1 To mcolPlugIn.Count
            Set objTmp = mcolPlugIn.Item(i)
            Call objTmp.ExecuteFunc(lngSys, lngModual, strFuncName, lngPatiID, varRecId, varKeyId, strReserve, int����)
        Next i
    ElseIf bytType = 4 Then
        For i = 1 To mcolPlugIn.Count
            Set objTmp = mcolPlugIn.Item(i)
            Call objTmp.Terminate(lngSys, lngModual, int����)
        Next i
    End If
    Err.Clear: On Error GoTo 0
End Function

Private Function CkeckUseable(ByVal str��λ���� As String) As Boolean
'���ܣ���չ���ʹ������ʾ������
'������ʹ�õ�λ��ȫ��
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errH
    strSQL = "Select Item,Text From Table(Cast(zltools.f_Reg_Info([1]) As zlTools.t_Reg_Rowset))"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "ZlplugIn", 0)
    If rsTmp.EOF Then Exit Function
    
    rsTmp.Filter = "Item='��λ����'"
    If rsTmp.RecordCount <> 1 Then Exit Function
    
    If InStr("," & rsTmp!Text & ",", "," & str��λ���� & ",") > 0 Then CkeckUseable = True
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
