Attribute VB_Name = "mdlCISBase"
Option Explicit
Public gcnOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gstrPrivs As String                   '��ǰ�û����еĵ�ǰģ��Ĺ���
Public gstrProductName As String
Public gstrSysName As String                'ϵͳ����
Public glngModul As Long
Public glngSys As Long
Public gblnCancel As Boolean                '��¼�����е�ȡ����ť�Ƿ񱻵����

Public gstrDBOwner As String                '��ǰϵͳ������
Public gstrDBUser As String                 '��ǰ���ݿ��û�
Public glngUserId As Long                   '��ǰ�û�id
Public gstrUserCode As String               '��ǰ�û�����
Public gstrUserName As String               '��ǰ�û�����
Public gstrUserAbbr As String               '��ǰ�û�����

Public glngDeptId As Long                   '��ǰ�û�����id
Public gstrDeptCode As String               '��ǰ�û����ű���
Public gstrDeptName As String               '��ǰ�û���������
Public gstrItemName As String

Public gstrUnitName As String               '�û���λ����
Public gfrmMain As Object


Public glngPreHWnd As Long '����֧�������ֹ���

Public gstrSql As String
Public gstrMatch As String                  '���ݱ��ز�����ƥ��ģʽ��ȷ������ƥ�����
Public gblnOK As Boolean

Public gobjKernel As New clsCISKernel       '�ٴ����Ĳ���
Public gobjLogisticPlatform As Object       '����ƽ̨�ӿ�

Public gobjRIS As Object                    '����RIS�ӿڶ���
Public Enum RISBaseItemOper                 '����RIS�������ݲ������ͣ�1-������2-�޸ģ�3-ɾ��
    AddNew = 1
    Modify = 2
    Delete = 3
End Enum
Public Enum RISBaseItemType                 '����RIS�����������ͣ�1��������ĿĿ¼��2��������Ŀ��λ
    ClinicItem = 1
    ClinicItemPart = 2
End Enum

Public gblnKSSStrict As Boolean             '�Ƿ����ÿ���ҩ���ϸ����
Public gblnIncomeItem As Boolean            '��¼������Ŀ�Ƿ�����

Public Type type_user_Digits
    dig_�ɱ��� As Double
    dig_���ۼ� As Double
    dig_���� As Double
    dig_��� As Double
End Type
Public gtype_MaxDigits As type_user_Digits  '������¼��󾫶�

Public Type TYPE_USER_INFO
    ID As Long
    ����ID As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
    ��ҩ���� As Long
End Type
Public UserInfo As TYPE_USER_INFO
Public Const gstrLisHelp As String = "zl9LisWork"               'LIS���ð���ʱʹ�õĲ�����
Public glngTXTProc As Long '����Ĭ�ϵ���Ϣ�����ĵ�ַ
Public Const WM_CONTEXTMENU = &H7B ' ���һ��ı���ʱ������������Ϣ
Public Const GCST_INVALIDCHAR = "'"             '�����������Ч�ַ�

'֧�ֻ��ֵĳ���
Public Const WM_MOUSEWHEEL = &H20A


Public Const GWL_STYLE = (-16)
Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetClientRect Lib "User32" (ByVal hWnd As Long, lpRect As RECT) As Long

'˽�С�����ģ�����
Public Enum ����_ҩƷĿ¼����_����
    P1_����ҩ������Ŀ = 1
    P2_�г�ҩ������Ŀ = 2
    P3_�в�ҩ������Ŀ = 3
    P4_Ӧ�÷�Χ = 4
    P5_ʱ��ҩƷ�����ε��� = 5
End Enum

Public Sub IniRIS(Optional ByVal blnMsg As Boolean)
'���ܣ���ʼ�������ӿڲ���
'������blnMsg������ʧ��ʱ�Ƿ���ʾ
    If gobjRIS Is Nothing Then
        On Error Resume Next
        Set gobjRIS = CreateObject("zl9XWInterface.clsHISInner")
        Err.Clear: On Error GoTo 0
    End If
    If gobjRIS Is Nothing Then
        If blnMsg Then
            MsgBox "RIS�ӿڲ���(zl9XWInterface)δ�����ɹ���", vbInformation, gstrSysName
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
    
    '��ȡע������������
    strCurrent = Format(Now, "yyyyMMddHHmm")
    intAtom = GetSetting("ZLSOFT", "����ȫ��", "����", 0)
    Call SaveSetting("ZLSOFT", "����ȫ��", "����", 0)
    blnValid = (intAtom <> 0)
    
    '������ڣ���Դ����н���
    If blnValid Then
        Call GlobalGetAtomName(intAtom, strBuffer, 255)
        strSource = Trim(Replace(strBuffer, Chr(0), ""))
        '���Ϊ�գ����ʾ�Ƿ�
        If strSource <> "" Then
            If Left(strSource, 1) <> "#" Then
                strSource = TranPasswd(Mid(strSource, 1, 12))
                If strSource <> strCurrent Then '�ж�ʱ�����Ƿ����1
                    If CStr(Mid(strSource, 11, 2) + 1) = CStr(Mid(strCurrent, 11, 2) + 0) Then
                        '�����ȣ���ͨ��
                    Else
                        '���ȣ���ʾ���ڽ�λ�����Ӧ��Ϊ��
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
        MsgBox "The component is lapse��", vbInformation, gstrSysName
        Exit Function
    End If
    CheckValid = True
End Function

Public Function GetFormat(ByVal dblInput As Double, ByVal intDotBit As Integer) As String
    'ȡ��ֵ��С��λ��
    GetFormat = Format(dblInput, "#0." & String(intDotBit, "0"))
End Function

Public Function TranPasswd(strOld As String) As String
    '------------------------------------------------
    '���ܣ� ����ת������
    '������
    '   strOld��ԭ����
    '���أ� �������ɵ�����
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
    '����ȡҩƷ�ĸ�����󾫶�
    Dim rsTemp As ADODB.Recordset
    On Error GoTo ErrHand
    
    gstrSql = "Select ���۽��, �ɱ���, ���ۼ�, ʵ������ From ҩƷ�շ���¼ Where Rownum < 1"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "��󾫶�")
    If rsTemp.RecordCount = 0 Then
        gtype_MaxDigits.dig_�ɱ��� = 7
        gtype_MaxDigits.dig_��� = 2
        gtype_MaxDigits.dig_���ۼ� = 7
        gtype_MaxDigits.dig_���� = 7
    Else
        gtype_MaxDigits.dig_�ɱ��� = rsTemp.Fields(1).NumericScale
        gtype_MaxDigits.dig_��� = rsTemp.Fields(0).NumericScale
        gtype_MaxDigits.dig_���ۼ� = rsTemp.Fields(2).NumericScale
        gtype_MaxDigits.dig_���� = rsTemp.Fields(3).NumericScale
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

'ȡҩƷ���۸��������С��λ��
Public Function GetDigit(ByVal int��� As Integer, ByVal int���� As Integer, Optional ByVal int��λ As Integer) As Integer
    'int���1-ҩƷ;2-����
    'int���ݣ�1-�ɱ���;2-���ۼ�;3-����;4-���
    'int��λ�������ȡ���λ�������Բ�����ò���
    '         ҩƷ��λ:1-�ۼ�;2-����;3-סԺ;4-ҩ��;
    '         ���ĵ�λ:1-ɢװ;2-��װ
    '���أ���С2�����Ϊ���ݿ����С��λ��
    
    Dim rsTmp As ADODB.Recordset
    Dim intMax��� As Integer
    Dim intMax�ɱ��� As Integer
    Dim intMax���ۼ� As Integer
    Dim intMax���� As Integer
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    gstrSql = "Select ���۽��, �ɱ���, ���ۼ�, ʵ������ From ҩƷ�շ���¼ Where Rownum < 1"
    Set rs = zlDatabase.OpenSQLRecord(gstrSql, "ȡҩƷ����")
    
    intMax��� = rs.Fields(0).NumericScale
    intMax�ɱ��� = rs.Fields(1).NumericScale
    intMax���ۼ� = rs.Fields(2).NumericScale
    intMax���� = rs.Fields(3).NumericScale
    
    gstrSql = "Select Nvl(����, 0) ���� From ҩƷ���ľ��� Where ��� = [1] And ���� = [2] And ��λ = [3] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, "ȡҩƷ" & Choose(int����, "�ɱ���", "���ۼ�", "����") & "С��λ��", int���, int����, int��λ)
    
    If rsTmp.RecordCount > 0 Then
        GetDigit = rsTmp!����
    End If
    
    If GetDigit = 0 Then
        '���û�����þ��ȣ���ȡ���ݿ���������λ��
        GetDigit = Choose(int����, intMax�ɱ���, intMax���ۼ�, intMax����)
    End If
    
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    GetDigit = Choose(int����, intMax�ɱ���, intMax���ۼ�, intMax����, intMax���)
End Function


Public Function GetUserInfo() As Boolean
    '���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As New ADODB.Recordset
    
    Set rsTmp = zlDatabase.GetUserInfo
    
    UserInfo.�û��� = gstrDBUser
    UserInfo.���� = gstrDBUser
    If Not rsTmp.EOF Then
        UserInfo.ID = rsTmp!ID
        UserInfo.��� = rsTmp!���
        UserInfo.����ID = IIf(IsNull(rsTmp!����ID), 0, rsTmp!����ID)
        UserInfo.���� = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        UserInfo.���� = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        gstrUserName = UserInfo.����
        GetUserInfo = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function MoveSpecialChar(ByVal strInputString As String, Optional ByVal blnMoveSpace As Boolean = True) As String
    '1 ȥ��һ���ַ�: " '_%?"����_%?ת��Ϊ��Ӧ��ȫ���ַ�
    '2 ȥ�������ַ�:�˸��Ʊ����С��س�
    '3 blnMoveSpace���Ƿ�ȥ���ַ��еĿո�Ture-ȥ���ո�ע��ͷβ�ո�Ĭ��ȥ��
    Dim n As Integer
    Dim intStrLen As Integer
    Dim intAsc As Integer
    Dim strText As String
    Dim strTmp As String
    Const CST_SPECIALCHAR = "_%?"      '����ת�����ַ�
    
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
                    strTmp = strTmp & "��"
                Case "%"
                    strTmp = strTmp & "��"
                Case "_"
                    strTmp = strTmp & "��"
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
                '�ո���
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
    '���ܣ������ַ����ļ���
    '��Σ�strInput-�����ַ�����bytIsWB-�Ƿ����(����Ϊƴ��)
    '���Σ���ȷ�����ַ��������󷵻�"-"
    '----------------------------------
    Dim strReturn As String
    
    strReturn = zlCommFun.zlGetSymbol(strInput, bytIsWB)
    
    zlGetSymbol = Mid(strReturn, 1, intOutNum)
End Function

Public Function zlClinicCodeRepeat(strInputCode As String, Optional lngSelfID As Long) As Boolean
    '----------------------------------
    '���ܣ����������Ŀ������Ƿ������б����ظ����ظ��������ʾ
    '��Σ�strInputCode-����ı��룻lngSelfID-�Լ���ID�ţ����޸�ʱ����Ҫ��������������ж�
    '���Σ��ظ�����True��������Flase
    '----------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    strSql = "select K.����||' ['||I.����||']'||I.���� as ����" & _
            " from ������ĿĿ¼ I,������Ŀ��� K" & _
            " where I.���=K.���� and I.����=[1] " & _
            "       and I.ID<>[2]"
    Err = 0: On Error GoTo ErrHand
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISBase", strInputCode, lngSelfID)
        
    With rsTmp
        If .RecordCount <> 0 Then
            MsgBox "����Ŀ�롰" & !���� & "�������ظ���", vbExclamation, gstrSysName
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
    '���ܣ�����շ���Ŀ������Ƿ������б����ظ����ظ��������ʾ
    '��Σ�strInputCode-����ı��룻lngSelfID-�Լ���ID�ţ����޸�ʱ����Ҫ��������������ж�
    '���Σ��ظ�����True��������Flase
    '----------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    strSql = "select K.����||' ['||I.����||']'||I.���� as ����" & _
            " from �շ���ĿĿ¼ I,�շ���Ŀ��� K" & _
            " where I.���=K.���� and I.����=[1] " & _
            "       and I.ID<>[2]"
    Err = 0: On Error GoTo ErrHand
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISBase", strInputCode, lngSelfID)
    
    With rsTmp
        If .RecordCount <> 0 Then
            MsgBox "����Ŀ�롰" & !���� & "�������ظ���", vbExclamation, gstrSysName
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
    '���ܣ������Ŀ�Ƿ����,���ڲ�������ʱ�ļ��
    '��Σ�strTableName ���� ,strField �ֶ��� , ,lngItemID,�ֶε�ֵ,strItemName ��ʾʱ��ʾ����Ŀ����
    '���Σ����ڷ���True��������Flase
    '----------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    Err = 0: On Error GoTo ErrHand
    strSql = "Select " & strField & " From " & strTbleName & " Where " & strField & "=[1]"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISBase", varValues)
    If rsTmp.RecordCount > 0 Then
        zlExistItem = True
    Else
         MsgBox "��" & strItemName & "���Ѿ�����������Աɾ����", vbExclamation, gstrSysName
        zlExistItem = False
    End If
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlExistItem = False
End Function

Public Function CheckIsInclude(strSource As String, strTarge As String) As Boolean
'���strSource�е�ÿһ���ַ��Ƿ���strTarge��
    Dim i As Long
    CheckIsInclude = False
    
    Select Case strTarge
    Case "����"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "ʱ��"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"";|\=+-_)(*&^%$#@!`~"
    Case "����ʱ��"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"";|\=+_)(*&^%$#@!`~"
    Case "����"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "С��"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+_)(*&^%$#@!`~"
    Case "������"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},.<>?/'"":;|\=+-_)(*&^%$#@!`~"
    Case "��С��"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/'"":;|\=+-_)(*&^%$#@!`~"
    Case "�ɴ�ӡ�ַ�"
        strTarge = "ZXCVBNMASDFGHJKLQWERTYUIOP[]{},<>?/.'"":;|\=+-_)(*&^%$#@!`~0123456789"
    End Select
    For i = 1 To Len(strSource)
        If InStr(strTarge, Mid(strSource, i, 1)) <= 0 Then Exit Function
    Next
    CheckIsInclude = True
End Function

Public Function Between(x, a, b) As Boolean
'���ܣ��ж�x�Ƿ���a��b֮��
    If a < b Then
        Between = x >= a And x <= b
    Else
        Between = x >= b And x <= a
    End If
End Function

Public Function FormatEx(ByVal vNumber As Variant, ByVal intBit As Integer) As String
'���ܣ��������뷽ʽ��ʽ����ʾ����,��֤С������󲻳���0,С����ǰҪ��0
'������vNumber=Single,Double,Currency���͵�����,intBit=���С��λ��
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
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function

Public Function ZVal(ByVal varValue As Variant) As String
'���ܣ���0��ת��Ϊ"NULL"��,������SQL���ʱ��
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
'����: ���ָ�������ָ����ָ���е�����
'����: obj=Ҫ����������ؼ�
'      intRow=Ҫ������к�
'      intCol=Ҫ������к��б���Array(1,2,3),������������Ա�ʾΪArray()
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
'����ַ����Ƿ��зǷ��ַ�������ṩ���ȣ��Գ��ȵĺϷ���Ҳ����⡣
    If InStr(strInput, "'") > 0 Then
        MsgBox "���������ݺ��зǷ��ַ���", vbExclamation, gstrSysName
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox "���������ݲ��ܳ���" & Int(intMax / 2) & "������" & "��" & intMax & "���ַ���", vbExclamation, gstrSysName
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
'���ܣ��򿪼�¼��
'    If rsTemp.State = adStateOpen Then rsTemp.Close
    
    On Error GoTo ErrHandle
'    Call SQLTest(App.ProductName, strCaption, IIf(strSql = "", gstrSql, strSql))
    Set rsTemp = zlDatabase.OpenSQLRecord(IIf(strSql = "", gstrSql, strSql), "cmd����_Click")
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
    '���ܣ� ����X,Y��ʵ�����꣬��������Ļ���������
    '������ X---���غ��������
    '       Y---�������������
    '----------------------------------------------------------------------
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(objBill.hWnd, objPoint)
    
    x = objPoint.x * 15 + objBill.CellLeft
    y = objPoint.y * 15 + objBill.CellTop + objBill.CellHeight
End Sub

Public Function FillGrid(ByRef objMsf As Object, ByVal rsData As ADODB.Recordset, Optional ByVal MaskArray As Variant, Optional ByVal blnClear As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------
    '����:������ݵ�����
    '����:
    '����:
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
    '����:
    '����:
    '����:
    '-------------------------------------------------------------------------------------------------------------
    Dim objItem As ListItem
    Dim lngLoop As Long
    
    On Error GoTo ErrHand
    
    LockWindowUpdate objLvw.hWnd
    
    Do While Not rs.EOF
        Set objItem = objLvw.ListItems.Add(, "K" & rs("ID").Value, rs("����").Value, _
                      IIf(rs("��Ŀ���") = "΢����" Or rs("���") = "��", "ItemGroup", "Item"), _
                      IIf(rs("��Ŀ���") = "΢����" Or rs("���") = "��", "ItemGroup", "Item"))
                      
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
    Case 1      '������
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
'����: װ��������ָ�������������������е���������
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
    '���ܣ�
    '------------------------------------------------------------------------------------------------------
    MsgBox strInfo, vbInformation, gstrSysName
    
End Sub

Public Sub ClearGrid(vsf As Object, Optional ByVal Row As Long = 1)
    '--------------------------------------------------------------------------------------------------------
    '����:����������
    '--------------------------------------------------------------------------------------------------------
    vsf.Rows = Row + 1
    vsf.RowData(Row) = 0
    vsf.Cell(flexcpText, Row, 0, Row, vsf.Cols - 1) = ""
    
End Sub

Public Function CheckNumeric(ByVal strText As String, ByVal lngLength As Long, Optional ByVal lngDecLength As Long = 0, Optional ByVal bytMode As Byte = 1) As String
    '--------------------------------------------------------------------------------------------------------
    '����:����ַ�������ֵ��Ч��
    '--------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    
    Dim str�������� As String
    Dim strС������ As String
    
    If lngDecLength = 0 Then
        '����
        Select Case bytMode
        Case 1      '������
            str�������� = strText
        Case 2      '������
            If Left(strText, 1) <> "-" And strText <> "0" Then
                CheckNumeric = "ӦΪ���������㣡"
                Exit Function
            End If
            str�������� = Mid(strText, 2)
            
        Case 3      '��������
            If Left(strText, 1) = "-" Then str�������� = Mid(strText, 2)
        End Select
    Else
        'С��
        Select Case bytMode
        Case 1      '��С��
            If Len(strText) > lngLength + 1 Then
                CheckNumeric = "���ȳ�����" & lngLength & "λ��"
                Exit Function
            End If
            
            If InStr(strText, ".") > 0 Then
                '��С������
                str�������� = Left(strText, InStr(strText, ".") - 1)
                strС������ = Mid(strText, InStr(strText, ".") + 1)
            Else
                str�������� = strText
            End If
            
        Case 2      '��С��
            If Len(strText) > lngLength + 2 Then
                CheckNumeric = "���ȳ�����" & lngLength & "λ��"
                Exit Function
            End If
            
            If Left(strText, 1) <> "-" Then
                CheckNumeric = "���Ǹ�����"
                Exit Function
            End If
            
            If InStr(strText, ".") > 0 Then
                '��С������
                str�������� = Mid(strText, 2, InStr(strText, ".") - 2)
                strС������ = Mid(strText, InStr(strText, ".") + 1)
            Else
                str�������� = Mid(strText, 2)
            End If
            
        Case 3      '����С��
            If Left(strText, 1) = "-" Then
                If Len(strText) > lngLength + 2 Then
                    CheckNumeric = "���ȳ�����" & lngLength & "λ��"
                    Exit Function
                End If
                If InStr(strText, ".") > 0 Then
                    '��С������
                    str�������� = Mid(strText, 2, InStr(strText, ".") - 2)
                    strС������ = Mid(strText, InStr(strText, ".") + 1)
                Else
                    str�������� = Mid(strText, 2)
                End If
            Else
                If Len(strText) > lngLength + 1 Then
                    CheckNumeric = "���ȳ�����" & lngLength & "λ��"
                    Exit Function
                End If
                If InStr(strText, ".") > 0 Then
                    '��С������
                    str�������� = Mid(strText, 1, InStr(strText, ".") - 1)
                    strС������ = Mid(strText, InStr(strText, ".") + 1)
                Else
                    str�������� = strText
                End If
                
            End If
        End Select
    End If
    
    If Len(str��������) > (lngLength - lngDecLength) Then
        If lngDecLength = 0 Then
            CheckNumeric = "���ȳ�����" & (lngLength - lngDecLength) & "λ��"
        Else
            CheckNumeric = "�������ݳ��ȳ�����" & (lngLength - lngDecLength) & "λ��"
        End If
        Exit Function
    End If
    
    If Len(strС������) > lngDecLength Then
        CheckNumeric = "С�����ݳ��ȳ�����" & lngDecLength & "λ��"
        Exit Function
    End If
    
    For lngLoop = 1 To Len(str��������)
        If Mid(str��������, lngLoop, 1) < "0" Or Mid(str��������, lngLoop, 1) > "9" Then
            CheckNumeric = "ӦΪ�����ͣ�"
            Exit Function
        End If
    Next
    
    For lngLoop = 1 To Len(strС������)
        If Mid(strС������, lngLoop, 1) < "0" Or Mid(strС������, lngLoop, 1) > "9" Then
            CheckNumeric = "ӦΪ�����ͣ�"
            Exit Function
        End If
    Next
    
    
    CheckNumeric = ""
End Function

'ȥ��TextBox��Ĭ���Ҽ��˵�
Public Function WndMessage(ByVal hWnd As OLE_HANDLE, ByVal Msg As OLE_HANDLE, ByVal wp As OLE_HANDLE, ByVal lp As Long) As Long
    ' �����Ϣ����WM_CONTEXTMENU���͵���Ĭ�ϵĴ��ں�������
    If Msg <> WM_CONTEXTMENU Then WndMessage = CallWindowProc(glngTXTProc, hWnd, Msg, wp, lp)
End Function

Public Function GetSysPara(ByVal int��� As Integer) As String
    Dim rsTemp As New ADODB.Recordset
    '��ȡϵͳ����
    On Error GoTo ErrHandle
    gstrSql = "Select Nvl(����ֵ,ȱʡֵ) From Zlparameters Where ϵͳ = [1] And Nvl(˽��, 0) = 0 And ģ�� Is Null And ������=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "��ȡϵͳ����ֵ", glngSys, int���)
    
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
'���ܣ���ItemData����ComboBox������ֵ
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
'���ܣ����ش�д�ĵ��ݺ���ǰ׺
    If curDate = #1/1/1900# Then
        PreFixNO = CStr(CInt(Format(zlDatabase.Currentdate, "YYYY")) - 1990)
    Else
        PreFixNO = CStr(CInt(Format(curDate, "YYYY")) - 1990)
    End If
    PreFixNO = IIf(CInt(PreFixNO) < 10, PreFixNO, Chr(55 + CInt(PreFixNO)))
End Function

Public Function GetFullNO(ByVal strNo As String, ByVal intNum As Integer) As String
'���ܣ����û�����Ĳ��ݵ��ţ�����ȫ���ĵ��š�
'������intNum=��Ŀ���,Ϊ0ʱ�̶��������
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
    
    strSql = "Select ��Ź���,Sysdate as ���� From ������Ʊ� Where ��Ŀ���=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, App.ProductName, intNum)
    If Not rsTmp.EOF Then
        intType = Val("" & rsTmp!��Ź���)
        curDate = rsTmp!����
    End If

    If intType = 1 Then
        '���ձ��
        strSql = Format(CDate(Format(rsTmp!����, "YYYY-MM-dd")) - CDate(Format(rsTmp!����, "YYYY") & "-01-01") + 1, "000")
        GetFullNO = PreFixNO & strSql & Format(Right(strNo, 4), "0000")
    Else
        '������
        GetFullNO = PreFixNO & Format(Right(strNo, 7), "0000000")
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'################################################################################################################
'## ���ܣ�  �����ݴ�һ��XtremeReportControl�ؼ����Ƶ�VSFlexGrid���Ա���д�ӡ
'################################################################################################################
Public Function zlReportToVSFlexGrid(vfgList As VSFlexGrid, rptList As ReportControl) As Boolean
    '-------------------------------------------------
    '��ȫ����ǿ��չ��,�������ݱ��
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
        
        '�����и���
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
        
        '�����и���
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
                    .TextMatrix(.Rows - 1, lngCol) = String(intTiers, "��") & rptList.GroupsOrder(intTiers).Caption & ": "
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
                        .TextMatrix(.Rows - 1, lngCol) = String(intTiers, "��") & rptRow.Record(.ColData(lngCol)).Value
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
    'ɾ���Ƿ��ַ�
    'strChar: Ҫ������ַ�
    'strInvalidChar���Ƿ��ַ��������Ϊ�գ���Ϊ~!@#$%^&*()_+|=-`;'"":/.,<>?{}[]\<>,���򰴴�����ַ�����
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
'���ܣ����ϵͳ�Ƿ���ڿ���ҩ����Ȩ����Ա���������õ�ǰ����Ա����ҩ����UserInfo����
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    UserInfo.��ҩ���� = 0
    
    On Error GoTo errH
    strSql = "Select ���� From ��Ա����ҩ��Ȩ�� Where ��¼״̬=1 and ��ԱID = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISKernel", UserInfo.ID)
    If rsTmp.RecordCount > 0 Then
        UserInfo.��ҩ���� = Val("" & rsTmp!����)
        CheckKSSPrivilege = True
    Else
        strSql = "Select 1 From ��Ա����ҩ��Ȩ�� Where ��¼״̬=1 and Rownum<2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "mdlCISKernel")
        CheckKSSPrivilege = rsTmp.RecordCount > 0
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function IntEx(vNumber As Variant) As Variant
'���ܣ�ȡ����ָ����ֵ����С����
    IntEx = -1 * Int(-1 * Val(vNumber))
End Function



Public Function FmgFlexScroll(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'֧��frmDoctorManage������ֵĹ���
    On Error GoTo errH
    Select Case wMsg
    Case WM_MOUSEWHEEL
        Select Case wParam
            Case -7864320  '���¹�
                If frmDoctorManage.vscBar.Value <> frmDoctorManage.vscBar.Max Then
                    frmDoctorManage.vscBar.SetFocus
                    zlCommFun.PressKey vbKeyPageDown
                End If
            Case 7864320   '���Ϲ�
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
'���ܣ���ģ̬�������������ַ�����
'������frmParent=���ø�����
'���أ�ѡ��������ַ�����ȡ���������ؿ�
    Dim frmNew As frmSpecChar
    Set frmNew = New frmSpecChar
    frmNew.Show 1, frmParent
    If gblnOK Then ShowSpecChar = frmNew.mstrChar
End Function

Public Sub ArrayIcons(objLvw As ListView, Optional intBegin As Integer = 1, Optional blnShow As Boolean)
'���ܣ����ݵ�һ��ͼ���λ��������������ͼ��
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
    
    '�����ͼ��
    t = objLvw.ListItems(intBegin).Top
    For i = intBegin To objLvw.ListItems.Count
        With objLvw.ListItems(i)
            'Item��Width�������ֲ���,Left��ָͼ��
            .Left = ((r.Right - r.Left - objLvw.Icons.ImageWidth) * Screen.TwipsPerPixelX) / 2
            .Top = t
            t = t + .Height
        End With
    Next
    
    '�����ͼ��
    t = objLvw.ListItems(intBegin).Top
    For i = intBegin To 1 Step -1
        With objLvw.ListItems(i)
            'Item��Width�������ֲ���,Left��ָͼ��
            .Left = ((r.Right - r.Left - objLvw.Icons.ImageWidth) * Screen.TwipsPerPixelX) / 2
            .Top = t
            t = t - .Height
        End With
    Next
End Sub

Public Function GetArrayByStr(ByVal strInput As String, ByVal lngLength As Long, ByVal strSplitChar As String) As Variant
    '���ݴ�����ַ������зֽ⣬����ָ���ַ����Ⱦ���Ҫ���зֽ⣬������浽������
    '��Σ�strInput-������ַ�����strSplitChar-�ַ��������ݵķָ���
    '���أ����飬���������Ա���ַ����Ȳ�����ָ������
    Dim strArray As Variant
    Dim arrTmp As Variant
    Dim strTmp As String
    Dim lngCount As Long
    Dim i As Long
    
    strArray = Array()
   
    '����ָ���ַ�ʱ����Ҫ�ֽ�
    If Len(strInput) > lngLength Then
        If strSplitChar = "" Then
            '�޷ָ���ʱ
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
            '�зָ���ʱ
            arrTmp = Split(strInput & strSplitChar, strSplitChar)
            lngCount = UBound(arrTmp)
        
            For i = 0 To lngCount
                If arrTmp(i) <> "" Then
                    '�зָ�������Ҫ���ַָ���֮���ַ��������ԣ����ܰѷָ���֮����ַ���
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

Public Function CheckBatches(ByVal blnҩ����� As Boolean, ByVal blnҩ������ As Boolean) As Boolean
    '���ܣ����ҩ�����ҩ��������ʱ�����������Ƿ�ͬʱ��������ҩ��ҩ��
    
    Dim rs�������� As ADODB.Recordset
    
    On Error GoTo ErrHand
    
    If blnҩ����� = True And blnҩ������ = False Then
        gstrSql = "Select 1" & vbNewLine & _
                        "From ��������˵�� T" & vbNewLine & _
                        "Where t.����id In" & vbNewLine & _
                        "      (Select Distinct t.����id From ��������˵�� T Where t.�������� Like '%ҩ��')" & vbNewLine & _
                        "      And (t.�������� Like '%ҩ��' or t.�������� Like '%�Ƽ���')"
                        
        Set rs�������� = zlDatabase.OpenSQLRecord(gstrSql, "�Ƿ��в�������ͬʱ������ҩ��ҩ��")
        If rs��������.RecordCount > 0 Then
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




