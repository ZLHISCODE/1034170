Attribute VB_Name = "mdlStuff"
 Option Explicit

Public gcnOracle As ADODB.Connection        '�������ݿ����ӣ��ر�ע�⣺��������Ϊ�µ�ʵ��
Public gstrPrivs As String                   '��ǰ�û����еĵ�ǰģ��Ĺ���
Public gstrSysName As String                'ϵͳ����
Public glngModul As Long
Public glngSys As Long
Public gstrAviPath As String
Public gstrVersion As String
Public gstrMatchMethod As String
Public gbytSimpleCodeTrans As Byte          '��Ƭ�����Ƿ���������л�����

Public gstrDBUser As String                 '��ǰ���ݿ��û�
Public glngUserId As Long                   '��ǰ�û�id
Public gstrUserCode As String               '��ǰ�û�����
Public gstrUserName As String               '��ǰ�û�����
Public gstrUserAbbr As String               '��ǰ�û�����

Public glngDeptId As Long                   '��ǰ�û�����id
Public gstrDeptCode As String               '��ǰ�û����ű���
Public gstrDeptName As String               '��ǰ�û���������

Public gstrUnitName As String '�û���λ����
Public gfrmMain As Object

Public gstrSQL As String
Public gblnOK As Boolean
Public gstrIme As String

Public gobjSquareCard As Object             'һ��ͨ�ӿ�
Public gstrCardType As String           '���п���𣬸�ʽ������|ȫ��|ˢ����־|�����ID|���ų���|ȱʡ��־(1-��ǰȱʡ;0-��ȱʡ)|�Ƿ�����ʻ�(1-�����ʻ�;0-�������ʻ�)|��������(�ڼ�λ���ڼ�λ����,��Ϊ������);��
Public gintCardCount As Integer  '������
Public gblnIncomeItem As Boolean            '��¼����Ŀ¼�������Ƿ�������������Ŀ

Public gobjPlugIn As Object             '��ҽӿ�

Public Const glngRowByFocus = &HFFE3C8
Public Const glngRowByNotFocus = &HF4F4EA
Public Const glngFixedForeColorByFocus = &HFF0000
Public Const glngFixedForeColorNotFocus = &H80000012

'ҩƷ���۸�������󾫶�
Public Type Type_Digits
    Digit_��� As Integer
    Digit_�ɱ��� As Integer
    Digit_���ۼ� As Integer
    Digit_���� As Integer
End Type
Public gtype_UserDrugDigits As Type_Digits

'���ѿ���ʽ
Public Enum gCardFormat
    ���� = 0
    ȫ�� = 1
    ˢ����־ = 2
    �����ID = 3
    ���ų��� = 4
    ȱʡ��־ = 5
    �Ƿ�����ʻ� = 6
    �������� = 7
End Enum

Public Type TYPE_USER_INFO
    Id As Long
    ����ID As Long
    ��� As String
    ���� As String
    ���� As String
    �û��� As String
End Type
Public gOraFmt_Max As g_FmtString


Public UserInfo As TYPE_USER_INFO
Public Declare Function SetFocusHwnd Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

'�л���ָ�������뷨��
Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal hkl As Long, ByVal flags As Long) As Long
'����ϵͳ�п��õ����뷨�����������뷨����Layout,����Ӣ�����뷨��
Public Declare Function GetKeyboardLayoutList Lib "user32" (ByVal nBuff As Long, lpList As Long) As Long
'��ȡĳ�����뷨������
Public Declare Function ImmGetDescription Lib "imm32.dll" Alias "ImmGetDescriptionA" (ByVal hkl As Long, ByVal lpsz As String, ByVal uBufLen As Long) As Long
'�ж�ĳ�����뷨�Ƿ��������뷨
Public Declare Function ImmIsIME Lib "imm32.dll" (ByVal hkl As Long) As Long
'��ȡָ�����뷨����Layout,����Ϊ0ʱ��ʾ��ǰ���뷨��
Public Declare Function GetKeyboardLayout Lib "user32" (ByVal dwLayout As Long) As Long
'��ȡ��ǰ���뷨����Layout��
Public Declare Function GetKeyboardLayoutName Lib "user32" Alias "GetKeyboardLayoutNameA" (ByVal pwszKLID As String) As Long
'�������뷨Layout���������뷨�л������뷨�л�˳�����ǰͷ(������������Ч),flags����=KLF_REORDER
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
'��ʼ���ڵı�־
Public Enum StartDayFlag
    FirstDayOfWeek = 0
    FirstDayOfMonth = 1
    FirstDayOfQuarter = 2
    FirstDayOfHalfYear = 3
    FirstDayOfyear = 4
End Enum
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Public Sub zlPlugIn_Ini(ByVal lngSys As Long, ByVal lngModul As Long, objPlugIn As Object)
    '�����չ�ӿڳ�ʼ��
    If objPlugIn Is Nothing Then
        On Error Resume Next
        Set objPlugIn = CreateObject("zlPlugIn.clsPlugIn")
        If Not objPlugIn Is Nothing Then
            Call objPlugIn.Initialize(gcnOracle, lngSys, lngModul)
            If InStr(",438,0,", "," & err.Number & ",") = 0 Then
                MsgBox "zlPlugIn ��Ҳ���ִ�� Initialize ʱ����" & vbCrLf & err.Number & vbCrLf & err.Description, vbInformation, gstrSysName
            End If
        End If
        err.Clear: On Error GoTo 0
    End If
End Sub

Public Sub zlPlugIn_Unload(objPlugIn As Object)
    'ж����ҽӿ�
    Set objPlugIn = Nothing
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
Public Function GetUserInfo() As Boolean
'���ܣ���ȡ��½�û���Ϣ
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = zlDatabase.GetUserInfo
    
    UserInfo.�û��� = gstrDBUser
    UserInfo.���� = gstrDBUser
    If Not rsTmp.EOF Then
        UserInfo.Id = rsTmp!Id
        UserInfo.��� = rsTmp!���
        UserInfo.����ID = IIf(IsNull(rsTmp!����ID), 0, rsTmp!����ID)
        UserInfo.���� = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        UserInfo.���� = IIf(IsNull(rsTmp!����), "", rsTmp!����)
        UserInfo.�û��� = UserInfo.����
        gstrUserName = IIf(IsNull(rsTmp!����), "", rsTmp!����)
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
'���ܣ���ItemData��Text����ComboBox������ֵ
    Dim strType As String, i As Integer
    
    SeekCboIndex = -1
    
    strType = TypeName(varData)
    If strType = "Field" Then
        If IsType(varData.Type, adVarChar) Then strType = "String"
    End If
    
    If strType = "String" Then
        If varData <> "" Then
            '�Ⱦ�ȷ����
            For i = 0 To objCbo.ListCount - 1
                If objCbo.List(i) = varData Then
                    SeekCboIndex = i: Exit Function
                ElseIf NeedName(objCbo.List(i)) = varData And varData <> "" Then
                    SeekCboIndex = i: Exit Function
                End If
            Next
            '��ģ������
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
'���ܣ��ж�ĳ��ADO�ֶ����������Ƿ���ָ���ֶ�������ͬһ��(������,����,�ַ�,������)
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
    '������������ȡָ����ı����������󳤶�
    '�������������ID������
    '����������ɹ����� �¼�������; ���߷��� 0
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    err = 0
    On Error GoTo Error_Handle
    If strID = "" Then
        strSQL = "select nvl(max(Vsize(����)),0) as LenCode from " & strTableName & " start with �ϼ�ID is null " & strWhere & " connect by prior id=�ϼ�id"
    Else
        strSQL = "select nvl(max(Vsize(����)),0) as LenCode from " & strTableName & " start with ID=" & strID & strWhere & " connect by prior id=�ϼ�id"
    End If
    
    Call zlDatabase.OpenRecordset(rsTemp, strSQL, "��ȡָ����ı����������󳤶�")
    
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

Public Function GetLocalCodeLength(ByVal str�ϼ�ID As String, ByVal strTableName As String, Optional ByVal strWhere As String) As Long
    '������������ȡָ����ı����������󳤶�
    '����������ϼ�ID������
    '����������ɹ����� ������; ���߷��� 0
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    err = 0
    On Error GoTo Error_Handle
    If str�ϼ�ID = "" Then
        strSQL = "select nvl(max(Vsize(����)),0) as LenCode from " & strTableName & " where �ϼ�ID is null" & strWhere
    Else
        strSQL = "select nvl(max(Vsize(����)),0) as LenCode from " & strTableName & " where �ϼ�ID=" & str�ϼ�ID & strWhere
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

Public Function GetParentCode(ByVal str�ϼ�ID As String, ByVal strTableName As String) As String
    '������������ȡ�ϼ�����
    '����������ϼ�ID,����
    '����������ɹ����� �ϼ�����; ���߷��� ��
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    err = 0
    On Error GoTo Error_Handle
    If str�ϼ�ID = "" Then
        GetParentCode = ""
        Exit Function
    End If
    
    strSQL = "select ���� from " & strTableName & " where ID=[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�ϼ�����", str�ϼ�ID)
    
    If rsTemp.EOF Then
        GetParentCode = ""
    Else
        GetParentCode = rsTemp.Fields("����").Value
    End If
    Exit Function
Error_Handle:
    Call ErrCenter
    Call SaveErrLog
    GetParentCode = ""
End Function

Public Function GetMaxLocalCode(ByVal str�ϼ�ID As String, ByVal strTableName As String, Optional ByVal strWhere As String) As String
    '��������������ָ������ϼ�ID ��ȡ������������
    '����������ϼ�ID,����
    '����������ɹ����� ������; ���߷��� ��
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim intCode As Integer, strCode As String, strAllCode As String
    Dim intLength   As Integer
    err = 0
    On Error GoTo Error_Handle
    If str�ϼ�ID = "" Then
        strSQL = "select max(to_number(����))+1 as MaxCode from " & strTableName & " where �ϼ�ID is null" & strWhere
    Else
        strSQL = "select nvl(max(to_number(����)),0)+1 as MaxCode from " & strTableName & " where �ϼ�ID=" & str�ϼ�ID & strWhere
    End If
    intCode = GetLocalCodeLength(str�ϼ�ID, strTableName, strWhere)
    
    Call zlDatabase.OpenRecordset(rsTemp, strSQL, "����ָ������ϼ�ID ��ȡ������������")
    If rsTemp.EOF Then
        GetMaxLocalCode = ""
        Exit Function
    End If
    intLength = intCode - Len(IIf(IsNull(rsTemp.Fields("MaxCode").Value), 0, rsTemp.Fields("MaxCode").Value))
    strAllCode = String(IIf(intLength < 0, 0, intLength), "0") & rsTemp.Fields("MaxCode").Value
    GetMaxLocalCode = Mid(strAllCode, Len(GetParentCode(str�ϼ�ID, strTableName)) + 1)
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
    '���ܣ���ʾ��Ϣ��
    '������strMsgInfor-��ʾ��Ϣ
    '     blnYesNo-�Ƿ��ṩYES��NO��ť
    '���أ�blnYes-����ṩYESNO��ť,�򷵻�YES(True)��NO(False)
    '----------------------------------------------------------------------------------------------------------------
        
    If blnYesNo = False Then
        MsgBox strMsgInfor, vbInformation + vbDefaultButton1, gstrSysName
    Else
        blnYes = MsgBox(strMsgInfor, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes
    End If
End Sub

Public Sub zlChangeCode(ByVal strTableName As String, _
    ByVal lng�ϼ�id As Long, _
    ByVal txtUpCode As TextBox, _
    ByVal txtCode As TextBox, _
    Optional ByVal chkChangeCode As CheckBox = Nothing, _
    Optional ByVal strCaption As String = "")
    '------------------------------------------------------------------------------------
    '���ܣ�����ѡ����ϼ�ȷ����ǰ�ı��룬�����ϼ�����������ʾ����
    '������strTableName-���ڷ���ı���
    '      lng�ϼ�ID-ѡ����ϼ�
    '      TxtUpCode-��ʾ���ϼ��ı���
    '      TxtUpCode-��ʾ�ı����ı���
    '      chkChangeCode-�����Ƿ�ı�ԭ�����ݿ��е���ʷ����ѡ��ؼ�
    '      strCaption-���ô����Capiton
    'ע�⣺���б�����ID,�ϼ�id,����
    '------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim intMaxCodeLen As Integer  'ȷ�������ʵ�ʳ���
    err = 0: On Error GoTo ErrHand
    
   chkChangeCode.Value = 0
   chkChangeCode.Enabled = True
   
    If lng�ϼ�id = 0 Then
        txtUpCode.Text = ""
        gstrSQL = "select max(����) as ���� From " & strTableName & " Where �ϼ�ID is null "
        zlDatabase.OpenRecordset rsTemp, gstrSQL, strCaption
            
        With rsTemp
            intMaxCodeLen = .Fields("����").DefinedSize
            If IsNull(!����) Then
                txtCode.Text = "01"
                txtCode.MaxLength = intMaxCodeLen
                txtCode.Tag = txtCode.MaxLength
                chkChangeCode.Value = 1
                chkChangeCode.Enabled = False
            Else
                txtCode.MaxLength = Len(Trim(!����))
                txtCode.Tag = txtCode.MaxLength
                If !���� = String(txtCode.MaxLength, "9") Then
                    If txtCode.MaxLength >= intMaxCodeLen Then
                        ShowMsgBox "������ͱ��볤���Ѿ��ﵽ������ƣ��޷���������"
                        txtCode.Text = Space(txtCode.MaxLength)
                       chkChangeCode.Value = 0
                       chkChangeCode.Enabled = False
                    Else
                        ShowMsgBox "�������Ѿ��ﵽ�������ƣ������������볤����������Ҫ"
                        txtCode.Text = "1" & String(txtCode.MaxLength, "0")
                        txtCode.MaxLength = txtCode.MaxLength + 1
                        txtCode.Tag = txtCode.MaxLength
                       chkChangeCode.Value = 1
                    End If
                Else
                    txtCode.Text = Format(Mid(!����, Len(txtUpCode.Text) + 1) + 1, String(txtCode.MaxLength, "0"))
                End If
            End If
        End With
        Exit Sub
   End If
   'ȷ���ϼ�����
   
    gstrSQL = "Select ���� From " & strTableName & " where id=[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, strCaption, lng�ϼ�id)
    
    If Not rsTemp.EOF Then
        txtUpCode.Text = zlCommFun.NVL(rsTemp!����)
    End If
    
    '��ȷ���Ƿ����¼�
    gstrSQL = "select nvl(max(����),'') as ����  From " & strTableName & " Where  �ϼ�ID =[1] "
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, strCaption, lng�ϼ�id)
    
    intMaxCodeLen = rsTemp.Fields("����").DefinedSize

    If zlCommFun.NVL(rsTemp!����) = "" Then
        '�������¼�
        '�����ϼ�IDȡ�ϼ�����
'        gstrSQL = "Select ���� From " & strTableName & " where id=" & lng�ϼ�id
'        zlDatabase.OpenRecordset rsTemp, gstrSQL, strCaption
'        txtUpCode.Text = zlCommFun.Nvl(rsTemp!����)
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
        txtCode.MaxLength = Len(!����) - Len(txtUpCode.Text)
        txtCode.Tag = txtCode.MaxLength
        If Mid(!����, Len(txtUpCode.Text) + 1) = String(txtCode.MaxLength, "9") Then
            If Len(txtUpCode.Text) + txtCode.MaxLength >= intMaxCodeLen Then
                ShowMsgBox "�÷����¼�������ͱ��볤���Ѿ��ﵽ������ƣ��޷���������"
                txtCode.Text = Space(txtCode.MaxLength)
               chkChangeCode.Value = 0
               chkChangeCode.Enabled = False
            Else
                ShowMsgBox "�÷����¼��������Ѿ��ﵽ�������ƣ������������볤����������Ҫ"
                txtCode.Text = "1" & String(txtCode.MaxLength, "0")
                txtCode.MaxLength = txtCode.MaxLength + 1
                txtCode.Tag = txtCode.MaxLength
               chkChangeCode.Value = 1
            End If
        Else
            txtCode.Text = Format(Mid(!����, Len(txtUpCode.Text) + 1) + 1, String(txtCode.MaxLength, "0"))
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
    '����: ��/�ر����뷨
    '����: blnOpen-�Ǵ򿪻��ǹر�(trueΪ��,falseΪ�ر�)
    '���أ�
    '-----------------------------------------------------------------------------------
    If blnOpen Then
        zlCommFun.OpenIme (True)
    Else
        zlCommFun.OpenIme False
    End If
End Sub


Public Sub SetTxtGotFocus(ByVal objTxt As Object, Optional blnOpenIme As Boolean = False)
    '--------------------------------------------------------------------------------------------------------
    '���ܣ����ı���ĵ��ı�ѡ�л����������뷨
    '����:blnOpenIme-�Ƿ�����뷨
    '����:
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

Public Function DepotProperty(ByVal lng��Աid As Long) As Boolean
    Dim rsCheck As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    '����ָ����Ա�Ƿ����ҩ������
    gstrSQL = "Select Distinct �������� From ������Ա B,��������˵�� A " & _
             " Where A.�������� = '���Ŀ�' And " & _
             " A.����id = B.����id And B.��Աid = [1]"
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��������", lng��Աid)
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
    '�Ƿ�������ҩ����Ա�鿴���ݵĳɱ���
    blnCostPrice = Val(zlDatabase.GetPara(190, 100, , 0))
    
    'ҩ����Ա���ܣ�ֻ��ҩ����Ա���Բ�������Ϊ׼
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
    '����:ȡĳ�ֶε�ֵ
    '����:rsObj          �������ֶ�
    '     varValue       ��rsObjΪNULLֵʱ��ȡ��ֵ
    '����:�����Ϊ��ֵ,����ԭ����ֵ,���Ϊ��ֵ,�򷵻�ָ����varValueֵ
    '-----------------------------------------------------------------------------------
    If IsNull(rsObj) Then
        NVL = varValue
    Else
        NVL = rsObj
    End If
End Function
Public Function Dec2Bin(bDec As Byte) As String

    '���ܣ�ʮ����תΪ�����ƺ���
    '�÷���String  Dec2Bin(Bdec as Byte)
    '���أ�  ʮ���ƵĶ����� �ַ���(String)
    '����  ����"0"
    Dim strBin As String

    On Error GoTo err
    If bDec > 255 Then
        Dec2Bin = "-1"
        Exit Function
    End If
    strBin = ""
    'תΪ�ַ���
    While bDec > 0
        strBin = bDec Mod 2 & strBin
        bDec = Fix(bDec / 2)
    Wend
    '������8λ
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
    '���ܣ�������תΪʮ���ƺ���
    '�÷���Long  bin2dec(strBin as String)
    '���أ�  �����Ƶ�ʮ���� ��������Long��
    '����  ����-1
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
    '������������ָ�����н�������
    '���������mshFilter-ָ��������
    '          intPreCol-�ϴ���
    '           intPreSort-�ϴ�����
    '           blnNum-�Ƿ�Ϊ������
    '���������
    '���أ�
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
    '�������������ҷ�����������
    '���������varTemp-ָ����ֵ
    '           mshGrd-ָ������
    '           intCol-ָ������
    '���������
    '���أ��ɹ������ҵ�����
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

'�����룬���ƣ���������ĳһ��
Public Function FindRownew(ByVal mshBill As BillEdit, ByVal int�Ƚ��� As Integer, _
    ByVal str�Ƚ�ֵ As String, ByVal blnFirst As Boolean) As Boolean
    Dim intStartRow As Integer
    Dim intRow As Integer
    Dim strSpell As String
    Dim strCode As String
    Dim rsCode As New Recordset
    
    On Error GoTo ErrHandle
    FindRownew = True
    With mshBill
        If .Rows = 2 Then Exit Function
        If str�Ƚ�ֵ = "" Then Exit Function
        
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
            If .TextMatrix(intRow, int�Ƚ���) <> "" Then
                strCode = .TextMatrix(intRow, int�Ƚ���)
                If InStr(1, UCase(strCode), UCase(str�Ƚ�ֵ)) <> 0 Then
                    .SetFocus
                    .Row = intRow
                    .Col = int�Ƚ���
                    .MsfObj.TopRow = .Row
                    .SetRowColor CLng(.Row), &HFFCECE, True
                    Exit Function
                End If
            End If
        Next
        gstrSQL = "" & _
        " SELECT DISTINCT b.���� " & _
        " FROM (SELECT DISTINCT A.�շ�ϸĿid " & _
        "       FROM �շ���Ŀ���� A" & _
        "       Where A.���� LIKE upper([1]) " & _
        "      ) A, �շ���ĿĿ¼ B " & _
        " Where a.�շ�ϸĿid = b.ID And (b.վ��=[2] or b.վ�� is null) "
        
        Set rsCode = zlDatabase.OpenSQLRecord(gstrSQL, "����ָ����������", GetMatchingSting(str�Ƚ�ֵ, False), gstrNodeNo)
        If rsCode.EOF Then
            FindRownew = False
            Exit Function
        End If
        
        For intRow = intStartRow To .Rows - 1
            If .TextMatrix(intRow, int�Ƚ���) <> "" Then
                strCode = .TextMatrix(intRow, int�Ƚ���)
                rsCode.MoveFirst
                Do While Not rsCode.EOF
                    If InStr(1, UCase(strCode), UCase(rsCode!����)) <> 0 Then
                        .SetFocus
                        .Row = intRow
                        .Col = int�Ƚ���
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
    '����:  ��ָ������Ϣ������ע�����
    '����:  RegType-ע������
    '       strSection-ע���Ŀ¼
    '       StrKey-����
    '       strKeyValue-��ֵ
    '����:
    '--------------------------------------------------------------------------------------------------------------
    err = 0: On Error GoTo ErrHand:
    Select Case RegType
        Case gע����Ϣ
            SaveSetting "ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue
        Case g����ȫ��
            SaveSetting "ZLSOFT", "����ȫ��\" & strSection, strKey, strKeyValue
        Case g����ģ��
            SaveSetting "ZLSOFT", "����ģ��" & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
        Case g˽��ȫ��
            SaveSetting "ZLSOFT", "˽��ȫ��\" & gstrDBUser & "\" & strSection, strKey, strKeyValue
        Case g˽��ģ��
            SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, strKeyValue
    End Select
ErrHand:
End Sub
Public Sub GetRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKey As String, ByRef strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '����:  ��ָ����ע����Ϣ��ȡ����
    '�����:  RegType-ע������
    '       strSection-ע���Ŀ¼
    '       StrKey-����
    '������:
    '       strKeyValue-���صļ�ֵ
    '����:
    '--------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    err = 0
    On Error GoTo ErrHand:
    Select Case RegType
        Case gע����Ϣ
            SaveSetting "ZLSOFT", "ע����Ϣ\" & strSection, strKey, strKeyValue
            strKeyValue = GetSetting("ZLSOFT", "ע����Ϣ\" & strSection, strKey, "")
        Case g����ȫ��
            strKeyValue = GetSetting("ZLSOFT", "����ȫ��\" & strSection, strKey, "")
        Case g����ģ��
            strKeyValue = GetSetting("ZLSOFT", "����ģ��" & "\" & App.ProductName & "\" & strSection, strKey, "")
        Case g˽��ȫ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser & "\" & strSection, strKey, "")
        Case g˽��ģ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKey, "")
    End Select
ErrHand:
End Sub
Public Function �ж�ֻ�߱����ϲ���(ByVal lng����ID As Long) As Boolean
    '�ж�ֻ�߱����ϱ����ʵ�:����ȡ���Ŀ���Ƽ������Ƶ����о߱����ϲ������ʵĲ���
    'lng����id-����id
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    �ж�ֻ�߱����ϲ��� = False
    gstrSQL = "select ��������, ����id, ������� from ��������˵�� where ����id =[1] And ��������='���ϲ���'"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ϲ��ŵĹ�������", lng����ID)
    
    
    If rsTemp.RecordCount = 0 Then
        Exit Function
    End If
    gstrSQL = "select ��������, ����id, ������� from ��������˵�� where ����id =[1] And �������� in( '���Ŀ�','�Ƽ���','����ⷿ')"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ϲ��ŵĹ�������", lng����ID)
    
    If rsTemp.RecordCount <> 0 Then
        Exit Function
    End If
    �ж�ֻ�߱����ϲ��� = True
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function CheckNOExists(ByVal int���� As Integer, ByVal strNo As String) As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "Select 1 From ҩƷ�շ���¼ Where NO=[2] And ����=[1] And Rownum<2"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�ж��Ƿ���ڸõ���", int����, strNo)
    If rsTemp.RecordCount = 0 Then Exit Function
    ShowMsgBox "�Ѿ����ڸõ��ݺ�(" & strNo & ")"
    CheckNOExists = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function GetFullNO(ByVal strNo As String, ByVal intNum As Integer, Optional ByVal lng����ID As Long) As String
'���ܣ����û�����Ĳ��ݵ��ţ�����ȫ���ĵ��š�
'������intNum=��Ŀ���,Ϊ0ʱ�̶��������
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, intType As Integer
    Dim curDate As Date
    Dim intYear As Integer
    Dim PreFixNO As String  '���ǰ׺
    Dim strPre As String    '���������ǰ2λ
    Dim str��� As String
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
    
    strSQL = "Select ��Ź���,������,Sysdate as ���� From ������Ʊ� Where ��Ŀ���=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "GetFullNO", intNum)
        
    If Not rsTmp.EOF Then
        intType = NVL(rsTmp!��Ź���, 0)
        curDate = rsTmp!����
        strPre = Left(NVL(rsTmp!������, PreFixNO & "0"), 2)
    End If
    
    If intType = 0 Then
        '������
        GetFullNO = strPre & Format(Right(strNo, 6), "000000")
    ElseIf intType = 1 Then
        '���ձ��
        strSQL = Format(CDate("1992-" & Format(rsTmp!����, "MM-dd")) - CDate("1992-01-01") + 1, "000")
        GetFullNO = PreFixNO & strSQL & Format(Right(strNo, 4), "0000")
    ElseIf intType = 2 Then
        '�����ҷ��±���
        gstrSQL = "Select ��� From ���Һ���� Where ��Ŀ���=[1] And Nvl(����ID,0)=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "GetFullNO", intNum, lng����ID)
        
        If rsTmp.RecordCount = 0 Then
            MsgBox "��δ���ÿ��ұ�ţ��޷��������룡", vbInformation, gstrSysName
            Exit Function
        End If
        If NVL(rsTmp!���) = "" Then
            MsgBox "��δ���ÿ��ұ�ţ��޷��������룡", vbInformation, gstrSysName
            Exit Function
        End If
        str��� = NVL(rsTmp!���)
        
        'С����λ�������²�������
        '��λ����λ������Ϊ��ָ���·ݵĺ���
        '��λ������Ϊ�ǲ�������ָ�����ҡ��·ݵĺ���
        '���ڵ��ڰ�λ��������
        If Len(strNo) <= 4 Then
            GetFullNO = PreFixNO & str��� & strMonth & String(4 - Len(strNo), "0") & strNo
        ElseIf Len(strNo) <= 6 Then
            GetFullNO = String(6 - Len(strNo), "0") & GetFullNO
            GetFullNO = PreFixNO & str��� & GetFullNO
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
    '����:ִ����ص�Oracle���̼�
    '����:cllProcs-oracle���̼�
    '     strCaption -ִ�й��̵ĸ����ڱ���
    '     blnNOCommit-ִ������̺�,���ύ����
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

Public Function Check�����ⰴ�����ۼ���() As Boolean
    '����:ȷ��ϵͳ�����ڸ�������µĳɱ����㷽ʽ
    Check�����ⰴ�����ۼ��� = Val(zlDatabase.GetPara(120, glngSys, 0)) = 1
End Function
Public Function ��֤�����ۼ���(ByVal lng�ⷿID As Long, ByVal lng����ID As Long, ByVal lng���� As Long, ByVal lng����ϵ�� As Long, _
                    ByVal dbl����� As Double, ByVal dbl����� As Double, _
                    ByVal dblָ������� As Double, ByVal dbl���� As Double, ByVal dbl���۽�� As Double, _
                    ByRef dblOut��� As Double, ByRef dblOut���� As Double, ByRef dblOut�ɱ���� As Double) As Boolean
    '------------------------------------------------------------------------------------------------------------
    ' ����:��ȡ���εĳɱ��ۺͲ��
    ' ���㹫ʽ:
    '       1.�����<=0��
    '         1) �����-ʵ�ʲ��<=0 Or dbl������� < 0
    '               a.���ĸ���������㷽ʽ=1:
    '                      a)�����ۣ�0��
    '                           ���=���۽��*ָ�������
    '                           �ɱ���=��������-�����ۣ�/����
    '                      b)������>0
    '                           �ɱ���=������
    '                           ��ۣ����۽��-����*�ɱ���
    '               b.���ĸ���������㷽ʽ<>1
    '                           ���=���۽��*ָ�������
    '                           �ɱ���=��������-�����ۣ�/����
    '          2)�����-ʵ�ʲ��>0
    '                �ɱ���= (�����-ʵ�ʲ��)/�������
    '                ��ۣ����۽��-����*�ɱ���
    '        2.�����>0
    '                   ������=������*��ʵ�ʲ��/ʵ�ʽ�
    '                  �ɱ���=��������-�����ۣ�/����
    '------------------------------------------------------------------------------------------------------------
    Dim dbl��� As Double, dbl���� As Double, dbl������� As Double
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    If dbl���� = 0 Then Exit Function
    dbl���� = Get�ɱ���(lng����ID, lng�ⷿID, lng����) * lng����ϵ��
    dbl��� = dbl���۽�� - dbl���� * dbl����
    
'    If dbl����� <= 0 Then
'        If dbl����� - dbl����� > 0 Then
'            gstrSQL = "Select (ʵ�ʽ��-ʵ�ʲ��)/ʵ������ as �ɱ��� From ҩƷ��� where �ⷿid=[1] and ҩƷid=[2] and nvl(����,0)=[3] and nvl(ʵ������,0)>0"
'            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ɱ���", lng�ⷿID, lng����ID, lng����)
'            If rsTemp.EOF = False Then
'                dbl���� = Val(NVL(rsTemp!�ɱ���)) * lng����ϵ��
'            End If
'        End If
'
'        If dbl����� - dbl����� <= 0 Or dbl���� <= 0 Then
'            If Check�����ⰴ�����ۼ��� = True Then
'                dbl���� = Get������(lng����ID) * lng����ϵ��
'                If dbl���� = 0 Then
'                    dbl��� = dbl���۽�� * dblָ�������
'                    dbl���� = (dbl���۽�� - dbl���) / Dbl����
'                Else
'                    dbl��� = dbl���۽�� - Dbl���� * dbl����
'                End If
'            Else
'                    dbl��� = dbl���۽�� * dblָ�������
'                    dbl���� = (dbl���۽�� - dbl���) / Dbl����
'            End If
'        Else
'            'dbl����� - dbl�����>0
'            dbl��� = dbl���۽�� - dbl���� * Dbl����
'        End If
'    Else
'                dbl��� = dbl���۽�� * (dbl����� / dbl�����)
'                dbl���� = (dbl���۽�� - dbl���) / Dbl����
'    End If
    
    dblOut�ɱ���� = Round(dbl���� * dbl����, 7)
    dblOut��� = Round(dbl���, 7)
    dblOut���� = Round(dbl����, 7)
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get������(ByVal lng����ID As Long) As Double
    '����:��ȡ������
    '����:lng����ID
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "Select �ɱ��� From �������� where ����id=[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ɱ���", lng����ID)
    
    If rsTemp.EOF Then
        Get������ = 0
    Else
        Get������ = Val(NVL(rsTemp!�ɱ���))
    End If
    rsTemp.Close
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ISCHECK��ǿ�ƿ���ָ���۸�() As Boolean
    '����:�ж��Ƿ�ǿ��Ҫ��������ۼ��ۼ�
     ISCHECK��ǿ�ƿ���ָ���۸� = Val(zlDatabase.GetPara(123, glngSys, 0)) = 1
End Function

Public Function ISCHECK�⹺��ǰ����() As Boolean
    '����:�ж��Ƿ�ǿ��Ҫ��������ۼ��ۼ�
    ISCHECK�⹺��ǰ���� = Val(zlDatabase.GetPara(127, glngSys, 0)) = 1
End Function
 
Public Function Check��ͨ����() As Boolean
    '------------------------------------------------------------------------------------------------------------------------------------
    '���ܣ���֤��ǰ��Ա����ͨ���ҵ������Ա
    '����:�Ƿ���true,���򷵻�false
    '------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, bln���ϲ������� As Boolean, strStock As String
    
    On Error GoTo ErrHandle
    bln���ϲ������� = Val(zlDatabase.GetPara(132, glngSys, 0)) = 1

    If bln���ϲ������� = False Then
        strStock = "K,V,12"
    Else
        strStock = "K,V,W,12"
    End If
    
    Check��ͨ���� = False
    gstrSQL = "" & _
        "   SELECT /*+ Rule*/ DISTINCT a.id, a.���� " & _
        "   FROM ��������˵�� c, �������ʷ��� b, ���ű� a " & _
        "       , Table(cast(f_Str2List([3]) as zlTools.t_StrList)) D " & _
        "   Where c.�������� = b.���� And (a.վ��=[2] or a.վ�� is null) " & _
        "       And b.����=D.Column_value " & _
        "       AND a.id = c.����id " & _
        "       AND TO_CHAR (a.����ʱ��, 'yyyy-MM-dd') = '3000-01-01' " & _
        "       And a.ID IN (Select ����ID From ������Ա Where ��ԱID=[1]) "
        
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Ա�ⷿ����", UserInfo.Id, gstrNodeNo, strStock)
    If rsTemp.EOF Then
        Check��ͨ���� = True
    Else
        Check��ͨ���� = False
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Get�ɱ���(ByVal lng����ID As Long, ByVal lng�ⷿID As Long, ByVal lng���� As Long) As Double
'���ܣ���ȡ��ǰҩƷ�ĳɱ��۸�
'������ҩƷid,�ⷿid,����
'����ֵ�� �ɱ��۸�
    Dim rsData As ADODB.Recordset
    Dim blnNullPrice As Boolean
    
    On Error GoTo ErrHandle
    
    gstrSQL = "select ƽ���ɱ��� from ҩƷ��� where ҩƷid=[1] and �ⷿid=[2] and nvl(����,0)=[3] and ����=1"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "�ɱ���", lng����ID, lng�ⷿID, lng����)
    
    If rsData.EOF Then
        blnNullPrice = True
    ElseIf IsNull(rsData!ƽ���ɱ���) = True Then
        blnNullPrice = True
    ElseIf Val(rsData!ƽ���ɱ���) < 0 Then
        blnNullPrice = True
    End If
    
    If Not blnNullPrice Then
        Get�ɱ��� = rsData!ƽ���ɱ���
    Else
        '����޷��ӿ����ȡ�ɱ��ۣ���Ӳ���������ȡ
        gstrSQL = "select �ɱ��� from �������� where ����id=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "�ɱ���", lng����ID)
        If Not rsData.EOF Then
            If Val(NVL(rsData!�ɱ���, 0)) > 0 Then
                Get�ɱ��� = rsData!�ɱ���
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

Public Function Get�ۼ�(ByVal bln�Ƿ�ʱ�� As Boolean, lng����ID As Long, ByVal lng�ⷿID As Long, ByVal lng���� As Long) As Double
    '���ܣ���ȡԭʼ���ۼ۵�λ�ۼۣ���Ҫ���ڳ���
    '����: bln�Ƿ�ʱ��:false-����,true-ʱ��
    '����ֵ����С��λ�ļ۸�
    Dim rsData As ADODB.Recordset
    Dim dbl���ۼ� As Double, dblָ�����ۼ� As Double, dbl��������� As Double, dbl�ӳ��� As Double
    Dim dbl�ɱ��� As Double
    
    On Error GoTo ErrHandle

    'ȡ����ҩƷ�ۼ�
    If bln�Ƿ�ʱ�� = False Then
        gstrSQL = "Select a.�ּ� " & _
            " From �շѼ�Ŀ A " & _
            " Where A.�շ�ϸĿID=[1] And Sysdate Between A.ִ������ And Nvl(A.��ֹ����,Sysdate) "
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "Get�ۼ�-ȡ����ҩƷ�ۼ�", lng����ID)
        
        If Not rsData.EOF Then
            Get�ۼ� = rsData!�ּ�
        End If
    Else
        'ȡʱ��ҩƷ�ۼ�
        gstrSQL = "select Decode(���ۼ�, Null, Decode(Nvl(ʵ������, 0), 0, 0, ʵ�ʽ�� / ʵ������), ���ۼ�) as ���ۼ� " & _
            " from ҩƷ��� where ����=1 and  ҩƷid=[1] and �ⷿid=[2] and nvl(����,0)=[3]"
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "GetOriPrice-���ۼ�", lng����ID, lng�ⷿID, lng����)
        
        If rsData.EOF Then
            
            '�����ݣ��ӹ����ȡ���һ�μ۸�
            gstrSQL = "Select �ϴ��ۼ�,ָ�����ۼ�,nvl(ָ�������,0) as ָ�������,nvl(�ӳ���,0) as �ӳ���,Nvl(���������,100) ��������� From �������� Where ����ID=[1]"
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "���ۼ�", lng����ID)
            
            If IsNull(rsData!�ϴ��ۼ�) Then
                dblָ�����ۼ� = rsData!ָ�����ۼ�
                dbl��������� = rsData!���������
                
                Get�ۼ� = 0
                dbl�ɱ��� = Get�ɱ���(lng����ID, lng�ⷿID, lng����)
                dbl�ӳ��� = rsData!�ӳ��� / 100
                dbl���ۼ� = dbl�ɱ��� * (1 + dbl�ӳ���)
                dbl���ۼ� = dbl���ۼ� + (dblָ�����ۼ� - dbl���ۼ�) * (1 - dbl��������� / 100)
                Get�ۼ� = IIf(dbl���ۼ� > dblָ�����ۼ�, dblָ�����ۼ�, dbl���ۼ�)
            Else
                Get�ۼ� = rsData!�ϴ��ۼ�
            End If
        Else '���������
            If rsData!���ۼ� < 0 Then
                gstrSQL = "Select �ϴ��ۼ�,ָ�����ۼ�,nvl(ָ�������,0) as ָ�������,nvl(�ӳ���,0) as �ӳ���,Nvl(���������,100) ��������� From �������� Where ����ID=[1]"
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "���ۼ�", lng����ID)
                
                If IsNull(rsData!�ϴ��ۼ�) Then
                    dblָ�����ۼ� = rsData!ָ�����ۼ�
                    dbl��������� = rsData!���������
                    
                    Get�ۼ� = 0
                    dbl�ɱ��� = Get�ɱ���(lng����ID, lng�ⷿID, lng����)
                    dbl�ӳ��� = rsData!�ӳ��� / 100
                    dbl���ۼ� = dbl�ɱ��� * (1 + dbl�ӳ���)
                    dbl���ۼ� = dbl���ۼ� + (dblָ�����ۼ� - dbl���ۼ�) * (1 - dbl��������� / 100)
                    Get�ۼ� = IIf(dbl���ۼ� > dblָ�����ۼ�, dblָ�����ۼ�, dbl���ۼ�)
                Else
                    Get�ۼ� = rsData!�ϴ��ۼ�
                End If
            Else
                Get�ۼ� = rsData!���ۼ�
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

Public Function Get���ۼ�(ByVal lng����ID As Long, ByVal lng�ⷿID As Long, ByVal lng���� As Long, ByVal dbl����ϵ�� As Double) As Double
    '���ܣ���ȡʱ��ҩƷ��ǰҩƷ�����ۼ�
    '����:ҩƷid,�ⷿid,����
    '����ֵ�����ۼ�
    Dim rsData As ADODB.Recordset
    Dim dbl���ۼ� As Double, dblָ�����ۼ� As Double, dbl��������� As Double, dbl�ӳ��� As Double
    Dim dbl�ɱ��� As Double
    
    On Error GoTo ErrHandle
    If lng���� <> 0 Then
        gstrSQL = "select ���ۼ� from ҩƷ��� where ҩƷid=[1] and �ⷿid=[2] and nvl(����,0)=[3] and ����=1"
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "���ۼ�", lng����ID, lng�ⷿID, lng����)
    Else
        gstrSQL = "Select ʵ�ʽ�� / ʵ������ As ���ۼ�" & vbNewLine & _
                "   From ҩƷ���" & vbNewLine & _
                "   Where �ⷿid = [2] And ҩƷid = [1] And ���� = 1 And ʵ������ > 0"
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "���ۼ�", lng����ID, lng�ⷿID)
    End If
    
    If rsData.EOF Or IsNull(rsData!���ۼ�) = True Then
        'ʱ��ҩƷ���ۼۼ��㹫ʽ:�ɹ���*(1+�ӳ���)
        '��Ϊ:�ɹ���*(1+�ӳ���)+(ָ�����ۼ�-�ɹ���*(1+�ӳ���))*(1-���������)
        '���ڲ�������ȵĴ���,��ǰ���а�ָ������ʼ���ĵط�,����Ҫ�������ת���ɼӳ��ʽ��м���,�˺������ڷ��ر��ι�ʽ���ӵĲ��ֽ�(ָ�����ۼ�-�ɹ���*(1+�ӳ���))*(1-���������)
        gstrSQL = "Select �ϴ��ۼ�,ָ�����ۼ�,nvl(ָ�������,0) as ָ�������,nvl(�ӳ���,0) as �ӳ���,Nvl(���������,100) ��������� From �������� Where ����ID=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "���ۼ�", lng����ID)
        
        If IsNull(rsData!�ϴ��ۼ�) Then
            dblָ�����ۼ� = rsData!ָ�����ۼ�
            dbl��������� = rsData!���������
            
            Get���ۼ� = 0
            dbl�ɱ��� = Get�ɱ���(lng����ID, lng�ⷿID, lng����)
            dbl�ӳ��� = rsData!�ӳ��� / 100
            dbl���ۼ� = dbl�ɱ��� * (1 + dbl�ӳ���)
            dbl���ۼ� = dbl���ۼ� + (dblָ�����ۼ� - dbl���ۼ�) * (1 - dbl��������� / 100)
            Get���ۼ� = IIf(dbl���ۼ� > dblָ�����ۼ�, dblָ�����ۼ�, dbl���ۼ�) * dbl����ϵ��
        Else
            Get���ۼ� = rsData!�ϴ��ۼ� * dbl����ϵ��
        End If
    Else
        Get���ۼ� = rsData!���ۼ� * dbl����ϵ��
    End If
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function ZVal(ByVal varValue As Variant, Optional ByVal blnForceNum As Boolean) As String
'���ܣ���0��ת��Ϊ"NULL"��,������SQL���ʱ��
'������blnForceNum=��ΪNullʱ���Ƿ�ǿ�Ʊ�ʾΪ������
    ZVal = IIf(Val(varValue) = 0, IIf(blnForceNum, "-NULL", "NULL"), Val(varValue))
End Function
Public Function GetControlRect(ByVal lngHwnd As Long) As RECT
'���ܣ���ȡָ���ؼ�����Ļ�е�λ��(Twip)
    Dim vRect As RECT
    Call GetWindowRect(lngHwnd, vRect)
    vRect.Left = vRect.Left * Screen.TwipsPerPixelX
    vRect.Right = vRect.Right * Screen.TwipsPerPixelX
    vRect.Top = vRect.Top * Screen.TwipsPerPixelY
    vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    GetControlRect = vRect
End Function
 
Public Function CheckPrint(ByVal strNo As String, ByVal int���� As Integer, lng�ⷿID As Long, Optional lng�Է�����ID As Long = 0, Optional bln�Է��ⷿ As Boolean = False) As Boolean
    '����Ƿ��Ѿ���ӡ���������Ƿ������ӡ�������򷵻�true�����򷵻�false
    Dim rsTemp As ADODB.Recordset
    
    If bln�Է��ⷿ = False Then
        gstrSQL = "Select 1 From ҩƷ�շ����� Where NO = [1] And ���� = [2] And �ⷿid = [3]"
    Else
        gstrSQL = "Select 1 From ҩƷ�շ����� Where NO = [1] And ���� = [2] And �Է�����id = [4]"
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ��Ѿ���ӡ����", strNo, int����, lng�ⷿID, lng�Է�����ID)
    If rsTemp.RecordCount > 0 Then
        If MsgBox("�˵����Ѿ���ӡ�����Ƿ������ӡ��", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            CheckPrint = False
            Exit Function
        Else
            CheckPrint = True
        End If
    Else
        gstrSQL = "Zl_ҩƷ�շ�����_Insert("
        gstrSQL = gstrSQL & "'" & strNo & "'"
        gstrSQL = gstrSQL & "," & int����
        gstrSQL = gstrSQL & "," & lng�ⷿID
        gstrSQL = gstrSQL & "," & lng�Է�����ID & ")"
        zlDatabase.ExecuteProcedure gstrSQL, "���ݴ�ӡ"
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
    '���ܣ� ����X,Y��ʵ�����꣬��������Ļ���������
    '������ X---���غ��������
    '       Y---�������������
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


Public Function ReturnParaData(ByVal lngSys As Long, ByVal str������IN As String) As ADODB.Recordset
    '-------------------------------------------------------------------------------------------
    '����:��ȡָ���Ĳ���ֵ,����һ����¼��
    '����:lngSys-ϵͳ
    '     str������IN-������In,�Զ��ŷ���
    '
    '����:������¼��
    '����:���˺�
    '����:2007/12/17
    '-------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    strSQL = "" & _
        "   Select  /*+ Rule*/ ������,nvl(����ֵ,ȱʡֵ) as ����ֵ,����˵�� " & _
        "   From zlParameters A,Table(Cast(f_Num2List([2]) As zlTools.t_NumList)) B" & _
        "   where A.������ = B.Column_Value and a.ϵͳ=[1] and nvl(A.˽��,0)=0 and nvl(a.ģ��,0)=0  " & _
        "   order by ������"
        
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����ֵ", lngSys, str������IN)
    
    Set ReturnParaData = rsTemp
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
'ȡ�ܣ��£��������꣬��ĵ�һ��
Public Function GetFirstDate(ByVal intInteval As Integer, ByVal datCurrent As Date) As Date
    Dim datReturn As Date
    
    Select Case intInteval
        Case FirstDayOfWeek       '��ǰ�ܵĵ�һ��
            datReturn = DateAdd("d", -Weekday(datCurrent) + 1, Now)
        Case FirstDayOfMonth       '��ǰ�µĵ�һ��
            datReturn = DateAdd("d", -Day(datCurrent) + 1, datCurrent)
        Case FirstDayOfQuarter       '��ǰ���ĵ�һ��
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
        Case FirstDayOfHalfYear       '��ǰ����ĵ�һ��
            If Month(datCurrent) > 6 Then
                datReturn = DateSerial(Year(datCurrent), 7, 1)
            Else
                datReturn = DateSerial(Year(datCurrent), 1, 1)
            End If
        Case FirstDayOfyear       '��ǰ��ĵ�һ��
            datReturn = DateSerial(Year(datCurrent), 1, 1)
    End Select
    GetFirstDate = datReturn
End Function



Public Function Check��������(ByVal lng�ⷿID As Long, ByVal lng����ID As Long, ByVal lng���� As Long, _
    ByVal dbl�������� As Double, ByVal int����� As Integer, Optional ByVal intType As Integer = 0) As Boolean
    '------------------------------------------------------------------------------
    '����:���������ʱ�Ŀ��������Ƿ��㹻
    '����:���㷵�ط���true,���򷵻�False
    '����:
    '    int�����:0-�����;1-��飬��������,2-��飬�����ֹ
    '����:���˺�
    '����:2008/02/15
    '------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, dbl���� As Double
    
    err = 0: On Error GoTo ErrHand:
    '0-�����
    If int����� = 0 Then Check�������� = True: Exit Function
    
    gstrSQL = "Select A.��������,A.ʵ������,B.����,B.���� From ҩƷ��� A,�շ���ĿĿ¼ B where A.ҩƷid=B.id And A.ҩƷid=[1] and A.�ⷿid=[2] and nvl(A.����,0)=[3] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�����ÿɴ�", lng����ID, lng�ⷿID, lng����)
    
    If rsTemp.EOF Then
        dbl���� = 0
        gstrSQL = "Select 0 as ��������,B.����,B.���� From �շ���ĿĿ¼ B where B.id=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�����ÿɴ�", lng����ID, lng�ⷿID, lng����)
        If rsTemp.EOF Then ShowMsgBox "ָ�����������ϲ�����,����!": Exit Function
    Else
        If intType = 0 Then
            dbl���� = Round(Val(zlStr.NVL(rsTemp!��������, 0)), g_С��λ��.obj_���С��.����С��)
        Else
            dbl���� = Round(Val(zlStr.NVL(rsTemp!ʵ������, 0)), g_С��λ��.obj_���С��.����С��)
        End If
    End If
    
    If dbl���� < Round(dbl��������, g_С��λ��.obj_���С��.����С��) Then
        If intType = 0 Then
            If int����� = 1 Then
                '1-��飬��������
                If MsgBox("��[" & rsTemp!���� & "]" & zlStr.NVL(rsTemp!����) & "���Ŀ��ÿ�治�㣬�Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            Else
                '2-��飬�����ֹ
                ShowMsgBox "��[" & rsTemp!���� & "]" & zlStr.NVL(rsTemp!����) & "���Ŀ��ÿ�治�㣬���ܼ�����"
                Exit Function
            End If
        Else
            If int����� = 1 Then
                '1-��飬��������
                If MsgBox("��[" & rsTemp!���� & "]" & zlStr.NVL(rsTemp!����) & "����ʵ�ʿ�治�㣬�Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            Else
                '2-��飬�����ֹ
                ShowMsgBox "��[" & rsTemp!���� & "]" & zlStr.NVL(rsTemp!����) & "����ʵ�ʿ�治�㣬���ܼ�����"
                Exit Function
            End If
        End If
    End If
    
    Check�������� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function
Public Function ȡ��������(ByVal int���� As Integer, _
    ByVal strNo As String, _
    lng����ID As Long, int��� As Integer, Optional lng���ϵ�� As Long = 1) As Long
    '------------------------------------------------------------------------------
    '����:��ȡ��������
    '����:����ָ���е�����
    '����:���˺�
    '����:2008/02/15
    '------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    err = 0: On Error GoTo ErrHand:
    gstrSQL = "Select Nvl(����, 0) ���� From ҩƷ�շ���¼ Where ���� = [1] And NO = [2] And ��� = [3] And ҩƷid = [4] And ���ϵ�� = [5]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�������", int����, strNo, int���, lng����ID, lng���ϵ��)
    If rsTemp.EOF Then
        ȡ�������� = 0
    Else
        ȡ�������� = Val(NVL(rsTemp!����))
    End If
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog

End Function

 
  
Public Function SelectItem(ByVal frmMain As Form, ByVal objCtl As Control, ByVal strKey As String, _
    ByVal strTable As String, ByVal strTittle As String, Optional ByVal blnNotMsg As Boolean = False) As Boolean
    '------------------------------------------------------------------------------
    '����:�๦��ѡ����
    '����:objCtl-�ı���ؼ�
    '     strKey-Ҫ������ֵ
    '     strTable-����
    '     strTittle-ѡ��������
    '     blnNotMsg-����ʾ.
    '����:
    '����:���˺�
    '����:2008/02/18
    '------------------------------------------------------------------------------
    Dim blnCancel As Boolean, lngH As Long
    Dim vRect As RECT, sngX As Single, sngY As Single
    Dim rsTemp  As ADODB.Recordset
    'zlDatabase.ShowSelect
    '���ܣ��๦��ѡ����
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    
    gstrSQL = "Select rownum as ID,a.* From " & strTable & " a"
    
    If strKey <> "" Then
        gstrSQL = gstrSQL & _
        "   Where ((����) like [1] or  ����  like [1] or  ����  like  upper([1]))  " & _
        "    "
    End If
    gstrSQL = gstrSQL & _
    "   order by ����"
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
            ShowMsgBox "û���ҵ���������������,����!"
        End If
        Call zlCtlSetFocus(objCtl, True)
        Exit Function
    End If
    
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Then
        With objCtl
            .TextMatrix(.Row, .Col) = NVL(rsTemp!����) & "-" & NVL(rsTemp!����)
            .Cell(flexcpData, .Row, .Col) = NVL(rsTemp!����)
        End With
    Else
        Call zlCtlSetFocus(objCtl, True)
        objCtl.Text = NVL(rsTemp!����)
        objCtl.Tag = NVL(rsTemp!����)
        zlCommFun.PressKey vbKeyTab
    End If
    SelectItem = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub zlCtlSetFocus(ByVal objCtl As Object, Optional blnDoEvnts As Boolean = False)
    '����:�������ƶ��ؼ���:2008-07-08 16:48:35
    err = 0: On Error Resume Next
    If blnDoEvnts Then DoEvents
    If IsCtrlSetFocus(objCtl) = True Then: objCtl.SetFocus
End Sub

Public Function Select����ѡ����(ByVal frmMain As Form, ByVal objCtl As Control, ByVal strSearch As String, _
    Optional str�������� As String = "", _
    Optional bln����Ա As Boolean = False, _
    Optional strSQL As String = "") As Boolean
    '------------------------------------------------------------------------------
    '����:����ѡ����
    '����:objCtl-ָ���ؼ�
    '     strSearch-Ҫ����������
    '     str��������-��������:��"V,W,K"
    '     bln����Ա-�Ƿ�Ӳ���Ա����
    '     strSQL-ֱ�Ӹ���SQL��ȡ����(�����ű�ı���һ��Ҫ��A)
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008/02/18
    '------------------------------------------------------------------------------
    Dim i As Long
    Dim blnCancel As Boolean, strKey As String, strTittle As String, lngH As Long, strFind As String
    Dim vRect As RECT
    Dim rsTemp  As ADODB.Recordset
    'zlDatabase.ShowSelect
    '���ܣ��๦��ѡ����
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    
    strTittle = "����ѡ����"
    vRect = GetControlRect(objCtl.hwnd)
    lngH = objCtl.Height
    
    strKey = GetMatchingSting(strSearch, False)
    
    If strSQL <> "" Then
    
        gstrSQL = strSQL
    Else
'        gstrSQL = "" & _
'        "   Select /*+ Rule*/ distinct a.Id,a.�ϼ�id,a.����,a.����,a.����,a.λ�� ,To_Char(a.����ʱ��, 'yyyy-mm-dd') As ����ʱ��, " & _
'        "          decode(To_Char(a.����ʱ��, 'yyyy-mm-dd'),'3000-01-01','',To_Char(a.����ʱ��, 'yyyy-mm-dd')) ����ʱ��"
    
        If str�������� = "" And bln����Ա = False Then
            gstrSQL = "" & _
            "   Select a.Id,a.�ϼ�id,a.����,a.����,a.����,a.λ�� ,To_Char(a.����ʱ��, 'yyyy-mm-dd') As ����ʱ��, " & _
            "          decode(To_Char(a.����ʱ��, 'yyyy-mm-dd'),'3000-01-01','',To_Char(a.����ʱ��, 'yyyy-mm-dd')) ����ʱ��"
        
            gstrSQL = gstrSQL & vbCrLf & _
            "   From ���ű� a" & _
            "   Where 1=1"
        Else
            gstrSQL = "" & _
            "   Select /*+ Rule*/ distinct a.Id,a.�ϼ�id,a.����,a.����,a.����,a.λ�� ,To_Char(a.����ʱ��, 'yyyy-mm-dd') As ����ʱ��, " & _
            "          decode(To_Char(a.����ʱ��, 'yyyy-mm-dd'),'3000-01-01','',To_Char(a.����ʱ��, 'yyyy-mm-dd')) ����ʱ��"
        
            gstrSQL = gstrSQL & vbCrLf & _
            "   From ���ű� a, �������ʷ��� b,��������˵�� c," & _
            IIf(str�������� = "", "", "       (Select Column_Value From Table(Cast(f_Str2list([2]) As zlTools.t_Strlist))) J") & _
            "   Where c.�������� = b.����" & IIf(str�������� = "", "(+)", " and B.����=J.column_value ") & _
            "         AND a.id = c.����id " & _
            IIf(bln����Ա = False, "", " And a.ID IN (Select ����ID From ������Ա Where ��ԱID=[1])")
        End If
        gstrSQL = gstrSQL & vbCrLf & _
            "   and  (a.����ʱ��>=to_date('3000-01-01','yyyy-mm-dd') or a.����ʱ�� is null ) And (a.վ��=[4] or a.վ�� is null) "
    End If
    
    strFind = ""
    If strSearch <> "" Then
        strFind = "   and  (a.���� like upper([3]) or a.���� like upper([3]) or a.���� like [3] )"
        If IsNumeric(strSearch) Then                         '���������,��ֻȡ����
            If Mid(gSystem_Para.Para_���뷽ʽ, 1, 1) = "1" Then strFind = " And (A.���� Like Upper([3]))"
        ElseIf zlCommFun.IsCharAlpha(strSearch) Then           '01,11.����ȫ����ĸʱֻƥ�����
            '0-ƴ����,1-�����,2-����
            '.int���뷽ʽ = Val(zlDatabase.GetPara("���뷽ʽ" ))
            If Mid(gSystem_Para.Para_���뷽ʽ, 2, 1) = "1" Then strFind = " And  (a.���� Like Upper([3]))"
        ElseIf zlCommFun.IsCharChinese(strSearch) Then  'ȫ����
            strFind = " And a.���� Like [3] "
        End If
    End If
    
    If strSearch = "" And str�������� = "" And bln����Ա = False And strSQL = "" Then
        gstrSQL = gstrSQL & _
        "   Start With A.�ϼ�id Is Null Connect By Prior A.ID = A.�ϼ�id "
    Else
        gstrSQL = gstrSQL & vbCrLf & strFind & vbCrLf & " Order by A.����"
    End If
    
    If strSearch = "" And str�������� = "" And bln����Ա = False And strSQL = "" Then
        '�����¼�
        Set rsTemp = zlDatabase.ShowSQLSelect(frmMain, gstrSQL, 1, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, UserInfo.Id, str��������, strKey, gstrNodeNo)
    Else
        Set rsTemp = zlDatabase.ShowSQLSelect(frmMain, gstrSQL, 0, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, UserInfo.Id, str��������, strKey, gstrNodeNo)
    End If
    If blnCancel = True Then
        Call zlCtlSetFocus(objCtl, True)
        Exit Function
    End If
    If rsTemp Is Nothing Then
        ShowMsgBox "û�����������Ĳ���,����!"
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
            ShowMsgBox "��ѡ��Ĳ����������б��в�����,����!"
            If objCtl.Enabled Then objCtl.SetFocus
            Exit Function
        End If
    Else
        objCtl.Text = NVL(rsTemp!����) & "-" & NVL(rsTemp!����)
        objCtl.Tag = Val(rsTemp!Id)
    End If
    zlCommFun.PressKey vbKeyTab
    Select����ѡ���� = True
End Function
  
Public Function zlDblIsValid(ByVal StrInput As String, ByVal intMax As Integer, Optional bln������� As Boolean = True, Optional bln���� As Boolean = True, _
        Optional ByVal hwnd As Long = 0, Optional str��Ŀ As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����ַ����Ƿ�Ϸ��Ľ��
    '���:strInput        ������ַ���
    '     intMax          ������λ��
    '     bln�������     �Ƿ���и������
    '     bln����         �Ƿ������ļ��
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-10-20 15:16:08
    '-----------------------------------------------------------------------------------------------------------
   
    Dim dblValue As Double
    If bln���� = True Then
        If StrInput = "" Then
            ShowMsgBox str��Ŀ & "δ���룬����!"
            If hwnd <> 0 Then SetFocusHwnd hwnd
            Exit Function
        End If
    End If
    If StrInput = "" Then zlDblIsValid = True: Exit Function
    
    If IsNumeric(StrInput) = False Then
        MsgBox str��Ŀ & "������Ч�����ָ�ʽ��", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '���ý���
        Exit Function
    End If
    
    dblValue = Val(StrInput)
    If dblValue >= 10 ^ intMax - 1 Then
        MsgBox str��Ŀ & "��ֵ���󣬲��ܳ���" & 10 ^ intMax - 1 & "��", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '���ý���
        Exit Function
    End If
    If bln������� = True And dblValue < 0 Then
        MsgBox str��Ŀ & "�������븺����", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '���ý���
        Exit Function
    End If
    
    If Abs(dblValue) >= 10 ^ intMax And dblValue < 0 Then
        MsgBox str��Ŀ & "��ֵ��С������С��-" & 10 ^ intMax - 1 & "λ��", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '���ý���
        Exit Function
    End If
    
    
    If bln���� = True And dblValue = 0 Then
        MsgBox str��Ŀ & "���������㡣", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '���ý���
        Exit Function
    End If
    zlDblIsValid = True
End Function

Public Function zlCheckIsDate(ByVal strKey As String, ByVal strTittle As String) As String
    '------------------------------------------------------------------------------
    '����:����Ƿ�Ϸ���������,����Ϊ:(20070101��2007-01-01)����(01-01��0101)����(01<01-31>)
    '����:strKey-��Ҫ���Ĺؽ���
    '����:�Ϸ�������,���ر�׼��ʽ(yyyy-mm-dd),���򷵻�""
    '����:���˺�
    '����:2008/01/24
    '------------------------------------------------------------------------------
    If Len(strKey) = 4 And InStr(1, strKey, "-") = 0 Then
        '0101,��Ҫ��ǰ�����
        strKey = Year(Now) & strKey
    ElseIf Len(Replace(strKey, "-", "")) = 4 And InStr(1, strKey, "-") > 0 Then
        '01-01��ʽ,��Ҫ����
        strKey = Year(Now) & Replace(strKey, "-", "")
    ElseIf Len(strKey) <= 2 And IsNumeric(strKey) Then
        'ָ����
        strKey = Format(Now, "YYYYMM") & IIf(Len(strKey) = 2, strKey, "0" & strKey)
    End If
    If Len(strKey) = 8 And InStr(1, strKey, "-") = 0 Then
        strKey = TranNumToDate(strKey)
        If strKey = "" Then
            ShowMsgBox strTittle & "����Ϊ������,���飡"
            Exit Function
        End If
    End If
    If Not IsDate(strKey) Then
        ShowMsgBox strTittle & "����Ϊ��������(2000-10-10) ��20001010��,���飡"
        Exit Function
    End If
    zlCheckIsDate = strKey
End Function

Public Function zl����δ��˵���(ByVal lng����ID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����Ƿ����δ��˵ĵ���
    '���:
    '����:
    '����:����,����true,���򷵻�False
    '����:���˺�
    '����:2008-11-07 15:33:14
    '-----------------------------------------------------------------------------------------------------------

    '���ҩƷ�Ƿ����δ��˵���
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "Select 1 From ҩƷ�շ���¼ Where ҩƷid = [1] And Rownum = 1 And ������� Is Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "������������Ƿ����δ��˵���", lng����ID)
    zl����δ��˵��� = rsTemp.RecordCount <> 0
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function Select��Ӧ��(ByVal frmMain As Form, ByVal objCtl As Control, Optional ByVal strSearch As String = "") As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:��Ӧ��ѡ��
    '���:frmMain-���õ�������
    '    objCtl-���õĿؼ�
    '    strSearch-��������(""��ʾ����ѡ��)
    '����:
    '����:ѡ��ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-11-10 10:38:26
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As Recordset, strKey As String
    Dim blnCancel As Boolean, lngH As Long
    Dim vRect As RECT, sngX As Single, sngY As Single
    Dim bytStyle As Byte, blnĩ�� As Boolean
    
    
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
    
    
 
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    gstrSQL = "" & _
        "   Select id,�ϼ�ID,����, ����, ����, ĩ��, ���֤��, ���֤Ч��, ִ�պ�, ִ��Ч��, ˰��ǼǺ�, ��ַ, �绰, ��������," & _
        "           �ʺ�, ��ϵ��, ����, ������, ���ö�, ����ί����, to_char(����ί������,'yyyy-mm-dd') as ����ί������, ������֤��, to_char(������֤����,'yyyy-mm-dd') as ������֤����," & _
        "           ҩ��ֱ�����, to_char(ҩ��ֱ�������,'yyyy-mm-dd') as ҩ��ֱ�������, ��Ȩ��, ��Ȩ��, վ��," & _
        "           to_char(����ʱ��,'yyyy-mm-dd') as ����ʱ��, decode(To_Char(����ʱ��,'yyyy-MM-dd'),'3000-01-01','', to_char(����ʱ��,'yyyy-mm-dd')) as ����ʱ��" & _
        "   From ��Ӧ�� " & _
        "   Where  (To_Char(����ʱ��,'yyyy-MM-dd')='3000-01-01' or ����ʱ�� is null)  "
    If strSearch = "" Then
        gstrSQL = gstrSQL & _
            "           And (substr(����,5,1)=1 And (վ��=[2] or վ�� is null) Or Nvl(ĩ��,0)=0) " & _
            "   Start with �ϼ�ID is null connect by prior ID =�ϼ�ID " & _
            "   Order by level,ID"
        blnĩ�� = True
        bytStyle = 2
    Else
        gstrSQL = gstrSQL & _
            "    And (վ��=[2] or վ�� is null) And ĩ��=1 And substr(����,5,1)=1 " & _
            "    And (���� like upper([1]) Or ���� like [1] or ���� like [1]) "
        bytStyle = 0
        blnĩ�� = False
    End If
    Set rsTemp = zlDatabase.ShowSQLSelect(frmMain, gstrSQL, bytStyle, "��Ӧ��ѡ����", Not blnĩ��, "", "��ѡ������������ϵĹ�Ӧ��", False, True, Not blnĩ��, sngX, sngY, lngH, blnCancel, False, False, strKey, gstrNodeNo)
        
    If blnCancel Then
        Call zlCtlSetFocus(objCtl, True)
        Exit Function
    End If
    If rsTemp Is Nothing Then
        ShowMsgBox "û���ҵ����������Ĺ�Ӧ��,����!"
        Call zlCtlSetFocus(objCtl, True)
        Exit Function
    End If
    
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Then
        With objCtl
            .TextMatrix(.Row, .Col) = NVL(rsTemp!����) & "-" & NVL(rsTemp!����)
            .Cell(flexcpData, .Row, .Col) = NVL(rsTemp!Id)
        End With
    Else
        Call zlCtlSetFocus(objCtl, True)
        objCtl.Text = NVL(rsTemp!����)
        objCtl.Tag = NVL(rsTemp!Id)
        zlCommFun.PressKey vbKeyTab
    End If
    Select��Ӧ�� = True
End Function

'�����룬���ƣ���������ĳһ��
Public Function FindVsRowNew(ByVal vsBill As VSFlexGrid, ByVal int�Ƚ��� As Integer, _
    ByVal str�Ƚ�ֵ As String, ByVal blnFirst As Boolean) As Boolean
    Dim intStartRow As Integer
    Dim intRow As Integer
    Dim strSpell As String
    Dim strCode As String
    Dim rsCode As New Recordset
    
    On Error GoTo ErrHandle
    FindVsRowNew = True
    With vsBill
        If .Rows = 2 Then Exit Function
        If str�Ƚ�ֵ = "" Then Exit Function
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
            If .TextMatrix(intRow, int�Ƚ���) <> "" Then
                strCode = .TextMatrix(intRow, int�Ƚ���)
                If InStr(1, UCase(strCode), UCase(str�Ƚ�ֵ)) <> 0 Then
                    .SetFocus
                    .Row = intRow
                    .Col = int�Ƚ���
                    .TopRow = .Row
                    Exit Function
                End If
            End If
        Next
        
        gstrSQL = "" & _
        " SELECT DISTINCT b.���� " & _
        " FROM (    SELECT DISTINCT A.�շ�ϸĿid " & _
        "           FROM �շ���Ŀ���� A" & _
        "           Where A.���� LIKE upper([1]) " & _
        "       ) a, �շ���ĿĿ¼ B " & _
        " Where a.�շ�ϸĿid = b.ID And (b.վ��=[2] or b.վ�� is null) "
        
        Set rsCode = zlDatabase.OpenSQLRecord(gstrSQL, "����ָ����������", GetMatchingSting(str�Ƚ�ֵ, False), gstrNodeNo)
        If rsCode.EOF Then
            FindVsRowNew = False
            Exit Function
        End If
        
        For intRow = intStartRow To .Rows - 1
            If .TextMatrix(intRow, int�Ƚ���) <> "" Then
                strCode = .TextMatrix(intRow, int�Ƚ���)
                rsCode.MoveFirst
                Do While Not rsCode.EOF
                    If InStr(1, UCase(strCode), UCase(rsCode!����)) <> 0 Then
                        .SetFocus
                        .Row = intRow
                        .Col = int�Ƚ���
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
    Optional blnδ�ҵ����� As Boolean = False, Optional strOra���� As String, Optional strWhere As String) As Boolean
    '------------------------------------------------------------------------------
    '����:�๦��ѡ����
    '����:objCtl-�ı���ؼ�
    '     strKey-Ҫ������ֵ
    '     strTable-����
    '     strTittle-ѡ��������
    '����:
    '����:���˺�
    '����:2008/02/18
    '------------------------------------------------------------------------------
    Dim blnCancel As Boolean, lngH As Long, str���� As String, str���� As String
    Dim vRect As RECT, sngX As Single, sngY As Single
    Dim rsTemp  As ADODB.Recordset
    'zlDatabase.ShowSelect
    '���ܣ��๦��ѡ����
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    str���� = strKey
    
    gstrSQL = "Select rownum as ID,a.* From " & strTable & " a where 1=1 "
    
    If strKey <> "" Then
        gstrSQL = gstrSQL & _
        "   And ((����) like [1] or  ����  like [1] or  ����  like  upper([1]))  " & _
        "    "
    End If
    gstrSQL = gstrSQL & strWhere & _
    "   order by ����"
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
        
        If blnδ�ҵ����� Then
            If zlCommFun.IsCharChinese(str����) = False Then GoTo NOAdd::
            If MsgBox("ע��:" & vbCrLf & _
                   "     δ�ҵ���ص�" & strTable & ",�Ƿ����ӡ�" & str���� & "����", vbQuestion + vbYesNo + vbDefaultButton2, strTable) = vbNo Then
                If objCtl.Enabled Then objCtl.SetFocus
                Exit Function
            End If
            
            If AutoAddBaseItem(strTable, str����, str����, strTable & "����", False) = False Then
                If objCtl.Enabled Then objCtl.SetFocus
                Exit Function
            End If
            
            If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Or UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
                With objCtl
                    .TextMatrix(.Row, .Col) = IIf(blnOnlyName, str����, str���� & "-" & str����)
                    If Not (UCase(TypeName(objCtl)) = UCase("BILLEDIT")) Then
                        .Cell(flexcpData, .Row, .Col) = IIf(blnOnlyName, str����, str���� & "-" & str����)
                    End If
                End With
            Else
                If IsCtrlSetFocus(objCtl) Then objCtl.SetFocus
                objCtl.Text = IIf(blnOnlyName, str����, str���� & "-" & str����)
                objCtl.Tag = str����
                zlCommFun.PressKey vbKeyTab
            End If
            SelectAndNotAddItem = True
            Exit Function
        Else
NOAdd:
            ShowMsgBox "û���ҵ�����������" & strTable & ",����!"
            If objCtl.Enabled Then objCtl.SetFocus
            Exit Function
        End If
    End If
    If UCase(TypeName(objCtl)) = UCase("VSFlexGrid") Or UCase(TypeName(objCtl)) = UCase("BILLEDIT") Then
        With objCtl
            .TextMatrix(.Row, .Col) = IIf(blnOnlyName, NVL(rsTemp!����), NVL(rsTemp!����) & "-" & NVL(rsTemp!����))
            If Not (UCase(TypeName(objCtl)) = UCase("BILLEDIT")) Then
                .EditText = NVL(rsTemp!����)
                .Cell(flexcpData, .Row, .Col) = NVL(rsTemp!����)
            End If
        End With
    Else
        If IsCtrlSetFocus(objCtl) Then objCtl.SetFocus
        objCtl.Text = NVL(rsTemp!����)
        objCtl.Tag = NVL(rsTemp!����)
        zlCommFun.PressKey vbKeyTab
    End If
    SelectAndNotAddItem = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function




Public Function AutoAddBaseItem(ByVal strTable As String, str���� As String, str���� As String, _
    Optional strTittle As String = "������Ŀ", Optional blnMsg As Boolean = False) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--��  ��:�Զ�������Ŀ��Ϣ(ֻ����б���,���Ƶ���Ϣ����(ֻ���ӣ����������,����)
    '--�����:
    '--������:
    '--��  ��:���ӳɹ�,����true,���򷵻�false
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New Recordset
    Dim int���� As Integer, strCode As String, strSpecify As String
    AutoAddBaseItem = False
    If blnMsg = True Then
        If MsgBox("û���ҵ��������" & strTable & "����Ҫ��������" & strTable & "����", vbYesNo + vbQuestion, strTittle) = vbNo Then
            Exit Function
        End If
    End If
    
    err = 0: On Error GoTo ErrHand:
    
    gstrSQL = "SELECT Nvl(MAX(LENGTH(����)), 2) As Length FROM  " & strTable
    zlDatabase.OpenRecordset rsTemp, gstrSQL, strTittle
    
    int���� = rsTemp!Length
    
    gstrSQL = "SELECT Nvl(MAX(LPAD(����," & int���� & ",'0')),'00') As Code FROM  " & strTable
    zlDatabase.OpenRecordset rsTemp, gstrSQL, strTittle
    strCode = rsTemp!Code
    
    int���� = Len(strCode)
    strCode = strCode + 1
    
    If int���� >= Len(strCode) Then
    strCode = String(int���� - Len(strCode), "0") & strCode
    End If
    strSpecify = zlCommFun.SpellCode(str����)
    
    
    gstrSQL = "ZL_" & strTable & "_INSERT('" & strCode & "','" & str���� & "','" & strSpecify & "')"
    zlDatabase.ExecuteProcedure gstrSQL, strTittle
    str���� = strCode
    AutoAddBaseItem = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub Logogram(ByVal staVal As StatusBar, ByVal bytType As Byte)
'���뷽ʽ
'staVal: StartusBar�ؼ�
'bytType: 0=ƴ��; 1=���;  ��ǰ����״̬
    Dim i As Integer
    For i = 1 To staVal.Panels.Count
        If staVal.Panels(i).Key = "PY" And staVal.Panels(i).Visible = True Then
            If bytType = 0 Then
                staVal.Panels(i).Bevel = sbrInset
                zlDatabase.SetPara "���뷽ʽ", 0
                gSystem_Para.int���뷽ʽ = 0
            ElseIf bytType = 1 Then
                staVal.Panels(i).Bevel = sbrRaised
            End If
        ElseIf staVal.Panels(i).Key = "WB" And staVal.Panels(i).Visible = True Then
            If bytType = 0 Then
                staVal.Panels(i).Bevel = sbrRaised
            ElseIf bytType = 1 Then
                staVal.Panels(i).Bevel = sbrInset
                zlDatabase.SetPara "���뷽ʽ", 1
                gSystem_Para.int���뷽ʽ = 1
            End If
        End If
    Next
End Sub


Public Function CheckQualifications(ByVal lngModule As Long, ByVal intType As Integer, ByVal StrInput As String) As Boolean
    'У�����ģ������̣���Ӧ����Ϣ������Ч��
    'intType��0�����ģ�1�������̣�2����Ӧ��
    'strInput���ַ���ʱΪ���ƣ�����ʱΪID
    Dim rsTmp As ADODB.Recordset
    Dim strMsgInfo As String
    Dim strMsgDate As String
    Dim dateCurrent As Date
    Dim strMsg As String
    
    Dim intCheckType As Integer
    Dim arrColumn
    Dim strCheck As String
    Dim strCheck_���� As String
    Dim strCheck_������ As String
    Dim strCheck_��Ӧ�� As String
    Dim n As Integer
    Dim strTmp As String
    
    On Error GoTo ErrHandle
    If StrInput = "" Then
        CheckQualifications = True
        Exit Function
    End If
        
    '����У����Ŀ�ͷ�ʽ�ı����ʽ��У�鷽ʽ|���1,��Ŀ1,�Ƿ�У��;���1,��Ŀ2,�Ƿ�У��;���2,��Ŀ1,�Ƿ�У��;���2,��Ŀ2....
    strCheck = zlDatabase.GetPara("����У��", glngSys, lngModule, "")
    
    '����Ĳ�����ʽ����ȷʱ�˳�
    If InStr(1, strCheck, "|") = 0 Then
        CheckQualifications = True
        Exit Function
    End If

    'ȡУ�鷽ʽ��0-����飻1�����ѣ�2����ֹ
    intCheckType = Val(Mid(strCheck, 1, InStr(1, strCheck, "|") - 1))
    
    '�����ʱ�˳�
    If intCheckType = 0 Then
        CheckQualifications = True
        Exit Function
    End If

    'ȡУ�����ݣ�
    strCheck = Mid(strCheck, InStr(1, strCheck, "|") + 1)

    If strCheck = "" Then
        CheckQualifications = True
        Exit Function
    End If

    '�ֱ�ȡ���ģ������̣���Ӧ����ҪУ�������
    strCheck = strCheck & ";"
    arrColumn = Split(strCheck, ";")
    For n = 0 To UBound(arrColumn)
        If arrColumn(n) <> "" Then
            If Split(arrColumn(n), ",")(0) = "����" And Split(arrColumn(n), ",")(2) = 1 Then
                strCheck_���� = IIf(strCheck_���� = "", "", strCheck_���� & ";") & Split(arrColumn(n), ",")(1)
            End If

            If Split(arrColumn(n), ",")(0) = "����������" And Split(arrColumn(n), ",")(2) = 1 Then
                strCheck_������ = IIf(strCheck_������ = "", "", strCheck_������ & ";") & Split(arrColumn(n), ",")(1)
            End If

            If Split(arrColumn(n), ",")(0) = "���Ĺ�Ӧ��" And Split(arrColumn(n), ",")(2) = 1 Then
                strCheck_��Ӧ�� = IIf(strCheck_��Ӧ�� = "", "", strCheck_��Ӧ�� & ";") & Split(arrColumn(n), ",")(1)
            End If
        End If
    Next
    
    '��У������ʱ�˳�
    If (intType = 0 And strCheck_���� = "") Or (intType = 1 And strCheck_������ = "") Or (intType = 2 And strCheck_��Ӧ�� = "") Then
        CheckQualifications = True
        Exit Function
    End If
    
    dateCurrent = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd"))
    
    '����
    If intType = 0 Then
        gstrSQL = "Select ('[' || B.���� || ']' || B.����) AS ������Ϣ, A.���֤��, A.���֤��Ч��,ע��֤��,ע��֤��Ч�� " & _
            " From �շ���ĿĿ¼ B,�������� A " & _
            " Where B.ID = A.����ID And A.����ID = [1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "У����������", Val(StrInput))
        
        If Not rsTmp.EOF Then
            If NVL(rsTmp!���֤��) = "" And InStr(strCheck_����, "���֤��") > 0 Then
                strTmp = rsTmp!������Ϣ & "��" & "�����֤��"
            End If
            
            If NVL(rsTmp!���֤��Ч��) <> "" Then
                If DateDiff("d", rsTmp!���֤��Ч��, dateCurrent) > 0 And InStr(strCheck_����, "���֤��Ч��") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!������Ϣ & "��", strTmp & ",") & "���֤�ѹ���"
                End If
            End If
        End If
        If NVL(rsTmp!ע��֤��) = "" And InStr(strCheck_����, "ע��֤��") > 0 Then
            strTmp = rsTmp!������Ϣ & "��" & "��ע��֤��"
        End If
        
        If NVL(rsTmp!ע��֤��Ч��) <> "" Then
            If DateDiff("d", rsTmp!ע��֤��Ч��, dateCurrent) > 0 And InStr(strCheck_����, "ע��֤��Ч��") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!������Ϣ & "��", strTmp & ",") & "ע��֤�ѹ���"
            End If
        End If
    End If
    
    '������
    If intType = 1 Then
        gstrSQL = "Select ('[' || A.���� || ']' || A.����) AS ������, A.������ҵ���֤, A.������ҵ���֤Ч��,a.��Ӫ���֤,a.��Ӫ���֤Ч��,a.��ҵ����ִ��,a.��ҵ����ִ��Ч�� " & _
                        " From ���������� A " & _
                        " Where A.���� = [1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "У����������", StrInput)
        
        If Not rsTmp.EOF Then
            If NVL(rsTmp!������ҵ���֤) = "" And InStr(strCheck_������ & ";", "������ҵ���֤" & ";") > 0 Then
                strTmp = rsTmp!������ & "��" & "��������ҵ���֤"
            End If
            
            If NVL(rsTmp!������ҵ���֤Ч��) <> "" Then
                If DateDiff("d", rsTmp!������ҵ���֤Ч��, dateCurrent) > 0 And InStr(strCheck_������ & ";", "������ҵ���֤Ч��" & ";") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!������ & "��", strTmp & ",") & "������ҵ���֤�ѹ���"
                End If
            End If
        End If
        If Not rsTmp.EOF Then
            If NVL(rsTmp!��Ӫ���֤) = "" And InStr(strCheck_������ & ";", "��Ӫ���֤" & ";") > 0 Then
                strTmp = rsTmp!������ & "��" & "�޾�Ӫ���֤"
            End If
            
            If NVL(rsTmp!��Ӫ���֤Ч��) <> "" Then
                If DateDiff("d", rsTmp!������ҵ���֤Ч��, dateCurrent) > 0 And InStr(strCheck_������ & ";", "��Ӫ���֤Ч��" & ";") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!������ & "��", strTmp & ",") & "��Ӫ���֤�ѹ���"
                End If
            End If
        End If
        If Not rsTmp.EOF Then
            If NVL(rsTmp!��ҵ����ִ��) = "" And InStr(strCheck_������ & ";", "��ҵ����ִ��" & ";") > 0 Then
                strTmp = rsTmp!������ & "��" & "����ҵ����ִ��"
            End If
            
            If NVL(rsTmp!��ҵ����ִ��Ч��) <> "" Then
                If DateDiff("d", rsTmp!������ҵ���֤Ч��, dateCurrent) > 0 And InStr(strCheck_������ & ";", "��ҵ����ִ��Ч��" & ";") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!������ & "��", strTmp & ",") & "��ҵ����ִ���ѹ���"
                End If
            End If
        End If
    End If
    
    '��Ӧ��
    If intType = 2 Then
        gstrSQL = "Select ('[' || ���� || ']' || ����) AS ��Ӧ��, ˰��ǼǺ�, ���֤��, ִ�պ�, ��Ȩ��, ������֤��, ������֤����, ҩ��ֱ�����, ҩ��ֱ�������, ���֤Ч��, ִ��Ч��, ��Ȩ�� " & _
            " From ��Ӧ�� " & _
            " Where (����ʱ�� Is Null Or ����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And ID = [1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "��Ӧ����Ϣ", Val(StrInput))
        
        strTmp = ""
        
        If Not rsTmp.EOF Then
            If NVL(rsTmp!˰��ǼǺ�) = "" And InStr(strCheck_��Ӧ��, "˰��ǼǺ�") > 0 Then
                strTmp = rsTmp!��Ӧ�� & "��" & "��˰��ǼǺ�"
            End If
            
            If NVL(rsTmp!���֤��) = "" And InStr(strCheck_��Ӧ��, "���֤��") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "�����֤��"
            End If
            
            If NVL(rsTmp!ִ�պ�) = "" And InStr(strCheck_��Ӧ��, "ִ�պ�") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "��ִ�պ�"
            End If
            
            If NVL(rsTmp!��Ȩ��) = "" And InStr(strCheck_��Ӧ��, "��Ȩ��") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "����Ȩ��"
            End If
            
            If NVL(rsTmp!������֤��) = "" And InStr(strCheck_��Ӧ��, "������֤��") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "��������֤��"
            End If
            
            If NVL(rsTmp!������֤����) <> "" Then
                If DateDiff("d", rsTmp!������֤����, dateCurrent) > 0 And InStr(strCheck_��Ӧ��, "������֤����") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "������֤���ѹ���"
                End If
            End If
            
            If NVL(rsTmp!ҩ��ֱ�����) = "" And InStr(strCheck_��Ӧ��, "ҩ��ֱ�����") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "��ҩ��ֱ�����"
            End If
            
            If NVL(rsTmp!ҩ��ֱ�������) <> "" Then
                If DateDiff("d", rsTmp!ҩ��ֱ�������, dateCurrent) > 0 And InStr(strCheck_��Ӧ��, "ҩ��ֱ�������") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "ҩ��ֱ������ѹ���"
                End If
            End If
            
            If NVL(rsTmp!���֤Ч��) <> "" Then
                If DateDiff("d", rsTmp!���֤Ч��, dateCurrent) > 0 And InStr(strCheck_��Ӧ��, "���֤Ч��") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "���֤�ѹ���"
                End If
            End If
            
            If NVL(rsTmp!ִ��Ч��) <> "" Then
                If DateDiff("d", rsTmp!ִ��Ч��, dateCurrent) > 0 And InStr(strCheck_��Ӧ��, "ִ��Ч��") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "ִ���ѹ���"
                End If
            End If
            
            If NVL(rsTmp!��Ȩ��) <> "" Then
                If DateDiff("d", rsTmp!ִ��Ч��, dateCurrent) > 0 And InStr(strCheck_��Ӧ��, "��Ȩ��") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!��Ӧ�� & "��", strTmp & ",") & "��Ȩ�ѹ���"
                End If
            End If
        End If
    End If
    
    '��ʾ���ֹ
    If strTmp <> "" Then
        If intCheckType = 1 Then
            If MsgBox("δͨ������У�飬�Ƿ������" & vbCrLf & strTmp, vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
                CheckQualifications = True
                Exit Function
            Else
                Exit Function
            End If
        ElseIf intCheckType = 2 Then
            MsgBox "δͨ������У�飬������⣡" & vbCrLf & strTmp, vbOKOnly, gstrSysName
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
    'vsf�����ã��������п��ж��뷽ʽ���̶��ж��뷽ʽ��Ĭ��Ϊ���ж��룩
    
    With objGrid
        .TextMatrix(0, intCol) = strColName
        .ColWidth(intCol) = lngColWidth
        .ColAlignment(intCol) = intColAlignment
'        .ColData(intCol) = lngColWidth
        
        .ColKey(intCol) = strColKey
        .FixedAlignment(intCol) = intFixedColAlignment
    End With
End Sub


'ȡϵͳ����ֵ
Public Sub GetSysParms()
    Dim rs As New ADODB.Recordset
    Dim n As Integer
    Dim strMsg As String
    
    On Error GoTo errH
    
    'ȡ�������������
    gstrSQL = "Select ���۽��, �ɱ���, ���ۼ�, ʵ������ From ҩƷ�շ���¼ Where Rownum < 1"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "ȡҩƷ����")
    gtype_UserDrugDigits.Digit_��� = rs.Fields(0).NumericScale
    gtype_UserDrugDigits.Digit_�ɱ��� = rs.Fields(1).NumericScale
    gtype_UserDrugDigits.Digit_���ۼ� = rs.Fields(2).NumericScale
    gtype_UserDrugDigits.Digit_���� = rs.Fields(3).NumericScale
    
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Public Function StuffWork_GetCheckStockRule(ByVal lng�ⷿID As Long) As Integer
    'ȡ���������
    Dim rsData As ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = " Select Nvl(��鷽ʽ,0) ����� From ���ϳ����� Where �ⷿID=[1]"
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ���������", lng�ⷿID)

    If Not rsData.EOF Then
        StuffWork_GetCheckStockRule = rsData!�����
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function Get��������(ByVal lng�ⷿID As Long, ByVal lng����ID As Long) As Integer
    '����ָ���ⷿ��ָ�����ϵķ�������
    '���أ�0-��������1-����
    Dim rsCheck As New ADODB.Recordset
    Dim int���� As Integer
    Dim bln���ϲ��� As Boolean
    Dim strSQL As String
        
    On Error GoTo ErrHandle
    
    '�ж��Ƿ��Ƿ��ϲ���
    strSQL = "select ����ID from ��������˵�� where (�������� =  '���ϲ���' or �������� =  '�Ƽ���') And ����id=[1]"
    Set rsCheck = zlDatabase.OpenSQLRecord(strSQL, "Get��������", lng�ⷿID)

    bln���ϲ��� = (Not rsCheck.EOF)
        
    '�ж϶�Ӧ��ҩƷĿ¼�еķ�������
    strSQL = " Select Nvl(�ⷿ����,0) As �ⷿ����,nvl(���÷���,0) As ���÷��� " & _
              " From �������� Where ����ID=[1]"
    Set rsCheck = zlDatabase.OpenSQLRecord(strSQL, "Get��������", lng����ID)
              
    If bln���ϲ��� Then
        int���� = rsCheck!���÷���
    Else
        int���� = rsCheck!�ⷿ����
    End If
    
    Get�������� = int����
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function CheckNoStock(ByVal lng�ⷿID As Long, ByVal lng����ID As Long, Optional ByVal lng���� As Long = -1) As Boolean
    '����Ƿ��޿��
    '���أ�true-�޿��;false-�п��
    Dim rsData As ADODB.Recordset
    
    On Error GoTo ErrHandle
    
    gstrSQL = "Select 1 From ҩƷ��� " & _
        " Where ���� = 1 And �ⷿid = [1] And ҩƷid = [2] And (Nvl(ʵ������, 0) <> 0 Or Nvl(ʵ�ʽ��, 0) <> 0 Or Nvl(ʵ�ʲ��, 0) <> 0) "
    
    If lng���� <> -1 Then
        gstrSQL = gstrSQL & " And Nvl(����,0) = [3] "
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "CheckNoStock", lng�ⷿID, lng����ID, lng����)
    
    CheckNoStock = rsData.EOF
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CheckNumStock(ByVal objVSF As Object, ByVal lng�ⷿID As Long, ByVal lntCol����id As Integer, _
    ByVal intCol���� As Integer, ByVal intCol���� As Integer, ByVal intCol����ϵ�� As Integer, _
    ByVal intMethod As Integer, Optional int���ҵ�� As Integer, Optional ByVal int���� As Integer, _
    Optional ByVal intType As Integer = 0) As String
    '���ܣ���˳����൥��ʱ�����ŵ��ݼ�����ʵ��(����)�����Ƿ��㹻
    '������objVSF-��Ҫ���ı��;lng�ⷿid��intcol����-���������У�intCol����-���������У�intCol����ϵ��-����ϵ��������
    '������intMethod��1-������ˣ�2-������3-�˿����
    '������int���ҵ��0-��⣻1-����
    '������intType��0-ʵ��������1-��������
    '����ֵ�����о���Ĳ������ƣ�Ϊ��-���ͨ�����������㣻��Ϊ��-���δͨ��������������
    Dim objCol As Collection         '��ʹ�õ���������
    Dim dblNum As Double
    Dim varNum As Variant
    Dim varTemp As Variant
    Dim strTemp As String
    Dim lng����ID As Long
    Dim lng���� As Long
    Dim rsData As ADODB.Recordset
    Dim strKey As String
    Dim vardrug As Variant
    Dim lngRow As Long
    Dim strArray As String
    Dim dbl����ϵ�� As Double
    Dim intSum As Integer '����
    
    On Error GoTo ErrHandle
    
    '����ϱ�������������������Ҫ�ǿ��ǲ����������
    Set objCol = New Collection
    With objVSF
        If .Rows < 2 Then Exit Function
        For lngRow = 1 To .Rows - 1
            dblNum = 0
            If .TextMatrix(lngRow, lntCol����id) <> "" Then
                For Each vardrug In objCol
                    If vardrug(0) = .TextMatrix(lngRow, lntCol����id) & "," & Val(.TextMatrix(lngRow, intCol����)) & "," & Val(.TextMatrix(lngRow, intCol����ϵ��)) Then
                        dblNum = vardrug(1)
                        objCol.Remove vardrug(0)
                        Exit For
                    End If
                Next
                strKey = .TextMatrix(lngRow, lntCol����id) & "," & Val(.TextMatrix(lngRow, intCol����)) & "," & Val(.TextMatrix(lngRow, intCol����ϵ��))
                '����С��λ�����������������ʱ�����������ݱȽ�
                strArray = dblNum + (Val(.TextMatrix(lngRow, intCol����)))
                objCol.Add Array(strKey, strArray), strKey
            End If
        Next
    End With
    
    For Each varNum In objCol
        strTemp = varNum(0)  '��ʽ�ǲ���id,����,����ϵ��
        dblNum = varNum(1)
        varTemp = Split(strTemp, ",")
        If int���ҵ�� = 0 Then '���
            If intMethod = 1 Then '�������
                If dblNum < 0 Then
                    '������⣬��Ҫ����棬������Ҫ�жϿ���Ƿ����
                    dblNum = Abs(dblNum)
                Else
                    '������⣬������棬���Բ����
                    dblNum = 0
                End If
            ElseIf intMethod = 2 Then
                '����
                If dblNum < 0 Then
                    dblNum = 0
                Else
                    dblNum = dblNum
                End If
            ElseIf intMethod = 3 Then
                '�˿���ˣ��˿����¼������
                dblNum = dblNum
            End If
        Else    '����
            If intMethod = 1 Then '�������
                If dblNum < 0 Then
                    '�������⣬������棬���Բ����
                    dblNum = 0
                Else
                    '������⣬��Ҫ����棬������Ҫ�жϿ���Ƿ����
                    dblNum = dblNum
                End If
            ElseIf intMethod = 2 Then
                '����
                If dblNum < 0 Then
                    dblNum = Abs(dblNum)
                Else
                    dblNum = 0
                End If
            End If
        End If
        
        'ֻ�����������ж�
        If dblNum > 0 Then
            lng����ID = varTemp(0)
            lng���� = varTemp(1)
            dbl����ϵ�� = varTemp(2)

            If Get��������(lng�ⷿID, lng����ID) = 0 Then
                lng���� = 0
            End If
            
            gstrSQL = "Select a.��������,a.ʵ������, '[' || b.���� || ']' || b.���� ����" & vbNewLine & _
                        "From ҩƷ��� A, �շ���ĿĿ¼ B" & vbNewLine & _
                        "Where a.ҩƷid = b.Id And a.ҩƷid = [2] And a.�ⷿid = [3] And Nvl(a.����, 0) = [4] And b.��� = '4' And a.���� = 1"
            Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "�����", dblNum, lng����ID, lng�ⷿID, lng����)
            If rsData.RecordCount = 0 Then '�޿���¼
                gstrSQL = "select '[' || ���� || ']' || ���� ���� from �շ���ĿĿ¼ where id=[1]"
                Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, "�����", lng����ID)
                
                intSum = intSum + 1
                If intSum <= 3 Then CheckNumStock = IIf(CheckNumStock = "", "", CheckNumStock & "��" & vbCrLf) & rsData!����
            Else '�п���¼
                If intType = 0 Then '�Ƚ�ʵ������
                    If zlStr.FormatEx(rsData!ʵ������ / dbl����ϵ��, int����, , False) >= dblNum Then
                    Else
                        intSum = intSum + 1
                        If intSum <= 3 Then CheckNumStock = IIf(CheckNumStock = "", "", CheckNumStock & "��" & vbCrLf) & rsData!����
                    End If
                Else '�ȽϿ�������
                    If zlStr.FormatEx(rsData!�������� / dbl����ϵ��, int����, , False) >= dblNum Then
                    Else
                        intSum = intSum + 1
                        If intSum <= 3 Then CheckNumStock = IIf(CheckNumStock = "", "", CheckNumStock & "��" & vbCrLf) & rsData!����
                    End If
                End If
            End If
'            Next
        End If
    Next
    CheckNumStock = CheckNumStock & IIf(intSum > 3, "��" & intSum & "��", "")
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

