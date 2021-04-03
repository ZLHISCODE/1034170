VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrivilege"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################
'��������
Private mstrRole As String
Private mstrOwners As String
Private mstrSQL As String
Private mrsModule As ADODB.Recordset                'ģ�����������¼
Private mrsTable  As ADODB.Recordset                '��������������¼
Private mrsFunction  As ADODB.Recordset             '�������������¼
Private mrsPrivs As ADODB.Recordset
Private mrsGranded As ADODB.Recordset            '������ǰ��Ȩ��ģ��
Private mrsFuncPars As ADODB.Recordset           '�������������ʵı�
Private mcnOracle As ADODB.Connection
Private mobjCollectionMoudal As Collection       '
Private mobjCollectionGranded As Collection      '
Private mobjColGrantObject As Collection         '���е�Ȩ��
Private mobjColRevokeObject As Collection        '�����е�Ȩ��
Private mstrFailInfo As String
Private mblnHaveBase As Boolean                  '�Ƿ�����ֵ��������ģ��Ȩ��

'�¼�����
Public Event BeforeProgress(ByVal Title As String, ByVal Max As Long)
Public Event Progressing(ByVal Progress As Long)
Public Event AfterProgress()

'######################################################################################################################

Public Function InitOracle(ByVal cnOracle As ADODB.Connection)
    Set mcnOracle = cnOracle
    InitOracle = True
End Function

Public Function InitPrivilege(strOwners As String, ByVal strRole As String, ByVal rsModule As ADODB.Recordset, ByVal rsTable As ADODB.Recordset, ByVal rsFunction As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '���ܣ��ӿں�������ʼ��Ȩ�޴�������
    '������
    '���أ��ɹ�����true�����򷵻�false
    '******************************************************************************************************************
    Dim strKey As String
    
    On Error GoTo errHand
    
    mstrFailInfo = ""
    mstrRole = strRole
    mstrOwners = strOwners
    Set mrsModule = rsModule
    Set mrsTable = rsTable
    Set mrsFunction = rsFunction
    
    '------------------------------------------------------------------------------------------------------------------
    Set mrsGranded = OpenCursor(mcnOracle, "ZLTOOLS.B_Popedom.Get_Role_Grant", mstrRole)
    Set mrsPrivs = mrsGranded.NextRecordset
    Set mrsFuncPars = mrsGranded.NextRecordset
    '�ж��Ƿ�����ֵ��������Ȩ��
    mrsModule.Filter = "���=11 And ϵͳ=0 And ��Ȩ��=1"
    mblnHaveBase = Not mrsModule.EOF
    
    mrsModule.Filter = ""
    If mrsModule.BOF = False Then
        mrsModule.MoveFirst
        Do While Not mrsModule.EOF
            strKey = "K" & mrsModule("ϵͳ").value & "'" & mrsModule("���").value & "'" & mrsModule("����").value & "'" & mrsModule("��Ȩ��").value
            mobjCollectionMoudal.Add "1", strKey
            mrsModule.MoveNext
        Loop
    End If
        
    If mrsGranded.BOF = False Then
        Do While Not mrsGranded.EOF
            strKey = "K" & mrsGranded("ϵͳ").value & "'" & mrsGranded("���").value & "'" & mrsGranded("����").value
            mobjCollectionGranded.Add "1", strKey
            mrsGranded.MoveNext
        Loop
    End If
    
    InitPrivilege = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
End Function

Public Function RevokePrivilege() As Boolean
    '******************************************************************************************************************
    '���ܣ��ӿں�����������еĶ���Ȩ��
    '������
    '���أ�
    '******************************************************************************************************************
    Dim blnReturn As Boolean
    
    On Error GoTo errHand
    
    '1  ȡ����ǰ����Ķ����ݱ��ķ���Ȩ��
    '------------------------------------------------------------------------------------------------------------------
    blnReturn = GetDictRevokeObjectData(mrsTable)

    '2 ȡ����ǰ����ĶԺ����ķ���Ȩ��
    '------------------------------------------------------------------------------------------------------------------
    If blnReturn Then blnReturn = GetFunctionRevokeObjectData(mrsFunction, mrsFuncPars)
    
    '3  ȡ����ǰ����Ķ�ģ��ķ���Ȩ��
    '------------------------------------------------------------------------------------------------------------------
    If blnReturn Then blnReturn = GetMoudalRevokeObjectData(mrsGranded, mrsPrivs)
           
    '4  ��ʼִ�����Ȩ��
    '------------------------------------------------------------------------------------------------------------------
    If blnReturn Then blnReturn = DeleteGrant(mrsGranded, mrsPrivs)
    If blnReturn Then blnReturn = ExecuteRevokeObject(mobjColRevokeObject)
    
    RevokePrivilege = blnReturn

    Exit Function
    
errHand:
    MsgBox err.Description
End Function

Public Property Get FailInfo() As String
    FailInfo = mstrFailInfo
End Property

Public Function ReGrantPrivilege(ByVal strRole As String, Optional ByVal strOwners As String) As Boolean
    '******************************************************************************************************************
    '���ܣ��ӿڣ�������Ȩ
    '������
    '���أ�
    '******************************************************************************************************************
    Dim rsTemp As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim lngCount As Long
    Dim objCollection As New Collection
    Dim strKey As String
    
    
    mstrFailInfo = ""
    
    mstrSQL = "select Distinct P.����,P.������,P.Ȩ�� " & _
              "  from  zlProgPrivs p,zlRolegrant R " & _
              "  Where nvl(p.ϵͳ, 0) = nvl(r.ϵͳ, 0) And p.��� = r.��� And p.���� = r.���� And R.��ɫ='" & strRole & "'"

    rsTemp.Open mstrSQL, mcnOracle, adOpenStatic, adLockReadOnly
    If rsTemp.BOF = False Then
        mstrSQL = "select Upper(Owner) As ������,Upper(table_name) As ����,Upper(Privilege) As Ȩ�� from role_tab_privs where role='" & strRole & "' And table_name Not Like 'BIN$%'"
        rs.Open mstrSQL, mcnOracle
        If rs.BOF = False Then
            Do While Not rs.EOF
                objCollection.Add "1", "K" & rs("������").value & "'" & rs("����").value & "'" & rs("Ȩ��").value
                rs.MoveNext
            Loop
        End If
                
        Do While Not rsTemp.EOF
            strKey = "K" & UCase(rsTemp("������").value) & "'" & UCase(rsTemp("����").value) & "'" & UCase(rsTemp("Ȩ��").value)
            
            If IsCollectionElement(objCollection, strKey) = False Then
                mstrSQL = "Grant " & rsTemp("Ȩ��").value & " on " & rsTemp("������").value & "." & rsTemp("����").value & " to " & strRole
                Call ExecuteSQL(mstrSQL)
            End If
            
            rsTemp.MoveNext
        Loop

        '����һЩ�������Ȩ��
        '------------------------------------------------------------------------------------------------------------------
        Call GrantSpecialPrivilege(strRole, strOwners)
    End If
    
    ReGrantPrivilege = True
    
End Function

Public Function GrantPrivilege() As Boolean
    '******************************************************************************************************************
    '���ܣ��ӿں������������Ȩ��
    '������
    '���أ�
    '******************************************************************************************************************
    Dim blnReturn As Boolean
    On Error GoTo errHand
    
    '1 �����ݱ�����������Ȩ
    '------------------------------------------------------------------------------------------------------------------
    blnReturn = GetDictGrantObjectData(mrsTable)
    
    '2 �Ժ�������������Ȩ
    '------------------------------------------------------------------------------------------------------------------
    If blnReturn Then blnReturn = GetFunctionGrantObjectData(mrsFunction, mrsFuncPars)
        
    '3 ��ģ�����������Ȩ
    '------------------------------------------------------------------------------------------------------------------
    If blnReturn Then blnReturn = GetMoudalGrantObjectData(mrsGranded, mrsPrivs)

    '4 ִ����Ȩ
    '------------------------------------------------------------------------------------------------------------------
    If blnReturn Then blnReturn = SaveGrant(mrsGranded, mrsPrivs)
    If blnReturn Then blnReturn = ExecuteGrantObject(mobjColGrantObject)
    If blnReturn Then blnReturn = GrantSpecialPrivilege(mstrRole, mstrOwners)
    
    GrantPrivilege = blnReturn
    
    Exit Function
    
errHand:
    MsgBox err.Description
End Function

Private Function GrantSpecialPrivilege(ByVal strRole As String, Optional ByVal strOwners As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim lngLoop As Long
    
    On Error GoTo errHand
    
    GrantSpecialPrivilege = GrantSpecialToRole(mcnOracle, strRole, mblnHaveBase, strOwners)

    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
    If 0 = 1 Then
        Resume
    End If
End Function

Private Function GetDictRevokeObjectData(ByVal rsTable As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '���ܣ���ȡ��������ض���Ȩ������
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Long
    Dim intCount As Long
    Dim intPercent As Integer
    Dim intSvrPercent As Integer
    
    On Error GoTo errHand
        
    rsTable.Filter = "��Ȩ�� = 0"
    intCount = rsTable.RecordCount
    If intCount > 0 Then
    
        RaiseEvent BeforeProgress("�ռ��ֵ�ȡ��Ȩ��", 100)
        
        intLoop = 0
        Do Until rsTable.EOF
            
            If rsTable("�����Ȩ") = 1 Then
                
                Call AddPrivilegeData(rsTable("������") & "'" & rsTable("����") & "'select", 2)
                Call AddPrivilegeData(rsTable("������") & "'" & rsTable("����") & "'insert", 2)
                Call AddPrivilegeData(rsTable("������") & "'" & rsTable("����") & "'update", 2)
                Call AddPrivilegeData(rsTable("������") & "'" & rsTable("����") & "'delete", 2)
                        
            End If
            
            intLoop = intLoop + 1
            intPercent = 100 * intLoop / intCount
            If intSvrPercent <> intPercent Then
                intSvrPercent = intPercent
                RaiseEvent Progressing(intSvrPercent)
            End If
                    
            rsTable.MoveNext
        Loop
        
        RaiseEvent AfterProgress
    End If
    
    GetDictRevokeObjectData = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
'    Resume
End Function

Private Function GetDictGrantObjectData(ByVal rsTable As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '���ܣ���ȡ��������ض���Ȩ������
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Long
    Dim intCount As Long
    Dim intPercent As Integer
    Dim intSvrPercent As Integer
    
    On Error GoTo errHand

    rsTable.Filter = "��Ȩ�� = 1"
    intCount = rsTable.RecordCount
    If intCount > 0 Then
        
        RaiseEvent BeforeProgress("�����ռ��ֵ�Ȩ��", 100)
        
        intLoop = 0
        Do Until rsTable.EOF
            Call AddPrivilegeData(rsTable("������") & "'" & rsTable("����") & "'select", 1)
            Call AddPrivilegeData(rsTable("������") & "'" & rsTable("����") & "'insert", 1)
            Call AddPrivilegeData(rsTable("������") & "'" & rsTable("����") & "'update", 1)
            Call AddPrivilegeData(rsTable("������") & "'" & rsTable("����") & "'delete", 1)

            intLoop = intLoop + 1
            intPercent = 100 * intLoop / intCount
            If intSvrPercent <> intPercent Then
                intSvrPercent = intPercent
                RaiseEvent Progressing(intSvrPercent)
            End If
                    
            rsTable.MoveNext
        Loop
        
        RaiseEvent AfterProgress
    End If
    
    GetDictGrantObjectData = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
'    Resume
End Function

Private Function GetFunctionRevokeObjectData(ByVal rsFunction As ADODB.Recordset, ByVal rsFuncPars As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '���ܣ���ȡ��������ض���Ȩ������
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strFail As String
    Dim blnAllow As Boolean
    Dim lngMoudal As Long
    Dim strFunc As String
    Dim intLoop As Long
    Dim intCount As Long
    Dim var���� As Variant
    Dim strKey As String
    Dim i As Long
    Dim varTemp As Variant
    Dim intPercent As Integer
    Dim intSvrPercent As Integer
    
    On Error GoTo errHand
    
    rsFunction.Filter = "��Ȩ�� = 0"
    
    intCount = rsFunction.RecordCount
    If intCount > 0 Then
        RaiseEvent BeforeProgress("�ռ�����ȡ��Ȩ��", 100)
        intLoop = 0
        
        Do Until rsFunction.EOF
            If rsFunction("�����Ȩ") = 1 Then
                'ֻ����ǰ���ڹ������ݱ��Ľ�ɫ������Ҫִ��ȡ������
                Call AddPrivilegeData(rsFunction("������") & "'" & UCase(rsFunction("������")) & "'execute", 2)
                
                '��ȡ������ʹ�ö����б�
                rsFuncPars.Filter = "ϵͳ = " & mrsFunction("ϵͳ") & " and ������ = '" & mrsFunction("������") & "'"
                Do Until rsFuncPars.EOF
                    var���� = Split(Replace(rsFuncPars("����"), "|", ","), ",")
                    For i = LBound(var����) To UBound(var����)
                        If var����(i) <> "" Then
                            varTemp = Split(var����(i), ".")
                            Call AddPrivilegeData(CStr(varTemp(0)) & "'" & CStr(varTemp(1)) & "'select", 2)
                        End If
                    Next
                    rsFuncPars.MoveNext
                Loop
            End If
            
            intLoop = intLoop + 1
            intPercent = 100 * intLoop / intCount
            If intSvrPercent <> intPercent Then
                intSvrPercent = intPercent
                RaiseEvent Progressing(intSvrPercent)
            End If
                        
            rsFunction.MoveNext
        Loop
        RaiseEvent AfterProgress
    End If
        
    GetFunctionRevokeObjectData = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
    If 0 = 1 Then
        Resume
    End If
End Function

Private Function GetFunctionGrantObjectData(ByVal rsFunction As ADODB.Recordset, ByVal rsFuncPars As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '���ܣ���ȡ��������ض���Ȩ������
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strFail As String
    Dim blnAllow As Boolean
    Dim lngMoudal As Long
    Dim strFunc As String
    Dim intLoop As Long
    Dim intCount As Long
    Dim var���� As Variant
    Dim strKey As String
    Dim i As Long
    Dim varTemp As Variant
    Dim intPercent As Integer
    Dim intSvrPercent As Integer
    
    On Error GoTo errHand
    
    rsFunction.Filter = "��Ȩ�� = 1"
    
    intCount = rsFunction.RecordCount
    If intCount > 0 Then
        RaiseEvent BeforeProgress("�����ռ�����Ȩ��", 100)
        intLoop = 0
        Do Until mrsFunction.EOF
'            'ֻ����ǰ���ڹ������ݱ��Ľ�ɫ������Ҫִ��ȡ������
            Call AddPrivilegeData(rsFunction("������") & "'" & UCase(rsFunction("������")) & "'execute", 1)
            
            '��ȡ������ʹ�ö����б�
            rsFuncPars.Filter = "ϵͳ = " & rsFunction("ϵͳ") & " and ������ = '" & rsFunction("������") & "'"
            Do Until rsFuncPars.EOF
                var���� = Split(Replace(rsFuncPars("����"), "|", ","), ",")
                For i = LBound(var����) To UBound(var����)
                    If var����(i) <> "" Then
                        varTemp = Split(var����(i), ".")
                        Call AddPrivilegeData(CStr(varTemp(0)) & "'" & CStr(varTemp(1)) & "'select", 1)
                    End If
                Next
                rsFuncPars.MoveNext
            Loop

            intLoop = intLoop + 1
            intPercent = 100 * intLoop / intCount
            If intSvrPercent <> intPercent Then
                intSvrPercent = intPercent
                RaiseEvent Progressing(intSvrPercent)
            End If
                        
            rsFunction.MoveNext
        Loop
        RaiseEvent AfterProgress
    End If
    
    GetFunctionGrantObjectData = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
'    Resume
End Function

Private Function GetMoudalRevokeObjectData(ByVal rsGranded As ADODB.Recordset, ByVal rsPrivs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '���ܣ���ȡģ�����ض���Ȩ������
    '������
    '���أ�
    '******************************************************************************************************************
    Dim blnAllow As Boolean
    Dim lngMoudal As Long
    Dim strFunc As String
    Dim intLoop As Long
    Dim intCount As Long
    Dim strKey As String
    Dim lngSys As Long
    Dim intPercent As Integer
    Dim intSvrPercent As Integer
    
    On Error GoTo errHand
    
    rsPrivs.Sort = "ϵͳ,���,����"
    intCount = rsPrivs.RecordCount
    If intCount > 0 Then
        RaiseEvent BeforeProgress("�ռ�ģ��ȡ��Ȩ��", 100)
    
        Do While Not rsPrivs.EOF
            If lngSys <> rsPrivs("ϵͳ").value Or lngMoudal <> rsPrivs("���").value Or strFunc <> rsPrivs("����").value Then
                lngSys = rsPrivs("ϵͳ").value
                lngMoudal = rsPrivs("���").value
                strFunc = IIf(IsNull(rsPrivs("����").value), "", rsPrivs("����").value)
                
                blnAllow = False
                
                strKey = "K" & lngSys & "'" & lngMoudal & "'" & strFunc & "'0"
                If IsCollectionElement(mobjCollectionMoudal, strKey) = True Then
                    strKey = "K" & lngSys & "'" & lngMoudal & "'" & strFunc
                    blnAllow = IsCollectionElement(mobjCollectionGranded, strKey)
                End If
                
            End If
            
            If blnAllow = True Then
                Call AddPrivilegeData(rsPrivs("������").value & "'" & rsPrivs("����").value & "'" & rsPrivs("Ȩ��").value, 2)
            End If
            
            intLoop = intLoop + 1
            intPercent = 100 * intLoop / intCount
            If intSvrPercent <> intPercent Then
                intSvrPercent = intPercent
                RaiseEvent Progressing(intSvrPercent)
            End If
                
            rsPrivs.MoveNext
        Loop
        
        RaiseEvent AfterProgress
        
    End If
    GetMoudalRevokeObjectData = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
End Function

Private Function GetMoudalGrantObjectData(ByVal rsGranded As ADODB.Recordset, ByVal rsPrivs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '���ܣ���ȡģ�����ض���Ȩ������
    '������
    '���أ�
    '******************************************************************************************************************
    Dim blnAllow As Boolean
    Dim lngMoudal As Long
    Dim strFunc As String
    Dim intLoop As Long
    Dim intCount As Long
    Dim strKey As String
    Dim lngSys As Long
    Dim intPercent As Integer
    Dim intSvrPercent As Integer
    
    On Error GoTo errHand
    
    rsPrivs.Sort = "ϵͳ,���,����"
    intCount = rsPrivs.RecordCount
    If intCount > 0 Then
        RaiseEvent BeforeProgress("�����ռ�ģ��Ȩ��", 100)
                
        Do While Not rsPrivs.EOF
            If lngSys <> rsPrivs("ϵͳ").value Or lngMoudal <> rsPrivs("���").value Or strFunc <> rsPrivs("����").value Then
                
                lngSys = rsPrivs("ϵͳ").value
                lngMoudal = rsPrivs("���").value
                strFunc = IIf(IsNull(rsPrivs("����").value), "", rsPrivs("����").value)

                strKey = "K" & lngSys & "'" & lngMoudal & "'" & strFunc & "'1"
                blnAllow = IsCollectionElement(mobjCollectionMoudal, strKey)

            End If
            
            If blnAllow = True Then
                Call AddPrivilegeData(rsPrivs("������").value & "'" & rsPrivs("����").value & "'" & rsPrivs("Ȩ��").value, 1)
            End If
            
            intLoop = intLoop + 1
            intPercent = 100 * intLoop / intCount
            If intSvrPercent <> intPercent Then
                intSvrPercent = intPercent
                RaiseEvent Progressing(intSvrPercent)
            End If
            
            rsPrivs.MoveNext
        Loop
        
        RaiseEvent AfterProgress
    
    End If
    
    GetMoudalGrantObjectData = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
End Function

Private Function ExecuteRevokeObject(ByVal objColGrangObject As Collection) As Boolean
    '******************************************************************************************************************
    '���ܣ�ִ����Ȩ
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Long
    Dim intCount As Long
    Dim strTemp As String
    Dim varTemp As Variant
    Dim rs As New ADODB.Recordset
    Dim strFail As String
    Dim objCollection As New Collection
    Dim strKey As String
    Dim intPercent As Integer
    Dim intSvrPercent As Integer
    
    On Error GoTo errHand
    
    intCount = objColGrangObject.Count
    If intCount > 0 Then
        
        RaiseEvent BeforeProgress("����ȡ����ɫȨ��", 100)
        
        mstrSQL = "select Upper(Owner) As ������,Upper(table_name) As ����,Upper(Privilege) As Ȩ�� from role_tab_privs where role='" & mstrRole & "' And table_name Not Like 'BIN$%'"
        rs.Open mstrSQL, mcnOracle
        If rs.BOF = False Then
            Do While Not rs.EOF
                objCollection.Add "1", "K" & rs("������").value & "'" & rs("����").value & "'" & rs("Ȩ��").value
                rs.MoveNext
            Loop
        End If
        
        For intLoop = 1 To intCount
            strTemp = objColGrangObject.Item(intLoop)
            varTemp = Split(strTemp, "'")
            
            strKey = "K" & UCase(CStr(varTemp(0))) & "'" & UCase(CStr(varTemp(1))) & "'" & UCase(CStr(varTemp(2)))

            If IsCollectionElement(objCollection, strKey) = True Then
                Call ExecuteSQL("revoke " & CStr(varTemp(2)) & " on " & CStr(varTemp(0)) & "." & CStr(varTemp(1)) & " from " & mstrRole)
                
                If err <> 0 Then
                    strFail = strFail & vbCrLf & CStr(varTemp(2)) & " on " & CStr(varTemp(0)) & "." & CStr(varTemp(1))
                    err.Clear
                End If
            End If
            
            intPercent = 100 * intLoop / intCount
            If intSvrPercent <> intPercent Then
                intSvrPercent = intPercent
                RaiseEvent Progressing(intSvrPercent)
            End If
            
        Next
        
        RaiseEvent AfterProgress
   End If
    
    ExecuteRevokeObject = True
    
    Exit Function
   '------------------------------------------------------------------------------------------------------------------
errHand:
   MsgBox err.Description
End Function

Private Function ExecuteGrantObject(ByVal objColGrangObject As Collection)
    '******************************************************************************************************************
    '���ܣ�ִ����Ȩ
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Long
    Dim intCount As Long
    Dim strTemp As String
    Dim varTemp As Variant
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim strKey As String
    Dim blnAllow As Boolean
    Dim intPercent As Integer
    Dim intSvrPercent As Integer
    
    Dim objCollection As New Collection
    
    On Error GoTo errHand
    
    intCount = objColGrangObject.Count
    If intCount > 0 Then
        
        RaiseEvent BeforeProgress("���������ɫȨ��", 100)
        
        mstrSQL = "select Upper(Owner) As ������,Upper(table_name) As ����,Upper(Privilege) As Ȩ�� from role_tab_privs where role='" & mstrRole & "' And table_name Not Like 'BIN$%'"
        rs.Open mstrSQL, mcnOracle
        If rs.BOF = False Then
            Do While Not rs.EOF
                objCollection.Add "1", "K" & rs("������").value & "'" & rs("����").value & "'" & rs("Ȩ��").value
                rs.MoveNext
            Loop
        End If
        
        For intLoop = 1 To intCount
            strTemp = objColGrangObject.Item(intLoop)
            varTemp = Split(strTemp, "'")
            strKey = "K" & UCase(CStr(varTemp(0))) & "'" & UCase(CStr(varTemp(1))) & "'" & UCase(CStr(varTemp(2)))
            
            If IsCollectionElement(objCollection, strKey) = False Then
                
                strSQL = "grant " & CStr(varTemp(2)) & " on " & CStr(varTemp(0)) & "." & CStr(varTemp(1)) & " to " & mstrRole
                
                On Error Resume Next
                err = 0
                Call ExecuteSQL("Grant " & CStr(varTemp(2)) & " on " & CStr(varTemp(0)) & "." & CStr(varTemp(1)) & " to " & mstrRole)
                
                
                If err <> 0 Then
                    mstrFailInfo = mstrFailInfo & vbCrLf & CStr(varTemp(2)) & " on " & CStr(varTemp(0)) & "." & CStr(varTemp(1))
                    err.Clear
                End If
                
                On Error GoTo errHand
            End If
            
            intPercent = 100 * intLoop / intCount
            If intSvrPercent <> intPercent Then
                intSvrPercent = intPercent
                RaiseEvent Progressing(intSvrPercent)
            End If
        Next
        
        RaiseEvent AfterProgress
    End If
    
    ExecuteGrantObject = True
    
    Exit Function
   '------------------------------------------------------------------------------------------------------------------
errHand:
   MsgBox err.Description
'   Resume
End Function

Private Function IsCollectionElement(ByVal objCollection As Collection, ByVal strKey As String) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim strTemp As String
    
    On Error GoTo errHand
    err = 0
    strTemp = objCollection.Item(strKey)
    IsCollectionElement = (err = 0)
    
    Exit Function
    
errHand:
    
End Function

Private Function DeleteGrant(ByVal rsGranded As ADODB.Recordset, ByVal rsPrivs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '���ܣ�ȡ����ǰ����Ķ�ģ��ķ���Ȩ��
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Long
    Dim intCount As Long
    Dim strKey As String
    Dim intPercent As Integer
    Dim intSvrPercent As Integer
    Dim strTemp As String
    
    On Error GoTo errHand
    
    mrsModule.Filter = "��Ȩ�� = 0 "
    intCount = mrsModule.RecordCount
    
    If intCount > 0 Then
        
        RaiseEvent BeforeProgress("����ɾ����ɫȨ��", 100)
        
        Do Until mrsModule.EOF

            strKey = "K" & mrsModule("ϵͳ").value & "'" & mrsModule("���").value & "'" & mrsModule("����").value
            If IsCollectionElement(mobjCollectionGranded, strKey) = True Then
                '��ǰ���ڹ��ģ����
                If strTemp <> "" Then strTemp = strTemp & "''"
                strTemp = strTemp & mrsModule("ϵͳ").value & "''" & mrsModule("���").value & "''" & mrsModule("����").value

                If LenB(strTemp) > 1800 Then
                    mstrSQL = "zl_zlRoleGrant_BatchDelete('" & mstrRole & "','" & strTemp & "')"
                    Call ExecuteProcedure(mstrSQL, "��Ȩ")
                    strTemp = ""
                End If
                
            End If
            
            intLoop = intLoop + 1
            intPercent = 100 * intLoop / intCount
            If intSvrPercent <> intPercent Then
                intSvrPercent = intPercent
                RaiseEvent Progressing(intSvrPercent)
            End If
                        
            mrsModule.MoveNext
        Loop
        
        If strTemp <> "" Then
            mstrSQL = "zl_zlRoleGrant_BatchDelete('" & mstrRole & "','" & strTemp & "')"
            Call ExecuteProcedure(mstrSQL, "��Ȩ")
        End If
        
        RaiseEvent AfterProgress
        
    End If
    
    DeleteGrant = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
'    Resume
End Function

Private Function SaveGrant(ByVal rsGranded As ADODB.Recordset, ByVal rsPrivs As ADODB.Recordset) As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim intLoop As Long
    Dim intCount As Long
    Dim strKey As String
    Dim intPercent As Integer
    Dim intSvrPercent As Integer
    Dim strTemp As String
    
    On Error GoTo errHand
    
    mrsModule.Filter = "��Ȩ�� = 1 "
    intCount = mrsModule.RecordCount
    If intCount > 0 Then
        
        RaiseEvent BeforeProgress("���ڱ����ɫȨ��", 100)
        
        Do Until mrsModule.EOF

            strKey = "K" & mrsModule("ϵͳ").value & "'" & mrsModule("���").value & "'" & mrsModule("����").value
            If IsCollectionElement(mobjCollectionGranded, strKey) = False Then
                '��ǰû�ڹ��ģ�����
                
                If strTemp <> "" Then strTemp = strTemp & "''"
                strTemp = strTemp & IIf(mrsModule("ϵͳ") = 0, "null", mrsModule("ϵͳ")) & "''" & mrsModule("���").value & "''" & mrsModule("����").value
                
                If LenB(strTemp) > 1800 Then
                    mstrSQL = "zl_zlRoleGrant_BatchInsert('" & mstrRole & "','" & strTemp & "')"
                    Call ExecuteProcedure(mstrSQL, "��Ȩ")
                    strTemp = ""
                End If

'                '��մ�����Ϊ���ܲ�����䲻������ִ��(����ΨһԼ��)
'                Err.Clear
            End If
            
            intLoop = intLoop + 1
            intPercent = 100 * intLoop / intCount
            If intSvrPercent <> intPercent Then
                intSvrPercent = intPercent
                RaiseEvent Progressing(intSvrPercent)
            End If
                        
            mrsModule.MoveNext
        Loop
        
        If strTemp <> "" Then
            mstrSQL = "zl_zlRoleGrant_BatchInsert('" & mstrRole & "','" & strTemp & "')"
            Call ExecuteProcedure(mstrSQL, "��Ȩ")
        End If
                
        RaiseEvent AfterProgress
        
    End If
    
    SaveGrant = True
    
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    MsgBox err.Description
'    Resume
End Function

Private Sub AddPrivilegeData(ByVal strPrivilege As String, Optional ByVal bytMode As Byte = 1)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    On Error Resume Next
    
    Select Case bytMode
    Case 1
        mobjColGrantObject.Add strPrivilege, "K" & strPrivilege
    Case 2
        mobjColRevokeObject.Add strPrivilege, "K" & strPrivilege
    End Select
End Sub

Private Sub ExecuteSQL(ByVal strSQL As String)
    mcnOracle.Execute strSQL
End Sub

Private Sub Class_Initialize()
    Set mobjCollectionMoudal = New Collection
    Set mobjCollectionGranded = New Collection
    Set mobjColGrantObject = New Collection
    Set mobjColRevokeObject = New Collection
    
    Set mrsPrivs = New ADODB.Recordset
    Set mrsGranded = New ADODB.Recordset           '������ǰ��Ȩ��ģ��
    Set mrsFuncPars = New ADODB.Recordset          '�������������ʵı�
End Sub

Private Sub Class_Terminate()
    Set mobjCollectionMoudal = Nothing
    Set mobjCollectionGranded = Nothing
    Set mobjColGrantObject = Nothing
    Set mobjColRevokeObject = Nothing
    
    Set mrsPrivs = Nothing
    Set mrsGranded = Nothing
    Set mrsFuncPars = Nothing
    Set mrsModule = Nothing
    Set mrsTable = Nothing
    Set mrsFunction = Nothing
    Set mcnOracle = Nothing
End Sub
