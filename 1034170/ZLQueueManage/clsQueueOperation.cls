VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQueueOperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mintValidDays           As Integer '������Ч����
Private mintCallWay             As Integer '���з�ʽ 0-ϵͳĬ�ϣ�1-΢������
Private mintCallCount           As Integer '���д��� Ĭ��Ϊ1
Private mintCallSpeed           As Integer '�����ٶ�
Private mintCallLen             As Integer '����ʱ�䳤�ȣ�Ĭ��10��
Private mstrCallStation         As String  '������վ����,Ĭ�ϵ�ǰվ��
Private mstrCallTarget          As String  '����ǰ���ò��˵ļ��Ŀ�ĵ�
Private mstrLocalStation        As String  '����վ����
Private mintWorkType            As String  'ҵ������
Private mcnOracle               As ADODB.Connection
Private mlngSysNo               As Long
Private mstrCustomOrder         As String  '�Զ�������
Private mstrDefQueryCols        As String  'Ĭ�ϲ�ѯ�У�Ĭ��Ϊ�ŶӽкŶ��б��е�������
Private mstrCallScript          As String
Private mblnUseVbsCall          As Boolean

Public Event OnCallPreBefore(ByVal lngQueueId As Long, ByVal lngCallWay As Long, ByRef strCallContext As String, blnCancel As Boolean)
Public Event OnCallPreAfter(ByVal lngQueueId As Long, ByVal lngCallWay As Long)
Public Event OnWorkBefore(ByVal lngQueueId As Long, ByVal lngOperationType As Long, blnCancel As Boolean)
Public Event OnWorkAfter(ByVal lngQueueId As Long, ByVal lngOperationType As Long)
Public Event OnCallingBefore(ByVal lngCallId As Long, ByVal strCallContext As String, blnCancel As Boolean)
Public Event OnCallingAfter(ByVal lngCallId As Long, ByVal strCallContext As String)

'�������ŵĺ���
Private Declare Function StartTextPlay Lib "StrSound.dll" (ByVal PlayText As String, ByVal intxx As Integer) As Long
Private Declare Function StopPlayStr Lib "StrSound" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long



'��������
Public Enum OperationType
    otStart = 0       '�Ŷ�
    otPause = 1       '��ͣ
    otAbstain = 2     '����
    otRestore = 3     '����
    otComplete = 4    '���
    otPriority = 5    '����
    otInsertQueue = 6 '���
    otDiagnose = 7   '����
End Enum

Public Sub InitQueue(cnOracle As ADODB.Connection, ByVal lngSysNo As Long, ByVal intWorkType As Integer)
'���ܣ���ʼ������
    Set mcnOracle = cnOracle
    
    mlngSysNo = lngSysNo
    mintWorkType = intWorkType
    
    mintValidDays = 0
    mintCallWay = 0
    mintCallCount = 1
    mintCallSpeed = 10
    mintCallLen = 10
    mstrCallScript = ""
    mblnUseVbsCall = False
    
    mstrCustomOrder = GetCustomOrderWhereStr
    mstrLocalStation = AnalyseComputer
    mstrCallStation = AnalyseComputer
    mstrDefQueryCols = GetAllCols

End Sub

Public Function InsertQueue(ByVal strQueueName As String, _
                            Optional ByVal strQueueNo As String, _
                            Optional ByVal strBusinessId As String, _
                            Optional ByVal strName As String, _
                            Optional ByVal strTarget As String, _
                            Optional ByVal strCustomOrder As String, _
                            Optional ByVal strExpandData As String) As Long
'strQueueName: Ҫ����Ķ�������
'strQueueNo: �ŶӺ���
'strBusinessId: ҵ��ID
'strName: ����
'strTarget: Ŀ�ĵ�
'strCustomOrder: �Զ����Ŷ����
'strExpandData:��չ���� ��ʽΪ/* "����='����', �Ա�='Ů',����=10"*/
'���ܣ����ز������Ŷ�ID��ʧ�ܷ���-1
    Dim strSql          As String
    Dim rsQueue         As ADODB.Recordset
    
    On Error GoTo ErrorHand
    
    InsertQueue = -1
    
    strSql = "select zl_�ŶӽкŶ���_����([1],[2],[3],[4],[5],[6],[7],[8]) as �Ŷ�id from dual"
    
    Set rsQueue = zlDatabase.OpenSQLRecord(strSql, "���ز������Ŷ�ID", strQueueName, mintWorkType, _
                                        strQueueNo, strBusinessId, strName, strTarget, strCustomOrder, strExpandData)
                                        
    If rsQueue.RecordCount > 0 Then InsertQueue = Nvl(rsQueue!�Ŷ�id)
    
    Exit Function
ErrorHand:
    InsertQueue = -1
End Function

Public Sub DeleteQueue(ByVal lngQueueId As Long)
'lngQueueID: �Ŷ�ID
'���ܣ�ɾ����������
    Dim strSql          As String
    
    strSql = "ZL_�ŶӽкŶ���_ɾ��(" & lngQueueId & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "ɾ����������")

End Sub

Public Sub ClearQueueforQueueName(ByVal strQueueName As String)
'strQueueName: ��������
'���ܣ������Ч������ָ�����еĶ�������
    Dim strSql          As String

    strSql = "ZL_�ŶӽкŶ���_���('" & strQueueName & "'," & ValidDays & "," & mintWorkType & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "�����������")

End Sub

Public Sub ClearQueueData()
'intWorkType: ҵ������
'intValidDays: ��Ч����
'���ܣ�����������Ч��������Ŷ�����

    Dim strSql As String
    On Error GoTo errHandle

    strSql = "ZL_�Ŷ����(" & mintWorkType & "," & ValidDays & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "����Ŷ�����")
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Public Sub UpdateQueue(ByVal lngQueueId As Long, ByVal strExpandData As String)
'lngQueueID: �Ŷ�ID
'arrUpdateColData: �ĸ�ʽΪ/* "����='����', �Ա�='Ů',����=10"*/
'���ܣ����¶�������
    Dim strSql          As String
    
    strSql = "ZL_�ŶӽкŶ���_����(" & lngQueueId & ",'" & strExpandData & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "���¶�������")
    
End Sub

Public Function GetQueueInf(ByVal lngQueueId As Long, Optional strQueueFields As String) As ADODB.Recordset
'lngQueueId: �Ŷ�ID
'strQueueFields:��Ҫ��ѯ���У���ʽ��'�������Ա����䣬...'
'���ܣ���ȡ��ǰ�Ŷ�����
    Dim rsRecord        As ADODB.Recordset
    
    On Error GoTo ErrorHand:
    Call SerachQueue("id=" & lngQueueId, rsRecord, strQueueFields)
    Set GetQueueInf = rsRecord
    
    Exit Function
ErrorHand:
    Set GetQueueInf = Nothing
End Function

Public Sub WriteTarget(ByVal lngQueueId As Long)
'lngQueueId: �Ŷ�ID
'���ܣ�д����Ŀ�ĵ�
    Dim strSql          As String
    
    strSql = "ZL_�ŶӽкŶ���_����Ŀ�ĵ�(" & lngQueueId & ",'" & CallTarget & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "д����Ŀ�ĵ�")

End Sub


Private Function GetAllCols() As String
'��ȡ�Ŷӽкű��е������ֶ���
On Error GoTo errHandle
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim strCols As String
    
    GetAllCols = "*"
    
    strSql = "select Column_Name as col from User_Tab_Columns where Table_Name='�ŶӽкŶ���'"
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "��ѯ�Ŷӽк���")
    If rsTmp.RecordCount <= 0 Then Exit Function
    
    strCols = ""
    rsTmp.MoveFirst
    
    While Not rsTmp.EOF
        If strCols <> "" Then strCols = strCols & ","
        strCols = strCols & Nvl(rsTmp!Col)
        
        rsTmp.MoveNext
    Wend
    
    GetAllCols = strCols
Exit Function
errHandle:
    GetAllCols = "*"
End Function


Public Sub OrderCall(ByVal strQueueName As String)
'strQueueName: ��������
'���ܣ�˳�����
    Dim strSql                  As String
    Dim blnCancel               As Boolean
    Dim rsQueue                 As ADODB.Recordset
    Dim rsTemp                  As ADODB.Recordset
    Dim strCallContext          As String
    
    blnCancel = False
    strCallContext = ""
    
    strSql = "select " & mstrDefQueryCols & _
             " from �ŶӽкŶ��� where ��������=[1] and ҵ������=[2] and �Ŷ�״̬=0 and rownum=1" & CustomOrder
             
    Set rsQueue = zlDatabase.OpenSQLRecord(strSql, "˳�����", strQueueName, mintWorkType)
    
    If rsQueue.RecordCount <= 0 Then Exit Sub
    
    rsQueue.MoveFirst
        
    RaiseEvent OnCallPreBefore(Nvl(rsQueue!Id), CallWay, strCallContext, blnCancel)
    
    If blnCancel = True Then Exit Sub
    
    If strCallContext = "" Then
        '��ȡ��������
        strSql = "select zl_�ŶӽкŶ���_��ȡ��������([1]) as �������� from dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ��������", Val(rsQueue!Id))
        strCallContext = Nvl(rsTemp!��������)
    End If
    
    strSql = "ZL_�ŶӽкŶ���_Call(" & Nvl(rsQueue!Id) & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "���ݺ����޸��Ŷ�״̬")
    
    strSql = "Zl_�Ŷ���������_����(" & Nvl(rsQueue!Id) & ",'" & LocalStation & "','" & strCallContext & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "д���������")
    
    RaiseEvent OnCallPreAfter(Nvl(rsQueue!Id), CallWay)

End Sub

Public Sub SpecifiedCall(ByVal lngQueueId As Long)
'lngQueueId: �Ŷ�ID
'���ܣ�ֱ�Ӻ���
    Dim strSql              As String
    Dim blnCancel           As Boolean
    Dim strCallContext      As String
    Dim rsTemp              As ADODB.Recordset

    blnCancel = False
    strCallContext = ""
    
    RaiseEvent OnCallPreBefore(lngQueueId, CallWay, strCallContext, blnCancel)
    
    If blnCancel = True Then Exit Sub
    
    If strCallContext = "" Then
        '��ȡ��������
        strSql = "select zl_�ŶӽкŶ���_��ȡ��������([1]) as �������� from dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ��������", lngQueueId)
        strCallContext = Nvl(rsTemp!��������)
    End If
    
    strSql = "ZL_�ŶӽкŶ���_Call(" & lngQueueId & ")"
    Call zlDatabase.ExecuteProcedure(strSql, "���ݺ����޸��Ŷ�״̬")
    
    strSql = "Zl_�Ŷ���������_����(" & lngQueueId & ",'" & LocalStation & "','" & strCallContext & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "д���������")
    
    RaiseEvent OnCallPreAfter(lngQueueId, CallWay)
    
End Sub

Public Sub BroadcastCall(ByVal lngQueueId As Long)
'lngQueueId: �Ŷ�ID
'���ܣ��㲥����
    Dim strSql              As String
    Dim blnCancel           As Boolean
    Dim strCallContext      As String
    Dim rsTemp              As ADODB.Recordset

    blnCancel = False
    strCallContext = ""
    
    RaiseEvent OnCallPreBefore(lngQueueId, CallWay, strCallContext, blnCancel)
    
    If blnCancel = True Then Exit Sub
    
    If strCallContext = "" Then
        '��ȡ��������
        strSql = "select zl_�ŶӽкŶ���_��ȡ��������([1]) as �������� from dual"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "��ȡ��������", lngQueueId)
        strCallContext = Nvl(rsTemp!��������)
    End If
    
    strSql = "Zl_�Ŷ���������_����(" & lngQueueId & ",'" & LocalStation & "','" & strCallContext & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "д���������")
    
    RaiseEvent OnCallPreAfter(lngQueueId, CallWay)
    
End Sub


Private Function StartCall(ByVal strCallingContext As String) As Boolean
'lngQueueId������id
'strCallingContext:��������
'���ܣ���ɺ���
    Dim start                   As Date
    Dim intCurCallCount         As Long
    
    StartCall = False
    If strCallingContext = "" Then Exit Function
                        
    intCurCallCount = 0
    While (intCurCallCount < CallCount)
        
        start = Timer
        If CallWay = 1 Then
            Call MSSoundPlay(strCallingContext, CallSpeed)
        Else
            Call StartTextPlay(strCallingContext, CallSpeed * 10)
        End If
        
        intCurCallCount = intCurCallCount + 1
                                        
        Do While Timer < start + CallLen
            Call Sleep(5)
            
            DoEvents
        Loop
    Wend
    
    StartCall = True
End Function

Private Sub MSSoundPlay(ByVal strConnetxt As String, ByVal lngSoundSpeed As Long)
    Dim objVoice        As Object
    
    On Error Resume Next
    
    If objVoice Is Nothing Then
        Set objVoice = CreateObject("SAPI.SpVoice")
    End If
    
    objVoice.Rate = lngSoundSpeed   '�ٶ�:-10,10  0
    objVoice.Volume = 100 '����:0,100   100
    objVoice.Speak strConnetxt, 1

End Sub

Public Sub CustomCall(ByVal strContext As String)
'strContext: �������� '
'���ܣ��Զ�����
    Dim strSql          As String
    Dim blnCancel       As Boolean
    
'    id(��������),���ݣ�վ�㣬ʱ��(ϵͳʱ��)
    strSql = "Zl_�Ŷ���������_�Զ������('" & strContext & "','" & CallStation & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "�Զ�����")
        
End Sub

Public Sub StartQueue(ByVal lngQueueId As Long)
'lngQueueId: �Ŷ�ID
'���ܣ���ʼ�Ŷ�
    Dim strSql                  As String
    Dim blnCancel               As Boolean
    
    blnCancel = False
    
    RaiseEvent OnWorkBefore(lngQueueId, OperationType.otStart, blnCancel)
    
    If blnCancel = True Then Exit Sub
    
    strSql = "zl_�ŶӽкŶ���_�Ŷ�״̬����(" & lngQueueId & ",0)"
    Call zlDatabase.ExecuteProcedure(strSql, "��ʼ�Ŷ�")
    
    RaiseEvent OnWorkAfter(lngQueueId, OperationType.otStart)

End Sub

Public Sub PauseQueue(ByVal lngQueueId As Long)
'lngQueueID: �Ŷ�ID
'���ܣ���ͣ�Ŷ�
    Dim strSql                  As String
    Dim blnCancel               As Boolean
    
    blnCancel = False

    RaiseEvent OnWorkBefore(lngQueueId, OperationType.otPause, blnCancel)
    
    If blnCancel = True Then Exit Sub
    
    strSql = "zl_�ŶӽкŶ���_�Ŷ�״̬����(" & lngQueueId & ",3)"
    Call zlDatabase.ExecuteProcedure(strSql, "��ͣ�Ŷ�")
    
    RaiseEvent OnWorkAfter(lngQueueId, OperationType.otPause)

End Sub

Public Sub AbstainQueue(ByVal lngQueueId As Long)
'lngQueueID: �Ŷ�ID
'���ܣ������Ŷ�
    Dim strSql                      As String
    Dim blnCancel                   As Boolean
    
    blnCancel = False
    
    RaiseEvent OnWorkBefore(lngQueueId, OperationType.otAbstain, blnCancel)
    
    If blnCancel = True Then Exit Sub
    
    strSql = "zl_�ŶӽкŶ���_�Ŷ�״̬����(" & lngQueueId & ",2)"
    Call zlDatabase.ExecuteProcedure(strSql, "�����Ŷ�")
    
    RaiseEvent OnWorkAfter(lngQueueId, OperationType.otAbstain)

End Sub

Public Sub RestoreQueue(ByVal lngQueueId As Long, Optional ByVal strCustomOrder As String)
'lngQueueID: �Ŷ�ID
'���ܣ������Ŷӣ��ŵ����λ�ã�
    Dim strSql                  As String
    Dim blnCancel               As Boolean
    
    blnCancel = False
    
    RaiseEvent OnWorkBefore(lngQueueId, OperationType.otRestore, blnCancel)
    
    If blnCancel = True Then Exit Sub
    
    '���¼����Ŷ����
    strSql = "zl_�ŶӽкŶ���_�Ŷ��������(" & lngQueueId & ",'" & strCustomOrder & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "���ü����Ŷ����")
    
    '�ָ��Ŷ�״̬
    strSql = "zl_�ŶӽкŶ���_�Ŷ�״̬����(" & lngQueueId & ",0)"
    Call zlDatabase.ExecuteProcedure(strSql, "�����Ŷ�")


    RaiseEvent OnWorkAfter(lngQueueId, OperationType.otRestore)
    
End Sub

Public Sub CompleteQueue(ByVal lngQueueId As Long)
'lngQueueID: �Ŷ�ID
'���ܣ�����Ŷ�
    Dim strSql                      As String
    Dim blnCancel                   As Boolean
    
    blnCancel = False

    strSql = "zl_�ŶӽкŶ���_�Ŷ�״̬����(" & lngQueueId & ",4)"
    
    RaiseEvent OnWorkBefore(lngQueueId, OperationType.otComplete, blnCancel)
    
    If blnCancel = True Then Exit Sub
    
    Call zlDatabase.ExecuteProcedure(strSql, "����Ŷ�")
    RaiseEvent OnWorkAfter(lngQueueId, OperationType.otComplete)
    
End Sub

Public Sub DoCustomWork(ByVal lngQueueId As Long, ByVal lngOperationType As Long)
'lngQueueId: �Ŷ�ID
'lngOperationType���������ͣ�Ϊ5���Ժ������
'���ܣ�ִ���Զ���ҵ��������������ת���

    Dim blnCancel       As Boolean

    blnCancel = False

    RaiseEvent OnWorkBefore(lngQueueId, lngOperationType, blnCancel)
    
    If blnCancel = True Then Exit Sub
    
    RaiseEvent OnWorkAfter(lngQueueId, lngOperationType)
End Sub

Public Function GetQueueState(ByVal lngQueueId As Long) As Long
'lngQueueID: �Ŷ�ID
'-1-ռλ�У�0-�Ŷ��У�1-�����У�2-�����ţ�3-����ͣ��4-��ɾ���,5-���ã�6-���ã�7-�Ѻ��У�
'���ܣ���ȡ�Ŷ�״̬��ʧ�ܷ���-2
    Dim strSql              As String
    Dim rsQueue             As ADODB.Recordset
    
    On Error GoTo ErrorHand
    GetQueueState = -2
    strSql = "select �Ŷ�״̬ from �ŶӽкŶ��� where id=[1]"
    Set rsQueue = zlDatabase.OpenSQLRecord(strSql, "�Ŷ�״̬", lngQueueId)
    If rsQueue.RecordCount > 0 Then GetQueueState = Nvl(rsQueue!�Ŷ�״̬)
    
    Exit Function
ErrorHand:
    GetQueueState = -2
End Function

Public Sub DoLoopCall()
'���ܣ��������ڱ�վ��������������ݣ���������������
    Dim strSql              As String
    Dim rsQueue             As ADODB.Recordset
    Dim blnCancel           As Boolean
    
    blnCancel = False
    
    strSql = "select id,�������� from �Ŷ���������  where վ��=[1] order by ����ʱ��"
    Set rsQueue = zlDatabase.OpenSQLRecord(strSql, "������������", CallStation)

    Do While Not rsQueue.EOF
        RaiseEvent OnCallingBefore(Nvl(rsQueue!Id), Nvl(rsQueue!��������), blnCancel)
        
        If blnCancel = True Then Exit Sub
        
        '�����Զ�����к�Ͳ�����������
        If UseVbsCall = True Then
            Call DoScriptCall(Nvl(rsQueue!Id), Nvl(rsQueue!��������))
        Else
            If StartCall(Nvl(rsQueue!��������)) = False Then Exit Sub
        End If
        
        RaiseEvent OnCallingAfter(Nvl(rsQueue!Id), Nvl(rsQueue!��������))
        
        '���гɹ���ɾ�����й�������
        strSql = "ZL_�Ŷ���������_DELETE(" & Nvl(rsQueue!Id) & ")"
        zlDatabase.ExecuteProcedure strSql, "�����������"
        
        rsQueue.MoveNext
    Loop
End Sub

Private Sub DoScriptCall(ByVal lngCallId As Long, ByVal strCallContext As String)
'����vbs�ű�ʵ���Զ������
'lngCallId:������id
'strCallContext: ��������
    Dim strMethod As String
    Dim lngIndex As Long
    Dim objCall
   
    If strCallContext = "" Then Exit Sub
    
    Set objCall = CreateObject("ScriptControl")
   
    strMethod = CallScript
    lngIndex = InStr(UCase(strMethod), "FUNCTION")
    
    If lngIndex <= 0 Then
        lngIndex = InStr(UCase(strMethod), "SUB")
        '���ù������Ŀ�ʼλ��
        lngIndex = lngIndex + 4
    Else
        lngIndex = lngIndex + 9
    End If
    
    strMethod = Mid(strMethod, lngIndex, InStr(strMethod, "(") - lngIndex)
    
    objCall.Language = "vbscript"
   
    Call objCall.AddCode(CallScript)
    Call objCall.Run(Trim(strMethod), lngCallId, strCallContext)
End Sub

 
Public Sub ChangeOrder(ByVal lngQueueId As Long, ByVal strNewOrder As String)
'lngQueueID: �Ŷ�ID
'strNewOrder: �µ��Ŷ����
'���ܣ������Ŷ�˳��
    Dim strSql As String
    
    strSql = "zl_�ŶӽкŶ���_�Ŷ���Ÿ���(" & lngQueueId & ",'" & strNewOrder & "')"
    Call zlDatabase.ExecuteProcedure(strSql, "�����Ŷ����")
    
End Sub

Public Function GetOrder(ByVal lngQueueId As Long) As String
'lngQueueID: �Ŷ�ID
'���ܣ����ض�Ӧ���Ŷ���ţ�ʧ�ܷ���-1
    Dim strSql          As String
    Dim rsReocrd        As ADODB.Recordset
    
    On Error GoTo ErrorHand
    GetOrder = ""
    
    strSql = "select �Ŷ���� from �ŶӽкŶ��� where id=[1]"
    Set rsReocrd = zlDatabase.OpenSQLRecord(strSql, "��ȡ�Ŷ����", lngQueueId)
    
    If rsReocrd.RecordCount > 0 Then GetOrder = Nvl(rsReocrd!�Ŷ����)
    
    Exit Function
ErrorHand:
    GetOrder = ""
End Function

Public Function GetCustomOrderWhereStr() As String
'���ܣ���ȡ�Զ�����������
    Dim strSql          As String
    Dim rsQueue         As ADODB.Recordset
    
    On Error GoTo ErrorHand
    
    GetCustomOrderWhereStr = ""
    
    strSql = "select zl_�ŶӽкŶ���_��ȡ����ʽ([1]) as �������� from dual"
    Set rsQueue = zlDatabase.OpenSQLRecord(strSql, "��ȡ�Զ�����������", mintWorkType)
    
    If rsQueue.RecordCount > 0 Then GetCustomOrderWhereStr = Nvl(rsQueue!��������)
    
    Exit Function
ErrorHand:
    GetCustomOrderWhereStr = ""
End Function

Public Function GetInsertOrder(ByVal lngInsertQueueId As Long, _
                               ByVal strOrderNo1 As String, _
                               ByVal strOrderNo2 As String) As String
'���ܣ���ȡ������
    Dim strSql          As String
    Dim rsQueue         As ADODB.Recordset
    
On Error GoTo ErrorHand
    GetInsertOrder = ""
    
    strSql = "select zl_�ŶӽкŶ���_��ȡ������([1],[2],[3]) as ������ from dual"
    Set rsQueue = zlDatabase.OpenSQLRecord(strSql, "��ȡ������", lngInsertQueueId, strOrderNo1, strOrderNo2)
    
    If rsQueue.RecordCount > 0 Then GetInsertOrder = Nvl(rsQueue!������)
        
    Exit Function
ErrorHand:
    GetInsertOrder = ""
End Function


Public Function GetMinOrder(ByVal strQueueName As String) As String
'strQueueName: ��������
'���ܣ���ȡ��ǰ������С�ŶӺ���
    Dim strSql          As String
    Dim rsReocrd        As ADODB.Recordset
    
On Error GoTo ErrorHand
    GetMinOrder = ""
    
    
    strSql = "select �Ŷ���� from (select �Ŷ���� from �ŶӽкŶ��� where ��������=[1] and ҵ������=[2] " & CustomOrder & " ) a where rownum =1"
    
    Set rsReocrd = zlDatabase.OpenSQLRecord(strSql, "��ȡ��ǰ������С�ŶӺ���", strQueueName, mintWorkType)
    
    If rsReocrd.RecordCount > 0 Then GetMinOrder = Nvl(rsReocrd!�Ŷ����)
    
    Exit Function
ErrorHand:
    GetMinOrder = ""
End Function

Public Function GetMaxQueueNo(Optional ByVal strFilter As String) As String
'strFilter:��ȡ��������Զ���������Ĭ��Ϊ�ձ�ʾ��ǰҵ���µ�����ŶӺ���,��ʽΪ/*"�Ա�='��',����=10,..."*/
'��: '��������=������'�����ȡ��������Ϊ�����Ƶĵ�ǰ�����룻���Ŷӱ��=�ڰ׳�'�����ȡ�Ŷӱ��Ϊ�ڰ׳��ĵ�ǰ������
'���ܣ���ȡָ�������µ�����ŶӺ��루�ŶӺţ�
    Dim strSql          As String
    Dim strNewFilter    As String
    Dim varData(20)     As Variant
    Dim rsQueue         As ADODB.Recordset
   
On Error GoTo ErrorHand
    GetMaxQueueNo = ""
    
    If strFilter = "" Then
        strFilter = "ҵ������=[1]"
        varData(0) = mintWorkType
    Else
        Call GetFilterCfg(strFilter, strNewFilter, varData)
    End If
    
    strSql = "select max(to_number(�ŶӺ���)) as �ŶӺ��� from �ŶӽкŶ��� where " & strFilter
        
    Set rsQueue = zlDatabase.OpenSQLRecord(strSql, "��ȡָ�������µ�����ŶӺ���", varData(0), varData(1), varData(2), varData(3), _
                                                    varData(4), varData(5), varData(6), varData(7), varData(8), varData(9), _
                                                    varData(10), varData(11), varData(12), varData(13), varData(14), varData(15), _
                                                    varData(16), varData(17), varData(18), varData(19), varData(20))
    
    If rsQueue.RecordCount > 0 Then GetMaxQueueNo = Nvl(rsQueue!�ŶӺ���)
    
    Exit Function
ErrorHand:
    GetMaxQueueNo = ""
End Function

Public Function SerachQueue(ByVal strFilter As String, rsRecord As ADODB.Recordset, Optional strQueueFields As String) As Long
'strFilter: ��������,��ʽΪ/*"�Ա�='��',����=10,..."*/
'���ؼ��������Ŷ�ID
'���ܣ����������������������Ŷ�ID,ʧ�ܷ���-1
    Dim strSql          As String
    Dim strNewFilter As String
    Dim varData(20)     As Variant

On Error GoTo ErrorHand
    
    SerachQueue = -1
    Set rsRecord = Nothing
    
    If strFilter = "" Then Exit Function
    
    Call GetFilterCfg(strFilter, strNewFilter, varData)
    
    If strQueueFields <> "" Then
        strSql = "select id," & strQueueFields & " from �ŶӽкŶ��� where " & strNewFilter
    Else
        strSql = "select " & mstrDefQueryCols & " from �ŶӽкŶ��� where " & strFilter
    End If

    Set rsRecord = zlDatabase.OpenSQLRecord(strSql, "���ؼ��������Ŷ�ID", varData(0), varData(1), varData(2), varData(3), _
                                                    varData(4), varData(5), varData(6), varData(7), varData(8), varData(9), _
                                                    varData(10), varData(11), varData(12), varData(13), varData(14), varData(15), _
                                                    varData(16), varData(17), varData(18), varData(19), varData(20))
    If rsRecord.RecordCount > 0 Then SerachQueue = Nvl(rsRecord!Id)
    
    Exit Function
ErrorHand:
    SerachQueue = -1
End Function

Private Sub GetFilterCfg(ByVal strFilter As String, ByRef strNewFilter As String, ByRef varParament() As Variant)
    Dim i               As Integer
    Dim varPara(20)     As Variant
    Dim varRela(20)     As Variant
    Dim strResult       As String
    Dim strValue        As String
    
    strFilter = Trim(strFilter) & " "
    strResult = strFilter
    strFilter = ""
    
    Do While InStr(strResult, "=") > 0
        varPara(i) = Mid(strResult, 1, InStr(strResult, "=") - 1)
        
        strValue = Trim(Mid(strResult, InStr(strResult, "=") + 1, InStr(strResult, " ") - InStr(strResult, "=") - 1))
            
        If IsNumeric(strValue) = True Then
            varParament(i) = Val(strValue)
        ElseIf IsDate(strValue) = True Then
            varParament(i) = CDate(strValue)
        Else
            If InStr(strValue, "(") > 0 Then
                varParament(i) = Mid(strValue, 3, Len(strValue) - 4)
            Else
                varParament(i) = Mid(strValue, 2, Len(strValue) - 2)
            End If
        End If
        
        strResult = LTrim(Mid(strResult, InStr(strResult, "=") + Len(strValue) + 1))
        
        If strResult <> "" Then
            varRela(i) = Mid(strResult, 1, InStr(strResult, " ") - 1)
            strResult = Mid(strResult, InStr(strResult, " ") + 1)
        End If
        
        strFilter = strFilter & varPara(i) & "=[" & i + 1 & "] " & varRela(i) & " "
        i = i + 1
    Loop
    
    strNewFilter = Trim(strFilter)
End Sub

Private Function AnalyseComputer() As String
'��ȡ���ؼ��������
    Dim strComputer As String * 256
    Call GetComputerName(strComputer, 255)
    AnalyseComputer = strComputer
    AnalyseComputer = Trim(Replace(AnalyseComputer, Chr(0), ""))
End Function

Private Function Nvl(ByVal varValue As Variant, Optional DefaultValue As Variant = "") As Variant
'���ܣ��൱��Oracle��NVL����Nullֵ�ĳ�����һ��Ԥ��ֵ
    Nvl = IIf(IsNull(varValue), DefaultValue, varValue)
End Function


'��ʼ��������ֵ
Private Sub Class_Initialize()
'    mintCallCount = 1
'    mintCallLen = 10
End Sub

'���ָ������ǰ������,��������Ч����
Public Property Get ValidDays() As Integer
    ValidDays = mintValidDays
End Property

Public Property Let ValidDays(ByVal vNewValue As Integer)
    mintValidDays = vNewValue
End Property

'���з�ʽ
Public Property Get CallWay() As Integer
    CallWay = mintCallWay
End Property

Public Property Let CallWay(ByVal vNewValue As Integer)
    mintCallWay = vNewValue
End Property

'���д���
Public Property Get CallCount() As Integer
    CallCount = mintCallCount
End Property

Public Property Let CallCount(ByVal vNewValue As Integer)
    mintCallCount = vNewValue
End Property

'�����ٶ�
Public Property Get CallSpeed() As Integer
    CallSpeed = mintCallSpeed
End Property

Public Property Let CallSpeed(ByVal vNewValue As Integer)
    mintCallSpeed = vNewValue
End Property

'����ʱ�䳤��
Public Property Get CallLen() As Integer
    CallLen = mintCallLen
End Property

Public Property Let CallLen(ByVal vNewValue As Integer)
    mintCallLen = vNewValue
End Property

'����վ����
Public Property Get LocalStation() As Variant
    LocalStation = mstrLocalStation
End Property

Public Property Let LocalStation(ByVal vNewValue As Variant)
    mstrLocalStation = vNewValue
End Property

'������վ����
Public Property Get CallStation() As String
    CallStation = mstrCallStation
End Property

Public Property Let CallStation(ByVal vNewValue As String)
    mstrCallStation = vNewValue
End Property

'����ǰ���ò��˵ļ��Ŀ�ĵ�
Public Property Get CallTarget() As String
    CallTarget = mstrCallTarget
End Property

Public Property Let CallTarget(ByVal vNewValue As String)
    mstrCallTarget = vNewValue
End Property

'�Զ�������
Public Property Get CustomOrder() As String
    CustomOrder = mstrCustomOrder
End Property

Public Property Let CustomOrder(ByVal vNewValue As String)
    mstrCustomOrder = vNewValue
End Property

'Ĭ�ϲ�ѯ��
Public Property Get DefQueryCols() As String
    DefQueryCols = mstrDefQueryCols
End Property

Public Property Get CallScript() As String
    CallScript = mstrCallScript
End Property

Public Property Let CallScript(ByVal vNewValue As String)
    mstrCallScript = vNewValue
End Property

Public Property Get UseVbsCall() As Boolean
    UseVbsCall = mblnUseVbsCall
End Property

Public Property Let UseVbsCall(ByVal vNewValue As Boolean)
    UseVbsCall = vNewValue
End Property



