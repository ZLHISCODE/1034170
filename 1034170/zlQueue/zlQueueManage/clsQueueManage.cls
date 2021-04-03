VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQueueManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public VERSION As String

Private WithEvents mfrmQueueStation As frmQueueStation
Attribute mfrmQueueStation.VB_VarHelpID = -1

Public Event OnQueueExecuteBefore(ByVal strҵ��ID As String, ByVal byt�������� As Byte, blnCancel As Boolean, strNewQueueName As String)
Public Event OnQueueExecuteAfter(ByVal strҵ��ID As String, ByVal byt�������� As Byte)
Public Event OnRecevieDiagnose(ByVal strҵ��ID As String, ByVal lngҵ������ As Long)
Public Event OnSelectionChanged(ByVal blnIsCallingList As Boolean, objReportRow As Object, cbrMain As Object)


Property Get Enable() As Boolean
    Enable = mfrmQueueStation.Enabled
End Property

Property Let Enable(value As Boolean)
    mfrmQueueStation.Enabled = value
End Property

'�Ƿ���ʾ���ҹ�����
Public Property Get IsShowFindTools() As Boolean
    IsShowFindTools = mfrmQueueStation.mblnIsShowFindTools
End Property

Public Property Let IsShowFindTools(ByVal vNewValue As Boolean)
    mfrmQueueStation.mblnIsShowFindTools = vNewValue
End Property

Public Function FindQueue(ByVal strLocateType As String, ByVal strLocateValue As String) As Boolean
    On Error GoTo errHandle
    
    FindQueue = mfrmQueueStation.FindQueue(strLocateType, strLocateValue)
    
    Exit Function
errHandle:
    If ErrCenter = 1 Then Resume
End Function

'Public Sub ShwoMeTest(cnOracle As ADODB.Connection)
'
'    Dim strQueue(2) As String
'
'    '���е��±��1��ʼ
'    strQueue(1) = "����1"
'    strQueue(2) = "����2"
'
'    mfrmQueueStation.zlShowMe cnOracle, strQueue, "����1", 0
'End Sub


Public Sub zlInitVar(cnOracle As ADODB.Connection, Optional lngSys As Long = 100, _
    Optional intҵ������ As Integer = 0, Optional intValidDays As Integer = 1, _
    Optional strPrivs As String = "", Optional strOption As String = "", Optional blnIsGroup As Boolean = True)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���ʼ��ϵͳ����
    '��Σ�strOption-����,�Ժ���չ
    '���ƣ����˺�
    '���ڣ�2010-06-11 11:01:09
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    If Not mfrmQueueStation Is Nothing Then
        Call mfrmQueueStation.zlInitVar(cnOracle, lngSys, intҵ������, intValidDays, strPrivs, strOption, blnIsGroup)
    End If
End Sub

'���ý���
Public Sub SetFocus()
    On Error Resume Next
    
    If Not mfrmQueueStation Is Nothing Then
        If mfrmQueueStation.Visible Then
            If mfrmQueueStation.rptQueueList.Visible Then Call mfrmQueueStation.rptQueueList.SetFocus
        End If
    End If
    
    On Error GoTo 0
End Sub


Public Function CloseWindows() As Boolean
    '--------------------------------------
    '����:�ر������Ӵ���
    '--------------------------------------
    Dim frmThis As Form
    
    On Error Resume Next
    
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = (Forms.Count = 0)
End Function


Public Sub zlCommandBarSet(ByVal intFuncType As Integer, ByVal blnUseState As Boolean)
'************************************************************************************
'
'���ù���״̬
'
'intFuncType���������� 0-���1-ֱ��/˳����2-���� ��3-��ͣ��4-��ɾ��5,-�㲥
'blnUseState���Ƿ�����
'
'************************************************************************************
    If Not mfrmQueueStation Is Nothing Then
        Call mfrmQueueStation.zlCommandBarSet(intFuncType, blnUseState)
    End If
End Sub


Private Sub Class_Initialize()
    VERSION = App.Major & "." & App.Minor & "." & App.Revision
    Set mfrmQueueStation = New frmQueueStation
    mfrmQueueStation.mblnIsShowFindTools = True
End Sub


Private Sub Class_Terminate()
    Unload mfrmQueueStation
    Set mfrmQueueStation = Nothing
End Sub

Public Function zlGetForm() As Object
    Set zlGetForm = mfrmQueueStation
End Function

Public Sub zlDefCommandBars(ByVal CommandBars As Object)
    Call mfrmQueueStation.zlDefCommandBars(CommandBars)
End Sub

Public Sub zlUpdateCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call mfrmQueueStation.zlUpdateCommandBars(Control)
End Sub

Public Sub zlExecuteCommandBars(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Call mfrmQueueStation.zlExecuteCommandBars(Control)
End Sub

Public Sub zlSetToolIcon(ByVal intIconSize As Integer, ByVal blnIsDisplayText As Boolean)
  Call mfrmQueueStation.zlSetToolIcon(intIconSize, blnIsDisplayText)
End Sub

Public Function zlRefresh(str��������() As String, ByVal strCur�������� As String, ByVal strCurҵ��ID As String, _
    Optional str���� As String = "", Optional strҽ�� As String = "", Optional strExcludeData As String = "", Optional intViewDataType As Integer = 0) As Long
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ�����ˢ��ָ��ҽ��id�ı������ݣ�����������ṩ�༭����
    '��Σ�str��������():�����ָ����������(��1��ʼ)
    '         strCur��������-��ǰ��������
    '         lngCurҵ��ID-ҵ��ID
    '         str����-����Ϊָ��������,����Ϊ�������:��"һ����,������,..."
    '         strҽ��-����Ϊ�ƶ���ҽ��,���Դ����ҽ��,�ö��ŷָ�,��"����,����,..."
    '         strExcludeData-�Ŷӵ�ָ��ҵ��ID
    '         intViewDataType������ʾ���ͣ�0��ʾ��ǰ�����µ��������ݣ�
    '                                      1��ʾ����Ϊ��ǰ������ҽ������Ϊ�գ�����ҽ���������ڵ�ǰҽ������������Ϊ�պ�ҽ��Ϊ�յ�����
    '                                      2��ʾ����Ϊ��ǰ���Һ�ҽ������Ϊ�ջ�ҽ���������ڵ�ǰҽ��������
    '                                      3��ʾ��ǰҽ��������
    '���ƣ����˺�
    '���ڣ�2010-06-11 20:54:55
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    If Not mfrmQueueStation Is Nothing Then
        'Call DebugMsg("ZLQueueInterface", "ZLRefresh", "str����=" & str���� & ", strҽ��=" & strҽ�� & ", strExcludeData=" & strExcludeData & ", intViewDataType=" & intViewDataType)
        zlRefresh = mfrmQueueStation.zlRefresh(str��������, strCur��������, strCurҵ��ID, str����, strҽ��, strExcludeData, intViewDataType)
    End If
End Function

Private Function To_Date(ByVal dat���� As Date) As String
'����:������е����ڴ�����ORACLE��Ҫ�����ڸ�ʽ��
    To_Date = "To_Date('" & Format(dat����, "YYYY-MM-DD hh:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
End Function

Public Function zlInQueue(ByVal str�������� As String, ByVal lngҵ������ As Long, ByVal strҵ��ID As String, ByVal lng����ID As Long, _
        ByVal str�������� As String, ByVal lng����id As Long, ByVal str���� As String, ByVal strҽ������ As String, _
        Optional str�Ŷӱ�� As String = "", Optional str�ŶӺ��� As String = "", Optional dat�Ŷ����� As Date) As Boolean
'*************************************************************************************
'
'����Ҫ���нкŲ������������
'
'str�������ƣ�
'
'lngҵ�����ͣ�
'
'lngҵ��ID��
'
'lng����ID��
'
'str����������
'
'str���ң�
'
'strҽ��������
'
'str�Ŷӱ�ǣ�
'
'lng�ŶӺ��룺
'
'dat�Ŷ����ڣ�
'
'*************************************************************************************
    
    Dim strSql As String
    Dim datTest As Date
    
    On Error GoTo err
    
    If dat�Ŷ����� = datTest Then
      dat�Ŷ����� = zlDatabase.Currentdate
    End If
        
    strSql = "ZL_�ŶӽкŶ���_INSERT('" & str�������� & "', " & lngҵ������ & ",'" & strҵ��ID & "'," & lng����ID & ",'" & str�ŶӺ��� & "','" & str�Ŷӱ�� & "','" _
             & str�������� & "'," & lng����id & ",'" & str���� & "','" & strҽ������ & "',to_date('" & Format(dat�Ŷ�����, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'))"

    zlDatabase.ExecuteProcedure strSql, "�������"
    zlInQueue = True
    Exit Function
    
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Sub zlQueueExec(ByVal str��ǰ������ As String, ByVal lngҵ������ As Long, ByVal strҵ��ID As String, ByVal byt�������� As Byte)
'************************************************************************************
'˳�������Ŷӵĺ���״̬���и���

'str��ǰ����������Σ��ַ����͡���Ҫ�����Ķ������ơ�
'lngҵ�����ͣ���Σ��������͡�Ŀǰ�����Ŷ�Ϊ0��PACS�Ŷ�Ϊ1��LIS�Ŷ�Ϊ2��
'lngҵ��ID����Σ��ַ����͡�����ŶӼ�¼��ҵ�����ݵĹ�ϵ��������������д����ţ�PACS����дҽ��ID��
'byt�������ͣ���Σ��������͡��кŲ��������� 0-���1-ֱ��/˳����Lngҵ��ID=0Ϊ˳������2-���� ��3-��ͣ��4-��ɾ��5,-�㲥, 6,���
    

'*************************************************************************************
        Call mfrmQueueStation.zlQueueExec(str��ǰ������, lngҵ������, strҵ��ID, byt��������)
End Sub








Public Function zlLISQueueJoin(ByVal str�������� As String, ByVal lngҵ������ As Long) As Long
'************************************************************************************
'����Ҫ���нкŲ�����������ӣ�����������ŶӺ���

'str�������ƣ���Σ��ַ����͡���Ҫ�����Ķ������ơ�
'lngҵ�����ͣ���Σ��������͡�Ŀǰ�����Ŷ�Ϊ0��PACS�Ŷ�Ϊ1��LIS�Ŷ�Ϊ2��
'lng����ID����Σ��������ͣ����ұ��
'dat�Ŷ����ڣ���ѡ��Σ��������ʱ�䡣

    
'���ƣ��޹���
'���ڣ�2012-6-26
'˵��������ʹ����lngWorkOrderID������ҵ��ID�����û�в��ҵ��ŶӺ������ݣ���ô�ӿڷ��ص��ŶӺ���ֵΪ0
'*************************************************************************************

    Dim strSql As String
    Dim datTest As Date
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
        
    strSql = "ZL_�ŶӽкŶ���_INSERT('" & str�������� & "', " & lngҵ������ & "," & Val(Format(zlDatabase.Currentdate, "HHMMSS")) & " ,null,'','','',null,'','',to_date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'))"
    zlDatabase.ExecuteProcedure strSql, "�������"
    
    
    strSql = "select max(�ŶӺ���) as �ŶӺ��� from �ŶӽкŶ��� where ��������=[1] and ҵ������=[2] and �Ŷ�״̬=0 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "�õ���ǰ���˵��Ŷ�������", str��������, lngҵ������)
    
    zlLISQueueJoin = Nvl(rsTemp!�ŶӺ���, 0)
    
    Exit Function
    
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function zlGetQueuePatientCount(ByVal str�������� As String, ByVal lngҵ������ As Long) As Long
'************************************************************************************
'�õ�ָ�������еȴ���������

'str�������ƣ���Σ��ַ����͡���Ҫ�����Ķ������ơ�
'lngҵ�����ͣ���Σ��������͡�Ŀǰ�����Ŷ�Ϊ0��PACS�Ŷ�Ϊ1��LIS�Ŷ�Ϊ2��
    
'���ƣ��޹���
'���ڣ�2012-6-26
'˵�������û�в��ҵ����ݣ���ô�ӿڷ��صĵȴ���������Ϊ0
'*************************************************************************************

 Dim strSql As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo err
    
    strSql = "select count(*) as �ȴ��������� from �ŶӽкŶ��� where ��������=[1] and ҵ������=[2] and �Ŷ�״̬=0"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "�õ��ȴ���������", str��������, lngҵ������)
    
    zlGetQueuePatientCount = Nvl(CLng(rsTemp!�ȴ���������) - 1, 0)
    
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Sub zlQueueParameterSetup(frm������ As Object, lngϵͳ�� As Long)
'************************************************
'���ŶӽкŲ������ý���

'frm�����壺��Σ�������󡣴����ý�����Ҫ�ĸ��������
'lngϵͳ�ţ���Σ��������͡���ǰϵͳ�š�
    
'���ƣ��޹���
'���ڣ�2012-6-26
'˵����ֱ�ӵ��ò������ý���
'************************************************
    
    Call mfrmQueueStation.QueueParameterSetup(frm������, lngϵͳ��)

End Sub


Public Function zlQueueOrderCall(ByVal str��ǰ������ As String, ByVal lngҵ������ As Long, ByVal byt�������� As Byte) As Long
'************************************************************************************
'˳�������Ŷӵĺ���״̬���и���

'str��ǰ����������Σ��ַ����͡���Ҫ�����Ķ������ơ�
'lngҵ�����ͣ���Σ��������͡�Ŀǰ�����Ŷ�Ϊ0��PACS�Ŷ�Ϊ1��LIS�Ŷ�Ϊ2��
'byt�������ͣ���Σ��������͡��кŲ��������� 0-���1-ֱ��/˳����Lngҵ��ID=0Ϊ˳������2-���� ��3-��ͣ��4-��ɾ��5,-�㲥, 6,���
    
'���ƣ��޹���
'���ڣ�2012-6-26
'˵����
'*************************************************************************************
    
    Dim strSql As String
    Dim rsWorkID As ADODB.Recordset
    Dim rsQueueNo As ADODB.Recordset
    
    On Error GoTo err
    
    strSql = "select ҵ��ID from �ŶӽкŶ��� where ��������=[1] and ҵ������=[2] and �Ŷ�״̬=0 and rownum=1 order by �ŶӺ��� "
    Set rsWorkID = zlDatabase.OpenSQLRecord(strSql, "�õ��ȴ���������", str��ǰ������, lngҵ������)
    
    If rsWorkID.RecordCount < 1 Then
        zlQueueOrderCall = 0
        Exit Function
    End If
    
    '���ҵ�ǰ�������ŶӺ���
    strSql = "select min(�ŶӺ���) as �ŶӺ��� from �ŶӽкŶ��� where ��������=[1] and ҵ������=[2] and �Ŷ�״̬=0 "
    Set rsQueueNo = zlDatabase.OpenSQLRecord(strSql, "�õ���ǰ���˵��ŶӺ���", str��ǰ������, lngҵ������)
    
     Call mfrmQueueStation.zlQueueExec(str��ǰ������, lngҵ������, rsWorkID!ҵ��ID, byt��������)
    
    zlQueueOrderCall = Nvl(rsQueueNo!�ŶӺ���, 0)

    Exit Function
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function zlQueueBroadcastCall(ByVal str�����ı� As String) As Boolean
'*************************************************************************************
'�Ժ����ı����й㲥����

'str�������ƣ���Σ��ַ����͡���Ҫ�����Ķ������ơ�
'lngҵ��ID����Σ��ַ����͡�����ŶӼ�¼��ҵ�����ݵĹ�ϵ��������������д����ţ�PACS����дҽ��ID��

'�޸ģ��޹���
'���ڣ�2012-6-27
'˵����������гɹ�����True ʧ�ܷ���false
'*************************************************************************************
    
    '���ù㲥���к���
    zlQueueBroadcastCall = mfrmQueueStation.QueueBroadcastCall(str�����ı�)

End Function


Public Function zlDelQueue(ByVal str�������� As String, Optional ByVal strҵ��ID As String = "") As Boolean
'*************************************************************************************
'�Ӷ�����ɾ���Ŷ���Ϣ

'str�������ƣ���Σ��ַ����͡���Ҫ�����Ķ������ơ�
'lngҵ��ID����Σ��ַ����͡�����ŶӼ�¼��ҵ�����ݵĹ�ϵ��������������д����ţ�PACS����дҽ��ID��

'�޸ģ��޹���
'���ڣ�2012-6-26
'˵�������û�в��ҵ����ݣ���ô�ӿڷ��ص��ŶӺ���Ϊ0
'*************************************************************************************
    
    Dim strSql As String
    
    On Error GoTo err
    
    strSql = "ZL_�ŶӽкŶ���_DELETE('" & str�������� & "','" & strҵ��ID & "')"
    zlDatabase.ExecuteProcedure strSql, "ɾ���Ŷ���Ϣ"
    zlDelQueue = True
    
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function






Public Function zlUpdateQueue(ByVal str�������� As String, ByVal strҵ��ID As String, ByVal lng����ID As Long, _
        ByVal str�������� As String, ByVal str���� As String, ByVal strҽ������ As String) As Boolean
'*************************************************************************************
'
'���¶����е���Ϣ
'
'str�������ƣ�
'
'lngҵ��ID��
'
'lng����ID��
'
'str����������
'
'str���ң�
'
'strҽ��������
'
'*************************************************************************************
    
    Dim strSql As String
    
    On Error GoTo err
  
    strSql = "ZL_�ŶӽкŶ���_UPDATE('" & str�������� & "','" & strҵ��ID & "'," & lng����ID & ",'" _
            & str�������� & "','" & str���� & "','" & strҽ������ & "')"
    zlDatabase.ExecuteProcedure strSql, "�޸��Ŷ���Ϣ"
    zlUpdateQueue = True
    Exit Function
err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function GetQueueBusinessDataIDs() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡҵ��IDs
    '���:bytType-0-�Һ�;1...
    '����:
    '����:�ɹ�����ҵ��IDs,����ö��ŷ���,��:22,33,44
    '����:���˺�
    '����:2014-03-11 16:48:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
 

    GetQueueBusinessDataIDs = mfrmQueueStation.GetQueueBusinessDataIDs()

End Function


Private Sub mfrmQueueStation_OnQueueExecuteAfter(ByVal strҵ��ID As String, ByVal byt�������� As Byte)
    RaiseEvent OnQueueExecuteAfter(strҵ��ID, byt��������)
End Sub

Private Sub mfrmQueueStation_OnQueueExecuteBefore(ByVal strҵ��ID As String, ByVal byt�������� As Byte, blnCancel As Boolean, strNewQueueName As String)
    RaiseEvent OnQueueExecuteBefore(strҵ��ID, byt��������, blnCancel, strNewQueueName)
End Sub

Private Sub mfrmQueueStation_OnRecevieDiagnose(ByVal strҵ��ID As String, ByVal lngҵ������ As Long)
    RaiseEvent OnRecevieDiagnose(strҵ��ID, lngҵ������)
End Sub

Private Sub mfrmQueueStation_OnSelectionChanged(ByVal blnIsCallingList As Boolean, objDataRow As XtremeReportControl.ReportRow, cbrMain As XtremeCommandBars.CommandBars)
    RaiseEvent OnSelectionChanged(blnIsCallingList, objDataRow, cbrMain)
End Sub