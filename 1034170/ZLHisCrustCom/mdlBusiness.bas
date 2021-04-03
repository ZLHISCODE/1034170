Attribute VB_Name = "mdlBusiness"
Option Explicit
'���øó���Ҫʵ�ֵĲ���
Public Enum OperateType
    OT_Repair = 0                                       '�����޸����൱������,���ж��Ƿ�Ԥ�������
    OT_PreUpgrade = 1                                   '��ǰ�������������ļ�������ʱĿ¼
    OT_OfficialUpgrade = 2                              '����ǰ����Ŀ¼�л��߷�����Ŀ¼����ȡ�ļ�����װ·��
    OT_CheckFile = 3                                        '��ʱֻ���ļ��ռ����ռ�APPSOFTĿ¼�µ�ָ�������ļ����������������������ۣ��������͵���Ϊ���ͻ��˲����Ƿ���Ҫ����
End Enum

Public Enum OperateStep
    OS_NotInProcessing = 0                              'δִ��
    OS_Completed = 1                                    'ִ�����,����OT_CheckFile,Ϊ�����ϣ���������
    OS_Failure = 2                                      'ִ��ʧ��,����OT_CheckFile,Ϊ�����ϣ�������
    OS_InProcessing = 3                                 'ִ����
End Enum

'��������
Public Enum MsgType
    MT_MsgHeader = 0                                    '��Ϣͷ
    MT_InitEnv = 1                                      '�ô�������δ��ʶ
    MT_SvrConn = 2                                      '���ӷ���������
    MT_ChcekUpdate = 3                                  '���¼��
    MT_DownAndDec = 4                                   '���ؽ�ѹ��������
    MT_SetUp = 5                                        '���������ڰ�װĿ¼����
    MT_RegCom = 6                                       '����ע�����
    MT_ExeBat = 7                                       'ִ�����������
    MT_MsgFoot = 8                                      '��Ϣβ��
End Enum

'�ļ�����
Public Enum FileType
    FT_Public = 0                   '��Ʒ��������
    FT_Apply = 1                    '��ƷӦ�ò���
    FT_Help = 2                     '��Ʒ�����ļ�
    FT_AdditionFile = 3             '��Ʒ�����ļ�
    FT_Other = 4                    '������Ʒ�ļ�
    FT_System = 5                   'ϵͳ�ļ�
End Enum
Public Function SetOperateProcess(ByVal otCurType As OperateType, ByVal osCurStep As OperateStep, Optional ByVal strMsg As String, Optional ByVal lngBeach As Long) As Boolean
'���ܣ����²������ȡ�
'������otCurType=��ǰ��������
'      osCurStep=��ǰ����
'      lngBeach=����������
'      strMsg=������Ϣ
'���أ��Ƿ�ִ�гɹ�
    Dim blnComplete As Boolean, strSql As String
    Dim strBeach As String
    
    gobjTrace.WriteSection "�����������", SL_LevelThree
    strMsg = MidB(strMsg, 1, glngNoteLength - 30)
    On Error Resume Next
    strSql = "zlTOOLS.Zl_Zlclients_UpdateProcess('" & gstrComputerName & "'," & otCurType & "," & osCurStep & "," & SQLAdjust(strMsg) & "," & IIf(lngBeach <> 0 And osCurStep = OS_Completed, lngBeach, "Null") & ")"
    Call ExecuteProcedure(strSql, "SetOperateProcess")
    If Err.Number <> 0 Then
        gobjTrace.WriteInfo "SetOperateProcess", "��ǽ��", "����", "���SQL", Replace(Replace(strSql, Chr(10), ""), Chr(13), ""), "������Ϣ", Err.Description
        Err.Clear
        blnComplete = osCurStep = OS_Completed Or osCurStep = OS_Failure And otCurType = OT_CheckFile
        Select Case otCurType
            Case OT_OfficialUpgrade '��ʽ������������Ԥ���������Ϣ�������޸������Ϣ����ȡ��Ԥ������־��������־
                strSql = "Update zlTOOLS.zlClients Set �������=" & osCurStep & " ,����˵��=" & SQLAdjust(strMsg) & "" & IIf(lngBeach <> 0 And osCurStep = OS_Completed, ",����=" & lngBeach, "") & IIf(blnComplete, ",������־=0,�Ƿ�Ԥ����=0,�޸�״̬=0,Ԥ�����=0,�Ƿ���������=0,�ռ���־=NULL,�ռ�״̬=NULL", "") & " Where ����վ = '" & gstrComputerName & "'"
            Case OT_PreUpgrade
                strSql = "Update zlTOOLS.zlClients Set Ԥ�����=" & osCurStep & " ,Ԥ����˵��=" & SQLAdjust(strMsg) & "" & IIf(blnComplete, ",�Ƿ�Ԥ����=0,�ռ���־=NULL,�ռ�״̬=NULL", "") & " Where ����վ = '" & gstrComputerName & "'"
            Case OT_Repair '�����޸���������Ԥ���������Ϣ�������޸������Ϣ����ȡ��Ԥ������־��������־
                strSql = "Update zlTOOLS.zlClients Set �޸�״̬=" & osCurStep & " ,�޸�˵��=" & SQLAdjust(strMsg) & "" & IIf(lngBeach <> 0 And osCurStep = OS_Completed, ",����=" & lngBeach, "") & IIf(blnComplete, ",������־=0,�Ƿ�Ԥ����=0,�������=0,Ԥ�����=0,�Ƿ���������=0,�ռ���־=NULL,�ռ�״̬=NULL", "") & " Where ����վ = '" & gstrComputerName & "'"
            Case OT_CheckFile
                strSql = "Update zlTOOLS.zlClients Set �ռ�״̬=" & osCurStep & " ,�ռ�˵��=" & SQLAdjust(strMsg) & "" & IIf(blnComplete, ",�ռ���־=0", "") & " Where ����վ = '" & gstrComputerName & "'"
                
        End Select
        gcnOracle.Execute strSql, , adCmdText
        If Err.Number <> 0 Then 'ִ��SQL����˵���ṹ��û������������ִ���Ͻṹ����
            gobjTrace.WriteInfo "SetOperateProcess", "��ǽ��", "����", "���SQL", Replace(Replace(strSql, Chr(10), ""), Chr(13), ""), "������Ϣ", Err.Description
            Err.Clear
            Select Case otCurType
                Case OT_OfficialUpgrade '��ʽ������������Ԥ���������Ϣ�������޸������Ϣ����ȡ��Ԥ������־��������־
                    strSql = "Update zlTOOLS.zlClients Set �������=" & osCurStep & " ,˵��=" & SQLAdjust(strMsg) & "" & IIf(blnComplete, ",������־=0,Ԥ�����=0,�Ƿ���������=0,�ռ���־=NULL,�ռ�״̬=NULL", "") & " Where ����վ = '" & gstrComputerName & "'"
                Case OT_PreUpgrade
                    strSql = "Update zlTOOLS.zlClients Set Ԥ�����=" & osCurStep & " ,˵��=" & SQLAdjust(strMsg) & "" & IIf(blnComplete, ",������־=0,�ռ���־=NULL,�ռ�״̬=NULL", "") & " Where ����վ = '" & gstrComputerName & "'"
                Case OT_Repair '�����޸���������Ԥ���������Ϣ�������޸������Ϣ����ȡ��Ԥ������־��������־
                    strSql = "Update zlTOOLS.zlClients Set �������=" & osCurStep & " ,˵��=" & SQLAdjust(strMsg) & "" & IIf(blnComplete, ",������־=0,Ԥ�����=0,�Ƿ���������=0,�ռ���־=NULL,�ռ�״̬=NULL", "") & " Where ����վ = '" & gstrComputerName & "'"
                Case OT_CheckFile
                    strSql = "Update zlTOOLS.zlClients Set �ռ�״̬=" & osCurStep & ",˵��=" & SQLAdjust(strMsg) & "" & IIf(blnComplete, ",�ռ���־=0", "") & " Where ����վ = '" & gstrComputerName & "'"
            End Select
            gcnOracle.Execute strSql, , adCmdText
            If Err.Number <> 0 Then
                gobjTrace.WriteInfo "SetOperateProcess", "��ǽ��", "ʧ��", "���SQL", Replace(Replace(strSql, Chr(10), ""), Chr(13), ""), "������Ϣ", Err.Description
                Call RecordErrMsg(MT_InitEnv, "�������ִ�����", "��ȷ�Ϲ����߶�����Ȩ��������" & Err.Description)
                Call RecordErrMsg(MT_MsgFoot, "��Ϣβ", "���:����ʧ�� ʱ��:" & Format(Currentdate, "yyyy-MM-dd HH:mm:ss"))
                Err.Clear
                MsgBox "�޷��������ִ�����������ϵ����Աȷ�Ϲ����߶���Ȩ��������", vbInformation, App.Title
                Exit Function
            ElseIf osCurStep = OS_InProcessing And otCurType = OT_CheckFile Then
                strSql = "Delete Zlclientupdatelog A Where a.����վ ='" & gstrComputerName & "' And ���� = 1"
                gcnOracle.Execute strSql, , adCmdText
                If Err.Number <> 0 Then Err.Clear
            End If
        ElseIf osCurStep = OS_InProcessing And otCurType = OT_CheckFile Then
            strSql = "Delete Zlclientupdatelog A Where a.����վ ='" & gstrComputerName & "' And ���� = 1"
            gcnOracle.Execute strSql, , adCmdText
            If Err.Number <> 0 Then Err.Clear
        End If
    End If
    gobjTrace.WriteInfo "SetOperateProcess", "��ǽ��", "�ɹ�", "���SQL", Replace(Replace(strSql, Chr(10), ""), Chr(13), "")
    SetOperateProcess = True
End Function

Public Function CheckJobs() As Boolean
'����:��鲢��ȡ�������������
    Dim rsTmp       As adodb.Recordset, strSql  As String
    Dim datCur      As Date, blnOnlyOfficialUp  As Boolean, blnOnlyPreUp    As Boolean
    Dim blnPreUp    As Boolean, blnOfficialUp   As Boolean, blnPreComplete  As Boolean, blnCollect  As Boolean
    Dim strMsg      As String
    
    On Error GoTo errH
    '���´���һ�㲻���ܳ���
    datCur = Currentdate
    '�ж������Ƿ������ȡ�Ƿ������˶�ʱ����
    strSql = "Select Max(����) ���� From ZLTOOLS.zlRegInfo Where ��Ŀ='�ͻ�����������'"
    Set rsTmp = OpenSQLRecord(strSql, "��鶨ʱ����")
    If rsTmp!���� & "" <> "" Then
        If CDate(Format(datCur, "YYYY-MM-DD hh:mm:ss")) >= CDate(Format(NVL(rsTmp!����), "YYYY-MM-DD hh:mm:ss")) Then
            blnOnlyOfficialUp = True 'ֻ����ʽ����
        Else
            blnOnlyPreUp = True 'ֻ��Ԥ����
        End If
    Else
        blnOnlyOfficialUp = True
    End If
    gobjTrace.WriteInfo "CheckJobs", "�Ƿ�ֻ����ʽ����", blnOnlyOfficialUp, "�Ƿ�ֻ��Ԥ����", blnOnlyPreUp
    On Error Resume Next
    Set rsTmp = Nothing
    '����û���Ƿ�Ԥ�����ֶ�(��ΪԤ����ʱ�����ݿ⻹û�������������Ҫ�������
    strSql = "Select Nvl(�Ƿ�Ԥ����,0) �Ƿ�Ԥ����, Nvl(Ԥ�����, 0) Ԥ�����, Nvl(������־, 0) ������־, Nvl(�ռ���־, 0) �ռ���־ From ZLTOOLS.Zlclients Where ����վ = [1]"
    Set rsTmp = OpenSQLRecord(strSql, "��鵱ǰ����", gstrComputerName)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo errH
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            blnPreUp = rsTmp!�Ƿ�Ԥ���� = 1
            blnOfficialUp = rsTmp!������־ = 1
            blnPreComplete = rsTmp!Ԥ����� = 1
            blnCollect = rsTmp!�ռ���־ = 1
        End If
    Else
        '�����·�ʽ��ȡ��ʧ����ʹ���Ϸ�ʽ�����Ӽ�����
        strSql = "Select Nvl(Ԥ�����, 0) Ԥ�����, Nvl(������־, 0) ������־, Nvl(�ռ���־, 0) �ռ���־ From ZLTOOLS.Zlclients Where ����վ = [1]"
        Set rsTmp = OpenSQLRecord(strSql, "��鵱ǰ����", gstrComputerName)
        If Not rsTmp.EOF Then
            blnPreUp = rsTmp!������־ = 1
            blnOfficialUp = rsTmp!������־ = 1
            blnPreComplete = rsTmp!Ԥ����� = 1
            blnCollect = rsTmp!�ռ���־ = 1
        End If
    End If
    gobjTrace.WriteInfo "CheckJobs", "�Ƿ���ҪԤ����", blnPreUp, "�Ƿ���Ҫ��ʽ����", blnOnlyPreUp, "�ϴ�Ԥ�����Ƿ����", blnPreComplete, "�Ƿ�����ļ��ռ�", blnCollect
    If gotCurType = OT_Repair Then
        If blnOnlyPreUp Then
            gotCurType = OT_PreUpgrade
        End If
    ElseIf (blnOfficialUp Or blnPreUp) And blnOnlyPreUp Then
        gotCurType = OT_PreUpgrade
    ElseIf (blnOfficialUp Or blnPreUp) And blnOnlyOfficialUp Then
        gotCurType = OT_OfficialUpgrade
    ElseIf blnCollect Then
        gotCurType = OT_CheckFile
    Else
        gobjTrace.WriteInfo "CheckJobs", "�����", "��ǰû���κ�����ϵͳ���Զ��˳�"
        Call RecordErrMsg(MT_InitEnv, "������", "��ǰû���κ�����ϵͳ���Զ��˳�")
        CheckJobs = True
        Exit Function
    End If
    'Ԥ�����Ѿ����
    If blnPreComplete And gotCurType = OT_PreUpgrade Then
        gobjTrace.WriteInfo "CheckJobs", "�����", "��ǰֻ��Ԥ����������Ԥ�����Ѿ���ɣ�ϵͳ���Զ��˳���"
        Call RecordErrMsg(MT_InitEnv, "������", "��ǰֻ��Ԥ����������Ԥ�����Ѿ���ɣ�ϵͳ���Զ��˳���")
        CheckJobs = True
        Exit Function
    End If
    gblnSilence = gotCurType = OT_CheckFile Or gotCurType = OT_PreUpgrade
    gobjTrace.WriteInfo "CheckJobs", "�����", Decode(gotCurType, OT_OfficialUpgrade, "��ʽ����", OT_PreUpgrade, "Ԥ����", OT_Repair, "�޸���ǿ������", OT_CheckFile, "�ռ�������")
    If gotCurType <> OT_CheckFile Then
        Set gclsConnect = GetFileConnect(strMsg)
        If gclsConnect Is Nothing Then
            gobjTrace.WriteInfo "CheckJobs", "����ʧ��", strMsg
            Call RecordErrMsg(MT_InitEnv, "������", "�޷������ļ�������,���ܼ������в�������Ϣ��" & strMsg)
            MsgBox "�޷������ļ�������������ϵ����Ա����Ϣ��" & vbNewLine & strMsg, vbInformation, App.Title
            Exit Function
        End If
    Else
        Set gclsConnect = New clsConnect
    End If
    CheckJobs = True
    Exit Function
errH:
    strMsg = Err.Description
    gobjTrace.WriteInfo "CheckJobs", "�����ⷢ����������", strMsg
    MsgBox "�����ⷢ��������������ϵ����Ա����Ϣ��" & vbNewLine & strMsg, vbInformation, App.Title
    Err.Clear
End Function

Private Function GetFileConnect(ByRef strMsg As String) As clsConnect
'���ܣ���ȡ�������ļ�����
    Dim objConn As New clsConnect
    Dim sctConnType As ServerConnectType
    Dim strServerID As String, strServer As String, strUser As String, strPwd As String, strPort As String, strCollectType As String
    Dim rsTmp As adodb.Recordset, strSql As String
    Dim blnDefalut As Boolean, blnConnOK As Boolean
    Dim blnOldStype As Boolean
    On Error Resume Next
    If gotCurType = OT_CheckFile Then
        strSql = "Select ����, λ��, �û���, ����, �˿�, �ռ����� From Zltools.Zlupgradeserver Where Nvl(�Ƿ��ռ�, 0) = 1"
        Set rsTmp = OpenSQLRecord(strSql, "��ȡ�������������", gstrComputerName)
        If Err.Number = 0 Then
            If Not rsTmp.EOF Then
                strServerID = rsTmp!��� & ""
                sctConnType = IIf(rsTmp!���� = 0, SCT_Share, SCT_FTP)
                strServer = rsTmp!λ��
                strUser = rsTmp!�û���
                strPwd = DeCipher(rsTmp!���� & "")
                strPort = rsTmp!�˿� & ""
                strCollectType = rsTmp!�ռ����� & ""
            End If
        Else
            Err.Clear
            blnOldStype = True
        End If
    Else
        strSql = "Select �����ļ������� From ZLTools.zlClients Where ����վ=[1]"
        Set rsTmp = OpenSQLRecord(strSql, "��ȡ�������������", gstrComputerName)
        If Err.Number = 0 Then
            If Not rsTmp.EOF Then strServerID = rsTmp!�����ļ������� & ""
        Else
            Err.Clear
            blnOldStype = True
        End If
        If strServerID <> "" Then
            strSql = "Select ���,����, λ��, �û���, ����, �˿�,Nvl(�Ƿ�ȱʡ,0) �Ƿ�ȱʡ , ���� From Zltools.Zlupgradeserver Where ��� = [1]"
            Set rsTmp = OpenSQLRecord(strSql, "��ȡ����������", Val(strServerID))
            If Not rsTmp.EOF Then
                strServerID = rsTmp!��� & ""
                sctConnType = IIf(rsTmp!���� = 0, SCT_Share, SCT_FTP)
                strServer = rsTmp!λ��
                strUser = rsTmp!�û���
                strPwd = DeCipher(rsTmp!���� & "")
                strPort = rsTmp!�˿� & ""
                glngFileBatch = Val(rsTmp!���� & "")
                blnDefalut = rsTmp!�Ƿ�ȱʡ = 1
            Else
                strServerID = ""
            End If
        End If
    End If
    If blnOldStype Then
        Set GetFileConnect = GetFileConnectOld(strMsg)
    Else
        If strServerID <> "" Then
            gobjTrace.WriteInfo "�ļ�������", "�������ļ�����", glngFileBatch, "���������", strServerID, "�Ƿ�Ĭ��", blnDefalut
            blnConnOK = objConn.ToConnect(sctConnType, strServer, strUser, strPwd, strPort, strCollectType, strMsg)
        End If
        '���Ӳ��ɹ��������������Զ�����Ĭ�Ϸ�����
        If Not blnConnOK And gotCurType <> OT_CheckFile And Not blnDefalut Then
            strSql = "Select ���,����, λ��, �û���, ����, �˿�, ���� From Zltools.Zlupgradeserver Where Nvl(�Ƿ�ȱʡ,0) = 1"
            Set rsTmp = OpenSQLRecord(strSql, "��ȡĬ������������")
            If Err.Number = 0 Then
                If Not rsTmp.EOF Then
                    strServerID = rsTmp!��� & ""
                    sctConnType = IIf(rsTmp!���� = 0, SCT_Share, SCT_FTP)
                    strServer = rsTmp!λ��
                    strUser = rsTmp!�û���
                    strPwd = DeCipher(rsTmp!���� & "")
                    strPort = rsTmp!�˿� & ""
                    glngFileBatch = Val(rsTmp!���� & "")
                    gobjTrace.WriteInfo "Ĭ�Ϸ�����", "�������ļ�����", glngFileBatch, "���������", strServerID
                    blnConnOK = objConn.ToConnect(sctConnType, strServer, strUser, strPwd, strPort, , strMsg)
                End If
            Else
                Err.Clear
            End If
        End If
        If blnConnOK Then Set GetFileConnect = objConn
    End If
    Exit Function
errH:
    strMsg = Err.Description
End Function

Private Function GetFileConnectOld(ByRef strMsg As String) As clsConnect
'���ܣ���ȡ�ļ����������ӣ��Ϸ�ʽ
'������blnUpgrade=True-Ԥ���������������� ��false-�ļ��ռ�������
    Dim rsTmp As adodb.Recordset, strSql As String
    Dim sctConnType As ServerConnectType, strServerID As String
    Dim objConn As New clsConnect
    Dim arrParas() As Variant, arrValues(4) As String
    Dim strSQLPars As String, i As Integer
    Dim blnReadOk As Boolean, blnConnOK As Boolean, blnGo As Boolean
    
    On Error GoTo errH
    '��ȡ��������
    sctConnType = SCT_Share
    strSql = "Select ��Ŀ,���� From ZLTools.zlregInfo where ��Ŀ=[1]"
    Set rsTmp = OpenSQLRecord(strSql, "��������", IIf(gotCurType <> OT_CheckFile, "��������", "�ռ���ʽ"))
    If Not rsTmp.EOF Then
        If NVL(rsTmp!����, 0) = 1 Then sctConnType = SCT_FTP
    End If
    If gotCurType = OT_CheckFile Then
        '�ļ��ռ�������ID
        strServerID = IIf(sctConnType = SCT_FTP, "F", "S")
    Else
        '��ȡ������ID
        strSql = "Select ����������,FTP������ From ZLTools.zlClients Where ����վ=[1]"
        Set rsTmp = OpenSQLRecord(strSql, "��ȡ�������������", gstrComputerName)
        If Not rsTmp.EOF Then strServerID = IIf(sctConnType = SCT_FTP, rsTmp!FTP������ & "", rsTmp!���������� & "")
    End If
    '��ȡ��������Ϣ
    If gotCurType <> OT_CheckFile Then
        If sctConnType = SCT_FTP Then
            arrParas = Array("FTP������", "FTP�û�", "FTP����", "FTP�˿�", "")
        Else
            arrParas = Array("������Ŀ¼", "�����û�", "��������", "", "")
        End If
    Else
        arrParas = Array("�ռ�Ŀ¼", "�����û�", "��������", "���ʶ˿�", "�ռ�����")
    End If
ReGetParas:
    '�Ȼ�ȡSQL����
    strSQLPars = ""
    For i = LBound(arrParas) To UBound(arrParas)
        If arrParas(i) <> "" Then
            strSQLPars = strSQLPars & ",'" & arrParas(i) & IIf(i <> UBound(arrParas), strServerID, "") & "'"
        End If
    Next
    strSQLPars = Mid(strSQLPars, 2)
    strSql = "Select ��Ŀ,���� From ZLTools.zlregInfo where ��Ŀ in(" & strSQLPars & ")"
    Set rsTmp = OpenSQLRecord(strSql, "��ȡ������")
    If Not rsTmp.EOF Then
        For i = LBound(arrParas) To UBound(arrParas)
            If arrParas(i) <> "" Then
                rsTmp.Filter = "��Ŀ='" & arrParas(i) & IIf(i <> UBound(arrParas), strServerID, "") & "'"
                If Not rsTmp.EOF Then arrValues(i) = rsTmp!���� & ""
            End If
        Next
    End If
    
    blnReadOk = True
    '���������û�������Ϊ�գ����ܽ����ռ�������
    If arrValues(0) = "" Or arrValues(1) = "" Or arrValues(2) = "" Then
        blnReadOk = False
    'FTP��ʽ��Ҫһ���˿�
    ElseIf sctConnType = SCT_FTP And arrValues(3) = "" Then
        blnReadOk = False
    '�ռ�ʱ���ռ����Ͳ���Ϊ��
    ElseIf gotCurType = OT_CheckFile And arrValues(4) = "" Then
        blnReadOk = False
    End If
    If blnReadOk Then
        gobjTrace.WriteInfo "GetFileConnectOld", "�ɷ�ʽ���������", strServerID
        blnConnOK = objConn.ToConnect(sctConnType, arrValues(0), arrValues(1), arrValues(2), arrValues(3), arrValues(4), strMsg)
    End If
    If (Not blnConnOK Or Not blnReadOk) And gotCurType <> OT_CheckFile Then
        If strServerID <> "" And strServerID <> "0" Then
            strServerID = "0"
            GoTo ReGetParas '���»�ȡ���ӷ������Ĳ���
        ElseIf (strServerID = "0" Or strServerID = "") And Not blnGo Then
            blnGo = True '��ֹѭ��
            strServerID = IIf(strServerID = "0", "", "0")
            GoTo ReGetParas '���»�ȡ���ӷ������Ĳ���
        End If
    End If
    If blnConnOK Then Set GetFileConnectOld = objConn
    Exit Function
errH:
    strMsg = Err.Description
End Function

Public Function CheckAndAdjustFolder() As Boolean
'���ܣ����а�װ·�����޸�
    Dim strSql              As String, rsTmp        As adodb.Recordset
    Dim strPath             As String, arrTmp       As Variant
    Dim i                   As Integer
    Dim strErrInfo          As String
    
    Err.Clear: On Error GoTo errH
    strSql = "Select Distinct Upper(��װ·��) ��װ·�� From Zlfilesupgrade"
    Set rsTmp = OpenSQLRecord(strSql, "��ȡ·���ļ���")
    
    Do While Not rsTmp.EOF
        arrTmp = Split(rsTmp!��װ·�� & "", "\")
        strPath = ""
        If UBound(arrTmp) <> -1 Then
            arrTmp(0) = Trim(arrTmp(0))
            If arrTmp(0) = "[APPSOFT]" Then
                strPath = gstrSetupPath
            ElseIf arrTmp(0) = "[PUBLIC]" Then
                If Not gobjFSO.FolderExists(gstrSetupPath & "\PUBLIC") Then
                    gobjFSO.CreateFolder (gstrSetupPath & "\PUBLIC")
                End If
                strPath = gstrSetupPath & "\PUBLIC"
            ElseIf arrTmp(0) = "[APPLY]" Then
                strPath = gstrSetupPath & "\APPLY"
            ElseIf arrTmp(0) = "[OS:]" Then 'ϵͳ��
                strPath = Left(gstrSystemPath, 2)
            ElseIf arrTmp(0) = "[APP:]" Then  '��ǰ��װ��
                strPath = Left(gstrSetupPath, 2)
            End If
            If strPath <> "" Then
                For i = 1 To UBound(arrTmp)
                    If arrTmp(i) <> "" Then
                        strPath = strPath & "\" & arrTmp(i)
                        If Not gobjFSO.FolderExists(strPath) Then
                            gobjFSO.CreateFolder (strPath)
                        End If
                    End If
                Next
                '���氲װ·�����Ż�ת���ٶȡ�
                gcllSetPath.Add strPath, "K_" & rsTmp!��װ·��
            End If
        End If
        rsTmp.MoveNext
    Loop
    '���������װ·�����Ż�ת���ٶȡ�
    On Error Resume Next
    gcllSetPath.Add gstrSetupPath, "K_[APPSOFT]"
    gcllSetPath.Add gstrSetupPath & "\PUBLIC", "K_[PUBLIC]"
    gcllSetPath.Add gstrSetupPath & "\APPLY", "K_[APPLY]"
    gcllSetPath.Add Left(gstrSystemPath, 2), "K_[OS:]"
    gcllSetPath.Add Left(gstrSetupPath, 2), "K_[APP:]"
    gcllSetPath.Add gstrSystemPath, "K_[SYSTEM]"
    gcllSetPath.Add gobjFSO.GetParentFolderName(gstrSystemPath) & "\Help", "K_[HELP]"
    gcllSetPath.Add gstrSetupPath & "\APPLY", "K_[APPSOFT]\APPLY"
    If Err.Number Then Err.Clear
    On Error Resume Next
    '���������ļ�·��
    strSql = "Select distinct upper(��װ·��) ��װ·�� From zlFilesExpired"
    Set rsTmp = OpenSQLRecord(strSql, "��ȡ·���ļ���")
    If Not rsTmp Is Nothing Then
        Err.Clear
        Do While Not rsTmp.EOF
            strPath = gcllSetPath("K_" & rsTmp!��װ·��)
            If Err.Number <> 0 Then
                Err.Clear
                arrTmp = Split(rsTmp!��װ·�� & "", "\")
                strPath = ""
                If UBound(arrTmp) <> -1 Then
                    arrTmp(0) = Trim(arrTmp(0))
                    If arrTmp(0) = "[APPSOFT]" Then
                        strPath = gstrSetupPath
                    ElseIf arrTmp(0) = "[PUBLIC]" Then
                        If Not gobjFSO.FolderExists(gstrSetupPath & "\PUBLIC") Then
                            gobjFSO.CreateFolder (gstrSetupPath & "\PUBLIC")
                        End If
                        strPath = gstrSetupPath & "\PUBLIC"
                    ElseIf arrTmp(0) = "[APPLY]" Then
                        strPath = gstrSetupPath & "\APPLY"
                    ElseIf arrTmp(0) = "[OS:]" Then 'ϵͳ��
                        strPath = Left(gstrSystemPath, 2)
                    ElseIf arrTmp(0) = "[APP:]" Then '��ǰ��װ��
                        strPath = Left(gstrSetupPath, 2)
                    End If
                    If strPath <> "" Then
                        For i = 1 To UBound(arrTmp)
                            If arrTmp(i) <> "" Then
                                strPath = strPath & "\" & arrTmp(i)
                                If Not gobjFSO.FolderExists(strPath) Then
                                    gobjFSO.CreateFolder (strPath)
                                End If
                            End If
                        Next
                        '���氲װ·�����Ż�ת���ٶȡ�
                        gcllSetPath.Add strPath, "K_" & rsTmp!��װ·��
                    End If
                End If
            End If
            rsTmp.MoveNext
        Loop
    End If
    If Err.Number Then Err.Clear
    CheckAndAdjustFolder = True
    Exit Function
errH:
    strErrInfo = Err.Description
    gobjTrace.WriteInfo "CheckAndAdjustFolder", "����޸���װĿ¼ʧ��", strErrInfo
    Call RecordErrMsg(MT_InitEnv, "�޸���װĿ¼", strErrInfo)
    MsgBox "����޸���װĿ¼����������������ϵ����Ա����Ϣ��" & vbNewLine & strErrInfo, vbInformation, App.Title
End Function

Public Function UpgradeBase(Optional ByVal blnUpgrade As Boolean = True) As Boolean
'���ܣ������Զ���������Ҫ�Ļ�������
    Dim strFile As String, blnAdmin As Boolean
    Dim strErr As String
    Dim strSql As String, rsTmp As adodb.Recordset
    Dim strReturn As String
    Dim strMsg As String
    Dim strCommand As String, strTmp As String
    Dim objText As TextStream, blnMust  As Boolean, blnErr  As Boolean
    
    If blnUpgrade Then
        gobjTrace.WriteSection "������������", SL_LevelTwo
        On Error GoTo errH
        strSql = "Select ���, �ļ���, Upper(�ļ���) ��׼�ļ���," & IIf(gblnHaveVersion, "�ļ��汾��", " ") & " �汾��, �޸�����, �ļ�����, ҵ�񲿼�, Upper(��װ·��) ��װ·��, Md5, �Զ�ע��, ǿ�Ƹ���" & vbNewLine & _
                "From ZLTOOLS.Zlfilesupgrade" & vbNewLine & _
                "Where Upper(�ļ���) In ('ZLRUNAS.EXE','ZLHISCRUST.EXE','ZLHISCRUSTCOM.DLL','7Z.EXE','7Z.DLL','AAMD532.DLL','GACUTIL.EXE','GACUTIL.EXE.CONFIG','ZL7Z.DLL')"
        Set rsTmp = OpenSQLRecord(strSql, App.Title)
        '1����������ZLRUNAS.EXE��ȡ����ԱȨ�ޣ��ɴ˿�������MD5���㲿��������ZlHISCrust������MD5
        On Error Resume Next
        strFile = gstrSetupPath & "\zlTestAdmin.txt"
        Call gobjFSO.CreateTextFile(strFile, True)
        Call gobjFSO.CopyFile(strFile, gstrSystemPath & "\zlTestAdmin.txt", True)
        If Err.Number = 75 Then
            blnAdmin = False
        ElseIf Dir(gstrSystemPath & "\zlTestAdmin.txt", vbNormal) <> "" Then
            blnAdmin = True
            Call gobjFSO.DeleteFile(gstrSystemPath & "\zlTestAdmin.txt", True)
        Else
            blnAdmin = False
        End If
        Call gobjFSO.DeleteFile(strFile, True)
        If Err.Number <> 0 Then Err.Clear
        gobjTrace.WriteInfo "UpgradeBase", "SystemĿ¼д��Ȩ��", blnAdmin
        If Not blnAdmin Then
            rsTmp.Filter = "��׼�ļ���='ZLRUNAS.EXE'"
            If Not rsTmp.EOF Then
                strFile = GetActualPath(rsTmp!��װ·��, Val(rsTmp!�ļ����� & ""), rsTmp!�ļ���)
                If Not gobjFSO.FileExists(strFile) Then
                    gobjTrace.WriteInfo "UpgradeBase", "���������ļ�", rsTmp!�ļ���
                    If gclsConnect.IsServerFileExists(rsTmp!��׼�ļ���) Then
                        If Not gclsConnect.DownloadFile(rsTmp!��׼�ļ���, gobjFSO.GetParentFolderName(strFile), strErr) Then
                            strMsg = "�������ļ��ļ�����ʧ��(ZLRUNAS.EXE(USERȨ��ִ�й���))" & strErr
                        Else
                            gobjTrace.WriteInfo "UpgradeBase", "���ذ�װ", strFile
                        End If
                    Else
                        strMsg = "�������ļ�ȱʧZLRUNAS.EXE(USERȨ��ִ�й���)"
                    End If
                End If
                If gobjFSO.FileExists(strFile) Then
                    '�ȱ��������У����´�����ʹ��
                    If gobjFSO.FileExists(gstrSetupPath & "\ZLRUNAS.ini") Then
                        gobjFSO.DeleteFile gstrSetupPath & "\ZLRUNAS.ini", True
                    End If
                    Set objText = gobjFSO.CreateTextFile(gstrSetupPath & "\ZLRUNAS.ini")
                    objText.WriteLine Cipher(gstrCommand)
                    objText.Close
                    Set objText = Nothing
                    strMsg = StartZLRunAs(strFile)
                End If
            Else
                strMsg = "������Ŀ¼(ZLfilesUpgrade)��ȱʧZLRUNAS.EXE(USERȨ��ִ�й���)"
            End If
            If strMsg <> "" Then
                gobjTrace.WriteInfo "UpgradeBase", "����Ա���й��߼��", strMsg
                Call RecordErrMsg(MT_InitEnv, "����Ա���й��߼��", strMsg)
                MsgBox strMsg & vbNewLine & "������ϵ����Ա��", vbInformation, App.Title
                Exit Function
            End If
        End If
        '2������AAMD532.dll�ò�������������MD5,��������ZLHISCrust.exe�������޷����ZLHISCrust.exe�Ƿ���Ҫ������
        strMsg = ""
        rsTmp.Filter = "��׼�ļ���='AAMD532.DLL'"
        If Not rsTmp.EOF Then
            strFile = GetActualPath(rsTmp!��װ·�� & "", Val(rsTmp!�ļ����� & ""), rsTmp!�ļ���)
            If Not gobjFSO.FileExists(strFile) Then
                If gclsConnect.IsServerFileExists(rsTmp!��׼�ļ���) Then
                    gobjTrace.WriteInfo "UpgradeBase", "���������ļ�", rsTmp!�ļ���
                    If Not gclsConnect.DownloadFile(rsTmp!��׼�ļ���, gobjFSO.GetParentFolderName(strFile), strErr) Then
                        strMsg = "�������ļ��ļ�����ʧ��AAMD532.DLL(MD5���㹤��)" & strErr
                    Else
                        gobjTrace.WriteInfo "UpgradeBase", "���ذ�װ", strFile
                    End If
                Else
                    strMsg = "�������ļ�ȱʧAAMD532.DLL(MD5���㹤��)"
                End If
            End If
        Else
            strMsg = "������Ŀ¼(ZLfilesUpgrade)��ȱʧAAMD532.DLL(MD5���㹤��)"
        End If
        
        If strMsg <> "" Then
            gobjTrace.WriteInfo "UpgradeBase", "MD5���㹤�߼��", strMsg
            Call RecordErrMsg(MT_InitEnv, "MD5���㹤�߼��", strMsg)
            MsgBox strMsg & vbNewLine & "������ϵ����Ա��", vbInformation, App.Title
            Exit Function
        End If
        strMsg = ""
        '3������ZLHISCrust.exe���ò������Խ��м��������
        If Val(GetSetting("ZLSOFT", "����ģ��\�Զ�����", "���ߵ���", "0")) = 0 Then
            If gintCallTimes = 0 Then '�ڶ��ε����������߽�������������ZLRUNAS���õ���һ��
                rsTmp.Filter = "��׼�ļ���='ZLHISCRUST.EXE'"
                If Not rsTmp.EOF Then
                    strFile = GetActualPath(rsTmp!��װ·�� & "", Val(rsTmp!�ļ����� & ""), rsTmp!�ļ���)
                    If IsFileUpgade(gstrAppPath & "\ZLHISCRUST.EXE", rsTmp!�汾�� & "", rsTmp!�޸����� & "", rsTmp!MD5 & "") Then
                        If gclsConnect.IsServerFileExists(rsTmp!��׼�ļ���) Then
                            gobjTrace.WriteInfo "UpgradeBase", "���������ļ�", rsTmp!�ļ���
                            If Not gclsConnect.DownloadFile(rsTmp!��׼�ļ���, gstrTempPath, strErr) Then
                                strMsg = "�������ļ��ļ�����ʧ��:ZLHISCRUST.EXE(�Զ�����������)" & strErr
                            Else
                                gobjTrace.WriteInfo "UpgradeBase", "���ذ�װ", gstrTempPath & "\" & rsTmp!�ļ���
                                '�ļ��ֱ���ϲ��������ļ��ƶ���APPSOft\APPLY��
                                strTmp = UCase(GetVersionInfo(gstrTempPath & "\" & rsTmp!�ļ���, FVN_ProductName))
                                If strTmp = "" Then strTmp = "ZLHISINSTALLUPDATE"
                                If strTmp <> "ZLHISINSTALLUPDATE" Then 'zlHisInstallUpdate
                                    gobjTrace.WriteInfo "UpgradeBase", "ZLHISCRUST.EXE�����ع����ϵͰ汾", True
                                    strFile = gstrSetupPath & "\Apply\" & rsTmp!�ļ���
                                    If gobjFSO.FileExists(strFile) Then
                                        If FileSystem.GetAttr(strFile) <> vbNormal Then
                                             Call FileSystem.SetAttr(strFile, vbNormal)
                                        End If
                                        Call gobjFSO.DeleteFile(strFile)
                                    End If
                                    gobjFSO.CopyFile gstrTempPath & "\" & rsTmp!�ļ���, strFile, False
                                    strCommand = GetHisUpdateCommand(True)
                                Else
                                    gobjTrace.WriteInfo "UpgradeBase", "ZLHISCRUST.EXE�����ع����ϵͰ汾", False
                                    strFile = gstrTempPath & "\" & rsTmp!�ļ���
                                    strCommand = GetHisUpdateCommand()
                                End If
                                '���غ���Ҫʹ���µ�ZLHISCRUST.EXE����������
                                On Error Resume Next
                                Call gobjTrace.CloseLog
                                If Shell(strFile & " " & strCommand, vbNormalFocus) <> 0 Then
                                    Call gclsConnect.CloseConnect
                                    Call gobjMe.ExitApp
                                Else
                                End If
                            End If
                        Else
                            strMsg = "�������ļ�ȱʧZLHISCRUST.EXE(�Զ�����������)"
                        End If
                    End If
                Else
                    strMsg = "������Ŀ¼(ZLfilesUpgrade)��ȱʧZLHISCRUST.EXE(�Զ�����������)"
                End If
            End If
        End If
        If strMsg <> "" Then
            gobjTrace.WriteInfo "UpgradeBase", "�Զ��������߼��", strMsg
            Call RecordErrMsg(MT_InitEnv, "�Զ��������߼��", strMsg)
            MsgBox strMsg & vbNewLine & "������ϵ����Ա��", vbInformation, App.Title
            Exit Function
        End If
        strMsg = ""
        '3.1 �Զ�����DLL
        rsTmp.Filter = "��׼�ļ���='ZLHISCRUSTCOM.DLL'"
        If Not rsTmp.EOF Then
            strFile = GetActualPath(rsTmp!��װ·�� & "", Val(rsTmp!�ļ����� & ""), rsTmp!�ļ���)
            If IsFileUpgade(strFile, rsTmp!�汾�� & "", rsTmp!�޸����� & "", rsTmp!MD5 & "") Then
                If gclsConnect.IsServerFileExists(rsTmp!��׼�ļ���) Then
                    gobjTrace.WriteInfo "UpgradeBase", "���������ļ�", rsTmp!�ļ���
                    If Not gclsConnect.DownloadFile(rsTmp!��׼�ļ���, gstrTempPath, strErr) Then
                        strMsg = "�������ļ��ļ�����ʧ��(ZLHISCRUSTCOM.DLL(�Զ�����ҵ������))" & strErr
                    Else
                        gobjTrace.WriteInfo "UpgradeBase", "����(�쳣)", gstrTempPath & "\" & rsTmp!�ļ���
                    End If
                Else
                    strMsg = "�������ļ�ȱʧZLHISCRUSTCOM.DLL(�Զ�����ҵ������)"
                End If
            End If
        Else
            strMsg = "������Ŀ¼(ZLfilesUpgrade)��ȱʧZLHISCRUSTCOM.DLL(�Զ�����ҵ������)"
        End If
        If strMsg <> "" Then
            gobjTrace.WriteInfo "UpgradeBase", "�Զ��������߼��", strMsg
            Call RecordErrMsg(MT_InitEnv, "�Զ��������߼��", strMsg)
            MsgBox strMsg & vbNewLine & "������ϵ����Ա��", vbInformation, App.Title
            Exit Function
        End If
        
        strMsg = ""
        '4������ѹ�����ߣ��Ա��������������Ľ�ѹ
        rsTmp.Filter = "��׼�ļ���='7Z.DLL'"
        If Not rsTmp.EOF Then
            strFile = GetActualPath(rsTmp!��װ·�� & "", Val(rsTmp!�ļ����� & ""), rsTmp!�ļ���)
            If Not gobjFSO.FileExists(strFile) Then
                If gclsConnect.IsServerFileExists(rsTmp!��׼�ļ���) Then
                    gobjTrace.WriteInfo "UpgradeBase", "���������ļ�", rsTmp!�ļ���
                    If Not gclsConnect.DownloadFile(rsTmp!��׼�ļ���, gobjFSO.GetParentFolderName(strFile), strErr) Then
                        strMsg = "�������ļ��ļ�����ʧ��(7Z.DLL(��ѹ������������))" & strErr
                    Else
                        gobjTrace.WriteInfo "UpgradeBase", "���ذ�װ", strFile
                    End If
                Else
                    strMsg = "�������ļ�ȱʧ7Z.DLL(��ѹ������������)"
                End If
            End If
        Else
            strMsg = "������Ŀ¼(ZLfilesUpgrade)��ȱʧ7Z.DLL(��ѹ������������)"
        End If
        If strMsg <> "" Then
            gobjTrace.WriteInfo "��ѹ���߼��", "��Ϣ", strMsg
            Call RecordErrMsg(MT_InitEnv, "�Զ��������߼��", strMsg)
            MsgBox strMsg & vbNewLine & "������ϵ����Ա��", vbInformation, App.Title
            Exit Function
        End If
        strMsg = ""
        '4������ѹ�����ߣ��Ա��������������Ľ�ѹ
        rsTmp.Filter = "��׼�ļ���='ZL7Z.DLL'"
        If Not rsTmp.EOF Then
            strFile = GetActualPath(rsTmp!��װ·�� & "", Val(rsTmp!�ļ����� & ""), rsTmp!�ļ���)
            If Not gobjFSO.FileExists(strFile) Then
                If gclsConnect.IsServerFileExists(rsTmp!��׼�ļ���) Then
                    gobjTrace.WriteInfo "UpgradeBase", "���������ļ�", rsTmp!�ļ���
                    If Not gclsConnect.DownloadFile(rsTmp!��׼�ļ���, gobjFSO.GetParentFolderName(strFile), strErr) Then
                        strMsg = "�������ļ��ļ�����ʧ��(ZL7Z.DLL(����ѹ������))" & strErr
                    Else
                        strMsg = ""
                        gobjTrace.WriteInfo "UpgradeBase", "���ذ�װ", strFile
                        If Not gclsRegCom.RegCom(strFile, strMsg, RFT_NormalReg) Then
                            gobjTrace.WriteInfo "UpgradeBase", "ZL7Zע��ʧ��", strMsg
                            Call RecordErrMsg(MT_InitEnv, "ZL7Zע��ʧ��", strMsg)
                        Else
                            gobjTrace.WriteInfo "UpgradeBase", "ZL7Zע��ɹ�", ""
                        End If
                        strMsg = ""
                    End If
                Else
                    strMsg = "�������ļ�ȱʧZL7Z.DLL(����ѹ������)"
                End If
            End If
        Else
            strMsg = "������Ŀ¼(ZLfilesUpgrade)��ȱʧZL7Z.DLL(����ѹ������)"
        End If
        If strMsg <> "" Then
            gobjTrace.WriteInfo "��ѹ���߼��", "��Ϣ", strMsg
            Call RecordErrMsg(MT_InitEnv, "�Զ��������߼��", strMsg)
            MsgBox strMsg & vbNewLine & "������ϵ����Ա��", vbInformation, App.Title
            Exit Function
        End If
    End If
    strMsg = ""
    rsTmp.Filter = "��׼�ļ���='7Z.EXE'"
    If Not rsTmp.EOF Then
        strFile = GetActualPath(rsTmp!��װ·�� & "", Val(rsTmp!�ļ����� & ""), rsTmp!�ļ���)
        gobj7zZip.Path7z = strFile
        If blnUpgrade Then '������������
            If Not gobjFSO.FileExists(strFile) Then
                If gclsConnect.IsServerFileExists(rsTmp!��׼�ļ���) Then
                    gobjTrace.WriteInfo "UpgradeBase", "���������ļ�", rsTmp!�ļ���
                    If Not gclsConnect.DownloadFile(rsTmp!��׼�ļ���, gobjFSO.GetParentFolderName(strFile), strErr) Then
                        strMsg = "�������ļ��ļ�����ʧ��(7Z.EXE(��ѹ����))" & strErr
                    Else
                        gobjTrace.WriteInfo "UpgradeBase", "���ذ�װ", strFile
                    End If
                Else
                    strMsg = "�������ļ�ȱʧ7Z.EXE(��ѹ����)"
                End If
            End If
        End If
    Else
        strMsg = "������Ŀ¼(ZLfilesUpgrade)��ȱʧ7Z.EXE(��ѹ����)"
    End If
    If strMsg <> "" Then
        gobjTrace.WriteInfo "UpgradeBase", "��ѹ���߼��", strMsg
        Call RecordErrMsg(MT_InitEnv, "�Զ��������߼��", strMsg)
        MsgBox strMsg & vbNewLine & "������ϵ����Ա��", vbInformation, App.Title
        Exit Function
    End If
    '5������
    strMsg = ""
    blnMust = IsMustGACUTIL(): blnErr = False
    rsTmp.Filter = "��׼�ļ���='GACUTIL.EXE'"
    If Not rsTmp.EOF Then
        strFile = GetActualPath(rsTmp!��װ·�� & "", Val(rsTmp!�ļ����� & ""), rsTmp!�ļ���)
        gclsRegCom.GACUPath = strFile
        If blnUpgrade Then '������������
            If Not gobjFSO.FileExists(strFile) Then
                If gclsConnect.IsServerFileExists(rsTmp!��׼�ļ���) Then
                    gobjTrace.WriteInfo "UpgradeBase", "���������ļ�", rsTmp!�ļ���
                    If Not gclsConnect.DownloadFile(rsTmp!��׼�ļ���, gobjFSO.GetParentFolderName(strFile), strErr) Then
                        strMsg = "�������ļ��ļ�����ʧ��(GACUTIL.EXE(ȫ�ֻ�����ӹ���))" & strErr
                        blnErr = True
                    Else
                        gobjTrace.WriteInfo "UpgradeBase", "���ذ�װ", strFile
                    End If
                Else
                    strMsg = "�������ļ�ȱʧGACUTIL.EXE(ȫ�ֻ�����ӹ���)"
                End If
            End If
        End If
    Else
        strMsg = "������Ŀ¼(ZLfilesUpgrade)��ȱʧGACUTIL.EXE(ȫ�ֻ�����ӹ���)"
    End If
    If strMsg <> "" Then
        gobjTrace.WriteInfo "UpgradeBase", "ȫ�ֻ�����ӹ��߼��", strMsg
        If blnMust Or blnErr Then
            Call RecordErrMsg(MT_InitEnv, "ȫ�ֻ�����ӹ��߼��", strMsg)
            MsgBox strMsg & vbNewLine & ",����ϵ����Ա��", vbInformation, App.Title
            Exit Function
        End If
    End If
    
    If blnUpgrade Then '������������
        strMsg = ""
        blnErr = False
        rsTmp.Filter = "��׼�ļ���='GACUTIL.EXE.CONFIG'"
        If Not rsTmp.EOF Then
            strFile = GetActualPath(rsTmp!��װ·�� & "", Val(rsTmp!�ļ����� & ""), rsTmp!�ļ���)
            If Not gobjFSO.FileExists(strFile) Then
                If gclsConnect.IsServerFileExists(rsTmp!��׼�ļ���) Then
                    gobjTrace.WriteInfo "UpgradeBase", "���������ļ�", rsTmp!�ļ���
                    If Not gclsConnect.DownloadFile(rsTmp!��׼�ļ���, gobjFSO.GetParentFolderName(strFile), strErr) Then
                        strMsg = "�������ļ��ļ�����ʧ��(GACUTIL.EXE.CONFIG(ȫ�ֻ�����ӹ��������ļ�))" & strErr
                        blnErr = True
                    Else
                        gobjTrace.WriteInfo "UpgradeBase", "���ذ�װ", strFile
                    End If
                Else
                    strMsg = "�������ļ�ȱʧGACUTIL.EXE.CONFIG(ȫ�ֻ�����ӹ��������ļ�)"
                End If
            End If
        Else
            strMsg = "������Ŀ¼(ZLfilesUpgrade)��ȱʧGACUTIL.EXE.CONFIG(ȫ�ֻ�����ӹ��������ļ�)"
        End If
        If strMsg <> "" Then
            gobjTrace.WriteInfo "UpgradeBase", "ȫ�ֻ�����ӹ��߼��", strMsg
            If blnMust Or blnErr Then
                Call RecordErrMsg(MT_InitEnv, "ȫ�ֻ�����ӹ��߼��", strMsg)
                MsgBox strMsg & vbNewLine & ",����ϵ����Ա��", vbInformation, App.Title
                Exit Function
            End If
        End If
    End If
    If Not gobj7zZip.Init7zZip Then
        gobjTrace.WriteInfo "UpgradeBase", "7zZip��ʼ��", "�޷�����ZL7z������û��7z.exe"
        Call RecordErrMsg(MT_InitEnv, "�Զ��������߼��", "�޷�����ZL7z������û��7z.exe")
        MsgBox "�޷�����ZL7z������û��7z.exe" & vbNewLine & "������ϵ����Ա��", vbInformation, App.Title
        Exit Function
    End If
    UpgradeBase = True
    Exit Function
errH:
    gobjTrace.WriteInfo "UpgradeBase", "������������������������", Err.Description
    Call RecordErrMsg(MT_InitEnv, "������������������������", Err.Description)
    MsgBox "������������������������" & vbNewLine & "������ϵ����Ա����Ϣ��" & Err.Description, vbInformation, App.Title
    Err.Clear
End Function

Private Function StartZLRunAs(ByVal strPath As String) As String
'���ܣ�����ZLRunas
    Dim strSql          As String, rsTmp    As adodb.Recordset
    Dim strUser         As String, strPwd   As String
    Dim strCommandPara  As String, strMsg   As String, strReturn As String
    Dim blnOk           As Boolean
    Dim objShell        As New clsShell
    
    On Error Resume Next
    strSql = "Select Max(����Ա�û�) ����Ա, Max(����Ա����)  ���� From ZLTOOLS.zlClients Where ����վ = [1]"
    Set rsTmp = OpenSQLRecord(strSql, "��ȡ��ǰ�ͻ��˵�¼���")
    '����ģʽ���Ͱ汾û���������ֶ�
    If Err.Number = 0 Then
        strUser = NVL(rsTmp!����Ա, "Administrator")
        strPwd = Trim(rsTmp!���� & "")
    Else
        Err.Clear
    End If
    On Error GoTo errH
    '�������
    If strPwd <> "" And strUser <> "" Then
        strPwd = DeCipher(strPwd)
        strCommandPara = "-u " & strUser & " -p " & strPwd  '����ZLRunas.EXE������
        gobjTrace.WriteInfo "StartZLRunAs", "�ͻ��˹������", Cipher(strCommandPara)
        '���������������
        If objShell.Run(strPath & " " & strCommandPara & " -ex """ & gstrAppPath & "\ZLHISCRUST.EXE"" -lwp", strReturn, , 30000) Then
            If InStr(strReturn, (1326)) > 0 Then
                strMsg = "��¼ʧ��: δ֪���û�����������롣"
            ElseIf InStr(strReturn, (1058)) > 0 Then
                strMsg = "�޷���������ԭ�������SecLogon���񱻽��á�"
            ElseIf InStr(strReturn, (1717)) > 0 Then
                strMsg = "'·���в��������ģ�����ִ�в��ɹ�"
            Else
                blnOk = True
            End If
        End If
    Else
        gobjTrace.WriteInfo "StartZLRunAs", "�ͻ��˹������", "û��ͳһ��������"
    End If
    'ʹ��ÿ���ͻ��˵ĸ�������
    If Not blnOk Then
        strSql = "Select Max(Decode(��Ŀ, '����Ա�˺�', ����, '')) As ����Ա, Max(Decode(��Ŀ, '����Ա����', ����, '')) As ����" & vbNewLine & _
                "From Zltools.Zlreginfo" & vbNewLine & _
                "Where ��Ŀ = '����Ա�˺�' Or ��Ŀ = '����Ա����'"
        Set rsTmp = OpenSQLRecord(strSql, "��ȡͳһ���")
        strUser = NVL(rsTmp!����Ա, "Administrator")
        strPwd = Trim(rsTmp!���� & "")
        If strPwd <> "" And strUser <> "" Then
            strPwd = DeCipher(strPwd)
            strCommandPara = "-u " & strUser & " -p " & strPwd  '����ZLRunas.EXE������
            gobjTrace.WriteInfo "StartZLRunAs", "��ǰ�ͻ��˵�¼���", Cipher(strCommandPara)
            '���������������
            If objShell.Run(strPath & " " & strCommandPara & " -ex """ & gstrAppPath & "\ZLHISCRUST.EXE"" -lwp", strReturn, , 30000) Then
                If InStr(strReturn, (1326)) > 0 Then
                    strMsg = "��¼ʧ��: δ֪���û�����������롣"
                ElseIf InStr(strReturn, (1058)) > 0 Then
                    strMsg = "�޷���������ԭ�������SecLogon���񱻽��á�"
                ElseIf InStr(strReturn, (1717)) > 0 Then
                    strMsg = "'·���в��������ģ�����ִ�в��ɹ�"
                Else
                    blnOk = True
                End If
            End If
        Else
            gobjTrace.WriteInfo "StartZLRunAs", "��ǰ�ͻ��˵�¼���", "û�е�¼�������"
        End If
    End If
    StartZLRunAs = strMsg
    Exit Function
errH:
    gobjTrace.WriteInfo "StartZLRunAs", "��ȡ�ͻ�����ɷ�����������", Err.Description
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function GetUpgradeFileList() As Boolean
'���ܣ���ȡZLFIleUpgrade
    Dim strSql As String, rsTmp As adodb.Recordset
    Dim strTmp As String, strMsg As String
    
    On Error GoTo errH
    '���ͬ���ļ�
    strSql = "Select Upper(a.�ļ���) �ļ��� From Zlfilesupgrade a Group By Upper(a.�ļ���) Having Count(1) > 1"
    Set rsTmp = OpenSQLRecord(strSql, "��ȡ�ļ��嵥")
    Do While Not rsTmp.EOF
        strTmp = strTmp & "," & rsTmp!�ļ���
        rsTmp.MoveNext
    Loop
    If strTmp <> "" Then
        strMsg = "����ͬ��(��Сд����)������" & Mid(Mid(strTmp, 2), 1, 100)
        gobjTrace.WriteInfo "GetUpgradeFileList", "�����嵥�Ϸ��Լ��", strMsg
        Call RecordErrMsg(MT_InitEnv, "�����嵥�Ϸ��Լ��", strMsg)
        MsgBox "�����嵥�������⣬����ϵ����Ա���д���" & vbNewLine & strMsg, vbInformation + vbDefaultButton1, App.Title
        Exit Function
    End If
    On Error Resume Next
    strSql = "Select a.�ļ���, Upper(a.�ļ���) ��׼�ļ���," & IIf(gblnHaveVersion, "a.�ļ��汾�� ", " a.") & "�汾��, a.�޸�����, a.�ļ�����, a.ҵ�񲿼�, a.��װ·��, a.Md5, NVL(a.�Զ�ע��,0) �Զ�ע��, NVL(a.ǿ�Ƹ���,0) ǿ�Ƹ���,���Ӱ�װ·��" & vbNewLine & _
            "From Zltools.Zlfilesupgrade a" & vbNewLine & _
            "Where Upper(a.�ļ���) Not In ('ZLRUNAS.EXE', 'ZLHISCRUST.EXE','ZLHISCRUSTCOM.DLL', '7Z.EXE', '7Z.DLL', 'AAMD532.DLL', 'GACUTIL.EXE','GACUTIL.EXE.CONFIG','ZL7Z.DLL')"
    Set rsTmp = OpenSQLRecord(strSql, "��ȡ�ļ��嵥")
    If Err.Number <> 0 Then
        Err.Clear
        strSql = "Select a.�ļ���, Upper(a.�ļ���) ��׼�ļ���, " & IIf(gblnHaveVersion, "a.�ļ��汾�� ", " a.") & "�汾��, a.�޸�����, a.�ļ�����, a.ҵ�񲿼�, a.��װ·��, a.Md5, NVL(a.�Զ�ע��,0) �Զ�ע��, NVL(a.ǿ�Ƹ���,0) ǿ�Ƹ���,Null ���Ӱ�װ·��" & vbNewLine & _
                "From Zltools.Zlfilesupgrade a" & vbNewLine & _
                "Where Upper(a.�ļ���) Not In ('ZLRUNAS.EXE', 'ZLHISCRUST.EXE','ZLHISCRUSTCOM.DLL', '7Z.EXE', '7Z.DLL', 'AAMD532.DLL', 'GACUTIL.EXE','GACUTIL.EXE.CONFIG','ZL7Z.DLL')"
        Set rsTmp = OpenSQLRecord(strSql, "��ȡ�ļ��嵥")
    End If
    'ʵ��·��-��װ·��ת��Ϊʵ��·��
    '�����ļ�·��-����·���ļ�
    Set grsFileUpgrade = CopyNewRec(rsTmp, , , Array("����", adInteger, 1, 0, "ʵ��·��", adVarChar, 500, Empty, "�����ļ�·��", adVarChar, 1000, Empty, "����ʵ��·��", adVarChar, 4000, Empty, _
                                                "�ж�����", adInteger, 3, 0, "Ԥ��������", adInteger, 1, 0, "������Ϣ", adVarChar, 1000, Empty, "�����Ϣ", adVarChar, 1000, Empty, _
                                                "�޺�׺�ļ���", adVarChar, 100, Empty, "��������", adInteger, 1, 0, "ע�����", adInteger, 1, 0))
    GetUpgradeFileList = True
    Exit Function
errH:
    gobjTrace.WriteInfo "GetUpgradeFileList", "�����嵥��ȡʧ��", Err.Description
    Call RecordErrMsg(MT_InitEnv, "�ļ��嵥��ȡ", Err.Description)
    MsgBox "�����嵥��ȡʧ�ܣ�" & vbNewLine & "����ϵ����Ա����Ϣ��" & Err.Description, vbInformation, App.Title
End Function

Public Function GetKILLProcess() As Boolean
'���ܣ���ȡҪɱ���Ľ���
    Dim strSql As String, rsTmp As adodb.Recordset
    Dim strTmp As String

    On Error Resume Next
    strSql = "Select ���, ����,���� From Zltools.ZlkillProcess Order By ���"
    Set rsTmp = OpenSQLRecord(strSql, "��ȡ�ļ��嵥")
    If rsTmp Is Nothing Then
        If Err.Number <> 0 Then Err.Clear
    Else
        Do While Not rsTmp.EOF
            strTmp = strTmp & ";" & UCase(rsTmp!����)
            rsTmp.MoveNext
        Loop
    End If
    
    If strTmp = "" Then
        strTmp = "zl9LabPrintSvr.exe;zl9LabReceiv.exe;zl9LabTcpSvr.exe;Zl9LISComm.exe;zl9PacsCapture.exe;zl9WizardMain.exe;zl9WizardStart.exe;ZL9Xls.exe;zlActMain.exe;ZLBAExport.exe;zlCDOpen.exe;zlCisAuditPrint.exe;zlDrugMachineManage.exe;zlGetImage.exe;zlGetImageEx.exe;zlHQMSDCollect.exe;zlLisReceiveSend.exe;zlMipClientManage.exe;zlMipClientPoll.exe;zlMipClientShell.exe;zlMsgBuilderStart.exe;zlMsgReceiver.exe;zlMsgSender.exe;ZLNewQuery.exe;zlOrclConfig.exe;ZLPacsBrowserStation.exe;ZlPacsSrv.exe;zlPeisAutoAnalyse.exe;zlQueueShow.exe;ZLRPTSQLAdjust.exe;ZLRUNAS.EXE;zlScreenKeyboard.exe;zlSoftShowArchive.exe;zlSvrNotice.exe;zlSvrStudio.exe;zlUpgradeReader.exe;zlWizardStart.exe;ZLPacsServerCenter.exe"
    Else
        strTmp = Mid(2, strTmp)
    End If
    gobjTrace.WriteInfo "GetKILLProcess", "�����嵥", strTmp
    garrKillProcess = Split(UCase(strTmp), ";")
    If Err.Number <> 0 Then Err.Clear
End Function

Public Function IsMustGACUTIL() As Boolean
'���ܣ��Ƿ����ҪGACUTIL.EXE��GACUTIL.EXE.CONFIG
    Dim strSql As String, rsTmp As adodb.Recordset

    On Error GoTo errH
    strSql = "Select Count(1) ���� From Zlfilesupgrade a Where a.�Զ�ע�� = [1] And a.Md5 Is Not Null"
    Set rsTmp = OpenSQLRecord(strSql, "��ȡ�ļ��嵥", RFT_NETGAC)
    IsMustGACUTIL = rsTmp!���� > 0
    Exit Function
errH:
    gobjTrace.WriteInfo "IsMustGACUTIL", "��ȡGACUTILע�Ჿ��", Err.Description
    Err.Clear
End Function

