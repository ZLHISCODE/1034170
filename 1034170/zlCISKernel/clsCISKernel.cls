VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCISKernel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'��ҩ�䷽��������Ҫ�Ĳ�����Ϣ
Public Type TYPE_PatiInfoEx
        lng����ID As Long
        lng��ҳID As Long '��ҳID
        str�Һŵ� As String '�Һŵ���
        str�Ա� As String  '�����Ա�
        intӤ�� As Integer
        int���� As Integer 'ҽ�����˵�����
        blnҽ�� As Boolean '�Ƿ�ҽ������
        lng���˿���ID As Long '����ȷ����ҩ�䷽��ȱʡҩ��
End Type

Public Sub InitCISKernel(cnMain As ADODB.Connection, frmMain As Object, ByVal lngSys As Long, ByVal MainPrivs As String)
    Set gfrmMain = frmMain
    Set gcnOracle = cnMain
    
    glngSys = lngSys
    gMainPrivs = MainPrivs
    
    gstrSysName = GetSetting("ZLSOFT", "ע����Ϣ", "gstrSysName", "")
    gstrProductName = GetSetting("ZLSOFT", "ע����Ϣ", "��Ʒ����", "����")
    gstrUnitName = GetUnitName
    
    '��ʼ������Ա��Ϣ
    Call GetUserInfo
    
    '��ʼ��ϵͳ����:�ɷ�ɢ�������GetForm��ȥ��ʼ��
    Call InitSysPar
End Sub


Public Sub ShowDiagHelp(ByVal bytModal As Byte, ByVal frmParent As Object, Optional ByVal lngItemID As Long, Optional ByVal blnShowInTaskBar As Boolean)
    Call frmDiagHelp.ShowMe(bytModal, frmParent, lngItemID, blnShowInTaskBar)
End Sub

Public Sub ShowClincHelp(ByVal bytModal As Byte, ByVal frmParent As Object, Optional ByVal lngItemID As Long, Optional ByVal blnShowInTaskBar As Boolean, Optional ByVal strItems As String)
    Call frmClinicHelp.ShowMe(bytModal, frmParent, lngItemID, blnShowInTaskBar, strItems)
End Sub

Public Function ShowSchemeEdit(frmParent As Object, ByVal int��Χ As Long, Optional rsScheme As ADODB.Recordset, _
    Optional ByVal blnView As Boolean, Optional ByVal bln��ʾȱʡ�� As Boolean, Optional ByVal strʹ�ÿ��� As String, _
    Optional ByVal byt���� As Byte, Optional ByVal str���Ʒ��� As String, Optional ByVal str�������� As String, _
    Optional ByVal strִ�з��� As String) As ADODB.Recordset
'���ܣ��༭���׷����ӿ�
'������int��Χ=1-����ʹ��,2-סԺʹ��,3-�����סԺ������ʹ��
'      blnView=�Ƿ���鿴
'      bln��ʾȱʡ��=�ٴ�·����Ŀ�������ʱ"ѡ��ʹ��"
'���أ����"������Ŀ���"��ͬ�ṹ�Ķ�̬��¼��,���ȡ���򷵻�Nothing
    Set ShowSchemeEdit = frmSchemeEdit.ShowMe(frmParent, int��Χ, rsScheme, blnView, bln��ʾȱʡ��, strʹ�ÿ���, byt����, str���Ʒ���, str��������, strִ�з���)
End Function

Public Function ShowAdviceEdit(frmParent As Object, ByVal int���� As Integer, ByVal bytUseType As Byte, ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
      ByVal strAdivceOfPath As String, ByVal dat��ʼʱ�� As Date, arrSQL As Variant, strAdviceOfItem As String, Optional ByVal rsLastAdvice As ADODB.Recordset, _
       Optional ByVal datSendTime As Date, Optional ByRef str·����ĿIDs As String, Optional ByRef objMip As Object, Optional ByRef datPathOut As Date) As Boolean
'���ܣ������ٴ�·����Ŀ��Ӧ��ҽ��
'������
'      int����          =0-ҽ��վ����,1-��ʿվ����
'      bytUseType       =1:·������,2=����·������Ŀ
'      strAdivceOfPath  =bytUseType=1ʱ������ҽ������ID:Ӥ�����:·����ĿID,...������227:0:38,335:1:69
'                       =bytUseType=2ʱ������ҽ��ID��,��ʱ����û���ύ������2315,2316
'      dat��ʼʱ��      =bytUseType=1ʱ����������ҽ���Ŀ�ʼʱ��
'                       =bytUseType=2ʱ�����벡��·�����һ�����ɵĵ�ǰ����
'      rs��ѡ��Ŀ��ҽ�� =bytUseType=1ʱ��������ѡ��Ŀ��ҽ�������ֶΰ�������Ŀid,����ҽ��ID����id,������Ŀid,�ǰ�ҽ�����ɵ���������ŵ�
'      datSendTime    =bytUseType=1ʱ������·��������ʱ��
'      objMip ��Ϣ���� zl9ComLib.clsMipModule
'���أ�arrSQL           =ҽ�������У�Ե�SQL���
'      strAdviceOfItem  =bytUseType=1ʱ��·����Ŀ��ҽ��ID�Ķ�Ӧ,����38:1983,69:1978
'                       =bytUseType=2ʱ������ҽ��ID��,��ʱ����û���ύ������2315,2316
'      ShowAdviceEdit   =true:���棬false:ȡ��
'      str·����ĿIDs   =·������ʱ��ҽ�޸��˵��䷽�ģ��ҳ����������޸��䷽�ı�������Ŀ����Ӧ�ı���ԭ����ĿID1|�������1,��Ŀ2|�������2��������
'      datPathOut      =bytUseType=2ʱ,����ҽ���Ŀ�ʼ����,����������ȷ����·������Ŀ���ӵ�·��������ĳһ�졣
    ShowAdviceEdit = frmInAdviceEdit.ShowMeByPath(frmParent, int����, gMainPrivs, bytUseType, lng����ID, lng��ҳID, strAdivceOfPath, dat��ʼʱ��, arrSQL, strAdviceOfItem, rsLastAdvice, datSendTime, str·����ĿIDs, objMip, datPathOut)
End Function

Public Function ShowAdviceOperate(frmParent As Object, ByVal int���� As Integer, _
    ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lng����ID As Long, ByVal strAdviceOfItem As String, ByVal dateStop As Date, Optional ByRef objMip As Object) As Boolean
'���ܣ������ٴ�·��ִ�к����ֹͣ�����Ĺ���
'      strAdviceOfItem=·����Ŀ��ص�ҽ��ID
'      dateStop=ֹͣ������ʱ��
'      objMip ��Ϣ���� zl9ComLib.clsMipModule
    ShowAdviceOperate = frmAdviceOperate.ShowMe(frmParent, gMainPrivs, 1, lng����ID, lng��ҳID, lng����ID, 0, int���� = 1, False, 1, strAdviceOfItem, dateStop, , , , , , lng����ID, , objMip)
End Function

Public Function AdviceOperate(frmParent As Object, ByVal strMainPrivs As String, ByVal int���� As Integer, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lng����ID As Long, _
    Optional ByVal lngҽ��ID As Long, Optional ByRef objMip As Object, Optional ByRef strPatisOut As String, Optional ByVal blnOnePati As Boolean) As Boolean
'���ܣ�˫����Ϣ�����б�ʱ����У�Ի���ȷ��ֹͣ����
    AdviceOperate = frmAdviceOperate.ShowMe(frmParent, strMainPrivs, int����, lng����ID, lng��ҳID, lng����ID, lngҽ��ID, True, False, 0, , , blnOnePati, , , , , , , objMip, strPatisOut)
End Function

Public Function AdviceSend(frmParent As Object, ByVal lng����ID As Long, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strMainPrivs As String, Optional ByRef objMip As Object) As Boolean
'���ܣ�˫����Ϣ�����б�ʱ����ҽ�����͹���
    AdviceSend = frmAdviceSendALL.ShowMe(frmParent, lng����ID, lng����ID, lng��ҳID, strMainPrivs, False, True, , , objMip)
End Function

Public Function AdviceRisReport(frmParent As Object, ByVal lng����ID As Long) As Boolean
'���ܣ�������ӡRISԤԼ��
    AdviceRisReport = frmAdviceRisReport.ShowMe(frmParent, lng����ID)
End Function

Public Function ShowBloodApply(frmParent As Object, ByVal lng��Ѫҽ��ID As Long) As Boolean
'���ܣ��ⲿ������Ѫ���뵥�鿴���ܣ����������סԺֱ���ɹҺŵ��ж�
    Dim rsTmp As Recordset, strSQL As String
    strSQL = "select ����ID,��ҳID,�Һŵ�,Ӥ�� From ����ҽ����¼ Where ID=[1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "BloodApplyView", lng��Ѫҽ��ID)
    
    If rsTmp.RecordCount > 0 Then
        If rsTmp!�Һŵ� & "" <> "" Then
            If gblnѪ��ϵͳ Then
                Call frmBloodApplyNew.ShowMe(frmParent, Val(rsTmp!����ID & ""), 0, 1, 2, lng��Ѫҽ��ID, , , , , , , , 1, rsTmp!�Һŵ� & "")
            Else
                Call frmBloodApply.ShowMe(frmParent, Val(rsTmp!����ID & ""), 0, 1, 2, lng��Ѫҽ��ID, , , , , , , , 1, rsTmp!�Һŵ� & "")
            End If
        Else
            If gblnѪ��ϵͳ Then
                Call frmBloodApplyNew.ShowMe(frmParent, Val(rsTmp!����ID & ""), Val(rsTmp!��ҳID & ""), 0, 2, lng��Ѫҽ��ID, , , , , , , , , , , , , Val(rsTmp!Ӥ�� & ""))
            Else
                Call frmBloodApply.ShowMe(frmParent, Val(rsTmp!����ID & ""), Val(rsTmp!��ҳID & ""), 0, 2, lng��Ѫҽ��ID, , , , , , , , , , , , , Val(rsTmp!Ӥ�� & ""))
            End If
        End If
    Else
        MsgBox "û���ҵ���ָ������Ѫҽ����", vbInformation, gstrSysName
        Exit Function
    End If
    
    ShowBloodApply = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ShowPacsApplication(frmParent As Object, ByVal lng���ҽ��ID As Long, Optional ByVal lngNo As Long) As Boolean
'���ܣ��ⲿ���ü�����뵥�鿴����
    ShowPacsApplication = ShowApply���(frmParent, lngNo, lng���ҽ��ID)
End Function

Public Function ChargeDelAudit(frmParent As Object, ByVal lngִ�п���ID As Long, ByVal lng����ID As Long) As Boolean
'���ܣ�˫����Ϣ�����б�ʱ����������˹���
    Dim strCommon As String
    Dim intAtom As Integer
    
    '���÷��ò�������
    On Error Resume Next
    
    If gobjInExse Is Nothing Then Set gobjInExse = CreateObject("zl9InExse.clsInExse")
    err.Clear: On Error GoTo 0
    
    If gobjInExse Is Nothing Then Exit Function
    
    '�������úϷ�������
    strCommon = Format(Now, "yyyyMMddHHmm")
    strCommon = TranPasswd(strCommon) & "||" & AnalyseComputer
    intAtom = GlobalAddAtom(strCommon)
    Call SaveSetting("ZLSOFT", "����ȫ��", "����", intAtom)
    ChargeDelAudit = gobjInExse.CallReCharge(frmParent, gcnOracle, gstrDBUser, glngSys, 1, 1, lngִ�п���ID, GetInsidePrivs(pҽ�����ѹ���), lng����ID)
    Call GlobalDeleteAtom(intAtom)
    
End Function

Public Function zlPrintAdvice(frmParent As Object, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lngBabyNum As Long, ByVal int��Ч As Integer, Optional ByVal strPrinterName As String, Optional ByVal intMode As Integer) As Boolean
'���ܣ���ӡ����ҽ������������ӡ�ӵ�һ������һҳ����ӡ֮������ҽ���Ĵ�ӡ���
'������int��Ч 0��������1��������strPrinterName ��ӡ�����ƣ�intMode=1-Ԥ��,2-��ӡ��Ĭ��Ϊ��ӡ
    Dim strSQL As String
    Dim strReport As String
    Dim strWhere As String
    Dim strPrinterTmp As String
    
    Dim dat�������ʱ�� As Date
    Dim dat���� As Date
    
    Dim rsTmp As ADODB.Recordset
    Dim rsTable As ADODB.Recordset
    Dim rsColumn As ADODB.Recordset
    Dim rsPLR As ADODB.Recordset
    
    Dim blnHead As Boolean, blnBody As Boolean
    
    Dim i As Long
    Dim lngBodyH As Long, lngHeadH As Long
    Dim lngPrintPos As Long
    Dim lngPrintType As Long
    Dim lngPages As Long
    Dim lng�������� As Long
    
    On Error GoTo errH
    
    strReport = IIF(int��Ч = 0, "ZL1_INSIDE_1254_1", "ZL1_INSIDE_1254_2")
    
    '1���ж�����ҽ��
    strSQL = "select 1 from ����ҽ����¼ where ����ID=[1] and ��ҳID=[2] and nvl(Ӥ��,0)=[3] and ҽ����Ч=[4]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlPrintAdvice", lng����ID, lng��ҳID, lngBabyNum, int��Ч)
    If rsTmp.EOF Then Exit Function
    
    '2������ʱ��
    strSQL = "Select ҽ������ʱ�� as ʱ�� From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����ʱ��", lng����ID, lng��ҳID)
    dat�������ʱ�� = Nvl(rsTmp!ʱ��, CDate("1900-01-01"))
    
    '3�����㱨��������
    strSQL = "Select A.ID as ����ID,B.ID,B.W,B.H,B.�и�,B.����" & _
        " From zlReports A,zlRPTItems B" & _
        " Where A.ID=B.����ID And B.����=4 And Nvl(A.ϵͳ,0)=[1] And A.���=[2] And B.��ʽ��=[3]" & _
        " Order by B.W*B.H Desc"
    Set rsTable = zlDatabase.OpenSQLRecord(strSQL, "zlPrintAdvice", glngSys, strReport, 1)
    If rsTable.EOF Then Exit Function
    
    strSQL = "Select ���,��ͷ,���� From zlRPTItems Where ����ID=[1] And ��ʽ��=[2] And �ϼ�ID=[3] And ����=6 Order by ���"
    Set rsColumn = zlDatabase.OpenSQLRecord(strSQL, "zlPrintAdvice", Val(rsTable!����ID), 1, Val(rsTable!ID))
    If rsColumn.EOF Then Exit Function
    
    '���´�������Զ��屨���еķ�����д
    '----------------------------------
    '�����ͷ�߶�:�Ե�һ��Ϊ׼
    For i = 0 To UBound(Split(rsColumn!��ͷ, "|"))
        lngHeadH = lngHeadH + Val(Split(Split(rsColumn!��ͷ, "|")(i), "^")(1))
    Next
    
    '�������߶�
    blnHead = False: blnBody = False
    rsColumn.MoveFirst
    Do While Not rsColumn.EOF
        i = UBound(Split(rsColumn!��ͷ, "|"))
        If i > 0 Then
            blnHead = True
        ElseIf i = 0 Then
            blnHead = blnHead Or (Split(Split(rsColumn!��ͷ, "|")(i), "^")(2) <> "#")
        End If
        blnBody = blnBody Or Not IsNull(rsColumn!����)
        rsColumn.MoveNext
    Loop
    If Not blnHead And blnBody Then '���б���
        lngBodyH = rsTable!H
    Else
        If rsTable!H - lngHeadH + 15 < 0 Then
            lngBodyH = 0
        Else
            lngBodyH = rsTable!H - lngHeadH + 15
        End If
    End If
    '�������
    lng�������� = Int(lngBodyH / rsTable!�и�) * Nvl(rsTable!����, 1)
    
    '4.����ҽ����ӡ����
    strSQL = "Zl_����ҽ����ӡ_Insert(" & lng����ID & "," & lng��ҳID & "," & lngBabyNum & "," & int��Ч & "," & lng�������� & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, "zlPrintAdvice")
    
    '5����ӡҽ��    ������ж���ҳ
    strSQL = "select max(ҳ��) as ҳ�� from ����ҽ����ӡ  where ����ID=[1] And ��ҳID=[2] And Nvl(Ӥ��,0)=[3] And ��Ч=[4]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "zlPrintAdvice", lng����ID, lng��ҳID, lngBabyNum, int��Ч)
    
    lngPages = Val(rsTmp!ҳ�� & "")
    
    '��ӡ������
    If strPrinterName <> "" Then
        strPrinterTmp = Printer.DeviceName
        For i = 0 To Printers.Count - 1
            If Printers(i).DeviceName = strPrinterName Then
                Set Printer = Printers(i)
                Exit For
            End If
        Next
    End If
    
    Call ReportOpen(gcnOracle, glngSys, strReport, frmParent, "����ID=" & lng����ID, "��ҳID=" & lng��ҳID, "Ӥ��=" & lngBabyNum, "��ӡģʽ=" & lngPrintType, _
        "ͣ����ӡ=0", "��ʼ�к�=1", "StartPageNum=1", "��ʼҳ��=1", "����ҳ��=" & lngPages, "PressWorkFirst=0", IIF(intMode = 1, 1, 2))
        
    '��ӡ���ָ�
    If strPrinterTmp <> "" Then
        For i = 0 To Printers.Count - 1
            If Printers(i).DeviceName = strPrinterTmp Then
                Set Printer = Printers(i)
                Exit For
            End If
        Next
    End If
    zlPrintAdvice = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function CloseWindows() As Boolean
    '--------------------------------------
    '����:�ر������Ӵ���
    '--------------------------------------
    Dim frmThis As Form
    err = 0: On Error Resume Next
    For Each frmThis In Forms
        Unload frmThis
    Next
    CloseWindows = (Forms.Count = 0)
End Function

Private Sub Class_Terminate()
'�ڵ���̨�˳�ʱ����,��ʱע��
    Set grsҽ�Ƹ��ʽ = Nothing
End Sub

Public Function ShowApplyCustom(frmParent As Object, ByVal lngFileID As Long, Optional ByRef lng������� As Long) As Boolean
'���ܣ��鿴�Զ������뵥
'������frmParent �������壻
    Dim objForm As New frmApplyCustom
    
    ShowApplyCustom = objForm.ShowMe(frmParent, 0, 2, 0, 0, 0, lngFileID, lng�������)
End Function

Public Function ShowAppCritical(frmParent As Object, ByVal blnModal As Boolean, ByVal intType As Integer, ByVal int�������� As Integer, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal str�Һŵ� As String, ByVal lngӤ�� As Long, ByRef lngΣ��ֵID As Long, ByVal lngҽ��ID As Long, _
    Optional ByVal lng�걾id As Long, Optional ByVal strΣ��ֵ���� As String, Optional ByVal dat����ʱ�� As Date, Optional ByVal lng�������ID As Long, Optional ByVal str������ As String, Optional ByRef objMip As Object) As Boolean
'���ܣ�Σ��ֵ�༭
    ShowAppCritical = frmCriticalEdit.ShowMe(frmParent, blnModal, intType, int��������, lng����ID, lng��ҳID, str�Һŵ�, lngӤ��, lngΣ��ֵID, lngҽ��ID, lng�걾id, strΣ��ֵ����, dat����ʱ��, lng�������ID, str������, objMip)
End Function

Public Function ShowQueryCritical(frmParent As Object, ByVal blnModal As Boolean, ByVal int���� As Integer, ByVal int���� As Integer, ByVal lng����id As Long, ByVal lng����ID As Long, ByRef objMip As Object) As Boolean
'���ܣ���ѯΣ��ֵ
    ShowQueryCritical = frmCriticalQuery.ShowMe(frmParent, blnModal, int����, int����, lng����id, lng����ID, objMip)
End Function

Public Function ShowDealCritical(frmParent As Object, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal str�Һŵ� As String, ByRef lng��¼ID As Long) As Boolean
'���ܣ��ٴ�����Σ��ֵ���������˴���
'������lng��¼ID �����������δ�����Σ��ֵ��¼
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim blnOK As Boolean
    Dim lngΣ��ֵID As Long
    Dim int�������� As Integer
    Dim lngҽ��ID As Long
 
    If lng��ҳID = 0 Then
        int�������� = 1
        strSQL = "select a.id,a.ҽ��ID,a.����,a.�Ա�,a.����,a.����ʱ��,a.״̬,a.�Ƿ�Σ��ֵ, a.Σ��ֵ���� from ����Σ��ֵ��¼ a where a.�Һŵ�=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "clsCIKernel", str�Һŵ�)
    Else
        int�������� = 2
        strSQL = "select a.id,a.ҽ��id,a.����,a.�Ա�,a.����,a.����ʱ��,a.״̬,a.�Ƿ�Σ��ֵ, a.Σ��ֵ����  from ����Σ��ֵ��¼ a where a.����ID=[1] and a.��ҳID=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "clsCIKernel", lng����ID, lng��ҳID)
    End If
    
    If rsTmp.EOF Then
        MsgBox "�û���δ�ҵ���ص�Σ��ֵ��¼��", vbInformation, gstrSysName
        Exit Function
    End If
    
    If rsTmp.RecordCount > 1 Then
        lngΣ��ֵID = frmCriticalQuery.ShowMeSelCard(frmParent, rsTmp)
        If lngΣ��ֵID = 0 Then Exit Function
        rsTmp.Filter = "ID=" & lngΣ��ֵID
        If rsTmp.EOF Then Exit Function
    End If
    
    lngΣ��ֵID = Val(rsTmp!ID & "")
    lngҽ��ID = Val(rsTmp!ҽ��ID & "")
    
    blnOK = frmCriticalEdit.ShowMe(frmParent, True, 3, int��������, lng����ID, lng��ҳID, str�Һŵ�, 0, lngΣ��ֵID, lngҽ��ID)
    If blnOK Then
        lng��¼ID = lngΣ��ֵID
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function ShowMainQueryCritical(ByVal lngSys As Long, ByVal lngModul As Long, ByRef frmParent As Object, ByVal strPrivs As String) As Boolean
'���ܣ�������ѯ����
    ShowMainQueryCritical = frmCriticalQuery.ShowMeQuery(lngSys, lngModul, frmParent, strPrivs)
End Function
