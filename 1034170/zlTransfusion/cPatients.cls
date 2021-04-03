VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPatients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cPatient"
Attribute VB_Ext_KEY = "Member0" ,"cPatient"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'�ֲ����������漯��
Private mCol As Collection
Public mSeatings As Seatings
Public Enum PATI_COLUM
    col_�Ŷ�״̬ = 0
    col_ͼ�� = 1
    col_calling = 2
    col_���� = 3
    
    col_����ʱ�� = 4
    col_˳��� = 5
    col_����̨ = 6
    col_�Һŵ� = 7
    col_��ҳID = 8
    col_��λ�� = 9
    col_���￨�� = 10
    
    col_��Ȩ�� = 11
    col_�Ա� = 12
    col_���� = 13
    
    col_����ID = 14
    col_����� = 15
    col_���˿��� = 16
    col_�ѱ� = 17
    col_������� = 18
    col_ҽ�� = 19
    col_����֤�� = 20
    col_IC���� = 21
    col_�Һ�ʱ�� = 22
    col_��ʼʱ�� = 23
    col_��ʼ����Ա = 24
    col_����ʱ�� = 25
    col_��������Ա = 26
    col_key = 27
End Enum

Private blnRptInit As Boolean

Public Function ClearPatients(ByRef Rpt As Object) As Boolean
'������ϼ�ָ���ؼ��е�����
    Dim lngIndex As Long
    
    For lngIndex = 1 To Me.Count
        Remove 1
    Next
    
    If Not Rpt Is Nothing Then
        If TypeName(Rpt) = "ReportControl" Then
            Rpt.Records.DeleteAll
        End If
    End If
End Function

Public Function FetchPatients(ByVal deptID As Long, ByVal BeginTime As Date, ByVal EndTime As Date, _
        Optional ByVal blnIsReady As Boolean = False, Optional ByVal strInfo As String, Optional ByRef Rpt As Object, _
        Optional ByRef img As Object, _
        Optional ByRef objSquareCard As Object)
'�����ݿ���ȡ�����ݼ��뼯�ϣ����������ָ���ؼ���
'Ŀǰֻ֧��ReportControl
'д��ǰ�����ClearPatients���̡�
'������
'  strInfo����ȡָ��������Ϣ����ʽ�� ����(1-���￨;2-�����;3-���ݺ�;4-����;5-����֤;6-IC��)��һ����|����|һ��ͨ�����ID

    Dim blnToRpt As Boolean
    Dim strSQL As String, rsTmp As ADODB.Recordset, objPati As cPatient
    Dim str��� As String, rsZD As ADODB.Recordset, strPar As String, strType As String, i As Integer
    Dim strSeatKey As String, strTmp As String, strTmpA As String, strTmpB As String
    Dim strVal As String
    Dim intType As Integer
    Dim lng����ID As Long, lngCardType As Long
    Dim dateTmp As Date
    Dim arrVal As Variant
    Dim strWatchSQL As String

    On Error GoTo errHandle
    strPar = zlDatabase.GetPara("��ʾ��������", glngSys, 1264, "1,1,1,1")
    For i = 0 To 3
        strType = strType & IIf(Val(Split(strPar, ",")(i)) = 1, "," & i, "")
    Next
    
    If Not Rpt Is Nothing Then
        If TypeName(Rpt) = "ReportControl" Then
            blnToRpt = True
        End If
    End If
    If EndTime < BeginTime Then Exit Function
    
    If Trim(strInfo) = "" And blnIsReady = False Then
        'ˢ������
        dateTmp = DateAdd("d", -3, BeginTime)
    Else
        '�ӵ�������ָ������
        arrVal = Split(strInfo, "|")
        intType = Val(arrVal(0))        '���п��ķ���
        strVal = arrVal(1)              '����
        Select Case intType
            Case 2
                strTmpA = " And c.�����=[5] "
            Case 3
                strTmpA = " And a.�Һŵ�=[5] "
            Case 4
                strTmpA = " And c.����=[5] "
            Case 5
                strTmpA = " And c.����֤��=[5] "
            Case 1, 6 To 99
                '1-���￨��6-99ҽ�ƿ���������Ϣ����һ����ͨ��IC���ҵ����ˣ��˴�ͨ�����������õ�����ID
                lngCardType = Val(arrVal(2))    'һ��ͨ���ID
                If objSquareCard Is Nothing Then
                    strTmpA = " And c.IC����=[5] "
                Else
                    strTmpA = " And c.����id=[6] "
                    If objSquareCard.zlGetPatiID(lngCardType, UCase(strVal), , lng����ID) = False Then lng����ID = 0
                End If
            Case Else
                strTmpA = " And a.�Һŵ� is null "
        End Select
    End If
    
    
    '��ʼ����λ��
    mSeatings.getSeatings (deptID)
    
    Call ClearPatients(Rpt)
           
'    'δ�ӵ���ָ��ҽ���������ڷ�Χ�Ĳ�����Ϣ��������ӵ�����ҽ���������ڷ�ΧΪһ����
'    strSQL = "Select a.*, f_List2str(Cast(Collect(b.�������) As t_Strlist), '��') ������� " & _
'             "From ( " & _
'             "Select Distinct *" & vbNewLine & _
'             "From (Select b.����id, Nvl(b.�����, 0) As �����, Nvl(b.���￨��, '��') ���￨��, b.����֤��, b.Ic����, Nvl(b.�ѱ�, '��') �ѱ�, b.����," & vbNewLine & _
'             "              Nvl(b.�Ա�, 'δ֪') �Ա�, Nvl(b.����, 'δ֪') ����, To_Char(Nvl(d.����,B.ִ��ʱ��), 'yyyy-MM-dd HH24:MI:SS') As ����ʱ��,b.�Ǽ�ʱ�� as �Һ�ʱ��, " & vbNewLine & _
'             "              Nvl(b.ִ����, '��') ҽ��, Nvl(b.�Һŵ�, '��') �Һŵ�, Nvl(c.���, 0) ��λ���, Nvl(c.���, '��') As ��λ��, LPad(To_Char(Nvl(d.˳��, 0)), 3, '0') ˳��," & vbNewLine & _
'             "              Decode(Nvl(d.��Ȩ��, 0), 0, Replace(Nvl(˳��, 0), '_', ''), d.��Ȩ��) As ��Ȩ��," & vbNewLine & _
'             "              Decode(Nvl(d.״̬, -1), 1, '1-����Һ', 2, '2-����', 3, '3-�˺�', 4, '4-����', 5, '5-������', 6, '6-��ִ��', 7, '7-ִ����'," & vbNewLine & _
'             "                      '0-δ�ӵ�') As �Ŷ�״̬, Nvl(b.���˿���, '��') ���˿���, b.�Һŵ�id, d.���б�־, d.��ʼʱ��, d.��ʼ����Ա, d.����ʱ��, d.��������Ա,d.����̨ " & vbNewLine & _
'             "       From (Select ����id, ����, To_Char(����, 'hhmmss') || '_' || LPad(To_Char(Nvl(˳���, 0)), 3, '0') As ˳��, ��Ȩ��, ״̬, ���б�־, ��ʼʱ��," & vbNewLine & _
'             "                     ��ʼ����Ա, �Һŵ�, ��������Ա, ����ʱ��,����̨ " & vbNewLine & _
'             "              From �ŶӼ�¼" & vbNewLine & _
'             "              Where ���� Between [2] And [3] And ����id = [1]) D," & vbNewLine & _
'             "            (Select ����id, ���, ��� From ��λ״����¼ Where ����id = [1]) C," & vbNewLine & _
'             "            (Select b.����id, c.�����, c.���￨��, c.�ѱ�, c.����, c.�Ա�, c.����, d.ִ��ʱ��, d.�Ǽ�ʱ��, c.��������, d.ִ����, d.Id As �Һŵ�id, b.�Һŵ�," & vbNewLine & _
'             "                     e.���� || '[' || e.���� || ']' As ���˿���, c.����֤��, c.Ic����" & vbNewLine & _
'             "              From ������ĿĿ¼ F, ���ű� E, ���˹Һż�¼ D, ������Ϣ C, ����ҽ����¼ B, ����ҽ������ A" & vbNewLine & _
'             "              Where Instr([4], Nvl(f.ִ�з���, 0)) > 0 And b.������Ŀid = f.Id And d.ִ�в���id = e.Id And b.����id = c.����id And" & vbNewLine & _
'             "                    a.ҽ��id = b.Id And b.�Һŵ� = d.No And b.������Դ = 1 And d.��¼���� = 1 And d.��¼״̬ = 1 And" & vbNewLine & _
'             "                    a.����ʱ�� Between [2] And [3] And a.ִ�в���id = [1]" & vbCr & _
'             strTmpA & vbCr & _
'             "      ) B " & vbNewLine & _
'             "Where b.����id = c.����id(+) And b.����id = d.����id(+) And b.�Һŵ� = d.�Һŵ�(+) And d.�Һŵ�(+) Is Null "

    'δ�ӵ�������������ۣ���ָ��ҽ���������ڷ�Χ�Ĳ�����Ϣ��������ӵ�����ҽ���������ڷ�ΧΪһ����
    strSQL = "Select a.*, f_List2str(Cast(Collect(b.�������) As t_Strlist), '��') ������� " & _
             "From ( " & _
             "Select Distinct *" & vbNewLine & _
             "From (Select b.����id, Nvl(b.�����, 0) As �����, Nvl(b.���￨��, '��') ���￨��, b.����֤��, b.Ic����, Nvl(b.�ѱ�, '��') �ѱ�, b.����," & vbNewLine & _
             "          Nvl(b.�Ա�, 'δ֪') �Ա�, Nvl(b.����, 'δ֪') ����, To_Char(Nvl(d.����,B.ִ��ʱ��), 'yyyy-MM-dd HH24:MI:SS') As ����ʱ��,b.�Ǽ�ʱ�� as �Һ�ʱ��, " & vbNewLine & _
             "           Nvl(b.ִ����, '��') ҽ��, Nvl(b.�Һŵ�, '��') �Һŵ�, Nvl(c.���, 0) ��λ���, Nvl(c.���, '��') As ��λ��, LPad(To_Char(Nvl(d.˳��, 0)), 3, '0') ˳��," & vbNewLine & _
             "           Decode(Nvl(d.��Ȩ��, 0), 0, Replace(Nvl(˳��, 0), '_', ''), d.��Ȩ��) As ��Ȩ��," & vbNewLine & _
             "           Decode(Nvl(d.״̬, -1), 1, '1-����Һ', 2, '2-����', 3, '3-�˺�', 4, '4-����', 5, '5-������', 6, '6-��ִ��', 7, '7-ִ����','0-δ�ӵ�') As �Ŷ�״̬, " & vbNewLine & _
             "           Nvl(b.���˿���, '��') ���˿���, b.��ҳid, b.�Һ�id, d.���б�־, d.��ʼʱ��, d.��ʼ����Ա, d.����ʱ��, d.��������Ա,d.����̨,b.������Դ " & vbNewLine & _
             "      From (Select ����id, ��ҳid, ����, To_Char(����, 'hhmmss') || '_' || LPad(To_Char(Nvl(˳���, 0)), 3, '0') As ˳��, " & vbNewLine & _
             "                ��Ȩ��, ״̬, ���б�־, ��ʼʱ��, ��ʼ����Ա, �Һŵ�, ��������Ա, ����ʱ��,����̨ " & vbNewLine & _
             "            From �ŶӼ�¼ " & vbNewLine & _
             "            Where ���� Between [2] And [3] And ����id = [1]) D," & vbNewLine & _
             "      (Select ����id, ���, ��� From ��λ״����¼ Where ����id = [1]) C," & vbNewLine & _
             "      (Select b.ҽ��id,a.����id, c.�����, c.���￨��, c.�ѱ�, c.����, c.�Ա�, c.����, Decode(D1.ִ��ʱ��, Null, a.��ʼִ��ʱ��, D1.ִ��ʱ��) ִ��ʱ��," & vbNewLine & _
             "           Decode(D1.�Ǽ�ʱ��, Null, a.����ʱ��, D1.�Ǽ�ʱ��) �Ǽ�ʱ��, c.��������, decode(D1.ִ����, null, a.����ҽ��, D1.ִ����) ִ����," & vbNewLine & _
             "           D2.��ҳid, D1.Id �Һ�id, a.�Һŵ�, f.���� || '[' || f.���� || ']' As ���˿���, c.����֤��, c.Ic����, " & vbNewLine & _
             "           Decode(Nvl(D2.��ҳid, 0), 0, 0, 1) ������Դ, Nvl(d2.��������, -1) �������� " & vbNewLine & _
             "       From ����ҽ����¼ A, ����ҽ������ B, ������Ϣ C, ���˹Һż�¼ D1, ������ҳ D2, ������ĿĿ¼ E, ���ű� F " & vbNewLine & _
             "       Where a.Id = b.ҽ��id And a.����id = c.����id And a.�Һŵ� = D1.No(+) And a.����id = D2.����id(+) And a.��ҳid = D2.��ҳid(+) " & vbNewLine & _
             "           And a.������Ŀid = e.Id And a.ִ�п���id = f.Id And a.������Դ In (1, 2) " & vbNewLine & _
             "           And b.ִ�в���id = [1] And b.����ʱ�� between [2] and [3] " & vbNewLine & _
             "           And D1.��¼����(+) = 1 And D1.��¼״̬(+) = 1 " & vbNewLine & _
             "           And Instr([4], Nvl(e.ִ�з���, 0)) > 0 "
    strSQL = strSQL & strTmpA & vbCr & ") B " & vbCr & _
             "Where b.����id = c.����id(+) And b.����id = d.����id(+) And b.�Һŵ� = d.�Һŵ�(+) And nvl(b.��ҳid,0)=nvl(d.��ҳid(+),0) And b.�������� in (-1,1) "

'    '�ѽӵ��������ڵĲ����ŶӼ�¼��1,5,6,7)��ָ�����ҡ�ָ�����ڵĲ����ŶӼ�¼
'    strSQL = strSQL & " Union all " & vbNewLine & _
'             "Select c.����id, nvl(b.�����,0) �����, nvl(b.���￨��, '��') ���￨��, b.����֤��, b.ic����, nvl(b.�ѱ�, '��') �ѱ�, " & _
'             "    b.����, nvl(b.�Ա�, 'δ֪') �Ա�, nvl(b.����, 'δ֪') ����, To_Char(c.����, 'yyyy-MM-dd HH24:MI:SS') As ����ʱ��, " & _
'             "    a.�Ǽ�ʱ�� �Һ�ʱ��, nvl(a.ִ����, '��') ҽ��, nvl(c.�Һŵ�, '��') �Һŵ�, Nvl(e.���, 0) ��λ���, Nvl(e.���, '��') As ��λ��, " & _
'             "    LPad(To_Char(nvl(c.˳���, 0)), 3, '0') ˳��, To_Char(nvl(c.��Ȩ��, 0)) ��Ȩ��, " & _
'             "    Decode(Nvl(c.״̬, -1), 1, '1-����Һ', 2, '2-����', 3, '3-�˺�', 4, '4-����', 5, '5-������', 6, '6-��ִ��', 7, '7-ִ����', '0-δ�ӵ�') As �Ŷ�״̬, " & _
'             "    Nvl(d.���� || '[' || d.���� || ']', '��') ���˿���, a.id �Һŵ�id, c.���б�־, c.��ʼʱ��, c.��ʼ����Ա, c.����ʱ��, c.��������Ա, c.����̨ " & vbNewLine & _
'             "From ���˹Һż�¼ A, ������Ϣ B, �ŶӼ�¼ C, ���ű� D, ��λ״����¼ E " & vbNewLine & _
'             "Where a.No = c.�Һŵ� And a.ִ�в���id = d.Id And b.����id = c.����id And c.����id = e.����id(+) And c.����id = e.����id(+) " & _
'             "    And a.ִ��״̬ In (1, 2) And a.��¼���� = 1 And a.��¼״̬ = 1 " & _
'             IIf(Trim(strInfo) = "" And blnIsReady = False, _
'                " And ( c.���� Between [5] And [2] And instr('1,5,6,7',c.״̬)>0 Or c.���� Between [2] And [3] ) ", _
'                " And c.���� Between [2] And [3] ") & _
'             "    And c.����id = [1] " & strTmpB & vbNewLine & _
'             ")  "
             
    '�ѽӵ�������������ۣ��������ڵĲ����ŶӼ�¼��1,5,6,7)��ָ�����ҡ�ָ�����ڵĲ����ŶӼ�¼
    strSQL = strSQL & vbNewLine & _
             "Union all " & vbNewLine & _
             "Select a.����id, nvl(c.�����,0) �����, nvl(c.���￨��, '��') ���￨��, c.����֤��, c.Ic����, nvl(c.�ѱ�, '��') �ѱ�, c.����, nvl(c.�Ա�, 'δ֪') �Ա�, nvl(c.����, 'δ֪') ����," & vbNewLine & _
             "    To_Char(a.����, 'yyyy-MM-dd HH24:MI:SS') ����ʱ��, Decode(b.�Ǽ�ʱ��, Null, d.�Ǽ�ʱ��, b.�Ǽ�ʱ��) �Һ�ʱ��," & vbNewLine & _
             "    nvl(b.ִ����, d.����ҽʦ) ҽ��, nvl(a.�Һŵ�, '��') �Һŵ�, Nvl(e.���, 0) ��λ���, nvl(e.���, '��') ��λ��," & vbNewLine & _
             "    LPad(To_Char(Nvl(a.˳���, 0)), 3, '0') ˳��, To_Char(Nvl(a.��Ȩ��, 0)) ��Ȩ��," & vbNewLine & _
             "    Decode(Nvl(a.״̬, -1), 1, '1-����Һ', 2, '2-����', 3, '3-�˺�', 4, '4-����', 5, '5-������', 6, '6-��ִ��', 7, '7-ִ����', '0-δ�ӵ�') As �Ŷ�״̬," & vbNewLine & _
             "    Decode(F1.����, Null, F2.���� || '[' || F2.���� || ']', F1.���� || '[' || F1.���� || ']') ���˿���, a.��ҳid, b.id �Һ�id, a.���б�־," & vbNewLine & _
             "    a.��ʼʱ��, a.��ʼ����Ա, a.����ʱ��, a.��������Ա, a.����̨, Decode(Nvl(a.��ҳid, 0), 0, 0, 1) ������Դ " & vbNewLine & _
             "From �ŶӼ�¼ A, ���˹Һż�¼ B, ������Ϣ C, ������ҳ D, ��λ״����¼ E, ���ű� F1, ���ű� F2 " & vbNewLine & _
             "Where a.����id = b.����id(+) And a.�Һŵ� = b.NO(+) And a.����id = d.����id(+) And a.��ҳid = d.��ҳid(+) And a.����id = c.����id " & vbNewLine & _
             "    And a.����id = e.����id(+) And a.����id = e.����id(+) And b.ִ�в���id = F1.Id(+) And d.��Ժ����id = F2.Id(+) " & vbNewLine & _
             "    And b.��¼״̬(+) = 1 And b.��¼����(+) = 1 And Instr(',1,2', b.ִ��״̬(+)) > 0 " & _
             IIf(Trim(strInfo) = "" And blnIsReady = False, _
                 " And (a.���� between [5] And [2] And instr('1,5,6,7', a.״̬) > 0 or a.���� between [2] and [3])", _
                 " And a.���� Between [2] And [3] ") & _
             "    And a.����id = [1] " & strTmpA & vbNewLine
    
    strSQL = strSQL & _
             ") ) A, ������ϼ�¼ B " & vbNewLine & _
             "Where a.����id=b.����id(+) and a.�Һ�id = b.��ҳid(+) " & _
             "Group by a.����id, a.�����, a.���￨��, a.����֤��, a.Ic����, a.�ѱ�, a.����, a.�Ա�, a.����, a.����ʱ��, a.�Һ�ʱ��, a.ҽ��, a.�Һŵ�, a.��λ���" & _
             "  , a.��λ��, a.˳��, a.��Ȩ��, a.�Ŷ�״̬, a.���˿���, a.��ҳid, a.�Һ�id, a.������Դ, a.���б�־, a.��ʼʱ��,a.��ʼ����Ա, a.����ʱ��, a.��������Ա, a.����̨ " & _
             "Order by a.�Ŷ�״̬, a.����ʱ�� "
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "FetchPatients", deptID, BeginTime, EndTime, _
                            strType, IIf(Trim(strInfo) = "" And blnIsReady = False, dateTmp, strVal), lng����ID)
    
    If blnToRpt And Not blnRptInit Then
        Call initObjRpt(Rpt, img)
    End If
    
    Do Until rsTmp.EOF
        Set objPati = New cPatient
        With rsTmp

            str��� = zlCommFun.NVL(.Fields("�������"), "(��)")
            
            objPati.����ID = .Fields("����ID")
            objPati.���˿��� = .Fields("���˿���")
            objPati.�ѱ� = .Fields("�ѱ�")
            objPati.�Һŵ� = zlCommFun.NVL(.Fields("�Һŵ�"))
            objPati.���￨�� = .Fields("���￨��")
            
            objPati.������Դ = .Fields("������Դ")
            If objPati.������Դ = 1 Then
                objPati.����ID = zlCommFun.NVL(.Fields("��ҳID"), 0)
            Else
                objPati.����ID = zlCommFun.NVL(.Fields("�Һ�ID"), 0)
            End If
            
            objPati.����ʱ�� = CDate(zlCommFun.NVL(.Fields("����ʱ��"), 0))
            objPati.��������Ա = "" & zlCommFun.NVL(.Fields("��ʼ����Ա"))
            objPati.����ʱ�� = CDate(IIf("" & .Fields("��ʼʱ��") = "", 0, "" & .Fields("��ʼʱ��")))
            objPati.��ʼ����Ա = "" & zlCommFun.NVL(.Fields("��������Ա"))
            objPati.��ʼʱ�� = CDate(IIf("" & .Fields("����ʱ��") = "", 0, "" & .Fields("����ʱ��")))
            objPati.����̨ = "" & zlCommFun.NVL(.Fields("����̨"))

            objPati.����� = CStr(zlCommFun.NVL(.Fields("�����"), "��"))
            objPati.������� = str���
            objPati.���� = .Fields("����")
            objPati.˳��� = Format(.Fields("����ʱ��"), "yyMMddHHmmss") & "_" & .Fields("˳��")
            
            objPati.���� = "" & zlCommFun.NVL(.Fields("����"))
            objPati.�Ա� = "" & zlCommFun.NVL(.Fields("�Ա�"))
            objPati.ҽ�� = "" & zlCommFun.NVL(.Fields("ҽ��"))
            objPati.�Ŷ�״̬ = .Fields("�Ŷ�״̬")
            objPati.��λ�� = .Fields("��λ��")
            objPati.��Ȩ�� = .Fields("��Ȩ��")
            
            objPati.���б�־ = Val("" & .Fields("���б�־"))
            
            objPati.����֤�� = "" & zlCommFun.NVL(.Fields("����֤��"))
            objPati.IC���� = "" & zlCommFun.NVL(.Fields("IC����"))
            
            objPati.�Һ�ʱ�� = CDate(zlCommFun.NVL(.Fields("�Һ�ʱ��"), 0))
            
            If "" & .Fields("��λ��") <> "" And "" & .Fields("��λ��") <> "��" Then
                strSeatKey = "" & .Fields("��λ���") & "_" & .Fields("��λ��")
                strTmp = Trim$("" & .Fields("��ʼ����Ա"))
                If strTmp <> "" Then strTmp = strTmp & " " & Format("" & .Fields("��ʼʱ��"), "HH:mm")
                    
                mSeatings(strSeatKey).��ʼʱ�� = strTmp
                mSeatings(strSeatKey).��� = str���
                mSeatings(strSeatKey).�Һŵ� = .Fields("�Һŵ�")
            End If
            
            '���벡����Ϣ����ļ���
            If objPati.������Դ = 1 Then
                Call Add(objPati, "" & objPati.����ID & "_" & objPati.����ID)
            Else
                Call Add(objPati, objPati.�Һŵ�)
            End If
            Call AddToRpt(objPati, Rpt)
            
            .MoveNext
        End With
        
    Loop
    If blnToRpt Then Rpt.Populate

    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Add(ByVal objPati As cPatient, Optional sKey As String) As cPatient
    '���Ӷ��󵽼��Ϻ�rpt�ؼ���
    '�����¶���
    Dim objNewMember As cPatient
    Dim blnAdd As Boolean
    
    On Error Resume Next
    
    blnAdd = False
    Set objNewMember = mCol(sKey)
    If objNewMember Is Nothing Then
        Set objNewMember = New cPatient
        blnAdd = True
    End If
    
    On Error GoTo 0
    
    '���ô��뷽��������
    objNewMember.����ID = objPati.����ID
    objNewMember.���˿��� = objPati.���˿���
    objNewMember.�ѱ� = objPati.�ѱ�
    objNewMember.�Һŵ� = objPati.�Һŵ�
    objNewMember.���￨�� = objPati.���￨��
    
    objNewMember.����ʱ�� = objPati.����ʱ��
    objNewMember.�Һ�ʱ�� = objPati.�Һ�ʱ��
    objNewMember.����� = objPati.�����
    objNewMember.������� = objPati.�������
    objNewMember.���� = objPati.����
    objNewMember.˳��� = objPati.˳���
    
    objNewMember.���� = objPati.����
    objNewMember.�Ա� = objPati.�Ա�
    objNewMember.ҽ�� = objPati.ҽ��
    objNewMember.�Ŷ�״̬ = objPati.�Ŷ�״̬
    objNewMember.��λ�� = objPati.��λ��
    objNewMember.��Ȩ�� = objPati.��Ȩ��
    objNewMember.���б�־ = objPati.���б�־
    objNewMember.����֤�� = objPati.����֤��
    objNewMember.IC���� = objPati.IC����
    objNewMember.����̨ = objPati.����̨
    objNewMember.��ʼ����Ա = objPati.��ʼ����Ա
    objNewMember.��������Ա = objPati.��������Ա
    objNewMember.��ʼʱ�� = objPati.��ʼʱ��
    objNewMember.����ʱ�� = objPati.����ʱ��
    
    objNewMember.������Դ = objPati.������Դ
    objNewMember.����ID = objPati.����ID
    
    objNewMember.Key = sKey
    
    If blnAdd Then
        If Len(sKey) = 0 Then
            mCol.Add objNewMember
        Else
            mCol.Add objNewMember, sKey
        End If
    End If

    '�����Ѵ����Ķ���
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Private Sub AddToRpt(ByVal objPati As cPatient, ByRef Rpt As Object)
    Dim objRecord As ReportRecord, objItem As ReportRecordItem
    Dim intIcon As Integer
    Dim str˳��� As String
    If Not Rpt Is Nothing Then
        If TypeName(Rpt) = "ReportControl" Then
            Set objRecord = Rpt.Records.Add
            Call Add_rptItem(objRecord, objPati.�Ŷ�״̬)
            
            
            Set objItem = Add_rptItem(objRecord, "")
            Select Case objPati.��λ��
            Case "��"
                intIcon = 3
            Case Else
                intIcon = 0
            End Select
            objItem.Icon = intIcon
            'calling
            Set objItem = Add_rptItem(objRecord, "")
            'objItem.Icon = intIcon
            If objPati.���б�־ = 1 Then
                objItem.Icon = 5
            Else
                objItem.Icon = 6
            End If
            
            Call Add_rptItem(objRecord, objPati.����)
            
            Call Add_rptItem(objRecord, Format(objPati.����ʱ��, "yy-MM-dd hh:mm"))
            'If InStr(objPati.˳���, "_") > 0 Then
            '    str˳��� = Mid(objPati.˳���, InStr(objPati.˳���, "_") + 1)
            'Else
                str˳��� = objPati.˳���
            'End If
            Call Add_rptItem(objRecord, IIf(Val(str˳���) = "0", "", str˳���))
            Call Add_rptItem(objRecord, objPati.����̨)
            
            '2015-05-06������Ϊ���Һŵ��š�����������Ϊ������ID_��ҳID��
            If objPati.������Դ = 1 Then
                Call Add_rptItem(objRecord, "")
                Call Add_rptItem(objRecord, objPati.����ID)
            Else
                Call Add_rptItem(objRecord, objPati.�Һŵ�)
                Call Add_rptItem(objRecord, "")
            End If
            
            Call Add_rptItem(objRecord, objPati.��λ��)
            Call Add_rptItem(objRecord, objPati.���￨��)
            Call Add_rptItem(objRecord, objPati.��Ȩ��)
            Call Add_rptItem(objRecord, objPati.�Ա�)
            Call Add_rptItem(objRecord, objPati.����)
            Call Add_rptItem(objRecord, objPati.����ID)
            Call Add_rptItem(objRecord, objPati.�����)
            Call Add_rptItem(objRecord, objPati.���˿���)
            Call Add_rptItem(objRecord, objPati.�ѱ�)
            Call Add_rptItem(objRecord, objPati.�������)
            Call Add_rptItem(objRecord, objPati.ҽ��)
            
            Call Add_rptItem(objRecord, objPati.����֤��)
            Call Add_rptItem(objRecord, objPati.IC����)
            Call Add_rptItem(objRecord, objPati.�Һ�ʱ��)
            
            Call Add_rptItem(objRecord, objPati.��ʼʱ��)
            Call Add_rptItem(objRecord, objPati.��ʼ����Ա)
            Call Add_rptItem(objRecord, objPati.����ʱ��)
            Call Add_rptItem(objRecord, objPati.��������Ա)
            Call Add_rptItem(objRecord, objPati.Key)
        End If
        '�ڴ˿������Ӷ������ؼ���֧��
    End If
End Sub

Public Property Get Item(vntIndexKey As Variant) As cPatient
Attribute Item.VB_UserMemId = 0
  On Error Resume Next
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    ' x.Count
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    '������������ For...Each �﷨ö�ٸü��ϡ�
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    '������󴴽�����
    Set mCol = New Collection
    Set mSeatings = New Seatings
End Sub

Private Sub Class_Terminate()
    '����ֹ���ƻ�����
    Set mCol = Nothing
    Set mSeatings = Nothing
End Sub

Private Function Add_rptItem(ByRef objRecord As ReportRecord, ByVal strValues As String) As ReportRecordItem
    
    Set Add_rptItem = objRecord.AddItem(strValues)
    Add_rptItem.Caption = strValues
    
End Function

Public Sub initObjRpt(ByRef rptPati As Object, ByRef img16 As Object)
    
    Dim objCol As ReportColumn

    With rptPati
        .Records.DeleteAll
        .Columns.DeleteAll
        .GroupsOrder.DeleteAll
        .SortOrder.DeleteAll
        
        '����˳�������(�������Ϊ����)�ı��,Ҫ��Find(�к�)������,���Կ���Record(�к�)����������
        Set objCol = .Columns.Add(col_�Ŷ�״̬, "״̬", 0, False): objCol.Sortable = False: objCol.Visible = False
        Set objCol = .Columns.Add(col_ͼ��, "", 18, False): objCol.Sortable = False: objCol.Alignment = xtpAlignmentCenter: objCol.AllowDrag = False
        Set objCol = .Columns.Add(col_calling, "", 18, False): objCol.Sortable = False: objCol.Alignment = xtpAlignmentCenter: objCol.AllowDrag = False
        Set objCol = .Columns.Add(col_����, "����", 55, True)
        
        Set objCol = .Columns.Add(col_����ʱ��, "����ʱ��", 95, True)
        Set objCol = .Columns.Add(col_˳���, "˳���", 60, True)
        Set objCol = .Columns.Add(col_����̨, "����̨", 35, True)
        Set objCol = .Columns.Add(col_�Һŵ�, "�Һŵ�", 68, True)
        Set objCol = .Columns.Add(col_��ҳID, "��ҳID", 45, True)
        Set objCol = .Columns.Add(col_��λ��, "��λ��", 45, True)
        Set objCol = .Columns.Add(col_���￨��, "���￨��", 60, True)
        
        '����������
        Set objCol = .Columns.Add(col_��Ȩ��, "��Ȩ��", 0, False)
        Set objCol = .Columns.Add(col_�Ա�, "�Ա�", 0, False)
        Set objCol = .Columns.Add(col_����, "����", 0, False)
        Set objCol = .Columns.Add(col_����ID, "����ID", 0, False)
        Set objCol = .Columns.Add(col_�����, "�����", 0, False)
        Set objCol = .Columns.Add(col_���˿���, "���˿���", 0, False)
        Set objCol = .Columns.Add(col_�ѱ�, "�ѱ�", 0, False)
        
        Set objCol = .Columns.Add(col_�������, "�������", 0, False)
        Set objCol = .Columns.Add(col_ҽ��, "ҽ��", 0, False)
        
        Set objCol = .Columns.Add(col_����֤��, "����֤��", 0, False)
        Set objCol = .Columns.Add(col_IC����, "IC����", 0, False)
        
        Set objCol = .Columns.Add(col_�Һ�ʱ��, "�Һ�ʱ��", 0, False)
        Set objCol = .Columns.Add(col_��ʼʱ��, "��ʼʱ��", 0, False)
        Set objCol = .Columns.Add(col_��ʼ����Ա, "��ʼ����Ա", 0, False)
        Set objCol = .Columns.Add(col_����ʱ��, "����ʱ��", 0, False)
        Set objCol = .Columns.Add(col_��������Ա, "��������Ա", 0, False)
        Set objCol = .Columns.Add(col_key, "Key", 0, False)
        
        For Each objCol In .Columns
            objCol.Editable = False
            objCol.Groupable = objCol.Index = col_�Ŷ�״̬
            If objCol.Width = 0 Then objCol.Visible = False
        Next
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "�϶��б��⵽����,�����з���..."
            .NoItemsText = "û�п���ʾ�Ĳ���..."
            
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '������SelectionChanged�¼�
        .ShowItemsInGroups = False
        .SetImageList img16
        
        .GroupsOrder.Add .Columns(col_�Ŷ�״̬)
        '.GroupsOrder.Add .Columns(col_����ʱ��)
        .GroupsOrder(0).SortAscending = True '����֮��,��������в���ʾ,�����е������ǲ����
        
        '����֮�����ʧȥ��¼���е�˳��,���ǿ�м���������
        .SortOrder.Add .Columns(col_��Ȩ��)
        .SortOrder(0).SortAscending = True
        .SortOrder.Add .Columns(col_˳���)
        .SortOrder(1).SortAscending = True
    End With
    blnRptInit = True
    
End Sub
Public Function DeptToCbo(ByRef cboDept As Object, ByVal strPrivs As String) As Boolean

'���ܣ���ʼ���ٴ�����

    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str����IDs As String
    
    On Error GoTo errH
    
    '��������ҽ������
    If Not cboDept Is Nothing Then
        If TypeName(cboDept) <> "ComboBox" Then
            Call Err.Raise(ccTypeMismatch, "cPatients.DeptToCbo", "DeptToCbo����Ҫ����ComboBox�Ϳؼ���")
        End If
    End If
    
    If InStr(strPrivs, "���п���") > 0 Then
        strSQL = _
            " Select /*+Rule */ Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B " & _
            " Where B.����ID = A.ID " & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
            " And B.������� IN(1,3) And B.�������� IN('����','�ٴ�')" & _
            " And (A.վ��=[2] Or A.վ�� is Null)" & _
            " Order by A.����"
    Else
        strSQL = _
            " Select /*+Rule */ Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B,������Ա C " & _
            " Where B.����ID = A.ID And A.ID=C.����ID And C.��ԱID=[1]" & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
            " And B.������� IN(1,3) And B.�������� IN('����','�ٴ�')" & _
            " And (A.վ��=[2] Or A.վ�� is Null)" & _
            " Order by A.����"
    End If
    
    cboDept.Clear
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "DeptToCbo", UserInfo.ID, zl9ComLib.gstrNodeNo)
    str����IDs = GetUser����IDs
    For i = 1 To rsTmp.RecordCount
        cboDept.AddItem rsTmp!���� & "-" & rsTmp!����
        cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID
        
        If rsTmp!ID = UserInfo.����ID Then
            Call zlControl.CboSetIndex(cboDept.hwnd, cboDept.NewIndex) 'ֱ����������
        End If
        If InStr("," & str����IDs & ",", "," & rsTmp!ID & ",") > 0 And cboDept.ListIndex = -1 Then
            Call zlControl.CboSetIndex(cboDept.hwnd, cboDept.NewIndex)
        End If
        
        rsTmp.MoveNext
    Next
    If cboDept.ListIndex = -1 And cboDept.ListCount > 0 Then
        Call zlControl.CboSetIndex(cboDept.hwnd, 0)
    End If
    
    
    DeptToCbo = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    
End Function

Private Function GetUser����IDs(Optional ByVal bln���� As Boolean) As String
'���ܣ���ȡ����Ա�����Ŀ���(�������ڿ���+�������������Ŀ���),�����ж��
'�������Ƿ�ȡ���������µĿ���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    If gbln�������Ҷ��� Then
        GetUser����IDs = UserInfo.����ID
    Else
        strSQL = "Select ����ID From ������Ա Where ��ԱID=[1]"
        If bln���� Then
            strSQL = strSQL & " Union" & _
                " Select Distinct B.����ID From ������Ա A,��λ״����¼ B" & _
                " Where A.����ID=B.����ID And A.��ԱID=[1]"
        End If
        
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlTransfusion", UserInfo.ID)
        For i = 1 To rsTmp.RecordCount
            GetUser����IDs = GetUser����IDs & "," & rsTmp!����ID
            rsTmp.MoveNext
        Next
        GetUser����IDs = Mid(GetUser����IDs, 2)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub PatiListRefresh(ByRef Rpt As Object, Optional ByVal strStat As String = "")
    '���ݵ�ǰ���ݣ����RPT�ؼ�
    'rpt ����ؼ�
    
    'strStat : ������һ��״̬�����ݵ��ؼ�  �ԣ��ָ���״̬��־
    
    Dim objPati As cPatient, blnToRpt As Boolean
    Dim intLoop As Integer
    
    If Not Rpt Is Nothing Then
        If TypeName(Rpt) = "ReportControl" Then
            blnToRpt = True
        End If
    End If
    If Not blnToRpt Then Exit Sub
    
    If strStat <> "" Then
        If Right(strStat, 1) <> "," Then strStat = strStat & ","
        If Left(strStat, 1) <> "," Then strStat = "," & strStat
    End If
    
    Call Rpt.Records.DeleteAll
    Rpt.Populate
    For intLoop = 1 To Me.Count
        blnToRpt = True
        Set objPati = Me.Item(intLoop)
        If strStat <> "" Then
            If InStr(strStat, "," & Val(objPati.�Ŷ�״̬) & ",") <= 0 Then blnToRpt = False
        End If
        If blnToRpt Then Call AddToRpt(objPati, Rpt)
    Next
    Rpt.Populate
End Sub
