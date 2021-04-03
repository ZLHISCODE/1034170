Attribute VB_Name = "mdlPivaNormal"
Option Explicit


Public Function PIVA_GetAdvice(ByVal bln�˲� As Boolean, ByVal lngCenterID As Long, ByVal str����id As String, ByVal dateExeStart As Date, ByVal dateExeEnd As Date) As ADODB.Recordset
    'ȡ�����ѷ��͵���δ��ҩ��ҽ����¼��Ҫ�ų��ѷֽ�Ϊ��Һ���ݵļ�¼
    'lngCenterID����Һ��������ID
    'dateExeStart��dateExeEnd��ҽ���Ŀ�ʼִ��ʱ�䷶Χ
    'ע�⣺���ص���ҽ����ҽ�����ݣ�ҩƷ����¼����ȡҽ����ϸʱ��Ҫ�����ID����ͺϲ�
    On Error GoTo errHandle
    gstrSQL = "Select /*+ Rule*/ Distinct A.ID As ҽ��id, Nvl(A.���id, A.ID) As ���id, M.���ͺ�, E.�ⷿid, D.����, D.�Ա�, D.����, D.��ʶ�� As סԺ��, D.����, D.���˲���id, " & _
        " D.���˿���id, A.��ʼִ��ʱ��,  M.����ʱ��, M.������, B.���� As ���˲���, C.���� As ���˿���, E.ID As �շ�id, E.����, E.NO, F.���� As ҩƷ����, F.���� As ͨ����, " & _
        " H.���� As ��Ʒ��, I.���� As Ӣ����, F.���, E.����, E.����, E.����, J.���㵥λ As ������λ, E.Ƶ��, " & _
        " (E.ʵ������ / G.סԺ��װ) As ����, G.סԺ��λ As ��λ , E.����, Decode(A.ҽ����Ч, 0, '����', '��ʱ') As ҽ������, A.ִ��Ƶ��," & _
        " K.���� As ��������, A.����ҽ��, A.ҽ������, A.����ʱ��, A.У�Ի�ʿ, A.У��ʱ��, Nvl(A.�����,-1) �����, E.�÷�, E.ҩƷid " & _
        " From ����ҽ����¼ A, ����ҽ������ M, ���ű� B, ���ű� C, סԺ���ü�¼ D, ҩƷ�շ���¼ E, �շ���ĿĿ¼ F, ҩƷ��� G, �շ���Ŀ���� H, ������Ŀ���� I, ������ĿĿ¼ J, ���ű� K "
        
    If str����id <> "" Then
        gstrSQL = gstrSQL & ",Table(Cast(f_Num2List([2]) As zlTools.t_NumList)) L "
    End If
    
    gstrSQL = gstrSQL & " Where D.���˲���id = B.ID And A.ID = M.ҽ��id And M.NO = D.NO And D.���˿���id = C.ID And A.��������id = K.ID And A.ID = D.ҽ����� And D.ID = E.����id And E.ҩƷid = F.ID And F.ID = G.ҩƷid And " & _
        " G.ҩƷid = H.�շ�ϸĿid(+) And H.����(+) = 3 And G.ҩ��id = I.������Ŀid(+) And I.����(+) = 2 And G.ҩ��id = J.ID And " & _
        " E.������� Is Null And E.ʵ������ > 0 And A.�������� > 0 And Not Exists (Select 1 From ��Һ��ҩ���� Where �շ�id = E.ID And Rownum = 1) " & _
        " And E.�ⷿid = [1] And M.����ʱ�� Between [3] And [4]  " & _
        " And Exists (Select 1 From ������ĿĿ¼ N, ����ҽ����¼ O " & _
        " Where N.��� = 'E' And N.�������� = '2' And N.ִ�з��� = 1 And O.������Ŀid = N.ID And Nvl(A.���id, A.ID) = O.ID) "
        
    If str����id <> "" Then
        gstrSQL = gstrSQL & " And D.���˲���id + 0 =L.Column_Value "
    End If
    
    If bln�˲� = True Then
        gstrSQL = Replace(gstrSQL, "Not Exists", "Exists")
    End If
    
    Set PIVA_GetAdvice = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ����¼", lngCenterID, str����id, dateExeStart, dateExeEnd)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function PIVA_GetExcStatus(ByVal str��ҩids As String, ByVal intStatus As Integer) As ADODB.Recordset
    '��鲻���ϵ�ǰ״̬����Һ��
    'str��ҩids����Һ��ID��
    'intStatus����ǰӦ�õ�ҵ��״̬
    Dim i As Integer
    Dim arrExecute As Variant
    
    On Error GoTo errHandle
    arrExecute = GetArrayByStr(str��ҩids, 3950, ",")
    For i = 0 To UBound(arrExecute)
        gstrSQL = " Select ID, ƿǩ��, ����״̬,�Ƿ��� " & _
            " From ��Һ��ҩ��¼ Where (����״̬ <> [2] " & IIf(intStatus = 2, " or �Ƿ���<>0", "") & ") And ID In (Select * From Table(Cast(f_Num2list([1]) As Zltools.t_Numlist))) "
        Set PIVA_GetExcStatus = zlDatabase.OpenSQLRecord(gstrSQL, "PIVA_GetStatus", CStr(arrExecute(i)), intStatus)
    Next
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Public Function PIVA_GetAdviceCount(ByVal lngCenterID As Long, ByVal dateExeStart As Date, ByVal dateExeEnd As Date) As ADODB.Recordset
    'ȡ�����ѷ��͵���δ��ҩ��ҽ����¼����Ҫ�ų��ѷֽ�Ϊ��Һ���ݵļ�¼
    'lngCenterID����Һ��������ID
    'dateExeStart��dateExeEnd��ҽ����ִ��ʱ�䷶Χ������ʱ�䣩
    Dim strTmp As String
    
    On Error GoTo errHandle
    gstrSQL = "Select /*+ rule*/ ����id, ����, Count(����id) As ����, 0 �˲��־ " & _
        " From (Select  Distinct B.���˲���id As ����id, '[' || D.���� || ']' || D.���� As ����, Nvl(A.���id, A.ID) As ���id, E.���ͺ� " & _
        " From ����ҽ����¼ A, ����ҽ������ E, סԺ���ü�¼ B, ҩƷ�շ���¼ C, ���ű� D " & _
        " Where A.Id = B.ҽ����� And A.ID = E.ҽ��id And E.NO = B.NO And B.Id = C.����id And B.���˲���id = D.Id And C.������� Is Null And " & _
        " C.ʵ������ > 0 And A.�������� > 0 And C.�ⷿid = [1] And Not Exists " & _
        " (Select 1 From ��Һ��ҩ���� Where �շ�id = C.ID And Rownum = 1) And E.����ʱ�� Between [2] And [3] " & _
        " And Exists (Select 1 From ������ĿĿ¼ F, ����ҽ����¼ G " & _
        " Where F.��� = 'E' And F.�������� = '2' And F.ִ�з��� = 1 And G.������Ŀid = F.ID And Nvl(A.���id, A.ID) = G.ID)" & _
        " Order By '[' || D.���� || ']' || D.����) " & _
        " Group By ����id, ����"
    
    strTmp = Replace(gstrSQL, "0 �˲��־", "1 �˲��־")
    strTmp = Replace(strTmp, "Not Exists", "Exists")
    
    '�ϲ�δ�˲���Ѻ˲��ҽ��
    gstrSQL = gstrSQL & " Union All " & strTmp
    
    Set PIVA_GetAdviceCount = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ҽ����¼��", lngCenterID, dateExeStart, dateExeEnd)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function Piva_GetMedi(ByVal lngCenterID As Long, ByVal str����id As String, ByVal intStemp As Integer, ByVal dateExeStart As Date, _
        ByVal dateExeEnd As Date, ByVal int��ʾ���� As Integer) As Recordset
    
    On Error GoTo errHandle
    If int��ʾ���� = 0 Then
        gstrSQL = "Select Distinct a.id,a.���Id,A.����id,A.��ҳid,A.����ҽ��,A.�����,A.ҩʦ���ԭ��,H.���˲���ID,H.���˿���ID, b.���� ��������, f.��ǰ���� ����,P.��ҩ����,decode(A.ҽ����Ч,0,'����',1,'��ʱ') ҽ����Ч,M.���� ��ҩ;��, " & _
            " g.��ʶ�� As סԺ��, a.����, a.�Ա�, a.����,c.id ҩƷid,  c.���� ҩƷ����, c.���, a.��������,I.���㵥λ,I.id ҩ��id,a.ִ��Ƶ��,nvl(a.ҩʦ��˱�־,0) ��˱�־,a.ִ��ʱ�䷽��,A.Ƥ�Խ��,A.����ʱ��,nvl(T.�Ƿ�Ƥ��,0) �Ƿ�Ƥ�� " & vbNewLine & _
            "From ����ҽ����¼ A, ���ű� B, �շ���ĿĿ¼ C, ��Һ��ҩ���� D, ҩƷ�շ���¼ E, ������Ϣ F, סԺ���ü�¼ G,��Һ��ҩ��¼ H,������ĿĿ¼ I,ҩƷ��� J, ҩƷ���� T,��ҺҩƷ���� P,����ҽ����¼ L,������ĿĿ¼ M " & vbNewLine & _
             ",Table(Cast(f_Num2List([2]) As zlTools.t_NumList)) K " & vbNewLine & _
            "Where a.����id = f.����id  And e.����id = g.Id And a.���˿���id = b.Id And E.����=9 and a.Id = g.ҽ����� And g.�շ�ϸĿid = c.Id and J.ҩƷid=c.id and J.ҩƷid=P.ҩƷid and J.ҩ��id=I.id And A.���id=L.id and L.������Ŀid=M.id And H.id=D.��¼id And J.ҩ��id=T.ҩ��id " & vbNewLine & _
            "      And e.Id = d.�շ�id And (nvl(a.ҩʦ��˱�־,0)=[1] " & IIf(intStemp = 0, " or nvl(a.ҩʦ��˱�־,0)=3", "") & ") " & IIf(intStemp = 0, " And H.����״̬ = 1 ", "") & "  And H.���˲���id + 0 =K.Column_Value  and H.ִ��ʱ�� between [3] and [4] And h.����id=[5] " & vbNewLine & _
            " order by b.����,A.����id,a.���Id"
            
        Set Piva_GetMedi = zlDatabase.OpenSQLRecord(gstrSQL, "Piva_GetMedi", intStemp, str����id, dateExeStart, dateExeEnd, lngCenterID)
    Else
        If intStemp = 0 Then
            gstrSQL = "Select Distinct a.id,a.���Id,A.����id,A.��ҳid,A.����ҽ��,A.�����,A.ҩʦ���ԭ��,G.���˲���ID,G.���˿���ID, b.���� ��������, f.��ǰ���� ����,P.��ҩ����,decode(A.ҽ����Ч,0,'����',1,'��ʱ') ҽ����Ч,M.���� ��ҩ;��, " & _
                " g.��ʶ�� As סԺ��, a.����, a.�Ա�, a.����,c.id ҩƷid,  c.���� ҩƷ����, c.���, a.��������,I.���㵥λ,I.id ҩ��id,a.ִ��Ƶ��,nvl(a.ҩʦ��˱�־,0) ��˱�־,a.ִ��ʱ�䷽��,A.Ƥ�Խ��,A.����ʱ��,nvl(T.�Ƿ�Ƥ��,0) �Ƿ�Ƥ�� " & vbNewLine & _
                "From ����ҽ����¼ A, ���ű� B, �շ���ĿĿ¼ C,  ҩƷ�շ���¼ E, ������Ϣ F, סԺ���ü�¼ G,������ĿĿ¼ I,ҩƷ��� J, ҩƷ���� T,��ҺҩƷ���� P,����ҽ����¼ L,������ĿĿ¼ M " & vbNewLine & _
                ",Table(Cast(f_Num2List([4]) As zlTools.t_NumList)) K " & vbNewLine & _
                "Where a.����id = f.����id  And e.����id = g.Id And a.���˿���id = b.Id And E.����=9 and a.Id = g.ҽ����� And g.�շ�ϸĿid = c.Id and J.ҩƷid=c.id and J.ҩƷid=P.ҩƷid and J.ҩ��id=I.id And A.���id=L.id and L.������Ŀid=M.id And J.ҩ��id=T.ҩ��id " & vbNewLine & _
                "      And G.���˲���id  =K.Column_Value  and E.�������� between [1] and [2] And E.�ⷿid=[3] " & vbNewLine & _
                " order by b.����,A.����id,a.���Id"
            Set Piva_GetMedi = zlDatabase.OpenSQLRecord(gstrSQL, "Piva_GetMedi", dateExeStart, dateExeEnd, lngCenterID, str����id)
        Else
            gstrSQL = "Select Distinct a.id,a.���Id,A.����id,A.��ҳid,A.����ҽ��,A.�����,A.ҩʦ���ԭ��,G.���˲���ID,G.���˿���ID, b.���� ��������, f.��ǰ���� ����,P.��ҩ����,decode(A.ҽ����Ч,0,'����',1,'��ʱ') ҽ����Ч,M.���� ��ҩ;��, " & _
                " g.��ʶ�� As סԺ��, a.����, a.�Ա�, a.����,c.id ҩƷid,  c.���� ҩƷ����, c.���, a.��������,I.���㵥λ,I.id ҩ��id,a.ִ��Ƶ��,nvl(a.ҩʦ��˱�־,0) ��˱�־,a.ִ��ʱ�䷽��,A.Ƥ�Խ��,A.����ʱ��,nvl(T.�Ƿ�Ƥ��,0) �Ƿ�Ƥ�� " & vbNewLine & _
                "From ����ҽ����¼ A, ���ű� B, �շ���ĿĿ¼ C,  ҩƷ�շ���¼ E, ������Ϣ F, סԺ���ü�¼ G,������ĿĿ¼ I,ҩƷ��� J, ҩƷ���� T,��ҺҩƷ���� P,����ҽ����¼ L,������ĿĿ¼ M " & vbNewLine & _
                ",Table(Cast(f_Num2List([4]) As zlTools.t_NumList)) K " & vbNewLine & _
                "Where a.����id = f.����id  And e.����id = g.Id And a.���˿���id = b.Id And E.����=9 and a.Id = g.ҽ����� And g.�շ�ϸĿid = c.Id and J.ҩƷid=c.id and J.ҩƷid=P.ҩƷid and J.ҩ��id=I.id And A.���id=L.id and L.������Ŀid=M.id And J.ҩ��id=T.ҩ��id " & vbNewLine & _
                "      And G.���˲���id  =K.Column_Value  and E.�������� between [1] and [2] And E.�ⷿid=[3] And (nvl(a.ҩʦ��˱�־,0)=[5] " & IIf(intStemp = 0, " or nvl(a.ҩʦ��˱�־,0)=3", "") & ") " & vbNewLine & _
                " order by b.����,A.����id,a.���Id"
            Set Piva_GetMedi = zlDatabase.OpenSQLRecord(gstrSQL, "Piva_GetMedi", dateExeStart, dateExeEnd, lngCenterID, str����id, intStemp)
        End If
    End If
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Public Function Piva_GetTrans(ByVal lngCenterID As Long, ByVal lng����id As Long, ByVal dateExeStart As Date, _
        ByVal dateExeEnd As Date, ByVal strStep As String, ByVal intPack As Integer, ByVal intSend As Integer, ByVal bln��� As Boolean, ByVal bln������ As Boolean) As ADODB.Recordset
        
    'ȡ��Һ��ҩ��¼
    'lngCenterID����Һ��������ID
    'str����ID������ID��
    'dateExeStart��dateExeEnd����Һ��ҩ���ݵ�ִ��ʱ�䷶Χ
    'strStep(��������)��01-��ҩӡǩ(1)��02-��ҩ�˲�(2)��03-���ͺ˲�(4)��04-�������(9)��10-�����ͨ��ҽ��(10)��11-���δͨ��ҽ��(10)��12-�ѷ��Ͳ鿴(5), 13-��ǩ�ղ鿴(6)��14-�ܾ�ǩ�ղ�(7)��15-�����ϲ鿴
    '�������ͣ�1�����ƣ�2����ҩ��3��У�ԣ�4����ҩ��5�����ͣ�6��ǩ�գ�7���ܾ�ǩ��  8��ȷ�Ͼ��գ�9���������룬10���������
    'intPack���룺0-���У�1-����ҩ��2-�����
    '�Ƿ�����0-�����(��Һ),1-�������,2-�������Ĵ��
    On Error GoTo errHandle
    
    If strStep = "15" Then
        '�Ѱ�ҩ״̬
        '1.�������ͨ��
        gstrSQL = "Select Distinct A.ID As ��ҩID,A.���α��,A.���ȼ�,A.�Ƿ�ȷ�ϵ���, A.����id, A.���, A.��ҩ����,S.��ɫ, A.����, A.�Ա�, A.����, A.סԺ��,A.����,LPad(A.����, 10, ' ') ��������,P.����,M.��� ҽ�����,M.ҩʦ���ʱ��,M.ִ��Ƶ��, A.���˲���id, A.���˿���id, A.ִ��ʱ��, A.ƿǩ��,A.���ʱ��,M.����id,M.��ҳid,A.�Ƿ��������,A.�Ƿ�����,A.�ֹ���������,'' ����ԭ��," & _
            " A.������Ա,A.����ʱ��, Nvl(A.��ӡ��־,0) As ��ӡ��־, A.�Ƿ���, B.���� As ���˲���, C.���� As ���˿���, 0 As �շ�id, 9 As ����, '' NO, F.���� As ҩƷ����, " & _
            " F.���� As ͨ����, H.���� As ��Ʒ��, I.���� As Ӣ����, F.���, e.����, e.����, M.�������� As ����, J.���㵥λ As ������λ,J.id ҩ��id, e.Ƶ��, '�������ͨ��' As ��������, " & _
            " 0 As ��ҩ����, (e.���ϵ��*e.ʵ������ / G.סԺ��װ) As ����,e.���ϵ��*e.ʵ������ As ʵ������, G.סԺ��λ As ��λ,0 As ����, 0 As �������, Nvl(M.�����,-1) �����, e.�÷�, e.ҩƷid,0 as �������,0 As ����id,null As ����, A.��ҩ����,L.����ʱ�� As ҽ������ʱ��,nvl(T.������,'0') ��ҩ����,T.��ý,M.Ƥ�Խ��,M.����ʱ��,A.ҽ��id,A.���ͺ�,nvl(T.�Ƿ�Ƥ��,0) �Ƿ�Ƥ��,x.��ҩ���� As ��ҩ����1 " & _
            " From ��Һ��ҩ��¼ A, ���ű� B, ���ű� C, �շ���ĿĿ¼ F, ҩƷ��� G,��ҺҩƷ���� X, �շ���Ŀ���� H, ������Ŀ���� I, ������ĿĿ¼ J, ����ҽ����¼ M, סԺ���ü�¼ D, ҩƷ�շ���¼ E, ����ҽ������ L ,��ҩ�������� S,ҩƷ���� T,��λ״����¼ O,��λ���Ʒ��� P,��Һ��ҩ���� Z "

        gstrSQL = gstrSQL & " Where A.ҽ��id = M.���id And A.���˲���id = B.ID And A.���˿���id = C.ID And F.ID = G.ҩƷid And G.ҩƷid=X.ҩƷid(+) And g.ҩƷid = e.ҩƷid And T.ҩ��id=J.id And A.����=O.����(+) And  A.���˲���id=O.����id(+) And A.���˿���id=O.����id(+) and O.��λ����=P.����(+) And " & _
            " G.ҩƷid = H.�շ�ϸĿid(+) And H.����(+) = 3 And A.��ҩ����=S.����(+) And a.����id = s.��������id(+) And G.ҩ��id = I.������Ŀid(+) And I.����(+) = 2 And G.ҩ��id = J.ID " & _
            " And m.Id = d.ҽ����� And d.Id = e.����id And a.ҽ��id = l.ҽ��id(+) And a.���ͺ� = l.���ͺ�(+) And a.id = z.��¼id And z.�շ�id = e.id " & _
            " And a.����״̬=10 And A.����id = [1] And A.ִ��ʱ�� Between [3] And [4] And Exists (Select 1 From ��Һ��ҩ���� D, ҩƷ�շ���¼ E Where d.�շ�id = e.Id And d.��¼id = a.Id)"
            
        If lng����id <> 0 Then
            gstrSQL = gstrSQL & " And A.���˲���id + 0 =[2] "
        End If

        If intPack = 1 Then
            '�����
            gstrSQL = gstrSQL & " And Nvl(A.�Ƿ���,0)=0 "
        ElseIf intPack = 2 Then
            '�������������������������Ĵ��
            gstrSQL = gstrSQL & " And Nvl(A.�Ƿ���,0) In (1,2) "
        End If
        
        '2.������˾ܾ�
        gstrSQL = gstrSQL & " Union All " & _
            "Select Distinct A.ID As ��ҩID,A.���α��,A.���ȼ�,A.�Ƿ�ȷ�ϵ���, A.����id, A.���, A.��ҩ����,S.��ɫ, A.����, A.�Ա�, A.����, A.סԺ��,A.����,LPad(A.����, 10, ' ') ��������,P.����,M.��� ҽ�����,M.ҩʦ���ʱ��,M.ִ��Ƶ��, A.���˲���id, A.���˿���id, A.ִ��ʱ��, A.ƿǩ��,A.���ʱ��,M.����id,M.��ҳid,A.�Ƿ��������,A.�Ƿ�����,A.�ֹ���������,'' ����ԭ��," & _
            " A.������Ա,A.����ʱ��, Nvl(A.��ӡ��־,0) As ��ӡ��־, A.�Ƿ���, B.���� As ���˲���, C.���� As ���˿���, 0 As �շ�id, 9 As ����, '' NO, F.���� As ҩƷ����, " & _
            " F.���� As ͨ����, H.���� As ��Ʒ��, I.���� As Ӣ����, F.���, e.����, e.����, M.�������� As ����, J.���㵥λ As ������λ,J.id ҩ��id, e.Ƶ��, '������˾ܾ�' As ��������, " & _
            " 0 As ��ҩ����, (e.���ϵ��*e.ʵ������ / G.סԺ��װ) As ����,e.���ϵ��*e.ʵ������ As ʵ������, G.סԺ��λ As ��λ,0 As ����, 0 As �������, Nvl(M.�����,-1) �����, e.�÷�, e.ҩƷid,0 as �������,0 As ����id,null As ����, A.��ҩ����,L.����ʱ�� As ҽ������ʱ��,nvl(T.������,'0') ��ҩ����,T.��ý,M.Ƥ�Խ��,M.����ʱ��,A.ҽ��id,A.���ͺ�,nvl(T.�Ƿ�Ƥ��,0) �Ƿ�Ƥ��,x.��ҩ���� As ��ҩ����1 " & _
            " From ��Һ��ҩ��¼ A, ���ű� B, ���ű� C, �շ���ĿĿ¼ F, ҩƷ��� G,��ҺҩƷ���� X, �շ���Ŀ���� H, ������Ŀ���� I, ������ĿĿ¼ J, ����ҽ����¼ M, סԺ���ü�¼ D, ҩƷ�շ���¼ E, ����ҽ������ L ,��ҩ�������� S,ҩƷ���� T,��λ״����¼ O,��λ���Ʒ��� P,��Һ��ҩ���� Z "

        gstrSQL = gstrSQL & " Where A.ҽ��id = M.���id And A.���˲���id = B.ID And A.���˿���id = C.ID And F.ID = G.ҩƷid And G.ҩƷid=X.ҩƷid(+) And g.ҩƷid = e.ҩƷid And T.ҩ��id=J.id And A.����=O.����(+) And  A.���˲���id=O.����id(+) And A.���˿���id=O.����id(+) and O.��λ����=P.����(+) And " & _
            " G.ҩƷid = H.�շ�ϸĿid(+) And H.����(+) = 3 And A.��ҩ����=S.����(+) And a.����id = s.��������id(+) And G.ҩ��id = I.������Ŀid(+) And I.����(+) = 2 And G.ҩ��id = J.ID " & _
            " And m.Id = d.ҽ����� And d.Id = e.����id And a.ҽ��id = l.ҽ��id(+) And a.���ͺ� = l.���ͺ�(+)  and e.ʵ������>0 And a.id = z.��¼id And z.�շ�id = e.id " & _
            " And a.����״̬=11 And A.����id = [1] And A.ִ��ʱ�� Between [3] And [4] And Exists (Select 1 From ��Һ��ҩ���� D, ҩƷ�շ���¼ E Where d.�շ�id = e.Id And d.��¼id = a.Id)"
            
        If lng����id <> 0 Then
            gstrSQL = gstrSQL & " And A.���˲���id + 0 =[2] "
        End If

        If intPack = 1 Then
            '�����
            gstrSQL = gstrSQL & " And Nvl(A.�Ƿ���,0)=0 "
        ElseIf intPack = 2 Then
            '�������������������������Ĵ��
            gstrSQL = gstrSQL & " And Nvl(A.�Ƿ���,0) In (1,2) "
        End If
        
        'δ��ҩ״̬
        '�����
        gstrSQL = gstrSQL & " Union All " & _
            " Select Distinct A.ID As ��ҩID,A.���α��,A.���ȼ�,A.�Ƿ�ȷ�ϵ���, A.����id, A.���, A.��ҩ����,S.��ɫ, A.����, A.�Ա�, A.����, A.סԺ��,A.����,LPad(A.����, 10, ' ') ��������,P.����,M.��� ҽ�����,M.ҩʦ���ʱ��,M.ִ��Ƶ��,  A.���˲���id, A.���˿���id, A.ִ��ʱ��, A.ƿǩ��,A.���ʱ��,M.����id,M.��ҳid,A.�Ƿ��������,A.�Ƿ�����,A.�ֹ���������,'' ����ԭ��," & _
            " A.������Ա,A.����ʱ��, Nvl(A.��ӡ��־,0) As ��ӡ��־, A.�Ƿ���, B.���� As ���˲���, C.���� As ���˿���, 0 As �շ�id, 9 As ����, '' As NO, F.���� As ҩƷ����, " & _
            " F.���� As ͨ����, H.���� As ��Ʒ��, I.���� As Ӣ����, F.���, '' As ����, '' As ����, M.�������� As ����, J.���㵥λ As ������λ,J.id ҩ��id, '' As Ƶ��, 'δ��ҩ����' As ��������, " & _
            " 0 As ��ҩ����, (M.��������/ G.����ϵ�� / G.סԺ��װ) As ����,M.��������/ G.����ϵ�� As ʵ������, G.סԺ��λ As ��λ,0 As ����, 0 As �������, Nvl(M.�����,-1) �����, '' As �÷�, M.�շ�ϸĿid As ҩƷid,0 as �������,0 As ����id,null As ����, " & _
            " A.��ҩ����,Null As ҽ������ʱ��,nvl(T.������,'0') ��ҩ����,T.��ý,M.Ƥ�Խ��,M.����ʱ��,A.ҽ��id,A.���ͺ�,nvl(T.�Ƿ�Ƥ��,0) �Ƿ�Ƥ��,x.��ҩ���� As ��ҩ����1 " & _
            " From ��Һ��ҩ��¼ A, ���ű� B, ���ű� C, �շ���ĿĿ¼ F, ҩƷ��� G,��ҺҩƷ���� X, �շ���Ŀ���� H, ������Ŀ���� I, ������ĿĿ¼ J, ����ҽ����¼ M ,��ҩ�������� S,ҩƷ���� T,��λ״����¼ O,��λ���Ʒ��� P "
        
        gstrSQL = gstrSQL & " Where A.ҽ��id = M.���id And A.���˲���id = B.ID  And A.���˿���id = C.ID And F.ID = G.ҩƷid And G.ҩƷid=X.ҩƷid(+) And M.�շ�ϸĿid = F.ID And T.ҩ��id=J.id And A.����=O.����(+) And  A.���˲���id=O.����id(+) And A.���˿���id=O.����id(+) and O.��λ����=P.����(+) And " & _
            " G.ҩƷid = H.�շ�ϸĿid(+) And H.����(+) = 3 And A.��ҩ����=S.����(+) And a.����id = s.��������id(+) And G.ҩ��id = I.������Ŀid(+) And I.����(+) = 2 And G.ҩ��id = J.ID And a.����״̬=10  " & _
            " And A.����id = [1] And A.ִ��ʱ�� Between [3] And [4] And Not Exists (Select 1 From ��Һ��ҩ���� D, ҩƷ�շ���¼ E Where d.�շ�id = e.Id And d.��¼id = a.Id) "
            
        If lng����id <> 0 Then
            gstrSQL = gstrSQL & " And A.���˲���id + 0 =[2] "
        End If
        
        If intPack = 1 Then
            '�����
            gstrSQL = gstrSQL & " And Nvl(A.�Ƿ���,0)=0 "
        ElseIf intPack = 2 Then
            '�������������������������Ĵ��
            gstrSQL = gstrSQL & " And Nvl(A.�Ƿ���,0) In (1,2) "
        End If
                
        '��Ʒ��
        gstrSQL = gstrSQL & " Union All " & _
            " Select Distinct A.ID As ��ҩID,A.���α��,A.���ȼ�,A.�Ƿ�ȷ�ϵ���, A.����id, A.���, A.��ҩ����,S.��ɫ, A.����, A.�Ա�, A.����, A.סԺ��,A.����,LPad(A.����, 10, ' ') ��������,P.����,M.��� ҽ�����,M.ҩʦ���ʱ��,M.ִ��Ƶ��,  A.���˲���id, A.���˿���id, A.ִ��ʱ��, A.ƿǩ��,A.���ʱ��,M.����id,M.��ҳid,A.�Ƿ��������,A.�Ƿ�����,A.�ֹ���������,'' ����ԭ��," & _
            " A.������Ա,A.����ʱ��, Nvl(A.��ӡ��־,0) As ��ӡ��־, A.�Ƿ���, B.���� As ���˲���, C.���� As ���˿���, 0 As �շ�id, 9 As ����, '' As NO, J.���� As ҩƷ����, " & _
            " J.���� As ͨ����, '' As ��Ʒ��, I.���� As Ӣ����, '' as ���, '' As ����, '' As ����, M.�������� As ����, J.���㵥λ As ������λ,J.id ҩ��id, '' As Ƶ��, 'δ��ҩ����' As ��������, " & _
            " 0 As ��ҩ����, 0 As ����,0 As ʵ������, '' ��λ,0 As ����, 0 As �������, Nvl(M.�����,-1) �����, '' As �÷�, Decode(Nvl(m.�շ�ϸĿid, 0), 0, j.Id, m.�շ�ϸĿid) As ҩƷid,0 as �������,0 As ����id,null As ����, " & _
            " A.��ҩ����,Null As ҽ������ʱ��,0 ��ҩ����,T.��ý,M.Ƥ�Խ��,M.����ʱ��,A.ҽ��id,A.���ͺ�,nvl(T.�Ƿ�Ƥ��,0) �Ƿ�Ƥ��,'' As ��ҩ����1 " & _
            " From ��Һ��ҩ��¼ A, ���ű� B, ���ű� C,������Ŀ���� I, ������ĿĿ¼ J, ����ҽ����¼ M ,��ҩ�������� S,ҩƷ���� T,��λ״����¼ O,��λ���Ʒ��� P "

        gstrSQL = gstrSQL & " Where A.ҽ��id = M.���id And A.���˲���id = B.ID  And A.���˿���id = C.ID and M.�շ�ϸĿid is null and M.������Ŀid=J.id And A.����=O.����(+) And  A.���˲���id=O.����id(+) And A.���˿���id=O.����id(+) and O.��λ����=P.����(+) And a.����״̬=10  " & _
            " And A.��ҩ����=S.����(+) And a.����id = s.��������id(+) And J.id = I.������Ŀid(+) And I.����(+) = 2 And j.Id = t.ҩ��id " & _
            " And A.����id = [1] And A.ִ��ʱ�� Between [3] And [4] And Not Exists (Select 1 From ��Һ��ҩ���� D, ҩƷ�շ���¼ E Where d.�շ�id = e.Id And d.��¼id = a.Id) "

        If lng����id <> 0 Then
            gstrSQL = gstrSQL & " And A.���˲���id + 0 =[2] "
        End If

        If intPack = 1 Then
            '�����
            gstrSQL = gstrSQL & " And Nvl(A.�Ƿ���,0)=0 "
        ElseIf intPack = 2 Then
            '�������������������������Ĵ��
            gstrSQL = gstrSQL & " And Nvl(A.�Ƿ���,0) In (1,2) "
        End If
    ElseIf strStep = "16" Then
        'ҽ������(�����)
        gstrSQL = " Select Distinct A.ID As ��ҩID,A.���α��,A.���ȼ�,A.�Ƿ�ȷ�ϵ���, A.����id, A.���, A.��ҩ����,S.��ɫ, A.����, A.�Ա�, A.����, A.סԺ��,A.����,LPad(A.����, 10, ' ') ��������,P.����,M.��� ҽ�����,M.ҩʦ���ʱ��,M.ִ��Ƶ��,  A.���˲���id, A.���˿���id, A.ִ��ʱ��, A.ƿǩ��,A.���ʱ��,M.����id,M.��ҳid,A.�Ƿ��������,A.�Ƿ�����,A.�ֹ���������,'' ����ԭ��," & _
            " A.������Ա,A.����ʱ��, Nvl(A.��ӡ��־,0) As ��ӡ��־, A.�Ƿ���, B.���� As ���˲���, C.���� As ���˿���, 0 As �շ�id, 9 As ����, '' As NO, F.���� As ҩƷ����, " & _
            " F.���� As ͨ����, H.���� As ��Ʒ��, I.���� As Ӣ����, F.���, '' As ����, '' As ����, M.�������� As ����, J.���㵥λ As ������λ,J.id ҩ��id, '' As Ƶ��, 'ҽ������' As ��������, " & _
            " 0 As ��ҩ����, (M.��������/ G.����ϵ�� / G.סԺ��װ) As ����,M.��������/ G.����ϵ�� As ʵ������, G.סԺ��λ As ��λ,0 As ����, 0 As �������, Nvl(M.�����,-1) �����, '' As �÷�, M.�շ�ϸĿid As ҩƷid,0 as �������,0 As ����id,null As ����, " & _
            " A.��ҩ����,Null As ҽ������ʱ��,nvl(T.������,'0') ��ҩ����,T.��ý,M.Ƥ�Խ��,M.����ʱ��,A.ҽ��id,A.���ͺ�,nvl(T.�Ƿ�Ƥ��,0) �Ƿ�Ƥ��,x.��ҩ���� As ��ҩ����1 " & _
            " From ��Һ��ҩ��¼ A, ���ű� B, ���ű� C, �շ���ĿĿ¼ F, ҩƷ��� G,��ҺҩƷ���� X, �շ���Ŀ���� H, ������Ŀ���� I, ������ĿĿ¼ J, ����ҽ����¼ M ,��ҩ�������� S,ҩƷ���� T,��λ״����¼ O,��λ���Ʒ��� P "
        
        gstrSQL = gstrSQL & " Where A.ҽ��id = M.���id And A.���˲���id = B.ID  And A.���˿���id = C.ID And F.ID = G.ҩƷid And G.ҩƷid=X.ҩƷid(+) And M.�շ�ϸĿid = F.ID And T.ҩ��id=J.id And A.����=O.����(+) And  A.���˲���id=O.����id(+) And A.���˿���id=O.����id(+) and O.��λ����=P.����(+) And " & _
            " G.ҩƷid = H.�շ�ϸĿid(+) And H.����(+) = 3 And A.��ҩ����=S.����(+) And a.����id = s.��������id(+) And G.ҩ��id = I.������Ŀid(+) And I.����(+) = 2 And G.ҩ��id = J.ID And a.����״̬=12  " & _
            " And A.����id = [1] And A.ִ��ʱ�� Between [3] And [4] "
            
        If lng����id <> 0 Then
            gstrSQL = gstrSQL & " And A.���˲���id + 0 =[2] "
        End If
        
        If intPack = 1 Then
            '�����
            gstrSQL = gstrSQL & " And Nvl(A.�Ƿ���,0)=0 "
        ElseIf intPack = 2 Then
            '�������������������������Ĵ��
            gstrSQL = gstrSQL & " And Nvl(A.�Ƿ���,0) In (1,2) "
        End If
                
        '�ϲ�ҽ������(��Ʒ�ַ���)
        gstrSQL = gstrSQL & " Union All " & _
            " Select Distinct A.ID As ��ҩID,A.���α��,A.���ȼ�,A.�Ƿ�ȷ�ϵ���, A.����id, A.���, A.��ҩ����,S.��ɫ, A.����, A.�Ա�, A.����, A.סԺ��,A.����,LPad(A.����, 10, ' ') ��������,P.����,M.��� ҽ�����,M.ҩʦ���ʱ��,M.ִ��Ƶ��,  A.���˲���id, A.���˿���id, A.ִ��ʱ��, A.ƿǩ��,A.���ʱ��,M.����id,M.��ҳid,A.�Ƿ��������,A.�Ƿ�����,A.�ֹ���������,'' ����ԭ��," & _
            " A.������Ա,A.����ʱ��, Nvl(A.��ӡ��־,0) As ��ӡ��־, A.�Ƿ���, B.���� As ���˲���, C.���� As ���˿���, 0 As �շ�id, 9 As ����, '' As NO, J.���� As ҩƷ����, " & _
            " J.���� As ͨ����, '' As ��Ʒ��, I.���� As Ӣ����, '' as ���, '' As ����, '' As ����, M.�������� As ����, J.���㵥λ As ������λ,J.id ҩ��id, '' As Ƶ��, 'ҽ������' As ��������, " & _
            " 0 As ��ҩ����, 0 As ����,0 As ʵ������, '' ��λ,0 As ����, 0 As �������, Nvl(M.�����,-1) �����, '' As �÷�, Decode(Nvl(m.�շ�ϸĿid, 0), 0, j.Id, m.�շ�ϸĿid) As ҩƷid,0 as �������,0 As ����id,null As ����, " & _
            " A.��ҩ����,Null As ҽ������ʱ��,0 ��ҩ����,T.��ý,M.Ƥ�Խ��,M.����ʱ��,A.ҽ��id,A.���ͺ�,nvl(T.�Ƿ�Ƥ��,0) �Ƿ�Ƥ��,'' As ��ҩ����1 " & _
            " From ��Һ��ҩ��¼ A, ���ű� B, ���ű� C,������Ŀ���� I, ������ĿĿ¼ J, ����ҽ����¼ M ,��ҩ�������� S,ҩƷ���� T,��λ״����¼ O,��λ���Ʒ��� P "

        gstrSQL = gstrSQL & " Where A.ҽ��id = M.���id And A.���˲���id = B.ID  And A.���˿���id = C.ID and M.�շ�ϸĿid is null and M.������Ŀid=J.id And A.����=O.����(+) And  A.���˲���id=O.����id(+) And A.���˿���id=O.����id(+) and O.��λ����=P.����(+) And a.����״̬=12  " & _
            " And A.��ҩ����=S.����(+) And a.����id = s.��������id(+) And J.id = I.������Ŀid(+) And I.����(+) = 2 And j.Id = t.ҩ��id " & _
            " And A.����id = [1] And A.ִ��ʱ�� Between [3] And [4] "

        If lng����id <> 0 Then
            gstrSQL = gstrSQL & " And A.���˲���id + 0 =[2] "
        End If

        If intPack = 1 Then
            '�����
            gstrSQL = gstrSQL & " And Nvl(A.�Ƿ���,0)=0 "
        ElseIf intPack = 2 Then
            '�������������������������Ĵ��
            gstrSQL = gstrSQL & " And Nvl(A.�Ƿ���,0) In (1,2) "
        End If
    Else
        '����
        gstrSQL = "Select Distinct A.ID As ��ҩID,A.���α��,A.���ȼ�,A.�Ƿ�ȷ�ϵ���, A.����id, A.���, A.��ҩ����, S.��ɫ,A.����, A.�Ա�, A.����, A.סԺ��,A.����,LPad(A.����, 10, ' ') ��������,P.����,M.��� ҽ�����,M.ҩʦ���ʱ��,M.ִ��Ƶ��,  A.���˲���id, A.���˿���id, A.ִ��ʱ��, A.ƿǩ��,A.���ʱ��,A.�Ƿ��������,A.�Ƿ�����,A.�ֹ���������," & IIf(strStep = "13", "W.����˵�� ����ԭ��,", "'' ����ԭ��,") & _
            "  A.������Ա,A.����ʱ��,Nvl(A.��ӡ��־,0) As ��ӡ��־, A.�Ƿ���, B.���� As ���˲���, C.���� As ���˿���, D.�շ�id, E.����, E.NO, F.���� As ҩƷ����, " & _
            " F.���� As ͨ����, H.���� As ��Ʒ��, I.���� As Ӣ����, F.���, E.����, E.����, E.����, J.���㵥λ As ������λ,J.id ҩ��id, E.Ƶ��, '' As ��������, " & _
            " Case Nvl(E.�����, 'δ���') When 'δ���' Then E.ʵ������ * Nvl(E.����, 1) / G.סԺ��װ Else 0 End As ��ҩ����,M.����id,M.��ҳid,T.��ý,M.Ƥ�Խ��,M.����ʱ��,A.ҽ��id,A.���ͺ�, " & _
            " (D.���� / G.סԺ��װ)  As ����,D.���� As ʵ������, G.סԺ��λ As ��λ,Nvl(E.����,0) As ����, Nvl(L.ʵ������, 0)/ G.סԺ��װ As �������, Nvl(M.�����,-1) �����, E.�÷�, E.ҩƷid, n.��� As �������,E.����id, o.����, A.��ҩ����,r.����ʱ�� As ҽ������ʱ��,nvl(T.������,'0') ��ҩ����,nvl(T.�Ƿ�Ƥ��,0) �Ƿ�Ƥ��,x.��ҩ���� As ��ҩ����1 " & _
            " From  ��Һ��ҩ��¼ A, ���ű� B, ���ű� C, ��Һ��ҩ���� D, ҩƷ�շ���¼ E, �շ���ĿĿ¼ F, ҩƷ��� G,��ҺҩƷ���� X,  �շ���Ŀ���� H, ������Ŀ���� I, ������ĿĿ¼ J, ����ҽ����¼ M, סԺ���ü�¼ N, ������ҳ O ,��ҩ�������� S,ҩƷ���� T,��λ״����¼ Q,��λ���Ʒ��� P "
        
        If strStep = "13" Then gstrSQL = gstrSQL & ",��Һ��ҩ״̬ W "
        
        If strStep = "01" And bln������ Then
            gstrSQL = gstrSQL & ",��������¼ Q,���������ϸ K "
        End If
        
        gstrSQL = gstrSQL & ",(Select �ⷿid, ҩƷid, Nvl(����, 0) As ����, Nvl(ʵ������, 0) As ʵ������ " & _
            " From ҩƷ��� Where ���� = 1 And �ⷿid = [1]) L, ҩƷ�շ���¼ P, ����ҽ������ R "
        
        gstrSQL = gstrSQL & " Where A.���˲���id = B.ID And A.���˿���id = C.ID And A.ID = D.��¼id And D.�շ�id = E.ID And E.ҩƷid = F.ID And F.ID = G.ҩƷid And G.ҩƷid=X.ҩƷid(+) And E.����id = N.ID And N.ҽ����� = M.ID And " & IIf(strStep = "13", "W.��ҩid=A.id And A.����״̬=W.�������� And A.����ʱ��=W.����ʱ�� And ", "") & _
            " G.ҩƷid = H.�շ�ϸĿid(+) And H.����(+) = 3 And G.ҩ��id = I.������Ŀid(+) And I.����(+) = 2 And G.ҩ��id = J.ID And T.ҩ��id=J.ID And A.��ҩ����=S.����(+) And a.����id = s.��������id(+) And E.�ⷿid = L.�ⷿid(+) And E.ҩƷid = L.ҩƷid(+) And A.����=Q.����(+) And  A.���˲���id=Q.����id(+) And A.���˿���id=Q.����id(+) and Q.��λ����=P.����(+) And Nvl(E.����, 0) = L.����(+) " & _
            " And n.����id = o.����id(+) And n.��ҳid = o.��ҳid(+) And A.����id = [1] And a.ҽ��id = r.ҽ��id And a.���ͺ� = r.���ͺ� And " & IIf(strStep = "04", "A.ִ��ʱ��", "A.ִ��ʱ��") & " Between [3] And [4] " & _
            " And e.���� = p.���� And e.No = p.No And e.�ⷿid + 0 = p.�ⷿid And e.ҩƷid + 0 = p.ҩƷid And e.��� = p.��� And (p.��¼״̬ = 1 Or Mod(p.��¼״̬, 3) = 0) "
            
        If lng����id <> 0 Then
            gstrSQL = gstrSQL & " And A.���˲���id + 0 =[2] "
        End If
        
        If strStep = "01" Then
            '����ҩ
            If bln������ Then gstrSQL = gstrSQL & " And Q.id=K.��ID and K.ҽ��id=M.id and Q.�����=1 and K.����ύ=1 "
            
            gstrSQL = gstrSQL & " And (" & IIf(bln���, "M.ҩʦ��˱�־=1 And", "") & " A.����״̬=1) "
        ElseIf strStep = "02" Then
            '����ҩ
            gstrSQL = gstrSQL & " And A.����״̬=2 "
        ElseIf strStep = "03" Then
            '������
            gstrSQL = gstrSQL & " And A.����״̬=4 "
        ElseIf strStep = "11" Then
            '���������
            gstrSQL = gstrSQL & " And A.����״̬=10 "
        ElseIf strStep = "12" Then
            '�ѷ���
            gstrSQL = gstrSQL & " And A.����״̬=5 "
        ElseIf strStep = "13" Then
            '��ǩ��
            gstrSQL = gstrSQL & " And A.����״̬=6 "
        ElseIf strStep = "14" Then
            '�Ѿܾ�ǩ��
            gstrSQL = gstrSQL & " And A.����״̬=7 "
        ElseIf strStep = "04" Then
            gstrSQL = gstrSQL & " And A.����״̬=9 "
        End If
        
        If intPack = 1 Then
            '�����
            gstrSQL = gstrSQL & " And Nvl(A.�Ƿ���,0)=0 "
        ElseIf intPack = 2 Then
            '�������������������������Ĵ��
            gstrSQL = gstrSQL & " And Nvl(A.�Ƿ���,0) In (1,2) "
        End If
    End If
    
    Set Piva_GetTrans = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Һ��ҩ��¼", lngCenterID, lng����id, dateExeStart, dateExeEnd)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Public Function PIVA_GetTransCount(ByVal lngCenterID As Long, ByVal dateExeStart As Date, ByVal dateExeEnd As Date, ByVal bln��� As Boolean, ByVal bln������ As Boolean, Optional intCheck As Integer) As ADODB.Recordset
    'ȡ������Һ����Ŀ
    'lngCenterID����Һ��������ID
    'dateExeStart��dateExeEnd����Һ��ҩ���ݵ�ִ��ʱ�䷶Χ
    On Error GoTo errHandle
    
    gstrSQL = "select ����, ����id, ����,  ����,ҩʦ��˱�־,����,���� from (with W as (Select Distinct a.����״̬,c.ҩʦ��˱�־, a.���˲���id As ����id, '[' || b.���� || ']' || b.���� As ����,b.����,b.����, c.���id As ҽ��id,A.id " & vbNewLine & _
        "       From ��Һ��ҩ��¼ A, ���ű� B, ����ҽ����¼ C" & IIf(bln������, ",��������¼ Q,���������ϸ K ", "") & vbNewLine & _
        "       Where a.���˲���id = b.Id And a.ҽ��id = c.���id And c.ִ������ <> 5 And a.����id = [1] And" & IIf(bln������, " c.id=k.ҽ��id and Q.id=K.��id and K.����ύ=1 and Q.�����=1 and", "") & vbNewLine & _
        "             a.ִ��ʱ�� Between [2] And [3]" & vbNewLine & _
        "             And Exists" & vbNewLine & _
        "        (Select 1 From ��Һ��ҩ���� D Where d.��¼id = a.Id))," & vbNewLine & _
        "       R as (Select Distinct a.����״̬, a.���˲���id As ����id, '[' || b.���� || ']' || b.���� As ����,b.����,b.����," & vbNewLine & _
        "                                 a.Id" & vbNewLine & _
        "                 From ��Һ��ҩ��¼ A, ���ű� B" & vbNewLine & _
        "                 Where a.���˲���id = b.Id  And a.����id = [1] And" & vbNewLine & _
        "                       a.ִ��ʱ�� Between [2]  and" & vbNewLine & _
        "                       [3] And Exists" & vbNewLine & _
        "                  (Select 1 From ��Һ��ҩ���� D Where d.��¼id = a.Id))"


    If bln��� = True Then
        '���ҽ��
        If intCheck = 0 Then
            gstrSQL = gstrSQL & " Select ����, ����id, ����, Count(ҽ��id) As ����,0 ҩʦ��˱�־,����,���� " & vbNewLine & _
            "From ( select Distinct '00' ����,����id,����,ҽ��id,����,���� from  W where (Nvl(ҩʦ��˱�־, 0) = 0 or Nvl(ҩʦ��˱�־, 0)=3)  and ����״̬=1)" & vbNewLine & _
            "Group By ����, ����id, ����,����,����" & vbNewLine & _
            "union all"
        Else
            gstrSQL = gstrSQL & "select ����,����id,����,count(ҽ��id) as ����,Nvl(ҩʦ��˱�־,0) ҩʦ��˱�־,����,���� from (" & _
                " Select distinct '00' As ����, D.���˲���id As ����id, '[' || B.���� || ']' || B.���� As ����, c.���id As ҽ��id,c.ҩʦ��˱�־,B.����,b.���� " & _
                " From ҩƷ�շ���¼ A, ���ű� B,����ҽ����¼ C,סԺ���ü�¼ D " & IIf(bln������, ",��������¼ Q,���������ϸ K ", "") & vbNewLine & _
                " Where D.���˲���id = B.ID And D.ҽ�����=C.id And A.����id=D.id And C.ִ������<>5  And A.�ⷿid = [1] and A.����=9 And A.�������� Between [2] And [3]) " & IIf(bln������, " c.id=k.ҽ��id and Q.id=K.��id and K.����ύ=1 and Q.�����=1 and", "") & vbNewLine & _
                " Group By ����,����id,����,Nvl(ҩʦ��˱�־,0),����,���� " & vbNewLine & _
                "union all"
        End If
        
'        gstrSQL = gstrSQL & " Select ����, ����id, ����, Count(ҽ��id) As ����,1 ҩʦ��˱�־" & vbNewLine & _
'            "From ( select Distinct '00' ����,����id,����,ҽ��id from  W where Nvl(ҩʦ��˱�־, 0) = 0 and ����״̬=1)" & vbNewLine & _
'            "Group By ����, ����id, ����" & vbNewLine & _
'            "union all"
            
         '��ҩ
        gstrSQL = gstrSQL & " Select ����, ����id, ����, Count(id) As ����,1 ҩʦ��˱�־,����,���� " & vbNewLine & _
            "From ( select '01' ����,����id,����,ҽ��id,id,����,���� from  W where Nvl(ҩʦ��˱�־, 0) =1 and ����״̬=1)" & vbNewLine & _
            "Group By ����, ����id, ����,����,����"
    Else
        '��ҩ
        gstrSQL = gstrSQL & " Select ����, ����id, ����, Count(id) As ����,1 ҩʦ��˱�־,����,���� " & vbNewLine & _
            "From ( select '01' ����,����id,����,ҽ��id,id,����,���� from  W where ����״̬=1)" & vbNewLine & _
            "Group By ����, ����id, ����,����,����"
    End If
    '��ҩ
    gstrSQL = gstrSQL & " Union All " & _
        "Select ����, ����id, ����, Count(id) As ����,1 ҩʦ��˱�־,����,���� " & vbNewLine & _
        "From ( select '02' ����,����id,����,id,����,���� from  R where ����״̬=2)" & vbNewLine & _
        "Group By ����, ����id, ����,����,����"

    '����
    gstrSQL = gstrSQL & " Union All " & _
        "Select ����, ����id, ����, Count(id) As ����,1 ҩʦ��˱�־,����,���� " & vbNewLine & _
        "From ( select '03' ����,����id,���� ,id,����,���� from  R where ����״̬=4)" & vbNewLine & _
        "Group By ����, ����id, ����,����,����"

    '�������
    gstrSQL = gstrSQL & " Union All " & _
        "Select ����, ����id, ����, Count(id) As ����,1 ҩʦ��˱�־ ,����,����" & vbNewLine & _
        "From ( select '04' ����,����id,����,id,����,���� from  R where ����״̬=9)" & vbNewLine & _
        "Group By ����, ����id, ����,����,����"

        
    If bln��� = True Then
        If intCheck = 0 Then
            '�����ͨ��ҽ���鿴
            gstrSQL = gstrSQL & " Union All " & _
                "Select ����, ����id, ����, Count(ҽ��id) As ����,1 ҩʦ��˱�־,����,���� " & vbNewLine & _
                "From ( select Distinct  '10' ����,����id,����,ҽ��id,����,���� from  W where  Nvl(ҩʦ��˱�־, 0) =1)" & vbNewLine & _
                "Group By ����, ����id, ����,����,����"
    
            'δ���ͨ��ҽ���鿴
            gstrSQL = gstrSQL & " Union All " & _
                "Select ����, ����id, ����, Count(ҽ��id) As ����,2 ҩʦ��˱�־,����,���� " & vbNewLine & _
                "From ( select Distinct  '11' ����,����id,����,ҽ��id,����,���� from  W where  Nvl(ҩʦ��˱�־, 0) =2)" & vbNewLine & _
                "Group By ����, ����id, ����,����,����"
        Else
            gstrSQL = gstrSQL & " Union All " & _
                "select ����,����id,����,count(ҽ��id) as ����,ҩʦ��˱�־,����,���� from (" & _
                " Select distinct '10' As ����, D.���˲���id As ����id, '[' || B.���� || ']' || B.���� As ����, c.���id As ҽ��id,c.ҩʦ��˱�־,B.����,B.���� " & _
                " From ҩƷ�շ���¼ A, ���ű� B,����ҽ����¼ C,סԺ���ü�¼ D " & _
                " Where D.���˲���id = B.ID And D.ҽ�����=C.id And A.����id=D.id and A.����=9  And C.ִ������<>5 and c.ҩʦ��˱�־=1 And A.�ⷿid = [1] And A.�������� Between [2] And [3]) " & _
                " Group By ����,����id,����,ҩʦ��˱�־,����,���� "
                
            gstrSQL = gstrSQL & " Union All " & _
                "select ����,����id,����,count(ҽ��id) as ����,ҩʦ��˱�־,����,���� from (" & _
                " Select distinct '11' As ����, D.���˲���id As ����id, '[' || B.���� || ']' || B.���� As ����, c.���id As ҽ��id,c.ҩʦ��˱�־,B.����,B.���� " & _
                " From ҩƷ�շ���¼ A, ���ű� B,����ҽ����¼ C,סԺ���ü�¼ D " & _
                " Where D.���˲���id = B.ID And D.ҽ�����=C.id And A.����id=D.id and A.����=9 And C.ִ������<>5 and c.ҩʦ��˱�־=2 And A.�ⷿid = [1] And A.�������� Between [2] And [3]) " & _
                " Group By ����,����id,����,ҩʦ��˱�־,����,���� "
        End If

    End If
    '�ѷ��Ͳ鿴
    gstrSQL = gstrSQL & " Union All " & _
        "Select ����, ����id, ����, Count(id) As ����,1 ҩʦ��˱�־,����,���� " & vbNewLine & _
        "From ( select '12' ����,����id,����,id,����,���� from  R where ����״̬=5)" & vbNewLine & _
        "Group By ����, ����id, ����,����,����"

    '��ǩ�ղ鿴
    gstrSQL = gstrSQL & " Union All " & _
        "Select ����, ����id, ����, Count(id) As ����,1 ҩʦ��˱�־,����,���� " & vbNewLine & _
        "From ( select '13' ����,����id,����,id,����,���� from  R where ����״̬=6)" & vbNewLine & _
        "Group By ����, ����id, ����,����,����"

    '�ܾ�ǩ�ղ鿴
    gstrSQL = gstrSQL & " Union All " & _
        "Select ����, ����id, ����, Count(id) As ����,1 ҩʦ��˱�־,����,���� " & vbNewLine & _
        "From ( select '14' ����,����id,����,id,����,���� from  R where ����״̬=7)" & vbNewLine & _
        "Group By ����, ����id, ����,����,����"

    '��������˲鿴
    gstrSQL = gstrSQL & " Union All " & _
        "Select '15' As ����, ����id, ����, Sum(����) As ����,1 ҩʦ��˱�־,����,���� " & vbNewLine & _
        "From (Select a.���˲���id As ����id, '[' || b.���� || ']' || b.���� As ����,����,b.����, Count(a.Id) As ����" & vbNewLine & _
        "       From (Select ID, ���˲���id" & vbNewLine & _
        "              From ��Һ��ҩ��¼ A" & vbNewLine & _
        "              Where a.����id = [1] And a.ִ��ʱ�� Between [2] And [3] And Nvl(a.����״̬, 0) In (10,11)) A, ���ű� B" & vbNewLine & _
        "       Where a.���˲���id = b.Id" & vbNewLine & _
        "       Group By a.���˲���id, '[' || b.���� || ']' || b.����,b.����,����)" & vbNewLine & _
        "   Group By ����id, ����,����,����"
    'ҽ�����˲鿴
    gstrSQL = gstrSQL & " Union All " & _
        " Select '16' As ����, A.���˲���id As ����id, '[' || B.���� || ']' || B.���� As ����, Count(A.ID) As ����,1 ҩʦ��˱�־,b.����,b.���� " & _
        " From ��Һ��ҩ��¼ A, ���ű� B " & _
        " Where A.���˲���id = B.ID And A.����״̬=12 And A.����id = [1] And A.ִ��ʱ�� Between [2] And [3] " & _
        " Group By A.���˲���id, '[' || B.���� || ']' || B.����,����,���� "
        
    gstrSQL = gstrSQL & " Order By ����, ���� )"
        
    Set PIVA_GetTransCount = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ������Һ����Ŀ", lngCenterID, dateExeStart, dateExeEnd)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function



Public Sub PIVA_AnalysisTrans(ByVal lngCenterID As Long, ByVal dateStart As String, ByVal dateEnd As String)
    'PIVA��̨�������ֽⷢҩ����������Һ��
    'lngCenterID����Һ��������ID
    'dateStart��dateEnd����ҩ���ݵ�����ʱ�䷶Χ
    On Error GoTo ErrHand
    gstrSQL = "Zl_��Һ��ҩ��¼_Insert("
    '��������ID
    gstrSQL = gstrSQL & lngCenterID
    '��ʼʱ��
    gstrSQL = gstrSQL & "," & dateStart
    '����ʱ��
    gstrSQL = gstrSQL & "," & dateEnd
    gstrSQL = gstrSQL & ")"

    Call zlDatabase.ExecuteProcedure(gstrSQL, "������Һ��ҩ��¼")
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Public Function DeptSendWork_Get��������() As Recordset
'��ȡ���˿������ƣ�ȡ��������Ϊ�ٴ�����Ĳ���
    On Error GoTo ErrHand
    
    gstrSQL = "Select distinct a.Id, a.����, a.����,zlSpellCode(a.����) ����,zlWBCode(a.����) ��ʼ���, a.����ʱ��" & vbNewLine & _
            "From ���ű� A, ��������˵�� B" & vbNewLine & _
            "Where a.Id = b.����id And (b.�������� = '�ٴ�' Or b.�������� = '����') And" & vbNewLine & _
            "      (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000/1/1', 'yyyy/mm/dd'))"
    
    
    Set DeptSendWork_Get�������� = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ϣ")
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Public Function DeptSendWork_Get��ҩ����() As Recordset
'��ȡҩƷ����ҩ����
    On Error GoTo ErrHand
    gstrSQL = "select ����,���� from ��Һ��ҩ����"
    
    Set DeptSendWork_Get��ҩ���� = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ҩ����")
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function DeptSendWork_GetƵ��() As Recordset
'��ȡҩƷ����ҩ����
    On Error GoTo ErrHand
    gstrSQL = "select ����,����,Ӣ������ from ����Ƶ����Ŀ where ���� not like '-%'"
    
    Set DeptSendWork_GetƵ�� = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡƵ��")
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function DeptSendWork_Get�շ���Ŀ() As Recordset
'��ȡ�շ���Ŀ
    On Error GoTo ErrHand
    gstrSQL = "select id,����,����,���㵥λ,˵�� from �շ���ĿĿ¼ where ���='Z' and nvl(�Ƿ���,0)=0"
    
    Set DeptSendWork_Get�շ���Ŀ = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�շ���Ŀ")
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function DeptSendWork_��ҩ;��() As Recordset
'��ȡ��ҩ;��,Ŀǰֻ��ԡ�����Ӫ������
    On Error GoTo ErrHand
    gstrSQL = "select ID, ���� from ������ĿĿ¼ where ��� = 'E' and �������� = '2' and ִ�з��� = '1' and ִ�б�� = 2"
    
    Set DeptSendWork_��ҩ;�� = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ҩ;��")
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function PIVA_�Ѱ�ҩ��Һ��(ByVal lngCenterID As Long, ByVal dateExeStart As Date, ByVal lng����ID As Long) As Recordset
'��ȡ�ò��˵�����Ѿ���ҩ��δ��ҩ����Һ��
    On Error GoTo errHandle

    gstrSQL = "Select Distinct A.ID As ��ҩID, A.����id, A.���, A.��ҩ����, S.��ɫ,A.����, A.�Ա�, A.����, A.סԺ��, A.����,M.ҩʦ���ʱ��, A.���˲���id, A.���˿���id, A.ִ��ʱ��, A.ƿǩ��,A.���ʱ��,M.ִ��Ƶ��,A.�Ƿ��������,A.�Ƿ�����,A.�ֹ���������," & _
            " A.������Ա, A.����ʱ��,Nvl(A.��ӡ��־,0) As ��ӡ��־, A.�Ƿ���, B.���� As ���˲���, C.���� As ���˿���, D.�շ�id, E.����, E.NO, F.���� As ҩƷ����, " & _
            " F.���� As ͨ����, H.���� As ��Ʒ��, I.���� As Ӣ����, F.���, E.����, E.����, E.����, J.���㵥λ As ������λ,J.id ҩ��id, E.Ƶ��," & _
            " Case Nvl(E.�����, 'δ���') When 'δ���' Then E.ʵ������ * Nvl(E.����, 1) / G.סԺ��װ Else 0 End As ��ҩ����,M.����id,M.��ҳid, " & _
            " (D.���� / G.סԺ��װ)  As ����,D.���� As ʵ������, G.סԺ��λ As ��λ,Nvl(E.����,0) As ����, Nvl(L.ʵ������, 0)/ G.סԺ��װ As �������, Nvl(M.�����,-1) �����, E.�÷�, E.ҩƷid, n.��� As �������,E.����id, o.����, A.��ҩ����,r.����ʱ�� As ҽ������ʱ��,nvl(X.��ҩ����,'') ��ҩ���� " & _
            " From  ��Һ��ҩ��¼ A, ���ű� B, ���ű� C, ��Һ��ҩ���� D, ҩƷ�շ���¼ E, �շ���ĿĿ¼ F, ҩƷ��� G,��ҺҩƷ���� X,  �շ���Ŀ���� H, ������Ŀ���� I, ������ĿĿ¼ J, ����ҽ����¼ M, סԺ���ü�¼ N, ������ҳ O ,��ҩ�������� S "


        gstrSQL = gstrSQL & ",(Select �ⷿid, ҩƷid, Nvl(����, 0) As ����, Nvl(ʵ������, 0) As ʵ������ " & _
            " From ҩƷ��� Where ���� = 1 And �ⷿid = [1]) L, ҩƷ�շ���¼ P, ����ҽ������ R "

        gstrSQL = gstrSQL & " Where A.���˲���id = B.ID And A.���˿���id = C.ID And A.ID = D.��¼id And D.�շ�id = E.ID And E.ҩƷid = F.ID And F.ID = G.ҩƷid And G.ҩƷid=X.ҩƷid(+) And E.����id = N.ID And N.ҽ����� = M.ID And " & _
            " G.ҩƷid = H.�շ�ϸĿid(+) And H.����(+) = 3 And G.ҩ��id = I.������Ŀid(+) And I.����(+) = 2 And G.ҩ��id = J.ID And A.��ҩ����=S.����(+) And E.�ⷿid = L.�ⷿid(+) And E.ҩƷid = L.ҩƷid(+) And Nvl(E.����, 0) = L.����(+) " & _
            " And n.����id = o.����id(+) And n.��ҳid = o.��ҳid(+) And A.����id = [1] And a.ҽ��id = r.ҽ��id And a.���ͺ� = r.���ͺ� And A.ִ��ʱ�� between [2] and [3] " & _
            " And e.���� = p.���� And e.No = p.No And e.�ⷿid + 0 = p.�ⷿid And e.ҩƷid + 0 = p.ҩƷid And e.��� = p.��� And (p.��¼״̬ = 1 Or Mod(p.��¼״̬, 3) = 0) "



        gstrSQL = gstrSQL & " And A.����״̬=2 and M.����id=[4] "
        

        Set PIVA_�Ѱ�ҩ��Һ�� = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��Һ��ҩ��¼", lngCenterID, CDate(Format(dateExeStart, "yyyy-mm-dd 00:00:00")), CDate(Format(dateExeStart, "yyyy-mm-dd 23:59:59")), lng����ID)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function






