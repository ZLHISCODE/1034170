Attribute VB_Name = "mdl������Ժ"
Option Explicit

Public Function �����������_������Ժ(rs��ϸ As ADODB.Recordset, str���㷽ʽ As String, Optional ByRef strAdvance As String = "") As Boolean
'������rsDetail     ������ϸ(����)
'      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
'�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    Dim rs�㷨 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim clsҽ�� As New clsInsure
    Dim rs������� As New ADODB.Recordset
    Dim dblȫ�Է� As Currency, dbl�����Ը� As Currency, dbl����ͳ�� As Currency, dblTemp As Double
    Dim dbl����� As Double, cur���Ʊ��� As Currency, curҩƷ���� As Currency
    Dim dbl�����ʻ� As Double, cur������ As Currency
    Dim lng����ID As Long, cur�ܶ� As Currency
    Dim rs��׼��Ŀ As New ADODB.Recordset
    Dim dblTemp1 As Double, datCurr As Date
    
    datCurr = zlDatabase.Currentdate
    If rs��ϸ.RecordCount > 0 Then
        rs��ϸ.MoveFirst
        lng����ID = rs��ϸ("����ID")
    End If
    cur�ܶ� = 0: curҩƷ���� = 0: cur���Ʊ��� = 0
    While Not rs��ϸ.EOF
        gstrSQL = "select a.����,b.����,b.ͳ��ȶ�,b.�㷨,a.��� from �շ�ϸĿ a,����֧������ b,����֧����Ŀ c where " & _
            "b.id=c.����id and a.id=c.�շ�ϸĿid and c.����=" & TYPE_������Ժ & " and a.id=" & rs��ϸ!�շ�ϸĿID
        Call OpenRecordset(rsTemp, gstrSysName)
        If rsTemp.EOF Then
            dblȫ�Է� = dblȫ�Է� + rs��ϸ!ʵ�ս��
            cur������ = 0
        ElseIf rsTemp!�㷨 = 1 Then
            dblȫ�Է� = dblȫ�Է� + rs��ϸ!ʵ�ս�� * (1 - rsTemp!ͳ��ȶ� / 100)
            dbl����ͳ�� = dbl����ͳ�� + rs��ϸ!ʵ�ս�� * rsTemp!ͳ��ȶ� / 100
            cur������ = rs��ϸ!ʵ�ս�� * rsTemp!ͳ��ȶ� / 100
        ElseIf rsTemp!�㷨 = 2 Then
            dbl����ͳ�� = dbl����ͳ�� + IIf(rs��ϸ!ʵ�ս�� < rsTemp!ͳ��ȶ�, rs��ϸ!ʵ�ս��, rsTemp!ͳ��ȶ�)
            dblȫ�Է� = dblȫ�Է� + IIf(rs��ϸ!ʵ�ս�� < rsTemp!ͳ��ȶ�, 0, rs��ϸ!ʵ�ս�� - rsTemp!ͳ��ȶ�)
            cur������ = IIf(rs��ϸ!ʵ�ս�� < rsTemp!ͳ��ȶ�, rs��ϸ!ʵ�ս��, rsTemp!ͳ��ȶ�)
        End If
        If rsTemp!��� = "5" Or rsTemp!��� = "6" Or rsTemp!��� = "7" Then
            curҩƷ���� = curҩƷ���� + cur������
        Else
            cur���Ʊ��� = cur���Ʊ��� + cur������
        End If
        
        cur�ܶ� = cur�ܶ� + Nvl(rs��ϸ!ʵ�ս��, 0)
        rs��ϸ.MoveNext
    Wend
    g��������.�������ý�� = cur�ܶ�
'    dblTemp = dbl����ͳ��
    
    'ÿ�챨��������80
'    gstrSQL = "Select nvl(sum(a.ͳ�ﱨ�����),0) From ���ս����¼ a,���˷��ü�¼ b Where a.��¼ID=b.����id " & _
'        "and to_char(b.����ʱ��,'yyyy-mm-dd')='" & _
'        Format(datCurr, "yyyy-mm-dd") & "' And a.����=1 And b.����ID=" & lng����id & " And a.����=" & TYPE_������Ժ
'    Call OpenRecordset(rsTemp, gstrSysName)
'    If dblTemp + rsTemp(0) > 80 Then dblTemp = 80 - rsTemp(0)
    'ÿ�ŵ���ҩƷ��������80Ԫ
    
    '20051220 �¶� ���ݷ��ؼ�
    Dim cur�޶� As Currency, bln�����ֹ As Boolean
    gstrSQL = "Select ����ֵ From ���ղ��� Where ����=" & TYPE_������Ժ & " And ������='�����޶�'"
    Call OpenRecordset(rsTemp, "���ղ���")
    If rsTemp.EOF Then
        cur�޶� = 80
    Else
        If Val(rsTemp!����ֵ) > 0 Then
            cur�޶� = Val(rsTemp!����ֵ)
        Else
            cur�޶� = 80
        End If
    End If
    
    gstrSQL = "Select ����ֵ From ���ղ��� Where ����=" & TYPE_������Ժ & " And ������='�����ֹ'"
    Call OpenRecordset(rsTemp, "���ղ���")
    If rsTemp.EOF Then
        bln�����ֹ = False
    Else
        If Val(rsTemp!����ֵ) = 1 Then
            bln�����ֹ = True
        Else
            bln�����ֹ = False
        End If
    End If
    If curҩƷ���� > cur�޶� Then
        If bln�����ֹ = True Then
            MsgBox "�ѳ��������޶�" & Format(cur�޶�, "0.00") & "�������շѣ�", vbInformation, gstrSysName
            �����������_������Ժ = False
            Exit Function
        Else
            dblTemp = cur���Ʊ��� + cur�޶�
            dblȫ�Է� = dblȫ�Է� + curҩƷ���� - cur�޶�
        End If
    Else
        dblTemp = curҩƷ���� + cur���Ʊ���
    End If
    
    g��������.����ͳ���� = dbl����ͳ��
    g��������.ȫ�Էѽ�� = dblȫ�Է�
    g��������.�����Ը���� = 0
    g��������.ͳ�ﱨ����� = dblTemp
    str���㷽ʽ = "����ҽ��;" & dblTemp & ";0"
   
    �����������_������Ժ = True
End Function

Public Function סԺ�������_������Ժ(rs��ϸ As ADODB.Recordset) As String
'������rsDetail     ������ϸ(����)
'      cur���㷽ʽ  "������ʽ;���;�Ƿ������޸�|...."
'�ֶΣ�����ID,�շ�ϸĿID,����,����,ʵ�ս��,ͳ����,����֧������ID,�Ƿ�ҽ��
    Dim rs�㷨 As New ADODB.Recordset, rsTemp As New ADODB.Recordset
    Dim clsҽ�� As New clsInsure
    Dim rs������� As New ADODB.Recordset
    Dim dblȫ�Է� As Currency, dbl�����Ը� As Currency, dbl����ͳ�� As Currency, dblTemp As Double
    Dim dbl����� As Double, lng���� As Long, lng��ְ As Long, lng���� As Long
    Dim dbl�����ʻ� As Double, lng�����  As Long, blnȫ��ͳ�� As Boolean, bln������ As Boolean, bln�޷ⶥ�� As Boolean
    Dim lng����ID As Long, curȫ�� As Currency, dblͳ�� As Currency
    Dim rs��׼��Ŀ As New ADODB.Recordset, strTemp As String, str���� As String
    Dim dblTemp1 As Double, datCurr As Date

    datCurr = zlDatabase.Currentdate
    If rs��ϸ.RecordCount > 0 Then
        rs��ϸ.MoveFirst
        lng����ID = rs��ϸ("����ID")
    End If
    gstrSQL = "Select max(��ҳid) from ������ҳ Where ����id=" & lng����ID
    Call OpenRecordset(rsTemp, gstrSysName)
    g��������.��ҳID = Nvl(rsTemp(0), 1)
    g��������.����ID = lng����ID
    g��������.��� = Format(datCurr, "yyyy")
    
    gstrSQL = "select ��Ժ����,nvl(��Ժ����,to_date('3000-01-01','yyyy-MM-dd')) as ��Ժ���� " & _
              "from ������ҳ where ����ID=" & g��������.����ID & " and ��ҳID=" & g��������.��ҳID
    Call OpenRecordset(rsTemp, "�������")
    If rsTemp("��Ժ����") = CDate("3000-01-01") Then
        g��������.��;���� = 1
    Else
        '��ʾ�ò����Ѿ���Ժ
        g��������.��;���� = 0
    End If
    
    With g��������
        gstrSQL = "select A.����,A.��Ա���,A.��ְ,A.�����," & _
                  "      B.סԺ�����ۼ�,B.�ʻ������ۼ�,B.�ʻ�֧���ۼ�,B.����ͳ���ۼ�,B.ͳ�ﱨ���ۼ�" & _
                  " from �����ʻ� A,�ʻ������Ϣ B" & _
                  " where A.����ID=B.����ID(+) and A.����=B.����(+) " & _
                  "     and B.���(+)=" & .��� & " and A.����ID=" & .����ID & " and A.����=" & TYPE_������Ժ
        Call OpenRecordset(rsTemp, "�������")
        
        lng���� = IIf(IsNull(rsTemp("����")), 0, rsTemp("����"))
        lng��ְ = IIf(IsNull(rsTemp("��ְ")), 1, rsTemp("��ְ"))
        lng���� = IIf(IsNull(rsTemp("�����")), 0, rsTemp("�����"))
        .סԺ���� = IIf(IsNull(rsTemp("סԺ�����ۼ�")), 0, rsTemp("סԺ�����ۼ�"))
        .�ʻ��ۼ����� = IIf(IsNull(rsTemp("�ʻ������ۼ�")), 0, rsTemp("�ʻ������ۼ�"))
        .�ʻ��ۼ�֧�� = IIf(IsNull(rsTemp("�ʻ�֧���ۼ�")), 0, rsTemp("�ʻ�֧���ۼ�"))
        .�ۼƽ���ͳ�� = IIf(IsNull(rsTemp("����ͳ���ۼ�")), 0, rsTemp("����ͳ���ۼ�"))
        .�ۼ�ͳ�ﱨ�� = IIf(IsNull(rsTemp("ͳ�ﱨ���ۼ�")), 0, rsTemp("ͳ�ﱨ���ۼ�"))
    
        
        gstrSQL = "select �����,nvl(ȫ��ͳ��,0) as ȫ��ͳ�� ,nvl(������,0) as ������ ,nvl(�޷ⶥ��,0) as �޷ⶥ�� " & _
                " from ���������" & _
                " where ����=" & TYPE_������Ժ & " and nvl(����,0)=" & lng���� & _
                "       and ��ְ=" & lng��ְ & " and ����<=" & lng���� & " and (" & lng���� & "<=���� or ����=0)"
        Call OpenRecordset(rsTemp, "�������")
        If rsTemp.RecordCount = 0 Then
            MsgBox "���ڡ��������������������������õ���", vbInformation, gstrSysName
            Exit Function
        End If
        lng����� = rsTemp("�����")
        blnȫ��ͳ�� = (rsTemp("ȫ��ͳ��") = 1)
        bln������ = (rsTemp("������") = 1)
        bln�޷ⶥ�� = (rsTemp("�޷ⶥ��") = 1)
    End With
    
    While Not rs��ϸ.EOF
        gstrSQL = "select a.����,b.����,b.ͳ��ȶ�,b.�㷨 from �շ�ϸĿ a,����֧������ b,����֧����Ŀ c where " & _
            "b.id=c.����id and a.id=c.�շ�ϸĿid and c.����=" & TYPE_������Ժ & " and a.id=" & rs��ϸ!�շ�ϸĿID
        Call OpenRecordset(rsTemp, gstrSysName)
        If rsTemp.EOF Then
            dblȫ�Է� = dblȫ�Է� + rs��ϸ!���
        ElseIf rsTemp!�㷨 = 1 Then
            dblȫ�Է� = dblȫ�Է� + rs��ϸ!��� * (1 - rsTemp!ͳ��ȶ� / 100)
            dbl����ͳ�� = dbl����ͳ�� + rs��ϸ!��� * rsTemp!ͳ��ȶ� / 100
        ElseIf rsTemp!�㷨 = 2 Then
'            dbl����ͳ�� = dbl����ͳ�� + IIf(rs��ϸ!��� < rsTemp!ͳ��ȶ�, rs��ϸ!���, rsTemp!ͳ��ȶ�)
'            dblȫ�Է� = dblȫ�Է� + IIf(rs��ϸ!��� < rsTemp!ͳ��ȶ�, 0, rs��ϸ!��� - rsTemp!ͳ��ȶ�)

            'Beging 20051228 �¶� ԭ������¼�͸�����¼�˹�ʽ��������
            If rs��ϸ!��� >= 0 Then
                dbl����ͳ�� = dbl����ͳ�� + IIf(rs��ϸ!��� < rsTemp!ͳ��ȶ�, rs��ϸ!���, rsTemp!ͳ��ȶ�)
                dblȫ�Է� = dblȫ�Է� + IIf(rs��ϸ!��� < rsTemp!ͳ��ȶ�, 0, rs��ϸ!��� - rsTemp!ͳ��ȶ�)
            Else
                dbl����ͳ�� = dbl����ͳ�� + IIf(Abs(rs��ϸ!���) < rsTemp!ͳ��ȶ�, rs��ϸ!���, -rsTemp!ͳ��ȶ�)
                dblȫ�Է� = dblȫ�Է� + IIf(Abs(rs��ϸ!���) < rsTemp!ͳ��ȶ�, 0, rs��ϸ!��� + rsTemp!ͳ��ȶ�)
            End If
            'End    20051228 �¶� ԭ������¼�͸�����¼�˹�ʽ��������
        End If
        curȫ�� = curȫ�� + Nvl(rs��ϸ!���, 0)
        rs��ϸ.MoveNext
    Wend
    dblTemp = dbl����ͳ��
    
    g��������.�������ý�� = curȫ��
    g��������.����ͳ���� = dbl����ͳ��
    g��������.ȫ�Էѽ�� = dblȫ�Է�
    g��������.�����Ը���� = 0
    g��������.ͳ�ﱨ����� = dbl����ͳ��
    
    gstrSQL = "Select * From ���˷��ü�¼ Where �����־=2 And ��¼״̬<>0 And nvl(���ӱ�־,0)<>9 and nvl(ʵ�ս��,0)<>0 and ����id=" & lng����ID & " And ��ҳid=" & g��������.��ҳID & " order by ��ҳID,���"
    Call OpenRecordset(rsTemp, gstrSysName)
    While Not rsTemp.EOF
        gstrSQL = "select a.����,b.����,b.ͳ��ȶ�,b.�㷨,c.����ID from �շ�ϸĿ a,����֧������ b,����֧����Ŀ c where " & _
            "b.id=c.����id and a.id=c.�շ�ϸĿid and c.����=" & TYPE_������Ժ & " and a.id=" & rsTemp!�շ�ϸĿID
        Call OpenRecordset(rs�㷨, gstrSysName)
        If rs�㷨.EOF Then
            dblͳ�� = 0
        ElseIf rs�㷨!�㷨 = 1 Then
            dblͳ�� = rsTemp!ʵ�ս�� * rs�㷨!ͳ��ȶ� / 100
        ElseIf rs�㷨!�㷨 = 2 Then
'            dblͳ�� = IIf(rsTemp!ʵ�ս�� < rs�㷨!ͳ��ȶ�, rsTemp!ʵ�ս��, rs�㷨!ͳ��ȶ�)
            'Beging 20051228 �¶� ԭ������¼�͸�����¼�˹�ʽ��������
            If rsTemp!ʵ�ս�� >= 0 Then
                dblͳ�� = IIf(rsTemp!ʵ�ս�� < rs�㷨!ͳ��ȶ�, rsTemp!ʵ�ս��, rs�㷨!ͳ��ȶ�)
            Else
                dblͳ�� = IIf(Abs(rsTemp!ʵ�ս��) < rs�㷨!ͳ��ȶ�, rsTemp!ʵ�ս��, -rs�㷨!ͳ��ȶ�)
            End If
            'End    20051228 �¶� ԭ������¼�͸�����¼�˹�ʽ��������
        End If
        If Not rs�㷨.EOF Then
            strTemp = rs�㷨!����id
            str���� = rs�㷨(1)
        Else
            str���� = "�Է�"
            strTemp = "NULL"
        End If
        gcnOracle.Execute "Delete From ������ϸ Where ��¼ID=" & rsTemp!ID
        gcnOracle.Execute "insert into ������ϸ values (" & dblͳ�� & "," & strTemp & ",'" & str���� & "'," & rsTemp!ID & ")"
        rsTemp.MoveNext
    Wend
    
    'ѭ������������Ŀ�ķ�������
    Call UpdateClass(g��������.����ID, g��������.��ҳID)
    
    סԺ�������_������Ժ = "����ҽ��;" & dblTemp & ";0"
End Function

Public Function ҽ����Ŀ_������Ժ(����ID As Long, �շ�ϸĿID As Long, ��� As Currency, _
    ByVal bln���� As Boolean, Optional ByVal intInsure As Integer) As String
    '��ȡҽ��������Ϊ�������ͷ��ظ���������
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = "Select B.����  " & _
             " From ����֧����Ŀ A,����֧������ B " & _
             " Where A.����ID=B.Id And A.����=B.���� And A.����=" & intInsure & " And A.�շ�ϸĿID=" & �շ�ϸĿID
    Call OpenRecordset(rsTemp, "��ȡҽ��������Ϊ�������ͷ��ظ���������")
    
    If rsTemp.RecordCount <> 0 Then
        ҽ����Ŀ_������Ժ = Nvl(rsTemp!����)
    End If
End Function

Private Sub UpdateClass(ByVal lng����ID As Long, ByVal lng��ҳID As Long)
    Dim str�������� As String
    Dim rsTemp As New ADODB.Recordset
    'ѭ������������Ŀ�ķ�������
    gstrSQL = "Select ID,����ID,�շ�ϸĿID,�������� From ���˷��ü�¼" & _
        " Where ����ID=" & lng����ID & " And ��ҳID=" & lng��ҳID & _
        " And Nvl(�Ƿ��ϴ�,0)=1 And Nvl(���ӱ�־,0)<>9 And Nvl(��¼״̬,0)<>0 And Nvl(ʵ�ս��,0)<>0 And �������� is null"
    Call OpenRecordset(rsTemp, "ѭ������������Ŀ�ķ�������")
    
    With rsTemp
        Do While Not .EOF
            str�������� = ҽ����Ŀ_������Ժ(!����ID, !�շ�ϸĿID, 0, False, TYPE_������Ժ)
            gstrSQL = "ZL_���˼��ʼ�¼_�ϴ�(" & !ID & ",NULL,NULL,'" & str�������� & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, "���·�������")
            .MoveNext
        Loop
    End With
End Sub
