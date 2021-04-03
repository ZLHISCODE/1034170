--[��������]1
--[�����߰汾��]10.34.90
--���ű�֧�ִ�ZLHIS+ v10.34.90 ������ v10.34.100
--����ϵͳ�����ߵ�¼PLSQL��ִ�����нű�
--�ű�ִ�к����ֹ�������������
Define n_System=100;
-------------------------------------------------------------------------------
------------------------------------------------------------------------------


-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--103000:���ϴ�,2017-03-06,����ͥ�绰�������ֻ�����ʽ��������д���ֻ�����
--���÷�Χ��������Ϣ�������ֻ��ŵİ汾
--��������:��ͥ�绰�����ֻ��Ÿ�ʽ������ͥ�绰��д���ֻ�����
--������Χ���������в�����Ϣ��¼
--��ʱ˵��: ���ݹ�451W�������������ݼ�¼345652���������������ű���24������ִ����ɣ����Ի�������:
--1.Ӳ������
--     IBM 3650 M4,cpu E5-2620 2.1GHz,2*6�ˣ�32G�ڴ�
--     V3700�洢,SASӲ��,10K RPM,Raid 10
--2.�������
--     Windows 2008,Oracle 10.2.0.4 64bit
--     ��־�ļ�500M/����Log Buffer����Ϊ200M,PGAΪ9G,SGAΪ�Զ��������25G
--3.���ݻ���
--     ������Ϣ����Ҫ���������ݼ�¼��345652
Create Or Replace Procedure Zl1_Optional_������Ϣ���� As
  Cursor c_Pati Is
    Select ����ID From ������Ϣ Where  Length(��ͥ�绰) = 11 and substr(��ͥ�绰,1,3) in 
            ('139','138','137','136','135','134','159','158','157',
            '150','151','152','147','188','187','182','183','184','178',
            '130','131','132','156','155','186','185','145','176','133','153','189','180','181','177','173','170') And �ֻ��� is Null;

  t_PatiId      t_Strlist := t_Strlist();
  n_Array_Size Number := 100000; --ÿ��ʮ�򣬶��˿���PGA����
  I            Number(8) := 0; --ÿ����100������¼�ύһ��,���˿���Undo����,�����ύ����Ƶ��
  v_���� varchar2(500);
Begin
  Select Max(����) Into v_���� From Zlupgradeconfig Where ��Ŀ = '������Ϣ�ֻ�������_20170228';
  Update Zlupgradeconfig Set ���� = Null Where ��Ŀ = '������Ϣ�ֻ�������_20170228';
  If Sql%Notfound  Then
    Insert Into Zlupgradeconfig (��Ŀ, ����) Values ('������Ϣ�ֻ�������_20170228', Null);
  End If;
  If Nvl(v_����, 0) = '�ɹ�' Then
    --�����������ɹ�
    Return;
  End If;
  Open c_Pati;
  Loop
    Fetch c_Pati Bulk Collect
      Into t_PatiId Limit n_Array_Size;
    Exit When t_PatiId.Count = 0;
  
    Forall I In 1 .. t_PatiId.Count
      Update ������Ϣ set �ֻ��� = ��ͥ�绰 Where ����ID = t_PatiId(I);
  
    If I = 9 Then
      Update Zlupgradeconfig Set ���� = To_Number(Nvl(����, 0)) + I * n_Array_Size Where ��Ŀ = '������Ϣ�ֻ�������_20170228';
      Commit;
      I := 0;
    Else
      I := I + 1;
    End If;
  End Loop;
  Update Zlupgradeconfig Set ���� = '�ɹ�' Where ��Ŀ = '������Ϣ�ֻ�������_20170228';
  Commit;
  Close c_Pati;
End Zl1_Optional_������Ϣ����;
/

--101262:������,2017-02-13,�С�����ҽ�����桱�Ĵ������ݣ�ͬʱ������������
--������������ʹ�õ�����LISϵͳ(���磺�Ƿ�LIS)�У����õ�����LIS�ӿڳ���zlLISInterface���е�Zl_���鱨�浥_Insert����
--���ڵ��Ӳ�����¼�в����������ݣ��������ﲡ����������û����ҳid�ģ���ʱ������ڵ��Ӳ�����¼�У���д����ҳΪ0����ȷӦ����д�Һŵ�id��
--���÷�Χ��ʹ�õ�����LISϵͳ��������Ӱ�����а汾�������ű������а汾ͨ��
--��������:
--1.�������Ӳ�����¼�У� ����Ϊ����Ĳ��ˣ�������д�Һŵ�id��
--2.������Դ�������ģ���д��Ӧ����ҳid��
--������Χ��
--1.���Ӳ�����¼���Ѿ���������ʷ��������,��ҳid�ǿջ���0 �����ݣ�����������
--2.ֱ�ӵǼǵļ��顢��鲡��û�йҺŵ���,��ҳIDҲ��Ϊ0���������ݲ�������
--��ʱ˵��: �������ݼ�¼2225269���������������ű���45������ִ����ɣ����Ի�������:
--1.Ӳ������
--     IBM 3650 M4,cpu E5-2620 2.1GHz,2*6�ˣ�32G�ڴ�
--     V3700�洢,SASӲ��,10K RPM,Raid 10
--2.�������
--     Windows 2008,Oracle 10.2.0.4 64bit
--     ��־�ļ�500M/����Log Buffer����Ϊ200M,PGAΪ9G,SGAΪ�Զ��������25G
--3.���ݻ���
--     ���Ӳ�����¼����Ҫ���������ݼ�¼��2225269
Create Or Replace Procedure Zl1_Optional_���ﲡ������_1 As
  n_��ҳid ���Ӳ�����¼.��ҳid%Type;
  I        Number(8) := 0;
  v_���� varchar2(500);
Begin
  Update Zlupgradeconfig Set ���� = Null Where ��Ŀ = '����������벡������_20160217';
  If Sql%NotFound Then
    Insert Into Zlupgradeconfig (��Ŀ, ����) Values ('����������벡������_20160217', Null);
  End If;
  Commit;
  If Nvl(v_����, '��') = '�ɹ�' Then
    --�����������ɹ�
    Return;
  End If;
  For c_Rec In (Select ID, ����id, ����id, ����ʱ��, ������Դ From ���Ӳ�����¼ A Where Nvl(��ҳid, 0) = 0) Loop
  
    If c_Rec.������Դ = 1 Then
	  --�йҺŵ�id�Ĳ��ˣ�������������
      Select Nvl(Max(d.Id), 0)
      Into n_��ҳid
      From ����ҽ������ B, ����ҽ����¼ C, ���˹Һż�¼ D
      Where b.����id = c_Rec.Id And b.ҽ��id = c.Id And c.�Һŵ� = d.No;
    
      --ֱ�ӵǼǵļ��顢��鲡��û�йҺŵ���,��ҳIDҲ��Ϊ0���������ݲ�����
    Else
      Select Nvl(Max(��ҳid), 0)
      Into n_��ҳid
      From ������ҳ
      Where ����id = c_Rec.����id And c_Rec.����ʱ�� Between ��Ժ���� AND  Nvl(��Ժ����,c_Rec.����ʱ��);        
    End If;
  
    If n_��ҳid > 0 Then
      Update ���Ӳ�����¼ Set ��ҳid = n_��ҳid Where Nvl(��ҳid, 0) = 0 And ID = c_Rec.Id;
    
      --ÿһ�����ύһ��
      I := I + 1;
      If I = 10000 Then
        Update Zlupgradeconfig Set ���� = To_Number(Nvl(����, 0)) + I Where ��Ŀ = '����������벡������_20160217';
        Commit;
        I := 0;
      End If;
    End If;
  End Loop;
  Update Zlupgradeconfig Set ���� = To_Number(Nvl(����, 0)) + I Where ��Ŀ = '����������벡������_20160217';
  Update Zlupgradeconfig Set ���� = '�ɹ�' Where ��Ŀ = User || '����������벡������_20160217';
  Commit;
End Zl1_Optional_���ﲡ������_1;
/
--101262:������,2017-02-13,ȱ�١�����ҽ�����桱�Ĵ������ݣ�ͬʱ������������
--������������ʹ�õ�����LISϵͳ(���磺�Ƿ�LIS)�У����õ�����LIS�ӿڳ���zlLISInterface��
--����ɲ���ȱ�١�����ҽ�����桱�Ĵ�������
--���÷�Χ��ʹ�õ�����LISϵͳ��������Ӱ�����а汾�������ű������а汾ͨ��
--��������:
--1.�������Ӳ�����¼�У� ����Ϊ����Ĳ��ˣ�������д�Һŵ�id��
--2.������Դ�������ģ���д��Ӧ����ҳid��
--������Χ��
--1.���Ӳ�����¼���Ѿ���������ʷ��������,��ҳid�ǿջ���0 �����ݣ�����������
--2.ֱ�ӵǼǵļ��顢��鲡��û�йҺŵ���,��ҳIDҲ��Ϊ0���������ݲ�������
--��ʱ˵��: �����С�����ҽ�����桱�Ĵ������ݣ����ݼ�¼2225269���������������ű���2������ִ����ɣ����Ի�������:
--1.Ӳ������
--     IBM 3650 M4,cpu E5-2620 2.1GHz,2*6�ˣ�32G�ڴ�
--     V3700�洢,SASӲ��,10K RPM,Raid 10
--2.�������
--     Windows 2008,Oracle 10.2.0.4 64bit
--     ��־�ļ�500M/����Log Buffer����Ϊ200M,PGAΪ9G,SGAΪ�Զ��������25G
--3.���ݻ���
--     XXҽԺ����10�������
--     ���Ӳ�����¼����Ҫ���������ݼ�¼��2225269
--����������������1891
--ʣ��127�в�����ԴΪ2������δ��������ȱ������ҳ���ݣ�
Create Or Replace Procedure Zl1_Optional_���ﲡ������_2 As
  n_��ҳid ���Ӳ�����¼.��ҳid%Type;
  I        Number(8) := 0;
  v_���� varchar2(500);
Begin
  Update Zlupgradeconfig Set ���� = Null Where ��Ŀ = '����������벡������_20160218';
  If Sql%NotFound Then
    Insert Into Zlupgradeconfig (��Ŀ, ����) Values ('����������벡������_20160218', Null);
  End If;
  Commit;
  If Nvl(v_����, '��') = '�ɹ�' Then
    --�����������ɹ�
    Return;
  End If;
  For c_Rec In (Select ID, ����id, ����id, ����ʱ��, ������Դ
                From ���Ӳ�����¼ A
                Where Nvl(��ҳid, 0) = 0 And Not Exists (Select 1 From ����ҽ������ B Where a.Id = b.����id)) Loop
  
    If c_Rec.������Դ = 1 Then
      Select Nvl(Max(ID), 0)
      Into n_��ҳid
      From ���˹Һż�¼
      Where ����id = c_Rec.����id And ִ�в���id = c_Rec.����id And �Ǽ�ʱ�� < c_Rec.����ʱ��;
    
      If n_��ҳid = 0 Then
        Select Nvl(Max(ID), 0)
        Into n_��ҳid
        From ���˹Һż�¼
        Where ����id = c_Rec.����id And �Ǽ�ʱ�� < c_Rec.����ʱ��;
      End If;
    Else
      Select Nvl(Max(��ҳid), 0)
      Into n_��ҳid
      From ������ҳ
      Where ����id = c_Rec.����id And c_Rec.����ʱ�� Between ��Ժ���� And Nvl(��Ժ����, c_Rec.����ʱ��);
    
      --���û�в�����ҳ���ݣ���3����������Һ����ݣ�����ִ�п��ң���˵����������Դ=2�������Ǵ���ģ�Ӧ����1
      If n_��ҳid = 0 Then
        Select Nvl(Max(ID), 0)
        Into n_��ҳid
        From ���˹Һż�¼
        Where ����id = c_Rec.����id And �Ǽ�ʱ�� Between Trunc(c_Rec.����ʱ�� - 3) And c_Rec.����ʱ��;
        
        If n_��ҳid <> 0 Then
          Update ���Ӳ�����¼ Set ��ҳid = n_��ҳid, ������Դ = 1 Where Nvl(��ҳid, 0) = 0 And ID = c_Rec.Id;
        End If;      
      End If;
    End If;
  
    If n_��ҳid > 0 Then
      Update ���Ӳ�����¼ Set ��ҳid = n_��ҳid Where Nvl(��ҳid, 0) = 0 And ID = c_Rec.Id;
    
      --ÿһ�����ύһ��
      I := I + 1;
      If I = 10000 Then
        Update Zlupgradeconfig Set ���� = To_Number(Nvl(����, 0)) + I Where ��Ŀ = '����������벡������_20160218';
        Commit;
        I := 0;
      End If;
    End If;
  End Loop;
  Update Zlupgradeconfig Set ���� = To_Number(Nvl(����, 0)) + I Where ��Ŀ = '����������벡������_20160218';
  Update Zlupgradeconfig Set ���� = '�ɹ�' Where ��Ŀ = User || '����������벡������_20160218';
  Commit;
End Zl1_Optional_���ﲡ������_2;
/


--98570:��ҵ��,2017-03-15,����ҩƷ�����ϴι�Ӧ��Ϊ�յļ�¼
--���÷�Χ���������⹺��⣬��ҩƷ������ϴι�Ӧ��Ϊ�յ����
--��������:������⹺���Ĺ�Ӧ����Ϣ���µ���Ӧ��ҩƷ����ϴι�Ӧ����Ϣ�У�ԭҩƷ����ϴι�Ӧ��Ϊ��ʱ��
--������Χ������ҩƷ������ϴι�Ӧ��Ϊ�յ����ݣ������⹺�������ʱ��
--��ʱ˵��: ҩƷ�շ���¼���ݹ�83911261����Ҫ�����⹺���в�ѯ���ݣ�ʵ������238��ҩƷ������2�ְ���ִ����ɣ��û�ʵ����Ҫ������ҩƷ����
--���Ի�������:
--1.Ӳ������
--     ��ͨPC��
--2.�������
--3.���ݻ���
--     ҩƷ�շ���¼��83911261
Create Or Replace Procedure Zl1_Optional_ҩƷ���_��Ӧ�� Is
Begin
  For r_ҩƷ��� In (Select ҩƷid, ��ҩ��λid
                 From (Select b.ҩƷid, a.��ҩ��λid, Row_Number() Over(Partition By a.ҩƷid Order By a.������� Desc) Top
                        From ҩƷ�շ���¼ A, ҩƷ��� B
                        Where b.�ϴι�Ӧ��id Is Null And b.ҩƷid = a.ҩƷid And a.���� = 1 And a.������� Is Not Null)
                 Where Top = 1) Loop
    Update ҩƷ��� Set �ϴι�Ӧ��id = r_ҩƷ���.��ҩ��λid Where ҩƷid = r_ҩƷ���.ҩƷid;
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl1_Optional_ҩƷ���_��Ӧ��;
/

--103024:������,2017-03-17,��Ѫ���뵥�в������¼���޸�
--���ڵ���ԭ����Ѫ�����¼.�в�����ֶ�Ϊ�ַ����͵ĺ��������ù�������ɾ����Ѫ�����¼�з����ֶ��в����_bak
Create Or Replace Procedure Zl1_Optional_��Ѫ�����¼_ɾ�� Is
Begin
  If Zl_Checkobject(2, '��Ѫ�����¼', '�в����_bak') > 0 Then
    Execute Immediate 'ALTER TABLE ��Ѫ�����¼ DROP COLUMN �в����_bak';
  End If;
End Zl1_Optional_��Ѫ�����¼_ɾ��;
/
-------------------------------------------------------------------------------------------------------
--����ϵͳ�������İ汾��
-------------------------------------------------------------------------------------------------------
--ϵͳ�汾��
--��ѡ�ű����ø���
--�����汾��
Commit;
