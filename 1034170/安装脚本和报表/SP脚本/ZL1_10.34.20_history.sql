--[��������]1
--[�����߰汾��]10.34.0
--���ű�֧�ִ�ZLHIS+ v10.34.10 ������ v10.34.20
--����ϵͳ�����ߵ�¼PLSQL��ִ�����нű�
--�ű�ִ�к����ֹ�������������
Define n_System=100;
-------------------------------------------------------------------------------
--�ṹ��������
------------------------------------------------------------------------------
--82209:��˶,2015-03-16,�����¼����
alter table ���������¼ add �������� number(2);

--82934:Ƚ����,2015-04-09,����Ԥ����¼������"��������"�ֶΣ�ͬʱ������������
Alter Table ����Ԥ����¼ Add(�������� Number(2));

-------------------------------------------------------------------------------
--������������
------------------------------------------------------------------------------
--82592:Ƚ����,2015-04-09,�շѲ�����ģ�������������Ԥ�������ݴ��������������
--82934:Ƚ����,2015-04-09,����Ԥ����¼������"��������"�ֶΣ�ͬʱ������������
--��ʱ˵��:�����������ű���15������ִ����ɣ����Ի�������:
--1.Ӳ������
--     IBM 3650 M4,cpu E5-2620 2.1GHz,2*6�ˣ�32G�ڴ�
--     V3700�洢,SASӲ��,10K RPM,Raid 10
--2.�������
--     Windows 2008,Oracle 10.2.0.4 64bit
--     ��־�ļ�500M/����Log Buffer����Ϊ500M,PGAΪ9G,SGAΪ�Զ��������25G
--3.���ݻ���
--     XXҽԺ����10�������
--     סԺ���ü�¼1������������ü�¼2ǧ������Ԥ����¼1ǧ1������
Declare
  --���ܣ�����ʹ��Ԥ����ļ�¼
  --���α����ڻ�ȡʹ��Ԥ����Ľ����¼
  Cursor c_�������� Is
    Select ����id, ����Ա���, ����Ա����, �տ�ʱ��, �ɿ���id, ��������
    From (With �����¼ As (Select Distinct ����id
                        From ����Ԥ����¼
                        Where ��¼���� In (1, 11) And ����id Is Not Null And Nvl(��Ԥ��, 0) <> 0)
           Select /*+ FULL(A)*/ a.Id As ����id, Max(a.����Ա���) As ����Ա���, Max(a.����Ա����) As ����Ա����, Max(a.�շ�ʱ��) As �տ�ʱ��, Max(a.�ɿ���id) As �ɿ���id,
                  2 As ��������
           From ���˽��ʼ�¼ A, �����¼ B
           Where a.Id = b.����id
           Group By a.Id
           Union All
           Select /*+ FULL(A)*/ a.����id, Max(a.����Ա���), Max(a.����Ա����), Max(a.�Ǽ�ʱ��), Max(a.�ɿ���id), 5
           From סԺ���ü�¼ A, �����¼ B
           Where a.����id = b.����id And a.���ʷ��� = 0
           Group By a.����id
           Union All
           Select /*+ FULL(A)*/ a.����id, Max(a.����Ա���), Max(a.����Ա����), Max(a.�Ǽ�ʱ��), Max(a.�ɿ���id),
                  Decode(Mod(Max(��¼����), 10), 1, 3, Mod(Max(��¼����), 10))
           From ������ü�¼ A, �����¼ B
           Where a.����id = b.����id And a.���ʷ��� = 0
           Group By a.����id);


  Type t_����id Is Table Of ����Ԥ����¼.����id%Type;
  Type t_�տ�ʱ�� Is Table Of ����Ԥ����¼.�տ�ʱ��%Type;
  Type t_����Ա��� Is Table Of ����Ԥ����¼.����Ա���%Type;
  Type t_����Ա���� Is Table Of ����Ԥ����¼.����Ա����%Type;
  Type t_�ɿ���id Is Table Of ����Ԥ����¼.�ɿ���id%Type;
  Type t_�������� Is Table Of ����Ԥ����¼.��������%Type;
  c_����id     t_����id;
  c_�տ�ʱ��   t_�տ�ʱ��;
  c_����Ա��� t_����Ա���;
  c_����Ա���� t_����Ա����;
  c_�ɿ���id   t_�ɿ���id;
  c_��������   t_��������;
  n_Array_Size Number := 10000; --ÿ����ȡһ�������ID,���˿���PGA����
  I            Number(8) := 0; --ÿ����10�����ID�ύһ��,���˿���Undo����,�����ύ����Ƶ��
  J            Number(16) := 0;
  v_����       Zlupgradeconfig.����%Type;
Begin
  Begin
    Select ���� Into v_���� From Zlupgradeconfig Where ��Ŀ = User || '_����Ԥ����¼����_20150409_1';
  Exception
    When Others Then
      v_���� := Null;
  End;
  If Nvl(v_����, 'RJM') = '�ɹ�' Then
    --�����������ɹ�
    Return;
  End If;

  --��������
  If Zl_Checkobject(1, Null, '����Ԥ����¼_20150409_bak') = 0 Then
    Execute Immediate 'Create Table ����Ԥ����¼_20150409_bak As Select * From ����Ԥ����¼';
  End If;

  Update Zlupgradeconfig Set ���� = Null Where ��Ŀ = User || '_����Ԥ����¼����_20150409_1';
  If Sql%NotFound Then
    Insert Into Zlupgradeconfig (��Ŀ, ����) Values (User || '_����Ԥ����¼����_20150409_1', Null);
  End If;

  Open c_��������;
  Loop
    Fetch c_�������� Bulk Collect
      Into c_����id, c_����Ա���, c_����Ա����, c_�տ�ʱ��, c_�ɿ���id, c_�������� Limit n_Array_Size;
    Exit When c_����id.Count = 0;
  
    --1�ڶ��μ�֮��ʹ��Ԥ����
    Forall K In 1 .. c_����id.Count
      Update ����Ԥ����¼ A
      Set �տ�ʱ�� = Nvl(c_�տ�ʱ��(K), �տ�ʱ��), ����Ա��� = Nvl(c_����Ա���(K), ����Ա���), ����Ա���� = Nvl(c_����Ա����(K), ����Ա����),
          �ɿ���id = Nvl(c_�ɿ���id(K), �ɿ���id), �������� = Nvl(c_��������(K), ��������)
      Where ����id = c_����id(K) And ��¼���� = 11;
  
    --2��һ��ʹ��Ԥ����
    ----2.1����Ԥ����¼
    Forall K In 1 .. c_����id.Count
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ����id, ��ҳid, ����id, �ɿλ, ��λ������, ��λ�ʺ�, ժҪ, ���, ���㷽ʽ, �������, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��,
         ����id, �ɿ�, �Ҳ�, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, У�Ա�־, ��ת��, ��������)
        Select ����Ԥ����¼_Id.Nextval, 11, NO, ʵ��Ʊ��, ��¼״̬, ����id, ��ҳid, ����id, �ɿλ, ��λ������, ��λ�ʺ�, ժҪ, Null, ���㷽ʽ, �������,
               Nvl(c_�տ�ʱ��(K), �տ�ʱ��), Nvl(c_����Ա���(K), ����Ա���), Nvl(c_����Ա����(K), ����Ա����), ��Ԥ��, ����id, �ɿ�, �Ҳ�,
               Nvl(c_�ɿ���id(K), �ɿ���id), Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, У�Ա�־, ��ת��, Nvl(c_��������(K), ��������)
        From ����Ԥ����¼
        Where ����id = c_����id(K) And ��¼���� = 1 And Nvl(��Ԥ��, 0) <> 0;
  
    ----2.2��ԭԤ����¼�ĳ�Ԥ�����Ϊ0
    Forall K In 1 .. c_����id.Count
      Update ����Ԥ����¼
      Set ��Ԥ�� = 0, �������� = Nvl(c_��������(K), ��������)
      Where ����id = c_����id(K) And ��¼���� = 1;
  
    J := J + c_����id.Count;
    If I = 10 Then
      Update Zlupgradeconfig Set ���� = '�Ѵ���' || J || '������ID' Where ��Ŀ = User || '_����Ԥ����¼����_20150409_1';
      Commit;
      I := 0;
    Else
      I := I + 1;
    End If;
  End Loop;
  Update Zlupgradeconfig Set ���� = '��������' || J || '������ID,���ڹر��α�' Where ��Ŀ = User || '_����Ԥ����¼����_20150409_1';
  Commit;
  Close c_��������;

  Update Zlupgradeconfig Set ���� = '�ɹ�' Where ��Ŀ = User || '_����Ԥ����¼����_20150409_1';
  Commit;
End;
/

Declare
  --���ܣ������������ϼ�¼��Ԥ����¼����Ա��һ�µļ�¼
  --���α����ڻ�ȡ�������ϼ�¼��Ԥ����¼����Ա��һ�µĽ������ϼ�¼
  Cursor c_�������� Is
    Select ID, a.����Ա���, a.����Ա����, a.�շ�ʱ��, a.�ɿ���id
    From ���˽��ʼ�¼ A
    Where ��¼״̬ = 2 And Exists (Select 1 From ����Ԥ����¼ Where ����id = a.Id And ����Ա���� <> a.����Ա����);

  Type t_����id Is Table Of ����Ԥ����¼.����id%Type;
  Type t_�տ�ʱ�� Is Table Of ����Ԥ����¼.�տ�ʱ��%Type;
  Type t_����Ա��� Is Table Of ����Ԥ����¼.����Ա���%Type;
  Type t_����Ա���� Is Table Of ����Ԥ����¼.����Ա����%Type;
  Type t_�ɿ���id Is Table Of ����Ԥ����¼.�ɿ���id%Type;
  c_����id     t_����id;
  c_�տ�ʱ��   t_�տ�ʱ��;
  c_����Ա��� t_����Ա���;
  c_����Ա���� t_����Ա����;
  c_�ɿ���id   t_�ɿ���id;
  n_Array_Size Number := 10000; --ÿ����ȡһ�������ID,���˿���PGA����
  I            Number(8) := 0; --ÿ����10�����ID�ύһ��,���˿���Undo����,�����ύ����Ƶ��
  v_����       Zlupgradeconfig.����%Type;
Begin
  Begin
    Select ���� Into v_���� From Zlupgradeconfig Where ��Ŀ = User || '_����Ԥ����¼����_20150409_2';
  Exception
    When Others Then
      v_���� := Null;
  End;
  If Nvl(v_����, 'RJM') = '�ɹ�' Then
    --�����������ɹ�
    Return;
  End If;

  --��������
  If Zl_Checkobject(1, Null, '����Ԥ����¼_20150409_bak') = 0 Then
    Execute Immediate 'Create Table ����Ԥ����¼_20150409_bak As Select * From ����Ԥ����¼';
  End If;

  Update Zlupgradeconfig Set ���� = Null Where ��Ŀ = User || '_����Ԥ����¼����_20150409_2';
  If Sql%NotFound Then
    Insert Into Zlupgradeconfig (��Ŀ, ����) Values (User || '_����Ԥ����¼����_20150409_2', Null);
  End If;

  Open c_��������();
  Loop
    Fetch c_�������� Bulk Collect
      Into c_����id, c_����Ա���, c_����Ա����, c_�տ�ʱ��, c_�ɿ���id Limit n_Array_Size;
    Exit When c_����id.Count = 0;
  
    --�ų�ʹ��Ԥ�����¼����Ϊʹ��Ԥ�������ǰ��������
    Forall K In 1 .. c_����id.Count
      Update ����Ԥ����¼
      Set �տ�ʱ�� = Nvl(c_�տ�ʱ��(K), �տ�ʱ��), ����Ա��� = Nvl(c_����Ա���(K), ����Ա���), ����Ա���� = Nvl(c_����Ա����(K), ����Ա����),
          �ɿ���id = Nvl(c_�ɿ���id(K), �ɿ���id)
      Where ����id = c_����id(K) And ��¼���� Not In (1, 11);
  
    If I = 10 Then
      Commit;
      I := 0;
    Else
      I := I + 1;
    End If;
  End Loop;
  Close c_��������;

  Update Zlupgradeconfig Set ���� = '�ɹ�' Where ��Ŀ = User || '_����Ԥ����¼����_20150409_2';
  Commit;
End;
/

Declare
  --���ܣ��������������ʡ��ֶ�
  --Ԥ������NULL,2-����,3-�շ�,4-�Һ�,5-���￨,6-����ҽ������
  --���α������������������ֶ�,��¼����Ϊ1��11������ǰ������
  Cursor c_�������� Is
    Select Rowid, Mod(��¼����, 10) As �������� From ����Ԥ����¼ Where ��¼���� Not In (1, 11);

  Type t_�������� Is Table Of ����Ԥ����¼.��������%Type;
  c_�������� t_��������;

  c_Rowid      t_Strlist := t_Strlist();
  n_Array_Size Number := 10000; --ÿ��һ��,���˿���PGA����
  I            Number(8) := 0; --ÿ����10������¼�ύһ��,���˿���Undo����,�����ύ����Ƶ��
  J            Number(16) := 0;
  v_����       Zlupgradeconfig.����%Type;
Begin
  Begin
    Select ���� Into v_���� From Zlupgradeconfig Where ��Ŀ = User || '_����Ԥ����¼����_20150409_3';
  Exception
    When Others Then
      v_���� := Null;
  End;
  If Nvl(v_����, 'RJM') = '�ɹ�' Then
    --�����������ɹ�
    Return;
  End If;

  Update Zlupgradeconfig Set ���� = Null Where ��Ŀ = User || '_����Ԥ����¼����_20150409_3';
  If Sql%NotFound Then
    Insert Into Zlupgradeconfig (��Ŀ, ����) Values (User || '_����Ԥ����¼����_20150409_3', Null);
  End If;

  Open c_��������();
  Loop
    Fetch c_�������� Bulk Collect
      Into c_Rowid, c_�������� Limit n_Array_Size;
    Exit When c_Rowid.Count = 0;
  
    Forall K In 1 .. c_Rowid.Count
      Update ����Ԥ����¼ Set �������� = Nvl(c_��������(K), ��������) Where Rowid = c_Rowid(K);
  
    J := J + c_Rowid.Count;
    If I = 10 Then
      Update Zlupgradeconfig Set ���� = '�Ѵ���' || J || '������ID' Where ��Ŀ = User || '_����Ԥ����¼����_20150409_3';
      Commit;
      I := 0;
    Else
      I := I + 1;
    End If;
  End Loop;
  Close c_��������;

  Update Zlupgradeconfig Set ���� = '�ɹ�' Where ��Ŀ = User || '_����Ԥ����¼����_20150409_3';
  Commit;
End;
/



-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------




-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------------------------------
--������ʷ���ݿռ�ϵͳ�汾��
-------------------------------------------------------------------------------------------------------
Update zlBakInfo Set �汾��='10.34.20',��������=Sysdate Where ϵͳ=&n_System;
Commit;