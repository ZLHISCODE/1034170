--[��������]1
--[�����߰汾��]10.34.160
--���ű�֧�ִ�ZLHIS+ v10.34.150 ������ v10.34.160
--����ϵͳ�����ߵ�¼PLSQL��ִ�����нű�
--�ű�ִ�к����ֹ�������������
Define n_System=100;
-------------------------------------------------------------------------------
--�ṹ��������
-------------------------------------------------------------------------------
--128157:����,2018-07-04,������Ŀ��λ����������
Create Index ������Ŀ��λ_IX_��λ on ������Ŀ��λ(��λ,����) Tablespace zl9indexhis;

--119477:����,2018-06-11,�������֮��,�ٴη�����ܵ����µ�

Alter Table �����¼��Ŀ add ������� VARCHAR2(100);

--126057:����,2018-06-06,ҩƷ�շ����������ֶΡ��Է��ⷿid��
Alter Table ҩƷ�շ����� Add �Է�����id Number(18);

Alter Table ҩƷ�շ����� Drop Constraint ҩƷ�շ�����_UQ_NO Cascade drop Index;

Alter Table ҩƷ�շ����� Add Constraint ҩƷ�շ�����_UQ_NO Unique (NO, ����, �ⷿID, �Է�����ID) Using Index Tablespace zl9Indexcis;

--126017:����,2018-05-23,������Ŀ�����鲿λ������ѡ����
Alter Table ������Ŀ��λ Add �ϼ����� Varchar2(30);
Alter Table ������Ŀ��λ Drop Constraint ������Ŀ��λ_UQ Cascade Drop Index;
Alter Table ������Ŀ��λ Add Constraint ������Ŀ��λ_UQ_��Ŀid Unique(��Ŀid,��λ,����,����,�ϼ�����)Using Index Tablespace Zl9indexhis;

--124269:������,2018-05-07,���ڻ��������´�ҽ�����
Alter Table ����ҽ����¼ Add ����ҽ��ID number(18);  

--120692:������,2018-04-17,�����¼֧�ּ�����Ŀ����
create table �������ݵ��붨��
(
��� number(1),
���� varchar2(100),
��ʽ varchar2(500)
)tablespace zl9BaseItem;
alter table �������ݵ��붨�� add constraint �������ݵ��붨��_PK primary key (���) using index tablespace zl9Indexhis;

--111037:��ΰ��,2018-06-24,�������Ǽ�����¼������ʱ��
Alter Table ������������¼ Add ����ʱ�� Date;

--124487:����һ,2018-04-18,�Ӵ󲿷ֱ��������������(ִ�к�����Ѿ����ٵ����ݶ���Ч,�������������(��)��Ч,��Ҫ�ؽ�������Move��)
Declare
  Cursor c_Sql Is
    Select 'Alter table ' || Table_Name || ' Initrans 20' Executesql
    From User_Tables
    Where Ini_Trans = 1 And
          Table_Name In (Select /*+ cardinality(a,10)*/
                          Upper(Column_Value) Tblname
                         From Table(f_Str2list('���Ӳ�����¼,���Ӳ�����ʽ,���Ӳ�������,���Ӳ�������,���Ӳ���ͼ��,���Ӳ�����ӡ,�Һ����״̬', ',')) A)
    Union All
    Select Distinct 'Alter index ' || Index_Name || ' Initrans 20' Executesql
    From User_Indexes
    Where Ini_Trans = 2 And Index_Type = 'NORMAL' And
          Table_Name In (Select /*+ cardinality(a,10)*/
                          Upper(Column_Value)
                         From Table(f_Str2list('���Ӳ�����¼,���Ӳ�����ʽ,���Ӳ�������,���Ӳ�������,���Ӳ���ͼ��,���Ӳ�����ӡ,�Һ����״̬', ',')) A);
  c_Row c_Sql%RowType;
Begin
  For c_Row In c_Sql Loop
    Execute Immediate c_Row.Executesql;
  End Loop;
End;
/

-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--127720:����,2018-06-29,������ҩ���۷���ҩƷ�������Զ��л�����
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1341, 0, 0, 0, 0, 75, '��ҩ���θ���', '0', '0',
         '0-������,1-���á����ú󣬶��۷���ҩƷ�����Ϊ�ϸ���ʱ���ҷ�ҩʱ��ҩƷ��ʵ���������㣬���Զ�Ѱ�ҿ���㹻���������β��滻����'
  From Dual;

--127487:����,2018-06-25,������������Ƿ������Ų���
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, -null, 0, 0, 0, 0, 305, '�����������Ų��ؿ���', '1', '1',
         '���øò�����,���漰���������ĵط������������Ƿ�¼�������źͲ��أ�0-�������������Ƿ�¼�������źͲ��أ�1-�����������Ƿ�¼�������źͲ��ء�'
  From Dual;

--127480:����,2018-06-20,����Oracle����Zl_�ҺŰ���_Autoupdate
Update �ҺŰ��żƻ� Set ʵ����Ч = To_Date('3000-01-01', 'yyyy-mm-dd') Where ��Чʱ�� > Sysdate And ʵ����Ч < Sysdate;

--124184:����,2018-04-13,����Ѫ��ҵ����Ϣ
insert into ҵ����Ϣ����(����,����,˵��,��������) values('ZLHIS_BLOOD_008','�ύ��Ѫ��Ӧ','ҽ����д���ύ����Ѫ��Ӧʱ����������Ϣ',7);

--124195:����,2018-04-13,����Ѫ��ҵ����Ϣ
insert into ҵ����Ϣ����(����,����,˵��,��������) values('ZLHIS_BLOOD_005','ѪҺ������','ѪҺ�����ɺ�ѪҺ���ڴ���״̬,�������Ԥ����Ѫ���ڣ���ʾ��Ӧ����',7);

--124189:����,2018-04-20,����Ѫ��ҵ����Ϣ
insert into ҵ����Ϣ����(����,����,˵��,��������) values('ZLHIS_BLOOD_006','������Ѫ��Ӧ','��ʿִ����Ѫʱ��������Ѫ��Ӧ����ʾ��Ӧҽ��վ',7);

--124187:����,2018-04-24,����Ѫ��ҵ����Ϣ
insert into ҵ����Ϣ����(����,����,˵��,��������) values('ZLHIS_BLOOD_007','Ѫ��������ʾ','��ʿִ����ɺ���ʾ��ʿվ��ҽ��վ����Ѫ��',7);

--127049:����,2018-06-12,���������Ż�
Update Zlparameters
Set ������ = '���ո��ӷ�', ����˵�� = '�������Һŵ�ʱ���ж��Ƿ��ڹҺŷ��õĻ����ϼ��ո��ӷ�'
Where ϵͳ = &n_System And ģ�� = 1802 And ������ = '����ҩ�·����';

Update Zlparameters
Set ������ = '���ո��ӷ�',����˵�� = '������ԤԼ��ʱ���ж��Ƿ��ڹҺŷ��õĻ����ϼ��ո��ӷ�'
Where ϵͳ = &n_System And ģ�� = 1803 And ������ = '����ҩ�·����';

--126057:����,2018-06-06,ҩƷ�շ����������ֶΡ��Է��ⷿid���ƿ�������������
Declare
  v_No         ҩƷ�շ�����.No%Type;
  n_����       ҩƷ�շ�����.����%Type;
  n_�ⷿid     ҩƷ�շ�����.�ⷿid%Type;
  n_�Է�����id ҩƷ�շ�����.�ⷿid%Type;

  Cursor c_ҩƷ�շ����� Is
    Select a.No, a.����, a.�ⷿid From ҩƷ�շ����� A Where a.���� = 19;
Begin
  For r_ҩƷ�շ����� In c_ҩƷ�շ����� Loop
    Begin
      Select �Է�����id
      Into n_�Է�����id
      From ҩƷ�շ���¼
      Where NO = r_ҩƷ�շ�����.No And �ⷿid = r_ҩƷ�շ�����.�ⷿid And ���� = r_ҩƷ�շ�����.���� And ���ϵ�� = -1 And Rownum < 2;
    Exception
      When Others Then
        n_�Է�����id := 0;
    End;
  
    Update ҩƷ�շ�����
    Set �Է�����id = n_�Է�����id
    Where NO = r_ҩƷ�շ�����.No And ���� = r_ҩƷ�շ�����.���� And �ⷿid = r_ҩƷ�շ�����.�ⷿid;

    Commit;
  End Loop;
End;
/

--124269:������,2018-05-07,���ڻ��������´�ҽ�����
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, -null, 0, 0, 0, 0, 302, '��������´�ҽ���ɻ���������Ҵ���', '0',
         '0', '0-��ʾ������,1-��ʾ����,�����ò��������������µ�ҽ�������ɻ���������ҵĻ�ʿ����У�Ի��߷���' 
  From Dual;

--124273:������,2018-05-07,����δ����ҽ��ʱ��ֹ����ת��ҽ��
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1254, 0, 0, 0, 0, 85, '����δ����ҽ��ʱ��ֹ����ת��ҽ��', '0', '0',
         '0-��ʾ������,1-��ʾ����,�����ò��������ڿ��Է��͵ĳ���ҽ��ʱ�ͻ��ֹУ�Ի��߷���ת��ҽ�������������жϵ�ʱ��ֻ�жϳ�����������Խ��(����ҽ������ǰ���δ��Чҽ��)�������ʹ��'
  From Dual;    

--124467:����,2018-04-23,������ҩģ�������Զ���ҩ�Ĺ������
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1341, 0, 1, 0, 0, 70, '�Զ���ҩ����', '0', '0',
         '0-ȫ�������Զ���ҩ;1-���Ӵ���(��ҽ���Ĵ���)�Զ���ҩ;2-�ֹ�����(��ҽ���Ĵ���)�Զ���ҩ'
  From Dual;

--124699:������,2018-04-23,��ʷ����ת��������ģʽ��Ӧ������������ͳһ����
Insert Into zlBakTableindex(ϵͳ,����,������) Select 100,'����Ԥ����¼','����Ԥ����¼_IX_��ҳID' From Dual;

Insert Into zlBakTableindex(ϵͳ,����,������) Select 100,'������ü�¼','������ü�¼_IX_����ID' From Dual;

Insert Into zlBakTableindex(ϵͳ,����,������) Select 100,'����ҽ����¼','����ҽ����¼_IX_����ʱ��' From Dual;

Insert Into zlBakTableindex(ϵͳ,����,������) Select 100,'�������Լ�¼','�������Լ�¼_IX_ҽ��ID' From Dual;

Insert Into zlBakTableindex(ϵͳ,����,������) Select 100,'������ˮ�߱걾','������ˮ�߱걾_IX_�걾ID' From Dual;

Insert Into zlBakTableindex(ϵͳ,����,������) Select 100,'������ˮ��ָ��','������ˮ��ָ��_IX_�걾ID' From Dual;

--120692:������,2018-04-17,�����¼֧�ּ�����Ŀ����
Insert into zlTables(ϵͳ,����,��ռ�,����) Values(&n_System,'�������ݵ��붨��','ZL9BASEITEM','A2');


--123971:������,2018-04-18,���Ӳ�������ѪҺ���պ������ִ�еǼ�
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ,����˵��)
  Select Zlparameters_Id.Nextval, &n_System, -Null, 0, 0, 0, 0, 301, 'ѪҺ���պ������ִ�еǼ�', '1', '1','����Ѫ��ϵͳʱҽ����ԱȡѪ���Һ��Ƿ���Ҫ����ѪҺ���պ˶Ի��ڲ����������Ѫִ������Ǽǣ�0-������н��ջ��ڼ��ɽ���ִ������Ǽ�,1-�������ѪҺ���պ˶Ի��ڲ��������ִ������Ǽ�'
  From Dual;
-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------
--119477:����,2018-07-09,�������֮��,�ٴη�����ܵ����µ�
Insert Into Zlprogprivs(ϵͳ, ���, ����, ������, ����, Ȩ��)Values(&n_System, 1255, '�����¼��ǩ', User, 'Zl_������λ���_Update', 'EXECUTE');

--127340:����,2018-06-29,���ŷ�ҩ����Ȩ����������
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1342,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0
Union All Select 'Zl_Gettransexenumber','EXECUTE' From Dual) A;

--121712:����,2018-05-21,������Ŀ����ֿ���Ȩ
Insert Into Zlprogfuncs (ϵͳ, ���, ����, ����, ˵��, ȱʡֵ) Values (&n_System, 1054, '������Ŀ�༭', 5, '���ӡ�ɾ�����޸�������Ŀ�Ĳ���Ȩ�ޡ��и�Ȩ��ʱ������Է��༰������Ŀ�������ӡ�ɾ�����޸ġ����á�ͣ�ã����������ü�鲿λ���ɼ���ʽ���걾���ա��ų��ϵ����Ӧ����', 1);
Insert Into Zlprogfuncs (ϵͳ, ���, ����, ����, ˵��, ȱʡֵ) Values (&n_System, 1054, '��ҩ�䷽�༭', 7, '���ӡ�ɾ�����޸���ҩ�䷽�Ĳ���Ȩ�ޡ��и�Ȩ��ʱ������Է��༰��ҩ�䷽�������ӡ�ɾ�����޸ġ����á�ͣ��', 1);
Insert Into Zlprogfuncs (ϵͳ, ���, ����, ����, ˵��, ȱʡֵ) Values (&n_System, 1054, '���׷����༭', 11, '���ӡ�ɾ�����޸ĳ��׷����Ĳ���Ȩ�ޡ��и�Ȩ��ʱ������Է��༰���׷����������ӡ�ɾ�����޸ġ����á�ͣ��', 1);
Update Zlprogfuncs Set ���� = 2 Where ϵͳ = &n_System And ��� = 1054 And ���� = '��������';
Update Zlprogfuncs Set ���� = 6 Where ϵͳ = &n_System And ��� = 1054 And ���� = '������ҩ�䷽';
Update Zlprogfuncs Set ���� = 8 Where ϵͳ = &n_System And ��� = 1054 And ���� = '������׷���';
Update Zlprogfuncs Set ���� = 9 Where ϵͳ = &n_System And ��� = 1054 And ���� = 'ȫԺ���׷���';
Update Zlprogfuncs Set ���� = 10 Where ϵͳ = &n_System And ��� = 1054 And ���� = '���Ƴ��׷���';
Update Zlprogfuncs Set ���� = 12 Where ϵͳ = &n_System And ��� = 1054 And ���� = '�޸�ȫԺ���׷���';
Update Zlprogfuncs Set ���� = 13 Where ϵͳ = &n_System And ��� = 1054 And ���� = '�޸Ŀ��ҳ��׷���';
Update Zlprogfuncs Set ���� = 14 Where ϵͳ = &n_System And ��� = 1054 And ���� = '�޸ĸ��˳��׷���';

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1054,'���׷����༭',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0
Union All Select 'ZL_���׷�������_Insert','EXECUTE' From Dual
Union All Select 'ZL_���׷�����Ŀ_UPDATE','EXECUTE' From Dual
Union All Select 'ZL_���Ʒ���Ŀ¼_DELETE','EXECUTE' From Dual
Union All Select 'ZL_���Ʒ���Ŀ¼_Insert','EXECUTE' From Dual
Union All Select 'ZL_���Ʒ���Ŀ¼_UPDATE','EXECUTE' From Dual
Union All Select 'ZL_������Ŀ_DELETE','EXECUTE' From Dual
Union All Select 'ZL_������Ŀ_Insert','EXECUTE' From Dual
Union All Select 'ZL_������Ŀ_REUSE','EXECUTE' From Dual
Union All Select 'ZL_������Ŀ_STOP','EXECUTE' From Dual
Union All Select 'ZL_������Ŀ_UPDATE','EXECUTE' From Dual
Union All Select 'ZL_������Ŀ_DELETE','EXECUTE' From Dual
Union All Select 'ZL_������Ŀ_Insert','EXECUTE' From Dual
Union All Select 'ZL_������Ŀ_UPDATE','EXECUTE' From Dual) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1054,'������Ŀ�༭',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0
Union All Select 'ZL_��������Ŀ_UPDATE','EXECUTE' From Dual
Union All Select 'ZL_���鱨����Ŀ_UPDATE','EXECUTE' From Dual
Union All Select 'ZL_������Ŀ_UPDATE','EXECUTE' From Dual
Union All Select 'ZL_�������_UPDATE','EXECUTE' From Dual
Union All Select 'ZL_������Ŀ_DELETE','EXECUTE' From Dual
Union All Select 'ZL_������Ŀ_Insert','EXECUTE' From Dual
Union All Select 'ZL_������Ŀ_UPDATE','EXECUTE' From Dual
Union All Select 'ZL_�÷�����_UPDATE','EXECUTE' From Dual
Union All Select 'ZL_���Ƶ���Ӧ��_UPDATE','EXECUTE' From Dual
Union All Select 'ZL_���Ʒ���Ŀ¼_DELETE','EXECUTE' From Dual
Union All Select 'ZL_���Ʒ���Ŀ¼_Insert','EXECUTE' From Dual
Union All Select 'ZL_���Ʒ���Ŀ¼_UPDATE','EXECUTE' From Dual
Union All Select 'ZL_���ƻ�����Ŀ_SAVE','EXECUTE' From Dual
Union All Select 'ZL_������Ŀ_DELETE','EXECUTE' From Dual
Union All Select 'ZL_������Ŀ_Insert','EXECUTE' From Dual
Union All Select 'ZL_������Ŀ_REUSE','EXECUTE' From Dual
Union All Select 'ZL_������Ŀ_STOP','EXECUTE' From Dual
Union All Select 'ZL_������Ŀ_UPDATE','EXECUTE' From Dual
Union All Select 'Zl_������Ŀ��λ_DELETE','EXECUTE' From Dual
Union All Select 'Zl_������Ŀ��λ_Insert','EXECUTE' From Dual
Union All Select '������Ŀ','SELECT' From Dual
Union All Select '������Ŀ�ο�','SELECT' From Dual
Union All Select '���Ʒ���Ŀ¼_ID','SELECT' From Dual
Union All Select '������ĿĿ¼_ID','SELECT' From Dual
Union All Select '����������Ŀ_ID','SELECT' From Dual) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1054,'��ҩ�䷽�༭',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0
Union All Select 'ZL_��ҩ�䷽_UPDATE','EXECUTE' From Dual
Union All Select 'ZL_���Ʒ���Ŀ¼_DELETE','EXECUTE' From Dual
Union All Select 'ZL_���Ʒ���Ŀ¼_Insert','EXECUTE' From Dual
Union All Select 'ZL_���Ʒ���Ŀ¼_UPDATE','EXECUTE' From Dual
Union All Select 'ZL_������Ŀ_DELETE','EXECUTE' From Dual
Union All Select 'ZL_������Ŀ_Insert','EXECUTE' From Dual
Union All Select 'ZL_������Ŀ_REUSE','EXECUTE' From Dual
Union All Select 'ZL_������Ŀ_STOP','EXECUTE' From Dual
Union All Select 'ZL_������Ŀ_UPDATE','EXECUTE' From Dual
Union All Select 'ZL_������Ŀ_DELETE','EXECUTE' From Dual
Union All Select 'ZL_������Ŀ_Insert','EXECUTE' From Dual
Union All Select 'ZL_������Ŀ_UPDATE','EXECUTE' From Dual) A;

Insert Into zlProgRelas(ϵͳ,���,���,����,��ϵ,����,�����ϵ)
Select &n_System,1054,1,A.* From (
Select ����,��ϵ,����,�����ϵ From zlProgRelas Where 1 = 0
Union All Select '�޸ĸ��˳��׷���',2,0,0 From Dual) A;

Insert Into zlProgRelas(ϵͳ,���,���,����,��ϵ,����,�����ϵ)
Select &n_System,1054,2,A.* From (
Select ����,��ϵ,����,�����ϵ From zlProgRelas Where 1 = 0
Union All Select '���Ƴ��׷���',2,1,0 From Dual
Union All Select '�޸Ŀ��ҳ��׷���',2,0,0 From Dual) A;

Insert Into zlProgRelas(ϵͳ,���,���,����,��ϵ,����,�����ϵ)
Select &n_System,1054,3,A.* From (
Select ����,��ϵ,����,�����ϵ From zlProgRelas Where 1 = 0
Union All Select 'ȫԺ���׷���',2,1,0 From Dual
Union All Select '�޸�ȫԺ���׷���',2,0,0 From Dual) A;

Insert Into zlProgRelas(ϵͳ,���,���,����,��ϵ,����,�����ϵ)
Select &n_System,1054,4,A.* From (
Select ����,��ϵ,����,�����ϵ From zlProgRelas Where 1 = 0
Union All Select '����������Ŀ',2,1,0 From Dual
Union All Select '������Ŀ�༭',2,0,0 From Dual) A;

Insert Into zlProgRelas(ϵͳ,���,���,����,��ϵ,����,�����ϵ)
Select &n_System,1054,5,A.* From (
Select ����,��ϵ,����,�����ϵ From zlProgRelas Where 1 = 0
Union All Select '������ҩ�䷽',2,1,0 From Dual
Union All Select '��ҩ�䷽�༭',2,0,0 From Dual) A;

Insert Into zlProgRelas(ϵͳ,���,���,����,��ϵ,����,�����ϵ)
Select &n_System,1054,6,A.* From (
Select ����,��ϵ,����,�����ϵ From zlProgRelas Where 1 = 0
Union All Select '������׷���',2,1,0 From Dual
Union All Select '���׷����༭',2,0,0 From Dual) A;

Insert Into zlProgRelas(ϵͳ,���,���,����,��ϵ,����,�����ϵ)
Select &n_System,1054,7,A.* From (
Select ����,��ϵ,����,�����ϵ From zlProgRelas Where 1 = 0
Union All Select '���׷����༭',2,1,0 From Dual
Union All Select '�޸ĸ��˳��׷���',2,0,0 From Dual
Union All Select '�޸Ŀ��ҳ��׷���',2,0,0 From Dual
Union All Select '�޸�ȫԺ���׷���',2,0,0 From Dual) A;

Insert Into Zlrolegrant
  (ϵͳ, ���, ��ɫ, ����)
  Select ϵͳ, ���, ��ɫ, '������Ŀ�༭'
  From Zlrolegrant
  Where ϵͳ = &n_System And ��� = 1054 And ���� = '��Ŀ�༭';
Insert Into Zlrolegrant
  (ϵͳ, ���, ��ɫ, ����)
  Select ϵͳ, ���, ��ɫ, '��ҩ�䷽�༭'
  From Zlrolegrant
  Where ϵͳ = &n_System And ��� = 1054 And ���� = '��Ŀ�༭';
Insert Into Zlrolegrant
  (ϵͳ, ���, ��ɫ, ����)
  Select ϵͳ, ���, ��ɫ, '���׷����༭'
  From Zlrolegrant
  Where ϵͳ = &n_System And ��� = 1054 And ���� = '��Ŀ�༭';

Insert Into zlModuleRelas(ϵͳ,ģ��,����,���ϵͳ,���ģ��,�������,��ع���,ȱʡֵ)
Select &n_System,1252,A.* From (
Select ����,���ϵͳ,���ģ��,�������,��ع���,ȱʡֵ From zlModuleRelas Where 1 = 0
Union All Select NULL,&n_System,1054,0,'���׷����༭',1 From Dual) A;

Insert Into zlModuleRelas(ϵͳ,ģ��,����,���ϵͳ,���ģ��,�������,��ع���,ȱʡֵ)
Select &n_System,1253,A.* From (
Select ����,���ϵͳ,���ģ��,�������,��ع���,ȱʡֵ From zlModuleRelas Where 1 = 0
Union All Select NULL,&n_System,1054,0,'���׷����༭',1 From Dual) A;

Delete Zlmodulerelas Where ϵͳ = &n_System And ģ�� = 1252 And ���� Is Null And ���ϵͳ = &n_System And ���ģ�� = 1054 And ��ع��� = '��Ŀ�༭';
Delete Zlmodulerelas Where ϵͳ = &n_System And ģ�� = 1253 And ���� Is Null And ���ϵͳ = &n_System And ���ģ�� = 1054 And ��ع��� = '��Ŀ�༭';
Delete Zlprogfuncs Where ϵͳ = &n_System And ��� = 1054 And ���� = '��Ŀ�༭';

--120692:������,2018-04-17,�����¼֧�ּ�����Ŀ����
Insert Into Zlprogprivs(ϵͳ, ���, ����, ������, ����, Ȩ��)Values(&n_System, 1255, '����', User, '�������ݵ��붨��', 'SELECT');
Insert Into Zlprogprivs(ϵͳ, ���, ����, ������, ����, Ȩ��) Values (&n_System, 1255, '�����¼�Ǽ�', User, 'Zl_�������ݵ��붨��_Update', 'EXECUTE');
--124418:������,2018-04-17,����LISȨ��
DELETE zlProgFuncs WHERE ϵͳ = &n_System And ��� = 1215 ;

Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��,ȱʡֵ)
Select &n_System,1215,A.* From (
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0
Union All Select '����',-NULL,NULL,1 From Dual
Union All Select '���ձ걾',1,'���ռ������뵥,��ȷ�������˼�����ʱ�䡣',1 From Dual
Union All Select '���ճ���',2,'�Ƿ���Գ����Ѿ����յı걾��',1 From Dual
Union All Select '��˱걾',3,'���Ѿ�����ı걾�������ȷ�ϡ�',1 From Dual
Union All Select 'δ�շ����',4,'�ܹ����δ��ȡ������ط��õļ��鵥��',1 From Dual
Union All Select '���ȡ��',5,'���Ѿ�����˵ı걾���г�������',1 From Dual
Union All Select '�����Ѵ�ӡ�ɻع�',6,'�д�Ȩ�ޣ�����Իع�����˲����Ѵ�ӡ�ı��档',1 From Dual
Union All Select '���ʼ�����',-NULL,NULL,1 From Dual) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1215,'���ձ걾',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0
Union All Select 'ZL_����걾��¼_�걾����','EXECUTE' From Dual
Union All Select 'Zl_������ͨ���_Write','EXECUTE' From Dual
) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1215,'���ճ���',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0
Union All Select 'Zl_����걾��¼_ȡ������','EXECUTE' From Dual
) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1215,'��˱걾',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0
Union All Select 'ZL_������ͨ���_BATCHUPDATE','EXECUTE' From Dual
Union All Select 'ZL_����걾��¼_�������','EXECUTE' From Dual
) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1215,'δ�շ����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0
Union All Select 'ZL_������ͨ���_BATCHUPDATE','EXECUTE' From Dual
Union All Select 'ZL_����걾��¼_�������','EXECUTE' From Dual
) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1215,'���ȡ��',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0
Union All Select 'ZL_����걾��¼_���ȡ��','EXECUTE' From Dual
) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1215,'�����Ѵ�ӡ�ɻع�',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0
Union All Select 'ZL_����걾��¼_���ȡ��','EXECUTE' From Dual
) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1215,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0
Union All Select 'Zl_����ҽ�����_Edit','EXECUTE' From Dual
Union All Select 'Zl_סԺ���ʼ�¼_Verify','EXECUTE' From Dual
Union All Select 'Zl_���鱨�浥_Insert','EXECUTE' From Dual
Union All Select 'Zl_���Ӳ�����ʽ_Insert','EXECUTE' From Dual
Union All Select '���ű�','SELECT' From Dual
Union All Select 'ҩƷ���ľ���','SELECT' From Dual
Union All Select 'ҩƷ���','SELECT' From Dual
Union All Select '����δ�����','SELECT' From Dual
Union All Select '�������','SELECT' From Dual
Union All Select 'ҩƷ�շ���¼','SELECT' From Dual
Union All Select '������ˮ�߱걾','SELECT' From Dual
Union All Select '������ˮ��ָ��','SELECT' From Dual
Union All Select '�����Լ���¼','SELECT' From Dual
Union All Select '����������Ŀ','SELECT' From Dual
Union All Select '����ϸ��','SELECT' From Dual
Union All Select '���鿹������ҩ','SELECT' From Dual
Union All Select '����ϸ��������','SELECT' From Dual
Union All Select '����ҩ�����','SELECT' From Dual
Union All Select '������ͨ���','SELECT' From Dual
Union All Select '����������Ŀ','SELECT' From Dual
Union All Select '����ϲ�����','SELECT' From Dual
Union All Select '���������¼','SELECT' From Dual
Union All Select '���鱨����Ŀ','SELECT' From Dual
Union All Select '������������¼','SELECT' From Dual
Union All Select '��������','SELECT' From Dual
Union All Select 'δ��ҩƷ��¼','SELECT' From Dual
Union All Select '������Ŀ�ֲ�','SELECT' From Dual
Union All Select '����ҽ������','SELECT' From Dual
Union All Select '������ĿĿ¼','SELECT' From Dual
Union All Select '������ҳ','SELECT' From Dual
Union All Select '������Ϣ','SELECT' From Dual
Union All Select 'סԺ���ü�¼','SELECT' From Dual
Union All Select '������ü�¼','SELECT' From Dual
Union All Select '���Ӳ�����¼','SELECT' From Dual
Union All Select '����ҽ������','SELECT' From Dual
Union All Select '�����ļ��б�','SELECT' From Dual
Union All Select '����걾��¼','SELECT' From Dual
Union All Select '���Ӳ�������','SELECT' From Dual
Union All Select '��Ա��','SELECT' From Dual
Union All Select '������Ա','SELECT' From Dual
Union All Select '����ҽ������','SELECT' From Dual
Union All Select '��������Ӧ��','SELECT' From Dual
Union All Select '����ҽ����¼','SELECT' From Dual
Union All Select '���Ӳ�����ʽ','SELECT' From Dual
) A;






-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------






-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--128596:����,2018-07-09,�޷�����¼��Ӥ�����µ�
CREATE OR REPLACE Procedure Zl_���µ���������_Update
(
  �ļ�id_In   In ���˻����ļ�.Id%Type, --���˻����ļ�ID
  ����ʱ��_In In ���˻�������.����ʱ��%Type,
  ��Ŀ���_In In ���˻�����ϸ.��Ŀ���%Type, --������Ŀ����ţ��ǻ�����Ŀ�̶�Ϊ0
  ��¼id_In   In ���˻�������.Id%Type,
  �༭_In     In Number := 0
) As
  n_����         �����ļ��б�.����%Type;
  n_��ʼʱ��     Number;
  n_������     Number;
  n_ʱ����     Number;
  v_��Ժʱ��     ���˱䶯��¼.��ʼʱ��%Type;
  Ncount         Number;
  v_��¼����     ���˻�������.����ʱ��%Type;
  v_���¼�¼     ���˻�����ϸ.��¼����%Type;
  v_��ʼʱ��     Varchar2(20);
  v_����ʱ��     Varchar2(20);
  v_�м�ʱ��     Varchar2(20);
  v_��ʾʱ��pre  Varchar2(20);
  v_��ʾʱ��next Varchar2(20);
  v_��Ժ��ʼʱ�� Varchar2(20);
  v_��Ժ����ʱ�� Varchar2(20);
  v_��ֵ         ���˻�����ϸ.��¼����%Type;
  v_����ʱ��     Varchar2(20);
  n_����         Number(1);
  n_��ϸid       ���˻�����ϸ.Id%Type;
  v_Error        Varchar2(255);
  n_��Ժ         Number(1);
  n_Time         Number(2);
  n_p            Number(2);
  Err_Custom Exception;
  --��ǰʱ����ʾ����������
  Function f_Nowshow
  (
    ��ʼʱ��_In   In Varchar2,
    ����ʱ��_In   In Varchar2,
    �м�ʱ��_In   In Varchar2,
    Id_In         In ���˻�����ϸ.Id%Type,
    �����ļ�id_In In ���˻����ļ�.Id%Type
  ) Return Varchar2 Is
    n_ʱ���   Number;
    n_��ʾ     Number(1);
    v_��¼���� ���˻�����ϸ.��¼����%Type;
    v_ʱ��     Varchar2(20);
  Begin
    n_ʱ��� := -1;
    For r_Temp In (Select g.����ʱ��, f.��¼����, f.��ʾ, f.δ��˵��
                   From ���˻����ļ� B, ���˻������� G, ���˻�����ϸ F
                   Where b.Id = g.�ļ�id And g.Id = f.��¼id And b.Id = �����ļ�id_In And f.��Ŀ��� = 1 And f.��¼���� = 1 And
                         f.��¼��� = 0 And g.����ʱ�� Between To_Date(��ʼʱ��_In, 'YYYY-MM-DD hh24:mi:ss') And
                         To_Date(����ʱ��_In, 'YYYY-MM-DD hh24:mi:ss') And f.Id <> Id_In
                   Order By g.����ʱ��) Loop
      If n_ʱ��� = -1 Then
        n_ʱ���   := Abs((r_Temp.����ʱ�� - To_Date(�м�ʱ��_In, 'YYYY-MM-DD hh24:mi:ss')) * 24 * 60 * 60);
        v_��¼���� := r_Temp.��¼����;
        n_��ʾ     := r_Temp.��ʾ;
        v_ʱ��     := To_Char(r_Temp.����ʱ��, 'YYYY-MM-DD hh24:mi:ss');
      Else
        If r_Temp.��ʾ = 1 Then
          If n_��ʾ = 1 And Abs((r_Temp.����ʱ�� - To_Date(�м�ʱ��_In, 'YYYY-MM-DD hh24:mi:ss')) * 24 * 60 * 60) < n_ʱ��� Then
            n_ʱ���   := Abs((r_Temp.����ʱ�� - To_Date(�м�ʱ��_In, 'YYYY-MM-DD hh24:mi:ss')) * 24 * 60 * 60);
            v_��¼���� := r_Temp.��¼����;
            n_��ʾ     := r_Temp.��ʾ;
            v_ʱ��     := To_Char(r_Temp.����ʱ��, 'YYYY-MM-DD hh24:mi:ss');
          Else
            n_ʱ���   := Abs((r_Temp.����ʱ�� - To_Date(�м�ʱ��_In, 'YYYY-MM-DD hh24:mi:ss')) * 24 * 60 * 60);
            v_��¼���� := r_Temp.��¼����;
            n_��ʾ     := r_Temp.��ʾ;
            v_ʱ��     := To_Char(r_Temp.����ʱ��, 'YYYY-MM-DD hh24:mi:ss');
          End If;
        Else
          If Abs((r_Temp.����ʱ�� - To_Date(�м�ʱ��_In, 'YYYY-MM-DD hh24:mi:ss')) * 24 * 60 * 60) < n_ʱ��� And n_��ʾ = 0 Then
            n_ʱ���   := Abs((r_Temp.����ʱ�� - To_Date(�м�ʱ��_In, 'YYYY-MM-DD hh24:mi:ss')) * 24 * 60 * 60);
            v_��¼���� := r_Temp.��¼����;
            n_��ʾ     := r_Temp.��ʾ;
            v_ʱ��     := To_Char(r_Temp.����ʱ��, 'YYYY-MM-DD hh24:mi:ss');
          End If;
        End If;

      End If;
      If r_Temp.δ��˵�� Is Not Null And r_Temp.��¼���� Is Null Then
        Return Null;
      End If;
    End Loop;
    If v_ʱ�� Is Not Null Then
      Return v_ʱ�� || '|' || v_��¼����;
    Else
      Return Null;
    End If;
  Exception
    When Others Then
      Return Null;
  End f_Nowshow;

Begin
  n_��Ժ := 0;
  If ��Ŀ���_In <> 1 Then
    Return;
  End If;

  If �༭_In = 1 Then
    Update ���˻�����ϸ Set ��¼���� = 1 Where ��¼id = ��¼id_In And ��Ŀ��� = ��Ŀ���_In;
  End If;

  Begin
    Select Max(a.����)
    Into n_����
    From �����ļ��б� A, ���˻����ļ� B
    Where a.���� = 3 And a.���� <> 1 And a.Id = b.��ʽid And b.Id = �ļ�id_In;
  End;

  --��ѯ��Ժʱ��
  If n_���� = 1 Then
    For r_List In (Select c.Ҫ������, c.�����ı�
                   From ���˻����ļ� A, �����ļ��ṹ C, �����ļ��ṹ D
                   Where c.��id = d.Id And d.��id Is Null And d.������� = 1 And a.��ʽid = c.�ļ�id And a.Id = �ļ�id_In
                   Order By c.Id) Loop
      Case r_List.Ҫ������
        When '��ʼʱ��' Then
          n_��ʼʱ�� := To_Number(r_List.�����ı�);
        When '������' Then
          n_������ := To_Number(r_List.�����ı�);
        When 'ʱ����' Then
          n_ʱ���� := To_Number(r_List.�����ı�);
    Else
          v_Error := '';
      End Case;
    End Loop;
  Else
    n_��ʼʱ�� :=  Zl_To_Number(zl_GetSysParameter('���¿�ʼʱ��', 1255)) ;
    n_������ := 6;
    n_ʱ���� := 4;
  End If;

  Select Min(h.��ʼʱ��)
  Into v_��Ժʱ��
  From ���˱䶯��¼ H, ���˻����ļ� B
  Where h.��ʼʱ�� Is Not Null And h.����id = b.����id And h.��ҳid = b.��ҳid And b.Id = �ļ�id_In
  Group By h.����id, h.��ҳid;

  v_��¼���� := To_Date(To_Char(v_��Ժʱ��, 'YYYY-MM-DD'), 'YYYY-MM-DD hh24:mi:ss');
  Ncount     := Floor(((v_��Ժʱ�� - v_��¼����) * 24 - n_��ʼʱ��) / n_ʱ����);

  If Ncount > n_������ Then
    Ncount := n_������;
  End If;
  v_��Ժ��ʼʱ�� := To_Char(v_��¼���� + ((n_��ʼʱ�� + Ncount * n_ʱ���� - (n_ʱ���� / 4)) / 24), 'YYYY-MM-DD hh24:mi:ss');
  v_��Ժ����ʱ�� := To_Char(v_��¼���� + ((n_��ʼʱ�� + Ncount * n_ʱ���� + (n_ʱ���� / 4)) / 24), 'YYYY-MM-DD hh24:mi:ss');
  If v_��Ժʱ�� <= To_Date(v_��Ժ��ʼʱ��, 'YYYY-MM-DD hh24:mi:ss') And
     v_��Ժʱ�� >= To_Date(v_��Ժ��ʼʱ��, 'YYYY-MM-DD hh24:mi:ss') - n_ʱ���� / 4 / 24 Then
    v_��Ժ����ʱ�� := v_��Ժ��ʼʱ��;
    v_��Ժ��ʼʱ�� := To_Char(v_��¼���� + (n_��ʼʱ�� + Ncount * n_ʱ����) / 24, 'YYYY-MM-DD hh24:mi:ss');
    n_��Ժ         := 1;
  Elsif v_��Ժʱ�� <= To_Date(v_��Ժ��ʼʱ��, 'YYYY-MM-DD hh24:mi:ss') + n_ʱ���� / 24 And
        v_��Ժʱ�� >= To_Date(v_��Ժ����ʱ��, 'YYYY-MM-DD hh24:mi:ss') Then
    v_��Ժ��ʼʱ�� := v_��Ժ����ʱ��;
    v_��Ժ����ʱ�� := To_Char(v_��¼���� + ((n_��ʼʱ�� + (Ncount + 1) * n_ʱ����) / 24), 'YYYY-MM-DD hh24:mi:ss');
    n_��Ժ         := 1;
  End If;

  v_��¼���� := To_Date(To_Char(����ʱ��_In, 'YYYY-MM-DD'), 'YYYY-MM-DD hh24:mi:ss');
  Ncount     := Floor(((����ʱ��_In - v_��¼����) * 24 - n_��ʼʱ��) / n_ʱ����);

  If Ncount > n_������ Then
    Ncount := n_������;
  End If;

  --��ǰ��������ʱ���
  v_��ʼʱ�� := To_Char(v_��¼���� + ((n_��ʼʱ�� + Ncount * n_ʱ����) / 24), 'YYYY-MM-DD hh24:mi:ss');
  v_����ʱ�� := To_Char(v_��¼���� + ((n_��ʼʱ�� + n_ʱ���� * (Ncount + 1)) / 24), 'YYYY-MM-DD hh24:mi:ss');
  v_�м�ʱ�� := To_Char(To_Date(v_��ʼʱ��, 'YYYY-MM-DD hh24:mi:ss') +
                    (To_Date(v_����ʱ��, 'YYYY-MM-DD hh24:mi:ss') - To_Date(v_��ʼʱ��, 'YYYY-MM-DD hh24:mi:ss')) / 2,
                    'YYYY-MM-DD hh24:mi:ss');

  Select Max(f.��¼����), Max(f.Id)
  Into v_���¼�¼, n_��ϸid
  From ���˻����ļ� B, ���˻������� G, ���˻�����ϸ F
  Where b.Id = g.�ļ�id And g.Id = f.��¼id And b.Id = �ļ�id_In And f.��Ŀ��� = 1 And f.��¼��� = 0 And g.����ʱ�� = ����ʱ��_In;

  v_��ֵ         := f_Nowshow(v_��ʼʱ��, v_����ʱ��, v_�м�ʱ��, n_��ϸid, �ļ�id_In);
  v_��ʾʱ��next := '';
  While v_��ʾʱ��next Is Null Loop
    If v_��ֵ Is Null Then
      If v_����ʱ�� Is Not Null Then
        v_����ʱ�� := To_Char(To_Date(v_�м�ʱ��, 'YYYY-MM-DD hh24:mi:ss') + n_ʱ���� / 24, 'YYYY-MM-DD hh24:mi:ss');
      Else
        v_����ʱ�� := v_�м�ʱ��;
      End If;
    Else
      n_p        := Instr(v_��ֵ, '|');
      v_����ʱ�� := Substr(v_��ֵ, 1, n_p - 1);
      v_��ֵ     := Substr(v_��ֵ, n_p + 1);
    End If;
    If To_Date(v_����ʱ��, 'YYYY-MM-DD hh24:mi:ss') < ����ʱ��_In Then
      v_��ֵ := f_Nowshow(To_Char(To_Date(v_��ʼʱ��, 'YYYY-MM-DD hh24:mi:ss') + n_ʱ���� / 24, 'YYYY-MM-DD hh24:mi:ss'),
                        To_Char(To_Date(v_����ʱ��, 'YYYY-MM-DD hh24:mi:ss') + n_ʱ���� / 24, 'YYYY-MM-DD hh24:mi:ss'),
                        To_Char(To_Date(v_�м�ʱ��, 'YYYY-MM-DD hh24:mi:ss') + n_ʱ���� / 24, 'YYYY-MM-DD hh24:mi:ss'), n_��ϸid,
                        �ļ�id_In);
    Else
      v_��ʾʱ��next := v_����ʱ��;
    End If;
  End Loop;
  v_��ֵ     := '';
  v_����ʱ�� := '';

  --ѭ����ѯ��ǰʱ��֮ǰ����ͨ����
  n_Time := 0;
  While n_Time * n_ʱ���� <= 24 Loop
    If To_Date(v_��ʼʱ��, 'YYYY-MM-DD hh24:mi:ss') < Trunc(To_Date(v_��ʼʱ��, 'YYYY-MM-DD hh24:mi:ss')) + n_��ʼʱ�� / 24 And
       To_Date(v_��ʼʱ��, 'YYYY-MM-DD hh24:mi:ss') <> Trunc(To_Date(v_��ʼʱ��, 'YYYY-MM-DD hh24:mi:ss')) Then
      v_����ʱ�� := To_Char(Trunc(To_Date(v_��ʼʱ��, 'YYYY-MM-DD hh24:mi:ss')) + n_��ʼʱ�� / 24, 'YYYY-MM-DD hh24:mi:ss');
      v_��ʼʱ�� := To_Char(Trunc(To_Date(v_��ʼʱ��, 'YYYY-MM-DD hh24:mi:ss')), 'YYYY-MM-DD hh24:mi:ss');
      v_�м�ʱ�� := To_Char(To_Date(v_��ʼʱ��, 'YYYY-MM-DD hh24:mi:ss') +
                        (To_Date(v_����ʱ��, 'YYYY-MM-DD hh24:mi:ss') - To_Date(v_��ʼʱ��, 'YYYY-MM-DD hh24:mi:ss')) / 2,
                        'YYYY-MM-DD hh24:mi:ss');
      v_��ֵ     := f_Nowshow(v_��ʼʱ��, v_����ʱ��, v_�м�ʱ��, n_��ϸid, �ļ�id_In);
      v_����ʱ�� := To_Char(Trunc(To_Date(v_����ʱ��, 'YYYY-MM-DD hh24:mi:ss')), 'YYYY-MM-DD hh24:mi:ss');
      v_��ʼʱ�� := To_Char(To_Date(v_����ʱ��, 'YYYY-MM-DD hh24:mi:ss') - n_ʱ���� / 24, 'YYYY-MM-DD hh24:mi:ss');
      v_�м�ʱ�� := To_Char(To_Date(v_����ʱ��, 'YYYY-MM-DD hh24:mi:ss') - n_ʱ���� / 2 / 24, 'YYYY-MM-DD hh24:mi:ss');
    Else
      v_��ֵ     := f_Nowshow(v_��ʼʱ��, v_����ʱ��, v_�м�ʱ��, n_��ϸid, �ļ�id_In);
      v_��ʼʱ�� := To_Char(To_Date(v_��ʼʱ��, 'YYYY-MM-DD hh24:mi:ss') - n_ʱ���� / 24, 'YYYY-MM-DD hh24:mi:ss');
      v_����ʱ�� := To_Char(To_Date(v_����ʱ��, 'YYYY-MM-DD hh24:mi:ss') - n_ʱ���� / 24, 'YYYY-MM-DD hh24:mi:ss');
      v_�м�ʱ�� := To_Char(To_Date(v_�м�ʱ��, 'YYYY-MM-DD hh24:mi:ss') - n_ʱ���� / 24, 'YYYY-MM-DD hh24:mi:ss');

    End If;
    n_Time := n_Time + 1;
    If v_��ֵ Is Not Null Then
      n_p           := Instr(v_��ֵ, '|');
      v_����ʱ��    := Substr(v_��ֵ, 1, n_p - 1);
      v_��ֵ        := Substr(v_��ֵ, n_p + 1);
      v_��ʾʱ��pre := v_����ʱ��;
      Exit When To_Date(v_����ʱ��, 'YYYY-MM-DD hh24:mi:ss') < ����ʱ��_In;
    End If;
  End Loop;
  If v_����ʱ�� Is Not Null Then
    If v_��ֵ < v_���¼�¼ And v_���¼�¼ > 37.5 Then
      Select Count(f.Id)
      Into n_����
      From ���˻����ļ� B, ���˻������� G, ���˻�����ϸ F
      Where b.Id = g.�ļ�id And g.Id = f.��¼id And b.Id = �ļ�id_In And f.��Ŀ��� = ��Ŀ���_In And g.����ʱ�� <> ����ʱ��_In And
            f.��¼���� = 7 And g.����ʱ�� Between To_Date(v_��ʾʱ��pre, 'YYYY-MM-DD hh24:mi:ss') And
            To_Date(v_��ʾʱ��next, 'YYYY-MM-DD hh24:mi:ss');

      If n_���� < 1 Then
        Update ���˻�����ϸ Set ��¼���� = 7 Where ��¼id = ��¼id_In And ��Ŀ��� = ��Ŀ���_In And ��ֹ�汾 Is Null;
      End If;

    End If;
  Else
    If ����ʱ��_In >= To_Date(v_��Ժ��ʼʱ��, 'YYYY-MM-DD hh24:mi:ss') And ����ʱ��_In <= To_Date(v_��Ժ����ʱ��, 'YYYY-MM-DD hh24:mi:ss') And
       n_��Ժ = 1 Then
      Select Count(f.Id)
      Into n_����
      From ���˻����ļ� B, ���˻������� G, ���˻�����ϸ F
      Where b.Id = g.�ļ�id And g.Id = f.��¼id And b.Id = �ļ�id_In And f.��Ŀ��� = ��Ŀ���_In And g.����ʱ�� <> ����ʱ��_In And
            f.��¼���� = 7 And g.����ʱ�� Between To_Date(v_��Ժ��ʼʱ��, 'YYYY-MM-DD hh24:mi:ss') And
            To_Date(v_��Ժ����ʱ��, 'YYYY-MM-DD hh24:mi:ss');

      If n_���� < 1 Then
        Update ���˻�����ϸ Set ��¼���� = 7 Where ��¼id = ��¼id_In And ��Ŀ��� = ��Ŀ���_In And ��ֹ�汾 Is Null;
      End If;
    End If;
  End If;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���µ���������_Update;
/
--128391:����,2018-07-05,�������״̬Ϊ2�����
CREATE OR REPLACE PROCEDURE Zl_ҽ����˹���_Update
(
  ҽ��id_In   ����ҽ��״̬.ҽ��id%TYPE,
  ����ʱ��_In ����ҽ��״̬.����ʱ��%TYPE,
  ����˵��_In ����ҽ��״̬.����˵��%TYPE := NULL,
  ��˶���_In NUMBER := 1, --1=����ҽ����2=��Ѫҽ�� 
  ������Ա_In VARCHAR2 := NULL
) IS
  --�޸�ֻ��������˲�ͨ����ҽ�����޸������˵�� 
  Err_Item EXCEPTION;
  v_Err_Msg  VARCHAR2(200);
  n_Count    NUMBER;
  n_���״̬ NUMBER;
BEGIN
  SELECT COUNT(1) INTO n_Count FROM ����ҽ����¼ WHERE Id = ҽ��id_In;
  SELECT ���״̬ INTO n_���״̬ FROM ����ҽ����¼ WHERE Id = ҽ��id_In;
  IF n_Count = 0 THEN
    v_Err_Msg := '��ҽ���Ѿ�ɾ��,���֤��';
    RAISE Err_Item;
  END IF;

  IF ��˶���_In = 1 THEN
    UPDATE ����ҽ��״̬
    SET ����ʱ�� = ����ʱ��_In, ����˵�� = ����˵��_In
    WHERE ҽ��id IN (SELECT Id FROM ����ҽ����¼ WHERE Id = ҽ��id_In OR ���id = ҽ��id_In) AND �������� = 12 AND
          ����ʱ�� = (SELECT MAX(����ʱ��) FROM ����ҽ��״̬ WHERE ҽ��id = ҽ��id_In AND �������� = 12);
  ELSE
    IF n_���״̬ = 1 THEN
      UPDATE ����ҽ��״̬
      SET ����ʱ�� = ����ʱ��_In, ����˵�� = ����˵��_In
      WHERE ҽ��id IN (SELECT Id FROM ����ҽ����¼ WHERE Id = ҽ��id_In OR ���id = ҽ��id_In) AND �������� = 19 AND
            ����ʱ�� = (SELECT MAX(����ʱ��) FROM ����ҽ��״̬ WHERE ҽ��id = ҽ��id_In AND �������� = 19);
      IF SQL%NOTFOUND THEN
        INSERT INTO ����ҽ��״̬
          (ҽ��id, ��������, ������Ա, ����ʱ��, ����˵��)
          SELECT Id, 19, ������Ա_In, ����ʱ��_In, ����˵��_In
          FROM ����ҽ����¼
          WHERE Id = ҽ��id_In OR ���id = ҽ��id_In;
      END IF;
    ELSIF n_���״̬ = 7 THEN
      UPDATE ����ҽ��״̬
      SET ����ʱ�� = ����ʱ��_In, ����˵�� = ����˵��_In
      WHERE ҽ��id IN (SELECT Id FROM ����ҽ����¼ WHERE Id = ҽ��id_In OR ���id = ҽ��id_In) AND �������� = 18 AND
            ����ʱ�� = (SELECT MAX(����ʱ��) FROM ����ҽ��״̬ WHERE ҽ��id = ҽ��id_In AND �������� = 18);
      IF SQL%NOTFOUND THEN
        INSERT INTO ����ҽ��״̬
          (ҽ��id, ��������, ������Ա, ����ʱ��, ����˵��)
          SELECT Id, 18, ������Ա_In, ����ʱ��_In, ����˵��_In
          FROM ����ҽ����¼
          WHERE Id = ҽ��id_In OR ���id = ҽ��id_In;
      END IF;
    ELSIF n_���״̬ = 3 THEN
      UPDATE ����ҽ��״̬
      SET ����ʱ�� = ����ʱ��_In, ����˵�� = ����˵��_In
      WHERE ҽ��id IN (SELECT Id FROM ����ҽ����¼ WHERE Id = ҽ��id_In OR ���id = ҽ��id_In) AND �������� = 12 AND
            ����ʱ�� = (SELECT MAX(����ʱ��) FROM ����ҽ��״̬ WHERE ҽ��id = ҽ��id_In AND �������� = 12);
      IF SQL%NOTFOUND THEN
        INSERT INTO ����ҽ��״̬
          (ҽ��id, ��������, ������Ա, ����ʱ��, ����˵��)
          SELECT Id, 12, ������Ա_In, ����ʱ��_In, ����˵��_In
          FROM ����ҽ����¼
          WHERE Id = ҽ��id_In OR ���id = ҽ��id_In;
      END IF;
	ELSIF n_���״̬ = 4 OR n_���״̬ = 2 THEN
      UPDATE ����ҽ��״̬
      SET ����ʱ�� = ����ʱ��_In, ����˵�� = ����˵��_In
      WHERE ҽ��id IN (SELECT Id FROM ����ҽ����¼ WHERE Id = ҽ��id_In OR ���id = ҽ��id_In) AND �������� = 11 AND
            ����ʱ�� = (SELECT MAX(����ʱ��) FROM ����ҽ��״̬ WHERE ҽ��id = ҽ��id_In AND �������� = 11);
      IF SQL%NOTFOUND THEN
        INSERT INTO ����ҽ��״̬
          (ҽ��id, ��������, ������Ա, ����ʱ��, ����˵��)
          SELECT Id, 11, ������Ա_In, ����ʱ��_In, ����˵��_In
          FROM ����ҽ����¼
          WHERE Id = ҽ��id_In OR ���id = ҽ��id_In;
      END IF;
    END IF;
  END IF;
EXCEPTION
  WHEN Err_Item THEN
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  WHEN OTHERS THEN
    Zl_Errorcenter(SQLCODE, SQLERRM);
END Zl_ҽ����˹���_Update;
/

--128391:����,2018-07-05,���Ӵ�������״̬Ϊ2�����
CREATE OR REPLACE PROCEDURE Zl_ҽ����˹���_Cancel
(
  ҽ��ids_In  VARCHAR2,
  ��˶���_In NUMBER := 1, --1=����ҽ����2=��Ѫҽ��
  ִ�����_In NUMBER := 0 --0=�ϰ�Ѫ�����̣���Ϊ0ʱ����ΪĿ�����״̬��1=����ˣ�7=��ǩ����4-��ǩ����3-�Ѿܾ���
) IS
  --ȡ�����
  CURSOR c_Advice IS
    SELECT * FROM TABLE(CAST(f_Num2list(ҽ��ids_In) AS t_Numlist));
  Err_Item EXCEPTION;
  v_Err_Msg  VARCHAR2(200);
  n_Count    NUMBER;
  n_ҽ��״̬ NUMBER;
  n_���״̬ NUMBER;
  n_�������� NUMBER;
BEGIN
  FOR r_Advice IN c_Advice LOOP
    SELECT COUNT(1), MAX(ҽ��״̬), Nvl(MAX(���״̬), 0)
    INTO n_Count, n_ҽ��״̬, n_���״̬
    FROM ����ҽ����¼
    WHERE Id = r_Advice.Column_Value;
  
    IF n_Count = 0 THEN
      v_Err_Msg := '��ҽ���Ѿ�ɾ��,���֤��';
      RAISE Err_Item;
    END IF;
  
    IF n_ҽ��״̬ <> 1 THEN
      v_Err_Msg := '��ѡ���ҽ���а�����У�Ե�ҽ��������ȡ����ˡ�';
      RAISE Err_Item;
    END IF;
  
    IF n_���״̬ = 1 THEN
      n_�������� := 19;
    ELSIF n_���״̬ = 7 THEN
      n_�������� := 18;
    ELSIF n_���״̬ = 3 THEN
      n_�������� := 12;
    ELSIF n_���״̬ = 4 OR n_���״̬ = 2 THEN
      n_�������� := 11;
    END IF;
  
    IF ��˶���_In = 1 OR ִ�����_In = 0 THEN
      UPDATE ����ҽ����¼ SET ���״̬ = 1 WHERE Id = r_Advice.Column_Value OR ���id = r_Advice.Column_Value;
      DELETE FROM ����ҽ��״̬
      WHERE ҽ��id IN (SELECT Id FROM ����ҽ����¼ WHERE Id = r_Advice.Column_Value OR ���id = r_Advice.Column_Value) AND
            �������� IN (11, 12) AND
            ����ʱ�� = (SELECT MAX(����ʱ��) FROM ����ҽ��״̬ WHERE ҽ��id = r_Advice.Column_Value AND �������� IN (11, 12));
    ELSIF ��˶���_In = 2 AND ִ�����_In <> 0 THEN
      UPDATE ����ҽ����¼
      SET ���״̬ = ִ�����_In
      WHERE Id = r_Advice.Column_Value OR ���id = r_Advice.Column_Value;
      DELETE FROM ����ҽ��״̬
      WHERE ҽ��id IN (SELECT Id FROM ����ҽ����¼ WHERE Id = r_Advice.Column_Value OR ���id = r_Advice.Column_Value) AND
            �������� = n_�������� AND
            ����ʱ�� =
            (SELECT MAX(����ʱ��) FROM ����ҽ��״̬ WHERE ҽ��id = r_Advice.Column_Value AND �������� = n_��������);
    END IF;
  
  END LOOP;
EXCEPTION
  WHEN Err_Item THEN
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  WHEN OTHERS THEN
    Zl_Errorcenter(SQLCODE, SQLERRM);
END Zl_ҽ����˹���_Cancel;
/

--128152:������,2018-07-05,���֤��������У��
Create Or Replace Function Zl_Fun_Checkidcard
(
  Idcard_In   In Varchar2,
  Calcdate_In In Date := Null
) Return Varchar2 Is
  ------------------------------------------------------------------------------- 
  --���ܣ����֤����Ϸ���У��,���������֤�ŵĳ������ڡ��Ա����� 
  --����˵��: 
  -- ��� IDcard_In:���֤���� 
  --    Calcdate_In:��������,ȱʡʱ��ϵͳʱ�� 
  -- ����ֵ���̶���ʽXML�� 
  --<OUTPUT> 
  --       <BIRTHDAY></BIRTHDAY>                //�������� 
  --       <SEX></SEX>                  //�Ա� 
  --       <AGE></AGE>                //���� 
  --     <MSG></MSG>         //�մ�-���֤����Ч(�ɴ����֤���л�ȡ�������ں��Ա�)���ǿմ�-���ش�����Ϣ 
  --</OUTPUT> 
  ------------------------------------------------------------------------------- 
  n_Count     Number(5);
  n_Sum       Number(5);
  v_У��λ    Varchar2(50);
  v_Pattern   Varchar2(500);
  v_Err_Msg   Varchar2(2000);
  v_�Ա�      Varchar2(100);
  v_����      Varchar2(100);
  d_Curr_Time Date;
  d_��������  Date;
  v_Temp      Varchar2(20);

Begin
  Select Sysdate Into d_Curr_Time From Dual;

  If Idcard_In Is Null Then
    v_Err_Msg := '�������֤��Ϊ��!';
    Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
  Else
    --���֤�Ϸ���֤ 
    v_Pattern := '11,12,13,14,15,21,22,23,31,32,33,34,35,36,37,41,42,43,44,45,46,50,51,52,53,54,61,62,63,64,65,71,81,82,91';
    --�������� 
    If Instr(v_Pattern, Substr(Idcard_In, 1, 2)) = 0 Then
      v_Err_Msg := '���֤ǰ��λ�����벻��ȷ!';
      Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
    End If;
    --���֤���ȼ�� 
    If Length(Idcard_In) = 15 Then
      --������֤��:15λ���֤��Ҫ��ȫ��Ϊ���� 
      v_Pattern := '^\d{15}$';
      Select Count(1) Into n_Count From Dual Where Regexp_Like(Idcard_In, v_Pattern);
      If n_Count = 0 Then
        v_Err_Msg := '���֤�а����Ƿ��ַ�������!';
        Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
      End If;
      --��ȡ�Ա� 
      If Mod(To_Number(Substr(Idcard_In, 15, 1)), 2) = 1 Then
        v_�Ա� := '��';
      Else
        v_�Ա� := 'Ů';
      End If;
      --�������ڵĺϷ��Լ�� 
      v_Pattern := '^19[0-9]{2}((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|[1-2][0-9]))$';
      Select Count(1) Into n_Count From Dual Where Regexp_Like('19' || Substr(Idcard_In, 7, 6), v_Pattern);
      If n_Count = 0 Then
        v_Err_Msg := '���֤�еĳ���������Ч������!';
        Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
      Else
        Begin
          d_�������� := To_Date('19' || Substr(Idcard_In, 7, 6), 'yyyy-mm-dd');
        Exception
          When Others Then
            --��ǰ�������֤û�����������������ݴ����������ڸ�Ϊ2��28�ţ��磺19470229�������
            v_Temp := '19' || Substr(Idcard_In, 7, 6);
            If Instr(v_Temp || ',', '0229,') > 0 Then
              v_Temp := '19' || Substr(Idcard_In, 7, 5) || '8';
            End If;
            d_�������� := To_Date(v_Temp, 'yyyy-mm-dd');
        End;
        If d_�������� > To_Date(To_Char(d_Curr_Time, 'YYYY-MM-DD'), 'YYYY-MM-dd') Then
          v_Err_Msg := '���֤�еĳ���������Ч������!';
          Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
        End If;
      End If;
    Elsif Length(Idcard_In) = 18 Then
      -- 18 λ���֤��ǰ17 λȫ��Ϊ���֣����1λ��Ϊ���ֻ�x 
      v_Pattern := '^\d{17}[0-9Xx]$';
      Select Count(1) Into n_Count From Dual Where Regexp_Like(Idcard_In, v_Pattern);
      If n_Count = 0 Then
        v_Err_Msg := '���֤�а����Ƿ��ַ�!';
        Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
      End If;
      --��ȡ�Ա� 
      If Mod(To_Number(Substr(Idcard_In, 17, 1)), 2) = 1 Then
        v_�Ա� := '��';
      Else
        v_�Ա� := 'Ů';
      End If;
      --�������ڵĺϷ��Լ�� 
      v_Pattern := '^(1[6-9]|[2-9][0-9])[0-9]{2}((01|03|05|07|08|10|12)(0[1-9]|[1-2][0-9]|3[0-1])|(04|06|09|11)(0[1-9]|[1-2][0-9]|30)|02(0[1-9]|[1-2][0-9]))$';
      Select Count(1) Into n_Count From Dual Where Regexp_Like(Substr(Idcard_In, 7, 8), v_Pattern);
      If n_Count = 0 Then
        v_Err_Msg := '���֤�еĳ���������Ч������!';
        Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
      Else
        Begin
          d_�������� := To_Date(Substr(Idcard_In, 7, 8), 'yyyy-mm-dd');
        Exception
          When Others Then
            --��ǰ�������֤û�����������������ݴ����������ڸ�Ϊ2��28�ţ��磺19470229�������
            v_Temp := Substr(Idcard_In, 7, 8);
            If Instr(v_Temp || ',', '0229,') > 0 Then
              v_Temp := Substr(Idcard_In, 7, 7) || '8';
            End If;
            d_�������� := To_Date(v_Temp, 'yyyy-mm-dd');
        End;
      
        If d_�������� > To_Date(To_Char(d_Curr_Time, 'YYYY-MM-DD'), 'YYYY-MM-dd') Then
          v_Err_Msg := '���֤�еĳ���������Ч������!';
          Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
        End If;
        --����У��λ 
        n_Sum     := (To_Number(Substr(Idcard_In, 1, 1)) + To_Number(Substr(Idcard_In, 11, 1))) * 7 +
                     (To_Number(Substr(Idcard_In, 2, 1)) + To_Number(Substr(Idcard_In, 12, 1))) * 9 +
                     (To_Number(Substr(Idcard_In, 3, 1)) + To_Number(Substr(Idcard_In, 13, 1))) * 10 +
                     (To_Number(Substr(Idcard_In, 4, 1)) + To_Number(Substr(Idcard_In, 14, 1))) * 5 +
                     (To_Number(Substr(Idcard_In, 5, 1)) + To_Number(Substr(Idcard_In, 15, 1))) * 8 +
                     (To_Number(Substr(Idcard_In, 6, 1)) + To_Number(Substr(Idcard_In, 16, 1))) * 4 +
                     (To_Number(Substr(Idcard_In, 7, 1)) + To_Number(Substr(Idcard_In, 17, 1))) * 2 +
                     To_Number(Substr(Idcard_In, 8, 1)) * 1 + To_Number(Substr(Idcard_In, 9, 1)) * 6 +
                     To_Number(Substr(Idcard_In, 10, 1)) * 3;
        n_Count   := Mod(n_Sum, 11);
        v_Pattern := '10X98765432';
        v_У��λ  := Substr(v_Pattern, n_Count + 1, 1);
        If v_У��λ <> Upper(Substr(Idcard_In, 18, 1)) Then
          v_Err_Msg := '���֤���벻��ȷ�����顣';
          Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
        End If;
      End If;
    Else
      v_Err_Msg := '���֤���Ȳ���,���顣';
      Return '<OUTPUT><BIRTHDAY></BIRTHDAY><SEX></SEX><AGE></AGE><MSG>' || v_Err_Msg || '</MSG></OUTPUT>';
    End If;
    v_���� := Zl_Age_Calc(0, d_��������, Calcdate_In);
  End If;

  Return '<OUTPUT><BIRTHDAY>' || To_Char(d_��������, 'YYYY-MM-DD') || '</BIRTHDAY><SEX>' || v_�Ա� || '</SEX><AGE>' || v_���� || '</AGE><MSG></MSG></OUTPUT>';
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Fun_Checkidcard;
/

--127817:��¶¶,2018-07-03,������ҳ��ȡ����ʱ������ȡ���������
CREATE OR REPLACE Function Zl_Adderss_Structure(v_Addressinfo Varchar2,n_Type Number :=Null) Return Varchar2 Is
  --���ؽṹ��ʡ,ʡ����,�Ƿ�����,�Ƿ���ʾ,�Ƿ�ֻ�����⼶|��,�б���,�Ƿ�����,�Ƿ���ʾ,�Ƿ�ֻ�����⼶
  --          |����,���ر���,�Ƿ�����,�Ƿ���ʾ,�Ƿ�ֻ�����⼶|����,�������,�Ƿ�����,�Ƿ���ʾ,�Ƿ�ֻ�����⼶
  --          |�ֵ�,�ֵ�����,�Ƿ�����,�Ƿ���ʾ,�Ƿ�ֻ�����⼶
  v_ʡ       Varchar2(100);
  v_Codeʡ   Varchar2(15);
  v_Infoʡ   Varchar2(150);
  v_��       Varchar2(100);
  v_Code��   Varchar2(15);
  v_Info��   Varchar2(150);
  v_����     Varchar2(100);
  v_Code���� Varchar2(15);
  v_Info���� Varchar2(150);
  v_����     Varchar2(100);
  v_Code���� Varchar2(15);
  v_Info���� Varchar2(150);
  v_�ֵ�     Varchar2(500);
  v_Code�ֵ� Varchar2(15);
  v_Info�ֵ� Varchar2(550);
  v_Tmp      Varchar2(100);
  v_Adrstmp  Varchar2(500);
  n_Pos      Number(5);
  n_����     Number(1);
  n_����ʾ   Number(1);
  n_Count    Number(3);
  v_Return   Varchar2(700);
Begin
  --����ṹ���ĵ�ַ�����ý��е�ַ��׼���ָ����
  v_Adrstmp := v_Addressinfo;
  If v_Addressinfo Like '%,%,%,%,%' Then
    n_Pos     := Instr(v_Adrstmp, ',');
    v_ʡ      := Substr(v_Adrstmp, 1, n_Pos - 1);
    v_Adrstmp := Substr(v_Adrstmp, n_Pos + 1);
    n_Pos     := Instr(v_Adrstmp, ',');
    v_��      := Substr(v_Adrstmp, 1, n_Pos - 1);
    v_Adrstmp := Substr(v_Adrstmp, n_Pos + 1);
    n_Pos     := Instr(v_Adrstmp, ',');
    v_����    := Substr(v_Adrstmp, 1, n_Pos - 1);
    v_Adrstmp := Substr(v_Adrstmp, n_Pos + 1);
    n_Pos     := Instr(v_Adrstmp, ',');
    v_����    := Substr(v_Adrstmp, 1, n_Pos - 1);
    v_�ֵ�    := Substr(v_Adrstmp, n_Pos + 1);
    Select Max(����) Into v_Codeʡ From ���� Where ���� = v_ʡ And Nvl(����, 0) = 0;
    --ʡ����ַ��û�У��Ͳ�������
    If v_Codeʡ Is Not Null Then
      Select Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ)
      Into v_Code��, n_����, n_����ʾ
      From ����
      Where ���� = v_�� And Nvl(����, 0) = 1 And �ϼ����� = v_Codeʡ;
      If v_Code�� Is Not Null Then
        v_Info�� := v_�� || ',' || v_Code�� || ',' || n_���� || ',' || n_����ʾ;
        Select Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ)
        Into v_Code����, n_����, n_����ʾ
        From ����
        Where ���� = v_���� And Nvl(����, 0) = 2 And �ϼ����� = v_Code��;
        --�����������ַ
      Else
        Select Max(����), Max(�ϼ�����)
        Into v_Code����, v_Code��
        From ����
        Where ���� = v_���� And Nvl(����, 0) = 2 And �ϼ����� In (Select ���� From ���� Where �ϼ����� = v_Codeʡ);
        If v_Code�� Is Not Null Then
          Select Max(����), Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ)
          Into v_��, v_Code��, n_����, n_����ʾ
          From ����
          Where ���� = v_Code��;
        End If;
        v_Info�� := v_�� || ',' || v_Code�� || ',' || n_���� || ',' || n_����ʾ;
        Select Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ)
        Into v_Code����, n_����, n_����ʾ
        From ����
        Where ���� = v_Code����;
      End If;
      v_Info���� := v_���� || ',' || v_Code���� || ',' || n_���� || ',' || n_����ʾ;
      If v_Code���� Is Not Null Then
        --������������ϸ��ַ�У��������������ַ�ṹ��¼��
        Select Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ)
        Into v_Code����, n_����, n_����ʾ
        From ����
        Where ���� = v_���� And Nvl(����, 0) = 3 And �ϼ����� = v_Code����;
        --�����������ַ
        If v_Code���� Is Null Then
          Select Max(����), Max(�ϼ�����)
          Into v_Code�ֵ�, v_Code����
          From ����
          Where ���� = v_�ֵ� And Nvl(����, 0) = 4 And �ϼ����� In (Select ���� From ���� Where �ϼ����� = v_Code����);
          If v_Code���� Is Not Null Then
            Select Max(����), Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ)
            Into v_����, v_Code����, n_����, n_����ʾ
            From ����
            Where ���� = v_Code����;
          End If;
        End If;
        v_Info���� := v_���� || ',' || v_Code���� || ',' || n_���� || ',' || n_����ʾ;
        If v_Code���� Is Not Null Then
          Select Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ)
          Into v_Code�ֵ�, n_����, n_����ʾ
          From ����
          Where ���� = v_�ֵ� And Nvl(����, 0) = 4 And �ϼ����� = v_Code����;
          v_Info�ֵ� := v_�ֵ� || ',' || v_Code�ֵ� || ',' || n_���� || ',' || n_����ʾ;
        End If;
      End If;
    End If;
    --�Ǳ�׼��ַ����������ַ����Ҫ�ָ�ʡ���У���,
  Else
    v_Adrstmp := v_Addressinfo;
    v_Tmp     := Substr(v_Adrstmp, 1, 2);
    Select Max(����), Max(����) Into v_ʡ, v_Codeʡ From ���� Where ���� Like v_Tmp || '%' And Nvl(����, 0) = 0;
    --��ʡ����ַ��˵�����Խṹ��
    If v_Codeʡ Is Not Null Then
      --ʡ����ַ�Ǳ�׼��
      If Substr(v_Adrstmp, 1, Length(v_ʡ)) = v_ʡ Then
        v_Adrstmp := Substr(v_Adrstmp, Length(v_ʡ) + 1);
        --ʡ����ַ����׼,�����½�ʡ����������,��ʱ���м���ַ�����Ǳ�׼���ġ�
      Else
        --���ж϶�����ַ�Ƿ���������ַ�벻��ʾ�ĵ�ַ
        If v_Tmp = '����' Then
          v_Tmp := '���ɹ�';
        Elsif v_Tmp = '����' Then
          v_Tmp := '������';
        End If;
        v_Adrstmp := Substr(v_Adrstmp, Length(v_Tmp) + 1);
      End If;
      --�Ƚ�ȡ�м������������ؼ��֣���ƥ��
      v_Tmp := Substr(v_Adrstmp, 1, 2);
      If n_Type is Null Then
        Select Max(����), Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ), Count(1) 
        Into v_��, v_Code��, n_����, n_����ʾ, n_Count 
        From ���� 
        Where ���� Like v_Tmp || '%' And Nvl(����, 0) = 1 And �ϼ����� = v_Codeʡ; 
      End If;
      --���ڶ���ƥ�䣬��������ӳ��ȣ���ʱ��2�������ӵ�3����ƥ��
      If n_Count > 1 Then
        v_Tmp := Substr(v_Adrstmp, 1, 3);
        Select Max(����), Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ)
        Into v_��, v_Code��, n_����, n_����ʾ
        From ����
        Where ���� Like v_Tmp || '%' And Nvl(����, 0) = 1 And �ϼ����� = v_Codeʡ;
      End If;
      --�ж��Ƿ���������ַ����ʾ�ĵ�ַ���µ�,������ڣ�����ݵ�������ַ��ȷ�������ַ
      --������û�еڶ����������Ҫ�������ж�
      If v_Code�� Is Null Then
        Select Max(����), Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ), Count(1), Max(�ϼ�����)
        Into v_����, v_Code����, n_����, n_����ʾ, n_Count, v_Code��
        From ����
        Where ���� Like v_Tmp || '%' And Nvl(����, 0) = 2 And �ϼ����� In (Select ���� From ���� Where �ϼ����� = v_Codeʡ);
        --���ڶ���ƥ�䣬��������ӳ��ȣ���ʱ��2�������ӵ�3����ƥ��
        If n_Count > 1 Then
          v_Tmp := Substr(v_Adrstmp, 1, 3);
          Select Max(����), Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ), Max(�ϼ�����)
          Into v_����, v_Code����, n_����, n_����ʾ, v_Code��
          From ����
          Where ���� Like v_Tmp || '%' And Nvl(����, 0) = 2 And �ϼ����� In (Select ���� From ���� Where �ϼ����� = v_Codeʡ);
        End If;
        If v_Code�� Is Not Null Then
          v_Info���� := v_���� || ',' || v_Code���� || ',' || n_���� || ',' || n_����ʾ;
          v_Adrstmp  := Substr(v_Adrstmp, Length(v_����) + 1);
          Select Max(����), Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ)
          Into v_��, v_Code��, n_����, n_����ʾ
          From ����
          Where ���� = v_Code��;
          v_Info�� := v_�� || ',' || v_Code�� || ',' || n_���� || ',' || n_����ʾ;
        End If;
      Else
        v_Info��  := v_�� || ',' || v_Code�� || ',' || n_���� || ',' || n_����ʾ;
        v_Adrstmp := Substr(v_Adrstmp, Length(v_��) + 1);
      End If;
      --û�����أ����������
      If Not v_Code�� Is Null And v_Code���� Is Null Then
        --�Ƚ�ȡ�ؼ������������ؼ��֣���ƥ��
        v_Tmp := Substr(v_Adrstmp, 1, 2);
        Select Max(����), Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ), Count(1)
        Into v_����, v_Code����, n_����, n_����ʾ, n_Count
        From ����
        Where ���� Like v_Tmp || '%' And Nvl(����, 0) = 2 And �ϼ����� = v_Code��;
        --���ڶ���ƥ�䣬��������ӳ��ȣ���ʱ��2�������ӵ�3����ƥ��
        If n_Count > 1 Then
          v_Tmp := Substr(v_Adrstmp, 1, 3);
          Select Max(����), Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ)
          Into v_����, v_Code����, n_����, n_����ʾ
          From ����
          Where ���� Like v_Tmp || '%' And Nvl(����, 0) = 2 And �ϼ����� = v_Code��;
        End If;
        If v_Code���� Is Null Then
          Select Max(�Ƿ�����), Max(�Ƿ���ʾ) Into n_����, n_����ʾ From ���� Where �ϼ����� = v_Code��;
          If Nvl(n_����, 0) = 1 Or Nvl(n_����ʾ, 0) = 1 Then
            Select Max(����), Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ), Count(1), Max(�ϼ�����)
            Into v_����, v_Code����, n_����, n_����ʾ, n_Count, v_Code����
            From ����
            Where ���� Like v_Tmp || '%' And Nvl(����, 0) = 3 And �ϼ����� In (Select ���� From ���� Where �ϼ����� = v_Code��);
            --���ڶ���ƥ�䣬��������ӳ��ȣ���ʱ��2�������ӵ�3����ƥ��
            If n_Count > 1 Then
              v_Tmp := Substr(v_Adrstmp, 1, 3);
              Select Max(����), Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ), Max(�ϼ�����)
              Into v_����, v_Code����, n_����, n_����ʾ, v_Code����
              From ����
              Where ���� Like v_Tmp || '%' And Nvl(����, 0) = 3 And �ϼ����� In (Select ���� From ���� Where �ϼ����� = v_Code��);
            End If;

            If v_Code���� Is Not Null Then
              v_Info���� := v_���� || ',' || v_Code���� || ',' || n_���� || ',' || n_����ʾ;
              v_Adrstmp  := Substr(v_Adrstmp, Length(v_����) + 1);
              Select Max(����), Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ)
              Into v_����, v_Code����, n_����, n_����ʾ
              From ����
              Where ���� = v_Code����;
              v_Info���� := v_���� || ',' || v_Code���� || ',' || n_���� || ',' || n_����ʾ;
            End If;
          End If;
        Else
          v_Info���� := v_���� || ',' || v_Code���� || ',' || n_���� || ',' || n_����ʾ;
          v_Adrstmp  := Substr(v_Adrstmp, Length(v_����) + 1);
        End If;
      End If;
      If v_Code���� Is Not Null And v_Code���� Is Null Then
        --�Ƚ�ȡ���򼶵����������ؼ��֣���ƥ��
        v_Tmp := Substr(v_Adrstmp, 1, 2);
        Select Max(����), Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ), Count(1)
        Into v_����, v_Code����, n_����, n_����ʾ, n_Count
        From ����
        Where ���� Like v_Tmp || '%' And Nvl(����, 0) = 3 And �ϼ����� = v_Code����;
        --���ڶ���ƥ�䣬��������ӳ��ȣ���ʱ��2�������ӵ�3����ƥ��
        If n_Count > 1 Then
          v_Tmp := Substr(v_Adrstmp, 1, 3);
          Select Max(����), Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ), Count(1)
          Into v_����, v_Code����, n_����, n_����ʾ, n_Count
          From ����
          Where ���� Like v_Tmp || '%' And Nvl(����, 0) = 3 And �ϼ����� = v_Code����;
        End If;
        If v_Code���� Is Null Then
          Select Max(�Ƿ�����), Max(�Ƿ���ʾ) Into n_����, n_����ʾ From ���� Where �ϼ����� = v_Code����;
          If Nvl(n_����, 0) = 1 Or Nvl(n_����ʾ, 0) = 1 Then
            Select Max(����), Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ), Count(1), Max(�ϼ�����)
            Into v_�ֵ�, v_Code�ֵ�, n_����, n_����ʾ, n_Count, v_Code����
            From ����
            Where ���� = v_Adrstmp And �ϼ����� In (Select ���� From ���� Where �ϼ����� = v_Code����);
          End If;
          If v_Code���� Is Not Null Then
            v_Info�ֵ� := v_�ֵ� || ',' || v_Code�ֵ� || ',' || n_���� || ',' || n_����ʾ;
            Select Max(����), Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ)
            Into v_����, v_Code����, n_����, n_����ʾ
            From ����
            Where ���� = v_Code����;
            v_Info���� := v_���� || ',' || v_Code���� || ',' || n_���� || ',' || n_����ʾ;
          End If;
        Else
          v_Info���� := v_���� || ',' || v_Code���� || ',' || n_���� || ',' || n_����ʾ;
          v_Adrstmp  := Substr(v_Adrstmp, Length(v_����) + 1);
        End If;
        If v_Code���� Is Not Null And v_Code�ֵ� Is Null Then
          Select Max(����), Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ)
          Into v_�ֵ�, v_Code�ֵ�, n_����, n_����ʾ
          From ����
          Where ���� = v_Adrstmp And Nvl(����, 0) = 4 And �ϼ����� = v_Code����;
          If v_Code�ֵ� Is Not Null Then
            v_Info�ֵ� := v_�ֵ� || ',' || v_Code�ֵ� || ',' || n_���� || ',' || n_����ʾ;
          End If;
        End If;
      End If;
    End If;
    If v_�ֵ� Is Null Then
      v_�ֵ� := v_Adrstmp;
    End If;
  End If;
  v_Infoʡ := v_ʡ || ',' || v_Codeʡ || ',,,';
  If v_Info�� Is Null Then
    v_Info�� := v_�� || ',,,';
  End If;
  --ֻ��ʡû���У��ж����Ƿ�ֻ�����⼶
  If Not v_Codeʡ Is Null And v_�� Is Null Then
    Select Count(1)
    Into n_Count
    From ����
    Where �ϼ����� = v_Codeʡ And Nvl(�Ƿ�����, 0) = 0 And Nvl(�Ƿ���ʾ, 0) = 0 And Rownum < 2;
    If n_Count = 0 Then
      Select Count(1) Into n_Count From ���� Where �ϼ����� = v_Codeʡ And Rownum < 2;
      If n_Count = 0 Then
        v_Info�� := v_Info�� || ',';
      Else
        v_Info�� := v_Info�� || ',1';
      End If;
    Else
      v_Info�� := v_Info�� || ',';
    End If;
  Else
    v_Info�� := v_Info�� || ',';
  End If;
  If v_Info���� Is Null Then
    v_Info���� := v_���� || ',,,';
  End If;
  --ֻ����û�����أ��ж�����ֻ�����⼶
  If Not v_Code�� Is Null And v_���� Is Null Then
    Select Count(1)
    Into n_Count
    From ����
    Where �ϼ����� = v_Code�� And Nvl(�Ƿ�����, 0) = 0 And Nvl(�Ƿ���ʾ, 0) = 0 And Rownum < 2;
    If n_Count = 0 Then
      Select Count(1) Into n_Count From ���� Where �ϼ����� = v_Code�� And Rownum < 2;
      If n_Count = 0 Then
        v_Info���� := v_Info���� || ',';
      Else
        v_Info���� := v_Info���� || ',1';
      End If;
    Else
      v_Info���� := v_Info���� || ',';
    End If;
  Else
    v_Info���� := v_Info���� || ',';
  End If;
  If v_Info���� Is Null Then
    v_Info���� := v_���� || ',,,';
  End If;
  --ֻ������û�������ж������Ƿ�ֻ��������¼�
  If Not v_Code���� Is Null And v_���� Is Null Then
    Select Count(1)
    Into n_Count
    From ����
    Where �ϼ����� = v_Code���� And Nvl(�Ƿ�����, 0) = 0 And Nvl(�Ƿ���ʾ, 0) = 0 And Rownum < 2;
    If n_Count = 0 Then
      Select Count(1) Into n_Count From ���� Where �ϼ����� = v_Code���� And Rownum < 2;
      If n_Count = 0 Then
        v_Info���� := v_Info���� || ',';
      Else
        v_Info���� := v_Info���� || ',1';
      End If;
    Else
      v_Info���� := v_Info���� || ',';
    End If;
  Else
    v_Info���� := v_Info���� || ',';
  End If;
  If v_Info�ֵ� Is Null Then
    v_Info�ֵ� := v_�ֵ� || ',,,,';
  Else
    v_Info�ֵ� := v_Info�ֵ� || ',';
  End If;
  v_Return := v_Infoʡ || '|' || v_Info�� || '|' || v_Info���� || '|' || v_Info���� || '|' || v_Info�ֵ�;
  Return(v_Return);
End;
/

--125588:����,2018-07-02,����ʵ������Ϊ0�����ĵ���
Create Or Replace Procedure Zl_�����շ���¼_Adjust
(
  ����id_In   In Number, --���ۼ�¼��ID
  ����_In     In Number := 0, --�Ƿ�תΪ�������ۣ����²������ԡ��շ�ϸĿ�еı�ۣ�
  ����id_In   In Number := 0, --����Ϊ0ʱ��ʾ�ǳɱ��۵��ۣ��������ۼ��������
  Billinfo_In In Varchar2 := Null --����ʱ�����İ����ε��ۡ���ʽ:"����1,�ּ�1|����2,�ּ�2|....."
) As
  n_������id ҩƷ�շ���¼.������id%Type; --������
  v_���۵��ݺ� ҩƷ�շ���¼.No%Type; --���۵���
  d_��Ч����   Date; --������Чʱ��
  n_ִ�е���   Number(1); --����ʱ�̵���
  n_ʵ�۲���   Number(1); --ʱ��ҩƷ
  n_�շ�ϸĿid Number(18); --�շ�ϸĿID
  d_�������   ҩƷ�շ���¼.�������%Type;
  n_���۽��   ҩƷ���.ʵ�ʽ��%Type;
  n_���ۼ�     ҩƷ���.���ۼ�%Type;
  n_���       Integer(8);
  v_Infotmp    Varchar2(4000);
  v_Fields     Varchar2(4000);
  n_����       Number(18);
  n_�ּ�       �շѼ�Ŀ.�ּ�%Type;
  n_ԭ��       �շѼ�Ŀ.ԭ��%Type;
  n_�շ�id     ҩƷ�շ���¼.Id%Type;
  n_ʱ�۷���   Number(1);

  Cursor c_Price --��ͨ����
  Is
    Select 1 ��¼״̬, 13 ����, v_���۵��ݺ� NO, Rownum ���, n_������id ������id, m.����id ҩƷid, s.���� ����, Null ����, s.Ч��,
           Decode(s.�ϴβ���, Null, q.����, s.�ϴβ���) ����, 1 ����, s.ʵ������ ��д����, 0 ʵ������, a.ԭ�� �ɱ���, 0 �ɱ����, a.�ּ� ���ۼ�, 0 ����,
           Nvl(s.���ۼ�, 0) As ������ۼ�, s.ʵ�ʽ�� As �����, s.ʵ�ʲ�� As �����, '���ĵ���' ժҪ, User ������, Sysdate ��������, s.�ⷿid �ⷿid,
           1 ���ϵ��, a.Id �۸�id, s.�ϴ���������, s.���Ч��, s.��׼�ĺ�, s.�ϴι�Ӧ��id,
           Nvl(s.���ۼ�, Decode(Nvl(s.ʵ������, 0), 0, a.ԭ��, Nvl(s.ʵ�ʽ��, 0) / s.ʵ������)) As ԭ�ۼ�
    From ҩƷ��� S, �������� M, �շѼ�Ŀ A, �շ���ĿĿ¼ Q
    Where s.ҩƷid = m.����id And m.����id = q.Id And m.����id = a.�շ�ϸĿid And s.���� = 1 And a.�䶯ԭ�� = 0 And a.Id = ����id_In And
          a.ִ������ <= Sysdate;

  Cursor c_ʱ�۰����ε��� --ʱ�����İ����ε���
  Is
    Select 1 ��¼״̬, 13 ����, v_���۵��ݺ� NO, n_��� + Rownum ���, n_������id ������id, s.ҩƷid ҩƷid, s.���� ����, Null ����, s.Ч��,
           Decode(s.�ϴβ���, Null, b.����, s.�ϴβ���) ����, 1 ����, Nvl(s.ʵ������, 0) ��д����, 0 ʵ������, a.ԭ�� �ɱ���, 0 �ɱ����, n_�ּ� ���ۼ�, 0 ����,
           '���ĵ���' ժҪ, User ������, Sysdate ��������, s.�ⷿid �ⷿid, 1 ���ϵ��, a.Id �۸�id, Nvl(b.�Ƿ���, 0) As ʱ��, s.ʵ�ʽ�� As �����,
           s.ʵ�ʲ�� As �����, Nvl(s.���ۼ�, Decode(Nvl(s.ʵ������, 0), 0, a.ԭ��, Nvl(s.ʵ�ʽ��, 0) / s.ʵ������)) As ԭ�ۼ�
    From ҩƷ��� S, �������� M, �շѼ�Ŀ A, �շ���ĿĿ¼ B
    Where s.ҩƷid = m.����id And m.����id = a.�շ�ϸĿid And a.�շ�ϸĿid = b.Id And s.���� = 1 And a.�䶯ԭ�� = 0 And a.Id = ����id_In And
          a.ִ������ <= Sysdate And Nvl(s.����, 0) = n_����;
Begin

  If ����id_In <> 0 Then
    --�ɱ��۵���
    Zl_�����շ���¼_�ɱ��۵���(����id_In);
    Return;
  End If;

  --ȡ������ID
  Select ���id Into n_������id From ҩƷ�������� Where ���� = 13;

  --ȡ����
  Select Nextno(147) Into v_���۵��ݺ� From Dual;
  --ȡ���ۼ�¼��Ч����
  Select �շ�ϸĿid, ִ������ Into n_�շ�ϸĿid, d_��Ч���� From �շѼ�Ŀ Where ID = ����id_In;
  --ȡ�ò����Ƿ���ʱ��ҩƷ
  Select Nvl(�Ƿ���, 0) Into n_ʵ�۲��� From �շ���ĿĿ¼ Where ID = n_�շ�ϸĿid;

  If Sysdate >= d_��Ч���� Then
    n_ִ�е��� := 1;
  Else
    n_ִ�е��� := 0;
  End If;

  If n_ִ�е��� = 1 Then
    d_������� := Sysdate;
    --��ͨ���۴���
    If Billinfo_In = '' Or Billinfo_In Is Null Then
      --��ʱ��ҩƷ����
      For c_���� In c_Price Loop
        /*If Nvl(c_����.��д����, 0) = 0 And Nvl(c_����.�����, 0) = 0 And Nvl(c_����.�����, 0) = 0 Then
        Null;*/
        If Nvl(c_����.��д����, 0) = 0 And (Nvl(c_����.�����, 0) <> 0 Or Nvl(c_����.�����, 0) <> 0) Then
          --����=0 ������<>0ʱֻ���¿����ж�Ӧ�����ۼ�,�������ۼ��������ݵ��ǽ���=0��ֻ��¼�����ۼۣ�����Ͳ�۲������

        
          --��������Ӱ���¼
          Insert Into ҩƷ�շ���¼
            (ID, ��¼״̬, ����, NO, ���, ������id, ҩƷid, ����, ����, Ч��, ����, ����, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ����, ժҪ, ������, ��������,
             �ⷿid, ���ϵ��, �۸�id, �����, �������, ��������, ���Ч��, ��׼�ĺ�, ��ҩ��λid, ����, Ƶ��)
          Values
            (ҩƷ�շ���¼_Id.Nextval, c_����.��¼״̬, c_����.����, c_����.No, c_����.���, c_����.������id, c_����.ҩƷid, c_����.����, c_����.����, c_����.Ч��,
             c_����.����, c_����.����, c_����.��д����, c_����.ʵ������, Decode(n_ʵ�۲���, 1, c_����.ԭ�ۼ�, c_����.�ɱ���), c_����.�ɱ����, c_����.���ۼ�, c_����.����,
             c_����.ժҪ, c_����.������, c_����.��������, c_����.�ⷿid, c_����.���ϵ��, c_����.�۸�id, User, d_�������, c_����.�ϴ���������, c_����.���Ч��,
             c_����.��׼�ĺ�, c_����.�ϴι�Ӧ��id, c_����.�����, c_����.�����);
        
          --���²��Ͽ�� ��ֻ��ʱ�����ĲŸ������ۼ�
          Update ҩƷ���
          Set ���ۼ� = Decode(n_ʵ�۲���, 1, Decode(Nvl(c_����.����, 0), 0, Null, c_����.���ۼ�), Null)
          Where �ⷿid = c_����.�ⷿid And ҩƷid = c_����.ҩƷid And ���� = 1 And Nvl(����, 0) = Nvl(c_����.����, 0);
        Else
          If n_ʵ�۲��� = 1 Then
            If c_����.������ۼ� = 0 Then
              n_���ۼ� := c_����.����� / c_����.��д����;
            Else
              n_���ۼ� := c_����.������ۼ�;
            End If;
          Else
            n_���ۼ� := c_����.�ɱ���;
          End If;
          n_���۽�� := Round((c_����.���ۼ� - n_���ۼ�) * c_����.��д����, 2);
        
          --��������Ӱ���¼
          Insert Into ҩƷ�շ���¼
            (ID, ��¼״̬, ����, NO, ���, ������id, ҩƷid, ����, ����, Ч��, ����, ����, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ����, ���۽��, ���, ժҪ, ������,
             ��������, �ⷿid, ���ϵ��, �۸�id, �����, �������, ��������, ���Ч��, ��׼�ĺ�, ��ҩ��λid, ����, Ƶ��)
          Values
            (ҩƷ�շ���¼_Id.Nextval, c_����.��¼״̬, c_����.����, c_����.No, c_����.���, c_����.������id, c_����.ҩƷid, c_����.����, c_����.����, c_����.Ч��,
             c_����.����, c_����.����, c_����.��д����, c_����.ʵ������, Decode(n_ʵ�۲���, 1, c_����.ԭ�ۼ�, c_����.�ɱ���), c_����.�ɱ����, c_����.���ۼ�, c_����.����,
             n_���۽��, n_���۽��, c_����.ժҪ, c_����.������, c_����.��������, c_����.�ⷿid, c_����.���ϵ��, c_����.�۸�id, User, d_�������, c_����.�ϴ���������,
             c_����.���Ч��, c_����.��׼�ĺ�, c_����.�ϴι�Ӧ��id, c_����.�����, c_����.�����);
        
          --���²��Ͽ��
          Update ҩƷ���
          Set ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) + n_���۽��, ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + n_���۽��,
              ���ۼ� = Decode(n_ʵ�۲���, 1, Decode(Nvl(c_����.����, 0), 0, Null, c_����.���ۼ�), Null)
          Where �ⷿid = c_����.�ⷿid And ҩƷid = c_����.ҩƷid And ���� = 1 And Nvl(����, 0) = Nvl(c_����.����, 0);
        
          If Sql%RowCount = 0 Then
            Insert Into ҩƷ���
              (�ⷿid, ҩƷid, ����, ����, ��������, ʵ������, ʵ�ʽ��, ʵ�ʲ��, Ч��, ���Ч��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴ���������, �ϴβ���, ��׼�ĺ�, ���ۼ�)
            Values
              (c_����.�ⷿid, c_����.ҩƷid, c_����.����, 1, 0, 0, n_���۽��, n_���۽��, c_����.Ч��, c_����. ���Ч��, c_����.�ϴι�Ӧ��id, c_����.�ɱ���,
               c_����.����, c_����.�ϴ���������, c_����.����, c_����.��׼�ĺ�,
               Decode(n_ʵ�۲���, 1, Decode(Nvl(c_����.����, 0), 0, Null, c_����.���ۼ�), Null));
          End If;
        End If;
      End Loop;
    Else
      --ʱ�۷������۴���
      n_��� := 0;
      --ʱ��ҩƷ�����ε���
      v_Infotmp := Billinfo_In || '|';
      While v_Infotmp Is Not Null Loop
        --�ֽⵥ��ID��
        v_Fields  := Substr(v_Infotmp, 1, Instr(v_Infotmp, '|') - 1);
        n_����    := Substr(v_Fields, 1, Instr(v_Fields, ',') - 1);
        n_�ּ�    := Substr(v_Fields, Instr(v_Fields, ',') + 1);
        v_Infotmp := Replace('|' || v_Infotmp, '|' || v_Fields || '|');
      
        For v_ʱ�۰����ε��� In c_ʱ�۰����ε��� Loop
          If v_ʱ�۰����ε���.��д���� <> 0 Then
            n_ԭ�� := Nvl(v_ʱ�۰����ε���.�����, 0) / v_ʱ�۰����ε���.��д����;
          Else
            n_ԭ�� := v_ʱ�۰����ε���.�ɱ���;
          End If;
        
          Select ҩƷ�շ���¼_Id.Nextval Into n_�շ�id From Dual;
          If Nvl(v_ʱ�۰����ε���.��д����, 0) = 0 And Nvl(v_ʱ�۰����ε���.�����, 0) = 0 And Nvl(v_ʱ�۰����ε���.�����, 0) = 0 Then
            Null;
          Elsif Nvl(v_ʱ�۰����ε���.��д����, 0) = 0 And (Nvl(v_ʱ�۰����ε���.�����, 0) <> 0 Or Nvl(v_ʱ�۰����ε���.�����, 0) <> 0) Then
            --����=0 ������<>0ʱֻ���¿����ж�Ӧ�����ۼ�,�������ۼ��������ݵ��ǽ���=0��ֻ��¼�����ۼۣ�����Ͳ�۲������

          
            --��������Ӱ���¼
            Insert Into ҩƷ�շ���¼
              (ID, ��¼״̬, ����, NO, ���, ������id, ҩƷid, ����, ����, Ч��, ����, ����, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ����, ժҪ, ������, ��������,
               �ⷿid, ���ϵ��, �۸�id, �����, �������, ����, Ƶ��)
            Values
              (n_�շ�id, v_ʱ�۰����ε���.��¼״̬, v_ʱ�۰����ε���.����, v_ʱ�۰����ε���.No, v_ʱ�۰����ε���.���, v_ʱ�۰����ε���.������id, v_ʱ�۰����ε���.ҩƷid,
               v_ʱ�۰����ε���.����, v_ʱ�۰����ε���.����, v_ʱ�۰����ε���.Ч��, v_ʱ�۰����ε���.����, v_ʱ�۰����ε���.����, v_ʱ�۰����ε���.��д����, v_ʱ�۰����ε���.ʵ������,
               Decode(n_ʵ�۲���, 1, v_ʱ�۰����ε���.ԭ�ۼ�, v_ʱ�۰����ε���.�ɱ���), v_ʱ�۰����ε���.�ɱ����, v_ʱ�۰����ε���.���ۼ�, v_ʱ�۰����ε���.����,
               v_ʱ�۰����ε���.ժҪ, v_ʱ�۰����ε���.������, v_ʱ�۰����ε���.��������, v_ʱ�۰����ε���.�ⷿid, v_ʱ�۰����ε���.���ϵ��, v_ʱ�۰����ε���.�۸�id, User, d_�������,
               v_ʱ�۰����ε���.�����, v_ʱ�۰����ε���.�����);
            n_��� := n_��� + 1;
            --������
            --���¿�����ۼ�,ֻ��ʱ�۷���ҩƷ���ܸ������ۼ��ֶ�
            Update ҩƷ���
            Set ���ۼ� = Decode(v_ʱ�۰����ε���.ʱ��, 1, Decode(Nvl(v_ʱ�۰����ε���.����, 0), 0, Null, v_ʱ�۰����ε���.���ۼ�), Null)
            Where �ⷿid = v_ʱ�۰����ε���.�ⷿid And ҩƷid = v_ʱ�۰����ε���.ҩƷid And ���� = 1 And Nvl(����, 0) = Nvl(v_ʱ�۰����ε���.����, 0);
          Else
            n_���ۼ�   := v_ʱ�۰����ε���.����� / v_ʱ�۰����ε���.��д����;
            n_���۽�� := Round((n_�ּ� - n_���ۼ�) * v_ʱ�۰����ε���.��д����, 2);
            --��������Ӱ���¼
            Insert Into ҩƷ�շ���¼
              (ID, ��¼״̬, ����, NO, ���, ������id, ҩƷid, ����, ����, Ч��, ����, ����, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ����, ���۽��, ���, ժҪ,
               ������, ��������, �ⷿid, ���ϵ��, �۸�id, �����, �������, ����, Ƶ��)
            Values
              (n_�շ�id, v_ʱ�۰����ε���.��¼״̬, v_ʱ�۰����ε���.����, v_ʱ�۰����ε���.No, v_ʱ�۰����ε���.���, v_ʱ�۰����ε���.������id, v_ʱ�۰����ε���.ҩƷid,
               v_ʱ�۰����ε���.����, v_ʱ�۰����ε���.����, v_ʱ�۰����ε���.Ч��, v_ʱ�۰����ε���.����, v_ʱ�۰����ε���.����, v_ʱ�۰����ε���.��д����, v_ʱ�۰����ε���.ʵ������,
               Decode(n_ʵ�۲���, 1, v_ʱ�۰����ε���.ԭ�ۼ�, v_ʱ�۰����ε���.�ɱ���), v_ʱ�۰����ε���.�ɱ����, v_ʱ�۰����ε���.���ۼ�, v_ʱ�۰����ε���.����, n_���۽��,
               n_���۽��, v_ʱ�۰����ε���.ժҪ, v_ʱ�۰����ε���.������, v_ʱ�۰����ε���.��������, v_ʱ�۰����ε���.�ⷿid, v_ʱ�۰����ε���.���ϵ��, v_ʱ�۰����ε���.�۸�id, User,
               d_�������, v_ʱ�۰����ε���.�����, v_ʱ�۰����ε���.�����);
            n_��� := n_��� + 1;
            --������
            If v_ʱ�۰����ε���.ʱ�� = 1 And Nvl(v_ʱ�۰����ε���.����, 0) > 0 Then
              n_ʱ�۷��� := 1;
            Else
              n_ʱ�۷��� := 0;
            End If;
          
            If Nvl(v_ʱ�۰����ε���.����, 0) = 0 Then
              Update ҩƷ���
              Set ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) + n_���۽��, ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + n_���۽��
              Where �ⷿid = v_ʱ�۰����ε���.�ⷿid And ҩƷid = v_ʱ�۰����ε���.ҩƷid And ���� = 1 And (���� Is Null Or ���� = 0);
            Else
              Update ҩƷ���
              Set ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) + n_���۽��, ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + n_���۽��,
                  ���ۼ� = Decode(n_ʱ�۷���, 1, v_ʱ�۰����ε���.���ۼ�, ���ۼ�)
              Where �ⷿid = v_ʱ�۰����ε���.�ⷿid And ҩƷid = v_ʱ�۰����ε���.ҩƷid And ���� = 1 And ���� = v_ʱ�۰����ε���.����;
            End If;
          
            If Sql%RowCount = 0 Then
              Insert Into ҩƷ���
                (�ⷿid, ҩƷid, ����, ����, ��������, ʵ������, ʵ�ʽ��, ʵ�ʲ��, ���ۼ�)
              Values
                (v_ʱ�۰����ε���.�ⷿid, v_ʱ�۰����ε���.ҩƷid, v_ʱ�۰����ε���.����, 1, 0, 0, n_���۽��, n_���۽��,
                 Decode(n_ʱ�۷���, 1, v_ʱ�۰����ε���.���ۼ�, Null));
            End If;
          End If;
        End Loop;
      End Loop;
    End If;
  
    Update ҩƷ�շ���¼ Set ����� = User, ������� = Sysdate Where �۸�id = ����id_In;
    Update �շѼ�Ŀ Set �䶯ԭ�� = 1 Where ID = ����id_In;
  
    --����ҩƷĿ¼���շ�ϸĿ�еı��
    If ����_In = 1 Then
      Update �շ���ĿĿ¼ Set �Ƿ��� = 0 Where ID = n_�շ�ϸĿid;
    End If;
    --�ɱ��۵���
    Zl_�����շ���¼_�ɱ��۵���(n_�շ�ϸĿid);
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����շ���¼_Adjust;
/

--125588:����,2018-07-02,��������ʵ������Ϊ0�ĵ���
Create Or Replace Procedure Zl_�����շ���¼_�ɱ��۵���(����id_In In ҩƷ�շ���¼.ҩƷid%Type) As
  v_No         ҩƷ�շ���¼.No%Type;
  v_Ӧ��id     Ӧ����¼.Id%Type; --Ӧ����¼��ID 
  v_Ӧ�����ݺ� Ӧ����¼.No%Type;
  d_����ʱ��   Date;
  n_���       Number(8);
  n_�ⷿid     ҩƷ�շ���¼.�ⷿid%Type;
  n_������id ҩƷ�շ���¼.������id%Type;
  n_���ϵ��   ҩƷ�շ���¼.���ϵ��%Type;
  n_�շ�id     ҩƷ�շ���¼.Id%Type;
  n_������     ҩƷ�շ���¼.���۽��%Type;
  n_ԭ�ɱ���   ҩƷ�շ���¼.�ɱ���%Type;
  n_�³ɱ���   ҩƷ�շ���¼.�ɱ���%Type;
  n_ƽ���ɱ��� ҩƷ���.ƽ���ɱ���%Type;
  v_����id     �ɱ��۵�����Ϣ.Id%Type;
  v_���ۻ��ܺ� �ɱ��۵�����Ϣ.���ۻ��ܺ�%Type;
  n_Count      Number(1) := 0;

  Cursor c_Stock Is --��ǰ��� 
    Select �ϴι�Ӧ��id, a.�ⷿid, a.ҩƷid As ����id, Nvl(a.����, 0) As ����, a.�ϴ�����, a.Ч��, a.�ϴβ���, a.���Ч��,
           Decode(Sign(Nvl(a.����, 0)), 1, a.�ϴβɹ���, a.ƽ���ɱ���) As ԭ�ɱ���
    From ҩƷ��� A
    Where a.���� = 1 And Nvl(a.ʵ������, 0) <> 0 And a.ҩƷid = ����id_In
    Order By a.�ⷿid;

  v_Stock c_Stock%RowType;
Begin
  d_����ʱ�� := Sysdate;
  n_�ⷿid   := 0;

  --�ж��Ƿ�����޿����� 
  Begin
    Select ID, �³ɱ���, ���ۻ��ܺ�
    Into v_����id, n_�³ɱ���, v_���ۻ��ܺ�
    From �ɱ��۵�����Ϣ
    Where ִ������ Is Null And Nvl(�ⷿid, 0) = 0 And ҩƷid = ����id_In;
  Exception
    When Others Then
      v_����id   := 0;
      n_�³ɱ��� := Null;
  End;

  --�޿����� 
  If v_����id > 0 Then
    --���ݵ�ǰ������²���������Ϣ 
    For v_Stock In c_Stock Loop
      Zl_���ϳɱ�����_Insert(v_Stock.�ϴι�Ӧ��id, v_Stock.�ⷿid, v_Stock.����id, v_Stock.����, v_Stock.�ϴ�����, v_Stock.ԭ�ɱ���, n_�³ɱ���,
                       Null, Null, 0, 0, v_���ۻ��ܺ�);
      n_Count := n_Count + 1;
    End Loop;
  
    If n_Count > 0 Then
      --�����ǰ�п���¼����ɾ���޿����ۼ�¼ 
      Delete �ɱ��۵�����Ϣ Where ID = v_����id;
    Else
      Update �ɱ��۵�����Ϣ Set ִ������ = d_����ʱ�� Where ID = v_����id;
    
      Update �������� Set �ɱ��� = n_�³ɱ��� Where ����id = ����id_In And �ɱ��� <> n_�³ɱ���;
    End If;
  End If;

  --ȡ����۵�����������ID 
  Select b.Id, b.ϵ��
  Into n_������id, n_���ϵ��
  From ҩƷ�������� A, ҩƷ������ B
  Where a.���id = b.Id And a.���� = 33 And Rownum < 2;

  For c_�ɱ����� In (Select a.�ⷿid, a.ҩƷid As ����id, Nvl(a.����, 0) ����, a.�ϴι�Ӧ��id, a.ʵ������, a.ʵ�ʽ��, a.ʵ�ʲ��, a.�ϴβ��� As ����,
                        a.�ϴ����� As ����, a.���Ч��, a.Ч��, a.�ϴ��������� As ��������, a.��׼�ĺ�, Nvl(a.ƽ���ɱ���, 0) As ԭ�ɱ���, b.�³ɱ���, b.��Ʊ��,
                        b.��Ʊ����, b.��Ʊ���, Nvl(a.�ϴβɹ���, 0) As �ϴβɹ���, b.Id As ����id
                 From ҩƷ��� A, �ɱ��۵�����Ϣ B
                 Where a.ҩƷid = b.ҩƷid And Nvl(a.�ϴι�Ӧ��id, 0) = Nvl(b.��ҩ��λid, 0) And a.�ⷿid = b.�ⷿid And
                       Nvl(a.����, 0) = Nvl(b.����, 0) And a.���� = 1 And b.ִ������ Is Null And a.ҩƷid = ����id_In
                 Order By a.�ⷿid) Loop
    If n_�ⷿid <> c_�ɱ�����.�ⷿid Then
      n_���   := 1;
      n_�ⷿid := c_�ɱ�����.�ⷿid;
      v_No     := Nextno(71, n_�ⷿid);
    Else
      n_��� := n_��� + 1;
    End If;
  
    Select ҩƷ�շ���¼_Id.Nextval Into n_�շ�id From Dual;
  
    /*If Nvl(c_�ɱ�����.ʵ������, 0) = 0 And Nvl(c_�ɱ�����.ʵ�ʽ��, 0) = 0 And Nvl(c_�ɱ�����.ʵ�ʲ��, 0) = 0 Then
    --����,����۶�Ϊ0�����ʾ��������¿���������������ĵ��ݣ��˵��ݻ�û����ˣ����ֻ��Ҫ���µ�����Ϣ������������
    Update �������� Set �ɱ��� = c_�ɱ�����.�³ɱ��� Where ����id = c_�ɱ�����.����id;
    
    Update �ɱ��۵�����Ϣ
    Set �շ�id = n_�շ�id, ִ������ = d_����ʱ��, Ч�� = c_�ɱ�����.Ч��, ���Ч�� = c_�ɱ�����.���Ч��, ���� = c_�ɱ�����.����, ���� = c_�ɱ�����.����
    Where ID = c_�ɱ�����.����id;*/
    If Nvl(c_�ɱ�����.ʵ������, 0) = 0 And (Nvl(c_�ɱ�����.ʵ�ʽ��, 0) <> 0 Or Nvl(c_�ɱ�����.ʵ�ʲ��, 0) <> 0) Then
      --����=0 ������<>0ʱֻ���¿����ж�Ӧ��ƽ���ɱ��ۺ����Ա��гɱ��ۣ��������ɱ����������ݵ��ǲ�۲�=0��ֻ��¼���³ɱ���
      --�������ۼ�¼��ֻ��¼���³ɱ���
      Insert Into ҩƷ�շ���¼
        (ID, ��¼״̬, ����, NO, ���, �ⷿid, ������id, ��ҩ��λid, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ��д����, ���ۼ�, �ɱ���, ���, ժҪ, ������, ��������, �����,
         �������, ��������, ��׼�ĺ�, ����, ��ҩ��ʽ, ����)
      Values
        (n_�շ�id, 1, 18, v_No, n_���, c_�ɱ�����.�ⷿid, n_������id, c_�ɱ�����.�ϴι�Ӧ��id, n_���ϵ��, c_�ɱ�����.����id, c_�ɱ�����.����, c_�ɱ�����.����,
         c_�ɱ�����.����, c_�ɱ�����.Ч��, 0, c_�ɱ�����.ʵ�ʽ��, c_�ɱ�����.ʵ�ʲ��, 0, '�������ϳɱ��۵���', Zl_Username, d_����ʱ��, Zl_Username, d_����ʱ��,
         c_�ɱ�����.��������, c_�ɱ�����.��׼�ĺ�, c_�ɱ�����.�³ɱ���, 1, c_�ɱ�����.ԭ�ɱ���);
      --���¿��      
      Update ҩƷ���
      Set ƽ���ɱ��� = c_�ɱ�����.�³ɱ���, �ϴβɹ��� = c_�ɱ�����.�³ɱ���
      Where �ⷿid = c_�ɱ�����.�ⷿid And ҩƷid = c_�ɱ�����.����id And Nvl(����, 0) = c_�ɱ�����.���� And ���� = 1;
      Update �������� Set �ɱ��� = c_�ɱ�����.�³ɱ��� Where ����id = c_�ɱ�����.����id;
    
      Update �ɱ��۵�����Ϣ
      Set �շ�id = n_�շ�id, ִ������ = d_����ʱ��, Ч�� = c_�ɱ�����.Ч��, ���Ч�� = c_�ɱ�����.���Ч��, ���� = c_�ɱ�����.����, ���� = c_�ɱ�����.����
      Where ID = c_�ɱ�����.����id;
    Else
      --������Ӧ�Ŀ��:ԭ�ɱ����-ʵ�³ɱ���� 
      n_������   := (c_�ɱ�����.ʵ�ʽ�� - c_�ɱ�����.ʵ�ʲ��) - Round(c_�ɱ�����.�³ɱ��� * c_�ɱ�����.ʵ������, 2);
      n_ԭ�ɱ��� := c_�ɱ�����.ԭ�ɱ���;
    
      If n_ԭ�ɱ��� <= 0 Then
        n_ԭ�ɱ��� := c_�ɱ�����.�ϴβɹ���;
      End If;
    
      --Ŀǰ���շ���¼��Ӧ: 
      -- ����--> ԭ�ɱ��� 
      -- ����-->�³ɱ��� 
      -- ��д����-->���ʵ������ 
      -- ���ۼ�-->���ʵ�ʽ�� 
      -- �ɱ���-->���ʵ�ʲ�� 
      -- ���-->���ε����� 
    
      Insert Into ҩƷ�շ���¼
        (ID, ��¼״̬, ����, NO, ���, �ⷿid, ������id, ��ҩ��λid, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ��д����, ���ۼ�, �ɱ���, ���, ժҪ, ������, ��������, �����,
         �������, ��������, ��׼�ĺ�, ����, ��ҩ��ʽ, ����)
      Values
        (n_�շ�id, 1, 18, v_No, n_���, c_�ɱ�����.�ⷿid, n_������id, c_�ɱ�����.�ϴι�Ӧ��id, n_���ϵ��, c_�ɱ�����.����id, c_�ɱ�����.����, c_�ɱ�����.����,
         c_�ɱ�����.����, c_�ɱ�����.Ч��, c_�ɱ�����.ʵ������, c_�ɱ�����.ʵ�ʽ��, c_�ɱ�����.ʵ�ʲ��, n_������, '�������ϳɱ��۵���', Zl_Username, d_����ʱ��,
         Zl_Username, d_����ʱ��, c_�ɱ�����.��������, c_�ɱ�����.��׼�ĺ�, c_�ɱ�����.�³ɱ���, 1, n_ԭ�ɱ���);
    
      --���¿�� 
      Update ҩƷ���
      Set ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + n_������
      Where �ⷿid = c_�ɱ�����.�ⷿid And ҩƷid = c_�ɱ�����.����id And Nvl(����, 0) = Nvl(c_�ɱ�����.����, 0) And ���� = 1;
    
      If Sql%NotFound Then
        Insert Into ҩƷ���
          (�ⷿid, ҩƷid, ����, ����, ʵ�ʲ��, �ϴ�����, Ч��, �ϴβ���, �ϴι�Ӧ��id, �ϴ���������, ��׼�ĺ�, ���Ч��)
        Values
          (c_�ɱ�����.�ⷿid, c_�ɱ�����.����id, c_�ɱ�����.����, 1, n_������, c_�ɱ�����.����, c_�ɱ�����.Ч��, c_�ɱ�����.����, c_�ɱ�����.�ϴι�Ӧ��id, c_�ɱ�����.��������,
           c_�ɱ�����.��׼�ĺ�, c_�ɱ�����.���Ч��);
      End If;
    
      Update ҩƷ���
      Set �ϴβɹ��� = c_�ɱ�����.�³ɱ���
      Where ҩƷid = c_�ɱ�����.����id And �ϴβɹ��� <> c_�ɱ�����.�³ɱ���;
    
      Update ��������
      Set �ɱ��� = c_�ɱ�����.�³ɱ���
      Where ����id = c_�ɱ�����.����id And �ɱ��� <> c_�ɱ�����.�³ɱ���;
    
      --���¼�������е�ƽ���ɱ��� 
      Update ҩƷ���
      Set ƽ���ɱ��� = Decode(Nvl(����, 0), 0, Decode((ʵ�ʽ�� - ʵ�ʲ��) / ʵ������, 0, �ϴβɹ���, (ʵ�ʽ�� - ʵ�ʲ��) / ʵ������), �ϴβɹ���)
      Where ҩƷid = c_�ɱ�����.����id And Nvl(����, 0) = Nvl(c_�ɱ�����.����, 0) And �ⷿid = c_�ɱ�����.�ⷿid And ���� = 1 And
            Nvl(ʵ������, 0) <> 0;
      If Sql%NotFound Then
        Select �ɱ��� Into n_ƽ���ɱ��� From �������� Where ����id = c_�ɱ�����.����id;
        Update ҩƷ���
        Set ƽ���ɱ��� = n_ƽ���ɱ���
        Where ҩƷid = c_�ɱ�����.����id And �ⷿid = c_�ɱ�����.�ⷿid And Nvl(����, 0) = Nvl(c_�ɱ�����.����, 0) And ���� = 1;
      End If;
    
      --���³ɱ��۵�����Ϣ 
      Update �ɱ��۵�����Ϣ
      Set �շ�id = n_�շ�id, ִ������ = d_����ʱ��, ԭ�ɱ��� = n_ԭ�ɱ���, Ч�� = c_�ɱ�����.Ч��, ���Ч�� = c_�ɱ�����.���Ч��, ���� = c_�ɱ�����.����,
          ���� = c_�ɱ�����.����
      Where ID = c_�ɱ�����.����id;
    End If;
  End Loop;

  --����Ӧ����¼ 
  For c_Ӧ�� In (Select Distinct a.��ҩ��λid, a.ҩƷid, a.��Ʊ��, a.��Ʊ����, a.��Ʊ���, b.����, b.���㵥λ, b.���
               From �ɱ��۵�����Ϣ A, �շ���ĿĿ¼ B
               Where a.ҩƷid = b.Id And Nvl(a.Ӧ����䶯, 0) = 1 And Nvl(a.��ҩ��λid, 0) <> 0 And a.ҩƷid = ����id_In
               Order By a.��ҩ��λid) Loop
  
    v_Ӧ�����ݺ� := Nextno(67);
  
    Select Ӧ����¼_Id.Nextval Into v_Ӧ��id From Dual;
  
    Insert Into Ӧ����¼
      (ID, ��¼����, ��¼״̬, ��λid, NO, ϵͳ��ʶ, ��Ʊ��, ��Ʊ����, ��Ʊ���, Ʒ��, ���, ������, ��������, �����, �������, ժҪ)
    Values
      (v_Ӧ��id, 1, 1, c_Ӧ��.��ҩ��λid, v_Ӧ�����ݺ�, 5, c_Ӧ��.��Ʊ��, c_Ӧ��.��Ʊ����, c_Ӧ��.��Ʊ���, c_Ӧ��.����, c_Ӧ��.���, Zl_Username, d_����ʱ��,
       Zl_Username, d_����ʱ��, '�ɱ��۵����Զ�����Ӧ����䶯��¼');
  
    If Nvl(c_Ӧ��.��ҩ��λid, 0) <> 0 Then
      Update Ӧ����� Set ��� = Nvl(���, 0) + Nvl(c_Ӧ��.��Ʊ���, 0) Where ��λid = c_Ӧ��.��ҩ��λid And ���� = 1;
      If Sql%NotFound Then
        Insert Into Ӧ����� (��λid, ����, ���) Values (c_Ӧ��.��ҩ��λid, 1, Nvl(c_Ӧ��.��Ʊ���, 0));
      End If;
    End If;
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����շ���¼_�ɱ��۵���;
/

--127450:���ϴ�,2018-06-27,�ҺŰ��Ƚ��ȳ�ԭ��ʹ��Ԥ����
Create Or Replace Procedure Zl_���˹Һż�¼_����_Insert
(
  �����¼id_In    �ٴ������¼.Id%Type,
  ����id_In        ������ü�¼.����id%Type,
  �����_In        ������ü�¼.��ʶ��%Type,
  ����_In          ������ü�¼.����%Type,
  �Ա�_In          ������ü�¼.�Ա�%Type,
  ����_In          ������ü�¼.����%Type,
  ���ʽ_In      ������ü�¼.���ʽ%Type, --���ڴ�Ų��˵�ҽ�Ƹ��ʽ���
  �ѱ�_In          ������ü�¼.�ѱ�%Type,
  ���ݺ�_In        ������ü�¼.No%Type,
  Ʊ�ݺ�_In        ������ü�¼.ʵ��Ʊ��%Type,
  ���_In          ������ü�¼.���%Type,
  �۸񸸺�_In      ������ü�¼.�۸񸸺�%Type,
  ��������_In      ������ü�¼.��������%Type,
  �շ����_In      ������ü�¼.�շ����%Type,
  �շ�ϸĿid_In    ������ü�¼.�շ�ϸĿid%Type,
  ����_In          ������ü�¼.����%Type,
  ��׼����_In      ������ü�¼.��׼����%Type,
  ������Ŀid_In    ������ü�¼.������Ŀid%Type,
  �վݷ�Ŀ_In      ������ü�¼.�վݷ�Ŀ%Type,
  ���㷽ʽ_In      Varchar2,
  Ӧ�ս��_In      ������ü�¼.Ӧ�ս��%Type,
  ʵ�ս��_In      ������ü�¼.ʵ�ս��%Type,
  ���˿���id_In    ������ü�¼.���˿���id%Type,
  ��������id_In    ������ü�¼.��������id%Type,
  ִ�в���id_In    ������ü�¼.ִ�в���id%Type,
  ����Ա���_In    ������ü�¼.����Ա���%Type,
  ����Ա����_In    ������ü�¼.����Ա����%Type,
  ����ʱ��_In      ������ü�¼.����ʱ��%Type,
  �Ǽ�ʱ��_In      ������ü�¼.�Ǽ�ʱ��%Type,
  ҽ������_In      �ҺŰ���.ҽ������%Type,
  ҽ��id_In        �ҺŰ���.ҽ��id%Type,
  ������_In        Number, --������¼�Ƿ���������
  ����_In          Number,
  �ű�_In          �ҺŰ���.����%Type,
  ����_In          ������ü�¼.��ҩ����%Type,
  ����id_In        ������ü�¼.����id%Type,
  ����id_In        Ʊ��ʹ����ϸ.����id%Type,
  Ԥ��֧��_In      ����Ԥ����¼.��Ԥ��%Type, --ˢ���Һ�ʱʹ�õ�Ԥ�����,���Ϊ1����.
  �ֽ�֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�ֽ�֧�����ݽ��,���Ϊ1����.
  ����֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�����ʻ�֧�����,,���Ϊ1����.
  ���մ���id_In    ������ü�¼.���մ���id%Type,
  ������Ŀ��_In    ������ü�¼.������Ŀ��%Type,
  ͳ����_In      ������ü�¼.ͳ����%Type,
  ժҪ_In          ������ü�¼.ժҪ%Type, --ԤԼ�Һ�ժҪ��Ϣ
  ԤԼ�Һ�_In      Number := 0, --ԤԼ�Һ�ʱ��(��¼״̬=0,����ʱ��ΪԤԼʱ��),��ʱ����Ҫ���������ز���
  �շ�Ʊ��_In      Number := 0, --�Һ��Ƿ�ʹ���շ�Ʊ��
  ���ձ���_In      ������ü�¼.���ձ���%Type,
  ����_In          ���˹Һż�¼.����%Type := 0,
  ����_In          �Һ����״̬.���%Type := Null, --ԤԼʱ������ü�¼�ķ�ҩ�����ֶ�,�Һ�ʱ����Һż�¼
  ����_In          ���˹Һż�¼.����%Type := Null,
  ԤԼ����_In      Number := 0,
  ԤԼ��ʽ_In      ԤԼ��ʽ.����%Type := Null,
  ���ɶ���_In      Number := 0,
  �����id_In      ����Ԥ����¼.�����id%Type := Null,
  ���㿨���_In    ����Ԥ����¼.���㿨���%Type := Null,
  ����_In          ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In    ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In      ����Ԥ����¼.����˵��%Type := Null,
  ������λ_In      ����Ԥ����¼.������λ%Type := Null,
  ��������_In      Number := 0,
  ����_In          ���˹Һż�¼.����%Type := Null,
  ����ģʽ_In      Number := 0,
  ���ʷ���_In      Number := 0,
  �˺�����_In      Number := 1,
  ��Ԥ������ids_In Varchar2 := Null,
  �������˷ѱ�_In  Number := 0,
  ���½������_In  Number := 0, --�Ƿ������Ա��������Ҫ�Ǵ���ͳһ����Ա��¼��̨�����������
  ԤԼ˳���_In    �ٴ�������ſ���.ԤԼ˳���%Type := Null,
  ������������_In  Number := 0,
  �շѵ�_In        ���˹Һż�¼.�շѵ�%Type := Null
) As
  ---------------------------------------------------------------------------
  --
  --����:
  --     ��������_in:0-�����ҺŻ���ԤԼ 1-����Աӵ�мӺ�Ȩ�޼Ӻ�
  --     �������˷ѱ�_In:0-���޸Ĳ��˷ѱ� 1-�޸Ĳ��˷ѱ�
  ----------------------------------------------------------------------------
  --���α������շѳ�Ԥ���Ŀ���Ԥ���б�
  --��ID�������ȳ��ϴ�δ����ġ�
  Cursor c_Deposit(v_����id ������Ϣ.����id%Type) Is
    Select NO, Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) As ���, Min(��¼״̬) As ��¼״̬, Nvl(Max(����id), 0) As ����id,
           Max(Decode(��¼����, 1, ID, 0) * Decode(��¼״̬, 2, 0, 1)) As ԭԤ��id, Min(Decode(��¼����, 1, �տ�ʱ��, NULL)) as �տ�ʱ��
    From ����Ԥ����¼
    Where ��¼���� In (1, 11) And ����id = v_����id And Nvl(Ԥ�����, 2) = 1 Having Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0
    Group By NO
    Order By �տ�ʱ��;
  --���ܣ�����һ�в��˹Һŷ��ã�����������ܵ�����Ԥ����¼
  --       ͬʱ������صĻ��ܱ�(���˹ҺŻ��ܡ����û���)
  --       ��һ�з��ô���Ʊ��ʹ�����(����ID_IN>0)

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_�ŶӺ��� �ŶӽкŶ���.�ŶӺ���%Type;
  v_�ֽ�     ���㷽ʽ.����%Type;
  v_�����ʻ� ���㷽ʽ.����%Type;
  v_�������� �ŶӽкŶ���.��������%Type;

  n_��ʱ��       Number;
  n_ԭʼ��ʱ��   Number;
  n_ʱ���޺�     Number;
  n_ʱ����Լ     Number;
  d_ʱ��ʱ��     Date;
  d_������ʱ�� Date;
  n_׷�Ӻ�       Number := 0; --����ʱ�ι��� ׷�ӹҺŵ����
  n_��Լ��       ���˹ҺŻ���.��Լ��%Type;
  n_�ѽ���       ���˹ҺŻ���.�����ѽ���%Type;
  n_ԤԼ��Чʱ�� Number;
  n_ʧЧ��       Number;
  n_ʧԼ�Һ�     Number := 0;
  n_��������     Number;
  n_����         Number := 0;

  n_��ӡid   Ʊ�ݴ�ӡ����.Id%Type;
  n_����id   ������ü�¼.Id%Type;
  n_Ԥ����� ����Ԥ����¼.���%Type;
  n_��ǰ��� ����Ԥ����¼.���%Type;
  n_����ֵ   ����Ԥ����¼.���%Type;
  n_Ԥ��id   ����Ԥ����¼.Id%Type;
  n_���ѿ�id ���ѿ�Ŀ¼.Id%Type;
  n_�Һ�id   ���˹Һż�¼.Id%Type;

  n_��id           ����ɿ����.Id%Type;
  n_���           �Һ����״̬.���%Type;
  n_�������       �Һ����״̬.���%Type;
  n_��ſ���       �ҺŰ���.��ſ���%Type;
  n_����̨ǩ���Ŷ� Number;
  n_Count          Number;
  n_�޺���         Number(18);
  n_���ƿ�         Number;
  d_�Ŷ�ʱ��       Date;
  v_���㷽ʽ��¼   Varchar2(1000);
  d_���ʱ��       Date;
  v_�ѱ�           �ѱ�.����%Type;
  n_ʱ�����       Number := -1;
  n_ԤԼ���ɶ���   Number;
  v_���㷽ʽ       ���㷽ʽ.����%Type;
  v_��������       Varchar2(1000);
  v_��ǰ����       Varchar2(200);
  v_�������       ����Ԥ����¼.�������%Type;
  n_������       ����Ԥ����¼.��Ԥ��%Type;
  n_��������־     Number(2);
  n_����id         �ҺŰ���.Id%Type;
  n_ԤԼ˳���     �ٴ�������ſ���.ԤԼ˳���%Type;
  n_�ƻ�id         �ҺŰ��żƻ�.Id%Type := 0;
  v_����           �ҺŰ�������.������Ŀ%Type;
  n_��Լ��         Number(18);
  n_�ѹ���         Number(4) := 0;
  n_Exists         Number;
  n_�ҳ��������� Number(4) := 0;
  n_��ʱ����ʾ     Number(3);
  n_����ģʽ       ������Ϣ.����ģʽ%Type;
  v_�Ŷ����       �ŶӽкŶ���.�Ŷ����%Type;
  v_������         �Һ����״̬.������%Type;
  v_��Ų���Ա     �Һ����״̬.����Ա����%Type;
  v_��Ż�����     �Һ����״̬.������%Type;
  v_���ʽ       ���˹Һż�¼.ҽ�Ƹ��ʽ%Type;
  n_״̬           �ٴ�������ſ���.�Һ�״̬%Type;
Begin
 
  --��ȡ��ǰ��������
  Select Terminal Into v_������ From V$session Where Audsid = Userenv('sessionid');
  n_��id := Zl_Get��id(����Ա����_In);

  If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
    --�ҺŻ���ԤԼ����
    --��Ϊ�����а��ձ൥�ݺŹ���,�չҺ������ܳ���10000��,����Ҫ���ΨһԼ����
    Select Count(*)
    Into n_Count
    From ������ü�¼
    Where ��¼���� = 4 And ��¼״̬ In (1, 3) And ��� = ���_In And NO = ���ݺ�_In;
    If n_Count <> 0 Then
      v_Err_Msg := '�Һŵ��ݺ��ظ�,���ܱ��棡' || Chr(13) || '���ʹ���˰���˳����,���չҺ������ܳ���10000�˴Ρ�';
      Raise Err_Item;
    End If;
  
    --��ȡ���㷽ʽ����
    Begin
      Select ���� Into v_�ֽ� From ���㷽ʽ Where ���� = 1;
    Exception
      When Others Then
        v_�ֽ� := '�ֽ�';
    End;
    Begin
      Select ���� Into v_�����ʻ� From ���㷽ʽ Where ���� = 3;
    Exception
      When Others Then
        v_�����ʻ� := '�����ʻ�';
    End;
  End If;

  n_��� := ����_In;
  v_�ѱ� := �ѱ�_In;
  d_ʱ��ʱ�� := ����ʱ��_In;

  If Nvl(���_In, 0) = 1 Then
    If �����¼id_In Is Not Null Then
      Begin
        Select 1
        Into n_Exists
        From �ٴ������¼
        Where ID = �����¼id_In And Nvl(�Ƿ񷢲�, 0) = 1 And Nvl(�Ƿ�����, 0) = 0;
      Exception
        When Others Then
          v_Err_Msg := '�޷�ȷ�������¼����������¼�Ƿ���ڻ�������';
          Raise Err_Item;
      End;
    End If;

    If �ѱ�_In Is Null Then
      Begin
        Select ���� Into v_�ѱ� From �ѱ� Where ȱʡ��־ = 1 And Rownum < 2;
      Exception
        When Others Then
          v_Err_Msg := '�޷�ȷ�����˷ѱ�����ȱʡ�ѱ��Ƿ���ȷ���ã�';
          Raise Err_Item;
      End;
    End If;
    If Nvl(�������˷ѱ�_In, 0) = 1 And v_�ѱ� Is Not Null Then
      Update ������Ϣ Set �ѱ� = v_�ѱ� Where ����id = ����id_In;
    End If;
  
    If Nvl(������������_In, 0) = 1 Then
      Update ������Ϣ Set ���� = ����_In Where ����id = ����id_In;
    End If;
  
    If �����_In Is Not Null Then
      Update ������Ϣ Set ����� = �����_In Where ����id = ����id_In And Nvl(�����, 0) = 0;
    End If;
  
    Update �ٴ�������ſ���
    Set �Һ�״̬ = 0
    Where ��¼id = �����¼id_In And ��� = ����_In And Nvl(�Һ�״̬, 0) = 3 And ����Ա���� = ����Ա����_In;
  
    --��ȡ�Ƿ��ʱ��
    Begin
      Select Nvl(�Ƿ��ʱ��, 0), Nvl(�Ƿ���ſ���, 0), �޺���, ��Լ��, �ѹ���, ��Լ��
      Into n_��ʱ��, n_��ſ���, n_�޺���, n_��Լ��, n_�ѹ���, n_��Լ��
      From �ٴ������¼
      Where ID = �����¼id_In;
      n_ԭʼ��ʱ�� := n_��ʱ��;
    Exception
      When Others Then
        n_��ʱ��     := 0;
        n_ԭʼ��ʱ�� := n_��ʱ��;
        n_��ſ���   := 0;
        n_�޺���     := Null;
        n_��Լ��     := Null;
    End;
  
    --��ȡ��ǰδʹ�õ����
    If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
      n_ԤԼ��Чʱ�� := Zl_To_Number(Zl_Getsysparameter('ԤԼ��Чʱ��', 1111));
      n_ʧԼ�Һ�     := Zl_To_Number(Zl_Getsysparameter('ʧԼ���ڹҺ�', 1111));
    End If;
    n_ʧЧ�� := 0;
  
    If Nvl(n_��ſ���, 0) = 1 And n_��ʱ�� = 0 Then
      --<��ſ��� δ����ʱ�� ��ȡ���õ�������,�Լ��Ѿ�ʹ�õ�����>
      If Nvl(ԤԼ�Һ�_In, 0) = 0 And Nvl(n_ԤԼ��Чʱ��, 0) <> 0 And Nvl(n_ʧԼ�Һ�, 0) > 0 Then
        Select Sum(Decode(Sign((Sysdate - n_ԤԼ��Чʱ�� / 24 / 60) - ԤԼʱ��), 1, 1, 0))
        Into n_ʧЧ��
        From ���˹Һż�¼
        Where �����¼id = �����¼id_In And ��¼״̬ = 1 And ��¼���� = 2;
      End If;
      If n_��� Is Null Then
        n_������� := Null;
        If n_ԭʼ��ʱ�� = 0 Then
          Select Min(���) Into n_������� From �ٴ�������ſ��� Where ��¼id = �����¼id_In And Nvl(�Һ�״̬, 0) = 0;
        End If;
        If n_������� Is Null Then
          Select Nvl(Max(���), 0) + 1 Into n_������� From �ٴ�������ſ��� Where ��¼id = �����¼id_In;
        End If;
        n_��� := Nvl(n_�������, 0);
      End If;
      --<��ſ��� δ����ʱ�� ��ȡ���õ������� �Լ��Ѿ�ʹ�õ����� --end
      --�ǼӺŵ������Ҫ����Ƿ񳬹�������
      If ��������_In = 0 Then
        If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
          --�Һ�ʱ���
          --������ſ���δ��ʱ�� �ﵽ������
          If n_�޺��� <= n_�ѹ��� And n_�޺��� > 0 Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(Trunc(����ʱ��_In), 'yyyy-mm-dd ') || '�Ѵﵽ����޺�����';
            Raise Err_Item;
          End If;
        Else
          --ԤԼʱ���
          If n_��Լ�� = 0 Then
            n_��Լ�� := n_�޺���;
          End If;
        
          If n_��Լ�� <= n_��Լ�� And n_��Լ�� > 0 Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(Trunc(����ʱ��_In), 'yyyy-mm-dd') || '�Ѵﵽ�����Լ����';
            Raise Err_Item;
          End If;
        End If;
      Else
        Null;
        --������ſ���,δ��ʱ�� �Ӻ����   ������,����Ժ������������Ժ󲹳�
      End If;
    Elsif Nvl(n_��ʱ��, 0) > 0 And Nvl(n_��ſ���, 0) = 0 Then
      --�Һ�ʱ��� �Ƿ����ʱ��,����ʱ��,��ʱ�ε�����������ȡ����
      If n_��� Is Null And Nvl(ԤԼ�Һ�_In, 0) = 1 Then
        Begin
          Select ���
          Into n_���
          From �ٴ�������ſ���
          Where ��¼id = �����¼id_In And ��ʼʱ�� = ����ʱ��_In And Rownum < 2;
        Exception
          When Others Then
            n_��� := Null;
        End;
      End If;
    
      If Nvl(ԤԼ�Һ�_In, 0) = 1 Then
        Begin
          Select Nvl(���, 0),
                 To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
                 ����, Decode(Nvl(�Ƿ�ԤԼ, 0), 0, 0, ����)
          Into n_ʱ�����, d_ʱ��ʱ��, n_ʱ���޺�, n_ʱ����Լ
          From �ٴ�������ſ���
          Where ��¼id = �����¼id_In And ��� = n_��� And ԤԼ˳��� Is Null;
        Exception
          When Others Then
            n_ʱ����� := -1;
            n_��ʱ��   := 0;
            d_ʱ��ʱ�� := ����ʱ��_In;
            n_ʱ���޺� := 0;
            n_ʱ����Լ := 0;
        End;
      End If;
    
      --<--��ͨ�ŷ�ʱ�� ����ֻ��ԤԼһ�����-->
      If ��������_In = 0 And Nvl(ԤԼ�Һ�_In, 0) = 1 Then
        --<����ԤԼ�Һ�-->
      
        Select Nvl(Sum(Decode(Nvl(Sign(a.��ʼʱ�� - d_ʱ��ʱ��), 0), 0, 1, 0)), 0) As ��Լ��
        Into n_��Լ��
        From �ٴ�������ſ��� A
        Where ��¼id = �����¼id_In And �Һ�״̬ Not In (0, 4, 5);
      
        n_ʱ����Լ := n_ʱ���޺�; --��ͨ�ŷ�ʱ�ε����,29��n_ʱ����Լʼ����0 �������⴦��
        --�����������
        If n_��Լ�� = 0 Then
          n_��Լ�� := n_�޺���;
        End If;
        If n_ʱ����Լ <= n_��Լ�� Or n_��Լ�� <= n_��Լ�� Then
          v_Err_Msg := '�ű�' || �ű�_In || '��ʱ��' || To_Char(d_ʱ��ʱ��, 'hh24:mi') || '�Ѵﵽ�����Լ����';
          Raise Err_Item;
        End If;
      Elsif ��������_In = 0 And n_�޺��� <= n_�ѹ��� - n_ʧЧ�� Then
        v_Err_Msg := '�ű�' || �ű�_In || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
        Raise Err_Item;
      End If;
      If Nvl(ԤԼ�Һ�_In, 0) = 1 Then
        --��ȡ����ҳ���������
        Select Nvl(Max(���), 0)
        Into n_�ҳ���������
        From �ٴ�������ſ��� A
        Where ��¼id = �����¼id_In And ԤԼ˳��� Is Null And �Һ�״̬ Not In (0, 5);
        If ԤԼ˳���_In Is Not Null Then
          n_ԤԼ˳��� := ԤԼ˳���_In;
        Else
        
          Select Nvl(Max(ԤԼ˳���), 0) + 1
          Into n_ԤԼ˳���
          From �ٴ�������ſ���
          Where ��¼id = �����¼id_In And ��� = n_ʱ����� And ԤԼ˳��� Is Not Null;
        
        End If;
        --�������
        n_��� := RPad(Nvl(n_ʱ�����, 0), Length(n_�޺���) + Length(Nvl(n_ʱ�����, 0)), 0) + n_ԤԼ˳���;
        If n_ԤԼ˳��� Is Null Then
          n_��� := Nvl(n_�ҳ���������, 0) + 1;
        End If;
      End If;
      --<--��ͨ�ŷ�ʱ��--End>
    Elsif Nvl(n_��ʱ��, 0) > 0 And Nvl(n_��ſ���, 0) = 1 Then
      --<������ſ��� ����ʱ��
      --ר�Һŷ�ʱ��
      Begin
        Select Nvl(���, 0),
               To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
               ����, Decode(Nvl(�Ƿ�ԤԼ, 0), 0, 0, ����)
        Into n_ʱ�����, d_ʱ��ʱ��, n_ʱ���޺�, n_ʱ����Լ
        From �ٴ�������ſ���
        Where ��¼id = �����¼id_In And ��� = n_���;
      Exception
        When Others Then
          n_ʱ����� := -1;
          n_��ʱ��   := 0;
          d_ʱ��ʱ�� := ����ʱ��_In;
          n_ʱ���޺� := 0;
          n_ʱ����Լ := 0;
      End;
    
      Select Max(���), Sum(Decode(���, Nvl(����_In, 0), 0, 1)), Sum(Decode(Sign(��ʼʱ�� - d_ʱ��ʱ��), 0, 1, 0))
      Into n_�������, n_�ѹ���, n_��������
      From �ٴ�������ſ���
      Where ��¼id = �����¼id_In And �Һ�״̬ Not In (0, 4, 5);
    
      n_ʧЧ�� := 0;
      If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
        If Nvl(n_ԤԼ��Чʱ��, 0) <> 0 And Nvl(n_ʧԼ�Һ�, 0) > 0 Then
        
          Select Sum(Decode(Sign((Sysdate - n_ԤԼ��Чʱ�� / 24 / 60) - ��ʼʱ��), 1, 1, 0))
          Into n_ʧЧ��
          From �ٴ�������ſ���
          Where ��¼id = �����¼id_In And ��ʼʱ�� Between Trunc(Sysdate) And Sysdate And Nvl(�Һ�״̬, 0) = 2;
        
        End If;
      End If;
    
      If ��������_In = 0 Then
        If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
        
          --��ʱ�ιҺ�ʱ�ж��Ƿ��ǹ��ڹҺ� Ҳ����׷�Ӻŵ���� Ŀǰֻ���ר�Һŷ�ʱ�ν��д���
          If ����_In Is Null Then
          
            Select Max(To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'),
                                'yyyy-mm-dd hh24:mi:ss'))
            Into d_������ʱ��
            From �ٴ�������ſ���
            Where ��¼id = �����¼id_In And Nvl(����, 0) <> 0;
          
            n_׷�Ӻ� := Case Sign(����ʱ��_In - d_������ʱ��)
                       When -1 Then
                        0
                       Else
                        1
                     End;
          
          End If;
        
          --�Һ� ׷�Ӻ���ʱ�����ʱ���޺���
          If n_ʱ���޺� <= n_�������� And Nvl(n_׷�Ӻ�, 0) = 0 Then
            v_Err_Msg := '�ű�' || �ű�_In || '��ʱ��' || To_Char(d_ʱ��ʱ��, 'hh24:mi') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
          If n_�޺��� <= n_�ѹ��� - n_ʧЧ�� Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        Else
          --ԤԼ�Һ�
          If n_��Լ�� = 0 Then
            n_��Լ�� := n_�޺���;
          End If;
          If n_��Լ�� <= n_�������� Then
            v_Err_Msg := '�ű�' || �ű�_In || '��ʱ��' || To_Char(d_ʱ��ʱ��, 'hh24:mi') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
          If n_�޺��� <= n_�ѹ��� Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        End If;
      End If;
      If n_��� Is Null Then
        --�������
        If Nvl(n_�������, 0) < Nvl(n_ʱ�����, 0) Then
          n_������� := Nvl(n_ʱ�����, 0);
        End If;
        n_��� := Nvl(n_�������, 0) + 1;
      End If;
    Elsif Nvl(n_��ʱ��, 0) = 0 And Nvl(n_��ſ���, 0) = 0 And Nvl(������_In, 0) = 0 And Nvl(�ű�_In, 0) > 0 Then
      ---<--��ͨ��  -->
      If Nvl(��������_In, 0) = 0 Then
        If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
          --�Һ�
          If Nvl(n_�ѹ���, 0) >= Nvl(n_�޺���, 0) And Nvl(n_�޺���, 0) > 0 Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        Else
          --ԤԼ
          If (Nvl(n_��Լ��, 0) > 0 And Nvl(n_��Լ��, 0) >= Nvl(n_��Լ��, 0)) Or
             (Nvl(n_�ѹ���, 0) >= Nvl(n_�޺���, 0) And Nvl(n_�޺���, 0) > 0) Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������';
            Raise Err_Item;
          End If;
        End If;
      End If;
      ---<--��ͨ��  -->
    End If;
  
    --���¹Һ����״̬
    If Not n_��� Is Null Then
      If n_��ʱ�� = 1 Then
        d_���ʱ�� := ����ʱ��_In;
      Else
        d_���ʱ�� := Trunc(����ʱ��_In);
      End If;
    
      --������ŵĴ���
      Begin
        If n_ԤԼ˳��� Is Null Then
          Select ����Ա����, ����վ����
          Into v_��Ų���Ա, v_��Ż�����
          From �ٴ�������ſ���
          Where Nvl(�Һ�״̬, 0) = 5 And ��¼id = �����¼id_In And ��� = n_���;
        Else
          Select ����Ա����, ����վ����
          Into v_��Ų���Ա, v_��Ż�����
          From �ٴ�������ſ���
          Where Nvl(�Һ�״̬, 0) = 5 And ��¼id = �����¼id_In And ��� = n_ʱ����� And ԤԼ˳��� = n_ԤԼ˳���;
        End If;
        n_���� := 1;
      Exception
        When Others Then
          v_��Ų���Ա := Null;
          v_��Ż����� := Null;
          n_����       := 0;
      End;
    
      If n_���� = 0 Then
        If n_ԤԼ˳��� Is Null Then
          Update �ٴ�������ſ���
          Set �Һ�״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1)
          Where ��¼id = �����¼id_In And ��� = n_��� And Nvl(�Һ�״̬, 0) = 3 And ����Ա���� = ����Ա����_In;
        Else
          Update �ٴ�������ſ���
          Set �Һ�״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1)
          Where ��¼id = �����¼id_In And ��� = n_��� And ԤԼ˳��� = n_ԤԼ˳��� And Nvl(�Һ�״̬, 0) = 3 And ����Ա���� = ����Ա����_In;
        End If;
      
        If Sql%Rowcount = 0 Then
          Begin
            If Nvl(n_��ʱ��, 0) > 0 Then
              If Nvl(n_��ſ���, 0) = 1 Then
                --��ʱ�κ�ר�Һ� ʧԼ��ԤԼ������Һ�
                Update �ٴ�������ſ���
                Set �Һ�״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1), ����Ա���� = ����Ա����_In
                Where ��¼id = �����¼id_In And ��� = n_��� And Nvl(�Һ�״̬, 0) In (0, 2);
                If Sql%NotFound Then
                  Begin
                    Select �Һ�״̬ Into n_״̬ From �ٴ�������ſ��� Where ��¼id = �����¼id_In And ��� = n_���;
                  Exception
                    When Others Then
                      n_״̬ := -1;
                  End;
                
                  If n_״̬ <> -1 Then
                    v_Err_Msg := '���' || n_��� || '�ѱ�ʹ��,������ѡ��һ�����.';
                    Raise Err_Item;
                  End If;
                
                  Insert Into �ٴ�������ſ���
                    (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, �Һ�״̬, ����ʱ��, ����, ����, ����Ա����, ��ע)
                    Select �����¼id_In, n_���, d_���ʱ��, d_���ʱ��, 1, Decode(ԤԼ�Һ�_In, 1, 1, 0), Decode(ԤԼ�Һ�_In, 1, 2, 1), Null,
                           Null, Null, ����Ա����_In, '׷�Ӻ�'
                    From Dual;
                End If;
              Else
                If Nvl(ԤԼ����_In, 0) = 1 Then
                  Insert Into �ٴ�������ſ���
                    (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, �Һ�״̬, ����ʱ��, ����, ����, ����Ա����, ��ע, ԤԼ˳���)
                    Select ��¼id, ���, ��ʼʱ��, ��ֹʱ��, 1, 1, Decode(ԤԼ�Һ�_In, 1, 2, 1), Null, Null, Null, ����Ա����_In, n_���,
                           n_ԤԼ˳���
                    From �ٴ�������ſ���
                    Where ��¼id = �����¼id_In And ��� = n_ʱ����� And ԤԼ˳��� Is Null;
                End If;
              End If;
            Else
              If Nvl(n_��ſ���, 0) = 1 Then
                Update �ٴ�������ſ���
                Set �Һ�״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1), ����Ա���� = ����Ա����_In
                Where ��¼id = �����¼id_In And ��� = n_��� And Nvl(�Һ�״̬, 0) = 0;
              
                If Sql%Rowcount = 0 Then
                  Begin
                    Select �Һ�״̬ Into n_״̬ From �ٴ�������ſ��� Where ��¼id = �����¼id_In And ��� = n_���;
                  Exception
                    When Others Then
                      n_״̬ := -1;
                  End;
                  If n_״̬ <> -1 Then
                    v_Err_Msg := '���' || n_��� || '�ѱ�ʹ��,������ѡ��һ�����.';
                    Raise Err_Item;
                  End If;
                  Insert Into �ٴ�������ſ���
                    (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, �Һ�״̬, ����ʱ��, ����, ����, ����Ա����, ��ע)
                    Select �����¼id_In, n_���, ����ʱ��_In, ����ʱ��_In, 1, Decode(ԤԼ�Һ�_In, 1, 1, 0), Decode(ԤԼ�Һ�_In, 1, 2, 1),
                           Null, Null, Null, ����Ա����_In, '׷�Ӻ�'
                    From Dual;
                
                End If;
              End If;
            End If;
          Exception
            When Others Then
              v_Err_Msg := '���' || n_��� || '�ѱ�ʹ��,������ѡ��һ�����.';
              Raise Err_Item;
          End;
        End If;
      Else
        If ����Ա����_In <> v_��Ų���Ա Or v_������ <> v_��Ż����� Then
          v_Err_Msg := '���' || n_��� || '�ѱ�����վ��(' || v_������ || ')����,������ѡ��һ�����.';
          Raise Err_Item;
        End If;
        If n_ԤԼ˳��� Is Null Then
          Update �ٴ�������ſ���
          Set �Һ�״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1)
          Where ��¼id = �����¼id_In And ��� = n_��� And Nvl(�Һ�״̬, 0) = 5 And ����Ա���� = ����Ա����_In And ����վ���� = v_������;
        Else
          Update �ٴ�������ſ���
          Set �Һ�״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1)
          Where ��¼id = �����¼id_In And ��� = n_ʱ����� And ԤԼ˳��� = n_ԤԼ˳��� And Nvl(�Һ�״̬, 0) = 5 And ����Ա���� = ����Ա����_In And
                ����վ���� = v_������;
        End If;
      End If;
    End If;
  End If;


  --�������˹Һŷ���(���ܵ����ǻ������������)
  Select ���˷��ü�¼_Id.Nextval Into n_����id From Dual; --Ӧ��ͨ������õ�

  Insert Into ������ü�¼
    (ID, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, NO, ʵ��Ʊ��, �����־, �Ӱ��־, ���ӱ�־, ��ҩ����, ����id, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, ���˿���id, �շ����,
     ���㵥λ, �շ�ϸĿid, ������Ŀid, �վݷ�Ŀ, ����, ����, ��׼����, Ӧ�ս��, ʵ�ս��, ���ʽ��, ����id, ���ʷ���, ��������id, ������, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����,
     ����ʱ��, �Ǽ�ʱ��, ���մ���id, ������Ŀ��, ���ձ���, ͳ����, ժҪ, ����, �ɿ���id)
  Values
    (n_����id, 4, Decode(ԤԼ�Һ�_In, 1, 0, 1), ���_In, Decode(�۸񸸺�_In, 0, Null, �۸񸸺�_In), ��������_In, ���ݺ�_In, Ʊ�ݺ�_In, 1, ����_In,
     ������_In, Decode(ԤԼ�Һ�_In, 1, To_Char(n_���), ����_In), Decode(����id_In, 0, Null, ����id_In),
     Decode(�����_In, 0, Null, �����_In), ���ʽ_In, ����_In, Decode(����_In, Null, Null, �Ա�_In), Decode(����_In, Null, Null, ����_In),
     v_�ѱ�, ���˿���id_In, �շ����_In, �ű�_In, �շ�ϸĿid_In, ������Ŀid_In, �վݷ�Ŀ_In, 1, ����_In, ��׼����_In, Ӧ�ս��_In, ʵ�ս��_In,
     Decode(ԤԼ�Һ�_In, 1, Null, Decode(Nvl(���ʷ���_In, 0), 1, Null, ʵ�ս��_In)),
     Decode(ԤԼ�Һ�_In, 1, Null, Decode(Nvl(���ʷ���_In, 0), 1, Null, ����id_In)), Decode(Nvl(���ʷ���_In, 0), 1, 1, 0), ��������id_In,
     ����Ա����_In, Decode(ԤԼ�Һ�_In, 1, ����Ա����_In, Null), ִ�в���id_In, ҽ������_In, ����Ա���_In, ����Ա����_In, ����ʱ��_In, �Ǽ�ʱ��_In, ���մ���id_In,
     ������Ŀ��_In, ���ձ���_In, ͳ����_In, ժҪ_In, ԤԼ��ʽ_In, Decode(ԤԼ�Һ�_In, 1, Null, n_��id));

  --���ܽ��㵽����Ԥ����¼
  If Nvl(ԤԼ�Һ�_In, 0) = 0 And Nvl(���ʷ���_In, 0) = 0 Then
    If Nvl(�ֽ�֧��_In, 0) = 0 And Nvl(����֧��_In, 0) = 0 And Nvl(Ԥ��֧��_In, 0) = 0 And ���_In = 1 Then
      Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      Insert Into ����Ԥ����¼
        (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ,
         ��������)
      Values
        (n_Ԥ��id, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), v_�ֽ�, 0, �Ǽ�ʱ��_In, ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�',
         n_��id, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, ������λ_In, 4);
    End If;
    If Nvl(�ֽ�֧��_In, 0) <> 0 And ���_In = 1 Then
      v_��������     := ���㷽ʽ_In || '|'; --�Կո�ֿ���|��β,û�н�������
      v_���㷽ʽ��¼ := '';
      While v_�������� Is Not Null Loop
        v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '|') - 1);
        v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
      
        v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
        n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1));
      
        v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
        v_������� := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
      
        v_��ǰ����   := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
        n_��������־ := To_Number(v_��ǰ����);
      
        If Instr('|' || v_���㷽ʽ��¼ || '|', '|' || Nvl(v_���㷽ʽ, v_�ֽ�) || '|') <> 0 Then
          v_Err_Msg := 'ʹ�����ظ��Ľ��㷽ʽ,����!';
          Raise Err_Item;
        Else
          v_���㷽ʽ��¼ := v_���㷽ʽ��¼ || '|' || Nvl(v_���㷽ʽ, v_�ֽ�);
        End If;
      
        If n_��������־ = 0 Then
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          Insert Into ����Ԥ����¼
            (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��,
             ������λ, ��������, �������)
          Values
            (n_Ԥ��id, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), Nvl(v_���㷽ʽ, v_�ֽ�), Nvl(n_������, 0), �Ǽ�ʱ��_In,
             ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�', n_��id, Null, Null, Null, Null, Null, ������λ_In, 4, v_�������);
        Else
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          Insert Into ����Ԥ����¼
            (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��,
             ������λ, ��������, �������)
          Values
            (n_Ԥ��id, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), Nvl(v_���㷽ʽ, v_�ֽ�), Nvl(n_������, 0), �Ǽ�ʱ��_In,
             ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�', n_��id, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, ������λ_In, 4,
             v_�������);
          If Nvl(���㿨���_In, 0) <> 0 And Nvl(�ֽ�֧��_In, 0) <> 0 Then
            n_���ѿ�id := Null;
            Begin
              Select Nvl(���ƿ�, 0), 1 Into n_���ƿ�, n_Count From �����ѽӿ�Ŀ¼ Where ��� = ���㿨���_In;
            Exception
              When Others Then
                n_Count := 0;
            End;
            If n_Count = 0 Then
              v_Err_Msg := 'û�з���ԭ���㿨����Ӧ���,���ܼ���������';
              Raise Err_Item;
            End If;
            If n_���ƿ� = 1 Then
              Select ID
              Into n_���ѿ�id
              From ���ѿ�Ŀ¼
              Where �ӿڱ�� = ���㿨���_In And ���� = ����_In And
                    ��� = (Select Max(���) From ���ѿ�Ŀ¼ Where �ӿڱ�� = ���㿨���_In And ���� = ����_In);
            End If;
            Zl_���˿������¼_Insert(���㿨���_In, n_���ѿ�id, Nvl(v_���㷽ʽ, v_�ֽ�), Nvl(n_������, 0), ����_In, Null, �Ǽ�ʱ��_In, Null, ����id_In,
                              n_Ԥ��id);
          End If;
        End If;
      
        If Nvl(���½������_In, 0) = 0 Then
          Update ��Ա�ɿ����
          Set ��� = Nvl(���, 0) + n_������
          Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = Nvl(v_���㷽ʽ, v_�ֽ�)
          Returning ��� Into n_����ֵ;
        
          If Sql%RowCount = 0 Then
            Insert Into ��Ա�ɿ����
              (�տ�Ա, ���㷽ʽ, ����, ���)
            Values
              (����Ա����_In, Nvl(v_���㷽ʽ, v_�ֽ�), 1, n_������);
            n_����ֵ := n_������;
          End If;
          If Nvl(n_����ֵ, 0) = 0 Then
            Delete From ��Ա�ɿ����
            Where �տ�Ա = ����Ա����_In And ���㷽ʽ = Nvl(v_���㷽ʽ, v_�ֽ�) And ���� = 1 And Nvl(���, 0) = 0;
          End If;
        End If;
      
        v_�������� := Substr(v_��������, Instr(v_��������, '|') + 1);
      End Loop;
    End If;
  
    --����ҽ���Һ�
    If Nvl(����֧��_In, 0) <> 0 And ���_In = 1 Then
      Insert Into ����Ԥ����¼
        (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, ��������)
      Values
        (����Ԥ����¼_Id.Nextval, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), v_�����ʻ�, ����֧��_In, �Ǽ�ʱ��_In, ����Ա���_In,
         ����Ա����_In, ����id_In, 'ҽ���Һ�', n_��id, 4);
    End If;
  
    --���ھ��￨ͨ��Ԥ����Һ�
    If Nvl(Ԥ��֧��_In, 0) <> 0 And ���_In = 1 Then
      n_Ԥ����� := Ԥ��֧��_In;
      For r_Deposit In c_Deposit(����id_In) Loop
        n_��ǰ��� := Case
                    When r_Deposit.��� - n_Ԥ����� < 0 Then
                     r_Deposit.���
                    Else
                     n_Ԥ�����
                  End;
      
        If r_Deposit.����id = 0 Then
          --��һ�γ�Ԥ��(���Ͻ���ID,���Ϊ0)
          Update ����Ԥ����¼ Set ��Ԥ�� = 0, ����id = ����id_In, �������� = 4 Where ID = r_Deposit.ԭԤ��id;
        
        End If;
        --���ϴ�ʣ���
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���,
           ��Ԥ��, ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
          Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�,
                 �Ǽ�ʱ��_In, ����Ա����_In, ����Ա���_In, n_��ǰ���, ����id_In, n_��id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
          From ����Ԥ����¼
          Where NO = r_Deposit.No And ��¼״̬ = r_Deposit.��¼״̬ And ��¼���� In (1, 11) And Rownum = 1;
      
        --���²���Ԥ�����
        Update �������
        Set Ԥ����� = Nvl(Ԥ�����, 0) - n_��ǰ���
        Where ����id = ����id_In And ���� = 1 And ���� = Nvl(1, 2);
        --����Ƿ��Ѿ�������
        If r_Deposit.��� < n_Ԥ����� Then
          n_Ԥ����� := n_Ԥ����� - r_Deposit.���;
        Else
          n_Ԥ����� := 0;
        End If;
      
        If n_Ԥ����� = 0 Then
          Exit;
        End If;
      End Loop;
      If n_Ԥ����� > 0 Then
        v_Err_Msg := 'Ԥ���಻��֧������֧�����,���ܼ���������';
        Raise Err_Item;
      
      End If;
      Delete From ������� Where ����id = ����id_In And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
    End If;
  
    --��ػ��ܱ�Ĵ���
    --��Ա�ɿ����
    If ���_In = 1 And Nvl(���½������_In, 0) = 0 Then
      If Nvl(����֧��_In, 0) <> 0 Then
        Update ��Ա�ɿ����
        Set ��� = Nvl(���, 0) + ����֧��_In
        Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = v_�����ʻ�
        Returning ��� Into n_����ֵ;
      
        If Sql%RowCount = 0 Then
          Insert Into ��Ա�ɿ���� (�տ�Ա, ���㷽ʽ, ����, ���) Values (����Ա����_In, v_�����ʻ�, 1, ����֧��_In);
          n_����ֵ := ����֧��_In;
        End If;
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete From ��Ա�ɿ����
          Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_�����ʻ� And Nvl(���, 0) = 0;
        End If;
      End If;
    End If;
  End If;

  --���˹ҺŻ���(ֻ����һ��,�ҵ�����ȡ�����Ѳ�����)
  If Nvl(ԤԼ�Һ�_In, 0) = 0 Then
    --���˱��ξ���(�Է���ʱ��Ϊ׼)
    If Nvl(����id_In, 0) <> 0 And ���_In = 1 Then
      Update ������Ϣ Set ����ʱ�� = ����ʱ��_In, ����״̬ = 1, �������� = ����_In Where ����id = ����id_In;
    End If;
  End If;

  If Nvl(ԤԼ�Һ�_In, 0) = 0 And Nvl(���ʷ���_In, 0) = 1 Then
    --����
    If Nvl(����id_In, 0) = 0 Then
      v_Err_Msg := 'Ҫ��Բ��˵ĹҺŷѽ��м��ʣ������ǽ������˲��ܼ��ʹҺš�';
      Raise Err_Item;
    End If;
  
    --�������
    Update �������
    Set ������� = Nvl(�������, 0) + Nvl(ʵ�ս��_In, 0)
    Where ����id = Nvl(����id_In, 0) And ���� = 1 And ���� = 1;
  
    If Sql%RowCount = 0 Then
      Insert Into ������� (����id, ����, ����, �������, Ԥ�����) Values (����id_In, 1, 1, Nvl(ʵ�ս��_In, 0), 0);
    End If;
  
    --����δ�����
    Update ����δ�����
    Set ��� = Nvl(���, 0) + Nvl(ʵ�ս��_In, 0)
    Where ����id = ����id_In And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���˿���id, 0) = Nvl(���˿���id_In, 0) And
          Nvl(��������id, 0) = Nvl(��������id_In, 0) And Nvl(ִ�в���id, 0) = Nvl(ִ�в���id_In, 0) And ������Ŀid + 0 = ������Ŀid_In And
          ��Դ;�� + 0 = 1;
  
    If Sql%RowCount = 0 Then
      Insert Into ����δ�����
        (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
      Values
        (����id_In, Null, Null, ���˿���id_In, ��������id_In, ִ�в���id_In, ������Ŀid_In, 1, Nvl(ʵ�ս��_In, 0));
    End If;
  End If;

  --���˹Һż�¼
  If �ű�_In Is Not Null And ���_In = 1 Then
    --And Nvl(ԤԼ�Һ�_In, 0) = 0
    Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
    Begin
      Select ���� Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = ���ʽ_In And Rownum < 2;
    Exception
      When Others Then
        v_���ʽ := Null;
    End;
    Insert Into ���˹Һż�¼
      (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ԤԼʱ��, ����Ա���,
       ����Ա����, ����, ����, ����, ԤԼ, ԤԼ��ʽ, ժҪ, ������ˮ��, ����˵��, ������λ, ����ʱ��, ������, ԤԼ����Ա, ԤԼ����Ա���, ����, ҽ�Ƹ��ʽ, �����¼id, �շѵ�)
    Values
      (n_�Һ�id, ���ݺ�_In, Decode(Nvl(ԤԼ�Һ�_In, 0), 1, 2, 1), 1, ����id_In, �����_In, ����_In, �Ա�_In, ����_In, �ű�_In, ����_In, ����_In,
       Null, ִ�в���id_In, ҽ������_In, 0, Null, �Ǽ�ʱ��_In, ����ʱ��_In, Decode(Nvl(ԤԼ�Һ�_In, 0), 1, ����ʱ��_In, Null), ����Ա���_In,
       ����Ա����_In, ����_In, n_���, ����_In, Decode(ԤԼ����_In, 1, 1, 0), ԤԼ��ʽ_In, ժҪ_In, ������ˮ��_In, ����˵��_In, ������λ_In,
       Decode(Nvl(ԤԼ�Һ�_In, 0), 0, �Ǽ�ʱ��_In, Null), Decode(Nvl(ԤԼ�Һ�_In, 0), 0, ����Ա����_In, Null),
       Decode(Nvl(ԤԼ�Һ�_In, 0), 1, ����Ա����_In, Null), Decode(Nvl(ԤԼ�Һ�_In, 0), 1, ����Ա���_In, Null),
       Decode(Nvl(����_In, 0), 0, Null, ����_In), v_���ʽ, �����¼id_In, �շѵ�_In);
  
    If Nvl(ԤԼ�Һ�_In, 0) = 0 And ԤԼ��ʽ_In Is Not Null Then
      Update ���˹Һż�¼
      Set ԤԼ = 1, ԤԼʱ�� = ����ʱ��_In, ԤԼ����Ա = ����Ա����_In, ԤԼ����Ա��� = ����Ա���_In
      Where ID = n_�Һ�id;
    End If;
  
    n_ԤԼ���ɶ��� := 0;
    If Nvl(ԤԼ�Һ�_In, 0) = 1 Then
      n_ԤԼ���ɶ��� := Zl_To_Number(zl_GetSysParameter('ԤԼ���ɶ���', 1113));
    End If;
  
    --0-����������;1-��ҽ�������̨�Ŷ�;2-�ȷ���,��ҽ��վ
    If Nvl(���ɶ���_In, 0) <> 0 And Nvl(ԤԼ�Һ�_In, 0) = 0 Or n_ԤԼ���ɶ��� = 1 Then
      n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113));
      If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Or n_ԤԼ���ɶ��� = 1 Then
        n_��ʱ����ʾ := Nvl(Zl_To_Number(zl_GetSysParameter(270)), 0);
        If Nvl(ԤԼ�Һ�_In, 0) = 1 And n_��ʱ����ʾ = 1 And n_��ʱ�� = 1 Then
          n_��ʱ����ʾ := 1;
        Else
          n_��ʱ����ʾ := Null;
        End If;
        --��������
        --.����ִ�в��š� �ķ�ʽ���ɶ���
        v_�������� := ִ�в���id_In;
        v_�ŶӺ��� := Zlgetnextqueue(ִ�в���id_In, n_�Һ�id, �ű�_In || '|' || n_���);
        v_�Ŷ���� := Zlgetsequencenum(0, n_�Һ�id, 0);
        d_�Ŷ�ʱ�� := Zl_Get_Queuedate(n_�Һ�id, �ű�_In, n_���, ����ʱ��_In);
        --  ��������_In , ҵ������_In, ҵ��id_In,����id_In,�ŶӺ���_In,�Ŷӱ��_In,��������_In,����ID_IN, ����_In, ҽ������_In, �Ŷ�ʱ��_In
        Zl_�ŶӽкŶ���_Insert(v_��������, 0, n_�Һ�id, ִ�в���id_In, v_�ŶӺ���, Null, ����_In, ����id_In, ����_In, ҽ������_In, d_�Ŷ�ʱ��, ԤԼ��ʽ_In,
                         n_��ʱ����ʾ, v_�Ŷ����);
      End If;
    End If;
  End If;
  --���˵�����Ϣ
  If ����id_In Is Not Null And ���_In = 1 Then
    --ȡ����:
    If Nvl(n_����ģʽ, 0) <> Nvl(����ģʽ_In, 0) Then
      --����ģʽ��ȷ��
    
      v_Err_Msg := Null;
      Begin
        Select Nvl(����ģʽ, 0) Into n_����ģʽ From ������Ϣ Where ����id = ����id_In;
      Exception
        When Others Then
          v_Err_Msg := 'δ�ҵ�ָ���Ĳ�����Ϣ,������Һ�';
      End;
    
      If v_Err_Msg Is Not Null Then
        Raise Err_Item;
      End If;
      If n_����ģʽ = 1 And Nvl(����ģʽ_In, 0) = 0 Then
        --�����Ѿ���"�����ƺ�����",������"�Ƚ�������Ƶ�",�����Ƿ����δ������
        Select Count(1)
        Into n_Count
        From ����δ�����
        Where ����id = ����id_In And (��Դ;�� In (1, 4) Or ��Դ;�� = 3 And Nvl(��ҳid, 0) = 0) And Nvl(���, 0) <> 0 And Rownum < 2;
        If Nvl(n_Count, 0) <> 0 Then
          --����δ�������ݣ������Ƚ���������ִ��
          v_Err_Msg := '��ǰ���˵ľ���ģʽΪ�����ƺ�����Ҵ���δ����ã�����������ò��˵ľ���ģʽ,������ȶ�δ����ý��ʺ��ٹҺŻ򲻵������˵ľ���ģʽ!';
          Raise Err_Item;
        End If;
        --���
        --δ����ҽ��ҵ��ģ�����ʱ�͹Һŵ�,��Ҫ��֤ͬһ�εľ���ģʽ��һ����(�����Ѿ���飬�����ٴ���)
      End If;
      Update ������Ϣ Set ����ģʽ = ����ģʽ_In Where ����id = ����id_In;
    End If;
  
    Update ������Ϣ
    Set ������ = Null, ������ = Null, �������� = Null
    Where ����id = ����id_In And Nvl(��Ժ, 0) = 0 And Exists
     (Select 1
           From ���˵�����¼
           Where ����id = ����id_In And ��ҳid Is Not Null And
                 �Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From ���˵�����¼ Where ����id = ����id_In));
  
    If Sql%RowCount > 0 Then
      Update ���˵�����¼
      Set ����ʱ�� = Sysdate
      Where ����id = ����id_In And ��ҳid Is Not Null And Nvl(����ʱ��, Sysdate) >= Sysdate;
    End If;
  End If;
  If ���_In = 1 Then
    --��Ϣ����
    Begin
      Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
        Using 1, n_�Һ�id;
    Exception
      When Others Then
        Null;
    End;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˹Һż�¼_����_Insert;
/

--127720:����,2018-06-29,ͬ������ҩƷ�շ���¼���ݸ�������
Create Or Replace Procedure Zl_ҩƷ�շ���¼_��������
(
  Id_In     In ҩƷ�շ���¼.Id%Type,
  ҩƷid_In In ҩƷ�շ���¼.ҩƷid%Type,
  ����_In   ҩƷ�շ���¼.����%Type := Null
) Is
  StrҩƷ     Varchar2(500);
  Lng�ⷿid   ҩƷ�շ���¼.�ⷿid%Type;
  Lngcur����  ҩƷ�շ���¼.����%Type;
  Lnglast���� ҩƷ�շ���¼.����%Type;
  Str����     ҩƷ�շ���¼.����%Type;
  StrЧ��     ҩƷ�շ���¼.Ч��%Type;
  Lng��Ӧ��id ҩƷ�շ���¼.��ҩ��λid%Type;
  Dat�������� ҩƷ�շ���¼.��������%Type;
  Str����     ҩƷ�շ���¼.����%Type;
  Dbl�������� ҩƷ�շ���¼.��д����%Type;
  Dblʵ������ ҩƷ�շ���¼.ʵ������%Type;
  Str��׼�ĺ� ҩƷ�շ���¼.��׼�ĺ�%Type;
  v_Error     Varchar2(255);
  Err_Custom Exception;
Begin
  If Nvl(����_In, 0) = 0 Then
    Return;
  End If;

  Select Nvl(a.����, 0), a.�ⷿid, Nvl(a.ʵ������, 0) * Nvl(a.����, 1), '[' || c.���� || ']' || c.����
  Into Lnglast����, Lng�ⷿid, Dblʵ������, StrҩƷ
  From ҩƷ�շ���¼ A, �շ���ĿĿ¼ C
  Where a.Id = Id_In And a.ҩƷid = c.Id;
  If Nvl(����_In, 0) = 0 Then
    Lngcur���� := Nvl(Lnglast����, 0);
  Else
    Lngcur���� := Nvl(����_In, 0);
  End If;

  Begin
    v_Error := '��һ�ʷ��������ҩƷ��ָ����������ʧЧ,������ɲ�����';
    --ȡ����ҩƷ������ 
    Select �ϴ�����, Ч��, Nvl(��������, 0) ��������, �ϴι�Ӧ��id, �ϴ���������, �ϴβ���, ��׼�ĺ�
    Into Str����, StrЧ��, Dbl��������, Lng��Ӧ��id, Dat��������, Str����, Str��׼�ĺ�
    From ҩƷ���
    Where �ⷿid = Lng�ⷿid And ҩƷid = ҩƷid_In And ���� = 1 And Nvl(����, 0) = Lngcur���� And
          (Nvl(����, 0) = 0 Or Ч�� Is Null Or Ч�� > Trunc(Sysdate));
  
  Exception
    When Others Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  End;

  If Lngcur���� <> Nvl(Lnglast����, 0) Then
    If Dbl�������� < Dblʵ������ And Lngcur���� <> 0 Then
      v_Error := StrҩƷ || '�Ŀ����������㣬������ֹ��';
      Raise Err_Custom;
    End If;
  End If;
  --����ҩƷ�շ���¼��������Ϣ 
  Update ҩƷ�շ���¼
  Set ���� = Lngcur����, ���� = Str����, Ч�� = StrЧ��, ��ҩ��λid = Lng��Ӧ��id, �������� = Dat��������, ���� = Str����, ��׼�ĺ� = Str��׼�ĺ�
  Where ID = Id_In;

  --����ԭ���ο��Ŀ������� 
  --���·�ҩ���ο��Ŀ������� 
  If Lnglast���� <> Lngcur���� Then
    Update ҩƷ���
    Set �������� = Nvl(��������, 0) + Dblʵ������
    Where �ⷿid + 0 = Lng�ⷿid And ҩƷid = ҩƷid_In And ���� = 1 And Nvl(����, 0) = Lnglast����;
  
    --�쳣���ݴ���
    Zl_ҩƷ���_���������쳣����(Lng�ⷿid, ҩƷid_In, Lnglast����);
  
    Update ҩƷ���
    Set �������� = Nvl(��������, 0) - Dblʵ������
    Where �ⷿid + 0 = Lng�ⷿid And ҩƷid = ҩƷid_In And ���� = 1 And Nvl(����, 0) = Lngcur����;
  End If;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�շ���¼_��������;
/

--127655:������,2018-06-22,������Commit������
CREATE OR REPLACE Procedure Zl_���ﴩ��̨_Liquid
(
  ����id_In In �ŶӼ�¼.����id%Type,
  ����id_In In �ŶӼ�¼.����id%Type
) Is

  --���ܣ����ӵ���Ϊ���˵��ŶӼ�¼����ȴ����̵�һ������̨

  n_Sn  ���ﴩ��̨.���%Type;
  v_Tmp ����ҽ����¼.����%Type;
  d_Tmp Date;

  Err_Item Exception;
  v_Err_Msg varchar2(200);
Begin

  -- ���24Сʱ���ڵ����ﴩ��̨
  Update ���ﴩ��̨
  Set ��������id = Null
  Where ID In
        (Select a.Id
         From ���ﴩ��̨ A, �ŶӼ�¼ B
         Where a.����id = b.����id And a.��������id = b.����id And a.��Ч = 1 And a.����id = ����id_In And b.���� < Sysdate - 1);

  -- ���24Сʱ���ڣ�����״̬��Ϊ��1-����Һ��5-�����̡������ﴩ��̨
  Update ���ﴩ��̨
  Set ��������id = Null
  Where ID In (Select a.Id
               From ���ﴩ��̨ A, �ŶӼ�¼ B
               Where a.����id = b.����id And a.��������id = b.����id And a.��Ч = 1 And a.����id = ����id_In And Not b.״̬ In (1, 5) And
                     b.���� < Sysdate - 1);
    

  -- Ϊ���ˡ��ŶӼ�¼�����䴩��̨
  Begin
    -- �����ŶӼ�¼���������ƣ�
    Select ����
    Into d_Tmp
    From �ŶӼ�¼
    Where ����id = ����id_In And ����id = ����id_In And ״̬ = 1
    For Update Nowait;

    -- ���ҿ�����ͬ����������̨δ������ŶӼ�¼
    Begin
      Select ���
      Into n_Sn
      From ���ﴩ��̨
      Where ����id = ����id_In And
            Not ��� In (Select ����̨ From �ŶӼ�¼ Where ״̬ = 1 And ����id = ����id_In And ����̨ > 0) And
            (��������id Is Null Or ��������id = 0) And ��Ч = 1 And Rownum < 2;
    Exception
      When Others Then
        n_Sn := Null;
    End;

    If n_Sn Is Not Null Then
      -- �ҵ������
      Update �ŶӼ�¼ Set ����̨ = n_Sn, ���� = Sysdate Where ����id = ����id_In And ����id = ����id_In And ״̬ = 1;
        
    Else
      -- δ�ҵ�����ƽ������һ������̨�����˵��ŶӼ�¼
      Begin
        Select ����̨
        Into n_Sn
        From (Select a.����̨, Count(1) ����
               From �ŶӼ�¼ A, ���ﴩ��̨ B
               Where a.����id = b.����id And a.����̨ = b.��� And a.����id = ����id_In And a.���� Between Sysdate - 1 And Sysdate And
                     a.״̬ = 1 And b.��Ч = 1
               Group By ����̨
               Order By ����, ����̨) A
        Where Rownum < 2;
      Exception
        When Others Then
          n_Sn := Null;
      End;

      If n_Sn Is Not Null Then
        -- �ҵ������
        Update �ŶӼ�¼ Set ����̨ = n_Sn, ���� = Sysdate Where ����id = ����id_In And ����id = ����id_In And ״̬ = 1;
          
      Else
        -- δ�ҵ�������һ����С�ŵĴ���̨
        Begin
          Select Min(���) Into n_Sn From ���ﴩ��̨ Where ����id = ����id_In And ��Ч = 1;
        Exception
          When Others Then
            n_Sn := Null;
        End;
        If n_Sn Is Not Null Then
          Update �ŶӼ�¼
          Set ����̨ = n_Sn, ���� = Sysdate
          Where ����id = ����id_In And ����id = ����id_In And ״̬ = 1;
            
        Else
            
          v_Err_Msg := '��ǰ����δ���ô���̨����Ч�Ĵ���̨��';
          Raise Err_Item;
        End If;
      End If;

    End If;

  Exception
    When Err_Item Then
      Raise Err_Item;
    When Others Then
      Begin
        Select ���� Into v_Tmp From ������Ϣ Where ����id = ����id_In;
      Exception
        When Others Then
          v_Tmp := 'δ֪';
      End;
      v_Err_Msg := '[' || v_Tmp || ']���ڴ���Һ�����У�';
      Raise Err_Item;
  End;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���ﴩ��̨_Liquid;
/

--127651:���˺�,2018-06-22,ɾ�������к���Commit��䣬������������. 
Create Or Replace Procedure Zl1_Autocptall(ǿ�Ƽ���_In In Number := 0) As
  Modilast Number(1); --�Ƿ����������Զ��ƷѲ���
  Period   Varchar2(6); --��Ҫ�������С�ڼ�
  Cursor Patitab Is
    Select Distinct ����id, ��ҳid
    From ��Ժ�����Զ�����
    Where Trunc(��ֹ����) >= (Select Min(��ʼ����) From �ڼ�� Where �ڼ� >= Period);
Begin
  If f_Is_Primary_Node = 0 Then
    Return;
  End If;
  Begin
    Select �ڼ� Into Period From �ڼ�� Where Trunc(Sysdate) - 1 Between Trunc(��ʼ����) And Trunc(��ֹ����);
  Exception
    When Others Then
      Return;
  End;
  Select zl_GetSysParameter(7) Into Modilast From Dual;

  If Modilast = 1 Then
    Period := To_Char(Add_Months(To_Date(Period || '05', 'yyyymmdd'), -1), 'yyyymm');
  End If;

  For Patifld In Patitab Loop
    If Patifld.����id Is Not Null And Patifld.��ҳid Is Not Null Then
      Zl1_Autocptone(Patifld.����id, Patifld.��ҳid, Period, 1, ǿ�Ƽ���_In);
    End If;
  End Loop;
End Zl1_Autocptall;
/

--127651:���˺�,2018-06-22,ɾ�������к���Commit��䣬������������. 
Create Or Replace Procedure Zl1_Autocptpati
(
  Patiid      In Number,
  Pageid      In Number,
  Recalcbdate In ���˱䶯��¼.�ϴμ���ʱ��%Type := Null,
  ǿ�Ƽ���_In In Number := 0
) As
  Modilast Number(1); --�Ƿ����������Զ��ƷѲ���
  Period   Varchar2(6); --��Ҫ�������С�ڼ�
Begin
  Begin
    Select �ڼ� Into Period From �ڼ�� Where Trunc(Sysdate) Between Trunc(��ʼ����) And Trunc(��ֹ����);
  Exception
    When Others Then
      Return;
  End;

  Select Zl_To_Number(zl_GetSysParameter(7)) Into Modilast From Dual;

  If Modilast = 1 Then
    Period := To_Char(Add_Months(To_Date(Period || '05', 'yyyymmdd'), -1), 'yyyymm');
  End If;

  If Recalcbdate Is Not Null Then
    Update ���˱䶯��¼
    Set �ϴμ���ʱ�� = Null
    Where ����id = Patiid And ��ҳid = Pageid And �ϴμ���ʱ�� >= Recalcbdate;
  End If;

  Zl1_Autocptone(Patiid, Pageid, Period, 0, ǿ�Ƽ���_In);
End Zl1_Autocptpati;
/

--127651:���˺�,2018-06-22,ɾ�������к���Commit��䣬������������. 
Create Or Replace Procedure Zl1_Autocptward
(
  Wardid      In Number,
  Recalcbdate In ���˱䶯��¼.�ϴμ���ʱ��%Type := Null,
  ǿ�Ƽ���_In In Number := 0
) As
  Modilast Number(1); --�Ƿ����������Զ��ƷѲ���
  Period   Varchar2(6); --��Ҫ�������С�ڼ�

  Cursor Patitab Is
    Select Distinct ����id, ��ҳid
    From ��Ժ�����Զ�����
    Where ����id = Wardid And Trunc(��ֹ����) >= (Select Min(��ʼ����) From �ڼ�� Where �ڼ� >= Period);
Begin
  Begin
    Select �ڼ� Into Period From �ڼ�� Where Trunc(Sysdate) - 1 Between Trunc(��ʼ����) And Trunc(��ֹ����);
  Exception
    When Others Then
      Return;
  End;
  Select zl_GetSysParameter(7) Into Modilast From Dual;

  If Modilast = 1 Then
    Period := To_Char(Add_Months(To_Date(Period || '05', 'yyyymmdd'), -1), 'yyyymm');
  End If;

  If Recalcbdate Is Not Null Then
    Update ���˱䶯��¼
    Set �ϴμ���ʱ�� = Null
    Where (����id, ��ҳid) In (Select ����id, ��ҳid From ������ҳ Where ��ǰ����id = Wardid And ��Ժ���� Is Null) And
          �ϴμ���ʱ�� >= Recalcbdate;
  End If;

  For Patifld In Patitab Loop
    If Patifld.����id Is Not Null And Patifld.��ҳid Is Not Null Then
      Zl1_Autocptone(Patifld.����id, Patifld.��ҳid, Period, 1, ǿ�Ƽ���_In);
    End If;
  End Loop;
End Zl1_Autocptward;
/

--127651:���˺�,2018-06-22,ɾ�������к���Commit��䣬������������. 
 Create Or Replace Procedure Zl_Third_Swapstaut
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --����:��ȡ����֧����״̬
  --���:Xml_In:
  --<IN>
  --        <JYLB>�������</JYLB> //֧������΢�ŵ�
  --        <JYLSH>������ˮ��</JYLSH>
  --        <BRID>����ID</BRID>
  --</IN>
  --����:Xml_Out
  --  <OUT>
  --    <ZT>״̬</ZT>  �D�D0-����ʧ��,1-���׳ɹ�;2-�������ڽ�����;3-�����ڸý��׼�¼
  --    <JYSJ>����ʱ��</JYSJ>  �D�D��״̬Ϊ1ʱ�ŷ���,����Ϊ��  ��ʽΪ'YYYY-MM-DD hh24:mi:ss'
  --    <JYID>ҵ����ID</JYID>  �D�D��״̬Ϊ1ʱ�ŷ���,����Ϊ��   ��ԹҺźͽ���Ϊ����ID;���Ԥ��ΪԤ��ID;����շ�Ϊ�������
  --    <YWLX>ҵ������</YWLX>   �D�D null-��ʷ����;1-Ԥ��;2-����;3-�շ�;4-�Һ�
  --    <DJH>���ݺ�</DJH>  �D�D ����ö��ŷָ�  ��ԹҺ�Ϊ�Һŵ��ݺ�,��Խ���Ϊ���ʵ��ݺ�,���Ԥ��ΪԤ�����ݺ�,����շ�Ϊ�շѵ��ݺ�


  --  </OUT>
  --------------------------------------------------------------------------------------------------
  v_Temp    Varchar2(32767); --��ʱXML
  x_Templet Xmltype; --ģ��XML

  v_�������   �������׼�¼.���%Type;
  v_������ˮ�� �������׼�¼.��ˮ��%Type;
  n_����id     ����Ԥ����¼.����id%Type;

  n_Count    Number(18);
  n_����id   �������׼�¼.ҵ�����id%Type;
  n_ҵ������ �������׼�¼.ҵ������%Type;
  v_Nos      Varchar2(3000);
  d_����ʱ�� Date;

Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Nvl(Extractvalue(Value(A), 'IN/JYLB'), '-'), Nvl(Extractvalue(Value(A), 'IN/JYLSH'), '-'),
         Nvl(Extractvalue(Value(A), 'IN/BRID'), 0)
  Into v_�������, v_������ˮ��, n_����id
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  Begin
    Select ״̬, ����ʱ��, ҵ�����id, ҵ������
    Into n_Count, d_����ʱ��, n_����id, n_ҵ������
    From �������׼�¼
    Where ��� = v_������� And ��ˮ�� = v_������ˮ��
    For Update Nowait;
  Exception
    When Others Then
      n_Count := -1;
  End;

  If n_Count = -1 Then
    Select Count(1) Into n_Count From �������׼�¼ Where ��� = v_������� And ��ˮ�� = v_������ˮ��;
    If n_Count = 0 Then
      Select Count(1), Max(�տ�ʱ��), Max(Decode(��¼����, 1, ID, 3, �������, ����id)), Max(Mod(��¼����, 10))
      Into n_Count, d_����ʱ��, n_����id, n_ҵ������
      From ����Ԥ����¼
      Where ��¼���� <> 11 And Nvl(У�Ա�־, 0) <> 1 And ����id = n_����id And ������ˮ�� = v_������ˮ�� And
            �����id In (Select ID From ҽ�ƿ���� Where ���� = v_�������);
      If n_Count = 0 Then
        n_Count := 3;
      Else
        --�������������׼�¼�������ڲ���Ԥ����¼��Ҳ��ʾ���׳ɹ�
        n_Count := 1;
      End If;
    Else
      n_Count := 2;
    End If;
  End If;

  If n_Count = 1 Then
    If n_ҵ������ = 1 Then
      Select Max(NO) Into v_Nos From ����Ԥ����¼ Where ��¼���� = 1 And ID = n_����id;
    Elsif n_ҵ������ = 2 Then
      Select Max(NO) Into v_Nos From ���˽��ʼ�¼ Where ID = n_����id;
    Elsif n_ҵ������ = 3 Then
      Select f_List2str(Cast(Collect(NO) As t_Strlist))
      Into v_Nos
      From (Select Distinct a.No As NO
             From ������ü�¼ A, ����Ԥ����¼ B
             Where a.����id = b.����id And b.������� = n_����id);
    Elsif n_ҵ������ = 4 Then
      Select Max(NO) Into v_Nos From ������ü�¼ Where ��¼���� = 4 And ����id = n_����id;
    End If;
  End If;

  v_Temp := '<ZT>' || n_Count || '</ZT>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<JYSJ>' || To_Char(d_����ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '</JYSJ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<JYID>' || Nvl(n_����id, 0) || '</JYID>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<YWLX>' || n_ҵ������ || '</YWLX>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<DJH>' || Nvl(v_Nos, '') || '</DJH>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Xml_Out := x_Templet;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Swapstaut;
/

--127651:���˺�,2018-06-22,ɾ�������к���Commit��䣬������������. 
Create Or Replace Procedure Zl_���ﻮ�ۼ�¼_Clear(Day_In Number) As
  --���ܣ��Զ�������۵� 
  --������Day_IN=ɾ�����ۺ󳬹�Day_IN��δ�շѵĵ��� 
  Cursor c_Price Is
    Select Distinct a.No, f_List2str(Cast(Collect(To_Char(a.���)) As t_Strlist)) As ���
    From ������ü�¼ A, δ��ҩƷ��¼ B
    Where a.��¼���� = 1 And a.��¼״̬ = 0 And a.ִ��״̬ Not In (1, 2) And a.������ Is Not Null And a.����Ա���� Is Null And
          b.���� In (8, 24) And Nvl(b.���շ�, 0) = 0 And a.No = b.No And Nvl(a.ִ�в���id, 0) = Nvl(b.�ⷿid, 0) And
          Sysdate - b.�������� >= Day_In
    Group By a.No;
Begin
  For r_Price In c_Price Loop
    If Not r_Price.��� Is Null Then
      Zl_���ﻮ�ۼ�¼_Delete(r_Price.No, r_Price.���, 1);
    End If;
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���ﻮ�ۼ�¼_Clear;
/

--127620:Ϳ����,2018-06-22,commit��䴦��
CREATE OR REPLACE Procedure Zl_Ӱ����ͼ��_У��
( 
  ҽ��id_In   In Ӱ�����¼.ҽ��id%Type, 
  ͼ��uid_In  In Ӱ����ͼ��.ͼ��uid%Type, 
  У������_In In Ӱ�����¼.У������%Type, 
  У�Խ��_In In Ӱ����ͼ��.У�Խ��%Type 
) Is 
  n_У��״̬ Number(1); 
  n_Tag      Number(1); 
  n_Count    Number(4); 
  v_���uid  Ӱ�����¼.���uid%Type; 
Begin 
 
  Select Nvl(У��״̬, 0), ���uid Into n_У��״̬, v_���uid From Ӱ�����¼ Where ҽ��id = ҽ��id_In; 
  Update Ӱ����ͼ�� Set У�Խ�� = У�Խ��_In Where ͼ��uid = ͼ��uid_In; 
 
  If У�Խ��_In = 5 Or У�Խ��_In = 6 Then 
    n_Tag := 1; 
  Else 
    n_Tag := 2; 
  End If; 
 
  If n_У��״̬ = 0 Then 
    Update Ӱ�����¼ Set У��״̬ = n_Tag, У������ = У������_In Where ҽ��id = ҽ��id_In; 
  Elsif n_У��״̬ = 1 And n_Tag = 2 Then 
    Update Ӱ�����¼ Set У��״̬ = n_Tag Where ҽ��id = ҽ��id_In; 
  Elsif n_У��״̬ = 2 And (У�Խ��_In = 5 Or У�Խ��_In = 6) Then 
    Select Count(1) 
    Into n_Count 
    From Ӱ�������� b, Ӱ����ͼ�� c 
    Where b.����uid = c.����uid And b.���uid = v_���uid And (c.У�Խ�� > 0 And c.У�Խ�� < 5); 
 
    If n_Count = 0 Then 
      Update Ӱ�����¼ Set У��״̬ = 1 Where ҽ��id = ҽ��id_In; 
    End If; 
  End If; 
Exception 
  When Others Then 
    Zl_Errorcenter(Sqlcode, Sqlerrm); 
End Zl_Ӱ����ͼ��_У��;
/

--126802:���ϴ�,2018-06-21,���ӷѷ��ع̶��Ľ����Ϣ
Create Or Replace Procedure Zl_Third_Getregfeedetail
(
  Xml_In  In Xmltype,
  Xml_Out In Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --����:��ȡ�Һŷ�����ϸ
  --���:Xml_In:
  --<IN>
  --  <BRID></BRID> //����ID
  --  <XM></XM>     //����
  --  <SFZH></SFZH> //���֤��
  --  <SFYY></SFYY> //�Ƿ��ԤԼ��֧��,1-��ԤԼ��֧��,0-�Һ�,ԤԼ֧��,ԤԼ���գ�Ĭ��Ϊ0
  --  <GHDH></GHDH> //�Һŵ���,ԤԼ����ʱ����
  --  <GHHM></GHHM> //�ҺŰ��ź���,�Һź�ԤԼʱ����
  --  <XMID></XMID> //�ҺŰ��ŵ���ĿID,�Һź�ԤԼʱ����
  --  <FB></FB>     //���˷ѱ�
  --  <FKFS></FKFS> //���ʽ
  --  <RQ></RQ>     //����
  --  <ZD></ZD>     //վ��
  --</IN>
  --����:Xml_Out
  -- <OUTPUT>
  --  <ZJE></ZJE>   //��ʵ�ս��
  --  <XMMX>        //��Ŀ��ϸ
  --    <XM>
  --      <DJH></DJH>       //���ݺ�
  --      <MC></MC>   //��Ŀ����
  --      <ID></ID>   //��ĿID
  --      <SL></SL>   //����������*����
  --      <YSJE></YSJE>   //Ӧ�ս��
  --      <SSJE></SSJE>   //ʵ�ս��
  --      <SJFM></SJFM>       //�վݷ�Ŀ
  --    </XM>
  --    <XM>
  --    ...
  --    </XM>
  --  </XMMX>
  -- </OUTPUT>

  --------------------------------------------------------------------------------------------------
  x_Templet    Xmltype; --ģ��XML
  v_Err_Msg    Varchar2(200);
  v_Temp       Varchar2(4000);
  n_��Ŀid     �ҺŰ���.��Ŀid%Type;
  v_No         ������ü�¼.No%Type;
  n_ԤԼ       Number(3);
  n_����id     ������Ϣ.����id%Type;
  v_����       ������Ϣ.����%Type;
  v_�Ա�       ������Ϣ.�Ա�%Type;
  v_����       ������Ϣ.����%Type;
  v_���֤��   ������Ϣ.���֤��%Type;
  v_�ѱ�       ������Ϣ.�ѱ�%Type;
  v_���ʽ   ҽ�Ƹ��ʽ.����%Type;
  v_��ʽ       Varchar2(20);
  d_����       Date;
  v_վ��       ���ű�.վ��%Type;
  v_����       �ҺŰ���.����%Type;
  n_�ܽ��     ������ü�¼.ʵ�ս��%Type;
  n_ʵ�ս��   ������ü�¼.ʵ�ս��%Type;
  v_ʵ��       Varchar(500);
  v_������Ŀid Varchar2(500);
  v_��������   Varchar2(500);
  v_����ֵ     Varchar2(100);
  n_cursor     Number(3);
  Err_Item     Exception;
  
  TYPE Price_type IS RECORD(��ĿID ������ü�¼.�շ�ϸĿID%Type,
                              ���� ������ü�¼.����%TYPE, 
                              ���� ������ü�¼.��׼����%TYPE, 
                              Ӧ�� ������ü�¼.Ӧ�ս��%TYPE, 
                              ʵ�� ������ü�¼.ʵ�ս��%TYPE);--����Price��¼���� 
  TYPE Price_type_array IS TABLE OF Price_type INDEX BY BINARY_INTEGER;--������Price��¼���������� 
  Price_rec Price_type;--�������������ͣ�Price��¼����
  Price_rec_array Price_type_array;--�������������ͣ����Price��¼����������
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');
  Select Extractvalue(Value(A), 'IN/BRID'), Extractvalue(Value(A), 'IN/XMID'), Extractvalue(Value(A), 'IN/GHHM'),
         Extractvalue(Value(A), 'IN/FB'), Extractvalue(Value(A), 'IN/ZD'), Extractvalue(Value(A), 'IN/SFYY'),
         Extractvalue(Value(A), 'IN/GHDH'), Extractvalue(Value(A), 'IN/SFZH'), Extractvalue(Value(A), 'IN/XM'),
         To_Date(Extractvalue(Value(A), 'IN/RQ'), 'yyyy-mm-dd hh24:mi:ss'), Extractvalue(Value(A), 'IN/FKFS')
  Into n_����id, n_��Ŀid, v_����, v_�ѱ�, v_վ��, n_ԤԼ, v_No, v_���֤��, v_����, d_����, v_��ʽ
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If v_��ʽ Is Null Then
    Select Max(����) Into v_���ʽ From ҽ�Ƹ��ʽ Where ȱʡ��־ = 1;
  Else
    Select Max(����) Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = v_��ʽ;
    If v_���ʽ Is Null Then
      v_���ʽ := v_��ʽ;
    End If;
  End If;
  If Nvl(n_����id, 0) = 0 And Not v_���֤�� Is Null And Not v_���� Is Null Then
    n_����id := Zl_Third_Getpatiid(v_���֤��, v_����);
  End If;
  If Nvl(n_����id, 0) = 0 Then
    v_Err_Msg := '�޷�ȷ��������Ϣ,����';
    Raise Err_Item;
  End If;
  Select Max(�Ա�), Max(����) Into v_�Ա�, v_���� From ������Ϣ Where ����ID = n_����id;
  
  n_�ܽ�� := 0;
  If v_No Is Null Then
    --�ҺŻ���ԤԼ
    For c_�Һ���Ŀ In (Select 1 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, 1 As ����,
                          c.Id As ������Ŀid, c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, -1 As ִ�п�������
                   From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C
                   Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = n_��Ŀid And d_���� Between b.ִ������ And
                         Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))
                   Union All
                   Select 2 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, d.�������� As ����,
                          c.Id As ������Ŀid, c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, -1 As ִ�п�������
                   From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C, �շѴ�����Ŀ D
                   Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = d.����id And d.����id = n_��Ŀid And
                         d_���� Between b.ִ������ And Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))) Loop
      v_ʵ��     := Zl_Actualmoney(v_�ѱ�, c_�Һ���Ŀ.��Ŀid, c_�Һ���Ŀ.������Ŀid, c_�Һ���Ŀ.���� * c_�Һ���Ŀ.����);
      n_ʵ�ս�� := To_Number(Substr(v_ʵ��, Instr(v_ʵ��, ':') + 1));
      n_�ܽ��   := n_�ܽ�� + Nvl(n_ʵ�ս��, 0);
      v_Temp     := v_Temp || '<XM><DJH></DJH><MC>' || c_�Һ���Ŀ.��Ŀ���� || '</MC>' || '<ID>' || c_�Һ���Ŀ.��Ŀid || '</ID>' ||
                    '<SL>' || c_�Һ���Ŀ.���� || '</SL>' || '<YSJE>' || c_�Һ���Ŀ.���� * c_�Һ���Ŀ.���� || '</YSJE>' || '<SSJE>' ||
                    n_ʵ�ս�� || '</SSJE>' || '<SJFM>' || c_�Һ���Ŀ.�վݷ�Ŀ || '</SJFM></XM>';
    End Loop;
  Else
    --ԤԼ����
    For c_�Һ���Ŀ In (Select a.�շ�ϸĿid As ��Ŀid, a.Ӧ�ս��, a.ʵ�ս��, a.���㵥λ, a.�վݷ�Ŀ, b.���� As ��Ŀ����, a.No, Nvl(a.����, 1) As ����, a.����
                   From ������ü�¼ A, �շ���ĿĿ¼ B
                   Where a.�շ�ϸĿid = b.Id And a.No = v_No And a.��¼���� = 4 And a.��¼״̬ = 0) Loop
      n_�ܽ�� := n_�ܽ�� + Nvl(c_�Һ���Ŀ.ʵ�ս��, 0);
      v_����   := c_�Һ���Ŀ.���㵥λ;
      v_Temp   := v_Temp || '<XM><DJH>' || c_�Һ���Ŀ.No || '</DJH><MC>' || c_�Һ���Ŀ.��Ŀ���� || '</MC>' || '<ID>' || c_�Һ���Ŀ.��Ŀid ||
                  '</ID>' || '<SL>' || c_�Һ���Ŀ.���� * c_�Һ���Ŀ.���� || '</SL>' || '<YSJE>' || c_�Һ���Ŀ.Ӧ�ս�� || '</YSJE>' ||
                  '<SSJE>' || c_�Һ���Ŀ.ʵ�ս�� || '</SSJE>' || '<SJFM>' || c_�Һ���Ŀ.�վݷ�Ŀ || '</SJFM></XM>';
    End Loop;
  End If;

  If Nvl(n_ԤԼ, 0) = 0 Then
    Begin
      Select Zl_Fun_Customregexpenses(n_����id, 0, v_����, v_����, v_�Ա�, v_����, v_���֤��, v_�ѱ�, v_���ʽ) Into v_������Ŀid From Dual;
    Exception
      When Others Then
        v_������Ŀid := Null;
    End;
    If v_������Ŀid Is Not Null Then
      IF Instr(v_������Ŀid, '|') > 0 Then
        v_�������� := v_������Ŀid || ','; --�Կո�ֿ���|��β,û�н�������
        v_������Ŀid := '';
        n_cursor   := 0;
        While v_�������� Is Not Null Loop
          v_����ֵ := Substr(v_��������, 1, Instr(v_��������, ',') - 1);
          Price_rec.��ĿID := To_Number(Substr(v_����ֵ, 1, Instr(v_����ֵ, '|') - 1));
        
          v_����ֵ := Substr(v_����ֵ, Instr(v_����ֵ, '|') + 1);
          Price_rec.���� := To_Number(Substr(v_����ֵ, 1, Instr(v_����ֵ, '|') - 1));
        
          v_����ֵ := Substr(v_����ֵ, Instr(v_����ֵ, '|') + 1);
          Price_rec.���� := To_Number(Substr(v_����ֵ, 1, Instr(v_����ֵ, '|') - 1));
          
          v_����ֵ := Substr(v_����ֵ, Instr(v_����ֵ, '|') + 1);
          Price_rec.Ӧ�� := To_Number(Substr(v_����ֵ, 1, Instr(v_����ֵ, '|') - 1));
        
          v_����ֵ := Substr(v_����ֵ, Instr(v_����ֵ, '|') + 1);
          Price_rec.ʵ�� := To_Number(v_����ֵ);
          n_cursor := n_cursor + 1;
          Price_rec_array(n_cursor):=Price_rec;
          v_�������� := Substr(v_��������, Instr(v_��������, ',') + 1);
          v_������Ŀid := v_������Ŀid || ',' || Price_rec_array(n_cursor).��ĿID;
        End Loop;
        
        If v_������Ŀid is not null then
          v_������Ŀid := substr(v_������Ŀid, 2);
        End if;
        
        For c_������Ŀ In (Select /*+cardinality(D,10)*/
                        5 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, 1 As ����, c.Id As ������Ŀid,
                        c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, -1 As ִ�п�������
                       From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C, Table(f_Str2list(v_������Ŀid)) D
                       Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = d.Column_Value And d_���� Between b.ִ������ And
                             Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')))  Loop
          FOR n_cursor IN 1..Price_rec_array.count LOOP        
            IF c_������Ŀ.��Ŀid = Price_rec_array(n_cursor).��ĿID Then
              n_ʵ�ս�� := Price_rec_array(n_cursor).ʵ��;
              n_�ܽ��   := n_�ܽ�� + Nvl(n_ʵ�ս��, 0);
              v_Temp     := v_Temp || '<XM><DJH></DJH><MC>' || c_������Ŀ.��Ŀ���� || '</MC>' || '<ID>' || c_������Ŀ.��Ŀid || '</ID>' ||
                          '<SL>' || Price_rec_array(n_cursor).���� || '</SL>' || '<YSJE>' || Price_rec_array(n_cursor).Ӧ�� || '</YSJE>' || '<SSJE>' ||
                          Price_rec_array(n_cursor).ʵ�� || '</SSJE>' || '<SJFM>' || c_������Ŀ.�վݷ�Ŀ || '</SJFM></XM>';
              EXIT;
            End IF;
          End LOOP;
        End Loop;  
      Else
        For c_������Ŀ In (Select /*+cardinality(D,10)*/
                        5 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, 1 As ����, c.Id As ������Ŀid,
                        c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, -1 As ִ�п�������
                       From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C, Table(f_Str2list(v_������Ŀid)) D
                       Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = d.Column_Value And d_���� Between b.ִ������ And
                             Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))
                       Union All
                       Select /*+cardinality(E,10)*/
                        5 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, d.�������� As ����,
                        c.Id As ������Ŀid, c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, -1 As ִ�п�������
                       From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C, �շѴ�����Ŀ D, Table(f_Str2list(v_������Ŀid)) E
                       Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = d.����id And d.����id = e.Column_Value And
                             d_���� Between b.ִ������ And Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))) Loop
          v_ʵ��     := Zl_Actualmoney(v_�ѱ�, c_������Ŀ.��Ŀid, c_������Ŀ.������Ŀid, c_������Ŀ.���� * c_������Ŀ.����);
          n_ʵ�ս�� := To_Number(Substr(v_ʵ��, Instr(v_ʵ��, ':') + 1));
          n_�ܽ��   := n_�ܽ�� + Nvl(n_ʵ�ս��, 0);
          v_Temp     := v_Temp || '<XM><DJH></DJH><MC>' || c_������Ŀ.��Ŀ���� || '</MC>' || '<ID>' || c_������Ŀ.��Ŀid || '</ID>' ||
                        '<SL>' || c_������Ŀ.���� || '</SL>' || '<YSJE>' || c_������Ŀ.���� * c_������Ŀ.���� || '</YSJE>' || '<SSJE>' ||
                        n_ʵ�ս�� || '</SSJE>' || '<SJFM>' || c_������Ŀ.�վݷ�Ŀ || '</SJFM></XM>';
        End Loop;
      End IF;
    End If;
  End If;

  v_Temp := '<XMMX>' || v_Temp || '</XMMX>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;

  v_Temp := '<ZJE>' || n_�ܽ�� || '</ZJE>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;

  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getregfeedetail;
/

--126802:���ϴ�,2018-06-21,���ӷѷ��ع̶��Ľ����Ϣ
Create Or Replace Procedure Zl_Third_Payment
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is

  -------------------------------------------------------------------------------------------------- 
  --����:�����ӿ�֧�� 
  --���:Xml_In: 
  --<IN>
  --        <NO></NO>                       //�շѵ��ݺŴ�,���ŷָ�������ݺ�
  --        <JE></JE>                       //�ܽ��
  --        <BRID>����ID</BRID>             //����ID
  --        <XM>����</XM>                   //����
  --        <SFZH>���֤��</SFZH>           //���֤��
  --        <SFGH></SFGH>                   //�Ƿ�Һŵ�
  --        <WCJE>����</WCJE>             //������ʱ,���ܽ��-���ν�������ܶ�Ϊ׼
  --       <JSLIST>
  --         <JS>
  --              <JSKLB>֧�������</JSKLB >
  --              <JSKH>֧������</ JSKH >
  --              <JSFS>֧����ʽ</JSFS> //֧����ʽ:�ֽ�;֧Ʊ,�����������,���Դ���
  --              <JSJE>֧�����</JSJE>
  --              <JYLSH>������ˮ��</JYLSH>
  --              <ZY>ժҪ</ZY>
  --              <SFCYJ>�Ƿ��Ԥ��</SFCYJ>  //�ʳ�Ԥ��ʱ,ֻ��JSJE�ڵ�:1-��Ԥ��
  --              <SFXFK>�Ƿ����ѿ�</SFXFK>  //(1-�����ѿ�),���ѿ�ʱ,������㿨���,���㿨��,������Ƚӵ�
  --              <EXPENDLIST>  //��չ������Ϣ
  --                  <EXPEND>
  --                        <JYMC >��������</��������>
  --                        <JYLR>��������</JYLR>
  --                  </EXPEND>
  --              </EXPENDLIST>
  --         </JS>
  --       </JSLIST >
  --</IN>

  --����:Xml_Out 
  --  <OUT> 
  --    <CZSJ>����ʱ��</CZSJ>          //HIS�ĵǼ�ʱ��
  --    <JZID>����ID</JZID>          //���ν���ID
  --    �D�D�������д�������˵����ȷִ�� 
  --    <ERROR> 
  --      <MSG>������Ϣ</MSG> 
  --    </ERROR> 
  --  </OUT> 
  -------------------------------------------------------------------------------------------------- 
  v_Nos      Varchar2(4000);
  n_�շ��ܶ� ������ü�¼.ʵ�ս��%Type;

  n_�����id ҽ�ƿ����.Id%Type;
  v_���㷽ʽ Varchar2(2000);
  n_����id   ������ü�¼.����id%Type;
  v_���֤�� ������Ϣ.���֤��%Type;
  v_����     ������ü�¼.����%Type;
  v_�Ա�     ������ü�¼.�Ա�%Type;
  v_����     ������ü�¼.����%Type;

  v_ҽ�Ƹ��ʽ���� ҽ�Ƹ��ʽ.����%Type;
  v_���ʽ         ҽ�Ƹ��ʽ.����%Type;
  v_����Ա����       ������ü�¼.����Ա���%Type;
  v_����Ա����       ������ü�¼.����Ա����%Type;
  n_����id           ������ü�¼.����id%Type;
  n_���ʽ��         ������ü�¼.���ʽ��%Type;
  d_�շ�ʱ��         ����Ԥ����¼.�տ�ʱ��%Type;
  n_���ѿ�id         ���ѿ�Ŀ¼.Id%Type;
  v_�շѽ���         Varchar2(2000);
  v_��ͨ����         Varchar2(4000);
  n_�Ƿ�Һ�         Number(3);
  n_Ԥ��֧��         ������ü�¼.ʵ�ս��%Type;
  n_��֧ͨ��         ������ü�¼.ʵ�ս��%Type;
  v_���㿨��         ����Ԥ����¼.����%Type;
  n_���㿨���       ����Ԥ����¼.���㿨���%Type;
  v_������ˮ��       ����Ԥ����¼.������ˮ��%Type;
  v_����˵��         ����Ԥ����¼.����˵��%Type;
  v_ժҪ             ����Ԥ����¼.ժҪ%Type;
  n_����id           �ҺŰ���.����id%Type;
  n_��Ŀid           �ҺŰ���.��Ŀid%Type;
  n_ҽ��id           �ҺŰ���.ҽ��id%Type;
  v_ҽ������         �ҺŰ���.ҽ������%Type;
  v_����             �ҺŰ���.����%Type;
  n_�����           ������Ϣ.�����%Type;
  d_����ʱ��         ���˹Һż�¼.����ʱ��%Type;
  v_�ѱ�             ������Ϣ.�ѱ�%Type;
  n_����             ���˹Һż�¼.����%Type;
  n_���ɶ���         Number(3);
  v_Para             Varchar2(500);
  n_�Һ�ģʽ         Number(3);
  d_����ʱ��         Date;
  v_��ʱ���㷽ʽ     ����Ԥ����¼.���㷽ʽ%Type;
  n_�����¼id       �ٴ������¼.Id%Type;
  n_���             ������ü�¼.���%Type;
  v_������Ŀid       Varchar2(500);
  v_��������         Varchar2(500);
  v_����ֵ           Varchar2(100);
  n_cursor           Number(3);
  n_ʵ�ս��         ������ü�¼.ʵ�ս��%Type;
  v_ʵ��             Varchar2(500);
  n_��������         ������ü�¼.��������%Type;
  n_���˿���id       ������ü�¼.���˿���id%Type;
  n_ִ�в���id       ������ü�¼.ִ�в���id%Type;
  v_No               ������ü�¼.No%Type;
  n_ҽ��֧��         ����Ԥ����¼.��Ԥ��%Type;
  n_Exists           Number;
  v_�����           �������׼�¼.���%Type;
  n_ҵ������         �������׼�¼.ҵ������%Type;
  n_�������         ����Ԥ����¼.�������%Type;
  v_Temp             Varchar2(32767); --��ʱXML 
  x_Templet          Xmltype; --ģ��XML 

  v_Err_Msg Varchar2(200);
  Err_Item Exception;
  Err_Special Exception;
  n_Count    Number(18);
  v_����Ա   ������ü�¼.����Ա����%Type;
  v_��ҩ���� Varchar2(4000);
  n_����   ����Ԥ����¼.��Ԥ��%Type;
  
  TYPE Price_type IS RECORD(��ĿID ������ü�¼.�շ�ϸĿID%Type,
                              ���� ������ü�¼.����%TYPE, 
                              ���� ������ü�¼.��׼����%TYPE, 
                              Ӧ�� ������ü�¼.Ӧ�ս��%TYPE, 
                              ʵ�� ������ü�¼.ʵ�ս��%TYPE);--����Price��¼���� 
  TYPE Price_type_array IS TABLE OF Price_type INDEX BY BINARY_INTEGER;--������Price��¼���������� 
  Price_rec Price_type;--�������������ͣ�Price��¼����
  Price_rec_array Price_type_array;--�������������ͣ����Price��¼����������

  Function Zl_��������(��¼id_In �ٴ������¼.Id%Type) Return Varchar2 As
    n_���﷽ʽ �ٴ������¼.���﷽ʽ%Type;
    v_����     ���˹Һż�¼.����%Type;
    v_Rowid    Varchar2(500);
    n_Next     Integer;
    n_First    Integer;
  Begin
    Begin
      Select Nvl(���﷽ʽ, 0) Into n_���﷽ʽ From �ٴ������¼ Where ID = ��¼id_In;
    Exception
      When Others Then
        v_Err_Msg := '�����¼(' || ��¼id_In || ')δ�ҵ�!';
        Raise Err_Item;
    End;
  
    --0-�����1-ָ�����ҡ�2-��̬���3-ƽ������,��Ӧ������������
    v_���� := Null;
    If n_���﷽ʽ = 1 Then
      --1-ָ������
      Begin
        Select b.���� Into v_���� From �ٴ��������Ҽ�¼ A, �������� B Where a.����id = b.Id And a.��¼id = ��¼id_In;
      Exception
        When Others Then
          v_���� := Null;
      End;
    End If;
    If n_���﷽ʽ = 2 Then
      --2-��̬����:�ø��ű���Һ�δ�������ٵ�����
      For c_���� In (Select ��������, Sum(Num) As Num
                   From (Select b.���� As ��������, 0 As Num
                          From �ٴ��������Ҽ�¼ A, �������� B
                          Where a.����id = b.Id And a.��¼id = ��¼id_In
                          Union All
                          Select ����, Count(����) As Num
                          From ���˹Һż�¼
                          Where Nvl(ִ��״̬, 0) = 0 And ����ʱ�� Between Trunc(Sysdate) And Sysdate And �����¼id = ��¼id_In And
                                ���� In (Select d.����
                                       From �ٴ��������Ҽ�¼ C, �������� D
                                       Where c.����id = d.Id And c.��¼id = ��¼id_In)
                          Group By ����)
                   Group By ��������
                   Order By Num) Loop
        v_���� := c_����.��������;
        Exit;
      End Loop;
    End If;
    If n_���﷽ʽ = 3 Then
      --ƽ�������ǰ����=1��ʾ�´�Ӧȡ�ĵ�ǰ����
      n_Next  := 0;
      n_First := 1;
      For c_���� In (Select a.Rowid As Rid, b.���� As ��������, a.��ǰ����
                   From �ٴ��������Ҽ�¼ A, �������� B
                   Where a.����id = b.Id And a.��¼id = ��¼id_In) Loop
        If n_First = 1 Then
          v_Rowid := c_����.Rid;
        End If;
        If n_Next = 1 Then
          v_���� := c_����.��������;
          Update �ٴ��������Ҽ�¼ Set ��ǰ���� = 1 Where Rowid = c_����.Rid;
          Exit;
        End If;
        If Nvl(c_����.��ǰ����, 0) = 1 Then
          Update �ٴ��������Ҽ�¼ Set ��ǰ���� = 0 Where Rowid = c_����.Rid;
          n_Next := 1;
        End If;
      End Loop;
      If v_���� Is Null Then
        Update �ٴ��������Ҽ�¼ Set ��ǰ���� = 1 Where Rowid = v_Rowid Returning ����id Into v_����;
        Select ���� Into v_���� From �������� Where ID = v_����;
      End If;
    End If;
    Return v_����;
  End;

  Function Zl_����(����_In �ҺŰ���.����%Type) Return Varchar2 As
    n_���﷽ʽ �ҺŰ���.���﷽ʽ%Type;
    n_����id   �ҺŰ���.Id%Type;
    v_����     ���˹Һż�¼.����%Type;
    v_Rowid    Varchar2(500);
    n_Next     Integer;
    n_First    Integer;
  Begin
  
    Begin
      Select ID, Nvl(���﷽ʽ, 0) Into n_����id, n_���﷽ʽ From �ҺŰ��� Where ���� = ����_In;
    Exception
      When Others Then
        n_����id := -1;
    End;
  
    If n_����id = -1 Then
      v_Err_Msg := '����(' || ����_In || ')δ�ҵ�!';
      Raise Err_Item;
    End If;
    --0-�����1-ָ�����ҡ�2-��̬���3-ƽ������,��Ӧ������������
    v_���� := Null;
    If n_���﷽ʽ = 1 Then
      --1-ָ������
      Begin
        Select �������� Into v_���� From �ҺŰ������� Where �ű�id = n_����id;
      Exception
        When Others Then
          v_���� := Null;
      End;
    End If;
    If n_���﷽ʽ = 2 Then
      --2-��̬����:�ø��ű���Һ�δ�������ٵ�����
      For c_���� In (Select ��������, Sum(Num) As Num
                   From (Select ��������, 0 As Num
                          From �ҺŰ�������
                          Where �ű�id = n_����id
                          Union All
                          Select ����, Count(����) As Num
                          From ���˹Һż�¼
                          Where Nvl(ִ��״̬, 0) = 0 And ����ʱ�� Between Trunc(Sysdate) And Sysdate And �ű� = ����_In And
                                ���� In (Select �������� From �ҺŰ������� Where �ű�id = n_����id)
                          Group By ����)
                   Group By ��������
                   Order By Num) Loop
        v_���� := c_����.��������;
        Exit;
      End Loop;
    End If;
    If n_���﷽ʽ = 3 Then
    
      --ƽ�������ǰ����=1��ʾ�´�Ӧȡ�ĵ�ǰ����
      n_Next  := 0;
      n_First := 1;
      For c_���� In (Select Rowid As Rid, �ű�id, ��������, ��ǰ���� From �ҺŰ������� Where �ű�id = n_����id) Loop
        If n_First = 1 Then
          v_Rowid := c_����.Rid;
        End If;
        If n_Next = 1 Then
          v_���� := c_����.��������;
          Update �ҺŰ������� Set ��ǰ���� = 1 Where Rowid = c_����.Rid;
          Exit;
        End If;
        If Nvl(c_����.��ǰ����, 0) = 1 Then
          Update �ҺŰ������� Set ��ǰ���� = 0 Where Rowid = c_����.Rid;
          n_Next := 1;
        End If;
      End Loop;
      If v_���� Is Null Then
        Update �ҺŰ������� Set ��ǰ���� = 1 Where Rowid = v_Rowid Returning �������� Into v_����;
      End If;
    End If;
  
    Return v_����;
  End;
  Procedure Third_Cardbalance_Modfiy
  (
    ����id_In     ����Ԥ����¼.����id%Type,
    �����_In     Varchar2,
    ����_In       ����Ԥ����¼.����%Type,
    ֧�����_In   ����Ԥ����¼.��Ԥ��%Type,
    ������ˮ��_In ����Ԥ����¼.������ˮ��%Type,
    ����˵��_In   ����Ԥ����¼.����˵��%Type,
    Xmlexpned_In  Xmltype
  ) Is
    n_�����id ҽ�ƿ����.Id%Type;
    v_���㷽ʽ ����Ԥ����¼.���㷽ʽ%Type;
    v_����     ҽ�ƿ����.����%Type;
  Begin
    v_Err_Msg := Null;
    Begin
      n_�����id := To_Number(�����_In);
    Exception
      When Others Then
        n_�����id := 0;
    End;
    If n_�����id = 0 Then
      Begin
        Select ID, ���㷽ʽ, Decode(Nvl(�Ƿ�����, 0), 1, Null, ���� || 'δ����,��������нɷ�!'), ����
        Into n_�����id, v_���㷽ʽ, v_Err_Msg, v_����
        From ҽ�ƿ����
        Where ���� = �����_In;
      Exception
        When Others Then
          n_�����id := -1;
          v_Err_Msg  := �����_In || '������!';
      End;
    Else
      Begin
        Select ID, ���㷽ʽ, Decode(Nvl(�Ƿ�����, 0), 1, Null, ���� || 'δ����,��������нɷ�!'), ����
        Into n_�����id, v_���㷽ʽ, v_Err_Msg, v_����
        From ҽ�ƿ����
        Where ID = n_�����id;
      Exception
        When Others Then
          n_�����id := -1;
          v_Err_Msg  := 'δ�ҵ�ָ���Ľ���֧����Ϣ!';
      End;
    End If;
    If Not v_Err_Msg Is Null Then
      Raise Err_Item;
    End If;
    If v_���㷽ʽ Is Null Then
      v_Err_Msg := Nvl(v_����, '') || 'δ���ý��㷽ʽ,����ҽ�ƿ���������ý��㷽ʽ';
      Raise Err_Item;
    End If;
  
    v_�շѽ��� := v_���㷽ʽ || '|' || ֧�����_In || '|' || ' |' || ' ';
    --���㷽ʽ|������|�������|����ժҪ 
    Zl_�����շѽ���_Modify(1, n_����id, ����id_In, v_�շѽ���, 0, 0, n_�����id, ����_In, ������ˮ��_In, ����˵��_In, 0, 0, 0, 0);
  
    --������չ������Ϣ
    For c_��չ In (Select Extractvalue(j.Column_Value, '/EXPEND/JYMC') As Jymc,
                        Extractvalue(j.Column_Value, '/EXPEND/JYLR') As Jylr
                 From Table(Xmlsequence(Extract(Xmlexpned_In, '/EXPENDLIST/EXPEND'))) J) Loop
      Zl_�������㽻��_Insert(n_�����id, 0, ����_In, ����id_In, c_��չ.Jymc || '|' || c_��չ.Jylr, 0);
    End Loop;
  End Third_Cardbalance_Modfiy;

  Procedure Square_Cardbalance_Modfiy
  (
    ����id_In     ����Ԥ����¼.����id%Type,
    �����_In     Varchar2,
    ����_In       ����Ԥ����¼.����%Type,
    ֧�����_In   ����Ԥ����¼.��Ԥ��%Type,
    ������ˮ��_In ����Ԥ����¼.������ˮ��%Type,
    ����˵��_In   ����Ԥ����¼.����˵��%Type,
    Xmlexpned_In  Xmltype
  ) Is
    n_�����id ҽ�ƿ����.Id%Type;
    v_���㷽ʽ ����Ԥ����¼.���㷽ʽ%Type;
    v_����     �����ѽӿ�Ŀ¼.����%Type;
  Begin
    v_Err_Msg := Null;
    Begin
      n_�����id := To_Number(�����_In);
    Exception
      When Others Then
        n_�����id := 0;
    End;
  
    If n_�����id = 0 Then
      Begin
        Select ���, ���㷽ʽ, Decode(Nvl(����, 0), 1, Null, ���� || 'δ����,��������нɷ�!'), ����
        Into n_�����id, v_���㷽ʽ, v_Err_Msg, v_����
        From �����ѽӿ�Ŀ¼
        Where ���� = �����_In;
      Exception
        When Others Then
          v_Err_Msg := '����:' || �����_In || '������!';
      End;
    
    Else
    
      Begin
        Select ���, ���㷽ʽ, Decode(Nvl(����, 0), 1, Null, ���� || 'δ����,��������нɷ�!'), ����
        Into n_�����id, v_���㷽ʽ, v_Err_Msg, v_����
        From �����ѽӿ�Ŀ¼
        Where ��� = �����_In;
      Exception
        When Others Then
          v_Err_Msg := 'δ�ҵ�ָ���Ľ���֧����Ϣ!';
      End;
    
    End If;
    If Not v_Err_Msg Is Null Then
      Raise Err_Item;
    End If;
    If v_���㷽ʽ Is Null Then
      v_Err_Msg := Nvl(v_����, '') || 'δ���ý��㷽ʽ,����ҽ�ƿ���������ý��㷽ʽ';
      Raise Err_Item;
    End If;
  
    Select ID
    Into n_���ѿ�id
    From ���ѿ�Ŀ¼
    Where �ӿڱ�� = n_�����id And ���� = ����_In And
          ��� = (Select Max(���) From ���ѿ�Ŀ¼ Where �ӿڱ�� = n_�����id And ���� = ����_In);
  
    --���㷽ʽ_IN��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||.... 
    v_�շѽ��� := n_�����id || '|' || ����_In || '|' || n_���ѿ�id || '|' || ֧�����_In;
    Zl_�����շѽ���_Modify(3, n_����id, ����id_In, v_�շѽ���, 0, 0, n_�����id, ����_In, ������ˮ��_In, ����˵��_In, 0, 0, 0, 0);
    --������չ������Ϣ
    For c_��չ In (Select Extractvalue(j.Column_Value, '/EXPEND/JYMC') As Jymc,
                        Extractvalue(j.Column_Value, '/EXPEND/JYLR') As Jylr
                 From Table(Xmlsequence(Extract(Xmlexpned_In, '/EXPENDLIST/EXPEND'))) J) Loop
      Zl_�������㽻��_Insert(n_�����id, 1, ����_In, ����id_In, c_��չ.Jymc || '|' || c_��չ.Jylr, 0);
    End Loop;
  End Square_Cardbalance_Modfiy;

Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/NO'), To_Number(Extractvalue(Value(A), 'IN/BRID')),
         To_Number(Extractvalue(Value(A), 'IN/JE')), To_Number(Extractvalue(Value(A), 'IN/WCJE')),
         To_Number(Extractvalue(Value(A), 'IN/SFGH')), Extractvalue(Value(A), 'IN/SFZH'),
         Extractvalue(Value(A), 'IN/XM')
  Into v_Nos, n_����id, n_�շ��ܶ�, n_����, n_�Ƿ�Һ�, v_���֤��, v_����
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  If Nvl(n_����id, 0) = 0 And Not v_���֤�� Is Null And Not v_���� Is Null Then
    n_����id := Zl_Third_Getpatiid(v_���֤��, v_����);
  End If;
  --0.��ؼ��

  If Nvl(n_����id, 0) = 0 Then
    v_Err_Msg := '������Чʶ�������,������ɷ�!';
    Raise Err_Item;
  
  End If;

  If v_Nos Is Null Then
    v_Err_Msg := 'û��ָ����ص��շѵ���,������ɷ�!';
    Raise Err_Item;
  
  End If;

  --��Աid,��Ա���,��Ա���� 
  v_Temp := Zl_Identity(1);
  If Nvl(v_Temp, '0') = '0' Or Nvl(v_Temp, '_') = '_' Then
    v_Err_Msg := 'ϵͳ�����ϱ���Ч�Ĳ���Ա,������ɷ�!';
    Raise Err_Item;
  End If;
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_����Ա���� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_����Ա���� := v_Temp;
  v_Err_Msg    := Null;
  Begin
    Select b.����, b.����, a.����, a.�Ա�, a.����
    Into v_ҽ�Ƹ��ʽ����, v_���ʽ, v_����, v_�Ա�, v_����
    From ������Ϣ A, ҽ�Ƹ��ʽ B
    Where a.ҽ�Ƹ��ʽ = b.����(+) And a.����id = n_����id;
  Exception
    When Others Then
      v_Err_Msg := 'ָ���Ľɷѵ����в�����Чʶ����,������ɷ�!';
  End;
  If Not v_Err_Msg Is Null Then
    Raise Err_Item;
  End If;

  Select Decode(Nvl(n_�Ƿ�Һ�, 0), 0, 3, 4) Into n_ҵ������ From Dual;

  For c_���׼�¼ In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As ���㿨���,
                        Extractvalue(b.Column_Value, '/JS/JSFS') As ���㷽ʽ,
                        Extractvalue(b.Column_Value, '/JS/JYLSH') As ������ˮ��,
                        Extractvalue(b.Column_Value, '/JS/JSKH') As ���㿨��, Extractvalue(b.Column_Value, '/JS/ZY') As ժҪ
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
  
    If c_���׼�¼.���㿨��� Is Null Then
      v_����� := c_���׼�¼.���㷽ʽ;
    Else
      Select Decode(Translate(Nvl(c_���׼�¼.���㿨���, 'abcd'), '#1234567890', '#'), Null, 1, 0)
      Into n_Count
      From Dual;
    
      If Nvl(n_Count, 0) = 1 Then
        Select Max(����) Into v_����� From ҽ�ƿ���� Where ID = To_Number(c_���׼�¼.���㿨���);
      Else
        Select Max(����) Into v_����� From ҽ�ƿ���� Where ���� = c_���׼�¼.���㿨���;
      End If;
    End If;
    If v_����� Is Null Then
      v_Err_Msg := '��֧�ֵĽ��㷽ʽ,���飡';
      Raise Err_Item;
    End If;
  
    If Zl_Fun_�������׼�¼_Locked(v_�����, c_���׼�¼.������ˮ��, c_���׼�¼.���㿨��, c_���׼�¼.ժҪ, n_ҵ������) = 0 Then
      v_Err_Msg := '������ˮ��Ϊ:' || c_���׼�¼.������ˮ�� || '�Ľ������ڽ����У��������ٴ��ύ�˽���!';
      Raise Err_Special;
    End If;
  End Loop;

  Select ���˽��ʼ�¼_Id.Nextval, Sysdate Into n_����id, d_�շ�ʱ�� From Dual;
  n_������� := -1 * n_����id;

  If Nvl(n_�Ƿ�Һ�, 0) = 0 Then
    --���õ���
    v_��ҩ���� := Zl_Getclinicchargepaywins(v_Nos);
  
    --1.���з����շѴ���
    --��ȡ��ҩ����
  
    n_���ʽ�� := 0;
    For c_�ɷѵ� In (Select /*+ rule */
                   a.No, Max(a.��������id) As ��������id, Max(a.���˿���id) As ���˿���id, Max(a.����id) As ����id, Sum(ʵ�ս��) As ʵ�ս��,
                   Max(a.������) As ������
                  From ������ü�¼ A, Table(f_Str2list(v_Nos)) J
                  Where a.��¼���� = 1 And a.No = j.Column_Value And a.��¼״̬ = 0
                  Group By a.No) Loop
      If Nvl(c_�ɷѵ�.����id, 0) <> n_����id Then
        v_Err_Msg := '�ɷѵ���:' || c_�ɷѵ�.No || '�뵱ǰ������ݲ���,������ɷ�!';
        Raise Err_Item;
      End If;
    
      n_���ʽ�� := n_���ʽ�� + Nvl(c_�ɷѵ�.ʵ�ս��, 0);
      Zl_���˻����շ�_Insert(c_�ɷѵ�.No, n_����id, 1, v_ҽ�Ƹ��ʽ����, v_����, v_�Ա�, v_����, c_�ɷѵ�.���˿���id, c_�ɷѵ�.��������id, c_�ɷѵ�.������, n_����id,
                       d_�շ�ʱ��, v_����Ա����, v_����Ա����, v_��ҩ����, 0, d_�շ�ʱ��);
    
    End Loop;
  
    --����ܽ���Ƿ���ȷ 
    If Nvl(n_����, 0) = 0 Then
      n_���� := Nvl(n_���ʽ��, 0) - Nvl(n_�շ��ܶ�, 0);
      If Abs(n_����) > 1.00 Then
        v_Err_Msg := '���ݽɿ�����ʵ�ʽ�������̫��!';
        Raise Err_Item;
      End If;
    End If;
    If Nvl(n_���ʽ��, 0) <> Nvl(n_�շ��ܶ�, 0) + Nvl(n_����, 0) Then
      v_Err_Msg := 'ָ���Ľɷѵ��ݵ��ܽ���,������ѡ��ɷѵ���!';
      Raise Err_Item;
    End If;
  
    --2.ȷ��֧����ʽ
    n_Count := 0;
    For c_���㷽ʽ In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As ���㿨���,
                          Extractvalue(b.Column_Value, '/JS/JSKH') As ���㿨��,
                          Extractvalue(b.Column_Value, '/JS/JSFS') As ���㷽ʽ,
                          Extractvalue(b.Column_Value, '/JS/JSJE') As ������,
                          Extractvalue(b.Column_Value, '/JS/JYLSH') As ������ˮ��,
                          Extractvalue(b.Column_Value, '/JS/JYSM') As ����˵��, Extractvalue(b.Column_Value, '/JS/ZY') As ժҪ,
                          Extractvalue(b.Column_Value, '/JS/SFCYJ') As �Ƿ��Ԥ��,
                          Extractvalue(b.Column_Value, '/JS/SFXFK') As �Ƿ����ѿ�,
                          Extract(b.Column_Value, '/JS/EXPENDLIST') As Expend
                   From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
    
      --1.����������
    
      If c_���㷽ʽ.���㿨��� Is Not Null And Nvl(c_���㷽ʽ.�Ƿ����ѿ�, 0) = 0 Then
        --1.����������
        Third_Cardbalance_Modfiy(n_����id, c_���㷽ʽ.���㿨���, c_���㷽ʽ.���㿨��, c_���㷽ʽ.������, c_���㷽ʽ.������ˮ��, c_���㷽ʽ.����˵��,
                                 c_���㷽ʽ.Expend);
      Elsif c_���㷽ʽ.���㿨��� Is Not Null And Nvl(c_���㷽ʽ.�Ƿ����ѿ�, 0) = 1 Then
        --2.���ѿ�����
        Square_Cardbalance_Modfiy(n_����id, c_���㷽ʽ.���㿨���, c_���㷽ʽ.���㿨��, c_���㷽ʽ.������, c_���㷽ʽ.������ˮ��, c_���㷽ʽ.����˵��,
                                  c_���㷽ʽ.Expend);
      Elsif Nvl(c_���㷽ʽ.�Ƿ��Ԥ��, 0) = 1 Then
        --3.��Ԥ����
        Zl_�����շѽ���_Modify(0, n_����id, n_����id, Null, c_���㷽ʽ.������, 0, Null, Null, Null, Null, 0, 0, 0, 0);
      Else
        --4.��ͨ����
        If c_���㷽ʽ.���㷽ʽ Is Null Then
          v_Err_Msg := 'δָ��ָ����ʽ�����ʽɿ�!';
          Raise Err_Item;
        End If;
        --���㷽ʽ|������|�������|����ժҪ||..
        v_�շѽ��� := c_���㷽ʽ.���㷽ʽ || '|' || c_���㷽ʽ.������ || '| | ';
        v_��ͨ���� := Nvl(v_��ͨ����, '') || '||' || v_�շѽ���;
      End If;
    
      If c_���㷽ʽ.���㿨��� Is Null Then
        v_����� := c_���㷽ʽ.���㷽ʽ;
      Else
        Select Decode(Translate(Nvl(c_���㷽ʽ.���㿨���, 'abcd'), '#1234567890', '#'), Null, 1, 0)
        Into n_Count
        From Dual;
      
        If Nvl(n_Count, 0) = 1 Then
          Select Max(����) Into v_����� From ҽ�ƿ���� Where ID = To_Number(c_���㷽ʽ.���㿨���);
        Else
          Select Max(����) Into v_����� From ҽ�ƿ���� Where ���� = c_���㷽ʽ.���㿨���;
        End If;
      End If;
    
      Update �������׼�¼
      Set ҵ�����id = n_�������
      Where ��ˮ�� = c_���㷽ʽ.������ˮ�� And ��� = v_����� And ҵ������ = n_ҵ������;
      n_Count := n_Count + 1;
    End Loop;
    If n_Count = 0 Then
      v_Err_Msg := '������Чȷ�ϵ�ǰ��֧����ʽ!';
      Raise Err_Item;
    End If;
    --5.��ͨ���㼰��ɽ�
    If v_��ͨ���� Is Not Null Then
      v_��ͨ���� := Substr(v_��ͨ����, 3);
    End If;
    Zl_�����շѽ���_Modify(0, n_����id, n_����id, v_��ͨ����, Null, 0, Null, Null, Null, Null, 0, 0, n_����, 1);
  Else
    n_���ʽ�� := 0;
    --�Һŵ���
    v_Para     := zl_GetSysParameter(256);
    n_�Һ�ģʽ := Substr(v_Para, 1, 1);
    Begin
      d_����ʱ�� := To_Date(Substr(v_Para, 3), 'yyyy-mm-dd hh24:mi:ss');
    Exception
      When Others Then
        d_����ʱ�� := Null;
    End;
    For c_���� In (Select 1 As ˳���, b.No, b.�վݷ�Ŀ, b.����id, b.ִ�в���id, b.���˿���id, b.������, b.�շ����, b.������Ŀid, b.���ӱ�־,
                        To_Char(b.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��, b.�۸񸸺�, b.��������, b.���, b.�շ�ϸĿid, b.���㵥λ,
                        Max(m.����) As ����, Max(m.���) As ���, Sum(b.��׼����) As ����, Avg(Nvl(b.����, 1) * b.����) As ����,
                        Sum(b.Ӧ�ս��) As Ӧ�ս��, Sum(b.ʵ�ս��) As ʵ�ս��, Max(j.����) As ��������, Max(q.����) As ִ�п���
                 From ������ü�¼ B, �շ���ĿĿ¼ M, ���ű� J, ���ű� Q
                 Where b.No = v_Nos And b.��¼���� = 4 And Nvl(b.����״̬, 0) = 0 And b.��¼״̬ = 0 And b.�շ�ϸĿid = m.Id And
                       b.��������id = j.Id(+) And b.ִ�в���id = q.Id(+)
                 Group By b.No, b.�վݷ�Ŀ, b.����id, b.ִ�в���id, b.���˿���id, b.������, b.������Ŀid, b.�շ����, b.�Ǽ�ʱ��, b.�۸񸸺�, b.��������, b.���,
                          b.�շ�ϸĿid, b.���㵥λ, b.���ӱ�־
                 Order By ���) Loop
      Zl_����ԤԼ�Һż�¼_Update(c_����.No, c_����.���, c_����.�۸񸸺�, c_����.��������, c_����.�շ����, c_����.�շ�ϸĿid, c_����.����, c_����.����, c_����.������Ŀid,
                         c_����.�վݷ�Ŀ, c_����.Ӧ�ս��, c_����.ʵ�ս��, c_����.���ӱ�־, Null, Null, Null, Null, c_����.���˿���id, c_����.ִ�в���id);
      n_���ʽ��   := n_���ʽ�� + c_����.ʵ�ս��;
      n_���       := c_����.���;
      n_���˿���id := c_����.���˿���id;
      n_ִ�в���id := c_����.ִ�в���id;
      v_No         := c_����.No;
    End Loop;
  
    Select a.ִ�в���id, a.�շ�ϸĿid, c.Id, a.ִ����, b.�ű�, b.�����, b.����ʱ��, a.�ѱ�, b.����, b.�����¼id
    Into n_����id, n_��Ŀid, n_ҽ��id, v_ҽ������, v_����, n_�����, d_����ʱ��, v_�ѱ�, n_����, n_�����¼id
    From ������ü�¼ A, ���˹Һż�¼ B, ��Ա�� C
    Where a.No = v_Nos And a.��¼���� = 4 And a.��� = 1 And a.No = b.No And a.ִ���� = c.����(+);
  
    Begin
      Select Zl_Fun_Customregexpenses(n_����id, 0, v_����, v_����, v_�Ա�, v_����, v_���֤��, v_�ѱ�, v_���ʽ) Into v_������Ŀid From Dual;
    Exception
      When Others Then
        v_������Ŀid := Null;
    End;
    If v_������Ŀid Is Not Null Then
      IF Instr(v_������Ŀid, '|') > 0 Then
        v_�������� := v_������Ŀid || ','; --�Կո�ֿ���|��β,û�н�������
        v_������Ŀid := '';
        n_cursor   := 0;
        While v_�������� Is Not Null Loop
          v_����ֵ := Substr(v_��������, 1, Instr(v_��������, ',') - 1);
          Price_rec.��ĿID := To_Number(Substr(v_����ֵ, 1, Instr(v_����ֵ, '|') - 1));
        
          v_����ֵ := Substr(v_����ֵ, Instr(v_����ֵ, '|') + 1);
          Price_rec.���� := To_Number(Substr(v_����ֵ, 1, Instr(v_����ֵ, '|') - 1));
        
          v_����ֵ := Substr(v_����ֵ, Instr(v_����ֵ, '|') + 1);
          Price_rec.���� := To_Number(Substr(v_����ֵ, 1, Instr(v_����ֵ, '|') - 1));
          
          v_����ֵ := Substr(v_����ֵ, Instr(v_����ֵ, '|') + 1);
          Price_rec.Ӧ�� := To_Number(Substr(v_����ֵ, 1, Instr(v_����ֵ, '|') - 1));
        
          v_����ֵ := Substr(v_����ֵ, Instr(v_����ֵ, '|') + 1);
          Price_rec.ʵ�� := To_Number(v_����ֵ);
          n_cursor := n_cursor + 1;
          Price_rec_array(n_cursor):=Price_rec;
          v_�������� := Substr(v_��������, Instr(v_��������, ',') + 1);
          v_������Ŀid := v_������Ŀid || ',' || Price_rec_array(n_cursor).��ĿID;
        End Loop;
        
        If v_������Ŀid is not null then
          v_������Ŀid := substr(v_������Ŀid, 2);
        End if;
        
        For c_������Ŀ In (Select /*+cardinality(D,10)*/
                        5 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, 1 As ����, c.Id As ������Ŀid,
                        c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, -1 As ִ�п�������
                       From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C, Table(f_Str2list(v_������Ŀid)) D
                       Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = d.Column_Value And Sysdate Between b.ִ������ And
                             Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))) Loop
        
          n_��� := n_��� + 1;
          Zl_����ԤԼ�Һż�¼_Update(v_No, n_���, Null, Null, c_������Ŀ.���, c_������Ŀ.��Ŀid, Price_rec_array(n_cursor).����, Price_rec_array(n_cursor).����, c_������Ŀ.������Ŀid,
                               c_������Ŀ.�վݷ�Ŀ, Price_rec_array(n_cursor).Ӧ��, Price_rec_array(n_cursor).ʵ��, Null, Null, Null, Null, Null, n_���˿���id,
                               n_ִ�в���id);

          n_ʵ�ս�� := Price_rec_array(n_cursor).ʵ��;                  
          n_���ʽ�� := n_���ʽ�� + n_ʵ�ս��;
        End Loop;                        
      Else
        For c_������Ŀ In (Select /*+cardinality(D,10)*/
                        5 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, 1 As ����, c.Id As ������Ŀid,
                        c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, -1 As ִ�п�������
                       From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C, Table(f_Str2list(v_������Ŀid)) D
                       Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = d.Column_Value And Sysdate Between b.ִ������ And
                             Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))
                       Union All
                       Select /*+cardinality(E,10)*/
                        6 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, d.�������� As ����,
                        c.Id As ������Ŀid, c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, -1 As ִ�п�������
                       From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C, �շѴ�����Ŀ D, Table(f_Str2list(v_������Ŀid)) E
                       Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = d.����id And d.����id = e.Column_Value And
                             Sysdate Between b.ִ������ And Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))) Loop
          n_��� := n_��� + 1;
          If c_������Ŀ.���� = 5 Then
            n_�������� := n_���;
          End If;
        
          v_ʵ��     := Zl_Actualmoney(v_�ѱ�, c_������Ŀ.��Ŀid, c_������Ŀ.������Ŀid, c_������Ŀ.���� * c_������Ŀ.����);
          n_ʵ�ս�� := To_Number(Substr(v_ʵ��, Instr(v_ʵ��, ':') + 1));
        
          If c_������Ŀ.���� = 5 Then
            Zl_����ԤԼ�Һż�¼_Update(v_No, n_���, Null, Null, c_������Ŀ.���, c_������Ŀ.��Ŀid, c_������Ŀ.����, c_������Ŀ.����, c_������Ŀ.������Ŀid,
                               c_������Ŀ.�վݷ�Ŀ, c_������Ŀ.���� * c_������Ŀ.����, n_ʵ�ս��, Null, Null, Null, Null, Null, n_���˿���id,
                               n_ִ�в���id);
          Else
            Zl_����ԤԼ�Һż�¼_Update(v_No, n_���, Null, n_��������, c_������Ŀ.���, c_������Ŀ.��Ŀid, c_������Ŀ.����, c_������Ŀ.����, c_������Ŀ.������Ŀid,
                               c_������Ŀ.�վݷ�Ŀ, c_������Ŀ.���� * c_������Ŀ.����, n_ʵ�ս��, Null, Null, Null, Null, Null, n_���˿���id,
                               n_ִ�в���id);
          End If;
          n_���ʽ�� := n_���ʽ�� + n_ʵ�ս��;
        
        End Loop;
      End IF;
    End If;
  
    --����ܽ���Ƿ���ȷ 
    If Nvl(n_����, 0) = 0 Then
      n_���� := Nvl(n_���ʽ��, 0) - Nvl(n_�շ��ܶ�, 0);
      If Abs(n_����) > 1.00 Then
        v_Err_Msg := '���ݽɿ�����ʵ�ʽ�������̫��!';
        Raise Err_Item;
      End If;
    End If;
    If Nvl(n_���ʽ��, 0) <> Nvl(n_�շ��ܶ�, 0) + Nvl(n_����, 0) Then
      Select Max(����Ա����) Into v_����Ա From ������ü�¼ Where ��¼���� = 4 And NO = v_Nos;
      If v_����Ա = v_����Ա���� Then
        v_Err_Msg := 'ָ���Ľɷѵ��ݵ��ܽ���,������ѡ��ɷѵ���!';
        Raise Err_Special;
      Else
        v_Err_Msg := 'ָ���Ľɷѵ��ݵ��ܽ���,������ѡ��ɷѵ���!';
        Raise Err_Item;
      End If;
    End If;
  
    --ԤԼ����
    If n_�Һ�ģʽ = 1 Then
      If d_����ʱ�� > d_����ʱ�� And n_�����¼id Is Null Then
        n_�Һ�ģʽ := 0;
      End If;
    End If;
  
    Select Decode(To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113, 100)), 0, 0, 1) Into n_���ɶ��� From Dual;
    If n_�Һ�ģʽ = 0 Then
      For c_���㷽ʽ In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As ���㿨���,
                            Extractvalue(b.Column_Value, '/JS/JSKH') As ���㿨��,
                            Extractvalue(b.Column_Value, '/JS/JSFS') As ���㷽ʽ,
                            Extractvalue(b.Column_Value, '/JS/JSJE') As ������,
                            Extractvalue(b.Column_Value, '/JS/JYLSH') As ������ˮ��,
                            Extractvalue(b.Column_Value, '/JS/JYSM') As ����˵��,
                            Extractvalue(b.Column_Value, '/JS/ZY') As ժҪ,
                            Extractvalue(b.Column_Value, '/JS/SFCYJ') As �Ƿ��Ԥ��,
                            Extractvalue(b.Column_Value, '/JS/SFXFK') As �Ƿ����ѿ�
                     From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
        If Nvl(c_���㷽ʽ.�Ƿ��Ԥ��, 0) = 1 Then
          n_Ԥ��֧�� := c_���㷽ʽ.������;
        Else
          If c_���㷽ʽ.���㷽ʽ Is Not Null Then
            Select Nvl(Max(1), 0) Into n_Exists From ���㷽ʽ Where ���� = c_���㷽ʽ.���㷽ʽ And ���� In (3, 4);
            If n_Exists = 1 Then
              n_ҽ��֧�� := c_���㷽ʽ.������;
            Else
              If Nvl(n_��֧ͨ��, 0) = 0 Then
                n_��֧ͨ�� := c_���㷽ʽ.������;
                v_���㷽ʽ := c_���㷽ʽ.���㷽ʽ;
                If Nvl(c_���㷽ʽ.�Ƿ����ѿ�, 0) = 1 Then
                  Begin
                    n_���㿨��� := To_Number(c_���㷽ʽ.���㿨���);
                  Exception
                    When Others Then
                      n_���㿨��� := 0;
                  End;
                  If n_���㿨��� = 0 Then
                    Begin
                      Select ���
                      Into n_���㿨���
                      From �����ѽӿ�Ŀ¼
                      Where ���� = c_���㷽ʽ.���㿨��� And Nvl(����, 0) = 1;
                    Exception
                      When Others Then
                        v_Err_Msg := 'δ�ҵ���Ӧ�����ѿ�!';
                        Raise Err_Item;
                    End;
                  End If;
                  If v_���㷽ʽ Is Null Then
                    Select ���㷽ʽ Into v_���㷽ʽ From �����ѽӿ�Ŀ¼ Where ��� = n_���㿨���;
                  End If;
                Else
                  Begin
                    n_�����id := To_Number(c_���㷽ʽ.���㿨���);
                  Exception
                    When Others Then
                      n_�����id := 0;
                  End;
                  If n_�����id = 0 Then
                    Begin
                      Select ID
                      Into n_�����id
                      From ҽ�ƿ����
                      Where ���� = c_���㷽ʽ.���㿨��� And Nvl(�Ƿ�����, 0) = 1;
                    Exception
                      When Others Then
                        v_Err_Msg := 'δ�ҵ���Ӧ��ҽ�ƿ�!';
                        Raise Err_Item;
                    End;
                  End If;
                  If v_���㷽ʽ Is Null Then
                    Select ���㷽ʽ Into v_���㷽ʽ From ҽ�ƿ���� Where ID = n_�����id;
                  End If;
                End If;
                v_���㿨��   := c_���㷽ʽ.���㿨��;
                v_������ˮ�� := c_���㷽ʽ.������ˮ��;
                v_����˵��   := c_���㷽ʽ.����˵��;
                v_ժҪ       := c_���㷽ʽ.ժҪ;
              Else
                v_Err_Msg := '�ҺŽ����ݲ�֧�ֶ��ֽ��㷽ʽ!';
                Raise Err_Item;
              End If;
            End If;
          Else
            If Nvl(n_��֧ͨ��, 0) = 0 Then
              n_��֧ͨ�� := c_���㷽ʽ.������;
              v_���㷽ʽ := c_���㷽ʽ.���㷽ʽ;
              If Nvl(c_���㷽ʽ.�Ƿ����ѿ�, 0) = 1 Then
                Begin
                  n_���㿨��� := To_Number(c_���㷽ʽ.���㿨���);
                Exception
                  When Others Then
                    n_���㿨��� := 0;
                End;
                If n_���㿨��� = 0 Then
                  Begin
                    Select ���
                    Into n_���㿨���
                    From �����ѽӿ�Ŀ¼
                    Where ���� = c_���㷽ʽ.���㿨��� And Nvl(����, 0) = 1;
                  Exception
                    When Others Then
                      v_Err_Msg := 'δ�ҵ���Ӧ�����ѿ�!';
                      Raise Err_Item;
                  End;
                End If;
                If v_���㷽ʽ Is Null Then
                  Select ���㷽ʽ Into v_���㷽ʽ From �����ѽӿ�Ŀ¼ Where ��� = n_���㿨���;
                End If;
              Else
                Begin
                  n_�����id := To_Number(c_���㷽ʽ.���㿨���);
                Exception
                  When Others Then
                    n_�����id := 0;
                End;
                If n_�����id = 0 Then
                  Begin
                    Select ID
                    Into n_�����id
                    From ҽ�ƿ����
                    Where ���� = c_���㷽ʽ.���㿨��� And Nvl(�Ƿ�����, 0) = 1;
                  Exception
                    When Others Then
                      v_Err_Msg := 'δ�ҵ���Ӧ��ҽ�ƿ�!';
                      Raise Err_Item;
                  End;
                End If;
                If v_���㷽ʽ Is Null Then
                  Select ���㷽ʽ Into v_���㷽ʽ From ҽ�ƿ���� Where ID = n_�����id;
                End If;
              End If;
              v_���㿨��   := c_���㷽ʽ.���㿨��;
              v_������ˮ�� := c_���㷽ʽ.������ˮ��;
              v_����˵��   := c_���㷽ʽ.����˵��;
              v_ժҪ       := c_���㷽ʽ.ժҪ;
            Else
              v_Err_Msg := '�ҺŽ����ݲ�֧�ֶ��ֽ��㷽ʽ!';
              Raise Err_Item;
            End If;
          End If;
        End If;
        If c_���㷽ʽ.���㿨��� Is Null Then
          v_����� := c_���㷽ʽ.���㷽ʽ;
        Else
          Select Decode(Translate(Nvl(c_���㷽ʽ.���㿨���, 'abcd'), '#1234567890', '#'), Null, 1, 0)
          Into n_Count
          From Dual;
        
          If Nvl(n_Count, 0) = 1 Then
            Select Max(����) Into v_����� From ҽ�ƿ���� Where ID = To_Number(c_���㷽ʽ.���㿨���);
          Else
            Select Max(����) Into v_����� From ҽ�ƿ���� Where ���� = c_���㷽ʽ.���㿨���;
          End If;
        End If;
        Update �������׼�¼
        Set ҵ�����id = n_����id
        Where ��ˮ�� = c_���㷽ʽ.������ˮ�� And ��� = v_����� And ҵ������ = n_ҵ������;
      End Loop;
      Zl_ԤԼ�ҺŽ���_Insert(v_Nos, Null, Null, n_����id, Zl_����(v_����), n_����id, n_�����, v_����, v_�Ա�, v_����, v_ҽ�Ƹ��ʽ����, v_�ѱ�,
                       v_���㷽ʽ, n_��֧ͨ��, n_Ԥ��֧��, n_ҽ��֧��, d_����ʱ��, n_����, v_����Ա����, v_����Ա����, n_���ɶ���, d_�շ�ʱ��, n_�����id, n_���㿨���,
                       v_���㿨��, v_������ˮ��, v_����˵��, Null, 0, 0, Null, 1);
    Else
      For c_���㷽ʽ In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As ���㿨���,
                            Extractvalue(b.Column_Value, '/JS/JSKH') As ���㿨��,
                            Extractvalue(b.Column_Value, '/JS/JSFS') As ���㷽ʽ,
                            Extractvalue(b.Column_Value, '/JS/JSJE') As ������,
                            Extractvalue(b.Column_Value, '/JS/JYLSH') As ������ˮ��,
                            Extractvalue(b.Column_Value, '/JS/JYSM') As ����˵��,
                            Extractvalue(b.Column_Value, '/JS/ZY') As ժҪ,
                            Extractvalue(b.Column_Value, '/JS/SFCYJ') As �Ƿ��Ԥ��,
                            Extractvalue(b.Column_Value, '/JS/SFXFK') As �Ƿ����ѿ�
                     From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
        If Nvl(c_���㷽ʽ.�Ƿ��Ԥ��, 0) = 1 Then
          n_Ԥ��֧�� := c_���㷽ʽ.������;
        Else
          n_��֧ͨ�� := Nvl(n_��֧ͨ��, 0) + c_���㷽ʽ.������;
          If c_���㷽ʽ.���㷽ʽ Is Null Then
            --���������㷽ʽ
            If Nvl(c_���㷽ʽ.�Ƿ����ѿ�, 0) = 1 Then
              Begin
                n_���㿨��� := To_Number(c_���㷽ʽ.���㿨���);
              Exception
                When Others Then
                  n_���㿨��� := 0;
              End;
              If n_���㿨��� = 0 Then
                Begin
                  Select ���
                  Into n_���㿨���
                  From �����ѽӿ�Ŀ¼
                  Where ���� = c_���㷽ʽ.���㿨��� And Nvl(����, 0) = 1;
                Exception
                  When Others Then
                    v_Err_Msg := 'δ�ҵ���Ӧ�����ѿ�!';
                    Raise Err_Item;
                End;
              End If;
              Select ���㷽ʽ Into v_��ʱ���㷽ʽ From �����ѽӿ�Ŀ¼ Where ��� = n_���㿨���;
            Else
              Begin
                n_�����id := To_Number(c_���㷽ʽ.���㿨���);
              Exception
                When Others Then
                  n_�����id := 0;
              End;
              If n_�����id = 0 Then
                Begin
                  Select ID
                  Into n_�����id
                  From ҽ�ƿ����
                  Where ���� = c_���㷽ʽ.���㿨��� And Nvl(�Ƿ�����, 0) = 1;
                Exception
                  When Others Then
                    v_Err_Msg := 'δ�ҵ���Ӧ��ҽ�ƿ�!';
                    Raise Err_Item;
                End;
              End If;
              Select ���㷽ʽ Into v_��ʱ���㷽ʽ From ҽ�ƿ���� Where ID = n_�����id;
            End If;
            v_���㿨��   := c_���㷽ʽ.���㿨��;
            v_������ˮ�� := c_���㷽ʽ.������ˮ��;
            v_����˵��   := c_���㷽ʽ.����˵��;
            v_ժҪ       := c_���㷽ʽ.ժҪ;
            v_���㷽ʽ   := v_���㷽ʽ || '|' || v_��ʱ���㷽ʽ || ',' || c_���㷽ʽ.������ || ',,1';
          Else
            --�������㷽ʽ
            v_���㷽ʽ := v_���㷽ʽ || '|' || c_���㷽ʽ.���㷽ʽ || ',' || c_���㷽ʽ.������ || ',,1';
          End If;
        End If;
        If c_���㷽ʽ.���㿨��� Is Null Then
          v_����� := c_���㷽ʽ.���㷽ʽ;
        Else
          Select Decode(Translate(Nvl(c_���㷽ʽ.���㿨���, 'abcd'), '#1234567890', '#'), Null, 1, 0)
          Into n_Count
          From Dual;
        
          If Nvl(n_Count, 0) = 1 Then
            Select Max(����) Into v_����� From ҽ�ƿ���� Where ID = To_Number(c_���㷽ʽ.���㿨���);
          Else
            Select Max(����) Into v_����� From ҽ�ƿ���� Where ���� = c_���㷽ʽ.���㿨���;
          End If;
        End If;
        Update �������׼�¼
        Set ҵ�����id = n_����id
        Where ��ˮ�� = c_���㷽ʽ.������ˮ�� And ��� = v_����� And ҵ������ = n_ҵ������;
      End Loop;
      If v_���㷽ʽ Is Not Null Then
        v_���㷽ʽ := Substr(v_���㷽ʽ, 2);
      End If;
      Zl_ԤԼ�ҺŽ���_����_Insert(v_Nos, Null, Null, n_����id, Zl_��������(n_�����¼id), n_����id, n_�����, v_����, v_�Ա�, v_����, v_ҽ�Ƹ��ʽ����,
                          v_�ѱ�, v_���㷽ʽ, n_��֧ͨ��, n_Ԥ��֧��, Null, d_����ʱ��, n_����, v_����Ա����, v_����Ա����, n_���ɶ���, d_�շ�ʱ��, n_�����id,
                          n_���㿨���, v_���㿨��, v_������ˮ��, v_����˵��, Null, 0, 0, Null, 1);
    End If;
    --������չ��Ϣ
    If Nvl(n_�����id, 0) <> 0 Then
      For c_��չ��Ϣ In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As ��������,
                            Extractvalue(b.Column_Value, '/EXPEND/JYLR') As ��������
                     From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS/EXPENDLIST/EXPEND'))) B) Loop
        Zl_�������㽻��_Insert(n_�����id, 0, v_���㿨��, n_����id, c_��չ��Ϣ.�������� || '|' || c_��չ��Ϣ.��������, 0);
      End Loop;
    End If;
    If Nvl(n_���㿨���, 0) <> 0 Then
      For c_��չ��Ϣ In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As ��������,
                            Extractvalue(b.Column_Value, '/EXPEND/JYLR') As ��������
                     From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS/EXPENDLIST/EXPEND'))) B) Loop
        Zl_�������㽻��_Insert(n_���㿨���, 1, v_���㿨��, n_����id, c_��չ��Ϣ.�������� || '|' || c_��չ��Ϣ.��������, 0);
      End Loop;
    End If;
    --�������
    Zl_���˹ҺŻ���_Update(v_ҽ������, n_ҽ��id, n_��Ŀid, n_����id, d_����ʱ��, 2, v_����, 1, n_�����¼id);
  End If;
  v_Temp := '<CZSJ>' || To_Char(d_�շ�ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<JZID>' || n_����id || '</JZID>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Err_Special Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20105, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Payment;
/

--126802:���ϴ�,2018-06-21,���ӷѷ��ع̶��Ľ����Ϣ
Create Or Replace Function Zl_Fun_Customregexpenses
(
  ����id_In       In ������Ϣ.����id%Type,
  ����_In         In ������Ϣ.����%Type,
  ����_In         In �ҺŰ���.����%Type,
  ����_In         In ������Ϣ.����%Type := Null,
  �Ա�_In         In ������Ϣ.�Ա�%Type := Null,
  ����_In         In ������Ϣ.����%Type := Null,
  ���֤��_In     In ������Ϣ.���֤��%Type := Null,
  �ѱ�_In         In ������Ϣ.�ѱ�%Type := Null,
  ҽ�Ƹ��ʽ_In In ������Ϣ.ҽ�Ƹ��ʽ%Type := Null
) Return Varchar2
--    ���ܣ��ҺŸ��ӷѴ�����Ŀ�û��Զ��庯��
  --    ������
  --        ����ID_In��������Ϣ.����ID
  --        ����_In��������Ϣ.����
  --        ����_In: �ҺŰ���.����
  --    ����: ��ʽһ���շ�ϸĿID1|����1|����1|Ӧ��1|ʵ��1,�շ�ϸĿID2|����2....����շ�ϸĿ�ö��ŷָ�,��Ŀ��Ӧ�ա�ʵ�յ���Ϣ���Է��ص�ֵΪ׼��
  --          ��ʽ�����շ�ϸĿID1,�շ�ϸĿID2...ֻ�����շ�ϸĿIDʱ���շѼ�ĿΪ׼��
  --    ����NULLʱ��������,���ܷ�����ͬ���շ�ϸĿID
 Is
Begin
  Return Null;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Fun_Customregexpenses;
/

--127450:���ϴ�,2018-06-20,����˿�ʱ�����˿��¼�ĳ�Ԥ����Ϣ�����ⱻ��Ԥ���ٴ�ʹ��
Create Or Replace Procedure Zl_����Ԥ����¼_Insert
(
  Id_In           ����Ԥ����¼.Id%Type,
  ���ݺ�_In       ����Ԥ����¼.No%Type,
  Ʊ�ݺ�_In       Ʊ��ʹ����ϸ.����%Type,
  ����id_In       ����Ԥ����¼.����id%Type,
  ��ҳid_In       ����Ԥ����¼.��ҳid%Type,
  ����id_In       ����Ԥ����¼.����id%Type,
  ���_In         ����Ԥ����¼.���%Type,
  ���㷽ʽ_In     ����Ԥ����¼.���㷽ʽ%Type,
  �������_In     ����Ԥ����¼.�������%Type,
  �ɿλ_In     ����Ԥ����¼.�ɿλ%Type,
  ��λ������_In   ����Ԥ����¼.��λ������%Type,
  ��λ�ʺ�_In     ����Ԥ����¼.��λ�ʺ�%Type,
  ժҪ_In         ����Ԥ����¼.ժҪ%Type,
  ����Ա���_In   ����Ԥ����¼.����Ա���%Type,
  ����Ա����_In   ����Ԥ����¼.����Ա����%Type,
  ����id_In       Ʊ��ʹ����ϸ.����id%Type,
  Ԥ�����_In     ����Ԥ����¼.Ԥ�����%Type := Null,
  �����id_In     ����Ԥ����¼.�����id%Type := Null,
  ���㿨���_In   ����Ԥ����¼.���㿨���%Type := Null,
  ����_In         ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In   ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In     ����Ԥ����¼.����˵��%Type := Null,
  ������λ_In     ����Ԥ����¼.������λ%Type := Null,
  �տ�ʱ��_In     ����Ԥ����¼.�տ�ʱ��%Type := Null,
  ��������_In     Integer := 0,
  ��������_In     ����Ԥ����¼.��������%Type := Null,
  ���½������_In Number := 0,
  �˿���_In     Number := 0,
  ǿ������_In     Number := 0,
  �Ƿ�ת��_In     Number := 0
) As
  ----------------------------------------------
  --��������_In:0-������Ԥ��;1-��Ϊ���۵�;3-����˿�
  --�˿���_In;0-�����˿����Ƿ�����˲�����1-����˿���
  --���½������_In��0-�ڱ������и��£�1-�� zl_��Ա�ɿ����_Update �и���
  --ǿ������_In:0-��ǿ�ƣ�1-�����������ѿ����������ֵ�ǿ�����ֽ������
  --�Ƿ�ת��_In:0-ԭ���˻����֣�1-ת�˵�֧�ֵ���������
  v_Err_Msg Varchar2(200);
  Err_Item Exception;

  v_����     ���㷽ʽ.����%Type;
  v_��ӡid   Ʊ�ݴ�ӡ����.Id%Type;
  v_����     ������Ϣ.��������%Type;
  v_Date     Date;
  n_����ֵ   �������.Ԥ�����%Type;
  n_��id     ����ɿ����.Id%Type;
  n_������� �������.Ԥ�����%Type;
  n_����Ԥ�� �������.Ԥ�����%Type;
  n_�˿���        ����Ԥ����¼.���%Type;
  n_ʣ���          ����Ԥ����¼.���%Type;
  n_����id          ���˽��ʼ�¼.ID%Type;
  
  Cursor C_��Ԥ�� is
    Select A.NO, A.����id, A.Ԥ�����, A.�����id, A.����, A.������ˮ��, A.����˵��, 0 as ���, A.�տ�ʱ��, A.��� AS Ԥ����
    From ����Ԥ����¼ A Where RowNum < 2;
  r_��Ԥ�� C_��Ԥ��%Rowtype;
  
  Type Ty_ʣ��� Is Ref Cursor;
  C_ʣ��� Ty_ʣ���; --��̬�α���� 
Begin
  v_Date := �տ�ʱ��_In;
  If v_Date Is Null Then
    Select Sysdate Into v_Date From Dual;
  End If;
  n_��id := Zl_Get��id(����Ա����_In);

  --����Ԥ���ɿ��¼
  Insert Into ����Ԥ����¼
    (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ, �ɿ���id, Ԥ�����,
     �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
  Values
    (Id_In, ���ݺ�_In, Ʊ�ݺ�_In, 1, Decode(��������_In, 1, 0, 1), ����id_In, Decode(��ҳid_In, 0, Null, ��ҳid_In),
     Decode(����id_In, 0, Null, ����id_In), ���_In, ���㷽ʽ_In, �������_In, v_Date, �ɿλ_In, ��λ������_In, ��λ�ʺ�_In, ����Ա���_In, ����Ա����_In,
     ժҪ_In, n_��id, Ԥ�����_In, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, ������λ_In, ��������_In);

  If ��������_In = 1 Then
    --�ݲ�������ܱ�
    Return;
  Elsif ��������_In = 3 Then
    --����һ��ԭԤ��ID�ĳ�����¼��ͬʱҲ����һ������˿�ĳ�����¼
    Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
    IF Nvl(�����id_In, 0) = 0 And Nvl(���㿨���_In, 0) =0 then
      --���֣�������ͨ���㷽ʽ���֡�ǿ�����֡���������������
      Open C_ʣ��� For
           Select A.NO, A.����id, A.Ԥ�����, A.�����id, A.����, A.������ˮ��, A.����˵��, 
                   Min(decode(sign(A.���),-1,0,1)) AS ���, Min(decode(A.��¼����,1,A.�տ�ʱ��,null)) AS �տ�ʱ��,  
                   Nvl(Sum(A.���), 0) - Nvl(Sum(A.��Ԥ��), 0) as Ԥ����
              From ����Ԥ����¼ A, ҽ�ƿ���� B, �����ѽӿ�Ŀ¼ C
             Where A.����ID = ����id_In And A.��¼���� In (1,11) And A.Ԥ����� = Nvl(Ԥ�����_In, 2)
               And A.�����ID = B.ID(+) And Decode(ǿ������_In, 1, 1, Nvl(B.�Ƿ�����, 1)) = 1
               And A.�����ID = C.���(+) And Decode(ǿ������_In, 1, 1, Nvl(C.�Ƿ�����, 1)) = 1
             Group By A.NO, A.����id, A.Ԥ�����, A.�����id, A.����, A.������ˮ��, A.����˵��
            Having Nvl(Sum(A.���), 0) - Nvl(Sum(A.��Ԥ��), 0) <> 0
             Order By ���,�տ�ʱ��;
    ElsIF Nvl(�Ƿ�ת��_In, 0) = 1 Then
      --ת�ˣ��������������ֻ���ǿ�����֣�����Ŀ��ſ��ܲ���ԭ����,�����ͬ�ֿ�����Ԥ���ɿ��̯
      --Ŀǰֻ֧��ͬһ�ֿ�ת��
      Open C_ʣ��� For
           Select A.NO, A.����id, A.Ԥ�����, A.�����id, A.����, A.������ˮ��, A.����˵��, 
                   Min(decode(sign(A.���),-1,0,1)) AS ���, Min(decode(A.��¼����,1,A.�տ�ʱ��,null)) AS �տ�ʱ��,  
                   Nvl(Sum(A.���), 0) - Nvl(Sum(A.��Ԥ��), 0) as Ԥ����
              From ����Ԥ����¼ A, ҽ�ƿ���� B
             Where A.����ID = ����id_In And A.��¼���� In (1,11) And A.Ԥ����� = Nvl(Ԥ�����_In, 2)
               And A.�����ID = B.ID(+)
               And Nvl(�����id, 0) = Nvl(�����id_In, 0) And Nvl(������ˮ��, '-') = Nvl(������ˮ��_In, '-')
             Group By A.NO, A.����id, A.Ԥ�����, A.�����id, A.����, A.������ˮ��, A.����˵��
            Having Nvl(Sum(A.���), 0) - Nvl(Sum(A.��Ԥ��), 0) <> 0
             Order By ���,�տ�ʱ��;
    Else
      --�����������������ѿ������ݿ����ID�����㿨��š����š�������ˮ��ȱʡԭԤ����¼���������ȷ��Ψһ����з�̯
      Open C_ʣ��� For
           Select A.NO, A.����id, A.Ԥ�����, A.�����id, A.����, A.������ˮ��, A.����˵��, 
                   Min(decode(sign(A.���),-1,0,1)) AS ���, Min(decode(A.��¼����,1,A.�տ�ʱ��,null)) AS �տ�ʱ��,  
                   Nvl(Sum(A.���), 0) - Nvl(Sum(A.��Ԥ��), 0) as Ԥ����
              From ����Ԥ����¼ A
             Where A.����ID = ����id_In And A.��¼���� In (1,11) And A.Ԥ����� = Nvl(Ԥ�����_In, 2)
               And Nvl(A.�����id, 0) = Nvl(�����id_In, 0) And Nvl(A.���㿨���, 0) = Nvl(���㿨���_In, 0) 
               And Nvl(A.����, '-') = Nvl(����_In, '-') And Nvl(������ˮ��, '-') = Nvl(������ˮ��_In, '-')
             Group By A.NO, A.����id, A.Ԥ�����, A.�����id, A.����, A.������ˮ��, A.����˵��
            Having Nvl(Sum(A.���), 0) - Nvl(Sum(A.��Ԥ��), 0) <> 0
             Order By ���,�տ�ʱ��;
    End IF;
    
    n_ʣ��� := -1 * ���_In;
    n_�˿��� := 0;
    Loop
      Fetch C_ʣ���
        Into r_��Ԥ��;
      Exit When C_ʣ���%NotFound;
      IF r_��Ԥ��.NO <> ���ݺ�_In Then
        IF n_ʣ��� > r_��Ԥ��.Ԥ���� then
           n_�˿��� := r_��Ԥ��.Ԥ����;
           n_ʣ��� := n_ʣ��� - n_�˿���;
        Else
           n_�˿��� := n_ʣ���;
           n_ʣ��� := 0;
        End IF;
          	  
        IF nvl(n_�˿���, 0) <> 0 THEN 
          UPDATE ����Ԥ����¼  SET ����ID = n_����id WHERE NO = r_��Ԥ��.NO AND ��¼���� = 1 AND ����ID IS NULL;
          Insert Into ����Ԥ����¼
             (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, ���, ���㷽ʽ, �������, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���,
             �տ�ʱ��, ����Ա����, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, ��Ԥ��, ��������)
          Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, 1, ����id, ��ҳid, Ԥ�����, ����id, Null, ���㷽ʽ, �������, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���_In,
             v_Date, ����Ա����_In, ժҪ, n_��id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, n_����id, n_�˿���, NULL
          From ����Ԥ����¼
          Where NO = r_��Ԥ��.NO And ��¼���� In (1, 11) And RowNum < 2;
        END IF;

        IF n_ʣ��� = 0 Then 
          Exit;
        End IF;
      End IF;
    END LOOP;

    IF n_ʣ��� <> 0 And Nvl(�˿���_In, 0) = 1 THEN 
      v_Err_Msg := '�˿�����ڲ���ʣ��Ԥ����';
      Raise Err_Item;
    END IF;
    
    n_�˿��� := -1 * (-1 * ���_In - n_ʣ���);
    IF n_�˿��� <> 0 Then
      Update ����Ԥ����¼ Set ����id = n_����id Where ID = Id_In;
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ, �ɿ���id, Ԥ�����,
         �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, ��Ԥ��, ��������)
      Values
        (����Ԥ����¼_Id.Nextval, ���ݺ�_In, Ʊ�ݺ�_In, 11, 1, ����id_In, Decode(��ҳid_In, 0, Null, ��ҳid_In),
         Decode(����id_In, 0, Null, ����id_In), NULL, ���㷽ʽ_In, �������_In, v_Date, �ɿλ_In, ��λ������_In, ��λ�ʺ�_In, ����Ա���_In, ����Ա����_In,
         ժҪ_In, n_��id, Ԥ�����_In, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, ������λ_In, n_����id, n_�˿���, NULL);
    End IF;
  End If;

  --����Ʊ��
  If Ʊ�ݺ�_In Is Not Null Then
    Select Ʊ�ݴ�ӡ����_Id.Nextval Into v_��ӡid From Dual;
  
    --����Ʊ��
    Insert Into Ʊ�ݴ�ӡ���� (ID, ��������, NO) Values (v_��ӡid, 2, ���ݺ�_In);
  
    Insert Into Ʊ��ʹ����ϸ
      (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����)
    Values
      (Ʊ��ʹ����ϸ_Id.Nextval, 2, Ʊ�ݺ�_In, 1, 1, ����id_In, v_��ӡid, v_Date, ����Ա����_In);
  
    --״̬�Ķ�
    Update Ʊ�����ü�¼
    Set ��ǰ���� = Ʊ�ݺ�_In, ʣ������ = Decode(Sign(ʣ������ - 1), -1, 0, ʣ������ - 1), ʹ��ʱ�� = Sysdate
    Where ID = Nvl(����id_In, 0);
  End If;

  --��ػ��ܱ���

  --�������(Ԥ���������)
  Begin
    Select ���� Into v_���� From ���㷽ʽ Where ���� = ���㷽ʽ_In;
  Exception
    When Others Then
      Null;
  End;
  If Nvl(v_����, 1) <> 5 Then
    Update �������
    Set Ԥ����� = Nvl(Ԥ�����, 0) + ���_In
    Where ���� = 1 And ����id = ����id_In And Nvl(����, 0) = Nvl(Ԥ�����_In, 0)
    Returning Ԥ����� Into n_����ֵ;
    If Sql%RowCount = 0 Then
      Insert Into �������
        (����id, ����, ����, Ԥ�����, �������)
      Values
        (����id_In, 1, Nvl(Ԥ�����_In, 0), ���_In, 0);
      n_����ֵ := ���_In;
    End If;
    If Nvl(���_In, 0) = 0 Then
      Delete From ������� Where ����id = ����id_In And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
    End If;
  End If;

  If ���_In < 0 Then
    Begin
      Select Nvl(Ԥ�����, 0) - Nvl(�������, 0)
      Into n_�������
      From �������
      Where ���� = 1 And ����id = ����id_In And Nvl(����, 0) = Nvl(Ԥ�����_In, 0);
    Exception
      When Others Then
        Null;
    End;
    --����˿�Ҫ��������Ԥ���Ƿ�֧������
    If ��������_In = 3 And Nvl(ǿ������_In, 0) = 0 Then
      For c_����Ԥ�� In (Select a.Ԥ��id, a.Ԥ�����, a.�����id, a.���㿨��� As ���ѽӿ�id, Nvl(b.����, c.���) As ����, Nvl(b.����, c.����) As ����,
                            Decode(b.����, Null, c.�Ƿ�ȫ��, b.�Ƿ�ȫ��) As �Ƿ�ȫ��, Decode(b.����, Null, c.�Ƿ�����, b.�Ƿ�����) As �Ƿ�����, a.����,
                            a.������ˮ��, a.����˵��, a.Ԥ�����
                     From (Select a.Ԥ�����, Nvl(a.�����id, 0) As �����id, Nvl(a.���㿨���, 0) As ���㿨���, a.����, a.������ˮ��, a.����˵��,
                                   Max(Decode(Sign(���), -1, Decode(a.��¼״̬, 1, 0, 2, 0, ID), ID)) As Ԥ��id,
                                   Nvl(Sum(���), 0) - Nvl(Sum(Nvl(��Ԥ��, 0)), 0) As Ԥ�����
                            From ����Ԥ����¼ A
                            Where a.����id = ����id_In And (Nvl(a.���㿨���, 0) <> 0 Or Nvl(�����id, 0) <> 0)
                            Group By a.Ԥ�����, Nvl(a.�����id, 0), Nvl(a.���㿨���, 0), a.����, a.������ˮ��, a.����˵��
                            Having Nvl(Sum(���), 0) - Nvl(Sum(Nvl(��Ԥ��, 0)), 0) <> 0) A, ҽ�ƿ���� B, �����ѽӿ�Ŀ¼ C
                     Where a.Ԥ����� = Nvl(Ԥ�����_In, 0) And a.�����id = b.Id(+) And a.���㿨��� = c.���(+) And Nvl(a.Ԥ�����, 0) <> 0
                     Order By ����, a.����, a.������ˮ��, a.����˵��) Loop
      
        If Instr(',7,8,', ',' || v_���� || ',') = 0 And Nvl(c_����Ԥ��.�Ƿ�����, 0) = 0 And Nvl(c_����Ԥ��.Ԥ�����, 0) > 0 Then
          n_����Ԥ�� := Nvl(n_����Ԥ��, 0) + Nvl(c_����Ԥ��.Ԥ�����, 0);
        Elsif Instr(',7,8,', ',' || v_���� || ',') > 0 Then
          If Nvl(c_����Ԥ��.����, '0') = Nvl(����_In, '0') And Nvl(c_����Ԥ��.������ˮ��, '0') = Nvl(������ˮ��_In, '0') And
             Nvl(c_����Ԥ��.����˵��, '0') = Nvl(����˵��_In, '0') Then
            n_����Ԥ�� := Nvl(n_����Ԥ��, 0) + Nvl(c_����Ԥ��.Ԥ�����, 0);
          End If;
        End If;
      End Loop;
    End If;
  
    If Instr(',7,8,', ',' || v_���� || ',') > 0 And Nvl(n_����Ԥ��, 0) < 0 And ��������_In = 3 Then
      v_Err_Msg := '�˿�����ڲ�������Ԥ����';
      Raise Err_Item;
    Elsif Nvl(n_�������, 0) < 0 And �˿���_In = 1 Then
      v_Err_Msg := '�˿�����ڲ���ʣ��Ԥ����';
      Raise Err_Item;
    Elsif Instr(',7,8,', ',' || v_���� || ',') = 0 And Nvl(n_�������, 0) - Nvl(n_����Ԥ��, 0) < 0 And ��������_In = 3 And �˿���_In = 1 Then
      v_Err_Msg := '�˿�����ڲ���ʣ��Ԥ����';
      Raise Err_Item;
    End If;
  End If;

  --��Ա�ɿ����(����)
  If Nvl(���½������_In, 0) = 0 Then
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) + ���_In
    Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = ���㷽ʽ_In
    Returning ��� Into n_����ֵ;
  
    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ���� (�տ�Ա, ���㷽ʽ, ����, ���) Values (����Ա����_In, ���㷽ʽ_In, 1, ���_In);
      n_����ֵ := ���_In;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ��Ա�ɿ����
      Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = ���㷽ʽ_In And Nvl(���, 0) = 0;
    End If;
  End If;
  --����ʱ�����Ĵ���
  Select Nvl(��������, 0) Into v_���� From ������Ϣ Where ����id = ����id_In;
  If v_���� = 1 And Nvl(���_In, 0) > 0 Then
    Update ������Ϣ
    Set ������ = Decode(Sign(Nvl(������, 0) - Nvl(���_In, 0)), 1, Nvl(������, 0) - Nvl(���_In, 0), Null),
        ������ = Decode(Sign(Nvl(������, 0) - Nvl(���_In, 0)), 1, ������, Null),
        �������� = Decode(Sign(Nvl(������, 0) - Nvl(���_In, 0)), 1, ��������, Null)
    Where ����id = ����id_In;
  End If;
  If ��������_In <> 1 Then
    --��Ϣ����;
    Begin
      Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
        Using 11, Id_In;
    Exception
      When Others Then
        Null;
    End;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Ԥ����¼_Insert;
/

--127450:���ϴ�,2018-06-20,�ҺŰ��Ƚ��ȳ�ԭ��ʹ��Ԥ����
CREATE OR REPLACE Procedure Zl_���˹Һż�¼_Insert
( 
  ����id_In       ������ü�¼.����id%Type, 
  �����_In       ������ü�¼.��ʶ��%Type, 
  ����_In         ������ü�¼.����%Type, 
  �Ա�_In         ������ü�¼.�Ա�%Type, 
  ����_In         ������ü�¼.����%Type, 
  ���ʽ_In     ������ü�¼.���ʽ%Type, --���ڴ�Ų��˵�ҽ�Ƹ��ʽ��� 
  �ѱ�_In         ������ü�¼.�ѱ�%Type, 
  ���ݺ�_In       ������ü�¼.No%Type, 
  Ʊ�ݺ�_In       ������ü�¼.ʵ��Ʊ��%Type, 
  ���_In         ������ü�¼.���%Type, 
  �۸񸸺�_In     ������ü�¼.�۸񸸺�%Type, 
  ��������_In     ������ü�¼.��������%Type, 
  �շ����_In     ������ü�¼.�շ����%Type, 
  �շ�ϸĿid_In   ������ü�¼.�շ�ϸĿid%Type, 
  ����_In         ������ü�¼.����%Type, 
  ��׼����_In     ������ü�¼.��׼����%Type, 
  ������Ŀid_In   ������ü�¼.������Ŀid%Type, 
  �վݷ�Ŀ_In     ������ü�¼.�վݷ�Ŀ%Type, 
  ���㷽ʽ_In     ����Ԥ����¼.���㷽ʽ%Type, --�ֽ�Ľ������� 
  Ӧ�ս��_In     ������ü�¼.Ӧ�ս��%Type, 
  ʵ�ս��_In     ������ü�¼.ʵ�ս��%Type, 
  ���˿���id_In   ������ü�¼.���˿���id%Type, 
  ��������id_In   ������ü�¼.��������id%Type, 
  ִ�в���id_In   ������ü�¼.ִ�в���id%Type, 
  ����Ա���_In   ������ü�¼.����Ա���%Type, 
  ����Ա����_In   ������ü�¼.����Ա����%Type, 
  ����ʱ��_In     ������ü�¼.����ʱ��%Type, 
  �Ǽ�ʱ��_In     ������ü�¼.�Ǽ�ʱ��%Type, 
  ҽ������_In     �ҺŰ���.ҽ������%Type, 
  ҽ��id_In       �ҺŰ���.ҽ��id%Type, 
  ������_In       Number, --������¼�Ƿ��������� 
  ����_In         Number, 
  �ű�_In         �ҺŰ���.����%Type, 
  ����_In         ������ü�¼.��ҩ����%Type, 
  ����id_In       ������ü�¼.����id%Type, 
  ����id_In       Ʊ��ʹ����ϸ.����id%Type, 
  Ԥ��֧��_In     ����Ԥ����¼.��Ԥ��%Type, --ˢ���Һ�ʱʹ�õ�Ԥ�����,���Ϊ1����. 
  �ֽ�֧��_In     ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�ֽ�֧�����ݽ��,���Ϊ1����. 
  ����֧��_In     ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�����ʻ�֧�����,,���Ϊ1����. 
  ���մ���id_In   ������ü�¼.���մ���id%Type, 
  ������Ŀ��_In   ������ü�¼.������Ŀ��%Type, 
  ͳ����_In     ������ü�¼.ͳ����%Type, 
  ժҪ_In         ������ü�¼.ժҪ%Type, --ԤԼ�Һ�ժҪ��Ϣ 
  ԤԼ�Һ�_In     Number := 0, --ԤԼ�Һ�ʱ��(��¼״̬=0,����ʱ��ΪԤԼʱ��),��ʱ����Ҫ���������ز��� 
  �շ�Ʊ��_In     Number := 0, --�Һ��Ƿ�ʹ���շ�Ʊ�� 
  ���ձ���_In     ������ü�¼.���ձ���%Type, 
  ����_In         ���˹Һż�¼.����%Type := 0, 
  ����_In         �Һ����״̬.���%Type := Null, --ԤԼʱ������ü�¼�ķ�ҩ�����ֶ�,�Һ�ʱ����Һż�¼ 
  ����_In         ���˹Һż�¼.����%Type := Null, 
  ԤԼ����_In     Number := 0, 
  ԤԼ��ʽ_In     ԤԼ��ʽ.����%Type := Null, 
  ���ɶ���_In     Number := 0, 
  �����id_In     ����Ԥ����¼.�����id%Type := Null, 
  ���㿨���_In   ����Ԥ����¼.���㿨���%Type := Null, 
  ����_In         ����Ԥ����¼.����%Type := Null, 
  ������ˮ��_In   ����Ԥ����¼.������ˮ��%Type := Null, 
  ����˵��_In     ����Ԥ����¼.����˵��%Type := Null, 
  ������λ_In     ����Ԥ����¼.������λ%Type := Null, 
  ��������_In     Number := 0, 
  ����_In         ���˹Һż�¼.����%Type := Null, 
  ����ģʽ_In     Number := 0, 
  ���ʷ���_In     Number := 0, 
  �˺�����_In     Number := 1, 
  �������˷ѱ�_In Number := 0, 
  ���½������_In Number := 0, --�Ƿ������Ա��������Ҫ�Ǵ���ͳһ����Ա��¼��̨����������� 
  ������������_In Number := 0, 
  �շѵ�_In       ���˹Һż�¼.�շѵ�%Type := Null 
) As 
  --------------------------------------------------------------------------- 
  -- 
  --����: 
  --     ��������_in:0-�����ҺŻ���ԤԼ 1-����Աӵ�мӺ�Ȩ�޼Ӻ� 
  --     �������˷ѱ�_In:0-���޸Ĳ��˷ѱ� 1-�޸Ĳ��˷ѱ� 
  ---------------------------------------------------------------------------- 
  --���α������շѳ�Ԥ���Ŀ���Ԥ���б� 
  --��ID�������ȳ��ϴ�δ����ġ� 
  Cursor c_Deposit(v_����id ������Ϣ.����id%Type) Is 
    Select NO, Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) As ���, Min(��¼״̬) As ��¼״̬, Nvl(Max(����id), 0) As ����id, 
           Max(Decode(��¼����, 1, ID, 0) * Decode(��¼״̬, 2, 0, 1)) As ԭԤ��id, Min(Decode(��¼����, 1, �տ�ʱ��, NULL)) as �տ�ʱ��
    From ����Ԥ����¼ 
    Where ��¼���� In (1, 11) And ����id = v_����id And Nvl(Ԥ�����, 2) = 1 Having Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0 
    Group By NO 
    Order By �տ�ʱ��; 
  --���ܣ�����һ�в��˹Һŷ��ã�����������ܵ�����Ԥ����¼ 
  --       ͬʱ������صĻ��ܱ�(���˹ҺŻ��ܡ����û���) 
  --       ��һ�з��ô���Ʊ��ʹ�����(����ID_IN>0) 
 
  v_Err_Msg Varchar2(255); 
  Err_Item Exception; 
 
  v_�ŶӺ��� �ŶӽкŶ���.�ŶӺ���%Type; 
  v_�ֽ�     ���㷽ʽ.����%Type; 
  v_�����ʻ� ���㷽ʽ.����%Type; 
  v_�������� �ŶӽкŶ���.��������%Type; 
 
  n_��ʱ��       Number; 
  n_ʱ���޺�     Number; 
  n_ʱ����Լ     Number; 
  d_ʱ��ʱ��     Date; 
  d_������ʱ�� Date; 
  n_׷�Ӻ�       Number := 0; --����ʱ�ι��� ׷�ӹҺŵ���� 
  n_��Լ��       ���˹ҺŻ���.��Լ��%Type; 
  n_�ѽ���       ���˹ҺŻ���.�����ѽ���%Type; 
  n_ԤԼ��Чʱ�� Number; 
  n_ʧЧ��       Number; 
  n_ʧԼ�Һ�     Number := 0; 
  n_��������     Number; 
  n_����         Number := 0; 
 
  n_��ӡid   Ʊ�ݴ�ӡ����.Id%Type; 
  n_����id   ������ü�¼.Id%Type; 
  n_Ԥ����� ����Ԥ����¼.���%Type; 
  n_��ǰ��� ����Ԥ����¼.���%Type; 
  n_����ֵ   ����Ԥ����¼.���%Type; 
  n_Ԥ��id   ����Ԥ����¼.Id%Type; 
  n_���ѿ�id ���ѿ�Ŀ¼.Id%Type; 
  n_�Һ�id   ���˹Һż�¼.Id%Type; 
 
  n_��id           ����ɿ����.Id%Type; 
  n_���           �Һ����״̬.���%Type; 
  n_�������       �Һ����״̬.���%Type; 
  n_��ſ���       �ҺŰ���.��ſ���%Type; 
  n_����̨ǩ���Ŷ� Number; 
  n_Count          Number; 
  n_�޺���         Number(18); 
  n_���ƿ�         Number; 
  d_�Ŷ�ʱ��       Date; 
  d_���ʱ��       Date; 
  v_�ѱ�           �ѱ�.����%Type; 
  n_ʱ�����       Number := -1; 
  n_ԤԼ���ɶ���   Number; 
  n_����id         �ҺŰ���.Id%Type; 
  n_�ƻ�id         �ҺŰ��żƻ�.Id%Type := 0; 
  v_����           �ҺŰ�������.������Ŀ%Type; 
  n_��Լ��         Number(18); 
  n_�ѹ���         Number(4) := 0; 
 
  n_�ҳ��������� Number(4) := 0; 
  n_����ģʽ       ������Ϣ.����ģʽ%Type; 
  v_�Ŷ����       �ŶӽкŶ���.�Ŷ����%Type; 
  v_������         �Һ����״̬.������%Type; 
  v_��Ų���Ա     �Һ����״̬.����Ա����%Type; 
  v_��Ż�����     �Һ����״̬.������%Type; 
  v_���ʽ       ���˹Һż�¼.ҽ�Ƹ��ʽ%Type; 
  v_Temp           Varchar2(3000); 
  n_��ʱ����ʾ     Number(3); 
  n_����ʹ�÷�     Number(3); 
  d_����ʱ��       Date; 
Begin 
  --��ȡ��ǰ�������� 
  Select Terminal Into v_������ From V$session Where Audsid = Userenv('sessionid'); 
  n_��id := Zl_Get��id(����Ա����_In); 
  If �ѱ�_In Is Null Then 
    Begin 
      Select ���� Into v_�ѱ� From �ѱ� Where ȱʡ��־ = 1 And Rownum < 2; 
      Update ������Ϣ Set �ѱ� = v_�ѱ� Where ����id = ����id_In; 
    Exception 
      When Others Then 
        v_Err_Msg := '�޷�ȷ�����˷ѱ�����ȱʡ�ѱ��Ƿ���ȷ���ã�'; 
        Raise Err_Item; 
    End; 
  Else 
    v_�ѱ� := �ѱ�_In; 
    If Nvl(�������˷ѱ�_In, 0) = 1 Then 
      Begin 
        Update ������Ϣ Set �ѱ� = v_�ѱ� Where ����id = ����id_In; 
      Exception 
        When Others Then 
          v_Err_Msg := 'û���ҵ���Ӧ�Ĳ��ˣ�'; 
          Raise Err_Item; 
      End; 
    End If; 
  End If; 
 
  If Nvl(������������_In, 0) = 1 Then 
    Begin 
      Update ������Ϣ Set ���� = ����_In Where ����id = ����id_In; 
    Exception 
      When Others Then 
        v_Err_Msg := 'û���ҵ���Ӧ�Ĳ��ˣ�'; 
        Raise Err_Item; 
    End; 
  End If; 
 
  Begin 
    Delete From �Һ����״̬ 
    Where ���� = �ű�_In And ���� = ����ʱ��_In And ��� = ����_In And ״̬ = 3 And ����Ա���� = ����Ա����_In; 
  Exception 
    When Others Then 
      Null; 
  End; 
  v_Temp := zl_GetSysParameter(256); 
  If v_Temp Is Null Or Substr(v_Temp, 1, 1) = '0' Then 
    Null; 
  Else 
    Begin 
      d_����ʱ�� := To_Date(Substr(v_Temp, 3), 'YYYY-MM-DD hh24:mi:ss'); 
    Exception 
      When Others Then 
        d_����ʱ�� := Null; 
    End; 
    If d_����ʱ�� Is Not Null Then 
      If ����ʱ��_In > d_����ʱ�� Then 
        v_Err_Msg := '��ǰ�Һŵķ���ʱ��' || To_Char(����ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '�Ѿ������˳�����Ű�ģʽ,������ʹ�üƻ��Ű�ģʽ�Һ�!'; 
        Raise Err_Item; 
      End If; 
    End If; 
  End If; 
 
  If Nvl(ԤԼ�Һ�_In, 0) = 0 Then 
    --�ҺŻ���ԤԼ���� 
    --��Ϊ�����а��ձ൥�ݺŹ���,�չҺ������ܳ���10000��,����Ҫ���ΨһԼ���� 
    Select Count(*) 
    Into n_Count 
    From ������ü�¼ 
    Where ��¼���� = 4 And ��¼״̬ In (1, 3) And ��� = ���_In And NO = ���ݺ�_In; 
    If n_Count <> 0 Then 
      v_Err_Msg := '�Һŵ��ݺ��ظ�,���ܱ��棡' || Chr(13) || '���ʹ���˰���˳����,���չҺ������ܳ���10000�˴Ρ�'; 
      Raise Err_Item; 
    End If; 
 
    --��ȡ���㷽ʽ���� 
    Begin 
      Select ���� Into v_�ֽ� From ���㷽ʽ Where ���� = 1; 
    Exception 
      When Others Then 
        v_�ֽ� := '�ֽ�'; 
    End; 
    Begin 
      Select ���� Into v_�����ʻ� From ���㷽ʽ Where ���� = 3; 
    Exception 
      When Others Then 
        v_�����ʻ� := '�����ʻ�'; 
    End; 
  End If; 
 
  n_��� := ����_In; 
  Select Decode(To_Char(����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����', Null) 
  Into v_���� 
  From Dual; 
 
  --�ҺŻ�ȡ���� 
  Begin 
    Select a.Id, a.��ſ���, Nvl(b.�޺���, 0), Nvl(b.��Լ��, 0) 
    Into n_����id, n_��ſ���, n_�޺���, n_��Լ�� 
    From �ҺŰ��� A, �ҺŰ������� B 
    Where a.Id = b.����id(+) And b.������Ŀ(+) = v_���� And a.���� = �ű�_In; 
 
  Exception 
    When Others Then 
      n_����id := -1; 
  End; 
 
  --����ǲ����ѻ��ߺű�Ϊ��ʱ����� 
  If Nvl(������_In, 0) = 0 Or �ű�_In Is Not Null Then 
    If n_����id = -1 Then 
      v_Err_Msg := '��������Ӧ�ĹҺŰ�������,����'; 
      Raise Err_Item; 
    End If; 
  End If; 
 
  If Nvl(ԤԼ�Һ�_In, 0) = 1 Then 
    --���Ȼ�ȡ�ƻ� 
    Begin 
      Select ID 
      Into n_�ƻ�id 
      From �ҺŰ��żƻ� 
      Where ����id = n_����id And ���ʱ�� Is Not Null And 
            Nvl(��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) = 
            (Select Max(a.��Чʱ��) As ��Ч 
             From �ҺŰ��żƻ� A 
             Where a.���ʱ�� Is Not Null And ����ʱ��_In Between Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And 
                   Nvl(a.ʧЧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) And a.����id = n_����id) And 
            ����ʱ��_In Between Nvl(��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And 
            Nvl(ʧЧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')); 
 
    Exception 
      When Others Then 
        n_�ƻ�id := 0; 
    End; 
    If Nvl(n_�ƻ�id, 0) <> 0 Then 
      Begin 
        --��ȡ�ƻ������� 
        Select a.Id, a.��ſ���, Nvl(b.�޺���, 0) As �޺���, Nvl(b.��Լ��, 0) As ��Լ�� 
        Into n_�ƻ�id, n_��ſ���, n_�޺���, n_��Լ�� 
        From �ҺŰ��żƻ� A, �Һżƻ����� B 
        Where a.���� = �ű�_In And a.Id = n_�ƻ�id And a.���ʱ�� Is Not Null And a.Id = b.�ƻ�id(+) And b.������Ŀ(+) = v_����; 
      Exception 
        When Others Then 
          v_Err_Msg := '������Ӧ�ĹҺŰ��Ż�ƻ�����,����'; 
          Raise Err_Item; 
      End; 
    End If; 
  End If; 
 
  --��ȡ�Ƿ��ʱ�� 
  If Nvl(n_�ƻ�id, 0) = 0 Then 
    Select Count(Rownum) Into n_��ʱ�� From �ҺŰ���ʱ�� Where ���� = v_���� And ����id = n_����id And Rownum <= 1; 
  Else 
    Select Count(Rownum) Into n_��ʱ�� From �Һżƻ�ʱ�� Where ���� = v_���� And �ƻ�id = n_�ƻ�id And Rownum <= 1; 
  End If; 
 
  --��ʱ�ιҺ�ʱ�ж��Ƿ��ǹ��ڹҺ� Ҳ����׷�Ӻŵ���� Ŀǰֻ���ר�Һŷ�ʱ�ν��д��� 
  If Nvl(ԤԼ�Һ�_In, 0) = 0 And n_��ʱ�� > 0 And Nvl(n_��ſ���, 0) = 1 And ����_In Is Null And Nvl(��������_In, 0) = 0 Then 
    --����ʱ��_in>Sysdate ����ʱ��>����ʱ��ʱ��--����_in is null 
    Begin 
      Select Max(To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')) 
      Into d_������ʱ�� 
      From �ҺŰ���ʱ�� 
      Where ����id = n_����id And ���� = v_���� And Nvl(��������, 0) <> 0; 
      n_׷�Ӻ� := Case Sign(����ʱ��_In - d_������ʱ��) 
                 When -1 Then 
                  0 
                 Else 
                  1 
               End; 
    Exception 
      When Others Then 
        n_׷�Ӻ� := 0; 
    End; 
  End If; 
  d_ʱ��ʱ�� := ����ʱ��_In; 
 
  If ���_In = 1 And Nvl(ԤԼ�Һ�_In, 0) = 0 And n_��ʱ�� > 0 Then 
    --�Һ�ʱ��� �Ƿ����ʱ��,����ʱ��,��ʱ�ε�����������ȡ���� 
    Begin 
      Select Nvl(���, 0), 
             To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'), 
             ��������, Decode(Nvl(�Ƿ�ԤԼ, 0), 0, 0, ��������) 
      Into n_ʱ�����, d_ʱ��ʱ��, n_ʱ���޺�, n_ʱ����Լ 
      From �ҺŰ���ʱ�� 
      Where ����id = n_����id And ���� = v_���� And 
            (���, ����id, ����) In (Select Nvl(Max(���), -1), ����id, ���� 
                               From �ҺŰ���ʱ�� 
                               Where ����id = n_����id And ���� = v_���� And 
                                     Decode(��������_In + n_׷�Ӻ�, 0, To_Char(����ʱ��_In, 'hh24:mi'), 
                                            To_Char(Nvl(��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi')) = 
                                     To_Char(Nvl(��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi') 
                               Group By ����id, ����); 
    Exception 
      When Others Then 
        n_ʱ����� := -1; 
        n_��ʱ��   := 0; 
        d_ʱ��ʱ�� := ����ʱ��_In; 
        n_ʱ���޺� := 0; 
        n_ʱ����Լ := 0; 
    End; 
  End If; 
 
  If ���_In = 1 And Nvl(ԤԼ�Һ�_In, 0) = 1 And n_��ʱ�� > 0 Then 
    --ԤԼ��,ȡ�ƻ� 
    Begin 
      If Nvl(n_�ƻ�id, 0) = 0 Then 
        --û�ƻ���Ч,ȡ���ŵ����� 
        Select Nvl(���, 0), 
               To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(c.��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ʱ��ʱ��, 
               ��������, Decode(Nvl(�Ƿ�ԤԼ, 0), 0, 0, ��������) 
        Into n_ʱ�����, d_ʱ��ʱ��, n_ʱ���޺�, n_ʱ����Լ 
        From �ҺŰ���ʱ�� C 
        Where ����id = n_����id And ���� = v_���� And 
              (���, ����id, ����) In 
              (Select Nvl(Max(c.���), -1), ����id, ���� 
               From �ҺŰ���ʱ�� C 
               Where ����id = n_����id And c.���� = v_���� And 
                     Decode(��������_In, 1, To_Char(Nvl(c.��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi'), 
                            To_Char(����ʱ��_In, 'hh24:mi')) = 
                     To_Char(Nvl(c.��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi') 
               Group By ����id, ����); 
      Else 
        --�мƻ���Чȡ�ƻ� 
        --û��Ч�������ǴӹҺżƻ�ʱ�β�ѯ 
        Select Nvl(���, -1), 
               To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(c.��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ʱ��ʱ��, 
               ��������, Decode(Nvl(�Ƿ�ԤԼ, 0), 0, 0, ��������) 
        Into n_ʱ�����, d_ʱ��ʱ��, n_ʱ���޺�, n_ʱ����Լ 
        From �Һżƻ�ʱ�� C 
        Where �ƻ�id = n_�ƻ�id And ���� = v_���� And 
              (���, �ƻ�id, ����) In 
              (Select Nvl(Max(c.���), -1), �ƻ�id, ���� 
               From �Һżƻ�ʱ�� C 
               Where �ƻ�id = n_�ƻ�id And c.���� = v_���� And 
                     Decode(��������_In, 1, To_Char(Nvl(c.��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi'), 
                            To_Char(����ʱ��_In, 'hh24:mi')) = 
                     To_Char(Nvl(c.��ʼʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')), 'hh24:mi') 
               Group By �ƻ�id, ����); 
      End If; 
    Exception 
      When Others Then 
        n_ʱ����� := -1; 
        n_��ʱ��   := 0; 
        d_ʱ��ʱ�� := ����ʱ��_In; 
        n_ʱ���޺� := 0; 
        n_ʱ����Լ := 0; 
    End; 
  End If; 
 
  If ���_In = 1 Then 
 
    --��ȡ��ǰδʹ�õ���� 
    If Nvl(n_��ſ���, 0) = 1 And n_��ʱ�� = 0 Then 
      --<��ſ��� δ����ʱ�� ��ȡ���õ�������,�Լ��Ѿ�ʹ�õ�����> 
      Begin 
        --������ 
        If �˺�����_In = 1 Then 
          Select Max(Nvl(���, 0)), Sum(Decode(���, Nvl(����_In, 0), 0, 1)), Sum(Nvl(ԤԼ, 0)) 
          Into n_�������, n_��������, n_��Լ�� 
          From �Һ����״̬ 
          Where ���� = �ű�_In And ���� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ״̬ Not In (4, 5); 
        Else 
          Select Max(Nvl(���, 0)), Sum(Decode(���, Nvl(����_In, 0), 0, 1)), Sum(Nvl(ԤԼ, 0)) 
          Into n_�������, n_��������, n_��Լ�� 
          From �Һ����״̬ 
          Where ���� = �ű�_In And ���� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ״̬ <> 5; 
        End If; 
      Exception 
        When Others Then 
          n_������� := 0; 
          n_�������� := 0; 
      End; 
      If n_��� Is Null Then 
        n_��� := Nvl(n_�������, 0) + 1; 
      End If; 
      --<��ſ��� δ����ʱ�� ��ȡ���õ������� �Լ��Ѿ�ʹ�õ����� --end> 
 
      --�ǼӺŵ������Ҫ����Ƿ񳬹������� 
      If ��������_In = 0 Then 
        If Nvl(ԤԼ�Һ�_In, 0) = 0 Then 
          --�Һ�ʱ��� ���ǼӺŵ�����£���⵱ǰ�����Ƿ�ʹ�ã���ֹ��������µ�������ظ��Ŀ��ܡ�
          --������ſ���δ��ʱ�� �ﵽ������ 
          --���Һż�¼�е�ǰ����Ƿ��Ѿ�ʹ�ã���δʹ���򲻼��Һ�����
          Select Count(1)
          Into n_����ʹ�÷�
          From �Һ����״̬ 
          Where ���� = �ű�_In And ���= nvl(����_In,0) And ���� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ����Ա���� <> ����Ա����_In; 
          if nvl(����_In,0)>0 And n_����ʹ�÷�=1 then
              v_Err_Msg := '�ű�' || �ű�_In || '����' || to_char(n_����ʹ�÷�) || '��' || To_Char(Trunc(����ʱ��_In), 'yyyy-mm-dd ') || '�ѱ������û�ʹ�ã�'; 
              Raise Err_Item; 
          end if;
        Else 
          --ԤԼʱ��� 
          If n_��Լ�� = 0 Then 
            n_��Լ�� := n_�޺���; 
          End If; 
 
          If n_��Լ�� <= n_��Լ�� And n_��Լ�� > 0 Then 
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(Trunc(����ʱ��_In), 'yyyy-mm-dd') || '�Ѵﵽ�����Լ����'; 
            Raise Err_Item; 
          End If; 
        End If; 
 
      Else 
        Null; 
        --������ſ���,δ��ʱ�� �Ӻ����   ������,����Ժ������������Ժ󲹳� 
      End If; 
 
    Elsif Nvl(n_��ʱ��, 0) > 0 And Nvl(n_��ſ���, 0) = 0 Then 
      --<--��ͨ�ŷ�ʱ�� ����ֻ��ԤԼһ�����--> 
      If ��������_In = 0 Then 
        --<����ԤԼ�Һ�--> 
        Begin 
          Select Count(0) As �ѹ���, Nvl(Sum(Decode(Nvl(Sign(a.���� - d_ʱ��ʱ��), 0), 0, 1, 0)), 0) As ��Լ�� 
          Into n_�ѹ���, n_��Լ�� 
          From �Һ����״̬ A 
          Where a.���� = �ű�_In And ���� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And 
                ״̬ Not In (4, 5); 
        Exception 
          When Others Then 
            n_�ѹ��� := 0; 
            n_��Լ�� := 0; 
        End; 
 
        n_ʱ����Լ := n_ʱ���޺�; --��ͨ�ŷ�ʱ�ε����,29��n_ʱ����Լʼ����0 �������⴦�� 
        --����������� 
        If n_��Լ�� = 0 Then 
          n_��Լ�� := n_�޺���; 
        End If; 
        If n_ʱ����Լ <= n_��Լ�� Or n_��Լ�� <= n_��Լ�� Then 
          v_Err_Msg := '�ű�' || �ű�_In || '��ʱ��' || To_Char(d_ʱ��ʱ��, 'hh24:mi') || '�Ѵﵽ�����Լ����'; 
          Raise Err_Item; 
        End If; 
        If n_�޺��� <= n_�ѹ��� Then 
          v_Err_Msg := '�ű�' || �ű�_In || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������'; 
          Raise Err_Item; 
        End If; 
      End If; 
 
      --û�дﵽʱ�ε��޺��� �����ڵ�ǰʱ������׷�� 
 
      --��ȡ����ҳ��������� 
      If �˺�����_In = 1 Then 
        Select Nvl(Max(���), 0) 
        Into n_�ҳ��������� 
        From �Һ����״̬ A 
        Where a.���� = d_ʱ��ʱ�� And ���� = �ű�_In And ״̬ Not In (4, 5); 
      Else 
        Select Nvl(Max(���), 0) 
        Into n_�ҳ��������� 
        From �Һ����״̬ A 
        Where a.���� = d_ʱ��ʱ�� And ���� = �ű�_In And ״̬ <> 5; 
      End If; 
 
      --������� 
      n_��� := RPad(Nvl(n_ʱ�����, 0), Length(n_�޺���) + Length(Nvl(n_ʱ�����, 0)), 0) + n_��Լ�� + 1; 
      If n_��� <= Nvl(n_�ҳ���������, 0) Then 
        n_��� := Nvl(n_�ҳ���������, 0) + 1; 
      End If; 
 
      --<--��ͨ�ŷ�ʱ��--End> 
    Elsif Nvl(n_��ʱ��, 0) > 0 And Nvl(n_��ſ���, 0) = 1 Then 
      --<������ſ��� ����ʱ�� 
      --ר�Һŷ�ʱ�� 
      Begin 
        If �˺�����_In = 1 Then 
          Select Max(���), Sum(Decode(���, Nvl(����_In, 0), 0, 1)), Sum(Decode(Sign(���� - d_ʱ��ʱ��), 0, 1, 0)) 
          Into n_�������, n_�ѹ���, n_�������� 
          From �Һ����״̬ 
          Where ���� = �ű�_In And ���� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ״̬ Not In (4, 5); 
        Else 
          Select Max(���), Sum(Decode(���, Nvl(����_In, 0), 0, 1)), Sum(Decode(Sign(���� - d_ʱ��ʱ��), 0, 1, 0)) 
          Into n_�������, n_�ѹ���, n_�������� 
          From �Һ����״̬ 
          Where ���� = �ű�_In And ���� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ״̬ <> 5; 
        End If; 
      Exception 
        When Others Then 
          n_������� := 0; 
          n_�������� := 0; 
          n_�ѹ���   := 0; 
      End; 
 
      n_ʧЧ�� := 0; 
      If Nvl(ԤԼ�Һ�_In, 0) = 0 Then 
        n_ԤԼ��Чʱ�� := Zl_To_Number(zl_GetSysParameter('ԤԼ��Чʱ��', 1111)); 
        n_ʧԼ�Һ�     := Zl_To_Number(zl_GetSysParameter('ʧԼ���ڹҺ�', 1111)); 
        If Nvl(n_ԤԼ��Чʱ��, 0) <> 0 And Nvl(n_ʧԼ�Һ�, 0) > 0 Then 
          Begin 
            Select Sum(Decode(Sign((Sysdate - n_ԤԼ��Чʱ�� / 24 / 60) - ����), 1, 1, 0)) 
            Into n_ʧЧ�� 
            From �Һ����״̬ 
            Where ���� = �ű�_In And ���� Between Trunc(Sysdate) And Sysdate And Nvl(ԤԼ, 0) = 1 And ״̬ = 2; 
          Exception 
            When Others Then 
              n_ʧЧ�� := 0; 
          End; 
        End If; 
      End If; 
 
      If ��������_In = 0 Then 
        If Nvl(ԤԼ�Һ�_In, 0) = 0 Then 
 
          --�Һ� ׷�Ӻ���ʱ�����ʱ���޺��� 
          If n_ʱ���޺� <= n_�������� And Nvl(n_׷�Ӻ�, 0) = 0 Then 
            v_Err_Msg := '�ű�' || �ű�_In || '��ʱ��' || To_Char(d_ʱ��ʱ��, 'hh24:mi') || '�Ѵﵽ�����������'; 
            Raise Err_Item; 
          End If; 
          If n_�޺��� <= n_�ѹ��� - n_ʧЧ�� Then 
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������'; 
            Raise Err_Item; 
          End If; 
 
        Else 
          --�Һ� 
          If n_��Լ�� = 0 Then 
            n_��Լ�� := n_�޺���; 
          End If; 
          If n_��Լ�� <= n_�������� Then 
            v_Err_Msg := '�ű�' || �ű�_In || '��ʱ��' || To_Char(d_ʱ��ʱ��, 'hh24:mi') || '�Ѵﵽ�����������'; 
            Raise Err_Item; 
          End If; 
          If n_�޺��� <= n_�ѹ��� Then 
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������'; 
            Raise Err_Item; 
          End If; 
        End If; 
      End If; 
      If n_��� Is Null Then 
        --������� 
        If Nvl(n_�������, 0) < Nvl(n_ʱ�����, 0) Then 
          n_������� := Nvl(n_ʱ�����, 0); 
        End If; 
        n_��� := Nvl(n_�������, 0) + 1; 
      End If; 
    Elsif Nvl(n_��ʱ��, 0) = 0 And Nvl(n_��ſ���, 0) = 0 And Nvl(������_In, 0) = 0 And Nvl(�ű�_In, 0) > 0 Then 
      ---<--��ͨ��  --> 
      Begin 
        Select �ѹ���, ��Լ�� 
        Into n_��������, n_��Լ�� 
        From ���˹ҺŻ��� 
        Where ���� = Trunc(����ʱ��_In) And ���� = �ű�_In; 
      Exception 
        When Others Then 
          n_�������� := 0; 
          n_��Լ��   := 0; 
      End; 
      If Nvl(��������_In, 0) = 0 Then 
        If Nvl(ԤԼ�Һ�_In, 0) = 0 Then 
          --�Һ� 
          If Nvl(n_��������, 0) >= Nvl(n_�޺���, 0) And Nvl(n_�޺���, 0) > 0 Then 
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������'; 
            Raise Err_Item; 
          End If; 
        Else 
          --ԤԼ 
          If (Nvl(n_��Լ��, 0) > 0 And Nvl(n_��Լ��, 0) >= Nvl(n_��Լ��, 0)) Or 
             (Nvl(n_��������, 0) >= Nvl(n_�޺���, 0) And Nvl(n_�޺���, 0) > 0) Then 
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || '�Ѵﵽ�����������'; 
            Raise Err_Item; 
          End If; 
        End If; 
      End If; 
      ---<--��ͨ��  --> 
    End If; 
  End If; 
 
  --���¹Һ����״̬ 
  If ���_In = 1 And Not n_��� Is Null Then 
    If n_��ʱ�� = 1 Then 
      d_���ʱ�� := ����ʱ��_In; 
    Else 
      d_���ʱ�� := Trunc(����ʱ��_In); 
    End If; 
    --������ŵĴ��� 
    Begin 
      Select ����Ա����, ������ 
      Into v_��Ų���Ա, v_��Ż����� 
      From �Һ����״̬ 
      Where ״̬ = 5 And ���� = �ű�_In And ���� = d_���ʱ�� And ��� = n_���; 
      n_���� := 1; 
    Exception 
      When Others Then 
        v_��Ų���Ա := Null; 
        v_��Ż����� := Null; 
        n_����       := 0; 
    End; 
    If n_���� = 0 Then 
      Update �Һ����״̬ 
      Set ״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1), ԤԼ = Decode(ԤԼ����_In, 1, 1, 0), �Ǽ�ʱ�� = Sysdate 
      Where ���� = �ű�_In And ���� = d_���ʱ�� And ��� = n_��� And ״̬ = 3 And ����Ա���� = ����Ա����_In; 
      If Sql%RowCount = 0 Then 
        Begin 
          If Nvl(n_��ʱ��, 0) = 0 Or Nvl(ԤԼ�Һ�_In, 0) = 1 Or (Nvl(n_��ſ���, 0) = 0 And Nvl(����_In, 0) = 0) Then 
            Insert Into �Һ����״̬ 
              (����, ����, ���, ״̬, ����Ա����, ԤԼ, �Ǽ�ʱ��) 
            Values 
              (�ű�_In, d_���ʱ��, n_���, Decode(ԤԼ�Һ�_In, 1, 2, 1), ����Ա����_In, Decode(ԤԼ����_In, 1, 1, 0), Sysdate); 
            If Nvl(ԤԼ�Һ�_In, 0) = 0 And ԤԼ��ʽ_In Is Not Null Then 
              Update �Һ����״̬ Set ԤԼ = 1 Where ���� = �ű�_In And ���� = d_���ʱ�� And ��� = n_���; 
            End If; 
          Elsif Nvl(n_��ʱ��, 0) > 0 Then 
            --��ʱ�κ�ר�Һ� ʧԼ��ԤԼ������Һ� 
            Update �Һ����״̬ 
            Set ״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1), ����Ա���� = ����Ա����_In, ԤԼ = Decode(ԤԼ����_In, 1, 1, 0), �Ǽ�ʱ�� = Sysdate 
            Where ���� = �ű�_In And ���� = d_���ʱ�� And ��� = n_��� And ״̬ = 2; 
            If Sql%NotFound Then 
              Insert Into �Һ����״̬ 
                (����, ����, ���, ״̬, ����Ա����, ԤԼ, �Ǽ�ʱ��) 
              Values 
                (�ű�_In, d_���ʱ��, n_���, Decode(ԤԼ�Һ�_In, 1, 2, 1), ����Ա����_In, Decode(ԤԼ����_In, 1, 1, 0), Sysdate); 
            End If; 
            If Nvl(ԤԼ�Һ�_In, 0) = 0 And ԤԼ��ʽ_In Is Not Null Then 
              Update �Һ����״̬ Set ԤԼ = 1 Where ���� = �ű�_In And ���� = d_���ʱ�� And ��� = n_���; 
            End If; 
          End If; 
        Exception 
          When Others Then 
            v_Err_Msg := '���' || n_��� || '�ѱ�ʹ��,������ѡ��һ�����.'; 
            Raise Err_Item; 
        End; 
      End If; 
    Else 
      If ����Ա����_In <> v_��Ų���Ա Or v_������ <> v_��Ż����� Then 
        v_Err_Msg := '���' || n_��� || '�ѱ�������' || v_������ || '����,������ѡ��һ�����.'; 
        Raise Err_Item; 
      Else 
        Update �Һ����״̬ 
        Set ״̬ = Decode(ԤԼ�Һ�_In, 1, 2, 1), ԤԼ = Decode(ԤԼ����_In, 1, 1, 0), �Ǽ�ʱ�� = Sysdate 
        Where ���� = �ű�_In And ���� = d_���ʱ�� And ��� = n_��� And ״̬ = 5 And ����Ա���� = ����Ա����_In And ������ = v_������; 
        If Nvl(ԤԼ�Һ�_In, 0) = 0 And ԤԼ��ʽ_In Is Not Null Then 
          Update �Һ����״̬ Set ԤԼ = 1 Where ���� = �ű�_In And ���� = d_���ʱ�� And ��� = n_���; 
        End If; 
      End If; 
    End If; 
  End If; 
 
  --�������˹Һŷ���(���ܵ����ǻ������������) 
  Select ���˷��ü�¼_Id.Nextval Into n_����id From Dual; --Ӧ��ͨ������õ� 
 
  Insert Into ������ü�¼ 
    (ID, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, NO, ʵ��Ʊ��, �����־, �Ӱ��־, ���ӱ�־, ��ҩ����, ����id, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, ���˿���id, �շ����, 
     ���㵥λ, �շ�ϸĿid, ������Ŀid, �վݷ�Ŀ, ����, ����, ��׼����, Ӧ�ս��, ʵ�ս��, ���ʽ��, ����id, ���ʷ���, ��������id, ������, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, 
     ����ʱ��, �Ǽ�ʱ��, ���մ���id, ������Ŀ��, ���ձ���, ͳ����, ժҪ, ����, �ɿ���id) 
  Values 
    (n_����id, 4, Decode(ԤԼ�Һ�_In, 1, 0, 1), ���_In, Decode(�۸񸸺�_In, 0, Null, �۸񸸺�_In), ��������_In, ���ݺ�_In, Ʊ�ݺ�_In, 1, ����_In, 
     ������_In, Decode(ԤԼ�Һ�_In, 1, To_Char(n_���), ����_In), Decode(����id_In, 0, Null, ����id_In), 
     Decode(�����_In, 0, Null, �����_In), ���ʽ_In, ����_In, Decode(����_In, Null, Null, �Ա�_In), Decode(����_In, Null, Null, ����_In), 
     v_�ѱ�, ���˿���id_In, �շ����_In, �ű�_In, �շ�ϸĿid_In, ������Ŀid_In, �վݷ�Ŀ_In, 1, ����_In, ��׼����_In, Ӧ�ս��_In, ʵ�ս��_In, 
     Decode(ԤԼ�Һ�_In, 1, Null, Decode(Nvl(���ʷ���_In, 0), 1, Null, ʵ�ս��_In)), 
     Decode(ԤԼ�Һ�_In, 1, Null, Decode(Nvl(���ʷ���_In, 0), 1, Null, ����id_In)), Decode(Nvl(���ʷ���_In, 0), 1, 1, 0), ��������id_In, 
     ����Ա����_In, Decode(ԤԼ�Һ�_In, 1, ����Ա����_In, Null), ִ�в���id_In, ҽ������_In, ����Ա���_In, ����Ա����_In, ����ʱ��_In, �Ǽ�ʱ��_In, ���մ���id_In, 
     ������Ŀ��_In, ���ձ���_In, ͳ����_In, ժҪ_In, ԤԼ��ʽ_In, Decode(ԤԼ�Һ�_In, 1, Null, n_��id)); 
 
  --���ܽ��㵽����Ԥ����¼ 
  If Nvl(ԤԼ�Һ�_In, 0) = 0 And Nvl(���ʷ���_In, 0) = 0 Then 
 
    If (Nvl(�ֽ�֧��_In, 0) <> 0 Or (Nvl(�ֽ�֧��_In, 0) = 0 And Nvl(����֧��_In, 0) = 0 And Nvl(Ԥ��֧��_In, 0) = 0)) And ���_In = 1 Then 
      Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual; 
      Insert Into ����Ԥ����¼ 
        (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 
         ��������) 
      Values 
        (n_Ԥ��id, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), Nvl(���㷽ʽ_In, v_�ֽ�), Nvl(�ֽ�֧��_In, 0), �Ǽ�ʱ��_In, 
         ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�', n_��id, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, ������λ_In, 4); 
 
      If Nvl(���㿨���_In, 0) <> 0 And Nvl(�ֽ�֧��_In, 0) <> 0 Then 
 
        n_���ѿ�id := Null; 
        Begin 
          Select Nvl(���ƿ�, 0), 1 Into n_���ƿ�, n_Count From �����ѽӿ�Ŀ¼ Where ��� = ���㿨���_In; 
        Exception 
          When Others Then 
            n_Count := 0; 
        End; 
        If n_Count = 0 Then 
          v_Err_Msg := 'û�з���ԭ���㿨����Ӧ���,���ܼ���������'; 
          Raise Err_Item; 
        End If; 
        If n_���ƿ� = 1 Then 
          Select ID 
          Into n_���ѿ�id 
          From ���ѿ�Ŀ¼ 
          Where �ӿڱ�� = ���㿨���_In And ���� = ����_In And 
                ��� = (Select Max(���) From ���ѿ�Ŀ¼ Where �ӿڱ�� = ���㿨���_In And ���� = ����_In); 
        End If; 
        Zl_���˿������¼_Insert(���㿨���_In, n_���ѿ�id, ���㷽ʽ_In, �ֽ�֧��_In, ����_In, Null, �Ǽ�ʱ��_In, Null, ����id_In, n_Ԥ��id); 
      End If; 
 
    End If; 
 
    --����ҽ���Һ� 
    If Nvl(����֧��_In, 0) <> 0 And ���_In = 1 Then 
      Insert Into ����Ԥ����¼ 
        (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, ��������) 
      Values 
        (����Ԥ����¼_Id.Nextval, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), v_�����ʻ�, ����֧��_In, �Ǽ�ʱ��_In, ����Ա���_In, 
         ����Ա����_In, ����id_In, 'ҽ���Һ�', n_��id, 4); 
    End If; 
 
    --���ھ��￨ͨ��Ԥ����Һ� 
    If Nvl(Ԥ��֧��_In, 0) <> 0 And ���_In = 1 Then 
      n_Ԥ����� := Ԥ��֧��_In; 
      For r_Deposit In c_Deposit(����id_In) Loop 
        n_��ǰ��� := Case 
                    When r_Deposit.��� - n_Ԥ����� < 0 Then 
                     r_Deposit.��� 
                    Else 
                     n_Ԥ����� 
                  End; 
 
        If r_Deposit.����id = 0 Then 
          --��һ�γ�Ԥ��(82592,����һ�α��Ͻ���ID,��Ԥ�����Ϊ0) 
          Update ����Ԥ����¼ Set ��Ԥ�� = 0, ����id = ����id_In, �������� = 4 Where ID = r_Deposit.ԭԤ��id; 
        End If; 
        --���ϴ�ʣ��� 
        Insert Into ����Ԥ����¼ 
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, 
           ��Ԥ��, ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������) 
          Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, 
                 �Ǽ�ʱ��_In, ����Ա����_In, ����Ա���_In, n_��ǰ���, ����id_In, n_��id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4 
          From ����Ԥ����¼ 
          Where NO = r_Deposit.No And ��¼״̬ = r_Deposit.��¼״̬ And ��¼���� In (1, 11) And Rownum = 1; 
 
        --���²���Ԥ����� 
        Update ������� 
        Set Ԥ����� = Nvl(Ԥ�����, 0) - n_��ǰ��� 
        Where ����id = ����id_In And ���� = 1 And ���� = Nvl(1, 2); 
        --����Ƿ��Ѿ������� 
        If r_Deposit.��� < n_Ԥ����� Then 
          n_Ԥ����� := n_Ԥ����� - r_Deposit.���; 
        Else 
          n_Ԥ����� := 0; 
        End If; 
 
        If n_Ԥ����� = 0 Then 
          Exit; 
        End If; 
      End Loop; 
      If n_Ԥ����� > 0 Then 
        v_Err_Msg := 'Ԥ���಻��֧������֧�����,���ܼ���������'; 
        Raise Err_Item; 
      End If; 
      Delete From ������� Where ����id = ����id_In And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0; 
    End If; 
 
    --��ػ��ܱ�Ĵ��� 
    --��Ա�ɿ���� 
    If ���_In = 1 And Nvl(���½������_In, 0) = 0 Then 
      If Nvl(�ֽ�֧��_In, 0) <> 0 Then 
        Update ��Ա�ɿ���� 
        Set ��� = Nvl(���, 0) + �ֽ�֧��_In 
        Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = Nvl(���㷽ʽ_In, v_�ֽ�) 
        Returning ��� Into n_����ֵ; 
 
        If Sql%RowCount = 0 Then 
          Insert Into ��Ա�ɿ���� 
            (�տ�Ա, ���㷽ʽ, ����, ���) 
          Values 
            (����Ա����_In, Nvl(���㷽ʽ_In, v_�ֽ�), 1, �ֽ�֧��_In); 
          n_����ֵ := �ֽ�֧��_In; 
        End If; 
        If Nvl(n_����ֵ, 0) = 0 Then 
          Delete From ��Ա�ɿ���� 
          Where �տ�Ա = ����Ա����_In And ���㷽ʽ = Nvl(���㷽ʽ_In, v_�ֽ�) And ���� = 1 And Nvl(���, 0) = 0; 
        End If; 
      End If; 
 
      If Nvl(����֧��_In, 0) <> 0 Then 
        Update ��Ա�ɿ���� 
        Set ��� = Nvl(���, 0) + ����֧��_In 
        Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = v_�����ʻ� 
        Returning ��� Into n_����ֵ; 
 
        If Sql%RowCount = 0 Then 
          Insert Into ��Ա�ɿ���� (�տ�Ա, ���㷽ʽ, ����, ���) Values (����Ա����_In, v_�����ʻ�, 1, ����֧��_In); 
          n_����ֵ := ����֧��_In; 
        End If; 
        If Nvl(n_����ֵ, 0) = 0 Then 
          Delete From ��Ա�ɿ���� 
          Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_�����ʻ� And Nvl(���, 0) = 0; 
        End If; 
      End If; 
    End If; 
  End If; 
 
  --���˹ҺŻ���(ֻ����һ��,�ҵ�����ȡ�����Ѳ�����) 
  If Nvl(ԤԼ�Һ�_In, 0) = 0 Then 
    --���˱��ξ���(�Է���ʱ��Ϊ׼) 
    If Nvl(����id_In, 0) <> 0 And ���_In = 1 Then 
      Update ������Ϣ Set ����ʱ�� = ����ʱ��_In, ����״̬ = 1, �������� = ����_In Where ����id = ����id_In; 
    End If; 
  End If; 
 
  If Nvl(ԤԼ�Һ�_In, 0) = 0 And Nvl(���ʷ���_In, 0) = 1 Then 
    --���� 
    If Nvl(����id_In, 0) = 0 Then 
      v_Err_Msg := 'Ҫ��Բ��˵ĹҺŷѽ��м��ʣ������ǽ������˲��ܼ��ʹҺš�'; 
      Raise Err_Item; 
    End If; 
 
    --������� 
    Update ������� 
    Set ������� = Nvl(�������, 0) + Nvl(ʵ�ս��_In, 0) 
    Where ����id = Nvl(����id_In, 0) And ���� = 1 And ���� = 1; 
 
    If Sql%RowCount = 0 Then 
      Insert Into ������� (����id, ����, ����, �������, Ԥ�����) Values (����id_In, 1, 1, Nvl(ʵ�ս��_In, 0), 0); 
    End If; 
 
    --����δ����� 
    Update ����δ����� 
    Set ��� = Nvl(���, 0) + Nvl(ʵ�ս��_In, 0) 
    Where ����id = ����id_In And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���˿���id, 0) = Nvl(���˿���id_In, 0) And 
          Nvl(��������id, 0) = Nvl(��������id_In, 0) And Nvl(ִ�в���id, 0) = Nvl(ִ�в���id_In, 0) And ������Ŀid + 0 = ������Ŀid_In And 
          ��Դ;�� + 0 = 1; 
 
    If Sql%RowCount = 0 Then 
      Insert Into ����δ����� 
        (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���) 
      Values 
        (����id_In, Null, Null, ���˿���id_In, ��������id_In, ִ�в���id_In, ������Ŀid_In, 1, Nvl(ʵ�ս��_In, 0)); 
    End If; 
  End If; 
 
  --���˹Һż�¼ 
  If �ű�_In Is Not Null And ���_In = 1 Then 
    --And Nvl(ԤԼ�Һ�_In, 0) = 0 
    Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual; 
    Begin 
      Select ���� Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = ���ʽ_In And Rownum < 2; 
    Exception 
      When Others Then 
        v_���ʽ := Null; 
    End; 
    Insert Into ���˹Һż�¼ 
      (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ԤԼʱ��, ����Ա���, 
       ����Ա����, ����, ����, ����, ԤԼ, ԤԼ��ʽ, ժҪ, ������ˮ��, ����˵��, ������λ, ����ʱ��, ������, ԤԼ����Ա, ԤԼ����Ա���, ����, ҽ�Ƹ��ʽ, �շѵ�) 
    Values 
      (n_�Һ�id, ���ݺ�_In, Decode(Nvl(ԤԼ�Һ�_In, 0), 1, 2, 1), 1, ����id_In, �����_In, ����_In, �Ա�_In, ����_In, �ű�_In, ����_In, ����_In, 
       Null, ִ�в���id_In, ҽ������_In, 0, Null, �Ǽ�ʱ��_In, ����ʱ��_In, Decode(Nvl(ԤԼ�Һ�_In, 0), 1, ����ʱ��_In, Null), ����Ա���_In, 
       ����Ա����_In, ����_In, n_���, ����_In, Decode(ԤԼ����_In, 1, 1, 0), ԤԼ��ʽ_In, ժҪ_In, ������ˮ��_In, ����˵��_In, ������λ_In, 
       Decode(Nvl(ԤԼ�Һ�_In, 0), 0, �Ǽ�ʱ��_In, Null), Decode(Nvl(ԤԼ�Һ�_In, 0), 0, ����Ա����_In, Null), 
       Decode(Nvl(ԤԼ�Һ�_In, 0), 1, ����Ա����_In, Null), Decode(Nvl(ԤԼ�Һ�_In, 0), 1, ����Ա���_In, Null), 
       Decode(Nvl(����_In, 0), 0, Null, ����_In), v_���ʽ, �շѵ�_In); 
    If Nvl(ԤԼ�Һ�_In, 0) = 0 And ԤԼ��ʽ_In Is Not Null Then 
      Update ���˹Һż�¼ 
      Set ԤԼ = 1, ԤԼʱ�� = ����ʱ��_In, ԤԼ����Ա = ����Ա����_In, ԤԼ����Ա��� = ����Ա���_In 
      Where ID = n_�Һ�id; 
    End If; 
    n_ԤԼ���ɶ��� := 0; 
    If Nvl(ԤԼ�Һ�_In, 0) = 1 Then 
      n_ԤԼ���ɶ��� := Zl_To_Number(zl_GetSysParameter('ԤԼ���ɶ���', 1113)); 
    End If; 
 
    --0-����������;1-��ҽ�������̨�Ŷ�;2-�ȷ���,��ҽ��վ 
    If Nvl(���ɶ���_In, 0) <> 0 And Nvl(ԤԼ�Һ�_In, 0) = 0 Or n_ԤԼ���ɶ��� = 1 Then 
      n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113)); 
      If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Or n_ԤԼ���ɶ��� = 1 Then 
        n_��ʱ����ʾ := Nvl(Zl_To_Number(zl_GetSysParameter(270)), 0); 
        If Nvl(ԤԼ�Һ�_In, 0) = 1 And n_��ʱ����ʾ = 1 And n_��ʱ�� = 1 Then 
          n_��ʱ����ʾ := 1; 
        Else 
          n_��ʱ����ʾ := Null; 
        End If; 
        --�������� 
        --.����ִ�в��š� �ķ�ʽ���ɶ��� 
        v_�������� := ִ�в���id_In; 
        v_�ŶӺ��� := Zlgetnextqueue(ִ�в���id_In, n_�Һ�id, �ű�_In || '|' || n_���); 
        v_�Ŷ���� := Zlgetsequencenum(0, n_�Һ�id, 0); 
        d_�Ŷ�ʱ�� := Zl_Get_Queuedate(n_�Һ�id, �ű�_In, n_���, ����ʱ��_In); 
        --  ��������_In , ҵ������_In, ҵ��id_In,����id_In,�ŶӺ���_In,�Ŷӱ��_In,��������_In,����ID_IN, ����_In, ҽ������_In, �Ŷ�ʱ��_In 
        Zl_�ŶӽкŶ���_Insert(v_��������, 0, n_�Һ�id, ִ�в���id_In, v_�ŶӺ���, Null, ����_In, ����id_In, ����_In, ҽ������_In, d_�Ŷ�ʱ��, ԤԼ��ʽ_In, 
                         n_��ʱ����ʾ, v_�Ŷ����); 
      End If; 
    End If; 
  End If; 
  --���˵�����Ϣ 
  If ����id_In Is Not Null And ���_In = 1 Then 
    --ȡ����: 
    If Nvl(n_����ģʽ, 0) <> Nvl(����ģʽ_In, 0) Then 
      --����ģʽ��ȷ�� 
 
      v_Err_Msg := Null; 
      Begin 
        Select Nvl(����ģʽ, 0) Into n_����ģʽ From ������Ϣ Where ����id = ����id_In; 
      Exception 
        When Others Then 
          v_Err_Msg := 'δ�ҵ�ָ���Ĳ�����Ϣ,������Һ�'; 
      End; 
 
      If v_Err_Msg Is Not Null Then 
        Raise Err_Item; 
      End If; 
      If n_����ģʽ = 1 And Nvl(����ģʽ_In, 0) = 0 Then 
        --�����Ѿ���"�����ƺ�����",������"�Ƚ�������Ƶ�",�����Ƿ����δ������ 
        Select Count(1) 
        Into n_Count 
        From ����δ����� 
        Where ����id = ����id_In And (��Դ;�� In (1, 4) Or ��Դ;�� = 3 And Nvl(��ҳid, 0) = 0) And Nvl(���, 0) <> 0 And Rownum < 2; 
        If Nvl(n_Count, 0) <> 0 Then 
          --����δ�������ݣ������Ƚ���������ִ�� 
          v_Err_Msg := '��ǰ���˵ľ���ģʽΪ�����ƺ�����Ҵ���δ����ã�����������ò��˵ľ���ģʽ,������ȶ�δ����ý��ʺ��ٹҺŻ򲻵������˵ľ���ģʽ!'; 
          Raise Err_Item; 
        End If; 
        --��� 
        --δ����ҽ��ҵ��ģ�����ʱ�͹Һŵ�,��Ҫ��֤ͬһ�εľ���ģʽ��һ����(�����Ѿ���飬�����ٴ���) 
      End If; 
      Update ������Ϣ Set ����ģʽ = ����ģʽ_In Where ����id = ����id_In; 
    End If; 
 
    Update ������Ϣ 
    Set ������ = Null, ������ = Null, �������� = Null 
    Where ����id = ����id_In And Nvl(��Ժ, 0) = 0 And Exists 
     (Select 1 
           From ���˵�����¼ 
           Where ����id = ����id_In And ��ҳid Is Not Null And 
                 �Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From ���˵�����¼ Where ����id = ����id_In)); 
 
    If Sql%RowCount > 0 Then 
      Update ���˵�����¼ 
      Set ����ʱ�� = Sysdate 
      Where ����id = ����id_In And ��ҳid Is Not Null And Nvl(����ʱ��, Sysdate) >= Sysdate; 
    End If; 
  End If; 
  If ���_In = 1 Then 
    --��Ϣ���� 
    Begin 
      Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;' 
        Using 1, n_�Һ�id; 
    Exception 
      When Others Then 
        Null; 
    End; 
  End If; 
Exception 
  When Err_Item Then 
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]'); 
  When Others Then 
    zl_ErrorCenter(SQLCode, SQLErrM); 
End Zl_���˹Һż�¼_Insert;
/

--127480:����,2018-06-20,����Oracle����Zl_�ҺŰ���_Autoupdate
CREATE OR REPLACE Procedure Zl_�ҺŰ���_Autoupdate Is
  Err_Item Exception;
  v_Date Date;
  -- v_Err_Msg Varchar2(100);
  v_Unitscount Number;
Begin
  --n_����ִ���� ���Ƿ���²��˹Һż�¼ ��������ü�¼�е�ִ����
  --               ����ƻ��и����� �Һ���Ŀ ��������� ���˹Һż�¼��������ü�¼�е�����
  Select Sysdate Into v_Date From Dual;
  Select Count(0) Into v_Unitscount From ������λ���ſ��� Where Rownum = 1;

  For v_��Ч In (Select ID, ����id, ����, ��Чʱ��, ʧЧʱ��, ����, ��һ, �ܶ�, ����, ����, ����, ����, ���﷽ʽ, ��ſ���, ִ��ʱ�� As �ϴ���Чʱ��, ��Ŀid, ҽ������, ҽ��id,
                      ���, ����id, �Ƿ���ͬ
               From (Select a.Id, a.����id, a.����, a.��Чʱ��, a.ʧЧʱ��, a.����, a.��һ, a.�ܶ�, a.����, a.����, a.����, a.����, a.���﷽ʽ, a.��ſ���,
                             b.ִ��ʱ��, a.��Ŀid, a.ҽ������, a.ҽ��id, Nvl(b.ִ�мƻ�id, 0) As ִ�мƻ�id,
                             Row_Number() Over(Partition By a.����id Order By a.��Чʱ�� Desc) As ˳���, b.���, b.����id,
                             Case
                               When b.��Ŀid = a.��Ŀid And Nvl(a.ҽ��id, 0) = Nvl(b.ҽ��id, 0) And
                                    Nvl(a.ҽ������, '-') = Nvl(b.ҽ������, '-') Then
                                1
                               Else
                                0
                             End As �Ƿ���ͬ
                      From �ҺŰ��żƻ� A, �ҺŰ��� B
                      Where Sysdate Between a.��Чʱ�� And a.ʧЧʱ�� And a.����id = b.Id And
                            a.ʵ����Ч >= To_Date('3000-01-01', 'yyyy-mm-dd') And a.��Чʱ�� + 0 <= Sysdate And ����� Is Not Null And
                            b.ͣ������ Is Null)
               Where ˳��� = 1 And ID <> Nvl(ִ�мƻ�id, 0)) Loop
    Update �ҺŰ��żƻ�
    Set ʵ����Ч = v_��Ч.�ϴ���Чʱ��
    Where ����id = v_��Ч.����id And ʧЧʱ�� <= v_��Ч.ʧЧʱ�� And ��Чʱ�� < Sysdate And ID <> v_��Ч.Id And
          ʵ����Ч >= To_Date('3000-01-01', 'yyyy-mm-dd');
  
    Update �ҺŰ���
    Set ���� = v_��Ч.����, ��һ = v_��Ч.��һ, �ܶ� = v_��Ч.�ܶ�, ���� = v_��Ч.����, ���� = v_��Ч.����, ���� = v_��Ч.����, ���� = v_��Ч.����,
        ���﷽ʽ = v_��Ч.���﷽ʽ, ��ſ��� = v_��Ч.��ſ���, ��ʼʱ�� = Sysdate, ��ֹʱ�� = v_��Ч.ʧЧʱ��, ��Ŀid = Nvl(v_��Ч.��Ŀid, ��Ŀid), ִ��ʱ�� = v_Date,
        ִ�мƻ�id = v_��Ч.Id, ��� = Decode(v_��Ч.�Ƿ���ͬ, 1, ���, 9999999), ҽ������ = v_��Ч.ҽ������, ҽ��id = v_��Ч.ҽ��id
    Where ID = v_��Ч.����id;
  
    --���µ������
    If Nvl(v_��Ч.�Ƿ���ͬ, 0) <> 1 Then
    
      Update �ҺŰ��� A
      Set ��� = -1 * ���
      Where ��Ŀid = v_��Ч.��Ŀid And a.����id = v_��Ч.����id And Nvl(a.ҽ������, '-') = Nvl(v_��Ч.ҽ������, '-') And
            Nvl(a.ҽ��id, 0) = Nvl(v_��Ч.ҽ��id, 0);
      For v_��� In (Select a.Id, Rownum As ���
                   From �ҺŰ��� A
                   Where a.��Ŀid = v_��Ч.��Ŀid And a.����id = v_��Ч.����id And Nvl(a.ҽ������, '-') = Nvl(v_��Ч.ҽ������, '-') And
                         Nvl(a.ҽ��id, 0) = Nvl(v_��Ч.ҽ��id, 0)
                   Order By a.Id) Loop
        Update �ҺŰ��� A Set ��� = v_���.��� Where ID = v_���.Id;
      End Loop;
    End If;
    Delete �ҺŰ������� Where �ű�id = v_��Ч.����id;
    Insert Into �ҺŰ�������
      (�ű�id, ��������)
      Select v_��Ч.����id, �������� From �Һżƻ����� Where �ƻ�id = v_��Ч.Id;
    Delete �ҺŰ������� Where ����id = v_��Ч.����id;
    Insert Into �ҺŰ�������
      (����id, ������Ŀ, �޺���, ��Լ��)
      Select v_��Ч.����id, ������Ŀ, �޺���, ��Լ�� From �Һżƻ����� Where �ƻ�id = v_��Ч.Id;
    Delete �ҺŰ���ʱ�� Where ����id = v_��Ч.����id;
    Insert Into �ҺŰ���ʱ��
      (����id, ���, ��ʼʱ��, ����ʱ��, ��������, �Ƿ�ԤԼ, ����)
      Select v_��Ч.����id, ���, ��ʼʱ��, ����ʱ��, ��������, �Ƿ�ԤԼ, ����
      From �Һżƻ�ʱ��
      Where �ƻ�id = v_��Ч.Id;
    If Nvl(v_Unitscount, 0) > 0 Then
      Delete ������λ���ſ��� Where ����id = v_��Ч.����id;
      Insert Into ������λ���ſ���
        (����id, ������λ, ������Ŀ, ���, ����)
        Select v_��Ч.����id, ������λ, ������Ŀ, ���, ���� From ������λ�ƻ����� Where �ƻ�id = v_��Ч.Id;
    End If;
  End Loop;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ҺŰ���_Autoupdate;
/

--126104:����,2018-06-14,���µ�סԺ��������һ�����µ���ʼʱ�����
Create Or Replace Function Zl_Calcindaysnew
(
  �ļ�id_In   In ���˻����ļ�.Id%Type,
  ����id_In   In ������ҳ.����id%Type,
  ��ҳid_In   In ������ҳ.��ҳid%Type,
  סԺ����_In In Date := Sysdate --ָ��ĳһ��������ʱ��סԺ���� 
) Return Number As

  d_��Ժʱ�� ������ҳ.��Ժ����%Type;
  d_��Ժʱ�� ������ҳ.��Ժ����%Type;
  d_��ʼʱ�� ���˻����ļ�.��ʼʱ��%Type;
  n_Days     Number(18);
  n_Bady     Number(18);
  n_Badybill Number(18);
  n_Addday   Number(18);
Begin

  n_Days     := 0;
  n_Bady     := 0;
  n_Badybill := 0;
  n_Addday   := 1;
  d_��Ժʱ�� := Null;
  d_��Ժʱ�� := Null;
  d_��ʼʱ�� := Null;
  --��ȡ���µ���ʼʱ�� 
  Begin
    Select Ӥ�� Into n_Bady From ���˻����ļ� Where ID = �ļ�id_In;
  Exception
    When Others Then
      n_Bady := 0;
  End;
  --��ȡ��һ�����µ��Ŀ�ʼʱ��,106122-CL-07-02-21 
  Begin
    Select ��ʼʱ��
    Into d_��ʼʱ��
    From (Select ��ʼʱ��
           From ���˻����ļ� A, �����ļ��б� B
           Where ����id = ����id_In And ��ҳid = ��ҳid_In And a.��ʽid = b.Id And b.���� = -1
           Order By ��ʼʱ��)
    Where Rownum < 2;
  Exception
    When Others Then
      d_��ʼʱ�� := Null;
  End;

  --�����Ӥ����ʼʱ���Գ���ʱ��Ϊ׼ 
  If n_Bady <> 0 Then
    Begin
      Select a.����ʱ��
      Into d_��Ժʱ��
      From ������������¼ A
      Where a.����id = ����id_In And a.��ҳid = ��ҳid_In And a.��� = n_Bady;
    Exception
      When Others Then
        d_��Ժʱ�� := d_��ʼʱ��;
    End;
    Begin
      Select Nvl(����ֵ, 0) Into n_Badybill From zlParameters Where ģ�� = 1255 And ������ = 'Ӥ�����µ�����������ʾ0'; --Ӥ�����µ�����������0��ʼ���Ǵ�1��ʼ 
    Exception
      When Others Then
        n_Badybill := 0;
    End;
  End If;

  If d_��Ժʱ�� Is Null Then
    d_��Ժʱ�� := d_��ʼʱ��;
  End If;

  --��ȡ���µ���ʵ�ʽ���ʱ�� 
  Begin
    Select Decode(Sign(a.��Ժʱ�� - b.����ʱ��), 1, a.��Ժʱ��, b.����ʱ��)
    Into d_��Ժʱ��
    From (Select Max(Nvl(��ֹʱ��, Sysdate)) As ��Ժʱ��, Max(����id) ����id, Max(��ҳid) ��ҳid
           From ���˱䶯��¼
           Where ��ʼʱ�� Is Not Null And ����id = ����id_In And ��ҳid = ��ҳid_In) A,
         (Select Nvl(����ʱ��, Sysdate) ����ʱ��, ����id, ��ҳid
           From (Select Max(����ʱ��) ����ʱ��, Max(a.����id) ����id, Max(a.��ҳid) ��ҳid
                  From ���˻����ļ� A, ���˻������� B
                  Where a.Id = b.�ļ�id And a.Id = �ļ�id_In)) B
    Where a.����id = b.����id And a.��ҳid = b.��ҳid;
  
  Exception
    When Others Then
      d_��Ժʱ�� := Sysdate;
  End;

  If n_Badybill = 1 Then
    n_Addday := 0;
  Else
    n_Addday := 1;
  End If;

  If d_��Ժʱ�� Is Not Null Then
    If Trunc(סԺ����_In) > Trunc(d_��Ժʱ��) Then
      Select Trunc(d_��Ժʱ��) - Trunc(d_��Ժʱ��) + n_Addday Into n_Days From Dual;
    Else
      Select Trunc(סԺ����_In) - Trunc(d_��Ժʱ��) + n_Addday Into n_Days From Dual;
    End If;
  End If;

  Return(n_Days);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Calcindaysnew;
/


--119477:����,2018-07-09,�������֮��,�ٴη�����ܵ����µ�
Create Or Replace Procedure Zl_������λ���_Update
(
  �ļ�id_In   In ���˻�������.�ļ�id%Type,
  ����ʱ��_In In ���˻�������.����ʱ��%Type,
  ����ʱ��_In In ���˻�������.����ʱ��%Type, --���ܼ�¼�ķ���ʱ��
  ��¼����_In In ���˻�����ϸ.��¼����%Type, --������Ŀ=1��ǩ����¼=5����ǩ��¼=15
  ��Ŀ���_In In ���˻�����ϸ.��Ŀ���%Type, --������Ŀ����ţ��ǻ�����Ŀ�̶�Ϊ0
  ��¼����_In In ���˻�����ϸ.��¼����%Type := Null, --��¼���ݣ��������Ϊ�գ��������ǰ�����ݣ�37��38/37
  ���²�λ_In In ���˻�����ϸ.���²�λ%Type := Null,
  ����Ա_In   In ���˻�������.������%Type := Null,
  ��¼���_In In ���˻�����ϸ.��¼���%Type := Null, --���÷������(һ�����ݶ�Ӧ������ͬ��Ŀ����ϸ)
  ������Դ_In In ���˻�����ϸ.������Դ%Type := 0,
  ɾ��_In     In Number := 0
) Is
  Intins     Number(18);
  Int����    Number(1);
  n_Exists   Number(1);
  n_Newid    ���˻�������.Id%Type;
  n_Oldid    ���˻�������.Id%Type;
  n_Synchro  Number(1);
  n_����id   ���˻�������.Id%Type;
  n_�ļ�id   ���˻�������.�ļ�id%Type;
  v_�����ı� ���˻�������.�����ı�%Type;

  n_������� ���˻�������.�������%Type;
  v_����id   ���ű�.Id%Type;
  v_������   ��Ա��.����%Type;
  n_��¼id   ���˻�������.Id%Type;
  n_��ϸid   ���˻�����ϸ.Id%Type;
  v_������Դ ���˻�����ϸ.������Դ%Type;
  n_��Ŀ���� �����¼��Ŀ.��Ŀ����%Type;
  --��ȡ�ò��˵�ǰ��������δ�����Ļ����ļ������ļ���ʼʱ��С�ڵ��ڼ�¼����ʱ����ļ��б�ͬ������ʹ��
  Cursor Cur_Fileformats Is
    Select a.Id As ��ʽid, b.Id As �ļ�id, a.����, a.����, b.Ӥ��
    From �����ļ��б� A, ���˻����ļ� B, ���˻����ļ� C, ���˻������� D
    Where a.���� = 3 And a.���� <> 1 And a.Id = b.��ʽid And b.Id <> c.Id And b.����ʱ�� Is Null And b.��ʼʱ�� <= d.����ʱ�� And
          (a.ͨ�� = 1 Or (a.ͨ�� = 2 And b.����id = c.����id)) And c.����id = b.����id And c.��ҳid = b.��ҳid And c.Ӥ�� = b.Ӥ�� And
          c.Id = d.�ļ�id And d.Id = n_��¼id And c.Id = �ļ�id_In
    Order By a.���;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --ȡ��¼ID
  Int����  := 0;
  n_��¼id := 0;
  Intins   := 0;
  If ����Ա_In Is Null Then
    v_������ := Zl_Username;
  Else
    v_������ := ����Ա_In;
  End If;

  Begin
    Select Max(ID) Into n_����id From ���˻������� Where �ļ�id = �ļ�id_In And ����ʱ�� = ����ʱ��_In;
  End;
  Begin
    Select ID, �������
    Into n_��¼id, n_�������
    From ���˻�������
    Where �ļ�id = �ļ�id_In And ����ʱ�� = ����ʱ��_In;
  Exception
    When Others Then
      n_��¼id := 0;
  End;

  Intins := 0;
  --���ҳ��˱���Ҫɾ�������ݣ��Ƿ񻹴�������Ч�����ݣ��������ֻɾ���������ݣ�����ɾ���˷���ʱ���Ӧ���������ݡ�
  If ɾ��_In = 1 Then
  
    Select a.Id
    Into n_�ļ�id
    From ���˻����ļ� A, ���˻����ļ� B, �����ļ��б� C
    Where b.Id = �ļ�id_In And a.����id = b.����id And a.��ҳid = b.��ҳid And a.Ӥ�� = b.Ӥ�� And a.��ʽid = c.Id And c.���� = 3 And
          c.���� = -1 And a.��ʼʱ�� < ����ʱ��_In And (a.����ʱ�� > ����ʱ��_In Or a.����ʱ�� Is Null);
    Select Max(1)
    Into Intins
    From ���˻�������
    Where Instr(�����ı� || ',', ',' || n_����id || ',') > 0 And �ļ�id = n_�ļ�id;
    Begin
      Select ID, �������
      Into n_��¼id, n_�������
      From ���˻�������
      Where �ļ�id = n_�ļ�id And ����ʱ�� = ����ʱ��_In;
    Exception
      When Others Then
        n_��¼id := 0;
    End;
    
    Begin
      Select ID
      Into n_��ϸid
      From ���˻�����ϸ
      Where ��¼id = n_��¼id And Nvl(��Ŀ���, 0) = ��Ŀ���_In And Nvl(���²�λ, 'TWBW') = Nvl(���²�λ_In, 'TWBW') and  Nvl(��¼���, 1) = Nvl(��¼���_In, 1)  And ��ֹ�汾 Is Null ;
    Exception
      --�������˳�
      When Others Then
        Return;
    End;
  
    If Intins = 0 Then
      Select Count(ID)
      Into Intins
      From ���˻�����ϸ
      Where ��¼id = n_��¼id And Mod(��¼����, 10) <> 5 And ��ֹ�汾 Is Null And ID <> n_��ϸid;
      If Intins = 0 Then
        Delete From ���˻�����ϸ Where ��¼id = n_��¼id;
      Else
        Delete From ���˻�����ϸ Where ID = n_��ϸid;
      End If;
      Delete From ���˻������� A
      Where a.Id = n_��¼id And Not Exists (Select 1 From ���˻�����ϸ B Where b.��¼id = a.Id);
    Else
      Select �����ı� Into v_�����ı� From ���˻������� Where ID = n_��¼id;
      v_�����ı� := Replace(v_�����ı� || ',', ',' || n_����id || ',', ',');
      v_�����ı� := Substr(v_�����ı�, 1, Length(v_�����ı�) - 1);
      Update ���˻������� Set �����ı� = v_�����ı� Where ID = n_��¼id;
      If v_�����ı� is null Then
        Select Count(ID)
        Into Intins
        From ���˻�����ϸ
        Where ��¼id = n_��¼id And Mod(��¼����, 10) <> 5 And ��ֹ�汾 Is Null And ID <> n_��ϸid;
        If Intins = 0 Then
          Delete From ���˻�����ϸ Where ��¼id = n_��¼id;
        Else
          Delete From ���˻�����ϸ Where ID = n_��ϸid;
        End If;
        Delete From ���˻������� A
        Where a.Id = n_��¼id And Not Exists (Select 1 From ���˻�����ϸ B Where b.��¼id = a.Id);
      End If;
    End If;
    Return; --ɾ����ֱ���˳�
  End If;

  --############
  --���λ��ܱ�������
  --############
  --�������µ���һ������ʼ��ֻ����һ����Ч�����µ��ļ���
  --������±������ͬ����ʱ������ݣ�ʹ������ID
  For Row_Format In Cur_Fileformats Loop
    If Row_Format.���� = -1 Then
      If Row_Format.���� = '1' Then
        Begin
          Select 1, h.��Ŀ����
          Into Intins, n_��Ŀ����
          From (Select To_Char(f.��Ŀ���) As ���, g.��Ŀ����
                 From ���¼�¼��Ŀ F, �����¼��Ŀ G
                 Where f.��Ŀ��� = g.��Ŀ��� And g.��Ŀ���� = 2 And
                       (g.���ÿ��� = 1 Or (g.���ÿ��� = 2 And Exists
                        (Select 1 From �������ÿ��� D Where g.��Ŀ��� = d.��Ŀ��� And d.����id = v_����id))) And Nvl(g.Ӧ�÷�ʽ, 0) <> 0 And
                       (Nvl(g.���ò���, 0) = 0 Or Nvl(g.���ò���, 0) = Decode(Nvl(Row_Format.Ӥ��, 0), 0, 1, 2))
                 Union All
                 Select b.�����ı� As ���, 1 As ��Ŀ����
                 From �����ļ��ṹ A, �����ļ��ṹ B
                 Where a.�ļ�id = Row_Format.��ʽid And a.��id Is Null And a.������� In (2, 3) And b.��id = a.Id) H
          Where Instr(',' || h.��� || ',', ',' || ��Ŀ���_In || ',', 1) > 0;
        Exception
          When Others Then
            Intins := 0;
        End;
      Else
        Begin
          Select 1, g.��Ŀ����
          Into Intins, n_��Ŀ����
          From ���¼�¼��Ŀ F, �����¼��Ŀ G
          Where f.��Ŀ��� = g.��Ŀ��� And Nvl(g.Ӧ�÷�ʽ, 0) <> 0 And g.����ȼ� >= 0 And
                (Nvl(g.���ò���, 0) = 0 Or Nvl(g.���ò���, 0) = Decode(Nvl(Row_Format.Ӥ��, 0), 0, 1, 2)) And f.��Ŀ��� = ��Ŀ���_In And
                (g.���ÿ��� = 1 Or (g.���ÿ��� = 2 And Exists
                 (Select 1 From �������ÿ��� D Where g.��Ŀ��� = d.��Ŀ��� And d.����id = v_����id)));
        Exception
          When Others Then
            Intins := 0;
        End;
      End If;
    
      If Intins > 0 Then
      
        --LPF,2013-01-23,������Ŀ�Ƿ���Ҫ����ͬ��(������ǰ�Ѿ�ͬ���������ݣ�Ϊ�˱�֤��¼�������µ�����һֱ�������ݴ˺����жϡ�)
        n_Synchro := Zl_Temperatureprogram(�ļ�id_In, v_����id, ��Ŀ���_In, ����ʱ��_In);
        Begin
          Select b.Id
          Into n_Newid
          From ���˻����ļ� A, ���˻������� B
          Where a.Id = Row_Format.�ļ�id And b.�ļ�id = a.Id And b.����ʱ�� = ����ʱ��_In;
        Exception
          When Others Then
            n_Newid := 0;
        End;
        n_Oldid := n_Newid;
        If n_Newid = 0 And n_Synchro = 1 Then
          Select ���˻�������_Id.Nextval Into n_Newid From Dual;
          --�������µ�����¼
          Insert Into ���˻�������
            (ID, �ļ�id, ������, ����ʱ��, ����ʱ��, ���汾, �����ı�)
          Values
            (n_Newid, Row_Format.�ļ�id, v_������, Sysdate, ����ʱ��_In, 1, ',' || n_����id);
        Else
          If n_Oldid <> 0 Then
            Select Max(1)
            Into n_Exists
            From ���˻�������
            Where ID = n_Oldid And Instr(�����ı� || ',', ',' || n_����id || ',') > 0;
            If n_Exists Is Null Then
              Update ���˻������� Set �����ı� = �����ı� || ',' || n_����id Where ID = n_Oldid;
            End If;
          
          End If;
        End If;
      
        If n_Newid > 0 Then
          --����δͬ�������µ�����(��ȻҪ���Ӷ���ѯ)
          Select Count(*)
          Into v_������Դ
          From ���˻�����ϸ
          Where ��¼id = n_Newid And ��Ŀ��� = ��Ŀ���_In And
                Decode(n_��Ŀ����, 2, Nvl(���²�λ, '��'), Nvl(���²�λ_In, '��')) = Nvl(���²�λ_In, '��');
          If v_������Դ = 0 Then
            --˵����ͬ����ʼ�Ѿ����й����
            If n_Synchro = 1 Then
              --û�м�����Ŀ�Ƿ���Ҫͬ��
            
              Insert Into ���˻�����ϸ
                (ID, ��¼id, ��¼����, ��Ŀ����, ��Ŀid, ��Ŀ���, ��Ŀ����, ��Ŀ����, ��¼����, ��Ŀ��λ, ��¼���, ���²�λ, ������Դ, ��Դid, ��ʼ�汾, ��ֹ�汾, ��¼��,
                 ��¼ʱ��, ��¼���)
                Select ���˻�����ϸ_Id.Nextval, n_Newid, ��¼����_In, a.������, a.��Ŀid, a.��Ŀ���, a.��Ŀ����, a.��Ŀ����, ��¼����_In, a.��Ŀ��λ, 0,
                       ���²�λ_In, ������Դ_In, Null, 1, Null, b.��¼��, Sysdate, 1
                From �����¼��Ŀ A, ���˻�����ϸ B
                Where a.��Ŀ��� = b.��Ŀ���(+) And b.��ֹ�汾(+) Is Null And b.��¼id(+) = n_��¼id And a.��Ŀ��� = ��Ŀ���_In And
                      Rownum < 2;
            
              If Sql%RowCount > 0 Then
                Int���� := 1;
              End If;
            End If;
          Else
            Update ���˻�����ϸ
            Set ��¼���� = ��¼����_In
            Where ��¼id = n_Newid And ��Ŀ��� = ��Ŀ���_In And
                  Decode(n_��Ŀ����, 2, Nvl(���²�λ, '��'), Nvl(���²�λ_In, '��')) = Nvl(���²�λ_In, '��');
            If Sql%RowCount > 0 Then
              Int���� := 1;
            End If;
          End If;
        End If;
      End If;
    End If;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_������λ���_Update;
/

--119477:����,2018-06-11,�������֮��,�ٴη�����ܵ����µ�

Create Or Replace Procedure Zl_���˻�������_Collect
(
  �ļ�id_In   In ���˻�������.�ļ�id%Type,
  ����ʱ��_In In ���˻�������.����ʱ��%Type,
  �������_In In ���˻�������.�������%Type,
  �����ı�_In In ���˻�������.�����ı�%Type,
  ���ܱ��_In In ���˻�������.���ܱ��%Type,
  ��ʼʱ��_In In ���˻�������.��ʼʱ��%Type,
  ����ʱ��_In In ���˻�������.����ʱ��%Type,
  ɾ��_In     Number := 0
) Is
  n_Exist  Number(1);
  v_��¼id ���˻�������.Id%Type;
  v_��Դid ���˻�������.Id%Type;
  v_User   ��Ա��.����%Type;
  n_�ļ�id ���˻�������.�ļ�id%Type;
Begin
  If ɾ��_In = 0 Then
    v_User := Zl_Username;
    Begin
      Select 1 Into n_Exist From ���˻������� Where �ļ�id = �ļ�id_In And ����ʱ�� = ����ʱ��_In;
    Exception
      When Others Then
        n_Exist := 0;
    End;
  
    If n_Exist = 0 Then
      Insert Into ���˻�������
        (ID, �ļ�id, ����ʱ��, ���汾, ������, ����ʱ��, �������, �����ı�, ���ܱ��, ��ʼʱ��, ����ʱ��)
      Values
        (���˻�������_Id.Nextval, �ļ�id_In, ����ʱ��_In, 1, v_User, Sysdate, �������_In, �����ı�_In, ���ܱ��_In, ��ʼʱ��_In, ����ʱ��_In);
    End If;
  Else
    Select ID Into v_��¼id From ���˻������� Where �ļ�id = �ļ�id_In And ����ʱ�� = ����ʱ��_In;
    Select a.Id
    Into n_�ļ�id
    From ���˻����ļ� A, ���˻����ļ� B, �����ļ��б� C
    Where b.Id = �ļ�id_In And a.����id = b.����id And a.��ҳid = b.��ҳid And a.Ӥ�� = b.Ӥ�� And A.��ʽid = c.Id And c.���� = 3 And
          c.���� = -1 And a.��ʼʱ�� < ����ʱ��_In And (a.����ʱ�� > ����ʱ��_In Or a.����ʱ�� Is Null);
    Select Max(a.��¼id)
    Into v_��Դid
    From ���˻�����ϸ A, ���˻�����ϸ B
    Where a.��Դid = b.Id(+) And b.��¼id = v_��¼id;
  
    For r_List In (Select a.����ʱ��, b.��Ŀ���, b.��Ŀ����
                   From ���˻������� A, ���˻�����ϸ B
                   Where �ļ�id = n_�ļ�id And a.Id = b.��¼id And Instr(�����ı�, v_��¼id) > 0 And b.������Դ = 1 And ��Դid Is Null) Loop
      Zl_������λ���_Update(�ļ�id_In, r_List.����ʱ��, ����ʱ��_In, 1, r_List.��Ŀ���, Null, Null, Null, Null, 1, 1);
    End Loop;
    Delete ���˻�����ϸ Where ��¼id = v_��¼id;
    Delete ���˻������� Where �ļ�id = �ļ�id_In And ����ʱ�� = ����ʱ��_In;
  
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˻�������_Collect;
/
--119477:����,2018-06-11,�������֮��,�ٴη�����ܵ����µ�

Create Or Replace Procedure Zl_�����¼��Ŀ_Update
(
  ��Ŀ���_In In �����¼��Ŀ.��Ŀ���%Type,
  ��Ŀ����_In In �����¼��Ŀ.��Ŀ����%Type,
  ��Ŀ����_In In �����¼��Ŀ.��Ŀ����%Type,
  ��Ŀ����_In In �����¼��Ŀ.��Ŀ����%Type,
  ��ĿС��_In In �����¼��Ŀ.��ĿС��%Type,
  ��Ŀ��λ_In In �����¼��Ŀ.��Ŀ��λ%Type,
  ��Ŀ��ʾ_In In �����¼��Ŀ.��Ŀ��ʾ%Type,
  ��Ŀֵ��_In In �����¼��Ŀ.��Ŀֵ��%Type,
  ����ȼ�_In In �����¼��Ŀ.����ȼ�%Type,
  ������_In   In �����¼��Ŀ.������%Type,
  ��Ŀid_In   In �����¼��Ŀ.��Ŀid%Type,
  Ӧ�÷�ʽ_In In �����¼��Ŀ.Ӧ�÷�ʽ%Type,
  ���ò���_In In �����¼��Ŀ.���ò���%Type,
  ��Ŀ����_In In �����¼��Ŀ.��Ŀ����%Type := 1,
  Ӧ�ó���_In In �����¼��Ŀ.Ӧ�ó���%Type := 0,
  ˵��_In     In �����¼��Ŀ.˵��%Type := Null,
  ȱʡֵ_In   In �����¼��Ŀ.ȱʡֵ%Type := Null,
  �������_In In �����¼��Ŀ.�������%Type := Null
) Is
  n_���� Number(1);
Begin
  n_���� := 0;
  Select Count(��Ŀ���) Into n_���� From �����¼��Ŀ Where ��Ŀ��� = ��Ŀ���_In And ��Ŀ��ʾ = 4;
  If �������_In Is Null Then
    Update �����¼��Ŀ
    Set ��Ŀ���� = ��Ŀ����_In, ��Ŀ���� = ��Ŀ����_In, ��Ŀ���� = ��Ŀ����_In, ��ĿС�� = ��ĿС��_In, ��Ŀ��λ = ��Ŀ��λ_In, ��Ŀ��ʾ = ��Ŀ��ʾ_In, ��Ŀֵ�� = ��Ŀֵ��_In,
        ����ȼ� = ����ȼ�_In, ������ = ������_In, ��Ŀid = ��Ŀid_In, Ӧ�÷�ʽ = Ӧ�÷�ʽ_In, ���ò��� = ���ò���_In, ��Ŀ���� = ��Ŀ����_In, Ӧ�ó��� = Ӧ�ó���_In,
        ˵�� = ˵��_In, ȱʡֵ = ȱʡֵ_In
    Where ��Ŀ��� = ��Ŀ���_In;
  Else
    Update �����¼��Ŀ Set ������� = �������_In Where ��Ŀ��� = ��Ŀ���_In;
  End If;
  If ��Ŀ���_In = 2 Then
    Update �����¼��Ŀ
    Set ��Ŀ���� = ��Ŀ����_In, ��Ŀ���� = ��Ŀ����_In, ��ĿС�� = ��ĿС��_In, ��Ŀ��λ = ��Ŀ��λ_In, ��Ŀ��ʾ = ��Ŀ��ʾ_In, ��Ŀֵ�� = ��Ŀֵ��_In, ����ȼ� = ����ȼ�_In,
        ������ = ������_In, ��Ŀ���� = ��Ŀ����_In, Ӧ�ó��� = Ӧ�ó���_In, ˵�� = ˵��_In
    Where ��Ŀ��� = -1;
  End If;

  If ��Ŀ���_In = 4 Or ��Ŀ���_In = 5 Then
    Update �����¼��Ŀ
    Set ��Ŀ���� = ��Ŀ����_In, ��Ŀ���� = ��Ŀ����_In, ��ĿС�� = ��ĿС��_In, ��Ŀ��λ = ��Ŀ��λ_In, ��Ŀ��ʾ = ��Ŀ��ʾ_In, ����ȼ� = ����ȼ�_In, ������ = ������_In,
        Ӧ�÷�ʽ = Ӧ�÷�ʽ_In, ���ò��� = ���ò���_In, ��Ŀ���� = ��Ŀ����_In, Ӧ�ó��� = Ӧ�ó���_In, ˵�� = ˵��_In
    Where ��Ŀ��� In (4, 5);
  End If;
  If ��Ŀ��ʾ_In = 4 Then
    Insert Into ���������Ŀ
      (���, �����)
      Select ��Ŀ���_In, Null From Dual Where Not Exists (Select 1 From ���������Ŀ Where ��� = ��Ŀ���_In);
  Else
    If n_���� = 1 Then
      Delete ���������Ŀ Where ��� = ��Ŀ���_In;
      Update ���������Ŀ Set ����� = Null Where ����� = ��Ŀ���_In;
    End If;
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����¼��Ŀ_Update;
/
--119477:����,2018-06-11,�������֮��,�ٴη�����ܵ����µ�

CREATE OR REPLACE Procedure Zl_���˻�������_Update
(
  �ļ�id_In   In ���˻�������.�ļ�id%Type,
  ����ʱ��_In In ���˻�������.����ʱ��%Type,
  ��¼����_In In ���˻�����ϸ.��¼����%Type, --������Ŀ=1��ǩ����¼=5����ǩ��¼=15
  ��Ŀ���_In In ���˻�����ϸ.��Ŀ���%Type, --������Ŀ����ţ��ǻ�����Ŀ�̶�Ϊ0
  ��¼����_In In ���˻�����ϸ.��¼����%Type := Null, --��¼���ݣ��������Ϊ�գ��������ǰ�����ݣ�37��38/37
  ���²�λ_In In ���˻�����ϸ.���²�λ%Type := Null,
  ���˼�¼_In In Number := 1,
  ������Դ_In In ���˻�����ϸ.������Դ%Type := 0,
  ��ǩ_In     In Number := 0,
  ����Ա_In   In ���˻�������.������%Type := Null,
  ��¼���_In In ���˻�����ϸ.��¼���%Type := Null, --���÷������(һ�����ݶ�Ӧ������ͬ��Ŀ����ϸ)
  ������_In In ���˻�����ϸ.������%Type := Null, --���÷������(��¼������Ŀ������������Ŀ���)
  δ��˵��_In In ���˻�����ϸ.δ��˵��%Type := Null, --��������洢ҽ��ID:���ͺ�
  �������_In In number :=0
) Is
  Intins      Number(18);
  Int����     Number(1);
  n_Newid     ���˻�������.Id%Type;
  n_Oldid     ���˻�������.Id%Type;
  n_����      ���˻����ӡ.����%Type;
  n_Mutilbill Number(1);
  n_Syntend   Number(1);
  n_Synchro   Number(1);

  n_�������     ���˻�������.�������%Type;
  v_����id       ���ű�.Id%Type;
  v_������       ��Ա��.����%Type;
  v_��¼��       ��Ա��.����%Type;
  n_�ļ�id       ���˻�������.�ļ�id%Type;
  n_��¼id       ���˻�������.Id%Type;
  n_��ϸid       ���˻�����ϸ.Id%Type;
  n_��Դid       ���˻�����ϸ.��Դid%Type;
  v_������Դ     ���˻�����ϸ.������Դ%Type;
  n_��߰汾     ���˻�����ϸ.��ʼ�汾%Type;
  n_��Ŀ����     �����¼��Ŀ.��Ŀ����%Type;
  n_����id       ���˻����ļ�.����id%Type;
  n_��ҳid       ���˻����ļ�.��ҳid%Type;
  n_Ӥ��         ���˻����ļ�.Ӥ��%Type;
  d_Ӥ����Ժʱ�� ����ҽ����¼.��ʼִ��ʱ��%Type;
  d_�ļ���ʼʱ�� ���˻����ļ�.��ʼʱ��%Type;
  --��ȡ�ò��˵�ǰ��������δ�����Ļ����ļ������ļ���ʼʱ��С�ڵ��ڼ�¼����ʱ����ļ��б�ͬ������ʹ��
  Cursor Cur_Fileformats Is
    Select a.Id As ��ʽid, b.Id As �ļ�id, a.����, a.����, b.Ӥ��
    From �����ļ��б� A, ���˻����ļ� B, ���˻����ļ� C, ���˻������� D
    Where a.���� = 3 And a.���� <> 1 And a.Id = b.��ʽid And b.Id <> c.Id And b.����ʱ�� Is Null And b.��ʼʱ�� <= d.����ʱ�� And
          (a.ͨ�� = 1 Or (a.ͨ�� = 2 And b.����id = c.����id)) And c.����id = b.����id And c.��ҳid = b.��ҳid And c.Ӥ�� = b.Ӥ�� And
          c.Id = d.�ļ�id And d.Id = n_��¼id And c.Id = �ļ�id_In
    Order By a.���;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --ȡ��¼ID
  Int����     := 0;
  n_��¼id    := 0;
  n_Mutilbill := 0;
  n_Syntend   := 0;
  If ����Ա_In Is Null Then
    v_������ := Zl_Username;
  Else
    v_������ := ����Ա_In;
  End If;

  --����Ƕ�Ӧ��ݻ����ļ�ֵΪ1����ʾ��ͬ�����������ļ������򲻴����ļ�ͬ��
  n_Mutilbill := Zl_To_Number(zl_GetSysParameter('��Ӧ��ݻ����ļ�', 1255));
  --��������ݻ����ļ�֮������ͬ��,���Զ�ͬ��,����ͬ��
  n_Syntend := Zl_To_Number(zl_GetSysParameter('��������ͬ��', 1255));

  Begin
    Select ID, �������
    Into n_��¼id, n_�������
    From ���˻�������
    Where �ļ�id = �ļ�id_In And ����ʱ�� = ����ʱ��_In;
  Exception
    When Others Then
      n_��¼id := 0;
  End;

  --����ǲ��Ǳ��˵ļ�¼
  ---------------------------------------------------------------------------------------------------------------------
  If ���˼�¼_In = 0 And n_��¼id > 0 And ��ǩ_In = 0 Then
    v_��¼�� := '';
    Begin
      Select ��¼��
      Into v_��¼��
      From ���˻�����ϸ
      Where ��¼id = n_��¼id And ��Ŀ��� = ��Ŀ���_In And Nvl(���²�λ, 'TWBW') = Nvl(���²�λ_In, 'TWBW') And
            Nvl(��¼���, 0) = Nvl(��¼���_In, 0) And ��ֹ�汾 Is Null;
    Exception
      When Others Then
        v_��¼�� := '';
    End;
    If v_��¼�� Is Not Null And v_��¼�� <> v_������ Then
      v_Error := '����Ȩ�޸����˵ǼǵĻ������ݣ�';
      Raise Err_Custom;
    End If;
  End If;

  --����Ƿ����
  Select ����id, ��ҳid, Nvl(Ӥ��, 0), ��ʼʱ��
  Into n_����id, n_��ҳid, n_Ӥ��, d_�ļ���ʼʱ��
  From ���˻����ļ�
  Where ID = �ļ�id_In;
  d_Ӥ����Ժʱ�� := Null;
  If n_Ӥ�� <> 0 Then
    Begin
      Select ��ʼִ��ʱ��
      Into d_Ӥ����Ժʱ��
      From ����ҽ����¼ B, ������ĿĿ¼ C
      Where b.������Ŀid + 0 = c.Id And b.ҽ��״̬ = 8 And Nvl(b.Ӥ��, 0) <> 0 And c.��� = 'Z' And
            Instr(',3,5,11,', ',' || c.�������� || ',', 1) > 0 And b.����id = n_����id And b.��ҳid = n_��ҳid And b.Ӥ�� = n_Ӥ��;
    Exception
      When Others Then
        d_Ӥ����Ժʱ�� := Null;
    End;
  End If;
  If d_Ӥ����Ժʱ�� Is Null Then
    v_����id := 0;
    Begin
      Select a.����id
      Into v_����id
      From ���˱䶯��¼ A, ���˻����ļ� B
      Where a.����id Is Not Null And a.����id = b.����id And a.��ҳid = b.��ҳid And b.Id = �ļ�id_In And
            (To_Date(To_Char(����ʱ��_In, 'YYYY-MM-DD HH24:MI') || '59', 'YYYY-MM-DD HH24:MI:SS') >= a.��ʼʱ�� And
            (To_Date(To_Char(����ʱ��_In, 'YYYY-MM-DD HH24:MI') || '00', 'YYYY-MM-DD HH24:MI:SS') < = Nvl(a.��ֹʱ��, Sysdate) Or
            a.��ֹʱ�� Is Null)) And Rownum < 2;
    Exception
      When Others Then
        v_����id := 0;
    End;
    If v_����id = 0 Then
      v_Error := '���ݷ���ʱ�� ' || To_Char(����ʱ��_In, 'YYYY-MM-DD HH24:MI:SS') || ' ���ڲ�����Ч�䶯ʱ�䷶Χ�ڣ����ܲ�����';
      Raise Err_Custom;
    End If;
  Else
    If ����ʱ��_In < d_�ļ���ʼʱ�� Or ����ʱ��_In > d_Ӥ����Ժʱ�� Then
      v_Error := '���ݷ���ʱ�� ' || To_Char(����ʱ��_In, 'YYYY-MM-DD HH24:MI:SS') || ' ���ڲ�����Ч�䶯ʱ�䷶Χ�ڣ����ܲ�����';
      Raise Err_Custom;
    End If;
  End If;

  --���������Դ<>0���˳�
  n_��Դid := 0;
  If n_��¼id > 0 Then
    Begin
      Select ������Դ, Nvl(��Դid, 0)
      Into v_������Դ, n_��Դid
      From ���˻�����ϸ
      Where ��¼id = n_��¼id And Nvl(��Ŀ���, 0) = ��Ŀ���_In And Nvl(���²�λ, 'TWBW') = Nvl(���²�λ_In, 'TWBW') And
            Nvl(��¼���, 0) = Nvl(��¼���_In, 0);
    Exception
      When Others Then
        v_������Դ := 0;
    End;
    If v_������Դ > 0 And n_��Դid > 0 Then
      Return;
    End If;
  End If;

  --ȡ��߰汾
  Select Nvl(Max(Nvl(a.��ʼ�汾, 1)), 0) + 1, Count(b.Id)
  Into n_��߰汾, Intins
  From ���˻�����ϸ A, ���˻������� B
  Where b.Id = n_��¼id And a.��¼id = b.Id And Mod(a.��¼����, 10) = 5;

  --Ŀǰ�Ѿ�ǩ�������ݲ����޸ģ�ֻ������ǩģʽ�½����޸ģ�����ǩ_In=1
  If ��ǩ_In <> 1 And Intins > 0 Then
    v_Error := '����ʱ�� ' || To_Char(����ʱ��_In, 'YYYY-MM-DD HH24:MI:SS') || ' ����Ӧ�������Ѿ�ǩ������ǩ�����ܼ���������' || Chr(13) || Chr(10) ||
               '��������������粢����������ģ���ˢ�º����ԣ�';
    Raise Err_Custom;
  End If;
  Intins := 0;

  --������ʱ,Ҫ������ݣ���ǩ����ʱ���Զ������ǩ�������޸ĵ����ݣ����Դ˴�ֻ�迼����ǩ���ɣ�
  If ��¼����_In Is Null Then
    Begin
      Select ID
      Into n_��ϸid
      From ���˻�����ϸ
      Where ��¼id = n_��¼id And Nvl(��Ŀ���, 0) = ��Ŀ���_In And Nvl(���²�λ, 'TWBW') = Nvl(���²�λ_In, 'TWBW') And
            Nvl(��¼���, 0) = Nvl(��¼���_In, 0) And ��ֹ�汾 Is Null;
    Exception
      --�������˳�
      When Others Then
        Return;
    End;
  
    --���ҳ��˱���Ҫɾ�������ݣ��Ƿ񻹴�������Ч�����ݣ��������ֻɾ���������ݣ�����ɾ���˷���ʱ���Ӧ���������ݡ�
    Select Count(ID)
    Into Intins
    From ���˻�����ϸ
    Where ��¼id = n_��¼id And Mod(��¼����, 10) <> 5 And ��ֹ�汾 Is Null And ID <> n_��ϸid;
    If Intins = 0 Then
      Delete From ���˻�����ϸ Where ��¼id = n_��¼id;
    Else
      Delete From ���˻�����ϸ Where ID = n_��ϸid;
    End If;
  
    Delete From ���˻������� A
    Where a.Id = n_��¼id And Not Exists (Select 1 From ���˻�����ϸ B Where b.��¼id = a.Id);
  
    --�����ɾ��ǩ�����޸Ĳ��������һ������,��Ӧ��ǩ����¼����ֹ�汾��Ϊ��
    Begin
      Select 1
      Into Intins
      From ���˻�����ϸ
      Where ��ʼ�汾 = n_��߰汾 And ��ֹ�汾 Is Null And ��¼���� = 1 And ��¼id = n_��¼id;
    Exception
      When Others Then
        Intins := 0;
    End;
    If Intins = 0 Then
      Update ���˻�����ϸ Set ��ֹ�汾 = Null Where ��¼���� = 5 And ��ʼ�汾 = n_��߰汾 - 1 And ��¼id = n_��¼id;
    End If;
    If Nvl(n_�������, 0) <> 0 and  �������_In=0 Then
      Return;
    End If;
  
    --############
    --�����������
    --############
    For Rsdel In (Select Distinct ��¼id From ���˻�����ϸ Where ��Դid = n_��ϸid) Loop
    
      Delete ���˻�����ϸ Where ��Դid = n_��ϸid And ��¼id = Rsdel.��¼id;
      --ɾ����Ӧ�Ĵ�ӡ����
      Begin
        Select Count(*) Into Intins From ���˻�����ϸ Where ��¼id = Rsdel.��¼id;
      Exception
        When Others Then
          Intins := 0;
      End;
      If Intins = 0 Then
        --��ȡ������ݶ�Ӧ���ļ�ID
        Begin
          Select b.Id, a.����
          Into n_�ļ�id, Intins
          From �����ļ��б� A, ���˻����ļ� B, ���˻������� C
          Where a.Id = b.��ʽid And b.Id = c.�ļ�id And c.Id = Rsdel.��¼id;
        Exception
          When Others Then
            n_�ļ�id := 0;
        End;
        Delete ���˻������� Where ID = Rsdel.��¼id;
        If Intins <> -1 Then
          Zl_���˻����ӡ_Update(n_�ļ�id, ����ʱ��_In, 1, 1);
        End If;
      End If;
    End Loop;
  Else
    --���¼�����Ŀ�Ƿ����ڸü�¼��
    Begin
      Select 1
      Into Intins
      From (Select b.��Ŀ���
             From �����ļ��ṹ A, �����¼��Ŀ B
             Where a.Ҫ������ = b.��Ŀ���� And b.��Ŀ��� = ��Ŀ���_In And
                   ��id = (Select b.Id
                          From ���˻����ļ� A, �����ļ��ṹ B
                          Where a.Id = �ļ�id_In And a.��ʽid = b.�ļ�id And b.��id Is Null And b.������� = 4)
             Union
             Select ��Ŀ���
             From �����¼��Ŀ
             Where ��Ŀ���� = 2 And ��Ŀ��� = ��Ŀ���_In
             Union
             Select ��Ŀ���
             From �����¼��Ŀ
             Where ��Ŀ��ʾ = 4 And ��Ŀ��� = ��Ŀ���_In);
    Exception
      When Others Then
        Intins := 0;
    End;
    If Intins = 0 Then
      Return;
    End If;
    If n_��¼id = 0 Then
      Select ���˻�������_Id.Nextval Into n_��¼id From Dual;
    
      Insert Into ���˻�������
        (ID, �ļ�id, ����ʱ��, ���汾, ������, ����ʱ��)
      Values
        (n_��¼id, �ļ�id_In, ����ʱ��_In, n_��߰汾, v_������, Sysdate);
    End If;
  
    --���뱾�εǼǵĲ��˻�����ϸ
    Update ���˻�����ϸ
    Set ��¼���� = ��¼����_In, δ��˵�� = δ��˵��_In, ��¼�� = v_������, ��¼ʱ�� = Sysdate
    Where ��¼id = n_��¼id And ��Ŀ��� = ��Ŀ���_In And Nvl(���²�λ, 'TWBW') = Nvl(���²�λ_In, 'TWBW') And
          Nvl(��¼���, 0) = Nvl(��¼���_In, 0) And ��ʼ�汾 = n_��߰汾 And ��ֹ�汾 Is Null;
    If Sql%RowCount = 0 Then
      Select ���˻�����ϸ_Id.Nextval Into n_��ϸid From Dual;
      Insert Into ���˻�����ϸ
        (ID, ��¼id, ��¼����, ��Ŀ����, ��Ŀid, ������, ��Ŀ���, ��Ŀ����, ��Ŀ����, ��¼����, ��Ŀ��λ, ��¼���, ��¼���, ���²�λ, ������Դ, ����, δ��˵��, ��ʼ�汾, ��ֹ�汾,
         ��¼��, ��¼ʱ��)
        Select n_��ϸid, n_��¼id, ��¼����_In, a.������, a.��Ŀid, ������_In, a.��Ŀ���, Upper(a.��Ŀ����), a.��Ŀ����, ��¼����_In, a.��Ŀ��λ, 0,
               ��¼���_In, ���²�λ_In, ������Դ_In, Nvl(b.����, 0), δ��˵��_In, n_��߰汾, Null, v_������, Sysdate
        From �����¼��Ŀ A, ���˻�����ϸ B
        Where a.��Ŀ��� = b.��Ŀ���(+) And b.��ֹ�汾(+) Is Null And b.��¼id(+) = n_��¼id And a.��Ŀ��� = ��Ŀ���_In And Rownum < 2;
    End If;
    Select ID
    Into n_��ϸid
    From ���˻�����ϸ
    Where ��¼id = n_��¼id And ��Ŀ��� = ��Ŀ���_In And Nvl(���²�λ, 'TWBW') = Nvl(���²�λ_In, 'TWBW') And
          Nvl(��¼���, 0) = Nvl(��¼���_In, 0) And ��ʼ�汾 = n_��߰汾 And ��ֹ�汾 Is Null;
    --��д��ʷ���ݼ�ǩ����¼����ֹ�汾
    Update ���˻�����ϸ
    Set ��ֹ�汾 = n_��߰汾
    Where ��¼id = n_��¼id And ((Mod(��¼����, 10) <> 5 And ��Ŀ��� = ��Ŀ���_In And Nvl(���²�λ, 'TWBW') = Nvl(���²�λ_In, 'TWBW') And
          Nvl(��¼���, 0) = Nvl(��¼���_In, 0)) Or ��¼���� = Decode(��ǩ_In, 1, 15, 5)) And ��ʼ�汾 <= n_��߰汾 - 1 And ��ֹ�汾 Is Null;
  
    --�����δǩ�����ݣ�����޸Ĳ���Ա��Ϊ�ü�¼�ı����˸���
    If n_��߰汾 = 1 Then
      Update ���˻������� Set ������ = v_������, ����ʱ�� = Sysdate Where ID = n_��¼id;
    End If;
  
    If Nvl(n_�������, 0) <> 0 and  �������_In=0  Then
      Return;
    End If;
  
    --############
    --ͬ����������
    --############
    --1\�ȴ������µ���һ������ʼ��ֻ����һ����Ч�����µ��ļ���
    --������±������ͬ����ʱ������ݣ�ʹ������ID
    --CL,2015-12-30,��¼��ͬ��������Ŀ�����µ�
    For Row_Format In Cur_Fileformats Loop
      If Row_Format.���� = -1 Then
        If Row_Format.���� = '1' Then
          Begin
            Select 1, h.��Ŀ����
            Into Intins, n_��Ŀ����
            From (Select To_Char(f.��Ŀ���) As ���, g.��Ŀ����
                   From ���¼�¼��Ŀ F, �����¼��Ŀ G
                   Where f.��Ŀ��� = g.��Ŀ��� And g.��Ŀ���� = 2 And
                         (g.���ÿ��� = 1 Or
                         (g.���ÿ��� = 2 And Exists
                          (Select 1 From �������ÿ��� D Where g.��Ŀ��� = d.��Ŀ��� And d.����id = v_����id))) And Nvl(g.Ӧ�÷�ʽ, 0) <> 0 And
                         (Nvl(g.���ò���, 0) = 0 Or Nvl(g.���ò���, 0) = Decode(Nvl(Row_Format.Ӥ��, 0), 0, 1, 2))
                   Union All
                   Select b.�����ı� As ���, 1 As ��Ŀ����
                   From �����ļ��ṹ A, �����ļ��ṹ B
                   Where a.�ļ�id = Row_Format.��ʽid And a.��id Is Null And a.������� In (2, 3) And b.��id = a.Id) H
            Where Instr(',' || h.��� || ',', ',' || ��Ŀ���_In || ',', 1) > 0;
          Exception
            When Others Then
              Intins := 0;
          End;
        Else
          Begin
            Select 1, g.��Ŀ����
            Into Intins, n_��Ŀ����
            From ���¼�¼��Ŀ F, �����¼��Ŀ G
            Where f.��Ŀ��� = g.��Ŀ��� And Nvl(g.Ӧ�÷�ʽ, 0) <> 0 And g.����ȼ� >= 0 And
                  (Nvl(g.���ò���, 0) = 0 Or Nvl(g.���ò���, 0) = Decode(Nvl(Row_Format.Ӥ��, 0), 0, 1, 2)) And f.��Ŀ��� = ��Ŀ���_In And
                  (g.���ÿ��� = 1 Or (g.���ÿ��� = 2 And Exists
                   (Select 1 From �������ÿ��� D Where g.��Ŀ��� = d.��Ŀ��� And d.����id = v_����id)));
          Exception
            When Others Then
              Intins := 0;
          End;
        End If;
      
        If Intins > 0 Then
          --LPF,2013-01-23,������Ŀ�Ƿ���Ҫ����ͬ��(������ǰ�Ѿ�ͬ���������ݣ�Ϊ�˱�֤��¼�������µ�����һֱ�������ݴ˺����жϡ�)
          n_Synchro := Zl_Temperatureprogram(�ļ�id_In, v_����id, ��Ŀ���_In, ����ʱ��_In);
          Begin
            Select b.Id
            Into n_Newid
            From ���˻����ļ� A, ���˻������� B
            Where a.Id = Row_Format.�ļ�id And b.�ļ�id = a.Id And b.����ʱ�� = ����ʱ��_In;
          Exception
            When Others Then
              n_Newid := 0;
          End;
          n_Oldid := n_Newid;
          If n_Newid = 0 And n_Synchro = 1 Then
            Select ���˻�������_Id.Nextval Into n_Newid From Dual;
            --�������µ�����¼
            Insert Into ���˻�������
              (ID, �ļ�id, ������, ����ʱ��, ����ʱ��, ���汾)
            Values
              (n_Newid, Row_Format.�ļ�id, v_������, Sysdate, ����ʱ��_In, 1);
          End If;
        
          If n_Newid > 0 Then
            --����δͬ�������µ�����(��ȻҪ���Ӷ���ѯ)
            Select Count(*)
            Into v_������Դ
            From ���˻�����ϸ
            Where ��¼id = n_Newid And ��Ŀ��� = ��Ŀ���_In And
                  Decode(n_��Ŀ����, 2, Nvl(���²�λ, '��'), Nvl(���²�λ_In, '��')) = Nvl(���²�λ_In, '��');
            If v_������Դ = 0 Then
              --˵����ͬ����ʼ�Ѿ����й����
              If n_Synchro = 1 Then
                --û�м�����Ŀ�Ƿ���Ҫͬ��
                Insert Into ���˻�����ϸ
                  (ID, ��¼id, ��¼����, ��Ŀ����, ��Ŀid, ��Ŀ���, ��Ŀ����, ��Ŀ����, ��¼����, ��Ŀ��λ, ��¼���, ���²�λ, ������Դ, ��Դid, ��ʼ�汾, ��ֹ�汾, ��¼��,
                   ��¼ʱ��, ��¼���)
                  Select ���˻�����ϸ_Id.Nextval, n_Newid, b.��¼����, b.��Ŀ����, b.��Ŀid, b.��Ŀ���, b.��Ŀ����, b.��Ŀ����, b.��¼����, b.��Ŀ��λ,
                         b.��¼���, b.���²�λ, 1, b.Id, 1, Null, b.��¼��, Sysdate, 1
                  From (Select ��Ŀ���_In As ��Ŀ���, Nvl(���²�λ_In, '��') As ���²�λ
                         From Dual
                         Minus
                         Select f.��Ŀ���, Decode(Nvl(f.��Ŀ����, 1), 2, Nvl(���²�λ, '��'), Nvl(���²�λ_In, '��'))
                         From ���˻�����ϸ E, �����¼��Ŀ F
                         Where e.��¼id = n_Newid And e.��Ŀ��� = f.��Ŀ���) A, ���˻�����ϸ B
                  Where a.��Ŀ��� = b.��Ŀ��� And b.��¼id = n_��¼id And b.Id = n_��ϸid;
                If Sql%RowCount > 0 Then
                  Int���� := 1;
                End If;
              End If;
            Else
              Update ���˻�����ϸ
              Set ��¼���� = ��¼����_In, ��Դid = n_��ϸid
              Where ��¼id = n_Newid And ��Ŀ��� = ��Ŀ���_In And
                    Decode(n_��Ŀ����, 2, Nvl(���²�λ, '��'), Nvl(���²�λ_In, '��')) = Nvl(���²�λ_In, '��') And ������Դ > 0;
              If Sql%RowCount > 0 Then
                Int���� := 1;
              End If;
            End If;
          End If;
        End If;
        --2\��ѭ�������¼��
      Else
        If n_Mutilbill = 1 And n_Syntend = 1 Then
          --��ȡ��¼���뵱ǰ��¼�������ص����������ݵĹ̶���Ŀ
          Select Count(*)
          Into Intins
          From (Select b.��Ŀ���
                 From �����ļ��ṹ A, �����¼��Ŀ B
                 Where a.Ҫ������ = b.��Ŀ���� And b.��Ŀ��ʾ In (0, 4, 5) And
                       ��id =
                       (Select ID From �����ļ��ṹ Where �ļ�id = Row_Format.��ʽid And ��id Is Null And ������� = 4)
                 Intersect
                 Select b.��Ŀ���
                 From �����ļ��ṹ A, �����¼��Ŀ B, ���˻����ļ� C, ���˻������� D, ���˻�����ϸ G
                 Where c.Id = d.�ļ�id And a.�ļ�id = c.��ʽid And d.Id = g.��¼id And d.Id = n_��¼id And g.Id = n_��ϸid And
                       b.��Ŀ��� = g.��Ŀ��� And b.��Ŀ��ʾ In (0, 4, 5) And g.��¼���� = 1 And a.Ҫ������ = b.��Ŀ���� And
                       a.��id = (Select ID From �����ļ��ṹ E Where e.�ļ�id = c.��ʽid And ��id Is Null And ������� = 4));
        
          If Intins > 0 Then
            n_Newid := 0;
            --����ָ���ļ��Ѿ�������ͬ����ʱ������ݣ�ֱ��������ID����
            Begin
              Select c.Id
              Into n_Newid
              From ���˻������� C
              Where c.�ļ�id = Row_Format.�ļ�id And c.����ʱ�� = ����ʱ��_In;
            Exception
              When Others Then
                n_Newid := 0;
            End;
          
            If n_Newid = 0 Then
              --������¼������¼
              Select ���˻�������_Id.Nextval Into n_Newid From Dual;
            
              Insert Into ���˻�������
                (ID, �ļ�id, ������, ����ʱ��, ����ʱ��, ���汾)
                Select n_Newid, Row_Format.�ļ�id, c.������, c.����ʱ��, c.����ʱ��, 1
                From ���˻������� C
                Where c.Id = n_��¼id;
            End If;
          
            If n_Newid > 0 Then
              --����δͬ���ļ�¼������
              Select Count(*) Into v_������Դ From ���˻�����ϸ Where ��¼id = n_Newid And ��Ŀ��� = ��Ŀ���_In;
              If v_������Դ = 0 Then
                Insert Into ���˻�����ϸ
                  (ID, ��¼id, ��¼����, ��Ŀ����, ��Ŀid, ��Ŀ���, ��Ŀ����, ��Ŀ����, ��¼����, ��Ŀ��λ, ��¼���, ���²�λ, ������Դ, ��Դid, δ��˵��, ��ʼ�汾, ��ֹ�汾,
                   ��¼��, ��¼ʱ��)
                  Select ���˻�����ϸ_Id.Nextval, n_Newid, b.��¼����, b.��Ŀ����, b.��Ŀid, b.��Ŀ���, b.��Ŀ����, b.��Ŀ����, b.��¼����, b.��Ŀ��λ,
                         b.��¼���, b.���²�λ, 1, b.Id, b.δ��˵��, 1, Null, b.��¼��, Sysdate
                  From (Select b.��Ŀ���
                         From �����ļ��ṹ A, �����¼��Ŀ B
                         Where a.Ҫ������ = b.��Ŀ���� And b.��Ŀ��ʾ In (0, 4, 5) And
                               ��id = (Select ID
                                      From �����ļ��ṹ
                                      Where �ļ�id = Row_Format.��ʽid And ��id Is Null And ������� = 4)
                         Intersect
                         Select b.��Ŀ���
                         From �����ļ��ṹ A, �����¼��Ŀ B, ���˻����ļ� C, ���˻������� D, ���˻�����ϸ G
                         Where c.Id = d.�ļ�id And a.�ļ�id = c.��ʽid And d.Id = g.��¼id And d.Id = n_��¼id And g.Id = n_��ϸid And
                               b.��Ŀ��� = g.��Ŀ��� And b.��Ŀ��ʾ In (0, 4, 5) And g.��¼���� = 1 And a.Ҫ������ = b.��Ŀ���� And
                               a.��id =
                               (Select ID From �����ļ��ṹ E Where e.�ļ�id = c.��ʽid And ��id Is Null And ������� = 4)) A, ���˻�����ϸ B
                  Where a.��Ŀ��� = b.��Ŀ��� And b.��¼id = n_��¼id And b.Id = n_��ϸid;
                If Sql%RowCount > 0 Then
                  Int���� := 1;
                  --ԭ������Ҫ��
                  Begin
                    Select ���� Into n_���� From ���˻����ӡ Where �ļ�id = Row_Format.�ļ�id And ��¼id = n_Newid;
                  Exception
                    When Others Then
                      n_���� := 1;
                  End;
                  Zl_���˻����ӡ_Update(Row_Format.�ļ�id, ����ʱ��_In, n_����, 0);
                End If;
              Else
                Update ���˻�����ϸ
                Set ��¼���� = ��¼����_In, δ��˵�� = δ��˵��_In, ��Դid = n_��ϸid
                Where ��¼id = n_Newid And ��Ŀ��� = ��Ŀ���_In And (������Դ > 0 Or ������Դ <> 3);
                If Sql%RowCount > 0 Then
                  Int���� := 1;
                End If;
              End If;
            End If;
          End If;
        End If;
      End If;
    End Loop;
  
    If Int���� = 1 Then
      Update ���˻�����ϸ Set ���� = 1 Where ID = n_��ϸid;
      --����ʷ���ݵĹ��ñ�־����ΪNULL
      Update ���˻�����ϸ Set ���� = Null Where ��¼id = n_��¼id And ��Ŀ��� = ��Ŀ���_In And ID <> n_��ϸid;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˻�������_Update;
/

--126662:����,2018-06-11,��ʷ�޸İ�֮ǰ�Ĺ��̸�����
Create Or Replace Procedure Zl_������Ŀ_Insert
(
  ���_In             In ������ĿĿ¼.���%Type := Null,
  ����id_In           In ������ĿĿ¼.����id%Type := Null,
  Id_In               In ������ĿĿ¼.Id%Type,
  ����_In             In ������ĿĿ¼.����%Type := Null,
  ����_In             In ������ĿĿ¼.����%Type := Null,
  ����ƴ��_In         In ������Ŀ����.����%Type := Null,
  �������_In         In ������Ŀ����.����%Type := Null,
  ����_In             ������ĿĿ¼.����%Type := Null,
  ����ƴ��_In         ������Ŀ����.����%Type := Null,
  �������_In         ������Ŀ����.����%Type := Null,
  ��������_In         In ������ĿĿ¼.��������%Type := Null,
  ִ��Ƶ��_In         In ������ĿĿ¼.ִ��Ƶ��%Type := Null,
  ����Ӧ��_In         In ������ĿĿ¼.����Ӧ��%Type := Null,
  ���㷽ʽ_In         In ������ĿĿ¼.���㷽ʽ%Type := Null,
  ���㵥λ_In         In ������ĿĿ¼.���㵥λ%Type := Null,
  �����Ա�_In         In ������ĿĿ¼.�����Ա�%Type := Null,
  ִ�а���_In         In ������ĿĿ¼.ִ�а���%Type := Null,
  �������_In         In ������ĿĿ¼.�������%Type := Null,
  �����Ŀ_In         In ������ĿĿ¼.�����Ŀ%Type := Null,
  �걾��λ_In         In ������ĿĿ¼.�걾��λ%Type := Null,
  ��������id_In       In ������϶���.����id%Type := Null,
  ִ�п���_In         In ������ĿĿ¼.ִ�п���%Type := Null,
  ����ִ��_In         In ����ִ�п���.ִ�п���id%Type := Null,
  סԺִ��_In         In ����ִ�п���.ִ�п���id%Type := Null,
  ����ִ��_In         In Varchar2, --�������Ҷ���ִ�е�˵��������'|'�ָÿ������'��������id^ִ�п���id'��ʽ��֯
  �ο�Ŀ¼id_In       In ������ĿĿ¼.�ο�Ŀ¼id%Type := Null,
  Ӧ�÷�Χ_In         In Number := 0,
  ¼������_In         In ������ĿĿ¼.¼������%Type := Null,
  ������Χ_In         In Number := 0,
  ִ�б��_In         In Number := 0,
  ִ�з���_In         In ������ĿĿ¼.ִ�з���%Type := 0,
  վ��_In             In ������ĿĿ¼.վ��%Type := Null,
  ��ĿƵ��_In         In Varchar2 := Null, --����Ŀ��Ƶ�����ô�������|����......
  �������_In         In ������ĿĿ¼.�������%Type := Null,
  ʹ�ÿ���_In         In Varchar2 := Null, --ʹ�ÿ��ҵ�IDs,�ö��ŷָ�
  ʹ�ÿ���Ӧ�÷�Χ_In In Number := 0, --ʹ�ÿ���Ӧ�õķ�Χ  0-���1-Ӧ����ͬ����2-���������У�3-Ӧ���ڵ�ǰ���
  First_In            In Number := 1, --First��1-��Ҫɾ��ִ�п��ң���������0-��ɾ��ִ�п��ң�ֱ������
  ����ϵ��_In         In ������ĿĿ¼.����ϵ��%Type := Null,
  ��Ѫ�������_In     In Varchar2 :=Null,
  ԭʼid_IN           In ������ĿĿ¼.Id%Type:=0,
  �Թܱ���_In         In ������ĿĿ¼.�Թܱ���%Type := Null  
) Is
  Type t_������Ŀ Is Ref Cursor;
  c_������Ŀ   t_������Ŀ;
  t_Id         t_Numlist;
  v_Id         ������ĿĿ¼.Id%Type;
  v_Records    Varchar2(4000); --��ʱ��¼�������Ҷ���ִ�п��ҵ��ַ���
  v_Currrec    Varchar2(1000); --�����ڶ���ִ�п����ַ����е�һ������
  v_Fields     Varchar2(1000);
  v_��������id ����ִ�п���.��������id%Type := Null;
  v_ִ�п���id ����ִ�п���.ִ�п���id%Type := Null;
  n_���       Number;
  v_���       Varchar2(1000);
  v_Strtmp     Varchar2(1000);
  v_Strinput   Varchar2(1000);
Begin
  If First_In = 1 Then
    Insert Into ������ĿĿ¼
      (���, ����id, ID, ����, ����, ��������, ִ��Ƶ��, ����Ӧ��, ���㷽ʽ, ���㵥λ, �����Ա�, ִ�а���, �������, ִ�п���, �����Ŀ, �걾��λ, ����ʱ��, ����ʱ��, �ο�Ŀ¼id, ¼������,
       ִ�б��, ִ�з���, �������, վ��, ����ϵ��,�Թܱ���)
    Values
      (���_In, ����id_In, Id_In, ����_In, ����_In, ��������_In, ִ��Ƶ��_In, ����Ӧ��_In, ���㷽ʽ_In, ���㵥λ_In, �����Ա�_In, ִ�а���_In, �������_In,
       ִ�п���_In, �����Ŀ_In, Decode(���_In, 'D', Decode(�����Ŀ_In, 1, '', �걾��λ_In), �걾��λ_In), Sysdate,
       To_Date('3000-01-01', 'YYYY-MM-DD'), �ο�Ŀ¼id_In, ¼������_In, ִ�б��_In, ִ�з���_In, �������_In, վ��_In, ����ϵ��_In,�Թܱ���_In);
    If ��������id_In Is Not Null Then
      Insert Into ������϶��� (����id, ���id, ����id) Values (��������id_In, Null, Id_In);
    End If;
    If ����ƴ��_In Is Not Null Then
      Insert Into ������Ŀ���� (������Ŀid, ����, ����, ����, ����) Values (Id_In, ����_In, 1, ����ƴ��_In, 1);
    End If;
    If �������_In Is Not Null Then
      Insert Into ������Ŀ���� (������Ŀid, ����, ����, ����, ����) Values (Id_In, ����_In, 1, �������_In, 2);
    End If;
    If ����_In Is Not Null And ����ƴ��_In Is Not Null Then
      Insert Into ������Ŀ���� (������Ŀid, ����, ����, ����, ����) Values (Id_In, ����_In, 9, ����ƴ��_In, 1);
    End If;
    If ����_In Is Not Null And �������_In Is Not Null Then
      Insert Into ������Ŀ���� (������Ŀid, ����, ����, ����, ����) Values (Id_In, ����_In, 9, �������_In, 2);
    End If;
  End If;
  If Ӧ�÷�Χ_In = 1 Then
    If ����id_In Is Null Then
      Open c_������Ŀ For
        Select ID From ������ĿĿ¼ Where ����id Is Null Order By ����;
    Else
      Open c_������Ŀ For
        Select ID From ������ĿĿ¼ Where ����id = ����id_In Order By ����;
    End If;
  Elsif Ӧ�÷�Χ_In = 2 Then
    If ����id_In Is Null Then
      Open c_������Ŀ For
        Select c.Id
        From ������ĿĿ¼ C, (Select ID From ���Ʒ���Ŀ¼ Start With �ϼ�id Is Null Connect By Prior ID = �ϼ�id) D
        Where d.Id = c.����id
        Order By ����;
    Else
      Open c_������Ŀ For
        Select c.Id
        From ������ĿĿ¼ C, (Select ID From ���Ʒ���Ŀ¼ Start With ID = ����id_In Connect By Prior ID = �ϼ�id) D
        Where d.Id = c.����id
        Order By ����;
    End If;
  Elsif Ӧ�÷�Χ_In = 3 Then
    Open c_������Ŀ For
      Select ID From ������ĿĿ¼ Where ��� = ���_In Order By ����;
  Else
    Open c_������Ŀ For
      Select ID From ������ĿĿ¼ Where ID = Id_In;
  End If;

  Loop
    Fetch c_������Ŀ
      Into v_Id;
    Exit When c_������Ŀ%NotFound;
  
    If First_In = 1 Then
      Delete From ����ִ�п��� Where ������Ŀid = v_Id;
      If ִ�п���_In = 4 And ����ִ��_In Is Not Null Then
        Insert Into ����ִ�п��� (������Ŀid, ������Դ, ��������id, ִ�п���id) Values (v_Id, 1, Null, ����ִ��_In);
      End If;
      If ִ�п���_In = 4 And סԺִ��_In Is Not Null Then
        Insert Into ����ִ�п��� (������Ŀid, ������Դ, ��������id, ִ�п���id) Values (v_Id, 2, Null, סԺִ��_In);
      End If;
    End If;
    If ִ�п���_In <> 4 Or ����ִ��_In Is Null Then
      v_Records := Null;
    Else
      v_Records := ����ִ��_In || '|';
    End If;
  
    While v_Records Is Not Null Loop
      v_Currrec    := Substr(v_Records, 1, Instr(v_Records, '|') - 1);
      v_Fields     := v_Currrec;
      v_��������id := To_Number(Substr(v_Fields, 1, Instr(v_Fields, '^') - 1));
      v_Fields     := Substr(v_Fields, Instr(v_Fields, '^') + 1);
      v_ִ�п���id := To_Number(v_Fields);
      Insert Into ����ִ�п���
        (������Ŀid, ������Դ, ��������id, ִ�п���id)
      Values
        (v_Id, Null, Decode(v_��������id, 0, Null, v_��������id), v_ִ�п���id);
      v_Records := Replace('|' || v_Records, '|' || v_Currrec || '|');
    End Loop;
    If Ӧ�÷�Χ_In <> 0 Then
      Update ������ĿĿ¼ Set ִ�п��� = ִ�п���_In Where ID = v_Id;
    End If;
  End Loop;
  Close c_������Ŀ;

  If First_In = 1 Then
    If ���_In = 'C' Or ���_In = 'F' Or ���_In = 'K' Then
      Insert Into ��������Ӧ��
        (�����ļ�id, Ӧ�ó���, ������Ŀid)
        Select a.�����ļ�id, 1, Id_In
        From ��������Ӧ�� A, ������ĿĿ¼ I
        Where a.������Ŀid = i.Id And i.��� = ���_In And Ӧ�ó��� = 1 And (�������_In = 0 Or �������_In = 1) And Rownum < 2;
      Insert Into ��������Ӧ��
        (�����ļ�id, Ӧ�ó���, ������Ŀid)
        Select a.�����ļ�id, 2, Id_In
        From ��������Ӧ�� A, ������ĿĿ¼ I
        Where a.������Ŀid = i.Id And i.��� = ���_In And Ӧ�ó��� = 2 And (�������_In = 0 Or �������_In = 2) And Rownum < 2;
    Elsif ���_In = 'D' Or ���_In = 'E' Then
      Insert Into ��������Ӧ��
        (�����ļ�id, Ӧ�ó���, ������Ŀid)
        Select a.�����ļ�id, 1, Id_In
        From ��������Ӧ�� A, ������ĿĿ¼ I
        Where a.������Ŀid = i.Id And i.��� = ���_In And �������� = ��������_In And Ӧ�ó��� = 1 And (�������_In = 0 Or �������_In = 1) And
              Rownum < 2;
      Insert Into ��������Ӧ��
        (�����ļ�id, Ӧ�ó���, ������Ŀid)
        Select a.�����ļ�id, 2, Id_In
        From ��������Ӧ�� A, ������ĿĿ¼ I
        Where a.������Ŀid = i.Id And i.��� = ���_In And �������� = ��������_In And Ӧ�ó��� = 2 And (�������_In = 0 Or �������_In = 2) And
              Rownum < 2;
    End If;
  End If;

  If ������Χ_In = 1 Then
    If ����id_In Is Null Then
      Update ������ĿĿ¼ Set ¼������ = ¼������_In Where ����id Is Null;
    Else
      Update ������ĿĿ¼ Set ¼������ = ¼������_In Where ����id = ����id_In;
    End If;
  Elsif ������Χ_In = 2 Then
    If ����id_In Is Null Then
      Update ������ĿĿ¼
      Set ¼������ = ¼������_In
      Where ����id In (Select ID From ���Ʒ���Ŀ¼ Start With �ϼ�id Is Null Connect By Prior ID = �ϼ�id);
    Else
      Update ������ĿĿ¼
      Set ¼������ = ¼������_In
      Where ����id In (Select ID From ���Ʒ���Ŀ¼ Start With ID = ����id_In Connect By Prior ID = �ϼ�id);
    End If;
  Elsif ������Χ_In = 3 Then
    Update ������ĿĿ¼ Set ¼������ = ¼������_In Where ��� = ���_In;
  Elsif ������Χ_In = 4 Then
    Update ������ĿĿ¼ Set ¼������ = ¼������_In;
  End If;

  --����Ŀ��Ƶ������
  If ���_In <> 'C' Then
    Delete �����÷����� Where ��Ŀid = Id_In;
    If ��ĿƵ��_In Is Not Null Then
      v_Strinput := ��ĿƵ��_In || '|';
      n_���     := 0;
    
      While v_Strinput Is Not Null Loop
        v_Strtmp := Substr(v_Strinput, 1, Instr(v_Strinput, '|') - 1);
        v_���   := v_Strtmp;
        n_���   := n_��� + 1;
      
        Insert Into �����÷����� (��Ŀid, ����, Ƶ��) Values (Id_In, n_���, v_���);
        v_Strinput := Replace('|' || v_Strinput, '|' || v_Strtmp || '|');
      End Loop;
    End If;
  End If;
  --ʹ�ÿ���
  If ʹ�ÿ���Ӧ�÷�Χ_In = 1 Then
    If ����id_In Is Null Then
      Open c_������Ŀ For
        Select ID From ������ĿĿ¼ Where ����id Is Null Order By ����;
    Else
      Open c_������Ŀ For
        Select ID From ������ĿĿ¼ Where ����id = ����id_In Order By ����;
    End If;
  Elsif ʹ�ÿ���Ӧ�÷�Χ_In = 2 Then
    If ����id_In Is Null Then
      Open c_������Ŀ For
        Select c.Id
        From ������ĿĿ¼ C, (Select ID From ���Ʒ���Ŀ¼ Start With �ϼ�id Is Null Connect By Prior ID = �ϼ�id) D
        Where d.Id = c.����id
        Order By ����;
    Else
      Open c_������Ŀ For
        Select c.Id
        From ������ĿĿ¼ C, (Select ID From ���Ʒ���Ŀ¼ Start With ID = ����id_In Connect By Prior ID = �ϼ�id) D
        Where d.Id = c.����id
        Order By ����;
    End If;
  Elsif ʹ�ÿ���Ӧ�÷�Χ_In = 3 Then
    Open c_������Ŀ For
      Select ID From ������ĿĿ¼ Where ��� = ���_In Order By ����;
  Else
    Open c_������Ŀ For
      Select ID From ������ĿĿ¼ Where ID = Id_In;
  End If;
  Fetch c_������Ŀ Bulk Collect
    Into t_Id;
  Close c_������Ŀ;

  Forall I In 1 .. t_Id.Count
    Delete �������ÿ��� Where ��Ŀid = t_Id(I) And Instr(',' || ʹ�ÿ���_In || ',', ',' || ����id || ',') = 0;

  If ʹ�ÿ���_In Is Not Null Then
    Forall I In 1 .. t_Id.Count
      Insert Into �������ÿ���
        (��Ŀid, ����id)
        Select t_Id(I), Column_Value
        From Table(f_Num2list(ʹ�ÿ���_In)) A
        Where Not Exists (Select 1 From �������ÿ��� Where ����id = Column_Value And ��Ŀid = t_Id(I));
  End If;
  --��Ѫ�������
  If ���_In = 'K' And ��Ѫ�������_In Is Not Null Then
    v_Strinput := ��Ѫ�������_In || '|';
  
    While v_Strinput Is Not Null Loop
      v_Strtmp := Substr(v_Strinput, 1, Instr(v_Strinput, '|') - 1);
      v_Id     := v_Strtmp;
    
      Insert Into ��Ѫ������� (��Ŀid, ������Ŀid) Values (Id_In, v_Id);
      v_Strinput := Replace('|' || v_Strinput, '|' || v_Strtmp || '|');
    End Loop;
  End If;
  
  if ԭʼid_IN<>0 then
    Zl_�����շ�_Insert(id_In,ԭʼid_IN);
  end if;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_������Ŀ_Insert;
/
--125837:Ϳ����,2018-06-07,�Ŷ���ʾ��ʽ����
CREATE OR REPLACE Procedure Zl_�ŶӽкŶ���_���ҵ��
(
       ҵ������_IN �ŶӽкŶ���.ҵ������%Type,
       ��Ч����_IN Number := 1
)
Is
Begin
  case ҵ������_IN
    when -1 then Null;
    else
      --�����ǰҵ�����ͣ�����ʱ������Чʱ��֮ǰ���Ŷ���Ϣ
      delete from �Ŷ��������� where ҵ������ = ҵ������_IN And ����ʱ�� <=  sysdate - (1 / 48);
     
      Delete From �ŶӽкŶ��� 
      Where ҵ������ = ҵ������_IN And To_Number(Trunc(Sysdate - �ŶӽкŶ���.�Ŷ�ʱ��)) >= ��Ч����_In;
  end case;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ŶӽкŶ���_���ҵ��;
/

--126057:����,2018-06-07,Zl_ҩƷ�շ�����_Insert�޸�
Create Or Replace Procedure Zl_ҩƷ�շ�����_Insert
(
  No_In     In ҩƷ�շ�����.No%Type,
  ����_In   In ҩƷ�շ�����.����%Type,
  �ⷿid_In In ҩƷ�շ�����.�ⷿid%Type,
  �Է�����id_In In ҩƷ�շ�����.�Է�����id%Type
) Is
  n_Count Number;
  n_Id    ҩƷ�շ�����.Id%Type;
Begin
  n_Count := 0;
  Begin
    Select 1 Into n_Count From ҩƷ�շ����� Where NO = No_In And ���� = ����_In And �ⷿid = �ⷿid_In And �Է�����id = �Է�����id_In;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If n_Count = 0 Then
    Select ҩƷ�շ�����_Id.Nextval Into n_Id From Dual;
    Insert Into ҩƷ�շ����� (ID, NO, ����, �ⷿid, ��ӡ״̬, �Է�����id) Values (n_Id, No_In, ����_In, �ⷿid_In, 1, �Է�����id_In);
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�շ�����_Insert;
/

--126591:���˺�,2018-06-04,�������ѵĴ���.
Create Or Replace Procedure Zl_Third_Settlement
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is

  --------------------------------------------------------------------------------------------------
  --����:�����ӿڽ���
  --���:Xml_In:
  --<IN>
  --        <BRID>����ID</BRID>         //����ID
  --        <XM>����</XM>               //����
  --        <SFZH>���֤��</SFZH>       //���֤��
  --        <ZYID>��ҳID</ZYID>         //��ҳID
  --        <JSLX>2</JSLX>         //��������,1-����,2-סԺ.Ŀǰ�̶���2
  --        <JE></JE>         //���ν����ܽ��
  --       <JSLIST>
  --         <JS>
  --              <JSKLB>֧�������</JSKLB >
  --              <JSKH>֧������</ JSKH >
  --              <JSFS>֧����ʽ</JSFS> //֧����ʽ:�ֽ�;֧Ʊ,�����������,���Դ���
  --              <JSJE>������</JSJE> //������(�������˲������ҽԺ�˿�)<SFCYJ>Ϊ1ʱΪ��Ԥ�����
  --              <JYLSH>������ˮ��</JYLSH>
  --              <ZY>ժҪ</ZY>
  --              <SFCYJ>�Ƿ��Ԥ��</SFCYJ>  //�Ƿ��Ԥ����0-���㣬1-��Ԥ��.�ʳ�Ԥ��ʱ,ֻ��JSJE�ڵ�
  --              <SFXFK>�Ƿ����ѿ�</SFXFK>  //(1-�����ѿ�),���ѿ�ʱ,������㿨���,���㿨��,������Ƚӵ�
  --              <EXPENDLIST>  //��չ������Ϣ
  --                  <EXPEND>
  --                        <JYMC>��������</JYMC> //��������   �˿�ʱ,�����Ԥ������ˮ��
  --                        <JYLR>��������</JYLR> //��������   �˿�ʱ,�����Ԥ���Ľ��
  --                  </EXPEND>
  --              </EXPENDLIST>
  --         </JS>
  --       </JSLIST >
  --</IN>

  --����:Xml_Out
  --  <OUT>
  --       <CZSJ>����ʱ��</CZSJ>          //HIS�ĵǼ�ʱ��
  --    �D�D�������д�������˵����ȷִ��
  --    <ERROR>
  --      <MSG>������Ϣ</MSG>
  --    </ERROR>
  --  </OUT>
  --------------------------------------------------------------------------------------------------
  n_��ҳid     ������ҳ.��ҳid%Type;
  n_����id     ������ҳ.����id%Type;
  v_����       ������Ϣ.����%Type;
  v_���֤��   ������Ϣ.���֤��%Type;
  n_�����ܶ�   ����Ԥ����¼.��Ԥ��%Type;
  n_�����ʽ�� ����Ԥ����¼.��Ԥ��%Type;
  n_��������   Number(3);
  v_����Ա���� ���˽��ʼ�¼.����Ա���%Type;
  v_����Ա���� ���˽��ʼ�¼.����Ա����%Type;
  n_����id     ���˽��ʼ�¼.Id%Type;
  n_��Ԥ����� ����Ԥ����¼.��Ԥ��%Type;
  d_����ʱ��   Date;
  n_Ԥ����ֵ   ����Ԥ����¼.���%Type;
  d_��ʼ����   Date;
  d_��������   Date;
  d_��С����   Date;
  d_�������   Date;

  n_Ԥ��id     ����Ԥ����¼.Id%Type;
  n_����id     ������ҳ.��Ժ����id%Type;
  n_���㿨��� �����ѽӿ�Ŀ¼.���%Type;
  n_ʱ������   Number(3);
  v_Ids        Varchar2(20000);
  v_No         ���˽��ʼ�¼.No%Type;
  v_Ԥ��no     ����Ԥ����¼.No%Type;
  n_�����id   ҽ�ƿ����.Id%Type;
  v_���㿨��   ����Ԥ����¼.����%Type;
  v_���㷽ʽ   ����Ԥ����¼.���㷽ʽ%Type;
  v_Temp       Varchar2(500);
  x_Templet    Xmltype; --ģ��XML
  v_Err_Msg    Varchar2(200);
  Err_Item    Exception;
  Err_Special Exception;

  n_Count    Number(18);
  n_Number   Number(2);
  n_����id   ������ü�¼.Id%Type;
  n_��¼���� ������ü�¼.��¼����%Type;
  v_����no   ������ü�¼.No%Type;
  n_���     ������ü�¼.���%Type;
  n_��¼״̬ ������ü�¼.��¼״̬%Type;
  n_ִ��״̬ ������ü�¼.ִ��״̬%Type;
  n_δ���� ������ü�¼.ʵ�ս��%Type;
  n_���ʽ�� ������ü�¼.ʵ�ս��%Type;
  n_����   ������ü�¼.ʵ�ս��%Type;

  v_�����     �������׼�¼.���%Type;
  v_���ѿ����� Varchar2(20000);

  Type t_���ý�����ϸ Is Ref Cursor;
  c_���ý�����ϸ t_���ý�����ϸ;

Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/ZYID'), To_Number(Extractvalue(Value(A), 'IN/BRID')),
         To_Number(Extractvalue(Value(A), 'IN/JE')), To_Number(Extractvalue(Value(A), 'IN/JSLX')),
         Extractvalue(Value(A), 'IN/SFZH'), Extractvalue(Value(A), 'IN/XM')
  Into n_��ҳid, n_����id, n_�����ܶ�, n_��������, v_���֤��, v_����
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  n_�������� := Nvl(n_��������, 2);
  If n_�������� = 1 And Nvl(n_����id, 0) = 0 And Not v_���֤�� Is Null And Not v_���� Is Null Then
    n_����id := Zl_Third_Getpatiid(v_���֤��, v_����);
  End If;
  --0.��ؼ��
  If Nvl(n_����id, 0) = 0 Then
    v_Err_Msg := '������Чʶ�������,������ɷ�!';
    Raise Err_Item;
  End If;

  --��Աid,��Ա���,��Ա����
  v_Temp := Zl_Identity(1);
  If Nvl(v_Temp, '0') = '0' Or Nvl(v_Temp, '_') = '_' Then
    v_Err_Msg := 'ϵͳ�����ϱ���Ч�Ĳ���Ա,���������!';
    Raise Err_Item;
  End If;
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_����Ա���� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_����Ա���� := v_Temp;
  v_Err_Msg    := Null;

  For c_���׼�¼ In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As ���㿨���,
                        Extractvalue(b.Column_Value, '/JS/JSFS') As ���㷽ʽ,
                        Extractvalue(b.Column_Value, '/JS/JYLSH') As ������ˮ��,
                        Extractvalue(b.Column_Value, '/JS/JSKH') As ���㿨��, Extractvalue(b.Column_Value, '/JS/ZY') As ժҪ,
                        Extractvalue(b.Column_Value, '/JS/SFCYJ') As �Ƿ��Ԥ��,
                        Extractvalue(b.Column_Value, '/JS/SFXFK') As �Ƿ����ѿ�
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
  
    If Not (c_���׼�¼.���㿨��� Is Null Or Nvl(c_���׼�¼.�Ƿ����ѿ�, '0') = '1' Or Nvl(c_���׼�¼.�Ƿ��Ԥ��, 0) = 1) Then
    
      Select Decode(Translate(Nvl(c_���׼�¼.���㿨���, 'abcd'), '#1234567890', '#'), Null, 1, 0)
      Into n_Count
      From Dual;
    
      If Nvl(n_Count, 0) = 1 Then
        Select Max(����) Into v_����� From ҽ�ƿ���� Where ID = To_Number(c_���׼�¼.���㿨���);
      Else
        Select Max(����) Into v_����� From ҽ�ƿ���� Where ���� = c_���׼�¼.���㿨���;
      End If;
    
      If v_����� Is Null Then
        v_Err_Msg := '��֧�ֵĽ��㷽ʽ,���飡';
        Raise Err_Item;
      End If;
    
      If Zl_Fun_�������׼�¼_Locked(v_�����, c_���׼�¼.������ˮ��, c_���׼�¼.���㿨��, c_���׼�¼.ժҪ, 2) = 0 Then
        v_Err_Msg := '������ˮ��Ϊ:' || c_���׼�¼.������ˮ�� || '�Ľ������ڽ����У��������ٴ��ύ�˽���!';
        Raise Err_Special;
      End If;
    End If;
  End Loop;

  Select Max(��Ժ����id) Into n_����id From ������ҳ Where ����id = n_����id And ��ҳid = n_��ҳid;
  Select ���˽��ʼ�¼_Id.Nextval, Sysdate, Nextno(15) Into n_����id, d_����ʱ��, v_No From Dual;

  n_ʱ������ := Zl_Getsysparameter('���ʷ���ʱ��', 1137);

  Select ���˽��ʼ�¼_Id.Nextval, Sysdate, Nextno(15) Into n_����id, d_����ʱ��, v_No From Dual;

  If n_�������� = 2 Then
    Open c_���ý�����ϸ For
      Select Max(Decode(����id, Null, ID, Null)) As ID, Mod(��¼����, 10) As ��¼����, NO, ���, ��¼״̬, ִ��״̬,
             Trunc(Min(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ��Сʱ��, Trunc(Max(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ���ʱ��,
             Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) As ���, Sum(Nvl(���ʽ��, 0)) As ���ʽ��
      From סԺ���ü�¼
      Where ����id = n_����id And ��¼״̬ <> 0 And ��ҳid = n_��ҳid And ���ʷ��� = 1
      Group By Mod(��¼����, 10), NO, ���, ��¼״̬, ִ��״̬
      Having(Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) <> 0) Or (Sum(Nvl(ʵ�ս��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Sum(Nvl(���ʽ��, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(���ʽ��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Mod(Count(*), 2) = 0
      Order By NO, ���;
  Else
  
    Open c_���ý�����ϸ For
      Select Min(Decode(����id, Null, ID, Null)) As ID, Mod(��¼����, 10) As ��¼����, NO, ���, ��¼״̬, ִ��״̬,
             Trunc(Min(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ��Сʱ��, Trunc(Max(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ���ʱ��,
             Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) As ���, Sum(Nvl(���ʽ��, 0)) As ���ʽ��
      From ������ü�¼
      Where ����id = n_����id And ��¼״̬ <> 0 And ���ʷ��� = 1
      Group By Mod(��¼����, 10), NO, ���, ��¼״̬, ִ��״̬
      Having(Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) <> 0) Or (Sum(Nvl(ʵ�ս��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Sum(Nvl(���ʽ��, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(���ʽ��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Mod(Count(*), 2) = 0
      Union All
      Select Min(Decode(����id, Null, ID, Null)) As ID, Mod(��¼����, 10) As ��¼����, NO, ���, ��¼״̬, ִ��״̬,
             Trunc(Min(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ��Сʱ��, Trunc(Max(Decode(n_ʱ������, 0, �Ǽ�ʱ��, ����ʱ��))) As ���ʱ��,
             Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) As ���, Sum(Nvl(���ʽ��, 0)) As ���ʽ��
      From סԺ���ü�¼
      Where ����id = n_����id And ��¼״̬ <> 0 And ���ʷ��� = 1 And Mod(��¼����, 10) = 5
      Group By Mod(��¼����, 10), NO, ���, ��¼״̬, ִ��״̬
      Having(Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) <> 0) Or (Sum(Nvl(ʵ�ս��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Sum(Nvl(���ʽ��, 0)) = 0 And Mod(Count(*), 2) = 0) Or Sum(Nvl(���ʽ��, 0)) = 0 And Sum(Nvl(Ӧ�ս��, 0)) <> 0 And Mod(Count(*), 2) = 0
      Order By NO, ���;
  End If;

  n_�����ʽ�� := 0;
  Loop
    Fetch c_���ý�����ϸ
      Into n_����id, n_��¼����, v_����no, n_���, n_��¼״̬, n_ִ��״̬, d_��С����, d_�������, n_δ����, n_���ʽ��;
    Exit When c_���ý�����ϸ%NotFound;
  
    n_�����ʽ�� := n_�����ʽ�� + Nvl(n_δ����, 0);
    If d_��ʼ���� Is Null Then
      d_��ʼ���� := d_��С����;
    Elsif d_��ʼ���� > d_��С���� Then
      d_��ʼ���� := d_��С����;
    End If;
    If d_�������� Is Null Then
      d_�������� := d_�������;
    Elsif d_�������� < d_������� Then
      d_�������� := d_�������;
    End If;
  
    If Nvl(n_���ʽ��, 0) = 0 Then
      If n_����id Is Not Null Then
        If Length(v_Ids || ',' || n_����id) > 4000 Then
          v_Ids := Substr(v_Ids, 2);
          Zl_���ʷ��ü�¼_Batch(v_Ids, n_����id, n_����id);
          v_Ids := '';
        End If;
        v_Ids := v_Ids || ',' || n_����id;
      Else
        Zl_���ʷ��ü�¼_Insert(0, v_����no, n_��¼����, n_��¼״̬, n_ִ��״̬, n_���, n_δ����, n_����id);
      End If;
    Else
      Zl_���ʷ��ü�¼_Insert(0, v_����no, n_��¼����, n_��¼״̬, n_ִ��״̬, n_���, n_δ����, n_����id);
    End If;
  
  End Loop;

  If v_Ids Is Not Null Then
    v_Ids := Substr(v_Ids, 2);
    Zl_���ʷ��ü�¼_Batch(v_Ids, n_����id, n_����id);
  End If;
  n_�����ʽ�� := Round(n_�����ʽ��, 6);

  If n_�����ʽ�� <> Nvl(n_�����ܶ�, 0) Then
    v_Err_Msg := '����Ľ��ʽ����ʵ�ʽ��ʽ���,���������!';
    Raise Err_Item;
  End If;

  Zl_���˽��ʼ�¼_Insert(n_����id, v_No, n_����id, d_����ʱ��, d_��ʼ����, d_��������, 0, 0, n_��ҳid, Null, 2, Null, n_��������);

  n_���ʽ�� := 0;
  n_Count    := 0;
  For r_���㷽ʽ In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As ���㿨���,
                        Extractvalue(b.Column_Value, '/JS/JSKH') As ���㿨��,
                        Extractvalue(b.Column_Value, '/JS/JSFS') As ���㷽ʽ,
                        Extractvalue(b.Column_Value, '/JS/JSJE') As ������,
                        Extractvalue(b.Column_Value, '/JS/JYLSH') As ������ˮ��,
                        Extractvalue(b.Column_Value, '/JS/JYSM') As ����˵��, Extractvalue(b.Column_Value, '/JS/ZY') As ժҪ,
                        Extractvalue(b.Column_Value, '/JS/SFCYJ') As �Ƿ��Ԥ��,
                        Extractvalue(b.Column_Value, '/JS/SFXFK') As �Ƿ����ѿ�,
                        Extract(b.Column_Value, '/JS/EXPENDLIST') As Expend
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
  
    v_�����   := r_���㷽ʽ.���㷽ʽ;
    n_���ʽ�� := n_���ʽ�� + Nvl(r_���㷽ʽ.������, 0);
  
    If Nvl(r_���㷽ʽ.�Ƿ��Ԥ��, 0) = 0 Then
      --����
      If n_Count = 1 Then
        v_Err_Msg := '���ʽ����ݲ�֧�ֶ��ֽ��㷽ʽ!';
        Raise Err_Item;
      End If;
      n_�����id := Null;
      If r_���㷽ʽ.���㿨��� Is Not Null Then
        Select Decode(Translate(Nvl(r_���㷽ʽ.���㿨���, 'abcd'), '#1234567890', '#'), Null, 1, 0)
        Into n_Number
        From Dual;
      
        If Nvl(r_���㷽ʽ.�Ƿ����ѿ�, 0) = 1 Then
          If Nvl(n_Number, 0) = 1 Then
            Select Max(���), Max(���㷽ʽ), Max(����)
            Into n_���㿨���, v_���㷽ʽ, v_�����
            From �����ѽӿ�Ŀ¼
            Where ��� = n_�����id And Nvl(����, 0) = 1;
          Else
            Select Max(���), Max(���㷽ʽ), Max(����)
            Into n_���㿨���, v_���㷽ʽ, v_�����
            From �����ѽӿ�Ŀ¼
            Where ���� = r_���㷽ʽ.���㿨��� And Nvl(����, 0) = 1;
          
          End If;
        
          If n_���㿨��� Is Null Then
            v_Err_Msg := 'δ�ҵ���Ӧ�����ѿ���Ϣ';
            Raise Err_Item;
          
          End If;
          n_�����id := Null;
        
        Else
          If Nvl(n_Number, 0) = 1 Then
            Select Max(ID), Max(���㷽ʽ), Max(����)
            Into n_�����id, v_���㷽ʽ, v_�����
            From ҽ�ƿ����
            Where ID = n_�����id And Nvl(�Ƿ�����, 0) = 1;
          Else
            Select Max(ID), Max(���㷽ʽ), Max(����)
            Into n_�����id, v_���㷽ʽ, v_�����
            From ҽ�ƿ����
            Where ���� = r_���㷽ʽ.���㿨��� And Nvl(�Ƿ�����, 0) = 1;
          End If;
        
          If n_�����id Is Null Then
            v_Err_Msg := 'δ�ҵ���Ӧ��ҽ�ƿ���Ϣ!';
            Raise Err_Item;
          End If;
        End If;
      End If;
    
      If n_�����id Is Not Null Then
        --������,����סԺԤ����
        v_���㿨�� := r_���㷽ʽ.���㿨��;
        If r_���㷽ʽ.������ > 0 Then
          --��ֵ���ֲ�Ӧ���������ν���
          n_���ʽ�� := n_���ʽ�� - Nvl(r_���㷽ʽ.������, 0);
          Select ����Ԥ����¼_Id.Nextval, Nextno(11) Into n_Ԥ��id, v_Ԥ��no From Dual;
          Zl_����Ԥ����¼_Insert(n_Ԥ��id, v_Ԥ��no, Null, n_����id, n_��ҳid, n_����id, r_���㷽ʽ.������, v_���㷽ʽ, '', '', '', '', '',
                           v_����Ա����, v_����Ա����, Null, n_��������, n_�����id, Null, r_���㷽ʽ.���㿨��, r_���㷽ʽ.������ˮ��, r_���㷽ʽ.����˵��, Null,
                           d_����ʱ��, 0);
          n_Ԥ����ֵ := Nvl(n_Ԥ����ֵ, 0) + r_���㷽ʽ.������;
          For c_��չ��Ϣ In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As ��������,
                                Extractvalue(b.Column_Value, '/EXPEND/JYLR') As ��������
                         From Table(Xmlsequence(Extract(r_���㷽ʽ.Expend, '/EXPENDLIST/EXPEND'))) B) Loop
            Zl_�������㽻��_Insert(n_�����id, 0, r_���㷽ʽ.���㿨��, n_Ԥ��id, c_��չ��Ϣ.�������� || '|' || c_��չ��Ϣ.��������, 1);
          
          End Loop;
        
        Else
        
          Zl_���ʽɿ��¼_Insert(v_No, n_����id, n_��ҳid, n_����id, v_���㷽ʽ, Null, r_���㷽ʽ.������, n_����id, v_����Ա����, v_����Ա����, d_����ʱ��,
                           Null, Null, Null, Null, Null, n_�����id, r_���㷽ʽ.���㿨��, r_���㷽ʽ.������ˮ��, r_���㷽ʽ.����˵��);
        
          For c_��չ��Ϣ In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As ��������,
                                Extractvalue(b.Column_Value, '/EXPEND/JYLR') As ��������
                         From Table(Xmlsequence(Extract(r_���㷽ʽ.Expend, '/EXPENDLIST/EXPEND'))) B) Loop
            Zl_�������㽻��_Insert(n_�����id, 0, r_���㷽ʽ.���㿨��, n_����id, c_��չ��Ϣ.�������� || '|' || c_��չ��Ϣ.��������, 0);
          End Loop;
        
        End If;
      
      Else
        If n_���㿨��� Is Not Null Then
          --���ѿ�
          v_���ѿ����� := Nvl(v_���ѿ�����, '') || '||' || n_���㿨��� || '|' || r_���㷽ʽ.���㿨�� || '|0|' || r_���㷽ʽ.������;
        Else
          --��������
          Zl_���ʽɿ��¼_Insert(v_No, n_����id, n_��ҳid, n_����id, v_���㷽ʽ, Null, r_���㷽ʽ.������, n_����id, v_����Ա����, v_����Ա����, d_����ʱ��,
                           Null, Null, Null, Null, Null, n_�����id, r_���㷽ʽ.���㿨��, r_���㷽ʽ.������ˮ��, r_���㷽ʽ.����˵��);
        End If;
      End If;
      n_Count := 1;
    Else
      --��Ԥ��,ĿǰĬ��ȫ��
      n_��Ԥ����� := r_���㷽ʽ.������ + Nvl(n_Ԥ����ֵ, 0);
      For r_Ԥ�� In (Select Min(ID) As ID, NO, ���㷽ʽ, Sum(Nvl(���, 0)) - Sum(Nvl(��Ԥ��, 0)) As ���, ������ˮ��
                   From ����Ԥ����¼
                   Where ����id = n_����id And Mod(��¼����, 10) = 1 And Nvl(Ԥ�����, 2) = 2 And (��ҳid = n_��ҳid Or ��ҳid Is Null)
                   Group By NO, ���㷽ʽ, ������ˮ��
                   Having Sum(Nvl(���, 0)) - Sum(Nvl(��Ԥ��, 0)) <> 0) Loop
        Zl_����Ԥ����¼_Insert(r_Ԥ��.Id, r_Ԥ��.No, 1, r_Ԥ��.���, n_����id, n_����id);
        n_��Ԥ����� := n_��Ԥ����� - Nvl(r_Ԥ��.���, 0);
      End Loop;
      If n_��Ԥ����� <> 0 Then
        v_Err_Msg := '�����Ԥ�����������ʵ�ʲ���,����!';
        Raise Err_Item;
      End If;
    End If;
  
    Update �������׼�¼
    Set ҵ�����id = n_����id
    Where ��ˮ�� = Nvl(r_���㷽ʽ.������ˮ��, '-') And ��� = v_����� And ҵ������ = 2;
  End Loop;

  --���ѿ�����
  If v_���ѿ����� Is Not Null Then
    v_���ѿ����� := Substr(v_���ѿ�����, 3);
  End If;

  n_����   := Round(Nvl(n_�����ܶ�, 0) - Nvl(n_���ʽ��, 0), 6);
  v_���㷽ʽ := Null;
  If Abs(Nvl(n_����, 0)) > 1 Then
    v_Err_Msg := '���������������1.00��С��-1.00Ԫ,��������ʲ���,����!';
    Raise Err_Item;
  End If;

  n_�����ܶ� := n_���ʽ��;

  n_���ʽ�� := 0;
  If Nvl(n_����, 0) <> 0 Then
    Select Nvl(Max(����), '����') Into v_���㷽ʽ From ���㷽ʽ Where Nvl(����, 0) = 9;
    n_���ʽ�� := Nvl(n_����, 0);
  End If;
  If Nvl(n_����, 0) <> 0 Or v_���ѿ����� Is Not Null Then
    Zl_���ʽɿ��¼_Insert(v_No, n_����id, n_��ҳid, n_����id, v_���㷽ʽ, Null, n_���ʽ��, n_����id, v_����Ա����, v_����Ա����, d_����ʱ��, Null, Null,
                     Null, Null, Null, Null, Null, Null, Null, v_���ѿ�����);
  End If;

  --��������Ϣ�ܶ�������ܶ��Ƿ���ȷ
  Select Sum(��Ԥ��) Into n_���ʽ�� From ����Ԥ����¼ Where ����id = n_����id;
  If Round(n_���ʽ��, 6) <> Round(n_�����ܶ�, 6) Then
  
    v_Err_Msg := '����Ľ���ϼƽ����ʵ�ʽ��ʽ��ϼƲ���,���������!';
    Raise Err_Item;
  End If;

  Update ����Ԥ����¼ Set У�Ա�־ = 0 Where ����id = n_����id And Nvl(У�Ա�־, 0) <> 0;
  v_Temp := '<CZSJ>' || To_Char(d_����ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Err_Special Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20105, v_Temp);
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Third_Settlement;
/

--13874:���ϴ�,2018-06-04,�ҺŲ����˷��˵�ָ���Ľ��㷽ʽ
Create Or Replace Procedure Zl_���˹Һż�¼_Delete
(
  ���ݺ�_In       ������ü�¼.No%Type,
  ����Ա���_In   ������ü�¼.����Ա���%Type,
  ����Ա����_In   ������ü�¼.����Ա����%Type,
  ժҪ_In         ������ü�¼.ժҪ%Type := Null, --ԤԼȡ��ʱ ��д ���ԤԼȡ��ԭ��
  ɾ�������_In   Number := 0,
  ��ԭ���˽���_In Varchar2 := Null,
  �˷�����_In     In Number := 0, --0-ȫ�� 1-�˹Һŷ� 2-�˲����� 3-�˸��ӷ� 4-�˹Һ��벡�� 5-�˹Һ��븽��
  ��ָ������_In   ����Ԥ����¼.���㷽ʽ%Type := Null,
  �˺�����_In     Number := 1,
  �ջ�Ʊ�ݺ�_In   Varchar2 := Null,
  ����˵��_In     ����Ԥ����¼.����˵��%Type := Null,
  ���㷽ʽ_In     Varchar2 := Null,
  ��Ԥ��_In       ����Ԥ����¼.��Ԥ��%Type := Null
) As
  --�˷�����_In,��һ�¼�������²�׼���в����˷�
  --    2.�����ӿ�,��ʱ��֧��
  -- �ҺŷѲ����ѷֿ���,����
  --    ��ͨ���㷽ʽ:ԭ���㷽ʽ�˲��ַ���
  --    Ԥ����:Ԥ����,�˲���
  --    Ԥ��������ͨ���㷽ʽ���:�˿����ͨ���㷽ʽ������
  --    ���ѿ�:ԭ�������ò����������ѿ�
  --��ԭ���˽���_In:ָ�����˻���ԭ�����㷽ʽ(��ҽ���ĸ����˻�,�����˻������ֵ�),����ö�����
  --��ָ������_IN:ָ��ԭ���˽��㲿��,Ӧ���˸����ֽ��㷽ʽ,Ϊ��ʱȱʡ�˸��ֽ�,�����˸�ָ���Ľ��㷽ʽ

  --���α������ж��Ƿ񵥶��ղ�����,���ҺŻ��ܱ���
  Cursor c_Registinfo(v_״̬ ������ü�¼.��¼״̬%Type) Is
    Select a.����ʱ��, a.�Ǽ�ʱ��, c.����ʱ��, a.�շ�ϸĿid As ��Ŀid, c.ִ�в���id As ����id, c.ִ���� As ҽ������, d.Id As ҽ��id, c.�ű� As ����
    From ������ü�¼ A, �ҺŰ��� B, ���˹Һż�¼ C, ��Ա�� D
    Where a.��¼���� = 4 And a.��¼״̬ = v_״̬ And c.No = a.No And c.ִ���� = d.����(+) And a.No = ���ݺ�_In And
          Nvl(a.���㵥λ, '�ű�') = c.�ű� And Rownum < 2;
  r_Registrow c_Registinfo%RowType;

  --���α������жϼ�¼�Ƿ����,�����û��ܱ���
  Cursor c_Moneyinfo(v_״̬ ������ü�¼.��¼״̬%Type) Is
    Select ���˿���id, ��������id, ִ�в���id, ������Ŀid, Nvl(Sum(Ӧ�ս��), 0) As Ӧ��, Nvl(Sum(ʵ�ս��), 0) As ʵ��, Nvl(Sum(���ʽ��), 0) As ����
    From ������ü�¼
    Where ��¼���� = 4 And ��¼״̬ = v_״̬ And NO = ���ݺ�_In
    Group By ���˿���id, ��������id, ִ�в���id, ������Ŀid;
  r_Moneyrow c_Moneyinfo%RowType;

  --�ù�����ڴ�����Ա�ɿ�������˵Ĳ�ͬ���㷽ʽ�Ľ��
  Cursor c_Opermoney(n_Id ����Ԥ����¼.����id%Type) Is
    Select Distinct b.���㷽ʽ, -1 * Nvl(b.��Ԥ��, 0) As ��Ԥ��
    From ����Ԥ����¼ B
    Where b.����id = n_Id And b.��¼���� = 4 And b.��¼״̬ = 2 And Nvl(b.��Ԥ��, 0) <> 0;

  v_Err_Msg Varchar(255);
  Err_Item Exception;

  n_����id ����Ԥ����¼.����id%Type;
  n_����id ������ü�¼.����id%Type;

  v_��ָ�����㷽ʽ ����Ԥ����¼.���㷽ʽ%Type;
  n_�˿���       ����Ԥ����¼.��Ԥ��%Type;
  n_��ӡid         Ʊ�ݴ�ӡ����.Id%Type;
  n_����id         ������Ϣ.����id%Type;
  n_�˷ѽ��       ����Ԥ����¼.��Ԥ��%Type;
  n_Ԥ�����       ����Ԥ����¼.��Ԥ��%Type; --ԭ��¼ Ԥ���ɿ���
  n_����ֵ         �������.Ԥ�����%Type;
  n_�Һ�id         ���˹Һż�¼.Id%Type;
  n_��id           ����ɿ����.Id%Type;

  n_�����˷�       Number; --��¼�Ƿ��Ǵ˵��ݵĵڶ����˷�
  n_����̨ǩ���Ŷ� Number;
  n_ԤԼ���ɶ���   Number;
  n_ԤԼ�Һ�       Number;
  n_�Һ����ɶ���   Number;
  d_Date           Date;
  n_����           ������ü�¼.���ʷ���%Type;
  n_����id1        ������Ϣ.����id%Type;
  n_���ض�         ������ü�¼.ʵ�ս��%Type;
  n_�ѽ���         Number;
  n_����id         �ҺŰ���.Id%Type;
  n_�ƻ�id         �ҺŰ��żƻ�.Id%Type;
  d_����ʱ��       Date;
  d_����ʱ��       ���˹Һż�¼.����ʱ��%Type;
  n_�����¼id     �ٴ������¼.Id%Type;
  v_����           �ҺŰ���.����%Type;
  n_���           ���˹Һż�¼.����%Type;
  v_ʱ���         ʱ���.ʱ���%Type;
  d_��鿪ʼʱ��   Date;
  d_������ʱ��   Date;
  v_Temp           Varchar2(500);
  v_����ids        Varchar2(500);
  n_���ﲡ��id     ������Ϣ.����id%Type;
  d_����ʱ��       ����ǼǼ�¼.����ʱ��%Type;
  v_��������       Varchar2(5000);
  v_��ǰ����       Varchar2(1000);
  v_���㷽ʽ       ����Ԥ����¼.���㷽ʽ%Type;
  n_��������־     Number;
  n_������       ����Ԥ����¼.��Ԥ��%Type;
Begin
  n_��id           := Zl_Get��id(����Ա����_In);
  v_��ָ�����㷽ʽ := ��ָ������_In;

  --�����ж�Ҫ�˺�/ȡ��ԤԼ�ļ�¼�Ƿ����
  Open c_Moneyinfo(1);
  Fetch c_Moneyinfo
    Into r_Moneyrow;
  If c_Moneyinfo%RowCount = 0 Then
    Close c_Moneyinfo;
    Open c_Moneyinfo(0);
    Fetch c_Moneyinfo
      Into r_Moneyrow;
    If c_Moneyinfo%RowCount = 0 Then
      v_Err_Msg := 'Ҫ����ĵ��ݲ����ڡ�';
      Raise Err_Item;
    End If;
    n_ԤԼ�Һ� := 1;
  End If;
  Close c_Moneyinfo;

  Begin
    Select Zl_Fun_Regcustomname Into v_Temp From Dual;
    If v_Temp Is Not Null Then
      v_����ids := Substr(v_Temp, Instr(v_Temp, '|') + 1);
    End If;
  Exception
    When Others Then
      v_����ids := Null;
  End;

  v_Temp := zl_GetSysParameter(256);
  If v_Temp Is Null Or Substr(v_Temp, 1, 1) = '0' Then
    Null;
  Else
    Begin
      d_����ʱ�� := To_Date(Substr(v_Temp, 3), 'YYYY-MM-DD hh24:mi:ss');
    Exception
      When Others Then
        d_����ʱ�� := Null;
    End;
  End If;

  Select �ű�, ����, ����ʱ�� Into v_����, n_���, d_����ʱ�� From ���˹Һż�¼ Where NO = ���ݺ�_In And Rownum < 2;

  Begin
    Select a.Id Into n_����id From �ҺŰ��� A Where a.���� = v_����;
  Exception
    When Others Then
      n_����id := -1;
  End;

  Begin
    Select ID
    Into n_�ƻ�id
    From �ҺŰ��żƻ�
    Where ����id = n_����id And ���ʱ�� Is Not Null And
          Nvl(��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) =
          (Select Max(a.��Чʱ��) As ��Ч
           From �ҺŰ��żƻ� A
           Where a.���ʱ�� Is Not Null And d_����ʱ�� Between Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                 Nvl(a.ʧЧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) And a.����id = n_����id) And
          d_����ʱ�� Between Nvl(��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And
          Nvl(ʧЧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd'));
  Exception
    When Others Then
      n_�ƻ�id := 0;
  End;

  Begin
    If Nvl(n_�ƻ�id, 0) = 0 Then
      Select Decode(To_Char(d_����ʱ��, 'D'), '1', ����, '2', ��һ, '3', �ܶ�, '4', ����, '5', ����, '6', ����, '7', ����, Null)
      Into v_ʱ���
      From �ҺŰ���
      Where ID = n_����id;
    Else
      Select Decode(To_Char(d_����ʱ��, 'D'), '1', ����, '2', ��һ, '3', �ܶ�, '4', ����, '5', ����, '6', ����, '7', ����, Null)
      Into v_ʱ���
      From �ҺŰ��żƻ�
      Where ID = n_�ƻ�id;
    End If;
  Exception
    When Others Then
      v_ʱ��� := Null;
  End;

  If v_ʱ��� Is Not Null And d_����ʱ�� Is Not Null Then
    --����Ƿ��ģʽ�ҺŰ���
    Select To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
           To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || ' ' || To_Char(��ֹʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
    Into d_��鿪ʼʱ��, d_������ʱ��
    From ʱ���
    Where ʱ��� = v_ʱ��� And վ�� Is Null And ���� Is Null;
    If d_��鿪ʼʱ�� > d_������ʱ�� Then
      d_������ʱ�� := d_������ʱ�� + 1;
    End If;
    If d_��鿪ʼʱ�� < d_����ʱ�� And d_������ʱ�� > d_����ʱ�� Then
      --��ȡ�����¼id
      Begin
        Select a.Id
        Into n_�����¼id
        From �ٴ������¼ A, �ٴ������Դ B
        Where a.��Դid = b.Id And b.���� = v_���� And �ϰ�ʱ�� = v_ʱ��� And d_����ʱ�� Between ��ʼʱ�� And ��ֹʱ��;
      Exception
        When Others Then
          n_�����¼id := Null;
      End;
    End If;
  End If;

  --1.ԤԼ����
  If Nvl(n_ԤԼ�Һ�, 0) = 1 Then
    --������Լ��
    Open c_Registinfo(0);
    Fetch c_Registinfo
      Into r_Registrow;
  
    Update ���˹ҺŻ���
    Set ��Լ�� = Nvl(��Լ��, 0) - 1
    Where ���� = Trunc(r_Registrow.����ʱ��) And ����id = r_Registrow.����id And ��Ŀid = r_Registrow.��Ŀid And
          (���� = r_Registrow.���� Or ���� Is Null);
  
    If Sql%RowCount = 0 Then
      Insert Into ���˹ҺŻ���
        (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, ��Լ��)
      Values
        (Trunc(r_Registrow.����ʱ��), r_Registrow.����id, r_Registrow.��Ŀid, r_Registrow.ҽ������,
         Decode(r_Registrow.ҽ��id, 0, Null, r_Registrow.ҽ��id), r_Registrow.����, -1);
    End If;
    Close c_Registinfo;
  
    --���¹Һ����״̬
    Delete �Һ����״̬
    Where ״̬ = 2 And
          (����, ���, ����) = (Select ���㵥λ, ��ҩ����, Trunc(����ʱ��)
                          From ������ü�¼
                          Where ��¼���� = 4 And ��¼״̬ = 0 And ��� = 1 And Rownum = 1 And NO = ���ݺ�_In) Or
          (����, ���, ����) = (Select ���㵥λ, ��ҩ����, ����ʱ��
                          From ������ü�¼
                          Where ��¼���� = 4 And ��¼״̬ = 0 And ��� = 1 And Rownum = 1 And NO = ���ݺ�_In);
  
    --��Ӳ��˹Һż�¼�� ������¼
    Select ���˹Һż�¼_Id.Nextval, Sysdate Into n_�Һ�id, d_Date From Dual;
  
    Update ���˹Һż�¼ Set ��¼״̬ = 3 Where NO = ���ݺ�_In And ��¼״̬ = 1 And ��¼���� = 2;
    If Sql%NotFound Then
      v_Err_Msg := 'ԤԼ����' || ���ݺ�_In || '�������ڻ����ڲ���ԭ���Ѿ���ȡ��ԤԼ';
      Raise Err_Item;
    End If;
  
    Insert Into ���˹Һż�¼
      (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����,
       ����, ����, ����, ԤԼ, ժҪ, ԤԼ��ʽ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ҽ�Ƹ��ʽ)
      Select n_�Һ�id, NO, ��¼����, 2, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, d_Date, ����ʱ��,
             ����Ա���_In, ����Ա����_In, ����, ����, ����, ԤԼ, Nvl(ժҪ_In, ժҪ) As ժҪ, ԤԼ��ʽ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ҽ�Ƹ��ʽ
      From ���˹Һż�¼
      Where NO = ���ݺ�_In;
  
    If n_�����¼id Is Not Null Then
      Update �ٴ������¼ Set ��Լ�� = ��Լ�� - 1 Where ID = n_�����¼id And Nvl(��Լ��, 0) > 0;
      Update �ٴ�������ſ��� Set �Һ�״̬ = Null, ����Ա���� = Null Where ��¼id = n_�����¼id And ��� = n_���;
    End If;
  
    --Update ���˹Һż�¼ set ժҪ=nvl(ժҪ_IN,ժҪ) where NO=���ݺ�_IN;
    --ɾ��������ü�¼
    Delete From ������ü�¼ Where NO = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 0;
    --���ԤԼ���ɶ���ʱ��Ҫ�������
  
    n_ԤԼ���ɶ��� := Zl_To_Number(zl_GetSysParameter('ԤԼ���ɶ���', 1113));
    If Nvl(n_ԤԼ���ɶ���, 0) = 1 Then
      --Ҫɾ������
      For v_�Һ� In (Select ID, ����, ִ�в���id, ִ���� From ���˹Һż�¼ Where NO = ���ݺ�_In) Loop
        Zl_�ŶӽкŶ���_Delete(v_�Һ�.ִ�в���id, v_�Һ�.Id);
      End Loop;
    End If;
    Return;
  End If;

  Select Nvl(���ʷ���, 0), ����id, Decode(Sign(Nvl(����id, 0)), 0, 0, 1)
  Into n_����, n_����id, n_�ѽ���
  From ������ü�¼
  Where ��¼���� = 4 And NO = ���ݺ�_In And ��¼״̬ In (1, 3) And Rownum < 2;

  --2.�ҺŴ���
  n_�ѽ��� := Nvl(n_�ѽ���, 0);

  If n_�ѽ��� = 1 And n_���� = 1 Then
    Select Sysdate, Null Into d_Date, n_����id From Dual;
  Else
    Select Sysdate, ���˽��ʼ�¼_Id.Nextval Into d_Date, n_����id From Dual;
  End If;

  ----0-ȫ�� 1-�˹Һŷ� 2-�˲�����
  If Nvl(�˷�����_In, 0) Not In (2, 3) Then
    --���ǹ��˲�����ʱ����
    --���¹Һ����״̬
    If �˺�����_In = 1 Then
      Delete �Һ����״̬
      Where ״̬ = 1 And
            (����, ���, ����) = (Select �ű�, ����, Trunc(����ʱ��) From ���˹Һż�¼ Where NO = ���ݺ�_In And Rownum = 1) Or
            (����, ���, ����) = (Select �ű�, ����, ����ʱ�� From ���˹Һż�¼ Where NO = ���ݺ�_In And Rownum = 1);
    Else
      Update �Һ����״̬
      Set ״̬ = 4
      Where ״̬ = 1 And
            (����, ���, ����) = (Select �ű�, ����, Trunc(����ʱ��) From ���˹Һż�¼ Where NO = ���ݺ�_In And Rownum = 1) Or
            (����, ���, ����) = (Select �ű�, ����, ����ʱ�� From ���˹Һż�¼ Where NO = ���ݺ�_In And Rownum = 1);
    End If;
  
    --���˾���״̬
    If n_����id Is Not Null Then
      Update ������Ϣ Set ����״̬ = 0, �������� = Null Where ����id = n_����id;
    
      --ɾ���������ش���,ֻ�е�ֻ��һ���Һż�¼���Ҳ��˽���������Һ����ڽ���ʱ�Żᴦ��
      If ɾ�������_In = 1 Then
        Delete ���ﲡ����¼ Where ����id = n_����id;
        Update ������Ϣ Set ����� = Null Where ����id = n_����id;
        --���ü�¼�����Һż����������￨����,�Լ����˽��Ѻ��˷ѻ����ʵķ���,�Һż�¼�������
        Update ������ü�¼ Set ��ʶ�� = Null Where �����־ = 1 And ����id = n_����id;
      End If;
    End If;
  
    --�����ʱ���˾��￨��,�˷�ʱ������￨��,�ڷǹ��˲�����ʱ
    n_����id1 := Null;
    Begin
      Select ����id
      Into n_����id1
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And ���ӱ�־ = 2 And Rownum < 2;
    Exception
      When Others Then
        Null;
    End;
  
    If n_����id1 Is Not Null And Nvl(�˷�����_In, 0) Not In (2, 3) Then
      Update ������Ϣ
      Set ���￨�� = Null, ����֤�� = Null, Ic���� = Decode(Ic����, ���￨��, Null, Ic����)
      Where ����id = n_����id1;
    End If;
  
  End If;

  --���ǰ���Ƿ��Ѿ������˹�����
  Begin
    Select 1 Into n_�����˷� From ������ü�¼ Where ��¼���� = 4 And NO = ���ݺ�_In And ��¼״̬ = 3 And Rownum < 2;
  Exception
    When Others Then
      n_�����˷� := 0;
  End;

  If Nvl(�˷�����_In, 0) = 0 Or Nvl(�˷�����_In, 0) = 2 Then
    --ȫ��,�˲�����
    --������ü�¼��������¼
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
       ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��,
       ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����,
             �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ִ����,
             ����Ա���_In, ����Ա����_In, ����ʱ��, d_Date, n_����id,
             Decode(n_����, 1, Decode(Nvl(n_�ѽ���, 0), 0, -1 * ʵ�ս��, Null), -1 * ���ʽ��), ������Ŀ��, ���մ���id, -1 * ͳ����,
             Nvl(ժҪ_In, ժҪ) As ժҪ, Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
  
    --ԭʼ��¼
    If n_���� = 1 And Nvl(n_�ѽ���, 0) = 0 Then
      Update ������ü�¼
      Set ��¼״̬ = 3, ����id = n_����id, ���ʽ�� = ʵ�ս��
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    Else
      Update ������ü�¼
      Set ��¼״̬ = 3
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    End If;
  Elsif Nvl(�˷�����_In, 0) = 1 Or Nvl(�˷�����_In, 0) = 4 Then
    --�˹Һŷ�,�˹Һ��벡����
    --������ü�¼��������¼
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
       ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��,
       ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����,
             �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ִ����,
             ����Ա���_In, ����Ա����_In, ����ʱ��, d_Date, n_����id,
             Decode(n_����, 1, Decode(Nvl(n_�ѽ���, 0), 0, -1 * ʵ�ս��, Null), -1 * ���ʽ��), ������Ŀ��, ���մ���id, -1 * ͳ����,
             Nvl(ժҪ_In, ժҪ) As ժҪ, Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') = 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
  
    --ԭʼ��¼
    If n_���� = 1 And Nvl(n_�ѽ���, 0) = 0 Then
      Update ������ü�¼
      Set ��¼״̬ = 3, ����id = n_����id, ���ʽ�� = ʵ�ս��
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') = 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    Else
      Update ������ü�¼
      Set ��¼״̬ = 3
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') = 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    End If;
  Elsif Nvl(�˷�����_In, 0) = 3 Then
    --�˸��ӷ�
    --������ü�¼��������¼
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
       ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��,
       ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����,
             �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ִ����,
             ����Ա���_In, ����Ա����_In, ����ʱ��, d_Date, n_����id,
             Decode(n_����, 1, Decode(Nvl(n_�ѽ���, 0), 0, -1 * ʵ�ս��, Null), -1 * ���ʽ��), ������Ŀ��, ���մ���id, -1 * ͳ����,
             Nvl(ժҪ_In, ժҪ) As ժҪ, Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') > 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
  
    --ԭʼ��¼
    If n_���� = 1 And Nvl(n_�ѽ���, 0) = 0 Then
      Update ������ü�¼
      Set ��¼״̬ = 3, ����id = n_����id, ���ʽ�� = ʵ�ս��
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') > 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    Else
      Update ������ü�¼
      Set ��¼״̬ = 3
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') > 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    End If;
  Elsif Nvl(�˷�����_In, 0) = 5 Then
    --�˹Һ��븽�ӷ�
    --������ü�¼��������¼
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
       ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��,
       ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����,
             �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ִ����,
             ����Ա���_In, ����Ա����_In, ����ʱ��, d_Date, n_����id,
             Decode(n_����, 1, Decode(Nvl(n_�ѽ���, 0), 0, -1 * ʵ�ս��, Null), -1 * ���ʽ��), ������Ŀ��, ���մ���id, -1 * ͳ����,
             Nvl(ժҪ_In, ժҪ) As ժҪ, Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Nvl(���ӱ�־, 0) <> 1;
  
    --ԭʼ��¼
    If n_���� = 1 And Nvl(n_�ѽ���, 0) = 0 Then
      Update ������ü�¼
      Set ��¼״̬ = 3, ����id = n_����id, ���ʽ�� = ʵ�ս��
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Nvl(���ӱ�־, 0) <> 1;
    Else
      Update ������ü�¼
      Set ��¼״̬ = 3
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Nvl(���ӱ�־, 0) <> 1;
    End If;
  End If;

  n_����id := 0;
  If n_���� = 0 Then
    --��ȡ����ID
    Select Nvl(����id, 0)
    Into n_����id
    From ������ü�¼
    Where ��¼���� = 4 And ��¼״̬ = 3 And NO = ���ݺ�_In And Rownum < 2;
  End If;

  If n_���� = 1 Then
    --����
    For c_���� In (Select ʵ�ս��, ���˿���id, ��������id, ִ�в���id, ������Ŀid
                 From ������ü�¼
                 Where ��¼���� = 4 And ��¼״̬ = 3 And NO = ���ݺ�_In And
                       Nvl(���ӱ�־, 0) =
                       Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0)) And
                       Nvl(���ʷ���, 0) = 1) Loop
      --�������
      Update �������
      Set ������� = Nvl(�������, 0) - Nvl(c_����.ʵ�ս��, 0)
      Where ����id = Nvl(n_����id, 0) And ���� = 1 And ���� = 1
      Returning ������� Into n_���ض�;
    
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, ����, ����, �������, Ԥ�����)
        Values
          (n_����id, 1, 1, -1 * Nvl(c_����.ʵ�ս��, 0), 0);
        n_���ض� := Nvl(c_����.ʵ�ս��, 0);
      End If;
      If Nvl(n_���ض�, 0) = 0 Then
        Delete �������
        Where ����id = Nvl(n_����id, 0) And ���� = 1 And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
      End If;
      --����δ�����
      Update ����δ�����
      Set ��� = Nvl(���, 0) - Nvl(c_����.ʵ�ս��, 0)
      Where ����id = n_����id And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���˿���id, 0) = Nvl(c_����.���˿���id, 0) And
            Nvl(��������id, 0) = Nvl(c_����.��������id, 0) And Nvl(ִ�в���id, 0) = Nvl(c_����.ִ�в���id, 0) And ������Ŀid + 0 = c_����.������Ŀid And
            ��Դ;�� + 0 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into ����δ�����
          (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
        Values
          (n_����id, Null, Null, c_����.���˿���id, c_����.��������id, c_����.ִ�в���id, c_����.������Ŀid, 1, -1 * Nvl(c_����.ʵ�ս��, 0));
      End If;
    End Loop;
    Delete ����δ�����
    Where ����id = n_����id And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���, 0) = 0 And ��Դ;�� + 0 = 1;
  End If;

  If n_���� = 0 Then
    --1.�˷�
    --���˹ҺŽ���:�ֽ�͸����ʻ�����
    If ���㷽ʽ_In Is Null And Nvl(��Ԥ��_In, 0) = 0 Then
      If ��ԭ���˽���_In Is Not Null Then
        --�˿����ȡ
        If Nvl(�˷�����_In, 0) <> 0 Or Nvl(n_�����˷�, 0) <> 0 Then
          --����ǵ����˲�����,����ֻ�˹Һŷ�,�Ȼ�ȡ�˷ѽ��
          Begin
            --��ȡ�����˿���
            Select -1 * Sum(Nvl(ʵ�ս��, 0)) As �տ���
            Into n_�˿���
            From ������ü�¼
            Where NO = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 2 And ����id = n_����id;
          
          Exception
            When Others Then
              v_Err_Msg := '���ݡ�' || ���ݺ�_In || '����' || Case Nvl(�˷�����_In, 0)
                             When 1 Then
                              '�Һŷ���'
                             When 2 Then
                              '������'
                           End || '�������ڲ���ԭ���Ѿ��������˷ѻ��ߵ��ݲ�����!';
              Raise Err_Item;
          End;
          Begin
            Select ��Ԥ��
            Into n_�˷ѽ��
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') = 0;
          Exception
            When Others Then
              n_�˷ѽ�� := 0;
          End;
        
          --a.����Ľ��㷽ʽ
        
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
             ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In, -n_�˿���,
                   n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') = 0;
        
          If n_�˷ѽ�� = 0 Then
            --b.����������ֽ�
            If n_�˿��� <> 0 Then
              If v_��ָ�����㷽ʽ Is Null Then
                --�˸��ֽ�
                Begin
                  Select ���� Into v_��ָ�����㷽ʽ From ���㷽ʽ Where ���� = 1;
                Exception
                  When Others Then
                    v_��ָ�����㷽ʽ := '�ֽ�';
                End;
              End If;
              Update ����Ԥ����¼
              Set ��Ԥ�� = ��Ԥ�� + (-1 * n_�˿���), ����˵�� = Nvl(����˵��_In, ����˵��)
              Where ��¼���� = 4 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_��ָ�����㷽ʽ;
              If Sql%RowCount = 0 Then
                Insert Into ����Ԥ����¼
                  (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����,
                   �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
                  Select ����Ԥ����¼_Id.Nextval, a.No, a.ʵ��Ʊ��, a.��¼����, 2, a.����id, a.��ҳid, a.����id, a.ժҪ, v_��ָ�����㷽ʽ, d_Date,
                         ����Ա���_In, ����Ա����_In, -1 * n_�˿���, n_����id, n_��id, Ԥ�����, Decode(����˵��_In, Null, �����id, Null), ���㿨���,
                         Decode(����˵��_In, Null, ����, Null), Decode(����˵��_In, Null, ������ˮ��, Null), Nvl(����˵��_In, ����˵��), ������λ, 4
                  From ����Ԥ����¼ A
                  Where a.��¼���� = 4 And a.��¼״̬ = 1 And a.����id = n_����id And Rownum < 2;
              End If;
            End If;
          End If;
        Else
          --a.����Ľ��㷽ʽԭ����
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
             ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In, -��Ԥ��,
                   n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') = 0;
        
          --b.����������ֽ�
          Begin
            Select Sum(��Ԥ��)
            Into n_�˷ѽ��
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') > 0;
          Exception
            When Others Then
              n_�˷ѽ�� := 0;
          End;
          If n_�˷ѽ�� <> 0 Then
            If v_��ָ�����㷽ʽ Is Null Then
              --�˸��ֽ�
              Begin
                Select ���㷽ʽ
                Into v_��ָ�����㷽ʽ
                From ����Ԥ����¼
                Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') = 0;
              
              Exception
                When Others Then
                  Begin
                    Select ���� Into v_��ָ�����㷽ʽ From ���㷽ʽ Where ���� = 1;
                  Exception
                    When Others Then
                      v_��ָ�����㷽ʽ := '�ֽ�';
                  End;
              End;
            End If;
            Update ����Ԥ����¼
            Set ��Ԥ�� = ��Ԥ�� + (-1 * n_�˷ѽ��), ����˵�� = Nvl(����˵��_In, ����˵��)
            Where ��¼���� = 4 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_��ָ�����㷽ʽ;
            If Sql%RowCount = 0 Then
              Insert Into ����Ԥ����¼
                (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
                 ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
                Select ����Ԥ����¼_Id.Nextval, a.No, a.ʵ��Ʊ��, a.��¼����, 2, a.����id, a.��ҳid, a.����id, a.ժҪ, v_��ָ�����㷽ʽ, d_Date,
                       ����Ա���_In, ����Ա����_In, -1 * n_�˷ѽ��, n_����id, n_��id, Ԥ�����, Decode(����˵��_In, Null, �����id, Null), ���㿨���,
                       Decode(����˵��_In, Null, ����, Null), Decode(����˵��_In, Null, ������ˮ��, Null), Nvl(����˵��_In, ����˵��), ������λ, 4
                From ����Ԥ����¼ A
                Where a.��¼���� = 4 And a.��¼״̬ = 1 And a.����id = n_����id And Rownum < 2;
            End If;
          End If;
        End If;
      Else
        --�˿����ȡ
        If Nvl(�˷�����_In, 0) <> 0 Or Nvl(n_�����˷�, 0) <> 0 Then
          --����ǵ����˲�����,����ֻ�˹Һŷ�,�Ȼ�ȡ�˷ѽ��
          Begin
            --��ȡ�����˿���
            Select -1 * Sum(Nvl(ʵ�ս��, 0)) As �տ���
            Into n_�˿���
            From ������ü�¼
            Where NO = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 2 And ����id = n_����id;
          Exception
            When Others Then
              v_Err_Msg := '���ݡ�' || ���ݺ�_In || '����' || Case Nvl(�˷�����_In, 0)
                             When 1 Then
                              '�Һŷ���'
                             When 2 Then
                              '������'
                           End || '�������ڲ���ԭ���Ѿ��������˷ѻ��ߵ��ݲ�����!';
              Raise Err_Item;
          End;
        End If;
        If Nvl(n_�����˷�, 0) = 0 And Nvl(�˷�����_In, 0) = 0 Then
          --�״�ȫ��
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
             ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In, -1 * ��Ԥ��,
                   n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id;
        Else
          --�����˷�,���߱��ε���һ����
          --�����˷�ʱ,��¼״̬=3 ,�״β�����,��¼״̬Ϊ1
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
             ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                   -1 * n_�˿���, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id And ժҪ = 'ҽ���Һ�' And ��Ԥ�� = n_�˿��� And
                  Rownum < 2;
          If Sql%RowCount = 0 Then
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
               ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
              Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                     -1 * n_�˿���, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
              From ����Ԥ����¼
              Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id And ��Ԥ�� = n_�˿��� And Rownum < 2;
          End If;
          If Sql%RowCount = 0 Then
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
               ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
              Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                     -1 * n_�˿���, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
              From ����Ԥ����¼
              Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id And Rownum < 2;
            If Sql%RowCount = 0 Then
              --�����˷�,����ȫ��ʹ��Ԥ����ɷ�ʱ�Ŵ��ڴ������
              n_Ԥ����� := n_�˿���;
            End If;
          End If;
        
        End If;
      End If;
    Else
      --�����㷽ʽ��
      If ���㷽ʽ_In is Not Null then
         v_�������� := ���㷽ʽ_In || '|'; --�Կո�ֿ���|��β,û�н�������
        While v_�������� Is Not Null Loop
          v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '|') - 1);
          v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
        
          v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
          n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1));
        
          v_��ǰ����   := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
          n_��������־ := To_Number(v_��ǰ����);
        
          If n_��������־ = 0 Then
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
               ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
              Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, v_���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                     -1 * n_������, n_����id, n_��id, Ԥ�����, Null, Null, Null, Null, ����˵��_In, ������λ, 4
              From ����Ԥ����¼
              Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id And Rownum < 2;
          Else
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
               ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
              Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, v_���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                     -1 * n_������, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, Nvl(����˵��_In, ����˵��), ������λ, 4
              
              From ����Ԥ����¼
              Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id And
                    (�����id Is Not Null Or ���㿨��� Is Not Null) And Rownum < 2;
          End If;
          v_�������� := Substr(v_��������, Instr(v_��������, '|') + 1);
        End Loop;
      End IF;
      n_Ԥ����� := Nvl(��Ԥ��_In, 0);
    End IF;
    --�״��˷�ʱ,��¼״̬�����Ϊ��3
    Update ����Ԥ����¼
    Set ��¼״̬ = 3
    Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id;
  
    --��Ԥ�� 1-ȫ�� 2-������,������ʱ��ȫ��ʹ��Ԥ�����нɿ�
    If Nvl(�˷�����_In, 0) = 0 Or (Nvl(�˷�����_In, 0) <> 0 And n_Ԥ����� <> 0) Then
      --���˹ҺŽ���:��Ԥ�����
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
         ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
        Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, d_Date,
               ����Ա����_In, ����Ա���_In, -1 * Decode(Nvl(�˷�����_In, 0) + Nvl(n_�����˷�, 0), 0, ��Ԥ��, n_Ԥ�����), n_����id, n_��id, Ԥ�����,
               �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
        From ����Ԥ����¼
        Where ��¼���� In (1, 11) And ����id = n_����id And Nvl(��Ԥ��, 0) <> 0 And
              Rownum = Decode(Nvl(�˷�����_In, 0) + Nvl(n_�����˷�, 0), 0, Rownum, 1);
    End If;
  
    --������Ԥ�����
    For c_Ԥ�� In (Select ����id, Ԥ�����, -1 * Sum(Nvl(��Ԥ��, 0)) As ��Ԥ��
                 From ����Ԥ����¼
                 Where ��¼���� In (1, 11) And ����id = n_����id
                 Group By ����id, Ԥ�����) Loop
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(c_Ԥ��.��Ԥ��, 0)
      Where ����id = c_Ԥ��.����id And ���� = Nvl(c_Ԥ��.Ԥ�����, 2) And ���� = 1
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, Ԥ�����, ����, ����)
        Values
          (c_Ԥ��.����id, Nvl(c_Ԥ��.��Ԥ��, 0), 1, Nvl(c_Ԥ��.Ԥ�����, 2));
        n_����ֵ := Nvl(c_Ԥ��.��Ԥ��, 0);
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From �������
        Where ����id = c_Ԥ��.����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
      End If;
    End Loop;
  
    If �ջ�Ʊ�ݺ�_In Is Not Null Then
      --���˹Һŷ�,������Ʊ��
      --�˿��ջ�Ʊ��(�����ϴιҺ�ʹ��Ʊ��,�����ջ�)
      Begin
        --�����һ�δ�ӡ��������ȡ
        --81907
        Select ID
        Into n_��ӡid
        From (Select b.Id
               From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
               Where a.��ӡid = b.Id And a.���� = 1 And a.ԭ�� In (1, 3) And b.�������� = 4 And b.No = ���ݺ�_In
               Order By a.ʹ��ʱ�� Desc)
        Where Rownum < 2;
      Exception
        When Others Then
          n_��ӡid := Null;
      End;
    
      --���ջ�ԭƱ��
      If n_��ӡid Is Not Null Then
        Begin
          Insert Into Ʊ��ʹ����ϸ
            (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����)
            Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_Date, ����Ա����_In
            From Ʊ��ʹ����ϸ
            Where ��ӡid = n_��ӡid And ���� = 1 And Instr(',' || �ջ�Ʊ�ݺ�_In || ',', ',' || ���� || ',') > 0;
        Exception
          When Others Then
            Delete From Ʊ��ʹ����ϸ Where ��ӡid = n_��ӡid And ���� = 2 And ԭ�� = 2;
            Insert Into Ʊ��ʹ����ϸ
              (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����)
              Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_Date, ����Ա����_In
              From Ʊ��ʹ����ϸ
              Where ��ӡid = n_��ӡid And ���� = 1 And Instr(',' || �ջ�Ʊ�ݺ�_In || ',', ',' || ���� || ',') > 0;
        End;
      End If;
    End If;
  End If;

  --�����˲�������,��������ܼ�¼
  --��ػ��ܱ�Ĵ���

  --���˹ҺŻ���
  If Nvl(�˷�����_In, 0) Not In (2, 3) Then
    Open c_Registinfo(3);
    Fetch c_Registinfo
      Into r_Registrow;
  
    If c_Registinfo%RowCount = 0 Then
      --ֻ�ղ�����ʱ�޺ű�,������
      Close c_Registinfo;
    Else
    
      --��Ҫȷ���Ƿ�ԤԼ�Һ�
      --1.�������ԤԼ�ҺŲ����ĹҺż�¼,����Ҫ���ѹ����������ѽ���
      --2.����������Һ�,��ֻ���ѹ���
    
      Begin
        Select Decode(ԤԼ, Null, 0, 0, 0, 1) Into n_ԤԼ�Һ� From ���˹Һż�¼ Where NO = ���ݺ�_In And Rownum = 1;
      Exception
        When Others Then
          n_ԤԼ�Һ� := 0;
      End;
    
      Update ���˹ҺŻ���
      Set �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - n_ԤԼ�Һ�, ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ�Һ�
      Where ���� = Trunc(r_Registrow.����ʱ��) And ����id = r_Registrow.����id And ��Ŀid = r_Registrow.��Ŀid And
            (���� = r_Registrow.���� Or ���� Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into ���˹ҺŻ���
          (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���, �����ѽ���, ��Լ��)
        Values
          (Trunc(r_Registrow.����ʱ��), r_Registrow.����id, r_Registrow.��Ŀid, r_Registrow.ҽ������,
           Decode(r_Registrow.ҽ��id, 0, Null, r_Registrow.ҽ��id), r_Registrow.����, -1, -1 * n_ԤԼ�Һ�, -1 * n_ԤԼ�Һ�);
      End If;
    
      If n_�����¼id Is Not Null Then
        Update �ٴ������¼
        Set �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - n_ԤԼ�Һ�, ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ�Һ�
        Where ID = n_�����¼id And Nvl(��Լ��, 0) > 0;
        Update �ٴ�������ſ��� Set �Һ�״̬ = Null, ����Ա���� = Null Where ��¼id = n_�����¼id And ��� = n_���;
      End If;
    
      Close c_Registinfo;
    End If;
  End If;

  If n_���� = 0 Then
    --��Ա�ɿ����(���������ʻ��ȵĽ�����,�����˳�Ԥ����)
    For r_Opermoney In c_Opermoney(n_����id) Loop
      Update ��Ա�ɿ����
      Set ��� = Nvl(���, 0) + (-1 * r_Opermoney.��Ԥ��)
      Where �տ�Ա = ����Ա����_In And ���㷽ʽ = r_Opermoney.���㷽ʽ And ���� = 1
      Returning ��� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ��Ա�ɿ����
          (�տ�Ա, ���㷽ʽ, ����, ���)
        Values
          (����Ա����_In, r_Opermoney.���㷽ʽ, 1, -1 * r_Opermoney.��Ԥ��);
        n_����ֵ := r_Opermoney.��Ԥ��;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From ��Ա�ɿ����
        Where �տ�Ա = ����Ա����_In And ���㷽ʽ = r_Opermoney.���㷽ʽ And ���� = 1 And Nvl(���, 0) = 0;
      End If;
    End Loop;
  End If;
  If Nvl(�˷�����_In, 0) Not In (2, 3) Then
    n_�Һ����ɶ��� := Zl_To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113));
    If n_�Һ����ɶ��� <> 0 Then
      n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113));
      If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Then
      
        --Ҫɾ������
        For v_�Һ� In (Select ID, ����, ִ�в���id, ִ���� From ���˹Һż�¼ Where NO = ���ݺ�_In) Loop
          Zl_�ŶӽкŶ���_Delete(v_�Һ�.ִ�в���id, v_�Һ�.Id);
        End Loop;
      End If;
    End If;
  
    --ҽ�������ľ���ǼǼ�¼
    Begin
      Select ����id, ����ʱ�� Into n_���ﲡ��id, d_����ʱ�� From ���˹Һż�¼ Where NO = ���ݺ�_In;
      Delete From ����ǼǼ�¼ Where ����id = n_���ﲡ��id And ����ʱ�� = d_����ʱ�� And ��ҳid Is Null;
    Exception
      When Others Then
        Null;
    End;
    --���˹Һż�¼
    Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
  
    Update ���˹Һż�¼ Set ��¼״̬ = 3 Where NO = ���ݺ�_In And ��¼���� = 1 And ��¼״̬ = 1;
    If Sql%NotFound Then
      v_Err_Msg := '�Һŵ���' || ���ݺ�_In || '�������ڻ����ڲ���ԭ���Ѿ����˺�';
      Raise Err_Item;
    End If;
    Insert Into ���˹Һż�¼
      (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����,
       ����, ����, ����, ԤԼ, ժҪ, ԤԼ��ʽ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ����, ҽ�Ƹ��ʽ)
      Select n_�Һ�id, NO, ��¼����, 2, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, d_Date, ����ʱ��,
             ����Ա���_In, ����Ա����_In, ����, ����, ����, ԤԼ, Nvl(ժҪ_In, ժҪ) As ժҪ, ԤԼ��ʽ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ����, ҽ�Ƹ��ʽ
      From ���˹Һż�¼
      Where NO = ���ݺ�_In;
  End If;
  --��Ϣ����
  Begin
    Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
      Using 2, ���ݺ�_In;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˹Һż�¼_Delete;
/

--126862:����,2018-06-20,�˲����ѻ򸽼ӷ�ʱ,��Ӧ�ø��²��˹ҺŻ���
--13874:���ϴ�,2018-06-04,�ҺŲ����˷��˵�ָ���Ľ��㷽ʽ
Create Or Replace Procedure Zl_���˹Һż�¼_����_Delete
(
  ���ݺ�_In       ������ü�¼.No%Type,
  ����Ա���_In   ������ü�¼.����Ա���%Type,
  ����Ա����_In   ������ü�¼.����Ա����%Type,
  ժҪ_In         ������ü�¼.ժҪ%Type := Null, --ԤԼȡ��ʱ ��д ���ԤԼȡ��ԭ�� 
  ɾ�������_In   Number := 0,
  ��ԭ���˽���_In Varchar2 := Null,
  �˷�����_In     In Number := 0, --0-ȫ�� 1-�˹Һŷ� 2-�˲����� 3-�˸��ӷ� 4-�˹Һ��벡�� 5-�˹Һ��븽�� 
  ��ָ������_In   ����Ԥ����¼.���㷽ʽ%Type := Null,
  �˺�����_In     Number := 1,
  ���㷽ʽ_In     Varchar2 := Null,
  ��Ԥ��_In       ����Ԥ����¼.��Ԥ��%Type := Null,
  �ջ�Ʊ�ݺ�_In   Varchar2 := Null,
  ����˵��_In     ����Ԥ����¼.����˵��%Type := Null
) As
  --�˷�����_In,��һ�¼�������²�׼���в����˷� 
  --    2.�����ӿ�,��ʱ��֧�� 
  -- �ҺŷѲ����ѷֿ���,���� 
  --    ��ͨ���㷽ʽ:ԭ���㷽ʽ�˲��ַ��� 
  --    Ԥ����:Ԥ����,�˲��� 
  --    Ԥ��������ͨ���㷽ʽ���:�˿����ͨ���㷽ʽ������ 
  --    ���ѿ�:ԭ�������ò����������ѿ� 
  --��ԭ���˽���_In:ָ�����˻���ԭ�����㷽ʽ(��ҽ���ĸ����˻�,�����˻������ֵ�),����ö����� 
  --��ָ������_IN:ָ��ԭ���˽��㲿��,Ӧ���˸����ֽ��㷽ʽ,Ϊ��ʱȱʡ�˸��ֽ�,�����˸�ָ���Ľ��㷽ʽ 

  --���α������ж��Ƿ񵥶��ղ�����,���ҺŻ��ܱ��� 
  Cursor c_Registinfo(v_״̬ ������ü�¼.��¼״̬%Type) Is
    Select a.����ʱ��, a.�Ǽ�ʱ��, c.����ʱ��, a.�շ�ϸĿid As ��Ŀid, c.ִ�в���id As ����id, c.ִ���� As ҽ������, d.Id As ҽ��id, c.�ű� As ����
    From ������ü�¼ A, ���˹Һż�¼ C, ��Ա�� D
    Where a.��¼���� = 4 And a.No = ���ݺ�_In And a.No = c.No And a.��¼״̬ = v_״̬ And c.ִ���� = d.����(+) And Rownum < 2;
  r_Registrow c_Registinfo%RowType;

  --���α������жϼ�¼�Ƿ����,�����û��ܱ��� 
  Cursor c_Moneyinfo(v_״̬ ������ü�¼.��¼״̬%Type) Is
    Select ���˿���id, ��������id, ִ�в���id, ������Ŀid, Nvl(Sum(Ӧ�ս��), 0) As Ӧ��, Nvl(Sum(ʵ�ս��), 0) As ʵ��, Nvl(Sum(���ʽ��), 0) As ����
    From ������ü�¼
    Where ��¼���� = 4 And ��¼״̬ = v_״̬ And NO = ���ݺ�_In
    Group By ���˿���id, ��������id, ִ�в���id, ������Ŀid;
  r_Moneyrow c_Moneyinfo%RowType;

  --�ù�����ڴ�����Ա�ɿ�������˵Ĳ�ͬ���㷽ʽ�Ľ�� 
  Cursor c_Opermoney(n_Id ����Ԥ����¼.����id%Type) Is
    Select Distinct b.���㷽ʽ, -1 * Nvl(b.��Ԥ��, 0) As ��Ԥ��
    From ����Ԥ����¼ B
    Where b.����id = n_Id And b.��¼���� = 4 And b.��¼״̬ = 2 And Nvl(b.��Ԥ��, 0) <> 0;

  v_Err_Msg Varchar(255);
  Err_Item Exception;

  n_����id ����Ԥ����¼.����id%Type;
  n_����id ������ü�¼.����id%Type;

  v_��ָ�����㷽ʽ ����Ԥ����¼.���㷽ʽ%Type;
  n_�˿���       ����Ԥ����¼.��Ԥ��%Type;
  n_��ӡid         Ʊ�ݴ�ӡ����.Id%Type;
  n_����id         ������Ϣ.����id%Type;
  n_�˷ѽ��       ����Ԥ����¼.��Ԥ��%Type;
  n_Ԥ�����       ����Ԥ����¼.��Ԥ��%Type; --ԭ��¼ Ԥ���ɿ��� 
  n_����ֵ         �������.Ԥ�����%Type;
  n_�Һ�id         ���˹Һż�¼.Id%Type;
  n_��id           ����ɿ����.Id%Type;

  n_�����˷�       Number; --��¼�Ƿ��Ǵ˵��ݵĵڶ����˷� 
  n_����̨ǩ���Ŷ� Number;
  n_ԤԼ���ɶ���   Number;
  n_ԤԼ�Һ�       Number;
  n_�Һ����ɶ���   Number;
  d_Date           Date;
  n_����           ������ü�¼.���ʷ���%Type;
  n_����id1        ������Ϣ.����id%Type;
  n_���ض�         ������ü�¼.ʵ�ս��%Type;
  n_�ѽ���         Number;
  n_���           ���˹Һż�¼.����%Type;
  n_���ﲡ��id     ������Ϣ.����id%Type;
  d_����ʱ��       ����ǼǼ�¼.����ʱ��%Type;
  n_�����¼id     �ٴ������¼.Id%Type;
  v_��������       Varchar2(5000);
  v_��ǰ����       Varchar2(1000);
  v_����ids        Varchar2(500);
  v_Temp           Varchar2(500);
  v_���㷽ʽ       ����Ԥ����¼.���㷽ʽ%Type;
  n_��������־     Number;
  n_������       ����Ԥ����¼.��Ԥ��%Type;
  n_�����         Number;
Begin
  n_��id           := Zl_Get��id(����Ա����_In);
  v_��ָ�����㷽ʽ := ��ָ������_In;

  Select �����¼id, ���� Into n_�����¼id, n_��� From ���˹Һż�¼ Where NO = ���ݺ�_In And Rownum < 2;

  --�����ж�Ҫ�˺�/ȡ��ԤԼ�ļ�¼�Ƿ���� 
  Open c_Moneyinfo(1);
  Fetch c_Moneyinfo
    Into r_Moneyrow;
  If c_Moneyinfo%RowCount = 0 Then
    Close c_Moneyinfo;
    Open c_Moneyinfo(0);
    Fetch c_Moneyinfo
      Into r_Moneyrow;
    If c_Moneyinfo%RowCount = 0 Then
      v_Err_Msg := 'Ҫ����ĵ��ݲ����ڡ�';
      Raise Err_Item;
    End If;
    n_ԤԼ�Һ� := 1;
  End If;
  Close c_Moneyinfo;

  Begin
    Select Zl_Fun_Regcustomname Into v_Temp From Dual;
    If v_Temp Is Not Null Then
      v_����ids := Substr(v_Temp, Instr(v_Temp, '|') + 1);
    End If;
  Exception
    When Others Then
      v_����ids := Null;
  End;

  --1.ԤԼ���� 
  If Nvl(n_ԤԼ�Һ�, 0) = 1 Then
    --������Լ�� 
    Open c_Registinfo(0);
    Fetch c_Registinfo
      Into r_Registrow;
  
    n_����� := Null;
    Update �ٴ������¼ Set ��Լ�� = Nvl(��Լ��, 0) - 1 Where ID = n_�����¼id Returning ��Լ�� Into n_�����;
    If Nvl(n_�����, 0) < 0 Then
      Update �ٴ������¼ Set ��Լ�� = 0 Where ID = n_�����¼id;
    End If;
  
    Update ���˹ҺŻ���
    Set ��Լ�� = Nvl(��Լ��, 0) - 1
    Where ���� = Trunc(r_Registrow.����ʱ��) And Nvl(ҽ��id, 0) = Nvl(r_Registrow.ҽ��id, 0) And
          Nvl(ҽ������, '-') = Nvl(r_Registrow.ҽ������, '-') And ����id = r_Registrow.����id And ��Ŀid = r_Registrow.��Ŀid And
          (���� = r_Registrow.���� Or ���� Is Null);
  
    If Sql%RowCount = 0 Then
      Insert Into ���˹ҺŻ���
        (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, ��Լ��)
      Values
        (Trunc(r_Registrow.����ʱ��), r_Registrow.����id, r_Registrow.��Ŀid, r_Registrow.ҽ������,
         Decode(r_Registrow.ҽ��id, 0, Null, r_Registrow.ҽ��id), r_Registrow.����, -1);
    End If;
  
    Close c_Registinfo;
  
    --���¹Һ����״̬ 
    Update �ٴ�������ſ���
    Set �Һ�״̬ = 0, ����Ա���� = Null
    Where �Һ�״̬ = 2 And ��¼id = n_�����¼id And ��� = n_���;
  
    Update �ٴ�������ſ���
    Set �Һ�״̬ = 4, ����Ա���� = Null
    Where �Һ�״̬ = 2 And ��¼id = n_�����¼id And ��ע = To_Char(n_���);
  
    --��Ӳ��˹Һż�¼�� ������¼ 
    Select ���˹Һż�¼_Id.Nextval, Sysdate Into n_�Һ�id, d_Date From Dual;
  
    Update ���˹Һż�¼ Set ��¼״̬ = 3 Where NO = ���ݺ�_In And ��¼״̬ = 1 And ��¼���� = 2;
    If Sql%NotFound Then
      v_Err_Msg := 'ԤԼ����' || ���ݺ�_In || '�������ڻ����ڲ���ԭ���Ѿ���ȡ��ԤԼ';
      Raise Err_Item;
    End If;
  
    Insert Into ���˹Һż�¼
      (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����,
       ����, ����, ����, ԤԼ, ժҪ, ԤԼ��ʽ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ҽ�Ƹ��ʽ, �����¼id, ԤԼ����Ա, ԤԼ����Ա���)
      Select n_�Һ�id, NO, ��¼����, 2, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, d_Date, ����ʱ��,
             ����Ա���_In, ����Ա����_In, ����, ����, ����, ԤԼ, Nvl(ժҪ_In, ժҪ) As ժҪ, ԤԼ��ʽ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ҽ�Ƹ��ʽ,
             n_�����¼id, ԤԼ����Ա, ԤԼ����Ա���
      From ���˹Һż�¼
      Where NO = ���ݺ�_In;
  
    Update ������ü�¼ Set ��¼״̬ = 3 Where NO = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 0;
    Insert Into ������ü�¼
      (ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, ���ʵ�id, ����id, ҽ�����, �����־, ���ʷ���, ����, �Ա�, ����, ��ʶ��, ���ʽ, ���˿���id, �ѱ�, �շ����,
       �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id,
       ִ����, ִ��״̬, ִ��ʱ��, ����, ����Ա���, ����Ա����, ����id, ���ʽ��, ���մ���id, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���, �ɿ���id, ����״̬, ��ת��)
      Select ���˷��ü�¼_Id.Nextval, ��¼����, NO, ʵ��Ʊ��, 2, ���, ��������, �۸񸸺�, ���ʵ�id, ����id, ҽ�����, �����־, ���ʷ���, ����, �Ա�, ����, ��ʶ��, ���ʽ,
             ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, -1 * ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, -1 * Ӧ�ս��,
             -1 * ʵ�ս��, ������, ��������id, ������, ����ʱ��, d_Date, ִ�в���id, ִ����, -1, ִ��ʱ��, ����, ����Ա���_In, ����Ա����_In, Null, Null,
             ���մ���id, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���, �ɿ���id, ����״̬, ��ת��
      From ������ü�¼
      Where NO = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 3;
  
    --���ԤԼ���ɶ���ʱ��Ҫ������� 
  
    n_ԤԼ���ɶ��� := Zl_To_Number(zl_GetSysParameter('ԤԼ���ɶ���', 1113));
    If Nvl(n_ԤԼ���ɶ���, 0) = 1 Then
      --Ҫɾ������ 
      For v_�Һ� In (Select ID, ����, ִ�в���id, ִ���� From ���˹Һż�¼ Where NO = ���ݺ�_In) Loop
        Zl_�ŶӽкŶ���_Delete(v_�Һ�.ִ�в���id, v_�Һ�.Id);
      End Loop;
    End If;
    Return;
  End If;

  Select Nvl(���ʷ���, 0), ����id, Decode(Sign(Nvl(����id, 0)), 0, 0, 1)
  Into n_����, n_����id, n_�ѽ���
  From ������ü�¼
  Where ��¼���� = 4 And NO = ���ݺ�_In And ��¼״̬ In (1, 3) And Rownum < 2;

  --2.�ҺŴ��� 
  n_�ѽ��� := Nvl(n_�ѽ���, 0);

  If n_�ѽ��� = 1 And n_���� = 1 Then
    Select Sysdate, Null Into d_Date, n_����id From Dual;
  Else
    Select Sysdate, ���˽��ʼ�¼_Id.Nextval Into d_Date, n_����id From Dual;
  End If;

  ----0-ȫ�� 1-�˹Һŷ� 2-�˲����� 
  If Nvl(�˷�����_In, 0) <> 2 Then
    --���ǹ��˲�����ʱ���� 
    --���¹Һ����״̬ 
    If �˺�����_In = 1 Then
      Update �ٴ�������ſ���
      Set �Һ�״̬ = 0, ����Ա���� = Null
      Where �Һ�״̬ = 1 And ��¼id = n_�����¼id And ��� = n_���;
    
      Update �ٴ�������ſ���
      Set �Һ�״̬ = 4, ����Ա���� = Null
      Where �Һ�״̬ = 1 And ��¼id = n_�����¼id And ��ע = To_Char(n_���);
    Else
      Update �ٴ�������ſ���
      Set �Һ�״̬ = 4, ����Ա���� = ����Ա����_In
      Where �Һ�״̬ = 1 And ��¼id = n_�����¼id And (��� = n_��� Or ��ע = To_Char(n_���));
    End If;
  
    --���˾���״̬ 
    If n_����id Is Not Null Then
      Update ������Ϣ Set ����״̬ = 0, �������� = Null Where ����id = n_����id;
    
      --ɾ���������ش���,ֻ�е�ֻ��һ���Һż�¼���Ҳ��˽���������Һ����ڽ���ʱ�Żᴦ�� 
      If ɾ�������_In = 1 Then
        Delete ���ﲡ����¼ Where ����id = n_����id;
        Update ������Ϣ Set ����� = Null Where ����id = n_����id;
        --���ü�¼�����Һż����������￨����,�Լ����˽��Ѻ��˷ѻ����ʵķ���,�Һż�¼������� 
        Update ������ü�¼ Set ��ʶ�� = Null Where �����־ = 1 And ����id = n_����id;
      End If;
    End If;
  
    --�����ʱ���˾��￨��,�˷�ʱ������￨��,�ڷǹ��˲�����ʱ 
    n_����id1 := Null;
    Begin
      Select ����id
      Into n_����id1
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And ���ӱ�־ = 2 And Rownum < 2;
    Exception
      When Others Then
        Null;
    End;
  
    If n_����id1 Is Not Null And Nvl(�˷�����_In, 0) <> 2 Then
      Update ������Ϣ
      Set ���￨�� = Null, ����֤�� = Null, Ic���� = Decode(Ic����, ���￨��, Null, Ic����)
      Where ����id = n_����id1;
    End If;
  
  End If;

  --���ǰ���Ƿ��Ѿ������˹����� 
  Begin
    Select 1 Into n_�����˷� From ������ü�¼ Where ��¼���� = 4 And NO = ���ݺ�_In And ��¼״̬ = 3 And Rownum < 2;
  Exception
    When Others Then
      n_�����˷� := 0;
  End;

  --������ü�¼ 
  --������¼ 
  If Nvl(�˷�����_In, 0) = 0 Or Nvl(�˷�����_In, 0) = 2 Then
    --ȫ��,�˲����� 
    --������ü�¼��������¼ 
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
       ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��,
       ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����,
             �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ִ����,
             ����Ա���_In, ����Ա����_In, ����ʱ��, d_Date, n_����id,
             Decode(n_����, 1, Decode(Nvl(n_�ѽ���, 0), 0, -1 * ʵ�ս��, Null), -1 * ���ʽ��), ������Ŀ��, ���մ���id, -1 * ͳ����,
             Nvl(ժҪ_In, ժҪ) As ժҪ, Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
  
    --ԭʼ��¼ 
    If n_���� = 1 And Nvl(n_�ѽ���, 0) = 0 Then
      Update ������ü�¼
      Set ��¼״̬ = 3, ����id = n_����id, ���ʽ�� = ʵ�ս��
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    Else
      Update ������ü�¼
      Set ��¼״̬ = 3
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    End If;
  Elsif Nvl(�˷�����_In, 0) = 1 Or Nvl(�˷�����_In, 0) = 4 Then
    --�˹Һŷ�,�˹Һ��벡���� 
    --������ü�¼��������¼ 
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
       ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��,
       ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����,
             �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ִ����,
             ����Ա���_In, ����Ա����_In, ����ʱ��, d_Date, n_����id,
             Decode(n_����, 1, Decode(Nvl(n_�ѽ���, 0), 0, -1 * ʵ�ս��, Null), -1 * ���ʽ��), ������Ŀ��, ���մ���id, -1 * ͳ����,
             Nvl(ժҪ_In, ժҪ) As ժҪ, Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') = 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
  
    --ԭʼ��¼ 
    If n_���� = 1 And Nvl(n_�ѽ���, 0) = 0 Then
      Update ������ü�¼
      Set ��¼״̬ = 3, ����id = n_����id, ���ʽ�� = ʵ�ս��
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') = 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    Else
      Update ������ü�¼
      Set ��¼״̬ = 3
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') = 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    End If;
  Elsif Nvl(�˷�����_In, 0) = 3 Then
    --�˸��ӷ� 
    --������ü�¼��������¼ 
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
       ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��,
       ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����,
             �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ִ����,
             ����Ա���_In, ����Ա����_In, ����ʱ��, d_Date, n_����id,
             Decode(n_����, 1, Decode(Nvl(n_�ѽ���, 0), 0, -1 * ʵ�ս��, Null), -1 * ���ʽ��), ������Ŀ��, ���մ���id, -1 * ͳ����,
             Nvl(ժҪ_In, ժҪ) As ժҪ, Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') > 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
  
    --ԭʼ��¼ 
    If n_���� = 1 And Nvl(n_�ѽ���, 0) = 0 Then
      Update ������ü�¼
      Set ��¼״̬ = 3, ����id = n_����id, ���ʽ�� = ʵ�ս��
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') > 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    Else
      Update ������ü�¼
      Set ��¼״̬ = 3
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Instr(',' || v_����ids || ',', ',' || �շ�ϸĿid || ',') > 0 And
            Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
    End If;
  Elsif Nvl(�˷�����_In, 0) = 5 Then
    --�˹Һ��븽�ӷ� 
    --������ü�¼��������¼ 
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
       ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ִ����, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��,
       ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, �۸񸸺�, ��������, ����id, ���˿���id, �����־, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, �շ����,
             �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ִ����,
             ����Ա���_In, ����Ա����_In, ����ʱ��, d_Date, n_����id,
             Decode(n_����, 1, Decode(Nvl(n_�ѽ���, 0), 0, -1 * ʵ�ս��, Null), -1 * ���ʽ��), ������Ŀ��, ���մ���id, -1 * ͳ����,
             Nvl(ժҪ_In, ժҪ) As ժҪ, Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id
      From ������ü�¼
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Nvl(���ӱ�־, 0) <> 1;
  
    --ԭʼ��¼ 
    If n_���� = 1 And Nvl(n_�ѽ���, 0) = 0 Then
      Update ������ü�¼
      Set ��¼״̬ = 3, ����id = n_����id, ���ʽ�� = ʵ�ս��
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Nvl(���ӱ�־, 0) <> 1;
    Else
      Update ������ü�¼
      Set ��¼״̬ = 3
      Where ��¼���� = 4 And ��¼״̬ = 1 And NO = ���ݺ�_In And Nvl(���ӱ�־, 0) <> 1;
    End If;
  End If;

  n_����id := 0;
  If n_���� = 0 Then
    --��ȡ����ID 
    Select Nvl(����id, 0)
    Into n_����id
    From ������ü�¼
    Where ��¼���� = 4 And ��¼״̬ = 3 And NO = ���ݺ�_In And
          Nvl(���ӱ�־, 0) = Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0)) And
          Rownum = 1;
  End If;

  If n_���� = 1 Then
    --���� 
    For c_���� In (Select ʵ�ս��, ���˿���id, ��������id, ִ�в���id, ������Ŀid
                 From ������ü�¼
                 Where ��¼���� = 4 And ��¼״̬ = 3 And NO = ���ݺ�_In And
                       Nvl(���ӱ�־, 0) =
                       Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0)) And
                       Nvl(���ʷ���, 0) = 1) Loop
      --������� 
      Update �������
      Set ������� = Nvl(�������, 0) - Nvl(c_����.ʵ�ս��, 0)
      Where ����id = Nvl(n_����id, 0) And ���� = 1 And ���� = 1
      Returning ������� Into n_���ض�;
    
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, ����, ����, �������, Ԥ�����)
        Values
          (n_����id, 1, 1, -1 * Nvl(c_����.ʵ�ս��, 0), 0);
        n_���ض� := Nvl(c_����.ʵ�ս��, 0);
      End If;
      If Nvl(n_���ض�, 0) = 0 Then
        Delete �������
        Where ����id = Nvl(n_����id, 0) And ���� = 1 And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
      End If;
      --����δ����� 
      Update ����δ�����
      Set ��� = Nvl(���, 0) - Nvl(c_����.ʵ�ս��, 0)
      Where ����id = n_����id And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���˿���id, 0) = Nvl(c_����.���˿���id, 0) And
            Nvl(��������id, 0) = Nvl(c_����.��������id, 0) And Nvl(ִ�в���id, 0) = Nvl(c_����.ִ�в���id, 0) And ������Ŀid + 0 = c_����.������Ŀid And
            ��Դ;�� + 0 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into ����δ�����
          (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
        Values
          (n_����id, Null, Null, c_����.���˿���id, c_����.��������id, c_����.ִ�в���id, c_����.������Ŀid, 1, -1 * Nvl(c_����.ʵ�ս��, 0));
      End If;
    End Loop;
    Delete ����δ�����
    Where ����id = n_����id And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���, 0) = 0 And ��Դ;�� + 0 = 1;
  End If;

  If n_���� = 0 Then
    --1.�˷� 
    --���˹ҺŽ���:�ֽ�͸����ʻ����� 
    If ���㷽ʽ_In Is Null And Nvl(��Ԥ��_In, 0) = 0 Then
      If ��ԭ���˽���_In Is Not Null Then
        --�˿����ȡ 
        If Nvl(�˷�����_In, 0) <> 0 Or Nvl(n_�����˷�, 0) <> 0 Then
          --����ǵ����˲�����,����ֻ�˹Һŷ�,�Ȼ�ȡ�˷ѽ�� 
          Begin
            --��ȡ�����˿��� 
            Select Sum(Nvl(ʵ�ս��, 0)) As �տ���
            Into n_�˿���
            From ������ü�¼
            Where NO = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 3 And
                  Nvl(���ӱ�־, 0) =
                  Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
          
          Exception
            When Others Then
              v_Err_Msg := '���ݡ�' || ���ݺ�_In || '����' || Case Nvl(�˷�����_In, 0)
                             When 1 Then
                              '�Һŷ���'
                             When 2 Then
                              '������'
                           End || '�������ڲ���ԭ���Ѿ��������˷ѻ��ߵ��ݲ�����!';
              Raise Err_Item;
          End;
          Begin
            Select ��Ԥ��
            Into n_�˷ѽ��
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') = 0;
          Exception
            When Others Then
              n_�˷ѽ�� := 0;
          End;
        
          --a.����Ľ��㷽ʽ 
        
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
             ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In, -n_�˿���,
                   n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') = 0;
        
          If n_�˷ѽ�� = 0 Then
            --b.����������ֽ� 
            If n_�˿��� <> 0 Then
              If v_��ָ�����㷽ʽ Is Null Then
                --�˸��ֽ� 
                Begin
                  Select ���� Into v_��ָ�����㷽ʽ From ���㷽ʽ Where ���� = 1;
                Exception
                  When Others Then
                    v_��ָ�����㷽ʽ := '�ֽ�';
                End;
              End If;
              Update ����Ԥ����¼
              Set ��Ԥ�� = ��Ԥ�� + (-1 * n_�˿���), ����˵�� = Nvl(����˵��_In, ����˵��)
              Where ��¼���� = 4 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_��ָ�����㷽ʽ;
              If Sql%RowCount = 0 Then
                Insert Into ����Ԥ����¼
                  (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����,
                   �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
                  Select ����Ԥ����¼_Id.Nextval, a.No, a.ʵ��Ʊ��, a.��¼����, 2, a.����id, a.��ҳid, a.����id, a.ժҪ, v_��ָ�����㷽ʽ, d_Date,
                         ����Ա���_In, ����Ա����_In, -1 * n_�˿���, n_����id, n_��id, Ԥ�����, Decode(����˵��_In, Null, �����id, Null), ���㿨���,
                         Decode(����˵��_In, Null, ����, Null), Decode(����˵��_In, Null, ������ˮ��, Null), Nvl(����˵��_In, ����˵��), ������λ,
                         4
                  From ����Ԥ����¼ A
                  Where a.��¼���� = 4 And a.��¼״̬ = 1 And a.����id = n_����id And Rownum < 2;
              End If;
            End If;
          End If;
        Else
          --a.����Ľ��㷽ʽԭ���� 
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
             ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In, -��Ԥ��,
                   n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') = 0;
        
          --b.����������ֽ� 
          Begin
            Select Sum(��Ԥ��)
            Into n_�˷ѽ��
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') > 0;
          Exception
            When Others Then
              n_�˷ѽ�� := 0;
          End;
          If n_�˷ѽ�� <> 0 Then
            If v_��ָ�����㷽ʽ Is Null Then
              --�˸��ֽ� 
              Begin
                Select ���㷽ʽ
                Into v_��ָ�����㷽ʽ
                From ����Ԥ����¼
                Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id And
                      Instr(',' || ��ԭ���˽���_In || ',', ',' || ���㷽ʽ || ',') = 0;
              
              Exception
                When Others Then
                  Begin
                    Select ���� Into v_��ָ�����㷽ʽ From ���㷽ʽ Where ���� = 1;
                  Exception
                    When Others Then
                      v_��ָ�����㷽ʽ := '�ֽ�';
                  End;
              End;
            End If;
            Update ����Ԥ����¼
            Set ��Ԥ�� = ��Ԥ�� + (-1 * n_�˷ѽ��), ����˵�� = Nvl(����˵��_In, ����˵��)
            Where ��¼���� = 4 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_��ָ�����㷽ʽ;
            If Sql%RowCount = 0 Then
              Insert Into ����Ԥ����¼
                (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����,
                 �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
                Select ����Ԥ����¼_Id.Nextval, a.No, a.ʵ��Ʊ��, a.��¼����, 2, a.����id, a.��ҳid, a.����id, a.ժҪ, v_��ָ�����㷽ʽ, d_Date,
                       ����Ա���_In, ����Ա����_In, -1 * n_�˷ѽ��, n_����id, n_��id, Ԥ�����, Decode(����˵��_In, Null, �����id, Null), ���㿨���,
                       Decode(����˵��_In, Null, ����, Null), Decode(����˵��_In, Null, ������ˮ��, Null), Nvl(����˵��_In, ����˵��), ������λ, 4
                From ����Ԥ����¼ A
                Where a.��¼���� = 4 And a.��¼״̬ = 1 And a.����id = n_����id And Rownum < 2;
            End If;
          End If;
        End If;
      Else
        --�˿����ȡ 
        If Nvl(�˷�����_In, 0) <> 0 Or Nvl(n_�����˷�, 0) <> 0 Then
          --����ǵ����˲�����,����ֻ�˹Һŷ�,�Ȼ�ȡ�˷ѽ�� 
          Begin
            --��ȡ�����˿��� 
            Select Sum(Nvl(ʵ�ս��, 0)) As �տ���
            Into n_�˿���
            From ������ü�¼
            Where NO = ���ݺ�_In And ��¼���� = 4 And ��¼״̬ = 3 And
                  Nvl(���ӱ�־, 0) =
                  Decode(Nvl(�˷�����_In, 0), 2, 1, 1, Decode(Nvl(���ӱ�־, 0), 1, -1, Nvl(���ӱ�־, 0)), Nvl(���ӱ�־, 0));
          Exception
            When Others Then
              v_Err_Msg := '���ݡ�' || ���ݺ�_In || '����' || Case Nvl(�˷�����_In, 0)
                             When 1 Then
                              '�Һŷ���'
                             When 2 Then
                              '������'
                           End || '�������ڲ���ԭ���Ѿ��������˷ѻ��ߵ��ݲ�����!';
              Raise Err_Item;
          End;
        End If;
        If Nvl(n_�����˷�, 0) = 0 And Nvl(�˷�����_In, 0) = 0 Then
          --�״�ȫ�� 
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
             ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                   -1 * ��Ԥ��, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = 1 And ����id = n_����id;
        Else
          --�����˷�,���߱��ε���һ���� 
          --�����˷�ʱ,��¼״̬=3 ,�״β�����,��¼״̬Ϊ1 
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
             ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                   -1 * n_�˿���, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
            From ����Ԥ����¼
            Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id And ժҪ = 'ҽ���Һ�' And
                  ��Ԥ�� = n_�˿��� And Rownum < 2;
          If Sql%RowCount = 0 Then
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
               ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
              Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                     -1 * n_�˿���, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
              From ����Ԥ����¼
              Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id And ��Ԥ�� = n_�˿��� And
                    Rownum < 2;
          End If;
          If Sql%RowCount = 0 Then
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
               ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
              Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                     -1 * n_�˿���, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
              From ����Ԥ����¼
              Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id And Rownum < 2;
            If Sql%RowCount = 0 Then
              --�����˷�,����ȫ��ʹ��Ԥ����ɷ�ʱ�Ŵ��ڴ������ 
              n_Ԥ����� := n_�˿���;
            End If;
          End If;
        
        End If;
      End If;
    Else
      --�����㷽ʽ�� 
      If ���㷽ʽ_In Is Not Null Then
        v_�������� := ���㷽ʽ_In || '|'; --�Կո�ֿ���|��β,û�н������� 
        While v_�������� Is Not Null Loop
          v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '|') - 1);
          v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
        
          v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
          n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1));
        
          v_��ǰ����   := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
          n_��������־ := To_Number(v_��ǰ����);
        
          If n_��������־ = 0 Then
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
               ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������, �������)
              Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, v_���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                     -1 * n_������, n_����id, n_��id, Ԥ�����, Null, Null, Null, Null, ����˵��_In, ������λ, 4, �������
              From ����Ԥ����¼
              Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id And Rownum < 2;
          Else
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
               ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������, �������)
              Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, v_���㷽ʽ, d_Date, ����Ա���_In, ����Ա����_In,
                     -1 * n_������, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, Nvl(����˵��_In, ����˵��), ������λ, 4, �������
              
              From ����Ԥ����¼
              Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id And
                    (�����id Is Not Null Or ���㿨��� Is Not Null) And Rownum < 2;
          End If;
          v_�������� := Substr(v_��������, Instr(v_��������, '|') + 1);
        End Loop;
      End If;
      n_Ԥ����� := Nvl(��Ԥ��_In, 0);
    End If;
    --�״��˷�ʱ,��¼״̬�����Ϊ��3 
    Update ����Ԥ����¼
    Set ��¼״̬ = 3
    Where ��¼���� = 4 And ��¼״̬ = Decode(Nvl(n_�����˷�, 0), 0, 1, 3) And ����id = n_����id;
  
    --��Ԥ�� 1-ȫ�� 2-������,������ʱ��ȫ��ʹ��Ԥ�����нɿ� 
    If Nvl(�˷�����_In, 0) = 0 Or (Nvl(�˷�����_In, 0) <> 0 And n_Ԥ����� <> 0) Then
      --���˹ҺŽ���:��Ԥ����� 
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
         ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
        Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, d_Date,
               ����Ա����_In, ����Ա���_In, -1 * Decode(Nvl(�˷�����_In, 0) + Nvl(n_�����˷�, 0), 0, ��Ԥ��, n_Ԥ�����), n_����id, n_��id, Ԥ�����,
               �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 4
        From ����Ԥ����¼
        Where ��¼���� In (1, 11) And ����id = n_����id And Nvl(��Ԥ��, 0) <> 0 And
              Rownum = Decode(Nvl(�˷�����_In, 0) + Nvl(n_�����˷�, 0), 0, Rownum, 1);
    End If;
  
    --������Ԥ����� 
    For c_Ԥ�� In (Select ����id, Ԥ�����, -1 * Sum(Nvl(��Ԥ��, 0)) As ��Ԥ��
                 From ����Ԥ����¼
                 Where ��¼���� In (1, 11) And ����id = n_����id
                 Group By ����id, Ԥ�����) Loop
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(c_Ԥ��.��Ԥ��, 0)
      Where ����id = c_Ԥ��.����id And ���� = Nvl(c_Ԥ��.Ԥ�����, 2) And ���� = 1
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, Ԥ�����, ����, ����)
        Values
          (c_Ԥ��.����id, Nvl(c_Ԥ��.��Ԥ��, 0), 1, Nvl(c_Ԥ��.Ԥ�����, 2));
        n_����ֵ := Nvl(c_Ԥ��.��Ԥ��, 0);
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From �������
        Where ����id = c_Ԥ��.����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
      End If;
    End Loop;
  
    If �ջ�Ʊ�ݺ�_In Is Not Null Then
      --���˹Һŷ�,������Ʊ�� 
      --�˿��ջ�Ʊ��(�����ϴιҺ�ʹ��Ʊ��,�����ջ�) 
      Begin
        --�����һ�δ�ӡ��������ȡ 
        --81907 
        Select ID
        Into n_��ӡid
        From (Select b.Id
               From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
               Where a.��ӡid = b.Id And a.���� = 1 And a.ԭ�� In (1, 3) And b.�������� = 4 And b.No = ���ݺ�_In
               Order By a.ʹ��ʱ�� Desc)
        Where Rownum < 2;
      Exception
        When Others Then
          n_��ӡid := Null;
      End;
    
      --���ջ�ԭƱ�� 
      If n_��ӡid Is Not Null Then
        Begin
          Insert Into Ʊ��ʹ����ϸ
            (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����)
            Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_Date, ����Ա����_In
            From Ʊ��ʹ����ϸ
            Where ��ӡid = n_��ӡid And ���� = 1 And Instr(',' || �ջ�Ʊ�ݺ�_In || ',', ',' || ���� || ',') > 0;
        Exception
          When Others Then
            Delete From Ʊ��ʹ����ϸ Where ��ӡid = n_��ӡid And ���� = 2 And ԭ�� = 2;
            Insert Into Ʊ��ʹ����ϸ
              (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����)
              Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_Date, ����Ա����_In
              From Ʊ��ʹ����ϸ
              Where ��ӡid = n_��ӡid And ���� = 1 And Instr(',' || �ջ�Ʊ�ݺ�_In || ',', ',' || ���� || ',') > 0;
        End;
      End If;
    End If;
  End If;

  --�����˲�������,��������ܼ�¼ 
  --��ػ��ܱ�Ĵ��� 
  If Nvl(�˷�����_In, 0) Not In (2, 3) Then
    --���˹ҺŻ��� 
    Open c_Registinfo(3);
    Fetch c_Registinfo
      Into r_Registrow;
    If c_Registinfo%RowCount = 0 Then
      --ֻ�ղ�����ʱ�޺ű�,������ 
      Close c_Registinfo;
    Else
    
      --��Ҫȷ���Ƿ�ԤԼ�Һ� 
      --1.�������ԤԼ�ҺŲ����ĹҺż�¼,����Ҫ���ѹ����������ѽ��� 
      --2.����������Һ�,��ֻ���ѹ��� 
      Begin
        Select Decode(ԤԼ, Null, 0, 0, 0, 1) Into n_ԤԼ�Һ� From ���˹Һż�¼ Where NO = ���ݺ�_In And Rownum = 1;
      Exception
        When Others Then
          n_ԤԼ�Һ� := 0;
      End;
      n_����� := Null;
      Update �ٴ������¼
      Set �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - n_ԤԼ�Һ�, ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ�Һ�
      Where ID = n_�����¼id
      Returning �ѹ��� Into n_�����;
    
      If Nvl(n_�����, 0) < 0 Then
        Update �ٴ������¼ Set �ѹ��� = 0 Where ID = n_�����¼id;
      End If;
    
      Update ���˹ҺŻ���
      Set �ѹ��� = Nvl(�ѹ���, 0) - 1, �����ѽ��� = Nvl(�����ѽ���, 0) - n_ԤԼ�Һ�, ��Լ�� = Nvl(��Լ��, 0) - n_ԤԼ�Һ�
      Where ���� = Trunc(r_Registrow.����ʱ��) And Nvl(ҽ��id, 0) = Nvl(r_Registrow.ҽ��id, 0) And
            Nvl(ҽ������, '-') = Nvl(r_Registrow.ҽ������, '-') And ����id = r_Registrow.����id And ��Ŀid = r_Registrow.��Ŀid And
            (���� = r_Registrow.���� Or ���� Is Null);
    
      If Sql%RowCount = 0 Then
        Insert Into ���˹ҺŻ���
          (����, ����id, ��Ŀid, ҽ������, ҽ��id, ����, �ѹ���, �����ѽ���, ��Լ��)
        Values
          (Trunc(r_Registrow.����ʱ��), r_Registrow.����id, r_Registrow.��Ŀid, r_Registrow.ҽ������,
           Decode(r_Registrow.ҽ��id, 0, Null, r_Registrow.ҽ��id), r_Registrow.����, -1, -1 * n_ԤԼ�Һ�, -1 * n_ԤԼ�Һ�);
      End If;
    
      Close c_Registinfo;
    End If;
  End If;

  If n_���� = 0 Then
    --��Ա�ɿ����(���������ʻ��ȵĽ�����,�����˳�Ԥ����) 
    For r_Opermoney In c_Opermoney(n_����id) Loop
      Update ��Ա�ɿ����
      Set ��� = Nvl(���, 0) + (-1 * r_Opermoney.��Ԥ��)
      Where �տ�Ա = ����Ա����_In And ���㷽ʽ = r_Opermoney.���㷽ʽ And ���� = 1
      Returning ��� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ��Ա�ɿ����
          (�տ�Ա, ���㷽ʽ, ����, ���)
        Values
          (����Ա����_In, r_Opermoney.���㷽ʽ, 1, -1 * r_Opermoney.��Ԥ��);
        n_����ֵ := r_Opermoney.��Ԥ��;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From ��Ա�ɿ����
        Where �տ�Ա = ����Ա����_In And ���㷽ʽ = r_Opermoney.���㷽ʽ And ���� = 1 And Nvl(���, 0) = 0;
      End If;
    End Loop;
  End If;
  If Nvl(�˷�����_In, 0) Not In (2, 3) Then
    n_�Һ����ɶ��� := Zl_To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113));
    If n_�Һ����ɶ��� <> 0 Then
      n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113));
      If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Then
      
        --Ҫɾ������ 
        For v_�Һ� In (Select ID, ����, ִ�в���id, ִ���� From ���˹Һż�¼ Where NO = ���ݺ�_In) Loop
          Zl_�ŶӽкŶ���_Delete(v_�Һ�.ִ�в���id, v_�Һ�.Id);
        End Loop;
      End If;
    End If;
  
    --ҽ�������ľ���ǼǼ�¼ 
    Begin
      Select ����id, ����ʱ�� Into n_���ﲡ��id, d_����ʱ�� From ���˹Һż�¼ Where NO = ���ݺ�_In;
      Delete From ����ǼǼ�¼ Where ����id = n_���ﲡ��id And ����ʱ�� = d_����ʱ�� And ��ҳid Is Null;
    Exception
      When Others Then
        Null;
    End;
  
    --���˹Һż�¼ 
    Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
  
    Update ���˹Һż�¼ Set ��¼״̬ = 3 Where NO = ���ݺ�_In And ��¼���� = 1 And ��¼״̬ = 1;
    If Sql%NotFound Then
      v_Err_Msg := '�Һŵ���' || ���ݺ�_In || '�������ڻ����ڲ���ԭ���Ѿ����˺�';
      Raise Err_Item;
    End If;
    Insert Into ���˹Һż�¼
      (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����,
       ����, ����, ����, ԤԼ, ժҪ, ԤԼ��ʽ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ����, ҽ�Ƹ��ʽ, �����¼id)
      Select n_�Һ�id, NO, ��¼����, 2, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, d_Date, ����ʱ��,
             ����Ա���_In, ����Ա����_In, ����, ����, ����, ԤԼ, Nvl(ժҪ_In, ժҪ) As ժҪ, ԤԼ��ʽ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ����, ҽ�Ƹ��ʽ,
             n_�����¼id
      From ���˹Һż�¼
      Where NO = ���ݺ�_In;
  End If;
  --��Ϣ���� 
  Begin
    Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
      Using 2, ���ݺ�_In;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˹Һż�¼_����_Delete;
/

--127548:�ƽ�,2018-06-21,RISȡ���Ǽ�ʱɾ��ԤԼ
--125867:�ƽ�,2018-05-23,RIS�ӿڳ�Ժ������δ�ɷ��ò�����ִ�з���
Create Or Replace Package b_Zlxwinterface Is
  Type t_Refcur Is Ref Cursor;

  --1������RIS״̬�ı�
  Procedure Receiverisstate
  (
    ҽ��id_In   ����ҽ������.ҽ��id%Type,
    Risid_In    ����ҽ������.Risid%Type,
    ״̬_In     Number,
    ������Ա_In ����ҽ������.�����%Type,
    ִ��ʱ��_In ����ҽ������.���ʱ��%Type := Null,
    ִ��˵��_In ����ҽ������.ִ��˵��%Type := Null,
    ����ִ��_In Number := 0
  );

  --2������ȷ��
  Procedure Ӱ�����ִ��
  (
    ҽ��id_In     Ӱ�����¼.ҽ��id%Type,
    ����ִ��_In   Number := 0,
    ����Ա���_In ��Ա��.���%Type := Null,
    ����Ա����_In ��Ա��.����%Type := Null,
    ִ�в���id_In ������ü�¼.ִ�в���id%Type := Null
  );

  --3��ȡ������ȷ��
  Procedure Ӱ�����ִ��_Cancel
  (
    ҽ��id_In     Ӱ�����¼.ҽ��id%Type,
    ����ִ��_In   Number := 0,
    ����Ա���_In ��Ա��.���%Type := Null,
    ����Ա����_In ��Ա��.����%Type := Null,
    ִ�в���id_In ������ü�¼.ִ�в���id%Type := Null
  );

  --4������RIS�ı���
  Procedure Receivereport
  (
    ҽ��id_In   ����ҽ������.ҽ��id%Type,
    Risid_In    ����ҽ������.Risid%Type,
    ��������_In ���Ӳ�������.�����ı�%Type,
    �������_In ���Ӳ�������.�����ı�%Type,
    ���潨��_In ���Ӳ�������.�����ı�%Type,
    ����ҽ��_In ���Ӳ�����¼.������%Type
  );

  --5���޸����뵥��Ϣ
  Procedure Ӱ������Ϣ_�޸�
  (
    ҽ��id_In       ����ҽ����¼.Id%Type,
    ����_In         ������Ϣ.����%Type,
    �Ա�_In         ������Ϣ.�Ա�%Type,
    ����_In         ������Ϣ.����%Type,
    �ѱ�_In         ������Ϣ.�ѱ�%Type,
    ҽ�Ƹ��ʽ_In ������Ϣ.ҽ�Ƹ��ʽ%Type,
    ����_In         ������Ϣ.����%Type,
    ����_In         ������Ϣ.����״��%Type,
    ְҵ_In         ������Ϣ.ְҵ%Type,
    ���֤��_In     ������Ϣ.���֤��%Type,
    ��ͥ��ַ_In     ������Ϣ.��ͥ��ַ%Type,
    ��ͥ�绰_In     ������Ϣ.��ͥ�绰%Type,
    ��ͥ��ַ�ʱ�_In ������Ϣ.��ͥ��ַ�ʱ�%Type,
    ��������_In     ������Ϣ.��������%Type := Null
  );

  --6��ȡ�����뵥��Ϣ
  Procedure ȡ��������뵥
  (
    ҽ��id_In     ����ҽ��ִ��.ҽ��id%Type,
    ����Ա���_In ��Ա��.���%Type := Null,
    ����Ա����_In ��Ա��.����%Type := Null,
    ִ�в���id_In ������ü�¼.ִ�в���id%Type := 0,
    �ܾ�ԭ��_In   ����ҽ������.ִ��˵��%Type := Null
  );

  --7������ҽ������ʧ�ܼ�¼
  Procedure Risҽ��ʧ�ܼ�¼_Insert
  (
    ������Դ_In   In Risҽ��ʧ�ܼ�¼.������Դ%Type,
    ����id_In     In Risҽ��ʧ�ܼ�¼.����id%Type,
    ��ҳid_In     In Risҽ��ʧ�ܼ�¼.��ҳid%Type,
    �Һŵ���_In   In Risҽ��ʧ�ܼ�¼.�Һŵ���%Type,
    ���ͺ�_In     In Risҽ��ʧ�ܼ�¼.���ͺ�%Type,
    �������id_In In Risҽ��ʧ�ܼ�¼.�������id%Type,
    ��챨����_In In Risҽ��ʧ�ܼ�¼.��챨����%Type,
    ��������_In   In Risҽ��ʧ�ܼ�¼.��������%Type
  );

  --8������ҽ������ʧ�ܼ�¼
  Procedure Risҽ��ʧ�ܼ�¼_�ط�
  (
    Id_In       In Risҽ��ʧ�ܼ�¼.Id%Type,
    ��������_In In Number
  );

  --9�����˺��½�סԺ���˵���
  Procedure ����ҽ��_�ؽ�����
  (
    ҽ��id_In In ����ҽ������.ҽ��id%Type,
    No_In     In ����ҽ������.No%Type,
    Action_In In Number
  );

  --10����ӡRIS���ԤԼ֪ͨ��
  Procedure Ris���ԤԼ_��ӡ(ҽ��id_In In Ris���ԤԼ.ҽ��id%Type);

  --11������RIS�ֿ������ò���
  Procedure Ris���ÿ���_Update
  (
    �������_In Ris���ÿ���.�������%Type,
    ����_In     Ris���ÿ���.����%Type,
    ����ids_In  Varchar2,
    ��������_In Number
  );

  --12��ɾ��RIS�ֿ������ò���
  Procedure Ris���ÿ���_Delete;

  --13������Ԫ������ȡ��Ϣ
  Function Ris_Replace_Element_Value
  (
    Ԫ����_In   In ����������Ŀ.������%Type,
    ����id_In   In ���Ӳ�����¼.����id%Type,
    ����id_In   In ���Ӳ�����¼.��ҳid%Type,
    ������Դ_In In ���Ӳ�����¼.������Դ%Type,
    ҽ��id_In   In ����ҽ������.ҽ��id%Type
  ) Return Varchar2;

  --14��ɾ��RIS��Ժ���ò���
  Procedure Ris��Ժ����_Delete;

  --15������RISRis��Ժ���ò���
  Procedure Ris��Ժ����_Update
  (
    Id_In           Ris��Ժ����.Id%Type,
    ҽԺ����_In     Ris��Ժ����.ҽԺ����%Type,
    ҽԺ����_In     Ris��Ժ����.ҽԺ����%Type,
    �û���_In       Ris��Ժ����.�û���%Type,
    ����_In         Ris��Ժ����.����%Type,
    ���ݿ������_In Ris��Ժ����.���ݿ������%Type
  );
End b_Zlxwinterface;
/
Create Or Replace Package Body b_Zlxwinterface Is

  --1������RIS״̬�ı�
  Procedure Receiverisstate
  (
    ҽ��id_In   ����ҽ������.ҽ��id%Type,
    Risid_In    ����ҽ������.Risid%Type,
    ״̬_In     Number,
    ������Ա_In ����ҽ������.�����%Type,
    ִ��ʱ��_In ����ҽ������.���ʱ��%Type := Null,
    ִ��˵��_In ����ҽ������.ִ��˵��%Type := Null,
    ����ִ��_In Number := 0
  ) Is
  
    --������ҽ��ID_IN - ����ִ�е�ҽ��ID��
    --      ״̬_IN - -1-ɾ����0-ԤԼ��1-�Ǽǣ�3-�����ɣ�4-�����ֹ��9-�������棻12-������ˣ�15-����
    --     ����ִ��_In -0-ȫ��ִ�У�1-����ִ�У����ҽ������Ƿ���ö�ÿ����Ŀ��ɢ����ִ�еķ�ʽ
  
    Cursor c_Adviceinfo Is
      Select a.Id, a.���id, Nvl(a.���id, a.Id) As ��id, a.�������, a.������Դ, a.ִ�п���id, b.ִ�й���
      From ����ҽ����¼ A, ����ҽ������ B
      Where a.Id = b.ҽ��id And ID = ҽ��id_In;
    r_Adviceinfo c_Adviceinfo%RowType;
  
    v_ִ��״̬ ����ҽ������.ִ��״̬%Type;
    v_ִ�й��� ����ҽ������.ִ�й���%Type;
    n_ִ��     Number; --����Ƿ���Ҫ����״̬��1����Ҫ���£���������Ҫ����
    v_Count    Number;
    v_�����   ����ҽ������.�����%Type;
    v_���ʱ�� ����ҽ������.���ʱ��%Type;
    v_Error    Varchar2(255);
    Err_Custom Exception;
  
  Begin
  
    v_ִ��״̬ := 0;
    v_ִ�й��� := 0;
  
    --��ȡҽ������ҽ��ID������ID
    Open c_Adviceinfo;
    Fetch c_Adviceinfo
      Into r_Adviceinfo;
    Close c_Adviceinfo;
  
    --����״̬_INִ��ҽ��
    ---1-ɾ����0-ԤԼ(��RIS��ʵ���Ͼ���ɾ��)��1-�Ǽǣ�3-�����ɣ�4-�����ֹ��9-�������棻12-������ˣ�13-ȡ����ˣ�14-����ɾ����15-����
  
    If ״̬_In = -1 Or ״̬_In = 0 Then
      v_ִ��״̬ := 0; --δִ��
      v_ִ�й��� := 0;
    Elsif ״̬_In = 1 Then
      v_ִ��״̬ := 3; --����ִ��
      v_ִ�й��� := 2; --�ѱ���
    Elsif ״̬_In = 3 Or ״̬_In = 14 Then
      v_ִ��״̬ := 3; --����ִ��
      v_ִ�й��� := 3; --�Ѽ��
    Elsif ״̬_In = 4 Then
      --���ı�
      v_ִ��״̬ := v_ִ��״̬;
    Elsif ״̬_In = 9 Or ״̬_In = 13 Then
      v_ִ��״̬ := 3; --����ִ��
      v_ִ�й��� := 4; --�ѱ���
    Elsif ״̬_In = 12 Then
      v_ִ��״̬ := 3; --����ִ��
      v_ִ�й��� := 5; --�����
    Elsif ״̬_In = 15 Then
      v_ִ��״̬ := 1; --��ȫִ��
      v_ִ�й��� := 6; --�����
      v_�����   := ������Ա_In;
      v_���ʱ�� := ִ��ʱ��_In;
    End If;
  
    n_ִ�� := 1; --Ĭ�϶�Ҫ����״̬
  
    If ״̬_In = 13 Or ״̬_In = 14 Then
      --ɾ����Ӧ��������
      Delete From ���Ӳ�����¼
      Where ID = (Select ����id From ����ҽ������ Where ҽ��id = ҽ��id_In And Risid = Risid_In);
      Delete From ����ҽ������ Where ҽ��id = ҽ��id_In And Risid = Risid_In;
    
      --ɾ�����ж��Ƿ񻹴��ڱ��棬��������ҽ��״̬���ֲ��䣬������ȫ��ɾ�������ҽ��״̬
      Select Count(1) Into v_Count From ����ҽ������ Where ҽ��id = ҽ��id_In;
    
      If v_Count > 0 Then
        n_ִ�� := 0; --��������ҽ��״̬���ֲ���
      End If;
    End If;
  
    --�����ɾ������ɾ�����е�ԤԼ��Ϣ
    If ״̬_In = -1 Or ״̬_In = 0 Then
      Zl_Ris���ԤԼ_Delete(ҽ��id_In);
    End If;
  
    --����ǵǼǣ����жϴ˼���Ƿ�δִ��
    If ״̬_In = 1 Then
      If r_Adviceinfo.ִ�й��� >= 3 Then
        v_Error := '�����Ѿ���������ˣ������ظ��Ǽǡ�';
        Raise Err_Custom;
      End If;
    End If;
  
    --��ʼִ��ҽ��
    If n_ִ�� = 1 Then
      If Nvl(����ִ��_In, 0) = 1 Then
        -- ������λҽ������ִ��
        Update ����ҽ������
        Set ִ��״̬ = v_ִ��״̬, ִ�й��� = v_ִ�й���, ִ��˵�� = ִ��˵��_In, ����� = v_�����, ���ʱ�� = v_���ʱ��
        Where ҽ��id = ҽ��id_In;
      Else
        Update ����ҽ������
        Set ִ��״̬ = v_ִ��״̬, ִ�й��� = v_ִ�й���, ִ��˵�� = ִ��˵��_In, ����� = v_�����, ���ʱ�� = v_���ʱ��
        Where ҽ��id In (Select ID From ����ҽ����¼ Where (ID = r_Adviceinfo.��id Or ���id = r_Adviceinfo.��id));
      End If;
    End If;
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Receiverisstate;

  --2������ȷ��
  Procedure Ӱ�����ִ��
  (
    ҽ��id_In     Ӱ�����¼.ҽ��id%Type,
    ����ִ��_In   Number := 0,
    ����Ա���_In ��Ա��.���%Type := Null,
    ����Ա����_In ��Ա��.����%Type := Null,
    ִ�в���id_In ������ü�¼.ִ�в���id%Type := Null
  ) Is
    --������ҽ��ID_IN=����ִ�е�ҽ��ID��
    --      ����ִ��_In=���ҽ������Ƿ���ö�ÿ����Ŀ��ɢ����ִ�еķ�ʽ,0-������ִ��
    Cursor c_Advice Is
      Select ID, ���id, Nvl(���id, ID) As ��id, �������, ������Դ From ����ҽ����¼ Where ID = ҽ��id_In;
    r_Advice c_Advice%RowType;
  
    v_Temp     Varchar2(255);
    v_��Ա��� ��Ա��.���%Type;
    v_��Ա���� ��Ա��.����%Type;
    v_����id   ���ű�.Id%Type;
    v_�������� ����ҽ������.��¼����%Type;
    v_���ͺ�   ����ҽ������.���ͺ�%Type;
    v_ִ�й��� ����ҽ������.ִ�й���%Type;
    v_Error    Varchar2(255);
    Err_Custom Exception;
  Begin
  
    --ȡ��ҽ��ID
    Open c_Advice;
    Fetch c_Advice
      Into r_Advice;
    Close c_Advice;
  
    Select ���ͺ�, ִ�й��� Into v_���ͺ�, v_ִ�й��� From ����ҽ������ Where ҽ��id = r_Advice.��id;
  
    --�ǼǺ���ɲ�ִ�з���  2-�Ǽǣ�3-��飬4-���棬5-��ˣ�6-���
    If v_ִ�й��� >= 2 Or v_ִ�й��� <= 6 Then
      --ȡ��ǰ������Ա
      If ����Ա���_In Is Not Null And ����Ա����_In Is Not Null And ִ�в���id_In Is Not Null Then
        v_��Ա��� := ����Ա���_In;
        v_��Ա���� := ����Ա����_In;
        v_����id   := ִ�в���id_In;
      Else
        v_Temp     := Zl_Identity;
        v_����id   := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
        v_��Ա��� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
        v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
      End If;
    
      If r_Advice.������Դ = 2 Then
        Select Decode(��¼����, 1, 1, Decode(�������, 1, 1, 2))
        Into v_��������
        From ����ҽ������
        Where ���ͺ� = v_���ͺ� And ҽ��id = ҽ��id_In;
      Else
        v_�������� := 1;
      End If;
    
      --ִ�з��ú��Զ�����
      If v_�������� = 1 Then
        Zl_����ҽ��ִ��_Finish(ҽ��id_In, v_���ͺ�, ����ִ��_In, v_��Ա���, v_��Ա����, r_Advice.��id, r_Advice.�������, v_����id);
      Else
        Zl_סԺҽ��ִ��_Finish(ҽ��id_In, v_���ͺ�, ����ִ��_In, v_��Ա���, v_��Ա����, r_Advice.��id, r_Advice.�������, v_����id);
      End If;
    End If;
  
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ӱ�����ִ��;

  --3��ȡ������ȷ��
  Procedure Ӱ�����ִ��_Cancel
  (
    ҽ��id_In     Ӱ�����¼.ҽ��id%Type,
    ����ִ��_In   Number := 0,
    ����Ա���_In ��Ա��.���%Type := Null,
    ����Ա����_In ��Ա��.����%Type := Null,
    ִ�в���id_In ������ü�¼.ִ�в���id%Type := Null
  ) Is
    --������
    --      ҽ��ID_IN=����ִ�е�ҽ��ID��
    --      ����ִ��_In=���ҽ������Ƿ���ö�ÿ����Ŀ��ɢ����ִ�еķ�ʽ,0-������ִ��
  
    Cursor c_Advice Is
      Select ID, ���id, Nvl(���id, ID) As ��id From ����ҽ����¼ Where ID = ҽ��id_In;
    r_Advice c_Advice%RowType;
  
    v_���ͺ� ����ҽ������.���ͺ�%Type;
    v_Count  Number;
    v_Error  Varchar2(255);
    Err_Custom Exception;
  
  Begin
  
    --ȡ��ҽ��ID
    Open c_Advice;
    Fetch c_Advice
      Into r_Advice;
    Close c_Advice;
  
    --�ȼ���Ƿ��Ѿ���Ժ��סԺ���ˣ��Ѿ�Ԥ��Ժ���߳�Ժ�ļ�����룬����ִ�з���
    Select Count(*)
    Into v_Count
    From ����ҽ����¼ A, ������ҳ B
    Where a.����id = b.����id And a.��ҳid = b.��ҳid And (b.��Ժ���� Is Not Null Or b.״̬ = 3) And a.Id = r_Advice.��id;
  
    If v_Count > 0 Then
      v_Error := 'סԺ�����Ѿ���Ժ����Ԥ��Ժ������ȡ�����á�';
      Raise Err_Custom;
    End If;
  
    Select ���ͺ� Into v_���ͺ� From ����ҽ������ Where ҽ��id = r_Advice.��id;
  
    --����ͳһ��ҽ��ִ��Cancel����
    Zl_����ҽ��ִ��_Cancel(ҽ��id_In, v_���ͺ�, Null, ����ִ��_In, ִ�в���id_In, ����Ա���_In, ����Ա����_In);
  
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ӱ�����ִ��_Cancel;

  --4������RIS�ı���
  Procedure Receivereport
  (
    ҽ��id_In   ����ҽ������.ҽ��id%Type,
    Risid_In    ����ҽ������.Risid%Type,
    ��������_In ���Ӳ�������.�����ı�%Type,
    �������_In ���Ӳ�������.�����ı�%Type,
    ���潨��_In ���Ӳ�������.�����ı�%Type,
    ����ҽ��_In ���Ӳ�����¼.������%Type
  ) Is
    --��ȡ����ҽ��������������Ϣ
    Cursor c_Advice
    (
      v_��id  Number,
      v_Risid Number
    ) Is
      Select e.Id, e.������Դ, e.����id, e.��ҳid, e.Ӥ��, e.���˿���id, e.�ļ�id, e.��������, e.��������, f.����id, e.ִ�п���id
      From (Select c.Id, c.������Դ, c.����id, c.��ҳid, c.Ӥ��, c.���˿���id, c.�ļ�id, d.���� ��������, d.���� ��������, c.ִ�п���id
             From (Select a.Id, a.������Դ, a.����id, a.��ҳid, a.Ӥ��, a.���˿���id, b.�����ļ�id �ļ�id, a.ִ�п���id
                    From ����ҽ����¼ A, ��������Ӧ�� B
                    Where a.Id = v_��id And a.������Ŀid = b.������Ŀid(+) And b.Ӧ�ó���(+) = Decode(a.������Դ, 2, 2, 4, 4, 1)) C,
                  �����ļ��б� D
             Where c.�ļ�id = d.Id(+)) E, ����ҽ������ F
      Where e.Id = f.ҽ��id(+) And f.Risid(+) = v_Risid;
  
    --�����ļ������Ԫ��
    Cursor c_File(v_File Number) Is
      Select a.Id, a.�ļ�id, a.��id, a.�������, a.��������, a.������, a.��������, a.��������, a.�����д�, a.�����ı�, a.�Ƿ���, a.Ԥ�����id, a.�������,
             a.ʹ��ʱ��, a.����Ҫ��id, a.�滻��, a.Ҫ������, a.Ҫ������, a.Ҫ�س���, a.Ҫ��С��, a.Ҫ�ص�λ, a.Ҫ�ر�ʾ, a.������̬, a.Ҫ��ֵ��
      From �����ļ��ṹ A
      Where a.�ļ�id = v_File
      Order By a.�������;
  
    Cursor c_Report(v_���Ӳ�����¼id Number) Is
      Select b.Id, a.�����ı�
      From ���Ӳ������� A, ���Ӳ������� B
      Where a.�������� = 3 And a.Id = b.��id And b.�������� = 2 And b.��ֹ�� = 0 And a.�ļ�id = v_���Ӳ�����¼id;
  
    Cursor c_Content
    (
      v_�ļ�id Number,
      v_���id Number
    ) Is
      Select a.Id, a.�ļ�id, a.��id, a.�������, a.��������, a.������, a.��������, a.��������, a.�����д�, a.�����ı�, a.�Ƿ���, a.Ԥ�����id, a.�������,
             a.ʹ��ʱ��, a.����Ҫ��id, a.�滻��, a.Ҫ������, a.Ҫ������, a.Ҫ�س���, a.Ҫ��С��, a.Ҫ�ص�λ, a.Ҫ�ر�ʾ, a.������̬, a.Ҫ��ֵ��
      From �����ļ��ṹ A
      Where �ļ�id = v_�ļ�id And ��id = v_���id;
  
    r_Advice        c_Advice%RowType;
    v_����id        ���Ӳ�������.�ļ�id%Type;
    v_��������id    ���Ӳ�������.Id%Type;
    v_��������idnew ���Ӳ�������.Id%Type;
    v_�������      ���Ӳ�������.�������%Type;
    v_��id          ���Ӳ�������.��id%Type;
    v_�����ı�      ���Ӳ�������.�����ı�%Type;
    v_�������id    ���Ӳ�������.�������id%Type;
    --v_��ʽ����    ���Ӳ�����ʽ.����%Type;
    v_Error Varchar2(255);
    Err_Custom Exception;
    v_��ҽ��id ����ҽ������.ҽ��id%Type;
    v_���     Varchar2(300);
    n_����     Number;
    n_Rptcount Number;
    v_�������� ���Ӳ�����¼.��������%Type;
    v_�Һŵ�id ���˹Һż�¼.Id%Type;
  
    Function Getrptno
    (
      v_ҽ��idin   ����ҽ������.ҽ��id%Type,
      v_��������in ���Ӳ�����¼.��������%Type
    ) Return Varchar As
      v_Return Number;
      v_No     Number;
      v_Count  Number;
    Begin
      Select Count(ҽ��id) + 1 Into v_No From ����ҽ������ Where ҽ��id = v_ҽ��idin;
      v_Count := 1;
      While v_Count = 1 Loop
        Select Count(ID)
        Into v_Count
        From ����ҽ������ A, ���Ӳ�����¼ B
        Where a.ҽ��id = v_ҽ��idin And a.����id = b.Id And b.�������� = v_��������in || v_No;
        If v_Count = 1 Then
          v_No := v_No + 1;
        End If;
      End Loop;
      v_Return := v_No;
      Return v_Return;
    End Getrptno;
  
  Begin
  
    -- ��ȡ��ҽ��ID ����ֹ��Ϊ���벿λҽ�������±��汣�����
    Select Nvl(���id, ID) As ��id Into v_��ҽ��id From ����ҽ����¼ Where ID = ҽ��id_In;
  
    Open c_Advice(v_��ҽ��id, Nvl(Risid_In, 0));
    Fetch c_Advice
      Into r_Advice;
  
    If Nvl(r_Advice.�ļ�id, 0) = 0 Then
      v_Error := '���μ����Ŀû�ж�Ӧ��صļ�鱨�棬�������Ա��ϵ��';
      Raise Err_Custom;
    Else
      If Nvl(r_Advice.����id, 0) > 0 Then
        ----����������
        --�ҳ��������д�ı�������к���"%����%","%����%","%����%","%���%",���ô���Ĳ�������
        For r_Report In c_Report(r_Advice.����id) Loop
          If r_Report.�����ı� Like '%����%' Then
            Update ���Ӳ������� Set �����ı� = ��������_In || Chr(13) || Chr(13) Where ID = r_Report.Id;
          Elsif r_Report.�����ı� Like '%���%' Then
            Update ���Ӳ������� Set �����ı� = �������_In || Chr(13) || Chr(13) Where ID = r_Report.Id;
          Elsif r_Report.�����ı� Like '%����%' Then
            Update ���Ӳ������� Set �����ı� = ���潨��_In || Chr(13) || Chr(13) Where ID = r_Report.Id;
          End If;
        End Loop;
        --���±���ʱ��
        Update ���Ӳ�����¼
        Set ���ʱ�� = Sysdate, ������ = ����ҽ��_In, ����ʱ�� = Sysdate
        Where ID = r_Advice.����id;
      Else
        --���жϵ������Ƿ��ж�Ӧ����ٺͱ��
        If Nvl(��������_In, ' ') <> ' ' Then
          Select Count(1)
          Into n_����
          From �����ļ��ṹ A, �����ļ��ṹ B
          Where a.��id = b.Id And a.�������� = 3 And b.�������� = 1 And a.�����ı� Like '%����%' And a.�ļ�id = r_Advice.�ļ�id;
        
          If n_���� <= 0 Then
            v_Error := '�����Ƶ�����û���ҵ�����������Ӧ����ٻ�������ϵ����Ա���ã�';
            Raise Err_Custom;
          End If;
        End If;
        If Nvl(�������_In, ' ') <> ' ' Then
          Select Count(1)
          Into n_����
          From �����ļ��ṹ A, �����ļ��ṹ B
          Where a.��id = b.Id And a.�������� = 3 And b.�������� = 1 And a.�����ı� Like '%���%' And a.�ļ�id = r_Advice.�ļ�id;
        
          If n_���� <= 0 Then
            v_Error := '�����Ƶ�����û���ҵ����������Ӧ����ٻ�������ϵ����Ա���ã�';
            Raise Err_Custom;
          End If;
        End If;
        If Nvl(���潨��_In, ' ') <> ' ' Then
          Select Count(1)
          Into n_����
          From �����ļ��ṹ A, �����ļ��ṹ B
          Where a.��id = b.Id And a.�������� = 3 And b.�������� = 1 And a.�����ı� Like '%����%' And a.�ļ�id = r_Advice.�ļ�id;
        
          If n_���� <= 0 Then
            v_Error := '�����Ƶ�����û���ҵ������顿��Ӧ����ٻ�������ϵ����Ա���ã�';
            Raise Err_Custom;
          End If;
        End If;
      
        If r_Advice.������Դ = 1 Then
          --�����ȡ�Һŵ�ID
          Select Nvl(c.Id, 0)
          Into v_�Һŵ�id
          From ����ҽ����¼ B, ���˹Һż�¼ C
          Where b.�Һŵ� = c.No(+) And c.��¼״̬ In (1, 3) And b.Id = v_��ҽ��id;
        Else
          --����������޹Һŵ�ID��ֱ������Ϊ0
          v_�Һŵ�id := 0;
        End If;
      
        --�������Ӳ�����¼
        Select ���Ӳ�����¼_Id.Nextval Into v_����id From Dual;
        n_Rptcount := Getrptno(ҽ��id_In, r_Advice.��������);
        If n_Rptcount > 1 Then
          v_�������� := r_Advice.�������� || n_Rptcount;
        Else
          v_�������� := r_Advice.��������;
        End If;
        Insert Into ���Ӳ�����¼
          (ID, ������Դ, ����id, ��ҳid, Ӥ��, ����id, ��������, �ļ�id, ��������, ������, ����ʱ��, ���ʱ��, ������, ����ʱ��, ���汾, ǩ������)
        Values
          (v_����id, r_Advice.������Դ, r_Advice.����id, Decode(r_Advice.������Դ, 2, r_Advice.��ҳid, v_�Һŵ�id), r_Advice.Ӥ��,
           r_Advice.���˿���id, r_Advice.��������, r_Advice.�ļ�id, v_��������, ����ҽ��_In, Sysdate, Sysdate, ����ҽ��_In, Sysdate, 1, 2);
      
        --����ҽ�������¼
        Insert Into ����ҽ������ (ҽ��id, ����id, Risid) Values (v_��ҽ��id, v_����id, Risid_In);
      
        v_������� := 0;
      
        --�²�����������
        For r_File In c_File(r_Advice.�ļ�id) Loop
          Select ���Ӳ�������_Id.Nextval Into v_��������id From Dual;
          v_�����ı�   := r_File.�����ı�;
          v_�������id := 0;
        
          If Nvl(r_File.��������, 0) = 1 And Nvl(r_File.��id, 0) = 0 Then
            --���
            v_�������id := r_File.Id;
            v_��id       := v_��������id;
          End If;
        
          If Nvl(r_File.��������, 0) = 4 And r_File.Ҫ������ Is Not Null Then
            --Ԫ��
            v_�����ı� := Zl_Replace_Element_Value(r_File.Ҫ������, r_Advice.����id, r_Advice.��ҳid, r_Advice.������Դ, r_Advice.Id);
          End If;
        
          If Nvl(r_File.��id, 0) <> 0 Then
            v_�������id := 0;
          End If;
        
          v_������� := v_������� + 1;
        
          If Instr(v_���, '|' || r_File.��id || '|') > 0 Then
            Null;
          Else
            Insert Into ���Ӳ�������
              (ID, �ļ�id, ��ʼ��, ��ֹ��, ��id, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���, Ԥ�����id, �������, ʹ��ʱ��, ����Ҫ��id, �滻��,
               Ҫ������, Ҫ������, Ҫ�س���, Ҫ��С��, Ҫ�ص�λ, Ҫ�ر�ʾ, ������̬, Ҫ��ֵ��, �������id)
            Values
              (v_��������id, v_����id, 1, 0, Decode(v_�������id, 0, v_��id, Null), v_�������, r_File.��������, r_File.������, r_File.��������,
               r_File.��������, Null, v_�����ı�, r_File.�Ƿ���, r_File.Ԥ�����id, r_File.�������, r_File.ʹ��ʱ��, r_File.����Ҫ��id,
               r_File.�滻��, r_File.Ҫ������, r_File.Ҫ������, r_File.Ҫ�س���, r_File.Ҫ��С��, r_File.Ҫ�ص�λ, r_File.Ҫ�ر�ʾ, r_File.������̬,
               r_File.Ҫ��ֵ��, Decode(v_�������id, 0, Null, v_�������id));
          End If;
        
          --Ϊ���ʱ�������ı�����
          If Nvl(r_File.��������, 0) = 3 And Nvl(r_File.��id, 0) <> 0 Then
            v_��� := v_��� || ',|' || r_File.Id || '|';
          
            If r_File.�����ı� Like '%����%' Then
              v_�����ı� := ��������_In || Chr(13) || Chr(13);
            Elsif r_File.�����ı� Like '%���%' Then
              v_�����ı� := �������_In || Chr(13) || Chr(13);
            Else
              v_�����ı� := ���潨��_In || Chr(13) || Chr(13);
            End If;
          
            For r_Con In c_Content(r_Advice.�ļ�id, r_File.Id) Loop
              Select ���Ӳ�������_Id.Nextval Into v_��������idnew From Dual;
              v_������� := v_������� + 1;
            
              Insert Into ���Ӳ�������
                (ID, �ļ�id, ��ʼ��, ��ֹ��, ��id, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���, Ԥ�����id, �������, ʹ��ʱ��, ����Ҫ��id,
                 �滻��, Ҫ������, Ҫ������, Ҫ�س���, Ҫ��С��, Ҫ�ص�λ, Ҫ�ر�ʾ, ������̬, Ҫ��ֵ��, �������id)
              Values
                (v_��������idnew, v_����id, 1, 0, v_��������id, v_�������, 2, r_Con.������, r_Con.��������, r_Con.��������, Null, v_�����ı�,
                 r_Con.�Ƿ���, r_Con.Ԥ�����id, r_Con.�������, r_Con.ʹ��ʱ��, r_Con.����Ҫ��id, r_Con.�滻��, r_Con.Ҫ������, r_Con.Ҫ������,
                 r_Con.Ҫ�س���, r_Con.Ҫ��С��, r_Con.Ҫ�ص�λ, r_Con.Ҫ�ر�ʾ, r_Con.������̬, r_Con.Ҫ��ֵ��,
                 Decode(v_�������id, 0, Null, v_�������id));
            End Loop;
          End If;
        End Loop;
      
        --����Ӳ�����ʽ�к����������ָ�ʽ�����ַ�������֮���������ֽ����ɼ�
        --Select ���� Into v_��ʽ���� From �����ļ���ʽ Where �ļ�ID=r_Advice.�ļ�ID;
        --Insert Into ���Ӳ�����ʽ (�ļ�ID,����) Values (v_����id,v_��ʽ����);
      
      End If;
    End If;
    Close c_Advice;
  
  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Receivereport;

  --5���޸����뵥��Ϣ
  Procedure Ӱ������Ϣ_�޸�
  (
    ҽ��id_In       ����ҽ����¼.Id%Type,
    ����_In         ������Ϣ.����%Type,
    �Ա�_In         ������Ϣ.�Ա�%Type,
    ����_In         ������Ϣ.����%Type,
    �ѱ�_In         ������Ϣ.�ѱ�%Type,
    ҽ�Ƹ��ʽ_In ������Ϣ.ҽ�Ƹ��ʽ%Type,
    ����_In         ������Ϣ.����%Type,
    ����_In         ������Ϣ.����״��%Type,
    ְҵ_In         ������Ϣ.ְҵ%Type,
    ���֤��_In     ������Ϣ.���֤��%Type,
    ��ͥ��ַ_In     ������Ϣ.��ͥ��ַ%Type,
    ��ͥ�绰_In     ������Ϣ.��ͥ�绰%Type,
    ��ͥ��ַ�ʱ�_In ������Ϣ.��ͥ��ַ�ʱ�%Type,
    ��������_In     ������Ϣ.��������%Type := Null
  ) As
  
    v_����     Varchar2(20);
    v_���䵥λ Varchar2(20);
    v_�������� Date;
    v_������Դ ����ҽ����¼.������Դ%Type;
    v_����id   ����ҽ����¼.����id%Type;
  Begin
    Begin
      Select ������Դ, ����id Into v_������Դ, v_����id From ����ҽ����¼ Where ID = ҽ��id_In;
    Exception
      When Others Then
        Return;
    End;
  
    If ��������_In Is Null And ����_In Is Not Null Then
      --�����������������
      v_���䵥λ := Substr(����_In, Length(����_In), 1);
      If Instr('��,��,��', v_���䵥λ) <= 0 Then
        v_���䵥λ := Null;
      Else
        v_���� := Replace(����_In, v_���䵥λ, '');
      End If;
      Begin
        v_���� := To_Number(v_����);
      Exception
        When Others Then
          v_���� := Null;
      End;
      If v_���� Is Not Null And v_���䵥λ Is Not Null Then
        Select Decode(v_���䵥λ, '��', Add_Months(Sysdate, -12 * v_����), '��', Add_Months(Sysdate, -1 * v_����), '��',
                       Sysdate - v_����)
        Into v_��������
        From Dual;
      End If;
    Else
      v_�������� := ��������_In;
    End If;
  
    If v_������Դ = 3 Then
      Update ������Ϣ
      Set ���� = ����_In, �Ա� = Nvl(�Ա�_In, �Ա�), ���� = ����_In, �������� = v_��������, �ѱ� = Nvl(�ѱ�_In, �ѱ�),
          ҽ�Ƹ��ʽ = Nvl(ҽ�Ƹ��ʽ_In, ҽ�Ƹ��ʽ), ���� = Nvl(����_In, ����), ����״�� = Nvl(����_In, ����״��), ְҵ = Nvl(ְҵ_In, ְҵ),
          ���֤�� = ���֤��_In, ��ͥ��ַ = ��ͥ��ַ_In, ��ͥ�绰 = ��ͥ�绰_In, ��ͥ��ַ�ʱ� = ��ͥ��ַ�ʱ�_In
      Where ����id = v_����id;
    
      --�޸Ķ�Ӧ��ҽ����¼
      Update ����ҽ����¼
      Set ���� = ����_In, �Ա� = �Ա�_In, ���� = ����_In
      Where ID = ҽ��id_In Or ���id = ҽ��id_In;
    Else
      Update ������Ϣ
      Set ���� = Nvl(����_In, ����), ����״�� = Nvl(����_In, ����״��), ְҵ = Nvl(ְҵ_In, ְҵ), ��ͥ��ַ = ��ͥ��ַ_In, ��ͥ�绰 = ��ͥ�绰_In,
          ��ͥ��ַ�ʱ� = ��ͥ��ַ�ʱ�_In
      Where ����id = v_����id;
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ӱ������Ϣ_�޸�;

  --6��ȡ�����뵥��Ϣ
  Procedure ȡ��������뵥
  (
    ҽ��id_In     ����ҽ��ִ��.ҽ��id%Type,
    ����Ա���_In ��Ա��.���%Type := Null,
    ����Ա����_In ��Ա��.����%Type := Null,
    ִ�в���id_In ������ü�¼.ִ�в���id%Type := 0,
    �ܾ�ԭ��_In   ����ҽ������.ִ��˵��%Type := Null
  ) As
    --������ҽ��ID_IN=����ִ�е�ҽ��ID
  
    v_���ͺ� ����ҽ��ִ��.���ͺ�%Type;
  
  Begin
  
    Begin
      Select ���ͺ� Into v_���ͺ� From ����ҽ������ Where ҽ��id = ҽ��id_In;
    Exception
      When Others Then
        Return;
    End;
  
    Zl_����ҽ��ִ��_�ܾ�ִ��(ҽ��id_In, v_���ͺ�, ����Ա���_In, ����Ա����_In, ִ�в���id_In, �ܾ�ԭ��_In);
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End ȡ��������뵥;

  --7������ҽ������ʧ�ܼ�¼
  Procedure Risҽ��ʧ�ܼ�¼_Insert
  (
    ������Դ_In   In Risҽ��ʧ�ܼ�¼.������Դ%Type,
    ����id_In     In Risҽ��ʧ�ܼ�¼.����id%Type,
    ��ҳid_In     In Risҽ��ʧ�ܼ�¼.��ҳid%Type,
    �Һŵ���_In   In Risҽ��ʧ�ܼ�¼.�Һŵ���%Type,
    ���ͺ�_In     In Risҽ��ʧ�ܼ�¼.���ͺ�%Type,
    �������id_In In Risҽ��ʧ�ܼ�¼.�������id%Type,
    ��챨����_In In Risҽ��ʧ�ܼ�¼.��챨����%Type,
    ��������_In   In Risҽ��ʧ�ܼ�¼.��������%Type
  ) Is
  Begin
    Insert Into Risҽ��ʧ�ܼ�¼
      (ID, ������Դ, ����id, ��ҳid, �Һŵ���, ���ͺ�, �������id, ��챨����, ��������, ����ʱ��, �ط�����)
    Values
      (Risҽ��ʧ�ܼ�¼_Id.Nextval, ������Դ_In, ����id_In, ��ҳid_In, �Һŵ���_In, ���ͺ�_In, �������id_In, ��챨����_In, ��������_In, Sysdate, 0);
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Risҽ��ʧ�ܼ�¼_Insert;

  --8������ҽ������ʧ�ܼ�¼
  Procedure Risҽ��ʧ�ܼ�¼_�ط�
  (
    Id_In       In Risҽ��ʧ�ܼ�¼.Id%Type,
    ��������_In In Number
  ) Is
    v_�ط����� Risҽ��ʧ�ܼ�¼.�ط�����%Type;
  Begin
    --��������_In -- 1 �ط��ɹ���ɾ����¼��2--�ط�ʧ��
  
    If ��������_In = 1 Then
      Delete From Risҽ��ʧ�ܼ�¼ Where ID = Id_In;
    Else
      Select �ط����� Into v_�ط����� From Risҽ��ʧ�ܼ�¼ Where ID = Id_In;
      If v_�ط����� >= 99 Then
        v_�ط����� := 99;
      Else
        v_�ط����� := v_�ط����� + 1;
      End If;
      Update Risҽ��ʧ�ܼ�¼ Set ����ʱ�� = Sysdate, �ط����� = v_�ط����� Where ID = Id_In;
    End If;
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Risҽ��ʧ�ܼ�¼_�ط�;

  --9�����˺��½�סԺ���˵���
  Procedure ����ҽ��_�ؽ�����
  (
    ҽ��id_In In ����ҽ������.ҽ��id%Type,
    No_In     In ����ҽ������.No%Type,
    Action_In In Number
  ) Is
    -- Action_In: 1 �ؽ����ݣ�2 ȡ���ؽ�����
    v_No ����ҽ������.No%Type;
  Begin
    If Action_In = 1 Then
      Select Nextno(14) Into v_No From Dual;
    
      Update ����ҽ������
      Set NO = v_No, �Ʒ�״̬ = 0
      Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = ҽ��id_In Or ���id = ҽ��id_In);
      Update סԺ���ü�¼ Set ҽ����� = Null Where NO = No_In;
    Elsif Action_In = 2 Then
      Update סԺ���ü�¼ Set ҽ����� = ҽ��id_In Where NO = No_In;
      Update ����ҽ������
      Set NO = No_In, �Ʒ�״̬ = 4
      Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = ҽ��id_In Or ���id = ҽ��id_In);
    End If;
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End ����ҽ��_�ؽ�����;

  --10����ӡRIS���ԤԼ֪ͨ��
  Procedure Ris���ԤԼ_��ӡ(ҽ��id_In In Ris���ԤԼ.ҽ��id%Type) Is
    v_Temp     Varchar2(255);
    v_��Ա���� ��Ա��.����%Type;
  Begin
    --ȡ��ǰ������Ա
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  
    Update Ris���ԤԼ Set �Ƿ��ӡ = 1, ��ӡ�� = v_��Ա����, ��ӡʱ�� = Sysdate Where ҽ��id = ҽ��id_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris���ԤԼ_��ӡ;

  --11������RIS�ֿ������ò���
  Procedure Ris���ÿ���_Update
  (
    �������_In Ris���ÿ���.�������%Type,
    ����_In     Ris���ÿ���.����%Type,
    ����ids_In  Varchar2,
    ��������_In Number
  ) Is
  
    l_����id   t_Numlist := t_Numlist();
    v_����ris  Ris���ÿ���.�Ƿ�����ris%Type;
    v_����ԤԼ Ris���ÿ���.�Ƿ�����ԤԼ%Type;
  
    Cursor c_Dept(Dept_In Varchar2) Is
      Select Column_Value From Table(f_Num2list(Dept_In));
  Begin
  
    If ��������_In = 1 Then
      v_����ris  := 1;
      v_����ԤԼ := Null;
      Delete From Ris���ÿ��� Where ������� = �������_In And ���� = ����_In And �Ƿ�����ris = 1;
    Else
      v_����ris  := Null;
      v_����ԤԼ := 1;
      Delete From Ris���ÿ��� Where ������� = �������_In And ���� = ����_In And �Ƿ�����ԤԼ = 1;
    End If;
  
    If ����ids_In Is Null Then
      Insert Into Ris���ÿ���
        (ID, �������, ����, ����id, �Ƿ�����ris, �Ƿ�����ԤԼ)
      Values
        (Ris���ÿ���_Id.Nextval, �������_In, ����_In, Null, v_����ris, v_����ԤԼ);
    Else
      Open c_Dept(����ids_In);
      Fetch c_Dept Bulk Collect
        Into l_����id;
      Close c_Dept;
    
      Forall I In 1 .. l_����id.Count
        Insert Into Ris���ÿ���
          (ID, �������, ����, ����id, �Ƿ�����ris, �Ƿ�����ԤԼ)
        Values
          (Ris���ÿ���_Id.Nextval, �������_In, ����_In, l_����id(I), v_����ris, v_����ԤԼ);
    End If;
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris���ÿ���_Update;

  --12��ɾ��RIS�ֿ������ò���
  Procedure Ris���ÿ���_Delete Is
  
  Begin
    Delete From Ris���ÿ���;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris���ÿ���_Delete;

  --13������Ԫ������ȡ��Ϣ
  Function Ris_Replace_Element_Value
  (
    Ԫ����_In   In ����������Ŀ.������%Type,
    ����id_In   In ���Ӳ�����¼.����id%Type,
    ����id_In   In ���Ӳ�����¼.��ҳid%Type,
    ������Դ_In In ���Ӳ�����¼.������Դ%Type,
    ҽ��id_In   In ����ҽ������.ҽ��id%Type
  ) Return Varchar2 Is
    v_Return Varchar2(4000) := Null;
    Cursor c_Patient Is
      Select ����, �Ա�, Decode(�Ա�, '��', 'M', 'Ů', 'F', 'O') As �Ա����, ��������, ����id, ��ϵ�˵�ַ, ��ͥ�绰, ��ϵ�˵绰, ����״��, ���֤��, ��ǰ����id,
             ��ǰ����id, ��ǰ���� As ����, ���￨��, ��Ժʱ��, ��Ժʱ��
      From ������Ϣ
      Where ����id = ����id_In;
    r_Patient c_Patient%RowType;
  
    Cursor c_Order Is
      Select ��ҳid, Ӥ��, Decode(������Դ, 1, 'OUTPAT', 2, 'INPAT', 'UNK') As ������Դ, ����ҽ��, ����ʱ��, У�Ի�ʿ, ҽ������, ������־, ִ�п���id
      From ����ҽ����¼
      Where ID = ҽ��id_In;
    r_Order c_Order%RowType;
  
    Cursor c_Diagnose Is
      Select ������� || Decode(Nvl(�Ƿ�����, 0), 0, '', ' (��)') As �ٴ����
      From �������ҽ�� A, ������ϼ�¼ B
      Where a.ҽ��id = ҽ��id_In And a.���id = b.Id;
    r_Diagnose c_Diagnose%RowType;
  
    --��ȡָ�����������
    Procedure p_Get_Rowtype(Table_In In Varchar2) Is
    Begin
      If Table_In = '������Ϣ' Then
        Open c_Patient;
        Fetch c_Patient
          Into r_Patient;
      Elsif Table_In = '����ҽ����¼' Then
        Open c_Order;
        Fetch c_Order
          Into r_Order;
      Elsif Table_In = '������ϼ�¼' Then
        Open c_Diagnose;
        Fetch c_Diagnose
          Into r_Diagnose;
      End If;
    Exception
      When Others Then
        Null;
    End p_Get_Rowtype;
  
  Begin
    Case
    --ֱ�ӷ��ص�����Ԫ��
      When Ԫ����_In = 'ҽ��ID' Then
        v_Return := ҽ��id_In;
      When Ԫ����_In = '����ID' Then
        v_Return := ����id_In;
      
    --�������Ա𵥶�����������Ӥ��
      When Instr(',����,�Ա�,�Ա����,��������,', ',' || Ԫ����_In || ',') > 0 Then
        p_Get_Rowtype('����ҽ����¼');
        p_Get_Rowtype('������Ϣ');
        If Nvl(r_Order.Ӥ��, 0) = 0 Then
          If Ԫ����_In = '����' Then
            v_Return := r_Patient.����;
          Elsif Ԫ����_In = '�Ա�' Then
            v_Return := r_Patient.�Ա�;
          Elsif Ԫ����_In = '�Ա����' Then
            v_Return := r_Patient.�Ա����;
          Elsif Ԫ����_In = '��������' Then
            v_Return := To_Char(r_Patient.��������, 'YYYYMMDDMISS');
          End If;
        Else
          If Ԫ����_In = '����' Then
            Select Decode(Ӥ������, Null, r_Patient.���� || '֮Ӥ' || Trim(To_Char(���, '9')), Ӥ������) As Ӥ������
            Into v_Return
            From ������������¼
            Where ����id = ����id_In And ��ҳid = r_Order.��ҳid And ��� = Nvl(r_Order.Ӥ��, 0);
          Elsif Instr('�Ա�', Ԫ����_In) > 0 Then
            Select Ӥ���Ա�
            Into v_Return
            From ������������¼
            Where ����id = ����id_In And ��ҳid = r_Order.��ҳid And ��� = Nvl(r_Order.Ӥ��, 0);
            If Ԫ����_In = '�Ա����' Then
              Select Decode(v_Return, '��', 'M', 'Ů', 'F', 'O') Into v_Return From Dual;
            End If;
          Elsif Ԫ����_In = '��������' Then
            Select ����ʱ��
            Into v_Return
            From ������������¼
            Where ����id = ����id_In And ��ҳid = r_Order.��ҳid And ��� = Nvl(r_Order.Ӥ��, 0);
            v_Return := To_Char(v_Return, 'YYYYMMDDMISS');
          End If;
        End If;
      
    --��ѯ������Ϣ���ص�Ԫ��
      When Instr(',��ϵ�˵�ַ,��ͥ�绰,��ϵ�˵绰,����״��,���֤��,����,���￨��,��Ժʱ��,��Ժʱ��,', ',' || Ԫ����_In || ',') > 0 Then
        p_Get_Rowtype('������Ϣ');
        Case Ԫ����_In
          When '��ϵ�˵�ַ' Then
            v_Return := r_Patient.��ϵ�˵�ַ;
          When '��ͥ�绰' Then
            v_Return := r_Patient.��ͥ�绰;
          When '��ϵ�˵绰' Then
            v_Return := r_Patient.��ϵ�˵绰;
          When '����״��' Then
            v_Return := r_Patient.����״��;
          When '���֤��' Then
            v_Return := r_Patient.���֤��;
          When '����' Then
            v_Return := r_Patient.����;
          When '���￨��' Then
            v_Return := r_Patient.���￨��;
          When '��Ժʱ��' Then
            v_Return := To_Char(r_Patient.��Ժʱ��, 'YYYYMMDDMISS');
          When '��Ժʱ��' Then
            v_Return := To_Char(r_Patient.��Ժʱ��, 'YYYYMMDDMISS');
          Else
            v_Return := '';
        End Case;
        --��ѯҽ�����ص�Ԫ��
      When Instr(',������Դ,����ҽ��,����ʱ��,У�Ի�ʿ,ҽ������,������־,������־����,', ',' || Ԫ����_In || ',') > 0 Then
        p_Get_Rowtype('����ҽ����¼');
        Case Ԫ����_In
          When '������Դ' Then
            v_Return := r_Order.������Դ;
          When '����ҽ��' Then
            v_Return := r_Order.����ҽ��;
          When '����ʱ��' Then
            v_Return := To_Char(r_Order.����ʱ��, 'YYYYMMDDMISS');
          When 'У�Ի�ʿ' Then
            v_Return := r_Order.У�Ի�ʿ;
          When 'ҽ������' Then
            v_Return := r_Order.ҽ������;
          When '������־' Then
            v_Return := r_Order.������־;
        End Case;
        --��ѯ��ϼ�¼���ص�Ԫ��
      When Ԫ����_In = '�ٴ����' Then
        p_Get_Rowtype('������ϼ�¼');
        v_Return := r_Diagnose.�ٴ����;
      
      Else
        --���в�ѯSQL����ֵ��Ԫ��
        If Ԫ����_In = 'ִ��վ��' Then
          p_Get_Rowtype('����ҽ����¼');
          Select Decode(վ��, 1, 'SITE0002', 2, 'SITE0001', 3, 'SITE0003', 'SITE0001')
          Into v_Return
          From ���ű�
          Where ID = r_Order.ִ�п���id;
        End If;
        If Ԫ����_In = '��ǰ��������' Then
          p_Get_Rowtype('������Ϣ');
          Select ���� Into v_Return From ���ű� Where ID = r_Patient.��ǰ����id;
        End If;
        If Ԫ����_In = '��������' Then
          p_Get_Rowtype('������Ϣ');
          Select ���� Into v_Return From ���ű� Where ID = r_Patient.��ǰ����id;
        End If;
        If Ԫ����_In = '��ʶ��' Then
          Select Decode(a.������Դ, 1, c.�����, 2, Decode(c.סԺ��, Null, c.�����, c.סԺ��), 4, c.������, c.�����)
          Into v_Return
          From ����ҽ����¼ A, ������Ϣ C
          Where a.����id = c.����id And a.Id = ҽ��id_In;
        End If;
    End Case;
  
    Return Trim(v_Return);
  Exception
    When Others Then
      Return Null;
  End Ris_Replace_Element_Value;

  --14��ɾ��RIS��Ժ���ò���
  Procedure Ris��Ժ����_Delete Is
  Begin
    Delete From Ris��Ժ����;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris��Ժ����_Delete;

  --15������RISRis��Ժ���ò���
  Procedure Ris��Ժ����_Update
  (
    Id_In           Ris��Ժ����.Id%Type,
    ҽԺ����_In     Ris��Ժ����.ҽԺ����%Type,
    ҽԺ����_In     Ris��Ժ����.ҽԺ����%Type,
    �û���_In       Ris��Ժ����.�û���%Type,
    ����_In         Ris��Ժ����.����%Type,
    ���ݿ������_In Ris��Ժ����.���ݿ������%Type
  ) Is
  
  Begin
  
    Insert Into Ris��Ժ����
      (ID, ҽԺ����, ҽԺ����, �û���, ����, ���ݿ������)
    Values
      (Id_In, ҽԺ����_In, ҽԺ����_In, �û���_In, ����_In, ���ݿ������_In);
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ris��Ժ����_Update;

End b_Zlxwinterface;
/
--128157:����,2018-07-04,������Ŀ��λɾ�����ݲ���ȷ
--126017:����,2018-05-23,������Ŀ�����鲿λ������ѡ����
Create Or Replace Procedure Zl_���Ƽ�鲿λ_Edit
(
  ����_In     In Number, --1:����;2:�޸�;3:ɾ��
  ����_In     In ���Ƽ�鲿λ.����%Type,
  ԭ����_In   In ���Ƽ�鲿λ.����%Type,
  �±���_In   In ���Ƽ�鲿λ.����%Type := Null,
  ����_In     In ���Ƽ�鲿λ.����%Type := Null,
  ����_In     In ���Ƽ�鲿λ.����%Type := Null,
  ��ע_In     In ���Ƽ�鲿λ.��ע%Type := Null,
  ����_In     In ���Ƽ�鲿λ.����%Type := Null,
  �����Ա�_In In ���Ƽ�鲿λ.�����Ա�%Type := Null,
  �ϼ�����_In In ���Ƽ�鲿λ.����%Type := Null --��ʽ���ϼ�����|����;�ϼ�����|����...(���ϼ�����Ϊ�գ���Ϊ|����)    
) Is
  v_ԭ���� ���Ƽ�鲿λ.����%Type := Null;
  e_Notfind Exception;
  v_����   Varchar2(1000);
  v_Fields Varchar2(1000);
  v_Tmp    Varchar2(1000);
  n_Count  Number;
  n_��¼id ������Ŀ��λ.Id%Type;
Begin
  If ����_In = 1 Then
    Insert Into ���Ƽ�鲿λ
      (����, ����, ����, ����, ��ע, ����, �����Ա�)
    Values
      (����_In, �±���_In, ����_In, ����_In, ��ע_In, ����_In, �����Ա�_In);
  Elsif ����_In = 2 Then
    Begin
      Select ���� Into v_ԭ���� From ���Ƽ�鲿λ Where ���� = ԭ����_In And ���� = ����_In;
    Exception
      When Others Then
        Null;
    End;
    If v_ԭ���� Is Null Then
      Raise e_Notfind;
    End If;
    Update ���Ƽ�鲿λ
    Set ���� = �±���_In, ���� = ����_In, ���� = ����_In, ��ע = ��ע_In, ���� = ����_In, �����Ա� = �����Ա�_In
    Where ���� = ԭ����_In And ���� = ����_In;
  
    --�����޸�
    v_���� := ';' || ����_In;
    v_���� := Replace(v_����, ',', Chr(10));
    v_���� := Replace(v_����, Chr(9), ';');
    v_���� := Replace(v_����, ';0', Chr(10));
    v_���� := Replace(v_����, ';1', Chr(10));
    v_���� := Replace(v_����, Chr(10), ';');
    v_���� := Replace(v_����, ';;', ';');
    v_���� := v_���� || ';';
  
    v_���� := Substr(v_����, 2);
    --ԭ�еķ����������Ѿ�ɾ���˻�ԭ�еĲ�λ�������Ѿ��ı���
    For r_Used In (Select ID, ��Ŀid, ��λ, ����, ����, Ĭ��,�ϼ����� From ������Ŀ��λ Where ��λ = v_ԭ���� And ���� = ����_In) Loop
      If Instr(';' || v_����, ';' || r_Used.���� || ';') = 0 Then
        Delete ������Ŀ��λ
        Where ID=r_Used.id;
      Else
        Update ������Ŀ��λ
        Set ��λ = ����_In
        Where ID=r_Used.id;
      End If;
    End Loop;
  
    --ԭ��û�еķ�����������
    v_Tmp := v_����;
    While v_Tmp Is Not Null Loop
      --����ȡÿ����Ŀ
      v_Fields := Substr(v_Tmp, 1, Instr(v_Tmp, ';') - 1);
      v_Tmp    := Substr(v_Tmp, Instr(v_Tmp, ';') + 1);
    
      If v_Fields Is Not Null Then
        For r_Used In (Select Distinct ��Ŀid From ������Ŀ��λ Where ��λ = ����_In And ���� = ����_In) Loop
          Select Count(ID)
          Into n_Count
          From ������Ŀ��λ
          Where ��Ŀid = r_Used.��Ŀid And ��λ = ����_In And ���� = ����_In And ���� = v_Fields;
        
          If n_Count = 0 Then
            Select ������Ŀ��λ_Id.Nextval Into n_��¼id From Dual;
            Insert Into ������Ŀ��λ
              (ID, ��Ŀid, ����, ��λ, ����)
            Values
              (n_��¼id, r_Used.��Ŀid, ����_In, ����_In, v_Fields);
          End If;
        End Loop;
      End If;
    End Loop;
  Elsif ����_In = 3 Then
    Delete ���Ƽ�鲿λ Where ���� = ԭ����_In And ���� = ����_In;
  End If;

Exception
  When e_Notfind Then
    Raise_Application_Error(-20101, '[ZLSOFT]�ò�λ�����ڣ������ѱ������û�ɾ���޸ģ�[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���Ƽ�鲿λ_Edit;
/
--126017:����,2018-05-23,������Ŀ�����鲿λ������ѡ����
CREATE OR REPLACE Procedure Zl_������Ŀ��λ_Insert
(
  ��Ŀid_In In ������Ŀ��λ.��Ŀid%Type,
  ����_In   In ������Ŀ��λ.����%Type,
  ��λ_In   In ������Ŀ��λ.��λ%Type,
  ����_In   In ������Ŀ��λ.����%Type,
  Ĭ��_In   In ������Ŀ��λ.Ĭ��%Type := Null,
  �ϼ�����_In  In ������Ŀ��λ.�ϼ�����%Type := Null
) As
  v_Code Varchar2(20); --����
  Err_Notfind Exception;
Begin
  Select Rtrim(����) Into v_Code From ������ĿĿ¼ Where ��� = 'D' And Id = ��Ŀid_In;
  If v_Code Is Null Then
    Raise Err_Notfind;
  End If;
  Insert Into ������Ŀ��λ (ID, ��Ŀid, ����, ��λ, ����, Ĭ��,�ϼ�����) Values (������Ŀ��λ_ID.Nextval, ��Ŀid_In, ����_In, ��λ_In, ����_In, Ĭ��_In,�ϼ�����_In);
Exception
  When Err_Notfind Then
    Raise_Application_Error(-20101, '[ZLSOFT]����Ŀ�����ڣ������ѱ������û�ɾ����[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_������Ŀ��λ_Insert;
/
--119268:��С��,2018-05-11,ȥ��������ɾ����ش���
Create Or Replace Procedure Zl_����ͼ����_Update
(
  �걾id_In   In ����ͼ����.�걾id%Type,
  ͼ������_In In ����ͼ����.ͼ������%Type,
  ͼ���_In   In Varchar2, -- ͼ������^ͼ������;���1;���2;���3....   ͼ������ 1=ֱ��ͼ 2=ɢ��ͼ 
  ͼ��lob_In  In Number, -- 0=С��4000  1=����4000 ��Ҫ���⴦�� 
  ��ʼ_In     In Number, -- 1=��ʼ 
  ͼ��λ��_In In ����ͼ����.ͼ��λ��%Type := Null
) Is
  l_Clob Clob;
Begin
  -- ���浽FTP 
  If ͼ��λ��_In Is Not Null Then
    Update ����ͼ���� Set ͼ��λ�� = ͼ��λ��_In Where �걾id = �걾id_In And ͼ������ = ͼ������_In;
    If Sql%RowCount = 0 Then
      Insert Into ����ͼ����
        (ID, �걾id, ͼ������, ͼ��λ��)
      Values
        (����ͼ����_Id.Nextval, �걾id_In, ͼ������_In, ͼ��λ��_In);
    End If;
    Return;
  End If;

  -- ���浽���ݿ� 
  If ͼ���_In Is Null Then
    Return;
  End If;

  If ͼ��lob_In = 0 Then
    Update ����ͼ���� Set ͼ��� = ͼ���_In Where �걾id = �걾id_In And ͼ������ = ͼ������_In;
    If Sql%RowCount = 0 Then
      Insert Into ����ͼ����
        (ID, �걾id, ͼ������, ͼ���)
      Values
        (����ͼ����_Id.Nextval, �걾id_In, ͼ������_In, ͼ���_In);
    End If;
  Else
    If ��ʼ_In = 1 Then
      Update ����ͼ���� Set ͼ��� = Empty_Clob() Where �걾id = �걾id_In And ͼ������ = ͼ������_In;
      If Sql%RowCount = 0 Then
        Insert Into ����ͼ����
          (ID, �걾id, ͼ������, ͼ���)
        Values
          (����ͼ����_Id.Nextval, �걾id_In, ͼ������_In, Empty_Clob());
      End If;
    End If;
    Select ͼ��� Into l_Clob From ����ͼ���� Where �걾id = �걾id_In And ͼ������ = ͼ������_In For Update;
    Dbms_Lob.Writeappend(l_Clob, Length(ͼ���_In), ͼ���_In);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ͼ����_Update;
/

--119268:��С��,2018-05-11,ȥ��������ɾ����ش���
Create Or Replace Procedure Zl_���Ӳ�����ʽ_Insert
(
  Id_In   In ���Ӳ�����ʽ.�ļ�id%Type,
  Txt_In  In Varchar2,
  ��ʼ_In In Number -- 1=��ʼ 
) Is
  l_Blob Blob;
Begin
  If ��ʼ_In = 1 Then
    Update ���Ӳ�����ʽ Set ���� = Empty_Blob() Where �ļ�id = Id_In;
    If Sql%RowCount = 0 Then
      Insert Into ���Ӳ�����ʽ (�ļ�id, ����) Values (Id_In, Empty_Blob());
    End If;
  End If;
  Select ���� Into l_Blob From ���Ӳ�����ʽ Where �ļ�id = Id_In For Update;
  Dbms_Lob.Writeappend(l_Blob, Length(Hextoraw(Txt_In)) / 2, Hextoraw(Txt_In));
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���Ӳ�����ʽ_Insert;
/

--125261:������,2018-05-08,ת��ҽ��У�Է��ʹ����Զ�ͣ����
Create Or Replace Procedure Zl_����ҽ����¼_У��
(
  --���ܣ�У��ָ����ҽ��
  --������ҽ��ID_IN=Nvl(���ID,ID)
  --      ״̬_IN=У��ͨ��3��У������2
  --      �Զ�У��_IN=����֮������Զ�У��,�Զ���д�Ƽ�����
  --˵����һ��ҽ��ֻ�ܵ���һ��,����ͬʱ��ɴ���һ��ҽ����У��
  ҽ��id_In     ����ҽ����¼.Id%Type,
  ״̬_In       ����ҽ����¼.ҽ��״̬%Type,
  У��ʱ��_In   ����ҽ��״̬.����ʱ��%Type,
  У��˵��_In   ����ҽ��״̬.����˵��%Type := Null,
  �Զ�У��_In   Number := Null,
  ����Ա���_In ��Ա��.���%Type := Null,
  ����Ա����_In ��Ա��.����%Type := Null
) Is
  --����ҽ�����
  v_״̬       ����ҽ����¼.ҽ��״̬%Type;
  v_��Ч       ����ҽ����¼.ҽ����Ч%Type;
  v_����id     ����ҽ����¼.����id%Type;
  v_��ҳid     ����ҽ����¼.��ҳid%Type;
  v_Ӥ��       ����ҽ����¼.Ӥ��%Type;
  v_ҽ������   ����ҽ����¼.ҽ������%Type;
  v_����ʱ��   ����ҽ����¼.����ʱ��%Type;
  v_��ʼʱ��   ����ҽ����¼.��ʼִ��ʱ��%Type;
  v_����ҽ��   ����ҽ����¼.����ҽ��%Type;
  v_ǰ��id     ����ҽ����¼.ǰ��id%Type;
  v_ִ�б��   ����ҽ����¼.ִ�б��%Type;
  v_ִ�п���id ����ҽ����¼.ִ�п���id%Type;
  v_�걾��λ   ����ҽ����¼.�걾��λ%Type;
  v_ֹͣʱ��   ����ҽ����¼.����ʱ��%Type;
  v_��������id ����ҽ����¼.��������id%Type;

  --���ڱ������ȼ�
  v_�������   ����ҽ����¼.�������%Type;
  v_������Ŀid ����ҽ����¼.������Ŀid%Type;
  v_��������   ������ĿĿ¼.��������%Type;
  v_����ȼ�id ������ҳ.����ȼ�id%Type;
  v_������־   ����ҽ����¼.������־%Type;
  v_��Ժ��ʽ   ��Ժ��ʽ.����%Type;

  v_Stopadviceids ����ҽ����¼.ҽ������%Type;
  n_Adviceid      ����ҽ����¼.����id%Type;
  n_���          Number(18);
  --�����Ŀͬһ�Զ�ֹͣ���������Ŀ:����Ӧ�ö��ǳ���(������ǰҽ��),����Ӧ�Ѽ�顣
  --ע��Ӧ��Ӥ������,ͬʱҲӦֹͣ����ǰҽ�����������ͬ������Ŀ��ҽ����
  Cursor c_Exclude Is
    Select Distinct b.Id As ҽ��id, b.��ʼִ��ʱ��, b.ִ����ֹʱ��, b.�ϴ�ִ��ʱ��, b.����ҽ��, b.ִ��ʱ�䷽��, b.Ƶ�ʼ��, b.Ƶ�ʴ���, b.�����λ
    From ���ƻ�����Ŀ A, ����ҽ����¼ B
    Where a.���� = 3 And a.��Ŀid = b.������Ŀid And b.Id <> ҽ��id_In And Nvl(b.ҽ����Ч, 0) = 0 And b.ҽ��״̬ In (3, 5, 6, 7) And
          b.����id = v_����id And Nvl(b.��ҳid, 0) = Nvl(v_��ҳid, 0) And Nvl(b.Ӥ��, 0) = Nvl(v_Ӥ��, 0) And
          a.���� In (Select Distinct ���� From ���ƻ�����Ŀ Where ���� = 3 And ��Ŀid = v_������Ŀid)
    Order By b.Id;
  v_��ֹʱ�� ����ҽ����¼.ִ����ֹʱ��%Type;

  --����ȼ�����
  Cursor c_Nurse Is
    Select a.Id As ҽ��id, a.��ʼִ��ʱ��, a.ִ����ֹʱ��, a.�ϴ�ִ��ʱ��, a.����ҽ��
    From ����ҽ����¼ A, ������ĿĿ¼ B
    Where a.������Ŀid = b.Id And a.������� = 'H' And b.�������� = '1' And a.����id = v_����id And Nvl(a.��ҳid, 0) = Nvl(v_��ҳid, 0) And
          Nvl(a.Ӥ��, 0) = Nvl(v_Ӥ��, 0) And Nvl(a.ҽ����Ч, 0) = 0 And a.ҽ��״̬ In (3, 5, 6, 7) And a.Id <> ҽ��id_In;

  --��¼���������
  Cursor c_Patiio Is
    Select a.Id As ҽ��id, a.��ʼִ��ʱ��, a.ִ����ֹʱ��, a.�ϴ�ִ��ʱ��, a.����ҽ��
    From ����ҽ����¼ A, ������ĿĿ¼ B
    Where a.������Ŀid = b.Id And a.������� = 'Z' And b.�������� = '12' And a.����id = v_����id And Nvl(a.��ҳid, 0) = Nvl(v_��ҳid, 0) And
          Nvl(a.Ӥ��, 0) = Nvl(v_Ӥ��, 0) And Nvl(a.ҽ����Ч, 0) = 0 And a.ҽ��״̬ In (3, 5, 6, 7) And a.Id <> ҽ��id_In;

  --��¼���黥��
  Cursor c_Patistate Is
    Select a.Id As ҽ��id, a.��ʼִ��ʱ��, a.ִ����ֹʱ��, a.�ϴ�ִ��ʱ��, a.����ҽ��
    From ����ҽ����¼ A, ������ĿĿ¼ B
    Where a.������Ŀid = b.Id And a.������� = 'Z' And b.�������� In ('9', '10') And a.����id = v_����id And
          Nvl(a.��ҳid, 0) = Nvl(v_��ҳid, 0) And Nvl(a.Ӥ��, 0) = Nvl(v_Ӥ��, 0) And Nvl(a.ҽ����Ч, 0) = 0 And
          a.ҽ��״̬ In (3, 5, 6, 7) And a.Id <> ҽ��id_In;
  --�䶯��Ч��¼
  Cursor c_Oldinfo Is
    Select b.*
    From (Select c.*
           From ���˱䶯��¼ C
           Where c.����id = v_����id And c.��ҳid = v_��ҳid And
                 c.��ʼʱ�� = (Select Min(��ʼʱ��)
                           From ���˱䶯��¼
                           Where ����id = v_����id And ��ҳid = v_��ҳid And ��ʼʱ�� > v_��ʼʱ��) And
                 NVL(c.��ֹʱ��|| '','��') = (Select  NVL(Min(��ֹʱ��)|| '','��')
                           From ���˱䶯��¼
                           Where ����id = v_����id And ��ҳid = v_��ҳid And ��ʼʱ�� > v_��ʼʱ��)) A, ���˱䶯��¼ B
    
    Where b.����id = v_����id And b.��ҳid = v_��ҳid And a.��ʼʱ�� = b.��ֹʱ�� And a.��ʼԭ�� = b.��ֹԭ�� And a.���Ӵ�λ = b.���Ӵ�λ
    Union
    Select *
    From ���˱䶯��¼
    Where ����id = v_����id And ��ҳid = v_��ҳid And ��ֹʱ�� Is Null And ��ʼʱ�� <= v_��ʼʱ��;

  Cursor c_Endinfo Is
    Select * From ���˱䶯��¼ Where ����id = v_����id And ��ҳid = v_��ҳid And ��ֹʱ�� Is Null;
  r_Oldinfo      c_Oldinfo%RowType;
  r_Endinfo      c_Endinfo%RowType;
  v_�䶯��ֹԭ�� ���˱䶯��¼.��ֹԭ��%Type;
  v_�䶯��ֹʱ�� ���˱䶯��¼.��ֹʱ��%Type;
  v_�䶯��ֹ��Ա ���˱䶯��¼.��ֹ��Ա%Type;

  --��������(Ӥ��)������δͣ����(���䷽����)
  Cursor c_Needstop
  (
    v_����id   ����ҽ����¼.����id%Type,
    v_��ҳid   ����ҽ����¼.��ҳid%Type,
    v_Ӥ��     ����ҽ����¼.Ӥ��%Type,
    v_Stoptime Date
  ) Is
    Select a.Id, a.�������, b.��������, b.ִ��Ƶ��
    From ����ҽ����¼ A, ������ĿĿ¼ B
    Where a.������Ŀid = b.Id(+) And a.����id = v_����id And a.��ҳid = v_��ҳid And (v_Ӥ�� = -1 Or Nvl(a.Ӥ��, 0) = Nvl(v_Ӥ��, 0)) And
          Nvl(a.ҽ����Ч, 0) = 0 And a.ҽ��״̬ Not In (1, 2, 4, 8, 9) And a.��ʼִ��ʱ�� <= v_Stoptime
    Order By a.���;
  --��������(Ӥ��)����ͣ��δȷ�ϵĳ���,��ִֹ��ʱ����ָ��ʱ��֮��
  Cursor c_Havestop
  (
    v_����id   ����ҽ����¼.����id%Type,
    v_��ҳid   ����ҽ����¼.��ҳid%Type,
    v_Ӥ��     ����ҽ����¼.Ӥ��%Type,
    v_Stoptime Date
  ) Is
    Select ID
    From ����ҽ����¼
    Where ����id = v_����id And ��ҳid = v_��ҳid And (v_Ӥ�� = -1 Or Nvl(Ӥ��, 0) = Nvl(v_Ӥ��, 0)) And Nvl(ҽ����Ч, 0) = 0 And
          ҽ��״̬ = 8 And ִ����ֹʱ�� > v_Stoptime And ��ʼִ��ʱ�� < v_Stoptime
    Order By ���;

  --ȡһ��ҽ���ļƼ�����
  Cursor c_Price Is
    Select a.Id, b.�շ���Ŀid, b.�շ�����, b.������Ŀ, b.��������, b.�շѷ�ʽ, c.��� As �շ����, a.�������, e.��������, e.�Թܱ���,
           Sum(Decode(Nvl(c.�Ƿ���, 0), 1, Nvl(d.ȱʡ�۸�, d.ԭ��), Null)) As ����
    From ����ҽ����¼ A, �����շѹ�ϵ B, �շ���ĿĿ¼ C, �շѼ�Ŀ D, ������ĿĿ¼ E
    Where a.������Ŀid = b.������Ŀid And b.�շ���Ŀid = c.Id And c.Id = d.�շ�ϸĿid And
          (a.���id Is Null And a.ִ�б�� In (1, 2) And b.�������� = 1 Or
          a.�걾��λ = b.��鲿λ And a.��鷽�� = b.��鷽�� And Nvl(b.��������, 0) = 0 Or
          a.��鷽�� Is Null And Nvl(b.��������, 0) = 0 And b.��鲿λ Is Null And b.��鷽�� Is Null) And
          a.������� Not In ('5', '6', '7') And Nvl(a.�Ƽ�����, 0) = 0 And Nvl(a.ִ������, 0) Not In (0, 5) And c.������� In (2, 3) And
          (c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.����ʱ�� Is Null) And Sysdate Between d.ִ������ And
          Nvl(d.��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD')) And Nvl(b.�շ�����, 0) <> 0 And
          Not (Nvl(c.�Ƿ���, 0) = 1 And Nvl(Nvl(d.ȱʡ�۸�, d.ԭ��), 0) = 0) And a.������Ŀid = e.Id And
          (a.Id = ҽ��id_In Or a.���id = ҽ��id_In)
    Group By a.Id, b.�շ���Ŀid, b.�շ�����, b.������Ŀ, b.��������, b.�շѷ�ʽ, c.���, a.�������, e.��������, e.�Թܱ���;

  Cursor c_Pati(v_����id ������Ϣ.����id%Type) Is
    Select * From ������Ϣ Where ����id = v_����id;
  r_Pati c_Pati%RowType;

  v_����id ��Ѫ������.����id%Type;

  --������ʱ����
  v_Count    Number;
  v_Date     Date;
  v_Temp     Varchar2(255);
  v_��Ա��� ��Ա��.���%Type;
  v_��Ա���� ��Ա��.����%Type;
  v_����ִ�� Varchar2(5);

  v_Error Varchar2(255);
  Err_Custom Exception;

  Function Getadvicetext(v_ҽ��id ����ҽ����¼.Id%Type) Return Varchar2 Is
    v_Text ����ҽ����¼.ҽ������%Type;
    v_��� ����ҽ����¼.�������%Type;
    v_�䷽ Number;
  Begin
    Select �������, ҽ������ Into v_���, v_Text From ����ҽ����¼ Where ID = v_ҽ��id;
    If v_��� = 'E' Then
      --��ҩ���г�ҩ��ҽ������
      Begin
        Select �������, Decode(�������, '7', v_Text, ҽ������)
        Into v_���, v_Text
        From ����ҽ����¼
        Where ���id = v_ҽ��id And ������� In ('5', '6', '7') And Rownum = 1;
      Exception
        When Others Then
          Null;
      End;
      If v_��� = '7' Then
        v_�䷽ := 1;
      End If;
    End If;
    If Length(v_Text) > 30 Then
      v_Text := Substr(v_Text, 1, 30) || '...';
    End If;
    If Length(v_Text) > 20 Then
      v_Text := '"' || v_Text || '"' || Chr(13) || Chr(10);
    Else
      v_Text := '"' || v_Text || '"';
    End If;
    If v_�䷽ = 1 Then
      v_Text := '��ҩ�䷽' || v_Text;
    End If;
    Return(v_Text);
  End;
Begin
  --���ҽ��״̬�Ƿ���ȷ:��������
  Begin
    Select a.ҽ����Ч, a.ҽ��״̬, a.����ʱ��, a.����ҽ��, a.��ʼִ��ʱ��, a.����id, a.��ҳid, a.Ӥ��, a.ҽ������, a.�������, a.������Ŀid, a.ǰ��id,
           Nvl(b.��������, '0'), Nvl(a.ִ�б��, 0), a.ִ�п���id, a.�걾��λ, a.��������id, Nvl(a.������־, 0) As ������־
    Into v_��Ч, v_״̬, v_����ʱ��, v_����ҽ��, v_��ʼʱ��, v_����id, v_��ҳid, v_Ӥ��, v_ҽ������, v_�������, v_������Ŀid, v_ǰ��id, v_��������, v_ִ�б��,
         v_ִ�п���id, v_�걾��λ, v_��������id, v_������־
    From ����ҽ����¼ A, ������ĿĿ¼ B
    Where a.������Ŀid = b.Id(+) And a.Id = ҽ��id_In;
  Exception
    When Others Then
      Begin
        v_Error := 'ҽ���ѱ�ɾ�������ܽ���У�ԡ�' || Chr(13) || Chr(10) || '������ǲ�����������ģ������¶�ȡУ�����ݡ�';
        Raise Err_Custom;
      End;
  End;
  If v_״̬ <> 1 Then
    v_Error := 'ҽ��"' || Getadvicetext(ҽ��id_In) || '"�����¿���ҽ��������ͨ��У�ԡ�' || Chr(13) || Chr(10) || '������ǲ�����������ģ������¶�ȡУ�����ݡ�';
    Raise Err_Custom;
  End If;
  --�ٴμ��У��ʱ�����Ч��:��������
  If To_Char(v_����ʱ��, 'YYYY-MM-DD HH24:MI') <= To_Char(v_��ʼʱ��, 'YYYY-MM-DD HH24:MI') Then
    If To_Char(У��ʱ��_In, 'YYYY-MM-DD HH24:MI') < To_Char(v_����ʱ��, 'YYYY-MM-DD HH24:MI') Then
      v_Error := 'ҽ��"' || Getadvicetext(ҽ��id_In) || '"��У��ʱ�䲻��С�ڿ���ʱ�� ' || To_Char(v_����ʱ��, 'YYYY-MM-DD HH24:MI') || '��' ||
                 Chr(13) || Chr(10) || '������ǲ�����������ģ������¶�ȡУ�����ݡ�';
      Raise Err_Custom;
    End If;
  Else
    If To_Char(У��ʱ��_In, 'YYYY-MM-DD HH24:MI') < To_Char(v_��ʼʱ��, 'YYYY-MM-DD HH24:MI') Then
      v_Error := 'ҽ��"' || Getadvicetext(ҽ��id_In) || '"��У��ʱ�䲻��С�ڿ�ʼִ��ʱ�� ' || To_Char(v_��ʼʱ��, 'YYYY-MM-DD HH24:MI') || '��' ||
                 Chr(13) || Chr(10) || '������ǲ�����������ģ������¶�ȡУ�����ݡ�';
      Raise Err_Custom;
    End If;
  End If;

  --���Ҫ��ǩ�������У��ʱ�Ƿ���ǩ��(����ȡ��ǩ��)
  If ״̬_In = 3 Then
    Select Zl_Fun_Getsignpar(Decode(v_ǰ��id, Null, 1, 3), v_��������id) Into v_Count From Dual;
    If v_Count = 1 Then
      --֤��ͣ�û�δע��֤�鲻����ǩ������ֻ�ж�һ�����ݼ���
      For C In (Select a.�Ƿ�ͣ��
                From ��Ա֤���¼ A, ��Ա�� B
                Where a.��Աid = b.Id And b.���� = v_����ҽ��
                Order By a.ע��ʱ�� Desc) Loop
        If Nvl(c.�Ƿ�ͣ��, 0) = 0 Then
          Select Count(*)
          Into v_Count
          From ����ҽ��״̬ A
          Where �������� = 1 And ҽ��id = ҽ��id_In And
                (ǩ��id Is Null And Exists
                 (Select 1
                  From ��Ա�� R, ��Ա����˵�� X
                  Where r.Id = x.��Աid And r.���� = a.������Ա And x.��Ա���� = '��ʿ') And Not Exists
                 (Select 1
                  From ��Ա�� R, ��Ա����˵�� Y
                  Where r.Id = y.��Աid And r.���� = a.������Ա And y.��Ա���� = 'ҽ��') Or ǩ��id Is Not Null Or a.������Ա <> v_����ҽ��);
          If Nvl(v_Count, 0) = 0 Then
            v_Error := 'ҽ��"' || Getadvicetext(ҽ��id_In) || '"��û�е���ǩ��������ͨ��У�ԡ�';
            Raise Err_Custom;
          End If;
        End If;
        Exit;
      End Loop;
    End If;
  End If;

  --��ǰ������Ա
  If ����Ա���_In Is Not Null And ����Ա����_In Is Not Null Then
    v_��Ա��� := ����Ա���_In;
    v_��Ա���� := ����Ա����_In;
  Else
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_��Ա��� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
    v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  End If;

  --��Ϊ����ͬʱ���¿�->�Զ�У��->�����Զ�ֹͣ,��˷ֱ�-2,-1��
  Select Sysdate - 1 / 60 / 60 / 24 Into v_Date From Dual;

  Update ����ҽ����¼
  Set ҽ��״̬ = ״̬_In, У�Ի�ʿ = v_��Ա����, У��ʱ�� = У��ʱ��_In
  Where ID = ҽ��id_In Or ���id = ҽ��id_In;

  Insert Into ����ҽ��״̬
    (ҽ��id, ��������, ������Ա, ����ʱ��, ����˵��)
    Select ID, ״̬_In, v_��Ա����, v_Date, У��˵��_In From ����ҽ����¼ Where ID = ҽ��id_In Or ���id = ҽ��id_In;

  --У��ͨ��ʱ����������
  If ״̬_In = 3 Then
    --�Զ�У��ʱ���Զ���дȱʡ�ļƼ�����
    If Nvl(�Զ�У��_In, 0) = 1 Then
      --1.��۵ļƼ���Ŀ,�������޼۲�Ϊ0,��ȱʡΪ����޼�,���򲻼���;�����ֹ��Ƽ�.
      --2.���ڷ�ҩ��ҩƷ����������δ��ִ�п���,����ʱ��ȡȱʡ��,�����ֹ����á�
      For r_Price In c_Price Loop
        --ȡ(����)ҽ���Ĺ���Ͳ���,�ɼ���ʽ�Լ�����Ŀ��Ϊ׼
        v_����id := Null;
        If r_Price.������� = 'E' And r_Price.�������� = '6' Then
          Begin
            Select c.����id
            Into v_����id
            From ����ҽ����¼ A, ������ĿĿ¼ B, ��Ѫ������ C
            Where a.������Ŀid = b.Id And b.�Թܱ��� = c.���� And a.���id = r_Price.Id And Rownum = 1;
          Exception
            When Others Then
              Null;
          End;
        Elsif r_Price.������� = 'C' And r_Price.�Թܱ��� Is Not Null Then
          Begin
            Select ����id Into v_����id From ��Ѫ������ Where ���� = r_Price.�Թܱ���;
          Exception
            When Others Then
              Null;
          End;
        End If;
      
        --�жϴ�������Թܷ��õ���ȡ
        If (Nvl(r_Price.�շѷ�ʽ, 0) = 1 And r_Price.�շ���� = '4' And r_Price.�շ���Ŀid = Nvl(v_����id, 0) Or
           Not (Nvl(r_Price.�շѷ�ʽ, 0) = 1 And r_Price.�շ���� = '4' And Nvl(v_����id, 0) <> 0)) Then
          Insert Into ����ҽ���Ƽ�
            (ҽ��id, �շ�ϸĿid, ����, ����, ����, ִ�п���id, ��������, �շѷ�ʽ)
          Values
            (r_Price.Id, r_Price.�շ���Ŀid, r_Price.�շ�����, r_Price.����, r_Price.������Ŀ, Null, r_Price.��������, r_Price.�շѷ�ʽ);
        End If;
      End Loop;
    End If;
  
    --����¼�������ҽ�����Ϊֹͣ
    If Nvl(v_��Ч, 0) = 1 And v_������Ŀid Is Null Then
      Update ����ҽ����¼
      Set ҽ��״̬ = 8, ͣ��ʱ�� = У��ʱ��_In, ͣ��ҽ�� = v_����ҽ��
      Where ID = ҽ��id_In Or ���id = ҽ��id_In;
    
      Insert Into ����ҽ��״̬
        (ҽ��id, ��������, ������Ա, ����ʱ��)
        Select ID, 8, v_��Ա����, Sysdate From ����ҽ����¼ Where ID = ҽ��id_In Or ���id = ҽ��id_In;
      End If;  
      
    --�ж��Ƿ���������Ҫִ��
    v_����ִ��:= zl_GetSysParameter(288);
    if v_����ִ��=1 and v_������Ŀid Is Null then
        Insert Into ����ҽ������
          (ҽ��id, ���ͺ�, ��¼����, NO, ��¼���, ��������, ������, ����ʱ��, ִ��״̬, ִ�в���id, �Ʒ�״̬, �״�ʱ��, ĩ��ʱ��)
        Values
          (ҽ��id_In, NextNO('10','0','','1'), '2',NextNO('14','0','','1'), '1', '1', v_��Ա����, sysdate, '0', v_ִ�п���id,'0',sysdate,sysdate);    
     End If;

    --��ͬһ�Զ�ֹͣ�������еĲ�������ҽ��ֹͣ(�����δֹͣ)
    For r_Exclude In c_Exclude Loop
      Select Decode(Sign(r_Exclude.��ʼִ��ʱ�� - v_��ʼʱ��), 1, r_Exclude.��ʼִ��ʱ��, v_��ʼʱ��)
      Into v_��ֹʱ��
      From Dual;
      Select Decode(Sign(r_Exclude.ִ����ֹʱ�� - v_��ʼʱ��), -1, r_Exclude.ִ����ֹʱ��, v_��ʼʱ��)
      Into v_��ֹʱ��
      From Dual;
      Zl_����ҽ����¼_ֹͣ(r_Exclude.ҽ��id, v_��ֹʱ��, v_����ҽ��, 1);
      v_Stopadviceids := v_Stopadviceids || ',' || r_Exclude.ҽ��id;
    End Loop;
  
    --��һЩ����ҽ���Ĵ���
    If v_������� = 'H' And v_�������� = '1' And Nvl(v_��Ч, 0) = 0 Then
      --У�Ի���ȼ�ʱ,ͬ�����Ĳ��˻���ȼ�
      If Nvl(v_Ӥ��, 0) = 0 Then
        --���˵�ǰӦ��������סԺ״̬
        v_Temp := Null;
        Begin
          Select Decode(״̬, 1, '�ȴ����', 2, '����ת��', 3, '��Ԥ��Ժ', Null)
          Into v_Temp
          From ������ҳ
          Where ����id = v_����id And ��ҳid = v_��ҳid;
        Exception
          When Others Then
            Null;
        End;
        If v_Temp Is Not Null Then
          v_Error := '���˵�ǰ����' || v_Temp || '״̬,ҽ��"' || v_ҽ������ || '"����ͨ��У�ԡ�';
          Raise Err_Custom;
        End If;
      
        Begin
          --�����շѶ��մ�����ǰҽ���Ƽ۱�û����д
          --δ����ʱ,��������ͬʱ,�������ж��ʱ,ֻȡһ����
          Select a.�շ���Ŀid
          Into v_����ȼ�id
          From �����շѹ�ϵ A, �շ���ĿĿ¼ B
          Where a.�շ���Ŀid = b.Id And b.��� = 'H' And Nvl(b.��Ŀ����, 0) <> 0 And a.������Ŀid = v_������Ŀid And Rownum = 1 And
                Not Exists
           (Select 1 From ������ҳ Where ����id = v_����id And ��ҳid = v_��ҳid And ����ȼ�id = a.�շ���Ŀid);
        Exception
          When Others Then
            Null;
        End;
      End If;
    
      --�䶯��¼��ʱ������룬�Ա���˲���ʱ����ͬһ���ֵ�У�ԡ�ֹͣ�Ȳ���
      v_��ʼʱ�� := To_Date(To_Char(v_��ʼʱ��, 'yyyy-mm-dd hh24:mi') || To_Char(Sysdate, 'ss'), 'yyyy-mm-dd hh24:mi:ss');
      If v_����ȼ�id Is Not Null Then
        Zl_���˱䶯��¼_Nurse(v_����id, v_��ҳid, v_����ȼ�id, v_��ʼʱ��, v_��Ա���, v_��Ա����);
      End If;
    
      --��ֹͣ��������ȼ�ҽ��(����ȼ�Ӧ�ö�Ϊ"������"����,��ֻ��һ��δͣ)
      For r_Nurse In c_Nurse Loop
        Select Decode(Sign(r_Nurse.��ʼִ��ʱ�� - v_��ʼʱ��), 1, r_Nurse.��ʼִ��ʱ��, v_��ʼʱ��)
        Into v_��ֹʱ��
        From Dual;
        Select Decode(Sign(r_Nurse.ִ����ֹʱ�� - v_��ʼʱ��), -1, r_Nurse.ִ����ֹʱ��, v_��ʼʱ��)
        Into v_��ֹʱ��
        From Dual;
        Zl_����ҽ����¼_ֹͣ(r_Nurse.ҽ��id, v_��ֹʱ��, v_����ҽ��, 1);
        v_Stopadviceids := v_Stopadviceids || ',' || r_Nurse.ҽ��id;
      End Loop;
    Elsif v_������� = 'Z' And v_�������� In ('9', '10') And Nvl(v_��Ч, 0) = 0 And Nvl(v_Ӥ��, 0) = 0 Then
      --���ز�Σҽ����9-����;10-��Σ
      --ֹͣ��ͬҽ��
      For r_Patistate In c_Patistate Loop
        Select Decode(Sign(r_Patistate.��ʼִ��ʱ�� - v_��ʼʱ��), 1, r_Patistate.��ʼִ��ʱ��, v_��ʼʱ��)
        Into v_��ֹʱ��
        From Dual;
        Select Decode(Sign(r_Patistate.ִ����ֹʱ�� - v_��ʼʱ��), -1, r_Patistate.ִ����ֹʱ��, v_��ʼʱ��)
        Into v_��ֹʱ��
        From Dual;
        Zl_����ҽ����¼_ֹͣ(r_Patistate.ҽ��id, v_��ֹʱ��, v_����ҽ��, 1);
        v_Stopadviceids := v_Stopadviceids || ',' || r_Patistate.ҽ��id;
      End Loop;
    
      --��������䶯
      Open c_Oldinfo; --�����ڴ���֮ǰ�ȴ�
      Fetch c_Oldinfo
        Into r_Oldinfo;
      Open c_Endinfo;
      Fetch c_Endinfo
        Into r_Endinfo;
      If c_Endinfo%RowCount = 0 Then
        Close c_Endinfo;
        v_Error := 'δ���ָò��˵�ǰ��Ч�ı䶯��¼��';
        Raise Err_Custom;
      End If;
      Select Count(*)
      Into v_Count
      From ���˱䶯��¼
      Where ����id = v_����id And ��ҳid = v_��ҳid And ��ʼʱ�� Is Null And ��ֹʱ�� Is Null;
      If v_Count > 0 Then
        v_Error := '���˵�ǰ����ת��״̬�����Ȱ���ת��ȷ�ϻ���ȡ��ת��״̬��';
        Raise Err_Custom;
      End If;
    
      Update ������ҳ
      Set ��ǰ���� = Decode(v_��������, '9', '��', '10', 'Σ')
      Where ����id = v_����id And ��ҳid = v_��ҳid;
    
      --ȡ���ϴα䶯
      If r_Oldinfo.��ֹʱ�� Is Not Null Then
        v_�䶯��ֹʱ�� := r_Oldinfo.��ֹʱ��;
        v_�䶯��ֹԭ�� := r_Oldinfo.��ֹԭ��;
        v_�䶯��ֹ��Ա := r_Oldinfo.��ֹ��Ա;
        --ȡ���ϴα䶯
        Update ���˱䶯��¼
        Set ��ֹʱ�� = v_��ʼʱ��, ��ֹԭ�� = 13, ��ֹ��Ա = v_��Ա����, �ϴμ���ʱ�� = Null
        Where ����id = v_����id And ��ҳid = v_��ҳid And ��ֹʱ�� = v_�䶯��ֹʱ�� And ��ֹԭ�� = v_�䶯��ֹԭ��;
        --���½����ļ�¼�����ֹͣ����������ɾ���ϴμ���ʱ��
        Update ���˱䶯��¼
        Set ���� = Decode(v_��������, '9', '��', '10', 'Σ'), �ϴμ���ʱ�� = Null
        Where ����id = v_����id And ��ҳid = v_��ҳid And ��ʼʱ�� > v_��ʼʱ��;
      Else
        Update ���˱䶯��¼
        Set ��ֹʱ�� = v_��ʼʱ��, ��ֹԭ�� = 13, ��ֹ��Ա = v_��Ա����,
            �ϴμ���ʱ�� = Decode(Sign(Nvl(�ϴμ���ʱ��, v_��ʼʱ��) - v_��ʼʱ��), 1, Null, �ϴμ���ʱ��)
        Where ����id = v_����id And ��ҳid = v_��ҳid And ��ֹʱ�� Is Null;
      End If;
    
      While c_Oldinfo%Found Loop
        Insert Into ���˱䶯��¼
          (ID, ����id, ��ҳid, ��ʼʱ��, ��ʼԭ��, ���Ӵ�λ, ����id, ����id, ����ȼ�id, ��λ�ȼ�id, ����, ���λ�ʿ, ����ҽʦ, ����ҽʦ, ����ҽʦ, ����, ����Ա���, ����Ա����,
           ��ֹʱ��, ��ֹԭ��, ��ֹ��Ա)
        Values
          (���˱䶯��¼_Id.Nextval, v_����id, v_��ҳid, v_��ʼʱ��, 13, r_Oldinfo.���Ӵ�λ, r_Oldinfo.����id, r_Oldinfo.����id,
           r_Oldinfo.����ȼ�id, r_Oldinfo.��λ�ȼ�id, r_Oldinfo.����, r_Oldinfo.���λ�ʿ, r_Oldinfo.����ҽʦ, r_Oldinfo.����ҽʦ,
           r_Oldinfo.����ҽʦ, Decode(v_��������, '9', '��', '10', 'Σ'), v_��Ա���, v_��Ա����, v_�䶯��ֹʱ��, v_�䶯��ֹԭ��, v_�䶯��ֹ��Ա);
      
        Fetch c_Oldinfo
          Into r_Oldinfo;
      End Loop;
    
      Close c_Oldinfo;
      Close c_Endinfo;
    Elsif v_������� = 'Z' And v_�������� = '12' And Nvl(v_��Ч, 0) = 0 And Nvl(v_Ӥ��, 0) = 0 Then
      --��¼�������ҽ��������
      For r_Patiio In c_Patiio Loop
        Select Decode(Sign(r_Patiio.��ʼִ��ʱ�� - v_��ʼʱ��), 1, r_Patiio.��ʼִ��ʱ��, v_��ʼʱ��)
        Into v_��ֹʱ��
        From Dual;
        Select Decode(Sign(r_Patiio.ִ����ֹʱ�� - v_��ʼʱ��), -1, r_Patiio.ִ����ֹʱ��, v_��ʼʱ��)
        Into v_��ֹʱ��
        From Dual;
        Zl_����ҽ����¼_ֹͣ(r_Patiio.ҽ��id, v_��ֹʱ��, v_����ҽ��, 1);
        v_Stopadviceids := v_Stopadviceids || ',' || r_Patiio.ҽ��id;
      End Loop;
    Elsif (v_������� = 'Z' And v_�������� In ('3', '4', '5', '6', '11', '14') And
          (v_�������� <> '14' Or v_�������� = '14' And v_ִ�б�� = 1)) Or (v_������� = 'F' And v_ִ�б�� = 1) Then
      v_Count := 0;
      If v_�������� = '4' Or v_�������� = '14' Or v_������� = 'F' Then
        --��������ǰУ��ʱ��ͬ�Ĵ���
        If Nvl(v_Ӥ��, 0) = 0 Then
          v_Count := 1;
        End If;
      Else
        --�⼸������ҽ����У����ֹͣҽ�����¼ӵ����ݣ������뷢������ͬ�Ĵ���
        v_Count := 1;
        If Nvl(v_Ӥ��, 0) = 0 Then
          v_Ӥ�� := -1;
        Else
          v_Ӥ�� := Nvl(v_Ӥ��, 0);
        End If;
      End If;
      If v_Count = 1 Then
        If v_������� = 'F' And v_ִ�б�� = 1 Then
          --����������(ȡ��)ֹͣ
          v_��ʼʱ�� := Trunc(To_Date(v_�걾��λ, 'yyyy-mm-dd hh24:mi:ss'));
        End If;
      
        --��������ҽ��У��ʱֹͣǰ��ĳ���,��ҽ����ʼʱ��ֹ��3-ת��;4-����;5-��Ժ;6-תԺ,11-����,14-��ǰ
        For r_Needstop In c_Needstop(v_����id, v_��ҳid, v_Ӥ��, v_��ʼʱ��) Loop
          Select Decode(Sign(��ʼִ��ʱ�� - v_��ʼʱ��), 1, ��ʼִ��ʱ��, v_��ʼʱ��)
          Into v_ֹͣʱ��
          From ����ҽ����¼
          Where ID = r_Needstop.Id;
          Update ����ҽ����¼
          Set ҽ��״̬ = 8, ִ����ֹʱ�� = v_ֹͣʱ��, ͣ��ʱ�� = У��ʱ��_In, ͣ��ҽ�� = v_����ҽ��
          Where ID = r_Needstop.Id;
        
          Insert Into ����ҽ��״̬
            (ҽ��id, ��������, ������Ա, ����ʱ��)
            Select ID, 8, v_��Ա����, У��ʱ��_In From ����ҽ����¼ Where ID = r_Needstop.Id;
          v_Stopadviceids := v_Stopadviceids || ',' || r_Needstop.Id;
        End Loop;
        --��ֹͣδȷ�ϵĳ���,��ֹʱ����ҽ����ʼ���,��ǰ����ֹʱ��(ͬʱ�������ҽ�������)
        For r_Havestop In c_Havestop(v_����id, v_��ҳid, v_Ӥ��, v_��ʼʱ��) Loop
          Update ����ҽ����¼
          Set ִ����ֹʱ�� = Decode(Sign(��ʼִ��ʱ�� - v_��ʼʱ��), 1, ��ʼִ��ʱ��, v_��ʼʱ��), ͣ��ʱ�� = У��ʱ��_In, ͣ��ҽ�� = v_����ҽ��
          Where ID = r_Havestop.Id;
        
          --���޸�ֹͣҽ���Ĳ�����Ա����Ϊֹͣʱ��ҽ�������ѽ��е���ǩ��
          Update ����ҽ��״̬ Set ����ʱ�� = У��ʱ��_In Where ҽ��id = r_Havestop.Id And �������� = 8;
          v_Stopadviceids := v_Stopadviceids || ',' || r_Havestop.Id;
        End Loop;
        --�����ڱ���ҽ��(û��ִ�У����ͣ����ı��δ�ã�
        Update ����ҽ����¼
        Set ִ�б�� = -1
        Where ����id = v_����id And ��ҳid = v_��ҳid And ҽ����Ч = 0 And ִ��Ƶ�� = '��Ҫʱ' And �ϴ�ִ��ʱ�� Is Null And ҽ��״̬ In (3, 5, 6, 7) And
              ִ�б�� <> -1;
        --�����תԺת��������Ժҽ��ͬʱ������ʱ����ҽ����
        If v_�������� In ('3', '5', '6', '11') Then
          Update ����ҽ����¼
          Set ִ�б�� = -1
          Where ����id = v_����id And ��ҳid = v_��ҳid And ҽ����Ч = 1 And ִ��Ƶ�� = '��Ҫʱ' And ҽ��״̬ = 3 And ִ�б�� <> -1;
        End If;
      End If;
    Elsif v_������� = 'Z' And v_�������� = '2' Then
      --�����۲����´���Ժ֪ͨ;
      --ԤԼ�Ǽǵ�������1.��ǰ��ԤԼ,2.��ǰ���������۲��ˣ���ԺʱҲ������Ϊ��Ҫ��ԤԼ,��Ժ����ʱ����˱����Ժ����ܽ��գ�
      Select Count(*) Into v_Count From ������ҳ Where ����id = v_����id And Nvl(��ҳid, 0) = 0;
      If v_Count = 0 Then
        Select Count(*) Into v_Count From ������ҳ Where ����id = v_����id And ��ҳid = v_��ҳid And �������� <> 1;
      End If;
      If v_Count = 0 Then
        Open c_Pati(v_����id);
        Fetch c_Pati
          Into r_Pati;
        Close c_Pati;
      
        v_��Ժ��ʽ := Null;
        If v_������־ = 1 Then
          v_��Ժ��ʽ := '����';
        End If;
      
        Zl_��Ժ������ҳ_Insert(1, 0, r_Pati.����id, r_Pati.סԺ��, Null, r_Pati.����, r_Pati.�Ա�, r_Pati.����, r_Pati.�ѱ�, r_Pati.��������,
                         r_Pati.����, r_Pati.����, r_Pati.ѧ��, r_Pati.����״��, r_Pati.ְҵ, r_Pati.���, r_Pati.���֤��, r_Pati.�����ص�,
                         r_Pati.��ͥ��ַ, r_Pati.��ͥ��ַ�ʱ�, r_Pati.��ͥ�绰, r_Pati.���ڵ�ַ, r_Pati.���ڵ�ַ�ʱ�, r_Pati.��ϵ������, r_Pati.��ϵ�˹�ϵ,
                         r_Pati.��ϵ�˵�ַ, r_Pati.��ϵ�˵绰, r_Pati.������λ, r_Pati.��ͬ��λid, r_Pati.��λ�绰, r_Pati.��λ�ʱ�, r_Pati.��λ������,
                         r_Pati.��λ�ʺ�, r_Pati.������, r_Pati.������, r_Pati.��������, v_ִ�п���id, Null, Null, v_��Ժ��ʽ, Null, Null,
                         v_����ҽ��, r_Pati.����, r_Pati.����, v_��ʼʱ��, Null, Null, r_Pati.ҽ�Ƹ��ʽ, Null, Null, Null, Null, Null,
                         Null, r_Pati.����, v_��Ա���, v_��Ա����, 0, Null, Null, 0);
      End If;
    End If;
    --ҽ��ֹͣ��Ϣ�Ĵ���
    If v_Stopadviceids Is Not Null Then
      v_Stopadviceids := Substr(v_Stopadviceids, 2);
      Select Max(a.Id)
      Into n_���
      From ����ҽ����¼ A
      Where a.Id In (Select Column_Value From Table(f_Num2list(v_Stopadviceids))) And a.ҽ����Ч = 0 And a.ҽ��״̬ = 8 And
            Nvl(a.ִ�б��, 0) <> -1 And a.������Դ <> 3;
      If n_��� Is Not Null Then
        Select Max(a.Id)
        Into n_Adviceid
        From ����ҽ����¼ A
        Where a.Id In (Select Column_Value From Table(f_Num2list(v_Stopadviceids))) And a.������־ = 1 And a.ҽ����Ч = 0 And
              a.ҽ��״̬ = 8 And Nvl(a.ִ�б��, 0) <> -1 And a.������Դ <> 3;
        If n_Adviceid Is Not Null Then
          n_Adviceid := n_���;
          Select Nvl(Max(0), 2)
          Into n_���
          From ҵ����Ϣ�嵥 A
          Where a.����id = v_����id And a.����id = v_��ҳid And a.���ͱ��� = 'ZLHIS_CIS_002' And a.���ȳ̶� = 2 And a.�Ƿ����� = 0;
        Else
          Select Nvl(Max(0), 1)
          Into n_���
          From ҵ����Ϣ�嵥 A
          Where a.����id = v_����id And a.����id = v_��ҳid And a.���ͱ��� = 'ZLHIS_CIS_002' And a.�Ƿ����� = 0;
        End If;
        If n_��� > 0 Then
          For R In (Select a.�������� As ����, a.��Ժ����id As ����id, a.��ǰ����id As ����id
                    From ������ҳ A
                    Where a.����id = v_����id And a.��ҳid = v_��ҳid) Loop
            Zl_ҵ����Ϣ�嵥_Insert(v_����id, v_��ҳid, r.����id, r.����id, r.����, '����ֹͣҽ����', '0010', 'ZLHIS_CIS_002', n_Adviceid, n_���,
                             0, Null, r.����id);
          End Loop;
        End If;
      End If;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ҽ����¼_У��;
/
--111037:��ΰ��,2018-06-24,�������Ǽ�����¼������ʱ��
--124866:������,2018-06-06,���ز�Σҽ���䶯
--125261:������,2018-05-08,ת��ҽ��У�Է��ʹ����Զ�ͣ����
Create Or Replace Procedure Zl_����ҽ������_Insert
(
  ҽ��id_In     ����ҽ������.ҽ��id%Type,
  ���ͺ�_In     ����ҽ������.���ͺ�%Type,
  ��¼����_In   ����ҽ������.��¼����%Type,
  No_In         ����ҽ������.No%Type,
  ��¼���_In   ����ҽ������.��¼���%Type,
  ��������_In   ����ҽ������.��������%Type,
  �״�ʱ��_In   ����ҽ������.�״�ʱ��%Type,
  ĩ��ʱ��_In   ����ҽ������.ĩ��ʱ��%Type,
  ����ʱ��_In   ����ҽ������.����ʱ��%Type,
  ִ��״̬_In   ����ҽ������.ִ��״̬%Type,
  ִ�в���id_In ����ҽ������.ִ�в���id%Type,
  �Ʒ�״̬_In   ����ҽ������.�Ʒ�״̬%Type,
  First_In      Number := 0,
  ��������_In   ����ҽ������.��������%Type := Null,
  ����Ա���_In ��Ա��.���%Type := Null,
  ����Ա����_In ��Ա��.����%Type := Null,
  ��ҩ��_In     δ��ҩƷ��¼.��ҩ��%Type := Null,
  �������_In   ����ҽ������.�������%Type := Null,
  �ֽ�ʱ��_In   Varchar2 := Null
  --���ܣ���д����ҽ�����ͼ�¼
  --������
  --      ҽ��id_In=Ҫ���͵�ÿ��ҽ��ID
  --      First_IN=��ʾ�Ƿ�һ��ҽ���ĵ�һҽ����,�Ա㴦��ҽ���������(���ҩ,�䷽�ĵ�һ��,��Ϊ��ҩ;��,�䷽�巨,�÷�����Ϊ����������)
  --      ��������_IN,�״�ʱ��_IN,ĩ��ʱ��_IN:��"������"����,����д��������,����д��ĩ��ʱ��(���ڻ���)��
  --      �������_In,סԺ�������͵��������ʱ����дΪ1����Ϊ��¼������2����������סԺ���ʣ��������������ա�
) Is
  --�������˼�ҽ��(һ��ҽ���е�һ��)�����Ϣ���α�
  Cursor c_Advice Is
    Select Nvl(a.���id, a.Id) As ��id, a.���, a.����id, a.��ҳid, a.Ӥ��, a.����, a.���˿���id, c.��������, a.�������, a.ҽ����Ч, a.ҽ��״̬, a.ҽ������,
           a.����ҽ��, a.����ʱ��, a.��ʼִ��ʱ��, a.�ϴ�ִ��ʱ��, a.ִ����ֹʱ��, a.ִ��ʱ�䷽��, a.Ƶ�ʴ���, a.Ƶ�ʼ��, a.�����λ, a.��������id, a.�걾��λ, a.ִ�п���id,
           a.���id, a.������Ŀid
    From ����ҽ����¼ A, ������ĿĿ¼ C
    Where a.������Ŀid = c.Id And a.Id = ҽ��id_In;
  r_Advice c_Advice%RowType;

  --��������(Ӥ��)������δͣ����(���䷽����),Ӥ������-1��ʾ������
  Cursor c_Needstop
  (
    v_����id   ����ҽ����¼.����id%Type,
    v_��ҳid   ����ҽ����¼.��ҳid%Type,
    v_Ӥ��     ����ҽ����¼.Ӥ��%Type,
    v_Stoptime Date
  ) Is
    Select a.Id, a.�������, b.��������, b.ִ��Ƶ��
    From ����ҽ����¼ A, ������ĿĿ¼ B
    Where a.������Ŀid = b.Id(+) And a.����id = v_����id And a.��ҳid = v_��ҳid And (v_Ӥ�� = -1 Or Nvl(a.Ӥ��, 0) = Nvl(v_Ӥ��, 0)) And
          Nvl(a.ҽ����Ч, 0) = 0 And a.ҽ��״̬ Not In (1, 2, 4, 8, 9) And a.��ʼִ��ʱ�� <= v_Stoptime
    Order By a.���;
  --��������(Ӥ��)����ͣ��δȷ�ϵĳ���,��ִֹ��ʱ����ָ��ʱ��֮��,Ӥ������-1��ʾ������
  Cursor c_Havestop
  (
    v_����id   ����ҽ����¼.����id%Type,
    v_��ҳid   ����ҽ����¼.��ҳid%Type,
    v_Ӥ��     ����ҽ����¼.Ӥ��%Type,
    v_Stoptime Date
  ) Is
    Select ID
    From ����ҽ����¼
    Where ����id = v_����id And ��ҳid = v_��ҳid And (v_Ӥ�� = -1 Or Nvl(Ӥ��, 0) = Nvl(v_Ӥ��, 0)) And Nvl(ҽ����Ч, 0) = 0 And
          ҽ��״̬ = 8 And ִ����ֹʱ�� > v_Stoptime And ��ʼִ��ʱ�� < v_Stoptime
    Order By ���;

  --������ʱ����
  v_Ӥ��     ����ҽ����¼.Ӥ��%Type;
  v_������   Number(1); --�Ƿ�����Գ���
  v_Autostop Number(1);
  v_Date     Date;
  v_Temp     Varchar2(255);
  v_��Ա��� ��Ա��.���%Type;
  v_��Ա���� ��Ա��.����%Type;
  v_ֹͣʱ�� ����ҽ����¼.����ʱ��%Type;
  n_ִ��״̬ ����ҽ������.ִ��״̬%Type;
  d_��ʼʱ�� ����ҽ����¼.��ʼִ��ʱ��%Type;

  v_Stopadviceids ����ҽ����¼.ҽ������%Type;
  n_Adviceid      ����ҽ����¼.����id%Type;
  n_���          Number(18);
  v_Error         Varchar2(255);
  Err_Custom Exception;
Begin
  --��ǰ������Ա
  If ����Ա���_In Is Not Null And ����Ա����_In Is Not Null Then
    v_��Ա��� := ����Ա���_In;
    v_��Ա���� := ����Ա����_In;
  Else
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_��Ա��� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
    v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  End If;
  --����״�ʱ��Ϊ�������뿪ʼִ��ʱ��
  If �״�ʱ��_In Is Null Or �ֽ�ʱ��_In Is Null Or ĩ��ʱ��_In Is Null Then
    Select ��ʼִ��ʱ�� Into d_��ʼʱ�� From ����ҽ����¼ Where ID = ҽ��id_In;
  End If;

  Open c_Advice;
  Fetch c_Advice
    Into r_Advice;
  Close c_Advice;

  --��һ��ҽ���ĵ�һ��ʱ����ҽ������
  If Nvl(First_In, 0) = 1 Then
    --�����������
    ---------------------------------------------------------------------------------------
    If Nvl(r_Advice.ҽ��״̬, 0) = 4 Then
      --���Ҫ���͵�ҽ���Ƿ�����
      v_Error := '"' || r_Advice.���� || '"��ҽ��"' || r_Advice.ҽ������ || '"�Ѿ������������ϡ�' || Chr(13) || Chr(10) ||
                 '�ò��˵�ҽ������ʧ�ܡ������¶�ȡ�����嵥���ԡ�';
      Raise Err_Custom;
    End If;
  
    If Nvl(r_Advice.ҽ����Ч, 0) = 0 Then
      --����������ҩ����,�䷽����,��ҩ"��ѡƵ��"����,��ҩ"������"����
    
      --��鳤���Ƿ��ѱ�����
      If r_Advice.�ϴ�ִ��ʱ�� Is Not Null Then
        If r_Advice.�ϴ�ִ��ʱ�� >= �״�ʱ��_In Then
          v_Error := '"' || r_Advice.���� || '"��ҽ��"' || r_Advice.ҽ������ || '"�Ѿ��������˷��͡�' || Chr(13) || Chr(10) ||
                     '�ò��˵�ҽ������ʧ�ܡ������¶�ȡ�����嵥���ԡ�';
          Raise Err_Custom;
        End If;
      End If;
    
      --��鳤������ǰ�Ƿ��ѱ��Զ�ֹͣ(������)
      If r_Advice.ִ����ֹʱ�� Is Not Null Then
        If �״�ʱ��_In > r_Advice.ִ����ֹʱ�� Then
          v_Error := '"' || r_Advice.���� || '"��ҽ��"' || r_Advice.ҽ������ || '"�Ѿ���ֹͣ��' || Chr(13) || Chr(10) ||
                     '�ò��˵�ҽ������ʧ�ܡ������¶�ȡ�����嵥���ԡ�';
          Raise Err_Custom;
        End If;
      End If;
    Elsif Nvl(r_Advice.ҽ��״̬, 0) In (8, 9) Then
      --���������䷽����
    
      --����Ƿ��ѱ�����(��������ԭ���Զ�ֹͣ)
      v_Error := '"' || r_Advice.���� || '"��ҽ��"' || r_Advice.ҽ������ || '"�Ѿ��������˷��͡�' || Chr(13) || Chr(10) ||
                 '�ò��˵�ҽ������ʧ�ܡ������¶�ȡ�����嵥���ԡ�';
      Raise Err_Custom;
    End If;
  
    --���ͺ��ҽ������
    ---------------------------------------------------------------------------------------
    If Nvl(r_Advice.ҽ����Ч, 0) = 0 Then
      --����ҽ��:�����ϴ�ִ��ʱ��
      Update ����ҽ����¼ Set �ϴ�ִ��ʱ�� = ĩ��ʱ��_In Where ID = r_Advice.��id Or ���id = r_Advice.��id;
    
      --�ж��Ƿ�����Գ���
      v_������ := 0;
      If r_Advice.ִ��ʱ�䷽�� Is Null And (Nvl(r_Advice.Ƶ�ʴ���, 0) = 0 Or Nvl(r_Advice.Ƶ�ʼ��, 0) = 0 Or r_Advice.�����λ Is Null) Then
        v_������ := 1;
      End If;
    
      --Ԥ������ֹʱ����δֹͣ���Զ�ֹͣ
      If r_Advice.ִ����ֹʱ�� Is Not Null And Nvl(r_Advice.ҽ��״̬, 0) Not In (8, 9) Then
        v_Autostop := 0;
        If v_������ = 1 Then
          --��ҩ"������"����
          If Trunc(ĩ��ʱ��_In) = Trunc(r_Advice.ִ����ֹʱ�� - 1) Then
            v_Autostop := 1; --��ֹ���첻ִ��
          End If;
        Elsif Zl_Advicenexttime(ҽ��id_In) > r_Advice.ִ����ֹʱ�� Then
          --��ҩ�������ҩ"��ѡƵ��"����
          v_Autostop := 1; --����ǵ���,������ִ��һ��
        End If;
      
        If v_Autostop = 1 Then
          Update ����ҽ����¼
          Set ҽ��״̬ = 8, ͣ��ʱ�� = ĩ��ʱ��_In, ͣ��ҽ�� = r_Advice.����ҽ��
          Where ID = r_Advice.��id Or ���id = r_Advice.��id;
        
          Insert Into ����ҽ��״̬
            (ҽ��id, ��������, ������Ա, ����ʱ��)
            Select ID, 8, r_Advice.����ҽ��, ����ʱ��_In
            From ����ҽ����¼
            Where ID = r_Advice.��id Or ���id = r_Advice.��id;
          v_Stopadviceids := v_Stopadviceids || ',' || r_Advice.��id;
        End If;
      End If;
    Else
      --����ֹͣ��
      --סԺҽ������ʱ�Զ�У�ԡ�ֹͣ��У������Sysdateȡ��,Ϊ�����ظ�,ֹͣʱ��ҲȡSysdate
      Select Sysdate Into v_Date From Dual;
      Update ����ҽ����¼
      Set ҽ��״̬ = 8, ִ����ֹʱ�� = ĩ��ʱ��_In,
          --Ϊһ��������ʱû��
          �ϴ�ִ��ʱ�� = ĩ��ʱ��_In,
          --Ϊһ��������ʱû��
          ͣ��ʱ�� = v_Date,
          --����ʱ��_IN,
          ͣ��ҽ�� = r_Advice.����ҽ��
      Where ID = r_Advice.��id Or ���id = r_Advice.��id;
    
      Insert Into ����ҽ��״̬
        (ҽ��id, ��������, ������Ա, ����ʱ��)
        Select ID, 8, v_��Ա����, v_Date --����ʱ��_IN
        From ����ҽ����¼
        Where ID = r_Advice.��id Or ���id = r_Advice.��id;
    End If;
  
    --����ҽ���Ĵ���
    ---------------------------------------------------------------------------------------
    If r_Advice.������� = 'Z' And Nvl(r_Advice.��������, '0') <> '0' Then
      --(1-����;2-סԺ;)3-ת��;4-����(������);5-��Ժ;6-תԺ,7-����,11-����
    
      --��������ҽ��Ҫ�Զ�ֹͣ���˸�ҽ��֮ǰ(��ʱ����)����δͣ�ĳ���
      If r_Advice.�������� In ('3', '5', '6', '11') Then
        If Nvl(r_Advice.Ӥ��, 0) = 0 Then
          v_Ӥ�� := -1;
        Else
          v_Ӥ�� := Nvl(r_Advice.Ӥ��, 0);
        End If;
        For r_Needstop In c_Needstop(r_Advice.����id, r_Advice.��ҳid, v_Ӥ��, r_Advice.��ʼִ��ʱ��) Loop
          Select Decode(Sign(��ʼִ��ʱ�� - r_Advice.��ʼִ��ʱ��), 1, ��ʼִ��ʱ��, r_Advice.��ʼִ��ʱ��)
          Into v_ֹͣʱ��
          From ����ҽ����¼
          Where ID = r_Needstop.Id;
          Update ����ҽ����¼
          Set ҽ��״̬ = 8, ִ����ֹʱ�� = v_ֹͣʱ��, ͣ��ʱ�� = ����ʱ��_In, ͣ��ҽ�� = r_Advice.����ҽ��
          Where ID = r_Needstop.Id;
        
          Insert Into ����ҽ��״̬
            (ҽ��id, ��������, ������Ա, ����ʱ��)
            Select ID, 8, v_��Ա����, ����ʱ��_In From ����ҽ����¼ Where ID = r_Needstop.Id;
          v_Stopadviceids := v_Stopadviceids || ',' || r_Needstop.Id;
        End Loop;
        --��ֹͣδȷ�ϵĳ���,��ֹʱ����ҽ����ʼ���,��ǰ����ֹʱ��(ͬʱ�������ҽ�������)
        For r_Havestop In c_Havestop(r_Advice.����id, r_Advice.��ҳid, v_Ӥ��, r_Advice.��ʼִ��ʱ��) Loop
          Update ����ҽ����¼
          Set ִ����ֹʱ�� = Decode(Sign(��ʼִ��ʱ�� - r_Advice.��ʼִ��ʱ��), 1, ��ʼִ��ʱ��, r_Advice.��ʼִ��ʱ��), ͣ��ʱ�� = ����ʱ��_In,
              ͣ��ҽ�� = r_Advice.����ҽ��
          Where ID = r_Havestop.Id;
        
          --���޸�ֹͣҽ���Ĳ�����Ա����Ϊֹͣʱ��ҽ�������ѽ��е���ǩ��
          Update ����ҽ��״̬ Set ����ʱ�� = ����ʱ��_In Where ҽ��id = r_Havestop.Id And �������� = 8;
          v_Stopadviceids := v_Stopadviceids || ',' || r_Havestop.Id;
        End Loop;
        --�����ڱ���ҽ��(û��ִ�У����ͣ����ı��δ�ã�,ͬʱ��������
        Update ����ҽ����¼
        Set ִ�б�� = -1
        Where ����id = r_Advice.����id And ��ҳid = r_Advice.��ҳid And
              (ҽ����Ч = 0 And ִ��Ƶ�� = '��Ҫʱ' And �ϴ�ִ��ʱ�� Is Null And ҽ��״̬ In (3, 5, 6, 7) Or
              ҽ����Ч = 1 And ִ��Ƶ�� = '��Ҫʱ' And ҽ��״̬ = 3) And ִ�б�� <> -1;
      End If;
    
      --��������⴦��
      If Nvl(r_Advice.Ӥ��, 0) = 0 Then
        If r_Advice.�������� = '3' And ִ�в���id_In Is Not Null And r_Advice.���˿���id Is Not Null And
           Nvl(r_Advice.���˿���id, 0) <> Nvl(ִ�в���id_In, 0) Then
          --ת��ҽ��,�����˵Ǽ�ת�Ƶ�"ִ�п���ID"(��Ժ�����ҵ�ǰ������ת����Ҳ�ͬ�Ŵ���)
          Zl_���˱䶯��¼_Change(r_Advice.����id, r_Advice.��ҳid, ִ�в���id_In, v_��Ա���, v_��Ա����);
        Elsif r_Advice.�������� In ('5', '6', '11') Then
          --��Ժ��תԺ������ҽ��,�����˱��ΪԤ��Ժ
          Begin
            Select ��ʼʱ��
            Into v_Date
            From ���˱䶯��¼
            Where ��ʼʱ�� Is Not Null And ��ֹʱ�� Is Null And ����id = r_Advice.����id And ��ҳid = r_Advice.��ҳid;
          Exception
            When Others Then
              v_Date := To_Date('1900-01-01', 'YYYY-MM-DD');
          End;
          If r_Advice.��ʼִ��ʱ�� < v_Date Then
            v_Error := 'ҽ��"' || r_Advice.ҽ������ || '"�Ŀ�ʼʱ��Ӧ���ڸò����ϴα䶯ʱ�� ' || To_Char(v_Date, 'YYYY-MM-DD HH24:Mi') || ' ��';
            Raise Err_Custom;
          End If;
          Zl_���˱䶯��¼_Preout(r_Advice.����id, r_Advice.��ҳid, r_Advice.��ʼִ��ʱ��);
        End If;
      Else
        If r_Advice.�������� = '11' Then
          Update ������������¼
          Set ����ʱ�� = r_Advice.��ʼִ��ʱ��
          Where ����id = r_Advice.����id And Nvl(��ҳid, 0) = Nvl(r_Advice.��ҳid, 0) And Nvl(���, 0) = Nvl(r_Advice.Ӥ��, 0);
        End If;
      End If;
    End If;
    --12Сʱδִ�еı�����������Ϊ���δ��
    If r_Advice.ҽ����Ч = 1 Then
      Update ����ҽ����¼
      Set ִ�б�� = -1
      Where ����id = r_Advice.����id And ��ҳid = r_Advice.��ҳid And ִ�б�� <> -1 And ҽ����Ч = 1 And ִ��Ƶ�� = '��Ҫʱ' And
            Sysdate - ��ʼִ��ʱ�� > 0.5 And ҽ��״̬ = 3;
    End If;
  End If;

  --��д���ͼ�¼
  ---------------------------------------------------------------------------------------
  n_ִ��״̬ := ִ��״̬_In;
  If ִ��״̬_In = 1 Then
    v_Temp := zl_GetSysParameter(186);
    If v_Temp = '11' Then
      If r_Advice.������� = 'E' And r_Advice.�������� in ('1','8') Or r_Advice.������� = 'K' Then
        n_ִ��״̬ := 0;
      End If;
    Elsif v_Temp = '01' Then
      If r_Advice.������� = 'E' And r_Advice.�������� = '1' Then
        n_ִ��״̬ := 0;
      End If;
    Elsif v_Temp = '10' Then
      If r_Advice.������� = 'E' And r_Advice.�������� = '8' Or r_Advice.������� = 'K' Then
        n_ִ��״̬ := 0;
      End If;    
    End If;
  End If;

  Insert Into ����ҽ������
    (ҽ��id, ���ͺ�, ��¼����, NO, ��¼���, ��������, ������, ����ʱ��, ִ��״̬, ִ�в���id, �Ʒ�״̬, �״�ʱ��, ĩ��ʱ��, ��������, �������)
  Values
    (ҽ��id_In, ���ͺ�_In, ��¼����_In, No_In, ��¼���_In, ��������_In, v_��Ա����, ����ʱ��_In, n_ִ��״̬, ִ�в���id_In, �Ʒ�״̬_In,
     Nvl(�״�ʱ��_In, d_��ʼʱ��), Nvl(ĩ��ʱ��_In, d_��ʼʱ��), ��������_In, �������_In);

  --�����ͼ��ҽ��ͬ��������ҽ���ļƷ�״̬   
  If �Ʒ�״̬_In = 1 And  r_Advice.��ID <> ҽ��id_In  And (r_Advice.������� = 'D' Or r_Advice.������� = 'F') Then   
     Update ����ҽ������ Set �Ʒ�״̬ = 1 Where ҽ��ID = r_Advice.��ID And ���ͺ� = ���ͺ�_In;
  End If;

  --��ҩ�ŵ���д
  If ��ҩ��_In Is Not Null Then
    Update δ��ҩƷ��¼ Set ��ҩ�� = ��ҩ��_In Where NO = No_In And ���� = 9 And ��ҩ�� Is Null;
    Update ҩƷ�շ���¼ Set ��Ʒ�ϸ�֤ = ��ҩ��_In Where NO = No_In And ���� = 9 And ��Ʒ�ϸ�֤ Is Null;
  End If;

  --�Զ���Ϊ��ִ��ʱ����Ҫͬ���������ִ��״̬����˻���״̬
  If ִ��״̬_In = 1 Then
    Zl_����ҽ��ִ��_Finish(ҽ��id_In, ���ͺ�_In, Null, Null, v_��Ա���, v_��Ա����, ִ�в���id_In);
  End If;

  --����ҽ��ִ��ʱ���¼(ֻ��������¼��)
  If Nvl(�ֽ�ʱ��_In, To_Char(d_��ʼʱ��, 'yyyy-mm-dd hh24:mi:ss')) Is Not Null Then
    If r_Advice.���id Is Null Then
      Insert Into ҽ��ִ��ʱ��
        (Ҫ��ʱ��, ҽ��id, ���ͺ�)
        Select To_Date(Column_Value, 'yyyy-mm-dd hh24:mi:ss'), ҽ��id_In, ���ͺ�_In
        From Table(f_Str2list(Nvl(�ֽ�ʱ��_In, To_Char(d_��ʼʱ��, 'yyyy-mm-dd hh24:mi:ss'))));
    End If;
  End If;

  --������дʱ������д
  If r_Advice.������� = 'F' Then
    --һ������ֻ��һ��
    If r_Advice.���id Is Null Then
      If Not r_Advice.�걾��λ Is Null Then
        v_Date := To_Date(r_Advice.�걾��λ, 'yyyy-mm-dd hh24:mi:ss');
      Else
        v_Date := r_Advice.��ʼִ��ʱ��;
      End If;
      Zl_���Ӳ���ʱ��_Insert(r_Advice.����id, r_Advice.��ҳid, 2, '����', r_Advice.��������id, r_Advice.����ҽ��, v_Date, v_Date,
                       r_Advice.ִ�п���id);
    End If;
  Elsif r_Advice.������� = 'Z' And r_Advice.�������� = '7' Then
    Zl_���Ӳ���ʱ��_Insert(r_Advice.����id, r_Advice.��ҳid, 2, '����', r_Advice.��������id, r_Advice.����ҽ��, r_Advice.��ʼִ��ʱ��,
                     r_Advice.��ʼִ��ʱ��, r_Advice.ִ�п���id);
  Elsif r_Advice.������� = 'Z' And r_Advice.�������� = '8' Then
    Zl_���Ӳ���ʱ��_Insert(r_Advice.����id, r_Advice.��ҳid, 2, '����', r_Advice.��������id, r_Advice.����ҽ��, r_Advice.��ʼִ��ʱ��,
                     r_Advice.��ʼִ��ʱ��, r_Advice.ִ�п���id);
  Elsif r_Advice.������� = 'Z' And r_Advice.�������� = '11' Then
    Zl_���Ӳ���ʱ��_Insert(r_Advice.����id, r_Advice.��ҳid, 2, '����', r_Advice.��������id, r_Advice.����ҽ��, r_Advice.��ʼִ��ʱ��,
                     r_Advice.��ʼִ��ʱ��, r_Advice.ִ�п���id);
  End If;
  --�������(֪���ļ�������������ŵ���)
  If Instr('C,D,E,F,G,K,L', r_Advice.�������) > 0 Then
    Zl_���Ӳ���ʱ��_Insert(r_Advice.����id, r_Advice.��ҳid, 2, '֪������', r_Advice.��������id, r_Advice.����ҽ��, r_Advice.��ʼִ��ʱ��,
                     r_Advice.��ʼִ��ʱ��, r_Advice.ִ�п���id, r_Advice.������Ŀid, r_Advice.ҽ������);
  End If;
  --ҽ��ֹͣ��Ϣ�Ĵ���
  If v_Stopadviceids Is Not Null Then
    v_Stopadviceids := Substr(v_Stopadviceids, 2);
    Select Max(a.Id)
    Into n_���
    From ����ҽ����¼ A
    Where a.Id In (Select Column_Value From Table(f_Num2list(v_Stopadviceids))) And a.ҽ����Ч = 0 And a.ҽ��״̬ = 8 And
          Nvl(a.ִ�б��, 0) <> -1 And a.������Դ <> 3;
    If n_��� Is Not Null Then
      Select Max(a.Id)
      Into n_Adviceid
      From ����ҽ����¼ A
      Where a.Id In (Select Column_Value From Table(f_Num2list(v_Stopadviceids))) And a.������־ = 1 And a.ҽ����Ч = 0 And
            a.ҽ��״̬ = 8 And Nvl(a.ִ�б��, 0) <> -1 And a.������Դ <> 3;
      If n_Adviceid Is Not Null Then
        Select Nvl(Max(0), 2)
        Into n_���
        From ҵ����Ϣ�嵥 A
        Where a.����id = r_Advice.����id And a.����id = r_Advice.��ҳid And a.���ͱ��� = 'ZLHIS_CIS_002' And a.���ȳ̶� = 2 And
              a.�Ƿ����� = 0;
      Else
        n_Adviceid := n_���;
        Select Nvl(Max(0), 1)
        Into n_���
        From ҵ����Ϣ�嵥 A
        Where a.����id = r_Advice.����id And a.����id = r_Advice.��ҳid And a.���ͱ��� = 'ZLHIS_CIS_002' And a.�Ƿ����� = 0;
      End If;
      If n_��� > 0 Then
        For R In (Select a.�������� As ����, a.��Ժ����id As ����id, a.��ǰ����id As ����id
                  From ������ҳ A
                  Where a.����id = r_Advice.����id And a.��ҳid = r_Advice.��ҳid) Loop
          Zl_ҵ����Ϣ�嵥_Insert(r_Advice.����id, r_Advice.��ҳid, r.����id, r.����id, r.����, '����ֹͣҽ����', '0010', 'ZLHIS_CIS_002',
                           n_Adviceid, n_���, 0, Null, r.����id);
        End Loop;
      End If;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ҽ������_Insert;
/

--124269:������,2018-05-07,���ڻ��������´�ҽ�����
Create Or Replace Procedure Zl_����ҽ����¼_Insert
(
  Id_In           In ����ҽ����¼.Id%Type,
  ���id_In       In ����ҽ����¼.���id%Type,
  ���_In         In ����ҽ����¼.���%Type,
  ������Դ_In     In ����ҽ����¼.������Դ%Type,
  ����id_In       In ����ҽ����¼.����id%Type,
  ��ҳid_In       In ����ҽ����¼.��ҳid%Type,
  Ӥ��_In         In ����ҽ����¼.Ӥ��%Type,
  ҽ��״̬_In     In ����ҽ����¼.ҽ��״̬%Type,
  ҽ����Ч_In     In ����ҽ����¼.ҽ����Ч%Type,
  �������_In     In ����ҽ����¼.�������%Type,
  ������Ŀid_In   In ����ҽ����¼.������Ŀid%Type,
  �շ�ϸĿid_In   In ����ҽ����¼.�շ�ϸĿid%Type,
  ����_In         In ����ҽ����¼.����%Type,
  ��������_In     In ����ҽ����¼.��������%Type,
  �ܸ�����_In     In ����ҽ����¼.�ܸ�����%Type,
  ҽ������_In     In ����ҽ����¼.ҽ������%Type,
  ҽ������_In     In ����ҽ����¼.ҽ������%Type,
  �걾��λ_In     In ����ҽ����¼.�걾��λ%Type,
  ִ��Ƶ��_In     In ����ҽ����¼.ִ��Ƶ��%Type,
  Ƶ�ʴ���_In     In ����ҽ����¼.Ƶ�ʴ���%Type,
  Ƶ�ʼ��_In     In ����ҽ����¼.Ƶ�ʼ��%Type,
  �����λ_In     In ����ҽ����¼.�����λ%Type,
  ִ��ʱ�䷽��_In In ����ҽ����¼.ִ��ʱ�䷽��%Type,
  �Ƽ�����_In     In ����ҽ����¼.�Ƽ�����%Type,
  ִ�п���id_In   In ����ҽ����¼.ִ�п���id%Type,
  ִ������_In     In ����ҽ����¼.ִ������%Type,
  ������־_In     In ����ҽ����¼.������־%Type,
  ��ʼִ��ʱ��_In In ����ҽ����¼.��ʼִ��ʱ��%Type,
  ִ����ֹʱ��_In In ����ҽ����¼.ִ����ֹʱ��%Type,
  ���˿���id_In   In ����ҽ����¼.���˿���id%Type,
  ��������id_In   In ����ҽ����¼.��������id%Type,
  ����ҽ��_In     In ����ҽ����¼.����ҽ��%Type,
  ����ʱ��_In     In ����ҽ����¼.����ʱ��%Type,
  �Һŵ�_In       In ����ҽ����¼.�Һŵ�%Type := Null,
  ǰ��id_In       In ����ҽ����¼.ǰ��id%Type := Null,
  ��鷽��_In     In ����ҽ����¼.��鷽��%Type := Null,
  ִ�б��_In     In ����ҽ����¼.ִ�б��%Type := Null,
  �ɷ����_In     In ����ҽ����¼.�ɷ����%Type := Null,
  ժҪ_In         In ����ҽ����¼.ժҪ%Type := Null,
  ����Ա����_In   In ����ҽ��״̬.������Ա%Type := Null,
  ��Ѽ���_In     In ����ҽ����¼.��Ѽ���%Type := Null,
  ��ҩĿ��_In     In ����ҽ����¼.��ҩĿ��%Type := Null,
  ��ҩ����_In     In ����ҽ����¼.��ҩ����%Type := Null,
  ���״̬_In     In ����ҽ����¼.���״̬%Type := Null,
  �������_In     In ����ҽ����¼.�������%Type := Null,
  ����˵��_In     In ����ҽ����¼.����˵��%Type := Null,
  �״�����_In     In ����ҽ����¼.�״�����%Type := Null,
  �䷽id_In       In ����ҽ����¼.�䷽id%Type := Null,
  �������_In     In ����ҽ����¼.�������%Type := Null,
  �����Ŀid_In   In ����ҽ����¼.�����Ŀid%Type := Null,
  Ƥ�Խ��_In     In ����ҽ����¼.Ƥ�Խ��%Type := Null,
  �������_In     In ����ҽ����¼.�������%Type := Null,
  ����ҽ��id_In   In ����ҽ����¼.����ҽ��id%Type := Null
  --���ܣ�ҽ����ʿ�¿�,��¼ҽ��ʱ�²�����ҽ����¼�������������סԺ��
) Is
  v_Temp     Varchar2(255);
  v_��Ա���� ����ҽ��״̬.������Ա%Type;

  v_����     ������Ϣ.����%Type;
  v_�Ա�     ������Ϣ.�Ա�%Type;
  v_����     ������Ϣ.����%Type;
  d_����ʱ�� ����ҽ����¼.����ʱ��%Type;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --��ǰ������Ա
  If ����Ա����_In Is Not Null Then
    v_��Ա���� := ����Ա����_In;
  Else
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  End If;

  If Nvl(��ҳid_In, 0) <> 0 Then
    Select ����, �Ա�, ���� Into v_����, v_�Ա�, v_���� From ������ҳ Where ����id = ����id_In And ��ҳid = ��ҳid_In;
  Else
    Select ����, �Ա�, ���� Into v_����, v_�Ա�, v_���� From ������Ϣ Where ����id = ����id_In;
  End If;

  If Instr(',F,K,', �������_In) > 0 Then
    d_����ʱ�� := To_Date(�걾��λ_In, 'yyyy-mm-dd hh24:mi:ss');
  End If;

  --����ҽ����¼
  Insert Into ����ҽ����¼
    (ID, ���id, ���, ������Դ, ����id, ��ҳid, ����, �Ա�, ����, Ӥ��, ҽ��״̬, ҽ����Ч, �������, ������Ŀid, �շ�ϸĿid, ����, ��������, �ܸ�����, ҽ������, ҽ������, �걾��λ,
     ��鷽��, ִ�б��, ִ��Ƶ��, Ƶ�ʴ���, Ƶ�ʼ��, �����λ, ִ��ʱ�䷽��, �Ƽ�����, ִ�п���id, ִ������, ������־, �ɷ����, ��ʼִ��ʱ��, ִ����ֹʱ��, ���˿���id, ��������id, ����ҽ��,
     ����ʱ��, �Һŵ�, ǰ��id, ժҪ, ��Ѽ���, ����ʱ��, ��ҩĿ��, ��ҩ����, ���״̬, �������, ����˵��, �״�����, �䷽id, �������, �����Ŀid, Ƥ�Խ��, �������, ����ҽ��id)
  Values
    (Id_In, ���id_In, ���_In, ������Դ_In, ����id_In, ��ҳid_In, v_����, v_�Ա�, v_����, Ӥ��_In, ҽ��״̬_In, ҽ����Ч_In, �������_In, ������Ŀid_In,
     �շ�ϸĿid_In, ����_In, ��������_In, �ܸ�����_In, ҽ������_In, ҽ������_In, �걾��λ_In, ��鷽��_In, ִ�б��_In, ִ��Ƶ��_In, Ƶ�ʴ���_In, Ƶ�ʼ��_In, �����λ_In,
     ִ��ʱ�䷽��_In, �Ƽ�����_In, ִ�п���id_In, ִ������_In, ������־_In, �ɷ����_In, ��ʼִ��ʱ��_In, ִ����ֹʱ��_In, ���˿���id_In, ��������id_In, ����ҽ��_In,
     ����ʱ��_In, �Һŵ�_In, ǰ��id_In, ժҪ_In, ��Ѽ���_In, d_����ʱ��, ��ҩĿ��_In, ��ҩ����_In, ���״̬_In, �������_In, ����˵��_In, �״�����_In, �䷽id_In,
     �������_In, �����Ŀid_In, Ƥ�Խ��_In, �������_In, ����ҽ��id_In);

  --����ҽ��״̬
  If ҽ��״̬_In <> -1 Then
    Delete From ����ҽ��״̬ Where ҽ��id = Id_In And �������� = 1;
    If Sql%RowCount <> 0 Then
      v_Error := '��ͬID���¿�ҽ���Ѿ����ڡ�';
      Raise Err_Custom;
    End If;
    --��Ϊ����ͬʱ���¿�->�Զ�У��(סԺҽ������)->�����Զ�ֹͣ(סԺҽ����������ֹͣ),��˷ֱ�-2,-1��
    Insert Into ����ҽ��״̬
      (ҽ��id, ��������, ������Ա, ����ʱ��)
    Values
      (Id_In, 1, v_��Ա����, Sysdate - 2 / 60 / 60 / 24);
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ҽ����¼_Insert;
/

--124269:������,2018-05-07,���ڻ��������´�ҽ�����
Create Or Replace Procedure Zl_����ҽ����¼_Update
(
  Id_In           In ����ҽ����¼.Id%Type,
  ���id_In       In ����ҽ����¼.���id%Type,
  ���_In         In ����ҽ����¼.���%Type,
  ҽ��״̬_In     In ����ҽ����¼.ҽ��״̬%Type,
  ҽ����Ч_In     In ����ҽ����¼.ҽ����Ч%Type,
  ������Ŀid_In   In ����ҽ����¼.������Ŀid%Type,
  �շ�ϸĿid_In   In ����ҽ����¼.�շ�ϸĿid%Type,
  ����_In         In ����ҽ����¼.����%Type,
  ��������_In     In ����ҽ����¼.��������%Type,
  �ܸ�����_In     In ����ҽ����¼.�ܸ�����%Type,
  ҽ������_In     In ����ҽ����¼.ҽ������%Type,
  ҽ������_In     In ����ҽ����¼.ҽ������%Type,
  �걾��λ_In     In ����ҽ����¼.�걾��λ%Type,
  ִ��Ƶ��_In     In ����ҽ����¼.ִ��Ƶ��%Type,
  Ƶ�ʴ���_In     In ����ҽ����¼.Ƶ�ʴ���%Type,
  Ƶ�ʼ��_In     In ����ҽ����¼.Ƶ�ʼ��%Type,
  �����λ_In     In ����ҽ����¼.�����λ%Type,
  ִ��ʱ�䷽��_In In ����ҽ����¼.ִ��ʱ�䷽��%Type,
  �Ƽ�����_In     In ����ҽ����¼.�Ƽ�����%Type,
  ִ�п���id_In   In ����ҽ����¼.ִ�п���id%Type,
  ִ������_In     In ����ҽ����¼.ִ������%Type,
  ������־_In     In ����ҽ����¼.������־%Type,
  ��ʼִ��ʱ��_In In ����ҽ����¼.��ʼִ��ʱ��%Type,
  ִ����ֹʱ��_In In ����ҽ����¼.ִ����ֹʱ��%Type,
  ���˿���id_In   In ����ҽ����¼.���˿���id%Type,
  ��������id_In   In ����ҽ����¼.��������id%Type,
  ����ҽ��_In     In ����ҽ����¼.����ҽ��%Type,
  ����ʱ��_In     In ����ҽ����¼.����ʱ��%Type,
  ��鷽��_In     In ����ҽ����¼.��鷽��%Type := Null,
  ִ�б��_In     In ����ҽ����¼.ִ�б��%Type := Null,
  �ɷ����_In     In ����ҽ����¼.�ɷ����%Type := Null,
  ժҪ_In         In ����ҽ����¼.ժҪ%Type := Null,
  ��Ա������_In   In ����ҽ��״̬.������Ա%Type := Null,
  ��Ѽ���_In     In ����ҽ����¼.��Ѽ���%Type := Null,
  ��ҩĿ��_In     In ����ҽ����¼.��ҩĿ��%Type := Null,
  ��ҩ����_In     In ����ҽ����¼.��ҩ����%Type := Null,
  ���״̬_In     In ����ҽ����¼.���״̬%Type := Null,
  ����˵��_In     In ����ҽ����¼.����˵��%Type := Null,
  �״�����_In     In ����ҽ����¼.�״�����%Type := Null,
  �������_In     In ����ҽ����¼.�������%Type := Null,
  �����Ŀid_In   In ����ҽ����¼.�����Ŀid%Type := Null,
  Ƥ�Խ��_In     In ����ҽ����¼.Ƥ�Խ��%Type := Null,
  �������_In     In ����ҽ����¼.�������%Type := Null,
  ����ҽ��id_In   In ����ҽ����¼.����ҽ��id%Type := Null
  --���ܣ���ҽ����ʿ�޸��˲������ݵ�ҽ����¼�������������סԺ��
  --˵����Updateʱ֮�����漰������ĿID,�Ƽ����Ա仯,����Ϊ��ҩ;��,�÷��ı仯
  --      Updateʱ֮�����漰��Ч�仯,����Ϊ����¼��ҽ��������ı���Ч
) Is
  v_Count Number;

  v_Temp            Varchar2(255);
  v_��Ա����        ����ҽ��״̬.������Ա%Type;
  v_�����������ids Varchar2(4000);

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --����ҽ��״̬:��������
  Begin
    Select ҽ��״̬ Into v_Count From ����ҽ����¼ Where ID = Id_In;
  Exception
    When Others Then
      Begin
        v_Error := 'ҽ��"' || ҽ������_In || '"�Ѿ�������,�����ѱ�������ɾ����';
        Raise Err_Custom;
      End;
  End;
  If v_Count Not In (-1, 1, 2) Then
    v_Error := 'ҽ��"' || ҽ������_In || '"�Ѿ�У�Ի���,�������޸ġ�';
    Raise Err_Custom;
  End If;

  Select Count(*) Into v_Count From ����ҽ��״̬ Where ҽ��id = Id_In And �������� = 1 And ǩ��id Is Not Null;
  If Nvl(v_Count, 0) > 0 Then
    v_Error := 'ҽ��"' || ҽ������_In || '"�Ѿ�����ǩ��,�������޸ġ�';
    Raise Err_Custom;
  End If;

  --������鳷��
  If ���id_In Is Null Then
    Zl_�������_Cancel(Id_In, v_�����������ids);
  End If;

  If v_�����������ids Is Not Null Then
    v_Error := 'ҽ��"' || ҽ������_In || '"�����������ڽ��д�����飬�������޸ġ�';
    Raise Err_Custom;
  End If;

  --��ǰ������Ա
  If ��Ա������_In Is Not Null Then
    v_��Ա���� := ��Ա������_In;
  Else
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  End If;

  --����ҽ����¼
  Update ����ҽ����¼
  Set ���id = ���id_In,
      --����һ����ҩ���������ü�鲿λ����������ID�仯
      ��� = ���_In, ҽ��״̬ = ҽ��״̬_In,
      --!��Ϊֻ���޸�δУ��ҽ��������Ӧ��Ϊ�¿���У�����ʵ�ҽ���޸ĺ�Ϊ�¿�
      ҽ����Ч = ҽ����Ч_In, ������Ŀid = ������Ŀid_In, �շ�ϸĿid = �շ�ϸĿid_In, ���� = ����_In, �������� = ��������_In, �ܸ����� = �ܸ�����_In, ҽ������ = ҽ������_In,
      ҽ������ = ҽ������_In, �걾��λ = �걾��λ_In, ��鷽�� = ��鷽��_In, ִ�б�� = ִ�б��_In, ִ��Ƶ�� = ִ��Ƶ��_In, Ƶ�ʴ��� = Ƶ�ʴ���_In, Ƶ�ʼ�� = Ƶ�ʼ��_In,
      �����λ = �����λ_In, ִ��ʱ�䷽�� = ִ��ʱ�䷽��_In, �Ƽ����� = �Ƽ�����_In, ִ�п���id = ִ�п���id_In, ִ������ = ִ������_In, �ɷ���� = �ɷ����_In,
      --ҩƷ�����⹺ҩ,��Ժ��ҩ�ĵ���ʱ�ᷢ���仯
      ������־ = ������־_In, ��ʼִ��ʱ�� = ��ʼִ��ʱ��_In, ִ����ֹʱ�� = ִ����ֹʱ��_In,
      --!��������ֹʱ������޸�,����Ӧ��Ϊ��
      ���˿���id = ���˿���id_In,
      --�޸�ʱ����Ϊ���˵ĵ�ǰ����
      ��������id = ��������id_In,
      --�޸ĺ����ݵ�ǰ���ұ仯
      ����ҽ�� = ����ҽ��_In, ��˱�� = Decode(Nvl(Instr(����ҽ��_In, '/'), 0), 0, Decode(��˱��, 1, Null, ��˱��), 1),
      --��ʿ��ҽ��ʱ���Ը���
      ����ʱ�� = ����ʱ��_In,
      --��¼�Ŀ����޸�
      ժҪ = ժҪ_In, ��Ѽ��� = ��Ѽ���_In, ����ʱ�� = Decode(�������, 'F', To_Date(�걾��λ_In, 'yyyy-mm-dd hh24:mi:ss'), Null),
      ��ҩĿ�� = ��ҩĿ��_In, ��ҩ���� = ��ҩ����_In, ���״̬ = ���״̬_In, ����˵�� = ����˵��_In, �״����� = �״�����_In, ������� = �������_In, �����Ŀid = �����Ŀid_In,
      Ƥ�Խ�� = Ƥ�Խ��_In,
      --������ҩ���
      ������� = �������_In, ����ҽ��id = ����ҽ��id_In
  Where ID = Id_In;

  --����ҽ��״̬:����ҽ���¿�����
  --��Ϊ����ͬʱ���¿�(�޸�)->�Զ�У��(סԺҽ������)->�����Զ�ֹͣ(סԺҽ����������ֹͣ),��˷ֱ�-2,-1��
  If ҽ��״̬_In <> -1 Then
    Update ����ҽ��״̬
    Set ������Ա = v_��Ա����, ����ʱ�� = Sysdate - 2 / 60 / 60 / 24
    Where ҽ��id = Id_In And �������� = 1; --�¿�����ʼ����,У�����ʱ�����Ϊ��ʷ��¼
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ҽ����¼_Update;
/

--125083:������,2018-05-03,Σ��ֵɾ��
Create Or Replace Procedure Zl_����Σ��ֵ��¼_Delete(Id_In In ����Σ��ֵ��¼.Id%Type) Is
Begin
  Delete ����Σ��ֵ��¼ Where ID = Id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Σ��ֵ��¼_Delete;
/

--113688:��¶¶,2018-04-28,ȡ������Ժ���������۲��˵ĵǼǺ�����Ϣ�е����Ժ��Ϣû�и���
Create Or Replace Procedure Zl_��Ժ������ҳ_Delete
(
  ����id_In     ������ҳ.����id%Type,
  ��ҳid_In     ������ҳ.��ҳid%Type,
  ת����_In     Number := 0,
  ���סԺ��_In Number := 0
  --���ܣ�ȡ��������Ժ/ԤԼ�Ǽ�
  --     ��ҳID_IN:Ϊ0ʱ��ʾȡ��ԤԼ�Ǽ�
  --     ת����_IN:��������Ժ�Ǽǲ���תΪסԺ���۲���
  --     ���סԺ��_In:��һ��סԺ�Ĳ���ת����ʱ�Ƿ����סԺ��
) As
  v_��Ժʱ��   ������ҳ.��Ժ����%Type;
  v_��Ժ����   ������ҳ.��Ժ����id%Type;
  v_��Ժʱ��   ������ҳ.��Ժ����%Type;
  v_סԺ��     ������ҳ.סԺ��%Type;
  v_����Ժ     ������ҳ.����Ժ%Type;
  v_��Ժ����id ������ҳ.��Ժ����id%Type;
  n_��������   ������ҳ.��������%Type;
  n_��ҳid     ������ҳ.��ҳid%Type;

  v_Count Number;
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  Select Nvl(״̬, 0), Nvl(��������, 0)
  Into v_Count, n_��������
  From ������ҳ
  Where ����id = ����id_In And ��ҳid = ��ҳid_In;
  If v_Count <> 1 Then
    v_Error := '�ò����Ѿ����,���Ƚ����˳�������Ժ״̬��';
    Raise Err_Custom;
  End If;

  --ɾ�����Ӳ���ʱ��
  Select ��Ժ����id, ����Ժ Into v_��Ժ����id, v_����Ժ From ������ҳ Where ����id = ����id_In And ��ҳid = ��ҳid_In;
  If v_����Ժ = 0 Then
    Zl_���Ӳ���ʱ��_Delete(����id_In, ��ҳid_In, '��Ժ', v_��Ժ����id);
  Else
    Zl_���Ӳ���ʱ��_Delete(����id_In, ��ҳid_In, '�ٴ���Ժ', v_��Ժ����id);
  End If;

  --��ȡ���һ�β�Ϊ�յ�סԺ��
  Begin
    If ��ҳid_In = 0 Then
      Select סԺ��
      Into v_סԺ��
      From ������ҳ
      Where ����id = ����id_In And
            ��ҳid =
            (Select Max(��ҳid) From ������ҳ Where ����id = ����id_In And Nvl(��ҳid, 0) <> 0 And Nvl(סԺ��, 0) <> 0);
    Else
      Select סԺ��
      Into v_סԺ��
      From ������ҳ
      Where ����id = ����id_In And
            ��ҳid =
            (Select Max(��ҳid) From ������ҳ Where ����id = ����id_In And ��ҳid < ��ҳid_In And Nvl(סԺ��, 0) <> 0);
    End If;
  Exception
    When Others Then
      Null;
  End;

  If ת����_In = 1 And Nvl(��ҳid_In, 0) <> 0 Then
    Update ������ҳ
    Set �������� = 2, סԺ�� = Decode(���סԺ��_In, 1, Null, סԺ��)
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And Nvl(��������, 0) = 0;
  
    --����סԺ����
    Update ������Ϣ Set סԺ���� = Decode(Sign(סԺ���� - 1), 1, סԺ���� - 1, Null) Where ����id = ����id_In;
    If ���סԺ��_In = 1 Then
      Update ������Ϣ Set סԺ�� = v_סԺ�� Where ����id = ����id_In;
    End If;
  Else
    Begin
      Select b.��Ժ����, b.��Ժ����, b.��Ժ����id
      Into v_��Ժʱ��, v_��Ժʱ��, v_��Ժ����
      From ������Ϣ A, ������ҳ B
      Where a.����id = ����id_In And a.����id = b.����id And a.��ҳid = b.��ҳid And Nvl(b.��ҳid, 0) <> 0;
    Exception
      When Others Then
        Null;
    End;
    --����ԤԼ�Ǽǲ��˲����סԺ�ձ�
    If Nvl(��ҳid_In, 0) <> 0 Then
      Select Zl_סԺ�ձ�_Count(v_��Ժ����, v_��Ժʱ��) Into v_Count From Dual;
      If v_Count > 0 Then
        v_Error := '�Ѳ���ҵ��ʱ���ڵ�סԺ�ձ�,���ܰ����ҵ��!';
        Raise Err_Custom;
      End If;
    End If;
    --�������۲����´���Ժ֪ͨ�����������Ч�Ĳ�����ҳ��¼��36549��
    Select Count(*) Into v_Count From ������ҳ Where ����id = ����id_In And ��Ժ���� Is Not Null And ��Ժ���� Is Null;
    If Not v_Count > 1 Then
      v_Count := 0;
      If Nvl(��ҳid_In, 0) <> 0 And Nvl(n_��������, 0) = 0 Then
        v_Count := 1;
      End If;
      --����Ժ����,ȡ����Ժ�Ǽ�ʱ,������Ϣ����Ժʱ��ͳ�Ժʱ��Ӧ�û��˵���һ����Ժ���ںͳ�Ժ����
      If v_����Ժ = 1 Then
        Begin
          Select ��Ժ����, ��Ժ����
          Into v_��Ժʱ��, v_��Ժʱ��
          From ������ҳ
          Where ����id = ����id_In And
                ��ҳid = (Select Max(��ҳid)
                        From ������ҳ
                        Where ����id = ����id_In And ��ҳid < ��ҳid_In);
    	Exception
      		When Others Then
        	Null;
        End;
      End If;    
      Update ������Ϣ
      Set סԺ�� = v_סԺ��, סԺ���� = Decode(v_Count, 0, סԺ����, Decode(Sign(סԺ���� - 1), 1, סԺ���� - 1, Null)), ��ǰ����id = Null,
          ��ǰ����id = Null, ��ǰ���� = Null, ��Ժʱ�� = v_��Ժʱ��, ��Ժʱ�� = v_��Ժʱ��, ������ = Null, ������ = Null, �������� = Null, ��Ժ = Null
      Where ����id = ����id_In;
      Delete From ��Ժ���� Where ����id = ����id_In;
    End If;
    Delete From ���˱䶯��¼ Where ����id = ����id_In And ��ҳid = ��ҳid_In;
    Delete From ������ϼ�¼ Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼��Դ = 2;
  
    --����סԺ�������Ԥ����,��Ϊ�������ｻ��
    Update ����Ԥ����¼ Set ��ҳid = Null Where ����id = ����id_In And ��ҳid = ��ҳid_In;
  
    --���η�����,�ı����﷢��
    Update סԺ���ü�¼ Set ��ҳid = Null Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼���� = 5;
  
    --����סԺ�����з��ü�¼�޽�������ȫ���������򽫶�Ӧ���ü�¼�е�"��ҳID"�����
    v_Count := 0;
    Select Nvl(Count(*), 0)
    Into v_Count
    From סԺ���ü�¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ���ʷ��� = 1 And ����id Is Not Null;
  
    If v_Count = 0 Then
      Begin
        Select Nvl(Count(*), 0)
        Into v_Count
        From סԺ���ü�¼
        Where ����id = ����id_In And ��ҳid = ��ҳid_In And ���ʷ��� = 1
        Group By NO, ��¼����, ���
        Having Nvl(Sum(ʵ�ս��), 0) <> 0;
      Exception
        When Others Then
          v_Count := 0;
      End;
    
      If v_Count = 0 Then
        Delete ����δ����� Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��� = 0;
        Update סԺ���ü�¼ Set ��ҳid = Null Where ����id = ����id_In And ��ҳid = ��ҳid_In And ���ʷ��� = 1;
      End If;
    End If;
  
    --����סԺ����ҽ����¼��������
    v_Count := 0;
    Select Nvl(Count(*), 0)
    Into v_Count
    From ����ҽ����¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And Nvl(ҽ��״̬, 0) <> 4;
    If v_Count = 0 Then
      Delete From ����ҽ����¼ Where ����id = ����id_In And ��ҳid = ��ҳid_In;
    End If;
  
    --���±�,û�н�������ҳ(����ID,��ҳID)�����,��Ϊ����ҳID�����ǹҺ�ID
    Delete From ���˹�����¼ Where ����id = ����id_In And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0);
    Delete From ������ϼ�¼ Where ����id = ����id_In And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0);
    Delete From ������������¼ Where ����id = ����id_In And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0);
    Delete From ���Ӳ�����¼ Where ����id = ����id_In And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0);
    Delete From ���Ӳ�����ӡ Where ����id = ����id_In And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0);
    --�����Ժ�����˾��￨,��ɾ����ʧ��(���˷��ü�¼��ҳID�����Լ��)
    Delete From ������ҳ Where ����id = ����id_In And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0);
    --�޸Ĳ�����Ϣ����ҳID��סԺ����
    Select Max(��ҳid) Into n_��ҳid From ������ҳ Where ����id = ����id_In And Nvl(��ҳid, 0) <> 0;
    Update ������Ϣ Set ��ҳid = n_��ҳid Where ����id = ����id_In;
    If n_��ҳid Is Null Then
      Update ������Ϣ Set סԺ���� = Null Where ����id = ����id_In;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Ժ������ҳ_Delete;
/
--124963:Ƚ����,2018-04-28,Ԥ����������ʹ�ú��ش��˷�Ʊ�����������תסԺʱ����
Create Or Replace Procedure Zl_����תסԺ_�շ�ת��
(
  No_In         סԺ���ü�¼.No%Type,
  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
  �˷�ʱ��_In   סԺ���ü�¼.����ʱ��%Type,
  �����˷�_In   Number := 0,
  ��Ժ����id_In סԺ���ü�¼.��������id%Type := Null,
  ��ҳid_In     סԺ���ü�¼.��ҳid%Type := Null,
  ���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type := Null,
  ����id_In     ����Ԥ����¼.����id%Type := Null,
  ԭ����id_In   ����Ԥ����¼.����id%Type := Null,
  ����_In     ����Ԥ����¼.��Ԥ��%Type := Null
) As
  --�����˷�_In:0-����תסԺ��������;1-�����˷�ģʽ
  -- �����˷�_InΪ1ʱ:��Ժ����id_In����ҳID_IN���Բ�����
  n_Count      Number(5);
  n_ԭ����id   סԺ���ü�¼.����id%Type;
  n_ʵ�ս��   ������ü�¼.ʵ�ս��%Type;
  n_Ԥ��ʹ�ö� ����Ԥ����¼.��Ԥ��%Type;
  n_ʵ�ʳ���   ����Ԥ����¼.��Ԥ��%Type;
  n_��id       ����ɿ����.Id%Type;
  n_����id     ������Ϣ.����id%Type;
  v_Ԥ��no     ����Ԥ����¼.No%Type;
  n_Ԥ�����   ����Ԥ����¼.��Ԥ��%Type;
  n_��ӡid     Ʊ��ʹ����ϸ.��ӡid%Type;
  n_��������id סԺ���ü�¼.��������id%Type;
  v_������     ������ü�¼.������%Type;
  n_����id     ������ü�¼.����id%Type;
  n_����     ����Ԥ����¼.��Ԥ��%Type;
  v_����     ���㷽ʽ.����%Type;
  n_����ֵ     �������.�������%Type;
  v_���㷽ʽ   ���㷽ʽ.����%Type;
  v_Nos        Varchar2(3000);
  v_����ids    Varchar2(3000);
  v_ԭ����ids  Varchar2(3000);
  n_Tempid     ����Ԥ����¼.Id%Type;
  n_Ԥ��id     ����Ԥ����¼.Id%Type;
  n_ҽ��       Number;
  n_����       Number;
  n_����       Number;
  n_�����˷�   Number;
  n_�˷�����   Number;
  n_�쳣��־   Number;
  n_�������   Number;
  n_����״̬   ������ü�¼.����״̬%Type;

  Err_Item Exception;
  v_Err_Msg Varchar2(200);

  Procedure Zl_Square_Update
  (
    ����ids_In    Varchar2,
    �ֽ���id_In   ����Ԥ����¼.����id%Type,
    �ɿ���id_In   ����Ԥ����¼.�ɿ���id%Type,
    �˿�ʱ��_In   ����Ԥ����¼.�տ�ʱ��%Type,
    �������_In   ����Ԥ����¼.�������%Type,
    ��������_In   Varchar2 := Null,
    �˷ѽ��_In   ����Ԥ����¼.��Ԥ��%Type := Null,
    ���㿨���_In ����Ԥ����¼.���㿨���%Type := Null
  ) As
    n_��¼״̬ ���˿������¼.��¼״̬%Type;
    n_Ԥ��id   ����Ԥ����¼.Id%Type;
    v_����     ���˿������¼.����%Type;
    n_���ڿ�Ƭ Number;
    d_ͣ������ ���ѿ�Ŀ¼.ͣ������%Type;
    n_������ ���˿������¼.���%Type;
    n_���     ���˿������¼.���%Type;
    n_���     ���ѿ�Ŀ¼.���%Type;
    n_�ӿڱ�� ���˿������¼.�ӿڱ��%Type;
    d_����ʱ�� ���ѿ�Ŀ¼.����ʱ��%Type;
    n_Id       ����Ԥ����¼.Id%Type;
  Begin
    n_Ԥ��id := 0;
  
    --�������ѿ�,���㿨��������Ѿ�������
    For v_У�� In (Select Min(a.Id) As Ԥ��id, c.���ѿ�id, Sum(c.������) As ������, c.�ӿڱ��, c.����, Max(c.���) As ���, Max(c.Id) As ID
                 From ����Ԥ����¼ A, ���˿�������� B, ���˿������¼ C
                 Where a.Id = b.Ԥ��id And a.���㿨��� = ���㿨���_In And b.������id = c.Id And a.��¼���� = 3 And
                       Instr(Nvl(��������_In, '_LXH'), ',' || a.���㷽ʽ || ',') = 0 And
                       a.����id In (Select Column_Value From Table(f_Str2list(����ids_In)))
                 Group By c.���ѿ�id, c.�ӿڱ��, c.����) Loop
    
      If Nvl(v_У��.���ѿ�id, 0) <> 0 Then
        Select Max(��¼״̬)
        Into n_��¼״̬
        From ���˿������¼
        Where �ӿڱ�� = v_У��.�ӿڱ�� And ���ѿ�id = Nvl(v_У��.���ѿ�id, 0) And ���� = v_У��.���� And Nvl(���, 0) = Nvl(v_У��.���, 0);
      Else
        Select Max(��¼״̬)
        Into n_��¼״̬
        From ���˿������¼
        Where �ӿڱ�� = v_У��.�ӿڱ�� And ���ѿ�id Is Null And ���� = v_У��.���� And Nvl(���, 0) = Nvl(v_У��.���, 0);
      End If;
    
      If n_��¼״̬ = 1 Then
        n_��¼״̬ := 2;
      Else
        n_��¼״̬ := n_��¼״̬ + 2;
      End If;
      --����ʱ,ֻ����һ��
      If n_Ԥ��id = 0 Then
        Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id,
           Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, У�Ա�־, �������, ��������)
          Select n_Ԥ��id, ��¼����, NO, 2, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �˿�ʱ��_In, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���_In, ����Ա����_In,
                 -1 * �˷ѽ��_In, �ֽ���id_In, �ɿ���id_In, Ԥ�����, �����id, Nvl(���㿨���, v_У��.�ӿڱ��), ����, ������ˮ��, ����˵��, ������λ, 2, �������_In,
                 ��������
          From ����Ԥ����¼ A
          Where ID = v_У��.Ԥ��id;
      End If;
    
      If Nvl(v_У��.���ѿ�id, 0) <> 0 Then
        --���ѿ�,ֱ���˻ؿ�������
        Begin
          Select ����, 1, ͣ������, (Select Max(���) From ���ѿ�Ŀ¼ B Where a.���� = b.���� And a.�ӿڱ�� = b.�ӿڱ��), ���, ���, �ӿڱ��, ����ʱ��
          Into v_����, n_���ڿ�Ƭ, d_ͣ������, n_������, n_���, n_���, n_�ӿڱ��, d_����ʱ��
          From ���ѿ�Ŀ¼ A
          Where ID = v_У��.���ѿ�id;
        Exception
          When Others Then
            n_���ڿ�Ƭ := 0;
        End;
      
        --ȡ��ͣ��
        If n_���ڿ�Ƭ = 0 Then
          v_Err_Msg := '����Ϊ"' || v_���� || '"�����ѿ�������ɾ�������������øÿ�Ƭ,���飡';
          Raise Err_Item;
        End If;
        If Nvl(n_���, 0) < Nvl(n_������, 0) Then
          v_Err_Msg := '����������ʷ������¼(����Ϊ"' || v_���� || '"),���飡';
          Raise Err_Item;
        End If;
        If Nvl(d_ͣ������, To_Date('3000-01-01', 'yyyy-mm-dd')) < To_Date('3000-01-01', 'yyyy-mm-dd') Then
          v_Err_Msg := '����Ϊ"' || v_���� || '"�����ѿ��Ѿ�������ͣ�ã������ٽ����˷�,���飡';
          Raise Err_Item;
        End If;
      
        If d_����ʱ�� < To_Date('3000-01-01', 'yyyy-mm-dd') Then
          v_Err_Msg := '����Ϊ"' || v_���� || '"�����ѿ��Ѿ����գ������˷�,���飡';
          Raise Err_Item;
        End If;
        Update ���ѿ�Ŀ¼ Set ��� = Nvl(���, 0) + �˷ѽ��_In Where ID = Nvl(v_У��.���ѿ�id, 0);
      End If;
    
      Select ���˿������¼_Id.Nextval Into n_Id From Dual;
      Insert Into ���˿������¼
        (ID, �ӿڱ��, ���ѿ�id, ���, ��¼״̬, ���㷽ʽ, ������, ����, ������ˮ��, ����ʱ��, ��ע, �����־)
        Select n_Id, �ӿڱ��, ���ѿ�id, ���, n_��¼״̬, ���㷽ʽ, -1 * �˷ѽ��_In, ����, ������ˮ��, ����ʱ��, ��ע,
               Decode(���ѿ�id, Null, 0, 0, 0, 1) As ��־
        From ���˿������¼
        Where ID = v_У��.Id;
      Insert Into ���˿�������� (Ԥ��id, ������id) Values (n_Ԥ��id, n_Id);
    
      If n_��¼״̬ <> 2 And n_��¼״̬ <> 1 Then
        Update ���˿������¼ Set ��¼״̬ = 3 Where ID = v_У��.Id;
      End If;
    End Loop;
  End;
Begin
  n_��id := Zl_Get��id(����Ա����_In);
  --����
  Begin
    Select ���� Into v_���� From ���㷽ʽ Where ���� = 9 And Rownum < 2;
  Exception
    When Others Then
      v_Err_Msg := 'û�з��������㷽ʽ�������Ƿ���ȷ���ã�';
      Raise Err_Item;
  End;

  If ԭ����id_In Is Null Then
  
    Select Count(NO), Sum(ʵ�ս��)
    Into n_Count, n_ʵ�ս��
    From ������ü�¼
    Where NO = No_In And Mod(��¼����, 10) = 1;
    If n_Count = 0 Or n_ʵ�ս�� = 0 Then
      v_Err_Msg := '����' || No_In || '�����շѵ��ݻ��򲢷�ԭ�����˲����˸õ���,����תΪסԺ����.';
      Raise Err_Item;
    End If;
  
    Select ����id, ����id, ��������id, ������
    Into n_ԭ����id, n_����id, n_��������id, v_������
    From ������ü�¼
    Where NO = No_In And Mod(��¼����, 10) = 1 And ��¼״̬ In (1, 3) And Rownum < 2;
  
    --1.1���Ϸ��ü�¼
    If ����id_In Is Null Then
      Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
    Else
      n_����id := ����id_In;
    End If;
  
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�, ���˿���id, �շ����, �շ�ϸĿid,
       ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ������, ִ����, ִ��״̬, ִ��ʱ��,
       ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id, ����״̬)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�, ���˿���id,
             �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, -1 * ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -1 * Ӧ�ս��, -1 * ʵ�ս��, ��������id,
             ������, ִ�в���id, ������, ִ����, -1, ִ��ʱ��, ����Ա���_In, ����Ա����_In, ����ʱ��, �˷�ʱ��_In, n_����id, -1 * ���ʽ��, ������Ŀ��, ���մ���id, ͳ����,
             ժҪ, Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id, 0
      From ������ü�¼
      Where NO = No_In And Mod(��¼����, 10) = 1 And ��¼״̬ = 1;
  
    --Update ������ü�¼ Set ��¼״̬ = 3 Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 1;
  
    --1.2����Ԥ����¼
    --���ϳ�Ԥ������
    For r_����id In (Select Distinct ����id
                   From ������ü�¼
                   Where NO In (Select Distinct NO
                                From ������ü�¼
                                Where ����id In (Select ����id
                                               From ����Ԥ����¼
                                               Where ������� In (Select b.�������
                                                              From ������ü�¼ A, ����Ԥ����¼ B
                                                              Where a.No = No_In And b.������� < 0 And Mod(a.��¼����, 10) = 1 And
                                                                    a.��¼״̬ <> 0 And a.����id = b.����id))) And
                         Mod(��¼����, 10) = 1 And ��¼״̬ <> 0
                   Union
                   Select Distinct ����id
                   From ������ü�¼
                   Where NO In (Select Distinct NO
                                From ������ü�¼
                                Where ����id In (Select a.����id
                                               From ������ü�¼ A, ����Ԥ����¼ B
                                               Where a.No = No_In And b.������� > 0 And Mod(a.��¼����, 10) = 1 And a.��¼״̬ <> 0 And
                                                     a.����id = b.����id)) And Mod(��¼����, 10) = 1 And ��¼״̬ <> 0) Loop
      v_ԭ����ids := v_ԭ����ids || ',' || r_����id.����id;
    End Loop;
    v_ԭ����ids := Substr(v_ԭ����ids, 2);
  
    Begin
      Select 1
      Into n_ҽ��
      From ���ս����¼
      Where ��¼id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And Rownum < 2;
    Exception
      When Others Then
        n_ҽ�� := 0;
    End;
  
    If n_ҽ�� = 1 Then
      Begin
        Select 1
        Into n_����
        From ҽ��������ϸ
        Where NO = No_In And ����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And Rownum < 2;
      Exception
        When Others Then
          v_Err_Msg := '��ǰ����' || No_In || '������ҽ��������ϸ,�޷���������תסԺ!';
          Raise Err_Item;
      End;
    End If;
  
    --ҽ���˿�
    For r_ҽ�� In (Select ����id, NO, ���㷽ʽ, ���, ��ע
                 From ҽ��������ϸ
                 Where NO = No_In And ����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids)))) Loop
      Update ��Ա�ɿ����
      Set ��� = Nvl(���, 0) - r_ҽ��.���
      Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = r_ҽ��.���㷽ʽ
      Returning ��� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ��Ա�ɿ����
          (�տ�Ա, ���㷽ʽ, ����, ���)
        Values
          (����Ա����_In, r_ҽ��.���㷽ʽ, 1, -1 * r_ҽ��.���);
        n_����ֵ := r_ҽ��.���;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From ��Ա�ɿ����
        Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = r_ҽ��.���㷽ʽ And Nvl(���, 0) = 0;
      End If;
    
      Update ����Ԥ����¼
      Set ��Ԥ�� = ��Ԥ�� + (-1 * r_ҽ��.���)
      Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = r_ҽ��.���㷽ʽ;
      If Sql%RowCount = 0 Then
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ,
           �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, У�Ա�־, ��������)
        Values
          (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * r_ҽ��.���, r_ҽ��.���㷽ʽ, Null, �˷�ʱ��_In,
           Null, Null, Null, ����Ա���_In, ����Ա����_In, r_ҽ��.��ע, n_��id, Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id,
           0, 3);
      End If;
    
      Update ����Ԥ����¼
      Set ��¼״̬ = 3
      Where ��¼���� = 3 And ��¼״̬ = 1 And ����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And
            ���㷽ʽ = r_ҽ��.���㷽ʽ;
    
      Update ҽ��������ϸ
      Set ��� = ��� + (-1 * r_ҽ��.���)
      Where NO = No_In And ����id = n_����id And ���㷽ʽ = r_ҽ��.���㷽ʽ;
      If Sql%RowCount = 0 Then
        Insert Into ҽ��������ϸ
          (����id, NO, ���㷽ʽ, ���)
        Values
          (n_����id, No_In, r_ҽ��.���㷽ʽ, -1 * r_ҽ��.���);
      End If;
      n_ʵ�ս�� := n_ʵ�ս�� - r_ҽ��.���;
    End Loop;
  
    Begin
      Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And ���� Like '%�ֽ�%' And Rownum < 2;
    Exception
      When Others Then
        Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And Rownum < 2;
    End;
  
    If n_ʵ�ս�� <> 0 Then
      For r_Prepay In (Select NO, Max(Decode(��¼����, 1, ʵ��Ʊ��, Null)) As ʵ��Ʊ��, ����id, ��ҳid, ����id, ���㷽ʽ, �������, �ɿλ, ��λ������,
                              ��λ�ʺ�, Sum(��Ԥ��) As ��Ԥ��, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ
                       From ����Ԥ����¼ A
                       Where ��¼���� In (1, 11) And a.����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids)))
                       Group By NO, ����id, ��ҳid, ����id, ���㷽ʽ, �������, �ɿλ, ��λ������, ��λ�ʺ�, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ) Loop
        If n_ʵ�ս�� <> 0 Then
          If r_Prepay.��Ԥ�� >= n_ʵ�ս�� Then
            Select ����Ԥ����¼_Id.Nextval Into n_Tempid From Dual;
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���,
               ��Ԥ��, ����id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, Ԥ�����, �������, �ɿ���id)
              Select n_Tempid, r_Prepay.No, r_Prepay.ʵ��Ʊ��, 11, 1, r_Prepay.����id, r_Prepay.��ҳid, r_Prepay.����id, Null,
                     r_Prepay.���㷽ʽ, r_Prepay.�������, Null, r_Prepay.�ɿλ, r_Prepay.��λ������, r_Prepay.��λ�ʺ�, �˷�ʱ��_In, ����Ա����_In,
                     ����Ա���_In, -1 * n_ʵ�ս��, n_����id, r_Prepay.�����id, r_Prepay.���㿨���, r_Prepay.����, r_Prepay.������ˮ��,
                     r_Prepay.����˵��, r_Prepay.������λ, 1, -1 * n_����id, n_��id
              From Dual;
            Update �������
            Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(n_ʵ�ս��, 0)
            Where ����id = n_����id And ���� = 1 And ���� = 1
            Returning Ԥ����� Into n_����ֵ;
            If Sql%RowCount = 0 Then
              Insert Into ������� (����id, ����, Ԥ�����, ����) Values (n_����id, 1, n_ʵ�ս��, 1);
              n_����ֵ := n_ʵ�ս��;
            End If;
            If Nvl(n_����ֵ, 0) = 0 Then
              Delete From �������
              Where ����id = n_����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
            End If;
            n_ʵ�ս�� := 0;
          Else
            Select ����Ԥ����¼_Id.Nextval Into n_Tempid From Dual;
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���,
               ��Ԥ��, ����id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, Ԥ�����, �������, �ɿ���id)
              Select n_Tempid, r_Prepay.No, r_Prepay.ʵ��Ʊ��, 11, 1, r_Prepay.����id, r_Prepay.��ҳid, r_Prepay.����id, Null,
                     r_Prepay.���㷽ʽ, r_Prepay.�������, Null, r_Prepay.�ɿλ, r_Prepay.��λ������, r_Prepay.��λ�ʺ�, �˷�ʱ��_In, ����Ա����_In,
                     ����Ա���_In, -1 * r_Prepay.��Ԥ��, n_����id, r_Prepay.�����id, r_Prepay.���㿨���, r_Prepay.����, r_Prepay.������ˮ��,
                     r_Prepay.����˵��, r_Prepay.������λ, 1, -1 * n_����id, n_��id
              From Dual;
            Update �������
            Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(r_Prepay.��Ԥ��, 0)
            Where ����id = n_����id And ���� = 1 And ���� = 1
            Returning Ԥ����� Into n_����ֵ;
            If Sql%RowCount = 0 Then
              Insert Into ������� (����id, ����, Ԥ�����, ����) Values (n_����id, 1, r_Prepay.��Ԥ��, 1);
              n_����ֵ := r_Prepay.��Ԥ��;
            End If;
            If Nvl(n_����ֵ, 0) = 0 Then
              Delete From �������
              Where ����id = n_����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
            End If;
            n_ʵ�ս�� := n_ʵ�ս�� - r_Prepay.��Ԥ��;
          End If;
        End If;
      End Loop;
    End If;
    --2.Ʊ���ջ�
    --������ǰû�д�ӡ,���ջ�
    Select Nvl(Max(ID), 0)
    Into n_��ӡid
    From (Select b.Id
           From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
           Where a.��ӡid = b.Id And a.���� = 1 And a.ԭ�� In (1, 3) And b.�������� = 1 And b.No = No_In
           Order By a.ʹ��ʱ�� Desc)
    Where Rownum < 2;
    If n_��ӡid > 0 Then
      --���ŵ���ѭ������ʱֻ���ջ�һ��
      Select Count(��ӡid) Into n_Count From Ʊ��ʹ����ϸ Where Ʊ�� = 1 And ���� = 2 And ��ӡid = n_��ӡid;
      If n_Count = 0 Then
        Insert Into Ʊ��ʹ����ϸ
          (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����)
          Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, �˷�ʱ��_In, ����Ա����_In
          From Ʊ��ʹ����ϸ
          Where ��ӡid = n_��ӡid And Ʊ�� = 1 And ���� = 1;
      End If;
    End If;
  
    --3.�ɿ����ݴ���(
    --   �����������:
    --    1. ת������ֱ�����ʵ�,��ɿ����ݲ�����;
    --    2. ��ת��,�ٵ������˿���Ʊ,����Ҫ���нɿ����ݴ���
    If Nvl(�����˷�_In, 0) = 1 Then
      For c_Ԥ�� In (Select a.���㷽ʽ, Sum(a.��Ԥ��) As ��Ԥ��, 2 As Ԥ�����, a.�����id, a.���㿨���, a.����, Min(a.������ˮ��) As ������ˮ��,
                          Min(a.����˵��) As ����˵��, Min(a.������λ) As ������λ, b.����
                   From ����Ԥ����¼ A, ���㷽ʽ B
                   Where a.��¼���� = 3 And a.����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And
                         a.���㷽ʽ = b.���� And b.���� In (1, 2, 7, 8) And a.���㷽ʽ Is Not Null
                   Group By a.���㷽ʽ, Ԥ�����, a.�����id, a.���㿨���, a.����, b.����
                   Having Sum(a.��Ԥ��) <> 0
                   Order By a.�����id, ���� Desc) Loop
        If n_ʵ�ս�� <> 0 Then
          Begin
            Select �Ƿ����� Into n_���� From ҽ�ƿ���� Where ID = c_Ԥ��.�����id;
          Exception
            When Others Then
              n_���� := 0;
          End;
          If (c_Ԥ��.���� = 7 Or (c_Ԥ��.���� = 8 And c_Ԥ��.�����id Is Not Null)) And n_���� = 0 Then
            If c_Ԥ��.��Ԥ�� > n_ʵ�ս�� Then
              Update ����Ԥ����¼
              Set ��Ԥ�� = ��Ԥ�� + (-1 * n_ʵ�ս��), ժҪ = ժҪ || '1' || ',' || c_Ԥ��.�����id || ',' || -1 * n_ʵ�ս�� || '|'
              Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ Is Null;
              If Sql%RowCount = 0 Then
                Insert Into ����Ԥ����¼
                  (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                   ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��������, У�Ա�־)
                Values
                  (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * n_ʵ�ս��, Null, Null, �˷�ʱ��_In,
                   Null, Null, Null, ����Ա���_In, ����Ա����_In, '1' || ',' || c_Ԥ��.�����id || ',' || -1 * n_ʵ�ս�� || '|', n_��id,
                   Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 3, 1);
              End If;
              Update ����Ԥ����¼
              Set ��¼״̬ = 3
              Where ��¼���� = 3 And ��¼״̬ = 1 And ����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And
                    ���㷽ʽ = c_Ԥ��.���㷽ʽ;
              n_����״̬ := 1;
              n_ʵ�ս�� := 0;
            Else
              Update ����Ԥ����¼
              Set ��Ԥ�� = ��Ԥ�� + (-1 * c_Ԥ��.��Ԥ��), ժҪ = ժҪ || '1' || ',' || c_Ԥ��.�����id || ',' || -1 * c_Ԥ��.��Ԥ�� || '|'
              Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ Is Null;
              If Sql%RowCount = 0 Then
                Insert Into ����Ԥ����¼
                  (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                   ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��������, У�Ա�־)
                Values
                  (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * c_Ԥ��.��Ԥ��, Null, Null, �˷�ʱ��_In,
                   Null, Null, Null, ����Ա���_In, ����Ա����_In, '1' || ',' || c_Ԥ��.�����id || ',' || -1 * c_Ԥ��.��Ԥ�� || '|', n_��id,
                   Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 3, 1);
              End If;
            
              Update ����Ԥ����¼
              Set ��¼״̬ = 3
              Where ��¼���� = 3 And ��¼״̬ = 1 And ����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And
                    ���㷽ʽ = c_Ԥ��.���㷽ʽ;
              n_����״̬ := 1;
              n_ʵ�ս�� := n_ʵ�ս�� - c_Ԥ��.��Ԥ��;
            End If;
          Else
            n_ʵ�ʳ��� := 0;
            If c_Ԥ��.���� In (3, 4) Or (c_Ԥ��.���� = 8 And c_Ԥ��.���㿨��� Is Not Null) Then
              v_���㷽ʽ := c_Ԥ��.���㷽ʽ;
            Else
              If ���㷽ʽ_In Is Null Then
                Begin
                  Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And ���� Like '%�ֽ�%' And Rownum < 2;
                Exception
                  When Others Then
                    Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And Rownum < 2;
                End;
              Else
                v_���㷽ʽ := ���㷽ʽ_In;
              End If;
            End If;
          
            If c_Ԥ��.���� = 8 And c_Ԥ��.���㿨��� Is Not Null Then
              If n_ʵ�ս�� >= c_Ԥ��.��Ԥ�� Then
                --Zl_Square_Update(v_ԭ����ids, n_����id, n_��id, �˷�ʱ��_In, -1 * n_����id, Null, c_Ԥ��.��Ԥ��, c_Ԥ��.���㿨���);
                Update ����Ԥ����¼
                Set ��Ԥ�� = ��Ԥ�� + (-1 * c_Ԥ��.��Ԥ��), ժҪ = ժҪ || '0' || ',' || c_Ԥ��.���㿨��� || ',' || -1 * c_Ԥ��.��Ԥ�� || '|'
                Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ Is Null;
                If Sql%RowCount = 0 Then
                  Insert Into ����Ԥ����¼
                    (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                     ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��������, У�Ա�־)
                  Values
                    (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * c_Ԥ��.��Ԥ��, Null, Null,
                     �˷�ʱ��_In, Null, Null, Null, ����Ա���_In, ����Ա����_In,
                     '0' || ',' || c_Ԥ��.���㿨��� || ',' || -1 * c_Ԥ��.��Ԥ�� || '|', n_��id, Null, Null, Null, Null, Null, Null,
                     n_����id, -1 * n_����id, 3, 1);
                End If;
                n_����״̬ := 1;
                n_ʵ�ʳ��� := c_Ԥ��.��Ԥ��;
              Else
                --Zl_Square_Update(v_ԭ����ids, n_����id, n_��id, �˷�ʱ��_In, -1 * n_����id, Null, n_ʵ�ս��, c_Ԥ��.���㿨���);
                Update ����Ԥ����¼
                Set ��Ԥ�� = ��Ԥ�� + (-1 * n_ʵ�ս��), ժҪ = ժҪ || '0' || ',' || c_Ԥ��.���㿨��� || ',' || -1 * n_ʵ�ս�� || '|'
                Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ Is Null;
                If Sql%RowCount = 0 Then
                  Insert Into ����Ԥ����¼
                    (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                     ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��������, У�Ա�־)
                  Values
                    (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * n_ʵ�ս��, Null, Null, �˷�ʱ��_In,
                     Null, Null, Null, ����Ա���_In, ����Ա����_In, '0' || ',' || c_Ԥ��.���㿨��� || ',' || -1 * n_ʵ�ս�� || '|', n_��id,
                     Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 3, 1);
                End If;
                n_����״̬ := 1;
                n_ʵ�ʳ��� := n_ʵ�ս��;
              End If;
            Else
              If c_Ԥ��.��Ԥ�� > n_ʵ�ս�� Then
                n_ʵ�ʳ��� := n_ʵ�ս��;
              Else
                n_ʵ�ʳ��� := c_Ԥ��.��Ԥ��;
              End If;
            End If;
          
            If c_Ԥ��.���㿨��� Is Null Then
              Update ��Ա�ɿ����
              Set ��� = Nvl(���, 0) - n_ʵ�ʳ���
              Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_���㷽ʽ
              Returning ��� Into n_����ֵ;
              If Sql%RowCount = 0 Then
                Insert Into ��Ա�ɿ����
                  (�տ�Ա, ���㷽ʽ, ����, ���)
                Values
                  (����Ա����_In, v_���㷽ʽ, 1, -1 * n_ʵ�ʳ���);
                n_����ֵ := n_ʵ�ʳ���;
              End If;
              If Nvl(n_����ֵ, 0) = 0 Then
                Delete From ��Ա�ɿ����
                Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_���㷽ʽ And Nvl(���, 0) = 0;
              End If;
            
              --��ԭԤ����¼
              Update ����Ԥ����¼
              Set ��Ԥ�� = ��Ԥ�� + (-1 * n_ʵ�ʳ���)
              Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_���㷽ʽ;
              If Sql%RowCount = 0 Then
                Insert Into ����Ԥ����¼
                  (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                   ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, У�Ա�־, ��������)
                Values
                  (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * n_ʵ�ʳ���, v_���㷽ʽ, Null, �˷�ʱ��_In,
                   Null, Null, Null, ����Ա���_In, ����Ա����_In, '', n_��id, Null, Null, Null, Null, Null, c_Ԥ��.������λ, n_����id,
                   -1 * n_����id, 0, 3);
              End If;
            End If;
            Update ����Ԥ����¼
            Set ��¼״̬ = 3
            Where ��¼���� = 3 And ��¼״̬ = 1 And ����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And
                  ���㷽ʽ = c_Ԥ��.���㷽ʽ;
            n_ʵ�ս�� := n_ʵ�ս�� - n_ʵ�ʳ���;
          End If;
        End If;
      End Loop;
    
      --���·�����˼�¼
      Update ������˼�¼
      Set ��¼״̬ = 2
      Where ����id In (Select ID From ������ü�¼ Where NO = No_In And ��¼���� = 1 And ��¼״̬ In (1, 3)) And ���� = 1;
      --���������¼
      Update ������ü�¼ Set ��¼״̬ = 3 Where NO = No_In And Mod(��¼����, 10) = 1 And ��¼״̬ = 1;
      For r_Clinic In (Select ���, ��������, �۸񸸺�, ����id, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ���ձ���, ��������,
                              ��ҩ����, ����, Sum(����) As ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, Sum(Ӧ�ս��) As Ӧ�ս��,
                              Sum(ʵ�ս��) As ʵ�ս��, Sum(ͳ����) As ͳ����, ��������id, ������, ִ�в���id, ������, Max(���ʵ�id) As ���ʵ�id, ����ʱ��,
                              ʵ��Ʊ��
                       From ������ü�¼
                       Where NO = No_In And Mod(��¼����, 10) = 1 And ��¼״̬ In (2, 3) And Nvl(���ӱ�־, 0) Not In (8, 9)
                       Group By ���, ��������, �۸񸸺�, ����id, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ���ձ���,
                                ��������, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, ��������id, ������, ִ�в���id, ������, ����ʱ��, ʵ��Ʊ��
                       Having Sum(����) <> 0) Loop
        Insert Into ������ü�¼
          (ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, �����־, ����id, ��ʶ��, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��,
           ���մ���id, ���ձ���, ��������, ��ҩ����, ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ͳ����, ���ʷ���, ��������id, ������, ����ʱ��,
           �Ǽ�ʱ��, ִ�в���id, ������, ����Ա���, ����Ա����, ���ʵ�id, ժҪ, �ɿ���id, ����id, ���ʽ��, ����״̬)
        Values
          (���˷��ü�¼_Id.Nextval, 1, No_In, r_Clinic.ʵ��Ʊ��, 2, r_Clinic.���, r_Clinic.��������, r_Clinic.�۸񸸺�, 1, r_Clinic.����id,
           '', r_Clinic.����, r_Clinic.�Ա�, r_Clinic.����, r_Clinic.���˿���id, r_Clinic.�ѱ�, r_Clinic.�շ����, r_Clinic.�շ�ϸĿid,
           r_Clinic.���㵥λ, r_Clinic.������Ŀ��, r_Clinic.���մ���id, r_Clinic.���ձ���, r_Clinic.��������, r_Clinic.��ҩ����, r_Clinic.����,
           -1 * r_Clinic.����, r_Clinic.�Ӱ��־, r_Clinic.���ӱ�־, r_Clinic.������Ŀid, r_Clinic.�վݷ�Ŀ, r_Clinic.��׼����,
           -1 * r_Clinic.Ӧ�ս��, -1 * r_Clinic.ʵ�ս��, -1 * r_Clinic.ͳ����, 0, r_Clinic.��������id, r_Clinic.������, r_Clinic.����ʱ��,
           �˷�ʱ��_In, r_Clinic.ִ�в���id, r_Clinic.������, ����Ա���_In, ����Ա����_In, r_Clinic.���ʵ�id, '', n_��id, n_����id,
           -1 * r_Clinic.ʵ�ս��, 0);
      End Loop;
    Else
      --4.�˿�תԤ��(������Ʊ��,�ɲ���Աͨ���ش����)
      For r_Pay In (Select Min(a.Id) As Ԥ��id, a.���㷽ʽ, Sum(a.��Ԥ��) As ��Ԥ��, 2 As Ԥ�����, a.�����id, a.���㿨���, a.����, a.������ˮ��,
                           a.����˵��, a.������λ, b.����
                    From ����Ԥ����¼ A, ���㷽ʽ B
                    Where a.��¼���� = 3 And a.����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And
                          a.���㷽ʽ = b.���� And (b.���� In (1, 2, 7, 8)) And a.���㷽ʽ Is Not Null
                    Group By a.���㷽ʽ, Ԥ�����, a.�����id, a.���㿨���, a.����, b.����, a.������ˮ��, a.����˵��, a.������λ

                    
                    Having Sum(a.��Ԥ��) <> 0
                    Order By a.�����id, ���� Desc) Loop
        --4.1����Ԥ����� (�����ڲ����˷ѵ����)
        --���е���,����������Ԥ�����
        --��Ϊ�տ�������ɿ�,������Ա�ɿ�����ޱ仯
        If n_ʵ�ս�� <> 0 Then
          If r_Pay.���� = 7 Or (r_Pay.���� = 8 And r_Pay.�����id Is Not Null) Then
            If r_Pay.��Ԥ�� > n_ʵ�ս�� Then
              Update ����Ԥ����¼
              Set ��Ԥ�� = ��Ԥ�� + (-1 * n_ʵ�ս��), ժҪ = ժҪ || '1' || ',' || r_Pay.�����id || ',' || -1 * n_ʵ�ս�� || '|'
              Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ Is Null;
              If Sql%RowCount = 0 Then
                Insert Into ����Ԥ����¼
                  (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                   ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��������, У�Ա�־)
                Values
                  (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * n_ʵ�ս��, Null, Null, �˷�ʱ��_In,
                   Null, Null, Null, ����Ա���_In, ����Ա����_In, '1' || ',' || r_Pay.�����id || ',' || -1 * n_ʵ�ս�� || '|', n_��id,
                   Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 3, 1);
              End If;
            
              Update ����Ԥ����¼
              Set ��¼״̬ = 3
              Where ��¼���� = 3 And ��¼״̬ = 1 And ����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And
                    ���㷽ʽ = r_Pay.���㷽ʽ;
              n_����״̬ := 1;
              n_ʵ�ս�� := 0;
            Else
              Update ����Ԥ����¼
              Set ��Ԥ�� = ��Ԥ�� + (-1 * r_Pay.��Ԥ��), ժҪ = ժҪ || '1' || ',' || r_Pay.�����id || ',' || -1 * r_Pay.��Ԥ�� || '|'
              Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ Is Null;
              If Sql%RowCount = 0 Then
                Insert Into ����Ԥ����¼
                  (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                   ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��������, У�Ա�־)
                Values
                  (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * r_Pay.��Ԥ��, Null, Null, �˷�ʱ��_In,
                   Null, Null, Null, ����Ա���_In, ����Ա����_In, '1' || ',' || r_Pay.�����id || ',' || -1 * r_Pay.��Ԥ�� || '|',
                   n_��id, Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 3, 1);
              End If;
            
              Update ����Ԥ����¼
              Set ��¼״̬ = 3
              Where ��¼���� = 3 And ��¼״̬ = 1 And ����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And
                    ���㷽ʽ = r_Pay.���㷽ʽ;
              n_����״̬ := 1;
              n_ʵ�ս�� := n_ʵ�ս�� - r_Pay.��Ԥ��;
            End If;
          Else
            n_ʵ�ʳ��� := 0;
            If r_Pay.���� In (3, 4) Or (r_Pay.���� = 8 And r_Pay.���㿨��� Is Not Null) Then
              v_���㷽ʽ := r_Pay.���㷽ʽ;
            Else
              If ���㷽ʽ_In Is Null Then
                Begin
                  Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And ���� Like '%�ֽ�%' And Rownum < 2;
                Exception
                  When Others Then
                    Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And Rownum < 2;
                End;
              Else
                v_���㷽ʽ := ���㷽ʽ_In;
              End If;
            End If;
          
            If r_Pay.���� = 8 And r_Pay.���㿨��� Is Not Null Then
              If n_ʵ�ս�� >= r_Pay.��Ԥ�� Then
                --Zl_Square_Update(v_ԭ����ids, n_����id, n_��id, �˷�ʱ��_In, -1 * n_����id, Null, r_Pay.��Ԥ��, r_Pay.���㿨���);
                Update ����Ԥ����¼
                Set ��Ԥ�� = ��Ԥ�� + (-1 * r_Pay.��Ԥ��), ժҪ = ժҪ || '0' || ',' || r_Pay.���㿨��� || ',' || -1 * r_Pay.��Ԥ�� || '|'
                Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ Is Null;
                If Sql%RowCount = 0 Then
                  Insert Into ����Ԥ����¼
                    (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                     ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��������, У�Ա�־)
                  Values
                    (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * r_Pay.��Ԥ��, Null, Null,
                     �˷�ʱ��_In, Null, Null, Null, ����Ա���_In, ����Ա����_In,
                     '0' || ',' || r_Pay.���㿨��� || ',' || -1 * r_Pay.��Ԥ�� || '|', n_��id, Null, Null, Null, Null, Null,
                     Null, n_����id, -1 * n_����id, 3, 1);
                End If;
                n_����״̬ := 1;
                n_ʵ�ʳ��� := r_Pay.��Ԥ��;
              Else
                --Zl_Square_Update(v_ԭ����ids, n_����id, n_��id, �˷�ʱ��_In, -1 * n_����id, Null, n_ʵ�ս��, r_Pay.���㿨���);
                Update ����Ԥ����¼
                Set ��Ԥ�� = ��Ԥ�� + (-1 * n_ʵ�ս��), ժҪ = ժҪ || '0' || ',' || r_Pay.���㿨��� || ',' || -1 * n_ʵ�ս�� || '|'
                Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ Is Null;
                If Sql%RowCount = 0 Then
                  Insert Into ����Ԥ����¼
                    (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                     ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��������, У�Ա�־)
                  Values
                    (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * n_ʵ�ս��, Null, Null, �˷�ʱ��_In,
                     Null, Null, Null, ����Ա���_In, ����Ա����_In, '0' || ',' || r_Pay.���㿨��� || ',' || -1 * n_ʵ�ս�� || '|', n_��id,
                     Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 3, 1);
                End If;
                n_����״̬ := 1;
                n_ʵ�ʳ��� := n_ʵ�ս��;
              End If;
            Else
              If r_Pay.��Ԥ�� > n_ʵ�ս�� Then
                n_ʵ�ʳ��� := n_ʵ�ս��;
              Else
                n_ʵ�ʳ��� := r_Pay.��Ԥ��;
              End If;
            End If;
          
            If r_Pay.���� Not In (3, 4, 7, 8) Then
              Update ����Ԥ����¼
              Set ��� = ��� + n_ʵ�ʳ���
              Where ��¼���� = 1 And ��¼״̬ = 1 And �տ�ʱ�� = �˷�ʱ��_In And ����id + 0 = n_����id And ���㷽ʽ = v_���㷽ʽ;
              If Sql%RowCount = 0 Then
                v_Ԥ��no := Nextno(11);
                Insert Into ����Ԥ����¼
                  (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                   ժҪ, �ɿ���id, Ԥ�����)
                Values
                  (����Ԥ����¼_Id.Nextval, v_Ԥ��no, Null, 1, 1, n_����id, ��ҳid_In, ��Ժ����id_In, n_ʵ�ʳ���, v_���㷽ʽ, Null, �˷�ʱ��_In,
                   Null, Null, Null, ����Ա���_In, ����Ա����_In, '����תסԺԤ��', n_��id, r_Pay.Ԥ�����);
              End If;
            
              --�������
              Update �������
              Set Ԥ����� = Nvl(Ԥ�����, 0) + n_ʵ�ʳ���
              Where ���� = 1 And ����id = n_����id And ���� = 2
              Returning Ԥ����� Into n_����ֵ;
              If Sql%RowCount = 0 Then
                Insert Into ������� (����id, ����, ����, Ԥ�����, �������) Values (n_����id, 1, 2, n_ʵ�ʳ���, 0);
                n_����ֵ := n_ʵ�ʳ���;
              End If;
              If Nvl(n_����ֵ, 0) = 0 Then
                Delete From �������
                Where ����id = n_����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
              End If;
            End If;
            --4.2�ɿ����ݴ���
            --   ��Ϊû��ʵ���ղ��˵�Ǯ,���Բ�����
            --�����˷��������ԭԤ����¼
            If r_Pay.���� In (3, 4) Then
              Update ��Ա�ɿ����
              Set ��� = Nvl(���, 0) - n_ʵ�ʳ���
              Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = r_Pay.���㷽ʽ
              Returning ��� Into n_����ֵ;
              If Sql%RowCount = 0 Then
                Insert Into ��Ա�ɿ����
                  (�տ�Ա, ���㷽ʽ, ����, ���)
                Values
                  (����Ա����_In, r_Pay.���㷽ʽ, 1, -1 * n_ʵ�ʳ���);
                n_����ֵ := n_ʵ�ʳ���;
              End If;
              If Nvl(n_����ֵ, 0) = 0 Then
                Delete From ��Ա�ɿ����
                Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = r_Pay.���㷽ʽ And Nvl(���, 0) = 0;
              End If;
            End If;
          
            If r_Pay.���� <> 8 Then
              Update ����Ԥ����¼
              Set ��Ԥ�� = ��Ԥ�� + (-1 * n_ʵ�ʳ���)
              Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_���㷽ʽ;
              If Sql%RowCount = 0 Then
                Insert Into ����Ԥ����¼
                  (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                   ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, У�Ա�־, ��������)
                Values
                  (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * n_ʵ�ʳ���, v_���㷽ʽ, Null, �˷�ʱ��_In,
                   Null, Null, Null, ����Ա���_In, ����Ա����_In, '', n_��id, r_Pay.�����id, r_Pay.���㿨���, r_Pay.����, r_Pay.������ˮ��,
                   r_Pay.����˵��, r_Pay.������λ, n_����id, -1 * n_����id, 0, 3);
              End If;
            End If;
          
            Update ����Ԥ����¼
            Set ��¼״̬ = 3
            Where ��¼���� = 3 And ��¼״̬ = 1 And ����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And
                  ���㷽ʽ = r_Pay.���㷽ʽ;
            n_ʵ�ս�� := n_ʵ�ս�� - n_ʵ�ʳ���;
          
          End If;
        End If;
      End Loop;
    End If;
  
    If ����_In Is Not Null Then
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ, �ɿ���id,
         �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, У�Ա�־, ��������)
      Values
        (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, ����_In, v_����, Null, �˷�ʱ��_In, Null, Null,
         Null, ����Ա���_In, ����Ա����_In, '', n_��id, Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 0, 3);
    End If;
    Delete From ����Ԥ����¼
    Where ����id = n_����id And ��¼���� = 3 And ��¼״̬ = 2 And ��Ԥ�� = 0 And ���㷽ʽ Is Not Null;
    Delete From ����Ԥ����¼ Where ����id = n_ԭ����id And ժҪ = 'Ԥ����ʱ��¼' And ��¼���� = 3;
    Update ������ü�¼ Set ����״̬ = Nvl(n_����״̬, 0) Where NO = No_In And Mod(��¼����, 10) = 1 And ��¼״̬ = 2;
  Else
    --ҽ��������ת��
    For r_Nos In (Select Distinct a.No
                  From ������ü�¼ A
                  Where Mod(a.��¼����, 10) = 1 And a.��¼״̬ In (1, 3) And a.����id = ԭ����id_In) Loop
      v_Nos := v_Nos || ',' || r_Nos.No;
    End Loop;
    v_Nos := Substr(v_Nos, 2);
  
    For r_����ids In (Select Distinct a.����id
                    From ������ü�¼ A
                    Where a.No In (Select Column_Value From Table(f_Str2list(v_Nos))) And Mod(a.��¼����, 10) = 1 And
                          a.��¼״̬ <> 0) Loop
      v_����ids := v_����ids || ',' || r_����ids.����id;
    End Loop;
    v_����ids := Substr(v_����ids, 2);
    Select Count(a.No), Sum(a.ʵ�ս��)
    Into n_Count, n_ʵ�ս��
    From ������ü�¼ A
    Where a.No In (Select Column_Value From Table(f_Str2list(v_Nos))) And Mod(a.��¼����, 10) = 1;
    If n_Count = 0 Or n_ʵ�ս�� = 0 Then
      v_Err_Msg := '���ν��㲻���շѻ��򲢷�ԭ�����˲����˸ý���,����תΪסԺ����.';
      Raise Err_Item;
    End If;
  
    Select ����id, ����id, ��������id, ������
    Into n_ԭ����id, n_����id, n_��������id, v_������
    From ������ü�¼
    Where ����id = ԭ����id_In And Mod(��¼����, 10) = 1 And ��¼״̬ In (1, 3) And Rownum < 2;
  
    Begin
      Select 1
      Into n_�����˷�
      From ������ü�¼ A
      Where Mod(a.��¼����, 10) = 1 And a.��¼״̬ = 2 And a.����id In (Select Column_Value From Table(f_Str2list(v_����ids))) And
            Rownum < 2;
    Exception
      When Others Then
        n_�����˷� := 0;
    End;
  
    Begin
      Select 0
      Into n_�����˷�
      From ������ü�¼ A
      Where ��¼���� = 11 And a.����id In (Select Column_Value From Table(f_Str2list(v_����ids))) And Rownum < 2;
    Exception
      When Others Then
        Null;
    End;
    Begin
      Select Count(Avg(1))
      Into n_�˷�����
      From ����Ԥ����¼ A
      Where a.��¼���� = 3 And a.��¼״̬ <> 0 And ����id In (Select Column_Value From Table(f_Str2list(v_����ids)))
      Group By a.���㷽ʽ;
    Exception
      When Others Then
        n_�˷����� := 0;
    End;
    --1.1���Ϸ��ü�¼
    If ����id_In Is Null Then
      Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
    Else
      n_����id := ����id_In;
    End If;
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�, ���˿���id, �շ����, �շ�ϸĿid,
       ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ������, ִ����, ִ��״̬, ִ��ʱ��,
       ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id, ����״̬)
      Select ���˷��ü�¼_Id.Nextval, a.No, a.ʵ��Ʊ��, a.��¼����, 2, a.���, a.��������, a.�۸񸸺�, a.����id, a.ҽ�����, a.�����־, a.����, a.�Ա�, a.����,
             a.��ʶ��, a.���ʽ, a.�ѱ�, a.���˿���id, a.�շ����, a.�շ�ϸĿid, a.���㵥λ, a.����, a.��ҩ����, -1 * a.����, a.�Ӱ��־, a.���ӱ�־, a.������Ŀid,
             a.�վݷ�Ŀ, a.���ʷ���, a.��׼����, -1 * a.Ӧ�ս��, -1 * a.ʵ�ս��, a.��������id, a.������, a.ִ�в���id, a.������, a.ִ����, -1, a.ִ��ʱ��,
             ����Ա���_In, ����Ա����_In, a.����ʱ��, �˷�ʱ��_In, n_����id, -1 * a.���ʽ��, a.������Ŀ��, a.���մ���id, a.ͳ����, a.ժҪ,
             Decode(Nvl(a.���ӱ�־, 0), 9, 1, 0), a.���ձ���, a.��������, n_��id, 0
      From ������ü�¼ A
      Where a.No In (Select Column_Value From Table(f_Str2list(v_Nos))) And Mod(a.��¼����, 10) = 1 And a.��¼״̬ = 1;
  
    --����ҽ��
    For r_ҽ�� In (Select ����id, NO, ���㷽ʽ, ���, ��ע
                 From ҽ��������ϸ
                 Where NO In (Select Column_Value From Table(f_Str2list(v_Nos))) And
                       ����id In (Select Column_Value From Table(f_Str2list(v_����ids)))) Loop
      Update ҽ��������ϸ
      Set ��� = ��� + (-1 * r_ҽ��.���)
      Where NO = r_ҽ��.No And ����id = r_ҽ��.����id And ���㷽ʽ = r_ҽ��.���㷽ʽ;
      If Sql%RowCount = 0 Then
        Insert Into ҽ��������ϸ
          (����id, NO, ���㷽ʽ, ���)
        Values
          (r_ҽ��.����id, r_ҽ��.No, r_ҽ��.���㷽ʽ, -1 * r_ҽ��.���);
      End If;
    End Loop;
  
    --Update ������ü�¼ Set ��¼״̬ = 3 Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 1;
    --1.2����Ԥ����¼
    --���ϳ�Ԥ������
    If n_�����˷� = 0 And Nvl(�����˷�_In, 0) = 0 Then
      For r_Prepay In (Select NO, Max(Decode(��¼����, 1, ʵ��Ʊ��, Null)) As ʵ��Ʊ��, ����id, ��ҳid, ����id, ���㷽ʽ, �������, �ɿλ, ��λ������,
                              ��λ�ʺ�, �տ�ʱ��, -1 * Sum(��Ԥ��) As ��Ԥ��, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������
                       From ����Ԥ����¼ A
                       Where ��¼���� In (1, 11) And a.����id In (Select Column_Value From Table(f_Str2list(v_����ids))) And
                             Nvl(��Ԥ��, 0) <> 0
                       Group By NO, ����id, ��ҳid, ����id, ���㷽ʽ, �������, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, �����id, ���㿨���, ����, ������ˮ��, ����˵��,
                                ������λ, ��������) Loop
        Select ����Ԥ����¼_Id.Nextval Into n_Tempid From Dual;
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
           ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, Ԥ�����, ��������)
          Select n_Tempid, r_Prepay.No, r_Prepay.ʵ��Ʊ��, 11, 1, r_Prepay.����id, r_Prepay.��ҳid, r_Prepay.����id, Null,
                 r_Prepay.���㷽ʽ, r_Prepay.�������, Null, r_Prepay.�ɿλ, r_Prepay.��λ������, r_Prepay.��λ�ʺ�, �˷�ʱ��_In, ����Ա����_In,
                 ����Ա���_In, r_Prepay.��Ԥ��, n_����id, n_��id, r_Prepay.�����id, r_Prepay.���㿨���, r_Prepay.����, r_Prepay.������ˮ��,
                 r_Prepay.����˵��, r_Prepay.������λ, -1 * n_����id, 1, r_Prepay.��������
          From Dual;
      End Loop;
    
      For v_Ԥ�� In (Select ����id, Nvl(Ԥ�����, 2) As Ԥ�����, Nvl(Sum(Nvl(��Ԥ��, 0)), 0) As Ԥ�����
                   From ����Ԥ����¼ A
                   Where ��¼���� In (1, 11) And a.����id In (Select Column_Value From Table(f_Str2list(v_����ids))) And
                         a.����id <> n_����id
                   Group By ����id, Nvl(Ԥ�����, 2)
                   Having Sum(Nvl(��Ԥ��, 0)) <> 0) Loop
      
        Update �������
        Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(v_Ԥ��.Ԥ�����, 0)
        Where ����id = v_Ԥ��.����id And ���� = Nvl(v_Ԥ��.Ԥ�����, 2) And ���� = 1
        Returning Ԥ����� Into n_����ֵ;
        If Sql%RowCount = 0 Then
          Insert Into �������
            (����id, ����, Ԥ�����, ����)
          Values
            (v_Ԥ��.����id, Nvl(v_Ԥ��.Ԥ�����, 2), v_Ԥ��.Ԥ�����, 1);
          n_����ֵ := v_Ԥ��.Ԥ�����;
        End If;
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete From �������
          Where ����id = v_Ԥ��.����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
        End If;
      End Loop;
    Else
      If n_�˷����� = 0 And Nvl(�����˷�_In, 0) = 0 Then
        --ֻʹ����Ԥ����ԭ���˻�Ԥ��
        For r_Prepay In (Select NO, Max(Decode(��¼����, 1, ʵ��Ʊ��, Null)) As ʵ��Ʊ��, ����id, ��ҳid, ����id, Max(���㷽ʽ) As ���㷽ʽ, �������,
                                �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, -1 * Sum(��Ԥ��) As ��Ԥ��, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������
                         From ����Ԥ����¼ A
                         Where ��¼���� In (1, 11) And a.����id In (Select Column_Value From Table(f_Str2list(v_����ids))) And
                               Nvl(��Ԥ��, 0) <> 0
                         Group By NO, ����id, ��ҳid, ����id, �������, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, �����id, ���㿨���, ����, ������ˮ��, ����˵��,
                                  ������λ, ��������) Loop
          Select ����Ԥ����¼_Id.Nextval Into n_Tempid From Dual;
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
             ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, Ԥ�����, ��������)
            Select n_Tempid, r_Prepay.No, r_Prepay.ʵ��Ʊ��, 11, 1, r_Prepay.����id, r_Prepay.��ҳid, r_Prepay.����id, Null,
                   r_Prepay.���㷽ʽ, r_Prepay.�������, Null, r_Prepay.�ɿλ, r_Prepay.��λ������, r_Prepay.��λ�ʺ�, �˷�ʱ��_In, ����Ա����_In,
                   ����Ա���_In, r_Prepay.��Ԥ��, n_����id, n_��id, r_Prepay.�����id, r_Prepay.���㿨���, r_Prepay.����, r_Prepay.������ˮ��,
                   r_Prepay.����˵��, r_Prepay.������λ, -1 * n_����id, 1, r_Prepay.��������
            From Dual;
        
          Select -1 * ��Ԥ�� Into n_Ԥ����� From ����Ԥ����¼ Where ID = n_Tempid;
          Update �������
          Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(n_Ԥ�����, 0)
          Where ����id = r_Prepay.����id And ���� = 1 And ���� = 1
          Returning Ԥ����� Into n_����ֵ;
          If Sql%RowCount = 0 Then
            Insert Into ������� (����id, ����, Ԥ�����, ����) Values (n_����id, 1, n_Ԥ�����, 1);
            n_����ֵ := n_Ԥ�����;
          End If;
          If Nvl(n_����ֵ, 0) = 0 Then
            Delete From �������
            Where ����id = r_Prepay.����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
          End If;
        End Loop;
      Else
        Begin
          Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And ���� Like '%�ֽ�%' And Rownum < 2;
        Exception
          When Others Then
            Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And Rownum < 2;
        End;
        Select ����Ԥ����¼_Id.Nextval Into n_Tempid From Dual;
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
           ����id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, ��������)
          Select n_Tempid, Max(NO), Max(ʵ��Ʊ��), 3, 3, ����id, ��ҳid, ����id, Null, v_���㷽ʽ, Max(�������), 'Ԥ����ʱ��¼', Null, Null,
                 Null, Max(�տ�ʱ��), ����Ա����_In, ����Ա���_In, Sum(��Ԥ��), n_ԭ����id, Null, Null, Null, Null, Null, Null,
                 -1 * n_ԭ����id, 3
          From ����Ԥ����¼ A
          Where ��¼���� In (1, 11) And a.����id In (Select Column_Value From Table(f_Str2list(v_����ids))) And
                Nvl(��Ԥ��, 0) <> 0
          Group By n_Tempid, 3, 3, ����id, ��ҳid, ����id, Null, v_���㷽ʽ, 'Ԥ����ʱ��¼', ����Ա����_In, ����Ա���_In, n_ԭ����id;
      End If;
    End If;
  
    --��������ɷѼ�ҽ������
    Insert Into ����Ԥ����¼
      (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����,
       �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, ��������)
      Select ����Ԥ����¼_Id.Nextval, ��¼����, NO, 2, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �˷�ʱ��_In, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���_In, ����Ա����_In,
             0, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, -1 * n_����id, ��������
      From ����Ԥ����¼ A, ���㷽ʽ B
      Where a.��¼���� = 3 And a.��¼״̬ = 1 And a.����id In (Select Column_Value From Table(f_Str2list(v_����ids))) And
            a.���㷽ʽ = b.���� And b.���� Not In (7, 8);
  
    Insert Into ����Ԥ����¼
      (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����,
       �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, ��������, У�Ա�־)
      Select ����Ԥ����¼_Id.Nextval, ��¼����, NO, 2, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �˷�ʱ��_In, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���_In, ����Ա����_In,
             0, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, -1 * n_����id, ��������, 1
      From ����Ԥ����¼ A, ���㷽ʽ B
      Where a.��¼���� = 3 And a.��¼״̬ = 1 And a.����id In (Select Column_Value From Table(f_Str2list(v_����ids))) And
            a.���㷽ʽ = b.���� And b.���� = 7;
    If Sql%RowCount <> 0 Then
      n_����״̬ := 1;
    End If;
  
    Update ����Ԥ����¼
    Set ��¼״̬ = 3
    Where ��¼���� = 3 And ��¼״̬ = 1 And ����id In (Select Column_Value From Table(f_Str2list(v_����ids)));
  
    --2.Ʊ���ջ�
    --������ǰû�д�ӡ,���ջ�
    For r_Nos In (Select Distinct a.No
                  From ������ü�¼ A
                  Where Mod(a.��¼����, 10) = 1 And a.��¼״̬ In (1, 3) And
                        a.����id In (Select Column_Value From Table(f_Str2list(v_����ids)))) Loop
    
      Select Nvl(Max(ID), 0)
      Into n_��ӡid
      From (Select b.Id
             From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
             Where a.��ӡid = b.Id And a.���� = 1 And a.ԭ�� In (1, 3) And b.�������� = 1 And b.No = r_Nos.No
             Order By a.ʹ��ʱ�� Desc)
      Where Rownum < 2;
      If n_��ӡid > 0 Then
        --���ŵ���ѭ������ʱֻ���ջ�һ��
        Select Count(��ӡid) Into n_Count From Ʊ��ʹ����ϸ Where Ʊ�� = 1 And ���� = 2 And ��ӡid = n_��ӡid;
        If n_Count = 0 Then
          Insert Into Ʊ��ʹ����ϸ
            (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����)
            Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, �˷�ʱ��_In, ����Ա����_In
            From Ʊ��ʹ����ϸ
            Where ��ӡid = n_��ӡid And Ʊ�� = 1 And ���� = 1;
        End If;
      End If;
    End Loop;
  
    --3.�ɿ����ݴ���(
    --   �����������:
    --    1. ת������ֱ�����ʵ�,��ɿ����ݲ�����;
    --    2. ��ת��,�ٵ������˿���Ʊ,����Ҫ���нɿ����ݴ���
    If Nvl(�����˷�_In, 0) = 1 Then
      For c_Ԥ�� In (Select a.���㷽ʽ, Sum(a.��Ԥ��) As ��Ԥ��, 2 As Ԥ�����, a.�����id, a.���㿨���, a.����, Min(a.������ˮ��) As ������ˮ��,
                          Min(a.����˵��) As ����˵��, Min(a.������λ) As ������λ, b.����
                   From ����Ԥ����¼ A, ���㷽ʽ B
                   Where a.��¼���� = 3 And a.��¼״̬ In (2, 3) And
                         a.����id In (Select Column_Value From Table(f_Str2list(v_����ids))) And a.���㷽ʽ = b.���� And
                         b.���� In (1, 2, 3, 4, 7, 8) And a.���㷽ʽ Is Not Null
                   Group By a.���㷽ʽ, Ԥ�����, a.�����id, a.���㿨���, a.����, b.����
                   Having Sum(a.��Ԥ��) <> 0) Loop
        Begin
          Select �Ƿ����� Into n_���� From ҽ�ƿ���� Where ID = c_Ԥ��.�����id;
        Exception
          When Others Then
            n_���� := 0;
        End;
        If (c_Ԥ��.���� = 7 Or (c_Ԥ��.���� = 8 And c_Ԥ��.�����id Is Not Null)) And n_���� = 0 Then
          Update ����Ԥ����¼
          Set ��Ԥ�� = ��Ԥ�� + (-1 * c_Ԥ��.��Ԥ��), ժҪ = ժҪ || '1' || ',' || c_Ԥ��.�����id || ',' || -1 * c_Ԥ��.��Ԥ�� || '|'
          Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ Is Null;
          If Sql%RowCount = 0 Then
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ,
               �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��������, У�Ա�־)
            Values
              (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * c_Ԥ��.��Ԥ��, Null, Null, �˷�ʱ��_In,
               Null, Null, Null, ����Ա���_In, ����Ա����_In, '1' || ',' || c_Ԥ��.�����id || ',' || -1 * c_Ԥ��.��Ԥ�� || '|', n_��id,
               Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 3, 1);
          End If;
          n_����״̬ := 1;
        Else
          If c_Ԥ��.���� In (3, 4) Or (c_Ԥ��.���� = 8 And c_Ԥ��.���㿨��� Is Not Null) Then
            v_���㷽ʽ := c_Ԥ��.���㷽ʽ;
          Else
            If ���㷽ʽ_In Is Null Then
              Begin
                Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And ���� Like '%�ֽ�%' And Rownum < 2;
              Exception
                When Others Then
                  Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And Rownum < 2;
              End;
            Else
              v_���㷽ʽ := ���㷽ʽ_In;
            End If;
          End If;
        
          If c_Ԥ��.���� = 8 And c_Ԥ��.���㿨��� Is Not Null Then
            --Zl_Square_Update(v_����ids, n_����id, n_��id, �˷�ʱ��_In, -1 * n_����id, Null, c_Ԥ��.��Ԥ��, c_Ԥ��.���㿨���);
            Update ����Ԥ����¼
            Set ��Ԥ�� = ��Ԥ�� + (-1 * c_Ԥ��.��Ԥ��), ժҪ = ժҪ || '0' || ',' || c_Ԥ��.���㿨��� || ',' || -1 * c_Ԥ��.��Ԥ�� || '|'
            Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ Is Null;
            If Sql%RowCount = 0 Then
              Insert Into ����Ԥ����¼
                (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ,
                 �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��������, У�Ա�־)
              Values
                (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * c_Ԥ��.��Ԥ��, Null, Null, �˷�ʱ��_In,
                 Null, Null, Null, ����Ա���_In, ����Ա����_In, '0' || ',' || c_Ԥ��.���㿨��� || ',' || -1 * c_Ԥ��.��Ԥ�� || '|', n_��id,
                 Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 3, 1);
            End If;
            n_����״̬ := 1;
          End If;
          If c_Ԥ��.���㿨��� Is Null Then
            Update ��Ա�ɿ����
            Set ��� = Nvl(���, 0) - c_Ԥ��.��Ԥ��
            Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_���㷽ʽ
            Returning ��� Into n_����ֵ;
            If Sql%RowCount = 0 Then
              Insert Into ��Ա�ɿ����
                (�տ�Ա, ���㷽ʽ, ����, ���)
              Values
                (����Ա����_In, v_���㷽ʽ, 1, -1 * c_Ԥ��.��Ԥ��);
              n_����ֵ := c_Ԥ��.��Ԥ��;
            End If;
            If Nvl(n_����ֵ, 0) = 0 Then
              Delete From ��Ա�ɿ����
              Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_���㷽ʽ And Nvl(���, 0) = 0;
            End If;
            --�����˷��������ԭԤ����¼
            Update ����Ԥ����¼
            Set ��Ԥ�� = ��Ԥ�� + (-1 * c_Ԥ��.��Ԥ��)
            Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_���㷽ʽ;
            If Sql%RowCount = 0 Then
              Insert Into ����Ԥ����¼
                (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ,
                 �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, У�Ա�־, ��������)
              Values
                (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * c_Ԥ��.��Ԥ��, v_���㷽ʽ, Null, �˷�ʱ��_In,
                 Null, Null, Null, ����Ա���_In, ����Ա����_In, '', n_��id, Null, Null, Null, Null, Null, c_Ԥ��.������λ, n_����id,
                 -1 * n_����id, 0, 3);
            End If;
          End If;
        End If;
      End Loop;
    
      --���·�����˼�¼
      Update ������˼�¼
      Set ��¼״̬ = 2
      Where ����id In (Select a.Id
                     From ������ü�¼ A
                     Where a.No In (Select Column_Value From Table(f_Str2list(v_Nos))) And Mod(a.��¼����, 10) = 1 And
                           a.��¼״̬ In (1, 3)) And ���� = 1;
      --���������¼
      For r_Nos In (Select Distinct NO
                    From ������ü�¼
                    Where Mod(��¼����, 10) = 1 And ��¼״̬ In (1, 3) And
                          ����id In (Select Column_Value From Table(f_Str2list(v_����ids)))) Loop
        Update ������ü�¼ Set ��¼״̬ = 3 Where NO = r_Nos.No And Mod(��¼����, 10) = 1 And ��¼״̬ = 1;
      End Loop;
      For r_Clinic In (Select Min(a.��¼����) As ��¼����, a.No, a.���, a.��������, a.�۸񸸺�, a.����id, a.����, a.�Ա�, a.����, a.���˿���id, a.�ѱ�,
                              a.�շ����, a.�շ�ϸĿid, a.���㵥λ, a.������Ŀ��, a.���մ���id, a.���ձ���, a.��������, a.��ҩ����, a.����, Sum(a.����) As ����,
                              a.�Ӱ��־, a.���ӱ�־, a.������Ŀid, a.�վݷ�Ŀ, a.��׼����, Sum(a.Ӧ�ս��) As Ӧ�ս��, Sum(a.ʵ�ս��) As ʵ�ս��,
                              Sum(a.ͳ����) As ͳ����, a.��������id, a.������, a.ִ�в���id, a.������, Max(a.���ʵ�id) As ���ʵ�id,
                              Max(a.�Ƿ���) As �Ƿ���, a.����ʱ��, Min(a.ʵ��Ʊ��) As ʵ��Ʊ��
                       From ������ü�¼ A
                       Where a.No In (Select Column_Value From Table(f_Str2list(v_Nos))) And Mod(a.��¼����, 10) = 1 And
                             a.��¼״̬ In (2, 3) And Nvl(a.���ӱ�־, 0) Not In (8, 9)
                       Group By a.No, a.���, a.��������, a.�۸񸸺�, a.����id, a.����, a.�Ա�, a.����, a.���˿���id, a.�ѱ�, a.�շ����, a.�շ�ϸĿid,
                                a.���㵥λ, a.������Ŀ��, a.���մ���id, a.���ձ���, a.��������, a.��ҩ����, a.����, a.�Ӱ��־, a.���ӱ�־, a.������Ŀid, a.�վݷ�Ŀ,
                                a.��׼����, a.��������id, a.������, a.ִ�в���id, a.������, a.����ʱ��
                       Having Sum(a.����) <> 0) Loop
        Insert Into ������ü�¼
          (ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, �����־, ����id, ��ʶ��, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��,
           ���մ���id, ���ձ���, ��������, ��ҩ����, ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ͳ����, ���ʷ���, ��������id, ������, ����ʱ��,
           �Ǽ�ʱ��, ִ�в���id, ������, ����Ա���, ����Ա����, ���ʵ�id, ժҪ, �Ƿ���, �ɿ���id, ����id, ���ʽ��, ִ��״̬, ����״̬)
        Values
          (���˷��ü�¼_Id.Nextval, r_Clinic.��¼����, r_Clinic.No, r_Clinic.ʵ��Ʊ��, 2, r_Clinic.���, r_Clinic.��������, r_Clinic.�۸񸸺�,
           1, r_Clinic.����id, '', r_Clinic.����, r_Clinic.�Ա�, r_Clinic.����, r_Clinic.���˿���id, r_Clinic.�ѱ�, r_Clinic.�շ����,
           r_Clinic.�շ�ϸĿid, r_Clinic.���㵥λ, r_Clinic.������Ŀ��, r_Clinic.���մ���id, r_Clinic.���ձ���, r_Clinic.��������, r_Clinic.��ҩ����,
           r_Clinic.����, -1 * r_Clinic.����, r_Clinic.�Ӱ��־, r_Clinic.���ӱ�־, r_Clinic.������Ŀid, r_Clinic.�վݷ�Ŀ, r_Clinic.��׼����,
           -1 * r_Clinic.Ӧ�ս��, -1 * r_Clinic.ʵ�ս��, -1 * r_Clinic.ͳ����, 0, r_Clinic.��������id, r_Clinic.������, r_Clinic.����ʱ��,
           �˷�ʱ��_In, r_Clinic.ִ�в���id, r_Clinic.������, ����Ա���_In, ����Ա����_In, r_Clinic.���ʵ�id, '', r_Clinic.�Ƿ���, n_��id, n_����id,
           -1 * r_Clinic.ʵ�ս��, -1, 0);
      End Loop;
    Else
      --4.�˿�תԤ��(������Ʊ��,�ɲ���Աͨ���ش����)
    
      For r_Pay In (Select Min(a.Id) As Ԥ��id, a.���㷽ʽ, Sum(a.��Ԥ��) As ��Ԥ��, 2 As Ԥ�����, a.�����id, a.���㿨���, a.����, a.������ˮ��,
                           a.����˵��, a.������λ, b.����
                    From ����Ԥ����¼ A, ���㷽ʽ B
                    Where a.��¼���� = 3 And a.��¼״̬ In (2, 3) And
                          a.����id In (Select Column_Value From Table(f_Str2list(v_����ids))) And a.���㷽ʽ = b.���� And
                          b.���� In (1, 2, 3, 4, 7, 8) And a.���㷽ʽ Is Not Null
                    Group By a.���㷽ʽ, Ԥ�����, a.�����id, a.���㿨���, a.����, b.����, a.������ˮ��, a.����˵��, a.������λ

                    
                    Having Sum(a.��Ԥ��) <> 0) Loop
        --4.1����Ԥ����� (�����ڲ����˷ѵ����)
        --���е���,����������Ԥ�����
        --��Ϊ�տ�������ɿ�,������Ա�ɿ�����ޱ仯
        If r_Pay.���� = 7 Or (r_Pay.���� = 8 And r_Pay.�����id Is Not Null) Then
          Update ����Ԥ����¼
          Set ��Ԥ�� = ��Ԥ�� + (-1 * r_Pay.��Ԥ��), ժҪ = ժҪ || '1' || ',' || r_Pay.�����id || ',' || -1 * r_Pay.��Ԥ�� || '|'
          Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ Is Null;
          If Sql%RowCount = 0 Then
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ,
               �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��������, У�Ա�־)
            Values
              (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * r_Pay.��Ԥ��, Null, Null, �˷�ʱ��_In,
               Null, Null, Null, ����Ա���_In, ����Ա����_In, '1' || ',' || r_Pay.�����id || ',' || -1 * r_Pay.��Ԥ�� || '|', n_��id,
               Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 3, 1);
          End If;
          n_����״̬ := 1;
        Else
          If r_Pay.���� In (3, 4) Or (r_Pay.���� = 8 And r_Pay.���㿨��� Is Not Null) Then
            v_���㷽ʽ := r_Pay.���㷽ʽ;
          Else
            Begin
              Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And ���� Like '%�ֽ�%' And Rownum < 2;
            Exception
              When Others Then
                Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And Rownum < 2;
            End;
          End If;
        
          If r_Pay.���� = 8 Then
            --Zl_Square_Update(v_����ids, n_����id, n_��id, �˷�ʱ��_In, -1 * n_����id, Null, r_Pay.��Ԥ��, r_Pay.���㿨���);
            Update ����Ԥ����¼
            Set ��Ԥ�� = ��Ԥ�� + (-1 * r_Pay.��Ԥ��), ժҪ = ժҪ || '0' || ',' || r_Pay.���㿨��� || ',' || -1 * r_Pay.��Ԥ�� || '|'
            Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ Is Null;
            If Sql%RowCount = 0 Then
              Insert Into ����Ԥ����¼
                (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ,
                 �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��������, У�Ա�־)
              Values
                (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * r_Pay.��Ԥ��, Null, Null, �˷�ʱ��_In,
                 Null, Null, Null, ����Ա���_In, ����Ա����_In, '0' || ',' || r_Pay.���㿨��� || ',' || -1 * r_Pay.��Ԥ�� || '|', n_��id,
                 Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 3, 1);
            End If;
            n_����״̬ := 1;
          End If;
          If r_Pay.���� Not In (3, 4, 7, 8) Then
            Update ����Ԥ����¼
            Set ��� = ��� + r_Pay.��Ԥ��
            Where ��¼���� = 1 And ��¼״̬ = 1 And �տ�ʱ�� = �˷�ʱ��_In And ����id + 0 = n_����id And ���㷽ʽ = v_���㷽ʽ;
            If Sql%RowCount = 0 Then
              v_Ԥ��no := Nextno(11);
              Insert Into ����Ԥ����¼
                (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ,
                 �ɿ���id, Ԥ�����)
              Values
                (����Ԥ����¼_Id.Nextval, v_Ԥ��no, Null, 1, 1, n_����id, ��ҳid_In, ��Ժ����id_In, r_Pay.��Ԥ��, v_���㷽ʽ, Null, �˷�ʱ��_In,
                 Null, Null, Null, ����Ա���_In, ����Ա����_In, '����תסԺԤ��', n_��id, r_Pay.Ԥ�����);
            End If;
          
            --�������
            Update �������
            Set Ԥ����� = Nvl(Ԥ�����, 0) + r_Pay.��Ԥ��
            Where ���� = 1 And ����id = n_����id And ���� = 2
            Returning Ԥ����� Into n_����ֵ;
            If Sql%RowCount = 0 Then
              Insert Into ������� (����id, ����, ����, Ԥ�����, �������) Values (n_����id, 1, 2, r_Pay.��Ԥ��, 0);
              n_����ֵ := r_Pay.��Ԥ��;
            End If;
            If Nvl(n_����ֵ, 0) = 0 Then
              Delete From �������
              Where ����id = n_����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
            End If;
          End If;
          --4.2�ɿ����ݴ���
          --   ��Ϊû��ʵ���ղ��˵�Ǯ,���Բ�����
          --�����˷��������ԭԤ����¼
          If r_Pay.���� In (3, 4) Then
            Update ��Ա�ɿ����
            Set ��� = Nvl(���, 0) - r_Pay.��Ԥ��
            Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = r_Pay.���㷽ʽ
            Returning ��� Into n_����ֵ;
            If Sql%RowCount = 0 Then
              Insert Into ��Ա�ɿ����
                (�տ�Ա, ���㷽ʽ, ����, ���)
              Values
                (����Ա����_In, r_Pay.���㷽ʽ, 1, -1 * r_Pay.��Ԥ��);
              n_����ֵ := r_Pay.��Ԥ��;
            End If;
            If Nvl(n_����ֵ, 0) = 0 Then
              Delete From ��Ա�ɿ����
              Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = r_Pay.���㷽ʽ And Nvl(���, 0) = 0;
            End If;
          End If;
        
          If r_Pay.���㿨��� Is Null Then
            Update ����Ԥ����¼
            Set ��Ԥ�� = ��Ԥ�� + (-1 * r_Pay.��Ԥ��)
            Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_���㷽ʽ;
            If Sql%RowCount = 0 Then
              Insert Into ����Ԥ����¼
                (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ,
                 �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, У�Ա�־, ��������)
              Values
                (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * r_Pay.��Ԥ��, v_���㷽ʽ, Null, �˷�ʱ��_In,
                 Null, Null, Null, ����Ա���_In, ����Ա����_In, '', n_��id, r_Pay.�����id, r_Pay.���㿨���, r_Pay.����, r_Pay.������ˮ��,
                 r_Pay.����˵��, r_Pay.������λ, n_����id, -1 * n_����id, 0, 3);
            End If;
          End If;
        End If;
      End Loop;
    End If;
    If ����_In Is Not Null Then
      Begin
        Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And ���� Like '%�ֽ�%' And Rownum < 2;
      Exception
        When Others Then
          Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And Rownum < 2;
      End;
      Update ����Ԥ����¼
      Set ��Ԥ�� = ��Ԥ�� - ����_In
      Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_���㷽ʽ;
      Update ����Ԥ����¼
      Set ��Ԥ�� = ��Ԥ�� + ����_In
      Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_����;
      If Sql%RowCount = 0 Then
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ,
           �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, У�Ա�־, ��������)
        Values
          (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, ����_In, v_����, Null, �˷�ʱ��_In, Null, Null,
           Null, ����Ա���_In, ����Ա����_In, '', n_��id, Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 0, 3);
      End If;
    End If;
    Delete From ����Ԥ����¼ Where ����id = n_ԭ����id And ժҪ = 'Ԥ����ʱ��¼' And ��¼���� = 3;
    Delete From ����Ԥ����¼
    Where ����id = n_����id And ��¼���� = 3 And ��¼״̬ = 2 And ��Ԥ�� = 0 And ���㷽ʽ Is Not Null;
    Update ������ü�¼
    Set ����״̬ = Nvl(n_����״̬, 0)
    Where NO In (Select Column_Value From Table(f_Str2list(v_Nos))) And Mod(��¼����, 10) = 1 And ��¼״̬ = 2;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����תסԺ_�շ�ת��;
/

--123754:Ƚ����,2018-04-26,ҽ������վԤԼ�Һ���������
Create Or Replace Procedure Zl1_Auto_Buildingregisterplan
(
  �Һ�ʱ��_In In Date := Null,
  ��Դid_In   �ٴ������Դ.Id%Type := Null
) As
  -------------------------------------------------------------------------
  --����˵�����Զ������ٴ������¼
  --          1�����ݺ�Դ�Զ�����ԤԼ���ڵ��ٴ������¼;
  --          2��ԤԼ������ȷ��:��ԴԤԼ����-->ԤԼ��ʽ��������ȡ���)-->ϵͳԤԼ����
  --���:�Һ�ʱ��_IN:NULLʱ���Զ�����;����ֻ���ָ�������Ƿ������˳����¼û��
  --    ��Դid_In:NULLʱ�������к�Դ������ֻ����ָ����Դ
  -------------------------------------------------------------------------
  n_ȱʡԤԼ���� �ٴ������Դ.ԤԼ����%Type;
  v_����Ա����   �ٴ����ﰲ��.����Ա����%Type;
  d_�Ǽ�����     �ٴ����ﰲ��.�Ǽ�ʱ��%Type;
  n_����id       �ٴ����ﰲ��.Id%Type;
  n_��Ŀid       �ٴ����ﰲ��.��Ŀid %Type;

  n_��¼id   �ٴ������¼.Id%Type;
  d_��ǰ���� �ٴ������¼.��������%Type;

  l_�̶�ʱ�� t_Strlist := t_Strlist();
  n_Count    Number(18);

  n_��ԤԼ���� Number := 0;
  d_��ʼʱ��   �ٴ������¼.��ʼʱ��%Type;
Begin

  Select Max(ԤԼ����) Into n_ȱʡԤԼ���� From ԤԼ��ʽ;
  If Nvl(n_ȱʡԤԼ����, 0) = 0 Then
    n_ȱʡԤԼ���� := To_Number(Nvl(zl_GetSysParameter('�Һ�����ԤԼ����'), '0'));
  End If;
  If Nvl(n_ȱʡԤԼ����, 0) = 0 Then
    n_ȱʡԤԼ���� := 7;
  End If;

  --�԰���Ϊ��λ,�����������Դ����ʱ�䡱��12:00:00-23:59:59�ڼ�ģ��򿪷�ԤԼ����+1��
  n_��ԤԼ���� := Zl_Fun_Getappointmentdays;

  d_��ǰ����   := Trunc(Nvl(�Һ�ʱ��_In, Sysdate));
  d_�Ǽ�����   := Sysdate;
  v_����Ա���� := Zl_Username;

  --��һ��ѭ������Դ��Ϣ
  For c_��Դ In (Select c.Id, c.����, c.����, c.��Ŀid, c.����id, c.ҽ������,
                      Decode(Nvl(c.ԤԼ����, 0), 0, n_ȱʡԤԼ����, c.ԤԼ����) + n_��ԤԼ���� As ԤԼ����, Nvl(b.վ��, '-') As վ��,
                      Nvl(c.�Ƿ���ջ���, 0) As �Ƿ���ջ���, Nvl(c.���տ���״̬, 0) As ���տ���״̬, Nvl(c.�Ű෽ʽ, 0) As �Ű෽ʽ
               From �ٴ������Դ C, ���ű� B, ��Ա�� A, �շ���ĿĿ¼ D
               Where c.����id = b.Id And c.ҽ��id = a.Id(+) And c.��Ŀid = d.Id And Nvl(c.�Ƿ�ɾ��, 0) = 0 And
                     Nvl(c.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(b.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(a.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(d.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     (��Դid_In Is Null Or c.Id = ��Դid_In)
                    --
                     And Exists (Select 1
                      From �ٴ����ﰲ�� M, �ٴ������ N
                      Where m.����id = n.Id And m.��Դid = c.Id And Nvl(n.�Ű෽ʽ, 0) = 0 And n.����ʱ�� Is Not Null And
                            m.���ʱ�� Is Not Null And d_��ǰ���� <= m.��ֹʱ��)) Loop
  
    --��鵱ǰ�������ڵİ��ŵ��շ���Ŀ�Ƿ�Ϊ��Դ�е��շ���Ŀ��������ǣ�����º�Դ�е��շ���Ŀ
    Begin
      Select ��Ŀid
      Into n_��Ŀid
      From (Select a.��Ŀid
             From �ٴ����ﰲ�� A, �ٴ������ B
             Where a.����id = b.Id And a.��Դid = c_��Դ.Id And a.���ʱ�� Is Not Null And d_��ǰ���� Between a.��ʼʱ�� And a.��ֹʱ�� And
                   Nvl(b.�Ű෽ʽ, 0) = 0 And b.����ʱ�� Is Not Null
             Order By a.�Ǽ�ʱ�� Desc)
      Where Rownum < 2;
    Exception
      When Others Then
        n_��Ŀid := Null;
    End;
    If Nvl(n_��Ŀid, 0) <> 0 Then
      If Nvl(c_��Դ.��Ŀid, 0) <> n_��Ŀid Then
        Update �ٴ������Դ Set ��Ŀid = n_��Ŀid Where ID = c_��Դ.Id;
        Commit;
      End If;
    End If;
  
    --�ڶ���ѭ������������
    --��ͷһ�쿪ʼ���ɣ�������ȫ��(8:00-7:59)��0:00-7:59û�г����¼
    --1.δָ����ԴID�������������ɳ����¼���г����¼�����ڽ����ٴ���
    --2.ָ���˺�ԴID���϶��Ƿ�������������ʱ�����������ɳ����¼
    For c_���� In (Select m.����,
                        Decode(To_Char(m.����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7',
                                '����', Null) As ����
                 From (Select Trunc(d_��ǰ����) + ���� As ����
                        From (Select Level - 1 As ���� From Dual Connect By Level <= c_��Դ.ԤԼ���� + 1)
                        Where ��Դid_In Is Not Null
                        Union All
                        Select Trunc(d_��ǰ���� - 1) + ���� As ����
                        From (Select Level - 1 As ���� From Dual Connect By Level <= c_��Դ.ԤԼ���� + 2)
                        Where ��Դid_In Is Null And Not Exists
                         (Select 1
                               From �ٴ������¼ A
                               Where a.��Դid = c_��Դ.Id And a.�������� = Trunc(d_��ǰ���� - 1) + ����)) M
                 Where �Һ�ʱ��_In Is Null Or Trunc(�Һ�ʱ��_In) = m.����) Loop
    
      l_�̶�ʱ�� := t_Strlist();
      --��鵱���Ƿ�����/�ܳ������,���ڣ������ɳ����¼
      Select Count(1)
      Into n_Count
      From �ٴ����ﰲ�� A, �ٴ������ B
      Where a.����id = b.Id And a.��Դid = c_��Դ.Id And c_����.���� Between Trunc(a.��ʼʱ��) And Trunc(a.��ֹʱ��) And
            Nvl(b.�Ű෽ʽ, 0) In (1, 2) And Rownum < 2;
    
      --��ǰ��ԴΪ����/���Ű࣬�ҵ�ǰ����֮ǰ���а���/���Ű�ĳ����¼�Ͳ��ٰ��̶��������ɳ����¼��
      If n_Count = 0 And Nvl(c_��Դ.�Ű෽ʽ, 0) <> 0 Then
        Select Count(1)
        Into n_Count
        From �ٴ����ﰲ�� A, �ٴ������ B
        Where a.����id = b.Id And Nvl(b.�Ű෽ʽ, 0) In (1, 2) And a.��Դid = c_��Դ.Id And a.��ʼʱ�� < c_����.���� And Rownum < 2;
      End If;
    
      If n_Count = 0 Then
        If ��Դid_In Is Null Then
          --���ﰲ��,ȡ���Ǽǵ�һ��
          Begin
            Select ����id
            Into n_����id
            From (Select a.Id As ����id
                   From �ٴ����ﰲ�� A, �ٴ������ B
                   Where a.��Դid = c_��Դ.Id And a.����id = b.Id And Nvl(b.�Ű෽ʽ, 0) = 0 And b.����ʱ�� Is Not Null And
                         a.���ʱ�� Is Not Null And c_����.���� Between a.��ʼʱ�� And a.��ֹʱ��
                   Order By a.�Ǽ�ʱ�� Desc)
            Where Rownum < 2;
          Exception
            When Others Then
              n_����id := 0;
          End;
        Else
          --���ָ���˺�ԴID���϶��Ƿ�������������ʱ�����������ɳ����¼�����Ǽǵ�һ���϶��Ǳ��������ģ�
          --ֻ��Ҫ����������ż��ɣ��������������Чʱ�䷶Χ�ڵľͲ�����
          Begin
            Select ����id
            Into n_����id
            From (Select a.Id As ����id, a.��ʼʱ��, a.��ֹʱ��, Row_Number() Over(Order By a.�Ǽ�ʱ�� Desc) As �к�
                   From �ٴ����ﰲ�� A, �ٴ������ B
                   Where a.��Դid = c_��Դ.Id And a.����id = b.Id And Nvl(b.�Ű෽ʽ, 0) = 0 And b.����ʱ�� Is Not Null And
                         a.���ʱ�� Is Not Null And c_����.���� Between ��ʼʱ�� And ��ֹʱ��)
            Where �к� = 1;
          Exception
            When Others Then
              n_����id := 0;
          End;
        End If;
      
        If Nvl(n_����id, 0) <> 0 Then
          If ��Դid_In Is Not Null Then
            --2.ָ���˺�ԴID���϶��Ƿ�������������ʱ�����������ɳ����¼
            --�����г����¼����Ҫ�����´���
            For c_��¼ In (Select a.����id, a.Id As ��¼id, a.��������, a.�ϰ�ʱ��, a.�Ƿ��ʱ��, a.�Ƿ���ſ���
                         From �ٴ������¼ A
                         Where a.��Դid = c_��Դ.Id And a.�������� = c_����.����) Loop
            
              Select Count(1) Into n_Count From ���˹Һż�¼ Where �����¼id = c_��¼.��¼id;
              If n_Count = 0 Then
                --2.2.1���ʱ�β�����ԤԼ�Һ����ݣ���ɾ����������
                Zl_�ٴ������ϰ�ʱ��_Delete(c_��¼.����id, To_Char(c_��¼.��������, 'yyyy-mm-dd'), 1, c_��¼.�ϰ�ʱ��);
              Else
                --2.2.2���ʱ�δ���ԤԼ�Һ����ݣ���ֻ����������¼�İ���ID����
                Update �ٴ������¼ Set ����id = n_����id Where ID = c_��¼.��¼id;
                l_�̶�ʱ��.Extend();
                l_�̶�ʱ��(l_�̶�ʱ��.Count) := c_��¼.�ϰ�ʱ��;
              End If;
            End Loop;
          End If;
        
          --��������Ƿ����
          Select Count(1) Into n_Count From �ٴ��������� Where ����id = n_����id And ������Ŀ = c_����.����;
          If n_Count = 0 Then
            --����������ٴ������¼���������ٴ������¼(ʱ���ΪNULL �Ŀռ�¼)
            Insert Into �ٴ������¼
              (ID, ����id, ��Դid, ��������, �Ǽ���, �Ǽ�ʱ��)
              Select �ٴ������¼_Id.Nextval, n_����id, a.Id As ID, c_����.����, v_����Ա����, d_�Ǽ����� As �Ǽ�ʱ��
              From �ٴ������Դ A, �ٴ����ﰲ�� B
              Where a.Id = b.��Դid And b.Id = n_����id And Not Exists
               (Select 1 From �ٴ������¼ Where ��Դid = a.Id And �������� = c_����.����);
          Else
            For c_��¼ In (With c_ʱ��� As
                            (Select ʱ���, ��ʼʱ��, ��ֹʱ��, ����, վ��, ȱʡʱ��, ��ǰʱ��
                            From (Select ʱ���, ��ʼʱ��, ��ֹʱ��, ����, վ��, ȱʡʱ��, ��ǰʱ��,
                                          Row_Number() Over(Partition By ʱ��� Order By ʱ���, վ�� Asc, ���� Asc) As ���
                                   From ʱ���
                                   Where Nvl(վ��, c_��Դ.վ��) = c_��Դ.վ�� And Nvl(����, c_��Դ.����) = c_��Դ.����)
                            Where ��� = 1)
                           Select n_����id As ����id, B1.��Դid, c_����.���� As ��������, m.�ϰ�ʱ��, m.Id As ����id,
                                  To_Date(To_Char(c_����.����, 'yyyy-mm-dd ') || To_Char(j.��ʼʱ��, 'hh24:mi:ss'),
                                           'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                                  To_Date(To_Char(c_����.����, 'yyyy-mm-dd ') || To_Char(j.��ֹʱ��, 'hh24:mi:ss'),
                                          'yyyy-mm-dd hh24:mi:ss') + Case
                                    When j.��ֹʱ�� <= j.��ʼʱ�� Then
                                     1
                                    Else
                                     0
                                  End As ��ֹʱ��, Null As ͣ�￪ʼʱ��, Null As ͣ����ֹʱ��, Null As ͣ��ԭ��,
                                  To_Date(To_Char(c_����.����, 'yyyy-mm-dd ') || To_Char(Nvl(j.ȱʡʱ��, j.��ʼʱ��), 'hh24:mi:ss'),
                                          'yyyy-mm-dd hh24:mi:ss') + Case
                                    When j.ȱʡʱ�� < j.��ʼʱ�� Then
                                     1
                                    Else
                                     0
                                  End As ȱʡԤԼʱ��,
                                  To_Date(To_Char(c_����.����, 'yyyy-mm-dd ') || To_Char(Nvl(j.��ǰʱ��, j.��ʼʱ��), 'hh24:mi:ss'),
                                          'yyyy-mm-dd hh24:mi:ss') + Case
                                    When j.��ʼʱ�� < j.��ǰʱ�� Then
                                     -1
                                    Else
                                     0
                                  End As ��ǰ�Һ�ʱ��, m.�޺���, 0 As �ѹ���, m.��Լ��, 0 As ��Լ��, 0 As �����ѽ���, m.�Ƿ���ſ���, m.�Ƿ��ʱ��, m.ԤԼ����,
                                  m.�Ƿ��ռ, B1.��Ŀid, B1.ҽ��id, B1.ҽ������, Null As ����ҽ��id, Null As ����ҽ������, m.���﷽ʽ, m.����id,
                                  0 As �Ƿ�����, 0 As �Ƿ���ʱ����, v_����Ա���� As ����Ա����, d_�Ǽ����� As �Ǽ�ʱ��, c_����.���� As ������Ŀ
                           From �ٴ����ﰲ�� B1, �ٴ��������� M, c_ʱ��� J
                           Where B1.Id = n_����id And B1.Id = m.����id And m.������Ŀ = c_����.���� And m.�ϰ�ʱ�� = j.ʱ��� And
                                 To_Date(To_Char(c_����.����, 'yyyy-mm-dd ') || To_Char(j.��ʼʱ��, 'hh24:mi:ss'),
                                         'yyyy-mm-dd hh24:mi:ss') >= B1.��ʼʱ�� And Not Exists
                            (Select 1 From Table(l_�̶�ʱ��) Where Column_Value = m.�ϰ�ʱ��)) Loop
            
              Select �ٴ������¼_Id.Nextval Into n_��¼id From Dual;
              Insert Into �ٴ������¼
                (ID, ����id, ��Դid, ��������, �ϰ�ʱ��, ��ʼʱ��, ��ֹʱ��, ͣ�￪ʼʱ��, ͣ����ֹʱ��, ͣ��ԭ��, ȱʡԤԼʱ��, ��ǰ�Һ�ʱ��, �޺���, �ѹ���, ��Լ��, ��Լ��,
                 �����ѽ���, �Ƿ���ſ���, �Ƿ��ʱ��, ԤԼ����, �Ƿ��ռ, ��Ŀid, ����id, ҽ��id, ҽ������, ����ҽ��id, ����ҽ������, ���﷽ʽ, ����id, �Ƿ�����, �Ƿ���ʱ����, �Ǽ���,
                 �Ǽ�ʱ��, �Ƿ񷢲�)
              Values
                (n_��¼id, c_��¼.����id, c_��¼.��Դid, c_��¼.��������, c_��¼.�ϰ�ʱ��, c_��¼.��ʼʱ��, c_��¼.��ֹʱ��, c_��¼.ͣ�￪ʼʱ��, c_��¼.ͣ����ֹʱ��,
                 c_��¼.ͣ��ԭ��, c_��¼.ȱʡԤԼʱ��, c_��¼.��ǰ�Һ�ʱ��, c_��¼.�޺���, c_��¼.�ѹ���, c_��¼.��Լ��, c_��¼.��Լ��, c_��¼.�����ѽ���, c_��¼.�Ƿ���ſ���,
                 c_��¼.�Ƿ��ʱ��, c_��¼.ԤԼ����, c_��¼.�Ƿ��ռ, c_��¼.��Ŀid, c_��Դ.����id, c_��¼.ҽ��id, c_��¼.ҽ������, c_��¼.����ҽ��id, c_��¼.����ҽ������,
                 c_��¼.���﷽ʽ, c_��¼.����id, c_��¼.�Ƿ�����, c_��¼.�Ƿ���ʱ����, c_��¼.����Ա����, d_�Ǽ�����, 1);
            
              d_��ʼʱ�� := c_��¼.��ʼʱ��;
              --�����ٴ�������ſ���
              If Nvl(c_��¼.�Ƿ��ʱ��, 0) = 1 And Nvl(c_��¼.�Ƿ���ſ���, 0) = 1 Then
                --��ʱ����������ſ��ƣ�ʹ��"ԤԼ˳���"��¼"�Ƿ�ԤԼ"
                Insert Into �ٴ�������ſ���
                  (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, ԤԼ˳���)
                  Select n_��¼id, ���,
                         To_Date(To_Char(c_��¼.��������, 'yyyy-mm-dd ') || To_Char(��ʼʱ��, 'hh24:mi:ss'),
                                  'yyyy-mm-dd hh24:mi:ss') + Case
                            When d_��ʼʱ�� > To_Date(To_Char(c_��¼.��������, 'yyyy-mm-dd ') || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') Then
                             1
                            Else
                             0
                          End,
                         To_Date(To_Char(c_��¼.��������, 'yyyy-mm-dd ') || To_Char(��ֹʱ��, 'hh24:mi:ss'),
                                  'yyyy-mm-dd hh24:mi:ss') + Case
                            When d_��ʼʱ�� >= To_Date(To_Char(c_��¼.��������, 'yyyy-mm-dd ') || To_Char(��ֹʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') Then
                             1
                            Else
                             0
                          End, ��������, �Ƿ�ԤԼ, �Ƿ�ԤԼ
                  From �ٴ�����ʱ��
                  Where ����id = c_��¼.����id;
              Else
                Insert Into �ٴ�������ſ���
                  (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ)
                  Select n_��¼id, ���,
                         To_Date(To_Char(c_��¼.��������, 'yyyy-mm-dd ') || To_Char(��ʼʱ��, 'hh24:mi:ss'),
                                 'yyyy-mm-dd hh24:mi:ss') + Case
                           When d_��ʼʱ�� > To_Date(To_Char(c_��¼.��������, 'yyyy-mm-dd ') || To_Char(��ʼʱ��, 'hh24:mi:ss'),
                                                 'yyyy-mm-dd hh24:mi:ss') Then
                            1
                           Else
                            0
                         End,
                         To_Date(To_Char(c_��¼.��������, 'yyyy-mm-dd ') || To_Char(��ֹʱ��, 'hh24:mi:ss'),
                                 'yyyy-mm-dd hh24:mi:ss') + Case
                           When d_��ʼʱ�� >= To_Date(To_Char(c_��¼.��������, 'yyyy-mm-dd ') || To_Char(��ֹʱ��, 'hh24:mi:ss'),
                                                  'yyyy-mm-dd hh24:mi:ss') Then
                            1
                           Else
                            0
                         End, ��������, �Ƿ�ԤԼ
                  From �ٴ�����ʱ��
                  Where ����id = c_��¼.����id;
              End If;
            
              --���������λ�Һſ��Ƽ�¼
              Insert Into �ٴ�����Һſ��Ƽ�¼
                (����, ����, ����, ��¼id, ���, ���Ʒ�ʽ, ����)
                Select ����, ����, ����, n_��¼id, ���, ���Ʒ�ʽ, ����
                From �ٴ�����Һſ���
                Where ����id = c_��¼.����id;
            
              --�����ٴ��������Ҽ�¼
              Insert Into �ٴ��������Ҽ�¼
                (��¼id, ����id)
                Select n_��¼id, ����id From �ٴ��������� Where ����id = c_��¼.����id;
            End Loop;
          
            --����ͣ�ﰲ�źͷ����ڼ��յ��������¼�ĳ���/ԤԼ���
            Zl_Clinicvisitmodify(c_��Դ.Id, n_����id, c_����.����, v_����Ա����, d_�Ǽ�����);
          End If;
        End If;
      End If;
      --һ��һ�ύ
      Commit;
    End Loop;
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl1_Auto_Buildingregisterplan;
/

--123726:Ƚ����,2018-04-26,�������ķ�������ʱ����
Create Or Replace Procedure Zl_���˷�������_Audit
(
  Id_In       ���˷�������.����id%Type,
  ����ʱ��_In ���˷�������.����ʱ��%Type,
  �����_In   ���˷�������.�����%Type,
  ���ʱ��_In ���˷�������.���ʱ��%Type,
  ״̬_In     ���˷�������.״̬%Type,
  Int�Զ����� Integer := 1,
  �������_In ���˷�������.�������%Type := 1 --��ҩƷ��������Ч,ȱʡΪ��ִ�е�ҩƷ������ 
) As
  n_ִ��״̬       סԺ���ü�¼.ִ��״̬%Type;
  n_�������       ���˷�������.�������%Type;
  v_�շ����       סԺ���ü�¼.�շ����%Type;
  v_No             סԺ���ü�¼.No%Type;
  n_ʵ������       ҩƷ�շ���¼.ʵ������%Type;
  n_����           ���˷�������.����%Type;
  n_�շ�id         ҩƷ�շ���¼.Id%Type;
  n_ҽ��id         סԺ���ü�¼.Id%Type;
  v_��������       ��������.��������%Type;
  n_�շ�ϸĿid     סԺ���ü�¼.�շ�ϸĿid%Type;
  n_��˲���id     ���˷�������.��˲���id%Type;
  n_ִ�в���id     סԺ���ü�¼.ִ�в���id%Type;
  n_����id         סԺ���ü�¼.����id%Type;
  n_��ҳid         סԺ���ü�¼.��ҳid%Type;
  n_��˱�־       ������ҳ.��˱�־%Type;
  n_סԺ״̬       ������ҳ.״̬%Type;
  n_������˷�ʽ   Number(2);
  n_δ��ƽ�ֹ���� Number(2);

  n_Cnt     Number(18);
  n_Temp    Number(18);
  v_Err_Msg Varchar2(300);
  Err_Item Exception;
Begin

  n_������� := 0;
  Select a.ִ��״̬, a.�շ����, a.�շ�ϸĿid, a.ִ�в���id, a.No, Nvl(b.��������, 0), a.ҽ�����, ����id, ��ҳid
  Into n_ִ��״̬, v_�շ����, n_�շ�ϸĿid, n_ִ�в���id, v_No, v_��������, n_ҽ��id, n_����id, n_��ҳid
  From סԺ���ü�¼ A, �������� B
  Where a.Id = Id_In And a.�շ�ϸĿid = b.����id(+);

  If Nvl(n_��ҳid, 0) <> 0 Then
  
    n_������˷�ʽ   := Nvl(zl_GetSysParameter(185), 0);
    n_δ��ƽ�ֹ���� := Nvl(zl_GetSysParameter(215), 0);
    If n_������˷�ʽ = 1 Or n_δ��ƽ�ֹ���� = 1 Then
      Begin
        Select ��˱�־, ״̬
        Into n_��˱�־, n_סԺ״̬
        From ������ҳ
        Where ����id = Nvl(n_����id, 0) And ��ҳid = Nvl(n_��ҳid, 0);
      Exception
        When Others Then
          n_��˱�־ := 0;
          n_סԺ״̬ := 0;
      End;
      If n_δ��ƽ�ֹ���� = 1 And n_סԺ״̬ = 1 Then
        v_Err_Msg := '����δ���,��ֹ�Բ�����ط��õĲ���!';
        Raise Err_Item;
      End If;
    
      If n_������˷�ʽ = 1 Then
        If Nvl(n_��˱�־, 0) = 1 Then
          v_Err_Msg := '�ò���Ŀǰ������˷���,���ܽ��з�����ص���!';
          Raise Err_Item;
        End If;
        If Nvl(n_��˱�־, 0) = 2 Then
          v_Err_Msg := '�ò���Ŀǰ�Ѿ�����˷������,���ܽ��з�����ص���!';
          Raise Err_Item;
        End If;
      End If;
    End If;
  
  End If;
  If Instr(',5,6,7', ',' || v_�շ����) > 0 Or (v_�շ���� = '4' And Nvl(v_��������, 0) = 1) Then
    n_������� := �������_In;
  End If;

  Update ���˷�������
  Set ����� = �����_In, ���ʱ�� = ���ʱ��_In, ״̬ = ״̬_In
  Where ����id = Id_In And ������� = n_������� And ����ʱ�� = ����ʱ��_In And ״̬ = 0
  Returning ����, ��˲���id Into n_����, n_��˲���id;

  If Sql%RowCount = 0 Then
    v_Err_Msg := '�������ʧ��,��ǰ�����ļ�¼������Ϊ���������Ѿ������˴���,����ˢ����Ϣ!';
    Raise Err_Item;
  End If;

  If n_������� = 0 And (Instr(',5,6,7', ',' || v_�շ����) > 0 Or (v_�շ���� = '4' And Nvl(v_��������, 0) = 1)) Then
    --��Ҫ���δִ�е���������ȫ������,�Ż�ͨ�� 
    Select Sum(Nvl(����, 0) * Nvl(ʵ������, 0))
    Into n_ʵ������
    From ҩƷ�շ���¼
    Where ������� Is Null And ����id = Id_In And Instr(',8,9,10,21,24,25,26,', ',' || ���� || ',') > 0;
    If Nvl(n_ʵ������, 0) < Nvl(n_����, 0) Then
      Select '�ڵ��ݺ�<<' || v_No || '>>��' || Decode(v_�շ����, '4', '����', 'ҩƷ') || 'Ϊ:' || Chr(13) || ���� || '-' || ���� ||
              Chr(13) || '����������(' || LTrim(To_Char(n_����, '9999999990.99')) || ')�����˴���' || Decode(v_�շ����, '4', '��', 'ҩ') ||
              '����(' || LTrim(To_Char(Nvl(n_ʵ������, 0), '9999999990.99')) || '),���������!'
      Into v_Err_Msg
      From �շ���ĿĿ¼
      Where ID = n_�շ�ϸĿid;
      Raise Err_Item;
    End If;
  
    If n_ҽ��id <> 0 Then
      Select Nvl(Max(d.Id), 0)
      Into n_Cnt
      From ����ҽ����¼ A, ����ҽ������ B, ��Һ��ҩ��¼ D
      Where a.Id = n_ҽ��id And a.Id = b.ҽ��id And b.No = v_No And a.���id = d.ҽ��id And b.���ͺ� = d.���ͺ� And b.��¼���� = 2 And
            d.����ʱ�� = ����ʱ��_In And d.����״̬ = 9;
    
      If n_Cnt <> 0 Then
        Select Count(1)
        Into n_Temp
        From ��Һ��ҩ״̬
        Where ��ҩid = n_Cnt And �������� = 10 And ����ʱ�� = ���ʱ��_In;
        If n_Temp = 0 Then
          Insert Into ��Һ��ҩ״̬ (��ҩid, ��������, ������Ա, ����ʱ��) Values (n_Cnt, 10, �����_In, ���ʱ��_In);
        End If;
        Update ��Һ��ҩ��¼ Set ������Ա = �����_In, ����ʱ�� = ���ʱ��_In, ����״̬ = 10 Where ID = n_Cnt;
      End If;
    End If;
  End If;

  If n_ִ��״̬ <> 0 Then
    If Instr(',5,6,7,', ',' || v_�շ���� || ',') > 0 And n_������� = 1 Then
      If n_ִ�в���id <> n_��˲���id Then
        Begin
          Select '[' || ���� || ']' || ���� Into v_Err_Msg From �շ���ĿĿ¼ Where ID = n_�շ�ϸĿid;
        Exception
          When Others Then
            v_Err_Msg := '';
        End;
        v_Err_Msg := '���������ʱ,ҩƷΪ' || v_Err_Msg || ' ���Ѿ���ִ�п���ִ��,�����ٽ����������,��ȡ�����!';
        Raise Err_Item;
      End If;
    End If;
  
    If v_�շ���� = '4' Then
      If v_�������� = 1 Then
        If n_ִ�в���id <> n_��˲���id And n_������� = 1 And Int�Զ����� <> 1 Then
          Begin
            Select '[' || ���� || ']' || ���� Into v_Err_Msg From �շ���ĿĿ¼ Where ID = n_�շ�ϸĿid;
          Exception
            When Others Then
              v_Err_Msg := '';
          End;
          v_Err_Msg := '���������ʱ,����Ϊ' || v_Err_Msg || ' ���Ѿ���ִ�п���ִ��,�����ٽ����������,��ȡ�����!';
          Raise Err_Item;
        End If;
      
        If n_������� = 1 And Int�Զ����� = 1 Then
          n_�շ�id := -1;
          --���������ڶ������ 
          For c_�շ���¼ In (Select ID, ����, Nvl(Sum(Nvl(����, 1) * ʵ������), 0) As ����
                         From ҩƷ�շ���¼
                         Where ����id = Id_In And ���� In (25, 26) And (��¼״̬ = 1 Or Mod(��¼״̬, 3) = 0)
                         Group By ID, ����) Loop
            n_�շ�id := c_�շ���¼.Id;
            If n_���� = 0 Then
              Exit;
            End If;
          
            If n_���� > c_�շ���¼.���� Then
              n_Temp := c_�շ���¼.����;
              n_���� := n_���� - c_�շ���¼.����;
            Else
              n_Temp := n_����;
              n_���� := 0;
            End If;
            Zl_�����շ���¼_��������(c_�շ���¼.Id, �����_In, ���ʱ��_In, c_�շ���¼.����, Null, Null, n_Temp, 0);
          End Loop;
          If n_�շ�id = -1 Then
            v_Err_Msg := '���������ʱ,����Ϊ' || v_Err_Msg || ' ��δ�ҵ���ص�ҩƷ�շ���Ϣ,��������Ϊ��;' || Chr(13) ||
                         '���������ĵĸ�������,�����ٽ����������,��ȡ�����!';
            Raise Err_Item;
          End If;
        End If;
      Else
        --���Ǹ��ٵ����� 
        Update סԺ���ü�¼ Set ִ��״̬ = 0 Where ID = Id_In;
      End If;
    Elsif Instr(',5,6,7,', ',' || v_�շ���� || ',') = 0 Then
      --���ܴ��ڲ�������,�����Ƚ���ҩƷ�Ĵ���ɲ���ִ��,����������˹���(ZL_סԺ���ʼ�¼_Delete)�д���,�����������: 
      --�ڵ��ñ�����ʱ: 
      --   1.������Ѿ�ִ�е�,���Ϊ����ִ��(ִ��״̬=2);�������ʹ����д����ⲿ������(ZL_סԺ���ʼ�¼_Delete):��:���ִ��״̬=2,���Ҳ������ʵ�,���Ϊ1(��ִ��) 
      --      ԭ������Ϊ��ҩƷ��ֻ�ܴ�������״̬.��ִ��;2-δִ�� 
      --   2.�����δִ�е�,��ִ��״̬����Ϊ0,�������ʹ����м�¼״̬���ֲ��� 
      Update סԺ���ü�¼ Set ִ��״̬ = Decode(Nvl(ִ��״̬, 0), 0, 0, 2) Where ID = Id_In; --��ҩƷ����û��ȡ��ִ�еĲ���,���Զ���ִ�е�Ҫ�ȸ�״̬���ܵ����� 
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˷�������_Audit;
/

--124675:������,2018-04-20,��ת��3201ҽԺ����ʷ���ݲ����з��ֵ��������ݵĴ���
CREATE OR REPLACE Procedure Zl1_Datamove_Tag
(
  d_End            In Date,
  n_����           In Number,
  n_System         In Number,
  n_Ԥ��ʣ������� In ����Ԥ����¼.���%Type := 10 --�����˲�����δ����ã�Ҳ������Ժ����ʱ������δ�����Ԥ������ָ��ֵ���µ�����ǿ��ת���������������δת���Ӷ�Ӱ��ת���ٶ�


) As
  --���ܣ���Ǵ�ת��������
  --˵����Ϊ����Undo��ռ����͹��󣬷ֶ��ύ
  d_Lastend Date; --����ת����ֹʱ�䣨d_EndΪ����ת����ֹʱ�䣩

  --�ݹ�ȡ����һ��Ԥ������е�һ���ֱ����Ϊ��ת����������
  Procedure Datamove_Tag_Update
  (
    ����id_In t_Numlist,
    d_End     In Date,
    n_����    In Number
  ) As

    c_����id t_Numlist := t_Numlist();
    c_No     t_Strlist := t_Strlist();
  Begin
    --1.1һ��Ԥ�����ݱ��������ID���ˣ��ҳ����е�һ���ֱ����Ϊ��ת�������ݣ��磺
    --   NO=A001 ��¼����=11 ����ID=10 ��ת��=1
    --   NO=A001 ��¼����=11 ����ID=11 ��ת��=NULL
    If ����id_In Is Null Then
      Select Distinct a.No
      Bulk Collect
      Into c_No
      From ����Ԥ����¼ A
      Where a.��¼���� In (1, 11) And a.��ת�� = n_���� And Exists
       (Select 1 From ����Ԥ����¼ Where NO = a.No And ��¼���� In (1, 11) And ��ת�� Is Null);
    Else
      Select Distinct a.No
      Bulk Collect
      Into c_No
      From ����Ԥ����¼ A
      Where a.����id In (Select /*+cardinality(b,10) */
                        Column_Value
                       From Table(����id_In) B) And a.��¼���� In (1, 11) And a.��ת�� Is Null And Exists
       (Select 1 From ����Ԥ����¼ Where NO = a.No And ��¼���� In (1, 11) And ��ת�� + 0 = n_����);
    End If;

    If c_No.Count = 0 Then
      Return;
    End If;

    --1.2ȡ�����
    Forall I In 1 .. c_No.Count
      Update ����Ԥ����¼ Set ��ת�� = Null Where NO = c_No(I) And ��¼���� In (1, 11);

    --------------------------------------------------------------------------------------------------------
    --2.1һ������ID���˶���Ԥ�����ݣ��ҳ����е�һ���ֱ����Ϊ��ת�������ݣ��磺
    --   NO=A001 ��¼����=11 ����ID=20 ��ת��=1
    --   NO=A002 ��¼����=11 ����ID=20 ��ת��=NULL
    Select Distinct a.����id
    Bulk Collect
    Into c_����id
    From ����Ԥ����¼ A
    Where a.No In (Select /*+cardinality(b,10) */
                    Column_Value
                   From Table(c_No) B) And a.��¼���� In (1, 11) And a.��ת�� Is Null And a.�տ�ʱ�� + 0 < d_End And Exists
     (Select 1 From ����Ԥ����¼ Where ����id = a.����id And ��ת�� + 0 = n_����);

    If c_����id.Count = 0 Then
      Return;
    End If;

    --2.2ȡ�����(����һ�ν��ʵ��������㷽ʽ�ļ�¼)
    Forall I In 1 .. c_����id.Count
      Update ����Ԥ����¼ Set ��ת�� = Null Where ����id = c_����id(I);

    --�ݹ����
    Datamove_Tag_Update(c_����id, d_End, n_����);
  End Datamove_Tag_Update;
Begin
  Select ������������ Into d_Lastend From zlDataMove Where ϵͳ = n_System And ��� = 1;
  If d_Lastend Is Null Then
    Return;
  End If;
  --�¼��Ӳ�ѯע�������Ż������ܹ������ݹ��˵���С�������ŵ����Exists��������ǰ��

  --1.���ú��㣨����,ҩƷ,�տ��Ʊ�ݵȣ�
  --����ҵ����ԭʼҵ��ķ���ʱ����ͬ���Ǽ�ʱ�䲻ͬ������Ҫ������ʱ������ѯ.
  --��������������ж������ID�����漰������õ��ݣ���Щ����Ҫһ��ת�����ų�ת��������Ӱ������ж��Ƿ����
  --1.һ�ŷ��õ��ݵ�һ�з��û���з��ÿ��ֶܷ�ν��ʣ��ж����ͬ�Ľ���ID��
  --2.�������Ϻ�Ҳ���ֶܷ�ν���(һ�ŵ��ݶ����ͬ�Ľ���ID)
  --3.�������Ϻ�������������õ���һ���(һ�ŵ��ݵĶ������ID���漰�������NO����ЩNO����֮ǰ�������Ϲ�������������ID)
  --���ǵ�������ĸ����ԣ�Ϊ���߼���������ѯ���ܣ�������ID���ų�(�ò��˵Ľ������ݶ���ת��)

  Update /*+ rule*/ ����Ԥ����¼ L
  Set ��ת�� = n_����
  Where ����id In
        (Select Distinct a.����id --1.�����շѺ͹Һŵ��շѽ����¼
         From ������ü�¼ A
         Where (a.��¼״̬ In (1, 2) Or a.��¼״̬ = 3 And Not Exists
                (Select 1
                 From ������ü�¼ B
                 Where a.No = b.No And a.��¼���� = b.��¼���� And b.��¼״̬ = 2 And b.�Ǽ�ʱ�� >= d_Lastend)) And a.��ת�� Is Null And
               a.��¼���� In (1, 4) And a.����ʱ�� < d_End And a.�Ǽ�ʱ�� < d_Lastend
         Union All
         Select Distinct b.����id --2.ҽ��������(û�з���ʱ���ֶ�,���ϼ�¼�ĵǼ�ʱ�䲻ͬ��Ϊ�˰��շѺ����ϵ�һ����ת��������Ҫ����B��)
         From ���ò����¼ A, ���ò����¼ B
         Where a.��ת�� Is Null And a.No = b.No And a.��¼���� = b.��¼���� And a.�Ǽ�ʱ�� < d_End
         Union All
         Select Distinct a.����id --3.���￨���շѽ����¼(�ų�֮���˿��ѵ�,һ�ŵ�����ֻҪ����һ������)
         From סԺ���ü�¼ A
         Where (a.��¼״̬ In (1, 2) Or a.��¼״̬ = 3 And Not Exists
                (Select 1
                 From סԺ���ü�¼ B
                 Where a.No = b.No And a.��¼���� = b.��¼���� And b.��¼״̬ = 2 And b.�Ǽ�ʱ�� >= d_Lastend)) And a.��ת�� Is Null And
               a.���ʷ��� = 0 And a.��¼���� = 5 And a.����ʱ�� < d_End
         Union All --4.סԺ���ʷ��õĽ��ʽ����¼
         Select ����id
         From (With Settle As (Select Distinct c.����id
                               From (Select Distinct b.No, b.���, Mod(b.��¼����, 10) As ��¼����
                                      From (Select Distinct b.Id
                                             From ���˽��ʼ�¼ A, ���˽��ʼ�¼ B --���ϵĽ��ʵ����շ�ʱ�������ָ��ʱ��֮������Ҫ����B��
                                             Where (a.��¼״̬ In (1, 2) Or a.��¼״̬ = 3 And Not Exists
                                                    (Select 1
                                                     From ���˽��ʼ�¼ C
                                                     Where a.No = c.No And c.��¼״̬ = 2 And c.�շ�ʱ�� >= d_Lastend)) And
                                                   a.��ת�� Is Null And a.No = b.No And (a.�������� = 2 Or Nvl(a.��������, 0) = 0) And
                                                   a.�շ�ʱ�� < d_End) A, סԺ���ü�¼ B
                                      Where a.Id = b.����id) B, סԺ���ü�¼ C --ͨ��C���ҵ���Щ���õ��ݵ����н���IDһ��ת(������ת��ʱ��֮��)
                               Where c.No = b.No And Mod(c.��¼����, 10) = b.��¼���� And c.��� = b.���)
                Select ����id
                From Settle
                Minus
                Select Distinct a.Id
                From ���˽��ʼ�¼ A,
                     (Select Distinct ����id
                       From (Select c.����id, c.No, Mod(c.��¼����, 10) As ��¼����, Nvl(Sum(c.ʵ�ս��), 0) As ʵ�ս��,
                                     Nvl(Sum(c.���ʽ��), 0) As ���ʽ��
                              From סԺ���ü�¼ C, Settle S
                              Where c.����id = s.����id
                              Group By c.No, Mod(c.��¼����, 10), c.����id) C
                       Where c.ʵ�ս�� <> c.���ʽ�� And Exists (Select 1 From ��Ժ���� F Where c.����id = f.����id) --��Ժ����û�н����Ҳת�ߣ�����Ҫʱ�ٳ�أ��������ų���������̫��
                             Or Exists (Select 1
                              From סԺ���ü�¼ E, ���˽��ʼ�¼ S
                              Where e.No = c.No And Mod(e.��¼����, 10) = c.��¼���� And e.����id = s.Id And
                                    s.��ת�� Is Null And s.�շ�ʱ�� >= d_Lastend)) N --��ʹ���ڱ���ת��ʱ��֮����壬ֻҪ����������ת��ʱ��֮�󣬾Ͳ��ų�


                Where a.����id = n.����id And (a.�������� = 2 Or Nvl(a.��������, 0) = 0))
                Union All --5.������ʷ��õĽ��ʽ����¼
                Select ����id
                From (With Settle As (Select Distinct c.����id
                                      From (Select Distinct b.No, b.���, Mod(b.��¼����, 10) As ��¼����
                                             From (Select Distinct b.Id
                                                    From ���˽��ʼ�¼ A, ���˽��ʼ�¼ B
                                                    Where a.��ת�� Is Null And a.No = b.No And (a.�������� = 1 Or Nvl(a.��������, 0) = 0) And
                                                          a.�շ�ʱ�� < d_End) A, ������ü�¼ B
                                             Where a.Id = b.����id) B, ������ü�¼ C
                                      Where c.No = b.No And Mod(c.��¼����, 10) = b.��¼���� And c.��� = b.���)
                       Select ����id
                       From Settle
                       Minus
                       Select Distinct a.Id
                       From ���˽��ʼ�¼ A,
                            (Select Distinct c.����id
                              From (Select c.����id, c.No, Mod(c.��¼����, 10) As ��¼����, Nvl(Sum(c.ʵ�ս��), 0) As ʵ�ս��,
                                            Nvl(Sum(c.���ʽ��), 0) As ���ʽ��
                                     From ������ü�¼ C, Settle S
                                     Where c.����id = s.����id
                                     Group By c.No, Mod(c.��¼����, 10), c.����id) C
                              Where c.ʵ�ս�� <> c.���ʽ�� --���ﲡ��û�н���Ĳ�ת��
                                    Or Exists (Select 1
                                     From ������ü�¼ E, ���˽��ʼ�¼ S
                                     Where e.No = c.No And Mod(e.��¼����, 10) = c.��¼���� And e.����id = s.Id And
                                           s.��ת�� Is Null And s.�շ�ʱ�� >= d_Lastend)) N
                       Where a.����id = n.����id And (a.�������� = 1 Or Nvl(a.��������, 0) = 0))
         );

  --�ų�Ԥ����δ�����
  --Ϊ�˽����߼��ĸ����ԣ����ų���ת��ʱ��֮��ҩ��δ��ҩ�ķ��ü�¼��Ӧ�Ľ���ID������������Ľ������ݺͷ�������ǿ��ת��
  --��Ϊǰ���SQL����Ľ���ID���ܲ�ȫ�ǳ�Ԥ����(�����շѺ�סԺ���ʲ��ѵ�)�����ԣ���Ҫ����һ��SQL���ų�
  Update /*+ rule*/ ����Ԥ����¼
  Set ��ת�� = Null
  Where ��ת�� = n_���� And
        ����id In
        (Select Distinct d.����id --�õ�����ص����г�Ԥ���Ľ���ID����ת��
         From ����Ԥ����¼ D,
              (Select Distinct l.No
                From (Select l.No, l.����id, l.Ԥ�����, Nvl(Sum(l.���), 0) As ���, Nvl(Sum(l.��Ԥ��), 0) As ��Ԥ��,
                              Sum(Decode(l.��ת��, Null, Decode(����ID,Null,Decode(��¼״̬,2,0,1),1), 0)) As δת��
                       From ����Ԥ����¼ L --���ܰ�����IDȷ�ϱ��δ�ת���ĳ��ֻ��ʣ��������Ҫ����L����ԭʼ��Ԥ���ĵ��ݣ��Լ���¼����Ϊ11�Ŀ��ܻ���ת��ʱ��֮��������ʣ���Ľ���ID
                       Where l.��¼���� In (1, 11) And
                             l.No In
                             (Select Distinct p.No From ����Ԥ����¼ P Where p.��¼���� In (1, 11) And p.��ת�� = n_����)
                       Group By l.No, l.����id, l.Ԥ�����) L --���סԺ����һ�ν��壬���ԣ����ܼ���ҳID
                Where δת�� > 0 --ֻҪ��Ԥ�����ݻ���δת����Ԥ�����Ԥ����¼����ת��������ת��һ���ֵ��º����жϴ���
                      Or
                      l.��� <> l.��Ԥ�� And
                      (Exists (Select 1
                               From ����Ԥ����¼ E --ʣ��Ԥ���һ���ø�����Ԥ�����˿NO�Ų�ͬ���������൱���ǳ����ˣ����ų�
                               Where e.����id = l.����id And e.Ԥ����� = l.Ԥ����� And e.��¼���� In (1, 11) And
                                     (e.��ת�� = n_���� Or e.��ת�� Is Null And e.����id Is Null And e.��¼���� = 1 And �տ�ʱ�� < d_End)
                                Having abs(Nvl(Sum(e.���), 0) - Nvl(Sum(e.��Ԥ��), 0)) > n_Ԥ��ʣ�������) --���С�ڵ���n���ų����������3�ֽ���IDΪ�յ�Ҫ����һ��
                       Or l.Ԥ����� = 2 And Exists (Select 1 From ��Ժ���� E Where l.����id = e.����id) Or Exists
                       (Select 1
                        From ����δ����� E
                        Where l.����id = e.����id And (l.Ԥ����� = 1 And e.��ҳid Is Null Or l.Ԥ����� = 2 And e.��ҳid Is Not Null)))) N
         Where d.No = n.No And d.��¼���� In (1, 11));

  --��������3�ֽ���IDΪ�յ�Ԥ����¼
  --1.Ԥ����û��ʹ�þ�ֱ�����˵ļ�¼(����IDΪ��)
  Update /*+ rule*/ ����Ԥ����¼
  Set ��ת�� = n_����
  Where ��¼���� = 1 And
        NO In (Select a.No
               From ����Ԥ����¼ A
               Where a.����id Is Null And a.��¼���� = 1 And a.��¼״̬ In (2, 3) And a.��ת�� Is Null And a.�տ�ʱ�� < d_End
               Group By a.No
               Having Sum(a.���) = 0);

  --2.��Ԥ������˿�ļ�¼������IDΪ�գ���¼״̬Ϊ2��
  Update /*+ rule*/ ����Ԥ����¼
  Set ��ת�� = n_����
  Where ����id Is Null And ��¼���� = 1 And ��¼״̬ = 2 And
        NO In (Select a.No From ����Ԥ����¼ A Where a.��¼���� = 1 And a.��¼״̬ = 3 And a.��ת�� = n_����);

  --�ų�ͬһ��Ԥ����ݲ��ּ�¼�����Ϊת����,ֻҪ�в�ת���ģ������ŵ��ݶ���ת��
  --����2���й���Ӱ�죬����Ҫ������֮��ִ��
  --ҪӰ���3��������жϣ�����Ҫ������֮ǰִ��
  Datamove_Tag_Update(Null, d_End, n_����);

  --3.Ԥ����δ����ʱ�ý�����Ԥ�����˿�(����IDΪ�գ����Ҹ�ԭʼ�ĳ�Ԥ����NOû�й�����ϵ)
  --��������"��� < 0"����Ϊ����Ԥ����û��ʹ�ù�����ֱ���ý�����Ԥ�����˿�����
  Update /*+ rule*/ ����Ԥ����¼ L
  Set ��ת�� = n_����
  Where Exists (Select 1
         From ����Ԥ����¼ E
         Where e.����id = l.����id And e.Ԥ����� = l.Ԥ����� And e.��¼���� In (1, 11) And
               (e.��ת�� = n_���� Or e.��ת�� Is Null And e.����id Is Null And e.��¼���� = 1 And ��¼״̬ = 1 And �տ�ʱ�� < d_End)
         Group By e.����id
         Having abs(Nvl(Sum(e.���), 0) - Nvl(Sum(e.��Ԥ��), 0)) <= n_Ԥ��ʣ�������) --���С�ڵ���nҪת������ǰ�桰�ų�Ԥ����δ����ġ�Ҫ����һ��
        And Exists (Select 1
         From ����Ԥ����¼ E
         Where e.����id = l.����id And e.Ԥ����� = l.Ԥ����� And e.��¼���� In (1, 11) And e.��ת�� = n_����) And
        ��ת�� Is Null And ����id Is Null And ��¼���� = 1 And ��¼״̬ = 1 And �տ�ʱ�� < d_End;

  Update Zldatamovelog
  Set ��ǰ���� = '(1/10)�������ݱ����ɣ����ڱ�Ƿ�������'
  Where ϵͳ = n_System And ���� = n_����;
  Commit;

  Update /*+ rule*/ ���˽��ʼ�¼
  Set ��ת�� = n_����
  Where ID In (Select ����id From ����Ԥ����¼ Where ��ת�� = n_����);

  --�����޽���ļ�¼(Ϊ���������ܣ����жϷ��ã�ֻҪ����������Ԥ����¼�͵���������ý���)
  Update /*+ rule*/ ���˽��ʼ�¼ L
  Set ��ת�� = n_����
  Where �շ�ʱ�� < d_End And ��ת�� Is Null And Not Exists (Select 1 From ����Ԥ����¼ P Where l.Id = p.����id);

  Update /*+ rule*/ ���˿��������
  Set ��ת�� = n_����
  Where Ԥ��id In (Select a.Id From ����Ԥ����¼ A Where ��ת�� = n_����);

  Update /*+ rule*/ ���˿������¼
  Set ��ת�� = n_����
  Where ID In (Select ������id From ���˿�������� Where ��ת�� = n_����);

  Update /*+ rule*/ �������㽻��
  Set ��ת�� = n_����
  Where ����id In (Select a.Id From ����Ԥ����¼ A Where ��ת�� = n_����);

  Update /*+ rule*/ �����˿���Ϣ
  Set ��ת�� = n_����
  Where (��¼id, ����id) In (Select a.Id, a.����id From ����Ԥ����¼ A Where ��ת�� = n_����);

  --1.�Һŷ����쳣����
  --a.����IDΪ�գ�ʵ�ս����ܲ�Ϊ�㣩
  --b.����ID��Ϊ�գ����ۺ�ʵ�ս��Ϊ0��Ӧ�ս�������������ĹҺŷ��ã�û�йҺż�¼��Ҳû��Ԥ����¼
  --������ʱ��ת������Ϊ�պ��˵ķ���ʱ����ͬ���Ǽ�ʱ�䲻ͬ��
  Update /*+ rule*/ ������ü�¼
  Set ��ת�� = n_����
  Where ��ת�� Is Null And ����ʱ�� < d_End And ��¼���� = 4 And (ʵ�ս�� = 0 Or ����id Is Null);

  --2.ֱ���շѵĺͽ����޽��㣨Ԥ������¼�ģ�Union����allȥ���ظ��Լ���in������
  Update /*+ rule*/ ������ü�¼
  Set ��ת�� = n_����
  Where ����id In
        (Select ����id From ����Ԥ����¼ Where ��ת�� = n_���� Union Select ID From ���˽��ʼ�¼ Where ��ת�� = n_����);

  --3.û�н���id������(������ʱ��)
  --a.δ���ʵĻ��ۼ�¼
  --b.δ�շѵ������
  --������"��ת�� Is Null"��Ϊ�˴���������α��ת�������
  Update /*+ rule*/ ������ü�¼ A
  Set ��ת�� = n_����
  Where (��¼״̬ = 0 Or ��¼���� = 1 And ʵ�ս�� = 0 And ���ʽ�� = 0) And ����id Is Null And ��ת�� Is Null And ����ʱ�� < d_End;

  --4.û�н���id������(������ʱ��)
  --δ���ʵ�������ʷ���(����)���ò���û��Ԥ�������Ҳ���������ת��ʱ��֮����δ��������ʷ���
  Update /*+ rule*/ ������ü�¼ A
  Set ��ת�� = n_����
  Where Not Exists (Select 1
         From ����Ԥ����¼ B
         Where b.����id = a.����id And b.��ת�� Is Null And b.Ԥ����� = 1 And b.��¼���� In (1, 11) Having
          Nvl(Sum(b.���), 0) <> Nvl(Sum(b.��Ԥ��), 0)) And Not Exists
   (Select 1
         From ������ü�¼ B
         Where a.����id = b.����id And b.��¼���� = 2 And b.����id Is Null And b.��ת�� Is Null And b.�Ǽ�ʱ�� > = d_Lastend) And
        ��¼���� = 2 And ����id Is Null And ��ת�� Is Null And ����ʱ�� < d_End;

  --5.û�н���id������(������ʱ��)
  --���������ļ��ʼ�¼����¼״̬Ϊ2�����Ǽ�ʱ������ڵ�ǰָ��ת��ʱ��֮�󣬶�ԭʼ���ʼ�¼����¼״̬Ϊ3�����Ǽ�ʱ����ָ��ת��ʱ��֮ǰ��ǰ�����ߵķ���ʱ������ͬ�ġ�
  --a.δ���ʵ�����ʷ��û���ۺ�ʵ�ս��Ϊ��ģ�����ģ�����û�й�ѡ������ý��ʣ�
  --b.�������Ϻ󣬼��ʵ����ʵļ�¼������IDΪ���Ҽ�¼״̬Ϊ2�ģ�����¼״̬Ϊ3�����н���ID������ǰ����ת��.
  Update /*+ rule*/ ������ü�¼ A
  Set ��ת�� = n_����
  Where (Exists (Select 1
                 From ������ü�¼ B
                 Where a.No = b.No And a.��¼���� = b.��¼���� And a.��� = b.��� And b.��¼״̬ = 3 And b.����id Is Not Null And
                       b.��ת�� + 0 = n_����) And ��¼״̬ = 2 Or Exists
         (Select 1
          From ������ü�¼ B
          Where a.No = b.No And a.��¼���� = b.��¼���� And a.��� = b.��� And b.����id Is Null
          Group By b.No, b.��¼����, b.���
          Having Nvl(Sum(b.ʵ�ս��), 0) = 0)) And ��¼���� = 2 And ����id Is Null And ��ת�� Is Null And ����ʱ�� < d_End;

  --6.�н���id�������(������ʱ��)
  --a.���ѱ���ۺ���ʽ��Ϊ����շѼ�¼,
  --b.һ�ŵ�����ͬ����ID�Ľ��ʽ��֮��Ϊ0(������Ϊ��)
  --��ʹ��ת��ʱ��֮��ҩ�ģ�Ҳǿ��ת����Ϊ�˼����߼������ԣ���߲�ѯ���ܣ�
  Update /*+ rule*/ ������ü�¼ A
  Set ��ת�� = n_����
  Where (���ʽ�� = 0 Or Exists
         (Select 1 From ������ü�¼ C Where a.����id = c.����id Group By c.����id, c.No Having Sum(c.���ʽ��) = 0)) And Not Exists
   (Select 1 From ����Ԥ����¼ B Where a.����id = b.����id And b.��ת�� Is Null) And ��¼���� = 1 And ����id Is Not Null And
        ��ת�� Is Null And ����ʱ�� < d_End;

  Update /*+ rule*/ ҽ��������ϸ
  Set ��ת�� = n_����
  Where ����id In (Select ����id From ����Ԥ����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ���ò����¼
  Set ��ת�� = n_����
  Where ����id In (Select ����id From ����Ԥ����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ƾ����ӡ��¼
  Set ��ת�� = n_����
  Where (NO, ��¼����) In (Select NO, ��¼���� From ������ü�¼ Where ��ת�� = n_����);

  --1.��Ԥ����¼����Ϊ��ȡ���￨ֱ���շѵģ��޽���ID��,�ټӽ��ʼ�¼��Ϊ��ȡ�����޽��㣨Ԥ������¼��
  Update /*+ rule*/ סԺ���ü�¼
  Set ��ת�� = n_����
  Where ����id In
        (Select ����id From ����Ԥ����¼ Where ��ת�� = n_���� Union Select ID From ���˽��ʼ�¼ Where ��ת�� = n_����);

  --2.û�н���id������(������ʱ��)
  --���������ļ��ʼ�¼����¼״̬Ϊ2����ԭʼ��¼�ͳ�����¼�ķ���ʱ������ͬ�ġ�
  --1)ת���������Ϻ󣬼��ʵ����ʵļ�¼����¼״̬Ϊ2����û�н���ID����(��¼״̬Ϊ3���н���ID��)����ǰ����ת����
  --2)δ���ʵ������(�ѳ����ļ��ʵ�����ۺ�ʵ�ս��Ϊ��)
  --3)û�н���ID�Ļ��ۼ�¼����Ϊת��
  Update /*+ rule*/ סԺ���ü�¼ A
  Set ��ת�� = n_����
  Where ((Exists (Select 1
                  From סԺ���ü�¼ B
                  Where a.No = b.No And a.��¼���� = b.��¼���� And a.��� = b.��� And b.��¼״̬ = 3 And b.����id Is Not Null And
                        b.��ת�� + 0 = n_����) And ��¼״̬ = 2 Or Exists
         (Select 1
           From סԺ���ü�¼ B
           Where a.No = b.No And a.��¼���� = b.��¼���� And a.��� = b.��� And b.����id Is Null
           Group By b.No, b.��¼����, b.���
           Having Nvl(Sum(b.ʵ�ս��), 0) = 0)) And a.��¼���� In (2, 3, 5) Or a.��¼״̬ = 0) And a.����id Is Null And a.��ת�� Is Null And
        a.����ʱ�� < d_End;

  --3.��Ժδ���ʵģ����ʲ��ˣ�����Ϊ�Ǻܾ���ǰ����Щ���ݣ����Ԥ���ѳ��꣬����ΪҪת��
  --ȥ��������ҳ�е�"����ת�� is null"������������ΪһЩ���˿�����֮ǰ����������ת����
  Update /*+ rule*/ סԺ���ü�¼ A
  Set ��ת�� = n_����
  Where ��ת�� Is Null And ����id Is Null And
        (����id, ��ҳid) In (Select ����id, ��ҳid
                         From ������ҳ C
                         Where ��Ժ���� < d_End And ��ת�� Is Null And Not Exists
                          (Select 1
                                From ����Ԥ����¼ B
                                Where b.����id = c.����id And b.��ת�� Is Null And b.Ԥ����� = 2 And b.��¼���� In (1, 11) Having
                                 Nvl(Sum(b.���), 0) <> Nvl(Sum(b.��Ԥ��), 0)));

  Update Zldatamovelog
  Set ��ǰ���� = '(2/10)�������ݱ����ɣ����ڱ��ҩƷ����'
  Where ϵͳ = n_System And ���� = n_����;
  Commit;

  Update /*+ Rule*/ ҩƷ�շ���¼ A
  Set ��ת�� = n_����
  Where Rowid In (Select m.Rowid
                  From ҩƷ�շ���¼ M, ������ü�¼ E
                  Where m.����id = e.Id And (e.��¼���� = 1 And m.���� In (8, 24) Or e.��¼���� = 2 And m.���� In (9, 25)) And
                        e.�շ���� In ('4', '5', '6', '7') And e.��ת�� = n_����
                  Union All
                  Select m.Rowid
                  From ҩƷ�շ���¼ M, סԺ���ü�¼ E
                  Where m.����id = e.Id And m.���� In (9, 10, 25, 26) And e.��¼���� = 2 And e.�շ���� In ('4', '5', '6', '7') And
                        e.��ת�� = n_����);

  Update /*+ rule*/ �շ���¼������Ϣ
  Set ��ת�� = n_����
  Where �շ�id In (Select ID From ҩƷ�շ���¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ��Һ��ҩ����
  Set ��ת�� = n_����
  Where �շ�id In (Select ID From ҩƷ�շ���¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ��Һ��ҩ��¼
  Set ��ת�� = n_����
  Where ID In (Select ��¼id From ��Һ��ҩ���� Where ��ת�� = n_����);

  Update /*+ rule*/ ��Һ��ҩ����
  Set ��ת�� = n_����
  Where ��ҩid In (Select ID From ��Һ��ҩ��¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ��Һ��ҩ״̬
  Set ��ת�� = n_����
  Where ��ҩid In (Select ID From ��Һ��ҩ��¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ҩƷ����ƻ�
  Set ��ת�� = n_����
  Where ����id In (Select ID From ҩƷ�շ���¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ҩƷǩ����ϸ
  Set ��ת�� = n_����
  Where �շ�id In (Select ID From ҩƷ�շ���¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ҩƷǩ����¼
  Set ��ת�� = n_����
  Where ID In (Select ǩ��id From ҩƷǩ����ϸ Where ��ת�� = n_����);

  Update /*+ rule*/ ҩƷ�շ������־ A
  Set ��ת�� = n_����
  Where (a.������, a.����) In (Select b.No, b.���� From ҩƷ�շ���¼ B Where b.��ת�� = n_����);

  Update /*+ rule*/ ҩƷ�շ�סԺ��־
  Set ��ת�� = n_����
  Where �շ�id In (Select ID From ҩƷ�շ���¼ Where ��ת�� = n_����);

  Update Zldatamovelog
  Set ��ǰ���� = '(3/10)ҩƷ���ݱ����ɣ����ڱ�ǽɿ���Ʊ������'
  Where ϵͳ = n_System And ���� = n_����;
  Commit;

  Update /*+ rule*/ ��Ա����¼ Set ��ת�� = n_���� Where ��ת�� Is Null And ���ʱ�� < d_End;

  Update /*+ rule*/ ��Ա�սɼ�¼ Set ��ת�� = n_���� Where ��ת�� Is Null And �Ǽ�ʱ�� < d_End;

  Update /*+ rule*/ ��Ա�սɶ���
  Set ��ת�� = n_����
  Where �ս�id In (Select ID From ��Ա�սɼ�¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ��Ա�ս���ϸ
  Set ��ת�� = n_����
  Where �ս�id In (Select ID From ��Ա�սɼ�¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ��Ա�ս�Ʊ��
  Set ��ת�� = n_����
  Where �ս�id In (Select ID From ��Ա�սɼ�¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ��Ա�ݴ��¼
  Set ��ת�� = n_����
  Where �ս�id In (Select ID From ��Ա�սɼ�¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ��Ա�ݴ��¼ Set ��ת�� = n_���� Where ��ת�� Is Null And ��¼���� = 1 And �Ǽ�ʱ�� < d_End;

  Update /*+ rule*/ Ʊ�����ü�¼ A
  Set ��ת�� = n_����
  Where Not Exists (Select 1 From Ʊ��ʹ����ϸ B Where b.����id = a.Id And b.ʹ��ʱ�� >= d_Lastend) And ��ת�� Is Null And ʣ������ = 0 And
        �Ǽ�ʱ�� < d_End;

  Update /*+ rule*/ Ʊ��ʹ����ϸ
  Set ��ת�� = n_����
  Where ����id In (Select ID From Ʊ�����ü�¼ Where ��ת�� = n_����);
  Update /*+ rule*/ Ʊ�ݴ�ӡ����
  Set ��ת�� = n_����
  Where ID In (Select ��ӡid From Ʊ��ʹ����ϸ Where ��ת�� = n_����);

  Update /*+ rule*/ Ʊ�ݴ�ӡ��ϸ
  Set ��ת�� = n_����
  Where ʹ��id In (Select ID From Ʊ��ʹ����ϸ Where ��ת�� = n_����);

  Update Zldatamovelog
  Set ��ǰ���� = '(4/10)�ɿ���Ʊ�����ݱ����ɣ����ڱ�Ǿ��Ｐ��������'
  Where ϵͳ = n_System And ���� = n_����;
  Commit;

  --2.���Ｐ��������
  --��ת�����������Һŷ���δת���ģ�����ת��ʱ��֮�����ҽ������Щҽ����Ϊʱ��û�е�����Ӧת������ҽ����Ӧ�ķ���δת����
  --��ʹ���ھ���(r.ִ��״̬ <> 2 )��Ҳǿ��ת��(ҽ������û��ʹ����ɾ��﹦��)
  Update /*+ rule*/ ���˹Һż�¼ T
  Set ��ת�� = n_����
  Where Rowid In
        (Select Rowid
         From ���˹Һż�¼ R
         Where Not Exists (Select 1 From ������ü�¼ A Where r.No = a.No And a.��¼���� = 4 And a.��ת�� Is Null) And Not Exists
          (Select 1
                From ����ҽ����¼ A
                Where a.�Һŵ� = r.No And a.��ת�� Is Null And a.������Դ <> 4 And Nvl(a.ͣ��ʱ��, a.����ʱ��) >= d_Lastend) And
               Not Exists (Select 1
                From ������ü�¼ E, ����ҽ����¼ A
                Where r.No = a.�Һŵ� And a.Id = e.ҽ����� And a.������Դ <> 4 And e.��ת�� Is Null) And
               r.��ת�� Is Null And r.����ʱ�� < d_End);

  --������һ���ֹҺ�����δת�������ԣ����ܱ�����ݿ�����Һ����ݲ�ƥ��
  Update ���˹ҺŻ��� Set ��ת�� = n_���� Where ��ת�� Is Null And ���� < d_End;
  Update /*+ rule*/ ����ת���¼ Set ��ת�� = n_���� Where NO In (Select NO From ���˹Һż�¼ Where ��ת�� = n_����);

  --ͨ��"סԺ���ü�¼"����ѯ��������"���˽��ʼ�¼",��Ϊ��Ժδ������ʲ���Ҳת���˷���
  --��Ժ����������Ȼ��Ҫ����Ϊ����ĳ�ν���ת���ˣ�������������ת����ֹʱ��֮ǰ��δ��Ժ(һ��סԺ��ν���)��
  --ͨ��ָ��������ʽ���������Ż���ȱʡ����"������ҳIX_��Ժ����"������Ч��̫�ͣ�
  --����"����ת�� is null"����������Ϊһ��סԺ��ν���ʱ������粻ͬ��ת������(ת����ֹʱ��)�����ֶν��ᱻ���¶�Ρ�
  Update /*+ rule*/ ������ҳ P
  Set ��ת�� = n_����
  Where Not Exists
   (Select 1 From סԺ���ü�¼ A Where a.����id = p.����id And a.��ҳid = p.��ҳid And a.��ת�� Is Null) And ��ת�� Is Null And
        ��Ժ���� < d_Lastend And (����id, ��ҳid) In (Select Distinct ����id, ��ҳid From סԺ���ü�¼ Where ��ת�� = n_����);

  --�ѳ�Ժ����û�з��õģ�Ҳ���Ϊת�����Ա�ת����������
  Update /*+ rule*/ ������ҳ P
  Set ��ת�� = n_����
  Where Not Exists (Select 1 From סԺ���ü�¼ A Where a.����id = p.����id And a.��ҳid = p.��ҳid) And ��ת�� Is Null And ����ת�� Is Null And
        ��Ժ���� < d_End;

  Update /*+ rule*/ ���˹�����¼
  Set ��ת�� = n_����
  Where (����id, ��ҳid) In (Select ����id, ID
                         From ���˹Һż�¼
                         Where ��ת�� = n_����
                         Union All
                         Select ����id, ��ҳid
                         From ������ҳ
                         Where ��ת�� = n_����);

  Update /*+ rule*/ ������ϼ�¼
  Set ��ת�� = n_����
  Where (����id, ��ҳid) In (Select ����id, ID
                         From ���˹Һż�¼
                         Where ��ת�� = n_����
                         Union All
                         Select ����id, ��ҳid
                         From ������ҳ
                         Where ��ת�� = n_����);

  Update /*+ rule*/ ���������¼
  Set ��ת�� = n_����
  Where (����id, ��ҳid) In (Select ����id, ID
                         From ���˹Һż�¼
                         Where ��ת�� = n_����
                         Union All
                         Select ����id, ��ҳid
                         From ������ҳ
                         Where ��ת�� = n_����);

  Update Zldatamovelog
  Set ��ǰ���� = '(5/10)���Ｐ�������ݱ����ɣ����ڱ�ǻ�������'
  Where ϵͳ = n_System And ���� = n_����;
  Commit;

  --3.��������
  Update /*+ rule*/ ���˻����ļ�
  Set ��ת�� = n_����
  Where (����id, ��ҳid) In (Select ����id, ��ҳid From ������ҳ Where ��ת�� = n_����);
  Update /*+ rule*/ ���˻�������
  Set ��ת�� = n_����
  Where �ļ�id In (Select ID From ���˻����ļ� Where ��ת�� = n_����);
  Update /*+ rule*/ ���˻�����ϸ
  Set ��ת�� = n_����
  Where ��¼id In (Select ID From ���˻������� Where ��ת�� = n_����);

  Update /*+ rule*/ ���˻����ӡ
  Set ��ת�� = n_����
  Where �ļ�id In (Select ID From ���˻����ļ� Where ��ת�� = n_����);
  Update /*+ rule*/ ���˻�����Ŀ
  Set ��ת�� = n_����
  Where �ļ�id In (Select ID From ���˻����ļ� Where ��ת�� = n_����);
  Update /*+ rule*/ ���˻���Ҫ������
  Set ��ת�� = n_����
  Where �ļ�id In (Select ID From ���˻����ļ� Where ��ת�� = n_����);
  Update /*+ rule*/ ����Ҫ������
  Set ��ת�� = n_����
  Where �ļ�id In (Select ID From ���˻����ļ� Where ��ת�� = n_����);

  --�ϰ滤��ϵͳ����
  Update /*+ rule*/ ���˻����¼
  Set ��ת�� = n_����
  Where (����id, ��ҳid) In (Select ����id, ��ҳid From ������ҳ Where ��ת�� = n_����);
  Update /*+ rule*/ ���˻�������
  Set ��ת�� = n_����
  Where ��¼id In (Select ID From ���˻����¼ Where ��ת�� = n_����);

  Update Zldatamovelog
  Set ��ǰ���� = '(6/10)�������ݱ����ɣ����ڱ�ǲ�������'
  Where ϵͳ = n_System And ���� = n_����;
  Commit;

  --4.��������
  Update /*+ rule*/ ���Ӳ�����¼
  Set ��ת�� = n_����
  Where ������Դ = 1 And (����id, ��ҳid) In (Select ����id, ID From ���˹Һż�¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ���Ӳ�����¼
  Set ��ת�� = n_����
  Where ������Դ = 2 And (����id, ��ҳid) In (Select ����id, ��ҳid From ������ҳ Where ��ת�� = n_����);

  --�ԵǼ��ಡ��(�޹Һŵ���)
  --����ID�����ظ�����Ϊ���鱨��֮��ģ���ι�����������һ�ű��棬���ڲ���ҽ��������У����ҽ��id��Ӧͬһ����ID
  --Ϊ�������ܣ�����ҽ�����ͼ�¼�ķ���ʱ���ѯ�������þ�ȷ��ʱ�䣬��Ϊֱ�ӵǼǵļ���ҽ����һ�㿪��ʱ���뷢��ʱ������
  --��Щ���⣨�������ݣ��Һŵ�Ϊ�յ�ҽ����������ԴΪ3�ģ�ֱ�ӵǼǵļ�����ҽ����������������ԴΪ1��4�ģ���������ҽ��������ҳID���ܲ���0
  Update /*+ rule*/ ���Ӳ�����¼
  Set ��ת�� = n_����
  Where ��ת�� Is Null And �������� = 7 And ID In (Select c.����id
               From ����ҽ����¼ B, ����ҽ������ C
               Where c.ҽ��id = b.Id And b.������Դ<>2 And b.�Һŵ� Is Null And b.���id Is Null And b.��ת�� Is Null And
                     b.����ʱ�� < d_End);

  Update /*+ rule*/ ���Ӳ�������
  Set ��ת�� = n_����
  Where ����id In (Select ID From ���Ӳ�����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ���Ӳ�����ʽ
  Set ��ת�� = n_����
  Where �ļ�id In (Select ID From ���Ӳ�����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ���Ӳ�������
  Set ��ת�� = n_����
  Where �ļ�id In (Select ID From ���Ӳ�����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ���Ӳ���ͼ��
  Set ��ת�� = n_����
  Where ����id In (Select ID From ���Ӳ������� Where ��ת�� = n_���� And �������� = 5);

  Update /*+ rule*/ ����ҽ������
  Set ��ת�� = n_����
  Where ����id In (Select ID From ���Ӳ�����¼ Where �������� = 7 And ��ת�� = n_����);
  Update /*+ rule*/ Ӱ�񱨸沵��
  Set ��ת�� = n_����
  Where (ҽ��id, ����id) In (Select ҽ��id, ����id From ����ҽ������ Where ��ת�� = n_����);
  Update /*+ rule*/ ҽ����������
  Set ��ת�� = n_����
  Where ID In (Select ����id From ����ҽ������ Where ��ת�� = n_����);

  Update /*+ rule*/ ������ļ�¼
  Set ��ת�� = n_����
  Where ����id In (Select ID From ���Ӳ�����¼ Where �������� = 7 And ��ת�� = n_����);

  Update /*+ rule*/ �����걨��¼
  Set ��ת�� = n_����
  Where �ļ�id In (Select ID From ���Ӳ�����¼ Where �������� = 5 And ��ת�� = n_����);

  Update /*+ rule*/ �������淴��
  Set ��ת�� = n_����
  Where �ļ�id In (Select ID From ���Ӳ�����¼ Where �������� = 5 And ��ת�� = n_����);

  Update /*+ rule*/ �����걨����
  Set ��ת�� = n_����
  Where �걨id In (Select ID From ���Ӳ�����¼ Where �������� = 5 And ��ת�� = n_����);

  Update /*+ rule*/ ������ϼ�¼
  Set ��ת�� = n_����
  Where ����id In (Select ID From ���Ӳ�����¼ Where ��ת�� = n_����);

  Update Zldatamovelog
  Set ��ǰ���� = '(7/10)�������ݱ����ɣ����ڱ���ٴ�·������'
  Where ϵͳ = n_System And ���� = n_����;
  Commit;

  --5.�ٴ�·��
  Update /*+ rule*/ �����ٴ�·��
  Set ��ת�� = n_����
  Where (����id, ��ҳid) In (Select ����id, ��ҳid From ������ҳ Where ��ת�� = n_����);
  Update /*+ rule*/ ���˺ϲ�·��
  Set ��ת�� = n_����
  Where ��Ҫ·����¼id In (Select ID From �����ٴ�·�� Where ��ת�� = n_����);
  Update /*+ rule*/ ���˺ϲ�·������
  Set ��ת�� = n_����
  Where ·����¼id In (Select ID From �����ٴ�·�� Where ��ת�� = n_����);

  Update /*+ rule*/ ���˳�����¼
  Set ��ת�� = n_����
  Where ·����¼id In (Select ID From �����ٴ�·�� Where ��ת�� = n_����);

  Update /*+ rule*/ ����·��ִ��
  Set ��ת�� = n_����
  Where ·����¼id In (Select ID From �����ٴ�·�� Where ��ת�� = n_����);
  Update /*+ rule*/ ����·������
  Set ��ת�� = n_����
  Where ·����¼id In (Select ID From �����ٴ�·�� Where ��ת�� = n_����);
  Update /*+ rule*/ ����·������
  Set ��ת�� = n_����
  Where ·����¼id In (Select ID From �����ٴ�·�� Where ��ת�� = n_����);
  Update /*+ rule*/ ����·��ָ��
  Set ��ת�� = n_����
  Where ·����¼id In (Select ID From �����ٴ�·�� Where ��ת�� = n_����);
  Update /*+ rule*/ ����·��ҽ��
  Set ��ת�� = n_����
  Where ·��ִ��id In (Select ID From ����·��ִ�� Where ��ת�� = n_����);

  Update Zldatamovelog
  Set ��ǰ���� = '(8/10)�ٴ�·�����ݱ����ɣ����ڱ��ҽ������'
  Where ϵͳ = n_System And ���� = n_����;
  Commit;

  --6.ҽ�������飬���
  --���ϲ�����Դ��������ԴΪ3���ԵǼ��ಡ�������˹Һŵ���ҽ����ת���˶�ҽ������û��ת��
  Update /*+ rule*/ ����ҽ����¼
  Set ��ת�� = n_����
  Where �Һŵ� In (Select NO From ���˹Һż�¼ Where ��ת�� = n_����) And ������Դ =1;

  --���ϲ�����Դ������ ��ҽ����ת���˶�ҽ������û��ת��
  Update /*+ rule*/ ����ҽ����¼
  Set ��ת�� = n_����
  Where (����id, ��ҳid) In (Select ����id, ��ҳid From ������ҳ Where ��ת�� = n_����) And ������Դ = 2;

  --�ԵǼ��ಡ��(�޹Һŵ�)������ҽ��������ǰ��ת����ʱ��ת��
  Update /*+ rule*/ ����ҽ����¼
  Set ��ת�� = n_����
  Where Rowid In (Select b.Rowid
                  From ����ҽ����¼ B, ����ҽ������ C
                  Where (b.���id = c.ҽ��id Or b.Id = c.ҽ��id) And c.��ת�� = n_����);

  --�ԵǼ��ಡ��(�޹Һŵ�)��û��ҽ������
  Update /*+ rule*/ ����ҽ����¼ A
  Set ��ת�� = n_����
  Where Not Exists (Select 1 From ����ҽ������ B Where a.Id = b.ҽ��id) And Not Exists
     (Select 1 From ����ҽ������ B Where a.���id = b.ҽ��id) And �Һŵ� Is Null And ������Դ = 3 And ��ת�� Is Null And ����ʱ�� < d_End;

  Update /*+ rule*/ ����ҽ���Ƽ�
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ����ҽ������
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ����ҽ������
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ��Ѫ�����¼
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ��Ѫ������
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ��Ѫ������Ŀ
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ����ҽ��ִ��
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ����ҽ����ӡ
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ҽ��ִ�д�ӡ
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ �������ҽ��
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ������ϼ�¼
  Set ��ת�� = n_����
  Where ID In (Select ���id From �������ҽ�� Where ��ת�� = n_����);

  Update /*+ rule*/ ����ҽ��״̬
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ҽ��ǩ����¼
  Set ��ת�� = n_����
  Where ID In (Select ǩ��id From ����ҽ��״̬ Where ��ת�� = n_���� And ǩ��id Is Not Null);

  Update /*+ rule*/ ����ҽ������
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ���Ƶ��ݴ�ӡ
  Set ��ת�� = n_����
  Where (NO, ��¼����) In (Select NO, ��¼���� From ����ҽ������ Where ��ת�� = n_����);

  Update /*+ rule*/ ҽ��ִ��ʱ��
  Set ��ת�� = n_����
  Where (ҽ��id, ���ͺ�) In (Select ҽ��id, ���ͺ� From ����ҽ������ Where ��ת�� = n_����);

  Update /*+ rule*/ ҽ��ִ�мƼ�
  Set ��ת�� = n_����
  Where (ҽ��id, ���ͺ�) In (Select ҽ��id, ���ͺ� From ����ҽ������ Where ��ת�� = n_����);

  Update /*+ rule*/ ִ�д�ӡ��¼
  Set ��ת�� = n_����
  Where (ҽ��id, ���ͺ�) In (Select ҽ��id, ���ͺ� From ����ҽ������ Where ��ת�� = n_����);

  Update /*+ rule*/ ���������ϸ
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ��������¼
  Set ��ת�� = n_����
  Where ID In (Select ��id From ���������ϸ Where ��ת�� = n_����);

  Update /*+ rule*/ ���������
  Set ��ת�� = n_����
  Where ��id In (Select ID From ��������¼ Where ��ת�� = n_����);

  Update /*+ rule*/ Ris���ԤԼ
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ �������Լ�¼
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ҽ�����뵥�ļ�
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  
  Update /*+ rule*/ ����Σ��ֵ��¼
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ����Σ��ֵ����
  Set ��ת�� = n_����
  Where Σ��ֵid In (Select ID From ����Σ��ֵ��¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ����Σ��ֵҽ��
  Set ��ת�� = n_����
  Where Σ��ֵid In (Select ID From ����Σ��ֵ��¼ Where ��ת�� = n_����);

  Update Zldatamovelog
  Set ��ǰ���� = '(9/10)ҽ�����ݱ����ɣ����ڱ�Ǽ���������'
  Where ϵͳ = n_System And ���� = n_����;
  Commit;

  Update /*+ rule*/ Ӱ�����¼
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ Ӱ�񱨸��¼
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ Ӱ�񱨸������¼
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ Ӱ��������
  Set ��ת�� = n_����
  Where ���uid In (Select ���uid From Ӱ�����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ Ӱ����ͼ��
  Set ��ת�� = n_����
  Where ����uid In (Select ����uid From Ӱ�������� Where ��ת�� = n_����);

  Update /*+ rule*/ Ӱ�����뵥ͼ��
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ Ӱ���ղ�����
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ Ӱ��Σ��ֵ��¼
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);

  Update Zldatamovelog
  Set ��ǰ���� = '(10/10)Ӱ�����ݱ����ɣ����ڱ�Ǽ�������'
  Where ϵͳ = n_System And ���� = n_����;
  Commit;

  Update /*+ rule*/ ����걾��¼
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ����������Ŀ
  Set ��ת�� = n_����
  Where �걾id In (Select ID From ����걾��¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ������Ŀ�ֲ�
  Set ��ת�� = n_����
  Where �걾id In (Select ID From ����걾��¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ���������¼
  Set ��ת�� = n_����
  Where �걾id In (Select ID From ����걾��¼ Where ��ת�� = n_����);
  Update /*+ rule*/ �����ʿؼ�¼
  Set ��ת�� = n_����
  Where �걾id In (Select ID From ����걾��¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ���������¼
  Set ��ת�� = n_����
  Where �걾id In (Select ID From ����걾��¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ����ǩ����¼
  Set ��ת�� = n_����
  Where ����걾id In (Select ID From ����걾��¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ����ͼ����
  Set ��ת�� = n_����
  Where �걾id In (Select ID From ����걾��¼ Where ��ת�� = n_����);

  Update /*+ rule*/ �����Լ���¼
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ������ռ�¼
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);

  Update /*+ rule*/ ������ͨ���
  Set ��ת�� = n_����
  Where ����걾id In (Select ID From ����걾��¼ Where ��ת�� = n_����);
  Update /*+ rule*/ �����ʿر���
  Set ��ת�� = n_����
  Where ���id In (Select ID From ������ͨ��� Where ��ת�� = n_����);
  Update /*+ rule*/ ����ҩ�����
  Set ��ת�� = n_����
  Where ϸ�����id In (Select ID From ������ͨ��� Where ��ת�� = n_����);
  Update /*+ rule*/ ������ˮ�߱걾
  Set ��ת�� = n_����
  Where �걾id In (Select ID From ����걾��¼ Where ��ת�� = n_����);
  Update /*+ rule*/ ������ˮ��ָ��
  Set ��ת�� = n_����
  Where �걾id In (Select ID From ����걾��¼ Where ��ת�� = n_����);

  Commit;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl1_Datamove_Tag;
/

--124292:����,2018-04-19,����ҽ�����ʻ��˺���Һ��ҩ��¼�еĲ���״̬δ�ı�����
Create Or Replace Procedure Zl_��Һ��ҩ��¼_ҽ������
(
  ҽ��id_In   In ��Һ��ҩ��¼.ҽ��id%Type,
  ���ͺ�_In   In ��Һ��ҩ��¼.���ͺ�%Type,
  ������Ա_In In ��Һ��ҩ��¼.������Ա%Type := Null,
  ����ʱ��_In In ��Һ��ҩ��¼.����ʱ��%Type := Null
) Is
  n_Count Number(5);
Begin
  --ֻ��״̬=1(δ��ҩ)�ļ�¼��������Ѿ���ҩ�ˣ���ͨ�����˷�ʽ����
  Select Count(ID) Into n_Count From ��Һ��ҩ��¼ Where ����״̬ in (1,10)  And ҽ��id = ҽ��id_In And ���ͺ� = ���ͺ�_In;

  If n_Count > 0 Then
    Update ��Һ��ҩ��¼
    Set ����״̬ = 12, ������Ա = Nvl(������Ա_In, Zl_Username), ����ʱ�� = Nvl(����ʱ��_In, Sysdate)
    Where ҽ��id = ҽ��id_In And ���ͺ� = ���ͺ�_In;
  
    Insert Into ��Һ��ҩ״̬
      (��ҩid, ��������, ������Ա, ����ʱ��)
      Select ID, 12, Nvl(������Ա_In, Zl_Username), Nvl(����ʱ��_In, Sysdate)
      From ��Һ��ҩ��¼
      Where ҽ��id = ҽ��id_In And ���ͺ� = ���ͺ�_In;
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Һ��ҩ��¼_ҽ������;
/

--120692:������,2018-04-17,�����¼֧�ּ�����Ŀ����
Create Or Replace Procedure Zl_�������ݵ��붨��_Update
(
  ���_In �������ݵ��붨��.���%Type,
  ����_In �������ݵ��붨��.����%Type,
  ��ʽ_In �������ݵ��붨��.��ʽ%Type
) Is
Begin
  Update �������ݵ��붨�� Set ���� = ����_In, ��ʽ = ��ʽ_In Where ��� = ���_In;
  If Sql%Rowcount = 0 Then
    Insert Into �������ݵ��붨�� (���, ����, ��ʽ) Values (���_In, ����_In, ��ʽ_In);
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�������ݵ��붨��_Update;
/
--123732:������,2018-04-09,�޸���˱걾֮��ҽ��ִ���˲���������
Create Or Replace Procedure Zl_����걾��¼_�������
(
  Id_In       ����걾��¼.Id%Type,
  �����_In   ����걾��¼.�����%Type := Null,
  ��Ա���_In ��Ա��.���%Type := Null,
  ��Ա����_In ��Ա��.����%Type := Null
) Is

  --δ��˵ķ�����(������ҩƷ) 
  Cursor c_Verify(v_ҽ��id In Number) Is
    Select Distinct 2 As ��¼����, NO, ���, ��¼״̬, �����־
    From סԺ���ü�¼
    Where �շ���� Not In ('5', '6', '7') And ���ʷ��� = 1 And �۸񸸺� Is Null And
          (��¼����, NO) In (Select ��¼����, NO
                         From ����ҽ������
                         Where ҽ��id = v_ҽ��id
                         Union All
                         Select ��¼����, NO
                         From ����ҽ������
                         Where ҽ��id In (Select ID From ����ҽ����¼ Where v_ҽ��id In (ID, ���id))) And ҽ����� = v_ҽ��id
    Union All
    Select Distinct 1 As ��¼����, NO, ���, ��¼״̬,�����־
    From ������ü�¼
    Where �շ���� Not In ('5', '6', '7') And ���ʷ��� = 1 And �۸񸸺� Is Null And
          (��¼����, NO) In (Select ��¼����, NO
                         From ����ҽ������
                         Where ҽ��id = v_ҽ��id
                         Union All
                         Select ��¼����, NO
                         From ����ҽ������
                         Where ҽ��id In (Select ID From ����ҽ����¼ Where v_ҽ��id In (ID, ���id))) And ҽ����� = v_ҽ��id
    Order By ��¼����, NO, ���;

  --���ҵ�ǰ�걾��������� 
  Cursor c_Samplequest(v_΢���� In Number) Is
    Select Distinct ҽ��id, ������Դ
    From (Select a.ҽ��id, b.������Դ
           From ������Ŀ�ֲ� A, ����걾��¼ B
           Where 0 = v_΢���� And a.�걾id = Id_In And a.ҽ��id Is Not Null And a.�걾id = b.Id
           Union
           Select a.ҽ��id, b.������Դ
           From ������Ŀ�ֲ� A, ����걾��¼ B
           Where 1 = v_΢���� And a.�걾id = Id_In And a.ҽ��id Is Not Null And a.�걾id = b.Id
           Union
           Select b.Id As ҽ��id, a.������Դ
           From ����걾��¼ A, ����ҽ����¼ B
           Where a.Id = Id_In And a.ҽ��id = b.���id);

  Cursor c_Stuff
  (
    v_No     Varchar2,
    v_��ҳid Number
  ) Is
    Select NO, ����, �ⷿid
    From δ��ҩƷ��¼
    Where NO = v_No And ���� In (24, 25, 26) And �ⷿid Is Not Null And Not Exists
     (Select 1 From Dual Where zl_GetSysParameter(Decode(v_��ҳid, Null, 92, 63)) = '1') And Exists
     (Select a.���
           From סԺ���ü�¼ A, �������� B
           Where a.��¼���� = 2 And a.��¼״̬ = 1 And a.No = v_No And a.�շ�ϸĿid = b.����id And b.�������� = 1
           Union All
           Select a.���
           From ������ü�¼ A, �������� B
           Where a.��¼���� = 2 And a.��¼״̬ = 1 And a.No = v_No And a.�շ�ϸĿid = b.����id And b.�������� = 1)
    Order By �ⷿid;

  v_ִ��  Number(1);
  v_No    ����ҽ������.No%Type;
  v_Nonew ����ҽ������.No%Type;
  v_����  ����ҽ������.��¼����%Type;
  v_���  Varchar2(1000);

  v_Count      Number(18);
  v_Counts     Number(18);
  v_΢����걾 Number(1) := 0;
  v_��ҳid     Number(18);
  v_Ӥ��       Number(1);
  v_����       Varchar2(100);
  v_����       Number(18);
  v_Intloop    Number;
  Err_Custom Exception;
  v_Error Varchar2(100);

  n_Par Number;
Begin
  Select Nvl(Ӥ��, 0), ���� Into v_Ӥ��, v_���� From ����걾��¼ Where ID = Id_In;

  --ִ�к��Զ���˶�Ӧ�ļ��ʻ��۵�(������ҩƷ) 
  Select Zl_To_Number(Nvl(zl_GetSysParameter(81), '0')) Into v_ִ�� From Dual;

  v_΢����걾 := 0;
  Begin
    Select 1 Into v_΢����걾 From ����걾��¼ Where ΢����걾 = 1 And ID = Id_In;
  Exception
    When Others Then
      v_΢����걾 := 0;
  End;

  --���ж�ҽ���Ƿ��շ�
  n_Par := Zl_To_Number(Nvl(zl_GetSysParameter(163), '0'));
  If n_Par = 1 Then
    For r_Samplequest In c_Samplequest(v_΢����걾) Loop
      For r_���ҽ�� In (Select ID As ҽ��id From ����ҽ����¼ Where ���id = r_Samplequest.ҽ��id) Loop
        For r_Verify In c_Verify(r_���ҽ��.ҽ��id) Loop
          If r_Verify.��¼״̬ = 0 Then
            If r_Verify.�����־ = 1 Then
              v_Error := '�걾δ�շѣ�������ִ�У�����ϵ����Ա��';
              Raise Err_Custom;
            Elsif r_Verify.�����־ = 2 Then
              v_Error := '�걾δ���ˣ�������ִ�У�����ϵ����Ա��';
              Raise Err_Custom;
            End If;
          End If;
        End Loop;
      End Loop;
    End Loop;
  End If;

  --1.�ñ��걾��״̬������˺�ʱ�� 
  Update ����걾��¼
  Set ����� = Decode(�����_In, Null, ��Ա����_In, �����_In), ���ʱ�� = Sysdate, ����״̬ = 2
  Where ID = Id_In;

  --��¼��˹��� 
  Insert Into ���������¼
    (ID, �걾id, ��������, ����Ա, ����ʱ��)
  Values
    (���������¼_Id.Nextval, Id_In, 0, Decode(�����_In, Null, ��Ա����_In, �����_In), Sysdate);

  --2.��鵱ǰ�걾��ص��������ر걾�Ƿ������� 
  For r_Samplequest In c_Samplequest(v_΢����걾) Loop
  
    v_Count := 0;
  
    If v_΢����걾 = 0 Then
      Begin
        Select Nvl(Count(1), 0)
        Into v_Count
        From ����걾��¼
        Where ����״̬ < 2 And ID In (Select �걾id From ������Ŀ�ֲ� Where ҽ��id = r_Samplequest.ҽ��id);
      Exception
        When Others Then
          v_Count := 0;
      End;
    End If;
  
    --r_SampleQuest.ҽ��id�����Ѿ����,����������� 
    If v_Count = 0 Then
    
      --1.�����뵥��ִ��״̬ 
      Update ����ҽ������
      Set ִ��״̬ = 1, ����� = Decode(�����_In, Null, ��Ա����_In, �����_In), ���ʱ�� = Sysdate
      Where ҽ��id In (Select ID From ����ҽ����¼ Where r_Samplequest.ҽ��id In (ID, ���id));
    
      Update ����ҽ������
      Set ִ��״̬ = 1, ����� = ��Ա����_In, ���ʱ�� = Sysdate
      Where ҽ��id In (Select ���id
                     From ����ҽ����¼
                     Where ID In (Select ID From ����ҽ����¼ Where r_Samplequest.ҽ��id In (ID, ���id)));
    
      If r_Samplequest.������Դ = 2 Then
        --2.����ִ�д��� 
        Update סԺ���ü�¼
        Set ִ��״̬ = 1, ִ��ʱ�� = Sysdate, ִ���� = ��Ա����_In
        Where �շ���� Not In ('5', '6', '7') And
              (ҽ�����, ��¼����, NO) In
              (Select ҽ��id, ��¼����, NO
               From ����ҽ������
               Where ҽ��id = r_Samplequest.ҽ��id
               Union All
               Select ҽ��id, ��¼����, NO
               From ����ҽ������
               Where ҽ��id In (Select ID From ����ҽ����¼ Where r_Samplequest.ҽ��id In (ID, ���id)));
      Else
        Update ������ü�¼
        Set ִ��״̬ = 1, ִ��ʱ�� = Sysdate, ִ���� = ��Ա����_In
        Where �շ���� Not In ('5', '6', '7') And
              (ҽ�����, ��¼����, NO) In
              (Select ҽ��id, ��¼����, NO
               From ����ҽ������
               Where ҽ��id = r_Samplequest.ҽ��id
               Union All
               Select ҽ��id, ��¼����, NO
               From ����ҽ������
               Where ҽ��id In (Select ID From ����ҽ����¼ Where r_Samplequest.ҽ��id In (ID, ���id)));
      End If;
    
      --3.�Զ���˼��� 
      If v_ִ�� = 1 Then
        Select Count(*) Into v_Counts From ����ҽ����¼ Where ���id = r_Samplequest.ҽ��id;
        If v_Counts > 0 Then
          For r_���ҽ�� In (Select ID As ҽ��id From ����ҽ����¼ Where ���id = r_Samplequest.ҽ��id) Loop
            For r_Verify In c_Verify(r_���ҽ��.ҽ��id) Loop
              If r_Verify.No || ',' || r_Verify.��¼���� <> v_No || ',' || v_���� Then
                If v_��� Is Not Null Then
                  If v_���� = 1 Then
                    Zl_������ʼ�¼_Verify(v_No, ��Ա���_In, ��Ա����_In, Substr(v_���, 2));
                  Elsif v_���� = 2 Then
                    Zl_סԺ���ʼ�¼_Verify(v_No, ��Ա���_In, ��Ա����_In, Substr(v_���, 2));
                  End If;
                End If;
                v_��� := Null;
              End If;
              v_No   := r_Verify.No;
              v_���� := r_Verify.��¼����;
              v_��� := v_��� || ',' || r_Verify.���;
            End Loop;
          End Loop;
        Else
          For r_Verify In c_Verify(r_Samplequest.ҽ��id) Loop
            If r_Verify.No || ',' || r_Verify.��¼���� <> v_No || ',' || v_���� Then
              If v_��� Is Not Null Then
                If v_���� = 1 Then
                  Zl_������ʼ�¼_Verify(v_No, ��Ա���_In, ��Ա����_In, Substr(v_���, 2));
                Elsif v_���� = 2 Then
                  Zl_סԺ���ʼ�¼_Verify(v_No, ��Ա���_In, ��Ա����_In, Substr(v_���, 2));
                End If;
              End If;
              v_��� := Null;
            End If;
            v_No   := r_Verify.No;
            v_���� := r_Verify.��¼����;
            v_��� := v_��� || ',' || r_Verify.���;
          End Loop;
        End If;
        If v_��� Is Not Null Then
          If v_���� = 1 Then
            Zl_������ʼ�¼_Verify(v_No, ��Ա���_In, ��Ա����_In, Substr(v_���, 2));
          Elsif v_���� = 2 Then
            Zl_סԺ���ʼ�¼_Verify(v_No, ��Ա���_In, ��Ա����_In, Substr(v_���, 2));
          End If;
          v_��� := Null;
          --  v_���� := null; 
        End If;
      End If;
    
      --����Լ����ĵ� 
      v_Intloop := 1;
    
      Select ����id Into v_���� From ����걾��¼ Where ID = Id_In;
      For r_�����Լ� In (Select c.����id, c.����
                     From ����ҽ����¼ A, ���鱨����Ŀ B, �����Լ���ϵ C
                     Where a.���id = r_Samplequest.ҽ��id And a.������Ŀid = b.������Ŀid And b.������Ŀid = c.��Ŀid And c.����id = v_����) Loop
        Zl_�����Լ���¼_Insert(r_Samplequest.ҽ��id, v_Intloop, r_�����Լ�.����id, r_�����Լ�.����);
        v_Intloop := v_Intloop + 1;
      End Loop;
      Select Count(*) Into v_Intloop From �����Լ���¼ Where ҽ��id = r_Samplequest.ҽ��id And NO Is Null;
      If v_Intloop > 1 Then
        v_Nonew := Nextno(14);
        Update �����Լ���¼ Set NO = v_Nonew Where ҽ��id = r_Samplequest.ҽ��id;
      End If;
      If v_Nonew Is Not Null Then
      
        Zl_�����Լ���¼_Bill(r_Samplequest.ҽ��id, v_Nonew);
      
        v_��ҳid := Null;
        Select ��ҳid Into v_��ҳid From ����ҽ����¼ A Where ID = r_Samplequest.ҽ��id;
      
        If v_��ҳid Is Null Then
          Zl_������ʼ�¼_Verify(v_Nonew, ��Ա���_In, ��Ա����_In);
        Else
          Zl_סԺ���ʼ�¼_Verify(v_Nonew, ��Ա���_In, ��Ա����_In);
        End If;
      
        --�������û���Զ�����,���Զ�����,���򲻴��� 
        For r_Stuff In c_Stuff(v_Nonew, v_��ҳid) Loop
          Zl_�����շ���¼_��������(r_Stuff.�ⷿid, 25, v_Nonew, ��Ա����_In, ��Ա����_In, ��Ա����_In, 1, Sysdate);
        End Loop;
      End If;
    End If;
  End Loop;
  Begin
    Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
      Using 9, 0 || ',' || Id_In;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����걾��¼_�������;
/

--122764:��ҵ��,2018-04-12,�����޸�
Create Or Replace Procedure Zl_ҩƷЭ�����ճ���_Insert
(
  No_In         In ҩƷ�շ���¼.No%Type,
  �������id_In In ҩƷ�շ���¼.������id%Type,
  �ⷿid_In     In ҩƷ�շ���¼.�ⷿid%Type,
  ��������_In   In Number := 2,
  �ɱ��۾���_In In Number := 2,
  �ۼ۾���_In   In Number := 2,
  ����_In   In Number := 2
) As
  v_Maxserial ҩƷ�շ���¼.���%Type;
  n_�շ�id    ҩƷ�շ���¼.Id%Type;

  Cursor c_���ҩƷ Is
    Select (Rownum + v_Maxserial) As ���, Э��ҩƷid, �ϴβ���, �ϴι�Ӧ��id, �ϴ�����, Ч��, �ϴ���������, ��׼�ĺ�, ����, ժҪ, ������, ��������, ҩƷid, ҩƷ���,
           �Է�����id, �ⷿid, �ɱ���, Round(Round(�ɱ���, �ɱ��۾���_In) * ����, ����_In) As �ɱ����, �ۼ�,
           Round(Round(�ۼ�, �ۼ۾���_In) * ����, ����_In) As �ۼ۽��, Round(�ۼ� * ����, ����_In) - Round(�ɱ��� * ����, ����_In) As ���
    From (Select a.Э��ҩƷid, c.�ϴβ���, c.�ϴι�Ӧ��id, c.�ϴ�����, c.Ч��, c.�ϴ���������, c.��׼�ĺ�,
                  Round(Round(e.ʵ������, ��������_In) * (a.���� / a.��ĸ), ��������_In) As ����, e.ժҪ, e.������, e.��������, e.ҩƷid, e.��� As ҩƷ���,
                  e.�Է�����id, e.�ⷿid,
                  Decode(Sign(Nvl(c.ʵ�ʽ��, 0)), 1, (d.�ּ� - d.�ּ� * (c.ʵ�ʲ�� / c.ʵ�ʽ��)), (d.�ּ� - d.�ּ� * (b.ָ������� / 100))) As �ɱ���,
                  d.�ּ� As �ۼ�
           From Э��ҩƷ���� A, (Select b.* From �շ���ĿĿ¼ A, ҩƷ��� B Where a.Id = b.ҩƷid And Nvl(�Ƿ���, 0) = 0) B,
                (Select �ⷿid, ҩƷid, ʵ�ʽ��, ʵ�ʲ��, �ϴβɹ���, �ϴβ���, �ϴι�Ӧ��id, �ϴ�����, Ч��, �ϴ���������, ��׼�ĺ�
                  From ҩƷ���
                  Where ���� = 1 And �ⷿid = �ⷿid_In) C,
                (Select �շ�ϸĿid, �ּ�
                  From �շѼ�Ŀ
                  Where ((Sysdate Between ִ������ And ��ֹ����) Or (Sysdate >= ִ������ And ��ֹ���� Is Null))) D,
                (Select * From ҩƷ�շ���¼ Where NO = No_In And ���� = 3 And ���ϵ�� = 1) E
           Where a.Э��ҩƷid = b.ҩƷid And b.ҩƷid = d.�շ�ϸĿid And b.ҩƷid = c.ҩƷid(+) And a.ҩƷid = e.ҩƷid
           Union All
           Select a.Э��ҩƷid, c.�ϴβ���, c.�ϴι�Ӧ��id, c.�ϴ�����, c.Ч��, c.�ϴ���������, c.��׼�ĺ�,
                  Round(Round(e.ʵ������, ��������_In) * (a.���� / a.��ĸ), ��������_In) As ����, e.ժҪ, e.������, e.��������, e.ҩƷid, e.��� As ҩƷ���,
                  e.�Է�����id, e.�ⷿid,
                  Decode(Sign(Nvl(c.ʵ�ʽ��, 0)), 1, (c.�ּ� - c.�ּ� * (c.ʵ�ʲ�� / c.ʵ�ʽ��)), (c.�ּ� - c.�ּ� * (b.ָ������� / 100))) As �ɱ���,
                  c.�ּ� As �ۼ�
           From Э��ҩƷ���� A, (Select b.* From �շ���ĿĿ¼ A, ҩƷ��� B Where a.Id = b.ҩƷid And Nvl(�Ƿ���, 0) = 1) B,
                (Select �ⷿid, ҩƷid, ʵ�ʽ��, ʵ�ʲ��, �ϴβɹ���, �ϴβ���, �ϴι�Ӧ��id, �ϴ�����, Ч��, �ϴ���������, ��׼�ĺ�, ʵ�ʽ�� / ʵ������ As �ּ�
                  From ҩƷ���
                  Where ���� = 1 And �ⷿid = �ⷿid_In And ʵ������ > 0) C,
                (Select * From ҩƷ�շ���¼ Where NO = No_In And ���� = 3 And ���ϵ�� = 1) E
           Where a.Э��ҩƷid = b.ҩƷid And b.ҩƷid = c.ҩƷid And a.ҩƷid = e.ҩƷid)
    Order By Э��ҩƷid;
Begin
  Select Max(���) Into v_Maxserial From ҩƷ�շ���¼ Where NO = No_In And ���� = 3 And ���ϵ�� = 1;
  For v_���ҩƷ In c_���ҩƷ Loop
    Select ҩƷ�շ���¼_Id.Nextval Into n_�շ�id From Dual;
  
    Insert Into ҩƷ�շ���¼
      (ID, ��¼״̬, ����, NO, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ���۽��, ���, ժҪ, ������, ��������,
       ����id, ����, ��ҩ��λid, ����, Ч��, ��������, ��׼�ĺ�)
    Values
      (n_�շ�id, 1, 3, No_In, v_���ҩƷ.���, v_���ҩƷ.�Է�����id, v_���ҩƷ.�ⷿid, �������id_In, -1, v_���ҩƷ.Э��ҩƷid, v_���ҩƷ.�ϴβ���, v_���ҩƷ.����,
       v_���ҩƷ.����, v_���ҩƷ.�ɱ���, v_���ҩƷ.�ɱ����, v_���ҩƷ.�ۼ�, v_���ҩƷ.�ۼ۽��, v_���ҩƷ.���, v_���ҩƷ.ժҪ, v_���ҩƷ.������, v_���ҩƷ.��������,
       v_���ҩƷ.ҩƷid, v_���ҩƷ.ҩƷ���, v_���ҩƷ.�ϴι�Ӧ��id, v_���ҩƷ.�ϴ�����, v_���ҩƷ.Ч��, v_���ҩƷ.�ϴ���������, v_���ҩƷ.��׼�ĺ�);
  
    --����Ϊ1��ʾ���ʱ�¿�������
    Zl_ҩƷ���_Update(n_�շ�id, 0);
  End Loop;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷЭ�����ճ���_Insert;
/

--127766:����,2018-07-03,������ҩʱע��֤�ŵ���д��©
--124583:��ҵ��,2018-04-20,���ŷ�ҩ,��ҩ��д��ҩ����
CREATE OR REPLACE Procedure Zl_ҩƷ�շ���¼_������ҩ
(
  Billinfo_In   In Varchar2, --��ʽ:"id1,����1|id2,����2|....."
  Partid_In     In ҩƷ�շ���¼.�ⷿid%Type,
  People_In     In ҩƷ�շ���¼.�����%Type,
  Date_In       In ҩƷ�շ���¼.�������%Type,
  ��ҩ��ʽ_In   In ҩƷ�շ���¼.��ҩ��ʽ%Type := 3,
  ��ҩ��_In     In ҩƷ�շ���¼.������%Type := Null,
  ���ܷ�ҩ��_In In ҩƷ�շ���¼.���ܷ�ҩ��%Type := Null,
  Intdigit_In   In Number := 2,
  ��ҩ��_In     In ҩƷ�շ���¼.��ҩ��%Type := Null,
  �˲���_In     In ҩƷ�շ���¼.�˲���%Type := Null
) Is
  --ֻ������
  v_Infotmp     Varchar2(4000);
  v_Fields      Varchar2(4000);
  n_Billid      ҩƷ�շ���¼.Id%Type;
  n_����        ҩƷ�շ���¼.����%Type;
  Lng������id Number(18);
  Int���ϵ��   Number;
  Intִ��״̬   Number;
  Int����       ҩƷ�շ���¼.����%Type;
  Strno         ҩƷ�շ���¼.No%Type;
  Lng�ⷿid     ҩƷ�շ���¼.�ⷿid%Type;
  LngҩƷid     ҩƷ�շ���¼.ҩƷid%Type;
  Lng����id     ҩƷ�շ���¼.����id%Type;
  Dbl�����     Number;
  v_���ۼ�      ҩƷ�շ���¼.���ۼ�%Type;
  Intδ����     δ��ҩƷ��¼.δ����%Type;
  v_�˲�����    ҩƷ�շ���¼.�˲�����%Type;
  --��д����
  Dblʵ������ ҩƷ�շ���¼.ʵ������%Type;
  Dblʵ�ʽ�� ҩƷ�շ���¼.���۽��%Type;
  Dbl�ɱ���� ҩƷ�շ���¼.�ɱ����%Type;
  Dblʵ�ʲ�� ҩƷ�շ���¼.���%Type;
  --2002-07-31����
  --LNGLAST���� ��ҩǰȷ��������(�Ѽ���������)
  Strҩ��           Varchar2(200);
  Dbl��������       ҩƷ�շ���¼.��д����%Type;
  Lnglast����       ҩƷ�շ���¼.����%Type;
  Lngcur����        ҩƷ�շ���¼.����%Type;
  Str����           ҩƷ�շ���¼.����%Type;
  StrЧ��           ҩƷ�շ���¼.Ч��%Type;
  n_�ϴι�Ӧ��id    ҩƷ���.�ϴι�Ӧ��id%Type;
  n_�ϴβɹ���      ҩƷ���.�ϴβɹ���%Type;
  v_�ϴβ���        ҩƷ���.�ϴβ���%Type;
  d_�ϴ���������    ҩƷ���.�ϴ���������%Type;
  v_��׼�ĺ�        ҩƷ���.��׼�ĺ�%Type;
  n_��¼״̬        ҩƷ�շ���¼.��¼״̬%Type;
  n_ƽ���ɱ���      ҩƷ���.ƽ���ɱ���%Type;
  n_��ҩ��ʽ        ҩƷ�շ���¼.��ҩ��ʽ%Type;
  v_ժҪ            ҩƷ�շ���¼.ժҪ%Type;
  Bln�շ��뷢ҩ���� Number(1);
  v_Error           Varchar2(255);
  Err_Custom Exception;
  n_ʱ��     Number(1) := 0;
  n_ʱ�۷��� Number(1) := 0;
  n_�������� δ��ҩƷ��¼.��������%Type;
Begin
  Select Sysdate Into v_�˲����� From Dual;
  If Billinfo_In Is Null Then
    v_Infotmp := Null;
  Else
    v_Infotmp := Billinfo_In || '|';
  End If;
  While v_Infotmp Is Not Null Loop
    --�ֽⵥ��ID��
    v_Fields  := Substr(v_Infotmp, 1, Instr(v_Infotmp, '|') - 1);
    n_Billid  := Substr(v_Fields, 1, Instr(v_Fields, ',') - 1);
    n_����    := Substr(v_Fields, Instr(v_Fields, ',') + 1);
    v_Infotmp := Replace('|' || v_Infotmp, '|' || v_Fields || '|');

    --��ȡ���շ���¼�ĵ��ݡ�ҩƷID���ⷿID,���۽�ʵ��������������ID
    Begin
      Select a.����, a.No, a.ҩƷid, a.�ⷿid, a.����id, Nvl(a.���ۼ�, 0), Nvl(a.���۽��, 0), Nvl(a.ʵ������, 0) * Nvl(a.����, 1), a.������id,
             a.���ϵ��, Nvl(a.����, 0), a.����, a.Ч��, a.��ҩ��λid, a.����, a.��������, a.��׼�ĺ�, Nvl(a.��ҩ��ʽ, 0), a.ժҪ, a.��¼״̬
      Into Int����, Strno, LngҩƷid, Lng�ⷿid, Lng����id, v_���ۼ�, Dblʵ�ʽ��, Dblʵ������, Lng������id, Int���ϵ��, Lnglast����, Str����, StrЧ��,
           n_�ϴι�Ӧ��id, v_�ϴβ���, d_�ϴ���������, v_��׼�ĺ�, n_��ҩ��ʽ, v_ժҪ, n_��¼״̬
      From ҩƷ�շ���¼ A
      Where a.Id = n_Billid And a.������� Is Null
      For Update Nowait;

      Select '[' || c.���� || ']' || c.����, Nvl(�Ƿ���, 0) ʱ��
      Into Strҩ��, n_ʱ��
      From �շ���ĿĿ¼ C
      Where c.Id = LngҩƷid;
    Exception
      When Others Then
        Int���� := 0;
        v_Error := '���������û���ִ�з�ҩ�������ظ�������';
        Raise Err_Custom;
    End;

    If n_��ҩ��ʽ = -1 Or v_ժҪ = '�ܷ�' Then
      Int���� := 0;
    End If;

    If Int���� > 0 Then
      If Nvl(n_����, 0) = 0 Then
        Lngcur���� := Lnglast����;
      Else
        Lngcur���� := Nvl(n_����, 0);
      End If;

      --����Ƿ��Ѿ���д�ⷿ
      Bln�շ��뷢ҩ���� := 0;
      If Lng�ⷿid Is Null Then
        Bln�շ��뷢ҩ���� := 1;
      End If;
      Lng�ⷿid := Partid_In;

      --ȡ����ҩƷ������
      Begin
        Select �ϴ�����, Ч��, Nvl(��������, 0), �ϴι�Ӧ��id, �ϴβ���, �ϴ���������, ��׼�ĺ�, �ϴβɹ���
        Into Str����, StrЧ��, Dbl��������, n_�ϴι�Ӧ��id, v_�ϴβ���, d_�ϴ���������, v_��׼�ĺ�, n_�ϴβɹ���
        From ҩƷ���
        Where �ⷿid + 0 = Lng�ⷿid And ҩƷid = LngҩƷid And ���� = 1 And Nvl(����, 0) = Lngcur����;
      Exception
        When Others Then
          n_�ϴβɹ��� := 0;
          Dbl��������  := 0;
      End;

      --���������������˳�
      If Lngcur���� <> Nvl(Lnglast����, 0) Then
        If Dbl�������� < Dblʵ������ And Lngcur���� <> 0 Then
          v_Error := Strҩ�� || '�Ŀ����������㣬������ֹ��';
          Raise Err_Custom;
        End If;
      End If;

      If n_��¼״̬ = 1 Then
        --ԭʼ��ҩ��¼��ȡ���¼۸�
        n_ƽ���ɱ��� := Round(Zl_Fun_Getoutcost(LngҩƷid, Lngcur����, Lng�ⷿid), 5);
      Else
        --��ҩ�ٷ���¼��ȡԭʼ���ݼ۸�
        Select a.�ɱ���
        Into n_ƽ���ɱ���
        From ҩƷ�շ���¼ A, ҩƷ�շ���¼ B
        Where b.Id = n_Billid And a.���� = b.���� And a.No = b.No And a.�ⷿid = b.�ⷿid And a.ҩƷid + 0 = b.ҩƷid And
              a.��� = b.��� And Nvl(a.����, 0) = Nvl(b.����, 0) And (a.��¼״̬ = 1 Or Mod(a.��¼״̬, 3) = 0);
      End If;

      Dbl�ɱ���� := Round(n_ƽ���ɱ��� * Nvl(Dblʵ������, 0), Intdigit_In);
      Dblʵ�ʲ�� := Round(Dblʵ�ʽ�� - Dbl�ɱ����, Intdigit_In);

      --��ѯ��������
      Select ��������
      Into n_��������
      From δ��ҩƷ��¼
      Where NO = Strno And ���� = Int���� And (�ⷿid + 0 = Lng�ⷿid Or �ⷿid Is Null);
      
      --����ҩƷ�շ���¼�����۽��ɱ������
      Update ҩƷ�շ���¼
      Set �ⷿid = Lng�ⷿid, �ɱ��� = n_ƽ���ɱ���, �ɱ���� = Dbl�ɱ����, ��� = Dblʵ�ʲ��, ���� = Lngcur����, ���� = Str����, Ч�� = StrЧ��,
          ��ҩ�� = ��ҩ��_In, �˲��� = �˲���_In, �˲����� = v_�˲�����, ����� = People_In, ������� = Date_In, ��ҩ��ʽ = ��ҩ��ʽ_In, ������ = ��ҩ��_In,
          ���ܷ�ҩ�� = ���ܷ�ҩ��_In, ��ҩ��λid = n_�ϴι�Ӧ��id, ���� = v_�ϴβ���, �������� = d_�ϴ���������, ��׼�ĺ� = v_��׼�ĺ�,ע��֤�� = n_��������
      Where ID = n_Billid;
      --�����������
      If Sql%RowCount = 0 Then
        v_Error := 'Ҫ��ҩ��ҩƷ��¼"' || Strҩ�� || '"�����ڣ�������ֹ��';
        Raise Err_Custom;
      End If;

      --����סԺ���ü�¼��ִ��״̬(��ִ��)
      Select Decode(Sum(Nvl(����, 1) * ʵ������), Null, 1, 0, 1, 2)
      Into Intִ��״̬
      From ҩƷ�շ���¼
      Where ���� = Int���� And NO = Strno And ����id = Lng����id And ����� Is Null;
      Update סԺ���ü�¼
      Set ִ��״̬ = Intִ��״̬, ִ���� = Decode(People_In, Null, Zl_Username, People_In), ִ��ʱ�� = Date_In, ִ�в���id = Partid_In
      Where ID = Lng����id;

      --����δ��ҩƷ��¼(���δ����Ϊ����ɾ��)
      Select Count(*)
      Into Intδ����
      From ҩƷ�շ���¼
      Where ���� = Int���� And (�ⷿid + 0 = Lng�ⷿid Or �ⷿid Is Null) And NO = Strno And ����� Is Null And
            Nvl(LTrim(RTrim(ժҪ)), 'С��') <> '�ܷ�';

      If Intδ���� = 0 Then
        Delete δ��ҩƷ��¼ Where NO = Strno And ���� = Int���� And (�ⷿid + 0 = Lng�ⷿid Or �ⷿid Is Null);
      End If;

      --����ԭ���ο��Ŀ�������
      --���·�ҩ���ο��Ŀ��ü�ʵ������
      If Lnglast���� <> Lngcur���� Then
        Update ҩƷ���
        Set �������� = Nvl(��������, 0) + Dblʵ������
        Where �ⷿid + 0 = Lng�ⷿid And ҩƷid = LngҩƷid And ���� = 1 And Nvl(����, 0) = Lnglast����;

        Zl_ҩƷ���_���������쳣����(Lng�ⷿid, LngҩƷid, Lnglast����);

        Update ҩƷ���
        Set �������� = Nvl(��������, 0) - Dblʵ������
        Where �ⷿid + 0 = Lng�ⷿid And ҩƷid = LngҩƷid And ���� = 1 And Nvl(����, 0) = Lngcur����;

        Zl_ҩƷ���_���������쳣����(Lng�ⷿid, LngҩƷid, Lngcur����);
      End If;

      If n_ʱ�� = 1 And Lngcur���� > 0 Then
        n_ʱ�۷��� := 1;
      Else
        n_ʱ�۷��� := 0;
      End If;

      If Bln�շ��뷢ҩ���� = 1 Then
        Update ҩƷ���
        Set �������� = Nvl(��������, 0) - Dblʵ������, ʵ������ = Nvl(ʵ������, 0) - Dblʵ������, ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) - Dblʵ�ʽ��,
            ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) - Dblʵ�ʲ��, ƽ���ɱ��� = Decode(ƽ���ɱ���, Null, n_ƽ���ɱ���, ƽ���ɱ���),
            �ϴβɹ��� = Decode(�ϴβɹ���, Null, n_ƽ���ɱ���, �ϴβɹ���), ���ۼ� = Decode(n_ʱ�۷���, 1, Decode(���ۼ�, Null, v_���ۼ�, ���ۼ�), ���ۼ�)
        Where �ⷿid + 0 = Lng�ⷿid And ҩƷid = LngҩƷid And ���� = 1 And Nvl(����, 0) = Lngcur����;
      Else
        Update ҩƷ���
        Set ʵ������ = Nvl(ʵ������, 0) - Dblʵ������, ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) - Dblʵ�ʽ��, ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) - Dblʵ�ʲ��,
            ƽ���ɱ��� = Decode(ƽ���ɱ���, Null, n_ƽ���ɱ���, ƽ���ɱ���), �ϴβɹ��� = Decode(�ϴβɹ���, Null, n_ƽ���ɱ���, �ϴβɹ���),
            ���ۼ� = Decode(n_ʱ�۷���, 1, Decode(���ۼ�, Null, v_���ۼ�, ���ۼ�), ���ۼ�)
        Where �ⷿid + 0 = Lng�ⷿid And ҩƷid = LngҩƷid And ���� = 1 And Nvl(����, 0) = Lngcur����;
      End If;

      If Sql%RowCount = 0 Then
        If n_�ϴβɹ��� = 0 Then
          If Dblʵ������ = 0 Then
            Dblʵ������ := 1;
          End If;
          n_�ϴβɹ��� := Round(Dbl�ɱ���� / Dblʵ������, 5);
        End If;

        If Bln�շ��뷢ҩ���� = 1 Then
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, ��������, ʵ������, ʵ�ʽ��, ʵ�ʲ��, �ϴ�����, �ϴβ���, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ���������, ��׼�ĺ�, Ч��, ƽ���ɱ���, ���ۼ�)
          Values
            (Lng�ⷿid, LngҩƷid, Lngcur����, 1, 0 - Dblʵ������, 0 - Dblʵ������, 0 - Dblʵ�ʽ��, 0 - Dblʵ�ʲ��, Str����, v_�ϴβ���,
             n_�ϴι�Ӧ��id, n_ƽ���ɱ���, d_�ϴ���������, v_��׼�ĺ�, StrЧ��, n_ƽ���ɱ���, Decode(n_ʱ�۷���, 1, v_���ۼ�, Null));
        Else
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, ʵ������, ʵ�ʽ��, ʵ�ʲ��, �ϴ�����, �ϴβ���, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ���������, ��׼�ĺ�, Ч��, ƽ���ɱ���, ���ۼ�)
          Values
            (Lng�ⷿid, LngҩƷid, Lngcur����, 1, 0 - Dblʵ������, 0 - Dblʵ�ʽ��, 0 - Dblʵ�ʲ��, Str����, v_�ϴβ���, n_�ϴι�Ӧ��id, n_ƽ���ɱ���,
             d_�ϴ���������, v_��׼�ĺ�, StrЧ��, n_ƽ���ɱ���, Decode(n_ʱ�۷���, 1, v_���ۼ�, Null));
        End If;
      End If;

      Zl_ҩƷ���_���������쳣����(Lng�ⷿid, LngҩƷid, Lngcur����);

      Delete ҩƷ���
      Where �ⷿid + 0 = Lng�ⷿid And ҩƷid = LngҩƷid And ���� = 1 And Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And
            Nvl(ʵ�ʽ��, 0) = 0 And Nvl(ʵ�ʲ��, 0) = 0;

      --�����������
      Zl_ҩƷ�շ���¼_��������(n_Billid);
    End If;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�շ���¼_������ҩ;
/

--127766:����,2018-07-02,����ҩƷ������ҩ�󵥾����͵Ĵ���
--124583:��ҵ��,2018-04-20,���ŷ�ҩ,��ҩ��д��ҩ����
CREATE OR REPLACE Procedure Zl_ҩƷ�շ���¼_������ҩ
(
  Billid_In     In ҩƷ�շ���¼.Id%Type,
  People_In     In ҩƷ�շ���¼.�����%Type,
  Date_In       In ҩƷ�շ���¼.�������%Type,
  ����_In       In ҩƷ���.�ϴ�����%Type := Null,
  Ч��_In       In ҩƷ���.Ч��%Type := Null,
  ����_In       In ҩƷ���.�ϴβ���%Type := Null,
  ��ҩ����_In   In ҩƷ�շ���¼.ʵ������%Type := Null,
  ��ҩ�ⷿ_In   In ҩƷ�շ���¼.�ⷿid%Type := Null,
  ��ҩ��_In     In ҩƷ�շ���¼.������%Type := Null,
  Intdigit_In   In Number := 2,
  ����_In       In Number := 2,
  ���ܷ�ҩ��_In In ҩƷ�շ���¼.���ܷ�ҩ��%Type := Null
) Is
  --ֻ������
  Int��¼״̬   ҩƷ�շ���¼.��¼״̬%Type;
  Intִ��״̬   סԺ���ü�¼.ִ��״̬%Type;
  Bln������ҩ   Number;
  Lng������id Number(18);
  Strno         ҩƷ�շ���¼.No%Type;
  Int����       ҩƷ�շ���¼.����%Type;
  Lng�ⷿid     ҩƷ�շ���¼.�ⷿid%Type;
  LngҩƷid     ҩƷ�շ���¼.ҩƷid%Type;
  Dblʵ������   ҩƷ�շ���¼.ʵ������%Type;
  Dblʵ�ʽ��   ҩƷ�շ���¼.���۽��%Type;
  Dblʵ�ʳɱ�   ҩƷ�շ���¼.�ɱ����%Type;
  Dblʵ�ʲ��   ҩƷ�շ���¼.���%Type;
  Lng����id     ҩƷ�շ���¼.����id%Type;
  n_���ۼ�      ҩƷ�շ���¼.���ۼ�%Type;
  n_�Ƿ���    Number;
  n_ʱ�۷���    Number;

  --20020731 Modified by zyb
  --������ҩʱ�������������ʸı��Ĵ���
  Lng������ ҩƷ�շ���¼.����%Type;
  Lng����   ҩƷ���.ҩ������%Type;
  Lng����   ҩƷ�շ���¼.����%Type; --ԭ����

  Str����        ҩƷ�շ���¼.����%Type; --ԭ����
  DateЧ��       ҩƷ�շ���¼.Ч��%Type; --ԭЧ��
  n_�ϴι�Ӧ��id ҩƷ���.�ϴι�Ӧ��id%Type;
  n_�ϴβɹ���   ҩƷ���.�ϴβɹ���%Type;
  v_�ϴβ���     ҩƷ���.�ϴβ���%Type;
  d_�ϴ��������� ҩƷ���.�ϴ���������%Type;
  v_��׼�ĺ�     ҩƷ���.��׼�ĺ�%Type;

  n_��¼����   סԺ���ü�¼.��¼����%Type;
  v_�շ����   סԺ���ü�¼.�շ����%Type;
  n_����       ҩƷ�շ���¼.����%Type;
  n_ԭʼ����   ҩƷ�շ���¼.ʵ������%Type;
  v_������¼id ҩƷ�շ���¼.Id%Type;
  Err_Custom Exception;
  v_Error    Varchar2(255);
  v_��ҩȷ�� ҩ����ҩ����.��ҩȷ��%Type;
  v_��ҩ     ҩ����ҩ����.��ҩ%Type;
  v_�Ŷ�״̬ Number(1);
  v_ִ��ʱ�� ҩƷ�շ���¼.�������%Type;
Begin
  If ��ҩ����_In Is Not Null Then
    If ��ҩ����_In = 0 Then
      Return;
    End If;
  End If;

  --��ȡ���շ���¼�ĵ��ݡ�ҩƷID���ⷿID
  Select a.����, a.No, a.�ⷿid, a.ҩƷid, a.����id, a.������id, a.��¼״̬, Nvl(a.����, 0), a.����, a.Ч��, a.��ҩ��λid, a.����, a.��������, a.��׼�ĺ�,
         a.�ɱ���, a.����, Nvl(a.ʵ������, 0) * Nvl(a.����, 1) As ʵ������, a.���ۼ�, Nvl(b.�Ƿ���, 0) �Ƿ���
  Into Int����, Strno, Lng�ⷿid, LngҩƷid, Lng����id, Lng������id, Int��¼״̬, Lng����, Str����, DateЧ��, n_�ϴι�Ӧ��id, v_�ϴβ���, d_�ϴ���������,
       v_��׼�ĺ�, n_�ϴβɹ���, n_����, n_ԭʼ����, n_���ۼ�, n_�Ƿ���
  From ҩƷ�շ���¼ A, �շ���ĿĿ¼ B
  Where a.ҩƷid = b.Id And a.Id = Billid_In;

  Begin
    Select Nvl(��ҩȷ��, 0), Nvl(��ҩ, 0)
    Into v_��ҩȷ��, v_��ҩ
    From ҩ����ҩ����
    Where ҩ��id = Lng�ⷿid And Rownum = 1;

  Exception
    When Others Then
      v_��ҩȷ�� := 0;
      v_��ҩ     := 0;
      Null;
  End;

  If v_��ҩȷ�� = 0 And v_��ҩ = 0 Then
    v_�Ŷ�״̬ := 2;
  Elsif v_��ҩȷ�� = 1 Then
    v_�Ŷ�״̬ := 0;
  Elsif v_��ҩ = 1 Then
    v_�Ŷ�״̬ := 1;
  End If;

  --��ȡ�ñʼ�¼ʣ��δ�������������
  --������������δ���������
  Select Sum(Nvl(ʵ������, 0) * Nvl(����, 1)), Sum(Nvl(���۽��, 0)), Sum(Nvl(�ɱ����, 0)), Sum(Nvl(���, 0))
  Into Dblʵ������, Dblʵ�ʽ��, Dblʵ�ʳɱ�, Dblʵ�ʲ��
  From ҩƷ�շ���¼
  Where ����� Is Not Null And NO = Strno And ���� = Int���� And ��� = (Select ��� From ҩƷ�շ���¼ Where ID = Billid_In);

  --���������ҩ��Ϊ�㣬��ʾ����ҩ
  If Dblʵ������ = 0 Then
    v_Error := '�õ����ѱ���������Ա��ҩ����ˢ�º����ԣ�';
    Raise Err_Custom;
  End If;
  If Nvl(��ҩ����_In, 0) > Dblʵ������ Then
    v_Error := '�õ����ѱ���������Ա������ҩ����ˢ�º����ԣ�';
    Raise Err_Custom;
  End If;

  --��ȡ��ҩƷ��ǰ�Ƿ��������Ϣ
  Select Nvl(ҩ������, 0) Into Lng���� From ҩƷ��� Where ҩƷid = LngҩƷid;
  --����ǲ�����ҩ�������¼������۽����
  Bln������ҩ := 0;
  If Not (��ҩ����_In Is Null Or Nvl(��ҩ����_In, 0) = Dblʵ������) Then
    Bln������ҩ := 1;
  End If;
  If Bln������ҩ = 1 Then
    Dblʵ�ʽ�� := Round(Dblʵ�ʽ�� * ��ҩ����_In / Dblʵ������, Intdigit_In);
    Dblʵ�ʳɱ� := Round(Dblʵ�ʳɱ� * ��ҩ����_In / Dblʵ������, Intdigit_In);
    Dblʵ�ʲ�� := Round(Dblʵ�ʲ�� * ��ҩ����_In / Dblʵ������, Intdigit_In);
    Dblʵ������ := ��ҩ����_In;
  End If;

  If n_ԭʼ���� = ��ҩ����_In Then
    Dblʵ������ := ��ҩ����_In / n_����;
  Else
    n_���� := 1;
  End If;

  --lng����:0-������;1-����;2-ԭ�������ֲ�������������������;3-ԭ���������ַ���������������
  If Lng���� = 0 And Lng���� <> 0 Then
    --ԭ�������ֲ�������������������
    Lng���� := 2;
  Elsif Lng���� <> 0 And Lng���� = 0 Then
    --ԭ������,�ַ���,�����µ����Σ������²����ķ�ҩ��¼��ʹ��
    Lng���� := 3;
  Else
    If Lng���� = 0 Then
      Lng���� := 0;
    Else
      Lng���� := 1;
    End If;
  End If;
  --�ж��Ƿ�ʱ�۷���
  If (Lng���� = 1 Or Lng���� = 3) And n_�Ƿ��� = 1 Then
    n_ʱ�۷��� := 1;
  Else
    n_ʱ�۷��� := 0;
  End If;

  --��¼״̬�ĺ��������仯
  --�����ļ�¼״̬        :iif(int��¼״̬=1,0,1)+1
  --�������ļ�¼״̬        :iif(int��¼״̬=1,0,1)+2
  --�ȴ���ҩ�ļ�¼״̬    :iif(int��¼״̬=1,0,1)+3

  --����������¼
  Select ҩƷ�շ���¼_Id.Nextval Into v_������¼id From Dual;
  Insert Into ҩƷ�շ���¼
    (ID, ��¼״̬, ����, NO, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ����, ��д����, ʵ������, �ɱ���, �ɱ����, ����, ���ۼ�, ���۽��,
     ���, ժҪ, ������, ��������, ��ҩ��, �����, �������, ����id, ����, Ƶ��, �÷�, ��ҩ����, ���, ������, ��ҩ��λid, ��������, ��׼�ĺ�, ���ܷ�ҩ��, ��ҩ��ʽ, ע��֤��)
    Select v_������¼id, Int��¼״̬ + Decode(Int��¼״̬, 1, 0, 1) + 1, Int����, Strno, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ����,
           ����, Ч��, n_����, -dblʵ������, -dblʵ������, �ɱ���, -dblʵ�ʳɱ�, ����, ���ۼ�, -dblʵ�ʽ��, -dblʵ�ʲ��, ժҪ, People_In, Date_In, ��ҩ��,
           People_In, Date_In, ����id, ����, Ƶ��, �÷�, ��ҩ����, ��ҩ�ⷿ_In, ��ҩ��_In, ��ҩ��λid, ��������, ��׼�ĺ�, ���ܷ�ҩ��_In, ��ҩ��ʽ, ע��֤��
    From ҩƷ�շ���¼
    Where ID = Billid_In;

  --����ǲ��ֳ�����������Ϊ1��ʵ������Ϊ������ʵ�������Ļ�
  --����������¼�Թ�������ҩ
  Select ҩƷ�շ���¼_Id.Nextval Into Lng������ From Dual;
  Insert Into ҩƷ�շ���¼
    (ID, ��¼״̬, ����, NO, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ����, ��д����, ʵ������, �ɱ���, �ɱ����, ����, ���ۼ�, ���۽��,
     ���, ժҪ, ������, ��������, ��ҩ��, �����, �������, ����id, ����, Ƶ��, �÷�, ��ҩ����, ��ҩ��λid, ��������, ��׼�ĺ�, ע��֤��)
    Select Lng������, Int��¼״̬ + Decode(Int��¼״̬, 1, 0, 1) + 3, Int����, Strno, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid,
           Decode(Lng����, 1, ����, 3, Lng������, 0), Decode(Lng����, 3, ����_In, 1, ����, ����), Decode(Lng����, 3, ����_In, 1, ����, Null),
           Decode(Lng����, 3, Ч��_In, 1, Ч��, Null), n_����, Dblʵ������, Dblʵ������, �ɱ���, Dblʵ�ʳɱ�, ����, ���ۼ�, Dblʵ�ʽ��, Dblʵ�ʲ��, ժҪ,
           ������, ��������, Null, Null, Null, ����id, ����, Ƶ��, �÷�, ��ҩ����, ��ҩ��λid, ��������, ��׼�ĺ�, ע��֤��
    From ҩƷ�շ���¼
    Where ID = Billid_In;

  --���·��ü�¼��ִ��״̬(0-δִ��;1-��ȫִ��;2-����ִ��)
  Select Decode(Sum(Nvl(����, 1) * ʵ������), Null, 0, 0, 0, 2)
  Into Intִ��״̬
  From ҩƷ�շ���¼
  Where ���� = Int���� And NO = Strno And ����id = Lng����id And ����� Is Not Null;

  If ����_In = 1 Then
    Select ��¼����, �շ���� Into n_��¼����, v_�շ���� From ������ü�¼ Where ID = Lng����id;
  Else
    Select ��¼����, �շ���� Into n_��¼����, v_�շ���� From סԺ���ü�¼ Where ID = Lng����id;
  End If;

  If Intִ��״̬ = 0 Then
    If ����_In = 1 Then
      Update ������ü�¼
      Set ִ��״̬ = Intִ��״̬, ִ���� = Null, ִ��ʱ�� = Null
      Where NO = Strno And
            ��� = (Select ��� From ������ü�¼ Where ID = (Select ����id From ҩƷ�շ���¼ Where ID = Billid_In)) And
            Mod(��¼����, 10) = n_��¼���� And ��¼״̬ <> 2 And ִ�в���id = Lng�ⷿid;
    Else
      Update סԺ���ü�¼ Set ִ��״̬ = Intִ��״̬, ִ���� = Null, ִ��ʱ�� = Null Where ID = Lng����id;
    End If;
  Else
    If ����_In = 1 Then
      Update ������ü�¼
      Set ִ��״̬ = Intִ��״̬
      Where NO = Strno And
            ��� = (Select ��� From ������ü�¼ Where ID = (Select ����id From ҩƷ�շ���¼ Where ID = Billid_In)) And
            Mod(��¼����, 10) = n_��¼���� And ��¼״̬ <> 2 And ִ�в���id = Lng�ⷿid;
    Else
      Update סԺ���ü�¼ Set ִ��״̬ = Intִ��״̬ Where ID = Lng����id;
    End If;
  End If;

  --����δ��ҩƷ��¼
  Begin
    If ����_In = 1 Then
      Insert Into δ��ҩƷ��¼
        (����, NO, ����id, ��ҳid, ����, ���ȼ�, �Է�����id, �ⷿid, ��ҩ����, ��������, ���շ�, ��ҩ��, ��ӡ״̬, δ����, ��ҩ��, �Ŷ�״̬)
        Select a.����, a.No, a.����id, Null, a.����, Nvl(b.���ȼ�, 0) ���ȼ�, a.�Է�����id, a.�ⷿid, a.��ҩ����, a.��������, a.���շ�, Null, 1, 1,
               a.��Ʒ�ϸ�֤, v_�Ŷ�״̬
        From (Select b.����, b.No, a.����id, a.����, Decode(a.��¼״̬, 0, 0, 1) ���շ�, b.�Է�����id, b.�ⷿid, b.��ҩ����, b.��������, c.���,
                      b.��Ʒ�ϸ�֤
               From ������ü�¼ A, ҩƷ�շ���¼ B, ������Ϣ C
               Where b.Id = Billid_In And a.Id = b.����id + 0 And a.����id = c.����id(+)) A, ��� B
        Where b.����(+) = a.���;
    Else
      Insert Into δ��ҩƷ��¼
        (����, NO, ����id, ��ҳid, ����, ���ȼ�, �Է�����id, �ⷿid, ��ҩ����, ��������, ���շ�, ��ҩ��, ��ӡ״̬, δ����, ��ҩ��, �Ŷ�״̬)
        Select a.����, a.No, a.����id, a.��ҳid, a.����, Nvl(b.���ȼ�, 0) ���ȼ�, a.�Է�����id, a.�ⷿid, a.��ҩ����, a.��������, a.���շ�, Null, 1, 1,
               a.��Ʒ�ϸ�֤, v_�Ŷ�״̬
        From (Select b.����, b.No, a.����id, a.��ҳid, a.����, Decode(a.��¼״̬, 0, 0, 1) ���շ�, b.�Է�����id, b.�ⷿid, b.��ҩ����, b.��������,
                      c.���, b.��Ʒ�ϸ�֤
               From סԺ���ü�¼ A, ҩƷ�շ���¼ B, ������Ϣ C
               Where b.Id = Billid_In And a.Id = b.����id + 0 And a.����id = c.����id(+)) A, ��� B
        Where b.����(+) = a.���;
    End If;
  Exception
    When Others Then
      Null;
  End;
  
  --�޸Ĵ�������
  Zl_Prescription_Type_Update(Strno, n_��¼����, LngҩƷid, v_�շ����);
    
  --�޸�ԭ��¼Ϊ��������¼
  Update ҩƷ�շ���¼ Set ��¼״̬ = Int��¼״̬ + Decode(Int��¼״̬, 1, 0, 1) + 2 Where ID = Billid_In;

  --�޸�ҩƷ���(������)
  If Lng���� <> 3 Then
    Update ҩƷ���
    Set ʵ������ = Nvl(ʵ������, 0) + Dblʵ������ * n_����, ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) + Dblʵ�ʽ��, ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + Dblʵ�ʲ��
    Where �ⷿid + 0 = Lng�ⷿid And ҩƷid = LngҩƷid And ���� = 1 And Nvl(����, 0) = Lng����;

    If Sql%RowCount = 0 Then
      Insert Into ҩƷ���
        (�ⷿid, ҩƷid, ����, ����, ʵ������, ʵ�ʽ��, ʵ�ʲ��, ���ۼ�, �ϴ�����, Ч��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴβ���, �ϴ���������, ��׼�ĺ�, ƽ���ɱ���)
      Values
        (Lng�ⷿid, LngҩƷid, Decode(Lng����, 2, 0, Lng����), 1, Dblʵ������ * n_����, Dblʵ�ʽ��, Dblʵ�ʲ��,
         Decode(n_ʱ�۷���, 1, n_���ۼ�, Null), Decode(Lng����, 1, Str����, Null), Decode(Lng����, 1, DateЧ��, Null), n_�ϴι�Ӧ��id,
         n_�ϴβɹ���, v_�ϴβ���, d_�ϴ���������, v_��׼�ĺ�, n_�ϴβɹ���);
    End If;

    Zl_ҩƷ���_���������쳣����(Lng�ⷿid, LngҩƷid, Lng����);
  Else
    Insert Into ҩƷ���
      (�ⷿid, ҩƷid, ����, Ч��, ����, ʵ������, ʵ�ʽ��, ʵ�ʲ��, ���ۼ�, �ϴ�����, �ϴβ���, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ���������, ��׼�ĺ�, ƽ���ɱ���)
    Values
      (Lng�ⷿid, LngҩƷid, Lng������, Ч��_In, 1, Dblʵ������ * n_����, Dblʵ�ʽ��, Dblʵ�ʲ��, Decode(n_ʱ�۷���, 1, n_���ۼ�, Null), ����_In,
       ����_In, n_�ϴι�Ӧ��id, n_�ϴβɹ���, d_�ϴ���������, v_��׼�ĺ�, n_�ϴβɹ���);
  End If;

  Delete ҩƷ���
  Where �ⷿid + 0 = Lng�ⷿid And ҩƷid = LngҩƷid And ���� = 1 And Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0 And
        Nvl(ʵ�ʲ��, 0) = 0;

  --�����������
  Zl_ҩƷ�շ���¼_��������(v_������¼id);

  Begin
    --�ƶ�֧������Ŀ�ڷ�ҩ��̬��������������Ϣ�Ĺ���
    Execute Immediate 'Begin zl_������Ϣ_����(:1,:2); End;'
      Using 7, Billid_In || ',' || ��ҩ����_In || ',' || ����_In;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�շ���¼_������ҩ;
/

--115765:����,2018-06-28,�۸񾫶ȴ���
--127911,��ҵ��,2018-06-28,��ֵ����ȡ����ⷿ�ĳɱ���
--124862:��ҵ��,2018-04-26,ֻ��������δ��˵ķ��ϵ���
Create Or Replace Procedure Zl_ҩƷ�շ���¼_��������
(
  �շ�id_In     In Varchar2, --��ʽ:id1,����1|id2,����2|.....
  �ⷿid_In     In ҩƷ�շ���¼.�ⷿid%Type,
  �����_In     In ҩƷ�շ���¼.�����%Type,
  �������_In   In ҩƷ�շ���¼.�������%Type,
  ���Ϸ�ʽ_In   In ҩƷ�շ���¼.��ҩ��ʽ%Type := 3, --1-��������;2-��������;3-���ŷ���;-1 ֹͣ����
  ������_In     In ҩƷ�շ���¼.������%Type := Null,
  ���ϱ�ʶ��_In In ҩƷ�շ���¼.���ܷ�ҩ��%Type := Null,
  ������_In     In ҩƷ�շ���¼.��ҩ��%Type := Null,
  ����˱���_In In ��Ա��.���%Type := Null
) Is
  --ֻ������
  Err_Item Exception;
  v_Err_Msg Varchar2(255);

  v_����     ҩƷ�շ���¼.����%Type;
  v_�ϴβ��� ҩƷ���.�ϴβ���%Type;
  v_��׼�ĺ� ҩƷ���.��׼�ĺ�%Type;

  v_Loop_Str Varchar2(4000);
  v_Fields   Varchar2(4000);

  n_Id       ҩƷ�շ���¼.Id%Type;
  n_����     ҩƷ�շ���¼.����%Type;
  n_�ɱ���   ҩƷ�շ���¼.�ɱ���%Type;
  n_�ⷿid   ҩƷ�շ���¼.�ⷿid%Type;
  n_����� ҩƷ���.ʵ�ʽ��%Type;
  n_����� ҩƷ���.ʵ�ʲ��%Type;
  n_δ����   δ��ҩƷ��¼.δ����%Type;
  --��д����
  n_�ɱ���� ҩƷ�շ���¼.�ɱ����%Type;
  n_ʵ�ʲ�� ҩƷ�շ���¼.���%Type;
  n_�������� ҩƷ�շ���¼.��д����%Type;
  n_����_Cur ҩƷ�շ���¼.����%Type;
  n_ʵ������ �շ���ĿĿ¼.�Ƿ���%Type;

  n_�ϴι�Ӧ��id       ҩƷ���.�ϴι�Ӧ��id%Type;
  n_�ϴβɹ���         ҩƷ���.�ϴβɹ���%Type;
  n_ִ��״̬           Number;
  n_�����             Number;
  n_�շ��뷢�Ϸ���     Number(1);
  n_С��               Number(2);
  n_�ɱ���С��         Number(2);
  n_����δ��˴������� Number(2);
  n_���               Number;

  d_Ч��                   ҩƷ�շ���¼.Ч��%Type;
  d_�ϴ���������           ҩƷ���.�ϴ���������%Type;
  n_�����־               Number(1);
  v_���no                 ҩƷ�շ���¼.No%Type;
  v_���ⷿid             ҩƷ�շ���¼.�ⷿid%Type := 0;
  v_������Ϣ               Varchar2(200);
  n_����ⷿ               ҩƷ���.�ⷿid%Type;
  v_����δ��˼��˵�����   Number(1);
  v_����δ�շѵĻ��۵����� Number(1);
  v_�Զ���˼��˵�         Number(1);
  n_ƽ���ɱ���             ҩƷ���.ƽ���ɱ���%Type;
Begin
  --��ȡ���С��λ��
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_С�� From Dual;
  --��ȡ�ɱ���С��λ��
  Select Zl_To_Number(Nvl(zl_GetSysParameter(157), '2')) Into n_�ɱ���С�� From Dual;

  Select zl_GetSysParameter('����δ��˵ļ��˴�������') Into v_����δ��˼��˵����� From Dual;
  Select zl_GetSysParameter('ִ�к��Զ���˻��۵�') Into v_�Զ���˼��˵� From Dual;
  Select zl_GetSysParameter('����δ�շѵ����ﻮ�۴�������') Into v_����δ�շѵĻ��۵����� From Dual;

  If �շ�id_In Is Null Then
    v_Loop_Str := Null;
  Else
    v_Loop_Str := �շ�id_In || '|';
  End If;

  While v_Loop_Str Is Not Null Loop
    --�ֽⵥ��ID��
    v_Fields   := Substr(v_Loop_Str, 1, Instr(v_Loop_Str, '|') - 1);
    n_Id       := Substr(v_Fields, 1, Instr(v_Fields, ',') - 1);
    n_����     := Substr(v_Fields, Instr(v_Fields, ',') + 1);
    v_Loop_Str := Replace('|' || v_Loop_Str, '|' || v_Fields || '|');
  
    --�����ز���
    v_Err_Msg := 'NO';
    For c_Check In (Select a.Id, a.����, b.Id ����id, a.No, b.No ����no, b.����id, Null ��ҳid, Null ���˲���id, b.���˿���id, b.��������id,
                           b.ִ�в���id, b.������Ŀid, b.ʵ�ս��, b.����Ա���, b.����Ա����, Nvl(b.��¼״̬, 0) As ��˱�־, a.�����,
                           Decode(Nvl(a.ժҪ, 'No�ܷ�'), '�ܷ�', 3, b.ִ��״̬) ִ��״̬, 1 �����־
                    From ҩƷ�շ���¼ A, ������ü�¼ B
                    Where a.����id = b.Id And a.Id = n_Id And a.����� Is Null And a.���� In (24, 25, 26)
                    Union All
                    Select a.Id, a.����, b.Id ����id, a.No, b.No ����no, b.����id, b.��ҳid, b.���˲���id, b.���˿���id, b.��������id, b.ִ�в���id,
                           b.������Ŀid, b.ʵ�ս��, b.����Ա���, b.����Ա����, Nvl(b.��¼״̬, 0) As ��˱�־, a.�����,
                           Decode(Nvl(a.ժҪ, 'No�ܷ�'), '�ܷ�', 3, b.ִ��״̬) ִ��״̬, 2 �����־
                    From ҩƷ�շ���¼ A, סԺ���ü�¼ B
                    Where a.����id = b.Id And a.Id = n_Id And a.����� Is Null And a.���� In (24, 25, 26)) Loop
      If Not (c_Check.����� Is Null) Then
        v_Err_Msg := '�ô���[' || c_Check.No || ']�ѱ���������Ա���ϣ�����������ֹ��';
        Raise Err_Item;
      End If;
      If Nvl(c_Check.ִ��״̬, 0) = 3 Then
        v_Err_Msg := '�ô���[' || c_Check.No || ']�Ѿܷ�������������ֹ��';
        Raise Err_Item;
      End If;
    
      If Nvl(c_Check.��˱�־, 0) = 0 And c_Check.���� = 25 Then
        If v_����δ��˼��˵����� = 0 Then
          v_Err_Msg := '�ô���[' || c_Check.No || ']��δ��ˣ�����������ֹ��';
          Raise Err_Item;
        Else
          If v_�Զ���˼��˵� = 1 Then
            --��������סԺ�ĵ���
            If c_Check.����Ա���� Is Null Then
              --��������סԺ�ĵ���
              Zl_���ʼ�¼_�������(c_Check.Id, c_Check.����id, c_Check.����no, c_Check.����id, c_Check.��ҳid, c_Check.���˲���id,
                           c_Check.���˿���id, c_Check.��������id, c_Check.ִ�в���id, c_Check.������Ŀid, c_Check.ʵ�ս��, ����˱���_In,
                           �����_In, c_Check.�����־, Null);
            Else
              --��������סԺ�ĵ���
              Zl_���ʼ�¼_�������(c_Check.Id, c_Check.����id, c_Check.����no, c_Check.����id, c_Check.��ҳid, c_Check.���˲���id,
                           c_Check.���˿���id, c_Check.��������id, c_Check.ִ�в���id, c_Check.������Ŀid, c_Check.ʵ�ս��, c_Check.����Ա���,
                           c_Check.����Ա����, c_Check.�����־, Null);
            End If;
          End If;
        End If;
      End If;
    
      If Nvl(c_Check.��˱�־, 0) = 0 And c_Check.���� = 24 And v_����δ�շѵĻ��۵����� = 0 Then
        v_Err_Msg := '�ô���[' || c_Check.No || ']��δ�շѣ�����������ֹ��';
        Raise Err_Item;
      End If;
    
      v_Err_Msg := 'Have';
    
      n_�����־ := c_Check.�����־;
    
    End Loop;
  
    If v_Err_Msg = 'NO' Then
      v_Err_Msg := 'δ�ҵ�ָ������,�����Ѿ�����������Ա����,����������ֹ��';
      Raise Err_Item;
    End If;
  
    --��ȡ���շ���¼�ĵ��ݡ�ҩƷID���ⷿID,���۽�ʵ��������������ID
    For c_�շ� In (Select a.����, a.No, a.ҩƷid, a.�ⷿid, a.����id, a.���ۼ�, Nvl(a.���۽��, 0) As ʵ�ʽ��,
                        Nvl(a.ʵ������, 0) * Nvl(a.����, 1) As ʵ������, a.������id, a.���ϵ��, Nvl(a.����, 0) As ����,
                        '[' || c.���� || ']' || c.���� As ����, a.����, a.Ч��, a.��ҩ��λid, a.����, a.��������, a.��׼�ĺ�, a.��Ʒ����, a.�ڲ�����,
                        b.��� As �������
                 From ҩƷ�շ���¼ A, �շ���ĿĿ¼ C, ������ü�¼ B
                 Where a.Id = n_Id And a.ҩƷid = c.Id And a.����id = b.Id And a.����� Is Null And a.���� In (24, 25, 26)
                 Union All
                 Select a.����, a.No, a.ҩƷid, a.�ⷿid, a.����id, a.���ۼ�, Nvl(a.���۽��, 0) As ʵ�ʽ��,
                        Nvl(a.ʵ������, 0) * Nvl(a.����, 1) As ʵ������, a.������id, a.���ϵ��, Nvl(a.����, 0) As ����,
                        '[' || c.���� || ']' || c.���� As ����, a.����, a.Ч��, a.��ҩ��λid, a.����, a.��������, a.��׼�ĺ�, a.��Ʒ����, a.�ڲ�����,
                        b.��� As �������
                 From ҩƷ�շ���¼ A, �շ���ĿĿ¼ C, סԺ���ü�¼ B
                 Where a.Id = n_Id And a.ҩƷid = c.Id And a.����id = b.Id And a.����� Is Null And a.���� In (24, 25, 26)) Loop
      If Nvl(n_����, 0) = 0 Then
        n_����_Cur := c_�շ�.����;
      Else
        n_����_Cur := Nvl(n_����, 0);
      End If;
    
      --����Ƿ��Ѿ���д�ⷿ
      n_�շ��뷢�Ϸ��� := 0;
      If c_�շ�.�ⷿid Is Null Then
        n_�շ��뷢�Ϸ��� := 1;
      End If;
    
      n_�ⷿid := �ⷿid_In;
      --ȡ�����������ϵ�����
      Begin
        Select �ϴ�����, Ч��, Nvl(��������, 0), �ϴι�Ӧ��id, �ϴβ���, �ϴ���������, ��׼�ĺ�, �ϴβɹ���, Nvl(ʵ�ʽ��, 0) ʵ�ʽ��, Nvl(ʵ�ʲ��, 0) ʵ�ʲ��
        Into v_����, d_Ч��, n_��������, n_�ϴι�Ӧ��id, v_�ϴβ���, d_�ϴ���������, v_��׼�ĺ�, n_�ϴβɹ���, n_�����, n_�����
        
        From ҩƷ���
        Where �ⷿid + 0 = n_�ⷿid And ҩƷid = c_�շ�.ҩƷid And ���� = 1 And Nvl(����, 0) = n_����_Cur;
      Exception
        When Others Then
          n_�����   := 0;
          n_�����   := 0;
          n_�ϴβɹ��� := 0;
          n_��������   := 0;
      End;
    
      --��ֵ�����������ģʽ
      Begin
        Select �ⷿid
        Into n_����ⷿ
        From ҩƷ�շ���¼
        Where ���� = 21 And ������� Is Null And ҩƷid = c_�շ�.ҩƷid And Nvl(����, 0) = c_�շ�.���� And ����id = c_�շ�.����id And
              Rownum = 1;
      Exception
        When Others Then
          n_����ⷿ := 0;
      End;
    
      --���������������˳�
      If n_����_Cur <> Nvl(c_�շ�.����, 0) Then
        If n_����ⷿ = 0 And n_�������� < Nvl(c_�շ�.ʵ������, 0) And n_����_Cur <> 0 Then
          v_Err_Msg := c_�շ�.���� || '�Ŀ����������㣬������ֹ��';
          Raise Err_Item;
        End If;
      End If;
    
      If n_����ⷿ = 0 Then
        --��ͨģʽȡ���ϲ��ż۸�
        n_�ɱ��� := Round(Zl_Fun_Getoutcost(c_�շ�.ҩƷid, c_�շ�.����, n_�ⷿid), n_�ɱ���С��);
      Else
        --��ֵ�����������ģʽȡ����ⷿ�۸�
        n_�ɱ��� := Round(Zl_Fun_Getoutcost(c_�շ�.ҩƷid, c_�շ�.����, n_����ⷿ), n_�ɱ���С��);
      End If;
      n_�ɱ���� := Round(n_�ɱ��� * c_�շ�.ʵ������, n_С��);
      n_ʵ�ʲ�� := Round(c_�շ�.ʵ�ʽ�� - n_�ɱ����, n_С��);
    
      --����ҩƷ�շ���¼�����۽��ɱ������
      Update ҩƷ�շ���¼
      Set �ɱ��� = n_�ɱ���, �ɱ���� = n_�ɱ����, ��� = n_ʵ�ʲ��, �ⷿid = n_�ⷿid, ���� = n_����_Cur, ���� = v_����, Ч�� = d_Ч��, ��ҩ�� = ������_In,
          ����� = �����_In, ������� = �������_In, ��ҩ��ʽ = ���Ϸ�ʽ_In, ������ = ������_In, ���ܷ�ҩ�� = ���ϱ�ʶ��_In, ��ҩ��λid = n_�ϴι�Ӧ��id, ���� = v_�ϴβ���,
          �������� = d_�ϴ���������, ��׼�ĺ� = v_��׼�ĺ�
      Where ID = n_Id;
    
      --�����������
      If Sql%RowCount = 0 Then
        v_Err_Msg := '��������صķ��ϼ�¼��������ϢΪ:' || c_�շ�.���� || '��������ֹ��';
        Raise Err_Item;
      End If;
    
      --���·��ü�¼��ִ��״̬(��ִ��)
      Select Decode(Sum(Nvl(����, 1) * ʵ������), Null, 1, 0, 1, 2)
      Into n_ִ��״̬
      From ҩƷ�շ���¼
      Where ���� = c_�շ�.���� And NO = c_�շ�.No And ����id = c_�շ�.����id And ����� Is Null;
    
      If n_�����־ = 1 Then
        Update ������ü�¼
        Set ִ��״̬ = n_ִ��״̬, ִ�в���id = �ⷿid_In, ִ���� = �����_In, ִ��ʱ�� = �������_In
        Where NO = c_�շ�.No And (Mod(��¼����, 10) = 1 Or Mod(��¼����, 10) = 2) And ��¼״̬ <> 2 And ��� = c_�շ�.�������;
      Else
        Update סԺ���ü�¼
        Set ִ��״̬ = n_ִ��״̬, ִ�в���id = �ⷿid_In, ִ���� = �����_In, ִ��ʱ�� = �������_In
        Where ID = c_�շ�.����id;
      End If;
    
      --����δ��ҩƷ��¼(���δ����Ϊ����ɾ��)
      Select Count(*)
      Into n_δ����
      From ҩƷ�շ���¼
      Where ���� = c_�շ�.���� And NO = c_�շ�.No And ����� Is Null And (�ⷿid + 0 = n_�ⷿid Or �ⷿid Is Null) And
            Nvl(LTrim(RTrim(ժҪ)), 'No_�ܷ�') <> '�ܷ�';
    
      If n_δ���� = 0 Then
        Delete δ��ҩƷ��¼ Where ���� = c_�շ�.���� And NO = c_�շ�.No And (�ⷿid + 0 = n_�ⷿid Or �ⷿid Is Null);
      End If;
    
      Select �Ƿ��� Into n_ʵ������ From �շ���ĿĿ¼ Where ID = c_�շ�.ҩƷid;
    
      --����ԭ���ο��Ŀ�������
      --���·������ο��Ŀ��ü�ʵ������
      If c_�շ�.���� <> n_����_Cur Then
        Update ҩƷ���
        Set �������� = Nvl(��������, 0) + c_�շ�.ʵ������,
            ���ۼ� = Decode(n_ʵ������, 1, Decode(Nvl(c_�շ�.����, 0), 0, Null, Decode(Nvl(���ۼ�, 0), 0, c_�շ�.���ۼ�, ���ۼ�)), Null)
        Where ���� = 1 And �ⷿid + 0 = n_�ⷿid And ҩƷid = c_�շ�.ҩƷid And Nvl(����, 0) = c_�շ�.����;
      
        Update ҩƷ���
        Set �������� = Nvl(��������, 0) - c_�շ�.ʵ������,
            ���ۼ� = Decode(n_ʵ������, 1, Decode(Nvl(n_����_Cur, 0), 0, Null, Decode(Nvl(���ۼ�, 0), 0, c_�շ�.���ۼ�, ���ۼ�)), Null)
        Where ���� = 1 And �ⷿid + 0 = n_�ⷿid And ҩƷid = c_�շ�.ҩƷid And Nvl(����, 0) = n_����_Cur;
      End If;
    
      If n_�շ��뷢�Ϸ��� = 1 Then
        Update ҩƷ���
        Set �������� = Nvl(��������, 0) - c_�շ�.ʵ������, ʵ������ = Nvl(ʵ������, 0) - c_�շ�.ʵ������, ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) - c_�շ�.ʵ�ʽ��,
            ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) - n_ʵ�ʲ��,
            ���ۼ� = Decode(n_ʵ������, 1, Decode(Nvl(n_����_Cur, 0), 0, Null, Decode(Nvl(���ۼ�, 0), 0, c_�շ�.���ۼ�, ���ۼ�)), Null),
            �ϴβɹ��� = Decode(�ϴβɹ���, Null, n_�ɱ���, �ϴβɹ���), ƽ���ɱ��� = Decode(ƽ���ɱ���, Null, n_�ɱ���, ƽ���ɱ���),
            ��Ʒ���� = Decode(��Ʒ����, Null, c_�շ�.��Ʒ����, ��Ʒ����), �ڲ����� = Decode(�ڲ�����, Null, c_�շ�.�ڲ�����, �ڲ�����),
            Ч�� = Decode(Ч��, Null, c_�շ�.Ч��, Ч��), �ϴ����� = Decode(�ϴ�����, Null, c_�շ�.����, �ϴ�����),
            �ϴ��������� = Decode(�ϴ���������, Null, c_�շ�.��������, �ϴ���������), �ϴβ��� = Decode(�ϴβ���, Null, c_�շ�.����, �ϴβ���)
        Where �ⷿid + 0 = n_�ⷿid And ҩƷid = c_�շ�.ҩƷid And ���� = 1 And Nvl(����, 0) = n_����_Cur;
      Else
        Update ҩƷ���
        Set ʵ������ = Nvl(ʵ������, 0) - c_�շ�.ʵ������, ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) - c_�շ�.ʵ�ʽ��, ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) - n_ʵ�ʲ��,
            ���ۼ� = Decode(n_ʵ������, 1, Decode(Nvl(n_����_Cur, 0), 0, Null, Decode(Nvl(���ۼ�, 0), 0, c_�շ�.���ۼ�, ���ۼ�)), Null),
            �ϴβɹ��� = Decode(�ϴβɹ���, Null, n_�ɱ���, �ϴβɹ���), ƽ���ɱ��� = Decode(ƽ���ɱ���, Null, n_�ɱ���, ƽ���ɱ���),
            ��Ʒ���� = Decode(��Ʒ����, Null, c_�շ�.��Ʒ����, ��Ʒ����), �ڲ����� = Decode(�ڲ�����, Null, c_�շ�.�ڲ�����, �ڲ�����),
            Ч�� = Decode(Ч��, Null, c_�շ�.Ч��, Ч��), �ϴ����� = Decode(�ϴ�����, Null, c_�շ�.����, �ϴ�����),
            �ϴ��������� = Decode(�ϴ���������, Null, c_�շ�.��������, �ϴ���������), �ϴβ��� = Decode(�ϴβ���, Null, c_�շ�.����, �ϴβ���)
        Where �ⷿid + 0 = n_�ⷿid And ҩƷid = c_�շ�.ҩƷid And ���� = 1 And Nvl(����, 0) = n_����_Cur;
      End If;
    
      If Sql%RowCount = 0 Then
        If n_�ϴβɹ��� = 0 Then
          If Nvl(c_�շ�.ʵ������, 0) = 0 Then
            n_�ϴβɹ��� := Round(n_�ɱ����, 5);
          Else
            n_�ϴβɹ��� := Round(n_�ɱ���� / c_�շ�.ʵ������, 5);
          End If;
        
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, ��������, ʵ������, ʵ�ʽ��, ʵ�ʲ��, �ϴ�����, �ϴβ���, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ���������, ��׼�ĺ�, Ч��, ���ۼ�, ��Ʒ����, �ڲ�����,
             ƽ���ɱ���)
          Values
            (n_�ⷿid, c_�շ�.ҩƷid, n_����_Cur, 1, 0 - c_�շ�.ʵ������, 0 - c_�շ�.ʵ������, 0 - c_�շ�.ʵ�ʽ��, 0 - n_ʵ�ʲ��, v_����, v_�ϴβ���,
             n_�ϴι�Ӧ��id, n_�ϴβɹ���, d_�ϴ���������, v_��׼�ĺ�, d_Ч��,
             Decode(n_ʵ������, 1, Decode(Nvl(n_����_Cur, 0), 0, Null, c_�շ�.���ۼ�), Null), c_�շ�.��Ʒ����, c_�շ�.�ڲ�����, n_�ϴβɹ���);
        End If;
      
      End If;
      Delete ҩƷ���
      Where ���� = 1 And �ⷿid + 0 = n_�ⷿid And ҩƷid = c_�շ�.ҩƷid And Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And
            Nvl(ʵ�ʽ��, 0) = 0 And Nvl(ʵ�ʲ��, 0) = 0;
    
      If n_����ⷿ > 0 Then
        --��˱�������������ⷿ���������ⵥ��
        For v_���� In (Select ���, NO, �ⷿid, ҩƷid, Nvl(����, 0) As ����, ʵ������, �ɱ���, �ɱ����, ���۽��, ���, ������id
                     From ҩƷ�շ���¼
                     Where ���� = 21 And ������� Is Null And ҩƷid = c_�շ�.ҩƷid And Nvl(����, 0) = c_�շ�.���� And ����id = c_�շ�.����id) Loop
        
          Update ҩƷ�շ���¼
          Set ���ܷ�ҩ�� = n_Id
          Where ���� = 21 And ������� Is Null And ҩƷid = c_�շ�.ҩƷid And Nvl(����, 0) = c_�շ�.���� And ����id = c_�շ�.����id;
        
          Zl_������������_Verify(v_����.���, v_����.No, v_����.�ⷿid, v_����.ҩƷid, v_����.����, v_����.ʵ������, v_����.�ɱ���, v_����.�ɱ����, v_����.���۽��,
                           v_����.���, v_����.������id, �����_In, �������_In);
        End Loop;
      
        --�����������������Ĳֿ���⹺��ⵥ��
        For v_��� In (Select NO, ���, ��ҩ��λid, ҩƷid, ����, ����, ��������, Ч��, �������, ���Ч��, ʵ������, �ɱ���, �ɱ����, ����, ���ۼ�, ���۽��, ���, ժҪ,
                            ע��֤��, Nvl(����, 0) As ����, ��Ʒ����, �ڲ�����
                     From ҩƷ�շ���¼
                     Where ���� = 21 And ������� Is Not Null And ҩƷid = c_�շ�.ҩƷid And Nvl(����, 0) = c_�շ�.���� And
                           ����id = c_�շ�.����id And ���ܷ�ҩ�� = n_Id) Loop
          Begin
            Select �ⷿid Into v_���ⷿid From ����ⷿ���� Where ����id = c_�շ�.�ⷿid;
          Exception
            When Others Then
              v_���ⷿid := 0;
          End;
        
          If v_���ⷿid > 0 Then
          
            --ͬһ�ŷ��ϵ���������ⵥ��NOҪһ��
            Select Max(NO), Max(���) + 1
            Into v_���no, n_���
            From ҩƷ�շ���¼
            Where ���� = 15 And ������� Is Null And ��ҩ��λid = v_���.��ҩ��λid And
                  ����id In
                  (Select Distinct ����id
                   From ҩƷ�շ���¼
                   Where ���� = 21 And ������� Is Not Null And
                         NO = (Select Distinct NO
                               From ҩƷ�շ���¼
                               Where ���� = 21 And ������� Is Not Null And ����id = c_�շ�.����id And ���ܷ�ҩ�� = n_Id));
          
            If v_���no Is Null Or v_���no = '' Then
              --������NOΪNull, �����µ���ⵥNO
              v_���no := Nextno(68, v_���ⷿid);
              n_���   := 1;
            End If;
          
            Begin
              If n_�����־ = 1 Then
                Select b.���� || ',' || a.���� || ',' || a.��ʶ�� || ',' || '' As ������Ϣ
                Into v_������Ϣ
                From ������ü�¼ A, ���ű� B
                Where a.���˿���id = b.Id And a.Id = c_�շ�.����id;
              Else
                Select b.���� || ',' || a.���� || ',' || a.��ʶ�� || ',' || a.���� As ������Ϣ
                Into v_������Ϣ
                From סԺ���ü�¼ A, ���ű� B
                Where a.���˿���id = b.Id And a.Id = c_�շ�.����id;
              End If;
            Exception
              When Others Then
                v_������Ϣ := '';
            End;
          
            Zl_�����⹺_Insert(v_���no, n_���, v_���ⷿid, v_���.��ҩ��λid, v_���.ҩƷid, v_���.����, v_���.����, v_���.��������, v_���.Ч��,
                           v_���.�������, v_���.���Ч��, v_���.ʵ������, v_���.�ɱ���, v_���.�ɱ����, v_���.����, v_���.���ۼ�, v_���.���۽��, v_���.���,
                           Null, '���Զ����ˡ�' || v_���.ժҪ, v_���.ע��֤��, �����_In, Null, Null, Null, Null, �������_In, Null, Null,
                           v_���.����, 1, v_������Ϣ, v_���.��Ʒ����, v_���.�ڲ�����, c_�շ�.����id);
          End If;
        End Loop;
      End If;
    
      --�����������
      Zl_�����շ���¼_��������(n_Id);
    End Loop;
  
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�շ���¼_��������;
/

--125779:����,2018-05-28,��ҩ��ҩƷid������
--125779:��ҵ��,2018-05-15,��ҩ��ҩƷid������
Create Or Replace Procedure Zl_��Һ��ҩ��¼_�������
(
  ��ҩid_In   In Varchar2, --ID��:ID1,��˱�־1,ID2,��˱�־2....
  ������Ա_In In ��Һ��ҩ��¼.������Ա%Type,
  ����ʱ��_In In ��Һ��ҩ��¼.����ʱ��%Type
) Is
  v_Tansid     Varchar2(20);
  v_Tansids    Varchar2(4000);
  v_Tmp        Varchar2(4000);
  v_Usercode   Varchar2(100);
  v_��ҩid     ҩƷ�շ���¼.Id%Type;
  n_Count      Number(1);
  d_���ʱ��   ҩƷ�շ���¼.�������%Type;
  v_No         ҩƷ�շ���¼.No%Type;
  v_�ϴ�no     ҩƷ�շ���¼.No%Type;
  n_��˱�־   Number(1);
  n_����״̬   Number(2);
  v_�շ�ids    Varchar2(4000);
  v_��ҩ����id ҩƷ�շ���¼.Id%Type;
  v_ԭʼid     ҩƷ�շ���¼.Id%Type;
  v_Error      Varchar2(255);
  Err_Custom Exception;

  Cursor c_���ʼ�¼ Is
    Select Distinct a.����id, b.����ʱ��
    From ҩƷ�շ���¼ A, ��Һ��ҩ��¼ B, ��Һ��ҩ���� C
    Where a.Id = c.�շ�id And b.Id = c.��¼id And b.Id = v_Tansid And b.����״̬ = 9;

  v_���ʼ�¼ c_���ʼ�¼%RowType;

  Cursor c_��ҩ��¼ Is
    Select /*+ rule*/
    Distinct a.Id As ��ҩid, c.�շ�id, c.����, a.ҩƷid, a.����,c.��¼id as ��ҩid
    From ҩƷ�շ���¼ A, ҩƷ�շ���¼ B, ��Һ��ҩ���� C, Table(Cast(f_Num2list(v_Tansids) As Zltools.t_Numlist)) D
    Where c.��¼id = d.Column_Value And b.Id = c.�շ�id And a.���� = b.���� And a.No = b.No And a.�ⷿid + 0 = b.�ⷿid And
          a.ҩƷid + 0 = b.ҩƷid And a.��� = b.��� And (a.��¼״̬ = 1 Or Mod(a.��¼״̬, 3) = 0)
    Order By a.ҩƷid, a.����;

  v_��ҩ��¼ c_��ҩ��¼%RowType;

  Cursor c_�������� Is
    Select /*+ rule*/
     a.No, a.��� || ':' || c.���� || ':' || c.��¼id As �������
    From סԺ���ü�¼ A, ҩƷ�շ���¼ B, ��Һ��ҩ���� C, Table(Cast(f_Num2list(v_Tansids) As Zltools.t_Numlist)) D
    Where a.Id = b.����id And b.Id = c.�շ�id And Mod(b.��¼״̬, 3) = 1 And c.��¼id = d.Column_Value;

  v_�������� c_��������%RowType;

Begin
  If ��ҩid_In Is Null Then
    v_Tmp := Null;
  Else
    v_Tmp := ��ҩid_In || ',';
  End If;

  v_Usercode := Zl_Identity;
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ';') + 1);
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ',') + 1);
  v_Usercode := Substr(v_Usercode, 1, Instr(v_Usercode, ',') - 1);

  While v_Tmp Is Not Null Loop
    --�ֽⵥ��ID��
    v_Tansid   := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp      := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
    n_��˱�־ := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp      := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
  
    v_�շ�ids := Null;
  
    --ͳ�����ȷ�ϵ���Һ��(n_��˱�־ = 1)
    If n_��˱�־ = 1 Then
      If v_Tansids Is Null Then
        v_Tansids := v_Tansid;
      Else
        v_Tansids := v_Tansids || ',' || v_Tansid;
      End If;
    End If;
  
    --��鵱ǰ��Һ����״̬�Ƿ�Ϊ����ҩ״̬
    Begin
      Select ����״̬ Into n_����״̬ From ��Һ��ҩ��¼ Where ID = v_Tansid;
    
      If n_����״̬ <> 9 Then
        v_Error := '�������ѱ����������ܽ���������ˣ�';
        Raise Err_Custom;
      End If;
    Exception
      When Others Then
        v_Error := '�������ѱ�������';
        Raise Err_Custom;
    End;
  
    If n_��˱�־ = 1 Then
      n_����״̬ := 10;
    Elsif n_��˱�־ = 2 Then
      n_����״̬ := 11;
    End If;
  
    --������Һ����Ӧ���շ�NO
    Begin
      Select NO
      Into v_No
      From ҩƷ�շ���¼
      Where ID In (Select �շ�id From ��Һ��ҩ���� Where ��¼id In (Select ID From ��Һ��ҩ��¼ Where ID = v_Tansid)) And Rownum < 2;
    Exception
      When Others Then
        v_No := '';
    End;
  
    --�շ�NO��ͬ����ҩID�����ʱ���Դ�����Ϊ�ӳ�1��
    If v_No = v_�ϴ�no Then
      d_���ʱ�� := d_���ʱ�� + 1 / 24 / 60 / 60;
    Else
      d_���ʱ�� := ����ʱ��_In;
      v_�ϴ�no   := v_No;
    End If;
  
    --���ʼ�¼����
    For v_���ʼ�¼ In c_���ʼ�¼ Loop
      Zl_���˷�������_Audit(v_���ʼ�¼.����id, v_���ʼ�¼.����ʱ��, ������Ա_In, d_���ʱ��, n_��˱�־);
    End Loop;
  
    Select Count(*) Into n_Count From ��Һ��ҩ״̬ Where ��ҩid = v_Tansid And ����ʱ�� = ����ʱ��_In;
  
    If n_Count <> 1 Then
      Insert Into ��Һ��ҩ״̬
        (��ҩid, ��������, ������Ա, ����ʱ��)
      Values
        (v_Tansid, n_����״̬, ������Ա_In, ����ʱ��_In);
    End If;
    Update ��Һ��ҩ��¼ Set ������Ա = ������Ա_In, ����ʱ�� = ����ʱ��_In, ����״̬ = n_����״̬ Where ID = v_Tansid;
  End Loop;

  --����ҩ
  For v_��ҩ��¼ In c_��ҩ��¼ Loop
    Zl_ҩƷ�շ���¼_������ҩ(v_��ҩ��¼.��ҩid, ������Ա_In, ����ʱ��_In, Null, Null, Null, v_��ҩ��¼.����, Null, ������Ա_In);
  
    --ȡ��ҩ����id
    Select a.Id
    Into v_��ҩid
    From ҩƷ�շ���¼ A, ҩƷ�շ���¼ B
    Where b.Id = v_��ҩ��¼.��ҩid And a.���� = b.���� And a.No = b.No And a.�ⷿid + 0 = b.�ⷿid And a.ҩƷid + 0 = b.ҩƷid And
          a.��� = b.��� And Mod(a.��¼״̬, 3) = 1 And a.������� Is Null;
  
    --��Һ��ҩ�����е��շ�ID����Ϊ��ҩ�������շ�ID
    Update ��Һ��ҩ���� Set �շ�id = v_��ҩid Where ��¼id = v_��ҩ��¼.��ҩid And �շ�id = v_��ҩ��¼.�շ�id;
  
    If v_�շ�ids Is Null Then
      v_�շ�ids := v_��ҩid;
    Else
      v_�շ�ids := v_�շ�ids || ',' || v_��ҩid;
    End If;
  
    --ȡԭʼid
    Select a.Id
    Into v_ԭʼid
    From ҩƷ�շ���¼ A, ҩƷ�շ���¼ B
    Where b.Id = v_��ҩ��¼.��ҩid And a.���� = b.���� And a.No = b.No And a.�ⷿid + 0 = b.�ⷿid And a.ҩƷid + 0 = b.ҩƷid And
          a.��� = b.��� And Mod(a.��¼״̬, 3) = 0 And a.������� Is Not Null;
  
    Insert Into ��Һ��ҩ����
      (��¼id, �շ�id, ����)
      Select ��¼id, v_ԭʼid, ���� From ��Һ��ҩ���� Where ��¼id = v_��ҩ��¼.��ҩid And �շ�id = v_��ҩid;
  
    v_�շ�ids := v_�շ�ids || ',' || v_ԭʼid;
  End Loop;

  --��������
  For v_�������� In c_�������� Loop
    Zl_סԺ���ʼ�¼_Delete(v_��������.No, v_��������.�������, v_Usercode, Zl_Username, 2, 1, 1, d_���ʱ��);
  End Loop;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Һ��ҩ��¼_�������;
/

--127629:��ҵ��,2018-06-22,ȥ�������е�commit
Create Or Replace Procedure Zl1_Autocloseaccount Is
  v_Lngid    ҩƷ����¼.Id%Type;
  d_��ʼ���� ҩƷ����¼.�ڳ�����%Type;
  d_�������� ҩƷ����¼.��ĩ����%Type;
  n_���ʱ�� Number(2);
  v_Error    Varchar2(255);
  Err_Custom Exception;
  d_��������     ҩƷ����¼.��ĩ����%Type;
  n_���id       ҩƷ����¼.Id%Type;
  n_δ��˽��id ҩƷ����¼.Id%Type;

  Cursor c_Stock Is
    Select Distinct b.Id
    From ��������˵�� A, ���ű� B
    Where a.����id = b.Id And a.�������� In ('��ҩ��', '��ҩ��', '��ҩ��', '��ҩ��', '��ҩ��', '��ҩ��', '�Ƽ���') And
          To_Char(b.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'
    Order By b.Id;
  r_Stock c_Stock%RowType;
Begin
  --ȡ���ʱ�㣬Ĭ��ÿ�����һ�ս��
  n_���ʱ�� := Nvl(zl_GetSysParameter(221), 0);

  --ֻ���Զ������ߴ˹��̣��ֹ���治���ߴ˹���
  If n_���ʱ�� <> -1 Then
    --���㱾�ν��Ľ������ڣ���Ϊ�Զ�����Ƕ�ǰһ�����ݽ��н�棬������Ҫ����ǰ������ǰһ����������жϣ�
    If n_���ʱ�� = 0 Or n_���ʱ�� > To_Number(To_Char(Trunc(Last_Day(Sysdate - 1)), 'dd')) Then
      --ָ����ÿ�����һ���棻���߽��ʱ������˱������������Ҳ���������һ����
      d_�������� := Trunc(Last_Day(Sysdate - 1)) + 1 - 1 / 24 / 60 / 60;
    Else
      d_�������� := Trunc(Sysdate - 1, 'MONTH') + n_���ʱ�� - 1 / 24 / 60 / 60;
    End If;
  
    --������ڣ��ڽ��ʱ�����ܽ����Զ����
    If Sysdate - d_�������� > 0 Then
      For r_Stock In c_Stock Loop
        --�ж��ڼ����Ƿ��н��(����ת��)
        --�˴�����ͨ�����ڼ䡱�ֶν����жϣ�����ͨ�����ʱ������жϣ���2016-05-28 23��59��59�������򲻽�棬������
        Select Nvl(Max(ID), 0)
        Into n_���id
        From ҩƷ����¼
        Where �ⷿid = r_Stock.Id And ��ĩ���� = d_�������� And ȡ���� Is Null;
      
        If n_���id > 0 Then
          --�����ǰ�ڼ��Ѿ������ˣ��Ͳ��ٽ�棬һ���ڼ�ֻ���һ��
          Null;
        Else
          --ȡ�ⷿ���Ľ��ID�ͱ��ν��Ŀ�ʼ����
          Select Nvl(Max(ID), 0), Max(��ĩ����) + 1 / 24 / 60 / 60
          Into n_���id, d_��ʼ����
          From ҩƷ����¼
          Where �ⷿid = r_Stock.Id And ȡ���� Is Null;
        
          --��ʼʱ�䲻�ܴ��ڽ���ʱ��
          If d_��ʼ���� <= d_�������� Then
            If n_���id > 0 Then
              --����Ƿ����δ��˵Ľ�棬����������Զ����(ͨ������������ڼ����ֹ����)
              Select Nvl(Max(ID), 0)
              Into n_δ��˽��id
              From ҩƷ����¼
              Where �ⷿid = r_Stock.Id And ������� Is Null;
            
              If n_δ��˽��id > 0 Then
                Zl_ҩƷ����¼_Verify(n_δ��˽��id, Zl_Username);
              End If;
            
              --�����µĽ���¼
              Select ҩƷ����¼_Id.Nextval Into v_Lngid From Dual;
            
              Insert Into ҩƷ����¼
                (ID, �ⷿid, �ڳ�����, ��ĩ����, ������, ��������, �ϴν��id, �ڼ�, ����)
              Values
                (v_Lngid, r_Stock.Id, d_��ʼ����, d_��������, Nvl(Zl_Username, 'zlhis'), Sysdate, n_���id,
                 To_Char(Trunc(d_��������), 'yyyymm'), 1);
            
              --����ҩƷ�����ϸ��������ĩ=�����ڳ�(������ĩ)+�ڼ䷢��
              Insert Into ҩƷ�����ϸ
                (���id, �ⷿid, ҩƷid, ����, �ڳ�����, �ڳ����, �ڳ����, ��ĩ����, ��ĩ���, ��ĩ���)
                Select v_Lngid, �ⷿid, ҩƷid, ����, Sum(�ڳ�����), Sum(�ڳ����), Sum(�ڳ����), Sum(��ĩ����), Sum(��ĩ���), Sum(��ĩ���)
                From (Select a.�ⷿid, a.ҩƷid, Nvl(a.����, 0) As ����, a.��ĩ���� As �ڳ�����, a.��ĩ��� As �ڳ����, a.��ĩ��� As �ڳ����, a.��ĩ����,
                              a.��ĩ���, a.��ĩ���
                       From ҩƷ�����ϸ A, ҩƷ��� B
                       Where a.ҩƷid = b.ҩƷid And a.���id = n_���id
                       Union All
                       Select a.�ⷿid, a.ҩƷid, Nvl(a.����, 0) As ����, 0 As �ڳ�����, 0 As �ڳ����, 0 As �ڳ����,
                              a.���ϵ�� * a.ʵ������ * Nvl(a.����, 1) As ��ĩ����, a.���ϵ�� * a.���۽�� As ��ĩ���, a.���ϵ�� * a.��� As ��ĩ���
                       From ҩƷ�շ���¼ A, ҩƷ��� B
                       Where a.ҩƷid = b.ҩƷid And a.�ⷿid + 0 = r_Stock.Id And a.������� Between d_��ʼ���� And d_��������)
                Group By �ⷿid, ҩƷid, ����
                Order By �ⷿid, ҩƷid, ����;
            
              --������������ĩ-����¼(��ȥ������ĩʱ�����������)
              Insert Into ҩƷ������
                (ID, ���id, �ⷿid, ҩƷid, ����, ������, ����, ��۲�)
                Select ҩƷ������_Id.Nextval, v_Lngid, a.�ⷿid, a.ҩƷid, a.����, a.ʵ������ As ������, a.ʵ�ʽ�� As ����, a.ʵ�ʲ�� As ��۲�
                From (Select �ⷿid, ҩƷid, ����, Sum(ʵ������) As ʵ������, Sum(ʵ�ʽ��) As ʵ�ʽ��, Sum(ʵ�ʲ��) As ʵ�ʲ��
                       From (Select a.�ⷿid, a.ҩƷid, Nvl(a.����, 0) As ����, Nvl(a.��ĩ����, 0) As ʵ������, Nvl(a.��ĩ���, 0) As ʵ�ʽ��,
                                     Nvl(a.��ĩ���, 0) As ʵ�ʲ��
                              From ҩƷ�����ϸ A, ҩƷ��� B
                              Where a.ҩƷid = b.ҩƷid And a.���id = v_Lngid
                              Union All
                              Select a.�ⷿid, a.ҩƷid, Nvl(a.����, 0) As ����, -1 * Nvl(a.ʵ������, 0) As ʵ������,
                                     -1 * Nvl(a.ʵ�ʽ��, 0) As ʵ�ʽ��, -1 * Nvl(ʵ�ʲ��, 0) As ʵ�ʲ��
                              From ҩƷ��� A, ҩƷ��� B
                              Where a.ҩƷid = b.ҩƷid And a.���� = 1 And a.�ⷿid = r_Stock.Id
                              Union All
                              Select a.�ⷿid, a.ҩƷid, Nvl(a.����, 0) As ����, a.���ϵ�� * a.ʵ������ * Nvl(a.����, 1) As ʵ������,
                                     a.���ϵ�� * a.���۽�� As ʵ�ʽ��, a.���ϵ�� * a.��� As ʵ�ʲ��
                              From ҩƷ�շ���¼ A, ҩƷ��� B
                              Where a.ҩƷid = b.ҩƷid And a.�ⷿid = r_Stock.Id And a.������� > d_��������) A
                       Group By �ⷿid, ҩƷid, ����) A
                Where a.ʵ������ <> 0 Or a.ʵ�ʽ�� <> 0 Or a.ʵ�ʲ�� <> 0;
              --�Զ�����������˽����Ϣ
              Zl_ҩƷ����¼_Verify(v_Lngid, Zl_Username);
            End If;
          End If;
        End If;
      End Loop;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl1_Autocloseaccount;
/

--127629:��ҵ��,2018-06-22,ȥ�������е�commit
Create Or Replace Procedure Zl1_Autostuffcloseaccount Is
  v_Lngid    ���Ͻ���¼.Id%Type;
  d_��ʼ���� ���Ͻ���¼.�ڳ�����%Type;
  d_�������� ���Ͻ���¼.��ĩ����%Type;
  n_���ʱ�� Number(2);
  v_Error    Varchar2(255);
  Err_Custom Exception;
  d_��������     ���Ͻ���¼.��ĩ����%Type;
  n_���id       ���Ͻ���¼.Id%Type;
  n_δ��˽��id ���Ͻ���¼.Id%Type;

  Cursor c_Stock Is
    Select Distinct b.Id
    From ��������˵�� A, ���ű� B
    Where a.����id = b.Id And a.�������� In ('���Ŀ�', '���ϲ���') And To_Char(b.����ʱ��, 'yyyy-MM-dd') = '3000-01-01'
    Order By b.Id;
  r_Stock c_Stock%RowType;
Begin
  --ȡ���ʱ�㣬Ĭ��ÿ�����һ�ս��
  n_���ʱ�� := Nvl(zl_GetSysParameter(281), 0);

  --ֻ���Զ������ߴ˹��̣��ֹ���治���ߴ˹���
  If n_���ʱ�� <> -1 Then
    --���㱾�ν��Ľ������ڣ���Ϊ�Զ�����Ƕ�ǰһ�����ݽ��н�棬������Ҫ����ǰ������ǰһ����������жϣ�
    If n_���ʱ�� = 0 Or n_���ʱ�� > To_Number(To_Char(Trunc(Last_Day(Sysdate - 1)), 'dd')) Then
      --ָ����ÿ�����һ���棻���߽��ʱ������˱������������Ҳ���������һ����
      d_�������� := Trunc(Last_Day(Sysdate - 1)) + 1 - 1 / 24 / 60 / 60;
    Else
      d_�������� := Trunc(Sysdate - 1, 'MONTH') + n_���ʱ�� - 1 / 24 / 60 / 60;
    End If;
  
    --������ڣ��ڽ��ʱ�����ܽ����Զ����
    If Sysdate - d_�������� > 0 Then
      For r_Stock In c_Stock Loop
        --�ж��ڼ����Ƿ��н��(����ת��)
        --�˴�����ͨ�����ڼ䡱�ֶν����жϣ�����ͨ�����ʱ������жϣ���2016-05-28 23��59��59�������򲻽�棬������
        Select Nvl(Max(ID), 0)
        Into n_���id
        From ���Ͻ���¼
        Where �ⷿid = r_Stock.Id And ��ĩ���� = d_�������� And ȡ���� Is Null;
      
        If n_���id > 0 Then
          --�����ǰ�ڼ��Ѿ������ˣ��Ͳ��ٽ�棬һ���ڼ�ֻ���һ��
          Null;
        Else
          --ȡ�ⷿ���Ľ��ID�ͱ��ν��Ŀ�ʼ����
          Select Nvl(Max(ID), 0), Max(��ĩ����) + 1 / 24 / 60 / 60
          Into n_���id, d_��ʼ����
          From ���Ͻ���¼
          Where �ⷿid = r_Stock.Id And ȡ���� Is Null;
        
          --��ʼʱ�䲻�ܴ��ڽ���ʱ��
          If d_��ʼ���� <= d_�������� Then
            If n_���id > 0 Then
              --����Ƿ����δ��˵Ľ�棬����������Զ����(ͨ������������ڼ����ֹ����)
              Select Nvl(Max(ID), 0)
              Into n_δ��˽��id
              From ���Ͻ���¼
              Where �ⷿid = r_Stock.Id And ������� Is Null;
            
              If n_δ��˽��id > 0 Then
                Zl_���Ͻ���¼_Verify(n_δ��˽��id, Zl_Username);
              End If;
            
              --�����µĽ���¼
              Select ���Ͻ���¼_Id.Nextval Into v_Lngid From Dual;
            
              Insert Into ���Ͻ���¼
                (ID, �ⷿid, �ڳ�����, ��ĩ����, ������, ��������, �ϴν��id, �ڼ�, ����)
              Values
                (v_Lngid, r_Stock.Id, d_��ʼ����, d_��������, Nvl(Zl_Username, 'zlhis'), Sysdate, n_���id,
                 To_Char(Trunc(d_��������), 'yyyymm'), 1);
            
              --����ҩƷ�����ϸ��������ĩ=�����ڳ�(������ĩ)+�ڼ䷢��
              Insert Into ���Ͻ����ϸ
                (���id, �ⷿid, ����id, ����, �ڳ�����, �ڳ����, �ڳ����, ��ĩ����, ��ĩ���, ��ĩ���)
                Select v_Lngid, �ⷿid, ����id, ����, Sum(�ڳ�����), Sum(�ڳ����), Sum(�ڳ����), Sum(��ĩ����), Sum(��ĩ���), Sum(��ĩ���)
                From (Select a.�ⷿid, a.����id, Nvl(a.����, 0) As ����, a.��ĩ���� As �ڳ�����, a.��ĩ��� As �ڳ����, a.��ĩ��� As �ڳ����, a.��ĩ����,
                              a.��ĩ���, a.��ĩ���
                       From ���Ͻ����ϸ A, �������� B
                       Where a.����id = b.����id And a.���id = n_���id
                       Union All
                       Select a.�ⷿid, a.ҩƷid, Nvl(a.����, 0) As ����, 0 As �ڳ�����, 0 As �ڳ����, 0 As �ڳ����,
                              a.���ϵ�� * a.ʵ������ * Nvl(a.����, 1) As ��ĩ����, a.���ϵ�� * a.���۽�� As ��ĩ���, a.���ϵ�� * a.��� As ��ĩ���
                       From ҩƷ�շ���¼ A, �������� B
                       Where a.ҩƷid = b.����id And a.�ⷿid + 0 = r_Stock.Id And a.������� Between d_��ʼ���� And d_��������)
                Group By �ⷿid, ����id, ����
                Order By �ⷿid, ����id, ����;
            
              --������������ĩ-����¼(��ȥ������ĩʱ�����������)
              Insert Into ���Ͻ�����
                (ID, ���id, �ⷿid, ����id, ����, ������, ����, ��۲�)
                Select ���Ͻ�����_Id.Nextval, v_Lngid, a.�ⷿid, a.����id, a.����, a.ʵ������ As ������, a.ʵ�ʽ�� As ����, a.ʵ�ʲ�� As ��۲�
                From (Select �ⷿid, ����id, ����, Sum(ʵ������) As ʵ������, Sum(ʵ�ʽ��) As ʵ�ʽ��, Sum(ʵ�ʲ��) As ʵ�ʲ��
                       From (Select a.�ⷿid, a.����id, Nvl(a.����, 0) As ����, Nvl(a.��ĩ����, 0) As ʵ������, Nvl(a.��ĩ���, 0) As ʵ�ʽ��,
                                     Nvl(a.��ĩ���, 0) As ʵ�ʲ��
                              From ���Ͻ����ϸ A, �������� B
                              Where a.����id = b.����id And a.���id = v_Lngid
                              Union All
                              Select a.�ⷿid, a.ҩƷid, Nvl(a.����, 0) As ����, -1 * Nvl(a.ʵ������, 0) As ʵ������,
                                     -1 * Nvl(a.ʵ�ʽ��, 0) As ʵ�ʽ��, -1 * Nvl(ʵ�ʲ��, 0) As ʵ�ʲ��
                              From ҩƷ��� A, �������� B
                              Where a.ҩƷid = b.����id And a.���� = 1 And a.�ⷿid = r_Stock.Id
                              Union All
                              Select a.�ⷿid, a.ҩƷid, Nvl(a.����, 0) As ����, a.���ϵ�� * a.ʵ������ * Nvl(a.����, 1) As ʵ������,
                                     a.���ϵ�� * a.���۽�� As ʵ�ʽ��, a.���ϵ�� * a.��� As ʵ�ʲ��
                              From ҩƷ�շ���¼ A, �������� B
                              Where a.ҩƷid = b.����id And a.�ⷿid = r_Stock.Id And a.������� > d_��������) A
                       Group By �ⷿid, ����id, ����) A
                Where a.ʵ������ <> 0 Or a.ʵ�ʽ�� <> 0 Or a.ʵ�ʲ�� <> 0;
              --�Զ�����������˽����Ϣ
              Zl_���Ͻ���¼_Verify(v_Lngid, Zl_Username);
            End If;
          End If;
        End If;
      End Loop;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl1_Autostuffcloseaccount;
/

--127629:��ҵ��,2018-06-22,ȥ�������е�commit
CREATE OR REPLACE Procedure Zl1_Autosend As
  v_�ⷿid       ҩƷ�շ���¼.�ⷿid%Type;
  v_�Զ���ҩ���� ҩ����ҩ����.�Զ���ҩ����%Type;

  Intdigit      Number(1);
  Intautoverify Number(1);
  Str����Ա��� ��Ա��.���%Type;
  Str����Ա���� ��Ա��.����%Type;

  Cursor Autosenddepid Is
    Select Nvl(ҩ��id, 0) ҩ��id, �Զ���ҩ���� From ҩ����ҩ���� Where ���� = 2 And Nvl(�Զ���ҩ����, 0) > 0;

  Cursor Autosendlist Is
    Select Distinct A.�ⷿid, A.ID, Nvl(A.����, 0) ����, C.����Ա����,a.ҩƷid
    From ҩƷ�շ���¼ A, δ��ҩƷ��¼ B, סԺ���ü�¼ C
    Where A.���� = B.���� And A.NO = B.NO And A.����id = C.ID And Nvl(A.�ⷿid, v_�ⷿid) + 0 = Nvl(B.�ⷿid, v_�ⷿid) And
          A.���� In (9, 10) And Mod(A.��¼״̬, 3) = 1 And A.����� Is Null And Nvl(A.�ⷿid, 0) + 0 = v_�ⷿid And
          B.�������� < Sysdate - v_�Զ���ҩ���� order by a.ҩƷid;

  v_Autosenddepid Autosenddepid%RowType;
  v_Autosendlist  Autosendlist%RowType;
Begin
  If f_Is_Primary_Node = 0 Then
    Return;
  End If;

  --ȡ����Ա���������
  Select ���, ����
  Into Str����Ա���, Str����Ա����
  From ��Ա�� A, �ϻ���Ա�� B
  Where A.ID = B.��Աid And B.�û��� = User;

  --��ȡ���С��λ��
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into Intdigit From Dual;
  --�жϻ��۵���ҩ���Ƿ��Զ����Ϊ���ʵ�
  Select Zl_To_Number(zl_GetSysParameter(81)) Into Intautoverify From Dual;

  For v_Autosenddepid In Autosenddepid Loop
    v_�ⷿid       := v_Autosenddepid.ҩ��id;
    v_�Զ���ҩ���� := v_Autosenddepid.�Զ���ҩ����;
    If v_�Զ���ҩ���� > 30 Then
      v_�Զ���ҩ���� := 30;
    End If;
    For v_Autosendlist In Autosendlist Loop
      Zl_ҩƷ�շ���¼_���ŷ�ҩ(v_Autosendlist.�ⷿid, v_Autosendlist.ID, v_Autosendlist.����Ա����, Sysdate, v_Autosendlist.����, 3, Null,
                     Null, Str����Ա���, Str����Ա����, Intdigit, Intautoverify);
    End Loop;
  End Loop;
End Zl1_Autosend;
/

--127629:��ҵ��,2018-06-22,ȥ�������е�commit
Create Or Replace Procedure Zl_���ϲ������_Update
(
  �ڼ�_In In �ڼ��.�ڼ�%Type
) Is
  Cursor c_�ڼ�� Is
    Select ��ʼ����, ��ֹ���� From �ڼ�� Where �ڼ� >= �ڼ�_In And Sysdate >= ��ʼ����;

  Cursor c_ƽ�������
  (
    v_��ʼ���� Date,
    v_��ֹ���� Date
  ) Is
    Select d.ҩ��, s.�ⷿid, s.����id, s.����, Decode(Sign(s.���), 1, s.��� / s.���, m.ָ������� / 100) As �����
    From (Select o.�ⷿid, o.����id, o.����, Nvl(e.��ǰ���, 0) - Nvl(j.�������, 0) - o.������ As ���,
                  Nvl(e.��ǰ���, 0) - Nvl(j.�������, 0) - o.������ As ���
           From (Select �ⷿid, ҩƷid ����id, Nvl(����, 0) As ����, Sum(���ϵ�� * ���۽��) As ������, Sum(���ϵ�� * ���) As ������
                  From ҩƷ�շ���¼ L
                  Where ������� Between Trunc(v_��ʼ����) And Trunc(v_��ֹ����) + 1 - 1 / 24 / 60 / 60 And
                        (���� = 19 And Exists
                         (Select 1 From ��������˵�� C Where c.����id = l.�ⷿid And c.�������� In ('���Ŀ�', '����ⷿ')) And Not Exists
                         (Select 1
                          From ��������˵�� C
                          Where c.����id = l.�Է�����id And c.�������� In ('���Ŀ�', '�Ƽ���', '����ⷿ')) Or ���� Between 8 And 11 Or
                         ���� In (20, 21))
                  Group By �ⷿid, ҩƷid, Nvl(����, 0)) O,
                (Select �ⷿid, ҩƷid ����id, Nvl(����, 0) As ����, Sum(���ϵ�� * ���۽��) As �������, Sum(���ϵ�� * ���) As �������
                  From ҩƷ�շ���¼
                  Where ������� >= Trunc(v_��ֹ����) + 1
                  Group By �ⷿid, ҩƷid, Nvl(����, 0)) J,
                (Select �ⷿid, ҩƷid ����id, Nvl(����, 0) As ����, Sum(ʵ�ʽ��) As ��ǰ���, Sum(ʵ�ʲ��) As ��ǰ���
                  From ҩƷ���
                  Where ���� = 1
                  Group By �ⷿid, ҩƷid, Nvl(����, 0)) E
           Where o.�ⷿid = j.�ⷿid(+) And o.����id = j.����id(+) And o.���� = j.����(+) And o.�ⷿid = e.�ⷿid(+) And
                 o.����id = e.����id(+) And o.���� = e.����(+)) S, �������� M,
         (Select ����id, Min(Decode(��������, '���Ŀ�', 1, Decode(��������, '����ⷿ', 1, 2))) As ҩ��
           From ��������˵��
           Where �������� In ('���Ŀ�', '����ⷿ', '���ϲ���')
           Group By ����id) D
    Where s.����id = m.����id And s.�ⷿid = d.����id
    Order By d.ҩ��, s.�ⷿid, s.����id, s.����;

  Cursor c_���ϳ����¼
  (
    v_��ʼ���� Date,
    v_��ֹ���� Date,
    v_�ⷿ     Integer,
    v_�ⷿid   Integer,
    v_����id   Integer,
    v_����     Integer
  ) Is
    Select ID, ����, NO, �������, ������id, ���ϵ��, �ɱ���, ʵ������ * ���� As ʵ������, ���۽��, ���, ����, ����, Ч��, ���Ч��, �Է�����id, ��������, ��׼�ĺ�,
           ��ҩ��λid
    From ҩƷ�շ���¼ L
    Where ������� Between Trunc(v_��ʼ����) And Trunc(v_��ֹ����) + 1 - 1 / 24 / 60 / 60 And �ⷿid = v_�ⷿid And ҩƷid = v_����id And
          Nvl(����, 0) = Nvl(v_����, 0) And
          (v_�ⷿ = 1 And ���� = 19 And Not Exists
           (Select 1 From ��������˵�� C Where c.����id = l.�Է�����id And c.�������� In ('���Ŀ�', '����ⷿ')) Or ���� Between 8 And 11 Or
           ���� In (20, 21));
  v_ԭ���     Number(18, 2);
  v_�ֲ��     Number(18, 2);
  v_�ɱ���     Number(18, 4);
  v_�Է����id Integer;
  v_С��       Number(2);
Begin
  Select nvl(����,2) into v_С�� From ҩƷ���ľ��� Where ����=0 and ��� = 2 And ���� = 4 And ��λ = 5;

  For v_Period In c_�ڼ�� Loop
    For v_Avgtax In c_ƽ�������(v_Period.��ʼ����, v_Period.��ֹ����) Loop
      For v_Outrec In c_���ϳ����¼(v_Period.��ʼ����, v_Period.��ֹ����, v_Avgtax.ҩ��, v_Avgtax.�ⷿid, v_Avgtax.����id, v_Avgtax.����) Loop
        v_ԭ��� := v_Outrec.���;
        v_�ֲ�� := Round(Nvl(v_Outrec.���۽��, 0) * v_Avgtax.�����, v_С��);
        If Nvl(v_Outrec.ʵ������, 0) = 0 Then
          v_�ɱ��� := v_Outrec.�ɱ���;
        Else
          v_�ɱ��� := Round((Nvl(v_Outrec.���۽��, 0) - v_�ֲ��) / v_Outrec.ʵ������, 4);
        End If;

        Update ҩƷ�շ���¼
        Set ��� = Round(v_�ֲ��, v_С��), �ɱ���� = Round(Nvl(v_Outrec.���۽��, 0) - v_�ֲ��, v_С��), �ɱ��� = v_�ɱ���
        Where ID = v_Outrec.Id;

        Update ҩƷ���
        Set ʵ�ʲ�� = Round(Nvl(ʵ�ʲ��, 0) + (v_�ֲ�� - v_ԭ���) * v_Outrec.���ϵ��, v_С��)
        Where �ⷿid = v_Avgtax.�ⷿid And ҩƷid = v_Avgtax.����id And Nvl(����, 0) = Nvl(v_Avgtax.����, 0) And ���� = 1;
        If Sql%NotFound Then
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, ��������, ʵ������, ʵ�ʽ��, ʵ�ʲ��, Ч��, ���Ч��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴ���������, �ϴβ���, ��׼�ĺ�)
          Values
            (v_Avgtax.�ⷿid, v_Avgtax.����id, v_Avgtax.����, 1, 0, 0, 0, Round((v_�ֲ�� - v_ԭ���) * v_Outrec.���ϵ��, v_С��),
             v_Outrec.Ч��, v_Outrec.���Ч��, v_Outrec.��ҩ��λid, v_�ɱ���, v_Outrec.����, v_Outrec.��������, v_Outrec.����, v_Outrec.��׼�ĺ�);
        End If;

        If v_Outrec.���� = 19 Then
          Update ҩƷ�շ���¼
          Set ��� = Round(v_�ֲ��, v_С��), �ɱ���� = Round(Nvl(v_Outrec.���۽��, 0), v_С��) - v_�ֲ��, �ɱ��� = v_�ɱ���
          Where NO = v_Outrec.No And ���� = 19 And ҩƷid + 0 = v_Avgtax.����id And Nvl(����, 0) = Nvl(v_Avgtax.����, 0) And
                �ⷿid + 0 = v_Outrec.�Է�����id And �Է�����id + 0 = v_Avgtax.�ⷿid And ���ϵ�� = -1 * v_Outrec.���ϵ��;
          If Sql%NotFound Then
            Null;
          Else
            Select ������id
            Into v_�Է����id
            From ҩƷ�շ���¼
            Where NO = v_Outrec.No And ���� = 19 And ҩƷid + 0 = v_Avgtax.����id And Nvl(����, 0) = Nvl(v_Avgtax.����, 0) And
                  �ⷿid + 0 = v_Outrec.�Է�����id And �Է�����id + 0 = v_Avgtax.�ⷿid And ���ϵ�� = -1 * v_Outrec.���ϵ�� And
                  Rownum < 2;

            Update ҩƷ���
            Set ʵ�ʲ�� = Round(Nvl(ʵ�ʲ��, 0) + (v_�ֲ�� - v_ԭ���) * v_Outrec.���ϵ�� * -1, v_С��)
            Where �ⷿid = v_Outrec.�Է�����id And ҩƷid = v_Avgtax.����id And Nvl(����, 0) = Nvl(v_Avgtax.����, 0) And ���� = 1;
            If Sql%NotFound Then
              Insert Into ҩƷ���
                (�ⷿid, ҩƷid, ����, ����, ��������, ʵ������, ʵ�ʽ��, ʵ�ʲ��, Ч��, ���Ч��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴ���������, �ϴβ���, ��׼�ĺ�)
              Values
                (v_Outrec.�Է�����id, v_Avgtax.����id, v_Avgtax.����, 1, 0, 0, 0,
                 Round((v_�ֲ�� - v_ԭ���) * v_Outrec.���ϵ�� * -1, v_С��), v_Outrec.Ч��, v_Outrec.���Ч��, v_Outrec.��ҩ��λid, v_�ɱ���,
                 v_Outrec.����, v_Outrec.��������, v_Outrec.����, v_Outrec.��׼�ĺ�);
            End If;
          End If;
        End If;
      End Loop;
      Delete From ҩƷ���
      Where �ⷿid = v_Avgtax.�ⷿid And ҩƷid = v_Avgtax.����id And Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And
            Nvl(ʵ�ʽ��, 0) = 0 And Nvl(ʵ�ʲ��, 0) = 0;
    End Loop;
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���ϲ������_Update;
/

--127629:��ҵ��,2018-06-22,ȥ�������е�commit
Create Or Replace Procedure Zl_ҩƷ�������_ȫ��ƽ��
(
  ��ʼʱ��_In In Date,
  �ⷿid_In   In ҩƷ�շ���¼.�ⷿid%Type,
  ����ʱ��_In In Date
) Is
  Cursor c_ƽ�������
  (
    v_��ʼ���� Date,
    v_��ֹ���� Date
  ) Is
    Select d.ҩ��, s.�ⷿid, s.ҩƷid, s.����, Decode(Sign(s.���), 1, s.��� / s.���, m.ָ������� / 100) As �����
    From (Select o.�ⷿid, o.ҩƷid, o.����, Nvl(e.��ǰ���, 0) - Nvl(j.�������, 0) - o.������ As ���,
                  Nvl(e.��ǰ���, 0) - Nvl(j.�������, 0) - o.������ As ���
           From (Select �ⷿid, ҩƷid, Nvl(����, 0) As ����, Sum(���ϵ�� * ���۽��) As ������, Sum(���ϵ�� * ���) As ������
                  From ҩƷ�շ���¼ L
                  Where �ⷿid = �ⷿid_In And ������� Between Trunc(v_��ʼ����) And Trunc(v_��ֹ����) + 1 - 1 / 24 / 60 / 60 And
                        (���� = 6 And Exists
                         (Select 1
                          From ��������˵�� C
                          Where c.����id = l.�ⷿid And c.�������� In ('��ҩ��', '��ҩ��', '��ҩ��')) And Not Exists
                         (Select 1
                          From ��������˵�� C
                          Where c.����id = l.�Է�����id And c.�������� In ('��ҩ��', '��ҩ��', '��ҩ��', '�Ƽ���')) Or ���� Between 7 And 11)
                  Group By �ⷿid, ҩƷid, Nvl(����, 0)) O,
                (Select �ⷿid, ҩƷid, Nvl(����, 0) As ����, Sum(���ϵ�� * ���۽��) As �������, Sum(���ϵ�� * ���) As �������
                  From ҩƷ�շ���¼
                  Where �ⷿid = �ⷿid_In And ������� >= Trunc(v_��ֹ����) + 1
                  Group By �ⷿid, ҩƷid, Nvl(����, 0)) J,
                (Select �ⷿid, ҩƷid, Nvl(����, 0) As ����, Sum(ʵ�ʽ��) As ��ǰ���, Sum(ʵ�ʲ��) As ��ǰ���
                  From ҩƷ���
                  Where ���� = 1
                  Group By �ⷿid, ҩƷid, Nvl(����, 0)) E
           Where o.�ⷿid = j.�ⷿid(+) And o.ҩƷid = j.ҩƷid(+) And o.���� = j.����(+) And o.�ⷿid = e.�ⷿid(+) And
                 o.ҩƷid = e.ҩƷid(+) And o.���� = e.����(+)) S, ҩƷ��� M,
         (Select ����id, Min(Decode(��������, '��ҩ��', 1, '��ҩ��', 1, '��ҩ��', 1, 2)) As ҩ��
           From ��������˵��
           Where �������� In ('��ҩ��', '��ҩ��', '��ҩ��', '��ҩ��', '��ҩ��', '��ҩ��')
           Group By ����id) D
    Where s.ҩƷid = m.ҩƷid And s.�ⷿid = d.����id
    Order By d.ҩ��, s.�ⷿid, s.ҩƷid, s.����;

  Cursor c_ҩƷ�����¼
  (
    v_��ʼ���� Date,
    v_��ֹ���� Date,
    v_ҩ��     Integer,
    v_�ⷿid   Integer,
    v_ҩƷid   Integer,
    v_����     Integer
  ) Is
    Select ID, ����, NO, �������, ������id, ���ϵ��, �ɱ���, ʵ������ * ���� As ʵ������, ���۽��, ���, ����, ����, Ч��, �Է�����id
    From ҩƷ�շ���¼ L
    Where �ⷿid = �ⷿid_In And ������� Between Trunc(v_��ʼ����) And Trunc(v_��ֹ����) + 1 - 1 / 24 / 60 / 60 And �ⷿid = v_�ⷿid And
          ҩƷid = v_ҩƷid And Nvl(����, 0) = Nvl(v_����, 0) And
          (v_ҩ�� = 1 And ���� = 6 And Not Exists
           (Select 1
            From ��������˵�� C
            Where c.����id = l.�Է�����id And c.�������� In ('��ҩ��', '��ҩ��', '��ҩ��', '�Ƽ���')) Or ���� Between 7 And 11);

  v_ԭ���     ҩƷ���.ʵ�ʲ��%Type;
  v_�ֲ��     ҩƷ���.ʵ�ʲ��%Type;
  v_�ɱ���     ҩƷ���.�ϴβɹ���%Type;
  v_�Է����id Integer;
  Intdigit     Number;
Begin
  --��ȡ���С��λ��
  Select Nvl(����, 2) Into Intdigit From ҩƷ���ľ��� Where ���� = 0 And ��� = 1 And ���� = 4 And ��λ = 5;

  For v_Avgtax In c_ƽ�������(��ʼʱ��_In, ����ʱ��_In) Loop
    For v_Outrec In c_ҩƷ�����¼(��ʼʱ��_In, ����ʱ��_In, v_Avgtax.ҩ��, v_Avgtax.�ⷿid, v_Avgtax.ҩƷid, v_Avgtax.����) Loop
      v_ԭ��� := v_Outrec.���;
      v_�ֲ�� := Round(Nvl(v_Outrec.���۽��, 0) * v_Avgtax.�����, Intdigit);
      If Nvl(v_Outrec.ʵ������, 0) = 0 Then
        v_�ɱ��� := v_Outrec.�ɱ���;
      Else
        v_�ɱ��� := Round((Nvl(v_Outrec.���۽��, 0) - v_�ֲ��) / v_Outrec.ʵ������, 7);
      End If;
    
      Update ҩƷ�շ���¼
      Set ��� = v_�ֲ��, �ɱ���� = Nvl(v_Outrec.���۽��, 0) - v_�ֲ��, �ɱ��� = v_�ɱ���
      Where ID = v_Outrec.Id;
    
      Update ҩƷ���
      Set ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + (v_�ֲ�� - v_ԭ���) * v_Outrec.���ϵ��
      Where �ⷿid = v_Avgtax.�ⷿid And ҩƷid = v_Avgtax.ҩƷid And Nvl(����, 0) = Nvl(v_Avgtax.����, 0) And ���� = 1;
      If Sql%NotFound Then
        Insert Into ҩƷ���
          (�ⷿid, ҩƷid, ����, ����, ��������, ʵ������, ʵ�ʽ��, ʵ�ʲ��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴβ���, Ч��)
        Values
          (v_Avgtax.�ⷿid, v_Avgtax.ҩƷid, v_Avgtax.����, 1, 0, 0, 0, (v_�ֲ�� - v_ԭ���) * v_Outrec.���ϵ��, Null, v_�ɱ���,
           v_Outrec.����, v_Outrec.����, v_Outrec.Ч��);
      End If;
    
      If v_Outrec.���� = 6 Then
        Update ҩƷ�շ���¼
        Set ��� = v_�ֲ��, �ɱ���� = Nvl(v_Outrec.���۽��, 0) - v_�ֲ��, �ɱ��� = v_�ɱ���
        Where NO = v_Outrec.No And ���� = 6 And ҩƷid + 0 = v_Avgtax.ҩƷid And Nvl(����, 0) = Nvl(v_Avgtax.����, 0) And
              �ⷿid + 0 = v_Outrec.�Է�����id And �Է�����id + 0 = v_Avgtax.�ⷿid And ���ϵ�� = -1 * v_Outrec.���ϵ��;
        If Sql%NotFound Then
          Null;
        Else
          Select ������id
          Into v_�Է����id
          From ҩƷ�շ���¼
          Where NO = v_Outrec.No And ���� = 6 And ҩƷid + 0 = v_Avgtax.ҩƷid And Nvl(����, 0) = Nvl(v_Avgtax.����, 0) And
                �ⷿid + 0 = v_Outrec.�Է�����id And �Է�����id + 0 = v_Avgtax.�ⷿid And ���ϵ�� = -1 * v_Outrec.���ϵ�� And Rownum < 2;
        
          Update ҩƷ���
          Set ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + (v_�ֲ�� - v_ԭ���) * v_Outrec.���ϵ�� * -1
          Where �ⷿid = v_Outrec.�Է�����id And ҩƷid = v_Avgtax.ҩƷid And Nvl(����, 0) = Nvl(v_Avgtax.����, 0) And ���� = 1;
          If Sql%NotFound Then
            Insert Into ҩƷ���
              (�ⷿid, ҩƷid, ����, ����, ��������, ʵ������, ʵ�ʽ��, ʵ�ʲ��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴβ���, Ч��)
            Values
              (v_Outrec.�Է�����id, v_Avgtax.ҩƷid, v_Avgtax.����, 1, 0, 0, 0, (v_�ֲ�� - v_ԭ���) * v_Outrec.���ϵ�� * -1, Null,
               v_�ɱ���, v_Outrec.����, v_Outrec.����, v_Outrec.Ч��);
          End If;
        End If;
      End If;
    End Loop;
    Delete From ҩƷ���
    Where �ⷿid = v_Avgtax.�ⷿid And ҩƷid = v_Avgtax.ҩƷid And Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0 And
          Nvl(ʵ�ʲ��, 0) = 0;
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�������_ȫ��ƽ��;
/
--111037:��ΰ��,2018-06-24,�������Ǽ�����¼������ʱ��
Create Or Replace Procedure Zl_������������¼_Insert
(
  ����id_In   ������������¼.����id%Type,
  ��ҳid_In   ������������¼.��ҳid%Type,
  ���_In     ������������¼.���%Type,
  Ӥ������_In ������������¼.Ӥ������%Type,
  Ӥ���Ա�_In ������������¼.Ӥ���Ա�%Type,
  �������_In ������������¼.�������%Type,
  ���䷽ʽ_In ������������¼.���䷽ʽ%Type,
  ̥��״��_In ������������¼.̥��״��%Type,
  ����ʱ��_In ������������¼.����ʱ��%Type,
  ��_In     ������������¼.��%Type,
  ����_In     ������������¼.����%Type,
  Ѫ��_In     ������������¼.Ѫ��%Type,
  ��ע˵��_In ������������¼.��ע˵��%Type := Null,
  ����ʱ��_In ������������¼.����ʱ��%Type := Null
) Is
Begin
  Insert Into ������������¼
    (����id, ��ҳid, ���, Ӥ������, Ӥ���Ա�, �������, ���䷽ʽ, ̥��״��, ��, ����, Ѫ��, ����ʱ��, ����ʱ��, ��ע˵��)
  Values
    (����id_In, ��ҳid_In, ���_In, Ӥ������_In, Ӥ���Ա�_In, �������_In, ���䷽ʽ_In, ̥��״��_In, ��_In, ����_In, Ѫ��_In, ����ʱ��_In, ����ʱ��_In,
     ��ע˵��_In);

  Zl_�����Զ����_Update(����id_In, ��ҳid_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_������������¼_Insert;
/

--111037:��ΰ��,2018-06-24,�������Ǽ�����¼������ʱ��

Create Or Replace Procedure Zl_����ҽ����¼_����
(
  Id_In         ����ҽ����¼.Id%Type,
  ����Ա���_In ��Ա��.���%Type := Null,
  ����Ա����_In ��Ա��.����%Type := Null,
  ����ҽ��id_In ����ҽ����¼.Id%Type := Null,
  ����ʱ��_In   ����ҽ��״̬.����ʱ��%Type := Null
) Is
  --���ܣ�����ָ����ҽ��(δ���͵ĳ���������)
  --˵����һ����ҩ��ֻ�ܵ���һ��(������ʾ�ж���)
  --������ID_IN=��ҽ��ID
  --      ����ҽ��id_In ȡ�����������ϵĻ���ȼ�ҽ�����������Զ�ֹͣ�Ļ���ȼ�ҽ��id

  v_���ͺ�       ����ҽ������.���ͺ�%Type;
  v_����no       ������ü�¼.No%Type;
  v_��¼����     ������ü�¼.��¼����%Type;
  v_�������     Varchar2(255);
  n_�Զ�ȡ��ִ�� Number(1) := 0;
  n_�����Ϻ���ҩ Number(1) := 0;

  v_Date     Date;
  v_Count    Number;
  v_Temp     Varchar2(255);
  v_��Ա��� ��Ա��.���%Type;
  v_��Ա���� ��Ա��.����%Type;

  --����ҽ�������Ϣ
  Cursor c_Advice Is
    Select a.����id, a.�Һŵ�, a.��ҳid, a.Ӥ��, a.ҽ��״̬, a.�ϴ�ִ��ʱ��, a.ҽ������, a.�������, b.��������, a.������Դ, a.ִ�п���id, b.ִ��Ƶ��, a.������Ŀid,
           a.��ʼִ��ʱ��
    From ����ҽ����¼ A, ������ĿĿ¼ B
    Where a.������Ŀid = b.Id And a.Id = Id_In;

  r_Advice c_Advice%RowType;

  --����ҽ������ʱ��ȡ��Ӧ�ķ������ʻ�����(�շѻ��۵�)��
  --����ҽ��������NO������λ���Ҫ���ʻ��˷ѵļ�¼
  --һ��ҽ�������Ƕ���д�˷��ͼ�¼,Ҳ��һ�����Ʒ���,�ҿ���NO��ͬ
  --ֻ�ܼ�¼״̬Ϊ1�ļ�¼,����Ѿ����ʻ򲿷����ʵļ�¼,���ٴ���
  --����ֻ��۸񸸺�Ϊ�յ�,�Ա�ȡ�������
  --���"����ҩ�������Ϻ���ҩ",�򲻶���Ӧ����(������ҩ;����)���м��ʹ���,�����ǻ�û��ִ�еļ��ʵ�,��δִ�С��շѵĻ��۵���������ɾ��


  Cursor c_Rollmoney(v_���ͺ� ����ҽ������.���ͺ�%Type) Is
    Select Decode(a.��¼����, 11, 1, a.��¼����) As ��¼����, a.��¼״̬, a.No, a.���, a.ִ��״̬ As ����ִ��, c.ִ��״̬ As ҽ��ִ��, c.ִ�в���id, b.���˿���id,
           b.�������, i.��������
    From ������ü�¼ A, ����ҽ����¼ B, ����ҽ������ C, ������ĿĿ¼ I
    Where c.ҽ��id = b.Id And c.���ͺ� = v_���ͺ� And (b.Id = Id_In Or b.���id = Id_In) And a.ҽ����� = b.Id And a.��¼״̬ In (0, 1) And
          a.No = c.No And (a.��¼���� = c.��¼���� Or a.��¼���� = 11 And c.��¼���� = 1) And b.������Ŀid = i.Id And a.�۸񸸺� Is Null And
          (n_�����Ϻ���ҩ = 0 Or
          n_�����Ϻ���ҩ = 1 And
          Not (Exists (Select 1
                        From ������ü�¼ D
                        Where d.ҽ����� = b.Id And d.��¼״̬ In (0, 1) And d.No = c.No And
                              (d.��¼���� = c.��¼���� Or d.��¼���� = 11 And c.��¼���� = 1) And d.�շ���� In ('5', '6', '7'))) Or
          Nvl(a.ִ��״̬, 0) = 0 And Not (a.��¼���� = 1 And a.��¼״̬ <> 0))
    Order By a.��¼����, a.No, a.���, a.�շ�ϸĿid;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --���ҽ��״̬�Ƿ���ȷ:��������
  Open c_Advice;
  Fetch c_Advice
    Into r_Advice;

  --��ǰ������Ա
  If ����Ա���_In Is Not Null And ����Ա����_In Is Not Null Then
    v_��Ա��� := ����Ա���_In;
    v_��Ա���� := ����Ա����_In;
  Else
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_��Ա��� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
    v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  End If;

  --����Ƿ�����Һ��Һ��¼�����Ƿ��Ѿ�����
  Select Count(1) Into v_Count From ��Һ��ҩ��¼ Where �Ƿ����� = 1 And ҽ��id = Id_In;
  If v_Count > 0 Then
    v_Error := 'ҽ��"' || r_Advice.ҽ������ || '"����ҺҩƷ���Ѿ�����Һ���������������������ϡ�';
    Raise Err_Custom;
  End If;

  If r_Advice.�Һŵ� Is Null And r_Advice.������Դ <> 3 Then
    If r_Advice.ҽ��״̬ In (4, 8, 9) Then
      v_Error := 'ҽ��"' || r_Advice.ҽ������ || '"�Ѿ������ϻ�ֹͣ�����������ϡ�';
      Raise Err_Custom;
    Elsif r_Advice.�ϴ�ִ��ʱ�� Is Not Null Then
      v_Error := 'ҽ��"' || r_Advice.ҽ������ || '"�Ѿ����ͣ����ܱ����ϡ�';
      Raise Err_Custom;
    End If;
  
    --�����Ի���ȼ����뷢�ͣ�У�Ժ�Ϳ������Զ��Ʒѣ����ϼ��������϶�Ӧ��ֹͣ���̴���
    If r_Advice.������� = 'H' And r_Advice.�������� = '1' And r_Advice.ִ��Ƶ�� = '2' And Nvl(r_Advice.Ӥ��, 0) = 0 Then
      --(��ȡ�������ڴ����޷���Ժ�����������ţ�45977)a.��ʼʱ���ǵ���֮ǰ�ģ�˵������Ч���Զ����ü��㣩�����������ϡ�
      --ҽ����ʱ��ֻ��ȷ���˷��ӣ����Ա䶯��¼�Ŀ�ʼʱ��Ҫȥ�������Ƚϡ�
      v_Count := 0;
      Begin
        Select b.��ֹʱ��
        Into v_Date
        From ���˱䶯��¼ B, ����ҽ���Ƽ� C
        Where b.����id = r_Advice.����id And b.��ҳid = r_Advice.��ҳid And c.ҽ��id = Id_In And c.�շ�ϸĿid = b.����ȼ�id And
              b.��ʼԭ�� = 6 And b.���Ӵ�λ = 0 And
              To_Char(b.��ʼʱ��, 'yyyy-mm-dd hh24:mi') = To_Char(r_Advice.��ʼִ��ʱ��, 'yyyy-mm-dd hh24:mi');
      Exception
        When Others Then
          v_Count := 1;
      End;
      If v_Count = 0 Then
        --d.�����������䶯����
        If v_Date Is Not Null Then
          v_Error := '���ڻ���ȼ�ҽ����Ч���Ѿ������������䶯��¼,�������ϸ�ҽ����';
          Raise Err_Custom;
        Else
          --������Ҫ�Զ����õĻ���ȼ��������ԭ������ȼ���ͬ���ó�������䶯��¼
          If Nvl(����ҽ��id_In, 0) <> 0 Then
            Delete ����ҽ��״̬ Where ҽ��id = ����ҽ��id_In And �������� In (8, 9);
            Select ��������
            Into v_Count
            From (Select �������� From ����ҽ��״̬ Where ҽ��id = ����ҽ��id_In Order By ����ʱ�� Desc)
            Where Rownum < 2;
            Update ����ҽ����¼
            Set ҽ��״̬ = v_Count, ִ����ֹʱ�� = Null, ͣ��ҽ�� = Null, ͣ��ʱ�� = Null, ȷ��ͣ��ʱ�� = Null, ȷ��ͣ����ʿ = Null
            Where ID = ����ҽ��id_In;
            --�ų�����Ƶ���Ĳ���
            Select Count(a.Id)
            Into v_Count
            From ����ҽ����¼ A, �����շѹ�ϵ B, ������ҳ C
            Where a.������Ŀid = b.������Ŀid And c.����ȼ�id = b.�շ���Ŀid And c.����id = a.����id And c.��ҳid = a.��ҳid And
                  a.Id = ����ҽ��id_In;
          End If;
          If v_Count = 0 Then
            --c.����ȼ������һ���䶯
            Zl_���˱䶯��¼_Undo(r_Advice.����id, r_Advice.��ҳid, v_��Ա���, v_��Ա����, '1', Null, Null, '����ȼ��䶯');
          End If;
        End If;
      Else
        --�ָ����һ�α��Զ�ֹͣ�Ļ���ȼ�
        If Nvl(����ҽ��id_In, 0) <> 0 Then
          Delete ����ҽ��״̬ Where ҽ��id = ����ҽ��id_In And �������� In (8, 9);
          Select ��������
          Into v_Count
          From (Select �������� From ����ҽ��״̬ Where ҽ��id = ����ҽ��id_In Order By ����ʱ�� Desc)
          Where Rownum < 2;
          Update ����ҽ����¼
          Set ҽ��״̬ = v_Count, ִ����ֹʱ�� = Null, ͣ��ҽ�� = Null, ͣ��ʱ�� = Null, ȷ��ͣ��ʱ�� = Null, ȷ��ͣ����ʿ = Null
          Where ID = ����ҽ��id_In;
        Else
          --������Ժʱָ���Ļ��������ı䶯��¼��ҽ���¿������ı䶯��¼��ͬ������Ҫ���ж�
          Select Count(a.Id)
          Into v_Count
          From ���˱䶯��¼ A
          Where a.����id = r_Advice.����id And a.��ҳid = r_Advice.��ҳid And a.��ʼԭ�� = 6;
          If v_Count <> 0 Then
            --b.�������ǰ�Ļ���ȼ���ͬ����У��ʱû�в�������ȼ��䶯,��������ȼ�ֹͣ�䶯
            Zl_���˱䶯��¼_Nurse(r_Advice.����id, r_Advice.��ҳid, Null, Sysdate, v_��Ա���, v_��Ա����);
          End If;
        End If;
      End If;
    End If;
  Else
    If r_Advice.ҽ��״̬ <> 8 Then
      v_Error := 'ҽ��"' || r_Advice.ҽ������ || '"��δ���ͻ��Ѿ����ϡ�';
      Raise Err_Custom;
    End If;
    --ҽ�������ж�
    Select Count(1)
    Into v_Count
    From ����ҽ������ A, ����ҽ����¼ B
    Where a.ҽ��id = b.Id And (b.Id = Id_In Or b.���id = Id_In);
    If v_Count <> 0 Then
      v_Error := 'ҽ��"' || r_Advice.ҽ������ || '"���ڸ��ӷ��ã��������ϡ�';
      Raise Err_Custom;
    End If;
  
    Begin
      --ҽ��IDΪ����ֵ������ҽ����һ�������˵�,�����޷��͡�
      Select Distinct ���ͺ�
      Into v_���ͺ�
      From ����ҽ������
      Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = Id_In Or ���id = Id_In);
    Exception
      When Others Then
        v_���ͺ� := Null;
    End;
  
    Select Zl_To_Number(Nvl(zl_GetSysParameter(68), 0)) Into n_�����Ϻ���ҩ From Dual;
    Select Zl_To_Number(Nvl(zl_GetSysParameter('���ﱾ���Զ�ִ��', '1252'), 0)) Into n_�Զ�ȡ��ִ�� From Dual;
    If n_�Զ�ȡ��ִ�� = 1 And v_���ͺ� Is Not Null Then
      --�ȸ���ҽ���ͷ��õ�ִ��״̬����Ϊ�������жϣ��Լ�����Zl_������ʼ�¼_Delete���м��
      For Rc In (Select a.ҽ��id, a.ִ�в���id
                 From ����ҽ������ A, ����ҽ����¼ B
                 Where a.ҽ��id = b.Id And (b.Id = Id_In Or b.���id = Id_In) And a.ִ�в���id = b.���˿���id) Loop
        Zl_����ҽ��ִ��_Cancel(Rc.ҽ��id, v_���ͺ�, Null, 1, Rc.ִ�в���id);
      End Loop;
    End If;
  
    --����ҽ��ֻ���ܷ���һ��
    --�����˷�ʱ���м�飬��Ϊ����ҽ��û�з��ã�����Ҫ���һ��ִ��״̬
    Select Count(*)
    Into v_Count
    From ����ҽ������ A, ����ҽ����¼ B, ������ĿĿ¼ I
    Where a.ҽ��id = b.Id And b.������Ŀid = i.Id And a.ִ��״̬ In (1, 3) And (b.Id = Id_In Or b.���id = Id_In) And
          (n_�����Ϻ���ҩ = 0 Or
          n_�����Ϻ���ҩ = 1 And Not (b.������� In ('5', '6', '7') Or b.������� = 'E' And i.�������� In ('2', '3', '4')));
    If v_Count > 0 Then
      v_Error := '��ҽ���Ѿ�ִ�л�����ִ�У��������ϡ�';
      Raise Err_Custom;
    End If;
  End If;

  If ����ʱ��_In Is Null Then
    Select Sysdate Into v_Date From Dual;
  Else
    v_Date := ����ʱ��_In;
  End If;

  Update ����ҽ����¼ Set ҽ��״̬ = 4 Where ID = Id_In Or ���id = Id_In;

  Insert Into ����ҽ��״̬
    (ҽ��id, ��������, ������Ա, ����ʱ��)
    Select ID, 4, v_��Ա����, v_Date From ����ҽ����¼ Where ID = Id_In Or ���id = Id_In;

  --סԺҽ������ʱ,δ��ӡ�������,ȱʡ����Ϊ���δ�ӡ
  If r_Advice.�Һŵ� Is Null Then
    Select Count(*)
    Into v_Count
    From ����ҽ����ӡ
    Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = Id_In Or ���id = Id_In);
    If Nvl(v_Count, 0) = 0 Then
      Zl_����ҽ����¼_���δ�ӡ(Id_In, 1);
    End If;
    If Nvl(r_Advice.Ӥ��, 0) > 0 And r_Advice.�������� = '11' Then
      Update ������������¼
      Set ����ʱ�� = Null
      Where ����id = r_Advice.����id And Nvl(��ҳid, 0) = Nvl(r_Advice.��ҳid, 0) And ��� = Nvl(r_Advice.Ӥ��, 0);
    End If;
  Else
    --����ҽ��(����)����ʱ����Ҫ�����������:ֻ��һ�η���
    --���˻��ۻ���ʷ���
    If v_���ͺ� Is Not Null Then
      --������ҽ���ķ���ɾ��������(��һ��ҽ�������в�ͬNO����)
      --������ʣ����ԭʼ�����ѱ�����(�򲿷�����),���ù��������ж�
      --���ﻮ�ۣ�������շѣ�������ɾ��
      v_����no   := Null;
      v_������� := Null;
      For r_Rollmoney In c_Rollmoney(v_���ͺ�) Loop
        If Nvl(r_Rollmoney.ҽ��ִ��, 0) In (1, 3) Then
          --1-��ȫִ��;3-����ִ��
          v_Error := 'ҽ��"' || r_Advice.ҽ������ || '"�Ѿ�ִ�л�����ִ�У��������ϡ�';
          Raise Err_Custom;
        End If;
        If Nvl(r_Rollmoney.����ִ��, 0) In (1, 2) Then
          --1-��ȫִ��;2-����ִ��
          v_Error := 'ҽ�����õ���"' || r_Rollmoney.No || '"�е������Ѿ�ȫ���򲿷�ִ�У��������ϡ�';
          Raise Err_Custom;
        End If;
        If r_Rollmoney.����ִ�� = 9 Then
          v_Error := 'ҽ�����õ���"' || r_Rollmoney.No || '"�е��շѽ�������쳣���������ϡ�';
          Raise Err_Custom;
        End If;
        v_Count := 1;
        If r_Rollmoney.��¼���� = 1 And r_Rollmoney.��¼״̬ <> 0 Then
          If 1 = n_�����Ϻ���ҩ And r_Rollmoney.������� = 'E' And r_Rollmoney.�������� In ('2', '3', '4') Then
            v_Count := 0;
          Else
            v_Error := 'ҽ�����õ���"' || r_Rollmoney.No || '"�Ѿ��շѣ��������ϡ�';
            Raise Err_Custom;
          End If;
        End If;
        If 1 = v_Count Then
          If Nvl(v_����no, '��') <> r_Rollmoney.No Then
            If v_������� Is Not Null And v_����no Is Not Null Then
              v_������� := Substr(v_�������, 2);
              If v_��¼���� = 1 Then
                Zl_���ﻮ�ۼ�¼_Delete(v_����no, v_�������);
              Elsif v_��¼���� = 2 Then
                Zl_������ʼ�¼_Delete(v_����no, v_�������, v_��Ա���, v_��Ա����);
              End If;
            End If;
            v_������� := Null;
          End If;
          v_��¼���� := r_Rollmoney.��¼����;
          v_����no   := r_Rollmoney.No;
          v_������� := v_������� || ',' || r_Rollmoney.���;
        End If;
      End Loop;
      If v_������� Is Not Null And v_����no Is Not Null Then
        v_������� := Substr(v_�������, 2);
        If v_��¼���� = 1 Then
          Zl_���ﻮ�ۼ�¼_Delete(v_����no, v_�������);
        Elsif v_��¼���� = 2 Then
          Zl_������ʼ�¼_Delete(v_����no, v_�������, v_��Ա���, v_��Ա����);
        End If;
      End If;
    
      --���"����ҩ�������Ϻ���ҩ"�����Ӧ�ĸ�ҩ;����������Ϊδִ�У��Ա��˷�
      If n_�����Ϻ���ҩ = 1 Then
        Update ������ü�¼
        Set ִ��״̬ = 0
        Where ִ��״̬ = 1 And ҽ����� = Id_In And Exists
         (Select 1
               From ����ҽ����¼ A, ������ĿĿ¼ B
               Where a.������Ŀid = b.Id And b.��� = 'E' And b.�������� In ('2', '3', '4') And a.Id = Id_In);
      End If;
    
      --����ҽ�����ͼ�¼(��ִ�м�¼)
      Delete From ����ҽ��ִ�� Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = Id_In Or ���id = Id_In);
      Delete From ����ҽ������ Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = Id_In Or ���id = Id_In);
    
      --��������ҽ���Ĵ���
      If r_Advice.������� = 'Z' And Nvl(r_Advice.��������, '0') <> '0' And Nvl(r_Advice.Ӥ��, 0) = 0 Then
        If r_Advice.�������� = '1' And r_Advice.ִ�п���id Is Not Null Then
          --����ҽ��
          Select Count(*)
          Into v_Count
          From ������ҳ
          Where ����id = r_Advice.����id And Nvl(��ҳid, 0) = 0 And ��Ժ����id = r_Advice.ִ�п���id And �������� In (1, 2);
          If v_Count = 1 Then
            Zl_��Ժ������ҳ_Delete(r_Advice.����id, 0);
          End If;
        Elsif r_Advice.�������� = '2' And r_Advice.ִ�п���id Is Not Null Then
          --סԺҽ��
          Select Count(*)
          Into v_Count
          From ������ҳ
          Where ����id = r_Advice.����id And Nvl(��ҳid, 0) = 0 And ��Ժ����id = r_Advice.ִ�п���id And Nvl(��������, 0) = 0;
          If v_Count = 1 Then
            Zl_��Ժ������ҳ_Delete(r_Advice.����id, 0);
          End If;
        End If;
      End If;
    End If;
  End If;

  --ɾ�������ǼǼ�¼
  If r_Advice.������� = 'E' And r_Advice.�������� = '1' Then
    --Update ����ҽ����¼ Set Ƥ�Խ��=Null Where ID=ID_IN; --��������Ƥ�Խ��
    --ɾ���������ļ�¼��������¼��������Ϊ����ҽ���Ƿ����ϣ����˶Ը�ҩ����
    For r_Test In (Select ����ʱ�� From ����ҽ��״̬ Where ҽ��id = Id_In And �������� = 10) Loop
      Delete From ���˹�����¼
      Where ����id = r_Advice.����id And ��¼��Դ = 2 And Nvl(��ҳid, 0) = Nvl(r_Advice.��ҳid, 0) And ��¼ʱ�� = r_Test.����ʱ��;
    End Loop;
  End If;

  Close c_Advice;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ҽ����¼_����;
/

--115765:����,2018-06-28,�ɱ��۾��ȴ���
--127911,��ҵ��,2018-06-28,��ֵ����ȡ����ⷿ�ĳɱ���
Create Or Replace Procedure Zl_�����շ���¼_��������
(
  Partid_In   In ҩƷ�շ���¼.�ⷿid%Type,
  Bill_In     In ҩƷ�շ���¼.����%Type,
  No_In       In ҩƷ�շ���¼.No%Type,
  People_In   In ҩƷ�շ���¼.�����%Type,
  ��ҩ��_In   In ҩƷ�շ���¼.��ҩ��%Type := Null,
  У����_In   In ҩƷ�շ���¼.������%Type := Null,
  ��ҩ��ʽ_In In ҩƷ�շ���¼.��ҩ��ʽ%Type := 1,
  ��ҩʱ��_In In ҩƷ�շ���¼.�������%Type := Null
) Is
  --���¼�����
  Cursor c_Modifybill Is
    Select a.Id, a.ҩƷid, a.������id, a.���ϵ��, a.����id, a.��ҩ��λid, a.��������, a.��׼�ĺ�, a.���Ч��, a.Ч��, a.����,
           Nvl(a.ʵ������, 0) * Nvl(a.����, 1) ����, Nvl(a.���۽��, 0) ���, Nvl(a.����, 0) ����, a.����, 2 As ������Դ, a.�ⷿid, a.�ڲ�����, a.��Ʒ����
    From ҩƷ�շ���¼ A, סԺ���ü�¼ B
    Where a.No = No_In And a.���� = Bill_In And (a.�ⷿid + 0 = Partid_In Or a.�ⷿid Is Null) And Nvl(a.ժҪ, '�ܷ���') <> '�ܷ�' And
          a.����id = b.Id And b.ִ��״̬ <> 1 And Mod(a.��¼״̬, 3) = 1 And a.����� Is Null
    Union All
    Select a.Id, a.ҩƷid, a.������id, a.���ϵ��, a.����id, a.��ҩ��λid, a.��������, a.��׼�ĺ�, a.���Ч��, a.Ч��, a.����,
           Nvl(a.ʵ������, 0) * Nvl(a.����, 1) ����, Nvl(a.���۽��, 0) ���, Nvl(a.����, 0) ����, a.����, 1 As ������Դ, a.�ⷿid, a.�ڲ�����, a.��Ʒ����
    From ҩƷ�շ���¼ A, ������ü�¼ B
    Where a.No = No_In And a.���� = Bill_In And (a.�ⷿid + 0 = Partid_In Or a.�ⷿid Is Null) And Nvl(a.ժҪ, '�ܷ���') <> '�ܷ�' And
          a.����id = b.Id And b.ִ��״̬ <> 1 And Mod(a.��¼״̬, 3) = 1 And a.����� Is Null
    Order By ҩƷid;

  v_Modifybill c_Modifybill%RowType;

  --ֻ������
  n_����� ҩƷ���.ʵ�ʽ��%Type;
  n_����� ҩƷ���.ʵ�ʲ��%Type;
  n_�����   ��������.ָ�������%Type;

  --��д����
  n_�ɱ����       ҩƷ�շ���¼.�ɱ����%Type;
  n_�ɱ���         ҩƷ�շ���¼.�ɱ���%Type;
  n_ʵ�ʲ��       ҩƷ�շ���¼.���%Type;
  d_����ʱ��       ҩƷ�շ���¼.�������%Type;
  n_�շ��뷢�Ϸ��� Number(1);
  n_С��           Number(1);
  n_�ɱ���С��           Number(1);
  v_���no         ҩƷ�շ���¼.No%Type;
  v_���ⷿid     ҩƷ�շ���¼.�ⷿid%Type := 0;
  v_������Ϣ       Varchar2(200);
  n_����ⷿ       ҩƷ���.�ⷿid%Type;
  n_���           Number;
  n_ƽ���ɱ���     ҩƷ���.ƽ���ɱ���%Type;
Begin
  If ��ҩʱ��_In Is Null Then
    Select Sysdate Into d_����ʱ�� From Dual;
  Else
    d_����ʱ�� := ��ҩʱ��_In;
  End If;

  Begin
    Select 0 Into n_�շ��뷢�Ϸ��� From δ��ҩƷ��¼ Where ���� = Bill_In And NO = No_In And �ⷿid + 0 = Partid_In;
  Exception
    When Others Then
      n_�շ��뷢�Ϸ��� := 1;
  End;

  --��ȡ���С��λ��
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_С�� From Dual;
  --��ȡ�ɱ���С��λ��
  Select Zl_To_Number(Nvl(zl_GetSysParameter(157), '2')) Into n_�ɱ���С�� From Dual;

  --��д�ѷ��ϴ�������ҩ��
  Update ҩƷ�շ���¼
  Set ��ҩ�� = Decode(��ҩ��_In, Null, ��ҩ��, ��ҩ��_In)
  Where NO = No_In And ���� = Bill_In And (�ⷿid + 0 = Partid_In Or �ⷿid Is Null) And Mod(��¼״̬, 3) = 1 And ����� Is Not Null;

  --���¼���ɱ��ۡ��ɱ������۽����
  For v_Modifybill In c_Modifybill Loop
    --��ֵ�����������ģʽ
    If n_����ⷿ Is Null Then
      Begin
        Select �ⷿid
        Into n_����ⷿ
        From ҩƷ�շ���¼
        Where ���� = 21 And ������� Is Null And ҩƷid = v_Modifybill.ҩƷid And Nvl(����, 0) = v_Modifybill.���� And
              ����id = v_Modifybill.����id And Rownum = 1;
      Exception
        When Others Then
          n_����ⷿ := 0;
      End;
    End If;
  
    If n_����ⷿ = 0 Then
      --��ͨģʽȡ���ϲ��ż۸�
      n_�ɱ��� := Round(Zl_Fun_Getoutcost(v_Modifybill.ҩƷid, v_Modifybill.����, Partid_In), n_�ɱ���С��);
    Else
      --��ֵ�����������ģʽȡ����ⷿ�۸�
      n_�ɱ��� := Round(Zl_Fun_Getoutcost(v_Modifybill.ҩƷid, v_Modifybill.����, n_����ⷿ), n_�ɱ���С��);
    End If;
    n_�ɱ���� := Round(n_�ɱ��� * v_Modifybill.����, n_С��);
    n_ʵ�ʲ�� := Round(Nvl(v_Modifybill.���, 0) - n_�ɱ����, n_С��);
  
    --����ҩƷ�շ���¼�����۽��ɱ������
    Update ҩƷ�շ���¼ Set �ɱ��� = n_�ɱ���, �ɱ���� = n_�ɱ����, ��� = n_ʵ�ʲ�� Where ID = v_Modifybill.Id;
  
    If n_�շ��뷢�Ϸ��� = 1 Then
      Update ҩƷ���
      Set �������� = Nvl(��������, 0) - Nvl(v_Modifybill.����, 0), ʵ������ = Nvl(ʵ������, 0) - Nvl(v_Modifybill.����, 0),
          ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) - Nvl(v_Modifybill.���, 0), ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) - n_ʵ�ʲ��,
          �ϴβɹ��� = Decode(�ϴβɹ���, Null, n_�ɱ���, �ϴβɹ���), ƽ���ɱ��� = Decode(ƽ���ɱ���, Null, n_�ɱ���, ƽ���ɱ���),
          ��Ʒ���� = Decode(��Ʒ����, Null, v_Modifybill.��Ʒ����, ��Ʒ����), �ڲ����� = Decode(�ڲ�����, Null, v_Modifybill.�ڲ�����, �ڲ�����),
          Ч�� = Decode(Ч��, Null, v_Modifybill.Ч��, Ч��), �ϴ����� = Decode(�ϴ�����, Null, v_Modifybill.����, �ϴ�����),
          �ϴ��������� = Decode(�ϴ���������, Null, v_Modifybill.��������, �ϴ���������), �ϴβ��� = Decode(�ϴβ���, Null, v_Modifybill.����, �ϴβ���)
      Where �ⷿid + 0 = Partid_In And ҩƷid = v_Modifybill.ҩƷid And ���� = 1 And Nvl(����, 0) = v_Modifybill.����;
    Else
      Update ҩƷ���
      Set ʵ������ = Nvl(ʵ������, 0) - Nvl(v_Modifybill.����, 0), ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) - Nvl(v_Modifybill.���, 0),
          ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) - n_ʵ�ʲ��, �ϴβɹ��� = Decode(�ϴβɹ���, Null, n_�ɱ���, �ϴβɹ���),
          ƽ���ɱ��� = Decode(ƽ���ɱ���, Null, n_�ɱ���, ƽ���ɱ���), ��Ʒ���� = Decode(��Ʒ����, Null, v_Modifybill.��Ʒ����, ��Ʒ����),
          �ڲ����� = Decode(�ڲ�����, Null, v_Modifybill.�ڲ�����, �ڲ�����), Ч�� = Decode(Ч��, Null, v_Modifybill.Ч��, Ч��),
          �ϴ����� = Decode(�ϴ�����, Null, v_Modifybill.����, �ϴ�����), �ϴ��������� = Decode(�ϴ���������, Null, v_Modifybill.��������, �ϴ���������),
          �ϴβ��� = Decode(�ϴβ���, Null, v_Modifybill.����, �ϴβ���)
      Where �ⷿid + 0 = Partid_In And ҩƷid = v_Modifybill.ҩƷid And ���� = 1 And Nvl(����, 0) = v_Modifybill.����;
    End If;
  
    If Sql%RowCount = 0 Then
      If n_�շ��뷢�Ϸ��� = 1 Then
        Insert Into ҩƷ���
          (�ⷿid, ҩƷid, ����, ����, ʵ������, ʵ�ʽ��, ʵ�ʲ��, Ч��, ���Ч��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴ���������, �ϴβ���, ��׼�ĺ�, ƽ���ɱ���)
        Values
          (Partid_In, v_Modifybill.ҩƷid, v_Modifybill.����, 1, 0 - Nvl(v_Modifybill.����, 0), 0 - Nvl(v_Modifybill.���, 0),
           0 - n_ʵ�ʲ��, v_Modifybill.Ч��, v_Modifybill.���Ч��, v_Modifybill.��ҩ��λid, n_�ɱ���, v_Modifybill.����,
           v_Modifybill.��������, v_Modifybill.����, v_Modifybill.��׼�ĺ�, n_�ɱ���);
      Else
        Insert Into ҩƷ���
          (�ⷿid, ҩƷid, ����, ����, ��������, ʵ������, ʵ�ʽ��, ʵ�ʲ��, Ч��, ���Ч��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴ���������, �ϴβ���, ��׼�ĺ�, ƽ���ɱ���)
        Values
          (Partid_In, v_Modifybill.ҩƷid, v_Modifybill.����, 1, 0 - Nvl(v_Modifybill.����, 0), 0 - Nvl(v_Modifybill.����, 0),
           0 - Nvl(v_Modifybill.���, 0), 0 - n_ʵ�ʲ��, v_Modifybill.Ч��, v_Modifybill.���Ч��, v_Modifybill.��ҩ��λid, n_�ɱ���,
           v_Modifybill.����, v_Modifybill.��������, v_Modifybill.����, v_Modifybill.��׼�ĺ�, n_�ɱ���);
      End If;
    End If;
  
    Delete ҩƷ���
    Where �ⷿid + 0 = Partid_In And ҩƷid = v_Modifybill.ҩƷid And Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And
          Nvl(ʵ�ʽ��, 0) = 0 And Nvl(ʵ�ʲ��, 0) = 0 And ���� = 1;
  
    --���²��˷��ü�¼��ִ��״̬(��ִ��)
    If v_Modifybill.������Դ = 2 Then
      Update סԺ���ü�¼
      Set ִ��״̬ = 1, ִ���� = Decode(People_In, Null, Zl_Username, People_In), ִ��ʱ�� = ��ҩʱ��_In
      Where ID = v_Modifybill.����id;
    Else
      Update ������ü�¼
      Set ִ��״̬ = 1, ִ���� = Decode(People_In, Null, Zl_Username, People_In), ִ��ʱ�� = ��ҩʱ��_In
      Where ID = v_Modifybill.����id;
    End If;
    --д�����
    Update ҩƷ�շ���¼
    Set �ⷿid = Partid_In, ��ҩ�� = Decode(��ҩ��_In, Null, ��ҩ��, ��ҩ��_In), ������ = Decode(У����_In, Null, ������, У����_In),
        ����� = People_In, ������� = d_����ʱ��, ��ҩ��ʽ = ��ҩ��ʽ_In
    Where ID = v_Modifybill.Id;
    --�����������
    Zl_�����շ���¼_��������(v_Modifybill.Id);
  
    If n_����ⷿ > 0 Then
      --��˱�������������ⷿ���������ⵥ��
      For v_���� In (Select ���, NO, �ⷿid, ҩƷid, Nvl(����, 0) As ����, ʵ������, �ɱ���, �ɱ����, ���۽��, ���, ������id
                   From ҩƷ�շ���¼
                   Where ���� = 21 And ������� Is Null And ҩƷid = v_Modifybill.ҩƷid And Nvl(����, 0) = v_Modifybill.���� And
                         ����id = v_Modifybill.����id) Loop
      
        Update ҩƷ�շ���¼
        Set ���ܷ�ҩ�� = v_Modifybill.Id
        Where ���� = 21 And ������� Is Null And ҩƷid = v_Modifybill.ҩƷid And Nvl(����, 0) = v_Modifybill.���� And
              ����id = v_Modifybill.����id;
      
        Zl_������������_Verify(v_����.���, v_����.No, v_����.�ⷿid, v_����.ҩƷid, v_����.����, v_����.ʵ������, v_����.�ɱ���, v_����.�ɱ����, v_����.���۽��,
                         v_����.���, v_����.������id, People_In, d_����ʱ��);
      End Loop;
    
      --�����������������Ĳֿ���⹺��ⵥ��
      For v_��� In (Select NO, ���, ��ҩ��λid, ҩƷid, ����, ����, ��������, Ч��, �������, ���Ч��, ʵ������, �ɱ���, �ɱ����, ����, ���ۼ�, ���۽��, ���, ժҪ,
                          ע��֤��, Nvl(����, 0) As ����, ��Ʒ����, �ڲ�����
                   From ҩƷ�շ���¼
                   Where ���� = 21 And ������� Is Not Null And ҩƷid = v_Modifybill.ҩƷid And Nvl(����, 0) = v_Modifybill.���� And
                         ����id = v_Modifybill.����id And ���ܷ�ҩ�� = v_Modifybill.Id) Loop
        Begin
          Select �ⷿid Into v_���ⷿid From ����ⷿ���� Where ����id = v_Modifybill.�ⷿid;
        Exception
          When Others Then
            v_���ⷿid := 0;
        End;
      
        If v_���ⷿid > 0 Then
        
          --ͬһ�ŷ��ϵ���������ⵥ��NOҪһ��
          Select Max(NO), Max(���) + 1
          Into v_���no, n_���
          From ҩƷ�շ���¼
          Where ���� = 15 And ������� Is Null And ��ҩ��λid = v_���.��ҩ��λid And
                ����id In (Select Distinct ����id
                         From ҩƷ�շ���¼
                         Where ���� = 21 And ������� Is Not Null And
                               NO = (Select Distinct NO
                                     From ҩƷ�շ���¼
                                     Where ���� = 21 And ������� Is Not Null And ����id = v_Modifybill.����id));
        
          If v_���no Is Null Or v_���no = '' Then
            --������NOΪNull, �����µ���ⵥNO
            v_���no := Nextno(68, v_���ⷿid);
            n_���   := 1;
          End If;
        
          Begin
            If v_Modifybill.������Դ = 1 Then
              Select b.���� || ',' || a.���� || ',' || a.��ʶ�� || ',' || '' As ������Ϣ
              Into v_������Ϣ
              From ������ü�¼ A, ���ű� B
              Where a.���˿���id = b.Id And a.Id = v_Modifybill.����id;
            Else
              Select b.���� || ',' || a.���� || ',' || a.��ʶ�� || ',' || a.���� As ������Ϣ
              Into v_������Ϣ
              From סԺ���ü�¼ A, ���ű� B
              Where a.���˿���id = b.Id And a.Id = v_Modifybill.����id;
            End If;
          Exception
            When Others Then
              v_������Ϣ := '';
          End;
        
          Zl_�����⹺_Insert(v_���no, n_���, v_���ⷿid, v_���.��ҩ��λid, v_���.ҩƷid, v_���.����, v_���.����, v_���.��������, v_���.Ч��,
                         v_���.�������, v_���.���Ч��, v_���.ʵ������, v_���.�ɱ���, v_���.�ɱ����, v_���.����, v_���.���ۼ�, v_���.���۽��, v_���.���,
                         Null, '���Զ����ˡ�' || v_���.ժҪ, v_���.ע��֤��, People_In, Null, Null, Null, Null, d_����ʱ��, Null, Null,
                         v_���.����, 1, v_������Ϣ, v_���.��Ʒ����, v_���.�ڲ�����, v_Modifybill.����id);
        End If;
      End Loop;
    End If;
  End Loop;

  --���»�ɾ��δ��ҩƷ��¼
  Delete δ��ҩƷ��¼ Where NO = No_In And ���� = Bill_In And (�ⷿid + 0 = Partid_In Or �ⷿid Is Null);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����շ���¼_��������;
/

--127738,��ҵ��,2018-07-02,������˲�������Ϊ0���
Create Or Replace Procedure Zl_ҩƷ�⹺_Verify
(
  Newno_In    In ҩƷ�շ���¼.No%Type := Null,
  Oldno_In    In ҩƷ�շ���¼.No%Type := Null,
  �����_In   In ҩƷ�շ���¼.�����%Type := Null,
  �������_In In ҩƷ�շ���¼.�������%Type := Sysdate
) Is
  Err_Isverified Exception;
  Err_Isbatch Exception;
  v_Druginf        Varchar2(50); --ԭ���������ڷ�����ҩƷ��Ϣ
  v_��ҩ��λid     ҩƷ�շ���¼.��ҩ��λid%Type;
  v_��Ʊ���       Ӧ����¼.��Ʊ���%Type;
  v_��������       ҩƷ���.��������%Type;
  v_ʱ�۷���       Number(1);
  n_ԭ�ɱ���       ҩƷ�շ���¼.�ɱ���%Type;
  v_Newno          ҩƷ�շ���¼.No%Type;
  n_New���        Number;
  n_�շ�id         ҩƷ�շ���¼.Id%Type;
  n_������         ҩƷ�շ���¼.���۽��%Type;
  n_������id     ҩƷ�շ���¼.������id%Type;
  n_�ۼ�������id ҩƷ�շ���¼.������id%Type;
  n_���ϵ��       ҩƷ�շ���¼.���ϵ��%Type;
  n_ƽ���ɱ���     ҩƷ���.ƽ���ɱ���%Type;
  n_�����ɱ���     ҩƷ�շ���¼.�ɱ���%Type;
  n_�����ۼ�       ҩƷ�շ���¼.���ۼ�%Type;
  v_Billno         ҩƷ�շ���¼.No%Type;

  Cursor c_ҩƷ�շ���¼ Is
    Select a.Id, a.���ۼ�, a.ʵ������, a.���۽��, a.���, a.�ⷿid, a.ҩƷid, a.����, a.��ҩ��λid, a.�ɱ���, a.����, a.Ч��, a.����, a.������id, a.��������,
           a.��׼�ĺ�, Nvl(b.�Ƿ���, 0) As ʱ��, Nvl(a.��ҩ��ʽ, 0) As �˿�, a.���Ч��, a.����, Nvl(a.�ƻ�id, 0) As �ƻ�id,
           Nvl(a.����id, 0) As ����id, a.���
    From ҩƷ�շ���¼ A, �շ���ĿĿ¼ B
    Where a.ҩƷid = b.Id And a.No = Newno_In And a.���� = 1 And a.��¼״̬ = 1
    Order By a.ҩƷid, a.����;
Begin

  n_New��� := 1;

  Select b.Id, b.ϵ��
  Into n_������id, n_���ϵ��
  From ҩƷ�������� A, ҩƷ������ B
  Where a.���id = b.Id And a.���� = 5 And Rownum < 2;

  Select ���id Into n_�ۼ�������id From ҩƷ�������� Where ���� = 13 And Rownum < 2;

  Update ҩƷ�շ���¼
  Set ����� = Nvl(�����_In, �����), ������� = �������_In
  Where NO = Newno_In And ���� = 1 And ��¼״̬ = 1 And ����� Is Null;

  If Sql%RowCount = 0 Then
    Raise Err_Isverified;
  End If;

  --��Ҫ���ԭ���������ڷ�����ҩƷ�����ܶ������
  Begin
    Select Distinct '(' || i.���� || ')' || Nvl(n.����, i.����) As ҩƷ��Ϣ
    Into v_Druginf
    From ҩƷ�շ���¼ A, ҩƷ��� B, �շ���ĿĿ¼ I, �շ���Ŀ���� N
    Where a.ҩƷid = b.ҩƷid And a.ҩƷid = i.Id And a.ҩƷid = n.�շ�ϸĿid(+) And n.����(+) = 3 And a.No = Newno_In And a.���� = 1 And
          a.��¼״̬ = 1 And Nvl(a.����, 0) = 0 And
          ((Nvl(b.ҩ�����, 0) = 1 And
          a.�ⷿid Not In (Select ����id From ��������˵�� Where (�������� Like '%ҩ��') Or (�������� Like '�Ƽ���'))) Or Nvl(b.ҩ������, 0) = 1) And
          Rownum = 1;
  Exception
    When Others Then
      v_Druginf := '';
  End;

  If v_Druginf Is Not Null Then
    Raise Err_Isbatch;
  End If;

  --ԭ�����ֲ�������ҩƷ,�����ʱ��Ҫ������
  Update ҩƷ�շ���¼
  Set ���� = 0
  Where ID In
        (Select ID
         From ҩƷ�շ���¼ A, ҩƷ��� B
         Where b.ҩƷid = a.ҩƷid And a.No = Newno_In And a.���� = 1 And a.��¼״̬ = 1 And Nvl(a.����, 0) > 0 And
               (Nvl(b.ҩ�����, 0) = 0 Or
               (Nvl(b.ҩ������, 0) = 0 And
               a.�ⷿid In (Select ����id From ��������˵�� Where (�������� Like '%ҩ��') Or (�������� Like '�Ƽ���')))));

  For v_ҩƷ�շ���¼ In c_ҩƷ�շ���¼ Loop
    --����ɹ��ƻ����е�ִ����������ε�������ۼ�ִ������
    If v_ҩƷ�շ���¼.�ƻ�id > 0 Then
      Update ҩƷ�ƻ�����
      Set ִ������ = Nvl(ִ������, 0) + v_ҩƷ�շ���¼.ʵ������
      Where �ƻ�id = v_ҩƷ�շ���¼.�ƻ�id And ҩƷid = v_ҩƷ�շ���¼.ҩƷid;
    End If;
    --���ÿ����¼�¼���¿���
    Zl_ҩƷ���_Update(v_ҩƷ�շ���¼.Id, 0);
  
    --����Ƿ���ҩƷ�˿⣬ȡԭ���ĳɱ���
    If v_ҩƷ�շ���¼.�˿� = 1 Then
      Begin
        Select ƽ���ɱ���
        Into n_ԭ�ɱ���
        From ҩƷ���
        Where ���� = 1 And �ⷿid = v_ҩƷ�շ���¼.�ⷿid And ҩƷid = v_ҩƷ�շ���¼.ҩƷid And Nvl(����, 0) = Nvl(v_ҩƷ�շ���¼.����, 0);
      Exception
        When Others Then
          n_ԭ�ɱ��� := 0;
      End;
    End If;
  
    If v_ҩƷ�շ���¼.ʱ�� = 1 Then
      Update ҩƷ��� Set �ϴ��ۼ� = v_ҩƷ�շ���¼.���ۼ� Where ҩƷid = v_ҩƷ�շ���¼.ҩƷid;
    End If;
  
    If v_ҩƷ�շ���¼.�˿� = 0 Then
      --���¸�ҩƷ�ĳɱ���
      Update ҩƷ���
      Set �ɱ��� = v_ҩƷ�շ���¼.�ɱ���, �ϴι�Ӧ��id = v_ҩƷ�շ���¼.��ҩ��λid, �ϴ����� = v_ҩƷ�շ���¼.����, �ϴ��������� = v_ҩƷ�շ���¼.��������, �ϴβ��� = v_ҩƷ�շ���¼.����,
          �ϴ���׼�ĺ� = v_ҩƷ�շ���¼.��׼�ĺ�
      Where ҩƷid = v_ҩƷ�շ���¼.ҩƷid;
    End If;
  
    --����Ƿ���ҩƷ�˿⣬����ɱ����Ƿ�䶯������䶯���������۵�����¼�����������
    If Oldno_In Is Null Then
      If v_ҩƷ�շ���¼.�˿� = 1 Then
        If n_ԭ�ɱ��� <> 0 And n_ԭ�ɱ��� <> v_ҩƷ�շ���¼.�ɱ��� Then
          If v_Newno Is Null Then
            v_Newno := Nextno(25, v_ҩƷ�շ���¼.�ⷿid);
          End If;
        
          --��������۵�����
          Select ҩƷ�շ���¼_Id.Nextval Into n_�շ�id From Dual;
        
          n_������ := (v_ҩƷ�շ���¼.���۽�� - v_ҩƷ�շ���¼.���) - Round(n_ԭ�ɱ��� * v_ҩƷ�շ���¼.ʵ������, 2);
        
          Insert Into ҩƷ�շ���¼
            (ID, ��¼״̬, ����, NO, ���, �ⷿid, ������id, ��ҩ��λid, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ��д����, ���ۼ�, �ɱ���, ���, ժҪ, ������, ��������,
             �����, �������, ��������, ��׼�ĺ�, ����, ��ҩ��ʽ, ����, ���Ч��, ����id)
          Values
            (n_�շ�id, 1, 5, v_Newno, n_New���, v_ҩƷ�շ���¼.�ⷿid, n_������id, v_ҩƷ�շ���¼.��ҩ��λid, n_���ϵ��, v_ҩƷ�շ���¼.ҩƷid,
             v_ҩƷ�շ���¼.����, v_ҩƷ�շ���¼.����, v_ҩƷ�շ���¼.����, v_ҩƷ�շ���¼.Ч��, v_ҩƷ�շ���¼.ʵ������, v_ҩƷ�շ���¼.���۽��, v_ҩƷ�շ���¼.���, n_������,
             '�⹺�˿�������Զ�����', Nvl(�����_In, Zl_Username), �������_In, Nvl(�����_In, Zl_Username), �������_In, v_ҩƷ�շ���¼.��������,
             v_ҩƷ�շ���¼.��׼�ĺ�, v_ҩƷ�շ���¼.�ɱ���, 0, n_ԭ�ɱ���, v_ҩƷ�շ���¼.���Ч��, v_ҩƷ�շ���¼.Id);
        
          n_New��� := n_New��� + 1;
          --���¿��
          Zl_ҩƷ���_Update(n_�շ�id);
        End If;
      End If;
    Else
      Select Distinct �ɱ���, ���ۼ�
      Into n_�����ɱ���, n_�����ۼ�
      From ҩƷ�շ���¼
      Where NO = Oldno_In And ҩƷid = v_ҩƷ�շ���¼.ҩƷid And Nvl(����, 0) = Nvl(v_ҩƷ�շ���¼.����, 0) And ��� = v_ҩƷ�շ���¼.��� And ���� = 1 And
            Mod(��¼״̬, 3) = 2;
      If n_�����ɱ��� <> v_ҩƷ�շ���¼.�ɱ��� Then
        --��������۵�����
        v_Newno := Nextno(25, v_ҩƷ�շ���¼.�ⷿid);
        Select ҩƷ�շ���¼_Id.Nextval Into n_�շ�id From Dual;
      
        n_������ := Round((v_ҩƷ�շ���¼.�ɱ��� - n_�����ɱ���) * v_ҩƷ�շ���¼.ʵ������, 2);
      
        Insert Into ҩƷ�շ���¼
          (ID, ��¼״̬, ����, NO, ���, �ⷿid, ������id, ��ҩ��λid, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ��д����, ���ۼ�, �ɱ���, ���, ժҪ, ������, ��������,
           �����, �������, ��������, ��׼�ĺ�, ����, ��ҩ��ʽ, ����, ���Ч��, ����id)
        Values
          (n_�շ�id, 1, 5, v_Newno, n_New���, v_ҩƷ�շ���¼.�ⷿid, n_������id, v_ҩƷ�շ���¼.��ҩ��λid, n_���ϵ��, v_ҩƷ�շ���¼.ҩƷid, v_ҩƷ�շ���¼.����,
           v_ҩƷ�շ���¼.����, v_ҩƷ�շ���¼.����, v_ҩƷ�շ���¼.Ч��, v_ҩƷ�շ���¼.ʵ������, v_ҩƷ�շ���¼.���۽��, v_ҩƷ�շ���¼.���, n_������, '������˼۸�䶯����',
           Nvl(�����_In, Zl_Username), �������_In, Nvl(�����_In, Zl_Username), �������_In, v_ҩƷ�շ���¼.��������, v_ҩƷ�շ���¼.��׼�ĺ�,
           v_ҩƷ�շ���¼.�ɱ���, 0, n_�����ɱ���, v_ҩƷ�շ���¼.���Ч��, v_ҩƷ�շ���¼.Id);
      
        n_New��� := n_New��� + 1;
      
        --���¿��
        Zl_ҩƷ���_Update(n_�շ�id);
      End If;
    
      --�����ۼ�
      If n_�����ۼ� <> v_ҩƷ�շ���¼.���ۼ� Then
        Select ҩƷ�շ���¼_Id.Nextval Into n_�շ�id From Dual;
        Select Nextno(147) Into v_Billno From Dual;
      
        n_������ := Round((n_�����ۼ� - v_ҩƷ�շ���¼.���ۼ�) * v_ҩƷ�շ���¼.ʵ������, 2);
      
        --��������������¼
        Insert Into ҩƷ�շ���¼
          (ID, ��¼״̬, ����, NO, ���, ������id, ҩƷid, ����, ����, Ч��, ����, ����, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ����, ���۽��, ���, ժҪ, ������,
           ��������, �ⷿid, ���ϵ��, �����, �������, ����id)
        Values
          (n_�շ�id, 1, 13, v_Billno, n_New���, n_�ۼ�������id, v_ҩƷ�շ���¼.ҩƷid, v_ҩƷ�շ���¼.����, v_ҩƷ�շ���¼.����, v_ҩƷ�շ���¼.Ч��,
           v_ҩƷ�շ���¼.����, 1, v_ҩƷ�շ���¼.ʵ������, 0, n_�����ۼ�, 0, v_ҩƷ�շ���¼.���ۼ�, 0, n_������, n_������, '������˼۸�䶯����',
           Nvl(�����_In, Zl_Username), �������_In, v_ҩƷ�շ���¼.�ⷿid, 1, Nvl(�����_In, Zl_Username), �������_In, v_ҩƷ�շ���¼.Id);
      
        n_New��� := n_New��� + 1;
        --����ҩƷ���
        Zl_ҩƷ���_Update(n_�շ�id);
      End If;
    End If;
  End Loop;

  --��Ӧ��������д���
  --�˴���һ���飬��Ҫ�ǽ��û�ж�Ӧ��Ʊ�ŵļ�¼
  Begin
    Update Ӧ����¼
    Set ����� = �����_In, ������� = �������_In
    Where ��ⵥ�ݺ� = Newno_In And ϵͳ��ʶ = 1 And ��¼���� = 0 And ��¼״̬ = 1;
  
    Select b.��λid, Sum(��Ʊ���)
    Into v_��ҩ��λid, v_��Ʊ���
    From ҩƷ�շ���¼ A, Ӧ����¼ B
    Where a.Id = b.�շ�id And a.No = Newno_In And a.���� = 1 And b.ϵͳ��ʶ = 1
    Group By b.��λid;
  
    If Nvl(v_��ҩ��λid, 0) <> 0 Then
      Update Ӧ����� Set ��� = Nvl(���, 0) + Nvl(v_��Ʊ���, 0) Where ��λid = v_��ҩ��λid And ���� = 1;
    
      If Sql%NotFound Then
        Insert Into Ӧ����� (��λid, ����, ���) Values (v_��ҩ��λid, 1, v_��Ʊ���);
      End If;
    End If;
  Exception
    When No_Data_Found Then
      Null;
  End;
Exception
  When Err_Isverified Then
    Raise_Application_Error(-20101, '[ZLSOFT]�õ����Ѿ���������ˣ�[ZLSOFT]');
  When Err_Isbatch Then
    Raise_Application_Error(-20102, '[ZLSOFT]�õ����а�����һ��ԭ�������������ڷ�����ҩƷ[' || v_Druginf || ']��������ˣ�[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�⹺_Verify;
/

--128236,��ΰ��,2018-07-11,�״����ɺϲ�·����ƥ��
CREATE OR REPLACE Procedure Zl_����·������_Insert
( 
  ����_In            Number, --1=����,2=�޸� 
  ·����¼id_In      �����ٴ�·��.Id%Type, 
  �׶�id_In          �ٴ�·���׶�.Id%Type, 
  ����_In            ����·������.����%Type, 
  ����_In            ����·������.����%Type, 
  ������_In          ����·������.������%Type, 
  �������_In        ����·������.�������%Type, --0=����,1=�������,2=������˳�,3=�������������������ý���·���Ĺ��̣� 
  ����˵��_In        ����·������.����˵��%Type, 
  �Ǽ���_In          ����·������.�Ǽ���%Type, 
  ���������_In      ����·������.���������%Type, 
  ����ԭ��_In        Varchar2, --���ڶ������ԭ��ʱ������ԭ��1,����ԭ��2 
  ʱ�����_In        ����·������.ʱ�����%Type, --0=������1=��һ�׶���ǰ�����죬2=��һ�׶���ǰ�����죬-1=�Ӻ� 
  ��·��id_In        �����ٴ�·��.·��id%Type, 
  ��·���汾_In      �����ٴ�·��.�汾��%Type, 
  ָ������_In        Varchar2, --ָ������|ָ����|ָ������||...,ĩβ��||,����Ϊ�� 
  ���_In            Number, 
  ��ת�����_In      ����·������.��ת�����%Type := Null, 
  �����ʷ��ת_In    Number := 0, 
  �����ϲ�·��ids_In Varchar2 := Null, --�������������ĺϲ�·����¼IDs 
  ����ʱ������_In    ����·��ִ��.����ʱ������%Type := Null --������_In =2,����ʱ������_IN=1-ʱ��ֻ�޸��������������ԭ�򣨴��ڶ������ԭ�������� 
) Is 
  Cursor c_Merge(·����¼id_In ����·��ִ��.·����¼id%Type) Is 
    Select a.Id, a.��ǰ�׶�id, a.��ǰ���� 
    From ���˺ϲ�·�� A 
    Where a.��Ҫ·����¼id = ·����¼id_In And a.����ʱ�� Is Null And a.��ǰ�׶�id is not Null; 
  t_�ϲ�·����¼id t_Numlist; 
  t_�ϲ�·���׶�id t_Numlist; 
  t_�ϲ�·������   t_Numlist; 
 
  v_Str   Varchar2(4000); 
  v_Tmp   Varchar2(1000); 
  n_Index Number; 
  I       Number(5) := 1; 
 
  l_ָ������ t_Strlist := t_Strlist(); 
  l_ָ���� t_Strlist := t_Strlist(); 
  l_ָ������ t_Numlist := t_Numlist(); 
 
  v_ԭ·��id     �����ٴ�·��.·��id%Type; 
  v_ԭ·���汾   �����ٴ�·��.�汾��%Type; 
  d_��ת���ʱ�� ����·������.��ת�����%Type; 
  d_�Ǽ�ʱ��     ����·������.�Ǽ�ʱ��%Type; 
  n_��ǰ�׶�id   ���˺ϲ�·��.��ǰ�׶�id%Type; 
  d_Date         Date; 
  n_����id       �����ٴ�·��.����id%Type; 
  n_��ҳid       �����ٴ�·��.��ҳid%Type; 
  n_Count        Number(5); 
  v_Error        Varchar2(255); 
  Err_Custom Exception; 
 
  Procedure p_�ݴ���Ŀ_Delete 
  ( 
    ·����¼id_In Number, 
    �׶�id_In     Number, 
    ����_In       Date 
  ) Is 
    n_Count Number(5); 
  Begin 
    --������˳����������Ҫɾ���ݴ�·������Ŀ��ȡ��ҽ������ 
    Select Count(1) 
    Into n_Count 
    From ����·��ִ�� T 
    Where t.·����¼id = ·����¼id_In And t.�׶�id = �׶�id_In And t.���� = ����_In And t.��Ŀid Is Null And t.����ʱ������ = 2; 
    If n_Count > 0 Then 
      --ȡ��ҽ������ 
      Delete From ����·��ҽ�� 
      Where ·��ִ��id In (Select a.Id 
                       From ����·��ִ�� A, ����·��ҽ�� B 
                       Where a.Id = b.·��ִ��id And a.·����¼id = ·����¼id_In And a.�׶�id = �׶�id_In And a.���� = ����_In And 
                             a.��Ŀid Is Null And a.����ʱ������ = 2); 
      --ɾ������Ŀ 
      Delete From ����·��ִ�� 
      Where ·����¼id = ·����¼id_In And �׶�id = �׶�id_In And ���� = ����_In And ��Ŀid Is Null And ����ʱ������ = 2; 
    End If; 
  End p_�ݴ���Ŀ_Delete; 
 
Begin 
  Select Sysdate Into d_Date From Dual; 
  If ��ת�����_In Is Not Null Then 
    d_��ת���ʱ�� := d_Date; 
  End If; 
  If ���_In = 1 Then 
    If ����_In = 1 Then 
      If Nvl(��·��id_In, 0) <> 0 Then 
        Select ·��id, �汾�� Into v_ԭ·��id, v_ԭ·���汾 From �����ٴ�·�� Where ID = ·����¼id_In; 
        Update �����ٴ�·�� Set ·��id = ��·��id_In, �汾�� = ��·���汾_In Where ID = ·����¼id_In; 
      End If; 
 
      If �������_In = 2 Or �������_In = 3 Then 
        p_�ݴ���Ŀ_Delete(·����¼id_In, �׶�id_In, ����_In); 
      End If; 
 
      Insert Into ����·������ 
        (·����¼id, �׶�id, ����, ����, ������, ����ʱ��, �������, ����˵��, �Ǽ���, �Ǽ�ʱ��, ����ԭ��, ʱ�����, ���������, �������ʱ��, ԭ·��id, ԭ·���汾, ��ת�����, ��ת���ʱ��) 
      Values 
        (·����¼id_In, �׶�id_In, ����_In, ����_In, ������_In, d_Date, Decode(�������_In, 0, 1, -1), ����˵��_In, �Ǽ���_In, d_Date, Null, 
         ʱ�����_In, ���������_In, d_Date, v_ԭ·��id, v_ԭ·���汾, ��ת�����_In, d_��ת���ʱ��); 
 
      If ����ԭ��_In Is Not Null Then 
        n_Index := 0; 
        For r_����ԭ�� In (Select Column_Value As ����ԭ�� From Table(f_Str2list(����ԭ��_In))) Loop 
          If n_Index = 0 Then 
            --����һ������ԭ�򵽲���·��������������ǰ 
            Update ����·������ T 
            Set t.����ԭ�� = r_����ԭ��.����ԭ�� 
            Where t.·����¼id = ·����¼id_In And t.�׶�id = �׶�id_In And t.���� = ����_In; 
            n_Index := 1; 
          End If; 
          Insert Into ����·������ 
            (·����¼id, �׶�id, ����, ����ԭ��) 
          Values 
            (·����¼id_In, �׶�id_In, ����_In, r_����ԭ��.����ԭ��); 
        End Loop; 
      End If; 
 
      --�洢�ϲ�·������ 
      Open c_Merge(·����¼id_In); 
      Fetch c_Merge Bulk Collect 
        Into t_�ϲ�·����¼id, t_�ϲ�·���׶�id, t_�ϲ�·������; 
      Close c_Merge; 
      If t_�ϲ�·����¼id.Count > 0 Then 
        Forall I In 1 .. t_�ϲ�·����¼id.Count 
          Insert Into ���˺ϲ�·������ 
            (·����¼id, �׶�id, ����, �ϲ�·����¼id, �ϲ�·���׶�id, �ϲ�·������, �Ǽ�ʱ��) 
          Values 
            (·����¼id_In, �׶�id_In, ����_In, t_�ϲ�·����¼id(I), t_�ϲ�·���׶�id(I), t_�ϲ�·������(I), d_Date); 
      End If; 
    Elsif ����_In = 2 Then 
      --����=2 
      If Nvl(����ʱ������_In, 0) <> 1 Then 
        Select �Ǽ�ʱ�� 
        Into d_�Ǽ�ʱ�� 
        From ����·������ 
        Where ·����¼id = ·����¼id_In And �׶�id = �׶�id_In And ���� = ����_In; 
        Update ����·������ 
        Set ������ = ������_In, ����ʱ�� = d_Date, ������� = Decode(�������_In, 0, 1, -1), ����˵�� = ����˵��_In, �Ǽ��� = �Ǽ���_In, �Ǽ�ʱ�� = d_Date, 
            ʱ����� = ʱ�����_In, ��������� = ���������_In, �������ʱ�� = d_Date, ��ת����� = ��ת�����_In, ��ת���ʱ�� = d_��ת���ʱ�� 
        Where ·����¼id = ·����¼id_In And �׶�id = �׶�id_In And ���� = ����_In; 
 
        If �������_In = 2 Or �������_In = 3 Then 
          p_�ݴ���Ŀ_Delete(·����¼id_In, �׶�id_In, ����_In); 
        End If; 
 
        --ɾ�����ٲ��루���ڶ������ԭ�� 
        If ����ԭ��_In Is Not Null Then 
          Delete From ����·������ Where ·����¼id = ·����¼id_In And �׶�id = �׶�id_In And ���� = ����_In; 
          n_Index := 0; 
          For r_����ԭ�� In (Select Column_Value As ����ԭ�� From Table(f_Str2list(����ԭ��_In))) Loop 
            If n_Index = 0 Then 
              Update ����·������ 
              Set ����ԭ�� = r_����ԭ��.����ԭ�� 
              Where ·����¼id = ·����¼id_In And �׶�id = �׶�id_In And ���� = ����_In; 
              n_Index := 1; 
            End If; 
            Insert Into ����·������ 
              (·����¼id, �׶�id, ����, ����ԭ��) 
            Values 
              (·����¼id_In, �׶�id_In, ����_In, r_����ԭ��.����ԭ��); 
          End Loop; 
        End If; 
        --��¼��·������Ŀ�Ѿ��������ʱ�����µ����������������ԭ�� 
        --1.��¼ǰ�������Ϊ1��������ʱ������Ϊ-1 ���������) 
        --2.����ԭ��ɾ��������¼�루ԭ�򣺱��ⲡ��·���������ظ�������ͬ����ԭ��ֵ�� 
      Elsif Nvl(����ʱ������_In, 0) = 1 Then 
        Delete From ����·������ Where ·����¼id = ·����¼id_In And �׶�id = �׶�id_In And ���� = ����_In; 
        n_Index := 0; 
        For r_�±���ԭ�� In (Select Distinct ����ԭ�� 
                        From ����·��ִ�� 
                        Where ·����¼id = ·����¼id_In And �׶�id = �׶�id_In And ���� = ����_In And ����ԭ�� Is Not Null And 
                              Nvl(����ʱ������, 0) < 2) Loop 
          If n_Index = 0 Then 
            Update ����·������ 
            Set ����ԭ�� = r_�±���ԭ��.����ԭ��, ������� = Decode(�������, 1, -1) 
            Where ·����¼id = ·����¼id_In And �׶�id = �׶�id_In And ���� = ����_In; 
            n_Index := 1; 
          End If; 
          Insert Into ����·������ 
            (·����¼id, �׶�id, ����, ����ԭ��) 
          Values 
            (·����¼id_In, �׶�id_In, ����_In, r_�±���ԭ��.����ԭ��); 
        End Loop; 
 
      End If; 
    End If; 
    If �����ʷ��ת_In = 1 Then 
      Update ����·������ 
      Set ��ת����� = ��ת�����_In, ��ת���ʱ�� = d_��ת���ʱ�� 
      Where ·����¼id = ·����¼id_In And ԭ·��id Is Not Null And ��ת����� Is Null; 
    End If; 
  End If; 
 
  If Not ָ������_In Is Null Then 
    v_Str := ָ������_In; 
    Loop 
      n_Index := Instr(v_Str, '||'); 
      Exit When(Nvl(n_Index, 0) = 0); 
      l_ָ������.Extend; 
      l_ָ����.Extend; 
      l_ָ������.Extend; 
 
      v_Tmp := Substr(v_Str, 1, n_Index - 1); 
      l_ָ������(I) := Substr(v_Tmp, 1, Instr(v_Tmp, '|') - 1); 
      v_Tmp := Substr(v_Tmp, Instr(v_Tmp, '|') + 1); 
      l_ָ����(I) := Trim(Substr(v_Tmp, 1, Instr(v_Tmp, '|') - 1)); 
      v_Tmp := Substr(v_Tmp, Instr(v_Tmp, '|') + 1); 
      l_ָ������(I) := To_Number(v_Tmp); 
 
      v_Str := Substr(v_Str, n_Index + 2); 
      I     := I + 1; 
    End Loop; 
 
    If ����_In = 1 Then 
      Forall I In 1 .. l_ָ������.Count 
 
        Insert Into ����·��ָ�� 
          (·����¼id, �׶�id, ����, ����, ��������, ����ָ��, ָ����, ָ������) 
        Values 
          (·����¼id_In, �׶�id_In, ����_In, ����_In, 2, l_ָ������(I), l_ָ����(I), l_ָ������(I)); 
    Else 
      Forall I In 1 .. l_ָ������.Count 
        Update ����·��ָ�� 
        Set ָ���� = l_ָ����(I) 
        Where ·����¼id = ·����¼id_In And �׶�id = �׶�id_In And ���� = ����_In And ����ָ�� = l_ָ������(I); 
    End If; 
  End If; 
 
  If ���_In = 1 And �������_In = 2 Then 
    If ����_In = 2 Then 
      n_Index := 0; 
      Select ��ǰ�׶�id Into n_Index From �����ٴ�·�� Where ID = ·����¼id_In; 
      If n_Index <> �׶�id_In Then 
        v_Error := '�ò����������˴��յ�·����Ŀ,�����޸��������������·����'; 
        Raise Err_Custom; 
      End If; 
    End If; 
    --��ǰ����,�����,����ͳ�Ʒ��� 
    Update �����ٴ�·�� 
    Set ����ʱ�� = d_Date, ״̬ = 3, ǰһ�׶�id = �׶�id_In, ��ǰ�׶�id = Null 
    Where ID = ·����¼id_In 
    Returning ����id, ��ҳid Into n_����id, n_��ҳid; 
 
    --���²�����ҳ��ǰ·����״̬ 
    Update ������ҳ Set ·��״̬ = 3 Where ����id = n_����id And ��ҳid = n_��ҳid; 
 
    --�����ϲ�·�� 
    Update ���˺ϲ�·�� 
    Set ����ʱ�� = d_Date, ǰһ�׶�id = ��ǰ�׶�id, ��ǰ�׶�id = Null 
    Where ��Ҫ·����¼id = ·����¼id_In And ����ʱ�� Is Null; 
  Elsif ���_In = 1 Then 
    --��Ҫ·���޸ĳ���������ȡ�������ϲ�·�� 
    If ����_In = 2 and  Nvl(����ʱ������_In, 0) <> 1Then 
      --���ȡ�������Ǽ�ʱ������ĺϲ�·�� 
      For c_Merge In (Select ID 
                      From ���˺ϲ�·�� 
                      Where ����ʱ�� = d_�Ǽ�ʱ�� And ��Ҫ·����¼id = ·����¼id_In And ����ʱ�� Is Not Null) Loop 
        Select b.�ϲ�·���׶�id 
        Into n_��ǰ�׶�id 
        From ���˺ϲ�·������ B 
        Where b.�Ǽ�ʱ�� = (Select Max(c.�Ǽ�ʱ��) 
                        From ���˺ϲ�·������ C 
                        Where c.·����¼id = b.·����¼id And c.�ϲ�·����¼id = b.�ϲ�·����¼id) And b.�ϲ�·����¼id = c_Merge.Id And 
              b.·����¼id = ·����¼id_In; 
 
        Update ���˺ϲ�·�� 
        Set ����ʱ�� = Null, ǰһ�׶�id = n_��ǰ�׶�id, ��ǰ�׶�id = ǰһ�׶�id 
        Where ����ʱ�� = d_�Ǽ�ʱ�� And ��Ҫ·����¼id = ·����¼id_In And ����ʱ�� Is Not Null; 
      End Loop; 
    End If; 
  End If; 
  If �����ϲ�·��ids_In Is Not Null Then 
    Update /*+ Rule */ ���˺ϲ�·�� 
    Set ����ʱ�� = d_Date, ǰһ�׶�id = ��ǰ�׶�id, ��ǰ�׶�id = Null 
    Where ID In (Select * From Table(Cast(f_Num2list(�����ϲ�·��ids_In) As Zltools.t_Numlist))); 
 
  End If; 
Exception 
  When Err_Custom Then 
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]'); 
  When Others Then 
    zl_ErrorCenter(SQLCode, SQLErrM); 
End Zl_����·������_Insert;
/
---------------------------------------------------------------------------------------------------
--����ϵͳ�������İ汾��
-------------------------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.34.160' Where ���=&n_System;
--�����汾��
Commit;
