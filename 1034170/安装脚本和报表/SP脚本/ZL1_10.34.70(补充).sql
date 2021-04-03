--[��������]1
--[�����߰汾��]10.34.30
--���ű�֧�ִ�ZLHIS+ v10.34.60 ������ v10.34.70
--����ϵͳ�����ߵ�¼PLSQL��ִ�����нű�
--�ű�ִ�к����ֹ�������������
Define n_System=100;
-------------------------------------------------------------------------------
--�ṹ��������
-------------------------------------------------------------------------------
--90040:������,2016-03-17,�������֤δ¼ԭ��������� ���֤��״̬ �������Ϊ ���֤δ¼ԭ��   
Declare
  n_Tableexist Number := 0;
Begin
  Begin
    Execute Immediate 'select count(*)  from USER_TABLES T WHERE T.TABLE_NAME = ''���֤��״̬'''
      Into n_Tableexist;
  
    If n_Tableexist = 1 Then
      Execute Immediate 'ALTER TABLE ���֤��״̬ RENAME TO ���֤δ¼ԭ��';
      Execute Immediate 'alter table ���֤δ¼ԭ�� drop Constraint ���֤��״̬_PK';
      Execute Immediate 'alter table ���֤δ¼ԭ�� drop Constraint ���֤��״̬_UQ_����';
    Else
      Execute Immediate 'Create Table ���֤δ¼ԭ��(
                                                    ���� VARCHAR2(2),
                                                    ���� VARCHAR2(50),
                                                    ���� VARCHAR2(10),
                                                  ȱʡ��־ NUMBER(1) default 0,
                                                    ˵�� VARCHAR2(50))
                                                    TABLESPACE zl9BaseItem';
    End If;
  Exception
    When Others Then
      Null;
  End;
End;
/
Alter Table ���֤δ¼ԭ�� Add Constraint ���֤δ¼ԭ��_PK Primary Key (����) Using Index Tablespace zl9Indexcis;
Alter Table ���֤δ¼ԭ�� Add Constraint ���֤δ¼ԭ��_UQ_���� Unique (����) Using Index Tablespace zl9Indexhis;

--92729:������,2016-03-08,רҵ��RIS�ӿڴ���
Create Table RIS���ԤԼ (
ҽ��ID  NUMBER(18),
ԤԼID   NUMBER(18),
ԤԼ����  DATE,
����豸ID  NUMBER(18),
����豸����  VARCHAR2(64),
ԤԼ��ʼʱ��  DATE,
ԤԼ����ʱ��  DATE,
ԤԼ��ʼʱ���  DATE,
ԤԼ����ʱ���  DATE,
��ת��  NUMBER(3))     
TABLESPACE zl9CisRec;

Alter Table RIS���ԤԼ Add Constraint RIS���ԤԼ_PK Primary Key(ҽ��ID) Using Index Tablespace zl9Indexhis;
Alter Table RIS���ԤԼ Add Constraint RIS���ԤԼ_FK_ҽ��ID Foreign Key (ҽ��ID) References ����ҽ����¼(ID) On Delete Cascade;
Create Index RIS���ԤԼ_IX_��ת�� On RIS���ԤԼ(��ת��) Tablespace zl9Indexcis;

--93380:������,2016-02-24,��������鳤����
Create Table �������鳤����(
    ��ID Number(18),
    �鳤ID Number(18),
    �ϴ�����ʱ�� Date)
    TABLESPACE zl9Expense
    PCTFREE 5;

Alter Table �������鳤���� Add Constraint �������鳤����_PK Primary Key (��ID,�鳤ID) Using Index Tablespace ZL9INDEXHIS;
Create Index �������鳤����_IX_�鳤ID On �������鳤����(�鳤ID) Pctfree 5 Tablespace zl9indexhis;

--91225:������,2016-02-16,�������Լ�¼ �����ֶ� ҽ��ID
alter table  �������Լ�¼ add (ҽ��ID number(18));
Create Index �������Լ�¼_IX_ҽ��ID On �������Լ�¼(ҽ��ID) Tablespace zl9Indexcis;
Alter Table �������Լ�¼ Add Constraint �������Լ�¼_FK_ҽ��ID Foreign Key (ҽ��ID) References ����ҽ����¼(ID);

--91316:�ŵ���,2016-01-21,��Һ������������
alter table ��ҩ�������� add ҩƷ���� varchar2(20);

--91333:�ŵ���,2016-01-21,��Һ�������Ĵ�ӡ��ˮ��
alter table ��Һ��ҩ��¼ modify ��ӡ��־ number(5);

--92808:����,2016-01-22,���ĵ��ݴ�ӡ����
Create table ҩƷ�շ����� 
(
ID number(18),
no varchar2(8),
���� number(2),
�ⷿid number(18),
��ӡ״̬ number(1)
) tablespace ZL9MEDLST;

Create Sequence ҩƷ�շ�����_ID Start With 1; 
Alter Table ҩƷ�շ����� Add Constraint ҩƷ�շ�����_PK Primary Key (ID) Using Index Tablespace zl9indexhis;
Alter Table ҩƷ�շ����� Add Constraint ҩƷ�շ�����_UQ_NO Unique (no,����,�ⷿid) Using Index Tablespace zl9indexhis;

--93380:������,2016-02-24,��������鳤����
Alter Table �������鳤���� Add Constraint �������鳤����_FK_��ID Foreign Key (��ID) References ����ɿ����(Id);
Alter Table �������鳤���� Add Constraint �������鳤����_FK_�鳤ID Foreign Key (�鳤ID) References ��Ա��(Id);

--91225:������,2016-3-7,�޸� �������Լ�¼����Ⱦ����¼��������ֶγ���
alter table ��Ⱦ��Ŀ¼ modify  ����  VARCHAR2(20);
alter table �������Լ�¼ modify  �걾���� varchar2(64);
alter table �������Լ�¼ modify  �ͼ�ҽ�� VARCHAR2(100);
alter table ��Ⱦ��Ŀ¼ modify  (���� varchar2(150),���� VARCHAR2(20),˵�� VARCHAR2(200));

-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--90040:������,2016-03-10,�������֤δ¼ԭ���ֵ���Ĭ������
Insert Into ���֤δ¼ԭ��
  (����, ����, ����)
  Select '01', 'δ��', 'WD' From Dual Where Not Exists (Select 1 From ���֤δ¼ԭ�� Where ���� = '01');
Insert Into ���֤δ¼ԭ��
  (����, ����, ����)
  Select '02', '��ʧ����', 'YSDB' From Dual Where Not Exists (Select 1 From ���֤δ¼ԭ�� Where ���� = '02');
Insert Into ���֤δ¼ԭ��
  (����, ����, ����)
  Select '03', 'δ��', 'WB' From Dual Where Not Exists (Select 1 From ���֤δ¼ԭ�� Where ���� = '03');


--90040:������,2016-03-17,���zlBaseCode�������Ѿ��������֤��״̬������ݾ�ֱ���޸�
update zlBaseCode set ���� = '���֤δ¼ԭ��' where ϵͳ = &n_sysTem and ���� = '���֤��״̬';
--90040:������,2016-03-10,�������֤δ¼ԭ���ֵ��
Insert Into zlBaseCode
  (ϵͳ, ����, �̶�, ˵��, ����)
  Select &n_System, '���֤δ¼ԭ��', 0, '�������֤����δ¼���ԭ��', '��Ա����'
  From Dual
  Where Not Exists (Select 1 From zlBaseCode Where ϵͳ =&n_System and  ���� = '���֤δ¼ԭ��');


--92729:������,2016-03-08,רҵ��RIS�ӿڴ���
Insert Into zlBakTables(ϵͳ,���,����,���,ֱ��ת��,ͣ�ô�����)
Select &n_System,8,A.* From (
Select ����,���,ֱ��ת��,ͣ�ô����� From zlBakTables Where 1 = 0 Union All
Select 'RIS���ԤԼ',21,1,-NULL From Dual Union All
Select ����,���,ֱ��ת��,ͣ�ô����� From zlBakTables Where 1 = 0) A;

--93917:����,2016-03-07,��RIS�����ݽ����ӿڽű�
Insert Into zlComponent(����,����,���汾,�ΰ汾,���汾,ϵͳ,ע���Ʒ����,ע���Ʒ����,ע���Ʒ�汾) Values('zl9XWInterface','Ӱ����Ϣϵͳ��������',10,35,10,&n_System,'����ҽԺ��Ϣϵͳ','ZLHIS+','10'); 
Insert Into zlPrograms(���,����,˵��,ϵͳ,����) Values( 1287,'ҽѧӰ����Ϣϵͳרҵ��','Ӱ��RIS��PACSϵͳ',&n_System,'zl9XWInterface'); 

Insert Into zlMenus(���,ID,�ϼ�ID,����,���,ϵͳ,ģ��,�̱���,ͼ��,˵��)
	Select ���,Zlmenus_Id.Nextval,ID,'ҽѧӰ����Ϣϵͳרҵ��' ,'B' ,&n_System,-NULL ,'ҽѧӰ����Ϣϵͳרҵ��' ,99 ,'Ӱ��RIS��PACSϵͳ' 
         From zlMenus Where ���� = 'ҽѧӰ��ϵͳ' And ��� = 'ȱʡ' And ϵͳ = &n_System And ģ�� Is Null;                 

Insert Into zlMenus(���,ID,�ϼ�ID,����,���,ϵͳ,ģ��,�̱���,ͼ��,˵��) 
	Select A.���,ZlMenus_ID.Nextval,A.ID,B.* From (
	     Select ���,ID From zlMenus Where ���� = 'ҽѧӰ����Ϣϵͳרҵ��' And ��� = 'ȱʡ' And ϵͳ = &n_System And ģ�� Is Null) A,
	    (Select ����,���,ϵͳ,ģ��,�̱���,ͼ��,˵�� From zlMenus Where 1 = 0 Union All
	Select 'ҽѧӰ����Ϣϵͳרҵ��' ,'R'  ,&n_System, 1287, 'Ӱ����Ϣ����վ' ,99, 'Ӱ��RIS��PACS����վ'  From Dual Union All						
                 Select ����,���,ϵͳ,ģ��,�̱���,ͼ��,˵�� From zlMenus Where 1 = 0) B; 

--78768:���ϴ�,2016-03-04,IDKind����ȱʡ�������
Insert Into Zlparameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1153, 1, 1, 0, 0, 1,'ȱʡ�������', Null, '0', '����ȱʡ�Ķ��������,�洢���ǿ����ID'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where ϵͳ = &n_System And ģ�� = 1153 And ������ = 'ȱʡ�������');

Insert Into Zlparameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1153, 1, 1, 0, 0, 2,'��ǰ����-���ܼ�', Null, Null, '����IDkind�Ĺ��ܼ�'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where ϵͳ = &n_System And ģ�� = 1153 And ������ = '��ǰ����-���ܼ�');

Insert Into Zlparameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1153, 1, 1, 0, 0, 3,'��ǰ����-���', Null, Null, '����IDkind�Ĺ��ܼ�'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where ϵͳ = &n_System And ģ�� = 1153 And ������ = '��ǰ����-���');

Insert Into Zlparameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1153, 1, 1, 0, 0, 4,'������-���ܼ�', Null, Null, '����IDkind�Ĺ��ܼ�'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where ϵͳ = &n_System And ģ�� = 1153 And ������ = '������-���ܼ�');

Insert Into Zlparameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1153, 1, 1, 0, 0, 5,'������-���', Null, Null, '����IDkind�Ĺ��ܼ�'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where ϵͳ = &n_System And ģ�� = 1153 And ������ = '������-���');

Insert Into Zlparameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1153, 1, 1, 0, 0, 6,'����-���ܼ�', Null, Null, '����IDkind�Ĺ��ܼ�'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where ϵͳ = &n_System And ģ�� = 1153 And ������ = '����-���ܼ�');

Insert Into Zlparameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1153, 1, 1, 0, 0, 7,'����-���', Null, Null, '����IDkind�Ĺ��ܼ�'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where ϵͳ = &n_System And ģ�� = 1153 And ������ = '����-���');

--91954:�ŵ���,2016-02-29,������ú�͸���ҽԺ���Զ���������
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1345, 0, 0, 0, 0, 35, '�Զ�����ʱ��Һ��������ֻ���������α䶯', '0', '0', '���ò�������֮���Զ��������Ὣ����������ŵ�ǰ�棬���統2#����ʱ�����Ὣ3���ķ��䵽2#'
  From Dual
  Where Not Exists (Select 1 From zlParameters Where ϵͳ = &n_System And ģ�� = 1345 And ������ = '�Զ�����ʱ��Һ��������ֻ���������α䶯');

--92718:����,2016-02-26,����Ӱ����Ϣϵͳ�ӿڿ���
Insert Into Zlparameters
  (Id, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, -null, -null, -null, -null, -null, 255, '����ҽѧӰ����Ϣϵͳרҵ��ӿ�', '0', '0',
         '�����Ƿ�����Ӱ����Ϣϵͳ�ӿڣ�0-δ����;1-���ã����밲װӰ����Ϣϵͳ�Ĳ���Ч��'
  From Dual
  Where Not Exists (Select 1
         From Zlparameters
         Where ������ = '����ҽѧӰ����Ϣϵͳרҵ��ӿ�' And Nvl(ģ��, 0) = 0 And Nvl(ϵͳ, 0) = &n_System);

--93311:�ŵ���,2016-02-24,����ҽ����ʱ�������û�ҩ��������
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1345, 0, 0, 0, 0, 34, '�������û�ҩ������Һ��������', '0', '0', '�˲��������󣬷���ҽ����ʱ�������û�ҩ������Һ��������'
  From Dual
  Where Not Exists
   (Select 1 From zlParameters Where Nvl(ϵͳ, 0) = &n_System And  Nvl(ģ��, 0) =1345 And ������ = '�������û�ҩ������Һ��������');

--92384:������,2016-02-24,ҽ����ִ����Ϣ
Insert Into ҵ����Ϣ����(����,����,˵��,��������)  Select 'ZLHIS_CIS_034','��ִ��ҽ������','ҽ�����ͺ�����Ҫִ�еǼǵ�ҽ����������һ����ִ��ͨ��Ϣ��',7 From Dual;

--91956:�ŵ���,2016-02-22,������ҩƷ�����յ�ҩƷҽ�������뾲������
Insert Into Zlparameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1345, 0, 0, 0, 0, 32,'�Ƿ����õĳ���ҩƷ����ҩƷ���˲���', Null, '1', '���������õĳ���ҩƷ������Һ����ʱ��ҩƷ�Ǹ��ݵ�����Һ���Ļ�������ȷ��'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where ϵͳ = &n_System And ģ�� = 1345 And ������ = '�Ƿ����õĳ���ҩƷ����ҩƷ���˲���');

--92537:������,2016-02-17,���Ӳ��� ��������������ɵ��� ����������������ɵ���
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, -null, 0, 0, 0, 0, 257, '��������������ɵ���', '0', '0', '�����Ƿ�������Ϻ����������ɵ������ܣ����ú�סԺ����������ѡ����������������������ƿ��������޸ġ�0-δ����;1-����'
  From Dual
  Where Not Exists
   (Select 1 From zlParameters Where Nvl(ϵͳ, 0) = &n_System And  Nvl(ģ��, 0) = 0 And ������ = '��������������ɵ���');


--93221:������,2016-02-02,ȱʡ����ʱ��
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1506, 1, 0, 0, 0, 2, 'ȱʡ����ʱ��', Null, Null,
         'ÿ�����ʵ�Ĭ��ʱ��,��ǰʱ�䳬���������õ�ʱ���,��ֹʱ��Ĭ��Ϊ�������õ�ʱ��;��ǰʱ��δ�����������õ�ʱ���,�Ե�ǰʱ��Ϊ׼'
  From Dual
  Where Not Exists (Select 1 From zlParameters Where ������ = 'ȱʡ����ʱ��' And ϵͳ = &n_System And ģ�� = 1506);


--92468:���ϴ�,2016-01-25,�Һŷ����ϸ���Ƶ�����£�����Ϊ���Ƿ��߷�Ʊ��
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1111, 0, 0, 0, 0, 67, '�㿨����Ʊ��', '0', '0', '�Һŷ����ϸ���Ƶ�����£�����Ϊ���Ƿ��߷�Ʊ��'
  From Dual
  Where Not Exists (Select 1 From zlParameters Where ������ = '�㿨����Ʊ��' And ϵͳ = &n_System And ģ�� = 1111);


--91584:�ŵ���,2016-01-22,��˾�������ҽ��
Insert Into Zlparameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1345, 0, 0, 0, 0, 30,'��˸�ҩ������������', Null, '0', '��˸�ҩ���Ĳ��ŷ�ҩ�;������ĵ�����'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where ϵͳ = &n_System And ģ�� = 1345 And ������ = '��˸�ҩ������������');

--91732:�ŵ���,2016-01-22,������ҩƷ�����յ�ҩƷҽ�������뾲������
Insert Into Zlparameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1345, 0, 0, 0, 0, 31,'������ҩƷ���������յ�ҩƷҽ�����ڲ��ŷ�ҩִ��', Null, '0', '������ҩƷ���������յ�ҩƷҽ�����ڲ��ŷ�ҩִ��'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where ϵͳ = &n_System And ģ�� = 1345 And ������ = '������ҩƷ���������յ�ҩƷҽ�����ڲ��ŷ�ҩִ��');

--91954:�ŵ���,2016-01-22,Zl_��Һ��ҩ��¼_�Զ�����
Insert Into Zlparameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1345, 0, 0, 0, 0, 33,'�����Զ�����', Null, '0', '�����Զ�����֮�󽫸����������ȼ��Զ�������Һ�������Σ�ͬʱ���ܱ����ϴ�����'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where ϵͳ = &n_System And ģ�� = 1345 And ������ = '�����Զ�����');

--92852:������,2016-01-21,��λ��Ƭ����ʽ����
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1265, 0, 0, 0, 0, 9, '��λ��Ƭ����ʽ', '', '1', '���ƴ�λ��Ƭ����ʾ˳��:�ǰ��Ҵ���������ʾ�����ǰ���λ���Ʒ���+��λ��������ʾ'
  From Dual
  Where Not Exists (Select 1 From zlParameters Where ������ = '��λ��Ƭ����ʽ' And ϵͳ = &n_System And ģ�� = 1265);

--91316:�ŵ���,2016-01-21,��Һ������������
Insert Into Zlparameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1345, 0, 0, 0, 0, 27,'����ҩƷ����������ҩƷ�����ݸ�ҩʱ��û����ҩ���ε���Һ��Ĭ��Ϊ0���β����', Null, '0', '����ҩƷ����������ҩƷ�����ݸ�ҩʱ��û����ҩ���ε���Һ��Ĭ��Ϊ0���β����'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where ϵͳ = &n_System And ģ�� = 1345 And ������ = '����ҩƷ����������ҩƷ�����ݸ�ҩʱ��û����ҩ���ε���Һ��Ĭ��Ϊ0���β����');

Insert Into Zlparameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1345, 0, 0, 0, 0, 28,'�����Σ�ҩƷ����', Null, '0', '�˲�����ѡ֮�󣬽����ϵ�����ҩƷ�����Σ�ҩƷ��������'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where ϵͳ = &n_System And ģ�� = 1345 And ������ = '�����Σ�ҩƷ����');

Insert Into Zlparameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1345, 0, 0, 0, 0, 29,'����ҩƷ��ҩƷ����ָ������', Null, '0', '����ҩ��Ӫ��ҩ������ҩ���ε�ҩƷ����ƥ������'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where ϵͳ = &n_System And ģ�� = 1345 And ������ = '����ҩƷ��ҩƷ����ָ������');

--92725:��ҵ��,2016-01-21,������ҩ����ӡ��ʾ��ʽ����
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1341, 0, 1, 0, 0, 55, '����ҩ���ݴ�ӡ��ʾ��ʽ', '0', '0', '0-��ʾ������ҩ��,1-ֻ��ʾδ��ӡ�Ĵ���ҩ����,2-ֻ��ʾ�Ѵ�ӡ�Ĵ���ҩ����'
  From Dual
  Where Not Exists (Select 1 From zlParameters Where ϵͳ = &n_System And ģ�� = 1341 And ������ = '����ҩ���ݴ�ӡ��ʾ��ʽ');

--92725:��ҵ��,2016-01-21,���Ӵ���ҩ����ɨ����Զ����в���
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1341, 0, 1, 0, 0, 56, '����ҩ����ɨ����Զ�����', '0', '0', '0-���Զ�����,1-ɨ����Զ���������'
  From Dual
  Where Not Exists (Select 1 From zlParameters Where ϵͳ = &n_System And ģ�� = 1341 And ������ = '����ҩ����ɨ����Զ�����');


-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------
--94102:��С��,2016-03-14,��Ӧ����Ȩ�޸���
Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1023, '��Ӧ����', User, '��������Ӧ��', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1023 And ���� = '��Ӧ����' And Upper(����) = Upper('��������Ӧ��'));

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1023, '��Ӧ����', User, '�����ļ��б�', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1023 And ���� = '��Ӧ����' And Upper(����) = Upper('�����ļ��б�'));

--93675:������,2016-03-14,���ѡ���������˴洢��������Ȩ��
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1252,'ҽ���´�',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All 
    Select 'Zl_������ҳ_��ҳ����EX','EXECUTE' From Dual Union All 
    Select 'Zl_������ҳ�ӱ�_��ҳ����','EXECUTE' From Dual Union All 
    Select '�ֻ��̶�','SELECT' From Dual Union All 
    Select '����������','SELECT' From Dual Union All 
	Select 'סԺ����ԭ��','SELECT' From Dual Union All 
	Select '������Ŀ','SELECT' From Dual Union All 
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1253,'ҽ���´�',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All 
    Select 'Zl_������ҳ_��ҳ����EX','EXECUTE' From Dual Union All 
    Select 'Zl_������ҳ�ӱ�_��ҳ����','EXECUTE' From Dual Union All 
    Select '�ֻ��̶�','SELECT' From Dual Union All 
    Select '����������','SELECT' From Dual Union All 
	Select 'סԺ����ԭ��','SELECT' From Dual Union All 
	Select '������Ŀ','SELECT' From Dual Union All 
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;


--90040:������,2016-03-17,���zlProgPrivs�������Ѿ��������֤��״̬�����ݾ�ֱ���޸�
update zlProgPrivs set ���� = '���֤δ¼ԭ��' where ϵͳ = &n_System  and ���� = '���֤��״̬';

--90040:������,2016-03-10,Ϊ�� ���֤δ¼ԭ�� ���Ȩ��
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1261,'��ҳ����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All 
    Select '���֤δ¼ԭ��','SELECT' From Dual Union All 
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1260,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All 
    Select '���֤δ¼ԭ��','SELECT' From Dual Union All 
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

--92729:������,2016-03-08,רҵ��RIS�ӿڴ���
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select  &n_System,1252,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All 
    Select 'RIS���ԤԼ','SELECT' From Dual Union All    
    Select 'Zl_Ris���ԤԼ_Insert','EXECUTE' From Dual Union All 
    Select 'Zl_Ris���ԤԼ_Delete','EXECUTE' From Dual Union All 
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select  &n_System,1253,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All 
    Select 'RIS���ԤԼ','SELECT' From Dual Union All    
    Select 'Zl_Ris���ԤԼ_Insert','EXECUTE' From Dual Union All 
    Select 'Zl_Ris���ԤԼ_Delete','EXECUTE' From Dual Union All 
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select  &n_System,1254,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All 
    Select 'RIS���ԤԼ','SELECT' From Dual Union All    
    Select 'Zl_Ris���ԤԼ_Insert','EXECUTE' From Dual Union All 
    Select 'Zl_Ris���ԤԼ_Delete','EXECUTE' From Dual Union All 
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

--93917:����,2016-03-08,��RIS�����ݽ����ӿڽű�
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��,ȱʡֵ)
	Select &n_System,1287,A.* From (
	Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0 Union All
	Select '����',-NULL,NULL,1 From Dual Union All       
	Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0) A;
       
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
	Select &n_System,1287,'����',User,A.* From (
	Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All
	Select '�ѱ�','SELECT' From Dual Union All
	Select '����','SELECT' From Dual Union All
	Select 'ְҵ','SELECT' From Dual Union All
	Select '�Ա�','SELECT' From Dual Union All
	Select '���ű�','SELECT' From Dual Union All
	Select '��Ա��','SELECT' From Dual Union All
	Select '������Ϣ','SELECT' From Dual Union All
	Select '������ҳ','SELECT' From Dual Union All
	Select '����״��','SELECT' From Dual Union All
	Select '�ϻ���Ա��','SELECT' From Dual Union All
	Select '��Ա֤���¼','SELECT' From Dual Union All
	Select '��������˵��','SELECT' From Dual Union All
	Select '���˹Һż�¼','SELECT' From Dual Union All
	Select '����ҽ����¼','SELECT' From Dual Union All
	Select '����ҽ������','SELECT' From Dual Union All
	Select '����ҽ������','SELECT' From Dual Union All
	Select '����ҽ������','SELECT' From Dual Union All
	Select '�������ҽ��','SELECT' From Dual Union All
	Select '������ϼ�¼','SELECT' From Dual Union All
	Select 'Ӱ������Ŀ','SELECT' From Dual Union All
	Select '�շ���ĿĿ¼','SELECT' From Dual Union All
	Select '������ü�¼','SELECT' From Dual Union All
	Select 'סԺ���ü�¼','SELECT' From Dual Union All
	Select '������ĿĿ¼','SELECT' From Dual Union All
	Select '������Ŀ��λ','SELECT' From Dual Union All
	Select '����ִ�п���','SELECT' From Dual Union All
	Select '�����շѹ�ϵ','SELECT' From Dual Union All
	Select '����֧����Ŀ','SELECT' From Dual Union All
	Select '���������Ա','SELECT' From Dual Union All
	Select '���������','SELECT' From Dual Union All
	Select 'ҽ�Ƹ��ʽ','SELECT' From Dual Union All
	Select 'ZlComponent','SELECT' From Dual Union All
	Select 'Zl_Lob_Append','EXECUTE' From Dual Union All 
	Select 'Zl_Lob_Read','EXECUTE' From Dual Union All
	Select 'Zl_Fun_Getsignpar','EXECUTE' From Dual Union All
	Select 'zl_Ӱ����Ϣ_XML���ݻ�ȡ','EXECUTE' From Dual Union All
	Select 'zl_�ҺŲ��˲���_INSERT','EXECUTE' From Dual Union All
	Select 'ZL_����ҽ����¼_Insert','EXECUTE' From Dual Union All
	Select 'ZL_����ҽ������_Insert','EXECUTE' From Dual Union All
	Select 'Zl_Ris���ԤԼ_Delete','EXECUTE' From Dual Union All
	Select 'Zl_Ris���ԤԼ_Insert','EXECUTE' From Dual Union All
	Select 'b_zlXWInterface','EXECUTE' From Dual Union All
	Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

--78768:���ϴ�,2016-03-04,IDKind����ȱʡ�������
Insert Into zlProgFuncs(ϵͳ,���,����,����,ȱʡֵ,˵��)
Select &n_System,1153,A.* From (
Select ����,����,ȱʡֵ,˵�� From zlProgFuncs Where 1 = 0 Union All
Select '��������',1,1,'���ò����Ĳ���Ȩ�ޡ��и�Ȩ��ʱ,������б��ز������á�' From Dual Union All
Select ����,����,ȱʡֵ,˵�� From zlProgFuncs Where 1 = 0) A;

--93758:����ԭ,2016-03-01,�ű�oracle����Ȩ�޷��ʴ���
Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1075, '����', User, '������������', 'SELECT'
  From Dual
  Where Not Exists (Select 1 From zlProgPrivs Where ϵͳ = &n_System And ��� = 1075 And ���� = '������������');

--92033:�ŵ���,2016-02-24,�����Ƿ񷢹�ҩ
Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1342, '����', User, 'Zl_��Һ��ҩ��¼_���', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1342 And ���� = '����' And Upper(����) = Upper('Zl_��Һ��ҩ��¼_���'));

--90447:������,2016-02-22,�����ȷ��
Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1111, '����', User, 'Zl1_Fun_Getreturnvisit', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1111 And ���� = '����' And Upper(����) = Upper('Zl1_Fun_Getreturnvisit'));

--91225:������,2016-02-18,ҽ��վʹ�ô�Ⱦ����������ӵ���Ȩ
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select  &n_System,1260,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All 
    Select '�������Լ�¼','SELECT' From Dual Union All 
    Select '�������淴��','SELECT' From Dual Union All
	Select '�����걨��¼','SELECT' From Dual Union All     
    Select 'Zl_�������Լ���¼_Update','EXECUTE' From Dual Union All 
    Select 'Zl_�����걨��¼_Update','EXECUTE' From Dual Union All 
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select  &n_System,1261,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All 
    Select '�������Լ�¼','SELECT' From Dual Union All 
    Select '�������淴��','SELECT' From Dual Union All 
	Select '�����걨��¼','SELECT' From Dual Union All    
    Select 'Zl_�������Լ���¼_Update','EXECUTE' From Dual Union All 
    Select 'Zl_�����걨��¼_Update','EXECUTE' From Dual Union All 
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

--93336:������,2016-02-16,��Ѫִ�еǼ�
Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1252, '����', User, 'Zl_Fun_Get��Ѫִ�еǼ�', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1252 And ���� = '����' And Upper(����) = Upper('Zl_Fun_Get��Ѫִ�еǼ�'));

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1253, '����', User, 'Zl_Fun_Get��Ѫִ�еǼ�', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1253 And ���� = '����' And Upper(����) = Upper('Zl_Fun_Get��Ѫִ�еǼ�'));

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1254, '����', User, 'Zl_Fun_Get��Ѫִ�еǼ�', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1254 And ���� = '����' And Upper(����) = Upper('Zl_Fun_Get��Ѫִ�еǼ�'));    

--92699:������,2016-01-22,�����������Ȩ������
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1265,'�����������',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All
Select 'zl_�����������_update','EXECUTE' From Dual Union All
Select 'zl_�����������_insert','EXECUTE' From Dual Union All
Select 'zl_�����������_delete','EXECUTE' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

--92808:����,2016-01-22,���ĵ��ݴ�ӡ����
Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1712, '����', User, 'ҩƷ�շ�����', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1712 And ���� = '����' And ���� = 'ҩƷ�շ�����');
Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1712, '����', User, 'Zl_ҩƷ�շ�����_Insert', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1712 And ���� = '����' And Upper(����) = Upper('Zl_ҩƷ�շ�����_Insert'));

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1713, '����', User, 'ҩƷ�շ�����', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1713 And ���� = '����' And ���� = 'ҩƷ�շ�����');
Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1713, '����', User, 'Zl_ҩƷ�շ�����_Insert', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1713 And ���� = '����' And Upper(����) = Upper('Zl_ҩƷ�շ�����_Insert'));    

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1714, '����', User, 'ҩƷ�շ�����', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1714 And ���� = '����' And ���� = 'ҩƷ�շ�����');
Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1714, '����', User, 'Zl_ҩƷ�շ�����_Insert', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1714 And ���� = '����' And Upper(����) = Upper('Zl_ҩƷ�շ�����_Insert'));  

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1716, '����', User, 'ҩƷ�շ�����', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1716 And ���� = '����' And ���� = 'ҩƷ�շ�����');
Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1716, '����', User, 'Zl_ҩƷ�շ�����_Insert', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1716 And ���� = '����' And Upper(����) = Upper('Zl_ҩƷ�շ�����_Insert'));  

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1717, '����', User, 'ҩƷ�շ�����', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1717 And ���� = '����' And ���� = 'ҩƷ�շ�����');
Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1717, '����', User, 'Zl_ҩƷ�շ�����_Insert', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1717 And ���� = '����' And Upper(����) = Upper('Zl_ҩƷ�շ�����_Insert'));  
         
Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1718, '����', User, 'ҩƷ�շ�����', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1718 And ���� = '����' And ���� = 'ҩƷ�շ�����');
Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1718, '����', User, 'Zl_ҩƷ�շ�����_Insert', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1718 And ���� = '����' And Upper(����) = Upper('Zl_ҩƷ�շ�����_Insert'));  

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1719, '����', User, 'ҩƷ�շ�����', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1719 And ���� = '����' And ���� = 'ҩƷ�շ�����');
Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1719, '����', User, 'Zl_ҩƷ�շ�����_Insert', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1719 And ���� = '����' And Upper(����) = Upper('Zl_ҩƷ�շ�����_Insert'));                    
         
Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1722, '����', User, 'ҩƷ�շ�����', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1722 And ���� = '����' And ���� = 'ҩƷ�շ�����');
Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1722, '����', User, 'Zl_ҩƷ�շ�����_Insert', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1722 And ���� = '����' And Upper(����) = Upper('Zl_ҩƷ�շ�����_Insert'));




-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--93917:����,2016-03-20,��RIS�����ݽ����ӿڽű�
Create Or Replace Function Zlpub_Pacs_��ȡ�������
(
  ҽ��id_In   In ����ҽ������.ҽ��id%Type,
  �������_In In ���Ӳ�������.�����ı�%Type
) Return Varchar2 Is

  v_Result        Varchar2(4000);
  v_Singleresult  Varchar2(4000);
  v_Reportcontent Ӱ�񱨸��¼.��������%Type;
  n_Count         Number(2);

  Cursor Cur_Report_Contents Is
    Select �������� From Ӱ�񱨸��¼ Where ҽ��id = ҽ��id_In;

Begin
  v_Result       := '';
  v_Singleresult := '';

  Select Count(1) Into n_Count From Ӱ�񱨸��¼ Where ҽ��id = ҽ��id_In;

  If n_Count > 0 Then
    For Row_Report_Contents In Cur_Report_Contents Loop
      v_Reportcontent := Row_Report_Contents.��������;
      Select Zlpub_Pacs_ȡ�������byxml(v_Reportcontent, �������_In) Into v_Singleresult From Dual;
      If v_Result Is Null And Not v_Singleresult Is Null Then
        v_Result := v_Singleresult;
      Elsif Not v_Singleresult Is Null Then
        v_Result := v_Result || ';' || v_Singleresult;
      End If;
    End Loop;
    Return v_Result;
  Else
    Select Count(1)
    Into n_Count
    From ���Ӳ������� A, ���Ӳ������� B, ����ҽ������ C
    Where a.�������� = 1 And a.Id = b.��id And b.�������� = 2 And b.��ʼ�� = 0 And a.�ļ�id = c.����id And c.ҽ��id = ҽ��id_In And
          a.�����ı� = �������_In;
  
    If n_Count > 0 Then
      Select b.�����ı�
      Into v_Result
      From ���Ӳ������� A, ���Ӳ������� B, ����ҽ������ C
      Where a.�������� = 1 And a.Id = b.��id And b.�������� = 2 And b.��ʼ�� = 0 And a.�ļ�id = c.����id And c.ҽ��id = ҽ��id_In And
            a.�����ı� = �������_In;
    Else
      Select b.�����ı�
      Into v_Result
      From ���Ӳ������� A, ���Ӳ������� B, ����ҽ������ C
      Where a.�������� = 3 And a.Id = b.��id And b.�������� = 2 And b.��ֹ�� = 0 And a.�ļ�id = c.����id And c.ҽ��id = ҽ��id_In And
            a.�����ı� = �������_In;
    End If;
  
    Return v_Result;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zlpub_Pacs_��ȡ�������;
/

--93675:������,2016-03-11,�ô洢����ֻ�������ѡ�����в�����ҳĳ���ֶεĸı��õ�
CREATE OR REPLACE Procedure Zl_������ҳ_��ҳ����EX
(
  ����id_In  In ������ҳ.����id%Type,
  ��ҳid_In  In ������ҳ.��ҳid%Type,
  ��Ϣ��_In  In varchar2,     /*������ҳ���ֶ���*/
  ��Ϣֵ_In  In varchar2      /*������ҳĳ�ֶε�ֵ*/
) Is
/*�ô洢����ֻ�������ѡ�����в�����ҳĳ���ֶεĸı��õ�*/
Begin
  If ��Ϣ��_In = '��Ժ��ʽ' Then
    update ������ҳ set ��Ժ��ʽ = ��Ϣֵ_In where ����id = ����id_In And ��ҳid = ��ҳid_In;
  elsif ��Ϣ��_In = 'ʬ���־' Then
     update ������ҳ set ʬ���־ = to_number(��Ϣֵ_In) where ����id = ����id_In And ��ҳid = ��ҳid_In;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_������ҳ_��ҳ����EX;
/

--92729:������,2016-03-18,רҵ��RIS�ӿ�
Create Or Replace Procedure Zl_Ris���ԤԼ_Insert
(
  ҽ��id_In     In Ris���ԤԼ.ҽ��id%Type,
  ԤԼid_In     In Ris���ԤԼ.ԤԼid%Type,
  ԤԼ����_In   In Ris���ԤԼ.ԤԼ����%Type,
  �豸id_In     In Ris���ԤԼ.����豸id%Type,
  �豸����_In   In Ris���ԤԼ.����豸����%Type,
  ��ʼʱ��_In   In Ris���ԤԼ.ԤԼ��ʼʱ��%Type,
  ����ʱ��_In   In Ris���ԤԼ.ԤԼ����ʱ��%Type,
  ��ʼʱ���_In In Ris���ԤԼ.ԤԼ��ʼʱ���%Type,
  ����ʱ���_In In Ris���ԤԼ.ԤԼ����ʱ���%Type
) Is
Begin
  Insert Into Ris���ԤԼ
    (ҽ��id, ԤԼid, ԤԼ����, ����豸id, ����豸����, ԤԼ��ʼʱ��, ԤԼ����ʱ��, ԤԼ��ʼʱ���, ԤԼ����ʱ���)
  Values
    (ҽ��id_In, ԤԼid_In, ԤԼ����_In, �豸id_In, �豸����_In, ��ʼʱ��_In, ����ʱ��_In, ��ʼʱ���_In, ����ʱ���_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Ris���ԤԼ_Insert;
/

--92729:������,2016-03-08,רҵ��RIS�ӿڴ���
Create Or Replace Procedure Zl_Ris���ԤԼ_Delete(ҽ��id_In In Ris���ԤԼ.ҽ��id%Type) Is
Begin
  Delete Ris���ԤԼ Where ҽ��id = ҽ��id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Ris���ԤԼ_Delete;
/

--92729:������,2016-03-08,רҵ��RIS�ӿڴ���
Create Or Replace Procedure Zl1_Datamove_Reb
(
  System_In    In Number,
  Speedmode_In In Number,
  Func_In      In Number,
  Enable_In    In Number := 0,
  Parallel_In  In Number := 0,
  Rebscope_In  In Number := 0
) As
  --���ܣ�����ʷ����ת��֮ǰ�����ô��������Զ���ҵ��Լ����������ת��֮��������Щ�����Լ��ؽ���ת���������ջر��ת�����������Ŀռ� 
  --������ 
  --System_In:    Ӧ��ϵͳ���,100=��׼�� 
  --speedmode_in������ת��ģʽ��0-����ģʽ��1-����ģʽ���ڿͻ���ͣ��ʱ��ת���ڼ����ת�����������Ψһ�������Լ�����������Լӿ���ת���ݵ�ɾ�������� 
  --func_in:      1=��������2=�Զ���ҵ��3=Լ����4=������5=�ؽ���ת��������6-�ջر��ת�����������Ŀռ䣬7-�����Ĵ洢�ռ䣨move�������ָ������õ�Լ�������� ,8-�ؽ����ת����ѯ��������������������� 
  --Enable_in:    0-���ã�1=���ã���func_inֵΪ1-4��Ч 
  --rebScope_in:   Func_In=6ʱ��ָ�ؽ������ķ�Χ(0-���ú�����,1-���ú����༰ҽ����,2-ȫ��)��Func_In=7ʱָMove��ķ�Χ(0-���ú����࣬1-ȫ��) 

  v_Sql      Varchar2(4000);
  n_Do       Number(1);
  n_Parallel Number(1);
  v_Tbs      Varchar2(100);

 --ת������е�SQL��ѯ���������
  v_Indexeswithtag Varchar2(4000) := '������ü�¼_IX_����ID,סԺ���ü�¼_IX_����ID,���ò����¼_IX_����ID,���ò����¼_IX_�Ǽ�ʱ��,����Ԥ����¼_IX_��ҳID,����Ԥ����¼_IX_����ID,����Ԥ����¼_IX_�տ�ʱ��,������ü�¼_IX_�Ǽ�ʱ��,������ü�¼_IX_ҽ�����,סԺ���ü�¼_IX_�Ǽ�ʱ��,���˽��ʼ�¼_IX_�շ�ʱ��,���˽��ʼ�¼_IX_����id' ||
                                     ',ҩƷ�շ���¼_IX_����ID,�շ���¼������Ϣ_IX_�շ�ID,��Һ��ҩ����_IX_�շ�ID,ҩƷ����ƻ�_IX_����ID,ҩƷǩ����ϸ_IX_�շ�ID' ||
                                     ',��Ա����¼_IX_���ʱ��,��Ա�սɼ�¼_IX_�Ǽ�ʱ��,��Ա�ݴ��¼_IX_�ս�ID,��Ա�ݴ��¼_IX_�Ǽ�ʱ��,Ʊ�����ü�¼_IX_�Ǽ�ʱ��,Ʊ��ʹ����ϸ_IX_����ID,Ʊ�ݴ�ӡ��ϸ_IX_ʹ��ID' ||
                                     ',���˹Һż�¼_IX_�Ǽ�ʱ��,����ҽ������_IX_����ʱ��,����ҽ����¼_IX_�Һŵ�,����ҽ����¼_IX_��ҳID,����ҽ����¼_IX_���ID' ||
                                     ',������ҳ_IX_��Ժ����,סԺ���ü�¼_IX_����ID,���˹�����¼_IX_����ID,������ϼ�¼_IX_����ID,���������¼_IX_��ҳID' ||
                                     ',���˻����¼_IX_��ҳID,���˻�������_IX_��¼id,���˻����ļ�_IX_��ҳID,���˻�������_IX_�ļ�ID,���˻�����ϸ_IX_��¼ID,���˻����ӡ_IX_�ļ�ID' ||
                                     ',���Ӳ�����¼_IX_����ID,����ҽ������_IX_����ID,Ӱ�񱨸沵��_IX_ҽ��ID,������ļ�¼_IX_����ID,������ϼ�¼_IX_����ID' ||
                                     ',�����ٴ�·��_IX_����ID,���˺ϲ�·��_IX_��Ҫ·����¼ID,����·��ִ��_IX_·����¼ID,���˳�����¼_IX_·����¼ID,�������ҽ��_IX_ҽ��ID' ||
                                     ',Ӱ�񱨸��¼_IX_ҽ��ID,Ӱ�񱨸������¼_IX_ҽ��ID,Ӱ�����뵥ͼ��_IX_ҽ��ID,Ӱ���ղ�����_IX_ҽ��ID,����걾��¼_IX_ҽ��ID,������Ŀ�ֲ�_IX_�걾ID,���������¼_IX_�걾ID' ||
                                     ',���������¼_IX_�걾ID,����ͼ����_IX_�걾ID,������ռ�¼_IX_ҽ��ID,������ͨ���_IX_����걾ID'; 

  --ת������е�SQL��ѯ���������(������Ψһ����Ӧ������)
  v_Constraintswithtag Varchar2(4000) := '����Ԥ����¼_UQ_NO,���˽��ʼ�¼_UQ_NO,���˽��ʼ�¼_PK,������ü�¼_UQ_NO,סԺ���ü�¼_UQ_NO,ҽ��������ϸ_PK' ||
                                         ',���˿��������_PK,���ò����¼_PK,���˿������¼_PK,�������㽻��_PK,�����˿���Ϣ_PK,��Һ��ҩ��¼_PK,ҩƷǩ����¼_PK,Ʊ�ݴ�ӡ����_PK,���˹Һż�¼_PK,���˹ҺŻ���_UQ_����,����ת���¼_UQ_NO' ||
                                         ',���˻�����Ŀ_UQ_ҳ��,���˻���Ҫ������_UQ_ҳ��,����Ҫ������_PK,���Ӳ�����¼_PK,���Ӳ�������_PK,���Ӳ�����ʽ_PK,���Ӳ�������_UQ_�������,���Ӳ���ͼ��_PK,�����걨��¼_PK,�������淴��_PK' ||
                                         ',���˺ϲ�·������_PK,����·������_PK,����·������_PK,����·��ָ��_UQ_����ָ��,����·��ҽ��_PK' ||
                                         ',����ҽ����¼_PK,����ҽ������_PK,����ҽ���Ƽ�_UQ_�շ�ϸĿID,����ҽ������_PK,����ҽ������_PK,����ҽ��ִ��_PK,ҽ��ִ��ʱ��_PK,ҽ��ִ�д�ӡ_PK,����ҽ����ӡ_UQ_ҽ��ID,��Ѫ�����¼_PK,��Ѫ������_PK' ||
                                         ',������ϼ�¼_PK,����ҽ��״̬_PK,ҽ��ǩ����¼_PK,����ҽ������_PK,���Ƶ��ݴ�ӡ_PK,ҽ��ִ�мƼ�_PK,ִ�д�ӡ��¼_PK,RIS���ԤԼ_PK' ||
                                         ',Ӱ�����¼_PK,Ӱ��������_UQ_���к�,Ӱ����ͼ��_UQ_ͼ���,Ӱ��Σ��ֵ��¼_UQ_ҽ��ID' ||
                                         ',����������Ŀ_PK,�����ʿؼ�¼_PK,����ǩ����¼_PK,�����Լ���¼_PK,�����ʿر���_PK,����ҩ�����_PK,��Ա�սɼ�¼_PK,��Ա�ս���ϸ_PK,��Ա�ս�Ʊ��_PK,��Ա�սɶ���_PK';

  --���ܣ�1.���û���������ת�����������������,����ɾ�������¼ʱ���ӱ�ÿ�м�¼ִ��һ��SQL��ѯ��ɾ�� 
  --      2.���û�����������Ψһ��Լ��������ʱ���Զ�ɾ����Ӧ������������ʱ�Զ������������������ɾ������ 
  --���磺����ҽ������_FK_ҽ��ID�������Щ������ڵı�����δת����δ��zlbaktables���ж��壩��ִ��ǰ���鲢����ת���� 
  Procedure Setconstraintstatus As
    v_Pcol Varchar2(50);
    v_Fcol Varchar2(50);
    v_Del  Varchar2(4000);
  Begin
    --����ʱ���Ƚ�������ת��������������������ٽ���ת��������� 
    If Enable_In = 0 Then
      --1.����ģʽת��ʱ��������ҵ�����ɾ�����������ԣ����ڼ���ɾ����������ô�������������ӱ����ݵ�ɾ������
      If Speedmode_In = 0 Then
        For Rp In (Select Distinct a.Table_Name As Ptable_Name, a.Constraint_Name
                   From User_Constraints A, User_Constraints C, zlBakTables B
                   Where a.Table_Name = b.���� And b.ֱ��ת�� = 1 And b.ϵͳ = System_In And a.Constraint_Type In ('P', 'U') And
                         c.r_Constraint_Name = a.Constraint_Name And c.Constraint_Type = 'R' And
                         c.Delete_Rule = 'CASCADE'
                   Order By a.Table_Name) Loop
        
          Select f_List2str(Cast(Collect(Column_Name Order By Position) As t_Strlist))
          Into v_Pcol
          From User_Cons_Columns
          Where Constraint_Name = Rp.Constraint_Name;
        
    v_Del := '';
          For Rf In (Select b.Table_Name, b.Constraint_Name,
                            f_List2str(Cast(Collect(b.Column_Name Order By b.Position) As t_Strlist)) As r_Col
                     From User_Constraints A, User_Cons_Columns B
                     Where a.r_Constraint_Name = Rp.Constraint_Name And a.Constraint_Name = b.Constraint_Name
                     Group By b.Table_Name, b.Constraint_Name) Loop
            If Instr(v_Pcol, ',') > 0 Then
              v_Del := v_Del || Chr(10) || '        Delete ' || Rf.Table_Name || ' Where (' || Rf.r_Col ||
                       ') in ((:Old.' || Replace(v_Pcol, ',', ',:Old.') || '));';
            Else
              v_Del := v_Del || Chr(10) || '        Delete ' || Rf.Table_Name || ' Where ' || Rf.r_Col || ' = :Old.' ||
                       v_Pcol || ';';
            End If;
          End Loop;
        
          v_Sql := 'Create Or Replace Trigger ' || Rp.Ptable_Name || '_Cascade_Del' || Chr(10) ||
                   '    After Delete On ' || Rp.Ptable_Name || Chr(10) || '    For Each Row' || Chr(10) || 'Begin' ||
                   Chr(10) || '    If :Old.��ת�� Is Null Then ' || v_Del || Chr(10) || '    End If; ' || Chr(10) ||
                   'End ' || Rp.Ptable_Name || '_Cascade_Del;';
          Execute Immediate v_Sql;
        End Loop;
      End If;
    
      --2.��������ת�����������������
      For R In (Select c.Table_Name, c.Constraint_Name, a.Table_Name As Ptable_Name
                From User_Constraints A, User_Constraints C, zlBakTables B
                Where a.Table_Name = b.���� And b.ֱ��ת�� = 1 And b.ϵͳ = System_In And a.Constraint_Type In ('P', 'U') And
                      c.r_Constraint_Name = a.Constraint_Name And c.Constraint_Type = 'R' And c.Status = 'ENABLED'
                Order By a.Table_Name) Loop
        v_Sql := 'Alter Table ' || r.Table_Name || ' Disable Constraint ' || r.Constraint_Name;
        Execute Immediate v_Sql;
      End Loop;
    
      --3.����������Ψһ������(����ת��ʱ)
      If Speedmode_In = 1 Then
        --����ɾ������������ʹskip_unusable_indexesΪtrue��Ҳ�޷�ɾ������Unusable״̬��Ψһ�������ı��еļ�¼
        --����ת������е�SQL��ѯ���������(������Ψһ����Ӧ������) 
        For R In (Select a.Table_Name, a.Constraint_Name
                  From User_Constraints A, zlBakTables T, User_Tables B
                  Where a.Table_Name = t.���� And t.ֱ��ת�� = 1 And t.ϵͳ = System_In And a.Status = 'ENABLED' And
                        a.Constraint_Type In ('P', 'U') And a.Table_Name = b.Table_Name And b.Iot_Type Is Null And
                        a.Constraint_Name Not In
                        (Select Upper(Column_Value) As Constraint_Name From Table(f_Str2list(v_Constraintswithtag)))
                  Order By Constraint_Name) Loop
          v_Sql := 'Alter Table ' || r.Table_Name || ' Disable Constraint ' || r.Constraint_Name ||
                   ' Cascade Drop Index';
          Execute Immediate v_Sql;
        End Loop;
      End If;
    Else
      --����ʱ
      --1.������������Ψһ��������������ת����������������� 
      If Speedmode_In = 1 Then
        --���ؽ�������������Լ�����Ա��ؽ�����ʱ���ò���ִ������ʱ�䣬��������Լ��ʱҲ���Բ���novalidate��ʽ 
        For R In (Select d.Table_Name, d.Constraint_Name,
                         f_List2str(Cast(Collect(d.Column_Name Order By d.Position) As t_Strlist)) Colstr
                  From User_Cons_Columns D,
                       (Select a.Table_Name, a.Constraint_Name
                         From User_Constraints A, zlBakTables T
                         Where a.Table_Name = t.���� And t.ֱ��ת�� = 1 And t.ϵͳ = System_In And a.Status = 'DISABLED' And
                               a.Constraint_Type In ('P', 'U')) A
                  Where a.Constraint_Name = d.Constraint_Name And a.Table_Name = d.Table_Name
                  Group By d.Table_Name, d.Constraint_Name
                  Order By Constraint_Name) Loop
          Update Zldatamovelog
          Set ��ǰ���� = '���ڻָ�Լ��:' || r.Constraint_Name
          Where ϵͳ = System_In And ���� = (Select Max(����) From Zldatamovelog Where ϵͳ = System_In);
        
          Select Tablespace_Name Into v_Tbs From User_Indexes Where Table_Name = r.Table_Name And Rownum < 2;
        
          --����������Ψһ��ʱ�������Ǳ�ɾ���˵ģ���������Ҫ��Create 
          v_Sql := 'Create Unique Index ' || r.Constraint_Name || ' On ' || r.Table_Name || '(' || r.Colstr ||
                   ') Tablespace ' || v_Tbs || ' Nologging';
          Begin
            Execute Immediate v_Sql;
          Exception
            When Others Then
              Null; --������Щ������Ψһ�����Ǳ���ת���ڼ䱻���õģ�֮ǰ�ʹ��ڲ�Ψһ���ݣ�����Ψһ��������� 
          End;
        
          --���Զ�����Լ���������Ĺ��� 
          v_Sql := 'Alter Table ' || r.Table_Name || ' Enable Novalidate Constraint ' || r.Constraint_Name;
          Execute Immediate v_Sql;
        End Loop;
      End If;
    
      --2.��������ת����������������� 
      For R In (Select c.Table_Name, c.Constraint_Name
                From User_Constraints A, User_Constraints C, zlBakTables B
                Where a.Table_Name = b.���� And b.ֱ��ת�� = 1 And b.ϵͳ = System_In And a.Constraint_Type In ('P', 'U') And
                      c.r_Constraint_Name = a.Constraint_Name And c.Constraint_Type = 'R' And c.Status = 'DISABLED'
                Order By a.Table_Name) Loop
        --Ϊ�˼ӿ��ٶȣ�����novalidate������֤�������� 
        --��������ת����������������zlbaktables�ж����ˣ���û�б�д��Ӧ������ת���ű���δ��֤�����ݿ�����Υ��Լ��������� 
        v_Sql := 'Alter Table ' || r.Table_Name || ' Enable Novalidate Constraint ' || r.Constraint_Name;
        Execute Immediate v_Sql;
      End Loop;
    
      --3.����ģʽת��ʱ��ɾ��֮ǰ�����������������ɾ������Ĵ�����
      If Speedmode_In = 0 Then
        For R In (Select a.Trigger_Name
                  From User_Triggers A, zlBakTables B
                  Where a.Table_Name = b.���� And b.ֱ��ת�� = 1 And b.ϵͳ = System_In And
                        Trigger_Name = Table_Name || '_CASCADE_DEL' And Triggering_Event = 'DELETE') Loop
          v_Sql := 'Drop Trigger ' || r.Trigger_Name;
          Execute Immediate v_Sql;
        End Loop;
      End If;
    End If;
  End Setconstraintstatus;

  --���ܣ�����ģʽʱ����LOB�������������������ģʽʱ������ת�������÷�ת������������(���磺����ҽ���Ƽ�_IX_�շ�ϸĿID) 
  --˵��������������Ϊ�����ɾ�����ݵ����� 
  Procedure Setindexstatus As
  Begin
    If Speedmode_In = 1 Then
      --����ת������е�SQL��ѯ��������� 
      For R In (Select /*+ rule*/
                 a.Index_Name
                From User_Indexes A, zlBakTables T
                Where a.Table_Name = t.���� And t.ֱ��ת�� = 1 And t.ϵͳ = System_In And t.ֱ��ת�� = 1 And
                      a.Index_Name <> a.Table_Name || '_IX_��ת��' And
                      a.Index_Name Not In
                      (Select Upper(Column_Value) As Index_Name From Table(f_Str2list(v_Indexeswithtag))) And
                      a.Status = Decode(Enable_In, 0, 'VALID', 'UNUSABLE') And a.Index_Type = 'NORMAL' And Not Exists
                 (Select 1
                       From User_Constraints C
                       Where c.Index_Name = a.Index_Name And c.Constraint_Type In ('P', 'U'))
                Order By Index_Name) Loop
      
        If Enable_In = 0 Then
          v_Sql := 'Alter Index ' || r.Index_Name || ' Unusable';
          Execute Immediate v_Sql;
        Else
          Update Zldatamovelog
          Set ��ǰ���� = '�����ؽ�����:' || r.Index_Name
          Where ϵͳ = System_In And ���� = (Select Max(����) From Zldatamovelog Where ϵͳ = System_In);
        
          v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Nologging';
          Begin
            Execute Immediate v_Sql;
            --�����ؽ��Ƚ������������ؽ�����Ҫ���������������������������ORA-00054: ��Դ��æ, ��ָ���� NOWAIT ��ʽ��ȡ��Դ 
          
          Exception
            When Others Then
              If SQLErrM Like 'ORA-00054%' Then
                v_Sql := Replace(v_Sql, 'Rebuild', 'Rebuild Online');
                Execute Immediate v_Sql;
              End If;
          End;
        End If;
      End Loop;
    Else
      For R In (Select a.Index_Name
                From (Select d.Table_Name, d.Index_Name,
                              f_List2str(Cast(Collect(d.Column_Name Order By d.Column_Position) As t_Strlist)) Colstr
                       From User_Ind_Columns D, zlBakTables T, User_Indexes C
                       Where c.Table_Name = t.���� And t.ֱ��ת�� = 1 And t.ϵͳ = System_In And c.Uniqueness = 'NONUNIQUE' And
                             c.Index_Type = 'NORMAL' And c.Status = Decode(Enable_In, 0, 'VALID', 'UNUSABLE') And
                             c.Index_Name = d.Index_Name And c.Table_Name = d.Table_Name
                       Group By d.Table_Name, d.Index_Name) A,
                     (Select e.Table_Name,
                              f_List2str(Cast(Collect(e.Column_Name Order By e.Position) As t_Strlist)) Colstr
                       From User_Cons_Columns E, User_Constraints F, zlBakTables T, User_Constraints C
                       Where e.Table_Name = t.���� And t.ֱ��ת�� = 1 And t.ϵͳ = System_In And
                             e.Constraint_Name = f.Constraint_Name And f.Constraint_Type = 'R' And
                             c.Constraint_Name = f.r_Constraint_Name And c.Table_Name Not In ('������ҳ', '������Ϣ') And
                             Not Exists
                        (Select 1 From zlBakTables G Where g.���� = c.Table_Name And g.ϵͳ = System_In)
                       Group By e.Table_Name, e.Constraint_Name) B
                Where a.Table_Name = b.Table_Name And a.Colstr = b.Colstr
                Order By Index_Name) Loop
      
        If Enable_In = 0 Then
          --���⴦�������������������ã�������ҩƷĿ¼�޸Ĺ�񣬲���ɿ���Ҫʹ�� 
          If r.Index_Name Not In ('����ҽ����¼_IX_�շ�ϸĿID', 'ҩƷ�շ���¼_IX_ҩƷID', 'ҩƷ�շ���¼_IX_�۸�ID') Then
            v_Sql := 'Alter Index ' || r.Index_Name || ' Unusable';
            Execute Immediate v_Sql;
          End If;
        Else
          Update Zldatamovelog
          Set ��ǰ���� = '�����ؽ�����:' || r.Index_Name
          Where ϵͳ = System_In And ���� = (Select Max(����) From Zldatamovelog Where ϵͳ = System_In);
        
          v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Online Nologging';
          Execute Immediate v_Sql;
          --�����ؽ��Ƚ������������ؽ�����Ҫ���������������������������ORA-00054: ��Դ��æ, ��ָ���� NOWAIT ��ʽ��ȡ��Դ 
        End If;
      End Loop;
    End If;
  End Setindexstatus;

  --���ܣ�ת�������ڼ䣬ͣ��ת�����ϵ����д�������ת�����ٻָ� 
  Procedure Settriggerstatus As
  Begin
    For R In (Select Distinct a.Table_Name, t.ͣ�ô�����
              From User_Triggers A, zlBakTables T
              Where a.Status = Decode(Enable_In, 0, 'ENABLED', 'DISABLED') And a.Table_Name = t.���� And t.ֱ��ת�� = 1 And
                    t.ϵͳ = System_In) Loop
      If Enable_In = 0 Then
        v_Sql := 'Alter Table ' || r.Table_Name || ' DISABLE ALL TRIGGERS';
        Update zlBakTables Set ͣ�ô����� = 1 Where ϵͳ = System_In And ���� = r.Table_Name;
      Elsif Nvl(r.ͣ�ô�����, 0) = 1 Then
        v_Sql := 'Alter Table ' || r.Table_Name || ' ENABLE ALL TRIGGERS';
        Update zlBakTables Set ͣ�ô����� = Null Where ϵͳ = System_In And ���� = r.Table_Name;
      End If;
      Execute Immediate v_Sql;
    End Loop;
    Commit;
  End Settriggerstatus;

  --���ܣ�ת�������ڼ䣬ͣ�õ�ǰ�����ߵ������Զ���ҵ��ת���������� 
  Procedure Setjobstatus As
    v_Jobs Varchar2(4000);
  Begin
    --ͣ�� 
    If Enable_In = 0 Then
      For R In (Select Job From User_Jobs Where Broken = 'N') Loop
        Dbms_Job.Broken(r.Job, True);
        v_Jobs := v_Jobs || ',' || r.Job;
      End Loop;
    
      If v_Jobs Is Not Null Then
        v_Jobs := Substr(v_Jobs, 2);
        Update zlDataMove Set ͣ����ҵ�� = v_Jobs Where ϵͳ = System_In And ��� = 1;
      End If;
    Else
      --���� 
      Select ͣ����ҵ�� Into v_Jobs From zlDataMove Where ϵͳ = System_In And ��� = 1;
      If v_Jobs Is Not Null Then
        For R In (Select Job
                  From User_Jobs
                  Where Broken = 'Y' And Job In (Select Column_Value From Table(f_Num2list(v_Jobs)))) Loop
          Dbms_Job.Broken(r.Job, False);
        End Loop;
        Update zlDataMove Set ͣ����ҵ�� = Null Where ϵͳ = System_In And ��� = 1;
      End If;
    End If;
    --��ҵ���ú�����ύ�������Ч 
    Commit;
  End Setjobstatus;
Begin
  If Parallel_In < 2 Then
    Execute Immediate 'Alter Session DISABLE PARALLEL DDL';
  Else
    If Func_In In (6, 7, 8) Or Func_In In (3, 4) And Enable_In = 1 Then
      --Ϊ�ؽ��������ò���ִ�У�����ͨ��������IO�豸�����ܣ�����̫�ߵĲ��жȷ����ή�����ܣ����и����ܴ洢�豸���ɼӴ��жȣ� 
      --ִ���ؽ���������Զ�Ϊ�������ϲ��ж����ԣ������ȡ������Ӱ�����SQL��ִ�мƻ�(ȫ��ɨ��+���в�ѯ������),�ں���ȡ�������Ĳ��ж� 
      --�ָ����߿��Լ��������ʱ�������ǲ�������ģʽ�������ϲ��У�����̫��
      Execute Immediate 'Alter Session FORCE PARALLEL DDL PARALLEL ' || Parallel_In;
      n_Parallel := 1;
    End If;
  End If;

  If Func_In = 1 Then
    --1.���ô����� 
    Settriggerstatus;
  Elsif Func_In = 2 Then
    --2.�����Զ���ҵ 
    Setjobstatus;
  Elsif Func_In = 3 Then
    --3.����Լ��״̬ 
    Setconstraintstatus;
  Elsif Func_In = 4 Then
    --4.��������״̬ 
    Setindexstatus;
  Elsif Func_In = 5 Then
    --5.�ؽ�"��ת��"���� 
    For R In (Select b.Index_Name
              From zlBakTables A, User_Indexes B
              Where a.���� = b.Table_Name And a.ֱ��ת�� = 1 And a.ϵͳ = System_In And b.Index_Name = b.Table_Name || '_IX_��ת��'
              Union All
              Select '������ҳ_IX_��ת��'
              From Dual
              Where System_In = 100) Loop
      Update Zldatamovelog
      Set ��ǰ���� = '�����ؽ�����:' || r.Index_Name
      Where ϵͳ = System_In And ���� = (Select Max(����) From Zldatamovelog Where ϵͳ = System_In);
    
      --��ʱ̫�̣����벢��DDL 
      --����ת��ʱ����ؽ����������������������������������ORA-00054: ��Դ��æ, ��ָ���� NOWAIT ��ʽ��ȡ��Դ 
      --�����ؽ�����̫�������ԣ���ʹ����ת��ģʽҲ���������ؽ�
      v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Nologging';
      Begin
        Execute Immediate v_Sql;
      Exception
        When Others Then
          If SQLErrM Like 'ORA-00054%' Then
            v_Sql := Replace(v_Sql, 'Rebuild', 'Rebuild Online');
            Execute Immediate v_Sql;
          End If;
      End;
    End Loop;
  
  Elsif Func_In = 6 Then
    --6.�ؽ����ת����ѯ���õ������������Ա����ؽ�����������һ��Ĳ�ѯʱ�䣩 
    --����ҵ������ý׶��������ؽ���Щ�������Ա���һЩ����Ҫ���ؽ���ʱ 
    For R In (Select b.Index_Name, a.���
              From User_Indexes B, zlBakTables A
              Where a.ϵͳ = System_In And a.���� = b.Table_Name And
                    b.Index_Name In (Select Upper(Column_Value)
                                     From Table(f_Str2list(v_Indexeswithtag))
                                     Union
                                     Select Upper(Column_Value)
                                     From Table(f_Str2list(v_Constraintswithtag)))
              Order By Index_Name) Loop
      n_Do := 0;
      If Rebscope_In = 0 Then
        If r.��� < 5 Then
          n_Do := 1; --�����ú����� 
        End If;
      Elsif Rebscope_In = 1 Then
        If r.��� < 5 Or r.��� = 8 Then
          n_Do := 1; --�����ú����ࡢҽ���� 
        End If;
      Else
        n_Do := 1;
      End If;
    
      If n_Do = 1 Then
        Update Zldatamovelog
        Set ��ǰ���� = '�����ؽ�����:' || r.Index_Name
        Where ϵͳ = System_In And ���� = (Select Max(����) From Zldatamovelog Where ϵͳ = System_In);
      
        --v_Sql := 'Alter Index ' || r.Index_Name || ' shrink Space'; 
        --ʹ��shrink��ʽ���ܲ���ִ��,��������ٶȱ�rebuild PARALLEL 8 ��6�� 
        If Speedmode_In = 1 Then
          v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Nologging';
        Else
          v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Online Nologging';
        End If;
        Begin
          Execute Immediate v_Sql;
          --�����ؽ��Ƚ������������ؽ�����Ҫ���������������������������ORA-00054: ��Դ��æ, ��ָ���� NOWAIT ��ʽ��ȡ��Դ 
        
        Exception
          When Others Then
            If SQLErrM Like 'ORA-00054%' Then
              v_Sql := Replace(v_Sql, 'Rebuild', 'Rebuild Online');
              Execute Immediate v_Sql;
            End If;
        End;
      End If;
    End Loop;
  
    --����������
  Elsif Func_In = 7 Then
    --rebScope_in=0,ֻ�������С��5�ľ��ú���������á�ҩƷ��Ʊ�ݣ�������ȫ������ 
    For R In (Select a.���� As Table_Name
              From zlBakTables A
              Where a.ֱ��ת�� = 1 And (��� < Decode(Rebscope_In, 0, 5, 100))
              Order By ���, ���) Loop
    
      Update Zldatamovelog
      Set ��ǰ���� = '���������:' || r.Table_Name
      Where ϵͳ = System_In And ���� = (Select Max(����) From Zldatamovelog Where ϵͳ = System_In);
    
      --����п��еĿռ䣬����Ƶ�������ռ䣬ֻ���������ܾ����ƶ��ļ�β�������ݿ飬�Ա���б�ռ��ļ������� 
      --��ǰ�������˻Ự����ǿ�Ʋ��� 
      v_Sql := 'Alter Table ' || r.Table_Name || ' Move Nologging';
      Execute Immediate v_Sql;
    
      --�����ƶ�Lob���� 
      For L In (Select Column_Name, Tablespace_Name From User_Lobs Where Table_Name = r.Table_Name) Loop
        v_Sql := 'Alter Table ' || r.Table_Name || ' Move Lob(' || l.Column_Name || ') Store as (Tablespace ' ||
                 l.Tablespace_Name || ') Nologging';
        Execute Immediate v_Sql;
      End Loop;
    
      v_Sql := 'Alter Table ' || r.Table_Name || ' Noparallel';
      Execute Immediate v_Sql;
    
      --move�󣬱���ص�������ȫ��ʧЧ����Ҫȫ���ؽ� 
      For S In (Select Index_Name
                From User_Indexes
                Where Table_Name = r.Table_Name And Status = 'UNUSABLE'
                Order By Index_Name) Loop
        Update Zldatamovelog
        Set ��ǰ���� = '���ڻָ�ʧЧ����:' || s.Index_Name
        Where ϵͳ = System_In And ���� = (Select Max(����) From Zldatamovelog Where ϵͳ = System_In);
      
        --��ǰ�������˻Ự����ǿ�Ʋ��� 
        v_Sql := 'Alter Index ' || s.Index_Name || ' Rebuild Nologging';
        Execute Immediate v_Sql;
      End Loop;
    End Loop;
    --�ؽ�ת�����ϱ��ת���������������������ת����ɺ��ջؿ��пռ䣩
    --ʧЧ���������ؽ�����Ϊת������е������ؽ�����
  Elsif Func_In = 8 Then
    For R In (Select b.Index_Name, a.���
              From User_Indexes B, zlBakTables A
              Where a.ϵͳ = System_In And a.���� = b.Table_Name And b.Status = 'VALID' And b.Index_Type = 'NORMAL' And
                    b.Index_Name Not Like 'BIN$%' And
                    b.Index_Name Not In (Select Upper(Column_Value)
                                         From Table(f_Str2list(v_Indexeswithtag))
                                         Union
                                         Select Upper(Column_Value)
                                         From Table(f_Str2list(v_Constraintswithtag)))
              Order By Index_Name) Loop
      Update Zldatamovelog
      Set ��ǰ���� = '�����ؽ�����:' || r.Index_Name
      Where ϵͳ = System_In And ���� = (Select Max(����) From Zldatamovelog Where ϵͳ = System_In);
    
      If Speedmode_In = 1 Then
        v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Nologging';
      Else
        v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Online Nologging';
      End If;
      Begin
        Execute Immediate v_Sql;
        --�����ؽ��Ƚ������������ؽ�����Ҫ���������������������������ORA-00054: ��Դ��æ, ��ָ���� NOWAIT ��ʽ��ȡ��Դ    
      Exception
        When Others Then
          If SQLErrM Like 'ORA-00054%' Then
            v_Sql := Replace(v_Sql, 'Rebuild', 'Rebuild Online');
            Execute Immediate v_Sql;
          End If;
      End;
    End Loop;
  End If;

  --ִ���ؽ���������Զ�Ϊ�������ϲ��ж����ԣ������ȡ������Ӱ�����SQL��ִ�мƻ�(ȫ��ɨ��+���в�ѯ������) 
  --------------------------------------------------------------------------------------------------- 
  If n_Parallel = 1 Then
    Execute Immediate 'ALTER Session DISABLE PARALLEL DDL';
  
    For R In (Select Index_Name From User_Indexes Where Degree Not In ('1', '0')) Loop
      v_Sql := 'Alter Index ' || r.Index_Name || ' Noparallel';
      Execute Immediate v_Sql;
    End Loop;
  End If;

  Update Zldatamovelog
  Set ��ǰ���� = '�ؽ����'
  Where ϵͳ = System_In And ���� = (Select Max(����) From Zldatamovelog Where ϵͳ = System_In);
  Commit;
  --�����̲����д����������ɵ��ù��̴��� 
End Zl1_Datamove_Reb;
/

--92729:������,2016-03-08,רҵ��RIS�ӿڴ���
Create Or Replace Procedure Zl1_Datamove_Tag
(
  d_End    In Date,
  n_����   In Number,
  n_System In Number
) As
  --���ܣ���Ǵ�ת�������� 
  --˵����Ϊ����Undo��ռ����͹��󣬷ֶ��ύ 
Begin
  --1.���ú��㣨����,ҩƷ,�տ��Ʊ�ݵȣ�  
  --�¼��Ӳ�ѯע�������Ż������ܹ������ݹ��˵���С�������ŵ����Exists��������ǰ��
  Update /*+ rule*/ ����Ԥ����¼ L
  Set ��ת�� = n_����
  Where ����id In
        (Select Distinct a.����id --1.�����շѺ͹Һŵ��շѽ����¼(�ų�֮���˺ź��˷ѵ�,һ�ŵ�����ֻҪ����һ������) 
         From ������ü�¼ A
         Where (a.��¼״̬ In (1, 2) Or a.��¼״̬ = 3 And Not Exists
                (Select 1
                 From ������ü�¼ B
                 Where a.No = b.No And a.��¼���� = b.��¼���� And b.��¼״̬ = 2 And b.�Ǽ�ʱ�� >= d_End))
     And a.��ת�� Is Null And a.��¼���� In (1, 4) And a.�Ǽ�ʱ�� < d_End
         Union All
         Select Distinct a.����id --2.ҽ�������� 
         From ���ò����¼ A
         Where (a.��¼״̬ In (1, 2) Or a.��¼״̬ = 3 And Not Exists
                (Select 1
                 From ���ò����¼ B
                 Where a.No = b.No And a.��¼���� = b.��¼���� And b.��¼״̬ In (1, 2) And b.�Ǽ�ʱ�� >= d_End))
     And a.��ת�� Is Null And a.��¼���� = 1 And a.�Ǽ�ʱ�� < d_End
         Union All
         Select Distinct a.����id --3.���￨���շѽ����¼(�ų�֮���˿��ѵ�,һ�ŵ�����ֻҪ����һ������) 
         From סԺ���ü�¼ A
         Where (a.��¼״̬ In (1, 2) Or a.��¼״̬ = 3 And Not Exists
                (Select 1
                 From סԺ���ü�¼ B
                 Where a.No = b.No And a.��¼���� = b.��¼���� And b.��¼״̬ = 2 And b.�Ǽ�ʱ�� >= d_End))
    And a.��ת�� Is Null And a.���ʷ��� = 0 And a.��¼���� = 5 And a.�Ǽ�ʱ�� < d_End
         Union All --4.����(���ʵ�)��סԺ�Ľ��ʽ����¼ 
         Select ����id
         From (With Settle As (Select Distinct a.Id As ����id, a.����id --3.����(���ʵ�)��סԺ�Ľ��ʽ����¼(�ų�֮��������ϵ�) 
                               From ���˽��ʼ�¼ A
                               Where (a.��¼״̬ In (1, 2) Or a.��¼״̬ = 3 And Not Exists
                                      (Select 1 From ���˽��ʼ�¼ B Where a.No = b.No And b.��¼״̬ = 2 And b.�շ�ʱ�� >= d_End))
              And a.��ת�� Is Null And a.�շ�ʱ�� < d_End)
                Select ����id
                From Settle
                Minus
                --���½���IDҪ�����ų�,���ⲿ�ַ�����ϸ��ת����Ӱ������ļ����Ƿ���� 
                --1.һ��Ԥ�����ʽ��ʳ��꣨����ID��ͬ��
                --2.���õ��ݵĽ���ID��صĿ��ܻ�������NO����������ID(�������Ϻ�ֶ�ν��ʽ��壬���ܲ�����ת��ʱ��֮��)
                --���ǵ�������ĸ����ԣ�Ϊ���߼���������ѯ���ܣ�������ID���ų� 
                Select Distinct d.Id
                From ���˽��ʼ�¼ D,
                     (Select Distinct c.����id --���סԺ����һ��ᣬ�Լ�������ʺ�סԺ���ʿ���һ����ҳ�ͬһ��Ԥ�����������ﲻ����ҳID 
                       From סԺ���ü�¼ C,
                            (Select Distinct d.No, d.���, Mod(d.��¼����, 10) As ��¼����
                              From סԺ���ü�¼ D,
                                   (Select s.����id From Settle S, ���˽��ʼ�¼ E --û�н����Ҹò���֮��û���ٽ���ͳ��˴��ʣ����־Ͳ��ų� 
                                     Where s.����id = e.����id And (e.�շ�ʱ�� > d_End Or Exists (Select 1 From ��Ժ���� F Where s.����id = f.����id))) S 
                              Where d.����id = s.����id) D
                       Where c.No = d.No And Mod(c.��¼����, 10) = d.��¼���� And c.��� = d.��� --���ʺ����Ϻ��ٶ԰������ʵ����ʵĽ���IDΪ�յļ�¼,һ����ܼ����Ƿ����,���ֽ���IDΪ�յ�����ת���ں��浥��ת�� 
                       Group By c.No, Mod(c.��¼����, 10), c.����id --һ�ŵ����е�һ�пɲ��ֽ��ʣ��Ե���Ϊ�������жϣ�����һ�ŵ��ݵ�����һ���ֱ�ת�� 
                       Having Nvl(Sum(c.ʵ�ս��), 0) <> Nvl(Sum(c.���ʽ��), 0) Or Exists (Select 1 --�ų�ת��ʱ��֮���ٴν��ʵ�(���Ϻ��ٴν���)������ԭʼ����ת�ߺ󣬺�������ʱ�޷���ȷ�ж� 
                                                                                   From סԺ���ü�¼ E, ���˽��ʼ�¼ S
                                                                                   Where e.No = c.No And Mod(e.��¼����, 10) = Mod(c.��¼����, 10) And
                                                                                         e.��¼���� In (12, 13, 15) And e.����id = s.Id  And s.��ת�� Is Null And s.�շ�ʱ�� >= d_End)
                       Union All
                       Select Distinct c.����id
                       From ������ü�¼ C,
                            (Select Distinct d.No, d.���, Mod(d.��¼����, 10) As ��¼����
                              From ������ü�¼ D, Settle S
                              Where d.����id = s.����id) D --��Ϊ�����ﲡ�ˣ����ԣ�ֻҪû�н���,�ò��˵Ķ���ת�� 
                       Where c.No = d.No And Mod(c.��¼����, 10) = d.��¼���� And c.��� = d.���
                       Group By c.No, Mod(c.��¼����, 10), c.����id
                       Having Nvl(Sum(c.ʵ�ս��), 0) <> Nvl(Sum(c.���ʽ��), 0) Or Exists (Select 1
                                                                                   From ������ü�¼ E, ���˽��ʼ�¼ S
                                                                                   Where e.No = c.No And Mod(e.��¼����, 10) = Mod(c.��¼����, 10) And
                                                                                         e.��¼���� In (12, 13, 15) And e.����id = s.Id And s.��ת�� Is Null And s.�շ�ʱ�� >= d_End)) N
                Where d.����id = n.����id)
         );

  --�ų�Ԥ����δ�����
  --Ϊ�˽����߼��ĸ����ԣ����ų���ת��ʱ��֮��ҩ��δ��ҩ�ķ��ü�¼��Ӧ�Ľ���ID������������Ľ������ݺͷ�������ǿ��ת�� 
  --��Ϊǰ���SQL����Ľ���ID���ܲ�ȫ�ǳ�Ԥ����(�����շѺ�סԺ���ʲ��ѵ�)�����ԣ���Ҫ����һ��SQL���ų� 
  --���ڿ��ܴ��������쳣(סԺ���ý��ʳ�Ԥ�����Ϊ1������Ԥ��)������û�м�Ԥ����������޶� 
  Update /*+ rule*/ ����Ԥ����¼
  Set ��ת�� = Null
  Where ��ת�� = n_���� And
        ����id In (Select Distinct d.����id
                 From ����Ԥ����¼ D,
                      --����D����Ϊ�˲��ͬһԤ�����ݵ���������ID����Ԥ�����Ԥ�����ϵģ��ٴγ�ͬһԤ�����ݣ� 
                      --��Ԥ�����Ԥ�������漰�����н���ID�Ķ���ת�������ⲿ�ֳ�Ԥ���Ľ���ID���ų���ԭʼԤ������ת�ߣ�������������ID�����õ��ݵ�һ����(ԭʼ���ʡ��������ϡ��ٴν�һ���֡��ٴν�ȫ��)ת�� 
                      (Select Distinct l.No
                        From ����Ԥ����¼ L, ����Ԥ����¼ P --���ܱ��ν��ʳ��ֻ��ʣ��������Ҫ����L����ԭʼ��Ԥ���ĵ��ݣ��Լ���¼����Ϊ11�Ŀ��ܻ���ת��ʱ��֮��������ʣ���Ľ���ID 
                        Where l.��¼���� = p.��¼���� And l.No = p.No And p.��¼���� In (1, 11) And p.��ת�� = n_����
                        Group By l.No, l.����id
                        Having Nvl(Sum(l.���), 0) <> Nvl(Sum(l.��Ԥ��), 0) And (Exists (Select 1
                                                                                  From ����Ԥ����¼ E --û�г�����֮��û���ٳ���������ͳ��˴��ʣ��Լ������ø��Ľ��ʲ�������ʾ��Ԥ�����ɳ��������������־Ͳ��ų�
                                                                                  Where l.����id = e.����id And e.��ת�� Is Null And e.�տ�ʱ�� > d_End)
                                                                                  Or Exists (Select 1 From ��Ժ���� E Where l.����id =e.����id)
                                                                                  Or Exists (Select 1 From ����δ����� E Where l.����id =e.����id))  
                        Or Nvl(Sum(l.���), 0) = Nvl(Sum(l.��Ԥ��), 0) And Exists (Select 1
                                                                                  From ����Ԥ����¼ E --�ų�ת��ʱ��֮�����������ID���,10.34.20�󣬳�Ԥ��ȫ������������һ����¼���շ�ʱ����ǳ�Ԥ��ʱ��(��ǰ����ԭʼ��Ԥ����ļ�¼�����Ԥ���ֶΣ�����ֱ�Ӳ鵽��Ԥ�����ʱ��)
                                                                                  Where e.No = l.No And e.��¼���� = 11 And e.��ת�� Is Null And e.�տ�ʱ�� >= d_End)) N
                 Where d.No = n.No And d.��¼���� In (1, 11));

  --Ԥ����û��ʹ�þ�ֱ�����˵ļ�¼(����IDΪ��) 
  Update /*+ rule*/ ����Ԥ����¼
  Set ��ת�� = n_����
  Where ��¼���� = 1 And
        NO In (Select a.No
               From ����Ԥ����¼ A
               Where a.����id Is Null And a.��¼���� = 1 And a.��¼״̬ In (2, 3) And a.��ת�� Is Null And a.�տ�ʱ�� < d_End
               Group By a.No
               Having Sum(a.���) = 0);

  --��Ԥ�������ϵļ�¼����¼����Ϊ2����û�н���ID 
  Update /*+ rule*/ ����Ԥ����¼
  Set ��ת�� = n_����
  Where ����id Is Null And ��¼���� = 2 And NO In (Select a.No From ����Ԥ����¼ A Where a.��ת�� = n_���� And a.��¼���� = 3);

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
  Where (��¼id,����ID) In (Select a.Id,A.����ID From ����Ԥ����¼ A Where ��ת�� = n_����);

  --1.�ҺŴ��ۺ�ʵ�ս��Ϊ0��(û�ж�Ӧ��Ԥ����¼),��ʹ֮�����˺ŷ���Ҳ���ܣ���Ϊ���Ϊ�㲻Ӱ�����),�����Ѽ�ʹΪ��Ҳ��Ԥ����¼ 
  --����IDΪ�յ����쳣���ݣ�����ҽԺ����3�ʴ������ݣ�
  --���ݹҺż�¼����������ã���ֱ�Ӱ�ʱ����������Ҫ�� 
  Update /*+ rule*/ ������ü�¼
  Set ��ת�� = n_����
  Where NO In (Select NO From ���˹Һż�¼ Where ��ת�� Is Null And �Ǽ�ʱ�� < d_End) And ��¼���� = 4 And (ʵ�ս�� = 0 Or ����id Is Null);

  --2.ֱ���շѵĺͽ����޽��㣨Ԥ������¼�ģ�Union����allȥ���ظ��Լ���in������ 
  Update /*+ rule*/ ������ü�¼
  Set ��ת�� = n_����
  Where ����id In
        (Select ����id From ����Ԥ����¼ Where ��ת�� = n_���� Union Select ID From ���˽��ʼ�¼ Where ��ת�� = n_����);

  --3.û�н���id������(���Ǽ�ʱ��)
  --1)δ���ʵ�������ʷ���(����)���ò���û��Ԥ����¼���Ԥ����¼�����Ҹ�ʱ��֮����������÷���
  --2)δ���ʵĻ��ۼ�¼
  --3)δ�շѣ�Ҳû�г�Ԥ�����������
  --������"��ת�� Is Null"��Ϊ�˴���������α��ת������� 
  Update /*+ rule*/ ������ü�¼ A
  Set ��ת�� = n_����
  Where (Not Exists (Select 1 From ����Ԥ����¼ B Where a.����id = b.����id And b.��ת�� Is Null And ��¼���� In (1, 11)) And Not Exists
         (Select 1 From ������ü�¼ B Where a.����id = b.����id And b.��ת�� Is Null And �Ǽ�ʱ�� > d_End) And ��¼���� = 2 Or ��¼״̬ = 0 Or
         ��¼���� = 1 And ʵ�ս�� = 0 And ���ʽ�� = 0) And ����id Is Null And ��ת�� Is Null And �Ǽ�ʱ�� < d_End;

  --4.û�н���id������(������ʱ��)
  --���������ļ��ʼ�¼����¼״̬Ϊ2�����Ǽ�ʱ������ڵ�ǰָ��ת��ʱ��֮�󣬶�ԭʼ���ʼ�¼����¼״̬Ϊ3�����Ǽ�ʱ����ָ��ת��ʱ��֮ǰ��ǰ�����ߵķ���ʱ������ͬ�ġ�
  --1)δ���ʵ�����ʷ��û���ۺ�ʵ�ս��Ϊ��ģ�����ģ�����û�й�ѡ������ý��ʣ�
  --2)�������Ϻ󣬼��ʵ����ʵļ�¼������IDΪ���Ҽ�¼״̬Ϊ2�ģ�����¼״̬Ϊ3�����н���ID������ǰ����ת��. 
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

  --5.�н���id�������(������ʱ��)
  --���ѱ���ۺ���ʽ��Ϊ����շѼ�¼,����һ�ŵ�����ͬ����ID�Ľ��ʽ��֮��Ϊ0(������Ϊ��)
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
  --���������ļ��ʼ�¼����¼״̬Ϊ2�����Ǽ�ʱ������ڵ�ǰָ��ת��ʱ��֮�󣬶�ԭʼ���ʼ�¼����¼״̬Ϊ3�����Ǽ�ʱ����ָ��ת��ʱ��֮ǰ��ǰ�����ߵķ���ʱ������ͬ�ġ�
  --1)ת���������Ϻ󣬼��ʵ����ʵļ�¼������״̬Ϊ2��û�н���ID��(��¼״̬Ϊ3���н���ID��)����ǰ����ת���� 
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
           Having Nvl(Sum(b.ʵ�ս��), 0) = 0)) And ��¼���� = 2 Or ��¼״̬ = 0) And ����id Is Null And ��ת�� Is Null And ����ʱ�� < d_End;

  --3.��Ժδ���ʵģ����ʲ��ˣ�����Ϊ�Ǻܾ���ǰ����Щ���ݣ����Ԥ���ѳ��꣬����ΪҪת�� 
  Update /*+ rule*/ סԺ���ü�¼ A
  Set ��ת�� = n_����
  Where ��ת�� Is Null And ����id Is Null And
        (����id, ��ҳid) In (Select ����id, ��ҳid
                         From ������ҳ C
                         Where ��Ժ���� < d_End And ��ת�� Is Null And ����ת�� Is Null And Not Exists
                          (Select 1
                                From ����Ԥ����¼ B
                                Where b.����id = c.����id And b.��ת�� Is Null And b.Ԥ����� = 2 And b.��¼���� In (1, 11)
                                Having Nvl(Sum(b.���), 0) - Nvl(Sum(b.��Ԥ��), 0) <> 0));

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
  Where Not Exists
   (Select 1 From Ʊ��ʹ����ϸ B Where b.����id = a.Id And b.ʹ��ʱ�� >= d_End) And ��ת�� Is Null And ʣ������ = 0 And �Ǽ�ʱ�� < d_End;

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
  --��ת�����������Һŷ���δת���ģ�ת��ʱ��֮�����ҽ����ҽ����Ӧ�ķ���δת���� 
  --��ʹ���ھ���(r.ִ��״̬ <> 2 )��Ҳǿ��ת�� 
  Update /*+ rule*/ ���˹Һż�¼ T
  Set ��ת�� = n_����
  Where Rowid In
        (Select Rowid
         From ���˹Һż�¼ R
         Where Not Exists (Select 1
                From ������ü�¼ A
                Where r.No = a.No And a.�Ǽ�ʱ�� < d_End And a.��¼���� = 4 And a.��ת�� Is Null) And Not Exists
          (Select 1
                From ����ҽ����¼ A
                Where a.�Һŵ� = r.No And a.��ת�� Is Null And a.������Դ <> 4 And Nvl(a.ͣ��ʱ��, a.����ʱ��) >= d_End) And Not Exists
          (Select 1
                From ������ü�¼ E, ����ҽ����¼ A
                Where r.No = a.�Һŵ� And a.Id = e.ҽ����� And a.������Դ <> 4 And e.��ת�� Is Null) And r.��ת�� Is Null And
               r.�Ǽ�ʱ�� < d_End);

  --������һ���ֹҺ�����δת�������ԣ����ܱ�����ݿ�����Һ����ݲ�ƥ�� 
  Update ���˹ҺŻ��� Set ��ת�� = n_���� Where ��ת�� Is Null And ���� < d_End;
  Update /*+ rule*/ ����ת���¼ Set ��ת�� = n_���� Where NO In (Select NO From ���˹Һż�¼ Where ��ת�� = n_����);

  --ͨ��"סԺ���ü�¼"����ѯ��������"���˽��ʼ�¼",��Ϊ��Ժδ������ʲ���Ҳת���˷��� 
  --��Ժ����������Ȼ��Ҫ����Ϊ����ĳ�ν���ת���ˣ������˵�ʱ��δ��Ժ(һ��סԺ��ν���)�� 
  --ͨ��ָ��������ʽ���������Ż���ȱʡ����"������ҳIX_��Ժ����"������Ч��̫�ͣ� 
  Update /*+ rule*/ ������ҳ P
  Set ��ת�� = n_����
  Where Not Exists (Select 1 From סԺ���ü�¼ A Where a.����id = p.����id And a.��ҳid = p.��ҳid And a.��ת�� Is Null) And ��ת�� Is Null And
        ����ת�� Is Null And ��Ժ���� < d_End And
        (����id, ��ҳid) In (Select Distinct ����id, ��ҳid From סԺ���ü�¼ Where ��ת�� = n_����);

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
  Where ������Դ <> 4 And (����id, ��ҳid) In (Select ����id, ID
                                       From ���˹Һż�¼
                                       Where ��ת�� = n_����
                                       Union All
                                       Select ����id, ��ҳid
                                       From ������ҳ
                                       Where ��ת�� = n_����);

  --�ԵǼ��ಡ��(�޹Һŵ���) 
  --����ID�����ظ�����Ϊ���鱨��֮��ģ���ι�����������һ�ű��棬���ڲ���ҽ��������У����ҽ��id��Ӧͬһ����ID 
  --Ϊ�������ܣ�����ҽ�����ͼ�¼�ķ���ʱ���ѯ�������þ�ȷ��ʱ�䣬��Ϊֱ�ӵǼǵļ���ҽ����һ�㿪��ʱ���뷢��ʱ������
  Update /*+ rule*/ ���Ӳ�����¼
  Set ��ת�� = N_����
  Where ID In (Select C.����id
             From ����ҽ����¼ B, ����ҽ������ C
             Where C.ҽ��id = B.Id And Nvl(B.��ҳid, 0) = 0 And B.�Һŵ� Is Null And B.���id Is Null And B.��ת�� Is Null And
                   B.����ʱ�� < d_End);

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
  Where ����id In (Select ID From ���Ӳ�����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ Ӱ�񱨸沵��
  Set ��ת�� = n_����
  Where (ҽ��id, ����id) In (Select ҽ��id, ����id From ����ҽ������ Where ��ת�� = n_����);

  Update /*+ rule*/ ������ļ�¼
  Set ��ת�� = n_����
  Where ����id In (Select ID From ���Ӳ�����¼ Where ��ת�� = n_����);
  Update /*+ rule*/ �����걨��¼
  Set ��ת�� = n_����  
  Where �ļ�id In (Select ID From ���Ӳ�����¼ Where ��ת�� = n_����);
  
  Update /*+ rule*/ �������淴��
  Set ��ת�� = n_����
  Where �ļ�id In (Select ID From ���Ӳ�����¼ Where ��ת�� = n_����);
  
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
  Update /*+ rule*/ ����ҽ����¼
  Set ��ת�� = n_����
  Where �Һŵ� In (Select NO From ���˹Һż�¼ Where ��ת�� = n_����) And ������Դ <> 4;
  Update /*+ rule*/ ����ҽ����¼
  Set ��ת�� = n_����
  Where (����id, ��ҳid) In (Select ����id, ��ҳid From ������ҳ Where ��ת�� = n_����);

  --�ԵǼ��ಡ��(�޹Һŵ�)������ҽ��������ǰ��ת����ʱ��ת�� 
  Update /*+ rule*/ ����ҽ����¼
  Set ��ת�� = n_����
  Where Rowid In (Select b.Rowid
                  From ����ҽ����¼ B, ����ҽ������ C
                  Where (b.���id = c.ҽ��id Or b.Id = c.ҽ��id) And c.��ת�� = n_����);

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
  
  Update /*+ rule*/ RIS���ԤԼ
  Set ��ת�� = n_����
  Where ҽ��id In (Select ID From ����ҽ����¼ Where ��ת�� = n_����);

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

--93964:���Ʊ�,2016-03-08,�䶯��¼����
Create Or Replace Procedure Zl_����ҽ����¼_ֹͣ
(
  --���ܣ�ָֹͣ����ҽ��
  --˵����һ����ҩ��ֻ�ܵ���һ��
  --������ID_IN=���IDΪNULL��ҽ����ID(��ҩ;��,��ҩ�÷�,�����Ŀ,��Ҫ����,������ҽ��)
  --      �ڲ�����_IN=�Ƿ����������ڲ��ڵ��ã���Ҫ�����Ƿ��ֹ�����ֹͣ����ȼ�
  Id_In         ����ҽ����¼.Id%Type,
  ��ֹʱ��_In   ����ҽ����¼.ִ����ֹʱ��%Type,
  ͣ��ҽ��_In   ����ҽ����¼.ͣ��ҽ��%Type,
  �ڲ�����_In   Number := 0,
  ҽʦ�ʸ�_In   Number := 0,
  ͣ�����_In   Number := 0,
  ����Ա���_In ��Ա��.���%Type := Null,
  ����Ա����_In ��Ա��.����%Type := Null
) Is
  v_״̬       ����ҽ����¼.ҽ��״̬%Type;
  v_ҽ������   ����ҽ����¼.ҽ������%Type;
  v_����ȼ�id ������ҳ.����ȼ�id%Type;
  v_����id     ������ҳ.����id%Type;
  v_��ҳid     ������ҳ.��ҳid%Type;
  v_Ӥ��       ����ҽ����¼.Ӥ��%Type;
  v_�������   ����ҽ����¼.�������%Type;
  v_��������   ������ĿĿ¼.��������%Type;

  Cursor c_Oldinfo Is
    Select b.*
    From (Select c.*
           From ���˱䶯��¼ C
           Where c.����id = v_����id And c.��ҳid = v_��ҳid And
                 c.��ʼʱ�� = (Select Min(��ʼʱ��)
                           From ���˱䶯��¼
                           Where ����id = v_����id And ��ҳid = v_��ҳid And ��ʼʱ�� > ��ֹʱ��_In) And
                 c.��ֹʱ�� = (Select Min(��ֹʱ��)
                           From ���˱䶯��¼
                           Where ����id = v_����id And ��ҳid = v_��ҳid And ��ʼʱ�� > ��ֹʱ��_In)) A, ���˱䶯��¼ B
    
    Where b.����id = v_����id And b.��ҳid = v_��ҳid And a.��ʼʱ�� = b.��ֹʱ�� And a.��ʼԭ�� = b.��ֹԭ�� And a.���Ӵ�λ = b.���Ӵ�λ
    Union
    Select *
    From ���˱䶯��¼
    Where ����id = v_����id And ��ҳid = v_��ҳid And ��ֹʱ�� Is Null And ��ʼʱ�� <= ��ֹʱ��_In;

  Cursor c_Endinfo Is
    Select * From ���˱䶯��¼ Where ����id = v_����id And ��ҳid = v_��ҳid And ��ֹʱ�� Is Null;
  r_Oldinfo  c_Oldinfo%Rowtype;
  r_Endinfo  c_Endinfo%Rowtype;
  v_��ֹԭ�� ���˱䶯��¼.��ֹԭ��%Type;
  v_��ֹʱ�� ���˱䶯��¼.��ֹʱ��%Type;
  v_��ֹ��Ա ���˱䶯��¼.��ֹ��Ա%Type;

  v_Date     Date;
  v_Count    Number;
  v_Temp     Varchar2(255);
  v_��Ա��� ���˱䶯��¼.����Ա���%Type;
  v_��Ա���� ���˱䶯��¼.����Ա����%Type;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --���ҽ��״̬�Ƿ���ȷ:��������
  Select a.ҽ��״̬, a.ҽ������, a.����id, a.��ҳid, Nvl(a.Ӥ��, 0), Nvl(a.�������, '*') As �������, b.��������
  Into v_״̬, v_ҽ������, v_����id, v_��ҳid, v_Ӥ��, v_�������, v_��������
  From ����ҽ����¼ A, ������ĿĿ¼ B
  Where a.������Ŀid = b.Id(+) And a.Id = Id_In;
  If v_״̬ In (4, 8, 9) Then
    v_Error := 'ҽ��"' || v_ҽ������ || '"�Ѿ������ϻ�ֹͣ��������ֹͣ��';
    Raise Err_Custom;
  End If;

  --����Ƿ�����Һ��Һ��¼�����Ƿ��Ѿ�����
  Select Count(1) Into v_Count From ��Һ��ҩ��¼ Where �Ƿ����� = 1 And ҽ��ID = Id_In And ִ��ʱ�� > ��ֹʱ��_In;
  If v_Count > 0 Then
    v_Error := 'ҽ��"' || v_ҽ������ || '"����ҺҩƷ���Ѿ�����Һ������������������ֹͣ��';
    Raise Err_Custom;
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

  Select Sysdate Into v_Date From Dual;

  --�ж��Ƿ����ִҵҽʦ������ҽʦ���ʸ�Ͳ����Ƿ�ѡ
  If ҽʦ�ʸ�_In > 0 Or ͣ�����_In = 0 Then
    Update ����ҽ����¼
    Set ҽ��״̬ = 8, ִ����ֹʱ�� = ��ֹʱ��_In, ͣ��ҽ�� = ͣ��ҽ��_In, ͣ��ʱ�� = v_Date, ��˱�� = Decode(��˱��, 2, 3, ��˱��)
    Where ID = Id_In Or ���id = Id_In;
  
    Insert Into ����ҽ��״̬
      (ҽ��id, ��������, ������Ա, ����ʱ��)
      Select ID, 8, v_��Ա����, v_Date --��ʿͣʱ��¼Ϊ��ʿ
      From ����ҽ����¼
      Where ID = Id_In Or ���id = Id_In;
  Else
    --����ֻ�޸���˱�ǣ������µ�״̬
    Update ����ҽ����¼ Set ��˱�� = 2 Where ID = Id_In Or ���id = Id_In;
  
    Insert Into ����ҽ��״̬
      (ҽ��id, ��������, ������Ա, ����ʱ��, ����˵��)
      Select ID, 13, v_��Ա����, v_Date, To_Char(��ֹʱ��_In, 'YYYY-MM-DD HH24:MI:SS') --��ʿͣʱ��¼Ϊ��ʿ
      From ����ҽ����¼
      Where ID = Id_In Or ���id = Id_In;
  End If;

  --�������⴦��
  If Nvl(�ڲ�����_In, 0) = 0 And (ҽʦ�ʸ�_In > 0 Or ͣ�����_In = 0) Then
    --ֹͣ����ҽ��ʱ���䶯���˲���
    If v_������� = 'Z' And v_�������� In ('9', '10') Then
      Open c_Oldinfo; --�����ڴ���֮ǰ�ȴ�
      Fetch c_Oldinfo
        Into r_Oldinfo;
      Open c_Endinfo;
      Fetch c_Endinfo
        Into r_Endinfo;
      If c_Endinfo%Rowcount = 0 Then
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
    
      Update ������ҳ Set ��ǰ���� = 'һ��' Where ����id = v_����id And ��ҳid = v_��ҳid;
    
      --ȡ���ϴα䶯
      If r_Oldinfo.��ֹʱ�� Is Not Null Then
        v_��ֹʱ�� := r_Oldinfo.��ֹʱ��;
        v_��ֹԭ�� := r_Oldinfo.��ֹԭ��;
        v_��ֹ��Ա := r_Oldinfo.��ֹ��Ա;
        --ȡ���ϴα䶯
        Update ���˱䶯��¼
        Set ��ֹʱ�� = ��ֹʱ��_In, ��ֹԭ�� = 13, ��ֹ��Ա = v_��Ա����, �ϴμ���ʱ�� = Null
        Where ����id = v_����id And ��ҳid = v_��ҳid And ��ֹʱ�� = v_��ֹʱ�� And ��ֹԭ�� = v_��ֹԭ��;
        --���½����ļ�¼�����ֹͣ����������ɾ���ϴμ���ʱ��
        Update ���˱䶯��¼
        Set ���� = 'һ��', �ϴμ���ʱ�� = Null
        Where ����id = v_����id And ��ҳid = v_��ҳid And ��ʼʱ�� > ��ֹʱ��_In;
      Else
        Update ���˱䶯��¼
        Set ��ֹʱ�� = ��ֹʱ��_In, ��ֹԭ�� = 13, ��ֹ��Ա = v_��Ա����
        Where ����id = v_����id And ��ҳid = v_��ҳid And ��ֹʱ�� Is Null;
      End If;
    
      While c_Oldinfo%Found Loop
        Insert Into ���˱䶯��¼
          (ID, ����id, ��ҳid, ��ʼʱ��, ��ʼԭ��, ���Ӵ�λ, ����id, ����id, ����ȼ�id, ��λ�ȼ�id, ����, ���λ�ʿ, ����ҽʦ, ����ҽʦ, ����ҽʦ, ����, ����Ա���, ����Ա����,
           ��ֹʱ��, ��ֹԭ��, ��ֹ��Ա)
        Values
          (���˱䶯��¼_Id.Nextval, v_����id, v_��ҳid, ��ֹʱ��_In, 13, r_Oldinfo.���Ӵ�λ, r_Oldinfo.����id, r_Oldinfo.����id,
           r_Oldinfo.����ȼ�id, r_Oldinfo.��λ�ȼ�id, r_Oldinfo.����, r_Oldinfo.���λ�ʿ, r_Oldinfo.����ҽʦ, r_Oldinfo.����ҽʦ,
           r_Oldinfo.����ҽʦ, 'һ��', v_��Ա���, v_��Ա����, v_��ֹʱ��, v_��ֹԭ��, v_��ֹ��Ա);
      
        Fetch c_Oldinfo
          Into r_Oldinfo;
      End Loop;
    
      Close c_Oldinfo;
      Close c_Endinfo;
    Elsif v_������� = 'H' And v_�������� = '1' And v_Ӥ�� = 0 Then
      --ֹͣ����ȼ�ʱ��ͬʱȡ�����˵Ļ���ȼ�����
      Begin
        Select c.�շ�ϸĿid
        Into v_����ȼ�id
        From ����ҽ����¼ A, ����ҽ���Ƽ� C, �շ���ĿĿ¼ D
        Where a.Id = c.ҽ��id And c.�շ�ϸĿid = d.Id And d.��� = 'H' And Nvl(d.��Ŀ����, 0) <> 0 And a.Id = Id_In And Rownum = 1 And
              Exists
         (Select 1 From ������ҳ Where ����id = v_����id And ��ҳid = v_��ҳid And ����ȼ�id = c.�շ�ϸĿid);
      Exception
        When Others Then
          Null;
      End;
      If v_����ȼ�id Is Not Null Then
        --�䶯��¼��ʱ������룬�Ա���˲���ʱ����ͬһ���ֵ�У�ԡ�ֹͣ�Ȳ���
        v_Date := To_Date(To_Char(��ֹʱ��_In, 'yyyy-mm-dd hh24:mi') || To_Char(v_Date, 'ss'), 'yyyy-mm-dd hh24:mi:ss');
        Zl_���˱䶯��¼_Nurse(v_����id, v_��ҳid, Null, v_Date, v_��Ա���, v_��Ա����);
      End If;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����ҽ����¼_ֹͣ;
/

--93964:���Ʊ�,2016-03-08,�䶯��¼��������
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
                 c.��ֹʱ�� = (Select Min(��ֹʱ��)
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
          Nvl(a.ҽ����Ч, 0) = 0 And a.ҽ��״̬ Not In (1, 2, 4, 8, 9) And a.��ʼִ��ʱ�� < v_Stoptime
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
  v_����ֵ   Zlparameters.����ֵ%Type;
  v_Count    Number;
  v_Date     Date;
  v_Temp     Varchar2(255);
  v_��Ա��� ��Ա��.���%Type;
  v_��Ա���� ��Ա��.����%Type;

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
    Select zl_GetSysParameter(25) Into v_����ֵ From Dual;
    If Nvl(v_����ֵ, '0') <> '0' Then
      Select zl_GetSysParameter(26) Into v_����ֵ From Dual;
      If v_ǰ��id Is Null Then
        Select Count(*) Into v_Count From ����ǩ�����ò��� Where ����id = v_��������id And ���� = 1;
      Else
        Select Count(*) Into v_Count From ����ǩ�����ò��� Where ����id = v_��������id And ���� = 3;
      End If;
      If Nvl(Substr(v_����ֵ, 2, 1), '0') = '1' And v_ǰ��id Is Null And v_Count > 0 Or
         Nvl(Substr(v_����ֵ, 3, 1), '0') = '1' And v_ǰ��id Is Not Null And v_Count > 0 Then
        Select Nvl(Max(�Ƿ�ͣ��), 0)
        Into v_Count
        From (Select a.�Ƿ�ͣ��, a.ע��ʱ��
               From ��Ա֤���¼ A, ��Ա�� B
               Where a.��Աid = b.Id And b.���� = v_����ҽ��
               Order By a.ע��ʱ�� Desc)
        Where Rownum < 2;
        If v_Count = 0 Then
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
      End If;
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
        Set ��ֹʱ�� = v_��ʼʱ��, ��ֹԭ�� = 13, ��ֹ��Ա = v_��Ա����
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

--93964:���Ʊ�,2016-03-07,�����䶯��¼ֹͣ��������
Create Or Replace Procedure Zl_����ҽ����¼_����
(
  ҽ��id_In     ����ҽ����¼.Id%Type,
  Flag_In       Number := 0,
  ҽ������_In   ����ҽ����¼.ҽ������%Type := Null,
  ��������_In   ����ҽ��״̬.��������%Type := Null,
  ����Ա���_In ��Ա��.���%Type := Null,
  ����Ա����_In ��Ա��.����%Type := Null
  --���ܣ�����סԺҽ����״̬�������Ͳ���(������������ͨ������Zl_����ҽ����¼_��������������)
  --������ҽ��ID_IN=һ��ҽ��ID
  --      FLAG_IN=�������ݡ�����ֹͣ��0=���ִ����ֹʱ��,1=�������е�ִ����ֹʱ�䡣
  --      ҽ������_IN=�ù��̱��������˵���ʱ���ã����ڴ�����ʾ��
  --      ��������_IN=�ù��̱��������˵���ʱ���ã����ں˶Ի������ݡ�0-���˷���,n=���˾���ҽ������
) Is
  --����ָ��ҽ���Ĳ�����¼,��һ��ΪҪ���˵�����(״̬��������)
  --���������˷��ͺ���Զ�ֹͣ,�ڻ��˷���ʱ�Զ�����ֹͣ����
  Cursor c_Rolladvice Is
    Select b.������Ա, b.����ʱ��, 0 As ���ͺ�, b.��������, 0 As ִ��״̬, Sysdate + Null As �״�ʱ��, Sysdate + Null As ĩ��ʱ��, a.�ϴ�ִ��ʱ��, a.ҽ����Ч,
           a.������� As ���, a.������Ŀid, Null As ����, a.����id, a.��ҳid, a.Ӥ��, 0 As ��¼����, 0 As �������, 0 As ��������id, a.��˱��, a.����ҽ��,
           a.ִ�п���id
    From ����ҽ����¼ A, ����ҽ��״̬ B
    Where a.Id = b.ҽ��id And (a.Id = ҽ��id_In Or a.���id = ҽ��id_In) And
          (Nvl(a.ҽ����Ч, 0) = 0 And b.�������� Not In (1, 2, 3) Or Nvl(a.ҽ����Ч, 0) = 1 And b.�������� Not In (1, 2, 3, 8))
    Union
    Select b.������ As ������Ա, b.����ʱ�� As ����ʱ��, b.���ͺ�, -null As ��������, b.ִ��״̬, b.�״�ʱ��, b.ĩ��ʱ��, a.�ϴ�ִ��ʱ��, a.ҽ����Ч, c.���, a.������Ŀid,
           c.�������� As ����, a.����id, a.��ҳid, a.Ӥ��, b.��¼����, b.�������, a.��������id, a.��˱��, a.����ҽ��, a.ִ�п���id
    From ����ҽ����¼ A, ����ҽ������ B, ������ĿĿ¼ C
    Where a.Id = b.ҽ��id And a.������Ŀid = c.Id And (a.Id = ҽ��id_In Or a.���id = ҽ��id_In)
    Order By ����ʱ�� Desc, ���ͺ�;
  r_Rolladvice c_Rolladvice%RowType;

  --��ʽͬc_Rolladvice��ֻȡ���Ͳ��������Զ����˴���
  Cursor c_Rollsend(v_���ͺ� ����ҽ������.���ͺ�%Type) Is
    Select Distinct b.ҽ��id, b.����ʱ�� As ����ʱ��, b.���ͺ�, b.ִ��״̬, a.������� As ���, c.��ǰ����id As ���˲���id, a.���˿���id,
                    b.ִ�в���id As ִ�п���id
    From ����ҽ����¼ A, ����ҽ������ B, ������ҳ C
    Where a.Id = b.ҽ��id And b.���ͺ� = v_���ͺ� And a.����id = c.����id And a.��ҳid = c.��ҳid And
          (a.Id = ҽ��id_In Or a.���id = ҽ��id_In)
    Order By b.����ʱ�� Desc, b.���ͺ�;

  --����ҽ��������NO������λ���Ҫ���ʵķ��ü�¼
  --һ��ҽ�������Ƕ���д�˷��ͼ�¼,�ҿ���NO��ͬ(ҩƷ��,�÷��巨��һ����)
  --���ܷ��ͼ�¼�ļƷ�״̬(��������Ʒ�),�з��ü�¼��Ȼ��������
  --����ֻ��۸񸸺�Ϊ�յ�,�Ա�ȡ�������
  --ֻ�ܼ�¼״̬Ϊ1�ķ���,���������ʻ򲿷����ʵļ�¼,���ٴ�������"��¼״̬=3"�Ķ�ȡ�����������жϣ�������
  Cursor c_Rollmoneyout
  (
    v_���ͺ�    ����ҽ������.���ͺ�%Type,
    v_ҽ��id    ����ҽ����¼.Id%Type,
    t_Adviceids t_Numlist
  ) Is
    Select /*+ Rule*/
     a.Id, a.��¼״̬, a.No, a.���, a.�շ����, a.ִ��״̬, d.��������, a.ִ�в���id, a.��¼����
    From ������ü�¼ A, Table(t_Adviceids) B, ����ҽ������ C, �������� D
    Where c.ҽ��id = b.Column_Value And c.���ͺ� = v_���ͺ� And a.ҽ����� = b.Column_Value And
          (a.ҽ����� = v_ҽ��id Or Nvl(v_ҽ��id, 0) = 0) And a.��¼״̬ In (0, 1, 3) And a.No = c.No And a.��¼���� = c.��¼���� And
          a.�۸񸸺� Is Null And a.�շ�ϸĿid = d.����id(+)
    Order By a.No, a.���;

  Cursor c_Rollmoneyin
  (
    v_���ͺ�    ����ҽ������.���ͺ�%Type,
    v_ҽ��id    ����ҽ����¼.Id%Type,
    t_Adviceids t_Numlist
  ) Is
    Select /*+ Rule*/
     a.Id, a.��¼״̬, a.No, a.���, a.�շ����, a.ִ��״̬, d.��������, a.ִ�в���id, a.��¼����
    From סԺ���ü�¼ A, Table(t_Adviceids) B, ����ҽ������ C, �������� D
    Where c.ҽ��id = b.Column_Value And c.���ͺ� = v_���ͺ� And a.ҽ����� = b.Column_Value And
          (a.ҽ����� = v_ҽ��id Or Nvl(v_ҽ��id, 0) = 0) And a.��¼״̬ In (0, 1, 3) And a.No = c.No And a.��¼���� = c.��¼���� And
          a.�۸񸸺� Is Null And a.�շ�ϸĿid = d.����id(+)
    Order By a.No, a.���;

  --ȡ����סԺ����ʱ�Զ����ŵ�����(��û�����ϵ�)
  Cursor c_Stuff_Drug(v_����id ҩƷ�շ���¼.����id%Type) Is
    Select ID
    From ҩƷ�շ���¼
    Where ����id = v_����id And (��¼״̬ = 1 Or Mod(��¼״̬, 3) = 0) And ����� Is Not Null
    Order By ҩƷid;

  --���ڴ�������ҽ���Ļ���
  Cursor c_Patilog
  (
    v_����id ���˱䶯��¼.����id%Type,
    v_��ҳid ���˱䶯��¼.��ҳid%Type
  ) Is
    Select *
    From ���˱䶯��¼
    Where ����id = v_����id And ��ҳid = v_��ҳid And ��ֹʱ�� Is Null
    Order By ��ʼʱ�� Desc;
  r_Patilog c_Patilog%RowType;

  Cursor c_Adviceids Is
    Select ID From ����ҽ����¼ Where ID = ҽ��id_In Or ���id = ҽ��id_In;
  t_Adviceids t_Numlist;

  v_ҽ��״̬     ����ҽ����¼.ҽ��״̬%Type;
  v_ҽ����Ч     ����ҽ����¼.ҽ����Ч%Type;
  v_����no       ����ҽ������.No%Type;
  v_�������     Varchar2(255);
  v_ĩ��ʱ��     ����ҽ������.ĩ��ʱ��%Type;
  v_����ʱ��     ����ҽ��״̬.����ʱ��%Type;
  v_��������     ������ĿĿ¼.��������%Type;
  v_ִ��Ƶ��     ������ĿĿ¼.ִ��Ƶ��%Type;
  v_�ϴ�ʱ��     ����ҽ����¼.�ϴ�ִ��ʱ��%Type;
  v_ִ��ʱ��     ����ҽ����¼.ִ��ʱ�䷽��%Type;
  v_��ʼִ��ʱ�� ����ҽ����¼.��ʼִ��ʱ��%Type;
  v_�ϴδ�ӡʱ�� ����ҽ����¼.�ϴδ�ӡʱ��%Type;
  v_Ƶ�ʼ��     ����ҽ����¼.Ƶ�ʼ��%Type;
  v_�����λ     ����ҽ����¼.�����λ%Type;
  v_���ͺ�       ����ҽ������.���ͺ�%Type;
  n_����ȼ�id   ���˱䶯��¼.����ȼ�id%Type;
  d_��ʼʱ��     ���˱䶯��¼.��ʼʱ��%Type;
  d_����ʱ��     ����ҽ��״̬.����ʱ��%Type;
  v_Tmp���ͺ�    ����ҽ������.���ͺ�%Type;
  n_ִ��         Number;

  Intdigit   Number(3);
  v_Update   Number(1);
  v_Count    Number(5);
  v_Temp     Varchar2(2000);
  v_��Ա��� ��Ա��.���%Type;
  v_��Ա���� ��Ա��.����%Type;
  v_Time     Varchar2(4000);
  n_Blndo    Number;

  v_Error Varchar2(2000);
  Err_Custom Exception;

  Function Checkmoneyundo
  (
    v_No       סԺ���ü�¼.No%Type,
    v_��¼���� סԺ���ü�¼.��¼����%Type,
    v_���     סԺ���ü�¼.���%Type,
    n_����     Number := 0 --0סԺ��1����
  ) Return Number Is
    n_Num      Number;
    n_ִ��״̬ Number;
  Begin
    n_Num := 0;
    If n_���� = 0 Then
      Select Nvl(Sum(Nvl(����, 1) * ����), 0) As ����
      Into n_Num
      From סԺ���ü�¼
      Where NO = v_No And ��¼���� = v_��¼���� And ��� = v_��� And ��¼״̬ In (2, 3);
      Select Nvl(ִ��״̬, 0)
      Into n_ִ��״̬
      From סԺ���ü�¼
      Where NO = v_No And ��¼���� = v_��¼���� And ��� = v_��� And ��¼״̬ = 3;
    Else
      Select Nvl(Sum(Nvl(����, 1) * ����), 0) As ����
      Into n_Num
      From ������ü�¼
      Where NO = v_No And ��¼���� = v_��¼���� And ��� = v_��� And ��¼״̬ In (2, 3);
      Select Nvl(ִ��״̬, 0)
      Into n_ִ��״̬
      From ������ü�¼
      Where NO = v_No And ��¼���� = v_��¼���� And ��� = v_��� And ��¼״̬ = 3;
    End If;
    If n_Num <> 0 Then
      n_Num := 1;
    End If;
    --�������¼����ִ�У�����ִ�еģ����Զ��ˡ�
    If n_ִ��״̬ <> 0 Then
      n_Num := 0;
    End If;
    Return(n_Num);
  End;
Begin
  v_Tmp���ͺ� := -1;
  Open c_Rolladvice;
  Loop
    Fetch c_Rolladvice
      Into r_Rolladvice;
    If c_Rolladvice%RowCount = 0 Then
      Close c_Rolladvice;
      v_Error := Nvl(ҽ������_In, '��ҽ��') || '��ǰû�п��Ի��˵����ݡ�';
      Raise Err_Custom;
    End If;
    Exit When c_Rolladvice%NotFound;
    Exit When d_����ʱ�� <> r_Rolladvice.����ʱ�� And d_����ʱ�� Is Not Null;
    d_����ʱ�� := r_Rolladvice.����ʱ��;
  
    --�������˵���ʱ�ж�
    If ҽ������_In Is Not Null Then
      If Nvl(r_Rolladvice.��������, 0) <> Nvl(��������_In, 0) Then
        v_Error := Nvl(ҽ������_In, '��ҽ��') || '�����뵱ǰҽ��һ����ˣ����ܸ�ҽ���Ѿ�ִ��������������';
        Raise Err_Custom;
      End If;
    End If;
  
    --һ�鷢�ͺ�ִֻ��һ��
    If v_Tmp���ͺ� <> r_Rolladvice.���ͺ� Then
      v_Tmp���ͺ� := r_Rolladvice.���ͺ�;
      n_ִ��      := 1;
    Else
      n_ִ�� := 0;
    End If;
  
    If n_ִ�� = 1 Then
      Open c_Adviceids;
      Fetch c_Adviceids Bulk Collect
        Into t_Adviceids;
      Close c_Adviceids;
    
      If r_Rolladvice.���ͺ� = 0 Then
        --����ҽ��״̬����(��ʱ��ؼ���)
        --4-���ϣ�5-������6-��ͣ��7-���ã�8-ֹͣ��9-ȷ��ֹͣ��10-Ƥ�Խ��;13-ͣ������
        ------------------------------------------------------------------
        --���ֻ���˻ص�У��״̬
        If r_Rolladvice.�������� = 3 Then
          v_Error := Nvl(ҽ������_In, '��ҽ��') || '��ǰ����ͨ��У��״̬�������ٻ��ˡ�';
          Raise Err_Custom;
        Elsif r_Rolladvice.�������� = 4 And Nvl(r_Rolladvice.Ӥ��, 0) = 0 Then
          If r_Rolladvice.��� = 'H' Then
            Select ��������, ִ��Ƶ�� Into v_��������, v_ִ��Ƶ�� From ������ĿĿ¼ Where ID = r_Rolladvice.������Ŀid;
            If v_�������� = '1' And v_ִ��Ƶ�� = '2' Then
              v_Error := '����ȼ����Ϻ����ٻ��ˡ�';
              Raise Err_Custom;
            End If;
          End If;
        End If;
      
        --����Ƿ�������������֮ǰ�Ĳ���
        If r_Rolladvice.�������� <> 5 Then
          --ȡ�������ʱ��
          Select Nvl(ҽ������ʱ��, To_Date('1900-01-01', 'YYYY-MM-DD'))
          Into v_����ʱ��
          From ������ҳ
          Where ����id = r_Rolladvice.����id And ��ҳid = r_Rolladvice.��ҳid;
        
          If r_Rolladvice.����ʱ�� < v_����ʱ�� Then
            v_Error := '�ò������������֮ǰ�Ĳ��������ٻ��ˡ�';
            Raise Err_Custom;
          End If;
        End If;
      
        --ɾ��(����ҽ��)�����״̬������¼
        Delete /*+ Rule*/
        From ����ҽ��״̬
        Where ҽ��id In (Select Column_Value From Table(t_Adviceids)) And ����ʱ�� = r_Rolladvice.����ʱ��;
      
        --ȡɾ����Ӧ�ָ���ҽ��״̬
        Select ��������
        Into v_ҽ��״̬
        From ����ҽ��״̬
        Where ����ʱ�� = (Select Max(����ʱ��) From ����ҽ��״̬ Where ҽ��id = ҽ��id_In) And ҽ��id = ҽ��id_In;
      
        --�ָ�(����ҽ��)���˺��״̬
        Update ����ҽ����¼ Set ҽ��״̬ = v_ҽ��״̬ Where ID = ҽ��id_In Or ���id = ҽ��id_In;
      
        --��������Ĵ���
        If r_Rolladvice.�������� = 8 Then
          --�����ڷ����ջع���ҽ�� ���������������ģʽ�����ж϶�Ӧ�ġ����˷������ʡ������Ƿ�ȡ��������������ˣ���������
          --                       ����ǲ�����������ģʽ���������ٻ��ˡ�
          --���ܳ��ڷ����ջ�ʱ��ȫ���ջ�(���ϴ�ִ��ʱ��)
          Select /*+ Rule*/
           Nvl(Count(*), 0)
          Into v_Count
          From ����ҽ����¼ A, ����ҽ������ B
          Where b.ҽ��id = a.Id And (a.Id = ҽ��id_In Or a.���id = ҽ��id_In) And
                b.���ͺ� =
                (Select Max(���ͺ�) From ����ҽ������ Where ҽ��id In (Select Column_Value From Table(t_Adviceids))) And
                a.ִ����ֹʱ�� Is Not Null And ((a.�ϴ�ִ��ʱ�� < b.ĩ��ʱ��) Or (a.�ϴ�ִ��ʱ�� Is Null And b.ĩ��ʱ�� Is Not Null));
          If v_Count > 0 Then
            If zl_GetSysParameter('�����ջز�����������', 1254) = '1' Then
              v_Error := Nvl(ҽ������_In, '��ҽ��') || '�ѱ����ڷ����ջأ������ٳ���ֹͣ������';
              Raise Err_Custom;
            Else
              --����Ѿ�ȡ���������룬���������.
              Select Count(1)
              Into v_Count
              From ���˷������� A, סԺ���ü�¼ B, ����ҽ����¼ C
              Where a.����id = b.Id And c.Id = b.ҽ����� And (c.Id = ҽ��id_In Or c.���id = ҽ��id_In);
              If v_Count > 0 Then
                v_Error := Nvl(ҽ������_In, '��ҽ��') || '�ѱ����ڷ����ջأ������ٳ���ֹͣ������';
                Raise Err_Custom;
              Else
                --�õ��ϴ�ִ��ʱ�����Ϣ
                Select �ϴ�ִ��ʱ��, ִ��ʱ�䷽��, ��ʼִ��ʱ��, �ϴδ�ӡʱ��, Ƶ�ʼ��, �����λ
                Into v_�ϴ�ʱ��, v_ִ��ʱ��, v_��ʼִ��ʱ��, v_�ϴδ�ӡʱ��, v_Ƶ�ʼ��, v_�����λ
                From ����ҽ����¼
                Where ID = ҽ��id_In;
                v_�ϴ�ʱ�� := To_Date(To_Char(v_�ϴ�ʱ�� + 1 / 24 / 60 / 60, 'yyyy-MM-dd hh24:mi:ss'), 'yyyy-MM-dd hh24:mi:ss');
              
                --�޸��ϴ�ִ��ʱ��Ϊ�ջغ��ĩ��ִ��ʱ�䡣
                v_ĩ��ʱ�� := Null;
                Begin
                  --һ��ҽ���ķ�����ĩʱ����ͬ,һ����ҩ��ȡ��С��
                  --ȡ���IDΪNULL��ҽ���ķ��ͼ�¼��ʱ��
                  --����ҩ;������ҩ�÷�����δ��д���ͼ�¼
                  Select /*+ Rule*/
                   ĩ��ʱ��, ���ͺ�
                  Into v_ĩ��ʱ��, v_���ͺ�
                  From ����ҽ������
                  Where ҽ��id In (Select Column_Value From Table(t_Adviceids)) And
                        ���ͺ� = (Select Max(���ͺ�)
                               From ����ҽ������
                               Where ҽ��id In (Select Column_Value From Table(t_Adviceids))) And Rownum = 1;
                Exception
                  When Others Then
                    Null;
                End;
                Update ����ҽ����¼ Set �ϴ�ִ��ʱ�� = v_ĩ��ʱ�� Where ID = ҽ��id_In Or ���id = ҽ��id_In;
              
                --��ԭҽ��ִ��ʱ��
                Select Zl_Adviceexetimes(ҽ��id_In, v_�ϴ�ʱ��, v_ĩ��ʱ��, v_ִ��ʱ��, v_��ʼִ��ʱ��, v_�ϴδ�ӡʱ��, v_Ƶ�ʼ��, v_�����λ, 0)
                Into v_Time
                From Dual;
                Insert Into ҽ��ִ��ʱ��
                  (Ҫ��ʱ��, ҽ��id, ���ͺ�)
                  Select To_Date(Column_Value, 'yyyy-mm-dd hh24:mi:ss'), ҽ��id_In, v_���ͺ�
                  From Table(f_Str2list(v_Time));
              End If;
            End If;
          End If;
        
          --����ȼ��䶯�������������䶯ʱ�����������
          If r_Rolladvice.��� = 'H' And Nvl(r_Rolladvice.Ӥ��, 0) = 0 Then
            Select ��������, ִ��Ƶ�� Into v_��������, v_ִ��Ƶ�� From ������ĿĿ¼ Where ID = r_Rolladvice.������Ŀid;
            If v_�������� = '1' And v_ִ��Ƶ�� = '2' Then
              Select Count(*), Max(a.����ȼ�id), Max(a.��ʼʱ��)
              Into v_Count, n_����ȼ�id, d_��ʼʱ��
              From ���˱䶯��¼ A
              Where a.����id = r_Rolladvice.����id And a.��ҳid = r_Rolladvice.��ҳid And a.��ʼԭ�� = 6 And a.��ֹʱ�� Is Null And
                    a.���Ӵ�λ = 0;
              --���û���ҵ����һ���ǻ���ȼ��䶯���ֹ
              If v_Count = 0 Then
                --ҽ������ȼ�����סʱ��Ļ���ȼ�һ��ʱҪ�����ж�
                Select Count(*)
                Into v_Count
                From ���˱䶯��¼ A
                Where a.����id = r_Rolladvice.����id And a.��ҳid = r_Rolladvice.��ҳid And a.��ʼԭ�� = 6;
                If v_Count > 0 Then
                  v_Error := '���ڻ���ȼ�ҽ��ֹͣ��ò����Ѿ������������䶯��¼,���ܻ��˸�ҽ����ֹͣ������';
                  Raise Err_Custom;
                End If;
              Else
                --���n_����ȼ�IDΪNull�������Ƿ��ǵ�ǰ���˵�ҽ����Ӧ�ı䶯��¼,Ŀ�����ж������ȼ�ҽ��ʱҪ��˳����ˡ�
                --���n_����ȼ�ID��ΪNull�����п�����У����һ������ȼ�ʱ���Զ�ֹͣ�ģ�δ�����䶯��¼��
                --     ����Ҫ��鵱ǰ���һ���䶯�Ļ���ȼ�ID�Ƿ��ǵ�ǰҽ���Ļ���ȼ�ID,Ŀ�����ж������ȼ�ҽ��ʱҪ��˳����ˣ����������Ҫ�ٳ������һ�α䶯��ֱ�ӻ���ҽ�����ɡ�
                If n_����ȼ�id Is Null Then
                  Select Count(*)
                  Into v_Count
                  From ���˱䶯��¼ B, ����ҽ���Ƽ� C
                  Where b.����id = r_Rolladvice.����id And b.��ҳid = r_Rolladvice.��ҳid And c.ҽ��id = ҽ��id_In And
                        c.�շ�ϸĿid = b.����ȼ�id And b.��ֹʱ�� = d_��ʼʱ�� And b.��ֹԭ�� = 6 And b.���Ӵ�λ = 0;
                Else
                  --��ʼʱ��ֻȡ���ӶԱȣ�У�Ե�ʱ����ȼ��Ŀ�ʼʱ����ҽ����ʼʱ��+��ǰʱ�������
                  Select Count(*)
                  Into v_Count
                  From ����ҽ���Ƽ� C, ����ҽ����¼ A
                  Where a.Id = c.ҽ��id And a.Id = ҽ��id_In And c.�շ�ϸĿid = n_����ȼ�id And
                        a.��ʼִ��ʱ�� = To_Date(To_Char(d_��ʼʱ��, 'yyyy-mm-dd hh24:mi'), 'yyyy-mm-dd hh24:mi');
                End If;
                If v_Count = 0 Then
                  v_Error := '�����˵�ҽ���������һ������ȼ�ҽ�����뽫����Ļ���ȼ�ҽ�����Ϻ��ٻ��˱���ҽ����';
                  Raise Err_Custom;
                End If;
              
                If n_����ȼ�id Is Null Then
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
                
                  Zl_���˱䶯��¼_Undo(r_Rolladvice.����id, r_Rolladvice.��ҳid, v_��Ա���, v_��Ա����, '1', Null, Null, '����ȼ��䶯');
                End If;
              End If;
            End If;
          End If;
          
          If r_Rolladvice.��� = 'Z' And  Instr(',9,10,', Nvl(r_Rolladvice.����, '0')) > 0 And Nvl(r_Rolladvice.Ӥ��, 0) = 0 Then 
            --���˲���ҽ��ʱ�����ñ䶯��¼����
            Zl_���˱䶯��¼_Undo(r_Rolladvice.����id, r_Rolladvice.��ҳid, v_��Ա���, v_��Ա����, Null, Null, Null, '�����䶯'); 
          End If;
        
          --����ҽ��ֹͣʱ,���ͣ��ҽ����ʱ��,�����ʵϰҽʦ�������˵ģ���ָ������״̬
          Update ����ҽ����¼
          Set ִ����ֹʱ�� = Decode(Flag_In, 1, ִ����ֹʱ��, Null), ͣ��ҽ�� = Null, ͣ��ʱ�� = Null,
              ��˱�� = Decode(r_Rolladvice.��˱��, 3, 2, r_Rolladvice.��˱��)
          Where ID = ҽ��id_In Or ���id = ҽ��id_In;
        Elsif r_Rolladvice.�������� = 9 Then
          --����ҽ��ȷ��ֹͣʱ,����Ƿ��Ѵ�ӡͣ��ʱ��
          Select /*+ Rule*/
           Count(*)
          Into v_Count
          From ����ҽ����ӡ
          Where ��ӡ��� = 1 And ҽ��id In (Select Column_Value From Table(t_Adviceids));
          If v_Count > 0 Then
            v_Error := Nvl(ҽ������_In, '��ҽ��') || '��ͣ��ʱ���Ѿ���ӡ�������ٳ���ȷ��ֹͣ������';
            Raise Err_Custom;
          End If;
        
          --����ҽ��ȷ��ֹͣʱ,���ͣ��ҽ����ʱ��
          Update ����ҽ����¼ Set ȷ��ͣ��ʱ�� = Null, ȷ��ͣ����ʿ = Null Where ID = ҽ��id_In Or ���id = ҽ��id_In;
        Elsif r_Rolladvice.�������� = 10 Then
          --���˱�עƤ�Խ��,ͬʱɾ�������Ǽ�(+)��(-),���ݼ�¼ʱ��
          Delete From ���˹�����¼
          Where ����id = r_Rolladvice.����id And Nvl(��ҳid, 0) = Nvl(r_Rolladvice.��ҳid, 0) And ��¼ʱ�� = r_Rolladvice.����ʱ��;
        
          Update ����ҽ����¼ Set Ƥ�Խ�� = Null Where ID = ҽ��id_In Or ���id = ҽ��id_In;
        Elsif r_Rolladvice.�������� = 13 Then
          If Instr(r_Rolladvice.����ҽ��, '/') > 0 Then
            Update ����ҽ����¼ Set ��˱�� = 1 Where ID = ҽ��id_In Or ���id = ҽ��id_In;
          Else
            Update ����ҽ����¼ Set ��˱�� = Null Where ID = ҽ��id_In Or ���id = ҽ��id_In;
          End If;
        End If;
      Else
        --����ҽ������(�Է��ͺŹؼ���)
        ------------------------------------------------------------------
        --��ǰ������Ա
        v_Temp     := Zl_Identity;
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
        v_��Ա��� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
        v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
      
        --����Ƿ�����Һ��Һ��¼�����Ƿ��Ѿ������������ѯ������˵������Һ��¼
        Begin
          Select Decode(max(�Ƿ�����), 1, 1, 0)
          Into v_Count
          From ��Һ��ҩ��¼
          Where ҽ��id = ҽ��id_In And ���ͺ� = r_Rolladvice.���ͺ�;
        Exception
          When Others Then
            v_Count := -1;
        End;
      
        If v_Count = 1 Then
          v_Error := 'ҽ��"' || ҽ������_In || '"����ҺҩƷ���Ѿ�����Һ�����������������ܻ��˷��͡�';
          Raise Err_Custom;
        Elsif v_Count = 0 Then
          Zl_��Һ��ҩ��¼_ҽ������(ҽ��id_In, r_Rolladvice.���ͺ�, v_��Ա����, Sysdate);
        End If;
      
        --���Ʒ����Զ�ִ��ʱ������Ҳ�Զ�����ִ��(����ʿվ�д˹���)
        --�Ǹ������õ�����ҽ����ͬ��ͨҽ��ִ�д���
        Select ҽ����Ч Into v_ҽ����Ч From ����ҽ����¼ Where ID = ҽ��id_In;
        If Substr(zl_GetSysParameter('����ִ���Զ����', 1254), v_ҽ����Ч + 1, 1) = '1' Then
          Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into Intdigit From Dual;
        
          For r_Rollsend In c_Rollsend(r_Rolladvice.���ͺ�) Loop
            If Nvl(r_Rollsend.ִ��״̬, 0) = 1 And
               (Nvl(r_Rollsend.ִ�п���id, 0) = Nvl(r_Rollsend.���˲���id, 0) Or
                Nvl(r_Rollsend.ִ�п���id, 0) = Nvl(r_Rollsend.���˿���id, 0)) Then
            
              --ҽ����ִ��״̬
              Update ����ҽ������ Set ִ��״̬ = 0 Where ���ͺ� = r_Rollsend.���ͺ� And ҽ��id = r_Rollsend.ҽ��id;
              v_Update := 1;
            
              If r_Rolladvice.��¼���� = 2 And Nvl(r_Rolladvice.�������, 0) = 0 Then
                --���õ�ִ��״̬
                For r_Rollmoney In c_Rollmoneyin(r_Rollsend.���ͺ�, r_Rollsend.ҽ��id, t_Adviceids) Loop
                  n_Blndo := 0;
                  If r_Rollmoney.��¼״̬ <> 3 Then
                    n_Blndo := 1;
                  Else
                    n_Blndo := Checkmoneyundo(r_Rollmoney.No, r_Rollmoney.��¼����, r_Rollmoney.���);
                  End If;
                  If n_Blndo > 0 Then
                    If Not (r_Rollmoney.�շ���� = '4' And Nvl(r_Rollmoney.��������, 0) = 1) And
                       Not r_Rollmoney.�շ���� In ('5', '6', '7') Then
                      --��ͨ����ֱ��ȡ��ִ��״̬������ҩƷ�͸������õ�����
                      Update סԺ���ü�¼
                      Set ִ��״̬ = 0, ִ��ʱ�� = Null, ִ���� = Null
                      Where NO = r_Rollmoney.No And ��¼���� = r_Rolladvice.��¼���� And ��¼״̬ = r_Rollmoney.��¼״̬ And
                            Nvl(�۸񸸺�, ���) = r_Rollmoney.��� And ҽ����� = r_Rollsend.ҽ��id;
                    Elsif r_Rollmoney.�շ���� = '4' And Nvl(r_Rollmoney.��������, 0) = 1 Then
                      --�������õ����ģ���ϵͳ����Ϊ�Զ�����ʱ�����Զ�����
                      If Nvl(zl_GetSysParameter(33), '0') = '1' Then
                        For r_Stuff In c_Stuff_Drug(r_Rollmoney.Id) Loop
                          Zl_�����շ���¼_��������(r_Stuff.Id, v_��Ա����, Sysdate, Null, Null, Null, Null, 0, v_��Ա����);
                        End Loop;
                      End If;
                    Elsif r_Rollmoney.�շ���� In ('5', '6', '7') Then
                      --סԺ���ҷ�ҩ��ҩƷ�Զ���ҩ
                      If r_Rollmoney.ִ�в���id = r_Rollsend.���˲���id Or r_Rollmoney.ִ�в���id = r_Rollsend.���˿���id Then
                        For r_Drug In c_Stuff_Drug(r_Rollmoney.Id) Loop
                          Zl_ҩƷ�շ���¼_������ҩ(r_Drug.Id, v_��Ա����, Sysdate, Null, Null, Null, Null, Null, Null, Intdigit, 2);
                        End Loop;
                      End If;
                    End If;
                  End If;
                End Loop;
              Else
                --סԺ���˷��÷��͵�����������������Դ����סԺ��
                For r_Rollmoney In c_Rollmoneyout(r_Rollsend.���ͺ�, r_Rollsend.ҽ��id, t_Adviceids) Loop
                  n_Blndo := 0;
                  If r_Rollmoney.��¼״̬ <> 3 Then
                    n_Blndo := 1;
                  Else
                    n_Blndo := Checkmoneyundo(r_Rollmoney.No, r_Rollmoney.��¼����, r_Rollmoney.���, 1);
                  End If;
                  If n_Blndo > 0 Then
                    If Not (r_Rollmoney.�շ���� = '4' And Nvl(r_Rollmoney.��������, 0) = 1) And
                       Not r_Rollmoney.�շ���� In ('5', '6', '7') Then
                      --��ͨ����ֱ��ȡ��ִ��״̬������ҩƷ�͸������õ�����
                      Update ������ü�¼
                      Set ִ��״̬ = 0, ִ��ʱ�� = Null, ִ���� = Null
                      Where NO = r_Rollmoney.No And ��¼���� = r_Rolladvice.��¼���� And ��¼״̬ = r_Rollmoney.��¼״̬ And
                            Nvl(�۸񸸺�, ���) = r_Rollmoney.��� And ҽ����� = r_Rollsend.ҽ��id;
                    Elsif r_Rollmoney.�շ���� = '4' And Nvl(r_Rollmoney.��������, 0) = 1 Then
                      If Nvl(zl_GetSysParameter(33), '0') = '1' Then
                        --�������õ����ģ���ϵͳ����Ϊ�Զ�����ʱ�����Զ�����
                        For r_Stuff In c_Stuff_Drug(r_Rollmoney.Id) Loop
                          Zl_�����շ���¼_��������(r_Stuff.Id, v_��Ա����, Sysdate, Null, Null, Null, Null, 0, v_��Ա����);
                        End Loop;
                      End If;
                    Elsif r_Rollmoney.�շ���� In ('5', '6', '7') Then
                      --�����ҷ�ҩ��ҩƷ�Զ���ҩ
                      If r_Rollmoney.ִ�в���id = r_Rollsend.���˲���id Or r_Rollmoney.ִ�в���id = r_Rollsend.���˿���id Then
                        For r_Drug In c_Stuff_Drug(r_Rollmoney.Id) Loop
                          Zl_ҩƷ�շ���¼_������ҩ(r_Drug.Id, v_��Ա����, Sysdate, Null, Null, Null, Null, Null, Null, Intdigit, 1);
                        End Loop;
                      End If;
                    End If;
                  End If;
                End Loop;
              End If;
            End If;
          End Loop;
        End If;
        ------------------------------------------------------------------
        --�������ջصĳ���ҩƷҽ�����������(���˷��þͶ�����)
        If Nvl(r_Rolladvice.ҽ����Ч, 0) = 0 Then
          If r_Rolladvice.�ϴ�ִ��ʱ�� Is Not Null And r_Rolladvice.ĩ��ʱ�� Is Not Null Then
            If r_Rolladvice.�ϴ�ִ��ʱ�� < r_Rolladvice.ĩ��ʱ�� Then
              v_Error := Nvl(ҽ������_In, '��ҽ��') || '������ڷ��͵������ѱ��ջأ������ٻ��ˡ�';
              Raise Err_Custom;
            End If;
          Elsif r_Rolladvice.�ϴ�ִ��ʱ�� Is Null And r_Rolladvice.ĩ��ʱ�� Is Not Null Then
            --�������ܱ�ȫ�������ջ�
            v_Error := Nvl(ҽ������_In, '��ҽ��') || 'δ�����ͣ����͵������ѱ�ȫ�������ջأ������ٻ��ˡ�';
            Raise Err_Custom;
          End If;
        End If;
      
        If Nvl(r_Rolladvice.ִ��״̬, 0) In (1, 3) And v_Update <> 1 Then
          --1-��ȫִ��;3-����ִ��
          v_Error := Nvl(ҽ������_In, '��ҽ��') || '������͵������Ѿ�ִ�л�����ִ�У����ܻ��ˡ�';
          Raise Err_Custom;
        Else
          --������ҽ����ִ�У���ҲҪ���ƻ��ˣ����磺����Ĳɼ���ʽ��
          Select /*+ Rule*/
           Count(1)
          Into v_Count
          From ����ҽ������
          Where ҽ��id In (Select Column_Value From Table(t_Adviceids)) And ִ��״̬ In (1, 3) And
                ���ͺ� =
                (Select Max(���ͺ�) From ����ҽ������ Where ҽ��id In (Select Column_Value From Table(t_Adviceids)));
          If v_Count > 0 Then
            v_Error := Nvl(ҽ������_In, '��ҽ��') || '������͵������Ѿ�ִ�л�����ִ�У����ܻ��ˡ�';
            Raise Err_Custom;
          End If;
        End If;
      
        ------------------------------------------------------------------
        --������ҽ���ķ�������(��һ��ҽ�������в�ͬNO����)
        --���ԭʼ�����ѱ�����(�򲿷�����),���ù��������ж�
        v_����no   := Null;
        v_������� := Null;
        If r_Rolladvice.��¼���� = 2 And Nvl(r_Rolladvice.�������, 0) = 0 Then
          For r_Rollmoney In c_Rollmoneyin(r_Rolladvice.���ͺ�, Null, t_Adviceids) Loop
            --��Ӧ�ķ�����ִ��
            If Nvl(r_Rollmoney.ִ��״̬, 0) <> 0 Then
              v_Error := Nvl(ҽ������_In, '��ҽ��') || '���͵ķ��õ���"' || r_Rollmoney.No || '"�е������ѱ����ֻ���ȫִ�У����ܻ��ˡ�';
              Raise Err_Custom;
            End If;
            n_Blndo := 0;
            If r_Rollmoney.��¼״̬ <> 3 Then
              n_Blndo := 1;
            Else
              n_Blndo := Checkmoneyundo(r_Rollmoney.No, r_Rollmoney.��¼����, r_Rollmoney.���);
            End If;
            If n_Blndo > 0 Then
              --���ֽ������жϲ�����ҩ
              If v_����no <> r_Rollmoney.No And v_������� Is Not Null Then
                Zl_סԺ���ʼ�¼_Delete(v_����no, Substr(v_�������, 2), v_��Ա���, v_��Ա����, 2, 0, 0);
                v_������� := Null;
              End If;
              v_����no   := r_Rollmoney.No;
              v_������� := v_������� || ',' || r_Rollmoney.���;
            End If;
          End Loop;
        Else
          For r_Rollmoney In c_Rollmoneyout(r_Rolladvice.���ͺ�, Null, t_Adviceids) Loop
            --��Ӧ�ķ�����ִ��
            If Nvl(r_Rollmoney.ִ��״̬, 0) <> 0 And Not (Nvl(r_Rollmoney.ִ��״̬, 0) = -1 And Nvl(r_Rollmoney.��¼״̬, 0) = 0) Then
              v_Error := Nvl(ҽ������_In, '��ҽ��') || '���͵ķ��õ���"' || r_Rollmoney.No || '"�е������ѱ����ֻ���ȫִ�У����ܻ��ˡ�';
              Raise Err_Custom;
            End If;
            --�շѵ������շ�
            If r_Rollmoney.��¼״̬ = 1 And Not (r_Rolladvice.��¼���� = 2 And Nvl(r_Rolladvice.�������, 0) = 1) Then
              v_Error := Nvl(ҽ������_In, '��ҽ��') || '���͵����ﵥ��"' || r_Rollmoney.No || '"���շѣ����ܻ��ˡ�';
              Raise Err_Custom;
            End If;
            n_Blndo := 0;
            If r_Rollmoney.��¼״̬ <> 3 Then
              n_Blndo := 1;
            Else
              n_Blndo := Checkmoneyundo(r_Rollmoney.No, r_Rollmoney.��¼����, r_Rollmoney.���, 1);
            End If;
            If n_Blndo > 0 Then
              --���ֽ������жϲ�����ҩ
              If v_����no <> r_Rollmoney.No And v_������� Is Not Null Then
                If r_Rolladvice.��¼���� = 2 And Nvl(r_Rolladvice.�������, 0) = 1 Then
                  --סԺ����Ϊ�������(���������ҽ������Ϊ������ʣ�����ҽ��û�л��˹���)
                  Zl_������ʼ�¼_Delete(v_����no, v_�������, v_��Ա���, v_��Ա����);
                Else
                  Zl_���ﻮ�ۼ�¼_Delete(v_����no, Substr(v_�������, 2));
                End If;
                v_������� := Null;
              End If;
              v_����no   := r_Rollmoney.No;
              v_������� := v_������� || ',' || r_Rollmoney.���;
            End If;
          End Loop;
        End If;
        If v_������� Is Not Null And v_����no Is Not Null Then
          v_������� := Substr(v_�������, 2);
          If r_Rolladvice.��¼���� = 2 And Nvl(r_Rolladvice.�������, 0) = 0 Then
            Zl_סԺ���ʼ�¼_Delete(v_����no, v_�������, v_��Ա���, v_��Ա����, 2, 0, 0);
          Elsif r_Rolladvice.��¼���� = 2 And Nvl(r_Rolladvice.�������, 0) = 1 Then
            --סԺ����Ϊ�������(���������ҽ������Ϊ������ʣ�����ҽ��û�л��˹���)
            Zl_������ʼ�¼_Delete(v_����no, v_�������, v_��Ա���, v_��Ա����);
          Else
            Zl_���ﻮ�ۼ�¼_Delete(v_����no, v_�������);
          End If;
        End If;
        --��Ѫҽ����ɾ������ҽ������
        Delete From ����ҽ������ Where ���ͺ� = r_Rolladvice.���ͺ� And ҽ��id = ҽ��id_In;
      
        --ɾ��ҽ��ִ��ʱ�� (����ҽ��ID�Ų����˼�¼)
        Delete From ҽ��ִ��ʱ�� Where ���ͺ� = r_Rolladvice.���ͺ� And ҽ��id = ҽ��id_In;
      
        --ɾ�����ͼ�¼(����ҽ����)
        Delete /*+ Rule*/
        From ����ҽ������
        Where ���ͺ� = r_Rolladvice.���ͺ� And ҽ��id In (Select Column_Value From Table(t_Adviceids));
      
        --���(����ҽ��)�ϴ�ִ��ʱ��(���ϴη��͵�ĩ��ִ��ʱ��)
        --���г���(���������Գ���)����ʱ����д��ĩ��ʱ��
        --��������û�У���ֻ���ܷ�����һ�Ρ�
        v_ĩ��ʱ�� := Null;
        Begin
          --һ��ҽ���ķ�����ĩʱ����ͬ,һ����ҩ��ȡ��С��
          --ȡ���IDΪNULL��ҽ���ķ��ͼ�¼��ʱ��
          --����ҩ;������ҩ�÷�����δ��д���ͼ�¼
          Select /*+ Rule*/
           ĩ��ʱ��
          Into v_ĩ��ʱ��
          From ����ҽ������
          Where ҽ��id In (Select Column_Value From Table(t_Adviceids)) And
                ���ͺ� =
                (Select Max(���ͺ�) From ����ҽ������ Where ҽ��id In (Select Column_Value From Table(t_Adviceids))) And
                Rownum = 1;
        Exception
          When Others Then
            Null;
        End;
        Update ����ҽ����¼ Set �ϴ�ִ��ʱ�� = v_ĩ��ʱ�� Where ID = ҽ��id_In Or ���id = ҽ��id_In;
      
        --������������ʱ��ͬʱ�Զ�����ֹͣ
        If Nvl(r_Rolladvice.ҽ����Ч, 0) = 1 Then
          --ɾ��(��������)�����ֹͣ״̬������¼
          Delete /*+ Rule*/
          From ����ҽ��״̬
          Where ҽ��id In (Select Column_Value From Table(t_Adviceids)) And
                ����ʱ�� = (Select Max(����ʱ��) From ����ҽ��״̬ Where ҽ��id = ҽ��id_In) And �������� = 8;
          --r_RollAdvice.����ʱ��:����ʱ����ܲ����Զ�ֹͣʱ����ͬ��
        
          --ȡɾ����Ӧ�ָ���ҽ��״̬
          Select ��������
          Into v_ҽ��״̬
          From ����ҽ��״̬
          Where ����ʱ�� = (Select Max(����ʱ��) From ����ҽ��״̬ Where ҽ��id = ҽ��id_In) And ҽ��id = ҽ��id_In;
        
          --�ָ�(����ҽ��)���˺��״̬
          Update ����ҽ����¼
          Set ҽ��״̬ = v_ҽ��״̬, ִ����ֹʱ�� = Null, ͣ��ҽ�� = Null, ͣ��ʱ�� = Null
          Where ID = ҽ��id_In Or ���id = ҽ��id_In;
        End If;
      
        --סԺ����ҽ�����ͺ�Ļ���(3-ת��;5-��Ժ;6-תԺ,11-����)
        If r_Rolladvice.��� = 'Z' And Instr(',3,5,6,11,', Nvl(r_Rolladvice.����, '0')) > 0 And Nvl(r_Rolladvice.Ӥ��, 0) = 0 Then
          Open c_Patilog(r_Rolladvice.����id, r_Rolladvice.��ҳid);
          Fetch c_Patilog
            Into r_Patilog;
          If c_Patilog%Found Then
            If r_Rolladvice.���� = '3' And r_Patilog.��ʼԭ�� = 3 Then
              --ȡ������ת��״̬
              If r_Patilog.��ʼʱ�� Is Null Then
                --ת��ҽ�������⴦����һ������������ת��ҽ��ʱ��ֻ�ܻ��������һ��,70443
                Select Count(1)
                Into v_Count
                From ����ҽ����¼ A, ������ĿĿ¼ B
                Where a.������Ŀid = b.Id And a.����id = r_Rolladvice.����id And a.��ҳid = r_Rolladvice.��ҳid And a.������� = 'Z' And
                      b.�������� = '3' And a.ҽ��״̬ = 8 And
                      a.��ʼִ��ʱ�� > (Select ��ʼִ��ʱ�� From ����ҽ����¼ Where ID = ҽ��id_In);
                If v_Count = 0 Then
                  Zl_���˱䶯��¼_Undo(r_Rolladvice.����id, r_Rolladvice.��ҳid, v_��Ա���, v_��Ա����, Null, Null, Null, 'ת��');
                Else
                  v_Error := '����ת���Ѿ���ƣ������ٻ��ˡ�';
                  Raise Err_Custom;
                End If;
              Else
                v_Error := '����ת���Ѿ���ƣ������ٻ��ˡ�';
                Raise Err_Custom;
              End If;
            Elsif r_Rolladvice.���� In ('5', '6', '11') And r_Patilog.��ʼԭ�� = 10 Then
              --ȡ������Ԥ��Ժ״̬
              Zl_���˱䶯��¼_Undo(r_Rolladvice.����id, r_Rolladvice.��ҳid, v_��Ա���, v_��Ա����, Null, Null, Null, 'Ԥ��Ժ');
            End If;
          End If;
          Close c_Patilog;
        End If;
      
        --���˲���ʱ��
        --1.�����¼�(ֻ��һ��ҽ����¼)��������7-����,8-����,11-����
        If r_Rolladvice.��� = 'F' Or r_Rolladvice.��� = 'Z' And Instr('7,8,11', r_Rolladvice.����) > 0 Then
          Zl_���Ӳ���ʱ��_Delete(r_Rolladvice.����id, r_Rolladvice.��ҳid, 'ҽ��', r_Rolladvice.��������id, ҽ��id_In);
        End If;
      
        --2.���⴦��֪��ͬ����(������ص�֪��ͬ�����ٴε��ã���Ϊ����������������Ŀ�����й�����֪��ͬ����)
        If Instr('C,D,E,F,G,K,L', r_Rolladvice.���) > 0 Then
          For R In (Select a.Id, a.������� From ����ҽ����¼ A Where a.Id = ҽ��id_In Or a.���id = ҽ��id_In) Loop
            --���id��һ��ҽ����һ����������ģ�����Ҫ���ж�һ�����
            If Instr('C,D,E,F,G,K,L', r.�������) > 0 Then
              Zl_���Ӳ���ʱ��_Delete(r_Rolladvice.����id, r_Rolladvice.��ҳid, 'ҽ��', r_Rolladvice.��������id, r.Id);
            End If;
          End Loop;
        End If;
      End If;
    End If;
    Exit When r_Rolladvice.���ͺ� = 0;
  End Loop;
  Close c_Rolladvice;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ҽ����¼_����;
/

--93917:����,2016-03-16,��RIS�����ݽ����ӿڽű�
Create Or Replace Package b_zlXWInterface Is
  Type t_Refcur Is Ref Cursor;

  --1������RIS״̬�ı�
  Procedure ReceiveRisState
  (
    ҽ��id_In   ����ҽ������.ҽ��id%Type,
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
  Procedure ReceiveReport
  (
    ҽ��id_In   ����ҽ������.ҽ��id%Type,
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

End b_zlXWInterface;
/

Create Or Replace Package Body b_zlXWInterface Is

  --1������RIS״̬�ı�
  Procedure ReceiveRisState
  (
    ҽ��id_In   ����ҽ������.ҽ��id%Type,
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
      Select ID, ���id, Nvl(���id, ID) As ��id, �������, ������Դ, ִ�п���id
      From ����ҽ����¼
      Where ID = ҽ��id_In;
    r_Adviceinfo c_Adviceinfo%RowType;
  
    v_Count Number;
    v_Error Varchar2(255);
    Err_Custom Exception;
    v_ִ��״̬ ����ҽ������.ִ��״̬%Type;
    v_ִ�й��� ����ҽ������.ִ�й���%Type;
  
  Begin
  
    v_ִ��״̬ := 0;
    v_ִ�й��� := 0;
  
    --��ȡҽ������ҽ��ID������ID
    Open c_Adviceinfo;
    Fetch c_Adviceinfo
      Into r_Adviceinfo;
    Close c_Adviceinfo;
  
    --�ȼ���Ƿ��Ѿ���Ժ��סԺ���ˣ��Ѿ�Ԥ��Ժ���߳�Ժ�ļ�����룬������ʼ������˷���
    Select Count(*)
    Into v_Count
    From ����ҽ����¼ A, ������ҳ B
    Where a.����id = b.����id And a.��ҳid = b.��ҳid And (b.��Ժ���� Is Not Null Or b.״̬ = 3) And a.Id = r_Adviceinfo.��id;
  
    If v_Count > 0 Then
      v_Error := 'סԺ�����Ѿ���Ժ����Ԥ��Ժ���޷���ʼ��顣';
      Raise Err_Custom;
    End If;
  
    --����״̬_INִ��ҽ��
    ---1-ɾ����0-ԤԼ��1-�Ǽǣ�3-�����ɣ�4-�����ֹ��9-�������棻12-������ˣ�15-����
  
    If ״̬_In = -1 Or ״̬_In = 0 Then
      v_ִ��״̬ := 0; --δִ��
      v_ִ�й��� := 0;
    Elsif ״̬_In = 1 Then
      v_ִ��״̬ := 3; --����ִ��
      v_ִ�й��� := 2; --�ѱ���
    Elsif ״̬_In = 3 Then
      v_ִ��״̬ := 3; --����ִ��
      v_ִ�й��� := 3; --�Ѽ��
    Elsif ״̬_In = 4 Then
      --���ı�
      v_ִ��״̬ := v_ִ��״̬;
    Elsif ״̬_In = 9 Then
      v_ִ��״̬ := 3; --����ִ��
      v_ִ�й��� := 4; --�ѱ���
    Elsif ״̬_In = 12 Then
      v_ִ��״̬ := 3; --����ִ��
      v_ִ�й��� := 5; --�����
    Elsif ״̬_In = 15 Then
      v_ִ��״̬ := 1; --��ȫִ��
      v_ִ�й��� := 6; --�����
    End If;
  
    --��ʼִ��ҽ��
    If Nvl(����ִ��_In, 0) = 1 Then
      -- ������λҽ������ִ��
      Update ����ҽ������
      Set ִ��״̬ = v_ִ��״̬, ִ�й��� = v_ִ�й���, ִ��˵�� = ִ��˵��_In
      Where ҽ��id = ҽ��id_In;
    Else
      Update ����ҽ������
      Set ִ��״̬ = v_ִ��״̬, ִ�й��� = v_ִ�й���, ִ��˵�� = ִ��˵��_In
      Where ҽ��id In (Select ID From ����ҽ����¼ Where (ID = r_Adviceinfo.��id Or ���id = r_Adviceinfo.��id));
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
  Begin
  
    Select ���ͺ�, ִ�й��� Into v_���ͺ�, v_ִ�й��� From ����ҽ������ Where ҽ��id = ҽ��id_In;
  
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
    
      --ȡ��ҽ��ID
      Open c_Advice;
      Fetch c_Advice
        Into r_Advice;
      Close c_Advice;
    
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
  
    v_���ͺ� ����ҽ������.���ͺ�%Type;
  Begin
  
    Select ���ͺ� Into v_���ͺ� From ����ҽ������ Where ҽ��id = ҽ��id_In;
  
    --����ͳһ��ҽ��ִ��Cancel����
    Zl_����ҽ��ִ��_Cancel(ҽ��id_In, v_���ͺ�, Null, ����ִ��_In, ִ�в���id_In, ����Ա���_In, ����Ա����_In);
  
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Ӱ�����ִ��_Cancel;

  --4������RIS�ı���
  Procedure ReceiveReport
  (
    ҽ��id_In   ����ҽ������.ҽ��id%Type,
    ��������_In ���Ӳ�������.�����ı�%Type,
    �������_In ���Ӳ�������.�����ı�%Type,
    ���潨��_In ���Ӳ�������.�����ı�%Type,
    ����ҽ��_In ���Ӳ�����¼.������%Type
  ) Is
  
    --��ȡ����ҽ��������������Ϣ
    Cursor c_Advice(v_��id Number) Is
      Select e.Id, e.������Դ, e.����id, e.��ҳid, e.Ӥ��, e.���˿���id, e.�ļ�id, e.��������, e.��������, f.����id, e.ִ�п���id
      From (Select c.Id, c.������Դ, c.����id, c.��ҳid, c.Ӥ��, c.���˿���id, c.�ļ�id, d.���� ��������, d.���� ��������, c.ִ�п���id
             From (Select a.Id, a.������Դ, a.����id, a.��ҳid, a.Ӥ��, a.���˿���id, b.�����ļ�id �ļ�id, a.ִ�п���id
                    From ����ҽ����¼ A, ��������Ӧ�� B
                    Where a.Id = v_��id And a.������Ŀid = b.������Ŀid(+) And b.Ӧ�ó���(+) = Decode(a.������Դ, 2, 2, 4, 4, 1)) C,
                  �����ļ��б� D
             Where c.�ļ�id = d.Id(+)) E, ����ҽ������ F
      Where e.Id = f.ҽ��id(+);
  
    --�����ļ������Ԫ��
    Cursor c_File(v_File Number) Is
      Select a.Id, a.�ļ�id, a.��id, a.�������, a.��������, a.������, a.��������, a.��������, a.�����д�, a.�����ı�, a.�Ƿ���, a.Ԥ�����id, a.�������,
             a.ʹ��ʱ��, a.����Ҫ��id, a.�滻��, a.Ҫ������, a.Ҫ������, a.Ҫ�س���, a.Ҫ��С��, a.Ҫ�ص�λ, a.Ҫ�ر�ʾ, a.������̬, a.Ҫ��ֵ��
      From �����ļ��ṹ A
      Where a.�ļ�id = v_File
      Order By a.�������;
  
    Cursor c_Report(v_���Ӳ�����¼id Number) Is
      Select /*+ rule */
       b.Id, a.�����ı�
      From ���Ӳ������� A, ���Ӳ������� B
      Where a.�ļ�id = v_���Ӳ�����¼id And Nvl(a.�������id, 0) <> 0 And (a.�����ı� Like '%����%' Or a.�����ı� Like '%���%' Or a.�����ı� Like '%����%') And
            b.��id = a.Id And b.�Ƿ��� = 1;
  
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
  Begin
  
    -- ��ȡ��ҽ��ID ����ֹ��Ϊ���벿λҽ�������±��汣�����
    Select Nvl(���id, ID) As ��id Into v_��ҽ��id From ����ҽ����¼ Where ID = ҽ��id_In;
  
    Open c_Advice(v_��ҽ��id);
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
            Update ���Ӳ������� Set �����ı� = ��������_In Where ID = r_Report.Id;
          Elsif r_Report.�����ı� Like '%���%' Then
            Update ���Ӳ������� Set �����ı� = �������_In Where ID = r_Report.Id;
          Elsif r_Report.�����ı� Like '%����%' Then
            Update ���Ӳ������� Set �����ı� = ���潨��_In Where ID = r_Report.Id;
          End If;
        End Loop;
        --���±���ʱ��
        Update ���Ӳ�����¼
        Set ���ʱ�� = Sysdate, ������ = ����ҽ��_In, ����ʱ�� = Sysdate
        Where ID = r_Advice.����id;
      Else
        --�������Ӳ�����¼
        Select ���Ӳ�����¼_Id.Nextval Into v_����id From Dual;
        Insert Into ���Ӳ�����¼
          (ID, ������Դ, ����id, ��ҳid, Ӥ��, ����id, ��������, �ļ�id, ��������, ������, ����ʱ��, ���ʱ��, ������, ����ʱ��, ���汾, ǩ������)
        Values
          (v_����id, r_Advice.������Դ, r_Advice.����id, r_Advice.��ҳid, r_Advice.Ӥ��, r_Advice.���˿���id, r_Advice.��������,
           r_Advice.�ļ�id, r_Advice.��������, ����ҽ��_In, Sysdate, Sysdate, ����ҽ��_In, Sysdate, 1, 2);
      
        --����ҽ�������¼
        Insert Into ����ҽ������ (ҽ��id, ����id) Values (v_��ҽ��id, v_����id);
      
	    v_������� := 0;
		
        --�²�����������
        For r_File In c_File(r_Advice.�ļ�id) Loop
          Select ���Ӳ�������_Id.Nextval Into v_��������id From Dual;
        
          If Nvl(r_File.��id, 0) <> 0 And (r_File.�����ı� Like '%����%') Then
            --����������(�����)
            v_�����ı�   := ��������_In || Chr(13) || Chr(13);
            v_�������id := 0;
          Elsif Nvl(r_File.��id, 0) <> 0 And (r_File.�����ı� Like '%���%') Then
            --���������(�����)
            v_�����ı�   := �������_In || Chr(13) || Chr(13);
            v_�������id := 0;
          Elsif Nvl(r_File.��id, 0) <> 0 And (r_File.�����ı� Like '%����%') Then
            --���鶨����(�����)
            v_�����ı�   := ���潨��_In || Chr(13) || Chr(13);
            v_�������id := 0;
          Elsif Nvl(r_File.��������, 0) = 1 And Nvl(r_File.��id, 0) = 0 Then
            --��ٶ�����
            v_��id       := v_��������id;
            v_�����ı�   := r_File.�����ı�;
            v_�������id := r_File.Id;
          Elsif Nvl(r_File.��������, 0) = 4 And r_File.Ҫ������ Is Not Null Then
            --�Զ��滻Ҫ��
            v_�����ı�   := Zl_Replace_Element_Value(r_File.Ҫ������, r_Advice.����id, r_Advice.��ҳid, r_Advice.������Դ, r_Advice.Id);
            v_�������id := 0;
          Else
            v_�����ı�   := r_File.�����ı�;
            v_�������id := 0;
          End If;
        
          --�������ݵ���дһ��
          If Nvl(r_File.��id, 0) <> 0 And (r_File.�����ı� Like '%����%' Or r_File.�����ı� Like '%���%' Or r_File.�����ı� Like '%����%') Then
            --��д�����ʾ���ƣ���д���ݣ�ͬʱ������ŷ����仯
            Select ���Ӳ�������_Id.Nextval Into v_��������idnew From Dual;
            v_������� := v_������� + 1;
            Insert Into ���Ӳ�������
              (ID, �ļ�id, ��ʼ��, ��ֹ��, ��id, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���, Ԥ�����id, �������, ʹ��ʱ��, ����Ҫ��id, �滻��,
               Ҫ������, Ҫ������, Ҫ�س���, Ҫ��С��, Ҫ�ص�λ, Ҫ�ر�ʾ, ������̬, Ҫ��ֵ��, �������id)
            Values
              (v_��������idnew, v_����id, 0, 0, Decode(v_�������id, 0, v_��id, Null), v_�������, r_File.��������, r_File.������,
               r_File.��������, 0, Null, v_�����ı�, r_File.�Ƿ���, r_File.Ԥ�����id, r_File.�������, r_File.ʹ��ʱ��, r_File.����Ҫ��id,
               r_File.�滻��, r_File.Ҫ������, r_File.Ҫ������, r_File.Ҫ�س���, r_File.Ҫ��С��, r_File.Ҫ�ص�λ, r_File.Ҫ�ر�ʾ, r_File.������̬,
               r_File.Ҫ��ֵ��, Decode(v_�������id, 0, Null, v_�������id));
            
            v_�����ı� := r_File.�����ı�;
          End If;
        
		  v_������� := v_������� + 1;
		  
          Insert Into ���Ӳ�������
            (ID, �ļ�id, ��ʼ��, ��ֹ��, ��id, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���, Ԥ�����id, �������, ʹ��ʱ��, ����Ҫ��id, �滻��,
             Ҫ������, Ҫ������, Ҫ�س���, Ҫ��С��, Ҫ�ص�λ, Ҫ�ر�ʾ, ������̬, Ҫ��ֵ��, �������id)
          Values
            (v_��������id, v_����id, 1, 0, Decode(v_�������id, 0, v_��id, Null), v_�������, r_File.��������, r_File.������, r_File.��������,
             r_File.��������, Null, v_�����ı�, r_File.�Ƿ���, r_File.Ԥ�����id, r_File.�������, r_File.ʹ��ʱ��, r_File.����Ҫ��id, r_File.�滻��,
             r_File.Ҫ������, r_File.Ҫ������, r_File.Ҫ�س���, r_File.Ҫ��С��, r_File.Ҫ�ص�λ, r_File.Ҫ�ر�ʾ, r_File.������̬, r_File.Ҫ��ֵ��,
             Decode(v_�������id, 0, Null, v_�������id));
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
end b_zlXWInterface;
/

--93623:���Ʊ�,2016-03-04,���ҽ���Ʒ�״̬����
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
          Nvl(a.ҽ����Ч, 0) = 0 And a.ҽ��״̬ Not In (1, 2, 4, 8, 9) And a.��ʼִ��ʱ�� < v_Stoptime
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
          If r_Advice.��ʼִ��ʱ�� <= v_Date Then
            v_Error := 'ҽ��"' || r_Advice.ҽ������ || '"�Ŀ�ʼʱ��Ӧ���ڸò����ϴα䶯ʱ�� ' || To_Char(v_Date, 'YYYY-MM-DD HH24:Mi') || ' ��';
            Raise Err_Custom;
          End If;
          Zl_���˱䶯��¼_Preout(r_Advice.����id, r_Advice.��ҳid, r_Advice.��ʼִ��ʱ��);
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

--93623:���Ʊ�,2016-03-04,���Ʒ�״̬����
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
  ����Ա����_In ��Ա��.����%Type := Null
  --���ܣ���д����ҽ�����ͼ�¼ 
  --������First_IN=��ʾ�Ƿ�һ��ҽ���ĵ�һҽ����,�Ա㴦��ҽ���������(���ҩ,�䷽�ĵ�һ��,��Ϊ��ҩ;��,�䷽�巨,�÷�����Ϊ����������) 
) Is
  --�������˼�ҽ��(һ��ҽ���е�һ��)�����Ϣ���α� 
  Cursor c_Advice Is
    Select Nvl(a.���id, a.Id) As ��id, a.���, a.����id, a.�Һŵ�, a.Ӥ��, b.����, c.��������, a.�������, a.ҽ��״̬, a.ҽ������, a.����ҽ��, a.��ʼִ��ʱ��,
           a.ִ��ʱ�䷽��, a.Ƶ�ʴ���, a.Ƶ�ʼ��, a.�����λ, Nvl(a.������־, 0) As ������־
    From ����ҽ����¼ A, ������Ϣ B, ������ĿĿ¼ C
    Where a.����id = b.����id And a.������Ŀid = c.Id And a.Id = ҽ��id_In
    Group By Nvl(a.���id, a.Id), a.���, a.����id, a.�Һŵ�, a.Ӥ��, b.����, c.��������, a.�������, a.ҽ��״̬, a.ҽ������, a.����ҽ��, a.��ʼִ��ʱ��,
             a.ִ��ʱ�䷽��, a.Ƶ�ʴ���, a.Ƶ�ʼ��, a.�����λ, a.������־;
  r_Advice c_Advice%RowType;

  Cursor c_Pati(v_����id ������Ϣ.����id%Type) Is
    Select * From ������Ϣ Where ����id = v_����id;
  r_Pati c_Pati%RowType;

  --������ʱ���� 
  v_Temp     Varchar2(255);
  v_Count    Number;
  v_�������� ������ҳ.��������%Type;
  v_��Ա��� ��Ա��.���%Type;
  v_��Ա���� ��Ա��.����%Type;
  v_��Ժ��ʽ ��Ժ��ʽ.����%Type;
  d_��ʼʱ�� ����ҽ����¼.��ʼִ��ʱ��%Type;

  v_Error Varchar2(255);
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
  If �״�ʱ��_In Is Null Then
    Select ��ʼִ��ʱ�� Into d_��ʼʱ�� From ����ҽ����¼ Where ID = ҽ��id_In;
  End If;
  Open c_Advice;
    Fetch c_Advice
      Into r_Advice;
  --��һ��ҽ���ĵ�һ��ʱ����ҽ������ 
  If Nvl(First_In, 0) = 1 Then
   
  
    --����������� 
    --------------------------------------------------------------------------------------- 
    If Nvl(r_Advice.ҽ��״̬, 0) <> 1 Then
      v_Error := '"' || r_Advice.���� || '"��ҽ��"' || r_Advice.ҽ������ || '"�Ѿ��������˷��͡�' || Chr(13) || Chr(10) ||
                 '�ò��˵�ҽ������ʧ�ܡ������¶�ȡ�����嵥���ԡ�';
      Raise Err_Custom;
    End If;
  
    --���ͺ��ҽ������:�������ͺ��Զ�ֹͣ 
    --------------------------------------------------------------------------------------- 
    Update ����ҽ����¼
    Set ҽ��״̬ = 8, ִ����ֹʱ�� = ĩ��ʱ��_In,
        --����û�� 
        ͣ��ʱ�� = ����ʱ��_In,
        --Ҫ��Ϊ����ʱ����ʾ 
        ͣ��ҽ�� = v_��Ա���� --Ҫ��Ϊ��������ʾ,��ͬ��סԺ,����ҽ���޻�ʿ���� 
    Where ID = r_Advice.��id Or ���id = r_Advice.��id;
  
    Insert Into ����ҽ��״̬
      (ҽ��id, ��������, ������Ա, ����ʱ��)
      Select ID, 8, v_��Ա����, ����ʱ��_In From ����ҽ����¼ Where ID = r_Advice.��id Or ���id = r_Advice.��id;
  
    --����ҽ���Ĵ��� 
    --------------------------------------------------------------------------------------- 
    If r_Advice.������� = 'Z' And Nvl(r_Advice.��������, '0') <> '0' And Nvl(r_Advice.Ӥ��, 0) = 0 Then
      --1-����;2-סԺ; 
      If Instr(',1,2,', r_Advice.��������) > 0 And ִ�в���id_In Is Not Null Then
        --��������µ�ԤԼ�Ǽǵ�������1.��ǰ��ԤԼ,2.��ǰ����Ժ,3-��Ҫ��ԤԼʱ���ڵ�סԺ��¼ 
      
        --ɾ�������Һ���Ч������ԤԼ�Ǽ� 
        Begin
          Select Count(*) Into v_Count From ������ҳ Where ����id = r_Advice.����id And Nvl(��ҳid, 0) = 0;
        Exception
          When Others Then
            v_Count := 0;
        End;
        If Nvl(v_Count, 0) > 0 Then
          Zl_��Ժ������ҳ_Delete(r_Advice.����id, 0, 0, 0);
          v_Count := 0;
        End If;
      
        If v_Count = 0 Then
          Select Count(*) Into v_Count From ������ҳ Where ����id = r_Advice.����id And ��Ժ���� Is Null;
        End If;
        If v_Count = 0 Then
          Select Count(*)
          Into v_Count
          From ������ҳ
          Where ����id = r_Advice.����id And (��Ժ���� >= r_Advice.��ʼִ��ʱ�� Or ��Ժ���� >= r_Advice.��ʼִ��ʱ��);
        End If;
        If v_Count = 0 Then
          If r_Advice.�������� = '1' Then
            --����ҽ��,��������"��ʼʱ��"���۵��ٴ�ִ�п��� 
            Begin
              v_�������� := 2;
              Select Decode(�������, 1, 1, 2)
              Into v_��������
              From ��������˵��
              Where �������� = '�ٴ�' And ����id = ִ�в���id_In;
            Exception
              When Others Then
                Null;
            End;
          Elsif r_Advice.�������� = '2' Then
            --סԺҽ��,��������"��ʼʱ��"�Ǽǵ��ٴ�ִ�п��� 
            v_�������� := 0;
          End If;
        
          Open c_Pati(r_Advice.����id);
          Fetch c_Pati
            Into r_Pati;
        
          v_��Ժ��ʽ := Null;
          If r_Advice.������־ = 1 Then
            v_��Ժ��ʽ := '����';
          Else
            Select Decode(����, 1, '����', Null)
            Into v_��Ժ��ʽ
            From ���˹Һż�¼
            Where NO = r_Advice.�Һŵ� And ��¼���� = 1 And ��¼״̬ = 1;
          End If;
        
          If v_�������� = 1 Then
            Zl_��Ժ������ҳ_Insert(1, v_��������, r_Pati.����id, r_Pati.�����, Null, r_Pati.����, r_Pati.�Ա�, r_Pati.����, r_Pati.�ѱ�,
                             r_Pati.��������, r_Pati.����, r_Pati.����, r_Pati.ѧ��, r_Pati.����״��, r_Pati.ְҵ, r_Pati.���,
                             r_Pati.���֤��, r_Pati.�����ص�, r_Pati.��ͥ��ַ, r_Pati.��ͥ��ַ�ʱ�, r_Pati.��ͥ�绰, r_Pati.���ڵ�ַ,
                             r_Pati.���ڵ�ַ�ʱ�, r_Pati.��ϵ������, r_Pati.��ϵ�˹�ϵ, r_Pati.��ϵ�˵�ַ, r_Pati.��ϵ�˵绰, r_Pati.������λ,
                             r_Pati.��ͬ��λid, r_Pati.��λ�绰, r_Pati.��λ�ʱ�, r_Pati.��λ������, r_Pati.��λ�ʺ�, r_Pati.������, r_Pati.������,
                             r_Pati.��������, ִ�в���id_In, Null, Null, v_��Ժ��ʽ, Null, Null, r_Advice.����ҽ��, r_Pati.����, r_Pati.����,
                             r_Advice.��ʼִ��ʱ��, Null, Null, r_Pati.ҽ�Ƹ��ʽ, Null, Null, Null, Null, Null, Null, r_Pati.����,
                             v_��Ա���, v_��Ա����, 0, Null, Null, 0);
          Else
            Zl_��Ժ������ҳ_Insert(1, v_��������, r_Pati.����id, r_Pati.סԺ��, Null, r_Pati.����, r_Pati.�Ա�, r_Pati.����, r_Pati.�ѱ�,
                             r_Pati.��������, r_Pati.����, r_Pati.����, r_Pati.ѧ��, r_Pati.����״��, r_Pati.ְҵ, r_Pati.���,
                             r_Pati.���֤��, r_Pati.�����ص�, r_Pati.��ͥ��ַ, r_Pati.��ͥ��ַ�ʱ�, r_Pati.��ͥ�绰, r_Pati.���ڵ�ַ,
                             r_Pati.���ڵ�ַ�ʱ�, r_Pati.��ϵ������, r_Pati.��ϵ�˹�ϵ, r_Pati.��ϵ�˵�ַ, r_Pati.��ϵ�˵绰, r_Pati.������λ,
                             r_Pati.��ͬ��λid, r_Pati.��λ�绰, r_Pati.��λ�ʱ�, r_Pati.��λ������, r_Pati.��λ�ʺ�, r_Pati.������, r_Pati.������,
                             r_Pati.��������, ִ�в���id_In, Null, Null, v_��Ժ��ʽ, Null, Null, r_Advice.����ҽ��, r_Pati.����, r_Pati.����,
                             r_Advice.��ʼִ��ʱ��, Null, Null, r_Pati.ҽ�Ƹ��ʽ, Null, Null, Null, Null, Null, Null, r_Pati.����,
                             v_��Ա���, v_��Ա����, 0, Null, Null, 0);
          End If;
          Close c_Pati;
        End If;
      End If;
    End If;
  
   
  End If;
  Close c_Advice;
  --��д���ͼ�¼ 
  --------------------------------------------------------------------------------------- 
  Insert Into ����ҽ������
    (ҽ��id, ���ͺ�, ��¼����, NO, ��¼���, ��������, ������, ����ʱ��, ִ��״̬, ִ�в���id, �Ʒ�״̬, �״�ʱ��, ĩ��ʱ��, ��������, �������)
  Values
    (ҽ��id_In, ���ͺ�_In, ��¼����_In, No_In, ��¼���_In, ��������_In, v_��Ա����, ����ʱ��_In, ִ��״̬_In, ִ�в���id_In, �Ʒ�״̬_In,
     Nvl(�״�ʱ��_In, d_��ʼʱ��), Nvl(ĩ��ʱ��_In, d_��ʼʱ��), ��������_In, Decode(��¼����_In, 2, 1, Null));

  --�����ͼ��ҽ��ͬ��������ҽ���ļƷ�״̬   
  If �Ʒ�״̬_In = 1 And  r_Advice.��ID <> ҽ��id_In  And (r_Advice.������� = 'D' Or r_Advice.������� = 'F') Then   
     Update ����ҽ������ Set �Ʒ�״̬ = 1 Where ҽ��ID = r_Advice.��ID And ���ͺ� = ���ͺ�_In;
  End If;
  
  --�Զ���Ϊ��ִ��ʱ����Ҫͬ���������ִ��״̬����˻���״̬ 
  If ִ��״̬_In = 1 Then
    Zl_����ҽ��ִ��_Finish(ҽ��id_In, ���ͺ�_In, Null, Null, v_��Ա���, v_��Ա����, ִ�в���id_In);
  End If;
  --��Ϣ����
  Begin
    Execute Immediate 'Begin zl_������Ϣ_����(:1,:2); End;'
      Using 3, ���ͺ�_In;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ҽ������_Insert;
/

--93868:������,2016-03-04,���ʧЧԤԼ����
Create Or Replace Procedure Zl_�Һ����״̬_Delete
(
  ������ʽ_In Number := 0,
  �ű�_In     ���˹Һż�¼.�ű�%Type := Null
) As
  n_ԤԼ��Чʱ�� Number(5);
  n_ʧԼ���ڹҺ� Number(2);
  n_�Һ���Ч���� Number(5);
Begin
  If ������ʽ_In = 0 Then
    --�����ʷ��¼
    Delete �Һ����״̬ Where ���� < Trunc(Sysdate);
  Else
    --���ʧԼ��
    n_ԤԼ��Чʱ�� := Nvl(zl_GetSysParameter('ԤԼ��Чʱ��', 1111), 0);
    n_ʧԼ���ڹҺ� := Nvl(zl_GetSysParameter('ʧԼ���ڹҺ�', 1111), 0);
    n_�Һ���Ч���� := Nvl(zl_GetSysParameter('�Һ���Ч����'), 7);
    If n_ԤԼ��Чʱ�� <> 0 And n_ʧԼ���ڹҺ� <> 0 Then
      If �ű�_In Is Null Then
        For c_ʧЧԤԼ In (Select b.����, b.����, b.���
                       From ���˹Һż�¼ A, �Һ����״̬ B
                       Where a.ԤԼʱ�� - 1 / 24 / 60 * n_ԤԼ��Чʱ�� < Sysdate And a.ԤԼʱ�� > Sysdate - n_�Һ���Ч���� And a.��¼���� = 2 And
                             a.�ű� = b.���� And a.���� = b.���) Loop
          Delete From �Һ����״̬
          Where ���� = c_ʧЧԤԼ.���� And ��� = c_ʧЧԤԼ.��� And ״̬ = 2 And ���� = c_ʧЧԤԼ.����;
        End Loop;
      Else
        For c_ʧЧԤԼ In (Select b.����, b.����, b.���
                       From ���˹Һż�¼ A, �Һ����״̬ B
                       Where a.ԤԼʱ�� - 1 / 24 / 60 * n_ԤԼ��Чʱ�� < Sysdate And a.ԤԼʱ�� > Sysdate - n_�Һ���Ч���� And a.��¼���� = 2 And
                             a.�ű� = b.���� And a.���� = b.��� And a.�ű� = �ű�_In) Loop
          Delete From �Һ����״̬
          Where ���� = c_ʧЧԤԼ.���� And ��� = c_ʧЧԤԼ.��� And ״̬ = 2 And ���� = c_ʧЧԤԼ.����;
        End Loop;
      End If;
    End If;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�Һ����״̬_Delete;
/

--93820:����,2016-03-03,�������ۼۺ�ԭ���ֶδ���
Create Or Replace Procedure Zl_ҩƷ�շ���¼_Adjust
(
  Adjustid    In Number, --���ۼ�¼��ID
  Bln����     In Number := 0, --�Ƿ�תΪ�������ۣ�����2004-06-08���շ�ϸĿ�еı�ۣ�
  Billinfo_In In Varchar2 := Null, --����ʱ��ҩƷ�����ε��ۡ���ʽ:"����1,�ּ�1|����2,�ּ�2|....."
  ҩƷid_In   In Number := 0 --����Ϊ0ʱ��ʾ�ǳɱ��۵��ۣ��������ۼ��������
) As
  Classid      Number(18); --������
  v_Billno     ҩƷ�շ���¼.No%Type; --���۵���
  Rundate      Date; --������Чʱ��
  Blnrun       Number(1); --����ʱ�̵���
  Blncurprice  Number(1); --ʱ��ҩƷ
  LngϸĿid    Number(18); --�շ�ϸĿID
  Adjustdate   Date; --����ʱ��
  n_����       Number(18);
  n_�ּ�       �շѼ�Ŀ.�ּ�%Type;
  n_ԭ��       �շѼ�Ŀ.ԭ��%Type;
  v_Infotmp    Varchar2(4000);
  v_Fields     Varchar2(4000);
  n_���       Number(8);
  n_ԭ��id     �շѼ�Ŀ.ԭ��id%Type;
  n_������Ŀid �շѼ�Ŀ.������Ŀid%Type;
  n_���۽��   ҩƷ���.ʵ�ʽ��%Type;
  n_���ۼ�     ҩƷ���.���ۼ�%Type;
  n_�շ�id     ҩƷ�շ���¼.Id%Type;
  n_�䶯ԭ��   �շѼ�Ŀ.�䶯ԭ��%Type;

  Cursor c_Price --��ͨ����
  Is
    Select 1 ��¼״̬, 13 ����, v_Billno NO, Rownum ���, Classid ������id, m.Id As ҩƷid, s.����, Null ����, Null Ч��, s.�ϴβ��� As ����,
           1 ����, s.�ϴι�Ӧ��id As ��Ӧ��id, s.ʵ������ ��д����, 0 ʵ������, a.ԭ�� �ɱ���, 0 �ɱ����, a.�ּ� ���ۼ�, 0 ����, 'ҩƷ����' ժҪ, Zl_Username ������,
           Sysdate ��������, s.�ⷿid �ⷿid, 1 ���ϵ��, a.Id �۸�id, Nvl(m.�Ƿ���, 0) As ʱ��, s.ʵ�ʽ�� As �����, s.ʵ�ʲ�� As �����,
           Nvl(s.���ۼ�, Decode(Nvl(s.ʵ������, 0), 0, a.ԭ��, Nvl(s.ʵ�ʽ��, 0) / s.ʵ������)) As ԭ�ۼ�
    From ҩƷ��� S, �շ���ĿĿ¼ M, �շѼ�Ŀ A
    Where s.ҩƷid = m.Id And m.Id = a.�շ�ϸĿid And s.���� = 1 And a.�䶯ԭ�� = 0 And a.Id = Adjustid And a.ִ������ <= Sysdate;

  v_Data c_Price%RowType;

  Cursor c_ʱ�۰����ε��� --ʱ��ҩƷ�����ε���
  Is
    Select 1 ��¼״̬, 13 ����, v_Billno NO, n_��� + Rownum ���, Classid ������id, m.ҩƷid ҩƷid, s.���� ����, Null ����, Null Ч��,
           s.�ϴβ��� As ����, s.�ϴι�Ӧ��id As ��Ӧ��id, 1 ����, Nvl(s.ʵ������, 0) ��д����, 0 ʵ������, a.ԭ�� �ɱ���, 0 �ɱ����, n_�ּ� ���ۼ�, 0 ����,
           'ҩƷ����' ժҪ, Zl_Username ������, Sysdate ��������, s.�ⷿid �ⷿid, 1 ���ϵ��, a.Id �۸�id, Nvl(b.�Ƿ���, 0) As ʱ��,
           s.ʵ�ʽ�� As �����, s.ʵ�ʲ�� As �����, Nvl(s.���ۼ�, Decode(Nvl(s.ʵ������, 0), 0, a.ԭ��, Nvl(s.ʵ�ʽ��, 0) / s.ʵ������)) As ԭ�ۼ�
    From ҩƷ��� S, ҩƷĿ¼ M, �շѼ�Ŀ A, �շ���ĿĿ¼ B
    Where s.ҩƷid = b.Id And s.ҩƷid = m.ҩƷid And m.ҩƷid = a.�շ�ϸĿid And s.���� = 1 And a.�䶯ԭ�� = 0 And a.Id = Adjustid And
          a.ִ������ <= Sysdate And Nvl(s.����, 0) = n_����;

  v_ʱ�۰����ε��� c_ʱ�۰����ε���%RowType;
Begin
  If ҩƷid_In <> 0 Then
    --�ɱ��۵���
    Zl_ҩƷ�շ���¼_�ɱ��۵���(ҩƷid_In);
  Else
    n_�䶯ԭ�� := 0;
    --ȡ���ۼ�¼��Ч����
    Select �շ�ϸĿid, ִ������, ������Ŀid Into LngϸĿid, Rundate, n_������Ŀid From �շѼ�Ŀ Where ID = Adjustid;
  
    If Sysdate >= Rundate Then
      Blnrun := 1;
    Else
      Blnrun := 0;
    End If;
  
    If Blnrun = 1 Then
      --ȡ������ID
      Select ���id Into Classid From ҩƷ�������� Where ���� = 13;
    
      --ȡ����
      Select Nextno(147) Into v_Billno From Dual;
    
      --ȡ��ҩƷ�Ƿ���ʱ��ҩƷ
      Select Nvl(�Ƿ���, 0) Into Blncurprice From �շ���ĿĿ¼ Where ID = LngϸĿid;
    
      --����Ƿ����ԭ�ۺ��ּ���ͬ���������ͬʱ��ִ�е��ۼ۹��ܣ�����ɾ�������շѼ�Ŀ��¼���ָ�ԭ�����շѼ�Ŀ
      Begin
        Select ԭ��id Into n_ԭ��id From �շѼ�Ŀ Where ID = Adjustid And ԭ�� = �ּ� And ԭ��id Is Not Null;
      Exception
        When Others Then
          n_ԭ��id := 0;
      End;
    
      If n_ԭ��id > 0 Then
        --����ּ�=ԭ�ۣ�����������ǵ�������������Ŀ������������ĿID��ɾ�����ۼ�¼
        Delete �շѼ�Ŀ Where ID = Adjustid;
        Update �շѼ�Ŀ
        Set ������Ŀid = n_������Ŀid, ��ֹ���� = To_Date('3000-01-01', 'yyyy-mm-dd')
        Where ID = n_ԭ��id;
      Else
        Adjustdate := Sysdate;
      
        Begin
          Select �䶯ԭ�� Into n_�䶯ԭ�� From �շѼ�Ŀ Where ID = Adjustid And �䶯ԭ�� = 1;
        Exception
          When Others Then
            n_�䶯ԭ�� := 0;
        End;
        If n_�䶯ԭ�� = 0 Then
          If Billinfo_In = '' Or Billinfo_In Is Null Then
            For v_Data In c_Price Loop
              Select ҩƷ�շ���¼_Id.Nextval Into n_�շ�id From Dual;
              If Nvl(v_Data.��д����, 0) = 0 And (Nvl(v_Data.�����, 0) <> 0 Or Nvl(v_Data.�����, 0) <> 0) Then
                --����=0 ������<>0ʱֻ���¿����ж�Ӧ�����ۼ�,�������ۼ��������ݵ��ǽ���=0��ֻ��¼�����ۼۣ�����Ͳ�۲������            
                --��������Ӱ���¼
                Insert Into ҩƷ�շ���¼
                  (ID, ��¼״̬, ����, NO, ���, ������id, ҩƷid, ����, ����, Ч��, ����, ����, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ����, ժҪ, ������,
                   ��������, �ⷿid, ���ϵ��, �۸�id, �����, �������, ����, Ƶ��, ��ҩ��λid)
                Values
                  (n_�շ�id, v_Data.��¼״̬, v_Data.����, v_Data.No, v_Data.���, v_Data.������id, v_Data.ҩƷid, v_Data.����,
                   v_Data.����, v_Data.Ч��, v_Data.����, v_Data.����, v_Data.��д����, v_Data.ʵ������,
                   Decode(Blncurprice, 1, v_Data.ԭ�ۼ�, v_Data.�ɱ���), v_Data.�ɱ����, v_Data.���ۼ�, v_Data.����, v_Data.ժҪ,
                   v_Data.������, v_Data.��������, v_Data.�ⷿid, v_Data.���ϵ��, v_Data.�۸�id, Zl_Username, Adjustdate, v_Data.�����,
                   v_Data.�����, v_Data.��Ӧ��id);
              
                --���¿�����ۼ�,ֻ��ʱ�۷���ҩƷ���ܸ������ۼ��ֶ�
                Zl_ҩƷ���_Update(n_�շ�id);
              Else
                If Blncurprice = 1 Then
                  n_���ۼ� := v_Data.����� / v_Data.��д����;
                Else
                  n_���ۼ� := v_Data.�ɱ���;
                End If;
                n_���۽�� := Round((v_Data.���ۼ� - n_���ۼ�) * v_Data.��д����, 2);
              
                --��������Ӱ���¼
                Insert Into ҩƷ�շ���¼
                  (ID, ��¼״̬, ����, NO, ���, ������id, ҩƷid, ����, ����, Ч��, ����, ����, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ����, ���۽��, ���, ժҪ,
                   ������, ��������, �ⷿid, ���ϵ��, �۸�id, �����, �������, ����, Ƶ��, ��ҩ��λid)
                Values
                  (n_�շ�id, v_Data.��¼״̬, v_Data.����, v_Data.No, v_Data.���, v_Data.������id, v_Data.ҩƷid, v_Data.����,
                   v_Data.����, v_Data.Ч��, v_Data.����, v_Data.����, v_Data.��д����, v_Data.ʵ������,
                   Decode(Blncurprice, 1, v_Data.ԭ�ۼ�, v_Data.�ɱ���), v_Data.�ɱ����, v_Data.���ۼ�, v_Data.����, n_���۽��, n_���۽��,
                   v_Data.ժҪ, v_Data.������, v_Data.��������, v_Data.�ⷿid, v_Data.���ϵ��, v_Data.�۸�id, Zl_Username, Adjustdate,
                   v_Data.�����, v_Data.�����, v_Data.��Ӧ��id);
              
                --����ҩƷ���
                Zl_ҩƷ���_Update(n_�շ�id);
              End If;
            End Loop;
          Else
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
                If Nvl(v_ʱ�۰����ε���.��д����, 0) = 0 And (Nvl(v_ʱ�۰����ε���.�����, 0) <> 0 Or Nvl(v_ʱ�۰����ε���.�����, 0) <> 0) Then
                  --����=0 ������<>0ʱֻ���¿����ж�Ӧ�����ۼ�,�������ۼ��������ݵ��ǽ���=0��ֻ��¼�����ۼۣ�����Ͳ�۲������              
                  --��������Ӱ���¼
                  Insert Into ҩƷ�շ���¼
                    (ID, ��¼״̬, ����, NO, ���, ������id, ҩƷid, ����, ����, Ч��, ����, ����, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ����, ժҪ, ������,
                     ��������, �ⷿid, ���ϵ��, �۸�id, �����, �������, ����, Ƶ��, ��ҩ��λid)
                  Values
                    (n_�շ�id, v_ʱ�۰����ε���.��¼״̬, v_ʱ�۰����ε���.����, v_ʱ�۰����ε���.No, v_ʱ�۰����ε���.���, v_ʱ�۰����ε���.������id, v_ʱ�۰����ε���.ҩƷid,
                     v_ʱ�۰����ε���.����, v_ʱ�۰����ε���.����, v_ʱ�۰����ε���.Ч��, v_ʱ�۰����ε���.����, v_ʱ�۰����ε���.����, v_ʱ�۰����ε���.��д����,
                     v_ʱ�۰����ε���.ʵ������, Decode(Blncurprice, 1, v_ʱ�۰����ε���.ԭ�ۼ�, v_ʱ�۰����ε���.�ɱ���), v_ʱ�۰����ε���.�ɱ����, v_ʱ�۰����ε���.���ۼ�,
                     v_ʱ�۰����ε���.����, v_ʱ�۰����ε���.ժҪ, v_ʱ�۰����ε���.������, v_ʱ�۰����ε���.��������, v_ʱ�۰����ε���.�ⷿid, v_ʱ�۰����ε���.���ϵ��,
                     v_ʱ�۰����ε���.�۸�id, Zl_Username, Adjustdate, v_ʱ�۰����ε���.�����, v_ʱ�۰����ε���.�����, v_ʱ�۰����ε���.��Ӧ��id);
                  n_��� := n_��� + 1;
                
                  --���¿�����ۼ�,ֻ��ʱ�۷���ҩƷ���ܸ������ۼ��ֶ�
                  Zl_ҩƷ���_Update(n_�շ�id);
                Else
                
                  n_���ۼ�   := v_ʱ�۰����ε���.����� / v_ʱ�۰����ε���.��д����;
                  n_���۽�� := Round((n_�ּ� - n_���ۼ�) * v_ʱ�۰����ε���.��д����, 2);
                  --��������Ӱ���¼
                  Insert Into ҩƷ�շ���¼
                    (ID, ��¼״̬, ����, NO, ���, ������id, ҩƷid, ����, ����, Ч��, ����, ����, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ����, ���۽��, ���,
                     ժҪ, ������, ��������, �ⷿid, ���ϵ��, �۸�id, �����, �������, ����, Ƶ��, ��ҩ��λid)
                  Values
                    (n_�շ�id, v_ʱ�۰����ε���.��¼״̬, v_ʱ�۰����ε���.����, v_ʱ�۰����ε���.No, v_ʱ�۰����ε���.���, v_ʱ�۰����ε���.������id, v_ʱ�۰����ε���.ҩƷid,
                     v_ʱ�۰����ε���.����, v_ʱ�۰����ε���.����, v_ʱ�۰����ε���.Ч��, v_ʱ�۰����ε���.����, v_ʱ�۰����ε���.����, v_ʱ�۰����ε���.��д����,
                     v_ʱ�۰����ε���.ʵ������, Decode(Blncurprice, 1, v_ʱ�۰����ε���.ԭ�ۼ�, v_ʱ�۰����ε���.�ɱ���), v_ʱ�۰����ε���.�ɱ����, v_ʱ�۰����ε���.���ۼ�,
                     v_ʱ�۰����ε���.����, n_���۽��, n_���۽��, v_ʱ�۰����ε���.ժҪ, v_ʱ�۰����ε���.������, v_ʱ�۰����ε���.��������, v_ʱ�۰����ε���.�ⷿid,
                     v_ʱ�۰����ε���.���ϵ��, v_ʱ�۰����ε���.�۸�id, Zl_Username, Adjustdate, v_ʱ�۰����ε���.�����, v_ʱ�۰����ε���.�����,
                     v_ʱ�۰����ε���.��Ӧ��id);
                  n_��� := n_��� + 1;
                
                  --���¿��
                  Zl_ҩƷ���_Update(n_�շ�id);
                End If;
              End Loop;
            End Loop;
          End If;
        
          Update �շѼ�Ŀ Set �䶯ԭ�� = 1 Where ID = Adjustid;
        
          --����ҩƷĿ¼���շ�ϸĿ�еı��
          If Bln���� = 1 Then
            Update �շ���ĿĿ¼ Set �Ƿ��� = 0 Where ID = LngϸĿid;
            Update �շ�ϸĿ Set �Ƿ��� = 0 Where ID = LngϸĿid;
          End If;
        End If;
      End If;
    
      If n_�䶯ԭ�� = 0 Then
        --�ɱ��۵���
        Zl_ҩƷ�շ���¼_�ɱ��۵���(LngϸĿid, Rundate);
      End If;
    End If;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�շ���¼_Adjust;
/
--93853:������,2016-03-03,��ʷ���ݳ�ش�������
Create Or Replace Procedure Zl_Retu_Exes
(
  v_No   In Varchar2,
  n_Type In Number
) As
  --------------------------------------------
  --����:v_No,���ݺ���
  --     n_Type,��������:1-�շ�,2-����,3-�Զ�����,4-�Һ�,5-���￨,6-Ԥ��,7-����,8-δ����õĲ���id,��ҳID
  --------------------------------------------
  n_Allow  Number(1); --�Ƿ��ܹ����ݷ���
  n_Patiid Number(18);
  n_Pageid Number(5);
  Err_Item Exception;
  v_Err_Msg Varchar2(100);
  n_System  Number(5);
  n_ֻ��    Number(2);

  v_Table  Varchar2(100);
  v_Field  Varchar2(100);
  v_Sql    Varchar2(4000);
  v_Fields Varchar2(4000);

  --���ܣ���ȡ����ֶ��ַ���
  Function Getfields(v_Table In Varchar2) Return Varchar2 As
    v_Colstr Varchar2(4000);
  Begin
    Select f_List2str(Cast(Collect(Column_Name) As t_Strlist)) As Colsstr
    Into v_Colstr
    From (Select Column_Name From User_Tab_Columns Where Table_Name = v_Table Order By Column_Id);
  
    Return v_Colstr;
  End Getfields;

  --------------------------------------------
  --����ָ��ID�Ĳ���Ԥ����¼�ӹ���
  --------------------------------------------
  Procedure Zl_Retu_Prepay(n_Settle_Id H����Ԥ����¼.����id%Type) As
  Begin
    For r_Rec In (Select * From H����Ԥ����¼ Where ����id = n_Settle_Id) Loop
      v_Table  := '���˿��������';
      v_Fields := Getfields(v_Table);
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || replace(v_Fields,'��ת��','Null as ��ת��') || ' From H' || v_Table ||
                  ' Where Ԥ��id = :1';
      Execute Immediate v_Sql
        Using r_Rec.Id;
    
      v_Table  := '�������㽻��';
      v_Fields := Getfields(v_Table);
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || replace(v_Fields,'��ת��','Null as ��ת��') || ' From H' || v_Table ||
                  ' Where ����ID = :1';
      Execute Immediate v_Sql
        Using r_Rec.Id;
    
      v_Table  := '�����˿���Ϣ';
      v_Fields := Getfields(v_Table);
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || replace(v_Fields,'��ת��','Null as ��ת��') || ' From H' || v_Table ||
                  ' Where ��¼ID = :1 And ����ID = :2';
      Execute Immediate v_Sql
        Using r_Rec.Id, n_Settle_Id;
    
      v_Table  := '���˿������¼';
      v_Fields := Getfields(v_Table);
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || replace(v_Fields,'��ת��','Null as ��ת��') || ' From H' || v_Table ||
                  ' Where ID In(Select ������id From H���˿�������� Where Ԥ��id = :1)';
      Execute Immediate v_Sql
        Using r_Rec.Id;
    
      Delete H���˿������¼ Where ID In (Select Distinct ������id From H���˿�������� Where Ԥ��id = r_Rec.Id);
      Delete From H���˿�������� Where Ԥ��id = r_Rec.Id;
      Delete From H�������㽻�� Where ����id = r_Rec.Id;
      Delete From H�����˿���Ϣ Where ��¼id = r_Rec.Id And ����id = n_Settle_Id;
    End Loop;
  
    v_Table  := '����Ԥ����¼';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || replace(v_Fields,'��ת��','Null as ��ת��') || ' From H' || v_Table ||
                ' Where ����id = :1';
    Execute Immediate v_Sql
      Using n_Settle_Id;
  
    Delete H����Ԥ����¼ Where ����id = n_Settle_Id;
  End Zl_Retu_Prepay;

  --------------------------------------------
  --����ָ��ID�Ĳ��˷��ü�¼�ӹ���
  --------------------------------------------
  Procedure Zl_Retu_Fee(n_Settle_Id HסԺ���ü�¼.����id%Type) As
  Begin
    --���ز��˷��ü�¼
    v_Table  := 'סԺ���ü�¼';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || replace(v_Fields,'��ת��','Null as ��ת��') || ' From H' || v_Table ||
                ' Where ����id = :1';
    Execute Immediate v_Sql
      Using n_Settle_Id;
  
    v_Table  := '������ü�¼';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || replace(v_Fields,'��ת��','Null as ��ת��') || ' From H' || v_Table ||
                ' Where ����id = :1';
    Execute Immediate v_Sql
      Using n_Settle_Id;
  
    v_Table  := '���ò����¼';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || replace(v_Fields,'��ת��','Null as ��ת��') || ' From H' || v_Table ||
                ' Where ����id = :1';
    Execute Immediate v_Sql
      Using n_Settle_Id;
  
    v_Table  := 'ҽ��������ϸ';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || replace(v_Fields,'��ת��','Null as ��ת��') || ' From H' || v_Table ||
                ' Where ����id = :1';
    Execute Immediate v_Sql
      Using n_Settle_Id;
  
    --ɾ���ѷ��صķ��ü�¼
    Delete H������ü�¼ Where ����id = n_Settle_Id;
    Delete HסԺ���ü�¼ Where ����id = n_Settle_Id;
    Delete H���ò����¼ Where ����id = n_Settle_Id;
    Delete Hҽ��������ϸ Where ����id = n_Settle_Id;
  End Zl_Retu_Fee;

  --------------------------------------------
  --����ָ��ID��ҩƷ�շ���¼�ӹ���
  --------------------------------------------
  Procedure Zl_Retu_Medilist(n_Rec_Id HҩƷ�շ���¼.Id%Type) As
  Begin
    --���������˳�򷵻�ҩƷ�շ���ر������     
    v_Table  := 'ҩƷ�շ���¼';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || replace(v_Fields,'��ת��','Null as ��ת��') || ' From H' || v_Table ||
                ' Where ID = :1';
    Execute Immediate v_Sql
      Using n_Rec_Id;
  
    v_Table  := '��Һ��ҩ��¼';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || replace(v_Fields,'��ת��','Null as ��ת��') || ' From H' || v_Table ||
                ' Where ID In(Select ��¼ID From H��Һ��ҩ���� Where �շ�ID =:1)';
    Execute Immediate v_Sql
      Using n_Rec_Id;
  
    For P In (Select ID From H��Һ��ҩ��¼ Where ID In (Select ��¼id From H��Һ��ҩ���� Where �շ�id = n_Rec_Id)) Loop
      For R In (Select Column_Value From Table(f_Str2list('��Һ��ҩ����,��Һ��ҩ״̬'))) Loop
        v_Table := r.Column_Value;
        v_Field := 'ID';
      
        v_Fields := Getfields(v_Table);
        v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || replace(v_Fields,'��ת��','Null as ��ת��') || ' From H' || v_Table ||
                    ' Where ' || v_Field || ' = :1';
        Execute Immediate v_Sql
          Using p.Id;
      
        v_Sql := 'Delete H' || v_Table || ' Where ' || v_Field || ' = :1';
        Execute Immediate v_Sql
          Using p.Id;
      End Loop;
    End Loop;
  
    Delete H��Һ��ҩ��¼ Where ID In (Select ��¼id From H��Һ��ҩ���� Where �շ�id = n_Rec_Id);
  
    v_Table  := 'ҩƷǩ����¼';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || replace(v_Fields,'��ת��','Null as ��ת��') || ' From H' || v_Table ||
                ' Where ID In(Select ǩ��ID From HҩƷǩ����ϸ Where �շ�ID =:1)';
    Execute Immediate v_Sql
      Using n_Rec_Id;
    Delete HҩƷǩ����¼ Where ID In (Select ǩ��id From HҩƷǩ����ϸ Where �շ�id = n_Rec_Id);
  
    For R In (Select Column_Value From Table(f_Str2list('�շ���¼������Ϣ,��Һ��ҩ����,ҩƷǩ����ϸ,ҩƷ����ƻ�'))) Loop
      v_Table := r.Column_Value;
      If v_Table = 'ҩƷ����ƻ�' Then
        v_Field := '����ID';
      Else
        v_Field := '�շ�ID';
      End If;
    
      v_Fields := Getfields(v_Table);
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || replace(v_Fields,'��ת��','Null as ��ת��') || ' From H' || v_Table ||
                  ' Where ' || v_Field || ' = :1';
      Execute Immediate v_Sql
        Using n_Rec_Id;
    
      v_Sql := 'Delete From H' || v_Table || ' Where ' || v_Field || ' = :1';
      Execute Immediate v_Sql
        Using n_Rec_Id;
    End Loop;
  
    --ɾ���ѷ��ص�ҩƷ�շ���¼
    Delete HҩƷ�շ���¼ Where ID = n_Rec_Id;
  End Zl_Retu_Medilist;

  --------------------------------------------
  --����Ϊ��������
  --------------------------------------------
Begin
  ----------------------------------------------------------------------------------------------------------
  --���˺�:��Ҫ�ǶԻ�����ͼ����ͼ��ת������������ֻ���ж�.
  Select ��� Into n_System From zlSystems Where Upper(������) = Zl_Owner And ��� Like '1%';
  Begin
    Select Nvl(ֻ��, 0) Into n_ֻ�� From zlBakSpaces Where ϵͳ = n_System And ��ǰ = 1;
  Exception
    When Others Then
      v_Err_Msg := '[ZLSOFT]��ǰû�п��õ���ʷ���ݿռ�,���ܼ���![ZLSOFT]';
      Raise Err_Item;
  End;
  If n_ֻ�� = 1 Then
    v_Err_Msg := '[ZLSOFT]��ʷ���ݿռ�Ŀǰ��״̬Ϊֻ��,���ܼ���![ZLSOFT]';
    Raise Err_Item;
  End If;

  If n_Type = 8 Then
    --8-���ָ�����˵�δ����ʷ���
    --�ų����������Ϻ󣬼��ʵ����ʵļ�¼����¼״̬Ϊ2��û�н���ID����¼״̬Ϊ3���н���ID��) 
    If Instr(v_No, ',') = 0 Then
      --a.������ID������ﲡ�˵�δ����ʷ���
      For Rno In (Select NO, ��¼����
                  From H������ü�¼ A
                  Where ����id = To_Number(v_No) And a.���ʷ��� = 1 And ����id Is Null And Not Exists
                   (Select 1
                         From H������ü�¼ B
                         Where a.No = b.No And a.��¼���� = b.��¼���� And a.��� = b.��� And b.��¼״̬ = 3 And b.����id Is Not Null)) Loop
        v_Table  := '������ü�¼';
        v_Fields := Getfields(v_Table);
        v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' ||
                    Replace(v_Fields, '��ת��', 'Null as ��ת��') || ' From H' || v_Table || ' Where NO = :1 And ��¼���� = :2';
        Execute Immediate v_Sql
          Using Rno.No, Rno.��¼����;
      
        For r_Rxlist In (Select m.Id
                         From HҩƷ�շ���¼ M, H������ü�¼ E
                         Where e.No = Rno.No And e.��¼���� = Rno.��¼���� And e.�շ���� In ('4', '5', '6', '7') And m.����id = e.Id And
                               m.���� In (9, 10, 25, 26)) Loop
          Zl_Retu_Medilist(r_Rxlist.Id);
        End Loop;
        Delete H������ü�¼ Where NO = Rno.No And ��¼���� = Rno.��¼����;
      End Loop;
    Else
      --b.������ID,��ҳID���סԺ���˵�δ����ʷ���
      n_Patiid := Substr(v_No, 1, Instr(v_No, ',') - 1);
      n_Pageid := Substr(v_No, Instr(v_No, ',') + 1);
    
      For Rno In (Select NO, ��¼����
                  From HסԺ���ü�¼ A
                  Where ����id = n_Patiid And ��ҳid = n_Pageid And a.���ʷ��� = 1 And ����id Is Null And Not Exists
                   (Select 1
                         From HסԺ���ü�¼ B
                         Where a.No = b.No And a.��¼���� = b.��¼���� And a.��� = b.��� And b.��¼״̬ = 3 And b.����id Is Not Null)) Loop
        v_Table  := 'סԺ���ü�¼';
        v_Fields := Getfields(v_Table);
        v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' ||
                    Replace(v_Fields, '��ת��', 'Null as ��ת��') || ' From H' || v_Table || ' Where NO = :1 And ��¼���� = :2';
        Execute Immediate v_Sql
          Using Rno.No, Rno.��¼����;
      
        For r_Rxlist In (Select m.Id
                         From HҩƷ�շ���¼ M, HסԺ���ü�¼ E
                         Where e.No = Rno.No And e.��¼���� = Rno.��¼���� And e.�շ���� In ('4', '5', '6', '7') And m.����id = e.Id And
                               m.���� In (9, 10, 25, 26)) Loop
          Zl_Retu_Medilist(r_Rxlist.Id);
        End Loop;
        Delete HסԺ���ü�¼ Where NO = Rno.No And ��¼���� = Rno.��¼����;
      End Loop;
    End If;
  Else
    --�ж��Ƿ��ܰ��յ��ݷ���
    Select Decode(Sum(Nvl(p.���, 0)) - Sum(Nvl(p.��Ԥ��, 0)), Null, 1, 0, 1, 0)
    Into n_Allow
    From H����Ԥ����¼ P,
         (Select ����id
           From H������ü�¼
           Where NO = v_No And (1 = n_Type And ��¼���� = 1 Or 2 = n_Type And ��¼���� = 2 Or 3 = n_Type And ��¼���� = 3 Or
                 4 = n_Type And ��¼���� = 4 Or 5 = n_Type And ��¼���� = 5)
           Union
           Select ����id
           From HסԺ���ü�¼
           Where NO = v_No And (1 = n_Type And ��¼���� = 1 Or 2 = n_Type And ��¼���� = 2 Or 3 = n_Type And ��¼���� = 3 Or
                 4 = n_Type And ��¼���� = 4 Or 5 = n_Type And ��¼���� = 5)
           Union
           Select ����id
           From H����Ԥ����¼
           Where NO = v_No And 6 = n_Type And ��¼���� In (1, 11)
           Union
           Select ID
           From H���˽��ʼ�¼
           Where NO = v_No And 7 = n_Type) L
    Where p.����id = l.����id And p.��¼���� In (1, 11);
    If n_Allow = 1 Then
      Select Decode(Sum(Nvl(e.ʵ�ս��, 0)) - Sum(Nvl(e.���ʽ��, 0)), Null, 1, 0, 1, 0)
      Into n_Allow
      From (Select e.ʵ�ս��, e.���ʽ��
             From H������ü�¼ E,
                  (Select ����id
                    From H������ü�¼
                    Where NO = v_No And (1 = n_Type And ��¼���� = 1 Or 2 = n_Type And ��¼���� = 2 Or 3 = n_Type And ��¼���� = 3 Or
                          4 = n_Type And ��¼���� = 4 Or 5 = n_Type And ��¼���� = 5)
                    Union
                    Select ����id
                    From H����Ԥ����¼
                    Where NO = v_No And 6 = n_Type And ��¼���� In (1, 11)
                    Union
                    Select ID
                    From H���˽��ʼ�¼
                    Where NO = v_No And 7 = n_Type) L
             Where e.����id = l.����id
             Union All
             Select e.ʵ�ս��, e.���ʽ��
             From HסԺ���ü�¼ E,
                  (Select ����id
                    From HסԺ���ü�¼
                    Where NO = v_No And (1 = n_Type And ��¼���� = 1 Or 2 = n_Type And ��¼���� = 2 Or 3 = n_Type And ��¼���� = 3 Or
                          4 = n_Type And ��¼���� = 4 Or 5 = n_Type And ��¼���� = 5)
                    Union
                    Select ����id
                    From H����Ԥ����¼
                    Where NO = v_No And 6 = n_Type And ��¼���� In (1, 11)
                    Union
                    Select ID
                    From H���˽��ʼ�¼
                    Where NO = v_No And 7 = n_Type) L
             Where e.����id = l.����id) E;
    End If;
  
    --���յ��ݻ��˻�ȡ�����α귵��
    If n_Allow = 1 Then
      For r_Settle In (Select ����id
                       From H������ü�¼
                       Where NO = v_No And
                             (1 = n_Type And ��¼���� = 1 Or 2 = n_Type And ��¼���� = 2 Or 3 = n_Type And ��¼���� = 3 Or
                             4 = n_Type And ��¼���� = 4 Or 5 = n_Type And ��¼���� = 5)
                       Union
                       Select ����id
                       From HסԺ���ü�¼
                       Where NO = v_No And
                             (1 = n_Type And ��¼���� = 1 Or 2 = n_Type And ��¼���� = 2 Or 3 = n_Type And ��¼���� = 3 Or
                             4 = n_Type And ��¼���� = 4 Or 5 = n_Type And ��¼���� = 5)
                       Union
                       Select ����id
                       From H����Ԥ����¼
                       Where NO = v_No And 6 = n_Type And ��¼���� In (1, 11)
                       Union All
                       Select ID
                       From H���˽��ʼ�¼
                       Where NO = v_No And 7 = n_Type) Loop
      
        v_Table  := '���˽��ʼ�¼';
        v_Fields := Getfields(v_Table);
        v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || replace(v_Fields,'��ת��','Null as ��ת��') || ' From H' || v_Table ||
                    ' Where id = :1';
        Execute Immediate v_Sql
          Using r_Settle.����id;
      
        Zl_Retu_Prepay(r_Settle.����id);
        For r_Rxlist In (Select m.Id
                         From HҩƷ�շ���¼ M,
                              (Select ID, NO, ���, ��¼����
                                From H������ü�¼
                                Where ����id = r_Settle.����id And �շ���� In ('4', '5', '6', '7') And ��¼���� In (1, 2)) E
                         Where m.No = e.No And m.����id = e.Id And
                               (e.��¼���� = 1 And m.���� In (8, 24) Or e.��¼���� <> 1 And m.���� In (9, 10, 25, 26))
                         Union All
                         Select m.Id
                         From HҩƷ�շ���¼ M,
                              (Select ID, NO, ���, ��¼����
                                From HסԺ���ü�¼
                                Where ����id = r_Settle.����id And �շ���� In ('4', '5', '6', '7') And ��¼���� In (1, 2)) E
                         Where m.No = e.No And m.����id = e.Id And
                               (e.��¼���� = 1 And m.���� In (8, 24) Or e.��¼���� <> 1 And m.���� In (9, 10, 25, 26))) Loop
          Zl_Retu_Medilist(r_Rxlist.Id);
        End Loop;
        Zl_Retu_Fee(r_Settle.����id);
      
        Delete H���˽��ʼ�¼ Where ID = r_Settle.����id;
      End Loop;
    Else
      Begin
        --n_Type,��������:1-�շ�,2-����,3-�Զ�����,4-�Һ�,5-���￨,6-Ԥ��,7-����
        If n_Type = 7 Then
          Select Distinct ����id Into n_Patiid From H���˽��ʼ�¼ Where NO = v_No;
        Elsif n_Type = 6 Then
          Select Distinct ����id
          Into n_Patiid
          From H����Ԥ����¼
          Where NO = v_No And 6 = n_Type And ��¼���� In (1, 11);
        Elsif n_Type = 5 Or n_Type = 3 Then
          Select Distinct ����id
          Into n_Patiid
          From HסԺ���ü�¼
          Where NO = v_No And (1 = n_Type And ��¼���� = 1 Or 2 = n_Type And ��¼���� = 2 Or 3 = n_Type And ��¼���� = 3 Or
                4 = n_Type And ��¼���� = 4 Or 5 = n_Type And ��¼���� = 5);
        Elsif n_Type = 4 Or n_Type = 1 Then
          If n_Type = 1 Then
            Select Distinct ����id
            Into n_Patiid
            From (Select Distinct ����id
                   From H������ü�¼
                   Where NO = v_No And ��¼���� = 1
                   Union All
                   Select Distinct ����id
                   From H���ò����¼
                   Where NO = v_No And ��¼���� = 1)
            Where Rownum < 2;
          
          Else
            Select Distinct ����id
            Into n_Patiid
            From H������ü�¼
            Where NO = v_No And (1 = n_Type And ��¼���� = 1 Or 2 = n_Type And ��¼���� = 2 Or 3 = n_Type And ��¼���� = 3 Or
                  4 = n_Type And ��¼���� = 4 Or 5 = n_Type And ��¼���� = 5);
          End If;
        Else
          Begin
            Select Distinct ����id
            Into n_Patiid
            From HסԺ���ü�¼
            Where NO = v_No And (1 = n_Type And ��¼���� = 1 Or 2 = n_Type And ��¼���� = 2 Or 3 = n_Type And ��¼���� = 3 Or
                  4 = n_Type And ��¼���� = 4 Or 5 = n_Type And ��¼���� = 5);
          Exception
            When Others Then
              n_Patiid := -1;
          End;
          If Nvl(n_Patiid, 0) <= 0 Then
            Select Distinct ����id
            Into n_Patiid
            From H������ü�¼
            Where NO = v_No And (1 = n_Type And ��¼���� = 1 Or 2 = n_Type And ��¼���� = 2 Or 3 = n_Type And ��¼���� = 3 Or
                  4 = n_Type And ��¼���� = 4 Or 5 = n_Type And ��¼���� = 5);
          End If;
        End If;
      Exception
        When Others Then
          n_Patiid := Null;
      End Zl_Patiid;
    
      For r_Settle In (Select Distinct ����id
                       From H������ü�¼
                       Where ����id = n_Patiid
                       Union
                       Select Distinct ����id
                       From HסԺ���ü�¼
                       Where ����id = n_Patiid
                       Union
                       Select Distinct ����id
                       From H���ò����¼
                       Where ����id = n_Patiid) Loop
      
        v_Table  := '���˽��ʼ�¼';
        v_Fields := Getfields(v_Table);
        v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || replace(v_Fields,'��ת��','Null as ��ת��') || ' From H' || v_Table ||
                    ' Where id = :1';
        Execute Immediate v_Sql
          Using r_Settle.����id;
      
        Zl_Retu_Prepay(r_Settle.����id);
        For r_Rxlist In (Select m.Id
                         From HҩƷ�շ���¼ M,
                              (Select ID, NO, ���, ��¼����
                                From H������ü�¼
                                Where ����id = r_Settle.����id And �շ���� In ('4', '5', '6', '7') And ��¼���� In (1, 2)) E
                         Where m.No = e.No And m.����id = e.Id And
                               (e.��¼���� = 1 And m.���� In (8, 24) Or e.��¼���� <> 1 And m.���� In (9, 10, 25, 26))
                         Union All
                         Select m.Id
                         From HҩƷ�շ���¼ M,
                              (Select ID, NO, ���, ��¼����
                                From HסԺ���ü�¼
                                Where ����id = r_Settle.����id And �շ���� In ('4', '5', '6', '7') And ��¼���� In (1, 2)) E
                         Where m.No = e.No And m.����id = e.Id And
                               (e.��¼���� = 1 And m.���� In (8, 24) Or e.��¼���� <> 1 And m.���� In (9, 10, 25, 26))
                         
                         ) Loop
          Zl_Retu_Medilist(r_Rxlist.Id);
        End Loop;
      
        Zl_Retu_Fee(r_Settle.����id);
        Delete H���˽��ʼ�¼ Where ID = r_Settle.����id;
      
      End Loop;
    End If;
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM || ':' || v_Sql);
End Zl_Retu_Exes;
/

--93846:������,2016-03-03,���֧���ж�
Create Or Replace Function Zl_Get_Threecardtypeid
(
  ģ���_In Number,
  ����id_In ������ü�¼.����id%Type,
  ҽ��id_In ����ҽ����¼.Id%Type
) Return Number Is
  ------------------------------------------------------------------------------------------- 
  --����:��ȡ��ǰ�����ӿ�֧���Ŀ����ID 
  --���: ģ���_IN-������ģ��� 
  --      ����ID_In-��ǰ����Ĳ���ID 
  --      ҽ��id_In-ҽ��ID
  --����:�����ʻ�֧���Ŀ����ID,��������˿����ID,���ٽ���ˢ������,��ֱ�ӵ�������֧���ӿ� 
  --˵��: 
  --  1.����ҽ��վ����: Ŀǰ��ʱ������ҽ������վ����,��Ҫ������֧���Ĳ������ò���Ч 
  --  2.����ִ�п���:��δ���øù���,������Ҫ������չ 
  --  3.�˹���,������û��ľ���ҵ��������з��� 
  ------------------------------------------------------------------------------------------- 
Begin
  If Nvl(ģ���_In, 0) = 1260 Then
    --����ҽ������վ 
    If Nvl(����id_In, 0) = 0 And Nvl(ҽ��id_In, 0) = 0 Then
      Return Null;
    End If;
    Return Null;
  End If;
  Return Null;
End Zl_Get_Threecardtypeid;
/

--93853:������,2016-03-03,��ʷ���ݳ�ش�������
--93801:����,2016-03-02,��ʷ���ݳ�صĴ���
Create Or Replace Procedure Zl_Retu_Clinic
(
  n_Patiid In Number,
  v_Times  In Varchar2,
  n_Flag   In Number
) As
  --------------------------------------------
  --����:n_Patiid,����id
  --     v_Times,�Һŵ��Ż�סԺ��ҳid�����ʱ���Һŵ�����쵥�ţ�
  --     n_Flag,�����סԺ��־:0-����,1-סԺ,2-��죨��ʱ��ֻ��n_Patiid������Ч��
  --------------------------------------------
  Err_Item Exception;
  v_Err_Msg    Varchar2(100);
  n_System     Number(5);
  n_Opersystem Number(5);
  n_ֻ��       Number(2);

  v_Table    Varchar2(100);
  v_Subtable Varchar2(100);
  v_Field    Varchar2(100);
  v_Subfield Varchar2(100);
  v_Sql      Varchar2(4000);
  v_Sqlchild Varchar2(4000);
  v_Fields   Varchar2(4000);

  --���ܣ���ȡ����ֶ��ַ���
  Function Getfields(v_Table In Varchar2) Return Varchar2 As
    v_Colstr Varchar2(4000);
  Begin
    Select f_List2str(Cast(Collect(Column_Name) As t_Strlist)) As Colsstr
    Into v_Colstr
    From (Select Column_Name From User_Tab_Columns Where Table_Name = v_Table Order By Column_Id);
  
    Return v_Colstr;
  End Getfields;

  --------------------------------------------
  --����ָ������ID����ҳ����ر���ӹ���
  --------------------------------------------
  Procedure Zl_Retu_Other
  (
    n_Pati_Id ������ҳ.����id%Type,
    n_Page_Id ������ҳ.��ҳid%Type
  ) As
  
  Begin
  
    For R In (Select Column_Value From Table(f_Str2list('���˹�����¼,������ϼ�¼,���������¼'))) Loop
      v_Table  := r.Column_Value;
      v_Fields := Getfields(v_Table);
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' ||
                  Replace(v_Fields, '��ת��', 'Null as ��ת��') || ' From H' || v_Table || ' Where ����id = :1 And ��ҳid = :2';
      Execute Immediate v_Sql
        Using n_Pati_Id, n_Page_Id;
    
      v_Sql := 'Delete From H' || v_Table || ' Where ����id = :1 And ��ҳid = :2';
      Execute Immediate v_Sql
        Using n_Pati_Id, n_Page_Id;
    End Loop;
  End Zl_Retu_Other;

  --------------------------------------------
  --����ָ������ID����ҳ���ٴ�·����ر���ӹ���
  --------------------------------------------
  Procedure Zl_Retu_Path
  (
    n_Pati_Id ������ҳ.����id%Type,
    n_Page_Id ������ҳ.��ҳid%Type
  ) As
  Begin
    v_Table  := '�����ٴ�·��';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '��ת��', 'Null as ��ת��') ||
                ' From H' || v_Table || ' Where ����id = :1 And ��ҳid = :2 ';
    Execute Immediate v_Sql
      Using n_Pati_Id, n_Page_Id;
  
    --����·��ҽ�����ڲ���ҽ����¼ת��֮��ִ��
    For P In (Select ID As ·����¼id From H�����ٴ�·�� Where ����id = n_Pati_Id And ��ҳid = n_Page_Id) Loop
      For R In (Select Column_Value
                From Table(f_Str2list('����·��ִ��,���˺ϲ�·��,����·������,����·������,����·��ָ��,���˺ϲ�·������,���˳�����¼'))) Loop
        v_Table := r.Column_Value;
        If v_Table = '���˺ϲ�·��' Then
          v_Field := '��Ҫ·����¼id';
        Else
          v_Field := '·����¼id';
        End If;
      
        v_Fields := Getfields(v_Table);
        v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' ||
                    Replace(v_Fields, '��ת��', 'Null as ��ת��') || ' From H' || v_Table || ' Where ' || v_Field || ' = :1';
        Execute Immediate v_Sql
          Using p.·����¼id;
      
        v_Sql := 'Delete H' || v_Table || ' Where ' || v_Field || ' = :1';
        Execute Immediate v_Sql
          Using p.·����¼id;
      End Loop;
    End Loop;
  
    Delete H�����ٴ�·�� Where ����id = n_Pati_Id And ��ҳid = n_Page_Id;
  End Zl_Retu_Path;

  --------------------------------------------
  --����ָ������ID����ҳ�Ļ�����ر���ӹ���
  --------------------------------------------
  Procedure Zl_Retu_Tend
  (
    n_Pati_Id ������ҳ.����id%Type,
    n_Page_Id ������ҳ.��ҳid%Type
  ) As
  Begin
  
    v_Table  := '���˻����ļ�';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '��ת��', 'Null as ��ת��') ||
                ' From H' || v_Table || ' Where ����id = :1 And ��ҳid = :2 ';
    Execute Immediate v_Sql
      Using n_Pati_Id, n_Page_Id;
  
    For P In (Select ID As �ļ�id From H���˻����ļ� Where ����id = n_Pati_Id And ��ҳid = n_Page_Id) Loop
      For R In (Select Column_Value
                From Table(f_Str2list('���˻�������,���˻����ӡ,���˻�����Ŀ,���˻���Ҫ������,����Ҫ������'))) Loop
        v_Table  := r.Column_Value;
        v_Fields := Getfields(v_Table);
        v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' ||
                    Replace(v_Fields, '��ת��', 'Null as ��ת��') || ' From H' || v_Table || ' Where �ļ�id = :1';
        Execute Immediate v_Sql
          Using p.�ļ�id;
      
        If v_Table = '���˻�������' Then
          v_Fields := Getfields('���˻�����ϸ');
          v_Sql    := 'Insert Into ���˻�����ϸ(' || v_Fields || ') Select ' || Replace(v_Fields, '��ת��', 'Null as ��ת��') ||
                      ' From H���˻�����ϸ Where ��¼id In (Select ID From H���˻������� Where �ļ�id = :1)';
          Execute Immediate v_Sql
            Using p.�ļ�id;
        
          v_Sql := 'Delete H���˻�����ϸ Where ��¼id In (Select ID From H���˻������� Where �ļ�id = :1)';
          Execute Immediate v_Sql
            Using p.�ļ�id;
        End If;
      
        v_Sql := 'Delete H' || v_Table || ' Where �ļ�id = :1';
        Execute Immediate v_Sql
          Using p.�ļ�id;
      End Loop;
    End Loop;
  
    Delete H���˻����ļ� Where ����id = n_Pati_Id And ��ҳid = n_Page_Id;
  
    --�ϰ滤��ϵͳ����
    ------------------------------------------------------------------------
    v_Table  := '���˻����¼';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '��ת��', 'Null as ��ת��') ||
                ' From H' || v_Table || ' Where ����id = :1 And ��ҳid = :2 ';
    Execute Immediate v_Sql
      Using n_Pati_Id, n_Page_Id;
  
    For P In (Select ID From H���˻����¼ Where ����id = n_Pati_Id And ��ҳid = n_Page_Id) Loop
      v_Table  := '���˻�������';
      v_Fields := Getfields(v_Table);
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' ||
                  Replace(v_Fields, '��ת��', 'Null as ��ת��') || ' From H' || v_Table || ' Where ��¼ID = :1';
      Execute Immediate v_Sql
        Using p.Id;
    
      v_Sql := 'Delete H' || v_Table || ' Where ��¼ID = :1';
      Execute Immediate v_Sql
        Using p.Id;
    End Loop;
  
    Delete H���˻����¼ Where ����id = n_Pati_Id And ��ҳid = n_Page_Id;
  End Zl_Retu_Tend;

  --------------------------------------------
  --����ָ��ID�Ĳ����°���Ӳ�����¼�ӹ���
  --------------------------------------------
  Procedure Zl_Retu_Epr(n_Rec_Id H���Ӳ�����¼.Id%Type) As
    v_Field Varchar(100);
  Begin
    v_Table  := '���Ӳ�����¼';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '��ת��', 'Null as ��ת��') ||
                ' From H' || v_Table || ' Where id = :1';
    Execute Immediate v_Sql
      Using n_Rec_Id;
  
    --������ϼ�¼��Zl_Retu_Other����ת�أ��޲���ID�����
    --Ӱ�񱨸沵��,����ҽ������,������ļ�¼,�⼸�ű��������Zl_Retu_Order��ת��ҽ�����ٴ���
    For R In (Select Column_Value
              From Table(f_Str2list('���Ӳ�������,���Ӳ�����ʽ,���Ӳ�������,�����걨��¼,�������淴��'))) Loop
      v_Table := r.Column_Value;
      If v_Table = '���Ӳ�������' Then
        v_Field := '����id';
      Else
        v_Field := '�ļ�id';
      End If;
      v_Fields := Getfields(v_Table);
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' ||
                  Replace(v_Fields, '��ת��', 'Null as ��ת��') || ' From H' || v_Table || ' Where ' || v_Field || ' = :1';
      Execute Immediate v_Sql
        Using n_Rec_Id;
    
      If v_Table = '���Ӳ�������' Then
        v_Fields := Getfields('���Ӳ���ͼ��');
        v_Sql    := 'Insert Into ���Ӳ���ͼ��(' || v_Fields || ') Select ' || Replace(v_Fields, '��ת��', 'Null as ��ת��') ||
                    ' From H���Ӳ���ͼ�� Where ����id In (Select ID From H���Ӳ������� Where �ļ�id = :1 And �������� = 5)';
        Execute Immediate v_Sql
          Using n_Rec_Id;
      
        Delete H���Ӳ���ͼ�� Where ����id In (Select ID From H���Ӳ������� Where �ļ�id = n_Rec_Id And �������� = 5);
      End If;
    
      v_Sql := 'Delete H' || v_Table || ' Where ' || v_Field || ' = :1';
      Execute Immediate v_Sql
        Using n_Rec_Id;
    End Loop;
  
    Delete H���Ӳ�����¼ Where ID = n_Rec_Id;
  End Zl_Retu_Epr;
  --------------------------------------------
  --����ָ��ID�Ĳ���ҽ����¼�ӹ��̣������ڲ������ٴ�·��ת��֮��ִ��(����ҽ������,Ӱ�񱨸沵�أ�����·��ҽ��)
  --��Zl_Retu_Other����ת����"������ϼ�¼",ת��"�������ҽ��"ʱ������ת
  --------------------------------------------
  Procedure Zl_Retu_Order(n_Rec_Id H����ҽ����¼.Id%Type) As
  Begin
    v_Table  := '����ҽ����¼';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '��ת��', 'Null as ��ת��') ||
                ' From H' || v_Table || ' Where id = :1';
    Execute Immediate v_Sql
      Using n_Rec_Id;
  
    --��"ҽ��ID,���ͺ�"Ϊ����ģ�����ҽ��IDֱ��ת�أ�ֻ��Ҫ����"����ҽ������"֮�󼴿�
    --���������ϵ��"������ļ�¼"����"����ҽ������"����
    For P In (Select Column_Value
              From Table(f_Str2list('����ҽ���Ƽ�,����ҽ��״̬,����ҽ������,����ҽ������,����ҽ������,����ҽ��ִ��,����ҽ����ӡ,��Ѫ�����¼,��Ѫ������,' ||
                                     'ҽ��ִ�д�ӡ,ҽ��ִ��ʱ��,ҽ��ִ�мƼ�,ִ�д�ӡ��¼,�������ҽ��,����·��ҽ��,����ҽ������,������ļ�¼,' ||
                                     'Ӱ�񱨸沵��,Ӱ�񱨸��¼,Ӱ�񱨸������¼,Ӱ�����¼,Ӱ�����뵥ͼ��,Ӱ���ղ�����,Ӱ��Σ��ֵ��¼,����걾��¼,�����Լ���¼,������ռ�¼'))) Loop
      v_Table := p.Column_Value;
      If Instr('����·��ҽ��', v_Table) > 0 Then
        v_Field := '����ҽ��ID';
      Else
        v_Field := 'ҽ��ID';
      End If;
    
      v_Fields := Getfields(v_Table);
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' ||
                  Replace(v_Fields, '��ת��', 'Null as ��ת��') || ' From H' || v_Table || ' Where ' || v_Field || ' = :1';
      If v_Table = '����ҽ��״̬' Then
        v_Sqlchild := v_Sql;
      Else
        Execute Immediate v_Sql
          Using n_Rec_Id;
      End If;
    
      If v_Table = '����ҽ��״̬' Then
        v_Fields := Getfields('ҽ��ǩ����¼');
        v_Sql    := 'Insert Into ҽ��ǩ����¼(' || v_Fields || ') Select ' || Replace(v_Fields, '��ת��', 'Null as ��ת��') ||
                    ' From Hҽ��ǩ����¼ Where ID In (Select ǩ��id From H����ҽ��״̬ Where ҽ��id = :1 And ǩ��id Is Not Null)';
        Execute Immediate v_Sql
          Using n_Rec_Id;
        Delete Hҽ��ǩ����¼
        Where ID In (Select ǩ��id From H����ҽ��״̬ Where ҽ��id = n_Rec_Id And ǩ��id Is Not Null);
      
        Execute Immediate v_Sqlchild
          Using n_Rec_Id;
      
      Elsif v_Table = '����ҽ������' Then
        v_Fields := Getfields('���Ƶ��ݴ�ӡ');
        v_Sql    := 'Insert Into ���Ƶ��ݴ�ӡ(' || v_Fields || ') Select ' || Replace(v_Fields, '��ת��', 'Null as ��ת��') ||
                    ' From H���Ƶ��ݴ�ӡ Where (NO, ��¼����) In (Select NO, ��¼���� From H����ҽ������ Where ҽ��id = :1)';
        Execute Immediate v_Sql
          Using n_Rec_Id;
        Delete H���Ƶ��ݴ�ӡ Where (NO, ��¼����) In (Select NO, ��¼���� From H����ҽ������ Where ҽ��id = n_Rec_Id);
      
      Elsif v_Table = 'Ӱ�����¼' Then
        v_Fields := Getfields('Ӱ��������');
        v_Sql    := 'Insert Into Ӱ��������(' || v_Fields || ') Select ' || Replace(v_Fields, '��ת��', 'Null as ��ת��') ||
                    ' From HӰ�������� Where ���uid In (Select ���uid From HӰ�����¼ Where ҽ��id = :1)';
        Execute Immediate v_Sql
          Using n_Rec_Id;
      
        v_Fields := Getfields('Ӱ����ͼ��');
        v_Sql    := 'Insert Into Ӱ����ͼ��(' || v_Fields || ') Select ' || Replace(v_Fields, '��ת��', 'Null as ��ת��') ||
                    ' From HӰ����ͼ�� Where ����uid In (Select b.����uid From HӰ�����¼ A, HӰ�������� B Where a.ҽ��id = :1 And a.���uid = b.���uid)';
        Execute Immediate v_Sql
          Using n_Rec_Id;
      
        Delete HӰ����ͼ��
        Where ����uid In (Select b.����uid
                        From HӰ�����¼ A, HӰ�������� B
                        Where a.ҽ��id = n_Rec_Id And a.���uid = b.���uid);
        Delete HӰ�������� Where ���uid In (Select ���uid From HӰ�����¼ Where ҽ��id = n_Rec_Id);
      
      Elsif v_Table = '����걾��¼' Then
        For R In (Select Column_Value
                  From Table(f_Str2list('����������Ŀ,���������¼,������Ŀ�ֲ�,�����ʿؼ�¼,���������¼,����ǩ����¼,����ͼ����'))) Loop
          v_Subtable := r.Column_Value;
          If v_Subtable = '����ǩ����¼' Then
            v_Subfield := '����걾ID';
          Else
            v_Subfield := '�걾ID';
          End If;
          v_Fields := Getfields(v_Subtable);
          v_Sql    := 'Insert Into ' || v_Subtable || '(' || v_Fields || ') Select ' ||
                      Replace(v_Fields, '��ת��', 'Null as ��ת��') || ' From H' || v_Subtable || ' Where ' || v_Subfield ||
                      ' In (Select ID From H����걾��¼ Where ҽ��id = :1)';
          Execute Immediate v_Sql
            Using n_Rec_Id;
        
          v_Sql := 'Delete H' || v_Subtable || ' Where ' || v_Subfield ||
                   ' In (Select ID From H����걾��¼ Where ҽ��id = :1)';
          Execute Immediate v_Sql
            Using n_Rec_Id;
        End Loop;
      
        v_Fields := Getfields('������ͨ���');
        v_Sql    := 'Insert Into ������ͨ���(' || v_Fields || ') Select ' || Replace(v_Fields, '��ת��', 'Null as ��ת��') ||
                    ' From H������ͨ��� Where ����걾id In (Select ID From H����걾��¼ Where ҽ��id = :1)';
        Execute Immediate v_Sql
          Using n_Rec_Id;
      
        v_Fields := Getfields('����ҩ�����');
        v_Sql    := 'Insert Into ����ҩ�����(' || v_Fields || ') Select ' || Replace(v_Fields, '��ת��', 'Null as ��ת��') ||
                    ' From H����ҩ����� Where ϸ�����id In (Select ID From H������ͨ��� Where ����걾id In (Select ID From H����걾��¼ Where ҽ��id = :1))';
        Execute Immediate v_Sql
          Using n_Rec_Id;
      
        v_Fields := Getfields('�����ʿر���');
        v_Sql    := 'Insert Into �����ʿر���(' || v_Fields || ') Select ' || Replace(v_Fields, '��ת��', 'Null as ��ת��') ||
                    ' From H�����ʿر��� Where ���ID In (Select ID From H������ͨ��� Where ����걾id In (Select ID From H����걾��¼ Where ҽ��id = :1))';
        Execute Immediate v_Sql
          Using n_Rec_Id;
      
        Delete H����ҩ�����
        Where ϸ�����id In
              (Select ID From H������ͨ��� Where ����걾id In (Select ID From H����걾��¼ Where ҽ��id = n_Rec_Id));
        Delete H�����ʿر���
        Where ���id In
              (Select ID From H������ͨ��� Where ����걾id In (Select ID From H����걾��¼ Where ҽ��id = n_Rec_Id));
      
        Delete H������ͨ��� Where ����걾id In (Select ID From H����걾��¼ Where ҽ��id = n_Rec_Id);
      End If;
    
      v_Sql := 'Delete H' || v_Table || ' Where ' || v_Field || ' = :1';
      Execute Immediate v_Sql
        Using n_Rec_Id;
    End Loop;
  
    --��������
    If n_Opersystem > 0 Then
      Execute Immediate 'Call zl24_Retu_Oper(:1)'
        Using n_Rec_Id;
    End If;
  
    Delete H����ҽ����¼ Where ID = n_Rec_Id;
  End Zl_Retu_Order;

  --------------------------------------------
  --����Ϊ��������
  --------------------------------------------
Begin
  ----------------------------------------------------------------------------------------------------------
  --�Ի�����ͼ��ת������������ֻ���ж�.
  Select ��� Into n_System From zlSystems Where Upper(������) = Zl_Owner And ��� Like '1%';
  Begin
    Select Nvl(ֻ��, 0) Into n_ֻ�� From zlBakSpaces Where ϵͳ = n_System And ��ǰ = 1;
  Exception
    When Others Then
      v_Err_Msg := '[ZLSOFT]��ǰû�п��õ���ʷ���ݿռ�,���ܼ���![ZLSOFT]';
      Raise Err_Item;
  End;
  If n_ֻ�� = 1 Then
    v_Err_Msg := '[ZLSOFT]��ʷ���ݿռ�Ŀǰ��״̬Ϊֻ��,���ܼ���![ZLSOFT]';
    Raise Err_Item;
  End If;

  --�Ի�����ͼ��ת������������ֻ���ж�.
  n_Opersystem := 0;
  Select ��� Into n_Opersystem From zlSystems Where Upper(������) = Zl_Owner And ��� Like '24%';
  If n_Opersystem > 0 Then
    Begin
      Select Nvl(ֻ��, 0) Into n_ֻ�� From zlBakSpaces Where ϵͳ = n_Opersystem And ��ǰ = 1;
    Exception
      When Others Then
        v_Err_Msg := '[ZLSOFT]��ǰû�п��õ�������ϵͳ��ʷ���ݿռ�,���ܼ���![ZLSOFT]';
        Raise Err_Item;
    End;
    If n_ֻ�� = 1 Then
      v_Err_Msg := '[ZLSOFT]������ϵͳ��ʷ���ݿռ�Ŀǰ��״̬Ϊֻ��,���ܼ���![ZLSOFT]';
      Raise Err_Item;
    End If;
  End If;

  --1.���ﲡ�ˣ����Һŵ����
  If n_Flag = 0 Then
    --���δ����ʷ���
    Zl_Retu_Exes(n_Patiid, 8);
  
    v_Table  := '���˹Һż�¼';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '��ת��', 'Null as ��ת��') ||
                ' From H' || v_Table || ' Where NO =:1 ';
    Execute Immediate v_Sql
      Using v_Times;
  
    For r_Other In (Select ID, ����id From H���˹Һż�¼ Where NO = v_Times) Loop
      Zl_Retu_Other(r_Other.����id, r_Other.Id);
    End Loop;
  
    For r_Epr In (Select /*+ Rule*/
                   b.Id
                  From H���˹Һż�¼ A, H���Ӳ�����¼ B
                  Where a.No = v_Times And a.����id = n_Patiid And b.����id = a.����id And b.��ҳid = a.Id) Loop
      Zl_Retu_Epr(r_Epr.Id);
    End Loop;
  
    For r_Order In (Select ID From H����ҽ����¼ Where ������Դ <> 4 And ����id = n_Patiid And �Һŵ� = v_Times) Loop
      Zl_Retu_Order(r_Order.Id);
    End Loop;
  
    --ת���¼
    v_Table  := '����ת���¼';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '��ת��', 'Null as ��ת��') ||
                ' From H' || v_Table || ' Where NO =:1';
    Execute Immediate v_Sql
      Using v_Times;
  
    Delete H����ת���¼ Where NO = v_Times;
    Delete H���˹Һż�¼ Where NO = v_Times;
  
    --2.סԺ���ˣ�������ID����ҳID���
  Elsif n_Flag = 1 Then
    --���δ����ʷ���
    Zl_Retu_Exes(n_Patiid || ',' || v_Times, 8);
  
    Zl_Retu_Other(n_Patiid, To_Number(v_Times));
    Zl_Retu_Path(n_Patiid, To_Number(v_Times));
  
    --��ת��������תҽ����Ӱ�񱨸沵�أ�����ҽ�������������в�������ҽ�����ӱ���ҽ��ת�غ���
    For r_Epr In (Select ID From H���Ӳ�����¼ Where ����id = n_Patiid And ��ҳid = To_Number(v_Times)) Loop
      Zl_Retu_Epr(r_Epr.Id);
    End Loop;
  
    Zl_Retu_Tend(n_Patiid, To_Number(v_Times));
  
    For r_Order In (Select ID From H����ҽ����¼ Where ����id = n_Patiid And ��ҳid = To_Number(v_Times)) Loop
      Zl_Retu_Order(r_Order.Id);
    End Loop;
    Update ������ҳ Set ����ת�� = 0 Where ����id = n_Patiid And ��ҳid = To_Number(v_Times);
  
    --3.��첡��
  Elsif n_Flag = 2 Then
    Zl_Retu_Other(n_Patiid, v_Times);
  
    For r_Cpr In (Select ID From H����ҽ����¼ Where ������Դ = 4 And �Һŵ� = v_Times) Loop
      Zl_Retu_Order(r_Cpr.Id);
    End Loop;
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM || ':' || v_Sql);
End Zl_Retu_Clinic;
/

--91646:�ŵ���,2016-02-29,������ȡ���Ϸ�
CREATE OR REPLACE Procedure Zl_��Һ��ҩ��¼_ȡ����ҩ(��ҩid_In In Varchar2 --ID��:ID1,ID2....
                                           ) Is
  v_Tansid   Varchar2(20);
  v_Tmp      Varchar2(4000);
  v_No       Varchar2(20);
  v_Usercode Varchar2(100);
  n_���     ��Һ��ҩ��¼.�Ƿ���%Type := 0;
  v_������Ա ��Һ��ҩ��¼.������Ա%Type;
  d_����ʱ�� ��Һ��ҩ��¼.����ʱ��%Type;

  v_Error    Varchar2(255);
  n_����״̬ ��Һ��ҩ��¼.����״̬%Type;
  Err_Custom Exception;
  n_row      number(10);
  n_Out      number(1);

Begin
  v_Usercode := Zl_Identity;
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ';') + 1);
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ',') + 1);
  v_Usercode := Substr(v_Usercode, 1, Instr(v_Usercode, ',') - 1);
  n_Out:=Nvl(zl_GetSysParameter('��Ժ���˲������÷�', 1345), 0);

  If ��ҩid_In Is Null Then
    v_Tmp := Null;
  Else
    v_Tmp := ��ҩid_In || ',';
  End If;

  While v_Tmp Is Not Null Loop
    --�ֽⵥ��ID��
    v_Tansid := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp    := Replace(',' || v_Tmp, ',' || v_Tansid || ',');

    --��鵱ǰ��Һ����״̬�Ƿ�Ϊ����ҩ״̬
    Begin
      Select ����״̬ Into n_����״̬ From ��Һ��ҩ��¼ Where ID = v_Tansid;

      if n_����״̬!=4 then
        v_Error := '�����ݵ�ǰ������ҩ״̬�����ܽ���ȡ����ҩ��';
        Raise Err_Custom;
      end if;
    Exception
      When Others Then
        v_Error := '�������ѱ�������';
        Raise Err_Custom;
    End;

    Select ������Ա, ����ʱ��
    Into v_������Ա, d_����ʱ��
    From (Select ������Ա, ����ʱ�� From ��Һ��ҩ״̬ Where ��ҩid = v_Tansid And �������� = 2 Order By ����ʱ�� Desc)
    Where Rownum = 1;

    Update ��Һ��ҩ��¼ Set ����״̬ = 2, ������Ա = v_������Ա, ����ʱ�� = d_����ʱ�� Where ID = v_Tansid;

    Select �Ƿ��� Into n_��� From ��Һ��ҩ��¼ Where ID = v_Tansid;
    If n_��� <> 1 Then
      for r_item in (Select A.NO,B.��� From ��Һ��ҩ���� A,סԺ���ü�¼ B Where A.����ID=B.����ID and A.NO=B.No and B.��¼״̬=1 and A.��ҩid = v_Tansid) loop
        if r_item.NO is not null then
          Zl_סԺ���ʼ�¼_Delete(r_item.NO,r_item.���, v_Usercode, Zl_Username);
        end if;
      end loop;
    Else
      Zl_��Һ��ҩ��¼_ȡ����ҩ(v_Tansid);
    End If;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Һ��ҩ��¼_ȡ����ҩ;
/

--93696:����,2016-03-03,�޿�����������������
Create Or Replace Procedure Zl_ҩƷ�շ���¼_��������(�շ�id_In In ҩƷ�շ���¼.Id%Type) Is
  v_����         ҩƷ�շ���¼.����%Type;
  v_������id   ҩƷ�շ���¼.������id%Type;
  v_�ۼ۾���     Number;
  v_����     Number;
  v_�շ�id       ҩƷ�շ���¼.Id%Type;
  v_ԭ��         ҩƷ�շ���¼.���ۼ�%Type;
  v_�ּ�         ҩƷ�շ���¼.���ۼ�%Type;
  v_�Ƿ���     �շ���ĿĿ¼.�Ƿ���%Type;
  v_�۸�id       �շѼ�Ŀ.Id%Type;
  v_�ⷿid       ҩƷ�շ���¼.�ⷿid%Type;
  v_ҩƷid       ҩƷ�շ���¼.ҩƷid%Type;
  v_����         ҩƷ�շ���¼.����%Type;
  v_ʵ������     ҩƷ�շ���¼.ʵ������%Type;
  v_�������     ҩƷ�շ���¼.���۽��%Type;
  v_���ϵ��     ҩƷ�շ���¼.���ϵ��%Type;
  v_ִ������     Number;
  v_�������     ҩƷ�շ���¼.�������%Type;
  v_���۽��     ҩƷ�շ���¼.���۽��%Type;
  v_���         ҩƷ�շ���¼.���%Type;
  v_No           ҩƷ�շ���¼.No%Type;
  n_ƽ���ɱ���   ҩƷ���.ƽ���ɱ���%Type;
  n_���õ��۾��� Number;
  d_��������     ҩƷ�շ���¼.��������%Type;
  d_����         ҩƷ�շ���¼.��������%Type;
  n_��¼״̬     ҩƷ�շ���¼.��¼״̬%Type;
  n_��ǰ����     ҩƷ���.ʵ������%Type;

  v_Billno ҩƷ�շ���¼.No%Type;
Begin
  --��������
  --�ۼۣ�
  --���ۣ����շѼ�Ŀ�ּ�������۸��жϣ�������Ȳ�������
  --ʱ�ۣ�ʱ���޿�治�����п�����жϵ�ǰ�۸�������۸��Ƿ�ͬ����ͬ������򲻴���

  --�ɱ��ۣ�
  --�е������������޵��ۼ����������۸��Ƿ�ͬ����ͬ����������������

  --1���ۼ۵�������
  --��ȡ������Ϣ��ԭ�ۣ��ּۼ�ҩ������
  Select a.����, a.��¼״̬, a.�ⷿid, a.ҩƷid, Nvl(a.����, 0) As ����, Nvl(a.ʵ������, 0) * Nvl(a.����, 1) As ʵ������, a.���ϵ��,
         Nvl(a.���ۼ�, 0) ԭ��, b.�ּ�, Nvl(c.�Ƿ���, 0) �Ƿ���, b.Id As �۸�id
  Into v_����, n_��¼״̬, v_�ⷿid, v_ҩƷid, v_����, v_ʵ������, v_���ϵ��, v_ԭ��, v_�ּ�, v_�Ƿ���, v_�۸�id
  From ҩƷ�շ���¼ A, �շѼ�Ŀ B, �շ���ĿĿ¼ C
  Where a.ҩƷid = b.�շ�ϸĿid And a.ҩƷid = c.Id And
        (Sysdate Between b.ִ������ And b.��ֹ���� Or Sysdate >= b.ִ������ And b.��ֹ���� Is Null) And a.Id = �շ�id_In;

  --��ȡ���С��λ��
  Select Zl_To_Number(Nvl(zl_GetSysParameter(157), '5')) Into n_���õ��۾��� From Dual;
  If v_���� = 8 Or v_���� = 9 Or v_���� = 10 Then
    --��ҩӦ��ȡ���þ��ȣ�����ҵ��Ӧ��ȡҩƷ���ľ����о���
    Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into v_���� From Dual;
  Else
    Select Nvl(����, 2) Into v_���� From ҩƷ���ľ��� Where ���� = 0 And ��� = 1 And ���� = 4 And ��λ = 5;
  End If;
  --����ֱ��ȡ�շѼ�Ŀ���ּ�
  v_ִ������ := 0;
  If v_�Ƿ��� = 0 Then
    v_ִ������ := 1;
  Else
    --ʱ��ҩƷ�����ۼۣ����ܴ��շѼ�Ŀ��ȡ����Ϊʱ�۵��ۿ����ǰ��ⷿ��������������
    v_�ּ� := 0;
  
    If v_���� > 0 Then
      --ʱ�۷������ȴӿ�����ȡ�����û�п����˵�������۸��뵱ǰ�۸�һ�����ô�����
      Begin
        Select Nvl(���ۼ�, 0)
        Into v_�ּ�
        From ҩƷ���
        Where ���� = 1 And �ⷿid + 0 = v_�ⷿid And ҩƷid = v_ҩƷid And Nvl(����, 0) = v_����;
      Exception
        When Others Then
          v_�ּ� := 0;
      End;
    End If;
  
    If v_�ּ� > 0 Then
      v_ִ������ := 1;
    End If;
  End If;

  If v_ִ������ = 1 Then
    --��ҩ�൥���ۼ۾���Ϊ5��������ͨ�൥��Ϊ7
    If v_���� = 8 Or v_���� = 9 Or v_���� = 10 Then
      v_�ۼ۾��� := n_���õ��۾���;
    Else
      v_�ۼ۾��� := 7;
    End If;
  
    --�Ƚ�ԭ�ۺ��ּۣ���ͬ����
    If v_ԭ�� <> Round(v_�ּ�, v_�ۼ۾���) Then
      Select ҩƷ�շ���¼_Id.Nextval Into v_�շ�id From Dual;
      Select Nextno(147) Into v_Billno From Dual;
    
      Select ���id Into v_������id From ҩƷ�������� Where ���� = 13;
    
      v_������� := Round(v_���ϵ�� * (Round(v_�ּ�, v_�ۼ۾���) - v_ԭ��) * v_ʵ������, v_����);
    
      --��������������¼
      Insert Into ҩƷ�շ���¼
        (ID, ��¼״̬, ����, NO, ���, ������id, ҩƷid, ����, ����, Ч��, ����, ����, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ����, ���۽��, ���, ժҪ, ������,
         ��������, �ⷿid, ���ϵ��, �۸�id, �����, �������, ����id)
        Select v_�շ�id, 1, 13, v_Billno, ���, v_������id, ҩƷid, ����, ����, Ч��, ����, ����, Abs(v_ʵ������), 0, v_ԭ��, 0,
               Round(v_�ּ�, v_�ۼ۾���), ����, v_�������, v_�������, '�Զ��������۱䶯', �����, �������, �ⷿid, 1, v_�۸�id, �����, �������, �շ�id_In
        From ҩƷ�շ���¼
        Where ID = �շ�id_In;
    
      --����ҩƷ���
      Zl_ҩƷ���_Update(v_�շ�id);
    
    End If;
  End If;

  --2���ɱ��۵��ۣ�ֻ�г���ҵ�������
  If Mod(n_��¼״̬, 3) = 2 Then
    Select �ⷿid, ҩƷid, Nvl(����, 0) ����, ���ϵ�� * Nvl(ʵ������, 0) * Nvl(����, 1) As ʵ������, ���ϵ�� * ���۽��, ���ϵ�� * ���, �ɱ���
    Into v_�ⷿid, v_ҩƷid, v_����, v_ʵ������, v_���۽��, v_���, v_ԭ��
    From ҩƷ�շ���¼
    Where ID = �շ�id_In;
  
    --ȡԭʼ���ݵ����ʱ��
    Select a.�������
    Into v_�������
    From ҩƷ�շ���¼ A, ҩƷ�շ���¼ B
    Where b.Id = �շ�id_In And a.���� = b.���� And a.No = b.No And a.�ⷿid = b.�ⷿid And a.ҩƷid + 0 = b.ҩƷid And a.��� = b.��� And
          (a.��¼״̬ = 1 Or Mod(a.��¼״̬, 3) = 0);
  
    v_ִ������ := 0;
    Begin
      Select 1, �³ɱ���
      Into v_ִ������, v_�ּ�
      From �ɱ��۵�����Ϣ
      Where �ⷿid + 0 = v_�ⷿid And ҩƷid + 0 = v_ҩƷid And Nvl(����, 0) = v_���� And ִ������ > v_������� And Rownum = 1
      Order By ִ������ Desc;
    
    Exception
      When Others Then
        v_ִ������ := 0;
        v_�ּ�     := 0;
      
        --�������޿����ۣ���ô����Ҫ�ⷿid�������жϣ�ֻ��Ҫ��ҩƷid������
        Begin
          Select 1, �³ɱ���
          Into v_ִ������, v_�ּ�
          From �ɱ��۵�����Ϣ
          Where ҩƷid + 0 = v_ҩƷid And ִ������ > v_������� And Rownum = 1
          Order By ִ������ Desc;
        
        Exception
          When Others Then
            --���ܳ��ֳ��������˿�ܾ���ǰ������,���ʱ��۸�����Ѿ������仯,������Ҫ����
            Begin
              Select ƽ���ɱ���, Nvl(ʵ������, 0) As ʵ������
              Into v_�ּ�, n_��ǰ����
              From ҩƷ���
              Where �ⷿid = v_�ⷿid And ҩƷid = v_ҩƷid And Nvl(����, 0) = v_���� And ���� = 1;
            
              --�޿�治�����п����������������1.ԭʼ�п�棬������������������2��ԭʼ���������¿��������Դ�ڳ�������
              --�����1�����ô��ͨ������ԭ�۸�������۸�Ƚϼ���
              --�����2�����ô����Ҫ�жϵ�ǰ����������������*���ϵ���Ƿ���ȣ���������˵���϶���ԭʼ�޿��Ȼ��������ݲ�����
              --���ʱ����Ҫ�����жϣ�ֻ��ͨ������������һ�¼۸��жϣ���۸�Ϊ����
              If n_��ǰ���� = v_ʵ������ Then
                --˵���ǵڶ�����������Ŀ�����ݣ����ʱ����ܲ��������ݺ����ף�����Ҫ���ж�
                --93696������Ҫ����������ε������޿���˿�ʱ���ٴ�����������
                /*Select �ɱ��� Into v_�ּ� From ҩƷ��� Where ҩƷid = v_ҩƷid;
                If Abs(Abs(v_�ּ�) - Abs(v_ԭ��)) > 1 Then
                  v_ִ������ := 1;
                End If;*/
                Null;
              Else
                --˵���ǵ�һ��������������ݣ�ֱ���õ�ǰ�����������ݼ۸�Ƚϼ���
                If Round(v_�ּ�, 2) <> Round(v_ԭ��, 2) Then
                  v_ִ������ := 1;
                End If;
              End If;
            Exception
              When Others Then
                v_ִ������ := 0;
            End;
        End;
    End;
  
    If v_ִ������ = 1 Then
      v_������� := (v_���۽�� - v_���) - Round(Round(v_�ּ�, 3) * v_ʵ������, v_����);
    
      If v_������� <> 0 Then
        Select ҩƷ�շ���¼_Id.Nextval Into v_�շ�id From Dual;
        Select b.Id, b.ϵ��
        Into v_������id, v_���ϵ��
        From ҩƷ�������� A, ҩƷ������ B
        Where a.���id = b.Id And a.���� = 5 And Rownum < 2;
      
        v_No := Nextno(25, v_�ⷿid);
      
        --��������۵�����
        Insert Into ҩƷ�շ���¼
          (ID, ��¼״̬, ����, NO, ���, �ⷿid, ������id, ��ҩ��λid, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ��д����, ���ۼ�, �ɱ���, ���, ժҪ, ������, ��������,
           �����, �������, ��������, ��׼�ĺ�, ����, ��ҩ��ʽ, ����, ���Ч��, ����id)
          Select v_�շ�id, 1, 5, v_No, 1, �ⷿid, v_������id, ��ҩ��λid, v_���ϵ��, ҩƷid, ����, ����, ����, Ч��, v_ʵ������, v_���۽��, v_���,
                 v_�������, '�Զ��������۱䶯', �����, �������, �����, �������, ��������, ��׼�ĺ�, v_�ּ�, 1, �ɱ���, ���Ч��, �շ�id_In
          From ҩƷ�շ���¼
          Where ID = �շ�id_In;
      
        --���¿��
        Zl_ҩƷ���_Update(v_�շ�id);
      
      End If;
    End If;
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�շ���¼_��������;
/

--93696:����,2016-03-03,�޿�����������������
Create Or Replace Procedure Zl_ҩƷ���_Update
(
  Id_In       In ҩƷ�շ���¼.Id%Type,
  Delete_In   In Number := 0,
  ������ʽ_In In Number := 0,
  ��ҩ��־_In In Number := 0,
  �������_In In Number := 0
) Is
  ----------------------------------------------------------------------------------------
  --����:������ϸ���ݸ��¿��
  --�ؼ��������¿��ÿ����������Ƿ����������
  --ҵ����򣺰���ģ��ֿ��������ݣ����ں���ά��
  --�������÷�Χ��ҩƷ��ͨҵ���漰������ҩƷ�շ���¼��ϸ���ٸ��¿����������ƽ���ɱ��۵�ҵ�񣬸ù���
  --ֻ�������������ڲ����ã�������Ϊ��������ֱ��ִ��
  --����:
  --     Id_In:ҩƷҵ��������ɾ������ˡ�����ʱ�����շ���¼��ϸ��id
  --     Delete_in: 0--��ɾ������ҵ����������ˡ������� 1--ɾ������ҵ��
  --     ������ʽ_In: 0--����������ʽ 1-�����������뵥�� 2-���� 3-���� Ŀǰֻ���ƿ�ģ����Ч
  --     ��ҩ��־_in: 0--�����  1--���  �˲���ֻ��ҩƷ���������ŷ�ҩģ����Ч
  --     �������_in:0,������˵���,1-����ҵ��
  ----------------------------------------------------------------------------------------
  v_�¿������� Zlparameters.����ֵ%Type;
  n_��������   ҩƷ���.ʵ������%Type;
  n_ʵ������   ҩƷ���.ʵ������%Type;
  n_���۽��   ҩƷ���.ʵ�ʽ��%Type;
  n_���       ҩƷ���.ʵ�ʲ��%Type;
  n_ʱ�۷���   Number(1);
  n_�ɱ���     ҩƷ�շ���¼.�ɱ���%Type;
  n_���ۼ�     ҩƷ���.���ۼ�%Type;

  n_�������   ҩƷ���.ʵ������%Type;
  n_���ƽ���� ҩƷ���.ƽ���ɱ���%Type;
  n_������     ҩƷ�շ���¼.ʵ������%Type;
  n_�ܳɱ���   ҩƷ�շ���¼.�ɱ���%Type;

  --ҵ����ϸ���ݣ��ѿ�����ݸ�����Ҫ�����ݶ��г���
  Cursor c_Detail Is
    Select a.Id, a.��¼״̬, a.����, a.No, a.���, a.�ⷿid, a.��ҩ��λid, a.������id, a.�Է�����id, a.���ϵ��, Nvl(a.��ҩ��ʽ, 0) As ��ҩ��ʽ, a.ҩƷid,
           Nvl(a.����, 0) ����, a.����, a.����, a.��������, a.Ч��, a.����, Nvl(a.��д����, 0) As ��д����, a.ʵ������, a.�ɱ���, a.�ɱ����, a.����, a.���ۼ�,
           Nvl(a.���۽��, 0) As ���۽��, Nvl(a.���, 0) As ���, a.��ҩ��, a.��ҩ����, a.�����, a.�������, a.�������, a.���Ч��, a.��׼�ĺ�, a.��Ʒ����,
           a.�ڲ�����, b.�Ƿ���, a.����, a.Ƶ��, a.ժҪ
    From ҩƷ�շ���¼ A, �շ���ĿĿ¼ B
    Where a.ҩƷid = b.Id And a.Id = Id_In;

  v_Detail c_Detail%RowType;
Begin
  --ȡ�¿��ÿ�����
  Select zl_GetSysParameter(96) Into v_�¿������� From Dual;

  For v_Detail In c_Detail Loop
    n_ʵ������ := v_Detail.���ϵ�� * v_Detail.ʵ������ * Nvl(v_Detail.����, 1);
    If n_ʵ������ Is Null Then
      n_ʵ������ := 0;
    End If;
    n_���۽�� := v_Detail.���ϵ�� * v_Detail.���۽��;
    n_���     := v_Detail.���ϵ�� * v_Detail.���;
  
    --��ȡ���͵��ݵ������ͳɱ���
    Begin
      Select Nvl(ʵ������, 0), Nvl(ƽ���ɱ���, 0)
      Into n_�������, n_���ƽ����
      From ҩƷ���
      Where ���� = 1 And �ⷿid = v_Detail.�ⷿid And ҩƷid = v_Detail.ҩƷid And Nvl(����, 0) = v_Detail.����;
    Exception
      When Others Then
        n_�������   := 0;
        n_���ƽ���� := 0;
    End;
  
    --�⹺��⣺����ҵ������⣬���ʱ��������������������ʱ��������ģ��˿�ģʽʱ���ʱҪ���ݲ������������������������෴�����������
    --ɾ������ʱҪ���ʱԤ���ļӻ�ȥ
    --����ʱֱ�Ӱ������Ӽ�����������
    --�������ж�����⻹���˿�
    If v_Detail.���� = 1 Then
      If v_Detail.������� Is Null Then
        --δ��˵��ݣ����ɾ��
        If Delete_In = 0 Then
          If n_ʵ������ < 0 Then
            If v_�¿������� = '1' Then
              n_�������� := n_ʵ������;
            Else
              n_�������� := 0;
            End If;
          Else
            n_�������� := 0;
          End If;
        Else
          If n_ʵ������ < 0 Then
            If v_�¿������� = '1' Then
              n_�������� := -1 * n_ʵ������;
            Else
              n_�������� := 0;
            End If;
          Else
            n_�������� := 0;
          End If;
        End If;
      Else
        --����˻��ѳ���
        If v_Detail.��¼״̬ = 1 Then
          --���
          If n_ʵ������ < 0 Then
            --�˿�Ҫ�����ʱ�Ѿ������˿�������
            If v_�¿������� = '1' Then
              n_�������� := 0;
            Else
              n_�������� := n_ʵ������;
            End If;
          Else
            --��ͨ���
            n_�������� := n_ʵ������;
          End If;
        Else
          --����
          n_�������� := n_ʵ������;
        End If;
      End If;
    End If;
  
    --������⣺��������ҩƷ��˵����⣬���ʱ��������������������ʱ��������ģ�����ԭ��ҩ��˵�ǳ��⣬���ʱ���ݲ������������������������෴�����������
    --ɾ������ʱҪ��ԭ��ҩԤ���������ӻ�ȥ
    --�����ϵ���ж�����⻹���˿�
    If v_Detail.���� = 2 Then
      If v_Detail.������� Is Null Then
        --���ɾ��
        If Delete_In = 0 Then
          If v_Detail.���ϵ�� < 0 Then
            If v_�¿������� = '1' Then
              n_�������� := n_ʵ������;
            Else
              n_�������� := 0;
            End If;
          Else
            n_�������� := 0;
          End If;
        Else
          If v_Detail.���ϵ�� < 0 Then
            If v_�¿������� = '1' Then
              n_�������� := -1 * n_ʵ������;
            Else
              n_�������� := 0;
            End If;
          Else
            n_�������� := 0;
          End If;
        End If;
      Else
        --��˺ͳ���
        If v_Detail.���ϵ�� < 0 Then
          If v_Detail.��¼״̬ = 1 Then
            If v_�¿������� = '1' Then
              n_�������� := 0;
            Else
              n_�������� := n_ʵ������;
            End If;
          Else
            n_�������� := n_ʵ������;
          End If;
        Else
          n_�������� := n_ʵ������;
        End If;
      End If;
    End If;
  
    --Э����⣺����Э��ҩƷ��˵����⣬���ʱ��������������������ʱ��������ģ��������ҩ��˵�ǳ��⣬���ʱ���ݲ������������������������෴�����������
    --ɾ������ʱҪ��ԭ��ҩԤ���������ӻ�ȥ
    --�����ϵ���ж�����⻹���˿�
    If v_Detail.���� = 3 Then
      If v_Detail.������� Is Null Then
        --���ɾ��
        If Delete_In = 0 Then
          If v_Detail.���ϵ�� < 0 Then
            If v_�¿������� = '1' Then
              n_�������� := n_ʵ������;
            Else
              n_�������� := 0;
            End If;
          Else
            n_�������� := 0;
          End If;
        Else
          If v_Detail.���ϵ�� < 0 Then
            If v_�¿������� = '1' Then
              n_�������� := -1 * n_ʵ������;
            Else
              n_�������� := 0;
            End If;
          Else
            n_�������� := 0;
          End If;
        End If;
      Else
        --��˺ͳ���
        If v_Detail.���ϵ�� < 0 Then
          If v_Detail.��¼״̬ = 1 Then
            If v_�¿������� = '1' Then
              n_�������� := 0;
            Else
              n_�������� := n_ʵ������;
            End If;
          Else
            n_�������� := n_ʵ������;
          End If;
        Else
          n_�������� := n_ʵ������;
        End If;
      End If;
    End If;
  
    --������⣺����ҵ������⣬���ʱ��������������������ʱ��������ģ��������ģʽʱҪ���ݲ���������������������ʱ���෴�����������
    --ɾ������ʱҪ���ʱԤ���ļӻ�ȥ
    --�������ж�����⻹���˿�
    If v_Detail.���� = 4 Then
      If v_Detail.������� Is Null Then
        --���ɾ��
        If Delete_In = 0 Then
          If n_ʵ������ < 0 Then
            If v_�¿������� = '1' Then
              n_�������� := n_ʵ������;
            Else
              n_�������� := 0;
            End If;
          Else
            n_�������� := 0;
          End If;
        Else
          If n_ʵ������ < 0 Then
            If v_�¿������� = '1' Then
              n_�������� := -1 * n_ʵ������;
            Else
              n_�������� := 0;
            End If;
          Else
            n_�������� := 0;
          End If;
        End If;
      Else
        If v_Detail.��¼״̬ = 1 Then
          --���
          If n_ʵ������ < 0 Then
            If v_�¿������� = '1' Then
              n_�������� := 0;
            Else
              n_�������� := n_ʵ������;
            End If;
          Else
            --��ͨ���
            n_�������� := n_ʵ������;
          End If;
        Else
          --����
          n_�������� := n_ʵ������;
        End If;
      End If;
    End If;
  
    --��۵������ɱ��۵��ۣ����漰��������仯�����ʱ�����������ʱֻ�������۵�����
    If v_Detail.���� = 5 Then
      n_�������� := 0;
    End If;
  
    --�ƿ⣺�ƿ����������ݣ�һ�����ⵥ�ݣ�һ����ⵥ�ݣ����ⵥ����Ҫ�����¿��ÿ����������Ƿ��¿��ÿ�棬����ǳ�������ÿ�����෴����
    --���������������ʱ���ڷ���ʱԤ���������������ʱ�������������
    --�������ģʽʱҲҪ���ݲ����������������
    --�ʱ����ҵ����ݲ��������Ƿ��¿�棬���ҵ���¿�棻ɾ��ʱ����ҵ����Ӳ���Ҫ�ѿ�滹��ȥ�����ҵ�񲻻����
    If v_Detail.���� = 6 Then
      If v_Detail.������� Is Null Then
        If Delete_In = 0 Then
          --�������޸ġ����͡����ˡ���������
          If v_Detail.��¼״̬ = 1 Then
            If ������ʽ_In = 2 Then
              --����
              If v_�¿������� = '0' And v_Detail.���ϵ�� = -1 Then
                n_�������� := n_ʵ������;
              Else
                n_�������� := 0;
              End If;
            Elsif ������ʽ_In = 3 Then
              --����
              If v_�¿������� = '0' And v_Detail.���ϵ�� = -1 Then
                n_�������� := -1 * n_ʵ������;
              Else
                n_�������� := 0;
              End If;
            Else
              --����
              If v_�¿������� = '1' And v_Detail.���ϵ�� = -1 Then
                n_�������� := n_ʵ������;
              Else
                n_�������� := 0;
              End If;
            End If;
          Else
            --�������
            If v_�¿������� = '1' And v_Detail.���ϵ�� = 1 Then
              n_�������� := n_ʵ������;
            Else
              n_�������� := 0;
            End If;
          End If;
        Else
          --ɾ��
          If v_Detail.��¼״̬ = 1 Then
            If v_�¿������� = '1' And v_Detail.���ϵ�� = -1 Then
              n_�������� := -1 * n_ʵ������;
            Elsif v_Detail.��ҩ���� Is Not Null And v_Detail.���ϵ�� = -1 Then
              n_�������� := -1 * n_ʵ������;
            Else
              n_�������� := 0;
            End If;
          Else
            If v_�¿������� = '1' And v_Detail.���ϵ�� = 1 Then
              n_�������� := -1 * n_ʵ������;
            Else
              n_�������� := 0;
            End If;
          End If;
        End If;
      Else
        If v_Detail.��¼״̬ = 1 Then
          --���
          If v_Detail.���ϵ�� = -1 Then
            --�����Ǳ�
            n_�������� := 0;
          Else
            --����Ǳ�
            n_�������� := n_ʵ������;
          End If;
        Else
          If ������ʽ_In = 0 Then
            --�����������
            n_�������� := n_ʵ������;
          Else
            --����������
            If v_�¿������� = '1' And v_Detail.���ϵ�� = 1 Then
              n_�������� := 0;
            Else
              n_�������� := n_ʵ������;
            End If;
          End If;
        End If;
      End If;
    End If;
  
    --���ã�����ҵ���ǳ��⣬���ʱ���ݲ���������������������ʱ�෴����
    --ɾ������ʱҪ���ʱԤ���ļӻ�ȥ
    If v_Detail.���� = 7 Then
      If v_Detail.������� Is Null Then
        --���ɾ��
        If Delete_In = 0 Then
          If v_�¿������� = '1' Then
            n_�������� := n_ʵ������;
          Else
            n_�������� := 0;
          End If;
        Else
          If v_�¿������� = '1' Then
            n_�������� := -1 * n_ʵ������;
          Else
            n_�������� := 0;
          End If;
        End If;
      Else
        --��˺ͳ���
        If v_Detail.��¼״̬ = 1 Then
          --���
          If v_�¿������� = '1' Then
            n_�������� := 0;
          Else
            n_�������� := n_ʵ������;
          End If;
        Else
          --����
          n_�������� := n_ʵ������;
        End If;
      End If;
    End If;
  
    --��ҩҵ�����ʱ�̶�������������������ʱ������
    --ɾ������ʱҪ���ʱԤ���ļӻ�ȥ
    --���ٷ�ҩ��ǵĿ������������ͬ��ɾ���������
    If v_Detail.���� = 8 Or v_Detail.���� = 9 Or v_Detail.���� = 10 Then
      If v_Detail.������� Is Null Then
        If Delete_In = 0 Then
          If ��ҩ��־_In = 0 Then
            n_�������� := n_ʵ������;
          Else
            n_�������� := -1 * n_ʵ������;
          End If;
        Else
          n_�������� := -1 * n_ʵ������;
        End If;
      Else
        n_�������� := 0;
      End If;
    End If;
  
    --�������⣺����ҵ���ǳ��⣬���ʱ���ݲ���������������������ʱ�෴����
    --ɾ������ʱҪ���ʱԤ���ļӻ�ȥ
    If v_Detail.���� = 11 Then
      If v_Detail.������� Is Null Then
        --���ɾ��
        If Delete_In = 0 Then
          If v_�¿������� = '1' Then
            n_�������� := n_ʵ������;
          Else
            n_�������� := 0;
          End If;
        Else
          If v_�¿������� = '1' Then
            n_�������� := -1 * n_ʵ������;
          Else
            n_�������� := 0;
          End If;
        End If;
      Else
        --���������
        If v_Detail.��¼״̬ = 1 Then
          --���
          If v_�¿������� = '1' Then
            n_�������� := 0;
          Else
            n_�������� := n_ʵ������;
          End If;
        Else
          --����
          n_�������� := n_ʵ������;
        End If;
      End If;
    End If;
  
    --�̵㣺�ʱ��ӯҵ�񲻴�������������̿�ҵ��̶�������������������ʱ�෴����
    --ɾ������ʱҪ���ʱԤ���ļӻ�ȥ
    --�����ϵ��������ӯ�̿�ҵ��
    If v_Detail.���� = 12 Then
      If v_Detail.������� Is Null Then
        --���ɾ��
        If Delete_In = 0 Then
          If v_Detail.���ϵ�� = -1 Then
            n_�������� := n_ʵ������;
          Else
            n_�������� := 0;
          End If;
        Else
          If v_Detail.���ϵ�� = -1 Then
            n_�������� := -1 * n_ʵ������;
          Else
            n_�������� := 0;
          End If;
        End If;
      Else
        --��˺ͳ���
        If v_Detail.��¼״̬ = 1 Then
          --���
          If v_Detail.���ϵ�� = '1' Then
            n_�������� := n_ʵ������;
          Else
            n_�������� := 0;
          End If;
        Else
          --����
          n_�������� := n_ʵ������;
        End If;
      End If;
    End If;
  
    --�ۼ۵��ۣ����漰��������仯�����ʱ�����������ʱֻ�������۵�����
    If v_Detail.���� = 13 Then
      n_�������� := 0;
    End If;
  
    --ҩƷ���棺������ҩ����ʱ���Ѿ����˿�棬���ŷ�ҩʱ����Ҫ�����ӻ�ȥ
    If v_Detail.���� = 27 Then
      n_�������� := n_ʵ������;
    End If;
  
    If v_Detail.���� > 0 And v_Detail.�Ƿ��� = 1 Then
      n_ʱ�۷��� := 1;
    Else
      n_ʱ�۷��� := 0;
    End If;
  
    n_���ۼ� := v_Detail.���ۼ�;
    --���ⵥ����Ҫ����ɱ��� ���ⵥ���е���=5 ����=12
    If v_Detail.���� = 5 Or v_Detail.���� = 12 Then
      If v_Detail.���� = 5 Then
        If v_Detail.��д���� <> 0 Then
          n_���ۼ� := Nvl(v_Detail.���ۼ�, 0) / v_Detail.��д����;
        Else
          n_���ۼ� := 0;
        End If;
        --���
        If v_Detail.��¼״̬ = 1 Then
          --��۵�����ҩ��ʽ=0���������ۡ��˻�����ҩ�����ĵ���������ҩ��ʽ=1
          n_�ɱ��� := v_Detail.����;
        Else
          --���� ��ԭԭʼ�ɱ���
          Begin
            --�ɱ���=(���-���)/����
            n_�ɱ��� := (Nvl(v_Detail.���ۼ�, 0) - Nvl(v_Detail.�ɱ���, 0)) / v_Detail.��д����;
          Exception
            When Others Then
              Select �ɱ��� Into n_�ɱ��� From ҩƷ��� Where ҩƷid = v_Detail.ҩƷid;
          End;
        End If;
      Else
        n_�ɱ��� := v_Detail.����;
      End If;
    Else
      If v_Detail.���� = 13 Then
        n_�ɱ��� := Nvl(v_Detail.����, 0) - Nvl(v_Detail.Ƶ��, 0);
      Else
        n_�ɱ��� := v_Detail.�ɱ���;
      End If;
    End If;
  
    --����ҵ�����ݸ��¿���¼
    If v_Detail.������� Is Null Then
      If n_�������� <> 0 Then
        --���ɾ��ʱֻ���¿�������
        Update ҩƷ���
        Set �������� = �������� + n_��������
        Where ���� = 1 And �ⷿid = v_Detail.�ⷿid And ҩƷid = v_Detail.ҩƷid And Nvl(����, 0) = v_Detail.����;
      
        If Sql%RowCount = 0 Then
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, Ч��, ����, ��������, ʵ������, ʵ�ʽ��, ʵ�ʲ��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴ���������, �ϴβ���, ���Ч��, ��׼�ĺ�, ���ۼ�, �ϴο���,
             ��Ʒ����, �ڲ�����, ƽ���ɱ���)
          Values
            (v_Detail.�ⷿid, v_Detail.ҩƷid, v_Detail.����, v_Detail.Ч��, 1, n_��������, 0, 0, 0, v_Detail.��ҩ��λid, n_�ɱ���,
             v_Detail.����, v_Detail.��������, v_Detail.����, v_Detail.���Ч��, v_Detail.��׼�ĺ�, Decode(n_ʱ�۷���, 1, n_���ۼ�, Null),
             v_Detail.����, v_Detail.��Ʒ����, v_Detail.�ڲ�����, n_�ɱ���);
        End If;
      End If;
    Else
      --���ʱ���¿�����������ʵ����������������۵�����
      If v_Detail.���� = 5 Then
        --����=5 �ĳɱ���������¼ ƽ���ɱ��۲���Ҫ���㣬��Ϊ���������¼۸��
        If v_Detail.ժҪ = '�⹺�˿�������Զ�����' Or v_Detail.ժҪ = '������˼۸�䶯����' Then
          --��һ���϶����⹺�˿⣬�⹺�˿�ֻ���³ɱ���,�ҿ϶��п��
          Update ҩƷ���
          Set ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + n_���
          Where ���� = 1 And �ⷿid = v_Detail.�ⷿid And ҩƷid = v_Detail.ҩƷid And Nvl(����, 0) = v_Detail.����;
        Else
          Update ҩƷ���
          Set ƽ���ɱ��� = n_�ɱ���, �ϴβɹ��� = n_�ɱ���, ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) + n_���۽��, ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + n_���
          Where ���� = 1 And �ⷿid = v_Detail.�ⷿid And ҩƷid = v_Detail.ҩƷid And Nvl(����, 0) = v_Detail.����;
          If Sql%NotFound Then
            Insert Into ҩƷ���
              (�ⷿid, ҩƷid, ����, ����, ʵ�ʲ��, �ϴ�����, Ч��, �ϴβ���, �ϴι�Ӧ��id, �ϴ���������, ��׼�ĺ�, ʵ�ʽ��, �ϴβɹ���, ƽ���ɱ���)
            Values
              (v_Detail.�ⷿid, v_Detail.ҩƷid, v_Detail.����, 1, n_���, v_Detail.����, v_Detail.Ч��, v_Detail.����,
               v_Detail.��ҩ��λid, v_Detail.��������, v_Detail.��׼�ĺ�, n_���۽��, n_�ɱ���, n_�ɱ���);
          End If;
        End If;
      Elsif v_Detail.���� = 13 Then
        --����=13 ���ۼ�������¼ ͬ�����µĽ��Ͳ�ۣ����Բ���Ҫ����ƽ���ɱ���
        Update ҩƷ���
        Set ���ۼ� = Decode(n_ʱ�۷���, 1, n_���ۼ�, Null), ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) + n_���۽��, ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + n_���
        Where ���� = 1 And �ⷿid = v_Detail.�ⷿid And ҩƷid = v_Detail.ҩƷid And Nvl(����, 0) = v_Detail.����;
        If Sql%RowCount = 0 Then
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, ��������, ʵ������, ʵ�ʽ��, ʵ�ʲ��, ���ۼ�)
          Values
            (v_Detail.�ⷿid, v_Detail.ҩƷid, v_Detail.����, 1, 0, 0, n_���۽��, n_���۽��, Decode(n_ʱ�۷���, 1, n_���ۼ�, Null));
        Else
          If n_ʱ�۷��� = 1 Then
            Update ҩƷ���
            Set ���ۼ� = ʵ�ʽ�� / ʵ������
            Where ���� = 1 And �ⷿid = v_Detail.�ⷿid And ҩƷid = v_Detail.ҩƷid And Nvl(����, 0) = v_Detail.���� And
                  Nvl(ʵ������, 0) <> 0;
          End If;
        End If;
      Else
        --�����ͳ��� ״̬�ֽ�
        --���ҵ��,������������������ּ۸���������Ҫ���¿���������Ϣ
        If (v_Detail.���ϵ�� = 1 And v_Detail.��¼״̬ = 1) Or (v_Detail.���ϵ�� = -1 And Mod(v_Detail.��¼״̬, 3) = 2) Or
           (v_Detail.���ϵ�� = 1 And Mod(v_Detail.��¼״̬, 3) = 2) Then
          Update ҩƷ���
          Set �������� = Nvl(��������, 0) + n_��������, ʵ������ = Nvl(ʵ������, 0) + n_ʵ������, ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) + n_���۽��,
              ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + n_���, �ϴι�Ӧ��id = v_Detail.��ҩ��λid,
              �ϴβɹ��� = Decode(v_Detail.����, 1, Decode(v_Detail.��ҩ��ʽ, 1, �ϴβɹ���, n_�ɱ���), n_�ɱ���),
              �ϴ����� = Nvl(v_Detail.����, �ϴ�����), �ϴ��������� = Nvl(v_Detail.��������, �ϴ���������), �ϴβ��� = Nvl(v_Detail.����, �ϴβ���),
              ���Ч�� = Nvl(v_Detail.���Ч��, ���Ч��), Ч�� = Nvl(v_Detail.Ч��, Ч��), ��׼�ĺ� = Nvl(v_Detail.��׼�ĺ�, ��׼�ĺ�),
              �ϴο��� = Decode(v_Detail.����, 1, v_Detail.����, �ϴο���), ��Ʒ���� = Nvl(v_Detail.��Ʒ����, ��Ʒ����),
              �ڲ����� = Nvl(v_Detail.�ڲ�����, �ڲ�����)
          Where �ⷿid = v_Detail.�ⷿid And ҩƷid = v_Detail.ҩƷid And Nvl(����, 0) = v_Detail.���� And ���� = 1;
          --�⹺��������������ʱ
          If (v_Detail.���� = 1 And v_Detail.��¼״̬ = 1 And �������_In = 0) Or (v_Detail.���� = 4 And v_Detail.��¼״̬ = 1) Then
            Update ҩƷ���
            Set ���ۼ� = Decode(n_ʱ�۷���, 1, n_���ۼ�, Null)
            Where �ⷿid = v_Detail.�ⷿid And ҩƷid = v_Detail.ҩƷid And Nvl(����, 0) = v_Detail.���� And ���� = 1;
          End If;
          --�����������Ҫ����ɱ���
          --�⹺�˻���������˺����г���ҵ�񲻸���ƽ���ɱ��ۣ����ֵ�ǰ�۸�
          If (v_Detail.���� = 1 And v_Detail.��ҩ��ʽ = 1) Or Mod(v_Detail.��¼״̬, 3) = 2 Or (v_Detail.���� = 1 And �������_In = 1) Then
            Null;
          Else
            --���ܽ��/��������ʽ����ƽ���ɱ��۶����ã����-��ۣ�/������Ϊ�����ݵ�׼ȷ��
            n_������ := (n_������� + n_ʵ������);
            If n_������ <> 0 Then
              n_�ܳɱ��� := (n_������� * n_���ƽ���� + n_ʵ������ * n_�ɱ���) / n_������;
              Update ҩƷ���
              Set ƽ���ɱ��� = n_�ܳɱ���
              Where ���� = 1 And �ⷿid = v_Detail.�ⷿid And ҩƷid = v_Detail.ҩƷid And Nvl(����, 0) = v_Detail.����;
            End If;
          End If;
        Else
          --����ҵ��ֻ��Ҫ���������������
          Update ҩƷ���
          Set �������� = Nvl(��������, 0) + n_��������, ʵ������ = Nvl(ʵ������, 0) + n_ʵ������, ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) + n_���۽��,
              ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + n_���
          Where �ⷿid = v_Detail.�ⷿid And ҩƷid = v_Detail.ҩƷid And Nvl(����, 0) = v_Detail.���� And ���� = 1;
        End If;
        --����δ�ҵ���������Ҫ��������������Ϣ
        If Sql%RowCount = 0 Then
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, Ч��, ����, ��������, ʵ������, ʵ�ʽ��, ʵ�ʲ��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴ���������, �ϴβ���, ���Ч��, ��׼�ĺ�, ���ۼ�, �ϴο���,
             ��Ʒ����, �ڲ�����, ƽ���ɱ���)
          Values
            (v_Detail.�ⷿid, v_Detail.ҩƷid, v_Detail.����, v_Detail.Ч��, 1, n_��������, n_ʵ������, n_���۽��, n_���, v_Detail.��ҩ��λid,
             n_�ɱ���, v_Detail.����, v_Detail.��������, v_Detail.����, v_Detail.���Ч��, v_Detail.��׼�ĺ�,
             Decode(n_ʱ�۷���, 1, n_���ۼ�, Null), v_Detail.����, v_Detail.��Ʒ����, v_Detail.�ڲ�����, n_�ɱ���);
        End If;       
      End If;
    End If;
  
    --ɾ������Ŀ������
    If �������_In = 0 Then
      Delete From ҩƷ���
      Where ���� = 1 And �ⷿid = v_Detail.�ⷿid And ҩƷid = v_Detail.ҩƷid And Nvl(����, 0) = v_Detail.���� And Nvl(��������, 0) = 0 And
            Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0 And Nvl(ʵ�ʲ��, 0) = 0;
    End If;
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ���_Update;
/


--93623:���˺�,2016-02-26,����ҽ�����͵ļƷ�״̬
Create Or Replace Procedure Zl_�����շѼ�¼_����
(
  No_In         ������ü�¼.No%Type,
  ����Ա���_In ������ü�¼.����Ա���%Type,
  ����Ա����_In ������ü�¼.����Ա����%Type,
  ���_In       Varchar2 := Null,
  �˷�ʱ��_In   ������ü�¼.�Ǽ�ʱ��%Type := Null,
  �˷�ժҪ_In   ������ü�¼.ժҪ%Type := Null,
  ����id_In     ����Ԥ����¼.����id%Type := Null,
  ����Ʊ��_In   Number := 0
) As
  --���ܣ�ɾ��һ�������շѵ��� 
  --������ 
  --        ���_IN           =Ҫ�˷ѵ���Ŀ���,��ʽΪ"1,3,5,6...",ȱʡNULL��ʾ��"δ�˵�"�����С� 
  --        ����Ʊ��_In       =0:ȫ�˻����һ��ȫ��ʱ,�ջ�Ʊ�ݡ� 
  --                           1:�����˷Ѳ�����Ʊ��,ͨ���ش���õ������� 
  --���α�ΪҪ�˷ѵ��ݵ�����ԭʼ��¼ 

  --ҽ��ȫ�˵�ĳ�ֽ������ֽ�Ӷ��������µ����ʱ,�ſ��˴�������,ִ���걾���̺�,��������е������������ 
  Cursor c_Bill Is
    Select a.Id, a.No, a.���ӱ�־, a.�շ�ϸĿid, a.���, a.�۸񸸺�, a.ִ��״̬, a.�շ����, a.����, a.����, a.ҽ�����, j.�������, m.��������,
           Nvl(a.���ӱ�־, 0) As ���
    From ������ü�¼ A, ����ҽ����¼ J, �������� M
    Where a.ҽ����� = j.Id(+) And a.No = No_In And a.��¼���� = 1 And a.��¼״̬ In (1, 3) And a.�շ�ϸĿid + 0 = m.����id(+)
    Order By a.�շ�ϸĿid, a.���;

  --:����ԭʼ�������,��Ӧ�ø��ݵ�ǰ�˷Ѳ������������д���
  -- Decode(Sign(���_In), 0, 999, 9)

  --���α����ڴ���ҩƷ���������� 
  --��Ҫ�ܷ��õ�ִ��״̬,��Ϊ���ڴ˲����� 
  Cursor c_Stock Is
    Select ID, ҩƷid, �ⷿid, ����, ����, ����, ʵ������, ����, ��ҩ��ʽ, ���Ч��, Ч��, ��Ʒ����, �ڲ�����, ����id
    From ҩƷ�շ���¼
    Where NO = No_In And ���� In (8, 24) --@@@ 
          And Mod(��¼״̬, 3) = 1 And ����� Is Null And
          ����id In (Select ID
                   From ������ü�¼
                   Where NO = No_In And Mod(��¼����, 10) = 1 And ��¼״̬ In (1, 3) And �շ���� In ('4', '5', '6', '7') --@@@ 
                         And (Instr(',' || ���_In || ',', ',' || ��� || ',') > 0 Or ���_In Is Null))
    Order By ҩƷid;

  --���α����ڴ���δ��ҩƷ��¼ 
  Cursor c_Spare Is
    Select NO, �ⷿid, ���� From δ��ҩƷ��¼ Where NO = No_In And ���� In (8, 24); --@@@ 

  --�ù�����ڴ�����Ա�ɿ�������˵Ĳ�ͬ���㷽ʽ�Ľ�� 

  n_����id ������ü�¼.����id%Type;
  n_��ӡid Ʊ�ݴ�ӡ����.Id%Type;

  --�����˷Ѽ������ 
  n_ʣ������ Number;
  n_ʣ��Ӧ�� Number;
  n_ʣ��ʵ�� Number;
  n_ʣ��ͳ�� Number;
  n_�������� Number;
  n_׼������ Number;
  n_�˷Ѵ��� Number;

  n_Ӧ�ս�� Number;
  n_ʵ�ս�� Number;
  n_ͳ���� Number;
  n_�ܽ��   Number;
  n_�����˷� Number; --�Ƿ��һ���˷���ȫ���˷�,��ÿ���˷ѹ������жϵõ��� 
  n_��id     ����ɿ����.Id%Type;

  l_����id   t_Numlist := t_Numlist();
  l_ҩƷ�շ� t_Numlist := t_Numlist();
  l_ʹ��id   t_Numlist := t_Numlist();

  l_���     t_Numlist := t_Numlist();
  l_ִ��״̬ t_Numlist := t_Numlist();

  n_Dec   Number;
  d_Date  Date;
  n_Count Number;

  Err_Item Exception;
  v_Err_Msg      Varchar2(255);
  n_����ģʽ     Number(3);
  v_Para         Varchar2(1000);
  n_ҽ��ִ�мƼ� Number;

Begin
  n_��id := Zl_Get��id(����Ա����_In);
  --�Ƿ��Ѿ�ȫ����ȫִ��(ֻ�Ǹõ������ŵ��ݵļ��) 
  Select Nvl(Count(*), 0)
  Into n_Count
  From ������ü�¼
  Where NO = No_In And ��¼���� = 1 And ��¼״̬ In (1, 3) And Nvl(ִ��״̬, 0) <> 1;
  If n_Count = 0 Then
    v_Err_Msg := '�õ����е���Ŀ�Ѿ�ȫ����ȫִ�У�';
    Raise Err_Item;
  End If;

  --δ��ȫִ�е���Ŀ�Ƿ���ʣ������(ֻ�����ŵ��ݵļ��) 
  --ִ��״̬��ԭʼ��¼���ж� 
  Select Nvl(Count(*), 0)
  Into n_Count
  From (Select ���, Sum(����) As ʣ������
         From (Select ��¼״̬, Nvl(�۸񸸺�, ���) As ���, Avg(Nvl(����, 1) * ����) As ����, ����id
                From ������ü�¼
                Where NO = No_In And Mod(��¼����, 10) = 1 And Nvl(���ӱ�־, 0) <> 9 And
                      Nvl(�۸񸸺�, ���) In
                      (Select Nvl(�۸񸸺�, ���)
                       From ������ü�¼
                       Where NO = No_In And Mod(��¼����, 10) = 1 And ��¼״̬ In (1, 3) And Nvl(ִ��״̬, 0) <> 1)
                Group By ��¼״̬, Nvl(�۸񸸺�, ���), ����id)
         Group By ���
         Having Sum(����) <> 0);
  If n_Count = 0 Then
    v_Err_Msg := '�õ�����δ��ȫִ�в�����Ŀʣ������Ϊ��,û�п����˷ѵķ��ã�';
    Raise Err_Item;
  End If;

  --ȷ���Ƿ���ҽ��ִ�мƼ��д�������,�����������,�����ҽ��ִ�мƼ۽����˷�,���򰴾ɷ�ʽ���д���
  Select Count(1)
  Into n_ҽ��ִ�мƼ�
  From ������ü�¼ A, ҽ��ִ�мƼ� B
  Where a.ҽ����� = b.ҽ��id And Mod(a.��¼����, 10) = 1 And a.No = No_In And a.��¼״̬ In (1, 3) And Rownum = 1;

  --------------------------------------------------------------------------------- 
  --���ñ��� 
  If �˷�ʱ��_In Is Not Null Then
    d_Date := �˷�ʱ��_In;
  Else
    Select Sysdate Into d_Date From Dual;
  End If;

  If ����id_In Is Null Then
    Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
  Else
    n_����id := ����id_In;
  End If;

  --���С��λ�� 
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_Dec From Dual;

  --ѭ������ÿ�з���(������Ŀ��) 
  n_�ܽ�� := 0;
  For r_Bill In c_Bill Loop
    If Instr(',' || ���_In || ',', ',' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || ',') > 0 Or ���_In Is Null Then
      If Nvl(r_Bill.ִ��״̬, 0) <> 1 Then
        --��ʣ������,ʣ��Ӧ��,ʣ��ʵ�� 
        Select Sum(Nvl(����, 1) * ����), Sum(Ӧ�ս��), Sum(ʵ�ս��), Sum(ͳ����)
        Into n_ʣ������, n_ʣ��Ӧ��, n_ʣ��ʵ��, n_ʣ��ͳ��
        From ������ü�¼
        Where NO = No_In And Mod(��¼����, 10) = 1 And ��� = r_Bill.���;
      
        If n_ʣ������ = 0 Then
          If ���_In Is Not Null Then
            v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����Ѿ�ȫ���˷ѣ�';
            Raise Err_Item;
          End If;
        Else
          --׼������(��ҩƷ��ĿΪʣ������,ԭʼ����) 
          If Instr(',4,5,6,7,', r_Bill.�շ����) = 0 Or (r_Bill.�շ���� = '4' And Nvl(r_Bill.��������, 0) = 0) Then
            --@@@ 
            --��ҩƷ����(�Ծ���ҽ��ִ��Ϊ׼���м��) 
            --: 1.����ҽ�����͵�,����ҽ��ִ��Ϊ׼(�����ܰ���:���;����;����;������Ѫ) 
            --: 2.������ҽ����,����ʣ������Ϊ׼ 
            n_Count := 0;
            If Instr(',C,D,F,G,K,', ',' || r_Bill.������� || ',') = 0 And r_Bill.������� Is Not Null Then
              If n_ҽ��ִ�мƼ� = 1 Then
                Select Decode(Sign(Sum(����)), -1, 0, Sum(����)), Count(*)
                Into n_׼������, n_Count
                From (Select Max(Decode(a.��¼״̬, 2, 0, a.Id)) As ID, Max(a.ҽ�����) As ҽ��id, Max(a.�շ�ϸĿid) As �շ�ϸĿid,
                              Sum(Nvl(a.����, 1) * Nvl(a.����, 1)) As ����,
                              Sum(Decode(a.��¼״̬, 2, 0, Nvl(a.����, 1) * Nvl(a.����, 1))) As ԭʼ����
                       From ������ü�¼ A, ����ҽ����¼ M
                       Where a.ҽ����� = m.Id And Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And
                             Instr('5,6,7', a.�շ����) = 0 And a.No = No_In And a.��� = r_Bill.��� And Mod(a.��¼����, 10) = 1 And
                             a.��¼״̬ In (1, 2, 3) And a.�۸񸸺� Is Null
                       Group By a.���
                       Union All
                       Select a.Id, a.ҽ����� As ҽ��id, a.�շ�ϸĿid, -1 * b.���� As ��ִ��, 0 ԭʼ����
                       From ������ü�¼ A, ҽ��ִ�мƼ� B, ����ҽ����¼ M
                       Where a.ҽ����� = b.ҽ��id And a.�շ�ϸĿid = b.�շ�ϸĿid + 0 And a.ҽ����� = m.Id And
                             Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And Instr('5,6,7', a.�շ����) = 0 And
                             (Exists
                              (Select 1
                               From ����ҽ��ִ��
                               Where b.ҽ��id = ҽ��id And b.���ͺ� = ���ͺ� And b.Ҫ��ʱ�� = Ҫ��ʱ�� And Nvl(ִ�н��, 0) = 1) Or Exists
                              (Select 1
                               From ����ҽ������
                               Where b.ҽ��id = ҽ��id And b.���ͺ� = ���ͺ� And Nvl(ִ��״̬, 0) = 1)) And Not Exists
                        (Select 1
                              From ����ҽ������
                              Where a.ҽ����� = ҽ��id And a.No = NO And Mod(a.��¼����, 10) = ��¼����) And a.No = No_In And
                             a.��� = r_Bill.��� And Mod(a.��¼����, 10) = 1 And a.��¼״̬ In (1, 3) ��and a.�۸񸸺� Is Null) Q1
                Where Not Exists (Select 1
                       From ҩƷ�շ���¼
                       Where ����id = Q1.Id And Instr(',8,9,10,21,24,25,26,', ',' || ���� || ',') > 0) Having
                 Max(ID) <> 0;
              Else
              
                Select Nvl(Sum(����), 0), Count(*)
                Into n_׼������, n_Count
                From (Select a.ҽ��id, a.�շ�ϸĿid, Nvl(a.����, 1) * Nvl(b.��������, 1) As ����
                       From ����ҽ���Ƽ� A, ����ҽ������ B, ������ü�¼ J, ����ҽ����¼ M
                       Where a.ҽ��id = b.ҽ��id And a.ҽ��id = m.Id And Nvl(b.ִ��״̬, 0) <> 1 And a.ҽ��id = j.ҽ����� And
                             a.�շ�ϸĿid = j.�շ�ϸĿid And j.No = No_In And Mod(j.��¼����, 10) = 1 And j.��� = r_Bill.��� And
                             j.��¼״̬ In (1, 3) And j.�۸񸸺� Is Null And Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And
                             Exists
                        (Select 1
                              From ����ҽ���Ƽ� A
                              Where a.ҽ��id = j.ҽ����� And a.�շ�ϸĿid = j.�շ�ϸĿid And Nvl(a.�շѷ�ʽ, 0) = 0) And Not Exists
                        (Select 1
                              From ҩƷ�շ���¼
                              Where ����id = j.Id And Instr(',8,9,10,21,24,25,26,', ',' || ���� || ',') > 0)
                       Union All
                       Select a.ҽ��id, a.�շ�ϸĿid, -1 * Nvl(a.����, 1) * Nvl(c.��������, 1) As ����
                       From ����ҽ���Ƽ� A, ����ҽ������ B, ����ҽ��ִ�� C, ������ü�¼ J, ����ҽ����¼ M
                       Where a.ҽ��id = b.ҽ��id And b.ҽ��id = c.ҽ��id And b.���ͺ� = c.���ͺ� And a.ҽ��id = m.Id And
                             Nvl(c.ִ�н��, 1) = 1 And Nvl(b.ִ��״̬, 0) <> 1 And a.ҽ��id = j.ҽ����� And a.�շ�ϸĿid = j.�շ�ϸĿid And
                             j.No = No_In And Mod(j.��¼����, 10) = 1 And Nvl(a.�շѷ�ʽ, 0) = 0 And j.��� = r_Bill.��� And
                             j.��¼״̬ In (1, 3) And j.�۸񸸺� Is Null And Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And
                             Not Exists
                        (Select 1
                              From ҩƷ�շ���¼
                              Where ����id = j.Id And Instr(',8,9,10,21,24,25,26,', ',' || ���� || ',') > 0) And Not Exists
                        (Select 1 From �������� Where ����id = j.�շ�ϸĿid And Nvl(��������, 0) = 1)
                       Union All
                       Select a.ҽ����� As ҽ��id, a.�շ�ϸĿid, Nvl(a.����, 1) * a.���� As ����
                       From ������ü�¼ A, ����ҽ����¼ M
                       Where a.ҽ����� = m.Id And Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And a.No = No_In And
                             Mod(a.��¼����, 10) = 1 And a.��� = r_Bill.��� And a.��¼״̬ = 2 And a.�۸񸸺� Is Null And Not Exists
                        (Select 1 From ҩƷ�շ���¼ Where NO = No_In And ���� In (8, 24) And ҩƷid = a.�շ�ϸĿid));
              End If;
            End If;
            If Nvl(n_Count, 0) <> 0 And n_׼������ = 0 Then
              v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з�����ִ��,�������˷ѣ�';
              Raise Err_Item;
            End If;
          
            If Nvl(n_Count, 0) = 0 Then
              n_׼������ := n_ʣ������;
            End If;
          
          Else
            Select Nvl(Sum(Nvl(����, 1) * ʵ������), 0), Count(*)
            Into n_׼������, n_Count
            From ҩƷ�շ���¼
            Where NO = No_In And ���� In (8, 24) And Mod(��¼״̬, 3) = 1 --@@@ 
                  And ����� Is Null And ����id = r_Bill.Id;
          
            --��ʣ��������׼������������������� 
            --1.���������õ������޶�Ӧ���շ���¼,��ʱʹ��ʣ������ 
            --2.��������,��ʱ�ѷ�ҩ���� 
            If n_׼������ = 0 Then
              If r_Bill.�շ���� = '4' Then
                If n_Count > 0 Then
                  v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����ѷ���,�����Ϻ����˷ѣ�';
                  Raise Err_Item;
                Else
                  n_׼������ := n_ʣ������;
                End If;
              Else
                v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����ѷ�ҩ,����ҩ�����˷ѣ�';
                Raise Err_Item;
              End If;
            End If;
          End If;
        
          If n_׼������ > n_ʣ������ Then
            v_Err_Msg := '����[' || No_In || '] �е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з��õ��˷�����(' || n_׼������ ||
                         ')������ʣ������(' || n_ʣ������ || ')���������˷ѣ�';
            Raise Err_Item;
          End If;
          If n_׼������ < 0 Then
            v_Err_Msg := '����[' || No_In || '] �е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з��õ��˷�����(' || n_׼������ ||
                         ')С�����㣬�������˷ѣ�';
            Raise Err_Item;
          End If;
        
          --�Ƿ񲿷��˷� 
          If r_Bill.ִ��״̬ = 2 Or n_׼������ <> Nvl(r_Bill.����, 1) * r_Bill.���� Then
            n_�����˷� := 0;
          End If;
        
          --�ñ���Ŀ�ڼ����˷� 
          Select Nvl(Max(Abs(ִ��״̬)), 0) + 1
          Into n_�˷Ѵ���
          From ������ü�¼
          Where NO = No_In And Mod(��¼����, 10) = 1 And ��¼״̬ = 2 And Nvl(ִ��״̬, 0) < 0 And ��� = r_Bill.���;
        
          --���=ʣ����*(׼����/ʣ����) 
          n_Ӧ�ս�� := Round(n_ʣ��Ӧ�� * (n_׼������ / n_ʣ������), n_Dec);
          n_ʵ�ս�� := Round(n_ʣ��ʵ�� * (n_׼������ / n_ʣ������), n_Dec);
          n_ͳ���� := Round(n_ʣ��ͳ�� * (n_׼������ / n_ʣ������), n_Dec);
          n_�ܽ��   := n_�ܽ�� + n_ʵ�ս��;
        
          --�����˷Ѽ�¼ 
          Insert Into ������ü�¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�, ���˿���id, �շ����,
             �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ������, ִ����,
             ִ��״̬, ����״̬, ִ��ʱ��, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, ����,
             �ɿ���id)
            Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�,
                   ���˿���id, �շ����, �շ�ϸĿid, ���㵥λ, Decode(Sign(n_׼������ - Nvl(����, 1) * ����), 0, ����, 1), ��ҩ����,
                   Decode(Sign(n_׼������ - Nvl(����, 1) * ����), 0, -1 * ����, -1 * n_׼������), �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����,
                   -1 * n_Ӧ�ս��, -1 * n_ʵ�ս��, ��������id, ������, ִ�в���id, ������, ִ����, -1 * n_�˷Ѵ���, 1, ִ��ʱ��, ����Ա���_In, ����Ա����_In,
                   ����ʱ��, d_Date, n_����id, -1 * n_ʵ�ս��, ������Ŀ��, ���մ���id, -1 * n_ͳ����, Nvl(�˷�ժҪ_In, ժҪ),
                   Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, ����, n_��id
            From ������ü�¼
            Where ID = r_Bill.Id;
        
          --���ԭ���ü�¼ 
          l_���.Extend;
          l_���(l_���.Count) := r_Bill.���;
          l_ִ��״̬.Extend;
          l_ִ��״̬(l_ִ��״̬.Count) := Case
                                    When Sign(n_׼������ - n_ʣ������) = 0 Then
                                     0
                                    Else
                                     1
                                  End;
        
          --          Update ������ü�¼ Set ��¼״̬ = 3 Where ID = r_Bill.Id;
        
        End If;
      Else
        If ���_In Is Not Null Then
          v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����Ѿ���ȫִ��,�����˷ѣ�';
          Raise Err_Item;
        End If;
        --���:û�޶��к�,ԭʼ�����а����Ѿ���ȫִ�е� 
        n_�����˷� := 0;
      End If;
    Else
      n_�����˷� := 0; --δָ���ñ�,���ڲ����˷� 
    End If;
  End Loop;
  --���ԭ���ü�¼ 
  Forall I In 1 .. l_���.Count
    Update ������ü�¼
    Set ��¼״̬ = 3, ִ��״̬ = l_ִ��״̬(I)
    Where Mod(��¼����, 10) = 1 And NO = No_In And ��� = l_���(I) And ��¼״̬ In (1, 3);

  l_���.Delete;
  For c_���� In (Select Distinct b.����id
               From ������ü�¼ A, ����Ԥ����¼ B
               Where a.����id = b.����id And a.No = No_In And Mod(a.��¼����, 10) = 1 And a.��¼״̬ In (1, 3) And
                     Nvl(b.��¼״̬, 0) = 1) Loop
    l_���.Extend;
    l_���(l_���.Count) := c_����.����id;
  End Loop;

  Forall I In 1 .. l_���.Count
    Update ����Ԥ����¼ Set ��¼״̬ = 3 Where ����id = l_���(I) And Mod(��¼����, 10) <> 1;

  --------------------------------------------------------------------------------- 
  --�˷�Ʊ�ݻ���(��ȫ��ʱ�Ż���,�����������ش�����л���) 
  If ����Ʊ��_In = 1 Then
  
    --���ñ�־||NO;ִ�п���(����);�վݷ�Ŀ(��ҳ����,����);�շ�ϸĿ(����)
    v_Para     := Nvl(zl_GetSysParameter('Ʊ�ݷ������', 1121), '0||0;0;0,0;0');
    n_����ģʽ := Zl_To_Number(Substr(v_Para, 1, 1));
    If n_����ģʽ <> 0 Then
      --�ջ�Ʊ��
      Select ʹ��id Bulk Collect
      Into l_ʹ��id
      From (Select Distinct b.ʹ��id From Ʊ�ݴ�ӡ��ϸ B Where b.No = No_In And Nvl(b.Ʊ��, 0) = 1);
    
      n_����ģʽ := l_ʹ��id.Count;
      If l_ʹ��id.Count <> 0 Then
        --������ռ�¼
        Forall I In 1 .. l_ʹ��id.Count
          Insert Into Ʊ��ʹ����ϸ
            (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ����, ʹ��ʱ��)
            Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, ����Ա����_In, d_Date
            From Ʊ��ʹ����ϸ A
            Where ID = l_ʹ��id(I) And ���� = 1 And Not Exists
             (Select 1 From Ʊ��ʹ����ϸ Where ���� = a.���� And Ʊ�� = a.Ʊ�� And Nvl(����, 0) <> 1);
      
        Forall I In 1 .. l_ʹ��id.Count
          Update Ʊ�ݴ�ӡ��ϸ Set �Ƿ���� = 1 Where ʹ��id = l_ʹ��id(I) And Nvl(�Ƿ����, 0) = 0;
      
      End If;
    End If;
    If n_����ģʽ = 0 Then
      --��ȡ�������һ�εĴ�ӡID(�����Ƕ��ŵ����շѴ�ӡ) 
      Begin
        Select ID
        Into n_��ӡid
        From (Select b.Id
               From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
               Where a.��ӡid = b.Id And a.���� = 1 And b.�������� = 1 And b.No = No_In
               Order By a.ʹ��ʱ�� Desc)
        Where Rownum < 2;
      Exception
        When Others Then
          Null;
      End;
      --������ǰû�д�ӡ,���ջ� 
      If n_��ӡid Is Not Null Then
        --a.���ŵ���ѭ������ʱֻ���ջ�һ�� 
        Select Count(*) Into n_Count From Ʊ��ʹ����ϸ Where Ʊ�� = 1 And ���� = 2 And ��ӡid = n_��ӡid;
        If n_Count = 0 Then
          Insert Into Ʊ��ʹ����ϸ
            (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����)
            Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_Date, ����Ա����_In
            From Ʊ��ʹ����ϸ
            Where ��ӡid = n_��ӡid And Ʊ�� = 1 And ���� = 1;
        Else
          --b.�����˷Ѷ���ջ�ʱ,���һ��ȫ���ջ�Ҫ�ſ����ջص� 
          Insert Into Ʊ��ʹ����ϸ
            (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����)
            Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_Date, ����Ա����_In
            From Ʊ��ʹ����ϸ A
            Where ��ӡid = n_��ӡid And Ʊ�� = 1 And ���� = 1 And Not Exists
             (Select 1 From Ʊ��ʹ����ϸ B Where a.���� = b.���� And ��ӡid = n_��ӡid And Ʊ�� = 1 And ���� = 2);
        End If;
      End If;
    End If;
  End If;

  --------------------------------------------------------------------------------- 
  --�������� 
  For v_���� In (Select ID, ҩƷid, �ⷿid, ����, ����, ����, ʵ������, ����, ��ҩ��ʽ, ���Ч��, Ч��, ��Ʒ����, �ڲ�����, ����id
               From ҩƷ�շ���¼
               Where ���� = 21 --@@@ 
                     And Mod(��¼״̬, 3) = 1 And ����� Is Null And
                     ����id In (Select ID
                              From ������ü�¼
                              Where NO = No_In And Mod(��¼����, 10) = 1 And ��¼״̬ In (1, 3) And �շ���� = '4' --@@@ 
                                    And (Instr(',' || ���_In || ',', ',' || ��� || ',') > 0 Or ���_In Is Null))
               Order By ҩƷid) Loop
    --����ҩƷ��� 
    If v_����.�ⷿid Is Not Null Then
      Update ҩƷ���
      Set �������� = Nvl(��������, 0) + Decode(v_����.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(v_����.����, 1) * Nvl(v_����.ʵ������, 0)
      Where �ⷿid = v_����.�ⷿid And ҩƷid = v_����.ҩƷid And Nvl(����, 0) = Nvl(v_����.����, 0) And ���� = 1;
      If Sql%RowCount = 0 Then
        Insert Into ҩƷ���
          (�ⷿid, ҩƷid, ����, ����, Ч��, ��������, �ϴ�����, �ϴβ���, ���Ч��, ��Ʒ����, �ڲ�����) --@@@ 
        Values
          (v_����.�ⷿid, v_����.ҩƷid, 1, v_����.����, v_����.Ч��,
           Decode(v_����.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(v_����.����, 1) * Nvl(v_����.ʵ������, 0), v_����.����, v_����.����, v_����.���Ч��,
           v_����.��Ʒ����, v_����.�ڲ�����);
      End If;
    End If;
  
    l_����id.Extend;
    l_����id(l_����id.Count) := v_����.����id;
    l_ҩƷ�շ�.Extend;
    l_ҩƷ�շ�(l_ҩƷ�շ�.Count) := v_����.Id;
  End Loop;

  --ҩƷ������� 
  For r_Stock In c_Stock Loop
    --����ҩƷ��� 
    If r_Stock.�ⷿid Is Not Null Then
    
      Select Decode(Count(Column_Value), Null, 0, 0, 0, 1)
      Into n_��������
      From Table(l_����id)
      Where Column_Value = r_Stock.����id;
      If Nvl(n_��������, 0) = 0 Then
      
        Update ҩƷ���
        Set �������� = Nvl(��������, 0) + Decode(r_Stock.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(r_Stock.����, 1) * Nvl(r_Stock.ʵ������, 0)
        Where �ⷿid = r_Stock.�ⷿid And ҩƷid = r_Stock.ҩƷid And Nvl(����, 0) = Nvl(r_Stock.����, 0) And ���� = 1;
        If Sql%RowCount = 0 Then
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, Ч��, ��������, �ϴ�����, �ϴβ���, ���Ч��, ��Ʒ����, �ڲ�����) --@@@ 
          Values
            (r_Stock.�ⷿid, r_Stock.ҩƷid, 1, r_Stock.����, r_Stock.Ч��,
             Decode(r_Stock.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(r_Stock.����, 1) * Nvl(r_Stock.ʵ������, 0), r_Stock.����, r_Stock.����,
             r_Stock.���Ч��, r_Stock.��Ʒ����, r_Stock.�ڲ�����);
        End If;
      End If;
    End If;
    l_ҩƷ�շ�.Extend;
    l_ҩƷ�շ�(l_ҩƷ�շ�.Count) := r_Stock.Id;
  End Loop;

  --ɾ��ҩƷ�շ���¼ 
  Forall I In 1 .. l_ҩƷ�շ�.Count
    Delete From ҩƷ�շ���¼ Where ID = l_ҩƷ�շ�(I) And ����� Is Null;

  --δ��ҩƷ��¼ 
  For r_Spare In c_Spare Loop
    Select Nvl(Count(*), 0)
    Into n_Count
    From ҩƷ�շ���¼
    Where NO = No_In And ���� = r_Spare.���� --@@@ 
          And Mod(��¼״̬, 3) = 1 And ����� Is Null And Nvl(�ⷿid, 0) = Nvl(r_Spare.�ⷿid, 0);
  
    If n_Count = 0 Then
      Delete From δ��ҩƷ��¼
      Where ���� = r_Spare.���� --@@@ 
            And NO = No_In And Nvl(�ⷿid, 0) = Nvl(r_Spare.�ⷿid, 0);
    End If;
  End Loop;
  --ҽ������
  --����_In    Integer:=0, --0:����;1-סԺ
  --����_In    Integer:=1, --1-�շѵ�;2-���ʵ�
  --����_In    Integer:=0, --0:ɾ�����۵�;1-�շѻ����;2-�˷ѻ�����
  --No_In      ������ü�¼.No%Type,
  --ҽ��ids_In Varchar2 := Null
  Zl_ҽ������_�Ʒ�״̬_Update(0, 1, 2, No_In);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����շѼ�¼_����;
/

--92736:������,2016-02-25,�˷�����ģʽ���񴰴���
Create Or Replace Procedure Zl_Third_Charge_Delcheck
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is

  -------------------------------------------------------------------------------------------------- 
  --����:�����˷Ѽ�� 
  --���:Xml_In: 
  --<IN>
  --    <BRID>����ID</BRID>
  --    <JE></JE> //�˿��ܽ��
  --    <JSKLB></JSKLB>     //���㿨���
  --    <TFZY>�˷�ժҪ</TFZY>
  --    <JCFP>1</JCFP>      //��鷢Ʊ,0-�����;1-���;Ϊ1ʱ����ӡ�˷�Ʊ�ĵ��ݲ����˷�
  --    <FYLIST>
  --        <FY>
  --           <DJH>�˿�ݺ�</DJH>
  --           <XH>�˿����(��ʽ:1,2,3..Ϊ�մ�����ʣ������)</DJH>
  --        <FY>
  --    </FYLIST>
  --    <TKLIST>
  --        <TK>
  --            <TKKLB>�˿���</TKKLB>
  --            <TKKH>�˿��</TKKH>
  --            <TKFS>�˿ʽ</TKFS> //�˿ʽ:�ֽ�;֧Ʊ,�����������,���Դ���
  --            <TKJE>֧�����</TKJE>
  --            <JYLSH>������ˮ��</JYLSH>
  --            <TKZY>ժҪ</TKZY>
  --            <TYJK>�˻�Ԥ����</TYJK> //�ʳ�Ԥ��ʱ,ֻ��JSJE�ڵ�:1-��Ԥ��
  --            <SFXFK>�Ƿ����ѿ�</SFXFK>   //(1-�����ѿ�),���ѿ�ʱ,������㿨���,���㿨��,������Ƚӵ�
  --            <EXPENDLIST>  //��չ������Ϣ
  --                <EXPEND>
  --                    <JYMC>��������</JYMC>
  --                    <JYLR>��������</JYLR>
  --                </EXPEND>
  --            </EXPENDLIST>
  --        </TK>
  --    </TKLIST>
  --</IN>

  --����:Xml_Out 
  --  <OUT> 
  --    �D�D�������д�������˵��ͨ�����
  --    <ERROR> 
  --      <MSG>������Ϣ</MSG> 
  --    </ERROR> 
  --  </OUT> 
  -------------------------------------------------------------------------------------------------- 
  n_�˿��ܶ� ������ü�¼.ʵ�ս��%Type;

  n_����id     ������ü�¼.����id%Type;
  n_���ݲ���id ������ü�¼.����id%Type;
  v_����Ա���� ������ü�¼.����Ա���%Type;
  v_����Ա���� ������ü�¼.����Ա����%Type;
  n_����id     ������ü�¼.����id%Type;
  n_ԭ������� ����Ԥ����¼.�������%Type;
  v_���㿨��� Varchar2(100);
  v_���㷽ʽ   ҽ�ƿ����.���㷽ʽ%Type;
  n_�����id   ҽ�ƿ����.Id%Type;

  v_ժҪ     ������ü�¼.ժҪ%Type;
  n_Count    Number(18);
  n_Temp     Number(18);
  n_��鷢Ʊ Number(3);
  n_�Ƿ��ӡ Number(3);
  n_�˷�ģʽ Number(3);

  v_Temp    Varchar2(32767); --��ʱXML 
  x_Templet Xmltype; --ģ��XML 

  v_Err_Msg Varchar2(200);
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  --0.��ȡ����еĲ���ID����Ϣ
  Select To_Number(Extractvalue(Value(A), 'IN/BRID')), To_Number(Extractvalue(Value(A), 'IN/JE')),
         Extractvalue(Value(A), 'IN/TFZY'), To_Number(Extractvalue(Value(A), 'IN/JCFP')),
         Extractvalue(Value(A), 'IN/JSKLB')
  Into n_����id, n_�˿��ܶ�, v_ժҪ, n_��鷢Ʊ, v_���㿨���
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  --0.��ؼ��
  If Nvl(n_����id, 0) = 0 Then
    v_Err_Msg := '������Чʶ�������,�������˷Ѳ���!';
    Raise Err_Item;
  End If;

  n_�˷�ģʽ := zl_GetSysParameter('�����˷���������');
  If Nvl(n_�˷�ģʽ, 0) = 1 Then
    v_Err_Msg := '��ǰΪ�˷�����ģʽ,����������˷�!';
    Raise Err_Item;
  End If;

  If v_���㿨��� Is Not Null Then
    Begin
      n_�����id := To_Number(v_���㿨���);
    Exception
      When Others Then
        n_�����id := 0;
    End;
    If n_�����id = 0 Then
      Begin
        Select ID Into n_�����id From ҽ�ƿ���� Where ���� = v_���㿨���;
      Exception
        When Others Then
          v_Err_Msg := '�޷�ȷ�ϴ���Ľ��㿨��';
          Raise Err_Item;
      End;
    End If;
  Else
    n_�����id := 0;
  End If;
  
  If Nvl(n_�����id, 0) <> 0 Then
    Select ���㷽ʽ Into v_���㷽ʽ From ҽ�ƿ���� Where ID = n_�����id;
  End If;

  --��Աid,��Ա���,��Ա���� 
  v_Temp := Zl_Identity(1);
  If Nvl(v_Temp, '0') = '0' Or Nvl(v_Temp, '_') = '_' Then
    v_Err_Msg := 'ϵͳ�����ϱ���Ч�Ĳ���Ա,�������˷�!';
    Raise Err_Item;
  End If;
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_����Ա���� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_����Ա���� := v_Temp;
  v_Err_Msg    := Null;

  --1.�˷Ѽ��

  n_Count      := 0;
  n_ԭ������� := 0;
  For c_���� In (Select Extractvalue(b.Column_Value, '/FY/DJH') As ���ݺ�, Extractvalue(b.Column_Value, '/FY/XH') As �˿����
               From Table(Xmlsequence(Extract(Xml_In, '/IN/FYLIST/FY'))) B) Loop
  
    If c_����.���ݺ� Is Null Then
      v_Err_Msg := 'δȷ��ָ���˷ѵĵ��ݺ�,�����˷�!';
      Raise Err_Item;
    End If;
    Begin
      Select a.�������, a.����id, a.����id
      Into n_Temp, n_����id, n_���ݲ���id
      From ����Ԥ����¼ A, ������ü�¼ B
      Where a.����id = b.����id And b.No = c_����.���ݺ� And b.��¼���� = 1 And Nvl(b.����״̬, 0) = 0 And b.��¼״̬ In (1, 3) And
            Rownum < 2;
    Exception
      When Others Then
        n_Temp := Null;
    End;
  
    If n_Temp Is Null Then
      v_Err_Msg := 'ָ���ĵ��ݺ�:' || c_����.���ݺ� || 'δ�ҵ�,�����˷�!';
      Raise Err_Item;
    End If;
  
    If Nvl(n_���ݲ���id, 0) = 0 Then
      Begin
        Select ����id
        Into n_���ݲ���id
        From ������ü�¼
        Where NO = c_����.���ݺ� And ��¼���� = 1 And Nvl(����״̬, 0) = 0 And ��¼״̬ In (1, 3) And Rownum < 2;
      Exception
        When Others Then
          n_���ݲ���id := 0;
      End;
    End If;
  
    If Nvl(n_����id, 0) <> Nvl(n_���ݲ���id, 0) Then
      v_Err_Msg := '�����˷ѵ��շѵ�:' || c_����.���ݺ� || '���ǵ�ǰ���˵��շѵ�,�����˷�!';
      Raise Err_Item;
    End If;
  
    If n_ԭ������� <> 0 And n_ԭ������� <> n_Temp Then
      v_Err_Msg := '�����˷ѵĵ��ݺŲ���һ���շѽ���,�����˷�!';
      Raise Err_Item;
    End If;
    n_ԭ������� := n_Temp;
  
    Select Count(*) Into n_Temp From ���ò����¼ Where �շѽ���id = n_����id;
    If Nvl(n_Temp, 0) <> 0 Then
      v_Err_Msg := '�����˷ѵĵ��ݺ��Ѿ������˱��ղ������,�����˷�!';
      Raise Err_Item;
    End If;
  
    If v_���㿨��� Is Not Null Then
      Select Count(*) Into n_Temp From ����Ԥ����¼ Where ����id = n_����id And ���㷽ʽ = v_���㷽ʽ;
      If Nvl(n_Temp, 0) = 0 Then
        v_Err_Msg := '�����˷ѵĵ��ݲ���' || v_���㷽ʽ || '�����,�����˷�!';
        Raise Err_Item;
      End If;
    End If;
  
    If Nvl(n_��鷢Ʊ, 0) = 1 Then
      Select Max(Decode(a.ʵ��Ʊ��, Null, 0, 1))
      Into n_�Ƿ��ӡ
      From ������ü�¼ A
      Where NO = c_����.���ݺ� And ��¼���� = 1;
      If Nvl(n_�Ƿ��ӡ, 0) = 1 Then
        v_Err_Msg := '�����˷ѵĵ��ݺ��ѿ���Ʊ,�����˷�!';
        Raise Err_Item;
      End If;
    End If;
  
    n_Count := n_Count + 1;
  End Loop;

  If n_Count = 0 Then
    v_Err_Msg := 'δȷ��������Ҫ�˷ѵĵ���,�����˷�!';
    Raise Err_Item;
  End If;

  --2.֧����ʽ���
  n_Count := 0;
  For c_���㷽ʽ In (Select Extractvalue(b.Column_Value, '/TK/TKKLB') As �����, Extractvalue(b.Column_Value, '/TK/TKKH') As ����,
                        Extractvalue(b.Column_Value, '/TK/TKFS') As ���㷽ʽ,
                        -1 * To_Number(Extractvalue(b.Column_Value, '/TK/TKJE')) As �˿���,
                        Extractvalue(b.Column_Value, '/TK/JYLSH') As ������ˮ��,
                        Extractvalue(b.Column_Value, '/TK/JYSM') As ����˵��, Extractvalue(b.Column_Value, '/TK/TKZY') As ժҪ,
                        Extractvalue(b.Column_Value, '/TK/TYJK') As �Ƿ���Ԥ��,
                        Extractvalue(b.Column_Value, '/TK/SFXFK') As �Ƿ����ѿ�,
                        Extract(b.Column_Value, '/TK/EXPENDLIST') As Expend
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/TKLIST/TK'))) B) Loop
  
    --1.�˻�������
    If c_���㷽ʽ.����� Is Not Null And Nvl(c_���㷽ʽ.�Ƿ����ѿ�, 0) = 0 Then
      --1.����������
      Null;
    Elsif c_���㷽ʽ.����� Is Not Null And Nvl(c_���㷽ʽ.�Ƿ����ѿ�, 0) = 1 Then
      --2.���ѿ�����
      Null;
    Elsif Nvl(c_���㷽ʽ.�Ƿ���Ԥ��, 0) = 1 Then
      --3.��Ԥ����
      Null;
    Else
      --4.��ͨ����
      If c_���㷽ʽ.���㷽ʽ Is Null Then
        v_Err_Msg := 'δָ��֧����ʽ,�����˷�!';
        Raise Err_Item;
      End If;
    End If;
    n_Count := n_Count + 1;
  End Loop;

  If n_Count = 0 Then
    v_Err_Msg := '������Чȷ�ϵ�ǰ��֧����ʽ,�����˷�!';
    Raise Err_Item;
  End If;

  Xml_Out := x_Templet;

Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Charge_Delcheck;
/

--93623:���˺�,2016-02-26,����ҽ�����͵ļƷ�״̬
Create Or Replace Procedure Zl_ҽ������_�Ʒ�״̬_Update
(
  ����_In    Integer := 0, --0:����;1-סԺ
  ����_In    Integer := 1, --1-�շѵ�;2-���ʵ�
  ����_In    Integer := 0, --0:ɾ�����۵�;1-�շѻ����;2-�˷ѻ�����
  No_In      ������ü�¼.No%Type,
  ҽ��ids_In Varchar2 := Null
) As
  Err_Item Exception;
  v_Err_Msg Varchar2(255);

  l_ҽ��id   t_Numlist := t_Numlist();
  l_ҽ��id1  t_Numlist := t_Numlist();
  l_�Ʒ�״̬ t_Numlist := t_Numlist();
  n_Count    Number(18);

Begin
  If ҽ��ids_In Is Null Then
    If ����_In = 0 Then
      Select ID Bulk Collect
      Into l_ҽ��id1
      From (Select Distinct ҽ����� As ID
             From ������ü�¼
             Where NO = No_In And ��¼���� = ����_In And ��¼״̬ In (0, 1, 3) And ҽ����� Is Not Null);
    Else
      Select ID Bulk Collect
      Into l_ҽ��id1
      From (Select Distinct ҽ����� As ID
             From סԺ���ü�¼
             Where NO = No_In And ��¼���� = ����_In And ��¼״̬ In (0, 1, 3) And ҽ����� Is Not Null);
    End If;
  Else
    Select ID Bulk Collect
    Into l_ҽ��id1
    From (Select Distinct Column_Value As ID From Table(f_Str2list(ҽ��ids_In)) B);
  End If;
  If l_ҽ��id1.Count = 0 Then
    Return;
  End If;

  For c_ҽ�� In (With c_ҽ����Ϣ As
                  (Select Column_Value As ҽ��id From Table(l_ҽ��id1))
                 Select ID, ���id, �������
                 From ����ҽ����¼ A
                 Where a.Id In (Select ҽ��id
                                From c_ҽ����Ϣ
                                Union All
                                Select Distinct ���id
                                From ����ҽ����¼ A1, c_ҽ����Ϣ B1
                                Where A1.Id = B1.ҽ��id And ���id Is Not Null)) Loop
    --1.���۵�ɾ��
    If Nvl(����_In, 0) = 0 Then
      --D    ���       JC
      --F    ����       SS
      --:-1-����Ʒ�(ͨ����ִ�к�Ժ��ִ�еĶ�����Ʒ�);0-δ�Ʒ�;1-�ѼƷѣ�2-�����˷�(����)��3-ȫ���շ�(�������շ���)��4-ȫ���˷�(����)
      If c_ҽ��.���id Is Null And Instr(',D,F,', ',' || c_ҽ��.������� || ',') > 0 Then
        If ����_In = 0 Then
          Select Nvl(Count(*), 0)
          Into n_Count
          From ������ü�¼ A, (Select ID From ����ҽ����¼ A Where ID = c_ҽ��.Id Or ���id = c_ҽ��.Id) B, ����ҽ������ C
          Where a.ҽ����� = b.Id And a.No = No_In And a.��¼���� = ����_In And a.No = c.No And a.��¼���� = c.��¼���� And c.�Ʒ�״̬ = 1 And
                Nvl(a.���ӱ�־, 0) <> 9;
        Else
          Select Nvl(Count(*), 0)
          Into n_Count
          From סԺ���ü�¼ A, (Select ID From ����ҽ����¼ A Where ID = c_ҽ��.Id Or ���id = c_ҽ��.Id) B, ����ҽ������ C
          Where a.ҽ����� = b.Id And a.No = No_In And a.��¼���� = ����_In And a.No = c.No And a.��¼���� = c.��¼���� And c.�Ʒ�״̬ = 1 And
                Nvl(a.���ӱ�־, 0) <> 9;
        End If;
      Else
        If ����_In = 0 Then
          Select Nvl(Count(*), 0)
          Into n_Count
          From ������ü�¼ A, (Select ID From ����ҽ����¼ A Where ID = c_ҽ��.Id) B, ����ҽ������ C
          Where a.ҽ����� = b.Id And a.No = No_In And a.��¼���� = ����_In And a.No = c.No And a.��¼���� = c.��¼���� And c.�Ʒ�״̬ = 1 And
                Nvl(a.���ӱ�־, 0) <> 9;
        Else
          Select Nvl(Count(*), 0)
          Into n_Count
          From סԺ���ü�¼ A, (Select ID From ����ҽ����¼ A Where ID = c_ҽ��.Id Or ���id = c_ҽ��.Id) B, ����ҽ������ C
          Where a.ҽ����� = b.Id And a.No = No_In And a.��¼���� = ����_In And a.No = c.No And a.��¼���� = c.��¼���� And c.�Ʒ�״̬ = 1 And
                Nvl(a.���ӱ�־, 0) <> 9;
        End If;
      End If;
      If Nvl(n_Count, 0) = 0 Then
        l_ҽ��id.Extend;
        l_ҽ��id(l_ҽ��id.Count) := c_ҽ��.Id;
        l_�Ʒ�״̬.Extend;
        l_�Ʒ�״̬(l_�Ʒ�״̬.Count) := 0; --δ�Ʒ�
      
      End If;
    End If;
  
    --1.�շѻ����
    If Nvl(����_In, 0) = 1 Then
      --:-1-����Ʒ�(ͨ����ִ�к�Ժ��ִ�еĶ�����Ʒ�);0-δ�Ʒ�;1-�ѼƷѣ�2-�����˷�(����)��3-ȫ���շ�(�������շ���)��4-ȫ���˷�(����)
      n_Count := 0;
      If c_ҽ��.���id Is Null And Instr(',D,F,', ',' || c_ҽ��.������� || ',') = 0 Then
        n_Count := 1;
      End If;
      If Nvl(n_Count, 0) = 0 Then
        l_ҽ��id.Extend;
        l_ҽ��id(l_ҽ��id.Count) := c_ҽ��.Id;
        l_�Ʒ�״̬.Extend;
        l_�Ʒ�״̬(l_�Ʒ�״̬.Count) := 3; --ȫ���շ�
      
      End If;
    End If;
  
    --2.�˷ѻ�����
    If Nvl(����_In, 0) = 2 Then
      --:-1-����Ʒ�(ͨ����ִ�к�Ժ��ִ�еĶ�����Ʒ�);0-δ�Ʒ�;1-�ѼƷѣ�2-�����˷�(����)��3-ȫ���շ�(�������շ���)��4-ȫ���˷�(����)
      If c_ҽ��.���id Is Null And Instr(',D,F,', ',' || c_ҽ��.������� || ',') > 0 Then
        If ����_In = 0 Then
          Select Case
                   When Min(״̬) = Max(״̬) Then
                    Max(״̬)
                   Else
                    3
                 End
          Into n_Count
          From (
                 
                 Select Case
                            When ʣ������ = ԭʼ���� And ԭʼ���� <> 0 Then
                             0
                            When ʣ������ = 0 Then
                             2
                            Else
                             3
                          End As ״̬
                 
                 From (Select ���, Sum(Nvl(Nvl(����, 1) * ����, 0)) As ʣ������,
                                Nvl(Sum(Decode(��¼״̬, 1, 1, 3, 1, 0) * Decode(��¼����, 11, 0, 1)  * Nvl(Nvl(����, 1) * ����, 0)), 0) As ԭʼ����
                         From ������ü�¼
                         Where Decode(����_In, 2, ��¼����, Mod(��¼����, 10)) = ����_In And NO = No_In And �۸񸸺� Is Null And
                               Nvl(���ӱ�־, 0) <> 9 And
                               ҽ����� + 0 In (Select ID From ����ҽ����¼ Where ID = c_ҽ��.Id Or ���id = c_ҽ��.Id)
                         Group By ���));
        Else
          Select Case
                   When Min(״̬) = Max(״̬) Then
                    Max(״̬)
                   Else
                    3
                 End
          Into n_Count
          From (
                 
                 Select Case
                           When ʣ������ = ԭʼ���� And ԭʼ���� <> 0 Then
                            0
                           When ʣ������ = 0 Then
                            2
                           Else
                            3
                         End As ״̬
                 
                 From (Select ���, Sum(Nvl(Nvl(����, 1) * ����, 0)) As ʣ������,
                                Nvl(Sum(Decode(��¼״̬, 1, 1, 3, 1, 0) * Decode(��¼����, 11, 0, 1) * Nvl(Nvl(����, 1) * ����, 0)), 0) As ԭʼ����
                         From סԺ���ü�¼
                         Where Decode(����_In, 2, ��¼����, Mod(��¼����, 10)) = ����_In And NO = No_In And �۸񸸺� Is Null And
                               Nvl(���ӱ�־, 0) <> 9 And
                               ҽ����� + 0 In (Select ID From ����ҽ����¼ Where ID = c_ҽ��.Id Or ���id = c_ҽ��.Id)
                         Group By ���));
        
        End If;
      Else
        If ����_In = 0 Then
          Select Case
                   When Min(״̬) = Max(״̬) Then
                    Max(״̬)
                   Else
                    3
                 End
          Into n_Count
          From (
                 
                 Select Case
                            When ʣ������ = ԭʼ���� And ԭʼ���� <> 0 Then
                             0
                            When ʣ������ = 0 Then
                             2
                            Else
                             3
                          End As ״̬
                 From (Select ���, Sum(Nvl(Nvl(����, 1) * ����, 0)) As ʣ������,
                                Nvl(Sum(Decode(��¼״̬, 1, 1, 3, 1, 0) * Decode(��¼����, 11, 0, 1) * Nvl(Nvl(����, 1) * ����, 0)), 0) As ԭʼ����
                         From ������ü�¼
                         Where Decode(����_In, 2, ��¼����, Mod(��¼����, 10)) = ����_In And NO = No_In And �۸񸸺� Is Null And
                               Nvl(���ӱ�־, 0) <> 9 And ҽ����� + 0 = c_ҽ��.Id
                         Group By ���));
        Else
        
          Select Case
                   When Min(״̬) = Max(״̬) Then
                    Max(״̬)
                   Else
                    3
                 End
          Into n_Count
          From (
                 
                 Select Case
                           When ʣ������ = ԭʼ���� And ԭʼ���� <> 0 Then
                            0
                           When ʣ������ = 0 Then
                            2
                           Else
                            3
                         End As ״̬
                 From (Select ���, Sum(Nvl(Nvl(����, 1) * ����, 0)) As ʣ������,
                                Nvl(Sum(Decode(��¼״̬, 1, 1, 3, 1, 0) * Decode(��¼����, 11, 0, 1) * Nvl(Nvl(����, 1) * ����, 0)), 0) As ԭʼ����
                         From סԺ���ü�¼
                         Where Decode(����_In, 2, ��¼����, Mod(��¼����, 10)) = ����_In And NO = No_In And �۸񸸺� Is Null And
                               Nvl(���ӱ�־, 0) <> 9 And ҽ����� + 0 = c_ҽ��.Id
                         Group By ���));
        
        End If;
      
      End If;
    
      If n_Count <> 0 Then
      
        l_ҽ��id.Extend;
        l_ҽ��id(l_ҽ��id.Count) := c_ҽ��.Id;
        l_�Ʒ�״̬.Extend;
        l_�Ʒ�״̬(l_�Ʒ�״̬.Count) := Case
                                  When Nvl(n_Count, 0) = 2 Then
                                   4
                                  Else
                                   2
                                End;
      End If;
    End If;
  End Loop;

  Forall I In 1 .. l_ҽ��id.Count
    Update ����ҽ������ A
    Set a.�Ʒ�״̬ = l_�Ʒ�״̬(I)
    Where ҽ��id = l_ҽ��id(I) And ��¼���� = ����_In And NO = No_In;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҽ������_�Ʒ�״̬_Update;
/

--93623:���˺�,2016-02-25,����ҽ�����͵ļƷ�״̬
Create Or Replace Procedure Zl_���ﻮ�ۼ�¼_Delete
(
  No_In   ������ü�¼.No%Type,
  ���_In Varchar2 := Null --��Ҫ��������ҽ��վ���ϵ���ҩƷ
) As
  --���ܣ�ɾ��һ�����ﻮ�۵���
  --�ù�����ڴ���ҩƷ����������
  Cursor c_Stock Is
    Select ��ҩ��ʽ, �ⷿid, ����, ҩƷid, ʵ������, ����, ���Ч��, ����, ����, Ч��, ID, ��Ʒ����, �ڲ�����, ����id
    From ҩƷ�շ���¼
    Where ���� In (8, 24) And NO = No_In And Mod(��¼״̬, 3) = 1 And ����� Is Null And
          ����id In (Select ID
                   From ������ü�¼
                   Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 0 And �շ���� In ('4', '5', '6', '7') And
                         (Instr(',' || ���_In || ',', ',' || ��� || ',') > 0 Or ���_In Is Null))
    Order By ҩƷid;
  --���α����ڴ�����ü�¼���
  Cursor c_Serial Is
    Select ID, �۸񸸺� From ������ü�¼ Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 0 Order By ���;

  v_ҽ��ids  Varchar2(4000);
  l_ҽ��id   t_Numlist := t_Numlist();
  l_ҩƷ�շ� t_Numlist := t_Numlist();
  v_ҽ��id   ����ҽ����¼.Id%Type;
  l_����id   t_Numlist := t_Numlist();
  n_�������� Number;

  n_����         ������ü�¼.���%Type;
  n_Count        Number;
  n_ҽ����       Number(5);
  n_��ִ��_Count Number;

  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin

  --�Ƿ��Ѿ�ɾ�����շ�
  Select Nvl(Count(ID), 0), Sum(Decode(ҽ�����, Null, 0, 1)), Max(ҽ�����), Sum(Decode(Nvl(ִ��״̬, 0), 1, 1, 2, 1, 0))
  Into n_Count, n_ҽ����, v_ҽ��id, n_��ִ��_Count
  From ������ü�¼
  Where ��¼���� = 1 And ��¼״̬ = 0 And NO = No_In And
        (Instr(',' || ���_In || ',', ',' || Nvl(�۸񸸺�, ���) || ',') > 0 Or ���_In Is Null);

  If n_Count = 0 Then
    v_Err_Msg := 'Ҫɾ���ķ��ü�¼�����ڣ������Ѿ�ɾ�����Ѿ��շѡ�';
    Raise Err_Item;
  End If;
  --�Ƿ��Ѿ�ִ��
  If Nvl(n_��ִ��_Count, 0) > 0 Then
    v_Err_Msg := 'Ҫɾ���ķ��ü�¼�а�����ִ�е����ݣ�';
    Raise Err_Item;
  End If;

  --ҽ�����ã��������ִ�е�ҽ��(ע����ִ�е������������,��Ϊ���� ���_IN ����������ý���������)
  Select Nvl(Count(*), 0)
  Into n_Count
  From ����ҽ������
  Where ִ��״̬ = 3 And (NO, ��¼����, ҽ��id) In
        (Select NO, ��¼����, ҽ�����
                      From ������ü�¼
                      Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 0 And ҽ����� Is Not Null And
                            (Instr(',' || ���_In || ',', ',' || ��� || ',') > 0 Or ���_In Is Null));
  If n_Count > 0 Then
    v_Err_Msg := 'Ҫɾ���ķ����д��ڶ�Ӧ��ҽ������ִ�е����������ɾ����';
    Raise Err_Item;
  End If;

  --ҩƷ�������
  --�ȴ���������
  For v_���� In (Select ��ҩ��ʽ, �ⷿid, ����, ҩƷid, ʵ������, ����, ���Ч��, ����, ����, Ч��, ID, ��Ʒ����, �ڲ�����, ����id
               From ҩƷ�շ���¼
               Where ���� = 21 And Mod(��¼״̬, 3) = 1 And ����� Is Null And
                     ����id In (Select ID
                              From ������ü�¼
                              Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 0 And �շ���� = '4' And
                                    (Instr(',' || ���_In || ',', ',' || ��� || ',') > 0 Or ���_In Is Null))
               Order By ҩƷid) Loop
  
    If v_����.�ⷿid Is Not Null Then
      Update ҩƷ���
      Set �������� = Nvl(��������, 0) + Decode(v_����.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(v_����.����, 1) * Nvl(v_����.ʵ������, 0)
      Where �ⷿid = v_����.�ⷿid And ҩƷid = v_����.ҩƷid And Nvl(����, 0) = Nvl(v_����.����, 0) And ���� = 1;
      If Sql%RowCount = 0 Then
        Insert Into ҩƷ���
          (�ⷿid, ҩƷid, ����, ����, Ч��, ��������, �ϴ�����, �ϴβ���, ���Ч��, ��Ʒ����, �ڲ�����)
        Values
          (v_����.�ⷿid, v_����.ҩƷid, 1, v_����.����, v_����.Ч��,
           Decode(v_����.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(v_����.����, 1) * Nvl(v_����.ʵ������, 0), v_����.����, v_����.����, v_����.���Ч��,
           v_����.��Ʒ����, v_����.�ڲ�����);
      End If;
    End If;
    l_ҩƷ�շ�.Extend;
    l_ҩƷ�շ�(l_ҩƷ�շ�.Count) := v_����.Id;
  
    l_����id.Extend;
    l_����id(l_����id.Count) := v_����.����id;
  End Loop;

  For r_Stock In c_Stock Loop
  
    If r_Stock.�ⷿid Is Not Null Then
    
      Select Decode(Count(Column_Value), Null, 0, 0, 0, 1)
      Into n_��������
      From Table(l_����id)
      Where Column_Value = r_Stock.����id;
      If Nvl(n_��������, 0) = 0 Then
        Update ҩƷ���
        Set �������� = Nvl(��������, 0) + Decode(r_Stock.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(r_Stock.����, 1) * Nvl(r_Stock.ʵ������, 0)
        Where �ⷿid = r_Stock.�ⷿid And ҩƷid = r_Stock.ҩƷid And Nvl(����, 0) = Nvl(r_Stock.����, 0) And ���� = 1;
        If Sql%RowCount = 0 Then
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, Ч��, ��������, �ϴ�����, �ϴβ���, ���Ч��, ��Ʒ����, �ڲ�����)
          Values
            (r_Stock.�ⷿid, r_Stock.ҩƷid, 1, r_Stock.����, r_Stock.Ч��,
             Decode(r_Stock.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(r_Stock.����, 1) * Nvl(r_Stock.ʵ������, 0), r_Stock.����, r_Stock.����,
             r_Stock.���Ч��, r_Stock.��Ʒ����, r_Stock.�ڲ�����);
        End If;
      End If;
    End If;
    l_ҩƷ�շ�.Extend;
    l_ҩƷ�շ�(l_ҩƷ�շ�.Count) := r_Stock.Id;
  End Loop;

  --ɾ��ҩƷ�շ���¼
  Forall I In 1 .. l_ҩƷ�շ�.Count
    Delete From ҩƷ�շ���¼ Where ID = l_ҩƷ�շ�(I);

  ------------------------------------------------------------------------------------------------------------------------
  --����ɾδ��ҩƷ��¼
  Delete From δ��ҩƷ��¼ A
  Where NO = No_In And ���� In (8, 24) And Not Exists
   (Select 1
         From ҩƷ�շ���¼
         Where ���� = a.���� And Nvl(�ⷿid, 0) = Nvl(a.�ⷿid, 0) And NO = No_In And Mod(��¼״̬, 3) = 1 And ����� Is Null);
  ------------------------------------------------------------------------------------------------------------------------

  --ɾ������ҽ������(���һ��ɾ��ʱ)
  If ���_In Is Null Then
    --Begin
    --  Select ҽ�����
    --  Into v_ҽ��id
    --  From ������ü�¼
    --  Where ��¼���� = 1 And ��¼״̬ = 0 And NO = No_In And Rownum = 1;
    -- Exception
    --  When Others Then
    --    Null;
    -- End;
  
    If v_ҽ��id Is Not Null Then
      Delete From ����ҽ������ Where ҽ��id = v_ҽ��id And NO = No_In And ��¼���� = 1;
    End If;
  End If;

  If n_ҽ���� > 0 Then
    If n_ҽ���� = 1 Then
      l_ҽ��id.Extend;
      l_ҽ��id(l_ҽ��id.Count) := v_ҽ��id;
    Else
      Select Distinct ҽ����� Bulk Collect
      Into l_ҽ��id
      From ������ü�¼
      Where ��¼���� = 1 And ��¼״̬ = 0 And NO = No_In And ҽ����� Is Not Null And
            (Instr(',' || ���_In || ',', ',' || Nvl(�۸񸸺�, ���) || ',') > 0 Or ���_In Is Null);
    End If;
  End If;

  --������ü�¼
  Delete From ������ü�¼
  Where ��¼���� = 1 And ��¼״̬ = 0 And NO = No_In And
        (Instr(',' || ���_In || ',', ',' || Nvl(�۸񸸺�, ���) || ',') > 0 Or ���_In Is Null);
  If Sql%RowCount = 0 Then
    v_Err_Msg := 'Ҫɾ���ķ��ü�¼�����ڣ������Ѿ�ɾ�����Ѿ��շѡ�';
    Raise Err_Item;
  End If;

  If ���_In Is Not Null Then
    --���µ���ʣ����÷��ü�¼�����
    n_Count := 1;
    For r_Serial In c_Serial Loop
      If r_Serial.�۸񸸺� Is Null Then
        n_���� := n_Count;
      End If;
      Update ������ü�¼ Set ��� = n_Count, �۸񸸺� = Decode(�۸񸸺�, Null, Null, n_����) Where ID = r_Serial.Id;
    
      n_Count := n_Count + 1;
    End Loop;
  End If;
  v_ҽ��ids := Null;
  For I In 1 .. l_ҽ��id.Count Loop
    v_ҽ��ids := Nvl(v_ҽ��ids, '') || ',' || l_ҽ��id(I);
  End Loop;
  If v_ҽ��ids Is Not Null Then
    v_ҽ��ids := Substr(v_ҽ��ids, 2);
    --����_In    Integer, --0:����;1-סԺ
    --����_In    Integer, --1-�շѵ�;2-���ʵ�
    --����_In    Integer, --0:ɾ�����۵�;1-�շѻ����;2-�˷ѻ�����
    --No_In      ������ü�¼.No%Type,
    --ҽ��ids_In Varchar2
    Zl_ҽ������_�Ʒ�״̬_Update(0, 1, 0, No_In, v_ҽ��ids);
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���ﻮ�ۼ�¼_Delete;
/

--93623:���˺�,2016-02-25,����ҽ�����͵ļƷ�״̬
Create Or Replace Procedure Zl_������շ�_Delete
(
  No_In         ������ü�¼.No%Type,
  ����Ա���_In ������ü�¼.����Ա���%Type,
  ����Ա����_In ������ü�¼.����Ա����%Type
) As
  --���ܣ�ɾ��һ��������շѵ���

  --���α�ΪҪ�˷ѵ��ݵ�����ԭʼ��¼
  Cursor c_Bill Is
    Select a.Id, a.No, a.���ӱ�־, a.�շ�ϸĿid, a.���, a.�۸񸸺�, a.ִ��״̬, a.�շ����, a.����, a.����, a.ҽ�����, j.�������, m.��������
    From ������ü�¼ A, ����ҽ����¼ J, �������� M
    Where a.ҽ����� = j.Id(+) And a.No = No_In And a.��¼���� = 1 And a.��¼״̬ In (1, 3) And a.�շ�ϸĿid + 0 = m.����id(+)
    Order By a.�շ�ϸĿid, a.���;

  --���α����ڴ���ҩƷ����������
  --��Ҫ�ܷ��õ�ִ��״̬,��Ϊ���ڴ˲�����
  Cursor c_Stock Is
    Select ID, ҩƷid, �ⷿid, ����, ����, ����, ʵ������, ����, ��ҩ��ʽ, ���Ч��, Ч��, ��Ʒ����, �ڲ�����, ����id
    From ҩƷ�շ���¼
    Where NO = No_In And ���� In (8, 24) --@@@
          And Mod(��¼״̬, 3) = 1 And ����� Is Null And
          ����id In (Select ID
                   From ������ü�¼
                   Where NO = No_In And ��¼���� = 1 And ��¼״̬ In (1, 3) And �շ���� In ('4', '5', '6', '7'))
    Order By ҩƷid;

  --���α����ڴ���δ��ҩƷ��¼
  Cursor c_Spare Is
    Select NO, �ⷿid, ���� From δ��ҩƷ��¼ Where NO = No_In And ���� In (8, 24); --@@@

  --�ù�����ڴ�����Ա�ɿ�������˵Ĳ�ͬ���㷽ʽ�Ľ��
  Cursor c_Money(����id_In ����Ԥ����¼.����id%Type) Is
    Select ���㷽ʽ, ��Ԥ��
    From ����Ԥ����¼
    Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = ����id_In And ���㷽ʽ Is Not Null And Nvl(��Ԥ��, 0) <> 0 And Nvl(У�Ա�־, 0) = 0;

  --���α����ڲ����շ�ʱʹ�ù��ĳ�Ԥ�����¼
  Cursor c_Deposit(V����id ����Ԥ����¼.����id%Type) Is
    Select ID, ��Ԥ�� As ���, Ԥ�����
    From ����Ԥ����¼
    Where ��¼���� In (1, 11) And ��¼״̬ In (1, 3) And ����id = V����id And Nvl(��Ԥ��, 0) <> 0
    Order By ID Desc;

  n_����id   ������Ϣ.����id%Type;
  n_����id   ������ü�¼.����id%Type;
  n_������� ����Ԥ����¼.�������%Type;
  n_��ӡid   Ʊ�ݴ�ӡ����.Id%Type;

  n_Ԥ����� ����Ԥ����¼.��Ԥ��%Type;
  n_����ֵ   ����Ԥ����¼.��Ԥ��%Type;
  --�����˷Ѽ������
  n_ʣ������ Number;
  n_ʣ��Ӧ�� Number;
  n_ʣ��ʵ�� Number;
  n_ʣ��ͳ�� Number;
  n_�������� Number;
  n_׼������ Number;
  n_�˷Ѵ��� Number;

  n_Ӧ�ս�� Number;
  n_ʵ�ս�� Number;
  n_ͳ���� Number;
  n_�ܽ��   Number;
  n_����״̬ ������ü�¼.����״̬%Type;
  n_�����˷� Number; --�Ƿ��һ���˷���ȫ���˷�,��ÿ���˷ѹ������жϵõ���
  n_��id     ����ɿ����.Id%Type;

  v_�˷ѽ��� ���㷽ʽ.����%Type;

  l_����id   t_Numlist := t_Numlist();
  l_ҩƷ�շ� t_Numlist := t_Numlist();
  l_ʹ��id   t_Numlist := t_Numlist();
  n_Dec      Number;
  d_Date     Date;
  n_Count    Number;
  n_ԭ����id Number;

  Err_Item Exception;
  v_Err_Msg      Varchar2(255);
  n_����ģʽ     Number(3);
  v_Para         Varchar2(1000);
  n_ҽ��ִ�мƼ� Number;

Begin
  n_��id := Zl_Get��id(����Ա����_In);

  --�Ƿ��Ѿ�ȫ����ȫִ��(ֻ�Ǹõ������ŵ��ݵļ��)
  Select Nvl(Count(*), 0)
  Into n_Count
  From ������ü�¼
  Where NO = No_In And ��¼���� = 1 And ��¼״̬ In (1, 3) And Nvl(ִ��״̬, 0) <> 1;
  If n_Count = 0 Then
    v_Err_Msg := '�õ����е���Ŀ�Ѿ�ȫ����ȫִ�У�';
    Raise Err_Item;
  End If;

  --δ��ȫִ�е���Ŀ�Ƿ���ʣ������(ֻ�����ŵ��ݵļ��)
  --ִ��״̬��ԭʼ��¼���ж�
  Select Nvl(Count(*), 0)
  Into n_Count
  From (Select ���, Sum(����) As ʣ������
         From (Select ��¼״̬, Nvl(�۸񸸺�, ���) As ���, Avg(Nvl(����, 1) * ����) As ����
                From ������ü�¼
                Where NO = No_In And ��¼���� = 1 And Nvl(���ӱ�־, 0) <> 9 And
                      Nvl(�۸񸸺�, ���) In
                      (Select Nvl(�۸񸸺�, ���)
                       From ������ü�¼
                       Where NO = No_In And ��¼���� = 1 And ��¼״̬ In (1, 3) And Nvl(ִ��״̬, 0) <> 1)
                Group By ��¼״̬, Nvl(�۸񸸺�, ���))
         Group By ���
         Having Sum(����) <> 0);
  If n_Count = 0 Then
    v_Err_Msg := '�õ�����δ��ȫִ�в�����Ŀʣ������Ϊ��,û�п����˷ѵķ��ã�';
    Raise Err_Item;
  End If;
  --ȷ���Ƿ���ҽ��ִ�мƼ��д�������,�����������,�����ҽ��ִ�мƼ۽����˷�,���򰴾ɷ�ʽ���д���
  Select Count(1)
  Into n_ҽ��ִ�мƼ�
  From ������ü�¼ A, ҽ��ִ�мƼ� B
  Where a.ҽ����� = b.ҽ��id And a.��¼���� = 1 And a.No = No_In And a.��¼״̬ In (1, 3) And Rownum = 1;

  ---------------------------------------------------------------------------------
  --���ñ���
  Select Sysdate Into d_Date From Dual;
  Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
  n_������� := Null;

  --���С��λ��
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_Dec From Dual;

  --��ȡ���㷽ʽ����
  Begin
    Select ���� Into v_�˷ѽ��� From ���㷽ʽ Where ���� = 1;
  Exception
    When Others Then
      v_�˷ѽ��� := '�ֽ�';
  End;

  ---------------------------------------------------------------------------------
  --ѭ������ÿ�з���(������Ŀ��)
  n_�ܽ��   := 0;
  n_�����˷� := 1;
  For r_Bill In c_Bill Loop
    If Nvl(r_Bill.ִ��״̬, 0) <> 1 Then
      --��ʣ������,ʣ��Ӧ��,ʣ��ʵ��
      Select Sum(Nvl(����, 1) * ����), Sum(Ӧ�ս��), Sum(ʵ�ս��), Sum(ͳ����)
      Into n_ʣ������, n_ʣ��Ӧ��, n_ʣ��ʵ��, n_ʣ��ͳ��
      From ������ü�¼
      Where NO = No_In And ��¼���� = 1 And ��� = r_Bill.���;
    
      If n_ʣ������ = 0 Then
        --�����δ�޶��к�,ԭʼ�����еĸñ��Ѿ�ȫ���˷�(ִ��״̬=0��һ�ֿ���)
        n_�����˷� := 0;
      Else
        --׼������(��ҩƷ��ĿΪʣ������,ԭʼ����)
        If Instr(',4,5,6,7,', r_Bill.�շ����) = 0 Or (r_Bill.�շ���� = '4' And Nvl(r_Bill.��������, 0) = 0) Then
          --@@@
          --��ҩƷ����(�Ծ���ҽ��ִ��Ϊ׼���м��)
          --: 1.����ҽ�����͵�,����ҽ��ִ��Ϊ׼(�����ܰ���:���;����;����;������Ѫ)
          --: 2.������ҽ����,����ʣ������Ϊ׼
          n_Count := 0;
          If Instr(',C,D,F,G,K,', ',' || r_Bill.������� || ',') = 0 And r_Bill.������� Is Not Null Then
            If n_ҽ��ִ�мƼ� = 1 Then
              Select Decode(Sign(Sum(����)), -1, 0, Sum(����)), Count(*)
              Into n_׼������, n_Count
              From (Select Max(Decode(a.��¼״̬, 2, 0, a.Id)) As ID, Max(a.ҽ�����) As ҽ��id, Max(a.�շ�ϸĿid) As �շ�ϸĿid,
                            Sum(Nvl(a.����, 1) * Nvl(a.����, 1)) As ����,
                            Sum(Decode(a.��¼״̬, 2, 0, Nvl(a.����, 1) * Nvl(a.����, 1))) As ԭʼ����
                     From ������ü�¼ A, ����ҽ����¼ M
                     Where a.ҽ����� = m.Id And Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And
                           Instr('5,6,7', a.�շ����) = 0 And a.No = No_In And a.��� = r_Bill.��� And a.��¼���� = 1 And
                           a.��¼״̬ In (1, 2, 3) And a.�۸񸸺� Is Null
                     Group By a.���
                     Union All
                     Select a.Id, a.ҽ����� As ҽ��id, a.�շ�ϸĿid, -1 * b.���� As ��ִ��, 0 ԭʼ����
                     From ������ü�¼ A, ҽ��ִ�мƼ� B, ����ҽ����¼ M
                     Where a.ҽ����� = b.ҽ��id And a.�շ�ϸĿid = b.�շ�ϸĿid + 0 And a.ҽ����� = m.Id And
                           Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And Instr('5,6,7', a.�շ����) = 0 And
                           (Exists
                            (Select 1
                             From ����ҽ��ִ��
                             Where b.ҽ��id = ҽ��id And b.���ͺ� = ���ͺ� And b.Ҫ��ʱ�� = Ҫ��ʱ�� And Nvl(ִ�н��, 0) = 1) Or Exists
                            (Select 1
                             From ����ҽ������
                             Where b.ҽ��id = ҽ��id And b.���ͺ� = ���ͺ� And Nvl(ִ��״̬, 0) = 1)) And Not Exists
                      (Select 1
                            From ����ҽ������
                            Where a.ҽ����� = ҽ��id And a.No = NO And Mod(a.��¼����, 10) = ��¼����) And a.No = No_In And
                           a.��� = r_Bill.��� And a.��¼���� = 1 And a.��¼״̬ In (1, 3) ��and a.�۸񸸺� Is Null) Q1
              Where Not Exists (Select 1 From ҩƷ�շ���¼ Where ����id = Q1.Id) Having Max(ID) <> 0;
            Else
              Select Nvl(Sum(����), 0), Count(*)
              Into n_׼������, n_Count
              From (Select a.ҽ��id, a.�շ�ϸĿid, Nvl(a.����, 1) * Nvl(b.��������, 1) As ����
                     From ����ҽ���Ƽ� A, ����ҽ������ B, ������ü�¼ J, ����ҽ����¼ M
                     Where a.ҽ��id = b.ҽ��id And a.ҽ��id = m.Id And Nvl(b.ִ��״̬, 0) <> 1 And a.ҽ��id = j.ҽ����� And
                           a.�շ�ϸĿid = j.�շ�ϸĿid And j.No = No_In And j.��¼���� = 1 And j.��� = r_Bill.��� And j.��¼״̬ In (1, 3) And
                           j.�۸񸸺� Is Null And Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And Exists
                      (Select 1
                            From ����ҽ���Ƽ� A
                            Where a.ҽ��id = j.ҽ����� And a.�շ�ϸĿid = j.�շ�ϸĿid And Nvl(a.�շѷ�ʽ, 0) = 0) And Not Exists
                      (Select 1 From ҩƷ�շ���¼ Where ����id = j.Id)
                     Union All
                     Select a.ҽ��id, a.�շ�ϸĿid, -1 * Nvl(a.����, 1) * Nvl(c.��������, 1) As ����
                     From ����ҽ���Ƽ� A, ����ҽ������ B, ����ҽ��ִ�� C, ������ü�¼ J, ����ҽ����¼ M
                     Where a.ҽ��id = b.ҽ��id And b.ҽ��id = c.ҽ��id And b.���ͺ� = c.���ͺ� And a.ҽ��id = m.Id And Nvl(c.ִ�н��, 1) = 1 And
                           Nvl(b.ִ��״̬, 0) <> 1 And a.ҽ��id = j.ҽ����� And a.�շ�ϸĿid = j.�շ�ϸĿid And j.No = No_In And
                           j.��¼���� = 1 And Nvl(a.�շѷ�ʽ, 0) = 0 And j.��� = r_Bill.��� And j.��¼״̬ In (1, 3) And j.�۸񸸺� Is Null And
                           Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And Not Exists
                      (Select 1 From ҩƷ�շ���¼ Where ����id = j.Id) And Not Exists
                      (Select 1 From �������� Where ����id = j.�շ�ϸĿid And Nvl(��������, 0) = 1)
                     Union All
                     Select a.ҽ����� As ҽ��id, a.�շ�ϸĿid, Nvl(a.����, 1) * a.���� As ����
                     From ������ü�¼ A, ����ҽ����¼ M
                     Where a.ҽ����� = m.Id And Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And a.No = No_In And
                           a.��¼���� = 1 And a.��� = r_Bill.��� And a.��¼״̬ = 2 And a.�۸񸸺� Is Null And Not Exists
                      (Select 1 From ҩƷ�շ���¼ Where NO = No_In And ���� In (8, 24) And ҩƷid = a.�շ�ϸĿid));
            End If;
          End If;
          If Nvl(n_Count, 0) <> 0 And n_׼������ = 0 Then
            v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з�����ִ��,�������˷ѣ�';
            Raise Err_Item;
          End If;
        
          If Nvl(n_Count, 0) = 0 Then
            n_׼������ := n_ʣ������;
          End If;
        
        Else
          Select Nvl(Sum(Nvl(����, 1) * ʵ������), 0), Count(*)
          Into n_׼������, n_Count
          From ҩƷ�շ���¼
          Where NO = No_In And ���� In (8, 24) And Mod(��¼״̬, 3) = 1 --@@@
                And ����� Is Null And ����id = r_Bill.Id;
        
          --��ʣ��������׼�������������������
          --1.���������õ������޶�Ӧ���շ���¼,��ʱʹ��ʣ������
          --2.��������,��ʱ�ѷ�ҩ����
          If n_׼������ = 0 Then
            If r_Bill.�շ���� = '4' Then
              If n_Count > 0 Then
                v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����ѷ���,�����Ϻ����˷ѣ�';
                Raise Err_Item;
              Else
                n_׼������ := n_ʣ������;
              End If;
            Else
              v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����ѷ�ҩ,����ҩ�����˷ѣ�';
              Raise Err_Item;
            End If;
          End If;
        End If;
      
        --�Ƿ񲿷��˷�
        If r_Bill.ִ��״̬ = 2 Or n_׼������ <> Nvl(r_Bill.����, 1) * r_Bill.���� Then
          n_�����˷� := 0;
        End If;
      
        --����������ü�¼
        n_����״̬ := 0;
        --�ñ���Ŀ�ڼ����˷�
        Select Nvl(Max(Abs(ִ��״̬)), 0) + 1
        Into n_�˷Ѵ���
        From ������ü�¼
        Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 2 And Nvl(ִ��״̬, 0) < 0 And ��� = r_Bill.���;
      
        n_Ӧ�ս�� := Round(n_ʣ��Ӧ�� * (n_׼������ / n_ʣ������), n_Dec);
        n_ʵ�ս�� := Round(n_ʣ��ʵ�� * (n_׼������ / n_ʣ������), n_Dec);
        n_ͳ���� := Round(n_ʣ��ͳ�� * (n_׼������ / n_ʣ������), n_Dec);
        n_�ܽ��   := n_�ܽ�� + n_ʵ�ս��;
      
        --�����˷Ѽ�¼
        Insert Into ������ü�¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�, ���˿���id, �շ����, �շ�ϸĿid,
           ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ������, ִ����, ִ��״̬,
           ����״̬, ִ��ʱ��, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, ����, �ɿ���id)
          Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�,
                 ���˿���id, �շ����, �շ�ϸĿid, ���㵥λ, Decode(Sign(n_׼������ - Nvl(����, 1) * ����), 0, ����, 1), ��ҩ����,
                 Decode(Sign(n_׼������ - Nvl(����, 1) * ����), 0, -1 * ����, -1 * n_׼������), �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����,
                 -1 * n_Ӧ�ս��, -1 * n_ʵ�ս��, ��������id, ������, ִ�в���id, ������, ִ����, -1 * n_�˷Ѵ���, n_����״̬, ִ��ʱ��, ����Ա���_In, ����Ա����_In,
                 ����ʱ��, d_Date, n_����id, -1 * n_ʵ�ս��, ������Ŀ��, ���մ���id, -1 * n_ͳ����, ժҪ, Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���,
                 ��������, ����, n_��id
          From ������ü�¼
          Where ID = r_Bill.Id;
      
        --���ԭ���ü�¼
        --ִ��״̬:ȫ������(׼����=ʣ����)���Ϊ0,������Ϊ1,�쳣�շѵ�,���Ǳ���9
        Update ������ü�¼
        Set ��¼״̬ = 3, ִ��״̬ = Decode(Nvl(ִ��״̬, 0), 9, 9, Decode(Sign(n_׼������ - n_ʣ������), 0, 0, 1))
        Where ID = r_Bill.Id;
      End If;
    Else
      --���:û�޶��к�,ԭʼ�����а����Ѿ���ȫִ�е�
      n_�����˷� := 0;
    End If;
  End Loop;

  ---------------------------------------------------------------------------------
  --������Ԥ����¼
  --�Զ���������,Ĭ�ϱ���һλ
  n_�ܽ�� := Round(n_�ܽ��, 1);
  --ԭ���ݵĽ���ID
  Select ����id, ����id
  Into n_ԭ����id, n_����id
  From ������ü�¼
  Where NO = No_In And ��¼���� = 1 And ��¼״̬ In (1, 3) And Rownum = 1;

  If n_�����˷� = 1 Then
    --���ݵ�һ���˷���ȫ������
    --��Ԥ�����ּ�¼
    Insert Into ����Ԥ����¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��, ����id,
       �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, У�Ա�־, �������, ��������)
      Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, d_Date,
             ����Ա����_In, ����Ա���_In, -1 * ��Ԥ��, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 0, n_�������, 3
      From ����Ԥ����¼
      Where ��¼���� In (1, 11) And ����id = n_ԭ����id And Nvl(��Ԥ��, 0) <> 0;
    --������Ԥ�����
    For v_Ԥ�� In (Select Ԥ�����, Nvl(Sum(Nvl(��Ԥ��, 0)), 0) As Ԥ�����
                 From ����Ԥ����¼
                 Where ��¼���� In (1, 11) And ����id = n_ԭ����id
                 Group By Ԥ�����
                 Having Sum(Nvl(��Ԥ��, 0)) <> 0) Loop
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(v_Ԥ��.Ԥ�����, 0)
      Where ����id = n_����id And ���� = 1 And ���� = Nvl(v_Ԥ��.Ԥ�����, 2)
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, ����, Ԥ�����, ����)
        Values
          (n_����id, Nvl(v_Ԥ��.Ԥ�����, 2), Nvl(v_Ԥ��.Ԥ�����, 0), 1);
        n_����ֵ := n_Ԥ�����;
      End If;
      If n_����ֵ = 0 Then
        Delete From ������� Where ����id = n_����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
      End If;
    End Loop;
  
    --ԭ���˻�(��Ԥ����ǰ���Ѵ���)
    Insert Into ����Ԥ����¼
      (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����,
       �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, У�Ա�־, �������, ��������)
      Select ����Ԥ����¼_Id.Nextval, ��¼����, NO, 2, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, d_Date, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���_In, ����Ա����_In,
             -1 * ��Ԥ��, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 0, n_�������, 3
      From ����Ԥ����¼ A, (Select ���� From ���㷽ʽ Where ���� In (3, 4)) J,
           (Select m.Id As Ԥ��id From ����Ԥ����¼ M Where m.����id = n_ԭ����id And m.��¼���� = 3 And m.��¼״̬ = 1) Q
      Where a.��¼���� = 3 And a.��¼״̬ = 1 And a.����id = n_ԭ����id And a.Id = q.Ԥ��id(+) And a.���㷽ʽ = j.����(+);
  Else
    --�����˷�ֱ����Ϊָ�����㷽ʽ
    Insert Into ����Ԥ����¼
      (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
       ���㿨���, ����, ������ˮ��, ����˵��, ������λ, У�Ա�־, �������, ��������)
      Select ����Ԥ����¼_Id.Nextval, 3, No_In, 2, ����id, ��ҳid, '�����˷ѽ���', v_�˷ѽ���, d_Date, Null, Null, Null, ����Ա���_In, ����Ա����_In,
             -1 * n_�ܽ��, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 0, n_�������, 3
      
      From ����Ԥ����¼
      Where ��¼���� = 3 And ��¼״̬ In (1, 3) And ����id = n_ԭ����id And Rownum = 1;
  
    --����շ�ʱֻʹ����Ԥ����,��Ҫ��Ԥ��,���ҿ����ж�ʳ�Ԥ��
    If Sql%RowCount = 0 Then
      n_Ԥ����� := n_�ܽ��;
    
      For r_Deposit In c_Deposit(n_ԭ����id) Loop
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
           ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, У�Ա�־, �������, ��������)
          Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�,
                 d_Date, ����Ա����_In, ����Ա���_In, Decode(Sign(r_Deposit.��� - n_Ԥ�����), -1, -1 * r_Deposit.���, -1 * n_Ԥ�����),
                 n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 0, n_�������, 3
          From ����Ԥ����¼
          Where ID = r_Deposit.Id;
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
    
      --���²���Ԥ�����
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) + n_�ܽ��
      Where ����id = n_����id And ���� = 1 And ���� = 1
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ������� (����id, ����, Ԥ�����, ����) Values (n_����id, 1, n_�ܽ��, 1);
        n_����ֵ := n_�ܽ��;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From ������� Where ����id = n_����id And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
      End If;
    End If;
  End If;
  --����ԭ��¼
  Update ����Ԥ����¼ Set ��¼״̬ = 3 Where ��¼���� = 3 And ��¼״̬ In (1, 3) And ����id = n_ԭ����id;

  Select Nvl(Sum(Nvl(���ʽ��, 0)), 0) Into n_ʵ�ս�� From ������ü�¼ Where ����id = n_����id;
  Select Nvl(Sum(Nvl(��Ԥ��, 0)), 0) Into n_����ֵ From ����Ԥ����¼ Where ����id = n_����id;

  n_ʵ�ս�� := n_ʵ�ս�� - n_����ֵ;

  If n_ʵ�ս�� <> 0 Then
    --δ�ҵ����²��������
    Zl_���շ����_Insert(No_In, n_����id, n_����id, n_ʵ�ս��, d_Date, ����Ա���_In, ����Ա����_In, 1);
  End If;

  --��Ա�ɿ����(ע����Ԥ����¼�����Ŵ������������ʻ��ȵĽ�����,�����˳�Ԥ����)
  For r_Moneyrow In c_Money(n_����id) Loop
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) + r_Moneyrow.��Ԥ��
    Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = r_Moneyrow.���㷽ʽ
    Returning ��� Into n_����ֵ;
    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ����
        (�տ�Ա, ���㷽ʽ, ����, ���)
      Values
        (����Ա����_In, r_Moneyrow.���㷽ʽ, 1, r_Moneyrow.��Ԥ��);
      n_����ֵ := r_Moneyrow.��Ԥ��;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ��Ա�ɿ����
      Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = r_Moneyrow.���㷽ʽ And Nvl(���, 0) = 0;
    End If;
  End Loop;

  ---------------------------------------------------------------------------------
  --�˷�Ʊ�ݻ���(��ȫ��ʱ�Ż���,�����������ش�����л���)
  --���ñ�־||NO;ִ�п���(����);�վݷ�Ŀ(��ҳ����,����);�շ�ϸĿ(����)
  v_Para     := Nvl(zl_GetSysParameter('Ʊ�ݷ������', 1121), '0||0;0;0,0;0');
  n_����ģʽ := Zl_To_Number(Substr(v_Para, 1, 1));
  If n_����ģʽ <> 0 Then
    --�ջ�Ʊ��
    Select ʹ��id Bulk Collect
    Into l_ʹ��id
    From (Select Distinct b.ʹ��id From Ʊ�ݴ�ӡ��ϸ B Where b.No = No_In And Nvl(b.Ʊ��, 0) = 1);
  
    n_����ģʽ := l_ʹ��id.Count;
    If l_ʹ��id.Count <> 0 Then
      --������ռ�¼
      Forall I In 1 .. l_ʹ��id.Count
        Insert Into Ʊ��ʹ����ϸ
          (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ����, ʹ��ʱ��)
          Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, ����Ա����_In, d_Date
          From Ʊ��ʹ����ϸ A
          Where ID = l_ʹ��id(I) And ���� = 1 And Not Exists
           (Select 1 From Ʊ��ʹ����ϸ Where ���� = a.���� And Ʊ�� = a.Ʊ�� And Nvl(����, 0) <> 1);
    
      Forall I In 1 .. l_ʹ��id.Count
        Update Ʊ�ݴ�ӡ��ϸ Set �Ƿ���� = 1 Where ʹ��id = l_ʹ��id(I) And Nvl(�Ƿ����, 0) = 0;
    
    End If;
  End If;
  If n_����ģʽ = 0 Then
    --��ȡ�������һ�εĴ�ӡID(�����Ƕ��ŵ����շѴ�ӡ)
    Begin
      Select ID
      Into n_��ӡid
      From (Select b.Id
             From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
             Where a.��ӡid = b.Id And a.���� = 1 And b.�������� = 1 And b.No = No_In
             Order By a.ʹ��ʱ�� Desc)
      Where Rownum < 2;
    Exception
      When Others Then
        Null;
    End;
    --������ǰû�д�ӡ,���ջ�
    If n_��ӡid Is Not Null Then
      --a.���ŵ���ѭ������ʱֻ���ջ�һ��
      Select Count(*) Into n_Count From Ʊ��ʹ����ϸ Where Ʊ�� = 1 And ���� = 2 And ��ӡid = n_��ӡid;
      If n_Count = 0 Then
        Insert Into Ʊ��ʹ����ϸ
          (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����)
          Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_Date, ����Ա����_In
          From Ʊ��ʹ����ϸ
          Where ��ӡid = n_��ӡid And Ʊ�� = 1 And ���� = 1;
      Else
        --b.�����˷Ѷ���ջ�ʱ,���һ��ȫ���ջ�Ҫ�ſ����ջص�
        Insert Into Ʊ��ʹ����ϸ
          (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����)
          Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_Date, ����Ա����_In
          From Ʊ��ʹ����ϸ A
          Where ��ӡid = n_��ӡid And Ʊ�� = 1 And ���� = 1 And Not Exists
           (Select 1 From Ʊ��ʹ����ϸ B Where a.���� = b.���� And ��ӡid = n_��ӡid And Ʊ�� = 1 And ���� = 2);
      End If;
    End If;
  End If;

  ---------------------------------------------------------------------------------
  --��������
  For v_���� In (Select ID, ҩƷid, �ⷿid, ����, ����, ����, ʵ������, ����, ��ҩ��ʽ, ���Ч��, Ч��, ��Ʒ����, �ڲ�����, ����id
               From ҩƷ�շ���¼
               Where ���� = 21 --@@@
                     And Mod(��¼״̬, 3) = 1 And ����� Is Null And
                     ����id In (Select ID
                              From ������ü�¼
                              Where NO = No_In And ��¼���� = 1 And ��¼״̬ In (1, 3) And �շ���� = '4')
               Order By ҩƷid) Loop
    --����ҩƷ���
    If v_����.�ⷿid Is Not Null Then
      Update ҩƷ���
      Set �������� = Nvl(��������, 0) + Decode(v_����.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(v_����.����, 1) * Nvl(v_����.ʵ������, 0)
      Where �ⷿid = v_����.�ⷿid And ҩƷid = v_����.ҩƷid And Nvl(����, 0) = Nvl(v_����.����, 0) And ���� = 1;
      If Sql%RowCount = 0 Then
        Insert Into ҩƷ���
          (�ⷿid, ҩƷid, ����, ����, Ч��, ��������, �ϴ�����, �ϴβ���, ���Ч��, ��Ʒ����, �ڲ�����) --@@@
        Values
          (v_����.�ⷿid, v_����.ҩƷid, 1, v_����.����, v_����.Ч��,
           Decode(v_����.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(v_����.����, 1) * Nvl(v_����.ʵ������, 0), v_����.����, v_����.����, v_����.���Ч��,
           v_����.��Ʒ����, v_����.�ڲ�����);
      End If;
    End If;
  
    l_����id.Extend;
    l_����id(l_����id.Count) := v_����.����id;
    l_ҩƷ�շ�.Extend;
    l_ҩƷ�շ�(l_ҩƷ�շ�.Count) := v_����.Id;
  End Loop;

  ---------------------------------------------------------------------------------
  --ҩƷ�������
  For r_Stock In c_Stock Loop
    --����ҩƷ���
    If r_Stock.�ⷿid Is Not Null Then
    
      Select Decode(Count(Column_Value), Null, 0, 0, 0, 1)
      Into n_��������
      From Table(l_����id)
      Where Column_Value = r_Stock.����id;
      If Nvl(n_��������, 0) = 0 Then
      
        Update ҩƷ���
        Set �������� = Nvl(��������, 0) + Decode(r_Stock.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(r_Stock.����, 1) * Nvl(r_Stock.ʵ������, 0)
        Where �ⷿid = r_Stock.�ⷿid And ҩƷid = r_Stock.ҩƷid And Nvl(����, 0) = Nvl(r_Stock.����, 0) And ���� = 1;
        If Sql%RowCount = 0 Then
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, Ч��, ��������, �ϴ�����, �ϴβ���, ���Ч��, ��Ʒ����, �ڲ�����) --@@@
          Values
            (r_Stock.�ⷿid, r_Stock.ҩƷid, 1, r_Stock.����, r_Stock.Ч��,
             Decode(r_Stock.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(r_Stock.����, 1) * Nvl(r_Stock.ʵ������, 0), r_Stock.����, r_Stock.����,
             r_Stock.���Ч��, r_Stock.��Ʒ����, r_Stock.�ڲ�����);
        End If;
      End If;
    End If;
    l_ҩƷ�շ�.Extend;
    l_ҩƷ�շ�(l_ҩƷ�շ�.Count) := r_Stock.Id;
  End Loop;

  --ɾ��ҩƷ�շ���¼
  Forall I In 1 .. l_ҩƷ�շ�.Count
    Delete From ҩƷ�շ���¼ Where ID = l_ҩƷ�շ�(I) And ����� Is Null;

  --δ��ҩƷ��¼
  For r_Spare In c_Spare Loop
    Select Nvl(Count(*), 0)
    Into n_Count
    From ҩƷ�շ���¼
    Where NO = No_In And ���� = r_Spare.���� --@@@
          And Mod(��¼״̬, 3) = 1 And ����� Is Null And Nvl(�ⷿid, 0) = Nvl(r_Spare.�ⷿid, 0);
  
    If n_Count = 0 Then
      Delete From δ��ҩƷ��¼
      Where ���� = r_Spare.���� --@@@
            And NO = No_In And Nvl(�ⷿid, 0) = Nvl(r_Spare.�ⷿid, 0);
    End If;
  End Loop;
  --ҽ������
  --����_In    Integer:=0, --0:����;1-סԺ
  --����_In    Integer:=1, --1-�շѵ�;2-���ʵ�
  --����_In    Integer:=0, --0:ɾ�����۵�;1-�շѻ����;2-�˷ѻ�����
  --No_In      ������ü�¼.No%Type,
  --ҽ��ids_In Varchar2 := Null
  Zl_ҽ������_�Ʒ�״̬_Update(0, 1, 2, No_In);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_������շ�_Delete;
/

--93623:���˺�,2016-02-25,����ҽ�����͵ļƷ�״̬
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
      Zl_���ﻮ�ۼ�¼_Delete(r_Price.No, r_Price.���);
      Commit;
    End If;
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���ﻮ�ۼ�¼_Clear;
/

--93623:���˺�,2016-02-25,����ҽ�����͵ļƷ�״̬
Create Or Replace Procedure Zl_���˻����շ�_Insert
(
  No_In         ������ü�¼.No%Type,
  ����id_In     ������ü�¼.����id%Type,
  ������Դ_In   Number,
  ���ʽ_In   ������ü�¼.���ʽ%Type,
  ����_In       ������ü�¼.����%Type,
  �Ա�_In       ������ü�¼.�Ա�%Type,
  ����_In       ������ü�¼.����%Type,
  ���˿���id_In ������ü�¼.���˿���id%Type,
  ��������id_In ������ü�¼.��������id%Type,
  ������_In     ������ü�¼.������%Type,
  ����id_In     ������ü�¼.����id%Type,
  ����ʱ��_In   ������ü�¼.����ʱ��%Type,
  ����Ա���_In ������ü�¼.����Ա���%Type,
  ����Ա����_In ������ü�¼.����Ա����%Type,
  ��ҩ����_In   Varchar2 := Null,
  �Ƿ���_In   ������ü�¼.�Ƿ���%Type := 0,
  �Ǽ�ʱ��_In   ������ü�¼.�Ǽ�ʱ��%Type := Null
) As
  --���ܣ������շ�ʱ��ȡ���۵�����
  --������
  --      ��ҩ����_In:ִ�в���ID1|��ҩ����1;...;ִ�в���IDn|��ҩ����n

  --        ������Դ_IN:1-����;2-סԺ
  --˵����
  --        1.��ȡ���۷���ʱ,�ż��������ػ���,�ڻ���ʱ������;��ҩƷ��ػ���(��������)����ʱ�Ѿ����㡣
  --        2.��ȡ���۷���ʱ,Ŀǰ���漰������δ������չ�����,�ɻ���ʱֱ�Ӵ���
  --���α�Ϊ����ԭ��������
  Cursor c_Price Is
    Select ID
    From ������ü�¼
    Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 0 And ����Ա���� Is Null
    Order By ���;

  n_Array_Size Number := 200;
  t_����id     t_Numlist;
  v_��������   ���ű�.����%Type;

  v_��ʶ��   ������ü�¼.��ʶ��%Type;
  v_���ʽ ҽ�Ƹ��ʽ.����%Type;

  --��ʱ����
  n_Count      Number;
  n_�²���ģʽ Number;
  v_����no     ҩƷ�շ���¼.No%Type;
  v_Date       Date;
  v_Err_Msg    Varchar2(255);
  Err_Item Exception;
  n_��id ����ɿ����.Id%Type;

Begin
  n_��id := Zl_Get��id(����Ա����_In);

  Select Count(ID)
  Into n_Count
  From ������ü�¼
  Where ��¼���� = 1 And ��¼״̬ = 0 And NO = No_In And ����Ա���� Is Null;
  If n_Count = 0 Then
    v_Err_Msg := '���ܶ�ȡ���۵�����,�õ��ݿ����Ѿ�ɾ�����Ѿ��շѣ�';
    Raise Err_Item;
  End If;
  v_Date := �Ǽ�ʱ��_In;
  If v_Date Is Null Then
    Select Sysdate Into v_Date From Dual;
  End If;
  If Nvl(����id_In, 0) <> 0 Then
    Select Decode(��ǰ����id, Null, �����, סԺ��) Into v_��ʶ�� From ������Ϣ Where ����id = ����id_In;
  End If;

  ------------------------------------------------------------------------------------------------------------------------
  --��������
  Open c_Price;
  Loop
    Fetch c_Price Bulk Collect
      Into t_����id Limit n_Array_Size;
    Exit When t_����id.Count = 0;
  
    --ѭ������������ü�¼
    Forall I In 1 .. t_����id.Count
    --ִ��״̬����ֶβ�����,�ڻ���ʱ����;��Ϊ����δ�շѷ�ҩ,������ִ�еĻ��۵��������շѲ����ġ�
    --Ϊ��֤��Ԥ�������¼��ʱ����ͬ,������д�Ǽ�ʱ��,��ҩƷ���ֲ��䶯��
      Update ������ü�¼
      Set ��¼״̬ = 1, ����id = Decode(����id_In, 0, Null, ����id_In), ��ʶ�� = v_��ʶ��, ���ʽ = ���ʽ_In, ���� = ����_In, ���� = ����_In,
          �Ա� = �Ա�_In,
          --���ܱ���ҽ�����͵�����
          ���˿���id = Nvl(���˿���id_In, ���˿���id), ��������id = Nvl(��������id_In, ��������id), ������ = Nvl(������_In, ������), ���ʽ�� = ʵ�ս��,
          ����id = ����id_In, ����ʱ�� = ����ʱ��_In, �Ǽ�ʱ�� = v_Date, ����Ա��� = ����Ա���_In, ����Ա���� = ����Ա����_In, �Ƿ��� = �Ƿ���_In,
          �ɿ���id = n_��id, ����״̬ = 1, ִ��״̬ = Decode(Nvl(ִ��״̬, 0), -1, Null, Nvl(ִ��״̬, 0))
      Where ID = t_����id(I) And ��¼״̬ = 0;
  
    If Sql%RowCount <> t_����id.Count Then
      v_Err_Msg := '���ڲ�������,�õ��ݿ����Ѿ�ɾ�����Ѿ��շѣ�';
      Raise Err_Item;
    End If;
  End Loop;

  Close c_Price;

  --��ػ��ܱ�Ĵ���
  --ҩƷ���ַǷ�����Ϣ���޸�
  --ҩƷδ����¼(����ѷ�ҩ���޸Ĳ���),���뷢ҩʱ�޿ⷿID
  --���ܴ��ڲ��Ϻ�ҩƷ�ⷿ��ͬ���������޷�ҩ����
  Update δ��ҩƷ��¼
  Set ����id = Decode(����id_In, 0, Null, ����id_In), ���� = ����_In, �Է�����id = ��������id_In, ���շ� = 1, �������� = v_Date
  Where ���� = 24 And NO = No_In And
        Nvl(�ⷿid, 0) In (Select Distinct Nvl(ִ�в���id, 0)
                         From ������ü�¼
                         Where ��¼���� = 1 And ��¼״̬ = 1 And NO = No_In And �շ���� = '4');

  Update δ��ҩƷ��¼
  Set ����id = Decode(����id_In, 0, Null, ����id_In), ���� = ����_In, �Է�����id = ��������id_In, ���շ� = 1, �������� = v_Date
  Where ���� = 8 And NO = No_In And
        Nvl(�ⷿid, 0) In (Select Distinct Nvl(ִ�в���id, 0)
                         From ������ü�¼
                         Where ��¼���� = 1 And ��¼״̬ = 1 And NO = No_In And �շ���� In ('5', '6', '7'));

  --ҩƷ�շ���¼(�����Ѿ���ҩ��ȡ����ҩ,���м�¼����)
  Update ҩƷ�շ���¼
  Set �Է�����id = ��������id_In, �������� = Decode(Sign(Nvl(�������, v_Date) - v_Date), -1, ��������, v_Date)
  Where ���� = 24 And NO = No_In And
        ����id + 0 In (Select ID From ������ü�¼ Where ��¼���� = 1 And ��¼״̬ = 1 And NO = No_In And �շ���� = '4');

  -------------------------------------------------------------------------------------------
  --����������
  n_Count := Null;
  Begin
    Select Count(*), Max(a.No)
    Into n_Count, v_����no
    From ҩƷ�շ���¼ A, ������ü�¼ B
    Where a.����id = b.Id And b.�շ���� = '4' And b.��¼���� = 1 And b.��¼״̬ = 1 And b.No = No_In And
          Instr(',8,9,10,21,24,25,26,', ',' || a.���� || ',') > 0 And Rownum <= 1;
  Exception
    When Others Then
      Null;
  End;
  If Nvl(n_Count, 0) > 0 Then
    If Nvl(���˿���id_In, 0) <> 0 Then
      Select ���� Into v_�������� From ���ű� Where ID = ���˿���id_In;
    End If;
    v_Err_Msg := LPad(' ', 4);
    v_Err_Msg := Substr('��������:' || ����_In || v_Err_Msg || '�Ա�:' || �Ա�_In || v_Err_Msg || '����' || ����_In || v_Err_Msg ||
                        '�����:' || Nvl(v_��ʶ��, '') || v_Err_Msg || '���˿���:' || v_��������, 1, 100);
  
    Update ҩƷ�շ���¼
    Set �Է�����id = ��������id_In, �������� = Decode(Sign(Nvl(�������, v_Date) - v_Date), -1, ��������, v_Date), ժҪ = v_Err_Msg
    Where ���� = 21 And NO = v_����no And
          ����id + 0 In (Select ID From ������ü�¼ Where ��¼���� = 1 And ��¼״̬ = 1 And NO = No_In And �շ���� = '4');
  End If;

  Update ҩƷ�շ���¼
  Set �Է�����id = ��������id_In, �������� = Decode(Sign(Nvl(�������, v_Date) - v_Date), -1, ��������, v_Date)
  Where ���� = 8 And NO = No_In And
        ����id + 0 In
        (Select ID From ������ü�¼ Where ��¼���� = 1 And ��¼״̬ = 1 And NO = No_In And �շ���� In ('5', '6', '7'));

  If Not ��ҩ����_In Is Null Then
    --���·�ҩ����
    For v_���� In (Select To_Number(C1) As C1, C2 From Table(f_Str2list2(��ҩ����_In, ';', '|'))) Loop
    
      Update ������ü�¼
      Set ��ҩ���� = Nvl(v_����.C2, ��ҩ����)
      Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 1 And ִ�в���id = Nvl(v_����.C1, ִ�в���id) And �շ���� In ('5', '6', '7');
    
      Update ҩƷ�շ���¼
      Set ��ҩ���� = Nvl(v_����.C2, ��ҩ����)
      Where ���� = 8 And NO = No_In And �ⷿid = Nvl(v_����.C1, �ⷿid) And
            ����id + 0 In (Select ID
                         From ������ü�¼
                         Where ��¼���� = 1 And ��¼״̬ = 1 And NO = No_In And �շ���� In ('5', '6', '7'));
    
      Update δ��ҩƷ��¼
      Set ��ҩ���� = Nvl(v_����.C2, ��ҩ����)
      Where ���� = 8 And NO = No_In And �ⷿid = Nvl(v_����.C1, �ⷿid) And
            Nvl(�ⷿid, 0) In (Select Distinct Nvl(ִ�в���id, 0)
                             From ������ü�¼
                             Where ��¼���� = 1 And ��¼״̬ = 1 And NO = No_In And �շ���� In ('5', '6', '7'));
    
    End Loop;
  End If;

  --���²��ݲ�����Ϣ
  If ����id_In Is Not Null Then
    If ���ʽ_In Is Not Null And ������Դ_In = 1 Then
      Select ���� Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = ���ʽ_In;
    End If;
    --ͨ�����۵��շ�ʱ������ķѱ�,��Ϊ���ò������
    Update ������Ϣ
    Set �Ա� = Decode(����, '�²���', Nvl(�Ա�_In, �Ա�), �Ա�), ���� = Decode(����, '�²���', Nvl(����_In, ����), ����),
        ���� = Decode(����, '�²���', ����_In, ����), ҽ�Ƹ��ʽ = Nvl(v_���ʽ, ҽ�Ƹ��ʽ)
    Where ����id = ����id_In;
    Select Zl_To_Number(Nvl(zl_GetSysParameter('�Զ���������', '1111'), '0')) Into n_�²���ģʽ From Dual;
    If n_�²���ģʽ = 1 Then
    
      Update ���˹Һż�¼
      Set ���� = ����_In, �Ա� = �Ա�_In, ���� = ����_In
      Where ����id = ����id_In And ���� = '�²���';
    
      Update ������ü�¼
      Set ���� = ����_In, �Ա� = �Ա�_In, ���� = ����_In, ���ʽ = ���ʽ_In
      Where ����id = ����id_In And ���� = '�²���';
    End If;
  End If;

  --ҽ������
  --����_In    Integer:=0, --0:����;1-סԺ
  --����_In    Integer:=1, --1-�շѵ�;2-���ʵ�
  --����_In    Integer:=0, --0:ɾ�����۵�;1-�շѻ����;2-�˷ѻ�����
  --No_In      ������ü�¼.No%Type,
  --ҽ��ids_In Varchar2 := Null
  Zl_ҽ������_�Ʒ�״̬_Update(0, 1, 1, No_In);

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˻����շ�_Insert;
/

--93623:���˺�,2016-02-25,����ҽ�����͵ļƷ�״̬
Create Or Replace Procedure Zl_�����շѼ�¼_Insert
(
  No_In         ������ü�¼.No%Type,
  ����id_In     ������ü�¼.����id%Type,
  ������Դ_In   Number,
  ���ʽ_In   ������ü�¼.���ʽ%Type,
  ����_In       ������ü�¼.����%Type,
  �Ա�_In       ������ü�¼.�Ա�%Type,
  ����_In       ������ü�¼.����%Type,
  ���˿���id_In ������ü�¼.���˿���id%Type,
  ��������id_In ������ü�¼.��������id%Type,
  ������_In     ������ü�¼.������%Type,
  �շѽ���_In   Varchar2,
  ��Ԥ����_In   ����Ԥ����¼.��Ԥ��%Type,
  ���ս���_In   Varchar2,
  ����id_In     ������ü�¼.����id%Type,
  ����ʱ��_In   ������ü�¼.����ʱ��%Type,
  ����Ա���_In ������ü�¼.����Ա���%Type,
  ����Ա����_In ������ü�¼.����Ա����%Type,
  ��ҩ����_In   Varchar2,
  �Ƿ���_In   ������ü�¼.�Ƿ���%Type := 0,
  �ɿ�_In       ����Ԥ����¼.�ɿ�%Type := Null,
  �Ҳ�_In       ����Ԥ����¼.�Ҳ�%Type := Null,
  ����������_In Varchar2 := Null,
  �Ǽ�ʱ��_In   ������ü�¼.�Ǽ�ʱ��%Type := Null,
  �������_In   ����Ԥ����¼.�������%Type := Null,
  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
  ���շ�_In   Number := 0
) As
  --���ܣ������շ�ʱ��ȡ���۵�����
  --������
  --      ��ҩ����_In:ִ�в���ID1|��ҩ����1;...;ִ�в���IDn|��ҩ����n
  --        ������Դ_IN:1-����;2-סԺ
  --        �շѽ���_IN:��ʽ="���㷽ʽ|������|�������|����ժҪ||.....",ע���޽�������ժҪʱҪ�ÿո����
  --        ���ս���_IN:��ʽ="���㷽ʽ|������||....."
  --        ����������_In:��ʽ=�����Id|�Ƿ����ѿ�|������|����|��ע||...
  --        ������ˮ��_In�ͽ���˵��_In:�շѽ���_INʱ��Ч.
  --˵����
  --        1.��ȡ���۷���ʱ,�ż��������ػ���,�ڻ���ʱ������;��ҩƷ��ػ���(��������)����ʱ�Ѿ����㡣
  --        2.��ȡ���۷���ʱ,Ŀǰ���漰������δ������չ�����,�ɻ���ʱֱ�Ӵ���
  --���α�Ϊ����ԭ��������

  --=================================
  --��ע���ù���Ŀǰֻ�м��շ�ʹ�ã�
  --=================================

  Cursor c_Price Is
    Select ID
    From ������ü�¼
    Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 0 And ����Ա���� Is Null
    Order By ���;

  n_Array_Size Number := 200;
  t_����id     t_Numlist;
  v_��������   ���ű�.����%Type;
  --���α������շѳ�Ԥ���Ŀ���Ԥ���б�(��SQL�ο�סԺ����)
  --��ID�������ȳ��ϴ�δ����ġ�
  Cursor c_Deposit(v_����id ������Ϣ.����id%Type) Is
    Select *
    From (Select a.Id, a.��¼״̬, Nvl(a.Ԥ�����, 2) As Ԥ�����, a.No, Nvl(a.���, 0) As ���
           From ����Ԥ����¼ A,
                (Select NO, Sum(Nvl(a.���, 0)) As ���
                  From ����Ԥ����¼ A
                  Where a.����id Is Null And Nvl(a.���, 0) <> 0 And a.����id = v_����id And Nvl(Ԥ�����, 2) = 1
                  Group By NO
                  Having Sum(Nvl(a.���, 0)) <> 0) B
           Where a.����id Is Null And Nvl(a.���, 0) <> 0 And a.No = b.No And a.����id = v_����id And Nvl(a.Ԥ�����, 2) = 1
           Union All
           Select 0 As ID, ��¼״̬, Nvl(Ԥ�����, 2) As Ԥ�����, NO, Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) As ���
           From ����Ԥ����¼
           Where ��¼���� In (1, 11) And ����id Is Not Null And Nvl(Ԥ�����, 2) = 1 And Nvl(���, 0) <> Nvl(��Ԥ��, 0) And
                 ����id = v_����id Having Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0
           Group By ��¼״̬, NO, Nvl(Ԥ�����, 2))
    Order By ID, Ԥ����� Desc, NO;

  --Ԥ���������ر���
  v_Ԥ����� ����Ԥ����¼.��Ԥ��%Type;
  v_�������� Varchar2(3000);
  v_��ǰ���� Varchar2(150);
  v_���㷽ʽ ����Ԥ����¼.���㷽ʽ%Type;
  n_������ ����Ԥ����¼.��Ԥ��%Type;
  v_������� ����Ԥ����¼.�������%Type;
  v_����ժҪ ����Ԥ����¼.ժҪ%Type;

  v_��ʶ��   ������ü�¼.��ʶ��%Type;
  v_���ʽ ҽ�Ƹ��ʽ.����%Type;
  n_����ֵ   �������.Ԥ�����%Type;

  --��ʱ����
  n_Count      Number;
  n_�²���ģʽ Number;
  v_����no     ҩƷ�շ���¼.No%Type;
  v_Date       Date;
  v_Err_Msg    Varchar2(255);
  Err_Item Exception;
  n_��id     ����ɿ����.Id%Type;
  n_�����id ҽ�ƿ����.Id%Type;
  n_���ѿ�   Number;
  v_����     ����Ԥ����¼.����%Type;
  v_������   Varchar2(100);
  n_���ƿ�   Number;
  n_Ԥ��id   ����Ԥ����¼.Id%Type;
  n_���ѿ�id ���ѿ�Ŀ¼.Id%Type;
  n_���     ����Ԥ����¼.���%Type;
Begin
  n_��id := Zl_Get��id(����Ա����_In);

  Select Count(ID)
  Into n_Count
  From ������ü�¼
  Where ��¼���� = 1 And ��¼״̬ = 0 And NO = No_In And ����Ա���� Is Null;
  If n_Count = 0 Then
    v_Err_Msg := '���ܶ�ȡ���۵�����,�õ��ݿ����Ѿ�ɾ�����Ѿ��շѣ�';
    Raise Err_Item;
  End If;
  v_Date := �Ǽ�ʱ��_In;
  If v_Date Is Null Then
    Select Sysdate Into v_Date From Dual;
  End If;
  If Nvl(����id_In, 0) <> 0 Then
    Select Decode(��ǰ����id, Null, �����, סԺ��) Into v_��ʶ�� From ������Ϣ Where ����id = ����id_In;
  End If;

  ------------------------------------------------------------------------------------------------------------------------
  --��������
  Open c_Price;
  Loop
    Fetch c_Price Bulk Collect
      Into t_����id Limit n_Array_Size;
    Exit When t_����id.Count = 0;
  
    --ѭ������������ü�¼
    Forall I In 1 .. t_����id.Count
    --ִ��״̬����ֶβ�����,�ڻ���ʱ����;��Ϊ����δ�շѷ�ҩ,������ִ�еĻ��۵��������շѲ����ġ�
    --Ϊ��֤��Ԥ�������¼��ʱ����ͬ,������д�Ǽ�ʱ��,��ҩƷ���ֲ��䶯��
      Update ������ü�¼
      Set ��¼״̬ = 1, ����id = Decode(����id_In, 0, Null, ����id_In), ��ʶ�� = v_��ʶ��, ���ʽ = ���ʽ_In, ���� = ����_In, ���� = ����_In,
          �Ա� = �Ա�_In,
          --���ܱ���ҽ�����͵�����
          ���˿���id = Nvl(���˿���id_In, ���˿���id), ��������id = Nvl(��������id_In, ��������id), ������ = Nvl(������_In, ������), ���ʽ�� = ʵ�ս��,
          ����id = ����id_In, ����ʱ�� = ����ʱ��_In, �Ǽ�ʱ�� = v_Date, ����Ա��� = ����Ա���_In, ����Ա���� = ����Ա����_In, �Ƿ��� = �Ƿ���_In,
          �ɿ���id = n_��id
      Where ID = t_����id(I) And ��¼״̬ = 0;
  
    If Sql%RowCount <> t_����id.Count Then
      v_Err_Msg := '���ڲ�������,�õ��ݿ����Ѿ�ɾ�����Ѿ��շѣ�';
      Raise Err_Item;
    End If;
  End Loop;
  Close c_Price;
  ------------------------------------------------------------------------------------------------------------------------

  --Ԥ������ؽ���
  --�շѽ���
  If �շѽ���_In Is Not Null Then
    v_�������� := �շѽ���_In || '||';
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
    
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
    
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      v_������� := LTrim(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_����ժҪ := LTrim(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
    
      If Nvl(n_������, 0) <> 0 Then
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ�, �Ҳ�, �ɿ���id, �������, ������ˮ��,
           ����˵��, ��������)
        Values
          (����Ԥ����¼_Id.Nextval, 3, No_In, 1, Decode(����id_In, 0, Null, ����id_In), Null, v_����ժҪ, v_���㷽ʽ, v_�������, v_Date,
           ����Ա���_In, ����Ա����_In, n_������, ����id_In, Decode(v_��������, �շѽ���_In || '||', �ɿ�_In, Null),
           Decode(v_��������, �շѽ���_In || '||', �Ҳ�_In, Null), n_��id, �������_In, ������ˮ��_In, ����˵��_In, 3);
      End If;
    
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;

  If ���ս���_In Is Not Null Then
    --�������ս���
    v_�������� := ���ս���_In || '||';
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
    
      v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
      n_������ := To_Number(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
    
      If Nvl(n_������, 0) <> 0 Then
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, ��������)
        Values
          (����Ԥ����¼_Id.Nextval, 3, No_In, 1, Decode(����id_In, 0, Null, ����id_In), '���ս���', v_���㷽ʽ, v_Date, ����Ա���_In,
           ����Ա����_In, n_������, ����id_In, n_��id, �������_In, 3);
      End If;
    
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;

  If ����������_In Is Not Null Then
    v_�������� := ����������_In || '||';
    While v_�������� Is Not Null Loop
      --�����Id|�Ƿ����ѿ�|������|����|��ע||...
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      n_�����id := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_���ѿ�   := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
    
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      v_����     := LTrim(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
    
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      v_����ժҪ := v_��ǰ����;
    
      If n_���ѿ� = 1 Then
        Select ���㷽ʽ, ����, Nvl(���ƿ�, 0)
        Into v_���㷽ʽ, v_������, n_���ƿ�
        From �����ѽӿ�Ŀ¼
        Where ��� = n_�����id;
        If v_���㷽ʽ Is Null Then
          v_Err_Msg := v_������ || 'δ���ý��㷽ʽ����,�������ѿ��н�������,����ʧ�ܣ�';
          Raise Err_Item;
        End If;
      Else
        Select ���㷽ʽ, ���� Into v_���㷽ʽ, v_������ From ҽ�ƿ���� Where ID = n_�����id;
        If v_���㷽ʽ Is Null Then
          v_Err_Msg := v_������ || 'δ���ý��㷽ʽ����,����ҽ�ƿ������н�������,����ʧ�ܣ�';
          Raise Err_Item;
        End If;
      End If;
      If Nvl(n_������, 0) <> 0 Then
        Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, �����id, ���㿨���, ���㷽ʽ, �������, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ�, �Ҳ�, �ɿ���id,
           �������, ����, ��������)
        Values
          (n_Ԥ��id, 3, No_In, 1, Decode(����id_In, 0, Null, ����id_In), Null, v_����ժҪ, Decode(n_���ѿ�, 1, Null, n_�����id),
           Decode(n_���ѿ�, 0, Null, n_�����id), v_���㷽ʽ, v_�������, v_Date, ����Ա���_In, ����Ա����_In, n_������, ����id_In, Null, Null,
           n_��id, �������_In, v_����, 3);
      
        --���������
        If n_���ѿ� = 1 Then
          n_���ѿ�id := Null;
          If n_���ƿ� = 1 Then
            Select ID
            Into n_���ѿ�id
            From ���ѿ�Ŀ¼
            Where �ӿڱ�� = n_�����id And ���� = v_���� And
                  ��� = (Select Max(���) From ���ѿ�Ŀ¼ Where �ӿڱ�� = n_�����id And ���� = v_����);
          End If;
          Zl_���˿������¼_Insert(n_�����id, n_���ѿ�id, v_���㷽ʽ, n_������, v_����, Null, Null, v_����ժҪ, ����id_In, n_Ԥ��id);
        End If;
      End If;
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;

  --Ԥ������
  If Nvl(��Ԥ����_In, 0) <> 0 Then
    If Nvl(����id_In, 0) = 0 Then
      v_Err_Msg := '����ȷ�����˲���ID,�շ�ʹ��Ԥ�������ʧ�ܣ�';
      Raise Err_Item;
    End If;
  
    v_Ԥ����� := ��Ԥ����_In;
    For r_Deposit In c_Deposit(����id_In) Loop
    
      n_��� := Case
                When r_Deposit.��� - v_Ԥ����� < 0 Then
                 r_Deposit.���
                Else
                 v_Ԥ�����
              End;
      If r_Deposit.Id <> 0 Then
        --��һ�γ�Ԥ��(82592,����һ�α��Ͻ���ID,��Ԥ�����Ϊ0)
        Update ����Ԥ����¼ Set ��Ԥ�� = 0, ����id = ����id_In, �������� = 3 Where ID = r_Deposit.Id;
      End If;
      --���ϴ�ʣ���
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
         ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, У�Ա�־, ��������)
        Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, v_Date,
               ����Ա����_In, ����Ա���_In, n_���, ����id_In, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������_In, У�Ա�־, 3
        From ����Ԥ����¼
        Where NO = r_Deposit.No And ��¼״̬ = r_Deposit.��¼״̬ And ��¼���� In (1, 11) And Rownum = 1;
    
      --���²���Ԥ�����
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) - n_���
      Where ����id = ����id_In And ���� = 1 And ���� = Nvl(r_Deposit.Ԥ�����, 2)
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ������� (����id, Ԥ�����, ����, ����) Values (����id_In, -n_���, Nvl(r_Deposit.Ԥ�����, 2), 1);
        n_����ֵ := -��Ԥ����_In;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From ������� Where ����id = ����id_In And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
      End If;
    
      --����Ƿ��Ѿ�������
      If r_Deposit.��� < v_Ԥ����� Then
        v_Ԥ����� := v_Ԥ����� - r_Deposit.���;
      Else
        v_Ԥ����� := 0;
      End If;
      If v_Ԥ����� = 0 Then
        Exit;
      End If;
    End Loop;
    --������Ƿ��㹻
    If v_Ԥ����� > 0 Then
      v_Err_Msg := '���˵ĵ�ǰԤ�������� ' || LTrim(To_Char(��Ԥ����_In, '9999999990.00')) || ' ��';
      Raise Err_Item;
    End If;
  
  End If;

  --��ػ��ܱ�Ĵ���

  --����"��Ա�ɿ����"
  --�շѽ���
  n_����ֵ := 0;
  If �շѽ���_In Is Not Null Then
    v_�������� := �շѽ���_In || '||';
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
    
      If Nvl(n_������, 0) <> 0 Then
        Update ��Ա�ɿ����
        Set ��� = Nvl(���, 0) + Nvl(n_������, 0)
        Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_���㷽ʽ
        Returning Nvl(���, 0) + n_����ֵ Into n_����ֵ;
        If Sql%RowCount = 0 Then
          Insert Into ��Ա�ɿ����
            (�տ�Ա, ���㷽ʽ, ����, ���)
          Values
            (����Ա����_In, v_���㷽ʽ, 1, Nvl(n_������, 0));
          n_����ֵ := Nvl(n_����ֵ, 0) + Nvl(n_������, 0);
        End If;
      
      End If;
    
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;

  --�������ս���
  If ���ս���_In Is Not Null Then
    v_�������� := ���ս���_In || '||';
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
    
      v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
      n_������ := To_Number(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
    
      If Nvl(n_������, 0) <> 0 Then
        Update ��Ա�ɿ����
        Set ��� = Nvl(���, 0) + Nvl(n_������, 0)
        Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_���㷽ʽ
        Returning Nvl(���, 0) + n_����ֵ Into n_����ֵ;
        If Sql%RowCount = 0 Then
          Insert Into ��Ա�ɿ����
            (�տ�Ա, ���㷽ʽ, ����, ���)
          Values
            (����Ա����_In, v_���㷽ʽ, 1, Nvl(n_������, 0));
          n_����ֵ := Nvl(n_����ֵ, 0) + Nvl(n_������, 0);
        End If;
      End If;
    
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;

  If ����������_In Is Not Null Then
    v_�������� := ����������_In || '||';
    While v_�������� Is Not Null Loop
      --�����Id|�Ƿ����ѿ�|������|����|��ע||...
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      n_�����id := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_���ѿ�   := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
    
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      v_����     := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') + 1);
    
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      v_����ժҪ := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
    
      If n_���ѿ� = 1 Then
        Select ���㷽ʽ, ����, Nvl(���ƿ�, 0)
        Into v_���㷽ʽ, v_������, n_���ƿ�
        From �����ѽӿ�Ŀ¼
        Where ��� = n_�����id;
        If v_���㷽ʽ Is Null Then
          v_Err_Msg := v_������ || 'δ���ý��㷽ʽ����,�������ѿ��н�������,����ʧ�ܣ�';
          Raise Err_Item;
        End If;
      Else
        Select ���㷽ʽ, ���� Into v_���㷽ʽ, v_������ From ҽ�ƿ���� Where ID = n_�����id;
        If v_���㷽ʽ Is Null Then
          v_Err_Msg := v_������ || 'δ���ý��㷽ʽ����,����ҽ�ƿ������н�������,����ʧ�ܣ�';
          Raise Err_Item;
        End If;
      End If;
    
      If Nvl(n_������, 0) <> 0 Then
        Update ��Ա�ɿ����
        Set ��� = Nvl(���, 0) + Nvl(n_������, 0)
        Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_���㷽ʽ
        Returning Nvl(���, 0) + n_����ֵ Into n_����ֵ;
        If Sql%RowCount = 0 Then
          Insert Into ��Ա�ɿ����
            (�տ�Ա, ���㷽ʽ, ����, ���)
          Values
            (����Ա����_In, v_���㷽ʽ, 1, Nvl(n_������, 0));
          n_����ֵ := Nvl(n_����ֵ, 0) + Nvl(n_������, 0);
        End If;
      End If;
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;

  If Nvl(n_����ֵ, 0) = 0 Then
    Delete From ��Ա�ɿ���� Where ���� = 1 And �տ�Ա = ����Ա����_In And Nvl(���, 0) = 0;
  End If;

  --ҩƷ���ַǷ�����Ϣ���޸�
  --ҩƷδ����¼(����ѷ�ҩ���޸Ĳ���),���뷢ҩʱ�޿ⷿID
  --���ܴ��ڲ��Ϻ�ҩƷ�ⷿ��ͬ���������޷�ҩ����
  Update δ��ҩƷ��¼
  Set ����id = Decode(����id_In, 0, Null, ����id_In), ���� = ����_In, �Է�����id = ��������id_In, ���շ� = 1, �������� = v_Date
  Where ���� = 24 And NO = No_In And
        Nvl(�ⷿid, 0) In (Select Distinct Nvl(ִ�в���id, 0)
                         From ������ü�¼
                         Where ��¼���� = 1 And ��¼״̬ = 1 And NO = No_In And �շ���� = '4');

  Update δ��ҩƷ��¼
  Set ����id = Decode(����id_In, 0, Null, ����id_In), ���� = ����_In, �Է�����id = ��������id_In, ���շ� = 1, �������� = v_Date
  Where ���� = 8 And NO = No_In And
        Nvl(�ⷿid, 0) In (Select Distinct Nvl(ִ�в���id, 0)
                         From ������ü�¼
                         Where ��¼���� = 1 And ��¼״̬ = 1 And NO = No_In And �շ���� In ('5', '6', '7'));

  --ҩƷ�շ���¼(�����Ѿ���ҩ��ȡ����ҩ,���м�¼����)
  Update ҩƷ�շ���¼
  Set �Է�����id = ��������id_In, �������� = Decode(Sign(Nvl(�������, v_Date) - v_Date), -1, ��������, v_Date)
  Where ���� = 24 And NO = No_In And
        ����id + 0 In (Select ID From ������ü�¼ Where ��¼���� = 1 And ��¼״̬ = 1 And NO = No_In And �շ���� = '4');

  -------------------------------------------------------------------------------------------
  --����������
  n_Count := Null;
  Begin
    Select Count(*), Max(a.No)
    Into n_Count, v_����no
    From ҩƷ�շ���¼ A, ������ü�¼ B
    Where a.����id = b.Id And b.�շ���� = '4' And b.��¼���� = 1 And b.��¼״̬ = 1 And
          Instr(',8,9,10,21,24,25,26,', ',' || a.���� || ',') > 0 And b.No = No_In And Rownum <= 1;
  Exception
    When Others Then
      Null;
  End;
  If Nvl(n_Count, 0) > 0 Then
    If Nvl(���˿���id_In, 0) <> 0 Then
      Select ���� Into v_�������� From ���ű� Where ID = ���˿���id_In;
    End If;
    v_Err_Msg := LPad(' ', 4);
    v_Err_Msg := Substr('��������:' || ����_In || v_Err_Msg || '�Ա�:' || �Ա�_In || v_Err_Msg || '����' || ����_In || v_Err_Msg ||
                        '�����:' || Nvl(v_��ʶ��, '') || v_Err_Msg || '���˿���:' || v_��������, 1, 100);
  
    Update ҩƷ�շ���¼
    Set �Է�����id = ��������id_In, �������� = Decode(Sign(Nvl(�������, v_Date) - v_Date), -1, ��������, v_Date), ժҪ = v_Err_Msg
    Where ���� = 21 And NO = v_����no And
          ����id + 0 In (Select ID From ������ü�¼ Where ��¼���� = 1 And ��¼״̬ = 1 And NO = No_In And �շ���� = '4');
  End If;

  Update ҩƷ�շ���¼
  Set �Է�����id = ��������id_In, �������� = Decode(Sign(Nvl(�������, v_Date) - v_Date), -1, ��������, v_Date)
  Where ���� = 8 And NO = No_In And
        ����id + 0 In
        (Select ID From ������ü�¼ Where ��¼���� = 1 And ��¼״̬ = 1 And NO = No_In And �շ���� In ('5', '6', '7'));
  If Not ��ҩ����_In Is Null Then
    --���·�ҩ����
    If Nvl(���շ�_In, 0) <> 0 Then
      Update ������ü�¼
      Set ��ҩ���� = ��ҩ����_In
      Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 1 And �շ���� = 'Z';
    Else
      For v_���� In (Select To_Number(C1) As C1, C2 From Table(f_Str2list2(��ҩ����_In, ';', '|'))) Loop
        Update ������ü�¼
        Set ��ҩ���� = Nvl(v_����.C2, ��ҩ����)
        Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 1 And ִ�в���id = Nvl(v_����.C1, ִ�в���id) And �շ���� In ('5', '6', '7');
      
        Update ҩƷ�շ���¼
        Set ��ҩ���� = Nvl(v_����.C2, ��ҩ����)
        Where ���� = 8 And NO = No_In And �ⷿid = Nvl(v_����.C1, �ⷿid) And
              ����id + 0 In (Select ID
                           From ������ü�¼
                           Where ��¼���� = 1 And ��¼״̬ = 1 And NO = No_In And �շ���� In ('5', '6', '7'));
      
        Update δ��ҩƷ��¼
        Set ��ҩ���� = Nvl(v_����.C2, ��ҩ����)
        Where ���� = 8 And NO = No_In And �ⷿid = Nvl(v_����.C1, �ⷿid) And
              Nvl(�ⷿid, 0) In (Select Distinct Nvl(ִ�в���id, 0)
                               From ������ü�¼
                               Where ��¼���� = 1 And ��¼״̬ = 1 And NO = No_In And �շ���� In ('5', '6', '7'));
      End Loop;
    End If;
  End If;

  --���²��ݲ�����Ϣ
  If ����id_In Is Not Null Then
    If ���ʽ_In Is Not Null And ������Դ_In = 1 Then
      Select ���� Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = ���ʽ_In;
    End If;
    --ͨ�����۵��շ�ʱ������ķѱ�,��Ϊ���ò������
    Update ������Ϣ
    Set �Ա� = Nvl(�Ա�_In, �Ա�), ���� = Nvl(����_In, ����), ���� = Decode(����, '�²���', ����_In, ����), ҽ�Ƹ��ʽ = Nvl(v_���ʽ, ҽ�Ƹ��ʽ)
    Where ����id = ����id_In;
    Select Zl_To_Number(Nvl(zl_GetSysParameter('�Զ���������', '1111'), '0')) Into n_�²���ģʽ From Dual;
    If n_�²���ģʽ = 1 Then
    
      Update ���˹Һż�¼
      Set ���� = ����_In, �Ա� = �Ա�_In, ���� = ����_In
      Where ����id = ����id_In And ���� = '�²���';
    
      Update ������ü�¼
      Set ���� = ����_In, �Ա� = �Ա�_In, ���� = ����_In, ���ʽ = ���ʽ_In
      Where ����id = ����id_In And ���� = '�²���';
    End If;
  End If;
  --ҽ������
  --����_In    Integer:=0, --0:����;1-סԺ
  --����_In    Integer:=1, --1-�շѵ�;2-���ʵ�
  --����_In    Integer:=0, --0:ɾ�����۵�;1-�շѻ����;2-�˷ѻ�����
  --No_In      ������ü�¼.No%Type,
  --ҽ��ids_In Varchar2 := Null
  Zl_ҽ������_�Ʒ�״̬_Update(0, 1, 1, No_In);

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����շѼ�¼_Insert;
/

--93623:���˺�,2016-02-25,����ҽ�����͵ļƷ�״̬
Create Or Replace Procedure Zl_������ʼ�¼_Delete
(
  No_In         ������ü�¼.No%Type,
  ���_In       Varchar2,
  ����Ա���_In ������ü�¼.����Ա���%Type,
  ����Ա����_In ������ü�¼.����Ա����%Type
) As
  --���ܣ�����һ��������ʵ�����ָ�������
  --��ţ���ʽ��"1,3,5,7,8",Ϊ�ձ�ʾ�������пɳ�����
  --�ù����������ָ��������

  --���α�ΪҪ�˷ѵ��ݵ�����ԭʼ��¼
  Cursor c_Bill(n_��־ Number) Is
    Select a.Id, a.�۸񸸺�, a.���, a.ִ��״̬, a.�շ����, a.ҽ�����, a.����id, a.������Ŀid, a.��������id, a.ִ�в���id, a.���˿���id, a.ʵ�ս��,
           Decode(a.��¼״̬, 0, 1, 0) As ����, j.�������, m.��������
    From ������ü�¼ A, ����ҽ����¼ J, �������� M
    Where a.ҽ����� = j.Id(+) And a.�շ�ϸĿid + 0 = m.����id(+) And a.No = No_In And a.��¼���� = 2 And a.��¼״̬ In (0, 1, 3) And
          a.�����־ = n_��־
    Order By a.�շ�ϸĿid, a.���;

  --���α����ڴ���ҩƷ����������
  --��Ҫ�ܷ��õ�ִ��״̬,��Ϊ���ڴ˲�����
  Cursor c_Stock(n_��־ Number) Is
    Select ID, �ⷿid, ҩƷid, ����, ����, ����, ʵ������, ����, ��ҩ��ʽ, ���Ч��, Ч��, ��Ʒ����, �ڲ�����, ����id
    From ҩƷ�շ���¼
    Where NO = No_In And ���� In (9, 25) And Mod(��¼״̬, 3) = 1 And ����� Is Null And
          ����id In (Select ID
                   From ������ü�¼
                   Where NO = No_In And ��¼���� = 2 And ��¼״̬ In (0, 1, 3) And �շ���� In ('4', '5', '6', '7') And �����־ = n_��־ And
                         (Instr(',' || ���_In || ',', ',' || ��� || ',') > 0 Or ���_In Is Null))
    Order By ҩƷid;

  --���α����ڴ�����ü�¼���
  Cursor c_Serial Is
    Select ���, �۸񸸺� From ������ü�¼ Where NO = No_In And ��¼���� = 2 And ��¼״̬ In (0, 1, 3) Order By ���;
  l_ҩƷ�շ� t_Numlist := t_Numlist();
  l_����     t_Numlist := t_Numlist();
  l_����id   t_Numlist := t_Numlist();
  n_�������� Number;

  v_ҽ��ids Varchar2(4000);

  n_ҽ��id   ����ҽ����¼.Id%Type;
  n_����     ������ü�¼.�۸񸸺�%Type;
  n_�����־ ������ü�¼.�����־%Type;

  --�����˷Ѽ������
  n_ʣ������ Number;
  n_ʣ��Ӧ�� Number;
  n_ʣ��ʵ�� Number;
  n_ʣ��ͳ�� Number;

  n_׼������ Number;
  n_�˷Ѵ��� Number;

  n_Ӧ�ս�� Number;
  n_ʵ�ս�� Number;
  n_ͳ���� Number;

  n_Dec Number;

  n_Count   Number;
  d_Curdate Date;
  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin
  --�Ƿ��Ѿ�ȫ����ȫִ��(ֻ�����ŵ��ݵļ��)
  Select Nvl(Count(*), 0), Max(Nvl(�����־, 1))
  Into n_Count, n_�����־
  From ������ü�¼
  Where NO = No_In And ��¼���� = 2 And ��¼״̬ In (0, 1, 3) And Nvl(ִ��״̬, 0) <> 1;
  If n_Count = 0 Then
    v_Err_Msg := '�õ����е���Ŀ�Ѿ�ȫ����ȫִ�У�';
    Raise Err_Item;
  End If;

  If Nvl(n_�����־, 0) = 0 Then
    n_�����־ := 1;
  End If;

  --δ��ȫִ�е���Ŀ�Ƿ���ʣ������(ֻ�����ŵ��ݵļ��)
  Select Nvl(Count(*), 0)
  Into n_Count
  From (Select ���, Sum(����) As ʣ������
         From (Select ��¼״̬, Nvl(�۸񸸺�, ���) As ���, Avg(Nvl(����, 1) * ����) As ����
                From ������ü�¼
                Where NO = No_In And ��¼���� = 2 And �����־ = n_�����־ And
                      Nvl(�۸񸸺�, ���) In
                      (Select Nvl(�۸񸸺�, ���)
                       From ������ü�¼
                       Where NO = No_In And ��¼���� = 2 And �����־ = n_�����־ And ��¼״̬ In (0, 1, 3) And Nvl(ִ��״̬, 0) <> 1)
                Group By ��¼״̬, Nvl(�۸񸸺�, ���))
         Group By ���
         Having Sum(����) <> 0);
  If n_Count = 0 Then
    v_Err_Msg := '�õ�����δ��ȫִ�в�����Ŀʣ������Ϊ��,û�п������ʵķ��ã�';
    Raise Err_Item;
  End If;

  ---------------------------------------------------------------------------------
  --���ñ���
  Select Sysdate Into d_Curdate From Dual;

  --���С��λ��
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_Dec From Dual;

  --ѭ������ÿ�з���(������Ŀ��)
  For r_Bill In c_Bill(n_�����־) Loop
    If Instr(',' || ���_In || ',', ',' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || ',') > 0 Or ���_In Is Null Then
    
      If r_Bill.���� = 0 Then
        If Nvl(r_Bill.ִ��״̬, 0) <> 1 Then
          --��ʣ������,ʣ��Ӧ��,ʣ��ʵ��
          Select Sum(Nvl(����, 1) * ����), Sum(Ӧ�ս��), Sum(ʵ�ս��), Sum(ͳ����)
          Into n_ʣ������, n_ʣ��Ӧ��, n_ʣ��ʵ��, n_ʣ��ͳ��
          From ������ü�¼
          Where NO = No_In And ��¼���� = 2 And ��� = r_Bill.���;
        
          If n_ʣ������ = 0 Then
            If ���_In Is Not Null Then
              v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����Ѿ�ȫ�����ʣ�';
              Raise Err_Item;
            End If;
            --�����δ�޶��к�,ԭʼ�����еĸñ��Ѿ�ȫ������(ִ��״̬=0��һ�ֿ���)
          Else
            --׼������(��ҩƷ��ĿΪʣ������,ԭʼ����)
            If Instr(',4,5,6,7,', r_Bill.�շ����) = 0 Or (r_Bill.�շ���� = '4' And Nvl(r_Bill.��������, 0) = 0) Then
            
              --@@@
              --��ҩƷ����(�Ծ���ҽ��ִ��Ϊ׼���м��)
              --: 1.����ҽ�����͵�,����ҽ��ִ��Ϊ׼(�����ܰ���:���;����;����;������Ѫ)
              --: 2.���ڲ���ҽ�ԼƼ��е��շѷ�ʽΪ:0-������ȡ ��,��֧�ֲ�����;�����������,��ֻ��ȫ��
              --: 3.������ҽ����,����ʣ������Ϊ׼
              n_Count := 0;
              If Instr(',C,D,F,G,K,', ',' || r_Bill.������� || ',') = 0 And r_Bill.������� Is Not Null Then
              
                Select Nvl(Sum(����), 0), Count(*)
                Into n_׼������, n_Count
                From (Select j.ҽ����� As ҽ��id, j.�շ�ϸĿid, Nvl(j.����, 1) * Nvl(j.����, 1) As ����
                       From ������ü�¼ J, ����ҽ����¼ M
                       Where j.ҽ����� = m.Id And j.No = No_In And j.��¼���� = 2 And j.��� = r_Bill.��� And j.��¼״̬ In (1, 3) And
                             Exists
                        (Select 1
                              From ����ҽ������ A
                              Where a.ҽ��id = j.ҽ����� And Nvl(a.ִ��״̬, 0) <> 1 And a.No || '' = No_In) And Exists
                        (Select 1
                              From ����ҽ���Ƽ� A
                              Where a.ҽ��id = j.ҽ����� And a.�շ�ϸĿid = j.�շ�ϸĿid And Nvl(a.�շѷ�ʽ, 0) = 0) And j.�۸񸸺� Is Null And
                             Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And
                             (j.��¼״̬ In (1, 3) And Not Exists
                              (Select 1
                               From ҩƷ�շ���¼
                               Where ����id = j.Id And Instr(',8,9,10,21,24,25,26,', ',' || ���� || ',') > 0) Or
                              j.��¼״̬ = 2 And Not Exists
                              (Select 1 From ҩƷ�շ���¼ Where NO = No_In And ���� In (8, 24) And ҩƷid = j.�շ�ϸĿid))
                       Union All
                       Select a.ҽ��id, a.�շ�ϸĿid, -1 * Nvl(a.����, 1) * Nvl(c.��������, 1) As ����
                       From ����ҽ���Ƽ� A, ����ҽ������ B, ����ҽ��ִ�� C, ������ü�¼ J, ����ҽ����¼ M
                       Where a.ҽ��id = b.ҽ��id And b.ҽ��id = c.ҽ��id And Nvl(a.�շѷ�ʽ, 0) = 0 And b.���ͺ� = c.���ͺ� And
                             a.ҽ��id = m.Id And Nvl(c.ִ�н��, 1) = 1 And Nvl(b.ִ��״̬, 0) <> 1 And a.ҽ��id = j.ҽ����� And
                             a.�շ�ϸĿid = j.�շ�ϸĿid And j.No = No_In And j.��¼���� = 2 And j.��� = r_Bill.��� And
                             j.��¼״̬ In (1, 3) And j.�۸񸸺� Is Null And Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And
                             Not Exists
                        (Select 1
                              From ҩƷ�շ���¼
                              Where ����id = j.Id And Instr(',8,9,10,21,24,25,26,', ',' || ���� || ',') > 0) And Not Exists
                        (Select 1 From �������� Where ����id = j.�շ�ϸĿid And Nvl(��������, 0) = 1)
                       Union All
                       Select a.ҽ��id, a.�շ�ϸĿid, 0 As ����
                       From ����ҽ���Ƽ� A, ������ü�¼ J, ����ҽ����¼ M
                       Where a.ҽ��id = m.Id And a.ҽ��id = j.ҽ����� And a.�շ�ϸĿid = j.�շ�ϸĿid And Nvl(a.�շѷ�ʽ, 0) <> 0 And
                             j.No = No_In And j.��¼���� = 2 And Nvl(j.ִ��״̬, 0) = 2 And Not Exists
                        (Select 1 From �������� Where ����id = j.�շ�ϸĿid And Nvl(��������, 0) = 1) And
                             Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0);
              
              End If;
            
              If Nvl(n_Count, 0) = 0 Then
                n_׼������ := n_ʣ������;
              End If;
            
            Else
              Select Sum(Nvl(����, 1) * ʵ������)
              Into n_׼������
              From ҩƷ�շ���¼
              Where NO = No_In And ���� In (9, 25) And Mod(��¼״̬, 3) = 1 And ����� Is Null And ����id = r_Bill.Id;
            
              --���������õ���������
              If r_Bill.�շ���� = '4' And Nvl(n_׼������, 0) = 0 Then
                n_׼������ := n_ʣ������;
              End If;
            End If;
          
            --����������ü�¼
          
            --�ñ���Ŀ�ڼ�������
            Select Nvl(Max(Abs(ִ��״̬)), 0) + 1
            Into n_�˷Ѵ���
            From ������ü�¼
            Where NO = No_In And ��¼���� = 2 And ��¼״̬ = 2 And ��� = r_Bill.���;
          
            --���=ʣ����*(׼����/ʣ����)
            n_Ӧ�ս�� := Round(n_ʣ��Ӧ�� * (n_׼������ / n_ʣ������), n_Dec);
            n_ʵ�ս�� := Round(n_ʣ��ʵ�� * (n_׼������ / n_ʣ������), n_Dec);
            n_ͳ���� := Round(n_ʣ��ͳ�� * (n_׼������ / n_ʣ������), n_Dec);
          
            --�����˷Ѽ�¼
            Insert Into ������ü�¼
              (ID, NO, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, Ӥ����, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�, ���˿���id, �շ����,
               �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ������,
               ִ����, ִ��״̬, ִ��ʱ��, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ������Ŀ��, ���մ���id, ͳ����, ���ʵ�id, ժҪ, ���ձ���, �Ƿ���, ����)
              Select ���˷��ü�¼_Id.Nextval, NO, ��¼����, 2, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, Ӥ����, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�,
                     ���˿���id, �շ����, �շ�ϸĿid, ���㵥λ, Decode(Sign(n_׼������ - Nvl(����, 1) * ����), 0, ����, 1), ��ҩ����,
                     Decode(Sign(n_׼������ - Nvl(����, 1) * ����), 0, -1 * ����, -1 * n_׼������), �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���,
                     ��׼����, -1 * n_Ӧ�ս��, -1 * n_ʵ�ս��, ��������id, ������, ִ�в���id, ������, ִ����, -1 * n_�˷Ѵ���, ִ��ʱ��, ����Ա���_In,
                     ����Ա����_In, ����ʱ��, d_Curdate, ������Ŀ��, ���մ���id, -1 * n_ͳ����, ���ʵ�id, ժҪ, ���ձ���, �Ƿ���, ����
              From ������ü�¼
              Where ID = r_Bill.Id;
          
            --��¼����ҽ�����Ѷ�Ӧ��ҽ��ID(����������)
            If n_ҽ��id Is Null And r_Bill.ҽ����� Is Not Null Then
              n_ҽ��id := r_Bill.ҽ�����;
            End If;
          
            --�������
            Update �������
            Set ������� = Nvl(�������, 0) - n_ʵ�ս��
            Where ����id = r_Bill.����id And ���� = 1 And ���� = 1;
            If Sql%RowCount = 0 Then
              Insert Into �������
                (����id, ����, ����, �������, Ԥ�����)
              Values
                (r_Bill.����id, 1, 1, -1 * n_ʵ�ս��, 0);
            End If;
          
            --����δ�����
            Update ����δ�����
            Set ��� = Nvl(���, 0) - n_ʵ�ս��
            Where ����id = r_Bill.����id And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And
                  Nvl(���˿���id, 0) = Nvl(r_Bill.���˿���id, 0) And Nvl(��������id, 0) = Nvl(r_Bill.��������id, 0) And
                  Nvl(ִ�в���id, 0) = Nvl(r_Bill.ִ�в���id, 0) And ������Ŀid + 0 = r_Bill.������Ŀid And ��Դ;�� + 0 = n_�����־;
            If Sql%RowCount = 0 Then
              Insert Into ����δ�����
                (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
              Values
                (r_Bill.����id, Null, Null, r_Bill.���˿���id, r_Bill.��������id, r_Bill.ִ�в���id, r_Bill.������Ŀid, n_�����־,
                 -1 * n_ʵ�ս��);
            End If;
          
            --���ԭ���ü�¼
            --ִ��״̬:ȫ������(׼����=ʣ����)���Ϊ0,������Ϊ1
            Update ������ü�¼
            Set ��¼״̬ = 3, ִ��״̬ = Decode(Sign(n_׼������ - n_ʣ������), 0, 0, 1)
            Where ID = r_Bill.Id;
          End If;
        Else
          If ���_In Is Not Null Then
            v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����Ѿ���ȫִ��,�������ʣ�';
            Raise Err_Item;
          End If;
          --���:û�޶��к�,ԭʼ�����а����Ѿ���ȫִ�е�
        End If;
      End If;
    End If;
  End Loop;

  ---------------------------------------------------------------------------------
  --ҩƷ�������
  ------------------------------------------------------------------------------------------------------------------------
  --�ȴ���������
  For v_���� In (Select ID, �ⷿid, ҩƷid, ����, ����, ����, ʵ������, ����, ��ҩ��ʽ, ���Ч��, Ч��, ��Ʒ����, �ڲ�����, ����id
               From ҩƷ�շ���¼
               Where ���� = 21 And Mod(��¼״̬, 3) = 1 And ����� Is Null And
                     ����id In (Select ID
                              From ������ü�¼
                              Where NO = No_In And ��¼���� = 2 And ��¼״̬ In (0, 1, 3) And �շ���� = '4' And �����־ = n_�����־ And
                                    (Instr(',' || ���_In || ',', ',' || ��� || ',') > 0 Or ���_In Is Null))
               Order By ҩƷid) Loop
    --����ҩƷ���
    If v_����.�ⷿid Is Not Null Then
      Update ҩƷ���
      Set �������� = Nvl(��������, 0) + Decode(v_����.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(v_����.����, 1) * Nvl(v_����.ʵ������, 0)
      Where �ⷿid = v_����.�ⷿid And ҩƷid = v_����.ҩƷid And Nvl(����, 0) = Nvl(v_����.����, 0) And ���� = 1;
      If Sql%RowCount = 0 Then
        Insert Into ҩƷ���
          (�ⷿid, ҩƷid, ����, ����, Ч��, ��������, �ϴ�����, �ϴβ���, ���Ч��, ��Ʒ����, �ڲ�����)
        Values
          (v_����.�ⷿid, v_����.ҩƷid, 1, v_����.����, v_����.Ч��,
           Decode(v_����.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(v_����.����, 1) * Nvl(v_����.ʵ������, 0), v_����.����, v_����.����, v_����.���Ч��,
           v_����.��Ʒ����, v_����.�ڲ�����);
      End If;
    End If;
    l_����id.Extend;
    l_����id(l_����id.Count) := v_����.����id;
    l_ҩƷ�շ�.Extend;
    l_ҩƷ�շ�(l_ҩƷ�շ�.Count) := v_����.Id;
  End Loop;

  For r_Stock In c_Stock(n_�����־) Loop
  
    --����ҩƷ���
    If r_Stock.�ⷿid Is Not Null Then
    
      Select Decode(Count(Column_Value), Null, 0, 0, 0, 1)
      Into n_��������
      From Table(l_����id)
      Where Column_Value = r_Stock.����id;
      If Nvl(n_��������, 0) = 0 Then
        Update ҩƷ���
        Set �������� = Nvl(��������, 0) + Decode(r_Stock.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(r_Stock.����, 1) * Nvl(r_Stock.ʵ������, 0)
        Where �ⷿid = r_Stock.�ⷿid And ҩƷid = r_Stock.ҩƷid And Nvl(����, 0) = Nvl(r_Stock.����, 0) And ���� = 1;
        If Sql%RowCount = 0 Then
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, Ч��, ��������, �ϴ�����, �ϴβ���, ���Ч��, ��Ʒ����, �ڲ�����)
          Values
            (r_Stock.�ⷿid, r_Stock.ҩƷid, 1, r_Stock.����, r_Stock.Ч��,
             Decode(r_Stock.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(r_Stock.����, 1) * Nvl(r_Stock.ʵ������, 0), r_Stock.����, r_Stock.����,
             r_Stock.���Ч��, r_Stock.��Ʒ����, r_Stock.�ڲ�����);
        End If;
      End If;
    End If;
  
    l_ҩƷ�շ�.Extend;
    l_ҩƷ�շ�(l_ҩƷ�շ�.Count) := r_Stock.Id;
  End Loop;

  --ɾ��ҩƷ�շ���¼
  Forall I In 1 .. l_ҩƷ�շ�.Count
    Delete From ҩƷ�շ���¼ Where ID = l_ҩƷ�շ�(I);

  ------------------------------------------------------------------------------------------------------------------------
  --����ɾδ��ҩƷ��¼

  Delete From δ��ҩƷ��¼ A
  Where NO = No_In And ���� In (9, 25) And Not Exists
   (Select 1
         From ҩƷ�շ���¼
         Where ���� = a.���� And Nvl(�ⷿid, 0) = Nvl(a.�ⷿid, 0) And NO = No_In And Mod(��¼״̬, 3) = 1 And ����� Is Null);
  ------------------------------------------------------------------------------------------------------------------------

  ---------------------------------------------------------------------------------
  --����ǻ���,ֱ��ɾ�����ü�¼(ҩƷ�����)
  n_Count   := 0;
  v_ҽ��ids := Null;
  For r_Bill In c_Bill(n_�����־) Loop
    If Instr(',' || ���_In || ',', ',' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || ',') > 0 Or ���_In Is Null Then
      If r_Bill.���� = 1 Then
        If Nvl(r_Bill.ִ��״̬, 0) <> 1 Then
          l_����.Extend;
          l_����(l_����.Count) := r_Bill.Id;
        
          --Delete From ������ü�¼ Where ID = r_Bill.ID;
          n_Count := n_Count + 1; --��¼�Ƿ���ɾ����
        
          If r_Bill.ҽ����� Is Not Null Then
            If Instr(',' || Nvl(v_ҽ��ids, '') || ',', ',' || r_Bill.ҽ����� || ',') = 0 Then
              v_ҽ��ids := Nvl(v_ҽ��ids, '') || ',' || r_Bill.ҽ�����;
            End If;
          
            --��¼����ҽ�����Ѷ�Ӧ��ҽ��ID(����������)
            If n_ҽ��id Is Null Then
              n_ҽ��id := r_Bill.ҽ�����;
            End If;
          End If;
        Else
          If ���_In Is Not Null Then
            v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����Ѿ���ȫִ��,�������ʣ�';
            Raise Err_Item;
          End If;
          --���:û�޶��к�,ԭʼ�����а����Ѿ���ȫִ�е�
        End If;
      End If;
    End If;
  End Loop;

  --ɾ�����ۼ�¼
  Forall I In 1 .. l_����.Count
    Delete From ������ü�¼ Where ID = l_����(I);

  --ɾ��֮����ͳһ�������
  If n_Count > 0 Then
    n_Count := 1;
    For r_Serial In c_Serial Loop
      If r_Serial.�۸񸸺� Is Null Then
        n_���� := n_Count;
      End If;
    
      Update ������ü�¼
      Set ��� = n_Count, �۸񸸺� = Decode(�۸񸸺�, Null, Null, n_����)
      Where NO = No_In And ��¼���� = 2 And ��� = r_Serial.���;
    
      Update ������ü�¼ Set �������� = n_Count Where NO = No_In And ��¼���� = 2 And �������� = r_Serial.���;
    
      n_Count := n_Count + 1;
    End Loop;
  
  End If;

  --���ŵ���ȫ������ʱ��ɾ������ҽ������
  If ���_In Is Null And n_ҽ��id Is Not Null Then
    Select Nvl(Count(*), 0)
    Into n_Count
    From (Select ���, Sum(����) As ʣ������
           From (Select ��¼״̬, Nvl(�۸񸸺�, ���) As ���, Avg(Nvl(����, 1) * ����) As ����
                  From ������ü�¼
                  Where NO = No_In And ��¼���� = 2 And ҽ����� + 0 = n_ҽ��id
                  Group By ��¼״̬, Nvl(�۸񸸺�, ���))
           Group By ���
           Having Nvl(Sum(����), 0) <> 0);
    If n_Count = 0 Then
      Delete From ����ҽ������ Where ҽ��id = n_ҽ��id And ��¼���� = 2 And NO = No_In;
    End If;
  End If;
  If v_ҽ��ids Is Not Null Then
    --ҽ������
    --����_In    Integer:=0, --0:����;1-סԺ
    --����_In    Integer:=1, --1-�շѵ�;2-���ʵ�
    --����_In    Integer:=0, --0:ɾ�����۵�;1-�շѻ����;2-�˷ѻ�����
    --No_In      ������ü�¼.No%Type,
    --ҽ��ids_In Varchar2 := Null
    v_ҽ��ids := Substr(v_ҽ��ids, 2);
    Zl_ҽ������_�Ʒ�״̬_Update(0, 2, 0, No_In, v_ҽ��ids);
  Else
    Zl_ҽ������_�Ʒ�״̬_Update(0, 2, 2, No_In, v_ҽ��ids);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_������ʼ�¼_Delete;
/

--93623:���˺�,2016-02-25,����ҽ�����͵ļƷ�״̬
Create Or Replace Procedure Zl_�����շѼ�¼_Delete
(
  No_In           ������ü�¼.No%Type,
  ����Ա���_In   ������ü�¼.����Ա���%Type,
  ����Ա����_In   ������ü�¼.����Ա����%Type,
  ҽ�����㷽ʽ_In Varchar2 := Null,
  ���_In         Varchar2 := Null,
  ���㷽ʽ_In     ����Ԥ����¼.���㷽ʽ%Type := Null,
  ���_In         ������ü�¼.ʵ�ս��%Type := 0,
  �˷�ʱ��_In     ������ü�¼.�Ǽ�ʱ��%Type := Null,
  ����Ʊ��_In     Number := 0,
  �˷�ժҪ_In     ������ü�¼.ժҪ%Type := Null,
  У�Ա�־_In     Number := 0,
  ����id_In       ����Ԥ����¼.����id%Type := Null,
  �������_In     ����Ԥ����¼.�������%Type := Null,
  һ��ͨ����_In   Varchar2 := Null,
  �˿����_In     Number := 0,
  �൥��ȫ��_In   Number := 0
) As
  --���ܣ�ɾ��һ�������շѵ��� 
  --������ 
  --        ҽ�����㷽ʽ_IN   =ҽ���˷�ʱ,��֧�ֽ������ϵĽ��㷽ʽ,���Ϊ�ձ�ʾ��ҽ���˷ѻ�ҽ���˷�ȫ�������������ϡ� 
  --        ���_IN           =Ҫ�˷ѵ���Ŀ���,��ʽΪ"1,3,5,6...",ȱʡNULL��ʾ��"δ�˵�"�����С� 
  --        ���㷽ʽ_IN       =��Ϊ�����˷�ʱ,�˷ѽ��Ľ��㷽ʽ�� 
  --        ���_IN           =ָ�˷�ʱ�²����������,�����˷ѻ�ҽ��ȫ�˵�ĳ�ֽ������ֽ�ʱ�Ż�����µ��� 
  --                           ��ʱ��������ڼ��㱾���˷ѵĽ�����,�����ü�¼�Ĵ����ڱ�����ִ��������Zl_�����շ����_Insert���� 
  --        ����Ʊ��_In       =0:����ȫ�˻����һ��ȫ��ʱ�ջ�Ʊ��,ע��,���ŵ����˷�ѭ����������ʱֻ�ջ�һ�Ρ� 
  --                           1:�����˷Ѳ�����Ʊ��,ͨ���ش���õ������� 
  --        У�Ա�־_IN:0-����Ҫ�϶�;1-��϶�(��������Ա�ɿ����,������Ʊ��,������Ԥ�����) 
  --        �˿����_In:1-���в�����(���˿ʽ�˵�ָ���Ľ��㷽ʽ<���㷽ʽ_In>��,0-��ָ���˿ʽ) 
  --        �൥��ȫ��_IN=1-�൥��ȫ��(���ŵ���ȫ��,ԭ����);0-��ԭ����
  --���α�ΪҪ�˷ѵ��ݵ�����ԭʼ��¼ 

  --ҽ��ȫ�˵�ĳ�ֽ������ֽ�Ӷ��������µ����ʱ,�ſ��˴�������,ִ���걾���̺�,��������е������������ 
  Cursor c_Bill Is
    Select a.Id, a.No, a.���ӱ�־, a.�շ�ϸĿid, a.���, a.�۸񸸺�, a.ִ��״̬, a.�շ����, a.����, a.����, a.ҽ�����, j.�������, m.��������,
           Nvl(a.���ӱ�־, 0) As ���
    From ������ü�¼ A, ����ҽ����¼ J, �������� M
    Where a.ҽ����� = j.Id(+) And a.No = No_In And a.��¼���� = 1 And a.��¼״̬ In (1, 3) And a.�շ�ϸĿid + 0 = m.����id(+) And
          Nvl(a.���ӱ�־, 0) <> Decode(�൥��ȫ��_In, 1, 999, 9)
    Order By a.�շ�ϸĿid, a.���;
  --:����ԭʼ�������,��Ӧ�ø��ݵ�ǰ�˷Ѳ������������д���
  -- Decode(Sign(���_In), 0, 999, 9)

  --���α����ڴ���ҩƷ���������� 
  --��Ҫ�ܷ��õ�ִ��״̬,��Ϊ���ڴ˲����� 
  Cursor c_Stock Is
    Select ID, ҩƷid, �ⷿid, ����, ����, ����, ʵ������, ����, ��ҩ��ʽ, ���Ч��, Ч��, ��Ʒ����, �ڲ�����, ����id
    From ҩƷ�շ���¼
    Where NO = No_In And ���� In (8, 24) --@@@ 
          And Mod(��¼״̬, 3) = 1 And ����� Is Null And
          ����id In (Select ID
                   From ������ü�¼
                   Where NO = No_In And ��¼���� = 1 And ��¼״̬ In (1, 3) And �շ���� In ('4', '5', '6', '7') --@@@ 
                         And (Instr(',' || ���_In || ',', ',' || ��� || ',') > 0 Or ���_In Is Null))
    Order By ҩƷid;

  --���α����ڴ���δ��ҩƷ��¼ 
  Cursor c_Spare Is
    Select NO, �ⷿid, ���� From δ��ҩƷ��¼ Where NO = No_In And ���� In (8, 24); --@@@ 

  --�ù�����ڴ�����Ա�ɿ�������˵Ĳ�ͬ���㷽ʽ�Ľ�� 
  Cursor c_Money(����id_In ����Ԥ����¼.����id%Type) Is
    Select ���㷽ʽ, ��Ԥ��
    From ����Ԥ����¼
    Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = ����id_In And ���㷽ʽ Is Not Null And Nvl(��Ԥ��, 0) <> 0 And Nvl(У�Ա�־, 0) = 0;

  --���α����ڲ����շ�ʱʹ�ù��ĳ�Ԥ�����¼ 
  Cursor c_Deposit(V����id ����Ԥ����¼.����id%Type) Is
    Select ID, ��Ԥ�� As ���, Ԥ�����
    From ����Ԥ����¼
    Where ��¼���� In (1, 11) And ��¼״̬ In (1, 3) And ����id = V����id And Nvl(��Ԥ��, 0) <> 0
    Order By ID Desc;

  n_����id   ������Ϣ.����id%Type;
  n_����id   ������ü�¼.����id%Type;
  n_������� ����Ԥ����¼.�������%Type;
  n_��ӡid   Ʊ�ݴ�ӡ����.Id%Type;

  n_���˽�� ����Ԥ����¼.��Ԥ��%Type;
  n_Ԥ����� ����Ԥ����¼.��Ԥ��%Type;
  n_����ֵ   ����Ԥ����¼.��Ԥ��%Type;
  n_ԭ���� ������ü�¼.ʵ�ս��%Type;
  --�����˷Ѽ������ 
  n_ʣ������ Number;
  n_ʣ��Ӧ�� Number;
  n_ʣ��ʵ�� Number;
  n_ʣ��ͳ�� Number;
  n_�������� Number;
  n_׼������ Number;
  n_�˷Ѵ��� Number;

  n_Ӧ�ս�� Number;
  n_ʵ�ս�� Number;
  n_ͳ���� Number;
  n_�ܽ��   Number;
  n_����״̬ ������ü�¼.����״̬%Type;
  n_�����˷� Number; --�Ƿ��һ���˷���ȫ���˷�,��ÿ���˷ѹ������жϵõ��� 
  n_��id     ����ɿ����.Id%Type;

  v_�˷ѽ��� ���㷽ʽ.����%Type;
  v_�������� Varchar2(500);
  n_������   Number(2);

  l_����id   t_Numlist := t_Numlist();
  l_ҩƷ�շ� t_Numlist := t_Numlist();
  l_ʹ��id   t_Numlist := t_Numlist();
  n_Dec      Number;
  d_Date     Date;
  n_Count    Number;
  n_ԭ����id Number;

  Err_Item Exception;
  v_Err_Msg      Varchar2(255);
  n_����ģʽ     Number(3);
  v_Para         Varchar2(1000);
  n_ҽ��ִ�мƼ� Number;

  Procedure Zl_Square_Update
  (
    ԭ����id_In ����Ԥ����¼.����id%Type,
    �ֽ���id_In ����Ԥ����¼.����id%Type,
    �ɿ���id_In ����Ԥ����¼.�ɿ���id%Type,
    �˿�ʱ��_In ����Ԥ����¼.�տ�ʱ��%Type,
    �������_In ����Ԥ����¼.�������%Type,
    ��������_In Varchar2 := Null
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
    For v_У�� In (Select a.Id As Ԥ��id, c.���ѿ�id, c.������, c.�ӿڱ��, c.����, c.���, c.Id
                 From ����Ԥ����¼ A, ���˿�������� B, ���˿������¼ C
                 Where a.Id = b.Ԥ��id And b.������id = c.Id And a.��¼���� = 3 And a.��¼״̬ = 1 And
                       Instr(Nvl(��������_In, '_LXH'), ',' || a.���㷽ʽ || ',') = 0 And a.����id = ԭ����id_In) Loop
    
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
                 -1 * ��Ԥ��, �ֽ���id_In, �ɿ���id_In, Ԥ�����, �����id, Nvl(���㿨���, v_У��.�ӿڱ��), ����, ������ˮ��, ����˵��, ������λ,
                 Decode(Nvl(У�Ա�־_In, 0), 0, 0, Decode(Nvl(v_У��.���ѿ�id, 0), 0, 1, 2)), �������_In, 3
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
        Update ���ѿ�Ŀ¼ Set ��� = Nvl(���, 0) + v_У��.������ Where ID = Nvl(v_У��.���ѿ�id, 0);
      End If;
    
      Select ���˿������¼_Id.Nextval Into n_Id From Dual;
      Insert Into ���˿������¼
        (ID, �ӿڱ��, ���ѿ�id, ���, ��¼״̬, ���㷽ʽ, ������, ����, ������ˮ��, ����ʱ��, ��ע, �����־)
        Select n_Id, �ӿڱ��, ���ѿ�id, ���, n_��¼״̬, ���㷽ʽ, -1 * v_У��.������, ����, ������ˮ��, ����ʱ��, ��ע,
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
  n_��id   := Zl_Get��id(����Ա����_In);
  n_������ := 0;
  --�Ƿ��Ѿ�ȫ����ȫִ��(ֻ�Ǹõ������ŵ��ݵļ��) 
  Select Nvl(Count(*), 0)
  Into n_Count
  From ������ü�¼
  Where NO = No_In And ��¼���� = 1 And ��¼״̬ In (1, 3) And Nvl(ִ��״̬, 0) <> 1;
  If n_Count = 0 Then
    v_Err_Msg := '�õ����е���Ŀ�Ѿ�ȫ����ȫִ�У�';
    Raise Err_Item;
  End If;

  --δ��ȫִ�е���Ŀ�Ƿ���ʣ������(ֻ�����ŵ��ݵļ��) 
  --ִ��״̬��ԭʼ��¼���ж� 
  Select Nvl(Count(*), 0)
  Into n_Count
  From (Select ���, Sum(����) As ʣ������
         From (Select ��¼״̬, Nvl(�۸񸸺�, ���) As ���, Avg(Nvl(����, 1) * ����) As ����
                From ������ü�¼
                Where NO = No_In And ��¼���� = 1 And Nvl(���ӱ�־, 0) <> 9 And
                      Nvl(�۸񸸺�, ���) In
                      (Select Nvl(�۸񸸺�, ���)
                       From ������ü�¼
                       Where NO = No_In And ��¼���� = 1 And ��¼״̬ In (1, 3) And Nvl(ִ��״̬, 0) <> 1)
                Group By ��¼״̬, Nvl(�۸񸸺�, ���))
         Group By ���
         Having Sum(����) <> 0);
  If n_Count = 0 Then
    v_Err_Msg := '�õ�����δ��ȫִ�в�����Ŀʣ������Ϊ��,û�п����˷ѵķ��ã�';
    Raise Err_Item;
  End If;
  --ȷ���Ƿ���ҽ��ִ�мƼ��д�������,�����������,�����ҽ��ִ�мƼ۽����˷�,���򰴾ɷ�ʽ���д���
  Select Count(1)
  Into n_ҽ��ִ�мƼ�
  From ������ü�¼ A, ҽ��ִ�мƼ� B
  Where a.ҽ����� = b.ҽ��id And a.��¼���� = 1 And a.No = No_In And a.��¼״̬ In (1, 3) And Rownum = 1;

  --------------------------------------------------------------------------------- 
  --���ñ��� 
  If �˷�ʱ��_In Is Not Null Then
    d_Date := �˷�ʱ��_In;
  Else
    Select Sysdate Into d_Date From Dual;
  End If;
  If ����id_In Is Null Then
    Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
  Else
    n_����id := ����id_In;
  End If;
  n_������� := �������_In;
  If n_������� Is Null Then
    n_������� := ����id_In;
  End If;
  --���С��λ�� 
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_Dec From Dual;

  --��ȡ���㷽ʽ���� 
  v_�˷ѽ��� := ���㷽ʽ_In;
  If v_�˷ѽ��� Is Null Then
    Begin
      Select ���� Into v_�˷ѽ��� From ���㷽ʽ Where ���� = 1;
    Exception
      When Others Then
        v_�˷ѽ��� := '�ֽ�';
    End;
  End If;
  --ѭ������ÿ�з���(������Ŀ��) 
  n_�ܽ��   := 0;
  n_�����˷� := 1;
  For r_Bill In c_Bill Loop
    If Instr(',' || ���_In || ',', ',' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || ',') > 0 Or ���_In Is Null Then
      If Nvl(r_Bill.ִ��״̬, 0) <> 1 Then
        --��ʣ������,ʣ��Ӧ��,ʣ��ʵ�� 
        Select Sum(Nvl(����, 1) * ����), Sum(Ӧ�ս��), Sum(ʵ�ս��), Sum(ͳ����)
        Into n_ʣ������, n_ʣ��Ӧ��, n_ʣ��ʵ��, n_ʣ��ͳ��
        From ������ü�¼
        Where NO = No_In And ��¼���� = 1 And ��� = r_Bill.���;
      
        If n_ʣ������ = 0 Then
          If ���_In Is Not Null Then
            v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����Ѿ�ȫ���˷ѣ�';
            Raise Err_Item;
          End If;
          --�����δ�޶��к�,ԭʼ�����еĸñ��Ѿ�ȫ���˷�(ִ��״̬=0��һ�ֿ���) 
          n_�����˷� := 0;
        Else
          --׼������(��ҩƷ��ĿΪʣ������,ԭʼ����) 
          If Instr(',4,5,6,7,', r_Bill.�շ����) = 0 Or (r_Bill.�շ���� = '4' And Nvl(r_Bill.��������, 0) = 0) Then
            --@@@ 
            --��ҩƷ����(�Ծ���ҽ��ִ��Ϊ׼���м��) 
            --: 1.����ҽ�����͵�,����ҽ��ִ��Ϊ׼(�����ܰ���:���;����;����;������Ѫ) 
            --: 2.������ҽ����,����ʣ������Ϊ׼ 
            n_Count := 0;
            If Instr(',C,D,F,G,K,', ',' || r_Bill.������� || ',') = 0 And r_Bill.������� Is Not Null Then
              If n_ҽ��ִ�мƼ� = 1 Then
                Select Decode(Sign(Sum(����)), -1, 0, Sum(����)), Count(*)
                Into n_׼������, n_Count
                From (Select Max(Decode(a.��¼״̬, 2, 0, a.Id)) As ID, Max(a.ҽ�����) As ҽ��id, Max(a.�շ�ϸĿid) As �շ�ϸĿid,
                              Sum(Nvl(a.����, 1) * Nvl(a.����, 1)) As ����,
                              Sum(Decode(a.��¼״̬, 2, 0, Nvl(a.����, 1) * Nvl(a.����, 1))) As ԭʼ����
                       From ������ü�¼ A, ����ҽ����¼ M
                       Where a.ҽ����� = m.Id And Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And
                             Instr('5,6,7', a.�շ����) = 0 And a.No = No_In And a.��� = r_Bill.��� And a.��¼���� = 1 And
                             a.��¼״̬ In (1, 2, 3) And a.�۸񸸺� Is Null
                       Group By a.���
                       Union All
                       Select a.Id, a.ҽ����� As ҽ��id, a.�շ�ϸĿid, -1 * b.���� As ��ִ��, 0 ԭʼ����
                       From ������ü�¼ A, ҽ��ִ�мƼ� B, ����ҽ����¼ M
                       Where a.ҽ����� = b.ҽ��id And a.�շ�ϸĿid = b.�շ�ϸĿid + 0 And a.ҽ����� = m.Id And
                             Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And Instr('5,6,7', a.�շ����) = 0 And
                             (Exists
                              (Select 1
                               From ����ҽ��ִ��
                               Where b.ҽ��id = ҽ��id And b.���ͺ� = ���ͺ� And b.Ҫ��ʱ�� = Ҫ��ʱ�� And Nvl(ִ�н��, 0) = 1) Or Exists
                              (Select 1
                               From ����ҽ������
                               Where b.ҽ��id = ҽ��id And b.���ͺ� = ���ͺ� And Nvl(ִ��״̬, 0) = 1)) And Not Exists
                        (Select 1
                              From ����ҽ������
                              Where a.ҽ����� = ҽ��id And a.No = NO And Mod(a.��¼����, 10) = ��¼����) And a.No = No_In And
                             a.��� = r_Bill.��� And a.��¼���� = 1 And a.��¼״̬ In (1, 3) ��and a.�۸񸸺� Is Null) Q1
                Where Not Exists (Select 1
                       From ҩƷ�շ���¼
                       Where ����id = Q1.Id And Instr(',8,9,10,21,24,25,26,', ',' || ���� || ',') > 0) Having
                 Max(ID) <> 0;
              Else
              
                Select Nvl(Sum(����), 0), Count(*)
                Into n_׼������, n_Count
                From (Select a.ҽ��id, a.�շ�ϸĿid, Nvl(a.����, 1) * Nvl(b.��������, 1) As ����
                       From ����ҽ���Ƽ� A, ����ҽ������ B, ������ü�¼ J, ����ҽ����¼ M
                       Where a.ҽ��id = b.ҽ��id And a.ҽ��id = m.Id And Nvl(b.ִ��״̬, 0) <> 1 And a.ҽ��id = j.ҽ����� And
                             a.�շ�ϸĿid = j.�շ�ϸĿid And j.No = No_In And j.��¼���� = 1 And j.��� = r_Bill.��� And
                             j.��¼״̬ In (1, 3) And j.�۸񸸺� Is Null And Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And
                             Exists
                        (Select 1
                              From ����ҽ���Ƽ� A
                              Where a.ҽ��id = j.ҽ����� And a.�շ�ϸĿid = j.�շ�ϸĿid And Nvl(a.�շѷ�ʽ, 0) = 0) And Not Exists
                        (Select 1
                              From ҩƷ�շ���¼
                              Where ����id = j.Id And Instr(',8,9,10,21,24,25,26,', ',' || ���� || ',') > 0)
                       Union All
                       Select a.ҽ��id, a.�շ�ϸĿid, -1 * Nvl(a.����, 1) * Nvl(c.��������, 1) As ����
                       From ����ҽ���Ƽ� A, ����ҽ������ B, ����ҽ��ִ�� C, ������ü�¼ J, ����ҽ����¼ M
                       Where a.ҽ��id = b.ҽ��id And b.ҽ��id = c.ҽ��id And b.���ͺ� = c.���ͺ� And a.ҽ��id = m.Id And
                             Nvl(c.ִ�н��, 1) = 1 And Nvl(b.ִ��״̬, 0) <> 1 And a.ҽ��id = j.ҽ����� And a.�շ�ϸĿid = j.�շ�ϸĿid And
                             j.No = No_In And j.��¼���� = 1 And Nvl(a.�շѷ�ʽ, 0) = 0 And j.��� = r_Bill.��� And j.��¼״̬ In (1, 3) And
                             j.�۸񸸺� Is Null And Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And Not Exists
                        (Select 1
                              From ҩƷ�շ���¼
                              Where ����id = j.Id And Instr(',8,9,10,21,24,25,26,', ',' || ���� || ',') > 0) And Not Exists
                        (Select 1 From �������� Where ����id = j.�շ�ϸĿid And Nvl(��������, 0) = 1)
                       Union All
                       Select a.ҽ����� As ҽ��id, a.�շ�ϸĿid, Nvl(a.����, 1) * a.���� As ����
                       From ������ü�¼ A, ����ҽ����¼ M
                       Where a.ҽ����� = m.Id And Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And a.No = No_In And
                             a.��¼���� = 1 And a.��� = r_Bill.��� And a.��¼״̬ = 2 And a.�۸񸸺� Is Null And Not Exists
                        (Select 1 From ҩƷ�շ���¼ Where NO = No_In And ���� In (8, 24) And ҩƷid = a.�շ�ϸĿid));
              End If;
            End If;
            If Nvl(n_Count, 0) <> 0 And n_׼������ = 0 Then
              v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з�����ִ��,�������˷ѣ�';
              Raise Err_Item;
            End If;
          
            If Nvl(n_Count, 0) = 0 Then
              n_׼������ := n_ʣ������;
            End If;
          
          Else
            Select Nvl(Sum(Nvl(����, 1) * ʵ������), 0), Count(*)
            Into n_׼������, n_Count
            From ҩƷ�շ���¼
            Where NO = No_In And ���� In (8, 24) And Mod(��¼״̬, 3) = 1 --@@@ 
                  And ����� Is Null And ����id = r_Bill.Id;
          
            --��ʣ��������׼������������������� 
            --1.���������õ������޶�Ӧ���շ���¼,��ʱʹ��ʣ������ 
            --2.��������,��ʱ�ѷ�ҩ���� 
            If n_׼������ = 0 Then
              If r_Bill.�շ���� = '4' Then
                If n_Count > 0 Then
                  v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����ѷ���,�����Ϻ����˷ѣ�';
                  Raise Err_Item;
                Else
                  n_׼������ := n_ʣ������;
                End If;
              Else
                v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����ѷ�ҩ,����ҩ�����˷ѣ�';
                Raise Err_Item;
              End If;
            End If;
          End If;
        
          --�Ƿ񲿷��˷� 
          If r_Bill.ִ��״̬ = 2 Or n_׼������ <> Nvl(r_Bill.����, 1) * r_Bill.���� Then
            n_�����˷� := 0;
          End If;
        
          --����������ü�¼ 
          n_����״̬ := 0;
          --�ñ���Ŀ�ڼ����˷� 
          If Nvl(У�Ա�־_In, 0) <> 0 Then
            n_�˷Ѵ��� := -9; --�ȱ���,�̶�Ϊ9 
            n_����״̬ := 1;
          Else
            Select Nvl(Max(Abs(ִ��״̬)), 0) + 1
            Into n_�˷Ѵ���
            From ������ü�¼
            Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 2 And Nvl(ִ��״̬, 0) < 0 And ��� = r_Bill.���;
          End If;
        
          --���=ʣ����*(׼����/ʣ����) 
          If Nvl(r_Bill.���, 0) = 9 Then
            --�����Գ������õ�С��λ(����:ҽ�����㳬��С��λ��,���Ϳ��ܳ���С��λ
            n_Ӧ�ս�� := Round(n_ʣ��Ӧ�� * (n_׼������ / n_ʣ������), 5);
            n_ʵ�ս�� := Round(n_ʣ��ʵ�� * (n_׼������ / n_ʣ������), 5);
            n_ͳ���� := Round(n_ʣ��ͳ�� * (n_׼������ / n_ʣ������), 5);
          Else
            n_Ӧ�ս�� := Round(n_ʣ��Ӧ�� * (n_׼������ / n_ʣ������), n_Dec);
            n_ʵ�ս�� := Round(n_ʣ��ʵ�� * (n_׼������ / n_ʣ������), n_Dec);
            n_ͳ���� := Round(n_ʣ��ͳ�� * (n_׼������ / n_ʣ������), n_Dec);
          End If;
          n_�ܽ�� := n_�ܽ�� + n_ʵ�ս��;
        
          --�����˷Ѽ�¼ 
          Insert Into ������ü�¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�, ���˿���id, �շ����,
             �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ������, ִ����,
             ִ��״̬, ����״̬, ִ��ʱ��, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, ����,
             �ɿ���id)
            Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�,
                   ���˿���id, �շ����, �շ�ϸĿid, ���㵥λ, Decode(Sign(n_׼������ - Nvl(����, 1) * ����), 0, ����, 1), ��ҩ����,
                   Decode(Sign(n_׼������ - Nvl(����, 1) * ����), 0, -1 * ����, -1 * n_׼������), �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����,
                   -1 * n_Ӧ�ս��, -1 * n_ʵ�ս��, ��������id, ������, ִ�в���id, ������, ִ����, -1 * n_�˷Ѵ���, n_����״̬, ִ��ʱ��, ����Ա���_In,
                   ����Ա����_In, ����ʱ��, d_Date, n_����id, -1 * n_ʵ�ս��, ������Ŀ��, ���մ���id, -1 * n_ͳ����, Nvl(�˷�ժҪ_In, ժҪ),
                   Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, ����, n_��id
            From ������ü�¼
            Where ID = r_Bill.Id;
        
          --���ԭ���ü�¼ 
          --ִ��״̬:ȫ������(׼����=ʣ����)���Ϊ0,������Ϊ1,�쳣�շѵ�,���Ǳ���9 
          Update ������ü�¼
          Set ��¼״̬ = 3, ִ��״̬ = Decode(Nvl(ִ��״̬, 0), 9, 9, Decode(Sign(n_׼������ - n_ʣ������), 0, 0, 1))
          Where ID = r_Bill.Id;
        End If;
      Else
        If ���_In Is Not Null Then
          v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����Ѿ���ȫִ��,�����˷ѣ�';
          Raise Err_Item;
        End If;
        --���:û�޶��к�,ԭʼ�����а����Ѿ���ȫִ�е� 
        n_�����˷� := 0;
      End If;
    Else
      n_�����˷� := 0; --δָ���ñ�,���ڲ����˷� 
    End If;
  End Loop;
  --------------------------------------------------------------------------------- 
  --������Ԥ����¼ 

  --ԭ���ݵĽ���ID 
  Select ����id, ����id
  Into n_ԭ����id, n_����id
  From ������ü�¼
  Where NO = No_In And ��¼���� = 1 And ��¼״̬ In (1, 3) And Rownum = 1;

  If n_�����˷� = 1 And Nvl(�˿����_In, 0) = 0 Then
    --���ݵ�һ���˷���ȫ������ 
    --��Ԥ�����ּ�¼ 
    Insert Into ����Ԥ����¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��, ����id,
       �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, У�Ա�־, �������, ��������)
      Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, d_Date,
             ����Ա����_In, ����Ա���_In, -1 * ��Ԥ��, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ,
             Decode(У�Ա�־_In, 1, 2, У�Ա�־_In), n_�������, 3
      From ����Ԥ����¼
      Where ��¼���� In (1, 11) And ����id = n_ԭ����id And Nvl(��Ԥ��, 0) <> 0;
    If Nvl(У�Ա�־_In, 0) = 0 Then
      --������Ԥ����� 
      For v_Ԥ�� In (Select Ԥ�����, Nvl(Sum(Nvl(��Ԥ��, 0)), 0) As Ԥ�����
                   From ����Ԥ����¼
                   Where ��¼���� In (1, 11) And ����id = n_ԭ����id
                   Group By Ԥ�����
                   Having Sum(Nvl(��Ԥ��, 0)) <> 0) Loop
        Update �������
        Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(v_Ԥ��.Ԥ�����, 0)
        Where ����id = n_����id And ���� = 1 And ���� = Nvl(v_Ԥ��.Ԥ�����, 2)
        Returning Ԥ����� Into n_����ֵ;
        If Sql%RowCount = 0 Then
          Insert Into �������
            (����id, ����, Ԥ�����, ����)
          Values
            (n_����id, Nvl(v_Ԥ��.Ԥ�����, 2), Nvl(v_Ԥ��.Ԥ�����, 0), 1);
          n_����ֵ := n_Ԥ�����;
        End If;
        If n_����ֵ = 0 Then
          Delete From ������� Where ����id = n_����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
        End If;
      End Loop;
    End If;
    --��ҽ��ȫ��,��ҽ�����н��㷽ʽ���������,ԭ���˻�(��Ԥ����ǰ���Ѵ���) 
    If ҽ�����㷽ʽ_In Is Null Then
      v_�������� := ',' || Nvl(һ��ͨ����_In, '-Lxh') || ',' || Nvl(һ��ͨ����_In, 'Lxh') || ',';
    
      --һ��ͨ�����ѿ������п������������Ҫ���⴦��,��Ҫ���϶�. 
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����,
         �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, У�Ա�־, �������, ��������)
        Select ����Ԥ����¼_Id.Nextval, ��¼����, NO, 2, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, d_Date, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���_In, ����Ա����_In,
               -1 * ��Ԥ��, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ,
               Case
                 When Nvl(�����id, 0) <> 0 Then
                  Decode(У�Ա�־_In, 1, 1, 0) * 1
                 When Nvl(���㿨���, 0) <> 0 Then
                  Decode(У�Ա�־_In, 1, 1, 0) * 1
                 When Nvl(q.Ԥ��id, 0) <> 0 Then
                  Decode(У�Ա�־_In, 1, 1, 0) * 1
                 When Nvl(j.����, '-') <> '-' Then
                  Decode(У�Ա�־_In, 1, 1, 0)
                 Else
                  Decode(У�Ա�־_In, 1, 2, 0)
               End As У�Ա�־, n_�������, 3
        From ����Ԥ����¼ A, (Select ���� From ���㷽ʽ Where ���� In (3, 4)) J,
             (Select m.Id As Ԥ��id
               From ����Ԥ����¼ M, һ��ͨĿ¼ C
               Where m.����id = n_ԭ����id And m.���㷽ʽ = c.���㷽ʽ And m.��¼���� = 3 And m.��¼״̬ = 1) Q
        Where a.��¼���� = 3 And a.��¼״̬ = 1 And a.����id = n_ԭ����id And a.Id = q.Ԥ��id(+) And a.���㷽ʽ = j.����(+) And
              Instr(v_��������, ',' || ���㷽ʽ || ',') = 0 And
              (Not Exists (Select 1 From ���˿�������� Where a.Id = Ԥ��id) Or Nvl(���㿨���, 0) = 0);
    
      --�������ѿ�,���㿨��������Ѿ������� 
      Zl_Square_Update(n_ԭ����id, n_����id, n_��id, d_Date, n_�������, v_��������);
      --b.���µľ��������ӿ�֧�ֵ�������,���������ϵĽ��㷽ʽ,���ϵ�ָ���Ľ��㷽ʽ��,�������(��Ϊ������������֮�������) 
      If һ��ͨ����_In Is Not Null Then
        Begin
          Select -1 * Nvl(Sum(��Ԥ��), 0) Into n_���˽�� From ����Ԥ����¼ Where ����id = n_����id;
        Exception
          When Others Then
            n_���˽�� := 0;
        End;
      
        If (n_�ܽ�� - n_���˽��) <> 0 Then
          --��ʱ���ܽ�û�а������,��Ϊ����������ڵ��ñ����̺�Ų��������ü�¼ 
          Insert Into ����Ԥ����¼
            (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����,
             ������ˮ��, ����˵��, ������λ, У�Ա�־, �������, ��������)
            Select ����Ԥ����¼_Id.Nextval, 3, NO, 2, ����id, ��ҳid, '�����˷ѽ���', v_�˷ѽ���, d_Date, ����Ա���_In, ����Ա����_In,
                   -1 * (n_�ܽ�� - n_���˽��), n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ,
                   Decode(У�Ա�־_In, 1, 2, 0), n_�������, 3
            From ����Ԥ����¼
            Where ��¼���� = 3 And ��¼״̬ = 1 And ����id = n_ԭ����id And Rownum = 1;
          n_������ := 1;
        End If;
      End If;
      --ҽ�����������ϵĽ��㷽ʽ��,�������,�˵�ָ���Ľ��㷽ʽ�� 
      --��Ҫ���������
    Else
      --a.ԭ���˻� 
      v_�������� := ',' || ҽ�����㷽ʽ_In || ',' || Nvl(һ��ͨ����_In, '-Lxh') || ',' || v_�˷ѽ��� || ',';
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����,
         ������ˮ��, ����˵��, ������λ, У�Ա�־, �������, ��������)
        Select ����Ԥ����¼_Id.Nextval, ��¼����, NO, 2, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, d_Date, ����Ա���_In, ����Ա����_In, -1 * ��Ԥ��, n_����id,
               n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ,
               
               Case
                 When Nvl(�����id, 0) <> 0 Then
                  Decode(У�Ա�־_In, 1, 1, 0) * 1
                 When Nvl(���㿨���, 0) <> 0 Then
                  Decode(У�Ա�־_In, 1, 1, 0) * 1
                 When Nvl(q.Ԥ��id, 0) <> 0 Then
                  Decode(У�Ա�־_In, 1, 1, 0) * 1
                 When Nvl(j.����, '-') <> '-' Then
                  Decode(У�Ա�־_In, 1, 1, 0)
                 Else
                  Decode(У�Ա�־_In, 1, 2, 0)
               End As У�Ա�־, n_�������, 3
        From ����Ԥ����¼ A, (Select ���� From ���㷽ʽ Where ���� In (3, 4)) J,
             (Select m.Id As Ԥ��id
               From ����Ԥ����¼ M, һ��ͨĿ¼ C
               Where m.����id = n_ԭ����id And m.���㷽ʽ = c.���㷽ʽ And m.��¼���� = 3 And m.��¼״̬ = 1) Q
        Where a.��¼���� = 3 And a.��¼״̬ = 1 And a.���㷽ʽ = j.����(+) And a.����id = n_ԭ����id And
              Instr(v_��������, ',' || a.���㷽ʽ || ',') = 0 And a.Id = q.Ԥ��id(+) And
              (Not Exists (Select 1 From ���˿�������� Where a.Id = Ԥ��id) Or Nvl(���㿨���, 0) = 0);
    
      --�������ѿ�,���㿨��������Ѿ������� 
      Zl_Square_Update(n_ԭ����id, n_����id, n_��id, d_Date, n_�������, v_��������);
    
      --b.���µľ���ҽ�����������ϵĽ��㷽ʽ,���ϵ�ָ���Ľ��㷽ʽ��,�������(��Ϊ������������֮�������) 
      Begin
        Select -1 * Nvl(Sum(��Ԥ��), 0) Into n_���˽�� From ����Ԥ����¼ Where ����id = n_����id;
      Exception
        When Others Then
          n_���˽�� := 0;
      End;
    
      If (n_�ܽ�� - n_���˽��) <> 0 Then
        --��ʱ���ܽ�û�а������,��Ϊ����������ڵ��ñ����̺�Ų��������ü�¼ 
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����,
           ������ˮ��, ����˵��, ������λ, У�Ա�־, �������, ��������)
          Select ����Ԥ����¼_Id.Nextval, 3, NO, 2, ����id, ��ҳid, Decode(һ��ͨ����_In, Null, '����ҽ���ӿ��˷�', '����ҽ���ӿں������ӿ��˷�'), v_�˷ѽ���,
                 d_Date, ����Ա���_In, ����Ա����_In, -1 * (n_�ܽ�� - n_���˽��), n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��,
                 ������λ, Decode(У�Ա�־_In, 1, 2, 0), n_�������, 3
          From ����Ԥ����¼
          Where ��¼���� = 3 And ��¼״̬ = 1 And ����id = n_ԭ����id And Rownum = 1;
        n_������ := 1;
      End If;
    
    End If;
  Else
    ------------------------------------------------- 
    --�����˷�ֱ����Ϊָ�����㷽ʽ 
    Insert Into ����Ԥ����¼
      (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id,
       ���㿨���, ����, ������ˮ��, ����˵��, ������λ, У�Ա�־, �������, ��������)
      Select ����Ԥ����¼_Id.Nextval, 3, No_In, 2, ����id, ��ҳid, '�����˷ѽ���', v_�˷ѽ���, d_Date, Null, Null, Null, ����Ա���_In, ����Ա����_In,
             -1 * (n_�ܽ�� + Nvl(���_In, 0)), n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ,
             Decode(У�Ա�־, 1, 2, 0), n_�������, 3
      From ����Ԥ����¼
      Where ��¼���� = 3 And ��¼״̬ In (1, 3) And ����id = n_ԭ����id And Rownum = 1;
  
    --����շ�ʱֻʹ����Ԥ����,��Ҫ��Ԥ��,���ҿ����ж�ʳ�Ԥ�� 
    If Sql%RowCount = 0 Then
      n_Ԥ����� := n_�ܽ�� + Nvl(���_In, 0);
    
      For r_Deposit In c_Deposit(n_ԭ����id) Loop
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
           ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, У�Ա�־, �������, ��������)
          Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�,
                 d_Date, ����Ա����_In, ����Ա���_In, Decode(Sign(r_Deposit.��� - n_Ԥ�����), -1, -1 * r_Deposit.���, -1 * n_Ԥ�����),
                 n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, Decode(У�Ա�־_In, 1, 2, 0), n_�������, 3
          From ����Ԥ����¼
          Where ID = r_Deposit.Id;
      
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
      If Nvl(У�Ա�־_In, 0) = 0 Then
        --���²���Ԥ����� 
        Update �������
        Set Ԥ����� = Nvl(Ԥ�����, 0) + n_�ܽ�� + Nvl(���_In, 0)
        Where ����id = n_����id And ���� = 1 And ���� = 1
        Returning Ԥ����� Into n_����ֵ;
        If Sql%RowCount = 0 Then
          Insert Into ������� (����id, ����, Ԥ�����, ����) Values (n_����id, 1, n_�ܽ�� + Nvl(���_In, 0), 1);
          n_����ֵ := n_�ܽ�� + Nvl(���_In, 0);
        End If;
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete From ������� Where ����id = n_����id And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
        End If;
      End If;
    End If;
  End If;

  --����ԭ��¼
  Update ����Ԥ����¼ Set ��¼״̬ = 3 Where ��¼���� = 3 And ��¼״̬ In (1, 3) And ����id = n_ԭ����id;

  If �൥��ȫ��_In <> 1 Then
    --���������,�൥��ȫ��ʱ ,��ԭ����������
    --�������ļ�¼״̬����Ϊ3
    If Nvl(���_In, 0) <> 0 Then
      n_Count := 1;
      If n_�����˷� = 1 And Nvl(�˿����_In, 0) = 0 Then
        n_ԭ���� := 0;
        --ԭ����,���������
        If n_������ = 0 Then
          Select -1 * Nvl(Sum(ʵ�ս��), 0)
          Into n_ԭ����
          From ������ü�¼ A
          Where NO = No_In And a.��¼���� = 1 And a.��¼״̬ In (1, 3) And Nvl(a.���ӱ�־, 0) = 9;
        End If;
        If Nvl(n_ԭ����, 0) <> 0 Or Nvl(���_In, 0) <> 0 Then
          Update ����Ԥ����¼
          Set ��Ԥ�� = ��Ԥ�� - n_ԭ���� - Nvl(���_In, 0)
          Where ���㷽ʽ = v_�˷ѽ��� And ����id = n_����id;
          If Sql%NotFound Then
            Insert Into ����Ԥ����¼
              (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����,
               �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, У�Ա�־, �������, ��������)
              Select ����Ԥ����¼_Id.Nextval, 3, No_In, 2, ����id, ��ҳid, '����', v_�˷ѽ���, d_Date, Null, Null, Null, ����Ա���_In,
                     ����Ա����_In, -1 * n_ԭ���� - Nvl(���_In, 0), n_����id, n_��id, Ԥ�����, Null, Null, Null, Null, Null, Null, 0,
                     n_�������, 3
              From ����Ԥ����¼
              Where ��¼���� = 3 And ��¼״̬ In (1, 3) And ����id = n_ԭ����id And Rownum = 1;
          End If;
        End If;
      End If;
    Elsif n_�����˷� = 1 And Nvl(�˿����_In, 0) = 0 Then
      --ԭ����ʱ,��Ҫ����Ԥ����¼��������
      Select Nvl(Sum(Nvl(���ʽ��, 0)), 0) Into n_ʵ�ս�� From ������ü�¼ Where ����id = n_����id;
      Select Nvl(Sum(Nvl(��Ԥ��, 0)), 0) Into n_����ֵ From ����Ԥ����¼ Where ����id = n_����id;
      If Abs(n_ʵ�ս��) <> Abs(n_����ֵ) Then
        n_ʵ�ս�� := n_ʵ�ս�� - n_����ֵ;
        Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� + Nvl(n_ʵ�ս��, 0) Where ���㷽ʽ = v_�˷ѽ��� And ����id = n_����id;
        If Sql%NotFound Then
          Insert Into ����Ԥ����¼
            (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����,
             �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, У�Ա�־, �������, ��������)
            Select ����Ԥ����¼_Id.Nextval, 3, No_In, 2, ����id, ��ҳid, '����', v_�˷ѽ���, d_Date, Null, Null, Null, ����Ա���_In,
                   ����Ա����_In, Nvl(n_ʵ�ս��, 0), n_����id, n_��id, Ԥ�����, Null, Null, Null, Null, Null, Null, 0, n_�������, 3
            From ����Ԥ����¼
            Where ��¼���� = 3 And ��¼״̬ In (1, 3) And ����id = n_ԭ����id And Rownum = 1;
        End If;
      End If;
    End If;
  
    Select Nvl(Sum(Nvl(���ʽ��, 0)), 0) Into n_ʵ�ս�� From ������ü�¼ Where ����id = n_����id;
    Select Nvl(Sum(Nvl(��Ԥ��, 0)), 0) Into n_����ֵ From ����Ԥ����¼ Where ����id = n_����id;
  
    n_ʵ�ս�� := n_ʵ�ս�� - n_����ֵ;
  
    If n_ʵ�ս�� <> 0 Then
      --δ�ҵ����²��������
      Zl_�����շ����_Insert(No_In, n_ʵ�ս��, 1);
    End If;
  End If;
  --------------------------------------------------------------------------------- 
  --��Ա�ɿ����(ע����Ԥ����¼�����Ŵ������������ʻ��ȵĽ�����,�����˳�Ԥ����) 
  --�������ҪУ�Ե�,�ݲ�������Ա�ɿ���� 
  If Nvl(У�Ա�־_In, 0) = 0 Then
    For r_Moneyrow In c_Money(n_����id) Loop
      Update ��Ա�ɿ����
      Set ��� = Nvl(���, 0) + r_Moneyrow.��Ԥ��
      Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = r_Moneyrow.���㷽ʽ
      Returning ��� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ��Ա�ɿ����
          (�տ�Ա, ���㷽ʽ, ����, ���)
        Values
          (����Ա����_In, r_Moneyrow.���㷽ʽ, 1, r_Moneyrow.��Ԥ��);
        n_����ֵ := r_Moneyrow.��Ԥ��;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From ��Ա�ɿ����
        Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = r_Moneyrow.���㷽ʽ And Nvl(���, 0) = 0;
      End If;
    End Loop;
  End If;

  --------------------------------------------------------------------------------- 
  --�˷�Ʊ�ݻ���(��ȫ��ʱ�Ż���,�����������ش�����л���) 
  If ����Ʊ��_In = 0 Then
  
    --���ñ�־||NO;ִ�п���(����);�վݷ�Ŀ(��ҳ����,����);�շ�ϸĿ(����)
    v_Para     := Nvl(zl_GetSysParameter('Ʊ�ݷ������', 1121), '0||0;0;0,0;0');
    n_����ģʽ := Zl_To_Number(Substr(v_Para, 1, 1));
    If n_����ģʽ <> 0 Then
      --�ջ�Ʊ��
      Select ʹ��id Bulk Collect
      Into l_ʹ��id
      From (Select Distinct b.ʹ��id From Ʊ�ݴ�ӡ��ϸ B Where b.No = No_In And Nvl(b.Ʊ��, 0) = 1);
    
      n_����ģʽ := l_ʹ��id.Count;
      If l_ʹ��id.Count <> 0 Then
        --������ռ�¼
        Forall I In 1 .. l_ʹ��id.Count
          Insert Into Ʊ��ʹ����ϸ
            (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ����, ʹ��ʱ��)
            Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, ����Ա����_In, d_Date
            From Ʊ��ʹ����ϸ A
            Where ID = l_ʹ��id(I) And ���� = 1 And Not Exists
             (Select 1 From Ʊ��ʹ����ϸ Where ���� = a.���� And Ʊ�� = a.Ʊ�� And Nvl(����, 0) <> 1);
      
        Forall I In 1 .. l_ʹ��id.Count
          Update Ʊ�ݴ�ӡ��ϸ Set �Ƿ���� = 1 Where ʹ��id = l_ʹ��id(I) And Nvl(�Ƿ����, 0) = 0;
      
      End If;
    End If;
    If n_����ģʽ = 0 Then
      --��ȡ�������һ�εĴ�ӡID(�����Ƕ��ŵ����շѴ�ӡ) 
      Begin
        Select ID
        Into n_��ӡid
        From (Select b.Id
               From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
               Where a.��ӡid = b.Id And a.���� = 1 And b.�������� = 1 And b.No = No_In
               Order By a.ʹ��ʱ�� Desc)
        Where Rownum < 2;
      Exception
        When Others Then
          Null;
      End;
      --������ǰû�д�ӡ,���ջ� 
      If n_��ӡid Is Not Null Then
        --a.���ŵ���ѭ������ʱֻ���ջ�һ�� 
        Select Count(*) Into n_Count From Ʊ��ʹ����ϸ Where Ʊ�� = 1 And ���� = 2 And ��ӡid = n_��ӡid;
        If n_Count = 0 Then
          Insert Into Ʊ��ʹ����ϸ
            (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����)
            Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_Date, ����Ա����_In
            From Ʊ��ʹ����ϸ
            Where ��ӡid = n_��ӡid And Ʊ�� = 1 And ���� = 1;
        Else
          --b.�����˷Ѷ���ջ�ʱ,���һ��ȫ���ջ�Ҫ�ſ����ջص� 
          Insert Into Ʊ��ʹ����ϸ
            (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����)
            Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, d_Date, ����Ա����_In
            From Ʊ��ʹ����ϸ A
            Where ��ӡid = n_��ӡid And Ʊ�� = 1 And ���� = 1 And Not Exists
             (Select 1 From Ʊ��ʹ����ϸ B Where a.���� = b.���� And ��ӡid = n_��ӡid And Ʊ�� = 1 And ���� = 2);
        End If;
      End If;
    End If;
  End If;

  --------------------------------------------------------------------------------- 
  --�������� 
  For v_���� In (Select ID, ҩƷid, �ⷿid, ����, ����, ����, ʵ������, ����, ��ҩ��ʽ, ���Ч��, Ч��, ��Ʒ����, �ڲ�����, ����id
               From ҩƷ�շ���¼
               Where ���� = 21 --@@@ 
                     And Mod(��¼״̬, 3) = 1 And ����� Is Null And
                     ����id In (Select ID
                              From ������ü�¼
                              Where NO = No_In And ��¼���� = 1 And ��¼״̬ In (1, 3) And �շ���� = '4' --@@@ 
                                    And (Instr(',' || ���_In || ',', ',' || ��� || ',') > 0 Or ���_In Is Null))
               Order By ҩƷid) Loop
    --����ҩƷ��� 
    If v_����.�ⷿid Is Not Null Then
      Update ҩƷ���
      Set �������� = Nvl(��������, 0) + Decode(v_����.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(v_����.����, 1) * Nvl(v_����.ʵ������, 0)
      Where �ⷿid = v_����.�ⷿid And ҩƷid = v_����.ҩƷid And Nvl(����, 0) = Nvl(v_����.����, 0) And ���� = 1;
      If Sql%RowCount = 0 Then
        Insert Into ҩƷ���
          (�ⷿid, ҩƷid, ����, ����, Ч��, ��������, �ϴ�����, �ϴβ���, ���Ч��, ��Ʒ����, �ڲ�����) --@@@ 
        Values
          (v_����.�ⷿid, v_����.ҩƷid, 1, v_����.����, v_����.Ч��,
           Decode(v_����.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(v_����.����, 1) * Nvl(v_����.ʵ������, 0), v_����.����, v_����.����, v_����.���Ч��,
           v_����.��Ʒ����, v_����.�ڲ�����);
      End If;
    End If;
  
    l_����id.Extend;
    l_����id(l_����id.Count) := v_����.����id;
    l_ҩƷ�շ�.Extend;
    l_ҩƷ�շ�(l_ҩƷ�շ�.Count) := v_����.Id;
  End Loop;

  --ҩƷ������� 
  For r_Stock In c_Stock Loop
    --����ҩƷ��� 
    If r_Stock.�ⷿid Is Not Null Then
    
      Select Decode(Count(Column_Value), Null, 0, 0, 0, 1)
      Into n_��������
      From Table(l_����id)
      Where Column_Value = r_Stock.����id;
      If Nvl(n_��������, 0) = 0 Then
      
        Update ҩƷ���
        Set �������� = Nvl(��������, 0) + Decode(r_Stock.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(r_Stock.����, 1) * Nvl(r_Stock.ʵ������, 0)
        Where �ⷿid = r_Stock.�ⷿid And ҩƷid = r_Stock.ҩƷid And Nvl(����, 0) = Nvl(r_Stock.����, 0) And ���� = 1;
        If Sql%RowCount = 0 Then
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, Ч��, ��������, �ϴ�����, �ϴβ���, ���Ч��, ��Ʒ����, �ڲ�����) --@@@ 
          Values
            (r_Stock.�ⷿid, r_Stock.ҩƷid, 1, r_Stock.����, r_Stock.Ч��,
             Decode(r_Stock.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(r_Stock.����, 1) * Nvl(r_Stock.ʵ������, 0), r_Stock.����, r_Stock.����,
             r_Stock.���Ч��, r_Stock.��Ʒ����, r_Stock.�ڲ�����);
        End If;
      End If;
    End If;
    l_ҩƷ�շ�.Extend;
    l_ҩƷ�շ�(l_ҩƷ�շ�.Count) := r_Stock.Id;
  End Loop;

  --ɾ��ҩƷ�շ���¼ 
  Forall I In 1 .. l_ҩƷ�շ�.Count
    Delete From ҩƷ�շ���¼ Where ID = l_ҩƷ�շ�(I) And ����� Is Null;

  --δ��ҩƷ��¼ 
  For r_Spare In c_Spare Loop
    Select Nvl(Count(*), 0)
    Into n_Count
    From ҩƷ�շ���¼
    Where NO = No_In And ���� = r_Spare.���� --@@@ 
          And Mod(��¼״̬, 3) = 1 And ����� Is Null And Nvl(�ⷿid, 0) = Nvl(r_Spare.�ⷿid, 0);
  
    If n_Count = 0 Then
      Delete From δ��ҩƷ��¼
      Where ���� = r_Spare.���� --@@@ 
            And NO = No_In And Nvl(�ⷿid, 0) = Nvl(r_Spare.�ⷿid, 0);
    End If;
  End Loop;

  --ҽ������
  --����_In    Integer:=0, --0:����;1-סԺ
  --����_In    Integer:=1, --1-�շѵ�;2-���ʵ�
  --����_In    Integer:=0, --0:ɾ�����۵�;1-�շѻ����;2-�˷ѻ�����
  --No_In      ������ü�¼.No%Type,
  --ҽ��ids_In Varchar2 := Null
  Zl_ҽ������_�Ʒ�״̬_Update(0, 1, 2, No_In);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����շѼ�¼_Delete;
/

--93623:���˺�,2016-02-25,����ҽ�����͵ļƷ�״̬
Create Or Replace Procedure Zl_סԺ���ʼ�¼_Delete
(
  No_In           סԺ���ü�¼.No%Type,
  ���_In         Varchar2,
  ����Ա���_In   סԺ���ü�¼.����Ա���%Type,
  ����Ա����_In   סԺ���ü�¼.����Ա����%Type,
  ��¼����_In     סԺ���ü�¼.��¼����%Type := 2,
  ����״̬_In     Number := 0,
  ��Һ��ҩ���_In Number := 1,
  �Ǽ�ʱ��_In     סԺ���ü�¼.�Ǽ�ʱ��%Type := Sysdate
) As
  --���ܣ�����һ��סԺ���ʵ�����ָ�������
  --��ţ���ʽ��"1,3,5,7,8",��"1:2:33456,3:2,5:2,7:2,8:2",ð��ǰ������ֱ�ʾ�к�,�м�����ֱ�ʾ�˵�����,��������ֱ�ʾ��ҩ��¼��ID,Ŀǰ�����������ʱ�Ŵ���
  --      Ϊ�ձ�ʾ�������пɳ�����
  --��¼����:    2-�˹����ʵ�,3-�Զ����ʵ�
  --��Һ��ҩ���:    0-ҽ�����ã������ҩƷ�Ƿ������Һ��ҩ���ģ�1-��ҽ�����ã����ҩƷ�Ƿ������ҩ����
  --�ù����������ָ��������
  --����״̬_In:0-��ʾֱ������;1-��ʾ�������(ͨ����������-->�����������)
  --���α�ΪҪ�˷ѵ��ݵ�����ԭʼ��¼
  Cursor c_Bill Is
    Select ID, �۸񸸺�, ���, ִ��״̬, ��¼����, �շ����, ҽ�����, �շ�ϸĿid, ����id, ��ҳid, ������Ŀid, ��������id, ���˿���id, ִ�в���id, ���˲���id, ����, ����
    From סԺ���ü�¼
    Where NO = No_In And ��¼���� = ��¼����_In And ��¼״̬ In (0, 1, 3) And �����־ = 2
    Order By �շ�ϸĿid, ���;

  --���α����ڴ���ҩƷ����������
  --��Ҫ�ܷ��õ�ִ��״̬,��Ϊ���ڴ˲�����
  Cursor c_Stock(v_���_In Varchar2) Is
    Select ID, ����, NO, �ⷿid, ҩƷid, ����, ��ҩ��ʽ, ����, ʵ������, ���Ч��, Ч��, ����, ����, ��������, ����id, ��Ʒ����, �ڲ�����
    From ҩƷ�շ���¼
    Where NO = No_In And ���� In (9, 10, 25, 26) And Mod(��¼״̬, 3) = 1 And ����� Is Null And
          ����id In (Select ID
                   From סԺ���ü�¼
                   Where NO = No_In And ��¼���� = ��¼����_In And ��¼״̬ In (0, 1, 3) And �շ���� In ('4', '5', '6', '7') And
                         �����־ = 2 And (Instr(',' || v_���_In || ',', ',' || ��� || ',') > 0 Or v_���_In Is Null))
    Order By ҩƷid, �������� Desc;

  r_Stock c_Stock%RowType;
  --���α����ڴ�����ü�¼���
  Cursor c_Serial Is
    Select ���, �۸񸸺�
    From סԺ���ü�¼
    Where NO = No_In And ��¼���� = ��¼����_In And ��¼״̬ In (0, 1, 3)
    Order By ���;

  Cursor Cr_ҩƷ Is
    Select ID, ����, NO, �ⷿid, ҩƷid, ����, ��ҩ��ʽ, 0 As ����, ���Ч��, Ч��, ����, ����, ��������, ����id
    From ҩƷ�շ���¼
    Where Rownum <= 1;
  v_ҩƷ Cr_ҩƷ%RowType;

  v_ҽ��id ����ҽ����¼.Id%Type;
  n_����   Number;
  v_����   סԺ���ü�¼.�۸񸸺�%Type;
  v_���   Varchar2(2000);
  v_Tmp    Varchar2(4000);

  v_ҽ��ids    Varchar2(4000);
  l_ҩƷ�շ�   t_Numlist := t_Numlist();
  l_����       t_Numlist := t_Numlist();
  l_����id     t_Numlist := t_Numlist();
  n_����       Number;
  n_����ⷿid ҩƷ�շ���¼.�ⷿid%Type;
  n_��������id ҩƷ�շ���¼.Id%Type;
  n_�ⷿid     ҩƷ�շ���¼.�ⷿid%Type;
  n_����ֵ     Number;
  --�����˷Ѽ������
  v_ʣ������ Number;
  v_ʣ��Ӧ�� Number;
  v_ʣ��ʵ�� Number;
  v_ʣ��ͳ�� Number;

  v_׼������ Number;
  v_�˷Ѵ��� Number;
  v_Ӧ�ս�� Number;
  v_ʵ�ս�� Number;
  v_ͳ���� Number;
  n_Temp     Number;
  n_�������� Number;
  v_Dec      Number;
  n_Count    Number;
  v_Curdate  Date;
  Err_Item Exception;
  v_Err_Msg        Varchar2(255);
  n_��������       Number;
  n_����id         ������ҳ.����id%Type;
  n_��ҳid         ������ҳ.��ҳid%Type;
  n_��˱�־       ������ҳ.��˱�־%Type;
  n_סԺ״̬       ������ҳ.״̬%Type;
  n_������˷�ʽ   Number(2);
  n_δ��ƽ�ֹ���� Number(2);
  v_��ҩid         Varchar2(4000);
  Type Ty_ҩƷ Is Ref Cursor;
  c_ҩƷ Ty_ҩƷ; --�α����

Begin
  --�������ʱ,��ҩƷ�ᴫ���кŵ���������
  If Not ���_In Is Null Then
    If Instr(���_In, ':') > 0 Then
      v_Tmp := ���_In || ',';
      While Not v_Tmp Is Null Loop
        v_��� := v_��� || ',' || Substr(v_Tmp, 1, Instr(v_Tmp, ':') - 1);
        If Instr(Substr(v_Tmp, Instr(v_Tmp, ':') + 1, Instr(v_Tmp, ',') - Instr(v_Tmp, ':') - 1), ':') > 0 Then
          v_��ҩid := v_��ҩid || ',' ||
                    Substr(v_Tmp, Instr(v_Tmp, ':', 1, 2) + 1, Instr(v_Tmp, ',') - Instr(v_Tmp, ':', 1, 2) - 1);
        End If;
        v_Tmp := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
      End Loop;
      v_��� := Substr(v_���, 2);
      If v_��ҩid Is Not Null Then
        v_��ҩid := Substr(v_��ҩid, 2);
      End If;
    Else
      v_��� := ���_In;
    End If;
  End If;

  --�Ƿ��Ѿ�ȫ����ȫִ��(ֻ�����ŵ��ݵļ��)
  Select Nvl(Count(*), 0), Nvl(Max(����id), 0), Nvl(Max(��ҳid), 0)
  Into n_Count, n_����id, n_��ҳid
  From סԺ���ü�¼
  Where NO = No_In And ��¼���� = ��¼����_In And ��¼״̬ In (0, 1, 3) And Nvl(ִ��״̬, 0) <> 1 And �����־ = 2;
  If n_Count = 0 Then
    v_Err_Msg := '�õ����е���Ŀ�Ѿ�ȫ����ȫִ�У�';
    Raise Err_Item;
  End If;

  n_������˷�ʽ   := Nvl(zl_GetSysParameter(185), 0);
  n_δ��ƽ�ֹ���� := Nvl(zl_GetSysParameter(215), 0);
  If n_������˷�ʽ = 1 Or n_δ��ƽ�ֹ���� = 1 Then
  
    Begin
      Select ��˱�־, ״̬ Into n_��˱�־, n_סԺ״̬ From ������ҳ Where ����id = n_����id And ��ҳid = n_��ҳid;
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

  --δ��ȫִ�е���Ŀ�Ƿ���ʣ������(ֻ�����ŵ��ݵļ��)
  Select Nvl(Count(*), 0)
  Into n_Count
  From (Select ���, Sum(����) As ʣ������
         From (Select ��¼״̬, Nvl(�۸񸸺�, ���) As ���, Avg(Nvl(����, 1) * ����) As ����
                From סԺ���ü�¼
                Where NO = No_In And ��¼���� = ��¼����_In And �����־ = 2 And
                      Nvl(�۸񸸺�, ���) In
                      (Select Nvl(�۸񸸺�, ���)
                       From סԺ���ü�¼
                       Where NO = No_In And ��¼���� = ��¼����_In And ��¼״̬ In (0, 1, 3) And Nvl(ִ��״̬, 0) <> 1)
                Group By ��¼״̬, Nvl(�۸񸸺�, ���))
         Group By ���
         Having Sum(����) <> 0);
  If n_Count = 0 Then
    v_Err_Msg := '�õ�����δ��ȫִ�в�����Ŀʣ������Ϊ��,û�п������ʵķ��ã�';
    Raise Err_Item;
  End If;

  --ҽ�����ã��������ִ�е�ҽ��(ע����ִ�е������������,��Ϊ���� ���_IN ����������ý���������)
  If Nvl(����״̬_In, 0) <> 1 Then
    --�������������̵ģ������ҽ��ִ��״̬
    Select Nvl(Count(*), 0)
    Into n_Count
    From ����ҽ������
    Where ִ��״̬ = 3 And (NO, ��¼����, ҽ��id) In
          (Select NO, ��¼����, ҽ�����
                        From סԺ���ü�¼
                        Where NO = No_In And ��¼���� = ��¼����_In And ��¼״̬ In (0, 1, 3) And ҽ����� Is Not Null And
                              (Instr(',' || v_��� || ',', ',' || ��� || ',') > 0 Or v_��� Is Null));
    If n_Count > 0 Then
      v_Err_Msg := 'Ҫ���ʵķ����д��ڶ�Ӧ��ҽ������ִ�е�������������ʣ�';
      Raise Err_Item;
    End If;
  End If;

  ---------------------------------------------------------------------------------
  --�ȴ�ҩƷ��Ӧ���ݼ�,��ȷ����ǰ������������,Ϊ�˴������ж�
  --�������α�������ȡ��"����� is Null"��������Ϊ�����ҩ���ܲ������ѷ�
  Open c_Stock(v_���);

  --���ñ���
  Select �Ǽ�ʱ��_In Into v_Curdate From Dual;

  --���С��λ��
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into v_Dec From Dual;

  For c_��Ŀ���� In (Select a.����
                 From ������Ϣ A, ������ҳ B
                 Where a.����id = b.����id And b.��Ŀ���� Is Not Null And
                       (b.����id, b.��ҳid) In
                       (Select Distinct ����id, ��ҳid
                        From סԺ���ü�¼
                        Where NO = No_In And ��¼���� = ��¼����_In And ��¼״̬ In (0, 1, 3) And �����־ = 2)) Loop
    v_Err_Msg := '���ˡ�' || c_��Ŀ����.���� || '�� �Ѿ���������Ŀ,���ܱ����ʣ�';
    Raise Err_Item;
  End Loop;
  v_ҽ��ids := Null;
  --ѭ������ÿ�з���(������Ŀ��)
  For r_Bill In c_Bill Loop
    --����Ѿ����ڲ�����Ŀ��,���ܽ������ʴ���
    If Instr(',' || v_��� || ',', ',' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || ',') > 0 Or v_��� Is Null Then
      Select Decode(��¼״̬, 0, 1, 0) Into n_���� From סԺ���ü�¼ Where ID = r_Bill.Id;
      If Nvl(r_Bill.ִ��״̬, 0) <> 1 Then
        --��ʣ������,ʣ��Ӧ��,ʣ��ʵ��
        Select Sum(Nvl(����, 1) * ����), Sum(Ӧ�ս��), Sum(ʵ�ս��), Sum(ͳ����)
        Into v_ʣ������, v_ʣ��Ӧ��, v_ʣ��ʵ��, v_ʣ��ͳ��
        From סԺ���ü�¼
        Where NO = No_In And ��¼���� = ��¼����_In And ��� = r_Bill.���;
        n_�������� := 0;
        If v_ʣ������ = 0 Then
          If v_��� Is Not Null Then
            v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����Ѿ�ȫ�����ʣ�';
            Raise Err_Item;
          End If;
          --�����δ�޶��к�,ԭʼ�����еĸñ��Ѿ�ȫ������(ִ��״̬=0��һ�ֿ���)
        Else
        
          If Instr(���_In, ':') > 0 Then
            v_Tmp := ',' || ���_In;
            v_Tmp := Substr(v_Tmp, Instr(v_Tmp, ',' || r_Bill.��� || ':') + Length(',' || r_Bill.��� || ':'));
            v_Tmp := Substr(v_Tmp, 1, Instr(v_Tmp || ',', ',') - 1);
            If Instr(v_Tmp, ':') > 0 Then
              v_Tmp := Substr(v_Tmp, 1, Instr(v_Tmp, ':') - 1);
            End If;
            v_׼������ := v_Tmp;
            n_�������� := 1;
          End If;
        
          --׼������(��ҩƷ��ĿΪʣ������,ԭʼ����)
          If Instr(',4,5,6,7,', r_Bill.�շ����) = 0 Then
            If Instr(���_In, ':') = 0 Or ���_In Is Null Then
              v_׼������ := v_ʣ������;
            End If;
          Else
            --ҽ�������ջ�ʱ,���Ŀ���û�з���,���������ʵ��ǲ�������,����Ҫ�������Ϊ׼
            If Instr(���_In, ':') = 0 Or ���_In Is Null Then
              Select Nvl(Sum(Nvl(����, 1) * ʵ������), 0), Count(*)
              Into v_׼������, n_Count
              From ҩƷ�շ���¼
              Where NO = No_In And ���� In (9, 10, 25, 26) And Mod(��¼״̬, 3) = 1 And ����� Is Null And ����id = r_Bill.Id;
            End If;
          
            --��ʣ��������׼�������������������
            --1.���������õ������޶�Ӧ���շ���¼,��ʱʹ��ʣ������
            --2.��������,��ʱ�ѷ�ҩ����
            If v_׼������ = 0 Then
              If r_Bill.�շ���� = '4' Then
                If n_Count > 0 Then
                  v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����ѷ���,�����Ϻ����˷ѣ�';
                  Raise Err_Item;
                Else
                  v_׼������ := v_ʣ������;
                End If;
              Else
                v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����ѷ�ҩ,����ҩ�����˷ѣ�';
                Raise Err_Item;
              End If;
            End If;
          End If;
        
          --����סԺ���ü�¼
          If Nvl(n_����, 0) = 0 Then
            --����ʱ,ֱ�Ӹ�������,���Բ���黮��������
            --�ñ���Ŀ�ڼ�������
            Select Nvl(Max(Abs(ִ��״̬)), 0) + 1
            Into v_�˷Ѵ���
            From סԺ���ü�¼
            Where NO = No_In And ��¼���� = ��¼����_In And ��¼״̬ = 2 And ��� = r_Bill.��� And �����־ = 2;
          End If;
        
          --���=ʣ����*(׼����/ʣ����)
          v_Ӧ�ս�� := Round(v_ʣ��Ӧ�� * (v_׼������ / v_ʣ������), v_Dec);
          v_ʵ�ս�� := Round(v_ʣ��ʵ�� * (v_׼������ / v_ʣ������), v_Dec);
          v_ͳ���� := Round(v_ʣ��ͳ�� * (v_׼������ / v_ʣ������), v_Dec);
          If Nvl(n_����, 0) = 1 Then
            If Nvl(n_��������, 0) = 0 Then
              l_����.Extend;
              l_����(l_����.Count) := r_Bill.Id;
              n_����ֵ := 0;
            Else
              --��������
              --���۵�,�Ƚ���ص����ݴ������ڲ�����
              n_���� := 0;
              If r_Bill.���� > 1 Then
                --�������ҩ,���ڻ��տ϶��ǻ��յĸ���,�����Ǵ���.���,��Ҫ���׼�������Ƿ������ ��
                If Trunc(v_׼������ / r_Bill.����) <> (v_׼������ / r_Bill.����) Then
                  v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з���Ϊ��ҩ,�밴���������˷ѣ�';
                  Raise Err_Item;
                End If;
                n_���� := Trunc(v_׼������ / r_Bill.����);
                If Nvl(r_Bill.����, 0) - n_���� < 0 Then
                  v_׼������ := r_Bill.����;
                Else
                  v_׼������ := 0;
                End If;
              End If;
              Update סԺ���ü�¼
              Set ���� = ���� - n_����, ���� = ���� - v_׼������, Ӧ�ս�� = Nvl(Ӧ�ս��, 0) - v_Ӧ�ս��, ʵ�ս�� = Nvl(ʵ�ս��, 0) - v_ʵ�ս��,
                  �Ǽ�ʱ�� = v_Curdate, ͳ���� = Nvl(ͳ����, 0) - v_ͳ����
              Where ID = r_Bill.Id
              Returning Nvl(����, 0) * Nvl(����, 0) Into n_����ֵ;
            End If;
            If Nvl(n_����ֵ, 0) <= 0 Then
              l_����.Extend;
              l_����(l_����.Count) := r_Bill.Id;
            End If;
            If r_Bill.ҽ����� Is Not Null Then
              If Instr(',' || Nvl(v_ҽ��ids, '') || ',', ',' || r_Bill.ҽ����� || ',') = 0 Then
                v_ҽ��ids := Nvl(v_ҽ��ids, '') || ',' || r_Bill.ҽ�����;
              End If;
              --��¼����ҽ�����Ѷ�Ӧ��ҽ��ID(����������)
              If v_ҽ��id Is Null Then
                v_ҽ��id := r_Bill.ҽ�����;
              End If;
            End If;
          
          End If;
        
          If Nvl(n_����, 0) = 0 Then
            --����ʱ,ֱ�Ӹ�������,���Բ���黮��������
            --�����˷Ѽ�¼
            Insert Into סԺ���ü�¼
              (ID, NO, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, ��ҳid, ����id, ҽ�����, �����־, �ಡ�˵�, Ӥ����, ����, �Ա�, ����, ��ʶ��, ����, �ѱ�, ���˲���id,
               ���˿���id, �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������,
               ִ�в���id, ������, ִ����, ִ��״̬, ִ��ʱ��, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ������Ŀ��, ���մ���id, ͳ����, ���ձ���, ���ʵ�id, ժҪ, ��������, �Ƿ���,
               ����, ҽ��С��id)
              Select ���˷��ü�¼_Id.Nextval, NO, ��¼����, 2, ���, ��������, �۸񸸺�, ��ҳid, ����id, ҽ�����, �����־, �ಡ�˵�, Ӥ����, ����, �Ա�, ����, ��ʶ��,
                     ����, �ѱ�, ���˲���id, ���˿���id, �շ����, �շ�ϸĿid, ���㵥λ, Decode(Sign(v_׼������ - Nvl(����, 1) * ����), 0, ����, 1), ��ҩ����,
                     Decode(Sign(v_׼������ - Nvl(����, 1) * ����), 0, -1 * ����, -1 * v_׼������), �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���,
                     ��׼����, -1 * v_Ӧ�ս��, -1 * v_ʵ�ս��, ��������id, ������, ִ�в���id, ������, ִ����, -1 * v_�˷Ѵ���, ִ��ʱ��, ����Ա���_In,
                     ����Ա����_In, ����ʱ��, v_Curdate, ������Ŀ��, ���մ���id, -1 * v_ͳ����, ���ձ���, ���ʵ�id, ժҪ, ��������, �Ƿ���, ����, ҽ��С��id
              From סԺ���ü�¼
              Where ID = r_Bill.Id;
          
            --��¼����ҽ�����Ѷ�Ӧ��ҽ��ID(����������)
            If v_ҽ��id Is Null And r_Bill.ҽ����� Is Not Null Then
              v_ҽ��id := r_Bill.ҽ�����;
            End If;
          
            Update ����������Ŀ
            Set �������� = Nvl(��������, 0) - v_׼������
            Where ����id = r_Bill.����id And ��ҳid = r_Bill.��ҳid And ��Ŀid = r_Bill.�շ�ϸĿid And Nvl(ʹ������, 0) <> 0;
          
            --�������
            Update �������
            Set ������� = Nvl(�������, 0) - v_ʵ�ս��
            Where ����id = r_Bill.����id And ���� = 2 And ���� = 1;
            If Sql%RowCount = 0 Then
              Insert Into �������
                (����id, ����, ����, �������, Ԥ�����)
              Values
                (r_Bill.����id, 2, 1, -1 * v_ʵ�ս��, 0);
            End If;
          
            --����δ�����
            Update ����δ�����
            Set ��� = Nvl(���, 0) - v_ʵ�ս��
            Where ����id = r_Bill.����id And Nvl(��ҳid, 0) = Nvl(r_Bill.��ҳid, 0) And Nvl(���˲���id, 0) = Nvl(r_Bill.���˲���id, 0) And
                  Nvl(���˿���id, 0) = Nvl(r_Bill.���˿���id, 0) And Nvl(��������id, 0) = Nvl(r_Bill.��������id, 0) And
                  Nvl(ִ�в���id, 0) = Nvl(r_Bill.ִ�в���id, 0) And ������Ŀid + 0 = r_Bill.������Ŀid And ��Դ;�� + 0 = 2;
            If Sql%RowCount = 0 Then
              Insert Into ����δ�����
                (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
              Values
                (r_Bill.����id, r_Bill.��ҳid, r_Bill.���˲���id, r_Bill.���˿���id, r_Bill.��������id, r_Bill.ִ�в���id, r_Bill.������Ŀid, 2,
                 -1 * v_ʵ�ս��);
            End If;
          
            --���ԭ���ü�¼
            --ִ��״̬:ȫ������(׼����=ʣ����)���Ϊ0,���򱣳�ԭ״̬
            If Instr(',4,5,6,7,', r_Bill.�շ����) = 0 Then
              --һ�������ҩƷ�����ĵ���Ŀ,�����ڲ������ʵ����,ֻ������������������ʱ,�Ż���ֲ�������,����
              --ִ��״ֻ̬������:0.δִ��;1��ִ��;
              --������������˹����н���ִ��ǿ�Ƹ�Ϊ��2����ִ��,�����Ҫ�ڴ˴���Ϊ1��ִ��.δִ�еĲ���.
              Update סԺ���ü�¼
              Set ��¼״̬ = 3, ִ��״̬ = Decode(Sign(v_׼������ - v_ʣ������), 0, 0, Decode(ִ��״̬, 2, 1, ִ��״̬))
              Where ID = r_Bill.Id;
            Else
              Update סԺ���ü�¼
              Set ��¼״̬ = 3, ִ��״̬ = Decode(Sign(v_׼������ - v_ʣ������), 0, 0, ִ��״̬)
              Where ID = r_Bill.Id;
            End If;
          End If;
        End If;
      Else
        If v_��� Is Not Null Then
          v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����Ѿ���ȫִ��,�������ʣ�';
          Raise Err_Item;
        End If;
        --���:û�޶��к�,ԭʼ�����а����Ѿ���ȫִ�е�
      End If;
    End If;
  End Loop;

  --��������ҩID,����ҩƷ�Ƿ�����Һ��ҩ����
  If v_��ҩid Is Null And ��Һ��ҩ���_In = 1 Then
    For v_���� In (Select ID
                 From סԺ���ü�¼
                 Where NO = No_In And ��¼���� = ��¼����_In And ��¼״̬ In (0, 1, 3) And �շ���� In ('4', '5', '6', '7') And �����־ = 2 And
                       (Instr(',' || v_��� || ',', ',' || ��� || ',') > 0 Or v_��� Is Null)) Loop
      Begin
        Select Count(1)
        Into n_Count
        From ��Һ��ҩ���� A, ҩƷ�շ���¼ B
        Where a.�շ�id = b.Id And b.����id = v_����.Id And Instr(',8,9,10,21,24,25,26,', ',' || b.���� || ',') > 0;
      Exception
        When Others Then
          n_Count := 0;
      End;
      If n_Count <> 0 Then
        v_Err_Msg := '�����Ѿ�������Һ��ҩ���ĵĴ�����ҩƷ���޷�������ʣ�';
        Raise Err_Item;
      End If;
    End Loop;
  End If;

  n_�������� := 0;
  ---------------------------------------------------------------------------------
  --ҩƷ��ش���:��Ҫ�Ƕ����������Ч.(�����ǲ���)
  For v_���� In (Select ID, ���, �շ����
               From סԺ���ü�¼
               Where NO = No_In And ��¼���� = ��¼����_In And ��¼״̬ In (0, 1, 3) And �շ���� In ('4', '5', '6', '7') And �����־ = 2 And
                     (Instr(',' || v_��� || ',', ',' || ��� || ',') > 0 Or v_��� Is Null)) Loop
    --���ݷ���ID��������صĴ���
    v_׼������ := 0;
    If Instr(���_In, ':') > 0 Then
      v_Tmp := ',' || ���_In;
      v_Tmp := Substr(v_Tmp, Instr(v_Tmp, ',' || v_����.��� || ':') + Length(',' || v_����.��� || ':'));
      v_Tmp := Substr(v_Tmp, 1, Instr(v_Tmp || ',', ',') - 1);
      If Instr(v_Tmp, ':') > 0 Then
        v_Tmp := Substr(v_Tmp, 1, Instr(v_Tmp, ':') - 1);
      End If;
      v_׼������ := v_Tmp;
    End If;
    If v_׼������ <> 0 Then
      n_�������� := 1;
      n_Temp     := 0;
      --------------------------------------------------------------------------------------
      --����Ƿ񱸻���������,��������
      -- a.������ڴ���δ��˵����������Ҳ�������ʱ,ֱ����ԭ���Ļ����ϸ���������������
      -- b.������ڴ���δ��˵�������������ȫ����ʱ,ֱ��ɾ��
      -- c.��洦��:��ԭΪ����ⷿ�Ŀ�������;���ϲ��Ų�����
      -- d.����Ѿ�������,���ʱ�������������ⵥ�Ѿ����,��˾Ͱ����������ת,���ָ������ϲ�����
      n_����ⷿid := Null;
      n_��������id := Null;
      If v_����.�շ���� = '4' Then
        Begin
          Select 1, �ⷿid, ID
          Into n_��������, n_����ⷿid, n_��������id
          From ҩƷ�շ���¼
          Where ����id = v_����.Id And ������� Is Null And ���� = 21 And Rownum = 1;
        Exception
          When Others Then
            n_�������� := 0;
        End;
      Else
        n_�������� := 0;
      End If;
      --------------------------------------------------------------------------------------
      If v_��ҩid Is Not Null Then
        Open c_ҩƷ For
          Select /*+ rule*/
           a.Id, a.����, a.No, a.�ⷿid, a.ҩƷid, a.����, a.��ҩ��ʽ,
           Decode(a.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(a.����, 1) * Nvl(a.ʵ������, 0) As ����, a.���Ч��, a.Ч��, a.����, a.����, a.��������,
           a.����id
          From ҩƷ�շ���¼ A, Table(f_Str2list(v_��ҩid)) B, ��Һ��ҩ���� C
          Where a.No = No_In And a.���� In (9, 10, 25, 26) And Mod(a.��¼״̬, 3) = 1 And a.����� Is Null And a.����id = v_����.Id And
                a.Id = c.�շ�id And c.��¼id = b.Column_Value
          Order By ��������;
      Else
        Open c_ҩƷ For
          Select ID, ����, NO, �ⷿid, ҩƷid, ����, ��ҩ��ʽ, Decode(��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(����, 1) * Nvl(ʵ������, 0) As ����,
                 ���Ч��, Ч��, ����, ����, ��������, ����id
          From ҩƷ�շ���¼
          Where NO = No_In And ���� In (9, 10, 25, 26) And Mod(��¼״̬, 3) = 1 And ����� Is Null And ����id = v_����.Id
          Order By ��������;
      End If;
      Loop
        Fetch c_ҩƷ
          Into v_ҩƷ;
        Exit When c_ҩƷ%NotFound;
        n_Temp := v_ҩƷ.����;
        If v_׼������ >= n_Temp Then
          l_ҩƷ�շ�.Extend;
          l_ҩƷ�շ�(l_ҩƷ�շ�.Count) := v_ҩƷ.Id;
          If Nvl(n_��������id, 0) > 0 Then
            l_ҩƷ�շ�.Extend;
            l_ҩƷ�շ�(l_ҩƷ�շ�.Count) := n_��������id;
          End If;
          v_׼������ := v_׼������ - n_Temp;
        Else
          If v_����.�շ���� = '7' Then
            --��ǰ�е�����Ҫ��
            Update ҩƷ�շ���¼
            Set ���� = 1, ʵ������ = Decode(����, Null, 1, 0, 1, ����) * Nvl(ʵ������, 0) - v_׼������,
                ��д���� = Decode(����, Null, 1, 0, 1, ����) * Nvl(��д����, 0) - v_׼������,
                �ɱ���� =
                 (Decode(����, Null, 1, 0, 1, ����) * Nvl(ʵ������, 0) - v_׼������) * �ɱ���,
                ���۽�� =
                 (Decode(����, Null, 1, 0, 1, ����) * Nvl(ʵ������, 0) - v_׼������) * ���ۼ�,
                ��� = Round((Decode(����, Null, 1, 0, 1, ����) * Nvl(ʵ������, 0) - v_׼������) * ���ۼ� -
                            (Decode(����, Null, 1, 0, 1, ����) * Nvl(ʵ������, 0) - v_׼������) * �ɱ���, 5)
            Where ID = v_ҩƷ.Id;
          Else
            Update ҩƷ�շ���¼
            Set ʵ������ = Nvl(ʵ������, 0) - v_׼������, ��д���� = Nvl(��д����, 0) - v_׼������,
                �ɱ���� =
                 (Nvl(ʵ������, 0) - v_׼������) * �ɱ���,
                ���۽�� =
                 (Nvl(ʵ������, 0) - v_׼������) * ���ۼ�,
                ��� = Round((Nvl(ʵ������, 0) - v_׼������) * ���ۼ� - (Nvl(ʵ������, 0) - v_׼������) * �ɱ���, 5)
            Where ID = v_ҩƷ.Id;
          End If;
          --�����������ⵥ
          If Nvl(n_��������id, 0) <> 0 Then
            If v_����.�շ���� = '7' Then
              Update ҩƷ�շ���¼
              Set ���� = 1, ʵ������ = Decode(����, Null, 1, 0, 1, ����) * Nvl(ʵ������, 0) - v_׼������,
                  ��д���� = Decode(����, Null, 1, 0, 1, ����) * Nvl(ʵ������, 0) - v_׼������,
                  �ɱ���� =
                   (Decode(����, Null, 1, 0, 1, ����) * Nvl(ʵ������, 0) - v_׼������) * �ɱ���,
                  ���۽�� =
                   (Decode(����, Null, 1, 0, 1, ����) * Nvl(ʵ������, 0) - v_׼������) * ���ۼ�,
                  ��� = Round((Decode(����, Null, 1, 0, 1, ����) * Nvl(ʵ������, 0) - v_׼������) * ���ۼ� -
                              (Decode(����, Null, 1, 0, 1, ����) * Nvl(ʵ������, 0) - v_׼������) * �ɱ���, 5)
              Where ID = Nvl(n_��������id, 0);
            Else
              Update ҩƷ�շ���¼
              Set ʵ������ = Nvl(ʵ������, 0) - v_׼������, ��д���� = Nvl(ʵ������, 0) - v_׼������,
                  �ɱ���� =
                   (Nvl(ʵ������, 0) - v_׼������) * �ɱ���,
                  ���۽�� =
                   (Nvl(ʵ������, 0) - v_׼������) * ���ۼ�,
                  ��� = Round((Nvl(ʵ������, 0) - v_׼������) * ���ۼ� - (Nvl(ʵ������, 0) - v_׼������) * �ɱ���, 5)
              Where ID = Nvl(n_��������id, 0);
            End If;
          End If;
          n_Temp     := v_׼������;
          v_׼������ := 0;
        End If;
        If Nvl(n_��������, 0) = 1 Then
          n_�ⷿid := n_����ⷿid;
        Else
          n_�ⷿid := v_ҩƷ.�ⷿid;
        End If;
      
        If n_�ⷿid Is Not Null Then
          Update ҩƷ���
          Set �������� = Nvl(��������, 0) + n_Temp
          Where �ⷿid = n_�ⷿid And ҩƷid = v_ҩƷ.ҩƷid And Nvl(����, 0) = Nvl(v_ҩƷ.����, 0) And ���� = 1;
          If Sql%RowCount = 0 Then
            Insert Into ҩƷ���
              (�ⷿid, ҩƷid, ����, ����, Ч��, ��������, �ϴ�����, �ϴβ���, ���Ч��)
            Values
              (n_�ⷿid, v_ҩƷ.ҩƷid, 1, v_ҩƷ.����, v_ҩƷ.Ч��, n_Temp, v_ҩƷ.����, v_ҩƷ.����, v_ҩƷ.���Ч��);
          End If;
          Delete ҩƷ���
          Where �ⷿid = n_�ⷿid And ҩƷid = v_ҩƷ.ҩƷid And Nvl(����, 0) = Nvl(v_ҩƷ.����, 0) And ���� = 1 And Nvl(��������, 0) = 0 And
                Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0 And Nvl(ʵ�ʲ��, 0) = 0;
        End If;
      
        If Nvl(n_��������, 0) = 1 Then
          Update ҩƷ���
          Set �������� = Nvl(��������, 0) + n_Temp
          Where �ⷿid = v_ҩƷ.�ⷿid And ҩƷid = v_ҩƷ.ҩƷid And Nvl(����, 0) = Nvl(v_ҩƷ.����, 0) And ���� = 1;
          If Sql%RowCount = 0 Then
            Insert Into ҩƷ���
              (�ⷿid, ҩƷid, ����, ����, Ч��, ��������, �ϴ�����, �ϴβ���, ���Ч��)
            Values
              (v_ҩƷ.�ⷿid, v_ҩƷ.ҩƷid, 1, v_ҩƷ.����, v_ҩƷ.Ч��, n_Temp, v_ҩƷ.����, v_ҩƷ.����, v_ҩƷ.���Ч��);
          End If;
        End If;
      
        If v_׼������ = 0 Then
          Exit;
        End If;
      End Loop;
      --���������ĵ�,�����:��Ϊ������Ļ�,������ҩƷ�շ���¼�д���
      If Nvl(v_׼������, 0) <> 0 And Not (v_����.�շ���� = '4' And n_Temp = 0) Then
        --δ�������,��ʾ��ҩƷ�����Ѿ�ִ��.
        v_Err_Msg := 'Ҫ���ʵķ����д����ѷ���ҩƷ�����ģ����ѱ����������ʣ�������ǲ�����������ġ�';
        Raise Err_Item;
      End If;
    End If;
  End Loop;

  If n_�������� = 0 Then
    ------------------------------------------------------------------------------------------------------------------------
    --�ȴ���������
    For v_���� In (Select ID, ����, NO, �ⷿid, ҩƷid, ����, ��ҩ��ʽ, ����, ʵ������, ���Ч��, Ч��, ����, ����, ��������, ����id, ��Ʒ����, �ڲ�����
                 From ҩƷ�շ���¼
                 Where ���� = 21 And Mod(��¼״̬, 3) = 1 And ����� Is Null And
                       ����id In (Select ID
                                From סԺ���ü�¼
                                Where NO = No_In And ��¼���� = ��¼����_In And ��¼״̬ In (0, 1, 3) And �շ���� = '4' And �����־ = 2 And
                                      (Instr(',' || v_��� || ',', ',' || ��� || ',') > 0 Or v_��� Is Null))
                 Order By ҩƷid, �������� Desc) Loop
      --����ҩƷ���
      If v_����.�ⷿid Is Not Null Then
        Update ҩƷ���
        Set �������� = Nvl(��������, 0) + Decode(v_����.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(v_����.����, 1) * Nvl(v_����.ʵ������, 0)
        Where �ⷿid = v_����.�ⷿid And ҩƷid = v_����.ҩƷid And Nvl(����, 0) = Nvl(v_����.����, 0) And ���� = 1;
        If Sql%RowCount = 0 Then
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, Ч��, ��������, �ϴ�����, �ϴβ���, ���Ч��, ��Ʒ����, �ڲ�����)
          Values
            (v_����.�ⷿid, v_����.ҩƷid, 1, v_����.����, v_����.Ч��,
             Decode(v_����.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(v_����.����, 1) * Nvl(v_����.ʵ������, 0), v_����.����, v_����.����, v_����.���Ч��,
             v_����.��Ʒ����, v_����.�ڲ�����);
        End If;
        Delete ҩƷ���
        Where �ⷿid = v_����.�ⷿid And ҩƷid = v_����.ҩƷid And Nvl(����, 0) = Nvl(v_����.����, 0) And ���� = 1 And Nvl(��������, 0) = 0 And
              Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0 And Nvl(ʵ�ʲ��, 0) = 0;
      End If;
      l_����id.Extend;
      l_����id(l_����id.Count) := v_����.����id;
      l_ҩƷ�շ�.Extend;
      l_ҩƷ�շ�(l_ҩƷ�շ�.Count) := v_����.Id;
    End Loop;
  
    --ҩƷ�������
    Fetch c_Stock
      Into r_Stock;
    While c_Stock%Found Loop
    
      --����ҩƷ���
      If r_Stock.�ⷿid Is Not Null Then
      
        Select Decode(Count(Column_Value), Null, 0, 0, 0, 1)
        Into n_��������
        From Table(l_����id)
        Where Column_Value = r_Stock.����id;
      
        Update ҩƷ���
        Set �������� = Nvl(��������, 0) + Decode(r_Stock.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(r_Stock.����, 1) * Nvl(r_Stock.ʵ������, 0)
        Where �ⷿid = r_Stock.�ⷿid And ҩƷid = r_Stock.ҩƷid And Nvl(����, 0) = Nvl(r_Stock.����, 0) And ���� = 1;
        If Sql%RowCount = 0 Then
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, Ч��, ��������, �ϴ�����, �ϴβ���, ���Ч��)
          Values
            (r_Stock.�ⷿid, r_Stock.ҩƷid, 1, r_Stock.����, r_Stock.Ч��,
             Decode(r_Stock.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(r_Stock.����, 1) * Nvl(r_Stock.ʵ������, 0), r_Stock.����, r_Stock.����,
             r_Stock.���Ч��);
        End If;
        Delete ҩƷ���
        Where �ⷿid = r_Stock.�ⷿid And ҩƷid = r_Stock.ҩƷid And Nvl(����, 0) = Nvl(r_Stock.����, 0) And ���� = 1 And
              Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0 And Nvl(ʵ�ʲ��, 0) = 0;
      End If;
    
      --ɾ��ҩƷ�շ���¼(���ϲ����������:����� Is Null)
      --Delete From ҩƷ�շ���¼ Where ID = r_Stock.ID And ����� Is Null;
    
      l_ҩƷ�շ�.Extend;
      l_ҩƷ�շ�(l_ҩƷ�շ�.Count) := r_Stock.Id;
      Fetch c_Stock
        Into r_Stock;
    End Loop;
    Close c_Stock;
  
    --ɾ��ҩƷ�շ���¼
    Forall I In 1 .. l_ҩƷ�շ�.Count
      Delete From ҩƷ�շ���¼ Where ID = l_ҩƷ�շ�(I) And ����� Is Null;
    If Sql%RowCount <> l_ҩƷ�շ�.Count And l_ҩƷ�շ�.Count <> 0 Then
      v_Err_Msg := 'Ҫ���ʵķ����д����ѷ���ҩƷ�����ģ����ѱ����������ʣ�������ǲ�����������ġ�';
      Raise Err_Item;
    End If;
  Else
    --ɾ��ҩƷ�շ���¼
    Forall I In 1 .. l_ҩƷ�շ�.Count
      Delete From ҩƷ�շ���¼ Where ID = l_ҩƷ�շ�(I) And ����� Is Null;
  End If;
  --δ��ҩƷ��¼
  Delete From δ��ҩƷ��¼ A
  Where NO = No_In And ���� In (9, 10, 25, 26) And Not Exists
   (Select 1
         From ҩƷ�շ���¼
         Where ���� = a.���� And Nvl(�ⷿid, 0) = Nvl(a.�ⷿid, 0) And NO = No_In And Mod(��¼״̬, 3) = 1 And ����� Is Null);

  ---------------------------------------------------------------------------------
  --����ǻ���,ֱ��ɾ�����ü�¼(ҩƷ�����)
  n_Count := l_����.Count;
  --ɾ�����ۼ�¼
  Forall I In 1 .. l_����.Count
    Delete From סԺ���ü�¼ Where ID = l_����(I);

  --ɾ��֮����ͳһ�������
  If n_Count > 0 Then
    n_Count := 1;
    For r_Serial In c_Serial Loop
      If r_Serial.�۸񸸺� Is Null Then
        v_���� := n_Count;
      End If;
    
      Update סԺ���ü�¼
      Set ��� = n_Count, �۸񸸺� = Decode(�۸񸸺�, Null, Null, v_����)
      Where NO = No_In And ��¼���� = ��¼����_In And ��� = r_Serial.���;
    
      Update סԺ���ü�¼
      Set �������� = n_Count
      Where NO = No_In And ��¼���� = ��¼����_In And �������� = r_Serial.���;
    
      n_Count := n_Count + 1;
    End Loop;
  
  End If;
  --���ŵ���ȫ������ʱ��ɾ������ҽ������
  If v_��� Is Null And v_ҽ��id Is Not Null Then
    Select Nvl(Count(*), 0)
    Into n_Count
    From (Select ���, Sum(����) As ʣ������
           From (Select ��¼״̬, Nvl(�۸񸸺�, ���) As ���, Avg(Nvl(����, 1) * ����) As ����
                  From סԺ���ü�¼
                  Where NO = No_In And ��¼���� = 2 And ҽ����� + 0 = v_ҽ��id
                  Group By ��¼״̬, Nvl(�۸񸸺�, ���))
           Group By ���
           Having Nvl(Sum(����), 0) <> 0);
    If n_Count = 0 Then
      Delete From ����ҽ������ Where ҽ��id = v_ҽ��id And ��¼���� = 2 And NO = No_In;
    End If;
  End If;

  If v_ҽ��ids Is Not Null Then
    --ҽ������
    --����_In    Integer:=0, --0:����;1-סԺ
    --����_In    Integer:=1, --1-�շѵ�;2-���ʵ�
    --����_In    Integer:=0, --0:ɾ�����۵�;1-�շѻ����;2-�˷ѻ�����
    --No_In      ������ü�¼.No%Type,
    --ҽ��ids_In Varchar2 := Null
    v_ҽ��ids := Substr(v_ҽ��ids, 2);
    Zl_ҽ������_�Ʒ�״̬_Update(1, 2, 0, No_In, v_ҽ��ids);
  Else
    Zl_ҽ������_�Ʒ�״̬_Update(1, 2, 2, No_In);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_סԺ���ʼ�¼_Delete;
/
 

--93606:Ƚ����,2016-02-25,ҽ�����˲����˷�ʱ�ؽ���õ����ѵļ�¼״̬��Ϊ��2��Ӧ����1
Create Or Replace Procedure Zl_�����˷ѽ���_Modify
(
  ��������_In     Number,
  ����id_In       ������ü�¼.����id%Type,
  ����id_In       ����Ԥ����¼.����id%Type,
  ���㷽ʽ_In     Varchar2,
  ��Ԥ��_In       ����Ԥ����¼.��Ԥ��%Type := Null,
  �����id_In     ����Ԥ����¼.�����id%Type := Null,
  ����_In         ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In   ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In     ����Ԥ����¼.����˵��%Type := Null,
  �ɿ�_In         ����Ԥ����¼.�ɿ�%Type := Null,
  �Ҳ�_In         ����Ԥ����¼.�Ҳ�%Type := Null,
  �����_In     ������ü�¼.ʵ�ս��%Type := Null,
  ����˷�_In     Number := 0,
  ԭ����id_In     ����Ԥ����¼.����id%Type := Null,
  ʣ��תԤ��_In   Number := 0,
  ȱʡ���㷽ʽ_In ���㷽ʽ.����%Type := Null
) As
  ------------------------------------------------------------------------------------------------------------------------------
  --����:�շѽ���ʱ,�޸Ľ���������Ϣ
  --��������_In:
  --   0-ԭ����
  --      ԭ������һ��ȫ��,����У�Ա�־��Ϊ1,ҽ�����óɹ���,����Ϊ2,��ɺ���0
  --   1-��ͨ�˷ѷ�ʽ:
  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
  --   2.�������˷ѽ���:
  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
  --     �ڿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
  --   3-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.."
  --     ����֧Ʊ��_In:������
  --   4-���ѿ�����:
  --     �ٽ��㷽ʽ_IN:����һ��ˢ���ſ�,��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||."
  --     ����֧Ʊ��_In:������

  -- ��Ԥ��_In:����漰Ԥ����,���뱾�ε���Ԥ�� ������<0ʱ ��ʾ��Ԥ������ֵ;>0 ʱ:��ʾ��Ԥ����
  -- ʣ��תԤ��_In: 1��ʾ��ʣ���˿��ת��Ϊ��ֵ���;0��ʾ��Ԥ��
  -- �����_In:��������ʱ,����
  -- ����˷�_In:0-δ����˷�;1-�쳣����˷�;2-����˷�
  -- ԭ����ID_IN:ԭ����ʱ,����(���ԭ����δ����ʱ,�������һ�ν���Ϊ׼)
  ------------------------------------------------------------------------------------------------------------------------------
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_�������� Varchar2(500);
  v_��ǰ���� Varchar2(50);
  v_����     ����ҽ�ƿ���Ϣ.����%Type;
  n_���ѿ�id ���ѿ�Ŀ¼.Id%Type;
  n_�����id ����Ԥ����¼.���㿨���%Type;
  v_����     Varchar2(100);
  n_���ƿ�   �����ѽӿ�Ŀ¼.���ƿ�%Type;
  n_���     ���˿������¼.���%Type;
  n_Id       ���˿������¼.Id%Type;
  n_Ԥ��id   ����Ԥ����¼.Id%Type;
  v_���㷽ʽ ����Ԥ����¼.���㷽ʽ%Type;
  n_������ ����Ԥ����¼.��Ԥ��%Type;
  n_����ֵ   ��Ա�ɿ����.���%Type;
  n_Ԥ����� ����Ԥ����¼.��Ԥ��%Type;
  n_��Ԥ��   ����Ԥ����¼.��Ԥ��%Type;
  v_������� ����Ԥ����¼.�������%Type;
  v_����ժҪ ����Ԥ����¼.ժҪ%Type;
  v_����   ���㷽ʽ.����%Type;
  n_��¼״̬ ����Ԥ����¼.��¼״̬%Type;

  v_�˷ѽ��� ���㷽ʽ.����%Type;
  v_No       ����Ԥ����¼.No%Type;
  n_Dec      Number; --���С��λ�� 

  n_Count    Number;
  n_Havenull Number;
  l_Ԥ��id   t_Numlist := t_Numlist();
  n_ԭ����id ����Ԥ����¼.����id%Type;
  n_�ؽ�id   ����Ԥ����¼.����id%Type;
  n_����id   ����Ԥ����¼.����id%Type;
  n_������� ����Ԥ����¼.����id%Type;
  v_Msg      Varchar2(5000);
  Cursor c_Feedata Is
    Select Max(NO) As NO, Max(m.����id) As ����id, Max(m.�Ǽ�ʱ��) As �Ǽ�ʱ��, Max(m.����Ա���) As ����Ա���, Max(m.����Ա����) As ����Ա����,
           Sum(���ʽ��) As ������, Max(m.�ɿ���id) As �ɿ���id
    From ������ü�¼ M
    Where m.����id = ����id_In;
  r_Feedata c_Feedata%RowType;

  Cursor c_Balancedata Is
    Select ��¼����, NO, ��¼״̬, ����id, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������
    From ����Ԥ����¼
    Where ����id = ����id_In And ���㷽ʽ Is Null;
  r_Balancedata c_Balancedata%RowType;

  Procedure Zl_Square_Update
  (
    ԭ����id_In ����Ԥ����¼.����id%Type,
    �ֽ���id_In ����Ԥ����¼.����id%Type,
    �ɿ���id_In ����Ԥ����¼.�ɿ���id%Type,
    �˿�ʱ��_In ����Ԥ����¼.�տ�ʱ��%Type,
    �������_In ����Ԥ����¼.�������%Type,
    ��������_In Varchar2 := Null
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
    For v_У�� In (Select a.Id As Ԥ��id, c.���ѿ�id, c.������, c.�ӿڱ��, c.����, c.���, c.Id
                 From ����Ԥ����¼ A, ���˿�������� B, ���˿������¼ C
                 Where a.Id = b.Ԥ��id And b.������id = c.Id And a.��¼���� = 3 And a.��¼״̬ In (1, 3) And
                       Instr(Nvl(��������_In, '_LXH'), ',' || a.���㷽ʽ || ',') = 0 And a.����id = ԭ����id_In) Loop
    
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
          Select n_Ԥ��id, ��¼����, NO, 2, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �˿�ʱ��_In, �ɿλ, ��λ������, ��λ�ʺ�, r_Balancedata.����Ա���,
                 r_Balancedata.����Ա����, -1 * ��Ԥ��, �ֽ���id_In, �ɿ���id_In, Ԥ�����, �����id, Nvl(���㿨���, v_У��.�ӿڱ��), ����, ������ˮ��, ����˵��,
                 ������λ, 2, �������_In, Mod(��¼����, 10)
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
        Update ���ѿ�Ŀ¼ Set ��� = Nvl(���, 0) + v_У��.������ Where ID = Nvl(v_У��.���ѿ�id, 0);
      End If;
    
      Select ���˿������¼_Id.Nextval Into n_Id From Dual;
      Insert Into ���˿������¼
        (ID, �ӿڱ��, ���ѿ�id, ���, ��¼״̬, ���㷽ʽ, ������, ����, ������ˮ��, ����ʱ��, ��ע, �����־)
        Select n_Id, �ӿڱ��, ���ѿ�id, ���, n_��¼״̬, ���㷽ʽ, -1 * v_У��.������, ����, ������ˮ��, ����ʱ��, ��ע,
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

  Begin
    Select ���� Into v_�˷ѽ��� From ���㷽ʽ Where ���� = 1;
  Exception
    When Others Then
      v_�˷ѽ��� := '�ֽ�';
  End;

  Open c_Feedata;
  Fetch c_Feedata
    Into r_Feedata;

  If r_Feedata.No Is Null Then
    v_Err_Msg := 'δ�ҵ�ָ�����˷Ѽ�¼��';
    Raise Err_Item;
  End If;

  --���С��λ�� 
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_Dec From Dual;

  --0.��ʽ����
  Select Count(1), Max(Decode(���㷽ʽ, Null, 1, 0)), Max(�������)
  Into n_Count, n_Havenull, n_�������
  From ����Ԥ����¼
  Where ����id = ����id_In;

  If Nvl(n_Count, 0) = 0 Or Nvl(�����_In, 0) <> 0 Then
    --���ӽ��㷽ʽΪNULL�ļ�¼
    Begin
      Select ���� Into v_���� From ���㷽ʽ Where Nvl(����, 0) = 9;
    Exception
      When Others Then
        v_���� := '����';
    End;
  End If;

  --1.���ӽ��㷽ʽΪ�յĽ�������
  If Nvl(n_Havenull, 0) = 0 Then
    n_Count := 0;
    Begin
      n_������ := Round(Nvl(r_Feedata.������, 0), n_Dec);
    
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, У�Ա�־, ��������)
      Values
        (����Ԥ����¼_Id.Nextval, 3, Null, 2, Decode(����id_In, 0, Null, ����id_In), Null, r_Feedata.�Ǽ�ʱ��, r_Feedata.����Ա���,
         r_Feedata.����Ա����, n_������, ����id_In, r_Feedata.�ɿ���id, -1 * ����id_In, 1, 3);
      --����(�Ȼ��ܺ���������
      If n_������ <> Nvl(r_Feedata.������, 0) Then
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, У�Ա�־, ��������)
        Values
          (����Ԥ����¼_Id.Nextval, 3, Null, 1, Decode(����id_In, 0, Null, ����id_In), v_����, r_Feedata.�Ǽ�ʱ��, r_Feedata.����Ա���,
           r_Feedata.����Ա����, Nvl(r_Feedata.������, 0) - n_������, ����id_In, r_Feedata.�ɿ���id, -1 * ����id_In, 1, 3);
      End If;
      n_������� := -1 * ����id_In;
    Exception
      When Others Then
        n_Count := 1;
    End;
    If n_Count = 1 Then
      v_Err_Msg := 'δ�ҵ�ָ�����˷���ϸ����,�������ʧ�ܣ�';
      Raise Err_Item;
    End If;
  End If;

  Open c_Balancedata;
  Fetch c_Balancedata
    Into r_Balancedata;

  If ��������_In = 0 Then
    --0.ԭ����
    n_ԭ����id := ԭ����id_In;
    If Nvl(n_ԭ����id, 0) = 0 Then
      Select Max(����id)
      Into n_ԭ����id
      From ������ü�¼ A,
           (Select �Ǽ�ʱ�� From ������ü�¼ Where NO = r_Feedata.No And Mod(��¼����, 10) = 1 And ��¼״̬ In (1, 3)) B
      Where a.No = r_Feedata.No And Mod(a.��¼����, 10) = 1 And a.�Ǽ�ʱ�� = b.�Ǽ�ʱ��;
    End If;
    If Nvl(n_ԭ����id, 0) = 0 Then
      v_Err_Msg := 'δ�ҵ�ԭ��������,����ԭ���ˣ�';
      Raise Err_Item;
    End If;
  
    --1.�ȴ���Ԥ����
    n_������ := 0;
    For v_��Ԥ�� In (Select a.Id, Nvl(a.��Ԥ��, 0) As ���
                  From ����Ԥ����¼ A
                  Where Mod(��¼����, 10) = 1 And ����id = n_ԭ����id And Nvl(��Ԥ��, 0) <> 0
                  Order By �տ�ʱ�� Desc) Loop
    
      n_������ := n_������ + v_��Ԥ��.���;
    
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, У�Ա�־, �����id, ���㿨���,
         ����, ������ˮ��, ����˵��, �������, Ԥ�����, ��������)
        Select ����Ԥ����¼_Id.Nextval, 11, NO, ʵ��Ʊ��, ��¼״̬, ����id, Null, ժҪ, ���㷽ʽ, r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա���,
               r_Balancedata.����Ա����, -1 * v_��Ԥ��.���, r_Balancedata.����id, r_Balancedata.�ɿ���id, r_Balancedata.�������, 2, Null,
               Null, ����, ������ˮ��, ����˵��, Null, Ԥ�����, 3
        From ����Ԥ����¼
        Where ID = v_��Ԥ��.Id;
    
      Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - v_��Ԥ��.��� Where ����id = ����id_In And ���㷽ʽ Is Null;
      If Sql%NotFound Then
        v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[�շѽ��㴰��]�������շѣ�';
        Raise Err_Item;
      End If;
    End Loop;
    If Nvl(n_������, 0) <> 0 Then
    
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) + (-1 * n_������)
      Where ����id = ����id_In And ���� = 1 And ���� = 1
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ������� (����id, ����, Ԥ�����, ����) Values (����id_In, 1, (-1 * n_������), 1);
        n_����ֵ := (-1 * n_������);
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From �������
        Where ����id = r_Balancedata.����id And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
      End If;
    End If;
    --2.�������ѿ�����
    Zl_Square_Update(n_ԭ����id, r_Balancedata.����id, r_Balancedata.�ɿ���id, r_Balancedata.�տ�ʱ��, r_Balancedata.�������, v_��������);
    --3.�����������㲿��
    Insert Into ����Ԥ����¼
      (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����,
       ������ˮ��, ����˵��, ������λ, У�Ա�־, �������, ��������)
      Select ����Ԥ����¼_Id.Nextval, ��¼����, NO, 2, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա���,
             r_Balancedata.����Ա����, -1 * ��Ԥ��, r_Balancedata.����id, r_Balancedata.�ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��,
             ������λ,
             Case
               When Nvl(a.�����id, 0) <> 0 Then
                1
               When Nvl(a.���㿨���, 0) <> 0 Then
                1
               When Nvl(q.Ԥ��id, 0) <> 0 Then
                1
               When Nvl(j.����, '-') <> '-' Then
               --ҽ��
                1
               Else
                2
             End As У�Ա�־, r_Balancedata.�������, 3
      From ����Ԥ����¼ A, (Select ���� From ���㷽ʽ Where ���� In (3, 4)) J,
           (Select m.Id As Ԥ��id
             From ����Ԥ����¼ M, һ��ͨĿ¼ C
             Where m.����id = n_ԭ����id And m.���㷽ʽ = c.���㷽ʽ And m.��¼���� = 3 And m.��¼״̬ In (1, 3)) Q
      Where Mod(a.��¼����, 10) <> 1 And a.��¼״̬ In (1, 3) And a.���㷽ʽ = j.����(+) And a.���㷽ʽ Is Not Null And a.����id = n_ԭ����id And
            a.Id = q.Ԥ��id(+) And (Not Exists (Select 1 From ���˿�������� Where a.Id = Ԥ��id) Or Nvl(���㿨���, 0) = 0);
  
    --���½��㷽ʽΪNULL �ļ�¼
    Select Sum(��Ԥ��) Into n_����ֵ From ����Ԥ����¼ Where ����id = r_Balancedata.����id And ���㷽ʽ Is Not Null;
    Select Sum(���ʽ��)
    Into n_������
    From ������ü�¼
    Where ����id = r_Balancedata.����id And Mod(��¼����, 10) = 1;
  
    Update ����Ԥ����¼
    Set ��Ԥ�� = Nvl(n_������, 0) - Nvl(n_����ֵ, 0)
    Where ����id = ����id_In And ���㷽ʽ Is Null;
    If Sql%NotFound Then
      v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[�շѽ��㴰��]�������շѣ�';
      Raise Err_Item;
    End If;
  
  End If;

  n_�ؽ�id := 0;
  If ��������_In <> 0 Then
    --����ȫ��ʱ,����Ƿ�����������շ����ݵ�
    Begin
      Select ����id Into n_�ؽ�id From ����Ԥ����¼ Where ������� = n_������� And ����id <> ����id_In And Rownum < 2;
    Exception
      When Others Then
        n_�ؽ�id := 0;
    End;
  End If;

  --��Ҫ���������
  If Nvl(�����_In, 0) <> 0 Then
    --���ѷ������յĽ����¼��
    n_����id   := ����id_In;
    n_��¼״̬ := 2;
    If Nvl(n_�ؽ�id, 0) <> 0 Then
      n_����id   := n_�ؽ�id;
      n_��¼״̬ := 1;
    End If;
  
    Update ����Ԥ����¼
    Set ��Ԥ�� = Nvl(��Ԥ��, 0) + Nvl(�����_In, 0)
    Where ����id = n_����id And ���㷽ʽ = v_����;
    If Sql%NotFound Then
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, У�Ա�־, �����id, ���㿨���, ����,
         ������ˮ��, ����˵��, �������, ��������)
      Values
        (����Ԥ����¼_Id.Nextval, 3, Null, n_��¼״̬, r_Balancedata.����id, Null, Null, v_����, r_Balancedata.�տ�ʱ��,
         r_Balancedata.����Ա���, r_Balancedata.����Ա����, �����_In, n_����id, r_Balancedata.�ɿ���id, r_Balancedata.�������, 2, Null,
         Null, ����_In, ������ˮ��_In, ����˵��_In, Null, 3);
    End If;
  
    Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - Nvl(�����_In, 0) Where ����id = n_����id And ���㷽ʽ Is Null;
    If Sql%NotFound Then
      v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[�շѽ��㴰��]�������շѣ�';
      Raise Err_Item;
    End If;
  End If;

  --Ԥ�����:����ǳ�Ԥ��,��Ҫ�ȴ����Ԥ����
  If Nvl(��Ԥ��_In, 0) <> 0 Then
    If Nvl(r_Balancedata.����id, 0) = 0 Then
      v_Err_Msg := '����ȷ��������Ϣ,����ʹ��Ԥ������㣡';
      Raise Err_Item;
    End If;
  
    n_Ԥ����� := ��Ԥ��_In;
    If n_Ԥ����� < 0 And Nvl(ʣ��תԤ��_In, 0) = 1 Then
    
      --     �ڳ�Ԥ��_In:����漰Ԥ����,���뱾�ε���Ԥ�� ������<0ʱ ��ʾ��Ԥ������ֵ;>0 ʱ:��ʾ��Ԥ����
      --     ��ʣ��תԤ��_In: 1��ʾ��ʣ���˿��ת��Ϊ��ֵ���;0��ʾ��Ԥ��
    
      --1.�����ɳ�ֵԤ��:
      v_No := Nextno(11);
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ���, ����id, �ɿ���id, �������, У�Ա�־, �����id, ���㿨���, ����,
         ������ˮ��, ����˵��, �������, Ԥ�����, ��������)
      Values
        (����Ԥ����¼_Id.Nextval, 1, v_No, 1, r_Balancedata.����id, Null, '�˷�����Ԥ��', v_�˷ѽ���, r_Balancedata.�տ�ʱ��,
         r_Balancedata.����Ա���, r_Balancedata.����Ա����, -1 * n_Ԥ�����, r_Balancedata.����id, r_Balancedata.�ɿ���id,
         r_Balancedata.�������, 0, Null, Null, ����_In, ������ˮ��_In, ����˵��_In, Null, 1, Null);
    
      --���²������
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) + n_Ԥ�����
      Where ����id = ����id_In And ���� = 1 And ���� = 1
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ������� (����id, ����, Ԥ�����, ����) Values (����id_In, 1, n_Ԥ�����, 1);
        n_����ֵ := n_Ԥ�����;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From �������
        Where ����id = r_Balancedata.����id And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
      End If;
    
      --2.�����˷Ѽ�¼
      If Nvl(n_�ؽ�id, 0) <> 0 Then
        Select Sum(��Ԥ��)
        Into n_����ֵ
        From ����Ԥ����¼
        Where ����id = ����id_In And ���㷽ʽ Is Null And Nvl(��Ԥ��, 0) <> 0;
        If Nvl(n_����ֵ, 0) <> 0 Then
          Insert Into ����Ԥ����¼
            (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, У�Ա�־, �����id, ���㿨���,
             ����, ������ˮ��, ����˵��, �������, ��������)
          Values
            (����Ԥ����¼_Id.Nextval, 3, Null, 2, r_Balancedata.����id, Null, v_����ժҪ, v_�˷ѽ���, r_Balancedata.�տ�ʱ��,
             r_Balancedata.����Ա���, r_Balancedata.����Ա����, n_����ֵ, ����id_In, r_Balancedata.�ɿ���id, r_Balancedata.�������, 2, Null,
             Null, ����_In, ������ˮ��_In, ����˵��_In, Null, 3);
        
          Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - Nvl(n_����ֵ, 0) Where ����id = ����id_In And ���㷽ʽ Is Null;
        End If;
        n_������ := n_������ - Nvl(n_����ֵ, 0);
      
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, У�Ա�־, �����id, ���㿨���, ����,
           ������ˮ��, ����˵��, �������, ��������)
        Values
          (����Ԥ����¼_Id.Nextval, 3, Null, 1, r_Balancedata.����id, Null, v_����ժҪ, v_�˷ѽ���, r_Balancedata.�տ�ʱ��,
           r_Balancedata.����Ա���, r_Balancedata.����Ա����, n_������, n_�ؽ�id, r_Balancedata.�ɿ���id, r_Balancedata.�������, 2, Null,
           Null, ����_In, ������ˮ��_In, ����˵��_In, Null, 3);
        Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� + n_Ԥ����� Where ����id = ����id_In And ���㷽ʽ Is Null;
        If Sql%NotFound Then
          v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[�շѽ��㴰��]�������շѣ�';
          Raise Err_Item;
        End If;
      Else
      
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, У�Ա�־, �����id, ���㿨���, ����,
           ������ˮ��, ����˵��, �������, ��������)
        Values
          (����Ԥ����¼_Id.Nextval, 3, Null, 2, r_Balancedata.����id, Null, v_����ժҪ, v_���㷽ʽ, r_Balancedata.�տ�ʱ��,
           r_Balancedata.����Ա���, r_Balancedata.����Ա����, -1 * n_Ԥ�����, r_Balancedata.����id, r_Balancedata.�ɿ���id,
           r_Balancedata.�������, 2, Null, Null, ����_In, ������ˮ��_In, ����˵��_In, Null, 3);
        Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� + n_Ԥ����� Where ����id = ����id_In And ���㷽ʽ Is Null;
        If Sql%NotFound Then
          v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[�շѽ��㴰��]�������շѣ�';
          Raise Err_Item;
        End If;
      End If;
    End If;
  
    If Nvl(n_Ԥ�����, 0) < 0 And Nvl(ʣ��תԤ��_In, 0) = 0 Then
    
      n_ԭ����id := ԭ����id_In;
      If Nvl(n_ԭ����id, 0) = 0 Then
        Select Max(b.����id)
        Into n_ԭ����id
        From ������ü�¼ A, ������ü�¼ B
        Where a.����id = ����id_In And a.No = b.No And b.��¼���� = 1 And b.��¼״̬ In (1, 3);
      End If;
    
      If Nvl(n_ԭ����id, 0) = 0 Then
        v_Err_Msg := 'δ�ҵ�ԭ��������,����ԭ���ˣ�';
        Raise Err_Item;
      End If;
      If Nvl(n_�ؽ�id, 0) <> 0 Then
        Select Sum(��Ԥ��)
        Into n_����ֵ
        From ����Ԥ����¼
        Where ����id = ����id_In And ���㷽ʽ Is Null And Nvl(��Ԥ��, 0) <> 0;
        If Nvl(n_����ֵ, 0) <> 0 Then
          For v_��Ԥ�� In (Select a.Id, Nvl(a.��Ԥ��, 0) As ���
                        From ����Ԥ����¼ A
                        Where Mod(��¼����, 10) = 1 And ����id = n_ԭ����id And Nvl(��Ԥ��, 0) <> 0
                        Order By �տ�ʱ�� Desc) Loop
            Insert Into ����Ԥ����¼
              (ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, У�Ա�־, �����id,
               ���㿨���, ����, ������ˮ��, ����˵��, �������, Ԥ�����, ��������)
              Select ����Ԥ����¼_Id.Nextval, 11, NO, ʵ��Ʊ��, ��¼״̬, ����id, Null, ժҪ, ���㷽ʽ, r_Balancedata.�տ�ʱ��,
                     r_Balancedata.����Ա���, r_Balancedata.����Ա����, Nvl(n_����ֵ, 0), r_Balancedata.����id, r_Balancedata.�ɿ���id,
                     r_Balancedata.�������, 2, Null, Null, ����, ������ˮ��, ����˵��, Null, Ԥ�����, 3
              From ����Ԥ����¼
              Where ID = v_��Ԥ��.Id;
            Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - Nvl(n_����ֵ, 0) Where ����id = ����id_In And ���㷽ʽ Is Null;
            If Sql%NotFound Then
              v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[�շѽ��㴰��]�������շѣ�';
              Raise Err_Item;
            End If;
            n_Ԥ����� := n_Ԥ����� - Nvl(n_����ֵ, 0);
          End Loop;
        End If;
        n_����ֵ := 0;
        --2.��Ԥ����
        For v_��Ԥ�� In (Select Max(a.Id) As ID, Max(a.�տ�ʱ��) As �տ�ʱ��, Sum(Nvl(a.��Ԥ��, 0)) As ���
                      From ����Ԥ����¼ A,
                           (Select Distinct a.����id
                             From ������ü�¼ A, ������ü�¼ B
                             Where a.No = b.No And Mod(b.��¼����, 10) = 1 And b.����id = n_ԭ����id) B
                      Where a.����id = b.����id And Mod(a.��¼����, 10) = 1 And Nvl(a.Ԥ�����, 0) = 1 And a.����id <> ����id_In
                      Group By NO
                      Order By �տ�ʱ�� Desc) Loop
        
          If v_��Ԥ��.��� + n_Ԥ����� < 0 Then
            n_������ := v_��Ԥ��.���;
            n_Ԥ����� := n_Ԥ����� + v_��Ԥ��.���;
          Else
            n_������ := n_Ԥ�����;
            n_Ԥ����� := 0;
          End If;
        
          Insert Into ����Ԥ����¼
            (ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, У�Ա�־, �����id,
             ���㿨���, ����, ������ˮ��, ����˵��, �������, Ԥ�����, ��������)
            Select ����Ԥ����¼_Id.Nextval, 11, NO, ʵ��Ʊ��, ��¼״̬, ����id, Null, ժҪ, ���㷽ʽ, r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա���,
                   r_Balancedata.����Ա����, n_������, n_�ؽ�id, r_Balancedata.�ɿ���id, r_Balancedata.�������, 2, Null, Null, ����,
                   ������ˮ��, ����˵��, Null, Ԥ�����, 3
            From ����Ԥ����¼
            Where ID = v_��Ԥ��.Id;
          Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - n_������ Where ����id = n_�ؽ�id And ���㷽ʽ Is Null;
          n_����ֵ := 1;
          If Sql%NotFound Then
            v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[�շѽ��㴰��]�������շѣ�';
            Raise Err_Item;
          End If;
          If n_Ԥ����� = 0 Then
            Exit;
          End If;
        End Loop;
      Else
        --��Ԥ����
        n_����ֵ   := 0;
        n_Ԥ����� := -1 * n_Ԥ�����;
      
        For v_��Ԥ�� In (Select Max(a.Id) As ID, Max(a.�տ�ʱ��) As �տ�ʱ��, Sum(Nvl(a.��Ԥ��, 0)) As ���
                      From ����Ԥ����¼ A,
                           (Select Distinct a.����id
                             From ������ü�¼ A, ������ü�¼ B
                             Where a.No = b.No And Mod(b.��¼����, 10) = 1 And b.����id = n_ԭ����id) B
                      Where a.����id = b.����id And Mod(a.��¼����, 10) = 1 And Nvl(a.Ԥ�����, 0) = 1
                      Group By NO
                      Having Sum(Nvl(a.��Ԥ��, 0)) > 0
                      Order By �տ�ʱ�� Desc) Loop
        
          If v_��Ԥ��.��� - n_Ԥ����� < 0 Then
            n_������ := -1 * v_��Ԥ��.���;
            n_Ԥ����� := n_Ԥ����� - v_��Ԥ��.���;
          Else
            n_������ := -1 * n_Ԥ�����;
            n_Ԥ����� := 0;
          End If;
        
          Insert Into ����Ԥ����¼
            (ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, У�Ա�־, �����id,
             ���㿨���, ����, ������ˮ��, ����˵��, �������, Ԥ�����, ��������)
            Select ����Ԥ����¼_Id.Nextval, 11, NO, ʵ��Ʊ��, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա���,
                   r_Balancedata.����Ա����, n_������, r_Balancedata.����id, r_Balancedata.�ɿ���id, r_Balancedata.�������, 2, Null,
                   Null, ����, ������ˮ��, ����˵��, Null, Ԥ�����, 3
            From ����Ԥ����¼
            Where ID = v_��Ԥ��.Id;
        
          Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - n_������ Where ����id = ����id_In And ���㷽ʽ Is Null;
          n_����ֵ := 1;
          If Sql%NotFound Then
            v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[�շѽ��㴰��]�������շѣ�';
            Raise Err_Item;
          End If;
          If n_Ԥ����� = 0 Then
            Exit;
          End If;
        End Loop;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        v_Err_Msg := 'δ�ҵ�ԭʼ�ĳ�Ԥ����¼,���ܻ���Ԥ���';
        Raise Err_Item;
      End If;
    
      If Nvl(n_Ԥ�����, 0) <> 0 Then
        v_Err_Msg := '��ǰ��Ԥ���������շѽ����еĳ�Ԥ����,���ܻ���Ԥ���';
        Raise Err_Item;
      End If;
    
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) + (-1 * ��Ԥ��_In)
      Where ����id = ����id_In And ���� = 1 And ���� = 1
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ������� (����id, ����, Ԥ�����, ����) Values (����id_In, 1, (-1 * ��Ԥ��_In), 1);
        n_����ֵ := (-1 * ��Ԥ��_In);
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From �������
        Where ����id = r_Balancedata.����id And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
      End If;
    
    End If;
  
    n_Ԥ����� := ��Ԥ��_In;
    If Nvl(n_Ԥ�����, 0) > 0 Then
      --��Ԥ����
      --���������
      Begin
        Select Nvl(Ԥ�����, 0) - Nvl(�������, 0)
        Into n_Ԥ�����
        From �������
        Where ����id = ����id_In And Nvl(����, 0) = 1 And ���� = 1;
      Exception
        When Others Then
          n_Ԥ����� := 0;
      End;
      If n_Ԥ����� < ��Ԥ��_In Then
        v_Err_Msg := '���˵ĵ�ǰԤ�����Ϊ ' || LTrim(To_Char(n_Ԥ�����, '9999999990.00')) || '��С�ڱ���֧����� ' ||
                     LTrim(To_Char(��Ԥ��_In, '9999999990.00')) || ' ��';
        Raise Err_Item;
      End If;
    
      n_Ԥ����� := ��Ԥ��_In;
      n_����id   := ����id_In;
      If Nvl(n_�ؽ�id, 0) <> 0 Then
        n_����id := n_�ؽ�id;
        --�ܵĳ�Ԥ����� = ���γ�Ԥ����� + δ�������
        --��Ϊ�ں���Ὣδ�������ȫ����ΪԤ����
        Select Sum(��Ԥ��)
        Into n_����ֵ
        From ����Ԥ����¼
        Where ����id = ����id_In And ���㷽ʽ Is Null And Nvl(��Ԥ��, 0) <> 0;
        If Nvl(n_����ֵ, 0) <> 0 Then
          n_Ԥ����� := n_Ԥ����� - Nvl(n_����ֵ, 0);
        End If;
      End If;
    
      For c_��Ԥ�� In (Select *
                    From (Select a.Id, a.��¼״̬, a.No, Nvl(a.���, 0) As ���
                           From ����Ԥ����¼ A,
                                (Select NO, Sum(Nvl(a.���, 0)) As ���
                                  From ����Ԥ����¼ A
                                  Where a.����id Is Null And Nvl(a.���, 0) <> 0 And a.����id = ����id_In And Ԥ����� = 1
                                  Group By NO
                                  Having Sum(Nvl(a.���, 0)) <> 0) B
                           Where a.����id Is Null And Nvl(a.���, 0) <> 0 And a.No = b.No And a.����id = ����id_In And a.Ԥ����� = 1
                           Union All
                           Select 0 As ID, ��¼״̬, NO, Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) As ���
                           From ����Ԥ����¼
                           Where ��¼���� In (1, 11) And ����id Is Not Null And Nvl(���, 0) <> Nvl(��Ԥ��, 0) And ����id = ����id_In And
                                 Ԥ����� = 1 Having Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0
                           Group By ��¼״̬, NO)
                    Order By ID, NO) Loop
      
        If c_��Ԥ��.��� - n_Ԥ����� < 0 Then
          n_��Ԥ�� := c_��Ԥ��.���;
        Else
          n_��Ԥ�� := n_Ԥ�����;
        End If;
      
        If c_��Ԥ��.Id <> 0 Then
          --��һ�γ�Ԥ��(����һ�α��Ͻ���ID,��Ԥ�����Ϊ0)
          Update ����Ԥ����¼
          Set ��Ԥ�� = 0, ����id = n_����id, ������� = n_�������, �������� = 3
          Where ID = c_��Ԥ��.Id;
        End If;
        --���ϴ�ʣ���
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
           ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, ��������)
          Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�,
                 r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա����, r_Balancedata.����Ա���, n_��Ԥ��, n_����id, r_Balancedata.�ɿ���id, Ԥ�����,
                 �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, n_�������, 3
          From ����Ԥ����¼
          Where NO = c_��Ԥ��.No And ��¼״̬ = c_��Ԥ��.��¼״̬ And ��¼���� In (1, 11) And Rownum = 1;
      
        --��������(���㷽ʽΪNULL��)
        Update ����Ԥ����¼
        Set ��Ԥ�� = ��Ԥ�� - n_��Ԥ��
        Where ����id = n_����id And ���㷽ʽ Is Null
        Returning Nvl(��Ԥ��, 0) Into n_����ֵ;
      
        --����Ƿ��Ѿ�������
        If c_��Ԥ��.��� < n_Ԥ����� Then
          n_Ԥ����� := n_Ԥ����� - c_��Ԥ��.���;
        Else
          n_Ԥ����� := 0;
        End If;
        If n_Ԥ����� = 0 Then
          Exit;
        End If;
      End Loop;
    
      --������Ƿ��㹻
      If Abs(n_Ԥ�����) > 0 Then
        v_Err_Msg := '���˵ĵ�ǰԤ�����С�ڱ���֧����� ' || LTrim(To_Char(��Ԥ��_In, '9999999990.00')) || ' ��';
        Raise Err_Item;
      End If;
    
      If Nvl(n_�ؽ�id, 0) <> 0 Then
        Select Sum(��Ԥ��)
        Into n_����ֵ
        From ����Ԥ����¼
        Where ����id = ����id_In And ���㷽ʽ Is Null And Nvl(��Ԥ��, 0) <> 0;
        If Nvl(n_����ֵ, 0) <> 0 Then
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
             ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�,
                   r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա����, r_Balancedata.����Ա���, n_����ֵ, ����id_In, r_Balancedata.�ɿ���id,
                   Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, n_�������, 3
            From ����Ԥ����¼
            Where ����id = n_�ؽ�id And ��¼���� In (1, 11) And Rownum = 1;
        
          Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - Nvl(n_����ֵ, 0) Where ����id = ����id_In And ���㷽ʽ Is Null;
          If Sql%NotFound Then
            v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[�շѽ��㴰��]�������շѣ�';
            Raise Err_Item;
          End If;
        End If;
      
      End If;
      --���²���Ԥ�����
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) - ��Ԥ��_In
      Where ����id = ����id_In And ���� = 1 And ���� = 1
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ������� (����id, ����, Ԥ�����, ����) Values (����id_In, 1, -1 * ��Ԥ��_In, 1);
        n_����ֵ := -1 * ��Ԥ��_In;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From ������� Where ����id = ����id_In And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
      End If;
    End If;
  End If;

  If ��������_In = 1 Then
    --   1-��ͨ�˷ѷ�ʽ:
    --�����շѽ��� :��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.."
    v_�������� := ���㷽ʽ_In || '||';
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
    
      v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
    
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
    
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      v_������� := LTrim(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_����ժҪ := LTrim(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
    
      --���жϡ�������Ƿ�Ϊ�㣬�п����Ѿ����꣬����ʱ���㷽ʽΪ�յ��ؽ�ͳ�����¼�ĳ�Ԥ��֮��Ϊ��
      If v_���㷽ʽ Is Null Then
        v_���㷽ʽ := ȱʡ���㷽ʽ_In;
      End If;
      --If Nvl(n_������, 0) <> 0 Then
      n_������ := Nvl(n_������, 0);
      If Nvl(n_�ؽ�id, 0) <> 0 Then
        --�϶����տ�
        --1.�Ȱ����ַ�ʽȫ��
        --2.�ٰ����ַ�ʽ�տ�
        --3:1+2=�����˿�
        --1.�Ƚ��˷ѵ�ȫ�����ϵ�
        Select Sum(��Ԥ��)
        Into n_����ֵ
        From ����Ԥ����¼
        Where ����id = ����id_In And ���㷽ʽ Is Null And Nvl(��Ԥ��, 0) <> 0;
      
        If Nvl(n_����ֵ, 0) <> 0 Then
        
          Insert Into ����Ԥ����¼
            (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, У�Ա�־, �����id, ���㿨���,
             ����, ������ˮ��, ����˵��, �������, ��������)
          Values
            (����Ԥ����¼_Id.Nextval, 3, Null, 2, r_Balancedata.����id, Null, v_����ժҪ, v_���㷽ʽ, r_Balancedata.�տ�ʱ��,
             r_Balancedata.����Ա���, r_Balancedata.����Ա����, n_����ֵ, ����id_In, r_Balancedata.�ɿ���id, r_Balancedata.�������, 2, Null,
             Null, ����_In, ������ˮ��_In, ����˵��_In, Null, 3);
          Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - Nvl(n_����ֵ, 0) Where ����id = ����id_In And ���㷽ʽ Is Null;
        End If;
        n_������ := n_������ - Nvl(n_����ֵ, 0);
        --2.�˿�
        If Nvl(n_������, 0) <> 0 Then
          Insert Into ����Ԥ����¼
            (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, У�Ա�־, �����id, ���㿨���,
             ����, ������ˮ��, ����˵��, �������, ��������)
          Values
            (����Ԥ����¼_Id.Nextval, 3, Null, 1, r_Balancedata.����id, Null, v_����ժҪ, v_���㷽ʽ, r_Balancedata.�տ�ʱ��,
             r_Balancedata.����Ա���, r_Balancedata.����Ա����, n_������, n_�ؽ�id, r_Balancedata.�ɿ���id, r_Balancedata.�������, 2, Null,
             Null, ����_In, ������ˮ��_In, ����˵��_In, Null, 3);
          Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - n_������ Where ����id = n_�ؽ�id And ���㷽ʽ Is Null;
        End If;
      Else
        --:>�˿�
        If Nvl(n_������, 0) <> 0 Then
          Insert Into ����Ԥ����¼
            (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, У�Ա�־, �����id, ���㿨���,
             ����, ������ˮ��, ����˵��, �������, ��������)
          Values
            (����Ԥ����¼_Id.Nextval, 3, Null, 2, r_Balancedata.����id, Null, v_����ժҪ, v_���㷽ʽ, r_Balancedata.�տ�ʱ��,
             r_Balancedata.����Ա���, r_Balancedata.����Ա����, n_������, r_Balancedata.����id, r_Balancedata.�ɿ���id,
             r_Balancedata.�������, 2, Null, Null, ����_In, ������ˮ��_In, ����˵��_In, Null, 3);
        
          Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - n_������ Where ����id = ����id_In And ���㷽ʽ Is Null;
          If Sql%NotFound Then
            v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[�շѽ��㴰��]�������շѣ�';
            Raise Err_Item;
          End If;
        End If;
      End If;
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  
  End If;

  If ��������_In = 2 Then
    --   2.�������˷ѽ���:
  
    v_��ǰ���� := ���㷽ʽ_In;
    v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
    v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
    n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
    v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
    v_������� := LTrim(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
    v_����ժҪ := LTrim(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
  
    If Nvl(n_������, 0) <> 0 Then
    
      If Nvl(n_�ؽ�id, 0) <> 0 Then
        --1.�Ȱ����ַ�ʽȫ��
        --2.�ٰ����ַ�ʽ�տ�
        --3:1+2=�����˿�
        --1.�Ƚ��˷ѵ�ȫ�����ϵ�
        Select Sum(��Ԥ��)
        Into n_����ֵ
        From ����Ԥ����¼
        Where ����id = ����id_In And ���㷽ʽ Is Null And Nvl(��Ԥ��, 0) <> 0;
        If Nvl(n_����ֵ, 0) <> 0 Then
          Insert Into ����Ԥ����¼
            (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, У�Ա�־, �����id, ���㿨���,
             ����, ������ˮ��, ����˵��, �������, ��������)
          Values
            (����Ԥ����¼_Id.Nextval, 3, Null, 2, r_Balancedata.����id, Null, v_����ժҪ, v_���㷽ʽ, r_Balancedata.�տ�ʱ��,
             r_Balancedata.����Ա���, r_Balancedata.����Ա����, n_����ֵ, ����id_In, r_Balancedata.�ɿ���id, r_Balancedata.�������, 2,
             �����id_In, Null, ����_In, ������ˮ��_In, ����˵��_In, Null, 3);
          Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - Nvl(n_����ֵ, 0) Where ����id = ����id_In And ���㷽ʽ Is Null;
        End If;
        n_������ := n_������ - Nvl(n_����ֵ, 0);
      
        --2.�˿�
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, У�Ա�־, �����id, ���㿨���, ����,
           ������ˮ��, ����˵��, �������, ��������)
        Values
          (����Ԥ����¼_Id.Nextval, 3, Null, 1, r_Balancedata.����id, Null, v_����ժҪ, v_���㷽ʽ, r_Balancedata.�տ�ʱ��,
           r_Balancedata.����Ա���, r_Balancedata.����Ա����, n_������, n_�ؽ�id, r_Balancedata.�ɿ���id, r_Balancedata.�������, 2,
           �����id_In, Null, ����_In, ������ˮ��_In, ����˵��_In, Null, 3);
        Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - n_������ Where ����id = n_�ؽ�id And ���㷽ʽ Is Null;
      
      Else
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, У�Ա�־, �����id, ���㿨���, ����,
           ������ˮ��, ����˵��, �������, ��������)
        Values
          (����Ԥ����¼_Id.Nextval, 3, Null, 2, r_Balancedata.����id, Null, v_����ժҪ, v_���㷽ʽ, r_Balancedata.�տ�ʱ��,
           r_Balancedata.����Ա���, r_Balancedata.����Ա����, n_������, r_Balancedata.����id, r_Balancedata.�ɿ���id, r_Balancedata.�������,
           2, �����id_In, Null, ����_In, ������ˮ��_In, ����˵��_In, v_�������, 3);
      
        Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - n_������ Where ����id = ����id_In And ���㷽ʽ Is Null;
        If Sql%NotFound Then
          v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[�շѽ��㴰��]�������շѣ�';
          Raise Err_Item;
        End If;
      End If;
    End If;
  End If;

  If ��������_In = 3 Then
    --   3-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
    --3.1����Ƿ��Ѿ�����ҽ����������,������ɾ��
    n_������ := 0;
    For v_ҽ�� In (Select ID, ���㷽ʽ, ��Ԥ��
                 From ����Ԥ����¼ A
                 Where ����id = ����id_In And Exists
                  (Select 1 From ���㷽ʽ Where a.���㷽ʽ = ���� And ���� In (3, 4)) And Mod(��¼����, 10) <> 1) Loop
      n_������ := n_������ + Nvl(v_ҽ��.��Ԥ��, 0);
      l_Ԥ��id.Extend;
      l_Ԥ��id(l_Ԥ��id.Count) := v_ҽ��.Id;
    End Loop;
    If Nvl(n_������, 0) <> 0 Then
      Update ����Ԥ����¼
      Set ��Ԥ�� = Nvl(��Ԥ��, 0) + Nvl(n_������, 0)
      Where ����id = ����id_In And ���㷽ʽ Is Null;
      If Sql%NotFound Then
        v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[�շѽ��㴰��]�������շѣ�';
        Raise Err_Item;
      End If;
    End If;
    If l_Ԥ��id.Count <> 0 Then
      Forall I In 1 .. l_Ԥ��id.Count
        Delete ����Ԥ����¼ Where ID = l_Ԥ��id(I);
    End If;
  
    If ���㷽ʽ_In Is Not Null Then
      v_�������� := ���㷽ʽ_In || '||';
    End If;
  
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
      n_������ := To_Number(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
    
      For c_������Ϣ In (Select ��¼����, NO, ��¼״̬, ����id, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������
                     From ����Ԥ����¼
                     Where ����id = ����id_In And ���㷽ʽ Is Null) Loop
      
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, У�Ա�־, ��������)
        Values
          (����Ԥ����¼_Id.Nextval, 3, Null, 2, c_������Ϣ.����id, Null, '���ս���', v_���㷽ʽ, c_������Ϣ.�տ�ʱ��, c_������Ϣ.����Ա���, c_������Ϣ.����Ա����,
           n_������, c_������Ϣ.����id, c_������Ϣ.�ɿ���id, c_������Ϣ.�������, 1, 3);
      End Loop;
      --��������(���㷽ʽΪNULL��)
      Update ����Ԥ����¼
      Set ��Ԥ�� = ��Ԥ�� - n_������
      Where ����id = ����id_In And ���㷽ʽ Is Null
      Returning Nvl(��Ԥ��, 0) Into n_����ֵ;
    
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;

  --4-���ѿ���������
  If ��������_In = 4 Then
    v_�������� := ���㷽ʽ_In || '||';
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      --�����ID|����|���ѿ�ID|���ѽ��
      n_�����id := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      v_����     := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_���ѿ�id := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_������ := To_Number(v_��ǰ����);
      Begin
        Select ����, ���ƿ�, ���㷽ʽ Into v_����, n_���ƿ�, v_���㷽ʽ From �����ѽӿ�Ŀ¼ Where ��� = n_�����id;
      Exception
        When Others Then
          v_���� := Null;
      End;
      If v_���� Is Null Then
        v_Err_Msg := 'δ�ҵ���Ӧ�Ľ��㿨�ӿ�,����ˢ������ʧ��!';
        Raise Err_Item;
      End If;
      If v_���㷽ʽ Is Null Then
        v_Err_Msg := v_���� || 'δ���ö�Ӧ�Ľ��㷽ʽ,����ˢ������ʧ��!';
        Raise Err_Item;
      End If;
      If Nvl(n_������, 0) <> 0 Then
        n_����id := ����id_In;
      
        If Nvl(n_�ؽ�id, 0) <> 0 Then
        
          Select Sum(��Ԥ��)
          Into n_����ֵ
          From ����Ԥ����¼
          Where ����id = ����id_In And ���㷽ʽ Is Null And Nvl(��Ԥ��, 0) <> 0;
          If Nvl(n_����ֵ, 0) <> 0 Then
          
            Update ����Ԥ����¼
            Set ��Ԥ�� = Nvl(��Ԥ��, 0) + Nvl(n_����ֵ, 0)
            Where ����id = ����id_In And ���㷽ʽ = v_���㷽ʽ And ���㿨��� = n_�����id
            Returning ID Into n_Ԥ��id;
          
            If Sql%NotFound Then
              Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
              Insert Into ����Ԥ����¼
                (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, ���㿨���, У�Ա�־,
                 ��������)
              Values
                (n_Ԥ��id, 3, Null, 2, r_Balancedata. ����id, Null, Null, v_���㷽ʽ, r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա���,
                 r_Balancedata. ����Ա����, n_����ֵ, r_Balancedata.����id, r_Balancedata. �ɿ���id, r_Balancedata.�������, n_�����id, 2,
                 3);
            End If;
            Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - Nvl(n_����ֵ, 0) Where ����id = ����id_In And ���㷽ʽ Is Null;
          
            --���뿨�����¼
            --������ѿ��Ƿ���ȷ
            n_��� := Zl_���ѿ�Ŀ¼_Check(n_�����id, v_����, n_���ѿ�id, v_Err_Msg);
            If Nvl(n_���, 0) = 0 Then
              Raise Err_Item;
            End If;
            Begin
              Select Nvl(Max(Nvl(���, 0)), 0) + 1
              Into n_���
              From ���˿������¼
              Where �ӿڱ�� = n_�����id And Nvl(���ѿ�id, 0) = Nvl(n_���ѿ�id, 0) And ���� = v_����;
            Exception
              When Others Then
                n_��� := 1;
            End;
          
            Select ���˿������¼_Id.Nextval Into n_Id From Dual;
          
            Insert Into ���˿������¼
              (ID, �ӿڱ��, ���ѿ�id, ���, ��¼״̬, ���㷽ʽ, ������, ����, ������ˮ��, ����ʱ��, ��ע, �����־)
            Values
              (n_Id, n_�����id, n_���ѿ�id, n_���, 1, v_���㷽ʽ, n_����ֵ, v_����, Null, r_Balancedata.�տ�ʱ��, Null, 0);
          
            --������ѿ�,��ͬʱ���������
            If Nvl(n_���ѿ�id, 0) <> 0 Then
              Update ���ѿ�Ŀ¼ Set ��� = ��� - n_����ֵ Where ID = n_���ѿ�id;
              If Sql%NotFound Then
                v_Err_Msg := '����Ϊ' || v_���� || '��' || v_���� || 'δ�ҵ�!';
                Raise Err_Item;
              End If;
            End If;
            Insert Into ���˿�������� (Ԥ��id, ������id) Values (n_Ԥ��id, n_Id);
            n_������ := n_������ - Nvl(n_����ֵ, 0);
          End If;
          n_����id := n_�ؽ�id;
        
        End If;
      
        Update ����Ԥ����¼
        Set ��Ԥ�� = Nvl(��Ԥ��, 0) + n_������
        Where ����id = n_����id And ���㷽ʽ = v_���㷽ʽ And ���㿨��� = n_�����id
        Returning ID Into n_Ԥ��id;
      
        If Sql%NotFound Then
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          Insert Into ����Ԥ����¼
            (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, ���㿨���, У�Ա�־, ��������)
          Values
            (n_Ԥ��id, 3, Null, Decode(Nvl(n_�ؽ�id, 0), 0, 2, 1), r_Balancedata. ����id, Null, Null, v_���㷽ʽ,
             r_Balancedata. �տ�ʱ��, r_Balancedata. ����Ա���, r_Balancedata. ����Ա����, n_������, n_����id, r_Balancedata. �ɿ���id,
             r_Balancedata. �������, n_�����id, 2, 3);
        End If;
      
        --���뿨�����¼
        n_��� := Zl_���ѿ�Ŀ¼_Check(n_�����id, v_����, n_���ѿ�id, v_Err_Msg);
        If Nvl(n_���, 0) = 0 Then
          Raise Err_Item;
        End If;
      
        Begin
          Select Nvl(Max(Nvl(���, 0)), 0) + 1
          Into n_���
          From ���˿������¼
          Where �ӿڱ�� = n_�����id And Nvl(���ѿ�id, 0) = Nvl(n_���ѿ�id, 0) And ���� = v_����;
        Exception
          When Others Then
            n_��� := 1;
        End;
      
        Select ���˿������¼_Id.Nextval Into n_Id From Dual;
      
        Insert Into ���˿������¼
          (ID, �ӿڱ��, ���ѿ�id, ���, ��¼״̬, ���㷽ʽ, ������, ����, ������ˮ��, ����ʱ��, ��ע, �����־)
        Values
          (n_Id, n_�����id, n_���ѿ�id, n_���, 1, v_���㷽ʽ, n_������, v_����, Null, r_Balancedata.�տ�ʱ��, Null, 0);
        --������ѿ�,��ͬʱ���������
        If Nvl(n_���ѿ�id, 0) <> 0 Then
          Update ���ѿ�Ŀ¼ Set ��� = ��� - n_������ Where ID = n_���ѿ�id;
          If Sql%NotFound Then
            v_Err_Msg := '����Ϊ' || v_���� || '��' || v_���� || 'δ�ҵ�!';
            Raise Err_Item;
          End If;
        End If;
        Insert Into ���˿�������� (Ԥ��id, ������id) Values (n_Ԥ��id, n_Id);
      
        --��������(���㷽ʽΪNULL��)
        Update ����Ԥ����¼
        Set ��Ԥ�� = ��Ԥ�� - n_������
        Where ����id = n_����id And ���㷽ʽ Is Null
        Returning Nvl(��Ԥ��, 0) Into n_����ֵ;
      End If;
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;

  If Nvl(����˷�_In, 0) = 0 Then
    Return;
  End If;

  -----------------------------------------------------------------------------------------
  --����շ�,��Ҫ������Ա�ɿ����,Ԥ����¼(���㷽ʽ=NULL)
  If Nvl(����˷�_In, 0) = 1 Then
    Update ����Ԥ����¼ Set �ɿ� = �ɿ�_In, �Ҳ� = �Ҳ�_In, У�Ա�־ = 0 Where ����id = ����id_In;
    Return;
  End If;

  --1.ɾ�����㷽ʽΪNULL��Ԥ����¼
  Delete ����Ԥ����¼ Where ����id = ����id_In And ���㷽ʽ Is Null And Nvl(��Ԥ��, 0) = 0;
  If Sql%NotFound Then
    Select Count(*) Into n_Count From ����Ԥ����¼ A Where ����id = ����id_In And ���㷽ʽ Is Null;
    If n_Count <> 0 Then
      v_Err_Msg := '������δ�ɿ������,������ɽ���!';
    Else
      v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[�շѽ��㴰��]�������շѣ�!';
    End If;
    Raise Err_Item;
  End If;
  If Nvl(n_�ؽ�id, 0) <> 0 Then
    Delete ����Ԥ����¼ Where ����id = n_�ؽ�id And ���㷽ʽ Is Null And Nvl(��Ԥ��, 0) = 0;
    If Sql%NotFound Then
      Select Count(*) Into n_Count From ����Ԥ����¼ A Where ����id = n_�ؽ�id And ���㷽ʽ Is Null;
      If n_Count <> 0 Then
        v_Err_Msg := '������δ�ɿ������,������ɽ���!';
      Else
        v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[�շѽ��㴰��]�������շѣ�!';
      End If;
      Raise Err_Item;
    End If;
    Update ����Ԥ����¼ Set �ɿ� = �ɿ�_In, �Ҳ� = �Ҳ�_In, У�Ա�־ = 0 Where ����id = n_�ؽ�id;
  
  End If;

  --������Ϊ��ʱ������һ�����Ϊ0�Ĳ���Ԥ����¼
  Select Count(*) Into n_Count From ����Ԥ����¼ A Where ����id = ����id_In;

  If n_Count = 0 Then
    v_���㷽ʽ := ȱʡ���㷽ʽ_In;
    If v_���㷽ʽ Is Null Then
      Begin
        Select ���㷽ʽ Into v_���㷽ʽ From ���㷽ʽӦ�� Where Ӧ�ó��� = '�շ�' And Nvl(ȱʡ��־, 0) = 1;
      Exception
        When Others Then
          v_���㷽ʽ := Null;
      End;
      If v_���㷽ʽ Is Null Then
        Begin
          Select ���� Into v_���㷽ʽ From ���㷽ʽ Where Nvl(����, 0) = 1;
        Exception
          When Others Then
            v_���㷽ʽ := '�ֽ�';
        End;
      End If;
    End If;
    Insert Into ����Ԥ����¼
      (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, У�Ա�־, �����id, ���㿨���, ����,
       ������ˮ��, ����˵��, �������, ��������)
    Values
      (����Ԥ����¼_Id.Nextval, 3, Null, 1, r_Balancedata.����id, Null, Null, v_���㷽ʽ, r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա���,
       r_Balancedata.����Ա����, 0, r_Balancedata.����id, r_Balancedata.�ɿ���id, r_Balancedata.�������, 2, Null, Null, Null, Null,
       ����˵��_In, Null, 3);
  End If;

  --2.����ɿ����ݺ��Ҳ����ݼ�У�Ա�־����Ϊ0
  Update ����Ԥ����¼ Set �ɿ� = �ɿ�_In, �Ҳ� = �Ҳ�_In, У�Ա�־ = 0 Where ����id = ����id_In;

  --3.���·���״̬
  Update ������ü�¼ Set ����״̬ = 0 Where ����id = ����id_In;
  If Nvl(n_�ؽ�id, 0) <> 0 Then
    Update ������ü�¼ Set ����״̬ = 0 Where ����id = n_�ؽ�id;
  End If;

  --4.������Ա�ɿ�����
  If n_�ؽ�id <> 0 Then
    For c_�ɿ� In (Select ���㷽ʽ, ����Ա����, Sum(Nvl(a.��Ԥ��, 0)) As ��Ԥ��
                 From ����Ԥ����¼ A
                 Where a.����id In (����id_In, n_�ؽ�id) And Mod(a.��¼����, 10) <> 1
                 Group By ���㷽ʽ, ����Ա����) Loop
    
      Update ��Ա�ɿ����
      Set ��� = Nvl(���, 0) + Nvl(c_�ɿ�.��Ԥ��, 0)
      Where �տ�Ա = c_�ɿ�.����Ա���� And ���� = 1 And ���㷽ʽ = c_�ɿ�.���㷽ʽ;
      If Sql%RowCount = 0 Then
        Insert Into ��Ա�ɿ����
          (�տ�Ա, ���㷽ʽ, ����, ���)
        Values
          (c_�ɿ�.����Ա����, c_�ɿ�.���㷽ʽ, 1, Nvl(c_�ɿ�.��Ԥ��, 0));
      End If;
    End Loop;
  
  Else
    For c_�ɿ� In (Select ���㷽ʽ, ����Ա����, Sum(Nvl(a.��Ԥ��, 0)) As ��Ԥ��
                 From ����Ԥ����¼ A
                 Where a.����id = ����id_In And Mod(a.��¼����, 10) <> 1
                 Group By ���㷽ʽ, ����Ա����) Loop
    
      Update ��Ա�ɿ����
      Set ��� = Nvl(���, 0) + Nvl(c_�ɿ�.��Ԥ��, 0)
      Where �տ�Ա = c_�ɿ�.����Ա���� And ���� = 1 And ���㷽ʽ = c_�ɿ�.���㷽ʽ;
      If Sql%RowCount = 0 Then
        Insert Into ��Ա�ɿ����
          (�տ�Ա, ���㷽ʽ, ����, ���)
        Values
          (c_�ɿ�.����Ա����, c_�ɿ�.���㷽ʽ, 1, Nvl(c_�ɿ�.��Ԥ��, 0));
      End If;
    End Loop;
  End If;
  --��Ϣ����
  Select ����id_In || ',' || ����id_In || ',' || Decode(����˷�_In, 2, 0, 0, 0, 1) Into v_Msg From Dual;
  Begin
    Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
      Using 5, v_Msg;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����˷ѽ���_Modify;
/

--89305:�ŵ���,2015-02-24,ɨ����ƿǩ�Զ���ҩ
CREATE OR REPLACE Procedure Zl_��Һ��ҩ��¼_����
(
  ��ҩid_In   In Varchar2, --ID��:ID1,ID2....
  ������Ա_In In ��Һ��ҩ��¼.������Ա%Type := Null,
  ����ʱ��_In In ��Һ��ҩ��¼.����ʱ��%Type := Null,
  ����˵��_In In ��Һ��ҩ״̬.����˵��%Type := Null
) Is
  v_Tansid Varchar2(20);
  v_Tmp    Varchar2(4000);
  v_Error    Varchar2(255);
  n_����״̬ ��Һ��ҩ��¼.����״̬%Type;
  Err_Custom Exception;
Begin
  If ��ҩid_In Is Null Then
    v_Tmp := Null;
  Else
    v_Tmp := ��ҩid_In || ',';
  End If;
  While v_Tmp Is Not Null Loop
    --�ֽⵥ��ID��
    v_Tansid := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp    := Replace(',' || v_Tmp, ',' || v_Tansid || ',');

     Begin
      Select ����״̬ Into n_����״̬ From ��Һ��ҩ��¼ Where ID = v_Tansid;

      if n_����״̬>4 then
        v_Error := '�������ѱ����������ܽ��з��Ͳ�����';
        Raise Err_Custom;
      end if;
    Exception
      When Others Then
        v_Error := '�������ѱ�������';
        Raise Err_Custom;
    End;

    Insert Into ��Һ��ҩ״̬ (��ҩid, ��������, ������Ա, ����ʱ��,����˵��) Values (v_Tansid, 5, ������Ա_In, ����ʱ��_In,����˵��_In);
    Update ��Һ��ҩ��¼ Set ����״̬ = 5, ������Ա = ������Ա_In, ����ʱ�� = ����ʱ��_In Where ID = v_Tansid;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Һ��ҩ��¼_����;
/


--93235:������,2016-02-24,����������Ϣ
Create Or Replace Procedure Zl_���˽��ʼ�¼_Insert
(
  Id_In           ���˽��ʼ�¼.Id%Type,
  ���ݺ�_In       ���˽��ʼ�¼.No%Type,
  ����id_In       ���˽��ʼ�¼.����id%Type,
  �շ�ʱ��_In     ���˽��ʼ�¼.�շ�ʱ��%Type,
  ��ʼ����_In     ���˽��ʼ�¼.��ʼ����%Type,
  ��������_In     ���˽��ʼ�¼.��������%Type,
  ��;����_In     ���˽��ʼ�¼.��;����%Type := 0,
  �ಡ�˽���_In   Number := 0,
  �����ʴ���_In Number := 0,
  ��ע_In         ���˽��ʼ�¼.��ע%Type := Null,
  ��Դ_In         Number := 1,
  ԭ��_In         ���˽��ʼ�¼.ԭ��%Type := Null,
  ��������_In     ���˽��ʼ�¼.��������%Type := 2
  --���ܣ�����һ�����˽��ʼ�¼
  --1.��Դ_In:1-����;2-סԺ
) As
  v_Temp     Varchar2(255);
  v_��Ա��� ��Ա��.���%Type;
  v_��Ա���� ��Ա��.����%Type;
  v_Error    Varchar2(255);
  Err_Custom Exception;
  n_��id ����ɿ����.Id%Type;
Begin
  --����Ա��Ϣ:����ID,��������;��ԱID,��Ա���,��Ա����
  v_Temp     := Zl_Identity;
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_��Ա��� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  n_��id     := Zl_Get��id(v_��Ա����);

  --���˽��ʼ�¼
  Insert Into ���˽��ʼ�¼
    (ID, NO, ʵ��Ʊ��, ��¼״̬, ��;����, ����id, ��ʼ����, ��������, �շ�ʱ��, ����Ա���, ����Ա����, ��ע, ԭ��, �ɿ���id, ��������)
  Values
    (Id_In, ���ݺ�_In, Null, 1, ��;����_In, Decode(�ಡ�˽���_In, 1, Null, ����id_In), ��ʼ����_In, ��������_In, �շ�ʱ��_In, v_��Ա���, v_��Ա����,
     ��ע_In, ԭ��_In, n_��id, ��������_In);
     
  Begin
    Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
      Using 15, Id_In;
  Exception
    When Others Then
      Null;
  End;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˽��ʼ�¼_Insert;
/

--93380:������,2016-02-24,��������鳤����
Create Or Replace Procedure Zl_�ɿ��Ա���_Insert
(
  ��id_In     In �ɿ��Ա���.��id%Type,
  ��Աid_In   In �ɿ��Ա���.��Աid%Type,
  ��������_In Number := 0
) Is
  --��������_IN:0-�����ɿ��Ա,1-�����ɿ��鳤
  v_Err_Msg Varchar2(500);
  Err_Item Exception;
  n_Count Number(18);
Begin
  If ��������_In = 0 Then
    --����Ƿ��Ѿ�������������
    Select Count(b.������)
    Into n_Count
    From �ɿ��Ա��� A, ����ɿ���� B
    Where a.��id = b.Id And a.��id <> ��id_In And ��Աid = ��Աid_In And ɾ������ >= Sysdate;
    If n_Count <> 0 Then
      v_Err_Msg := '[ZLSOFT]����Ա�����򲢷���ԭ���Ѿ������䵽������,�����ٷ��䵽������,�����¶�ȡ����Ա![ZLSOFT]';
      Raise Err_Item;
    End If;
    Insert Into �ɿ��Ա��� (��id, ��Աid) Values (��id_In, ��Աid_In);
  Else
    Select Count(b.������)
    Into n_Count
    From �ɿ��Ա��� A, ����ɿ���� B
    Where a.��id = b.Id And a.��id <> ��id_In And ��Աid = ��Աid_In And ɾ������ >= Sysdate;
    If n_Count <> 0 Then
      v_Err_Msg := '[ZLSOFT]����Ա�����򲢷���ԭ���Ѿ������䵽������,�����ٷ����Ϊ�鳤��,�����¶�ȡ����Ա![ZLSOFT]';
      Raise Err_Item;
    End If;
    Insert Into �������鳤���� (��id, �鳤id) Values (��id_In, ��Աid_In);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ɿ��Ա���_Insert;
/

--93380:������,2016-02-24,��������鳤����
Create Or Replace Procedure Zl_�ɿ��Ա���_Move
(
  ��Աid_In   Varchar2,
  ԭ��id_In   In �ɿ��Ա���.��id%Type,
  ����id_In   In �ɿ��Ա���.��id%Type := -1,
  ��������_In In Number := 0
) Is
  --��������_IN:0-����ɿ��Ա,1-����ɿ��鳤
  --��ԱID_IN:�����Աʱ,�ö��ŷ���
  --����ID_IN:-1��ʾ�Ƴ�;
  n_Pos     Number;
  v_Temp    Varchar2(4000);
  n_��Աid  �ɿ��Ա���.��Աid%Type;
  n_���    ��Ա�ɿ����.���%Type;
  v_����    ��Ա��.����%Type;
  v_Err_Msg Varchar2(500);
  Err_Item Exception;
  l_��Աid t_Numlist := t_Numlist();
Begin
  --ѭ������
  v_Temp := ��Աid_In;
  If ��������_In = 0 Then
    While v_Temp Is Not Null Loop
      n_Pos := Instr(v_Temp, ',');
      If n_Pos = 0 Then
        If Not v_Temp Is Null Then
          n_��Աid := To_Number(v_Temp);
          v_Temp   := Null;
          Select Sum(Nvl(a.���, 0)), Max(a.�տ�Ա)
          Into n_���, v_����
          From ��Ա�ɿ���� A, ��Ա�� B
          Where a.���� = 1 And a.�տ�Ա = b.���� And b.Id = n_��Աid;
          If Nvl(n_���, 0) <> 0 And ����id_In > 0 Then
            v_Err_Msg := '[ZLSOFT]��Ա��' || v_���� || '���������ݴ��,���ܵ�����������![ZLSOFT]';
            Raise Err_Item;
          End If;
          l_��Աid.Extend;
          l_��Աid(l_��Աid.Count) := n_��Աid;
        End If;
      Else
        --�õ���ԱID
        n_��Աid := To_Number(Substr(v_Temp, 1, n_Pos - 1));
        v_Temp   := Substr(v_Temp, n_Pos + 1);
        Select Sum(Nvl(a.���, 0)), Max(a.�տ�Ա)
        Into n_���, v_����
        From ��Ա�ɿ���� A, ��Ա�� B
        Where a.���� = 1 And a.�տ�Ա = b.���� And b.Id = n_��Աid;
        If Nvl(n_���, 0) <> 0 And ����id_In > 0 Then
          v_Err_Msg := '[ZLSOFT]��Ա��' || v_���� || '���������ݴ��,���ܵ�����������![ZLSOFT]';
          Raise Err_Item;
        End If;
        l_��Աid.Extend;
        l_��Աid(l_��Աid.Count) := n_��Աid;
      End If;
    End Loop;
  
    Forall I In 1 .. l_��Աid.Count
      Delete �ɿ��Ա��� Where ��id = ԭ��id_In And ��Աid = l_��Աid(I);
    If ����id_In > 0 Then
      Forall I In 1 .. l_��Աid.Count
        Insert Into �ɿ��Ա��� (��id, ��Աid) Values (����id_In, l_��Աid(I));
    End If;
  Else
    While v_Temp Is Not Null Loop
      n_Pos := Instr(v_Temp, ',');
      If n_Pos = 0 Then
        If Not v_Temp Is Null Then
          n_��Աid := To_Number(v_Temp);
          v_Temp   := Null;
          l_��Աid.Extend;
          l_��Աid(l_��Աid.Count) := n_��Աid;
        End If;
      Else
        --�õ���ԱID
        n_��Աid := To_Number(Substr(v_Temp, 1, n_Pos - 1));
        v_Temp   := Substr(v_Temp, n_Pos + 1);
        l_��Աid.Extend;
        l_��Աid(l_��Աid.Count) := n_��Աid;
      End If;
    End Loop;
  
    Forall I In 1 .. l_��Աid.Count
      Delete �������鳤���� Where ��id = ԭ��id_In And �鳤id = l_��Աid(I);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ɿ��Ա���_Move;
/

--93380:������,2016-02-24,��������鳤����
Create Or Replace Procedure Zl_С�����ʼ�¼_Insert
(
  Id_In       In ��Ա�սɼ�¼.Id%Type,
  ���ݺ�_In   In ��Ա�սɼ�¼.No%Type,
  �ɿ���id_In In ��Ա�սɼ�¼.�ɿ���id%Type,
  ��ʼʱ��_In In ��Ա�սɼ�¼.��ʼʱ��%Type,
  ��ֹʱ��_In In ��Ա�սɼ�¼.��ֹʱ��%Type,
  �տ�Ա_In   In ��Ա�սɼ�¼.�տ�Ա%Type,
  �տ�Աid_In In ��Ա��.Id%Type,
  �տ�ʱ��_In In ��Ա�սɼ�¼.С���տ�ʱ��%Type,
  ������Ϣ_In In Varchar2,
  ��������_In In Number := 0
) Is
  ----------------------------------------------------------------------------------------
  --����:���������ʼ�¼д��
  --����:
  --     ������Ϣ_IN:��¼ID1,��¼ID2,...��¼IDn
  --     ��������_IN:0-��������� 1-������ӵ�һ����¼ 2-��������м�¼ 3-����������һ����¼
  ----------------------------------------------------------------------------------------
  v_Err_Msg Varchar2(100);
  Err_Item Exception;
  n_Exists   Number(2);
  v_No       ��Ա�սɼ�¼.No%Type;
  v_�տ�Ա   ��Ա�սɼ�¼.�տ�Ա%Type;
  n_��Ԥ���� ��Ա�սɼ�¼.��Ԥ����%Type := 0;
  n_����ϼ� ��Ա�սɼ�¼.����ϼ�%Type := 0;
  n_����ϼ� ��Ա�սɼ�¼.����ϼ�%Type := 0;

Begin
  --����ǰ�Ĳ������  
  n_Exists := 0;
  Begin
    Select /*+ Rule*/
     a.No, a.�տ�Ա, 1
    Into v_No, v_�տ�Ա, n_Exists
    From ��Ա�սɼ�¼ A, Table(f_Num2list(������Ϣ_In)) B
    Where a.Id = b.Column_Value And С������id Is Not Null And Rownum < 2;
  Exception
    When Others Then
      n_Exists := 0;
  End;
  If n_Exists = 1 Then
    v_Err_Msg := v_�տ�Ա || '�տ�Ա���տ��Ϊ[' || v_No || ']�ļ�¼�ѱ����ʣ��������ٴ����ʣ�';
    Raise Err_Item;
  End If;
  Begin
    Select /*+ Rule*/
     a.No, a.�տ�Ա, 1
    Into v_No, v_�տ�Ա, n_Exists
    From ��Ա�սɼ�¼ A, Table(f_Num2list(������Ϣ_In)) B
    Where a.Id = b.Column_Value And �����տ�ʱ�� Is Not Null And Rownum < 2;
  Exception
    When Others Then
      n_Exists := 0;
  End;
  If n_Exists = 1 Then
    v_Err_Msg := v_�տ�Ա || '�տ�Ա���տ��Ϊ[' || v_No || ']�ļ�¼�ѱ�������տ������С��������ʣ�';
    Raise Err_Item;
  End If;
  Begin
    Select /*+ Rule*/
     a.No, a.�տ�Ա, 1
    Into v_No, v_�տ�Ա, n_Exists
    From ��Ա�սɼ�¼ A, Table(f_Num2list(������Ϣ_In)) B
    Where a.Id = b.Column_Value And ����ʱ�� Is Not Null And Rownum < 2;
  Exception
    When Others Then
      n_Exists := 0;
  End;
  If n_Exists = 1 Then
    v_Err_Msg := v_�տ�Ա || '�տ�Ա���տ��Ϊ[' || v_No || ']�ļ�¼�ѱ����ϣ�������������ʣ�';
    Raise Err_Item;
  End If;

  --���ܽ��
  Select /*+ Rule*/
   Sum(a.��Ԥ����), Sum(a.����ϼ�), Sum(a.����ϼ�)
  Into n_��Ԥ����, n_����ϼ�, n_����ϼ�
  From ��Ա�սɼ�¼ A, Table(f_Num2list(������Ϣ_In)) B
  Where a.Id = b.Column_Value
  Group By a.��¼����;

  If ��������_In = 0 Or ��������_In = 1 Then
    --�������ʼ�¼
    Insert Into ��Ա�սɼ�¼
      (ID, ��¼����, NO, �տ�Ա, �ɿ���id, �Ǽ�ʱ��, ��ʼʱ��, ��ֹʱ��, ��Ԥ����, ����ϼ�, ����ϼ�, С������id, �Ǽ���)
      Select Id_In, 3, ���ݺ�_In, �տ�Ա_In, �ɿ���id_In, �տ�ʱ��_In, ��ʼʱ��_In, ��ֹʱ��_In, n_��Ԥ����, n_����ϼ�, n_����ϼ�, Id_In, �տ�Ա_In
      From Dual;
    Update ����ɿ���� Set �ϴ�����ʱ�� = ��ֹʱ��_In Where ID = �ɿ���id_In;
    Update �������鳤���� Set �ϴ�����ʱ�� = ��ֹʱ��_In Where ��id = �ɿ���id_In And �鳤id = �տ�Աid_In;
  End If;
  --�������ʼ�¼
  If ��������_In = 2 Or ��������_In = 3 Then
    Update ��Ա�սɼ�¼
    Set ��Ԥ���� = ��Ԥ���� + n_��Ԥ����, ����ϼ� = ����ϼ� + n_����ϼ�, ����ϼ� = ����ϼ� + n_����ϼ�
    Where ID = Id_In;
  End If;

  Update ��Ա�սɼ�¼
  Set С������id = Id_In
  Where С���տ�id In (Select Column_Value From Table(f_Num2list(������Ϣ_In)));

  If ��������_In = 0 Or ��������_In = 3 Then
    --������ϸ
    Insert Into ��Ա�ս���ϸ
      (�ս�id, ���㷽ʽ, ���)
      Select Id_In, a.���㷽ʽ, Sum(a.���)
      From ��Ա�ս���ϸ A, ��Ա�սɼ�¼ B
      Where a.�ս�id = b.Id And b.��¼���� = 2 And b.С������id = Id_In
      Group By Id_In, ���㷽ʽ;
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_С�����ʼ�¼_Insert;
/

--93380:������,2016-02-24,��������鳤����
CREATE OR REPLACE Procedure Zl_С�����ʼ�¼_Cancel
(
  Id_In       In ��Ա�սɼ�¼.Id%Type,
  ������_In   In ��Ա�սɼ�¼.������%Type,
  ������ID_IN In ��Ա��.ID%Type,
  ����ʱ��_In In ��Ա�սɼ�¼.����ʱ��%Type,
  �ɿ���id_In In ����ɿ����.Id%Type
) Is
  ----------------------------------------------------------------------------------------
  --����:���������ʼ�¼����
  --����:ID_IN:Ҫ���ϼ�¼��ID
  ----------------------------------------------------------------------------------------
  v_Err_Msg Varchar2(100);
  Err_Item Exception;
  n_Exists   Number(2);
  n_Count    Number(18);
  v_No       ��Ա�սɼ�¼.No%Type;
  d_��ֹʱ�� ��Ա�սɼ�¼.��ֹʱ��%Type;
  v_�տ�Ա   ��Ա�սɼ�¼.�տ�Ա%Type;
Begin
  ---����ǰ�������
  n_Exists := 0;
  Begin
    Select NO, �տ�Ա Into v_No, v_�տ�Ա From ��Ա�սɼ�¼ Where ID = Id_In;
  Exception
    When Others Then
      n_Exists := 1;
  End;
  If n_Exists = 1 Then
    v_Err_Msg := '��¼δ���ҵ��������ѱ�ɾ�����޷������������ϲ�����';
    Raise Err_Item;
  End If;
  Begin
    Select 1 Into n_Exists From ��Ա�սɼ�¼ Where �����տ�ʱ�� Is Not Null And ID = Id_In;
  Exception
    When Others Then
      n_Exists := 0;
  End;
  If n_Exists = 1 Then
    v_Err_Msg := v_�տ�Ա || '�տ�Ա�����ʵ���Ϊ[' || v_No || ']�ļ�¼�ѱ�������տ���������ϣ�';
    Raise Err_Item;
  End If;
  Begin
    Select 1 Into n_Exists From ��Ա�սɼ�¼ Where ����ʱ�� Is Not Null And ID = Id_In;
  Exception
    When Others Then
      n_Exists := 0;
  End;
  If n_Exists = 1 Then
    v_Err_Msg := v_�տ�Ա || '�տ�Ա�����ʵ���Ϊ[' || v_No || ']�ļ�¼�ѱ����ϣ��������ٴ����ϣ�';
    Raise Err_Item;
  End If;

  --����Ƿ����һ�����ʼ�¼
  Select Count(*)
  Into n_Count
  From ��Ա�սɼ�¼
  Where �Ǽ�ʱ�� > (Select �Ǽ�ʱ�� From ��Ա�սɼ�¼ Where ID = Id_In) And ��¼���� = 3 And ID + 0 <> Id_In And Rownum < 2 And
        �տ�Ա || '' = v_�տ�Ա And ����ʱ�� Is Null;

  If n_Count >= 1 Then
    --�ǲ������һ�ε����ʼ�¼
    v_Err_Msg := '���ʵ���Ϊ:' || v_No || '�����ʼ�¼���������һ�ε����ʼ�¼,����������!';
    Raise Err_Item;
  End If;

  --�������ʲ���
  Update ��Ա�սɼ�¼ Set ������ = ������_In, ����ʱ�� = ����ʱ��_In Where ID = Id_In And ��¼���� = 3;
  Insert Into ��Ա�սɶ���
    (�ս�id, ����, ��¼id)
    Select Id_In, 8, ID From ��Ա�սɼ�¼ Where С������id = Id_In And ��¼���� = 2;
  Update ��Ա�սɼ�¼ Set С������id = Null Where С������id = Id_In;

  --�ָ����һ����Ч������ʱ��
  Select Max(��ֹʱ��)
  Into d_��ֹʱ��
  From ��Ա�սɼ�¼
  Where �Ǽ�ʱ�� <= (Select �Ǽ�ʱ�� From ��Ա�սɼ�¼ Where ID = Id_In) And ID + 0 <> Id_In And ����ʱ�� Is Null And �����տ�ʱ�� Is Null And
        �տ�Ա || '' = v_�տ�Ա And ��¼���� = 3;
  If d_��ֹʱ�� Is Null Then
    --ȡ������Сһ���տ��¼�ĵǼ�ʱ��
    Select Min(�Ǽ�ʱ��)
    Into d_��ֹʱ��
    From ��Ա�սɼ�¼
    Where ��¼���� = 2 And ����ʱ�� Is Null And �����տ�ʱ�� Is Null And �ɿ���id = �ɿ���id_In;
  End If;
  Update ����ɿ���� Set �ϴ�����ʱ�� = d_��ֹʱ�� Where ID = �ɿ���id_In;
  Update �������鳤���� Set �ϴ�����ʱ�� = d_��ֹʱ�� Where ��ID = �ɿ���id_In And �鳤ID=������ID_IN;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101,  '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_С�����ʼ�¼_Cancel;
/

--93380:������,2016-02-24,��������鳤����
Create Or Replace Procedure Zl_�շ�Ա���ʼ�¼_Insert
(
  Id_In         In ��Ա�սɼ�¼.Id%Type,
  No_In         In ��Ա�սɼ�¼.No%Type,
  �տ�Ա_In     In ��Ա�սɼ�¼.�տ�Ա%Type,
  �տ��id_In In ��Ա�սɼ�¼.�տ��id%Type,
  �鳤id_In     In ��Ա��.Id%Type,
  ��Ԥ����_In   In ��Ա�սɼ�¼.��Ԥ����%Type,
  ����ϼ�_In   In ��Ա�սɼ�¼.����ϼ�%Type,
  ����ϼ�_In   In ��Ա�սɼ�¼.����ϼ�%Type,
  ժҪ_In       In ��Ա�սɼ�¼.ժҪ%Type,
  ��ʼʱ��_In   In ��Ա�սɼ�¼.��ʼʱ��%Type,
  ��ֹʱ��_In   In ��Ա�սɼ�¼.��ֹʱ��%Type,
  �Ǽ���_In     In ��Ա�սɼ�¼.�Ǽ���%Type,
  �Ǽ�ʱ��_In   In ��Ա�սɼ�¼.�Ǽ�ʱ��%Type,
  �սɱ�־_In   In ��Ա�սɼ�¼.�սɱ�־%Type,
  �շѶ���_In   In Varchar2,
  �������_In   In Integer := 0,
  �ݴ��_In     In ��Ա�ݴ��¼.���%Type := 0,
  ���_In       In ��Ա�սɼ�¼.���%Type := 0
) Is
  --------------------------------------------------------------------------------------------
  --����:�շ�Ա���ʼ�¼д��
  --����:��¼����_IN:
  --     �շѶ���_IN:����1,��¼ID1|����2,��¼ID2|...|����n,��¼IDn
  --     �������_In:0-�������ʼ�¼�Ͷ���;1-ֻ�������
  --------------------------------------------------------------------------------------------
  v_Err_Msg Varchar2(500);
  Err_Item Exception;

  n_��id       ��Ա�սɼ�¼.�ɿ���id%Type;
  n_Count      Number(18);
  n_�ݴ�id     ��Ա�ݴ��¼.Id%Type;
  v_�ݴ浥��   ��Ա�ݴ��¼.No%Type;
  v_С���տ��� ��Ա�սɼ�¼.С���տ���%Type;
  v_�ֽ�       ���㷽ʽ.����%Type;

Begin
  --�������,�Ƿ�ǰʱ�䷶Χ���Ѿ������ʼ�¼
  If �������_In = 0 Then
    Select Count(1)
    Into n_Count
    From ��Ա�սɼ�¼
    Where �տ�Ա = �տ�Ա_In And ��ʼʱ�� > ��ʼʱ��_In And ����ʱ�� Is Null And Nvl(���, 0) = Nvl(���_In, 0) And Rownum < 2 And ��¼���� = 1;
    If n_Count <> 0 Then
      v_Err_Msg := '�շ�Ա:"' || �տ�Ա_In || '"�ڵ�ǰ���ʷ�Χ���Ѿ������˽�������,�����ٽ������ʣ�';
      Raise Err_Item;
    End If;
  End If;

  --����Ƿ��Ѵ���������ϸ
  For c_��� In (Select /*+ rule*/
                a.����, a.��¼id, m.�տ�Ա
               From ��Ա�սɶ��� A, Table(f_Str2list2(�շѶ���_In, '|', ',')) B, ��Ա�սɼ�¼ M
               Where a.���� = b.C1 And a.��¼id = b.C2 And a.�ս�id = m.Id And m.����ʱ�� Is Null And Rownum < 2) Loop
    --1-�շ�(���Һ�),2-����,3-Ԥ��,4-���;5-���ѿ���ֵ;6--���ѿ���ֵ;7-�ݴ��(��������)�������տ���������϶��գ��������ӣ�
    If c_���.���� = 1 Then
      Select Decode(��¼����, 1, '�շ�', 4, '�Һ�', '') || '����Ϊ:' || NO || '��' || Decode(��¼����, 1, '�շ�', 4, '�Һ�', '') ||
              '��¼�Ѿ������ʣ������ٽ������ʴ���'
      Into v_Err_Msg
      From ������ü�¼
      Where ����id = c_���.��¼id And Rownum < 2;
    End If;
  
    If c_���.���� = 2 Then
      Select '���ʵ���Ϊ:' || NO || '�Ľ��ʼ�¼�Ѿ������ʣ������ٽ������ʴ���'
      Into v_Err_Msg
      From ���˽��ʼ�¼
      Where ID = c_���.��¼id And Rownum < 2;
    End If;
    If c_���.���� = 3 Then
      Select 'Ԥ������Ϊ:' || NO || '��Ԥ����¼�Ѿ������ʣ������ٽ������ʴ���'
      Into v_Err_Msg
      From ����Ԥ����¼
      Where ID = c_���.��¼id And Rownum < 2;
    End If;
    If c_���.���� = 4 And �տ�Ա_In = c_���.�տ�Ա Then
      Select '�շ�Ա��' || To_Char(���ʱ��, 'yyyy-mm-dd hh24:mi:ss') || '��' || Decode(�����, Null, '�����', '�����') ||
              '�Ѿ������ʣ������ٽ������ʴ���'
      Into v_Err_Msg
      From ��Ա����¼
      Where ID = c_���.��¼id And Rownum < 2;
    End If;
  
    If c_���.���� = 5 Then
      Select '�շ�Ա��' || To_Char(��ֵʱ��, 'yyyy-mm-dd hh24:mi:ss') || '�����ѿ���ֵ��¼�Ѿ������ʣ������ٽ������ʴ���'
      Into v_Err_Msg
      From ���ѿ���ֵ��¼
      Where ID = c_���.��¼id And Rownum < 2;
    End If;
    If c_���.���� = 6 Then
      Select '�շ�Ա�ڷ�����Ϊ:' || ���� || '���ҷ���ʱ��Ϊ:' || To_Char(����ʱ��, 'yyyy-mm-dd hh24:mi:ss') || '������¼�Ѿ������ʣ������ٽ������ʴ���'
      Into v_Err_Msg
      From ���ѿ�Ŀ¼
      Where ID = c_���.��¼id And Rownum < 2;
    End If;
    If c_���.���� = 7 Then
      Select '�ݴ浥��Ϊ:' || NO || '���ݴ��¼�Ѿ������ʣ������ٽ������ʴ���'
      Into v_Err_Msg
      From ��Ա�ݴ��¼
      Where ID = c_���.��¼id And Rownum < 2;
    End If;
    If c_���.���� = 9 Then
      Select '����Ϊ:' || NO || '�Ĳ�������¼�Ѿ������ʣ������ٽ������ʴ���'
      Into v_Err_Msg
      From ���ò����¼
      Where ����id = c_���.��¼id And Rownum < 2;
    End If;
    If v_Err_Msg Is Not Null Then
      Raise Err_Item;
    End If;
  End Loop;

  If �������_In = 0 Then
    n_��id := Zl_Get��id(�տ�Ա_In);
    If �鳤id_In Is Not Null Then
      Select ���� Into v_С���տ��� From ��Ա�� Where ID = �鳤id_In;
    End If;
    Insert Into ��Ա�սɼ�¼
      (ID, ��¼����, NO, �տ�Ա, �տ��id, ��Ԥ����, ����ϼ�, ����ϼ�, ժҪ, ��ʼʱ��, ��ֹʱ��, �ɿ���id, �Ǽ���, �Ǽ�ʱ��, �սɱ�־, ���, С���տ���)
    Values
      (Id_In, 1, No_In, �տ�Ա_In, �տ��id_In, ��Ԥ����_In, ����ϼ�_In, ����ϼ�_In, ժҪ_In, ��ʼʱ��_In, ��ֹʱ��_In, n_��id, �Ǽ���_In, �Ǽ�ʱ��_In,
       �սɱ�־_In, ���_In, v_С���տ���);
  
    Update ��Ա�ɿ���� Set �ϴ�����ʱ�� = ��ֹʱ��_In Where �տ�Ա = �տ�Ա_In;
  End If;
  --�����շѶ���
  Insert Into ��Ա�սɶ���
    (�ս�id, ����, ��¼id)
    Select Id_In, C1, C2 From Table(f_Str2list2(�շѶ���_In, '|', ','));
  If �ݴ��_In <> 0 Then
    Begin
      Select ���� Into v_�ֽ� From ���㷽ʽ Where ���� = 1 And Rownum < 2;
    Exception
      When Others Then
        v_�ֽ� := '�ֽ�';
    End;
    Select ��Ա�ݴ��¼_Id.Nextval Into n_�ݴ�id From Dual;
    v_�ݴ浥�� := Nextno(141);
    Insert Into ��Ա�ݴ��¼
      (ID, �ս�id, ��¼����, NO, ���㷽ʽ, ���, �տ�Ա, �Ǽ���, �Ǽ�ʱ��)
    Values
      (n_�ݴ�id, Id_In, 2, v_�ݴ浥��, v_�ֽ�, -1 * �ݴ��_In, �տ�Ա_In, �Ǽ���_In, �Ǽ�ʱ��_In);
  
    Insert Into ��Ա�սɶ��� (�ս�id, ����, ��¼id) Values (Id_In, 7, n_�ݴ�id);
    Select ��Ա�ݴ��¼_Id.Nextval Into n_�ݴ�id From Dual;
    v_�ݴ浥�� := Nextno(141);
    Insert Into ��Ա�ݴ��¼
      (ID, �ս�id, ��¼����, NO, ���㷽ʽ, ���, �տ�Ա, �Ǽ���, �Ǽ�ʱ��)
    Values
      (n_�ݴ�id, Id_In, 2, v_�ݴ浥��, v_�ֽ�, �ݴ��_In, �տ�Ա_In, �Ǽ���_In, �Ǽ�ʱ��_In + 1 / 24 / 60 / 60);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�շ�Ա���ʼ�¼_Insert;
/

--93564:Ƚ����,2016-02-24,���ý��Ϊ0ʱδ�ɹ�������㷽ʽΪNULL�Ĳ���Ԥ����¼��
Create Or Replace Procedure Zl_�����շѽ���_Modify
(
  ��������_In     Number,
  ����id_In       ������ü�¼.����id%Type,
  ����id_In       ����Ԥ����¼.����id%Type,
  ���㷽ʽ_In     Varchar2,
  ��Ԥ��_In       ����Ԥ����¼.��Ԥ��%Type := Null,
  ��֧Ʊ��_In     ����Ԥ����¼.��Ԥ��%Type := Null,
  �����id_In     ����Ԥ����¼.�����id%Type := Null,
  ����_In         ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In   ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In     ����Ԥ����¼.����˵��%Type := Null,
  �ɿ�_In         ����Ԥ����¼.�ɿ�%Type := Null,
  �Ҳ�_In         ����Ԥ����¼.�Ҳ�%Type := Null,
  �����_In     ������ü�¼.ʵ�ս��%Type := Null,
  ��ɽ���_In     Number := 0,
  ȱʡ���㷽ʽ_In ���㷽ʽ.����%Type := Null,
  ���½������_In  Number := 0--�Ƿ������Ա��������Ҫ�Ǵ���ͳһ����Ա��¼��̨�����������
) As
  ------------------------------------------------------------------------------------------------------------------------------
  --����:�շѽ���ʱ,�޸Ľ���������Ϣ
  --��������_In:
  --   0-��ͨ�շѷ�ʽ:
  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.." ;Ҳ�������.
  --     ����֧Ʊ��_In:����漰��֧Ʊ,���뱾�ε���֧Ʊ��,�������շ�ʱ,������
  --   1.����������:
  --     �ٽ��㷽ʽ_IN:ֻ�ܴ���һ�����㷽ʽ,���������һЩ������Ϣ,��ʽΪ:"���㷽ʽ|������|�������|����ժҪ"
  --     ����֧Ʊ��_In:������
  --     �ۿ����ID_IN,����_IN,������ˮ��_IN,����˵��_In:��Ҫ����
  --   2-ҽ������(�������ҽ���Ľ���,��Ҫ��ɾ��ԭҽ������,���´���ĸ���)
  --     �ٽ��㷽ʽ_IN:��������,��ʽΪ:���㷽ʽ|������||.." 
  --     ����֧Ʊ��_In:������
  --   3-���ѿ�����:
  --     �ٽ��㷽ʽ_IN:����һ��ˢ���ſ�,��ʽΪ:�����ID|����|���ѿ�ID|���ѽ��||."  ���ѿ�ID:Ϊ��ʱ,���ݿ����Զ���λ
  --     �ڳ�Ԥ��_In: ������
  --     ����֧Ʊ��_In:������
  -- ��Ԥ��_In: ���ڳ�Ԥ��ʱ,����
  -- �����_In:��������ʱ,����
  -- ��ɽ���_In:1-����շ�;0-δ����շ�
  ------------------------------------------------------------------------------------------------------------------------------
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_�������� Varchar2(500);
  v_��ǰ���� Varchar2(50);
  v_����     ����ҽ�ƿ���Ϣ.����%Type;
  n_���ѿ�id ���ѿ�Ŀ¼.Id%Type;
  n_�����id ����Ԥ����¼.���㿨���%Type;
  v_����     Varchar2(100);
  n_���ƿ�   �����ѽӿ�Ŀ¼.���ƿ�%Type;
  n_���     ���˿������¼.���%Type;
  n_Id       ���˿������¼.Id%Type;
  n_Ԥ��id   ����Ԥ����¼.Id%Type;
  v_���㷽ʽ ����Ԥ����¼.���㷽ʽ%Type;
  n_������ ����Ԥ����¼.��Ԥ��%Type;
  n_����ֵ   ��Ա�ɿ����.���%Type;
  n_Ԥ����� ����Ԥ����¼.��Ԥ��%Type;
  n_��Ԥ��   ����Ԥ����¼.��Ԥ��%Type;
  v_��֧Ʊ   ����Ԥ����¼.���㷽ʽ%Type;
  v_������� ����Ԥ����¼.�������%Type;
  v_����ժҪ ����Ԥ����¼.ժҪ%Type;
  v_����   ���㷽ʽ.����%Type;
  n_Count    Number;
  n_Havenull Number;
  l_Ԥ��id   t_Numlist := t_Numlist();

  Cursor c_Feedata Is
    Select Max(m.����id) As ����id, Max(m.�Ǽ�ʱ��) As �Ǽ�ʱ��, Max(m.����Ա���) As ����Ա���, Max(m.����Ա����) As ����Ա����, Sum(���ʽ��) As ������,
           Max(m.�ɿ���id) As �ɿ���id
    From ������ü�¼ M
    Where m.����id = ����id_In;
  r_Feedata c_Feedata%RowType;

  Cursor c_Balancedata Is
    Select ��¼����, NO, ��¼״̬, ����id, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������
    From ����Ԥ����¼
    Where ����id = ����id_In And ���㷽ʽ Is Null;
  r_Balancedata c_Balancedata%RowType;

Begin
  Begin
    Select ���� Into v_���� From ���㷽ʽ Where Nvl(����, 0) = 9;
  Exception
    When Others Then
      v_���� := '����';
  End;

  --0.��ʽ����
  Select Count(1), Max(Decode(���㷽ʽ, Null, 1, 0))
  Into n_Count, n_Havenull
  From ����Ԥ����¼
  Where ����id = ����id_In;

  --1.���ӽ��㷽ʽΪ�յĽ�������
  n_������ := 0;
  n_Count    := 0;
  Open c_Feedata;
  Begin
    Fetch c_Feedata
      Into r_Feedata;
    --�������������㷽ʽΪnull�ļ�¼
    Select Nvl(Sum(��Ԥ��), 0) Into n_������ From ����Ԥ����¼ Where ����id = ����id_In;
    If Nvl(n_Havenull, 0) = 0 Or Round(Nvl(r_Feedata.������, 0), 6) <> Round(Nvl(n_������, 0), 6) Then
      --��ɾ�����ڵĽ��㷽ʽΪnull�ļ�¼
      Delete From ����Ԥ����¼ Where ����id = ����id_In And ���㷽ʽ Is Null;
      Select Nvl(Sum(��Ԥ��), 0) Into n_������ From ����Ԥ����¼ Where ����id = ����id_In;
    
      n_������ := Round(Nvl(r_Feedata.������, 0) - n_������, 6);
    
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, У�Ա�־, ��������)
      Values
        (����Ԥ����¼_Id.Nextval, 3, Null, 1, Decode(����id_In, 0, Null, ����id_In), Null, r_Feedata.�Ǽ�ʱ��, r_Feedata.����Ա���,
         r_Feedata.����Ա����, n_������, ����id_In, r_Feedata.�ɿ���id, -1 * ����id_In, 1, 3);
    End If;
  Exception
    When Others Then
      n_Count := 1;
  End;
  Close c_Feedata;
  If n_Count = 1 Then
    v_Err_Msg := 'δ�ҵ�ָ�����շ���ϸ����,�������ʧ�ܣ�';
    Raise Err_Item;
  End If;

  If ��������_In = 0 And Nvl(��֧Ʊ��_In, 0) <> 0 Then
    Begin
      Select b.����
      Into v_��֧Ʊ
      From ���㷽ʽӦ�� A, ���㷽ʽ B
      Where a.Ӧ�ó��� = '�շ�' And b.���� = a.���㷽ʽ And Nvl(b.Ӧ����, 0) = 1 And Rownum <= 1;
    Exception
      When Others Then
        v_��֧Ʊ := '��';
    End;
    If v_��֧Ʊ = '��' Then
      v_Err_Msg := '�ڽ��㳡����,�����ڽ�������ΪӦ����Ľ��㷽ʽ,����[���㷽ʽ]�����ã�';
      Raise Err_Item;
    End If;
  End If;

  Open c_Balancedata;
  Fetch c_Balancedata
    Into r_Balancedata;

  If Nvl(�����_In, 0) <> 0 Then
    Update ����Ԥ����¼
    Set ��Ԥ�� = Nvl(��Ԥ��, 0) + Nvl(�����_In, 0)
    Where ����id = ����id_In And ���㷽ʽ = v_����;
    If Sql%NotFound Then
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, У�Ա�־, �����id, ���㿨���, ����,
         ������ˮ��, ����˵��, �������, ��������)
      Values
        (����Ԥ����¼_Id.Nextval, 3, Null, 1, r_Balancedata.����id, Null, Null, v_����, r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա���,
         r_Balancedata.����Ա����, �����_In, r_Balancedata.����id, r_Balancedata.�ɿ���id, r_Balancedata.�������, 2, Null, Null, ����_In,
         ������ˮ��_In, ����˵��_In, Null, 3);
    End If;
    Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - Nvl(�����_In, 0) Where ����id = ����id_In And ���㷽ʽ Is Null;
    If Sql%NotFound Then
      v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[�շѽ��㴰��]�������շѣ�';
      Raise Err_Item;
    End If;
  End If;

  --Ԥ�����
  If Nvl(��Ԥ��_In, 0) <> 0 Then
    If Nvl(����id_In, 0) = 0 Then
      v_Err_Msg := '����ȷ�����˵Ĳ���ID,�շѲ���ʹ��Ԥ�������,�������ʧ�ܣ�';
      Raise Err_Item;
    End If;
  
    --���������
    Begin
      Select Nvl(Ԥ�����, 0) - Nvl(�������, 0)
      Into n_Ԥ�����
      From �������
      Where ����id = ����id_In And Nvl(����, 0) = 1 And ���� = 1;
    Exception
      When Others Then
        n_Ԥ����� := 0;
    End;
    If n_Ԥ����� < ��Ԥ��_In Then
      v_Err_Msg := '���˵ĵ�ǰԤ�����Ϊ ' || LTrim(To_Char(n_Ԥ�����, '9999999990.00')) || '��С�ڱ���֧����� ' ||
                   LTrim(To_Char(��Ԥ��_In, '9999999990.00')) || '��֧��ʧ�ܣ�';
      Raise Err_Item;
    End If;
  
    n_Ԥ����� := ��Ԥ��_In;
  
    For c_��Ԥ�� In (Select *
                  From (Select a.Id, a.��¼״̬, a.No, Nvl(a.���, 0) As ���
                         From ����Ԥ����¼ A,
                              (Select NO, Sum(Nvl(a.���, 0)) As ���
                                From ����Ԥ����¼ A
                                Where a.����id Is Null And Nvl(a.���, 0) <> 0 And a.����id = ����id_In And Ԥ����� = 1
                                Group By NO
                                Having Sum(Nvl(a.���, 0)) <> 0) B
                         Where a.����id Is Null And Nvl(a.���, 0) <> 0 And a.No = b.No And a.����id = ����id_In And a.Ԥ����� = 1
                         Union All
                         Select 0 As ID, ��¼״̬, NO, Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) As ���
                         From ����Ԥ����¼
                         Where ��¼���� In (1, 11) And ����id Is Not Null And Nvl(���, 0) <> Nvl(��Ԥ��, 0) And ����id = ����id_In And
                               Ԥ����� = 1 Having Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0
                         Group By ��¼״̬, NO)
                  Order By ID, NO) Loop
    
      If c_��Ԥ��.��� - n_Ԥ����� < 0 Then
        n_��Ԥ�� := c_��Ԥ��.���;
      Else
        n_��Ԥ�� := n_Ԥ�����;
      End If;
    
      If c_��Ԥ��.Id <> 0 Then
        --��һ�γ�Ԥ��(����һ�α��Ͻ���ID,��Ԥ�����Ϊ0)
        Update ����Ԥ����¼
        Set ��Ԥ�� = 0, ����id = ����id_In, ������� = -1 * ����id_In, �������� = 3
        Where ID = c_��Ԥ��.Id;
      End If;
      --���ϴ�ʣ���
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
         ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, ��������)
        Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�,
               r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա����, r_Balancedata.����Ա���, n_��Ԥ��, ����id_In, r_Balancedata.�ɿ���id, Ԥ�����,
               �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, -1 * ����id_In, 3
        From ����Ԥ����¼
        Where NO = c_��Ԥ��.No And ��¼״̬ = c_��Ԥ��.��¼״̬ And ��¼���� In (1, 11) And Rownum = 1;
    
      --��������(���㷽ʽΪNULL��)
      Update ����Ԥ����¼
      Set ��Ԥ�� = ��Ԥ�� - n_��Ԥ��
      Where ����id = ����id_In And ���㷽ʽ Is Null
      Returning Nvl(��Ԥ��, 0) Into n_����ֵ;
    
      --����Ƿ��Ѿ�������
      If c_��Ԥ��.��� < n_Ԥ����� Then
        n_Ԥ����� := n_Ԥ����� - c_��Ԥ��.���;
      Else
        n_Ԥ����� := 0;
      End If;
      If n_Ԥ����� = 0 Then
        Exit;
      End If;
    
    End Loop;
    --������Ƿ��㹻
    If Abs(n_Ԥ�����) > 0 Then
      v_Err_Msg := '���˵ĵ�ǰԤ�����С�ڱ���֧����� ' || LTrim(To_Char(��Ԥ��_In, '9999999990.00')) || '��֧��ʧ�ܣ�';
      Raise Err_Item;
    End If;
  
    --���²���Ԥ�����
    Update �������
    Set Ԥ����� = Nvl(Ԥ�����, 0) - ��Ԥ��_In
    Where ����id = ����id_In And ���� = 1 And ���� = 1
    Returning Ԥ����� Into n_����ֵ;
    If Sql%RowCount = 0 Then
      Insert Into ������� (����id, ����, Ԥ�����, ����) Values (����id_In, 1, -1 * ��Ԥ��_In, 1);
      n_����ֵ := -1 * ��Ԥ��_In;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ������� Where ����id = ����id_In And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
    End If;
  End If;
  If ��������_In = 0 Then
  
    If Nvl(��֧Ʊ��_In, 0) <> 0 Then
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, У�Ա�־, �����id, ���㿨���, ����,
         ������ˮ��, ����˵��, �������, ��������)
      Values
        (����Ԥ����¼_Id.Nextval, 3, Null, 1, r_Balancedata.����id, Null, Null, v_��֧Ʊ, r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա���,
         r_Balancedata.����Ա����, ��֧Ʊ��_In, r_Balancedata.����id, r_Balancedata.�ɿ���id, r_Balancedata.�������, 2, Null, Null, ����_In,
         ������ˮ��_In, ����˵��_In, Null, 3);
    
      Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - ��֧Ʊ��_In Where ����id = ����id_In And ���㷽ʽ Is Null;
      If Sql%NotFound Then
        v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[�շѽ��㴰��]�������շѣ�';
        Raise Err_Item;
      End If;
    End If;
  
    --�����շѽ��� :��ʽΪ:"���㷽ʽ|������|�������|����ժҪ||.."
    v_�������� := ���㷽ʽ_In || '||';
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
    
      v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
    
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
    
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      v_������� := LTrim(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_����ժҪ := LTrim(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
    
      If Nvl(n_������, 0) <> 0 Then
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, У�Ա�־, �����id, ���㿨���, ����,
           ������ˮ��, ����˵��, �������, ��������)
        Values
          (����Ԥ����¼_Id.Nextval, 3, Null, 1, r_Balancedata.����id, Null, v_����ժҪ, v_���㷽ʽ, r_Balancedata.�տ�ʱ��,
           r_Balancedata.����Ա���, r_Balancedata.����Ա����, n_������, r_Balancedata.����id, r_Balancedata.�ɿ���id, r_Balancedata.�������,
           2, Null, Null, ����_In, ������ˮ��_In, ����˵��_In, v_�������, 3);
      
        Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - n_������ Where ����id = ����id_In And ���㷽ʽ Is Null;
        If Sql%NotFound Then
          v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[�շѽ��㴰��]�������շѣ�';
          Raise Err_Item;
        End If;
      End If;
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;

  If ��������_In = 1 Then
    --���������㽻��
  
    v_��ǰ���� := ���㷽ʽ_In;
  
    v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
    v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
    n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
    v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
    v_������� := LTrim(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
    v_����ժҪ := LTrim(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
  
    If Nvl(n_������, 0) <> 0 Then
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, У�Ա�־, �����id, ���㿨���, ����,
         ������ˮ��, ����˵��, �������, ��������)
      Values
        (����Ԥ����¼_Id.Nextval, 3, Null, 1, r_Balancedata.����id, Null, v_����ժҪ, v_���㷽ʽ, r_Balancedata.�տ�ʱ��,
         r_Balancedata.����Ա���, r_Balancedata.����Ա����, n_������, r_Balancedata.����id, r_Balancedata.�ɿ���id, r_Balancedata.�������,
         2, �����id_In, Null, ����_In, ������ˮ��_In, ����˵��_In, v_�������, 3);
    
      Update ����Ԥ����¼ Set ��Ԥ�� = ��Ԥ�� - n_������ Where ����id = ����id_In And ���㷽ʽ Is Null;
      If Sql%NotFound Then
        v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[�շѽ��㴰��]�������շѣ�';
        Raise Err_Item;
      End If;
    End If;
  End If;

  --2.ҽ������(���ô˹���,��ȡƽ����̯�ķ�ʽ��̯�������):�������ҽ���ᴦ��,����ȫ��
  If ��������_In = 2 Then
    --2.1����Ƿ��Ѿ�����ҽ����������,������ɾ��
    n_������ := 0;
    For v_ҽ�� In (Select ID, ���㷽ʽ, ��Ԥ��
                 From ����Ԥ����¼ A
                 Where ����id = ����id_In And Exists
                  (Select 1 From ���㷽ʽ Where a.���㷽ʽ = ���� And ���� In (3, 4)) And Mod(��¼����, 10) <> 1)
    
     Loop
      n_������ := n_������ + Nvl(v_ҽ��.��Ԥ��, 0);
      l_Ԥ��id.Extend;
      l_Ԥ��id(l_Ԥ��id.Count) := v_ҽ��.Id;
    End Loop;
    If Nvl(n_������, 0) <> 0 Then
      Update ����Ԥ����¼
      Set ��Ԥ�� = Nvl(��Ԥ��, 0) + Nvl(n_������, 0)
      Where ����id = ����id_In And ���㷽ʽ Is Null;
      If Sql%NotFound Then
        v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[�շѽ��㴰��]�������շѣ�';
        Raise Err_Item;
      End If;
    End If;
    If l_Ԥ��id.Count <> 0 Then
      Forall I In 1 .. l_Ԥ��id.Count
        Delete ����Ԥ����¼ Where ID = l_Ԥ��id(I);
    End If;
  
    If ���㷽ʽ_In Is Not Null Then
      v_�������� := ���㷽ʽ_In || '||';
    End If;
  
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
      n_������ := To_Number(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
    
      For c_������Ϣ In (Select ��¼����, NO, ��¼״̬, ����id, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������
                     From ����Ԥ����¼
                     Where ����id = ����id_In And ���㷽ʽ Is Null) Loop
      
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, У�Ա�־, ��������)
        Values
          (����Ԥ����¼_Id.Nextval, 3, Null, 1, c_������Ϣ.����id, Null, '���ս���', v_���㷽ʽ, c_������Ϣ.�տ�ʱ��, c_������Ϣ.����Ա���, c_������Ϣ.����Ա����,
           n_������, c_������Ϣ.����id, c_������Ϣ.�ɿ���id, c_������Ϣ.�������, 1, 3);
      End Loop;
      --��������(���㷽ʽΪNULL��)
      Update ����Ԥ����¼
      Set ��Ԥ�� = ��Ԥ�� - n_������
      Where ����id = ����id_In And ���㷽ʽ Is Null
      Returning Nvl(��Ԥ��, 0) Into n_����ֵ;
    
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;

  --3-���ѿ���������
  If ��������_In = 3 Then
    v_�������� := ���㷽ʽ_In || '||';
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      --�����ID|����|���ѿ�ID|���ѽ��
      n_�����id := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      v_����     := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_���ѿ�id := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_������ := To_Number(v_��ǰ����);
      Begin
        Select ����, ���ƿ�, ���㷽ʽ Into v_����, n_���ƿ�, v_���㷽ʽ From �����ѽӿ�Ŀ¼ Where ��� = �����id_In;
      Exception
        When Others Then
          v_���� := Null;
      End;
      If v_���� Is Null Then
        v_Err_Msg := 'δ�ҵ���Ӧ�Ľ��㿨�ӿ�,����ˢ������ʧ��!';
        Raise Err_Item;
      End If;
      If v_���㷽ʽ Is Null Then
        v_Err_Msg := v_���� || 'δ���ö�Ӧ�Ľ��㷽ʽ,����ˢ������ʧ��!';
        Raise Err_Item;
      End If;
      If Nvl(n_���ѿ�id, 0) = 0 Then
        --δ�������ѿ�IDʱ,�Կ���Ϊ׼���в���(���ŵĺϷ���,�ڳ��������ж�)
        Begin
          Select ID
          Into n_���ѿ�id
          From ���ѿ�Ŀ¼
          Where �ӿڱ�� = n_�����id And ���� = v_���� And
                ��� = (Select Max(���) From ���ѿ�Ŀ¼ Where �ӿڱ�� = n_�����id And ���� = v_����);
        Exception
          When Others Then
            n_���ѿ�id := 0;
        End;
        If Nvl(n_���ѿ�id, 0) = 0 Then
          v_Err_Msg := 'δ�ҵ�����Ϊ:' || v_���� || '��' || v_���� || '.,����ˢ������ʧ��!';
          Raise Err_Item;
        End If;
      End If;
    
      If Nvl(n_������, 0) <> 0 Then
      
        Update ����Ԥ����¼
        Set ��Ԥ�� = Nvl(��Ԥ��, 0) + n_������
        Where ����id = r_Balancedata. ����id And ���㷽ʽ = v_���㷽ʽ And ���㿨��� = n_�����id
        Returning ID Into n_Ԥ��id;
        If Sql%NotFound Then
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          Insert Into ����Ԥ����¼
            (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, ���㿨���, У�Ա�־, ��������)
          Values
            (n_Ԥ��id, 3, Null, 1, r_Balancedata. ����id, Null, Null, v_���㷽ʽ, r_Balancedata. �տ�ʱ��, r_Balancedata. ����Ա���,
             r_Balancedata. ����Ա����, n_������, r_Balancedata. ����id, r_Balancedata. �ɿ���id, r_Balancedata. �������, n_�����id, 2, 3);
        End If;
      
        --���뿨�����¼
        Begin
          Select Nvl(Max(Nvl(���, 0)), 0) + 1
          Into n_���
          From ���˿������¼
          Where �ӿڱ�� = n_�����id And Nvl(���ѿ�id, 0) = Nvl(n_���ѿ�id, 0) And ���� = v_����;
        Exception
          When Others Then
            n_��� := 1;
        End;
      
        Select ���˿������¼_Id.Nextval Into n_Id From Dual;
      
        Insert Into ���˿������¼
          (ID, �ӿڱ��, ���ѿ�id, ���, ��¼״̬, ���㷽ʽ, ������, ����, ������ˮ��, ����ʱ��, ��ע, �����־)
        Values
          (n_Id, n_�����id, n_���ѿ�id, n_���, 1, v_���㷽ʽ, n_������, v_����, Null, r_Balancedata. �տ�ʱ��, Null, 0);
        --������ѿ�,��ͬʱ���������
        If Nvl(n_���ѿ�id, 0) <> 0 Then
          Update ���ѿ�Ŀ¼ Set ��� = ��� - n_������ Where ID = n_���ѿ�id;
          If Sql%NotFound Then
            v_Err_Msg := '����Ϊ' || v_���� || '��' || v_���� || 'δ�ҵ�!';
            Raise Err_Item;
          End If;
        End If;
        Insert Into ���˿�������� (Ԥ��id, ������id) Values (n_Ԥ��id, n_Id);
      
        --��������(���㷽ʽΪNULL��)
        Update ����Ԥ����¼
        Set ��Ԥ�� = ��Ԥ�� - n_������
        Where ����id = r_Balancedata. ����id And ���㷽ʽ Is Null And Nvl(У�Ա�־, 0) = 1
        Returning Nvl(��Ԥ��, 0) Into n_����ֵ;
      
      End If;
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;

  If Nvl(��ɽ���_In, 0) = 0 Then
    Return;
  End If;

  -----------------------------------------------------------------------------------------
  --����շ�,��Ҫ������Ա�ɿ����,Ԥ����¼(���㷽ʽ=NULL)

  --1.ɾ�����㷽ʽΪNULL��Ԥ����¼
  Delete ����Ԥ����¼ Where ����id = ����id_In And ���㷽ʽ Is Null And Nvl(��Ԥ��, 0) = 0;
  If Sql%NotFound Then
    Select Count(*) Into n_Count From ����Ԥ����¼ A Where ����id = ����id_In And ���㷽ʽ Is Null;
    If n_Count <> 0 Then
      v_Err_Msg := '������δ�ɿ������,������ɽ���!';
    Else
      v_Err_Msg := '������Ϣ����,������Ϊ����ԭ����ɽ�����Ϣ����,����[�շѽ��㴰��]�������շѣ�!';
    End If;
    Raise Err_Item;
  End If;

  --������Ϊ��ʱ������һ�����Ϊ0�Ĳ���Ԥ����¼
  Select Count(*) Into n_Count From ����Ԥ����¼ A Where ����id = ����id_In;

  If n_Count = 0 Then
    v_���㷽ʽ := ȱʡ���㷽ʽ_In;
    If v_���㷽ʽ Is Null Then
      Begin
        Select ���㷽ʽ Into v_���㷽ʽ From ���㷽ʽӦ�� Where Ӧ�ó��� = '�շ�' And Nvl(ȱʡ��־, 0) = 1;
      Exception
        When Others Then
          v_���㷽ʽ := Null;
      End;
      If v_���㷽ʽ Is Null Then
        Begin
          Select ���� Into v_���㷽ʽ From ���㷽ʽ Where Nvl(����, 0) = 1;
        Exception
          When Others Then
            v_���㷽ʽ := '�ֽ�';
        End;
      End If;
    End If;
    Insert Into ����Ԥ����¼
      (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, У�Ա�־, �����id, ���㿨���, ����,
       ������ˮ��, ����˵��, �������, ��������)
    Values
      (����Ԥ����¼_Id.Nextval, 3, Null, 1, r_Balancedata.����id, Null, Null, v_���㷽ʽ, r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա���,
       r_Balancedata.����Ա����, 0, r_Balancedata.����id, r_Balancedata.�ɿ���id, r_Balancedata.�������, 2, Null, Null, Null, Null,
       ����˵��_In, Null, 3);
  End If;

  --2.����ɿ����ݺ��Ҳ����ݼ�У�Ա�־����Ϊ0
  Update ����Ԥ����¼ Set �ɿ� = �ɿ�_In, �Ҳ� = �Ҳ�_In, У�Ա�־ = 0 Where ����id = ����id_In;

  --3.���·���״̬
  Update ������ü�¼ Set ����״̬ = 0 Where ����id = ����id_In;

  --4.������Ա�ɿ�����
  If Nvl(���½������_In,0)=0 then
    For c_�ɿ� In (Select ���㷽ʽ, ����Ա����, Sum(Nvl(a.��Ԥ��, 0)) As ��Ԥ��
                 From ����Ԥ����¼ A
                 Where a.����id = ����id_In And Mod(a.��¼����, 10) <> 1
                 Group By ���㷽ʽ, ����Ա����) Loop
    
      Update ��Ա�ɿ����
      Set ��� = Nvl(���, 0) + Nvl(c_�ɿ�.��Ԥ��, 0)
      Where �տ�Ա = c_�ɿ�.����Ա���� And ���� = 1 And ���㷽ʽ = c_�ɿ�.���㷽ʽ;
      If Sql%RowCount = 0 Then
        Insert Into ��Ա�ɿ����
          (�տ�Ա, ���㷽ʽ, ����, ���)
        Values
          (c_�ɿ�.����Ա����, c_�ɿ�.���㷽ʽ, 1, Nvl(c_�ɿ�.��Ԥ��, 0));
      End If;
    End Loop;
  End if;
  --�շѺ��������
  Begin
    Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
      Using 4, ����id_In;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����շѽ���_Modify;
/

--93333:������,2016-02-24,ȡ���������Ʒ���
Create Or Replace Procedure Zl_����걾��¼_�걾����
(
  Id_In         In ����걾��¼.Id%Type,
  ҽ��id_In     In ����걾��¼.ҽ��id%Type,
  ���ҽ��_In   In Varchar2, --���ڸ��¶��ҽ����ִ��״̬
  ���Ǳ걾id_In In ����걾��¼.Id%Type := 0, --����ʱָ����һ���걾ʱ����ָ��ı걾
  �걾���_In   In ����걾��¼.�걾���%Type,
  ����ʱ��_In   In ����걾��¼.����ʱ��%Type,
  ������_In     In ����걾��¼.������%Type,
  ����id_In     In ����걾��¼.����id%Type,
  ����ʱ��_In   In ����걾��¼.����ʱ��%Type,
  �걾��̬_In   In ����걾��¼.�걾��̬%Type,
  ������_In     In ����걾��¼.������%Type := Null,
  ����ʱ��_In   In ����걾��¼.����ʱ��%Type := Null,
  ΢����걾_In In ����걾��¼.΢����걾%Type := Null,
  �걾���_In   In ����걾��¼.�걾���%Type := 0,
  ���鱸ע_In   In ����걾��¼.���鱸ע%Type := Null,
  ����_In       In ����걾��¼.����%Type := Null,
  �Ա�_In       In ����걾��¼.�Ա�%Type := Null,
  ����_In       In ����걾��¼.����%Type := Null,
  No_In         In ����걾��¼.No%Type := Null,
  �걾����_In   In ����걾��¼.�걾����%Type := Null,
  �������id_In In ����걾��¼.�������id%Type := Null,
  ������_In     In ����걾��¼.������%Type := Null,
  ��ʶ��_In     In ����걾��¼.��ʶ��%Type := Null,
  ����_In       In ����걾��¼.����%Type := Null,
  ���˿���_In   In ����걾��¼.���˿���%Type := Null,
  ������Ŀ_In   In ����걾��¼.������Ŀ%Type := Null,
  ��������_In   In ����걾��¼.��������%Type := Null,
  ����id_In     In ����걾��¼.����id%Type := Null,
  ִ�п���_In   In ����걾��¼.ִ�п���id%Type := Null,
  ��Ա���_In   In ��Ա��.���%Type := Null,
  ��Ա����_In   In ��Ա��.����%Type := Null
) Is

  Cursor v_Advice Is
    Select /*+ Rule */
    Distinct a.Id, a.����ʱ��, a.�걾��λ, f.��������, a.ִ�п���id, a.������Ŀid, a.��������id, a.����ҽ��, a.����id, a.������Դ, a.Ӥ��, a.������־ As ����,
             b.�����, b.סԺ��, b.��������, a.�Һŵ�, Decode(c.��ҳid, 0, Null, c.��ҳid) As ��ҳid, d.��������, f.������, f.����ʱ��
    From ����ҽ����¼ A, ����ҽ������ F, ������Ϣ B, ������ҳ C, ������ĿĿ¼ D
    Where a.���id In (Select * From Table(Cast(f_Num2list(���ҽ��_In) As Zltools.t_Numlist))) And a.Id = f.ҽ��id And
          a.����id = b.����id And a.����id = c.����id(+) And a.��ҳid = c.��ҳid(+) And a.������Ŀid = d.Id(+);

  Cursor v_Advice_1 Is
    Select /*+ Rule */
    Distinct b.No As ���ݺ�
    From ����ҽ����¼ A, ����ҽ������ B
    Where a.Id = b.ҽ��id And a.���id In (Select * From Table(Cast(f_Num2list(���ҽ��_In) As Zltools.t_Numlist)))
    Union All
    Select /*+ Rule */
    Distinct b.No As ���ݺ�
    From ����ҽ����¼ A, ����ҽ������ B
    Where a.Id = b.ҽ��id And a.Id In (Select * From Table(Cast(f_Num2list(���ҽ��_In) As Zltools.t_Numlist)));

  Cursor v_Patient Is
    Select ����id, סԺ��, �����, �������� From ������Ϣ Where ����id = ����id_In;

  --δ��˵ķ�����(������ҩƷ)
  Cursor c_Verify(v_ҽ��id In Number) Is
    Select /*+ Rule */
    Distinct a.��¼����, a.No, a.���, a.ҽ�����, a.�����־
    From סԺ���ü�¼ A, ����ҽ������ B,
         (Select ID
           From ����ҽ����¼
           Where ID In (Select * From Table(Cast(f_Num2list(���ҽ��_In) As Zltools.t_Numlist)))
           Union All
           Select ID
           From ����ҽ����¼
           Where ���id In (Select * From Table(Cast(f_Num2list(���ҽ��_In) As Zltools.t_Numlist)))) C
    Where a.�շ���� Not In ('5', '6', '7') And a.ҽ����� = c.Id And a.��¼״̬ = 0 And �۸񸸺� Is Null And a.ҽ����� = b.ҽ��id And
          a.��¼���� = b.��¼���� And a.No = b.No And a.���ʷ��� = 1
    Union All
    Select /*+ Rule */
    Distinct a.��¼����, a.No, a.���, a.ҽ�����, a.�����־
    From סԺ���ü�¼ A, ����ҽ������ B,
         (Select ID
           From ����ҽ����¼
           Where ID In (Select * From Table(Cast(f_Num2list(���ҽ��_In) As Zltools.t_Numlist)))
           Union All
           Select ID
           From ����ҽ����¼
           Where ���id In (Select * From Table(Cast(f_Num2list(���ҽ��_In) As Zltools.t_Numlist)))) C
    Where a.�շ���� Not In ('5', '6', '7') And a.ҽ����� = c.Id And a.��¼״̬ = 0 And �۸񸸺� Is Null And a.ҽ����� = b.ҽ��id And
          a.��¼���� = b.��¼���� And a.No = b.No And a.���ʷ��� = 1
    Union All
    Select /*+ Rule */
    Distinct a.��¼����, a.No, a.���, a.ҽ�����, a.�����־
    From ������ü�¼ A, ����ҽ������ B,
         (Select ID
           From ����ҽ����¼
           Where ID In (Select * From Table(Cast(f_Num2list(���ҽ��_In) As Zltools.t_Numlist)))
           Union All
           Select ID
           From ����ҽ����¼
           Where ���id In (Select * From Table(Cast(f_Num2list(���ҽ��_In) As Zltools.t_Numlist)))) C
    Where a.�շ���� Not In ('5', '6', '7') And a.ҽ����� = c.Id And a.��¼״̬ = 0 And �۸񸸺� Is Null And a.ҽ����� = b.ҽ��id And
          a.��¼���� = b.��¼���� And a.No = b.No And a.���ʷ��� = 1
    Union All
    Select /*+ Rule */
    Distinct a.��¼����, a.No, a.���, a.ҽ�����, a.�����־
    From ������ü�¼ A, ����ҽ������ B,
         (Select ID
           From ����ҽ����¼
           Where ID In (Select * From Table(Cast(f_Num2list(���ҽ��_In) As Zltools.t_Numlist)))
           Union All
           Select ID
           From ����ҽ����¼
           Where ���id In (Select * From Table(Cast(f_Num2list(���ҽ��_In) As Zltools.t_Numlist)))) C
    Where a.�շ���� Not In ('5', '6', '7') And a.ҽ����� = c.Id And a.��¼״̬ = 0 And �۸񸸺� Is Null And a.ҽ����� = b.ҽ��id And
          a.��¼���� = b.��¼���� And a.No = b.No And a.���ʷ��� = 1
    Order By ��¼����, NO, ���;

  --���ҵ�ǰ�걾���������
  Cursor c_Samplequest(v_΢���� In Number) Is
    Select Distinct ҽ��id, ������Դ
    From (Select Decode(a.ҽ��id, Null, b.ҽ��id, a.ҽ��id) As ҽ��id, b.������Դ
           From ������Ŀ�ֲ� A, ����걾��¼ B
           Where Nvl(v_΢����, 0) = 0 And a.�걾id = b.Id And b.ҽ��id In (Select ҽ��id From ����걾��¼ Where ID = Id_In) And
                 a.ҽ��id Is Not Null
           Union
           Select Decode(a.ҽ��id, Null, b.ҽ��id, a.ҽ��id) As ҽ��id, b.������Դ
           From ������Ŀ�ֲ� A, ����걾��¼ B
           Where 1 = v_΢���� And b.Id = a.�걾id And b.Id = Id_In
           Union
           Select b.Id As ҽ��id, b.������Դ
           From ����걾��¼ A, ����ҽ����¼ B
           Where a.Id = Id_In And a.ҽ��id In (b.Id, b.���id));

  Cursor c_Stuff
  (
    v_No     Varchar2,
    v_��ҳid Number
  ) Is
    Select NO As ���ݺ�, ����, �ⷿid
    From δ��ҩƷ��¼
    Where NO = v_No And ���� In (24, 25, 26) And �ⷿid Is Not Null And Not Exists
     (Select 1 From Dual Where Zl_Getsysparameter(Decode(v_��ҳid, Null, 92, 63)) = '1') And Exists
     (Select a.���
           From סԺ���ü�¼ A, �������� B
           Where a.��¼���� = 2 And a.No = v_No And a.�շ�ϸĿid = b.����id And b.�������� = 1
           Union All
           Select a.���
           From ������ü�¼ A, �������� B
           Where a.��¼���� = 2 And a.No = v_No And a.�շ�ϸĿid = b.����id And b.�������� = 1)
    Order By �ⷿid;

  r_Advice   v_Advice%Rowtype;
  r_Advice_1 v_Advice_1%Rowtype;
  r_Patient  v_Patient%Rowtype;

  Err_Custom Exception;
  v_Error Varchar2(1000);
  v_Flag  Number(18);

  v_Temp      Varchar2(255);
  v_Seq       Number;
  v_Union     Number;
  v_Patientid Number;
  v_Itemid    Number;
  v_Count     Number;
  v_ִ��      Number;
  v_No        ����ҽ������.No%Type;
  v_����      ����ҽ������.��¼����%Type;
  v_���      Varchar2(1000);
  v_��ҳid    Number(18);
  v_�����־  סԺ���ü�¼.�����־%Type;
  n_Count     Number;
  v_����      ����ҽ����¼.����%Type;
  v_�Ա�      ����ҽ����¼.�Ա�%Type;
  v_����      ����ҽ����¼.����%Type;
  v_������Դ  ����ҽ����¼.������Դ%Type;
  v_Ӥ��      ����ҽ����¼.Ӥ��%Type;
  v_Ӥ������  ����ҽ����¼.����%Type;
  v_Ӥ���Ա�  ����ҽ����¼.�Ա�%Type;
Begin

  If Nvl(���Ǳ걾id_In, 0) > 0 Then
    Begin
      Select ���� Into v_Temp From ����걾��¼ Where ID = ���Ǳ걾id_In And ���� Is Null;
    Exception
      When Others Then
        v_Error := 'ָ�����ǵı걾�ѱ����ջ���ɾ����������ָ����';
        Raise Err_Custom;
    End;
  End If;

  If Nvl(ҽ��id_In, 0) > 0 Then
    Select ����, �Ա�, ����, ������Դ, Ӥ��
    Into v_����, v_�Ա�, v_����, v_������Դ, v_Ӥ��
    From ����ҽ����¼
    Where ID = ҽ��id_In;
  
    If v_������Դ <> 3 Then
      If Nvl(v_Ӥ��, 0) = 0 Then
        If v_���� <> ����_In Or v_�Ա� <> �Ա�_In  Then
          v_Error := '�����������Ա������ҽ���������ܱ��棬������޸Ĳ�����Ϣ���ٽ��б��棡';
          Raise Err_Custom;
        End If;
      Else
        Select b.Ӥ������, b.Ӥ���Ա�
        Into v_Ӥ������, v_Ӥ���Ա�
        From ����ҽ����¼ A, ������������¼ B
        Where a.����id = b.����id And a.��ҳid = b.��ҳid And a.Ӥ�� = b.��� And
              a.���id In (Select * From Table(Cast(f_Num2list(���ҽ��_In) As Zltools.t_Numlist))) And Rownum = 1;
      
        If v_Ӥ������ <> ����_In Or v_Ӥ���Ա� <> �Ա�_In Then
          v_Error := '�����������Ա������ҽ���������ܱ��棬������޸Ĳ�����Ϣ���ٽ��б��棡';
          Raise Err_Custom;
        End If;
      End If;
    End If;
  
    Select Count(ID) Into v_Flag From ����걾��¼ Where ҽ��id = ҽ��id_In And ID <> Id_In;
    If v_Flag > 0 Then
      Select Count(Distinct b.������Ŀid)
      Into v_Flag
      From ����ҽ����¼ A, ���鱨����Ŀ B
      Where a.������Ŀid = b.������Ŀid And a.���id In (Select * From Table(Cast(f_Num2list(���ҽ��_In) As Zltools.t_Numlist)));
    
      Select Count(a.��Ŀid)
      Into n_Count
      From ������Ŀ�ֲ� A
      Where a.ҽ��id In (Select * From Table(Cast(f_Num2list(���ҽ��_In) As Zltools.t_Numlist))) And a.�걾id <> Id_In;
      If (v_Flag - n_Count) <= 0 Then
        v_Error := '��ǰҽ���ѱ����գ������ظ����գ�';
        Raise Err_Custom;
      End If;
    End If;
  End If;

  If ҽ��id_In = 0 Then
    Open v_Patient;
    Fetch v_Patient
      Into r_Patient;
  
    If v_Patient%Found Then
      Zl_������Ϣ_�������(r_Patient.����id);
    End If;
  
    Update ����걾��¼
    Set ����ʱ�� = Decode(����ʱ��_In, Null, ����ʱ��, ����ʱ��_In), ������ = Decode(������_In, Null, ������, ������_In), �걾���� = Nvl(�걾����_In, �걾����),
        ����ʱ�� = ����ʱ��_In, ���� = Decode(����_In, Null, ����, ����_In), �Ա� = Decode(�Ա�_In, Null, �Ա�, �Ա�_In),
        ���� = Decode(����_In, Null, ����, ����_In), �������� = Decode(����_In, Null, Null, Zl_Val(����_In)),
        ���䵥λ = Decode(����_In, Null, ���䵥λ,
                       Decode(����_In, Null, Null, '����', '����', 'Ӥ��', 'Ӥ��',
                               Decode(Sign(Instr(����_In, '��')), 1, '��',
                                       Decode(Sign(Instr(����_In, '��')), 1, '��',
                                               Decode(Sign(Instr(����_In, '��')), 1, '��',
                                                       Decode(Sign(Instr(����_In, 'Сʱ')), 1, 'Сʱ', Null)))))),
        �������id = Decode(�������id_In, Null, �������id, �������id_In), ������ = Decode(������_In, Null, ������, ������_In),
        �걾��̬ = Decode(�걾��̬_In, Null, �걾��̬, �걾��̬_In), ��ʶ�� = Decode(��ʶ��_In, Null, ��ʶ��, ��ʶ��_In),
        ���� = Decode(����_In, Null, ����, ����_In), ���˿��� = Decode(���˿���_In, Null, ���˿���, ���˿���_In),
        ������Ŀ = Decode(������Ŀ_In, Null, ������Ŀ, ������Ŀ_In), ����id = Decode(����id_In, Null, ����id, ����id_In),
        ҽ��id = Decode(ҽ��id_In, Null, ҽ��id, 0, ҽ��id, ҽ��id_In)
    Where ID = Id_In;
    If Sql%NotFound Then
      Insert Into ����걾��¼
        (ID, ҽ��id, �걾���, ����ʱ��, ������, �걾����, ������, ����ʱ��, ����״̬, ��������, ����id, ��������, ����ʱ��, �걾��̬, ������, ִ�п���id, ������, ����ʱ��, ΢����걾,
         �걾���, ���鱸ע, �������id, ������, ����, �Ա�, ����, ��������, ���䵥λ, ����id, ������Դ, Ӥ��, NO, �ϲ�id, ��ʶ��, ����, ���˿���, ����, �����, סԺ��, ��������,
         �Һŵ�, ��ҳid, ������Ŀ, ��������, ������, ����ʱ��)
      Values
        (Id_In, Decode(ҽ��id_In, 0, Null, ҽ��id_In), �걾���_In, ����ʱ��_In, ������_In, �걾����_In, ��Ա����_In, ����ʱ��_In, 1, ��������_In,
         Decode(����id_In, 0, Null, ����id_In), Null, Null, �걾��̬_In, 0, ִ�п���_In, ������_In, ����ʱ��_In, ΢����걾_In, �걾���_In, ���鱸ע_In,
         �������id_In, ������_In, ����_In, �Ա�_In, ����_In, Zl_Val(����_In),
         Decode(����_In, Null, Null, '����', '����', 'Ӥ��', 'Ӥ��',
                 Decode(Sign(Instr(����_In, '��')), 1, '��',
                         Decode(Sign(Instr(����_In, '��')), 1, '��',
                                 Decode(Sign(Instr(����_In, '��')), 1, '��', Decode(Sign(Instr(����_In, 'Сʱ')), 1, 'Сʱ', Null))))),
         ����id_In, Decode(r_Patient.סԺ��, Null, Decode(r_Patient.�����, Null, 3, 1), 2), 0, Null, Null, ��ʶ��_In, ����_In,
         ���˿���_In, �걾���_In, r_Patient.�����, r_Patient.סԺ��, r_Patient.��������, Null, Null, ������Ŀ_In, Null, Null, Null);
    End If;
    If Nvl(���Ǳ걾id_In, 0) > 0 Then
      Zl_����걾��¼_Union(Id_In, ���Ǳ걾id_In);
    End If;
    --��¼���պͲ������
    Insert Into ���������¼
      (ID, �걾id, ��������, ����Ա, ����ʱ��)
    Values
      (���������¼_Id.Nextval, Id_In, 2, ��Ա����_In, Sysdate);
    Close v_Patient;
  Else
    Open v_Advice;
    Fetch v_Advice
      Into r_Advice;
  
    If v_Advice%Found Then
      Zl_������Ϣ_�������(r_Advice.����id);
    End If;
  
    Update ����걾��¼
    Set ҽ��id = Decode(ҽ��id_In, Null, ҽ��id, 0, ҽ��id, ҽ��id_In), ����ʱ�� = Decode(����ʱ��_In, Null, ����ʱ��, ����ʱ��_In),
        ������ = Decode(������_In, Null, ������, ������_In), �걾��� = Decode(�걾���_In, Null, �걾���, �걾���_In),
        �걾���� = Decode(�걾����_In, Null, Decode(�걾����, Null, r_Advice.�걾��λ, �걾����), �걾����_In),
        ����ʱ�� = Decode(r_Advice.����ʱ��, Null, ����ʱ��, r_Advice.����ʱ��), ������ = Decode(������, Null, ��Ա����_In, ������),
        �������� = Decode(r_Advice.��������, Null, ��������, r_Advice.��������), �������� = Decode(��������_In, Null, ��������, ��������_In),
        ִ�п���id = Decode(ִ�п���_In, Null, ִ�п���id, ִ�п���_In), ������ = Decode(������_In, Null, ������, ������_In),
        ����ʱ�� = Decode(����ʱ��_In, Null, ����ʱ��, ����ʱ��_In), ���鱸ע = Decode(���鱸ע_In, Null, ���鱸ע, ���鱸ע_In),
        �������id = Decode(�������id_In, Null, �������id, �������id_In), ������ = Decode(������_In, Null, ������, ������_In),
        ���� = Decode(����_In, Null, ����, ����_In), �Ա� = Decode(�Ա�_In, Null, �Ա�, �Ա�_In), ���� = Decode(����_In, Null, ����, ����_In),
        �������� = Decode(����_In, Null, ��������, Zl_Val(����_In)),
        ���䵥λ = Decode(����_In, Null, ���䵥λ,
                       Decode(����_In, Null, Null, '����', '����', 'Ӥ��', 'Ӥ��',
                               Decode(Sign(Instr(����_In, '��')), 1, '��',
                                       Decode(Sign(Instr(����_In, '��')), 1, '��',
                                               Decode(Sign(Instr(����_In, '��')), 1, '��',
                                                       Decode(Sign(Instr(����_In, 'Сʱ')), 1, 'Сʱ', Null)))))),
        ����id = Decode(r_Advice.����id, Null, ����id, r_Advice.����id), ������Դ = Decode(r_Advice.������Դ, Null, ������Դ, r_Advice.������Դ),
        Ӥ�� = Decode(r_Advice.Ӥ��, Ӥ��, r_Advice.Ӥ��), NO = Decode(No_In, Null, NO, No_In), �ϲ�id = v_Union,
        �걾��̬ = Decode(�걾��̬_In, Null, �걾��̬, �걾��̬_In), ��ʶ�� = Decode(��ʶ��_In, Null, ��ʶ��, ��ʶ��_In),
        ���� = Decode(����_In, Null, ����, ����_In), ���˿��� = Decode(���˿���_In, Null, ���˿���, ���˿���_In), �걾��� = �걾���_In,
        ����� = r_Advice.�����, סԺ�� = r_Advice.סԺ��, �������� = r_Advice.��������, �Һŵ� = r_Advice.�Һŵ�, ��ҳid = r_Advice.��ҳid,
        ������Ŀ = Decode(������Ŀ_In, Null, ������Ŀ, ������Ŀ_In), �������� = r_Advice.��������, ������ = r_Advice.������, ����ʱ�� = r_Advice.����ʱ��
    Where ID = Id_In;
  
    If Sql%NotFound Then
      Insert Into ����걾��¼
        (ID, ҽ��id, �걾���, ����ʱ��, ������, �걾����, ������, ����ʱ��, ����״̬, ��������, ����id, ��������, ����ʱ��, �걾��̬, ������, ִ�п���id, ������, ����ʱ��, ΢����걾,
         �걾���, ���鱸ע, �������id, ������, ����, �Ա�, ����, ��������, ���䵥λ, ����id, ������Դ, Ӥ��, NO, �ϲ�id, ��ʶ��, ����, ���˿���, ����, �����, סԺ��, ��������,
         �Һŵ�, ��ҳid, ������Ŀ, ��������, ������, ����ʱ��)
      Values
        (Id_In, Decode(ҽ��id_In, 0, Null, ҽ��id_In), �걾���_In, ����ʱ��_In, ������_In, Nvl(�걾����_In, r_Advice.�걾��λ), ��Ա����_In,
         ����ʱ��_In, 1, ��������_In, Decode(����id_In, 0, Null, ����id_In), r_Advice.��������, r_Advice.����ʱ��, �걾��̬_In, 0, ִ�п���_In,
         ������_In, ����ʱ��_In, ΢����걾_In, �걾���_In, ���鱸ע_In, �������id_In, ������_In, ����_In, �Ա�_In, ����_In, Zl_Val(����_In),
         Decode(����_In, Null, Null, '����', '����', 'Ӥ��', 'Ӥ��',
                 Decode(Sign(Instr(����_In, '��')), 1, '��',
                         Decode(Sign(Instr(����_In, '��')), 1, '��',
                                 Decode(Sign(Instr(����_In, '��')), 1, '��', Decode(Sign(Instr(����_In, 'Сʱ')), 1, 'Сʱ', Null))))),
         r_Advice.����id, r_Advice.������Դ, r_Advice.Ӥ��, No_In, v_Union, ��ʶ��_In, ����_In, ���˿���_In, r_Advice.����, r_Advice.�����,
         r_Advice.סԺ��, r_Advice.��������, r_Advice.�Һŵ�, r_Advice.��ҳid, ������Ŀ_In, r_Advice.��������, r_Advice.������, r_Advice.����ʱ��);
    End If;
    If Nvl(���Ǳ걾id_In, 0) > 0 Then
      Zl_����걾��¼_Union(Id_In, ���Ǳ걾id_In);
    End If;
    Insert Into ���������¼
      (ID, �걾id, ��������, ����Ա, ����ʱ��)
    Values
      (���������¼_Id.Nextval, Id_In, 2, ��Ա����_In, Sysdate);
  
    --��������Ŀ��ʱ��д�ϲ�ID
    Begin
      Select a.Id
      Into v_Union
      From ����걾��¼ A, ����걾��¼ B, ����ҽ����¼ C, ����ϲ����� D, ����ҽ����¼ E
      Where a.����id = b.����id And b.Id = Id_In And a.����״̬ = 1 And Nvl(a.����id, 0) <> 0 And a.ҽ��id = c.���id And
            d.����Ŀid = c.������Ŀid And d.�ϲ���Ŀid = e.������Ŀid And e.Id = r_Advice.Id And Rownum = 1
      Order By a.����ʱ�� Desc;
    Exception
      When Others Then
        v_Union := Null;
    End;
    If Nvl(v_Union, 0) <> 0 Then
      Update ����걾��¼ Set �ϲ�id = v_Union Where (ID = Id_In Or ҽ��id = r_Advice.Id);
    End If;
    --������������Ŀʱ��д�ϲ���Ŀ
    Begin
      Select a.Id, a.����id, c.����Ŀid
      Into v_Union, v_Patientid, v_Itemid
      From ����걾��¼ A, ����ҽ����¼ B, ����ϲ����� C
      Where a.ҽ��id = b.���id And b.������Ŀid = c.����Ŀid And a.Id = Id_In And Rownum = 1;
    Exception
      When Others Then
        v_Union := Null;
    End;
    If Nvl(v_Union, 0) <> 0 Then
      Update ����걾��¼
      Set �ϲ�id = v_Union
      Where ID In (Select a.Id
                   From ����걾��¼ A, ����ҽ����¼ B, ����ϲ����� C
                   Where a.ҽ��id = b.���id And b.������Ŀid = c.�ϲ���Ŀid And c.����Ŀid = v_Itemid And a.����id = v_Patientid And
                         a.����״̬ = 1);
    End If;
  
    v_Seq := 1;
    Close v_Advice;
    v_Flag := 0;
    Begin
      Select Nvl(Max(1), 0) Into v_Flag From ����������Ŀ Where �걾id = Id_In;
    Exception
      When Others Then
        v_Flag := 0;
    End;
    If v_Flag = 0 Then
      For r_Advice In v_Advice Loop
        Update ����������Ŀ
        Set �걾id = Id_In, ������Ŀid = r_Advice.������Ŀid
        Where �걾id = Id_In And ������Ŀid = r_Advice.������Ŀid;
        If Sql%Rowcount = 0 Then
          Insert Into ����������Ŀ (�걾id, ������Ŀid, ���) Values (Id_In, r_Advice.������Ŀid, v_Seq);
        End If;
        v_Seq := v_Seq + 1;
      End Loop;
    End If;
  
  End If;

  --���ݲ������ж��Ƿ���
    For r_Advice_1 In v_Advice_1 Loop
      --�������û���Զ�����,���Զ�����,���򲻴���
      For r_Stuff In c_Stuff(r_Advice_1.���ݺ�, v_��ҳid) Loop
      
        Zl_�����շ���¼_��������(r_Stuff.�ⷿid, r_Stuff.����, r_Stuff.���ݺ�, ��Ա����_In, ��Ա����_In, ��Ա����_In, 1, Sysdate);
      End Loop;
    End Loop;


  Update /*+ Rule */ ����ҽ������
  Set ִ��״̬ = 3
  Where ִ��״̬ = 0 And
        ҽ��id In (Select ID
                 From ����ҽ����¼
                 Where ID In (Select * From Table(Cast(f_Num2list(���ҽ��_In) As Zltools.t_Numlist)))
                 Union All
                 Select ID
                 From ����ҽ����¼
                 Where ���id In (Select * From Table(Cast(f_Num2list(���ҽ��_In) As Zltools.t_Numlist))));

  --ִ�к��Զ���˶�Ӧ�ļ��ʻ��۵�(������ҩƷ)
  Select Zl_To_Number(Nvl(Zl_Getsysparameter(81), '0')) Into v_ִ�� From Dual;
  --2.��鵱ǰ�걾��ص��������ر걾�Ƿ�������
  For r_Samplequest In c_Samplequest(΢����걾_In) Loop
  
    v_Count := 0;
  
    --r_SampleQuest.ҽ��id�����Ѿ����,�����������
    If v_Count = 0 Then
    
      If r_Samplequest.������Դ = 2 Then
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
      If Nvl(v_ִ��, 0) = 1 Then
        For r_Verify In c_Verify(r_Samplequest.ҽ��id) Loop
          If r_Verify.No || ',' || r_Verify.��¼���� <> v_No || ',' || v_���� Then
            If v_��� Is Not Null Then
              If r_Verify.�����־ = 1 Then
                Zl_������ʼ�¼_Verify(v_No, ��Ա���_In, ��Ա����_In, Substr(v_���, 2));
              Elsif r_Verify.�����־ = 2 Then
                Zl_סԺ���ʼ�¼_Verify(v_No, ��Ա���_In, ��Ա����_In, Substr(v_���, 2));
              End If;
            End If;
            v_��� := Null;
          End If;
          v_�����־ := r_Verify.�����־;
          v_No       := r_Verify.No;
          v_����     := r_Verify.��¼����;
          v_���     := v_��� || ',' || r_Verify.���;
        End Loop;
        If v_��� Is Not Null Then
          If v_�����־ = 1 Then
            Zl_������ʼ�¼_Verify(v_No, ��Ա���_In, ��Ա����_In, Substr(v_���, 2));
          Elsif v_�����־ = 2 Then
            Zl_סԺ���ʼ�¼_Verify(v_No, ��Ա���_In, ��Ա����_In, Substr(v_���, 2));
          End If;
        End If;
      End If;
    
    End If;
  End Loop;

  If Nvl(��������_In, 0) = 1 Then
    Zl_����ҽ����¼_���δ�ӡ(ҽ��id_In, 1);
  End If;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����걾��¼_�걾����;
/

--93144:���ϴ�,2016-02-29,Ԥ���˿�����
CREATE OR REPLACE Procedure Zl_����Ԥ����¼_Insert
(
  Id_In         ����Ԥ����¼.Id%Type,
  ���ݺ�_In     ����Ԥ����¼.No%Type,
  Ʊ�ݺ�_In     Ʊ��ʹ����ϸ.����%Type,
  ����id_In     ����Ԥ����¼.����id%Type,
  ��ҳid_In     ����Ԥ����¼.��ҳid%Type,
  ����id_In     ����Ԥ����¼.����id%Type,
  ���_In       ����Ԥ����¼.���%Type,
  ���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type,
  �������_In   ����Ԥ����¼.�������%Type,
  �ɿλ_In   ����Ԥ����¼.�ɿλ%Type,
  ��λ������_In ����Ԥ����¼.��λ������%Type,
  ��λ�ʺ�_In   ����Ԥ����¼.��λ�ʺ�%Type,
  ժҪ_In       ����Ԥ����¼.ժҪ%Type,
  ����Ա���_In ����Ԥ����¼.����Ա���%Type,
  ����Ա����_In ����Ԥ����¼.����Ա����%Type,
  ����id_In     Ʊ��ʹ����ϸ.����id%Type,
  Ԥ�����_In   ����Ԥ����¼.Ԥ�����%Type := Null,
  �����id_In   ����Ԥ����¼.�����id%Type := Null,
  ���㿨���_In ����Ԥ����¼.���㿨���%Type := Null,
  ����_In       ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
  ������λ_In   ����Ԥ����¼.������λ%Type := Null,
  �տ�ʱ��_In   ����Ԥ����¼.�տ�ʱ��%Type := Null,
  ��������_In   Integer := 0,
  ��������_In   ����Ԥ����¼.��������%Type := Null,
  ���½������_In  Number := 0,--�Ƿ������Ա��������Ҫ�Ǵ���ͳһ����Ա��¼��̨�������������
  �˿ʽ_In   Number := 0
) As
  ----------------------------------------------
  --��������_In:0-������Ԥ��;1-��Ϊ���۵�;3-����˿�
  --�˿ʽ_In;0-��ʾ;1-��ֹ��2-����
  v_Err_Msg Varchar2(200);
  n_Err_Num Number;
  Err_Item Exception;

  v_����   ���㷽ʽ.����%Type;
  v_��ӡid Ʊ�ݴ�ӡ����.Id%Type;
  v_����   ������Ϣ.��������%Type;
  v_Date   Date;
  n_����ֵ �������.Ԥ�����%Type;
  n_��id   ����ɿ����.Id%Type;
  n_������� �������.Ԥ�����%Type;
  n_����Ԥ�� �������.Ԥ�����%Type;
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
  
  If ���_In<0 then
    Begin
      Select Nvl(Ԥ�����,0)-Nvl(�������,0) into n_������� From ������� Where ���� = 1 And ����id = ����id_In And Nvl(����, 0) = Nvl(Ԥ�����_In, 0);
    Exception
      When Others Then
        Null;
    End;
    --����˿�Ҫ��������Ԥ���Ƿ�֧������
    IF ��������_In = 3 then
      For c_����Ԥ�� In (Select A.Ԥ��ID,A.Ԥ�����,A.�����ID,A.���㿨��� as ���ѽӿ�ID,nvl(B.����,C.���) as ����,nvl(B.����,C.����) as ����,
                              Decode(B.����,NULL,C.�Ƿ�ȫ��,B.�Ƿ�ȫ��) as �Ƿ�ȫ��,Decode(B.����,NULL,C.�Ƿ�����,B.�Ƿ�����) as �Ƿ�����,
                              A.����,A.������ˮ��,A.����˵��,          A.Ԥ�����
                              From (   Select A.Ԥ�����,nvl(A.�����ID,0) as �����ID,nvl(A.���㿨���,0) as ���㿨���,
                              A.����,A.������ˮ��,A.����˵��,           max(decode(sign(���),-1, decode(A.��¼״̬,1,0,2,0,ID),ID)) as Ԥ��ID,
                              nvl(sum(���),0)-nvl(sum(nvl(��Ԥ��,0)),0) as Ԥ�����    From ����Ԥ����¼ A    Where   A.����ID=����id_In and (nvl(A.���㿨���,0)<>0 or nvl(�����ID,0)<>0)
                              Group by A.Ԥ�����,nvl(A.�����ID,0),nvl(A.���㿨���,0),A.����,A.������ˮ��,A.����˵��
                              Having nvl(sum(���),0)-nvl(sum(nvl(��Ԥ��,0)),0)  <>0 ) A,ҽ�ƿ���� B,�����ѽӿ�Ŀ¼ C
                              Where A.Ԥ����� =Nvl(Ԥ�����_In, 0) And  A.�����ID=B.ID(+)  and A.���㿨���=C.���(+)  and nvl(A.Ԥ�����,0)<>0   Order by ����,A.����,A.������ˮ��,A.����˵��) Loop

        IF instr(',7,8,',','|| v_���� || ',') =0 And Nvl(c_����Ԥ��.�Ƿ�����,0) = 0 And Nvl(c_����Ԥ��.Ԥ�����,0) > 0 Then
          n_����Ԥ��:= Nvl(n_����Ԥ��,0) + Nvl(c_����Ԥ��.Ԥ�����,0);
        ElsIf instr(',7,8,',','|| v_���� || ',') >0 Then
              If Nvl(c_����Ԥ��.����,'0') = Nvl(����_In,'0') And Nvl(c_����Ԥ��.������ˮ��,'0')= Nvl(������ˮ��_In,'0') And Nvl(c_����Ԥ��.����˵��,'0') = Nvl(����˵��_In,'0') then
                n_����Ԥ��:= Nvl(n_����Ԥ��,0) + Nvl(c_����Ԥ��.Ԥ�����,0);
              End if;
        End IF;
      End Loop;
    End if;
    
    If instr(',7,8,',','|| v_���� || ',') > 0 And Nvl(n_����Ԥ��,0) < 0 And ��������_In = 3 Then
        n_Err_Num := -20101;
        v_Err_Msg :='�˿�����ڲ�������Ԥ�������ܼ�����';
        Raise Err_Item;
    ElsIf Nvl(n_�������,0) < 0 And �˿ʽ_In <> 2 Then
        n_Err_Num := -20101;
        v_Err_Msg :='�˿�����ڲ���ʣ��Ԥ�������ܼ�����';
        If �˿ʽ_In = 0 Then
          n_Err_Num := -20111;
          v_Err_Msg :='�˿�����ڲ���ʣ��Ԥ�����Ƿ���ԣ�';
        End If;
        Raise Err_Item;
    ElsIf instr(',7,8,',','|| v_���� || ',') = 0 And Nvl(n_�������,0) - Nvl(n_����Ԥ��,0) < 0 And ��������_In = 3 And �˿ʽ_In <> 2 then
        n_Err_Num := -20101;
        v_Err_Msg :='�˿�����ڲ���ʣ��Ԥ�������ܼ�����';
        If �˿ʽ_In = 0 Then
          n_Err_Num := -20111;
          v_Err_Msg :='�˿�����ڲ���ʣ��Ԥ�����Ƿ���ԣ�';
        End If;
        Raise Err_Item;
    End if;
  End if;

  --��Ա�ɿ����(����)
  If Nvl(���½������_In,0)=0 then
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) + ���_In
    Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = ���㷽ʽ_In
    Returning ��� Into n_����ֵ;

    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ���� (�տ�Ա, ���㷽ʽ, ����, ���) Values (����Ա����_In, ���㷽ʽ_In, 1, ���_In);
      n_����ֵ := ���_In;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ��Ա�ɿ���� Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = ���㷽ʽ_In And Nvl(���, 0) = 0;
    End If;
  End if;
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
    Raise_Application_Error(n_Err_Num, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Ԥ����¼_Insert;
/

--94030:������,2016-03-16,Σ��ֵ��Ϣ�Ķ�
--92384:������,2016-02-24,ҽ����ִ����Ϣ
CREATE OR REPLACE Procedure Zl_ҵ����Ϣ�嵥_Read
(
  ����id_In     In ҵ����Ϣ�嵥.����id%Type,
  ����id_In     In ҵ����Ϣ�嵥.����id%Type,
  ���ͱ���_In   In ҵ����Ϣ�嵥.���ͱ���%Type,
  �Ķ�����_In   In ҵ����Ϣ״̬.�Ķ�����%Type,
  �Ķ���_In     In ҵ����Ϣ״̬.�Ķ���%Type,
  �Ķ�����id_In In ҵ����Ϣ״̬.�Ķ�����id%Type,
  �Ķ�ʱ��_In   In ҵ����Ϣ״̬.�Ķ�ʱ��%Type := Null,
  ��Ϣid_In     In ҵ����Ϣ״̬.��Ϣid%Type := Null,
  ҵ���ʶ_In   In ҵ����Ϣ�嵥.ҵ���ʶ %Type := Null
) Is
  d_Cur   Date;
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  If �Ķ�ʱ��_In Is Null Then
    Select Sysdate Into d_Cur From Dual;
  Else
    d_Cur := �Ķ�ʱ��_In;
  End If;
  If Nvl(��Ϣid_In, 0) <> 0 Then
    Insert Into ҵ����Ϣ״̬
      (��Ϣid, �Ķ�����, �Ķ���, �Ķ�ʱ��, �Ķ�����id)
    Values
      (��Ϣid_In, �Ķ�����_In, �Ķ���_In, d_Cur, �Ķ�����id_In);
    Update ҵ����Ϣ�嵥 Set �Ƿ����� = 1 Where ID = ��Ϣid_In;
  Elsif ҵ���ʶ_In Is Not Null Then
    For R In (Select a.Id
              From ҵ����Ϣ�嵥 A
              Where a.����id = ����id_In And a.����id = ����id_In And a.���ͱ��� = ���ͱ���_In And a.ҵ���ʶ = ҵ���ʶ_In And
                    Nvl(a.�Ƿ�����, 0) = 0) Loop
      Insert Into ҵ����Ϣ״̬
        (��Ϣid, �Ķ�����, �Ķ���, �Ķ�ʱ��, �Ķ�����id)
      Values
        (r.Id, �Ķ�����_In, �Ķ���_In, d_Cur, �Ķ�����id_In);
    End Loop;
    Update ҵ����Ϣ�嵥
    Set �Ƿ����� = 1
    Where ����id = ����id_In And ����id = ����id_In And ���ͱ��� = ���ͱ���_In And Nvl(�Ƿ�����, 0) = 0 And ҵ���ʶ = ҵ���ʶ_In;
  Elsif ���ͱ���_In = 'ZLHIS_CIS_034' Then
    --��ִ����Ϣ���⴦��
    For R In (Select ID
              From ҵ����Ϣ�嵥 A, ҵ����Ϣ���Ѳ��� B
              Where a.Id = b.��Ϣid And a.����id = ����id_In And a.����id = ����id_In And a.���ͱ��� = ���ͱ���_In And Nvl(a.�Ƿ�����, 0) = 0 And
                    b.����id = �Ķ�����id_In
              Group By ID) Loop
      Insert Into ҵ����Ϣ״̬
        (��Ϣid, �Ķ�����, �Ķ���, �Ķ�ʱ��, �Ķ�����id)
      Values
        (r.Id, �Ķ�����_In, �Ķ���_In, d_Cur, �Ķ�����id_In);
      Update ҵ����Ϣ�嵥 Set �Ƿ����� = 1 Where ID = r.Id;
    End Loop;
  Else
    For R In (Select a.Id
              From ҵ����Ϣ�嵥 A
              Where a.����id = ����id_In And a.����id = ����id_In And a.���ͱ��� = ���ͱ���_In And Nvl(a.�Ƿ�����, 0) = 0) Loop
      Insert Into ҵ����Ϣ״̬
        (��Ϣid, �Ķ�����, �Ķ���, �Ķ�ʱ��, �Ķ�����id)
      Values
        (r.Id, �Ķ�����_In, �Ķ���_In, d_Cur, �Ķ�����id_In);
    End Loop;
    Update ҵ����Ϣ�嵥
    Set �Ƿ����� = 1
    Where ����id = ����id_In And ����id = ����id_In And ���ͱ��� = ���ͱ���_In And Nvl(�Ƿ�����, 0) = 0;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҵ����Ϣ�嵥_Read;
/

--91316:�ŵ���,2015-02-23,��Һ������������
--91954:�ŵ���,2016-03-15,�������������Զ�����
CREATE OR REPLACE Procedure Zl_��Һ��ҩ��¼_�˲�
(
  ����id_In   In ��Һ��ҩ��¼.����id%Type,
  ҽ��id_In   In Varchar2, --��Һҽ����ҩ;����Ӧ��ҽ��ID:ҽ��ID1,ҽ��ID2...
  ���ͺ�_In   In ����ҽ������.���ͺ�%Type,
  �˲���_In   In ��Һ��ҩ״̬.������Ա%Type,
  �˲�ʱ��_In In ��Һ��ҩ״̬.����ʱ��%Type
) Is
  v_Count    Number;
  v_���     Number;
  v_ִ��ʱ�� Date;

  v_���id      Number;
  v_New���id   Number;
  v_Old���id   Number;
  v_���ͺ�      Number;
  v_Tmp         Varchar2(200);
  I             Number;
  v_��ҩid      Number;
  v_����        Number;
  v_Maxno       Varchar2(4000);
  v_Lableno     Varchar2(200);
  v_Maxbatch    Number;
  v_Curdose     Number;
  v_Sumdose     Number;
  v_Drugcount   Number;
  v_Currdate    Date;
  n_Needcheck   Number;
  n_Lngid       ҩƷ�շ���¼.Id%Type;
  n_Count       Number(3);
  n_����        ҩƷ�շ���¼.����%Type;
  v_No          ҩƷ�շ���¼.No%Type;
  n_���ʹ���    Number(5);
  n_����id      ������Ϣ.����id%Type := 0;
  b_Change      Boolean;
  n_Sum         Number(8);
  n_��������    Number(1);
  n_Cur         Number(5);
  v_�ϴη��ͺ�  ����ҽ������.���ͺ�%Type;
  v_ҽ��ids     Varchar2(4000);
  v_Tansid      Varchar2(12);
  v_��ǰ����    Varchar2(20);
  n_Num         Number(8);
  d_Oldִ��ʱ�� Date;
  n_�Ƿ���    Number(1);
  n_���        Number(1);
  n_��ҩ��      Number(2);
  --���Ʋ���
  v_ҽ������       Number;
  v_��Һ����       Number;
  v_����Һ����     Varchar2(2000);
  v_����Һ��ҩ;�� Varchar2(2000);
  v_��Դ����       Varchar2(4000);
  v_Continue       Number := 1;
  v_Nodosage       Number := 0;
  v_�����ϴ�����   Number := 0;
  d_�ֹ����ʱ��   Date;
  n_Tpn���÷�ʽ    Number := 0;
  v_ҩƷ����       varchar2(20);
  n_���ҩƷ����   number(1);
  n_����ҩƷ����   number(1);
  n_���ȼ�         number:=999;
  n_�Զ�����       number:=0;
  n_����id       number:=0;
  n_row            number(2);
  Err_Item Exception;

  Cursor c_ҽ����¼ Is
    Select /*+rule */
    Distinct e.ҽ��id As ���id, e.���ͺ�, b.Ƶ�ʼ��, b.�����λ, b.ִ��ʱ�䷽��, a.����, a.�Ա�, a.����, a.סԺ��, a.��ǰ���� As ����, a.��ǰ����id As ���˲���id,
             a.��ǰ����id As ���˿���id, e.�״�ʱ��, e.ĩ��ʱ��, b.��ʼִ��ʱ��, Nvl(e.��������, 0) As ����, e.����ʱ��, Decode(b.ҽ����Ч, 0, 1, 2) As ҽ������,
             b.������Ŀid As ��ҩ;��, b.����id, Nvl(c.ִ�б��, 0) As �Ƿ�tpn
    From ����ҽ������ E, ����ҽ����¼ B, ������Ϣ A, ������ĿĿ¼ C, Table(f_Num2list(ҽ��id_In)) D
    Where e.ҽ��id = b.Id And b.����id = a.����id And  c.��� = 'E' And c.�������� = '2' And c.ִ�з��� = 1 And b.������Ŀid = c.Id And
          e.ҽ��id = d.Column_Value And e.���ͺ� = ���ͺ�_In
    Order By b.����id, e.ҽ��id, e.���ͺ�;

  Cursor c_����ҽ����¼ Is
    Select /*+rule */
    Distinct e.ҽ��id As ���id, e.���ͺ�, b.Ƶ�ʼ��, b.�����λ, b.ִ��ʱ�䷽��, a.����, a.�Ա�, a.����, a.סԺ��, a.��ǰ���� As ����, a.��ǰ����id As ���˲���id,
             a.��ǰ����id As ���˿���id, e.�״�ʱ��, e.ĩ��ʱ��, b.��ʼִ��ʱ��, Nvl(e.��������, 0) As ����, e.����ʱ��, Decode(b.ҽ����Ч, 0, 1, 2) As ҽ������,
             b.������Ŀid As ��ҩ;��, b.����id
    From ����ҽ������ E, ����ҽ����¼ B, ������Ϣ A, ������ĿĿ¼ C, Table(f_Num2list(ҽ��id_In)) D
    Where e.ҽ��id = b.Id And b.����id = a.����id And c.��� = 'E' And c.�������� = '2' And c.ִ�з��� = 1 And b.������Ŀid = c.Id And
          e.ҽ��id = d.Column_Value And e.���ͺ� = ���ͺ�_In And b.����id = n_����id
    Order By e.ҽ��id, e.���ͺ�;

  Cursor c_�շ���¼ Is
    Select Distinct c.Id As �շ�id, c.���, c.ʵ������ As ����, Nvl(e.�Ƿ�������, 0) As �Ƿ�������, c.����, c.No,F.������,F.�Ƿ�����ҩ
    From ����ҽ����¼ A, ����ҽ������ B, ҩƷ�շ���¼ C, סԺ���ü�¼ D, ��ҺҩƷ���� E,ҩƷ���� F,ҩƷ��� G
    Where a.Id = b.ҽ��id And c.����id = d.Id And a.Id = d.ҽ����� And b.No = c.No And c.ҩƷid=G.ҩƷID and G.ҩ��ID=F.ҩ��ID And c.ҩƷid = e.ҩƷid(+) And
          b.ִ�в���id + 0 = ����id_In And c.���� = 9 And c.������� Is Null And a.���id = v_���id And b.���ͺ� = ���ͺ�_In And c.��� < 1000
    Order By c.No, c.���;

  v_ҽ����¼     c_ҽ����¼%RowType;
  v_�շ���¼     c_�շ���¼%RowType;
  v_����ҽ����¼ c_����ҽ����¼%RowType;

  Function Zl_Getpivaworkbatch(ִ��ʱ��_In In Date,ҩƷ����_In In varchar2:=null) Return Number As
    v_Exetime   Date;
    v_Starttime Date;
    v_Endtime   Date;
    v_Maxbatch  Number(2);
    v_Batch     Number;
    Cursor c_��ҩ���� Is
      Select ����, ��ҩʱ��, ��ҩʱ��, ���,ҩƷ���� From ��ҩ�������� Where ���� = 1 and ��������id=����id_In Order By ����;

    v_��ҩ���� c_��ҩ����%RowType;
  Begin
    v_Exetime := To_Date(Substr(To_Char(ִ��ʱ��_In, 'yyyy-mm-dd hh24:mi'), 12), 'hh24:mi');

    Select Max(Nvl(����, 0)) + 1 Into v_Maxbatch From ��ҩ�������� Where ���� = 1 and ��������id=����id_In;

    For v_��ҩ���� In c_��ҩ���� Loop
      v_Batch     := 0;

      if ҩƷ����_In is null then
        if v_��ҩ����.����<>'0'and v_��ҩ����.ҩƷ���� is null then
          v_Starttime := To_Date(Substr(v_��ҩ����.��ҩʱ��, 1, Instr(v_��ҩ����.��ҩʱ��, '-') - 1), 'hh24:mi');
          v_Endtime   := To_Date(Substr(v_��ҩ����.��ҩʱ��, Instr(v_��ҩ����.��ҩʱ��, '-') + 1), 'hh24:mi');

          If v_Exetime >= v_Starttime And v_Exetime <= v_Endtime Then
            v_Batch := v_��ҩ����.����;
            n_���  := v_��ҩ����.���;
            Exit When v_Batch > 0;
          End If;
        end if;
      else
        if ҩƷ����_In=v_��ҩ����.ҩƷ���� then
          v_Batch := v_��ҩ����.����;
          n_���  := v_��ҩ����.���;
          Exit When v_Batch > 0;
        end if;
      end if;
    End Loop;

    If v_Batch = 0 and n_���ҩƷ����<>1 Then
      v_Batch := v_Maxbatch;
    End If;
    Return(v_Batch);
  End;

  Function Zl_GetFirst(��ҩid_In In number,����id_In In number) Return Number As
    n_First     Number;
    n_����id    number;
    Cursor c_���ȼ� Is
      select ����id,��ҩ����,���ȼ�,Ƶ�� from ��ҺҩƷ���ȼ� where (����id=����id_In or ����id=0) order by ����id,���ȼ� desc;

    r_���ȼ� c_���ȼ�%RowType;
  Begin
   n_First:=0;
   for r_���ȼ� in c_���ȼ� loop
     if n_����id<>0 and r_���ȼ�.����id=0 then
       exit;
     end if;
     n_����id:=r_���ȼ�.����id;

     for r_��ҩ��¼ in(select distinct D.��ҩ����,E.ִ��Ƶ�� from ��Һ��ҩ��¼ A,��Һ��ҩ���� B,ҩƷ�շ���¼ C,��ҺҩƷ���� D,����ҽ����¼ E  where A.ҽ��id=E.Id and A.id=B.��¼ID and B.�շ�ID=C.id and C.ҩƷID=D.ҩƷID and a.id=��ҩid_In) loop
       if instr(r_��ҩ��¼.��ҩ����,r_���ȼ�.��ҩ����,1)>0 and instr(r_���ȼ�.Ƶ��,r_��ҩ��¼.ִ��Ƶ��,1)>0 then
         n_First:= r_���ȼ�.���ȼ�;
         exit;
       end if;
     end loop;
   end loop;

   if n_First=0 then
     n_First:=999;
   end if;
   Return(n_First);
  End;
Begin
  n_Count          := 0;
  v_ҽ������       := Zl_To_Number(Nvl(zl_GetSysParameter('ҽ������', 1345), 1));
  v_��Һ����       := Zl_To_Number(Nvl(zl_GetSysParameter('ͬ������Һ����', 1345), 0));
  v_����Һ����     := Nvl(zl_GetSysParameter('����ҺҩƷ����', 1345), '');
  v_����Һ��ҩ;�� := Nvl(zl_GetSysParameter('��Һ��ҩ;��', 1345), '');
  v_��Դ����       := Nvl(zl_GetSysParameter('��Դ����', 1345), '');
  v_�����ϴ�����   := Zl_To_Number(Nvl(zl_GetSysParameter('�����ϴ�����', 1345), 0));
  n_Tpn���÷�ʽ    := Zl_To_Number(Nvl(zl_GetSysParameter('����Ӫ��ҩ�ﴦ�÷�ʽ', 1345), 0));
  n_���ҩƷ����   := Zl_To_Number(Nvl(zl_GetSysParameter('����ҩƷ����������ҩƷ�����ݸ�ҩʱ��û����ҩ���ε���Һ��Ĭ��Ϊ0���β����', 1345), 0));
  n_����ҩƷ����   := Zl_To_Number(Nvl(zl_GetSysParameter('����ҩƷ��ҩƷ����ָ������', 1345), 0));
  n_�Զ�����:= Zl_To_Number(Nvl(zl_GetSysParameter('�����Զ�����', 1345), 0));
  v_ҽ��ids  := ҽ��id_In;
  v_��ǰ���� := '';

  Select Trunc(Sysdate) Into v_Currdate From Dual;

  Select Max(Nvl(����, 0)) + 1 Into v_Maxbatch From ��ҩ��������;

  --��鵱ǰ���˵�ҽ���Ƿ��н�����Ҫִ�е���Һ��������״̬��
  If Instr(v_ҽ��ids, ',') = 0 Then
    v_Tansid := v_ҽ��ids;
  Else
    v_Tansid := Substr(v_Tmp, 1, Instr(v_ҽ��ids, ',') - 1);
  End If;

  Select Count(ID)
  Into n_Num
  From ��Һ��ҩ��¼
  Where �Ƿ����� = 1 And ִ��ʱ�� Between Trunc(v_Currdate) And Trunc(v_Currdate + 1) - 1 / 24 / 60 / 60 And
        ҽ��id In
        (Select ���id
         From ����ҽ����¼
         Where ����id = (Select ����id From ����ҽ����¼ Where ���id = v_Tansid And Rownum < 2) And (������� = '5' Or ������� = '6')) And
        Rownum < 2;

  If n_Num > 0 Then
    Select ����
    Into v_��ǰ����
    From ��Һ��ҩ��¼
    Where �Ƿ����� = 1 And ִ��ʱ�� Between Trunc(v_Currdate) And Trunc(v_Currdate + 1) - 1 / 24 / 60 / 60 And
          ҽ��id In
          (Select ���id
           From ����ҽ����¼
           Where ����id = (Select ����id From ����ҽ����¼ Where ���id = v_Tansid And Rownum < 2) And (������� = '5' Or ������� = '6')) And
          Rownum < 2;
    Raise Err_Item;
  End If;

  --�Ƚ�ԭ�շ���¼����������µ��շ���¼��������ɾ��
  --Update ҩƷ�շ���¼
  --Set ��� = ��� + 10000
  --Where ID In (Select \*+rule *\
  --             Distinct c.Id
  --             From ����ҽ����¼ A, ����ҽ������ B, ҩƷ�շ���¼ C, סԺ���ü�¼ D, Table(f_Num2list(ҽ��id_In)) F
  --             Where a.Id = b.ҽ��id And c.����id = d.Id And a.Id = d.ҽ����� And b.No = c.No And b.ִ�в���id + 0 = ����id_In And
  --                   c.���� = 9 And c.������� Is Null And a.���id = f.Column_Value And b.���ͺ� = ���ͺ�_In And c.��� < 10000);

  For v_ҽ����¼ In c_ҽ����¼ Loop
    v_Continue := 1;
    n_����id := v_ҽ����¼.����id;
    n_����id:= v_ҽ����¼.���˿���id;

    Select Count(1)
    Into v_Continue
    From ����ҽ����¼ A, ��Һ������ҩƷ B,סԺ���ü�¼ C
    Where c.�շ�ϸĿid = b.ҩƷid And c.ҽ����� =A.id and A.���id= v_ҽ����¼.���id;
    If v_Continue = 0 Then
      v_Continue := 1;
    Else
      v_Continue := 0;
    End If;

    --�������Ʋ�����Һ��
    If (v_ҽ������ = 1 And v_ҽ����¼.ҽ������ <> 1) Or (v_ҽ������ = 2 And v_ҽ����¼.ҽ������ <> 2) Then
      v_Continue := 0;
    End If;

    If Not v_����Һ��ҩ;�� Is Null Then
      If Instr(',' || v_����Һ��ҩ;�� || ',', ',' || v_ҽ����¼.��ҩ;�� || ',') = 0 Then
        v_Continue := 0;
      End If;
    End If;

    If Not v_��Դ���� Is Null Then
      If Instr(',' || v_��Դ���� || ',', ',' || v_ҽ����¼.���˿���id || ',') = 0 Then
        v_Continue := 0;
      End If;
    End If;

    v_ҩƷ����:=null;
    for r_ҩƷ���� in (Select decode(nvl(D.�Ƿ�����ҩ,0),0,'','����ҩ') ҩƷ����
    From ����ҽ����¼ A, ҩƷ��� B,סԺ���ü�¼ C,ҩƷ���� D
    Where c.�շ�ϸĿid = b.ҩƷid And B.ҩ��ID=D.ҩ��ID And c.ҽ����� =A.id and A.���id= v_ҽ����¼.���id) loop
      if r_ҩƷ����.ҩƷ���� is not null then
        v_ҩƷ����:=r_ҩƷ����.ҩƷ����;
      end if;
    end loop;

    if v_ҩƷ���� is null then
       If v_ҽ����¼.�Ƿ�tpn = 2 Then
        v_ҩƷ����:='Ӫ��ҩ';
        v_Continue := 1;
      end if;
    end if;

    If v_Continue = 1 Then
      v_Old���id := v_New���id;
      v_���id    := v_ҽ����¼.���id;
      v_New���id := v_���id;
      v_���ͺ�    := v_ҽ����¼.���ͺ�;
      v_���      := 0;


      If v_Continue = 1 Then
        --v_Count := Zl_Gettransexenumber(v_ҽ����¼.��ʼִ��ʱ��, v_ҽ����¼.�״�ʱ��, v_ҽ����¼.ĩ��ʱ��, v_ҽ����¼.Ƶ�ʼ��, v_ҽ����¼.�����λ, v_ҽ����¼.ִ��ʱ�䷽��);
        Select Count(ҽ��id)
        Into v_Count
        From ҽ��ִ��ʱ��
        Where ҽ��id = v_ҽ����¼.���id And ���ͺ� = v_ҽ����¼.���ͺ�;

        v_Nodosage := 0;

        For I In 1 .. v_Count Loop
          Select ��Һ��ҩ��¼_Id.Nextval Into v_��ҩid From Dual;
          v_��� := v_��� + 1;

          If I > 1 Then
            --��ҽ��ִ��ʱ�����ȡҽ����ִ��ʱ��
            --v_ִ��ʱ�� := Zl_Gettransexetime(v_ҽ����¼.��ʼִ��ʱ��, v_ִ��ʱ��, v_ҽ����¼.Ƶ�ʼ��, v_ҽ����¼.�����λ, v_ҽ����¼.ִ��ʱ�䷽��);
            Select Ҫ��ʱ��
            Into v_ִ��ʱ��
            From ҽ��ִ��ʱ��
            Where ҽ��id = v_ҽ����¼.���id And ���ͺ� = v_ҽ����¼.���ͺ� And Ҫ��ʱ�� > v_ִ��ʱ�� And Rownum = 1
            Order By Ҫ��ʱ��;
          Else
            Select Min(Ҫ��ʱ��)
            Into v_ִ��ʱ��
            From ҽ��ִ��ʱ��
            Where ҽ��id = v_ҽ����¼.���id And ���ͺ� = v_ҽ����¼.���ͺ� And Rownum = 1
            Order By Ҫ��ʱ��;
          End If;

          v_���� := 0;
          --ҩƷ���Ͳ�Ϊ�գ������ҩƷ����ƥ������

          if (v_ҩƷ���� is null or n_����ҩƷ����=0) and n_�Զ�����=0 then
            If d_Oldִ��ʱ�� <> Trunc(v_ִ��ʱ��) Or d_Oldִ��ʱ�� Is Null Then
              b_Change      := True;
              d_Oldִ��ʱ�� := v_ִ��ʱ��;

              Select /*+ rule*/
               Count(a.Ҫ��ʱ��)
              Into n_Cur
              From ҽ��ִ��ʱ�� A
              Where a.ҽ��id In (Select ID
                               From ����ҽ����¼
                               Where ����id = (Select ����id From ����ҽ����¼ Where ID = v_ҽ����¼.���id And Rownum < 2)) And
                    a.Ҫ��ʱ�� Between Trunc(v_ִ��ʱ��) And Trunc(v_ִ��ʱ��+1) - 1 / 24 / 60 / 60;

              Select Count(a.Ҫ��ʱ��)
              Into n_Sum
              From ҽ��ִ��ʱ�� A
              Where a.ҽ��id In (Select ID
                               From ����ҽ����¼
                               Where ����id = (Select ����id From ����ҽ����¼ Where ID = v_ҽ����¼.���id And Rownum < 2)) And
                    a.Ҫ��ʱ�� Between Trunc(v_ִ��ʱ�� - 1) And Trunc(v_ִ��ʱ��) - 1 / 24 / 60 / 60;

              Select Count(Distinct a.��ҩ����)
              Into n_��ҩ��
              From ��Һ��ҩ��¼ A
              Where a.ҽ��id In (Select ID From ����ҽ����¼ Where ����id = v_ҽ����¼.����id And ���id Is Null) And
                    a.ִ��ʱ�� Between Trunc(v_ִ��ʱ�� - 1) And Trunc(v_ִ��ʱ��) - 1 / 24 / 60 / 60;

              If n_Cur <> n_Sum Or  n_��ҩ�� > 1 Then
                Update ��Һ��ҩ��¼
                Set �Ƿ�������� = 1
                Where ҽ��id In (Select ID
                               From ����ҽ����¼
                               Where ����id = (Select ����id From ����ҽ����¼ Where ID = v_ҽ����¼.���id And Rownum < 2)) And
                      ִ��ʱ�� Between Trunc(v_ִ��ʱ��) And Trunc(v_ִ��ʱ�� + 1) - 1 / 24 / 60 / 60 And ����״̬ < 2;
                b_Change := False;
              End If;
            End If;

            If b_Change = True Then
              b_Change := True;
              n_����id := v_ҽ����¼.����id;
              Select Count(ID)
              Into n_Sum
              From ��Һ��ҩ��¼
              Where ҽ��id = v_ҽ����¼.���id And ִ��ʱ�� Between Trunc(v_ִ��ʱ�� - 1) And Trunc(v_ִ��ʱ��) - 1 / 24 / 60 / 60;
              If n_Sum = 0 Then
                Update ��Һ��ҩ��¼
                Set �Ƿ�������� = 1
                Where ҽ��id In (Select ID
                               From ����ҽ����¼
                               Where ����id = (Select ����id From ����ҽ����¼ Where ID = v_ҽ����¼.���id And Rownum < 2)) And
                      ִ��ʱ�� Between Trunc(v_ִ��ʱ��) And Trunc(v_ִ��ʱ�� + 1) - 1 / 24 / 60 / 60 And ����״̬ < 2;
                b_Change := False;
              End If;

              If b_Change = True Then
                For v_����ҽ����¼ In c_����ҽ����¼ Loop
                  --�����Һ���Ƿ���������״̬
                  Select Count(ID)
                  Into n_Sum
                  From ��Һ��ҩ��¼
                  Where ҽ��id = v_����ҽ����¼.���id And ִ��ʱ�� Between Trunc(v_ִ��ʱ�� - 1) And Trunc(v_ִ��ʱ��) - 1 / 24 / 60 / 60 And
                        ���ʱ�� Is Not Null;
                  If n_Sum <> 0 Then
                    Update ��Һ��ҩ��¼
                    Set �Ƿ�������� = 1
                    Where ҽ��id In
                          (Select ID
                           From ����ҽ����¼
                           Where ����id = (Select ����id From ����ҽ����¼ Where ID = v_����ҽ����¼.���id And Rownum < 2)) And
                          ִ��ʱ�� Between Trunc(v_ִ��ʱ��) And Trunc(v_ִ��ʱ�� + 1) - 1 / 24 / 60 / 60 And ����״̬ < 2;
                    b_Change := False;
                    Exit;
                  End If;

                  Select Count(ҽ��id)
                  Into n_Cur
                  From ҽ��ִ��ʱ��
                  Where ҽ��id = v_����ҽ����¼.���id And Ҫ��ʱ�� Between Trunc(v_ִ��ʱ��) And Trunc(v_ִ��ʱ�� + 1) - 1 / 24 / 60 / 60;
                  Select Count(ҽ��id)
                  Into n_Sum
                  From ҽ��ִ��ʱ��
                  Where ҽ��id = v_����ҽ����¼.���id And Ҫ��ʱ�� Between Trunc(v_ִ��ʱ�� - 1) And Trunc(v_ִ��ʱ��) - 1 / 24 / 60 / 60;
                  If n_Sum <> n_Cur Then
                    Update ��Һ��ҩ��¼
                    Set �Ƿ�������� = 1
                    Where ҽ��id In
                          (Select ID
                           From ����ҽ����¼
                           Where ����id = (Select ����id From ����ҽ����¼ Where ID = v_����ҽ����¼.���id And Rownum < 2)) And
                          ִ��ʱ�� Between Trunc(v_ִ��ʱ��) And Trunc(v_ִ��ʱ�� + 1) - 1 / 24 / 60 / 60 And ����״̬ < 2;
                    b_Change := False;
                    Exit;
                  End If;
                End Loop;
              End If;
            End If;

            If (v_�����ϴ����� = 1 Or b_Change = True) and n_�Զ�����=0 Then
              --ȡ�ϴε�����
              Begin
                Select Distinct ��ҩ����
                Into v_����
                From ��Һ��ҩ��¼ A
                Where ҽ��id = v_ҽ����¼.���id And
                      ���ͺ� = (Select Distinct Max(���ͺ�)
                             From ��Һ��ҩ��¼
                             Where ҽ��id = v_ҽ����¼.���id And ���ͺ� <> v_ҽ����¼.���ͺ�) And
                      To_Char(a.ִ��ʱ��, 'hh24:mi:ss') = To_Char(v_ִ��ʱ��, 'hh24:mi:ss');
              Exception
                When Others Then
                  v_���� := 0;
              End;
            End If;

            If v_���� = 0 Then
              v_���� := Zl_Getpivaworkbatch(v_ִ��ʱ��);

              --ͬ����ͬ��������Һ�����ƣ���������䵽�¸�����
              If v_��Һ���� > 0 And Not v_����Һ���� Is Null And v_���� < v_Maxbatch Then
                Begin
                  Select /*+rule */
                   Sum(����) As ����
                  Into v_Curdose
                  From (Select Distinct c.Id, c.����
                         From ����ҽ����¼ A, ����ҽ������ B, ҩƷ�շ���¼ C, סԺ���ü�¼ D, ҩƷ��� E, ҩƷ���� F, Table(f_Str2list(v_����Һ����)) G
                         Where a.Id = b.ҽ��id And c.����id = d.Id And a.Id = d.ҽ����� And b.No = c.No And c.ҩƷid = e.ҩƷid And
                               e.ҩ��id = f.ҩ��id And b.ִ�в���id + 0 = ����id_In And c.���� = 9 And c.������� Is Null And
                               f.ҩƷ���� = g.Column_Value And a.���id = v_���id And b.���ͺ� = v_���ͺ�);
                Exception
                  When Others Then
                    v_Curdose := 0;
                End;

                Begin
                  Select /*+rule */
                   Sum(����) As ����
                  Into v_Sumdose
                  From (Select Distinct a.Id, a.����
                         From ҩƷ�շ���¼ A, ����ҽ����¼ B, ��Һ��ҩ��¼ C, ��Һ��ҩ���� D, ҩƷ��� E, ҩƷ���� F, Table(f_Str2list(v_����Һ����)) G
                         Where c.Id = d.��¼id And a.Id = d.�շ�id And c.ҽ��id = b.Id And a.ҩƷid + 0 = e.ҩƷid And
                               e.ҩ��id = f.ҩ��id And b.����id + 0 = v_ҽ����¼.����id And f.ҩƷ���� = g.Column_Value And
                               c.ִ��ʱ�� Between Trunc(v_ִ��ʱ��) And Trunc(v_ִ��ʱ�� + 1) - 1 / 24 / 60 / 60 And c.��ҩ���� = v_����);
                Exception
                  When Others Then
                    v_Sumdose := 0;
                End;

                If v_Sumdose > 0 And v_Sumdose + v_Curdose > v_��Һ���� Then
                  v_���� := v_���� + 1;
                End If;
              End If;
            End If;

          elsif v_ҩƷ���� is not null and n_����ҩƷ����=1 then
            --ҩƷ���Ͳ�Ϊ�գ�ֱ�Ӹ���ҩƷ����ƥ������
            v_���� := Zl_Getpivaworkbatch(sysdate,v_ҩƷ����);
          else
            v_���� := Zl_Getpivaworkbatch(v_ִ��ʱ��);
          end if;  

          Select Count(ҽ��id)
          Into n_���ʹ���
          From ҽ��ִ��ʱ��
          Where ҽ��id = v_ҽ����¼.���id And Ҫ��ʱ�� <= v_ִ��ʱ��
          Order By Ҫ��ʱ��;

          If n_���ʹ��� > 99 Then
            n_���ʹ��� := Mod(n_���ʹ���, 99);
          End If;

          If Length(v_ҽ����¼.���id) > 9 Then
            If n_���ʹ��� < 10 Then
              Select '91' || Substr(To_Char(v_ҽ����¼.���id), Length(v_ҽ����¼.���id) - 8) || To_Char(v_ҽ����¼.���id) || '0' ||
                      To_Char(n_���ʹ���)
              Into v_Maxno
              From Dual;
            Else
              Select '91' || Substr(To_Char(v_ҽ����¼.���id), Length(v_ҽ����¼.���id) - 8) || To_Char(v_ҽ����¼.���id) ||
                      To_Char(n_���ʹ���)
              Into v_Maxno
              From Dual;
            End If;
          Else
            If n_���ʹ��� < 10 Then
              Select '91' || Substr('000000000', Length(v_ҽ����¼.���id) + 1) || To_Char(v_ҽ����¼.���id) || '0' ||
                      To_Char(n_���ʹ���)
              Into v_Maxno
              From Dual;
            Else
              Select '91' || Substr('000000000', Length(v_ҽ����¼.���id) + 1) || To_Char(v_ҽ����¼.���id) || To_Char(n_���ʹ���)
              Into v_Maxno
              From Dual;
            End If;
          End If;
          n_�������� := 0;
          If b_Change = False Then
            n_�������� := 1;
          End If;

          If v_���� <> 0 Then
            Select Nvl(Max(���), 0) Into n_��� From ��ҩ�������� Where ���� = v_����;
          End If;

          If (Trunc(v_ִ��ʱ��) <= v_Currdate Or n_��� <> 0) and (v_ҩƷ���� is null or n_����ҩƷ����=0) Then
            n_�Ƿ���     := 1;
            d_�ֹ����ʱ�� := Sysdate;
          Else
            n_�Ƿ���     := 0;
            d_�ֹ����ʱ�� := Null;
          End If;

          --�����TPN��������������ζ�����Ϊ����
          If v_ҽ����¼.�Ƿ�tpn = 2 Then
            n_�Ƿ���     := 0;
            d_�ֹ����ʱ�� := Null;
          End If;

          if v_����=0 then
             n_�Ƿ���:=1;
          end if;
          --������ҩ��¼
          Insert Into ��Һ��ҩ��¼
            (ID, ����id, ���, ����, �Ա�, ����, סԺ��, ����, ���˲���id, ���˿���id, ִ��ʱ��, ҽ��id, ���ͺ�, ��ҩ����, ƿǩ��, �Ƿ��������, �Ƿ���, ���ʱ��, ����״̬,
             ������Ա, ����ʱ��)
          Values
            (v_��ҩid, ����id_In, v_���, v_ҽ����¼.����, v_ҽ����¼.�Ա�, v_ҽ����¼.����, v_ҽ����¼.סԺ��, v_ҽ����¼.����, v_ҽ����¼.���˲���id,
             v_ҽ����¼.���˿���id, v_ִ��ʱ��, v_ҽ����¼.���id, v_ҽ����¼.���ͺ�, v_����, v_Maxno, n_��������, n_�Ƿ���,
             d_�ֹ����ʱ��, 1, �˲���_In, �˲�ʱ��_In);

          Insert Into ��Һ��ҩ״̬ (��ҩid, ��������, ������Ա, ����ʱ��) Values (v_��ҩid, 1, �˲���_In, �˲�ʱ��_In);

          --������ҩ��¼��Ӧ��ҩƷ��¼
          For v_�շ���¼ In c_�շ���¼ Loop
            If v_�շ���¼.�Ƿ������� = 1 Then
              v_Nodosage := 1;
            End If;

            n_Count := n_Count + 1;

            Select ҩƷ�շ���¼_Id.Nextval Into n_Lngid From Dual;

            Insert Into ҩƷ�շ���¼
              (ID, ��¼״̬, ����, NO, ���, �ⷿid, ��ҩ��λid, ������id, �Է�����id, ���ϵ��, ҩƷid, ����, ����, ����, ��������, Ч��, ����, ��д����, ʵ������,
               �ɱ���, �ɱ����, ����, ���ۼ�, ���۽��, ���, ժҪ, ������, ��������, ��ҩ��, ��ҩ����, �����, �������, �۸�id, ����id, ����, Ƶ��, �÷�, ���, �������, ���Ч��,
               ��Ʒ�ϸ�֤, ��ҩ��ʽ, ��ҩ����, ������, ��׼�ĺ�, ���ܷ�ҩ��, ע��֤��, �ⷿ��λ, ��Ʒ����, �ڲ�����, �˲���, �˲�����, ǩ��ȷ����, ǩ��ʱ��)
              Select n_Lngid, ��¼״̬, ����, NO, n_Count + 1000, �ⷿid, ��ҩ��λid, ������id, �Է�����id, ���ϵ��, ҩƷid, ����, ����, ����, ��������,
                     Ч��, ����, ��д���� / v_Count, ʵ������ / v_Count, �ɱ���, �ɱ���� / v_Count, ����, ���ۼ�, ���۽�� / v_Count, ��� / v_Count,
                     '����', ������, ��������, ��ҩ��, ��ҩ����, �����, �������, �۸�id, ����id, ����, Ƶ��, �÷�, ���, �������, ���Ч��, ��Ʒ�ϸ�֤, ��ҩ��ʽ, ��ҩ����,
                     ������, ��׼�ĺ�, ���ܷ�ҩ��, ע��֤��, �ⷿ��λ, ��Ʒ����, �ڲ�����, �˲���, �˲�����, ǩ��ȷ����, ǩ��ʱ��



              From ҩƷ�շ���¼
              Where ID = v_�շ���¼.�շ�id;

            Insert Into ��Һ��ҩ���� (��¼id, �շ�id, ����) Values (v_��ҩid, n_Lngid, v_�շ���¼.���� / v_Count);

            n_���ȼ�:=Zl_GetFirst(v_��ҩid,v_ҽ����¼.���˿���id);
            update ��Һ��ҩ��¼ set ���ȼ�=n_���ȼ� where id=v_��ҩid;
          End Loop;

        End Loop;

        For v_�շ���¼ In c_�շ���¼ Loop
          n_���� := v_�շ���¼.����;

          v_No := v_�շ���¼.No;
          Delete From ҩƷ�շ���¼ Where ID = v_�շ���¼.�շ�id;
        End Loop;

        --����ҩƷ���߲������õ�ҩƷĬ��Ϊ0����
        select count(�շ�id) into n_Row from ��Һ��ҩ���� where ��¼id=v_��ҩid;
        if (v_Nodosage = 1 or n_row=1) and n_���ҩƷ����=1 then
          Update ��Һ��ҩ��¼ Set ��ҩ���� = 0,�Ƿ��� = 1 Where ҽ��id = v_ҽ����¼.���id And ���ͺ� = v_ҽ����¼.���ͺ� and ����״̬<2;
        end if;
        --������ڡ��������á����Ե�ҩƷ��Ҳ����Ϊ���
        If v_Nodosage = 1 Then
          Update ��Һ��ҩ��¼ Set �Ƿ��� = 1 Where ҽ��id = v_ҽ����¼.���id And ���ͺ� = v_ҽ����¼.���ͺ� and ����״̬<2;
        End If;
      End If;
    End If;
  End Loop;

  For v_�շ���¼ In (Select ID From ҩƷ�շ���¼ Where ��� < 1000 And ���� = n_���� And NO = v_No) Loop
    n_Count := n_Count + 1;
    Update ҩƷ�շ���¼ Set ��� = n_Count + 1000, ժҪ = '����' Where ID = v_�շ���¼.Id;
  End Loop;

  Update ҩƷ�շ���¼
  Set ��� = ��� - 1000, ժҪ = 'ҽ������'
  Where ժҪ = '����' And ��� > 1000 And ���� = n_���� And NO = v_No;

  if n_�Զ�����=1 then
    Zl_��Һ��ҩ��¼_�Զ�����(n_����id,n_����id,����id_In,v_ִ��ʱ��);
  end if;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]����' || v_��ǰ���� || '����Һ���������б���������Һ��������ʧ�ܣ�[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Һ��ҩ��¼_�˲�;
/


--90447:������,2016-02-22,�����ȷ��
Create Or Replace Function Zl1_Fun_Getreturnvisit
(
  ����id_In ������Ϣ.����id%Type,
  ����id_In ���ű�.Id%Type
) Return Number As
  --------------------------------------------------------------------------------------------------
  --����:�Һ�ʱ�жϲ����Ƿ���
  --Ĭ�Ϲ���:��ǰ�Һſ��ҵ��ٴ����ʻ��ٴ�����ȡ��λ�ı�������30���ڴ��ڹҺż�¼��,Ϊ����
  --���:
  --  ����ID_In   : �ҺŲ���ID
  --  ����ID_In   : �Һſ���ID
  --����:�Ƿ���,1-����,0-����
  --------------------------------------------------------------------------------------------------
  v_�Һ����� �ٴ�����.��������%Type;
  n_Exists   Number(3);
  v_����ids  Varchar2(4000);
  v_Err_Msg  Varchar2(200);
  Err_Item Exception;
Begin
  Begin
    Select �������� Into v_�Һ����� From �ٴ����� Where ����id = ����id_In And Rownum < 2;
  Exception
    When Others Then
      Return 0;
  End;
  v_�Һ����� := Substr(v_�Һ�����, 1, 2) || '%';
  For r_���� In (Select Distinct ����id From �ٴ����� Where �������� Like v_�Һ�����) Loop
    v_����ids := v_����ids || ',' || r_����.����id;
  End Loop;
  If v_����ids Is Not Null Then
    v_����ids := Substr(v_����ids, 2);
  End If;
  n_Exists := 0;
  Select Max(1)
  Into n_Exists
  From ���˹Һż�¼
  Where �Ǽ�ʱ�� > Sysdate - 30 And ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 1 And
        ִ�в���id In (Select Column_Value From Table(f_Str2list(v_����ids)));
  Return Nvl(n_Exists, 0);
Exception
  When Err_Item Then
    v_Err_Msg := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End;
/


--91225:������,2016-02-16,�������Լ�¼�����ֶ� ҽ��ID
CREATE OR REPLACE Procedure Zl_�������Լ���¼_Insert
(
  Id_In         In �������Լ�¼.Id%Type,
  ����id_In     In �������Լ�¼.����id%Type,
  ��ҳid_In     In �������Լ�¼.��ҳid%Type,
  �Һŵ�_In     In �������Լ�¼.�Һŵ�%Type,
  ҽ��ID_In     In �������Լ�¼.ҽ��ID%Type,
  �ͼ�ʱ��_In   In �������Լ�¼.�ͼ�ʱ��%Type,
  �ͼ����id_In In �������Լ�¼.�ͼ����id%Type,
  �ͼ�ҽ��_In   In �������Լ�¼.�ͼ�ҽ��%Type,
  �걾����_In   In �������Լ�¼.�걾����%Type,
  �������_In   In �������Լ�¼.�������%Type,
  ��Ⱦ��_In     In �������Լ�¼.��Ⱦ������%Type,
  ���ʱ��_In   In �������Լ�¼.���ʱ��%Type,
  �Ǽ�ʱ��_In   In �������Լ�¼.�Ǽ�ʱ��%Type,
  �Ǽ���_In     In �������Լ�¼.�Ǽ���%Type,
  �Ǽǿ���id_In In �������Լ�¼.�Ǽǿ���id%Type,
  ��¼״̬_In   In �������Լ�¼.��¼״̬%Type
) Is
Begin
  Insert Into �������Լ�¼
    (ID, ����id, ��ҳid, �Һŵ�, ҽ��ID,�ͼ�ʱ��, �ͼ����id, �ͼ�ҽ��, �걾����, �������, ��Ⱦ������, ���ʱ��, �Ǽ�ʱ��, �Ǽ���, �Ǽǿ���id, ��¼״̬)
  Values
    (Id_In, ����id_In, ��ҳid_In, �Һŵ�_In, ҽ��ID_In,�ͼ�ʱ��_In, �ͼ����id_In, �ͼ�ҽ��_In, �걾����_In, �������_In, ��Ⱦ��_In, ���ʱ��_In, �Ǽ�ʱ��_In, �Ǽ���_In,
     �Ǽǿ���id_In, ��¼״̬_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�������Լ���¼_Insert;
/

--93336:������,2016-02-16,��Ѫִ�еǼ�
Create Or Replace Function Zl_Fun_Get��Ѫִ�еǼ�(ҽ��id_In In ����ҽ����¼.Id%Type) Return Varchar2 Is
  n_Count ����ҽ��ִ��.��������%Type;
  n_Boold Number; --�Ƿ���Ѫ��ϵͳ
  v_Tmp   Varchar2(4000);
  n_Bags  Number; --Ѫ����Ѫ����
  --���ܣ��� ��Ѫ;�� ҽ��ִ�еǼ�
  --������ҽ��id_In ��ҽ��ID
  --���أ��̶���ʽ�ַ����� �Ǽ�һ������д���������Ƿ�����Ѫ�⣬Ѫ������������0.33333<SPLIT>1<SPLIT>3
  --˵������̨������5λС���������뱣�棬������ʾ���ж�ʱҪ����Ѫ��������������ȡ�����бȽ�
Begin
  Select Count(1) Into n_Boold From zlSystems Where ��� = 2200;
  If n_Boold = 1 Then
    n_Boold := Nvl(zl_GetSysParameter(236), 0);
  End If;
  If n_Boold = 1 Then
    v_Tmp := 'Select Zl_Get_��Ѫִ�д���(:1) as ���� From Dual';
    Execute Immediate v_Tmp
      Into n_Bags
      Using ҽ��id_In;
  End If;
  If n_Bags Is Null Then
    n_Count := 1;
  Else
    n_Count := Round(1 / n_Bags, 5);
  End If;
  v_Tmp := To_Char(n_Count, '0.99999') || '<SPLIT>' || Nvl(n_Boold, 0) || '<SPLIT>' || Nvl(n_Bags, 0);
  Return v_Tmp;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Fun_Get��Ѫִ�еǼ�;
/  
--91646:�ŵ���,2016-02-22,��ҩ��ȡ���Ϸ�
CREATE OR REPLACE Function zl_fun_PIVACustom
( 
  ��ҩid_In     In ��Һ��ҩ��¼.ID%type
) Return Varchar2 Is 
  ---------------------------------- 
  --���ܣ�������ҩʱ������ҩƷ����ҩ���ͣ�ʹ�õ�������ҩ;��ȷ���շѵĲ���id��
  --���أ�
  --      �շ���Ŀid1m��ȡ���Σ��շ���Ŀid2,�շ�����   ���磺12,1;13,2  ��Ҫ�շ���ĿidΪ12����ȡһ�η��ã��շ���ĿidΪ13����ȡ���η���
  ----------------------------------- 
  v_Temp   varchar2(500);
  Err_Custom Exception; 
  v_Error Varchar2(255); 
  v_IsSpecial varchar2(20);
Begin
  v_Temp:='';
  for r_item in (select distinct A.ҩƷid,A.��ҩ����,F.���� ��ҩ;��,D.����,G.����ϵ�� from ��ҺҩƷ���� A,��Һ��ҩ��¼ B,��Һ��ҩ���� C,ҩƷ�շ���¼ D,����ҽ����¼ E,������ĿĿ¼ F,ҩƷ��� G where B.ҽ��id=E.id and E.������Ŀid=F.id and B.Id=C.��¼ID and C.�շ�ID=D.ID and D.ҩƷID=A.ҩƷID and D.ҩƷID=G.ҩƷID and
    B.Id= ��ҩid_In) loop
    --�ж��Ƿ�Ϊ�ȵ���
    if r_item.��ҩ����='�ȵ���' then
      --1mlע����,�շ�����
      v_IsSpecial:='';
    end if;
      
    if r_item.��ҩ����='����ҩ' then
      --���ݸ�ҩ;���ж�
       if  r_item.��ҩ;��='��ע' then
         --60mlע����,�շ�����
         v_Temp:='';
       else 
         --20mlע����,�շ�����
          v_Temp:='';
       end if;
       exit; 
    end if;
  
    if r_item.��ҩ����='Ӫ��Һ' then
       --�����շ���Ŀ�ֱ�����
       v_Temp:='';   
       
       exit;   
    end if;
  
    --�ж�����ҩƷ���ҳ���Ӧ���շ�id
    if  r_item.ҩƷid='1' then
      --60mlע����,�շ�����
      v_Temp:='';  
    elsif v_Temp is null then
      --20mlע����,�շ�����
      v_Temp:='';
    end if;
  
  end loop;
  Return v_Temp || v_IsSpecial; 
Exception 
  When Err_Custom Then 
    Return Null; 
  When Others Then 
    zl_ErrorCenter(SQLCode, SQLErrM); 
End zl_fun_PIVACustom;
/       

--91646:�ŵ���,2016-02-03,��ҩ��ȡ���Ϸ�
CREATE OR REPLACE Procedure Zl_��Һ��ҩ��¼_��ҩ
(
  ��ҩid_In   In Varchar2, --ID��:ID1,ID2....
  ������Ա_In In ��Һ��ҩ��¼.������Ա%Type := Null,
  ����ʱ��_In In ��Һ��ҩ��¼.����ʱ��%Type := Null
) Is
  v_Tansid   Varchar2(20);
  v_Tmp      Varchar2(4000);
  v_No       Varchar2(20);
  v_Usercode Varchar2(100);
  n_����״̬ ��Һ��ҩ��¼.����״̬%Type;
  v_Error    Varchar2(255);
  n_People      number(1);
  n_row      number(2);
  d_ִ��ʱ�� date;
  v_��ҩ���� varchar2(50);
  n_��Ŀid   number(18);
  v_�շ���Ŀid varchar2(200);
  v_info    varchar2(200);
  v_id varchar2(20);
  n_���� number(2);
  n_count number(18);
  n_Out number(10);
  n_OutNum number(10);
  n_���״̬ number(1);
  Err_Custom Exception;

  Cursor c_Bill Is
    Select a.����id, a.��ҳid, a.��ʶ��, a.����, a.�Ա�, a.����, a.����, a.�ѱ�, a.���˲���id, a.���˿���id, a.Ӥ����, e.ҩƷid, b.�ⷿid,f.��ҩ����
    From סԺ���ü�¼ A, ҩƷ�շ���¼ B, ��Һ��ҩ��¼ C, ��Һ��ҩ���� D, ҩƷ��� E, ��ҺҩƷ���� F
    Where a.Id = b.����id And b.Id = d.�շ�id And d.��¼id = c.Id And b.ҩƷid = e.ҩƷid And b.ҩƷid = f.ҩƷid And Nvl(c.�Ƿ���, 0) <> 1 And c.Id = v_Tansid;

Begin
  v_Usercode := Zl_Identity;
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ';') + 1);
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ',') + 1);
  v_Usercode := Substr(v_Usercode, 1, Instr(v_Usercode, ',') - 1);
  n_People:=Nvl(zl_GetSysParameter('���÷Ѱ�������ȡ', 1345), 0);
  n_Out:=Nvl(zl_GetSysParameter('��Ժ���˲������÷�', 1345), 0);

  If ��ҩid_In Is Null Then
    v_Tmp := Null;
  Else
    v_Tmp := ��ҩid_In || ',';
  End If;

  While v_Tmp Is Not Null Loop
    --�ֽⵥ��ID��
    v_Tansid := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp    := Replace(',' || v_Tmp, ',' || v_Tansid || ',');

    --��鵱ǰ��Һ����״̬�Ƿ�Ϊ����ҩ״̬
    Begin
      Select ����״̬,ִ��ʱ��,nvl(�Ƿ���,0)  Into n_����״̬,d_ִ��ʱ��,n_���״̬ From ��Һ��ҩ��¼ Where ID = v_Tansid;

      if n_����״̬>3 then
        v_Error := '�������ѱ����������ܽ��з�ҩ��';
        Raise Err_Custom;
      end if;
    Exception
      When Others Then
        v_Error := '�������ѱ�������';
        Raise Err_Custom;
    End;

    Update ��Һ��ҩ��¼ Set ����״̬ = 4, ������Ա = ������Ա_In, ����ʱ�� = ����ʱ��_In Where ID = v_Tansid;
    Insert Into ��Һ��ҩ״̬ (��ҩid, ��������, ������Ա, ����ʱ��) Values (v_Tansid, 4, ������Ա_In, ����ʱ��_In);

    if n_���״̬=0 then
      n_count:=0;
      Select Nextno(14) Into v_No From Dual;
      For r_Bill In c_Bill Loop
        Select count(����id) into n_OutNum From ������ҳ where ��ҳID=r_Bill.��ҳid And ����ID=r_Bill.����id  And (Nvl(״̬,0)=3 Or ��Ժ���� Is Not NULL);
        if n_count=0 and (n_OutNum=0 or n_out=0) then
          --��ȡ���Ϸ�
          --v_�շ���Ŀid:='6970,2;6971,1;';
          select zl_fun_PIVACustom(v_Tansid) into  v_�շ���Ŀid from dual;
          While v_�շ���Ŀid Is Not Null Loop
             v_info:= Substr(v_�շ���Ŀid, 1, Instr(v_�շ���Ŀid, ';') - 1);
             v_�շ���Ŀid := Replace(';' || v_�շ���Ŀid, ';' || v_info || ';');
             

             v_id:= Substr(v_info, 1, Instr(v_info, ',') - 1);
             v_info := Replace(',' || v_info, ',' || v_id || ',');

            For r_Item In (Select a.Id �շ�ϸĿid, a.��� �շ����, a.���㵥λ, a.�Ӱ�Ӽ� �Ӱ��־, d.Id ������Ŀid, d.�վݷ�Ŀ, b.�ּ�
                           From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ D
                           Where a.Id = b.�շ�ϸĿid And b.������Ŀid = d.Id And a.id=v_id and
                                 b.ִ������ <= Sysdate And
                                 (b.��ֹ���� >= Sysdate Or b.��ֹ���� Is Null)) Loop
              if n_count=0 then
                Insert Into ��Һ��ҩ���� (��ҩid, NO,����id) Values (v_Tansid, v_No, r_Bill.����id);
              end if;

              n_count:=n_count+1;
              Zl_סԺ���ʼ�¼_Insert(v_No, n_count, r_Bill.����id, r_Bill.��ҳid, r_Bill.��ʶ��, r_Bill.����, r_Bill.�Ա�, r_Bill.����, r_Bill.����,
                               r_Bill.�ѱ�, r_Bill.���˲���id, r_Bill.���˿���id, r_Item.�Ӱ��־, r_Bill.Ӥ����, r_Bill.�ⷿid, ������Ա_In, Null,
                               r_Item.�շ�ϸĿid, r_Item.�շ����, r_Item.���㵥λ, Null, Null, Null, 1,v_info, Null, r_Bill.�ⷿid, Null,
                               r_Item.������Ŀid, r_Item.�վݷ�Ŀ, r_Item.�ּ�, r_Item.�ּ�*v_info, r_Item.�ּ�*v_info, Null, Sysdate, Sysdate, Null, Null,
                               v_Usercode, ������Ա_In);
            End Loop;
          end loop;
        end if;


        select count(��Ŀid) into n_��Ŀid from �����շѷ��� where ��ҩ����=substr(r_Bill.��ҩ����,INSTR(r_Bill.��ҩ����,'-',1,1)+1);
        if n_��Ŀid<>0 then
          n_row:=0;
          select nvl(��Ŀid,0) into n_��Ŀid from �����շѷ��� where ��ҩ����=substr(r_Bill.��ҩ����,INSTR(r_Bill.��ҩ����,'-',1,1)+1);
          if n_People=1 then
            select count(��ҩid) into n_row from ��Һ��ҩ���� A,סԺ���ü�¼ B,��Һ��ҩ��¼ C where A.No=b.no and A.��ҩID=C.id and b.����id=r_Bill.����id And B.��¼״̬=1 and B.�շ�ϸĿid=n_��Ŀid and d_ִ��ʱ�� Between Trunc(c.ִ��ʱ��) And Trunc(c.ִ��ʱ��+1) - 1 / 24 / 60 / 60;
          end if;
        else
          n_row:=1;
        end if;

        if n_row=0 and (n_OutNum=0 or n_out=0) then
          For r_Item In (Select a.Id �շ�ϸĿid, a.��� �շ����, a.���㵥λ, a.�Ӱ�Ӽ� �Ӱ��־, d.Id ������Ŀid, d.�վݷ�Ŀ, b.�ּ�
                         From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ D
                         Where a.Id = b.�շ�ϸĿid And b.������Ŀid = d.Id And a.id=n_��Ŀid and
                               b.ִ������ <= Sysdate And
                               (b.��ֹ���� >= Sysdate Or b.��ֹ���� Is Null)) Loop
            if n_count=0 then
              Insert Into ��Һ��ҩ���� (��ҩid, NO,����id) Values (v_Tansid, v_No, r_Bill.����id);
            end if;

            n_count:=n_count+1;
            Zl_סԺ���ʼ�¼_Insert(v_No, n_count, r_Bill.����id, r_Bill.��ҳid, r_Bill.��ʶ��, r_Bill.����, r_Bill.�Ա�, r_Bill.����, r_Bill.����,
                             r_Bill.�ѱ�, r_Bill.���˲���id, r_Bill.���˿���id, r_Item.�Ӱ��־, r_Bill.Ӥ����, r_Bill.�ⷿid, ������Ա_In, Null,
                             r_Item.�շ�ϸĿid, r_Item.�շ����, r_Item.���㵥λ, Null, Null, Null, 1, 1, Null, r_Bill.�ⷿid, Null,
                             r_Item.������Ŀid, r_Item.�վݷ�Ŀ, r_Item.�ּ�, r_Item.�ּ�, r_Item.�ּ�, Null, Sysdate, Sysdate, Null, Null,
                             v_Usercode, ������Ա_In);
          End Loop;
        end if;

        if n_People<>1 and n_row=0 then
          Exit;
        end if;
      End Loop;
    end if;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Һ��ҩ��¼_��ҩ;
/

--92898:������,2016-02-03,����ʱ������Ϊ�գ����������ʱ��Ļ����Ե�ǰʱ���滻
CREATE OR REPLACE Procedure Zl_���˹�����¼_Insert
( 
  ����id_In     ���˹�����¼.����id%Type, 
  ��ҳid_In     ���˹�����¼.��ҳid%Type, 
  ��Դ_In       ���˹�����¼.��¼��Դ%Type, 
  ҩ��id_In     ���˹�����¼.ҩ��id%Type, 
  ҩ����_In     ���˹�����¼.ҩ����%Type, 
  ���_In       ���˹�����¼.���%Type := 1, 
  ����ʱ��_In   ���˹�����¼.����ʱ��%Type := Null, 
  ��¼ʱ��_In   ���˹�����¼.��¼ʱ��%Type := Null, 
  ������Ӧ_In   ���˹�����¼.������Ӧ%Type := Null, 
  ����Դ����_In ���˹�����¼.����Դ����%Type := Null 
) Is 
  V_Date     Date; 
  V_Temp     Varchar2(255); 
  V_��Ա���� ���˹�����¼.��¼��%Type; 
  N_Count    Number; 
Begin 
  V_Temp     := Zl_Identity; 
  V_Temp     := Substr(V_Temp, Instr(V_Temp, ';') + 1); 
  V_Temp     := Substr(V_Temp, Instr(V_Temp, ',') + 1); 
  V_��Ա���� := Substr(V_Temp, Instr(V_Temp, ',') + 1); 
 
  If ��¼ʱ��_In Is Not Null Then 
    V_Date := ��¼ʱ��_In; 
  Else 
    Select Sysdate Into V_Date From Dual; 
  End If; 
 
  Insert Into ���˹�����¼ 
    (ID, ����id, ��ҳid, ��¼��Դ, ҩ��id, ҩ����, ���, ����ʱ��, ��¼ʱ��, ��¼��, ������Ӧ,����Դ����) 
  Values 
    (���˹�����¼_Id.Nextval, ����id_In, ��ҳid_In, ��Դ_In, ҩ��id_In, ҩ����_In, ���_In, ����ʱ��_In, V_Date, V_��Ա����, ������Ӧ_In,����Դ����_In); 
 
  Select Count(1) Into N_Count From ���˹���ҩ�� Where ����id = ����id_In And ����ҩ�� = ҩ����_In; 
  If N_Count = 0 Then 
    Insert Into ���˹���ҩ�� 
      (����id, ����ҩ��id, ����ҩ��, ������Ӧ) 
    Values 
      (����id_In, ҩ��id_In, ҩ����_In, ������Ӧ_In); 
  Else 
    Update ���˹���ҩ�� 
    Set ����ҩ��id = ҩ��id_In, ������Ӧ = ������Ӧ_In 
    Where ����id = ����id_In And ����ҩ�� = ҩ����_In; 
  End If; 
Exception 
  When Others Then 
    zl_ErrorCenter(SQLCode, SQLErrM); 
End Zl_���˹�����¼_Insert;
/

--93209:������,2016-02-01,�������볤��,��ҽ��ID���ȴ���8ʱ,��������ԭ����12λ����Ϊ13λ
Create Or Replace Function Nextno
(
  ���_In     In ������Ʊ�.��Ŀ���%Type,
  ����id_In   In ���ű�.Id%Type := Null,
  v_Tag       In Varchar2 := Null,
  �������_In In Integer := 1
) Return Varchar2
--    ���ܣ������ض���������µĺ���,�������£�
  --    һ����Ŀ��ţ�
  --       1   ����ID         ����
  --       2   סԺ��         ����
  --       3   �����         ����
  --       10  ҽ�����ͺ�     ����,˳��������
  --       x   �������ݺ�     �ַ�,���ݱ�Ź���˳��������,���Զ���ȱ
  --    �������λȷ��ԭ��
  --       ��1990Ϊ���������������������0��9/A��Z��˳����Ϊ��ȱ���
  --
  --    ˵����������-10���������Ʊ�,���ڲ�������²�ȱ��(ȡ�˺�,��δʹ��)
  --          For Update�ڲ��������������,����Waitѡ���Ա���������߷��ؿ�
  --          v_Tag ��������δָ���Ĳ���,Ŀǰ ��Ӱ�������(����)
  --    ���أ�������
 Is
  Pragma Autonomous_Transaction;
  v_No     ������Ʊ�.������%Type;
  v_Maxno  ������Ʊ�.������%Type;
  n_Maxno  Number;
  n_Amt    Number;
  n_Mod    ������Ʊ�.��Ź���%Type;
  v_Deptno Varchar2(20);
  v_Year   Varchar2(1);
  v_Tmp    Varchar2(10);

  v_�Թܱ���   Number;
  v_��������   Varchar2(20);
  v_����       Varchar2(10);
  v_ҽ��       Varchar2(18);
  v_Error      Varchar2(255);
  n_Checkmaxno Number;

  Err_Custom Exception;
Begin

  --1.����ID
  If ���_In = 1 Then
    Select Nvl(��Ź���, 0) Into n_Mod From ������Ʊ� Where ��Ŀ��� = ���_In;
  
    --������ȡֵ�����ڲ�Ҫ����ID�����������û����ٲ�������
    If n_Mod = 1 Then
      Select ������Ϣ_Id.Nextval Into v_No From Dual;
    Else
      Select Nvl(������, '0') Into v_Maxno From ������Ʊ� Where ��Ŀ��� = ���_In For Update;
    
      Select Nvl(Max(����id), 0) + 1 Into n_Maxno From ������Ϣ Where ����id >= To_Number(v_Maxno);
    
      Update ������Ʊ� Set ������ = Decode(Sign(n_Maxno - 10), 1, n_Maxno - 10, 1) Where ��Ŀ��� = ���_In;
      v_No := To_Char(n_Maxno);
    End If;
    --2.סԺ��
  Elsif ���_In = 2 Then
    Select Nvl(������, '0'), Nvl(��Ź���, 0)
    Into v_Maxno, n_Mod
    From ������Ʊ�
    Where ��Ŀ��� = ���_In
    For Update;
  
    If n_Mod = 0 Then
      --0.˳����
      Select Nvl(Max(סԺ��), 0) + 1 Into n_Maxno From ������Ϣ Where סԺ�� >= To_Number(v_Maxno);
    
      Update ������Ʊ� Set ������ = Decode(Sign(n_Maxno - 10), 1, n_Maxno - 10, 1) Where ��Ŀ��� = ���_In;
    Elsif n_Mod = 1 Then
      --1.����(YYMM)+˳���(0000)
      v_Tmp := To_Char(Sysdate, 'YYMM');
    
      Select Nvl(Max(סԺ��), To_Number(v_Tmp || '0000')) + 1
      Into n_Maxno
      From ������Ϣ
      Where סԺ�� Like To_Number(v_Tmp) || '%' And סԺ�� >= To_Number(v_Maxno);
      Update ������Ʊ�
      Set ������ = Decode(Sign(n_Maxno - 10 - To_Number(v_Tmp || '0000')), 1, n_Maxno - 10, To_Number(v_Tmp || '0001'))
      Where ��Ŀ��� = ���_In;
    
    Elsif n_Mod = 2 Then
      --2.��(YYYY)+˳���(00000)
      v_Tmp := To_Char(Sysdate, 'YYYY');
    
      Select Nvl(Max(סԺ��), To_Number(v_Tmp || '00000')) + 1
      Into n_Maxno
      From ������Ϣ
      Where סԺ�� Like To_Number(v_Tmp) || '%' And סԺ�� >= To_Number(v_Maxno);
      Update ������Ʊ�
      Set ������ = Decode(Sign(n_Maxno - 10 - To_Number(v_Tmp || '00000')), 1, n_Maxno - 10, To_Number(v_Tmp || '00001'))
      Where ��Ŀ��� = ���_In;
    
    End If;
    v_No := To_Char(n_Maxno);
  
    --3.�����
  Elsif ���_In = 3 Then
    Select Nvl(������, '0'), Nvl(��Ź���, 0)
    Into v_Maxno, n_Mod
    From ������Ʊ�
    Where ��Ŀ��� = ���_In
    For Update;
  
    If n_Mod = 0 Then
      --0.˳����
    
      Select Nvl(Max(�����), 0) + 1 Into n_Maxno From ������Ϣ Where ����� >= To_Number(v_Maxno);
    
      Update ������Ʊ� Set ������ = Decode(Sign(n_Maxno - 10), 1, n_Maxno - 10, 1) Where ��Ŀ��� = ���_In;
    Elsif n_Mod = 1 Then
      --1.���ڱ��YYMMDD
      v_Tmp := To_Char(Sysdate, 'YYMMDD');
    
      Select Nvl(Max(�����), To_Number(v_Tmp || '0000')) + 1
      Into n_Maxno
      From ������Ϣ
      Where ����� Like To_Number(v_Tmp) || '%' And ����� >= To_Number(v_Maxno);
      Update ������Ʊ�
      Set ������ = Decode(Sign(n_Maxno - 10 - To_Number(v_Tmp || '0000')), 1, n_Maxno - 10, To_Number(v_Tmp || '0001'))
      Where ��Ŀ��� = ���_In;
    
    End If;
    v_No := To_Char(n_Maxno);
  
    --10.ҽ�����ͺ�
  Elsif ���_In = 10 Then
    Select Nvl(������, '0') Into v_Maxno From ������Ʊ� Where ��Ŀ��� = ���_In For Update;
    v_Maxno := To_Char(To_Number(v_Maxno) + 1);
    Update ������Ʊ� Set ������ = v_Maxno Where ��Ŀ��� = ���_In;
    v_No := v_Maxno;
  
    --���ܷ�ҩ��
  Elsif ���_In = 20 Then
    --YYYYMMDD+5λ˳���(00000)
    Select To_Char(Sysdate, 'yyyymmdd') Into v_Tmp From Dual;
    Select Nvl(������, '0') Into v_Maxno From ������Ʊ� Where ��Ŀ��� = ���_In For Update;
  
    If v_Maxno = '0' Then
      v_Maxno := v_Tmp || '00001';
    Else
      If Substr(v_Maxno, 1, 8) = v_Tmp Then
        If To_Number(Substr(v_Maxno, 9, 5)) = 99999 Then
          v_Maxno := v_Tmp || '00001';
        Else
          v_Maxno := v_Tmp || Trim(To_Char(To_Number(Substr(v_Maxno, 9, 5)) + 1, '00000'));
        End If;
      Else
        v_Maxno := v_Tmp || '00001';
      End If;
    End If;
    Update ������Ʊ� Set ������ = v_Maxno Where ��Ŀ��� = ���_In;
  
    v_No := v_Maxno;
  
  Elsif ���_In = 123 Then
  
    Select Nvl(Max(����ֵ), 0) Into n_Mod From Ӱ�����̲��� Where ����id = ����id_In And ������ = '�������ɷ�ʽ';
    Select Nvl(Max(����ֵ), 1)
    Into n_Checkmaxno
    From Ӱ�����̲���
    Where ����id = ����id_In And ������ = '��ȡʵ��������';
  
    If n_Mod = 1 Then
      --Ӱ����Ű����ҵ���
      --�Ӻ������ȡ������
      Select Nvl(Max(������), 0) Into n_Maxno From ���Һ���� Where ��Ŀ��� = ���_In And ����id = ����id_In;
      If n_Maxno = 0 Then
        --û�м�¼�����Զ����ӿ��Һ����
        Insert Into ���Һ���� (��Ŀ���, ����id, ���, ������) Values (���_In, ����id_In, 'A', '1');
      End If;
    
      --��ȡʵ��������,�����ʵ�������룬��������Ⱥ�����еĴ�10
      If n_Checkmaxno = 1 Then
      
        Select Nvl(Max(����), 0) + 1 Into n_Amt From Ӱ�����¼ Where ִ�п���id = ����id_In And ���� >= n_Maxno;
      
        If n_Amt > n_Maxno Then
          n_Maxno := n_Amt;
        End If;
      Else
        -- �������ȡʵ�������룬����10
        n_Maxno := n_Maxno + 10 + 1;
      End If;
    
      -- ����������
      If (n_Checkmaxno = 0 And �������_In = n_Maxno) Or n_Checkmaxno = 1 Then
        Update ���Һ����
        Set ������ = Decode(Sign(n_Maxno - 10), 1, n_Maxno - 10, 1)
        Where ��Ŀ��� = ���_In And ����id = ����id_In;
      End If;
    Else
      --Ӱ����Ű�������
      Select Nvl(Max(������), 0) Into n_Maxno From Ӱ������� Where ���� = v_Tag;
      --��ȡʵ�������룬�����ʵ�������룬��������Ⱥ�����еĴ�10
      If n_Checkmaxno = 1 Then
      
        Select Nvl(Max(����), 0) + 1 Into n_Amt From Ӱ�����¼ Where Ӱ����� = v_Tag And ���� >= n_Maxno;
      
        If n_Amt > n_Maxno Then
          n_Maxno := n_Amt;
        End If;
      Else
        -- �������ȡʵ�������룬����10
        n_Maxno := n_Maxno + 10 + 1;
      End If;
    
      -- ����������
      If (n_Checkmaxno = 0 And �������_In = n_Maxno) Or n_Checkmaxno = 1 Then
        Update Ӱ������� Set ������ = Decode(Sign(n_Maxno - 10), 1, n_Maxno - 10, 1) Where ���� = v_Tag;
      End If;
    End If;
    v_No := To_Char(n_Maxno);
  
  Elsif ���_In = 124 Then
    ----------------------------------------------------------------------------------------------------------------------------
    --��콡����
  
    Begin
      Select Nvl(��Ź���, 0) Into n_Mod From ������Ʊ� Where ��Ŀ��� = ���_In;
    Exception
      When Others Then
        v_Error := '������Ʊ��в��������Ϊ' || ���_In || '�ĺ���';
        Raise Err_Custom;
    End;
  
    If n_Mod = 0 Then
      --˳����
      Select Nvl(Zl_To_Number(������), 0) + �������_In
      Into v_Maxno
      From ������Ʊ�
      Where ��Ŀ��� = ���_In
      For Update;
      Update ������Ʊ� Set ������ = v_Maxno Where ��Ŀ��� = ���_In;
    Else
      --ǰ׺+˳���
      Update �����ű� Set ������ = ������ Where ��Ŀ��� = ���_In And ����ǰ׺ = v_Tag;
      If Sql%RowCount = 0 Then
        Insert Into �����ű�
          (��Ŀ���, ����ǰ׺, ����, ������)
          Select ���_In, v_Tag, Sysdate, '' From Dual;
      End If;
    
      Select Nvl(Zl_To_Number(������), 0) + �������_In
      Into n_Maxno
      From �����ű� A
      Where a.��Ŀ��� = ���_In And a.����ǰ׺ = v_Tag;
    
      If Substr(n_Maxno, 1, Length(v_Tag)) <> v_Tag Then
        --��λ��
        Select Nvl(Zl_To_Number(Substr(������, Length(v_Tag) + 1)), 0) + �������_In
        Into n_Maxno
        From �����ű� A
        Where a.��Ŀ��� = ���_In And a.����ǰ׺ = v_Tag;
        Select v_Tag || n_Maxno Into v_Maxno From Dual;
      Else
        If n_Maxno = �������_In Then
          v_Maxno := v_Tag || To_Char(�������_In);
        Else
          v_Maxno := n_Maxno;
        End If;
      End If;
    
      Update �����ű� Set ������ = v_Maxno Where ��Ŀ��� = ���_In And ����ǰ׺ = v_Tag;
    End If;
    v_No := v_Maxno;
  
  Elsif ���_In = 125 Then
    --������Ϊ�˼��ٺ�����Ʊ������ʱ��
    If n_Mod = 0 Then
      Begin
        Select �Թܱ���
        Into v_�Թܱ���
        From ����ҽ����¼ A, ������ĿĿ¼ B
        Where a.������Ŀid = b.Id And a.Id = ����id_In And �Թܱ��� Is Not Null;
      Exception
        When Others Then
          v_Error := 'û���ҵ�������Ŀ��Ӧ�Ĺ��룡';
          Raise Err_Custom;
      End;
    Else
      Begin
        Select �Թܱ���, c.����
        Into v_�Թܱ���, v_����
        From ����ҽ����¼ A, ������ĿĿ¼ B, ���Ƽ������� C
        Where a.������Ŀid = b.Id And a.Id = ����id_In And b.�������� = c.���� And �Թܱ��� Is Not Null;
      
      Exception
        When Others Then
          v_Error := 'û���ҵ�������Ŀ��Ӧ�Ĺ���ͱ��룡';
          Raise Err_Custom;
      End;
    End If;
  
    Select Nvl(������, '0'), Nvl(��Ź���, 0)
    Into v_Maxno, n_Mod
    From ������Ʊ�
    Where ��Ŀ��� = ���_In
    For Update;
  
    If n_Mod = 0 Then
      --��ҽ������
    
      v_ҽ�� := ����id_In;
      If Length(v_ҽ��) > 12 Then
        v_ҽ�� := Substr(v_ҽ��, Length(v_ҽ��) - 11);
      Else
        v_ҽ�� := LPad(v_ҽ��, 12, '0');
      End If;
      if Length(ltrim(v_ҽ��,'0'))>8 then
         Select v_�Թܱ��� || LPad(Substr(v_ҽ��, Length(v_ҽ��) - (13 - Length(v_�Թܱ���) - 2)), (13 - Length(v_�Թܱ���)), '0')
          Into v_��������
         From Dual;
      else
         Select v_�Թܱ��� || LPad(Substr(v_ҽ��, Length(v_ҽ��) - (12 - Length(v_�Թܱ���) - 2)), (12 - Length(v_�Թܱ���)), '0')
         Into v_��������
         From Dual;
      end if;
      v_No := v_��������;
    Else
      --��"С���ţ�1λ��+����(2λ)+����(6λ)+˳���(3)λ"��������
      Begin
        Select ������
        Into v_Maxno
        From �����ű�
        Where ��Ŀ��� = ���_In And ����ǰ׺ = v_���� || v_�Թܱ��� And Trunc(����) = Trunc(Sysdate)
        For Update;
        v_Maxno := v_Maxno + 1;
        If Length(v_Maxno) <= 3 Then
          v_Maxno := LPad(v_Maxno, 3, '0');
        End If;
        v_No := v_���� || v_�Թܱ��� || To_Char(Trunc(Sysdate), 'yymmdd') || v_Maxno;
        Update �����ű�
        Set ���� = Trunc(Sysdate), ������ = v_Maxno
        Where ����ǰ׺ = v_���� || v_�Թܱ��� And Trunc(����) = Trunc(Sysdate);
      Exception
        When Others Then
          Update �����ű� Set ���� = Trunc(Sysdate), ������ = 1 Where ����ǰ׺ = v_���� || v_�Թܱ���;
          If Sql%RowCount = 0 Then
            Insert Into �����ű�
              (��Ŀ���, ����ǰ׺, ����, ������)
            Values
              (���_In, v_���� || v_�Թܱ���, Trunc(Sysdate), 1);
          End If;
          v_No := v_���� || v_�Թܱ��� || To_Char(Trunc(Sysdate), 'yymmdd') || '001';
      End;
    End If;
    --�����������
  Elsif ���_In = 126 Then
    --12λ˳���
    Select Nvl(������, '0') Into v_Maxno From ������Ʊ� Where ��Ŀ��� = ���_In For Update;
  
    If v_Maxno = '0' Then
      v_Maxno := '000000000001';
    Else
      v_Maxno := Trim(To_Char(To_Number(v_Maxno) + 1, '000000000000'));
    End If;
    Update ������Ʊ� Set ������ = v_Maxno Where ��Ŀ��� = ���_In;
  
    v_No := v_Maxno;
  Elsif ���_In = 135 Then
    --ҩƷ���ĵ���
    Begin
      Select Nvl(��Ź���, 0) Into n_Mod From ������Ʊ� Where ��Ŀ��� = ���_In;
    Exception
      When Others Then
        v_Error := '������Ʊ��в��������Ϊ' || ���_In || '�ĺ���';
        Raise Err_Custom;
    End;
    --1.����(YYYYMM)+˳���(0000)
    v_Tmp := To_Char(Sysdate, 'YYYYMM');
    Select Nvl(������, '0') Into v_Maxno From ������Ʊ� Where ��Ŀ��� = ���_In For Update;
  
    If v_Maxno = '0' Then
      v_Maxno := v_Tmp || '0001';
    Else
      If Substr(v_Maxno, 1, 6) = v_Tmp Then
        v_Maxno := v_Tmp || Trim(To_Char(To_Number(Substr(v_Maxno, 7, 4)) + 1, '0000'));
      Else
        v_Maxno := v_Tmp || '0001';
      End If;
    End If;
    Update ������Ʊ� Set ������ = v_Maxno Where ��Ŀ��� = ���_In;
    v_No := v_Maxno;
  Elsif ���_In = 136 Then
      --YYMMDD+5λ˳���(00000)
      Select To_Char(sysdate, 'yymmdd') Into v_Tmp From Dual;
      Select Nvl(������, '0') Into v_Maxno From ������Ʊ� Where ��Ŀ��� = ���_In For Update;
    
      If v_Maxno = '0' Then
        v_Maxno := v_Tmp || '00001';
      Else
        If Substr(v_Maxno, 1, 6) = v_Tmp Then
          If To_Number(Substr(v_Maxno, 7, 5)) = 99999 Then
            v_Maxno := v_Tmp || '00001';
          Else
            v_Maxno := v_Tmp || Trim(To_Char(To_Number(Substr(v_Maxno, 7, 5)) + 1, '00000'));
          End If;
        Else
          v_Maxno := v_Tmp || '00001';
        End If;
      End If;
      Update ������Ʊ� Set ������ = v_Maxno Where ��Ŀ��� = ���_In;
      v_No := v_Maxno;
  Elsif ���_In = 131 Then
    --��챨����
    Select Nvl(������, '0') Into v_Maxno From ������Ʊ� Where ��Ŀ��� = ���_In For Update;
    v_Maxno := To_Char(To_Number(v_Maxno) + 1);
    Update ������Ʊ� Set ������ = v_Maxno Where ��Ŀ��� = ���_In;
    v_No := v_Maxno;
    --�������ݺ�
  Else
    Begin
      Select Nvl(��Ź���, 0) Into n_Mod From ������Ʊ� Where ��Ŀ��� = ���_In;
    Exception
      When Others Then
        v_Error := '������Ʊ��в��������Ϊ' || ���_In || '�ĺ���';
        Raise Err_Custom;
    End;
  
    --��������ǰ����ұ��˳�����������δ���ÿ��ұ��룬���ȡ����˳����
    If n_Mod = 2 And ���_In <> 122 Then
      Begin
        Select ��� Into v_Deptno From ���Һ���� Where ����id = ����id_In And ��Ŀ��� = ���_In;
      Exception
        When Others Then
          Null;
      End;
      If v_Deptno Is Null Then
        n_Mod := 0;
      End If;
    End If;
  
    Select Decode(Sign(Intyear - 10), -1, To_Char(Intyear, '9'), Chr(55 + Intyear))
    Into v_Year
    From (Select To_Number(To_Char(Sysdate, 'yyyy'), '9999') - 1990 As Intyear From Dual);
  
    If n_Mod = 0 Then
      --0.����˳����
      Select Nvl(������, '0') Into v_No From ������Ʊ� Where ��Ŀ��� = ���_In For Update;
    
      --��2λ������ʽ:0-9,A-Z
      Select Bit1 || Decode(Sign(Ascii(Bit2) - Ascii('9')), -1, LPad(To_Number(Bit28) + 1, 7, '0'),
                             Decode(Bit38, '999999', Decode(Bit2, '9', 'A', Chr(Ascii(Bit2) + 1)) || '000000',
                                     Bit2 || LPad(To_Number(Bit38) + 1, 6, '0')))
      Into v_No
      From (Select Substr(Maxno, 1, 1) As Bit1, Substr(Maxno, 2, 1) As Bit2, Substr(Maxno, 2) As Bit28,
                    Substr(Maxno, 3) As Bit38
             From (Select Decode(v_No, '0', v_Year || '0000000',
                                   Decode(Sign(Ascii(Substr(v_No, 1, 1)) - Ascii(v_Year)), -1, v_Year || '0000000', v_No)) As Maxno
                    From Dual));
    
      Update ������Ʊ� Set ������ = v_No Where ��Ŀ��� = ���_In;
    
    Elsif n_Mod = 1 Then
      --1.����+��˳����:YDDD0000
      Select Nvl(������, '0') Into v_No From ������Ʊ� Where ��Ŀ��� = ���_In For Update;
      Select v_Year || LPad(Trunc(Sysdate - Trunc(Sysdate, 'YYYY') + 1, 0), 3, '0') || '0000' Into v_Maxno From Dual;
      If v_No < v_Maxno Then
        v_No := v_Maxno;
      End If;
      v_No := Substr(v_No, 1, 4) || LPad(To_Number(Substr(v_No, 5, 4)) + 1, 4, '0');
      Update ������Ʊ� Set ������ = v_No Where ��Ŀ��� = ���_In;
    
    Elsif n_Mod = 2 Then
      If ���_In = 122 Then
        --2.�����ұ���+YYMMDD+3λ˳���:2201090728001
        Select Count(*) Into n_Maxno From ���Һ���� Where ��Ŀ��� = ���_In And ����id = ����id_In;
        If Nvl(n_Maxno, 0) = 0 Then
          Insert Into ���Һ���� (��Ŀ���, ����id, ������, ���) Values (���_In, ����id_In, Null, Null);
          Commit;
        End If;
      
        Select ���� Into v_Deptno From ���ű� Where ID = ����id_In;
        Select Nvl(������, '-') Into v_No From ���Һ���� Where ��Ŀ��� = ���_In And ����id = ����id_In For Update;
        v_Tmp := To_Char(Sysdate, 'YYMMDD');
      
        If Substr(v_No, 1, Length(v_Deptno || v_Tmp)) = v_Deptno || v_Tmp Then
          v_No := v_Deptno || v_Tmp || LPad(To_Number(Substr(v_No, Length(v_Deptno || v_Tmp) + 1)) + 1, 3, '0');
        Else
          v_No := v_Deptno || v_Tmp || LPad('1', 3, '0');
        End If;
        Update ���Һ���� Set ������ = v_No Where ��Ŀ��� = ���_In And ����id = ����id_In;
      
      Else
        --2.����+���ұ��+��+˳���:YKDD0000
        Begin
          --����-��assciiΪ45,���ں�year�Ƚ�(0��asciiΪ48)
          Select ���, Nvl(������, '-')
          Into v_Deptno, v_No
          From ���Һ����
          Where ��Ŀ��� = ���_In And Nvl(����id, 0) = Nvl(����id_In, 0)
          For Update;
        Exception
          When Others Then
            Null;
        End;
        If v_Deptno Is Null Then
          v_Error := '����δ���ñ�ţ��޷��������룡';
          Raise Err_Custom;
        Else
          v_Tmp := To_Char(Sysdate, 'MM');
          Select Substr(Maxno, 1, 4) || LPad(To_Number(Substr(Maxno, 5, 4)) + 1, 4, '0')
          Into v_No
          From (Select Decode(Sign(Ascii(Substr(v_No, 1, 1)) - Ascii(v_Year)), -1, v_Year || v_Deptno || v_Tmp || '0000',
                                Decode(Sign(To_Number(Substr(v_No, 3, 2)) - To_Number(v_Tmp)), -1,
                                        v_Year || v_Deptno || v_Tmp || '0000', v_No)) As Maxno
                 From Dual);
          Update ���Һ���� Set ������ = v_No Where ��Ŀ��� = ���_In And Nvl(����id, 0) = Nvl(����id_In, 0);
        
        End If;
      End If;
    Elsif n_Mod = 3 Then
    
      --��������+000001����
      Select Substr(To_Char(Sysdate, 'yyyymmdd'), 3, 6) Into v_Tmp From Dual;
      Select Nvl(������, '0') Into v_Maxno From ������Ʊ� Where ��Ŀ��� = ���_In For Update;
    
      If v_Maxno = '0' Then
        v_Maxno := v_Tmp || '000001';
      Else
        If Substr(v_Maxno, 1, 6) = v_Tmp Then
          v_Maxno := v_Tmp || Trim(To_Char(To_Number(Substr(v_Maxno, 7, 6)) + 1, '000000'));
        Else
          v_Maxno := v_Tmp || '000001';
        End If;
      End If;
      Update ������Ʊ� Set ������ = v_Maxno Where ��Ŀ��� = ���_In;
      v_No := v_Maxno;
    
    Elsif n_Mod = 5 Then
      --1.����(YYYYMM)+˳���(000000)
      v_Tmp := To_Char(Sysdate, 'YYYYMM');
      Select Nvl(������, '0') Into v_Maxno From ������Ʊ� Where ��Ŀ��� = ���_In For Update;
    
      If v_Maxno = '0' Then
        v_Maxno := v_Tmp || '000001';
      Else
        If Substr(v_Maxno, 1, 6) = v_Tmp Then
          v_Maxno := v_Tmp || Trim(To_Char(To_Number(Substr(v_Maxno, 7, 6)) + 1, '000000'));
        Else
          v_Maxno := v_Tmp || '000001';
        End If;
      End If;
      Update ������Ʊ� Set ������ = v_Maxno Where ��Ŀ��� = ���_In;
      v_No := v_Maxno;
    Else
      v_Error := '���Ϊ' || ���_In || '�ĺ���,�����ֵ:' || n_Mod || ',��ǰϵͳ��֧�֣�';
      Raise Err_Custom;
    End If;
  End If;

  Commit;
  Return v_No;
Exception
  When Err_Custom Then
    Rollback;
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Rollback;
    zl_ErrorCenter(SQLCode, SQLErrM);
End Nextno;
/

--92808:����,2016-01-22,���ĵ��ݴ�ӡ����
Create Or Replace Procedure Zl_ҩƷ�շ�����_Insert
(
  No_In     In ҩƷ�շ�����.No%Type,
  ����_In   In ҩƷ�շ�����.����%Type,
  �ⷿid_In In ҩƷ�շ�����.�ⷿid%Type
) Is
  n_Count Number;
  n_Id    ҩƷ�շ�����.Id%Type;
Begin
  n_Count := 0;
  Begin
    Select 1 Into n_Count From ҩƷ�շ����� Where NO = No_In And ���� = ����_In And �ⷿid = �ⷿid_In;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If n_Count = 0 Then
    Select ҩƷ�շ�����_Id.Nextval Into n_Id From Dual;
    Insert Into ҩƷ�շ����� (ID, NO, ����, �ⷿid, ��ӡ״̬) Values (n_Id, No_In, ����_In, �ⷿid_In, 1);
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End;
/

--93146:������,2016-01-27,��ȡ��Դʣ����������
Create Or Replace Procedure Zl_Third_Docarrange
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --����:ҽ���Ű�ƻ�
  --���:Xml_In:
  --<IN>
  --   <YSID>870</YSID>    //ҽ��ID
  --   <KDID>870</KSID>    //����ID
  --   <KSSJ>2014-10-29 </KSSJ>    //��ʼʱ��
  --   <CXTS>14</CXTS>    //��ѯ����
  --   <HZDW>֧����</HZDW> //������λ
  --   <HL>����</HL>      //���࣬�ɴ�������ö��ŷָ�����ʽ:��ͨ,ר��,...
  --</IN>

  --����:Xml_Out
  --<OUTPUT>
  --   <PBLIST>       //δ���ظýڵ��ʾû������
  --    <PB>
  --     <RQ>2014-10-29</RQ>     //����
  --     <SYHS>5</SYHS>    //ʣ�����
  --     <SBSJ>ȫ��</SBSJ>             //�ϰ�ʱ��
  --     <YGS>5</YGS>    //�ѹҺ���
  --    </PB>
  --   <PBLIST>
  --   <ERROR><MSG></MSG></ERROR> //�����������
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  d_����         Date;
  v_�Ű�         �ҺŰ���.����%Type;
  n_�޺���       �ҺŰ�������.�޺���%Type;
  n_�ѹ���       �ҺŰ�������.�޺���%Type;
  n_���ѹ���     �ҺŰ�������.�޺���%Type;
  n_��Լ��       �ҺŰ�������.�޺���%Type;
  n_��Լ��       �ҺŰ�������.�޺���%Type;
  n_ʣ����       �ҺŰ�������.�޺���%Type;
  v_�ϰ�ʱ��     Varchar2(300);
  n_ҽ��id       ��Ա��.Id%Type;
  n_����id       ���ű�.Id%Type;
  n_��ѯ����     Number(4);
  n_������λ���� Number(5);
  n_��Լ�ѹ���   Number(4);
  n_��Լ����     Number(3);
  n_���Ŵ���     Number(3);
  v_����         �ҺŰ���.����%Type;
  n_����id       �ҺŰ��żƻ�.����id%Type;
  n_�ƻ�id       �ҺŰ��żƻ�.Id%Type;
  v_������λ     �Һź�����λ.����%Type;
  n_Daycount     Number(4);
  d_��ʼʱ��     Date;
  d_ԭʼʱ��     Date;
  n_����         Number(3);
  v_Temp         Varchar2(32767); --��ʱXML
  x_Templet      Xmltype; --ģ��XML
  v_Err_Msg      Varchar2(200);
  v_����         Varchar2(200);
  n_Exists       Number(2);
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/YSID'), Extractvalue(Value(A), 'IN/KSID'), Extractvalue(Value(A), 'IN/CXTS'),
         To_Date(Extractvalue(Value(A), 'IN/KSSJ'), 'YYYY-MM-DD hh24:mi:ss'), Extractvalue(Value(A), 'IN/HZDW'),
         Extractvalue(Value(A), 'IN/HL')
  Into n_ҽ��id, n_����id, n_��ѯ����, d_��ʼʱ��, v_������λ, v_����
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  n_��ѯ���� := Nvl(n_��ѯ����, 14);
  d_ԭʼʱ�� := Trunc(d_��ʼʱ��);
  d_��ʼʱ�� := Trunc(d_��ʼʱ��);
  n_Daycount := 0;
  If Nvl(n_����id, 0) = 0 Then
    While (n_Daycount < n_��ѯ����) Loop
      If Trunc(d_��ʼʱ�� + n_Daycount) = Trunc(Sysdate) Then
        d_��ʼʱ�� := Sysdate - n_Daycount;
      Else
        d_��ʼʱ�� := d_ԭʼʱ��;
      End If;
      n_���Ŵ��� := 0;
      v_�ϰ�ʱ�� := Null;
      n_���ѹ��� := 0;
      n_�ѹ���   := 0;
      n_ʣ����   := 0;
      n_�޺���   := 0;
      n_��Լ��   := 0;
      n_��Լ��   := 0;
      For r_�Ű� In (Select d_��ʼʱ�� + n_Daycount As ����, a.�Ű�, a.�޺���, a.��Լ��, Nvl(Hz.�ѹ���, 0) As �ѹ���, Nvl(Hz.��Լ��, 0) As ��Լ��,
                          a.����id, a.�ƻ�id, a.����, a.����
                   
                   From (Select Ap.����id, Ap.����, Bm.���� As ��������, Ap.ҽ������, Nvl(Ap.ҽ��id, 0) As ҽ��id, Ry.רҵ����ְ�� As ְ��, Ap.����,
                                 Ap.����id, Ap.�ƻ�id, Ap.�Ű�, Ap.��Ŀid, Fy.���� As ��Ŀ����, Ap.��ſ���, Ap.�޺���, Ap.��Լ��

                          
                          From (Select Ap.����id, Ap.����, Bm.���� As ��������, Ap.ҽ������, Ap.ҽ��id, Ap.����, Ap.Id As ����id, 0 As �ƻ�id,
                                        Ap.��Ŀid, Nvl(Ap.��ſ���, 0) As ��ſ���,
                                        Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', Ap.����, '2', Ap.��һ, '3', Ap.�ܶ�, '4',
                                                Ap.����, '5', Ap.����, '6', Ap.����, '7', Ap.����, Null) As �Ű�,
                                        Nvl(Xz.��Լ��, Xz.�޺���) As ��Լ��, Nvl(Xz.�޺���, 0) As �޺���
                                 From �ҺŰ��� Ap, ���ű� Bm, �ҺŰ������� Xz
                                 Where Ap.����id = Bm.Id(+) And Ap.ҽ��id = n_ҽ��id And Ap.ͣ������ Is Null And
                                       d_��ʼʱ�� + n_Daycount Between Nvl(Ap.��ʼʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                       Nvl(Ap.��ֹʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And Xz.����id(+) = Ap.Id And
                                       Xz.������Ŀ(+) = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�',
                                                           '4', '����', '5', '����', '6', '����', '7', '����', Null) And Not Exists
                                  (Select Rownum
                                        From �ҺŰ���ͣ��״̬ Ty
                                        Where Ty.����id = Ap.Id And d_��ʼʱ�� + n_Daycount Between Ty.��ʼֹͣʱ�� And Ty.����ֹͣʱ��) And
                                       Not Exists
                                  (Select Rownum
                                        From �ҺŰ��żƻ� Jh
                                        Where Jh.����id = Ap.Id And Jh.���ʱ�� Is Not Null And
                                              d_��ʼʱ�� + n_Daycount Between Nvl(Jh.��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                              Nvl(Jh.ʧЧʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')))
                                 Union All
                                 Select Ap.����id, Ap.����, Bm.���� As ��������, Jh.ҽ������, Jh.ҽ��id, Ap.����, Ap.Id As ����id, Jh.Id As �ƻ�id,
                                        Jh.��Ŀid, Nvl(Jh.��ſ���, 0) As ��ſ���,
                                        Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', Jh.����, '2', Jh.��һ, '3', Jh.�ܶ�, '4',
                                                Jh.����, '5', Jh.����, '6', Jh.����, '7', Jh.����, Null) As �Ű�,
                                        Nvl(Xz.��Լ��, Xz.�޺���) As ��Լ��, Nvl(Xz.�޺���, 0) As �޺���
                                 From �ҺŰ��żƻ� Jh, �ҺŰ��� Ap, ���ű� Bm, �Һżƻ����� Xz
                                 Where Jh.����id = Ap.Id And Ap.����id = Bm.Id(+) And Ap.ͣ������ Is Null And Jh.ҽ��id = n_ҽ��id And
                                       d_��ʼʱ�� + n_Daycount Between Nvl(Jh.��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                       Nvl(Jh.ʧЧʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And Xz.�ƻ�id(+) = Jh.Id And
                                       Xz.������Ŀ(+) = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�',
                                                           '4', '����', '5', '����', '6', '����', '7', '����', Null) And Not Exists
                                  (Select Rownum
                                        From �ҺŰ���ͣ��״̬ Ty
                                        Where Ty.����id = Ap.Id And d_��ʼʱ�� + n_Daycount Between Ty.��ʼֹͣʱ�� And Ty.����ֹͣʱ��) And
                                       (Jh.��Чʱ��, Jh.����id) =
                                       (Select Max(Sxjh.��Чʱ��) As ��Чʱ��, ����id
                                        From �ҺŰ��żƻ� Sxjh
                                        Where Sxjh.���ʱ�� Is Not Null And
                                              d_��ʼʱ�� + n_Daycount Between Nvl(Sxjh.��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                              Nvl(Sxjh.ʧЧʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And Sxjh.����id = Jh.����id
                                        Group By Sxjh.����id)) Ap, ���ű� Bm, ��Ա�� Ry, �շ���ĿĿ¼ Fy
                          Where Ap.����id = Bm.Id(+) And Ap.ҽ��id = Ry.Id(+) And Ap.��Ŀid = Fy.Id And �Ű� Is Not Null) A,
                        ���˹ҺŻ��� Hz, �շѼ�Ŀ B
                   Where a.���� = Hz.����(+) And Hz.����(+) = Trunc(d_��ʼʱ�� + n_Daycount) And a.��Ŀid = b.�շ�ϸĿid And
                         b.��ֹ���� > d_��ʼʱ�� + n_Daycount And b.ִ������ <= d_��ʼʱ�� + n_Daycount
                   
                   ) Loop
        If v_���� Is Null Or Instr(',' || v_���� || ',', ',' || r_�Ű�.���� || ',') > 0 Then
          v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_�Ű�.�Ű�;
          n_���ѹ��� := n_���ѹ��� + r_�Ű�.�ѹ���;
          n_�ѹ���   := r_�Ű�.�ѹ���;
          n_�޺���   := r_�Ű�.�޺���;
          n_��Լ��   := r_�Ű�.��Լ��;
          n_��Լ��   := r_�Ű�.��Լ��;
          n_����id   := Nvl(r_�Ű�.����id, 0);
          n_�ƻ�id   := Nvl(r_�Ű�.�ƻ�id, 0);
          v_����     := r_�Ű�.����;
          n_���Ŵ��� := 1;
          If v_�ϰ�ʱ�� Is Not Null Then
            If v_������λ Is Not Null Then
              If n_�ƻ�id <> 0 Then
                Begin
                  Select 1
                  Into n_��Լ����
                  From ������λ�ƻ�����
                  Where �ƻ�id = n_�ƻ�id And ������λ = v_������λ And Rownum < 2;
                Exception
                  When Others Then
                    n_��Լ���� := 0;
                End;
              Else
                Begin
                  Select 1
                  Into n_��Լ����
                  From ������λ���ſ���
                  Where ����id = n_����id And ������λ = v_������λ And Rownum < 2;
                Exception
                  When Others Then
                    n_��Լ���� := 0;
                End;
              End If;
            End If;
          
            If v_������λ Is Null Or Nvl(n_��Լ����, 0) = 0 Then
              If n_�ƻ�id <> 0 Then
                Begin
                  Select Sum(����)
                  Into n_������λ����
                  From ������λ�ƻ�����
                  Where �ƻ�id = n_�ƻ�id And
                        ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                      '����', '6', '����', '7', '����', Null);
                Exception
                  When Others Then
                    n_������λ���� := 0;
                End;
              Else
                Begin
                  Select Sum(����)
                  Into n_������λ����
                  From ������λ���ſ���
                  Where ����id = n_����id And
                        ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                      '����', '6', '����', '7', '����', Null);
                Exception
                  When Others Then
                    n_������λ���� := 0;
                End;
              End If;
              Begin
                Select Count(1)
                Into n_��Լ�ѹ���
                From ���˹Һż�¼
                Where �ű� = v_���� And ��¼״̬ = 1 And ������λ Is Not Null And ����ʱ�� Between Trunc(d_��ʼʱ�� + n_Daycount) And
                      Trunc(d_��ʼʱ�� + n_Daycount + 1) - 1 / 60 / 60 / 24;
              Exception
                When Others Then
                  n_��Լ�ѹ��� := 0;
              End;
              If n_������λ���� = 0 Then
                n_������λ���� := Null;
              End If;
              If Trunc(d_��ʼʱ�� + n_Daycount) > Trunc(Sysdate) Then
                n_ʣ���� := Nvl(n_ʣ����, 0) + n_��Լ�� - n_��Լ�� - Nvl(n_������λ����, n_��Լ�ѹ���) + Nvl(n_��Լ�ѹ���, 0);
              Else
                n_ʣ���� := Nvl(n_ʣ����, 0) + n_�޺��� - n_�ѹ��� - Nvl(n_������λ����, n_��Լ�ѹ���) + Nvl(n_��Լ�ѹ���, 0);
              End If;
            Else
              --��Լ��λ
              If n_�ƻ�id <> 0 Then
                Begin
                  Select 1
                  Into n_����
                  From ������λ�ƻ�����
                  Where �ƻ�id = n_�ƻ�id And
                        ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                      '����', '6', '����', '7', '����', Null) And ������λ = v_������λ And ���� = 0 And Rownum < 2;
                Exception
                  When Others Then
                    n_���� := 0;
                End;
              Else
                Begin
                  Select 1
                  Into n_����
                  From ������λ���ſ���
                  Where ����id = n_����id And
                        ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                      '����', '6', '����', '7', '����', Null) And ������λ = v_������λ And ���� = 0 And Rownum < 2;
                Exception
                  When Others Then
                    n_���� := 0;
                End;
              End If;
              If Nvl(n_����, 0) = 0 Then
                Begin
                  Select Count(1)
                  Into n_��Լ�ѹ���
                  From ���˹Һż�¼
                  Where �ű� = v_���� And ��¼״̬ = 1 And ������λ = v_������λ And ����ʱ�� Between Trunc(d_��ʼʱ�� + n_Daycount) And
                        Trunc(d_��ʼʱ�� + n_Daycount + 1) - 1 / 60 / 60 / 24;
                Exception
                  When Others Then
                    n_��Լ�ѹ��� := 0;
                End;
                If n_�ƻ�id <> 0 Then
                  Begin
                    Select Sum(����)
                    Into n_������λ����
                    From ������λ�ƻ�����
                    Where �ƻ�id = n_�ƻ�id And
                          ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                        '5', '����', '6', '����', '7', '����', Null) And ������λ = v_������λ;
                  Exception
                    When Others Then
                      n_������λ���� := 0;
                  End;
                Else
                  Begin
                    Select Sum(����)
                    Into n_������λ����
                    From ������λ���ſ���
                    Where ����id = n_����id And
                          ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                        '5', '����', '6', '����', '7', '����', Null) And ������λ = v_������λ;
                  Exception
                    When Others Then
                      n_������λ���� := 0;
                  End;
                End If;
                If n_������λ���� = 0 Then
                  n_������λ���� := Null;
                End If;
                n_ʣ���� := Nvl(n_ʣ����, 0) + Nvl(n_������λ����, n_��Լ�ѹ���) - Nvl(n_��Լ�ѹ���, 0);
              
              End If;
            End If;
          End If;
          n_������λ���� := 0;
          n_��Լ����     := 0;
          n_����         := 0;
        End If;
      End Loop;
      v_�ϰ�ʱ�� := Substr(v_�ϰ�ʱ��, 2);
      If n_���Ŵ��� = 1 Then
        v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_��ʼʱ�� + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                  '<SYHS>' || n_ʣ���� || '</SYHS>' || '<SBSJ>' || v_�ϰ�ʱ�� || '</SBSJ>' || '<YGS>' || n_���ѹ��� || '</YGS>' ||
                  '</PB>';
      End If;
      n_Daycount := n_Daycount + 1;
    End Loop;
  Else
    While (n_Daycount < n_��ѯ����) Loop
      If Trunc(d_��ʼʱ�� + n_Daycount) = Trunc(Sysdate) Then
        d_��ʼʱ�� := Sysdate - n_Daycount;
      Else
        d_��ʼʱ�� := d_ԭʼʱ��;
      End If;
      v_�ϰ�ʱ�� := Null;
      n_���ѹ��� := 0;
      n_�ѹ���   := 0;
      n_ʣ����   := 0;
      n_�޺���   := 0;
      n_��Լ��   := 0;
      n_��Լ��   := 0;
      n_���Ŵ��� := 0;
      For r_�Ű� In (Select d_��ʼʱ�� + n_Daycount As ����, a.�Ű�, a.�޺���, a.��Լ��, Nvl(Hz.�ѹ���, 0) As �ѹ���, Nvl(Hz.��Լ��, 0) As ��Լ��,
                          a.����id, a.�ƻ�id, a.����, a.����
                   
                   From (Select Ap.����id, Ap.����, Bm.���� As ��������, Ap.ҽ������, Nvl(Ap.ҽ��id, 0) As ҽ��id, Ry.רҵ����ְ�� As ְ��, Ap.����,
                                 Ap.����id, Ap.�ƻ�id, Ap.�Ű�, Ap.��Ŀid, Fy.���� As ��Ŀ����, Ap.��ſ���, Ap.�޺���, Ap.��Լ��

                          
                          From (Select Ap.����id, Ap.����, Bm.���� As ��������, Ap.ҽ������, Ap.ҽ��id, Ap.����, Ap.Id As ����id, 0 As �ƻ�id,
                                        Ap.��Ŀid, Nvl(Ap.��ſ���, 0) As ��ſ���,
                                        Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', Ap.����, '2', Ap.��һ, '3', Ap.�ܶ�, '4',
                                                Ap.����, '5', Ap.����, '6', Ap.����, '7', Ap.����, Null) As �Ű�,
                                        Nvl(Xz.��Լ��, Xz.�޺���) As ��Լ��, Nvl(Xz.�޺���, 0) As �޺���
                                 From �ҺŰ��� Ap, ���ű� Bm, �ҺŰ������� Xz
                                 Where Ap.����id = Bm.Id(+) And Ap.ҽ��id = n_ҽ��id And Ap.����id = n_����id And Ap.ͣ������ Is Null And
                                       d_��ʼʱ�� + n_Daycount Between Nvl(Ap.��ʼʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                       Nvl(Ap.��ֹʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And Xz.����id(+) = Ap.Id And
                                       Xz.������Ŀ(+) = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�',
                                                           '4', '����', '5', '����', '6', '����', '7', '����', Null) And Not Exists
                                  (Select Rownum
                                        From �ҺŰ���ͣ��״̬ Ty
                                        Where Ty.����id = Ap.Id And d_��ʼʱ�� + n_Daycount Between Ty.��ʼֹͣʱ�� And Ty.����ֹͣʱ��) And
                                       Not Exists
                                  (Select Rownum
                                        From �ҺŰ��żƻ� Jh
                                        Where Jh.����id = Ap.Id And Jh.���ʱ�� Is Not Null And
                                              d_��ʼʱ�� + n_Daycount Between Nvl(Jh.��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                              Nvl(Jh.ʧЧʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')))
                                 Union All
                                 Select Ap.����id, Ap.����, Bm.���� As ��������, Jh.ҽ������, Jh.ҽ��id, Ap.����, Ap.Id As ����id, Jh.Id As �ƻ�id,
                                        Jh.��Ŀid, Nvl(Jh.��ſ���, 0) As ��ſ���,
                                        Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', Jh.����, '2', Jh.��һ, '3', Jh.�ܶ�, '4',
                                                Jh.����, '5', Jh.����, '6', Jh.����, '7', Jh.����, Null) As �Ű�,
                                        Nvl(Xz.��Լ��, Xz.�޺���) As ��Լ��, Nvl(Xz.�޺���, 0) As �޺���
                                 From �ҺŰ��żƻ� Jh, �ҺŰ��� Ap, ���ű� Bm, �Һżƻ����� Xz
                                 Where Jh.����id = Ap.Id And Ap.����id = Bm.Id(+) And Ap.ͣ������ Is Null And Jh.ҽ��id = n_ҽ��id And
                                       Ap.����id = n_����id And
                                       d_��ʼʱ�� + n_Daycount Between Nvl(Jh.��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                       Nvl(Jh.ʧЧʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And Xz.�ƻ�id(+) = Jh.Id And
                                       Xz.������Ŀ(+) = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�',
                                                           '4', '����', '5', '����', '6', '����', '7', '����', Null) And Not Exists
                                  (Select Rownum
                                        From �ҺŰ���ͣ��״̬ Ty
                                        Where Ty.����id = Ap.Id And d_��ʼʱ�� + n_Daycount Between Ty.��ʼֹͣʱ�� And Ty.����ֹͣʱ��) And
                                       (Jh.��Чʱ��, Jh.����id) =
                                       (Select Max(Sxjh.��Чʱ��) As ��Чʱ��, ����id
                                        From �ҺŰ��żƻ� Sxjh
                                        Where Sxjh.���ʱ�� Is Not Null And
                                              d_��ʼʱ�� + n_Daycount Between Nvl(Sxjh.��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                                              Nvl(Sxjh.ʧЧʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And Sxjh.����id = Jh.����id
                                        Group By Sxjh.����id)) Ap, ���ű� Bm, ��Ա�� Ry, �շ���ĿĿ¼ Fy
                          Where Ap.����id = Bm.Id(+) And Ap.ҽ��id = Ry.Id(+) And Ap.��Ŀid = Fy.Id And �Ű� Is Not Null) A,
                        ���˹ҺŻ��� Hz, �շѼ�Ŀ B
                   Where a.���� = Hz.����(+) And Hz.����(+) = Trunc(d_��ʼʱ�� + n_Daycount) And a.��Ŀid = b.�շ�ϸĿid And
                         b.��ֹ���� > d_��ʼʱ�� + n_Daycount And b.ִ������ <= d_��ʼʱ�� + n_Daycount
                   
                   ) Loop
        If v_���� Is Null Or Instr(',' || v_���� || ',', ',' || r_�Ű�.���� || ',') > 0 Then
          v_�ϰ�ʱ�� := v_�ϰ�ʱ�� || '+' || r_�Ű�.�Ű�;
          n_���ѹ��� := n_���ѹ��� + r_�Ű�.�ѹ���;
          n_�ѹ���   := r_�Ű�.�ѹ���;
          n_�޺���   := r_�Ű�.�޺���;
          n_��Լ��   := r_�Ű�.��Լ��;
          n_��Լ��   := r_�Ű�.��Լ��;
          n_����id   := Nvl(r_�Ű�.����id, 0);
          n_�ƻ�id   := Nvl(r_�Ű�.�ƻ�id, 0);
          v_����     := r_�Ű�.����;
          n_���Ŵ��� := 1;
        
          If v_�ϰ�ʱ�� Is Not Null Then
            If v_������λ Is Not Null Then
              If n_�ƻ�id <> 0 Then
                Begin
                  Select 1
                  Into n_��Լ����
                  From ������λ�ƻ�����
                  Where �ƻ�id = n_�ƻ�id And ������λ = v_������λ And Rownum < 2;
                Exception
                  When Others Then
                    n_��Լ���� := 0;
                End;
              Else
                Begin
                  Select 1
                  Into n_��Լ����
                  From ������λ���ſ���
                  Where ����id = n_����id And ������λ = v_������λ And Rownum < 2;
                Exception
                  When Others Then
                    n_��Լ���� := 0;
                End;
              End If;
            End If;
          
            If v_������λ Is Null Or Nvl(n_��Լ����, 0) = 0 Then
              If n_�ƻ�id <> 0 Then
                Begin
                  Select Sum(����)
                  Into n_������λ����
                  From ������λ�ƻ�����
                  Where �ƻ�id = n_�ƻ�id And
                        ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                      '����', '6', '����', '7', '����', Null);
                Exception
                  When Others Then
                    n_������λ���� := 0;
                End;
              Else
                Begin
                  Select Sum(����)
                  Into n_������λ����
                  From ������λ���ſ���
                  Where ����id = n_����id And
                        ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                      '����', '6', '����', '7', '����', Null);
                Exception
                  When Others Then
                    n_������λ���� := 0;
                End;
              End If;
              Begin
                Select Count(1)
                Into n_��Լ�ѹ���
                From ���˹Һż�¼
                Where �ű� = v_���� And ��¼״̬ = 1 And ������λ Is Not Null And ����ʱ�� Between Trunc(d_��ʼʱ�� + n_Daycount) And
                      Trunc(d_��ʼʱ�� + n_Daycount + 1) - 1 / 60 / 60 / 24;
              Exception
                When Others Then
                  n_��Լ�ѹ��� := 0;
              End;
              If n_������λ���� = 0 Then
                n_������λ���� := Null;
              End If;
              If Trunc(d_��ʼʱ�� + n_Daycount) > Trunc(Sysdate) Then
                n_ʣ���� := Nvl(n_ʣ����, 0) + n_��Լ�� - n_��Լ�� - Nvl(n_������λ����, n_��Լ�ѹ���) + Nvl(n_��Լ�ѹ���, 0);
              Else
                n_ʣ���� := Nvl(n_ʣ����, 0) + n_�޺��� - n_�ѹ��� - Nvl(n_������λ����, n_��Լ�ѹ���) + Nvl(n_��Լ�ѹ���, 0);
              End If;
            Else
              --��Լ��λ
              If n_�ƻ�id <> 0 Then
                Begin
                  Select 1
                  Into n_����
                  From ������λ�ƻ�����
                  Where �ƻ�id = n_�ƻ�id And
                        ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                      '����', '6', '����', '7', '����', Null) And ������λ = v_������λ And ���� = 0 And Rownum < 2;
                Exception
                  When Others Then
                    n_���� := 0;
                End;
              Else
                Begin
                  Select 1
                  Into n_����
                  From ������λ���ſ���
                  Where ����id = n_����id And
                        ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                      '����', '6', '����', '7', '����', Null) And ������λ = v_������λ And ���� = 0 And Rownum < 2;
                Exception
                  When Others Then
                    n_���� := 0;
                End;
              End If;
              If Nvl(n_����, 0) = 0 Then
                Begin
                  Select Count(1)
                  Into n_��Լ�ѹ���
                  From ���˹Һż�¼
                  Where �ű� = v_���� And ��¼״̬ = 1 And ������λ = v_������λ And ����ʱ�� Between Trunc(d_��ʼʱ�� + n_Daycount) And
                        Trunc(d_��ʼʱ�� + n_Daycount + 1) - 1 / 60 / 60 / 24;
                Exception
                  When Others Then
                    n_��Լ�ѹ��� := 0;
                End;
                If n_�ƻ�id <> 0 Then
                  Begin
                    Select Sum(����)
                    Into n_������λ����
                    From ������λ�ƻ�����
                    Where �ƻ�id = n_�ƻ�id And
                          ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                        '5', '����', '6', '����', '7', '����', Null) And ������λ = v_������λ;
                  Exception
                    When Others Then
                      n_������λ���� := 0;
                  End;
                Else
                  Begin
                    Select Sum(����)
                    Into n_������λ����
                    From ������λ���ſ���
                    Where ����id = n_����id And
                          ������Ŀ = Decode(To_Char(d_��ʼʱ�� + n_Daycount, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����',
                                        '5', '����', '6', '����', '7', '����', Null) And ������λ = v_������λ;
                  Exception
                    When Others Then
                      n_������λ���� := 0;
                  End;
                End If;
                If n_������λ���� = 0 Then
                  n_������λ���� := Null;
                End If;
                n_ʣ���� := Nvl(n_ʣ����, 0) + Nvl(n_������λ����, n_��Լ�ѹ���) - Nvl(n_��Լ�ѹ���, 0);
              
              End If;
            End If;
          End If;
          n_������λ���� := 0;
          n_��Լ����     := 0;
          n_����         := 0;
        End If;
      End Loop;
      v_�ϰ�ʱ�� := Substr(v_�ϰ�ʱ��, 2);
      If n_���Ŵ��� = 1 Then
        v_Temp := v_Temp || '<PB>' || '<RQ>' || To_Char(Trunc(d_��ʼʱ�� + n_Daycount), 'YYYY-MM-DD') || '</RQ>' ||
                  '<SYHS>' || n_ʣ���� || '</SYHS>' || '<SBSJ>' || v_�ϰ�ʱ�� || '</SBSJ>' || '<YGS>' || n_���ѹ��� || '</YGS>' ||
                  '</PB>';
      End If;
      n_Daycount := n_Daycount + 1;
    End Loop;
  End If;
  v_Temp := '<PBLIST>' || v_Temp || '</PBLIST>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Docarrange;
/

--93052:������,2016-02-02,���񴰽������⴦��
Create Or Replace Procedure Zl_Third_Getsettlement
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --����:��ȡHIS��������
  --���:Xml_In:
  --<IN>
  -- <BRID></BRID>       //����ID 
  -- <ZYID></ZYID>         //��ҳID
  -- <JSLX></JSLX>       //�������͡�1-����,2-סԺ���̶���2
  -- <JSKLB></JSKLB>       //���㿨���
  --</IN>
  --����:Xml_Out
  --<OUTPUT>
  --<JBXX>              //������Ϣ
  --   <XM></XM>           //����
  --   <XB></XB>           //�Ա�
  --   <NL></NL>         //����
  --   <ZYH></ ZYH>        //סԺ��
  --   <ZYKS></ ZYKS>          //סԺ����  
  --   <KSID></KSID>         //����ID
  --   <ZZYS></ ZZYS>          //����ҽ��  
  --   <RYSJ></ RYSJ>          //��Ժʱ��
  --   <CYSJ></ CYSJ >         //��Ժʱ�� 
  --   <JZSJ></JZSJ>         //����ʱ��(δ����Ϊ��)
  --   <DJH></DJH>         //���ݺ�(δ����Ϊ��)
  --   <JSZFY></JSZFY>         //�����ܷ���
  --</JBXX>
  --<YJKLIST>              //���Ԥ�ɿ��
  --   <ITEM>
  --     <DJH><DJH>        //Ԥ����ݺ�
  --     <JSFS></JSFS>     //���㷽ʽ��Ϊ���ƣ�����ʲô��ȡʲô��
  --     <JE></JE>           //Ԥ�ɿ���
  --     <JYLSH></JYLSH>       //������ˮ�ţ����ڳ���ʹ�ã�
  --     <SFJSK></SFJSK>       //�Ƿ���㿨��1-�ǣ�0-��������ɴ���Ŀ����ɷѣ�����1�����򷵻�0
  --   </ITEM>
  --</YJKLIST >
  --<TBQK>               //�˲����
  --   <TBLX></TBLX>         //�˲�����(1:���˲��2:ҽԺ�˿�)
  --   <TBJE></TBJE>         //�˲����
  --</TBQK>
  -- <ERROR><MSG></MSG></ERROR>    //���ִ���ʱ���ؾ���ԭ��error�ڵ�Ϊ�ձ�ʾ�ɹ�
  --</OUTPUT>  

  --------------------------------------------------------------------------------------------------
  n_����id     ������Ϣ.����id%Type;
  n_��ҳid     ������ҳ.��ҳid%Type;
  n_��������   Number(3);
  v_���㿨��� Varchar2(200);
  n_�����id   ҽ�ƿ����.Id%Type;
  n_�Ƿ����   Number(3); -- 1-δ����,0-����
  n_���ʽ��   סԺ���ü�¼.���ʽ��%Type;
  v_Temp       Varchar2(32767); --��ʱXML
  v_Subtemp    Varchar2(32767);
  v_����ids    Varchar2(5000);
  n_�˲����   ����Ԥ����¼.��Ԥ��%Type;
  n_�������   ����Ԥ����¼.���%Type;
  n_����id     ����Ԥ����¼.����id%Type;
  x_Templet    Xmltype; --ģ��XML
  v_Err_Msg    Varchar2(200);
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/BRID'), Extractvalue(Value(A), 'IN/ZYID'), Extractvalue(Value(A), 'IN/JSLX'),
         Extractvalue(Value(A), 'IN/JSKLB')
  Into n_����id, n_��ҳid, n_��������, v_���㿨���
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  --Ĭ��סԺ����
  n_�������� := Nvl(n_��������, 2);
  Begin
    Select ID Into n_�����id From ҽ�ƿ���� Where ���� = v_���㿨���;
  Exception
    When Others Then
      v_Err_Msg := '�޷�ȷ�ϴ���Ľ��㿨,����!';
      Raise Err_Item;
  End;
  If n_�������� = 2 Then
    Begin
      Select Distinct 1
      Into n_�Ƿ����
      From סԺ���ü�¼
      Where ����id = n_����id And ��¼״̬ <> 0 And ��ҳid = n_��ҳid And ���ʷ��� = 1 Having
       Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) <> 0;
    Exception
      When Others Then
        n_�Ƿ���� := 0;
    End;
    If n_�Ƿ���� = 0 Then
      --����,��ȡ��������
      For r_���� In (Select ����, �Ա�, ����, סԺ��, סԺ����, ����id, ����ҽ��, To_Char(��Ժʱ��, 'yyyy-mm-dd') As ��Ժʱ��,
                          To_Char(��Ժʱ��, 'yyyy-mm-dd') As ��Ժʱ��, To_Char(����ʱ��, 'yyyy-mm-dd') As ����ʱ��, ���ݺ�, �����ܷ���, ����id
                   From (Select c.����, c.�Ա�, c.����, c.סԺ��, e.���� As סԺ����, c.��Ժ����id As ����id, c.סԺҽʦ As ����ҽ��, c.��Ժ���� As ��Ժʱ��,
                                 c.��Ժ���� As ��Ժʱ��, a.�շ�ʱ�� As ����ʱ��, a.No As ���ݺ�, Sum(d.��Ԥ��) As �����ܷ���, a.Id As ����id
                          From ���˽��ʼ�¼ A, ������Ϣ B, ������ҳ C, ����Ԥ����¼ D, ���ű� E
                          Where a.��¼״̬ = 1 And a.����id = c.����id And a.����id = b.����id And c.��ҳid = n_��ҳid And a.����id = n_����id And
                                d.����id = a.Id And c.��Ժ����id = e.Id(+) And Exists
                           (Select 1 From ����Ԥ����¼ Where ����id = a.Id And ���㷽ʽ = v_���㿨��� And ��ҳid = n_��ҳid)
                          Group By c.����, c.�Ա�, c.����, c.סԺ��, e.����, c.��Ժ����id, c.סԺҽʦ, c.��Ժ����, c.��Ժ����, a.�շ�ʱ��, a.No, a.Id
                          Order By ����ʱ�� Desc)
                   Where Rownum < 2) Loop
        v_Temp := '<XM>' || r_����.���� || '</XM>';
        v_Temp := v_Temp || '<XB>' || r_����.�Ա� || '</XB>';
        v_Temp := v_Temp || '<NL>' || r_����.���� || '</NL>';
        v_Temp := v_Temp || '<ZYH>' || r_����.סԺ�� || '</ZYH>';
        v_Temp := v_Temp || '<ZYKS>' || r_����.סԺ���� || '</ZYKS>';
        v_Temp := v_Temp || '<KSID>' || r_����.����id || '</KSID>';
        v_Temp := v_Temp || '<ZZYS>' || r_����.����ҽ�� || '</ZZYS>';
        v_Temp := v_Temp || '<RYSJ>' || r_����.��Ժʱ�� || '</RYSJ>';
        v_Temp := v_Temp || '<CYSJ>' || r_����.��Ժʱ�� || '</CYSJ>';
        v_Temp := v_Temp || '<JZSJ>' || r_����.����ʱ�� || '</JZSJ>';
        v_Temp := v_Temp || '<DJH>' || r_����.���ݺ� || '</DJH>';
        v_Temp := v_Temp || '<JSZFY>' || r_����.�����ܷ��� || '</JSZFY>';
        v_Temp := '<JBXX>' || v_Temp || '</JBXX>';
        Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
        n_����id := r_����.����id;
      End Loop;
      If n_����id Is Null Then
        v_Err_Msg := '�ò���û�н�������!';
        Raise Err_Item;
      End If;
      v_Temp := '';
      For r_Ԥ�� In (Select NO As ���ݺ�, ���㷽ʽ, Sum(��Ԥ��) As ���, ������ˮ��, Max(�����id) As �����id
                   From ����Ԥ����¼
                   Where ����id = n_����id And Mod(��¼����, 10) = 1
                   Group By NO, ���㷽ʽ, ������ˮ��
                   Order By ���ݺ� Desc) Loop
        v_Temp := '<DJH>' || r_Ԥ��.���ݺ� || '</DJH>';
        v_Temp := v_Temp || '<JSFS>' || r_Ԥ��.���㷽ʽ || '</JSFS>';
        v_Temp := v_Temp || '<JE>' || r_Ԥ��.��� || '</JE>';
        v_Temp := v_Temp || '<JYLSH>' || r_Ԥ��.������ˮ�� || '</JYLSH>';
        If n_�����id = r_Ԥ��.�����id Then
          v_Temp := v_Temp || '<SFJSK>' || 1 || '</SFJSK>';
        Else
          v_Temp := v_Temp || '<SFJSK>' || 0 || '</SFJSK>';
        End If;
        v_Temp    := '<ITEM>' || v_Temp || '</ITEM>';
        v_Subtemp := v_Subtemp || v_Temp;
      End Loop;
      v_Subtemp := '<YJKLIST>' || v_Subtemp || '</YJKLIST>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Subtemp)) Into x_Templet From Dual;
      Select Nvl(Sum(��Ԥ��), 0)
      Into n_�˲����
      From ����Ԥ����¼
      Where ����id = n_����id And Mod(��¼����, 10) = 2 And Nvl(У�Ա�־, 0) = 0;
      If n_�˲���� < 0 Then
        v_Temp := '<TBLX>' || 2 || '</TBLX>';
      Else
        v_Temp := '<TBLX>' || 1 || '</TBLX>';
      End If;
      v_Temp := v_Temp || '<TBJE>' || Abs(n_�˲����) || '</TBJE>';
      v_Temp := '<TBQK>' || v_Temp || '</TBQK>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    Else
      --δ���壬��ȡδ������
      For r_Info In (Select c.����, c.�Ա�, c.����, c.סԺ��, d.���� As סԺ����, c.��Ժ����id As ����id, c.סԺҽʦ As ����ҽ��,
                            To_Char(c.��Ժ����, 'yyyy-mm-dd') As ��Ժʱ��, To_Char(c.��Ժ����, 'yyyy-mm-dd') As ��Ժʱ��
                     From ������ҳ C, ���ű� D
                     Where c.����id = n_����id And c.��Ժ����id = d.Id(+) And c.��ҳid = n_��ҳid And Rownum < 2) Loop
        v_Temp := '<XM>' || r_Info.���� || '</XM>';
        v_Temp := v_Temp || '<XB>' || r_Info.�Ա� || '</XB>';
        v_Temp := v_Temp || '<NL>' || r_Info.���� || '</NL>';
        v_Temp := v_Temp || '<ZYH>' || r_Info.סԺ�� || '</ZYH>';
        v_Temp := v_Temp || '<ZYKS>' || r_Info.סԺ���� || '</ZYKS>';
        v_Temp := v_Temp || '<KSID>' || r_Info.����id || '</KSID>';
        v_Temp := v_Temp || '<ZZYS>' || r_Info.����ҽ�� || '</ZZYS>';
        v_Temp := v_Temp || '<RYSJ>' || r_Info.��Ժʱ�� || '</RYSJ>';
        v_Temp := v_Temp || '<CYSJ>' || r_Info.��Ժʱ�� || '</CYSJ>';
        v_Temp := v_Temp || '<JZSJ>' || '' || '</JZSJ>';
        v_Temp := v_Temp || '<DJH>' || '' || '</DJH>';
      End Loop;
      Begin
        Select Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0))
        Into n_���ʽ��
        From סԺ���ü�¼
        Where ����id = n_����id And ��¼״̬ <> 0 And ��ҳid = n_��ҳid And ���ʷ��� = 1;
      Exception
        When Others Then
          n_���ʽ�� := 0;
      End;
      v_Temp := v_Temp || '<JSZFY>' || n_���ʽ�� || '</JSZFY>';
      v_Temp := '<JBXX>' || v_Temp || '</JBXX>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
      v_Subtemp := '';
      For r_Ԥ�� In (Select NO As ���ݺ�, ���㷽ʽ, Sum(Nvl(���, 0)) - Sum(Nvl(��Ԥ��, 0)) As ���, ������ˮ��, Max(�����id) As �����id
                   From ����Ԥ����¼
                   Where ����id = n_����id And Mod(��¼����, 10) = 1 And Nvl(Ԥ�����, 2) = 2 And (��ҳid = n_��ҳid Or ��ҳid Is Null)
                   Group By NO, ���㷽ʽ, ������ˮ��
                   Having Sum(Nvl(���, 0)) - Sum(Nvl(��Ԥ��, 0)) <> 0) Loop
        v_Temp := '<DJH>' || r_Ԥ��.���ݺ� || '</DJH>';
        v_Temp := v_Temp || '<JSFS>' || r_Ԥ��.���㷽ʽ || '</JSFS>';
        v_Temp := v_Temp || '<JE>' || r_Ԥ��.��� || '</JE>';
        v_Temp := v_Temp || '<JYLSH>' || r_Ԥ��.������ˮ�� || '</JYLSH>';
        If n_�����id = r_Ԥ��.�����id Then
          v_Temp := v_Temp || '<SFJSK>' || 1 || '</SFJSK>';
        Else
          v_Temp := v_Temp || '<SFJSK>' || 0 || '</SFJSK>';
        End If;
        v_Temp     := '<ITEM>' || v_Temp || '</ITEM>';
        v_Subtemp  := v_Subtemp || v_Temp;
        n_������� := Nvl(n_�������, 0) + Nvl(r_Ԥ��.���, 0);
      End Loop;
      v_Subtemp := '<YJKLIST>' || v_Subtemp || '</YJKLIST>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Subtemp)) Into x_Templet From Dual;
      If Nvl(n_�������,0) - Nvl(n_���ʽ��,0) > 0 Then
        v_Temp := '<TBLX>' || 2 || '</TBLX>';
      Else
        v_Temp := '<TBLX>' || 1 || '</TBLX>';
      End If;
      v_Temp := v_Temp || '<TBJE>' || Abs(Nvl(n_�������,0) - Nvl(n_���ʽ��,0)) || '</TBJE>';
      v_Temp := '<TBQK>' || v_Temp || '</TBQK>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    End If;
  End If;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getsettlement;
/

--93052:������,2016-01-25,֧��������У�Ա�־����
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
  n_����       Number(3);
  n_Ԥ��id     ����Ԥ����¼.Id%Type;
  n_����id     ������ҳ.��Ժ����id%Type;
  d_��������   Date;
  n_���㿨��� �����ѽӿ�Ŀ¼.���%Type;
  n_ʱ������   Number(3);
  v_Ids        Varchar2(20000);
  v_���ѿ����� Varchar2(5000);
  v_No         ���˽��ʼ�¼.No%Type;
  v_Ԥ��no     ����Ԥ����¼.No%Type;
  n_�����id   ҽ�ƿ����.Id%Type;
  v_���㿨��   ����Ԥ����¼.����%Type;
  v_���㷽ʽ   ����Ԥ����¼.���㷽ʽ%Type;
  v_Temp       Varchar2(500);
  x_Templet    Xmltype; --ģ��XML
  v_Err_Msg    Varchar2(200);
  Err_Item Exception;
  n_Count Number(18);

Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/ZYID'), To_Number(Extractvalue(Value(A), 'IN/BRID')),
         To_Number(Extractvalue(Value(A), 'IN/JE')), To_Number(Extractvalue(Value(A), 'IN/JSLX'))
  Into n_��ҳid, n_����id, n_�����ܶ�, n_��������
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  n_�������� := Nvl(n_��������, 2);

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
  If n_�������� = 2 Then
    Begin
      Select Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0))
      Into n_�����ʽ��
      From סԺ���ü�¼
      Where ����id = n_����id And ��¼״̬ <> 0 And ��ҳid = n_��ҳid And ���ʷ��� = 1;
    Exception
      When Others Then
        n_�����ʽ�� := 0;
    End;
  
    If n_�����ʽ�� <> n_�����ܶ� Then
      v_Err_Msg := '����Ľ��ʽ����ʵ�ʽ��ʽ���,���������!';
      Raise Err_Item;
    End If;
    Begin
      Select ��Ժ����id Into n_����id From ������ҳ Where ����id = n_����id And ��ҳid = n_��ҳid;
    Exception
      When Others Then
        n_����id := Null;
    End;
  
    Select ���˽��ʼ�¼_Id.Nextval, Sysdate, Nextno(15) Into n_����id, d_����ʱ��, v_No From Dual;
  
    n_ʱ������ := zl_GetSysParameter('���ʷ���ʱ��', 1137);
    If n_ʱ������ = 0 Then
      --���Ǽ�ʱ��
      Select Trunc(Min(�Ǽ�ʱ��)), Trunc(Max(�Ǽ�ʱ��))
      Into d_��ʼ����, d_��������
      From (Select NO, ���, �Ǽ�ʱ��, ����ʱ��
             From סԺ���ü�¼
             Where ����id = n_����id And ��¼״̬ <> 0 And ��ҳid = n_��ҳid And ���ʷ��� = 1
             Group By NO, ���, �Ǽ�ʱ��, ����ʱ��, Mod(��¼����, 10)
             Having Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) <> 0);
    Else
      --������ʱ��  
      Select Trunc(Min(����ʱ��)), Trunc(Max(����ʱ��))
      Into d_��ʼ����, d_��������
      From (Select NO, ���, �Ǽ�ʱ��, ����ʱ��
             From סԺ���ü�¼
             Where ����id = n_����id And ��¼״̬ <> 0 And ��ҳid = n_��ҳid And ���ʷ��� = 1
             Group By NO, ���, �Ǽ�ʱ��, ����ʱ��, Mod(��¼����, 10)
             Having Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) <> 0);
    End If;
  
    Zl_���˽��ʼ�¼_Insert(n_����id, v_No, n_����id, d_����ʱ��, d_��ʼ����, d_��������, 0, 0, n_��ҳid, Null, 2, Null, 2);
  
    For r_���� In (Select Min(ID) As ID, Mod(��¼����, 10) As ��¼����, NO, ���, ��¼״̬, ִ��״̬,
                        Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) As ���, Sum(Nvl(���ʽ��, 0)) As ���ʽ��
                 From סԺ���ü�¼
                 Where ����id = n_����id And ��¼״̬ <> 0 And ��ҳid = n_��ҳid And ���ʷ��� = 1
                 Group By Mod(��¼����, 10), NO, ���, ��¼״̬, ִ��״̬
                 Having Sum(Nvl(ʵ�ս��, 0)) - Sum(Nvl(���ʽ��, 0)) <> 0
                 Order By NO, ���) Loop
      If Nvl(r_����.���ʽ��, 0) = 0 Then
        Begin
          Select 1 Into n_���� From סԺ���ü�¼ Where ID = r_����.Id And ����id Is Null;
        Exception
          When Others Then
            n_���� := 0;
        End;
        If n_���� = 1 Then
          v_Ids := v_Ids || ',' || r_����.Id;
        Else
          Zl_���ʷ��ü�¼_Insert(0, r_����.No, r_����.��¼����, r_����.��¼״̬, r_����.ִ��״̬, r_����.���, r_����.���, n_����id);
        End If;
      Else
        Zl_���ʷ��ü�¼_Insert(0, r_����.No, r_����.��¼����, r_����.��¼״̬, r_����.ִ��״̬, r_����.���, r_����.���, n_����id);
      End If;
    End Loop;
  
    If v_Ids Is Not Null Then
      v_Ids := Substr(v_Ids, 2);
      Zl_���ʷ��ü�¼_Batch(v_Ids, n_����id, n_����id);
    End If;
  
    n_Count := 0;
    For r_���㷽ʽ In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As ���㿨���,
                          Extractvalue(b.Column_Value, '/JS/JSKH') As ���㿨��,
                          Extractvalue(b.Column_Value, '/JS/JSFS') As ���㷽ʽ,
                          Extractvalue(b.Column_Value, '/JS/JSJE') As ������,
                          Extractvalue(b.Column_Value, '/JS/JYLSH') As ������ˮ��,
                          Extractvalue(b.Column_Value, '/JS/JYSM') As ����˵��, Extractvalue(b.Column_Value, '/JS/ZY') As ժҪ,
                          Extractvalue(b.Column_Value, '/JS/SFCYJ') As �Ƿ��Ԥ��,
                          Extractvalue(b.Column_Value, '/JS/SFXFK') As �Ƿ����ѿ�
                   From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
      If Nvl(r_���㷽ʽ.�Ƿ��Ԥ��, 0) = 0 Then
        --����
        If n_Count = 1 Then
          v_Err_Msg := '���ʽ����ݲ�֧�ֶ��ֽ��㷽ʽ!';
          Raise Err_Item;
        End If;
        If Nvl(r_���㷽ʽ.�Ƿ����ѿ�, 0) = 1 Then
          Begin
            n_���㿨��� := To_Number(r_���㷽ʽ.���㿨���);
          Exception
            When Others Then
              n_���㿨��� := 0;
          End;
          If n_���㿨��� = 0 Then
            Begin
              Select ���
              Into n_���㿨���
              From �����ѽӿ�Ŀ¼
              Where ���� = r_���㷽ʽ.���㿨��� And Nvl(����, 0) = 1;
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
            n_�����id := To_Number(r_���㷽ʽ.���㿨���);
          Exception
            When Others Then
              n_�����id := 0;
          End;
          If n_�����id = 0 Then
            Begin
              Select ID Into n_�����id From ҽ�ƿ���� Where ���� = r_���㷽ʽ.���㿨��� And Nvl(�Ƿ�����, 0) = 1;
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
      
        If n_�����id Is Not Null Then
          --������,����סԺԤ����
          v_���㿨�� := r_���㷽ʽ.���㿨��;
          If r_���㷽ʽ.������ > 0 Then
            Select ����Ԥ����¼_Id.Nextval, Nextno(11) Into n_Ԥ��id, v_Ԥ��no From Dual;
            Zl_����Ԥ����¼_Insert(n_Ԥ��id, v_Ԥ��no, Null, n_����id, n_��ҳid, n_����id, r_���㷽ʽ.������, v_���㷽ʽ, '', '', '', '', '',
                             v_����Ա����, v_����Ա����, Null, 2, n_�����id, Null, r_���㷽ʽ.���㿨��, r_���㷽ʽ.������ˮ��, r_���㷽ʽ.����˵��, Null,
                             d_����ʱ��, 0);
            n_Ԥ����ֵ := Nvl(n_Ԥ����ֵ, 0) + r_���㷽ʽ.������;
          Else
          
            Zl_���ʽɿ��¼_Insert(v_No, n_����id, n_��ҳid, n_����id, v_���㷽ʽ, Null, r_���㷽ʽ.������, n_����id, v_����Ա����, v_����Ա����, d_����ʱ��,
                             Null, Null, Null, Null, Null, n_�����id, r_���㷽ʽ.���㿨��, r_���㷽ʽ.������ˮ��, r_���㷽ʽ.����˵��);
          
          End If;
        Else
          If n_���㿨��� Is Not Null Then
            --���ѿ�
            v_���㿨��   := r_���㷽ʽ.���㿨��;
            v_���ѿ����� := n_���㿨��� || '|' || r_���㷽ʽ.���㿨�� || '|0|' || r_���㷽ʽ.������ || '||';
            Zl_���ʽɿ��¼_Insert(v_No, n_����id, n_��ҳid, n_����id, v_���㷽ʽ, Null, r_���㷽ʽ.������, n_����id, v_����Ա����, v_����Ա����, d_����ʱ��,
                             Null, Null, Null, Null, Null, n_�����id, r_���㷽ʽ.���㿨��, r_���㷽ʽ.������ˮ��, r_���㷽ʽ.����˵��, v_���ѿ�����);
          Else
            --��������
            Zl_���ʽɿ��¼_Insert(v_No, n_����id, n_��ҳid, n_����id, v_���㷽ʽ, Null, r_���㷽ʽ.������, n_����id, v_����Ա����, v_����Ա����, d_����ʱ��,
                             Null, Null, Null, Null, Null, n_�����id, r_���㷽ʽ.���㿨��, r_���㷽ʽ.������ˮ��, r_���㷽ʽ.����˵��);
          End If;
        End If;
      
        n_Count := 1;
      End If;
    End Loop;
  
    For r_���㷽ʽ In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As ���㿨���,
                          Extractvalue(b.Column_Value, '/JS/JSKH') As ���㿨��,
                          Extractvalue(b.Column_Value, '/JS/JSFS') As ���㷽ʽ,
                          Extractvalue(b.Column_Value, '/JS/JSJE') As ������,
                          Extractvalue(b.Column_Value, '/JS/JYLSH') As ������ˮ��,
                          Extractvalue(b.Column_Value, '/JS/JYSM') As ����˵��, Extractvalue(b.Column_Value, '/JS/ZY') As ժҪ,
                          Extractvalue(b.Column_Value, '/JS/SFCYJ') As �Ƿ��Ԥ��,
                          Extractvalue(b.Column_Value, '/JS/SFXFK') As �Ƿ����ѿ�
                   From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
      If Nvl(r_���㷽ʽ.�Ƿ��Ԥ��, 0) = 1 Then
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
    End Loop;
  
    --������չ��Ϣ
    If Nvl(n_�����id, 0) <> 0 Then
      If n_Ԥ��id Is Not Null Then
        For c_��չ��Ϣ In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As ��������,
                              Extractvalue(b.Column_Value, '/EXPEND/JYLR') As ��������
                       From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS/EXPENDLIST/EXPEND'))) B) Loop
          Zl_�������㽻��_Insert(n_�����id, 0, v_���㿨��, n_Ԥ��id, c_��չ��Ϣ.�������� || '|' || c_��չ��Ϣ.��������, 1);
        End Loop;
      Else
        For c_��չ��Ϣ In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As ��������,
                              Extractvalue(b.Column_Value, '/EXPEND/JYLR') As ��������
                       From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS/EXPENDLIST/EXPEND'))) B) Loop
          Zl_�������㽻��_Insert(n_�����id, 0, v_���㿨��, n_����id, c_��չ��Ϣ.�������� || '|' || c_��չ��Ϣ.��������, 0);
        End Loop;
      End If;
    End If;
    If Nvl(n_���㿨���, 0) <> 0 Then
      For c_��չ��Ϣ In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As ��������,
                            Extractvalue(b.Column_Value, '/EXPEND/JYLR') As ��������
                     From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS/EXPENDLIST/EXPEND'))) B) Loop
        Zl_�������㽻��_Insert(n_���㿨���, 1, v_���㿨��, n_����id, c_��չ��Ϣ.�������� || '|' || c_��չ��Ϣ.��������, 0);
      End Loop;
    End If;
  End If;
  Update ����Ԥ����¼ Set У�Ա�־ = 0 Where ����id = n_����id And Nvl(У�Ա�־, 0) <> 0;
  v_Temp := '<CZSJ>' || To_Char(d_����ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Settlement;
/

--90143:���ϴ�,2016-01-25,�������ɻ��۵�ʱ���ɷ�����¼
Create Or Replace Procedure Zl_ҽ�ƿ���¼_Insert
(
  --��������������=0-����,1-����,2-����(�൱���ش�)
  --      ����ʱ,���ݺ�_IN�������ԭ��/�����ĵ��ݺš�
  --      ����/������,�ٻ���ʱ�������һ�ο���Ϊ׼��
  ��������_In   Number,
  ���ݺ�_In     סԺ���ü�¼.No%Type,
  ����id_In     סԺ���ü�¼.����id%Type,
  ��ҳid_In     סԺ���ü�¼.��ҳid%Type,
  ��ʶ��_In     סԺ���ü�¼.��ʶ��%Type,
  �ѱ�_In       סԺ���ü�¼.�ѱ�%Type,
  �����id_In   ҽ�ƿ����.Id%Type,
  ԭ����_In     ����ҽ�ƿ���Ϣ.����%Type,
  ҽ�ƿ���_In   ����ҽ�ƿ���Ϣ.����%Type,
  �䶯ԭ��_In   ����ҽ�ƿ��䶯.�䶯ԭ��%Type,
  ����_In       ������Ϣ.����֤��%Type,
  ����_In       סԺ���ü�¼.����%Type,
  �Ա�_In       סԺ���ü�¼.�Ա�%Type,
  ����_In       סԺ���ü�¼.����%Type,
  ���˲���id_In סԺ���ü�¼.���˲���id%Type,
  ���˿���id_In סԺ���ü�¼.���˿���id%Type,
  �շ�ϸĿid_In סԺ���ü�¼.�շ�ϸĿid%Type,
  �շ����_In   סԺ���ü�¼.�շ����%Type,
  ���㵥λ_In   סԺ���ü�¼.���㵥λ%Type,
  ������Ŀid_In סԺ���ü�¼.������Ŀid%Type,
  �վݷ�Ŀ_In   סԺ���ü�¼.�վݷ�Ŀ%Type,
  ��׼����_In   סԺ���ü�¼.��׼����%Type,
  ִ�в���id_In סԺ���ü�¼.ִ�в���id%Type,
  ��������id_In סԺ���ü�¼.��������id%Type,
  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
  �Ӱ��־_In   סԺ���ü�¼.�Ӱ��־%Type,
  ����ʱ��_In   סԺ���ü�¼.�Ǽ�ʱ��%Type,
  ����id_In     Ʊ��ʹ����ϸ.����id%Type,
  Ic����_In     ������Ϣ.Ic����%Type := Null,
  Ӧ�ս��_In   סԺ���ü�¼.Ӧ�ս��%Type,
  ʵ�ս��_In   סԺ���ü�¼.ʵ�ս��%Type,
  ���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type,
  ˢ�����id_In ����Ԥ����¼.�����id%Type,
  ���ѿ�_In     Integer := 0,
  ˢ������_In   ����ҽ�ƿ���Ϣ.����%Type,
  ����id_In     ����Ԥ����¼.����id%Type,
  ������ˮ��_In ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In   ����Ԥ����¼.����˵��%Type := Null,
  ������λ_In   ����Ԥ����¼.������λ%Type := Null,
  ���½������_In  Number := 0,--�Ƿ������Ա��������Ҫ�Ǵ���ͳһ����Ա��¼��̨�������������
  ժҪ_In       סԺ���ü�¼.ժҪ%Type := Null
) As

  Cursor c_Precard Is
    Select ID As ����id From סԺ���ü�¼ Where ��¼���� = 5 And ʵ��Ʊ�� = ԭ����_In And ����id = ����id_In;
  r_Cardrow c_Precard%RowType;

  Cursor c_ҽ�ƿ� Is
    Select ID, ����, ����, ����, ǰ׺�ı�, ���ų���, ȱʡ��־, �Ƿ�̶�, �Ƿ��ϸ����, Nvl(�Ƿ�ˢ��, 0) As �Ƿ�ˢ��, Nvl(�Ƿ�����, 0) As �Ƿ�����,
           Nvl(�Ƿ�����ʻ�, 0) As �Ƿ�����ʻ�, Nvl(�Ƿ�ȫ��, 0) As �Ƿ�ȫ��, ����, ��ע, �ض���Ŀ, ���㷽ʽ, �Ƿ�����, ��������, Nvl(�Ƿ��ظ�ʹ��, 0) As �Ƿ��ظ�ʹ��
    From ҽ�ƿ����
    Where ID = �����id_In;
  r_ҽ�ƿ� c_ҽ�ƿ�%RowType;

  v_����id         סԺ���ü�¼.Id%Type;
  v_����id         סԺ���ü�¼.����id%Type;
  v_�ջ�id         Ʊ�ݴ�ӡ����.Id%Type;
  v_��ӡid         Ʊ�ݴ�ӡ����.Id%Type;
  n_���մ���       Ʊ��ʹ����ϸ.���մ���%Type;
  n_����           Ʊ��ʹ����ϸ.����%Type;
  n_����ֵ         �������.�������%Type;
  n_Count          Number(18);
  n_Ԥ��id         ����Ԥ����¼.Id%Type;
  n_���ѿ�id       ���ѿ�Ŀ¼.Id%Type;
  n_���ƿ�         Number;
  n_ҽ�ƿ��ظ�ʹ�� Number(3);
  Err_Item Exception;
  v_Err_Msg  Varchar2(500);
  n_��id     ����ɿ����.Id%Type;
  n_�䶯���� Number;
Begin
  n_��id := Zl_Get��id(����Ա����_In);
  Open c_ҽ�ƿ�;
  Fetch c_ҽ�ƿ�
    Into r_ҽ�ƿ�;
  If c_ҽ�ƿ�%RowCount = 0 Then
    Close c_ҽ�ƿ�;
    v_Err_Msg := '[ZLSOFT]û�з���ԭҽ�ƿ�����Ӧ���,���ܼ���������[ZLSOFT]';
    Raise Err_Item;
  End If;

  n_ҽ�ƿ��ظ�ʹ�� := Nvl(r_ҽ�ƿ�.�Ƿ��ظ�ʹ��, 0);
  Close c_ҽ�ƿ�;
  If Not ���㷽ʽ_In Is Null Then
    If Nvl(����id_In, 0) <> 0 Then
      v_����id := ����id_In;
    Else
      Select ���˽��ʼ�¼_Id.Nextval Into v_����id From Dual;
    End If;
  End If;
  If ��������_In <> 2 Then
    --�����Ͳ���
    Select ���˷��ü�¼_Id.Nextval Into v_����id From Dual;
  
    Insert Into סԺ���ü�¼
      (ID, ��¼����, ��¼״̬, NO, ʵ��Ʊ��, ���, ����id, ��ҳid, ���˲���id, ���˿���id, ��ʶ��, ����, �Ա�, ����, �ѱ�, ���ʷ���, �����־, �Ӱ��־, ��������id, ������,
       ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, �շ�ϸĿid, �շ����, ���㵥λ, ����, ����, ��ҩ����, ���ӱ�־, ִ�в���id, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ����id,
       ���ʽ��, �ɿ���id, ����,ժҪ)
    Values
      (v_����id, 5, 1, ���ݺ�_In, ҽ�ƿ���_In, 1, ����id_In, Decode(��ҳid_In, 0, Null, ��ҳid_In),
       Decode(���˲���id_In, 0, Null, ���˲���id_In), Decode(���˿���id_In, 0, Null, ���˿���id_In), Decode(��ʶ��_In, 0, Null, ��ʶ��_In),
       ����_In, �Ա�_In, ����_In, �ѱ�_In, Decode(���㷽ʽ_In, Null, 1, 0), 3, �Ӱ��־_In, ��������id_In, ����Ա����_In, ����Ա���_In, ����Ա����_In,
       ����ʱ��_In, ����ʱ��_In, �շ�ϸĿid_In, �շ����_In, ���㵥λ_In, 1, 1, ҽ�ƿ���_In, ��������_In, ִ�в���id_In, ������Ŀid_In, �վݷ�Ŀ_In, ��׼����_In,
       Ӧ�ս��_In, ʵ�ս��_In, v_����id, Decode(���㷽ʽ_In, Null, Null, ʵ�ս��_In), n_��id, �����id_In, ժҪ_In);
  
    --���������ҽ�ƿ����ã��򽫽������벡��Ԥ����¼
    If Not ���㷽ʽ_In Is Null Then
      Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      Insert Into ����Ԥ����¼
        (ID, NO, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, ժҪ, �ɿ���id, �����id, ����, ���㿨���, ������ˮ��,
         ����˵��, �������, ������λ, ��������)
      Values
        (n_Ԥ��id, ���ݺ�_In, 5, 1, ����id_In, Decode(��ҳid_In, 0, Null, ��ҳid_In), Decode(���˿���id_In, 0, Null, ���˿���id_In),
         ���㷽ʽ_In, ����ʱ��_In, ����Ա���_In, ����Ա����_In, ʵ�ս��_In, v_����id, 'ҽ�ƿ�����', n_��id, Decode(���ѿ�_In, 0, ˢ�����id_In, Null),
         ˢ������_In, Decode(���ѿ�_In, 0, Null, ˢ�����id_In), ������ˮ��_In, ����˵��_In, v_����id, ������λ_In, 5);
    
      If ���ѿ�_In = 1 And ˢ������_In Is Not Null Then
      
        n_���ѿ�id := Null;
        Begin
          Select Nvl(���ƿ�, 0), 1 Into n_���ƿ�, n_Count From �����ѽӿ�Ŀ¼ Where ��� = ˢ�����id_In;
        Exception
          When Others Then
            n_Count := 0;
        End;
        If n_Count = 0 Then
          v_Err_Msg := '[ZLSOFT]û�з���ԭ���㿨����Ӧ���,���ܼ���������[ZLSOFT]';
          Raise Err_Item;
        End If;
        If n_���ƿ� = 1 Then
          Select ID
          Into n_���ѿ�id
          From ���ѿ�Ŀ¼
          Where �ӿڱ�� = ˢ�����id_In And ���� = ˢ������_In And
                ��� = (Select Max(���) From ���ѿ�Ŀ¼ Where �ӿڱ�� = ˢ�����id_In And ���� = ˢ������_In);
        End If;
        Zl_���˿������¼_Insert(ˢ�����id_In, n_���ѿ�id, ���㷽ʽ_In, ʵ�ս��_In, ˢ������_In, Null, Null, Null, v_����id, n_Ԥ��id);
      End If;
    End If;
  
    --����ʹ��Ʊ��
    Select Ʊ�ݴ�ӡ����_Id.Nextval Into v_��ӡid From Dual;
    Insert Into Ʊ�ݴ�ӡ���� (ID, ��������, NO) Values (v_��ӡid, 5, ���ݺ�_In);
    n_���մ��� := 0;
    If n_ҽ�ƿ��ظ�ʹ�� = 1 Then
      Select Nvl(Max(���մ���), 0), Nvl(Max(����), 0)
      Into n_���մ���, n_����
      From Ʊ��ʹ����ϸ
      Where Ʊ�� = 5 And ���� = ҽ�ƿ���_In;
      If n_���մ��� > 0 Or n_���� > 0 Then
        n_���մ��� := n_���մ��� + 1;
      End If;
    Else
      --��Ҫ����Ƿ����Ʊ��ʹ����ϸ��������ڣ��϶��ᷢ������
      Select Nvl(Max(����), 0)
      Into n_����
      From Ʊ��ʹ����ϸ A, Ʊ�����ü�¼ B
      Where a.Ʊ�� = 5 And a.���� = ҽ�ƿ���_In And Nvl(a.����id, 0) = Nvl(����id_In, 0) And a.����id = b.Id;
      If n_���� <> 0 Then
        v_Err_Msg := '[ZLSOFT]����:' || ҽ�ƿ���_In || ' �Ѿ�ʹ�ã������ٽ��з�������,����![ZLSOFT]';
        Raise Err_Item;
      End If;
    End If;
  
    Insert Into Ʊ��ʹ����ϸ
      (ID, Ʊ��, ����, ����, ԭ��, ����id, ���մ���, ��ӡid, ʹ��ʱ��, ʹ����)
    Values
      (Ʊ��ʹ����ϸ_Id.Nextval, 5, ҽ�ƿ���_In, 1, 1, ����id_In, Decode(n_���մ���, 0, Null, n_���մ���), v_��ӡid, ����ʱ��_In, ����Ա����_In);
    --����ǻ���,�ٷ���,�򲻼�ʣ������
    If Nvl(n_���մ���, 0) = 0 Then
      --��������״̬�仯
      Update Ʊ�����ü�¼
      Set ��ǰ���� = ҽ�ƿ���_In, ʣ������ = Decode(Sign(ʣ������ - 1), -1, 0, ʣ������ - 1), ʹ��ʱ�� = Sysdate
      Where ID = Nvl(����id_In, 0);
    End If;
  
    --��ػ��ܱ�Ĵ���
    If ���㷽ʽ_In Is Null Then
      --����'�������'
      Update �������
      Set ������� = Nvl(�������, 0) + ʵ�ս��_In
      Where ���� = 1 And ����id = ����id_In And Nvl(����, 2) = Decode(Nvl(��ҳid_In, 0), 0, 1, 2)
      Returning ������� Into n_����ֵ;
    
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, ����, ����, Ԥ�����, �������)
        Values
          (����id_In, 1, Decode(Nvl(��ҳid_In, 0), 0, 1, 2), 0, ʵ�ս��_In);
        n_����ֵ := ʵ�ս��_In;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From ������� Where ����id = ����id_In And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
      End If;
    
      --����'����δ�����'
      Update ����δ�����
      Set ��� = Nvl(���, 0) + ʵ�ս��_In
      Where ����id = ����id_In And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0) And Nvl(���˲���id, 0) = Nvl(���˲���id_In, 0) And
            Nvl(���˿���id, 0) = Nvl(���˿���id_In, 0) And Nvl(��������id, 0) = Nvl(��������id_In, 0) And
            Nvl(ִ�в���id, 0) = Nvl(ִ�в���id_In, 0) And ������Ŀid + 0 = ������Ŀid_In And ��Դ;�� = 3;
    
      If Sql%RowCount = 0 Then
        Insert Into ����δ�����
          (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
        Values
          (����id_In, Decode(��ҳid_In, 0, Null, ��ҳid_In), Decode(���˲���id_In, 0, Null, ���˲���id_In),
           Decode(���˿���id_In, 0, Null, ���˿���id_In), ��������id_In, ִ�в���id_In, ������Ŀid_In, 3, ʵ�ս��_In);
      End If;
    
    Else
      --����"��Ա�ɿ����"
      if Nvl(���½������_In,0)=0 then
        Update ��Ա�ɿ����
        Set ��� = Nvl(���, 0) + ʵ�ս��_In
        Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = ���㷽ʽ_In
        Returning ��� Into n_����ֵ;
      
        If Sql%RowCount = 0 Then
          Insert Into ��Ա�ɿ���� (�տ�Ա, ���㷽ʽ, ����, ���) Values (����Ա����_In, ���㷽ʽ_In, 1, ʵ�ս��_In);
          n_����ֵ := ʵ�ս��_In;
        End If;
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete From ��Ա�ɿ����
          Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = ���㷽ʽ_In And Nvl(���, 0) = 0;
        End If;
      End if;
    End If;
  
  Else
    --��������ʽ
    --���Ȳ�����Ҫ������ԭҽ�ƿ����ü�¼
    Open c_Precard;
    Fetch c_Precard
      Into r_Cardrow;
  
    If c_Precard%RowCount = 0 Then
      Close c_Precard;
      v_Err_Msg := '[ZLSOFT]û�з���ԭҽ�ƿ����ż�¼,��������ʧ�ܣ�[ZLSOFT]';
      Raise Err_Item;
    Else
      --������ԭ���ü�¼ʱ�Ŵ���
      --�ش��ջ�Ʊ��
      Begin
        Select ID
        Into v_�ջ�id
        From (Select b.Id
               From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
               Where a.��ӡid = b.Id And a.���� = 1 And b.�������� = 5 And b.No = ���ݺ�_In
               Order By a.ʹ��ʱ�� Desc)
        Where Rownum < 2;
      Exception
        When Others Then
          Null;
      End;
    
      If v_�ջ�id Is Not Null Then
        Insert Into Ʊ��ʹ����ϸ
          (ID, Ʊ��, ����, ����, ԭ��, ����id, ���մ���, ��ӡid, ʹ��ʱ��, ʹ����)
          Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 4, ����id, ���մ���, ��ӡid, ����ʱ��_In, ����Ա����_In
          From Ʊ��ʹ����ϸ
          Where ��ӡid = v_�ջ�id And Ʊ�� = 5 And ���� = 1;
      End If;
    
      --�ش򷢳�Ʊ��
      Select Ʊ�ݴ�ӡ����_Id.Nextval Into v_��ӡid From Dual;
    
      Insert Into Ʊ�ݴ�ӡ���� (ID, ��������, NO) Values (v_��ӡid, 5, ���ݺ�_In);
      n_���մ��� := 0;
      If n_ҽ�ƿ��ظ�ʹ�� = 1 Then
        Select Nvl(Max(���մ���), 0), Nvl(Max(����), 0)
        Into n_���մ���, n_����
        From Ʊ��ʹ����ϸ
        Where Ʊ�� = 5 And ���� = ҽ�ƿ���_In;
        If n_���մ��� > 0 Or n_���� > 0 Then
          n_���մ��� := n_���մ��� + 1;
        End If;
      Else
        --��Ҫ����Ƿ����Ʊ��ʹ����ϸ��������ڣ��϶��ᷢ������
        Select Nvl(Max(����), 0)
        Into n_����
        From Ʊ��ʹ����ϸ A, Ʊ�����ü�¼ B
        Where a.Ʊ�� = 5 And a.���� = ҽ�ƿ���_In And Nvl(a.����id, 0) = Nvl(����id_In, 0) And a.����id = b.Id;
        If n_���� <> 0 Then
          v_Err_Msg := '[ZLSOFT]�¿���:' || ҽ�ƿ���_In || ' �Ѿ�ʹ�ã��뻻һ���¿�,����![ZLSOFT]';
          Raise Err_Item;
        End If;
      
      End If;
    
      Insert Into Ʊ��ʹ����ϸ
        (ID, Ʊ��, ����, ����, ԭ��, ����id, ���մ���, ��ӡid, ʹ��ʱ��, ʹ����)
      Values
        (Ʊ��ʹ����ϸ_Id.Nextval, 5, ҽ�ƿ���_In, 1, Decode(v_�ջ�id, Null, 1, 3), ����id_In, Decode(n_���մ���, 0, Null, n_���մ���), v_��ӡid,
         ����ʱ��_In, ����Ա����_In);
      --����ǻ���,�ٷ���,�򲻼�ʣ������
      If Nvl(n_���մ���, 0) = 0 Then
        --����״̬�仯
        Update Ʊ�����ü�¼
        Set ��ǰ���� = ҽ�ƿ���_In, ʣ������ = Decode(Sign(ʣ������ - 1), -1, 0, ʣ������ - 1), ʹ��ʱ�� = Sysdate
        Where ID = Nvl(����id_In, 0);
      End If;
      --����ԭ������¼״̬
      Update סԺ���ü�¼
      Set ʵ��Ʊ�� = ҽ�ƿ���_In, ��ҩ���� = ҽ�ƿ���_In, ���ӱ�־ = 2, ���� = �����id_In
      Where ID = r_Cardrow.����id;
      Close c_Precard;
    End If;
  End If;

  --������صı䶯��Ϣ
  --Zl_ҽ�ƿ��䶯_Insert (�䶯����_In/����id_In ,�����id_In, ԭ����_In, ҽ�ƿ���_In, �䶯ԭ��_In, ����_In, ����Ա����_In, �䶯ʱ��_In
  --Ic����_In, ��ʧ��ʽ_In)
  --�䶯����_In:1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����); ��-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ)
  n_�䶯���� := Case
              When ��������_In = 0 Then
               1
              When ��������_In = 1 Then
               3
              Else
               2
            End;
  Zl_ҽ�ƿ��䶯_Insert(n_�䶯����, ����id_In, �����id_In, ԭ����_In, ҽ�ƿ���_In, �䶯ԭ��_In, ����_In, ����Ա����_In, ����ʱ��_In, Ic����_In, Null);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҽ�ƿ���¼_Insert;
/

Create Or Replace Procedure Zl_ҽ�ƿ���¼_Delete
(
  ���ݺ�_In     סԺ���ü�¼.No%Type,
  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
  ����Ա����_In סԺ���ü�¼.����Ա����%Type
) As
  Cursor c_Cardinfo Is
    Select a.Id As ����id, Nvl(a.���ʷ���, 0) As ����, a.����id, a.ʵ��Ʊ��, a.����id, Nvl(a.��ҳid, 0) As ��ҳid,
           Nvl(a.���˲���id, 0) As ���˲���id, Nvl(a.���˿���id, 0) As ���˿���id, Nvl(a.��������id, 0) As ��������id,
           Nvl(a.ִ�в���id, 0) As ִ�в���id, a.������Ŀid, a.ʵ�ս��, b.���㷽ʽ, b.��Ԥ��, b.�����id, b.����, b.���㿨���, b.�������, a.����,
           b.Id As Ԥ��id, a.ժҪ
    From סԺ���ü�¼ A, ����Ԥ����¼ B
    Where a.��¼���� = 5 And a.��¼״̬ = 1 And a.No = ���ݺ�_In And a.����id = b.����id(+);
  r_Cardrow c_Cardinfo%RowType;

  v_����id   סԺ���ü�¼.Id%Type;
  v_����id   סԺ���ü�¼.����id%Type;
  v_��ӡid   Ʊ�ݴ�ӡ����.Id%Type;
  n_����ֵ   �������.�������%Type;
  n_�����id Number(18);
  v_����״̬ ������ü�¼.��¼״̬%Type;

  v_Date Date;

  v_Err_Msg Varchar2(200);
  Err_Item Exception;
  n_��id ����ɿ����.Id%Type;

Begin
  Open c_Cardinfo;
  Fetch c_Cardinfo
    Into r_Cardrow;
  n_��id := Zl_Get��id(����Ա����_In);

  --�����ж�Ҫ�˿��ļ�¼�Ƿ����
  If c_Cardinfo%RowCount = 0 Then
    Close c_Cardinfo;
    v_Err_Msg := '[ZLSOFT]û�з���Ҫ�˿��ļ�¼,�ü�¼�����Ѿ��˳���[ZLSOFT]';
    Raise Err_Item;
  Else
    Select Sysdate Into v_Date From Dual;
    Select ���˷��ü�¼_Id.Nextval Into v_����id From Dual;
  
    If r_Cardrow.���� = 0 Then
      Select ���˽��ʼ�¼_Id.Nextval Into v_����id From Dual;
    End If;
  
    --�˳����￨���ü�¼
    Insert Into סԺ���ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, ����id, ��ҳid, ���˲���id, ���˿���id, �����־, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ����, �Ӱ��־,
       ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ����id, ���ʽ��,
       �ɿ���id, ����, ժҪ)
      Select v_����id, NO, ʵ��Ʊ��, ��¼����, 2, ���, ����id, ��ҳid, ���˲���id, ���˿���id, �����־, ����, �Ա�, ����, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
             -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ����Ա���_In, ����Ա����_In,
             ����ʱ��, v_Date, v_����id, Decode(v_����id, Null, Null, -���ʽ��), n_��id, ����, ժҪ
      From סԺ���ü�¼
      Where ID = r_Cardrow.����id;
  
    Update סԺ���ü�¼ Set ��¼״̬ = 3 Where ID = r_Cardrow.����id;
    
    --���������۵���������ۻ�δ�շѣ�ֱ��ɾ��
    Begin
      Select Nvl(��¼״̬,-1) into v_����״̬ From ������ü�¼ Where ����ID=r_Cardrow.����id And ��¼����=1 And NO=r_Cardrow.ժҪ;
    Exception
      When Others Then
        v_����״̬ := -1;
    End;
    if v_����״̬=0 then
      Zl_���ﻮ�ۼ�¼_Delete(r_Cardrow.ժҪ);
    end if;
    
    --Ԥ���������յĽ�����
    If r_Cardrow.���� = 0 Then
      Insert Into ����Ԥ����¼
        (ID, NO, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����,
         ������ˮ��, ����˵��, ������λ, ��������)
        Select ����Ԥ����¼_Id.Nextval, NO, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, v_Date, ����Ա���_In, ����Ա����_In, -��Ԥ��, v_����id,
               n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 5
        From ����Ԥ����¼
        Where ��¼���� = 5 And ��¼״̬ = 1 And ����id = r_Cardrow.����id;
    
      Update ����Ԥ����¼ Set ��¼״̬ = 3 Where ��¼���� = 5 And ��¼״̬ = 1 And ����id = r_Cardrow.����id;
    
      --Zl_���˿������¼_Strike(����id_In In Varchar2,  Ԥ��id_In ����Ԥ����¼.ID%Type := -1
      Zl_���˿������¼_Strike(r_Cardrow.����id, r_Cardrow.Ԥ��id);
    
    End If;
  
    --�˿��ջ�Ʊ��
    Begin
      Select ID
      Into v_��ӡid
      From (Select b.Id
             From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
             Where a.��ӡid = b.Id And a.���� = 1 And b.�������� = 5 And b.No = ���ݺ�_In
             Order By a.ʹ��ʱ�� Desc)
      Where Rownum < 2;
    Exception
      When Others Then
        Null;
    End;
    If v_��ӡid Is Not Null Then
      Insert Into Ʊ��ʹ����ϸ
        (ID, Ʊ��, ����, ����, ԭ��, ����id, ���մ���, ��ӡid, ʹ��ʱ��, ʹ����)
        Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ���մ���, ��ӡid, v_Date, ����Ա����_In
        From Ʊ��ʹ����ϸ
        Where ��ӡid = v_��ӡid And Ʊ�� = 5 And ���� = 1;
    End If;
  
    n_�����id := To_Number(Nvl(r_Cardrow.����, '0'));
    If n_�����id = 0 Then
      --ȡ���￨
      Select ID Into n_�����id From ҽ�ƿ���� Where ���� = '���￨' And Nvl(�Ƿ�̶�, 0) = 1;
    End If;
  
    --������صı䶯��Ϣ
    --Zl_ҽ�ƿ��䶯_Insert (�䶯����_In/����id_In ,�����id_In, ԭ����_In, ҽ�ƿ���_In, �䶯ԭ��_In, ����_In, ����Ա����_In, �䶯ʱ��_In
    --Ic����_In, ��ʧ��ʽ_In)
    --�䶯����_In:1-����(��11�󶨿�);2-����;3-����(13-����ͣ��);4-�˿�(��14ȡ����); ��-�������(ֻ��¼);6-��ʧ(16ȡ����ʧ)
    Zl_ҽ�ƿ��䶯_Insert(4, r_Cardrow.����id, n_�����id, r_Cardrow.ʵ��Ʊ��, r_Cardrow.ʵ��Ʊ��, Null, Null, ����Ա����_In, v_Date, Null,
                    Null);
  
    ----------------------------------------------------------------------------------------------------------------------------------------
  
    --��ػ��ܱ�Ĵ���
    If r_Cardrow.���� = 1 Then
      --����'�������'
      Update �������
      Set ������� = Nvl(�������, 0) + (-1 * r_Cardrow.ʵ�ս��)
      Where ���� = 1 And ����id = r_Cardrow.����id And Nvl(����, 2) = Decode(Nvl(r_Cardrow.��ҳid, 0), 0, 1, 2)
      Returning ������� Into n_����ֵ;
    
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, ����, ����, Ԥ�����, �������)
        Values
          (r_Cardrow.����id, 1, Decode(Nvl(r_Cardrow.��ҳid, 0), 0, 1, 2), 0, -1 * r_Cardrow.ʵ�ս��);
        n_����ֵ := -1 * r_Cardrow.ʵ�ս��;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete ������� Where ���� = 1 And ����id = r_Cardrow.����id And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
      End If;
    
      --����'����δ�����'
      Update ����δ�����
      Set ��� = Nvl(���, 0) + (-1 * r_Cardrow.ʵ�ս��)
      Where ����id = r_Cardrow.����id And Nvl(��ҳid, 0) = r_Cardrow.��ҳid And Nvl(���˲���id, 0) = r_Cardrow.���˲���id And
            Nvl(���˿���id, 0) = r_Cardrow.���˿���id And Nvl(��������id, 0) = r_Cardrow.��������id And
            Nvl(ִ�в���id, 0) = r_Cardrow.ִ�в���id And ������Ŀid + 0 = r_Cardrow.������Ŀid And ��Դ;�� = 3;
    
      If Sql%RowCount = 0 Then
        Insert Into ����δ�����
          (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
        Values
          (r_Cardrow.����id, Decode(r_Cardrow.��ҳid, 0, Null, r_Cardrow.��ҳid),
           Decode(r_Cardrow.���˲���id, 0, Null, r_Cardrow.���˲���id), Decode(r_Cardrow.���˿���id, 0, Null, r_Cardrow.���˿���id),
           Decode(r_Cardrow.��������id, 0, Null, r_Cardrow.��������id), Decode(r_Cardrow.ִ�в���id, 0, Null, r_Cardrow.ִ�в���id),
           r_Cardrow.������Ŀid, 3, -1 * r_Cardrow.ʵ�ս��);
      End If;
    
    Elsif r_Cardrow.���㷽ʽ Is Not Null Then
      --����"��Ա�ɿ����"
      Update ��Ա�ɿ����
      Set ��� = Nvl(���, 0) + (-1 * r_Cardrow.��Ԥ��)
      Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = r_Cardrow.���㷽ʽ
      Returning ��� Into n_����ֵ;
    
      If Sql%RowCount = 0 Then
        Insert Into ��Ա�ɿ����
          (�տ�Ա, ���㷽ʽ, ����, ���)
        Values
          (����Ա����_In, r_Cardrow.���㷽ʽ, 1, -1 * r_Cardrow.��Ԥ��);
        n_����ֵ := -1 * r_Cardrow.��Ԥ��;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From ��Ա�ɿ����
        Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = r_Cardrow.���㷽ʽ And Nvl(���, 0) = 0;
      End If;
    End If;
  
    Close c_Cardinfo;
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20999, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҽ�ƿ���¼_Delete;
/

--94132:������,2016-03-14,֧����ҽ�������Դ����
--93006:������,2016-01-22,��ֹԤԼ��Դ����
Create Or Replace Procedure Zl_Third_Getnolist
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --����:��ȡ��Դ�б�
  --���:Xml_In:
  --<IN>
  --  <RQ>����</RQ>
  --  <KSID>����ID</KSID>
  --  <YSID>ҽ��ID</YSID>
  --  <YSXM>ҽ������</YSXM>
  --  <HZDW>֧����</HZDW>    //������λ�������˵�ʱ��ֻȡ������λ�ĺ�;Ϊ��ʱ��ֻȡ�Ǻ�����λ�ĺ�
  --</IN>

  --����:Xml_Out
  --<OUTPUT>
  --  <GROUP>
  --    <RQ>����</RQ>
  --    <HBLIST>
  --     <HB>
  --        <HM>235</HM>       //����
  --        <YSID>549</YSID>      //ҽ��ID
  --        <YS>����</YS>       //ҽ������
  --        <KSID>123</KSID>   //����ID
  --        <KSMC>�ڿ�</KSMC>   //��������
  --        <ZC>����ҽʦ</ZC> //ְ��
  --        <XMID>10086<XMID> //�Һ���Ŀ��ID
  --        <XMMC>�Һŷ�</XMMC> //�Һ���Ŀ������
  --        <YGHS>0</YGHS>      //�ѹҺ���
  --        <SYHS>99</SYHS>   //ʣ�����
  --        <PRICE>15</PRICE>      //�۸�
  --        <HL>��ͨ</HL>       //�Һ�����
  --        <HCXH>1</HCXH>    //�Ƿ���ڻ������ʱ��Σ�1-���� 0���߿�-������
  --        <FSD>0</FSD>      //�Ƿ��ʱ��
  --        <FWMC>����</FWMC>     //�ű�ʱ��
  --        <HBTIME>(08:00-17:59)</HBTIME> //�ɹ�ʱ��
  --     <SPANLIST>
  --            <SPAN>
  --                  <SJD/>      //ʱ���
  --                  <SL/>      //����
  --            </SPAN>
  --            ����
  --          </SPANLIST>
  --      </HB>
  --      <HB>
  --      ����
  --      </HB>
  --    </HBLIST>
  --  </GROUP>
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  d_����         Date;
  n_����id       ���˹Һż�¼.ִ�в���id%Type;
  n_ҽ��id       ��Ա��.Id%Type;
  v_ҽ������     ��Ա��.����%Type;
  v_����         �ҺŰ�������.������Ŀ%Type;
  v_ʱ���       Varchar2(100);
  v_������λ     �Һź�����λ.����%Type;
  n_��ʱ��       Number(3);
  n_����ʣ��     Number(5);
  n_�ѹ���       Number(5);
  n_��Լ�ѹ���   Number(5);
  n_�ϼƽ��     �շѼ�Ŀ.�ּ�%Type;
  n_��Լ������   Number(5);
  n_��Լʣ������ Number(5);
  n_���������� Number(5);
  n_��Լģʽ     Number(3); --��Լģʽ:1-��Լ��λ������ģʽ 0-��Լ��λָ�����ģʽ
  n_�Ǻ�Լ       Number(3);
  n_�Ƿ�Ԥ��     Number(3);
  d_�Ӻ�ʱ��     Date;
  n_�������     Number(3);
  n_ʱ������     Number(5);
  n_Ԥ������     Number(5);
  n_����ԤԼ     Number(3);
  n_����         Number(3);
  v_ʣ������     Varchar2(100);
  v_Timetemp     Varchar2(100);
  v_Temp         Varchar2(32767); --��ʱXML
  v_Xmlmain      Clob; --��ʱXML
  c_Xmlmain      Clob; --��ʱXML
  x_Templet      Xmltype; --ģ��XML
  v_Err_Msg      Varchar2(200);
  n_Exists       Number(2);
  v_Sql          Varchar2(20000);
  Type c_Main Is Ref Cursor;
  r_����id   �ҺŰ���.����id%Type;
  r_����     �ҺŰ���.����%Type;
  r_�������� ���ű�.����%Type;
  r_ҽ������ �ҺŰ���.ҽ������%Type;
  r_ҽ��id   �ҺŰ���.ҽ��id%Type;
  r_ְ��     ��Ա��.רҵ����ְ��%Type;
  r_����     �ҺŰ���.����%Type;
  r_����id   �ҺŰ���.Id%Type;
  r_�ƻ�id   �ҺŰ��żƻ�.Id%Type;
  r_�Ű�     �ҺŰ���.����%Type;
  r_��Ŀid   �ҺŰ���.��Ŀid%Type;
  r_��Ŀ���� �շ���ĿĿ¼.����%Type;
  r_��ſ��� �ҺŰ���.��ſ���%Type;
  r_�޺���   �ҺŰ�������.�޺���%Type;
  r_��Լ��   �ҺŰ�������.��Լ��%Type;
  r_�ѹ���   ���˹ҺŻ���.�ѹ���%Type;
  r_��Լ��   ���˹ҺŻ���.��Լ��%Type;
  r_�ѽ���   ���˹ҺŻ���.�����ѽ���%Type;
  r_�۸�     �շѼ�Ŀ.�ּ�%Type;
  r_No       c_Main;
  n_Curcount Number(3);

  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select To_Date(Extractvalue(Value(A), 'IN/RQ'), 'YYYY-MM-DD hh24:mi:ss'), Extractvalue(Value(A), 'IN/KSID'),
         Extractvalue(Value(A), 'IN/YSID'), Extractvalue(Value(A), 'IN/YSXM'), Extractvalue(Value(A), 'IN/HZDW')
  Into d_����, n_����id, n_ҽ��id, v_ҽ������, v_������λ
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;
  --���ڽڵ�Ϊ�յ����
  If d_���� Is Null Then
    d_���� := Trunc(Sysdate);
  End If;

  Select Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����', Null)
  Into v_����
  From Dual;
  n_��Լʣ������ := 0;

  v_Sql := 'Select a.*, Nvl(Hz.�ѹ���, 0) As �ѹ���, Nvl(Hz.��Լ��, 0) As ��Լ��, Nvl(Hz.�����ѽ���, 0) As �ѽ���, b.�ּ� As �۸� ';
  v_Sql := v_Sql ||
           'From (Select Ap.����id, Ap.����, Bm.���� As ��������, Ap.ҽ������, Nvl(Ap.ҽ��id, 0) As ҽ��id, Ry.רҵ����ְ�� As ְ��, Ap.����, ';
  v_Sql := v_Sql || ' Ap.����id, Ap.�ƻ�id, Ap.�Ű�, Ap.��Ŀid, Fy.���� As ��Ŀ����, Ap.��ſ���, Ap.�޺���, Ap.��Լ�� ';
  v_Sql := v_Sql || 'From (Select Ap.����id, Ap.����, Bm.���� As ��������, Ap.ҽ������, Ap.ҽ��id, Ap.����, Ap.Id As ����id, 0 As �ƻ�id, ';
  v_Sql := v_Sql || 'Ap.��Ŀid, Nvl(Ap.��ſ���, 0) As ��ſ���, ';
  v_Sql := v_Sql || 'Decode(To_Char(:1, ''D''), ''1'', Ap.����, ''2'', Ap.��һ, ''3'', Ap.�ܶ�, ''4'', Ap.����, ''5'', Ap.����, ';
  v_Sql := v_Sql || ' ''6'', Ap.����, ''7'', Ap.����, Null) As �Ű�, Xz.��Լ��, Xz.�޺��� ';
  v_Sql := v_Sql || 'From �ҺŰ��� Ap, ���ű� Bm, �ҺŰ������� Xz ';
  v_Sql := v_Sql || 'Where Ap.����id = Bm.Id(+) ';

  n_Curcount := 2;
  If Nvl(n_����id, 0) <> 0 Then
    v_Sql      := v_Sql || 'And Ap.����id = :2 ';
    n_Curcount := n_Curcount + 1;
  End If;
  If Nvl(n_ҽ��id, 0) <> 0 Then
    If n_Curcount = 2 Then
      v_Sql := v_Sql || 'And Ap.ҽ��id = :2 ';
    Else
      v_Sql := v_Sql || 'And Ap.ҽ��id = :3 ';
    End If;
    n_Curcount := n_Curcount + 1;
  End If;
  If Nvl(v_ҽ������, '_') <> '_' Then
    If n_Curcount = 2 Then
      v_Sql := v_Sql || 'And Ap.ҽ������ = :2 ';
    End If;
    If n_Curcount = 3 Then
      v_Sql := v_Sql || 'And Ap.ҽ������ = :3 ';
    End If;
    If n_Curcount = 4 Then
      v_Sql := v_Sql || 'And Ap.ҽ������ = :4 ';
    End If;
    n_Curcount := n_Curcount + 1;
  End If;

  v_Sql      := v_Sql || 'And Ap.ͣ������ Is Null And :' || n_Curcount ||
                ' Between Nvl(Ap.��ʼʱ��, To_Date(''1900-01-01'', ''YYYY-MM-DD'')) And ';
  n_Curcount := n_Curcount + 1;
  v_Sql      := v_Sql || 'Nvl(Ap.��ֹʱ��, To_Date(''3000 - 01 - 01'', ''YYYY-MM-DD'')) And Xz.����id(+) = Ap.Id And ';
  v_Sql      := v_Sql || ' Xz.������Ŀ(+) = Decode(To_Char(:' || n_Curcount ||
                ', ''D''), ''1'', ''����'', ''2'', ''��һ'', ''3'', ''�ܶ�'', ''4'', ''����'', ''5'', ';
  n_Curcount := n_Curcount + 1;
  v_Sql      := v_Sql || ' ''����'', ''6'', ''����'', ''7'', ''����'', Null) And Not Exists ';
  v_Sql      := v_Sql || '(Select Rownum ';
  v_Sql      := v_Sql || 'From �ҺŰ���ͣ��״̬ Ty ';
  v_Sql      := v_Sql || 'Where Ty.����id = Ap.Id And :' || n_Curcount ||
                ' Between Ty.��ʼֹͣʱ�� And Ty.����ֹͣʱ��) And Not Exists ';
  n_Curcount := n_Curcount + 1;
  v_Sql      := v_Sql || '(Select Rownum ';
  v_Sql      := v_Sql || 'From �ҺŰ��żƻ� Jh Where Jh.����id = Ap.Id And Jh.���ʱ�� Is Not Null And ';
  v_Sql      := v_Sql || ':' || n_Curcount ||
                ' Between Nvl(Jh.��Чʱ��, To_Date(''1900-01-01'', ''YYYY-MM-DD'')) And Nvl(Jh.ʧЧʱ��, To_Date(''3000-01-01'', ''YYYY-MM-DD''))) ';
  n_Curcount := n_Curcount + 1;
  v_Sql      := v_Sql || 'Union All ';
  v_Sql      := v_Sql ||
                'Select Ap.����id, Ap.����, Bm.���� As ��������, Jh.ҽ������, Jh.ҽ��id, Ap.����, Ap.Id As ����id, Jh.Id As �ƻ�id, ';
  v_Sql      := v_Sql || 'Jh.��Ŀid, Nvl(Jh.��ſ���, 0) As ��ſ���,Decode(To_Char(:' || n_Curcount ||
                ', ''D''), ''1'', Jh.����, ''2'', Jh.��һ, ''3'', Jh.�ܶ�, ''4'', Jh.����, ''5'', Jh.����, ';
  n_Curcount := n_Curcount + 1;
  v_Sql      := v_Sql || ' ''6'', Jh.����, ''7'', Jh.����, Null) As �Ű�, Xz.��Լ��, Xz.�޺��� ';
  v_Sql      := v_Sql || 'From �ҺŰ��żƻ� Jh, �ҺŰ��� Ap, ���ű� Bm, �Һżƻ����� Xz ';
  v_Sql      := v_Sql || 'Where Jh.����id = Ap.Id And Ap.����id = Bm.Id(+) And Ap.ͣ������ Is Null ';

  If Nvl(n_����id, 0) <> 0 Then
    v_Sql      := v_Sql || 'And Ap.����id = :' || n_Curcount || ' ';
    n_Curcount := n_Curcount + 1;
  End If;
  If Nvl(n_ҽ��id, 0) <> 0 Then
    v_Sql      := v_Sql || 'And Jh.ҽ��id = :' || n_Curcount || ' ';
    n_Curcount := n_Curcount + 1;
  End If;
  If Nvl(v_ҽ������, '_') <> '_' Then
    v_Sql      := v_Sql || 'And Jh.ҽ������ = :' || n_Curcount || ' ';
    n_Curcount := n_Curcount + 1;
  End If;

  v_Sql      := v_Sql || ' And :' || n_Curcount ||
                ' Between Nvl(Jh.��Чʱ��, To_Date(''1900-01-01'', ''YYYY-MM-DD'')) And Nvl(Jh.ʧЧʱ��, To_Date(''3000-01-01'', ''YYYY-MM-DD'')) And Xz.�ƻ�id(+) = Jh.Id And ';
  n_Curcount := n_Curcount + 1;
  v_Sql      := v_Sql || 'Xz.������Ŀ(+) = Decode(To_Char(:' || n_Curcount ||
                ', ''D''), ''1'', ''����'', ''2'', ''��һ'', ''3'', ''�ܶ�'', ''4'', ''����'', ''5'', ''����'', ''6'', ''����'', ''7'', ''����'', Null) And Not Exists ';
  n_Curcount := n_Curcount + 1;
  v_Sql      := v_Sql || '(Select Rownum From �ҺŰ���ͣ��״̬ Ty Where Ty.����id = Ap.Id And :' || n_Curcount ||
                ' Between Ty.��ʼֹͣʱ�� And Ty.����ֹͣʱ��) And ';
  n_Curcount := n_Curcount + 1;
  v_Sql      := v_Sql || '(Jh.��Чʱ��, Jh.����id) = (Select Max(Sxjh.��Чʱ��) As ��Чʱ��, ����id From �ҺŰ��żƻ� Sxjh ';
  v_Sql      := v_Sql || ' Where Sxjh.���ʱ�� Is Not Null And :' || n_Curcount ||
                ' Between Nvl(Sxjh.��Чʱ��, To_Date(''1900-01-01'', ''YYYY-MM-DD'')) And ';
  n_Curcount := n_Curcount + 1;
  v_Sql      := v_Sql || 'Nvl(Sxjh.ʧЧʱ��, To_Date(''3000-01-01'', ''YYYY-MM-DD'')) And Sxjh.����id = Jh.����id ';
  v_Sql      := v_Sql || 'Group By Sxjh.����id)) Ap, ���ű� Bm, ��Ա�� Ry, �շ���ĿĿ¼ Fy ';
  v_Sql      := v_Sql || 'Where Ap.����id = Bm.Id(+) And Ap.ҽ��id = Ry.Id(+) And Ap.��Ŀid = Fy.Id And �Ű� Is Not Null) A, ';
  v_Sql      := v_Sql || '���˹ҺŻ��� Hz, �շѼ�Ŀ B ';
  v_Sql      := v_Sql || 'Where a.���� = Hz.����(+) And Hz.����(+) = Trunc(:' || n_Curcount || ') And a.��Ŀid = b.�շ�ϸĿid And ';
  n_Curcount := n_Curcount + 1;
  v_Sql      := v_Sql || 'Nvl(b.��ֹ����, To_Date(''3000-1-1'', ''YYYY-Mm-DD'')) > :' || n_Curcount || ' ';
  n_Curcount := n_Curcount + 1;
  v_Sql      := v_Sql || 'And b.ִ������ <= :' || n_Curcount || ' ';
  If Nvl(n_����id, 0) <> 0 And Nvl(n_ҽ��id, 0) <> 0 And Nvl(v_ҽ������, '_') <> '_' Then
    Open r_No For v_Sql
      Using d_����, n_����id, n_ҽ��id, v_ҽ������, d_����, d_����, d_����, d_����, d_����, n_����id, n_ҽ��id, v_ҽ������, d_����, d_����, d_����, d_����, d_����, d_����, d_����;
  End If;
  If Nvl(n_����id, 0) <> 0 And Nvl(n_ҽ��id, 0) = 0 And Nvl(v_ҽ������, '_') = '_' Then
    Open r_No For v_Sql
      Using d_����, n_����id, d_����, d_����, d_����, d_����, d_����, n_����id, d_����, d_����, d_����, d_����, d_����, d_����, d_����;
  End If;
  If Nvl(n_����id, 0) = 0 And Nvl(n_ҽ��id, 0) <> 0 And Nvl(v_ҽ������, '_') = '_' Then
    Open r_No For v_Sql
      Using d_����, n_ҽ��id, d_����, d_����, d_����, d_����, d_����, n_ҽ��id, d_����, d_����, d_����, d_����, d_����, d_����, d_����;
  End If;
  If Nvl(n_����id, 0) = 0 And Nvl(n_ҽ��id, 0) = 0 And Nvl(v_ҽ������, '_') <> '_' Then
    Open r_No For v_Sql
      Using d_����, v_ҽ������, d_����, d_����, d_����, d_����, d_����, v_ҽ������, d_����, d_����, d_����, d_����, d_����, d_����, d_����;
  End If;
  If Nvl(n_����id, 0) <> 0 And Nvl(n_ҽ��id, 0) <> 0 And Nvl(v_ҽ������, '_') = '_' Then
    Open r_No For v_Sql
      Using d_����, n_����id, n_ҽ��id, d_����, d_����, d_����, d_����, d_����, n_����id, n_ҽ��id, d_����, d_����, d_����, d_����, d_����, d_����, d_����;
  End If;
  If Nvl(n_����id, 0) <> 0 And Nvl(n_ҽ��id, 0) = 0 And Nvl(v_ҽ������, '_') <> '_' Then
    Open r_No For v_Sql
      Using d_����, n_����id, v_ҽ������, d_����, d_����, d_����, d_����, d_����, n_����id, v_ҽ������, d_����, d_����, d_����, d_����, d_����, d_����, d_����;
  End If;
  If Nvl(n_����id, 0) = 0 And Nvl(n_ҽ��id, 0) <> 0 And Nvl(v_ҽ������, '_') <> '_' Then
    Open r_No For v_Sql
      Using d_����, n_ҽ��id, v_ҽ������, d_����, d_����, d_����, d_����, d_����, n_ҽ��id, v_ҽ������, d_����, d_����, d_����, d_����, d_����, d_����, d_����;
  End If;
  Loop
    Fetch r_No
      Into r_����id, r_����, r_��������, r_ҽ������, r_ҽ��id, r_ְ��, r_����, r_����id, r_�ƻ�id, r_�Ű�, r_��Ŀid, r_��Ŀ����, r_��ſ���, r_�޺���, r_��Լ��,
           r_�ѹ���, r_��Լ��, r_�ѽ���, r_�۸�;
    Exit When r_No%NotFound;
    If r_�ƻ�id <> 0 Then
      Select Sign(Count(Rownum))
      Into n_��ʱ��
      From �ҺŰ��żƻ� Jh, �Һżƻ�ʱ�� Sd
      Where Jh.Id = Sd.�ƻ�id And Jh.Id = r_�ƻ�id And
            Sd.���� =
            Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����', Null) And
            Rownum < 2;
    Else
      Select Sign(Count(Rownum))
      Into n_��ʱ��
      From �ҺŰ��� Ap, �ҺŰ���ʱ�� Sd
      Where Ap.Id = Sd.����id And Ap.Id = r_����id And
            Sd.���� =
            Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����', Null) And
            Rownum < 2;
    End If;
    If n_��ʱ�� = 0 Then
      v_Temp := '';
      If v_������λ Is Not Null And r_��ſ��� = 1 Then
        If r_�ƻ�id <> 0 Then
          Select Nvl(Sum(����), 0)
          Into n_��Լ������
          From ������λ�ƻ�����
          Where �ƻ�id = r_�ƻ�id And ������λ = v_������λ And
                ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7',
                              '����', Null);
          Select Count(1)
          Into n_��Լģʽ
          From ������λ�ƻ�����
          Where �ƻ�id = r_�ƻ�id And ������λ = v_������λ And
                ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7',
                              '����', Null) And ��� = 0;
        Else
          Select Nvl(Sum(����), 0)
          Into n_��Լ������
          From ������λ���ſ���
          Where ����id = r_����id And ������λ = v_������λ And
                ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7',
                              '����', Null);
          Select Count(1)
          Into n_��Լģʽ
          From ������λ���ſ���
          Where ����id = r_����id And ������λ = v_������λ And
                ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7',
                              '����', Null) And ��� = 0;
        End If;
        If n_��Լģʽ = 0 Then
          If r_�ƻ�id <> 0 Then
            Select Count(1)
            Into n_��Լ�ѹ���
            From ���˹Һż�¼ A
            Where �ű� = r_���� And ��¼״̬ = 1 And ����ʱ�� Between Trunc(d_����) And Trunc(d_���� + 1) - 1 / 60 / 60 / 24 And Exists
             (Select 1
                   From ������λ�ƻ�����
                   Where �ƻ�id = r_�ƻ�id And ������λ = v_������λ And
                         ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6',
                                       '����', '7', '����', Null) And ��� = a.���� And ���� <> 0);
          Else
            Select Count(1)
            Into n_��Լ�ѹ���
            From ���˹Һż�¼ A
            Where �ű� = r_���� And ��¼״̬ = 1 And ����ʱ�� Between Trunc(d_����) And Trunc(d_���� + 1) - 1 / 60 / 60 / 24 And Exists
             (Select 1
                   From ������λ���ſ���
                   Where ����id = r_����id And ������λ = v_������λ And
                         ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6',
                                       '����', '7', '����', Null) And ��� = a.���� And ���� <> 0);
          End If;
        Else
          Begin
            Select Count(1)
            Into n_��Լ�ѹ���
            From ���˹Һż�¼
            Where �ű� = r_���� And ��¼״̬ = 1 And ������λ = v_������λ And ����ʱ�� Between Trunc(d_����) And
                  Trunc(d_���� + 1) - 1 / 60 / 60 / 24;
          Exception
            When Others Then
              n_��Լ�ѹ��� := 0;
          End;
        End If;
        If n_��Լ������ = 0 Then
          n_��Լʣ������ := 0;
        Else
          n_��Լʣ������ := n_��Լ������ - n_��Լ�ѹ���;
          If n_��Լʣ������ > (Nvl(r_�޺���, 0) - r_�ѹ���) Then
            n_��Լʣ������ := Nvl(r_�޺���, 0) - r_�ѹ���;
          End If;
        End If;
      End If;
    Else
      v_Temp := '<SPANLIST>';
      If r_�ƻ�id <> 0 Then
        Select Max(����ʱ��)
        Into d_�Ӻ�ʱ��
        From �Һżƻ�ʱ��
        Where �ƻ�id = r_�ƻ�id And ���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����',
                                            '6', '����', '7', '����', Null);
        If r_��ſ��� = 1 Then
          If Trunc(d_����) = Trunc(Sysdate) Then
            n_����ԤԼ := 0;
          Else
            Select Nvl(Max(Jh.�Ƿ�ԤԼ), 0)
            Into n_����ԤԼ
            From (Select Sd.�ƻ�id, Sd.���, Sd.����, Jh.����,
                          To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.��ʼʱ��, 'hh24:mi:ss'),
                                   'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                          To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.����ʱ��, 'hh24:mi:ss'),
                                   'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, Sd.��������, Sd.�Ƿ�ԤԼ
                   From �ҺŰ��żƻ� Jh, �Һżƻ�ʱ�� Sd
                   Where Jh.Id = Sd.�ƻ�id And Jh.Id = r_�ƻ�id And
                         Sd.���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6',
                                        '����', '7', '����', Null)) Jh, �Һ����״̬ Zt
            Where Zt.����(+) = Jh.��ʼʱ�� And Zt.����(+) = Jh.���� And Zt.���(+) = Jh.��� And
                  Decode(Sign(Sysdate - Jh.��ʼʱ��), -1, 0, 1) <> 1;
          End If;
        
          For r_Time In (Select Jh.����, Jh.���, Jh.����, Jh.��ʼʱ��, Jh.����ʱ��, Jh.��������, Jh.�Ƿ�ԤԼ, 0 As ��Լ��,
                                Decode(Nvl(Zt.���, 0), 0, 1, 0) As ʣ����, Decode(Sign(Sysdate - Jh.��ʼʱ��), -1, 0, 1) As ʧЧʱ��



                         From (Select Sd.�ƻ�id, Sd.���, Sd.����, Jh.����,
                                       To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.��ʼʱ��, 'hh24:mi:ss'),
                                                'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                                       To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.����ʱ��, 'hh24:mi:ss'),
                                                'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, Sd.��������, Sd.�Ƿ�ԤԼ
                                From �ҺŰ��żƻ� Jh, �Һżƻ�ʱ�� Sd
                                Where Jh.Id = Sd.�ƻ�id And Jh.Id = r_�ƻ�id And
                                      Sd.���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                                     '����', '6', '����', '7', '����', Null)) Jh, �Һ����״̬ Zt
                         Where Zt.����(+) = Jh.��ʼʱ�� And Zt.����(+) = Jh.���� And Zt.���(+) = Jh.��� And
                               Decode(Sign(Sysdate - Jh.��ʼʱ��), -1, 0, 1) <> 1
                         Order By ���) Loop
            If v_������λ Is Not Null Then
              Begin
                Select 1
                Into n_��Լģʽ
                From ������λ�ƻ�����
                Where ������Ŀ = r_Time.���� And �ƻ�id = r_�ƻ�id And ��� = 0 And ������λ = v_������λ;
              Exception
                When Others Then
                  n_��Լģʽ := 0;
              End;
            Else
              n_��Լģʽ := 0;
            End If;
            If r_Time.ʣ���� = 0 Then
              n_����ʣ�� := 0;
            Else
              n_����ʣ�� := r_Time.��������;
            End If;
            If v_������λ Is Null Or n_��Լģʽ = 1 Then
              Begin
                Select 1
                Into n_Exists
                From ������λ�ƻ�����
                Where ������Ŀ = r_Time.���� And �ƻ�id = r_�ƻ�id And ��� = r_Time.��� And Rownum < 2;
              Exception
                When Others Then
                  n_Exists := 0;
              End;
              If n_Exists = 0 Then
                If n_����ԤԼ = 1 And r_Time.�Ƿ�ԤԼ = 0 Then
                  Null;
                Else
                  Begin
                    Select 1
                    Into n_�Ƿ�Ԥ��
                    From �Һ����״̬
                    Where ״̬ In (3, 4) And ���� = r_���� And ��� = r_Time.��� And Trunc(����) = Trunc(d_����);
                  Exception
                    When Others Then
                      n_�Ƿ�Ԥ�� := 0;
                  End;
                  If n_�Ƿ�Ԥ�� = 0 Then
                    v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                  To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_����ʣ�� || '</SL>' ||
                                  '</SPAN>';
                    n_ʱ������ := Nvl(n_ʱ������, 0) + n_����ʣ��;
                  End If;
                End If;
              End If;
            Else
              Begin
                Select 1
                Into n_Exists
                From ������λ�ƻ�����
                Where ������Ŀ = r_Time.���� And �ƻ�id = r_�ƻ�id And ��� = r_Time.��� And ������λ = v_������λ And Rownum < 2;
              Exception
                When Others Then
                  n_Exists := 0;
              End;
              Begin
                Select 0
                Into n_�Ǻ�Լ
                From ������λ�ƻ�����
                Where �ƻ�id = r_�ƻ�id And ������λ = v_������λ And Rownum < 2;
              Exception
                When Others Then
                  n_�Ǻ�Լ := 1;
              End;
              If n_Exists = 1 Or n_�Ǻ�Լ = 1 Then
                If n_����ԤԼ = 1 And r_Time.�Ƿ�ԤԼ = 0 Then
                  Null;
                Else
                  Begin
                    Select 1
                    Into n_�Ƿ�Ԥ��
                    From �Һ����״̬
                    Where ״̬ In (3, 4) And ���� = r_���� And ��� = r_Time.��� And Trunc(����) = Trunc(d_����);
                  Exception
                    When Others Then
                      n_�Ƿ�Ԥ�� := 0;
                  End;
                  If n_�Ƿ�Ԥ�� = 0 Then
                    v_Temp         := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                      To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_����ʣ�� || '</SL>' ||
                                      '</SPAN>';
                    n_��Լʣ������ := n_��Լʣ������ + 1;
                  End If;
                End If;
              End If;
            End If;
          End Loop;
        Else
          n_���������� := Nvl(r_��Լ��, Nvl(r_�޺���, 0)) - Nvl(r_��Լ��, 0);
          For r_Time In (Select Jh.����, Jh.���, Jh.����, Jh.��ʼʱ��, Jh.����ʱ��, Jh.��������, Jh.�Ƿ�ԤԼ,
                                Sum(Decode(Nvl(Zt.���, 0), 0, 0, 1)) As ��Լ��,
                                Jh.�������� - Sum(Decode(Nvl(Zt.���, 0), 0, 0, 1)) As ʣ����,
                                Decode(Sign(Sysdate - Jh.��ʼʱ��), -1, 0, 1) As ʧЧʱ��
                         From (Select Sd.�ƻ�id, Sd.���, Sd.����, Jh.����,
                                       To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.��ʼʱ��, 'hh24:mi:ss'),
                                                'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                                       To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.����ʱ��, 'hh24:mi:ss'),
                                                'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, Sd.��������, Sd.�Ƿ�ԤԼ
                                From �ҺŰ��żƻ� Jh, �Һżƻ�ʱ�� Sd
                                Where Jh.Id = Sd.�ƻ�id And Jh.Id = r_�ƻ�id And
                                      Sd.���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                                     '����', '6', '����', '7', '����', Null)) Jh, �Һ����״̬ Zt
                         Where Jh.���� = Zt.����(+) And Jh.��ʼʱ�� = Zt.����(+) And
                               Decode(Sign(Sysdate - Jh.��ʼʱ��), -1, 0, 1) <> 1
                         Group By Jh.����, Jh.���, Jh.����, Jh.��ʼʱ��, Jh.����ʱ��, Jh.��������, Jh.�Ƿ�ԤԼ
                         Order By Jh.���) Loop
            If v_������λ Is Not Null Then
              Begin
                Select 1
                Into n_��Լģʽ
                From ������λ�ƻ�����
                Where ������Ŀ = r_Time.���� And �ƻ�id = r_�ƻ�id And ��� = 0 And ������λ = v_������λ;
              Exception
                When Others Then
                  n_��Լģʽ := 0;
              End;
            Else
              n_��Լģʽ := 0;
            End If;
            n_����ʣ�� := r_Time.ʣ����;
            If v_������λ Is Null Or n_��Լģʽ = 1 Then
              Begin
                Select 1
                Into n_Exists
                From ������λ�ƻ�����
                Where ������Ŀ = r_Time.���� And �ƻ�id = r_�ƻ�id And ��� = r_Time.��� And Rownum < 2;
              Exception
                When Others Then
                  n_Exists := 0;
              End;
              If n_Exists = 0 Then
                If n_���������� < n_����ʣ�� Then
                  v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_���������� || '</SL>' ||
                                '</SPAN>';
                  n_ʱ������ := Nvl(n_ʱ������, 0) + n_����������;
                Else
                  v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_����ʣ�� || '</SL>' ||
                                '</SPAN>';
                  n_ʱ������ := Nvl(n_ʱ������, 0) + n_����ʣ��;
                End If;
              End If;
            Else
              Begin
                Select 1
                Into n_Exists
                From ������λ�ƻ�����
                Where ������Ŀ = r_Time.���� And �ƻ�id = r_�ƻ�id And ��� = r_Time.��� And ������λ = v_������λ And Rownum < 2;
              Exception
                When Others Then
                  n_Exists := 0;
              End;
              Begin
                Select 0
                Into n_�Ǻ�Լ
                From ������λ�ƻ�����
                Where �ƻ�id = r_�ƻ�id And ������λ = v_������λ And Rownum < 2;
              Exception
                When Others Then
                  n_�Ǻ�Լ := 1;
              End;
              If n_Exists = 1 Or n_�Ǻ�Լ = 1 Then
                If n_���������� < n_����ʣ�� Then
                  v_Temp         := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                    To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_���������� || '</SL>' ||
                                    '</SPAN>';
                  n_��Լʣ������ := n_��Լʣ������ + Nvl(n_����ʣ��, 0);
                Else
                  v_Temp         := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                    To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_����ʣ�� || '</SL>' ||
                                    '</SPAN>';
                  n_��Լʣ������ := n_��Լʣ������ + Nvl(n_����ʣ��, 0);
                End If;
              End If;
            End If;
          End Loop;
        End If;
      Else
        Select Max(����ʱ��)
        Into d_�Ӻ�ʱ��
        From �ҺŰ���ʱ��
        Where ����id = r_����id And ���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����',
                                            '6', '����', '7', '����', Null);
        If r_��ſ��� = 1 Then
          If Trunc(d_����) = Trunc(Sysdate) Then
            n_����ԤԼ := 0;
          Else
            Select Nvl(Max(Ap.�Ƿ�ԤԼ), 0)
            Into n_����ԤԼ
            From (Select Sd.����id, Sd.���, Sd.����, Ap.����,
                          To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.��ʼʱ��, 'hh24:mi:ss'),
                                   'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                          To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.����ʱ��, 'hh24:mi:ss'),
                                   'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, Sd.��������, Sd.�Ƿ�ԤԼ
                   From �ҺŰ��� Ap, �ҺŰ���ʱ�� Sd
                   Where Ap.Id = Sd.����id And Ap.Id = r_����id And
                         Sd.���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6',
                                        '����', '7', '����', Null)) Ap, �Һ����״̬ Zt
            Where Zt.����(+) = Ap.��ʼʱ�� And Zt.����(+) = Ap.���� And Zt.���(+) = Ap.��� And
                  Decode(Sign(Sysdate - Ap.��ʼʱ��), -1, 0, 1) <> 1;
          End If;
          For r_Time In (Select Ap.����, Ap.���, Ap.����, Ap.��ʼʱ��, Ap.����ʱ��, Ap.��������, Ap.�Ƿ�ԤԼ, 0 As ��Լ��,
                                Decode(Nvl(Zt.���, 0), 0, 1, 0) As ʣ����, Decode(Sign(Sysdate - Ap.��ʼʱ��), -1, 0, 1) As ʧЧʱ��



                         From (Select Sd.����id, Sd.���, Sd.����, Ap.����,
                                       To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.��ʼʱ��, 'hh24:mi:ss'),
                                                'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                                       To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.����ʱ��, 'hh24:mi:ss'),
                                                'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, Sd.��������, Sd.�Ƿ�ԤԼ
                                From �ҺŰ��� Ap, �ҺŰ���ʱ�� Sd
                                Where Ap.Id = Sd.����id And Ap.Id = r_����id And
                                      Sd.���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                                     '����', '6', '����', '7', '����', Null)) Ap, �Һ����״̬ Zt
                         Where Zt.����(+) = Ap.��ʼʱ�� And Zt.����(+) = Ap.���� And Zt.���(+) = Ap.��� And
                               Decode(Sign(Sysdate - Ap.��ʼʱ��), -1, 0, 1) <> 1
                         Order By ���) Loop
            If v_������λ Is Not Null Then
              Begin
                Select 1
                Into n_��Լģʽ
                From ������λ���ſ���
                Where ������Ŀ = r_Time.���� And ����id = r_����id And ��� = 0 And ������λ = v_������λ;
              Exception
                When Others Then
                  n_��Լģʽ := 0;
              End;
            Else
              n_��Լģʽ := 0;
            End If;
            If r_Time.ʣ���� = 0 Then
              n_����ʣ�� := 0;
            Else
              n_����ʣ�� := r_Time.��������;
            End If;
            If v_������λ Is Null Or n_��Լģʽ = 1 Then
              Begin
                Select 1
                Into n_Exists
                From ������λ���ſ���
                Where ������Ŀ = r_Time.���� And ����id = r_����id And ��� = r_Time.��� And Rownum < 2;
              Exception
                When Others Then
                  n_Exists := 0;
              End;
              If n_Exists = 0 Then
                If n_����ԤԼ = 1 And r_Time.�Ƿ�ԤԼ = 0 Then
                  Null;
                Else
                  Begin
                    Select 1
                    Into n_�Ƿ�Ԥ��
                    From �Һ����״̬
                    Where ״̬ In (3, 4) And ���� = r_���� And ��� = r_Time.��� And Trunc(����) = Trunc(d_����);
                  Exception
                    When Others Then
                      n_�Ƿ�Ԥ�� := 0;
                  End;
                  If n_�Ƿ�Ԥ�� = 0 Then
                    v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                  To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_����ʣ�� || '</SL>' ||
                                  '</SPAN>';
                    n_ʱ������ := Nvl(n_ʱ������, 0) + n_����ʣ��;
                  End If;
                End If;
              End If;
            Else
              Begin
                Select 1
                Into n_Exists
                From ������λ���ſ���
                Where ������Ŀ = r_Time.���� And ����id = r_����id And ��� = r_Time.��� And ������λ = v_������λ And Rownum < 2;
              Exception
                When Others Then
                  n_Exists := 0;
              End;
              Begin
                Select 0
                Into n_�Ǻ�Լ
                From ������λ���ſ���
                Where ����id = r_����id And ������λ = v_������λ And Rownum < 2;
              Exception
                When Others Then
                  n_�Ǻ�Լ := 1;
              End;
              If n_Exists = 1 Or n_�Ǻ�Լ = 1 Then
                If n_����ԤԼ = 1 And r_Time.�Ƿ�ԤԼ = 0 Then
                  Null;
                Else
                  Begin
                    Select 1
                    Into n_�Ƿ�Ԥ��
                    From �Һ����״̬
                    Where ״̬ In (3, 4) And ���� = r_���� And ��� = r_Time.��� And Trunc(����) = Trunc(d_����);
                  Exception
                    When Others Then
                      n_�Ƿ�Ԥ�� := 0;
                  End;
                  If n_�Ƿ�Ԥ�� = 0 Then
                    v_Temp         := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                      To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_����ʣ�� || '</SL>' ||
                                      '</SPAN>';
                    n_��Լʣ������ := n_��Լʣ������ + 1;
                  End If;
                End If;
              End If;
            End If;
          End Loop;
        Else
          n_���������� := Nvl(r_��Լ��, Nvl(r_�޺���, 0)) - Nvl(r_��Լ��, 0);
          For r_Time In (Select Ap.����, Ap.���, Ap.����, Ap.��ʼʱ��, Ap.����ʱ��, Ap.��������, Ap.�Ƿ�ԤԼ,
                                Sum(Decode(Nvl(Zt.���, 0), 0, 0, 1)) As ��Լ��,
                                Ap.�������� - Sum(Decode(Nvl(Zt.���, 0), 0, 0, 1)) As ʣ����,
                                Decode(Sign(Sysdate - Ap.��ʼʱ��), -1, 0, 1) As ʧЧʱ��
                         From (Select Sd.����id, Sd.���, Sd.����, Ap.����,
                                       To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.��ʼʱ��, 'hh24:mi:ss'),
                                                'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                                       To_Date(To_Char(d_����, 'yyyy-mm-dd') || ' ' || To_Char(Sd.����ʱ��, 'hh24:mi:ss'),
                                                'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, Sd.��������, Sd.�Ƿ�ԤԼ
                                From �ҺŰ��� Ap, �ҺŰ���ʱ�� Sd
                                Where Ap.Id = Sd.����id And Ap.Id = r_����id And
                                      Sd.���� = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                                     '����', '6', '����', '7', '����', Null)) Ap, �Һ����״̬ Zt
                         Where Ap.���� = Zt.����(+) And Ap.��ʼʱ�� = Zt.����(+) And
                               Decode(Sign(Sysdate - Ap.��ʼʱ��), -1, 0, 1) <> 1
                         Group By Ap.����, Ap.���, Ap.����, Ap.��ʼʱ��, Ap.����ʱ��, Ap.��������, Ap.�Ƿ�ԤԼ
                         Order By Ap.���) Loop
            If v_������λ Is Not Null Then
              Begin
                Select 1
                Into n_��Լģʽ
                From ������λ���ſ���
                Where ������Ŀ = r_Time.���� And ����id = r_����id And ��� = 0 And ������λ = v_������λ;
              Exception
                When Others Then
                  n_��Լģʽ := 0;
              End;
            Else
              n_��Լģʽ := 0;
            End If;
            n_����ʣ�� := r_Time.ʣ����;
            If v_������λ Is Null Or n_��Լģʽ = 1 Then
              Begin
                Select 1
                Into n_Exists
                From ������λ���ſ���
                Where ������Ŀ = r_Time.���� And ����id = r_����id And ��� = r_Time.��� And Rownum < 2;
              Exception
                When Others Then
                  n_Exists := 0;
              End;
              If n_Exists = 0 Then
                If n_���������� < n_����ʣ�� Then
                  v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_���������� || '</SL>' ||
                                '</SPAN>';
                  n_ʱ������ := Nvl(n_ʱ������, 0) + n_����������;
                Else
                  v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_����ʣ�� || '</SL>' ||
                                '</SPAN>';
                  n_ʱ������ := Nvl(n_ʱ������, 0) + n_����ʣ��;
                End If;
              End If;
            Else
              Begin
                Select 1
                Into n_Exists
                From ������λ���ſ���
                Where ������Ŀ = r_Time.���� And ����id = r_����id And ��� = r_Time.��� And ������λ = v_������λ And Rownum < 2;
              Exception
                When Others Then
                  n_Exists := 0;
              End;
              Begin
                Select 0
                Into n_�Ǻ�Լ
                From ������λ���ſ���
                Where ����id = r_����id And ������λ = v_������λ And Rownum < 2;
              Exception
                When Others Then
                  n_�Ǻ�Լ := 1;
              End;
              If n_Exists = 1 Or n_�Ǻ�Լ = 1 Then
                If n_���������� < n_����ʣ�� Then
                  v_Temp         := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                    To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_���������� || '</SL>' ||
                                    '</SPAN>';
                  n_��Լʣ������ := n_��Լʣ������ + Nvl(n_����ʣ��, 0);
                Else
                  v_Temp         := v_Temp || '<SPAN>' || '<SJD>' || To_Char(r_Time.��ʼʱ��, 'hh24:mi:ss') || '-' ||
                                    To_Char(r_Time.����ʱ��, 'hh24:mi:ss') || '</SJD>' || '<SL>' || n_����ʣ�� || '</SL>' ||
                                    '</SPAN>';
                  n_��Լʣ������ := n_��Լʣ������ + Nvl(n_����ʣ��, 0);
                End If;
              End If;
            End If;
          End Loop;
        End If;
      End If;
    End If;
    If v_������λ Is Not Null Then
      If Nvl(r_�ƻ�id, 0) <> 0 Then
        Begin
          Select 0
          Into n_�Ǻ�Լ
          From ������λ�ƻ�����
          Where �ƻ�id = r_�ƻ�id And ������λ = v_������λ And Rownum < 2;
        Exception
          When Others Then
            n_�Ǻ�Լ := 1;
        End;
      Else
        Begin
          Select 0
          Into n_�Ǻ�Լ
          From ������λ���ſ���
          Where ����id = r_����id And ������λ = v_������λ And Rownum < 2;
        Exception
          When Others Then
            n_�Ǻ�Լ := 1;
        End;
      End If;
    End If;
    If v_������λ Is Null Or n_�Ǻ�Լ = 1 Then
      If r_�޺��� = 0 Then
        v_ʣ������ := '';
      Else
        If Nvl(r_�ƻ�id, 0) <> 0 Then
          Select Sum(����)
          Into n_��Լ������
          From ������λ�ƻ�����
          Where �ƻ�id = r_�ƻ�id And ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                                '����', '6', '����', '7', '����', Null);
        Else
          Select Sum(����)
          Into n_��Լ������
          From ������λ���ſ���
          Where ����id = r_����id And ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                                '����', '6', '����', '7', '����', Null);
        End If;
        Begin
          Select Count(1)
          Into n_��Լ�ѹ���
          From ���˹Һż�¼
          Where �ű� = r_���� And ��¼״̬ = 1 And ������λ Is Not Null And ����ʱ�� Between Trunc(d_����) And
                Trunc(d_���� + 1) - 1 / 60 / 60 / 24;
        Exception
          When Others Then
            n_��Լ�ѹ��� := 0;
        End;
        Select Count(1)
        Into n_Ԥ������
        From �Һ����״̬
        Where ״̬ = 3 And ���� = r_���� And Trunc(����) = Trunc(d_����);
        If Trunc(d_����) = Trunc(Sysdate) Then
          If Nvl(n_��Լ������, 0) = 0 Then
            v_ʣ������ := r_�޺��� - r_�ѹ��� - r_��Լ�� + r_�ѽ��� - n_Ԥ������;
          Else
            v_ʣ������ := r_�޺��� - r_�ѹ��� - r_��Լ�� + r_�ѽ��� - n_��Լ������ + n_��Լ�ѹ��� - n_Ԥ������;
          End If;
          n_�ѹ��� := r_�ѹ���;
          If Nvl(n_ʱ������, 0) < v_ʣ������ And n_��ʱ�� <> 0 Then
            n_������� := 1;
            v_Temp     := v_Temp || '<SPAN>' || '<SJD>' || To_Char(d_�Ӻ�ʱ��, 'hh24:mi:ss') || '-' || '</SJD>' || '<SL>' ||
                          To_Number(v_ʣ������ - Nvl(n_ʱ������, 0)) || '</SL>' || '</SPAN>';
          Else
            n_������� := 0;
          End If;
        Else
          If Nvl(n_��Լ������, 0) = 0 Then
            v_ʣ������ := r_��Լ�� - r_��Լ�� - n_Ԥ������;
            If v_ʣ������ Is Null Then
              v_ʣ������ := r_�޺��� - r_�ѹ��� - r_��Լ�� + r_�ѽ��� - n_Ԥ������;
            End If;
          Else
            v_ʣ������ := r_��Լ�� - r_��Լ�� - n_��Լ������ + n_��Լ�ѹ��� - n_Ԥ������;
            If v_ʣ������ Is Null Then
              v_ʣ������ := r_�޺��� - r_�ѹ��� - r_��Լ�� + r_�ѽ��� - n_��Լ������ + n_��Լ�ѹ��� - n_Ԥ������;
            End If;
          End If;
          n_�ѹ��� := r_�ѹ���;
        End If;
      End If;
    Else
      If Nvl(r_�ƻ�id, 0) <> 0 Then
        If v_������λ Is Not Null Then
          Begin
            Select 1
            Into n_��Լģʽ
            From ������λ�ƻ�����
            Where ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',
                                '7', '����', Null) And �ƻ�id = r_�ƻ�id And ��� = 0 And ������λ = v_������λ;
          Exception
            When Others Then
              n_��Լģʽ := 0;
          End;
        Else
          n_��Լģʽ := 0;
        End If;
        Select Sum(����)
        Into n_��Լ������
        From ������λ�ƻ�����
        Where �ƻ�id = r_�ƻ�id And ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����',
                                              '6', '����', '7', '����', Null) And ������λ = v_������λ;
      Else
        If v_������λ Is Not Null Then
          Begin
            Select 1
            Into n_��Լģʽ
            From ������λ���ſ���
            Where ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',
                                '7', '����', Null) And ����id = r_����id And ��� = 0 And ������λ = v_������λ;
          Exception
            When Others Then
              n_��Լģʽ := 0;
          End;
        Else
          n_��Լģʽ := 0;
        End If;
        Select Sum(����)
        Into n_��Լ������
        From ������λ���ſ���
        Where ����id = r_����id And ������Ŀ = Decode(To_Char(d_����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����',
                                              '6', '����', '7', '����', Null) And ������λ = v_������λ;
      End If;
      If n_��Լģʽ = 0 Then
        v_ʣ������   := n_��Լʣ������;
        n_�ѹ���     := r_�ѹ���;
        n_��Լ�ѹ��� := Nvl(n_��Լ������, 0) - n_��Լʣ������;
      Else
        n_�ѹ��� := r_�ѹ���;
        Begin
          Select Count(1)
          Into n_��Լ�ѹ���
          From ���˹Һż�¼
          Where �ű� = r_���� And ��¼״̬ = 1 And ������λ = v_������λ And ����ʱ�� Between Trunc(d_����) And
                Trunc(d_���� + 1) - 1 / 60 / 60 / 24;
        Exception
          When Others Then
            n_��Լ�ѹ��� := 0;
        End;
        If Nvl(n_��Լ������, 0) = 0 Then
          v_ʣ������ := '0';
        Else
          v_ʣ������ := n_��Լ������ - n_��Լ�ѹ���;
        End If;
      End If;
    End If;
    Select To_Char(��ʼʱ��, 'hh24:mi') Into v_Timetemp From ʱ��� Where ʱ��� = r_�Ű�;
    v_ʱ��� := v_Timetemp || '-';
    Select To_Char(��ֹʱ��, 'hh24:mi') Into v_Timetemp From ʱ��� Where ʱ��� = r_�Ű�;
    v_ʱ��� := v_ʱ��� || v_Timetemp;
    If v_Temp Is Not Null Then
      v_Temp := v_Temp || '</SPANLIST>';
    End If;
    If v_������λ Is Not Null Then
      If Nvl(r_�ƻ�id, 0) <> 0 Then
        Begin
          Select 1
          Into n_����
          From ������λ�ƻ�����
          Where �ƻ�id = r_�ƻ�id And ������λ = v_������λ And ���� = 0 And Rownum < 2;
        Exception
          When Others Then
            n_���� := 0;
        End;
      Else
        Begin
          Select 1
          Into n_����
          From ������λ���ſ���
          Where ����id = r_����id And ������λ = v_������λ And ���� = 0 And Rownum < 2;
        Exception
          When Others Then
            n_���� := 0;
        End;
      End If;
    End If;
	--��Լ��=0��ԤԼ��ֹ
    If Trunc(d_����) <> Trunc(Sysdate) Then
      If r_��Լ�� = 0 Then
        n_���� := 1;
      End If;
    End If;
    If Nvl(n_����, 0) = 0 Then
      --���������
      n_�ϼƽ�� := r_�۸�;
      For r_Subfee In (Select �ּ�, ��������
                       From �շѴ�����Ŀ A, �շѼ�Ŀ B
                       Where a.����id = r_��Ŀid And a.����id = b.�շ�ϸĿid And Sysdate Between b.ִ������ And
                             Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))) Loop
        n_�ϼƽ�� := n_�ϼƽ�� + r_Subfee.�ּ� * r_Subfee.��������;
      End Loop;
      If Trunc(Sysdate) = Trunc(d_����) Then
        Begin
          Select 1
          Into n_Exists
          From (Select ʱ���
                 From ʱ���
                 Where ('3000-01-10 ' || To_Char(Sysdate, 'HH24:MI:SS') < '3000-01-10 ' || To_Char(��ֹʱ��, 'HH24:MI:SS')) Or
                       ('3000-01-10 ' || To_Char(Sysdate, 'HH24:MI:SS') <
                       Decode(Sign(��ʼʱ�� - ��ֹʱ��), 1, '3000-01-11 ' || To_Char(��ֹʱ��, 'HH24:MI:SS'),
                               '3000-01-10 ' || To_Char(��ֹʱ��, 'HH24:MI:SS'))))
          Where ʱ��� = r_�Ű�;
        Exception
          When Others Then
            n_Exists := 0;
        End;
      Else
        n_Exists := 1;
      End If;
      If n_Exists = 1 Then
        If v_ʣ������ > 0 Then
          c_Xmlmain := '<HB>' || '<HM>' || r_���� || '</HM>' || '<YSID>' || r_ҽ��id || '</YSID>' || '<YS>' || r_ҽ������ ||
                       '</YS>' || '<KSID>' || r_����id || '</KSID>' || '<KSMC>' || r_�������� || '</KSMC>' || '<ZC>' || r_ְ�� ||
                       '</ZC>' || '<XMID>' || r_��Ŀid || '</XMID>' || '<XMMC>' || r_��Ŀ���� || '</XMMC>' || '<YGHS>' ||
                       n_�ѹ��� || '</YGHS>' || '<SYHS>' || v_ʣ������ || '</SYHS>' || '<PRICE>' || n_�ϼƽ�� || '</PRICE>' ||
                       '<HCXH>' || n_������� || '</HCXH>' || '<HL>' || r_���� || '</HL>' || '<FSD>' || n_��ʱ�� || '</FSD>' ||
                       '<HBTIME>' || v_ʱ��� || '</HBTIME>' || '<FWMC>' || r_�Ű� || '</FWMC>' || v_Temp || '</HB>';
          v_Xmlmain := v_Xmlmain || c_Xmlmain;
        Else
          c_Xmlmain := '<HB>' || '<HM>' || r_���� || '</HM>' || '<YSID>' || r_ҽ��id || '</YSID>' || '<YS>' || r_ҽ������ ||
                       '</YS>' || '<KSID>' || r_����id || '</KSID>' || '<KSMC>' || r_�������� || '</KSMC>' || '<ZC>' || r_ְ�� ||
                       '</ZC>' || '<XMID>' || r_��Ŀid || '</XMID>' || '<XMMC>' || r_��Ŀ���� || '</XMMC>' || '<YGHS>' ||
                       n_�ѹ��� || '</YGHS>' || '<SYHS>' || v_ʣ������ || '</SYHS>' || '<PRICE>' || n_�ϼƽ�� || '</PRICE>' ||
                       '<HL>' || r_���� || '</HL>' || '<FSD>' || n_��ʱ�� || '</FSD>' || '<HBTIME>' || v_ʱ��� || '</HBTIME>' ||
                       '<FWMC>' || r_�Ű� || '</FWMC>' || '</HB>';
          v_Xmlmain := v_Xmlmain || c_Xmlmain;
        End If;
      End If;
    End If;
    n_��Լʣ������ := 0;
    n_��Լ������   := 0;
    n_ʱ������     := 0;
    n_����         := 0;
    n_�Ǻ�Լ       := 0;
  End Loop;
  Close r_No;
  v_Xmlmain := '<GROUP>' || '<RQ>' || To_Char(d_����, 'yyyy-mm-dd') || '</RQ>' || '<HBLIST>' || v_Xmlmain || '</HBLIST>' ||
               '</GROUP>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Xmlmain)) Into x_Templet From Dual;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getnolist;
/

--92949:����,2016-01-22,��Ա��������̫��ʱ��������
Create Or Replace Procedure Zl_��Ա��_�޸�
(
  Id_In           In ��Ա��.Id%Type,
  ���_In         In ��Ա��.���%Type,
  ����_In         In ��Ա��.����%Type,
  ����_In         In ��Ա��.����%Type,
  ���֤��_In     In ��Ա��.���֤��%Type,
  ��������_In     In ��Ա��.��������%Type,
  �Ա�_In         In ��Ա��.�Ա�%Type,
  ����_In         In ��Ա��.����%Type,
  ��������_In     In ��Ա��.��������%Type,
  �칫�ҵ绰_In   In ��Ա��.�칫�ҵ绰%Type,
  �����ʼ�_In     In ��Ա��.�����ʼ�%Type,
  ִҵ���_In     In ��Ա��.ִҵ���%Type,
  ִҵ��Χ_In     In ��Ա��.ִҵ��Χ%Type,
  ����ְ��_In     In ��Ա��.����ְ��%Type,
  רҵ����ְ��_In In ��Ա��.רҵ����ְ��%Type,
  Ƹ�μ���ְ��_In In ��Ա��.Ƹ�μ���ְ��%Type,
  ѧ��_In         In ��Ա��.ѧ��%Type,
  ��ѧרҵ_In     In ��Ա��.��ѧרҵ%Type,
  ��ѧʱ��_In     In ��Ա��.��ѧʱ��%Type,
  ��ѧ����_In     In ��Ա��.��ѧ����%Type,
  ������ѵ_In     In ��Ա��.������ѵ%Type,
  ���п���_In     In ��Ա��.���п���%Type,
  ���˼��_In     In ��Ա��.���˼��%Type,
  �����б�_In     In Varchar2, --�����б�_IN��������д��ʽ���£�"12:1;23:0;"
  ��Ա����_In     In Varchar2, --��Ա����_IN��������д��ʽ���£�"����Һ�Ա;ҽ��;��ʿ;"
  ����_In         In ��Ա��.����%Type := Null,
  վ��_In         In ��Ա��.վ��%Type := Null,
  ǩ��_In         In ��Ա��.ǩ��%Type := Null,
  ִҵ֤��_In     In ��Ա��.ִҵ֤��%Type := Null,
  �ʸ�֤���_In   In ��Ա��.�ʸ�֤���%Type := Null,
  ִҵ��ʼ����_In In ��Ա��.ִҵ��ʼ����%Type := Null,
  ����Ȩ��־_In   In ��Ա��.����Ȩ��־%Type := Null,
  �����ȼ�_In     In ��Ա��.�����ȼ�%Type := Null,
  �ƶ��绰_In     In ��Ա��.�ƶ��绰%Type := Null
) Is
  Intpos    Pls_Integer;
  Intȱʡ   Number(1);
  Strtemp   Varchar2(2000);
  Str����   Varchar2(10);
  Lng����id ���ű�.Id%Type;
Begin
  --���Ȳ����޸ļ�¼
  Update ��Ա��
  Set ��� = ���_In, ���� = ����_In, ���� = ����_In, ���֤�� = ���֤��_In, �������� = ��������_In, �Ա� = �Ա�_In, ���� = ����_In, �������� = ��������_In,
      �칫�ҵ绰 = �칫�ҵ绰_In, �����ʼ� = �����ʼ�_In, ִҵ��� = ִҵ���_In, ִҵ��Χ = ִҵ��Χ_In, ����ְ�� = ����ְ��_In, רҵ����ְ�� = רҵ����ְ��_In,
      Ƹ�μ���ְ�� = Ƹ�μ���ְ��_In, ѧ�� = ѧ��_In, ��ѧרҵ = ��ѧרҵ_In, ��ѧʱ�� = ��ѧʱ��_In, ��ѧ���� = ��ѧ����_In, ������ѵ = ������ѵ_In, ���п��� = ���п���_In,
      ���˼�� = ���˼��_In, վ�� = վ��_In, ���� = ����_In, ǩ�� = ǩ��_In, ִҵ֤�� = ִҵ֤��_In, �ʸ�֤��� = �ʸ�֤���_In, ִҵ��ʼ���� = ִҵ��ʼ����_In,
      ����Ȩ��־ = ����Ȩ��־_In, �����ȼ� = �����ȼ�_In, �ƶ��绰 = �ƶ��绰_In
  Where ID = Id_In;

  --����ɾ�����е���������
  Delete From ������Ա Where ��Աid = Id_In;

  --�����޸���������
  Strtemp := �����б�_In;

  While Strtemp Is Not Null Loop
    Intpos := Instr(Strtemp, ':');
  
    If Intpos = 0 Then
      Strtemp := '';
    Else
      --�õ�����ID
      Str����   := Substr(Strtemp, 1, Intpos - 1);
      Lng����id := To_Number(Str����);
      Strtemp   := Substr(Strtemp, Intpos + 1);
      --�õ��Ƿ�ȱʡ
      Intpos  := Instr(Strtemp, ';');
      Intȱʡ := To_Number(Substr(Strtemp, 1, Intpos - 1));
      Strtemp := Substr(Strtemp, Intpos + 1);
    
      Insert Into ������Ա (����id, ��Աid, ȱʡ) Values (Lng����id, Id_In, Intȱʡ);
    End If;
  End Loop;

  --����ɾ�����е�����˵��
  Delete From ��Ա����˵�� Where ��Աid = Id_In;

  --����޸���Ա����˵��
  Strtemp := ��Ա����_In;

  While Strtemp Is Not Null Loop
    Intpos := Instr(Strtemp, ';');
  
    If Intpos = 0 Then
      Strtemp := '';
    Else
      --�õ���Ա����
      Str���� := Substr(Strtemp, 1, Intpos - 1);
      Strtemp := Substr(Strtemp, Intpos + 1);
    
      Insert Into ��Ա����˵�� (��Ա����, ��Աid) Values (Str����, Id_In);
    End If;
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Ա��_�޸�;
/

--91316:�ŵ���,2016-03-10,�����Զ���������
--91954:�ŵ���,2016-02-29,�Զ��������
CREATE OR REPLACE Procedure Zl_��Һ��ҩ��¼_�Զ�����
(
 ����id_In In number,
 ����id_In In number,
 ����id_In In number,
 ִ������_In in date
) Is
v_���δ� varchar2(500);
n_���� number(5);
v_��ҩid  varchar2(500);
v_batch varchar2(20);
v_Fields varchar2(100);
v_Tansid varchar(18);
n_����id number(18);
v_����   varchar2(10);
v_Id     varchar2(200);
n_�Զ�����ģʽ  number(1);
Begin
  n_�Զ�����ģʽ:=Zl_To_Number(Nvl(zl_GetSysParameter('�Զ�����ʱ��Һ��������ֻ���������α䶯', 1345), 0));
  --�ÿ��Ҹ������ζ�Ӧ������
  select max(����) into v_���� from ��ҩ�������� where ��������id=����id_In and ҩƷ���� is  null;
  for R_Batch in (select B.���� ��ҩ����,A.����,A.����id from ������������ A,��ҩ�������� B where A.��������ID=B.��������ID And A.��ҩ����=(B.���� || '#') and (A.����id=����id_In or A.����ID=0) and A.��������id=����id_In order by A.����id desc, A.��ҩ���� asc) loop

    n_����id:=R_Batch.����id;

  --�ò��˰���������ִ��ʱ�䣬���ȼ�����,�����������е����������ȼ��ڲ�����Һ��ҩ��¼��ʱ��д��
    n_����:=0;
    for r_item in (Select a.Id ��ҩid, d.����,A.ƿǩ��,A.��ҩ����
    From ��Һ��ҩ��¼ A, ����ҽ����¼ B, ��Һ��ҩ���� C, ҩƷ�շ���¼ D, ҩƷ��� E, ҩƷ���� F,��ҩ�������� G
    Where a.Id = c.��¼id And c.�շ�id = d.Id And d.ҩƷid = e.ҩƷid And e.ҩ��id = f.ҩ��id And a.����id = ����id_In and G.��������ID=a.����id And
          a.ҽ��id = b.Id And b.����id = ����id_In And A.��ҩ����=G.���� And G.����<>0 and G.ҩƷ���� is null And f.��ý = 1 And a.ִ��ʱ�� Between Trunc(ִ������_In ) And
          Trunc(ִ������_In+1) - 1 / 24 / 60 / 60 and  A.��ҩ����<=decode(n_�Զ�����ģʽ,1,R_Batch.��ҩ����,100)
    Order By a.���ȼ�,a.��ҩ����, a.ִ��ʱ��, d.���� desc) loop

      if instr(','|| v_��ҩid,','|| r_item.��ҩid || ',',1)<1then
        --������ҩid�״�ѭ����ʱ����䵥�������ۼ�
        v_��ҩid:=v_��ҩid || r_item.��ҩid || ',';
        n_����:=n_����+r_item.����;
        v_���δ�:=v_���δ� || r_item.��ҩid || ',' || R_Batch.��ҩ���� || '|';
      elsif instr('|'|| v_���δ�,'|'|| r_item.��ҩid || ',' || R_Batch.��ҩ���� || '|',1)>0 and n_����<>0 then
        --������ҩid�����γ��ֹ������ۼƸõ���
        n_����:=n_����+r_item.����;
      end if;

      if n_����>=R_Batch.���� then
        exit;
      end if;

    end loop;


  end loop;

  --����ÿ��������˵�����������Ϣ�򲻿������п��ҵ�ģʽ
  for r_item in (Select a.Id ��ҩid, d.����,A.ƿǩ��
  From ��Һ��ҩ��¼ A, ����ҽ����¼ B, ��Һ��ҩ���� C, ҩƷ�շ���¼ D, ҩƷ��� E, ҩƷ���� F,��ҩ�������� G
  Where a.Id = c.��¼id And c.�շ�id = d.Id And d.ҩƷid = e.ҩƷid And e.ҩ��id = f.ҩ��id And a.����id = ����id_In  And
        a.ҽ��id = b.Id And b.����id = ����id_In And A.��ҩ����=G.���� And G.����<>0 and G.��������ID=a.����id and G.ҩƷ���� is null And f.��ý = 1 And a.ִ��ʱ�� Between Trunc(ִ������_In) And
        Trunc(ִ������_In+1) - 1 / 24 / 60 / 60
  Order By a.��ҩ����,a.���ȼ�,a.ִ��ʱ��, d.����) loop
    if instr(','|| v_��ҩid,','|| r_item.��ҩid || ',',1)<1then
      --û�н����Զ���������Һ��ֱ�������һ������
      v_��ҩid:=v_��ҩid || r_item.��ҩid || ',';
      v_���δ�:=v_���δ� || r_item.��ҩid || ',' || v_���� || '|';
    end if;
  end loop;


  --�����Զ�������������
  while v_���δ� is not null loop
    --�ֽⵥ��ID��
    v_Fields := Substr(v_���δ�, 1, Instr(v_���δ�, '|') - 1);
    v_Tansid := Substr(v_Fields, 1, Instr(v_Fields, ',') - 1);
    v_batch   := Substr(v_Fields, Instr(v_Fields, ',') + 1);

    v_���δ� := Replace('|' || v_���δ�, '|' || v_Fields || '|');

    update ��Һ��ҩ��¼ set ��ҩ����=v_batch where id=v_Tansid;
  end loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Һ��ҩ��¼_�Զ�����;
/

--93011:��˶,2016-01-22,�����������չ����н��ղ������ڵ��Ӳ�������в��ܿ���
Create Or Replace Procedure Zl_�������ռ�¼_Insert
(
  ����id_In   In �������ռ�¼.����id%Type,
  ��ҳid_In   In �������ռ�¼.��ҳid%Type,
  ������_In   In �������ռ�¼.������%Type,
  ������_In   In �������ռ�¼.������%Type,
  ����ʱ��_In In �������ռ�¼.����ʱ��%Type,
  ��¼ʱ��_In In �������ռ�¼.��¼ʱ��%Type
) Is
  n_Id      �������ռ�¼.Id%Type;
  n_Count   Number := 0;
  v_Err_Msg Varchar2(2000);
  Err_Item Exception;
  n_�ύid Number(18);
Begin
  Select Count(����id) Into n_Count From �������ռ�¼ Where ����id = ����id_In And ��ҳid = ��ҳid_In;
  If Nvl(n_Count, 0) > 0 Then
    Update �������ռ�¼
    Set ����id = ����id_In, ��ҳid = ��ҳid_In, ������ = ������_In, ������ = ������_In, ����ʱ�� = ����ʱ��_In, ��¼ʱ�� = ��¼ʱ��_In
    Where ����id = ����id_In And ��ҳid = ��ҳid_In;
  Else
    Select �������ռ�¼_Id.Nextval Into n_Id From Dual;
    Insert Into �������ռ�¼
      (ID, ����id, ��ҳid, ������, ������, ����ʱ��, ��¼ʱ��)
    Values
      (n_Id, ����id_In, ��ҳid_In, ������_In, ������_In, ����ʱ��_In, ��¼ʱ��_In);
  End If;
  If Sql%NotFound Then
    v_Err_Msg := '[ZLSOFT]�������չ������Ӽ��޸Ľ��ռ�¼ʧ�ܣ���˲����ԣ�[ZLSOFT]';
    Raise Err_Item;
  End If;

  Begin
    --���Ӳ���������Ҫ����Ϊ���ܶ�����װ������ֻ�ܶ�ִ̬��
    Execute Immediate 'Select ID From �����ύ��¼ Where ����ID=:1 and ��ҳid=:2'
      Into n_�ύid
      Using ����id_In, ��ҳid_In;
    Execute Immediate 'CALL ZL_�����ύ��¼_RECEIVE(:1,:2)'
      Using n_�ύid, ������_In;
  Exception
    When Others Then
      Null;
  End;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�������ռ�¼_Insert;
/

--92918:������,2016-01-21,����䶯��¼��װ�ű�����
Create Or Replace Procedure Zl_����䶯��¼_Insert
(
  �Һŵ���_In   In ���˹Һż�¼.No%Type,
  �������_In   In ����䶯��¼.���%Type,
  �䶯ԭ��_In   In ����䶯��¼.�䶯ԭ��%Type,
  ����Ա����_In In ����䶯��¼.����Ա����%Type,
  ����Ա���_In In ����䶯��¼.����Ա���%Type,
  ����_In       In ����䶯��¼.�ֺ���%Type := Null,
  ����id_In     In ����䶯��¼.�ֿ���id%Type := Null,
  ��Ŀid_In     In ����䶯��¼.����Ŀid%Type := Null,
  ҽ��id_In     In ����䶯��¼.��ҽ��id%Type := Null,
  ҽ������_In   In ����䶯��¼.��ҽ������%Type := Null,
  ����_In       In ����䶯��¼.������%Type := Null,
  ����_In       In ����䶯��¼.�ֺ���%Type := Null,
  ԤԼʱ��_In   In ����䶯��¼.��ԤԼʱ��%Type := Null
) Is
  v_Err_Msg Varchar2(255);
  Err_Item Exception;
  -----------------------------------------------------------
  --˵�����ҺŻ���ǰ���ã���¼���ű䶯������_IN�Ժ��ֵ������������ޱ䶯
  --      �������_IN = 1-��������;2-���ﻻ��;3-ǿ�����ﻻ��
  -----------------------------------------------------------
Begin
  Insert Into ����䶯��¼
    (ID, ���, �Һŵ�, ����id, �䶯ԭ��, ԭ����, �ֺ���, ԭ����id, �ֿ���id, ԭ��Ŀid, ����Ŀid, ԭҽ��id, ��ҽ��id, ԭҽ������, ��ҽ������, ԭ����, ������, ԭ����, �ֺ���,
     ԭԤԼʱ��, ��ԤԼʱ��, �Ǽ�ʱ��, ����Ա����, ����Ա���)
    Select ����䶯��¼_Id.Nextval, �������_In, �Һŵ���_In, a.����id, �䶯ԭ��_In, a.�ű�, Nvl(����_In, a.�ű�), a.ִ�в���id,
           Nvl(����id_In, a.ִ�в���id), b.�շ�ϸĿid, Nvl(��Ŀid_In, b.�շ�ϸĿid), c.Id, ҽ��id_In, a.ִ����, ҽ������_In, a.����, ����_In, a.����,
           Nvl(����_In, a.����), a.ԤԼʱ��, Nvl(ԤԼʱ��_In, a.ԤԼʱ��), Sysdate, ����Ա����_In, ����Ա���_In
    From ���˹Һż�¼ A, ������ü�¼ B, ��Ա�� C
    Where a.No = �Һŵ���_In And a.��¼״̬ = 1 And a.ִ���� = c.����(+) And b.No = a.No And b.��¼���� = 4 And b.��� = 1 And Rownum < 2;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����䶯��¼_Insert;
/

--92984:����,2016-01-21,���Ź����׼��������ϵͳ����ƫ���
Create Or Replace Procedure Zl_���ű�_Stop(Id_In In ���ű�.ID%Type) Is
Begin
  Update ���ű� Set ����ʱ�� = Sysdate Where ID = Id_In;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���ű�_Stop;
/

--92984:����,2016-01-21,���Ź����׼��������ϵͳ����ƫ���
Create Or Replace Procedure Zl_���ű�_Reuse(Id_In In ���ű�.ID%Type) Is
Begin
  Update ���ű� Set ����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd') Where ID = Id_In;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���ű�_Reuse;
/

--91316:�ŵ���,2016-01-21,��Һ������������
CREATE OR REPLACE Procedure Zl_��ҩ��������_Save
( 
  ����id_In In Varchar2, --�������δ�:����1,��ҩʱ��1,��ҩʱ��1,���1,����1,��ɫ1,ҩƷ����1|.... 
  ��������ID_In In ��ҩ��������.��������ID%type 
) Is 
  v_����     Varchar2(20); 
  v_��ҩʱ�� Varchar2(20); 
  v_��ҩʱ�� Varchar2(20); 
  n_���     Number(1); 
  v_����     Number(1); 
  v_��ɫ     number(18); 
  v_Field    Varchar2(500); 
  v_Tmp      Varchar2(500); 
  v_ҩƷ���� varchar2(50);
Begin 
  If ����id_In Is Null Then 
    v_Tmp := Null; 
  Else 
    v_Tmp := ����id_In || '|'; 
  End If; 
 
  Delete From ��ҩ�������� where ��������ID=��������ID_In; 
 
  While v_Tmp Is Not Null Loop 
    --�ֽ� 
    v_Field := Substr(v_Tmp, 1, Instr(v_Tmp, '|') - 1); 
    v_Tmp   := Replace('|' || v_Tmp, '|' || v_Field || '|'); 
 
    v_����  := Substr(v_Field, 1, Instr(v_Field, ',') - 1); 
    v_Field := Substr(v_Field, Instr(v_Field, ',') + 1); 
 
    v_��ҩʱ�� := Substr(v_Field, 1, Instr(v_Field, ',') - 1); 
    v_Field    := Substr(v_Field, Instr(v_Field, ',') + 1); 
 
    v_��ҩʱ�� := Substr(v_Field, 1, Instr(v_Field, ',') - 1); 
    v_Field    := Substr(v_Field, Instr(v_Field, ',') + 1); 
 
    n_��� := Substr(v_Field, 1, Instr(v_Field, ',') - 1); 
    v_Field    := Substr(v_Field, Instr(v_Field, ',') + 1); 
 
    v_���� := Substr(v_Field,1, Instr(v_Field, ',')- 1); 
    v_Field    := Substr(v_Field, Instr(v_Field, ',') + 1);
    
    v_��ɫ :=Substr(v_Field,1, Instr(v_Field, ',')- 1); 
    v_ҩƷ���� := Substr(v_Field, Instr(v_Field, ',')+ 1); 
 
    Insert Into ��ҩ�������� 
      (����, ��ҩʱ��, ��ҩʱ��, ���, ����,��ɫ,ҩƷ����,��������ID) 
    Values 
      (v_����, v_��ҩʱ��, v_��ҩʱ��, n_���, v_����,v_��ɫ,v_ҩƷ����,��������ID_In); 
  End Loop; 
Exception 
  When Others Then 
    zl_ErrorCenter(SQLCode, SQLErrM); 
End Zl_��ҩ��������_Save;
/

--91333:�ŵ���,2016-01-21,��Һ�������Ĵ�ӡ��ˮ��
CREATE OR REPLACE Procedure Zl_��Һ��ҩ��¼_��ӡ
( 
  ��ҩid_In   In Varchar2, --ID��:ID1,ID2.... 
  ��ӡʱ��_In In ��Һ��ҩ��¼.��ӡʱ��%Type 
) Is 
  v_Tansid Varchar2(20); 
  v_Tmp    Varchar2(4000); 
  n_Count  Number(5); 
  n_Row    NUmber(5);
Begin 
  
  n_Count := 0; 
  If ��ҩid_In Is Null Then 
    v_Tmp := Null; 
  Else 
    v_Tmp := ��ҩid_In || ','; 
  End If; 
  While v_Tmp Is Not Null Loop 
    --�ֽⵥ��ID�� 
    v_Tansid := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1); 
    v_Tmp    := Substr(v_Tmp, Instr(v_Tmp, ',') + 1); 
 
    n_Count := n_Count + 1; 
    select count(id) into n_Row from ��Һ��ҩ��¼ where Nvl(��ӡ��־, 0)<>0 and ��ӡʱ�� between Trunc(��ӡʱ��_In ) And Trunc(��ӡʱ��_In+1) - 1 / 24 / 60 / 60;
    Update ��Һ��ҩ��¼ 
    Set ��ӡ��־ = decode(Nvl(��ӡ��־, 0),0,n_Row+1,��ӡ��־), ��ӡʱ�� = ��ӡʱ��_In, ��ӡ��� = n_Count 
    Where ID = v_Tansid; 
  End Loop; 
Exception 
  When Others Then 
    zl_ErrorCenter(SQLCode, SQLErrM); 
End Zl_��Һ��ҩ��¼_��ӡ;
/

--93964:���Ʊ�,2016-03-07,������ͬʱ��䶯��¼�޷���������
--92938:������,2016-01-21,���˱䶯��¼�������������ű��Ͱ�װ�ű���ͬ����
Create Or Replace Procedure Zl_���˱䶯��¼_Undo
(
  ����id_In     ������ҳ.����id%Type,
  ��ҳid_In     ������ҳ.��ҳid%Type,
  ����Ա���_In ���˱䶯��¼.����Ա���%Type,
  ����Ա����_In ���˱䶯��¼.����Ա����%Type,
  ����_In       Varchar2 := Null, --a.תΪסԺʱ,���סԺ��,b-����Զ����ʷ����Ƿ��ѽ���
  ����_In       Varchar2 := Null, --����ʱ��ʾ������Ժʱ���ŵ��µĴ��ţ�ԭ��λ��ռ���ڳ������ж�
  ����λ_In     Varchar2 := Null, --����ʱ��ʾ������Ժʱ���ŵ��µ�����λ��ԭ��λ��ռ���ڳ������ж�
  ������ʽ_In   Varchar2 := Null --ָ�����峷���������糷����Ժ��ת�Ƶȱ�������
) As
  -----------------------------------------------------------
  --˵����1.�����������һ�εı䶯
  --        2.ǰ�᣺�����˰���ʱ,������һ�Ŵ�λ���䶯,�����д�λ��Ӧ�����䶯
  -----------------------------------------------------------
  --Ҫ�����ı䶯��¼(�������,���ܶ���)
  Cursor c_Curlog Is
    Select *
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And (��ֹʱ�� Is Null Or ��ֹԭ�� = 1)
    Order By ��ֹʱ�� Desc, ��ʼʱ�� Desc;
  r_Curlogrow c_Curlog%Rowtype;

  --������Ҫ�ָ��ı䶯��¼(�������,���ܶ���)
  Cursor c_Prelog
  (
    v_��ֹʱ�� ���˱䶯��¼.��ֹʱ��%Type,
    v_��ֹԭ�� ���˱䶯��¼.��ֹԭ��%Type
  ) Is
    Select *
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� = v_��ֹʱ�� And ��ֹԭ�� = v_��ֹԭ��
    Order By ��ֹʱ�� Desc, ��ʼʱ�� Desc;
  r_Prelogrow c_Prelog%Rowtype;

  --��ȡ����ԭ��λ��ס������Ϣ
  Cursor c_Prebed
  (
    v_����id ������ҳ.����id%Type,
    v_��ҳid ������ҳ.��ҳid%Type
  ) Is
    Select a.����, c.��Ժ����, c.��Ժ����id, c.��ǰ����id
    From ���˱䶯��¼ a, ��λ״����¼ b, ������ҳ c
    Where a.����id = ����id_In And a.��ҳid = ��ҳid_In And ��ֹԭ�� = 4 And a.����id = b.����id And a.����id = c.����id And
          a.��ҳid = c.��ҳid And a.���� = b.����
    Order By ��ֹʱ�� Desc, ��ʼʱ�� Desc;
  r_Prebedrow c_Prebed%Rowtype;

  Cursor c_Prebedpati
  (
    v_��Ժ����id ���˱䶯��¼.����id%Type,
    v_����id     ���˱䶯��¼.����id%Type,
    v_ԭ����     ���˱䶯��¼.����%Type
  ) Is
    Select a.����id, c.��ҳid, a.����, c.��Ժ����
    From ���˱䶯��¼ a, ������ҳ c,
         (Select ����id
           From ��λ״����¼
           Where (����id Is Null Or ����id = v_��Ժ����id Or ���� = 1) And ����id = v_����id And ���� = v_ԭ����) d
    Where a.����id = d.����id And a.��ҳid = (Select ��ҳid From ������Ϣ Where ����id = d.����id) And a.��ֹԭ�� = 4 And a.����id = c.����id And
          a.��ҳid = c.��ҳid
    Order By ��ֹʱ�� Desc, ��ʼʱ�� Desc;
  r_Prebedpati c_Prebedpati%Rowtype;

  v_��ʼʱ�� ���˱䶯��¼.��ʼʱ��%Type;
  v_��ʼԭ�� ���˱䶯��¼.��ʼԭ��%Type;
  v_��ֹ��Ա ���˱䶯��¼.��ֹ��Ա%Type;

  v_Count       Number;
  v_Countcurlog Number;
  v_Countprelog Number;

  Err_Custom Exception;
  v_Error Varchar2(255);

  v_������ʽ     Varchar2(100);
  v_�����       Zlsystems.�����%Type;
  v_����״̬     Number(3);
  v_���Ŵ�       Varchar2(255);
  v_����         ���˱䶯��¼.����%Type;
  v_����id       ���˱䶯��¼.����id%Type;
  v_��ҳid       ���˱䶯��¼.��ҳid%Type;
  v_����id       ���˱䶯��¼.����id%Type;
  v_ԭ����1      ���˱䶯��¼.����%Type;
  v_ԭ����2      ���˱䶯��¼.����%Type;
  v_��ǰ����1    ���˱䶯��¼.����%Type;
  v_��ǰ����2    ���˱䶯��¼.����%Type;
  v_��Ժ����id   ���˱䶯��¼.����id%Type;
  v_��λ�ȼ�id   ���˱䶯��¼.��λ�ȼ�id%Type;
  v_����         ������ҳ.����%Type;
  v_����         ������Ϣ.����%Type;
  n_ԭ����id     ������ҳ.��Ժ����id%Type;
  n_ԭ����id     ������ҳ.��ǰ����id%Type;
  v_ĸӤת�Ʊ�־ ������ҳ.ĸӤת�Ʊ�־%Type;
  v_Tmp          Varchar2(100);
  d_��ʼʱ��     Date;
Begin
  If ������ʽ_In Is Null Then
    v_Error := '[ZLSOFT]û��ָ������ĳ���������[ZLSOFT]';
    Raise Err_Custom;
  Else
    v_������ʽ := ������ʽ_In;
  End If;

  Open c_Curlog;
  Fetch c_Curlog
    Into r_Curlogrow;
  If c_Curlog%Rowcount = 0 Then
    v_Error := '[ZLSOFT]���˵�ǰû�п��Գ����Ĳ�����[ZLSOFT]';
    Close c_Curlog;
    Raise Err_Custom;
  End If;
  
  Select Count(*) Into v_Count From ������ҳ Where ����id = ����id_In And ��ҳid > ��ҳid_In;
  If v_Count > 0 Then
    v_Error := '[ZLSOFT]��ֻ�ܶԲ��˵����һ��סԺ���г�������,���γ���������ֹ![ZLSOFT]';
    Raise Err_Custom;
  End If;
  
  Select Count(Id)
  Into v_Countcurlog
  From ���˱䶯��¼
  Where ����id = ����id_In And ��ҳid = ��ҳid_In And (��ֹʱ�� Is Null Or ��ֹԭ�� = 1)
  Order By ��ֹʱ�� Desc, ��ʼʱ�� Desc;

  Select Count(a.����)
  Into v_Countprelog
  From ���˱䶯��¼ a, ��λ״����¼ b, ������ҳ c
  Where a.����id = ����id_In And a.��ҳid = ��ҳid_In And ��ֹԭ�� = 4 And a.����id = b.����id And a.����id = c.����id And a.��ҳid = c.��ҳid And
        a.���� = b.����
  Order By ��ֹʱ�� Desc, ��ʼʱ�� Desc;

  --�ж��Ƿ�����λ�Ի�
  If v_������ʽ = '����' And v_Countcurlog <= 1 And v_Countprelog <= 1 Then
    Open c_Prebed(����id_In, ��ҳid_In);
    Fetch c_Prebed
      Into r_Prebedrow;
  
    v_��Ժ����id := r_Prebedrow.��Ժ����id;
    v_����id     := r_Prebedrow.��ǰ����id;
    v_ԭ����1    := r_Prebedrow.����;
    v_��ǰ����1  := r_Prebedrow.��Ժ����;
  
    For r_Prebedpati In c_Prebedpati(v_��Ժ����id, v_����id, v_ԭ����1) Loop
      v_����id    := r_Prebedpati.����id;
      v_��ҳid    := r_Prebedpati.��ҳid;
      v_ԭ����2   := r_Prebedpati.����;
      v_��ǰ����2 := r_Prebedpati.��Ժ����;
    
      If v_����id <> 0 And v_��ҳid <> 0 And v_ԭ����1 = v_��ǰ����2 And v_ԭ����2 = v_��ǰ����1 Then
        v_������ʽ := '��λ�Ի�';
        Select ���� Into v_���� From ������ҳ Where ����id = v_����id And ��ҳid = v_��ҳid;
        If v_���� Is Null Then
          Zl_���˱䶯��¼_Undo(v_����id, v_��ҳid, ����Ա���_In, ����Ա����_In, Null, ����_In, ����λ_In, v_������ʽ);
        Else
          Zl_���˱䶯��¼_Undo(v_����id, v_��ҳid, ����Ա���_In, ����Ա����_In, '1', ����_In, ����λ_In, v_������ʽ);
        End If;
      End If;
    
      --ֻ�����һ�δ�λ�ǶԷ���λ�ļ�¼���д���
      Exit;
    End Loop;
  End If;

  If r_Curlogrow.��ֹʱ�� Is Null And r_Curlogrow.��ʼʱ�� Is Null And r_Curlogrow.��ʼԭ�� = 3 And v_������ʽ = 'ת��' Then
    --����ת��(��־)
    Delete From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ʼʱ�� Is Null And ��ֹʱ�� Is Null And ��ʼԭ�� = 3;
  
    Update ������ҳ Set ״̬ = 0 Where ����id = ����id_In And ��ҳid = ��ҳid_In;
  
    Close c_Curlog;
  Elsif r_Curlogrow.��ֹʱ�� Is Null And r_Curlogrow.��ʼʱ�� Is Null And r_Curlogrow.��ʼԭ�� = 15 And v_������ʽ = 'ת����' Then
    --����ת��(��־)
    Delete From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ʼʱ�� Is Null And ��ֹʱ�� Is Null And ��ʼԭ�� = 15;
  
    Update ������ҳ Set ״̬ = 0 Where ����id = ����id_In And ��ҳid = ��ҳid_In;
  
    Close c_Curlog;
  Elsif r_Curlogrow.��ֹʱ�� Is Not Null And r_Curlogrow.��ֹԭ�� = 1 And v_������ʽ = '��Ժ' Then
    --������Ժ
    v_��ʼʱ�� := r_Curlogrow.��ֹʱ��; --�����ı䶯��¼�Ŀ�ʼʱ��
  
    Select Zl_סԺ�ձ�_Count(r_Curlogrow.����id, r_Curlogrow.��ֹʱ��) Into v_Count From Dual;
    If v_Count > 0 Then
      v_Error := '[ZLSOFT]�Ѳ���ҵ��ʱ���ڵ�סԺ�ձ�,���ܰ����ҵ��![ZLSOFT]';
      Raise Err_Custom;
    End If;
    --�Ƿ���й����Ӳ������
    Select Nvl(����״̬, 0) Into v_����״̬ From ������ҳ Where ����id = ����id_In And ��ҳid = ��ҳid_In;
    If v_����״̬ Not In (0, 2) Then
      v_Error := '[ZLSOFT]���˵ĵ��Ӳ������ύ��飬�����ٳ�����Ժ��[ZLSOFT]';
      Close c_Curlog;
      Raise Err_Custom;
    End If;
  
    --ɾ��������дʱ��
    Zl_���Ӳ���ʱ��_Delete(����id_In, ��ҳid_In, '��Ժ', r_Curlogrow.����id);
  
    --�ָ���Ժ
    Update ���˱䶯��¼
    Set ��ֹʱ�� = Null, ��ֹԭ�� = Null, ��ֹ��Ա = Null, �ϴμ���ʱ�� = Null
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� Is Not Null And ��ֹԭ�� = 1;
  
    Select ��ʼԭ��
    Into v_��ʼԭ��
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� Is Null And Nvl(���Ӵ�λ, 0) = 0;
  
    Update ������ҳ
    Set ״̬ = Decode(v_��ʼԭ��, 10, 3, ״̬), ��Ժ���� = Null, ��Ժ��ʽ = Null, �����־ = Null, �������� = Null
    Where ����id = ����id_In And ��ҳid = ��ҳid_In;
    --����λ
    If ����_In Is Null Then
      --ԭ��λû�б�ռ��,��ͥ����Ҳ���ᱻռ��(���������жϱ�ռ�����,ռ�ûᴫ�봲��_In)
      Close c_Curlog;
      For r_Curlogrow In c_Curlog Loop
        If r_Curlogrow.���� Is Not Null Then
          --��鴲λ
          Select Count(*)
          Into v_Count
          From ��λ״����¼
          Where ����id = r_Curlogrow.����id And ���� = r_Curlogrow.���� And ״̬ = '�մ�';
          If v_Count = 0 Then
            v_Error := '[ZLSOFT]����ʧ��,��λ ' || r_Curlogrow.���� || ' ���ǿմ���[ZLSOFT]';
            Raise Err_Custom;
          End If;
          --����ռ�ô�λ
          Update ��λ״����¼
          Set ״̬ = 'ռ��', ����id = ����id_In, �ȼ�id = r_Curlogrow.��λ�ȼ�id, ����id = r_Curlogrow.����id --ǿ�лָ���ǰ�Ŀ���,���ô�Ҳ���ô����ˡ�
          Where ����id = r_Curlogrow.����id And ���� = r_Curlogrow.����;
        End If;
      
        If Nvl(r_Curlogrow.���Ӵ�λ, 0) = 0 Then
          Update ������Ϣ
          Set ��Ժʱ�� = Null, ��ǰ����id = r_Curlogrow.����id, ��ǰ����id = r_Curlogrow.����id, ��ǰ���� = r_Curlogrow.����, ��Ժ = 1
          Where ����id = ����id_In;
        
          --������Ժ����
          Begin
            Update ��Ժ����
            Set ����id = Nvl(r_Curlogrow.����id, 0), ����id = r_Curlogrow.����id
            Where ����id = ����id_In;
            If Sql%Rowcount = 0 Then
              Insert Into ��Ժ����
                (����id, ����id, ����id)
              Values
                (����id_In, r_Curlogrow.����id, Nvl(r_Curlogrow.����id, 0));
            End If;
          Exception
            When Others Then
              Null;
          End;
        
        End If;
      End Loop;
    Else
      --ԭ��λ��ռ�ã������°��ŵĴ�λ,��סһ�Ż���Ų�������
      v_���Ŵ� := ����_In || ',';
      --������˳�Ժǰ״̬ΪԤ��Ժ������Ԥ��Ժ
      If v_��ʼԭ�� = 10 Then
        --����Ԥ��Ժ
        Update ������ҳ Set ״̬ = 0 Where ����id = ����id_In And ��ҳid = ��ҳid_In;
        --�ָ��䶯
        Delete From ���˱䶯��¼
        Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ʼԭ�� = 10 And ��ֹʱ�� Is Null;
      
        Update ���˱䶯��¼
        Set ��ֹʱ�� = v_��ʼʱ��, ��ֹԭ�� = 4, ��ֹ��Ա = ����Ա����_In
        Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹԭ�� = 10 And ��ֹʱ�� = r_Curlogrow.��ʼʱ��;
      Else
        Update ���˱䶯��¼
        Set ��ֹʱ�� = v_��ʼʱ��, ��ֹԭ�� = 4, ��ֹ��Ա = ����Ա����_In
        Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹʱ�� Is Null;
      End If;
    
      While v_���Ŵ� Is Not Null Loop
        v_���� := Substr(v_���Ŵ�, 1, Instr(v_���Ŵ�, ',') - 1);
        --ԭʼ��λ�ȼ����´�λ�ȼ��������ڳ������ж�
        --��鴲λ
        Select Count(*)
        Into v_Count
        From ��λ״����¼
        Where ����id = r_Curlogrow.����id And ���� = v_���� And ״̬ = '�մ�';
        If v_Count = 0 Then
          v_Error := '����ʧ��,��λ ' || v_���� || ' ���ǿմ���';
          Close c_Curlog;
          Raise Err_Custom;
        End If;
        --���´�λ״����¼
        Update ��λ״����¼
        Set ״̬ = 'ռ��', ����id = ����id_In, ����id = r_Curlogrow.����id
        Where ����id = r_Curlogrow.����id And ���� = v_����;
      
        Select �ȼ�id Into v_��λ�ȼ�id From ��λ״����¼ Where ����id = r_Curlogrow.����id And ���� = v_����;
        ----����ԭ��Ϊ4
        Insert Into ���˱䶯��¼
          (Id, ����id, ��ҳid, ��ʼʱ��, ��ʼԭ��, ���Ӵ�λ, ����id, ����id, ҽ��С��id, ����ȼ�id, ��λ�ȼ�id, ����, ���λ�ʿ, ����ҽʦ, ����ҽʦ, ����ҽʦ, ����, ����Ա���,
           ����Ա����)
        Values
          (���˱䶯��¼_Id.Nextval, ����id_In, ��ҳid_In, v_��ʼʱ��, 4, Decode(����λ_In, v_����, 0, 1), r_Curlogrow.����id,
           r_Curlogrow.����id, r_Curlogrow.ҽ��С��id, r_Curlogrow.����ȼ�id, v_��λ�ȼ�id, v_����, r_Curlogrow.���λ�ʿ, r_Curlogrow.����ҽʦ,
           r_Curlogrow.����ҽʦ, r_Curlogrow.����ҽʦ, r_Curlogrow.����, ����Ա���_In, ����Ա����_In);
      
        v_���Ŵ� := Substr(v_���Ŵ�, Instr(v_���Ŵ�, ',') + 1);
      End Loop;
      --���²�����Ϣ
      Update ������Ϣ
      Set ��Ժʱ�� = Null, ��ǰ����id = r_Curlogrow.����id, ��ǰ����id = r_Curlogrow.����id, ��ǰ���� = ����λ_In, ��Ժ = 1
      Where ����id = ����id_In;
    
      --������Ժ����
      Begin
        Update ��Ժ���� Set ����id = Nvl(r_Curlogrow.����id, 0), ����id = r_Curlogrow.����id Where ����id = ����id_In;
        If Sql%Rowcount = 0 Then
          Insert Into ��Ժ����
            (����id, ����id, ����id)
          Values
            (����id_In, r_Curlogrow.����id, Nvl(r_Curlogrow.����id, 0));
        End If;
      Exception
        When Others Then
          Null;
      End;
    
      --���²�����ҳ��Ժ����
      Update ������ҳ Set ��Ժ���� = ����λ_In Where ����id = ����id_In And ��ҳid = ��ҳid_In;
      Close c_Curlog;
    End If;
    --ɾ����Ժ��� �����������Ϣ
    --Delete From ������ϼ�¼ Where ����id = ����id_In And ��ҳid = ��ҳid_In And ������� in (3,13) And ��¼��Դ = 2;
  
    Begin
      Select ����� Into v_����� From Zlsystems Where Floor(��� / 100) = 3;
    Exception
      When Others Then
        Null;
    End;
    --ɾ���ò��˵������¼
    If v_����� = 100 Then
      Execute Immediate 'Delete From �����¼ Where ����id =:1 And ��ҳid =:2'
        Using ����id_In, ��ҳid_In;
    End If;
  Elsif r_Curlogrow.��ʼԭ�� = 1 And v_������ʽ = '��Ժ��ס' Then
    --�������(��Ժͬʱ���)
    v_��ʼʱ�� := r_Curlogrow.��ʼʱ��;
    Select Zl_סԺ�ձ�_Count(r_Curlogrow.����id, v_��ʼʱ��) Into v_Count From Dual;
    If v_Count > 0 Then
      v_Error := '[ZLSOFT]�Ѳ���ҵ��ʱ���ڵ�סԺ�ձ�,���ܰ����ҵ��![ZLSOFT]';
      Raise Err_Custom;
    End If;
  
    Select Count(1)
    Into v_Count
    From ���˻����ļ� a, ���˻������� b
    Where a.Id = b.�ļ�id And ����id = ����id_In And ��ҳid = ��ҳid_In And ����id = r_Curlogrow.����id And Rownum < 2;
    If v_Count > 0 Then
      v_Error := '[ZLSOFT]�����Ѳ��������ݵĻ����ļ������ܰ����ҵ��[ZLSOFT]';
      Raise Err_Custom;
    End If;
    Delete From ���˻����ļ� Where ����id = ����id_In And ��ҳid = ��ҳid_In And ����id = r_Curlogrow.����id;
  
    Close c_Curlog;
  
    If ����_In = '1' Then
      For r_Fee In (Select No
                    From סԺ���ü�¼
                    Where ����id = ����id_In And ��ҳid = ��ҳid_In And Mod(��¼����, 10) = 3 And �Ǽ�ʱ�� >= v_��ʼʱ��
                    Group By No, ���, Mod(��¼����, 10)
                    Having Sum(���ʽ��) <> 0) Loop
        v_Error := '[ZLSOFT]�ò��˵��Զ����ʷ����ѽ���,���ܽ��г���������[ZLSOFT]';
        Raise Err_Custom;
      End Loop;
    End If;
    --�˳���ǰ��λ
    For r_Curlogrow In c_Curlog Loop
      If r_Curlogrow.���� Is Not Null Then
        Update ��λ״����¼
        Set ״̬ = '�մ�', ����id = Null, ����id = Decode(����, 1, Null, ����id)
        Where ����id = r_Curlogrow.����id And ���� = r_Curlogrow.����;
      End If;
    End Loop;
    --�����Ϣ��ԭ
    Update ������ҳ Set ��Ժ���� = Null, ��Ժ���� = Null, ״̬ = 1 Where ����id = ����id_In And ��ҳid = ��ҳid_In;
  
    Update ������Ϣ Set ��ǰ���� = Null Where ����id = ����id_In;
  
    --�ָ��䶯(��Ժͬʱ��Ʋ����а���)
    --��Ϊ��ͬһ����¼�еĳ���,���Բ�������Ա
    Update ���˱䶯��¼
    Set ��λ�ȼ�id = Null, ���� = Null, ���λ�ʿ = Null, ����ҽʦ = Null, �ϴμ���ʱ�� = Null
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ʼԭ�� = 1 And ��ֹʱ�� Is Null;
  Elsif r_Curlogrow.��ʼԭ�� = 2 And v_������ʽ = '��ס' Then
    --������Ժ���
    v_��ʼʱ�� := r_Curlogrow.��ʼʱ��;
    Select Zl_סԺ�ձ�_Count(r_Curlogrow.����id, v_��ʼʱ��) Into v_Count From Dual;
    If v_Count > 0 Then
      v_Error := '[ZLSOFT]�Ѳ���ҵ��ʱ���ڵ�סԺ�ձ�,���ܰ����ҵ��![ZLSOFT]';
      Raise Err_Custom;
    End If;
  
    Select Count(1)
    Into v_Count
    From ���˻����ļ� a, ���˻������� b
    Where a.Id = b.�ļ�id And ����id = ����id_In And ��ҳid = ��ҳid_In And ����id = r_Curlogrow.����id And Rownum < 2;
    If v_Count > 0 Then
      v_Error := '[ZLSOFT]�����Ѳ��������ݵĻ����ļ������ܰ����ҵ��[ZLSOFT]';
      Raise Err_Custom;
    End If;
    Delete From ���˻����ļ� Where ����id = ����id_In And ��ҳid = ��ҳid_In And ����id = r_Curlogrow.����id;
  
    If ����_In = '1' Then
      For r_Fee In (Select No
                    From סԺ���ü�¼
                    Where ����id = ����id_In And ��ҳid = ��ҳid_In And Mod(��¼����, 10) = 3 And �Ǽ�ʱ�� >= v_��ʼʱ��
                    Group By No, ���, Mod(��¼����, 10)
                    Having Sum(���ʽ��) <> 0) Loop
        v_Error := '[ZLSOFT]�ò��˵��Զ����ʷ����ѽ���,���ܽ��г���������[ZLSOFT]';
        Raise Err_Custom;
      End Loop;
    End If;
    --ɾ��������дʱ��
    Zl_���Ӳ���ʱ��_Delete(����id_In, ��ҳid_In, '��Ժ', r_Curlogrow.����id);
    Close c_Curlog;
    --�˳���ǰ��λ
    For r_Curlogrow In c_Curlog Loop
      If r_Curlogrow.���� Is Not Null Then
        Update ��λ״����¼
        Set ״̬ = '�մ�', ����id = Null, ����id = Decode(����, 1, Null, ����id)
        Where ����id = r_Curlogrow.����id And ���� = r_Curlogrow.����;
      End If;
    End Loop;
    --�����Ϣ��ԭ
    Open c_Prelog(v_��ʼʱ��, 2);
    Fetch c_Prelog
      Into r_Prelogrow;
    Update ������ҳ
    Set ��Ժ���� = Null, ��Ժ���� = Null, ״̬ = 1, ��ǰ���� = r_Prelogrow.����, ��Ժ���� = r_Prelogrow.����, ҽ��С��id = r_Prelogrow.ҽ��С��id,
        ����ȼ�id = r_Prelogrow.����ȼ�id
    Where ����id = ����id_In And ��ҳid = ��ҳid_In;
    Close c_Prelog;
  
    Update ������Ϣ Set ��ǰ���� = Null Where ����id = ����id_In;
    Delete ������ҳ�ӱ�
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And (��Ϣ�� = '����ҽʦ' Or ��Ϣ�� = '����ҽʦ');
    --�ָ��䶯
    Delete From ���˱䶯��¼ Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ʼԭ�� = 2 And ��ֹʱ�� Is Null;
  
    Update ���˱䶯��¼
    Set ��ֹʱ�� = Null, ��ֹԭ�� = Null, ��ֹ��Ա = Null, �ϴμ���ʱ�� = Null
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹԭ�� = 2 And ��ֹʱ�� = v_��ʼʱ��;
  Elsif r_Curlogrow.��ʼԭ�� = 3 And v_������ʽ = 'ת����ס' Then
    --����ת�����
    v_��ʼʱ�� := r_Curlogrow.��ʼʱ��;
    Select Zl_סԺ�ձ�_Count(r_Curlogrow.����id, v_��ʼʱ��) Into v_Count From Dual;
    If v_Count > 0 Then
      v_Error := '[ZLSOFT]�Ѳ���ҵ��ʱ���ڵ�סԺ�ձ�,���ܰ����ҵ��![ZLSOFT]';
      Raise Err_Custom;
    End If;
  
    Select Count(1)
    Into v_Count
    From ���˻����ļ�
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ����id = r_Curlogrow.����id And ����ʱ�� >= r_Curlogrow.��ʼʱ��;
    If v_Count > 0 Then
      v_Error := '[ZLSOFT]�����Ѳ��������ļ������ܰ����ҵ��[ZLSOFT]';
      Raise Err_Custom;
    End If;
  
    If ����_In = '1' Then
      For r_Fee In (Select No
                    From סԺ���ü�¼
                    Where ����id = ����id_In And ��ҳid = ��ҳid_In And Mod(��¼����, 10) = 3 And �Ǽ�ʱ�� >= v_��ʼʱ��
                    Group By No, ���, Mod(��¼����, 10)
                    Having Sum(���ʽ��) <> 0) Loop
        v_Error := '[ZLSOFT]�ò��˵��Զ����ʷ����ѽ���,���ܽ��г���������[ZLSOFT]';
        Raise Err_Custom;
      End Loop;
    End If;
  
    --ɾ��������дʱ��
    Zl_���Ӳ���ʱ��_Delete(����id_In, ��ҳid_In, 'ת��', r_Curlogrow.����id);
    Close c_Curlog;
    --�˳���ǰ��λ
    For r_Curlogrow In c_Curlog Loop
      If r_Curlogrow.���� Is Not Null Then
        Update ��λ״����¼
        Set ״̬ = '�մ�', ����id = Null, ����id = Decode(����, 1, Null, ����id)
        Where ����id = r_Curlogrow.����id And ���� = r_Curlogrow.����;
      End If;
    End Loop;
    --��鼰��ԭԭ��λ
    For r_Prelogrow In c_Prelog(v_��ʼʱ��, 3) Loop
      d_��ʼʱ�� := r_Prelogrow.��ʼʱ��;
      If r_Prelogrow.���� Is Not Null Then
        Select Count(*)
        Into v_Count
        From ��λ״����¼
        Where ����id = r_Prelogrow.����id And ���� = r_Prelogrow.���� And ״̬ = '�մ�';
        If v_Count = 0 Then
          v_Error := '[ZLSOFT]����ת��ÿ���ǰ�Ĵ�λ ' || r_Prelogrow.���� || ' ��ǰ�ǿմ����Ѿ�������[ZLSOFT]';
          Raise Err_Custom;
        End If;
      
        Update ��λ״����¼
        Set ״̬ = 'ռ��', ����id = ����id_In, �ȼ�id = r_Prelogrow.��λ�ȼ�id, ����id = r_Prelogrow.����id --ǿ�лָ���ǰ�Ŀ���,���ô�Ҳ���ô����ˡ�
        Where ����id = r_Prelogrow.����id And ���� = r_Prelogrow.����;
      End If;
      --�����Ϣ��ԭ
      If Nvl(r_Prelogrow.���Ӵ�λ, 0) = 0 Then
        --�ж��Ƿ���Ӥ��
        Select Count(1) Into v_Count From ������������¼ Where ����id = ����id_In And ��ҳid = ��ҳid_In;
        If v_Count > 0 Then
          Select ĸӤת�Ʊ�־ Into v_ĸӤת�Ʊ�־ From ������ҳ Where ����id = ����id_In And ��ҳid = ��ҳid_In;
          If v_ĸӤת�Ʊ�־ Is Not Null Then
            If Substr(v_ĸӤת�Ʊ�־, Length(v_ĸӤת�Ʊ�־)) = '1' Then
              --�����1��ʾĸ�׺�Ӥ��δ�ֿ�������ա�������ҳ.Ӥ������ID���͡�������ҳ.Ӥ������ID����
              Update ������ҳ Set Ӥ������id = Null, Ӥ������id = Null Where ����id = ����id_In And ��ҳid = ��ҳid_In;
              If Length(v_ĸӤת�Ʊ�־) > 1 Then
                v_ĸӤת�Ʊ�־ := Substr(v_ĸӤת�Ʊ�־, 1, Length(v_ĸӤת�Ʊ�־) - 1);
              Else
                v_ĸӤת�Ʊ�־ := '';
              End If;
            Else
              --�����0����ʾ���ϴ�ת����ĸ�׵���ת�ߵģ������½�ԭ���ƿ��ҺͲ�����д����������ҳ.Ӥ������ID���͡�������ҳ.Ӥ������ID�����Ӳ��˱䶯��¼��ȡ������������һλ��ʶ
              If Length(v_ĸӤת�Ʊ�־) > 1 Then
                v_ĸӤת�Ʊ�־ := Substr(v_ĸӤת�Ʊ�־, 1, Length(v_ĸӤת�Ʊ�־) - 1);
                --�鿴��һ��ת�Ƶı�ʶ
                If Substr(v_ĸӤת�Ʊ�־, Length(v_ĸӤת�Ʊ�־)) = '1' Then
                  Update ������ҳ
                  Set Ӥ������id = Null, Ӥ������id = Null
                  Where ����id = ����id_In And ��ҳid = ��ҳid_In;
                Else
                  --ȡ��ת��ǰ��Ӥ�����Ҳ���ID
                  v_Tmp   := v_ĸӤת�Ʊ�־;
                  v_Count := 1;
                  While v_Tmp Is Not Null Loop
                    v_Count := v_Count + 1;
                    If Substr(v_Tmp, Length(v_Tmp)) = '1' Then
                      Select Max(a.����id) As ����id, Max(a.����id) As ����id
                      Into n_ԭ����id, n_ԭ����id
                      From (Select ����id, ����id, Rownum As ���
                             From (Select ����id, ����id
                                    From ���˱䶯��¼
                                    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ʼԭ�� = 3 And ���Ӵ�λ = 0
                                    Order By ��ʼʱ�� Desc)) a
                      Where ��� = v_Count;
                    End If;
                    If Length(v_Tmp) = 1 Then
                      v_Tmp := '';
                    Else
                      v_Tmp := Substr(v_Tmp, 1, Length(v_Tmp) - 1);
                    End If;
                  End Loop;
                  If Nvl(n_ԭ����id, 0) = 0 Then
                    --���û���ҵ�����ȡ��Ժ����
                    Select Max(b.����id), Max(b.����id)
                    Into n_ԭ����id, n_ԭ����id
                    From ���˱䶯��¼ b
                    Where b.����id = ����id_In And b.��ҳid = ��ҳid_In And b.����id Is Not Null And b.����id Is Not Null And
                          b.��ʼʱ�� = (Select Min(a.��ʼʱ��)
                                    From ���˱䶯��¼ a
                                    Where a.����id = b.����id And a.��ҳid = b.��ҳid And a.����id Is Not Null And
                                          a.����id Is Not Null And a.���Ӵ�λ = 0);
                  End If;
                
                  If Nvl(n_ԭ����id, 0) <> 0 Then
                    Update ������ҳ
                    Set Ӥ������id = n_ԭ����id, Ӥ������id = n_ԭ����id
                    Where ����id = ����id_In And ��ҳid = ��ҳid_In;
                  Else
                    --֮ǰû��ת�Ƽ�¼,���Ӥ������ID��Ӥ������ID
                    Update ������ҳ
                    Set Ӥ������id = Null, Ӥ������id = Null
                    Where ����id = ����id_In And ��ҳid = ��ҳid_In;
                  End If;
                End If;
              Else
                --ֻ����һ��ת��,���˺����Ӥ�����Ҳ���ID
                v_ĸӤת�Ʊ�־ := '';
                Update ������ҳ
                Set Ӥ������id = Null, Ӥ������id = Null
                Where ����id = ����id_In And ��ҳid = ��ҳid_In;
              End If;
            End If;
            --ȥ�����һλ��ʶ
            Update ������ҳ Set ĸӤת�Ʊ�־ = v_ĸӤת�Ʊ�־ Where ����id = ����id_In And ��ҳid = ��ҳid_In;
          End If;
        
        End If;
      
        Update ������ҳ
        Set ״̬ = 2, ��ǰ����id = r_Prelogrow.����id, ��Ժ����id = r_Prelogrow.����id, ҽ��С��id = r_Prelogrow.ҽ��С��id,
            ��Ժ���� = r_Prelogrow.����, ����ȼ�id = r_Prelogrow.����ȼ�id, ���λ�ʿ = r_Prelogrow.���λ�ʿ, סԺҽʦ = r_Prelogrow.����ҽʦ,
            ��ǰ���� = r_Curlogrow.����
        Where ����id = ����id_In And ��ҳid = ��ҳid_In;
      
        Update ������ҳ�ӱ�
        Set ��Ϣֵ = r_Prelogrow.����ҽʦ
        Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��Ϣ�� = '����ҽʦ';
        Update ������ҳ�ӱ�
        Set ��Ϣֵ = r_Prelogrow.����ҽʦ
        Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��Ϣ�� = '����ҽʦ';
      
        Update ������Ϣ
        Set ��ǰ����id = r_Prelogrow.����id, ��ǰ����id = r_Prelogrow.����id, ��ǰ���� = r_Prelogrow.����
        Where ����id = ����id_In;
      
        --������Ժ����
        Update ��Ժ���� Set ����id = Nvl(r_Prelogrow.����id, 0), ����id = r_Prelogrow.����id Where ����id = ����id_In;
      
      End If;
    End Loop;
    --�ָ��䶯(�ָ�����ʱת�Ʊ��״̬)
    Delete From ���˱䶯��¼
    Where ���Ӵ�λ = 1 And ����id = ����id_In And ��ҳid = ��ҳid_In And ��ʼԭ�� = 3 And ��ֹʱ�� Is Null;
  
    Select ��ֹ��Ա
    Into v_��ֹ��Ա
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹԭ�� = 3 And ��ֹʱ�� = v_��ʼʱ�� And Nvl(���Ӵ�λ, 0) = 0;
    --��ʱ��¼�Ĳ���Ա��Ϣ��¼������ֹ��Ա,��Ϊû�м�¼��ֹ��Ա���,�Ͳ��ָ�
    Update ���˱䶯��¼
    Set ��ʼʱ�� = Null, ҽ��С��id = Null, ����ȼ�id = Null, ��λ�ȼ�id = Null, ���� = Null, ���λ�ʿ = Null, ����ҽʦ = Null, ����Ա��� = Null,
        ����Ա���� = v_��ֹ��Ա, �ϴμ���ʱ�� = Null, ����id = Null
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ʼԭ�� = 3 And ��ֹʱ�� Is Null;
  
    Update ���˱䶯��¼
    Set ��ֹʱ�� = Null, ��ֹԭ�� = Null, ��ֹ��Ա = Null, �ϴμ���ʱ�� = Null
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹԭ�� = 3 And ��ֹʱ�� = v_��ʼʱ�� And ��ʼʱ�� = d_��ʼʱ��;
  
  Elsif r_Curlogrow.��ʼԭ�� = 4 And v_������ʽ = '����' Then
    --��������
    v_��ʼʱ�� := r_Curlogrow.��ʼʱ��;
    Close c_Curlog;
    If ����_In = '1' Then
      For r_Fee In (Select No
                    From סԺ���ü�¼
                    Where ����id = ����id_In And ��ҳid = ��ҳid_In And Mod(��¼����, 10) = 3 And �Ǽ�ʱ�� >= v_��ʼʱ��
                    Group By No, ���, Mod(��¼����, 10)
                    Having Sum(���ʽ��) <> 0) Loop
        v_Error := '[ZLSOFT]�ò��˵��Զ����ʷ����ѽ���,���ܽ��г���������[ZLSOFT]';
        Raise Err_Custom;
      End Loop;
    End If;
    --�˳���ǰ��λ
    For r_Curlogrow In c_Curlog Loop
      If r_Curlogrow.���� Is Not Null Then
        Update ��λ״����¼
        Set ״̬ = '�մ�', ����id = Null, ����id = Decode(����, 1, Null, ����id)
        Where ����id = r_Curlogrow.����id And ���� = r_Curlogrow.����;
      End If;
    End Loop;
    --��鼰��ԭԭ��λ
    For r_Prelogrow In c_Prelog(v_��ʼʱ��, 4) Loop
      d_��ʼʱ�� := r_Prelogrow.��ʼʱ��;
      If r_Prelogrow.���� Is Not Null Then
        Select Count(*)
        Into v_Count
        From ��λ״����¼
        Where ����id = r_Prelogrow.����id And ���� = r_Prelogrow.���� And ״̬ = '�մ�';
        If v_Count = 0 Then
          v_Error := '[ZLSOFT]�������һ�λ���ǰ����ס�Ĵ�λ ' || r_Prelogrow.���� || ' ��ǰ�ǿմ����Ѿ�������[ZLSOFT]';
          Raise Err_Custom;
        End If;
      
        Update ��λ״����¼
        Set ״̬ = 'ռ��', ����id = ����id_In, �ȼ�id = r_Prelogrow.��λ�ȼ�id, ����id = Decode(����, 1, r_Prelogrow.����id, ����id)
        Where ����id = r_Prelogrow.����id And ���� = r_Prelogrow.����;
      End If;
      --������Ϣ��������ҳ,������������Ҷ���ʱ,�����ſ��Ի�����,�˴�Ϊ���ж�,ͳһ��ԭ����
      If Nvl(r_Prelogrow.���Ӵ�λ, 0) = 0 Then
        Update ������ҳ
        Set ��Ժ���� = r_Prelogrow.����, ��ǰ����id = r_Prelogrow.����id
        Where ����id = ����id_In And ��ҳid = ��ҳid_In;
      
        Update ������Ϣ Set ��ǰ���� = r_Prelogrow.����, ��ǰ����id = r_Prelogrow.����id Where ����id = ����id_In;
        --������Ժ����
        Update ��Ժ���� Set ����id = Nvl(r_Prelogrow.����id, 0) Where ����id = ����id_In;
      End If;
    End Loop;
    --�ָ��䶯
    Delete From ���˱䶯��¼ Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ʼԭ�� = 4 And ��ֹʱ�� Is Null;
  
    Update ���˱䶯��¼
    Set ��ֹʱ�� = Null, ��ֹԭ�� = Null, ��ֹ��Ա = Null, �ϴμ���ʱ�� = Null
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹԭ�� = 4 And ��ֹʱ�� = v_��ʼʱ�� And ��ʼʱ�� = d_��ʼʱ��;
  Elsif r_Curlogrow.��ʼԭ�� = 4 And v_������ʽ = '��λ�Ի�' Then
    --������λ�Ի�
    v_��ʼʱ�� := r_Curlogrow.��ʼʱ��;
    Select ���� Into v_���� From ������Ϣ Where ����id = ����id_In;
    Close c_Curlog;
    If ����_In = '1' Then
      For r_Fee In (Select No
                    From סԺ���ü�¼
                    Where ����id = ����id_In And ��ҳid = ��ҳid_In And Mod(��¼����, 10) = 3 And �Ǽ�ʱ�� >= v_��ʼʱ��
                    Group By No, ���, Mod(��¼����, 10)
                    Having Sum(���ʽ��) <> 0) Loop
        v_Error := '[ZLSOFT]���� ' || v_���� || ' ���Զ����ʷ����ѽ���,���ܽ��г���������[ZLSOFT]';
        Raise Err_Custom;
      End Loop;
    End If;
    --�˳���ǰ��λ
    --��鼰��ԭԭ��λ
    For r_Prelogrow In c_Prelog(v_��ʼʱ��, 4) Loop
      d_��ʼʱ�� := r_Prelogrow.��ʼʱ��;
      If r_Prelogrow.���� Is Not Null Then
        Update ��λ״����¼
        Set ״̬ = 'ռ��', ����id = ����id_In, �ȼ�id = r_Prelogrow.��λ�ȼ�id, ����id = Decode(����, 1, r_Prelogrow.����id, ����id)
        Where ����id = r_Prelogrow.����id And ���� = r_Prelogrow.����;
      End If;
      --������Ϣ��������ҳ,������������Ҷ���ʱ,�����ſ��Ի�����,�˴�Ϊ���ж�,ͳһ��ԭ����
      If Nvl(r_Prelogrow.���Ӵ�λ, 0) = 0 Then
        Update ������ҳ
        Set ��Ժ���� = r_Prelogrow.����, ��ǰ����id = r_Prelogrow.����id
        Where ����id = ����id_In And ��ҳid = ��ҳid_In;
      
        Update ������Ϣ Set ��ǰ���� = r_Prelogrow.����, ��ǰ����id = r_Prelogrow.����id Where ����id = ����id_In;
        --������Ժ����
        Update ��Ժ���� Set ����id = Nvl(r_Prelogrow.����id, 0) Where ����id = ����id_In;
      End If;
    End Loop;
    --�ָ��䶯
    Delete From ���˱䶯��¼ Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ʼԭ�� = 4 And ��ֹʱ�� Is Null;
  
    Update ���˱䶯��¼
    Set ��ֹʱ�� = Null, ��ֹԭ�� = Null, ��ֹ��Ա = Null, �ϴμ���ʱ�� = Null
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹԭ�� = 4 And ��ֹʱ�� = v_��ʼʱ�� And ��ʼʱ�� = d_��ʼʱ��;
  Elsif r_Curlogrow.��ʼԭ�� = 5 And v_������ʽ = '��λ�ȼ��䶯' Then
    --������λ�ȼ��䶯
    If ����_In = '1' Then
      For r_Fee In (Select No
                    From סԺ���ü�¼
                    Where ����id = ����id_In And ��ҳid = ��ҳid_In And Mod(��¼����, 10) = 3 And �շ�ϸĿid = r_Curlogrow.��λ�ȼ�id And
                          �Ǽ�ʱ�� >= r_Curlogrow.��ʼʱ��
                    Group By No, ���, Mod(��¼����, 10)
                    Having Sum(���ʽ��) <> 0) Loop
        Close c_Curlog;
        v_Error := '[ZLSOFT]�ò��˵��Զ����ʷ����ѽ���,���ܽ��г���������[ZLSOFT]';
        Raise Err_Custom;
      End Loop;
    End If;
    --��ԭԭ��λ�ĵȼ�
    For r_Prelogrow In c_Prelog(r_Curlogrow.��ʼʱ��, 5) Loop
      d_��ʼʱ�� := r_Prelogrow.��ʼʱ��;
      If r_Prelogrow.���� Is Not Null Then
        Update ��λ״����¼
        Set �ȼ�id = r_Prelogrow.��λ�ȼ�id
        Where ����id = r_Prelogrow.����id And ���� = r_Prelogrow.����;
      End If;
    End Loop;
    --�ָ��䶯
    Delete From ���˱䶯��¼ Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ʼԭ�� = 5 And ��ֹʱ�� Is Null;
  
    Update ���˱䶯��¼
    Set ��ֹʱ�� = Null, ��ֹԭ�� = Null, ��ֹ��Ա = Null, �ϴμ���ʱ�� = Null
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹԭ�� = 5 And ��ֹʱ�� = r_Curlogrow.��ʼʱ�� And ��ʼʱ�� = d_��ʼʱ��;
  
    Close c_Curlog;
  Elsif r_Curlogrow.��ʼԭ�� = 6 And v_������ʽ = '����ȼ��䶯' Then
    --��������ȼ��䶯
    If ����_In = '1' Then
      For r_Fee In (Select No
                    From סԺ���ü�¼
                    Where ����id = ����id_In And ��ҳid = ��ҳid_In And Mod(��¼����, 10) = 3 And �շ�ϸĿid = r_Curlogrow.����ȼ�id And
                          �Ǽ�ʱ�� >= r_Curlogrow.��ʼʱ��
                    Group By No, ���, Mod(��¼����, 10)
                    Having Sum(���ʽ��) <> 0) Loop
        Close c_Curlog;
        v_Error := '[ZLSOFT]�ò��˵��Զ����ʷ����ѽ���,���ܽ��г���������[ZLSOFT]';
        Raise Err_Custom;
      End Loop;
    End If;
  
    Open c_Prelog(r_Curlogrow.��ʼʱ��, 6);
    Fetch c_Prelog
      Into r_Prelogrow;
    --�ָ�ԭ����ȼ�
    Update ������ҳ Set ����ȼ�id = r_Prelogrow.����ȼ�id Where ����id = ����id_In And ��ҳid = ��ҳid_In;
    --�ָ��䶯
    Delete From ���˱䶯��¼ Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ʼԭ�� = 6 And ��ֹʱ�� Is Null;
  
    --ҽ�������Ļ���ȼ��䶯û�м�¼�룬����ǰһ�ȼ���ֹͣʱ���뵱ǰ�ȼ��Ŀ�ʼʱ����ͬһ���ӣ�����Ҫȡmax(id)
    Update ���˱䶯��¼
    Set ��ֹʱ�� = Null, ��ֹԭ�� = Null, ��ֹ��Ա = Null, �ϴμ���ʱ�� = Null
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹԭ�� = 6 And ��ֹʱ�� = r_Curlogrow.��ʼʱ��  And ��ʼʱ�� = r_Prelogrow.��ʼʱ��;
  
    Close c_Prelog;
    Close c_Curlog;
  Elsif r_Curlogrow.��ʼԭ�� = 7 And v_������ʽ = '����ҽʦ�ı�' Then
    --ɾ��������дʱ��
    Zl_���Ӳ���ʱ��_Delete(����id_In, ��ҳid_In, '����', r_Curlogrow.����id);
    --��������ҽʦ�ı�
    Open c_Prelog(r_Curlogrow.��ʼʱ��, 7);
    Fetch c_Prelog
      Into r_Prelogrow;
    --�ָ�ԭҽʦ
    Update ������ҳ Set סԺҽʦ = r_Prelogrow.����ҽʦ Where ����id = ����id_In And ��ҳid = ��ҳid_In;
    --�ָ��䶯
    Delete From ���˱䶯��¼ Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ʼԭ�� = 7 And ��ֹʱ�� Is Null;
  
    Update ���˱䶯��¼
    Set ��ֹʱ�� = Null, ��ֹԭ�� = Null, ��ֹ��Ա = Null, �ϴμ���ʱ�� = Null
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹԭ�� = 7 And ��ֹʱ�� = r_Curlogrow.��ʼʱ��  And ��ʼʱ�� = r_Prelogrow.��ʼʱ��;
  
    Close c_Prelog;
    Close c_Curlog;
  Elsif r_Curlogrow.��ʼԭ�� = 8 And v_������ʽ = '���λ�ʿ�ı�' Then
    --�������λ�ʿ�ı�
    Open c_Prelog(r_Curlogrow.��ʼʱ��, 8);
    Fetch c_Prelog
      Into r_Prelogrow;
    --�ָ�ԭ���λ�ʿ
    Update ������ҳ Set ���λ�ʿ = r_Prelogrow.���λ�ʿ Where ����id = ����id_In And ��ҳid = ��ҳid_In;
    --�ָ��䶯
    Delete From ���˱䶯��¼ Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ʼԭ�� = 8 And ��ֹʱ�� Is Null;
  
    Update ���˱䶯��¼
    Set ��ֹʱ�� = Null, ��ֹԭ�� = Null, ��ֹ��Ա = Null, �ϴμ���ʱ�� = Null
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹԭ�� = 8 And ��ֹʱ�� = r_Curlogrow.��ʼʱ��  And ��ʼʱ�� = r_Prelogrow.��ʼʱ��;
  
    Close c_Prelog;
    Close c_Curlog;
  Elsif r_Curlogrow.��ʼԭ�� = 9 And v_������ʽ = 'תΪסԺ����' Then
    --����תΪסԺ����
    --ɾ��������дʱ��
    Open c_Prelog(r_Curlogrow.��ʼʱ��, 9);
    Fetch c_Prelog
      Into r_Prelogrow;
      
    Zl_���Ӳ���ʱ��_Delete(����id_In, ��ҳid_In, '��Ժ', r_Curlogrow.����id);
    Update ������ҳ Set �������� = 2 Where ����id = ����id_In And ��ҳid = ��ҳid_In;
    --�ָ��䶯
    Delete From ���˱䶯��¼ Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ʼԭ�� = 9 And ��ֹʱ�� Is Null;
  
    Update ���˱䶯��¼
    Set ��ֹʱ�� = Null, ��ֹԭ�� = Null, ��ֹ��Ա = Null, �ϴμ���ʱ�� = Null
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹԭ�� = 9 And ��ֹʱ�� = r_Curlogrow.��ʼʱ�� And ��ʼʱ�� = r_Prelogrow.��ʼʱ��;
  
    If ��ҳid_In = 1 And Nvl(����_In, '0') = '1' Then
      Update ������Ϣ Set סԺ�� = Null Where ����id = ����id_In;
      Update ������ҳ Set סԺ�� = Null Where ����id = ����id_In;
    End If;
  
    Close c_Curlog;
    Close c_Prelog;
  Elsif r_Curlogrow.��ʼԭ�� = 10 And v_������ʽ = 'Ԥ��Ժ' Then
    --����Ԥ��Ժ
    --ɾ��������дʱ��
    Open c_Prelog(r_Curlogrow.��ʼʱ��, 10);
    Fetch c_Prelog
      Into r_Prelogrow;
      
    Zl_���Ӳ���ʱ��_Delete(����id_In, ��ҳid_In, '��Ժ', r_Curlogrow.����id);
    --�ָ�סԺ״̬
    Update ������ҳ Set ״̬ = 0 Where ����id = ����id_In And ��ҳid = ��ҳid_In;
    --�ָ��䶯
    Delete From ���˱䶯��¼ Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ʼԭ�� = 10 And ��ֹʱ�� Is Null;
  
    Update ���˱䶯��¼
    Set ��ֹʱ�� = Null, ��ֹԭ�� = Null, �ϴμ���ʱ�� = Null, ��ֹ��Ա = Null
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹԭ�� = 10 And ��ֹʱ�� = r_Curlogrow.��ʼʱ�� And ��ʼʱ�� = r_Prelogrow.��ʼʱ��;
  
    Close c_Curlog;
    Close c_Prelog;
  Elsif r_Curlogrow.��ʼԭ�� = 11 And v_������ʽ = '����ҽʦ�䶯' Then
    --��������ҽʦ�ı�
    Open c_Prelog(r_Curlogrow.��ʼʱ��, 11);
    Fetch c_Prelog
      Into r_Prelogrow;
    --�ָ�ԭ����ҽʦ
    Update ������ҳ�ӱ�
    Set ��Ϣֵ = r_Prelogrow.����ҽʦ
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��Ϣ�� = '����ҽʦ';
    --�ָ��䶯
    Delete From ���˱䶯��¼ Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ʼԭ�� = 11 And ��ֹʱ�� Is Null;
  
    Update ���˱䶯��¼
    Set ��ֹʱ�� = Null, ��ֹԭ�� = Null, ��ֹ��Ա = Null, �ϴμ���ʱ�� = Null
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹԭ�� = 11 And ��ֹʱ�� = r_Curlogrow.��ʼʱ��  And ��ʼʱ�� = r_Prelogrow.��ʼʱ��;
  
    Close c_Prelog;
    Close c_Curlog;
  Elsif r_Curlogrow.��ʼԭ�� = 12 And v_������ʽ = '����ҽʦ�䶯' Then
    --��������ҽʦ�ı�
    Open c_Prelog(r_Curlogrow.��ʼʱ��, 12);
    Fetch c_Prelog
      Into r_Prelogrow;
    --�ָ�ԭ����ҽʦ
    Update ������ҳ�ӱ�
    Set ��Ϣֵ = r_Prelogrow.����ҽʦ
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��Ϣ�� = '����ҽʦ';
    --�ָ��䶯
    Delete From ���˱䶯��¼ Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ʼԭ�� = 12 And ��ֹʱ�� Is Null;
  
    Update ���˱䶯��¼
    Set ��ֹʱ�� = Null, ��ֹԭ�� = Null, ��ֹ��Ա = Null, �ϴμ���ʱ�� = Null
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹԭ�� = 12 And ��ֹʱ�� = r_Curlogrow.��ʼʱ��  And ��ʼʱ�� = r_Prelogrow.��ʼʱ��;
  
    Close c_Prelog;
    Close c_Curlog;
  Elsif r_Curlogrow.��ʼԭ�� = 13 And v_������ʽ = '�����䶯' Then
    --��������ı�
    Open c_Prelog(r_Curlogrow.��ʼʱ��, 13);
    Fetch c_Prelog
      Into r_Prelogrow;
    --�ָ�ԭ����
    Update ������ҳ Set ��ǰ���� = r_Prelogrow.���� Where ����id = ����id_In And ��ҳid = ��ҳid_In;
    --�ָ��䶯
    Delete From ���˱䶯��¼ Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ʼԭ�� = 13 And ��ֹʱ�� Is Null;
  
    Update ���˱䶯��¼
    Set ��ֹʱ�� = Null, ��ֹԭ�� = Null, ��ֹ��Ա = Null, �ϴμ���ʱ�� = Null
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹԭ�� = 13 And ��ֹʱ�� = r_Curlogrow.��ʼʱ�� And ��ʼʱ�� = r_Prelogrow.��ʼʱ��;
  
    Close c_Prelog;
    Close c_Curlog;
  
  Elsif r_Curlogrow.��ʼԭ�� = 14 And v_������ʽ = 'תҽ��С��' Then
    --ɾ��������дʱ��
    Zl_���Ӳ���ʱ��_Delete(����id_In, ��ҳid_In, '����', r_Curlogrow.����id);
  
    --����ҽ��С��ı�
    Open c_Prelog(r_Curlogrow.��ʼʱ��, 14);
    Fetch c_Prelog
      Into r_Prelogrow;
    --�ָ�ԭҽ��С��
    Update ������ҳ
    Set ҽ��С��id = r_Prelogrow.ҽ��С��id, סԺҽʦ = r_Prelogrow.����ҽʦ
    Where ����id = ����id_In And ��ҳid = ��ҳid_In;
    --�ָ�ԭ����ҽʦ
    Update ������ҳ�ӱ�
    Set ��Ϣֵ = r_Prelogrow.����ҽʦ
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��Ϣ�� = '����ҽʦ';
  
    --�ָ��䶯
    Delete From ���˱䶯��¼ Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ʼԭ�� = 14 And ��ֹʱ�� Is Null;
  
    Update ���˱䶯��¼
    Set ��ֹʱ�� = Null, ��ֹԭ�� = Null, ��ֹ��Ա = Null, �ϴμ���ʱ�� = Null
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹԭ�� = 14 And ��ֹʱ�� = r_Curlogrow.��ʼʱ��  And ��ʼʱ�� = r_Prelogrow.��ʼʱ��;
  
    Close c_Prelog;
    Close c_Curlog;
  Elsif r_Curlogrow.��ʼԭ�� = 15 And v_������ʽ = 'ת������ס' Then
    --�����벡��
    v_��ʼʱ�� := r_Curlogrow.��ʼʱ��;
  
    If ����_In = '1' Then
      For r_Fee In (Select No
                    From סԺ���ü�¼
                    Where ����id = ����id_In And ��ҳid = ��ҳid_In And Mod(��¼����, 10) = 3 And �Ǽ�ʱ�� >= v_��ʼʱ��
                    Group By No, ���, Mod(��¼����, 10)
                    Having Sum(���ʽ��) <> 0) Loop
        v_Error := '[ZLSOFT]�ò��˵��Զ����ʷ����ѽ���,���ܽ��г���������[ZLSOFT]';
        Raise Err_Custom;
      End Loop;
    End If;
  
    Open c_Prelog(v_��ʼʱ��, 15);
    Fetch c_Prelog
      Into r_Prelogrow;
  
    d_��ʼʱ�� := r_Prelogrow.��ʼʱ��;
    --����Ч��ҽ��(δֹͣ�����ϵĳ�����δ���͵�����)��Ϊԭ����ִ�еģ���ִ�п����Զ�����Ϊ�µĲ���
    Update ����ҽ����¼
    Set ִ�п���id = r_Prelogrow.����id
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ִ�п���id = r_Curlogrow.����id And ҽ��״̬ Not In (4, 8, 9) And ����ʱ�� < v_��ʼʱ��;
    --��δ��˵ļ��ʻ��۵���Ϊԭ����ִ�еģ���ִ�п����Զ�����Ϊ�µĲ���
    Update סԺ���ü�¼
    Set ִ�в���id = r_Prelogrow.����id
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ִ�в���id = r_Curlogrow.����id And ��¼״̬ = 0;
    Close c_Prelog;
    Close c_Curlog;
  
    --�˳���ǰ��λ
    For r_Curlogrow In c_Curlog Loop
      If r_Curlogrow.���� Is Not Null Then
        Update ��λ״����¼
        Set ״̬ = '�մ�', ����id = Null, ����id = Decode(����, 1, Null, ����id)
        Where ����id = r_Curlogrow.����id And ���� = r_Curlogrow.����;
      End If;
    End Loop;
    --��鼰��ԭԭ��λ
    For r_Prelogrow In c_Prelog(v_��ʼʱ��, 15) Loop
      If r_Prelogrow.���� Is Not Null Then
        Select Count(*)
        Into v_Count
        From ��λ״����¼
        Where ����id = r_Prelogrow.����id And ���� = r_Prelogrow.���� And ״̬ = '�մ�';
        If v_Count = 0 Then
          v_Error := '[ZLSOFT]����ת��ò���ǰ�Ĵ�λ ' || r_Prelogrow.���� || ' ��ǰ�ǿմ����Ѿ�������[ZLSOFT]';
          Raise Err_Custom;
        End If;
      
        Update ��λ״����¼
        Set ״̬ = 'ռ��', ����id = ����id_In, �ȼ�id = r_Prelogrow.��λ�ȼ�id, ����id = r_Prelogrow.����id --ǿ�лָ���ǰ�Ŀ���,���ô�Ҳ���ô����ˡ�
        Where ����id = r_Prelogrow.����id And ���� = r_Prelogrow.����;
      End If;
      --�����Ϣ��ԭ
      If Nvl(r_Prelogrow.���Ӵ�λ, 0) = 0 Then
        Update ������ҳ
        Set ״̬ = 2, ��ǰ����id = r_Prelogrow.����id, ��Ժ����id = r_Prelogrow.����id, ҽ��С��id = r_Prelogrow.ҽ��С��id,
            ��Ժ���� = r_Prelogrow.����, ����ȼ�id = r_Prelogrow.����ȼ�id, ���λ�ʿ = r_Prelogrow.���λ�ʿ, סԺҽʦ = r_Prelogrow.����ҽʦ,
            ��ǰ���� = r_Curlogrow.����
        Where ����id = ����id_In And ��ҳid = ��ҳid_In;
      
        Update ������Ϣ Set ��ǰ����id = r_Prelogrow.����id, ��ǰ���� = r_Prelogrow.���� Where ����id = ����id_In;
      
        --������Ժ����
        Update ��Ժ���� Set ����id = Nvl(r_Prelogrow.����id, 0) Where ����id = ����id_In;
      
      End If;
    End Loop;
  
    --�ָ��䶯(�ָ�����ʱת�Ʊ��״̬)
    Delete From ���˱䶯��¼
    Where ���Ӵ�λ = 1 And ����id = ����id_In And ��ҳid = ��ҳid_In And ��ʼԭ�� = 15 And ��ֹʱ�� Is Null;
  
    Select ��ֹ��Ա
    Into v_��ֹ��Ա
    From ���˱䶯��¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹԭ�� = 15 And ��ֹʱ�� = v_��ʼʱ�� And Nvl(���Ӵ�λ, 0) = 0;
    --��ʱ��¼�Ĳ���Ա��Ϣ��¼������ֹ��Ա,��Ϊû�м�¼��ֹ��Ա���,�Ͳ��ָ�
    Update ���˱䶯��¼
    Set ��ʼʱ�� = Null, ҽ��С��id = Null, ����ȼ�id = Null, ��λ�ȼ�id = Null, ���� = Null, ���λ�ʿ = Null, ����ҽʦ = Null, ����Ա��� = Null,
        ����Ա���� = v_��ֹ��Ա, �ϴμ���ʱ�� = Null, ���Ӵ�λ = Null, ����ҽʦ = Null, ���� = Null
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ʼԭ�� = 15 And ��ֹʱ�� Is Null;
  
    Update ���˱䶯��¼
    Set ��ֹʱ�� = Null, ��ֹԭ�� = Null, ��ֹ��Ա = Null, �ϴμ���ʱ�� = Null
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ֹԭ�� = 15 And ��ֹʱ�� = v_��ʼʱ��  And ��ʼʱ�� = d_��ʼʱ��;
  Else
    Close c_Curlog;
    v_Error := '[ZLSOFT]��ִ�еĳ���' || v_������ʽ || '�����Ѿ���������ִ��,��ˢ�½��棡[ZLSOFT]';
    Raise Err_Custom;
  End If;
  --�����������
  Select Count(*)
  Into v_Count
  From ���˱䶯��¼
  Where ����id = ����id_In And ��ҳid = ��ҳid_In And Nvl(���Ӵ�λ, 0) = 0 And ��ʼʱ�� Is Not Null And ��ֹʱ�� Is Null;
  If v_Count > 1 Then
    v_Error := '���ֲ��˴��ڷǷ��ı䶯��¼,��ǰ�������ܼ�����' || Chr(13) || Chr(10) || '��������������粢�����������,��ˢ�²���״̬�����ԣ�';
    Raise Err_Custom;
  End If;

  Select Count(*) Into v_Count From ������ҳ Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��Ժ���� Is Null;
  If v_Count = 0 Then
    v_Error := '����ʧ��,�ò����ѳ�Ժ,���ܽ��е�ǰ����' || Chr(13) || Chr(10) || '��������������粢�����������,��ˢ�²���״̬��';
    Raise Err_Custom;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, v_Error);
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_���˱䶯��¼_Undo;
/

--94041:����,2016-03-10,�����ۼ۵��۳ɱ����ֶ�����
--92931:����,2016-01-21,���ϵ�����Ч���������ű��Ͱ�װ�ű���ͬ����
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
        If Nvl(c_����.��д����, 0) = 0 And Nvl(c_����.�����, 0) = 0 And Nvl(c_����.�����, 0) = 0 Then
          Null;
        Elsif Nvl(c_����.��д����, 0) = 0 And (Nvl(c_����.�����, 0) <> 0 Or Nvl(c_����.�����, 0) <> 0) Then
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

--93587:��ҵ��,2016-02-25,����ִ��״̬��ֵ����
Create Or Replace Procedure Zl_ҩƷ�շ���¼_������ҩ
(
  Partid_In        In ҩƷ�շ���¼.�ⷿid%Type,
  Bill_In          In ҩƷ�շ���¼.����%Type,
  No_In            In ҩƷ�շ���¼.No%Type,
  People_In        In ҩƷ�շ���¼.�����%Type,
  ��ҩ��_In        In ҩƷ�շ���¼.��ҩ��%Type := Null,
  У����_In        In ҩƷ�շ���¼.������%Type := Null,
  ��ҩ��ʽ_In      In ҩƷ�շ���¼.��ҩ��ʽ%Type := 1,
  ��ҩʱ��_In      In ҩƷ�շ���¼.�������%Type := Null,
  ����Ա���_In    In ��Ա��.���%Type := Null,
  ����Ա����_In    In ��Ա��.����%Type := Null,
  Intdigit_In      In Number := 2,
  Intautoverify_In In Number := 0,
  ����_In          In Number := 1,
  �˲���_In        In ҩƷ�շ���¼.�˲���%Type := Null,
  δȡҩ_In        In ҩƷ�շ���¼.�Ƿ�δȡҩ%Type := Null
) Is
  --סԺ����
  Cursor c_Modifybillin Is
    Select a.Id, a.ҩƷid, a.������id, a.���ϵ��, a.����id, Nvl(a.ʵ������, 0) * Nvl(a.����, 1) ����, Nvl(a.���۽��, 0) ���, Nvl(a.����, 0) ����,
           a.��ҩ��λid, a.�ɱ���, a.����, a.����, a.Ч��, a.��������, a.��׼�ĺ�, b.����id, b.���, Nvl(c.��������, Nvl(a.ע��֤��, 0)) ��������,
           Nvl(a.���ۼ�, 0) As ���ۼ�
    From ҩƷ�շ���¼ A, סԺ���ü�¼ B, δ��ҩƷ��¼ C
    Where a.���� = c.���� And a.No = c.No And Nvl(a.�ⷿid, 0) = Nvl(c.�ⷿid, 0) And a.No = No_In And a.���� = Bill_In And
          (a.�ⷿid + 0 = Partid_In Or a.�ⷿid Is Null) And Nvl(a.ժҪ, 'С��') <> '�ܷ�' And a.����id = b.Id And Nvl(b.ִ��״̬,0) <> 1 And
          Mod(a.��¼״̬, 3) = 1 And a.����� Is Null;

  --���ﲡ��
  Cursor c_Modifybillout Is
    Select a.Id, a.ҩƷid, a.������id, a.���ϵ��, a.����id, Nvl(a.ʵ������, 0) * Nvl(a.����, 1) ����, Nvl(a.���۽��, 0) ���, Nvl(a.����, 0) ����,
           a.��ҩ��λid, a.�ɱ���, a.����, a.����, a.Ч��, a.��������, a.��׼�ĺ�, b.����id, b.���, Nvl(c.��������, Nvl(a.ע��֤��, 0)) ��������,
           Nvl(a.���ۼ�, 0) As ���ۼ�, b.No, b.��¼����
    From ҩƷ�շ���¼ A, ������ü�¼ B, δ��ҩƷ��¼ C
    Where a.���� = c.���� And a.No = c.No And Nvl(a.�ⷿid, 0) = Nvl(c.�ⷿid, 0) And a.No = No_In And a.���� = Bill_In And
          (a.�ⷿid + 0 = Partid_In Or a.�ⷿid Is Null) And Nvl(a.ժҪ, 'С��') <> '�ܷ�' And a.����id = b.Id And Nvl(b.ִ��״̬,0) <> 1 And
          Mod(a.��¼״̬, 3) = 1 And a.����� Is Null
    Order By ҩƷid;

  v_Modifybillin  c_Modifybillin%RowType;
  v_Modifybillout c_Modifybillout%RowType;

  --ֻ������
  Dbl�����  Number;
  v_�˲����� ҩƷ�շ���¼.�˲�����%Type;
  --��д����
  Dblʵ�ʽ��       ҩƷ�շ���¼.���۽��%Type;
  Dbl�ɱ����       ҩƷ�շ���¼.�ɱ����%Type;
  Dblʵ�ʲ��       ҩƷ�շ���¼.���%Type;
  Date����ʱ��      ҩƷ�շ���¼.�������%Type;
  Bln�շ��뷢ҩ���� Number(1);
  n_ƽ���ɱ���      ҩƷ���.ƽ���ɱ���%Type;
  v_������          ҩƷ�շ���¼.������%Type;
  v_Error           Varchar2(4000);
  Err_Custom Exception;
Begin
  --ȡ��ҩʱ��
  If ��ҩʱ��_In Is Null Then
    Select Sysdate Into Date����ʱ�� From Dual;
  Else
    Date����ʱ�� := ��ҩʱ��_In;
  End If;

  v_�˲����� := Date����ʱ��;
  Begin
    Select 0 Into Bln�շ��뷢ҩ���� From δ��ҩƷ��¼ Where ���� = Bill_In And NO = No_In And �ⷿid + 0 = Partid_In;
  Exception
    When Others Then
      Bln�շ��뷢ҩ���� := 1;
  End;

  --��д�ѷ�ҩ��������ҩ��
  Update ҩƷ�շ���¼
  Set ��ҩ�� = Decode(��ҩ��_In, Null, ��ҩ��, ��ҩ��_In), ��ҩ���� = Decode(��ҩ��_In, Null, ��ҩ����, Date����ʱ��)
  Where NO = No_In And ���� = Bill_In And (�ⷿid + 0 = Partid_In Or �ⷿid Is Null) And Mod(��¼״̬, 3) = 1 And ����� Is Not Null;

  Begin
    Select ������
    Into v_������
    From ҩƷ�շ���¼
    Where NO = No_In And ���� = Bill_In And �ⷿid + 0 = Partid_In And ������� Is Null And Rownum = 1
    For Update Nowait;
  Exception
    When Others Then
      v_Error := '���������û���ִ�з�ҩ�������ظ�������';
      Raise Err_Custom;
  End;

  --���¼���ɱ��ۡ��ɱ������۽����
  If ����_In = 1 Then
    --������������
    For v_Modifybillout In c_Modifybillout Loop
      n_ƽ���ɱ��� := Round(Zl_Fun_Getoutcost(v_Modifybillout.ҩƷid, v_Modifybillout.����, Partid_In), 5);
      Dbl�ɱ����  := Round(n_ƽ���ɱ��� * Nvl(v_Modifybillout.����, 0), Intdigit_In);
      --���۽��
      Dblʵ�ʽ�� := Nvl(v_Modifybillout.���, 0);
      --���
      Dblʵ�ʲ�� := Round(Dblʵ�ʽ�� - Dbl�ɱ����, Intdigit_In);
    
      --����ҩƷ�շ���¼�����۽��ɱ�����ۡ�����˵���Ϣ
      Update ҩƷ�շ���¼
      Set �ɱ��� = n_ƽ���ɱ���, �ɱ���� = Dbl�ɱ����, ��� = Dblʵ�ʲ��, �ⷿid = Partid_In, ��ҩ�� = Decode(��ҩ��_In, Null, ��ҩ��, ��ҩ��_In),
          �˲��� = �˲���_In, �˲����� = v_�˲�����, ��ҩ���� = Decode(��ҩ��_In, Null, ��ҩ����, Date����ʱ��),
          ������ = Decode(У����_In, Null, ������, У����_In), ����� = Decode(People_In, Null, Zl_Username, People_In),
          ������� = Date����ʱ��, ��ҩ��ʽ = ��ҩ��ʽ_In, ע��֤�� = v_Modifybillout.��������, �Ƿ�δȡҩ = δȡҩ_In
      Where ID = v_Modifybillout.Id;
    
      If Bln�շ��뷢ҩ���� = 1 Then
        Update ҩƷ���
        Set �������� = Nvl(��������, 0) - Nvl(v_Modifybillout.����, 0), ʵ������ = Nvl(ʵ������, 0) - Nvl(v_Modifybillout.����, 0),
            ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) - Nvl(v_Modifybillout.���, 0), ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) - Dblʵ�ʲ��
        Where �ⷿid + 0 = Partid_In And ҩƷid = v_Modifybillout.ҩƷid And ���� = 1 And Nvl(����, 0) = v_Modifybillout.����;
      Else
        Update ҩƷ���
        Set ʵ������ = Nvl(ʵ������, 0) - Nvl(v_Modifybillout.����, 0), ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) - Nvl(v_Modifybillout.���, 0),
            ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) - Dblʵ�ʲ��
        Where �ⷿid + 0 = Partid_In And ҩƷid = v_Modifybillout.ҩƷid And ���� = 1 And Nvl(����, 0) = v_Modifybillout.����;
      End If;
    
      If Sql%RowCount = 0 Then
        If Bln�շ��뷢ҩ���� = 1 Then
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, ��������, ʵ������, ʵ�ʽ��, ʵ�ʲ��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴβ���, Ч��, �ϴ���������, ��׼�ĺ�, ƽ���ɱ���)
          Values
            (Partid_In, v_Modifybillout.ҩƷid, v_Modifybillout.����, 1, 0 - Nvl(v_Modifybillout.����, 0),
             0 - Nvl(v_Modifybillout.����, 0), 0 - Nvl(v_Modifybillout.���, 0), 0 - Dblʵ�ʲ��, v_Modifybillout.��ҩ��λid,
             v_Modifybillout.�ɱ���, v_Modifybillout.����, v_Modifybillout.����, v_Modifybillout.Ч��, v_Modifybillout.��������,
             v_Modifybillout.��׼�ĺ�, v_Modifybillout.�ɱ���);
        Else
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, ʵ������, ʵ�ʽ��, ʵ�ʲ��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴβ���, Ч��, �ϴ���������, ��׼�ĺ�, ƽ���ɱ���)
          Values
            (Partid_In, v_Modifybillout.ҩƷid, v_Modifybillout.����, 1, 0 - Nvl(v_Modifybillout.����, 0),
             0 - Nvl(v_Modifybillout.���, 0), 0 - Dblʵ�ʲ��, v_Modifybillout.��ҩ��λid, v_Modifybillout.�ɱ���,
             v_Modifybillout.����, v_Modifybillout.����, v_Modifybillout.Ч��, v_Modifybillout.��������, v_Modifybillout.��׼�ĺ�,
             v_Modifybillout.�ɱ���);
        End If;
      End If;
    
      Delete ҩƷ���
      Where �ⷿid + 0 = Partid_In And ҩƷid = v_Modifybillout.ҩƷid And Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And
            Nvl(ʵ�ʽ��, 0) = 0 And Nvl(ʵ�ʲ��, 0) = 0 And ���� = 1;
    
      --���·��ü�¼��ִ��״̬(��ִ��)
      Update ������ü�¼
      Set ִ��״̬ = 1, ִ���� = Decode(People_In, Null, Zl_Username, People_In), ִ�в���id = Partid_In, ִ��ʱ�� = Date����ʱ��
      
      Where NO = v_Modifybillout.No And Mod(��¼����, 10) = v_Modifybillout.��¼���� And ��¼״̬ <> 2 And ��� = v_Modifybillout.���;
    
      --������ˣ��ظ����Ҳû�й�ϵ��
      If Intautoverify_In = 1 Then
        If Bill_In = 9 Then
          Zl_������ʼ�¼_Verify(No_In, ����Ա���_In, ����Ա����_In, v_Modifybillout.���, ��ҩʱ��_In);
        End If;
      End If;
    
      --�����������
      Zl_ҩƷ�շ���¼_��������(v_Modifybillout.Id);
    End Loop;
  Else
    --����סԺ����
    For v_Modifybillin In c_Modifybillin Loop
      n_ƽ���ɱ��� := Round(Zl_Fun_Getoutcost(v_Modifybillin.ҩƷid, v_Modifybillin.����, Partid_In), 5);
      Dbl�ɱ����  := Round(n_ƽ���ɱ��� * Nvl(v_Modifybillin.����, 0), Intdigit_In);
      --���۽��
      Dblʵ�ʽ�� := Nvl(v_Modifybillin.���, 0);
      --���
      Dblʵ�ʲ�� := Round(Dblʵ�ʽ�� - Dbl�ɱ����, Intdigit_In);
    
      --����ҩƷ�շ���¼�����۽��ɱ�����ۡ�����˵���Ϣ
      Update ҩƷ�շ���¼
      Set �ɱ��� = n_ƽ���ɱ���, �ɱ���� = Dbl�ɱ����, ��� = Dblʵ�ʲ��, �ⷿid = Partid_In, ��ҩ�� = Decode(��ҩ��_In, Null, ��ҩ��, ��ҩ��_In),
          �˲��� = �˲���_In, �˲����� = v_�˲�����, ��ҩ���� = Decode(��ҩ��_In, Null, ��ҩ����, Date����ʱ��),
          ������ = Decode(У����_In, Null, ������, У����_In), ����� = Decode(People_In, Null, Zl_Username, People_In),
          ������� = Date����ʱ��, ��ҩ��ʽ = ��ҩ��ʽ_In, ע��֤�� = v_Modifybillout.��������
      Where ID = v_Modifybillin.Id;
    
      If Bln�շ��뷢ҩ���� = 1 Then
        Update ҩƷ���
        Set �������� = Nvl(��������, 0) - Nvl(v_Modifybillin.����, 0), ʵ������ = Nvl(ʵ������, 0) - Nvl(v_Modifybillin.����, 0),
            ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) - Nvl(v_Modifybillin.���, 0), ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) - Dblʵ�ʲ��
        Where �ⷿid + 0 = Partid_In And ҩƷid = v_Modifybillin.ҩƷid And ���� = 1 And Nvl(����, 0) = v_Modifybillin.����;
      Else
        Update ҩƷ���
        Set ʵ������ = Nvl(ʵ������, 0) - Nvl(v_Modifybillin.����, 0), ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) - Nvl(v_Modifybillin.���, 0),
            ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) - Dblʵ�ʲ��
        Where �ⷿid + 0 = Partid_In And ҩƷid = v_Modifybillin.ҩƷid And ���� = 1 And Nvl(����, 0) = v_Modifybillin.����;
      End If;
    
      If Sql%RowCount = 0 Then
        If Bln�շ��뷢ҩ���� = 1 Then
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, ��������, ʵ������, ʵ�ʽ��, ʵ�ʲ��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴβ���, Ч��, �ϴ���������, ��׼�ĺ�, ƽ���ɱ���)
          Values
            (Partid_In, v_Modifybillin.ҩƷid, v_Modifybillin.����, 1, 0 - Nvl(v_Modifybillin.����, 0),
             0 - Nvl(v_Modifybillin.����, 0), 0 - Nvl(v_Modifybillin.���, 0), 0 - Dblʵ�ʲ��, v_Modifybillin.��ҩ��λid,
             v_Modifybillin.�ɱ���, v_Modifybillin.����, v_Modifybillin.����, v_Modifybillin.Ч��, v_Modifybillin.��������,
             v_Modifybillin.��׼�ĺ�, v_Modifybillin.�ɱ���);
        Else
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, ʵ������, ʵ�ʽ��, ʵ�ʲ��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴβ���, Ч��, �ϴ���������, ��׼�ĺ�, ƽ���ɱ���)
          Values
            (Partid_In, v_Modifybillin.ҩƷid, v_Modifybillin.����, 1, 0 - Nvl(v_Modifybillin.����, 0),
             0 - Nvl(v_Modifybillin.���, 0), 0 - Dblʵ�ʲ��, v_Modifybillin.��ҩ��λid, v_Modifybillin.�ɱ���, v_Modifybillin.����,
             v_Modifybillin.����, v_Modifybillin.Ч��, v_Modifybillin.��������, v_Modifybillin.��׼�ĺ�, v_Modifybillin.�ɱ���);
        End If;
      End If;
    
      Delete ҩƷ���
      Where �ⷿid + 0 = Partid_In And ҩƷid = v_Modifybillin.ҩƷid And Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And
            Nvl(ʵ�ʽ��, 0) = 0 And Nvl(ʵ�ʲ��, 0) = 0 And ���� = 1;
    
      --���·��ü�¼��ִ��״̬(��ִ��)
      Update סԺ���ü�¼
      Set ִ��״̬ = 1, ִ���� = Decode(People_In, Null, Zl_Username, People_In), ִ�в���id = Partid_In, ִ��ʱ�� = Date����ʱ��
      
      Where ID = v_Modifybillin.����id;
    
      --������ˣ��ظ����Ҳû�й�ϵ��
      If Intautoverify_In = 1 Then
        If Bill_In = 9 Then
          Zl_סԺ���ʼ�¼_Verify(No_In, ����Ա���_In, ����Ա����_In, v_Modifybillin.���, v_Modifybillin.����id, ��ҩʱ��_In);
        End If;
      End If;
    
      --�����������
      Zl_ҩƷ�շ���¼_��������(v_Modifybillin.Id);
    End Loop;
  End If;

  --���»�ɾ��δ��ҩƷ��¼
  Delete δ��ҩƷ��¼ Where NO = No_In And ���� = Bill_In And (�ⷿid + 0 = Partid_In Or �ⷿid Is Null);

  If Bill_In = 8 Then
    Begin
      --�ƶ�֧������Ŀ�ڷ�ҩ��̬��������������Ϣ�Ĺ���
      Execute Immediate 'Begin zl_������Ϣ_����(:1,:2); End;'
        Using 6, No_In || ',' || Partid_In;
    Exception
      When Others Then
        Null;
    End;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�շ���¼_������ҩ;
/

--93587:��ҵ��,2016-02-25,����ִ��״̬��ֵ����
Create Or Replace Procedure Zl_ҩƷ�շ���¼_���Ŀⷿ
(
  Partid_In       In ҩƷ�շ���¼.�ⷿid%Type,
  Bill_In         In ҩƷ�շ���¼.����%Type,
  No_In           In ҩƷ�շ���¼.No%Type,
  Otherstockid_In In ҩƷ�շ���¼.�ⷿid%Type,
  ����_In         In Number := 1,
  Date_In         In ҩƷ�շ���¼.��������%Type :=Null
) Is
  --���¼�����
  Cursor c_Modifybillout Is
    Select a.Id, a.ҩƷid, a.������id, a.���ϵ��, a.����id, Nvl(a.ʵ������, 0) * Nvl(a.����, 1) ����, Nvl(a.���۽��, 0) ���, Nvl(a.����, 0) ����,
           a.��ҩ��λid, a.�ɱ���, a.����, a.����, a.Ч��, a.��������, a.��׼�ĺ�
    From ҩƷ�շ���¼ a, ������ü�¼ b
    Where a.No = No_In And a.���� = Bill_In And (a.�ⷿid + 0 = Otherstockid_In Or a.�ⷿid Is Null) And
          Nvl(a.ժҪ, 'С��') <> '�ܷ�' And a.����id = b.Id And Nvl(b.ִ��״̬,0) <> 1 And a.����� Is Null;

  Cursor c_Modifybillin Is
    Select a.Id, a.ҩƷid, a.������id, a.���ϵ��, a.����id, Nvl(a.ʵ������, 0) * Nvl(a.����, 1) ����, Nvl(a.���۽��, 0) ���, Nvl(a.����, 0) ����,
           a.��ҩ��λid, a.�ɱ���, a.����, a.����, a.Ч��, a.��������, a.��׼�ĺ�
    From ҩƷ�շ���¼ a, סԺ���ü�¼ b
    Where a.No = No_In And a.���� = Bill_In And (a.�ⷿid + 0 = Otherstockid_In Or a.�ⷿid Is Null) And
          Nvl(a.ժҪ, 'С��') <> '�ܷ�' And a.����id = b.Id And Nvl(b.ִ��״̬,0) <> 1 And a.����� Is Null;

  --������������δ�����
  Cursor c_Billout Is
    Select b.ʵ�ս��, b.����id, 0 ��ҳid, 0 ���˲���id, b.���˿���id, b.��������id, b.ִ�в���id, b.������Ŀid, b.�����־
    From ҩƷ�շ���¼ a, ������ü�¼ b
    Where a.����id = b.Id And Nvl(b.ִ��״̬,0) <> 1 And a.No = No_In And a.���� = Bill_In And
          (a.�ⷿid + 0 = Otherstockid_In Or a.�ⷿid Is Null) And Nvl(a.ժҪ, 'С��') <> '�ܷ�' And a.����� Is Null And b.��¼���� = 2 And
          b.��¼״̬ = 1;

  Cursor c_Billin Is
    Select b.ʵ�ս��, b.����id, b.��ҳid, b.���˲���id, b.���˿���id, b.��������id, b.ִ�в���id, b.������Ŀid, b.�����־
    From ҩƷ�շ���¼ a, סԺ���ü�¼ b
    Where a.����id = b.Id And Nvl(b.ִ��״̬,0) <> 1 And a.No = No_In And a.���� = Bill_In And
          (a.�ⷿid + 0 = Otherstockid_In Or a.�ⷿid Is Null) And Nvl(a.ժҪ, 'С��') <> '�ܷ�' And a.����� Is Null And b.��¼���� = 2 And
          b.��¼״̬ = 1;

  r_Modifybillout   c_Modifybillout%Rowtype;
  r_Modifybillin    c_Modifybillin%Rowtype;
  r_Billout         c_Billout%Rowtype;
  r_Billin          c_Billin%Rowtype;
  Bln�շ��뷢ҩ���� Number(1);
  v_Count           Number;
Begin
  Begin
    Select 0
    Into Bln�շ��뷢ҩ����
    From δ��ҩƷ��¼
    Where ���� = Bill_In And No = No_In And �ⷿid + 0 = Otherstockid_In;
  Exception
    When Others Then
      Bln�շ��뷢ҩ���� := 1;
  End;

  --����ԭ�ⷿ�Ŀ��Կ�棬���ֿⷿ�Ŀ��ÿ��
  If ����_In = 1 Then
    --��������
    For r_Modifybillout In c_Modifybillout Loop
      If Bln�շ��뷢ҩ���� = 0 Then
        Update ҩƷ���
        Set �������� = Nvl(��������, 0) + Nvl(r_Modifybillout.����, 0)
        Where �ⷿid + 0 = Otherstockid_In And ҩƷid = r_Modifybillout.ҩƷid And ���� = 1 And Nvl(����, 0) = r_Modifybillout.����;
      
        Update ҩƷ���
        Set �������� = Nvl(��������, 0) - Nvl(r_Modifybillout.����, 0)
        Where �ⷿid + 0 = Partid_In And ҩƷid = r_Modifybillout.ҩƷid And ���� = 1 And Nvl(����, 0) = r_Modifybillout.����;
        
        If Sql%Rowcount = 0 Then
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, ��������, ʵ�ʽ��, ʵ�ʲ��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴβ���, Ч��, �ϴ���������, ��׼�ĺ�)
          Values
            (Partid_In, r_Modifybillout.ҩƷid, r_Modifybillout.����, 1, 0 - Nvl(r_Modifybillout.����, 0), 0, 0,
             r_Modifybillout.��ҩ��λid, r_Modifybillout.�ɱ���, r_Modifybillout.����, r_Modifybillout.����, r_Modifybillout.Ч��,
             r_Modifybillout.��������, r_Modifybillout.��׼�ĺ�);
        End If;
      End If;
    End Loop;
  Else
    --����סԺ
    For r_Modifybillin In c_Modifybillin Loop
      If Bln�շ��뷢ҩ���� = 0 Then
        Update ҩƷ���
        Set �������� = Nvl(��������, 0) + Nvl(r_Modifybillin.����, 0)
        Where �ⷿid + 0 = Otherstockid_In And ҩƷid = r_Modifybillin.ҩƷid And ���� = 1 And Nvl(����, 0) = r_Modifybillin.����;
      
        Update ҩƷ���
        Set �������� = Nvl(��������, 0) - Nvl(r_Modifybillin.����, 0)
        Where �ⷿid + 0 = Partid_In And ҩƷid = r_Modifybillin.ҩƷid And ���� = 1 And Nvl(����, 0) = r_Modifybillin.����;
      
        If Sql%Rowcount = 0 Then
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, ��������, ʵ�ʽ��, ʵ�ʲ��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴβ���, Ч��, �ϴ���������, ��׼�ĺ�)
          Values
            (Partid_In, r_Modifybillin.ҩƷid, r_Modifybillin.����, 1, 0 - Nvl(r_Modifybillin.����, 0), 0, 0,
             r_Modifybillin.��ҩ��λid, r_Modifybillin.�ɱ���, r_Modifybillin.����, r_Modifybillin.����, r_Modifybillin.Ч��,
             r_Modifybillin.��������, r_Modifybillin.��׼�ĺ�);
        End If;
      End If;
    End Loop;
  End If;

  --��������ҩ������������ı�ⷿID
  If ����_In = 1 Then
    --��������
    For r_Billout In c_Billout Loop
      --��ԭ�ⷿ��δ�����
      Update ����δ�����
      Set ��� = Nvl(���, 0) - Nvl(r_Billout.ʵ�ս��, 0)
      Where ����id = r_Billout.����id And Nvl(��ҳid, 0) = Nvl(r_Billout.��ҳid, 0) And
            Nvl(���˲���id, 0) = Nvl(r_Billout.���˲���id, 0) And Nvl(���˿���id, 0) = Nvl(r_Billout.���˿���id, 0) And
            Nvl(��������id, 0) = Nvl(r_Billout.��������id, 0) And Nvl(ִ�в���id, 0) = Nvl(r_Billout.ִ�в���id, 0) And
            ������Ŀid + 0 = r_Billout.������Ŀid And ��Դ;�� + 0 = r_Billout.�����־;
            
      If Sql%Rowcount <> 0 Then 
        --�����ֿⷿ��δ�����
        Update ����δ�����
        Set ��� = Nvl(���, 0) + Nvl(r_Billout.ʵ�ս��, 0)
        Where ����id = r_Billout.����id And Nvl(��ҳid, 0) = Nvl(r_Billout.��ҳid, 0) And
              Nvl(���˲���id, 0) = Nvl(r_Billout.���˲���id, 0) And Nvl(���˿���id, 0) = Nvl(r_Billout.���˿���id, 0) And
              Nvl(��������id, 0) = Nvl(r_Billout.��������id, 0) And Nvl(ִ�в���id, 0) = Partid_In And ������Ŀid + 0 = r_Billout.������Ŀid And
              ��Դ;�� + 0 = r_Billout.�����־;
        
        If Sql%Rowcount = 0 Then 
          Insert Into ����δ����� 
            (����id,���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���) 
          Values 
            (r_Billout.����id, r_Billout.���˿���id, r_Billout.��������id, Partid_In, 
             r_Billout.������Ŀid, r_Billout.�����־, Nvl(r_Billout.ʵ�ս��, 0)); 
        End If;
      end if;
    End Loop;
  Else
    --����סԺ
    For r_Billin In c_Billin Loop
      --��ԭ�ⷿ��δ�����
      Update ����δ�����
      Set ��� = Nvl(���, 0) - Nvl(r_Billin.ʵ�ս��, 0)
      Where ����id = r_Billin.����id And Nvl(��ҳid, 0) = Nvl(r_Billin.��ҳid, 0) And
            Nvl(���˲���id, 0) = Nvl(r_Billin.���˲���id, 0) And Nvl(���˿���id, 0) = Nvl(r_Billin.���˿���id, 0) And
            Nvl(��������id, 0) = Nvl(r_Billin.��������id, 0) And Nvl(ִ�в���id, 0) = Nvl(r_Billin.ִ�в���id, 0) And
            ������Ŀid + 0 = r_Billin.������Ŀid And ��Դ;�� + 0 = r_Billin.�����־;
      
      If Sql%Rowcount <> 0 Then 
        --�����ֿⷿ��δ�����
        Update ����δ�����
        Set ��� = Nvl(���, 0) + Nvl(r_Billin.ʵ�ս��, 0)
        Where ����id = r_Billin.����id And Nvl(��ҳid, 0) = Nvl(r_Billin.��ҳid, 0) And
              Nvl(���˲���id, 0) = Nvl(r_Billin.���˲���id, 0) And Nvl(���˿���id, 0) = Nvl(r_Billin.���˿���id, 0) And
              Nvl(��������id, 0) = Nvl(r_Billin.��������id, 0) And Nvl(ִ�в���id, 0) = Partid_In And ������Ŀid + 0 = r_Billin.������Ŀid And
              ��Դ;�� + 0 = r_Billin.�����־;
             
        If Sql%Rowcount = 0 Then 
          Insert Into ����δ����� 
            (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���) 
          Values 
            (r_Billin.����id, r_Billin.��ҳid, r_Billin.���˲���id, r_Billin.���˿���id, r_Billin.��������id, Partid_In, r_Billin.������Ŀid, 
             r_Billin.�����־, Nvl(r_Billin.ʵ�ս��, 0)); 
        End If; 
      end if;
    End Loop;
  End If;
  
  delete from ����δ����� Where ���=0;

  If ����_In = 1 Then
    Update ������ü�¼
    Set ִ�в���id = Partid_In
    Where Id In
          (Select Distinct ����id From ҩƷ�շ���¼ Where No = No_In And ���� = Bill_In And �ⷿid + 0 = Otherstockid_In);
  Else
    Update סԺ���ü�¼
    Set ִ�в���id = Partid_In
    Where Id In
          (Select Distinct ����id From ҩƷ�շ���¼ Where No = No_In And ���� = Bill_In And �ⷿid + 0 = Otherstockid_In);
  End If;

  --�޸ĸõ������м�¼(��ҩ���ٴ��������)
  Update ҩƷ�շ���¼ Set �ⷿid = Partid_In Where No = No_In And ���� = Bill_In And �ⷿid + 0 = Otherstockid_In;

  --�޸�δ��ҩƷ��¼
  Begin
    Select 1 Into v_Count From δ��ҩƷ��¼ Where �ⷿid + 0 = Partid_In And No = No_In And ���� = Bill_In;
  Exception
    When Others Then
      v_Count := 0;
  End;

  If v_Count = 0 Then
    Update δ��ҩƷ��¼ Set �ⷿid = Partid_In Where No = No_In And ���� = Bill_In And �ⷿid + 0 = Otherstockid_In;
  Else
    Delete δ��ҩƷ��¼ Where No = No_In And ���� = Bill_In And �ⷿid + 0 = Otherstockid_In;
  End If;
  
  If Date_In Is Not Null Then
     Delete From  ���˷��û��� Where ����>=Date_In;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_ҩƷ�շ���¼_���Ŀⷿ;
/
--93706:��ΰ��,2016-03-09,��������Ϊ2�����һ��ʱ����������������31������⴦��
Create Or Replace Function Zl_Age_Calc
(
  ����id_In   ������Ϣ.����id%Type,
  ��������_In Date := Null,
  ��������_In Date := Null
) Return Varchar2
--����:���ݳ������ڼ�������.����Ǽǲ���,�������䲻��.
  --����:1�����ڣ�XСʱ[X����],1����1�����ڣ�X��,1����1�����ڣ�X��[X��],1������ͯ�������ޣ�X��[X��],>=��ͯ�������ޣ�X��
  --˵��:1�����ڣ���ָ����������24Сʱ��;1�����ڣ���ָ������㣻����7.8�ճ�����8.8�ղ���1��;1�����ڣ�Ҳ�Ƕ�����㡣;�����ڡ�����ָ��<����
 As
  d_��������      Date;
  d_��������      Date;
  n_Days          Number;
  n_Months        Number;
  v_����          ������Ϣ.����%Type;
  n_Upperagelimit Number; --����:��������

  v_Return Varchar2(20); --���ڲ�����Ϣ����ر�������ֶ�Ϊ10���ַ��������������10���ַ���5������
Begin
  --����ǼǵĲ��˲�����������
  If Nvl(����id_In, 0) <> 0 Then
    Begin
      Select ����
      Into v_����
      From ������Ϣ
      Where ����id = ����id_In And Floor(Sysdate - �Ǽ�ʱ��) = 0 And ���� Is Not Null;
    Exception
      When Others Then
        Null;
    End;
    If v_���� Is Not Null Then
      v_Return := v_����;
      Return v_Return;
    End If;
  End If;

  If ��������_In Is Null Then
    If Nvl(����id_In, 0) <> 0 Then
      Select �������� Into d_�������� From ������Ϣ Where ����id = ����id_In;
    End If;
    If d_�������� Is Null Then
      Return Null;
    End If;
  Else
    d_�������� := ��������_In;
  End If;
  If ��������_In Is Null Then
    Select Sysdate Into d_�������� From Dual;
  Else
    d_�������� := ��������_In;
  End If;
  --����������ڴ��ڼ�������,��ֱ��Ϊ0Сʱ
  If (d_�������� - d_��������) < 0 Then
    v_Return := '0Сʱ';
    Return v_Return;
  End If;
  --��ȡ��ͯ���������
  Begin
    Select Nvl(����ֵ, 14)
    Into n_Upperagelimit
    From Zlparameters
    Where ϵͳ = 100 And Nvl(ģ��, 0) = 0 And ������ = 147;
  Exception
    When Others Then
      n_Upperagelimit := 14;
  End;

  n_Months := Trunc(Months_Between(d_��������, d_��������));
  If n_Months < 12 * n_Upperagelimit Then
    --С��1������
    If n_Months < 12 Then
      --С��1��
      If n_Months < 1 Then
        n_Days := Trunc(d_�������� - d_��������);
        --һ������
        If n_Days = 0 Then
          n_Days := Trunc((d_�������� - d_��������) * 24 * 60);
          If Mod(n_Days, 60) = 0 Then
            v_Return := n_Days / 60 || 'Сʱ';
          Else
            v_Return := Floor(n_Days / 60) || 'Сʱ' || Mod(n_Days, 60) || '����';
          End If;
        Else
          --һ����һ��
          v_Return := n_Days || '��';
        End If;
      Else
        --����1��
        n_Days := Trunc(Add_Months(d_��������, -1 * n_Months) - d_��������);
	If n_Days >= 31 Then
          --��Լ���������2�·����һ��,�������ڸպô���2�·����һ���ҵ��첻�Ǳ��µ����һ��
          --�磺�������ڣ�2016-02-29   �������ڣ�2015-01-30
          n_Months := n_Months + 1;
          n_Days   := n_Days - 31;
        End If;
        If n_Days = 0 Then
          v_Return := n_Months || '��';
        Else
          v_Return := n_Months || '��' || n_Days || '��';
        End If;
      End If;
    Else
      --1�굽С��Ӥ���������޵����
      If Mod(n_Months, 12) = 0 Then
        v_Return := n_Months / 12 || '��';
      Else
        v_Return := Floor(n_Months / 12) || '��' || Mod(n_Months, 12) || '��';
      End If;
    End If;
  Else
    --���ڵ���Ӥ����������(ֱ��X��)
    v_Return := Floor(n_Months / 12) || '��';
  End If;
  Return v_Return;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Age_Calc;
/
---------------------------------------------------------------------------------------------------
--����ϵͳ�������İ汾��
-------------------------------------------------------------------------------------------------------
Delete from zlFilesUpgrade Where UPPER(�ļ���)='REGCOM.DLL';
Insert Into zlFilesUpgrade
  (���, �ļ�����, �ļ���, �汾��, �޸�����, ��װ·��, �ļ�˵��, ǿ�Ƹ���, �Զ�ע��, ��������)
  Select Max(���) + 1, 4, 'REGCOM.DLL', '', '', '[System]', '����ע���ļ�', 1, 1, Sysdate
  From zlFilesUpgrade
  Where Upper(�ļ���) = 'REGCOM.DLL';

Delete from zlFilesUpgrade Where UPPER(�ļ���)='ZL9PACSIMAGECAP.DLL';
Insert Into zlFilesUpgrade (�ļ�����,�ļ���,�汾��,�޸�����,����ϵͳ,ҵ�񲿼�,��װ·��,�ļ�˵��,ǿ�Ƹ���,�Զ�ע��,��������,���) select 1,'ZL9PACSIMAGECAP.DLL','', Null ,'1','ZL9PACSWORK','[Appsoft]\Apply','','0','1',sysdate,��� from Dual a,(Select max(to_number(���))+1 ��� from zlfilesupgrade) b where not exists (select 1 from zlfilesupgrade where upper(�ļ���)='ZL9PACSIMAGECAP.DLL');

Delete from zlFilesUpgrade Where UPPER(�ļ���)='ZL9XWINTERFACE.DLL';
Insert Into zlFilesUpgrade (�ļ�����,�ļ���,�汾��,�޸�����,����ϵͳ,ҵ�񲿼�,��װ·��,�ļ�˵��,ǿ�Ƹ���,�Զ�ע��,��������,���) select 0,'ZL9XWINTERFACE.DLL','', Null ,'1','','[Appsoft]\Apply','','0','1',sysdate,��� from Dual a,(Select max(to_number(���))+1 ��� from zlfilesupgrade) b where not exists (select 1 from zlfilesupgrade where upper(�ļ���)='ZL9XWINTERFACE.DLL');

--ϵͳ�汾��
Update zlSystems Set �汾��='10.34.70' Where ���=&n_System;







--�����汾��
Commit;