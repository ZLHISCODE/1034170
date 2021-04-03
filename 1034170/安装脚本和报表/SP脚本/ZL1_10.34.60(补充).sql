--[��������]1
--[�����߰汾��]10.34.30
--���ű�֧�ִ�ZLHIS+ v10.34.50 ������ v10.34.60
--����ϵͳ�����ߵ�¼PLSQL��ִ�����нű�
--�ű�ִ�к����ֹ�������������
Define n_System=100;
-------------------------------------------------------------------------------
--�ṹ��������
-------------------------------------------------------------------------------
--91780:������,2015-12-28,��ͬ��Ʊ������
Alter Table ��Ա�ս�Ʊ�� Add ���� Varchar2(20);

--91427:����,2015-12-22,����ҩƷ����ϵͳ
create table ҩƷ���ռ�¼
(
id number(18),
NO varchar2(8),
�ⷿid number(18),
��ҩ��λid number(18),
������ varchar2(200),
�������� date,
������ varchar2(200),
�������� date,
�Ƿ�ϸ� number(1),
��ע  varchar2(1000)
) TABLESPACE zl9MedLst
    initrans 20;

create table ҩƷ������ϸ 
(
����id number(18),
ҩƷid number(18),
�ɱ��� number(16,7),
���ۼ� number(16,7),
��ҩ���� number(16,5),
���� varchar2(20),
�������� date,
Ч�� date,
���� varchar2(60),
��׼�ĺ� varchar2(40),
��ҩ���� date,
�Ƿ�ϸ� number(1)
) TABLESPACE zl9MedLst
    initrans 20;

Create Sequence ҩƷ���ռ�¼_ID Start With 1; 

Alter Table ҩƷ���ռ�¼ Add Constraint ҩƷ���ռ�¼_PK Primary Key (ID) Using Index Tablespace zl9indexhis;

Alter Table ҩƷ������ϸ Add Constraint ҩƷ������ϸ_UQ_����ID Unique (����id,ҩƷid) Using Index Tablespace zl9indexhis;

Alter Table ҩƷ������ϸ Modify ����id Constraint ҩƷ������ϸ_NN_����id Not Null;   

Create Index ҩƷ���ռ�¼_IX_�ⷿid On ҩƷ���ռ�¼(�ⷿid) Tablespace zl9Indexhis;

Create Index ҩƷ���ռ�¼_IX_��ҩ��λid On ҩƷ���ռ�¼(��ҩ��λid) Tablespace zl9Indexhis;   

Create Index ҩƷ���ռ�¼_IX_NO On ҩƷ���ռ�¼(NO) Tablespace zl9Indexhis;  

Create Index ҩƷ������ϸ_IX_ҩƷid On ҩƷ������ϸ(ҩƷid) Tablespace zl9Indexhis;  

--91225:������,2015-12-16,��Ⱦ������ϵͳ ��������
create table ��Ⱦ��Ŀ¼(
   ���� VARCHAR2(10),
   ���� VARCHAR2(200), 
   ���� VARCHAR2(200), 
   ˵�� VARCHAR2(500)
) TABLESPACE zl9EprDat;

create table �������Լ�¼(
   ID    Number(18),
   ����ID number(18), 
   ��ҳid NUMBER(5),
   �Һŵ� VARCHAR2(8),
   �ͼ�ʱ�� date,
   �ͼ����ID number(18), 
   �ͼ�ҽ�� VARCHAR2(201), 
   �걾���� VARCHAR2(60),
   ������� VARCHAR2(1000),
   ��Ⱦ������ VARCHAR2(200),
   ���ʱ�� date,
   �Ǽ�ʱ�� date,
   �Ǽ��� VARCHAR2(100),
   �Ǽǿ���ID number(18), 
   ��¼״̬ number(2),
   ������ VARCHAR2(100),
   ����ʱ�� date,
   �������˵�� VARCHAR2(1000),
   �ļ�ID number(18),
   ��ת�� Number(3)
) TABLESPACE zl9EprDat;

create table �������淴��(
   �ļ�ID NUMBER(18),
   �Ǽ�ʱ�� date, 
   �Ǽ��� VARCHAR2(100),
   ��¼״̬ NUMBER(3),
   �������� VARCHAR2 (500),
   ������ VARCHAR2(100),
   ����ʱ�� date,
   �������˵�� VARCHAR2(500),
   ��ת�� Number(3)
) TABLESPACE zl9EprDat;

alter table �����걨��¼ Add(�������� VARCHAR2(50),����ҽ�� VARCHAR2(100),������ VARCHAR2(100),����ʱ�� Date,����ID NUMBER(18),��ҳID NUMBER(18),������Դ NUMBER(3));

Create Sequence �������Լ�¼_ID Start With 1;

Alter Table ��Ⱦ��Ŀ¼ Add Constraint ��Ⱦ��Ŀ¼_PK Primary Key (����) Using Index Tablespace zl9Indexcis;

Alter Table ��Ⱦ��Ŀ¼ Add Constraint ��Ⱦ��Ŀ¼_UQ_���� Unique (����) Using Index Tablespace zl9Indexhis;

Alter Table �������Լ�¼ Add Constraint �������Լ�¼_PK Primary Key (ID) Using Index Tablespace zl9Indexcis;

Alter Table �������淴�� Add Constraint �������淴��_PK Primary Key (�ļ�ID,�Ǽ�ʱ��) Using Index Tablespace zl9Indexcis;

Create Index �������Լ�¼_IX_����ID On �������Լ�¼(����ID,��ҳID)  Tablespace zl9Indexcis;

Create Index �������Լ�¼_IX_�Ǽ�ʱ�� On �������Լ�¼(�Ǽ�ʱ��)  Tablespace zl9Indexcis;

Create Index �������Լ�¼_IX_�Һŵ� On �������Լ�¼(�Һŵ�)  Tablespace zl9Indexcis;

Create Index �������Լ�¼_IX_��ת�� On �������Լ�¼(��ת��) Tablespace zl9Indexcis;

Create Index �������Լ�¼_IX_�ļ�ID On �������Լ�¼(�ļ�ID) Tablespace zl9Indexcis;

Create Index �����걨��¼_IX_���� On �����걨��¼(����) Tablespace zl9Indexcis;

Create Index �����걨��¼_IX_����ID On �����걨��¼(����ID,��ҳID)  Tablespace zl9Indexcis;

Create Index �������淴��_IX_��ת�� On �������淴��(��ת��) Tablespace zl9Indexcis;

Create Index �������淴��_IX_�Ǽ�ʱ�� On �������淴��(�Ǽ�ʱ��) Tablespace zl9Indexcis;

--91687:������,2015-12-15,��������
Create Index ���ﴩ��̨_Ix_��������id On ���ﴩ��̨(��������id) Pctfree 5 Tablespace Zl9indexcis Nologging;

--90666:����,2015-12-07,������ʡ�����淶Ҫ��,�������²�λ:����
Alter Table ���¼�¼��Ŀ Modify ��¼�� Varchar2(20);

--92493:������,2016-01-08,��������ǰ�������ֶ� ���没��
Alter Table ��������ǰ�� Add ���没�� Varchar2(80);

--91712:Ϳ����,2015-12-16,����Ƭ��Ϣ�����������
Create Index ����Ƭ��Ϣ_IX_����ҽ��ID On ����Ƭ��Ϣ(����ҽ��ID) Tablespace zl9Indexcis nologging;

--91225:������,2015-12-16,��Ⱦ������ϵͳ ��������
Alter Table �������Լ�¼ Add Constraint �������Լ�¼_FK_����ID Foreign Key (����ID) References ������Ϣ(����ID);
Alter Table �������Լ�¼ Add Constraint �������Լ�¼_FK_��ҳID Foreign Key (����ID,��ҳID) References ������ҳ(����ID,��ҳID);
Alter Table �������Լ�¼ Add Constraint �������Լ�¼_FK_�ͼ����ID Foreign Key (�ͼ����ID) References ���ű�(ID);
Alter Table �������Լ�¼ Add Constraint �������Լ�¼_FK_�Ǽǿ���ID Foreign Key (�Ǽǿ���ID) References ���ű�(ID);
Alter Table �������淴�� Add Constraint �������淴��_FK_�ļ�ID Foreign Key (�ļ�ID) References �����걨��¼ (�ļ�ID) On Delete Cascade;
Alter Table �����걨��¼ Add Constraint �����걨��¼_FK_����ID Foreign Key (����ID) References ������Ϣ (����ID);

--91427:����,2015-12-22,����ҩƷ����ϵͳ
Alter Table ҩƷ���ռ�¼ Add Constraint ҩƷ���ռ�¼_FK_�ⷿid Foreign Key (�ⷿid) References ���ű�(ID) On Delete Cascade;
Alter Table ҩƷ���ռ�¼ Add Constraint ҩƷ���ռ�¼_FK_��ҩ��λid Foreign Key (��ҩ��λid) References ��Ӧ��(ID) On Delete Cascade;
Alter Table ҩƷ������ϸ Add Constraint ҩƷ������ϸ_FK_����id Foreign Key (����id) References ҩƷ���ռ�¼(ID) On Delete Cascade;
Alter Table ҩƷ������ϸ Add Constraint ҩƷ������ϸ_FK_ҩƷid Foreign Key (ҩƷid) References �շ���ĿĿ¼(ID) On Delete Cascade;


-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--89717:��ΰ��,2016-01-14,��Ժ������ȡ�����·��
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1256, 0, 0, 0, 0, 13, '��Ժ������ȡ�����·��', '0', '0','������ô˲�������Ժ�Ĳ��˲�����ȡ����ɵ�·����'
  From Dual
  Where Not Exists (Select 1
         From zlParameters
         Where ������ = '��Ժ������ȡ�����·��' And Nvl(ģ��, 0) = 1256 And Nvl(ϵͳ, 0) = &n_System);

--92321:������,2015-01-13,���뵥���û���
Insert Into Zlparameters
  (Id, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, -null, -null, -null, -null, -null, 238, '���뵥���û���', '111', '111',    
  '�������뵥���û��ڣ�������ö�Ӧ�����뵥����ҽ���´�ʱֻ��ͨ�����뵥��ʽ�´�����ҽ���´�������������Ŀ�����Һ�Ҳ���Զ��������뵥���������д��'|| Chr(13) ||'��һλ�����ǣ���顢���顢��Ѫ'
  From Dual Where Not Exists (Select 1
         From Zlparameters  Where ������ = '���뵥���û���' And Nvl(ģ��, 0) = 0 And Nvl(ϵͳ, 0) = &n_System); 

--89419:�ŵ���,2015-01-05,��Ժ���˲������÷�
Insert Into Zlparameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1345, 0, 1, 0, 0, 24,'��Ժ���˲������÷�', Null, '0', '�Ѿ���Ժ�Ĳ��˲������÷�'
  From Dual
  Where Not Exists (Select 1 From Zlparameters Where ϵͳ = &n_System And ģ�� = 1345 And ������ = '��Ժ���˲������÷�');

--91671:������,2015-08-30,����ҽʦ�ﵽ�����ȼ��������
Insert Into Zlparameters
  (Id, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, -null, -null, -null, -null, -null, 254, '����ҽʦ�ﵽ�����ȼ��������', '0', '0',        
  '���ò�������ƣ��´�����ҽ��ʱ���������ҽʦ���������ȼ�Ҫ��������Ȩ����δ�������ǰ�ҽ�������ȼ���������������ˣ���ֱ��У�ԣ�'
  || Chr(13) ||'�������ҽʦ������(����)������Ŀ�ȼ�ʱ������Ҫ��ˣ�' 
  From Dual Where Not Exists (Select 1
         From Zlparameters  Where ������ = '����ҽʦ�ﵽ�����ȼ��������' And Nvl(ģ��, 0) = 0 And Nvl(ϵͳ, 0) = &n_System);

--91665:Ƚ����,2015-12-29,���Ӷ൥�ݷֵ��ݽ���ʱҽ������ʧ��ʱֻ�Խ���ɹ������շѵ�ģʽ��
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1121, 0, 0, 0, 0, 104, 'ֻ��ҽ������ɹ������շ�', '0', '0',
         '���ݷֵ��ݽ���ģʽ�£���ҽ������ʧ�ܣ������ֵ��ݽ���ɹ�ʱ�Ƿ�Խ���ɹ��ĵ��ݽ����շѡ�0-ֻ�����е��ݶ�����ҽ������ɹ�����ܼ����շѣ�1-ҽ������ʧ�ܣ������ֵ��ݽ���ɹ�ʱֻ�Խ���ɹ��ĵ��ݽ����շѡ�'
  From Dual
  Where Not Exists (Select 1 From zlParameters Where ϵͳ = &n_System And ģ�� = 1121 And ������ = 'ֻ��ҽ������ɹ������շ�');

--91427:����,2015-12-22,����ҩƷ����ϵͳ
Insert Into zlPrograms(���,����,˵��,ϵͳ,����) Values( 1348,'ҩƷ������չ���','ҩƷ���ǰ�����ⵥҩƷ������Ϣ�Ƿ�ϸ�',&n_System,'zl9MediStore'); 

Insert Into zlMenus(���,ID,�ϼ�ID,����,���,ϵͳ,ģ��,�̱���,ͼ��,˵��)
  Select ���,Zlmenus_Id.Nextval,id,'ҩƷ������չ���' ,'I' ,&n_System,1348 ,'�������' ,114 ,'ҩƷ���ǰ�����ⵥҩƷ������Ϣ�Ƿ�ϸ�' 
         From zlMenus Where ���� = 'ҩ�������ҩƷ���ϵͳ' And ��� = 'ȱʡ' And ϵͳ = &n_System And ģ�� is null;          

Insert Into zlMenus(���,ID,�ϼ�ID,����,���,ϵͳ,ģ��,�̱���,ͼ��,˵��)
  Select ���,Zlmenus_Id.Nextval,id,'ҩƷ������չ���' ,'I' ,&n_System,1348 ,'�������' ,114 ,'ҩƷ���ǰ�����ⵥҩƷ������Ϣ�Ƿ�ϸ�' 
         From zlMenus Where ���� = '��������ҩ������ϵͳ' And ��� = 'ȱʡ' And ϵͳ = &n_System And ģ�� is null;          

Insert Into ������Ʊ�
  (��Ŀ���, ��Ŀ����, �Զ���ȱ, ��Ź���)
  Select 148, 'ҩƷ�������', 0, 0
  From Dual
  Where Not Exists (Select 1 From ������Ʊ� Where ��Ŀ��� = 148 And ��Ŀ���� = 'ҩƷ�������');

--89983:����,2015-12-22,�����������������ʽ,��������ֵ
Update zlParameters
Set ����˵�� = '����������Ϊ�����Ŀ���˲��������������������ʽ��1.����ֵΪ0���������������ʾR���š�2.����ֵΪ1����ÿ��������������ʼ��Ӧ�ĺ���������Ϸ��������"������"������"��"��ʶ��ʼ���ڶ�Ӧ�����ĺ���������Ϸ���"��"��ʶ��ֹ��3.����ֵΪ2,�������������ʾA+����ֵ'
Where ϵͳ = &n_System And ģ�� = 1255 And ������ = 85;

--91225:������,2015-12-22,��Ⱦ������ϵͳ���ģ��źͲ˵�
Insert Into zlComponent(����,����,���汾,�ΰ汾,���汾,ϵͳ,ע���Ʒ����,ע���Ʒ����,ע���Ʒ�汾) Values('zl9Disease','��Ⱦ��ϵͳ����',10,34,0,&n_System,'����ҽԺ��Ϣϵͳ','ZLHIS+','10');
Insert Into zlPrograms(���,����,˵��,ϵͳ,����) Values( 1278,'��Ⱦ��������վ','���ڶԴ�Ⱦ�����浥����ˡ��ϱ��ȹ�����',&n_System,'zl9Disease');

Insert Into zlMenus
  (���, ID, �ϼ�id, ����, ���, ˵��, ϵͳ, ģ��, �̱���, ͼ��)
  Select ���, Zlmenus_Id.Nextval, ID, '��Ⱦ������ϵͳ', 'D', ' ���ڶԴ�Ⱦ�����浥����ˡ��ϱ��ȹ����� ', &n_System, -null, '��Ⱦ������', 99
  From zlMenus
  Where ���� = '�ٴ���Ϣϵͳ' And ��� = 'ȱʡ' And ϵͳ = &n_System And ģ�� Is Null;

Insert Into zlMenus
  (���, ID, �ϼ�id, ����, ���, ˵��, ϵͳ, ģ��, �̱���, ͼ��)
  Select a.���, Zlmenus_Id.Nextval, a.Id, b.*
  From (Select ���, ID From zlMenus Where ���� = '��Ⱦ������ϵͳ' And ��� = 'ȱʡ' And ϵͳ = &n_System And ģ�� Is Null) A,
       (Select ����, ���, ˵��, ϵͳ, ģ��, �̱���, ͼ��
         From zlMenus
         Where 1 = 0
         Union All
         Select '��Ⱦ��������վ', 'D', ' ���ڶԴ�Ⱦ������Ľ��ա���ˡ��ϱ��ȹ�����', &n_System, 1278, '��Ⱦ������', 130
         From Dual
         Union All
         Select ����, ���, ˵��, ϵͳ, ģ��, �̱���, ͼ��
         From zlMenus
         Where 1 = 0) B;

--91225:������,2015-12-21,��Ⱦ������ϵͳ
Update zlParameters
Set ����˵�� = 'ÿλ���ֱ����ͬ��Ϣ���ͣ�1�������ġ�2ҽ�����š�3Σ��ֵ��4���泷����5ҽ����ˡ�6��Ⱦ������'
Where ������ = '�Զ�ˢ������' And ģ�� = 1261 and ϵͳ = &n_System;

Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1260, 1, 0, 0, 0, 28, '�Զ�ˢ�²������ļ��', '', '0',
         '����ÿ���ٷ����Զ�ˢ�²����������������е����ݣ�Ϊ0��ʾ���Զ�ˢ��(���ֹ�ˢ��)��'
  From Dual
  Where Not Exists
   (Select 1 From zlParameters Where ϵͳ = &n_System And ģ�� = 1260 And ������ = '�Զ�ˢ�²������ļ��');

Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1260, 1, 0, 0, 0, 29, '�Զ�ˢ������', '', '0', 'ÿλ���ֱ����ͬ��Ϣ���ͣ�1��Ⱦ������'
  From Dual
  Where Not Exists (Select 1 From zlParameters Where ϵͳ = &n_System And ģ�� = 1260 And ������ = '�Զ�ˢ������');

Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1260, 1, 0, 0, 0, 30, '�Զ�ˢ�²�����������', '', '1', '���ý�����������ɵĲ�����ʾ��������������'
  From Dual
  Where Not Exists
   (Select 1 From zlParameters Where ϵͳ = &n_System And ģ�� = 1260 And ������ = '�Զ�ˢ�²�����������');

--91225:������,2015-12-22,��Ⱦ������ϵͳ��Ӳ���
Insert Into zlParameters(ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
      Select Zlparameters_Id.Nextval, &n_System, 1278, a.* From (Select ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵�� From zlParameters Where 1 = 0 Union All
      Select 1, 0, 0, 0, 1, '������վ�ɹ����ļ�', Null, Null, '��Ⱦ������վ�ɹ�����ļ�,�ǲ����ļ��б��е��ļ�ID' From Dual Union All
      Select 1, 0, 0, 0, 2, '������ϱ�����״̬�²鿴��������ı���', '7', '7', '������ϱ�����״̬�²鿴���������ı���' From Dual Union All
      Select 1, 0, 0, 0, 3, '������ϱ�����״̬�²鿴ָ�������ı������ʼ����', Null, Null, '������ϱ�����״̬�²鿴ָ��ʱ��εı������ʼ����' From Dual Union All
      Select 1, 0, 0, 0, 4, '������ϱ�����״̬�²鿴ָ�������ı���Ľ�������', Null, Null, '������ϱ�����״̬�²鿴ָ��ʱ��εı���Ľ�������' From Dual Union All
      Select 1, 0, 0, 0, 5, '��Ⱦ��ϵͳ�鿴״̬��Χ', '1,1,1,0,1,0', '1,1,1,0,1,0','ѡ��鿴״̬��Χ�ĵı���,���ڵ�λΪ1�Ļ��������ò鿴��״̬�����棬Ϊ0�Ļ������鿴����1λ-�����,��2λ-������,��3λ-���ϱ�,��4λ-���ϱ�,��5λ-����д���濨,��6λ-�Ǵ�Ⱦ��' From Dual Union All
      Select 1, 0, 0, 0, 6, 'δ��д״̬�²鿴��������ı���', '0', '0', 'δ��д״̬�²鿴���������ı���' From Dual Union All
      Select 1, 0, 0, 0, 7, 'δ��д״̬�²鿴ָ�������ı������ʼ����', Null, Null, 'δ��д״̬�²鿴ָ��ʱ��εı������ʼ����' From Dual Union All
      Select 1, 0, 0, 0, 8, 'δ��д״̬�²鿴ָ�������ı���Ľ�������', Null, Null, 'δ��д״̬�²鿴ָ��ʱ��εı���Ľ�������' From Dual Union All
      Select 1, 0, 0, 0, 9, '��ɾ��״̬�²鿴��������ı���', '7', '7', '��ɾ��״̬�²鿴���������ı���' From Dual Union All
      Select 1, 0, 0, 0, 10, '��ɾ��״̬�²鿴ָ�������ı������ʼ����', Null, Null, '��ɾ��״̬�²鿴ָ��ʱ��εı������ʼ����' From Dual Union All
      Select 1, 0, 0, 0, 11, '��ɾ��״̬�²鿴ָ�������ı���Ľ�������', Null, Null, '��ɾ��״̬�²鿴ָ��ʱ��εı���Ľ�������' From Dual Union All
      Select 1, 0, 0, 0, 12, '��ǰ�鿴����Ĺ���״̬', '1', '1', '0-����д,1-��˹�����2-�ϱ�������3-��ɾ����4-���ع���' From Dual Union All
      Select ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵�� From zlParameters Where 1 = 0) A;


Insert Into ҵ����Ϣ����(����,����,˵��) 
Select 'ZLHIS_CIS_032','��Ⱦ�����Խ������','��ʦվ��д��Ⱦ������¼ʱ��������һ��֪ͨ��Ϣ��' From Dual Union All
Select 'ZLHIS_CIS_033','��Ⱦ�����淵������','��Ⱦ��������д������Ҫ�󣬲�����һ������֪ͨ��Ϣ��' From Dual;

--91225:������,2015-12-16,��Ⱦ������ϵͳ ��������
Insert Into zlBaseCode(ϵͳ,����,�̶�,˵��,����) Values( &n_sysTem,'��Ⱦ��Ŀ¼',0,'������Ⱦ��Ŀ¼','ҽ�ƹ���' ); 

Insert Into ��Ⱦ��Ŀ¼(����,����,����) 
Select '01','����','SY' From Dual Union All
Select '02','����','HL' From Dual Union All
Select '03','��Ⱦ�Էǵ��ͷ���','CRXFDXFY' From Dual Union All
Select '04','���̲�(HIV)','AZBHIV' From Dual Union All
Select '05','���̲�(AIDS)','AZBAIDS' From Dual Union All
Select '06','�����Ը���(����)','BDXGYJX' From Dual Union All
Select '07','�����Ը���(����)','BDXGYYX' From Dual Union All
Select '08','�����Ը���(����)','BDXGYBX' From Dual Union All
Select '09','�����Ը���(����)','BDXGYWX' From Dual Union All
Select '10','�����Ը���(δ����)','BDXGYWFX' From Dual Union All
Select '11','���������','GSHZY' From Dual Union All
Select '12','�˸�Ⱦ���²���������','RGRGZBXQLG' From Dual Union All
Select '13','����H1N1����','JXH1N1LG' From Dual Union All
Select '14','����','MZ' From Dual Union All
Select '15','�����Գ�Ѫ��','LXXCXR' From Dual Union All
Select '16','��Ȯ��','KQB' From Dual Union All
Select '17','��������������','LXXYXGY' From Dual Union All
Select '18','�Ǹ���','DGR' From Dual Union All
Select '19','̿��(��̿��)','TJFTJ' From Dual Union All
Select '20','̿��(δ����)','TJWFX' From Dual Union All
Select '21','����(ϸ����)','LJXJX' From Dual Union All
Select '22','����(���װ���)','LJAMBX' From Dual Union All
Select '23','�ν��(Ϳ��)','FJHTY' From Dual Union All
Select '24','�ν��(������)','FJHJPY' From Dual Union All
Select '25','�ν��(����)','FJHJY' From Dual Union All
Select '26','�ν��(δ̵��)','FJHWTJ' From Dual Union All
Select '27','�˺�(�˺�)','SHSH' From Dual Union All
Select '28','�˺�(���˺�)','SHFSH' From Dual Union All
Select '29','�������Լ���Ĥ��','LXXLJSMY' From Dual Union All
Select '30','���տ�','BRK' From Dual Union All
Select '31','�׺�','BH' From Dual Union All
Select '32','���������˷�','XSEPSF' From Dual Union All
Select '33','�ɺ���','XHR' From Dual Union All
Select '34','��³�Ͼ���','BLSJB' From Dual Union All
Select '35','�ܲ���÷��(����)','LBMDYQ' From Dual Union All
Select '36','�ܲ���÷��(����)','LBMDEQ' From Dual Union All
Select '37','�ܲ���÷��(����)','LBMDSQ' From Dual Union All
Select '38','�ܲ���÷��(̥��)','LBMDTC' From Dual Union All
Select '39','�ܲ���÷��(����)','LBMDYX' From Dual Union All
Select '40','���������岡','GDLXTB' From Dual Union All
Select '41','Ѫ���没','XXCB' From Dual Union All
Select '42','ű��(����ű)','LJJRL' From Dual Union All
Select '43','ű��(����ű)','LJEXL' From Dual Union All
Select '44','ű��(δ����)','LJWFX' From Dual Union All
Select '45','�����Ը�ð','LXXGM' From Dual Union All
Select '46','������������','LXXSXY' From Dual Union All
Select '47','����','FZ' From Dual Union All
Select '48','���Գ�Ѫ�Խ�Ĥ��','JXCXXJMY' From Dual Union All
Select '49','��粡','MFB' From Dual Union All
Select '50','�����Ժ͵ط��԰����˺�','LXXHDFXBZSH' From Dual Union All
Select '51','���Ȳ�','HRB' From Dual Union All
Select '52','���没','BCB' From Dual Union All
Select '53','˿�没','SCB' From Dual Union All
Select '54','�����ҡ�ϸ���ԺͰ��װ����������˺��͸��˺�����ĸ�Ⱦ�Ը�к��','CHLXJXHAMBX' From Dual Union All
Select '55','����ڲ�','SZKB' From Dual;

--91225:������,2015-12-16,��Ⱦ������ϵͳ
Insert Into zlBakTables(ϵͳ,���,����,���,ֱ��ת��,ͣ�ô�����)
Select &n_System,6,A.* From (
Select ����,���,ֱ��ת��,ͣ�ô����� From zlBakTables Where 1 = 0 Union All
Select '�������淴��',4,1,-NULL From Dual Union All
Select ����,���,ֱ��ת��,ͣ�ô����� From zlBakTables Where 1 = 0) A;

--91064:��˶,2015-12-08,��Ժҽ���������
Insert Into Zlparameters
  (Id, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, -null, -null, -null, -null, -null, 253, '��Ժҽ�������Ƚ���', '0', '0',
         '��ѡ��ҽ������Ϣ���Ƿ��������¼����Ժҽ����0-��������¼����Ժҽ����1-��������¼����Ժҽ������Ժҽ�������Ƚ���'
  From Dual
  Where Not Exists (Select 1
         From Zlparameters
         Where ������ = '��Ժҽ�������Ƚ���' And Nvl(ģ��, 0) = 0 And Nvl(ϵͳ, 0) = &n_System);

--90666:����,2015-12-07,������ʡ�����淶Ҫ��,�������²�λ:����
Insert Into ���²�λ (��Ŀ���, ��λ, ȱʡ��, �̶���) Values (1, '����', 0, 1);

Update ���¼�¼��Ŀ Set ��¼�� = '��,��,��,��' Where ��Ŀ��� = 1;

--78413:������,2015-12-01,ҽ���嵥��ӡ
Delete From zlProgPrivs Where Upper(����) = 'ZL_ҽ����ӡ��¼_INSERT' or ����='ҽ����ӡ��¼';

--91641:��ΰ��,2015-12-25,·��ƥ��ʱ������Ч
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1256, 0, 0, 0, 0, 12, 'ƥ��ʱ��Ч��ͬ��·������Ŀ', '0', '0',
         '0-������Ŀ��ͬʱ,��Ч����ͬ����·������Ŀ,������ƥ����ͬ��Ч,1-������Ŀ����Ч����ͬʱ������·������Ŀ'
  From Dual
  Where Not Exists (Select 1
         From zlParameters
         Where ������ = 'ƥ��ʱ��Ч��ͬ��·������Ŀ' And Nvl(ģ��, 0) = 1256 And Nvl(ϵͳ, 0) = &n_System);



-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------
--92335:���ϴ�,2016-01-18,����֧����ģʽ�����̲��
Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1801, '����', User, 'zl_��Ա�ɿ����_Update', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1801 And ���� = '����' And Upper(����) = Upper('zl_��Ա�ɿ����_Update'));

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1802, '����', User, 'zl_��Ա�ɿ����_Update', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1802 And ���� = '����' And Upper(����) = Upper('zl_��Ա�ɿ����_Update'));

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1803, '����', User, 'zl_��Ա�ɿ����_Update', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1803 And ���� = '����' And Upper(����) = Upper('zl_��Ա�ɿ����_Update'));

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1804, '����', User, 'zl_��Ա�ɿ����_Update', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1804 And ���� = '����' And Upper(����) = Upper('zl_��Ա�ɿ����_Update'));

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1805, '����', User, 'zl_��Ա�ɿ����_Update', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1805 And ���� = '����' And Upper(����) = Upper('zl_��Ա�ɿ����_Update'));

--89620:��ΰ��,2016-01-15,��ǰ����ٴ�·��ִ�е�Ȩ��
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��,ȱʡֵ)
Select &n_System,1256,A.* From (
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0 Union All
Select '��ǰ���',11,'��ǰ����ٴ�·��ִ�е�Ȩ�ޡ�',1 From Dual Union All
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0) A;

--91487:Ƚ����,2016-01-05,���ղ���������˷ѡ�
Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1124, 'ҽ������', User, '�����˿���Ϣ', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1124 And ���� = 'ҽ������' And Upper(����) = Upper('�����˿���Ϣ'));

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1124, '�����˷�', User, '�����˿���Ϣ', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1124 And ���� = '�����˷�' And Upper(����) = Upper('�����˿���Ϣ'));        

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1124, 'ҽ������', User, 'Zl_�����˿���Ϣ_Insert', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1124 And ���� = 'ҽ������' And Upper(����) = Upper('Zl_�����˿���Ϣ_Insert'));

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1124, '�����˷�', User, 'Zl_�����˿���Ϣ_Insert', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1124 And ���� = '�����˷�' And Upper(����) = Upper('Zl_�����˿���Ϣ_Insert'));

--91427:����,2015-12-22,����ҩƷ����ϵͳ
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��,ȱʡֵ)
Select &n_System,1348,A.* From (
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0 Union All 
    Select '����',-Null,NULL,1 From Dual Union All      
    Select '����',2,'����ҩƷ������չ���Ĳ���Ȩ�ޡ��и�Ȩ��ʱ����������������յ�',1 From Dual Union All 
    Select '�޸�',4,'��δ��˵�ҩƷ�����޸ĵĲ���Ȩ�ޡ��и�Ȩ��ʱ�������δ��˵�������յ������޸�',1 From Dual Union All 
    Select 'ɾ��',5,'ɾ��ҩƷ������չ����¼�Ĳ���Ȩ�ޡ��и�Ȩ��ʱ�������δ��˵�������յ�����ɾ��',1 From Dual Union All 
    Select '���',6,'����ҩƷ�⹺�ſ��¼��˵Ĳ���Ȩ�ޡ��и�Ȩ��ʱ�������������յ��������',1 From Dual Union All     
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1348,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All 
    Select 'NextNO','EXECUTE' From Dual Union All 
    Select '������Ʊ�','SELECT' From Dual Union All 
    Select '������Ʊ�','UPDATE' From Dual Union All 
    Select '���Һ����','SELECT' From Dual Union All 
    Select '���Һ����','UPDATE' From Dual Union All
    Select '���ű�','SELECT' From Dual Union All 
    Select '������Ա','SELECT' From Dual Union All 
    Select '�������ʷ���','SELECT' From Dual Union All 
    Select '��������˵��','SELECT' From Dual Union All 
    Select '��Ӧ��','SELECT' From Dual Union All 
    Select 'ҩƷ������','SELECT' From Dual Union All 
    Select '��Ա��','SELECT' From Dual Union All 
    Select '�ϻ���Ա��','SELECT' From Dual Union All 
    Select '�շѼ�Ŀ','SELECT' From Dual Union All 
    Select '�շ�ϸĿ','SELECT' From Dual Union All 
    Select '�շ���Ŀ����','SELECT' From Dual Union All 
    Select '�շ���ĿĿ¼','SELECT' From Dual Union All 
    Select '�շ�ִ�п���','SELECT' From Dual Union All 
    Select 'ҩƷ����','SELECT' From Dual Union All 
    Select 'ҩƷ���ʷ���','SELECT' From Dual Union All 
    Select 'ҩƷ������','SELECT' From Dual Union All 
    Select 'ҩƷ��������','SELECT' From Dual Union All 
    Select 'ҩƷ���','SELECT' From Dual Union All 
    Select 'ҩƷ����','SELECT' From Dual Union All     
    Select 'ҩƷĿ¼','SELECT' From Dual Union All 
    Select 'ҩƷ������','SELECT' From Dual Union All    
    Select 'ҩƷ����','SELECT' From Dual Union All 
    Select 'ҩƷ���','SELECT' From Dual Union All 
    Select 'ҩƷ���ľ���','SELECT' From Dual Union All     
    Select '���Ʒ���Ŀ¼','SELECT' From Dual Union All 
    Select '������Ŀ���','SELECT' From Dual Union All 
    Select '������ĿĿ¼','SELECT' From Dual Union All 
    Select '����ִ�п���','SELECT' From Dual Union All 
    Select 'ҩƷ���ռ�¼','SELECT' From Dual Union All 
    Select 'ҩƷ������ϸ','SELECT' From Dual Union All 
    Select 'ҩƷ�����޶�','SELECT' From Dual Union All
    Select 'ҩƷ���ռ�¼_ID','SELECT' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1348,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All
Select 'Zl_ҩƷ���ռ�¼_Insert','EXECUTE' From Dual Union All
Select 'Zl_ҩƷ������ϸ_Insert','EXECUTE' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1348,'�޸�',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All
Select 'Zl_ҩƷ���ռ�¼_Insert','EXECUTE' From Dual Union All
Select 'ZL_ҩƷ���ռ�¼_Delete','EXECUTE' From Dual Union All
Select 'Zl_ҩƷ������ϸ_Insert','EXECUTE' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1348,'ɾ��',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All
Select 'ZL_ҩƷ���ռ�¼_Delete','EXECUTE' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1348,'���',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All
Select 'ZL_ҩƷ���ռ�¼_Verify','EXECUTE' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

--91225:������,2015-12-22,��Ⱦ������ϵͳ���Ȩ��
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��,ȱʡֵ)
Select &n_System,1278,A.* From (
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0 Union All
Select '����',-NULL,NULL,1 From Dual Union All
Select '��Χ����',1,'���ù���վ�ɹ���ļ������淶Χ���и�Ȩ��ʱ���������ñ�����վ�ɹ�����ļ�',1 From Dual Union All
Select '����',2,'��������Ķ��ⱨ����Ϣ�Ǽǡ��и�Ȩ��ʱ������Լ�������Ķ��ⱨ����Ϣ���еǼ�',1 From Dual Union All
Select '����',3,'ȡ������ı��͵Ǽǻ���վܾ��������и�Ȩ��ʱ������Լ�������ĵǼǽ��վܾ��������л���',1 From Dual Union All
Select '���',4,'�����д�˵ļ������档�и�Ȩ��ʱ������Լ�������Ľ������',1 From Dual Union All
Select 'ɾ��',5,'ɾ���ظ��ļ������档�и�Ȩ��ʱ������Լ����������ɾ��',1 From Dual Union All
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1278,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All
Select 'ZL_Replace_Element_Value','EXECUTE' From Dual Union All
Select 'Zl_Lob_Read','EXECUTE' From Dual Union All
Select 'Zl_���Ӳ�����ӡ_Insert','EXECUTE' From Dual Union All
Select 'Zl_�����걨��Ӧ_UPDATE','EXECUTE' From Dual Union All
Select '������ҳ','SELECT' From Dual Union All
Select '����������ʽ','SELECT' From Dual Union All
Select '�����ļ��ṹ','SELECT' From Dual Union All
Select '�����ļ��б�','SELECT' From Dual Union All
Select '����ҳ���ʽ','SELECT' From Dual Union All
Select '������Ϣ','SELECT' From Dual Union All
Select '����ҽ�ƿ�����','SELECT' From Dual Union All
Select '����ҽ�ƿ���Ϣ','SELECT' From Dual Union All
Select '��������˵��','SELECT' From Dual Union All
Select '���Ӳ�������','SELECT' From Dual Union All
Select '���Ӳ�����¼','SELECT' From Dual Union All
Select '���Ӳ�������','SELECT' From Dual Union All
Select '�������͵�λ','SELECT' From Dual Union All
Select '�����걨��Ӧ','SELECT' From Dual Union All
Select '�����걨��¼','SELECT' From Dual Union All
Select '�����ѽӿ�Ŀ¼','SELECT' From Dual Union All
Select '��Ա����˵��','SELECT' From Dual Union All
Select '���ѿ�Ŀ¼','SELECT' From Dual Union All
Select 'ҽ�ƿ���ʧ��ʽ','SELECT' From Dual Union All
Select 'ҽ�ƿ����','SELECT' From Dual Union All
Select '�������淴��','SELECT' From Dual Union All
Select '������Ƭ','SELECT' From Dual Union All
Select '�������Լ�¼','SELECT' From Dual Union All
Select '���˹Һż�¼','SELECT' From Dual Union All
Select '���ű�','SELECT' From Dual Union All
Select '��Ա��','SELECT' From Dual Union All
Select '������Ա','SELECT' From Dual Union All
Select 'ְҵ','SELECT' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1278,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All
Select 'Zl_�����걨��¼_Send','EXECUTE' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1278,'���',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All
Select 'Zl_�����걨��¼_Update','EXECUTE' From Dual Union All
Select 'Zl_�����걨��¼_Incept','EXECUTE' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1278,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All
Select 'Zl_�����걨��¼_Untread','EXECUTE' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1278,'ɾ��',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All
Select 'Zl_�����걨��¼_Delete','EXECUTE' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;


--91866:����,2015-12-21,��Ⱦ������ϵͳ���Խ���������Ǽǡ���ѯ����
Insert Into zlModuleRelas(ϵͳ,ģ��,����,���ϵͳ,���ģ��,�������,��ع���,ȱʡֵ)
Select &n_System,1290,A.* From (
Select ����,���ϵͳ,���ģ��,�������,��ع���,ȱʡֵ From zlModuleRelas Where 1 = 0 Union All
Select NULL,&n_System,9001,1,'����',1 From Dual Union All
Select ����,���ϵͳ,���ģ��,�������,��ع���,ȱʡֵ From zlModuleRelas Where 1 = 0) A;

Insert Into zlModuleRelas(ϵͳ,ģ��,����,���ϵͳ,���ģ��,�������,��ع���,ȱʡֵ)
Select &n_System,1291,A.* From (
Select ����,���ϵͳ,���ģ��,�������,��ع���,ȱʡֵ From zlModuleRelas Where 1 = 0 Union All
Select NULL,&n_System,9001,1,'����',1 From Dual Union All
Select ����,���ϵͳ,���ģ��,�������,��ع���,ȱʡֵ From zlModuleRelas Where 1 = 0) A;

Insert Into zlModuleRelas(ϵͳ,ģ��,����,���ϵͳ,���ģ��,�������,��ع���,ȱʡֵ)
Select &n_System,1294,A.* From (
Select ����,���ϵͳ,���ģ��,�������,��ع���,ȱʡֵ From zlModuleRelas Where 1 = 0 Union All
Select NULL,&n_System,9001,1,'����',1 From Dual Union All
Select ����,���ϵͳ,���ģ��,�������,��ع���,ȱʡֵ From zlModuleRelas Where 1 = 0) A;

--91225:������,2015-12-21,��Ⱦ������ϵͳ�����ӱ�
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��,ȱʡֵ)
Select  &n_System,9001,A.* From (
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0 Union All
Select '��Ⱦ�����Խ���Ǽ�',2,'�д�Ȩ��ʱ��������ýӿڶԴ�Ⱦ�����Խ�����еǼ�',1 From Dual Union All
Select '��Ⱦ�����Խ����ѯ',3,'�д�Ȩ��ʱ��������ýӿڶԴ�Ⱦ�����Խ�����в�ѯ',1 From Dual Union All
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select  &n_System,9001,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All 
    Select '�������Լ�¼','SELECT' From Dual Union All 
    Select '��Ⱦ��Ŀ¼','SELECT' From Dual Union All 
    Select '��������ǰ��','SELECT' From Dual Union All     
    Select '���Ƽ���걾','SELECT' From Dual Union All 
    Select 'Zl_�������Լ���¼_Insert','EXECUTE' From Dual Union All 
    Select 'Zl_�������Լ���¼_Update','EXECUTE' From Dual Union All 
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

--91225:������,2015-12-21,��Ⱦ�����Խ����������ѯ
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,9001,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All 
    Select '������Ƭ','SELECT' From Dual Union All    
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

--91738:���ջ�,2015-12-17,���Ӳ����������°�PACS����
Insert Into zlModuleRelas(ϵͳ,ģ��,����,���ϵͳ,���ģ��,�������,��ع���,ȱʡֵ)
Select &n_System,1560,A.* From (
Select ����,���ϵͳ,���ģ��,�������,��ع���,ȱʡֵ From zlModuleRelas Where 1 = 0 Union All
Select '����',&n_System,9004,1,'����',1 From Dual Union All
Select ����,���ϵͳ,���ģ��,�������,��ع���,ȱʡֵ From zlModuleRelas Where 1 = 0) A;

Insert Into Zltools.Zlrolegrant
 (ϵͳ, ���, ��ɫ, ����)
 Select Distinct &n_System, 9004, ��ɫ, '����'
 From Zltools.Zlrolegrant A
 Where ��� = 1560 And Not Exists (Select 1 From Zltools.Zlrolegrant Where ��� = 9004 And ��ɫ = a.��ɫ);

--89242:���ϴ�,2015-12-08,ʹ�ýṹ����ַ�ؼ�
Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1107, '����', User, 'zl_���˵�ַ��Ϣ_update', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1107 And ���� = '����' And Upper(����) = Upper('zl_���˵�ַ��Ϣ_update'));

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1111, '����', User, 'zl_���˵�ַ��Ϣ_update', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1111 And ���� = '����' And Upper(����) = Upper('zl_���˵�ַ��Ϣ_update'));

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1113, '�����޸�', User, 'zl_���˵�ַ��Ϣ_update', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1113 And ���� = '�����޸�' And Upper(����) = Upper('zl_���˵�ַ��Ϣ_update'));

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1107, '����', User, 'Zl_Adderss_Structure', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1107 And ���� = '����' And Upper(����) = Upper('Zl_Adderss_Structure'));

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1111, '����', User, 'Zl_Adderss_Structure', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1111 And ���� = '����' And Upper(����) = Upper('Zl_Adderss_Structure'));

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1113, '�����޸�', User, 'Zl_Adderss_Structure', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1113 And ���� = '�����޸�' And Upper(����) = Upper('Zl_Adderss_Structure'));

--89620:��ΰ��,2016-01-15,��ǰ����ٴ�·��ִ�е�Ȩ��
Insert Into zlProgRelas(ϵͳ,���,���,����,��ϵ,����,�����ϵ)
Select &n_System,1256,2,A.* From (
Select ����,��ϵ,����,�����ϵ From zlProgRelas Where 1 = 0 Union All
Select '����·��',2,1,0 From Dual Union All
Select '��ǰ���',2,0,0 From Dual Union All
Select ����,��ϵ,����,�����ϵ From zlProgRelas Where 1 = 0) A;






-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------

--91225:���Ʊ�,2016-01-19,��Ⱦ������ϵͳ
--����ZL1_REPORT_1280/������Ⱦ������ǼǱ�
Insert Into zlReports(ID,���,����,˵��,����,��ӡ��,��ֽ,Ʊ��,��ӡ��ʽ,ϵͳ,����ID,����,�޸�ʱ��,����ʱ��,��ֹ��ʼʱ��,��ֹ����ʱ��) Values(zlReports_ID.NextVal,'ZL1_REPORT_1280','������Ⱦ������ǼǱ�','��ѯ��չʾ��Ⱦ������ǼǼ�¼','H`;~@e`~{( PlscuZ,\L','Microsoft XPS Document Writer',15,0,0,100,1280,'����',Sysdate,Sysdate,Null,Null);
Insert Into zlRPTFmts(����ID,���,˵��,W,H,ֽ��,ֽ��,��ֽ̬��,ͼ��) Values(zlReports_ID.CurrVal,1,'������Ⱦ������ǼǱ�',11906,16838,9,2,0,0);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����,˵��) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'�ǼǼ�¼','ROWNUM,139|����,202|�Ա�,202|��������,202|ְҵ,202|��ͥ��ַ,202|�绰,202|��������,202|ȷ������,202|ʵ��,130|�ٴ�,130|Я��,130|����,130|���,202|�������,202|�����,202|���,202|�տ�����,202|�տ���,202|���籨������,202|��ע,202',User||'.���Ӳ�����¼,'||User||'.�����걨��¼,'||User||'.������Ϣ,'||User||'.���ű�',0,Null);
Insert Into zlRPTSQLs(ԴID,�к�,����)  Select zlRPTDatas_ID.CurrVal,a.* From (
  Select 1,'Select Rownum, ����, �Ա�, ��������, ְҵ, ��ͥ��ַ, �绰, ��������, ȷ������, ʵ��, �ٴ�, Я��, ����, ���, �������, �����, ���, �տ�����, �տ���, ���籨������, ��ע' From Dual Union All
  Select 2,'From (Select p.����, p.�Ա�, to_char(p.��������,''yyyy-mm-dd'') ��������, p.ְҵ, p.��ͥ��ַ, p.��ϵ�˵绰 �绰, to_char(s.��������,''yyyy-mm-dd'') ��������, to_char(s.ȷ������,''yyyy-mm-dd'')ȷ������, '''' ʵ��, '''' �ٴ�, '''' Я��, '''' ����,' From Dual Union All
  Select 3,'              s.�������1 || s.�������2 ���, d.���� �������, to_char(l.���ʱ��,''yyyy-mm-dd'') �����, l.������ ���, to_char(Trunc(s.�վ�ʱ��),''yyyy-mm-dd'') �տ�����, s.�վ��� �տ���,' From Dual Union All
  Select 4,'              to_char(Trunc(s.����ʱ��),''yyyy-mm-dd'') ���籨������, s.���ע ��ע' From Dual Union All
  Select 5,'       From ���Ӳ�����¼ L, �����걨��¼ S, ������Ϣ P, ���ű� D' From Dual Union All
  Select 6,'       Where l.�������� = 5 And l.�ļ�id In ([0]) And l.���ʱ�� Between [1] And [2] And' From Dual Union All
  Select 7,'             l.Id = s.�ļ�id(+) And s.����״̬(+) <> -1 And l.����id = p.����id And l.����id = d.Id' From Dual Union All
  Select 8,'       Order By p.����)' From Dual Union All
  Select 9,Null From Dual) a;
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����,����) Values(zlRPTDatas_ID.CurrVal,Null,0,'�ļ�',0,'ѡ�������塭',0,Null,Null,'select id ,���� From �����ļ��б� where ����=5',Null,'ID,131,'||CHR(38)||'S'||CHR(38)||'B|����,202,'||CHR(38)||'D'||CHR(38)||'S',User||'.�����ļ��б�|',Null);
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����,����) Values(zlRPTDatas_ID.CurrVal,Null,1,'��ʼʱ��',2,CHR(38)||'���³�ʱ��',0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����,����) Values(zlRPTDatas_ID.CurrVal,Null,2,'����ʱ��',2,CHR(38)||'����ĩʱ��',0,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ2',2,Null,0,'�����1',11,'ͳ��ʱ��:[=��ʼʱ��]  -  [=����ʱ��]',Null,90,615,3780,210,0,0,1,'����',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ1',2,Null,0,'�����1',12,'������Ⱦ������ǼǱ�',Null,6720,120,3300,330,0,0,1,'����',16,1,0,0,0,16777215,0,Null,Null,Null,1,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'�����1',4,Null,0,Null,0,'�ǼǼ�¼',Null,90,975,16560,10470,450,0,1,'����',9,0,0,0,0,16777215,1,Null,Null,Null,1,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-1,0,Null,Null,'[�ǼǼ�¼.ROWNUM]','4^450^���',0,0,315,0,0,2,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-2,1,Null,Null,'[�ǼǼ�¼.����]','4^450^����',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-3,2,Null,Null,'[�ǼǼ�¼.�Ա�]','4^450^�Ա�',0,0,315,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-4,3,Null,Null,'[�ǼǼ�¼.��������]','4^450^��������',0,0,990,0,0,1,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-5,4,Null,Null,'[�ǼǼ�¼.ְҵ]','4^450^ְҵ',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-6,5,Null,Null,'[�ǼǼ�¼.��ͥ��ַ]','4^450^��ͥ��ַ',0,0,1245,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-7,6,Null,Null,'[�ǼǼ�¼.�绰]','4^450^�绰',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-8,7,Null,Null,'[�ǼǼ�¼.��������]','4^450^��������',0,0,1005,0,0,1,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-9,8,Null,Null,'[�ǼǼ�¼.ȷ������]','4^450^ȷ������',0,0,1065,0,0,1,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-10,9,Null,Null,'[�ǼǼ�¼.ʵ��]','4^450^ʵ��',0,0,285,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-11,10,Null,Null,'[�ǼǼ�¼.�ٴ�]','4^450^�ٴ�',0,0,270,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-12,11,Null,Null,'[�ǼǼ�¼.Я��]','4^450^Я��',0,0,255,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-13,12,Null,Null,'[�ǼǼ�¼.����]','4^450^����',0,0,285,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-14,13,Null,Null,'[�ǼǼ�¼.���]','4^450^���',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-15,14,Null,Null,'[�ǼǼ�¼.�������]','4^450^�������',0,0,825,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-16,15,Null,Null,'[�ǼǼ�¼.�����]','4^450^�����',0,0,990,0,0,1,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-17,16,Null,Null,'[�ǼǼ�¼.���]','4^450^���',0,0,840,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-18,17,Null,Null,'[�ǼǼ�¼.�տ�����]','4^450^�տ�����',0,0,990,0,0,1,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-19,18,Null,Null,'[�ǼǼ�¼.�տ���]','4^450^�տ���',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-20,19,Null,Null,'[�ǼǼ�¼.���籨������]','4^450^���籨������',0,0,810,0,0,1,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-21,20,Null,Null,'[�ǼǼ�¼.��ע]','4^450^��ע',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);

--����ZL1_REPORT_1280/������Ⱦ������ǼǱ�
Insert into zlPrograms(���,����,˵��,ϵͳ,����) Values(1280,'������Ⱦ������ǼǱ�','��ѯ��չʾ��Ⱦ������ǼǼ�¼',100,'zl9Report');
Insert into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1280,'����',Null);
Insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
  Select 100,1280,'����',User,'�����ļ��б�','SELECT' From Dual Union All
  Select 100,1280,'����',User,'������Ϣ','SELECT' From Dual Union All
  Select 100,1280,'����',User,'���ű�','SELECT' From Dual Union All
  Select 100,1280,'����',User,'���Ӳ�����¼','SELECT' From Dual Union All
  Select 100,1280,'����',User,'�����걨��¼','SELECT' From Dual;
Insert into zlMenus(���,ID,�ϼ�ID,����,�̱���,���,ͼ��,˵��,ϵͳ,ģ��) Select 'ȱʡ',zlMenus_ID.Nextval,ID,'������Ⱦ������ǼǱ�','������Ⱦ������ǼǱ�',Null,105,'��ѯ��չʾ��Ⱦ������ǼǼ�¼',100,1280 From zlMenus Where ϵͳ=100 And ���='ȱʡ' And ����='��Ⱦ������ϵͳ' And ģ�� is NULL;

--����ZL1_REPORT_1281/��Ⱦ�����Լ����һ����
Insert Into zlReports(ID,���,����,˵��,����,��ӡ��,��ֽ,Ʊ��,��ӡ��ʽ,ϵͳ,����ID,����,�޸�ʱ��,����ʱ��,��ֹ��ʼʱ��,��ֹ����ʱ��) Values(zlReports_ID.NextVal,'ZL1_REPORT_1281','��Ⱦ�����Լ����һ����','��ѯ��Ⱦ�����Լ����','Mv:uZldpv3%Fmxx}^"QW',Null,15,0,0,100,1281,'����',Sysdate,Sysdate,To_Date('2016-01-01 00:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('2016-01-01 00:00:00','YYYY-MM-DD HH24:MI:SS'));
Insert Into zlRPTFmts(����ID,���,˵��,W,H,ֽ��,ֽ��,��ֽ̬��,ͼ��) Values(zlReports_ID.CurrVal,1,'��Ⱦ�����Լ����һ����1',11904,16832,256,1,0,0);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����,˵��) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'�������Լ�¼_����','ID,131|��Դ,130|����ID,131|����,202|�Ա�,202|����,202|����,202|��ʶ��,131|�ͼ�ʱ��,202|�ͼ�ҽ��,202|�ͼ����,202|�걾����,202|�������,202|���Ƽ���,202|�Ǽ���,202|�Ǽ�ʱ��,202|������,202|����ʱ��,202|�������˵��,202',User||'.�������Լ�¼,'||User||'.������ҳ,'||User||'.���˹Һż�¼,'||User||'.���ű�',0,Null);
Insert Into zlRPTSQLs(ԴID,�к�,����)  Select zlRPTDatas_ID.CurrVal,a.* From (
  Select 1,'select A.Id, A.��Դ, ����id, A.����, A.�Ա�,A.����,e.���� as ����, A.��ʶ��,A.�ͼ�ʱ��, A.�ͼ�ҽ��,f.���� As �ͼ����,A.�걾����, A.�������,  A.���Ƽ���, A.�Ǽ���,A.�Ǽ�ʱ��,A.������, A.����ʱ��, A.�������˵��' From Dual Union All
  Select 2,'from ' From Dual Union All
  Select 3,'(Select a.Id,  ''סԺ'' As ��Դ, a.����id, c.����, c.�Ա�,c.����,' From Dual Union All
  Select 4,'      C.��Ժ����id as ����ID, c.סԺ�� As ��ʶ��,To_Char(a.�ͼ�ʱ��, ''yyyy-MM-dd hh24:mi'')  �ͼ�ʱ��, a.�ͼ�ҽ��, a.�ͼ����id, a.�걾����, a.�������, a.��Ⱦ������ As ���Ƽ���, a.�Ǽ���,' From Dual Union All
  Select 5,'       To_Char(a.�Ǽ�ʱ��, ''yyyy-MM-dd hh24:mi'') �Ǽ�ʱ��, a.������, To_Char(a.����ʱ��, ''yyyy-MM-dd hh24:mi'')  ����ʱ��, a.�������˵��' From Dual Union All
  Select 6,'From �������Լ�¼ A, ������ҳ C' From Dual Union All
  Select 7,'Where a.����id = c.����id And a.��ҳid = c.��ҳid and a.�Ǽ�ʱ�� Between [0] And [1]' From Dual Union All
  Select 8,'union all' From Dual Union All
  Select 9,'Select a.Id,  ''����'' As ��Դ, a.����id,  b.���� , b.�Ա� , b.����,' From Dual Union All
  Select 10,'        b.ִ�в���id as ����ID , b.����� As ��ʶ��,To_Char(a.�ͼ�ʱ��, ''yyyy-MM-dd hh24:mi'')  �ͼ�ʱ��, a.�ͼ�ҽ��, a.�ͼ����ID, a.�걾����, a.�������, a.��Ⱦ������ As ���Ƽ���, a.�Ǽ���,' From Dual Union All
  Select 11,'       To_Char(a.�Ǽ�ʱ��, ''yyyy-MM-dd hh24:mi'') �Ǽ�ʱ��, a.������, To_Char(a.����ʱ��, ''yyyy-MM-dd hh24:mi'')  ����ʱ��, a.�������˵��' From Dual Union All
  Select 12,'From �������Լ�¼ A, ���˹Һż�¼ B' From Dual Union All
  Select 13,'Where  a.����id = b.����id And a.�Һŵ� = b.No And a.�Ǽ�ʱ�� Between [0] And [1]) A ,���ű� E, ���ű� F' From Dual Union All
  Select 14,'where  a.�ͼ����id = f.Id(+) And  A.����ID = e.Id(+)' From Dual Union All
  Select 15,'Order By a.Id' From Dual Union All
  Select 16,Null From Dual) a;
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����,����) Values(zlRPTDatas_ID.CurrVal,Null,0,'��ʼ����',2,CHR(38)||'���³�ʱ��',0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����,����) Values(zlRPTDatas_ID.CurrVal,Null,1,'��������',2,CHR(38)||'����ĩʱ��',0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����Ա',2,Null,0,'�����1',21,'ͳ���ˣ�[����Ա����]',Null,435,14685,2100,225,0,0,1,'����',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,Null,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��������',2,Null,0,'�����1',12,'[��λ����]��Ⱦ�����Լ����һ����',Null,2778,390,6120,375,0,0,1,'����',18,0,0,0,0,16777215,0,Null,Null,Null,1,0,1,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ӡʱ��',2,Null,0,'�����1',23,'��ӡʱ�䣺[yyyy-MM-dd hh:mm:ss]',Null,7985,14670,3255,225,0,0,1,'����',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,Null,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ1',2,Null,0,'�����1',13,'���ڣ�[=��ʼ����]��[=��������]',Null,8090,1080,3150,225,0,0,1,'����',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,Null,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'�����1',4,Null,0,Null,0,'�������Լ�¼_����',Null,435,1500,10805,12960,255,0,0,'����',10.5,0,0,0,0,16777215,1,Null,Null,Null,1,0,Null,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-1,0,Null,Null,'[�������Լ�¼_����.��Դ]','4^225^��Դ',0,0,1005,0,0,1,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-2,1,Null,Null,'[�������Լ�¼_����.����]','4^225^����',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-3,2,Null,Null,'[�������Լ�¼_����.�Ա�]','4^225^�Ա�',0,0,1005,0,0,1,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-4,3,Null,Null,'[�������Լ�¼_����.����]','4^225^����',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-5,4,Null,Null,'[�������Լ�¼_����.����]','4^225^����',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-6,5,Null,Null,'[�������Լ�¼_����.��ʶ��]','4^225^��ʶ��',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-7,6,Null,Null,'[�������Լ�¼_����.�ͼ�ʱ��]','4^225^�ͼ�ʱ��',0,0,2040,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-8,7,Null,Null,'[�������Լ�¼_����.�ͼ�ҽ��]','4^225^�ͼ�ҽ��',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-9,8,Null,Null,'[�������Լ�¼_����.�ͼ����]','4^225^�ͼ����',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-10,9,Null,Null,'[�������Լ�¼_����.�걾����]','4^225^�걾����',0,0,1155,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-11,10,Null,Null,'[�������Լ�¼_����.�������]','4^225^�������',0,0,1785,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-12,11,Null,Null,'[�������Լ�¼_����.���Ƽ���]','4^225^���Ƽ���',0,0,1365,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-13,12,Null,Null,'[�������Լ�¼_����.�Ǽ���]','4^225^�Ǽ���',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-14,13,Null,Null,'[�������Լ�¼_����.�Ǽ�ʱ��]','4^225^�Ǽ�ʱ��',0,0,2010,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-15,14,Null,Null,'[�������Լ�¼_����.������]','4^225^������',0,0,1005,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-16,15,Null,Null,'[�������Լ�¼_����.����ʱ��]','4^225^����ʱ��',0,0,2010,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,6,zlRPTItems_ID.CurrVal-17,16,Null,Null,'[�������Լ�¼_����.�������˵��]','4^225^�������˵��',0,0,2205,0,0,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,0,0,Null,Null,Null,Null,Null,Null,Null,Null);

--����ZL1_REPORT_1281/��Ⱦ�����Լ����һ����
Insert into zlPrograms(���,����,˵��,ϵͳ,����) Values(1281,'��Ⱦ�����Լ����һ����','��ѯ��Ⱦ�����Լ����',100,'zl9Report');
Insert into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1281,'����',Null);
Insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
  Select 100,1281,'����',User,'������ҳ','SELECT' From Dual Union All
  Select 100,1281,'����',User,'���˹Һż�¼','SELECT' From Dual Union All
  Select 100,1281,'����',User,'���ű�','SELECT' From Dual Union All
  Select 100,1281,'����',User,'�������Լ�¼','SELECT' From Dual;
Insert into zlMenus(���,ID,�ϼ�ID,����,�̱���,���,ͼ��,˵��,ϵͳ,ģ��) Select 'ȱʡ',zlMenus_ID.Nextval,ID,'��Ⱦ�����Լ����һ����','��Ⱦ�����Լ����һ����',Null,105,'��ѯ��Ⱦ�����Լ����',100,1281 From zlMenus Where ϵͳ=100 And ���='ȱʡ' And ����='��Ⱦ������ϵͳ' And ģ�� is NULL;

--����ZL1_REPORT_1282/��Ⱦ����������ܱ�
Insert Into zlReports(ID,���,����,˵��,����,��ӡ��,��ֽ,Ʊ��,��ӡ��ʽ,ϵͳ,����ID,����,�޸�ʱ��,����ʱ��,��ֹ��ʼʱ��,��ֹ����ʱ��) Values(zlReports_ID.NextVal,'ZL1_REPORT_1282','��Ⱦ����������ܱ�','������Դ�Ⱦ�����л���','Mv:jLio`s)4ViooG*U\',Null,15,0,0,100,1282,'����',Sysdate,Sysdate,To_Date('2016-01-01 00:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('2016-01-01 00:00:00','YYYY-MM-DD HH24:MI:SS'));
Insert Into zlRPTFmts(����ID,���,˵��,W,H,ֽ��,ֽ��,��ֽ̬��,ͼ��) Values(zlReports_ID.CurrVal,1,'��Ⱦ���������Ա���ܱ�1',11904,16832,9,1,0,0);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����,˵��) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'�����걨��¼_����','����,202|��Ⱦ������,202|��,139|Ů,139|��,139',User||'.�����걨��¼,'||User||'.�������Լ�¼,'||User||'.�������淴��',1,Null);
Insert Into zlRPTSQLs(ԴID,�к�,����)  Select zlRPTDatas_ID.CurrVal,a.* From (
  Select 1,'Select nvl(a.����, ''δ֪����'') as ����,  b.��Ⱦ������, sum(decode(A.�Ա�, ''��'',1,0)) as ��,sum(decode(A.�Ա�, ''Ů'',1,0)) as Ů,sum(decode(A.�Ա�,''��'',1,1)) as �� ' From Dual Union All
  Select 2,'From �����걨��¼ A, �������Լ�¼ B,�������淴�� C ' From Dual Union All
  Select 3,'Where a.�ļ�id = b.�ļ�id  and A.�ļ�ID = C.�ļ�ID  and c.�Ǽ�ʱ�� Between [0] And [1]' From Dual Union All
  Select 4,'Group By nvl(a.����, ''δ֪����''),b.��Ⱦ������' From Dual) a;
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����,����) Values(zlRPTDatas_ID.CurrVal,Null,0,'��ʼ����',2,CHR(38)||'���³�ʱ��',0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����,����) Values(zlRPTDatas_ID.CurrVal,Null,1,'��������',2,CHR(38)||'����ĩʱ��',0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'ͳ����',2,Null,0,'���ܱ�1',21,'ͳ���ˣ�[����Ա����]',Null,825,15790,2100,225,0,0,1,'����',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,Null,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����',2,Null,0,'���ܱ�1',12,'[��λ����]��Ⱦ����������ܱ�',Null,2975,615,6105,450,0,0,1,'����',22,0,0,0,0,16777215,0,Null,Null,Null,1,0,1,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'ͳ��ʱ��',2,Null,0,'���ܱ�1',23,'ͳ��ʱ�䣺[yyyy-mm-dd HH:MM:SS]',Null,7975,15790,3255,225,0,0,1,'����',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,Null,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����',2,Null,0,'���ܱ�1',13,'���ڣ�[=��ʼ����]��[=��������]',Null,8080,1575,3150,225,0,0,1,'����',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,Null,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'���ܱ�1',5,Null,0,Null,0,'�����걨��¼_����',Null,825,1965,10405,13605,255,0,0,'����',10.5,0,0,0,0,16777215,1,Null,Null,Null,1,0,Null,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,7,zlRPTItems_ID.CurrVal-1,0,Null,Null,'����',Null,0,0,1005,0,255,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,8,zlRPTItems_ID.CurrVal-2,0,Null,Null,'��Ⱦ������',Null,0,0,1000,0,255,0,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-3,0,Null,Null,'��',Null,0,0,1005,0,255,2,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-4,1,Null,Null,'Ů',Null,0,0,1005,0,255,2,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-5,2,Null,Null,'��',Null,0,0,1005,0,255,2,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,Null,Null,Null,Null,Null,Null,Null,Null);

--����ZL1_REPORT_1282/��Ⱦ����������ܱ�
Insert into zlPrograms(���,����,˵��,ϵͳ,����) Values(1282,'��Ⱦ����������ܱ�','������Դ�Ⱦ�����л���',100,'zl9Report');
Insert into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1282,'����',Null);
Insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
  Select 100,1282,'����',User,'�������淴��','SELECT' From Dual Union All
  Select 100,1282,'����',User,'�����걨��¼','SELECT' From Dual Union All
  Select 100,1282,'����',User,'�������Լ�¼','SELECT' From Dual;
Insert into zlMenus(���,ID,�ϼ�ID,����,�̱���,���,ͼ��,˵��,ϵͳ,ģ��) Select 'ȱʡ',zlMenus_ID.Nextval,ID,'��Ⱦ����������ܱ�','��Ⱦ����������ܱ�',Null,105,'������Դ�Ⱦ�����л���',100,1282 From zlMenus Where ϵͳ=100 And ���='ȱʡ' And ����='��Ⱦ������ϵͳ' And ģ�� is NULL;

--����ZL1_REPORT_1283/��Ⱦ����ְҵ���ܱ�
Insert Into zlReports(ID,���,����,˵��,����,��ӡ��,��ֽ,Ʊ��,��ӡ��ʽ,ϵͳ,����ID,����,�޸�ʱ��,����ʱ��,��ֹ��ʼʱ��,��ֹ����ʱ��) Values(zlReports_ID.NextVal,'ZL1_REPORT_1283','��Ⱦ����ְҵ���ܱ�','��ְҵ�Դ�Ⱦ�����л���','Mv:jX}o`s)4Vi{{G*U\',Null,15,0,0,100,1283,'����',Sysdate,Sysdate,To_Date('2016-01-01 00:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('2016-01-01 00:00:00','YYYY-MM-DD HH24:MI:SS'));
Insert Into zlRPTFmts(����ID,���,˵��,W,H,ֽ��,ֽ��,��ֽ̬��,ͼ��) Values(zlReports_ID.CurrVal,1,'��Ⱦ����ְҵ���ܱ�1',11904,16832,9,1,0,0);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����,˵��) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'����ҽ����¼_����','ְҵ,202|��Ⱦ������,202|��,139|Ů,139|��,139',User||'.�����걨��¼,'||User||'.�������Լ�¼,'||User||'.�������淴��',1,Null);
Insert Into zlRPTSQLs(ԴID,�к�,����)  Select zlRPTDatas_ID.CurrVal,a.* From (
  Select 1,'Select nvl(a.ְҵ,''����'') as ְҵ,  b.��Ⱦ������, sum(decode(A.�Ա�, ''��'',1,0)) as ��,sum(decode(A.�Ա�, ''Ů'',1,0)) as Ů,sum(decode(A.�Ա�,''��'',1,1)) as �� ' From Dual Union All
  Select 2,'From �����걨��¼ A, �������Լ�¼ B,�������淴�� C ' From Dual Union All
  Select 3,'Where a.�ļ�id = b.�ļ�id  and A.�ļ�ID = C.�ļ�ID  and c.�Ǽ�ʱ�� Between [0] And [1]' From Dual Union All
  Select 4,'Group By nvl(a.ְҵ,''����''),b.��Ⱦ������' From Dual) a;
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����,����) Values(zlRPTDatas_ID.CurrVal,Null,0,'��ʼ����',2,CHR(38)||'���³�ʱ��',0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����,����) Values(zlRPTDatas_ID.CurrVal,Null,1,'��������',2,CHR(38)||'����ĩʱ��',0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����Ա',2,Null,0,'���ܱ�1',21,'ͳ���ˣ�[����Ա����]',Null,645,15600,2100,225,0,0,1,'����',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,Null,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ1',2,Null,0,'���ܱ�1',12,'[��λ����]��Ⱦ����ְҵ���ܱ�',Null,2705,570,6105,450,0,0,1,'����',22,0,0,0,0,16777215,0,Null,Null,Null,1,0,1,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����',2,Null,0,'���ܱ�1',13,'���ڣ�[=��ʼ����]��[=��������]',Null,7720,1350,3150,225,0,0,1,'����',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,Null,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ2',2,Null,0,'���ܱ�1',23,'��ӡʱ��:[yyyy-mm-dd HH:MM]',Null,8035,15585,2835,225,0,0,1,'����',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,Null,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'���ܱ�1',5,Null,0,Null,0,'����ҽ����¼_����',Null,645,1740,10225,13470,255,0,0,'����',10.5,0,0,0,0,16777215,1,Null,Null,Null,1,0,Null,Null,Null,0,0,0,0,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,7,zlRPTItems_ID.CurrVal-1,0,Null,Null,'ְҵ',Null,0,0,1005,0,255,0,0,'����',0,0,0,0,0,0,0,Null,Null,'SUM',1,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,8,zlRPTItems_ID.CurrVal-2,0,Null,Null,'��Ⱦ������',Null,0,0,1000,0,255,0,0,'����',0,0,0,0,0,0,0,Null,Null,'SUM',1,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-3,0,Null,Null,'��',Null,0,0,1335,0,255,2,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-4,1,Null,Null,'Ů',Null,0,0,1005,0,255,2,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,Null,Null,Null,Null,Null,Null,Null,Null);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,��ID,ԴID,���¼��,���Ҽ��,�������,�������,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,Null,9,zlRPTItems_ID.CurrVal-5,2,Null,Null,'��',Null,0,0,1005,0,255,2,0,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,Null,Null,Null,Null,Null,Null,Null,Null);

--����ZL1_REPORT_1283/��Ⱦ����ְҵ���ܱ�
Insert into zlPrograms(���,����,˵��,ϵͳ,����) Values(1283,'��Ⱦ����ְҵ���ܱ�','��ְҵ�Դ�Ⱦ�����л���',100,'zl9Report');
Insert into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1283,'����',Null);
Insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
  Select 100,1283,'����',User,'�������淴��','SELECT' From Dual Union All
  Select 100,1283,'����',User,'�����걨��¼','SELECT' From Dual Union All
  Select 100,1283,'����',User,'�������Լ�¼','SELECT' From Dual;
Insert into zlMenus(���,ID,�ϼ�ID,����,�̱���,���,ͼ��,˵��,ϵͳ,ģ��) Select 'ȱʡ',zlMenus_ID.Nextval,ID,'��Ⱦ����ְҵ���ܱ�','��Ⱦ����ְҵ���ܱ�',Null,105,'��ְҵ�Դ�Ⱦ�����л���',100,1283 From zlMenus Where ϵͳ=100 And ���='ȱʡ' And ����='��Ⱦ������ϵͳ' And ģ�� is NULL;






-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--92736:������,2016-01-19,���ʽӿ��޸�
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
                          Where a.��¼״̬ = 1 And a.����id = c.����id And a.����id = b.����id And b.��ҳid = c.��ҳid And a.����id = n_����id And
                                d.����id = a.Id And c.��Ժ����id = e.Id(+) And Exists
                           (Select 1 From ����Ԥ����¼ Where ����id = a.Id And ���㷽ʽ = v_���㿨���)
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
      If n_������� - n_���ʽ�� > 0 Then
        v_Temp := '<TBLX>' || 2 || '</TBLX>';
      Else
        v_Temp := '<TBLX>' || 1 || '</TBLX>';
      End If;
      v_Temp := v_Temp || '<TBJE>' || Abs(n_������� - n_���ʽ��) || '</TBJE>';
      v_Temp := '<TBQK>' || v_Temp || '</TBQK>';
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    End If;
  End If;
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getsettlement;
/

--92818:��˶,2016-01-18,���һ����Ժ������ҳID=NUll�Ľṹ����ַ
Create Or Replace Procedure Zl_���˵�ַ��Ϣ_Update
(
  ����_In     Number,
  ����id_In   ���˵�ַ��Ϣ.����id%Type,
  ��ҳid_In   ���˵�ַ��Ϣ.��ҳid%Type,
  ��ַ���_In ���˵�ַ��Ϣ.��ַ���%Type,
  ʡ_In       ���˵�ַ��Ϣ.ʡ%Type := Null,
  ��_In       ���˵�ַ��Ϣ.��%Type := Null,
  ��_In       ���˵�ַ��Ϣ.��%Type := Null,
  ����_In     ���˵�ַ��Ϣ.����%Type := Null,
  ����_In     ���˵�ַ��Ϣ.����%Type := Null,
  ��������_In ���˵�ַ��Ϣ.��������%Type := Null
) Is
  --���ܣ���ҳ�����нṹ�����˵�ַ��Ϣ���� 
  --����������_In 1-����,�޸�   2-ɾ�� 
  d_��Ժ���� ������ҳ.��Ժ����%Type;
  n_Count    Number(3);
Begin
  If ����_In = 1 Then
    Update ���˵�ַ��Ϣ
    Set ʡ = ʡ_In, �� = ��_In, �� = ��_In, ���� = ����_In, ���� = ����_In, �������� = ��������_In
    Where ����id = ����id_In And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0) And ��ַ��� = ��ַ���_In;
    If Sql%Rowcount = 0 Then
      Insert Into ���˵�ַ��Ϣ
        (����id, ��ҳid, ��ַ���, ʡ, ��, ��, ����, ����, ��������)
      Values
        (����id_In, ��ҳid_In, ��ַ���_In, ʡ_In, ��_In, ��_In, ����_In, ����_In, ��������_In);
    End If;
    --����ҳID�ǲ������һ���ڸ�Ժ����������ҳID=Null������
    If Not ��ҳid_In Is Null Then
      Select ��Ժ���� Into d_��Ժ���� From ������ҳ Where ����id = ����id_In And ��ҳid = ��ҳid_In;
      --���ڳ�Ժʱ�䣬���жϸó�Ժ���Ƿ���ھ����סԺ����
      If Not d_��Ժ���� Is Null Then
        --���ж�סԺ
        Select Count(1) Into n_Count From ������ҳ Where ����id = ����id_In And ��Ժ���� >= d_��Ժ����;
        If n_Count = 0 Then
          Begin
            --�ù��̲�������׼����С�����ϵͳ��������װû�в��˹Һż�¼
            Execute Immediate 'Select Count(1) From ���˹Һż�¼ Where ����id =:1  And �Ǽ�ʱ�� >=:2 '
              Into n_Count
              Using ����id_In, d_��Ժ����;
          Exception
            When Others Then
              Null;
          End;
        End If;
      End If;
      If d_��Ժ���� Is Null Or Nvl(n_Count, 0) = 0 Then
        Update ���˵�ַ��Ϣ
        Set ʡ = ʡ_In, �� = ��_In, �� = ��_In, ���� = ����_In, ���� = ����_In, �������� = ��������_In
        Where ����id = ����id_In And ��ҳid Is Null And ��ַ��� = ��ַ���_In;
      End If;
    End If;
  Else
    Delete From ���˵�ַ��Ϣ
    Where ����id = ����id_In And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0) And ��ַ��� = ��ַ���_In;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_���˵�ַ��Ϣ_Update;
/

--92335:���ϴ�,2016-01-18,����֧����ģʽ�����̲��
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
  ���½������_In  Number := 0--�Ƿ������Ա��������Ҫ�Ǵ���ͳһ����Ա��¼��̨�������������
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
       ���ʽ��, �ɿ���id, ����)
    Values
      (v_����id, 5, 1, ���ݺ�_In, ҽ�ƿ���_In, 1, ����id_In, Decode(��ҳid_In, 0, Null, ��ҳid_In),
       Decode(���˲���id_In, 0, Null, ���˲���id_In), Decode(���˿���id_In, 0, Null, ���˿���id_In), Decode(��ʶ��_In, 0, Null, ��ʶ��_In),
       ����_In, �Ա�_In, ����_In, �ѱ�_In, Decode(���㷽ʽ_In, Null, 1, 0), 3, �Ӱ��־_In, ��������id_In, ����Ա����_In, ����Ա���_In, ����Ա����_In,
       ����ʱ��_In, ����ʱ��_In, �շ�ϸĿid_In, �շ����_In, ���㵥λ_In, 1, 1, ҽ�ƿ���_In, ��������_In, ִ�в���id_In, ������Ŀid_In, �վݷ�Ŀ_In, ��׼����_In,
       Ӧ�ս��_In, ʵ�ս��_In, v_����id, Decode(���㷽ʽ_In, Null, Null, ʵ�ս��_In), n_��id, �����id_In);
  
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

Create Or Replace Procedure Zl_���˹Һż�¼_Insert
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
  ���½������_In  Number := 0--�Ƿ������Ա��������Ҫ�Ǵ���ͳһ����Ա��¼��̨�����������
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
    Select *
    From (Select a.Id, a.��¼״̬, a.Ԥ�����, a.No, Nvl(a.���, 0) As ���
           From ����Ԥ����¼ A,
                (Select NO, Sum(Nvl(a.���, 0)) As ���
                  From ����Ԥ����¼ A
                  Where a.����id Is Null And Nvl(a.���, 0) <> 0 And a.����id = v_����id And Nvl(a.Ԥ�����, 2) = 1
                  Group By NO
                  Having Sum(Nvl(a.���, 0)) <> 0) B
           Where a.����id Is Null And Nvl(a.���, 0) <> 0 And a.���㷽ʽ Not In (Select ���� From ���㷽ʽ Where ���� = 5) And
                 a.No = b.No And a.����id = v_����id And Nvl(a.Ԥ�����, 2) = 1
           Union All
           Select 0 As ID, ��¼״̬, Ԥ�����, NO, Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) As ���
           From ����Ԥ����¼
           Where ��¼���� In (1, 11) And ����id Is Not Null And Nvl(���, 0) <> Nvl(��Ԥ��, 0) And ����id = v_����id And
                 Nvl(Ԥ�����, 2) = 1 Having Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0
           Group By ��¼״̬, NO, Ԥ�����)
    Order By ID, NO;
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
  Begin
    Delete From �Һ����״̬
    Where ���� = �ű�_In And ���� = ����ʱ��_In And ��� = ����_In And ״̬ = 3 And ����Ա���� = ����Ա����_In;
  Exception
    When Others Then
      Null;
  End;
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
          --�Һ�ʱ���       
          --������ſ���δ��ʱ�� �ﵽ������
          If n_�޺��� <= n_�������� And n_�޺��� > 0 Then
            v_Err_Msg := '�ű�' || �ű�_In || '��' || To_Char(Trunc(����ʱ��_In), 'yyyy-mm-dd ') || '�Ѵﵽ�����������';
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
        (n_Ԥ��id, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), Nvl(���㷽ʽ_In, v_�ֽ�), Nvl(�ֽ�֧��_In, 0), �Ǽ�ʱ��_In, ����Ա���_In,
         ����Ա����_In, ����id_In, '�Һ��շ�', n_��id, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, ������λ_In, 4);
    
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
      
        If r_Deposit.Id <> 0 Then
          --��һ�γ�Ԥ��(82592,����һ�α��Ͻ���ID,��Ԥ�����Ϊ0)
          Update ����Ԥ����¼ Set ��Ԥ�� = 0, ����id = ����id_In, �������� = 4 Where ID = r_Deposit.Id;
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
        Where ����id = ����id_In And ���� = 1 And ���� = Nvl(r_Deposit.Ԥ�����, 2);
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
    If ���_In = 1 And Nvl(���½������_In,0)=0 Then
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
    If Nvl(���ʷ���_In, 0) = 0 Then
      --����Ʊ��ʹ�����
      If ���_In = 1 And Ʊ�ݺ�_In Is Not Null Then
        Select Ʊ�ݴ�ӡ����_Id.Nextval Into n_��ӡid From Dual;
      
        --����Ʊ��
        Insert Into Ʊ�ݴ�ӡ���� (ID, ��������, NO) Values (n_��ӡid, 4, ���ݺ�_In);
      
        Insert Into Ʊ��ʹ����ϸ
          (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����)
        Values
          (Ʊ��ʹ����ϸ_Id.Nextval, Decode(�շ�Ʊ��_In, 1, 1, 4), Ʊ�ݺ�_In, 1, 1, ����id_In, n_��ӡid, �Ǽ�ʱ��_In, ����Ա����_In);
      
        --״̬�Ķ�
        Update Ʊ�����ü�¼
        Set ��ǰ���� = Ʊ�ݺ�_In, ʣ������ = Decode(Sign(ʣ������ - 1), -1, 0, ʣ������ - 1), ʹ��ʱ�� = Sysdate
        Where ID = Nvl(����id_In, 0);
      End If;
    End If;
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
       ����Ա����, ����, ����, ����, ԤԼ, ԤԼ��ʽ, ժҪ, ������ˮ��, ����˵��, ������λ, ����ʱ��, ������, ԤԼ����Ա, ԤԼ����Ա���, ����, ҽ�Ƹ��ʽ)
    Values
      (n_�Һ�id, ���ݺ�_In, Decode(Nvl(ԤԼ�Һ�_In, 0), 1, 2, 1), 1, ����id_In, �����_In, ����_In, �Ա�_In, ����_In, �ű�_In, ����_In, ����_In,
       Null, ִ�в���id_In, ҽ������_In, 0, Null, �Ǽ�ʱ��_In, ����ʱ��_In, Decode(Nvl(ԤԼ�Һ�_In, 0), 1, ����ʱ��_In, Null), ����Ա���_In,
       ����Ա����_In, ����_In, n_���, ����_In, Decode(ԤԼ����_In, 1, 1, 0), ԤԼ��ʽ_In, ժҪ_In, ������ˮ��_In, ����˵��_In, ������λ_In,
       Decode(Nvl(ԤԼ�Һ�_In, 0), 0, �Ǽ�ʱ��_In, Null), Decode(Nvl(ԤԼ�Һ�_In, 0), 0, ����Ա����_In, Null),
       Decode(Nvl(ԤԼ�Һ�_In, 0), 1, ����Ա����_In, Null), Decode(Nvl(ԤԼ�Һ�_In, 0), 1, ����Ա���_In, Null),
       Decode(Nvl(����_In, 0), 0, Null, ����_In), v_���ʽ);
    n_ԤԼ���ɶ��� := 0;
    If Nvl(ԤԼ�Һ�_In, 0) = 1 Then
      n_ԤԼ���ɶ��� := Zl_To_Number(zl_GetSysParameter('ԤԼ���ɶ���', 1113));
    End If;
  
    --0-����������;1-��ҽ�������̨�Ŷ�;2-�ȷ���,��ҽ��վ
    If Nvl(���ɶ���_In, 0) <> 0 And Nvl(ԤԼ�Һ�_In, 0) = 0 Or n_ԤԼ���ɶ��� = 1 Then
      n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113));
      If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Or n_ԤԼ���ɶ��� = 1 Then
      
        --��������
        --.����ִ�в��š� �ķ�ʽ���ɶ���
        v_�������� := ִ�в���id_In;
        v_�ŶӺ��� := Zlgetnextqueue(ִ�в���id_In, n_�Һ�id, �ű�_In || '|' || n_���);
        v_�Ŷ���� := Zlgetsequencenum(0, n_�Һ�id, 0);
        d_�Ŷ�ʱ�� := Zl_Get_Queuedate(n_�Һ�id, �ű�_In, n_���, ����ʱ��_In);
        --  ��������_In , ҵ������_In, ҵ��id_In,����id_In,�ŶӺ���_In,�Ŷӱ��_In,��������_In,����ID_IN, ����_In, ҽ������_In, �Ŷ�ʱ��_In
        Zl_�ŶӽкŶ���_Insert(v_��������, 0, n_�Һ�id, ִ�в���id_In, v_�ŶӺ���, Null, ����_In, ����id_In, ����_In, ҽ������_In, d_�Ŷ�ʱ��, ԤԼ��ʽ_In,
                         Null, v_�Ŷ����);
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
    Where ����id = ����id_In And Exists (Select 1
           From ���˵�����¼
           Where ����id = ����id_In And ��ҳid Is Not Null And
                 �Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From ���˵�����¼ Where ����id = ����id_In));
  
    If Sql%RowCount > 0 Then
      Update ���˵�����¼
      Set ����ʱ�� = Sysdate
      Where ����id = ����id_In And ��ҳid Is Not Null And Nvl(����ʱ��, Sysdate) > Sysdate;
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
  ���½������_In  Number := 0--�Ƿ������Ա��������Ҫ�Ǵ���ͳһ����Ա��¼��̨�������������
) As
  ----------------------------------------------
  --��������_In:0-������Ԥ��;1-��Ϊ���۵�
  v_Err_Msg Varchar2(200);
  Err_Item Exception;

  v_����   ���㷽ʽ.����%Type;
  v_��ӡid Ʊ�ݴ�ӡ����.Id%Type;
  v_����   ������Ϣ.��������%Type;
  v_Date   Date;
  n_����ֵ �������.Ԥ�����%Type;
  n_��id   ����ɿ����.Id%Type;
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
  If ��������_In = 0 Then
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
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Ԥ����¼_Insert;
/

CREATE OR REPLACE Procedure zl_��Ա�ɿ����_Update(
����ģ��_In      Number,
����id_In        ����Ԥ����¼.����id%Type,
���㷽ʽ_In      ����Ԥ����¼.���㷽ʽ%Type,
�ֽ�֧��_In      ����Ԥ����¼.��Ԥ��%Type,
����֧��_In      ����Ԥ����¼.��Ԥ��%Type,
����Ա����_In    ����Ԥ����¼.����Ա����%Type
) as 
  Err_Item Exception;
  v_Err_Msg Varchar2(255);
  n_����ֵ �������.Ԥ�����%Type;
  v_�ֽ�     ���㷽ʽ.����%Type;
  v_�����ʻ� ���㷽ʽ.����%Type;
begin
  ---����ģ�������Ա�������
  ---����ģ�飺1-Ԥ����,2-���ʲ���,3-�շ��տ�,4-�Һ��տ�,5-���￨�տ�
 if ����ģ��_In=1 or ����ģ��_In=5 then
    --��Ա�ɿ����(����)
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) + �ֽ�֧��_In
    Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = ���㷽ʽ_In
    Returning ��� Into n_����ֵ;

    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ���� (�տ�Ա, ���㷽ʽ, ����, ���) Values (����Ա����_In, ���㷽ʽ_In, 1, �ֽ�֧��_In);
      n_����ֵ := �ֽ�֧��_In;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ��Ա�ɿ���� Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = ���㷽ʽ_In And Nvl(���, 0) = 0;
    End If;
 elsif ����ģ��_In=3 then
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
 elsif ����ģ��_In=4 then
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
 End if;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End zl_��Ա�ɿ����_Update;
/

--92776:������,2016-01-15,ɾ����Ⱦ�����濨ʱ��������濨�����ķ��������ļ�ID
CREATE OR REPLACE Procedure Zl_���Ӳ�����¼_Delete(Id_In In ���Ӳ�����¼.Id%Type) Is
  n_����״̬ ���Ӳ�����¼.����״̬%Type; 
  e_Submit Exception; 
Begin 
  Select Nvl(����״̬, 0) Into n_����״̬ From ���Ӳ�����¼ Where ID = Id_In; 
  If n_����״̬ > 0 Then 
    Raise e_Submit; 
  End If; 
  Delete ������ϼ�¼ T 
  Where t.Id In (Select a.Id 
                 From ������ϼ�¼ A, ���Ӳ�����¼ C 
                 Where a.����id = c.Id And a.����id = c.����id And a.��ҳid = c.��ҳid And c.Id = Id_In); 
  Update ���Ӳ���ʱ�� 
  Set ��ɼ�¼id = Null, ���ʱ�� = Null 
  Where (����id, ��ҳid, �ļ�id) = (Select ����id, ��ҳid, �ļ�id From ���Ӳ�����¼ Where ID = Id_In) And ��ɼ�¼id = Id_In; 
  update �������Լ�¼ set �ļ�ID = NULL where �ļ�ID = Id_In;  --��Ⱦ������ϵͳ��������ķ��������ļ�ID
  Delete ���Ӳ�����ӡ Where �ļ�id = Id_In; 
  Delete ���Ӳ�����¼ Where ID = Id_In; 
  Delete �����걨��¼ Where �ļ�id = Id_In; --Ϊ֧���°没����ɾ������� 
Exception 
  When e_Submit Then 
    Raise_Application_Error(-20101, '[ZLSOFT]����ɾ�����������յĲ�����[ZLSOFT]'); 
  When Others Then 
    zl_ErrorCenter(SQLCode, SQLErrM); 
End Zl_���Ӳ�����¼_Delete;
/

--92527:������,2016-01-15,������ò����������
Create Or Replace Procedure Zl_����δ���������_Recalc(����id_In סԺ���ü�¼.����id%Type) As
  v_�ѱ�     �ѱ�.����%Type;
  v_No       ������ü�¼.NO%Type;
  n_ʵ�ս�� ������ü�¼.ʵ�ս��%Type;
  n_������� �������.�������%Type;
  n_С��λ�� Number(2);
  v_Counter  Number(5);
  d_Sysdate  Date;
  v_Thisinfo Varchar(100);
  v_Lastinfo Varchar(100);

  Err_Custom Exception;
  v_Error Varchar2(255);
Begin
  Select �ѱ� Into v_�ѱ� From ������Ϣ Where ����id = ����id_In;

  --�����ж�
  --a.��ǰ���ǰ���������ܼ����ۿ�ģʽ
  v_Counter := To_Number(Nvl(Zl_Getsysparameter(93), 0));
  If v_Counter = 1 Then
    v_Error := '��ǰ�ѱ�ʹ����������ܼ����ۿ�ģʽ,��֧�ַ�������!';
    Raise Err_Custom;
  End If;

  --b.��ǰ�ѱ���ʹ��ҩƷ���ɱ��ۼ��մ��۵ķѱ�
  v_Counter := 0;
  Select Count(�ѱ�) Into v_Counter From �ѱ���ϸ Where �ѱ� = v_�ѱ� And ���㷽�� = 1;
  If v_Counter > 0 Then
    v_Error := '��ǰ�ѱ�ʹ��ҩƷ���ɱ��ۼ��մ���ģʽ,��֧�ַ�������!';
    Raise Err_Custom;
  End If;

  --c.û��δ�����
  Begin
    Select ������� Into n_������� From ������� Where ����id = ����id_In and ����=1 and ����=1;
  Exception
    When Others Then
      n_������� := 0;
  End;
  --������δ����ã������Ǳ���סԺ�����ģ��ں���ִ��ʱ���жϱ����Ƿ���δ����ϸ
  If n_������� = 0 Then
    v_Error := '���˲�����δ�����,���ý��з�������!';
    Raise Err_Custom;
  End If;

  --d.�������뱾��סԺ�ѱ�ͬ�ķ�����ϸ
  v_Counter := 0;
  Select Count(ID) Into v_Counter From ������ü�¼ Where ����id = ����id_In And �ѱ� <> v_�ѱ�;
  If v_Counter = 0 Then
    v_Error := '���˲������뱾��סԺ�ѱ�ͬ�ķ�����ϸ ,���ý��з�������!';
    Raise Err_Custom;
  End If;

  --ִ��
  v_Counter  := 0;
  d_Sysdate  := Sysdate;
  n_С��λ�� := To_Number(Nvl(Zl_Getsysparameter(9), 2));
  For r_Fee In (Select ����id, �����־, ���ʷ���, ����, �Ա�, ����, ��ʶ��, ���˿���id, �շ����, �շ�ϸĿid, ���㵥λ,
                       �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��������id, ������, ִ�в���id, ����ʱ��,
                       ����Ա���, ����Ա����, Nvl(Sum(Ӧ�ս��), 0) Ӧ�ս��, Nvl(Sum(ʵ�ս��), 0) ʵ�ս��
                From (Select ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, ���ʵ�id, ����id, ҽ�����,
                              �����־, ���ʷ���, ����, �Ա�, ����, ��ʶ��, 0 As ���˲���id, ���˿���id, �ѱ�, �շ����,
                              �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ,
                              ��׼����, Ӧ�ս��, ʵ�ս��, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id,
                              ִ����, ִ��״̬, ִ��ʱ��, ����, ����Ա���, ����Ա����, ����id, ���ʽ��, ���մ���id,
                              ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���
                       From ������ü�¼
                       Union All
                       Select ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, ���ʵ�id, ����id, ҽ�����,
                              �����־, ���ʷ���, ����, �Ա�, ����, ��ʶ��, 0 As ���˲���id, ���˿���id, �ѱ�, �շ����,
                              �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ,
                              ��׼����, Ӧ�ս��, ʵ�ս��, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id,
                              ִ����, ִ��״̬, ִ��ʱ��, ����, ����Ա���, ����Ա����, ����id, ���ʽ��, ���մ���id,
                              ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���
                       From H������ü�¼)
                Where ����id = ����id_In And ��¼״̬ <> 0 And ���ʷ��� = 1
                Group By ����id, �����־, ���ʷ���, ����, �Ա�, ����, ��ʶ��, ���˿���id, �շ����, �շ�ϸĿid, ���㵥λ,
                         �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��������id, ������, ִ�в���id, ����ʱ��,
                         ����Ա���, ����Ա����
                Having(Nvl(Sum(ʵ�ս��), 0) <> Nvl(Sum(���ʽ��), 0) Or Nvl(Sum(���ʽ��), 0) = 0) And Not(Nvl(Sum(Ӧ�ս��), 0) = 0 And Nvl(Sum(ʵ�ս��), 0) = 0)
                Order By ��������id, ������, ����Ա����) Loop
    --          ������δ��ķ���,������ϸ���ֽ���,�Լ����ʺ�����,��Щ��¼�п�����ת��󱸱�
    --          1.�ſ�����ȫ�����ʵļ�¼(Sum(Ӧ�ս��)=Sum(Ӧ�ս��))
    --          2.�ſ����޴��۳���ļ��ʺ������ʵļ�¼(Sum(Ӧ�ս��)=0,Sum(Ӧ�ս��)=0)
    --          3.���ſ����۳�������˵������ʵļ�¼��Ҫ��ԭ�����¼һ����������(Sum(Ӧ�ս��)=0,Sum(Ӧ�ս��)<>0)
    --          4.���ſ����۳���������ʵ�պͽ��ʶ�Ϊ��ļ�¼����Ϊ�Ļ�ԭ���ķѱ�ʱ��Ҫ�����ȥ
    If r_Fee.Ӧ�ս�� <> 0 Then
      Begin
        Select ʵ�ս��
        Into n_ʵ�ս��
        From (Select Round(r_Fee.Ӧ�ս�� * Nvl(ʵ�ձ���, 0) / 100, n_С��λ��) ʵ�ս��
               From �ѱ���ϸ
               Where �շ�ϸĿid = r_Fee.�շ�ϸĿid And �ѱ� = v_�ѱ� And Abs(r_Fee.Ӧ�ս��) Between Ӧ�ն���ֵ And
                     Ӧ�ն�βֵ And Nvl(���㷽��, 0) = 0
               Union All
               Select Round(r_Fee.Ӧ�ս�� * Nvl(ʵ�ձ���, 0) / 100, n_С��λ��) ʵ�ս��
               From �ѱ���ϸ A
               Where ������Ŀid = r_Fee.������Ŀid And �ѱ� = v_�ѱ� And Abs(r_Fee.Ӧ�ս��) Between Ӧ�ն���ֵ And
                     Ӧ�ն�βֵ And Nvl(���㷽��, 0) = 0 And Not Exists
                (Select 1 From �ѱ���ϸ B Where B.�ѱ� = A.�ѱ� And B.�շ�ϸĿid = r_Fee.�շ�ϸĿid));
      Exception
        When Others Then
          n_ʵ�ս�� := r_Fee.Ӧ�ս��;
      End;
    Else
      n_ʵ�ս�� := 0;
    End If;
    --�����������ԭʵ�յĲ��
    n_ʵ�ս�� := -1 * (r_Fee.ʵ�ս�� - n_ʵ�ս��);
  
    If n_ʵ�ս�� <> 0 Then
      --һ�ŵ��ݵĿ�������id,������,����Ա����,����Ҫ����ͬ���������֮һ����������µ��ݣ������û�б䣬һ�ŵ������100����ϸ
      v_Thisinfo := r_Fee.��������id || r_Fee.������ || r_Fee.����Ա���� || ' ';
      If v_Counter = 0 Or v_Counter = 100 Or v_Thisinfo <> v_Lastinfo Then
        v_No       := Nextno(14);
        v_Counter  := 1;
        v_Lastinfo := v_Thisinfo;
      Else
        v_Counter := v_Counter + 1;
      End If;
    
      Insert Into ������ü�¼
        (ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, ���ʵ�id, �����־, ����id, ��ʶ��, ����, �Ա�, ����,
         ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ����, ����, ��ҩ����, �Ӱ��־,
         ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ͳ����, ���ʷ���, ������, ��������id,
         ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ��״̬, ����id, ���ʽ��, ����Ա���, ����Ա����, ժҪ, �Ƿ���,
         ҽ�����)
      Values
        (���˷��ü�¼_Id.Nextval, 2, v_No, 1, v_Counter, Null, Null, Null, r_Fee.�����־, r_Fee.����id, r_Fee.��ʶ��,
         r_Fee.����, r_Fee.�Ա�, r_Fee.����, r_Fee.���˿���id, v_�ѱ�, r_Fee.�շ����, r_Fee.�շ�ϸĿid, r_Fee.���㵥λ,
         Null, Null, 0, 0, Null, r_Fee.�Ӱ��־, r_Fee.���ӱ�־, r_Fee.Ӥ����, r_Fee.������Ŀid, r_Fee.�վݷ�Ŀ, 0, 0,
         n_ʵ�ս��, Null, 1, Null, r_Fee.��������id, r_Fee.������, r_Fee.����ʱ��, d_Sysdate, r_Fee.ִ�в���id, 0, Null,
         Null, r_Fee.����Ա���, r_Fee.����Ա����, Decode(v_Counter, 1, 'ʵ��������', ''), 0, Null);
    End If;
  End Loop;

  If v_Counter = 0 Then
    v_Error := '��������ԭ��֮һ,û�н��з�������:' || Chr(13) || Chr(13) || 'a.û�з��ֲ��˱���סԺ��δ�����.' ||
               Chr(13) || 'b.����δ������ѽ����˷�������.' || Chr(13) || 'c.����ǰ�ѱ������ʵ�ճ����Ϊ��.';
    Raise Err_Custom;
  Else
    --�������
    n_ʵ�ս�� := 0;
    Select Sum(ʵ�ս��)
    Into n_ʵ�ս��
    From ������ü�¼
    Where ����id = ����id_In And ��¼���� = 2 And �Ǽ�ʱ�� = d_Sysdate;
    Update ������� Set ������� = Nvl(�������, 0) + n_ʵ�ս�� Where ����id = ����id_In And ���� = 1 And ���� = 1;
    If Sql%Rowcount = 0 Then
      Insert Into ������� (����id, ����, �������, Ԥ�����, ����) Values (����id_In, 1, n_ʵ�ս��, 0, 1);
    End If;
  
    --����δ�����
    For r_Fee In (Select Null As ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, Sum(ʵ�ս��) ʵ�ս��
                  From ������ü�¼
                  Where ����id = ����id_In And ��¼���� = 2 And �Ǽ�ʱ�� = d_Sysdate
                  Group By ���˿���id, ��������id, ִ�в���id, ������Ŀid) Loop
      Update ����δ�����
      Set ��� = Nvl(���, 0) + r_Fee.ʵ�ս��
      Where ����id = ����id_In And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = Nvl(r_Fee.���˲���id, 0) And
            Nvl(���˿���id, 0) = Nvl(r_Fee.���˿���id, 0) And Nvl(��������id, 0) = Nvl(r_Fee.��������id, 0) And
            Nvl(ִ�в���id, 0) = Nvl(r_Fee.ִ�в���id, 0) And ������Ŀid + 0 = Nvl(r_Fee.������Ŀid, 0) And
            ��Դ;�� + 0 = 2;
      If Sql%Rowcount = 0 Then
        Insert Into ����δ�����
          (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
        Values
          (����id_In, Null, r_Fee.���˲���id, r_Fee.���˿���id, r_Fee.��������id, r_Fee.ִ�в���id, r_Fee.������Ŀid, 2,
           r_Fee.ʵ�ս��);
      End If;
    End Loop;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����δ���������_Recalc;
/

--92527:������,2016-01-14,�ѱ���������������
Create Or Replace Procedure Zl_����δ�����_Recalc
(
  ����id_In סԺ���ü�¼.����id%Type,
  ��ҳid_In סԺ���ü�¼.��ҳid%Type
) As
  v_�ѱ�     �ѱ�.����%Type;
  v_No       סԺ���ü�¼.No%Type;
  n_ʵ�ս�� סԺ���ü�¼.ʵ�ս��%Type;
  n_������� �������.�������%Type;
  n_С��λ�� Number(2);
  v_Counter  Number(5);
  d_Sysdate  Date;
  v_Thisinfo Varchar(100);
  v_Lastinfo Varchar(100);

  Err_Custom Exception;
  v_Error Varchar2(255);
Begin
  Select �ѱ� Into v_�ѱ� From ������ҳ Where ����id = ����id_In And ��ҳid = ��ҳid_In;

  --�����ж�
  --a.��ǰ���ǰ���������ܼ����ۿ�ģʽ
  v_Counter := To_Number(Nvl(zl_GetSysParameter(93), 0));
  If v_Counter = 1 Then
    v_Error := '��ǰ�ѱ�ʹ����������ܼ����ۿ�ģʽ,��֧�ַ�������!';
    Raise Err_Custom;
  End If;

  --b.��ǰ�ѱ���ʹ��ҩƷ���ɱ��ۼ��մ��۵ķѱ�
  v_Counter := 0;
  Select Count(�ѱ�) Into v_Counter From �ѱ���ϸ Where �ѱ� = v_�ѱ� And ���㷽�� = 1;
  If v_Counter > 0 Then
    v_Error := '��ǰ�ѱ�ʹ��ҩƷ���ɱ��ۼ��մ���ģʽ,��֧�ַ�������!';
    Raise Err_Custom;
  End If;

  --c.û��δ�����
  Begin
    Select ������� Into n_������� From ������� Where ����id = ����id_In And ���� = 2 And ���� = 1;
  Exception
    When Others Then
      n_������� := 0;
  End;
  --������δ����ã������Ǳ���סԺ�����ģ��ں���ִ��ʱ���жϱ����Ƿ���δ����ϸ
  If n_������� = 0 Then
    v_Error := '���˲�����δ�����,���ý��з�������!';
    Raise Err_Custom;
  End If;

  --d.�������뱾��סԺ�ѱ�ͬ�ķ�����ϸ
  v_Counter := 0;
  Select Count(ID) Into v_Counter From סԺ���ü�¼ Where ����id = ����id_In And ��ҳid = ��ҳid_In And �ѱ� <> v_�ѱ�;
  If v_Counter = 0 Then
    v_Error := '���˲������뱾��סԺ�ѱ�ͬ�ķ�����ϸ ,���ý��з�������!';
    Raise Err_Custom;
  End If;

  --ִ��
  v_Counter  := 0;
  d_Sysdate  := Sysdate;
  n_С��λ�� := To_Number(Nvl(zl_GetSysParameter(9), 2));
  For r_Fee In (Select ����id, ��ҳid, �����־, ���ʷ���, ����, �Ա�, ����, ��ʶ��, ����, ���˲���id, ���˿���id, �շ����, �շ�ϸĿid, ���㵥λ, �Ӱ��־, ���ӱ�־, Ӥ����,
                       ������Ŀid, �վݷ�Ŀ, ��������id, ������, ִ�в���id, ����ʱ��, ����Ա���, ����Ա����, ҽ��С��id, Nvl(Sum(Ӧ�ս��), 0) Ӧ�ս��,
                       Nvl(Sum(ʵ�ս��), 0) ʵ�ս��
                From (Select ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, ���ʵ�id, ����id, ��ҳid, ҽ�����, �����־, ���ʷ���, ����, �Ա�,
                              ����, ��ʶ��, ����, ���˲���id, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid,
                              �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ����, ����Ա���,
                              ����Ա����, ����id, ���ʽ��, ���մ���id, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���, ҽ��С��id
                       From סԺ���ü�¼
                       Union All
                       Select ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, ���ʵ�id, ����id, ��ҳid, ҽ�����, �����־, ���ʷ���, ����, �Ա�,
                              ����, ��ʶ��, ����, ���˲���id, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid,
                              �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ����, ����Ա���,
                              ����Ա����, ����id, ���ʽ��, ���մ���id, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���, ҽ��С��id
                       From HסԺ���ü�¼)
                Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼״̬ <> 0 And ���ʷ��� = 1
                Group By ����id, ��ҳid, �����־, ���ʷ���, ����, �Ա�, ����, ��ʶ��, ����, ���˲���id, ���˿���id, �շ����, �շ�ϸĿid, ���㵥λ, �Ӱ��־, ���ӱ�־, Ӥ����,
                         ������Ŀid, �վݷ�Ŀ, ��������id, ������, ִ�в���id, ����ʱ��, ����Ա���, ����Ա����, ҽ��С��id
                Having(Nvl(Sum(ʵ�ս��), 0) <> Nvl(Sum(���ʽ��), 0) Or Nvl(Sum(���ʽ��), 0) = 0) And Not(Nvl(Sum(Ӧ�ս��), 0) = 0 And Nvl(Sum(ʵ�ս��), 0) = 0)
                Order By ��������id, ������, ����Ա����) Loop
    --          ������δ��ķ���,������ϸ���ֽ���,�Լ����ʺ�����,��Щ��¼�п�����ת��󱸱�
    --          1.�ſ�����ȫ�����ʵļ�¼(Sum(Ӧ�ս��)=Sum(Ӧ�ս��))
    --          2.�ſ����޴��۳���ļ��ʺ������ʵļ�¼(Sum(Ӧ�ս��)=0,Sum(Ӧ�ս��)=0)
    --          3.���ſ����۳�������˵������ʵļ�¼��Ҫ��ԭ�����¼һ����������(Sum(Ӧ�ս��)=0,Sum(Ӧ�ս��)<>0)
    --          4.���ſ����۳���������ʵ�պͽ��ʶ�Ϊ��ļ�¼����Ϊ�Ļ�ԭ���ķѱ�ʱ��Ҫ�����ȥ
    If r_Fee.Ӧ�ս�� <> 0 Then
      Begin
        Select ʵ�ս��
        Into n_ʵ�ս��
        From (Select Round(r_Fee.Ӧ�ս�� * Nvl(ʵ�ձ���, 0) / 100, n_С��λ��) ʵ�ս��
               From �ѱ���ϸ
               Where �շ�ϸĿid = r_Fee.�շ�ϸĿid And �ѱ� = v_�ѱ� And Abs(r_Fee.Ӧ�ս��) Between Ӧ�ն���ֵ And Ӧ�ն�βֵ And Nvl(���㷽��, 0) = 0
               Union All
               Select Round(r_Fee.Ӧ�ս�� * Nvl(ʵ�ձ���, 0) / 100, n_С��λ��) ʵ�ս��
               From �ѱ���ϸ A
               Where ������Ŀid = r_Fee.������Ŀid And �ѱ� = v_�ѱ� And Abs(r_Fee.Ӧ�ս��) Between Ӧ�ն���ֵ And Ӧ�ն�βֵ And Nvl(���㷽��, 0) = 0 And
                     Not Exists (Select 1 From �ѱ���ϸ B Where b.�ѱ� = a.�ѱ� And b.�շ�ϸĿid = r_Fee.�շ�ϸĿid));
      Exception
        When Others Then
          n_ʵ�ս�� := r_Fee.Ӧ�ս��;
      End;
    Else
      n_ʵ�ս�� := 0;
    End If;
    --�����������ԭʵ�յĲ��
    n_ʵ�ս�� := -1 * (r_Fee.ʵ�ս�� - n_ʵ�ս��);
  
    If n_ʵ�ս�� <> 0 Then
      --һ�ŵ��ݵĿ�������id,������,����Ա����,����Ҫ����ͬ���������֮һ����������µ��ݣ������û�б䣬һ�ŵ������100����ϸ
      v_Thisinfo := r_Fee.��������id || r_Fee.������ || r_Fee.����Ա���� || r_Fee.����;
      If v_Counter = 0 Or v_Counter = 100 Or v_Thisinfo <> v_Lastinfo Then
        v_No       := Nextno(14);
        v_Counter  := 1;
        v_Lastinfo := v_Thisinfo;
      Else
        v_Counter := v_Counter + 1;
      End If;
    
      Insert Into סԺ���ü�¼
        (ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, ���ʵ�id, �����־, ����id, ��ҳid, ��ʶ��, ����, ����, �Ա�, ����, ���˲���id, ���˿���id, �ѱ�,
         �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ����, ����, ��ҩ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ͳ����, ���ʷ���,
         ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ��״̬, ����id, ���ʽ��, ����Ա���, ����Ա����, ժҪ, �Ƿ���, ҽ�����, ҽ��С��id)
      Values
        (���˷��ü�¼_Id.Nextval, 2, v_No, 1, v_Counter, Null, Null, 0, Null, r_Fee.�����־, r_Fee.����id, r_Fee.��ҳid, r_Fee.��ʶ��,
         r_Fee.����, r_Fee.����, r_Fee.�Ա�, r_Fee.����, r_Fee.���˲���id, r_Fee.���˿���id, v_�ѱ�, r_Fee.�շ����, r_Fee.�շ�ϸĿid, r_Fee.���㵥λ,
         Null, Null, 0, 0, Null, r_Fee.�Ӱ��־, r_Fee.���ӱ�־, r_Fee.Ӥ����, r_Fee.������Ŀid, r_Fee.�վݷ�Ŀ, 0, 0, n_ʵ�ս��, Null, 1,
         Null, r_Fee.��������id, r_Fee.������, r_Fee.����ʱ��, d_Sysdate, r_Fee.ִ�в���id, 0, Null, Null, r_Fee.����Ա���, r_Fee.����Ա����,
         Decode(v_Counter, 1, 'ʵ��������', ''), 0, Null, r_Fee.ҽ��С��id);
    End If;
  End Loop;

  If v_Counter = 0 Then
    v_Error := '��������ԭ��֮һ,û�н��з�������:' || Chr(13) || Chr(13) || 'a.û�з��ֲ��˱���סԺ��δ�����.' || Chr(13) || 'b.����δ������ѽ����˷�������.' ||
               Chr(13) || 'c.����ǰ�ѱ������ʵ�ճ����Ϊ��.';
    Raise Err_Custom;
  Else
    --�������
    n_ʵ�ս�� := 0;
    Select Sum(ʵ�ս��)
    Into n_ʵ�ս��
    From סԺ���ü�¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼���� = 2 And �Ǽ�ʱ�� = d_Sysdate;
    Update ������� Set ������� = Nvl(�������, 0) + n_ʵ�ս�� Where ����id = ����id_In And ���� = 1 And ���� = 2;
    If Sql%RowCount = 0 Then
      Insert Into ������� (����id, ����, �������, Ԥ�����, ����) Values (����id_In, 1, n_ʵ�ս��, 0, 2);
    End If;
  
    --����δ�����
    For r_Fee In (Select ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, Sum(ʵ�ս��) ʵ�ս��
                  From סԺ���ü�¼
                  Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼���� = 2 And �Ǽ�ʱ�� = d_Sysdate
                  Group By ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid) Loop
      Update ����δ�����
      Set ��� = Nvl(���, 0) + r_Fee.ʵ�ս��
      Where ����id = ����id_In And Nvl(��ҳid, 0) = ��ҳid_In And Nvl(���˲���id, 0) = r_Fee.���˲���id And
            Nvl(���˿���id, 0) = r_Fee.���˿���id And Nvl(��������id, 0) = r_Fee.��������id And Nvl(ִ�в���id, 0) = r_Fee.ִ�в���id And
            ������Ŀid + 0 = r_Fee.������Ŀid And ��Դ;�� + 0 = 2;
      If Sql%RowCount = 0 Then
        Insert Into ����δ�����
          (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
        Values
          (����id_In, ��ҳid_In, r_Fee.���˲���id, r_Fee.���˿���id, r_Fee.��������id, r_Fee.ִ�в���id, r_Fee.������Ŀid, 2, r_Fee.ʵ�ս��);
      End If;
    End Loop;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����δ�����_Recalc;
/

--89717:��ΰ��,2016-01-14,��Ժ������ȡ�����·��
CREATE OR REPLACE Procedure Zl_����·������_Delete
(
  ·����¼id_In �����ٴ�·��.Id%Type,
  ��������_In   �����ٴ�·��.״̬%Type
) Is
  v_�׶�id     ����·������.�׶�id%Type;
  v_ǰһ�׶�id ����·������.�׶�id%Type;
  v_����       ����·������.����%Type;
  v_����       ����·������.����%Type;

  v_����id       �����ٴ�·��.����id%Type;
  v_��ҳid       �����ٴ�·��.��ҳid%Type;
  d_�Ǽ�ʱ��     ����·������.�Ǽ�ʱ��%Type;
  d_��Ժ����     ������ҳ.��Ժ����%Type;
  n_��ǰ�׶�id   ���˺ϲ�·��.��ǰ�׶�id%Type;
  v_�Ƿ����Ժ Varchar2(20);
  v_Error        Varchar2(255);
  Err_Custom Exception;
Begin
  --��Ժ���˲�����ȡ�����·��
  Select zl_GetSysParameter('��Ժ������ȡ�����·��', 1256) Into v_�Ƿ����Ժ From Dual;
  If v_�Ƿ����Ժ = '1' Then
    Select b.��Ժ����
    Into d_��Ժ����
    From �����ٴ�·�� A, ������ҳ B
    Where a.����id = b.����id And a.��ҳid = b.��ҳid And a.Id = ·����¼id_In;
    If d_��Ժ���� Is Not Null Then
      If d_��Ժ���� <= Sysdate Then
        v_Error := '�ò����Ѿ���Ժ,������ȡ�����·����';
        Raise Err_Custom;
      End If;
    End If;
  End If;
  --
  Select ǰһ�׶�id Into v_�׶�id From �����ٴ�·�� Where ID = ·����¼id_In;
  Select Max(����), Max(����)
  Into v_����, v_����
  From ����·��ִ��
  Where ·����¼id = ·����¼id_In And �׶�id = v_�׶�id;

  Select ����ʱ�� Into d_�Ǽ�ʱ�� From �����ٴ�·�� Where ID = ·����¼id_In;

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

  If ��������_In = 3 Then
    --�������Ϊ����ʱ�Զ�������,ȡ�������Զ�ȡ������
    Delete ����·������ Where ·����¼id = ·����¼id_In And �׶�id = v_�׶�id And ���� = v_����;
    Delete ����·��ָ�� Where ·����¼id = ·����¼id_In And �׶�id = v_�׶�id And ���� = v_����;
  End If;

  --b.���˵�ǰһ���׶�
  Select Max(�׶�id)
  Into v_ǰһ�׶�id
  From ����·��ִ��
  Where ·����¼id = ·����¼id_In And
        �Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From ����·��ִ�� Where ·����¼id = ·����¼id_In And �׶�id <> v_�׶�id);

  Update �����ٴ�·��
  Set ����ʱ�� = Null, ״̬ = 1, ǰһ�׶�id = v_ǰһ�׶�id, ��ǰ�׶�id = v_�׶�id, ��ǰ���� = v_����
  Where ID = ·����¼id_In
  Returning ����id, ��ҳid Into v_����id, v_��ҳid;

  --���²�����ҳ��ǰ·����״̬
  Update ������ҳ Set ·��״̬ = 1 Where ����id = v_����id And ��ҳid = v_��ҳid;

  Delete ���˳�����¼ Where ����id = v_����id And ��ҳid = v_��ҳid;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����·������_Delete;
/

--92518:������,2016-01-11,��Ⱦ�����Խ����Ϣ
Create Or Replace Procedure Zl_ҵ����Ϣ�嵥_Read
(
  ����id_In     In ҵ����Ϣ�嵥.����id%Type,
  ����id_In     In ҵ����Ϣ�嵥.����id%Type,
  ���ͱ���_In   In ҵ����Ϣ�嵥.���ͱ���%Type,
  �Ķ�����_In   In ҵ����Ϣ״̬.�Ķ�����%Type,
  �Ķ���_In     In ҵ����Ϣ״̬.�Ķ���%Type,
  �Ķ�����id_In In ҵ����Ϣ״̬.�Ķ�����id%Type,
  �Ķ�ʱ��_In   In ҵ����Ϣ״̬.�Ķ�ʱ��%Type := Null,
  ��Ϣid_In     In ҵ����Ϣ״̬.��Ϣid%Type := Null
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

--91752:������,2016-01-09,��¼ǩ��������������ǩ�����ݴ�����
Create Or Replace Procedure Zl_���ӻ����¼_Update
(
  ����id_In   In ���˻����¼.����id%Type,
  ��ҳid_In   In ���˻����¼.��ҳid%Type,
  Ӥ��_In     In ���˻����¼.Ӥ��%Type,
  ��ʼʱ��_In In ���˻����¼.����ʱ��%Type, --����¼��Ч��ȵĿ�ʼʱ�� 
  ����ʱ��_In In ���˻����¼.����ʱ��%Type, --����¼��Ч��ȵ���ֹʱ�䣬������¼Ϊÿ���ӣ����±�Ϊ4Сʱ,ʱ�����ڵ���ͬ��Ŀ��¼Ҫɾ�� 
  ��¼����_In In ���˻�������.��¼����%Type, --������Ŀ=1���ϱ�˵��=2�����ת���=3�������ձ��=4,ǩ����¼=5,�±�˵��=6 
  ��Ŀ���_In In ���˻�������.��Ŀ���%Type, --������Ŀ����ţ��ǻ�����Ŀ�̶�Ϊ0 
  ��¼���_In In ���˻�������.��¼���%Type, --��¼���ݵ������־ 
  ��¼����_In In ���˻�������.��¼����%Type := Null, --��¼���ݣ��������Ϊ�գ��������ǰ������ 
  ���²�λ_In In ���˻�������.���²�λ%Type := Null,
  ���˼�¼_In In Number := 1,
  ��Ŀ�״�_In In Number := 1,
  ���Ժϸ�_In In Number := 0,
  �Ƿ�˵��_In In Number := 0, --��˵��,����д��λ 
  ����ʱ��_In In ���˻����¼.����ʱ��%Type := Null, --�����¼�ķ���ʱ�� 
  δ��˵��_In In ���˻�������.δ��˵��%Type := Null,	--δ��˵��
  ����Ա_IN	  IN ���˻����¼.������%Type:=null
) Is
  v_������   ���˻����¼.������%Type;
  v_��¼��   ���˻�������.��¼��%Type;
  v_��¼���� ���˻�������.��¼����%Type;
  n_������ ���˻����¼.������%Type;
  d_����ʱ�� ���˻����¼.����ʱ��%Type;
  d_����ʱ�� ���˻����¼.����ʱ��%Type;
  n_��¼id   ���˻�������.��¼id%Type;
  v_����id   ���˻����¼.����id%Type;
  v_���     ���˻�������.��¼���%Type;
  v_���Ŀ �����¼��Ŀ.��Ŀ����%Type;
  n_��Ŀ���� �����¼��Ŀ.��Ŀ����%Type;
  n_��Ŀ��ʾ �����¼��Ŀ.��Ŀ��ʾ%Type;
  n_��ʼ�汾 ���˻�������.��ʼ�汾%Type;
  n_��ǰ�汾 ���˻�������.��ʼ�汾%Type;
  v_Records  Number;
  n_Add      Number;
  --������ 

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  Begin
    Select p.���� Into v_������ From �ϻ���Ա�� O, ��Ա�� P Where o.��Աid = p.Id And �û��� = User;
  Exception
    When Others Then
      v_������ := User;
  End;
  if ����Ա_IN is not null then 
	v_������ := ����Ա_IN;
  end if ;

  d_����ʱ�� := ����ʱ��_In;

  If d_����ʱ�� Is Null Then
    d_����ʱ�� := ��ʼʱ��_In;
  End If;

  If ����ʱ��_In Is Null Then
    d_����ʱ�� := ��ʼʱ��_In;
  Else
    d_����ʱ�� := ����ʱ��_In;
  End If;

  n_��Ŀ���� := 1;
  Begin
    Select ��Ŀ����, ��Ŀ��ʾ, ��Ŀ����
    Into n_��Ŀ����, n_��Ŀ��ʾ, v_���Ŀ
    From �����¼��Ŀ
    Where ��Ŀ��� = ��Ŀ���_In;
  Exception
    When Others Then
      v_���Ŀ := 1;
  End;
  --��鲡���ڱ��μ�¼ʱ�����ڣ�������ͬ��¼��Ŀ��������ʱ�䲻��ͬ�Ļ����¼���������� 
  --------------------------------------------------------------------------------------------------------------------- 
  If (��Ŀ�״�_In = 1) Or (��¼����_In Is Null And δ��˵��_In Is Null) Then
    For r_List In (Select l.Id, Count(*) As ��¼��
                   From ���˻����¼ L, ���˻������� D
                   Where l.Id = d.��¼id And l.����id = ����id_In And l.��ҳid = ��ҳid_In And Nvl(l.Ӥ��, 0) = Nvl(Ӥ��_In, 0) And
                         l.������Դ = 2 And d.��ֹ�汾 Is Null And d.��Ŀ��� = ��Ŀ���_In And d.��¼���� <> 5 And
                         (��¼����_In Is Null And l.����ʱ�� >= ��ʼʱ��_In Or ��¼����_In Is Not Null And l.����ʱ�� >= ��ʼʱ��_In) And
                         l.����ʱ�� <= d_����ʱ��
                   Group By l.Id) Loop
      n_��ǰ�汾 := 0;
      n_��¼id   := r_List.Id;
      Begin
        Select Nvl(��ʼ�汾, 1)
        Into n_��ǰ�汾
        From ���˻�������
        Where ��¼id = n_��¼id And ��Ŀ��� = ��Ŀ���_In And Nvl(��¼���, 0) = ��¼���_In And ��¼���� = ��¼����_In And
              Decode(v_���Ŀ, 2, Nvl(���²�λ, '��'), Nvl(���²�λ_In, '��')) = Nvl(���²�λ_In, '��') And ��ֹ�汾 Is Null;
      Exception
        When Others Then
          n_��ǰ�汾 := 0;
      End;
    
      If ��¼����_In = 2 Or ��¼����_In = 6 Then
        Delete ���˻�������
        Where ��¼id = n_��¼id And ��¼���� = ��¼����_In And ��Ŀ��� = ��Ŀ���_In And ��ֹ�汾 Is Null;
      Else
        If ���²�λ_In Is Not Null Then
          Delete ���˻�������
          Where ��¼id = n_��¼id And ��¼���� = ��¼����_In And ��Ŀ��� = ��Ŀ���_In And Nvl(���²�λ, '��') = Nvl(���²�λ_In, '��') And
                Nvl(��¼���, 0) = Nvl(��¼���_In, 0) And ��ֹ�汾 Is Null;
        Else
          Delete ���˻�������
          Where ��¼id = n_��¼id And ��¼���� = ��¼����_In And ��Ŀ��� = ��Ŀ���_In And Nvl(��¼���, 0) = Nvl(��¼���_In, 0) And ��ֹ�汾 Is Null;
        End If;
      End If;
    
      --����汾 
      Update ���˻�������
      Set ��ֹ�汾 = Null
      Where ��ֹ�汾 = n_��ǰ�汾 And ��¼id = n_��¼id And ��Ŀ��� = ��Ŀ���_In And ��¼���� = ��¼����_In And Nvl(��¼���, 0) = ��¼���_In And
            Decode(v_���Ŀ, 2, Nvl(���²�λ, '��'), Nvl(���²�λ_In, '��')) = Nvl(���²�λ_In, '��');
    
      --����Ƿ񻹴����ϴ�ǩ�����޸ĵļ�¼,���������,��ǩ����¼����ֹ�汾��Ϊ�� 
      Begin
        Select 1
        Into v_Records
        From ���˻�������
        Where ��ֹ�汾 = n_��ǰ�汾 And ��¼id = n_��¼id And ��¼���� <> 5 And Rownum < 2;
      Exception
        When Others Then
          v_Records := 0;
      End;
    
      If v_Records = 0 Then
        Update ���˻������� Set ��ֹ�汾 = Null Where ��ֹ�汾 = n_��ǰ�汾 And ��¼���� = 5 And ��¼id = n_��¼id;
      End If;
    
      Update ���˻����¼
      Set ���汾 = ���汾 - 1
      Where ID = n_��¼id And ���汾 Not In (Select ��ֹ�汾 From ���˻������� Where ��¼���� <> 5 And ��¼id = n_��¼id);
    
      Delete From ���˻�������
      Where ��¼id = n_��¼id And ��¼���� = 5 And
            Nvl(��ʼ�汾, 1) Not In
            (Select Nvl(��ʼ�汾, 1) From ���˻������� A Where a.��¼���� <> 5 And a.��¼id = n_��¼id);
    
      Delete From ���˻����¼ A
      Where a.Id = n_��¼id And Not Exists (Select 1 From ���˻������� B Where b.��¼id = a.Id);
    End Loop;
  End If;

  If ��¼����_In Is Null And δ��˵��_In Is Null Then
    Return;
  End If;
  --------------------------------------------------------------------------------------------------------------------- 
  n_��¼id := 0;
  Begin
    Select ID
    Into n_��¼id
    From ���˻����¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And Nvl(Ӥ��, 0) = Nvl(Ӥ��_In, 0) And ������Դ = 2 And ����ʱ�� = d_����ʱ�� And
          Rownum < 2;
  Exception
    When Others Then
      n_��¼id := 0;
  End;

  n_������ := Zl_Patittendgrade(����id_In, ��ҳid_In, ��ʼʱ��_In);

  --------------------------------------------------------------------------------------------------------------------- 
  v_����id := 0;
  Begin
    Select ����id
    Into v_����id
    From ���˱䶯��¼
    Where ����id Is Not Null And ����id = ����id_In And ��ҳid = ��ҳid_In And
          (��ʼʱ�� Between ��ʼʱ��_In And d_����ʱ�� Or ��ʼʱ�� <= ��ʼʱ��_In) And (��ʼʱ��_In <= ��ֹʱ�� Or ��ֹʱ�� Is Null) And Rownum < 2;
  Exception
    When Others Then
      v_����id := 0;
  End;
  If v_����id = 0 Then
    v_Error := '��' || To_Char(��ʼʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '��' || To_Char(d_����ʱ��, 'yyyy-mm-dd hh24:mi:ss') ||
               '�����޶�Ӧ���ң����ܲ�����';
    Raise Err_Custom;
  End If;

  --ȷ�Ͽ�ʼ�汾�� 
  --------------------------------------------------------------------------------------------------------------------- 
  Select Nvl(Max(Nvl(a.��ʼ�汾, 1)), 0) + 1
  Into n_��ʼ�汾
  From ���˻������� A, ���˻����¼ B
  Where b.����id = ����id_In And b.��ҳid = ��ҳid_In And Nvl(b.Ӥ��, 0) = Nvl(Ӥ��_In, 0) And b.������Դ = 2 And b.����ʱ�� = d_����ʱ�� And
        a.��¼id = b.Id And a.��¼���� = 5;

  n_��ǰ�汾 := n_��ʼ�汾;

  --����ǲ��Ǳ��˵ļ�¼ 
  n_Add      := 1;
  v_��¼��   := '';
  v_��¼���� := '';
  Begin
    Select ��¼��, ��¼����, Nvl(��ʼ�汾, 1)
    Into v_��¼��, v_��¼����, n_��ǰ�汾
    From ���˻�������
    Where ��¼id = n_��¼id And ��¼���� = ��¼����_In And ��Ŀ��� = ��Ŀ���_In And Nvl(��¼���, 0) = Nvl(��¼���_In, 0) And
          Decode(v_���Ŀ, 2, Nvl(���²�λ, '��'), Nvl(���²�λ_In, '��')) = Nvl(���²�λ_In, '��') And ��ֹ�汾 Is Null;
  Exception
    When Others Then
      v_��¼�� := '';
      n_Add    := 1;
  End;
  --------------------------------------------------------------------------------------------------------------------- 
  If ���˼�¼_In = 0 Then
    If v_��¼�� Is Not Null And v_��¼�� <> v_������ Then
      v_Error := '��' || To_Char(��ʼʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '��' || To_Char(d_����ʱ��, 'yyyy-mm-dd hh24:mi:ss') ||
                 '���ڼ�¼�˲��ǵ�ǰ�ˣ�����Ȩ�޸ģ�';
      Raise Err_Custom;
    End If;
  End If;
  --��д���˻����¼������Ѿ������벡�ˡ����Һͷ���ʱ����ͬ�ļ�¼���޸ģ����������µļ�¼ 
  --------------------------------------------------------------------------------------------------------------------- 
  If n_��¼id = 0 Then
    Select ���˻����¼_Id.Nextval Into n_��¼id From Dual;
    n_Add := 1;
  Else
    If n_��Ŀ���� = 0 And n_��Ŀ��ʾ = 0 Then
      If n_Add = 1 And Zl_To_Number(v_��¼����) = Zl_To_Number(��¼����_In) Then
        n_Add := 0;
      End If;
    Else
      If n_Add = 1 And v_��¼���� = ��¼����_In Then
        n_Add := 0;
      End If;
    End If;
  End If;

  If n_Add = 0 And n_��ʼ�汾 > n_��ǰ�汾 And n_��ʼ�汾 > 1 Then
    n_��ʼ�汾 := n_��ʼ�汾 - 1;
  End If;

  Update ���˻����¼ Set ������ = v_������, ����ʱ�� = Sysdate, ���汾 = n_��ʼ�汾 Where ID = n_��¼id;

  If Sql%RowCount = 0 Then
    Insert Into ���˻����¼
      (ID, ������Դ, ����id, ��ҳid, Ӥ��, ����id, ������, ����ʱ��, ������, ����ʱ��, ���汾)
    Values
      (n_��¼id, 2, ����id_In, ��ҳid_In, Ӥ��_In, v_����id, n_������, d_����ʱ��, v_������, Sysdate, n_��ʼ�汾);
  End If;

  --����汾���� 
  --------------------------------------------------------------------------------------------------------------------- 
  Update ���˻�������
  Set ��ֹ�汾 = n_��ʼ�汾, ��ʼ�汾 = Nvl(��ʼ�汾, 1)
  Where ��¼id = n_��¼id And ��¼���� = 5 And ��ֹ�汾 Is Null And n_Add = 1;

  If ��¼����_In = 1 Then
    Update ���˻�������
    Set ��ֹ�汾 = n_��ʼ�汾, ��ʼ�汾 = Nvl(��ʼ�汾, 1)
    Where ��¼id = n_��¼id And ��¼���� = ��¼����_In And ��Ŀ��� = ��Ŀ���_In And ��ֹ�汾 Is Null And n_Add = 1 And
          Decode(v_���Ŀ, 2, Nvl(���²�λ, '��'), Nvl(���²�λ_In, '��')) = Nvl(���²�λ_In, '��') And Nvl(��ʼ�汾, 1) <> n_��ʼ�汾;
  Else
    Update ���˻�������
    Set ��ֹ�汾 = n_��ʼ�汾, ��ʼ�汾 = Nvl(��ʼ�汾, 1)
    Where ��¼id = n_��¼id And ��¼���� = ��¼����_In And n_Add = 1 And
          ��Ŀ���� = Decode(��¼����_In, 2, '�ϱ�˵��', 6, '�±�˵��', 3, '���ת', 4, ��¼����_In) And ��ֹ�汾 Is Null And
          Nvl(��ʼ�汾, 1) <> n_��ʼ�汾;
  End If;

  --ɾ���Ѿ��Ǽǵĸ�����Ĳ��˻������� 
  --------------------------------------------------------------------------------------------------------------------- 
  If ��¼����_In = 2 Or ��¼����_In = 6 Then
    Delete ���˻�������
    Where ��¼id = n_��¼id And ��¼���� = ��¼����_In And ��Ŀ��� = ��Ŀ���_In And ��ֹ�汾 Is Null;
  Else
    Delete ���˻�������
    Where ��¼id = n_��¼id And ��¼���� = ��¼����_In And ��Ŀ��� = ��Ŀ���_In And
          Decode(v_���Ŀ, 2, Nvl(���²�λ, '��'), Nvl(���²�λ_In, '��')) = Nvl(���²�λ_In, '��') And Nvl(��¼���, 0) = Nvl(��¼���_In, 0) And
          ��ֹ�汾 Is Null;
  End If;

  --���뱾�εǼǵĲ��˻������� 
  If ��¼����_In = 1 Then
    --����ǻ��Ŀ����ݵ�ǰ��¼����Ŀ���,ȡ������(���Ŀ���ڲ�ͬ��λ������,��Ҫ�Զ���������Ա㱣���������) 
    v_��� := 1;
    If v_���Ŀ = 2 Then
      Begin
        Select Nvl(Max(��¼���), 0) + 1
        Into v_���
        From ���˻�������
        Where ��¼id = n_��¼id And ��Ŀ��� = ��Ŀ���_In;
      Exception
        When Others Then
          v_��� := 1;
      End;
    End If;
  
    Insert Into ���˻�������
      (ID, ��¼id, ��¼����, ��Ŀ����, ��Ŀid, ��Ŀ���, ��Ŀ����, ��Ŀ����, ��¼����, ��Ŀ��λ, ��¼���, ��¼��, ���²�λ, ���Ժϸ�, ��ʼ�汾, ��ֹ�汾, ��¼���, δ��˵��)
      Select ���˻�������_Id.Nextval, n_��¼id, ��¼����_In, ������, ��Ŀid, ��Ŀ���, ��Ŀ����, ��Ŀ����, ��¼����_In, Decode(�Ƿ�˵��_In, 1, Null, ��Ŀ��λ),
             ��¼���_In, v_������, ���²�λ_In, ���Ժϸ�_In, n_��ʼ�汾, Null, v_���, δ��˵��_In
      From �����¼��Ŀ
      Where ��Ŀ��� = ��Ŀ���_In;
  Else
    Insert Into ���˻�������
      (ID, ��¼id, ��¼����, ��Ŀ����, ��Ŀid, ��Ŀ���, ��Ŀ����, ��Ŀ����, ��¼����, ��Ŀ��λ, ��¼���, ��¼��, ���²�λ, ���Ժϸ�, ��ʼ�汾, ��ֹ�汾, ��¼���, δ��˵��)
    Values
      (���˻�������_Id.Nextval, n_��¼id, ��¼����_In, Null, Null, 0, Decode(��¼����_In, 2, '�ϱ�˵��', 6, '�±�˵��', 3, '���ת', 4, ��¼����_In),
       Decode(��¼����_In, 3, 0, 1), Decode(��¼����_In, 4, '1', ��¼����_In), '', ��¼���_In, v_������, ���²�λ_In, ���Ժϸ�_In, n_��ʼ�汾, Null, 1,
       δ��˵��_In);
  End If;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���ӻ����¼_Update;
/

--91225:������,2016-01-11,��Ⱦ������ϵͳ
Create Or Replace Procedure Zl_���Ӳ�����¼_Update
(
  Id_In       In ���Ӳ�����¼.Id%Type,
  ������Դ_In In ���Ӳ�����¼.������Դ%Type,
  ����id_In   In ���Ӳ�����¼.����id%Type,
  ��ҳid_In   In ���Ӳ�����¼.��ҳid%Type,
  Ӥ��_In     In ���Ӳ�����¼.Ӥ��%Type,
  ����id_In   In ���Ӳ�����¼.����id%Type,
  �ļ�id_In   In ���Ӳ�����¼.�ļ�id%Type,
  ҽ��id_In   In ����ҽ������.ҽ��id%Type := Null,
  ����ʱ��_In In ���Ӳ�����¼.����ʱ��%Type := Null
) Is
  v_������     ���Ӳ�����¼.������%Type;
  d_����ʱ��   ���Ӳ�����¼.����ʱ��%Type;
  d_����ʱ��   ���Ӳ�����¼.����ʱ��%Type;
  d_���ʱ��   ���Ӳ�����¼.���ʱ��%Type := Null;
  n_���汾   ���Ӳ�����¼.���汾%Type := 1;
  n_Ԥ�����id ���Ӳ�������.Ԥ�����id%Type;
  n_�������id ���Ӳ�������.�������id%Type;
  v_��������   ���Ӳ�������.��������%Type;
  n_����״̬   ���Ӳ�����¼.����״̬%Type;
  e_Submit Exception;
  e_Nofile Exception;
  e_Repeat Exception;

  n_���� �����ļ��б�.����%Type;
  v_���� �����ļ��б�.����%Type;
  v_�¼� ����ʱ��Ҫ��.�¼�%Type;
  n_Ψһ ����ʱ��Ҫ��.Ψһ%Type;
  n_��� Number(1);
  n_Num  Number;
  n_Lab  Number;

  --���Ͳ�����ϼ�¼
  Procedure Put_Pati_Diag
  (
    v_Kind_Emr  In Varchar2,
    n_Kind_Base In ������ϼ�¼.�������%Type,
    n_Del_Old   In Number
  ) Is
    n_����      ������ϼ�¼.�������%Type;
    n_��ҽ      Number(1); --�Ƿ���ҽ��0-��ҽ;1-��ҽ
    n_����id    ������ϼ�¼.����id%Type; --��Ӧ��������Ŀ¼(ICD����ҽ����)��ID
    n_���id    ������ϼ�¼.���id%Type; --��Ӧ�������Ŀ¼��ID
    n_֤��id    ������ϼ�¼.֤��id%Type; --��Ӧ�������Ŀ¼��ID
    n_����      ������ϼ�¼.�Ƿ�����%Type; --�Ƿ����0-ȷ��;1-����
    d_����      ������ϼ�¼.��¼����%Type; --��ϴ���
    n_����      ������ϼ�¼.��ϴ���%Type; --��ϴ���
    v_��Ժ����  ������ϼ�¼.��Ժ����%Type;
    v_��Ժ���  ������ϼ�¼.��Ժ���%Type;
    n_Syncpage  Number(1); --�Ƿ�ͬ�����²�����ҳ 0-��ͬ�� 1-ͬ��
    n_��ҽorder Number(2); --��ҳ��ϴ���
    n_��ҽorder Number(2); --��ҳ��ϴ���
  Begin
    --ȡ���Ƿ�ͬ�����²�����ҳ����
    n_Syncpage := Nvl(zl_GetSysParameter('SyncPage', 1070), 0);
  
    If n_Del_Old = 1 Then
      n_����      := 0;
      n_��ҽorder := 0;
      n_��ҽorder := 0;
    Else
      Select Nvl(Max(��ϴ���), 0)
      Into n_����
      From ������ϼ�¼
      Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼��Դ = 1 And ����id + 0 = Id_In And Nvl(ҽ��id, 0) = Nvl(ҽ��id_In, 0);
      If n_Syncpage = 1 Then
        Select Nvl(Max(��ϴ���), 0)
        Into n_��ҽorder
        From ������ϼ�¼
        Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼��Դ = 3;
      End If;
    End If;
  
    For r_Temp In (Select Rownum As ����, �������� As ����, �����ı� As ����
                   From ��ʱ��������
                   Where �������� = 7 And Substr(��������, 1, 2) = v_Kind_Emr And Nvl(��ֹ��, 0) = 0) Loop
      If n_Del_Old = 1 And r_Temp.���� = 1 Then
        Delete ������ϼ�¼
        Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼��Դ = 1 And ����id + 0 = Id_In And
              ������� In (n_Kind_Base, n_Kind_Base + 10) And Nvl(ҽ��id, 0) = Nvl(ҽ��id_In, 0);
        If n_Syncpage = 1 And (n_Kind_Base = 2 Or n_Kind_Base = 3) Then
          --ֻ������Ժ��Ϻͳ�Ժ���
          Delete ������ϼ�¼
          Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼��Դ = 3 And ������� In (n_Kind_Base, n_Kind_Base + 10);
        End If;
      End If;
      n_��ҽ   := To_Number(Substr(r_Temp.����, Instr(r_Temp.����, ';', 1, 1) + 1,
                                 Instr(r_Temp.����, ';', 1, 2) - Instr(r_Temp.����, ';', 1, 1) - 1));
      n_����id := To_Number(Substr(r_Temp.����, Instr(r_Temp.����, ';', 1, 2) + 1,
                                 Instr(r_Temp.����, ';', 1, 3) - Instr(r_Temp.����, ';', 1, 2) - 1));
      n_���id := To_Number(Substr(r_Temp.����, Instr(r_Temp.����, ';', 1, 3) + 1,
                                 Instr(r_Temp.����, ';', 1, 4) - Instr(r_Temp.����, ';', 1, 3) - 1));
      n_֤��id := To_Number(Substr(r_Temp.����, Instr(r_Temp.����, ';', 1, 4) + 1,
                                 Instr(r_Temp.����, ';', 1, 5) - Instr(r_Temp.����, ';', 1, 4) - 1));
      n_����   := To_Number(Substr(r_Temp.����, Instr(r_Temp.����, ';', 1, 5) + 1,
                                 Instr(r_Temp.����, ';', 1, 6) - Instr(r_Temp.����, ';', 1, 5) - 1));
      d_����   := To_Date(Substr(r_Temp.����, Instr(r_Temp.����, ';', 1, 6) + 1,
                               Instr(r_Temp.����, ';', 1, 7) - Instr(r_Temp.����, ';', 1, 6) - 1), 'yyyy-mm-dd hh24:mi:ss');
      If n_Kind_Base <> 1 And n_Kind_Base <> 2 And n_Kind_Base <> 3 Then
        n_��ҽ := 0;
      End If;
      If n_��ҽ = 1 Then
        n_���� := n_Kind_Base + 10;
      Else
        n_���� := n_Kind_Base;
      End If;
      Insert Into ������ϼ�¼
        (ID, ����id, ��ҳid, ҽ��id, ��¼��Դ, ��ϴ���, ����id, �������, ����id, ���id, ֤��id, �������, �Ƿ�����, ��¼����, ��¼��)
      Values
        (������ϼ�¼_Id.Nextval, ����id_In, ��ҳid_In, ҽ��id_In, 1, r_Temp.���� + n_����, Id_In, n_����,
         Decode(n_����id, 0, Null, n_����id), Decode(n_���id, 0, Null, n_���id), Decode(n_֤��id, 0, Null, n_֤��id), r_Temp.����,
         n_����, d_����, v_������);
      If n_Syncpage = 1 And (n_Kind_Base = 2 Or n_Kind_Base = 3) Then
        If n_Kind_Base = 3 Then
          v_��Ժ���� := Substr(r_Temp.����, Instr(r_Temp.����, ';', 1, 7) + 1,
                           Instr(r_Temp.����, ';', 1, 8) - Instr(r_Temp.����, ';', 1, 7) - 1);
          v_��Ժ��� := Substr(r_Temp.����, Instr(r_Temp.����, ';', 1, 8) + 1);
        End If;
        --�����Ҫͬ����ҳ��ϣ�ֻ������Ժ��Ϻͳ�Ժ���
        If n_��ҽ = 1 Then
          n_��ҽorder := n_��ҽorder + 1;
        Else
          n_��ҽorder := n_��ҽorder + 1;
        End If;
        Insert Into ������ϼ�¼
          (ID, ����id, ��ҳid, ��¼��Դ, ��ϴ���, �������, �������, ����id, ���id, ֤��id, �������, ��Ժ����, ��Ժ���, �Ƿ�����, ��¼����, ��¼��)
        Values
          (������ϼ�¼_Id.Nextval, ����id_In, ��ҳid_In, 3, Decode(n_��ҽ, 1, n_��ҽorder, n_��ҽorder), 1, n_����,
           Decode(n_����id, 0, Null, n_����id), Decode(n_���id, 0, Null, n_���id), Decode(n_֤��id, 0, Null, n_֤��id),
           Replace(r_Temp.����, '(?)', ''), v_��Ժ����, v_��Ժ���, n_����, d_����, v_������);
      End If;
    End Loop;
  End Put_Pati_Diag;

Begin
  Begin
    Select p.���� Into v_������ From �ϻ���Ա�� O, ��Ա�� P Where o.��Աid = p.Id And �û��� = User;
  Exception
    When Others Then
      v_������ := User;
  End;
  d_����ʱ�� := Sysdate;
  d_����ʱ�� := Nvl(����ʱ��_In, Sysdate);

  Select Greatest(Nvl(Max(��ʼ��), 1), Nvl(Max(��ֹ��), 1) + 1) Into n_���汾 From ��ʱ��������;
  If n_���汾 <= 0 Then
    n_���汾 := 1;
  End If;

  Select Count(*) Into n_Num From �����ļ��б� Where ID = �ļ�id_In;
  If n_Num = 0 Then
    Raise e_Nofile;
  End If;

  Select l.����, l.����, q.�¼�, q.Ψһ
  Into n_����, v_����, v_�¼�, n_Ψһ
  From �����ļ��б� L, ����ʱ��Ҫ�� Q
  Where l.Id = q.�ļ�id(+) And l.Id = �ļ�id_In;

  Update ���Ӳ�����¼
  Set ������Դ = ������Դ_In, ����id = ����id_In, ��ҳid = ��ҳid_In, Ӥ�� = Ӥ��_In, ����id = ����id_In, �ļ�id = �ļ�id_In, ����ʱ�� = d_����ʱ��
  Where ID = Id_In;
  If Sql%RowCount = 0 Then
    Insert Into ���Ӳ�����¼
      (ID, ������Դ, ����id, ��ҳid, Ӥ��, ����id, ��������, �ļ�id, ��������, ���汾, ������, ����ʱ��, ������, ����ʱ��)
    Values
      (Id_In, ������Դ_In, ����id_In, ��ҳid_In, Ӥ��_In, ����id_In, n_����, �ļ�id_In, v_����, n_���汾, v_������, d_����ʱ��, v_������, d_����ʱ��);
    If n_���� = 7 And Nvl(ҽ��id_In, 0) <> 0 Then
      --��鱨����ظ���
      Select Count(*)
      Into n_Num
      From ���Ӳ�����¼ L, ����ҽ������ R
      Where l.Id = r.����id And r.ҽ��id = ҽ��id_In And l.�ļ�id = �ļ�id_In;
      If n_Num > 0 Then
        Raise e_Repeat;
      End If;
      --������������ж�������µ�ҽ���ϲ�Ϊһ�����յ����
      Begin
        Select a.Id
        Into n_Lab
        From ����걾��¼ A, ����ҽ����¼ B
        Where a.ҽ��id = b.���id And Rownum <= 1 And a.ҽ��id = ҽ��id_In;
      Exception
        When Others Then
          n_Lab := 0;
      End;
      If n_Lab = 0 Then
        --������Ŀ������������
        Insert Into ����ҽ������ (ҽ��id, ����id) Values (ҽ��id_In, Id_In);
      Else
        --�������������Ŀ
        Insert Into ����ҽ������
          (ҽ��id, ����id)
          Select Distinct b.ҽ��id, Id_In
          From ����걾��¼ A, ������Ŀ�ֲ� B
          Where a.Id = b.�걾id And a.ҽ��id = ҽ��id_In And b.ҽ��id Is Not Null;
      End If;
    End If;
  Else
    Select Nvl(����״̬, 0) Into n_����״̬ From ���Ӳ�����¼ Where ID = Id_In;
    Select Max(����״̬) Into n_Num From �����걨��¼ Where �ļ�id = Id_In;
    If Nvl(n_Num, 0) <> 4 and Nvl(n_Num, 0) <> 5 Then
      If n_����״̬ > 0 Then
        Raise e_Submit;
      End If;
    End If;
  End If;

  Update ���Ӳ�������
  Set ������� = -1 * �������, �����д� = -1 * �����д�, ��ֹ�� = Decode(Nvl(��ֹ��, 0), 0, n_���汾, ��ֹ��)
  Where �ļ�id = Id_In;
  For r_Temp In (Select ID, ��id, ��ʼ��, ��ֹ��, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���, �������id, Ԥ�����id, �������, ʹ��ʱ��,
                        ����Ҫ��id, �滻��, Ҫ������, Ҫ������, Ҫ�س���, Ҫ��С��, Ҫ�ص�λ, Ҫ�ر�ʾ, ������̬, Ҫ��ֵ��
                 From ��ʱ��������
                 Order By ID) Loop
  
    --����Ԥ�����id������ǰ�ļ�(XML����ʷ�ļ�)������Ԥ������뵱ǰϵͳ�����ϡ�
    n_Ԥ�����id := r_Temp.Ԥ�����id;
    If r_Temp.�������� = 1 And Nvl(n_Ԥ�����id, 0) <> 0 Then
      Select Max(ID) Into n_Ԥ�����id From �����ļ��ṹ Where ID = n_Ԥ�����id And �ļ�id Is Null;
      If n_Ԥ�����id = 0 Then
        n_Ԥ�����id := Null;
      End If;
    End If;
    --�޸��������id������������id�����ڣ������������Ʋ��Ҷ�Ӧ�Ķ������id
    n_�������id := r_Temp.�������id;
    If r_Temp.�������� = 1 Then
      If Nvl(n_�������id, 0) <> 0 Then
        Select Max(ID) Into n_�������id From �����ļ��ṹ Where ID = n_�������id And �ļ�id = �ļ�id_In;
      End If;
      If Nvl(n_�������id, 0) = 0 Then
        Select Max(ID)
        Into n_�������id
        From �����ļ��ṹ
        Where �ļ�id = �ļ�id_In And �����ı� || Ԥ�����id = r_Temp.�����ı� || n_Ԥ�����id;
      End If;
      If n_�������id = 0 Then
        n_�������id := Null;
      End If;
    End If;
  
    v_�������� := r_Temp.��������;
    --��ǩ�������ñ����˺����ʱ��
    If r_Temp.�������� = 8 Then
      If Instr(v_��������, ';', 1, 5) = 0 Then
        v_�������� := v_�������� || ';';
      End If;
      If Instr(v_��������, ';', 1, 5) - Instr(v_��������, ';', 1, 4) = 1 Then
        v_�������� := Substr(v_��������, 1, Instr(v_��������, ';', 1, 4) - 1) || ';' || To_Char(d_����ʱ��, 'yyyy-mm-dd hh24:mi:ss') ||
                  Substr(v_��������, Instr(v_��������, ';', 1, 5));
      End If;
      If r_Temp.��ʼ�� >= n_���汾 Then
        If Nvl(Instr(r_Temp.�����ı�, ';'), 0) = 0 Then
          v_������ := r_Temp.�����ı�;
        Else
          --�����ı��д����ǩ����;ID,�п���ǩ��ͬ�����Ա���ʹ��ID,ͬʱȷ����ʷ���ݵĻ���������
          Begin
            Select ���� Into v_������ From ��Ա�� Where ID = Substr(r_Temp.�����ı�, Instr(r_Temp.�����ı�, ';') + 1);
          Exception
            When Others Then
              v_������ := Substr(r_Temp.�����ı�, 1, Instr(r_Temp.�����ı�, ';') - 1);
          End;
        End If;
      End If;
      If d_���ʱ�� Is Null And r_Temp.��ʼ�� = 1 Then
        d_���ʱ�� := To_Date(Substr(v_��������, Instr(v_��������, ';', 1, 4) + 1,
                                 Instr(v_��������, ';', 1, 5) - Instr(v_��������, ';', 1, 4) - 1), 'yyyy-mm-dd hh24:mi:ss');
      End If;
    End If;
  
    Update ���Ӳ�������
    Set ��id = r_Temp.��id, ��ʼ�� = r_Temp.��ʼ��, ��ֹ�� = r_Temp.��ֹ��, ������� = r_Temp.�������, �������� = r_Temp.��������, ������ = r_Temp.������,
        �������� = r_Temp.��������, �������� = v_��������, �����д� = r_Temp.�����д�, �����ı� = r_Temp.�����ı�, �Ƿ��� = r_Temp.�Ƿ���, �������id = n_�������id,
        Ԥ�����id = n_Ԥ�����id, ������� = r_Temp.�������, ʹ��ʱ�� = r_Temp.ʹ��ʱ��, ����Ҫ��id = r_Temp.����Ҫ��id, �滻�� = r_Temp.�滻��,
        Ҫ������ = r_Temp.Ҫ������, Ҫ������ = r_Temp.Ҫ������, Ҫ�س��� = r_Temp.Ҫ�س���, Ҫ��С�� = r_Temp.Ҫ��С��, Ҫ�ص�λ = r_Temp.Ҫ�ص�λ,
        Ҫ�ر�ʾ = r_Temp.Ҫ�ر�ʾ, ������̬ = r_Temp.������̬, Ҫ��ֵ�� = r_Temp.Ҫ��ֵ��
    Where ID = r_Temp.Id And �ļ�id + 0 = Id_In;
    If Sql%RowCount = 0 Then
      Insert Into ���Ӳ�������
        (ID, �ļ�id, ��id, ��ʼ��, ��ֹ��, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���, �������id, Ԥ�����id, �������, ʹ��ʱ��, ����Ҫ��id,
         �滻��, Ҫ������, Ҫ������, Ҫ�س���, Ҫ��С��, Ҫ�ص�λ, Ҫ�ر�ʾ, ������̬, Ҫ��ֵ��)
      Values
        (r_Temp.Id, Id_In, r_Temp.��id, r_Temp.��ʼ��, r_Temp.��ֹ��, r_Temp.�������, r_Temp.��������, r_Temp.������, r_Temp.��������,
         v_��������, r_Temp.�����д�, r_Temp.�����ı�, r_Temp.�Ƿ���, n_�������id, n_Ԥ�����id, r_Temp.�������, r_Temp.ʹ��ʱ��, r_Temp.����Ҫ��id,
         r_Temp.�滻��, r_Temp.Ҫ������, r_Temp.Ҫ������, r_Temp.Ҫ�س���, r_Temp.Ҫ��С��, r_Temp.Ҫ�ص�λ, r_Temp.Ҫ�ر�ʾ, r_Temp.������̬,
         r_Temp.Ҫ��ֵ��);
    Else
      --��ͨ������ڱ༭ʱû�кۼ�������Ԫ���棻�����Ҫ�ָ��ӵ�Ԫ����֤�汾��¼
      If r_Temp.�������� = 3 Then
        n_��� := 0;
        If Instr(v_��������, ';', 1, 18) = 0 Then
          n_��� := 1;
        Elsif Substr(v_��������, Instr(v_��������, ';', 1, 18) + 1, 1) = '0' Then
          n_��� := 1;
        End If;
        If n_��� = 1 Then
          Update ���Ӳ�������
          Set ������� = Abs(�������), �����д� = Abs(�����д�)
          Where �ļ�id = Id_In And ��id = r_Temp.Id And ��ʼ�� <= n_���汾 And �������� <> 5;
        End If;
      End If;
    End If;
  End Loop;
  Delete ���Ӳ�������
  Where (Nvl(�������, 0) < 0 Or Nvl(�����д�, 0) < 0 Or Nvl(��ʼ��, 1) > n_���汾) And �ļ�id = Id_In;

  Update ���Ӳ�����¼
  Set ���ʱ�� = d_���ʱ��, ������ = v_������, ���汾 = n_���汾,
      ǩ������ =
       (Select Nvl(Sum(Power(2, Ҫ�ر�ʾ - 1)), 0)
        From (Select Distinct Ҫ�ر�ʾ From ��ʱ�������� Where �������� = 8 And ��ʼ�� >= n_���汾))
  Where ID = Id_In;

  --��ɾ��ԭ����ϣ���Ϊ�п���ԭ����ϱ�ɾ�������
  Delete ������ϼ�¼ Where ����id = ����id_In And ��ҳid = ��ҳid_In And ����id + 0 = Id_In;
  --��д������ϼ�¼
  If n_���� = 1 Then
    Put_Pati_Diag('11', 1, 1);
  Elsif n_���� = 2 And (v_�¼� = '��Ժ' Or v_�¼� = '�״���Ժ' Or v_�¼� = '�ٴ���Ժ') And n_Ψһ = 1 Then
    Put_Pati_Diag('21', 2, 1);
    Put_Pati_Diag('22', 2, 1);
    Put_Pati_Diag('23', 2, 1);
    Put_Pati_Diag('24', 2, 0);
  Elsif n_���� = 2 And (v_�¼� = '24Сʱ��Ժ' Or v_�¼� = '24Сʱ����') Then
    Put_Pati_Diag('21', 2, 1);
    Put_Pati_Diag('22', 2, 1);
    Put_Pati_Diag('23', 2, 1);
    Put_Pati_Diag('24', 2, 0);
    Put_Pati_Diag('31', 3, 1);
  Elsif n_���� = 2 And (v_�¼� = '��Ժ' Or v_�¼� = '����') Then
    Put_Pati_Diag('31', 3, 1);
  Elsif n_���� = 2 And v_�¼� = '����' Then
    Put_Pati_Diag('41', 8, 1);
    Put_Pati_Diag('42', 9, 1);
  Elsif n_���� = 7 And (ҽ��id_In Is Not Null) Then
    Put_Pati_Diag('51', 6, 1);
    Put_Pati_Diag('52', 22, 1);
    --ֻ�������Ա�־
    --Update ����ҽ������ Set ������� = 0 Where ҽ��id = ҽ��id_In;
    Update ����ҽ������
    Set ������� = 1
    Where ҽ��id = ҽ��id_In And Exists
     (Select �����ı�
           From ��ʱ��������
           Where �������� = 7 And (Substr(��������, 1, 2) = '51' Or Substr(��������, 1, 2) = '52') And Nvl(��ֹ��, 0) = 0);
  End If;

  --������Ӳ���ʱ��
  If d_���ʱ�� Is Null Then
    Update ���Ӳ���ʱ�� Set ���ʱ�� = Null Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��ɼ�¼id = Id_In;
    If Sql%RowCount = 0 Then
      Zl_���Ӳ���ʱ��_Update(����id_In, ��ҳid_In, ������Դ_In, ����id_In, �ļ�id_In, Id_In, Null, v_������);
    End If;
  Else
    Zl_���Ӳ���ʱ��_Update(����id_In, ��ҳid_In, ������Դ_In, ����id_In, �ļ�id_In, Id_In, d_���ʱ��, v_������);
  End If;
Exception
  When e_Submit Then
    Raise_Application_Error(-20101, '[ZLSOFT]���ܸ��ı��������յĲ�����[ZLSOFT]');
  When e_Nofile Then
    Raise_Application_Error(-20101, '[ZLSOFT]�����ļ����嶪ʧ������ϵϵͳ����Ա��[ZLSOFT]');
  When e_Repeat Then
    Raise_Application_Error(-20101, '[ZLSOFT]�������Ѿ���д�������˱��棬�����ٱ��棡[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���Ӳ�����¼_Update;
/

--91844:����,2016-01-07,���µ�û�л�����ϸ,�����ļ�������ȷɾ��
CREATE OR REPLACE Procedure ZL_���˻����ļ�_DELETE(
	ID_IN IN ���˻����ļ�.ID%Type 
) 
IS 
	ERR_ITEM Exception; 
	V_ERR_MSG  VARCHAR2(500); 
	LNGSIGNED NUMBER ; 
Begin 
	--���������������ɾ�� 
	Begin 
		SELECT 1 INTO LNGSIGNED 
		FROM ���˻������� A,���˻����ļ� B,���˻�����ϸ C
		WHERE B.ID=ID_IN And A.�ļ�ID=B.ID And C.��¼ID = a.id And RowNum<2; 
	Exception 
		When Others Then LNGSIGNED:=0; 
	End ; 
 
	IF LNGSIGNED=1 THEN 
		V_ERR_MSG := '���ļ��Ѿ������������ݲ�����ɾ��,���飡'; 
		RAISE ERR_ITEM; 
	End IF ; 
 
	--ɾ����ӡ���� 
	DELETE ���˻����ӡ WHERE �ļ�ID=ID_IN; 
	--ɾ����ϸ���� 
	DELETE ���˻�����ϸ WHERE ��¼ID IN (SELECT ID FROM ���˻������� WHERE �ļ�ID=ID_IN); 
	--ɾ���м�¼ 
	DELETE ���˻������� WHERE �ļ�ID=ID_IN; 
	--ɾ�������ļ� 
	DELETE ���˻����ļ� WHERE ID=ID_IN; 
	--���ϼ����ݵ�����ID����Ϊ�� 
	UPDATE ���˻����ļ� SET ����ID=NULL WHERE ����ID=ID_IN; 
Exception 
	WHEN ERR_ITEM THEN 
		RAISE_APPLICATION_ERROR(-20101, '[ZLSOFT]' || V_ERR_MSG || '[ZLSOFT]'); 
	When Others Then 
		ZL_ERRORCENTER (SQLCODE, SQLERRM); 
End ZL_���˻����ļ�_DELETE;
/

--92469:������,2016-01-07,ȥ��һ������ �ͼ���_In
Create Or Replace Procedure Zl_����ҽ������_Sampleinput
(
  ҽ��id      In Varchar2,
  ������_In   In ����ҽ������.������%Type := Null,
  ��������_In In ����ҽ������.��������%Type := 0,
  ��Ա���_In In ��Ա��.���%Type := Null,
  ��Ա����_In In ��Ա��.����%Type := Null
) Is
  --δ��˵ķ�����(������ҩƷ)
  Cursor c_Verify(v_ҽ��id In Number) Is
    Select Distinct ��¼����, NO, ���
    From סԺ���ü�¼
    Where �շ���� Not In ('5', '6', '7') And ҽ����� + 0 = v_ҽ��id And ���ʷ��� = 1 And ��¼״̬ = 0 And �۸񸸺� Is Null And
          (��¼����, NO) In (Select ��¼����, NO
                         From ����ҽ������
                         Where ҽ��id = v_ҽ��id
                         Union All
                         Select ��¼����, NO
                         From ����ҽ������
                         Where ҽ��id In (Select ID From ����ҽ����¼ Where v_ҽ��id In (ID, ���id)))
    Union All
    Select Distinct ��¼����, NO, ���
    From ������ü�¼
    Where �շ���� Not In ('5', '6', '7') And ҽ����� + 0 = v_ҽ��id And ���ʷ��� = 1 And ��¼״̬ = 0 And �۸񸸺� Is Null And
          (��¼����, NO) In (Select ��¼����, NO
                         From ����ҽ������
                         Where ҽ��id = v_ҽ��id
                         Union All
                         Select ��¼����, NO
                         From ����ҽ������
                         Where ҽ��id In (Select ID From ����ҽ����¼ Where v_ҽ��id In (ID, ���id)))
    Order By ��¼����, NO, ���;

  --���ҵ�ǰ�걾���������
  Cursor c_Samplequest(v_ҽ��id In Number) Is
    Select Distinct ID As ҽ��id, ������Դ From ����ҽ����¼ Where v_ҽ��id In (ID, ���id);

  v_ִ�� Number(1);
  v_No   ����ҽ������.No%Type;
  v_���� ����ҽ������.��¼����%Type;
  v_��� Varchar2(1000);

  v_ҽ��id   ����ҽ������.ҽ��id%Type;
  v_���id   ����ҽ����¼.���id%Type;
  v_�������� ����ҽ������.��¼����%Type;
  v_�������� ����ҽ������.��������%Type;
  v_Records  Varchar2(2000);
  v_Currrec  Varchar2(50);
  v_Fields   Varchar2(50);
  v_Count    Number(18);
  v_����id   ����ҽ����¼.����id%Type;
  v_��ҳid   ����ҽ����¼.��ҳid%Type;
  v_�Ƿ��Ժ Number; --0=��Ժ,1=��Ժ
  v_��¼״̬ Number;
  v_������Դ ����ҽ����¼.������Դ%Type;
  v_Date     Date;
  Err_Custom Exception;
  v_Error Varchar2(100);
Begin
  Select Sysdate Into v_Date From Dual;
  --ִ�к��Զ���˶�Ӧ�ļ��ʻ��۵�(������ҩƷ)
  Select Zl_To_Number(Nvl(Zl_Getsysparameter(81), '0')) Into v_ִ�� From Dual;

  v_Records := ҽ��id || '|';

  While v_Records Is Not Null Loop
  
    v_Currrec := Substr(v_Records, 1, Instr(v_Records, '|') - 1);
    v_Fields  := v_Currrec;
    v_ҽ��id  := Substr(v_Fields, 1, Instr(v_Fields, ',') - 1);
    v_���id  := Substr(v_Fields, Instr(v_Fields, ',') + 1);
    If ������_In Is Null Then
      Update ����ҽ������
      Set ������ = Null, ����ʱ�� = Null, �������� = Null
      Where ҽ��id In (v_ҽ��id, v_���id);
      Update ����ҽ������
      Set ִ��״̬ = Decode(��������, Null, 0, 1)
      Where ҽ��id In (Select ID From ����ҽ����¼ Where ID In (v_ҽ��id, v_���id) And ���id Is Null);
      For r_Samplequest In c_Samplequest(v_���id) Loop
        If r_Samplequest.������Դ = 2 Then
          Select Decode(��¼����, 1, 1, Decode(�������, 1, 1, 2))
          Into v_��������
          From ����ҽ������
          Where ҽ��id = r_Samplequest.ҽ��id;
        Else
          v_�������� := 1;
        End If;
        If v_�������� = 2 Then
          --2.����ִ�д���
          Update סԺ���ü�¼
          Set ִ��״̬ = 0, ִ��ʱ�� = Null, ִ���� = ������_In
          Where �շ���� Not In ('5', '6', '7') And
                (ҽ�����, ��¼����, NO) In
                (Select ҽ��id, ��¼����, NO
                 From ����ҽ������
                 Where ҽ��id = r_Samplequest.ҽ��id
                 Union All
                 Select ҽ��id, ��¼����, NO
                 From ����ҽ������
                 Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = r_Samplequest.ҽ��id And ���id Is Not Null)
                 Union All
                 Select ҽ��id, ��¼����, NO
                 From ����ҽ������
                 Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = r_Samplequest.ҽ��id And ���id Is Null) And ������ Is Null);
        Else
          Update ������ü�¼
          Set ִ��״̬ = 0, ִ��ʱ�� = Null, ִ���� = ������_In
          Where �շ���� Not In ('5', '6', '7') And
                (ҽ�����, ��¼����, NO) In
                (Select ҽ��id, ��¼����, NO
                 From ����ҽ������
                 Where ҽ��id = r_Samplequest.ҽ��id
                 Union All
                 Select ҽ��id, ��¼����, NO
                 From ����ҽ������
                 Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = r_Samplequest.ҽ��id And ���id Is Not Null)
                 Union All
                 Select ҽ��id, ��¼����, NO
                 From ����ҽ������
                 Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = r_Samplequest.ҽ��id And ���id Is Null) And ������ Is Null);
        End If;
      End Loop;
    Else
      --�ж��Ƿ��ѳ�Ժ������ѳ�Ժ������ɵǼ�
      Begin
        If v_��ҳid Is Null Then
          Select a.����id, a.��ҳid, a.������Դ
          Into v_����id, v_��ҳid, v_������Դ
          From ����ҽ����¼ A, ������ҳ B
          Where a.����id = b.����id And a.��ҳid = b.��ҳid(+) And a.Id = v_ҽ��id;
        End If;
      Exception
        When Others Then
          v_������Դ := 1;
      End;
      If v_������Դ = 2 Then
        If Nvl(v_��ҳid, 0) > 0 Then
          Select Decode(��Ժ����, Null, 1, 0)
          Into v_�Ƿ��Ժ
          From ������ҳ
          Where ����id = v_����id And ��ҳid = v_��ҳid;
        Else
          v_�Ƿ��Ժ := 0;
        End If;
      
        If v_�Ƿ��Ժ = 0 Then
          --��Ժ�ĲŴ���
          Begin
            Select Nvl(��¼״̬, 0)
            Into v_��¼״̬
            From סԺ���ü�¼
            Where ҽ����� = v_ҽ��id And Nvl(��¼״̬, 0) = 0 And Rownum = 1;
          Exception
            When Others Then
              v_��¼״̬ := 1;
          End;
        
          Select Nvl(��������, 0) Into v_�������� From ����ҽ������ Where ҽ��id = v_ҽ��id;
          If v_�������� = 0 Then
            v_Error := '�����ѳ�Ժ������ɵǼ�!';
            Raise Err_Custom;
          End If;
        
        End If;
      End If;
    
      Update ����ҽ������
      Set ������ = ������_In, ����ʱ�� = v_Date, �������� = ��������_In,  �زɱ걾 = Null
      Where ҽ��id In (v_ҽ��id, v_���id);
      Update ����ҽ������
      Set ִ��״̬ = 1
      Where ҽ��id In (Select ID From ����ҽ����¼ Where ID In (v_ҽ��id, v_���id) And ���id Is Null);
      --���ʻ��۵��Ƿ�תΪ���ʵ�
      --2.��鵱ǰ�걾��ص��������ر걾�Ƿ�������
      For r_Samplequest In c_Samplequest(v_���id) Loop
        v_Count := 0;
        --r_SampleQuest.ҽ��id�����Ѿ����,�����������
        If v_Count = 0 Then
          If r_Samplequest.������Դ = 2 Then
            Select Decode(��¼����, 1, 1, Decode(�������, 1, 1, 2))
            Into v_��������
            From ����ҽ������
            Where ҽ��id = r_Samplequest.ҽ��id;
          Else
            v_�������� := 1;
          End If;
          If v_�������� = 2 Then
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
                   Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = r_Samplequest.ҽ��id And ���id Is Not Null)
                   Union All
                   Select ҽ��id, ��¼����, NO
                   From ����ҽ������
                   Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = r_Samplequest.ҽ��id And ���id Is Null) And ������ Is Null);
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
                   Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = r_Samplequest.ҽ��id And ���id Is Not Null)
                   Union All
                   Select ҽ��id, ��¼����, NO
                   From ����ҽ������
                   Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = r_Samplequest.ҽ��id And ���id Is Null) And ������ Is Null);
          End If;
          --3.�Զ���˼���
          If v_ִ�� = 1 Then
            For r_Verify In c_Verify(r_Samplequest.ҽ��id) Loop
              If r_Verify.No || ',' || r_Verify.��¼���� <> v_No || ',' || v_���� Then
                If v_��� Is Not Null Then
                  If v_�������� = 1 Then
                    Zl_������ʼ�¼_Verify(v_No, ��Ա���_In, ��Ա����_In, Substr(v_���, 2));
                  Elsif v_�������� = 2 Then
                    Zl_סԺ���ʼ�¼_Verify(v_No, ��Ա���_In, ��Ա����_In, Substr(v_���, 2));
                  End If;
                End If;
                v_��� := Null;
              End If;
              v_No   := r_Verify.No;
              v_���� := r_Verify.��¼����;
              v_��� := v_��� || ',' || r_Verify.���;
            End Loop;
            If v_��� Is Not Null Then
              If v_�������� = 1 Then
                Zl_������ʼ�¼_Verify(v_No, ��Ա���_In, ��Ա����_In, Substr(v_���, 2));
              Elsif v_�������� = 2 Then
                Zl_סԺ���ʼ�¼_Verify(v_No, ��Ա���_In, ��Ա����_In, Substr(v_���, 2));
              End If;
            End If;
          End If;
        End If;
      End Loop;
    End If;
    v_Records := Substr('|' || v_Records, Length('|' || v_Currrec || '|') + 1);
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_����ҽ������_Sampleinput;
/

--92410:������,2016-01-05,����״̬����
Create Or Replace Procedure Zl_����תסԺ_����������
(
  No_In         סԺ���ü�¼.No%Type,
  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
  �˷�ʱ��_In   סԺ���ü�¼.����ʱ��%Type,
  �����˷�_In   Number := 0,
  ��Ժ����id_In סԺ���ü�¼.��������id%Type := Null,
  ��ҳid_In     סԺ���ü�¼.��ҳid%Type := Null,
  �����˷�_In   Number := 0,
  ����id_In     ����Ԥ����¼.����id%Type := Null
) As
  v_����ids    Varchar2(3000);
  n_��id       ����ɿ����.Id%Type;
  n_����       Number;
  v_Ԥ��no     ����Ԥ����¼.No%Type;
  n_����id     ������ü�¼.����id%Type;
  v_Nos        Varchar2(3000);
  v_Info       Varchar2(5000);
  v_��ǰ����   Varchar2(3000);
  v_ԭ����ids  Varchar2(5000);
  n_Tempid     ����Ԥ����¼.Id%Type;
  v_��ˮ��     ����Ԥ����¼.������ˮ��%Type;
  v_˵��       ����Ԥ����¼.����˵��%Type;
  n_Ԥ��id     ����Ԥ����¼.Id%Type;
  n_ԭԤ��id   ����Ԥ����¼.Id%Type;
  n_����id     ������Ϣ.����id%Type;
  n_ԭ����id   ����Ԥ����¼.����id%Type;
  n_�������   ����Ԥ����¼.��Ԥ��%Type;
  n_�����     ����Ԥ����¼.�����id%Type;
  n_������     Number;
  n_����ֵ     ��Ա�ɿ����.���%Type;
  v_���㷽ʽ   ����Ԥ����¼.���㷽ʽ%Type;
  n_�����id   ҽ�ƿ����.Id%Type;
  v_����       ����Ԥ����¼.����%Type;
  v_������ˮ�� ����Ԥ����¼.������ˮ��%Type;
  v_����˵��   ����Ԥ����¼.����˵��%Type;
  n_ԭ����     Number;
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
    For v_У�� In (Select Min(a.Id) As Ԥ��id, c.���ѿ�id, Sum(c.������) As ������, c.�ӿڱ��, c.����, Min(c.���) As ���, Min(c.Id) As ID
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
  If ����id_In Is Null Then
    Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
  Else
    n_����id := ����id_In;
  End If;

  Select ����id, ����id
  Into n_ԭ����id, n_����id
  From ������ü�¼
  Where NO = No_In And ��¼���� = 1 And ��¼״̬ In (1, 3) And Rownum < 2;

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
                                                   a.����id = b.����id))) Loop
    v_ԭ����ids := v_ԭ����ids || ',' || r_����id.����id;
  End Loop;
  v_ԭ����ids := Substr(v_ԭ����ids, 2);

  Begin
    Select ժҪ
    Into v_Info
    From ����Ԥ����¼
    Where ���㷽ʽ Is Null And ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id;
  Exception
    When Others Then
      v_Info := '';
  End;
  --����������Ϣ
  If v_Info Is Not Null Then
    While v_Info Is Not Null Loop
      v_��ǰ���� := Substr(v_Info, 1, Instr(v_Info, '|') - 1);
      n_������   := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1));
    
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
      n_�����   := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1));
    
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
      n_������� := -1 * To_Number(v_��ǰ����);
    
      If n_������ = 0 Then
        --���ѿ�
        Select ���㷽ʽ
        Into v_���㷽ʽ
        From ����Ԥ����¼
        Where ����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And ���㿨��� = n_����� And Rownum < 2;
        Zl_Square_Update(v_ԭ����ids, n_����id, n_��id, �˷�ʱ��_In, -1 * n_����id, Null, n_�������, n_�����);
        Update ��Ա�ɿ����
        Set ��� = Nvl(���, 0) - n_�������
        Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_���㷽ʽ
        Returning ��� Into n_����ֵ;
        If Sql%RowCount = 0 Then
          Insert Into ��Ա�ɿ����
            (�տ�Ա, ���㷽ʽ, ����, ���)
          Values
            (����Ա����_In, v_���㷽ʽ, 1, -1 * n_�������);
          n_����ֵ := n_�������;
        End If;
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete From ��Ա�ɿ����
          Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_���㷽ʽ And Nvl(���, 0) = 0;
        End If;
      Else
        --���㿨
        Select ���㷽ʽ, �����id, ����, ������ˮ��, ����˵��
        Into v_���㷽ʽ, n_�����id, v_����, v_������ˮ��, v_����˵��
        From ����Ԥ����¼
        Where ����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And �����id = n_����� And Rownum < 2;
        If Nvl(�����˷�_In, 0) = 1 Then
          If �����˷�_In = 0 Then
            v_Err_Msg := '�����޷����ֵ������˻�,�޷������˷�!';
            Raise Err_Item;
          End If;
          Update ����Ԥ����¼
          Set ��Ԥ�� = ��Ԥ�� - n_�������
          Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_���㷽ʽ;
          If Sql%RowCount = 0 Then
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ,
               �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, У�Ա�־, ��������)
            Values
              (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * n_�������, v_���㷽ʽ, Null, �˷�ʱ��_In,
               Null, Null, Null, ����Ա���_In, ����Ա����_In, '', n_��id, n_�����id, Null, v_����, v_������ˮ��, v_����˵��, Null, n_����id,
               -1 * n_����id, 0, 3);
          End If;
          Update ��Ա�ɿ����
          Set ��� = Nvl(���, 0) - n_�������
          Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_���㷽ʽ
          Returning ��� Into n_����ֵ;
          If Sql%RowCount = 0 Then
            Insert Into ��Ա�ɿ����
              (�տ�Ա, ���㷽ʽ, ����, ���)
            Values
              (����Ա����_In, v_���㷽ʽ, 1, -1 * n_�������);
            n_����ֵ := -1 * n_�������;
          End If;
          If Nvl(n_����ֵ, 0) = 0 Then
            Delete From ��Ա�ɿ����
            Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_���㷽ʽ And Nvl(���, 0) = 0;
          End If;
        Else
          Begin
            Select 1 Into n_���� From ҽ�ƿ���� Where ID = n_�����id And �Ƿ����� = 1;
          Exception
            When Others Then
              n_���� := 0;
          End;
        
          If �����˷�_In = 1 Or n_���� = 0 Then
            v_���㷽ʽ := v_���㷽ʽ;
            n_ԭ����   := 1;
          Else
            n_ԭ���� := 0;
            Begin
              Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And ���� Like '%�ֽ�%' And Rownum < 2;
            Exception
              When Others Then
                Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And Rownum < 2;
            End;
          End If;
        
          If �����˷�_In = 0 Then
            If n_ԭ���� = 1 Then
              Select ������ˮ��, ����˵��, ID
              Into v_��ˮ��, v_˵��, n_ԭԤ��id
              From ����Ԥ����¼
              Where ����id = n_ԭ����id And ���㷽ʽ = v_���㷽ʽ And Rownum < 2;
            
              Update ����Ԥ����¼
              Set ��Ԥ�� = ��Ԥ�� - n_�������
              Where ��¼���� = 3 And ��¼״̬ = 2 And �տ�ʱ�� = �˷�ʱ��_In And ����id + 0 = n_����id And ���㷽ʽ = v_���㷽ʽ And ����id = n_����id;
              If Sql%RowCount = 0 Then
                Insert Into ����Ԥ����¼
                  (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                   ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, У�Ա�־, ��������)
                Values
                  (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * n_�������, v_���㷽ʽ, Null, �˷�ʱ��_In,
                   Null, Null, Null, ����Ա���_In, ����Ա����_In, '', n_��id, n_�����id, Null, v_����, v_������ˮ��, v_����˵��, Null, n_����id,
                   -1 * n_����id, 0, 3);
              End If;
            
              Update ����Ԥ����¼
              Set ��� = ��� + n_�������
              Where ��¼���� = 1 And ��¼״̬ = 1 And �տ�ʱ�� = �˷�ʱ��_In And ����id + 0 = n_����id And ���㷽ʽ = v_���㷽ʽ;
              If Sql%RowCount = 0 Then
                v_Ԥ��no := Nextno(11);
                Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
                Insert Into ����Ԥ����¼
                  (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                   ժҪ, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ)
                Values
                  (n_Ԥ��id, v_Ԥ��no, Null, 1, 1, n_����id, ��ҳid_In, ��Ժ����id_In, n_�������, v_���㷽ʽ, Null, �˷�ʱ��_In, Null, Null,
                   Null, ����Ա���_In, ����Ա����_In, '����תסԺԤ��', n_��id, 2, n_�����id, Null, v_����, v_��ˮ��, v_˵��, Null);
                Update �������㽻�� Set ����id = n_Ԥ��id Where ����id = n_ԭԤ��id;
              End If;
            Else
              Update ����Ԥ����¼
              Set ��Ԥ�� = ��Ԥ�� - n_�������
              Where ��¼���� = 3 And ��¼״̬ = 2 And �տ�ʱ�� = �˷�ʱ��_In And ����id + 0 = n_����id And ���㷽ʽ = v_���㷽ʽ And ����id = n_����id;
              If Sql%RowCount = 0 Then
                Insert Into ����Ԥ����¼
                  (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                   ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, У�Ա�־, ��������)
                Values
                  (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * n_�������, v_���㷽ʽ, Null, �˷�ʱ��_In,
                   Null, Null, Null, ����Ա���_In, ����Ա����_In, '', n_��id, Null, Null, Null, Null, Null, Null, n_����id,
                   -1 * n_����id, 0, 3);
              End If;
            
              Update ����Ԥ����¼
              Set ��� = ��� + n_�������
              Where ��¼���� = 1 And ��¼״̬ = 1 And �տ�ʱ�� = �˷�ʱ��_In And ����id + 0 = n_����id And ���㷽ʽ = v_���㷽ʽ;
              If Sql%RowCount = 0 Then
                v_Ԥ��no := Nextno(11);
                Insert Into ����Ԥ����¼
                  (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                   ժҪ, �ɿ���id, Ԥ�����)
                Values
                  (����Ԥ����¼_Id.Nextval, v_Ԥ��no, Null, 1, 1, n_����id, ��ҳid_In, ��Ժ����id_In, n_�������, v_���㷽ʽ, Null, �˷�ʱ��_In,
                   Null, Null, Null, ����Ա���_In, ����Ա����_In, '����תסԺԤ��', n_��id, 2);
              End If;
            End If;
          
            --�������
            Update �������
            Set Ԥ����� = Nvl(Ԥ�����, 0) + n_�������
            Where ���� = 1 And ����id = n_����id And ���� = 2
            Returning Ԥ����� Into n_����ֵ;
            If Sql%RowCount = 0 Then
              Insert Into ������� (����id, ����, ����, Ԥ�����, �������) Values (n_����id, 1, 2, n_�������, 0);
              n_����ֵ := n_�������;
            End If;
            If Nvl(n_����ֵ, 0) = 0 Then
              Delete From �������
              Where ����id = n_����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
            End If;
          End If;
          --4.2�ɿ����ݴ���
          --   ��Ϊû��ʵ���ղ��˵�Ǯ,���Բ�����
          --�����˷��������ԭԤ����¼
          If �����˷�_In = 1 Then
            Update ��Ա�ɿ����
            Set ��� = Nvl(���, 0) - n_�������
            Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_���㷽ʽ
            Returning ��� Into n_����ֵ;
            If Sql%RowCount = 0 Then
              Insert Into ��Ա�ɿ����
                (�տ�Ա, ���㷽ʽ, ����, ���)
              Values
                (����Ա����_In, v_���㷽ʽ, 1, -1 * n_�������);
              n_����ֵ := -1 * n_�������;
            End If;
            If Nvl(n_����ֵ, 0) = 0 Then
              Delete From ��Ա�ɿ����
              Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_���㷽ʽ And Nvl(���, 0) = 0;
            End If;
            Update ����Ԥ����¼
            Set ��Ԥ�� = ��Ԥ�� + (-1 * n_�������)
            Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_���㷽ʽ;
            If Sql%RowCount = 0 Then
              Insert Into ����Ԥ����¼
                (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ,
                 �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, У�Ա�־, ��������)
              Values
                (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * n_�������, v_���㷽ʽ, Null, �˷�ʱ��_In,
                 Null, Null, Null, ����Ա���_In, ����Ա����_In, '', n_��id, n_�����id, Null, v_����, v_������ˮ��, v_����˵��, Null, n_����id,
                 -1 * n_����id, 0, 3);
            End If;
          End If;
        End If;
      End If;
      v_Info := Substr(v_Info, Instr(v_Info, '|') + 1);
    End Loop;
  End If;

  Delete From ����Ԥ����¼ Where ����id = n_����id And ��¼״̬ = 2 And ���㷽ʽ Is Null;
  Update ������ü�¼ Set ����״̬ = 0 Where ����id = n_����id;
  Update ������ü�¼ Set ����״̬ = 0 Where NO = No_In And Mod(��¼����, 10) = 1 And ��¼״̬ = 2;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����תסԺ_����������;
/

--91225:������,2016-01-04,��Ⱦ������ϵͳ�ڼ����걨��¼����������ֶ�
CREATE OR REPLACE Procedure Zl_�����걨��¼_Incept
(
  �ļ�id_In     In �����걨��¼.�ļ�id%Type,
  Incept_In     In Number, --���ջ��Ǿܾ�
  ˵��_In       In �����걨��¼.�վ�˵��%Type,
  �ĵ�id_In     In Varchar2,
  ����id_In     In �����걨��¼.����ID%Type,
  ��ҳID_In     In �����걨��¼.��ҳID%Type,
  ������Դ_In   In �����걨��¼.������Դ%Type,
  Emrcontent_In In Varchar2  --�²�����ϴ�
) Is
  v_�վ���   ��Ա��.����%Type;

  v_����      �����걨��¼.����%Type;
  v_�Ա�      �����걨��¼.�Ա�%Type;
  v_����      �����걨��¼.����%Type;
  v_ְҵ      �����걨��¼.ְҵ%Type;
  v_��ͥ��ַ  �����걨��¼.��ͥ��ַ%Type;
  v_��ͥ�绰  �����걨��¼.��ͥ�绰%Type;
  v_��������  �����걨��¼.��������%Type;
  v_ȷ������  �����걨��¼.ȷ������%Type;
  v_�������1 �����걨��¼.�������1%Type;
  v_�������2 �����걨��¼.�������2%Type;
  v_���ע  �����걨��¼.���ע%Type;
  v_�����ı�  ���Ӳ�������.�����ı�%Type;
  v_��������  �����걨��¼.��������%Type;
  v_����ҽ��  �����걨��¼.����ҽ��%Type;

  v_Count Number;
  e_Changed Exception;

  Function Trimlen
  (
    Str_In Varchar2,
    Len_In Number
  ) Return Varchar2 Is
    v_Temp Varchar2(4000);
  Begin
    If Str_In Is Not Null Then
      For I In 1 .. Length(Str_In) Loop
        If Lengthb(v_Temp || Substr(Str_In, I, 1)) <= Len_In Then
          v_Temp := v_Temp || Substr(Str_In, I, 1);
        Else
          Exit;
        End If;
      End Loop;
    End If;
    Return v_Temp;
  End Trimlen;
Begin

  Select ���� Into v_�վ��� From ��Ա�� P, �ϻ���Ա�� U Where p.Id = u.��Աid And u.�û��� = User And Rownum < 2;

  If Length(�ĵ�id_In) <> 32 Then
    --�²���ID��32λGUID
    Update ���Ӳ�����¼ Set ����״̬ = Decode(Incept_In, 1, 1, -1) Where ID = �ļ�id_In;
    If Sql%RowCount = 0 Then
      Raise e_Changed;
    End If;
  End If;

  --�Զ���ȡ�걨�����е���Ŀ����
  If Incept_In = 1 Then
    If Length(�ĵ�id_In) <> 32 Then
      --�̶���ӦҪ��
      v_Count := 0;
      For r_Item In (Select Ҫ������, Ҫ������, �����д�, �����ı�
                     From ���Ӳ�������
                     Where (�������� = 4 or �������� = 8 )And �ļ�id = �ļ�id_In
                     Order By �������, �����д�) Loop

        If r_Item.Ҫ������ = '����' Then
          v_���� := Trimlen(r_Item.�����ı�, 20);
        Elsif r_Item.Ҫ������ = '�Ա�' Then
          v_�Ա� := Trimlen(r_Item.�����ı�, 4);
        Elsif r_Item.Ҫ������ = '����' Then
          v_���� := Trimlen(r_Item.�����ı�, 10);
        Elsif r_Item.Ҫ������ = 'ְҵ'  Then
          v_ְҵ := Trimlen(r_Item.�����ı�, 80);
        Elsif r_Item.Ҫ������ = '��ͥ��ַ' Then
          v_��ͥ��ַ := Trimlen(r_Item.�����ı�, 100);
        Elsif r_Item.Ҫ������ = '��ͥ�绰'  Then
          v_��ͥ�绰 := Trimlen(r_Item.�����ı�, 20);
        Elsif r_Item.Ҫ������ = '��ǰ����'  Then
          v_Count := v_Count + 1;
          If v_Count = 1  Then
            --�����е�1��"��ǰ����"��Ϊ��������
            Begin
              v_�������� := To_Date(Replace(Replace(Replace(r_Item.�����ı�, '��', '-'), '��', '-'), '��', ''), 'YYYY-MM-DD');
            Exception
              When Others Then
                Null;
            End;
          Elsif v_Count = 2  Then
            --�����е�2��"��ǰ����"��Ϊȷ������
            Begin
              v_ȷ������ := To_Date(Replace(Replace(Replace(r_Item.�����ı�, '��', '-'), '��', '-'), '��', ''), 'YYYY-MM-DD');
            Exception
              When Others Then
                Null;
            End;
          End If;
        Elsif r_Item.Ҫ������ = '������Ⱦ��' Then
          v_�������1 := Trimlen(r_Item.�����ı�, 150);
        End If;
      End Loop;

        --������ʱҪ�ض�Ӧ
      For r_Item In (Select �걨��Ŀ, ��ӦҪ�� From �����걨��Ӧ) Loop
        Begin
          Select �����ı�
          Into v_�����ı�
          From ���Ӳ�������
          Where �������� = 4 And ����Ҫ��id Is Null And Ҫ������ = r_Item.��ӦҪ�� And �ļ�id = �ļ�id_In;
        Exception
          When Others Then
            v_�����ı� := Null;
        End;

        If r_Item.�걨��Ŀ = '�������2' Then
          v_�������2 := Trimlen(v_�����ı�, 150);
        Elsif r_Item.�걨��Ŀ = '���ע' Then
          v_���ע := Trimlen(v_�����ı�, 100);
        End If;
      End Loop;
    Else
      Select ����, �Ա�, ����, ְҵ, ��ͥ��ַ, ��ͥ�绰, ��ͥ�绰
      Into v_����, v_�Ա�, v_����, v_ְҵ, v_��ͥ��ַ, v_��ͥ�绰, v_��ͥ�绰
      From ������Ϣ
      Where ����id = ����id_In;
      v_��������  := '';
      v_ȷ������  := '';
      v_�������1 := Substr(Emrcontent_In, 1, Instr(Emrcontent_In, '|') - 1);
      v_�������2 := '';
      v_���ע  := Substr(Emrcontent_In, Instr(Emrcontent_In, '|') + 1,Instr(Emrcontent_In, '|',1,2)-1-Instr(Emrcontent_In, '|'));
      v_��������  := '1 ���α���';
      v_����ҽ��  := Substr(Emrcontent_In, Instr(Emrcontent_In, '|',1,2) + 1);
    End If;
  End If;

  --��������
  Update �����걨��¼
  Set ����״̬ = Decode(Incept_In, 1, 1, -1), �վ��� = v_�վ���, �վ�ʱ�� = Sysdate, �վ�˵�� = ˵��_In, ���� = v_����, �Ա� = v_�Ա�, ���� = v_����,
      ְҵ = v_ְҵ, ��ͥ��ַ = v_��ͥ��ַ, ��ͥ�绰 = v_��ͥ�绰, �������� = v_��������, ȷ������ = v_ȷ������, �������1 = v_�������1, �������2 = v_�������1,
      ���ע = v_���ע, ����ҽ�� = v_����ҽ��,�������� = v_��������,����id= ����id_In,��ҳID = ��ҳID_In,������Դ = ������Դ_In
  Where �ļ�id = �ļ�id_In;
  If Sql%RowCount = 0 Then
    Insert Into �����걨��¼
      (�ļ�id, ����״̬, �վ���, �վ�ʱ��, �վ�˵��, ����, �Ա�, ����, ְҵ, ��ͥ��ַ, ��ͥ�绰, ��������, ȷ������, �������1, �������2, ���ע, �ĵ�id, ����ҽ��, ��������,����id,��ҳID,������Դ)
    Values
      (�ļ�id_In, Decode(Incept_In, 1, 1, -1), v_�վ���, Sysdate, ˵��_In, v_����, v_�Ա�, v_����, v_ְҵ, v_��ͥ��ַ, v_��ͥ�绰, v_��������,
       v_ȷ������, v_�������1, v_�������2, v_���ע, �ĵ�id_In, v_����ҽ��, v_��������,����id_In,��ҳID_In,������Դ_In);
  End If;
 
Exception
  When No_Data_Found Then
    Raise_Application_Error(-20101, '[ZLSOFT]�û���ݲ���ȷ��[ZLSOFT]');
  When e_Changed Then
    Raise_Application_Error(-20101, '[ZLSOFT]���������Ѿ��������û��ı䣡[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����걨��¼_Incept;
/

--92258:����,2015-12-31,����ͼ�������ʾ����ͼ��
--92392:����,2016-01-05,������ͼ����е�sql��ѯ���д����У�ͬһ����
--Ӱ�񱨸�������(---���岿��---)***************************************************
CREATE OR REPLACE Package b_PACS_RptPluginOriginal Is
  Type t_Refcur Is Ref Cursor;

  -- ��    �ܣ���ȡ��ʷ�����¼
  Procedure p_GetReportHistory(
    Val                   Out t_Refcur,
    ҽ��id_In             In ����ҽ����¼.ID%Type,
    ��Աid_In             In ������Ա.��Աid%Type,
    ��ǰ����id_In         In ������Ա.����ID%Type,
    �鿴��������ʷ����_In In number := 0
  );

  --��    �ܣ���ȡ��Ӧ��������
  Procedure p_GetReportContent(
    Val           Out t_Refcur,
    ����ID_In     In varchar2,
    EditorType_In Number := 0 --0:PACS����༭����1--���Ӳ����༭����2--�����ĵ��༭��
    );

  --��    �ܣ�����ҽ��ID��ȡ�����Ϣ
  Procedure p_GetStudyInfoByAdviceId(
    Val       Out t_Refcur,
    ҽ��id_In In Ӱ�����¼.ҽ��id%Type
  );

  --��    �ܣ���ȡ����ͼ������
  Procedure p_GetReportImageCount(
    Val Out t_Refcur,
    ��ѯ����_In In varchar2
  );

  --��    �ܣ���ȡ����ͼ������
  Procedure p_GetReportImageData(
    Val         Out t_Refcur,
    ��ѯ����_In In varchar2,
    ��ʼλ��_In In number,
    ����λ��_In In number
  );

  --��    �ܣ���ȡԤ��ͼ������
  Procedure p_GetStudyImageCount(
    Val Out t_Refcur,
    ��ѯ����_In In varchar2,
    �Ƿ���ʱ_In In number:=0
  );

  --��    �ܣ���ȡԤ��ͼ������
  Procedure p_GetStudyImageData(
    Val         Out t_Refcur,
    ��ѯ��ʽ_In In varchar2,
    ��ѯ����_In In varchar2,
    ��ʼλ��_In In number,
    ����λ��_In In number,
    �Ƿ���ʱ_In In number
  );

  --���ܣ���ȡ��ʱͼ������
  Procedure p_Get_TempImageSeries(
    Val         Out t_Refcur,
    ʱ�䷶Χ_In In Number,
    ����_In In Ӱ����ʱ��¼.����%Type:=null
  );

  --����;��ȡͼ��ע
  procedure P_Get_NormalNote(
    Val         Out t_Refcur
  );

  --���ܣ����볣��ͼ��ע
  Procedure p_Insert_Normalnote(
    note_in in Ӱ���ֵ�����.����%Type,
    code_In Ӱ���ֵ�����.����%Type
  );

  --���ܣ��޸ĳ���ͼ��ע
  Procedure p_Edit_Normalnote(
    note_in In Ӱ���ֵ�����.����%Type,
    num_In  Ӱ���ֵ�����.���%Type
  );

  --���ܣ�ɾ������ͼ��ע
  Procedure p_Del_Normalnote(
    num_In Ӱ���ֵ�����.���%Type
  );

  --���ܣ���ȡ��ע����һ������
  Procedure p_Get_NormalNum(
    Val Out t_Refcur
  );
  --���ܣ���ȡ���ID
  Procedure p_Get_PlugID(
    Val     Out t_Refcur,
    ����_In In Ӱ�񱨸���.����%Type
  );

  --���ܣ�����༭���������
  Procedure p_SetFontParam(
    font_In nvarchar2,
    user_In nvarchar2
  );

  --���ܣ���ȡ�༭���������
  Procedure p_GetFontParam(
    Val Out t_Refcur,
    user_In nvarchar2
  );

  --���ܣ�����༭���������
  Procedure p_SetFormParam(
    form_In nvarchar2,
    user_In nvarchar2
  );

  --���ܣ���ȡ�༭���������
  Procedure p_GetFormParam(
    Val Out t_Refcur,
    user_In nvarchar2
  );
  
  --���ܣ�����ͼ��UID��ȡ�����Ϣ
  Procedure p_GetStudyInfoByImageUID(
    Val Out t_Refcur,
    ҽ��ID_In In Ӱ�����¼.ҽ��ID%Type,
    ͼ��UID_In In Ӱ����ͼ��.ͼ��UID%Type
  );
  
  --���ܣ����ݼ��UID��ȡFTP��Ϣ
  Procedure p_GetFtpinfoByStudyUID(
    Val Out t_Refcur,
    ���UID_In In Ӱ�����¼.���UID%Type
  );
  
  --���ܣ����ݿ���ID��ȡFTP��Ϣ
  Procedure p_GetFtpinfoByDeptId(
    Val Out t_Refcur,
    ����ID_In In Ӱ�����̲���.����ID%Type
  );
  
  --���ܣ�����ҽ��ID��ȡFTP��Ϣ
  Procedure p_GetFtpinfoByAdvicetId(
    Val Out t_Refcur,
    ҽ��ID_In In Ӱ�����¼.ҽ��ID%Type
  );
  
  --���ܣ���ȡ���UID
  Procedure p_GetStudyUID(
    Val Out t_Refcur,
    ���UID_In In Ӱ�����¼.���UID%Type
  );
  
  --���ܣ���ȡ����UID
  Procedure p_GetSeriesUID(
    Val Out t_Refcur,
    ����UID_In In Ӱ��������.����UID%Type
  );
  
  --���ܣ������豸�Ż�ȡ�豸��Ϣ
  Procedure p_GetDeviceInfo(
    Val Out t_Refcur,
    �豸��_In In Ӱ���豸Ŀ¼.�豸��%Type
  );
  
  --��ȡҽ��վ�洢�豸��
  Procedure p_GetDeviceIdByAdviceId(
    Val Out t_Refcur,
    ҽ��ID_In In ����ҽ������.ҽ��ID%Type
  );
End b_PACS_RptPluginOriginal;
/

--Ӱ�񱨸淶�Ĺ���(---ʵ�ֲ���---)***************************************************
CREATE OR REPLACE Package Body b_PACS_RptPluginOriginal Is

  --��    �ܣ���ȡ��ʷ�����¼
  Procedure p_GetReportHistory(
    Val                   Out t_Refcur,
    ҽ��id_In             In ����ҽ����¼.ID%Type,
    ��Աid_In             In ������Ա.��Աid%Type,
    ��ǰ����id_In         In ������Ա.����ID%Type,
    �鿴��������ʷ����_In In number := 0
  ) Is
    strSql     varchar2(4000);
    strSqlBack varchar2(4000);
    strFilter  varchar2(400);
  Begin
    If �鿴��������ʷ����_In = 1 Then
      strFilter := ' ';
    Else
      strFilter := ' And c.ִ�п���id+0 in (select ����id from ������Ա where ��Աid = '|| ��Աid_In ||
                   ' union all select to_Number(' || ��ǰ����id_In || ') from dual) ';
    End If;

    strSql := 'Select 2 as ��������, f.����'||'||''-''||'||'f.���� As ��������, c.Id As ҽ��id, a.Ӱ����� as ���,b.������ as ������,' ||
              'to_char(b.����ʱ��,''yyyy-mm-dd hh24:mi:ss'') as ����ʱ��,b.�ĵ����� ��������, c.ҽ������, TO_CHAR(RAWTOHEX(b.id)) ����ID ' ||
              'From Ӱ�����¼ A, Ӱ�񱨸��¼ B, ����ҽ����¼ C, Ӱ�����¼ D, ����ҽ����¼ E, ���ű� F ' ||
              'Where a.ҽ��id = b.ҽ��id And d.ҽ��id = e.Id And e.Id =' ||
              ҽ��id_In || ' And e.ִ�п���ID = F.ID And b.ҽ��id = c.Id And ' ||
              '(c.����id = e.����id Or a.����id = d.����id) And c.���id Is Null ' || strFilter ||
              ' union all ' ||
              'Select 1 as ��������, g.����'||'||''-''||'||'g.���� As ��������, c.Id As ҽ��id, a.Ӱ����� as ���, a.������, ' ||
              'to_char(f.����ʱ��,''yyyy-mm-dd hh24:mi:ss'') as ����ʱ��, a.Ӱ�����||''����'' ��������, c.ҽ������,TO_CHAR( b.����id) as ����ID ' ||
              'From Ӱ�����¼ A, ����ҽ������ B, ����ҽ����¼ C, Ӱ�����¼ D, ����ҽ����¼ E, ���Ӳ�����¼ F, ���ű� G ' ||
              'Where a.ҽ��id = b.ҽ��id And d.ҽ��id = e.Id And e.Id = ' ||
              ҽ��id_In || ' And e.ִ�п���ID = g.ID And b.ҽ��id = c.Id And b.����ID Is Not Null And ' ||
              '(c.����id = e.����id Or a.����id = d.����id) And c.���id Is Null And b.����id = f.id ' || strFilter;

    strSqlBack := strSql;
    strSqlBack := replace(strSqlBack, 'Ӱ�����¼', 'HӰ�����¼');
    strSqlBack := replace(strSqlBack, '����ҽ������', 'H����ҽ������');
    strSqlBack := replace(strSqlBack, '����ҽ����¼', 'H����ҽ����¼');

    strSql := strSql || ' UNION ALL ' || strSQLBack || ' Order By ����ʱ�� Asc';

    Open Val For strSql;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetReportHistory;

  --��    �ܣ���ȡ��Ӧ��������
  Procedure p_GetReportContent(
    Val           Out t_Refcur,
    ����ID_In     varchar2,
    EditorType_In Number := 0 --0:���Ӳ����༭����1--PACS����༭����2--�����ĵ��༭��
    ) Is
    strSql varchar2(1000);
  Begin
    If EditorType_In = 1 Then
      strSql := 'Select a.�����ı� As ����, b.��������, b.�����ı� As ����,b.��ʼ�� as �汾 From ���Ӳ������� a,���Ӳ������� b ' ||
                'Where a.�ļ�id = ' || ����ID_In ||
                ' And a.�������� = 3 And a.Id = b.��ID And b.�������� = 2 and b.��ֹ��=0 ';
    ElsIf EditorType_In = 0 Then
      strSql := 'select ���� from ���Ӳ�����ʽ where �ļ�ID=' || ����ID_In;
    Else
      strSql := 'Select �������� As ���� From Ӱ�񱨸��¼ Where ID=HexToRaw(''' ||
                ����ID_In || ''')';
    End If;

    Open Val For strSql;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetReportContent;

  --��    �ܣ�����ҽ��ID��ȡ�����Ϣ
  Procedure p_GetStudyInfoByAdviceId(
    Val       Out t_Refcur,
    ҽ��id_In In Ӱ�����¼.ҽ��id%Type
  ) Is
    strSql varchar2(100);
  Begin
    strSql := 'Select ���UID,����ͼ��,��������,����,����,�Ա�,���� from Ӱ�����¼ where ҽ��ID =' || ҽ��id_In;
    Open Val For strSql;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetStudyInfoByAdviceId;

  --��    �ܣ���ȡ����ͼ������
  Procedure p_GetReportImageCount(
    Val Out t_Refcur,
    ��ѯ����_In In varchar2
  ) Is
  Begin
    Open Val For
      Select Count(B.Column_Value) ����ֵ
      From Ӱ�����¼ A, Table(Cast(f_Str2list(Replace(A.����ͼ��,';',',')) As zlTools.t_Strlist)) B Where ҽ��ID = ��ѯ����_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetReportImageCount;

  --��    �ܣ���ȡ����ͼ������
  Procedure p_GetReportImageData(
    Val         Out t_Refcur,
    ��ѯ����_In In varchar2,
    ��ʼλ��_In In number,
    ����λ��_In In number
  ) Is
  Begin
    Open Val For
         Select * from (Select rownum as ˳���, rownum as ͼ���, B.FTP�û��� As User1,B.FTP���� As Pwd1,B.IP��ַ As Host1,'/'||B.FtpĿ¼||'/' As Root1,
          Decode(A.��������,Null,'',to_Char(A.��������,'YYYYMMDD')||'/')||A.���UID||'/'||Replace(D.Column_Value,'.jpg','') As URL,B.�豸�� as �豸��1,
          C.FTP�û��� As User2,C.FTP���� As Pwd2,C.IP��ַ As Host2,'/'||C.FtpĿ¼||'/' As Root2,
          C.�豸�� as �豸��2,Replace(D.Column_Value,'.jpg','') AS ͼ��UID,A.���UID,'' ����UID,0 ��̬ͼ,'' ��������,'' �ɼ�ʱ��, '' ¼�Ƴ���
          From Ӱ�����¼ A, Ӱ���豸Ŀ¼ B, Ӱ���豸Ŀ¼ C, Table(Cast(f_Str2list(Replace(A.����ͼ��,';',',')) As zlTools.t_Strlist)) D
          Where A.λ��һ = B.�豸��(+) And A.λ�ö� = C.�豸��(+) And A.ҽ��id = ��ѯ����_In)
          Where ˳��� >= ��ʼλ��_In and ˳���<=����λ��_In;

  End p_GetReportImageData;

  --��    �ܣ���ȡԤ��ͼ������
  Procedure p_GetStudyImageCount(
    Val Out t_Refcur,
    ��ѯ����_In In varchar2,
    �Ƿ���ʱ_In In number:=0
  ) Is
    strSql varchar2(2000);
  Begin
    if �Ƿ���ʱ_In = 0 then
      strSql := 'select T1.����ֵ+T2.����ֵ as ����ֵ from ' ||
              '(select count(1) as ����ֵ from Ӱ����ͼ�� a, Ӱ�������� b, Ӱ�����¼ c ' ||
              'where a.����UID=b.����UID and b.���UID=c.���UID and c.ҽ��ID=''' ||
              ��ѯ����_In || ''') T1,' ||
              '(select count(1) as ����ֵ from HӰ����ͼ�� a, HӰ�������� b, Ӱ�����¼ c ' ||
              'where a.����UID=b.����UID and b.���UID=c.���UID and c.ҽ��ID=''' ||
              ��ѯ����_In || ''') T2';
    else
      strSql := 'select count(1)  as ����ֵ from Ӱ����ʱͼ��  where  ����UID='''||��ѯ����_In || '''';
    end if;

    Open Val For strSql;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetStudyImageCount;

  --��    �ܣ���ȡԤ��ͼ������
  Procedure p_GetStudyImageData(
    Val         Out t_Refcur,
    ��ѯ��ʽ_In In varchar2,
    ��ѯ����_In In varchar2,
    ��ʼλ��_In In number,
    ����λ��_In In number,
    �Ƿ���ʱ_In In number
  ) Is
    strSql    varchar2(2000);
    strFilter varchar2(100);
  Begin
    if ��ѯ��ʽ_In = 0 then
      strFilter := 'and c.ҽ��ID=''' || ��ѯ����_In || '''';
    elsif ��ѯ��ʽ_In = 1 then
      strFilter := 'and B.����UID=''' || ��ѯ����_In || '''';
    else
      strFilter := 'and A.ͼ��UID=''' || ��ѯ����_In || '''';
    end if;

    strSql := 'Select * from (Select rownum as ˳���, T.* from(' ||
              'Select A.ͼ���,D.FTP�û��� As User1,D.FTP���� As Pwd1,D.IP��ַ As Host1,''/''||D.FtpĿ¼||''/'' As Root1,' ||
              'Decode(C.��������,Null,'''',to_Char(C.��������,''YYYYMMDD'')||''/'')||C.���UID||''/''||A.ͼ��UID As URL,d.�豸�� as �豸��1,' ||
              'E.FTP�û��� As User2,E.FTP���� As Pwd2,E.IP��ַ As Host2,''/''||E.FtpĿ¼||''/'' As Root2,' ||
              'e.�豸�� as �豸��2, A.ͼ��UID,C.���UID,B.����UID,A.��̬ͼ,A.��������,A.�ɼ�ʱ��, A.¼�Ƴ��� ' ||
              'From Ӱ����ͼ�� A,Ӱ�������� B,Ӱ�����¼ C,Ӱ���豸Ŀ¼ D,Ӱ���豸Ŀ¼ E ' ||
              'Where A.����UID=B.����UID And B.���UID=C.���UID And C.λ��һ=D.�豸��(+) And C.λ�ö�=E.�豸��(+) ' ||
              strFilter || ' '|| 'Order by ����UID, ͼ���) T ) ' ||
              'Where ˳���>=' || ��ʼλ��_In || ' and ˳���<=' || ����λ��_In || '';

    if �Ƿ���ʱ_In = 1 then
      strSql:= replace(strSql,'Ӱ����','Ӱ����ʱ');
    end if;

    Open Val For strSql;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetStudyImageData;

  --���ܣ���ȡ��ʱͼ������
  Procedure p_Get_TempImageSeries(
    Val Out t_Refcur,
    ʱ�䷶Χ_In In Number,
    ����_In In Ӱ����ʱ��¼.����%Type:=null
  ) As
  Begin
    If ����_In Is Null Then
      Open Val For
        select B.����UID,A.����,A.���� As ���, A.�������� from Ӱ����ʱ��¼ A,Ӱ����ʱ���� B
        where A.���uid = B.���uid And A.�������� Between Sysdate-ʱ�䷶Χ_In And Sysdate
        order by ���;
    Else
      Open Val For
        select B.����UID,A.����,A.���� As ���, A.�������� from Ӱ����ʱ��¼ A,Ӱ����ʱ���� B
        where A.���uid = B.���uid And A.�������� Between Sysdate-ʱ�䷶Χ_In And Sysdate and a.���� = ����_In
        order by ���;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --���ܣ���ȡͼ��ע
  Procedure p_Get_Normalnote(
    Val Out t_Refcur
  ) As
  Begin
    Open Val For
      Select b.��� As ���, b.���� As ����
        From Ӱ���ֵ��嵥 A, Ӱ���ֵ����� B
       Where a.Id = b.�ֵ�id
         And a.���� = 'Ӱ��ͼ��ע';
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End;

  --���ܣ����볣��ͼ��ע
  Procedure p_Insert_Normalnote(
    note_in In Ӱ���ֵ�����.����%Type,
    code_In Ӱ���ֵ�����.����%Type
  ) As
    n_Num         Number;
    dictionary_id Varchar2(36);
  Begin
    Select id
      Into dictionary_id
      From Ӱ���ֵ��嵥
     Where ˵�� = 'Ӱ��ͼ��ע';
    Select Decode(Max(to_number(���)), Null, 0, Max(to_number(���)))
      Into n_Num
      From Ӱ���ֵ�����
     Where �ֵ�id = dictionary_id;
    n_Num := n_Num + 1;
    Insert Into Ӱ���ֵ�����
      (�ֵ�id, ���, ����, ˵��)
    Values
      (dictionary_id, to_char(n_Num), note_in, 'Ӱ��ͼ��ע');
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Insert_Normalnote;

  --���ܣ��޸ĳ���ͼ��ע
  Procedure p_Edit_Normalnote(
    note_in In Ӱ���ֵ�����.����%Type,
    num_In  Ӱ���ֵ�����.���%Type
  ) As
    dictionary_id Varchar2(36);
  Begin
    Select id
      Into dictionary_id
      From Ӱ���ֵ��嵥
     Where ˵�� = 'Ӱ��ͼ��ע';
    Update Ӱ���ֵ����� t
       Set t.���� = note_in
     Where t.�ֵ�id = dictionary_id
       And t.��� = num_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Edit_Normalnote;

  --���ܣ�ɾ������ͼ��ע
  Procedure p_Del_Normalnote(
    num_In Ӱ���ֵ�����.���%Type
  ) As
    dictionary_id Varchar2(36);
  Begin
    Select id
      Into dictionary_id
      From Ӱ���ֵ��嵥
     Where ˵�� = 'Ӱ��ͼ��ע';
    Delete Ӱ���ֵ����� t
     Where t.�ֵ�id = dictionary_id
       And t.��� = num_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Del_Normalnote;

  --���ܣ���ȡ��ע����һ������
  Procedure p_Get_NormalNum(
    Val Out t_Refcur
  ) As
    n_Num         Number;
    dictionary_id Varchar2(36);
  Begin
    Select id
      Into dictionary_id
      From Ӱ���ֵ��嵥
     Where ˵�� = 'Ӱ��ͼ��ע';
    Open Val For
      Select Decode(Max(to_number(���)), Null, 1, Max(to_number(���) + 1)) ���
        From Ӱ���ֵ����� t
       Where t.�ֵ�id = dictionary_id;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_NormalNum;

  --���ܣ���ȡ���ID
  Procedure p_Get_PlugID(
    Val     Out t_Refcur,
    ����_In In Ӱ�񱨸���.����%Type
  ) Is
  Begin
    Open Val For
      Select RawToHex(ID) ID From Ӱ�񱨸��� Where ���� = ����_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_PlugID;

  --���ܣ�����༭���������
  Procedure p_SetFontParam(
    font_In nvarchar2,
    user_In nvarchar2
  ) As
    m_ID     nvarchar2(36);
    numcount int;
  Begin
    Select RawToHex(ID)
      Into m_ID
      From Ӱ�����˵��
     Where ģ�� = 'ImageEditor'
       And ������ = '��������';
    Select Count(*)
      Into numcount
      From Ӱ�����ȡֵ t
     Where t.����id = m_ID
       And t.������ʶ = user_In;
    If numcount > 0 then
      Update Ӱ�����ȡֵ a
         Set a.����ֵ = font_In
       Where a.������ʶ = user_In
         And a.����id = m_ID;
    Else
      Insert Into Ӱ�����ȡֵ a
        (ID, ����ID, ������ʶ, ����ֵ)
      Values
        (sys_Guid(), m_ID, user_In, font_In);
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_SetFontParam;

  --���ܣ���ȡ�༭���������
  Procedure p_GetFontParam(
    Val Out t_Refcur,
    user_In nvarchar2
  ) As
    m_ID nvarchar2(36);
  Begin
    Select RawToHex(ID)
      Into m_ID
      From Ӱ�����˵��
     Where ģ�� = 'ImageEditor'
       And ������ = '��������';
    Open Val For
      Select a.����ֵ
        From Ӱ�����ȡֵ a
       Where a.����id = m_ID
         And a.������ʶ = user_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetFontParam;

  --���ܣ�����༭���������
  Procedure p_SetFormParam(
    form_In nvarchar2,
    user_In nvarchar2
  ) As
    m_ID     nvarchar2(36);
    numcount int;
  Begin
    Select RawToHex(ID)
      Into m_ID
      From Ӱ�����˵��
     Where ģ�� = 'ImageEditor'
       And ������ = '��������';
    Select Count(*)
      Into numcount
      From Ӱ�����ȡֵ t
     Where t.����id = m_ID
       And t.������ʶ = user_In;
    If numcount > 0 then
      Update Ӱ�����ȡֵ a
         Set a.����ֵ = form_In
       Where a.������ʶ = user_In
         And a.����id = m_ID;
    Else
      Insert Into Ӱ�����ȡֵ a
        (ID, ����ID, ������ʶ, ����ֵ)
      Values
        (sys_Guid(), m_ID, user_In, form_In);
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_SetFormParam;

  --���ܣ���ȡ�༭���������
  Procedure p_GetFormParam(
    Val Out t_Refcur,
    user_In nvarchar2
  ) As
    m_ID nvarchar2(36);
  Begin
    Select RawToHex(ID)
      Into m_ID
      From Ӱ�����˵��
     Where ģ�� = 'ImageEditor'
       And ������ = '��������';
    Open Val For
      Select a.����ֵ
        From Ӱ�����ȡֵ a
       Where a.����id = m_ID
         And a.������ʶ = user_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetFormParam;
  
  --���ܣ�����ͼ��UID��ȡ�����Ϣ
  Procedure p_GetStudyInfoByImageUID(
    Val Out t_Refcur,
    ҽ��ID_In In Ӱ�����¼.ҽ��ID%Type,
    ͼ��UID_In In Ӱ����ͼ��.ͼ��UID%Type
  )As
  Begin
    Open Val For
      Select D.���UID From Ӱ����ͼ�� A,Ӱ�������� B,Ӱ�����¼ C,Ӱ����ʱ���� D
      Where C.ҽ��ID=ҽ��ID_In And A.ͼ��UID=ͼ��UID_In And A.����UID=B.����UID And B.���UID=C.���UID And A.����UID = D.����UID;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetStudyInfoByImageUID;
  
  --���ܣ����ݼ��UID��ȡFTP��Ϣ
  Procedure p_GetFtpinfoByStudyUID(
    Val Out t_Refcur,
    ���UID_In In Ӱ�����¼.���UID%Type
  )As
  Begin
    Open Val For
      Select D.FTP�û��� As FtpUser,D.FTP���� As FtpPwd,C.λ��һ,C.λ�ö�,C.λ����,C.��������,
      D.IP��ַ As Host,'/'||D.FtpĿ¼||'/' As Root,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')||C.���UID As URL
      From Ӱ�����¼ C,Ӱ���豸Ŀ¼ D Where Decode(C.λ��һ,Null,C.λ�ö�,C.λ��һ)=D.�豸��(+) And C.���UID= ���UID_In Union All
      Select D.FTP�û��� As FtpUser,D.FTP���� As FtpPwd,C.λ��һ,C.λ�ö�,C.λ����,C.��������,
      D.IP��ַ As Host,'/'||D.FtpĿ¼||'/' As Root,Decode(C.��������,Null,'',to_Char(C.��������,'YYYYMMDD')||'/')||C.���UID As URL
      From Ӱ����ʱ��¼ C,Ӱ���豸Ŀ¼ D Where Decode(C.λ��һ,Null,C.λ�ö�,C.λ��һ)=D.�豸��(+) And C.���UID= ���UID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetFtpinfoByStudyUID;
  
  --���ܣ����ݿ���ID��ȡFTP��Ϣ
  Procedure p_GetFtpinfoByDeptId(
    Val Out t_Refcur,
    ����ID_In In Ӱ�����̲���.����ID%Type
  )As
  Begin
    Open Val For
      Select a.�豸��, a.ip��ַ, a.ftp�û���, a.ftp���� From Ӱ���豸Ŀ¼ a, Ӱ�����̲��� b
      Where a.�豸�� = b.����ֵ And b.������ = '�洢�豸��' And b.����id=����ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetFtpinfoByDeptId;
  
  --���ܣ�����ҽ��ID��ȡFTP��Ϣ
  Procedure p_GetFtpinfoByAdvicetId(
    Val Out t_Refcur,
    ҽ��ID_In In Ӱ�����¼.ҽ��ID%Type
  )As
  Begin
    Open Val For
      Select a.�豸��, a.ip��ַ, a.ftp�û���, a.ftp���� From Ӱ���豸Ŀ¼ a, Ӱ�����¼ b 
      Where b.λ��һ = a.�豸��(+) And b.ҽ��id =ҽ��ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetFtpinfoByAdvicetId;
  
  --���ܣ���ȡ���UID
  Procedure p_GetStudyUID(
    Val Out t_Refcur,
    ���UID_In In Ӱ�����¼.���UID%Type
  )As
  Begin
    Open Val For
      Select ���UID from Ӱ�����¼ where ���UID = ���UID_In Union All Select ���UID from Ӱ����ʱ��¼ where ���UID = ���UID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetStudyUID;
  
  --���ܣ���ȡ����UID
  Procedure p_GetSeriesUID(
    Val Out t_Refcur,
    ����UID_In In Ӱ��������.����UID%Type
  )As
  Begin
    Open Val For
      Select ����UID from Ӱ�������� where ����UID = ����UID_In Union All Select ����UID from Ӱ����ʱ���� where ����UID = ����UID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetSeriesUID;
  
  --���ܣ������豸�Ż�ȡ�豸��Ϣ
  Procedure p_GetDeviceInfo(
    Val Out t_Refcur,
    �豸��_In In Ӱ���豸Ŀ¼.�豸��%Type
  )As
  Begin
    Open Val For
      Select �豸��,�豸��,'/'||Decode(FtpĿ¼,Null,'',FtpĿ¼||'/') As URL,FTP�û���,FTP����,IP��ַ
      From Ӱ���豸Ŀ¼ Where ����=1 and �豸��=�豸��_In and NVL(״̬,0)=1;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetDeviceInfo;
  
  --��ȡҽ��վ�洢�豸��
  Procedure p_GetDeviceIdByAdviceId(
    Val Out t_Refcur,
    ҽ��ID_In In ����ҽ������.ҽ��ID%Type
  )As
  Begin
    Open Val For
      Select d.����ֵ From ҽ��ִ�з��� a, ����ҽ������ b, Ӱ��DICOM����� c, Ӱ��DICOM������� d
      Where a.����ID = b.ִ�в���id And a.ִ�м� = b.ִ�м� And a.����豸 = c.�豸��
      And c.������='ͼ�����' And c.����ID=d.����ID And d.��������='�洢�豸' And b.ҽ��id=ҽ��ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetDeviceIdByAdviceId;
End b_PACS_RptPluginOriginal;
/


--89676:����,2015-12-30,��¼��ͬ��������Ŀ�����µ�
--91458:������,2016-01-04,�������봦��
Create Or Replace Procedure Zl_���˻�������_Update
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
  δ��˵��_In In ���˻�����ϸ.δ��˵��%Type := Null --��������洢ҽ��ID:���ͺ�
) Is
  Intins      Number(18);
  Int����     Number(1);
  n_Newid     ���˻�������.Id%Type;
  n_Oldid     ���˻�������.Id%Type;
  n_����      ���˻����ӡ.����%Type;
  n_Mutilbill Number(1);
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
    From �����ļ��б� a, ���˻����ļ� b, ���˻����ļ� c, ���˻������� d
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
  If ����Ա_In Is Null Then
    v_������ := Zl_Username;
  Else
    v_������ := ����Ա_In;
  End If;

  --����Ƕ�Ӧ��ݻ����ļ�ֵΪ1����ʾ��ͬ�����������ļ������򲻴����ļ�ͬ�� 
  n_Mutilbill := Zl_To_Number(Zl_Getsysparameter('��Ӧ��ݻ����ļ�', 1255));

  Begin
    Select Id, �������
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
  Where Id = �ļ�id_In;
  d_Ӥ����Ժʱ�� := Null;
  If n_Ӥ�� <> 0 Then
    Begin
      Select ��ʼִ��ʱ��
      Into d_Ӥ����Ժʱ��
      From ����ҽ����¼ b, ������ĿĿ¼ c
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
      From ���˱䶯��¼ a, ���˻����ļ� b
      Where a.����id Is Not Null And a.����id = b.����id And a.��ҳid = b.��ҳid And b.Id = �ļ�id_In And
            (To_Date(To_Char(����ʱ��_In, 'YYYY-MM-DD HH24:MI') || '59', 'YYYY-MM-DD HH24:MI:SS') >= a.��ʼʱ�� And
            (To_Date(To_Char(����ʱ��_In, 'YYYY-MM-DD HH24:MI') || '00', 'YYYY-MM-DD HH24:MI:SS') < =
            Nvl(a.��ֹʱ��, Sysdate) Or a.��ֹʱ�� Is Null)) And Rownum < 2;
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
  From ���˻�����ϸ a, ���˻������� b
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
      Select Id
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
    Select Count(Id)
    Into Intins
    From ���˻�����ϸ
    Where ��¼id = n_��¼id And Mod(��¼����, 10) <> 5 And ��ֹ�汾 Is Null And Id <> n_��ϸid;
    If Intins = 0 Then
      Delete From ���˻�����ϸ Where ��¼id = n_��¼id;
    Else
      Delete From ���˻�����ϸ Where Id = n_��ϸid;
    End If;
  
    Delete From ���˻������� a
    Where a.Id = n_��¼id And Not Exists (Select 1 From ���˻�����ϸ b Where b.��¼id = a.Id);
  
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
    If Nvl(n_�������, 0) <> 0 Then
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
          From �����ļ��б� a, ���˻����ļ� b, ���˻������� c
          Where a.Id = b.��ʽid And b.Id = c.�ļ�id And c.Id = Rsdel.��¼id;
        Exception
          When Others Then
            n_�ļ�id := 0;
        End;
        Delete ���˻������� Where Id = Rsdel.��¼id;
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
             From �����ļ��ṹ a, �����¼��Ŀ b
             Where a.Ҫ������ = b.��Ŀ���� And b.��Ŀ��� = ��Ŀ���_In And
                   ��id = (Select b.Id
                          From ���˻����ļ� a, �����ļ��ṹ b
                          Where a.Id = �ļ�id_In And a.��ʽid = b.�ļ�id And b.��id Is Null And b.������� = 4)
             Union
             Select ��Ŀ���
             From �����¼��Ŀ
             Where ��Ŀ���� = 2 And ��Ŀ��� = ��Ŀ���_In);
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
        (Id, �ļ�id, ����ʱ��, ���汾, ������, ����ʱ��)
      Values
        (n_��¼id, �ļ�id_In, ����ʱ��_In, n_��߰汾, v_������, Sysdate);
    End If;
  
    --���뱾�εǼǵĲ��˻�����ϸ 
    Update ���˻�����ϸ
    Set ��¼���� = ��¼����_In, ������Դ = ������Դ_In, δ��˵�� = δ��˵��_In, ��¼�� = v_������, ��¼ʱ�� = Sysdate
    Where ��¼id = n_��¼id And ��Ŀ��� = ��Ŀ���_In And Nvl(���²�λ, 'TWBW') = Nvl(���²�λ_In, 'TWBW') And
          Nvl(��¼���, 0) = Nvl(��¼���_In, 0) And ��ʼ�汾 = n_��߰汾 And ��ֹ�汾 Is Null;
    If Sql%Rowcount = 0 Then
      Select ���˻�����ϸ_Id.Nextval Into n_��ϸid From Dual;
      Insert Into ���˻�����ϸ
        (Id, ��¼id, ��¼����, ��Ŀ����, ��Ŀid, ������, ��Ŀ���, ��Ŀ����, ��Ŀ����, ��¼����, ��Ŀ��λ, ��¼���, ��¼���, ���²�λ, ������Դ, ����, δ��˵��, ��ʼ�汾, ��ֹ�汾,
         ��¼��, ��¼ʱ��)
        Select n_��ϸid, n_��¼id, ��¼����_In, a.������, a.��Ŀid, ������_In, a.��Ŀ���, Upper(a.��Ŀ����), a.��Ŀ����, ��¼����_In, a.��Ŀ��λ, 0,
               ��¼���_In, ���²�λ_In, ������Դ_In, Nvl(b.����, 0), δ��˵��_In, n_��߰汾, Null, v_������, Sysdate
        From �����¼��Ŀ a, ���˻�����ϸ b
        Where a.��Ŀ��� = b.��Ŀ���(+) And b.��ֹ�汾(+) Is Null And b.��¼id(+) = n_��¼id And a.��Ŀ��� = ��Ŀ���_In And Rownum < 2;
    End If;
    Select Id
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
      Update ���˻������� Set ������ = v_������, ����ʱ�� = Sysdate Where Id = n_��¼id;
    End If;
  
    If Nvl(n_�������, 0) <> 0 Then
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
                   From ���¼�¼��Ŀ f, �����¼��Ŀ g
                   Where f.��Ŀ��� = g.��Ŀ��� And g.��Ŀ���� = 2 And
                         (g.���ÿ��� = 1 Or
                         (g.���ÿ��� = 2 And Exists
                          (Select 1 From �������ÿ��� d Where g.��Ŀ��� = d.��Ŀ��� And d.����id = v_����id))) And Nvl(g.Ӧ�÷�ʽ, 0) <> 0 And
                         (Nvl(g.���ò���, 0) = 0 Or Nvl(g.���ò���, 0) = Decode(Nvl(Row_Format.Ӥ��, 0), 0, 1, 2))
                   Union All
                   Select b.�����ı� As ���, 1 As ��Ŀ����
                   From �����ļ��ṹ a, �����ļ��ṹ b
                   Where a.�ļ�id = Row_Format.��ʽid And a.��id Is Null And a.������� In (2, 3) And b.��id = a.Id) h
            Where Instr(',' || h.��� || ',', ',' || ��Ŀ���_In || ',', 1) > 0;
          Exception
            When Others Then
              Intins := 0;
          End;
        Else
          Begin
            Select 1, g.��Ŀ����
            Into Intins, n_��Ŀ����
            From ���¼�¼��Ŀ f, �����¼��Ŀ g
            Where f.��Ŀ��� = g.��Ŀ��� And Nvl(g.Ӧ�÷�ʽ, 0) = 1 And g.����ȼ� >= 0 And
                  (Nvl(g.���ò���, 0) = 0 Or Nvl(g.���ò���, 0) = Decode(Nvl(Row_Format.Ӥ��, 0), 0, 1, 2)) And f.��Ŀ��� = ��Ŀ���_In And
                  (g.���ÿ��� = 1 Or (g.���ÿ��� = 2 And Exists
                   (Select 1 From �������ÿ��� d Where g.��Ŀ��� = d.��Ŀ��� And d.����id = v_����id)));
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
            From ���˻����ļ� a, ���˻������� b
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
              (Id, �ļ�id, ������, ����ʱ��, ����ʱ��, ���汾)
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
                  (Id, ��¼id, ��¼����, ��Ŀ����, ��Ŀid, ��Ŀ���, ��Ŀ����, ��Ŀ����, ��¼����, ��Ŀ��λ, ��¼���, ���²�λ, ������Դ, ��Դid, ��ʼ�汾, ��ֹ�汾, ��¼��,
                   ��¼ʱ��, ��¼���)
                  Select ���˻�����ϸ_Id.Nextval, n_Newid, b.��¼����, b.��Ŀ����, b.��Ŀid, b.��Ŀ���, b.��Ŀ����, b.��Ŀ����, b.��¼����, b.��Ŀ��λ,
                         b.��¼���, b.���²�λ, 1, b.Id, 1, Null, b.��¼��, Sysdate, 1
                  From (Select ��Ŀ���_In As ��Ŀ���, Nvl(���²�λ_In, '��') As ���²�λ
                         From Dual
                         Minus
                         Select f.��Ŀ���, Decode(Nvl(f.��Ŀ����, 1), 2, Nvl(���²�λ, '��'), Nvl(���²�λ_In, '��'))
                         From ���˻�����ϸ e, �����¼��Ŀ f
                         Where e.��¼id = n_Newid And e.��Ŀ��� = f.��Ŀ���) a, ���˻�����ϸ b
                  Where a.��Ŀ��� = b.��Ŀ��� And b.��¼id = n_��¼id And b.Id = n_��ϸid;
                If Sql%Rowcount > 0 Then
                  Int���� := 1;
                End If;
              End If;
            Else
              Update ���˻�����ϸ
              Set ��¼���� = ��¼����_In, ��Դid = n_��ϸid
              Where ��¼id = n_Newid And ��Ŀ��� = ��Ŀ���_In And
                    Decode(n_��Ŀ����, 2, Nvl(���²�λ, '��'), Nvl(���²�λ_In, '��')) = Nvl(���²�λ_In, '��') And ������Դ > 0;
              If Sql%Rowcount > 0 Then
                Int���� := 1;
              End If;
            End If;
          End If;
        End If;
        --2\��ѭ�������¼�� 
      Else
        If n_Mutilbill = 1 Then
          --��ȡ��¼���뵱ǰ��¼�������ص����������ݵĹ̶���Ŀ 
          Select Count(*)
          Into Intins
          From (Select b.��Ŀ���
                 From �����ļ��ṹ a, �����¼��Ŀ b
                 Where a.Ҫ������ = b.��Ŀ���� And b.��Ŀ��ʾ In (0, 4, 5) And
                       ��id =
                       (Select Id From �����ļ��ṹ Where �ļ�id = Row_Format.��ʽid And ��id Is Null And ������� = 4)
                 Intersect
                 Select b.��Ŀ���
                 From �����ļ��ṹ a, �����¼��Ŀ b, ���˻����ļ� c, ���˻������� d, ���˻�����ϸ g
                 Where c.Id = d.�ļ�id And a.�ļ�id = c.��ʽid And d.Id = g.��¼id And d.Id = n_��¼id And g.Id = n_��ϸid And
                       b.��Ŀ��� = g.��Ŀ��� And b.��Ŀ��ʾ In (0, 4, 5) And g.��¼���� = 1 And a.Ҫ������ = b.��Ŀ���� And
                       a.��id = (Select Id From �����ļ��ṹ e Where e.�ļ�id = c.��ʽid And ��id Is Null And ������� = 4));
        
          If Intins > 0 Then
            n_Newid := 0;
            --����ָ���ļ��Ѿ�������ͬ����ʱ������ݣ�ֱ��������ID���� 
            Begin
              Select c.Id
              Into n_Newid
              From ���˻������� c
              Where c.�ļ�id = Row_Format.�ļ�id And c.����ʱ�� = ����ʱ��_In;
            Exception
              When Others Then
                n_Newid := 0;
            End;
          
            If n_Newid = 0 Then
              --������¼������¼ 
              Select ���˻�������_Id.Nextval Into n_Newid From Dual;
            
              Insert Into ���˻�������
                (Id, �ļ�id, ������, ����ʱ��, ����ʱ��, ���汾)
                Select n_Newid, Row_Format.�ļ�id, c.������, c.����ʱ��, c.����ʱ��, 1
                From ���˻������� c
                Where c.Id = n_��¼id;
            End If;
          
            If n_Newid > 0 Then
              --����δͬ���ļ�¼������ 
              Select Count(*) Into v_������Դ From ���˻�����ϸ Where ��¼id = n_Newid And ��Ŀ��� = ��Ŀ���_In;
              If v_������Դ = 0 Then
                Insert Into ���˻�����ϸ
                  (Id, ��¼id, ��¼����, ��Ŀ����, ��Ŀid, ��Ŀ���, ��Ŀ����, ��Ŀ����, ��¼����, ��Ŀ��λ, ��¼���, ���²�λ, ������Դ, ��Դid, δ��˵��, ��ʼ�汾, ��ֹ�汾,
                   ��¼��, ��¼ʱ��)
                  Select ���˻�����ϸ_Id.Nextval, n_Newid, b.��¼����, b.��Ŀ����, b.��Ŀid, b.��Ŀ���, b.��Ŀ����, b.��Ŀ����, b.��¼����, b.��Ŀ��λ,
                         b.��¼���, b.���²�λ, 1, b.Id, b.δ��˵��, 1, Null, b.��¼��, Sysdate
                  From (Select b.��Ŀ���
                         From �����ļ��ṹ a, �����¼��Ŀ b
                         Where a.Ҫ������ = b.��Ŀ���� And b.��Ŀ��ʾ In (0, 4, 5) And
                               ��id = (Select Id
                                      From �����ļ��ṹ
                                      Where �ļ�id = Row_Format.��ʽid And ��id Is Null And ������� = 4)
                         Intersect
                         Select b.��Ŀ���
                         From �����ļ��ṹ a, �����¼��Ŀ b, ���˻����ļ� c, ���˻������� d, ���˻�����ϸ g
                         Where c.Id = d.�ļ�id And a.�ļ�id = c.��ʽid And d.Id = g.��¼id And d.Id = n_��¼id And g.Id = n_��ϸid And
                               b.��Ŀ��� = g.��Ŀ��� And b.��Ŀ��ʾ In (0, 4, 5) And g.��¼���� = 1 And a.Ҫ������ = b.��Ŀ���� And
                               a.��id =
                               (Select Id From �����ļ��ṹ e Where e.�ļ�id = c.��ʽid And ��id Is Null And ������� = 4)) a, ���˻�����ϸ b
                  Where a.��Ŀ��� = b.��Ŀ��� And b.��¼id = n_��¼id And b.Id = n_��ϸid;
                If Sql%Rowcount > 0 Then
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
                Where ��¼id = n_Newid And ��Ŀ��� = ��Ŀ���_In And ������Դ > 0;
                If Sql%Rowcount > 0 Then
                  Int���� := 1;
                End If;
              End If;
            End If;
          End If;
        End If;
      End If;
    End Loop;
  
    If Int���� = 1 Then
      Update ���˻�����ϸ Set ���� = 1 Where Id = n_��ϸid;
      --����ʷ���ݵĹ��ñ�־����ΪNULL 
      Update ���˻�����ϸ Set ���� = Null Where ��¼id = n_��¼id And ��Ŀ��� = ��Ŀ���_In And Id <> n_��ϸid;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_���˻�������_Update;
/

--91458:������,2016-01-04,�������봦��
Create Or Replace Procedure Zl_�����ļ���ʽ_Update
(
  �ļ�id_In     In �����ļ��ṹ.�ļ�id%Type,
  ��ͷ����_In   Number,
  ������_In     Number,
  ��С�и�_In   Number,
  �ı�����_In   Varchar2,
  �ı���ɫ_In   Number,
  �����ɫ_In   Number,
  �����ı�_In   Varchar2,
  ��������_In   Varchar2,
  ��ʼʱ��_In   Number,
  ��ֹʱ��_In   Number,
  ��������_In   Varchar2,
  ������ɫ_In   Number,
  ��Ч������_In Number,
  ���кϲ�_In   Number,
  ʱ������_In   Number, --��¼��Ԥ������ӡʱ����ʱ����(�磺Ѫ�Ǽ�¼������Ҫ��ʾ�����ʱ��)
  ҳ���ʽ_In   ����ҳ���ʽ.��ʽ%Type, --��.PaperKind;.PaperOrient;.PaperHeight;.PaperWidth;.MarginLeft;.MarginRight;.MarginTop;.MarginBottom��֯
  ҳü�ı�_In   ����ҳ���ʽ.ҳü%Type,
  ҳ���ı�_In   ����ҳ���ʽ.ҳ��%Type,
  ���ϱ�ǩ_In   Varchar2, --����"ǰ׺{��Ŀ}"��֯����"|"Ϊ�ָ��ı��ϱ�ǩ����
  ��ͷ��Ԫ_In   Varchar2, --����"�к�,���,�ı�"��֯����"|"Ϊ�ָ��ı�ͷ��Ԫ����
  ���м���_In   Varchar2, --����"�к�,�п�,��Ŀ����"��֯����"|"Ϊ�ָ��ı��м��ϣ�������Ŀ������֯Ϊ"ǰ׺{��Ŀ}��׺`�Ƿ����", �ո�ָ���
  ����ʱ��_In   Varchar2 := Null, --����"ʱ������,��ʼʱ���,����ʱ���"��֯����"|"Ϊ�ָ��ļ��ϡ�
  ���±�ǩ_In   Varchar2 := Null, --����"ǰ׺{��Ŀ}"��֯����"|"Ϊ�ָ��ı��ϱ�ǩ����
  �������_In   Number := Null
) Is
  v_Items    Varchar2(4000); --��Ŀ����
  v_Subitems Varchar2(4000); --��Ŀ����
  v_Fields   Varchar2(4000); --һ����Ŀ���������
  v_Colno    Varchar2(100); --��Ŀ�к�
  n_��id     �����ļ��ṹ.��id%Type;
  n_������� �����ļ��ṹ.�������%Type;
  n_������ �����ļ��ṹ.������%Type;
  v_�������� �����ļ��ṹ.��������%Type;
  n_�����д� �����ļ��ṹ.�����д�%Type;
  v_�����ı� �����ļ��ṹ.�����ı�%Type;
  v_�Ƿ��� �����ļ��ṹ.�Ƿ���%Type;
  v_Ҫ������ �����ļ��ṹ.Ҫ������%Type;
  v_Ҫ�ص�λ �����ļ��ṹ.Ҫ�ص�λ%Type;
  n_Ҫ�ر�ʾ �����ļ��ṹ.Ҫ�ر�ʾ%Type;
Begin
  Delete �����ļ��ṹ Where �ļ�id = �ļ�id_In;

  Update ����ҳ���ʽ
  Set ��ʽ = ҳ���ʽ_In, ҳü = ҳü�ı�_In, ҳ�� = ҳ���ı�_In
  Where ���� = 3 And ��� In (Select ҳ�� From �����ļ��б� Where Id = �ļ�id_In);

  Select �����ļ��ṹ_Id.Nextval Into n_��id From Dual;
  Insert Into �����ļ��ṹ
    (Id, �ļ�id, �������, ��������, ��������, �����ı�)
  Values
    (n_��id, �ļ�id_In, 1, 1, '���������Ժ���ʽ��˵��', '�����ʽ');
  Insert Into �����ļ��ṹ
    (Id, �ļ�id, ��id, ��������, �������, ��������, �����ı�, Ҫ������)
    Select �����ļ��ṹ_Id.Nextval, �ļ�id_In, n_��id, 4, ���, ����, �ı�, ����
    From (Select 1 As ���, 'Ŀǰ֧�ֵ���(1)�Ͷ���(2)' As ����, To_Char(��ͷ����_In) As �ı�, '��ͷ����' As ����
           From Dual
           Union All
           Select 2, '����ܹ��е�����', To_Char(������_In), '������'
           From Dual
           Union All
           Select 3, 'ÿ�е���С�߶�(�)', To_Char(��С�и�_In), '��С�и�'
           From Dual
           Union All
           Select 4, '����Ĭ������', �ı�����_In, '�ı�����'
           From Dual
           Union All
           Select 5, '����ı���ɫRGBֵ', To_Char(�ı���ɫ_In), '�ı���ɫ'
           From Dual
           Union All
           Select 6, '����ߵĻ�����ɫ', To_Char(�����ɫ_In), '�����ɫ'
           From Dual
           Union All
           Select 7, '�������������', �����ı�_In, '�����ı�'
           From Dual
           Union All
           Select 8, '���������', ��������_In, '��������'
           From Dual
           Union All
           Select 9, '��24Сʱ��ʾ��������ʼ��Χ', To_Char(��ʼʱ��_In), '��ʼʱ��'
           From Dual
           Union All
           Select 10, 'С�ڿ�ʼʱ���ʾ������ֹ', To_Char(��ֹʱ��_In), '��ֹʱ��'
           From Dual
           Union All
           Select 11, '�������������ݼ�¼������', ��������_In, '��������'
           From Dual
           Union All
           Select 13, '��Ч������', To_Char(��Ч������_In), '��Ч������'
           From Dual
           Union All
           Select 14, '����ʱ��ϲ�', To_Char(���кϲ�_In), '����ʱ��ϲ�'
           From Dual
           Union All
           Select 12, '���ϱ�������ݼ�¼����ɫ', To_Char(������ɫ_In), '������ɫ'
           From Dual
           Union All
           Select 15, 'ʱ��������', To_Char(ʱ������_In), 'ʱ��������'
           From Dual
           Union All
           Select 16, '�������', To_Char(�������_In), '�������'
           From Dual);

  Select �����ļ��ṹ_Id.Nextval Into n_��id From Dual;
  Insert Into �����ļ��ṹ
    (Id, �ļ�id, �������, ��������, ��������, �����ı�)
  Values
    (n_��id, �ļ�id_In, 2, 1, '���滻����ɵı�����Ŀ', '���ϱ�ǩ');
  If ��ͷ��Ԫ_In Is Null Then
    v_Items := Null;
  Else
    v_Items := Rtrim(���ϱ�ǩ_In) || '|';
  End If;
  n_������� := 0;
  While v_Items Is Not Null Loop
    v_Fields   := Substr(v_Items, 1, Instr(v_Items, '|') - 1);
    n_������� := n_������� + 1;
    v_�����ı� := Substr(v_Fields, 1, Instr(v_Fields, '{') - 1);
    v_Ҫ������ := Substr(v_Fields, Instr(v_Fields, '{') + 1, Instr(v_Fields, '}') - Instr(v_Fields, '{') - 1);
    If Substr(v_�����ı�, 1, 2) = Chr(13) || Chr(10) Then
      v_�Ƿ��� := 1;
      v_�����ı� := Substr(v_�����ı�, 3);
    Else
      v_�Ƿ��� := 0;
    End If;
    Insert Into �����ļ��ṹ
      (Id, �ļ�id, ��id, ��������, �������, �����ı�, Ҫ������, �Ƿ���)
    Values
      (�����ļ��ṹ_Id.Nextval, �ļ�id_In, n_��id, 4, n_�������, v_�����ı�, v_Ҫ������, v_�Ƿ���);
    v_Items := Substr(v_Items, Instr(v_Items, '|') + 1);
  End Loop;

  Select �����ļ��ṹ_Id.Nextval Into n_��id From Dual;
  Insert Into �����ļ��ṹ
    (Id, �ļ�id, �������, ��������, ��������, �����ı�)
  Values
    (n_��id, �ļ�id_In, 3, 1, '��ɱ�ͷ�ĸ���Ԫ����', '��ͷ��Ԫ');
  If ��ͷ��Ԫ_In Is Null Then
    v_Items := Null;
  Else
    v_Items := Rtrim(��ͷ��Ԫ_In) || '|';
  End If;
  While v_Items Is Not Null Loop
    v_Fields   := Substr(v_Items, 1, Instr(v_Items, '|') - 1);
    n_������� := To_Number(Substr(v_Fields, 1, Instr(v_Fields, ',', 1, 1) - 1));
    n_�����д� := To_Number(Substr(v_Fields,
                               Instr(v_Fields, ',', 1, 1) + 1,
                               Instr(v_Fields, ',', 1, 2) - Instr(v_Fields, ',', 1, 1) - 1));
    v_�����ı� := Substr(v_Fields, Instr(v_Fields, ',', 1, 2) + 1);
    Insert Into �����ļ��ṹ
      (Id, �ļ�id, ��id, ��������, �������, �����д�, �����ı�)
    Values
      (�����ļ��ṹ_Id.Nextval, �ļ�id_In, n_��id, 2, n_�������, n_�����д�, v_�����ı�);
    v_Items := Substr(v_Items, Instr(v_Items, '|') + 1);
  End Loop;

  Select �����ļ��ṹ_Id.Nextval Into n_��id From Dual;
  Insert Into �����ļ��ṹ
    (Id, �ļ�id, �������, ��������, ��������, �����ı�)
  Values
    (n_��id, �ļ�id_In, 4, 1, '����������еĶ�������', '���м���');
  If ���м���_In Is Null Then
    v_Items := Null;
  Else
    v_Items := Rtrim(���м���_In) || '|';
  End If;
  While v_Items Is Not Null Loop
    v_Fields := Substr(v_Items, 1, Instr(v_Items, '|') - 1);
    --�����������˶��չ�ϵ�к�Ϊ:��Ŀ�к�`�����к�
    v_Colno := Substr(v_Fields, 1, Instr(v_Fields, ',', 1, 1) - 1);
    If Instr(v_Colno, '`', 1, 1) > 0 Then
      n_������� := To_Number(Substr(v_Colno, 1, Instr(v_Colno, '`', 1, 1) - 1));
      n_������ := To_Number(Substr(v_Colno, Instr(v_Colno, '`', 1, 1) + 1));
    Else
      n_������� := To_Number(v_Colno);
      n_������ := Null;
    End If;
    v_�������� := Substr(v_Fields,
                     Instr(v_Fields, ',', 1, 1) + 1,
                     Instr(v_Fields, ',', 1, 2) - Instr(v_Fields, ',', 1, 1) - 1);
    v_Subitems := Substr(v_Fields, Instr(v_Fields, ',', 1, 2) + 1);
    If v_Subitems Is Null Then
      Insert Into �����ļ��ṹ
        (Id, �ļ�id, ��id, ��������, �������, ������, ��������, �����д�, �����ı�, Ҫ������, Ҫ�ص�λ)
      Values
        (�����ļ��ṹ_Id.Nextval, �ļ�id_In, n_��id, 4, n_�������, n_������, v_��������, 1, '', '', '');
    Else
      v_Subitems := Rtrim(v_Subitems) || ' ';
    End If;
    n_�����д� := 0;
    While v_Subitems Is Not Null Loop
      n_�����д� := n_�����д� + 1;
      v_Fields   := Substr(v_Subitems, 1, Instr(v_Subitems, ' ') - 1);
      v_�����ı� := Substr(v_Fields, 1, Instr(v_Fields, '{') - 1);
      v_Ҫ������ := Substr(v_Fields, Instr(v_Fields, '{') + 1, Instr(v_Fields, '}') - Instr(v_Fields, '{') - 1);
      If Instr(v_Fields, '`') > 0 Then
        v_Ҫ�ص�λ := Substr(v_Fields, Instr(v_Fields, '}') + 1, Instr(v_Fields, '`') - Instr(v_Fields, '}') - 1);
        n_Ҫ�ر�ʾ := To_Number(Substr(v_Fields, Instr(v_Fields, '`', 1, 1) + 1));
      Else
        v_Ҫ�ص�λ := Substr(v_Fields, Instr(v_Fields, '}') + 1);
        n_Ҫ�ر�ʾ := 0;
      End If;
      If n_�����д� > 1 Then
        n_������ := Null;
      End If;
      Insert Into �����ļ��ṹ
        (Id, �ļ�id, ��id, ��������, �������, ������, ��������, �����д�, �����ı�, Ҫ������, Ҫ�ص�λ, Ҫ�ر�ʾ)
      Values
        (�����ļ��ṹ_Id.Nextval, �ļ�id_In, n_��id, 4, n_�������, n_������, v_��������, n_�����д�, v_�����ı�, v_Ҫ������, v_Ҫ�ص�λ, n_Ҫ�ر�ʾ);
      v_Subitems := Substr(v_Subitems, Instr(v_Subitems, ' ') + 1);
    End Loop;
    v_Items := Substr(v_Items, Instr(v_Items, '|') + 1);
  End Loop;

  Select �����ļ��ṹ_Id.Nextval Into n_��id From Dual;
  Insert Into �����ļ��ṹ
    (Id, �ļ�id, �������, ��������, ��������, �����ı�)
  Values
    (n_��id, �ļ�id_In, 5, 1, '����ʱ�ķ�������', '����ʱ��');
  If ����ʱ��_In Is Null Then
    v_Items := Null;
  Else
    v_Items := Rtrim(����ʱ��_In) || '|';
  End If;

  n_�����д� := 0;
  While v_Items Is Not Null Loop
    v_Fields   := Substr(v_Items, 1, Instr(v_Items, '|') - 1);
    n_�����д� := n_�����д� + 1;
    Insert Into �����ļ��ṹ
      (Id, �ļ�id, ��id, ��������, �������, �����д�, �����ı�)
    Values
      (�����ļ��ṹ_Id.Nextval, �ļ�id_In, n_��id, 2, n_�����д�, n_�����д�, v_Fields);
    v_Items := Substr(v_Items, Instr(v_Items, '|') + 1);
  End Loop;

  Select �����ļ��ṹ_Id.Nextval Into n_��id From Dual;
  Insert Into �����ļ��ṹ
    (Id, �ļ�id, �������, ��������, ��������, �����ı�)
  Values
    (n_��id, �ļ�id_In, 6, 1, '���滻����ɵı�����Ŀ', '���±�ǩ');
  If ��ͷ��Ԫ_In Is Null Then
    v_Items := Null;
  Else
    v_Items := Rtrim(���±�ǩ_In) || '|';
  End If;
  n_������� := 0;
  While v_Items Is Not Null Loop
    v_Fields   := Substr(v_Items, 1, Instr(v_Items, '|') - 1);
    n_������� := n_������� + 1;
    v_�����ı� := Substr(v_Fields, 1, Instr(v_Fields, '{') - 1);
    v_Ҫ������ := Substr(v_Fields, Instr(v_Fields, '{') + 1, Instr(v_Fields, '}') - Instr(v_Fields, '{') - 1);
    If Substr(v_�����ı�, 1, 2) = Chr(13) || Chr(10) Then
      v_�Ƿ��� := 1;
      v_�����ı� := Substr(v_�����ı�, 3);
    Else
      v_�Ƿ��� := 0;
    End If;
    Insert Into �����ļ��ṹ
      (Id, �ļ�id, ��id, ��������, �������, �����ı�, Ҫ������, �Ƿ���)
    Values
      (�����ļ��ṹ_Id.Nextval, �ļ�id_In, n_��id, 4, n_�������, v_�����ı�, v_Ҫ������, v_�Ƿ���);
    v_Items := Substr(v_Items, Instr(v_Items, '|') + 1);
  End Loop;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�����ļ���ʽ_Update;
/

--92208:������,2015-12-29,���˽��ʲ�������������Ϣ
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

--92157:��ΰ��,2015-12-29,����סԺҽ���༭��������·��ҽ��ʱ�������ٴ�·���ĵ�ǰ�׶κ͵�ǰ������Ϊ��
Create Or Replace Procedure Zl_����·������_Delete
(
  ִ�м�¼id_In ����·��ִ��.Id%Type,
  ����ģʽ_In   Number := 0,
  ���ó���_In   Number := 0
) Is
  --����:����ģʽ_in=0:ȡ��·����Ŀʱ����,=1:��������ҽ��ʱ����,=2��ȡ�����ɱ������ɵ���Ŀʱ,
  --               =3:ZL_����ҽ����¼_Delete����,��ֹסԺҽ���༭����ɾ��·��ҽ��ʱ�������ٴ�·���ĵ�ǰ�׶κ͵�ǰ������Ϊ�ա�
  --     ���ó���_In =0:ҽ��վ  ;1-��ʿվ
  t_Id   t_Numlist;
  t_ʱ�� t_Strlist;
  --����ҽ��,�����׶δ���ʱ��ɾ��,δУ��ʱ��ɾ��ҽ��(������������У�Ե�δ���ϵĲ�����ɾ��·����Ŀ)
  Cursor c_Advice(����ʱ��_In �����ٴ�·��.����ʱ��%Type) Is
    Select a.����ҽ��id
    From ����·��ҽ�� A, ����ҽ����¼ C
    Where ·��ִ��id = ִ�м�¼id_In And a.����ҽ��id = c.Id And c.ҽ��״̬ = 1 And
          To_Date(To_Char(c.����ʱ�� + 59 / 24 / 60 / 60, 'yyyy-mm-dd hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') > ����ʱ��_In And
          Not Exists
     (Select 1 From ����·��ҽ�� B Where a.����ҽ��id = b.����ҽ��id And a.·��ִ��id <> b.·��ִ��id);

  Cursor c_Doc Is
    Select ID, To_Char(����ʱ��, 'yyyy-MM-dd hh24:mi:ss') From ���Ӳ�����¼ Where ·��ִ��id = ִ�м�¼id_In;

  --ɾ�����һ����Ŀʱ�����ǰ���Ƿ�����ǰ�����Ľ׶Ρ�
  Cursor c_Turn
  (
    ·����¼id_In ����·��ִ��.·����¼id%Type,
    ����_In       ����·��ִ��.����%Type,
    �׶�id_In     ����·��ִ��.�׶�id%Type
  ) Is
    Select a.�׶�id
    From ����·��ִ�� A
    Where a.·����¼id = ·����¼id_In And a.���� = ����_In And a.�׶�id <> �׶�id_In And a.��Ŀ���� = 'δ�����κ���Ŀ' And Exists
     (Select 1
           From ����·������ B
           Where a.·����¼id = b.·����¼id And a.�׶�id = b.�׶�id And a.���� = b.���� And b.ʱ����� = 1)
    Order By a.�Ǽ�ʱ�� Desc;
  t_�׶�id t_Numlist;

  Cursor c_Merge
  (
    ·����¼id_In ����·��ִ��.·����¼id%Type,
    �׶�id_In     ����·��ִ��.�׶�id%Type
  ) Is
    Select a.Id, Max(b.�ϲ�·���׶�id) As �׶�id
    From ���˺ϲ�·�� A, ���˺ϲ�·������ B
    Where a.Id = b.�ϲ�·����¼id(+) And a.��Ҫ·����¼id = b.·����¼id(+) And b.�׶�id(+) <> �׶�id_In And a.��Ҫ·����¼id = ·����¼id_In And
          (b.�Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��)
                     From ���˺ϲ�·������ C
                     Where c.·����¼id = b.·����¼id And c.�ϲ�·����¼id = b.�ϲ�·����¼id And c.�׶�id = b.�׶�id) Or b.�Ǽ�ʱ�� Is Null)
    Group By a.Id;
  t_�ϲ�·���׶�id t_Numlist;
  t_�ϲ�·����¼id t_Numlist;

  r_Pp_Item ����·��ִ��%RowType;

  v_�׶�id          ����·��ִ��.�׶�id%Type;
  v_ǰһ�׶�id      ����·��ִ��.�׶�id%Type;
  v_·����¼id      ����·��ִ��.·����¼id%Type;
  v_����            ����·��ִ��.����%Type;
  v_Last����        ����·��ִ��.����%Type;
  v_���id          ����ҽ����¼.���id%Type;
  v_Other·��ִ��id ����·��ִ��.Id%Type;
  n_Count           Number(5);
  n_�仯����        Number(5);
  d_����ʱ��        �����ٴ�·��.����ʱ��%Type;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin

  --���·��������Ŀ��һ����ҩ�ģ�ɾ������һ��·����Ŀ��Ҫ�ý����еĸ�ҩ;����ִ��ID����Ϊһ����ҩ������ҩƷ��ִ��ID
  Select Nvl(Max(b.���id), 0)
  Into v_���id
  From ����·��ҽ�� A, ����ҽ����¼ B
  Where a.·��ִ��id = ִ�м�¼id_In And a.����ҽ��id = b.Id And b.������� In ('5', '6');

  If v_���id <> 0 Then
    Select Nvl(Max(c.Id), 0)
    Into v_Other·��ִ��id
    From ����·��ҽ�� A, ����ҽ����¼ B, ����·��ִ�� C
    Where c.Id = a.·��ִ��id And b.���id = v_���id And a.����ҽ��id = b.Id And a.·��ִ��id <> ִ�м�¼id_In And
          c.�Ǽ�ʱ�� = (Select Max(d.�Ǽ�ʱ��)
                    From ����·��ִ�� D, ����ҽ����¼ E, ����·��ҽ�� F
                    Where e.Id = f.����ҽ��id And d.Id = f.·��ִ��id And e.���id = v_���id And f.·��ִ��id <> ִ�м�¼id_In);
  
    If v_Other·��ִ��id <> 0 Then
      Select Count(1)
      Into n_Count
      From ����·��ִ�� A, ����·��ִ�� B
      Where a.Id = v_Other·��ִ��id And b.Id = ִ�м�¼id_In And a.�׶�id = b.�׶�id And a.���� = b.����;
    
      If n_Count > 0 Then
        --�����������Ŀ����ͬҽ������Ӧ��ͬ��·����Ŀ���ϲ�·������Ҫ·�����ظ����ɵ�ҽ����ʱ�����޸�·��ִ��ID
        Select Count(1) Into n_Count From ����·��ҽ�� Where ·��ִ��id = v_Other·��ִ��id And ����ҽ��id = v_���id;
        If n_Count = 0 Then
          Update ����·��ҽ��
          Set ·��ִ��id = v_Other·��ִ��id
          Where ����ҽ��id = v_���id And ·��ִ��id = ִ�м�¼id_In;
        End If;
      End If;
    End If;
  End If;

  --����ʱ��֮ǰ��·����Ŀ��Ӧ�Ĳ���ҽ����ɾ��
  Select b.����ʱ��
  Into d_����ʱ��
  From ����·��ִ�� A, �����ٴ�·�� B
  Where a.·����¼id = b.Id And a.Id = ִ�м�¼id_In;
  If d_����ʱ�� Is Not Null Then
    --�Ƿ�����ȡ�����߼������ڽ�������м��
    Open c_Advice(d_����ʱ��);
    Fetch c_Advice Bulk Collect
      Into t_Id;
    Close c_Advice;
  
    Delete ����·��ҽ�� Where ·��ִ��id = ִ�м�¼id_In;
    If t_Id.Count > 0 Then
      Forall I In 1 .. t_Id.Count
        Delete From ����ҽ����¼ Where ID = t_Id(I) And ҽ��״̬ = 1;
    End If;
  
    If ����ģʽ_In = 0 Or ����ģʽ_In = 2 Then
      Open c_Doc;
      Fetch c_Doc Bulk Collect
        Into t_Id, t_ʱ��;
      Close c_Doc;
      If t_Id.Count > 0 Then
        For I In 1 .. t_Id.Count Loop
          If To_Date(t_ʱ��(I), 'yyyy-MM-dd hh24:mi:ss') > d_����ʱ�� Then
            Zl_���Ӳ�����¼_Delete(t_Id(I));
          Else
            Update ���Ӳ�����¼ Set ·��ִ��id = Null Where ID = t_Id(I);
          End If;
        End Loop;
      End If;
    End If;
  End If;

  --�����ȡ�����ɱ������ɵ���Ŀʱ����ɾ��ִ�м�¼
  If ����ģʽ_In = 3 Then
    Select * Into r_Pp_Item From ����·��ִ�� T Where ID = ִ�м�¼id_In;
    Delete ����·��ִ�� Where ID = ִ�м�¼id_In;
    Select Count(1)
    Into n_Count
    From ����·��ִ��
    Where ·����¼id = r_Pp_Item.·����¼id And �׶�id = r_Pp_Item.�׶�id And ���� = r_Pp_Item.����;
    If n_Count = 0 Then
      --����һ��������Ŀ[δ�����κ���Ŀ]
      Insert Into ����·��ִ��
        (ID, ·����¼id, �׶�id, ����, ����, ����, ��Ŀid, �Ǽ���, �Ǽ�ʱ��, ��Ŀ���, ��Ŀ����, ִ����, ������, ��Ŀ���)
      Values
        (����·��ִ��_Id.Nextval, r_Pp_Item.·����¼id, r_Pp_Item.�׶�id, r_Pp_Item.����, r_Pp_Item.����, r_Pp_Item.����, Null,
         Zl_Username, Sysdate, Null, 'δ�����κ���Ŀ', Null, 1, '�Ѿ�ִ��|1' || Chr(9) || '�Ѿ�ִ��');
    End If;
  Elsif ����ģʽ_In <> 2 Then
    Delete ����·��ִ��
    Where ID = ִ�м�¼id_In
    Returning ·����¼id, �׶�id, ���� Into v_·����¼id, v_�׶�id, v_����;
  End If;

  If ����ģʽ_In = 0 And ���ó���_In = 0 Then
    Select Count(1)
    Into n_Count
    From ����·��ִ��
    Where ·����¼id = v_·����¼id And �׶�id = v_�׶�id And ���� = v_����;
    If n_Count = 0 Then
      Select Max(����) Into v_���� From ����·��ִ�� Where ·����¼id = v_·����¼id And �׶�id = v_�׶�id;
      Select Max(����) Into v_Last���� From ����·��ִ�� Where ·����¼id = v_·����¼id;
      --��¼�仯������
      Select ��ǰ���� Into n_�仯���� From �����ٴ�·�� Where ID = v_·����¼id;
      --�����ǰ�׶ε����һ��ִ�м�¼��ɾ��(ȫ�����ǷǱ���ִ�е������)
      --����·����ת��һ���׶ε�������������һ��·���Ľ׶ν��棨���磺a·����3�׶�:3-5��,��ִ�е�3�죬��ת������·����������ִ�е�5�죩
      If v_���� Is Null Or v_���� <> v_Last���� Then
        --a.�����ǰû���κ�ִ�м�¼
        If v_Last���� Is Null Then
          Update �����ٴ�·��
          Set ǰһ�׶�id = Null, ��ǰ�׶�id = Null, ��ǰ���� = Null, ״̬ = 1
          Where ID = v_·����¼id;
          Update ���˺ϲ�·��
          Set ǰһ�׶�id = Null, ��ǰ�׶�id = Null, ��ǰ���� = Null
          Where ��Ҫ·����¼id = v_·����¼id;
        Else
          --b.���˵�ǰһ���׶�
          --���ǰһ�׶��������Ľ׶Σ���ֱ��ɾ��
          Open c_Turn(v_·����¼id, v_Last����, v_�׶�id);
          Fetch c_Turn Bulk Collect
            Into t_�׶�id;
          Close c_Turn;
          If t_�׶�id.Count > 0 Then
            Forall I In 1 .. t_�׶�id.Count
              Delete From ����·������ Where ·����¼id = v_·����¼id And �׶�id = t_�׶�id(I) And ���� = v_Last����;
            Forall I In 1 .. t_�׶�id.Count
              Delete From ����·��ִ�� Where ·����¼id = v_·����¼id And �׶�id = t_�׶�id(I) And ���� = v_Last����;
            --ɾ����ȡ���һ���׶�Ϊǰһ�׶�ID
            Select Max(�׶�id)
            Into v_ǰһ�׶�id
            From ����·��ִ��
            Where ·����¼id = v_·����¼id And �Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From ����·��ִ�� Where ·����¼id = v_·����¼id);
            Update �����ٴ�·�� Set ǰһ�׶�id = v_ǰһ�׶�id Where ID = v_·����¼id;
          End If;
          --�޸Ĳ����ٴ�·����Ϣ
          Select Max(�׶�id)
          Into v_�׶�id
          From ����·��ִ��
          Where ·����¼id = v_·����¼id And
                �Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��)
                        From ����·��ִ��
                        Where ·����¼id = v_·����¼id And �׶�id <> (Select ǰһ�׶�id From �����ٴ�·�� Where ID = v_·����¼id));
          --���»�ȡ��ǰ����
          --��ǰһ�׶�����Ϊ����һ�׶���ǰ�����죨ʱ�����=2���ҵڶ�������ʱ�ڿ�ѡ�׶����������м�׶����ɺ���Ľ׶�ʱ,
          --���ֳ������ɵ�·����ִ��ȡ����������ʱ��v_Last������Ҫ���»�ȡ��
          Select Max(����) Into v_Last���� From ����·��ִ�� Where ·����¼id = v_·����¼id;
          --
          Update �����ٴ�·��
          Set ��ǰ�׶�id = ǰһ�׶�id, ǰһ�׶�id = v_�׶�id, ��ǰ���� = v_Last����, ״̬ = 1
          Where ID = v_·����¼id;
        
          n_�仯���� := n_�仯���� - v_Last����;
          --�޸Ĳ����ٴ��ϲ�·����Ϣ
          Select Nvl(��ǰ�׶�id, 0) Into v_�׶�id From �����ٴ�·�� Where ID = v_·����¼id;
          Open c_Merge(v_·����¼id, v_�׶�id);
          Fetch c_Merge Bulk Collect
            Into t_�ϲ�·����¼id, t_�ϲ�·���׶�id;
          Close c_Merge;
          If t_�ϲ�·���׶�id.Count > 0 Then
            Forall I In 1 .. t_�ϲ�·���׶�id.Count
              Update ���˺ϲ�·��
              Set ��ǰ���� = Decode(ǰһ�׶�id, Null, Null, Nvl(��ǰ����, 0) - Nvl(n_�仯����, 0)), ��ǰ�׶�id = ǰһ�׶�id,
                  ǰһ�׶�id = t_�ϲ�·���׶�id(I)
              Where ��Ҫ·����¼id = v_·����¼id And ID = t_�ϲ�·����¼id(I);
          End If;
        End If;
      Else
        --���һ���׶��ж��죬ȡ�����һ����Ŀʱ��ֻ��������
        Update �����ٴ�·�� Set ��ǰ���� = v_���� Where ID = v_·����¼id And ��ǰ���� <> v_����;
        n_�仯���� := n_�仯���� - v_����;
        If n_�仯���� <> 0 Then
          --�޸Ĳ����ٴ��ϲ�·����Ϣ
          Select Nvl(��ǰ�׶�id, 0) Into v_�׶�id From �����ٴ�·�� Where ID = v_·����¼id;
          Open c_Merge(v_·����¼id, v_�׶�id);
          Fetch c_Merge Bulk Collect
            Into t_�ϲ�·����¼id, t_�ϲ�·���׶�id;
          Close c_Merge;
          If t_�ϲ�·���׶�id.Count > 0 Then
            Forall I In 1 .. t_�ϲ�·���׶�id.Count
              Update ���˺ϲ�·��
              Set ��ǰ���� = Decode(ǰһ�׶�id, Null, Null, Nvl(��ǰ����, 0) - Nvl(n_�仯����, 0)), ��ǰ�׶�id = ǰһ�׶�id,
                  ǰһ�׶�id = t_�ϲ�·���׶�id(I)
              Where ��Ҫ·����¼id = v_·����¼id And ID = t_�ϲ�·����¼id(I);
          End If;
        End If;
      End If;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����·������_Delete;
/

--92157:��ΰ��,2015-12-29,����סԺҽ���༭��������·��ҽ��ʱ�������ٴ�·���ĵ�ǰ�׶κ͵�ǰ������Ϊ��
CREATE OR REPLACE Procedure Zl_����ҽ����¼_Delete
(
  --���ܣ�ɾ��ָ��ҽ��,�����������סԺ��
  --������
  --      ҽ��ID_IN����ǰҪɾ����ҽ����ID(�ǿɼ��еĵ���ҽ��ID��������ID)
  --      ɾ���_IN=0ʱ,ֻɾ��ָ��ID��ҽ��(ҽ���༭�������)��
  --          1.���ҽ����ͬ��ɾ����ɾ��֮�����ŵ����ɳ��������ö�Ӧ���̡�
  --          2.��ɾ����ҽ��Ӧ��δУ��,����Ӧ�ѿ��ơ�
  --          3.����ҽ��״̬�����ݻ��Զ�ɾ��������ҽ���Ƽۣ�����ҽ������δУ�Ե�û�м�¼��
  --      ɾ���_IN=1ʱ,ɾ������ҽ��(����������)�����ҩ;���������ϣ����������ҩ�䷽��
  --          1.��Ҫ�ڹ�����ͬʱ������ؼ�¼����š�
  --          2.һ����ҩ��ֻɾ����ǰҩƷ��¼(��������ҩ;��)��
  ҽ��id_In ����ҽ����¼.Id%Type,
  ɾ���_In Number := 0
) Is
  v_״̬            ����ҽ����¼.ҽ��״̬%Type;
  v_���id          ����ҽ����¼.���id%Type;
  v_����id          ����ҽ����¼.����id%Type;
  v_�Һŵ�          ����ҽ����¼.�Һŵ�%Type;
  v_��ҳid          ����ҽ����¼.��ҳid%Type;
  v_Ӥ��            ����ҽ����¼.Ӥ��%Type;
  v_���            ����ҽ����¼.���%Type;
  v_����            ����ҽ����¼.ҽ������%Type;
  v_·��ִ��id      ����·��ִ��.Id%Type;
  v_Other·��ִ��id ����·��ִ��.Id%Type;
  v_·��ִ�з�ʽ    �ٴ�·����Ŀ.ִ�з�ʽ%Type;
  v_����Ҫ��        �ٴ�·����Ŀ.����Ҫ��%Type;
  v_Count           Number(5);
  v_·����¼id      �����ٴ�·��.Id%Type;
  v_����ԭ��        ����·��ִ��.����ԭ��%Type;
  n_�Ƿ�����        Number(5);
  n_·����Ŀid      ����·��ִ��.��Ŀid%Type;
  n_Islast          Number(5);
  n_Del_Count       Number(5);
  n_Del����         Number(2); --0-ֻɾ��ָ��ID��ҽ����1-ɾ������ҽ��
  v_�������        ����ҽ����¼.�������%Type;
  v_���״̬        ����ҽ����¼.���״̬%Type;
  v_������ĿID      ����ҽ����¼.������ĿID%Type;
  v_����Ѫ��        zlParameters.����ֵ%Type;
  v_ִ�з���        ������ĿĿ¼.ִ�з���%type;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --���ҽ��״̬:��������
  Begin
    Select ����id, �Һŵ�, ��ҳid, Ӥ��, ҽ��״̬, ���id, ҽ������, �������, ���״̬,������ĿID
    Into v_����id, v_�Һŵ�, v_��ҳid, v_Ӥ��, v_״̬, v_���id, v_����, v_�������, v_���״̬,v_������ĿID
    From ����ҽ����¼
    Where ID = ҽ��id_In;
  Exception
    When Others Then
      Begin
        v_Error := 'δ����Ҫɾ����ҽ����¼�������ѱ�������ɾ����';
        Raise Err_Custom;
      End;
  End;
  If v_�Һŵ� Is Null Then
    If Not v_״̬ In (1, 2, -1) Then
      v_Error := 'ҽ��"' || v_���� || '"�Ѿ���У�ԣ�������ɾ����';
      Raise Err_Custom;
    End If;
  Else
    If v_״̬ <> 1 Then
      v_Error := 'ҽ��"' || v_���� || '"�Ѿ������ͻ����ϣ�����ɾ����';
      Raise Err_Custom;
    End If;
  End If;

  --��Ѫҽ����������
  if v_������� = 'K' and v_���״̬ in (2,5) then 
    --�Ƿ�װ��Ѫ��ϵͳ
    Select count(1) into v_Count  From zlSystems Where ���=2200;
    If Nvl(v_Count, 0) > 0 Then
      select ִ�з��� into v_ִ�з��� from ������ĿĿ¼ where  ID = v_������ĿID;
      if not (nvl(v_ִ�з���,0) = 1) then
        --�Ƿ�������Ѫ�����ϵͳ
        Select Zl_Getsysparameter(236) into v_����Ѫ�� From Dual;
        if Nvl(v_����Ѫ��,'0') <> '0' then
          if v_���״̬ = 5 then
             v_Error := '������Ѫ��';
          else
             v_Error := '�����������Ѫ��';
          end if;
          v_Error := 'ҽ��"' || v_���� ||'"�ѱ�Ѫ����գ�' || v_Error || '����ɾ��������ɾ��������Ѫ����ϵ��';
          Raise Err_Custom;
        end if;
      end if;
    end if;
  end if;

  Select Count(*)
  Into v_Count
  From ����ҽ��״̬
  Where ҽ��id = ҽ��id_In And �������� In (1, 11) And ǩ��id Is Not Null;
  If Nvl(v_Count, 0) > 0 Then
    v_Error := 'ҽ��"' || v_���� || '"�Ѿ�����ǩ��,����ɾ����';
    Raise Err_Custom;
  End If;

  --�ж�ɾ���黹��ָ��ID��ҽ��
  If Nvl(ɾ���_In, 0) = 0 Then
    n_Del���� := 0;
  Else
    If v_���id Is Null Then
      --������,����������,��ҩ�䷽,�������,�Լ�����ҽ��
      Select Max(���), Count(*) Into v_���, n_Del_Count From ����ҽ����¼ Where ID = ҽ��id_In Or ���id = ҽ��id_In;
      n_Del���� := 1;
    Else
      --��ҩһ����ҩ�����(������)
      --���ж��Ƿ�һ����ҩ
      Select Count(*) Into n_Del_Count From ����ҽ����¼ Where ���id = v_���id;
      If n_Del_Count = 1 Then
        --������ҩ:ͬʱɾ�����ҩ;��
        Select Max(���), Count(*) Into v_���, n_Del_Count From ����ҽ����¼ Where ID = ҽ��id_In Or ID = v_���id;
        n_Del���� := 1;
      Else
        --һ����ҩ:ֻɾ����ǰҩƷ
        n_Del_Count := 1;
        Select ��� Into v_��� From ����ҽ����¼ Where ID = ҽ��id_In;
        n_Del���� := 0;
      End If;
    End If;
  End If;
  --
  Begin
    --�������·����ҽ�����򲻲�ѯ·��ִ�б�
    Select Count(1) Into v_Count From ����·��ҽ�� Where ����ҽ��id = ҽ��id_In;

    If v_Count > 0 Then
      --����������Ϊ·������Ŀ����Ŀid�ǿ�
      --�α�ѭ����������ͬһ�죬ͬһ��ҽ����Ӧ���·����Ŀ�������
      --�������ɵ���Ŀ��Ҫ��д����ԭ�򣬲�ɾ��Ŀ��ֻɾҽ�����Ǳ������ɵ���Ŀ��ֱ��ɾ��ҽ���Ͷ�Ӧ��Ŀ
      For Rs In (Select a.Id, d.ִ�з�ʽ, d.����Ҫ��, a.·����¼id, a.����ԭ��, a.��Ŀid
                 From ����·��ִ�� A, ����·��ҽ�� B, �ٴ�·����Ŀ D
                 Where b.����ҽ��id = ҽ��id_In And b.·��ִ��id = a.Id And a.��Ŀid = d.Id(+)) Loop
        v_·��ִ��id   := Rs.Id;
        v_·��ִ�з�ʽ := Rs.ִ�з�ʽ;
        v_����Ҫ��     := Rs.����Ҫ��;
        v_·����¼id   := Rs.·����¼id;
        v_����ԭ��     := Rs.����ԭ��;
        n_·����Ŀid   := Rs.��Ŀid;

        Select Count(1)
        Into n_�Ƿ�����
        From ����·��ִ�� A, ����·������ B
        Where a.·����¼id = b.·����¼id And a.�׶�id = b.�׶�id And a.���� = b.���� And a.Id = v_·��ִ��id;

        If n_�Ƿ����� > 0 Then
          v_Error := '��ҽ����Ӧ���ٴ�·����Ŀ�Ѿ���������ȡ��������ɾ����';
          Raise Err_Custom;
        End If;
        --����ʱ���˱���ԭ��ı������õ���Ŀ����ɾ��
        If Not v_·��ִ�з�ʽ Is Null And v_����ԭ�� Is Null Then
          If v_·��ִ�з�ʽ <> 3 Then
            --����������ɵ���Ŀ��ѡ�����ɵ�ҽ����ʣ���һ����������ɾ��
            If v_����Ҫ�� = 1 Then
              --·�������ҽ������һ����ҩʱ������ɾ��ԭ�еĸ�ҩ;��
              If Nvl(ɾ���_In, 0) = 0 And v_���id Is Null Then
                Select Count(*)
                Into v_Count
                From ����·��ҽ�� A
                Where a.·��ִ��id = v_·��ִ��id And a.����ҽ��id <> ҽ��id_In;
              Else
                Select Count(*)
                Into v_Count
                From ����·��ҽ�� A
                Where a.·��ִ��id = v_·��ִ��id And
                      a.����ҽ��id Not In
                      (Select ID From ����ҽ����¼ Where ID = ҽ��id_In Or ���id = ҽ��id_In Or ID = v_���id);
              End If;
              If v_Count = 0 Then
                v_Error := '��ҽ����Ӧ���ٴ�·����Ŀ���Ǳ�Ҫʱ���ɵģ�����ɾ����';
                Raise Err_Custom;
              End If;
            Else
              --ִ�з�ʽ��0-����ִ��(Ҳ������ʵ�ֱ�ǩ)��1-ÿ��ִ�У�2-����ִ��һ�Σ�3-��Ҫʱִ��,4-����ִ��һ�Σ����ڽ׶α����ҽ�ִ��һ�Σ�

              If v_·��ִ�з�ʽ = 2 Or v_·��ִ�з�ʽ = 4 Then
                --��������Ѿ��������׶�Ϊ����������,��ǰ�������ǽ׶����һ��ʱ,ִ�з�ʽΪ2��4ʱ����������ӱ���ԭ���ɾ��ҽ���ģ�
        Null;
        Else
                v_Error := '��ҽ����Ӧ���ٴ�·����Ŀ�Ǳ������ɵģ���Ҫ��ӱ���ԭ�����ɾ����';
                Raise Err_Custom;
              End If;
            End If;
          End If;
        End If;

        --�ж��Ƿ������һ��ҽ�����ּ��������һ��ҽ����һ��ҽ��
        If n_Del���� = 0 Then
          Select Count(1)
          Into n_Islast
          From �����ٴ�·�� A, ����·��ִ�� B, ����·��ҽ�� C
          Where a.Id = b.·����¼id And a.��ǰ�׶�id = b.�׶�id And a.��ǰ���� = b.���� And b.Id = c.·��ִ��id And a.Id = v_·����¼id;
        Else
          --n_Del����=1�Ķ���ɾ����ҽ�����п��ܴ���������ID=null��Ҳ�п����Ǵ�������ID<>null��
          If v_���id Is Null Then
            Select Decode(Count(1), 0, 1, 0)
            Into n_Islast
            From �����ٴ�·�� A, ����·��ִ�� B, ����·��ҽ�� C, ����ҽ����¼ D
            Where a.Id = b.·����¼id And a.��ǰ�׶�id = b.�׶�id And a.��ǰ���� = b.���� And a.Id = v_·����¼id And c.����ҽ��id = d.Id And
                  c.·��ִ��id = b.Id And (d.Id <> ҽ��id_In And d.���id <> ҽ��id_In);
          Else
            Select Decode(Count(1), 0, 1, 0)
            Into n_Islast
            From �����ٴ�·�� A, ����·��ִ�� B, ����·��ҽ�� C
            Where a.Id = b.·����¼id And a.��ǰ�׶�id = b.�׶�id And a.��ǰ���� = b.���� And c.·��ִ��id = b.Id And a.Id = v_·����¼id And
                  (c.����ҽ��id <> ҽ��id_In And c.����ҽ��id <> v_���id);
          End If;
        End If;

        If n_Islast = 1 Then
          --�����һ����Ŀ�����һ��ҽ�����͵���·����Ŀɾ���Ĺ���
          If v_����ԭ�� Is Null Or Nvl(n_·����Ŀid, 0) = 0 Then
            Zl_����·������_Delete(v_·��ִ��id,3);
          Else
            --�������ɵ�û��������д������ԭ��Ĳ�ɾ����Ŀ
            Zl_����·������_Delete(v_·��ִ��id, 2);
          End If;
        Else
          If n_Del���� = 0 Then
            Delete From ����·��ҽ�� Where ����ҽ��id = ҽ��id_In And ·��ִ��id = v_·��ִ��id;

            --�����ǰҩƷɾ���󣬸�ִ��id��ֻʣ��ҩ;������Ҫ��ִ��ID����Ϊһ����ҩ������ҩƷ��ִ��ID
            If v_���id Is Not Null Then
              Select Nvl(Max(c.Id), 0)
              Into v_Other·��ִ��id
              From ����·��ҽ�� A, ����ҽ����¼ B, ����·��ִ�� C
              Where c.Id = a.·��ִ��id And b.���id = v_���id And b.Id <> ҽ��id_In And a.����ҽ��id = b.Id And
                    a.·��ִ��id <> v_·��ִ��id And c.�Ǽ�ʱ�� = (Select Max(d.�Ǽ�ʱ��)
                                                       From ����·��ִ�� D, ����ҽ����¼ E, ����·��ҽ�� F
                                                       Where e.Id = f.����ҽ��id And d.Id = f.·��ִ��id And e.���id = v_���id And
                                                             f.·��ִ��id <> v_·��ִ��id);

              If v_Other·��ִ��id <> 0 Then
                Select Count(1)
                Into v_Count
                From ����·��ִ�� A, ����·��ִ�� B
                Where a.Id = v_Other·��ִ��id And b.Id = v_·��ִ��id And a.�׶�id = b.�׶�id And a.���� = b.����;

                If v_Count > 0 Then
                  Update ����·��ҽ��
                  Set ·��ִ��id = v_Other·��ִ��id
                  Where ����ҽ��id = v_���id And ·��ִ��id = v_·��ִ��id And Not Exists
                   (Select 1 From ����·��ҽ�� C Where ·��ִ��id = v_·��ִ��id And ����ҽ��id <> v_���id);
                End If;
              End If;
            End If;
          Else
            If v_���id Is Null Then
              Delete From ����·��ҽ��
              Where ·��ִ��id = v_·��ִ��id And
                    ����ҽ��id In (Select ID From ����ҽ����¼ Where ID = ҽ��id_In Or ���id = ҽ��id_In);
            Else
              --������ҩ:ͬʱɾ�����ҩ;��
              Delete From ����·��ҽ��
              Where ·��ִ��id = v_·��ִ��id And ����ҽ��id In (Select ID From ����ҽ����¼ Where ID = ҽ��id_In Or ID = v_���id);
            End If;
          End If;
          --����Ŀ��Ӧ��ҽ��ɾ�������ɾ��·����Ŀ���������ԭ��Ϊ���Ҳ���·������Ŀ��ɾ����Ŀ��ֻɾҽ��
          If v_����ԭ�� Is Null Or Nvl(n_·����Ŀid, 0) = 0 Then
            Delete From ����·��ִ��
            Where ID = v_·��ִ��id And Not Exists (Select 1 From ����·��ҽ�� Where ·��ִ��id = v_·��ִ��id);
          End If;
        End If;
      End Loop;
    End If;
  End;

  --ɾ��������Ϻ�ɾ��ҽ��
  If n_Del���� = 0 Then
    Delete From �������ҽ�� Where ҽ��id = ҽ��id_In;
    Delete From ����ҽ����¼ Where ID = ҽ��id_In;
  Else
    If v_���id Is Null Then
      Delete From �������ҽ��
      Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = ҽ��id_In Or ���id = ҽ��id_In);
      Delete From ����ҽ����¼ Where ID = ҽ��id_In Or ���id = ҽ��id_In;
    Else
      --������ҩ:ͬʱɾ�����ҩ;��
      Delete From �������ҽ�� Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = ҽ��id_In Or ID = v_���id);
      Delete From ����ҽ����¼ Where ID = ҽ��id_In Or ID = v_���id;
    End If;
  End If;

  If Nvl(ɾ���_In, 0) <> 0 Then
    --�������
    Update ����ҽ����¼
    Set ��� = ��� - n_Del_Count
    Where ����id = v_����id And Nvl(��ҳid, 0) = Nvl(v_��ҳid, 0) And Nvl(�Һŵ�, '��') = Nvl(v_�Һŵ�, '��') And
          Nvl(Ӥ��, 0) = Nvl(v_Ӥ��, 0) And ��� > v_���;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ҽ����¼_Delete;
/

--91385:Ƚ����,2015-12-29,�޸ġ��ǱҶ�����룬�������塱�����������.
Create Or Replace Procedure Zl_���˽����¼_Update
(
  ����id_In       ����Ԥ����¼.����id%Type,
  ���ս���_In     Varchar2, --"���㷽ʽ|������||....."
  ����_In         Number := 0,
  ȱʡ���㷽ʽ_In Varchar2 := Null,
  ȱʡ��Ԥ��_In   Number := 0, --0-���ֽ�ɿ�,1:ʣ�ڿ����ó�Ԥ��֧��(����Ԥ��),2-ʣ�ڿ����ó�Ԥ��֧��(סԺԤ��)
  �����id_In     ����Ԥ����¼.�����id%Type := Null,
  ���㿨���_In   ����Ԥ����¼.���㿨���%Type := Null,
  ����_In         ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In   ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In     ����Ԥ����¼.����˵��%Type := Null,
  ������λ_In     ����Ԥ����¼.������λ%Type := Null
  
) As
  --���α�ΪҪɾ�����ɷ��ü�¼�����Ľ����¼

  Cursor c_Del Is
    Select a.Id, a.��¼����, a.��Ԥ��, a.���㷽ʽ, b.����, a.Ԥ�����
    From ����Ԥ����¼ A, ���㷽ʽ B
    Where a.���㷽ʽ = b.���� And a.����id = ����id_In;

  --�����Ϣ
  v_No         ����Ԥ����¼.No%Type;
  v_����id     סԺ���ü�¼.����id%Type;
  v_��ҳid     סԺ���ü�¼.��ҳid%Type;
  v_����ʱ��   סԺ���ü�¼.����ʱ��%Type;
  v_�Ǽ�ʱ��   סԺ���ü�¼.�Ǽ�ʱ��%Type;
  v_����Ա��� סԺ���ü�¼.����Ա���%Type;
  v_����Ա���� סԺ���ü�¼.����Ա����%Type;

  --���ν������
  v_���ϼ� ����Ԥ����¼.��Ԥ��%Type;

  --���ս���
  v_���ս��� Varchar2(255);
  v_��ǰ���� Varchar2(50);
  v_���㷽ʽ ����Ԥ����¼.���㷽ʽ%Type;
  v_������ ����Ԥ����¼.��Ԥ��%Type;

  v_��¼���� ����Ԥ����¼.��¼����%Type;
  v_ȱʡ     ����Ԥ����¼.���㷽ʽ%Type;

  --�ֱҴ���������
  v_�ֽ���   ����Ԥ����¼.��Ԥ��%Type;
  v_Cashcented ����Ԥ����¼.��Ԥ��%Type;
  v_�����   ����Ԥ����¼.��Ԥ��%Type;
  v_����id     סԺ���ü�¼.Id%Type;
  v_���       סԺ���ü�¼.���%Type;
  v_�շ����   סԺ���ü�¼.�շ����%Type;
  v_�շ�ϸĿid סԺ���ü�¼.�շ�ϸĿid%Type;
  v_������Ŀid סԺ���ü�¼.������Ŀid%Type;
  v_�վݷ�Ŀ   סԺ���ü�¼.�վݷ�Ŀ%Type;
  n_Noexists   Number(3);
  n_ҽ��С��id סԺ���ü�¼.ҽ��С��id%Type;
  n_�������   ����Ԥ����¼.�������%Type;
  n_����״̬   ������ü�¼.����״̬%Type;
  n_Ԥ�����   ����Ԥ����¼.���%Type;
  n_��ǰ���   ����Ԥ����¼.���%Type;
  v_�����     ���㷽ʽ.����%Type;

  --��ʱ����
  Err_Item Exception;
  v_Err_Msg  Varchar2(255);
  n_��id     ����ɿ����.Id%Type;
  n_ִ��״̬ ������ü�¼.ִ��״̬%Type;
Begin
  --���ȱʡ���㷽ʽΪ�գ���ȡ�ֽ���㷽ʽ
  If ȱʡ���㷽ʽ_In Is Null Then
    Begin
      Select ���� Into v_ȱʡ From ���㷽ʽ Where ���� = 1 And Rownum < 2;
    Exception
      When Others Then
        v_ȱʡ := '�ֽ�';
    End;
  Else
    v_ȱʡ := ȱʡ���㷽ʽ_In;
  End If;

  --ȡ�ñ��ν���������Ϣ
  If Nvl(����_In, 0) = 1 Then
    Select NO, ����id, �շ�ʱ��, ����Ա���, ����Ա����, �ɿ���id, 0
    Into v_No, v_����id, v_�Ǽ�ʱ��, v_����Ա���, v_����Ա����, n_��id, n_ִ��״̬
    From ���˽��ʼ�¼
    Where ID = ����id_In;
  Else
    Begin
      n_Noexists := 0;
      Select NO, ����id, �Ǽ�ʱ��, ����Ա���, ����Ա����, �ɿ���id, ִ��״̬, ����״̬
      Into v_No, v_����id, v_�Ǽ�ʱ��, v_����Ա���, v_����Ա����, n_��id, n_ִ��״̬, n_����״̬
      From ������ü�¼
      Where ����id = ����id_In And Rownum < 2;
    Exception
      When Others Then
        n_Noexists := 1;
    End;
    If n_Noexists = 1 Then
      --���ü�¼�����ڣ��Ӳ����¼����
      Select NO, ����id, �Ǽ�ʱ��, ����Ա���, ����Ա����, �ɿ���id, ����״̬
      Into v_No, v_����id, v_�Ǽ�ʱ��, v_����Ա���, v_����Ա����, n_��id, n_����״̬
      From ���ò����¼
      Where ����id = ����id_In And Rownum < 2;
    End If;
    If Nvl(n_����״̬, 0) = 1 Then
      --�쳣����Ϊ��:
      v_ȱʡ := Null;
    End If;
  
    Begin
      --20051027 �¶�
      Select ��¼����
      Into v_��¼����
      From ����Ԥ����¼
      Where ����id = ����id_In And Rownum = 1 And Mod(��¼����, 10) <> 1;
    Exception
      When Others Then
        v_��¼���� := -1;
    End;
    If v_��¼���� = -1 Then
      Begin
        Select Decode(��¼����, 1, 3, 11, 3, 4, 4, ��¼����)
        Into v_��¼����
        From ������ü�¼
        Where ����id = ����id_In And Rownum = 1;
      Exception
        When Others Then
          --�����ǿ���
          Select ��¼���� Into v_��¼���� From סԺ���ü�¼ Where ����id = ����id_In And Rownum = 1;
      End;
    End If;
  End If;

  If Nvl(v_����id, 0) <> 0 And Nvl(����_In, 0) = 1 Then
    Select ��ҳid Into v_��ҳid From ������Ϣ Where ����id = v_����id;
  End If;
  Select ������� Into n_������� From ����Ԥ����¼ Where ����id = ����id_In And Rownum = 1;

  ----���˽ɿ�,Ԥ������,��Ϊû�иĳ�Ԥ����
  --�շ�δ��δ������ɵ�,�����쳣��������,��������Ա�ɿ����
  v_���ϼ� := 0;
  For r_Del In c_Del Loop
    If r_Del.��¼���� Not In (1, 11) Then
      If Nvl(n_����״̬, 0) <> 1 Then
        Update ��Ա�ɿ����
        Set ��� = Nvl(���, 0) - r_Del.��Ԥ��
        Where �տ�Ա = v_����Ա���� And ���� = 1 And ���㷽ʽ = r_Del.���㷽ʽ;
      
        If Sql%RowCount = 0 Then
          Insert Into ��Ա�ɿ����
            (�տ�Ա, ���㷽ʽ, ����, ���)
          Values
            (v_����Ա����, r_Del.���㷽ʽ, 1, -1 * r_Del.��Ԥ��);
        End If;
      End If;
      v_���ϼ� := v_���ϼ� + r_Del.��Ԥ��;
      Delete From ����Ԥ����¼ Where ID = r_Del.Id;
    Else
      --����Ƿ��Ԥ��
      If Nvl(ȱʡ��Ԥ��_In, 0) <> 0 Then
        v_���ϼ� := v_���ϼ� + r_Del.��Ԥ��;
        If Nvl(n_����״̬, 0) <> 1 Then
          Update �������
          Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(r_Del.��Ԥ��, 0)
          Where ����id = v_����id And ���� = Nvl(r_Del.Ԥ�����, 2);
          If Sql%NotFound Then
            Insert Into �������
              (����id, ����, Ԥ�����, �������, ����)
            Values
              (v_����id, 1, Nvl(r_Del.��Ԥ��, 0), 0, Nvl(r_Del.Ԥ�����, 2));
          End If;
        End If;
        If r_Del.��¼���� = 1 Then
          Update ����Ԥ����¼ Set ��Ԥ�� = 0 Where ID = r_Del.Id;
        Else
          Delete ����Ԥ����¼ Where ID = r_Del.Id;
        End If;
      End If;
    End If;
  End Loop;

  --------------------------------------------------------------------------------------------------------------
  --------------------------------------------------------------------------------------------------------------
  --����ҽ��֧������
  If ���ս���_In Is Not Null Then
    --�������ս���
    v_���ս��� := ���ս���_In || '||';
    While v_���ս��� Is Not Null Loop
      v_��ǰ���� := Substr(v_���ս���, 1, Instr(v_���ս���, '||') - 1);
    
      v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
      v_������ := To_Number(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
    
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, ��������)
      Values
        (����Ԥ����¼_Id.Nextval, Decode(����_In, 1, 2, v_��¼����), v_No, 1, v_����id, v_��ҳid, '���ղ���', v_���㷽ʽ, v_�Ǽ�ʱ��, v_����Ա���,
         v_����Ա����, v_������, ����id_In, n_��id, n_�������, Mod(Decode(����_In, 1, 2, v_��¼����), 10));
    
      v_���ϼ� := v_���ϼ� - v_������;
    
      v_���ս��� := Substr(v_���ս���, Instr(v_���ս���, '||') + 2);
    End Loop;
  End If;
  --ʣ�ಿ����Ԥ��
  If Nvl(ȱʡ��Ԥ��_In, 0) <> 0 And v_���ϼ� <> 0 Then
  
    n_Ԥ����� := v_���ϼ�;
    For c_Ԥ�� In (Select *
                 From (Select a.Id, a.��¼״̬, a.Ԥ�����, a.No, Nvl(a.���, 0) As ���
                        From ����Ԥ����¼ A,
                             (Select NO, Sum(Nvl(a.���, 0)) As ���
                               From ����Ԥ����¼ A
                               Where a.����id Is Null And Nvl(a.���, 0) <> 0 And a.����id = v_����id And Nvl(a.Ԥ�����, 2) = ȱʡ��Ԥ��_In
                               Group By NO
                               Having Sum(Nvl(a.���, 0)) <> 0) B
                        Where a.����id Is Null And Nvl(a.���, 0) <> 0 And
                              a.���㷽ʽ Not In (Select ���� From ���㷽ʽ Where ���� = 5) And a.No = b.No And a.����id = v_����id And
                              Nvl(a.Ԥ�����, 2) = ȱʡ��Ԥ��_In
                        Union All
                        Select 0 As ID, ��¼״̬, Ԥ�����, NO, Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) As ���
                        From ����Ԥ����¼
                        Where ��¼���� In (1, 11) And ����id Is Not Null And Nvl(���, 0) <> Nvl(��Ԥ��, 0) And ����id = v_����id And
                              Nvl(Ԥ�����, 2) = ȱʡ��Ԥ��_In Having Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0
                        Group By ��¼״̬, NO, Ԥ�����)
                 Order By ID, NO) Loop
    
      n_��ǰ��� := Case
                  When c_Ԥ��.��� - n_Ԥ����� < 0 Then
                   c_Ԥ��.���
                  Else
                   n_Ԥ�����
                End;
    
      If c_Ԥ��.Id <> 0 Then
        --��һ�γ�Ԥ��(82592,����һ�α��Ͻ���ID,��Ԥ�����Ϊ0)
        Update ����Ԥ����¼
        Set ��Ԥ�� = 0, ����id = ����id_In, ������� = n_�������, �������� = Mod(Decode(����_In, 1, 2, v_��¼����), 10)
        Where ID = c_Ԥ��.Id;
      End If;
      --���ϴ�ʣ���
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���,
         ��Ԥ��, ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, ��������)
        Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�,
               v_�Ǽ�ʱ��, v_����Ա����, v_����Ա���, n_��ǰ���, ����id_In, n_��id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, n_�������,
               Mod(Decode(����_In, 1, 2, v_��¼����), 10)
        From ����Ԥ����¼
        Where NO = c_Ԥ��.No And ��¼״̬ = c_Ԥ��.��¼״̬ And ��¼���� In (1, 11) And Rownum = 1;
    
      --���²���Ԥ�����
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) - n_��ǰ���
      Where ����id = v_����id And ���� = 1 And ���� = Nvl(c_Ԥ��.Ԥ�����, 2);
      --����Ƿ��Ѿ�������
      If c_Ԥ��.��� < n_Ԥ����� Then
        n_Ԥ����� := n_Ԥ����� - c_Ԥ��.���;
      Else
        n_Ԥ����� := 0;
      End If;
    
      If n_Ԥ����� = 0 Then
        Exit;
      End If;
    End Loop;
    If n_Ԥ����� <> 0 Then
      v_Err_Msg := '[ZLSOFT]Ԥ���಻��֧������֧�����,���ܼ���������[ZLSOFT]';
      Raise Err_Item;
    End If;
    Delete From ������� Where ����id = v_����id And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
    v_���ϼ� := n_Ԥ�����;
  End If;

  --ʣ�ಿ��ȫ����ȱʡ���㷽ʽ���㣬(С����Ҳ�����ж��⴦��)
  If v_���ϼ� <> 0 Then
    Update ����Ԥ����¼
    Set ��Ԥ�� = ��Ԥ�� + v_���ϼ�, �����id = �����id_In, ���㿨��� = ���㿨���_In, ���� = ����_In, ������ˮ�� = ������ˮ��_In, ����˵�� = ����˵��_In,
        ������λ = ������λ_In, ������� = n_�������
    
    Where ����id = ����id_In And Nvl(���㷽ʽ, 'LXH_Test') = Nvl(v_ȱʡ, 'LXH_Test') And ��¼���� = Decode(����_In, 1, 2, v_��¼����);
    If Sql%RowCount = 0 Then
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, �����id, ���㿨���, ����, ������ˮ��,
         ����˵��, ������λ, ��������)
      Values
        (����Ԥ����¼_Id.Nextval, Decode(����_In, 1, 2, v_��¼����), v_No, 1, v_����id, v_��ҳid, '���ս�������', v_ȱʡ, v_�Ǽ�ʱ��, v_����Ա���,
         v_����Ա����, v_���ϼ�, ����id_In, n_��id, n_�������, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, ������λ_In,
         Mod(Decode(����_In, 1, 2, v_��¼����), 10));
    End If;
  
    --�ҺŽ���,�ֱҴ���(���ڹҺŽ���û��Ԥ����,�����ڴ˹����и��ݷֱҴ������������)
    If v_��¼���� = 4 Then
    
      Begin
        Select a.��Ԥ��
        Into v_�ֽ���
        From ����Ԥ����¼ A, ���㷽ʽ B
        Where a.���㷽ʽ = b.���� And b.���� = 1 And a.����id = ����id_In;
      Exception
        When Others Then
          v_�ֽ��� := 0;
      End;
      If Floor(Abs(v_�ֽ���) * 10) <> Abs(v_�ֽ���) * 10 Then
        --����
        v_Cashcented := Zl_Cent_Money(v_�ֽ���);
        v_�����   := v_Cashcented - v_�ֽ���;
        If v_����� <> 0 Then
          If n_������� < 0 Then
            --10.34֮���������
            Begin
              Select ���� Into v_����� From ���㷽ʽ Where ���� = 9;
            Exception
              When Others Then
                v_Err_Msg := '������ȷ��ȡ��������Ϣ�����ȼ����㷽ʽ�������Ƿ�������ȷ��';
                Raise Err_Item;
            End;
            Insert Into ����Ԥ����¼
              (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, ��������)
            Values
              (����Ԥ����¼_Id.Nextval, Decode(����_In, 1, 2, v_��¼����), v_No, 1, v_����id, v_��ҳid, '����', v_�����, v_�Ǽ�ʱ��, v_����Ա���,
               v_����Ա����, v_�����, ����id_In, n_��id, n_�������, Mod(Decode(����_In, 1, 2, v_��¼����), 10));
          Else
            --1.����Ԥ����¼(һ�����ڼ�¼)
            Update ����Ԥ����¼
            Set ��Ԥ�� = v_Cashcented
            Where ���㷽ʽ = (Select ���� From ���㷽ʽ Where ���� = 1 And Rownum = 1) And ����id = ����id_In;
          
            --2.���������ü�¼(ע:���㵥λ��¼���Ǻű�,���Բ�ȡ������)
            Begin
              Select a.���, a.Id, c.Id, c.�վݷ�Ŀ
              Into v_�շ����, v_�շ�ϸĿid, v_������Ŀid, v_�վݷ�Ŀ
              From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C, �շ��ض���Ŀ D
              Where d.�ض���Ŀ = '�����' And d.�շ�ϸĿid = a.Id And a.Id = b.�շ�ϸĿid And b.������Ŀid = c.Id And
                    Sysdate Between b.ִ������ And Nvl(b.��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'));
            Exception
              When Others Then
                v_Err_Msg := '������ȷ��ȡ�շ���������Ϣ�����ȼ�����Ŀ�Ƿ�������ȷ��';
                Raise Err_Item;
            End;
            If Nvl(����_In, 0) = 1 Then
              Select ���˷��ü�¼_Id.Nextval Into v_����id From Dual;
              Select Max(���) + 1, Max(����ʱ��) Into v_���, v_����ʱ�� From סԺ���ü�¼ Where ����id = ����id_In;
              n_ҽ��С��id := Zl_ҽ��С��_Get(0, v_����Ա����, v_����id, v_��ҳid, v_����ʱ��);
            
              Insert Into סԺ���ü�¼
                (ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, �����־, ����id, ��ʶ��, ����, ����, �Ա�, ����, ���˲���id, ���˿���id, �ѱ�, �շ����,
                 �շ�ϸĿid, ���㵥λ, ��ҩ����, ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ���ʷ���, ������, ��������id, ������, ����ʱ��,
                 �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ����id, ���ʽ��, ����Ա���, ����Ա����, �Ƿ��ϴ�, �ɿ���id, ҽ��С��id)
                Select v_����id, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, v_���, Null, Null, �����־, ����id, ��ʶ��, ����, ����, �Ա�, ����, ���˲���id, ���˿���id,
                       �ѱ�, v_�շ����, v_�շ�ϸĿid, ���㵥λ, ��ҩ����, 1, 1, �Ӱ��־, 9, v_������Ŀid, v_�վݷ�Ŀ, v_�����, v_�����, v_�����, ���ʷ���,
                       ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ����id_In, v_�����, ����Ա���, ����Ա����, 1, �ɿ���id,
                       Decode(n_ҽ��С��id, Null, ҽ��С��id, n_ҽ��С��id)
                From סԺ���ü�¼
                Where ����id = ����id_In And Rownum = 1;
            Else
              Select ���˷��ü�¼_Id.Nextval Into v_����id From Dual;
              Select Max(���) + 1 Into v_��� From ������ü�¼ Where ����id = ����id_In;
              Insert Into ������ü�¼
                (ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, �����־, ����id, ��ʶ��, ���ʽ, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid,
                 ���㵥λ, ��ҩ����, ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ���ʷ���, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��,
                 ִ�в���id, ִ����, ִ��״̬, ����״̬, ����id, ���ʽ��, ����Ա���, ����Ա����, �Ƿ��ϴ�, �ɿ���id)
                Select v_����id, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, v_���, Null, Null, �����־, ����id, ��ʶ��, ���ʽ, ����, �Ա�, ����, ���˿���id, �ѱ�,
                       v_�շ����, v_�շ�ϸĿid, ���㵥λ, ��ҩ����, 1, 1, �Ӱ��־, 9, v_������Ŀid, v_�վݷ�Ŀ, v_�����, v_�����, v_�����, ���ʷ���, ������,
                       ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ����״̬, ����id_In, v_�����, ����Ա���, ����Ա����, 1, �ɿ���id
                From ������ü�¼
                Where ����id = ����id_In And Rownum = 1;
            End If;
          End If;
          --3.���»��ܱ�
          --ֻ���ܲ��������ı仯.��Ϊ�˱�������������α�
        End If;
      End If;
    End If;
  End If;

  --����ٴ���"��Ա�ɿ����"(û�ж���Ԥ���ǲ���,����"�������"��Ԥ�����ø���)
  For r_Del In c_Del Loop
    If r_Del.��¼���� Not In (1, 11) Then
      If Nvl(n_����״̬, 0) <> 1 Then
        Update ��Ա�ɿ����
        Set ��� = Nvl(���, 0) + r_Del.��Ԥ��
        Where �տ�Ա = v_����Ա���� And ���� = 1 And ���㷽ʽ = r_Del.���㷽ʽ;
        If Sql%RowCount = 0 Then
          Insert Into ��Ա�ɿ����
            (�տ�Ա, ���㷽ʽ, ����, ���)
          Values
            (v_����Ա����, r_Del.���㷽ʽ, 1, r_Del.��Ԥ��);
        End If;
      End If;
    End If;
  End Loop;
  Delete From ��Ա�ɿ���� Where ���� = 1 And �տ�Ա = v_����Ա���� And Nvl(���, 0) = 0;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˽����¼_Update;
/

--91385:Ƚ����,2015-12-29,�޸ġ��ǱҶ�����룬�������塱�����������.
Create Or Replace Function Zl_Cent_Money
(
  Money_In In Number,
  Type_In  In Number := 2
) Return Number As
  n_Sign Integer;
  n_Temp Number(16, 5);
  n_��� Number(16, 5);
  n_Mode Number(1);
Begin
  --         0.������
  --         1.��ȡ�������뷨,eg:0.51=0.50;0.56=0.60
  --         2.�����շ�,eg:0.51=0.60,0.56=0.60
  --         3.����շ�,eg:0.51=0.50,0.56=0.50
  --        4.�����������˫,eg:0.14=0.10,0.16=0.20,0.151=0.20,0.15=0.20,0.25=0.20
  --           �����������˫,����ҹ���ѧ����ίԱ����ʽ�䲼�ġ�������Լ����,������vb��Round����,�������������ְ�����λ����ʱ�����Ը����ֽ���������Լ 
  --           �����м����뷨:���������忼�ǣ�������ͽ�һ�������㿴��ż����ǰΪżӦ��ȥ����ǰΪ��Ҫ��һ
  --         5.�������塢�������,�Խǽ��д�������Ҫ�ȶԷֱҽ�������,��0.24(��)���¶�����ǣ�0.75(��)���϶����ǣ�0.25-0.74����Ϊ0.5��
  --         6.��������:eg:0.15=0.10:0.16=0.2:   ���˺� ����:34519  ����:2010-12-06 09:58:02

  n_Mode := To_Number(Substr(Nvl(zl_GetSysParameter(14) || '000', '000'), Type_In, 1));
  n_Sign := Sign(Money_In);
  n_��� := Abs(Money_In);
  If n_Mode = 1 Then
    --1.�������뷨,eg:0.51=0.50;0.56=0.60
    n_Temp := n_Sign * Round(n_���, 1);
    Return n_Temp;
  End If;
  If n_Mode = 2 Then
    ----2.�����շ�,eg:0.51=0.60,0.56=0.60
    n_Temp := n_Sign * Ceil(n_��� * 10) / 10;
    Return n_Temp;
  End If;
  If n_Mode = 3 Then
    ----3.����շ�,eg:0.51=0.50,0.56=0.50
    n_Temp := n_Sign * Floor(n_��� * 10) / 10;
    Return n_Temp;
  End If;
  If n_Mode = 4 Then
    ----4.�����������˫,����Oracleû����غ���,�㷨����,�ݲ�֧��
    n_Temp := n_Sign * n_���;
    Return n_Temp;
  End If;
  If n_Mode = 5 Then
    ----5.�������塢�������,eg:0.29=0,0.30=0.50,0.79=0.50,0.80=1.00
    n_Temp := Round(n_��� - Floor(n_���), 1);
    If n_Temp >= 0.8 Then
      n_Temp := 1;
    Elsif n_Temp < 0.3 Then
      n_Temp := 0;
    Else
      n_Temp := 0.5;
    End If;
    n_Temp := Floor(n_���) + n_Temp; --5.�������塢�������,eg:0.24=0,0.25=0.50,0.74=0.50,0.75=1.00
    n_Temp := n_Sign * n_Temp;
    Return n_Temp;
  End If;
  If n_Mode = 6 Then
    ----6.��������
    n_Temp := n_Sign * Round(n_��� - 0.01, 1);
    Return n_Temp;
  End If;
  Return Money_In;
Exception
  When Others Then
    Return Null;
End Zl_Cent_Money;
/

--91665:Ƚ����,2015-12-29,���Ӷ൥�ݷֵ��ݽ���ʱҽ������ʧ��ʱֻ�Խ���ɹ������շѵ�ģʽ��
Create Or Replace Procedure Zl_�����շ�Ʊ��_Insert
(
  No_In           Varchar2,
  Ʊ�ݺ�_In       Ʊ��ʹ����ϸ.����%Type,
  ����id_In       Ʊ��ʹ����ϸ.����id%Type,
  ʹ����_In       Ʊ��ʹ����ϸ.ʹ����%Type,
  ʹ��ʱ��_In     Ʊ��ʹ����ϸ.ʹ��ʱ��%Type,
  ��ӡid_In       Ʊ�ݴ�ӡ����.Id%Type := 0,
  Ʊ������_In     Number := 1,
  ҽ���ӿڴ�ӡ_In Number := 0
) As
  --���ܣ����������շ�Ʊ�ݵķ���
  --������
  --      NO_IN       =     �շѵĵ��ݺ�,�����Ƕ��ŵ���ͬʱ�շѡ���ʽΪ��A0000001,A0000002,....
  --      Ʊ�ݺ�_IN   =     Ҫʹ�õĿ�ʼƱ�ݺš���Ʊ�ݺ�Ӧ�ò�Ϊ�գ������ô���Ʊ�ݣ�Ҳ�������ֶ���һ���շѵĵ��ݡ�
  --      ����ID_IN   =     �ϸ����Ʊ��ʱ��Ϊʹ��Ʊ�ݵ��������Ρ����ϸ����ʱ��ΪNULL��
  --      ��ӡID_IN   =     ���޸Ķ൥���е�һ��ʱ,Ϊ�˱��������ش�,���õ��ݵĴ�ӡ������дΪ���˷ѵ�����ͬ,�������·���Ʊ��,���˷��ش򷢳�
  --      Ʊ������_In =     ʵ�������Ʊ�ݴ�ӡ����
  --      ҽ���ӿڴ�ӡ_In = �Ƿ�ҽ���ӿڴ�ӡ�ȴ���Ʊ�����ݣ����ǽ������ӡid_In

  --���α�����Ʊ�ݷ�Χ�ж�
  Cursor c_Fact Is
    Select * From Ʊ�����ü�¼ Where ID = Nvl(����id_In, 0);
  r_Factrow c_Fact%RowType;

  v_Ʊ�ݺ�     Ʊ��ʹ����ϸ.����%Type;
  v_��ǰƱ�ݺ� Ʊ��ʹ����ϸ.����%Type;
  v_��ӡid     Ʊ�ݴ�ӡ����.Id%Type;

  v_��ǰ�� ������ü�¼.No%Type;
  v_���ݺ� Varchar2(1000);

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  If Nvl(��ӡid_In, 0) = 0 Or Nvl(ҽ���ӿڴ�ӡ_In, 0) = 1 Then
    --��Ʊ�ݺ�ʱ,���ô���Ʊ��
    If Ʊ�ݺ�_In Is Null Then
      Return;
    End If;
    v_��ӡid := Nvl(��ӡid_In, 0);
    If Nvl(v_��ӡid, 0) = 0 Then
      Select Ʊ�ݴ�ӡ����_Id.Nextval Into v_��ӡid From Dual;
    End If;
  
    --���ɵ��ݵ�Ʊ�ݴ�ӡ����
    v_���ݺ� := No_In || ',';
    While v_���ݺ� Is Not Null Loop
      v_��ǰ�� := Substr(v_���ݺ�, 1, Instr(v_���ݺ�, ',') - 1);
      --Ʊ�ݴ�ӡ����
      Insert Into Ʊ�ݴ�ӡ���� (ID, ��������, NO) Values (v_��ӡid, 1, v_��ǰ��);
      --������ü�¼����д��ʼƱ�ݺ��Ա���ʾ
      Update ������ü�¼ Set ʵ��Ʊ�� = Ʊ�ݺ�_In Where ��¼���� = 1 And NO = v_��ǰ��;
      v_���ݺ� := Substr(v_���ݺ�, Instr(v_���ݺ�, ',') + 1);
    End Loop;
  
    --������Ʊ��
    v_Ʊ�ݺ� := Ʊ�ݺ�_In;
    If Nvl(����id_In, 0) <> 0 Then
      Open c_Fact;
      Fetch c_Fact
        Into r_Factrow;
      If c_Fact%RowCount = 0 Then
        v_Error := '��Ч��Ʊ���������Σ��޷�����շ�Ʊ�ݷ��������';
        Close c_Fact;
        Raise Err_Custom;
      Elsif Nvl(r_Factrow.ʣ������, 0) < Ʊ������_In Then
        v_Error := '��ǰ���ε�ʣ����������' || Ʊ������_In || '�ţ��޷�����շ�Ʊ�ݷ��������';
        Close c_Fact;
        Raise Err_Custom;
      End If;
    End If;
    For I In 1 .. Ʊ������_In Loop
      --���Ʊ�ݷ�Χ�Ƿ���ȷ
      If Nvl(����id_In, 0) <> 0 Then
        If Not (Upper(v_Ʊ�ݺ�) >= Upper(r_Factrow.��ʼ����) And Upper(v_Ʊ�ݺ�) <= Upper(r_Factrow.��ֹ����) And
            Length(v_Ʊ�ݺ�) = Length(r_Factrow.��ֹ����)) Then
          v_Error := '�õ�����Ҫ��ӡ����Ʊ��,��Ʊ�ݺ�"' || v_Ʊ�ݺ� || '"����Ʊ�����õĺ��뷶Χ��';
          Close c_Fact;
          Raise Err_Custom;
        End If;
      End If;
    
      --����Ʊ��
      Insert Into Ʊ��ʹ����ϸ
        (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ����, ʹ��ʱ��)
      Values
        (Ʊ��ʹ����ϸ_Id.Nextval, 1, v_Ʊ�ݺ�, 1, 1, ����id_In, v_��ӡid, ʹ����_In, ʹ��ʱ��_In);
    
      v_��ǰƱ�ݺ� := v_Ʊ�ݺ�;
      --��һ��Ʊ�ݺ�
      v_Ʊ�ݺ� := Zl_Incstr(v_Ʊ�ݺ�);
    End Loop;
  
    If Nvl(����id_In, 0) <> 0 Then
      Update Ʊ�����ü�¼
      Set ʹ��ʱ�� = ʹ��ʱ��_In, ��ǰ���� = v_��ǰƱ�ݺ�, ʣ������ = Nvl(ʣ������, 0) - Ʊ������_In
      Where ID = ����id_In;
    
      Close c_Fact;
    End If;
  Else
    Insert Into Ʊ�ݴ�ӡ���� (ID, ��������, NO) Values (��ӡid_In, 1, No_In);
    If Ʊ�ݺ�_In Is Null Then
      Return;
    End If;
    --������ü�¼����д��ʼƱ�ݺ��Ա���ʾ
    Update ������ü�¼ Set ʵ��Ʊ�� = Ʊ�ݺ�_In Where ��¼���� = 1 And NO = No_In;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����շ�Ʊ��_Insert;
/

--92335:���ϴ�,2016-01-18,����֧����ģʽ�����̲��
--91665:Ƚ����,2015-12-29,���Ӷ൥�ݷֵ��ݽ���ʱҽ������ʧ��ʱֻ�Խ���ɹ������շѵ�ģʽ��
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
    If n_Havenull = 0 Or Round(Nvl(r_Feedata.������, 0), 6) <> Round(Nvl(n_������, 0), 6) Then
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

--92177:������,2015-12-29,ҽ��У��ʱ��ҳID����
Create Or Replace Procedure Zl_סԺ�շѽ���_Update
(
  ����id_In       סԺ���ü�¼.����id%Type,
  ���ʽ���_In     Varchar2, --���ʽ���_IN-��ҽ��ʱ:���㷽ʽ|������|�������||.....ҽ��ʱ:���㷽ʽ|������|�������,��������,�����ʺ�||.....
  ��Ԥ��_In       Varchar2, --��Ԥ��_IN= ID|���ݺ�|���|��¼״̬||.....
  �ɿ�_In         ����Ԥ����¼.�ɿ�%Type := Null,
  �Ҳ�_In         ����Ԥ����¼.�Ҳ�%Type := Null,
  �����ʻ�����_In Varchar2 := Null --:���㷽ʽ|������|�����ID|����|������ˮ��|����˵��||...
) As
  --����:�������ʱ��ҽ����ʽ�����,��ؽ�����Ϣ�ĵ���
  --     ��Ϊ������ʺ�,���ɵ�ҽ���������ܶ��̯���ܻ�����ʽ����ʱ�в���,�����ṩ��У�Թ���,
  --   ����Ա�ڽ���У��ʱ,���Ե�����ҽ�����㷽ʽ�ĸ��ֽ������ʽ,�������ɽ��㴮,���ҿ��ܲ��������.

  --������Ϣ
  Cursor c_Pati(v_����id ������Ϣ.����id%Type) Is
    Select a.����, a.�Ա�, a.����, a.סԺ��, a.�����, b.��ҳid, b.��Ժ����, b.��ǰ����id, b.��Ժ����id, Nvl(b.�ѱ�, a.�ѱ�) As �ѱ�, c.���� As ���ʽ
    From ������Ϣ A, ������ҳ B, ҽ�Ƹ��ʽ C, (Select Max(��ҳid) As ��ҳid From סԺ���ü�¼ Where ����id = ����id_In) D
    Where a.����id = v_����id And a.����id = b.����id(+) And b.��ҳid = Nvl(d.��ҳid, 0) And a.ҽ�Ƹ��ʽ = c.����(+);
  r_Pati c_Pati%RowType;

  --���̱���
  v_�������� Varchar2(4000);
  v_��ǰ���� Varchar2(100);
  v_���㷽ʽ ����Ԥ����¼.���㷽ʽ%Type;
  n_������ ����Ԥ����¼.��Ԥ��%Type;
  v_������� Varchar2(100); --���ս����¼ʱ,����:�������,��������,�����ʺ�

  n_�����id   ����Ԥ����¼.�����id%Type;
  v_����       ����Ԥ����¼.����%Type;
  v_������ˮ�� ����Ԥ����¼.������ˮ��%Type;
  v_����˵��   ����Ԥ����¼.����˵��%Type;

  v_No         ����Ԥ����¼.No%Type;
  v_����id     ����Ԥ����¼.����id%Type;
  v_�տ�ʱ��   ����Ԥ����¼.�տ�ʱ��%Type;
  v_����Ա��� ����Ԥ����¼.����Ա���%Type;
  v_����Ա���� ����Ԥ����¼.����Ա����%Type;
  v_������   ����Ԥ����¼.���㷽ʽ%Type;

  v_Ԥ��id   ����Ԥ����¼.Id%Type;
  v_��¼״̬ ����Ԥ����¼.��¼״̬%Type;

  v_������� ����Ԥ����¼.�ɿλ%Type;
  v_�����ʺ� ����Ԥ����¼.��λ������%Type;
  v_�������� ����Ԥ����¼.��λ�ʺ�%Type;
  v_���ʽ ������ü�¼.���ʽ%Type;
  n_����ֵ   �������.Ԥ�����%Type;
  n_Dele     Number; --0-��ɾ��,1-ɾ��
  v_Error    Varchar2(255);
  Err_Custom Exception;
  n_�����־   Number; --1.����,2-סԺ,3-�����סԺ
  n_��id       ����ɿ����.Id%Type;
  n_���       Number;
  n_�����Ԥ�� ����Ԥ����¼.��Ԥ��%Type;
  n_סԺ��Ԥ�� ����Ԥ����¼.��Ԥ��%Type;

Begin

  --1.ȡԤ����¼�е���Ҫ�������Ϣ
  Select NO, ����id, �շ�ʱ��, ����Ա���, ����Ա����, �ɿ���id
  Into v_No, v_����id, v_�տ�ʱ��, v_����Ա���, v_����Ա����, n_��id
  From ���˽��ʼ�¼
  Where ID = ����id_In;

  Open c_Pati(v_����id);
  Fetch c_Pati
    Into r_Pati;

  --��������Ϣ
  Begin
    Select ���� Into v_������ From ���㷽ʽ Where ���� = 9;
  Exception
    When Others Then
      Begin
        v_Error := '������ȷ��ȡ�շ���������Ϣ�����ȼ�����Ŀ�Ƿ�������ȷ��';
        Raise Err_Custom;
      End;
  End;

  --2.ɾ���ɵļ�¼,���˻�������
  --������Ա�ɿ����,�������,
  For c_Del In (Select ���㷽ʽ, ����Ա����, ��Ԥ�� From ����Ԥ����¼ Where ����id = ����id_In And ��¼���� = 2) Loop
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) - Nvl(c_Del.��Ԥ��, 0)
    Where ���㷽ʽ = c_Del.���㷽ʽ And �տ�Ա = v_����Ա���� And ���� = 1;
    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ����
        (�տ�Ա, ���㷽ʽ, ����, ���)
      Values
        (v_����Ա����, c_Del.���㷽ʽ, 1, -1 * c_Del.��Ԥ��);
    End If;
  End Loop;

  If v_����id > 0 Then
    For v_Ԥ�� In (Select Ԥ�����, Sum(Nvl(��Ԥ��, 0)) As Ԥ�����
                 From ����Ԥ����¼
                 Where ����id = ����id_In And ��¼���� In (1, 11)
                 Group By Ԥ�����
                 Having Sum(Nvl(��Ԥ��, 0)) <> 0) Loop
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(v_Ԥ��.Ԥ�����, 0)
      Where ����id = v_����id And ���� = Nvl(v_Ԥ��.Ԥ�����, 2) And ���� = 1;
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, ����, Ԥ�����, ����)
        Values
          (v_����id, Nvl(v_Ԥ��.Ԥ�����, 2), Nvl(v_Ԥ��.Ԥ�����, 0), 1);
      End If;
    End Loop;
  End If;

  --ȡ�����־
  Select Case
           When �����־ = 1 And סԺ��־ = 1 Then
            3
           When �����־ = 1 Then
            1
           Else
            2
         End, ���ʽ
  Into n_�����־, v_���ʽ
  From (Select Nvl(Max(�����־), 0) As �����־, Nvl(Max(סԺ��־), 0) As סԺ��־, Max(���ʽ) As ���ʽ
         From (Select 1 As �����־, 0 As סԺ��־, ���ʽ
                From ������ü�¼
                Where ����id = ����id_In And Rownum = 1
                Union All
                Select 0 As �����־, 1 As סԺ��־, '' As ���ʽ
                From סԺ���ü�¼
                Where ����id = ����id_In And Rownum = 1));

  --���˻��ܱ�.         ����δ�����(��Ϊ������������,���Բ�����)
  --ֻ���ܲ��������ı仯. �����ֻ���ܴ���һ��,��Ϊ�˱�������������α�

  --ɾ�����ʽɿ�,���ս����¼
  Delete �������㽻�� Where ����id In (Select ID From ����Ԥ����¼ Where ����id = ����id_In And ��¼���� = 2);

  Delete ����Ԥ����¼ Where ����id = ����id_In And ��¼���� = 2;

  --��һ�γ�Ԥ����,��ճ����
  Update ����Ԥ����¼ Set ��Ԥ�� = Null, ����id = Null, �������� = Null Where ����id = ����id_In And ��¼���� = 1;
  --ɾ�������
  Delete ����Ԥ����¼ Where ����id = ����id_In And ��¼���� = 11;

  --ɾ������¼
  If n_�����־ = 1 Then
    Delete ������ü�¼ Where ����id = ����id_In And ���ӱ�־ = 9;
  Else
    Delete סԺ���ü�¼ Where ����id = ����id_In And ���ӱ�־ = 9;
  End If;

  --4.�������ɲ���Ԥ����¼�������
  --4.1.�������,���ս���
  If ���ʽ���_In Is Not Null Then
    v_�������� := ���ʽ���_In || '||';
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_������� := LTrim(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
      If Instr(v_�������, ',') > 0 Then
        --ҽ������:�������,��������,�����ʺ�
        v_������� := v_������� || ',';
        v_������� := Substr(v_�������, 1, Instr(v_�������, ',') - 1);
        v_������� := Substr(v_�������, Instr(v_�������, ',') + 1);
        v_�������� := Substr(v_�������, 1, Instr(v_�������, ',') - 1);
        v_������� := Substr(v_�������, Instr(v_�������, ',') + 1);
        v_�����ʺ� := Substr(v_�������, 1, Instr(v_�������, ',') - 1);
        v_������� := Null;
      Else
        v_������� := Null;
        v_�������� := Null;
        v_�����ʺ� := Null;
      End If;
    
      If Nvl(n_������, 0) <> 0 Then
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��,
           ����id, �ɿ�, �Ҳ�, �ɿ���id, ��������)
        Values
          (����Ԥ����¼_Id.Nextval, v_No, Null, 2, 1, v_����id, r_Pati.��ҳid, r_Pati.��Ժ����id, Null, v_���㷽ʽ, v_�������, '���ʽɿ�',
           v_�������, v_��������, v_�����ʺ�, v_�տ�ʱ��, v_����Ա���, v_����Ա����, n_������, ����id_In,
           Decode(v_��������, ���ʽ���_In || '||', �ɿ�_In, Null), Decode(v_��������, ���ʽ���_In || '||', �Ҳ�_In, Null), n_��id, 2);
      End If;
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;

  If �����ʻ�����_In Is Not Null Then
    --���㷽ʽ|������|�����ID|����|������ˮ��|����˵��||...
    v_�������� := �����ʻ�����_In || '||';
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
    
      v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
    
      n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
    
      n_�����id := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
    
      v_����     := LTrim(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
    
      v_������ˮ�� := LTrim(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_��ǰ����   := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      v_����˵��   := LTrim(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
    
      If Nvl(n_������, 0) <> 0 Then
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��,
           ����id, �ɿ�, �Ҳ�, �ɿ���id, ��������, �����id, ����, ������ˮ��, ����˵��)
        Values
          (����Ԥ����¼_Id.Nextval, v_No, Null, 2, 1, v_����id, r_Pati.��ҳid, r_Pati.��Ժ����id, Null, v_���㷽ʽ, v_�������, '���ʽɿ�', Null,
           Null, Null, v_�տ�ʱ��, v_����Ա���, v_����Ա����, n_������, ����id_In, Null, Null, n_��id, 2, n_�����id, v_����, v_������ˮ��, v_����˵��);
      End If;
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  End If;

  --4.2.Ԥ������
  If ��Ԥ��_In Is Not Null Then
    v_��������   := ��Ԥ��_In || '||';
    n_�����Ԥ�� := 0;
    n_סԺ��Ԥ�� := 0;
    While v_�������� Is Not Null Loop
      v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
      v_Ԥ��id   := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1)); --�Ǽ�¼��Ԥ����ID
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      v_������� := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1); --�Ǽ�¼��Ԥ����NO��
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1);
      n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1));
      v_��¼״̬ := To_Number(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
    
      If v_Ԥ��id <> 0 Then
        --��һ�γ�Ԥ��(����һ�α��Ͻ���ID,��Ԥ�����Ϊ0)
        Update ����Ԥ����¼ Set ��Ԥ�� = 0, ����id = ����id_In, �������� = 2 Where ID = v_Ԥ��id;
      End If;
      --���ϴ�ʣ���
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
         ����id, �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������)
        Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, v_��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�,
               v_�տ�ʱ��, v_����Ա����, v_����Ա���, n_������, ����id_In, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, 2
        From ����Ԥ����¼
        Where NO = v_������� And ��¼���� In (1, 11) And ��¼״̬ = v_��¼״̬ And Rownum = 1;
    
      Begin
        Select Nvl(Ԥ�����, 2)
        Into n_���
        From ����Ԥ����¼
        Where NO = v_������� And ��¼���� In (1, 11) And ��¼״̬ = v_��¼״̬ And Rownum = 1;
      Exception
        When Others Then
          n_��� := 2;
      End;
      If Nvl(n_���, 0) = 1 Then
        n_�����Ԥ�� := n_�����Ԥ�� + Nvl(n_������, 0);
      Else
        n_סԺ��Ԥ�� := n_סԺ��Ԥ�� + Nvl(n_������, 0);
      End If;
      v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
    End Loop;
  
    --���²������
    If n_�����Ԥ�� <> 0 Then
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) - n_�����Ԥ��
      Where ����id = v_����id And ���� = 1 And ���� = 1
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ������� (����id, ����, Ԥ�����, ����) Values (v_����id, 1, -1 * n_�����Ԥ��, 1);
        n_����ֵ := -1 * n_�����Ԥ��;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From ������� Where ����id = v_����id And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
      End If;
    End If;
    If n_סԺ��Ԥ�� <> 0 Then
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) - n_סԺ��Ԥ��
      Where ����id = v_����id And ���� = 1 And ���� = 2
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ������� (����id, ����, Ԥ�����, ����) Values (v_����id, 2, -1 * n_סԺ��Ԥ��, 1);
        n_����ֵ := -1 * n_סԺ��Ԥ��;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From ������� Where ����id = v_����id And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
      End If;
    End If;
  
  End If;

  --5.��ػ��ܱ�Ĵ���
  --������Ա�ɿ����
  --�ɿ����,���ս���
  n_Dele := 0;
  For c_���� In (Select ���㷽ʽ, ��Ԥ�� From ����Ԥ����¼ Where ����id = ����id_In And ��¼���� = 2) Loop
  
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) + Nvl(c_����.��Ԥ��, 0)
    Where �տ�Ա = v_����Ա���� And ���� = 1 And ���㷽ʽ = c_����.���㷽ʽ
    Returning ��� Into n_����ֵ;
    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ����
        (�տ�Ա, ���㷽ʽ, ����, ���)
      Values
        (v_����Ա����, c_����.���㷽ʽ, 1, Nvl(c_����.��Ԥ��, 0));
      n_����ֵ := Nvl(Nvl(c_����.��Ԥ��, 0), 0);
    End If;
  
    If Nvl(n_����ֵ, 0) = 0 Then
      n_Dele := 1;
    End If;
  
  End Loop;

  If n_Dele = 1 Then
    Delete From ��Ա�ɿ���� Where ���� = 1 And �տ�Ա = v_����Ա���� And Nvl(���, 0) = 0;
  End If;
  --���ܱ�,ֻ���ػ������,��Ϊ��������,δ����ò���(�²�����������ѽ���),ֻ��һ������¼,��Ϊʹ�ñ�����������α�

  --6.ҽ����ر�Ĵ���
  --Delete ҽ���˶Ա� Where ����Id=����Id_IN;
  Update ���ս�����ϸ Set ��־ = 2 Where ����id = ����id_In;

  Close c_Pati;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_סԺ�շѽ���_Update;
/


--91903:������,2015-12-28,������˵�δ������΢����ҽ����ִ��״̬��Ϊ��ִ��
CREATE OR REPLACE Procedure Zl_����걾��¼_�������
(
  Id_In       ����걾��¼.Id%Type,
  �����_In   ����걾��¼.�����%Type := Null,
  ��Ա���_In ��Ա��.���%Type := Null,
  ��Ա����_In ��Ա��.����%Type := Null
) Is

  --δ��˵ķ�����(������ҩƷ)
  Cursor c_Verify(v_ҽ��id In Number) Is
    Select Distinct ��¼����, NO, ���
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
    Select Distinct ��¼����, NO, ���
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

  v_ִ�� Number(1);
  v_No   ����ҽ������.No%Type;
  v_���� ����ҽ������.��¼����%Type;
  v_��� Varchar2(1000);

  v_Count Number(18);

  v_΢����걾 Number(1) := 0;
  v_��ҳid     Number(18);
  v_Ӥ��       Number(1);
  v_����       Varchar2(100);
  v_����       Number(18);
  v_Intloop    Number;
  Err_Custom Exception;
  v_Error Varchar2(100);
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
      Set ִ��״̬ = 1, ����� = ��Ա����_In, ���ʱ�� = Sysdate
      Where ҽ��id In (Select ID From ����ҽ����¼ Where r_Samplequest.ҽ��id In (ID, ���id));
      
      update ����ҽ������
      Set ִ��״̬ = 1, ����� = ��Ա����_In, ���ʱ�� = Sysdate
      Where ҽ��id In (select ���ID from ����ҽ����¼ where ID in(Select ID From ����ҽ����¼ Where r_Samplequest.ҽ��id In (ID, ���id)));

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
        If v_��� Is Not Null Then
          If v_���� = 1 Then
            Zl_������ʼ�¼_Verify(v_No, ��Ա���_In, ��Ա����_In, Substr(v_���, 2));
          Elsif v_���� = 2 Then
            Zl_סԺ���ʼ�¼_Verify(v_No, ��Ա���_In, ��Ա����_In, Substr(v_���, 2));
          End If;
          v_��� := Null;
        End If;
      End If;

      --����Լ����ĵ�
      v_Intloop := 1;
      v_No      := Null;
      Select ����id Into v_���� From ����걾��¼ Where ID = Id_In;
      For r_�����Լ� In (Select c.����id, c.����
                     From ����ҽ����¼ A, ���鱨����Ŀ B, �����Լ���ϵ C
                     Where a.���id = r_Samplequest.ҽ��id And a.������Ŀid = b.������Ŀid And b.������Ŀid = c.��Ŀid And c.����id = v_����) Loop
        Zl_�����Լ���¼_Insert(r_Samplequest.ҽ��id, v_Intloop, r_�����Լ�.����id, r_�����Լ�.����);
        v_Intloop := v_Intloop + 1;
      End Loop;
      Select Count(*) Into v_Intloop From �����Լ���¼ Where ҽ��id = r_Samplequest.ҽ��id And NO Is Null;
      If v_Intloop > 1 Then
        v_No := Nextno(14);
        Update �����Լ���¼ Set NO = v_No Where ҽ��id = r_Samplequest.ҽ��id;
      End If;
      If v_No Is Not Null Then

        Zl_�����Լ���¼_Bill(r_Samplequest.ҽ��id, v_No);

        v_��ҳid := Null;
        Select ��ҳid Into v_��ҳid From ����ҽ����¼ A Where ID = r_Samplequest.ҽ��id;

        If v_��ҳid Is Null Then
          Zl_������ʼ�¼_Verify(v_No, ��Ա���_In, ��Ա����_In);
        Else
          Zl_סԺ���ʼ�¼_Verify(v_No, ��Ա���_In, ��Ա����_In);
        End If;

        --�������û���Զ�����,���Զ�����,���򲻴���
        For r_Stuff In c_Stuff(v_No, v_��ҳid) Loop
          Zl_�����շ���¼_��������(r_Stuff.�ⷿid, 25, v_No, ��Ա����_In, ��Ա����_In, ��Ա����_In, 1, Sysdate);
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

--91225:������,2015-12-28,�ڼ����걨��¼���洦��״̬��1�ſɱ��ͣ������ڴ�Ⱦ������ϵͳ���洦��״̬��3���ܱ���
Create Or Replace Procedure Zl_�����걨��¼_Send
(
  �ļ�id_In   In Varchar2,
  ������_In   In �����걨��¼.������%Type,
  ����ʱ��_In In �����걨��¼.����ʱ��%Type,
  ���͵�λ_In In �����걨��¼.���͵�λ%Type,
  ���ͱ�ע_In In �����걨��¼.���ͱ�ע%Type
) Is
  v_����   ��Ա��.����%Type;
  n_�ļ�id Number;
  e_Changed Exception;
Begin

  If Length(�ļ�id_In) <> 32 Then
    n_�ļ�id := To_Number(�ļ�id_In); --�²���ID��32λGUID
  End If;

  Select ���� Into v_���� From ��Ա�� P, �ϻ���Ա�� U Where p.Id = u.��Աid And u.�û��� = User And Rownum < 2;
  If Length(�ļ�id_In) <> 32 Then
    --���û�й鵵������鵵 
    Update ���Ӳ�����¼ Set �鵵�� = v_����, �鵵���� = Sysdate Where ID = �ļ�id_In And �鵵�� Is Null;
  End If;

  Update �����걨��¼
  Set ����״̬ = 2, ������ = ������_In, ����ʱ�� = ����ʱ��_In, ���͵�λ = ���͵�λ_In, ���ͱ�ע = ���ͱ�ע_In, �Ǽ��� = v_����, �Ǽ�ʱ�� = Sysdate
  Where Decode(Length(�ļ�id_In), 32, �ĵ�id, �ļ�id) = Decode(Length(�ļ�id_In), 32, �ļ�id_In, n_�ļ�id) And (����״̬ = 1 or ����״̬ = 3);
  If Sql%RowCount = 0 Then
    Raise e_Changed;
  End If;
Exception
  When No_Data_Found Then
    Raise_Application_Error(-20101, '[ZLSOFT]�û���ݲ���ȷ��[ZLSOFT]');
  When e_Changed Then
    Raise_Application_Error(-20101, '[ZLSOFT]���������Ѿ��������û��ı䣡[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����걨��¼_Send;
/

--91780:������,2015-12-28,��ͬ��Ʊ������
Create Or Replace Procedure Zl_�շ�Ա����Ʊ��_Insert
(
  �ս�id_In   In ��Ա�ս�Ʊ��.�ս�id%Type,
  Ʊ����Ϣ_In Varchar2
) Is
  --------------------------------------------------------------------------------------------------------------------
  --����:�շ�Ա������ϸд��
  --����:������Ϣ_IN:Ʊ��,����,���,Ʊ������,��ʼƱ��,��ֹƱ��,���,����ʱ��|Ʊ��,����,���,Ʊ������,��ʼƱ��,��ֹƱ��,���,����ʱ��|...
  --                 Ʊ��:1-�շ��վ�,2-Ԥ���վ�,3-�����վ�,4-�Һ��վ�,5-���￨
  --         ����:1-����Ʊ��;2-�˷��ջ�Ʊ��;3-�ش��ջ�Ʊ��
  --                 ����ʱ��:yyyy-mm-dd hh24:mi:ss
  --
  --------------------------------------------------------------------------------------------------------------------
  v_�������� Varchar2(4000);
  v_��ǰ���� Varchar2(500);

  n_Ʊ��     ��Ա�ս�Ʊ��.Ʊ��%Type;
  n_����     ��Ա�ս�Ʊ��.����%Type;
  n_���     ��Ա�ս�Ʊ��.���%Type;
  n_Ʊ������ ��Ա�ս�Ʊ��.Ʊ������%Type;
  v_��ʼƱ�� ��Ա�ս�Ʊ��.��ʼƱ��%Type;
  v_��ֹƱ�� ��Ա�ս�Ʊ��.��ֹƱ��%Type;
  n_���     ��Ա�ս�Ʊ��.���%Type;
  v_����ʱ�� Varchar2(20);
  v_����     ��Ա�ս�Ʊ��.����%Type;

  t_��ʼƱ�� t_Strlist := t_Strlist();
  t_��ֹƱ�� t_Strlist := t_Strlist();
  t_����ʱ�� t_Strlist := t_Strlist();
  t_����     t_Strlist := t_Strlist();
  t_Ʊ��     t_Numlist := t_Numlist();
  t_����     t_Numlist := t_Numlist();
  t_���     t_Numlist := t_Numlist();
  t_���     t_Numlist := t_Numlist();
  t_Ʊ������ t_Numlist := t_Numlist();
Begin

  v_�������� := Ʊ����Ϣ_In || '|'; --�Կո�ֿ���|��β,û�н�������
  While v_�������� Is Not Null Loop
  
    v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '|') - 1);
    n_Ʊ��     := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1));
    t_Ʊ��.Extend;
    t_Ʊ��(t_Ʊ��.Count) := n_Ʊ��;
  
    v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
    n_����     := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1));
    t_����.Extend;
    t_����(t_����.Count) := n_����;
  
    v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
    n_���     := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1));
    t_���.Extend;
    t_���(t_���.Count) := n_���;
  
    v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
    n_Ʊ������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1));
    t_Ʊ������.Extend;
    t_Ʊ������(t_Ʊ������.Count) := n_Ʊ������;
  
    v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
    v_��ʼƱ�� := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
    t_��ʼƱ��.Extend;
    t_��ʼƱ��(t_��ʼƱ��.Count) := v_��ʼƱ��;
  
    v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
    v_��ֹƱ�� := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
    t_��ֹƱ��.Extend;
    t_��ֹƱ��(t_��ֹƱ��.Count) := v_��ֹƱ��;
  
    v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
    n_���     := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1));
    t_���.Extend;
    t_���(t_���.Count) := n_���;
  
    v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
    v_����ʱ�� := LTrim(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1));
    t_����ʱ��.Extend;
    If Nvl(v_����ʱ��, '-') = '-' Then
      t_����ʱ��(t_����ʱ��.Count) := Null;
    Else
      t_����ʱ��(t_����ʱ��.Count) := v_����ʱ��;
    End If;
  
    v_���� := LTrim(Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1));
    t_����.Extend;
    If Nvl(v_����, '-') = '-' Then
      t_����(t_����.Count) := Null;
    Else
      t_����(t_����.Count) := v_����;
    End If;
  
    v_�������� := Substr(v_��������, Instr(v_��������, '|') + 1);
  End Loop;
  --������������
  Forall I In 1 .. t_Ʊ��.Count
    Insert Into ��Ա�ս�Ʊ��
      (�ս�id, Ʊ��, ����, ���, Ʊ������, ��ʼƱ��, ��ֹƱ��, ���, ����ʱ��, ����)
    Values
      (�ս�id_In, t_Ʊ��(I), t_����(I), t_���(I), t_Ʊ������(I), t_��ʼƱ��(I), t_��ֹƱ��(I), t_���(I),
       To_Date(t_����ʱ��(I), 'yyyy-mm-dd hh24:mi:ss'), t_����(I));
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�շ�Ա����Ʊ��_Insert;
/

--92454:������,2016-01-07,�����Һ��Ŷ�����
--92007:������,2015-12-25,����ԤԼ��дԤԼ����Ա
Create Or Replace Procedure Zl_���������Һ�_Insert
(
  ������ʽ_In     Integer,
  ����id_In       ������ü�¼.����id%Type,
  ����_In         �ҺŰ���.����%Type,
  ����_In         �Һ����״̬.���%Type,
  ���ݺ�_In       ������ü�¼.No%Type,
  Ʊ�ݺ�_In       ������ü�¼.ʵ��Ʊ��%Type,
  ���㷽ʽ_In     ����Ԥ����¼.���㷽ʽ%Type, --�ֽ�Ľ�������
  ժҪ_In         ������ü�¼.ժҪ%Type, --ԤԼ�Һ�ժҪ��Ϣ
  ����ʱ��_In     ������ü�¼.����ʱ��%Type,
  �Ǽ�ʱ��_In     ������ü�¼.�Ǽ�ʱ��%Type,
  ������λ_In     �Һź�����λ.����%Type,
  �ҺŽ��ϼ�_In ������ü�¼.ʵ�ս��%Type,
  ����id_In       Ʊ��ʹ����ϸ.����id%Type,
  �շ�Ʊ��_In     Number := 0, --�Һ��Ƿ�ʹ���շ�Ʊ��
  ������ˮ��_In   ����Ԥ����¼.������ˮ��%Type,
  ����˵��_In     ����Ԥ����¼.����˵��%Type,
  ԤԼ��ʽ_In     ԤԼ��ʽ.����%Type := Null,
  Ԥ��id_In       ����Ԥ����¼.Id%Type := Null,
  �����id_In     ����Ԥ����¼.�����id%Type := Null,
  �������״̬_In Number := 0,
  �Ƿ������豸_In Number := 0,
  ����id_In       ������ü�¼.����id%Type := Null,
  ��������_In     Number := 0,
  ���ս���_In     Varchar2 := Null,
  ��Ԥ��_In       Number := Null,
  ֧������_In     ����Ԥ����¼.����%Type := Null,
  �˺�����_In     Number := 1,
  �ѱ�_In         ������ü�¼.�ѱ�%Type := Null,
  ������_In       �Һ����״̬.������%Type := Null,
  ��������_In     Number := 0
) As
  --���ܣ������������йҺ�(����ԤԼ;ԤԼ�ҺŲ��ۿ�;ԤԼ�Һſۿ�)
  --���:������ʽ_IN:1-��ʾ�Һ�,2-��ʾԤԼ�ҺŲ��ۿ�,3-��ʾԤԼ�Һ�,�ۿ�
  --      �������״̬_In:1-��ʾǿ�Ƽ���Һ����״̬����;�������������Ż�����ʱ��ʱ�ż���.
  --      �Ƿ������豸_In:1-��ʾ��ҽԺ�������豸���д˺����ĵ���,�����豸���ô˺��� ����Ӻ�,��������
  --      ��������_In :0-������������ 1-��ʾ�Ե��ݽ�������,����δ��Ч�ĵ�����Ϣ;2-�������ļ�¼���н���-������������:δ��Ч�ĵ��������пۿ���ɺ���н���
  --      ���ս���_IN:��ʽ="���㷽ʽ|������||....."
  Err_Item Exception;
  v_Err_Msg  Varchar2(255);
  n_��ӡid   Ʊ�ݴ�ӡ����.Id%Type;
  n_����ֵ   ����Ԥ����¼.���%Type;
  v_�ŶӺ��� Varchar2(20);
  v_�������� �ŶӽкŶ���.��������%Type;
  n_Ԥ��id   ����Ԥ����¼.Id%Type;
  n_�Һ�id   ���˹Һż�¼.Id%Type;
  v_�������� Varchar2(3000);
  v_��ǰ���� Varchar2(150);

  v_���㷽ʽ       ����Ԥ����¼.���㷽ʽ%Type;
  n_������       ����Ԥ����¼.��Ԥ��%Type;
  n_����ϼ�       Number(16, 5);
  n_Ԥ�����       ����Ԥ����¼.��Ԥ��%Type;
  n_��id           ����ɿ����.Id%Type;
  d_�Ŷ�ʱ��       Date;
  n_����           Number;
  n_ͬ����Լһ���� Number(18);
  n_����ԤԼ������ Number(18);
  n_��Լ����       Number(18);

  n_������λ����       Number(18);
  n_�Ƿ񿪷�           Number(1);
  n_Count              Number(18);
  n_�к�               Number(18);
  n_���               ���˹Һż�¼.����%Type;
  n_����id             ������ü�¼.Id%Type;
  n_�۸񸸺�           Number(18);
  n_ԭ��Ŀid           �շ���ĿĿ¼.Id%Type;
  n_ԭ������Ŀid       �շ���ĿĿ¼.Id%Type;
  v_����               ���˹Һż�¼.����%Type;
  n_����id             �ҺŰ���.Id%Type;
  n_ʵ�ս��ϼ�       ������ü�¼.ʵ�ս��%Type;
  n_��������id         ������ü�¼.��������id%Type;
  n_ʵ�ս��           ������ü�¼.ʵ�ս��%Type;
  n_Ӧ�ս��           ������ü�¼.ʵ�ս��%Type;
  n_����id             ���˽��ʼ�¼.Id%Type;
  v_Temp               Varchar2(500);
  n_ԤԼʱ�����       Number;
  n_ԤԼ����           Number;
  d_ʱ�ο�ʼʱ��       Date;
  n_ԤԼ����           ������λ�ҺŻ���.��Լ��%Type;
  n_����               ���˹Һż�¼.����%Type;
  d_�Ǽ�ʱ��           Date;
  v_����Ա���         ��Ա��.���%Type;
  v_����Ա����         ��Ա��.����%Type;
  n_ԤԼ               Integer;
  v_����               �ҺŰ���ʱ��.����%Type;
  n_���÷�ʱ��         Integer;
  n_�ѹ���             ���˹ҺŻ���.�ѹ���%Type;
  n_��Լ��             ���˹ҺŻ���.��Լ��%Type;
  n_�����ѽ���         ���˹ҺŻ���.��Լ��%Type;
  n_ԤԼ���ɶ���       Number;
  d_Date               Date;
  n_�Һ����           Number;
  v_�Ŷӱ��           �ŶӽкŶ���.�Ŷӱ��%Type;
  v_�Ŷ����           �ŶӽкŶ���.�Ŷ����%Type;
  v_������             �Һ����״̬.������%Type;
  v_��Ų���Ա         �Һ����״̬.����Ա����%Type;
  v_��Ż�����         �Һ����״̬.������%Type;
  n_�������           Number := 0;
  v_���ʽ           ���˹Һż�¼.ҽ�Ƹ��ʽ%Type;
  v_�ѱ�               ������ü�¼.�ѱ�%Type;
  n_���ηѱ�           Number(3) := 0;
  n_Tmp����id          �ҺŰ���.Id%Type;
  n_�ƻ�id             �ҺŰ��żƻ�.Id%Type;
  v_����               ������Ϣ.����%Type;
  n_������λ������ģʽ Number;

  Cursor c_Pati(n_����id ������Ϣ.����id%Type) Is
    Select a.����id, a.����, a.�Ա�, a.����, a.סԺ��, a.�����, a.�ѱ�, a.����, c.���� As ���ʽ
    From ������Ϣ A, ҽ�Ƹ��ʽ C
    Where a.����id = n_����id And a.ҽ�Ƹ��ʽ = c.����(+);

  r_Pati c_Pati%RowType;

  --���α������շѳ�Ԥ���Ŀ���Ԥ���б�
  --��ID�������ȳ��ϴ�δ����ġ�
  Cursor c_Deposit(n_����id ������Ϣ.����id%Type) Is
    Select *
    From (Select a.Id, a.��¼״̬, a.Ԥ�����, a.No, Nvl(a.���, 0) As ���
           From ����Ԥ����¼ A,
                (Select NO, Sum(Nvl(a.���, 0)) As ���
                  From ����Ԥ����¼ A
                  Where a.����id Is Null And Nvl(a.���, 0) <> 0 And a.����id = n_����id And Nvl(a.Ԥ�����, 2) = 1
                  Group By NO
                  Having Sum(Nvl(a.���, 0)) <> 0) B
           Where a.����id Is Null And Nvl(a.���, 0) <> 0 And a.���㷽ʽ Not In (Select ���� From ���㷽ʽ Where ���� = 5) And
                 a.No = b.No And a.����id = n_����id And Nvl(a.Ԥ�����, 2) = 1
           Union All
           Select 0 As ID, ��¼״̬, Ԥ�����, NO, Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) As ���
           From ����Ԥ����¼
           Where ��¼���� In (1, 11) And ����id Is Not Null And Nvl(���, 0) <> Nvl(��Ԥ��, 0) And ����id = n_����id And
                 Nvl(Ԥ�����, 2) = 1 Having Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0
           Group By ��¼״̬, NO, Ԥ�����)
    Order By ID, NO;

  Cursor c_����
  (
    v_����        �ҺŰ���.����%Type,
    d_����ʱ��_In Date
  ) Is
    Select *
    From (With ����ʱ��� As (Select ʱ���
                         From (Select ʱ���,
                                       To_Date(Decode(Sign(��ʼʱ�� - ��ֹʱ��), 1, '3000-01-09 ' || To_Char(��ʼʱ��, 'HH24:MI:SS'),
                                                       '3000-01-10 ' || To_Char(��ʼʱ��, 'HH24:MI:SS')), 'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                                       To_Date(Decode(Sign(��ʼʱ�� - ��ֹʱ��), 1, '3000-01-11 ' || To_Char(��ֹʱ��, 'HH24:MI:SS'),
                                                       '3000-01-10 ' || To_Char(��ֹʱ��, 'HH24:MI:SS')), 'yyyy-mm-dd hh24:mi:ss') As ��ֹʱ��,
                                       To_Date('3000-01-10 ' || To_Char(d_����ʱ��_In, 'HH24:MI:SS'), 'yyyy-mm-dd hh24:mi:ss') As ��ǰʱ��,
                                       To_Date('3000-01-10 ' || To_Char(��ʼʱ��, 'HH24:MI:SS'), 'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��1,
                                       To_Date('3000-01-10 ' || To_Char(��ֹʱ��, 'HH24:MI:SS'), 'yyyy-mm-dd hh24:mi:ss') As ��ֹʱ��1
                                From ʱ���)
                         Where ��ǰʱ�� Between ��ʼʱ�� And ��ֹʱ��1 Or ��ǰʱ�� Between ��ʼʱ��1 And ��ֹʱ��)
           Select Distinct p.Id, p.����, p.����, p.����id, b.���� As ���ұ���, b.���� As ��������, p.��Ŀid, c.���� As ��Ŀ����, c.���� As ��Ŀ����,
                           p.ҽ��id, d.��� As ҽ�����, p.ҽ������, p.�޺���, p.��Լ��, p.���� As ��, p.��һ As һ, p.�ܶ� As ��, p.���� As ��,
                           p.���� As ��, p.���� As ��, p.���� As ��, p.��ſ���
           From (Select p.Id, p.����, p.����, p.����id, p.��Ŀid, p.ҽ��id, p.ҽ������, b.�޺���, b.��Լ��, Nvl(p.��������, 0) As ��������, p.����, p.��һ,
                         p.�ܶ�, p.����, p.����, p.����, p.����, p.���﷽ʽ, p.��ſ���,
                         Decode(To_Char(d_����ʱ��_In, 'D'), '1', p.����, '2', p.��һ, '3', p.�ܶ�, '4', p.����, '5', p.����, '6', p.����,
                                 '7', p.����, Null) As �Ű�
                  From �ҺŰ��� P, �ҺŰ������� B
                  Where p.ͣ������ Is Null And p.Id = b.����id(+) And
                        b.������Ŀ(+) = Decode(To_Char(d_����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����',
                                           '6', '����', '7', '����', Null) And
                        d_����ʱ��_In Between Nvl(p.��ʼʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                        Nvl(p.��ֹʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And Not Exists
                   (Select 1
                         From �ҺŰ��żƻ�
                         Where ����id = p.Id And (d_����ʱ��_In Between ��Чʱ�� And ʧЧʱ��) And ���ʱ�� Is Not Null) And Not Exists
                   (Select 1
                         From �ҺŰ���ͣ��״̬
                         Where ����id = p.Id And d_����ʱ��_In Between ��ʼֹͣʱ�� And ����ֹͣʱ��) And p.���� = v_����
                  Union All
                  Select c.Id, c.����, c.����, c.����id, p.��Ŀid, p.ҽ��id, p.ҽ������, b.�޺���, b.��Լ��, Nvl(c.��������, 0) As ��������, p.����, p.��һ,
                         p.�ܶ�, p.����, p.����, p.����, p.����, p.���﷽ʽ, p.��ſ���,
                         Decode(To_Char(d_����ʱ��_In, 'D'), '1', p.����, '2', p.��һ, '3', p.�ܶ�, '4', p.����, '5', p.����, '6', p.����,
                                 '7', p.����, Null) As �Ű�
                  From �ҺŰ��żƻ� P, �ҺŰ��� C, �Һżƻ����� B,
                       (Select Max(a.��Чʱ��) As ��Ч, ����id
                         From �ҺŰ��żƻ� A, �ҺŰ��� B
                         Where a.����id = b.Id And a.���ʱ�� Is Not Null And
                               ����ʱ��_In Between Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                               Nvl(a.ʧЧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) And b.���� = ����_In
                         Group By ����id) E
                  Where p.����id = c.Id And p.Id = b.�ƻ�id(+) And p.��Чʱ�� = e.��Ч And p.����id = e.����id And
                        Nvl(p.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) = Nvl(e.��Ч, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                        b.������Ŀ(+) = Decode(To_Char(d_����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����',
                                           '6', '����', '7', '����', Null) And (d_����ʱ��_In Between p.��Чʱ�� And p.ʧЧʱ��) And
                        p.���ʱ�� Is Not Null And Not Exists
                   (Select 1
                         From �ҺŰ���ͣ��״̬
                         Where ����id = c.Id And d_����ʱ��_In Between ��ʼֹͣʱ�� And ����ֹͣʱ��) And p.���� = v_����) P, ���ű� B, �շ���ĿĿ¼ C,
                ��Ա�� D
           Where p.����id = b.Id And p.ҽ��id = d.Id(+) And p.��Ŀid = c.Id And
                 (c.����ʱ�� Is Null Or c.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) And
                 (Nvl(p.ҽ��id, 0) = 0 Or Exists
                  (Select 1
                   From ��Ա�� Q
                   Where p.ҽ��id = q.Id And (q.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or q.����ʱ�� Is Null))) And Exists
            (Select 1 From ����ʱ��� Where ʱ��� = p.�Ű�))
           Order By ����;


  r_���� c_����%RowType;

  Function Zl_����(����_In �ҺŰ���.����%Type) Return Varchar2 As
    n_���﷽ʽ �ҺŰ���.���﷽ʽ%Type;
    n_����id   �ҺŰ���.Id%Type;
    v_����     ���˹Һż�¼.����%Type;
    v_Rowid    Varchar2(500);
    n_Next     Integer;
    n_First    Integer;
  Begin
  
    If ��������_In = 2 Then
      --�Ե��ݽ��н���,���ȼ���Ƿ��������
      Select Count(Rowid) Into n_���� From ���˹Һż�¼ Where����¼״̬ = 0 And NO = ���ݺ�_In And ����id = ����id_In;
      If n_���� = 0 Then
        v_Err_Msg := '���ݺ�Ϊ(' || ���ݺ�_In || ')�ĵ���,�����ڻ����Ѿ�������!';
        Raise Err_Item;
      End If;
      Select Max(����) Into n_���� From ���˹Һż�¼ Where����¼״̬ = 0 And NO = ���ݺ�_In And ����id = ����id_In;
    End If;
  
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

  Function Zl_����Ա
  (
    Type_In     Integer,
    Splitstr_In Varchar2
  ) Return Varchar2 As
    n_Step Number(18);
    v_Sub  Varchar2(1000);
    --Type_In:0-��ȡȱʡ����ID;1-��ȡ����Ա���;2-��ȡ����Ա����
    -- SplitStr:��ʽΪ:����ID,��������;��ԱID,��Ա���,��Ա����(��Zl_Identity��ȡ��)
  Begin
    If Type_In = 0 Then
      --ȱʡ����
      n_Step := Instr(Splitstr_In, ',');
      v_Sub  := Substr(Splitstr_In, 1, n_Step - 1);
      Return v_Sub;
    End If;
    If Type_In = 1 Then
      --����Ա����
      n_Step := Instr(Splitstr_In, ';');
      v_Sub  := Substr(Splitstr_In, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, 1, n_Step - 1);
      Return v_Sub;
    End If;
    If Type_In = 2 Then
      --����Ա����
      n_Step := Instr(Splitstr_In, ';');
      v_Sub  := Substr(Splitstr_In, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      Return v_Sub;
    End If;
  End;

Begin
  If �ѱ�_In Is Null Then
    Select Zl_Custom_Getpatifeetype(1, ����id_In) Into v_�ѱ� From Dual;
  Else
    v_�ѱ� := �ѱ�_In;
  End If;
  If v_�ѱ� Is Null Then
    n_���ηѱ� := 1;
    Select ���� Into v_�ѱ� From �ѱ� Where ȱʡ��־ = 1 And Rownum < 2;
  End If;
  Update ������Ϣ Set �ѱ� = v_�ѱ� Where ����id = ����id_In;
  If ��������_In = 1 Then
    Select Zl_Age_Calc(����id_In) Into v_���� From Dual;
    If v_���� Is Not Null Then
      Update ������Ϣ Set ���� = v_���� Where ����id = ����id_In;
    End If;
  End If;
  --��ȡ��ǰ��������
  If ������_In Is Not Null Then
    v_������ := ������_In;
  Else
    Select Terminal Into v_������ From V$session Where Audsid = Userenv('sessionid');
  End If;
  n_ʵ�ս��ϼ� := 0;
  Select Count(*) + 1
  Into n_�Һ����
  From ���˹Һż�¼
  Where �ű� = ����_In And �Ǽ�ʱ�� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In + 1) - 1 / 24 / 60 / 60;
  --Begin
  --����ID,��������;��ԱID,��Ա���,��Ա����
  v_Temp := Zl_Identity(0);
  If Nvl(v_Temp, ' ') = ' ' Then
    v_Err_Msg := '��ǰ������Աδ���ö�Ӧ����Ա��ϵ,���ܼ�����';
    Raise Err_Item;
  End If;

  If �Ǽ�ʱ��_In Is Null Then
    d_�Ǽ�ʱ�� := Sysdate;
  Else
    d_�Ǽ�ʱ�� := �Ǽ�ʱ��_In;
  End If;
  If Trunc(Sysdate) > Trunc(����ʱ��_In) Then
    v_Err_Msg := '���ܹ���ǰ�ĺ�(' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || ')��';
    Raise Err_Item;
  End If;
  n_ͬ����Լһ���� := Nvl(zl_GetSysParameter('����ͬ����Լһ����', 1111), 0);
  n_����ԤԼ������ := Nvl(zl_GetSysParameter('����ԤԼ������', 1111), 0);
  n_��������id     := To_Number(Zl_����Ա(0, v_Temp));
  v_����Ա���     := Zl_����Ա(1, v_Temp);
  v_����Ա����     := Zl_����Ա(2, v_Temp);
  n_��id           := Zl_Get��id(v_����Ա����);

  If ������ʽ_In <> 1 Then
    --ԤԼ����Ƿ���Ӻ�����λ����
    --��������˺�����λ���� ��
    Begin
      Select ID
      Into n_�ƻ�id
      From �ҺŰ��żƻ�
      Where ���� = ����_In And ����ʱ��_In Between Nvl(��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
            Nvl(ʧЧʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And Rownum < 2
      Order By ��Чʱ�� Desc;
    Exception
      When Others Then
        Select ID Into n_Tmp����id From �ҺŰ��� Where ���� = ����_In;
    End;
    If Nvl(n_�ƻ�id, 0) <> 0 Then
      Select Count(0)
      Into n_������λ����
      From ������λ�ƻ�����
      Where ������λ = ������λ_In And �ƻ�id = n_�ƻ�id And Rownum < 2;
    Else
      Select Count(0)
      Into n_������λ����
      From ������λ���ſ���
      Where ������λ = ������λ_In And ����id = n_Tmp����id And Rownum < 2;
    End If;
  End If;

  If ������ʽ_In <> 2 Then
    v_���� := Zl_����(����_In);
  End If;
  If ������ʽ_In <> 2 And ���㷽ʽ_In Is Not Null Then
    --�����㷽ʽ�Ƿ��걸
    Select Count(*) Into n_Count From ���㷽ʽ Where ���� = Nvl(���㷽ʽ_In, 'Lxh') And ���� In (2, 7, 8);
    If Nvl(�����id_In, 0) <> 0 And n_Count = 0 Then
      Select Count(1)
      Into n_Count
      From ҽ�ƿ����
      Where ID = Nvl(�����id_In, 0) And ���㷽ʽ = Nvl(���㷽ʽ_In, 'lxh');
    End If;
    If n_Count = 0 Then
      v_Err_Msg := '���㷽ʽ(' || ���㷽ʽ_In || ')δ����,���ڽ��㷽ʽ���������á�';
      Raise Err_Item;
    End If;
  End If;

  --��Ϊ�����а��ձ൥�ݺŹ���,�չҺ������ܳ���10000��,����Ҫ���ΨһԼ����
  Select Count(*) Into n_Count From ������ü�¼ Where ��¼���� = 4 And ��¼״̬ In (1, 3) And NO = ���ݺ�_In;
  If n_Count <> 0 Then
    v_Err_Msg := '�Һŵ��ݺ��ظ�,���ܱ��棡' || Chr(13) || '���ʹ���˰���˳����,���չҺ������ܳ���10000�˴Ρ�';
    Raise Err_Item;
  End If;

  Open c_Pati(����id_In);
  n_Count := 0;
  Begin
    Fetch c_Pati
      Into r_Pati;
  Exception
    When Others Then
      n_Count := -1;
  End;
  If n_Count = -1 Then
    v_Err_Msg := '����δ�ҵ������ܼ�����';
    Raise Err_Item;
  End If;

  Open c_����(����_In, ����ʱ��_In);
  Begin
    Fetch c_����
      Into r_����;
  Exception
    When Others Then
      n_Count := -1;
  End;
  If n_Count = -1 Then
    v_Err_Msg := '�úű�û����' || To_Char(����ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '�н��а��š�';
    Raise Err_Item;
  End If;

  Select Decode(To_Char(����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����', '����')
  Into v_����
  From Dual;
  Begin
    Select 1 Into n_���÷�ʱ�� From �ҺŰ���ʱ�� Where ����id = r_����.Id And ���� = v_���� And Rownum <= 1;
  Exception
    When Others Then
      n_���÷�ʱ�� := 0;
  End;

  --�Բ������ƽ��м��
  --����ԤԼ���ۿ�ʱ���м��
  If ������ʽ_In = 2 Then
    If Nvl(n_ͬ����Լһ����, 0) <> 0 Or Nvl(n_����ԤԼ������, 0) <> 0 Then
      n_��Լ���� := 0;
      For c_Chkitem In (Select Count(1) As ��Լ, a.ִ�в���id As ����id, Nvl(k.����, '') As ����
                        From ���˹Һż�¼ A, ������Ϣ B, ���ű� K
                        Where a.����id = b.����id And a.����id = ����id_In And a.ִ�в���id = k.Id(+) And a.��¼���� = 2 And ��¼״̬ = 1 And
                              a.ԤԼʱ�� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60
                        Group By a.ִ�в���id, k.����) Loop
        If Nvl(n_ͬ����Լһ����, 0) <> 0 And c_Chkitem.����id = r_����.����id Then
        
          v_Err_Msg := '�ò����Ѿ��ڿ���[' || c_Chkitem.���� || ']������ԤԼ,������ԤԼ��';
          Raise Err_Item;
        
          If Nvl(n_����ԤԼ������, 0) > 0 And c_Chkitem.����id <> r_����.����id Then
            n_��Լ���� := n_��Լ���� + 1;
          End If;
        End If;
      End Loop;
      If n_��Լ���� >= Nvl(n_����ԤԼ������, 0) And Nvl(n_����ԤԼ������, 0) > 0 Then
        v_Err_Msg := 'ͬһ���������ͬʱԤԼ[' || Nvl(n_����ԤԼ������, 0) || ']������,������ԤԼ��';
        Raise Err_Item;
      End If;
    End If;
  End If;

  d_Date         := Null;
  d_ʱ�ο�ʼʱ�� := Null;

  If Nvl(r_����.�޺���, 0) >= 0 Or r_����.�޺��� Is Null Then
  
    Select Nvl(Sum(Nvl(b.�ѹ���, 0)), 0), Nvl(Sum(Nvl(b.�����ѽ���, 0)), 0), Nvl(Sum(Nvl(b.��Լ��, 0)), 0)
    Into n_�ѹ���, n_�����ѽ���, n_��Լ��
    From �ҺŰ��� A, ���˹ҺŻ��� B
    Where a.����id = b.����id And a.��Ŀid = b.��Ŀid And a.���� = ����_In And b.���� Between Trunc(����ʱ��_In) And
          Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And (a.���� = b.���� Or b.���� Is Null) And Nvl(a.ҽ��id, 0) = Nvl(b.ҽ��id, 0) And
          Nvl(a.ҽ������, 'ҽ��') = Nvl(b.ҽ������, 'ҽ��');
  
    If n_���÷�ʱ�� = 1 Then
      If Nvl(r_����.��ſ���, 0) = 1 Then
        If Nvl(�Ƿ������豸_In, 0) = 0 Then
          Select Count(*), Max(��ʼʱ��)
          Into n_Count, d_ʱ�ο�ʼʱ��
          From �ҺŰ���ʱ��
          Where ����id = r_����.Id And ���� = v_���� And ��� = Nvl(����_In, 0);
          v_Temp := '�Һ�';
          If ������ʽ_In > 1 Then
            v_Temp := 'ԤԼ�Һ�';
          End If;
        
          If n_Count = 0 Then
            v_Err_Msg := '�ű�Ϊ' || ����_In || '�ĹҺŰ����в��������Ϊ' || Nvl(����_In, 0) || '�İ���,������' || v_Temp || '��';
            Raise Err_Item;
          End If;
        End If;
        --�����,����ѡ��Һ�
        If Trunc(Sysdate) = Trunc(����ʱ��_In) Then
          --�ҵ���ĺ�
          v_Temp := To_Char(Sysdate, 'yyyy-mm-dd') || ' ';
          For v_ʱ�� In (Select To_Date(v_Temp || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                              To_Date(To_Char(Sysdate + Decode(Sign(��ʼʱ�� - ����ʱ��), 1, 1, 0), 'yyyy-mm-dd') || ' ' ||
                                       To_Char(����ʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, ��������, �Ƿ�ԤԼ
                       From �ҺŰ���ʱ��
                       Where ����id = r_����.Id And ���� = v_���� And ��� = Nvl(����_In, 0)) Loop
            If Sysdate > v_ʱ��.����ʱ�� Then
              v_Err_Msg := '�ű�Ϊ' || ����_In || '������Ϊ' || Nvl(����_In, 0) || '�İ���,�Ѿ�����ʱ��,������' || v_Temp || '��';
              Raise Err_Item;
            End If;
          End Loop;
        End If;
      Elsif ������ʽ_In > 1 Then
        --δ������ŵ�,��Ҫ���ԤԼ�����
      
        n_Count := 0;
        For v_ʱ�� In (Select ���, ��ʼʱ��, ����ʱ��, ��������, �Ƿ�ԤԼ
                     From �ҺŰ���ʱ��
                     Where ����id = r_����.Id And ���� = v_���� And
                           (('3000-01-10 ' || To_Char(����ʱ��_In, 'HH24:MI:SS') Between
                           Decode(Sign(��ʼʱ�� - ����ʱ�� - 1 / 24 / 60 / 60), 1, '3000-01-09 ' || To_Char(��ʼʱ��, 'HH24:MI:SS'),
                                    '3000-01-10 ' || To_Char(��ʼʱ��, 'HH24:MI:SS')) And
                           '3000-01-10 ' || To_Char(����ʱ�� - 1 / 24 / 60 / 60, 'HH24:MI:SS')) Or
                           ('3000-01-10 ' || To_Char(����ʱ��_In, 'HH24:MI:SS') Between
                           '3000-01-10 ' || To_Char(��ʼʱ��, 'HH24:MI:SS') And
                           Decode(Sign(��ʼʱ�� - ����ʱ�� - 1 / 24 / 60 / 60), 1,
                                    '3000-01-11 ' || To_Char(����ʱ�� - 1 / 24 / 60 / 60, 'HH24:MI:SS'),
                                    '3000-01-10 ' || To_Char(����ʱ�� - 1 / 24 / 60 / 60, 'HH24:MI:SS'))))) Loop
          n_ԤԼʱ����� := v_ʱ��.���;
          d_ʱ�ο�ʼʱ�� := v_ʱ��.��ʼʱ��;
        
          Select Count(*), Max(���)
          Into n_Count, n_ԤԼ����
          From �Һ����״̬
          Where ���� = ����_In And ���� = ����ʱ��_In And ״̬ Not In (4, 5);
        
          If Nvl(n_Count, 0) > Nvl(v_ʱ��.��������, 0) And ��������_In <> 2 Then
            v_Err_Msg := '�ű�Ϊ' || ����_In || '�ĹҺŰ�������' || To_Char(v_ʱ��.��ʼʱ��, 'hh24:mi:ss') || '��' ||
                         To_Char(v_ʱ��.����ʱ��, 'hh24:mi:ss') || '���ֻ��ԤԼ' || Nvl(v_ʱ��.��������, 0) || '��,�����ٽ���ԤԼ�Һţ�';
            Raise Err_Item;
          End If;
          n_Count := 1;
        End Loop;
      
        If n_Count = 0 Then
          v_Err_Msg := '�ű�Ϊ' || ����_In || '�ĹҺŰ�����û����صİ��żƻ�(' || To_Char(����ʱ��_In, 'YYYY-mm-dd HH24:MI:SS') ||
                       '),���ܽ���ԤԼ�Һţ�';
          Raise Err_Item;
        End If;
      End If;
    End If;
  End If;

  If ������ʽ_In = 1 And ��������_In <> 2 Then
    --�ҺŹ���:
    --  �ѹ������ܴ����޺���
    If n_�ѹ��� >= Nvl(r_����.�޺���, 0) And r_����.�޺��� Is Not Null Then
      v_Err_Msg := '�úű�����Ѵﵽ�޺��� ' || Nvl(r_����.�޺���, 0) || '�����ٹҺţ�';
      Raise Err_Item;
    End If;
  End If;

  If ������ʽ_In > 1 Then
    --ԤԼ����ؼ��
    --����:
    --   1.����Լ���ܳ�����Լ��
    --   2.����Ƿ�����ʱ�ε�
    If n_��Լ�� >= Nvl(r_����.��Լ��, 0) And Nvl(r_����.��Լ��, 0) <> 0 And r_����.��Լ�� Is Not Null And ��������_In <> 2 Then
      v_Err_Msg := '�úű��Ѵﵽ��Լ�� ' || Nvl(r_����.��Լ��, 0) || '������ԤԼ�Һţ�';
      Raise Err_Item;
    End If;
  End If;
  If n_������λ���� > 0 And ������ʽ_In <> 1 And ������λ_In Is Not Null Then
  
    If Nvl(r_����.��ſ���, 0) = 1 And Nvl(����_In, 0) = 0 Then
      v_Err_Msg := '��ǰ����ʹ������ſ���,��ȷ������ҪԤԼ�����,���ܼ�����';
      Raise Err_Item;
    End If; --Nvl(r_����.��ſ���, 0) =0
  
    n_��� := Case
              When Nvl(r_����.��ſ���, 0) = 1 Or n_���÷�ʱ�� = 1 And ������ʽ_In > 1 Then
               Nvl(����_In, 0)
              Else
               0
            End;
    --������λ������ģʽ
    Begin
      If Nvl(n_�ƻ�id, 0) <> 0 Then
        Select 0
        Into n_���
        From ������λ�ƻ�����
        Where ������λ = ������λ_In And �ƻ�id = n_�ƻ�id And
              ������Ŀ = Decode(To_Char(����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7',
                            '����', Null) And ���� <> 0 And ��� = 0 And Rownum < 2;
      Else
        Select 0
        Into n_���
        From ������λ���ſ���
        Where ������λ = ������λ_In And ����id = n_Tmp����id And
              ������Ŀ = Decode(To_Char(����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7',
                            '����', Null) And ���� <> 0 And ��� = 0 And Rownum < 2;
      End If;
      n_������λ������ģʽ := 1;
    Exception
      When Others Then
        n_������λ������ģʽ := 0;
    End;
    --������ż��
    For c_������λ In (Select c.���, ����
                   From �ҺŰ��� A, ������λ���ſ��� C
                   Where a.���� = ����_In And Decode(To_Char(����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                                 '����', '6', '����', '7', '����', Null) = c.������Ŀ(+) And a.Id = c.����id And
                         c.������λ = ������λ_In And c.��� = n_��� And Not Exists
                    (Select 1
                          From �ҺŰ��żƻ� D
                          Where d.����id = a.Id And d.���ʱ�� Is Not Null And
                                ����ʱ��_In Between Nvl(d.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                                Nvl(d.ʧЧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')))
                   Union All
                   Select c.���, ����
                   From �ҺŰ��żƻ� A, �ҺŰ��� D, ������λ�ƻ����� C,
                        (Select Max(a.��Чʱ��) As ��Ч, ����id
                          From �ҺŰ��żƻ� A, �ҺŰ��� B
                          Where a.����id = b.Id And a.���ʱ�� Is Not Null And
                                ����ʱ��_In Between Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                                Nvl(a.ʧЧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) And b.���� = ����_In
                          Group By ����id) E
                   Where a.����id = d.Id And a.���ʱ�� Is Not Null And d.���� = ����_In And a.����id = e.����id And
                         Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) =
                         Nvl(e.��Ч, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                         Decode(To_Char(����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',
                                '7', '����', Null) = c.������Ŀ(+) And a.Id = c.�ƻ�id And c.������λ = ������λ_In And c.��� = n_��� And
                         ����ʱ��_In Between Nvl(a.��Чʱ��, To_Date('1900-01-01', 'yyyy-mm-dd')) And
                         Nvl(a.ʧЧʱ��, To_Date('3000-01-01', 'yyyy-mm-dd'))) Loop
    
      If Nvl(r_����.��ſ���, 0) = 1 And c_������λ.��� = n_��� And n_������λ������ģʽ = 0 Then
        n_�Ƿ񿪷� := 1;
        Exit;
      Elsif (Nvl(r_����.��ſ���, 0) = 0 And c_������λ.��� = n_���) Or n_������λ������ģʽ = 1 Then
        Begin
          Select Nvl(��Լ��, 0)
          Into n_ԤԼ����
          From ������λ�ҺŻ���
          Where ������λ = ������λ_In And ���� = Trunc(����ʱ��_In) And ���� = ����_In;
        Exception
          When Others Then
            n_ԤԼ���� := 0;
        End;
        If c_������λ.���� <= n_ԤԼ���� And Nvl(c_������λ.����, 0) > 0 And ��������_In <> 2 Then
          v_Err_Msg := '�úű��Ѵﵽ��Լ�� ' || Nvl(c_������λ.����, 0) || '������ԤԼ�Һţ�';
          Raise Err_Item;
        End If;
        n_�Ƿ񿪷� := 1;
        Exit;
      End If;
    
    End Loop;
  
    If Nvl(n_�Ƿ񿪷�, 0) = 0 Then
      v_Err_Msg := '��ǰ���(' || Nvl(����_In, 0) || 'δ����,���ܼ�����';
      Raise Err_Item;
    End If;
  End If;

  --����޺�������Լ��
  n_�к�         := 1;
  n_ԭ��Ŀid     := 0;
  n_ԭ������Ŀid := 0;
  n_ʵ�ս��ϼ� := 0;
  If ��������_In <> 1 Then
    If ������ʽ_In <> 2 Then
      If Nvl(����id_In, 0) = 0 Then
        --����Ӧ�ó�����
        Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
      Else
        n_����id := ����id_In;
      End If;
    Else
      n_����id := Null;
    End If;
  End If;
  For c_Item In (Select 1 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, 1 As ����, c.Id As ������Ŀid,
                        c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, Null As ��������
                 From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C
                 Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = r_����.��Ŀid And Sysdate Between b.ִ������ And
                       Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))
                 Union All
                 Select 2 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, d.�������� As ����,
                        c.Id As ������Ŀid, c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, 1 As ��������
                 From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C, �շѴ�����Ŀ D
                 Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = d.����id And d.����id = r_����.��Ŀid And
                       Sysdate Between b.ִ������ And Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))
                 Order By ����, ��Ŀ����, �������) Loop
    n_�۸񸸺� := Null;
    If n_ԭ��Ŀid = c_Item.��Ŀid Then
      If n_ԭ������Ŀid <> c_Item.������Ŀid Then
        n_�۸񸸺� := n_�к�;
      End If;
      n_ԭ������Ŀid := c_Item.������Ŀid;
    End If;
    n_ԭ��Ŀid := c_Item.��Ŀid;
    n_Ӧ�ս�� := Round(c_Item.���� * c_Item.����, 5);
    n_ʵ�ս�� := n_Ӧ�ս��;
    If Nvl(c_Item.���ηѱ�, 0) <> 1 And n_���ηѱ� = 0 Then
      --����:
      v_Temp     := Zl_Actualmoney(r_Pati.�ѱ�, c_Item.��Ŀid, c_Item.������Ŀid, n_Ӧ�ս��);
      v_Temp     := Substr(v_Temp, Instr(v_Temp, ':') + 1);
      n_ʵ�ս�� := Zl_To_Number(v_Temp);
    End If;
    n_ʵ�ս��ϼ� := Nvl(n_ʵ�ս��ϼ�, 0) + n_ʵ�ս��;
  
    --�������ݲ���������
    If ��������_In <> 1 Then
      --�������˹Һŷ���(���ܵ����ǻ������������)
      Select ���˷��ü�¼_Id.Nextval Into n_����id From Dual;
      --:������ʽ_IN:1-��ʾ�Һ�,2-��ʾԤԼ�ҺŲ��ۿ�,3-��ʾԤԼ�Һ�,�ۿ�
      Insert Into ������ü�¼
        (ID, ��¼����, ��¼״̬, ���, �۸񸸺�, ��������, NO, ʵ��Ʊ��, �����־, �Ӱ��־, ���ӱ�־, ��ҩ����, ����id, ��ʶ��, ���ʽ, ����, �Ա�, ����, �ѱ�, ���˿���id,
         �շ����, ���㵥λ, �շ�ϸĿid, ������Ŀid, �վݷ�Ŀ, ����, ����, ��׼����, Ӧ�ս��, ʵ�ս��, ���ʽ��, ����id, ���ʷ���, ��������id, ������, ������, ִ�в���id, ִ����,
         ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ���մ���id, ������Ŀ��, ���ձ���, ͳ����, ժҪ, ����, �ɿ���id)
      Values
        (n_����id, 4, Decode(������ʽ_In, 2, 0, 1), n_�к�, n_�۸񸸺�, c_Item.��������, ���ݺ�_In, Ʊ�ݺ�_In, 1, Null, Null,
         Decode(������ʽ_In, 2, To_Char(����_In), v_����), r_Pati.����id, r_Pati.�����, r_Pati.���ʽ, r_Pati.����, r_Pati.�Ա�, r_Pati.����,
         r_Pati.�ѱ�, r_����.����id, c_Item.���, ����_In, c_Item.��Ŀid, c_Item.������Ŀid, c_Item.�վݷ�Ŀ, 1, c_Item.����, c_Item.����,
         n_Ӧ�ս��, n_ʵ�ս��, Decode(������ʽ_In, 2, Null, n_ʵ�ս��), n_����id, 0, n_��������id, v_����Ա����,
         Decode(������ʽ_In, 2, v_����Ա����, Null), r_����.����id, r_����.ҽ������, v_����Ա���, v_����Ա����, ����ʱ��_In, d_�Ǽ�ʱ��, Null, 0, Null, Null,
         ժҪ_In, ԤԼ��ʽ_In, Decode(������ʽ_In, 2, Null, n_��id));
    End If;
    n_�к� := n_�к� + 1;
  
  End Loop;

  If Round(Nvl(�ҺŽ��ϼ�_In, 0), 5) <> Round(Nvl(n_ʵ�ս��ϼ�, 0), 5) Then
    v_Err_Msg := '���ιҺŽ���ȷ,��������ΪҽԺ�����˼۸�,�����»�ȡ�Һ��շ���Ŀ�ļ۸�,���ܼ�����';
    Raise Err_Item;
  End If;

  If n_���÷�ʱ�� = 1 Then
    d_Date := To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(d_ʱ�ο�ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss');
  Else
    d_Date := Trunc(����ʱ��_In);
  End If;

  --���¹Һ����״̬
  If ��������_In <> 2 Then
    n_���� := ����_In;
  End If;
  Begin
    Select 1
    Into n_Count
    From �Һ����״̬
    Where Trunc(����) = Trunc(����ʱ��_In) And ���� = ����_In And ��� = n_���� And ״̬ <> 5;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If n_Count = 1 Then
    If n_���÷�ʱ�� = 0 And Nvl(r_����.��ſ���, 0) = 1 Then
      n_���� := Null;
    End If;
    If n_���÷�ʱ�� = 1 And Nvl(r_����.��ſ���, 0) = 1 Then
      v_Err_Msg := '��ǰ����ѱ�ʹ�ã�������ѡ��һ����ţ�';
      Raise Err_Item;
    End If;
  End If;
  n_Count := 0;
  If n_���÷�ʱ�� = 0 And Nvl(r_����.��ſ���, 0) = 1 And n_���� Is Null And ��������_In <> 2 Then
    If �˺�����_In = 1 Then
      Select Nvl(Max(���), 0) + 1
      Into n_����
      From �Һ����״̬
      Where ���� = Trunc(����ʱ��_In) And ���� = r_����.���� And ״̬ Not In (4, 5);
    Else
      Select Nvl(Max(���), 0) + 1
      Into n_����
      From �Һ����״̬
      Where ���� = Trunc(����ʱ��_In) And ���� = r_����.���� And ״̬ <> 5;
    End If;
  End If;
  If n_���÷�ʱ�� = 1 And ��������_In <> 2 Then
  
    If ������ʽ_In > 1 And Nvl(r_����.��ſ���, 0) = 0 Then
      --����:ԤԼʱ�����||ԤԼ��
      If Nvl(n_ԤԼ����, 0) = 0 Then
        v_Temp := Nvl(r_����.��Լ��, 0);
        v_Temp := LTrim(RTrim(v_Temp));
        v_Temp := LPad(Nvl(n_ԤԼ����, 0) + 1, Length(v_Temp), '0');
        v_Temp := n_ԤԼʱ����� || v_Temp;
        n_���� := To_Number(v_Temp);
      Else
        n_���� := n_ԤԼ���� + 1;
      End If;
    End If;
  End If;

  If Nvl(r_����.��ſ���, 0) = 1 Or (������ʽ_In > 1 And n_���÷�ʱ�� = 1) Or �������״̬_In = 1 Then
    --������ŵĴ���
    Begin
      Select ����Ա����, ������
      Into v_��Ų���Ա, v_��Ż�����
      From �Һ����״̬
      Where ״̬ = 5 And ���� = ����_In And Trunc(����) = Trunc(d_Date) And ��� = n_����;
      n_������� := 1;
    Exception
      When Others Then
        v_��Ų���Ա := Null;
        v_��Ż����� := Null;
        n_�������   := 0;
    End;
    If n_������� = 0 Then
      Update �Һ����״̬
      Set ״̬ = Decode(������ʽ_In, 2, 2, 1), ԤԼ = Decode(������ʽ_In, 1, 0, 1), �Ǽ�ʱ�� = Sysdate
      Where ���� = ����_In And ���� = d_Date And ��� = n_���� And ����Ա���� = v_����Ա����;
      If Sql%RowCount = 0 Then
        Begin
          Insert Into �Һ����״̬
            (����, ����, ���, ״̬, ����Ա����, ԤԼ, �Ǽ�ʱ��)
          Values
            (����_In, d_Date, n_����, Decode(������ʽ_In, 2, 2, 1), v_����Ա����, Decode(������ʽ_In, 1, 0, 1), Sysdate);
        
          If n_������λ���� > 0 And ������ʽ_In > 1 And Nvl(n_�Ƿ񿪷�, 0) = 1 Then
            Update ������λ�ҺŻ���
            Set ��Լ�� = ��Լ�� + Decode(������ʽ_In, 2, 1, 0), �ѽ��� = �ѽ��� + Decode(������ʽ_In, 3, 1, 0)
            Where ���� = ����_In And ���� = d_Date And ��� = n_���� And ������λ = ������λ_In;
            If Sql%NotFound Then
              Insert Into ������λ�ҺŻ���
                (����, ����, ���, ������λ, ��Լ��, �ѽ���)
              Values
                (����_In, d_Date, n_����, ������λ_In, Decode(������ʽ_In, 1, 0, 1), Decode(������ʽ_In, 3, 1, 0));
            End If;
          End If;
        Exception
          When Others Then
            v_Err_Msg := '���' || n_���� || '�ѱ�ʹ��,������ѡ��һ�����.';
            Raise Err_Item;
        End;
      End If;
    Else
      If v_����Ա���� <> v_��Ų���Ա Or v_������ <> v_��Ż����� Then
        v_Err_Msg := '���' || n_���� || '�ѱ�������' || v_������ || '����,������ѡ��һ�����.';
        Raise Err_Item;
      Else
        Update �Һ����״̬
        Set ״̬ = Decode(������ʽ_In, 2, 2, 1), ԤԼ = Decode(������ʽ_In, 1, 0, 1), �Ǽ�ʱ�� = Sysdate
        Where ���� = ����_In And Trunc(����) = Trunc(d_Date) And ��� = n_���� And ״̬ = 5 And ����Ա���� = v_����Ա���� And ������ = v_������;
      End If;
    End If;
  End If;

  --�������ݲ������κ� ����
  If ������ʽ_In <> 2 And ��������_In <> 1 Then
    --�Һ�,ԤԼ�Һ��Ѿ��ۿ��
    n_Ԥ��id := Ԥ��id_In;
    If Nvl(n_Ԥ��id, 0) = 0 Then
      Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
    End If;
    n_����ϼ� := 0;
    If ���ս���_In Is Not Null Then
      --�������ս���
      v_�������� := ���ս���_In || '||';
      n_����ϼ� := 0;
      While v_�������� Is Not Null Loop
        v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '||') - 1);
        v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
        n_������ := To_Number(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
        If Nvl(n_������, 0) <> 0 Then
          Insert Into ����Ԥ����¼
            (ID, ��¼����, NO, ��¼״̬, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, ��������)
          Values
            (n_Ԥ��id, 4, ���ݺ�_In, 1, Decode(����id_In, 0, Null, ����id_In), '���ս���', v_���㷽ʽ, d_�Ǽ�ʱ��, v_����Ա���, v_����Ա����, n_������,
             n_����id, n_��id, n_����id, 4);
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
        End If;
        n_����ϼ� := Nvl(n_����ϼ�, 0) + Nvl(n_������, 0);
        v_�������� := Substr(v_��������, Instr(v_��������, '||') + 2);
      End Loop;
    End If;
  
    If Nvl(��Ԥ��_In, 0) <> 0 Then
      --������Ԥ��
      n_����ϼ� := n_����ϼ� + Nvl(��Ԥ��_In, 0);
      n_Ԥ����� := ��Ԥ��_In;
      For r_Deposit In c_Deposit(����id_In) Loop
        n_������ := Case
                    When r_Deposit.��� - n_Ԥ����� < 0 Then
                     r_Deposit.���
                    Else
                     n_Ԥ�����
                  End;
        If r_Deposit.Id <> 0 Then
          --��һ�γ�Ԥ��(82592,����һ�α��Ͻ���ID,��Ԥ�����Ϊ0)
          Update ����Ԥ����¼ Set ��Ԥ�� = 0, ����id = n_����id, �������� = 4 Where ID = r_Deposit.Id;
        End If;
        --���ϴ�ʣ���
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���,
           ��Ԥ��, ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, ��������)
          Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�,
                 d_�Ǽ�ʱ��, v_����Ա����, v_����Ա���, n_������, n_����id, n_��id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, n_����id, 4
          From ����Ԥ����¼
          Where NO = r_Deposit.No And ��¼״̬ = r_Deposit.��¼״̬ And ��¼���� In (1, 11) And Rownum = 1;
      
        --���²���Ԥ�����
        Update �������
        Set Ԥ����� = Nvl(Ԥ�����, 0) - n_������
        Where ����id = ����id_In And ���� = 1 And ���� = Nvl(r_Deposit.Ԥ�����, 2);
        --����Ƿ��Ѿ�������
        If r_Deposit.��� <= n_������ Then
          n_Ԥ����� := n_Ԥ����� - r_Deposit.���;
        Else
          n_Ԥ����� := 0;
        End If;
        If n_Ԥ����� = 0 Then
          Exit;
        End If;
      End Loop;
      If n_Ԥ����� > 0 Then
        v_Err_Msg := 'Ԥ������֧������֧�����,���ܼ���������';
        Raise Err_Item;
      End If;
    End If;
    --ʣ�����,��ָ�����㷽֧��
    n_������ := Nvl(n_ʵ�ս��ϼ�, 0) - Nvl(n_����ϼ�, 0);
    If Nvl(n_������, 0) < 0 Then
      v_Err_Msg := '�Һŵ���ؽ�������˵�ǰʵ����,���ܼ���������';
      Raise Err_Item;
    End If;
    If Nvl(n_������, 0) <> 0 Or (Nvl(n_������, 0) = 0 And Nvl(��Ԥ��_In, 0) = 0) Then
      If ���㷽ʽ_In Is Null Then
        v_Err_Msg := 'δ����ָ���Ľ��㷽ʽ,���ܼ���������';
        Raise Err_Item;
      End If;
    
      If Nvl(Ԥ��id_In, 0) <> 0 Then
        --�����Ԥ��ID_In��Ҫ��Ϊ�˽����������,���ҽ������վ���˸�ID,��Ҫ���µ�ID���и���,����������ת���ID
        Update ����Ԥ����¼ Set ID = n_Ԥ��id Where ID = Nvl(Ԥ��id_In, 0);
        n_Ԥ��id := Nvl(Ԥ��id_In, 0);
      End If;
    
      Insert Into ����Ԥ����¼
        (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, ������ˮ��, ����˵��, �������, ������λ, �����id, ����,
         ��������)
      Values
        (n_Ԥ��id, 4, 1, ���ݺ�_In, r_Pati.����id, ���㷽ʽ_In, Nvl(n_������, 0), d_�Ǽ�ʱ��, v_����Ա���, v_����Ա����, n_����id, ������λ_In || '�ɿ�',
         n_��id, ������ˮ��_In, ����˵��_In, n_����id, ������λ_In, �����id_In, ֧������_In, 4);
    End If;
  
    --������Ա�ɿ�����
  
    For v_�ɿ� In (Select ���㷽ʽ, Sum(Nvl(a.��Ԥ��, 0)) As ��Ԥ��
                 From ����Ԥ����¼ A
                 Where a.����id = n_����id And Mod(a.��¼����, 10) <> 1 And Nvl(����id, 0) = Nvl(����id_In, 0)
                 Group By ���㷽ʽ) Loop
    
      Update ��Ա�ɿ����
      Set ��� = Nvl(���, 0) + Nvl(v_�ɿ�.��Ԥ��, 0)
      Where �տ�Ա = v_����Ա���� And ���� = 1 And ���㷽ʽ = v_�ɿ�.���㷽ʽ
      Returning ��� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ��Ա�ɿ����
          (�տ�Ա, ���㷽ʽ, ����, ���)
        Values
          (v_����Ա����, v_�ɿ�.���㷽ʽ, 1, Nvl(v_�ɿ�.��Ԥ��, 0));
        n_����ֵ := Nvl(v_�ɿ�.��Ԥ��, 0);
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From ��Ա�ɿ����
        Where �տ�Ա = v_����Ա���� And ���㷽ʽ = ���㷽ʽ_In And ���� = 1 And Nvl(���, 0) = 0;
      End If;
    
    End Loop;
  
  End If;

  --����Һż�¼
  If ��������_In = 2 Then
    Begin
      Select ID Into n_�Һ�id From ���˹Һż�¼ Where����¼״̬ = 0 And NO = ���ݺ�_In And ����id = ����id_In;
    Exception
      When Others Then
        Null;
    End;
  Else
    Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
  End If;

  Update ���˹Һż�¼
  Set ��¼���� = Decode(������ʽ_In, 2, 2, 1), ��¼״̬ = Decode(��������_In, 1, 0, 1), ����� = r_Pati.�����, ����Ա���� = v_����Ա����,
      ����Ա��� = v_����Ա���, ԤԼ = Decode(������ʽ_In, 1, 0, 1), ������ = Decode(��������_In, 1, Null, Decode(������ʽ_In, 2, Null, v_����Ա����)),
      ����ʱ�� = Case ��������_In
                When 1 Then
                 Null
                Else
                 Case ������ʽ_In
                   When 2 Then
                    Null
                   Else
                    d_�Ǽ�ʱ��
                 End
              End, ������ˮ�� = Nvl(������ˮ��_In, ������ˮ��), ����˵�� = Nvl(����˵��_In, ����˵��), ������λ = Nvl(������λ_In, ������λ),
      ԤԼ����Ա = Decode(������ʽ_In, 1, Nvl(ԤԼ����Ա, Null), Nvl(ԤԼ����Ա, v_����Ա����)),
      ԤԼ����Ա��� = Decode(������ʽ_In, 1, Nvl(ԤԼ����Ա���, Null), Nvl(ԤԼ����Ա���, v_����Ա���))
  Where ID = n_�Һ�id;
  If Sql%NotFound Then
    Begin
      Select ���� Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = r_Pati.���ʽ And Rownum < 2;
    Exception
      When Others Then
        v_���ʽ := Null;
    End;
    Insert Into ���˹Һż�¼
      (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ԤԼʱ��, ����Ա���,
       ����Ա����, ����, ����, ԤԼ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ҽ�Ƹ��ʽ, ԤԼ����Ա, ԤԼ����Ա���)
    Values
      (n_�Һ�id, ���ݺ�_In, Decode(������ʽ_In, 2, 2, 1), Decode(��������_In, 1, 0, 1), r_Pati.����id, r_Pati.�����, r_Pati.����,
       r_Pati.�Ա�, r_Pati.����, ����_In, 0, v_����, Null, r_����.����id, r_����.ҽ������, 0, Null, d_�Ǽ�ʱ��, ����ʱ��_In,
       Case When(Nvl(������ʽ_In, 0)) = 1 Then Null Else ����ʱ��_In End, v_����Ա���, v_����Ա����, 0, n_����, Decode(������ʽ_In, 1, 0, 1),
       Decode(������ʽ_In, 2, Null, v_����Ա����), Decode(������ʽ_In, 2, To_Date(Null), d_�Ǽ�ʱ��), ������ˮ��_In, ����˵��_In, ������λ_In, v_���ʽ,
       Decode(������ʽ_In, 1, Null, v_����Ա����), Decode(������ʽ_In, 1, Null, v_����Ա���));
  End If;
  --�������ݲ��ܲ�������
  If ��������_In <> 1 Then
    n_ԤԼ���ɶ��� := 0;
    If ������ʽ_In > 1 Then
      n_ԤԼ���ɶ��� := Zl_To_Number(zl_GetSysParameter('ԤԼ���ɶ���', 1113));
    End If;
    --�Һź��շѵ�ԤԼ��ֱ�ӽ������(�շ�ԤԼȱ�ٽ��չ���,����ֱ�Ӻ͹Һ�һ��ֱ�ӽ������)
    If ������ʽ_In <> 2 Or n_ԤԼ���ɶ��� = 1 Then
      If Zl_To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113)) <> 0 Then
        --�Ŷӽк�ģʽ:-0-����������;1-��ҽ�������̨�Ŷ�;2-�ȷ���,��ҽ��վ      
        If Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113)) = 0 Or n_ԤԼ���ɶ��� = 1 Then
          --��������
          --.����ִ�в��š� �ķ�ʽ���ɶ���
          v_�������� := r_����.����id;
          v_�ŶӺ��� := Zlgetnextqueue(r_����.����id, n_�Һ�id, ����_In || '|' || ����_In);
          --�Һ�id_In,����_In,����_In,ȱʡ����_In,��չ_In(������)
          v_�Ŷ���� := Zlgetsequencenum(0, n_�Һ�id, 0);
          --�Һ�id_In,����_In,����_In,ȱʡ����_In,��չ_In(������)  
          d_�Ŷ�ʱ�� := Zl_Get_Queuedate(n_�Һ�id, ����_In, ����_In, d_Date);
          --  ��������_In , ҵ������_In, ҵ��id_In,����id_In,�ŶӺ���_In,v_�Ŷӱ��,��������_In,����ID_IN, ����_In, ҽ������_In, �Ŷ�ʱ��_In
          Zl_�ŶӽкŶ���_Insert(v_��������, 0, n_�Һ�id, r_����.����id, v_�ŶӺ���, Null, r_Pati.����, r_Pati.����id, v_����, r_����.ҽ������,
                           d_�Ŷ�ʱ��, ԤԼ��ʽ_In, n_���÷�ʱ��, v_�Ŷ����);
        End If;
      End If;
    End If;
  
    If Nvl(������ʽ_In, 0) = 1 Then
      --����Ʊ��ʹ�����
      If Ʊ�ݺ�_In Is Not Null Then
        Select Ʊ�ݴ�ӡ����_Id.Nextval Into n_��ӡid From Dual;
        --����Ʊ��
        Insert Into Ʊ�ݴ�ӡ���� (ID, ��������, NO) Values (n_��ӡid, 4, ���ݺ�_In);
        Insert Into Ʊ��ʹ����ϸ
          (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����)
        Values
          (Ʊ��ʹ����ϸ_Id.Nextval, Decode(�շ�Ʊ��_In, 1, 1, 4), Ʊ�ݺ�_In, 1, 1, ����id_In, n_��ӡid, d_�Ǽ�ʱ��, v_����Ա����);
        --״̬�Ķ�
        Update Ʊ�����ü�¼
        Set ��ǰ���� = Ʊ�ݺ�_In, ʣ������ = Decode(Sign(ʣ������ - 1), -1, 0, ʣ������ - 1), ʹ��ʱ�� = Sysdate
        Where ID = Nvl(����id_In, 0);
      End If;
      --���˱��ξ���(�Է���ʱ��Ϊ׼)
      If Nvl(r_Pati.����id, 0) <> 0 Then
        Update ������Ϣ Set ����ʱ�� = ����ʱ��_In, ����״̬ = 1, �������� = v_���� Where ����id = r_Pati.����id;
      End If;
    End If;
  End If;
  --���˹ҺŻ���
  --��������ʱ�����ٶԻ��ܵ��ݽ���ͳ���� ����������ʱ�Ѿ������˻���
  If ��������_In <> 2 Then
    --������ʽ_IN:1-��ʾ�Һ�,2-��ʾԤԼ�ҺŲ��ۿ�,3-��ʾԤԼ�Һ�,�ۿ�
    --�Ƿ�ΪԤԼ����:0-��ԤԼ�Һ�; 1-ԤԼ�Һ�,2-ԤԼ����;3-�շ�ԤԼ
    n_ԤԼ := Case
              When Nvl(������ʽ_In, 0) = 1 Then
               0
              When Nvl(������ʽ_In, 0) = 2 Then
               1
              When Nvl(������ʽ_In, 0) = 3 Then
               3
              Else
               0
            End;
    Zl_���˹ҺŻ���_Update(r_����.ҽ������, r_����.ҽ��id, r_����.��Ŀid, r_����.����id, ����ʱ��_In, n_ԤԼ, ����_In);
  End If;
  --��Ϣ����
  Begin
    Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
      Using 1, n_�Һ�id;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]+' || v_Err_Msg || '+[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���������Һ�_Insert;
/

--91633:��ΰ��,2015-12-25,�ٴ�·��ȡ��������Ŀ��δ�����κ���Ŀ��
Create Or Replace Procedure Zl_����·������_Insert
(
  ���_In           Number, --ҽ���������·����Ŀʱ�����Ϊ0
  ����id_In         �����ٴ�·��.����id%Type,
  ��ҳid_In         �����ٴ�·��.��ҳid%Type,
  Ӥ��_In           ���Ӳ�����¼.Ӥ��%Type,
  ����id_In         �����ٴ�·��.����id%Type,
  ·����¼id_In     ����·��ִ��.·����¼id%Type,
  �׶�id_In         ����·��ִ��.�׶�id%Type,
  ����_In           ����·��ִ��.����%Type,
  ����_In           ����·��ִ��.����%Type,
  ����_In           ����·��ִ��.����%Type,
  ��Ŀid_In         ����·��ִ��.��Ŀid%Type,
  ҽ��ids_In        Varchar2,
  �����ļ�ids_In    Varchar2,
  ���˲���ids_In    Varchar2,
  �Ǽ���_In         ����·��ִ��.�Ǽ���%Type,
  �Ǽ�ʱ��_In       ����·��ִ��.�Ǽ�ʱ��%Type,
  ��Ŀ����_In       ����·��ִ��.��Ŀ����%Type := Null,
  ִ����_In         ����·��ִ��.ִ����%Type := Null,
  ��Ŀ���_In       ����·��ִ��.��Ŀ���%Type := Null,
  ͼ��id_In         ����·��ִ��.ͼ��id%Type := Null,
  ���ԭ��_In       ����·��ִ��.���ԭ��%Type := Null,
  ����ԭ��_In       ����·��ִ��.����ԭ��%Type := Null,
  �Զ�ִ��_In       Number := 0,
  ���Ӳ���id_In     ���Ӳ�����¼.Id%Type := Null,
  �ϲ�·���׶�s_In  Varchar2 := Null, --�����޸ĺϲ�·���ĵ�ǰ�׶�ID����ʽ���ϲ�·����¼ID:�׶�ID,�ϲ�·����¼ID:�׶�ID��������
  �ϲ�·����¼id_In ����·��ִ��.�ϲ�·����¼id%Type := Null,
  �ϲ�·���׶�id_In ����·��ִ��.�ϲ�·���׶�id%Type := Null,
  ����λ��id_In     ����·��ִ��.Id%Type := 0,
  ������_In         ����·��ִ��.������ %Type := 1,
  ����ids_In        Varchar2 := Null,
  ����ʱ������_In   ����·��ִ��.����ʱ������%Type := Null --1-��¼,2-�ݴ�
) Is
  v_��ǰ�׶�id �����ٴ�·��.��ǰ�׶�id%Type;
  v_·��ִ��id ����·��ִ��.Id%Type;
  v_����id     ���Ӳ�����¼.Id%Type;
  t_Advice     t_Numlist;
  t_File       t_Numlist;
  t_Doc        t_Numlist;

  v_Id             ���Ӳ�������.Id%Type;
  v_��id           ���Ӳ�������.��id%Type;
  v_��ǰ��id       ���Ӳ�������.��id%Type;
  v_ԭ�������     ���Ӳ�������.��id%Type;
  v_�����ı�       ���Ӳ�������.�����ı�%Type;
  v_ִ�л���       Varchar2(20);
  n_��ǰ����       �����ٴ�·��.��ǰ����%Type;
  n_�ϲ�·����¼id ����·��ִ��.�ϲ�·����¼id%Type;
  n_�ϲ�·���׶�id ����·��ִ��.�ϲ�·���׶�id%Type;
  n_����           �����ٴ�·��.��ǰ����%Type;
  v_�ϲ�·���׶�s  Varchar2(255);

  v_��Ŀ��� ����·��ִ��.��Ŀ���%Type;
  n_Count    Number;
  n_Minnum   Number;
  v_Error    Varchar2(255);
  Err_Custom Exception;

  --��Ŀ��Ŵ���
  Procedure p_Sort_��Ŀ���
  (
    ��Ŀ���_In In ����·��ִ��.��Ŀ���%Type,
    ִ��id_In   In ����·��ִ��.Id%Type
  ) Is
    n_Num Number;
  Begin
    n_Num := ��Ŀ���_In;
    For r_Outpathitem In (Select a.Id, Nvl(a.��Ŀ���, b.��Ŀ���) As ��Ŀ���
                          From ����·��ִ�� A, �ٴ�·����Ŀ B
                          Where a.·����¼id = ·����¼id_In And a.�׶�id = �׶�id_In And a.���� = ����_In And a.���� = ����_In And
                                a.��Ŀid = b.Id(+) And Nvl(a.��Ŀ���, b.��Ŀ���) >= ��Ŀ���_In
                          Order By Nvl(a.��Ŀ���, b.��Ŀ���)) Loop
      n_Num := n_Num + 1;
      --1-�Ӳ���λ�ô�֮�������·������Ŀ��ż� 1
      Update ����·��ִ�� A Set a.��Ŀ��� = n_Num Where a.Id = r_Outpathitem.Id;
    End Loop;
    Update ����·��ִ�� A Set a.��Ŀ��� = ��Ŀ���_In Where a.Id = ִ��id_In;
  Exception
    When Others Then
      Null;
  End p_Sort_��Ŀ���;
Begin
  If ���_In = 1 And (��Ŀ����_In Is Null Or ��Ŀ����_In = 'δ�����κ���Ŀ' Or ��Ŀ����_In = '·������Ŀ') Then
    --�ϲ�·��
    If �ϲ�·���׶�s_In Is Not Null Then
      Select Nvl(��ǰ����, 1) Into n_��ǰ���� From �����ٴ�·�� Where ID = ·����¼id_In;
      --�������(��Ҫ·����ǰ�ϲ�·������ǰ����Ҫ·���Ӻ󣬺ϲ�·�����Ӻ�)
      n_����          := ����_In - n_��ǰ����;
      v_�ϲ�·���׶�s := �ϲ�·���׶�s_In || ',';
      While v_�ϲ�·���׶�s Is Not Null Loop
        n_�ϲ�·����¼id := To_Number(Substr(v_�ϲ�·���׶�s, 1, Instr(v_�ϲ�·���׶�s, ':') - 1));
        n_�ϲ�·���׶�id := To_Number(Substr(v_�ϲ�·���׶�s, Instr(v_�ϲ�·���׶�s, ':') + 1,
                                       Instr(v_�ϲ�·���׶�s, ',') - Instr(v_�ϲ�·���׶�s, ':') - 1));
        Select Nvl(��ǰ�׶�id, 0) Into v_��ǰ�׶�id From ���˺ϲ�·�� Where ID = n_�ϲ�·����¼id;
        If v_��ǰ�׶�id <> n_�ϲ�·���׶�id Then
          Update ���˺ϲ�·�� Set ǰһ�׶�id = ��ǰ�׶�id, ��ǰ�׶�id = n_�ϲ�·���׶�id Where ID = n_�ϲ�·����¼id;
        End If;
        Update ���˺ϲ�·�� Set ��ǰ���� = Nvl(��ǰ����, 1) + n_���� Where ID = n_�ϲ�·����¼id;
      
        v_�ϲ�·���׶�s := Substr(v_�ϲ�·���׶�s, Instr(v_�ϲ�·���׶�s, ',') + 1);
      End Loop;
    End If;
    --��Ҫ·��
    If ������_In = 1 Then
      Select Nvl(��ǰ�׶�id, 0) Into v_��ǰ�׶�id From �����ٴ�·�� Where ID = ·����¼id_In;
      If v_��ǰ�׶�id <> �׶�id_In Then
        Update �����ٴ�·�� Set ǰһ�׶�id = ��ǰ�׶�id, ��ǰ�׶�id = �׶�id_In Where ID = ·����¼id_In;
      End If;
      Update �����ٴ�·�� Set ��ǰ���� = ����_In Where ID = ·����¼id_In;
    End If;
  End If;

  --��ӵ�·������Ŀ:��ʹ�п�ѡ����Ŀ���ܻ�δ����,���ռ�ü����油����Ŀ������
  If ��Ŀ����_In Is Not Null Then
    Select Max(Nvl(a.��Ŀ���, b.��Ŀ���)) + 1
    Into v_��Ŀ���
    From ����·��ִ�� A, �ٴ�·����Ŀ B
    Where a.·����¼id = ·����¼id_In And a.�׶�id = �׶�id_In And a.���� = ����_In And a.���� = ����_In And a.��Ŀid = b.Id(+);
  End If;

  v_·��ִ��id := 0;
  If ���_In = 0 And ��Ŀ����_In Is Null Then
    --��max��Ϊ���ݴ���ǰ�����ݣ�ʵ����ͬһ��Ŀ�ڵ���ֻ��һ��ִ�м�¼
    Select Nvl(Max(ID), 0)
    Into v_·��ִ��id
    From ����·��ִ��
    Where ·����¼id = ·����¼id_In And �׶�id = �׶�id_In And ���� = ����_In And ��Ŀid = ��Ŀid_In;
  End If;

  --ҽ��������ӵķ�·������Ŀ
  If v_·��ִ��id = 0 Then
    Select Count(1)
    Into n_Count
    From ����·��ִ��
    Where ·����¼id = ·����¼id_In And �׶�id = �׶�id_In And ���� = ����_In;
    If n_Count = 0 Then
      --�״�����·����Ŀ��ת��ǰһ����ݴ���Ŀ
      Update ����·��ִ��
      Set �׶�id = �׶�id_In, ���� = ����_In, ���� = ����_In, ��Ŀ��� = Null
      Where ID In (Select ID From ����·��ִ�� Where ·����¼id = ·����¼id_In And ����ʱ������ = 2);
      --�޸��ݴ��ʶ
      Update ����·��ִ��
      Set ����ʱ������ = Null
      Where ID In (Select a.Id
                   From ����·��ִ�� A, ����·��ҽ�� B, ����ҽ����¼ C
                   Where a.Id = b.·��ִ��id And b.����ҽ��id = c.Id And a.·����¼id = ·����¼id_In And a.����ʱ������ = 2 And
                         a.���� = Trunc(c.��ʼִ��ʱ��));
    End If;
    Select ����·��ִ��_Id.Nextval Into v_·��ִ��id From Dual;
    Insert Into ����·��ִ��
      (ID, ·����¼id, �׶�id, ����, ����, ����, ��Ŀid, �Ǽ���, �Ǽ�ʱ��, ��Ŀ���, ��Ŀ����, ִ����, ������, ��Ŀ���, ͼ��id, ���ԭ��, ����ԭ��, �ϲ�·����¼id, �ϲ�·���׶�id,
       ����ʱ������)
    Values
      (v_·��ִ��id, ·����¼id_In, �׶�id_In, ����_In, ����_In, ����_In, ��Ŀid_In, �Ǽ���_In, �Ǽ�ʱ��_In, v_��Ŀ���, ��Ŀ����_In, ִ����_In, ������_In,
       ��Ŀ���_In, ͼ��id_In, ���ԭ��_In, ����ԭ��_In, �ϲ�·����¼id_In, �ϲ�·���׶�id_In, ����ʱ������_In);
  
    --·������Ŀ��Ų��� ����
    If ����λ��id_In <> 0 Then
      --��ȡҪ��������
      Select Nvl(a.��Ŀ���, b.��Ŀ���)
      Into v_��Ŀ���
      From ����·��ִ�� A, �ٴ�·����Ŀ B
      Where a.Id = ����λ��id_In And a.��Ŀid = b.Id(+);
      --��ŵ���
      p_Sort_��Ŀ���(v_��Ŀ���, v_·��ִ��id);
    End If;
    --·����Ŀ��������ʱ,�������:�����ٴ�·����Ŀ����A1,A2,A3��3����Ŀ,�״�����A1,A2��,������·������ĿB1,B2,ͬʱ��B1,B2���뵽A1��λ��
    --         ��ô��ʱ����·��ִ���е���ű�Ϊ:B1(1),B2(2),A1(3),A2(4),����ٲ�������A3ʱ,·����ʾ˳���Ϊ��B1(1),B2(2),A1(3),A3(3),A2(4)
    --         �����ͻ����·����Ŀ�в������ɵ�A3���ܰ����ٴ�·����Ŀ��˳��A1,A2,A3 ��ȷ����
  
    --��ǰ�׶Σ���ǰ��������ǰ�����£�����·������Ŀ��·���ڵ���Ŀ��ű����µ���������δ���·������Ŀʱ��·������Ŀ�����Ϊ�գ�
    Select Nvl(Count(ID), 0)
    Into n_Count
    From ����·��ִ��
    Where ·����¼id = ·����¼id_In And �׶�id = �׶�id_In And ���� = ����_In And ���� = ����_In And ��Ŀid Is Not Null And ��Ŀ��� Is Not Null;
    --�������ɵ�·����Ŀ�������
    If n_Count > 0 And ��Ŀid_In Is Not Null Then
      --���Ҳ������ɵ�·����Ŀ,Ӧ�ò����λ��
      Select Min(b.��Ŀ���)
      Into n_Minnum
      From ����·��ִ�� A, �ٴ�·����Ŀ B
      Where a.·����¼id = ·����¼id_In And a.�׶�id = �׶�id_In And a.���� = ����_In And a.���� = ����_In And a.��Ŀid = b.Id;
    
      Select ��Ŀ��� Into v_��Ŀ��� From �ٴ�·����Ŀ Where ID = ��Ŀid_In;
      --ȷ����·����Ŀ��Ų����λ�ã�
      If v_��Ŀ��� = n_Minnum Then
        --v_��Ŀ��� = n_Minnum������·��ִ�м�¼���ڴ�����ǰ���뵽���ݿ⣬������������ݾ�����С����������
        Select ��Ŀ���
        Into v_��Ŀ���
        From (Select Nvl(a.��Ŀ���, b.��Ŀ���) As ��Ŀ���
               From ����·��ִ�� A, �ٴ�·����Ŀ B
               Where a.·����¼id = ·����¼id_In And a.�׶�id = �׶�id_In And a.���� = ����_In And a.���� = ����_In And a.��Ŀid = b.Id And
                     b.��Ŀ��� > n_Minnum
               Order By b.��Ŀ���)
        Where Rownum = 1;
      Else
        Select ��Ŀ���
        Into v_��Ŀ���
        From (Select Nvl(a.��Ŀ���, b.��Ŀ���) As ��Ŀ���
               
               From ����·��ִ�� A, �ٴ�·����Ŀ B
               Where a.·����¼id = ·����¼id_In And a.�׶�id = �׶�id_In And a.���� = ����_In And a.���� = ����_In And a.��Ŀid = b.Id And
                     b.��Ŀ��� < v_��Ŀ���
               Order By b.��Ŀ��� Desc)
        Where Rownum = 1;
        v_��Ŀ��� := v_��Ŀ��� + 1;
      End If;
      p_Sort_��Ŀ���(v_��Ŀ���, v_·��ִ��id);
    End If;
  
    --������Զ�ִ��ģʽ��������ǰ����׶�ʱ���ã�;��¼·������Ŀ
    If �Զ�ִ��_In = 1 Then
      Select zl_GetSysParameter('�Ƿ�����·��ִ�л���', 1256) Into v_ִ�л��� From Dual;
      If v_ִ�л��� = '1' Then
        Select zl_GetSysParameter('·��ִ�л������ó���', 1256) Into v_ִ�л��� From Dual;
      
        Select Nvl(Nvl(a.ִ����, b.ִ����), 0)
        Into n_Count
        From ����·��ִ�� A, �ٴ�·����Ŀ B
        Where a.��Ŀid = b.Id(+) And a.Id = v_·��ִ��id;
        --��ǰִ���߷������ó����Զ�ִ��,��ִ����ȡ����ֵʱ,ͳһ����
        If n_Count = 0 Or Substr(v_ִ�л���, n_Count, 1) = '1' Then
          Update ����·��ִ��
          Set ִ���� = �Ǽ���_In, ִ��ʱ�� = �Ǽ�ʱ��_In, ִ�н�� = '�Ѿ�ִ��', ִ��˵�� = '�Զ�ִ�С�'
          Where ID = v_·��ִ��id;
        End If;
      End If;
    End If;
  End If;
  --ɾ��������Ŀ��δ�����κ���Ŀ�������ǰ�׶Σ���ǰ���ڴ���������Ŀ����ɾ����δ�����κ���Ŀ����
  Select Count(ID)
  Into n_Count
  From ����·��ִ�� T
  Where t.·����¼id = ·����¼id_In And t.�׶�id = �׶�id_In And t.���� = ����_In And NVL(t.��Ŀ����,'·������Ŀ') = 'δ�����κ���Ŀ';

  If n_Count > 0 Then
    Select Count(ID)
    Into n_Count
    From ����·��ִ�� T
    Where t.·����¼id = ·����¼id_In And t.�׶�id = �׶�id_In And t.���� = ����_In And NVL(t.��Ŀ����,'·������Ŀ') <> 'δ�����κ���Ŀ';
    If n_Count > 0 Then
      Delete From ����·��ִ�� T
      Where t.·����¼id = ·����¼id_In And t.�׶�id = �׶�id_In And t.���� = ����_In And NVL(t.��Ŀ����,'·������Ŀ') = 'δ�����κ���Ŀ';
    End If;
  End If;

  If ҽ��ids_In Is Not Null Then
    Select Column_Value Bulk Collect Into t_Advice From Table(f_Num2list(ҽ��ids_In));
    Forall I In 1 .. t_Advice.Count
      Insert Into ����·��ҽ�� (·��ִ��id, ����ҽ��id) Values (v_·��ִ��id, t_Advice(I));
  End If;

  If ���˲���ids_In Is Not Null Then
    Select Column_Value Bulk Collect Into t_Doc From Table(f_Num2list(���˲���ids_In));
    Select Column_Value Bulk Collect Into t_File From Table(f_Num2list(�����ļ�ids_In));
    For I In 1 .. t_Doc.Count Loop
      v_����id := t_Doc(I);
    
      Insert Into ���Ӳ�����¼
        (ID, ������Դ, ����id, ��ҳid, Ӥ��, ����id, ��������, �ļ�id, ��������, ������, ����ʱ��, ������, ����ʱ��, ���汾, ǩ������, �༭��ʽ, ·��ִ��id)
        Select v_����id, 2, ����id_In, ��ҳid_In, Ӥ��_In, ����id_In, ����, ID, ����, �Ǽ���_In, �Ǽ�ʱ��_In, �Ǽ���_In, �Ǽ�ʱ��_In, 1, 0, Decode(����,2,1,0),
               v_·��ִ��id
        From �����ļ��б�
        Where ID = t_File(I);

      For Rs In (Select ID, �ļ�id, Nvl(��id, 0) As ��id, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���, Ԥ�����id, �������, ʹ��ʱ��,
                        ����Ҫ��id, �滻��, Ҫ������, Ҫ������, Ҫ�س���, Ҫ��С��, Ҫ�ص�λ, Ҫ�ر�ʾ, ������̬, Ҫ��ֵ��
                 From �����ļ��ṹ
                 Where �ļ�id = t_File(I)
                 Order By �������) Loop

        Select ���Ӳ�������_Id.Nextval Into v_Id From Dual;
      
        If Rs.��id = 0 Then
          v_��ǰ��id := v_Id;
          v_��id     := Null;
        Else
          --�������Ϊ�յ�ʱ�򣬸�ID�Ͳ��ǰ���˳����ˣ���Ҫ���²���
          If Rs.������� Is Null Then
            Select ������� Into v_ԭ������� From �����ļ��ṹ Where ID = Rs.��id;
            If v_ԭ������� Is Null Then
              v_��id := Null;
            Else
              Select ID Into v_��id From ���Ӳ������� Where �ļ�id = v_����id And ������� = v_ԭ�������;
            End If;
          Else
            v_��id := v_��ǰ��id;
          End If;
        End If;
      
        If Rs.�������� = 4 And Rs.�滻�� = 1 Then
          v_�����ı� := Zl_Replace_Element_Value(Rs.Ҫ������, ����id_In, ��ҳid_In, 2, Null, Ӥ��_In);
        Else
          v_�����ı� := Rs.�����ı�;
        End If;
      
        Insert Into ���Ӳ�������
          (ID, �ļ�id, ��ʼ��, ��ֹ��, ��id, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���, Ԥ�����id, �������id, �������, ʹ��ʱ��, ����Ҫ��id,
           �滻��, Ҫ������, Ҫ������, Ҫ�س���, Ҫ��С��, Ҫ�ص�λ, Ҫ�ر�ʾ, ������̬, Ҫ��ֵ��)
        Values
          (v_Id, v_����id, 1, 0, v_��id, Rs.�������, Rs.��������, Rs.������, Rs.��������, Rs.��������, Rs.�����д�, v_�����ı�, Rs.�Ƿ���, Rs.Ԥ�����id,
           Decode(Rs.��id, 0, Rs.Id, Null), Rs.�������, Rs.ʹ��ʱ��, Rs.����Ҫ��id, Rs.�滻��, Rs.Ҫ������, Rs.Ҫ������, Rs.Ҫ�س���, Rs.Ҫ��С��,
           Rs.Ҫ�ص�λ, Rs.Ҫ�ر�ʾ, Rs.������̬, Rs.Ҫ��ֵ��);
      
        If Rs.�������� = 5 Then
          Insert Into ���Ӳ���ͼ�� (����id, ͼ��) Values (v_Id, (Select ͼ�� From �����ļ�ͼ�� Where ����id = Rs.Id));
        End If;
      
      End Loop;
    
      Insert Into ���Ӳ�����ʽ
        (�ļ�id, ����)
      Values
        (v_����id, (Select ���� From �����ļ���ʽ Where �ļ�id = t_File(I)));
    End Loop;
  End If;

  If Nvl(���Ӳ���id_In, 0) <> 0 Then
    Update ���Ӳ�����¼ Set ·��ִ��id = v_·��ִ��id Where ID = ���Ӳ���id_In;
  End If;
  If ����ids_In Is Not Null Then
    For Rs In (Select /*+ Rule*/
                Column_Value As ����id
               From Table(Cast(f_Str2list(����ids_In, ',') As Zltools.t_Strlist))) Loop
      Insert Into ����·������ (·��ִ��id, ����id) Values (v_·��ִ��id, Rs.����id);
    End Loop;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����·������_Insert;
/

--92410:������,2016-01-05,����״̬����
--91842:������,2016-01-18,����תסԺ����
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
  
    Select Count(NO), Sum(ʵ�ս��) Into n_Count, n_ʵ�ս�� From ������ü�¼ Where NO = No_In And ��¼���� = 1;
    If n_Count = 0 Or n_ʵ�ս�� = 0 Then
      v_Err_Msg := '����' || No_In || '�����շѵ��ݻ��򲢷�ԭ�����˲����˸õ���,����תΪסԺ����.';
      Raise Err_Item;
    End If;
  
    Select ����id, ����id, ��������id, ������
    Into n_ԭ����id, n_����id, n_��������id, v_������
    From ������ü�¼
    Where NO = No_In And ��¼���� = 1 And ��¼״̬ In (1, 3) And Rownum < 2;
  
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
      Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 1;
  
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
                                                     a.����id = b.����id))) Loop
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
      For r_Prepay In (Select NO, ʵ��Ʊ��, ����id, ��ҳid, ����id, ���㷽ʽ, �������, �ɿλ, ��λ������, ��λ�ʺ�, Sum(��Ԥ��) As ��Ԥ��, �����id, ���㿨���,
                              ����, ������ˮ��, ����˵��, ������λ
                       From ����Ԥ����¼ A
                       Where ��¼���� In (1, 11) And a.����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids)))
                       Group By n_Tempid, NO, ʵ��Ʊ��, ����id, ��ҳid, ����id, ���㷽ʽ, �������, �ɿλ, ��λ������, ��λ�ʺ�, �����id, ���㿨���, ����,
                                ������ˮ��, ����˵��, ������λ) Loop
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
      Update ������ü�¼ Set ��¼״̬ = 3 Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 1;
      For r_Clinic In (Select ���, ��������, �۸񸸺�, ����id, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ���ձ���, ��������,
                              ��ҩ����, ����, Sum(����) As ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, Sum(Ӧ�ս��) As Ӧ�ս��,
                              Sum(ʵ�ս��) As ʵ�ս��, Sum(ͳ����) As ͳ����, ��������id, ������, ִ�в���id, ������, Max(���ʵ�id) As ���ʵ�id, ����ʱ��,
                              ʵ��Ʊ��
                       From ������ü�¼
                       Where NO = No_In And ��¼���� = 1 And ��¼״̬ In (2, 3) And Nvl(���ӱ�־, 0) Not In (8, 9)
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
                  ���㷽ʽ = v_���㷽ʽ;
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
      For r_Prepay In (Select NO, ʵ��Ʊ��, ����id, ��ҳid, ����id, ���㷽ʽ, �������, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, -1 * Sum(��Ԥ��) As ��Ԥ��,
                              �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������
                       From ����Ԥ����¼ A
                       Where ��¼���� In (1, 11) And a.����id In (Select Column_Value From Table(f_Str2list(v_����ids))) And
                             Nvl(��Ԥ��, 0) <> 0
                       Group By n_Tempid, NO, ʵ��Ʊ��, ����id, ��ҳid, ����id, ���㷽ʽ, �������, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, �����id, ���㿨���,
                                ����, ������ˮ��, ����˵��, ������λ, ��������) Loop
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
        For r_Prepay In (Select NO, ʵ��Ʊ��, ����id, ��ҳid, ����id, Max(���㷽ʽ) As ���㷽ʽ, �������, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��,
                                -1 * Sum(��Ԥ��) As ��Ԥ��, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������
                         From ����Ԥ����¼ A
                         Where ��¼���� In (1, 11) And a.����id In (Select Column_Value From Table(f_Str2list(v_����ids))) And
                               Nvl(��Ԥ��, 0) <> 0
                         Group By n_Tempid, NO, ʵ��Ʊ��, ����id, ��ҳid, ����id, �������, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, �����id, ���㿨���, ����,
                                  ������ˮ��, ����˵��, ������λ, ��������) Loop
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

--91427:����,2015-12-22,����ҩƷ����ϵͳ
Create Or Replace Procedure Zl_ҩƷ���ռ�¼_Insert
(
  Id_In         In ҩƷ���ռ�¼.Id%Type,
  No_In         In ҩƷ���ռ�¼.No%Type,
  �ⷿid_In     In ҩƷ���ռ�¼.�ⷿid%Type,
  ��ҩ��λid_In In ҩƷ���ռ�¼.��ҩ��λid%Type,
  ������_In     In ҩƷ���ռ�¼.������%Type,
  ��������_In   In ҩƷ���ռ�¼.��������%Type,
  �Ƿ�ϸ�_In   In ҩƷ���ռ�¼.�Ƿ�ϸ�%Type :=0,
  ��ע_in     in ҩƷ���ռ�¼.��ע%type :=null
) Is
Begin
  Insert Into ҩƷ���ռ�¼
    (ID, NO, �ⷿid, ��ҩ��λid, ������, ��������,  �Ƿ�ϸ�,��ע)
  Values
    (Id_In, No_In, �ⷿid_In, ��ҩ��λid_In, ������_In, ��������_In, �Ƿ�ϸ�_In,��ע_in);

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End;
/

--91427:����,2015-12-22,����ҩƷ����ϵͳ
Create Or Replace Procedure Zl_ҩƷ������ϸ_Insert
(
  ����id_In   In ҩƷ������ϸ.����id%Type,
  ҩƷid_In   In ҩƷ������ϸ.ҩƷid%Type,
  �ɱ���_In   In ҩƷ������ϸ.�ɱ���%Type :=null,
  ���ۼ�_In   In ҩƷ������ϸ.���ۼ�%Type :=null,
  ��ҩ����_In In ҩƷ������ϸ.��ҩ����%Type:=null,
  ����_In     In ҩƷ������ϸ.����%Type:=null,
  ��������_In In ҩƷ������ϸ.��������%Type:=null,
  Ч��_In     In ҩƷ������ϸ.Ч��%Type:=null,
  ����_In     In ҩƷ������ϸ.����%Type:=null,
  ��׼�ĺ�_In In ҩƷ������ϸ.��׼�ĺ�%Type:=null,
  ��ҩ����_In In ҩƷ������ϸ.��ҩ����%Type:=null,
  �Ƿ�ϸ�_In In ҩƷ������ϸ.�Ƿ�ϸ�%Type:=0
) Is
Begin
  Insert Into ҩƷ������ϸ
    (����id, ҩƷid, �ɱ���, ���ۼ�, ��ҩ����, ����, ��������, Ч��, ����, ��׼�ĺ�, ��ҩ����, �Ƿ�ϸ�)
  Values
    (����id_In, ҩƷid_In, �ɱ���_In, ���ۼ�_In, ��ҩ����_In, ����_In, ��������_In, Ч��_In, ����_In, ��׼�ĺ�_In, ��ҩ����_In, �Ƿ�ϸ�_In);

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End;
/

--91427:����,2015-12-22,����ҩƷ����ϵͳ
Create Or Replace Procedure Zl_ҩƷ���ռ�¼_Delete(����id_In In ҩƷ���ռ�¼.Id%Type) Is
  Err_Isverified Exception;
Begin
  Delete From ҩƷ���ռ�¼ Where ID = ����id_In And ������ Is Null;

  If Sql%RowCount = 0 Then
    Raise Err_Isverified;
  End If;
Exception
  When Err_Isverified Then
    Raise_Application_Error(-20101, '[ZLSOFT]�õ����Ѿ�������ɾ�����ѱ�����ˣ�[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End;
/

--91427:����,2015-12-22,����ҩƷ����ϵͳ
Create Or Replace Procedure Zl_ҩƷ���ռ�¼_Verify
(
  ����id_In   In ҩƷ���ռ�¼.Id%Type,
  ������_In   In ҩƷ���ռ�¼.������%Type,
  ��������_In In ҩƷ���ռ�¼.��������%Type
) Is
  Err_Isverified Exception;
Begin
  Update ҩƷ���ռ�¼ Set ������ = ������_In, �������� = ��������_In Where ID = ����id_In And ������ Is Null;

  If Sql%RowCount = 0 Then
    Raise Err_Isverified;
  End If;
Exception
  When Err_Isverified Then
    Raise_Application_Error(-20101, '[ZLSOFT]�õ����ѱ�����ˣ�[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End;
/

--92705:������,2016-01-15,��Ⱦ������վ�������ϱ��ı��棬Ӧ�û��˵����ϱ�״̬
--91225:������,2016-01-11,���պ����һ����к󣬻���ʱ���˽��պ����
CREATE OR REPLACE Procedure Zl_�����걨��¼_Untread
(
  �ļ�id_In   In Varchar2,
  IsStation_In   in Number:=NULL      --�Ƿ��Ǵ�Ⱦ������վ���ã�0�����ǣ�1�Ǵ�Ⱦ������վ����
) Is
  n_����״̬ �����걨��¼.����״̬%Type;
  n_�ļ�id   Number;
  n_Count    Number;
Begin
  If Length(�ļ�id_In) <> 32 Then
    n_�ļ�id := To_Number(�ļ�id_In); --�²���ID��32λGUID
  End If;

  Select count(1)
  Into n_Count
  From �����걨��¼
  Where ������ Is Not Null And ����ʱ�� Is Not Null And
        Decode(Length(�ļ�id_In), 32, �ĵ�id, �ļ�id) = Decode(Length(�ļ�id_In), 32, �ļ�id_In, n_�ļ�id);

  If n_Count > 0 Then      --ȡ��ɾ��
    Update �����걨��¼
    Set ������ = Null, ����ʱ�� = Null
    Where ������ Is Not Null And ����ʱ�� Is Not Null And
          Decode(Length(�ļ�id_In), 32, �ĵ�id, �ļ�id) = Decode(Length(�ļ�id_In), 32, �ļ�id_In, n_�ļ�id);
  Else
    Select ����״̬
    Into n_����״̬
    From �����걨��¼
    Where Decode(Length(�ļ�id_In), 32, �ĵ�id, �ļ�id) = Decode(Length(�ļ�id_In), 32, �ļ�id_In, n_�ļ�id);
    If n_����״̬ = 2 Then         --ȡ���ϱ�
      --������걨�Ǽ�ʱ���еĹ鵵���鵵�˺��걨�Ǽ����Ƿ���ͬ������ȡ���鵵
      If Length(�ļ�id_In) <> 32 Then
        Update ���Ӳ�����¼
        Set �鵵�� = Null, �鵵���� = Null
        Where ID = n_�ļ�id And �鵵�� = (Select �Ǽ��� From �����걨��¼ Where �ļ�id = n_�ļ�id);
      End If;
      if IsStation_In =1 then
        Update �����걨��¼
        Set ����״̬ = 3, ������ = '', ����ʱ�� = Null, ���͵�λ = Null, ���ͱ�ע = '', �Ǽ��� = '', �Ǽ�ʱ�� = ''
        Where Decode(Length(�ļ�id_In), 32, �ĵ�id, �ļ�id) = Decode(Length(�ļ�id_In), 32, �ļ�id_In, n_�ļ�id);
      else
        Update �����걨��¼
        Set ����״̬ = 1, ������ = '', ����ʱ�� = Null, ���͵�λ = Null, ���ͱ�ע = '', �Ǽ��� = '', �Ǽ�ʱ�� = ''
        Where Decode(Length(�ļ�id_In), 32, �ĵ�id, �ļ�id) = Decode(Length(�ļ�id_In), 32, �ļ�id_In, n_�ļ�id);
      end if;
    Elsif n_����״̬ = 1 Or n_����״̬ = -1 Then   --ȡ�����պ;ܾ�
      If Length(�ļ�id_In) <> 32 Then
        Update ���Ӳ�����¼ Set ����״̬ = 0 Where ID = n_�ļ�id;
      End If;
      Delete �����걨��¼
      Where Decode(Length(�ļ�id_In), 32, �ĵ�id, �ļ�id) = Decode(Length(�ļ�id_In), 32, �ļ�id_In, n_�ļ�id);
    Elsif n_����״̬ = 3 Or n_����״̬ = 4 Then  --ȡ�����
      If Length(�ļ�id_In) <> 32 Then
        Update ���Ӳ�����¼ Set ����״̬ = 0 Where ID = n_�ļ�id;
		Delete �������淴�� Where �ļ�id = n_�ļ�id And �Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From �������淴�� Where �ļ�id = n_�ļ�id);
      End If;
      
      Update �����걨��¼ Set ����״̬ = 1
      Where Decode(Length(�ļ�id_In), 32, �ĵ�id, �ļ�id) = Decode(Length(�ļ�id_In), 32, �ļ�id_In, n_�ļ�id);
    End If;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����걨��¼_Untread;
/

--91225:������,2015-12-22,��Ⱦ������ϵͳ��������
CREATE OR REPLACE Procedure Zl_�����걨��¼_Delete(Id_In In Varchar2) Is
  v_������ �����걨��¼.������%Type;
  n_�ļ�id Number;
  e_Changed Exception;
Begin

  If Length(Id_In) <> 32 Then
    n_�ļ�id := To_Number(Id_In); --�²���ID��32λGUID 
  End If;
  Select b.���� Into v_������ From �ϻ���Ա�� A, ��Ա�� B Where a.��Աid = b.Id And a.�û��� = User And Rownum < 2;

  Update �����걨��¼
  Set ������ = v_������, ����ʱ�� = Sysdate
  Where Decode(Length(Id_In), 32, �ĵ�id, �ļ�id) = Decode(Length(Id_In), 32, Id_In, n_�ļ�id);
  If Sql%RowCount = 0 Then
    Raise e_Changed;
  End If;

Exception
  When e_Changed Then
    Raise_Application_Error(-20101, '[ZLSOFT]���������Ѿ��������û��ı䣡[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����걨��¼_Delete;
/

--91225:������,2015-12-24,��Ⱦ������ϵͳ
CREATE OR REPLACE Procedure Zl_�����걨��¼_Update
(
  �ļ�id_In     In �����걨��¼.�ļ�id%Type,
  Aduitstate_In In Number,
  �Ǽ�ʱ��_In   In �������淴��.�Ǽ�ʱ��%Type,
  �Ǽ���_In     In �������淴��.�Ǽ���%Type,
  ��������_In   In �������淴��.��������%Type,
  ������_In     In �������淴��.������%Type,
  ����ʱ��_In   In �������淴��.����ʱ��%Type,
  ��������_In   In �������淴��.�������˵��%Type
) Is
Begin
  If Aduitstate_In = 3 Then
    Update �����걨��¼ Set ����״̬ = Aduitstate_In Where �ļ�id = �ļ�id_In;
    Insert Into �������淴��
      (�ļ�id, �Ǽ�ʱ��, �Ǽ���, ��¼״̬, ��������)
    Values
      (�ļ�id_In, �Ǽ�ʱ��_In, �Ǽ���_In, 3, ��������_In);
  Elsif Aduitstate_In = 4 Then
    Update �����걨��¼ Set ����״̬ = Aduitstate_In Where �ļ�id = �ļ�id_In;
    Insert Into �������淴��
      (�ļ�id, �Ǽ�ʱ��, �Ǽ���, ��¼״̬, ��������)
    Values
      (�ļ�id_In, �Ǽ�ʱ��_In, �Ǽ���_In, 1, ��������_In);
  Elsif Aduitstate_In = 5 Then
      Update �����걨��¼ Set ����״̬ = Aduitstate_In, ��������='2 ��������' Where �ļ�id = �ļ�id_In;
    
      Update �������淴��
      Set ��¼״̬ = 2, ������ = ������_In, ����ʱ�� = ����ʱ��_In, �������˵�� = ��������_In
      Where �ļ�id = �ļ�id_In And �Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From �������淴�� Where �ļ�id = �ļ�id_In);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����걨��¼_Update;
/

--91225:������,2015-12-18,���Լ��������͸��µĹ���
CREATE OR REPLACE Procedure Zl_�������Լ���¼_Insert
(
  Id_In         In �������Լ�¼.Id%Type,
  ����id_In     In �������Լ�¼.����id%Type,
  ��ҳid_In     In �������Լ�¼.��ҳid%Type,
  �Һŵ�_In     In �������Լ�¼.�Һŵ�%Type,
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
    (ID, ����id, ��ҳid, �Һŵ�, �ͼ�ʱ��, �ͼ����id, �ͼ�ҽ��, �걾����, �������, ��Ⱦ������, ���ʱ��, �Ǽ�ʱ��, �Ǽ���, �Ǽǿ���id, ��¼״̬)
  Values
    (Id_In, ����id_In, ��ҳid_In, �Һŵ�_In, �ͼ�ʱ��_In, �ͼ����id_In, �ͼ�ҽ��_In, �걾����_In, �������_In, ��Ⱦ��_In, ���ʱ��_In, �Ǽ�ʱ��_In, �Ǽ���_In,
     �Ǽǿ���id_In, ��¼״̬_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�������Լ���¼_Insert;
/

--91225:������,2016-01-11,���Լ������������ȡ������������
CREATE OR REPLACE Procedure Zl_�������Լ���¼_Update
(
  Operate_in         In  number,
  Id_In           In �������Լ�¼.Id%Type,
  �ļ�ID_In       In �������Լ�¼.�ļ�ID%Type,
  ��¼״̬_In     In �������Լ�¼.��¼״̬%Type,
  ������_In       In �������Լ�¼.������%Type,
  ����ʱ��_In     In �������Լ�¼.����ʱ��%Type,
  �������˵��_In In �������Լ�¼.�������˵��%Type
) Is
Begin
  if Operate_in = 1 then      /*���ô���˵�� */
      Update �������Լ�¼
      Set ������ = ������_In, ����ʱ�� = ����ʱ��_In, �������˵�� = �������˵��_In,��¼״̬ = ��¼״̬_In,�ļ�ID = �ļ�ID_In
      Where ID = Id_In;
  elsif Operate_in = 2 then   /*�������浥�����Խ��������*/
    if  �ļ�ID_In is not null then
        Update �������Լ�¼ Set �ļ�ID = �ļ�ID_In Where ID = Id_In;
    end if;
  elsif Operate_in = 3 then   /*ȡ�����浥�����Խ���������Ĺ���*/
    if  �ļ�ID_In is not null then
        Update �������Լ�¼ Set �ļ�ID = NULL Where �ļ�ID = �ļ�ID_In;
    end if;
  end if;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�������Լ���¼_Update;
/

--91709:Ƚ����,2015-12-17,�쳣�շѵ�������δ��������Ԥ����¼��
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
    n_����id := ����id_In;
    If Nvl(n_�ؽ�id, 0) <> 0 Then
      n_����id := n_�ؽ�id;
    End If;
  
    Update ����Ԥ����¼
    Set ��Ԥ�� = Nvl(��Ԥ��, 0) + Nvl(�����_In, 0)
    Where ����id = n_����id And ���㷽ʽ = v_����;
    If Sql%NotFound Then
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, У�Ա�־, �����id, ���㿨���, ����,
         ������ˮ��, ����˵��, �������, ��������)
      Values
        (����Ԥ����¼_Id.Nextval, 3, Null, 2, r_Balancedata.����id, Null, Null, v_����, r_Balancedata.�տ�ʱ��, r_Balancedata.����Ա���,
         r_Balancedata.����Ա����, �����_In, n_����id, r_Balancedata.�ɿ���id, r_Balancedata.�������, 2, Null, Null, ����_In, ������ˮ��_In,
         ����˵��_In, Null, 3);
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

--91225:������,2015-12-16,��Ⱦ������ϵͳ������ �������淴�� ��
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
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                  ' Where ����id = :1 And ��ҳid = :2';
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
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                ' Where ����id = :1 And ��ҳid = :2 ';
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
        v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                    ' Where ' || v_Field || ' = :1';
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
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                ' Where ����id = :1 And ��ҳid = :2 ';
    Execute Immediate v_Sql
      Using n_Pati_Id, n_Page_Id;
  
    For P In (Select ID As �ļ�id From H���˻����ļ� Where ����id = n_Pati_Id And ��ҳid = n_Page_Id) Loop
      For R In (Select Column_Value From Table(f_Str2list('���˻�������,���˻����ӡ,���˻�����Ŀ,���˻���Ҫ������,����Ҫ������'))) Loop
        v_Table  := r.Column_Value;
        v_Fields := Getfields(v_Table);
        v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                    ' Where �ļ�id = :1';
        Execute Immediate v_Sql
          Using p.�ļ�id;
      
        If v_Table = '���˻�������' Then
          v_Fields := Getfields('���˻�����ϸ');
          v_Sql    := 'Insert Into ���˻�����ϸ(' || v_Fields || ') Select ' || v_Fields ||
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
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                ' Where ����id = :1 And ��ҳid = :2 ';
    Execute Immediate v_Sql
      Using n_Pati_Id, n_Page_Id;
                   
                   For P In (Select ID From H���˻����¼ Where ����id = n_Pati_Id And ��ҳid = n_Page_Id) Loop      
        v_Table  := '���˻�������';
        v_Fields := Getfields(v_Table);
        v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                    ' Where ��¼ID = :1';
        Execute Immediate v_Sql
          Using p.ID;
      
        v_Sql := 'Delete H' || v_Table || ' Where ��¼ID = :1';
        Execute Immediate v_Sql
          Using p.ID;     
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
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                ' Where id = :1';
    Execute Immediate v_Sql
      Using n_Rec_Id;
  
    --������ϼ�¼��Zl_Retu_Other����ת�أ��޲���ID�����
    --Ӱ�񱨸沵��,����ҽ������,������ļ�¼,�⼸�ű��������Zl_Retu_Order��ת��ҽ�����ٴ���
    For R In (Select Column_Value From Table(f_Str2list('���Ӳ�������,���Ӳ�����ʽ,���Ӳ�������,�����걨��¼,�������淴��'))) Loop
      v_Table := r.Column_Value;
      If v_Table = '���Ӳ�������' Then
        v_Field := '����id';
      Else
        v_Field := '�ļ�id';
      End If;
      v_Fields := Getfields(v_Table);
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                  ' Where ' || v_Field || ' = :1';
      Execute Immediate v_Sql
        Using n_Rec_Id;
    
      If v_Table = '���Ӳ�������' Then
        v_Fields := Getfields('���Ӳ���ͼ��');
        v_Sql    := 'Insert Into ���Ӳ���ͼ��(' || v_Fields || ') Select ' || v_Fields ||
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
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                ' Where id = :1';
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
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                  ' Where ' || v_Field || ' = :1';
      Execute Immediate v_Sql
        Using n_Rec_Id;
    
      If v_Table = '����ҽ��״̬' Then
        v_Fields := Getfields('ҽ��ǩ����¼');
        v_Sql    := 'Insert Into ҽ��ǩ����¼(' || v_Fields || ') Select ' || v_Fields ||
                    ' From Hҽ��ǩ����¼ Where ID In (Select ǩ��id From H����ҽ��״̬ Where ҽ��id = :1 And ǩ��id Is Not Null)';
        Execute Immediate v_Sql
          Using n_Rec_Id;
        Delete Hҽ��ǩ����¼
        Where ID In (Select ǩ��id From H����ҽ��״̬ Where ҽ��id = n_Rec_Id And ǩ��id Is Not Null);
      
      Elsif v_Table = '����ҽ������' Then
        v_Fields := Getfields('���Ƶ��ݴ�ӡ');
        v_Sql    := 'Insert Into ���Ƶ��ݴ�ӡ(' || v_Fields || ') Select ' || v_Fields ||
                    ' From H���Ƶ��ݴ�ӡ Where (NO, ��¼����) In (Select NO, ��¼���� From H����ҽ������ Where ҽ��id = :1)';
        Execute Immediate v_Sql
          Using n_Rec_Id;
        Delete H���Ƶ��ݴ�ӡ Where (NO, ��¼����) In (Select NO, ��¼���� From H����ҽ������ Where ҽ��id = n_Rec_Id);
      
      Elsif v_Table = 'Ӱ�����¼' Then
        v_Fields := Getfields('Ӱ��������');
        v_Sql    := 'Insert Into Ӱ��������(' || v_Fields || ') Select ' || v_Fields ||
                    ' From HӰ�������� Where ���uid In (Select ���uid From HӰ�����¼ Where ҽ��id = :1)';
        Execute Immediate v_Sql
          Using n_Rec_Id;
      
        v_Fields := Getfields('Ӱ����ͼ��');
        v_Sql    := 'Insert Into Ӱ����ͼ��(' || v_Fields || ') Select ' || v_Fields ||
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
          v_Sql    := 'Insert Into ' || v_Subtable || '(' || v_Fields || ') Select ' || v_Fields || ' From H' ||
                      v_Subtable || ' Where ' || v_Subfield || ' In (Select ID From H����걾��¼ Where ҽ��id = :1)';
          Execute Immediate v_Sql
            Using n_Rec_Id;
        
          v_Sql := 'Delete H' || v_Subtable || ' Where ' || v_Subfield ||
                   ' In (Select ID From H����걾��¼ Where ҽ��id = :1)';
          Execute Immediate v_Sql
            Using n_Rec_Id;
        End Loop;
      
        v_Fields := Getfields('������ͨ���');
        v_Sql    := 'Insert Into ������ͨ���(' || v_Fields || ') Select ' || v_Fields ||
                    ' From H������ͨ��� Where ����걾id In (Select ID From H����걾��¼ Where ҽ��id = :1)';
        Execute Immediate v_Sql
          Using n_Rec_Id;
      
        v_Fields := Getfields('����ҩ�����');
        v_Sql    := 'Insert Into ����ҩ�����(' || v_Fields || ') Select ' || v_Fields ||
                    ' From H����ҩ����� Where ϸ�����id In (Select ID From H������ͨ��� Where ����걾id In (Select ID From H����걾��¼ Where ҽ��id = :1))';
        Execute Immediate v_Sql
          Using n_Rec_Id;
      
        v_Fields := Getfields('�����ʿر���');
        v_Sql    := 'Insert Into �����ʿر���(' || v_Fields || ') Select ' || v_Fields ||
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
      Execute Immediate 'zl24_Retu_Oper(:1)'
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
    Select Nvl(ֻ��, 0) Into n_ֻ�� From Zlbakspaces Where ϵͳ = n_System And ��ǰ = 1;
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
      Select Nvl(ֻ��, 0) Into n_ֻ�� From Zlbakspaces Where ϵͳ = n_Opersystem And ��ǰ = 1;
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
    v_Table  := '���˹Һż�¼';
    v_Fields := Getfields(v_Table);
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                ' Where NO =:1 ';
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
    v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || v_Fields || ' From H' || v_Table ||
                ' Where NO =:1';
   Execute Immediate v_Sql
      Using v_Times;
  
    Delete H����ת���¼ Where NO = v_Times;
    Delete H���˹Һż�¼ Where NO = v_Times;
  
    --2.סԺ���ˣ�������ID����ҳID���
  Elsif n_Flag = 1 Then
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

  Begin
    Execute Immediate 'Update zlbakInfo  set ���ת������=sysdate where ϵͳ=' || n_System;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm || ':' || v_Sql);
End Zl_Retu_Clinic;
/

--91225:������,2015-12-16,��Ⱦ������ϵͳ������ �������淴��
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
                                         ',������ϼ�¼_PK,����ҽ��״̬_PK,ҽ��ǩ����¼_PK,����ҽ������_PK,���Ƶ��ݴ�ӡ_PK,ҽ��ִ�мƼ�_PK,ִ�д�ӡ��¼_PK' ||
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

--91225:������,2015-12-16,��Ⱦ������ϵͳ���� �������淴�� ��
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

--92335:���ϴ�,2016-01-18,����֧����ģʽ�����̲��
--91561:������,2015-12-14,ԤԼ������ѺŲ���Ԥ����¼
Create Or Replace Procedure Zl_ԤԼ�ҺŽ���_Insert
(
  No_In            ������ü�¼.No%Type,
  Ʊ�ݺ�_In        ������ü�¼.ʵ��Ʊ��%Type,
  ����id_In        Ʊ��ʹ����ϸ.����id%Type,
  ����id_In        ������ü�¼.����id%Type,
  ����_In          ������ü�¼.��ҩ����%Type,
  ����id_In        ������ü�¼.����id%Type,
  �����_In        ������ü�¼.��ʶ��%Type,
  ����_In          ������ü�¼.����%Type,
  �Ա�_In          ������ü�¼.�Ա�%Type,
  ����_In          ������ü�¼.����%Type,
  ���ʽ_In      ������ü�¼.���ʽ%Type, --���ڴ�Ų��˵�ҽ�Ƹ��ʽ���
  �ѱ�_In          ������ü�¼.�ѱ�%Type,
  ���㷽ʽ_In      ����Ԥ����¼.���㷽ʽ%Type, --�ֽ�Ľ�������
  �ֽ�֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�ֽ�֧�����ݽ��
  Ԥ��֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱʹ�õ�Ԥ�����
  ����֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�����ʻ�֧�����
  ����ʱ��_In      ������ü�¼.����ʱ��%Type,
  ����_In          �Һ����״̬.���%Type,
  ����Ա���_In    ������ü�¼.����Ա���%Type,
  ����Ա����_In    ������ü�¼.����Ա����%Type,
  ���ɶ���_In      Number := 0,
  �Ǽ�ʱ��_In      ������ü�¼.�Ǽ�ʱ��%Type := Null,
  �����id_In      ����Ԥ����¼.�����id%Type := Null,
  ���㿨���_In    ����Ԥ����¼.���㿨���%Type := Null,
  ����_In          ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In    ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In      ����Ԥ����¼.����˵��%Type := Null,
  ����_In          ���˹Һż�¼.����%Type := Null,
  ����ģʽ_In      Number := 0,
  ���ʷ���_In      Number := 0,
  ��Ԥ������ids_In Varchar2 := Null,
  ��������_In      Number := 0,
  ���½������_In  Number := 0--�Ƿ������Ա��������Ҫ�Ǵ���ͳһ����Ա��¼��̨�����������
) As
  --���α������շѳ�Ԥ���Ŀ���Ԥ���б�
  --��ID�������ȳ��ϴ�δ����ġ�
  Cursor c_Deposit
  (
    v_����id        ������Ϣ.����id%Type,
    v_��Ԥ������ids Varchar2
  ) Is
    Select *
    From (Select a.Id, a.����id, a.��¼״̬, a.Ԥ�����, a.No, Nvl(a.���, 0) As ���
           From ����Ԥ����¼ A,
                (Select NO, Sum(Nvl(a.���, 0)) As ���
                  From ����Ԥ����¼ A
                  Where a.����id Is Null And Nvl(a.���, 0) <> 0 And
                        a.����id In (Select Column_Value From Table(f_Num2list(v_��Ԥ������ids))) And Nvl(a.Ԥ�����, 2) = 1
                  Group By NO
                  Having Sum(Nvl(a.���, 0)) <> 0) B
           Where a.����id Is Null And Nvl(a.���, 0) <> 0 And a.���㷽ʽ Not In (Select ���� From ���㷽ʽ Where ���� = 5) And
                 a.No = b.No And a.����id In (Select Column_Value From Table(f_Num2list(v_��Ԥ������ids))) And
                 Nvl(a.Ԥ�����, 2) = 1
           Union All
           Select 0 As ID, Max(����id) As ����id, ��¼״̬, Ԥ�����, NO, Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) As ���
           From ����Ԥ����¼
           Where ��¼���� In (1, 11) And ����id Is Not Null And Nvl(���, 0) <> Nvl(��Ԥ��, 0) And
                 ����id In (Select Column_Value From Table(f_Num2list(v_��Ԥ������ids))) And Nvl(Ԥ�����, 2) = 1 Having
            Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0
           Group By ��¼״̬, Ԥ�����, NO)
    Order By Decode(����id, Nvl(v_����id, 0), 0, 1), ID, NO, Ԥ�����;

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_�ֽ�     ���㷽ʽ.����%Type;
  v_�����ʻ� ���㷽ʽ.����%Type;
  v_�������� �ŶӽкŶ���.��������%Type;
  v_�ű�     ������ü�¼.���㵥λ%Type;
  v_����     ������ü�¼.��ҩ����%Type;
  v_�ŶӺ��� �ŶӽкŶ���.�ŶӺ��� %Type;
  v_ԤԼ��ʽ ���˹Һż�¼.ԤԼ��ʽ %Type;

  n_��ӡid        Ʊ�ݴ�ӡ����.Id%Type;
  n_Ԥ�����      ����Ԥ����¼.���%Type;
  n_����ֵ        ����Ԥ����¼.���%Type;
  v_��Ԥ������ids Varchar2(4000);

  n_�Һ�id         ���˹Һż�¼.Id%Type;
  n_����̨ǩ���Ŷ� Number;
  n_��id           ����ɿ����.Id%Type;
  n_Count          Number(18);
  n_�Ŷ�           Number;
  n_�����Ŷ�       Number;
  n_��ǰ���       ����Ԥ����¼.���%Type;
  n_Ԥ��id         ����Ԥ����¼.Id%Type;
  n_���ѿ�id       ���ѿ�Ŀ¼.Id%Type;
  n_���ƿ�         Number;

  d_Date     Date;
  d_ԤԼʱ�� ������ü�¼.����ʱ��%Type;
  d_����ʱ�� Date;
  d_�Ŷ�ʱ�� Date;
  n_ʱ��     Number := 0;
  n_����     Number := 0;
  v_�Ŷ���� �ŶӽкŶ���.�Ŷ����%Type;
  n_����ģʽ ������Ϣ.����ģʽ%Type;
  n_Ʊ��     Ʊ��ʹ����ϸ.Ʊ��%Type;
  v_���ʽ ���˹Һż�¼.ҽ�Ƹ��ʽ%Type;
  n_����ģʽ Number := 0;
Begin
  n_��id          := Zl_Get��id(����Ա����_In);
  v_��Ԥ������ids := Nvl(��Ԥ������ids_In, ����id_In);
  n_����ģʽ      := Nvl(zl_GetSysParameter(64, 1111), 0);

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
  If �Ǽ�ʱ��_In Is Null Then
    Select Sysdate Into d_Date From Dual;
  Else
    d_Date := �Ǽ�ʱ��_In;
  End If;

  --���¹Һ����״̬
  Begin
    Select �ű�, ����, Trunc(����ʱ��), ����ʱ��, ԤԼ��ʽ
    Into v_�ű�, v_����, d_ԤԼʱ��, d_����ʱ��, v_ԤԼ��ʽ
    From ���˹Һż�¼
    Where ��¼���� = 2 And ��¼״̬ = 1 And Rownum = 1 And NO = No_In;
  Exception
    When Others Then
      v_Err_Msg := '��ǰԤԼ�Һŵ��ѱ������˽���';
      Raise Err_Item;
  End;

  --�ж��Ƿ��ʱ��
  Begin
    Select 1
    Into n_ʱ��
    From Dual
    Where Exists (Select 1
           From �ҺŰ���ʱ�� A, �ҺŰ��� B
           Where a.����id = b.Id And b.���� = v_�ű� And Rownum < 2
           Union All
           Select 1
           From �Һżƻ�ʱ�� C, �ҺŰ��żƻ� D ��
           Where c.�ƻ�id = d.Id And d.���� = v_�ű� And d.��Чʱ�� > Sysdate And Rownum < 2);
  Exception
    When Others Then
      n_ʱ�� := 0;
  End;

  If n_ʱ�� = 0 And ��������_In = 0 Then
    If n_����ģʽ = 0 Then
      If Trunc(����ʱ��_In) = Trunc(Sysdate) Then
        d_����ʱ�� := ����ʱ��_In;
      Else
        d_����ʱ�� := Sysdate;
      End If;
    Else
      d_����ʱ�� := ����ʱ��_In;
    End If;
  Else
    If Not ����ʱ��_In Is Null Then
      d_����ʱ�� := ����ʱ��_In;
    End If;
  End If;
  If Not v_���� Is Null Then
    If ����_In Is Null Then
      Delete �Һ����״̬ Where ���� = v_�ű� And Trunc(����) = Trunc(d_ԤԼʱ��) And ��� = v_����;
    Else
      If Trunc(d_ԤԼʱ��) <> Trunc(Sysdate) And n_����ģʽ = 0 Then
      
        If n_ʱ�� = 0 And ��������_In = 0 Then
          --��ǰ���ջ��ӳٽ���
          Delete �Һ����״̬ Where ���� = v_�ű� And Trunc(����) = Trunc(d_ԤԼʱ��) And ��� = v_����;
          Begin
            Select 1 Into n_���� From �Һ����״̬ Where ���� = v_�ű� And ���� = Trunc(Sysdate) And ��� = v_����;
          Exception
            When Others Then
              n_���� := 0;
          End;
          If n_���� = 0 Then
            Insert Into �Һ����״̬
              (����, ����, ���, ״̬, ����Ա����, �Ǽ�ʱ��)
            Values
              (v_�ű�, Trunc(Sysdate), v_����, 1, ����Ա����_In, Sysdate);
          Else
            --�����ѱ�ʹ�õ����
            Begin
              v_���� := 1;
              Insert Into �Һ����״̬
                (����, ����, ���, ״̬, ����Ա����, �Ǽ�ʱ��)
              Values
                (v_�ű�, Trunc(Sysdate), v_����, 1, ����Ա����_In, Sysdate);
            Exception
              When Others Then
                Select Min(��� + 1)
                Into v_����
                From �Һ����״̬ A
                Where ���� = v_�ű� And ���� = Trunc(Sysdate) And Not Exists
                 (Select 1 From �Һ����״̬ Where ���� = a.���� And ���� = a.���� And ��� = a.��� + 1);
                Insert Into �Һ����״̬
                  (����, ����, ���, ״̬, ����Ա����, �Ǽ�ʱ��)
                Values
                  (v_�ű�, Trunc(Sysdate), v_����, 1, ����Ա����_In, Sysdate);
            End;
          End If;
        Else
          Update �Һ����״̬
          Set ״̬ = 1, �Ǽ�ʱ�� = Sysdate
          Where Trunc(����) = Trunc(d_ԤԼʱ��) And ��� = v_���� And ���� = v_�ű� And ״̬ = 2;
          If Sql% NotFound Then
            Begin
              Insert Into �Һ����״̬
                (����, ����, ���, ״̬, ����Ա����, �Ǽ�ʱ��)
              Values
                (v_�ű�, Trunc(Sysdate), v_����, 1, ����Ա����_In, Sysdate);
            Exception
              When Others Then
                v_Err_Msg := '���' || v_���� || '�ѱ�������ʹ��,������ѡ��һ�����.';
                Raise Err_Item;
            End;
          End If;
        
        End If;
      
      Else
        Update �Һ����״̬
        Set ��� = ����_In, ״̬ = 1, �Ǽ�ʱ�� = Sysdate
        Where ���� = v_�ű� And Trunc(����) = Trunc(d_ԤԼʱ��) And ��� = v_����;
        If Sql%RowCount = 0 Then
          Begin
            Insert Into �Һ����״̬
              (����, ����, ���, ״̬, ����Ա����, �Ǽ�ʱ��)
            Values
              (v_�ű�, Trunc(d_����ʱ��), v_����, 1, ����Ա����_In, Sysdate);
          Exception
            When Others Then
              v_Err_Msg := '���' || v_���� || '�ѱ�������ʹ��,������ѡ��һ�����.';
              Raise Err_Item;
          End;
        End If;
      End If;
    End If;
  Else
    If Not ����_In Is Null Then
      Begin
        Insert Into �Һ����״̬
          (����, ����, ���, ״̬, ����Ա����, �Ǽ�ʱ��)
        Values
          (v_�ű�, Trunc(Sysdate), ����_In, 1, ����Ա����_In, Sysdate);
      Exception
        When Others Then
          v_Err_Msg := '���' || ����_In || '�ѱ�������ʹ��,������ѡ��һ�����.';
          Raise Err_Item;
      End;
      v_���� := ����_In;
    Else
      v_���� := Null;
    End If;
  End If;

  --����������ü�¼
  Update ������ü�¼
  Set ��¼״̬ = 1, ʵ��Ʊ�� = Decode(Nvl(���ʷ���_In, 0), 1, Null, Ʊ�ݺ�_In), ����id = Decode(Nvl(���ʷ���_In, 0), 1, Null, ����id_In),
      ���ʽ�� = Decode(Nvl(���ʷ���_In, 0), 1, Null, ʵ�ս��), ��ҩ���� = ����_In, ����id = ����id_In, ��ʶ�� = �����_In, ���� = ����_In, ���� = ����_In,
      �Ա� = �Ա�_In, ���ʽ = ���ʽ_In, �ѱ� = �ѱ�_In, ����ʱ�� = d_����ʱ��, �Ǽ�ʱ�� = d_Date, ����Ա��� = ����Ա���_In, ����Ա���� = ����Ա����_In,
      �ɿ���id = n_��id, ���ʷ��� = Decode(Nvl(���ʷ���_In, 0), 1, 1, 0)
  Where ��¼���� = 4 And ��¼״̬ = 0 And NO = No_In;

  --���˹Һż�¼
  Update ���˹Һż�¼
  Set ������ = ����Ա����_In, ����ʱ�� = d_Date, ��¼���� = 1, ����id = ����id_In, ����� = �����_In, ����ʱ�� = d_����ʱ��, ���� = ����_In, �Ա� = �Ա�_In,
      ���� = ����_In, ����Ա��� = ����Ա���_In, ����Ա���� = ����Ա����_In, ���� = Decode(Nvl(����_In, 0), 0, Null, ����_In), ���� = v_����, ���� = ����_In
  Where ��¼״̬ = 1 And NO = No_In And ��¼���� = 2
  Returning ID Into n_�Һ�id;
  If Sql%NotFound Then
    Begin
      Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
      Begin
        Select ���� Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = ���ʽ_In And Rownum < 2;
      Exception
        When Others Then
          v_���ʽ := Null;
      End;
      Insert Into ���˹Һż�¼
        (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����,
         ժҪ, ����, ԤԼ, ԤԼ��ʽ, ������, ����ʱ��, ԤԼʱ��, ����, ҽ�Ƹ��ʽ)
        Select n_�Һ�id, No_In, 1, 1, ����id_In, �����_In, ����_In, �Ա�_In, ����_In, ���㵥λ, �Ӱ��־, ����_In, Null, ִ�в���id, ִ����, 0, Null,
               �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����, ժҪ, v_����, 1, Substr(����, 1, 10) As ԤԼ��ʽ, ����Ա����_In, Nvl(�Ǽ�ʱ��_In, Sysdate), ����ʱ��,
               Decode(Nvl(����_In, 0), 0, Null, ����_In), v_���ʽ
        From ������ü�¼
        Where ��¼���� = 4 And ��¼״̬ = 1 And Rownum = 1 And NO = No_In;
    Exception
      When Others Then
        v_Err_Msg := '���ڲ���ԭ��,���ݺ�Ϊ��' || No_In || '���Ĳ���' || ����_In || '�Ѿ�������';
        Raise Err_Item;
    End;
  End If;

  --0-����������;1-��ҽ�������̨�Ŷ�;2-�ȷ���,��ҽ��վ
  If Nvl(���ɶ���_In, 0) <> 0 Then
    n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113));
    If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Then
      For v_�Һ� In (Select ID, ����, ����, ִ����, ִ�в���id, ����ʱ��, �ű�, ���� From ���˹Һż�¼ Where NO = No_In) Loop
      
        Begin
          Select 1,
                 Case
                   When �Ŷ�ʱ�� < Trunc(Sysdate) Then
                    1
                   Else
                    0
                 End
          Into n_�Ŷ�, n_�����Ŷ�
          From �ŶӽкŶ���
          Where ҵ������ = 0 And ҵ��id = v_�Һ�.Id And Rownum <= 1;
        Exception
          When Others Then
            n_�Ŷ� := 0;
        End;
        If n_�Ŷ� = 0 Then
          --��������
          --����ִ�в��š���������
          n_�Һ�id   := v_�Һ�.Id;
          v_�������� := v_�Һ�.ִ�в���id;
          v_�ŶӺ��� := Zlgetnextqueue(v_�Һ�.ִ�в���id, n_�Һ�id, v_�Һ�.�ű� || '|' || v_�Һ�.����);
          v_�Ŷ���� := Zlgetsequencenum(0, n_�Һ�id, 0);
        
          --�Һ�id_In,����_In,����_In,ȱʡ����_In,��չ_In(������)
          d_�Ŷ�ʱ�� := Zl_Get_Queuedate(n_�Һ�id, v_�Һ�.�ű�, v_�Һ�.����, v_�Һ�.����ʱ��);
          --   ��������_In , ҵ������_In, ҵ��id_In,����id_In,�ŶӺ���_In,�Ŷӱ��_In,��������_In,����ID_IN, ����_In, ҽ������_In,
          Zl_�ŶӽкŶ���_Insert(v_��������, 0, n_�Һ�id, v_�Һ�.ִ�в���id, v_�ŶӺ���, Null, ����_In, ����id_In, v_�Һ�.����, v_�Һ�.ִ����, d_�Ŷ�ʱ��,
                           v_ԤԼ��ʽ, Null, v_�Ŷ����);
        Elsif Nvl(n_�����Ŷ�, 0) = 1 Then
          --���¶��к�
          v_�ŶӺ��� := Zlgetnextqueue(v_�Һ�.ִ�в���id, v_�Һ�.Id, v_�Һ�.�ű� || '|' || Nvl(v_�Һ�.����, 0));
          v_�Ŷ���� := Zlgetsequencenum(0, v_�Һ�.Id, 1);
          --�¶�������_IN, ҵ������_In, ҵ��id_In , ����id_In , ��������_In , ����_In, ҽ������_In ,�ŶӺ���_In
          Zl_�ŶӽкŶ���_Update(v_�Һ�.ִ�в���id, 0, v_�Һ�.Id, v_�Һ�.ִ�в���id, v_�Һ�.����, v_�Һ�.����, v_�Һ�.ִ����, v_�ŶӺ���, v_�Ŷ����);
        
        Else
          --�¶�������_IN, ҵ������_In, ҵ��id_In , ����id_In , ��������_In , ����_In, ҽ������_In ,�ŶӺ���_In
          Zl_�ŶӽкŶ���_Update(v_�Һ�.ִ�в���id, 0, v_�Һ�.Id, v_�Һ�.ִ�в���id, v_�Һ�.����, v_�Һ�.����, v_�Һ�.ִ����);
        End If;
      End Loop;
    End If;
  End If;

  --���ܽ��㵽����Ԥ����¼
  If (Nvl(�ֽ�֧��_In, 0) <> 0 Or (Nvl(�ֽ�֧��_In, 0) = 0 And Nvl(����֧��_In, 0) = 0 And Nvl(Ԥ��֧��_In, 0) = 0)) And
     Nvl(���ʷ���_In, 0) = 0 Then
    Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
    Insert Into ����Ԥ����¼
      (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, �������,
       ��������)
    Values
      (n_Ԥ��id, 4, 1, No_In, ����id_In, Nvl(���㷽ʽ_In, v_�ֽ�), �ֽ�֧��_In, d_Date, ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�', n_��id,
       �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, ����id_In, 4);
  
    If Nvl(���㿨���_In, 0) <> 0 And Nvl(�ֽ�֧��_In, 0) <> 0 Then
    
      n_���ѿ�id := Null;
      Begin
        Select Nvl(���ƿ�, 0), 1 Into n_���ƿ�, n_Count From �����ѽӿ�Ŀ¼ Where ��� = ���㿨���_In;
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
        Where �ӿڱ�� = ���㿨���_In And ���� = ����_In And
              ��� = (Select Max(���) From ���ѿ�Ŀ¼ Where �ӿڱ�� = ���㿨���_In And ���� = ����_In);
      End If;
      Zl_���˿������¼_Insert(���㿨���_In, n_���ѿ�id, ���㷽ʽ_In, �ֽ�֧��_In, ����_In, Null, �Ǽ�ʱ��_In, Null, ����id_In, n_Ԥ��id);
    End If;
  
  End If;

  --���ھ��￨ͨ��Ԥ����Һ�
  If Nvl(Ԥ��֧��_In, 0) <> 0 And Nvl(���ʷ���_In, 0) = 0 Then
    n_Ԥ����� := Ԥ��֧��_In;
    For r_Deposit In c_Deposit(����id_In, v_��Ԥ������ids) Loop
      n_��ǰ��� := Case
                  When r_Deposit.��� - n_Ԥ����� < 0 Then
                   r_Deposit.���
                  Else
                   n_Ԥ�����
                End;
      If r_Deposit.Id <> 0 Then
        --��һ�γ�Ԥ��(���Ͻ���ID,���Ϊ0)
        Update ����Ԥ����¼ Set ��Ԥ�� = 0, ����id = ����id_In, �������� = 4 Where ID = r_Deposit.Id;
      End If;
      --���ϴ�ʣ���
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
         ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, Ԥ�����, �������, ��������)
        Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �Ǽ�ʱ��_In,
               ����Ա����_In, ����Ա���_In, n_��ǰ���, ����id_In, n_��id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, Ԥ�����, ����id_In, 4
        From ����Ԥ����¼
        Where NO = r_Deposit.No And ��¼״̬ = r_Deposit.��¼״̬ And ��¼���� In (1, 11) And Rownum = 1;
    
      --���²���Ԥ�����
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) - n_��ǰ���
      Where ����id = r_Deposit.����id And ���� = 1 And ���� = Nvl(r_Deposit.Ԥ�����, 2)
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, ����, Ԥ�����, ����)
        Values
          (r_Deposit.����id, Nvl(r_Deposit.Ԥ�����, 2), -1 * n_��ǰ���, 1);
        n_����ֵ := -1 * n_��ǰ���;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From �������
        Where ����id = r_Deposit.����id And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
      End If;
    
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
  End If;

  --����ҽ���Һ�
  If Nvl(����֧��_In, 0) <> 0 And Nvl(���ʷ���_In, 0) = 0 Then
    Insert Into ����Ԥ����¼
      (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ,
       Ԥ�����, �������, ��������)
    Values
      (����Ԥ����¼_Id.Nextval, 4, 1, No_In, ����id_In, v_�����ʻ�, ����֧��_In, d_Date, ����Ա���_In, ����Ա����_In, ����id_In, 'ҽ���Һ�', n_��id,
       �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, Null, Null, ����id_In, 4);
  End If;

  --��ػ��ܱ�Ĵ���
  --��Ա�ɿ����
  If Nvl(�ֽ�֧��_In, 0) <> 0 And Nvl(���ʷ���_In, 0) = 0 And Nvl(���½������_In,0)=0 Then
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
      Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = Nvl(���㷽ʽ_In, v_�ֽ�) And Nvl(���, 0) = 0;
    End If;
  End If;

  If Nvl(����֧��_In, 0) <> 0 And Nvl(���ʷ���_In, 0) = 0 And Nvl(���½������_In,0)=0 Then
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) + ����֧��_In
    Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = v_�����ʻ�
    Returning ��� Into n_����ֵ;
  
    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ���� (�տ�Ա, ���㷽ʽ, ����, ���) Values (����Ա����_In, v_�����ʻ�, 1, ����֧��_In);
      n_����ֵ := ����֧��_In;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ��Ա�ɿ���� Where �տ�Ա = ����Ա����_In And ���� = 1 And Nvl(���, 0) = 0;
    End If;
  End If;

  --����Ʊ��ʹ�����
  If Ʊ�ݺ�_In Is Not Null And Nvl(���ʷ���_In, 0) = 0 Then
    Select Ʊ�ݴ�ӡ����_Id.Nextval Into n_��ӡid From Dual;
  
    --��ǰƱ�ݵ�Ʊ��
    Select Ʊ�� Into n_Ʊ�� From Ʊ�����ü�¼ Where ID = Nvl(����id_In, 0);
    --����Ʊ��
    Insert Into Ʊ�ݴ�ӡ���� (ID, ��������, NO) Values (n_��ӡid, 4, No_In);
  
    Insert Into Ʊ��ʹ����ϸ
      (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����)
    Values
      (Ʊ��ʹ����ϸ_Id.Nextval, n_Ʊ��, Ʊ�ݺ�_In, 1, 1, ����id_In, n_��ӡid, d_Date, ����Ա����_In);
  
    --״̬�Ķ�
    Update Ʊ�����ü�¼
    Set ��ǰ���� = Ʊ�ݺ�_In, ʣ������ = Decode(Sign(ʣ������ - 1), -1, 0, ʣ������ - 1), ʹ��ʱ�� = d_Date
    Where ID = Nvl(����id_In, 0);
  End If;

  If Nvl(���ʷ���_In, 0) = 1 Then
    --����
    If Nvl(����id_In, 0) = 0 Then
      v_Err_Msg := 'Ҫ��Բ��˵ĹҺŷѽ��м��ʣ������ǽ������˲��ܼ��ʹҺš�';
      Raise Err_Item;
    End If;
    For c_���� In (Select ʵ�ս��, ���˿���id, ��������id, ִ�в���id, ������Ŀid
                 From ������ü�¼
                 Where ��¼���� = 4 And ��¼״̬ = 1 And NO = No_In And Nvl(���ʷ���, 0) = 1) Loop
      --�������
      Update �������
      Set ������� = Nvl(�������, 0) + Nvl(c_����.ʵ�ս��, 0)
      Where ����id = Nvl(����id_In, 0) And ���� = 1 And ���� = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, ����, ����, �������, Ԥ�����)
        Values
          (����id_In, 1, 1, Nvl(c_����.ʵ�ս��, 0), 0);
      End If;
    
      --����δ�����
      Update ����δ�����
      Set ��� = Nvl(���, 0) + Nvl(c_����.ʵ�ս��, 0)
      Where ����id = ����id_In And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���˿���id, 0) = Nvl(c_����.���˿���id, 0) And
            Nvl(��������id, 0) = Nvl(c_����.��������id, 0) And Nvl(ִ�в���id, 0) = Nvl(c_����.ִ�в���id, 0) And ������Ŀid + 0 = c_����.������Ŀid And
            ��Դ;�� + 0 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into ����δ�����
          (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
        Values
          (����id_In, Null, Null, c_����.���˿���id, c_����.��������id, c_����.ִ�в���id, c_����.������Ŀid, 1, Nvl(c_����.ʵ�ս��, 0));
      End If;
    End Loop;
  End If;
  If Nvl(����id_In, 0) <> 0 Then
    n_����ģʽ := 0;
    Update ������Ϣ
    Set ����ʱ�� = d_����ʱ��, ����״̬ = 1, �������� = ����_In
    Where ����id = ����id_In
    Returning Nvl(����ģʽ, 0) Into n_����ģʽ;
    --ȡ����:
    If Nvl(n_����ģʽ, 0) <> Nvl(����ģʽ_In, 0) Then
      --����ģʽ��ȷ��
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
  End If;

  --���˵�����Ϣ
  If ����id_In Is Not Null Then
    Update ������Ϣ
    Set ������ = Null, ������ = Null, �������� = Null
    Where ����id = ����id_In And Exists (Select 1
           From ���˵�����¼
           Where ����id = ����id_In And ��ҳid Is Not Null And
                 �Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From ���˵�����¼ Where ����id = ����id_In));
    If Sql%RowCount > 0 Then
      Update ���˵�����¼
      Set ����ʱ�� = d_Date
      Where ����id = ����id_In And ��ҳid Is Not Null And Nvl(����ʱ��, d_Date) > d_Date;
    End If;
  End If;
  --��Ϣ����
  Begin
    Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
      Using 1, n_�Һ�id;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ԤԼ�ҺŽ���_Insert;
/

--91085:������,2015-12-08,ȡ����ˣ�ͬ��ȡ����ʾ
Create Or Replace Procedure Zl_����걾��¼_���ȡ��(Id_In ����걾��¼.Id%Type) Is
  --���ҵ�ǰ�걾���������
  Cursor c_Samplequest Is
    Select Distinct ҽ��id
    From (Select ҽ��id
           From ����걾��¼
           Where ID = Id_In
           Union
           Select ҽ��id From ������Ŀ�ֲ� Where �걾id = Id_In);

  v_��ҳid Number(18);
  v_No     Varchar2(20);
  v_Temp   Varchar2(255);
  v_Fileid Number(18);

  v_��Ա��� ��Ա��.���%Type;
  v_��Ա���� ��Ա��.����%Type;
  v_��ǰʱ�� Date;
Begin
  --����Ա��Ϣ:����ID,��������;��ԱID,��Ա���,��Ա����
  v_Temp     := Zl_Identity;
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_��Ա��� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_��ǰʱ�� := Sysdate;


  --1.ȡ���걾���
  Update ����걾��¼
  Set ����� = Null, ���ʱ�� = Null, ��ӡ���� = Null, ���δͨ�� = Null, ����״̬ = 1
  Where ID = Id_In;
  --Delete ����ǩ����¼ Where ����걾id = Id_In;
  --��¼��˹���
  Insert Into ���������¼
    (ID, �걾id, ��������, ����Ա, ����ʱ��)
  Values
    (���������¼_Id.Nextval, Id_In, 1, v_��Ա����, Sysdate);

  --2.��鵱ǰ�걾��ص��������ر걾
  For r_Samplequest In c_Samplequest Loop
  
    --1.�����뵥������ִ��״̬
    Update ����ҽ������
    Set ִ��״̬ = 3, ����� = Null, ���ʱ�� = Null
    Where ҽ��id In (Select ID From ����ҽ����¼ Where r_Samplequest.ҽ��id = ���id);
  
    Update ����ҽ������
    Set ִ��״̬ = 3, ����� = Null, ���ʱ�� = Null
    Where ҽ��id In (Select ID From ����ҽ����¼ Where r_Samplequest.ҽ��id = ID) And Nvl(������, '�տ�') = '�տ�';
  
    Begin
      Select ����id Into v_Fileid From ����ҽ������ Where ҽ��id = r_Samplequest.ҽ��id;
      Zl_������ļ�¼_Cancel(r_Samplequest.ҽ��id, v_Fileid, Null);
      Delete ����ҽ������ Where ҽ��id = r_Samplequest.ҽ��id;
    Exception
      When Others Then
        v_Fileid := 0;
    End;
    If v_Fileid <> 0 Then
      Delete ���Ӳ�����¼ Where ID = v_Fileid;
      Delete ���Ӳ������� Where �ļ�id = v_Fileid;
    End If;
  
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����걾��¼_���ȡ��;
/

--91259:��˶,2015-12-08,��ַ���ƣ���������С���⴦��
Create Or Replace Function Zl_Adderss_Structure(v_Addressinfo Varchar2) Return Varchar2 Is
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
        If v_���� Is Null And Not v_�ֵ� Is Null Then
          --�Ƚ�ȡ���򼶵����������ؼ��֣���ƥ��
          v_Tmp := Substr(v_�ֵ�, 1, 2);
          Select Max(����)
          Into v_����
          From ����
          Where ���� Like v_Tmp || '%' And Nvl(����, 0) = 3 And �ϼ����� = v_Code����;
          --���ڶ���ƥ�䣬��������ӳ��ȣ���ʱ��2�������ӵ�3����ƥ��
          If n_Count > 1 Then
            v_Tmp := Substr(v_�ֵ�, 1, 3);
            Select Max(����)
            Into v_����
            From ����
            Where ���� Like v_Tmp || '%' And Nvl(����, 0) = 3 And �ϼ����� = v_Code����;
          End If;
          If Not v_���� Is Null Then
            v_�ֵ� := Substr(v_�ֵ�, Length(v_����) + 1);
          End If;
        End If;
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
      Select Max(����), Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ), Count(1)
      Into v_��, v_Code��, n_����, n_����ʾ, n_Count
      From ����
      Where ���� Like v_Tmp || '%' And Nvl(����, 0) = 1 And �ϼ����� = v_Codeʡ;
      --���ڶ���ƥ�䣬��������ӳ��ȣ���ʱ��2�������ӵ�3����ƥ��
      If n_Count > 1 Then
        v_Tmp := Substr(v_Adrstmp, 1, 3);
        Select Max(����), Max(����), Max(�Ƿ�����), Max(�Ƿ���ʾ)
        Into v_��, v_Code��, n_����, n_����ʾ
        From ����
        Where ���� Like v_Tmp || '%' And Nvl(����, 0) = 1 And �ϼ����� = v_Codeʡ;
      End If;
      --�ж��Ƿ���������ַ����ʾ�ĵ�ַ���µ�,������ڣ�����ݵ�������ַ��ȷ�������ַ
      If v_Code�� Is Null Then
        Select Max(�Ƿ�����), Max(�Ƿ���ʾ) Into n_����, n_����ʾ From ���� Where �ϼ����� = v_Codeʡ;
        If Nvl(n_����, 0) = 1 Or Nvl(n_����ʾ, 0) = 1 Then
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

--91090:������,2015-12-07,����ʱ�����
Create Or Replace Procedure Zl_������ͨ���_Batchupdate
(
  ����걾id_In In ������ͨ���.����걾id%Type,
  ����id_In     In ������ͨ���.����id%Type := Null,
  �걾����_In   In Varchar2,
  �Ա�_In       In Number,
  ��������_In   In Date,
  ����ָ��_In   In Varchar2, --��ʽ����ĿID^ֵ|������
  ΢����_In     In Number := 0, --1=΢����
  ø���id_In   In ����ø���¼.Id%Type := Null
) Is
  v_��¼���� Number(2);

  v_Temp           Varchar2(255);
  v_��Ա����       ��Ա��.����%Type;
  v_Count          Number;
  v_ҽ��id         ����걾��¼.ҽ��id%Type;
  v_ҩ�����       ����ҩ�����.���%Type;
  v_ϸ��id         ����ϸ��.Id%Type;
  v_������ͨ���id ������ͨ���.Id%Type;
  v_ҩ������       ����ҩ�����.ҩ������%Type;
  v_Od             ������ͨ���.Od%Type;
  v_Cutoff         ������ͨ���.Cutoff%Type;
  v_Sco            ������ͨ���.Sco%Type;

  v_Records   Varchar2(4000);
  v_Currrec   Varchar2(100);
  v_����      ����걾��¼.����%Type;
  v_��Ŀid    ������ͨ���.������Ŀid%Type;
  v_������  ������ͨ���.������%Type;
  v_������1 ������ͨ���.������%Type;
  v_��ʱ���  ������ͨ���.������%Type;

  v_����ִ�     Varchar2(4000);
  v_�������     Number;
  v_���������� Number;
  v_����������� Varchar2(4000);

  v_Resultref  Varchar2(1000);
  v_�ο�ֵ     Varchar2(1000);
  v_�ο�ֵ1    Varchar2(1000);
  v_Σ���ο�   Varchar2(1000);
  v_�����־   Number;
  v_����ֵ     Number;
  v_�����     Number;
  v_С��       Number;
  v_��������   Number;
  v_��������   Number;
  v_��ο�     Number;
  v_�������id Number;

  v_Lower Number;
  v_Upper Number;

  Function Zlval(Vstr In Varchar2) Return Number Is
    Result Number(16, 6);
    Intbit Number(8);
    Strnum Varchar(10);
    Function Sub_Is_Number(v_In In Varchar2) Return Boolean Is
      n_Tmp Number;
    Begin
      n_Tmp := To_Number(v_In);
      If n_Tmp Is Not Null Then
        Return True;
      End If;
    Exception
      When Others Then
        Return False;
    End Sub_Is_Number;
  Begin
    Strnum := '';
    If Sub_Is_Number(Vstr) = True Then
      Result := To_Number(Nvl(Vstr, 0));
      Return(Result);
    Else
      For Intbit In 1 .. 10 Loop
        If Instr('0123456789.', Substr(Vstr, Intbit, 1)) = 0 Then
          Exit;
        End If;
        Strnum := Strnum || Substr(Vstr, Intbit, 1);
        Null;
      End Loop;
      Result := To_Number(Nvl(Strnum, 0));
      Return(Result);
    End If;
  End Zlval;
  -- >>>>>>>>>>>>>>>>>>  ����Ƿ����ֵĺ���  <<<<<<<<<<<<<<<<<<
  Function Sub_Is_Number(v_In In Varchar2) Return Boolean Is
    n_Tmp Number;
  Begin
    n_Tmp := To_Number(v_In);
    If n_Tmp Is Not Null Then
      Return True;
    Else
      Return False;
    End If;
  Exception
    When Others Then
      Return False;
  End Sub_Is_Number;
Begin
  --����Ա��Ϣ:����ID,��������;��ԱID,��Ա���,��Ա����
  v_Temp     := Zl_Identity;
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);

  Select Nvl(������, 0), Nvl(ҽ��id, 0), ����, �������id
  Into v_��¼����, v_ҽ��id, v_����, v_�������id
  From ����걾��¼
  Where ID = ����걾id_In;
  If Sql%Rowcount > 0 Then
    v_Records := ����ָ��_In || '|';
    While v_Records Is Not Null Loop
      If Nvl(΢����_In, 0) = 0 Then
        --��ͨ�걾����
        v_Currrec := Substr(v_Records, 1, Instr(v_Records, '|') - 1);
        If Instr(v_Currrec, '<Split>') > 0 Then
          v_����ִ� := Substr(v_Currrec, Instr(v_Currrec, '<Split>') + 7);
          v_Currrec  := Substr(v_Currrec, 1, Instr(v_Currrec, '<Split>') - 1);
        End If;
        v_��Ŀid := To_Number(Substr(v_Currrec, 1, Instr(v_Currrec, '^') - 1));
        v_Temp   := Substr(v_Currrec, Instr(v_Currrec, '^') + 1);
        If Instr(v_Temp, '^') > 0 Then
          v_������ := Substr(v_Temp, 1, Instr(v_Temp, '^') - 1);
          v_Temp     := Substr(v_Temp, Instr(v_Temp, '^') + 1);
          v_Od       := Substr(v_Temp, 1, Instr(v_Temp, '^') - 1);
          v_Temp     := Substr(v_Temp, Instr(v_Temp, '^') + 1);
          v_Cutoff   := Substr(v_Temp, 1, Instr(v_Temp, '^') - 1);
          v_Sco      := Substr(v_Temp, Instr(v_Temp, '^') + 1);
        Else
          v_������ := Substr(v_Currrec, Instr(v_Currrec, '^') + 1);
          v_Od       := Null;
          v_Cutoff   := Null;
          v_Sco      := Null;
        End If;
        If v_����ִ� Is Not Null Then
          If Instr(v_����ִ�, '^') > 0 Then
            v_�������     := Substr(v_����ִ�, 1, Instr(v_����ִ�, '^') - 1);
            v_����ִ�     := Substr(v_����ִ�, Instr(v_����ִ�, '^') + 1);
            v_����������� := v_����ִ�;
          Else
            v_�������     := v_����ִ�;
            v_����������� := Null;
          End If;
        End If;
        v_С�� := 2;
        Select b.��������, b.��������, Max(a.����ֵ), Max(a.�����), Max(Nvl(a.С��λ��, 2))
        Into v_��������, v_��������, v_����ֵ, v_�����, v_С��
        From ����������Ŀ A, ������Ŀ B
        Where a.��Ŀid = b.������Ŀid And a.����id = ����id_In And a.��Ŀid = v_��Ŀid
        Group By b.��������, b.��������;
        If v_����ֵ Is Null Then
          v_����ֵ := 0;
        End If;
        If v_����� Is Null Then
          v_����� := 1;
        End If;
      
        If Instr(v_������, '+') = 0 Then
          v_������1 := v_������;
          Begin
            If v_����ֵ <> 0 Or v_����� <> 1 Then
              v_������ := (v_������ + v_����ֵ) * v_�����;
            
            End If;
            If Instr(v_������, 'E') = 0 Then
              If Zlval(v_������) = v_������ Then
                If v_С�� = 0 Then
                  v_������ := Trim(To_Char(To_Number(Nvl(Trim(v_������), 0)), '999999999'));
                Else
                  v_������ := Trim(To_Char(To_Number(Nvl(Trim(v_������), 0)), '999999990' || Substr('.000000', 1, 1 + v_С��)));
                End If;
              End If;
            End If;
          Exception
            When Others Then
              v_������ := v_������1;
          End;
        End If;
      
        --��ȡ�ο����жϽ����־
        v_Resultref := Zlgetreference(v_��Ŀid, �걾����_In, �Ա�_In, ��������_In, ����id_In, v_����, v_�������id);
        v_�ο�ֵ1   := v_Resultref;
        Select Nvl(��ο�, 0) Into v_��ο� From ������Ŀ Where ������Ŀid = v_��Ŀid;
        If Instr(v_Resultref, Chr(13) || Chr(10)) > 0 Then
          v_Resultref := Substr(v_Resultref, 1, Instr(v_Resultref, Chr(13) || Chr(10)) - 1);
        Else
          v_��ο� := 0;
        End If;
        v_Σ���ο� := Zl_Get_Reference(2, v_��Ŀid, �걾����_In, �Ա�_In, ��������_In, ����id_In, v_����, v_�������id);
      
        v_�����־ := 1;
        v_��ʱ��� := v_������;
        v_������ := Replace(Replace(v_������, '>', ''), '<', '');
        --����">"��"<"���ж�
        --If Instr(v_������, '>') > 0 Then
        --  v_�ο�ֵ   := v_Resultref;
        --  v_�����־ := 3;
        --End If;
        --If Instr(v_������, '<') > 0 Then
        --  v_�ο�ֵ   := v_Resultref;
        --  v_�����־ := 2;
        --End If;
      
        If v_�����־ = 1 Then
          If (Instr(v_������, '+') > 0 Or Instr(v_������, '*') > 0) And Sub_Is_Number(v_������) = False Then
            v_�ο�ֵ   := v_Resultref;
            v_�����־ := 4;
          Else
            If v_Resultref Is Null Or Sub_Is_Number(v_������) = False Then
              v_�ο�ֵ   := Nvl(v_Resultref, '');
              v_�����־ := 1;
            Else
              v_�ο�ֵ := Nvl(v_Resultref, '');
              If Length(v_Resultref) > 0 Then
                If Instr(v_Resultref, '��') > 0 Then
                  If Instr(v_Resultref, '��') < Length(v_Resultref) Then
                    v_Upper := Zlval(Nvl(Substr(v_Resultref, Instr(v_Resultref, '��') + 1), 0));
                  Else
                    v_Upper := 0;
                  End If;
                  v_Lower := Zlval(Nvl(Substr(v_Resultref, 1, Instr(v_Resultref, '��') - 1), 0));
                Else
                  v_Upper := Zlval(v_Resultref);
                  v_Lower := Zlval(v_Resultref);
                End If;
                If Nvl(v_������, 0) > v_Upper And v_Upper <> 0 Then
                  v_�����־ := 3;
                Else
                  If Nvl(v_������, 0) < v_Lower And v_Lower <> 0 Then
                    v_�����־ := 2;
                  Else
                    v_�����־ := 1;
                  End If;
                End If;
                If v_�����־ <> 1 Then
                  If Sub_Is_Number(v_������) = True Then
                    If Instr(v_Σ���ο�, '��') > 0 Then
                    
                      If Nvl(Zlval(v_������), 0) < To_Number(Substr(v_Σ���ο�, 1, Instr(v_Σ���ο�, '��') - 1)) Then
                        v_�����־ := 5;
                      End If;
                    
                      If Nvl(Zlval(v_������), 0) > To_Number(Substr(v_Σ���ο�, 1, Instr(v_Σ���ο�, '��') - 1)) Then
                        v_�����־ := 6;
                      End If;
                    
                    End If;
                  End If;
                End If;
              Else
                v_�����־ := 1;
                If Sub_Is_Number(v_������) = True Then
                  If Instr(v_Σ���ο�, '��') > 0 Then
                  
                    If Nvl(Zlval(v_������), 0) < To_Number(Substr(v_Σ���ο�, 1, Instr(v_Σ���ο�, '��') - 1)) Then
                      v_�����־ := 5;
                    End If;
                  
                    If Nvl(Zlval(v_������), 0) > To_Number(Substr(v_Σ���ο�, 1, Instr(v_Σ���ο�, '��') - 1)) Then
                      v_�����־ := 6;
                    End If;
                  
                  End If;
                End If;
              End If;
            End If;
          End If;
        End If;
        v_������ := v_��ʱ���;
        Update ������ͨ���
        Set ������ = v_������, �����־ = Decode(v_��ο�, 0, v_�����־, 1), �޸��� = Decode(ԭʼ���, Null, Null, v_��Ա����),
            �޸�ʱ�� = Decode(ԭʼ���, Null, Null, Sysdate), ԭʼ��� = Decode(ԭʼ���, Null, v_������, ԭʼ���),
            ԭʼ��¼ʱ�� = Decode(ԭʼ���, Null, Sysdate, ԭʼ��¼ʱ��), ��¼�� = Decode(ԭʼ���, Null, v_��Ա����, ��¼��), ����id = ����id_In,
            ����ο� = Decode(v_��ο�, 0, v_�ο�ֵ, v_�ο�ֵ1), Od = v_Od, Cutoff = v_Cutoff, Sco = v_Sco, ø���id = ø���id_In
        Where ����걾id = ����걾id_In And ������Ŀid = v_��Ŀid And ��¼���� = v_��¼����;
      
        If Sql%Rowcount = 0 Then
          Insert Into ������ͨ���
            (ID, ����걾id, ������Ŀid, ������, �����־, ��¼����, ԭʼ���, ԭʼ��¼ʱ��, ��¼��, ����id, ����ο�, Od, Cutoff, Sco, ø���id)
          Values
            (������ͨ���_Id.Nextval, ����걾id_In, v_��Ŀid, v_������, Decode(v_��ο�, 0, v_�����־, 1), 0, v_������, Sysdate, v_��Ա����,
             ����id_In, Decode(v_��ο�, 0, v_�ο�ֵ, v_�ο�ֵ1), v_Od, v_Cutoff, v_Sco, ø���id_In);
        End If;
      
        Update ������ˮ��ָ��
        Set �����Ƿ���� = v_�������, ������� = v_�����������
        Where �걾id = ����걾id_In And ��Ŀid = v_��Ŀid;
      
        If Sql%Rowcount = 0 Then
          Insert Into ������ˮ��ָ��
            (ID, �걾id, ��Ŀid, �����Ƿ����, �������)
          Values
            (������ˮ��ָ��_Id.Nextval, ����걾id_In, v_��Ŀid, v_�������, v_�����������);
        End If;
      
        Select Count(*) Into v_Count From ������Ŀ�ֲ� Where �걾id = ����걾id_In And ��Ŀid + 0 = v_��Ŀid;
        If v_Count = 0 Then
          Insert Into ������Ŀ�ֲ�
            (ID, �걾id, ��Ŀid, ҽ��id, ��Χ)
          Values
            (������Ŀ�ֲ�_Id.Nextval, ����걾id_In, v_��Ŀid, Decode(v_ҽ��id, 0, Null, v_ҽ��id), 1);
        End If;
      Else
        --����΢����
        v_Currrec := Substr(v_Records, 1, Instr(v_Records, '|') - 1);
        If Instr(v_Currrec, '<Split>') > 0 Then
          v_����ִ� := Substr(v_Currrec, Instr(v_Currrec, '<Split>') + 7);
          v_Currrec  := Substr(v_Currrec, 1, Instr(v_Currrec, '<Split>') - 1);
        End If;
        v_Temp     := v_Currrec;
        v_��Ŀid   := To_Number(Substr(v_Temp, 1, Instr(v_Temp, '^') - 1));
        v_Temp     := Substr(v_Temp, Instr(v_Temp, '^') + 1);
        v_������ := Substr(v_Temp, 1, Instr(v_Temp, '^') - 1);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, '^') + 1);
        v_ҩ������ := Substr(v_Temp, 1, Instr(v_Temp, '^') - 1);
        v_Temp     := Substr(v_Temp, Instr(v_Temp, '^') + 1);
        v_ҩ����� := v_Temp;
        If v_����ִ� Is Not Null Then
          If Instr(v_����ִ�, '^') > 0 Then
            v_�������     := Substr(v_����ִ�, 1, Instr(v_����ִ�, '^') - 1);
            v_����ִ�     := Substr(v_����ִ�, Instr(v_����ִ�, '^') + 1);
            v_����������� := v_����ִ�;
          Else
            v_�������     := v_����ִ�;
            v_����������� := Null;
          End If;
        End If;
      
        Begin
          Select Distinct ID
          Into v_ϸ��id
          From ����ϸ�� A, ����ϸ������ B
          Where a.Id = b.ϸ��id(+) And (a.������ = �걾����_In Or a.Ӣ���� = �걾����_In Or a.���� = �걾����_In Or b.ͨ������ = �걾����_In);
        Exception
          When Others Then
            Return;
        End;
        If Sql%Rowcount > 0 Then
          Update ������ͨ���
          Set �޸��� = Decode(ԭʼ���, Null, Null, v_��Ա����), �޸�ʱ�� = Decode(ԭʼ���, Null, Null, Sysdate),
              ��¼�� = Decode(ԭʼ���, Null, v_��Ա����, ��¼��), ����id = ����id_In, ��¼���� = v_��¼����
          Where ����걾id = ����걾id_In And ϸ��id = v_ϸ��id;
        
          If Sql%Rowcount = 0 Then
            Select ������ͨ���_Id.Nextval Into v_������ͨ���id From Dual;
            Insert Into ������ͨ���
              (ID, ����걾id, ϸ��id, ԭʼ��¼ʱ��, ��¼��, ����id, ��¼����)
            Values
              (v_������ͨ���id, ����걾id_In, v_ϸ��id, Sysdate, v_��Ա����, ����id_In, v_��¼����);
          Else
            Select ID Into v_������ͨ���id From ������ͨ��� Where ����걾id = ����걾id_In And ϸ��id = v_ϸ��id;
          End If;
          --------------�ݲ�����΢������ˮ������-----------------------------------------------------------       
          --         Update ������ˮ��ָ��
          --          Set �����Ƿ���� = v_�������, ������� = v_�����������
          --          Where �걾id = ����걾id_In And ��Ŀid = v_��Ŀid;
          --       
          --          If Sql%Rowcount = 0 Then
          --            Insert Into ������ˮ��ָ��
          --              (ID, �걾id, ��Ŀid, �����Ƿ����, �������)
          --            Values
          --             (������ˮ��ָ��_Id.Nextval, ����걾id_In, v_��Ŀid, v_�������, v_�����������);
          --          End If;
          --------------------------------------------------------------------------------------------------       
          Select Count(*) Into v_Count From ������Ŀ�ֲ� Where �걾id = ����걾id_In And ��Ŀid + 0 = v_ϸ��id;
          If v_Count = 0 Then
            Insert Into ������Ŀ�ֲ�
              (ID, �걾id, ϸ��id, ҽ��id, ��Χ)
            Values
              (������Ŀ�ֲ�_Id.Nextval, ����걾id_In, v_ϸ��id, Decode(v_ҽ��id, 0, Null, v_ҽ��id), 1);
          End If;
          If Nvl(v_��Ŀid, 0) <> 0 Then
            Update ����ҩ�����
            Set �޸��� = v_��Ա����, �޸�ʱ�� = Sysdate, ��� = v_ҩ�����, ������� = v_������, ����id = ����id_In, ҩ������ = v_ҩ������
            Where ϸ�����id = v_������ͨ���id And ������id = v_��Ŀid;
          
            If Sql%Rowcount = 0 Then
              Insert Into ����ҩ�����
                (ϸ�����id, ������id, �޸���, �޸�ʱ��, ���, �������, ��¼����, ����id, ҩ������)
              Values
                (v_������ͨ���id, v_��Ŀid, v_��Ա����, Sysdate, v_ҩ�����, v_������, 0, ����id_In, v_ҩ������);
            End If;
          End If;
        End If;
      End If;
      --v_Records := Replace('|' || v_Records, '|' || v_Currrec || '|');;
      v_Records := Substr(v_Records, Instr(v_Records, '|') + 1);
    End Loop;
    If ����ָ��_In Is Not Null Then
      Update ����걾��¼ Set ����ʱ�� = Sysdate Where ID = ����걾id_In;
    End If;
    If Nvl(΢����_In, 0) = 0 Then
      Select Count(*)
      Into v_����������
      From ������ˮ��ָ��
      Where �걾id = ����걾id_In And Nvl(�����Ƿ����, 0) = 0;
      If v_���������� = 0 Then
        Update ������ˮ�߱걾 Set �����Ƿ���� = 1 Where �걾id = ����걾id_In;
        If Sql%Rowcount = 0 Then
          Insert Into ������ˮ�߱걾 (ID, �걾id, �����Ƿ����) Values (������ˮ��ָ��_Id.Nextval, ����걾id_In, 1);
        End If;
      Else
        Update ������ˮ�߱걾 Set �����Ƿ���� = 0 Where �걾id = ����걾id_In;
        If Sql%Rowcount = 0 Then
          Insert Into ������ˮ�߱걾 (ID, �걾id, �����Ƿ����) Values (������ˮ��ָ��_Id.Nextval, ����걾id_In, 0);
        End If;
      End If;
    End If;
  End If;

Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_������ͨ���_Batchupdate;
/


--89666:����,2015-12-03,��ʾ���洴��ʱ������ʱ��
--Ӱ�񱨸�ҵ��(---���岿��---)***************************************************
CREATE OR REPLACE Package b_PACS_RptManage Is
  Type t_Refcur Is Ref Cursor;

  --1������������
  Procedure p_Edit_Doc_Lockinfo(
    ����_Id_In Ӱ�񱨸��¼.Id%Type,
	������_In  Ӱ�񱨸��¼.������%Type
	);

  --2��������������
  Procedure p_Edit_Doc_EvaluatRptQuality(
    ����Id_In Ӱ�񱨸��¼.Id%Type,
	�����ȼ�_In Ӱ�񱨸��¼.��������%Type
	);
                                
  --3������������
  Procedure p_Edit_Doc_EvaluatResult(
    ����Id_In Ӱ�񱨸��¼.Id%Type,
	�����_In Ӱ�񱨸��¼.�������%Type
	);
                                
  --4�����淢��/����
  Procedure p_Edit_Doc_ReportRelease(
    ����Id_In Ӱ�񱨸��¼.Id%Type,
	��ǰ������_In Ӱ�񱨸��¼.���淢����%Type
	);

 --5���������޸ı���
  Procedure p_Ӱ�񱨸��¼_����(
    ԭ��ID_In     Ӱ�񱨸��¼.ԭ��ID%Type,
    ��������_In   Ӱ�񱨸��¼.��������%Type,
    ��¼��_In     Ӱ�񱨸��¼.��¼��%Type,
    ���༭��_In Ӱ�񱨸��¼.���༭��%Type,
    Id_In         Ӱ�񱨸��¼.Id%Type,
    ҽ��ID_In     Ӱ�񱨸��¼.ҽ��ID%Type 
	);

  --6����ȡ��д���ĵ�����
  Procedure p_Get_Doc_Content(
    Val           Out t_Refcur,
	DocID_In Ӱ�񱨸��¼.Id%Type
	);

  --7�����ñ����ӡ������Ϣ
  Procedure p_Checkrejectsignature(Signdate_In Date,
                                   ����ID_In   Ӱ�񱨸������¼.����Id%Type,
                                   ������_In   Ӱ�񱨸������¼.������%Type,
                                   ����˵��_In Ӱ�񱨸������¼.����˵��%Type,
                                   Val         Out Sys_Refcursor);

  --8����ѯ��Ӧԭ���µ�������
  Procedure p_Get_Samplelist_Maxseqnum(
    Val           Out t_Refcur,
	ԭ��ID_In Ӱ�񱨸淶���嵥.ԭ��ID%Type
	);

  --9��ɾ���ĵ�����
  Procedure p_Del_Ӱ�񱨸淶���嵥(
    Id_In Ӱ�񱨸淶���嵥.Id%Type
	);

 --10������ĵ��Ĳ�����־
 Procedure p_Ӱ�񱨸������¼_Add(Id_In       Ӱ�񱨸������¼.Id%Type,
                               ����ID_In   Ӱ�񱨸������¼.����ID%Type,                               
                               ������_In   Ӱ�񱨸������¼.������%Type,                               
                               ��������_In Ӱ�񱨸������¼.��������%Type);

  --11��ɾ������
  Procedure p_Ӱ�񱨸��¼_ɾ��(
    ����_Id_In Ӱ�񱨸��¼.Id%Type
	);

  --12����ȡǩ������
  Procedure p_Get_SysConfigSignature(
    Val           Out t_Refcur,
	����ID_In		In ���ű�.ID%Type
	);

--13����ȡ�˻�ǩ��ӡ��
Procedure p_Get_PersonSignImg(
  Val           Out t_Refcur,
  ID_In		In ��Ա��.ID%Type
  );


--14����ȡǩ����֤����Ϣ
Procedure p_Get_SignCertInfo(
  Val           Out t_Refcur,
  ֤��ID_In		��Ա֤���¼.ID%Type
  );

--15�����±���״̬
Procedure p_Update_ReportState(
  ����Id_In  Ӱ�񱨸��¼.ID%Type,
  ����״̬_In  Ӱ�񱨸��¼.����״̬%Type,
  �����_In   Ӱ�񱨸��¼.��������%Type
  );

--16����ȡ����״̬
Procedure p_Get_ReportState(
  Val           Out t_Refcur,
  ����Id_In	Ӱ�񱨸��¼.ID%Type
  );

--17�����沵��
Procedure p_Reject_Report(
  ҽ��ID_In	Ӱ�񱨸沵��.ҽ��ID%Type, 
  ����ID_In	Ӱ�񱨸沵��.��鱨��ID%Type, 
  ��������_In Ӱ�񱨸沵��.��������%Type, 
  ����ʱ��_In Ӱ�񱨸沵��.����ʱ��%Type, 
  ������_In   Ӱ�񱨸沵��.������%Type,
  ��������_In  Ӱ�񱨸��¼.��������%Type,
  ����״̬_In Ӱ�񱨸��¼.����״̬%Type
  );

--17.1���������沵��
Procedure p_Reject_Cancel(
  ID_In       Ӱ�񱨸沵��.ID%Type,
  ����ID_In    Ӱ�񱨸沵��.��鱨��ID%Type,
  ����״̬_In   Ӱ�񱨸��¼.����״̬%Type
  );

--18����ȡ���沵����Ϣ
Procedure p_Get_RejectInfo(
  Val           Out t_Refcur,
  ����ID_In	Ӱ�񱨸沵��.��鱨��ID%Type
  );

--19����ȡԭ�Ͷ���
Procedure p_Get_Doc_Process(
  Val           Out t_Refcur,
  ԭ��id_In Ӱ�񱨸涯��.ԭ��id%Type
  );

--20��ͨ��ѧ��ɸѡ�����Ӧ�ķ�����Ϣ
  Procedure p_Get_Samplelist_By_Conditions(
    Val           Out t_Refcur,
    ԭ��id_In       Varchar2,
    ѧ��_In  Varchar2,
    Condition_In Varchar2, --����ɸѡ
    ����_In    Varchar2
    );

  --21��ͨ������ID��ȡ��������
  Procedure p_Get_��������_By_ID(
    Val           Out t_Refcur,
    ID_IN ���ű�.ID%TYPE
    );

  --22����ȡ����Ԥ�����
  Procedure p_Get_AllPreOutlines(
    Val           Out t_Refcur
    );

  --23����ȡ�ĵ�����
  Procedure p_Get_reportTitle_By_ID(
    Val           Out t_Refcur,
    ID_IN  Ӱ�񱨸��¼.id%TYPE   
    );
  
  --24����ȡ����������
  Procedure p_Get_����������_By_ID(
    Val           Out t_Refcur,
    ID_IN  Ӱ�񱨸��¼.id%TYPE   
    );

  --25��ͨ��ҽ��ID��ȡ�����б�
  Procedure p_Get_Ӱ�񱨸��¼_By_ҽ��ID(
    Val           Out t_Refcur,
    ҽ��ID_IN  Ӱ�񱨸��¼.ҽ��ID%TYPE   
    );
  
  --26����ѯӰ�����̲���ֵ
  Procedure p_Get_Ӱ�����̲���ֵ(
    Val           Out t_Refcur,  
    ����ID_IN  Ӱ�����̲���.����ID%TYPE
    );

  --27������ҽ��ID����ѯ��Ӧ��ԭ���б�
  Procedure p_Get_Ӱ��ԭ���б�_By_ҽ��ID(
    Val           Out t_Refcur,
    ҽ��_IN  Ӱ�����¼.ҽ��ID%TYPE   
    );

  --28�����ݱ���ID��ѯ��ӡ��¼
  procedure p_Get_ReportPrintLog_By_����ID
  (
       val out sys_refcursor  ,
       ����_IN  Ӱ�񱨸������¼.����ID%TYPE
  );

  --29������ҽ��ID��ѯ���淢���б�
  Procedure p_Get_ReportReleaseList(
    Val           Out t_Refcur,
    ҽ��_IN  Ӱ�񱨸��¼.ҽ��ID%TYPE   
    );

  --30�����ݱ���ID��ѯ���ؼ�¼����
  Procedure p_Get_RejectedCount(
    Val           Out t_Refcur,
    ����_IN  Ӱ�񱨸沵��.��鱨��ID%TYPE
    );

  --31������ҽ��ID��ѯ���涯����Ҫ��һЩID��
  Procedure p_Get_DocProcess_IDs(
    Val           Out t_Refcur,
    ҽ��_IN  ����ҽ����¼.ID%TYPE
    );

  --32������ҽ��ID�ͱ���ID��ѯ�����һЩ����
  Procedure p_Get_DocInfo(
    Val           Out t_Refcur,
    ҽ��ID_IN  Ӱ�����¼.ҽ��ID%TYPE,
    ����ID_IN  Ӱ�񱨸��¼.ID%TYPE
    );
  
  --33����ѯһ���������ͬԭ��ID�ı�������
   Procedure p_Get_SameAntetypeDocCounts(
       Val           Out t_Refcur,
       ҽ��ID_IN  Ӱ�񱨸��¼.ҽ��ID%TYPE,
       ԭ��ID_IN  Ӱ�񱨸��¼.ԭ��ID%TYPE
  );

  --34����ȡ����ͼ�洢��Ϣ
  Procedure p_Get_DocImageSaveInof_By_ID(
    Val           Out t_Refcur,
	  ID_IN  Ӱ�񱨸��¼.id%TYPE
    );

end b_PACS_RptManage;
/

--Ӱ�񱨸�ҵ��(---ʵ�ֲ���---)***************************************************

CREATE OR REPLACE Package Body b_PACS_RptManage Is

  --1������������
  Procedure p_Edit_Doc_Lockinfo(
    ����_Id_In Ӱ�񱨸��¼.Id%Type,
	������_In  Ӱ�񱨸��¼.������%Type
	) Is
  Begin
  
    --  ����IDΪ�գ���������С�������_In�����������ı��
    If ����_Id_In Is Null Then
      Update Ӱ�񱨸��¼ A Set a.������ = '' Where a.������ = ������_In;
    Else
      Update Ӱ�񱨸��¼ A
         Set a.������ = ������_In
       Where a.Id = ����_Id_In;
    End If;
  
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Edit_Doc_Lockinfo;

  --2��������������
  Procedure p_Edit_Doc_EvaluatRptQuality(
    ����Id_In Ӱ�񱨸��¼.Id%Type,
	�����ȼ�_In  Ӱ�񱨸��¼.��������%Type
	) Is
  Begin
    Update Ӱ�񱨸��¼ Set �������� = �����ȼ�_In Where Id = ����Id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Edit_Doc_EvaluatRptQuality;
  
  --3������������
  Procedure p_Edit_Doc_EvaluatResult(
    ����Id_In Ӱ�񱨸��¼.Id%Type,
	�����_In Ӱ�񱨸��¼.�������%Type
	) Is
  Begin
     Update Ӱ�񱨸��¼ Set ������� = �����_In Where Id = ����Id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Edit_Doc_EvaluatResult;
  
  --4�����淢��/����
  Procedure p_Edit_Doc_ReportRelease(
    ����Id_In Ӱ�񱨸��¼.Id%Type,
	��ǰ������_In Ӱ�񱨸��¼.���淢����%Type
	) Is
    v_���淢��     Ӱ�񱨸��¼.���淢��%Type; 
  Begin
    
    Begin 
		  Select nvl(���淢��,0) Into v_���淢�� From Ӱ�񱨸��¼ where ID=����Id_In; 
    Exception 
      When Others Then 
        v_���淢�� :=0; 
    End; 
     
    Update Ӱ�񱨸��¼ Set ���淢�� =decode(v_���淢��,0,1,0),���淢����=decode(v_���淢��,0,��ǰ������_In,'') Where ID=����Id_In; 
     
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Edit_Doc_ReportRelease; 

 --5���������޸ı���
  Procedure p_Ӱ�񱨸��¼_����(
    ԭ��ID_In     Ӱ�񱨸��¼.ԭ��ID%Type,
    ��������_In   Ӱ�񱨸��¼.��������%Type,
    ��¼��_In     Ӱ�񱨸��¼.��¼��%Type,
    ���༭��_In Ӱ�񱨸��¼.���༭��%Type,
    Id_In         Ӱ�񱨸��¼.Id%Type,
    ҽ��ID_In     Ӱ�񱨸��¼.ҽ��ID%Type
  ) As
    --ԭ��ID_In ԭ��ID
    --�����ĵ���д��¼
    --1 ������������
    --2 �����ĵ���д��¼��״̬
    --3 ����༭��־
    --4 �����ĵ�����
    v_����id    Ӱ�񱨸��¼.Id%Type;
    v_ԭ������  Ӱ�񱨸�ԭ���嵥.����%Type;
    v_�豸��    Ӱ�񱨸�ԭ���嵥.�豸��%Type;
    v_�������  number;
    x_Editlog   Xmltype;
    Cur_Time    Date;
    To_Editlist t_Editlist;
    Tn_Editlist t_Editlist;
    v_Msg       Varchar2(200);
    v_New       number;
    Err_Custom  Exception;
    v_Result    Ӱ�񱨸��¼.������%Type;
    v_����ID    Ӱ�񱨸������¼.ID%Type;

    Function Elist_Filter(
    Source_t t_Editlist
    ) Return t_Editlist Is
      Target_t t_Editlist := t_Editlist();
    Begin

      --�Զ����ĵ���˵���������ֻ�ǽ� Source_t���ձ༭ʱ����������
      For Rs In (Select /*+rule*/
                  *
                   From Table(Cast(Source_t As t_Editlist)) A
                  Order By a.�༭ʱ��) Loop
        Target_t.Extend;
        Target_t(Target_t.Count) := t_Edits(Rs.�༭��,
                                            Rs.�༭ʱ��,
                                            Rs.ǩ��,
                                            Rs.��ǩ��);
      End Loop;
      Return Target_t;
    End;

    Function Build_Editlog(
    Tn_Edit t_Editlist,
    To_Edit t_Editlist,
    v_Did   Ӱ�񱨸��¼.Id%Type) Return Xmltype Is
      --Tn_Edit ���α�����±༭��¼��To_Edit�ϴα���ľɱ༭��¼
      --�����α༭��¼����ϳ�һ���༭��¼

      x_Return Xmltype;
      r_Saveid Raw(16);
      n_Class  Number;
      --n_Class �༭��־�еĲ������ 1-������2-ɾ����3-�༭��4-ǩ����5-�󶩡�6-��ǩ��7-��ǩ
      v_Signor  Ӱ�񱨸��¼.������%Type;
      v_Adjunct Ӱ�񱨸��¼.������%Type;
      Tns_Edit  t_Editlist;
      Tos_Edit  t_Editlist;

      Function Atitle(ԭ��ID Ӱ�񱨸�ԭ���嵥.Id%Type) Return Varchar2 Is
        v_ԭ������ Ӱ�񱨸�ԭ���嵥.����%Type;
      Begin
        --����ԭ��ID������ԭ������
        If ԭ��ID Is Null Then
          Return Null;
        Else
          Select ���� Into v_ԭ������ From Ӱ�񱨸�ԭ���嵥 Where ID = ԭ��ID;
          Return v_ԭ������;
        End If;
      End;

    Begin
      x_Return := Xmltype('<root></root>');
      If v_Did Is Null Then
        --�����������ĵ��������ĵ���null����
        Select Sys_Guid() Into r_Saveid From Dual;

        --PACS����û�����ĵ����������湹��XML����䱣���ɸ�EMR��ͬ�������v_Subiid��ֵΪ��
        Tns_Edit := Elist_Filter(Tn_Edit);
        Select Decode(Tns_Edit(Tns_Edit.Count).ǩ��, 0, 1, 4)
          Into n_Class
          From Dual;
        Select Appendchildxml(x_Return,
                              '/root',
                              Xmlelement("operate",
                                         Xmlforest(r_Saveid As "saving_id",
                                                   n_Class As "class",
                                                   To_Char(Cur_Time,
                                                           'yyyy-mm-dd hh24:mi:ss') As
                                                   "cur_time",
                                                   ���༭��_In As "operator",
                                                   Decode(n_Class,
                                                          4,
                                                          Tns_Edit(Tns_Edit.Count).�༭��,
                                                          '') As "signer",
                                                   '' As Adjunct)))
          Into x_Return
          From Dual;
      Else
        --�����������ĵ���
        Select Sys_Guid() Into r_Saveid From Dual;

        v_Signor  := '';
        v_Adjunct := '';
        Tns_Edit  := Elist_Filter(Tn_Edit);
        Tos_Edit  := Elist_Filter(To_Edit);
        If Tns_Edit(Tns_Edit.Count)
         .ǩ�� = 1 And Tns_Edit(Tns_Edit.Count).��ǩ�� = 0 Then
          --���һ����ǩ��
          If Tos_Edit.Count = 0 Then
            --�������ĵ�ֱ��ǩ��
            n_Class  := 4;
            v_Signor := Tns_Edit(Tns_Edit.Count).�༭��;
          Elsif Tos_Edit(Tos_Edit.Count).�༭�� Is Null Then
            --֮ǰûǩ��
            n_Class  := 4;
            v_Signor := Tns_Edit(Tns_Edit.Count).�༭��;
          Elsif Tos_Edit(Tos_Edit.Count)
           .ǩ�� = 1 And Tns_Edit(Tns_Edit.Count)
                .�༭ʱ�� > Tos_Edit(Tos_Edit.Count).�༭ʱ�� Then
            --�����ͨǩ��
            n_Class  := 4;
            v_Signor := Tns_Edit(Tns_Edit.Count).�༭��;
          Elsif Tos_Edit(Tos_Edit.Count)
           .ǩ�� = 1 And Tns_Edit(Tns_Edit.Count)
                .�༭ʱ�� < Tos_Edit(Tos_Edit.Count).�༭ʱ�� Then
            --�������ǩ��
            n_Class   := 7;
            v_Adjunct := Tos_Edit(Tos_Edit.Count).�༭��;
          Elsif Tos_Edit(Tos_Edit.Count)
           .ǩ�� = 1 And Tns_Edit(Tns_Edit.Count)
                .�༭ʱ�� = Tos_Edit(Tos_Edit.Count).�༭ʱ�� Then
            --�ޱ仯
            n_Class := -1;
          End If;
        Elsif Tns_Edit(Tns_Edit.Count)
         .ǩ�� = 1 And Tns_Edit(Tns_Edit.Count).��ǩ�� = 1 Then
          --��ǩ��
          If Tos_Edit(Tos_Edit.Count).��ǩ�� = 0 Then
            --֮ǰû��ǩ����������ǩ��������
            n_Class  := 6;
            v_Signor := Tns_Edit(Tns_Edit.Count).�༭��;
          Elsif Tos_Edit(Tos_Edit.Count)
           .��ǩ�� = 1 And Tns_Edit(Tns_Edit.Count)
                .�༭ʱ�� > Tos_Edit(Tos_Edit.Count).�༭ʱ�� Then
            --�����ǩ
            n_Class  := 6;
            v_Signor := Tns_Edit(Tns_Edit.Count).�༭��;
          Elsif Tos_Edit(Tos_Edit.Count)
           .��ǩ�� = 1 And Tns_Edit(Tns_Edit.Count)
                .�༭ʱ�� < Tos_Edit(Tos_Edit.Count).�༭ʱ�� Then
            --���������ǩ
            n_Class   := 7;
            v_Adjunct := Tos_Edit(Tos_Edit.Count).�༭��;
          Elsif Tos_Edit(Tos_Edit.Count)
           .��ǩ�� = 1 And Tns_Edit(Tns_Edit.Count)
                .�༭ʱ�� = Tos_Edit(Tos_Edit.Count).�༭ʱ�� Then
            --�ޱ仯
            n_Class := -1;
          End If;
        Elsif Tns_Edit(Tns_Edit.Count).�༭�� Is Null And Tos_Edit.Count = 0 Then
          n_Class := 1;
        Elsif Tns_Edit(Tns_Edit.Count)
         .�༭�� Is Null And Tos_Edit(Tos_Edit.Count).ǩ�� = 1 Then
          n_Class   := 7;
          v_Adjunct := Tos_Edit(Tos_Edit.Count).�༭��;
        Elsif Tns_Edit(Tns_Edit.Count)
         .�༭�� Is Null And Tos_Edit(Tos_Edit.Count).�༭�� Is Null Then
          n_Class := 3;
        Elsif Tns_Edit(Tns_Edit.Count)
         .��ǩ�� = 0 And Tos_Edit(Tos_Edit.Count).��ǩ�� = 0 Then
          n_Class := 5;
        Elsif Tns_Edit(Tns_Edit.Count)
         .��ǩ�� = 0 And Tos_Edit(Tos_Edit.Count).��ǩ�� = 1 Then
          n_Class   := 7;
          v_Adjunct := Tos_Edit(Tos_Edit.Count).�༭��;
        End If;

        If n_Class <> -1 Then
          Select Appendchildxml(x_Return,
                                '/root',
                                Xmlelement("operate",
                                           Xmlforest(r_Saveid As "saving_id",
                                                     n_Class As "class",
                                                     To_Char(Cur_Time,
                                                             'yyyy-mm-dd hh24:mi:ss') As
                                                     "cur_time",
                                                     ���༭��_In As "operator",
                                                     Decode(n_Class,
                                                            4,
                                                            v_Signor,
                                                            6,
                                                            v_Signor,
                                                            '') As "signer",
                                                     v_Adjunct As Adjunct)))
            Into x_Return
            From Dual;
        End If;

      End If;
      Return x_Return;
    End Build_Editlog;

    Function Get_NextRPTNum(
    AntetypeName Ӱ�񱨸�ԭ���嵥.����%Type,
    Order_ID Ӱ�񱨸��¼.ҽ��Id%Type
    )
      Return Number Is
        v_��� Number;
        v_count Number;
        v_num Number;
      Begin

        v_count :=0;
        v_num :=1;
        loop
             select count(*)+v_num into v_��� from Ӱ�񱨸��¼ where ҽ��ID=Order_ID;
             select count(*) into v_count from Ӱ�񱨸��¼ where ҽ��ID=Order_ID and �ĵ�����=AntetypeName||'_'||v_���;

             if v_count =0 then
               exit;
             end if;

             v_num := v_num +1;
         end loop;

         return v_���;
     End;

  Begin

    Select ����, �豸��,Sysdate
      Into v_ԭ������,v_�豸��, Cur_Time
      From Ӱ�񱨸�ԭ���嵥
     Where ID = ԭ��ID_In;

    --------------------1 �����ĵ���д��¼��״̬--------------------
    --��ȡ�ĵ���ǩ���ͱ༭���������޸ģ���¼
    Tn_Editlist := b_PACS_RptPublic.f_Geteditlist(��������_In);

    --------------------2 ����༭��־--------------------
    select count(*) into v_New from Ӱ�񱨸��¼ where ID=Id_In;

    v_����id := Id_In;
    select zlpub_pacs_ȡ�������byxml (��������_In,'������') into v_Result from dual;
    If v_New=0 Then
      --��������
      To_Editlist := t_Editlist();
      x_Editlog   := Build_Editlog(Tn_Editlist, To_Editlist, Null);

      --ȡ�������
      v_������� := Get_NextRPTNum(v_ԭ������,ҽ��ID_In);

      Insert Into Ӱ�񱨸��¼
        (ID,
         ԭ��ID,
         �ĵ�����,
         ��������,
         ����ʱ��,
         ������,
         ����״̬,
         ���༭ʱ��,
         ���༭��,
         �༭��־,
         ҽ��ID,
         ��¼��,
         ������,
         �豸��)
      Values
        (v_����id,
         ԭ��ID_In,
         v_ԭ������||'_'||v_�������,
         ��������_In,
         Cur_Time,
         ���༭��_In,
         1,
         Cur_Time,
         ���༭��_In,
         x_Editlog,
         ҽ��ID_In,
         ��¼��_In,
         v_Result,
         v_�豸��);
      Insert Into ����ҽ������(ҽ��ID,��鱨��ID)Values(ҽ��ID_In,v_����id);
      
      Select Sys_Guid() Into v_����ID From Dual;
      Insert Into Ӱ�񱨸������¼(ID, ����ID,ҽ��ID,�ĵ�����,������,����ʱ��,��������) 
             Values(v_����ID,v_����id,ҽ��ID_In,v_ԭ������||'_'||v_�������,���༭��_In,sysdate,6);

    Else
      --��ȡ�ļ�ԭʼ�༭��¼,�����ڸ���֮ǰ��ȡ
      Select b_PACS_RptPublic.f_Geteditlist(��������)
        Into To_Editlist
        From Ӱ�񱨸��¼
       Where ID = v_����id;

      x_Editlog := Build_Editlog(Tn_Editlist, To_Editlist, v_����id);
      Select Appendchildxml(�༭��־,
                            '/root',
                            Extract(x_Editlog, '/root/*'))
             Into x_Editlog From Ӱ�񱨸��¼ Where ID = v_����id;

       Update Ӱ�񱨸��¼
                Set ��������     = ��������_In,
                ���༭ʱ�� = Cur_Time,
                ���༭��   = ���༭��_In,
                �༭��־     = x_Editlog,
                ��¼��       =��¼��_In,
                ������     =v_Result
                Where ID = v_����id;
       end if;

  Exception
    When Err_Custom Then
      Raise_Application_Error(-20101, v_Msg);
    When Others Then
      Zl_Errorcenter(SQLCode, SQLErrM);
  End p_Ӱ�񱨸��¼_����;

  --6����ȡ��д���ĵ�����
  Procedure p_Get_Doc_Content(
    Val           Out t_Refcur,
	DocID_In Ӱ�񱨸��¼.Id%Type
	) As
  Begin
    Open Val For
      Select  Nvl(a.��������.GetClobVal(), '<ZLXML/>') As �������� From Ӱ�񱨸��¼ A Where a.Id = DocID_In;
  End;

  --7�����ñ����ӡ������Ϣ
  Procedure p_Checkrejectsignature(Signdate_In Date,
                                   ����ID_In   Ӱ�񱨸������¼.����Id%Type,
                                   ������_In   Ӱ�񱨸������¼.������%Type,
                                   ����˵��_In Ӱ�񱨸������¼.����˵��%Type,
                                   Val         Out Sys_Refcursor) As
  Begin
    Open Val For
      Select ������, ����ʱ��
        From Ӱ�񱨸������¼
       Where ����ID = ����ID_In
         And ��������=1
         And ����ʱ�� >= Signdate_In
         And ����ʱ�� Is Null
       Order By ����ʱ�� Asc;
    --���ϴ�ӡ��¼
    Update Ӱ�񱨸������¼ B
       Set ������ = ������_In, ����ʱ�� = Sysdate, b.����˵�� = ����˵��_In
     Where ����ID = ����ID_In And ��������=1
       And ����ʱ�� >= Signdate_In;

  End p_Checkrejectsignature;

  --8����ѯ��Ӧԭ���µ�������
  Procedure p_Get_Samplelist_Maxseqnum(
    Val           Out t_Refcur,
	ԭ��ID_In Ӱ�񱨸淶���嵥.ԭ��ID%Type
	) As
  Begin
    Open Val For
      Select Nvl(Max(a.���), 0) + 1 As Num
        From Ӱ�񱨸淶���嵥 A
       Where a.ԭ��ID = ԭ��ID_In;
  End;

  --9��ɾ���ĵ�����
  Procedure p_Del_Ӱ�񱨸淶���嵥(
    Id_In Ӱ�񱨸淶���嵥.Id%Type
	) As
  Begin
    Delete From Ӱ�񱨸淶���嵥 Where Id = Id_In;
  End;
  
 --10������ĵ��Ĳ�����־
  Procedure p_Ӱ�񱨸������¼_Add(Id_In       Ӱ�񱨸������¼.Id%Type,
                               ����ID_In   Ӱ�񱨸������¼.����ID%Type,
                               ������_In   Ӱ�񱨸������¼.������%Type,
                               ��������_In Ӱ�񱨸������¼.��������%Type) As
  n_ҽ��ID Ӱ�񱨸������¼.ҽ��ID%Type;
  n_�ĵ����� Ӱ�񱨸��¼.�ĵ�����%Type;
  Begin

  Begin
    Select ҽ��ID,�ĵ����� Into n_ҽ��ID,n_�ĵ����� From Ӱ�񱨸��¼ Where ID = ����ID_In;
  Exception
    When Others Then
      null;
  End;
  if n_ҽ��ID is not null then
    Insert Into Ӱ�񱨸������¼
      (ID, ����ID,ҽ��ID,�ĵ�����,������,����ʱ��,��������)
    Values
      (Id_In, ����ID_In, n_ҽ��ID,n_�ĵ�����,������_In, sysdate,��������_In);
    if ��������_In=1 then
        update Ӱ�񱨸��¼ set �����ӡ=1 where ID=����ID_In;
    end if;
  end if;
  Exception
    When Others Then
      Zl_Errorcenter(SQLCode, SQLErrM);
  End;

  --11��ɾ������
  Procedure p_Ӱ�񱨸��¼_ɾ��(
    ����_Id_In Ӱ�񱨸��¼.Id%Type
	) As
  Begin    

    Delete From Ӱ�񱨸��¼ Where Ӱ�񱨸��¼.Id = Hextoraw(����_Id_In);

    Delete From ����ҽ������ Where ��鱨��ID =hextoraw(����_Id_In);

  Exception   
    When Others Then
      Zl_Errorcenter(SQLCode, SQLErrM);
  End p_Ӱ�񱨸��¼_ɾ��;


--12����ȡǩ������
Procedure p_Get_SysConfigSignature(
  Val           Out t_Refcur,
  ����ID_In		In ���ű�.ID%Type
  )Is
Begin
    --�����û�, ģ���,����
	Open  Val For 
	    select Zl_Fun_Getsignpar(7, ����ID_In) as ǩ������ from dual;
Exception
  When Others Then
  Zl_Errorcenter(Sqlcode, Sqlerrm); 
End;


--13����ȡ�˻�ǩ��ӡ��
Procedure p_Get_PersonSignImg(
  Val           Out t_Refcur,
  ID_In		In ��Ա��.ID%Type
  )Is
  v_sql Varchar2(1000);
  n_count Number(5);
Begin                
  Select Count(*) Into n_Count From user_tables Where table_name =Upper('Ӱ��ǩ��ͼƬ');
  
  If n_Count > 0 Then
     v_sql := 'Truncate Table Ӱ��ǩ��ͼƬ';
     Execute Immediate v_sql;   
     
     v_sql := 'Insert Into Ӱ��ǩ��ͼƬ Select a.id, to_lob(a.ǩ��ͼƬ) as ǩ��ͼƬ From ��Ա�� a Where a.ID=' || ID_In;
     Execute Immediate v_sql;  
  Else
     v_sql := 'Create GLOBAL TEMPORARY TABLE Ӱ��ǩ��ͼƬ ON COMMIT PRESERVE ROWS AS Select a.id, to_lob(a.ǩ��ͼƬ) as ǩ��ͼƬ From ��Ա�� a Where a.ID=' || ID_In;  
     Execute Immediate v_sql;    
  End If; 
    
  v_sql := 'Select ǩ��ͼƬ From Ӱ��ǩ��ͼƬ Where Id=:ID';
    --�����û�, ģ���,����
	Open  Val For v_sql Using ID_In;  

Exception
  When Others Then
  Zl_Errorcenter(Sqlcode, Sqlerrm); 
End;


--14����ȡǩ����֤����Ϣ
Procedure p_Get_SignCertInfo(
  Val           Out t_Refcur,
  ֤��ID_In		��Ա֤���¼.ID%Type
  )Is
Begin
	Open  Val For 
	    Select ID, CertDN,CertSN,SignCert,EncCert From ��Ա֤���¼ Where ID=֤��ID_In;
Exception
  When Others Then
		Zl_Errorcenter(Sqlcode, Sqlerrm); 
End;

--15�����±���״̬
Procedure p_Update_ReportState(
  ����Id_In  Ӱ�񱨸��¼.ID%Type,
  ����״̬_In  Ӱ�񱨸��¼.����״̬%Type,
  �����_In   Ӱ�񱨸��¼.��������%Type
  )Is
Begin
  --����״̬1-δǩ����2-����ϣ�3-����ˣ�4-������5-��ϲ��أ�6-��˲���
  --�������״̬��1-δǩ����2-�����;5-��ϲ��أ���ʱ��û������˵�
  if (����״̬_In=1) or (����״̬_In=2) or (����״̬_In=5) then 
    Update Ӱ�񱨸��¼ Set ����״̬=����״̬_In,��������=null,������ʱ��=null Where ID=����Id_In;
  elsif (����״̬_In=3) or (����״̬_In=4) then 
    Update Ӱ�񱨸��¼ Set ����״̬=����״̬_In,��������=�����_In,������ʱ��=sysdate Where ID=����Id_In;
  else
    Update Ӱ�񱨸��¼ Set ����״̬=����״̬_In Where ID=����Id_In;
  end if;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End;

--16����ȡ����״̬
Procedure p_Get_ReportState(
  Val           Out t_Refcur,
  ����Id_In	Ӱ�񱨸��¼.ID%Type
  )Is
Begin
	Open  Val For 
	    Select ����״̬ From Ӱ�񱨸��¼ Where ID=����Id_In;
Exception
  When Others Then
		Zl_Errorcenter(Sqlcode, Sqlerrm); 
End;



--17�����沵��
Procedure p_Reject_Report(
  ҽ��ID_In  Ӱ�񱨸沵��.ҽ��ID%Type,
  ����ID_In  Ӱ�񱨸沵��.��鱨��ID%Type,
  ��������_In Ӱ�񱨸沵��.��������%Type,
  ����ʱ��_In Ӱ�񱨸沵��.����ʱ��%Type,
  ������_In   Ӱ�񱨸沵��.������%Type,
  ��������_In  Ӱ�񱨸��¼.��������%Type,
  ����״̬_In Ӱ�񱨸��¼.����״̬%Type
  )Is
Begin
  Insert Into Ӱ�񱨸沵��(ID, ҽ��ID,��鱨��ID,��������,����ʱ��,������)
  Values(Ӱ�񱨸沵��_ID.NEXTVAL, ҽ��ID_IN, ����ID_In, ��������_IN, ����ʱ��_IN, ������_IN);

  Update Ӱ�񱨸��¼ Set ����״̬=����״̬_In,��������=��������_In Where ID=����ID_In;

  --Update ����ҽ������ Set ִ�й���=-1 Where ҽ��ID= ҽ��ID_IN;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End;

--17.1���������沵��
Procedure p_Reject_Cancel(
  ID_In       Ӱ�񱨸沵��.ID%Type,
  ����ID_In    Ӱ�񱨸沵��.��鱨��ID%Type,
  ����״̬_In   Ӱ�񱨸��¼.����״̬%Type
  )Is
Begin
  Update Ӱ�񱨸沵�� Set �Ƿ���=1 Where ID=ID_In;
  Update Ӱ�񱨸��¼ Set ����״̬=����״̬_In,��������='' Where ID=����ID_In;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End;

--18����ȡ���沵����Ϣ
Procedure p_Get_RejectInfo(
  Val           Out t_Refcur,
  ����ID_In  Ӱ�񱨸沵��.��鱨��ID%Type
  )Is
Begin
  Open  Val For
    Select A.ID, A.��������, A.����ʱ��, A.������, Nvl( A.�Ƿ���,0) As ����״̬, B.����״̬
    From Ӱ�񱨸沵�� A, Ӱ�񱨸��¼ B Where A.��鱨��ID=����Id_In And A.��鱨��ID = B.ID Order by ����ʱ��;
End;

--19����ȡԭ�Ͷ���
Procedure p_Get_Doc_Process(
  Val           Out t_Refcur,
  ԭ��ID_In Ӱ�񱨸涯��.ԭ��id%Type
  ) As
  Begin
    Open Val For
      Select RawtoHex(p.id) ID,
             p.���� As ��������,
			 e.���� As �¼�����,
			 e.���� As �¼�����,
			 e.Ԫ��IID As Ԫ��IID,
             p.��������,
             p.���,
             p.˵��,
             p.�ɷ��ֹ�ִ��,
             To_Clob(Nvl(p.����.GetClobVal(),'<NULL/>')) As ����, 
             RawtoHex(p.�¼�ID) �¼�ID
        From Ӱ�񱨸涯�� P, Ӱ�񱨸��¼� E
       Where p.�¼�ID = e.Id(+) And p.ԭ��ID=ԭ��ID_In
       Order By ��������, ���;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Doc_Process;

  --20��ͨ��ѧ��ɸѡ�����Ӧ�ķ�����Ϣ
  Procedure p_Get_Samplelist_By_Conditions(
    Val           Out t_Refcur,
    ԭ��id_In       Varchar2,
    ѧ��_In          Varchar2,
    Condition_In Varchar2, --����ɸѡ
    ����_In          Varchar2
  ) As
  Begin

    Open Val For
      Select /*+ rule*/ Rawtohex(a.Id) ID, a.����, a.����, a.˵��,
             Nvl2(a.˵��, a.˵�� || '����:' || a.����, '����:' || a.����) Content, a.��ǩ, a.ѧ��
      From Ӱ�񱨸淶���嵥 A
      Where a.ԭ��ID = Hextoraw(ԭ��id_In) And
            ((a.ѧ�� Is Null And a.�Ƿ�˽�� = 0) Or ѧ��_In Is Null Or a.���� = ����_In Or
            (a.ѧ�� Is Not Null And  b_PACS_RptPublic.f_If_Intersect(a.ѧ��, ѧ��_In) > 0 And a.�Ƿ�˽�� = 0)) And
            (Condition_In Is Null Or
            (a.��ǩ Is Not Null And Condition_In Is Not Null And b_PACS_RptPublic.f_If_Intersect(a.��ǩ, Condition_In) > 0))
      Order By a.���;

  End p_Get_Samplelist_By_Conditions;

  --21��ͨ������ID��ȡ��������
  Procedure p_Get_��������_By_ID(
    Val           Out t_Refcur,
    ID_IN ���ű�.ID%TYPE
    )Is
  begin
       open val for
       select ���� from ���ű� where id=ID_IN;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  end p_Get_��������_By_ID;
    

 --22����ȡ����Ԥ�����
  Procedure p_Get_AllPreOutlines(
    Val           Out t_Refcur
  )Is
  begin
       open val for
       Select Rawtohex(ID) ID, a.����, a.���� From Ӱ�񱨸�Ԥ����� a Order By a.����;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  end p_Get_AllPreOutlines;

  --23����ȡ�ĵ�����
  Procedure p_Get_reportTitle_By_ID(
    Val           Out t_Refcur,
	ID_IN  Ӱ�񱨸��¼.id%TYPE   
    )Is
  begin
       open val for
       select �ĵ����� from Ӱ�񱨸��¼ where id=ID_IN;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  end p_Get_reportTitle_By_ID;

  --24����ȡ����������
  Procedure p_Get_����������_By_ID(
    Val           Out t_Refcur,
	ID_IN  Ӱ�񱨸��¼.id%TYPE   
    )Is
  Begin
       Open Val For
         Select ������ From Ӱ�񱨸��¼ Where id =ID_IN;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  end p_Get_����������_By_ID;

 --25��ͨ��ҽ��ID��ȡ�����б�
  Procedure p_Get_Ӱ�񱨸��¼_By_ҽ��ID(
    Val           Out t_Refcur,
	ҽ��ID_IN  Ӱ�񱨸��¼.ҽ��ID%TYPE
    )Is
  Begin
       Open Val For
       Select RawToHex(ID) As REPORTID, RawToHex(ԭ��ID) As ANTETYPEID, ҽ��ID As ORDERID,�ĵ����� As REPORTNAME,
              ����ʱ�� As REPORTDATE, Decode(Nvl(����״̬,0),1,'�༭��',2,'�����',3,'�����',4,'������',5,'��ϲ���','��˲���') As REPORTSTATE,
              ������ As CreateUser,������ʱ�� As ExamineyDate,�������� As ExamineyUser,Decode(Nvl(�������,0),1,'����','') As RESULTPOSITIVE,
              Nvl(��������,0) As INNERQUALITY,' ' As REPORTQUALITY, Decode(Nvl(�����ӡ,0),0,'δ��ӡ','�Ѵ�ӡ') As ReportPrint,
              Decode(Nvl(���淢��,0),0,'δ����','�ѷ���') As REPORTRELEASE ,��¼�� as RECDOCTOR From Ӱ�񱨸��¼ Where ҽ��ID =ҽ��ID_IN
              order by REPORTDATE desc;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Ӱ�񱨸��¼_By_ҽ��ID;
  
  --26����ѯӰ�����̲���ֵ
  Procedure p_Get_Ӱ�����̲���ֵ(
    Val           Out t_Refcur,
	����ID_IN  Ӱ�����̲���.����ID%TYPE
    )Is
  Begin
       Open val For
       Select ������,����ֵ From Ӱ�����̲��� Where ����ID=����ID_IN;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Ӱ�����̲���ֵ;

  --27������ҽ��ID����ѯ��Ӧ��ԭ���б�
  Procedure p_Get_Ӱ��ԭ���б�_By_ҽ��ID(
    Val           Out t_Refcur,
    ҽ��_IN  Ӱ�����¼.ҽ��ID%TYPE   
    )Is
  Begin
       Open Val For
       Select rawtohex(c.id) As ANTETYPEID , c.���� As ANTETYPENAME,c.˵�� 
       From ����ҽ����¼ a,Ӱ�񱨸�ԭ��Ӧ�� b,Ӱ�񱨸�ԭ���嵥 c 
       Where a.id=ҽ��_IN And a.������Ŀid=b.������ĿID And b.����ԭ��ID=c.id And a.������Դ =b.Ӧ�ó���;
       
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Ӱ��ԭ���б�_By_ҽ��ID;

  --28�����ݱ���ID��ѯ��ӡ��¼
  procedure p_Get_ReportPrintLog_By_����ID
  (
       val out sys_refcursor  ,
       ����_IN  Ӱ�񱨸������¼.����ID%TYPE
  )is
  begin
       open val for
       Select  c.�ĵ����� , b.������, To_Char(b.����ʱ��, 'yyyy-MM-dd HH24:mi') ��ӡʱ��, b.������,
               To_Char(b.����ʱ��, 'yyyy-MM-dd HH24:mi') ����ʱ��, b.����˵��
               From Ӱ�񱨸������¼ B, Ӱ�񱨸��¼ C
               Where c.Id = ����_IN And b.����ID = c.Id And ��������=1 Order By b.����ʱ��;

  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  end p_Get_ReportPrintLog_By_����ID;

  --29������ҽ��ID��ѯ���淢���б�
  Procedure p_Get_ReportReleaseList(
    Val           Out t_Refcur,
    ҽ��_IN  Ӱ�񱨸��¼.ҽ��ID%TYPE   
    )Is
  Begin
       Open val For
       Select rawtohex(ID) As ����ID, �ĵ����� As ��������,���༭ʱ�� as ��������,
              decode(nvl(���淢��,0),0,'δ����','�ѷ���') As ���淢�� 
              From Ӱ�񱨸��¼ Where ����״̬ Between 2 And 4 And ҽ��ID =ҽ��_IN;
       
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_ReportReleaseList;

  --30�����ݱ���ID��ѯ���ؼ�¼����
  Procedure p_Get_RejectedCount(
    Val           Out t_Refcur,
    ����_IN  Ӱ�񱨸沵��.��鱨��ID%TYPE
    )Is
  Begin
       Open val For
       Select count(*) As �������� From Ӱ�񱨸沵�� Where ��鱨��ID=����_IN;

  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_RejectedCount;

   --31������ҽ��ID��ѯ���涯����Ҫ��һЩID��
  Procedure p_Get_DocProcess_IDs(
    Val           Out t_Refcur,
    ҽ��_IN  ����ҽ����¼.ID%TYPE
    )Is
  Begin
       open val for
       select ID as ҽ��ID,��ҳID,�Һŵ� from ����ҽ����¼ where ID=ҽ��_IN;

  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_DocProcess_IDs;

  --32������ҽ��ID�ͱ���ID��ѯ�����һЩ����
  Procedure p_Get_DocInfo(
       Val           Out t_Refcur,
       ҽ��ID_IN  Ӱ�����¼.ҽ��ID%TYPE,
       ����ID_IN  Ӱ�񱨸��¼.ID%TYPE
  )Is
  Begin
      If ����ID_IN Is Null Then 
        Open Val For 
        Select ִ�п���ID,'������' As ������ From Ӱ�����¼ Where ҽ��ID=ҽ��ID_IN;
      Else
        Open Val For
        Select ִ�п���ID,������ From Ӱ�����¼ A,Ӱ�񱨸��¼ b Where a.ҽ��ID=B.ҽ��ID and b.id=����ID_IN;
      End if;
       

  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_DocInfo;

  --33����ѯһ���������ͬԭ��ID�ı�������
   Procedure p_Get_SameAntetypeDocCounts(
       Val           Out t_Refcur,
       ҽ��ID_IN  Ӱ�񱨸��¼.ҽ��ID%TYPE,
       ԭ��ID_IN  Ӱ�񱨸��¼.ԭ��ID%TYPE
  )Is
  Begin      
        Open Val For
        Select count(id) as DocCounts From Ӱ�񱨸��¼ Where ҽ��ID=ҽ��ID_IN and ԭ��ID=ԭ��ID_IN;    
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_SameAntetypeDocCounts;

  --34����ȡ����ͼ�洢��Ϣ
  Procedure p_Get_DocImageSaveInof_By_ID(
    Val           Out t_Refcur,
	  ID_IN  Ӱ�񱨸��¼.id%TYPE
    )Is
  begin
       open val for
       select �豸��,����ʱ�� from Ӱ�񱨸��¼ where id=ID_IN;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  end p_Get_DocImageSaveInof_By_ID;

End b_PACS_RptManage;
/

CREATE OR REPLACE Package b_PACS_RptFragments Is
  Type t_Refcur Is Ref Cursor;


  --���ܣ���ȡ����Ԥ�����
  Procedure p_Get_All_Phr_Onlines(
    Val Out t_Refcur
	);
  --���ܣ���ȡ���ж������
  Procedure p_Get_All_Fragment_Class(
    Val Out t_Refcur
	);
  --���ܣ���ȡ��ǰ�û�ѧ�����ж���������ڵ�
  Procedure p_Get_All_Fragment(
    Val           Out t_Refcur,
    Subjects_In Ӱ�񱨸�Ƭ���嵥.ѧ��%Type
	) ;


  --���ܣ����ݷ���ID���Ҷ���
  Procedure p_Get_Fragment_By_Typeid(
    Val           Out t_Refcur,
    Id_In Ӱ�񱨸�Ƭ���嵥.ID%Type
	) ;


   Procedure p_Get_Label_By_Typeid(
     Val           Out t_Refcur,
	 Id_In Ӱ�񱨸�Ƭ���嵥.ID%Type
	 ) ;

  --���ܣ������������
  Procedure p_Add_Fragmenttype(
    Id_In     Ӱ�񱨸�Ƭ���嵥.ID%Type,
    Pid_In    Ӱ�񱨸�Ƭ���嵥.�ϼ�ID%Type,
    Code_In   Ӱ�񱨸�Ƭ���嵥.����%Type,
    Title_In  Ӱ�񱨸�Ƭ���嵥.����%Type,
    Note_In   Ӱ�񱨸�Ƭ���嵥.˵��%Type,
    Leaf_In   Ӱ�񱨸�Ƭ���嵥.�ڵ�����%Type,
    Author_In Ӱ�񱨸�Ƭ���嵥.����%Type
    ) ;

  --���ܣ��޸Ķ������
  Procedure p_Edit_Fragmenttype(
    Id_In     Ӱ�񱨸�Ƭ���嵥.ID%Type,
    Pid_In    Ӱ�񱨸�Ƭ���嵥.�ϼ�ID%Type,
    Code_In   Ӱ�񱨸�Ƭ���嵥.����%Type,
    Title_In  Ӱ�񱨸�Ƭ���嵥.����%Type,
    Note_In   Ӱ�񱨸�Ƭ���嵥.˵��%Type,
    Leaf_In   Ӱ�񱨸�Ƭ���嵥.�ڵ�����%Type,
    Author_In Ӱ�񱨸�Ƭ���嵥.����%Type
    ) ;

  --���ܣ�ɾ���������
   Procedure p_Del_Fragmenttype(
     Id_In Ӱ�񱨸�Ƭ���嵥.ID%Type
	 );

    --���ܣ���Ӷ���
  Procedure p_Add_Fragment(
     Id_In      Ӱ�񱨸�Ƭ���嵥.ID%Type,
    Pid_In      Ӱ�񱨸�Ƭ���嵥.�ϼ�ID%Type,
    Code_In     Ӱ�񱨸�Ƭ���嵥.����%Type,
    Title_In    Ӱ�񱨸�Ƭ���嵥.����%Type,
    Note_In     Ӱ�񱨸�Ƭ���嵥.˵��%Type,
    Leaf_In     Ӱ�񱨸�Ƭ���嵥.�ڵ�����%Type,
    Content_In  Ӱ�񱨸�Ƭ���嵥.���%Type,
    Subjects_In Ӱ�񱨸�Ƭ���嵥.ѧ��%Type,
    Label_In    Ӱ�񱨸�Ƭ���嵥.��ǩ%Type,
    Private_In  Ӱ�񱨸�Ƭ���嵥.�Ƿ�˽��%Type,
    Author_In   Ӱ�񱨸�Ƭ���嵥.����%Type
    ) ;

   --���ܣ��޸Ķ���
  Procedure p_Edit_Fragment(
    Id_In       Ӱ�񱨸�Ƭ���嵥.ID%Type,
    Pid_In      Ӱ�񱨸�Ƭ���嵥.�ϼ�ID%Type,
    Code_In     Ӱ�񱨸�Ƭ���嵥.����%Type,
    Title_In    Ӱ�񱨸�Ƭ���嵥.����%Type,
    Note_In     Ӱ�񱨸�Ƭ���嵥.˵��%Type,
    Leaf_In     Ӱ�񱨸�Ƭ���嵥.�ڵ�����%Type,
    Content_In  Ӱ�񱨸�Ƭ���嵥.���%Type,
    Subjects_In Ӱ�񱨸�Ƭ���嵥.ѧ��%Type,
    Label_In    Ӱ�񱨸�Ƭ���嵥.��ǩ%Type,
    Private_In  Ӱ�񱨸�Ƭ���嵥.�Ƿ�˽��%Type,
    Author_In   Ӱ�񱨸�Ƭ���嵥.����%Type
    );
   --���ܣ�ɾ������
  Procedure p_Del_Fragment(
    Id_In Ӱ�񱨸�Ƭ���嵥.ID%Type
	);

  procedure p_Get_All_Fragment_List(
    Val Out t_Refcur
	);

  --���ܣ��������
  Procedure p_Import_Fragment(
    Id_In       Ӱ�񱨸�Ƭ���嵥.ID%Type,
    Pid_In      Ӱ�񱨸�Ƭ���嵥.�ϼ�ID%Type,
    Code_In     Ӱ�񱨸�Ƭ���嵥.����%Type,
    Title_In    Ӱ�񱨸�Ƭ���嵥.����%Type,
    Note_In     Ӱ�񱨸�Ƭ���嵥.˵��%Type,
    Leaf_In     Ӱ�񱨸�Ƭ���嵥.�ڵ�����%Type,
    Content_In  Ӱ�񱨸�Ƭ���嵥.���%Type,
    Subjects_In Ӱ�񱨸�Ƭ���嵥.ѧ��%Type,
    Label_In    Ӱ�񱨸�Ƭ���嵥.��ǩ%Type,
    Private_In  Ӱ�񱨸�Ƭ���嵥.�Ƿ�˽��%Type,
    Author_In   Ӱ�񱨸�Ƭ���嵥.����%Type
    ) ;

procedure p_Get_Data_Last_Edit_Time(
  Val           Out t_Refcur,
  Table_Name_In varchar2
  );

   --���ܣ��ж�Ƭ�η����ܷ�ɾ��
  Procedure p_IsCanDel_FragmentType(
    Val           Out t_Refcur,
	Id_In Ӱ�񱨸�Ƭ���嵥.Id%Type
	);

  --���ܣ�����Ƭ��ID�����õ�ǰƬ�ε���Ӧ����
  Procedure p_Edit_FragmentConditionById
  (
    ID_In      In Ӱ�񱨸�Ƭ���嵥.ID%Type,
    ��Ӧ����_In In Ӱ�񱨸�Ƭ���嵥.��Ӧ����%Type
  );
  
  --���ܣ�����Ƭ�εĸ�ID����������Ŀ¼����Ŀ¼Ƭ�ε���Ӧ����
  Procedure p_Edit_FragmentConditionByPid
  (
    �ϼ�ID_In      In Ӱ�񱨸�Ƭ���嵥.ID%Type,
    ��Ӧ����_In    In Ӱ�񱨸�Ƭ���嵥.��Ӧ����%Type
  );

  --���ܣ���ȡ��ǰ����Ƭ����Ӧ����
  Procedure p_Get_FraConditionByOrderId
  (
    Val           Out t_Refcur,
	ҽ��ID_In    Ӱ�����¼.ҽ��ID%Type
  );

  --���ܣ���ȡӰ�������
  Procedure p_Get_CheckLueKind
  (
    Val           Out t_Refcur
  );
  
  --���ܣ���������ȡ���Ƽ�鲿λ
  Procedure p_Get_CheckPartList
  (
    Val           Out t_Refcur,
    Kind_In       Varchar2
  );
  
  --���ܣ���������ȡӰ������Ŀ
  Procedure p_Get_CheckRadListByKind
  (
    Val           Out t_Refcur,
    Kind_In       Varchar2
  );
  
  --���ܣ��������Ʊ����ȡӰ������Ŀ
  Procedure p_Get_CheckRadListByCode
  (
    Val           Out t_Refcur,
    Code_In       Varchar2
  );

  --�ж��Ƿ�����ͬ�Ĵ���
  Procedure p_Get_HasSameCode
  (
  Val      Out t_Refcur,
  ID_In      In Ӱ�񱨸�Ƭ���嵥.ID%Type,
  Code_In  Ӱ�񱨸�Ƭ���嵥.����%Type
  );

  --�ж��Ƿ�����ͬ������
  Procedure p_Get_HasSameName
  (
  Val      Out t_Refcur,
  ID_In    In Ӱ�񱨸�Ƭ���嵥.ID%Type,
  PID_In    In Ӱ�񱨸�Ƭ���嵥.�ϼ�ID%Type,
  Name_In  In Ӱ�񱨸�Ƭ���嵥.����%Type,
  Author_In In  Ӱ�񱨸�Ƭ���嵥.����%Type
  );

  End  b_PACS_RptFragments;
/
CREATE OR REPLACE Package Body b_PACS_RptFragments Is

  ------------------------------------------------------------------------
  --Ƭ��ģ��
  ------------------------------------------------------------------------

  --���ܣ���ȡ����Ԥ�����
  Procedure p_Get_All_Phr_Onlines(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select Rawtohex(ID) ID, ����, ���� From Ӱ�񱨸�Ԥ����� Order By ����;
  End p_Get_All_Phr_Onlines;

  --���ܣ���ȡ���ж������
  Procedure p_Get_All_Fragment_Class(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select Rawtohex(a.Id) As ID, Rawtohex(a.�ϼ�id) As �ϼ�id, a.����, a.����, a.˵��, a.�ڵ�����
      From Ӱ�񱨸�Ƭ���嵥 A
      Where a.�ڵ����� = 0
      Start With �ϼ�id Is Null
      Connect By Prior ID = �ϼ�id
      Order By ����;
  End p_Get_All_Fragment_Class;

  --���ܣ���ȡ��ǰ�û�ѧ�����ж���������ڵ�
  Procedure p_Get_All_Fragment(
    Val           Out t_Refcur,
    Subjects_In Ӱ�񱨸�Ƭ���嵥.ѧ��%Type
	) As
  Begin
    If Subjects_In <> '' Then
      Open Val For
        Select Rawtohex(a.Id) As ID, Rawtohex(a.�ϼ�id) As �ϼ�id, a.����, a.����, a.˵��, a.�ڵ�����, Nvl(a.���.GetClobVal(), '<NULL/>') As ���, 
			a.ѧ��, a.��ǩ, a.�Ƿ�˽��, a.����, Nvl(a.��Ӧ����.GetClobVal(), '<NULL/>') As ��Ӧ����,a.���༭ʱ��, a.�ڵ����� As Image
        From Ӱ�񱨸�Ƭ���嵥 A
        Where (a.ѧ�� In (Select /*+rule*/
                         Column_Value As Lable
                        From Table(b_Pacs_Common.f_Str2list(Subjects_In, ','))
                        Intersect
                        Select /*+rule*/
                         Column_Value As Lable
                        From Table(b_Pacs_Common.f_Str2list(a.ѧ��, ','))) And a.�ڵ����� <> 0) Or a.�ڵ����� = 0 Or a.ѧ�� Is Null
        Order By ����, �ϼ�id;
    Else
      Open Val For
        Select Rawtohex(a.Id) As ID, Rawtohex(a.�ϼ�id) As �ϼ�id, a.����, a.����, a.˵��, a.�ڵ�����, Nvl(a.���.GetClobVal(), '<NULL/>') As ���, 
			a.ѧ��, a.��ǩ, a.�Ƿ�˽��, a.����, Nvl(a.��Ӧ����.GetClobVal(), '<NULL/>') As ��Ӧ����,a.���༭ʱ��, a.�ڵ����� As Image
        From Ӱ�񱨸�Ƭ���嵥 A
        Order By �ϼ�id, �ڵ�����, ����, ����;
    End If;
  End p_Get_All_Fragment;

  --���ܣ����ݷ���ID���Ҷ���
  Procedure p_Get_Fragment_By_Typeid(
    Val           Out t_Refcur,
    Id_In Ӱ�񱨸�Ƭ���嵥.Id%Type
  ) As
  Begin
    Open Val For
      Select Rawtohex(a.Id) As ID, a.�ϼ�ID,a.����, a.����, a.˵��, a.�ڵ�����, Nvl(a.���.GetClobVal(), '<NULL/>') As ���, 
				a.ѧ��, a.��ǩ, a.�Ƿ�˽��, a.����, Nvl(a.��Ӧ����.GetClobVal(), '<NULL/>') As ��Ӧ����, a.���༭ʱ��,a.�ڵ����� As Image
      From Ӱ�񱨸�Ƭ���嵥 A
      Where a.�ϼ�id = Hextoraw(Id_In) And a.�ڵ����� <> 0;
  End p_Get_Fragment_By_Typeid;

  --���ܣ�����ĳ���������ж����ǩ
  Procedure p_Get_Label_By_Typeid(
    Val           Out t_Refcur,
    Id_In Ӱ�񱨸�Ƭ���嵥.Id%Type
    ) As
  Begin
    Open Val For
      Select Distinct ��ǩ From Ӱ�񱨸�Ƭ���嵥 Where �ϼ�id = Hextoraw(Id_In) And ��ǩ Is Not Null;
  End p_Get_Label_By_Typeid;

  --���ܣ������������
  Procedure p_Add_Fragmenttype(
    Id_In     Ӱ�񱨸�Ƭ���嵥.Id%Type,
    Pid_In    Ӱ�񱨸�Ƭ���嵥.�ϼ�id%Type,
    Code_In   Ӱ�񱨸�Ƭ���嵥.����%Type,
    Title_In  Ӱ�񱨸�Ƭ���嵥.����%Type,
    Note_In   Ӱ�񱨸�Ƭ���嵥.˵��%Type,
    Leaf_In   Ӱ�񱨸�Ƭ���嵥.�ڵ�����%Type,
    Author_In Ӱ�񱨸�Ƭ���嵥.����%Type
    ) As
    n_Num     Number;
    v_Err_Msg Varchar2(100);
    Err_Item Exception;
  Begin
    Select Count(ID)
    Into n_Num
    From Ӱ�񱨸�Ƭ���嵥
    Where ���� = Code_In Or ���� = Title_In And �ڵ����� = 0 And �ϼ�id = Hextoraw(Pid_In);

    If n_Num > 0 Then
      v_Err_Msg := '[ZLSOFT]�������ƻ�����Ѿ����ڣ�[ZLSOFT]';
      Raise Err_Item;
    Else
      Insert Into Ӱ�񱨸�Ƭ���嵥
        (ID, �ϼ�id, ����, ����, ˵��, �ڵ�����, ����, ���༭ʱ��)
      Values
        (Hextoraw(Id_In), Hextoraw(Pid_In), Code_In, Title_In, Note_In, Leaf_In, Author_In, Sysdate);
    End If;

  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Add_Fragmenttype;

  --���ܣ��޸Ķ������
  Procedure p_Edit_Fragmenttype(
    Id_In     Ӱ�񱨸�Ƭ���嵥.Id%Type,
    Pid_In    Ӱ�񱨸�Ƭ���嵥.�ϼ�id%Type,
    Code_In   Ӱ�񱨸�Ƭ���嵥.����%Type,
    Title_In  Ӱ�񱨸�Ƭ���嵥.����%Type,
    Note_In   Ӱ�񱨸�Ƭ���嵥.˵��%Type,
    Leaf_In   Ӱ�񱨸�Ƭ���嵥.�ڵ�����%Type,
    Author_In Ӱ�񱨸�Ƭ���嵥.����%Type
    ) As
    n_Num     Number;
    v_Err_Msg Varchar2(100);
    Err_Item Exception;
  Begin
    Select Count(ID)
    Into n_Num
    From Ӱ�񱨸�Ƭ���嵥
    Where (���� = Code_In Or ���� = Title_In) And �ڵ����� = 0 And �ϼ�id = Hextoraw(Pid_In) And ID <> Hextoraw(Id_In);

	If n_Num > 0 Then
      v_Err_Msg := '[ZLSOFT]�������ƻ�����Ѿ����ڣ�[ZLSOFT]';
      Raise Err_Item;
    Else
      Update Ӱ�񱨸�Ƭ���嵥
      Set �ϼ�id = Hextoraw(Pid_In), ���� = Code_In, ���� = Title_In, ˵�� = Note_In, �ڵ����� = Leaf_In, ���� = Author_In,
          ���༭ʱ�� = Sysdate
      Where ID = Hextoraw(Id_In);
    End If;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Edit_Fragmenttype;

  --���ܣ�ɾ���������
  Procedure p_Del_Fragmenttype(
    Id_In Ӱ�񱨸�Ƭ���嵥.Id%Type
	) As
    n_Num     Number;
    v_Err_Msg Varchar2(100);
    Err_Item Exception;
  Begin
    Select Count(ID)
    Into n_Num
    From Ӱ�񱨸�Ƭ���嵥
    Where �ڵ����� <> 0 And
          ID In (Select ID From Ӱ�񱨸�Ƭ���嵥 Connect By Prior ID = �ϼ�id Start With ID = Hextoraw(Id_In));

	If n_Num > 0 Then
      v_Err_Msg := '[ZLSOFT]�÷����´��ڶ���ݲ���ɾ����[ZLSOFT]';
      Raise Err_Item;
    Else
      Delete Ӱ�񱨸�Ƭ���嵥 Where ID = Hextoraw(Id_In);
    End If;

  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Del_Fragmenttype;

  --���ܣ���Ӷ���
  Procedure p_Add_Fragment(
    Id_In       Ӱ�񱨸�Ƭ���嵥.Id%Type,
    Pid_In      Ӱ�񱨸�Ƭ���嵥.�ϼ�id%Type,
    Code_In     Ӱ�񱨸�Ƭ���嵥.����%Type,
    Title_In    Ӱ�񱨸�Ƭ���嵥.����%Type,
    Note_In     Ӱ�񱨸�Ƭ���嵥.˵��%Type,
    Leaf_In     Ӱ�񱨸�Ƭ���嵥.�ڵ�����%Type,
    Content_In  Ӱ�񱨸�Ƭ���嵥.���%Type,
    Subjects_In Ӱ�񱨸�Ƭ���嵥.ѧ��%Type,
    Label_In    Ӱ�񱨸�Ƭ���嵥.��ǩ%Type,
    Private_In  Ӱ�񱨸�Ƭ���嵥.�Ƿ�˽��%Type,
    Author_In   Ӱ�񱨸�Ƭ���嵥.����%Type
    ) As
  Begin

      Insert Into Ӱ�񱨸�Ƭ���嵥
        (ID, �ϼ�id, ����, ����, ˵��, �ڵ�����, ���, ѧ��, ��ǩ, �Ƿ�˽��, ����, ���༭ʱ��)
      Values
        (Hextoraw(Id_In), Hextoraw(Pid_In), Code_In, Title_In, Note_In, Leaf_In, Content_In, Subjects_In, Label_In,
         Private_In, Author_In, Sysdate);

  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Add_Fragment;

  --���ܣ��޸Ķ���
  Procedure p_Edit_Fragment(
    Id_In       Ӱ�񱨸�Ƭ���嵥.Id%Type,
    Pid_In      Ӱ�񱨸�Ƭ���嵥.�ϼ�id%Type,
    Code_In     Ӱ�񱨸�Ƭ���嵥.����%Type,
    Title_In    Ӱ�񱨸�Ƭ���嵥.����%Type,
    Note_In     Ӱ�񱨸�Ƭ���嵥.˵��%Type,
    Leaf_In     Ӱ�񱨸�Ƭ���嵥.�ڵ�����%Type,
    Content_In  Ӱ�񱨸�Ƭ���嵥.���%Type,
    Subjects_In Ӱ�񱨸�Ƭ���嵥.ѧ��%Type,
    Label_In    Ӱ�񱨸�Ƭ���嵥.��ǩ%Type,
    Private_In  Ӱ�񱨸�Ƭ���嵥.�Ƿ�˽��%Type,
    Author_In   Ӱ�񱨸�Ƭ���嵥.����%Type
    ) As
    n_Num     Number;
    v_Err_Msg Varchar2(100);
    Err_Item Exception;
  Begin
    Select Count(ID)
    Into n_Num
    From Ӱ�񱨸�Ƭ���嵥
    Where (���� = Code_In Or ���� = Title_In) And �ڵ����� <> 0 And �ϼ�id = Hextoraw(Pid_In) And ID <> Hextoraw(Id_In);

	If n_Num > 0 Then
      v_Err_Msg := '[ZLSOFT]��������ƻ�����Ѿ����ڣ�[ZLSOFT]';
      Raise Err_Item;
    Else
      Update Ӱ�񱨸�Ƭ���嵥
      Set �ϼ�id = Hextoraw(Pid_In), ���� = Code_In, ���� = Title_In, ˵�� = Note_In, �ڵ����� = Leaf_In, ��� = Content_In,
          ѧ�� = Subjects_In, ��ǩ = Label_In, �Ƿ�˽�� = Private_In, ���� = Author_In, ���༭ʱ�� = Sysdate
      Where ID = Hextoraw(Id_In);
    End If;

  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Edit_Fragment;

  --
  Procedure p_Get_All_Fragment_List(Val Out t_Refcur) As
  Begin
    Open Val For
      Select Rawtohex(t.Id) As ID, Rawtohex(t.�ϼ�id) As �ϼ�id, t.����, t.����, t.˵��, t.�ڵ�����, Nvl(t.���.GetClobVal(), '<NULL/>') As ���, t.ѧ��, t.��ǩ, t.�Ƿ�˽��, t.����,
             t.���༭ʱ��
      From Ӱ�񱨸�Ƭ���嵥 T;
  End p_Get_All_Fragment_List;

  --���ܣ�ɾ������
  Procedure p_Del_Fragment(
    Id_In Ӱ�񱨸�Ƭ���嵥.Id%Type
	) As
  Begin
    Delete Ӱ�񱨸�Ƭ���嵥 Where ID = Hextoraw(Id_In);
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Del_Fragment;

  --���ܣ��������
  Procedure p_Import_Fragment(
    Id_In       Ӱ�񱨸�Ƭ���嵥.Id%Type,
    Pid_In      Ӱ�񱨸�Ƭ���嵥.�ϼ�id%Type,
    Code_In     Ӱ�񱨸�Ƭ���嵥.����%Type,
    Title_In    Ӱ�񱨸�Ƭ���嵥.����%Type,
    Note_In     Ӱ�񱨸�Ƭ���嵥.˵��%Type,
    Leaf_In     Ӱ�񱨸�Ƭ���嵥.�ڵ�����%Type,
    Content_In  Ӱ�񱨸�Ƭ���嵥.���%Type,
    Subjects_In Ӱ�񱨸�Ƭ���嵥.ѧ��%Type,
    Label_In    Ӱ�񱨸�Ƭ���嵥.��ǩ%Type,
    Private_In  Ӱ�񱨸�Ƭ���嵥.�Ƿ�˽��%Type,
    Author_In   Ӱ�񱨸�Ƭ���嵥.����%Type
    ) As
    v_Num Number(2);
  Begin
    Select Count(ID)
    Into v_Num
    From Ӱ�񱨸�Ƭ���嵥
    Where ((���� = Code_In Or ���� = Title_In) And �ϼ�id = Hextoraw(Pid_In)) Or
          (�ϼ�id Is Null And (���� = Code_In Or ���� = Title_In));

    If v_Num > 0 Then
      Update Ӱ�񱨸�Ƭ���嵥
      Set ��� = Content_In, ���༭ʱ�� = Sysdate, �Ƿ�˽�� = 0
      Where ((���� = Code_In Or ���� = Title_In) And �ϼ�id = Hextoraw(Pid_In)) Or
            (�ϼ�id Is Null And (���� = Code_In Or ���� = Title_In));
    Else
      Insert Into Ӱ�񱨸�Ƭ���嵥
        (ID, �ϼ�id, ����, ����, ˵��, �ڵ�����, ���, ѧ��, ��ǩ, �Ƿ�˽��, ����, ���༭ʱ��)
      Values
        (Hextoraw(Id_In), Hextoraw(Pid_In), Code_In, Title_In, Note_In, Leaf_In, Content_In, Subjects_In, Label_In,
         Private_In, Author_In, Sysdate);
    End If;

  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Import_Fragment;

  --
  Procedure p_Get_Data_Last_Edit_Time(
    Val           Out t_Refcur,
    Table_Name_In Varchar2
    ) As
    v_Sql Varchar2(4000);
  Begin
    v_Sql := 'select max(���༭ʱ��) maxvalue from ' || Table_Name_In;
    Open Val For v_Sql;
  End p_Get_Data_Last_Edit_Time;
  
   --���ܣ��ж�Ƭ�η����ܷ�ɾ��
  Procedure p_IsCanDel_FragmentType(
    Val           Out t_Refcur,
	Id_In Ӱ�񱨸�Ƭ���嵥.Id%Type
	) As
  Begin
    Open Val For
      Select Count(t.id) Count
        From Ӱ�񱨸�Ƭ���嵥 t
       Where �ϼ�id = Hextoraw(Id_In);
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_IsCanDel_FragmentType;
  
  --���ܣ�����Ƭ��ID�����õ�ǰƬ�ε���Ӧ����
  Procedure p_Edit_FragmentConditionById
  (
    ID_In      In Ӱ�񱨸�Ƭ���嵥.ID%Type,
    ��Ӧ����_In In Ӱ�񱨸�Ƭ���嵥.��Ӧ����%Type
  )As
  Begin
    Update Ӱ�񱨸�Ƭ���嵥 Set ��Ӧ���� = ��Ӧ����_In Where ID = Hextoraw(ID_In) And �ڵ����� != 0;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Edit_FragmentConditionById;
  
  --���ܣ�����Ƭ�εĸ�ID����������Ŀ¼����Ŀ¼Ƭ�ε���Ӧ����
  Procedure p_Edit_FragmentConditionByPid
  (
    �ϼ�ID_In      In Ӱ�񱨸�Ƭ���嵥.ID%Type,
    ��Ӧ����_In In Ӱ�񱨸�Ƭ���嵥.��Ӧ����%Type
  )As
  Begin
    Update Ӱ�񱨸�Ƭ���嵥 Set ��Ӧ���� = ��Ӧ����_In Where �ϼ�ID = Hextoraw(�ϼ�ID_In) And �ڵ����� != 0;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Edit_FragmentConditionByPid;

  --���ܣ���ȡ��ǰ����Ƭ����Ӧ����
  Procedure p_Get_FraConditionByOrderId(
    Val           Out t_Refcur,
	  ҽ��ID_In    Ӱ�����¼.ҽ��ID%Type
	) As
  Begin
    Open Val For
	  Select a.id, a.�Ա�,c.Ӱ�����, d.����||' - '||d.���� ������, c.Ӱ�����||' - '||e.����||' - '||e.���� �����Ŀ, A.ҽ������
      From ����ҽ����¼ a, ����ҽ������ b, Ӱ�����¼ c, Ӱ������� d, ������ĿĿ¼ e
      Where a.id = b.ҽ��id and b.ҽ��id=c.ҽ��id and c.Ӱ����� = d.���� and a.������Ŀid = e.id and a.id = ҽ��ID_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Get_FraConditionByOrderId;

  --���ܣ���ȡӰ�������
  Procedure p_Get_CheckLueKind
  (
    Val           Out t_Refcur
  ) As
  Begin
    Open Val For
      Select ����||' - '||���� ������ From Ӱ�������;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Get_CheckLueKind;
  
  --���ܣ���������ȡ���Ƽ�鲿λ
  Procedure p_Get_CheckPartList
  (
    Val           Out t_Refcur,
    Kind_In       Varchar2
  ) As
  Begin
    Open Val For
      Select Distinct ����||���� IID, '' �ϼ�ID, ����||' - '||���� ���Ʋ�λ From ���Ƽ�鲿λ a,
      Table(Cast(f_Str2list(''||Kind_In||'') As zlTools.t_Strlist)) b Where a.���� = b.Column_Value
      Union Select ����||����||���� IID, ����||���� �ϼ�ID, ����||' - '||���� ���Ʋ�λ From ���Ƽ�鲿λ c,
      Table(Cast(f_Str2list(''||Kind_In||'') As zlTools.t_Strlist)) d Where c.���� = d.Column_Value;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Get_CheckPartList;
  
  --���ܣ���������ȡӰ������Ŀ
  Procedure p_Get_CheckRadListByKind
  (
    Val           Out t_Refcur,
    Kind_In       Varchar2
  ) As
  Begin
    Open Val For
      Select I.����, r.Ӱ�����||' - '||I.����||' - '||I.���� �����Ŀ
      From ������ĿĿ¼ I, Ӱ������Ŀ R, Table(Cast(f_Str2list(''||Kind_In||'') As zlTools.t_Strlist)) M
      Where I.ID = R.������Ŀid And R.Ӱ�����=M.Column_Value;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Get_CheckRadListByKind;
  
  --���ܣ��������Ʊ����ȡӰ������Ŀ
  Procedure p_Get_CheckRadListByCode
  (
    Val           Out t_Refcur,
    Code_In       Varchar2
  ) As
  Begin
    Open Val For
      Select I.����, r.Ӱ�����||' - '||I.����||' - '||I.���� �����Ŀ
      From ������ĿĿ¼ I, Ӱ������Ŀ R, Table(Cast(f_Str2list(''||Code_In||'') As zlTools.t_Strlist)) M
      Where I.ID = R.������Ŀid And I.����=M.Column_Value;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Get_CheckRadListByCode;

  --�ж��Ƿ�����ͬ�Ĵ���
  Procedure p_Get_HasSameCode
  (
  Val      Out t_Refcur,
  ID_In      In Ӱ�񱨸�Ƭ���嵥.ID%Type,
  Code_In  Ӱ�񱨸�Ƭ���嵥.����%Type
  ) As
  Begin
  Open Val For
    Select Count(1) From Ӱ�񱨸�Ƭ���嵥 Where ID<>ID_In And ����=Code_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Get_HasSameCode;

  --�ж��Ƿ�����ͬ������
  Procedure p_Get_HasSameName
  (
  Val      Out t_Refcur,
  ID_In    In Ӱ�񱨸�Ƭ���嵥.ID%Type,
  PID_In    In Ӱ�񱨸�Ƭ���嵥.�ϼ�ID%Type,
  Name_In  In Ӱ�񱨸�Ƭ���嵥.����%Type,
  Author_In In  Ӱ�񱨸�Ƭ���嵥.����%Type
  ) As
  Begin
  Open Val For
    Select Count(1) From Ӱ�񱨸�Ƭ���嵥 Where �ϼ�ID=PID_In And ����=Author_In And ID<>ID_In And ����=Name_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Get_HasSameName;

End  b_PACS_RptFragments;
/

--89419:�ŵ���,2015-01-05,��Ժ���˲������÷�
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
      Select Max(NO) Into v_No From ��Һ��ҩ���� Where ��ҩid = v_Tansid;
      if v_No is not null then
        select count(no) into n_row from סԺ���ü�¼ where NO=v_No and ���=1 and ��¼״̬=1;
        if n_row<>0 then
           Zl_סԺ���ʼ�¼_Delete(v_No, 1, v_Usercode, Zl_Username);
        end if;
      end if;
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

--89419:�ŵ���,2015-01-05,��Ժ���˲������÷�
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

--91083:�ŵ���,2015-12-02,��Һ�������Ĳ�����ҩƷ����ʹ��
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
  Err_Item Exception;

  Cursor c_ҽ����¼ Is
    Select /*+rule */
    Distinct e.ҽ��id As ���id, e.���ͺ�, b.Ƶ�ʼ��, b.�����λ, b.ִ��ʱ�䷽��, a.����, a.�Ա�, a.����, a.סԺ��, a.��ǰ���� As ����, a.��ǰ����id As ���˲���id,
             a.��ǰ����id As ���˿���id, e.�״�ʱ��, e.ĩ��ʱ��, b.��ʼִ��ʱ��, Nvl(e.��������, 0) As ����, e.����ʱ��, Decode(b.ҽ����Ч, 0, 1, 2) As ҽ������,
             b.������Ŀid As ��ҩ;��, b.����id, Nvl(c.ִ�б��, 0) As �Ƿ�tpn
    From ����ҽ������ E, ����ҽ����¼ B, ������Ϣ A, ������ĿĿ¼ C, Table(f_Num2list(ҽ��id_In)) D
    Where e.ҽ��id = b.Id And b.����id = a.����id And c.��� = 'E' And c.�������� = '2' And c.ִ�з��� = 1 And b.������Ŀid = c.Id And
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
    Select Distinct c.Id As �շ�id, c.���, c.ʵ������ As ����, Nvl(e.�Ƿ�������, 0) As �Ƿ�������, c.����, c.No
    From ����ҽ����¼ A, ����ҽ������ B, ҩƷ�շ���¼ C, סԺ���ü�¼ D, ��ҺҩƷ���� E
    Where a.Id = b.ҽ��id And c.����id = d.Id And a.Id = d.ҽ����� And b.No = c.No And c.ҩƷid = e.ҩƷid(+) And
          b.ִ�в���id + 0 = ����id_In And c.���� = 9 And c.������� Is Null And a.���id = v_���id And b.���ͺ� = ���ͺ�_In And c.��� < 1000
    Order By c.No, c.���;

  v_ҽ����¼     c_ҽ����¼%RowType;
  v_�շ���¼     c_�շ���¼%RowType;
  v_����ҽ����¼ c_����ҽ����¼%RowType;
  Function Zl_Getpivaworkbatch
  (
    ִ��ʱ��_In   In Date,
    ��������id_In In ��Һ��ҩ��¼.����id%Type
  ) Return Number As
    v_Exetime   Date;
    v_Starttime Date;
    v_Endtime   Date;
    v_Maxbatch  Number(2);
    v_Batch     Number;
    Cursor c_��ҩ���� Is
      Select ����, ��ҩʱ��, ��ҩʱ��, ���
      From ��ҩ��������
      Where ���� = 1 And ��������id = ��������id_In
      Order By ����;

    v_��ҩ���� c_��ҩ����%RowType;
  Begin
    v_Exetime := To_Date(Substr(To_Char(ִ��ʱ��_In, 'yyyy-mm-dd hh24:mi'), 12), 'hh24:mi');

    Select Max(Nvl(����, 0)) + 1 Into v_Maxbatch From ��ҩ�������� Where ���� = 1 And ��������id = ��������id_In;

    For v_��ҩ���� In c_��ҩ���� Loop
      v_Batch     := 0;
      v_Starttime := To_Date(Substr(v_��ҩ����.��ҩʱ��, 1, Instr(v_��ҩ����.��ҩʱ��, '-') - 1), 'hh24:mi');
      v_Endtime   := To_Date(Substr(v_��ҩ����.��ҩʱ��, Instr(v_��ҩ����.��ҩʱ��, '-') + 1), 'hh24:mi');

      If v_Exetime >= v_Starttime And v_Exetime <= v_Endtime Then
        v_Batch := v_��ҩ����.����;
        n_���  := v_��ҩ����.���;
        Exit When v_Batch > 0;
      End If;
    End Loop;

    If v_Batch = 0 Then
      v_Batch := v_Maxbatch;
    End If;
    Return(v_Batch);
  End;
Begin
  n_Count          := 0;
  v_ҽ������       := Zl_To_Number(Nvl(zl_GetSysParameter('ҽ������', 1345), 1));
  v_��Һ����       := Zl_To_Number(Nvl(zl_GetSysParameter('ͬ������Һ����', 1345), 0));
  v_����Һ����     := Nvl(zl_GetSysParameter('����ҺҩƷ����', 1345), '');
  v_����Һ��ҩ;�� := Nvl(zl_GetSysParameter('��Һ��ҩ;��', 1345), '');
  v_��Դ����       := Nvl(zl_GetSysParameter('��Դ����', 1345), '');
  v_�����ϴ�����   := Zl_To_Number(Nvl(zl_GetSysParameter('�����ϴ�����', 1345), 0));

  v_ҽ��ids  := ҽ��id_In;
  v_��ǰ���� := '';
  v_New���id:=0;

  Select Trunc(Sysdate) Into v_Currdate From Dual;

  Select Max(Nvl(����, 0)) + 1 Into v_Maxbatch From ��ҩ�������� Where ���� = 1 And ��������id = ����id_In;

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

  For v_ҽ����¼ In c_ҽ����¼ Loop
    v_Continue := 1;

    Select Count(1) into v_Continue
    From ����ҽ����¼ A, ��Һ������ҩƷ B,סԺ���ü�¼ C
    Where c.�շ�ϸĿid = b.ҩƷid and A.id=C.ҽ����� And a.���id = v_ҽ����¼.���id and C.��¼״̬=1;
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
      If Instr(',' || v_��Դ���� || ',', ',' || v_ҽ����¼.���˲���id || ',') = 0 Then
        v_Continue := 0;
      End If;
    End If;

    If v_ҽ����¼.�Ƿ�tpn = 2 Then
      v_Continue := 1;
    end if;

    If v_Continue = 1 Then
      v_Old���id := v_New���id;
      v_���id    := v_ҽ����¼.���id;
      v_New���id := v_���id;
      v_���ͺ�    := v_ҽ����¼.���ͺ�;
      v_���      := 0;

      If v_Continue = 1 Then
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

              For v_��Һ��¼ In (Select ID, ִ��ʱ��
                             From ��Һ��ҩ��¼
                             Where ҽ��id In
                                   (Select ID
                                    From ����ҽ����¼
                                    Where ����id =
                                          (Select ����id From ����ҽ����¼ Where ID = v_ҽ����¼.���id And Rownum < 2)) And
                                   ִ��ʱ�� Between Trunc(v_ִ��ʱ��) And Trunc(v_ִ��ʱ��+1) - 1 / 24 / 60 / 60 And ����״̬ < 2) Loop
                v_���� := Zl_Getpivaworkbatch(v_��Һ��¼.ִ��ʱ��, ����id_In);
                Update ��Һ��ҩ��¼ Set ��ҩ���� = v_���� Where ID = v_��Һ��¼.Id;
                v_���� := 0;
              End Loop;
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

          If v_�����ϴ����� = 1 Or b_Change = True Then
            --ȡ�ϴε�����
            Begin
              Select Distinct ��ҩ����
              Into v_����
              From ��Һ��ҩ��¼ A
              Where ҽ��id = v_ҽ����¼.���id And
                    ���ͺ� = (Select Distinct Max(���ͺ�)
                           From ��Һ��ҩ��¼
                           Where ҽ��id = v_ҽ����¼.���id And ���ͺ� <> v_ҽ����¼.���ͺ� And ִ��ʱ��<v_ִ��ʱ�� and To_Char(ִ��ʱ��, 'hh24:mi:ss') = To_Char(v_ִ��ʱ��, 'hh24:mi:ss')) And
                    To_Char(a.ִ��ʱ��, 'hh24:mi:ss') = To_Char(v_ִ��ʱ��, 'hh24:mi:ss');
            Exception
              When Others Then
                v_���� := 0;
            End;
          End If;

          If v_���� = 0 Then
            v_���� := Zl_Getpivaworkbatch(v_ִ��ʱ��, ����id_In);

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

          if v_Old���id<>v_ҽ����¼.���id then
            Select Count(ҽ��id)
            Into n_���ʹ���
            From ��Һ��ҩ��¼
            Where ҽ��id = v_ҽ����¼.���id
            Order By ִ��ʱ��;
          else
            n_���ʹ���:=n_���ʹ���+1;
          end if;

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
            Select Nvl(Max(���), 0)
            Into n_���
            From ��ҩ��������
            Where ���� = 1 And ��������id = ����id_In And ���� = v_����;
          End If;

          If Trunc(v_ִ��ʱ��) <= v_Currdate Or n_��� <> 0 Then
            n_�Ƿ���     := 1;
            d_�ֹ����ʱ�� := Sysdate;
          Else
            n_�Ƿ���     := 0;
            d_�ֹ����ʱ�� := Null;
          End If;

           --�����TPN�����ָ����Ҫ��������ã��򲻹�����������ζ�����Ϊ���������
          If v_ҽ����¼.�Ƿ�tpn = 2 Then
            n_�Ƿ���     := 0;
            d_�ֹ����ʱ�� := Null;
          End If;

          --������ҩ��¼
          Insert Into ��Һ��ҩ��¼
            (ID, ����id, ���, ����, �Ա�, ����, סԺ��, ����, ���˲���id, ���˿���id, ִ��ʱ��, ҽ��id, ���ͺ�, ��ҩ����, ƿǩ��, �Ƿ��������, �Ƿ���, ���ʱ��, ����״̬,
             ������Ա, ����ʱ��)
          Values
            (v_��ҩid, ����id_In, v_���, v_ҽ����¼.����, v_ҽ����¼.�Ա�, v_ҽ����¼.����, v_ҽ����¼.סԺ��, v_ҽ����¼.����, v_ҽ����¼.���˲���id,
             v_ҽ����¼.���˿���id, v_ִ��ʱ��, v_ҽ����¼.���id, v_ҽ����¼.���ͺ�, Decode(v_����, 0, Null, v_����), v_Maxno, n_��������, n_�Ƿ���,
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
          End Loop;

        End Loop;

        For v_�շ���¼ In c_�շ���¼ Loop
          n_���� := v_�շ���¼.����;

          v_No := v_�շ���¼.No;
          Delete From ҩƷ�շ���¼ Where ID = v_�շ���¼.�շ�id;
        End Loop;

        --������ڡ��������á����Ե�ҩƷ��Ҳ����Ϊ���
        If v_Nodosage = 1 Then
          Update ��Һ��ҩ��¼ Set �Ƿ��� = 1 Where ҽ��id = v_ҽ����¼.���id And ���ͺ� = v_ҽ����¼.���ͺ�;
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

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]����' || v_��ǰ���� || '����Һ���������б���������Һ��������ʧ�ܣ�[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Һ��ҩ��¼_�˲�;
/

--90943:������,2015-12-02,֧����ԤԼ����������
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
  --        <BRID>����ID</BRID>
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
  --    �D�D�������д�������˵����ȷִ�� 
  --    <ERROR> 
  --    <CZSJ>����ʱ��</CZSJ>          //HIS�ĵǼ�ʱ��
  --      <MSG>������Ϣ</MSG> 
  --    </ERROR> 
  --  </OUT> 
  -------------------------------------------------------------------------------------------------- 
  v_Nos      Varchar2(4000);
  n_�շ��ܶ� ������ü�¼.ʵ�ս��%Type;

  n_�����id ҽ�ƿ����.Id%Type;
  v_���㷽ʽ Varchar2(2000);
  n_����id   ������ü�¼.����id%Type;
  v_����     ������ü�¼.����%Type;
  v_�Ա�     ������ü�¼.�Ա�%Type;
  v_����     ������ü�¼.����%Type;

  v_ҽ�Ƹ��ʽ���� ҽ�Ƹ��ʽ.����%Type;
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

  v_Temp    Varchar2(32767); --��ʱXML 
  x_Templet Xmltype; --ģ��XML 

  v_Err_Msg Varchar2(200);
  Err_Item Exception;
  n_Count    Number(18);
  v_��ҩ���� Varchar2(4000);
  n_����   ����Ԥ����¼.��Ԥ��%Type;
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
         To_Number(Extractvalue(Value(A), 'IN/SFGH'))
  Into v_Nos, n_����id, n_�շ��ܶ�, n_����, n_�Ƿ�Һ�
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

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
    Select b.����, a.����, a.�Ա�, a.����
    Into v_ҽ�Ƹ��ʽ����, v_����, v_�Ա�, v_����
    From ������Ϣ A, ҽ�Ƹ��ʽ B
    Where a.ҽ�Ƹ��ʽ = b.����(+) And a.����id = n_����id;
  Exception
    When Others Then
      v_Err_Msg := 'ָ���Ľɷѵ����в�����Чʶ����,������ɷ�!';
  End;
  If Not v_Err_Msg Is Null Then
    Raise Err_Item;
  End If;
  Select ���˽��ʼ�¼_Id.Nextval, Sysdate Into n_����id, d_�շ�ʱ�� From Dual;

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
      n_���� := Nvl(n_�շ��ܶ�, 0) - Nvl(n_���ʽ��, 0);
      If Abs(n_����) > 1.00 Then
        v_Err_Msg := '���ݽɿ�����ʵ�ʽ�������̫��!';
        Raise Err_Item;
      End If;
    End If;
    If Nvl(n_���ʽ��, 0) + Nvl(n_����, 0) <> Nvl(n_�շ��ܶ�, 0) Then
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
    For c_���� In (Select 1 As ˳���, b.No, b.�վݷ�Ŀ, b.����id, b.ִ�в���id, b.���˿���ID, b.������, b.�շ����, b.������Ŀid, b.���ӱ�־,
                        To_Char(b.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��, b.�۸񸸺�, b.��������, b.���, b.�շ�ϸĿid, b.���㵥λ,
                        Max(m.����) As ����, Max(m.���) As ���, Sum(b.��׼����) As ����, Avg(Nvl(b.����, 1) * b.����) As ����,
                        Sum(b.Ӧ�ս��) As Ӧ�ս��, Sum(b.ʵ�ս��) As ʵ�ս��, Max(j.����) As ��������, Max(q.����) As ִ�п���
                 From ������ü�¼ B, �շ���ĿĿ¼ M, ���ű� J, ���ű� Q
                 Where b.No = v_Nos And b.��¼���� = 4 And Nvl(b.����״̬, 0) = 0 And
                       b.��¼״̬ = 0 And b.�շ�ϸĿid = m.Id And b.��������id = j.Id(+) And b.ִ�в���id = q.Id(+)
                 Group By b.No, b.�վݷ�Ŀ, b.����id, b.ִ�в���id, b.���˿���ID, b.������, b.������Ŀid, b.�շ����, b.�Ǽ�ʱ��, b.�۸񸸺�, b.��������, b.���,
                          b.�շ�ϸĿid, b.���㵥λ, b.���ӱ�־
                 Order By ���) Loop
      Zl_����ԤԼ�Һż�¼_Update(c_����.No, c_����.���, c_����.�۸񸸺�, c_����.��������, c_����.�շ����, c_����.�շ�ϸĿid, c_����.����, c_����.����, c_����.������Ŀid,
                         c_����.�վݷ�Ŀ, c_����.Ӧ�ս��, c_����.ʵ�ս��, c_����.���ӱ�־, Null, Null, Null, Null, c_����.���˿���ID, c_����.ִ�в���id);
      n_���ʽ�� := n_���ʽ�� + c_����.ʵ�ս��;
    End Loop;
  
    --����ܽ���Ƿ���ȷ 
    If Nvl(n_����, 0) = 0 Then
      n_���� := Nvl(n_�շ��ܶ�, 0) - Nvl(n_���ʽ��, 0);
      If Abs(n_����) > 1.00 Then
        v_Err_Msg := '���ݽɿ�����ʵ�ʽ�������̫��!';
        Raise Err_Item;
      End If;
    End If;
    If Nvl(n_���ʽ��, 0) + Nvl(n_����, 0) <> Nvl(n_�շ��ܶ�, 0) Then
      v_Err_Msg := 'ָ���Ľɷѵ��ݵ��ܽ���,������ѡ��ɷѵ���!';
      Raise Err_Item;
    End If;
  
    For c_���㷽ʽ In (Select Extractvalue(b.Column_Value, '/JS/JSKLB') As ���㿨���,
                          Extractvalue(b.Column_Value, '/JS/JSKH') As ���㿨��,
                          Extractvalue(b.Column_Value, '/JS/JSFS') As ���㷽ʽ,
                          Extractvalue(b.Column_Value, '/JS/JSJE') As ������,
                          Extractvalue(b.Column_Value, '/JS/JYLSH') As ������ˮ��,
                          Extractvalue(b.Column_Value, '/JS/JYSM') As ����˵��, Extractvalue(b.Column_Value, '/JS/ZY') As ժҪ,
                          Extractvalue(b.Column_Value, '/JS/SFCYJ') As �Ƿ��Ԥ��,
                          Extractvalue(b.Column_Value, '/JS/SFXFK') As �Ƿ����ѿ�
                   From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B) Loop
      If Nvl(c_���㷽ʽ.�Ƿ��Ԥ��, 0) = 1 Then
        n_Ԥ��֧�� := c_���㷽ʽ.������;
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
                Select ID Into n_�����id From ҽ�ƿ���� Where ���� = c_���㷽ʽ.���㿨��� And Nvl(�Ƿ�����, 0) = 1;
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
    End Loop;
  
    --ԤԼ����
    Select a.ִ�в���id, a.�շ�ϸĿid, c.Id, a.ִ����, b.�ű�, b.�����, b.����ʱ��, a.�ѱ�, b.����
    Into n_����id, n_��Ŀid, n_ҽ��id, v_ҽ������, v_����, n_�����, d_����ʱ��, v_�ѱ�, n_����
    From ������ü�¼ A, ���˹Һż�¼ B, ��Ա�� C
    Where a.No = v_Nos And a.��¼���� = 4 And a.��� = 1 And a.No = b.No And a.ִ���� = c.����(+);
    Select Decode(To_Number(zl_GetSysParameter('�Ŷӽк�ģʽ', 1113, 100)), 0, 0, 1) Into n_���ɶ��� From Dual;
  
    Zl_ԤԼ�ҺŽ���_Insert(v_Nos, Null, Null, n_����id, Zl_����(v_����), n_����id, n_�����, v_����, v_�Ա�, v_����, v_ҽ�Ƹ��ʽ����, v_�ѱ�, v_���㷽ʽ,
                     n_��֧ͨ��, n_Ԥ��֧��, Null, d_����ʱ��, n_����, v_����Ա����, v_����Ա����, n_���ɶ���, d_�շ�ʱ��, n_�����id, n_���㿨���, v_���㿨��,
                     v_������ˮ��, v_����˵��, Null, 0, 0, Null, 1);
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
    Zl_���˹ҺŻ���_Update(v_ҽ������, n_ҽ��id, n_��Ŀid, n_����id, d_����ʱ��, 2, v_����, 1);
  End If;
  v_Temp := '<CZSJ>' || To_Char(d_�շ�ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Payment;
/

--90943:������,2015-11-30,�Һ�ȡ�ƻ�ID����
Create Or Replace Procedure Zl_Third_Regist
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --����:HIS�Һ�
  --���:Xml_In:
  --<IN>
  --   <CZFS>3</CZFS>    //������ʽ
  --   <HM>����</HM>    //����
  --   <HX>����</HX>     //����
  --   <JKFS>0</JKFS>  //�ɿʽ,0-�ҺŻ�ԤԼ�ɿ�;1-ԤԼ���ɿ�
  --   <YYSJ>2014-10-21 </YYSJ>    //ԤԼ���� YYYY-MM-DD,��ʱ�η���ſ�����Ҫ����ʱ��
  --   <JE>���</JE>     //���
  --   <JSLIST>
  --     <JS>            //������Ϣ���Һ�Ŀǰ��֧��һ�����ṹ���շ�һ�£��Ժ����չ
  --       <JSKLB>���㿨���</JSKLB>    //���㿨���
  --       <JSKH>֧�����ʺ�</JSKH>           //���㿨��(֧�����ʺ�)
  --       <JYSM>����˵��</JYSM>            //˵�����̶���֧����
  --       <JYLSH>��ˮ��</JYLSH>           //��ˮ�ţ���������
  --       <JSFS>���㷽ʽ</JSFS>            //���㷽ʽ:�ֽ�֧Ʊ�������������,���Դ���
  --       <JSJE>������</JSJE>            //������
  --       <ZY>ժҪ</ZY>                  //ժҪ
  --       <SFCYJ></SFCYJ>              //�Ƿ��Ԥ�����Һ�Ŀǰ����
  --       <SFXFK></SFXFK>              //�Ƿ����ѿ�,�Һ�Ŀǰ����
  --       <EXPENDLIST>                 //��չ��Ϣ
  --         <EXPEND>
  --           <JYMC>��������</JYMC>        //��������
  --           <JYLR>��������<JYLR>         //��������
  --         </EXPEND>
  --         <EXPEND>
  --           ...
  --         </EXPEND>
  --       </EXPENDLIST>
  --     </JS>
  --   </JSLIST>
  --   <HZDW>������λ</HZDW>        //������λ����
  --   <YYFS>֧����<YYFS>    //ԤԼ��ʽ,����������֧����
  --   <BRID>����ID</BRID>     //����ID
  --   <BRLX></BRLX>             //ҽ����������
  --   <FB>��ͨ</FB>               //���˷ѱ𣬿��Բ���
  --   <JQM>������</JQM>            //������
  --</IN>

  --����:Xml_Out
  --<OUTPUT>
  -- <GHDH>�Һŵ���</GHDH>          //�Һŵ���
  -- <CZSJ>����ʱ��</CZSJ>          //HIS�ĵǼ�ʱ��
  -- <ERROR><MSG>������Ϣ</MSG></ERROR>  //����ʱ����
  --</ OUTPUT>
  --------------------------------------------------------------------------------------------------
  v_����       �ҺŰ���.����%Type;
  d_����ʱ��   Date;
  d_ԭʼʱ��   Date;
  d_�Ǽ�ʱ��   Date;
  v_���       Varchar2(200);
  n_Ӧ�ս��   ������ü�¼.Ӧ�ս��%Type;
  v_��ˮ��     ����Ԥ����¼.������ˮ��%Type;
  v_˵��       ������ü�¼.ժҪ%Type;
  n_����id     ������Ϣ.����id%Type;
  v_ԤԼ��ʽ   ԤԼ��ʽ.����%Type;
  v_��������� ҽ�ƿ����.����%Type;
  v_���㿨��   ����Ԥ����¼.����%Type;
  n_�����     ������ü�¼.��ʶ��%Type;
  v_����       ������ü�¼.����%Type;
  v_�Ա�       ������ü�¼.�Ա�%Type;
  v_����       ������ü�¼.����%Type;
  v_���ʽ   ������ü�¼.���ʽ%Type;
  v_�ѱ�       ������ü�¼.�ѱ�%Type;
  v_No         ���˹Һż�¼.No%Type;
  v_���㷽ʽ   ҽ�ƿ����.���㷽ʽ%Type;
  v_�շ����   ������ü�¼.�շ����%Type;
  n_�շ�ϸĿid ������ü�¼.�շ�ϸĿid%Type;
  n_��׼����   ������ü�¼.��׼����%Type;
  n_������Ŀid ������ü�¼.������Ŀid%Type;
  n_���ηѱ�   �շ���ĿĿ¼.���ηѱ�%Type;
  v_�վݷ�Ŀ   ������ü�¼.�վݷ�Ŀ%Type;
  n_���˿���id ������ü�¼.���˿���id%Type;
  n_��������id ������ü�¼.��������id%Type;
  v_����Ա��� ������ü�¼.����Ա���%Type;
  v_����Ա���� ������ü�¼.����Ա����%Type;
  v_ҽ������   �ҺŰ���.ҽ������%Type;
  n_ҽ��id     �ҺŰ���.ҽ��id%Type;
  n_����id     ������ü�¼.����id%Type;
  n_�����id   ҽ�ƿ����.Id%Type;
  v_�Ű�       �ҺŰ���.����%Type;
  n_����id     �ҺŰ���.Id%Type;
  n_�ƻ�id     �ҺŰ��żƻ�.Id%Type;
  n_Ԥ��id     ����Ԥ����¼.Id%Type;
  n_��ſ���   �ҺŰ���.��ſ���%Type;
  n_����       �Һ����״̬.���%Type;
  v_����       �ҺŰ�������.������Ŀ%Type;
  v_��������   ������Ϣ.��������%Type;
  n_����       Number(3);
  n_��ʱ��     Number(3);
  v_������λ   ���˹Һż�¼.������λ%Type;
  v_������     �Һ����״̬.������%Type;
  n_�ɿʽ   Number(3);
  v_Temp       Varchar2(32767); --��ʱXML
  x_Templet    Xmltype; --ģ��XML
  v_Err_Msg    Varchar2(200);
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/HM'), Extractvalue(Value(A), 'IN/HX'),
         To_Date(Extractvalue(Value(A), 'IN/YYSJ'), 'YYYY-MM-DD hh24:mi:ss'), Extractvalue(Value(A), 'IN/JE'),
         Extractvalue(Value(A), 'IN/YYFS'), Extractvalue(Value(A), 'IN/HZDW'), Extractvalue(Value(A), 'IN/BRID'),
         Extractvalue(Value(A), 'IN/BRLX'), Extractvalue(Value(A), 'IN/FB'), Extractvalue(Value(A), 'IN/JQM'),
         Extractvalue(Value(A), 'IN/JKFS')
  Into v_����, n_����, d_ԭʼʱ��, n_Ӧ�ս��, v_ԤԼ��ʽ, v_������λ, n_����id, v_��������, v_�ѱ�, v_������, n_�ɿʽ
  
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  Select Extractvalue(b.Column_Value, '/JS/JSKLB'), Extractvalue(b.Column_Value, '/JS/JSKH'),
         Extractvalue(b.Column_Value, '/JS/JSFS'), Extractvalue(b.Column_Value, '/JS/JYLSH'),
         Extractvalue(b.Column_Value, '/JS/JYSM')
  Into v_���������, v_���㿨��, v_���㷽ʽ, v_��ˮ��, v_˵��
  From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS'))) B;

  d_�Ǽ�ʱ�� := Sysdate;
  d_����ʱ�� := Trunc(d_ԭʼʱ��);
  If v_�������� Is Not Null Then
    Begin
      Select 1 Into n_���� From �������� Where ���� = v_��������;
    Exception
      When Others Then
        v_Err_Msg := 'û�з���Ϊ(' || v_�������� || ')�Ĳ�������';
        Raise Err_Item;
    End;
    Update ������Ϣ Set �������� = Nvl(��������, v_��������) Where ����id = n_����id;
  End If;
  Begin
    Select b.���㷽ʽ, b.Id Into v_���㷽ʽ, n_�����id From ҽ�ƿ���� B Where b.���� = v_��������� And Rownum < 2;
  Exception
    When Others Then
      v_Err_Msg := 'û�з��ָý��㿨�������Ϣ';
      Raise Err_Item;
  End;

  Select a.�����, a.����, a.�Ա�, a.����, Nvl(b.����, c.����)
  Into n_�����, v_����, v_�Ա�, v_����, v_���ʽ
  From ������Ϣ A, ҽ�Ƹ��ʽ B, (Select ���� From ҽ�Ƹ��ʽ Where ȱʡ��־ = '1' And Rownum < 2) C
  Where a.����id = n_����id And a.ҽ�Ƹ��ʽ = b.����(+);
  v_No   := Nextno(12);
  v_Temp := Zl_Identity(1);
  Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_Temp From Dual;
  Select Substr(v_Temp, 0, Instr(v_Temp, ',') - 1) Into v_����Ա��� From Dual;
  Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_����Ա���� From Dual;
  v_Temp := Zl_Identity(2);
  Select Substr(v_Temp, 0, Instr(v_Temp, ',') - 1) Into n_��������id From Dual;
  Select Decode(To_Char(d_ԭʼʱ��, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����', Null)
  Into v_����
  From Dual;
  Begin
    Select ID
    Into n_�ƻ�id
    From (Select ID
           From �ҺŰ��żƻ�
           Where ���� = v_���� And d_ԭʼʱ�� Between Nvl(��Чʱ��, To_Date('1900-01-01', 'YYYY-MM-DD')) And
                 Nvl(ʧЧʱ��, To_Date('3000-01-01', 'YYYY-MM-DD')) And ���ʱ�� Is Not Null
           Order By ��Чʱ�� Desc)
    Where Rownum < 2;
  Exception
    When Others Then
      Select ID Into n_����id From �ҺŰ��� Where ���� = v_����;
  End;
  If Nvl(n_�ƻ�id, 0) <> 0 Then
    --�Ӽƻ���ȡ��Ϣ
    Select a.��Ŀid, b.����id, a.ҽ������, a.ҽ��id,
           Decode(To_Char(d_����ʱ��, 'D'), '1', a.����, '2', a.��һ, '3', a.�ܶ�, '4', a.����, '5', a.����, '6', a.����, '7', a.����,
                   Null), Nvl(a.��ſ���, 0)
    Into n_�շ�ϸĿid, n_���˿���id, v_ҽ������, n_ҽ��id, v_�Ű�, n_��ſ���
    From �ҺŰ��żƻ� A, �ҺŰ��� B
    Where a.Id = n_�ƻ�id And b.Id = a.����id;
    Select Count(Rownum) Into n_��ʱ�� From �Һżƻ�ʱ�� Where ���� = v_���� And �ƻ�id = n_�ƻ�id And Rownum < 2;
    --������λ���
    If v_������λ Is Not Null Then
      Begin
        Select 1 Into n_���� From ������λ�ƻ����� Where �ƻ�id = n_�ƻ�id And ���� = 0 And ������λ = v_������λ;
      Exception
        When Others Then
          n_���� := 0;
      End;
    End If;
    If n_���� = 1 Then
      v_Err_Msg := '����ĺ�����λ�ڴ˺����ϱ����ã�';
      Raise Err_Item;
    End If;
    If n_��ʱ�� = 1 And n_��ſ��� = 0 Then
      d_����ʱ�� := d_ԭʼʱ��;
      Select ���
      Into n_����
      From �Һżƻ�ʱ��
      Where �ƻ�id = n_�ƻ�id And ���� = v_���� And To_Char(��ʼʱ��, 'hh24:mi:ss') = To_Char(d_����ʱ��, 'hh24:mi:ss');
    Else
      Begin
        Select To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || '' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'YYYY-MM-DD hh24:mi:ss')
        Into d_����ʱ��
        From �Һżƻ�ʱ��
        Where �ƻ�id = n_�ƻ�id And ���� = v_���� And ��� = Nvl(n_����, 0);
      Exception
        When Others Then
          If n_��ʱ�� = 1 And n_��ſ��� = 1 Then
            Select To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || '' || To_Char(Max(����ʱ��), 'hh24:mi:ss'),
                            'YYYY-MM-DD hh24:mi:ss')
            Into d_����ʱ��
            From �Һżƻ�ʱ��
            Where �ƻ�id = n_�ƻ�id And ���� = v_����;
          Else
            Select To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || '' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'YYYY-MM-DD hh24:mi:ss')
            Into d_����ʱ��
            From ʱ���
            Where ʱ��� = v_�Ű�;
          End If;
          If d_����ʱ�� < d_�Ǽ�ʱ�� Then
            d_����ʱ�� := d_�Ǽ�ʱ��;
          End If;
      End;
    End If;
  Else
    --�Ӱ��Ŷ�ȡ��Ϣ
    Select b.��Ŀid, b.����id, b.ҽ������, b.ҽ��id,
           Decode(To_Char(d_����ʱ��, 'D'), '1', b.����, '2', b.��һ, '3', b.�ܶ�, '4', b.����, '5', b.����, '6', b.����, '7', b.����,
                   Null), Nvl(b.��ſ���, 0)
    Into n_�շ�ϸĿid, n_���˿���id, v_ҽ������, n_ҽ��id, v_�Ű�, n_��ſ���
    From �ҺŰ��� B
    Where b.Id = n_����id;
    Select Count(Rownum) Into n_��ʱ�� From �ҺŰ���ʱ�� Where ���� = v_���� And ����id = n_����id And Rownum < 2;
    --������λ���
    If v_������λ Is Not Null Then
      Begin
        Select 1 Into n_���� From ������λ���ſ��� Where ����id = n_����id And ���� = 0 And ������λ = v_������λ;
      Exception
        When Others Then
          n_���� := 0;
      End;
    End If;
    If n_���� = 1 Then
      v_Err_Msg := '����ĺ�����λ�ڴ˺����ϱ����ã�';
      Raise Err_Item;
    End If;
    If n_��ʱ�� = 1 And n_��ſ��� = 0 Then
      d_����ʱ�� := d_ԭʼʱ��;
      Select ���
      Into n_����
      From �ҺŰ���ʱ��
      Where ����id = n_����id And ���� = v_���� And To_Char(��ʼʱ��, 'hh24:mi:ss') = To_Char(d_����ʱ��, 'hh24:mi:ss');
    Else
      Begin
        Select To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || '' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'YYYY-MM-DD hh24:mi:ss')
        Into d_����ʱ��
        From �ҺŰ���ʱ��
        Where ����id = n_����id And ���� = v_���� And ��� = Nvl(n_����, 0);
      Exception
        When Others Then
          If n_��ʱ�� = 1 And n_��ſ��� = 1 Then
            Select To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || '' || To_Char(Max(����ʱ��), 'hh24:mi:ss'),
                            'YYYY-MM-DD hh24:mi:ss')
            Into d_����ʱ��
            From �ҺŰ���ʱ��
            Where ����id = n_����id And ���� = v_����;
          Else
            Select To_Date(To_Char(d_����ʱ��, 'yyyy-mm-dd') || '' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'YYYY-MM-DD hh24:mi:ss')
            Into d_����ʱ��
            From ʱ���
            Where ʱ��� = v_�Ű�;
          End If;
          If d_����ʱ�� < d_�Ǽ�ʱ�� Then
            d_����ʱ�� := d_�Ǽ�ʱ��;
          End If;
      End;
    End If;
  End If;

  Select a.���, b.�ּ�, b.������Ŀid, c.�վݷ�Ŀ, a.���ηѱ�
  Into v_�շ����, n_��׼����, n_������Ŀid, v_�վݷ�Ŀ, n_���ηѱ�
  From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C
  Where a.Id = n_�շ�ϸĿid And b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And Sysdate Between b.ִ������ And
        Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD')) And Rownum < 2;

  Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
  Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;

  If Trunc(d_����ʱ��) <> Trunc(Sysdate) Then
    If Nvl(n_�ɿʽ, 0) = 0 Then
      Zl_���������Һ�_Insert(3, n_����id, v_����, n_����, v_No, Null, v_���㷽ʽ, Null, d_����ʱ��, d_�Ǽ�ʱ��, v_������λ, n_Ӧ�ս��, Null, Null,
                       v_��ˮ��, v_˵��, v_ԤԼ��ʽ, n_Ԥ��id, n_�����id, Null, 1, n_����id, 0, Null, Null, v_���㿨��, 1, v_�ѱ�, v_������, 1);
    Else
      Zl_���������Һ�_Insert(2, n_����id, v_����, n_����, v_No, Null, v_���㷽ʽ, Null, d_����ʱ��, d_�Ǽ�ʱ��, v_������λ, n_Ӧ�ս��, Null, Null,
                       v_��ˮ��, v_˵��, v_ԤԼ��ʽ, n_Ԥ��id, n_�����id, Null, 1, n_����id, 0, Null, Null, v_���㿨��, 1, v_�ѱ�, v_������, 1);
    End If;
  Else
    Zl_���������Һ�_Insert(1, n_����id, v_����, n_����, v_No, Null, v_���㷽ʽ, Null, d_����ʱ��, d_�Ǽ�ʱ��, v_������λ, n_Ӧ�ս��, Null, Null, v_��ˮ��,
                     v_˵��, v_ԤԼ��ʽ, n_Ԥ��id, n_�����id, Null, 1, n_����id, 0, Null, Null, v_���㿨��, 1, v_�ѱ�, v_������, 1);
  End If;

  For c_��չ��Ϣ In (Select Extractvalue(b.Column_Value, '/EXPEND/JYMC') As ��������,
                        Extractvalue(b.Column_Value, '/EXPEND/JYLR') As ��������
                 From Table(Xmlsequence(Extract(Xml_In, '/IN/JSLIST/JS/EXPENDLIST/EXPEND'))) B) Loop
    Zl_�������㽻��_Insert(n_�����id, 0, v_���㿨��, n_����id, c_��չ��Ϣ.�������� || '|' || c_��չ��Ϣ.��������, 0);
  End Loop;

  v_Temp := '<GHDH>' || v_No || '</GHDH>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<CZSJ>' || To_Char(d_�Ǽ�ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Regist;
/

--90512:����,2015-11-17,�ϴβɹ�����Ϣ����
Create Or Replace Procedure Zl_ҩƷ�������_Verify
(
  No_In     In ҩƷ�շ���¼.No%Type := Null,
  �����_In In ҩƷ�շ���¼.�����%Type := Null
) Is
  Err_Isverified Exception;
  Err_Isbatch Exception;
  v_Druginf Varchar2(50); --ԭ���������ڷ�����ҩƷ��Ϣ

  Cursor c_ҩƷ�շ���¼ Is
    Select a.Id, a.ʵ������, a.���۽��, a.���ۼ�, a.���, a.�ⷿid, a.ҩƷid, a.����, a.�ɱ���, a.����, a.Ч��, a.����, a.������id, a.��������, a.��׼�ĺ�,
           a.��ҩ��λid, Nvl(b.�Ƿ���, 0) As ʱ��
    From ҩƷ�շ���¼ A, �շ���ĿĿ¼ B
    Where a.ҩƷid = b.Id And a.No = No_In And a.���� = 4 And a.��¼״̬ = 1
    Order By a.ҩƷid;
Begin
  Update ҩƷ�շ���¼
  Set ����� = �����_In, ������� = Sysdate
  Where NO = No_In And ���� = 4 And ��¼״̬ = 1 And ����� Is Null;

  If Sql%RowCount = 0 Then
    Raise Err_Isverified;
  End If;

  --��Ҫ���ԭ���������ڷ�����ҩƷ�����ܶ������
  Begin
    Select Distinct '(' || i.���� || ')' || Nvl(n.����, i.����) As ҩƷ��Ϣ
    Into v_Druginf
    From ҩƷ�շ���¼ A, ҩƷ��� B, �շ���ĿĿ¼ I, �շ���Ŀ���� N
    Where a.ҩƷid = b.ҩƷid And a.ҩƷid = i.Id And a.ҩƷid = n.�շ�ϸĿid(+) And n.����(+) = 3 And a.No = No_In And a.���� = 4 And
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
         Where b.ҩƷid = a.ҩƷid And a.No = No_In And a.���� = 4 And a.��¼״̬ = 1 And Nvl(a.����, 0) > 0 And
               (Nvl(b.ҩ�����, 0) = 0 Or
               (Nvl(b.ҩ������, 0) = 0 And
               a.�ⷿid In (Select ����id From ��������˵�� Where (�������� Like '%ҩ��') Or (�������� Like '�Ƽ���')))));

  For v_ҩƷ�շ���¼ In c_ҩƷ�շ���¼ Loop
    --����ҩƷ�������Ӧ����
    Zl_ҩƷ���_Update(v_ҩƷ�շ���¼.Id, 0);
    Update ҩƷ���
    Set �ɱ��� = v_ҩƷ�շ���¼.�ɱ���, �ϴ��ۼ� = Decode(v_ҩƷ�շ���¼.ʱ��, 1, v_ҩƷ�շ���¼.���ۼ�, Null), �ϴι�Ӧ��id = v_ҩƷ�շ���¼.��ҩ��λid,
        �ϴ����� = v_ҩƷ�շ���¼.����, �ϴ��������� = v_ҩƷ�շ���¼.��������, �ϴβ��� = v_ҩƷ�շ���¼.����, �ϴ���׼�ĺ� = v_ҩƷ�շ���¼.��׼�ĺ�
    Where ҩƷid = v_ҩƷ�շ���¼.ҩƷid;
  End Loop;
Exception
  When Err_Isverified Then
    Raise_Application_Error(-20101, '[ZLSOFT]�õ����Ѿ���������ˣ�[ZLSOFT]');
  When Err_Isbatch Then
    Raise_Application_Error(-20102, '[ZLSOFT]�õ����а�����һ��ԭ�������������ڷ�����ҩƷ[' || v_Druginf || ']��������ˣ�[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�������_Verify;
/

--92493:������,2016-01-08,��������ǰ������������ֶ� ���没��
Create Or Replace Procedure Zl_��������ǰ��_Append
(
  �ļ�id_In   In ��������ǰ��.�ļ�id%Type,
  ����_In     In Varchar2 := Null, --���ֺŷָ��ļ���id��
  ���_In     In Varchar2 := Null, --���ֺŷָ��ļ���id��
  ���没��_In In Varchar2 := Null
) Is
  v_Disease Varchar2(4000);
  n_����id  ��������ǰ��.����id%Type;
  n_���id  ��������ǰ��.���id%Type;
Begin

  Update �����ļ��б� Set ���� = ���� Where ID = �ļ�id_In;

  If Sql%RowCount = 0 Then
    Raise No_Data_Found;
  End If;

  If ����_In Is Not Null Then
    v_Disease := ����_In || ';';
    While v_Disease Is Not Null Loop
      n_����id  := To_Number(Substr(v_Disease, 1, Instr(v_Disease, ';') - 1));
      v_Disease := Substr(v_Disease, Instr(v_Disease, ';') + 1);
      Update ��������ǰ�� Set �ļ�id = �ļ�id, ���没�� = ���没��_In Where �ļ�id = �ļ�id_In And ����id = n_����id;
      If Sql%RowCount = 0 Then
        Insert Into ��������ǰ�� (�ļ�id, ����id, ���没��) Values (�ļ�id_In, n_����id, ���没��_In);
      End If;
    End Loop;
  End If;

  If ���_In Is Not Null Then
    v_Disease := ���_In || ';';
    While v_Disease Is Not Null Loop
      n_���id  := To_Number(Substr(v_Disease, 1, Instr(v_Disease, ';') - 1));
      v_Disease := Substr(v_Disease, Instr(v_Disease, ';') + 1);
      Update ��������ǰ�� Set �ļ�id = �ļ�id, ���没�� = ���没��_In Where �ļ�id = �ļ�id_In And ���id = n_���id;
      If Sql%RowCount = 0 Then
        Insert Into ��������ǰ�� (�ļ�id, ���id, ���没��) Values (�ļ�id_In, n_���id, ���没��_In);
      End If;
    End Loop;
  End If;
Exception
  When No_Data_Found Then
    Raise_Application_Error(-20101, '[ZLSOFT]û���ҵ��ļ��������Ѿ���ɾ����[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��������ǰ��_Append;
/




Insert Into zlFilesUpgrade (�ļ�����,�ļ���,�汾��,�޸�����,����ϵͳ,ҵ�񲿼�,��װ·��,�ļ�˵��,ǿ�Ƹ���,�Զ�ע��,��������,���) select 1,'zl9Disease.dll','', Null ,'1','zl9Cisjob','[APPSOFT]\APPLY','�������沿��','0','1',sysdate,��� from Dual a,(Select max(to_number(���))+1 ��� from zlfilesupgrade) b where not exists (select 1 from zlfilesupgrade where upper(�ļ���)='ZL9DISEASE.DLL');
Insert Into zlFilesUpgrade (�ļ�����,�ļ���,�汾��,�޸�����,����ϵͳ,ҵ�񲿼�,��װ·��,�ļ�˵��,ǿ�Ƹ���,�Զ�ע��,��������,���) select 1,'zlDisReportCard.dll','', Null ,'1','zl9Cisjob','[APPSOFT]\APPLY','�����������ò���','0','1',sysdate,��� from Dual a,(Select max(to_number(���))+1 ��� from zlfilesupgrade) b where not exists (select 1 from zlfilesupgrade where upper(�ļ���)='ZLDISREPORTCARD.DLL');
Insert Into zlFilesUpgrade (�ļ�����,�ļ���,�汾��,�޸�����,����ϵͳ,ҵ�񲿼�,��װ·��,�ļ�˵��,ǿ�Ƹ���,�Զ�ע��,��������,���) select 1,'zl9PacsImageCap.dll','', Null ,'1','zl9PacsWork','[APPSOFT]\APPLY','��Ƶ�ɼ�����','0','1',sysdate,��� from Dual a,(Select max(to_number(���))+1 ��� from zlfilesupgrade) b where not exists (select 1 from zlfilesupgrade where upper(�ļ���)='ZL9PACSIMAGECAP.DLL');

---------------------------------------------------------------------------------------------------
--����ϵͳ�������İ汾��
-------------------------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.34.60' Where ���=&n_System;
--�����汾��
Commit;