--[��������]1
--[�����߰汾��]10.34.130
--���ű�֧�ִ�ZLHIS+ v10.34.130 ������ v10.34.140
--����ϵͳ�����ߵ�¼PLSQL��ִ�����нű�
--�ű�ִ�к����ֹ�������������
Define n_System=100;
-------------------------------------------------------------------------------
--�ṹ��������
-------------------------------------------------------------------------------
--116339:����,2017-12-14,��¼��Ŀ����̷��������޸�
alter table �����¼��Ŀ add ȱʡֵ VARCHAR2(100);

--115026:������,2017-12-04,����Σ��ֵ
Create Table ����Σ��ֵ��¼(
  ID Number(18),
  ������Դ varchar2(100),    
  ����ID number(18),
  ��ҳID NUMBER(5),
  �Һŵ� VARCHAR2(8),
  Ӥ�� number(3),
  ���� VARCHAR2(100),
  �Ա� VARCHAR2(4),
  ���� varchar2(20),    
  ҽ��ID number(18),
  �걾ID NUMBER(18),   
  Σ��ֵ���� varchar2(2000),       
  ����ʱ�� date,
  �������ID number(18),
  ������ VARCHAR2(20),    
  ������� varchar2(2000),
  ȷ��ʱ�� date,          
  ȷ���� VARCHAR2(20),
  ȷ�Ͽ���ID number(18),       
  ״̬ number(3),      
  �Ƿ�Σ��ֵ number(1),  
  ��ת�� Number(3)
) TABLESPACE zl9CisRec;

Create Sequence ����Σ��ֵ��¼_ID Start With 1;

Alter Table ����Σ��ֵ��¼ Add Constraint ����Σ��ֵ��¼_PK Primary Key (ID) Using Index Tablespace zl9Indexcis;
Alter Table ����Σ��ֵ��¼ Add Constraint ����Σ��ֵ��¼_FK_����ID Foreign Key (����ID) References ������Ϣ(����ID);
Alter Table ����Σ��ֵ��¼ Add Constraint ����Σ��ֵ��¼_FK_��ҳID Foreign Key (����ID,��ҳID) References ������ҳ(����ID,��ҳID);
Alter Table ����Σ��ֵ��¼ Add Constraint ����Σ��ֵ��¼_FK_�������ID Foreign Key (�������ID) References ���ű�(ID);
Alter Table ����Σ��ֵ��¼ Add Constraint ����Σ��ֵ��¼_FK_ȷ�Ͽ���ID Foreign Key (ȷ�Ͽ���ID) References ���ű�(ID);
Alter Table ����Σ��ֵ��¼ Add Constraint ����Σ��ֵ��¼_FK_ҽ��ID Foreign Key (ҽ��ID) References ����ҽ����¼(ID);
Create Index ����Σ��ֵ��¼_IX_����ID On ����Σ��ֵ��¼(����ID,��ҳID)  Tablespace zl9Indexcis;
Create Index ����Σ��ֵ��¼_IX_�Һŵ� On ����Σ��ֵ��¼(�Һŵ�)  Tablespace zl9Indexcis;
Create Index ����Σ��ֵ��¼_IX_ҽ��ID On ����Σ��ֵ��¼(ҽ��ID) Tablespace zl9Indexcis;
Create Index ����Σ��ֵ��¼_IX_����ʱ�� On ����Σ��ֵ��¼(����ʱ��)  Tablespace zl9Indexcis;
Create Index ����Σ��ֵ��¼_IX_��ת�� On ����Σ��ֵ��¼(��ת��) Tablespace zl9Indexcis;

CREATE TABLE ����Σ��ֵҽ��(
    Σ��ֵID NUMBER(18),
    ҽ��ID NUMBER(18),
    ��ת�� Number(3))
    TABLESPACE zl9CisRec;

Alter Table ����Σ��ֵҽ�� Add Constraint ����Σ��ֵҽ��_UQ_Σ��ֵID Unique (Σ��ֵID,ҽ��ID) Using Index Tablespace zl9Indexcis;
Alter Table ����Σ��ֵҽ�� Add Constraint ����Σ��ֵҽ��_FK_Σ��ֵID Foreign Key (Σ��ֵID) References ����Σ��ֵ��¼(ID) On Delete Cascade;
Alter Table ����Σ��ֵҽ�� Add Constraint ����Σ��ֵҽ��_FK_ҽ��ID Foreign Key (ҽ��ID) References ����ҽ����¼(ID) On Delete Cascade; 
Alter Table ����Σ��ֵҽ�� Modify Σ��ֵID Constraint ����Σ��ֵҽ��_NN_Σ��ֵID Not Null;
Create Index ����Σ��ֵҽ��_IX_ҽ��ID On ����Σ��ֵҽ��(ҽ��ID) Tablespace zl9Indexcis;
Create Index ����Σ��ֵҽ��_IX_��ת�� On ����Σ��ֵҽ��(��ת��) Tablespace zl9Indexcis;

CREATE TABLE ����Σ��ֵ����(
    Σ��ֵID NUMBER(18),
    �ĵ�ID VARCHAR2(32),
    ���ĵ�ID VARCHAR2(32),
    ���� varchar2(100),
    ����� varchar2(20),
    ���ʱ�� date,
    ��ת�� Number(3))
    TABLESPACE zl9EprDat;    

Alter Table ����Σ��ֵ���� Add Constraint ����Σ��ֵ����_UQ_Σ��ֵID Unique (Σ��ֵID,�ĵ�ID,���ĵ�ID) Using Index Tablespace zl9Indexcis;
Alter Table ����Σ��ֵ���� Add Constraint ����Σ��ֵ����_FK_Σ��ֵID Foreign Key (Σ��ֵID) References ����Σ��ֵ��¼(ID) On Delete Cascade;
Alter Table ����Σ��ֵ���� Modify Σ��ֵID Constraint ����Σ��ֵ����_NN_Σ��ֵID Not Null;
Create Index ����Σ��ֵ����_IX_��ת�� On ����Σ��ֵ����(��ת��) Tablespace zl9Indexcis;

--113432:�ƽ�,2017-11-09,�°�����޸���챨����Ϊ�ַ���
alter table RISҽ��ʧ�ܼ�¼ rename column ��챨���� to ��챨����_bak;
alter table RISҽ��ʧ�ܼ�¼ add ��챨���� VARCHAR2(20);
update RISҽ��ʧ�ܼ�¼ set ��챨����=to_char(��챨����_bak);

--115695:����,2017-11-09,�޸ı����ֶγ���
Alter Table ���Ʒ���Ŀ¼ Modify(���� Varchar2(20));

--111635:���Ʊ�,2017-07-14,XML�Զ������뵥
CREATE TABLE �Զ������뵥�ļ�(
  �ļ�ID NUMBER(18),
  �ļ��� VARCHAR2(200),
  ��� number(2),
  ���� CLOB,
  ������ VARCHAR2(20),
  ����ʱ�� DATE
  )
TABLESPACE zl9EprLob;

CREATE TABLE ҽ�����뵥�ļ�(
  ҽ��ID NUMBER(18),
  �ļ�ID NUMBER(18),
  �ļ��� VARCHAR2(200),
  ��� number(2),
  ���� CLOB,
  ��ת�� Number(3)
  )
TABLESPACE zl9EprLob;

Alter Table �����ļ��б� Add(��ʽ Number(5));

Alter Table �Զ������뵥�ļ� Add Constraint �Զ������뵥�ļ�_PK Primary Key (�ļ�ID,���) Using Index Tablespace zl9Indexcis;

Alter Table ҽ�����뵥�ļ� Add Constraint ҽ�����뵥�ļ�_PK Primary Key (ҽ��ID,�ļ�ID,���) Using Index Tablespace zl9Indexcis;

Create Index ҽ�����뵥�ļ�_IX_��ת�� On ҽ�����뵥�ļ�(��ת��) Tablespace zl9Indexcis;

Create Index ҽ�����뵥�ļ�_IX_�ļ�ID On ҽ�����뵥�ļ�(�ļ�ID) Tablespace zl9Indexcis;

--111635:���Ʊ�,2017-07-14,XML�Զ������뵥
Alter Table �Զ������뵥�ļ� Add Constraint �Զ������뵥�ļ�_FK_�ļ�ID Foreign Key (�ļ�ID) References �����ļ��б�(ID) On Delete Cascade;
Alter Table ҽ�����뵥�ļ� Add Constraint ҽ�����뵥�ļ�_FK_ҽ��ID Foreign Key (ҽ��ID) References ����ҽ����¼(ID) On Delete Cascade;
Alter Table ҽ�����뵥�ļ� Add Constraint ҽ�����뵥�ļ�_FK_�ļ�ID Foreign Key (�ļ�ID) References �����ļ��б�(ID) On Delete Cascade;

--114434:��ΰ��,2017-11-17,�������׺�����ҩ�������
Alter Table ����ҽ����¼ Add ������� Number(18);
Create Sequence ����ҽ����¼_������� Start With 1 Cache 100;
Create Index ����ҽ����¼_IX_������� On ����ҽ����¼(�������) Pctfree 5 Tablespace zl9Indexcis;


-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--118463:����,2017-12-18,������ҩ�����Զ�Ĭ��Ϊ��ҩ״̬������
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1342, 0, 0, 0, 0, 28, '��ҩ��������Ĭ��Ϊ��ҩ״̬', '0', '0', '��ҩ����������ҪĬ��Ϊ��ҩ״̬.1-��;0-����'
  From Dual;

--118267:����һ,2017-12-16,LISͼƬ����ת������ִ���ļ�
Insert Into Zltools.Zlfilesupgrade
  (���, ��������, ��װ·��, �ļ�����, �ļ���, �汾��, �޸�����, ����ϵͳ, ҵ�񲿼�, Md5, �ļ�˵��, �Զ�ע��, ǿ�Ƹ���, ���Ӱ�װ·��)
  Select ���, To_Date('2017-12-16 02:25:54', 'yyyy-mm-dd hh24:mi:ss'), '[APPSOFT]', 0, 'ZLLISPIC2FTP.EXE', Null, Null, Null,
         Null, Null, '��������:LISͼƬ����ת������ִ���ļ�', 0, 0, Null
  From Dual A, (Select Nvl(Max(To_Number(���)), 0) + 1 ��� From zlFilesUpgrade) B
  Where Not Exists (Select 1 From Zltools.Zlfilesupgrade Where Upper(�ļ���) = 'ZLLISPIC2FTP.EXE');

--101301:����,2017-12-06,ȡѪ�����Ϣ��ʾ��ʿվ
insert into ҵ����Ϣ����(����,����,˵��,��������) values ('ZLHIS_BLOOD_003','ȡѪ�������','��ʿȡѪ��ɣ����ѻ�ʿվ',1);

--116339:����,2017-12-14,��¼��Ŀ����̷��������޸�
Declare
  Strdata Varchar2(1000);
  Strpre  Varchar2(1000);
  Strtext Varchar2(30);
  Cursor Cur_Item Is
    Select ��Ŀ���, ��Ŀֵ��, ȱʡֵ From �����¼��Ŀ Where ��Ŀ��ʾ In (2, 3);
Begin

  For Row_Format In Cur_Item Loop
    Strdata := Row_Format.��Ŀֵ��;
    Strpre  := '';
    While Strdata Is Not Null Loop
      If Instr(Strdata, ';', 1) > 0 Then
        Strtext := Substr(Strdata, 1, Instr(Strdata, ';', 1) - 1);
      Else
        Strtext := Strdata;
        Strdata := '';
      End If;
      If Instr(Strtext, '��', 1) = 1 Then
        If Strpre Is Null Then
          Strpre := Substr(Strtext, 2);
        Else
          Strpre := Strpre || ';' || Substr(Strtext, 2);
        End If;
        Strtext := Substr(Strtext, 2);
        Strdata := Substr(Strdata, Instr(Strdata, ';', 1) + 1);
        Strdata := Strpre || ';' || Strdata;
        Update �����¼��Ŀ Set ȱʡֵ = Strtext, ��Ŀֵ�� = Strdata Where ��Ŀ��� = Row_Format.��Ŀ���;
        Exit;
      Else
        If Strpre Is Null Then
          Strpre := Strtext;
        Else
          Strpre := Strpre || ';' || Strtext;
        End If;
        Strdata := Substr(Strdata, Instr(Strdata, ';', 1) + 1);
      End If;
    End Loop;
  End Loop;
  Update �����¼��Ŀ Set ȱʡֵ = '��', ��Ŀֵ�� = '��;��(��)' Where ��Ŀ���� = '����' And ������Ŀ = 1;
End;
/

--116846:������,2017-12-14,��Ѫ���뱣���Զ��庯�����
Insert Into Zlprocedure(Id, ����, ����, ״̬, ������, ˵��) Values(Zlprocedure_Id.Nextval,2,'Zl1_EX_BloodApplyCheck',3,User,'�¿����޸���Ѫ����ʱ����������֮ǰ�������������ݽ��м�飬��������ʾ����������');

--117641:���ϴ�,2017-12-11,ɾ����Ч����
Delete from Zlparameters Where ϵͳ = &n_System And  ģ�� = 1111 and ������ = '�˺���ʾ��ϸ��Ϣ';

--94173:������,2017-12-08,У��������Ϣ
Insert Into ҵ����Ϣ����(����,����,˵��,��������) 
Select 'ZLHIS_CIS_035','У����������','��ʿУ��ҽ��ʱ��Ϊ����ʱ������һ��֪ͨ��Ϣ��',7 From Dual;

--000000:����,2017-12-06���ڼ���������
Insert Into �ڼ��(�ڼ�,��ʼ����,��ֹ����) 
Select '201802',to_date('2018-02-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2018-02-28 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201803',to_date('2018-03-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2018-03-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201804',to_date('2018-04-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2018-04-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201805',to_date('2018-05-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2018-05-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201806',to_date('2018-06-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2018-06-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201807',to_date('2018-07-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2018-07-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201808',to_date('2018-08-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2018-08-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201809',to_date('2018-09-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2018-09-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201810',to_date('2018-10-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2018-10-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201811',to_date('2018-11-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2018-11-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201812',to_date('2018-12-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2018-12-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201901',to_date('2019-01-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2019-01-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201902',to_date('2019-02-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2019-02-28 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201903',to_date('2019-03-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2019-03-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201904',to_date('2019-04-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2019-04-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201905',to_date('2019-05-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2019-05-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201906',to_date('2019-06-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2019-06-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201907',to_date('2019-07-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2019-07-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201908',to_date('2019-08-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2019-08-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201909',to_date('2019-09-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2019-09-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201910',to_date('2019-10-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2019-10-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201911',to_date('2019-11-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2019-11-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '201912',to_date('2019-12-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2019-12-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202001',to_date('2020-01-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2020-01-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202002',to_date('2020-02-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2020-02-29 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202003',to_date('2020-03-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2020-03-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202004',to_date('2020-04-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2020-04-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202005',to_date('2020-05-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2020-05-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202006',to_date('2020-06-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2020-06-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202007',to_date('2020-07-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2020-07-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202008',to_date('2020-08-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2020-08-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202009',to_date('2020-09-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2020-09-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202010',to_date('2020-10-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2020-10-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202011',to_date('2020-11-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2020-11-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202012',to_date('2020-12-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2020-12-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202101',to_date('2021-01-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2021-01-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202102',to_date('2021-02-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2021-02-28 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202103',to_date('2021-03-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2021-03-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202104',to_date('2021-04-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2021-04-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202105',to_date('2021-05-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2021-05-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202106',to_date('2021-06-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2021-06-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202107',to_date('2021-07-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2021-07-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202108',to_date('2021-08-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2021-08-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202109',to_date('2021-09-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2021-09-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202110',to_date('2021-10-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2021-10-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202111',to_date('2021-11-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2021-11-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202112',to_date('2021-12-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2021-12-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202201',to_date('2022-01-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2022-01-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202202',to_date('2022-02-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2022-02-28 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202203',to_date('2022-03-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2022-03-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202204',to_date('2022-04-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2022-04-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202205',to_date('2022-05-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2022-05-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202206',to_date('2022-06-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2022-06-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202207',to_date('2022-07-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2022-07-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202208',to_date('2022-08-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2022-08-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202209',to_date('2022-09-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2022-09-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202210',to_date('2022-10-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2022-10-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202211',to_date('2022-11-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2022-11-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202212',to_date('2022-12-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2022-12-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202301',to_date('2023-01-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2023-01-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202302',to_date('2023-02-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2023-02-28 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202303',to_date('2023-03-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2023-03-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202304',to_date('2023-04-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2023-04-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202305',to_date('2023-05-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2023-05-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202306',to_date('2023-06-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2023-06-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202307',to_date('2023-07-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2023-07-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202308',to_date('2023-08-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2023-08-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202309',to_date('2023-09-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2023-09-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202310',to_date('2023-10-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2023-10-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202311',to_date('2023-11-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2023-11-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202312',to_date('2023-12-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2023-12-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202401',to_date('2024-01-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2024-01-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202402',to_date('2024-02-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2024-02-29 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202403',to_date('2024-03-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2024-03-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202404',to_date('2024-04-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2024-04-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202405',to_date('2024-05-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2024-05-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202406',to_date('2024-06-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2024-06-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202407',to_date('2024-07-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2024-07-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202408',to_date('2024-08-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2024-08-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202409',to_date('2024-09-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2024-09-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202410',to_date('2024-10-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2024-10-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202411',to_date('2024-11-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2024-11-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202412',to_date('2024-12-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2024-12-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202501',to_date('2025-01-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2025-01-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202502',to_date('2025-02-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2025-02-28 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202503',to_date('2025-03-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2025-03-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202504',to_date('2025-04-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2025-04-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202505',to_date('2025-05-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2025-05-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202506',to_date('2025-06-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2025-06-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202507',to_date('2025-07-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2025-07-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202508',to_date('2025-08-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2025-08-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202509',to_date('2025-09-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2025-09-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202510',to_date('2025-10-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2025-10-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202511',to_date('2025-11-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2025-11-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202512',to_date('2025-12-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2025-12-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202601',to_date('2026-01-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2026-01-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202602',to_date('2026-02-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2026-02-28 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202603',to_date('2026-03-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2026-03-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202604',to_date('2026-04-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2026-04-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202605',to_date('2026-05-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2026-05-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202606',to_date('2026-06-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2026-06-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202607',to_date('2026-07-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2026-07-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202608',to_date('2026-08-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2026-08-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202609',to_date('2026-09-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2026-09-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202610',to_date('2026-10-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2026-10-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202611',to_date('2026-11-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2026-11-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202612',to_date('2026-12-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2026-12-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202701',to_date('2027-01-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2027-01-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202702',to_date('2027-02-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2027-02-28 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202703',to_date('2027-03-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2027-03-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202704',to_date('2027-04-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2027-04-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202705',to_date('2027-05-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2027-05-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202706',to_date('2027-06-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2027-06-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202707',to_date('2027-07-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2027-07-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202708',to_date('2027-08-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2027-08-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202709',to_date('2027-09-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2027-09-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202710',to_date('2027-10-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2027-10-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202711',to_date('2027-11-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2027-11-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202712',to_date('2027-12-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2027-12-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202801',to_date('2028-01-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2028-01-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202802',to_date('2028-02-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2028-02-29 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202803',to_date('2028-03-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2028-03-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202804',to_date('2028-04-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2028-04-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202805',to_date('2028-05-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2028-05-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202806',to_date('2028-06-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2028-06-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202807',to_date('2028-07-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2028-07-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202808',to_date('2028-08-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2028-08-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202809',to_date('2028-09-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2028-09-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202810',to_date('2028-10-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2028-10-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202811',to_date('2028-11-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2028-11-30 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual Union All
Select '202812',to_date('2028-12-01 00:00:00','yyyy-mm-dd hh24:mi:ss'),to_date('2028-12-31 00:00:00','yyyy-mm-dd hh24:mi:ss') From Dual;

--115026:������,2017-12-04,����Σ��ֵ
Insert Into zlPrograms(���,����,˵��,ϵͳ,����) Values( 1284,'����Σ��ֵ��ѯ','���ڶԲ���Σ��ֵ��ѯͳ�Ʒ�����',&n_System,'zl9CISJob');

Insert Into zlMenus
  (���, ID, �ϼ�id, ����, ���, ˵��, ϵͳ, ģ��, �̱���, ͼ��)
  Select ���, Zlmenus_Id.Nextval, ID, 'Σ��ֵ����ϵͳ', 'D', '���ڶԲ���Σ��ֵ�����ѯ������', &n_System, -null, 'Σ��ֵ����', 99
  From zlMenus
  Where ���� = '�ٴ���Ϣϵͳ' And ��� = 'ȱʡ' And ϵͳ = &n_System And ģ�� Is Null;

Insert Into zlMenus
  (���, ID, �ϼ�id, ����, ���, ˵��, ϵͳ, ģ��, �̱���, ͼ��)
  Select a.���, Zlmenus_Id.Nextval, a.Id, b.*
  From (Select ���, ID From zlMenus Where ���� = 'Σ��ֵ����ϵͳ' And ��� = 'ȱʡ' And ϵͳ = &n_System And ģ�� Is Null) A,
       (Select ����, ���, ˵��, ϵͳ, ģ��, �̱���, ͼ��
         From zlMenus
         Where 1 = 0          
         Union All         
         Select '����Σ��ֵ��ѯ', 'D', '���ڶԲ���Σ��ֵ��ѯͳ�Ʒ�����', &n_System, 1284, 'Σ��ֵ����', 99
         From Dual         
         Union All
         Select ����, ���, ˵��, ϵͳ, ģ��, �̱���, ͼ��
         From zlMenus
         Where 1 = 0) B;

--115026:������,2017-12-04,����Σ��ֵ
Insert Into zlBakTables(ϵͳ,���,����,���,ֱ��ת��,ͣ�ô�����)
Select &n_System,8,A.* From (
Select ����,���,ֱ��ת��,ͣ�ô����� From zlBakTables Where 1 = 0 Union All 
      Select '����Σ��ֵ��¼',25,1,-Null From Dual Union All  
      Select '����Σ��ֵҽ��',26,1,-Null From Dual Union All
      Select '����Σ��ֵ����',27,1,-Null From Dual Union All
Select ����,���,ֱ��ת��,ͣ�ô����� From zlBakTables Where 1 = 0) A;

Insert Into zlBakTableindex(ϵͳ,����,������)
Select &n_System,A.* From (
Select ����,������ From zlBakTableindex Where 1 = 0 Union All
Select '����Σ��ֵ��¼','����Σ��ֵ��¼_IX_ҽ��ID' From Dual Union All
Select '����Σ��ֵҽ��','����Σ��ֵҽ��_UQ_Σ��ֵID' From Dual Union All
Select '����Σ��ֵ����','����Σ��ֵ����_UQ_Σ��ֵID' From Dual Union All
Select ����,������ From zlBakTableindex Where 1 = 0) A;

--115026:������,2017-12-04,����Σ��ֵ
Insert into zlTables(ϵͳ,����,��ռ�,����) Values(100,'����Σ��ֵ��¼','ZL9CISREC','B1');

Insert into zlTables(ϵͳ,����,��ռ�,����) Values(100,'����Σ��ֵҽ��','ZL9CISREC','B1');

Insert into zlTables(ϵͳ,����,��ռ�,����) Values(100,'����Σ��ֵ����','zl9EprDat','B1');


--115806:����,2017-11-24,�޸�ʱ��ҩƷ���ʱȡ�ϴ��ۼ۵Ĳ���˵��
Update zlParameters
Set ����˵�� = '��������ҩƷ���⹺������������ʱ���ۼ��ǰ�ʲô��ʽ���ģ� 0-������ʽȡ��Ĭ�ϣ� 1-����ȡ�ϴ��⹺�����ۼ���Ϊ�����ۼ�.'
Where ������ = 'ʱ��ҩƷ���ʱȡ�ϴ��ۼ�';

--111635:���Ʊ�,2017-07-14,XML�Զ������뵥
Insert into zlTables(ϵͳ,����,��ռ�,����) Values(100,'ҽ�����뵥�ļ�','ZL9EPRLOB','B1');
Insert into zlTables(ϵͳ,����,��ռ�,����) Values(100,'�Զ������뵥�ļ�','ZL9EPRLOB','A2');

Insert Into zlBakTables(ϵͳ,���,����,���,ֱ��ת��,ͣ�ô�����)
Select &n_System,8,A.* From (
Select ����,���,ֱ��ת��,ͣ�ô����� From zlBakTables Where 1 = 0 Union All
Select 'ҽ�����뵥�ļ�',24,1,-NULL From Dual Union All
Select ����,���,ֱ��ת��,ͣ�ô����� From zlBakTables Where 1 = 0) A;

Insert Into zlBakTableindex(ϵͳ,����,������)
Select &n_System,A.* From (
Select ����,������ From zlBakTableindex Where 1 = 0 Union All
Select 'ҽ�����뵥�ļ�','ҽ�����뵥�ļ�_PK' From Dual Union All
Select ����,������ From zlBakTableindex Where 1 = 0) A;

--114364:���ϴ�,2017-11-14,�����������̿���
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ,����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1801, 0, 0, 0, 0, 21, '���������������', NULL, '0',
         '���Ʒ���ʱ¼�����뻹��ʹ��ȱʡ����.0-�����ɲ�����������;1-����ʹ��ȱʡ���룬�������������'
  From Dual;

Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ,����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1801, 0, 0, 0, 0, 22, 'ȱʡ����', NULL, NULL,
         '���Ʒ���ʱ¼�����뻹��ʹ��ȱʡ����.������ʽ�����,�����ID,����ȱʡ��ʽ,ȱʡ�̶�����||...'
  From Dual;


--115481:����,2017-11-2,�Ƿ��Զ�������ݽкŴ���
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ,����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1291, 0, 1, 0, 0, 57, '�Զ�������ݺ��д���', '1', '1','�����Ŷӽкź��Ƿ��Զ�������ݴ���,0-���Զ�����;1-�Զ�����'
  From Dual;
   
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ,����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1290, 0, 1, 0, 0, 54, '�Զ�������ݺ��д���', '1', '1','�����Ŷӽкź��Ƿ��Զ�������ݴ���,0-���Զ�����;1-�Զ�����'
  From Dual;

--114920:���˺�,2017-12-11,�������������,��Ҫ���������������
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1257, 1, 0, 0, 0, 18, '�ϴ�ѡ���������', Null, '0',
         '���ڿ����ϴ��Ƿ�ѡ������ʾ����������Ա��ٴν���ʱĬ��:1-�ϴ�ѡ����������ʾ,0-�ϴ�δѡ��������ʾ��'
  From Dual;

Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1150, 1, 1, 0, 0, 43, '�ϴ�ѡ���������', Null, '0',
         '���ڿ����ϴ��Ƿ�ѡ������ʾ����������Ա��ٴν���ʱĬ��:1-�ϴ�ѡ����������ʾ,0-�ϴ�δѡ��������ʾ��'
  From Dual;
 
-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------
--114601:����,2017-12-21,����ҩƷ���ŷ�ҩ�Ļ����鿴Ȩ��
Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1342, '����', User, '������Ǽ�¼', 'SELECT' From Dual;

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1342, '����', User, '�����������', 'SELECT' From Dual;

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1342, '����', User, '������������¼', 'SELECT' From Dual;

--104221:��ΰ��,2017-12-15,���֤�ż��
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,9003,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All 
    Select 'Zl_Fun_Checkidcard','EXECUTE' From Dual Union All 
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

--116846:������,2017-12-14,��Ѫ���뱣���Զ��庯�����
Insert Into zlProgPrivs(ϵͳ,���,������,����,����,Ȩ��) values(&n_System,1252,User,'ҽ���´�','Zl1_EX_BloodApplyCheck','EXECUTE');

Insert Into zlProgPrivs(ϵͳ,���,������,����,����,Ȩ��) values(&n_System,1253,User,'ҽ���´�','Zl1_EX_BloodApplyCheck','EXECUTE');

--115026:������,2017-12-04,����Σ��ֵ
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��,ȱʡֵ)
Select  &n_System,9001,A.* From (
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0 Union All
Select 'Σ��ֵ����',4,'�д�Ȩ��ʱ��������ýӿڶ�Σ��ֵ���еǼ�',1 From Dual Union All 
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select  &n_System,9001,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All 
    Select '����Σ��ֵ��¼_ID','SELECT' From Dual Union All 
    Select '����Σ��ֵҽ��','SELECT' From Dual Union All    
    Select '����Σ��ֵ��¼','SELECT' From Dual Union All    
    Select 'Zl_����Σ��ֵ��¼_Insert','EXECUTE' From Dual Union All 
    Select 'Zl_����Σ��ֵ��¼_Update','EXECUTE' From Dual Union All 
    Select 'Zl_����Σ��ֵ��¼_DELETE','EXECUTE' From Dual Union All
	Select 'Zl_����Σ��ֵ��¼_����','EXECUTE' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

--115026:������,2017-12-04,����Σ��ֵ
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��,ȱʡֵ)
Select 100,1284,A.* From (
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0 Union All
Select '����',-NULL,NULL,1 From Dual Union All
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select 100,1284,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All
Select '����ҽ����¼','SELECT' From Dual Union All
Select '������ҳ','SELECT' From Dual Union All
Select '���ű�','SELECT' From Dual Union All
Select '��Ժ����','SELECT' From Dual Union All
Select '������Ƭ','SELECT' From Dual Union All
Select '������ĿĿ¼','SELECT' From Dual Union All 
Select '������Ϣ','SELECT' From Dual Union All
Select '���˹Һż�¼','SELECT' From Dual Union All
Select '����Σ��ֵ��¼','SELECT' From Dual Union All
Select '����Σ��ֵҽ��','SELECT' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

--115026:������,2017-12-04,����Σ��ֵ
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��,ȱʡֵ)
Select &n_System,1261,A.* From (
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0 Union All 
    Select 'Σ��ֵ����',17,'�и�Ȩ��ʱ��סԺҽ��վ�Ŵ�����Σ��ֵ��¼��',1 From Dual Union All 
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0) A;

Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��,ȱʡֵ)
Select &n_System,1260,A.* From (
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0 Union All 
    Select 'Σ��ֵ����',23,'�и�Ȩ��ʱ������ҽ��վ�Ŵ�����Σ��ֵ��¼��',1 From Dual Union All 
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1260,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All
    Select '����Σ��ֵ��¼_ID','SELECT' From Dual Union All 
    Select '����Σ��ֵ��¼','SELECT' From Dual Union All 
    Select '����Σ��ֵҽ��','SELECT' From Dual Union All  
    Select '����Σ��ֵ����','SELECT' From Dual Union All   
    Select 'Zl_����Σ��ֵ��¼_Insert','EXECUTE' From Dual Union All 
    Select 'ZL_����Σ��ֵ��¼_UPDATE','EXECUTE' From Dual Union All
    Select 'Zl_����Σ��ֵ��¼_DELETE','EXECUTE' From Dual Union All
    Select 'Zl_����Σ��ֵ��¼_����','EXECUTE' From Dual Union All
    Select 'Zl_����Σ��ֵҽ��_Update','EXECUTE' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1261,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All
    Select '����Σ��ֵ��¼_ID','SELECT' From Dual Union All 
    Select '����Σ��ֵ��¼','SELECT' From Dual Union All 
    Select '����Σ��ֵҽ��','SELECT' From Dual Union All  
    Select '����Σ��ֵ����','SELECT' From Dual Union All   
    Select 'Zl_����Σ��ֵ��¼_Insert','EXECUTE' From Dual Union All 
    Select 'ZL_����Σ��ֵ��¼_UPDATE','EXECUTE' From Dual Union All
    Select 'Zl_����Σ��ֵ��¼_DELETE','EXECUTE' From Dual Union All
    Select 'Zl_����Σ��ֵ��¼_����','EXECUTE' From Dual Union All
    Select 'Zl_����Σ��ֵҽ��_Update','EXECUTE' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1252,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All 
    Select '����Σ��ֵҽ��','SELECT' From Dual Union All  
    Select 'Zl_����Σ��ֵҽ��_Update','EXECUTE' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1253,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All 
    Select '����Σ��ֵҽ��','SELECT' From Dual Union All  
    Select 'Zl_����Σ��ֵҽ��_Update','EXECUTE' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1254,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All 
    Select '����Σ��ֵҽ��','SELECT' From Dual Union All  
    Select 'Zl_����Σ��ֵҽ��_Update','EXECUTE' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

--117527:Ƚ����,2017-12-01,Ԥ�������ʹ�����ѿ�֧�����������δ�˻����ѿ�
Delete From zlProgPrivs
Where ϵͳ = &n_System And ��� = 1103 And ���� = 'Ԥ���˿�' And Upper(����) = 'ZL_���˿������¼_STRIKE';

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1103, 'Ԥ���˿�', User, 'ZL_���˿������¼_�˿�', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1103 And ���� = 'Ԥ���˿�' And Upper(����) = 'ZL_���˿������¼_�˿�');

--111635:���Ʊ�,2017-07-14,XML�Զ������뵥
Insert Into Zlprogprivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1253, '����', User, 'Zl_�Զ������뵥�ļ�_Edit', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From Zlprogprivs
         Where ϵͳ = &n_System And ��� = 1253 And ���� = '����' And Upper(����) = Upper('Zl_�Զ������뵥�ļ�_Edit'));

Insert Into Zlprogprivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1253, '����', User, 'Zl_ҽ�����뵥�ļ�_Edit', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From Zlprogprivs
         Where ϵͳ = &n_System And ��� = 1253 And ���� = '����' And Upper(����) = Upper('Zl_ҽ�����뵥�ļ�_Edit'));

Insert Into Zlprogprivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1252, '����', User, 'Zl_�Զ������뵥�ļ�_Edit', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From Zlprogprivs
         Where ϵͳ = &n_System And ��� = 1252 And ���� = '����' And Upper(����) = Upper('Zl_�Զ������뵥�ļ�_Edit'));

Insert Into Zlprogprivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1252, '����', User, 'Zl_ҽ�����뵥�ļ�_Edit', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From Zlprogprivs
         Where ϵͳ = &n_System And ��� = 1252 And ���� = '����' And Upper(����) = Upper('Zl_ҽ�����뵥�ļ�_Edit'));

Insert Into Zlprogprivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1252, '����', User, '������������', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From Zlprogprivs
         Where ϵͳ = &n_System And ��� = 1252 And ���� = '����' And Upper(����) = Upper('������������'));

Insert Into Zlprogprivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1253, '����', User, '������������', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From Zlprogprivs
         Where ϵͳ = &n_System And ��� = 1253 And ���� = '����' And Upper(����) = Upper('������������'));

Insert Into Zlprogprivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1252, '����', User, 'ҽ�����뵥�ļ�', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From Zlprogprivs
         Where ϵͳ = &n_System And ��� = 1252 And ���� = '����' And Upper(����) = Upper('ҽ�����뵥�ļ�'));

Insert Into Zlprogprivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1253, '����', User, 'ҽ�����뵥�ļ�', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From Zlprogprivs
         Where ϵͳ = &n_System And ��� = 1253 And ���� = '����' And Upper(����) = Upper('ҽ�����뵥�ļ�'));    

Insert Into Zlprogprivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1252, '����', User, '�Զ������뵥�ļ�', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From Zlprogprivs
         Where ϵͳ = &n_System And ��� = 1252 And ���� = '����' And Upper(����) = Upper('�Զ������뵥�ļ�'));

Insert Into Zlprogprivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1253, '����', User, '�Զ������뵥�ļ�', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From Zlprogprivs
         Where ϵͳ = &n_System And ��� = 1253 And ���� = '����' And Upper(����) = Upper('�Զ������뵥�ļ�'));   


--116034:���Ʊ�,2017-11-03,·�����ɲ�������������
Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1256, '����', User, 'Zl_Lob_ReadForPath', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1256 And ���� = '����' And Upper(����) = Upper('Zl_Lob_ReadForPath'));

--112953:���Ʊ�,2017-09-11,ҩƷ˵����֪ʶ��
Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1252, '����', User, 'Zl_Drugexplain_Readlob', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1252 And ���� = '����' And Upper(����) = Upper('Zl_Drugexplain_Readlob'));
Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1253, '����', User, 'Zl_Drugexplain_Readlob', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1253 And ���� = '����' And Upper(����) = Upper('Zl_Drugexplain_Readlob'));


--114434:��ΰ��,2017-11-17,�������׺�����ҩ�������
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select &n_System,1252,'����',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All 
    Select '����ҽ����¼_�������','SELECT' From Dual Union All 
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--115026:������,2017-12-04,����Σ��ֵ
--����ZL1_INSIDE_1254_20/Σ��ֵ��¼��
Insert Into zlReports(ID,���,����,˵��,����,��ӡ��,��ֽ,Ʊ��,��ӡ��ʽ,ϵͳ,����ID,����,�޸�ʱ��,����ʱ��,��ֹ��ʼʱ��,��ֹ����ʱ��) Values(zlReports_ID.NextVal,'ZL1_INSIDE_1254_20','Σ��ֵ��¼��','Σ��ֵ��¼��','Yn2t*l~v}1;F~et9C<AD',Null,15,0,0,100,Null,Null,Sysdate,Sysdate,To_Date('2017-11-01 00:00:00','YYYY-MM-DD HH24:MI:SS'),To_Date('2017-11-01 00:00:00','YYYY-MM-DD HH24:MI:SS'));
Insert Into zlRPTFmts(����ID,���,˵��,W,H,ֽ��,ֽ��,��ֽ̬��,ͼ��) Values(zlReports_ID.CurrVal,1,'Σ��ֵ��¼��',11904,16832,9,1,0,0);
Insert Into zlRPTDatas(ID,����ID,����,�ֶ�,����,����,˵��) Values(zlRPTDatas_ID.NextVal,zlReports_ID.CurrVal,'����Σ��ֵ','���,202|����,202|�Ա�,202|����,202|����,202|�����,202|���,139|������,202|����,202|Σ��ֵ����,202|����ʱ��,202|�������,202|������,202|�������,202|�Ƿ���Σ��ֵ,202|ȷ��ʱ��,202|ȷ�Ͽ���,202|ȷ����,202',User||'.����Σ��ֵ��¼,'||User||'.���ű�,'||User||'.������ҳ,'||User||'.���˹Һż�¼,'||User||'.����ҽ����¼',0,Null);
Insert Into zlRPTSQLs(ԴID,�к�,����)  Select zlRPTDatas_ID.CurrVal,a.* From (
  Select 1,'Select decode(h.�������,''D'',''�����'',''������'') as ���,a.����, a.�Ա�, a.����,' From Dual Union All
  Select 2,'decode(a.�Һŵ�,null,g.����,f.����) as ����,decode(a.�Һŵ�,null,''סԺ��'',''�����'') as �����,decode(a.�Һŵ�,null,d.סԺ��,e.�����) as ���' From Dual Union All
  Select 3,',decode(a.�Һŵ�,null,''�� ��'',''�� ��'') as ������,decode(a.�Һŵ�,null,d.��Ժ����,decode(e.����,1,''��'',''��'')) as ����,a.Σ��ֵ����,' From Dual Union All
  Select 4,'To_Char(a.����ʱ��, ''yyyy-mm-dd hh24:mi'') as ����ʱ��,b.���� as �������,a.������,a.�������, ' From Dual Union All
  Select 5,'decode(a.״̬,1,''   ��   �� '', decode(a.�Ƿ�Σ��ֵ,1,''   �ǡ� ��'',''   �� ���'')) as �Ƿ���Σ��ֵ,' From Dual Union All
  Select 6,'To_Char(a.ȷ��ʱ��, ''yyyy-mm-dd hh24:mi'') as ȷ��ʱ��,c.���� as ȷ�Ͽ���, a.ȷ���� ' From Dual Union All
  Select 7,'From ����Σ��ֵ��¼ A,���ű� b,���ű� c,������ҳ d,���˹Һż�¼ e,���ű� f,���ű� g,����ҽ����¼ h' From Dual Union All
  Select 8,'Where a.�������id=b.id and a.ȷ�Ͽ���id=c.id(+) and a.����id=d.����id(+) and a.��ҳid=d.��ҳid(+)' From Dual Union All
  Select 9,'and a.�Һŵ�=e.no(+) and e.ִ�в���id=f.id(+) and d.��Ժ����id=g.id(+) and a.ҽ��id=h.id(+) and a.Id = [0]' From Dual) a;
Insert Into zlRPTPars(ԴID,����,���,����,����,ȱʡֵ,��ʽ,ֵ�б�,����SQL,��ϸSQL,�����ֶ�,��ϸ�ֶ�,����,����) Values(zlRPTDatas_ID.CurrVal,Null,0,'��¼ID',1,'22',0,Null,Null,Null,Null,Null,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����14',1,Null,0,Null,0,Null,Null,3970,5565,210,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����10',1,Null,0,Null,0,Null,Null,2835,5565,210,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����11',1,Null,0,Null,0,Null,Null,3215,5565,210,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����13',1,Null,0,Null,0,Null,Null,4335,5565,210,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����1',1,Null,0,Null,0,Null,Null,4485,1230,2025,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����15',1,Null,0,Null,0,Null,Null,4705,5565,210,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����16',1,Null,0,Null,0,Null,Null,5070,5565,210,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����17',1,Null,0,Null,0,Null,Null,5460,5565,210,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����38',1,Null,0,Null,0,Null,Null,5465,2260,1575,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����35',1,Null,0,Null,0,Null,Null,5465,2905,1575,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����43',1,Null,0,Null,0,Null,Null,5675,8460,1575,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����12',1,Null,0,Null,0,Null,Null,3580,5565,210,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����40',1,Null,0,Null,0,Null,Null,5720,5390,1575,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����18',1,Null,0,Null,0,Null,Null,5825,5565,210,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����22',1,Null,0,Null,0,Null,Null,6175,5565,210,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����21',1,Null,0,Null,0,Null,Null,6540,5565,210,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����23',1,Null,0,Null,0,Null,Null,6910,5565,210,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����24',1,Null,0,Null,0,Null,Null,7275,5565,210,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����25',1,Null,0,Null,0,Null,Null,7665,5565,210,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����26',1,Null,0,Null,0,Null,Null,8030,5565,210,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����37',1,Null,0,Null,0,Null,Null,8265,2260,1575,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����45',1,Null,0,Null,0,Null,Null,8265,2905,1575,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����42',1,Null,0,Null,0,Null,Null,8265,5390,1575,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����3',1,Null,0,Null,0,Null,Null,8265,8460,1575,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����27',1,Null,0,Null,0,Null,Null,8410,5565,210,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����28',1,Null,0,Null,0,Null,Null,8775,5565,210,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����30',1,Null,0,Null,0,Null,Null,9165,5565,210,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����29',1,Null,0,Null,0,Null,Null,9530,5565,210,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����31',1,Null,0,Null,0,Null,Null,9900,5565,210,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����32',1,Null,0,Null,0,Null,Null,10265,5565,210,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����4',1,Null,0,Null,0,Null,Null,590,5565,210,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����6',1,Null,0,Null,0,Null,Null,980,5565,210,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����5',1,Null,0,Null,0,Null,Null,1345,5565,210,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����39',1,Null,0,Null,0,Null,Null,1675,2260,1575,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����36',1,Null,0,Null,0,Null,Null,1675,2905,1575,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����41',1,Null,0,Null,0,Null,Null,1675,5390,1905,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����44',1,Null,0,Null,0,Null,Null,1675,8460,1980,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����7',1,Null,0,Null,0,Null,Null,1715,5565,210,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����8',1,Null,0,Null,0,Null,Null,2080,5565,210,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����9',1,Null,0,Null,0,Null,Null,2470,5565,210,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ1',2,Null,0,Null,0,'Σ��ֵ��¼��',Null,4500,885,1980,330,0,0,1,'����',16,1,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ5',2,Null,0,Null,0,'��  ��',Null,4770,2010,630,225,0,0,1,'����',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����33',1,Null,0,Null,0,Null,Null,10655,5565,210,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ37',2,Null,0,Null,0,'[����Σ��ֵ.�����]',Null,4770,2655,1995,225,0,0,1,'����',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ29',2,Null,0,Null,0,'������� [����Σ��ֵ.�������]',Null,4770,5115,3150,225,0,0,1,'����',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ24',2,Null,0,Null,0,'ȷ�Ͽ��� [����Σ��ֵ.ȷ�Ͽ���]',Null,4770,8205,1935,210,0,0,1,'����',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����34',1,Null,0,Null,0,Null,Null,11020,5565,210,0,0,0,1,'����',0,0,0,0,0,0,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ26',2,Null,0,Null,0,'�Ƿ���Σ��ֵ [����Σ��ֵ.�Ƿ���Σ��ֵ]',Null,315,7755,3990,225,0,0,1,'����',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ17',2,Null,0,Null,0,'Σ��ֵ����',Null,525,3195,1125,225,0,0,1,'����',10.5,1,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ25',2,Null,0,Null,0,'ȷ��ʱ�� [����Σ��ֵ.ȷ��ʱ��]',Null,690,8205,3150,225,0,0,1,'����',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ18',2,Null,0,Null,0,'�������',Null,750,5970,900,225,0,0,1,'����',10.5,1,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ28',2,Null,0,Null,0,'����ʱ�� [����Σ��ֵ.����ʱ��]',Null,810,5115,3150,225,0,0,1,'����',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ2',2,Null,0,Null,0,'��  ��',Null,1020,2010,630,225,0,0,1,'����',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ8',2,Null,0,Null,0,'��  ��',Null,1020,2655,630,225,0,0,1,'����',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ33',2,Null,0,Null,0,'[����Σ��ֵ.����]',Null,1710,2010,1785,225,0,0,1,'����',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ36',2,Null,0,Null,0,'[����Σ��ֵ.����]',Null,1710,2655,1785,225,0,0,1,'����',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ27',2,Null,0,Null,0,'[����Σ��ֵ.�������]',Null,1725,6015,2205,225,0,0,1,'����',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ31',2,Null,0,Null,0,'[����Σ��ֵ.Σ��ֵ����]',Null,1755,3225,2415,225,0,0,1,'����',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ34',2,Null,0,Null,0,'[����Σ��ֵ.�Ա�]',Null,5475,2010,1785,225,0,0,1,'����',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ38',2,Null,0,Null,0,'[����Σ��ֵ.���]',Null,5475,2655,1785,225,0,2,1,'����',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ23',2,Null,0,Null,0,'ȷ���� [����Σ��ֵ.ȷ����]',Null,7560,8190,2280,210,0,0,1,'����',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ35',2,Null,0,Null,0,'[����Σ��ֵ.����]',Null,8310,2010,1785,225,0,0,1,'����',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ40',2,Null,0,Null,0,'[����Σ��ֵ.����]',Null,8310,2655,1785,225,0,0,1,'����',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ30',2,Null,0,Null,0,'������ [����Σ��ֵ.������]',Null,7560,5115,2730,225,0,0,1,'����',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����2',10,Null,0,Null,0,'����1',Null,1715,5975,8115,1440,0,0,0,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'����1',10,Null,0,Null,0,'����1',Null,1740,3195,8115,1440,0,0,0,'����',9,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ32',2,Null,0,Null,0,'[����Σ��ֵ.���]',Null,7560,1470,1785,225,0,0,1,'����',10.5,0,0,0,255,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ6',2,Null,0,Null,0,'��  ��',Null,7560,2010,630,225,0,0,1,'����',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);
Insert Into zlRPTItems(ID,����ID,��ʽ��,����,����,�ϼ�ID,���,����,����,����,��ͷ,X,Y,W,H,�и�,����,�Ե�,����,�ֺ�,����,б��,����,ǰ��,����,�߿�,����,��ʽ,����,����,����,ϵͳ,�������,�������,���Ҽ��,���¼��,ԴID,Դ�к�) Values(zlRPTItems_ID.NextVal,zlReports_ID.CurrVal,1,'��ǩ39',2,Null,0,Null,0,'[����Σ��ֵ.������]',Null,7560,2655,1995,225,0,0,1,'����',10.5,0,0,0,0,16777215,0,Null,Null,Null,1,0,0,0,0,0,0,Null,0);


--����ZL1_INSIDE_1254_20/Σ��ֵ��¼��
Insert into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1252,'Σ��ֵ��¼��','Σ��ֵ��¼��');
Insert into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1253,'Σ��ֵ��¼��','Σ��ֵ��¼��');
Insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
  Select 100,1252,'Σ��ֵ��¼��',User,'����Σ��ֵ��¼','SELECT' From Dual Union All
  Select 100,1252,'Σ��ֵ��¼��',User,'���˹Һż�¼','SELECT' From Dual Union All
  Select 100,1252,'Σ��ֵ��¼��',User,'����ҽ����¼','SELECT' From Dual Union All
  Select 100,1252,'Σ��ֵ��¼��',User,'���ű�','SELECT' From Dual Union All 
  Select 100,1252,'Σ��ֵ��¼��',User,'������ҳ','SELECT' From Dual Union All  
  Select 100,1253,'Σ��ֵ��¼��',User,'����Σ��ֵ��¼','SELECT' From Dual Union All
  Select 100,1253,'Σ��ֵ��¼��',User,'���˹Һż�¼','SELECT' From Dual Union All 
  Select 100,1253,'Σ��ֵ��¼��',User,'�������ҽ��','SELECT' From Dual Union All
  Select 100,1253,'Σ��ֵ��¼��',User,'���ű�','SELECT' From Dual Union All 
  Select 100,1253,'Σ��ֵ��¼��',User,'������ҳ','SELECT' From Dual;


-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--118566:��С��,2017-12-22,�ϰ�LIS������Ŀ�����޸�վ��Ϊ��
Create Or Replace Procedure Zl_������Ŀ_Edit
(
  �༭����_In   In Number, --1-���ӣ�2-�޸ģ�3-ɾ��
  Id_In         In ������ĿĿ¼.Id%Type,
  ���Ʒ���id_In In ������ĿĿ¼.����id%Type := Null,
  ��������_In   In ������ĿĿ¼.��������%Type := Null,
  ����_In       In ������ĿĿ¼.����%Type := Null,
  ����_In       In ������ĿĿ¼.����%Type := Null,
  ����ƴ��_In   In ������Ŀ����.����%Type := Null,
  �������_In   In ������Ŀ����.����%Type := Null,
  ����_In       In ������Ŀ����.����%Type := Null,
  Ӣ����_In     In ������Ŀ.��д%Type := Null,
  ���㵥λ_In   In ������ĿĿ¼.���㵥λ%Type := Null,
  �걾��λ_In   In ������ĿĿ¼.�걾��λ%Type := Null,
  �����Ա�_In   In ������ĿĿ¼.�����Ա�%Type := Null,
  ����Ӧ��_In   In ������ĿĿ¼.����Ӧ��%Type := Null,
  �����Ŀ_In   In ������ĿĿ¼.�����Ŀ%Type := Null,
  �������_In   In ������Ŀ.�������%Type := Null,
  ���鷽��_In   In ������Ŀ.���鷽��%Type := Null,
  
  ��Ŀ���_In In ������Ŀ.��Ŀ���%Type := Null,
  �������_In In ������Ŀ.�������%Type := Null,
  �����Χ_In In ������Ŀ.�����Χ%Type := Null,
  Ĭ��ֵ_In   In ������Ŀ.Ĭ��ֵ%Type := Null,
  ���㹫ʽ_In In ������Ŀ.���㹫ʽ%Type := Null,
  ȡֵ����_In In ������Ŀ.ȡֵ����%Type := Null,
  ��˽��Ŀ_In In ������Ŀ.��˽��Ŀ%Type := Null,
  ��ο�_In   In ������Ŀ.��ο�%Type := Null,
  
  ���Թ�ʽ_In   In ������Ŀ.���Թ�ʽ%Type := Null,
  �����Թ�ʽ_In In ������Ŀ.�����Թ�ʽ%Type := Null,
  Cutoff��ʽ_In In ������Ŀ.Cutoff��ʽ%Type := Null
  
) Is
  v_�������   ������ĿĿ¼.�������%Type;
  v_�����Ŀ   ������ĿĿ¼.�����Ŀ%Type;
  v_ִ�п���   ������ĿĿ¼.ִ�п���%Type;
  v_������Ŀid ���鱨����Ŀ.������Ŀid%Type := 0;
  v_վ��       ������ĿĿ¼.վ��%Type;

  Function Get_������Ŀid(������Ŀid_In In ������ĿĿ¼.Id%Type) Return Number Is
    v_������Ŀid ����������Ŀ.Id%Type;
  Begin
    Select ������Ŀid Into v_������Ŀid From ���鱨����Ŀ Where ������Ŀid = ������Ŀid_In And ������Ŀid Is Not Null;
    Return v_������Ŀid;
  Exception
    When Others Then
      Return Null;
  End Get_������Ŀid;

Begin
  If �༭����_In = 1 Then
    Zl_������Ŀ_Insert('C', ���Ʒ���id_In, Id_In, ����_In, ����_In, ����ƴ��_In, �������_In, ����_In, ����ƴ��_In, �������_In, ��������_In, 1, ����Ӧ��_In,
                   3, ���㵥λ_In, �����Ա�_In, 0, 3, �����Ŀ_In, �걾��λ_In, Null, 4, Null, Null, Null, Null, 0);
    --Update ������ĿĿ¼ Set ������� = �������_In Where ID = Id_In;
    If �����Ŀ_In = 0 Then
      Select ����������Ŀ_Id.Nextval Into v_������Ŀid From Dual;
    End If;
  Elsif �༭����_In = 2 Then
    Select �������, Nvl(�����Ŀ, 0), ִ�п���, վ��
    Into v_�������, v_�����Ŀ, v_ִ�п���, v_վ��
    From ������ĿĿ¼
    Where ID = Id_In;
    Zl_������Ŀ_Update('C', ���Ʒ���id_In, Id_In, ����_In, ����_In, ����ƴ��_In, �������_In, ����_In, ����ƴ��_In, �������_In, ��������_In, 1, ����Ӧ��_In,
                   3, ���㵥λ_In, �����Ա�_In, 0, v_�������, �����Ŀ_In, �걾��λ_In, Null, v_ִ�п���, Null, Null, Null, Null, 1, 0, Null, 0,
                   0, 0, v_վ��);
    --Update ������ĿĿ¼ Set ������� = �������_In Where ID = Id_In;
    If v_�����Ŀ = 0 Then
      v_������Ŀid := Get_������Ŀid(Id_In);
      If �����Ŀ_In = 1 Then
        Delete ���鱨����Ŀ Where ������Ŀid = Id_In And ϸ��id Is Null;
        Delete ����������Ŀ Where ID = v_������Ŀid;
      End If;
    Else
      If �����Ŀ_In = 0 Then
        Delete ���鱨����Ŀ Where ������Ŀid = Id_In;
        Select ����������Ŀ_Id.Nextval Into v_������Ŀid From Dual;
      End If;
    End If;
    -- ���ϰ�������ӵ���Ŀ,����û�б�����Ŀid 2007-07-13
    If Nvl(v_������Ŀid, 0) = 0 Then
      Select ����������Ŀ_Id.Nextval Into v_������Ŀid From Dual;
    End If;
  Elsif �༭����_In = 3 Then
    Select Nvl(�����Ŀ, 0) Into v_�����Ŀ From ������ĿĿ¼ Where ID = Id_In;
    If v_�����Ŀ = 0 Then
      v_������Ŀid := Get_������Ŀid(Id_In);
      Delete ���鱨����Ŀ Where ������Ŀid = Id_In;
      Delete ����������Ŀ Where ID = v_������Ŀid;
    End If;
    Delete ������ĿĿ¼ Where ID = Id_In;
    Return;
  End If;

  If �����Ŀ_In = 0 Then
    Update ����������Ŀ
    Set ���� = ����_In, ������ = ����_In, Ӣ���� = Ӣ����_In, �滻�� = 0, ���� = Decode(�������_In, 1, 0, 2, 1, 3, 3),
        ���� = Decode(�������_In, 1, 10, 2, 100, 3, 10), С�� = Decode(�������_In, 1, 3, 2, 0, 3, 0), ��λ = ���㵥λ_In, ��ʾ�� = 0,
        �Ա��� = �����Ա�_In
    Where ID = v_������Ŀid;
    If Sql%RowCount = 0 Then
      Insert Into ����������Ŀ
        (ID, ����, ������, Ӣ����, �滻��, ����, ����, С��, ��λ, ��ʾ��, �Ա���)
      Values
        (v_������Ŀid, ����_In, ����_In, Ӣ����_In, 0, Decode(�������_In, 1, 0, 2, 1, 3, 3), Decode(�������_In, 1, 10, 2, 100, 3, 10),
         Decode(�������_In, 1, 3, 2, 0, 3, 0), ���㵥λ_In, 0, �����Ա�_In);
      Insert Into ���鱨����Ŀ
        (ID, ������Ŀid, ������Ŀid, ����걾)
      Values
        (���鱨����Ŀ_Id.Nextval, Id_In, v_������Ŀid, �걾��λ_In);
    Else
      Update ���鱨����Ŀ Set ����걾 = �걾��λ_In Where ������Ŀid = Id_In And ������Ŀid = v_������Ŀid;
    End If;
  
    Update ������Ŀ
    Set ��д = Ӣ����_In, ��λ = ���㵥λ_In, ��Ŀ��� = ��Ŀ���_In, ������� = �������_In, �����Χ = �����Χ_In, Ĭ��ֵ = Ĭ��ֵ_In, ���㹫ʽ = ���㹫ʽ_In,
        ȡֵ���� = ȡֵ����_In, ��˽��Ŀ = ��˽��Ŀ_In, ���Թ�ʽ = ���Թ�ʽ_In, �����Թ�ʽ = �����Թ�ʽ_In, Cutoff��ʽ = Cutoff��ʽ_In, ������� = �������_In,
        ���鷽�� = ���鷽��_In, ��ο� = ��ο�_In
    Where ������Ŀid = v_������Ŀid;
    If Sql%RowCount = 0 Then
      Insert Into ������Ŀ
        (������Ŀid, ��д, ��λ, ��Ŀ���, �������, �����Χ, Ĭ��ֵ, ���㹫ʽ, ȡֵ����, ��˽��Ŀ, ���Թ�ʽ, �����Թ�ʽ, Cutoff��ʽ, �������, ���鷽��, ��ο�)
      Values
        (v_������Ŀid, Ӣ����_In, ���㵥λ_In, ��Ŀ���_In, �������_In, �����Χ_In, Ĭ��ֵ_In, ���㹫ʽ_In, ȡֵ����_In, ��˽��Ŀ_In, ���Թ�ʽ_In, �����Թ�ʽ_In,
         Cutoff��ʽ_In, �������_In, ���鷽��_In, ��ο�_In);
    End If;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_������Ŀ_Edit;
/

--118738:����,2017-12-20,�޸ķ�Ʊ��Ϣ����ϴι�Ӧ�̴���
Create Or Replace Procedure Zl_�����⹺��Ʊ��Ϣ_Update
(
  No_In         In ҩƷ�շ���¼.No%Type := Null,
  ��¼״̬_In   In ҩƷ�շ���¼.��¼״̬%Type := Null,
  ���_In       In ҩƷ�շ���¼.���%Type := Null,
  ��Ʊ��_In     In Ӧ����¼.��Ʊ��%Type := Null,
  ��Ʊ����_In   In Ӧ����¼.��Ʊ����%Type := Null,
  ��Ʊ���_In   In Ӧ����¼.��Ʊ���%Type := Null,
  ������λid_In In Ӧ����¼.��λid%Type := 0,
  ��Ʊ����_In   In Ӧ����¼.��Ʊ����%Type := Null
) Is
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  v_No         Ӧ����¼.No%Type;
  n_Ӧ��id     Ӧ����¼.Id%Type;
  n_�շ�id     Ӧ����¼.�շ�id%Type;
  n_�������   Ӧ����¼.�������%Type;
  n_��Ʊ���   Ӧ����¼.��Ʊ���%Type; --�ɷ�Ʊ���
  n_������λid Ӧ����¼.��λid%Type;
  n_Dec        Number;
  n_ʣ������   Ӧ����¼.��Ʊ���%Type;
Begin
  --���С��λ��
  Select Nvl(����, 2) Into n_Dec From ҩƷ���ľ��� Where ���� = 0 And ��� = 2 And ���� = 4 And ��λ = 5;

  --ȡ�Ƿ񸶿�ܶ�
  Begin
    Select Max(�������), Sum(Nvl(��Ʊ���, 0))
    Into n_�������, n_��Ʊ���
    From Ӧ����¼
    Where �շ�id In (Select ID From ҩƷ�շ���¼ Where NO = No_In And ��� = ���_In And ���� = 15) And ϵͳ��ʶ = 5 And ��¼���� = -1;
  Exception
    When Others Then
      n_��Ʊ��� := 0;
  End;

  n_������� := Nvl(n_�������, 0);

  If n_������� <> 0 Then
    v_Err_Msg := '[ZLSOFT]�õ����Ѿ������˿�������޸ķ�Ʊ��Ϣ[ZLSOFT]';
    Raise Err_Item;
  End If;

  If ��Ʊ���_In > n_��Ʊ��� And n_��Ʊ��� <> 0 Then
    v_Err_Msg := '[ZLSOFT]��Ʊ����С�ڼƻ�������[ZLSOFT]';
    Raise Err_Item;
  End If;
  n_��Ʊ��� := Nvl(n_��Ʊ���, 0);

  --�ж��Ƿ������ļ�¼
  If ��¼״̬_In <> 1 Then
    Begin
      Select Sum(Nvl(��Ʊ���, 0))
      Into n_��Ʊ���
      From Ӧ����¼
      Where �շ�id In (Select ID From ҩƷ�շ���¼ Where NO = No_In And ��� = ���_In And ���� = 15) And ϵͳ��ʶ = 5 And ��¼���� = 0;
    Exception
      When Others Then
        n_��Ʊ��� := 0;
    End;
    n_��Ʊ��� := Nvl(n_��Ʊ���, 0);
    If Nvl(��Ʊ��_In, ' ') = ' ' And ��Ʊ���_In <> 0 Then
      v_Err_Msg := '[ZLSOFT]���ܶԳ����򱻳�����¼�ķ�Ʊ�Ÿ�Ϊ��,���ܱ��棡[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    Select ��ҩ��λid, Sum(Nvl(ʵ������, 0))
    Into n_������λid, n_ʣ������
    From ҩƷ�շ���¼
    Where ���� = 15 And NO = No_In And ��� = ���_In
    Group By ��ҩ��λid;
  
    --������صķ�Ʊ��Ϣ,ֻ���ķ�Ʊ�ţ���Ʊ����
    For v_�շ� In (Select a.Id, a.�ⷿid, a.No, a.��¼״̬, a.���۽��, b.����, b.���, b.����, a.����, b.���㵥λ, a.ʵ������, a.�ɱ���, a.�ɱ����, a.������,
                        a.��������, a.�����, a.�������, a.ժҪ, a.ҩƷid, a.���, a.��ҩ��λid
                 From ҩƷ�շ���¼ A, �շ���ĿĿ¼ B
                 Where a.���� = 15 And a.No = No_In And a.��� = ���_In And a.ҩƷid = b.Id
                 Order By a.Id) Loop
      Update Ӧ����¼
      Set ��Ʊ�� = ��Ʊ��_In, ��Ʊ���� = ��Ʊ����_In, ��Ʊ���� = ��Ʊ����_In, ��Ʊ��� = Round((v_�շ�.ʵ������ / n_ʣ������) * ��Ʊ���_In, n_Dec),
          ��λid = ������λid_In, ��Ʊ�޸�ʱ�� = Sysdate
      Where �շ�id = v_�շ�.Id And ϵͳ��ʶ = 5 And ��¼���� = 0;
    
      If Sql%RowCount = 0 Then
        If ��Ʊ��_In Is Not Null Then
          --����ǵ�һ����ϸ,�����Ӧ����¼��NO
          Begin
            Select NO
            Into v_No
            From Ӧ����¼
            Where ϵͳ��ʶ = 5 And ��¼���� = 0 And ��ⵥ�ݺ� = No_In And Rownum < 2;
          Exception
            When Others Then
              v_No := Nextno(67);
          End;
        
          Select Ӧ����¼_Id.Nextval Into n_Ӧ��id From Dual;
          Insert Into Ӧ����¼
            (ID, ��¼����, ��¼״̬, ��λid, NO, ϵͳ��ʶ, �շ�id, ��ⵥ�ݺ�, ���ݽ��, ��Ʊ��, ��Ʊ����, ��Ʊ���, Ʒ��, ���, ����, ����, ������λ, ����, �ɹ���, �ɹ����,
             ������, ��������, �����, �������, ժҪ, ��Ŀid, ���, �ⷿid, ��Ʊ����, ��Ʊ�޸�ʱ��)
          Values
            (n_Ӧ��id, 0, v_�շ�.��¼״̬, ������λid_In, v_No, 5, v_�շ�.Id, v_�շ�.No, v_�շ�.���۽��, ��Ʊ��_In, ��Ʊ����_In,
             Round((v_�շ�.ʵ������ / n_ʣ������) * ��Ʊ���_In, n_Dec), v_�շ�.����, v_�շ�.���, v_�շ�.����, v_�շ�.����, v_�շ�.���㵥λ, v_�շ�.ʵ������,
             v_�շ�.�ɱ���, v_�շ�.�ɱ����, v_�շ�.������, v_�շ�.��������, v_�շ�.�����, v_�շ�.�������, v_�շ�.ժҪ, v_�շ�.ҩƷid, v_�շ�.���, v_�շ�.�ⷿid,
             ��Ʊ����_In, Sysdate);
        End If;
      End If;
    End Loop;
  Else
    --δ�����ĵ���
    Select a.Id, Nvl(b.��Ʊ���, 0), a.��ҩ��λid
    Into n_�շ�id, n_��Ʊ���, n_������λid
    From ҩƷ�շ���¼ A, (Select * From Ӧ����¼ Where ϵͳ��ʶ = 5 And ��¼���� = 0 And ��¼״̬ = 1 And ������� Is Null) B
    Where a.Id = b.�շ�id(+) And a.No = No_In And a.���� = 15 And a.��¼״̬ = 1 And a.��� = ���_In;
  
    Update Ӧ����¼
    Set ��Ʊ�� = ��Ʊ��_In, ��Ʊ���� = ��Ʊ����_In, ��Ʊ���� = ��Ʊ����_In, ��Ʊ��� = ��Ʊ���_In, ��λid = ������λid_In, ��Ʊ�޸�ʱ�� = Sysdate
    Where �շ�id = n_�շ�id And ϵͳ��ʶ = 5 And ��¼״̬ = 1 And ��¼���� = 0;
  
    If Sql%RowCount = 0 Then
      If ��Ʊ��_In Is Not Null Or ��Ʊ����_In Is Not Null Then
        --����ǵ�һ����ϸ,�����Ӧ����¼��NO
        Begin
          Select NO
          Into v_No
          From Ӧ����¼
          Where ϵͳ��ʶ = 5 And ��¼���� = 0 And ��¼״̬ = 1 And ��ⵥ�ݺ� = No_In And Rownum < 2;
        Exception
          When Others Then
            v_No := Nextno(67);
        End;
      
        Select Ӧ����¼_Id.Nextval Into n_Ӧ��id From Dual;
      
        Insert Into Ӧ����¼
          (ID, ��¼����, ��¼״̬, ��Ŀid, ���, ��λid, NO, ϵͳ��ʶ, �շ�id, ��ⵥ�ݺ�, ���ݽ��, ��Ʊ��, ��Ʊ����, ��Ʊ���, Ʒ��, ���, ����, ����, ������λ, ����, �ɹ���,
           �ɹ����, ������, ��������, �����, �������, ժҪ, �ⷿid, ��Ʊ����, ��Ʊ�޸�ʱ��)
          Select n_Ӧ��id, 0, 1, a.ҩƷid, a.���, ������λid_In, v_No, 5, n_�շ�id, a.No, a.���۽��, ��Ʊ��_In, ��Ʊ����_In, ��Ʊ���_In, b.����,
                 b.���, b.����, a.����, b.���㵥λ, a.ʵ������, a.�ɱ���, a.�ɱ����, a.������, a.��������, a.�����, a.�������, a.ժҪ, a.�ⷿid, ��Ʊ����_In,
                 Sysdate
          From ҩƷ�շ���¼ A, �շ���ĿĿ¼ B
          Where a.���� = 15 And a.No = No_In And a.��� = ���_In And a.ҩƷid = b.Id;
      End If;
    End If;
  End If;

  Update Ӧ����� Set ��� = Nvl(���, 0) - n_��Ʊ��� Where ��λid = n_������λid And ���� = 1;
  If Sql%NotFound Then
    Insert Into Ӧ����� (��λid, ����, ���) Values (n_������λid, 1, -n_��Ʊ���);
  End If;
  Update Ӧ����� Set ��� = Nvl(���, 0) + Nvl(��Ʊ���_In, 0) Where ��λid = ������λid_In And ���� = 1;

  If Sql%NotFound Then
    Insert Into Ӧ����� (��λid, ����, ���) Values (������λid_In, 1, ��Ʊ���_In);
  End If;

  --����ҩƷ�շ���¼�еĹ�ҩ��λ
  Update ҩƷ�շ���¼ Set ��ҩ��λid = ������λid_In Where NO = No_In And ���� = 15 And ��� = ���_In;

  --����ҩƷ�������ϴι�Ӧ��
  Update ҩƷ���
  Set �ϴι�Ӧ��id = ������λid_In
  Where ���� = 1 And (�ⷿid, ҩƷid, ����) In (Select �ⷿid, ҩƷid, nvl(����,0) as ���� From ҩƷ�շ���¼ Where NO = No_In And ���� = 15);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When No_Data_Found Then
    Raise_Application_Error(-20101, '[ZLSOFT]�õ����Ѿ������˳������Ѿ������[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����⹺��Ʊ��Ϣ_Update;
/

--118738:����,2017-12-20,�޸ķ�Ʊ��Ϣ����ϴι�Ӧ�̴���
Create Or Replace Procedure Zl_ҩƷ�⹺��Ʊ��Ϣ_Update
(
  No_In       In ҩƷ�շ���¼.No%Type := Null,
  ���_In     In ҩƷ�շ���¼.���%Type,
  ��Ʊ��_In   In Ӧ����¼.��Ʊ��%Type := Null,
  ��Ʊ����_In In Ӧ����¼.��Ʊ����%Type := Null,
  ��Ʊ���_In In Ӧ����¼.��Ʊ���%Type := Null,
  ��ҩ��λ_In In Ӧ����¼.��λid%Type := 0,
  ������־_In Number, --1��δ���������޸ķ�Ʊ��Ϣ; 2�����ֳ��������޸ķ�Ʊ��Ϣ
  ��Ʊ����_In In Ӧ����¼.��Ʊ����%Type := Null
) Is
  Errinfor Varchar2(255);
  Erritem Exception;

  v_No         Ӧ����¼.No%Type;
  v_Ӧ��id     Ӧ����¼.Id%Type;
  v_�շ�id     Ӧ����¼.�շ�id%Type;
  v_�������   Ӧ����¼.�������%Type;
  v_��Ʊ���   Ӧ����¼.��Ʊ���%Type; --�ɷ�Ʊ���
  v_��ҩ��λid Ӧ����¼.��λid%Type;
  n_Dec        Number;
  n_ʣ������   ҩƷ�շ���¼.ʵ������%Type;

  Cursor c_ҩƷ��¼ Is
    Select a.Id, a.�ⷿid, a.No, a.��¼״̬, a.���۽��, b.����, b.���, b.����, a.����, b.���㵥λ, a.ʵ������, a.�ɱ���, a.�ɱ����, a.������, a.��������,
           a.�����, a.�������, a.ժҪ, a.ҩƷid, a.���, a.��ҩ��λid
    From ҩƷ�շ���¼ A, �շ���ĿĿ¼ B
    Where a.���� = 1 And a.No = No_In And a.��� = ���_In And a.ҩƷid = b.Id
    Order By a.Id;
Begin
  --���С��λ��
  Select Nvl(����, 2) Into n_Dec From ҩƷ���ľ��� Where ���� = 0 And ��� = 1 And ���� = 4 And ��λ = 5;

  --ȡ�Ƿ񸶿�ܶ�
  Begin
    Select Max(�������), Sum(Nvl(��Ʊ���, 0))
    Into v_�������, v_��Ʊ���
    From Ӧ����¼
    Where �շ�id In (Select ID From ҩƷ�շ���¼ Where NO = No_In And ��� = ���_In And ���� = 1) And ϵͳ��ʶ = 1 And ��¼���� = -1;
  Exception
    When Others Then
      v_��Ʊ��� := 0;
      Null;
  End;
  v_��Ʊ��� := Nvl(v_��Ʊ���, 0);
  v_������� := Nvl(v_�������, 0);
  If v_������� <> 0 Then
    Errinfor := '[ZLSOFT]�õ����Ѿ������˿�������޸ķ�Ʊ��Ϣ[ZLSOFT]';
    Raise Erritem;
  End If;
  If ��Ʊ���_In > v_��Ʊ��� And v_��Ʊ��� <> 0 Then
    Errinfor := '[ZLSOFT]��Ʊ���ܴ��ڼƻ�������[ZLSOFT]';
    Raise Erritem;
  End If;

  If ������־_In = 1 Then
    --δ��������
    Select a.Id, Nvl(b.��Ʊ���, 0), a.��ҩ��λid
    Into v_�շ�id, v_��Ʊ���, v_��ҩ��λid
    From ҩƷ�շ���¼ A, (Select * From Ӧ����¼ Where ϵͳ��ʶ = 1 And ��¼���� = 0 And ��¼״̬ = 1 And ������� Is Null) B
    Where a.Id = b.�շ�id(+) And a.No = No_In And a.���� = 1 And a.��¼״̬ = 1 And a.��� = ���_In;
  
    Update Ӧ����¼
    Set ��Ʊ�� = ��Ʊ��_In, ��Ʊ���� = ��Ʊ����_In, ��Ʊ���� = ��Ʊ����_In, ��Ʊ��� = ��Ʊ���_In, ��λid = ��ҩ��λ_In, ��Ʊ�޸�ʱ�� = Sysdate
    Where �շ�id = v_�շ�id And ϵͳ��ʶ = 1 And ��¼״̬ = 1 And ��¼���� = 0;
  
    If Sql%RowCount = 0 Then
      If ��Ʊ��_In Is Not Null Then
        --����ǵ�һ����ϸ,�����Ӧ����¼��NO
        Begin
          Select NO
          Into v_No
          From Ӧ����¼
          Where ϵͳ��ʶ = 1 And ��¼���� = 0 And ��¼״̬ = 1 And ��ⵥ�ݺ� = No_In And Rownum < 2;
        Exception
          When Others Then
            v_No := Nextno(67);
        End;
      
        Select Ӧ����¼_Id.Nextval Into v_Ӧ��id From Dual;
        Insert Into Ӧ����¼
          (ID, ��¼����, ��¼״̬, ��λid, NO, ϵͳ��ʶ, �շ�id, ��ⵥ�ݺ�, ���ݽ��, ��Ʊ��, ��Ʊ����, ��Ʊ���, Ʒ��, ���, ����, ����, ������λ, ����, �ɹ���, �ɹ����, ������,
           ��������, �����, �������, ժҪ, ��Ŀid, ���, �ⷿid, ��Ʊ�޸�ʱ��, ��Ʊ����)
          Select v_Ӧ��id, 0, 1, ��ҩ��λ_In, v_No, 1, v_�շ�id, a.No, a.���۽��, ��Ʊ��_In, ��Ʊ����_In, ��Ʊ���_In, b.����, b.���, b.����, a.����,
                 b.���㵥λ, a.ʵ������, a.�ɱ���, a.�ɱ����, a.������, a.��������, a.�����, a.�������, a.ժҪ, a.ҩƷid, a.���, a.�ⷿid, Sysdate,
                 ��Ʊ����_In
          From ҩƷ�շ���¼ A, �շ���ĿĿ¼ B
          Where a.���� = 1 And a.No = No_In And a.��� = ���_In And a.ҩƷid = b.Id;
      End If;
    End If;
  Else
    --����ԭ���ݵķ�Ʊ���
    Begin
      Select Sum(Nvl(��Ʊ���, 0))
      Into v_��Ʊ���
      From Ӧ����¼
      Where �շ�id In (Select ID From ҩƷ�շ���¼ Where NO = No_In And ��� = ���_In And ���� = 1) And ϵͳ��ʶ = 1 And ��¼���� = 0;
    Exception
      When Others Then
        v_��Ʊ��� := 0;
        Null;
    End;
  
    v_��Ʊ��� := Nvl(v_��Ʊ���, 0);
  
    --���ֳ������ݣ���������̯��Ʊ���
    Select ��ҩ��λid, Sum(ʵ������)
    Into v_��ҩ��λid, n_ʣ������
    From ҩƷ�շ���¼
    Where ���� = 1 And NO = No_In And ��� = ���_In
    Group By ��ҩ��λid;
  
    For v_ҩƷ��¼ In c_ҩƷ��¼ Loop
      Update Ӧ����¼
      Set ��Ʊ�� = ��Ʊ��_In, ��Ʊ���� = ��Ʊ����_In, ��Ʊ���� = ��Ʊ����_In, ��Ʊ��� = Round((v_ҩƷ��¼.ʵ������ / n_ʣ������) * ��Ʊ���_In, n_Dec),
          ��λid = ��ҩ��λ_In, ��Ʊ�޸�ʱ�� = Sysdate
      Where �շ�id = v_ҩƷ��¼.Id And ϵͳ��ʶ = 1 And ��¼���� = 0;
    
      If Sql%RowCount = 0 Then
        If ��Ʊ��_In Is Not Null Then
          --����ǵ�һ����ϸ,�����Ӧ����¼��NO
          Begin
            Select NO
            Into v_No
            From Ӧ����¼
            Where ϵͳ��ʶ = 1 And ��¼���� = 0 And ��ⵥ�ݺ� = No_In And Rownum < 2;
          Exception
            When Others Then
              v_No := Nextno(67);
          End;
        
          Select Ӧ����¼_Id.Nextval Into v_Ӧ��id From Dual;
          Insert Into Ӧ����¼
            (ID, ��¼����, ��¼״̬, ��λid, NO, ϵͳ��ʶ, �շ�id, ��ⵥ�ݺ�, ���ݽ��, ��Ʊ��, ��Ʊ����, ��Ʊ���, Ʒ��, ���, ����, ����, ������λ, ����, �ɹ���, �ɹ����,
             ������, ��������, �����, �������, ժҪ, ��Ŀid, ���, �ⷿid, ��Ʊ�޸�ʱ��, ��Ʊ����)
          Values
            (v_Ӧ��id, 0, v_ҩƷ��¼.��¼״̬, ��ҩ��λ_In, v_No, 1, v_ҩƷ��¼.Id, v_ҩƷ��¼.No, v_ҩƷ��¼.���۽��, ��Ʊ��_In, ��Ʊ����_In,
             Round((v_ҩƷ��¼.ʵ������ / n_ʣ������) * ��Ʊ���_In, n_Dec), v_ҩƷ��¼.����, v_ҩƷ��¼.���, v_ҩƷ��¼.����, v_ҩƷ��¼.����, v_ҩƷ��¼.���㵥λ,
             v_ҩƷ��¼.ʵ������, v_ҩƷ��¼.�ɱ���, v_ҩƷ��¼.�ɱ����, v_ҩƷ��¼.������, v_ҩƷ��¼.��������, v_ҩƷ��¼.�����, v_ҩƷ��¼.�������, v_ҩƷ��¼.ժҪ,
             v_ҩƷ��¼.ҩƷid, v_ҩƷ��¼.���, v_ҩƷ��¼.�ⷿid, Sysdate, ��Ʊ����_In);
        End If;
      End If;
    
    End Loop;
  End If;

  Update Ӧ����� Set ��� = Nvl(���, 0) - v_��Ʊ��� Where ��λid = v_��ҩ��λid And ���� = 1;
  If Sql%NotFound Then
    Insert Into Ӧ����� (��λid, ����, ���) Values (v_��ҩ��λid, 1, -v_��Ʊ���);
  End If;
  Update Ӧ����� Set ��� = Nvl(���, 0) + ��Ʊ���_In Where ��λid = ��ҩ��λ_In And ���� = 1;
  If Sql%NotFound Then
    Insert Into Ӧ����� (��λid, ����, ���) Values (��ҩ��λ_In, 1, ��Ʊ���_In);
  End If;

  --����ҩƷ�շ���¼�еĹ�ҩ��λ
  Update ҩƷ�շ���¼ Set ��ҩ��λid = ��ҩ��λ_In Where NO = No_In And ���� = 1 And ��� = ���_In;

  --����ҩƷ�������ϴι�Ӧ��
  Update ҩƷ���
  Set �ϴι�Ӧ��id = ��ҩ��λ_In
  Where (�ⷿid, ҩƷid, ����) In (Select �ⷿid, ҩƷid, nvl(����,0) as ���� From ҩƷ�շ���¼ Where NO = No_In And ���� = 1) And ���� = 1;
Exception
  When Erritem Then
    Raise_Application_Error(-20101, Errinfor);
  When No_Data_Found Then
    Raise_Application_Error(-20101, '[ZLSOFT]�õ����Ѿ������˳������Ѿ������[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�⹺��Ʊ��Ϣ_Update;
/

--106747:��С��,2017-12-19,ֱ�ӵǼǵı걾�ع�������״̬ʱɾ������ҽ����¼
Create Or Replace Procedure Zl_����걾��¼_תΪ����
(
  ҽ��id_In   In ����걾��¼.ҽ��id%Type,
  ɾ��Ժ��_In In Number := 0
) Is
  --=0ɾ������=1��ɾ������

  Cursor c_Sample Is
    Select Distinct a.Id As �걾id, Decode(b.ҽ��id, Null, a.ҽ��id, b.ҽ��id) As ҽ��id, a.��������, a.����id, a.������Դ
    From ������Ŀ�ֲ� B, ����걾��¼ A
    Where a.Id = b.�걾id(+) And a.ҽ��id = ҽ��id_In;

  Cursor c_Stuff(Vno Varchar2) Is
    Select Distinct s.Id, s.����, s.ʵ������, s.�ѷ�����
    From (Select a.Id, a.����, a.ʵ������, b.�ѷ�����, a.��¼״̬, a.�����
           From (Select a.Id, a.ҩƷid, a.���, a.����, a.����, a.ʵ������, a.��¼״̬, a.�����
                  From ҩƷ�շ���¼ A
                  Where a.����� Is Not Null And Nvl(a.��ҩ��ʽ, 0) <> -1 And (a.��¼״̬ = 1 Or Mod(a.��¼״̬, 3) = 0) And a.No = Vno And
                        a.���� In (24, 25, 26)) A,
                (Select a.����, a.ҩƷid, a.���, Sum(a.ʵ������) As �ѷ�����
                  From ҩƷ�շ���¼ A
                  Where a.����� Is Not Null And Nvl(a.��ҩ��ʽ, 0) <> -1 And a.No = Vno And
                        ���� In (Select ����
                               From ҩƷ�շ���¼
                               Where NO = Vno And ����� Is Not Null And Nvl(��ҩ��ʽ, 0) <> -1 And (��¼״̬ = 1 Or Mod(��¼״̬, 3) = 0) And
                                     ���� In (24, 25, 26))
                  Group By a.����, a.ҩƷid, a.���) B
           Where a.���� = b.���� And a.ҩƷid + 0 = b.ҩƷid And a.��� = b.��� And b.�ѷ����� <> 0) S
    Where (s.��¼״̬ = 1 Or Mod(s.��¼״̬, 3) = 0) And s.ʵ������ > (s.ʵ������ - s.�ѷ�����) And s.����� Is Not Null;

  v_Temp       Varchar2(255);
  v_��Ա����id ������Ա.����id%Type;
  v_��Ա���   ��Ա��.���%Type;
  v_��Ա����   ��Ա��.����%Type;
  Err_Custom Exception;
  v_Error Varchar2(255);
  v_Flag  Number(1) := 0;

  v_No       Varchar2(20);
  v_��ǰʱ�� Date;
  v_��ҳid   Number(18);
Begin
  v_Temp       := Zl_Identity;
  v_��Ա����id := To_Number(Substr(v_Temp, 1, Instr(v_Temp, ',') - 1));
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp       := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_��Ա���   := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_��Ա����   := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_��ǰʱ��   := Sysdate;
  v_Flag       := 0;
  Begin
    Select Nvl(Max(1), 0) Into v_Flag From ����걾��¼ Where ΢����걾 = 1 And ҽ��id = ҽ��id_In;
  Exception
    When Others Then
      v_Flag := 0;
  End;

  If v_Flag = 1 Then
  
    For r_Sample In c_Sample Loop
      Update ������Ŀ�ֲ�
      Set ҽ��id = Null
      Where �걾id In (Select Distinct ID From ����걾��¼ Where ҽ��id = r_Sample.ҽ��id);
      Update ����걾��¼
      Set ҽ��id = Null, ���� = Null, �Ա� = Null, ���� = Null, ����id = Null, ������Դ = Null, Ӥ�� = Null, �ϲ�id = Null, ���� = Null,
          �Һŵ� = Null, ����� = Null, סԺ�� = Null, �������� = Null, ��ҳid = Null, ������Ŀ = Null, �������� = Null, ���䵥λ = Null,
          �������� = Null, ������ = Null, �������id = Null, ������ = Null, ����ʱ�� = Null, �걾���� = Null, �걾��̬ = Null, ������ = Null,
          ����ʱ�� = Null, �������� = Null
      Where ҽ��id = r_Sample.ҽ��id;
      If r_Sample.�������� = 1 Then
        --ɾ��ʱ������ʱ������ɾ��
        Begin
          Delete ҽ��ִ��ʱ�� Where ҽ��id In (Select ID From ����ҽ����¼ Where r_Sample.ҽ��id In (ID, ���id));
          Delete ����ҽ������ Where ҽ��id In (Select ID From ����ҽ����¼ Where r_Sample.ҽ��id In (ID, ���id));
          Delete ����ҽ����¼ Where r_Sample.ҽ��id In (ID, ���id);
          Null;
        Exception
          When Others Then
            Null;
        End;
      Else
        Update ����ҽ������
        Set ִ��״̬ = 0
        Where ִ��״̬ = 3 And ҽ��id In (Select ID From ����ҽ����¼ Where r_Sample.ҽ��id In (ID, ���id));
      
        If r_Sample.������Դ = 2 Then
          Update /*+ rule */ סԺ���ü�¼
          Set ִ��״̬ = 0, ִ��ʱ�� = Null, ִ���� = Null
          Where ����id = r_Sample.����id And �շ���� Not In ('5', '6', '7') And
                (ҽ�����, ��¼����, NO) In
                (Select ҽ��id, ��¼����, NO
                 From ����ҽ������
                 Where ҽ��id = r_Sample.ҽ��id
                 Union All
                 Select ҽ��id, ��¼����, NO
                 From ����ҽ������
                 Where ҽ��id In (Select ID From ����ҽ����¼ Where r_Sample.ҽ��id In (ID, ���id) And ���id Is Not Null) And
                       ������ Is Null
                 Union All
                 Select a.ҽ��id, a.��¼����, a.No
                 From ����ҽ������ A, סԺ���ü�¼ B
                 Where ҽ��id In (Select ID From ����ҽ����¼ Where r_Sample.ҽ��id In (ID, ���id) And ���id Is Null) And
                       ������ Is Null And a.ҽ��id = b.ҽ����� And a.��¼���� = b.��¼���� And a.No = b.No And
                       b.ִ���� In (Select Distinct ����
                                 From ��Ա�� A, ������Ա B, ��������˵�� C
                                 Where a.Id = b.��Աid And b.����id = c.����id And c.�������� = '����'));
        Else
          Update /*+ rule */ ������ü�¼
          Set ִ��״̬ = 0, ִ��ʱ�� = Null, ִ���� = Null
          Where ����id = r_Sample.����id And �շ���� Not In ('5', '6', '7') And
                (ҽ�����, ��¼����, NO) In
                (Select ҽ��id, ��¼����, NO
                 From ����ҽ������
                 Where ҽ��id = r_Sample.ҽ��id
                 Union All
                 Select ҽ��id, ��¼����, NO
                 From ����ҽ������
                 Where ҽ��id In (Select ID From ����ҽ����¼ Where r_Sample.ҽ��id In (ID, ���id) And ���id Is Not Null) And
                       ������ Is Null
                 Union All
                 Select a.ҽ��id, a.��¼����, a.No
                 From ����ҽ������ A, ������ü�¼ B
                 Where ҽ��id In (Select ID From ����ҽ����¼ Where r_Sample.ҽ��id In (ID, ���id) And ���id Is Null) And
                       ������ Is Null And a.ҽ��id = b.ҽ����� And a.��¼���� = b.��¼���� And a.No = b.No And
                       b.ִ���� In (Select Distinct ����
                                 From ��Ա�� A, ������Ա B, ��������˵�� C
                                 Where a.Id = b.��Աid And b.����id = c.����id And c.�������� = '����'));
        
        End If;
        --ȡ���Լ����ĵ������
        v_No := '';
        Begin
          Select Distinct NO Into v_No From �����Լ���¼ Where ҽ��id = r_Sample.ҽ��id;
        Exception
          When Others Then
            v_No := '';
        End;
      
        If v_No Is Not Null Then
        
          For r_Stuff In c_Stuff(v_No) Loop
            Zl_�����շ���¼_��������(r_Stuff.Id, v_��Ա����, v_��ǰʱ��, r_Stuff.����, Null, Null, r_Stuff.�ѷ�����, 0, Null, 0);
          End Loop;
        
          v_��ҳid := Null;
          Select ��ҳid Into v_��ҳid From ����ҽ����¼ A Where ID = r_Sample.ҽ��id;
        
          If v_��ҳid Is Null Then
            Zl_������ʼ�¼_Delete(v_No, '', v_��Ա���, v_��Ա����);
          Else
            Zl_סԺ���ʼ�¼_Delete(v_No, '', v_��Ա���, v_��Ա����);
          End If;
          Update �����Լ���¼ Set NO = '' Where ҽ��id = r_Sample.ҽ��id;
        End If;
      End If;
    
    End Loop;
  
  Else
  
    For r_Sample In c_Sample Loop
    
      --����Ƿ�����ȡ������
      v_Flag := 0;
      Begin
        Select 1 Into v_Flag From ����걾��¼ Where ID = r_Sample.�걾id And ����״̬ = 2;
      Exception
        When Others Then
          v_Flag := 0;
      End;
    
      If v_Flag = 1 Then
        v_Error := '��ǰ�������ڵı걾�����Ѿ�����˵ģ�����ȡ����ˣ�';
        Raise Err_Custom;
      End If;
    
      --ɾ���ϲ�������Ŀ
      Update ����걾��¼ Set �ϲ�id = Null Where �ϲ�id In (Select ID From ����걾��¼ Where ҽ��id = ҽ��id_In);
      --���ļ���걾��¼���¼��ҽ��id,��ʵ����Բ�Ҫ����Ϣ,�Ժ���ȡ��
      Update ����걾��¼
      Set ҽ��id = Null, ���� = Null, �Ա� = Null, ���� = Null, ����id = Null, ������Դ = Null, Ӥ�� = Null, �ϲ�id = Null, ���� = Null,
          �Һŵ� = Null, ����� = Null, סԺ�� = Null, �������� = Null, ��ҳid = Null, ������Ŀ = Null, �������� = Null, ���䵥λ = Null,
          �������� = Null, ������ = Null, �������id = Null, ������ = Null, ����ʱ�� = Null, �걾���� = Null, �걾��̬ = Null, ������ = Null,
          ����ʱ�� = Null, �������� = Null
      Where ID = r_Sample.�걾id;
    
      Update ������Ŀ�ֲ� Set ҽ��id = Null Where �걾id = r_Sample.�걾id;
    
      If r_Sample.�������� = 1 Then
        --ɾ��ʱ������ʱ������ɾ��
        Begin
          Delete ҽ��ִ��ʱ�� Where ҽ��id In (Select ID From ����ҽ����¼ Where r_Sample.ҽ��id In (ID, ���id));
          Delete ����ҽ������ Where ҽ��id In (Select ID From ����ҽ����¼ Where r_Sample.ҽ��id In (ID, ���id));
          Delete ����ҽ����¼ Where r_Sample.ҽ��id In (ID, ���id);
          Null;
        Exception
          When Others Then
            Null;
        End;
      Else
        Update ����ҽ������
        Set ִ��״̬ = 0
        Where ִ��״̬ = 3 And ҽ��id In (Select ID From ����ҽ����¼ Where r_Sample.ҽ��id In (ID, ���id));
      
        If r_Sample.������Դ = 2 Then
          Update /*+ rule */ סԺ���ü�¼
          Set ִ��״̬ = 0, ִ��ʱ�� = Null, ִ���� = Null
          Where ����id = r_Sample.����id And �շ���� Not In ('5', '6', '7') And
                (ҽ�����, ��¼����, NO) In
                (Select ҽ��id, ��¼����, NO
                 From ����ҽ������
                 Where ҽ��id = r_Sample.ҽ��id
                 Union All
                 Select ҽ��id, ��¼����, NO
                 From ����ҽ������
                 Where ҽ��id In (Select ID From ����ҽ����¼ Where r_Sample.ҽ��id In (ID, ���id) And ���id Is Not Null) And
                       ������ Is Null
                 Union All
                 Select a.ҽ��id, a.��¼����, a.No
                 From ����ҽ������ A, סԺ���ü�¼ B
                 Where ҽ��id In (Select ID From ����ҽ����¼ Where r_Sample.ҽ��id In (ID, ���id) And ���id Is Null) And
                       ������ Is Null And a.ҽ��id = b.ҽ����� And a.��¼���� = b.��¼���� And a.No = b.No And
                       b.ִ���� In (Select Distinct ����
                                 From ��Ա�� A, ������Ա B, ��������˵�� C
                                 Where a.Id = b.��Աid And b.����id = c.����id And c.�������� = '����'));
        Else
          Update /*+ rule */ ������ü�¼
          Set ִ��״̬ = 0, ִ��ʱ�� = Null, ִ���� = Null
          Where ����id = r_Sample.����id And �շ���� Not In ('5', '6', '7') And
                (ҽ�����, ��¼����, NO) In
                (Select ҽ��id, ��¼����, NO
                 From ����ҽ������
                 Where ҽ��id = r_Sample.ҽ��id
                 Union All
                 Select ҽ��id, ��¼����, NO
                 From ����ҽ������
                 Where ҽ��id In (Select ID From ����ҽ����¼ Where r_Sample.ҽ��id In (ID, ���id) And ���id Is Not Null) And
                       ������ Is Null
                 Union All
                 Select a.ҽ��id, a.��¼����, a.No
                 From ����ҽ������ A, ������ü�¼ B
                 Where ҽ��id In (Select ID From ����ҽ����¼ Where r_Sample.ҽ��id In (ID, ���id) And ���id Is Null) And
                       ������ Is Null And a.ҽ��id = b.ҽ����� And a.��¼���� = b.��¼���� And a.No = b.No And
                       b.ִ���� In (Select Distinct ����
                                 From ��Ա�� A, ������Ա B, ��������˵�� C
                                 Where a.Id = b.��Աid And b.����id = c.����id And c.�������� = '����'));
        End If;
      End If;
      --ȡ���Լ����ĵ������
      v_No := '';
      Begin
        Select Distinct NO Into v_No From �����Լ���¼ Where ҽ��id = r_Sample.ҽ��id;
      Exception
        When Others Then
          v_No := '';
      End;
      If v_No Is Not Null Then
        For r_Stuff In c_Stuff(v_No) Loop
          Zl_�����շ���¼_��������(r_Stuff.Id, v_��Ա����, v_��ǰʱ��, r_Stuff.����, Null, Null, r_Stuff.�ѷ�����, 0, Null, 0);
        End Loop;
      
        v_��ҳid := Null;
        Select ��ҳid Into v_��ҳid From ����ҽ����¼ A Where ID = r_Sample.ҽ��id;
      
        If v_��ҳid Is Null Then
          Zl_������ʼ�¼_Delete(v_No, '', v_��Ա���, v_��Ա����);
        Else
          Zl_סԺ���ʼ�¼_Delete(v_No, '', v_��Ա���, v_��Ա����);
        End If;
        Update �����Լ���¼ Set NO = '' Where ҽ��id = r_Sample.ҽ��id;
      End If;
      --ɾ���Լ����ĵ�
    --Delete From �����Լ���¼ Where ҽ��id = r_Sample.ҽ��id;
    End Loop;
  End If;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����걾��¼_תΪ����;
/

--115597:����,2017-12-18,�ϸ���Ʋ�֧���ظ�ʹ�õ�Ժ�ڿ��������ݿ�������ͬ����ʱ���ظ�ʹ�øÿ���Ʊ�����ü�¼�����ᷢ�������仯
Create Or Replace Procedure Zl_ҽ�ƿ���¼_Insert
(
  --��������������=0-����,1-����,2-����(�൱���ش�)
  --      ����ʱ,���ݺ�_IN�������ԭ��/�����ĵ��ݺš�
  --      ����/������,�ٻ���ʱ�������һ�ο���Ϊ׼��
  ��������_In     Number,
  ���ݺ�_In       סԺ���ü�¼.No%Type,
  ����id_In       סԺ���ü�¼.����id%Type,
  ��ҳid_In       סԺ���ü�¼.��ҳid%Type,
  ��ʶ��_In       סԺ���ü�¼.��ʶ��%Type,
  �ѱ�_In         סԺ���ü�¼.�ѱ�%Type,
  �����id_In     ҽ�ƿ����.Id%Type,
  ԭ����_In       ����ҽ�ƿ���Ϣ.����%Type,
  ҽ�ƿ���_In     ����ҽ�ƿ���Ϣ.����%Type,
  �䶯ԭ��_In     ����ҽ�ƿ��䶯.�䶯ԭ��%Type,
  ����_In         ������Ϣ.����֤��%Type,
  ����_In         סԺ���ü�¼.����%Type,
  �Ա�_In         סԺ���ü�¼.�Ա�%Type,
  ����_In         סԺ���ü�¼.����%Type,
  ���˲���id_In   סԺ���ü�¼.���˲���id%Type,
  ���˿���id_In   סԺ���ü�¼.���˿���id%Type,
  �շ�ϸĿid_In   סԺ���ü�¼.�շ�ϸĿid%Type,
  �շ����_In     סԺ���ü�¼.�շ����%Type,
  ���㵥λ_In     סԺ���ü�¼.���㵥λ%Type,
  ������Ŀid_In   סԺ���ü�¼.������Ŀid%Type,
  �վݷ�Ŀ_In     סԺ���ü�¼.�վݷ�Ŀ%Type,
  ��׼����_In     סԺ���ü�¼.��׼����%Type,
  ִ�в���id_In   סԺ���ü�¼.ִ�в���id%Type,
  ��������id_In   סԺ���ü�¼.��������id%Type,
  ����Ա���_In   סԺ���ü�¼.����Ա���%Type,
  ����Ա����_In   סԺ���ü�¼.����Ա����%Type,
  �Ӱ��־_In     סԺ���ü�¼.�Ӱ��־%Type,
  ����ʱ��_In     סԺ���ü�¼.�Ǽ�ʱ��%Type,
  ����id_In       Ʊ��ʹ����ϸ.����id%Type,
  Ic����_In       ������Ϣ.Ic����%Type := Null,
  Ӧ�ս��_In     סԺ���ü�¼.Ӧ�ս��%Type,
  ʵ�ս��_In     סԺ���ü�¼.ʵ�ս��%Type,
  ���㷽ʽ_In     ����Ԥ����¼.���㷽ʽ%Type,
  ˢ�����id_In   ����Ԥ����¼.�����id%Type,
  ���ѿ�_In       Integer := 0,
  ˢ������_In     ����ҽ�ƿ���Ϣ.����%Type,
  ����id_In       ����Ԥ����¼.����id%Type,
  ������ˮ��_In   ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In     ����Ԥ����¼.����˵��%Type := Null,
  ������λ_In     ����Ԥ����¼.������λ%Type := Null,
  ���½������_In Number := 0, --�Ƿ������Ա��������Ҫ�Ǵ���ͳһ����Ա��¼��̨�������������
  ժҪ_In         סԺ���ü�¼.ժҪ%Type := Null
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
       ���ʽ��, �ɿ���id, ����, ժҪ)
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
      Select Nvl(Max(a.���մ���), 0), Nvl(Max(a.����), 0)
      Into n_���մ���, n_����
      From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B, סԺ���ü�¼ C
      Where a.��ӡid = b.Id And b.No = c.No And a.Ʊ�� = 5 And c.���� = �����id_In And c.��¼���� = 5 And a.���� = ҽ�ƿ���_In;
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
      If Nvl(���½������_In, 0) = 0 Then
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
      End If;
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

--118465:��˼��,2017-12-15,�ṩ���µĹ���
Create Or Replace Procedure Zl_Ӱ��������_Update
(
  ���id_In           In Ӱ��������.���id%Type,
  ����_In             In Ӱ��������.����%Type,
  ����_In             In Ӱ��������.����%Type,
  ��������_In         In Ӱ��������.��������%Type,
  �Ƿ�����_In         In Ӱ��������.�Ƿ�����%Type,
  �Ƿ�����Ҽ��˵�_In In Ӱ��������.�Ƿ�����Ҽ��˵�%Type,
  �Ƿ���빤����_In   In Ӱ��������.�Ƿ���빤����%Type,
  �Զ�ִ��ʱ��_In     In Ӱ��������.�Զ�ִ��ʱ��%Type,
  Vbs�ű�_In          In Ӱ��������.Vbs�ű�%Type
) Is

  n_������� Ӱ��������.�������%Type;
  n_Id       Number;

Begin

  Select Nvl(Max(�������), 0) + 1 Into n_������� From Ӱ�������� Where ���id = ���id_In;
  Select Nvl(Max(ID), 0) + 1 Into n_Id From Ӱ��������;

  Insert Into Ӱ��������
    (ID, ���id, �������, ����, ����, ��������, �Ƿ�����, �Ƿ�����Ҽ��˵�, �Ƿ���빤����, �Զ�ִ��ʱ��, Vbs�ű�)
  Values
    (n_Id, ���id_In, n_�������, ����_In, ����_In, ��������_In, �Ƿ�����_In, �Ƿ�����Ҽ��˵�_In, �Ƿ���빤����_In, �Զ�ִ��ʱ��_In, Vbs�ű�_In);

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Ӱ��������_Update;
/

--116339:����,2017-12-14,��¼��Ŀ����̷��������޸�
Create Or Replace Procedure Zl_�����¼��Ŀ_Insert
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
  ȱʡֵ_In     In �����¼��Ŀ.ȱʡֵ%Type := Null
) Is
Begin
  Insert Into �����¼��Ŀ
    (��Ŀ���, ��Ŀ����, ��Ŀ����, ��Ŀ����, ��ĿС��, ��Ŀ��λ, ��Ŀ��ʾ, ��Ŀֵ��, ����ȼ�, ������, ��Ŀid, ���ÿ���, Ӧ�÷�ʽ, ���ò���, ��Ŀ����, Ӧ�ó���, ˵��, ȱʡֵ)
  Values
    (��Ŀ���_In, ��Ŀ����_In, ��Ŀ����_In, ��Ŀ����_In, ��ĿС��_In, ��Ŀ��λ_In, ��Ŀ��ʾ_In, ��Ŀֵ��_In, ����ȼ�_In, ������_In, ��Ŀid_In, 1, Ӧ�÷�ʽ_In,
     ���ò���_In, ��Ŀ����_In, Ӧ�ó���_In, ˵��_In, ȱʡֵ_In);

  If ��Ŀ��ʾ_In = 4 Then
    Insert Into ���������Ŀ
      (���, �����)
      Select ��Ŀ���_In, Null From Dual Where Not Exists (Select 1 From ���������Ŀ Where ��� = ��Ŀ���_In);
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����¼��Ŀ_Insert;
/

--116339:����,2017-12-14,��¼��Ŀ����̷��������޸�
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
  ȱʡֵ_In     In �����¼��Ŀ.ȱʡֵ%Type := Null
) Is
  n_���� Number(1);
Begin
  n_���� := 0;
  Select Count(��Ŀ���) Into n_���� From �����¼��Ŀ Where ��Ŀ��� = ��Ŀ���_In And ��Ŀ��ʾ = 4;
  Update �����¼��Ŀ
  Set ��Ŀ���� = ��Ŀ����_In, ��Ŀ���� = ��Ŀ����_In, ��Ŀ���� = ��Ŀ����_In, ��ĿС�� = ��ĿС��_In, ��Ŀ��λ = ��Ŀ��λ_In, ��Ŀ��ʾ = ��Ŀ��ʾ_In, ��Ŀֵ�� = ��Ŀֵ��_In,
      ����ȼ� = ����ȼ�_In, ������ = ������_In, ��Ŀid = ��Ŀid_In, Ӧ�÷�ʽ = Ӧ�÷�ʽ_In, ���ò��� = ���ò���_In, ��Ŀ���� = ��Ŀ����_In, Ӧ�ó��� = Ӧ�ó���_In,
      ˵�� = ˵��_In, ȱʡֵ = ȱʡֵ_In
  Where ��Ŀ��� = ��Ŀ���_In;

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

--116848:������,2017-12-14,��Ѫ���������Ѫ�Ͷ���
Create Or Replace Function Zl_Fun_BloodApplyCode
(
  ��������_In Number,
  ��Ѫ����_In Number,
  ģʽ_In     Number := 1
) Return Varchar2 As
  v_Return Varchar2(100);
Begin
  --����˵��:�µ�����ʱ��1��ǿ�ƿ����Ƿ������޸����뵥ABO��RH��2��������Ѫ���뵥�ϼ���ָ���е�ABO��RHָ�����(Ҳ��ֻ����ABOָ�����)��
  ----                                                        A����ȡ������ʱ���Զ�����ABO��RH��B:����ʱ���ABO��RH�Ƿ�ͼ���������һ��

  --���˵����
  ----��������_In=1-��Ѫ���뵥;2-ȡѪ֪ͨ��(����ҽԺ�����������Ϳ���)
  ----��Ѫ����_In=0-��ͨ��Ѫ;1-������Ѫ(����ҽԺ������Ѫ�����̶ȿ���)
  ----ģʽ_in:0=ͨ������ֵ�����Ƿ������������뵥ABO��RH��1=���ݼ���������ABO��RH�������Ǽ��ABO��RH�Ƿ�ͼ�����һ��(��Ѫ���뵥ʱ��Ч)
  --�������أ�ģʽ_in=0ʱ������0(�����޸�)��1(�������޸�)��ģʽ_in=1ʱ�������ַ�����ʽ: ABOָ�����:0(ѯ��)��1(��ֹ),RHָ�����:0(ѯ��)��1(��ֹ)��
  ----        �磺800001:1,��ʾ����ʱABO�ͼ�������һ��ʱ���ֹ���棬Ҳ��ֱ��дָ����룬�磺800001����ʾ����ʱABO�ͼ�������һ��ʱ�����ѯ�ʡ�

  If  ģʽ_In = 0 Then
    --0��ʾ�����޸�ABO;1�������޸�ABO
    v_Return := '0';
  Else
    --���ؿղ��Զ�ƥ��ABO��RH���ұ���ʱ�����м�顣
    v_Return := '';
  End If;
  Return v_Return;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Fun_BloodApplyCode;
/

--116846:������,2017-12-14,��Ѫ���뱣���Զ��庯�����
Create Or Replace Function Zl1_EX_BloodApplyCheck
(
  ���ó���_In         Number,
  ����id_In           ����ҽ����¼.����id%Type,
  ����id_In           Number,
  ��������_In         Number,
  ��Ѫ����_In         Number,
  �Ƿ����_In         ��Ѫ�����¼.�Ƿ����%Type,
  �������_In         ����ҽ������.����%Type,
  ���ids_In          Varchar2,
  ��Ѫ����_In         ��Ѫ����.����%Type,
  ��ѪĿ��_In         ��ѪĿ��.����%Type,
  ��Ѫ����_In         ��Ѫ����.����%Type,
  Ԥ����Ѫ����_In     ����ҽ����¼.�걾��λ%Type,
  Ѫ��_In             Ѫ��.����%Type,
  Rhd_In              Varchar2,
  ������Ŀ_In         Varchar2,
  ��Ѫִ�п���id_In   ����ҽ����¼.ִ�п���id%Type,
  ;��id_In           ����ҽ����¼.������Ŀid%Type,
  ;��ִ�п���id_In   ����ҽ����¼.ִ�п���id%Type,
  ��ע_In             ����ҽ����¼.ҽ������%Type,
  Ӥ�����_In         ����ҽ����¼.Ӥ��%Type := 0,
  ������Ѫʷ_In       ��Ѫ�����¼.������Ѫʷ%Type := Null,
  ������Ѫ��Ӧʷ_In   ��Ѫ�����¼.������Ѫ��Ӧʷ%Type := Null,
  ��Ѫ���ɼ�����ʷ_In ��Ѫ�����¼.��Ѫ���ɼ�����ʷ%Type := Null,
  �в����_In         ��Ѫ�����¼.�в����%Type := Null,
  ��Ѫ������_In       ��Ѫ�����¼.��Ѫ������%Type := Null,
  ������_In         Varchar2 := Null
) Return Varchar2
--����˵�����¿����޸���Ѫ����ʱ����������֮ǰ�������������ݽ��м�飬��������ʾ����������
--����˵��������������ݱ���֮ǰ�����뵥���ݽ����ض����Ϳ��ƣ�������ҽԺҵ���ض����ڣ���������˺���
--���˵����
  ----���ó���_in=1-����,2-סԺ 
  ----����id_In=����ʱ���Һż�¼id,סԺʱ������ҳid 
  ----��������_In=1-��Ѫ���뵥;2-ȡѪ֪ͨ��
  ----��Ѫ����_In=0-��ͨ��Ѫ;1-������Ѫ
  ----�Ƿ����_In=0-�Ǵ���;1-����
  ----�������_In=����ʱΪ�գ�����Ϊ���������Ϣ
  ----���ids_In=����ҳѡ�������������iD����������','�ŷָ����¼������Ϊ��
  ----��Ѫ����_In=��Ӧ��Ѫ�����ֵ�������
  ----��ѪĿ��_In=��Ӧ��ѪĿ���ֵ�������
  ----��Ѫ����_In=��Ӧ��Ѫ�����ֵ�������
  ----Ѫ��_In=��ӦѪ���ֵ�������
  ----Rhd_In=;+;-
  ----������Ŀ_In=��������Ŀ+�������ķ�ʽ���룬��������Ʒ������';'�ָ��ʽ�磺��Ѫ������ĿID,������;��Ѫ������ĿID,������
  ----;��id_In=��Ѫ���뵥���ǲɼ���ʽ��������Ŀid��ȡѪ֪ͨ��������Ѫ;����������ĿID
  ----�в����_In=��ʽ:�д�/����
  ----������_In=��Ѫ�����򷵻����뵥�·��ļ�������Ϣ���ֶ�������ο�:��Ѫ����������ȡѪ������Ϊ�ա�����ָ���������<SplitCol>�ָ��ָͬ��֮����<SplitRow>�ָ���ظ�ʽ���£�
  ----            ������ĿID<SplitCol>ָ�����<SplitCol>ָ��������<SplitCol>ָ��Ӣ����<SplitCol>ָ����<SplitCol>�����λ<SplitCol>�����־<SplitCol>����ο�
  ----            <SplitCol>ȡֵ����<SplitCol>�Ƿ��˹���д<SplitRow>������ĿID<SplitCol>ָ�����<SplitCol>ָ��������<SplitCol>ָ��Ӣ����<SplitCol>ָ����
  ----            <SplitCol>�����λ<SplitCol>�����־<SplitCol>����ο�<SplitCol>ȡֵ����<SplitCol>�Ƿ��˹���д
  ----    
--�������أ�"������|��ʾ��Ϣ",������=0-����,1-ѯ����ʾ,2-��ֹ��������Ϊ0ʱ�����践����ʾ��Ϣ���ָ����� 
 As
  v_Return Varchar2(200);
Begin
  v_Return := Null;
  Return v_Return;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl1_EX_BloodApplyCheck;
/

--118154:��ΰ��,2017-12-12,������ҩ��������
Create Or Replace Procedure Zl_����ҽ����¼_Update
(
  Id_In           ����ҽ����¼.Id%Type,
  ���id_In       ����ҽ����¼.���id%Type,
  ���_In         ����ҽ����¼.���%Type,
  ҽ��״̬_In     ����ҽ����¼.ҽ��״̬%Type,
  ҽ����Ч_In     ����ҽ����¼.ҽ����Ч%Type,
  ������Ŀid_In   ����ҽ����¼.������Ŀid%Type,
  �շ�ϸĿid_In   ����ҽ����¼.�շ�ϸĿid%Type,
  ����_In         ����ҽ����¼.����%Type,
  ��������_In     ����ҽ����¼.��������%Type,
  �ܸ�����_In     ����ҽ����¼.�ܸ�����%Type,
  ҽ������_In     ����ҽ����¼.ҽ������%Type,
  ҽ������_In     ����ҽ����¼.ҽ������%Type,
  �걾��λ_In     ����ҽ����¼.�걾��λ%Type,
  ִ��Ƶ��_In     ����ҽ����¼.ִ��Ƶ��%Type,
  Ƶ�ʴ���_In     ����ҽ����¼.Ƶ�ʴ���%Type,
  Ƶ�ʼ��_In     ����ҽ����¼.Ƶ�ʼ��%Type,
  �����λ_In     ����ҽ����¼.�����λ%Type,
  ִ��ʱ�䷽��_In ����ҽ����¼.ִ��ʱ�䷽��%Type,
  �Ƽ�����_In     ����ҽ����¼.�Ƽ�����%Type,
  ִ�п���id_In   ����ҽ����¼.ִ�п���id%Type,
  ִ������_In     ����ҽ����¼.ִ������%Type,
  ������־_In     ����ҽ����¼.������־%Type,
  ��ʼִ��ʱ��_In ����ҽ����¼.��ʼִ��ʱ��%Type,
  ִ����ֹʱ��_In ����ҽ����¼.ִ����ֹʱ��%Type,
  ���˿���id_In   ����ҽ����¼.���˿���id%Type,
  ��������id_In   ����ҽ����¼.��������id%Type,
  ����ҽ��_In     ����ҽ����¼.����ҽ��%Type,
  ����ʱ��_In     ����ҽ����¼.����ʱ��%Type,
  ��鷽��_In     ����ҽ����¼.��鷽��%Type := Null,
  ִ�б��_In     ����ҽ����¼.ִ�б��%Type := Null,
  �ɷ����_In     ����ҽ����¼.�ɷ����%Type := Null,
  ժҪ_In         ����ҽ����¼.ժҪ%Type := Null,
  ��Ա������_In   ����ҽ��״̬.������Ա%Type := Null,
  ��Ѽ���_In     ����ҽ����¼.��Ѽ���%Type := Null,
  ��ҩĿ��_In     ����ҽ����¼.��ҩĿ��%Type := Null,
  ��ҩ����_In     ����ҽ����¼.��ҩ����%Type := Null,
  ���״̬_In     ����ҽ����¼.���״̬%Type := Null,
  ����˵��_In     ����ҽ����¼.����˵��%Type := Null,
  �״�����_In     ����ҽ����¼.�״�����%Type := Null,
  �������_In     ����ҽ����¼.�������%Type := Null,
  �����Ŀid_In   ����ҽ����¼.�����Ŀid%Type := Null,
  Ƥ�Խ��_In     ����ҽ����¼.Ƥ�Խ��%Type := Null,
  �������_In     ����ҽ����¼.�������%Type := Null
  --���ܣ���ҽ����ʿ�޸��˲������ݵ�ҽ����¼�������������סԺ��
  --˵����Updateʱ֮�����漰������ĿID,�Ƽ����Ա仯,����Ϊ��ҩ;��,�÷��ı仯
  --      Updateʱ֮�����漰��Ч�仯,����Ϊ����¼��ҽ��������ı���Ч
) Is
  v_Count Number;

  v_Temp     Varchar2(255);
  v_��Ա���� ����ҽ��״̬.������Ա%Type;
  v_�����������IDs varChar2(4000);

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
     zl_�������_cancel(Id_In,v_�����������IDs);
  End If;

  If v_�����������IDs Is Not Null Then
  v_Error :=  'ҽ��"' || ҽ������_In || '"�����������ڽ��д�����飬�������޸ġ�';
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
      ������� = �������_In
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

--107484:��˼��,2017-12-12,����Zl_Ӱ����Ϣ_Xml���ݻ�ȡ�� Ӥ������������ȡ��ʽ
--107484:��˼��,2017-12-13,���ֻ�ԭ����ǰ����ʽ�����ӹ��̴�����
--117484:��˼��,2017-12-14,����Zl_Ӱ����Ϣ_Xml���ݻ�ȡ�� Ӥ������������ȡ��ʽ
Create Or Replace Function zl_Ӱ����Ϣ_XML���ݻ�ȡ
( 
    ҽ��ID_In ����ҽ����¼.id%Type, 
    ��Ϣ����_In varchar2, 
    ��ǰ�û�_In varchar2,
    ��Ϣ���_In varchar2:=Null         --�°棺��鱨��ID 
) Return varchar2 IS 
  v_Context varchar2(4000); 
  n_Ӥ����� ����ҽ����¼.Ӥ��%Type;
  v_���� ����ҽ����¼.����%Type;
  n_��ҳid   ����ҽ����¼.��ҳid%Type;
 
  --ZLHIS_CIS_005(ҽ��ִ�а������) 
  Function Get_Zlhis_Cis_005 Return varchar2 As 
    v_Return varchar2(4000); 
  Begin 
        Select 
          '<patient_info>' || 
             '<patient_id>' || a.����id || '</patient_id>' || 
             '<patient_name>' || v_���� ||'</patient_name>' || 
          '</patient_info>' || 
          '<patient_clinic>' || 
             '<patient_source>' || b.������Դ ||'</patient_source>' || 
             '<clinic_dept_id>' || b.���˿���id || '</clinic_dept_id>' || 
          '</patient_clinic>' || 
          '<patient_order>' || 
             '<order_id>' || c.ҽ��id || '</order_id>' || 
             '<order_expiry>' || b.ҽ����Ч ||'</order_expiry>' || 
             '<order_kind>' || b.������� || '</order_kind>' || 
             '<operation_kind>' || d.�������� ||'</operation_kind>' || 
             '<order_item_id>' || c.ҽ��id || '</order_item_id>' || 
             '<order_item_title>' || b.ҽ������ ||'</order_item_title>' || 
          '</patient_order>' || 
          '<arrange_result>' || 
             '<arrange_time>' ||To_Char(c.����ʱ��,'yyyy/mm/dd hh24:mi:ss')|| '</arrange_time>' || 
          '</arrange_result>'   Into v_Return 
 
      From ������Ϣ A, ����ҽ����¼ B, ����ҽ������ C, ������ĿĿ¼ D 
      Where a.����id = b.����id And c.ҽ��id = b.Id And b.������Ŀid = d.Id And c.����ʱ�� Is Not Null And b.���id Is Null And 
          b.������� = 'D' And b.Id = ҽ��id_In; 
    Return v_Return; 
  End Get_Zlhis_Cis_005; 
 
  --ZLHIS_CIS_017(���߼������) 
  Function Get_ZLHIS_CIS_017 Return varchar2 As 
    v_return varchar2(4000); 
  Begin 
        Select 
           '<patient_info>' || 
               '<patient_id>' || a.����id || '</patient_id>' || 
               '<patient_name>' || v_���� || '</patient_name>' || 
               '<in_number>' || nvl(a.סԺ��,0) || '</in_number>' || 
               '<out_number>' || nvl(a.�����,0) || '</out_number>' || 
           '</patient_info>' || 
           '<patient_clinic>' || 
               '<patient_source>' || b.������Դ || '</patient_source>' || 
               '<clinic_id>' || Case b.������Դ When 1 Then nvl(d.id,0) When 2 Then nvl(b.��ҳid, 0) Else 0 End || '</clinic_id>' || 
               '<clinic_dept_id>' || nvl(a.��ǰ����id,0) || '</clinic_dept_id>' || 
           '</patient_clinic>' || 
           '<check_request>' || 
               '<request_id>' || b.id || '</request_id>' || 
               '<check_item_id>' || b.������Ŀid || '</check_item_id>' || 
               '<check_item_title>' || b.ҽ������ || '</check_item_title>' || 
               '<execute_dept_id>' || nvl(c.ִ�в���id,0) || '</execute_dept_id>' || 
               '<send_serial>' || c.���ͺ� || '</send_serial>' || 
               '<bill_no>' || c.NO || '</bill_no>' || 
               '<bill_kind>' || c.��¼���� || '</bill_kind>' || 
               '<create_doctor>' || b.����ҽ�� || '</create_doctor>' || 
               '<create_time>' || b.����ʱ�� || '</create_time>' || 
               '<create_dept_id>' || nvl(b.��������id,0) || '</create_dept_id>' || 
           '</check_request>' into v_return 
 
        From ������Ϣ a, ����ҽ����¼ b, ����ҽ������ c, ���˹Һż�¼ d 
        Where a.����id=b.����id And b.id=c.ҽ��id And b.�Һŵ�=d.no(+) And b.���ID Is Null 
              And a.����id=b.����id And b.id=ҽ��ID_In; 
 
      Return v_return; 
  End Get_ZLHIS_CIS_017; 
  
  --ZLHIS_CIS_015(ҽ���ܾ�ִ��) 
  Function Get_ZLHIS_CIS_015 Return varchar2 As 
    v_return varchar2(4000); 
  Begin 
       Select 
           '<patient_info>' || 
               '<patient_id>' || a.����id || '</patient_id>' || 
               '<patient_name>' || v_���� || '</patient_name>' || 
               '<in_number>' || nvl(a.סԺ��,0) || '</in_number>' || 
               '<out_number>' || nvl(a.�����,0) || '</out_number>' || 
           '</patient_info>' || 
           '<patient_clinic>' || 
               '<patient_source>' || b.������Դ || '</patient_source>' || 
               '<clinic_id>' || Case b.������Դ When 1 Then nvl(e.id,0) When 2 Then nvl(b.��ҳid, 0) Else 0 End || '</clinic_id>' || 
               '<clinic_dept_id>' || nvl(a.��ǰ����id,0) || '</clinic_dept_id>' || 
           '</patient_clinic>' || 
           '<refuse_order>' || 
               '<order_id>' || b.id || '</order_id>' || 
               '<order_expiry>1</order_expiry>' || 
               '<order_kind>' || b.������� || '</order_kind>' || 
               '<operation_kind>' || d.�������� || '</operation_kind>' || 
               '<order_item_id>' || b.������ĿID || '</order_item_id>' || 
               '<order_item_title>' || d.���� || '</order_item_title>' || 
           '</refuse_order>' Into v_return 
 
       From ������Ϣ a, ����ҽ����¼ b, ����ҽ������ c, ������ĿĿ¼ d, ���˹Һż�¼ e 
       Where a.����id=b.����id And b.id=c.ҽ��id And b.������Ŀid=d.id And b.�Һŵ�=e.no(+) And b.���ID Is Null 
              And a.����id=b.����id And b.id=ҽ��ID_In; 
              
       Return v_return;     
  End Get_ZLHIS_CIS_015;   
   
 
  --ZLHIS_CIS_024(����ҽ������) 
  Function Get_ZLHIS_CIS_024 Return varchar2 As 
    v_return varchar2(4000); 
  Begin 
       Select 
           '<patient_info>' || 
               '<patient_id>' || a.����id || '</patient_id>' || 
               '<patient_name>' || v_���� || '</patient_name>' || 
               '<in_number>' || nvl(a.סԺ��,0) || '</in_number>' || 
               '<out_number>' || nvl(a.�����,0) || '</out_number>' || 
           '</patient_info>' || 
           '<patient_clinic>' || 
               '<patient_source>' || b.������Դ || '</patient_source>' || 
               '<clinic_id>' || Case b.������Դ When 1 Then nvl(e.id,0) When 2 Then nvl(b.��ҳid, 0) Else 0 End || '</clinic_id>' || 
               '<clinic_dept_id>' || nvl(a.��ǰ����id,0) || '</clinic_dept_id>' || 
           '</patient_clinic>' || 
           '<cancel_order>' || 
               '<order_id>' || b.id || '</order_id>' || 
               '<order_kind>' || b.������� || '</order_kind>' || 
               '<operation_kind>' || d.�������� || '</operation_kind>' || 
           '</cancel_order>' Into v_return 
 
       From ������Ϣ a, ����ҽ����¼ b, ����ҽ������ c, ������ĿĿ¼ d, ���˹Һż�¼ e 
       Where a.����id=b.����id And b.id=c.ҽ��id And b.������Ŀid=d.id And b.�Һŵ�=e.no(+) And b.���ID Is Null 
              And a.����id=b.����id And b.id=ҽ��ID_In; 
 
      Return v_return; 
  End Get_ZLHIS_CIS_024; 
 
 
  --ZLHIS_PACS_001(��鱨�����) 
  Function Get_ZLHIS_PACS_001 Return varchar2 As 
    v_return varchar2(4000); 
  Begin 
    If ��Ϣ���_In Is Null Then
        Select 
           '<patient_info>' || 
               '<patient_id>' || a.����id || '</patient_id>' || 
               '<patient_name>' || v_���� || '</patient_name>' || 
               '<in_number>' || nvl(a.סԺ��,0) || '</in_number>' || 
               '<out_number>' || nvl(a.�����,0) || '</out_number>' || 
           '</patient_info>' || 
           '<patient_clinic>' || 
               '<patient_source>' || b.������Դ || '</patient_source>' || 
               '<clinic_id>' || Case b.������Դ When 1 Then nvl(e.id,0) When 2 Then nvl(b.��ҳid, 0) Else 0 End || '</clinic_id>' || 
               '<clinic_dept_id>' || nvl(a.��ǰ����id,0) || '</clinic_dept_id>' || 
           '</patient_clinic>' || 
           '<advice_info>' || 
               '<order_id>' || b.id || '</order_id>' || 
               '<check_item_id>' || b.������Ŀid || '</check_item_id>' || 
               '<check_item_title>' || b.ҽ������ || '</check_item_title>' || 
               '<create_doctor>' || b.����ҽ�� || '</create_doctor>' || 
               '<create_dept_id>' || nvl(b.��������id, 0) || '</create_dept_id>' || 
               '<study_execute_id>' || nvl(b.ִ�п���id,0) || '</study_execute_id>' || 
           '</advice_info>' || 
           '<report_inf>' || 
               '<report_id>' || d.����id || '</report_id>' || 
               '<report_doctor>' || c.����� || '</report_doctor>' || 
               '<result_positive>' || c.������� || '</result_positive>' || 
           '</report_inf>'  Into v_return 
 
        From ������Ϣ a, ����ҽ����¼ b, ����ҽ������ c, ����ҽ������ d, ���˹Һż�¼ e 
        Where a.����id=b.����id And b.id=c.ҽ��id And b.id=d.ҽ��id And b.�Һŵ�=e.no(+) And b.���ID Is Null 
              And d.��鱨��id Is Null And a.����id=b.����id And b.id=ҽ��ID_In; 
    Else
        Select 
           '<patient_info>' || 
               '<patient_id>' || a.����id || '</patient_id>' || 
               '<patient_name>' || v_���� || '</patient_name>' || 
               '<in_number>' || nvl(a.סԺ��,0) || '</in_number>' || 
               '<out_number>' || nvl(a.�����,0) || '</out_number>' || 
           '</patient_info>' || 
           '<patient_clinic>' || 
               '<patient_source>' || b.������Դ || '</patient_source>' || 
               '<clinic_id>' || Case b.������Դ When 1 Then nvl(e.id,0) When 2 Then nvl(b.��ҳid, 0) Else 0 End || '</clinic_id>' || 
               '<clinic_dept_id>' || nvl(a.��ǰ����id,0) || '</clinic_dept_id>' || 
           '</patient_clinic>' || 
           '<advice_info>' || 
               '<order_id>' || b.id || '</order_id>' || 
               '<check_item_id>' || b.������Ŀid || '</check_item_id>' || 
               '<check_item_title>' || b.ҽ������ || '</check_item_title>' || 
               '<create_doctor>' || b.����ҽ�� || '</create_doctor>' || 
               '<create_dept_id>' || nvl(b.��������id, 0) || '</create_dept_id>' || 
               '<study_execute_id>' || nvl(b.ִ�п���id,0) || '</study_execute_id>' || 
           '</advice_info>' || 
           '<report_inf>' || 
               '<report_id>' || d.��鱨��id || '</report_id>' || 
               '<report_doctor>' || c.����� || '</report_doctor>' || 
               '<result_positive>' || c.������� || '</result_positive>' || 
           '</report_inf>'  Into v_return 
 
        From ������Ϣ a, ����ҽ����¼ b, ����ҽ������ c, ����ҽ������ d, ���˹Һż�¼ e 
        Where a.����id=b.����id And b.id=c.ҽ��id And b.id=d.ҽ��id And b.�Һŵ�=e.no(+) And b.���ID Is Null 
              And d.����id Is Null And a.����id=b.����id And b.id=ҽ��ID_In And d.��鱨��id=��Ϣ���_In; 
    End If;
    
    Return v_return; 
  End Get_ZLHIS_PACS_001; 
 
  --ZLHIS_PACS_002(����״̬ͬ��) 
  Function Get_ZLHIS_PACS_002 Return varchar2 As 
    v_return varchar2(4000); 
  Begin 
        Select 
           '<patient_info>' || 
               '<patient_id>' || a.����id || '</patient_id>' || 
               '<patient_name>' || v_���� || '</patient_name>' || 
               '<identity_card>' || a.���֤�� || '</identity_card>' || 
               '<in_number>' || nvl(a.סԺ��,0) || '</in_number>' || 
               '<out_number>' || nvl(a.�����,0) || '</out_number>' || 
           '</patient_info>' || 
           '<patient_clinic>' || 
               '<patient_source>' || b.������Դ || '</patient_source>' || 
               '<clinic_id>' || Case b.������Դ When 1 Then nvl(d.id,0) When 2 Then nvl(b.��ҳid, 0) Else 0 End || '</clinic_id>' || 
               '<clinic_dept_id>' || nvl(a.��ǰ����id,0) || '</clinic_dept_id>' || 
           '</patient_clinic>' || 
           '<study_state>' || 
               '<study_cur_state>' || nvl(c.ִ�й���,0) || '</study_cur_state>' || 
               '<study_cur_time>' || sysdate || '</study_cur_time>' || 
               '<order_id>' || b.id || '</order_id>' || 
               '<study_item_id>' || b.������Ŀid || '</study_item_id>' || 
               '<study_item_title>' || b.ҽ������ || '</study_item_title>' || 
               '<study_oper_person>' || ��ǰ�û�_In || '</study_oper_person>' || 
               '<study_execute_id>' || nvl(b.ִ�п���id,0) || '</study_execute_id>' || 
           '</study_state>' Into v_return 
 
        From ������Ϣ a, ����ҽ����¼ b, ����ҽ������ c, ���˹Һż�¼ d 
        Where a.����id=b.����id And b.id=c.ҽ��id And b.�Һŵ�=d.no(+) And b.���ID Is Null 
              And a.����id=b.����id And b.id=ҽ��ID_In; 
 
      Return v_return; 
  End Get_ZLHIS_PACS_002; 
 
 
  --ZLHIS_PACS_003(���״̬����) 
  Function Get_ZLHIS_PACS_003 Return varchar2 As 
    v_return varchar2(4000); 
  Begin 
        Select 
           '<patient_info>' || 
               '<patient_id>' || a.����id || '</patient_id>' || 
               '<patient_name>' || v_���� || '</patient_name>' || 
               '<identity_card>' || a.���֤�� || '</identity_card>' || 
               '<in_number>' || nvl(a.סԺ��,0) || '</in_number>' || 
               '<out_number>' || nvl(a.�����,0) || '</out_number>' || 
           '</patient_info>' || 
           '<patient_clinic>' || 
               '<patient_source>' || b.������Դ || '</patient_source>' || 
               '<clinic_id>' || Case b.������Դ When 1 Then nvl(d.id,0) When 2 Then nvl(b.��ҳid, 0) Else 0 End || '</clinic_id>' || 
               '<clinic_dept_id>' || nvl(a.��ǰ����id, 0) || '</clinic_dept_id>' || 
           '</patient_clinic>' || 
           '<study_state>' || 
               '<study_cur_state>' || nvl(c.ִ�й���,0) || '</study_cur_state>' || 
               '<study_cur_time>' || Sysdate || '</study_cur_time>' || 
               '<study_order_id>' || b.id || '</study_order_id>' || 
               '<study_item_id>' || b.������Ŀid || '</study_item_id>' || 
               '<study_item_title>' || b.ҽ������ || '</study_item_title>' || 
               '<study_oper_person>' || ��ǰ�û�_In || '</study_oper_person>' || 
               '<study_execute_id>' || nvl(b.ִ�п���id,0) || '</study_execute_id>' || 
           '</study_state>' Into v_return 
 
        From ������Ϣ a, ����ҽ����¼ b, ����ҽ������ c, ���˹Һż�¼ d 
        Where a.����id=b.����id And b.id=c.ҽ��id And b.�Һŵ�=d.no(+) And b.���ID Is Null 
              And a.����id=b.����id And b.id=ҽ��ID_In; 
 
      Return v_return; 
  End Get_ZLHIS_PACS_003; 
 
 
  --ZLHIS_PACS_004(��鱨�泷��) 
  Function Get_ZLHIS_PACS_004 Return varchar2 As 
    v_return varchar2(4000); 
  Begin 
    If ��Ϣ���_In Is Null Then
        Select 
           '<patient_info>' || 
               '<patient_id>' || a.����id || '</patient_id>' || 
               '<patient_name>' || v_���� || '</patient_name>' || 
               '<in_number>' || nvl(a.סԺ��,0) || '</in_number>' || 
               '<out_number>' || nvl(a.�����,0) || '</out_number>' || 
           '</patient_info>' || 
           '<patient_clinic>' || 
               '<patient_source>' || b.������Դ || '</patient_source>' || 
               '<clinic_id>' || Case b.������Դ When 1 Then nvl(e.id,0) When 2 Then nvl(b.��ҳid, 0) Else 0 End || '</clinic_id>' || 
               '<clinic_dept_id>' || nvl(a.��ǰ����id,0) || '</clinic_dept_id>' || 
           '</patient_clinic>' || 
           '<advice_info>' || 
               '<order_id>' || b.id || '</order_id>' || 
               '<cur_state>' || nvl(c.ִ�й���,0) || '</cur_state>' || 
               '<check_item_id>' || b.������Ŀid || '</check_item_id>' || 
               '<check_item_title>' || b.ҽ������ || '</check_item_title>' || 
               '<create_doctor>' || b.����ҽ�� || '</create_doctor>' || 
               '<create_dept_id>' || nvl(b.��������id,0) || '</create_dept_id>' || 
               '<study_execute_id>' || nvl(b.ִ�п���id,0) || '</study_execute_id>' || 
           '</advice_info>' || 
           '<report_info>' || 
               '<report_id>' || d.����id || '</report_id>' || 
           '</report_info>'  Into v_return 
 
        From ������Ϣ a, ����ҽ����¼ b, ����ҽ������ c, ����ҽ������ d, ���˹Һż�¼ e 
        Where a.����id=b.����id And b.id=c.ҽ��id And b.id=d.ҽ��id And b.�Һŵ�=e.no(+) And b.���ID Is Null 
              And d.��鱨��id Is Null And a.����id=b.����id And b.id=ҽ��ID_In; 
    Else
        Select 
           '<patient_info>' || 
               '<patient_id>' || a.����id || '</patient_id>' || 
               '<patient_name>' || v_���� || '</patient_name>' || 
               '<in_number>' || nvl(a.סԺ��,0) || '</in_number>' || 
               '<out_number>' || nvl(a.�����,0) || '</out_number>' || 
           '</patient_info>' || 
           '<patient_clinic>' || 
               '<patient_source>' || b.������Դ || '</patient_source>' || 
               '<clinic_id>' || Case b.������Դ When 1 Then nvl(e.id,0) When 2 Then nvl(b.��ҳid, 0) Else 0 End || '</clinic_id>' || 
               '<clinic_dept_id>' || nvl(a.��ǰ����id,0) || '</clinic_dept_id>' || 
           '</patient_clinic>' || 
           '<advice_info>' || 
               '<order_id>' || b.id || '</order_id>' || 
               '<cur_state>' || nvl(c.ִ�й���,0) || '</cur_state>' || 
               '<check_item_id>' || b.������Ŀid || '</check_item_id>' || 
               '<check_item_title>' || b.ҽ������ || '</check_item_title>' || 
               '<create_doctor>' || b.����ҽ�� || '</create_doctor>' || 
               '<create_dept_id>' || nvl(b.��������id,0) || '</create_dept_id>' || 
               '<study_execute_id>' || nvl(b.ִ�п���id,0) || '</study_execute_id>' || 
           '</advice_info>' || 
           '<report_info>' || 
               '<report_id>' || d.��鱨��id || '</report_id>' || 
           '</report_info>'  Into v_return 
 
        From ������Ϣ a, ����ҽ����¼ b, ����ҽ������ c, ����ҽ������ d, ���˹Һż�¼ e 
        Where a.����id=b.����id And b.id=c.ҽ��id And b.id=d.ҽ��id And b.�Һŵ�=e.no(+) And b.���ID Is Null 
              And d.����id Is Null And a.����id=b.����id And b.id=ҽ��ID_In And d.��鱨��id=��Ϣ���_In; 
    End If;
    
    Return v_return; 
  End Get_ZLHIS_PACS_004; 
  
  --ZLHIS_PACS_005(���Σ��ֵ֪ͨ) 
  Function Get_ZLHIS_PACS_005 Return varchar2 As 
    v_return varchar2(4000); 
  Begin 
        With t As (Select id, ���� From ��Ա��) 
        Select 
           '<patient_info>' || 
               '<patient_id>' || a.����id || '</patient_id>' || 
               '<patient_name>' || v_���� || '</patient_name>' || 
               '<in_number>' || nvl(a.סԺ��,0) || '</in_number>' || 
               '<out_number>' || nvl(a.�����,0) || '</out_number>' || 
           '</patient_info>' || 
           '<patient_clinic>' || 
               '<patient_source>' || b.������Դ || '</patient_source>' || 
               '<clinic_id>' || Case b.������Դ When 1 Then nvl(i.id,0) When 2 Then nvl(b.��ҳid, 0) Else 0 End || '</clinic_id>' || 
               '<clinic_area_id>' || nvl(a.��ǰ����id,0) || '</clinic_area_id>' || 
               '<clinic_dept_id>' || nvl(b.��������id,0) || '</clinic_dept_id>' || 
               '<in_doctor_id>' || nvl(e.id,0) || '</in_doctor_id>' || 
               '<director_doctor_id>' || nvl(g.id,0) || '</director_doctor_id>' || 
               '<treat_doctor_id>' || nvl(h.id,0) || '</treat_doctor_id>' || 
               '<duty_nurse_id>' || nvl(f.id,0) || '</duty_nurse_id>' || 
           '</patient_clinic>' || 
           '<check_order>' || 
               '<order_id>' || b.id || '</order_id>' || 
               '<check_item_id>' || b.������Ŀid || '</check_item_id>' || 
               '<check_item_title>' || b.ҽ������ || '</check_item_title>' || 
               '<study_execute_id>' || nvl(b.ִ�п���id,0) || '</study_execute_id>' || 
           '</check_order>'  Into v_return 
 
        From ������Ϣ a, ����ҽ����¼ b, ������ҳ c, ���˹Һż�¼ i, 
             (Select a.����ID,a.��ҳID,����ҽʦ,����ҽʦ From ���˱䶯��¼ a,����ҽ����¼ b 
              Where a.����id=b.����id And a.��ʼԭ�� Is Not Null And a.��ֹʱ�� Is Null And b.id=0) d, 
              t e, t f, t g, t h 
        Where b.����id = a.����id And b.�Һŵ�=i.no(+) 
              And b.����id=c.����id(+) And b.��ҳid=c.��ҳid(+) 
              And c.סԺҽʦ=e.����(+) And c.���λ�ʿ=f.����(+) 
              And c.����id =d.����id(+) And c.��ҳid=d.��ҳid(+) 
              And d.����ҽʦ=g.����(+) And d.����ҽʦ=h.����(+) 
              And b.���id Is Null And a.����id=b.����id And b.id=ҽ��ID_In; 
 
      Return v_return; 
  End Get_ZLHIS_PACS_005;
  
  --ZLHIS_PACS_006(���ԤԼ֪ͨ) 
  Function Get_ZLHIS_PACS_006 Return varchar2 As 
    v_return varchar2(4000); 
  Begin 
        Select 
           '<patient_info>' || 
               '<patient_id>' || a.����id || '</patient_id>' || 
               '<patient_name>' || v_���� || '</patient_name>' || 
               '<in_number>' || nvl(a.סԺ��,0) || '</in_number>' || 
               '<out_number>' || nvl(a.�����,0) || '</out_number>' || 
           '</patient_info>' || 
           '<patient_clinic>' || 
               '<patient_source>' || b.������Դ || '</patient_source>' || 
               '<clinic_id>' || Case b.������Դ When 1 Then nvl(i.id,0) When 2 Then nvl(b.��ҳid, 0) Else 0 End || '</clinic_id>' || 
               '<clinic_area_id>' || nvl(a.��ǰ����id,0) || '</clinic_area_id>' || 
               '<clinic_dept_id>' || nvl(b.��������id,0) || '</clinic_dept_id>' || 
           '</patient_clinic>' || 
           '<check_order>' || 
               '<order_id>' || b.id || '</order_id>' || 
               '<check_item_id>' || b.������Ŀid || '</check_item_id>' || 
               '<check_item_title>' || b.ҽ������ || '</check_item_title>' || 
               '<study_execute_id>' || nvl(b.ִ�п���id,0) || '</study_execute_id>' ||
               '<schedult_id>' || k.ԤԼid || '</schedult_id>' ||
               '<machine_id>' || nvl(k.����豸id,0) || '</machine_id>' ||
               '<machine_name>' || nvl(k.����豸����,'') || '</machine_name>' ||
               '<schedule_date>' || To_Char(k.ԤԼ����, 'YYYY-MM-DD HH24:MI:SS') || '</schedule_date>' ||
               '<schedule_begin_time>' || To_Char(k.ԤԼ��ʼʱ��, 'YYYY-MM-DD HH24:MI:SS') || '</schedule_begin_time>' ||
               '<schedule_end_time>' || To_Char(k.ԤԼ����ʱ��, 'YYYY-MM-DD HH24:MI:SS')  || '</schedule_end_time>' ||
               '<schedule_sec_begin>' || To_Char(k.ԤԼ��ʼʱ���, 'YYYY-MM-DD HH24:MI:SS')  || '</schedule_sec_begin>' ||
               '<schedule_sec_end>' || To_Char(k.ԤԼ����ʱ���, 'YYYY-MM-DD HH24:MI:SS')  || '</schedule_sec_end>' ||
               '<schedule_call_no>' || nvl(k.���,0) || '</schedule_call_no>' || 
           '</check_order>'  Into v_return 
 
        From ������Ϣ a, ����ҽ����¼ b, ������ҳ c, ���˹Һż�¼ i, Ris���ԤԼ k 
        Where b.����id = a.����id And b.�Һŵ�=i.no(+) 
              And b.����id=c.����id(+) And b.��ҳid=c.��ҳid(+) 
              And b.id =k.ҽ��id 
              And b.���id Is Null And a.����id=b.����id And b.id=ҽ��ID_In; 
 
      Return v_return; 
  End Get_ZLHIS_PACS_006; 
 
Begin 
  v_Context := ''; 
  
  --�����ж��Ƿ���Ӥ����������v_���� ��ȡӤ������������v_���� Ϊ����ҽ����¼��������
  Select Max(Ӥ��), Max(��ҳid) Into n_Ӥ�����, n_��ҳid From ����ҽ����¼ Where ID = ҽ��id_In;
  If n_Ӥ����� > 0 And n_��ҳid > 0 Then
    Select Nvl(b.Ӥ������, a.���� || '֮��' || Trim(To_Char(b.���, '9')))
    Into v_����
    From ����ҽ����¼ A, ������������¼ B
    Where a.����id = b.����id And b.��ҳid = n_��ҳid And b.��� = n_Ӥ����� And a.Id = ҽ��id_In;
  Else
    Select ���� Into v_���� From ����ҽ����¼
    Where Id = ҽ��id_In;
  End If;
 
  Case ��Ϣ����_In 
    When 'ZLHIS_CIS_005' Then 
      --ZLHIS_CIS_005(ҽ��ִ�а������) 
      v_Context := Get_ZLHIS_CIS_005; 
 
    When 'ZLHIS_CIS_015' Then 
        --ZLHIS_CIS_015(ҽ���ܾ�ִ��) 
        v_Context := Get_ZLHIS_CIS_015; 
        
    When 'ZLHIS_CIS_017' Then 
        --ZLHIS_CIS_017(���߼������) 
        v_Context := Get_ZLHIS_CIS_017; 
 
    When 'ZLHIS_CIS_024' Then 
        --ZLHIS_PACS_024(����ҽ������) 
        v_Context := Get_ZLHIS_CIS_024; 
 
    When 'ZLHIS_PACS_001' Then 
        --ZLHIS_PACS_001(��鱨�����) 
        v_Context := Get_ZLHIS_PACS_001; 
 
    When 'ZLHIS_PACS_002' Then 
        --ZLHIS_PACS_002(���״̬ͬ��) 
        v_Context := Get_ZLHIS_PACS_002; 
 
    When 'ZLHIS_PACS_003' Then 
        --ZLHIS_PACS_003(���״̬����) 
        v_Context := Get_ZLHIS_PACS_003; 
 
    When 'ZLHIS_PACS_004' Then 
        --ZLHIS_PACS_004(��鱨�泷��) 
        v_Context := Get_ZLHIS_PACS_004; 
 
    When 'ZLHIS_PACS_005' Then 
        --ZLHIS_PACS_005(���Σ��ֵ֪ͨ) 
        v_Context := Get_ZLHIS_PACS_005; 
    
    When 'ZLHIS_PACS_006' Then 
        --ZLHIS_PACS_006(���ԤԼ֪ͨ) 
        v_Context := Get_ZLHIS_PACS_006; 
    Else 
      Return ''; 
  End Case; 
 
  Return v_Context; 
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End zl_Ӱ����Ϣ_XML���ݻ�ȡ;
/

--117999:���ϴ�,2017-12-12,�������״̬���
Create Or Replace Procedure Zl_���������Һ�_Insert
(
  ������ʽ_In     Integer,
  ����id_In       ������ü�¼.����id%Type,
  ����_In         �ҺŰ���.����%Type,
  ����_In         �Һ����״̬.���%Type,
  ���ݺ�_In       ������ü�¼.No%Type,
  Ʊ�ݺ�_In       ������ü�¼.ʵ��Ʊ��%Type,
  ���㷽ʽ_In     Varchar2,
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
  ��������_In     Number := 0,
  ������_In     Number := 0,
  �����¼id_In   �ٴ������¼.Id%Type := Null
) As
  --���ܣ������������йҺ�(����ԤԼ;ԤԼ�ҺŲ��ۿ�;ԤԼ�Һſۿ�)
  --���:������ʽ_IN:1-��ʾ�Һ�,2-��ʾԤԼ�ҺŲ��ۿ�,3-��ʾԤԼ�Һ�,�ۿ�
  --      ���㷽ʽ_IN:֧�ֶ��ֽ��㷽ʽ,���ֽ��㷽ʽʱ�������ʽ����:���㷽ʽ����1,���,�������,��������־|���㷽ʽ����2,���,�������,��������־|...
  --      �������״̬_In:1-��ʾǿ�Ƽ���Һ����״̬����;�������������Ż�����ʱ��ʱ�ż���.
  --      �Ƿ������豸_In:1-��ʾ��ҽԺ�������豸���д˺����ĵ���,�����豸���ô˺��� ����Ӻ�,��������
  --      ��������_In :0-������������ 1-��ʾ�Ե��ݽ�������,����δ��Ч�ĵ�����Ϣ;2-�������ļ�¼���н���-������������:δ��Ч�ĵ��������пۿ���ɺ���н���
  --      ���ս���_IN:��ʽ="���㷽ʽ|������||....."
  Err_Item Exception;
  Err_Special Exception;
  v_Err_Msg            Varchar2(255);
  n_��ӡid             Ʊ�ݴ�ӡ����.Id%Type;
  n_����ֵ             ����Ԥ����¼.���%Type;
  v_�ŶӺ���           Varchar2(20);
  v_��������           �ŶӽкŶ���.��������%Type;
  n_Ԥ��id             ����Ԥ����¼.Id%Type;
  n_�Һ�id             ���˹Һż�¼.Id%Type;
  v_��������           Varchar2(3000);
  v_��ǰ����           Varchar2(150);
  d_����ʱ��           Date;
  v_���㷽ʽ           ����Ԥ����¼.���㷽ʽ%Type;
  n_������           ����Ԥ����¼.��Ԥ��%Type;
  n_����ϼ�           Number(16, 5);
  n_Ԥ�����           ����Ԥ����¼.��Ԥ��%Type;
  n_��id               ����ɿ����.Id%Type;
  d_�Ŷ�ʱ��           Date;
  n_����               Number;
  n_����ԤԼ������     Number(18);
  n_��Լ����           Number(18);
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
  v_�շ���Ŀids        Varchar2(300);
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
  v_�Ŷ����           �ŶӽкŶ���.�Ŷ����%Type;
  v_������             �Һ����״̬.������%Type;
  v_��Ų���Ա         �Һ����״̬.����Ա����%Type;
  v_��Ż�����         �Һ����״̬.������%Type;
  n_�������           Number := 0;
  n_������id           �շ��ض���Ŀ.�շ�ϸĿid%Type;
  v_���ʽ           ���˹Һż�¼.ҽ�Ƹ��ʽ%Type;
  v_�ѱ�               ������ü�¼.�ѱ�%Type;
  n_���ηѱ�           Number(3) := 0;
  n_Tmp����id          �ҺŰ���.Id%Type;
  n_�ƻ�id             �ҺŰ��żƻ�.Id%Type;
  v_����               ������Ϣ.����%Type;
  n_������λ������ģʽ Number;
  n_�����¼id         �ٴ������¼.Id%Type;
  n_�Һ�ģʽ           Number(3);
  n_ͬ���޺���         Number;
  n_ͬ����Լ��         Number;
  n_���˹Һſ�����     Number;
  n_��ʱ����ʾ         Number;
  d_����ʱ��           Date;
  n_Exists             Number;
  v_Para               Varchar2(2000);
  n_ר�ҺŹҺ�����     Number;
  n_ר�Һ�ԤԼ����     Number;
  v_ʱ���             ʱ���.ʱ���%Type;
  d_��鿪ʼʱ��       ʱ���.��ʼʱ��%Type;
  d_������ʱ��       ʱ���.��ֹʱ��%Type;

  Cursor c_Pati(n_����id ������Ϣ.����id%Type) Is
    Select a.����id, a.����, a.�Ա�, a.����, a.סԺ��, a.�����, a.�ѱ�, a.����, c.���� As ���ʽ
    From ������Ϣ A, ҽ�Ƹ��ʽ C
    Where a.����id = n_����id And a.ҽ�Ƹ��ʽ = c.����(+);

  r_Pati c_Pati%RowType;

  --���α������շѳ�Ԥ���Ŀ���Ԥ���б�
  --��ID�������ȳ��ϴ�δ����ġ�
  Cursor c_Deposit(v_����id ������Ϣ.����id%Type) Is
    Select NO, Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) As ���, Min(��¼״̬) As ��¼״̬, Nvl(Max(����id), 0) As ����id,
           Max(Decode(��¼����, 1, ID, 0) * Decode(��¼״̬, 2, 0, 1)) As ԭԤ��id
    From ����Ԥ����¼
    Where ��¼���� In (1, 11) And ����id = v_����id And Nvl(Ԥ�����, 2) = 1 Having Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0
    Group By NO
    Order By ����id, NO;

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
                           p.���� As ��, p.���� As ��, p.���� As ��, p.��ſ���, p.�ƻ�id
           From (Select p.Id, p.����, p.����, p.����id, p.��Ŀid, p.ҽ��id, p.ҽ������, b.�޺���, b.��Լ��, Nvl(p.��������, 0) As ��������, p.����, p.��һ,
                         p.�ܶ�, p.����, p.����, p.����, p.����, p.���﷽ʽ, p.��ſ���,
                         Decode(To_Char(d_����ʱ��_In, 'D'), '1', p.����, '2', p.��һ, '3', p.�ܶ�, '4', p.����, '5', p.����, '6', p.����,
                                 '7', p.����, Null) As �Ű�, Null As �ƻ�id
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
                                 '7', p.����, Null) As �Ű�, p.Id As �ƻ�id
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

  Procedure Zl_���������Һ�_����_Insert
  (
    ��¼id_In       �ٴ������¼.Id%Type,
    ������ʽ_In     Integer,
    ����id_In       ������ü�¼.����id%Type,
    ����_In         �ҺŰ���.����%Type,
    ����_In         �Һ����״̬.���%Type,
    ���ݺ�_In       ������ü�¼.No%Type,
    Ʊ�ݺ�_In       ������ü�¼.ʵ��Ʊ��%Type,
    ���㷽ʽ_In     Varchar2,
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
    ��������_In     Number := 0,
    ������_In     Number := 0
  ) As
    --���ܣ������������йҺ�(����ԤԼ;ԤԼ�ҺŲ��ۿ�;ԤԼ�Һſۿ�),������Ű�ģʽ��ʹ��
    --���: ������ʽ_IN:1-��ʾ�Һ�,2-��ʾԤԼ�ҺŲ��ۿ�,3-��ʾԤԼ�Һ�,�ۿ�
    --      �������״̬_In:1-��ʾǿ�Ƽ���Һ����״̬����;�������������Ż�����ʱ��ʱ�ż���.
    --      �Ƿ������豸_In:1-��ʾ��ҽԺ�������豸���д˺����ĵ���,�����豸���ô˺��� ����Ӻ�,��������
    --      ��������_In :0-������������ 1-��ʾ�Ե��ݽ�������,����δ��Ч�ĵ�����Ϣ;2-�������ļ�¼���н���-������������:δ��Ч�ĵ��������пۿ���ɺ���н���
    --      ���ս���_IN:��ʽ="���㷽ʽ|������||....."
    --      ��Ԥ������ids_In:����ö��ַ���,��Ԥ��ʱ��Ч(��Ԥ�������ҵ���������һ��),��Ҫ��ʹ�ü�����Ԥ����
    Err_Item Exception;
    Err_Special Exception;
    v_Err_Msg            Varchar2(255);
    n_��ӡid             Ʊ�ݴ�ӡ����.Id%Type;
    n_����ֵ             ����Ԥ����¼.���%Type;
    v_�ŶӺ���           Varchar2(20);
    v_��������           �ŶӽкŶ���.��������%Type;
    n_Ԥ��id             ����Ԥ����¼.Id%Type;
    n_�Һ�id             ���˹Һż�¼.Id%Type;
    v_��������           Varchar2(3000);
    v_��ǰ����           Varchar2(150);
    v_���㷽ʽ           ����Ԥ����¼.���㷽ʽ%Type;
    n_������           ����Ԥ����¼.��Ԥ��%Type;
    n_����ϼ�           Number(16, 5);
    n_Ԥ�����           ����Ԥ����¼.��Ԥ��%Type;
    n_��id               ����ɿ����.Id%Type;
    d_�Ŷ�ʱ��           Date;
    n_����               Number;
    n_����ԤԼ������     Number(18);
    n_��Լ����           Number(18);
    d_����ʱ��           Date;
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
    n_ʵ�ս��ϼ�       ������ü�¼.ʵ�ս��%Type;
    n_��������id         ������ü�¼.��������id%Type;
    n_ʵ�ս��           ������ü�¼.ʵ�ս��%Type;
    n_Ӧ�ս��           ������ü�¼.ʵ�ս��%Type;
    n_����id             ���˽��ʼ�¼.Id%Type;
    v_Temp               Varchar2(500);
    v_���㷽ʽ��¼       Varchar2(1000);
    n_ԤԼʱ�����       Number;
    n_��ſ���           �ٴ������¼.�Ƿ���ſ���%Type;
    n_��Լ��             �ٴ������¼.��Լ��%Type;
    n_��Ŀid             �ٴ������¼.��Ŀid%Type;
    n_����id             �ٴ������¼.����id%Type;
    d_��ֹʱ��           �ٴ������¼.��ֹʱ��%Type;
    v_ҽ������           �ٴ������¼.ҽ������%Type;
    n_ҽ��id             �ٴ������¼.ҽ��id%Type;
    n_ԤԼ˳���         �ٴ�������ſ���.ԤԼ˳���%Type;
    n_ԤԼ����           Number;
    d_ʱ�ο�ʼʱ��       Date;
    d_ʱ����ֹʱ��       Date;
    v_�շ���Ŀids        Varchar2(300);
    n_��������־         Number;
    n_����               ���˹Һż�¼.����%Type;
    d_�Ǽ�ʱ��           Date;
    n_���ʽ��           ����Ԥ����¼.��Ԥ��%Type;
    v_�������           ����Ԥ����¼.�������%Type;
    v_����Ա���         ��Ա��.���%Type;
    v_����Ա����         ��Ա��.����%Type;
    n_ԤԼ               Integer;
    v_�ֽ�               ����Ԥ����¼.���㷽ʽ%Type;
    n_���÷�ʱ��         Integer;
    n_�ѹ���             ���˹ҺŻ���.�ѹ���%Type;
    n_��Լ��             ���˹ҺŻ���.��Լ��%Type;
    n_�����ѽ���         ���˹ҺŻ���.��Լ��%Type;
    n_ԤԼ���ɶ���       Number;
    n_�޺���             �ٴ������¼.�޺���%Type;
    d_Date               Date;
    n_�Һ����           Number;
    v_�Ŷ����           �ŶӽкŶ���.�Ŷ����%Type;
    v_������             �Һ����״̬.������%Type;
    v_��Ų���Ա         �Һ����״̬.����Ա����%Type;
    v_��Ż�����         �Һ����״̬.������%Type;
    n_�������           Number := 0;
    n_������id           �շ��ض���Ŀ.�շ�ϸĿid%Type;
    v_���ʽ           ���˹Һż�¼.ҽ�Ƹ��ʽ%Type;
    v_�ѱ�               ������ü�¼.�ѱ�%Type;
    n_���ηѱ�           Number(3) := 0;
    v_����               ������Ϣ.����%Type;
    n_������λ������ģʽ Number;
    n_ͬ���޺���         Number;
    n_��ʱ����ʾ         Number;
    n_ͬ����Լ��         Number;
    n_���˹Һſ�����     Number;
    n_Exists             Number(5);
    n_����ҽ��id         �ٴ������¼.����ҽ��id%Type;
    v_����ҽ������       �ٴ������¼.����ҽ������%Type;
    d_���￪ʼʱ��       �ٴ������¼.���￪ʼʱ��%Type;
    d_������ֹʱ��       �ٴ������¼.������ֹʱ��%Type;
    n_ר�ҺŹҺ�����     Number;
    n_ר�Һ�ԤԼ����     Number;
  
    Cursor c_Pati(n_����id ������Ϣ.����id%Type) Is
      Select a.����id, a.����, a.�Ա�, a.����, a.סԺ��, a.�����, a.�ѱ�, a.����, c.���� As ���ʽ
      From ������Ϣ A, ҽ�Ƹ��ʽ C
      Where a.����id = n_����id And a.ҽ�Ƹ��ʽ = c.����(+);
  
    r_Pati c_Pati%RowType;
  
    --���α������շѳ�Ԥ���Ŀ���Ԥ���б�
    --��ID�������ȳ��ϴ�δ����ġ�
    Cursor c_Deposit(v_����id ������Ϣ.����id%Type) Is
      Select NO, Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) As ���, Min(��¼״̬) As ��¼״̬, Nvl(Max(����id), 0) As ����id,
             Max(Decode(��¼����, 1, ID, 0) * Decode(��¼״̬, 2, 0, 1)) As ԭԤ��id
      From ����Ԥ����¼
      Where ��¼���� In (1, 11) And ����id = v_����id And Nvl(Ԥ�����, 2) = 1 Having Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0
      Group By NO
      Order By ����id, NO;
  
    Function Zl_����(��¼id_In �ٴ������¼.Id%Type) Return Varchar2 As
      n_���﷽ʽ �ٴ������¼.���﷽ʽ%Type;
      v_����     ���˹Һż�¼.����%Type;
      v_Rowid    Varchar2(500);
      n_Next     Integer;
      n_First    Integer;
    Begin
    
      If ��������_In = 2 Then
        --�Ե��ݽ��н���,���ȼ���Ƿ��������
        Select Count(Rowid)
        Into n_����
        From ���˹Һż�¼ Where����¼״̬ = 0 And NO = ���ݺ�_In And ����id = ����id_In;
        If n_���� = 0 Then
          v_Err_Msg := '���ݺ�Ϊ(' || ���ݺ�_In || ')�ĵ���,�����ڻ����Ѿ�������!';
          Raise Err_Item;
        End If;
        Select Max(����) Into n_���� From ���˹Һż�¼ Where����¼״̬ = 0 And NO = ���ݺ�_In And ����id = ����id_In;
      End If;
    
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
                            Where Nvl(ִ��״̬, 0) = 0 And ����ʱ�� Between Trunc(Sysdate) And Sysdate And �ű� = ����_In And
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
    d_����ʱ�� := ����ʱ��_In;
  
    If d_����ʱ�� Is Null Then
      d_����ʱ�� := Sysdate;
    End If;
  
    Begin
      Select ���� Into v_�ֽ� From ���㷽ʽ Where ���� = 1;
    Exception
      When Others Then
        v_�ֽ� := '�ֽ�';
    End;
  
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
    Where �����¼id = ��¼id_In And �Ǽ�ʱ�� Between Trunc(����ʱ��_In) And Trunc(����ʱ��_In + 1) - 1 / 24 / 60 / 60;
  
    If �Ǽ�ʱ��_In Is Null Then
      d_�Ǽ�ʱ�� := Sysdate;
    Else
      d_�Ǽ�ʱ�� := �Ǽ�ʱ��_In;
    End If;
    If Trunc(Sysdate) > Trunc(����ʱ��_In) Then
      v_Err_Msg := '���ܹ���ǰ�ĺ�(' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || ')��';
      Raise Err_Item;
    End If;
  
    v_Temp           := Nvl(zl_GetSysParameter('����ͬ���޹�N����', 1111), '0|0') || '|';
    n_ͬ���޺���     := To_Number(Substr(v_Temp, 1, Instr(v_Temp, '|') - 1));
    n_ͬ����Լ��     := To_Number(Nvl(zl_GetSysParameter('����ͬ����ԼN����', 1111), '0'));
    n_����ԤԼ������ := To_Number(Nvl(zl_GetSysParameter('����ԤԼ������', 1111), '0'));
    n_���˹Һſ����� := To_Number(Nvl(zl_GetSysParameter('���˹Һſ�������', 1111), '0'));
    n_ר�ҺŹҺ����� := To_Number(Nvl(zl_GetSysParameter('ר�ҺŹҺ�����'), '0'));
    n_ר�Һ�ԤԼ���� := To_Number(Nvl(zl_GetSysParameter('ר�Һ�ԤԼ����'), '0'));
  
    --����ID,��������;��ԱID,��Ա���,��Ա����
    v_Temp := Zl_Identity(0);
    If Nvl(v_Temp, ' ') = ' ' Then
      v_Err_Msg := '��ǰ������Աδ���ö�Ӧ����Ա��ϵ,���ܼ�����';
      Raise Err_Item;
    End If;
    n_��������id := To_Number(Zl_����Ա(0, v_Temp));
    v_����Ա��� := Zl_����Ա(1, v_Temp);
    v_����Ա���� := Zl_����Ա(2, v_Temp);
    n_��id       := Zl_Get��id(v_����Ա����);
  
    --֧���������ύ���
    Select Nvl(Max(1), 0)
    Into n_Exists
    From ���˹Һż�¼
    Where ����id = ����id_In And �ű� = ����_In And ���� = ����_In And ����Ա���� = v_����Ա���� And Nvl(��¼id_In, 0) = Nvl(�����¼id, 0) And
          �Ǽ�ʱ�� > Sysdate - 0.01 And ��¼״̬ = 1 And ����ʱ�� = ����ʱ��_In;
    If n_Exists = 1 Then
      v_Err_Msg := '�����Ѿ��Һ�,�����ظ�����ͬ�ĺţ�';
      Raise Err_Special;
    End If;
  
    If ������ʽ_In <> 1 Then
      --ԤԼ����Ƿ���Ӻ�����λ����
      --��������˺�����λ���� ��
      Begin
        Select 1
        Into n_������λ����
        From �ٴ�����Һſ��Ƽ�¼
        Where ��¼id = ��¼id_In And ���� = 1 And ���� = 1 And ���Ʒ�ʽ <> 4 And Rownum < 2;
      Exception
        When Others Then
          n_������λ���� := 0;
      End;
    End If;
  
    If ������ʽ_In <> 2 Then
      v_���� := Zl_����(��¼id_In);
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
  
    Begin
      Select Nvl(�Ƿ��ʱ��, 0), �޺���, �ѹ���, �����ѽ���, ��Լ��, �Ƿ���ſ���, ��Լ��, ��Ŀid, ����id, ҽ��id, ҽ������, ����ҽ��id, ����ҽ������, ���￪ʼʱ��, ������ֹʱ��
      Into n_���÷�ʱ��, n_�޺���, n_�ѹ���, n_�����ѽ���, n_��Լ��, n_��ſ���, n_��Լ��, n_��Ŀid, n_����id, n_ҽ��id, v_ҽ������, n_����ҽ��id, v_����ҽ������,
           d_���￪ʼʱ��, d_������ֹʱ��
      From �ٴ������¼
      Where ID = ��¼id_In And Nvl(�Ƿ�����, 0) = 0;
    Exception
      When Others Then
        n_Count := -1;
    End;
    If n_Count = -1 Then
      v_Err_Msg := '�úű�û����' || To_Char(����ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '�н��а��š�';
      Raise Err_Item;
    End If;
  
    If ����ʱ��_In Between Nvl(d_���￪ʼʱ��, Sysdate) And Nvl(d_������ֹʱ��, Sysdate - 1) And v_����ҽ������ Is Not Null Then
      n_ҽ��id   := n_����ҽ��id;
      v_ҽ������ := v_����ҽ������;
    End If;
  
    --�Բ������ƽ��м��
    --����ԤԼ���ۿ�ʱ���м��
    If ������ʽ_In = 2 Then
      If Nvl(n_ͬ����Լ��, 0) <> 0 Or Nvl(n_����ԤԼ������, 0) <> 0 Then
        n_��Լ���� := 0;
        For c_Chkitem In (Select Distinct ִ�в���id
                          From ���˹Һż�¼
                          Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 2 And ԤԼʱ�� Between Trunc(����ʱ��_In) And
                                Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ִ�в���id <> n_����id) Loop
          n_��Լ���� := n_��Լ���� + 1;
        End Loop;
        If n_��Լ���� >= Nvl(n_����ԤԼ������, 0) And Nvl(n_����ԤԼ������, 0) > 0 Then
          v_Err_Msg := 'ͬһ�������ͬʱ��ԤԼ[' || Nvl(n_����ԤԼ������, 0) || ']������,������ԤԼ��';
          Raise Err_Item;
        End If;
      
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 2 And ԤԼʱ�� Between Trunc(����ʱ��_In) And
              Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ִ�в���id = n_����id;
        If n_Count >= Nvl(n_ͬ����Լ��, 0) And Nvl(n_ͬ����Լ��, 0) > 0 Then
          v_Err_Msg := '�ò����Ѿ��ڸÿ���ԤԼ��' || n_Count || '��,������ԤԼ��';
          Raise Err_Item;
        End If;
      End If;
      If Nvl(n_ר�Һ�ԤԼ����, 0) <> 0 Then
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 2 And �����¼id = ��¼id_In;
        If n_Count >= Nvl(n_ר�Һ�ԤԼ����, 0) And Nvl(n_ר�Һ�ԤԼ����, 0) > 0 Then
          v_Err_Msg := '�ò����Ѿ���������ԤԼ����,������ԤԼ��';
          Raise Err_Item;
        End If;
      End If;
    Else
      If Nvl(n_ͬ���޺���, 0) <> 0 Or Nvl(n_���˹Һſ�����, 0) <> 0 Then
        n_��Լ���� := 0;
        For c_Chkitem In (Select Distinct ִ�в���id
                          From ���˹Һż�¼
                          Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 1 And ����ʱ�� Between Trunc(����ʱ��_In) And
                                Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ִ�в���id <> n_����id) Loop
          n_��Լ���� := n_��Լ���� + 1;
        End Loop;
        If n_��Լ���� >= Nvl(n_���˹Һſ�����, 0) And Nvl(n_���˹Һſ�����, 0) > 0 Then
          v_Err_Msg := 'ͬһ�������ͬʱ�ܹҺ�[' || Nvl(n_���˹Һſ�����, 0) || ']������,�����ٹҺţ�';
          Raise Err_Item;
        End If;
      
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 1 And ����ʱ�� Between Trunc(����ʱ��_In) And
              Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ִ�в���id = n_����id;
        If n_Count >= Nvl(n_ͬ���޺���, 0) And Nvl(n_ͬ���޺���, 0) > 0 Then
          v_Err_Msg := '�ò����Ѿ��ڸÿ��ҹҺ���' || n_Count || '��,�����ٹҺţ�';
          Raise Err_Item;
        End If;
      End If;
      If Nvl(n_ר�ҺŹҺ�����, 0) <> 0 Then
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 1 And �����¼id = ��¼id_In;
        If n_Count >= Nvl(n_ר�ҺŹҺ�����, 0) And Nvl(n_ר�ҺŹҺ�����, 0) > 0 Then
          v_Err_Msg := '�ò����Ѿ��������ŹҺ�����,�����ٹҺţ�';
          Raise Err_Item;
        End If;
      End If;
    End If;
  
    d_Date         := Null;
    d_ʱ�ο�ʼʱ�� := Null;
  
    If Nvl(n_�޺���, 0) >= 0 Or n_�޺��� Is Null Then
      If n_���÷�ʱ�� = 1 Then
        If Nvl(n_��ſ���, 0) = 1 Then
          If Nvl(�Ƿ������豸_In, 0) = 0 Then
            Select Count(*), Max(��ʼʱ��)
            Into n_Count, d_ʱ�ο�ʼʱ��
            From �ٴ�������ſ���
            Where ��¼id = ��¼id_In And ��� = Nvl(����_In, 0);
          
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
                                To_Date(To_Char(Sysdate + Decode(Sign(��ʼʱ�� - ��ֹʱ��), 1, 1, 0), 'yyyy-mm-dd') || ' ' ||
                                         To_Char(��ֹʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ��ֹʱ��, ����, �Ƿ�ԤԼ
                         From �ٴ�������ſ���
                         Where ��¼id = ��¼id_In And ��� = Nvl(����_In, 0)) Loop
              If Sysdate > v_ʱ��.��ֹʱ�� Then
                v_Err_Msg := '�ű�Ϊ' || ����_In || '������Ϊ' || Nvl(����_In, 0) || '�İ���,�Ѿ�����ʱ��,������' || v_Temp || '��';
                Raise Err_Item;
              End If;
            End Loop;
          End If;
        Elsif ������ʽ_In > 1 Then
          --δ������ŵ�,��Ҫ���ԤԼ�����
          n_Count := 0;
          For v_ʱ�� In (Select ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ
                       From �ٴ�������ſ���
                       Where ��¼id = ��¼id_In And
                             (('3000-01-10 ' || To_Char(����ʱ��_In, 'HH24:MI:SS') Between
                             Decode(Sign(��ʼʱ�� - ��ֹʱ�� - 1 / 24 / 60 / 60), 1,
                                      '3000-01-09 ' || To_Char(��ʼʱ��, 'HH24:MI:SS'),
                                      '3000-01-10 ' || To_Char(��ʼʱ��, 'HH24:MI:SS')) And
                             '3000-01-10 ' || To_Char(��ֹʱ�� - 1 / 24 / 60 / 60, 'HH24:MI:SS')) Or
                             ('3000-01-10 ' || To_Char(����ʱ��_In, 'HH24:MI:SS') Between
                             '3000-01-10 ' || To_Char(��ʼʱ��, 'HH24:MI:SS') And
                             Decode(Sign(��ʼʱ�� - ��ֹʱ�� - 1 / 24 / 60 / 60), 1,
                                      '3000-01-11 ' || To_Char(��ֹʱ�� - 1 / 24 / 60 / 60, 'HH24:MI:SS'),
                                      '3000-01-10 ' || To_Char(��ֹʱ�� - 1 / 24 / 60 / 60, 'HH24:MI:SS'))))) Loop
            n_ԤԼʱ����� := v_ʱ��.���;
            d_ʱ�ο�ʼʱ�� := v_ʱ��.��ʼʱ��;
            d_ʱ����ֹʱ�� := v_ʱ��.��ֹʱ��;
          
            Select Count(*), Max(���), Max(ԤԼ˳���) + 1
            Into n_Count, n_ԤԼ����, n_ԤԼ˳���
            From �ٴ�������ſ���
            Where ��¼id = ��¼id_In And Nvl(�Һ�״̬, 0) Not In (0, 4, 5);
          
            If Nvl(n_Count, 0) > Nvl(v_ʱ��.����, 0) And ��������_In <> 2 Then
              v_Err_Msg := '�ű�Ϊ' || ����_In || '�ĹҺŰ�������' || To_Char(v_ʱ��.��ʼʱ��, 'hh24:mi:ss') || '��' ||
                           To_Char(v_ʱ��.��ֹʱ��, 'hh24:mi:ss') || '���ֻ��ԤԼ' || Nvl(v_ʱ��.����, 0) || '��,�����ٽ���ԤԼ�Һţ�';
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
      If n_�ѹ��� >= Nvl(n_�޺���, 0) And n_�޺��� Is Not Null Then
        v_Err_Msg := '�úű�����Ѵﵽ�޺��� ' || Nvl(n_�޺���, 0) || '�����ٹҺţ�';
        Raise Err_Item;
      End If;
    End If;
  
    If ������ʽ_In > 1 Then
      --ԤԼ����ؼ��
      --����:
      --   1.����Լ���ܳ�����Լ��
      --   2.����Ƿ�����ʱ�ε�
      If n_��Լ�� >= Nvl(n_��Լ��, 0) And Nvl(n_��Լ��, 0) <> 0 And n_��Լ�� Is Not Null And ��������_In <> 2 Then
        v_Err_Msg := '�úű��Ѵﵽ��Լ�� ' || Nvl(n_��Լ��, 0) || '������ԤԼ�Һţ�';
        Raise Err_Item;
      End If;
      If ԤԼ��ʽ_In Is Not Null Then
        Select To_Number(Substr(Zl_Fun_Get�ٴ�����ԤԼ״̬(��¼id_In, ����ʱ��_In, ����_In, ԤԼ��ʽ_In, NULL, 0, v_����Ա����, v_������), 1, 1))
        Into n_Exists
        From Dual;
        If n_Exists <> 0 Then
          v_Err_Msg := '�����ԤԼ��ʽ' || ԤԼ��ʽ_In || '������,���ܼ�����';
          Raise Err_Item;
        End If;
      End If;
    End If;
    If n_������λ���� > 0 And ������ʽ_In <> 1 And ������λ_In Is Not Null Then
      If Nvl(n_��ſ���, 0) = 1 And Nvl(����_In, 0) = 0 Then
        v_Err_Msg := '��ǰ����ʹ������ſ���,��ȷ������ҪԤԼ�����,���ܼ�����';
        Raise Err_Item;
      End If; --Nvl(r_����.��ſ���, 0) =0
    
      n_��� := Case
                When Nvl(n_��ſ���, 0) = 1 Or n_���÷�ʱ�� = 1 And ������ʽ_In > 1 Then
                 Nvl(����_In, 0)
                Else
                 0
              End;
    
      --������λ����ģʽ
      Select Nvl(���Ʒ�ʽ, 0)
      Into n_������λ������ģʽ
      From �ٴ�����Һſ��Ƽ�¼
      Where ��¼id = ��¼id_In And ���� = ������λ_In And ���� = 1 And ���� = 1 And Rownum < 2;
    
      If n_������λ������ģʽ = 0 Then
        v_Err_Msg := '��ǰ����(' || Nvl(����_In, 0) || 'δ����' || ������λ_In || '��ԤԼ,���ܼ�����';
        Raise Err_Item;
      End If;
      If n_������λ������ģʽ = 1 Or n_������λ������ģʽ = 2 Then
        Select ����
        Into n_Count
        From �ٴ�����Һſ��Ƽ�¼
        Where ��¼id = ��¼id_In And ���� = ������λ_In And ���� = 1 And ���� = 1;
        If n_������λ������ģʽ = 1 Then
          n_Count := Round(Nvl(n_��Լ��, n_�޺���) * n_Count / 100);
        End If;
        Select Count(1)
        Into n_Exists
        From ���˹Һż�¼
        Where ��¼״̬ = 1 And �����¼id = ��¼id_In And ������λ = ������λ_In;
        If n_Exists >= n_Count Then
          v_Err_Msg := '��ǰ����(' || Nvl(����_In, 0) || '�Ѿ�����' || ������λ_In || '��ԤԼ����,���ܼ�����';
          Raise Err_Item;
        End If;
      End If;
      --������ż��
      If n_������λ������ģʽ = 3 Then
        For c_������λ In (Select ���, ����
                       From �ٴ�����Һſ��Ƽ�¼
                       Where ��¼id = ��¼id_In And ���� = ������λ_In And ���� = 1 And ���� = 1 And ��� = ����_In) Loop
          If n_��ſ��� = 1 Then
            Begin
              Select 1
              Into n_Count
              From �ٴ�������ſ���
              Where ��¼id = ��¼id_In And ��� = ����_In And Nvl(�Һ�״̬, 0) = 0;
            Exception
              When Others Then
                n_Count := 0;
            End;
            If n_Count = 1 Then
              n_�Ƿ񿪷� := 1;
            Else
              v_Err_Msg := '��ǰ���(' || Nvl(����_In, 0) || '�Ѿ�����' || ������λ_In || '��ԤԼ����,���ܼ�����';
              Raise Err_Item;
            End If;
          Else
            Select Count(1)
            Into n_Count
            From �ٴ�������ſ���
            Where ��¼id = ��¼id_In And ��� = ����_In And ԤԼ˳��� Is Not Null And Nvl(�Һ�״̬, 0) <> 0;
            If n_Count >= c_������λ.���� Then
              v_Err_Msg := '��ǰ���(' || Nvl(����_In, 0) || '�Ѿ�����' || ������λ_In || '��ԤԼ����,���ܼ�����';
              Raise Err_Item;
            Else
              n_�Ƿ񿪷� := 1;
            End If;
          End If;
        End Loop;
      
        If Nvl(n_�Ƿ񿪷�, 0) = 0 Then
          v_Err_Msg := '��ǰ���(' || Nvl(����_In, 0) || 'δ����,���ܼ�����';
          Raise Err_Item;
        End If;
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
  
    If Nvl(������_In, 0) = 1 Then
      Begin
        Select �շ�ϸĿid Into n_������id From �շ��ض���Ŀ Where �ض���Ŀ = '������';
        v_�շ���Ŀids := n_��Ŀid || ',' || n_������id;
      Exception
        When Others Then
          v_Err_Msg := '����ȷ��������,�Һ�ʧ��!';
          Raise Err_Item;
      End;
    Else
      v_�շ���Ŀids := n_��Ŀid;
    End If;
  
    For c_Item In (Select 1 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, 1 As ����,
                          c.Id As ������Ŀid, c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, Null As ��������
                   From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C
                   Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = n_��Ŀid And d_����ʱ�� Between b.ִ������ And
                         Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))
                   Union All
                   Select 2 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, 1 As ����,
                          c.Id As ������Ŀid, c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, Null As ��������
                   From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C
                   Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = n_������id And d_����ʱ�� Between b.ִ������ And
                         Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))
                   Union All
                   Select 3 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, d.�������� As ����,
                          c.Id As ������Ŀid, c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, 1 As ��������
                   From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C, �շѴ�����Ŀ D
                   Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = d.����id And
                         d.����id In (Select Column_Value From Table(f_Str2list(v_�շ���Ŀids))) And d_����ʱ�� Between b.ִ������ And
                         Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))
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
           Decode(������ʽ_In, 2, To_Char(����_In), v_����), r_Pati.����id, r_Pati.�����, r_Pati.���ʽ, r_Pati.����, r_Pati.�Ա�,
           r_Pati.����, r_Pati.�ѱ�, n_����id, c_Item.���, ����_In, c_Item.��Ŀid, c_Item.������Ŀid, c_Item.�վݷ�Ŀ, 1, c_Item.����,
           c_Item.����, n_Ӧ�ս��, n_ʵ�ս��, Decode(������ʽ_In, 2, Null, n_ʵ�ս��), n_����id, 0, n_��������id, v_����Ա����,
           Decode(������ʽ_In, 2, v_����Ա����, Null), n_����id, v_ҽ������, v_����Ա���, v_����Ա����, ����ʱ��_In, d_�Ǽ�ʱ��, Null, 0, Null, Null,
           ժҪ_In, ԤԼ��ʽ_In, Decode(������ʽ_In, 2, Null, n_��id));
      End If;
      n_�к� := n_�к� + 1;
    
    End Loop;
  
    If Round(Nvl(�ҺŽ��ϼ�_In, 0), 5) <> Round(Nvl(n_ʵ�ս��ϼ�, 0), 5) Then
      v_Err_Msg := '���ιҺŽ���ȷ,��������ΪҽԺ�����˼۸�,�����»�ȡ�Һ��շ���Ŀ�ļ۸�,���ܼ�����';
      Raise Err_Item;
    End If;
  
    If n_���÷�ʱ�� = 1 Then
      d_Date := To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(d_ʱ�ο�ʼʱ��, 'hh24:mi:ss'),
                        'yyyy-mm-dd hh24:mi:ss');
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
      From �ٴ�������ſ���
      Where ��¼id = ��¼id_In And ��� = n_���� And Nvl(�Һ�״̬, 0) Not In (0, 5);
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count = 1 Then
      If n_���÷�ʱ�� = 0 And Nvl(n_��ſ���, 0) = 1 Then
        n_���� := Null;
      End If;
      If n_���÷�ʱ�� = 1 And Nvl(n_��ſ���, 0) = 1 Then
        v_Err_Msg := '��ǰ����ѱ�ʹ�ã�������ѡ��һ����ţ�';
        Raise Err_Item;
      End If;
    End If;
  
    If n_���÷�ʱ�� = 0 And Nvl(n_��ſ���, 0) = 1 And n_���� Is Null And ��������_In <> 2 Then
      Select Nvl(Min(���), 0)
      Into n_����
      From �ٴ�������ſ���
      Where ��¼id = ��¼id_In And Nvl(�Һ�״̬, 0) = 5 And ����Ա���� = v_����Ա���� And ����վ���� = v_������;
      If n_���� = 0 Then
        Select Nvl(Min(���), 0) Into n_���� From �ٴ�������ſ��� Where ��¼id = ��¼id_In And Nvl(�Һ�״̬, 0) = 0;
        If n_���� = 0 Then
          Select Nvl(Max(���), 0) + 1 Into n_���� From �ٴ�������ſ��� Where ��¼id = ��¼id_In;
        End If;
      End If;
    End If;
  
    If n_���÷�ʱ�� = 1 And ��������_In <> 2 Then
      If ������ʽ_In > 1 And Nvl(n_��ſ���, 0) = 0 Then
        --����:ԤԼʱ�����||ԤԼ��
        If Nvl(n_ԤԼ����, 0) = 0 Then
          v_Temp := Nvl(n_��Լ��, 0);
          v_Temp := LTrim(RTrim(v_Temp));
          v_Temp := LPad(Nvl(n_ԤԼ����, 0) + 1, Length(v_Temp), '0');
          v_Temp := n_ԤԼʱ����� || v_Temp;
          n_���� := To_Number(v_Temp);
        Else
          n_���� := n_ԤԼ���� + 1;
        End If;
      End If;
    End If;
  
    If Nvl(n_��ſ���, 0) = 1 Or (������ʽ_In > 1 And n_���÷�ʱ�� = 1) Or �������״̬_In = 1 Then
      --������ŵĴ���
      Begin
        Select ����Ա����, ����վ����
        Into v_��Ų���Ա, v_��Ż�����
        From �ٴ�������ſ���
        Where �Һ�״̬ = 5 And ��¼id = ��¼id_In And ��� = n_����;
        n_������� := 1;
      Exception
        When Others Then
          v_��Ų���Ա := Null;
          v_��Ż����� := Null;
          n_�������   := 0;
      End;
      If n_������� = 0 Then
        If n_���÷�ʱ�� = 1 And n_��ſ��� = 0 Then
          Insert Into �ٴ�������ſ���
            (��¼id, ���, ԤԼ˳���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, �Һ�״̬, ����, ����, ����Ա����, ��ע)
            Select ��¼id_In, n_ԤԼʱ�����, n_ԤԼ˳���, d_ʱ�ο�ʼʱ��, d_ʱ����ֹʱ��, 1, Decode(������ʽ_In, 1, 0, 1), Decode(������ʽ_In, 2, 2, 1),
                   1, ������λ_In, v_����Ա����, n_����
            From Dual;
        Else
          Update �ٴ�������ſ���
          Set �Һ�״̬ = Decode(������ʽ_In, 2, 2, 1), ����Ա���� = v_����Ա����
          Where ��¼id = ��¼id_In And ��� = n_����;
        End If;
        If Sql%RowCount = 0 Then
          Begin
            If n_���÷�ʱ�� = 1 Then
              --��ʱ��
              If n_��ſ��� = 1 Then
                --��ſ���
                Select Max(��ֹʱ��) Into d_��ֹʱ�� From �ٴ�������ſ��� Where ��¼id = ��¼id_In;
                If Sysdate > d_��ֹʱ�� Then
                  d_��ֹʱ�� := Sysdate;
                End If;
                Insert Into �ٴ�������ſ���
                  (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, �Һ�״̬, ����, ����, ����Ա����)
                  Select ��¼id_In, n_����, d_��ֹʱ��, d_��ֹʱ��, 1, Decode(������ʽ_In, 1, 0, 1), Decode(������ʽ_In, 2, 2, 1), 1,
                         ������λ_In, v_����Ա����
                  From Dual;
              Else
                --��ʱ��,����ſ���
                Null;
              End If;
            Else
              --����ʱ��
              Insert Into �ٴ�������ſ���
                (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, �Һ�״̬, ����, ����, ����Ա����)
                Select ��¼id_In, n_����, ��ʼʱ��, ��ֹʱ��, 1, Decode(������ʽ_In, 1, 0, 1), Decode(������ʽ_In, 2, 2, 1), 1, ������λ_In,
                       v_����Ա����
                From �ٴ�������ſ���
                Where ��¼id = ��¼id_In And ��� = 1;
            End If;
          Exception
            When Others Then
              v_Err_Msg := '���' || n_���� || '�ѱ�ʹ��,������ѡ��һ�����.';
              Raise Err_Item;
          End;
        End If;
      Else
        If v_����Ա���� <> v_��Ų���Ա Or v_������ <> v_��Ż����� Then
          v_Err_Msg := '���' || n_���� || '�ѱ�����' || v_������ || '����,������ѡ��һ�����.';
          Raise Err_Item;
        Else
          Update �ٴ�������ſ���
          Set �Һ�״̬ = Decode(������ʽ_In, 2, 2, 1), ����ʱ�� = Null
          Where ��¼id = ��¼id_In And ��� = n_���� And �Һ�״̬ = 5 And ����Ա���� = v_����Ա���� And ����վ���� = v_������;
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
              (n_Ԥ��id, 4, ���ݺ�_In, 1, Decode(����id_In, 0, Null, ����id_In), '���ս���', v_���㷽ʽ, d_�Ǽ�ʱ��, v_����Ա���, v_����Ա����,
               n_������, n_����id, n_��id, n_����id, 4);
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
          If r_Deposit.����id = 0 Then
            --��һ�γ�Ԥ��(���Ͻ���ID,���Ϊ0)
            Update ����Ԥ����¼ Set ��Ԥ�� = 0, ����id = n_����id, �������� = 4 Where ID = r_Deposit.ԭԤ��id;
          End If;
          --���ϴ�ʣ���
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����,
             ����Ա���, ��Ԥ��, ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������,
                   ��λ�ʺ�, d_�Ǽ�ʱ��, v_����Ա����, v_����Ա���, n_������, n_����id, n_��id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, n_����id, 4
            From ����Ԥ����¼
            Where NO = r_Deposit.No And ��¼״̬ = r_Deposit.��¼״̬ And ��¼���� In (1, 11) And Rownum = 1;
        
          --���²���Ԥ�����
          Update �������
          Set Ԥ����� = Nvl(Ԥ�����, 0) - n_������
          Where ����id = ����id_In And ���� = 1 And ���� = Nvl(1, 2)
          Returning Ԥ����� Into n_����ֵ;
          If Sql%RowCount = 0 Then
            Insert Into ������� (����id, Ԥ�����, ����, ����) Values (����id_In, -1 * n_������, 1, 1);
            n_����ֵ := -1 * n_������;
          End If;
          If Nvl(n_����ֵ, 0) = 0 Then
            Delete From �������
            Where ����id = ����id_In And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
          End If;
        
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
        If Instr(���㷽ʽ_In, ',') = 0 Then
          --ֻ����һ�ֽ��㷽ʽ��
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          Insert Into ����Ԥ����¼
            (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��,
             ������λ, ��������, �������)
          Values
            (n_Ԥ��id, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), Nvl(���㷽ʽ_In, v_�ֽ�), Nvl(n_������, 0), �Ǽ�ʱ��_In,
             v_����Ա���, v_����Ա����, n_����id, '�Һ��շ�', n_��id, �����id_In, Null, ֧������_In, ������ˮ��_In, ����˵��_In, ������λ_In, 4, v_�������);
        Else
          v_��������     := ���㷽ʽ_In || '|'; --�Կո�ֿ���|��β,û�н�������
          n_Exists       := 0;
          v_���㷽ʽ��¼ := '';
          While v_�������� Is Not Null Loop
            v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '|') - 1);
            v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
          
            v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
            n_���ʽ�� := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1));
          
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
                (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��,
                 ����˵��, ������λ, ��������, �������)
              Values
                (n_Ԥ��id, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), Nvl(v_���㷽ʽ, v_�ֽ�), Nvl(n_���ʽ��, 0), �Ǽ�ʱ��_In,
                 v_����Ա���, v_����Ա����, n_����id, '�Һ��շ�', n_��id, Null, Null, Null, Null, Null, ������λ_In, 4, v_�������);
            Else
              If n_Exists = 1 Then
                v_Err_Msg := 'Ŀǰ�ҺŽ�֧��һ���������㷽ʽ,���ܼ���������';
                Raise Err_Item;
              End If;
              Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
              Insert Into ����Ԥ����¼
                (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��,
                 ����˵��, ������λ, ��������, �������)
              Values
                (n_Ԥ��id, 4, 1, ���ݺ�_In, Decode(����id_In, 0, Null, ����id_In), Nvl(v_���㷽ʽ, v_�ֽ�), Nvl(n_���ʽ��, 0), �Ǽ�ʱ��_In,
                 v_����Ա���, v_����Ա����, n_����id, '�Һ��շ�', n_��id, �����id_In, Null, ֧������_In, ������ˮ��_In, ����˵��_In, ������λ_In, 4, v_�������);
              n_Exists := 1;
            End If;
          
            v_�������� := Substr(v_��������, Instr(v_��������, '|') + 1);
          End Loop;
        End If;
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
          Where �տ�Ա = v_����Ա���� And ���㷽ʽ = v_�ɿ�.���㷽ʽ And ���� = 1 And Nvl(���, 0) = 0;
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
        ����Ա��� = v_����Ա���, ԤԼ = Decode(������ʽ_In, 1, 0, 1),
        ������ = Decode(��������_In, 1, Null, Decode(������ʽ_In, 2, Null, v_����Ա����)),
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
        ԤԼ����Ա��� = Decode(������ʽ_In, 1, Nvl(ԤԼ����Ա���, Null), Nvl(ԤԼ����Ա���, v_����Ա���)), �����¼id = ��¼id_In
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
         ����Ա����, ����, ����, ԤԼ, ������, ����ʱ��, ������ˮ��, ����˵��, ������λ, ҽ�Ƹ��ʽ, ԤԼ����Ա, ԤԼ����Ա���, �����¼id)
      Values
        (n_�Һ�id, ���ݺ�_In, Decode(������ʽ_In, 2, 2, 1), Decode(��������_In, 1, 0, 1), r_Pati.����id, r_Pati.�����, r_Pati.����,
         r_Pati.�Ա�, r_Pati.����, ����_In, 0, v_����, Null, n_����id, v_ҽ������, 0, Null, d_�Ǽ�ʱ��, ����ʱ��_In,
         Case When(Nvl(������ʽ_In, 0)) = 1 Then Null Else ����ʱ��_In End, v_����Ա���, v_����Ա����, 0, n_����, Decode(������ʽ_In, 1, 0, 1),
         Decode(������ʽ_In, 2, Null, v_����Ա����), Decode(������ʽ_In, 2, To_Date(Null), d_�Ǽ�ʱ��), ������ˮ��_In, ����˵��_In, ������λ_In,
         v_���ʽ, Decode(������ʽ_In, 1, Null, v_����Ա����), Decode(������ʽ_In, 1, Null, v_����Ա���), ��¼id_In);
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
            n_��ʱ����ʾ := Nvl(Zl_To_Number(zl_GetSysParameter(270)), 0);
            If Nvl(������ʽ_In, 0) > 1 And n_��ʱ����ʾ = 1 And n_���÷�ʱ�� = 1 Then
              n_��ʱ����ʾ := 1;
            Else
              n_��ʱ����ʾ := Null;
            End If;
            --��������
            --.����ִ�в��š� �ķ�ʽ���ɶ���
            v_�������� := n_����id;
            v_�ŶӺ��� := Zlgetnextqueue(n_����id, n_�Һ�id, ����_In || '|' || ����_In);
            --�Һ�id_In,����_In,����_In,ȱʡ����_In,��չ_In(������)
            v_�Ŷ���� := Zlgetsequencenum(0, n_�Һ�id, 0);
            --�Һ�id_In,����_In,����_In,ȱʡ����_In,��չ_In(������)
            d_�Ŷ�ʱ�� := Zl_Get_Queuedate(n_�Һ�id, ����_In, ����_In, d_Date);
            --  ��������_In , ҵ������_In, ҵ��id_In,����id_In,�ŶӺ���_In,v_�Ŷӱ��,��������_In,����ID_IN, ����_In, ҽ������_In, �Ŷ�ʱ��_In
            Zl_�ŶӽкŶ���_Insert(v_��������, 0, n_�Һ�id, n_����id, v_�ŶӺ���, Null, r_Pati.����, r_Pati.����id, v_����, v_ҽ������, d_�Ŷ�ʱ��,
                             ԤԼ��ʽ_In, n_��ʱ����ʾ, v_�Ŷ����);
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
      Zl_���˹ҺŻ���_Update(v_ҽ������, n_ҽ��id, n_��Ŀid, n_����id, ����ʱ��_In, n_ԤԼ, ����_In, 0, ��¼id_In);
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
    When Err_Special Then
      Raise_Application_Error(-20105, '[ZLSOFT]+' || v_Err_Msg || '+[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End;

Begin
  n_�����¼id := �����¼id_In;
  v_Para       := zl_GetSysParameter(256);
  n_�Һ�ģʽ   := Substr(v_Para, 1, 1);
  Begin
    d_����ʱ�� := To_Date(Substr(v_Para, 3), 'yyyy-mm-dd hh24:mi:ss');
  Exception
    When Others Then
      d_����ʱ�� := Null;
  End;

  d_����ʱ�� := ����ʱ��_In;
  If d_����ʱ�� Is Null Then
    d_����ʱ�� := Sysdate;
  End If;

  If Sysdate - 10 > Nvl(d_����ʱ��, Sysdate - 30) Then
    If n_�Һ�ģʽ = 1 And Nvl(����ʱ��_In, Sysdate) > Nvl(d_����ʱ��, Sysdate - 30) And n_�����¼id Is Null Then
      v_Err_Msg := 'ϵͳ��ǰ���ڳ�����Ű�ģʽ������Ĳ����޷�ȷ���ҺŰ��ţ������ԣ�';
      Raise Err_Item;
    End If;
  Else
    If n_�Һ�ģʽ = 1 And Nvl(����ʱ��_In, Sysdate) > Nvl(d_����ʱ��, Sysdate - 30) And n_�����¼id Is Null Then
      Begin
        Select a.Id
        Into n_�����¼id
        From �ٴ������¼ A, �ٴ������Դ B
        Where a.��Դid = b.Id And b.���� = ����_In And Nvl(����ʱ��_In, Sysdate) Between a.��ʼʱ�� And a.��ֹʱ��;
      Exception
        When Others Then
          v_Err_Msg := 'ϵͳ��ǰ���ڳ�����Ű�ģʽ������Ĳ����޷�ȷ���ҺŰ��ţ������ԣ�';
          Raise Err_Item;
      End;
    End If;
  End If;

  If n_�����¼id Is Not Null Then
    --������Ű�ģʽ
    Zl_���������Һ�_����_Insert(n_�����¼id, ������ʽ_In, ����id_In, ����_In, ����_In, ���ݺ�_In, Ʊ�ݺ�_In, ���㷽ʽ_In, ժҪ_In, ����ʱ��_In, �Ǽ�ʱ��_In,
                        ������λ_In, �ҺŽ��ϼ�_In, ����id_In, �շ�Ʊ��_In, ������ˮ��_In, ����˵��_In, ԤԼ��ʽ_In, Ԥ��id_In, �����id_In, �������״̬_In,
                        �Ƿ������豸_In, ����id_In, ��������_In, ���ս���_In, ��Ԥ��_In, ֧������_In, �˺�����_In, �ѱ�_In, ������_In, ��������_In, ������_In);
  Else
    v_Temp := zl_GetSysParameter(256);
    If v_Temp Is Null Or Substr(v_Temp, 1, 1) = '0' Then
      Null;
    Else
      Begin
        d_����ʱ�� := To_Date(Substr(v_Temp, 3), 'YYYY-MM-DD hh24:mi:ss');
      Exception
        When Others Then
          Null;
      End;
      If ����ʱ��_In > d_����ʱ�� Then
        v_Err_Msg := '��ǰ�Һŵķ���ʱ��' || To_Char(����ʱ��_In, 'yyyy-mm-dd hh24:mi:ss') || '�Ѿ������˳�����Ű�ģʽ,������ʹ�üƻ��Ű�ģʽ�Һ�!';
        Raise Err_Item;
      End If;
    End If;
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
    v_Temp           := Nvl(zl_GetSysParameter('����ͬ���޹�N����', 1111), '0|0') || '|';
    n_ͬ���޺���     := To_Number(Substr(v_Temp, 1, Instr(v_Temp, '|') - 1));
    n_ͬ����Լ��     := To_Number(Nvl(zl_GetSysParameter('����ͬ����ԼN����', 1111), '0'));
    n_����ԤԼ������ := To_Number(Nvl(zl_GetSysParameter('����ԤԼ������', 1111), '0'));
    n_���˹Һſ����� := To_Number(Nvl(zl_GetSysParameter('���˹Һſ�������', 1111), '0'));
    n_ר�ҺŹҺ����� := To_Number(Nvl(zl_GetSysParameter('ר�ҺŹҺ�����'), '0'));
    n_ר�Һ�ԤԼ���� := To_Number(Nvl(zl_GetSysParameter('ר�Һ�ԤԼ����'), '0'));
  
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
    n_��������id := To_Number(Zl_����Ա(0, v_Temp));
    v_����Ա��� := Zl_����Ա(1, v_Temp);
    v_����Ա���� := Zl_����Ա(2, v_Temp);
    n_��id       := Zl_Get��id(v_����Ա����);
  
    --֧���������ύ���
    Select Nvl(Max(1), 0)
    Into n_Exists
    From ���˹Һż�¼
    Where ����id = ����id_In And �ű� = ����_In And ���� = ����_In And ����Ա���� = v_����Ա���� And Nvl(n_�����¼id, 0) = Nvl(�����¼id, 0) And
          �Ǽ�ʱ�� > Sysdate - 0.01 And ��¼״̬ = 1 And ����ʱ�� = ����ʱ��_In;
    If n_Exists = 1 Then
      v_Err_Msg := '�����Ѿ��Һ�,�����ظ�����ͬ�ĺţ�';
      Raise Err_Special;
    End If;
  
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
  
    Select Decode(To_Char(����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����',
                   '����')
    Into v_����
    From Dual;
    Begin
      If r_����.�ƻ�id Is Null Then
        Select Max(1) Into n_���÷�ʱ�� From �ҺŰ���ʱ�� Where ����id = r_����.Id And ���� = v_���� And Rownum < 2;
        Select Decode(To_Char(����ʱ��_In, 'D'), '1', ����, '2', ��һ, '3', �ܶ�, '4', ����, '5', ����, '6', ����, '7', ����, Null)
        Into v_ʱ���
        From �ҺŰ���
        Where ID = r_����.Id;
      Else
        Select Max(1)
        Into n_���÷�ʱ��
        From �Һżƻ�ʱ��
        Where �ƻ�id = r_����.�ƻ�id And ���� = v_���� And Rownum < 2;
        Select Decode(To_Char(����ʱ��_In, 'D'), '1', ����, '2', ��һ, '3', �ܶ�, '4', ����, '5', ����, '6', ����, '7', ����, Null)
        Into v_ʱ���
        From �ҺŰ��żƻ�
        Where ID = r_����.�ƻ�id;
      End If;
    Exception
      When Others Then
        n_���÷�ʱ�� := 0;
    End;
  
    If v_ʱ��� Is Not Null And d_����ʱ�� Is Not Null Then
      --����Ƿ��ģʽ�ҺŰ���
      Select To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss'),
             To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(��ֹʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss')
      Into d_��鿪ʼʱ��, d_������ʱ��
      From ʱ���
      Where ʱ��� = v_ʱ��� And վ�� Is Null And ���� Is Null;
      If d_��鿪ʼʱ�� > d_������ʱ�� Then
        d_������ʱ�� := d_������ʱ�� + 1;
      End If;
      If d_������ʱ�� > d_����ʱ�� Then
        --��ȡ�����¼id
        Begin
          Select a.Id
          Into n_�����¼id
          From �ٴ������¼ A, �ٴ������Դ B
          Where a.��Դid = b.Id And b.���� = ����_In And �ϰ�ʱ�� = v_ʱ��� And ����ʱ��_In Between ��ʼʱ�� And ��ֹʱ��;
        Exception
          When Others Then
            n_�����¼id := Null;
        End;
      End If;
    End If;
  
    --�Բ������ƽ��м��
    --����ԤԼ���ۿ�ʱ���м��
    If ������ʽ_In = 2 Then
      If Nvl(n_ͬ����Լ��, 0) <> 0 Or Nvl(n_����ԤԼ������, 0) <> 0 Then
        n_��Լ���� := 0;
        For c_Chkitem In (Select Distinct ִ�в���id
                          From ���˹Һż�¼
                          Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 2 And ԤԼʱ�� Between Trunc(����ʱ��_In) And
                                Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ִ�в���id <> r_����.����id) Loop
          n_��Լ���� := n_��Լ���� + 1;
        End Loop;
        If n_��Լ���� >= Nvl(n_����ԤԼ������, 0) And Nvl(n_����ԤԼ������, 0) > 0 Then
          v_Err_Msg := 'ͬһ�������ͬʱ��ԤԼ[' || Nvl(n_����ԤԼ������, 0) || ']������,������ԤԼ��';
          Raise Err_Item;
        End If;
      
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 2 And ԤԼʱ�� Between Trunc(����ʱ��_In) And
              Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ִ�в���id = r_����.����id;
        If n_Count >= Nvl(n_ͬ����Լ��, 0) And Nvl(n_ͬ����Լ��, 0) > 0 Then
          v_Err_Msg := '�ò����Ѿ��ڸÿ���ԤԼ��' || n_Count || '��,������ԤԼ��';
          Raise Err_Item;
        End If;
      End If;
      If Nvl(n_ר�Һ�ԤԼ����, 0) <> 0 Then
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 2 And ԤԼʱ�� Between Trunc(����ʱ��_In) And
              Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And �ű� = r_����.����;
        If n_Count >= Nvl(n_ר�Һ�ԤԼ����, 0) And Nvl(n_ר�Һ�ԤԼ����, 0) > 0 Then
          v_Err_Msg := '�ò����Ѿ���������ԤԼ����,������ԤԼ��';
          Raise Err_Item;
        End If;
      End If;
    Else
      If Nvl(n_ͬ���޺���, 0) <> 0 Or Nvl(n_���˹Һſ�����, 0) <> 0 Then
        n_��Լ���� := 0;
        For c_Chkitem In (Select Distinct ִ�в���id
                          From ���˹Һż�¼
                          Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 1 And ����ʱ�� Between Trunc(����ʱ��_In) And
                                Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ִ�в���id <> r_����.����id) Loop
          n_��Լ���� := n_��Լ���� + 1;
        End Loop;
        If n_��Լ���� >= Nvl(n_���˹Һſ�����, 0) And Nvl(n_���˹Һſ�����, 0) > 0 Then
          v_Err_Msg := 'ͬһ�������ͬʱ�ܹҺ�[' || Nvl(n_���˹Һſ�����, 0) || ']������,�����ٹҺţ�';
          Raise Err_Item;
        End If;
      
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 1 And ����ʱ�� Between Trunc(����ʱ��_In) And
              Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ִ�в���id = r_����.����id;
        If n_Count >= Nvl(n_ͬ���޺���, 0) And Nvl(n_ͬ���޺���, 0) > 0 Then
          v_Err_Msg := '�ò����Ѿ��ڸÿ��ҹҺ���' || n_Count || '��,�����ٹҺţ�';
          Raise Err_Item;
        End If;
      End If;
      If Nvl(n_ר�ҺŹҺ�����, 0) <> 0 Then
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 1 And ԤԼʱ�� Between Trunc(����ʱ��_In) And
              Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And �ű� = r_����.����;
        If n_Count >= Nvl(n_ר�ҺŹҺ�����, 0) And Nvl(n_ר�ҺŹҺ�����, 0) > 0 Then
          v_Err_Msg := '�ò����Ѿ��������ŹҺ�����,�����ٹҺţ�';
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
            If r_����.�ƻ�id Is Null Then
              Select Count(*), Max(��ʼʱ��)
              Into n_Count, d_ʱ�ο�ʼʱ��
              From �ҺŰ���ʱ��
              Where ����id = r_����.Id And ���� = v_���� And ��� = Nvl(����_In, 0);
            Else
              Select Count(*), Max(��ʼʱ��)
              Into n_Count, d_ʱ�ο�ʼʱ��
              From �Һżƻ�ʱ��
              Where �ƻ�id = r_����.�ƻ�id And ���� = v_���� And ��� = Nvl(����_In, 0);
            End If;
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
            If r_����.�ƻ�id Is Null Then
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
            Else
              For v_ʱ�� In (Select To_Date(v_Temp || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                                  To_Date(To_Char(Sysdate + Decode(Sign(��ʼʱ�� - ����ʱ��), 1, 1, 0), 'yyyy-mm-dd') || ' ' ||
                                           To_Char(����ʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, ��������, �Ƿ�ԤԼ
                           From �Һżƻ�ʱ��
                           Where �ƻ�id = r_����.�ƻ�id And ���� = v_���� And ��� = Nvl(����_In, 0)) Loop
                If Sysdate > v_ʱ��.����ʱ�� Then
                  v_Err_Msg := '�ű�Ϊ' || ����_In || '������Ϊ' || Nvl(����_In, 0) || '�İ���,�Ѿ�����ʱ��,������' || v_Temp || '��';
                  Raise Err_Item;
                End If;
              End Loop;
            End If;
          End If;
        Elsif ������ʽ_In > 1 Then
          --δ������ŵ�,��Ҫ���ԤԼ�����
          n_Count := 0;
          If r_����.�ƻ�id Is Null Then
            For v_ʱ�� In (Select ���, ��ʼʱ��, ����ʱ��, ��������, �Ƿ�ԤԼ
                         From �ҺŰ���ʱ��
                         Where ����id = r_����.Id And ���� = v_���� And
                               (('3000-01-10 ' || To_Char(����ʱ��_In, 'HH24:MI:SS') Between
                               Decode(Sign(��ʼʱ�� - ����ʱ�� - 1 / 24 / 60 / 60), 1,
                                        '3000-01-09 ' || To_Char(��ʼʱ��, 'HH24:MI:SS'),
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
          Else
            For v_ʱ�� In (Select ���, ��ʼʱ��, ����ʱ��, ��������, �Ƿ�ԤԼ
                         From �Һżƻ�ʱ��
                         Where �ƻ�id = r_����.�ƻ�id And ���� = v_���� And
                               (('3000-01-10 ' || To_Char(����ʱ��_In, 'HH24:MI:SS') Between
                               Decode(Sign(��ʼʱ�� - ����ʱ�� - 1 / 24 / 60 / 60), 1,
                                        '3000-01-09 ' || To_Char(��ʼʱ��, 'HH24:MI:SS'),
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
          End If;
        
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
                ������Ŀ = Decode(To_Char(����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',
                              '7', '����', Null) And ���� <> 0 And ��� = 0 And Rownum < 2;
        Else
          Select 0
          Into n_���
          From ������λ���ſ���
          Where ������λ = ������λ_In And ����id = n_Tmp����id And
                ������Ŀ = Decode(To_Char(����ʱ��_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',
                              '7', '����', Null) And ���� <> 0 And ��� = 0 And Rownum < 2;
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
  
    If Nvl(������_In, 0) = 1 Then
      Begin
        Select �շ�ϸĿid Into n_������id From �շ��ض���Ŀ Where �ض���Ŀ = '������';
        v_�շ���Ŀids := r_����.��Ŀid || ',' || n_������id;
      Exception
        When Others Then
          v_Err_Msg := '����ȷ��������,�Һ�ʧ��!';
          Raise Err_Item;
      End;
    Else
      v_�շ���Ŀids := r_����.��Ŀid;
    End If;
  
    For c_Item In (Select 1 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, 1 As ����,
                          c.Id As ������Ŀid, c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, Null As ��������
                   From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C
                   Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = r_����.��Ŀid And d_����ʱ�� Between b.ִ������ And
                         Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))
                   Union All
                   Select 2 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, 1 As ����,
                          c.Id As ������Ŀid, c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, Null As ��������
                   From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C
                   Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = n_������id And d_����ʱ�� Between b.ִ������ And
                         Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))
                   Union All
                   Select 3 As ����, a.���, a.Id As ��Ŀid, a.���� As ��Ŀ����, a.���� As ��Ŀ����, a.���㵥λ, a.���ηѱ�, d.�������� As ����,
                          c.Id As ������Ŀid, c.���� As ������Ŀ, c.���� As �������, c.�վݷ�Ŀ, b.�ּ� As ����, 1 As ��������
                   From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C, �շѴ�����Ŀ D
                   Where b.�շ�ϸĿid = a.Id And b.������Ŀid = c.Id And a.Id = d.����id And
                         d.����id In (Select Column_Value From Table(f_Str2list(v_�շ���Ŀids))) And d_����ʱ�� Between b.ִ������ And
                         Nvl(b.��ֹ����, To_Date('3000-1-1', 'YYYY-MM-DD'))
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
           Decode(������ʽ_In, 2, To_Char(����_In), v_����), r_Pati.����id, r_Pati.�����, r_Pati.���ʽ, r_Pati.����, r_Pati.�Ա�,
           r_Pati.����, r_Pati.�ѱ�, r_����.����id, c_Item.���, ����_In, c_Item.��Ŀid, c_Item.������Ŀid, c_Item.�վݷ�Ŀ, 1, c_Item.����,
           c_Item.����, n_Ӧ�ս��, n_ʵ�ս��, Decode(������ʽ_In, 2, Null, n_ʵ�ս��), n_����id, 0, n_��������id, v_����Ա����,
           Decode(������ʽ_In, 2, v_����Ա����, Null), r_����.����id, r_����.ҽ������, v_����Ա���, v_����Ա����, ����ʱ��_In, d_�Ǽ�ʱ��, Null, 0, Null,
           Null, ժҪ_In, ԤԼ��ʽ_In, Decode(������ʽ_In, 2, Null, n_��id));
      End If;
      n_�к� := n_�к� + 1;
    
    End Loop;
  
    If Round(Nvl(�ҺŽ��ϼ�_In, 0), 5) <> Round(Nvl(n_ʵ�ս��ϼ�, 0), 5) Then
      v_Err_Msg := '���ιҺŽ���ȷ,��������ΪҽԺ�����˼۸�,�����»�ȡ�Һ��շ���Ŀ�ļ۸�,���ܼ�����';
      Raise Err_Item;
    End If;
  
    If n_���÷�ʱ�� = 1 Then
      d_Date := To_Date(To_Char(����ʱ��_In, 'yyyy-mm-dd') || ' ' || To_Char(d_ʱ�ο�ʼʱ��, 'hh24:mi:ss'),
                        'yyyy-mm-dd hh24:mi:ss');
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
  
    If n_�����¼id Is Not Null Then
      Update �ٴ�������ſ���
      Set �Һ�״̬ = Decode(������ʽ_In, 2, 2, 1), ����Ա���� = v_����Ա����
      Where ��¼id = n_�����¼id And ��� = n_���;
      If ������ʽ_In = 2 Then
        Update �ٴ������¼ Set ��Լ�� = ��Լ�� + 1 Where ID = n_�����¼id;
      Else
        If ������ʽ_In <> 1 Then
          Update �ٴ������¼
          Set ��Լ�� = ��Լ�� + 1, �ѹ��� = �ѹ��� + 1, �����ѽ��� = �����ѽ��� + 1
          Where ID = n_�����¼id;
        Else
          Update �ٴ������¼ Set �ѹ��� = �ѹ��� + 1 Where ID = n_�����¼id;
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
              (n_Ԥ��id, 4, ���ݺ�_In, 1, Decode(����id_In, 0, Null, ����id_In), '���ս���', v_���㷽ʽ, d_�Ǽ�ʱ��, v_����Ա���, v_����Ա����,
               n_������, n_����id, n_��id, n_����id, 4);
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
          If r_Deposit.����id = 0 Then
            --��һ�γ�Ԥ��(���Ͻ���ID,���Ϊ0)
            Update ����Ԥ����¼ Set ��Ԥ�� = 0, ����id = n_����id, �������� = 4 Where ID = r_Deposit.ԭԤ��id;
          End If;
          --���ϴ�ʣ���
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����,
             ����Ա���, ��Ԥ��, ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, ��������)
            Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������,
                   ��λ�ʺ�, d_�Ǽ�ʱ��, v_����Ա����, v_����Ա���, n_������, n_����id, n_��id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, n_����id, 4
            From ����Ԥ����¼
            Where NO = r_Deposit.No And ��¼״̬ = r_Deposit.��¼״̬ And ��¼���� In (1, 11) And Rownum = 1;
        
          --���²���Ԥ�����
          Update �������
          Set Ԥ����� = Nvl(Ԥ�����, 0) - n_������
          Where ����id = ����id_In And ���� = 1 And ���� = Nvl(1, 2)
          Returning Ԥ����� Into n_����ֵ;
          If Sql%RowCount = 0 Then
            Insert Into ������� (����id, Ԥ�����, ����, ����) Values (����id_In, -1 * n_������, 1, 1);
            n_����ֵ := -1 * n_������;
          End If;
          If Nvl(n_����ֵ, 0) = 0 Then
            Delete From �������
            Where ����id = ����id_In And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
          End If;
        
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
          (n_Ԥ��id, 4, 1, ���ݺ�_In, r_Pati.����id, ���㷽ʽ_In, Nvl(n_������, 0), d_�Ǽ�ʱ��, v_����Ա���, v_����Ա����, n_����id,
           ������λ_In || '�ɿ�', n_��id, ������ˮ��_In, ����˵��_In, n_����id, ������λ_In, �����id_In, ֧������_In, 4);
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
        ����Ա��� = v_����Ա���, ԤԼ = Decode(������ʽ_In, 1, 0, 1),
        ������ = Decode(��������_In, 1, Null, Decode(������ʽ_In, 2, Null, v_����Ա����)),
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
         Decode(������ʽ_In, 2, Null, v_����Ա����), Decode(������ʽ_In, 2, To_Date(Null), d_�Ǽ�ʱ��), ������ˮ��_In, ����˵��_In, ������λ_In,
         v_���ʽ, Decode(������ʽ_In, 1, Null, v_����Ա����), Decode(������ʽ_In, 1, Null, v_����Ա���));
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
            n_��ʱ����ʾ := Nvl(Zl_To_Number(zl_GetSysParameter(270)), 0);
            If Nvl(������ʽ_In, 0) > 1 And n_��ʱ����ʾ = 1 And n_���÷�ʱ�� = 1 Then
              n_��ʱ����ʾ := 1;
            Else
              n_��ʱ����ʾ := Null;
            End If;
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
                             d_�Ŷ�ʱ��, ԤԼ��ʽ_In, n_��ʱ����ʾ, v_�Ŷ����);
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
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]+' || v_Err_Msg || '+[ZLSOFT]');
  When Err_Special Then
    Raise_Application_Error(-20105, '[ZLSOFT]+' || v_Err_Msg || '+[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���������Һ�_Insert;
/

--117999:���ϴ�,2017-12-13,�������״̬���
CREATE OR REPLACE Function Zl_Fun_Get�ٴ�����ԤԼ״̬
(
  ��¼id_In   In �ٴ������¼.Id%Type,
  ԤԼʱ��_In In ���˹Һż�¼.ԤԼʱ��%Type,
  ���_In     �ٴ�������ſ���.���%Type := Null,
  ԤԼ��ʽ_In ԤԼ��ʽ.����%Type := Null,
  ������λ_In �Һź�����λ.����%Type := Null,
  �շ�ԤԼ_In Number := 0,
  ����Ա����_In �Һ����״̬.����Ա����%Type := Null,
  ������_In   �Һ����״̬.������%Type := Null
) Return Varchar2 As
  --���ܣ��жϳ����¼��ԤԼʱ���Ƿ��ԤԼ
  --��Σ�
  --���أ�
  --     ��ʽ��ԤԼ״̬|��ʾ��Ϣ���磺"1|ԤԼʱ�䲻�ڵ�ǰ�ϰ�ʱ��ʱ�䷶Χ�ڡ�"
  --     ԤԼ״̬��
  --         0-��ԤԼ
  --         ======================================================
  --         1-����ԤԼ��ԤԼʱ�䲻�ڵ�ǰ�ϰ�ʱ��ʱ�䷶Χ��
  --         2-����ԤԼ����ǰ�ϰ�ʱ�ν�ֹԤԼ
  --         3-����ԤԼ����ǰ�ϰ�ʱ����ԤԼʱ��ʱ��ͣ��
  --         4-����ԤԼ����ǰ�ϰ�ʱ��ʣ���ԤԼ��Ϊ��
  --         ======================================================
  --         5-����ԤԼ����ǰԤԼʱ���ڷ����ڼ���ʱ�䷶Χ�ڣ����ϰ�
  --         6-����ԤԼ����ǰԤԼʱ���ڷ����ڼ���ʱ�䷶Χ�ڣ���ֹԤԼ
  --         7-����ԤԼ����ǰԤԼʱ���ڷ����ڼ��ղ�����ԤԼ��ʱ�䷶Χ��
  --         8-����ԤԼ����ǰԤԼʱ���ڷ����ڼ��ղ�����Һŵ�ʱ�䷶Χ��
  --         9-����ԤԼ����ǰԤԼʱ���ڷ����ڼ���ʱ�䷶Χ�ڣ���ͣ��
  --         ======================================================
  --         10-����ԤԼ����ǰԤԼ��ʽ��ֹԤԼ
  --         11-����ԤԼ����ǰԤԼ��ʽ��ԤԼ������
  --         ======================================================
  --         12-����ԤԼ����ǰ������λ��ֹԤԼ
  --         13-����ԤԼ����ǰ������λ��ԤԼ������
  --         ======================================================
  --         14-����ԤԼ����ǰ��Ž�ֹԤԼ
  --         15-����ԤԼ����ǰ����Ѿ���ʹ��
  --         16-����ԤԼ����ǰ��Ų�����
  --
  n_��Դid         �ٴ������¼.��Դid%Type;
  n_�Ƿ��ʱ��     �ٴ������¼.�Ƿ��ʱ��%Type;
  n_ԤԼ����       �ٴ������¼.ԤԼ����%Type;
  d_ͣ�￪ʼʱ��   �ٴ������¼.ͣ�￪ʼʱ��%Type;
  d_ͣ����ֹʱ��   �ٴ������¼.ͣ����ֹʱ��%Type;
  v_ͣ��ԭ��       �ٴ������¼.ͣ��ԭ��%Type;
  n_��Լ��         �ٴ������¼.��Լ��%Type;
  n_��Լ��         �ٴ������¼.��Լ��%Type;
  n_��ռ           �ٴ������¼.�Ƿ��ռ%Type;
  n_���Ʒ�ʽ       �ٴ�����Һſ��Ƽ�¼.���Ʒ�ʽ%Type;
  n_����           �ٴ�����Һſ��Ƽ�¼.����%Type;
  n_��������       �ٴ�����Һſ��Ƽ�¼.����%Type;
  n_��ſ���       �ٴ������¼.�Ƿ���ſ���%Type;
  v_ԤԼ��ʽ       �ٴ�����Һſ��Ƽ�¼.����%Type;
  n_����           �ٴ�����Һſ��Ƽ�¼.����%Type;
  n_ԤԼ��ʽ��Լ�� �ٴ������¼.��Լ��%Type;
  n_ԤԼ��ʽ��Լ�� �ٴ������¼.��Լ��%Type;
  n_�Һ�״̬       �ٴ�������ſ���.�Һ�״̬%Type;
  n_�Ƿ�ԤԼ       �ٴ�������ſ���.�Ƿ�ԤԼ%Type;

  n_���տ���״̬ �ٴ������Դ.���տ���״̬%Type;

  v_����ԤԼ �������ձ�.����ԤԼ����%Type;
  v_����Һ� �������ձ�.����Һ�����%Type;
  n_Count    Number(2);
  n_��ʹ��   Number(5);
  v_���Ż����� �Һ����״̬.������%Type;
  v_���Ų���Ա �Һ����״̬.����Ա����%Type;
Begin
  Begin
    Select a.��Դid, a.�Ƿ��ʱ��, a.ԤԼ����, a.ͣ�￪ʼʱ��, a.ͣ����ֹʱ��, a.ͣ��ԭ��, Nvl(��Լ��, �޺���), ��Լ��, �Ƿ��ռ, �Ƿ���ſ���
    Into n_��Դid, n_�Ƿ��ʱ��, n_ԤԼ����, d_ͣ�￪ʼʱ��, d_ͣ����ֹʱ��, v_ͣ��ԭ��, n_��Լ��, n_��Լ��, n_��ռ, n_��ſ���
    From �ٴ������¼ A
    Where a.Id = ��¼id_In And ԤԼʱ��_In Between ��ʼʱ�� And ��ֹʱ��;
  Exception
    When Others Then
      Return '1|ԤԼʱ�䲻�ڵ�ǰ�ϰ�ʱ��ʱ�䷶Χ�ڡ�';
  End;

  --ԤԼ��ʽ���
  If ԤԼ��ʽ_In Is Not Null Then
    Begin
      Select ���Ʒ�ʽ
      Into n_���Ʒ�ʽ
      From �ٴ�����Һſ��Ƽ�¼
      Where ���� = 2 And ���� = 1 And ��¼id = ��¼id_In And ���� = ԤԼ��ʽ_In And Rownum < 2;
    Exception
      When Others Then
        Begin
          Select ���Ʒ�ʽ
          Into n_���Ʒ�ʽ
          From �ٴ�����Һſ��Ƽ�¼
          Where ���� = 2 And ���� = 1 And ��¼id = ��¼id_In And Rownum < 2;
        Exception
          When Others Then
            Null;
        End;
    End;
    If n_���Ʒ�ʽ = 0 Then
      Return '10|��ǰԤԼ��ʽ��ֹԤԼ��';
    End If;
    If n_���Ʒ�ʽ = 1 Or n_���Ʒ�ʽ = 2 Then
      Select Nvl(��Լ��, �޺���) Into n_ԤԼ��ʽ��Լ�� From �ٴ������¼ Where ID = ��¼id_In;
      If n_��ռ = 0 Then
        Begin
          Select ����
          Into n_����
          From �ٴ�����Һſ��Ƽ�¼
          Where ���� = 2 And ���� = 1 And ���� = ԤԼ��ʽ_In And ��¼id = ��¼id_In;
        Exception
          When Others Then
            n_���� := Null;
        End;
        If n_���� Is Not Null Then
          If n_���Ʒ�ʽ = 1 Then
            n_ԤԼ��ʽ��Լ�� := Round(n_ԤԼ��ʽ��Լ�� * n_���� / 100);
          Else
            n_ԤԼ��ʽ��Լ�� := n_����;
          End If;
          Select Count(1)
          Into n_ԤԼ��ʽ��Լ��
          From ���˹Һż�¼
          Where �����¼id = ��¼id_In And ��¼״̬ = 1 And ԤԼ��ʽ = ԤԼ��ʽ_In;
          If n_ԤԼ��ʽ��Լ�� >= n_ԤԼ��ʽ��Լ�� Then
            Return '11|��ǰԤԼ��ʽ��ԤԼ�����㡣';
          End If;
        End If;
      Else
        --��������ռ
        Begin
          Select ����
          Into n_����
          From �ٴ�����Һſ��Ƽ�¼
          Where ���� = 2 And ���� = 1 And ���� = ԤԼ��ʽ_In And ��¼id = ��¼id_In;
        Exception
          When Others Then
            n_���� := Null;
        End;
        If n_���� Is Not Null Then
          If n_���Ʒ�ʽ = 1 Then
            n_ԤԼ��ʽ��Լ�� := Round(n_ԤԼ��ʽ��Լ�� * n_���� / 100);
          Else
            n_ԤԼ��ʽ��Լ�� := n_����;
          End If;
          Select Count(1)
          Into n_ԤԼ��ʽ��Լ��
          From ���˹Һż�¼
          Where �����¼id = ��¼id_In And ��¼״̬ = 1 And ԤԼ��ʽ = ԤԼ��ʽ_In;
          If n_ԤԼ��ʽ��Լ�� >= n_ԤԼ��ʽ��Լ�� Then
            Return '11|��ǰԤԼ��ʽ��ԤԼ�����㡣';
          End If;
        Else
          If �շ�ԤԼ_In = 0 Then
            For r_���� In (Select ����, ����, ���� From �ٴ�����Һſ��Ƽ�¼ Where ���� = 1 And ��¼id = ��¼id_In) Loop
              If r_����.���� = 1 Then
                Select Count(1)
                Into n_��ʹ��
                From ���˹Һż�¼
                Where �����¼id = ��¼id_In And ������λ = r_����.���� And ��¼״̬ = 1;
              Else
                Select Count(1)
                Into n_��ʹ��
                From ���˹Һż�¼
                Where �����¼id = ��¼id_In And ԤԼ��ʽ = r_����.���� And ��¼״̬ = 1;
              End If;
              If n_���Ʒ�ʽ = 1 Then
                n_�������� := Nvl(n_��������, 0) + Round(r_����.���� * n_ԤԼ��ʽ��Լ�� / 100) - Nvl(n_��ʹ��, 0);
              Else
                n_�������� := Nvl(n_��������, 0) + r_����.���� - Nvl(n_��ʹ��, 0);
              End If;
            End Loop;
            Select Count(1) Into n_��ʹ�� From ���˹Һż�¼ Where �����¼id = ��¼id_In And ��¼״̬ = 1;
            If n_ԤԼ��ʽ��Լ�� - n_�������� - n_��ʹ�� > 0 Then
              Null;
            Else
              Return '11|��ǰԤԼ��ʽ��ԤԼ�����㡣';
            End If;
          Else
            For r_���� In (Select ����, ����, ����
                         From �ٴ�����Һſ��Ƽ�¼
                         Where ���� = 1 And ���� = 2 And ��¼id = ��¼id_In) Loop
              Select Count(1)
              Into n_��ʹ��
              From ���˹Һż�¼
              Where �����¼id = ��¼id_In And ԤԼ��ʽ = r_����.���� And ��¼״̬ = 1;
              If n_���Ʒ�ʽ = 1 Then
                n_�������� := Nvl(n_��������, 0) + Round(r_����.���� * n_ԤԼ��ʽ��Լ�� / 100) - Nvl(n_��ʹ��, 0);
              Else
                n_�������� := Nvl(n_��������, 0) + r_����.���� - Nvl(n_��ʹ��, 0);
              End If;
            End Loop;
            Select Count(1) Into n_��ʹ�� From ���˹Һż�¼ Where �����¼id = ��¼id_In And ��¼״̬ = 1;
            If n_ԤԼ��ʽ��Լ�� - n_�������� - n_��ʹ�� > 0 Then
              Null;
            Else
              Return '11|��ǰԤԼ��ʽ��ԤԼ�����㡣';
            End If;
          End If;
        End If;
      End If;
    End If;
    If n_���Ʒ�ʽ = 3 Then
      If n_��ſ��� = 1 Then
        If �շ�ԤԼ_In = 0 Then
          Begin
            Select ����, ����, ����
            Into n_ԤԼ��ʽ��Լ��, v_ԤԼ��ʽ, n_����
            From �ٴ�����Һſ��Ƽ�¼
            Where ���� = 1 And ��¼id = ��¼id_In And ��� = ���_In;
          Exception
            When Others Then
              n_ԤԼ��ʽ��Լ�� := Null;
          End;
          If n_ԤԼ��ʽ��Լ�� Is Not Null Then
            If v_ԤԼ��ʽ <> ԤԼ��ʽ_In Or n_���� = 1 Then
              Return '11|��ǰԤԼ��ʽ��ԤԼ�����㡣';
            End If;
            Select Nvl(Max(1), 0)
            Into n_ԤԼ��ʽ��Լ��
            From ���˹Һż�¼
            Where �����¼id = ��¼id_In And ���� = ���_In;
            If n_ԤԼ��ʽ��Լ�� >= n_ԤԼ��ʽ��Լ�� Then
              Return '11|��ǰԤԼ��ʽ��ԤԼ�����㡣';
            End If;
          End If;
        Else
          Begin
            Select ����, ����, ����
            Into n_ԤԼ��ʽ��Լ��, v_ԤԼ��ʽ, n_����
            From �ٴ�����Һſ��Ƽ�¼
            Where ���� = 1 And ���� = 2 And ��¼id = ��¼id_In And ��� = ���_In;
          Exception
            When Others Then
              n_ԤԼ��ʽ��Լ�� := Null;
          End;
          If n_ԤԼ��ʽ��Լ�� Is Not Null Then
            If v_ԤԼ��ʽ <> ԤԼ��ʽ_In Then
              Return '11|��ǰԤԼ��ʽ��ԤԼ�����㡣';
            End If;
            Select Nvl(Max(1), 0)
            Into n_ԤԼ��ʽ��Լ��
            From ���˹Һż�¼
            Where �����¼id = ��¼id_In And ���� = ���_In;
            If n_ԤԼ��ʽ��Լ�� >= n_ԤԼ��ʽ��Լ�� Then
              Return '11|��ǰԤԼ��ʽ��ԤԼ�����㡣';
            End If;
          End If;
        End If;
      Else
        If �շ�ԤԼ_In = 0 Then
          For r_���� In (Select ����, ����, ����
                       From �ٴ�����Һſ��Ƽ�¼
                       Where ���� = 1 And ��¼id = ��¼id_In And ��� = ���_In) Loop
            If r_����.���� <> ԤԼ��ʽ_In Or r_����.���� = 1 Then
              If r_����.���� = 1 Then
                Select Count(1)
                Into n_��ʹ��
                From �ٴ�������ſ��� A, ���˹Һż�¼ B
                Where a.��¼id = ��¼id_In And a.ԤԼ˳��� Is Not Null And Nvl(a.�Һ�״̬, 0) <> 0 And a.��ע = b.���� And
                      b.������λ = r_����.���� And b.��¼״̬ = 1;
              Else
                Select Count(1)
                Into n_��ʹ��
                From �ٴ�������ſ��� A, ���˹Һż�¼ B
                Where a.��¼id = ��¼id_In And a.ԤԼ˳��� Is Not Null And Nvl(a.�Һ�״̬, 0) <> 0 And a.��ע = b.���� And
                      b.ԤԼ��ʽ = r_����.���� And b.��¼״̬ = 1;
              End If;
              n_�������� := Nvl(n_��������, 0) + r_����.���� - Nvl(n_��ʹ��, 0);
            Else
              Select Count(1)
              Into n_ԤԼ��ʽ��Լ��
              From �ٴ�������ſ��� A, ���˹Һż�¼ B
              Where a.��¼id = ��¼id_In And a.ԤԼ˳��� Is Not Null And Nvl(a.�Һ�״̬, 0) <> 0 And a.��ע = b.���� And
                    b.ԤԼ��ʽ = ԤԼ��ʽ_In And b.��¼״̬ = 1;
              If n_ԤԼ��ʽ��Լ�� >= n_ԤԼ��ʽ��Լ�� Then
                Return '11|��ǰԤԼ��ʽ��ԤԼ�����㡣';
              End If;
            End If;
          End Loop;
          Select Count(1)
          Into n_��ʹ��
          From �ٴ�������ſ��� A
          Where a.��¼id = ��¼id_In And a.ԤԼ˳��� Is Not Null And Nvl(a.�Һ�״̬, 0) <> 0 And ��� = ���_In;
          Select Nvl(��Լ��, �޺���) Into n_ԤԼ��ʽ��Լ�� From �ٴ������¼ Where ID = ��¼id_In;
          If n_ԤԼ��ʽ��Լ�� - n_�������� - n_��ʹ�� > 0 Then
            Null;
          Else
            Return '11|��ǰԤԼ��ʽ��ԤԼ�����㡣';
          End If;
        Else
          For r_���� In (Select ����, ����, ����
                       From �ٴ�����Һſ��Ƽ�¼
                       Where ���� = 1 And ���� = 2 And ��¼id = ��¼id_In And ��� = ���_In) Loop
            If r_����.���� <> ԤԼ��ʽ_In Then
              Select Count(1)
              Into n_��ʹ��
              From �ٴ�������ſ��� A, ���˹Һż�¼ B
              Where a.��¼id = ��¼id_In And a.ԤԼ˳��� Is Not Null And Nvl(a.�Һ�״̬, 0) <> 0 And a.��ע = b.���� And
                    b.ԤԼ��ʽ = r_����.���� And b.��¼״̬ = 1;
              n_�������� := Nvl(n_��������, 0) + r_����.���� - Nvl(n_��ʹ��, 0);
            Else
              Select Count(1)
              Into n_ԤԼ��ʽ��Լ��
              From �ٴ�������ſ��� A, ���˹Һż�¼ B
              Where a.��¼id = ��¼id_In And a.ԤԼ˳��� Is Not Null And Nvl(a.�Һ�״̬, 0) <> 0 And a.��ע = b.���� And
                    b.ԤԼ��ʽ = ԤԼ��ʽ_In And b.��¼״̬ = 1;
              If n_ԤԼ��ʽ��Լ�� >= n_ԤԼ��ʽ��Լ�� Then
                Return '11|��ǰԤԼ��ʽ��ԤԼ�����㡣';
              End If;
            End If;
          End Loop;
          Select Count(1)
          Into n_��ʹ��
          From �ٴ�������ſ��� A
          Where a.��¼id = ��¼id_In And a.ԤԼ˳��� Is Not Null And Nvl(a.�Һ�״̬, 0) <> 0 And ��� = ���_In;
          Select Nvl(��Լ��, �޺���) Into n_ԤԼ��ʽ��Լ�� From �ٴ������¼ Where ID = ��¼id_In;
          If n_ԤԼ��ʽ��Լ�� - n_�������� - n_��ʹ�� > 0 Then
            Null;
          Else
            Return '11|��ǰԤԼ��ʽ��ԤԼ�����㡣';
          End If;
        End If;
      End If;
    End If;
  End If;

  --������λ���
  If ������λ_In Is Not Null Then
    Begin
      Select ���Ʒ�ʽ
      Into n_���Ʒ�ʽ
      From �ٴ�����Һſ��Ƽ�¼
      Where ���� = 1 And ���� = 1 And ��¼id = ��¼id_In And ���� = ������λ_In And Rownum < 2;
    Exception
      When Others Then
        Begin
          Select ���Ʒ�ʽ
          Into n_���Ʒ�ʽ
          From �ٴ�����Һſ��Ƽ�¼
          Where ���� = 1 And ���� = 1 And ��¼id = ��¼id_In And Rownum < 2;
        Exception
          When Others Then
            Null;
        End;
    End;
    If n_���Ʒ�ʽ = 0 Then
      Return '12|��ǰ������λ��ֹԤԼ��';
    End If;
    If n_���Ʒ�ʽ = 1 Or n_���Ʒ�ʽ = 2 Then
      Select Nvl(��Լ��, �޺���) Into n_ԤԼ��ʽ��Լ�� From �ٴ������¼ Where ID = ��¼id_In;
      If n_��ռ = 0 Then
        Begin
          Select ����
          Into n_����
          From �ٴ�����Һſ��Ƽ�¼
          Where ���� = 1 And ���� = 1 And ���� = ������λ_In And ��¼id = ��¼id_In;
        Exception
          When Others Then
            n_���� := Null;
        End;
        If n_���� Is Not Null Then
          If n_���Ʒ�ʽ = 1 Then
            n_ԤԼ��ʽ��Լ�� := Round(n_ԤԼ��ʽ��Լ�� * n_���� / 100);
          Else
            n_ԤԼ��ʽ��Լ�� := n_����;
          End If;
          Select Count(1)
          Into n_ԤԼ��ʽ��Լ��
          From ���˹Һż�¼
          Where �����¼id = ��¼id_In And ��¼״̬ = 1 And ������λ = ������λ_In;
          If n_ԤԼ��ʽ��Լ�� >= n_ԤԼ��ʽ��Լ�� Then
            Return '13|��ǰ������λ��ԤԼ�����㡣';
          End If;
        End If;
      Else
        --��������ռ
        Begin
          Select ����
          Into n_����
          From �ٴ�����Һſ��Ƽ�¼
          Where ���� = 1 And ���� = 1 And ���� = ������λ_In And ��¼id = ��¼id_In;
        Exception
          When Others Then
            n_���� := Null;
        End;
        If n_���� Is Not Null Then
          If n_���Ʒ�ʽ = 1 Then
            n_ԤԼ��ʽ��Լ�� := Round(n_ԤԼ��ʽ��Լ�� * n_���� / 100);
          Else
            n_ԤԼ��ʽ��Լ�� := n_����;
          End If;
          Select Count(1)
          Into n_ԤԼ��ʽ��Լ��
          From ���˹Һż�¼
          Where �����¼id = ��¼id_In And ��¼״̬ = 1 And ������λ = ������λ_In;
          If n_ԤԼ��ʽ��Լ�� >= n_ԤԼ��ʽ��Լ�� Then
            Return '13|��ǰ������λ��ԤԼ�����㡣';
          End If;
        Else
          For r_���� In (Select ����, ����, ���� From �ٴ�����Һſ��Ƽ�¼ Where ���� = 1 And ��¼id = ��¼id_In) Loop
            If r_����.���� = 1 Then
              Select Count(1)
              Into n_��ʹ��
              From ���˹Һż�¼
              Where �����¼id = ��¼id_In And ������λ = r_����.���� And ��¼״̬ = 1;
            Else
              Select Count(1)
              Into n_��ʹ��
              From ���˹Һż�¼
              Where �����¼id = ��¼id_In And ԤԼ��ʽ = r_����.���� And ��¼״̬ = 1;
            End If;
            If n_���Ʒ�ʽ = 1 Then
              n_�������� := Nvl(n_��������, 0) + Round(r_����.���� * n_ԤԼ��ʽ��Լ�� / 100) - Nvl(n_��ʹ��, 0);
            Else
              n_�������� := Nvl(n_��������, 0) + r_����.���� - Nvl(n_��ʹ��, 0);
            End If;
          End Loop;
          Select Count(1) Into n_��ʹ�� From ���˹Һż�¼ Where �����¼id = ��¼id_In And ��¼״̬ = 1;
          If n_ԤԼ��ʽ��Լ�� - n_�������� - n_��ʹ�� > 0 Then
            Null;
          Else
            Return '13|��ǰ������λ��ԤԼ�����㡣';
          End If;
        End If;
      End If;
    End If;
    If n_���Ʒ�ʽ = 3 Then
      If n_��ſ��� = 1 Then
        Begin
          Select ����, ����, ����
          Into n_ԤԼ��ʽ��Լ��, v_ԤԼ��ʽ, n_����
          From �ٴ�����Һſ��Ƽ�¼
          Where ���� = 1 And ��¼id = ��¼id_In And ��� = ���_In;
        Exception
          When Others Then
            n_ԤԼ��ʽ��Լ�� := Null;
        End;
        If n_ԤԼ��ʽ��Լ�� Is Not Null Then
          If v_ԤԼ��ʽ <> ������λ_In Or n_���� = 1 Then
            Return '13|��ǰ������λ��ԤԼ�����㡣';
          End If;
          Select Nvl(Max(1), 0)
          Into n_ԤԼ��ʽ��Լ��
          From ���˹Һż�¼
          Where �����¼id = ��¼id_In And ���� = ���_In;
          If n_ԤԼ��ʽ��Լ�� >= n_ԤԼ��ʽ��Լ�� Then
            Return '13|��ǰ������λ��ԤԼ�����㡣';
          End If;
        End If;
      Else
        For r_���� In (Select ����, ����, ����
                     From �ٴ�����Һſ��Ƽ�¼
                     Where ���� = 1 And ��¼id = ��¼id_In And ��� = ���_In) Loop
          If r_����.���� <> ������λ_In Or r_����.���� = 1 Then
            If r_����.���� = 1 Then
              Select Count(1)
              Into n_��ʹ��
              From �ٴ�������ſ��� A, ���˹Һż�¼ B
              Where a.��¼id = ��¼id_In And a.ԤԼ˳��� Is Not Null And Nvl(a.�Һ�״̬, 0) <> 0 And a.��ע = b.���� And
                    b.������λ = r_����.���� And b.��¼״̬ = 1;
            Else
              Select Count(1)
              Into n_��ʹ��
              From �ٴ�������ſ��� A, ���˹Һż�¼ B
              Where a.��¼id = ��¼id_In And a.ԤԼ˳��� Is Not Null And Nvl(a.�Һ�״̬, 0) <> 0 And a.��ע = b.���� And
                    b.ԤԼ��ʽ = r_����.���� And b.��¼״̬ = 1;
            End If;
            n_�������� := Nvl(n_��������, 0) + r_����.���� - Nvl(n_��ʹ��, 0);
          Else
            Select Count(1)
            Into n_ԤԼ��ʽ��Լ��
            From �ٴ�������ſ��� A, ���˹Һż�¼ B
            Where a.��¼id = ��¼id_In And a.ԤԼ˳��� Is Not Null And Nvl(a.�Һ�״̬, 0) <> 0 And a.��ע = b.���� And b.������λ = ������λ_In And
                  b.��¼״̬ = 1;
            If n_ԤԼ��ʽ��Լ�� >= n_ԤԼ��ʽ��Լ�� Then
              Return '13|��ǰ������λ��ԤԼ�����㡣';
            End If;
          End If;
        End Loop;
        Select Count(1)
        Into n_��ʹ��
        From �ٴ�������ſ��� A
        Where a.��¼id = ��¼id_In And a.ԤԼ˳��� Is Not Null And Nvl(a.�Һ�״̬, 0) <> 0 And ��� = ���_In;
        Select Nvl(��Լ��, �޺���) Into n_ԤԼ��ʽ��Լ�� From �ٴ������¼ Where ID = ��¼id_In;
        If n_ԤԼ��ʽ��Լ�� - n_�������� - n_��ʹ�� > 0 Then
          Null;
        Else
          Return '13|��ǰ������λ��ԤԼ�����㡣';
        End If;
      End If;
    End If;
  End If;

  --0-����ԤԼ����;1-�úű��ֹԤԼ;2-����ֹ��������ƽ̨��ԤԼ
  If Nvl(n_ԤԼ����, 0) = 1 Then
    Return '2|��ǰ�ϰ�ʱ�ν�ֹԤԼ��';
  End If;

  If d_ͣ�￪ʼʱ�� Is Not Null And Not (Nvl(n_��ſ���, 0) = 1 And Nvl(n_�Ƿ��ʱ��, 0) = 1) Then
    If ԤԼʱ��_In >= d_ͣ�￪ʼʱ�� And ԤԼʱ��_In <= d_ͣ����ֹʱ�� Then
      Return '3|��ǰ�ϰ�ʱ����ԤԼʱ��ʱ��ͣ�����ԤԼ��';
    End If;
  End If;

  If Nvl(n_��Լ��, 0) > 0 Then
    If Nvl(n_��Լ��, 0) - Nvl(n_��Լ��, 0) <= 0 Then
      Return '4|��ǰ�ϰ�ʱ��ʣ���ԤԼ��Ϊ�㣬���ܼ���ԤԼ��';
    End If;
  End If;

  If Nvl(n_�Ƿ��ʱ��, 0) = 0 Then
    --����ʱ��
    Begin
      Select Nvl(b.���տ���״̬, 0) Into n_���տ���״̬ From �ٴ������Դ B Where b.Id = n_��Դid;
    Exception
      When Others Then
        n_���տ���״̬ := 0;
    End;

    --1.���Ұ���ԤԼʱ��Ľڼ���
    Begin
      Select a.����ԤԼ����, a.����Һ�����
      Into v_����ԤԼ, v_����Һ�
      From �������ձ� A
      Where a.���� = 0 And ԤԼʱ��_In Between a.��ʼ���� And a.��ֹ���� + 1 - 1 / 24 / 60 / 60 And Rownum < 2;
    Exception
      When Others Then
        Return '0|����ԤԼ��';
    End;

    --���տ���״̬��0-���ϰ�;1-�ϰ��ҿ���ԤԼ;2-�ϰ൫������ԤԼ;3-�ܽڼ������ÿ���
    If Nvl(n_���տ���״̬, 0) = 0 Then
      --���ϰ�Ŀ϶��ǲ���ԤԼ��
      Return '5|��ǰԤԼʱ���ڷ����ڼ���ʱ�䷶Χ�ڣ����ϰࡣ';
    Elsif Nvl(n_���տ���״̬, 0) = 1 Then
      Return '0|����ԤԼ��';
    Elsif Nvl(n_���տ���״̬, 0) = 2 Then
      --�ڽڼ���ʱ�䷶Χ�ڣ�����ԤԼ
      Return '6|��ǰԤԼʱ���ڷ����ڼ���ʱ�䷶Χ�ڣ���ֹԤԼ��';
    Elsif Nvl(n_���տ���״̬, 0) = 3 Then
      --û��"����Һ�"��һ��û��"����ԤԼ"
      If v_����Һ� Is Not Null Then
        --2.����Ƿ��а���ԤԼʱ���"����Һ�"
        Select Max(1)
        Into n_Count
        From Table(f_Str2list(v_����Һ�, ';'))
        Where ԤԼʱ��_In Between To_Date(Column_Value, 'yyyy-mm-dd') And
              To_Date(Column_Value, 'yyyy-mm-dd') + 1 - 1 / 24 / 60 / 60 And Rownum < 2;

        If Nvl(n_Count, 0) <> 0 Then
          --3.����Ƿ��а���ԤԼʱ���"����ԤԼ"
          Select Max(1)
          Into n_Count
          From Table(f_Str2list(v_����ԤԼ, ';'))
          Where ԤԼʱ��_In Between To_Date(Column_Value, 'yyyy-mm-dd') And
                To_Date(Column_Value, 'yyyy-mm-dd') + 1 - 1 / 24 / 60 / 60 And Rownum < 2;

          If Nvl(n_Count, 0) = 0 Then
            --����"����ԤԼ"ʱ�䷶Χ�ڣ�����ԤԼ
            Return '7|��ǰԤԼʱ���ڷ����ڼ��ղ�����ԤԼ��ʱ�䷶Χ�ڣ�����ԤԼ��';
          Else
            Return '0|����ԤԼ��';
          End If;
        Else
          Return '8|��ǰԤԼʱ���ڷ����ڼ��ղ�����Һŵ�ʱ�䷶Χ�ڣ�����ԤԼ��';
        End If;
      Else
        --û������"����Һ�"/"����ԤԼ"��ʾͣ��϶�����ԤԼ
        Return '9|��ǰԤԼʱ���ڷ����ڼ���ʱ�䷶Χ�ڣ���ͣ�����ԤԼ��';
      End If;
    End If;
  Else
    --��ʱ��
    If Nvl(���_In, 0) <> 0 Then
      Begin
        Select Nvl(�Ƿ�ԤԼ, 0), Nvl(�Һ�״̬, 0), ����Ա����, ����վ����
        Into n_�Ƿ�ԤԼ, n_�Һ�״̬, v_���Ų���Ա, v_���Ż�����
        From �ٴ�������ſ���
        Where ��¼id = ��¼id_In And ��� = ���_In;
      Exception
        When Others Then
          Return '16|��ǰѡ�����Ų����á�';
      End;
      If n_�Ƿ�ԤԼ = 0 Then
        Return '14|��ǰѡ�����Ž�ֹԤԼ��';
      End If;
      If n_�Һ�״̬ <> 0 Then
        If n_�Һ�״̬ = 5 And (Nvl(����Ա����_In, '-') <> Nvl(v_���Ų���Ա, '_') Or Nvl(������_In, '-') <> Nvl(v_���Ż�����, '_')) Then
           Return '15|��ǰѡ�������Ѿ���'|| Nvl(v_���Ż�����,'') ||'������';
        Elsif n_�Һ�״̬ <> 5 Then
           Return '15|��ǰѡ�������Ѿ���ʹ�á�';
        End if;
      End If;
    End If;
    Return '0|����ԤԼ��';
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Fun_Get�ٴ�����ԤԼ״̬;
/

--115264:Ƚ����,2017-12-11,סԺ���˰������շѣ�ֱ���շѣ������������ʱ�������־�ͱ�ʶ����д����
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
  n_�����־ ������ü�¼.�����־%Type;
  v_���ʽ ҽ�Ƹ��ʽ.����%Type;

  --��ʱ���� 
  n_Count      Number;
  v_����Ա���� ������ü�¼.����Ա����%Type;
  n_�²���ģʽ Number;
  v_����no     ҩƷ�շ���¼.No%Type;
  v_Date       Date;
  v_Err_Msg    Varchar2(255);
  Err_Item Exception;
  Err_Special Exception;
  n_��id ����ɿ����.Id%Type;
Begin
  n_��id := Zl_Get��id(����Ա����_In);

  Select Count(ID)
  Into n_Count
  From ������ü�¼
  Where ��¼���� = 1 And ��¼״̬ = 0 And NO = No_In And ����Ա���� Is Null;
  If n_Count = 0 Then
    Select Max(����Ա����) Into v_����Ա���� From ������ü�¼ Where ��¼���� = 1 And NO = No_In;
    If v_����Ա���� Is Not Null Then
      If v_����Ա���� = ����Ա����_In Then
        v_Err_Msg := '���ܶ�ȡ���۵�����,�õ����Ѿ����շѣ�';
        Raise Err_Special;
      Else
        v_Err_Msg := '���ܶ�ȡ���۵�����,�õ����Ѿ����շѣ�';
        Raise Err_Item;
      End If;
    Else
      v_Err_Msg := '���ܶ�ȡ���۵�����,�õ����Ѿ���ɾ����';
      Raise Err_Item;
    End If;
  End If;
  v_Date := �Ǽ�ʱ��_In;
  If v_Date Is Null Then
    Select Sysdate Into v_Date From Dual;
  End If;
  If Nvl(����id_In, 0) <> 0 Then
    --���������־��ȡ�����/סԺ��
    Select Max(�����־), Max(��ʶ��) Into n_�����־, v_��ʶ�� From ������ü�¼ Where ��¼���� = 1 And NO = No_In;
    If v_��ʶ�� Is Null Then
      Select Decode(n_�����־, 2, סԺ��, �����) Into v_��ʶ�� From ������Ϣ Where ����id = ����id_In;
    End If;
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
      Set ��¼״̬ = 1, ����id = Decode(����id_In, 0, Null, ����id_In), ��ʶ�� = Nvl(��ʶ��, v_��ʶ��), ���ʽ = ���ʽ_In, ���� = ����_In,
          ���� = ����_In, �Ա� = �Ա�_In,
          --���ܱ���ҽ�����͵����� 
          ���˿���id = Nvl(���˿���id_In, ���˿���id), ��������id = Nvl(��������id_In, ��������id), ������ = Nvl(������_In, ������), ���ʽ�� = ʵ�ս��,
          ����id = ����id_In, ����ʱ�� = ����ʱ��_In, �Ǽ�ʱ�� = v_Date, ����Ա��� = ����Ա���_In, ����Ա���� = ����Ա����_In, �Ƿ��� = �Ƿ���_In,
          �ɿ���id = n_��id, ����״̬ = 1, ִ��״̬ = Decode(Nvl(ִ��״̬, 0), -1, Null, Nvl(ִ��״̬, 0))
      Where ID = t_����id(I) And ��¼״̬ = 0;
  
    If Sql%RowCount <> t_����id.Count Then
      Select Count(1)
      Into n_Count
      From ������ü�¼
      Where ��¼״̬ = 1 And ID In (Select Column_Value From Table(t_����id));
      If n_Count <> t_����id.Count Then
        v_Err_Msg := '���ڲ�������,�õ����Ѿ�ɾ����';
        Raise Err_Item;
      Else
        Select Max(����Ա����)
        Into v_����Ա����
        From ������ü�¼
        Where ��¼״̬ = 1 And ID In (Select Column_Value From Table(t_����id));
        If v_����Ա���� = ����Ա����_In Then
          v_Err_Msg := '���ڲ�������,�õ����Ѿ��շѣ�';
          Raise Err_Special;
        Else
          v_Err_Msg := '���ڲ�������,�õ����Ѿ��շѣ�';
          Raise Err_Item;
        End If;
      End If;
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
  When Err_Special Then
    Raise_Application_Error(-20105, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˻����շ�_Insert;
/

--118128:����,2017-12-11,���������ε���0����
Create Or Replace Procedure Zl_�����������_Insert
(
  No_In         In ҩƷ�շ���¼.No%Type,
  ���_In       In ҩƷ�շ���¼.���%Type,
  �ⷿid_In     In ҩƷ�շ���¼.�ⷿid%Type,
  ������id_In In ҩƷ�շ���¼.������id%Type,
  ����id_In     In ҩƷ�շ���¼.ҩƷid%Type,
  ʵ������_In   In ҩƷ�շ���¼.ʵ������%Type,
  �ɱ���_In     In ҩƷ�շ���¼.�ɱ���%Type,
  �ɱ����_In   In ҩƷ�շ���¼.�ɱ����%Type,
  ���ۼ�_In     In ҩƷ�շ���¼.���ۼ�%Type,
  ���۽��_In   In ҩƷ�շ���¼.���۽��%Type,
  ���_In       In ҩƷ�շ���¼.���%Type,
  ���۲��_In   In ҩƷ�շ���¼.���%Type,
  ������_In     In ҩƷ�շ���¼.������%Type,
  ��������_In   In ҩƷ�շ���¼.��������%Type,
  ժҪ_In       In ҩƷ�շ���¼.ժҪ%Type := Null,
  ����_In       In ҩƷ�շ���¼.����%Type := Null,
  ����_In       In ҩƷ�շ���¼.����%Type := Null,
  ��������_In   In ҩƷ�շ���¼.��������%Type := Null,
  Ч��_In       In ҩƷ�շ���¼.Ч��%Type := Null,
  �������_In   In ҩƷ�շ���¼.�������%Type := Null,
  ���Ч��_In   In ҩƷ�շ���¼.���Ч��%Type := Null,
  ��Ʒ����_In   In ҩƷ�շ���¼.��Ʒ����%Type := Null,
  ��׼�ĺ�_In   In ҩƷ�շ���¼.��׼�ĺ�%Type := Null
) Is
  n_Id       ҩƷ�շ���¼.Id%Type; --�շ�ID
  n_���ϵ�� ҩƷ�շ���¼.���ϵ��%Type;
  n_����     ҩƷ�շ���¼.����%Type := Null; --����
  n_�ⷿ���� Integer; --�Ƿ��������    1:����;0��������
  n_���÷��� Integer; --�Ƿ��������    1:����;0��������
Begin
  If Not ��׼�ĺ�_In Is Null And Not ����_In Is Null Then
    Update ҩƷ�����̶��� Set ��׼�ĺ� = ��׼�ĺ�_In Where ҩƷid = ����id_In And �������� = ����_In;
    If Sql%RowCount = 0 Then
      Insert Into ҩƷ�����̶��� (ҩƷid, ��������, ��׼�ĺ�) Values (����id_In, ����_In, ��׼�ĺ�_In);
    End If;
  End If;

  n_���ϵ�� := 1;

  Select ҩƷ�շ���¼_Id.Nextval Into n_Id From Dual;

  Select Nvl(�ⷿ����, 0), Nvl(���÷���, 0) Into n_�ⷿ����, n_���÷��� From �������� Where ����id = ����id_In;

  If n_���÷��� = 0 Then
    If n_�ⷿ���� = 1 Then
      Begin
        Select Distinct 0
        Into n_�ⷿ����
        From ��������˵��
        Where (�������� = '���ϲ���' Or �������� Like '�Ƽ���') And ����id = �ⷿid_In;
      Exception
        When Others Then
          n_�ⷿ���� := 1;
      End;
    
      If n_�ⷿ���� = 1 Then
        n_���� := n_Id;
      End If;
	Else
      n_���� := 0;
    End If;
  Else
    n_���� := n_Id;
  End If;
  Insert Into ҩƷ�շ���¼
    (ID, ��¼״̬, ����, NO, ���, �ⷿid, ������id, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, �������, ���Ч��, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ���۽��, ���,
     ժҪ, ������, ��������, ��������, �÷�, ��Ʒ����, ��׼�ĺ�)
  Values
    (n_Id, 1, 17, No_In, ���_In, �ⷿid_In, ������id_In, n_���ϵ��, ����id_In, n_����, ����_In, ����_In, Ч��_In, �������_In, ���Ч��_In,
     ʵ������_In, ʵ������_In, �ɱ���_In, �ɱ����_In, ���ۼ�_In, ���۽��_In, ���_In, ժҪ_In, ������_In, ��������_In, ��������_In, ���۲��_In, ��Ʒ����_In,
     ��׼�ĺ�_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����������_Insert;
/

--118128:����,2017-12-11,���������ε���0����
Create Or Replace Procedure Zl_�����⹺_Insert
(
  No_In         In ҩƷ�շ���¼.No%Type,
  ���_In       In ҩƷ�շ���¼.���%Type,
  �ⷿid_In     In ҩƷ�շ���¼.�ⷿid%Type,
  ��ҩ��λid_In In ҩƷ�շ���¼.��ҩ��λid%Type,
  ����id_In     In ҩƷ�շ���¼.ҩƷid%Type,
  ����_In       In ҩƷ�շ���¼.����%Type := Null,
  ����_In       In ҩƷ�շ���¼.����%Type := Null,
  ��������_In   In ҩƷ�շ���¼.��������%Type := Null,
  Ч��_In       In ҩƷ�շ���¼.Ч��%Type := Null,
  �������_In   In ҩƷ�շ���¼.�������%Type := Null,
  ���Ч��_In   In ҩƷ�շ���¼.���Ч��%Type := Null,
  ʵ������_In   In ҩƷ�շ���¼.ʵ������%Type := Null,
  �ɱ���_In     In ҩƷ�շ���¼.�ɱ���%Type := Null,
  �ɱ����_In   In ҩƷ�շ���¼.�ɱ����%Type := Null,
  ����_In       In ҩƷ�շ���¼.����%Type := Null,
  ���ۼ�_In     In ҩƷ�շ���¼.���ۼ�%Type := Null,
  ���۽��_In   In ҩƷ�շ���¼.���۽��%Type := Null,
  ���_In       In ҩƷ�շ���¼.���%Type := Null,
  ���۲��_In   In ҩƷ�շ���¼.���%Type := Null,
  ժҪ_In       In ҩƷ�շ���¼.ժҪ%Type := Null,
  ע��֤��_In   In ҩƷ�շ���¼.ע��֤��%Type := Null,
  ������_In     In ҩƷ�շ���¼.������%Type := Null,
  �������_In   In Ӧ����¼.�������%Type := Null,
  ��Ʊ��_In     In Ӧ����¼.��Ʊ��%Type := Null,
  ��Ʊ����_In   In Ӧ����¼.��Ʊ����%Type := Null,
  ��Ʊ���_In   In Ӧ����¼.��Ʊ���%Type := Null,
  ��������_In   In ҩƷ�շ���¼.��������%Type := Null,
  �˲���_In     In ҩƷ�շ���¼.��ҩ��%Type := Null,
  �˲�����_In   In ҩƷ�շ���¼.��ҩ����%Type := Null,
  ����_In       In ҩƷ�շ���¼.����%Type := 0,
  �˻�_In       In Number := 1,
  ��ֵ����_In   In Varchar2 := Null,
  ��Ʒ����_In   In ҩƷ�շ���¼.��Ʒ����%Type := Null,
  �ڲ�����_In   In ҩƷ�շ���¼.�ڲ�����%Type := Null,
  ����id_In     In ҩƷ�շ���¼.����id%Type := 0,
  ��Ʊ����_In   In Ӧ����¼.��Ʊ����%Type := Null,
  �������_In   In Number := 0,
  ��׼�ĺ�_In   In ҩƷ�շ���¼.��׼�ĺ�%Type := Null,
  ���ս���_In   In ҩƷ�շ���¼.���ս���%Type := Null
) Is
  Err_Item Exception;
  v_Err_Msg Varchar2(200);

  v_No         Ӧ����¼.No%Type; --Ӧ����¼��NO 
  v_��Ʒ��     �շ���ĿĿ¼.����%Type; --ͨ������ 
  v_���       �շ���ĿĿ¼.���%Type;
  v_����       �շ���ĿĿ¼.���%Type;
  v_��λ       �շ���ĿĿ¼.���㵥λ%Type;
  v_Lngid      ҩƷ�շ���¼.Id%Type; --�շ�ID 
  n_Ӧ��id     Ӧ����¼.Id%Type; --Ӧ����¼��ID 
  n_������id ҩƷ�շ���¼.������id%Type; --������ID 
  n_���ϵ��   ҩƷ�շ���¼.���ϵ��%Type; --���ϵ�� 
  n_����       ҩƷ�շ���¼.����%Type := Null; --���� 
  n_�ⷿ����   Integer; --�Ƿ��������    1:������0�������� 
  n_���÷���   Integer; --�Ƿ����÷���       1:������0�������� 
  v_��������   ҩƷ���.��������%Type;
Begin

  If Not ��׼�ĺ�_In Is Null And Not ����_In Is Null Then
    Update ҩƷ�����̶��� Set ��׼�ĺ� = ��׼�ĺ�_In Where ҩƷid = ����id_In And �������� = ����_In;
    If Sql%RowCount = 0 Then
      Insert Into ҩƷ�����̶��� (ҩƷid, ��������, ��׼�ĺ�) Values (����id_In, ����_In, ��׼�ĺ�_In);
    End If;
  End If;

  --ȡ�ò��ϵ����� 
  v_���� := '';
  Select ����, ���, ���㵥λ Into v_��Ʒ��, v_���, v_��λ From �շ���ĿĿ¼ Where ID = ����id_In;

  If v_��� Is Not Null Then
    If Instr(v_���, '|') <> 0 Then
      v_���� := Substr(v_���, Instr(v_���, '|'));
      v_��� := Substr(v_���, Instr(v_���, '|') - 1);
    End If;
  End If;

  Select ҩƷ�շ���¼_Id.Nextval Into v_Lngid From Dual;

  Select Nvl(�ⷿ����, 0), Nvl(���÷���, 0) Into n_�ⷿ����, n_���÷��� From �������� Where ����id = ����id_In;

  --�������ֱ���ô�����������
  If �������_In = 0 Then
    If ����id_In > 0 And ����_In > 0 Then
      n_���� := ����_In;
    Else
      If n_���÷��� = 0 Then
        If n_�ⷿ���� = 1 Then
          Begin
            Select Distinct 0
            Into n_�ⷿ����
            From ��������˵��
            Where ((�������� Like '���ϲ���') Or (�������� Like '�Ƽ���')) And ����id = �ⷿid_In;
          Exception
            When Others Then
              n_�ⷿ���� := 1;
          End;
        
          If n_�ⷿ���� = 1 Then
            n_���� := v_Lngid;
          End If;
		Else
          n_���� := 0;
        End If;
      Else
        n_���� := v_Lngid;
      End If;
    End If;
  Else
    n_���� := ����_In;
  End If;

  Select b.Id, b.ϵ��
  Into n_������id, n_���ϵ��
  From ҩƷ�������� A, ҩƷ������ B
  Where a.���id = b.Id And a.���� = 30 And Rownum < 2;

  Insert Into ҩƷ�շ���¼
    (ID, ��¼״̬, ����, NO, ���, �ⷿid, ��ҩ��λid, ������id, ���ϵ��, ҩƷid, ����, ����, ����, ��������, Ч��, �������, ���Ч��, ��д����, ʵ������, �ɱ���, �ɱ����,
     ����, ���ۼ�, ���۽��, ���, ժҪ, ������, ��������, ��ҩ��ʽ, ��ҩ��, ��ҩ����, ע��֤��, �÷�, ��Ʒ����, �ڲ�����, ����id, ��׼�ĺ�, ���ս���)
  Values
    (v_Lngid, 1, 15, No_In, ���_In, �ⷿid_In, ��ҩ��λid_In, n_������id, n_���ϵ��, ����id_In, Decode(�˻�_In, -1, ����_In, n_����), ����_In,
     ����_In, ��������_In, Ч��_In, �������_In, ���Ч��_In, �˻�_In * ʵ������_In, �˻�_In * ʵ������_In, �ɱ���_In, �˻�_In * �ɱ����_In, ����_In, ���ۼ�_In,
     �˻�_In * ���۽��_In, �˻�_In * ���_In, ժҪ_In, ������_In, ��������_In, Decode(�˻�_In, -1, 1, 0), �˲���_In, �˲�����_In, ע��֤��_In, ���۲��_In,
     ��Ʒ����_In, �ڲ�����_In, ����id_In, ��׼�ĺ�_In, ���ս���_In);

  --��ֵ������Ϣ 
  If Length(��ֵ����_In) > 0 Then
    Insert Into �շ���¼������Ϣ
      (�շ�id, ����, ��������, סԺ��, ����)
    Values
      (v_Lngid, Substr(��ֵ����_In, 1, Instr(��ֵ����_In, ',', 1, 1) - 1),
       Substr(��ֵ����_In, Instr(��ֵ����_In, ',', 1, 1) + 1, Instr(��ֵ����_In, ',', 1, 2) - Instr(��ֵ����_In, ',', 1, 1) - 1),
       Substr(��ֵ����_In, Instr(��ֵ����_In, ',', 1, 2) + 1, Instr(��ֵ����_In, ',', 1, 3) - Instr(��ֵ����_In, ',', 1, 2) - 1),
       Substr(��ֵ����_In, Instr(��ֵ����_In, ',', 1, 3) + 1, Length(��ֵ����_In)));
  End If;

  If ��Ʊ��_In Is Not Null Or �������_In Is Not Null Then
  
    Select Ӧ����¼_Id.Nextval Into n_Ӧ��id From Dual;
  
    --����ǵ�һ����ϸ,�����Ӧ����¼��NO 
    Begin
      Select NO
      Into v_No
      From Ӧ����¼
      Where ϵͳ��ʶ = 5 And ��¼���� = 0 And ��¼״̬ = 1 And ��ⵥ�ݺ� = No_In And Rownum < 2;
    Exception
      When Others Then
        v_No := Nextno(67);
    End;
  
    Insert Into Ӧ����¼
      (ID, ��¼����, ��¼״̬, ��Ŀid, ���, ��λid, NO, ϵͳ��ʶ, �շ�id, ��ⵥ�ݺ�, ���ݽ��, �������, ��Ʊ��, ��Ʊ����, ��Ʊ���, Ʒ��, ���, ����, ����, ������λ, ����,
       �ɹ���, �ɹ����, ������, ��������, �����, �������, ժҪ, �ⷿid, ��Ʊ����)
    Values
      (n_Ӧ��id, 0, 1, ����id_In, ���_In, ��ҩ��λid_In, v_No, 5, v_Lngid, No_In, �˻�_In * ���۽��_In, �������_In, ��Ʊ��_In, ��Ʊ����_In,
       �˻�_In * Decode(��Ʊ��_In, Null, �ɱ����_In, ��Ʊ���_In), v_��Ʒ��, v_���, v_����, ����_In, v_��λ, �˻�_In * ʵ������_In, �ɱ���_In,
       �˻�_In * �ɱ����_In, ������_In, ��������_In, Null, Null, ժҪ_In, �ⷿid_In, ��Ʊ����_In);
  End If;

  --�˻�ʱ�¿������� 
  If �˻�_In = -1 And Nvl(����id_In, 0) <> 2 Then
    --����� 
    Begin
      Select Nvl(��������, 0)
      Into v_��������
      From ҩƷ���
      Where �ⷿid = �ⷿid_In And ҩƷid = ����id_In And Nvl(����, 0) = Nvl(����_In, 0) And ���� = 1;
    Exception
      When Others Then
        v_�������� := 0;
    End;
  
    If v_�������� - ʵ������_In < 0 Then
      v_Err_Msg := '[ZLSOFT]��' || ���_In || '�еĿ�����������,����[ZLSOFT]';
      Raise Err_Item;
    End If;
    Update ҩƷ���
    Set �������� = Nvl(��������, 0) - ʵ������_In
    Where �ⷿid = �ⷿid_In And ҩƷid = ����id_In And Nvl(����, 0) = ����_In And ���� = 1;
    Delete From ҩƷ���
    Where �ⷿid = �ⷿid_In And ҩƷid = ����id_In And Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0 And
          Nvl(ʵ�ʲ��, 0) = 0;
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����⹺_Insert;
/

--114951:����,2017-12-08,���µ�����¼���������������Զ����

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
      End Case;
    End Loop;
  Else
    n_��ʼʱ�� := 4;
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
--114951:����,2017-12-08,���µ�����¼���������������Զ����
CREATE OR REPLACE Procedure Zl_���µ�����_Update
(
  �ļ�id_In   In ���˻����ļ�.Id%Type, --���˻����ļ�ID
  ����ʱ��_In In ���˻�������.����ʱ��%Type, --�������ݵķ���ʱ��
  ��¼����_In In ���˻�����ϸ.��¼����%Type, --������Ŀ=1���ϱ�˵��=2�����ת���=3�������ձ��=4,�±�˵��=6
  ��Ŀ���_In In ���˻�����ϸ.��Ŀ���%Type, --������Ŀ����ţ��ǻ�����Ŀ�̶�Ϊ0
  ��¼����_In In ���˻�����ϸ.��¼����%Type := Null, --��¼���ݣ��������Ϊ�գ��������ǰ������  36��36/37
  ���²�λ_In In ���˻�����ϸ.���²�λ%Type := Null, --ɾ������ʱ������д��λ �����Ŀ��
  ���Ժϸ�_In In Number := 0,
  δ��˵��_In In ���˻�����ϸ.δ��˵��%Type := Null, --δ��˵��
  ���˼�¼_In In Number := 1,
  ������Դ_In In ���˻�����ϸ.������Դ%Type := 0,
  ��Դid_In   In ���˻�����ϸ.��Դid%Type := Null, --ʼ��Ϊԭʼ��¼����ԴID
  ����_In     In ���˻�����ϸ.����%Type := 0,
  ��Ŀ�״�_In In Number := 0, --������Ŀʹ�ã���������ǰ�Ƿ���ɾ��һ��ʱ���ڵ�������Ϣ�� 1 ɾ��
  ��ʼʱ��_In In ���˻�������.����ʱ��%Type := Null, --����¼��Ч��ȵĿ�ʼʱ��
  ����ʱ��_In In ���˻�������.����ʱ��%Type := Null, --����¼��Ч��ȵ���ֹʱ�䣬������¼Ϊÿ���ӣ����±�Ϊ4Сʱ,ʱ�����ڵ���ͬ��Ŀ��¼Ҫɾ��


  ����Ա_In   In ���˻�������.������%Type := Null,
  ������_In In Number := 1,
  ��ʾ_In     In Number := 0,
  ����_In     In Number := 0
) Is
  n_��Ŀ��� ���˻�����ϸ.��Ŀ���%Type;
  n_��¼��� ���˻�����ϸ.��¼���%Type; --��¼���ݵ������־
  v_������   ���˻�������.������%Type;
  v_��¼��   ���˻�����ϸ.��¼��%Type;
  d_����ʱ�� ���˻�������.����ʱ��%Type;
  d_����ʱ�� ���˻�������.����ʱ��%Type;
  d_��ʼʱ�� ���˻�������.����ʱ��%Type;
  n_��¼id   ���˻�����ϸ.��¼id%Type;
  v_����id   ���˻����ļ�.����id%Type;
  n_����Ӧ�� �����¼��Ŀ.Ӧ�÷�ʽ%Type;
  n_����     �����¼��Ŀ.��Ŀ���%Type := 2;
  n_����     �����¼��Ŀ.��Ŀ���%Type := 1;
  n_����     �����¼��Ŀ.��Ŀ���%Type := -1;
  n_��Ŀ���� �����¼��Ŀ.��Ŀ����%Type := 1;
  n_��ʼ�汾 ���˻�����ϸ.��ʼ�汾%Type;
  n_��ʹǿ�� �����¼��Ŀ.��Ŀ���%Type;
  n_Newid    ���˻�����ϸ.Id%Type;

  n_����id       ���˻����ļ�.����id%Type;
  n_��ҳid       ���˻����ļ�.��ҳid%Type;
  n_Ӥ��         ���˻����ļ�.Ӥ��%Type;
  d_Ӥ����Ժʱ�� ����ҽ����¼.��ʼִ��ʱ��%Type;
  d_�ļ���ʼʱ�� ���˻����ļ�.��ʼʱ��%Type;
  n_Preblue      Number;
  n_i            Number;
  n_Sqlrowcount  Number;
  n_Count        Number(1);
  v_��¼����     ���˻�����ϸ.��¼����%Type;
  v_Data         ���˻�����ϸ.��¼����%Type;
  --������
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  d_����ʱ�� := ����ʱ��_In;

  If d_����ʱ�� Is Null Then
    v_Error := '���ݷ���ʱ�䲻��Ϊ�գ�';
    Raise Err_Custom;
  End If;

  If ��ʼʱ��_In Is Null Then
    d_��ʼʱ�� := d_����ʱ��;
  Else
    d_��ʼʱ�� := ��ʼʱ��_In;
  End If;

  If ����ʱ��_In Is Null Then
    d_����ʱ�� := d_��ʼʱ��;
  Else
    d_����ʱ�� := ����ʱ��_In;
  End If;

  --��ȡ��¼ID
  n_��¼id := 0;
  If ����Ա_In Is Null Then
    v_������ := Zl_Username;
  Else
    v_������ := ����Ա_In;
  End If;
  ----------------------------------------------------------------------------------------------------------------------
  Begin
    Select ID Into n_��¼id From ���˻������� Where �ļ�id = �ļ�id_In And ����ʱ�� = ����ʱ��_In;
  Exception
    When Others Then
      n_��¼id := 0;
  End;

  --������ݵķ���ʱ���Ƿ��Ӧ����
  ---------------------------------------------------------------------------------------------------------------------
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
            (����ʱ��_In >= a.��ʼʱ�� And (����ʱ��_In < = Nvl(a.��ֹʱ��, Sysdate) Or a.��ֹʱ�� Is Null)) And Rownum < 2;
    Exception
      When Others Then
        v_����id := 0;
    End;
    If v_����id = 0 And ������_In = 1 Then
      v_Error := '���ݷ���ʱ�� ' || To_Char(����ʱ��_In, 'YYYY-MM-DD HH24:MI:SS') || ' ���ڲ�����Ч�䶯ʱ�䷶Χ�ڣ����ܲ�����';
      Raise Err_Custom;
    End If;
  Else
    If ����ʱ��_In < d_�ļ���ʼʱ�� Or ����ʱ��_In > d_Ӥ����Ժʱ�� Then
      v_Error := '���ݷ���ʱ�� ' || To_Char(����ʱ��_In, 'YYYY-MM-DD HH24:MI:SS') || ' ���ڲ�����Ч�䶯ʱ�䷶Χ�ڣ����ܲ�����';
      Raise Err_Custom;
    End If;
  End If;
  --����ǲ��Ǳ��˵ļ�¼
  ---------------------------------------------------------------------------------------------------------------------
  If ���˼�¼_In = 0 And n_��¼id > 0 Then
    v_��¼�� := '';
    Begin
      Select ��¼��
      Into v_��¼��
      From ���˻�����ϸ
      Where ��¼id = n_��¼id And ��¼���� = ��¼����_In And ��Ŀ��� = ��Ŀ���_In And ��ֹ�汾 Is Null And Rownum < 2
      Order By Nvl(��¼���, 0);
    Exception
      When Others Then
        v_��¼�� := '';
    End;
    If v_��¼�� Is Not Null And v_��¼�� <> v_������ Then
      v_Error := '��' || To_Char(d_��ʼʱ��, 'yyyy-mm-dd hh24:mi:ss') || '��' || To_Char(d_����ʱ��, 'yyyy-mm-dd hh24:mi:ss') ||
                 '���ڼ�¼�˲��ǵ�ǰ�ˣ�����Ȩ�޸ģ�';
      Raise Err_Custom;
    End If;
  End If;

  --��ȡ��ʹǿ��������Ŀ����Ŀ���
  Begin
    Select ��Ŀ��� Into n_��ʹǿ�� From ���¼�¼��Ŀ Where ��¼�� = '��ʹǿ��';
  Exception
    When Others Then
      n_��ʹǿ�� := -999;
  End;
  --������������Ƿ���
  If ��Ŀ���_In = n_���� Then
    n_��Ŀ��� := n_����;
  Else
    n_��Ŀ��� := ��Ŀ���_In;
  End If;
  Begin
    Select Ӧ�÷�ʽ, ��Ŀ���� Into n_����Ӧ��, n_��Ŀ���� From �����¼��Ŀ Where ��Ŀ��� = n_��Ŀ���;
  Exception
    When Others Then
      n_����Ӧ�� := 0;
  End;

  ----���ĳ��ʱ���ڵĻ���������Ϣ
  --��Ŀ�״�_In ������Ŀ���ݻ���ʱ��α���һ������ʱ������ڱ��� ��Ŀ�״�_In��=1
  --��¼����_In Is Null And δ��˵��_In Is Null ����Ϊɾ������
  ---------------------------------------------------------------------------------------------------------------------
  If (��Ŀ�״�_In = 1) Or (��¼����_In Is Null And δ��˵��_In Is Null) Then
    For r_List In (Select l.Id, Count(*) As ��¼��, Min(l.����ʱ��) ����ʱ��
                   From ���˻����ļ� A, ���˻������� L, ���˻�����ϸ D
                   Where a.Id = l.�ļ�id And l.Id = d.��¼id And a.Id = �ļ�id_In And d.��ֹ�汾 Is Null And l.����ʱ�� >= d_��ʼʱ�� And
                         l.����ʱ�� <= d_����ʱ��
                   Group By l.Id) Loop
      n_Sqlrowcount := 0;
      If ��¼����_In = 2 Or ��¼����_In = 6 Then
        Delete ���˻�����ϸ
        Where ��¼id = r_List.Id And ��¼���� = ��¼����_In And ��Ŀ��� = ��Ŀ���_In And ��ֹ�汾 Is Null;
        n_Sqlrowcount := Sql%RowCount;
      Else
        If ���²�λ_In Is Not Null Then
          --�˴���Ҫ��Ի��Ŀ
          Delete ���˻�����ϸ
          Where ��¼id = r_List.Id And ��¼���� = ��¼����_In And ��Ŀ��� = ��Ŀ���_In And Nvl(���²�λ, '��') = Nvl(���²�λ_In, '��') And
                ��ֹ�汾 Is Null;
        Else
          Delete ���˻�����ϸ
          Where ��¼id = r_List.Id And ��¼���� = ��¼����_In And ��Ŀ��� = ��Ŀ���_In And ��ֹ�汾 Is Null;
        End If;

        n_Sqlrowcount := Sql%RowCount;
        --������������ʹ���ɾ��������ͬʱɾ����������
        If ��Ŀ���_In = n_���� And n_����Ӧ�� = 2 Then
          Delete ���˻�����ϸ
          Where ��¼id = r_List.Id And ��¼���� = ��¼����_In And ��Ŀ��� = n_���� And ��ֹ�汾 Is Null;
          n_Sqlrowcount := n_Sqlrowcount + Sql%RowCount;
        End If;
        --���Ϊ����ѹ/����ѹɾ������ѹʱͬʱɾ������ѹ����
        If ��Ŀ���_In = 4 Then
          Delete ���˻�����ϸ
          Where ��¼id = r_List.Id And ��¼���� = ��¼����_In And ��Ŀ��� = 5 And ��ֹ�汾 Is Null;
          n_Sqlrowcount := n_Sqlrowcount + Sql%RowCount;
        End If;
      End If;
      If n_Sqlrowcount >= r_List.��¼�� Then
        Delete ���˻������� Where ID = r_List.Id;
      End If;
      --���´�ӡ
      Update ���µ���ӡ
      Set ��ӡ�� = Null, ��ӡʱ�� = Null
      Where �ļ�id = �ļ�id_In And
            ��ʼʱ�� = (Select Max(��ʼʱ��) From ���µ���ӡ Where �ļ�id = �ļ�id_In And ��ʼʱ�� <= r_List.����ʱ��);
    End Loop;
  End If;

  If ��¼����_In Is Null And δ��˵��_In Is Null Then
    Return;
  End If;

  --�ֽ���Ŀ��¼����
  n_Preblue := 0;
  If (��¼����_In = 1 Or ��¼����_In = 7) And Instr(',' || n_��ʹǿ�� || ',1,2,4,', ',' || ��Ŀ���_In || ',', 1) > 0 Then
    n_Preblue := Nvl(Instr(Nvl(��¼����_In, ''), '/', 1), 0);
    If n_Preblue > 1 Then
      n_Preblue := 1;
    End If;
  End If;

  If ��Ŀ���_In = 4 And n_Preblue = 0 Then
    v_Error := 'Ѫѹ���ݸ�ʽ����! ��ʽ:����ѹ/����ѹ��';
    Raise Err_Custom;
  End If;

  --ȷ�Ͽ�ʼ�汾��
  ---------------------------------------------------------------------------------------------------------------------
  n_��ʼ�汾 := 1;

  --��д���˻������ݣ�����Ѿ������벡�ˡ����Һͷ���ʱ����ͬ�ļ�¼���޸ģ����������µļ�¼
  ---------------------------------------------------------------------------------------------------------------------
  --������Ŀ��ɾ���������ӣ����ܿ�ʼ��ȡ�ļ�¼ID�Ѿ������ڡ�
  Begin
    Select ID Into n_��¼id From ���˻������� Where ID = n_��¼id;
  Exception
    When Others Then
      n_��¼id := 0;
  End;

  If n_��¼id = 0 Then
    Select ���˻�������_Id.Nextval Into n_��¼id From Dual;
    Insert Into ���˻�������
      (ID, �ļ�id, ��ʾ, ����ʱ��, ������, ����ʱ��, ���汾)
    Values
      (n_��¼id, �ļ�id_In, 0, d_����ʱ��, v_������, Sysdate, n_��ʼ�汾);
  End If;

  --���ɾ�����������ݻ�������������
  If (��Ŀ���_In = n_���� Or ��Ŀ���_In = n_��ʹǿ�� Or (��Ŀ���_In = n_���� And n_����Ӧ�� = 2)) And n_Preblue = 0 Then
    Delete From ���˻�����ϸ
    Where ��¼id = n_��¼id And ��Ŀ��� = Decode(��Ŀ���_In, n_����, n_����, ��Ŀ���_In) And Decode(��Ŀ���_In, n_����, 1, Nvl(��¼���, 0)) = 1 And
          ��¼���� = ��¼����_In And ��ֹ�汾 Is Null;
  End If;

  --��д���˻�����ϸ������Ѿ������벡�ˡ����Һͷ���ʱ����ͬ�ļ�¼���޸ģ����������µļ�¼
  -----------------------------------------------------------------------------------------------------------------------
  v_Data     := ��¼����_In;
  n_��Ŀ��� := ��Ŀ���_In;
  For n_i In 0 .. n_Preblue Loop
    If n_i = 0 Then
      If ��Ŀ���_In = n_���� Then
        n_��¼��� := 1;
      Else
        n_��¼��� := 0;
      End If;
    Else
      --����ѹ/����ѹ
      If ��Ŀ���_In = 4 Then
        n_��¼��� := 0;
        n_��Ŀ��� := 5;
      Else
        n_��¼��� := 1;
        If ��Ŀ���_In = n_���� Then
          n_��Ŀ��� := n_����;
        End If;
      End If;

    End If;
    If n_Preblue > 0 Then
      v_��¼���� := Substr(v_Data, 1, Instr(v_Data, '/', 1) - 1);
      If v_��¼���� Is Null Then
        v_��¼���� := v_Data;
      End If;
    Else
      v_��¼���� := v_Data;
    End If;

    --����Ƿ���Ҫ�������
    Select Count(b.��¼��)
    Into n_Count
    From �����¼��Ŀ A, ���¼�¼��Ŀ B
    Where a.��Ŀ��� = b.��Ŀ��� And a.��Ŀ��� = ��Ŀ���_In And b.��¼�� = 1 And a.������ = '1)����������Ŀ';

    --Ϊ�˼�����ǰͬ���������������ݼ�¼���Ϊ0
    if ��Ŀ���_In=0  then
    If n_i = 0 Then
      Update ���˻�����ϸ
      Set ��¼���� = v_��¼����, ���²�λ = ���²�λ_In, ���Ժϸ� = ���Ժϸ�_In,
          δ��˵�� = Decode(n_��Ŀ���, n_����, Decode(v_��¼����, '����', Null, δ��˵��_In), δ��˵��_In), ��¼�� = v_������, ��¼ʱ�� = Sysdate
      Where ��¼id = n_��¼id And ��Ŀ��� = n_��Ŀ��� And ��¼���� = ��¼����_In And
            Decode(��Ŀ���_In, n_����, Nvl(��¼���, 0), n_��ʹǿ��, Nvl(��¼���, 0), Nvl(n_��¼���, 0)) = Nvl(n_��¼���, 0) And
            Decode(n_��Ŀ����, 2, Nvl(���²�λ, '��'), Nvl(���²�λ_In, '��')) = Nvl(���²�λ_In, '��') And ��ֹ�汾 Is Null;
    Else
      Update ���˻�����ϸ
      Set ��¼���� = v_��¼����, ��¼�� = v_������, ��¼ʱ�� = Sysdate
      Where ��¼id = n_��¼id And ��Ŀ��� = n_��Ŀ��� And ��¼���� = ��¼����_In And
            Decode(��Ŀ���_In, n_����, Nvl(��¼���, 0), n_��ʹǿ��, Nvl(��¼���, 0), Nvl(n_��¼���, 0)) = Nvl(n_��¼���, 0) And ��ֹ�汾 Is Null;
    End If;
    else
      If n_i = 0 Then
      Update ���˻�����ϸ
      Set ��¼���� = v_��¼����, ���²�λ = ���²�λ_In, ���Ժϸ� = ���Ժϸ�_In, ��¼���� = ��¼����_In,
          δ��˵�� = Decode(n_��Ŀ���, n_����, Decode(v_��¼����, '����', Null, δ��˵��_In), δ��˵��_In), ��¼�� = v_������, ��¼ʱ�� = Sysdate
      Where ��¼id = n_��¼id And ��Ŀ��� = n_��Ŀ���  And
            Decode(��Ŀ���_In, n_����, Nvl(��¼���, 0), n_��ʹǿ��, Nvl(��¼���, 0), Nvl(n_��¼���, 0)) = Nvl(n_��¼���, 0) And
            Decode(n_��Ŀ����, 2, Nvl(���²�λ, '��'), Nvl(���²�λ_In, '��')) = Nvl(���²�λ_In, '��') And ��ֹ�汾 Is Null;
      Else
        Update ���˻�����ϸ
        Set ��¼���� = v_��¼����, ��¼�� = v_������,��¼���� = ��¼����_In, ��¼ʱ�� = Sysdate
        Where ��¼id = n_��¼id And ��Ŀ��� = n_��Ŀ���  And
              Decode(��Ŀ���_In, n_����, Nvl(��¼���, 0), n_��ʹǿ��, Nvl(��¼���, 0), Nvl(n_��¼���, 0)) = Nvl(n_��¼���, 0) And ��ֹ�汾 Is Null;
      End If;
    end if;
    If Sql%RowCount = 0 Then
      --���뱾�εǼǵĲ��˻�������
      If Mod(��¼����_In, 10) = 1 Or ��¼����_In = 7 Then
        Select ���˻�����ϸ_Id.Nextval Into n_Newid From Dual;
        Insert Into ���˻�����ϸ
          (ID, ��¼id, ��¼����, ��Ŀ����, ��Ŀid, ��Ŀ���, ��Ŀ����, ��Ŀ����, ��¼����, ��Ŀ��λ, ��¼���, ��¼��, ���²�λ, ���Ժϸ�, ��ʼ�汾, ��ֹ�汾, ��¼���, δ��˵��,
           ��¼ʱ��, ������Դ, ��ʾ, ��Դid, ����)
          Select n_Newid, n_��¼id, ��¼����_In, ������, ��Ŀid, ��Ŀ���, ��Ŀ����, ��Ŀ����, v_��¼����, ��Ŀ��λ, n_��¼���, v_������, ���²�λ_In, ���Ժϸ�_In,
                 n_��ʼ�汾, Null, Null, Decode(n_��Ŀ���, n_����, Decode(v_��¼����, '����', Null, δ��˵��_In), δ��˵��_In), Sysdate,
                 ������Դ_In, 0, ��Դid_In, ����_In
          From �����¼��Ŀ
          Where ��Ŀ��� = n_��Ŀ���;
        If ��ʾ_In = 1 Then
          Zl_���µ�����_������ʾ(n_Newid, 1);
        End If;
        If n_Count > 0 And ����_In = 1 Then
          Zl_���µ���������_Update(�ļ�id_In, ����ʱ��_In, ��Ŀ���_In, n_��¼id);
        End If;
      Else
        Insert Into ���˻�����ϸ
          (ID, ��¼id, ��¼����, ��Ŀ����, ��Ŀid, ��Ŀ���, ��Ŀ����, ��Ŀ����, ��¼����, ��Ŀ��λ, ��¼���, ��¼��, ���²�λ, ���Ժϸ�, ��ʼ�汾, ��ֹ�汾, ��¼���, δ��˵��,
           ��¼ʱ��, ������Դ, ��ʾ, ��Դid, ����)
        Values
          (���˻�����ϸ_Id.Nextval, n_��¼id, ��¼����_In, Null, Null, 0,
           Decode(��¼����_In, 2, '�ϱ�˵��', 6, '�±�˵��', 3, '���ת', 4, v_��¼����), Decode(��¼����_In, 3, 0, 1),
           Decode(��¼����_In, 4, '1', ��¼����_In), '', n_��¼���, v_������, ���²�λ_In, ���Ժϸ�_In, n_��ʼ�汾, Null, Null, δ��˵��_In, Sysdate,
           ������Դ_In, 0, ��Դid_In, ����_In);
      End If;
    Else

      If n_Count > 0 And ����_In = 1 Then
        Zl_���µ���������_Update(�ļ�id_In, ����ʱ��_In, ��Ŀ���_In, n_��¼id, 1);
      End If;
    End If;
    If n_Preblue > 0 Then
      v_Data := Substr(v_Data, Instr(v_Data, '/', 1) + 1);
    End If;
  End Loop;
  --���´�ӡ
  Update ���µ���ӡ
  Set ��ӡ�� = Null, ��ӡʱ�� = Null
  Where �ļ�id = �ļ�id_In And
        ��ʼʱ�� = (Select Max(��ʼʱ��) From ���µ���ӡ Where �ļ�id = �ļ�id_In And ��ʼʱ�� <= d_����ʱ��);
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���µ�����_Update;
/

--117925:����,2017-12-08,�ų�������������
Create Or Replace Procedure Zl_��������_Delete(
                                           --ɾ��ҩƷ�շ���¼���ָ���Ӧ�ı�ҩƷ���
                                           No_In In ҩƷ�շ���¼.No%Type) Is
  Err_Item Exception;
  v_Err_Msg  Varchar2(100);
  v_�¿��   Zlparameters.����ֵ%Type;
  n_ʵ������ �շ���ĿĿ¼.�Ƿ���%Type;

Begin
  Select Nvl(zl_GetSysParameter(95), '0') Into v_�¿�� From Dual;

  If v_�¿�� = 1 Then
    --ͨ��ѭ�����ָ�ԭ���Ŀ�������
    For v_���� In (Select ID, ��д����, �ⷿid, ���ۼ�, ����, ����, ҩƷid, ��ҩ��λid, �ɱ���, Ч��, ���Ч��, ����, ��������, ��׼�ĺ�
                 From ҩƷ�շ���¼
                 Where NO = No_In And ���� = 20 And ���ϵ�� = -1
                 Order By ҩƷid, ����) Loop
      Select �Ƿ��� Into n_ʵ������ From �շ���ĿĿ¼ Where ID = v_����.ҩƷid;
    
      Update ҩƷ���
      Set �������� = Nvl(��������, 0) + v_����.��д����
      Where �ⷿid = v_����.�ⷿid And ҩƷid = v_����.ҩƷid And Nvl(����, 0) = Nvl(v_����.����, 0) And ���� = 1;
    
      If Sql%NotFound Then
      
        Insert Into ҩƷ���
          (�ⷿid, ҩƷid, ����, ����, ��������, Ч��, ���Ч��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴ���������, �ϴβ���, ��׼�ĺ�, ���ۼ�)
        Values
          (v_����.�ⷿid, v_����.ҩƷid, v_����.����, 1, v_����.��д����, v_����.Ч��, v_����.���Ч��, v_����.��ҩ��λid, v_����.�ɱ���, v_����.����, v_����.��������,
           v_����.����, v_����.��׼�ĺ�, Decode(n_ʵ������, 1, Decode(Nvl(v_����.����, 0), 0, Null, v_����.���ۼ�), Null));
      End If;
    
      Delete From ҩƷ���
      Where �ⷿid = v_����.�ⷿid And ҩƷid = v_����.ҩƷid And Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0 And
            Nvl(ʵ�ʲ��, 0) = 0;
      Delete From ����������Ϣ Where �շ�id = v_����.Id;
    End Loop;
  End If;

  Delete From ҩƷ�շ���¼ Where NO = No_In And ���� = 20 And ��¼״̬ = 1 And ����� Is Null;

  If Sql%RowCount = 0 Then
    v_Err_Msg := '[ZLSOFT]�õ����Ѿ�������ɾ�����ѱ�����ˣ�[ZLSOFT]';
    Raise Err_Item;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��������_Delete;
/

--117925:����,2017-12-08,��������������
Create Or Replace Procedure Zl_�����������_Verify
(
  No_In     In ҩƷ�շ���¼.No%Type := Null,
  �����_In In ҩƷ�շ���¼.�����%Type := Null
) Is
  Err_Item Exception;
  v_Err_Msg Varchar2(500);

  n_ʵ������    �շ���ĿĿ¼.�Ƿ���%Type;
  n_Batch_Count Integer; --ԭ���������ڷ����Ĳ��ϵ�����
  v_����ǰ׺    Varchar2(20);
  v_�ڲ�����    ҩƷ���.�ڲ�����%Type;
  n_ƽ���ɱ���  ҩƷ���.ƽ���ɱ���%Type;
Begin
  v_����ǰ׺ := Nvl(Zl_Getsysparameter(159), '');

  Update ҩƷ�շ���¼
  Set ����� = �����_In, ������� = Sysdate
  Where NO = No_In And ���� = 17 And ��¼״̬ = 1 And ����� Is Null;

  If Sql%RowCount = 0 Then
    v_Err_Msg := '[ZLSOFT]�õ����Ѿ���������ˣ�[ZLSOFT]';
    Raise Err_Item;
  End If;

  --��Ҫ���ԭ���������ڷ����Ĳ��ϣ����ܶ������
  Select Count(*)
  Into n_Batch_Count
  From ҩƷ�շ���¼ A, �������� B
  Where a.ҩƷid = b.����id And a.No = No_In And a.���� = 17 And a.��¼״̬ = 1 And Nvl(a.����, 0) = 0 And
        ((Nvl(b.�ⷿ����, 0) = 1 And
        a.�ⷿid Not In (Select ����id From ��������˵�� Where (�������� Like '���ϲ���') Or (�������� Like '�Ƽ���'))) Or Nvl(b.���÷���, 0) = 1);

  If n_Batch_Count > 0 Then
    v_Err_Msg := '[ZLSOFT]�õ����а�����һ��ԭ�������������ڷ����Ĳ��ϣ�������ˣ�[ZLSOFT';
    Raise Err_Item;
  End If;

  --ԭ�����ֲ������Ĳ���,�����ʱ��Ҫ������
  Update ҩƷ�շ���¼
  Set ���� = 0
  Where ID =
        (Select ID
         From ҩƷ�շ���¼ A, �������� B
         Where b.����id = a.ҩƷid And a.No = No_In And a.���� = 17 And a.��¼״̬ = 1 And Nvl(a.����, 0) > 0 And
               (Nvl(b.�ⷿ����, 0) = 0 Or
               (Nvl(b.���÷���, 0) = 0 And
               a.�ⷿid In (Select ����id From ��������˵�� Where (�������� Like '���ϲ���') Or (�������� Like '�Ƽ���')))));

  For c_�շ� In (Select a.Id, a.ʵ������, a.���ۼ�, a.���۽��, a.���, a.�ⷿid, a.ҩƷid, a.����, a.�ɱ���, a.����, a.Ч��, a.���Ч��, a.�������, a.����,
                      a.������id, a.��������, a.��Ʒ����, a.�ڲ�����, Nvl(b.�Ƿ��������, 0) As �������, a.��׼�ĺ�
               From ҩƷ�շ���¼ A, �������� B
               Where a.ҩƷid = b.����id And a.No = No_In And a.���� = 17 And a.��¼״̬ = 1
               Order By a.ҩƷid, a.����) Loop
  
    v_�ڲ����� := Null;
    If c_�շ�.������� = 1 Then
      If Not v_����ǰ׺ Is Null Then
        v_�ڲ����� := v_����ǰ׺ || Nextno(126);
      Else
        v_�ڲ����� := Nextno(126);
      End If;
    End If;
  
    --����ҩƷ�������Ӧ����
    Select �Ƿ��� Into n_ʵ������ From �շ���ĿĿ¼ Where ID = c_�շ�.ҩƷid;
  
    Update ҩƷ���
    Set �������� = Nvl(��������, 0) + Nvl(c_�շ�.ʵ������, 0), ʵ������ = Nvl(ʵ������, 0) + Nvl(c_�շ�.ʵ������, 0),
        ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) + Nvl(c_�շ�.���۽��, 0), ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + Nvl(c_�շ�.���, 0), �ϴβɹ��� = Nvl(c_�շ�.�ɱ���, �ϴβɹ���),
        �ϴ����� = Nvl(c_�շ�.����, �ϴ�����), �ϴ��������� = Nvl(c_�շ�.��������, �ϴ���������), �ϴβ��� = Nvl(c_�շ�.����, �ϴβ���), Ч�� = Nvl(c_�շ�.Ч��, Ч��),
        ���Ч�� = Nvl(c_�շ�.���Ч��, ���Ч��), ���ۼ� = Decode(Nvl(c_�շ�.����, 0), 0, Null, Decode(n_ʵ������, 1, c_�շ�.���ۼ�, Null)),
        ��Ʒ���� = c_�շ�.��Ʒ����, �ڲ����� = v_�ڲ�����, ��׼�ĺ� = c_�շ�.��׼�ĺ�
    Where �ⷿid = c_�շ�.�ⷿid And ҩƷid = c_�շ�.ҩƷid And Nvl(����, 0) = Nvl(c_�շ�.����, 0) And ���� = 1;
  
    If Sql%NotFound Then
      Insert Into ҩƷ���
        (�ⷿid, ҩƷid, ����, ����, ��������, ʵ������, ʵ�ʽ��, ʵ�ʲ��, �ϴβɹ���, �ϴ�����, �ϴ���������, �ϴβ���, Ч��, ���Ч��, ���ۼ�, ��Ʒ����, �ڲ�����, ƽ���ɱ���, ��׼�ĺ�)
      Values
        (c_�շ�.�ⷿid, c_�շ�.ҩƷid, c_�շ�.����, 1, c_�շ�.ʵ������, c_�շ�.ʵ������, c_�շ�.���۽��, c_�շ�.���, c_�շ�.�ɱ���, c_�շ�.����, c_�շ�.��������,
         c_�շ�.����, c_�շ�.Ч��, c_�շ�.���Ч��, Decode(Nvl(c_�շ�.����, 0), 0, Null, Decode(n_ʵ������, 1, c_�շ�.���ۼ�, Null)), c_�շ�.��Ʒ����,
         v_�ڲ�����, c_�շ�.�ɱ���, c_�շ�.��׼�ĺ�);
    End If;
  
    Delete From ҩƷ���
    Where �ⷿid = c_�շ�.�ⷿid And ҩƷid = c_�շ�.ҩƷid And Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0 And
          Nvl(ʵ�ʲ��, 0) = 0;
  
    If Not v_�ڲ����� Is Null Then
      Update ҩƷ�շ���¼ Set �ڲ����� = v_�ڲ����� Where ID = c_�շ�.Id;
    End If;
  
    --���¸ò��ϵĳɱ���
    Update �������� Set �ɱ��� = c_�շ�.�ɱ���, �ϴ��ۼ� = c_�շ�.���ۼ� Where ����id = c_�շ�.ҩƷid;
  
    --���¼�������е�ƽ���ɱ���
    Update ҩƷ���
    Set ƽ���ɱ��� = Decode(Nvl(����, 0), 0, Decode((ʵ�ʽ�� - ʵ�ʲ��) / ʵ������, 0, �ϴβɹ���, (ʵ�ʽ�� - ʵ�ʲ��) / ʵ������), �ϴβɹ���)
    Where ҩƷid = c_�շ�.ҩƷid And Nvl(����, 0) = Nvl(c_�շ�.����, 0) And �ⷿid = c_�շ�.�ⷿid And Nvl(ʵ������, 0) <> 0 And ���� = 1;
    If Sql%NotFound Then
      Select �ɱ��� Into n_ƽ���ɱ��� From �������� Where ����id = c_�շ�.ҩƷid;
      Update ҩƷ���
      Set ƽ���ɱ��� = n_ƽ���ɱ���
      Where ҩƷid = c_�շ�.ҩƷid And �ⷿid = c_�շ�.�ⷿid And Nvl(����, 0) = Nvl(c_�շ�.����, 0) And ���� = 1;
    End If;
  
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����������_Verify;
/

--117925:����,2017-12-07,��������������
Create Or Replace Procedure Zl_��������_Delete(No_In In ҩƷ�շ���¼.NO%Type) Is
  Err_Item Exception;
  v_Err_Msg  Varchar2(500);
  v_�¿��   zlParameters.����ֵ%Type;
  v_��ȷ���� zlParameters.����ֵ%Type;
  d_�������� ҩƷ�շ���¼.��ҩ����%Type;
  n_ʵ������ �շ���ĿĿ¼.�Ƿ���%Type;

Begin
  Select Nvl(zl_GetSysParameter(95), '0') Into v_�¿�� From Dual;

  --ֻ������ȷ���ε�����²����¿��ÿ��
  Select Nvl(zl_GetSysParameter(83), '0') Into v_��ȷ���� From Dual;

  --����Ƿ��Ѿ����ϻ����
  Select ��ҩ���� Into d_�������� From ҩƷ�շ���¼ Where ���� = 19 And NO = No_In And Rownum < 2;

  If d_�������� Is Not Null Then
    v_Err_Msg := '[ZLSOFT]�����쵥�Ѿ������˷���,���ܽ���ɾ��![ZLSOFT]';
    Raise Err_Item;
  End If;

  If To_Number(v_�¿��, '9999') = 1 And To_Number(v_��ȷ����, '9999') = 1 Then
    --��Ҫ��ԭ���ÿ��
    --ͨ��ѭ�����ָ�ԭ���Ŀ�������
    For v_���� In (Select ʵ������, �ⷿid, ����, ҩƷid, ���ۼ�, ����, Ч��, ����, ��ҩ��λid, �ɱ���, ��������,
                          ��׼�ĺ�, ���Ч��
                   From ҩƷ�շ���¼
                   Where NO = No_In And ���� = 19 And ���ϵ�� = -1
                   Order By ҩƷid, ����) Loop
      Select �Ƿ��� Into n_ʵ������ From �շ���ĿĿ¼ Where ID = v_����.ҩƷid;

      Update ҩƷ���
      Set �������� = Nvl(��������, 0) + v_����.ʵ������
      Where �ⷿid = v_����.�ⷿid And ҩƷid = v_����.ҩƷid And Nvl(����, 0) = Nvl(v_����.����, 0) And ���� = 1;

      If Sql%NotFound Then

        Insert Into ҩƷ���
          (�ⷿid, ҩƷid, ����, ����, ��������, Ч��, ���Ч��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴ���������,
           �ϴβ���, ��׼�ĺ�, ���ۼ�)
        Values
          (v_����.�ⷿid, v_����.ҩƷid, v_����.����, 1, v_����.ʵ������, v_����.Ч��, v_����.���Ч��,
           v_����.��ҩ��λid, v_����.�ɱ���, v_����.����, v_����.��������, v_����.����, v_����.��׼�ĺ�,
           Decode(n_ʵ������, 1, Decode(Nvl(v_����.����, 0), 0, Null, v_����.���ۼ�), Null));
      End If;

      Delete From ҩƷ���
      Where �ⷿid = v_����.�ⷿid And ҩƷid = v_����.ҩƷid And Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And
            Nvl(ʵ�ʽ��, 0) = 0 And Nvl(ʵ�ʲ��, 0) = 0;
    End Loop;
  End If;

  Delete ҩƷ�շ���¼ Where NO = No_In And ���� = 19 And ��¼״̬ = 1 And ����� Is Null;

  If Sql%NotFound Then
    v_Err_Msg := '[ZLSOFT]�õ����Ѿ�������ɾ�����ѱ�����ˣ�[ZLSOFT]';
    Raise Err_Item;
  End If;
Exception

  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��������_Delete;
/

--117925:����,2017-12-07,��������������
Create Or Replace Procedure Zl_�����̵�_Strike
(
  No_In     In ҩƷ�շ���¼.No%Type,
  �����_In In ҩƷ�շ���¼.�����%Type
) Is
  Err_Item Exception;
  v_Err_Msg Varchar2(100);

  n_Batch_Count Integer; --ԭ���������ڷ����Ĳ��ϵ�����
  n_Count       Integer; --ԭ�����ֲ�����

  n_����       ҩƷ�շ���¼.����%Type;
  n_�ɱ���     ҩƷ�շ���¼.�ɱ���%Type;
  n_����id     ҩƷ�շ���¼.ҩƷid%Type;
  n_ʵ������   �շ���ĿĿ¼.�Ƿ���%Type;
  n_ƽ���ɱ��� ҩƷ���.ƽ���ɱ���%Type;

Begin
  Update ҩƷ�շ���¼ Set ��¼״̬ = 3 Where NO = No_In And ���� = 22 And ��¼״̬ = 1;

  If Sql%RowCount = 0 Then
    v_Err_Msg := '[ZLSOFT]�õ����Ѿ������˳�����[ZLSOFT]';
    Raise Err_Item;
  End If;

  --��Ҫ���ԭ���������ڷ����Ĳ��ϣ����ܶ������
  Select Count(*), Max(a.ҩƷid)
  Into n_Batch_Count, n_����id
  From ҩƷ�շ���¼ A, �������� B
  Where a.ҩƷid = b.����id And a.No = No_In And a.���� = 22 And a.��¼״̬ = 3 And Nvl(a.����, 0) = 0 And
        ((Nvl(b.�ⷿ����, 0) = 1 And
        a.�ⷿid Not In (Select ����id From ��������˵�� Where (�������� Like '���ϲ���') Or (�������� Like '�Ƽ���'))) Or Nvl(b.���÷���, 0) = 1);

  If n_Batch_Count > 0 Then
    Begin
      Select ���� || '-' || ���� Into v_Err_Msg From �շ���ĿĿ¼ Where ID = n_����id;
    Exception
      When Others Then
        Null;
    End;
    v_Err_Msg := '[ZLSOFT]�õ����в���Ϊ:' || v_Err_Msg || Chr(10) || Chr(13) || '�Ĳ���,ԭ��������,�����ڷ�������˲�����ˣ�[ZLSOFT]';
    Raise Err_Item;
  End If;

  Insert Into ҩƷ�շ���¼
    (ID, ��¼״̬, ����, NO, ���, �ⷿid, ������id, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ���Ч��, ��д����, ����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ���۽��, ���, ժҪ,
     ������, ��������, �����, �������, Ƶ��, ��ҩ��λid, ��������, ��׼�ĺ�, ����)
    Select ҩƷ�շ���¼_Id.Nextval, 2, ����, NO, ���, �ⷿid, ������id, ���ϵ��, a.ҩƷid,
           Decode(Nvl(a.����, 0), 0, Null, (Decode(Nvl(b.�ⷿ����, 0), 0, Null, a.����))), a.����, ����, a.Ч��, a.���Ч��, ��д����, a.����,
           -ʵ������, a.�ɱ���, �ɱ����, ���ۼ�, -���۽��, -���, ժҪ, �����_In, Sysdate, �����_In, Sysdate, a.Ƶ��, a.��ҩ��λid, a.��������, a.��׼�ĺ�,
           a.����
    From ҩƷ�շ���¼ A, �������� B
    Where NO = No_In And a.ҩƷid = b.����id And ���� = 22 And ��¼״̬ = 3;

  For c_���� In (Select ID, ʵ������, ���ۼ�, ���۽��, ���, �ⷿid, ҩƷid ����id, ����, ����, Ч��, ���Ч��, ����, ������id, ���ϵ��, ��ҩ��λid, ��������, ��׼�ĺ�,
                      ����
               From ҩƷ�շ���¼
               Where NO = No_In And ���� = 22 And ��¼״̬ = 2
               Order By ҩƷid, ����) Loop
    Select �Ƿ��� Into n_ʵ������ From �շ���ĿĿ¼ Where ID = c_����.����id;
  
    --ԭ�����ֲ������Ĳ���,��C����ʱ��Ҫ������
    Begin
      Select Count(*)
      Into n_Count
      From ҩƷ�շ���¼ A, �������� B
      Where b.����id + 0 = c_����.����id And a.No = No_In And a.ҩƷid = b.����id And a.���� = 22 And a.�ⷿid + 0 = c_����.�ⷿid And
            a.��¼״̬ = 3 And Nvl(a.����, 0) > 0 And
            (Nvl(b.�ⷿ����, 0) = 0 Or
            (Nvl(b.���÷���, 0) = 0 And
            a.�ⷿid In (Select ����id From ��������˵�� Where (�������� Like '���ϲ���') Or (�������� Like '�Ƽ���'))));
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count > 0 Then
      n_���� := 0;
    Else
      n_���� := Nvl(c_����.����, 0);
    End If;
  
    --����ҩƷ�������Ӧ����
    Update ҩƷ���
    Set �������� = Nvl(��������, 0) + Nvl(c_����.ʵ������, 0) * c_����.���ϵ��, ʵ������ = Nvl(ʵ������, 0) + Nvl(c_����.ʵ������, 0) * c_����.���ϵ��,
        ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) + Nvl(c_����.���۽��, 0) * c_����.���ϵ��, ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + Nvl(c_����.���, 0) * c_����.���ϵ��,
        ���ۼ� = Decode(n_ʵ������, 1, Decode(Nvl(n_����, 0), 0, Null, Decode(Nvl(���ۼ�, 0), 0, c_����.���ۼ�, ���ۼ�)), Null)
    Where �ⷿid = c_����.�ⷿid And ҩƷid = c_����.����id And Nvl(����, 0) = n_���� And ���� = 1;
  
    If Sql%NotFound Then
      If Nvl(c_����.ʵ������, 0) <> 0 Then
        n_�ɱ��� := Round((Nvl(c_����.���۽��, 0) - Nvl(c_����.���, 0)) / c_����.ʵ������, 7);
      Else
        n_�ɱ��� := 0;
      End If;
      Insert Into ҩƷ���
        (�ⷿid, ҩƷid, ����, ����, ��������, ʵ������, ʵ�ʽ��, ʵ�ʲ��, Ч��, ���Ч��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴ���������, �ϴβ���, ��׼�ĺ�, ���ۼ�, ƽ���ɱ���)
      
      Values
        (c_����.�ⷿid, c_����.����id, n_����, 1, c_����.ʵ������ * c_����.���ϵ��, c_����.ʵ������ * c_����.���ϵ��, c_����.���۽�� * c_����.���ϵ��,
         c_����.��� * c_����.���ϵ��, c_����.Ч��, c_����.���Ч��, c_����.��ҩ��λid, n_�ɱ���, c_����.����, c_����.��������, c_����.����, c_����.��׼�ĺ�,
         Decode(n_ʵ������, 1, Decode(Nvl(c_����.����, 0), 0, Null, c_����.���ۼ�), Null), n_�ɱ���);
    End If;
  
    Delete From ҩƷ���
    Where �ⷿid = c_����.�ⷿid And ҩƷid = c_����.����id And Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0 And
          Nvl(ʵ�ʲ��, 0) = 0;
    Zl_�����շ���¼_��������(c_����.Id);
  
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����̵�_Strike;
/

--117925:����,2017-12-07,��������������
Create Or Replace Procedure Zl_�����̵�_Verify
(
  No_In     In ҩƷ�շ���¼.No%Type := Null,
  �����_In In ҩƷ�շ���¼.�����%Type := Null
) Is
  Err_Item Exception;
  v_Err_Msg Varchar2(200);

  n_�ɱ���     ҩƷ�շ���¼.�ɱ���%Type;
  n_����id     ҩƷ�շ���¼.ҩƷid%Type;
  n_ʵ������   �շ���ĿĿ¼.�Ƿ���%Type;
  n_ƽ���ɱ��� ҩƷ���.ƽ���ɱ���%Type;

  n_Batch_Count Integer; --ԭ���������ڷ����Ĳ��ϵ�����

Begin
  Update ҩƷ�շ���¼
  Set ����� = �����_In, ������� = Sysdate
  Where NO = No_In And ���� = 22 And ��¼״̬ = 1 And ����� Is Null;

  If Sql%RowCount = 0 Then
    v_Err_Msg := '[ZLSOFT]�õ����Ѿ���������ˣ�[ZLSOFT]';
    Raise Err_Item;
  End If;

  For c_���� In (Select ID, ʵ������, ���ۼ�, ���۽��, ���, �ⷿid, ҩƷid ����id, ����, ����, Ч��, ���Ч��, ����, ������id, ���ϵ��, ��ҩ��λid, ��������, ��׼�ĺ�
               From ҩƷ�շ���¼
               Where NO = No_In And ���� = 22 And ��¼״̬ = 1
               Order By ����id, ����) Loop
  
    Select �Ƿ��� Into n_ʵ������ From �շ���ĿĿ¼ Where ID = c_����.����id;
  
    --����ҩƷ�������Ӧ����
    Update ҩƷ���
    Set �������� = Nvl(��������, 0) + Decode(c_����.���ϵ��, 1, Nvl(c_����.ʵ������, 0), 0),
        ʵ������ = Nvl(ʵ������, 0) + Nvl(c_����.ʵ������, 0) * c_����.���ϵ��, ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) + Nvl(c_����.���۽��, 0) * c_����.���ϵ��,
        ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + Nvl(c_����.���, 0) * c_����.���ϵ��,
        ���ۼ� = Decode(n_ʵ������, 1, Decode(Nvl(c_����.����, 0), 0, Null, Decode(Nvl(���ۼ�, 0), 0, c_����.���ۼ�, ���ۼ�)), Null)
    Where �ⷿid = c_����.�ⷿid And ҩƷid = c_����.����id And Nvl(����, 0) = Nvl(c_����.����, 0) And ���� = 1;
  
    If Sql%NotFound Then
      If Nvl(c_����.ʵ������, 0) <> 0 Then
        n_�ɱ��� := Round((Nvl(c_����.���۽��, 0) - Nvl(c_����.���, 0)) / c_����.ʵ������, 7);
      Else
        n_�ɱ��� := 0;
      End If;
    
      Insert Into ҩƷ���
        (�ⷿid, ҩƷid, ����, ����, ��������, ʵ������, ʵ�ʽ��, ʵ�ʲ��, Ч��, ���Ч��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴ���������, �ϴβ���, ��׼�ĺ�, ���ۼ�,ƽ���ɱ���)
      
      Values
        (c_����.�ⷿid, c_����.����id, c_����.����, 1, Decode(c_����.���ϵ��, 1, Nvl(c_����.ʵ������, 0), 0), c_����.ʵ������ * c_����.���ϵ��,
         c_����.���۽�� * c_����.���ϵ��, c_����.��� * c_����.���ϵ��, c_����.Ч��, c_����.���Ч��, c_����.��ҩ��λid, n_�ɱ���, c_����.����, c_����.��������, c_����.����,
         c_����.��׼�ĺ�, Decode(n_ʵ������, 1, Decode(Nvl(c_����.����, 0), 0, Null, c_����.���ۼ�), Null),n_�ɱ���);
    End If;
  
    Delete From ҩƷ���
    Where �ⷿid = c_����.�ⷿid And ҩƷid = c_����.����id And Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0 And
          Nvl(ʵ�ʲ��, 0) = 0;
  
    --���¼���ƽ���ɱ���
    Update ҩƷ���
    Set ƽ���ɱ��� = Decode(Nvl(����, 0), 0, decode((ʵ�ʽ�� - ʵ�ʲ��) / ʵ������,0,�ϴβɹ���,(ʵ�ʽ�� - ʵ�ʲ��) / ʵ������), �ϴβɹ���)
    Where �ⷿid = c_����.�ⷿid And ҩƷid = c_����.����id And Nvl(����, 0) = Nvl(c_����.����, 0) And ���� = 1 And Nvl(ʵ������, 0) <> 0;
    If Sql%NotFound Then
      Select �ɱ��� Into n_ƽ���ɱ��� From �������� Where ����id = c_����.����id;
      Update ҩƷ���
      Set ƽ���ɱ��� = n_ƽ���ɱ���
      Where ҩƷid = c_����.����id And �ⷿid = c_����.�ⷿid And Nvl(����, 0) = Nvl(c_����.����, 0) and ����=1;
    End If;
  
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����̵�_Verify;
/

--117925:����,2017-12-07,��������������
Create Or Replace Procedure Zl_�����̵�_Delete(

                                               --ɾ��ҩƷ�շ���¼���ָ���Ӧ�ı�ҩƷ���
                                               No_In In ҩƷ�շ���¼.NO%Type) Is
  Err_Item Exception;
  v_Err_Msg  Varchar2(100);
  n_�ɱ���   ҩƷ�շ���¼.�ɱ���%Type;
  n_ʵ������ �շ���ĿĿ¼.�Ƿ���%Type;

Begin
  --ͨ��ѭ�����ָ��������ԭ���Ŀ���������
  --ʵ�������������������
  For c_���� In (Select ʵ������, �ⷿid, ����, ҩƷid ����id, ���ۼ�, ��ҩ��λid, ���۽��, ���, Ч��, ���Ч��, ����,
                        ����, ��������, ��׼�ĺ�
                 From ҩƷ�շ���¼
                 Where NO = No_In And ���� = 22 And ���ϵ�� = -1
                 Order By ҩƷid, ����) Loop
    Select �Ƿ��� Into n_ʵ������ From �շ���ĿĿ¼ Where ID = c_����.����id;

    Update ҩƷ���
    Set �������� = Nvl(��������, 0) + c_����.ʵ������,
        ���ۼ� = Decode(n_ʵ������, 1,
                         Decode(Nvl(c_����.����, 0), 0, Null, Decode(Nvl(���ۼ�, 0), 0, c_����.���ۼ�, ���ۼ�)), Null)
    Where �ⷿid = c_����.�ⷿid And ҩƷid = c_����.����id And Nvl(����, 0) = Nvl(c_����.����, 0) And ���� = 1;

    If Sql%NotFound Then
      If Nvl(c_����.ʵ������, 0) <> 0 Then
        n_�ɱ��� := Round((Nvl(c_����.���۽��, 0) - Nvl(c_����.���, 0)) / c_����.ʵ������, 7);
      Else
        n_�ɱ��� := 0;
      End If;
      Insert Into ҩƷ���
        (�ⷿid, ҩƷid, ����, ����, ��������, �ϴι�Ӧ��id, �ϴβɹ���, ���ۼ�)
      Values
        (c_����.�ⷿid, c_����.����id, c_����.����, 1, c_����.ʵ������, c_����.��ҩ��λid, n_�ɱ���,
         Decode(n_ʵ������, 1, Decode(Nvl(c_����.����, 0), 0, Null, c_����.���ۼ�), Null));
    End If;

    Delete From ҩƷ���
    Where �ⷿid = c_����.�ⷿid And ҩƷid = c_����.����id And Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And
          Nvl(ʵ�ʽ��, 0) = 0 And Nvl(ʵ�ʲ��, 0) = 0;
  End Loop;

  Delete From ҩƷ�շ���¼ Where NO = No_In And ���� = 22 And ��¼״̬ = 1 And ����� Is Null;

  If Sql%RowCount = 0 Then
    v_Err_Msg := '[ZLSOFT]�õ����Ѿ�������ɾ�����ѱ�����ˣ�[ZLSOFT]';
    Raise Err_Item;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����̵�_Delete;
/

--117925:����,2017-12-07,��������������
Create Or Replace Procedure Zl_�����ƿ�_Back(No_In In ҩƷ�շ���¼.NO%Type) Is
  v_Err_Msg Varchar2(500);
  Err_Item Exception;

  d_�������� ҩƷ�շ���¼.��ҩ����%Type;
  v_����     ҩƷ�շ���¼.��ҩ��%Type;
  v_���     ҩƷ�շ���¼.�����%Type;
  v_�¿��   zlParameters.����ֵ%Type;
  n_ʵ������ �շ���ĿĿ¼.�Ƿ���%Type;

Begin
  Select Nvl(zl_GetSysParameter(95), '0') Into v_�¿�� From Dual;

  Select ��ҩ��, ��ҩ����, �����
  Into v_����, d_��������, v_���
  From ҩƷ�շ���¼
  Where ���� = 19 And NO = No_In And Rownum < 2;

  If v_��� Is Not Null Then
    v_Err_Msg := '[ZLSOFT]�õ����ѱ��ⷿ���գ�����������ˣ�[ZLSOFT]';
    Raise Err_Item;
  End If;

  If v_���� Is Null Then
    Return;
  End If;

  If d_�������� Is Null Then
    --��������ҩ��Ϊ�ռ���
    Update ҩƷ�շ���¼ Set ��ҩ�� = Null, ��� = Null Where ���� = 19 And NO = No_In;
  Else

    --��Ҫ�ָ�����ⷿ�Ŀ�������
    Update ҩƷ�շ���¼ Set ��ҩ���� = Null Where ���� = 19 And NO = No_In;
    --����������ӵ���ʱ�Ѿ����˿���,�򱾴λ��˲��ٻ�㹿��ÿɴ���.
    If To_Number(v_�¿��, '9999') <> 1 Then

      For v_���� In (Select ʵ������, �ⷿid, ���ۼ�, ����, ҩƷid, ����, Ч��, ����, ��ҩ��λid, �ɱ���, ��������,
                            ���Ч��, ��׼�ĺ�
                     From ҩƷ�շ���¼
                     Where NO = No_In And ���� = 19 And ���ϵ�� = -1
                     Order By ҩƷid, ����) Loop
        Select �Ƿ��� Into n_ʵ������ From �շ���ĿĿ¼ Where ID = v_����.ҩƷid;

        Update ҩƷ���
        Set �������� = Nvl(��������, 0) + v_����.ʵ������
        Where �ⷿid = v_����.�ⷿid And ҩƷid = v_����.ҩƷid And Nvl(����, 0) = Nvl(v_����.����, 0) And ���� = 1;

        If Sql%RowCount = 0 Then
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, ��������, Ч��, ���Ч��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴ���������,
             �ϴβ���, ��׼�ĺ�, ���ۼ�)
          Values
            (v_����.�ⷿid, v_����.ҩƷid, Nvl(v_����.����, 0), 1, v_����.ʵ������, v_����.Ч��, v_����.���Ч��,
             v_����.��ҩ��λid, v_����.�ɱ���, v_����.����, v_����.��������, v_����.����, v_����.��׼�ĺ�,
             Decode(n_ʵ������, 1, Decode(Nvl(v_����.����, 0), 0, Null, v_����.���ۼ�), Null));

        End If;

        Delete From ҩƷ���
        Where �ⷿid = v_����.�ⷿid And ҩƷid = v_����.ҩƷid And Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And
              Nvl(ʵ�ʽ��, 0) = 0 And Nvl(ʵ�ʲ��, 0) = 0;
      End Loop;
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����ƿ�_Back;
/

--117925:����,2017-12-07,��������������
Create Or Replace Procedure Zl_�����ƿ�_Prepare
(
  No_In     In ҩƷ�շ���¼.NO%Type,
  ����Ա_In Varchar2 := Null
) Is
  v_Err_Msg Varchar2(500);
  Err_Item Exception;
  v_�¿��   zlParameters.����ֵ%Type;
  n_ʵ������ �շ���ĿĿ¼.�Ƿ���%Type;

Begin
  Select Nvl(zl_GetSysParameter(95), '0') Into v_�¿�� From Dual;

  If ����Ա_In Is Not Null Then
    Update ҩƷ�շ���¼
    Set ��ҩ�� = ����Ա_In, ��� = To_Char(Sysdate, 'yyyy-MM-dd hh24:mi:ss')
    Where ���� = 19 And NO = No_In;

  Else

    Update ҩƷ�շ���¼ Set ��ҩ���� = Sysdate Where ���� = 19 And NO = No_In;

    If To_Number(v_�¿��, '9999') <> 1 Then
      For v_���� In (Select ʵ������, �ⷿid, ���ۼ�, ����, ҩƷid, ����, Ч��, ����, ��ҩ��λid, �ɱ���, ���Ч��,
                            ��������, ��׼�ĺ�
                     From ҩƷ�շ���¼
                     Where NO = No_In And ���� = 19 And ���ϵ�� = -1
                     Order By ҩƷid, ����) Loop

        Select �Ƿ��� Into n_ʵ������ From �շ���ĿĿ¼ Where ID = v_����.ҩƷid;

        Update ҩƷ���
        Set �������� = Nvl(��������, 0) - v_����.ʵ������
        Where �ⷿid = v_����.�ⷿid And ҩƷid = v_����.ҩƷid And Nvl(����, 0) = Nvl(v_����.����, 0) And ���� = 1;

        If Sql%RowCount = 0 Then
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, ��������, Ч��, ���Ч��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴ���������,
             �ϴβ���, ��׼�ĺ�, ���ۼ�)
          Values
            (v_����.�ⷿid, v_����.ҩƷid, Nvl(v_����.����, 0), 1, -1 * v_����.ʵ������, v_����.Ч��, v_����.���Ч��,
             v_����.��ҩ��λid, v_����.�ɱ���, v_����.����, v_����.��������, v_����.����, v_����.��׼�ĺ�,
             Decode(n_ʵ������, 1, Decode(Nvl(v_����.����, 0), 0, Null, v_����.���ۼ�), Null));

        End If;

        Delete From ҩƷ���
        Where �ⷿid = v_����.�ⷿid And ҩƷid = v_����.ҩƷid And Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And
              Nvl(ʵ�ʽ��, 0) = 0 And Nvl(ʵ�ʲ��, 0) = 0;

      End Loop;
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����ƿ�_Prepare;
/

--117925:����,2017-12-07,��������������
Create Or Replace Procedure Zl_�����ƿ�_Strike
(
  �д�_In       In Integer,
  ԭ��¼״̬_In In ҩƷ�շ���¼.��¼״̬%Type,
  No_In         In ҩƷ�շ���¼.No%Type,
  ���_In       In ҩƷ�շ���¼.���%Type,
  ����id_In     In ҩƷ�շ���¼.ҩƷid%Type,
  ��������_In   In ҩƷ�շ���¼.ʵ������%Type,
  ������_In     In ҩƷ�շ���¼.������%Type,
  ��������_In   In ҩƷ�շ���¼.��������%Type,
  ������ʽ_In   In Integer := 0
  --0������������ʽ��1�������������뵥�ݣ�2������Ѳ����ĳ������뵥��
) Is
  Err_Item Exception;
  v_Err_Msg Varchar2(500);

  n_Batch_Count Integer; --ԭ���������ڷ�����ҩƷ������

  n_�ⷿid       ҩƷ�շ���¼.�ⷿid%Type;
  n_�Է�����id   ҩƷ�շ���¼.�Է�����id%Type;
  n_����         ҩƷ�շ���¼.����%Type;
  n_�ɱ���       ҩƷ�շ���¼.�ɱ���%Type;
  n_�ɱ����     ҩƷ�շ���¼.�ɱ����%Type;
  n_���ۼ�       ҩƷ�շ���¼.���ۼ�%Type;
  n_���۽��     ҩƷ�շ���¼.���۽��%Type;
  n_�����       ҩƷ�շ���¼.���%Type;
  v_����         ҩƷ�շ���¼.����%Type;
  v_����         ҩƷ�շ���¼.����%Type;
  v_Ч��         ҩƷ�շ���¼.Ч��%Type;
  v_��Ʒ����     ҩƷ�շ���¼.��Ʒ����%Type;
  v_�ڲ�����     ҩƷ�շ���¼.�ڲ�����%Type;
  v_�������     ҩƷ�շ���¼.�������%Type;
  v_���Ч��     ҩƷ�շ���¼.���Ч��%Type;
  n_��ҩ��λid   ҩƷ�շ���¼.��ҩ��λid%Type;
  d_��������     ҩƷ�շ���¼.��������%Type;
  v_��׼�ĺ�     ҩƷ�շ���¼.��׼�ĺ�%Type;
  n_����         ҩƷ�շ���¼.����%Type;
  n_���         ҩƷ�շ���¼.���%Type;
  n_���ϵ��     ҩƷ�շ���¼.���ϵ��%Type;
  n_������id   ҩƷ�շ���¼.������id%Type;
  v_��ҩ��       ҩƷ�շ���¼.��ҩ��%Type;
  d_��������     ҩƷ�շ���¼.��ҩ����%Type;
  v_ժҪ         ҩƷ�շ���¼.ժҪ%Type;
  n_ʣ������     ҩƷ�շ���¼.ʵ������%Type;
  n_ʣ��ɱ���� ҩƷ�շ���¼.�ɱ����%Type;
  n_ʣ�����۽�� ҩƷ�շ���¼.���۽��%Type;
  n_�շ�id       ҩƷ�շ���¼.Id%Type;
  n_ʵ������     �շ���ĿĿ¼.�Ƿ���%Type;
  --�Գ����������м��
  n_�����     ҩƷ���.ʵ������%Type;
  n_�ⷿ����   Integer;
  n_���÷���   Integer;
  n_С��       Number;
  n_��¼��     Number;
  n_ƽ���ɱ��� ҩƷ���.ƽ���ɱ���%Type;
  v_�¿��ÿ�� Zlparameters.����ֵ%Type;
  n_��������   ҩƷ���.��������%Type;

  Cursor c_ҩƷ�շ���¼ Is
    Select a.Id, a.���, a.�ⷿid, a.�Է�����id, a.������id, a.���ϵ��, a.ҩƷid, Nvl(a.����, 0) As ����, a.����, a.����, a.Ч��, a.��ҩ��,
           a.��ҩ���� As ��������, a.ժҪ, a.��ҩ��λid, a.��׼�ĺ�, a.��������, a.�ɱ���, a.���ۼ�, Nvl(b.�Ƿ���, 0) As ʱ��, a.����, a.����, a.Ƶ��, a.��Ʒ����,
           a.�ڲ�����, a.�������, a.���Ч��
    From ҩƷ�շ���¼ A, �շ���ĿĿ¼ B
    Where a.ҩƷid = b.Id And a.No = No_In And a.���� = 19 And (a.��� >= ���_In And a.��� <= ���_In + 1) And
          (a.��¼״̬ = 1 Or Mod(a.��¼״̬, 3) = 0)
    Order By a.ҩƷid,a.����,a.���;

  Cursor c_���������¼ Is
    Select a.Id, a.���, a.�ⷿid, a.�Է�����id, a.������id, a.���ϵ��, a.ҩƷid, Nvl(a.����, 0) As ����, a.����, a.����, a.Ч��, a.��ҩ��,
           a.��ҩ���� As ��������, a.ժҪ, a.��ҩ��λid, a.��׼�ĺ�, a.��������, a.�ɱ���, a.ʵ������, a.���۽��, a.���, a.���ۼ�, Nvl(b.�Ƿ���, 0) As ʱ��,
           a.����, a.����, a.Ƶ��, a.��Ʒ����, a.�ڲ�����, a.�������, a.���Ч��
    From ҩƷ�շ���¼ A, �շ���ĿĿ¼ B
    Where a.ҩƷid = b.Id And a.No = No_In And a.���� = 19 And (a.��� >= ���_In And a.��� <= ���_In + 1) And
          (a.��¼״̬ = ԭ��¼״̬_In And Mod(a.��¼״̬, 3) = 2) And a.������� Is Null
    Order By a.ҩƷid,a.����;
Begin
  --��ȡ���С��λ��
  Select Nvl(����, 2) Into n_С�� From ҩƷ���ľ��� Where ���� = 0 And ��� = 2 And ���� = 4 And ��λ = 5;
  Select Nvl(Zl_Getsysparameter(95), '0') Into v_�¿��ÿ�� From Dual;

  If ������ʽ_In = 1 Then
    --���������ֻ�����������ݣ�����д����ˣ�������ڣ�Ҳ�����¿��
    If �д�_In = 1 Then
      Update ҩƷ�շ���¼
      Set ��¼״̬ = Decode(ԭ��¼״̬_In, 1, 3, ԭ��¼״̬_In + 3)
      Where NO = No_In And ���� = 19 And ��¼״̬ = ԭ��¼״̬_In;
    
      If Sql%RowCount = 0 Then
        v_Err_Msg := '[ZLSOFT]�õ����Ѿ������˳�����[ZLSOFT]';
        Raise Err_Item;
      End If;
    End If;
  
    --ԭ�������������ڷ������������ϣ����ܳ���
    Select Count(*)
    Into n_Batch_Count
    From ҩƷ�շ���¼ A, �������� B
    Where a.ҩƷid = b.����id And a.No = No_In And a.���� = 19 And a.ҩƷid + 0 = ����id_In And Mod(a.��¼״̬, 3) = 0 And
          Nvl(a.����, 0) = 0 And
          ((Nvl(b.�ⷿ����, 0) = 1 And
          a.�ⷿid Not In (Select ����id From ��������˵�� Where (�������� Like '%���ϲ���') Or (�������� Like '�Ƽ���'))) Or
          Nvl(b.���÷���, 0) = 1);
  
    If n_Batch_Count > 0 Then
      v_Err_Msg := '[ZLSOFT]�õ����а�����һ��ԭ�������������ڷ������������ϣ����ܳ�����[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    --��ȡ��ǰ��������ʣ������
    Select Sum(a.ʵ������) As ʣ������, Sum(a.�ɱ����) As ʣ��ɱ����, Sum(a.���۽��) As ʣ�����۽��, a.�ɱ���, a.���ۼ�, a.�Է�����id, Nvl(a.����, 0),
           b.�ⷿ����, b.���÷���
    Into n_ʣ������, n_ʣ��ɱ����, n_ʣ�����۽��, n_�ɱ���, n_���ۼ�, n_�ⷿid, n_����, n_�ⷿ����, n_���÷���
    From ҩƷ�շ���¼ A, �������� B
    Where a.No = No_In And a.ҩƷid = b.����id And a.���� = 19 And a.ҩƷid + 0 = ����id_In And a.��� = ���_In
    Group By a.�ɱ���, a.���ۼ�, a.�Է�����id, Nvl(a.����, 0), b.�ⷿ����, b.���÷���;
    --�жϸò����ǿⷿ���Ƿ��ϲ���
    --n_����:(ԭ����,�ַ���Ϊ����;����Ϊ��)ԭ������,�ַ���,�����̲�����
    --��Ϊ���ڶԷ��ⷿ���൱�ڳ��⣬�����ڵ�ǰ�ⷿ���൱����⣬���Ե�ǰ�ⷿ�����飬������˿�������¼
    Select Nvl(a.����, 0)
    Into n_����
    From ҩƷ�շ���¼ A
    Where a.No = No_In And a.���� = 19 And a.ҩƷid + 0 = ����id_In And a.��� = ���_In + 1 And Mod(a.��¼״̬, 3) = 0;
  
    --ȡ�����
    Begin
      Select Nvl(ʵ������, 0)
      Into n_�����
      From ҩƷ���
      Where �ⷿid = n_�ⷿid And ҩƷid = ����id_In And Nvl(����, 0) = n_���� And ���� = 1;
    Exception
      When Others Then
        n_����� := 0;
    End;
  
    If Nvl(n_ʣ������, 0) = 0 Then
      v_Err_Msg := '[ZLSOFT]�õ����е�' || Ceil(���_In / 2) || '�еĲ����Ѿ����������,�����ٳ壡[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    --������������ʣ������,ȡʣ������;����ȡ�����
    If n_����� < n_ʣ������ Then
      n_ʣ��ɱ���� := n_����� / n_ʣ������ * n_ʣ��ɱ����;
      n_ʣ�����۽�� := n_����� / n_ʣ������ * n_ʣ�����۽��;
      n_ʣ������     := n_�����;
    End If;
  
    --������������ʣ��������������
    If n_ʣ������ < ��������_In Then
      v_Err_Msg := '[ZLSOFT]�õ����е�' || Ceil(���_In / 2) || '�е��������ϳ���������������ʣ������ݣ����ܳ�����[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    n_�ɱ���� := Round(��������_In / n_ʣ������ * n_ʣ��ɱ����, n_С��);
    n_���۽�� := Round(��������_In / n_ʣ������ * n_ʣ�����۽��, n_С��);
    n_�����   := Round(n_���۽�� - n_�ɱ����, n_С��);
  
    For v_ҩƷ�շ���¼ In c_ҩƷ�շ���¼ Loop
      n_�ⷿid     := v_ҩƷ�շ���¼.�ⷿid;
      n_�Է�����id := v_ҩƷ�շ���¼.�Է�����id;
      n_����       := v_ҩƷ�շ���¼.����;
      n_���ۼ�     := v_ҩƷ�շ���¼.���ۼ�;
      n_���ϵ��   := v_ҩƷ�շ���¼.���ϵ��;
      n_�ɱ���     := v_ҩƷ�շ���¼.�ɱ���;
      v_����       := v_ҩƷ�շ���¼.����;
      v_����       := v_ҩƷ�շ���¼.����;
      v_Ч��       := v_ҩƷ�շ���¼.Ч��;
      v_��Ʒ����   := v_ҩƷ�շ���¼.��Ʒ����;
      v_�ڲ�����   := v_ҩƷ�շ���¼.�ڲ�����;
      v_���Ч��   := v_ҩƷ�շ���¼.���Ч��;
      v_�������   := v_ҩƷ�շ���¼.�������;
      n_��ҩ��λid := v_ҩƷ�շ���¼.��ҩ��λid;
      d_��������   := v_ҩƷ�շ���¼.��������;
      v_��׼�ĺ�   := v_ҩƷ�շ���¼.��׼�ĺ�;
      n_����       := v_ҩƷ�շ���¼.����;
      n_���       := v_ҩƷ�շ���¼.���;
      n_������id := v_ҩƷ�շ���¼.������id;
      v_��ҩ��     := v_ҩƷ�շ���¼.��ҩ��;
      d_��������   := v_ҩƷ�շ���¼.��������;
      v_ժҪ       := v_ҩƷ�շ���¼.ժҪ;
    
      Select ҩƷ�շ���¼_Id.Nextval Into n_�շ�id From Dual;
      Select �Ƿ��� Into n_ʵ������ From �շ���ĿĿ¼ Where ID = ����id_In;
    
      Insert Into ҩƷ�շ���¼
        (ID, ��¼״̬, ����, NO, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, �������, ���Ч��, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�,
         ���۽��, ���, ժҪ, ������, ��������, ��ҩ��, ��ҩ����, ��ҩ��λid, ��������, ��׼�ĺ�, ����, ��Ʒ����, �ڲ�����)
      Values
        (n_�շ�id, Decode(ԭ��¼״̬_In, 1, 2, ԭ��¼״̬_In + 2), 19, No_In, n_���, n_�ⷿid, n_�Է�����id, n_������id, n_���ϵ��, ����id_In,
         n_����, v_����, v_����, v_Ч��, v_�������, v_���Ч��, -��������_In, -��������_In, n_�ɱ���, -n_�ɱ����, n_���ۼ�, -n_���۽��, -n_�����, v_ժҪ,
         ������_In, ��������_In, v_��ҩ��, d_��������, n_��ҩ��λid, d_��������, v_��׼�ĺ�, n_����, v_��Ʒ����, v_�ڲ�����);
    End Loop;
  
    --����Ϊ1��ʾ�������ʱ�¿�������������ԭ����ⷿ
    If v_�¿��ÿ�� = '1' And n_���ϵ�� = 1 Then
      Update ҩƷ���
      Set �������� = Nvl(��������, 0) - ��������_In
      Where �ⷿid = n_�ⷿid And ҩƷid = ����id_In And Nvl(����, 0) = n_���� And ���� = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into ҩƷ���
          (�ⷿid, ҩƷid, ����, ����, ��������, �ϴ�����, Ч��, �ϴβ���)
        Values
          (n_�ⷿid, ����id_In, n_����, 1, -1 * ��������_In, v_����, v_Ч��, v_����);
      End If;
    
      Delete From ҩƷ���
      Where �ⷿid = n_�ⷿid And ҩƷid = ����id_In And Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0 And
            Nvl(ʵ�ʲ��, 0) = 0;
    End If;
  
  Elsif ������ʽ_In = 2 Then
  
    --���������������ĵ��ݣ�����д����ˣ�������ڣ������¿��
    For v_���������¼ In c_���������¼ Loop
      n_�ⷿid     := v_���������¼.�ⷿid;
      n_�Է�����id := v_���������¼.�Է�����id;
      n_����       := v_���������¼.����;
      n_���ۼ�     := v_���������¼.���ۼ�;
      n_���۽��   := v_���������¼.���۽��;
      n_�����     := v_���������¼.���;
      n_���ϵ��   := v_���������¼.���ϵ��;
      n_�ɱ���     := v_���������¼.�ɱ���;
      v_����       := v_���������¼.����;
      v_����       := v_���������¼.����;
      v_Ч��       := v_���������¼.Ч��;
      v_��Ʒ����   := v_���������¼.��Ʒ����;
      v_�ڲ�����   := v_���������¼.�ڲ�����;
      v_���Ч��   := v_���������¼.���Ч��;
      v_�������   := v_���������¼.�������;
      n_��ҩ��λid := v_���������¼.��ҩ��λid;
      d_��������   := v_���������¼.��������;
      v_��׼�ĺ�   := v_���������¼.��׼�ĺ�;
      n_����       := v_���������¼.����;
      --ԭ�����ֲ������Ĳ���,�ڳ���ʱ��Ҫ������
      Begin
        Select Count(*)
        Into n_��¼��
        From ҩƷ�շ���¼ A, �������� B
        Where b.����id = a.ҩƷid And a.ҩƷid = ����id_In And a.No = No_In And a.���� = 19 And a.�ⷿid = n_�ⷿid And
              Mod(a.��¼״̬, 3) = 0 And Nvl(a.����, 0) > 0 And
              (Nvl(b.�ⷿ����, 0) = 0 Or
              (Nvl(b.���÷���, 0) = 0 And
              a.�ⷿid In (Select ����id From ��������˵�� Where (�������� Like '%���ϲ���') Or (�������� Like '�Ƽ���'))));
      Exception
        When Others Then
          n_��¼�� := 0;
      End;
      If n_��¼�� > 0 Then
        n_���� := 0;
      Else
        n_���� := Nvl(n_����, 0);
      End If;
      Select �Ƿ��� Into n_ʵ������ From �շ���ĿĿ¼ Where ID = ����id_In;
      --����ʱ���Ѿ����˿��ÿ������ط��Ͳ���������
      If v_�¿��ÿ�� = '1' And n_���ϵ�� = 1 Then
        n_�������� := 0;
      Else
        n_�������� := ��������_In;
      End If;
    
      --����ҩƷ�������Ӧ����
      Update ҩƷ���
      Set �������� = Nvl(��������, 0) + Nvl(n_��������, 0) * n_���ϵ��, ʵ������ = Nvl(ʵ������, 0) + Nvl(��������_In, 0) * n_���ϵ��,
          ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) + Nvl(n_���۽��, 0) * n_���ϵ��, ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + Nvl(n_�����, 0) * n_���ϵ��,
          �ϴβɹ��� = Nvl(n_�ɱ���, �ϴβɹ���), �ϴ����� = Nvl(v_����, �ϴ�����), �ϴβ��� = Nvl(v_����, �ϴβ���), Ч�� = Nvl(v_Ч��, Ч��),
          ���ۼ� = Decode(n_ʵ������, 1, Decode(Nvl(����, 0), 0, Null, Decode(Nvl(���ۼ�, 0), 0, n_���ۼ�, ���ۼ�)), Null),
          ��Ʒ���� = Nvl(��Ʒ����, v_��Ʒ����), �ڲ����� = Nvl(�ڲ�����, v_�ڲ�����)
      Where �ⷿid = n_�ⷿid And ҩƷid = ����id_In And Nvl(����, 0) = n_���� And ���� = 1;
    
      If Sql%NotFound Then
        Insert Into ҩƷ���
          (�ⷿid, ҩƷid, ����, ����, ��������, ʵ������, ʵ�ʽ��, ʵ�ʲ��, �ϴβɹ���, �ϴ�����, �ϴβ���, Ч��, ���Ч��, �ϴι�Ӧ��id, �ϴ���������, ��׼�ĺ�, ���ۼ�, �ϴο���,
           ��Ʒ����, �ڲ�����, ƽ���ɱ���)
        Values
          (n_�ⷿid, ����id_In, n_����, 1, ��������_In * n_���ϵ��, ��������_In * n_���ϵ��, n_���۽�� * n_���ϵ��, n_����� * n_���ϵ��, n_�ɱ���, v_����,
           v_����, v_Ч��, v_���Ч��, n_��ҩ��λid, d_��������, v_��׼�ĺ�, Decode(n_ʵ������, 1, Decode(Nvl(n_����, 0), 0, Null, n_���ۼ�), Null),
           n_����, v_��Ʒ����, v_�ڲ�����, n_�ɱ���);
      End If;
    
      Delete From ҩƷ���
      Where �ⷿid = n_�ⷿid And ҩƷid = ����id_In And Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0 And
            Nvl(ʵ�ʲ��, 0) = 0;
    
      --��д����ˡ��������
      Update ҩƷ�շ���¼
      Set ����� = ������_In, ������� = ��������_In
      Where NO = No_In And ���� = 19 And ID = v_���������¼.Id;
    
      Zl_�����շ���¼_��������(v_���������¼.Id);
    End Loop;
  Else
    --��������ҵ�񣬲�������������˲����¿��
    If �д�_In = 1 Then
      Update ҩƷ�շ���¼
      Set ��¼״̬ = Decode(ԭ��¼״̬_In, 1, 3, ԭ��¼״̬_In + 3)
      Where NO = No_In And ���� = 19 And ��¼״̬ = ԭ��¼״̬_In;
      If Sql%RowCount = 0 Then
        v_Err_Msg := '[ZLSOFT]�õ����Ѿ������˳�����[ZLSOFT]';
        Raise Err_Item;
      End If;
    End If;
  
    Select Count(*)
    Into n_Batch_Count
    From ҩƷ�շ���¼ A, �������� B
    Where a.ҩƷid = b.����id And a.No = No_In And a.���� = 19 And a.ҩƷid + 0 = ����id_In And Mod(a.��¼״̬, 3) = 0 And
          Nvl(a.����, 0) = 0 And
          ((Nvl(b.�ⷿ����, 0) = 1 And
          a.�ⷿid Not In (Select ����id From ��������˵�� Where (�������� Like '%���ϲ���') Or (�������� Like '�Ƽ���'))) Or
          Nvl(b.���÷���, 0) = 1);
  
    If n_Batch_Count > 0 Then
      v_Err_Msg := '[ZLSOFT]�õ����а�����һ��ԭ�������������ڷ������������ϣ����ܳ�����[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    Select Sum(a.ʵ������) As ʣ������, Sum(a.�ɱ����) As ʣ��ɱ����, Sum(a.���۽��) As ʣ�����۽��, a.�ɱ���, a.���ۼ�, a.�Է�����id, Nvl(a.����, 0),
           b.�ⷿ����, b.���÷���
    Into n_ʣ������, n_ʣ��ɱ����, n_ʣ�����۽��, n_�ɱ���, n_���ۼ�, n_�ⷿid, n_����, n_�ⷿ����, n_���÷���
    From ҩƷ�շ���¼ A, �������� B
    Where a.No = No_In And a.ҩƷid = b.����id And a.���� = 19 And a.ҩƷid + 0 = ����id_In And a.��� = ���_In
    Group By a.�ɱ���, a.���ۼ�, a.�Է�����id, Nvl(a.����, 0), b.�ⷿ����, b.���÷���;
  
    --�жϸò����ǿⷿ���Ƿ��ϲ���
    --n_����:(ԭ����,�ַ���Ϊ����;����Ϊ��)ԭ������,�ַ���,�����̲�����
    --��Ϊ���ڶԷ��ⷿ���൱�ڳ��⣬�����ڵ�ǰ�ⷿ���൱����⣬���Ե�ǰ�ⷿ�����飬������˿�������¼
    Select Nvl(a.����, 0)
    Into n_����
    From ҩƷ�շ���¼ A
    Where a.No = No_In And a.���� = 19 And a.ҩƷid + 0 = ����id_In And a.��� = ���_In + 1 And Mod(a.��¼״̬, 3) = 0;
  
    --ȡ�����
    Begin
      Select Nvl(ʵ������, 0)
      Into n_�����
      From ҩƷ���
      Where �ⷿid = n_�ⷿid And ҩƷid = ����id_In And Nvl(����, 0) = n_���� And ���� = 1;
    Exception
      When Others Then
        n_����� := 0;
    End;
  
    If Nvl(n_ʣ������, 0) = 0 Then
      v_Err_Msg := '[ZLSOFT]�õ����е�' || Ceil(���_In / 2) || '�еĲ����Ѿ����������,�����ٳ壡[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    --������������ʣ������,ȡʣ������;����ȡ�����
    If n_����� < n_ʣ������ Then
      n_ʣ��ɱ���� := n_����� / n_ʣ������ * n_ʣ��ɱ����;
      n_ʣ�����۽�� := n_����� / n_ʣ������ * n_ʣ�����۽��;
      n_ʣ������     := n_�����;
    End If;
  
    --������������ʣ��������������
    If n_ʣ������ < ��������_In Then
      v_Err_Msg := '[ZLSOFT]�õ����е�' || Ceil(���_In / 2) || '�е��������ϳ���������������ʣ������ݣ����ܳ�����[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    n_�ɱ���� := Round(��������_In / n_ʣ������ * n_ʣ��ɱ����, n_С��);
    n_���۽�� := Round(��������_In / n_ʣ������ * n_ʣ�����۽��, n_С��);
    n_�����   := Round(n_���۽�� - n_�ɱ����, n_С��);
  
    For v_ҩƷ�շ���¼ In c_ҩƷ�շ���¼ Loop
      n_�ⷿid     := v_ҩƷ�շ���¼.�ⷿid;
      n_�Է�����id := v_ҩƷ�շ���¼.�Է�����id;
      n_����       := v_ҩƷ�շ���¼.����;
      n_���ۼ�     := v_ҩƷ�շ���¼.���ۼ�;
      n_���ϵ��   := v_ҩƷ�շ���¼.���ϵ��;
      n_�ɱ���     := v_ҩƷ�շ���¼.�ɱ���;
      v_����       := v_ҩƷ�շ���¼.����;
      v_����       := v_ҩƷ�շ���¼.����;
      v_Ч��       := v_ҩƷ�շ���¼.Ч��;
      v_��Ʒ����   := v_ҩƷ�շ���¼.��Ʒ����;
      v_�ڲ�����   := v_ҩƷ�շ���¼.�ڲ�����;
      v_���Ч��   := v_ҩƷ�շ���¼.���Ч��;
      v_�������   := v_ҩƷ�շ���¼.�������;
      n_��ҩ��λid := v_ҩƷ�շ���¼.��ҩ��λid;
      d_��������   := v_ҩƷ�շ���¼.��������;
      v_��׼�ĺ�   := v_ҩƷ�շ���¼.��׼�ĺ�;
      n_����       := v_ҩƷ�շ���¼.����;
      n_���       := v_ҩƷ�շ���¼.���;
      n_������id := v_ҩƷ�շ���¼.������id;
      v_��ҩ��     := v_ҩƷ�շ���¼.��ҩ��;
      d_��������   := v_ҩƷ�շ���¼.��������;
      v_ժҪ       := v_ҩƷ�շ���¼.ժҪ;
    
      Select ҩƷ�շ���¼_Id.Nextval Into n_�շ�id From Dual;
      Select �Ƿ��� Into n_ʵ������ From �շ���ĿĿ¼ Where ID = ����id_In;
    
      Insert Into ҩƷ�շ���¼
        (ID, ��¼״̬, ����, NO, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, �������, ���Ч��, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�,
         ���۽��, ���, ժҪ, ������, ��������, �����, �������, ��ҩ��, ��ҩ����, ��ҩ��λid, ��������, ��׼�ĺ�, ����, ��Ʒ����, �ڲ�����)
      Values
        (n_�շ�id, Decode(ԭ��¼״̬_In, 1, 2, ԭ��¼״̬_In + 2), 19, No_In, n_���, n_�ⷿid, n_�Է�����id, n_������id, n_���ϵ��, ����id_In,
         n_����, v_����, v_����, v_Ч��, v_�������, v_���Ч��, -��������_In, -��������_In, n_�ɱ���, -n_�ɱ����, n_���ۼ�, -n_���۽��, -n_�����, v_ժҪ,
         ������_In, ��������_In, ������_In, ��������_In, v_��ҩ��, d_��������, n_��ҩ��λid, d_��������, v_��׼�ĺ�, n_����, v_��Ʒ����, v_�ڲ�����);
    
      --ԭ�����ֲ������Ĳ���,�ڳ���ʱ��Ҫ������
      Begin
        Select Count(*)
        Into n_��¼��
        From ҩƷ�շ���¼ A, �������� B
        Where b.����id = a.ҩƷid And a.ҩƷid = ����id_In And a.No = No_In And a.���� = 19 And a.�ⷿid = n_�ⷿid And
              Mod(a.��¼״̬, 3) = 0 And Nvl(a.����, 0) > 0 And
              (Nvl(b.�ⷿ����, 0) = 0 Or
              (Nvl(b.���÷���, 0) = 0 And
              a.�ⷿid In (Select ����id From ��������˵�� Where (�������� Like '%���ϲ���') Or (�������� Like '�Ƽ���'))));
      Exception
        When Others Then
          n_��¼�� := 0;
      End;
      If n_��¼�� > 0 Then
        n_���� := 0;
      Else
        n_���� := Nvl(n_����, 0);
      End If;
    
      --����ҩƷ�������Ӧ����
      Update ҩƷ���
      Set �������� = Nvl(��������, 0) - Nvl(��������_In, 0) * n_���ϵ��, ʵ������ = Nvl(ʵ������, 0) - Nvl(��������_In, 0) * n_���ϵ��,
          ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) - Nvl(n_���۽��, 0) * n_���ϵ��, ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) - Nvl(n_�����, 0) * n_���ϵ��,
          �ϴβɹ��� = Nvl(n_�ɱ���, �ϴβɹ���), �ϴ����� = Nvl(v_����, �ϴ�����), �ϴβ��� = Nvl(v_����, �ϴβ���), Ч�� = Nvl(v_Ч��, Ч��),
          ���ۼ� = Decode(n_ʵ������, 1, Decode(Nvl(����, 0), 0, Null, Decode(Nvl(���ۼ�, 0), 0, n_���ۼ�, ���ۼ�)), Null),
          ��Ʒ���� = Nvl(��Ʒ����, v_��Ʒ����), �ڲ����� = Nvl(�ڲ�����, v_�ڲ�����)
      Where �ⷿid = n_�ⷿid And ҩƷid = ����id_In And Nvl(����, 0) = n_���� And ���� = 1;
    
      If Sql%NotFound Then
        Insert Into ҩƷ���
          (�ⷿid, ҩƷid, ����, ����, ��������, ʵ������, ʵ�ʽ��, ʵ�ʲ��, �ϴβɹ���, �ϴ�����, �ϴβ���, Ч��, ���Ч��, �ϴι�Ӧ��id, �ϴ���������, ��׼�ĺ�, ���ۼ�, �ϴο���,
           ��Ʒ����, �ڲ�����, ƽ���ɱ���)
        Values
          (n_�ⷿid, ����id_In, n_����, 1, -��������_In * n_���ϵ��, -��������_In * n_���ϵ��, -n_���۽�� * n_���ϵ��, -n_����� * n_���ϵ��, n_�ɱ���,
           v_����, v_����, v_Ч��, v_���Ч��, n_��ҩ��λid, d_��������, v_��׼�ĺ�,
           Decode(n_ʵ������, 1, Decode(Nvl(n_����, 0), 0, Null, n_���ۼ�), Null), n_����, v_��Ʒ����, v_�ڲ�����, n_�ɱ���);
      End If;
    
      Delete From ҩƷ���
      Where �ⷿid = n_�ⷿid And ҩƷid = ����id_In And Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0 And
            Nvl(ʵ�ʲ��, 0) = 0;
    
      --���¼�������е�ƽ���ɱ���
      Update ҩƷ���
      Set ƽ���ɱ��� = Decode(Nvl(����, 0), 0, Decode((ʵ�ʽ�� - ʵ�ʲ��) / ʵ������, 0, �ϴβɹ���, (ʵ�ʽ�� - ʵ�ʲ��) / ʵ������), �ϴβɹ���)
      Where ���� = 1 And �ⷿid = n_�ⷿid And ҩƷid = ����id_In And Nvl(����, 0) = n_���� And Nvl(ʵ������, 0) <> 0;
      If Sql%NotFound Then
        Select �ɱ��� Into n_ƽ���ɱ��� From �������� Where ����id = ����id_In;
        Update ҩƷ���
        Set ƽ���ɱ��� = n_ƽ���ɱ���
        Where ���� = 1 And �ⷿid = n_�ⷿid And ҩƷid = ����id_In And Nvl(����, 0) = n_���� And Nvl(ƽ���ɱ���, 0) <> n_�ɱ���;
      End If;
      --������ۺ����
      Zl_�����շ���¼_��������(n_�շ�id);
    End Loop;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����ƿ�_Strike;
/

--117925:����,2017-12-07,��������������
Create Or Replace Procedure Zl_�����ƿ�_Delete
(
  No_In       In ҩƷ�շ���¼.No%Type,
  ��¼״̬_In In ҩƷ�շ���¼.��¼״̬%Type := 1
) Is
  v_���� ҩƷ�շ���¼.��ҩ����%Type;

  v_Err_Msg Varchar2(200);
  Err_Item Exception;
  v_�¿��   Zlparameters.����ֵ%Type;
  n_ʵ������ �շ���ĿĿ¼.�Ƿ���%Type;

  Cursor c_ҩƷ�շ���¼ Is
    Select ʵ������, ���ۼ�, �ⷿid, Nvl(����, 0) As ����, ҩƷid, ����, Ч��, ����, ��ҩ��λid, �ɱ���, ���Ч��, ��������, ��׼�ĺ� 
    From ҩƷ�շ���¼
    Where NO = No_In And ���� = 19 And ���ϵ�� = -1
    Order By ҩƷid,����;

  Cursor c_���������¼ Is
    Select (-1 * ʵ������) ʵ������, �ⷿid, Nvl(����, 0) As ����, ҩƷid, ����, Ч��, ����, ��ҩ��λid, ��׼�ĺ�, �ɱ���, �������� 
    From ҩƷ�շ���¼
    Where NO = No_In And ���� = 19 And ���ϵ�� = 1 And ��¼״̬ = ��¼״̬_In
    Order By ҩƷid,����;
Begin
  Select Nvl(Zl_Getsysparameter(95), '0') Into v_�¿�� From Dual;

  If ��¼״̬_In = 1 Then
    --����Ƿ��ѷ��ͣ��ѷ��͵ĵ�����Ҫ��ԭ�������� 
    Select ��ҩ���� Into v_���� From ҩƷ�շ���¼ Where ���� = 19 And NO = No_In And Rownum < 2;
  
    If v_���� Is Not Null Or To_Number(v_�¿��, '9999') = 1 Then
      --ͨ��ѭ�����ָ�ԭ���Ŀ������� 
      For c_���� In c_ҩƷ�շ���¼ Loop
        Select �Ƿ��� Into n_ʵ������ From �շ���ĿĿ¼ Where ID = c_����.ҩƷid;
      
        Update ҩƷ���
        Set �������� = Nvl(��������, 0) + c_����.ʵ������
        Where �ⷿid = c_����.�ⷿid And ҩƷid = c_����.ҩƷid And Nvl(����, 0) = Nvl(c_����.����, 0) And ���� = 1;
      
        If Sql%NotFound Then
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, ��������, Ч��, ���Ч��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴ���������, �ϴβ���, ��׼�ĺ�, ���ۼ�)
          Values
            (c_����.�ⷿid, c_����.ҩƷid, c_����.����, 1, c_����.ʵ������, c_����.Ч��, c_����.���Ч��, c_����.��ҩ��λid, c_����.�ɱ���, c_����.����, c_����.��������,
             c_����.����, c_����.��׼�ĺ�, Decode(n_ʵ������, 1, Decode(Nvl(c_����.����, 0), 0, Null, c_����.���ۼ�), Null));
        End If;
      
        Delete From ҩƷ���
        Where �ⷿid = c_����.�ⷿid And ҩƷid = c_����.ҩƷid And Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0 And
              Nvl(ʵ�ʲ��, 0) = 0;
      End Loop;
    End If;
  Else
    --�����ƿ������������ 
  
    --�������ֵΪ1ҲҪ�ָ�ԭ���Ŀ������� 
    If v_�¿�� = '1' Then
      --ͨ��ѭ�����ָ�ԭ���Ŀ������� 
      For v_���������¼ In c_���������¼ Loop
        Update ҩƷ���
        Set �������� = Nvl(��������, 0) + v_���������¼.ʵ������
        Where �ⷿid = v_���������¼.�ⷿid And ҩƷid = v_���������¼.ҩƷid And Nvl(����, 0) = v_���������¼.���� And ���� = 1;
      
        If Sql%NotFound Then
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, ��������, �ϴ�����, Ч��, �ϴβ���, �ϴι�Ӧ��id, ��׼�ĺ�, �ϴβɹ���, �ϴ���������)
          Values
            (v_���������¼.�ⷿid, v_���������¼.ҩƷid, v_���������¼.����, 1, v_���������¼.ʵ������, v_���������¼.����, v_���������¼.Ч��, v_���������¼.����,
             v_���������¼.��ҩ��λid, v_���������¼.��׼�ĺ�, v_���������¼.�ɱ���, v_���������¼.��������);
        End If;
      
        Delete From ҩƷ���
        Where �ⷿid = v_���������¼.�ⷿid And ҩƷid = v_���������¼.ҩƷid And Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And
              Nvl(ʵ�ʽ��, 0) = 0 And Nvl(ʵ�ʲ��, 0) = 0;
      End Loop;
    End If;
  End If;

  Delete --����ͳ����������ƿⵥ��ɾ�� 
  From ҩƷ�շ���¼
  Where NO = No_In And ���� = 19 And ��¼״̬ = ��¼״̬_In And ����� Is Null;

  If Sql%RowCount = 0 Then
    v_Err_Msg := '[ZLSOFT]�õ����Ѿ�������ɾ�����ѱ�����ˣ�[ZLSOFT]';
    Raise Err_Item;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����ƿ�_Delete;
/

--117925:����,2017-12-07,��������������
CREATE OR REPLACE Procedure Zl_���Ͽ���۵���_Verify
(
  No_In     In ҩƷ�շ���¼.NO%Type := Null,
  �����_In In ҩƷ�շ���¼.�����%Type := Null
) Is
  Err_Item Exception;
  v_Err_Msg Varchar2(100);

  Cursor c_���ϵ�����Ϣ Is
    Select �ⷿid, ҩƷid ����id, ����, ������id, ���, ����, Ч��, ���Ч��, �ɱ���,���� as �³ɱ���, ��ҩ��λid, ��������, ��׼�ĺ�,
           ����
    From ҩƷ�շ���¼
    Where NO = No_In And ���� = 18 And ��¼״̬ = 1
    Order By ҩƷid, ����;
Begin
  Update ҩƷ�շ���¼
  Set ����� = �����_In, ������� = Sysdate
  Where NO = No_In And ���� = 18 And ��¼״̬ = 1 And ����� Is Null;

  If Sql%RowCount = 0 Then
    v_Err_Msg := '[ZLSOFT]�õ����Ѿ���������ˣ�[ZLSOFT]';
    Raise Err_Item;
  End If;

  For c_���� In c_���ϵ�����Ϣ Loop
    --����ҩƷ�������Ӧ����

    Update ҩƷ���
    Set ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + Nvl(c_����.���, 0),�ϴβɹ���=c_����.�³ɱ���,ƽ���ɱ���=c_����.�³ɱ���
    Where �ⷿid = c_����.�ⷿid And ҩƷid = c_����.����id And Nvl(����, 0) = Nvl(c_����.����, 0) And ���� = 1;

    If Sql%NotFound Then

      Insert Into ҩƷ���
        (�ⷿid, ҩƷid, ����, ����, ʵ�ʲ��, Ч��, ���Ч��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴ���������,
         �ϴβ���, ��׼�ĺ�,ƽ���ɱ���)
      Values
        (c_����.�ⷿid, c_����.����id, c_����.����, 1, c_����.���, c_����.Ч��, c_����.���Ч��, c_����.��ҩ��λid,
         c_����.�³ɱ���, c_����.����, c_����.��������, c_����.����, c_����.��׼�ĺ�,c_����.�³ɱ���);
    End If;

    Delete From ҩƷ���
    Where �ⷿid = c_����.�ⷿid And ҩƷid = c_����.����id And Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And
          Nvl(ʵ�ʽ��, 0) = 0 And Nvl(ʵ�ʲ��, 0) = 0;
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���Ͽ���۵���_Verify;
/

--117925:����,2017-12-07,��������������
Create Or Replace Procedure Zl_������������_Delete(
                                                   --ɾ��ҩƷ�շ���¼���ָ���Ӧ�ı�ҩƷ���
                                                   No_In In ҩƷ�շ���¼.NO%Type) Is
  v_Err_Msg Varchar2(500);
  Err_Item Exception;
  n_ʵ������ �շ���ĿĿ¼.�Ƿ���%Type;

Begin
  --ͨ��ѭ�����ָ�ԭ���Ŀ�������
  For c_���� In (Select ��д����, �ⷿid, ���ۼ�, ����, Ч��, ���Ч��, ҩƷid ����id, �ɱ���, ��ҩ��λid, ��������, ����,
                        ��׼�ĺ�, ����
                 From ҩƷ�շ���¼
                 Where NO = No_In And ���� = 21
                 Order By ҩƷid, ����) Loop
    Select �Ƿ��� Into n_ʵ������ From �շ���ĿĿ¼ Where ID = c_����.����id;

    Update ҩƷ���
    Set �������� = Nvl(��������, 0) + c_����.��д����,
        ���ۼ� = Decode(n_ʵ������, 1,
                         Decode(Nvl(c_����.����, 0), 0, Null, Decode(Nvl(���ۼ�, 0), 0, c_����.���ۼ�, ���ۼ�)), Null)
    Where �ⷿid = c_����.�ⷿid And ҩƷid = c_����.����id And Nvl(����, 0) = Nvl(c_����.����, 0) And ���� = 1;

    If Sql%NotFound Then

      Insert Into ҩƷ���
        (�ⷿid, ҩƷid, ����, ����, ��������, Ч��, ���Ч��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴ���������,
         �ϴβ���, ��׼�ĺ�, ���ۼ�)
      Values
        (c_����.�ⷿid, c_����.����id, c_����.����, 1, c_����.��д����, c_����.Ч��, c_����.���Ч��, c_����.��ҩ��λid,
         c_����.�ɱ���, c_����.����, c_����.��������, c_����.����, c_����.��׼�ĺ�,
         Decode(n_ʵ������, 1, Decode(Nvl(c_����.����, 0), 0, Null, c_����.���ۼ�), Null));
    End If;

    Delete From ҩƷ���
    Where �ⷿid = c_����.�ⷿid And ҩƷid = c_����.����id And Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And
          Nvl(ʵ�ʽ��, 0) = 0 And Nvl(ʵ�ʲ��, 0) = 0;
  End Loop;

  Delete From ҩƷ�շ���¼ Where NO = No_In And ���� = 21 And ��¼״̬ = 1 And ����� Is Null;

  If Sql%RowCount = 0 Then
    v_Err_Msg := '[ZLSOFT]�õ����Ѿ�������ɾ�����ѱ�����ˣ�[ZLSOFT]';
    Raise Err_Item;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_������������_Delete;
/

--117925:����,2017-12-07,��������������
Create Or Replace Procedure Zl_��������ԭ�ϳ���_Insert
(
  No_In         In ҩƷ�շ���¼.NO%Type,
  �Է�����id_In In ҩƷ�շ���¼.�Է�����id%Type
) As
  v_Err_Msg Varchar2(100);
  Err_Item Exception;

  n_����       ҩƷ�շ���¼.ʵ������%Type;
  n_�ɱ���     ҩƷ�շ���¼.�ɱ���%Type;
  n_�ɱ����   ҩƷ�շ���¼.�ɱ����%Type;
  n_���       ҩƷ�շ���¼.���%Type;
  n_�ۼ�       ҩƷ�շ���¼.���ۼ�%Type;
  n_���۽��   ҩƷ�շ���¼.���۽��%Type;
  n_�����   ҩƷ���.ʵ�ʽ��%Type;
  n_�����   ҩƷ���.ʵ�ʲ��%Type;
  n_��������   ҩƷ���.��������%Type;
  n_ʵ������   ҩƷ���.ʵ������%Type;
  n_�������id ҩƷ�շ���¼.������id%Type; --������ID
  n_Max_���   ҩƷ�շ���¼.���%Type;

  v_�ϴβ���   ҩƷ���.�ϴβ���%Type;
  v_���ɱ����� zlParameters.����ֵ%Type;
Begin
  Select B.ID
  Into n_�������id
  From ҩƷ�������� A, ҩƷ������ B
  Where A.���id = B.ID And A.���� = 31 And B.ϵ�� = -1 And Rownum < 2;

  Select zl_GetSysParameter(120) Into v_���ɱ����� From Dual;

  Select Max(���) Into n_Max_��� From ҩƷ�շ���¼ Where NO = No_In And ���� = 16 And ���ϵ�� = 1;

  For v_���� In (Select * From ҩƷ�շ���¼ Where NO = No_In And ���� = 16 And ���ϵ�� = 1 Order By ҩƷid, ����) Loop

    For v_��� In (Select A.*, B.�Ƿ���, C.ָ�������, C.�ɱ���
                   From ���Ʋ��Ϲ��� A, �շ���ĿĿ¼ B, �������� C
                   Where A.ԭ�ϲ���id = B.ID And A.���Ʋ���id = v_����.ҩƷid And A.ԭ�ϲ���id = C.����id
				   Order By A.ԭ�ϲ���id) Loop

      Begin
        Select ��������, ʵ������, ʵ�ʲ��, ʵ�ʽ��, �ϴβ���
        Into n_��������, n_ʵ������, n_�����, n_�����, v_�ϴβ���
        From ҩƷ���
        Where ҩƷid = v_���.ԭ�ϲ���id And ���� = 1 And �ⷿid = �Է�����id_In;
      Exception
        When Others Then
          n_�������� := 0;
          n_ʵ������ := 0;
          n_����� := 0;
          n_����� := 0;
      End;
      If Nvl(v_���.�Ƿ���, 0) = 1 Then
        --ʵ��
        If Nvl(n_ʵ������, 0) > 0 Then
          n_�ۼ� := Nvl(n_�����, 0) / n_ʵ������;
        Else
          --�޿���:����ʾ
          v_Err_Msg := '[ZLSOFT]�õ����д���һ������ԭ�ϵ�ʵ����������[ZLSOFT]';
          Raise Err_Item;
        End If;
      Else
        --����,���ּ�Ϊ׼
        Begin
          Select Nvl(�ּ�, 0)
          Into n_�ۼ�
          From �շѼ�Ŀ
          Where �շ�ϸĿid = v_���.ԭ�ϲ���id And
                ((Sysdate Between ִ������ And ��ֹ����) Or (Sysdate >= ִ������ And ��ֹ���� Is Null));

        Exception
          When Others Then
            v_Err_Msg := 'Err';
        End;
        If v_Err_Msg = 'Err' Then
          v_Err_Msg := '[ZLSOFT]�õ����д���һ������ԭ�ϻ�δ���ж��ۣ�[ZLSOFT]';
          Raise Err_Item;
        End If;
      End If;
      n_���� := Nvl(v_����.ʵ������, 0) * v_���.���� / v_���.��ĸ;

      If n_���� = 0 Then
        v_Err_Msg := '[ZLSOFT]�õ����д���һ������ԭ�ϵ�����Ϊ���ˣ�[ZLSOFT]';
        Raise Err_Item;
      End If;
      n_���۽�� := n_���� * n_�ۼ�;

      --��ɱ���
      If Nvl(n_�����, 0) <= 0 Then
        If v_���ɱ����� = '1' And Nvl(v_���.�ɱ���, 0) > 0 Then
          n_�ɱ��� := v_���.�ɱ���;
          n_���   := n_���۽�� - n_���� * n_�ɱ���;
        Else
          n_���   := n_���۽�� * v_���.ָ������� / 100;
          n_�ɱ��� := (n_���۽�� - n_���) / n_����;
        End If;
      Else
        n_���   := n_���۽�� * (n_����� / n_�����);
        n_�ɱ��� := (n_���۽�� - n_���) / n_����;
      End If;
      n_�ɱ��� := Nvl(n_�ɱ���, 0);

      n_�ɱ���� := n_�ɱ��� * n_����;
      n_Max_��� := n_Max_��� + 1;

      Insert Into ҩƷ�շ���¼
        (ID, ��¼״̬, ����, NO, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ��д����, ʵ������,
         �ɱ���, �ɱ����, ���ۼ�, ���۽��, ���, ժҪ, ������, ��������, ����id, ����)
      Values
        (ҩƷ�շ���¼_Id.Nextval, 1, 16, No_In, n_Max_���, v_����.�Է�����id, v_����.�ⷿid, n_�������id, -1,
         v_���.ԭ�ϲ���id, v_�ϴβ���, n_����, n_����, n_�ɱ���, n_�ɱ����, n_�ۼ�, n_���۽��, n_���, v_����.ժҪ,
         v_����.������, v_����.��������, v_����.ҩƷid, v_����.���);

      --IF n_��������<0 then
      --    v_Err_Msg:='[ZLSOFT]�õ����д���һ������ԭ�ϵĿ�����������[ZLSOFT]';
      --    RAISE Err_Item;
      --End IF ;

      Update ҩƷ���
      Set �������� = Nvl(��������, 0) - n_����
      Where �ⷿid = v_����.�Է�����id And ҩƷid = v_���.ԭ�ϲ���id And ���� = 1;

      If Sql%NotFound Then
        Insert Into ҩƷ���
          (�ⷿid, ҩƷid, ����, ��������)
        Values
          (v_����.�Է�����id, v_���.ԭ�ϲ���id, 1, -n_����);
      End If;

      Delete From ҩƷ���
      Where �ⷿid = v_����.�Է�����id And ҩƷid = v_���.ԭ�ϲ���id And Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And
            Nvl(ʵ�ʽ��, 0) = 0 And Nvl(ʵ�ʲ��, 0) = 0;

    End Loop;
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��������ԭ�ϳ���_Insert;
/

--117925:����,2017-12-07,��������������
Create Or Replace Procedure Zl_���Ʋ������_Insert
(
  No_In         In ҩƷ�շ���¼.No%Type,
  ���_In       In ҩƷ�շ���¼.���%Type,
  �ⷿid_In     In ҩƷ�շ���¼.�ⷿid%Type,
  �Է�����id_In In ҩƷ�շ���¼.�Է�����id%Type,
  ����id_In     In ҩƷ�շ���¼.ҩƷid%Type,
  ʵ������_In   In ҩƷ�շ���¼.ʵ������%Type,
  ���ۼ�_In     In ҩƷ�շ���¼.���ۼ�%Type,
  ���۽��_In   In ҩƷ�շ���¼.���۽��%Type,
  ������_In     In ҩƷ�շ���¼.������%Type,
  ����_In       In ҩƷ�շ���¼.����%Type := Null,
  Ч��_In       In ҩƷ�շ���¼.Ч��%Type := Null,
  �������_In   In ҩƷ�շ���¼.�������%Type := Null,
  ���Ч��_In   In ҩƷ�շ���¼.���Ч��%Type := Null,
  ժҪ_In       In ҩƷ�շ���¼.ժҪ%Type := Null,
  ��������_In   In ҩƷ�շ���¼.��������%Type := Null,
  ��¼��_In     In Integer := 0
) Is
  Err_Item Exception;
  v_Err_Msg Varchar2(100);

  n_Id         ҩƷ�շ���¼.Id%Type; --�շ�ID
  n_�������id ҩƷ�շ���¼.������id%Type; --������ID
  n_������id ҩƷ�շ���¼.������id%Type; --������ID
  n_ʣ������   ҩƷ���.��������%Type;
  n_��ǰ����   ҩƷ�շ���¼.ʵ������%Type;
  n_�ۼ�       ҩƷ�շ���¼.���ۼ�%Type;
  n_�ּ�       �շѼ�Ŀ.�ּ�%Type;
  n_���۽��   ҩƷ�շ���¼.���۽��%Type;
  n_�ɱ���     ҩƷ�շ���¼.�ɱ���%Type;
  n_�ɱ����   ҩƷ�շ���¼.�ɱ����%Type;
  n_�ܳ���ɱ� ҩƷ�շ���¼.�ɱ����%Type;
  n_���       ҩƷ�շ���¼.���%Type;
  n_�����     ҩƷ�շ���¼.���%Type;
  n_������id ҩƷ�շ���¼.Id%Type;
  n_ʵ������   �շ���ĿĿ¼.�Ƿ���%Type;
  n_�ⷿ����   Integer; --�Ƿ��������   1:����;0��������
  n_���÷���   Integer; --���÷���
  n_����       ҩƷ�շ���¼.����%Type := Null; --����
  v_���ɱ����� Zlparameters.����ֵ%Type;
Begin
  -------------------------------------------------------------------------------------------
  --1.�ȴ���ԭ�ϳ��ⲿ��
  Select b.Id
  Into n_�������id
  From ҩƷ�������� A, ҩƷ������ B
  Where a.���id = b.Id And a.���� = 31 And b.ϵ�� = -1 And Rownum < 2;

  Select Zl_Getsysparameter(120) Into v_���ɱ����� From Dual;

  Select Max(���) Into n_����� From ҩƷ�շ���¼ Where NO = No_In And ���� = 16 And ���ϵ�� = -1;
  If Nvl(n_�����, 0) < ��¼��_In Then
    n_����� := ��¼��_In;
  End If;
  n_�ܳ���ɱ� := 0;
  For v_��� In (Select a.*, b.�Ƿ���, c.ָ�������, c.�ɱ���, c.���÷���
               From ���Ʋ��Ϲ��� A, �շ���ĿĿ¼ B, �������� C
               Where a.ԭ�ϲ���id = b.Id And a.���Ʋ���id = ����id_In And a.ԭ�ϲ���id = c.����id
			   Order By a.ԭ�ϲ���id) Loop
    n_ʣ������ := Round(ʵ������_In * v_���.���� / v_���.��ĸ, 7);
  
    If n_ʣ������ = 0 Then
      v_Err_Msg := '[ZLSOFT]�õ����д���һ������ԭ�ϵ�����Ϊ���ˣ�[ZLSOFT]';
      Raise Err_Item;
    End If;
  
    If Nvl(v_���.�Ƿ���, 0) = 0 Then
      --���۴���
      Begin
        Select Nvl(�ּ�, 0)
        Into n_�ּ�
        From �շѼ�Ŀ
        Where �շ�ϸĿid = v_���.ԭ�ϲ���id And ((Sysdate Between ִ������ And ��ֹ����) Or (Sysdate >= ִ������ And ��ֹ���� Is Null));
      Exception
        When Others Then
          v_Err_Msg := 'Err';
      End;
      If Nvl(v_Err_Msg, ' ') = 'Err' Then
        v_Err_Msg := '[ZLSOFT]�õ����д���һ������ԭ�ϻ�δ���ж��ۣ�[ZLSOFT]';
        Raise Err_Item;
      End If;
    Else
      n_�ּ� := 0;
    End If;
    n_������id := -1;
    --���Ƚ��ȳ���ԭ�����
    For v_��� In (Select Nvl(����, 0) As ����, Max(���ۼ�) As ���ۼ�, Sum(Nvl(��������, 0)) As ��������, Sum(Nvl(ʵ������, 0)) As ʵ������,
                        Sum(Nvl(ʵ�ʲ��, 0)) As ʵ�ʲ��, Sum(Nvl(ʵ�ʽ��, 0)) As ʵ�ʽ��, Max(�ϴβ���) As �ϴβ���, Max(�ϴ�����) As �ϴ�����,
                        Max(�ϴ���������) As �ϴ���������, Max(Ч��) As Ч��, Max(���Ч��) As ���Ч��, Max(��׼�ĺ�) As ��׼�ĺ�
                 From ҩƷ���
                 Where ҩƷid = v_���.ԭ�ϲ���id And ���� = 1 And �ⷿid = �Է�����id_In
                 Group By Nvl(����, 0)
                 Order By Nvl(����, 0)) Loop
    
      Select �Ƿ��� Into n_ʵ������ From �շ���ĿĿ¼ Where ID = v_���.ԭ�ϲ���id;
    
      If Nvl(v_���.�Ƿ���, 0) = 1 Then
        --ʵ�۴���
        If Nvl(v_���.ʵ������, 0) > 0 Then
          If Nvl(v_���.����, 0) <> 0 And Nvl(v_���.���ۼ�, 0) <> 0 Then
            --����ʵ�ۣ������������ۼۣ���ֻ�������ۼ�Ϊ׼.
            n_�ۼ� := Nvl(v_���.���ۼ�, 0);
          Else
            n_�ۼ� := Nvl(v_���.ʵ�ʽ��, 0) / v_���.ʵ������;
          End If;
        Else
          --�޿���:����ʾ
          v_Err_Msg := '[ZLSOFT]�õ����д���һ������ԭ�ϵ�ʵ����������[ZLSOFT]';
          Raise Err_Item;
        End If;
      Else
        n_�ۼ� := n_�ּ�;
      End If;
      If Nvl(v_���.��������, 0) >= n_ʣ������ Then
        n_��ǰ���� := n_ʣ������;
      Else
        n_��ǰ���� := Nvl(v_���.��������, 0);
      End If;
      n_���۽�� := Round(n_��ǰ���� * n_�ۼ�, 7);
    
      --��ɱ���
      If Nvl(v_���.ʵ�ʽ��, 0) <= 0 Then
        If v_���ɱ����� = '1' And Nvl(v_���.�ɱ���, 0) > 0 Then
          n_�ɱ��� := v_���.�ɱ���;
          n_���   := Round(n_���۽�� - n_��ǰ���� * n_�ɱ���, 7);
        Else
          n_���   := n_���۽�� * v_���.ָ������� / 100;
          n_�ɱ��� := (n_���۽�� - n_���) / n_��ǰ����;
        End If;
      Else
        n_���   := n_���۽�� * (v_���.ʵ�ʲ�� / v_���.ʵ�ʽ��);
        n_�ɱ��� := (n_���۽�� - n_���) / n_��ǰ����;
      End If;
      n_�ɱ���     := Nvl(n_�ɱ���, 0);
      n_�ɱ����   := n_�ɱ��� * n_��ǰ����;
      n_�ܳ���ɱ� := n_�ܳ���ɱ� + n_�ɱ����;
    
      n_����� := n_����� + 1;
      Select ҩƷ�շ���¼_Id.Nextval Into n_������id From Dual;
    
      Insert Into ҩƷ�շ���¼
        (ID, ��¼״̬, ����, NO, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ����, Ч��, ���Ч��, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ���۽��, ���,
         ժҪ, ������, ��������, ����id, ����)
      Values
        (n_������id, 1, 16, No_In, n_�����, �Է�����id_In, �ⷿid_In, n_�������id, -1, v_���.ԭ�ϲ���id,
         Decode(v_���.����, 0, Null, v_���.����), v_���.�ϴ�����, v_���.Ч��, v_���.���Ч��, Nvl(n_��ǰ����, 0), n_��ǰ����, n_�ɱ���, n_�ɱ����, n_�ۼ�,
         n_���۽��, n_���, ժҪ_In, ������_In, ��������_In, ����id_In, ���_In);
    
      Update ҩƷ���
      Set �������� = Nvl(��������, 0) - n_��ǰ����
      Where �ⷿid = �Է�����id_In And ҩƷid = v_���.ԭ�ϲ���id And Nvl(����, 0) = Nvl(v_���.����, 0) And ���� = 1;
    
      If Sql%NotFound Then
        Insert Into ҩƷ���
          (�ⷿid, ҩƷid, ����, ����, ��������, ���ۼ�)
        Values
          (�Է�����id_In, v_���.ԭ�ϲ���id, Decode(Nvl(v_���.����, 0), 0, Null, v_���.����), 1, -n_��ǰ����,
           Decode(n_ʵ������, 1, Decode(Nvl(v_���.����, 0), 0, Null, n_�ۼ�), Null));
      End If;
    
      Delete From ҩƷ���
      Where �ⷿid = �Է�����id_In And ҩƷid = v_���.ԭ�ϲ���id And Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0 And
            Nvl(ʵ�ʲ��, 0) = 0;
      n_ʣ������ := n_ʣ������ - Nvl(v_���.��������, 0);
      If n_ʣ������ <= 0 Then
        Exit;
      End If;
    End Loop;
  
    If n_ʣ������ > 0 Then
      --�ȿ��������,��Ҫ��ʣ����������
      If Nvl(v_���.�Ƿ���, 0) = 1 Or Nvl(v_���.���÷���, 0) = 1 Then
        --ʵ�ۻ����÷�������Ҫ�п��
        v_Err_Msg := '[ZLSOFT]�õ����д���һ������ԭ�ϵĿ����������㣬����[ZLSOFT]';
        Raise Err_Item;
      End If;
    
      If n_������id = -1 Then
        --��ʾ����û�п�棬��Ҫ������ص�����
        n_�ۼ�     := n_�ּ�;
        n_���۽�� := Round(n_ʣ������ * n_�ۼ�, 7);
        If v_���ɱ����� = '1' And Nvl(v_���.�ɱ���, 0) > 0 Then
          n_�ɱ��� := v_���.�ɱ���;
          n_���   := Round(n_���۽�� - n_ʣ������ * n_�ɱ���, 7);
        Else
          n_���   := n_���۽�� * v_���.ָ������� / 100;
          n_�ɱ��� := (n_���۽�� - n_���) / n_ʣ������;
        End If;
        n_�ɱ���     := Nvl(n_�ɱ���, 0);
        n_�ɱ����   := n_�ɱ��� * n_ʣ������;
        n_�ܳ���ɱ� := n_�ܳ���ɱ� + n_�ɱ����;
      
        n_����� := n_����� + 1;
        Select ҩƷ�շ���¼_Id.Nextval Into n_������id From Dual;
      
        Insert Into ҩƷ�շ���¼
          (ID, ��¼״̬, ����, NO, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ����, Ч��, ���Ч��, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ���۽��,
           ���, ժҪ, ������, ��������, ����id, ����)
        Values
          (n_������id, 1, 16, No_In, n_�����, �Է�����id_In, �ⷿid_In, n_�������id, -1, v_���.ԭ�ϲ���id, Null, Null, Null, Null, n_ʣ������,
           n_ʣ������, n_�ɱ���, n_�ɱ����, n_�ۼ�, n_���۽��, n_���, ժҪ_In, ������_In, ��������_In, ����id_In, ���_In);
      Else
        --������ʣ��
        Select �ɱ���, ���ۼ�, ���۽��, ���, ��д����, �ɱ����
        Into n_�ɱ���, n_�ۼ�, n_���۽��, n_���, n_��ǰ����, n_�ɱ����
        From ҩƷ�շ���¼
        Where ID = n_������id;
      
        Update ҩƷ�շ���¼
        Set ��д���� = Nvl(��д����, 0) + n_ʣ������, ʵ������ = Nvl(ʵ������, 0) + n_ʣ������, �ɱ��� = Nvl(n_�ɱ���, 0),
            �ɱ���� = Nvl(n_�ɱ���, 0) * (n_��ǰ���� + n_ʣ������), ���ۼ� = n_�ۼ�, ���۽�� = n_�ۼ� * (n_��ǰ���� + n_ʣ������),
            ��� = Round((n_�ۼ� * (n_��ǰ���� + n_ʣ������)) - (Nvl(n_�ɱ���, 0) * (n_��ǰ���� + n_ʣ������)), 7)
        Where ID = n_������id;
        n_�ܳ���ɱ� := (n_�ܳ���ɱ� - n_�ɱ����) + Nvl(n_�ɱ���, 0) * (n_��ǰ���� + n_ʣ������);
      
      End If;
      Update ҩƷ���
      Set �������� = Nvl(��������, 0) - n_ʣ������
      Where �ⷿid = �Է�����id_In And ҩƷid = v_���.ԭ�ϲ���id And Nvl(����, 0) = 0 And ���� = 1;
    
      If Sql%NotFound Then
        Insert Into ҩƷ���
          (�ⷿid, ҩƷid, ����, ��������, ���ۼ�)
        Values
          (�Է�����id_In, v_���.ԭ�ϲ���id, 1, -n_ʣ������, Null);
      End If;
    
      Delete From ҩƷ���
      Where �ⷿid = �Է�����id_In And ҩƷid = v_���.ԭ�ϲ���id And Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0 And
            Nvl(ʵ�ʲ��, 0) = 0;
    End If;
  End Loop;

  n_�ɱ��� := n_�ܳ���ɱ� / ʵ������_In;
  n_���   := ���۽��_In - n_�ܳ���ɱ�;

  Select ҩƷ�շ���¼_Id.Nextval Into n_Id From Dual;

  --ȷ���Ƿ����  
  Select Nvl(�ⷿ����, 0), Nvl(���÷���, 0) Into n_�ⷿ����, n_���÷��� From �������� Where ����id = ����id_In;

  If n_���÷��� = 0 Then
    If n_�ⷿ���� = 1 Then
      Begin
        Select Distinct 0
        Into n_�ⷿ����
        From ��������˵��
        Where (�������� = '���ϲ���' Or �������� Like '�Ƽ���') And ����id = �ⷿid_In;
      Exception
        When Others Then
          n_�ⷿ���� := 1;
      End;
    
      If n_�ⷿ���� = 1 Then
        n_���� := n_Id;
      End If;
    End If;
  Else
    n_���� := n_Id;
  End If;

  Select b.Id
  Into n_������id
  From ҩƷ�������� A, ҩƷ������ B
  Where a.���id = b.Id And a.���� = 31 And b.ϵ�� = 1 And Rownum < 2;

  Insert Into ҩƷ�շ���¼
    (ID, ��¼״̬, ����, NO, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ����, Ч��, �������, ���Ч��, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ���۽��,
     ���, ժҪ, ������, ��������)
  Values
    (n_Id, 1, 16, No_In, ���_In, �ⷿid_In, �Է�����id_In, n_������id, 1, ����id_In, n_����, ����_In, Ч��_In, �������_In, ���Ч��_In, ʵ������_In,
     ʵ������_In, n_�ɱ���, n_�ܳ���ɱ�, ���ۼ�_In, ���۽��_In, n_���, ժҪ_In, ������_In, ��������_In);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���Ʋ������_Insert;
/

--117925:����,2017-12-07,��������������
Create Or Replace Procedure Zl_���Ʋ������_Strike
(
  No_In     In ҩƷ�շ���¼.No%Type,
  �����_In In ҩƷ�շ���¼.�����%Type
) Is
  Err_Item Exception;
  v_Err_Msg    Varchar2(500);
  n_ʵ������   �շ���ĿĿ¼.�Ƿ���%Type;
  n_ƽ���ɱ��� ҩƷ���.ƽ���ɱ���%Type;

Begin
  Update ҩƷ�շ���¼ Set ��¼״̬ = 3 Where NO = No_In And ���� = 16 And ��¼״̬ = 1;

  If Sql%RowCount = 0 Then
    v_Err_Msg := '[ZLSOFT]�õ����Ѿ������˳�����[ZLSOFT]';
    Raise Err_Item;
  End If;

  Insert Into ҩƷ�շ���¼
    (ID, ��¼״̬, ����, NO, ���, �ⷿid, ������id, �Է�����id, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ���۽��, ���, ժҪ,
     ������, ��������, �����, �������, ����id, ����)
    Select ҩƷ�շ���¼_Id.Nextval, 2, 16, No_In, ���, �ⷿid, ������id, �Է�����id, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, -��д����, -ʵ������, �ɱ���,
           -�ɱ����, ���ۼ�, -���۽��, -���, ժҪ, �����_In, Sysdate, �����_In, Sysdate, ����id, ����
    From ҩƷ�շ���¼
    Where NO = No_In And ���� = 16 And ��¼״̬ = 3;

  For c_���� In (Select ID, �ⷿid, ������id, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ���ۼ�, ��д����, ʵ������, �ɱ���, ���۽��, ���
               From ҩƷ�շ���¼ A
               Where NO = No_In And ���� = 16 And ��¼״̬ = 2
			   Order By ҩƷid,����) Loop
  
    Select �Ƿ��� Into n_ʵ������ From �շ���ĿĿ¼ Where ID = c_����.ҩƷid;
  
    --���Ĳ��Ͽ������Ӧ����
    --���Ʋ�����ԭ�ϲ��ϵĴ���ͨ�����ϵ����ʵ��
    Update ҩƷ���
    Set �������� = Nvl(��������, 0) + Nvl(c_����.��д����, 0) * c_����.���ϵ��, ʵ������ = Nvl(ʵ������, 0) + Nvl(c_����.��д����, 0) * c_����.���ϵ��,
        ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) + Nvl(c_����.���۽��, 0) * c_����.���ϵ��, ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + Nvl(c_����.���, 0) * c_����.���ϵ��,
        �ϴβɹ��� = Nvl(c_����.�ɱ���, �ϴβɹ���), �ϴ����� = Nvl(c_����.����, �ϴ�����), �ϴβ��� = Nvl(c_����.����, �ϴβ���), Ч�� = Nvl(c_����.Ч��, Ч��),
        ���ۼ� = Decode(n_ʵ������, 1, Decode(Nvl(c_����.����, 0), 0, Null, Decode(Nvl(���ۼ�, 0), 0, c_����.���ۼ�, ���ۼ�)), Null)
    Where �ⷿid = c_����.�ⷿid And ҩƷid = c_����.ҩƷid And Nvl(����, 0) = Nvl(c_����.����, 0) And ���� = 1;
  
    If Sql%NotFound Then
      Insert Into ҩƷ���
        (�ⷿid, ҩƷid, ����, ����, ��������, ʵ������, ʵ�ʽ��, ʵ�ʲ��, �ϴβɹ���, �ϴ�����, �ϴβ���, Ч��, ���ۼ�, ƽ���ɱ���)
      Values
        (c_����.�ⷿid, c_����.ҩƷid, c_����.����, 1, c_����.��д���� * c_����.���ϵ��, c_����.��д���� * c_����.���ϵ��, c_����.���۽�� * c_����.���ϵ��,
         c_����.��� * c_����.���ϵ��, c_����.�ɱ���, c_����.����, c_����.����, c_����.Ч��,
         Decode(n_ʵ������, 1, Decode(Nvl(c_����.����, 0), 0, Null, c_����.���ۼ�), Null), c_����.�ɱ���);
    End If;
  
    Delete From ҩƷ���
    Where �ⷿid = c_����.�ⷿid And ҩƷid = c_����.ҩƷid And Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0 And
          Nvl(ʵ�ʲ��, 0) = 0;
  
    --���¼�������е�ƽ���ɱ���
    Update ҩƷ���
    Set ƽ���ɱ��� = Decode(Nvl(����, 0), 0, Decode((ʵ�ʽ�� - ʵ�ʲ��) / ʵ������, 0, �ϴβɹ���, (ʵ�ʽ�� - ʵ�ʲ��) / ʵ������), �ϴβɹ���)
    Where ���� = 1 And �ⷿid = c_����.�ⷿid And ҩƷid = c_����.ҩƷid And Nvl(����, 0) = Nvl(c_����.����, 0) And Nvl(ʵ������, 0) <> 0;
    If Sql%NotFound Then
      Select �ɱ��� Into n_ƽ���ɱ��� From �������� Where ����id = c_����.ҩƷid;
      Update ҩƷ���
      Set ƽ���ɱ��� = n_ƽ���ɱ���
      Where ���� = 1 And �ⷿid = c_����.�ⷿid And ҩƷid = c_����.ҩƷid And Nvl(����, 0) = Nvl(c_����.����, 0) And
            Nvl(ƽ���ɱ���, 0) <> c_����.�ɱ���;
    End If;
    Zl_�����շ���¼_��������(c_����.Id);
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���Ʋ������_Strike;
/

--117925:����,2017-12-07,��������������
Create Or Replace Procedure Zl_���Ʋ������_Verify
(
  No_In     In ҩƷ�շ���¼.No%Type := Null,
  �����_In In ҩƷ�շ���¼.�����%Type := Null
) Is
  Err_Item Exception;
  v_Err_Msg Varchar2(100);

  v_���ɱ����� Zlparameters.����ֵ%Type;

  n_ʵ�ʿ���� ҩƷ���.ʵ�ʽ��%Type;
  n_ʵ�ʿ���� ҩƷ���.ʵ�ʲ��%Type;
  n_ʵ�ʿ������ ҩƷ���.ʵ������%Type;
  n_������     ҩƷ���.ʵ�ʲ��%Type;
  n_ʵ������     �շ���ĿĿ¼.�Ƿ���%Type;

  n_�ɱ���     ҩƷ�շ���¼.�ɱ���%Type;
  n_�ɱ����   ҩƷ�շ���¼.�ɱ����%Type;
  n_�����     Number(18, 8);
  n_С��       Number(2);
  n_ƽ���ɱ��� ҩƷ���.ƽ���ɱ���%Type;

Begin
  Select zl_GetSysParameter(120) Into v_���ɱ����� From Dual;

  Update ҩƷ�շ���¼
  Set ����� = Nvl(�����_In, �����), ������� = Sysdate
  Where NO = No_In And ���� = 16 And ��¼״̬ = 1 And ����� Is Null;

  If Sql%RowCount = 0 Then
    v_Err_Msg := '[ZLSOFT]�õ����Ѿ���������ˣ�[ZLSOFT]';
    Raise Err_Item;
  End If;

  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_С�� From Dual;

  Update ҩƷ�շ���¼ Set �ɱ���� = 0 Where NO = No_In And ���� = 16 And ��¼״̬ = 1 And ���ϵ�� = 1;

  For v_ԭ�� In (Select ID, ʵ������, ���ۼ�, ���۽��, ���, �ⷿid, ҩƷid, ����, �ɱ���, ����, Ч��, ����, ���Ч��, ��׼�ĺ�, ��������, ������id, ���ϵ��, �Է�����id,
                      ����id As ���Ʋ���id, Trunc(����) As ���
               From ҩƷ�շ���¼
               Where NO = No_In And ���� = 16 And ��¼״̬ = 1 And ���ϵ�� = -1
               Order By ҩƷid, ����) Loop
  
    Select �Ƿ��� Into n_ʵ������ From �շ���ĿĿ¼ Where ID = v_ԭ��.ҩƷid;
  
    Begin
      Select Nvl(ʵ�ʽ��, 0), Nvl(ʵ�ʲ��, 0), Nvl(ʵ������, 0)
      Into n_ʵ�ʿ����, n_ʵ�ʿ����, n_ʵ�ʿ������
      From ҩƷ���
      Where ҩƷid = v_ԭ��.ҩƷid And Nvl(����, 0) = Nvl(v_ԭ��.����, 0) And �ⷿid = v_ԭ��.�ⷿid And ���� = 1 And Rownum = 1;
    Exception
      When Others Then
        n_ʵ�ʿ���� := 0;
        n_ʵ�ʿ������ := 0;
    End;
  
    If n_ʵ�ʿ���� <= 0 Then
      If (n_ʵ�ʿ���� - n_ʵ�ʿ����) <= 0 Or n_ʵ�ʿ������ <= 0 Then
      
        Begin
          Select ָ������� / 100 Into n_����� From �������� Where ����id = v_ԭ��.ҩƷid;
        Exception
          When Others Then
            n_����� := 0;
        End;
        If v_���ɱ����� = '1' Then
          Begin
            Select Nvl(�ɱ���, 0) Into n_�ɱ��� From �������� Where ����id = v_ԭ��.ҩƷid;
          Exception
            When Others Then
              n_�ɱ��� := 0;
          End;
          If n_�ɱ��� = 0 Then
            n_������ := Round(v_ԭ��.���۽�� * n_�����, 4);
          Else
            n_������ := Round(v_ԭ��.���۽�� - v_ԭ��.ʵ������ * n_�ɱ���, 4);
          End If;
        Else
          n_������ := Round(v_ԭ��.���۽�� * n_�����, n_С��);
        End If;
      Else
        --��Ҫ�������ۼ�Ϊ���������Ӷ�����޳ɱ�������
        n_�ɱ���   := ((n_ʵ�ʿ���� - n_ʵ�ʿ����) / n_ʵ�ʿ������);
        n_������ := Round(v_ԭ��.���۽�� - n_�ɱ��� * v_ԭ��.ʵ������, n_С��);
      End If;
    Else
      n_�����   := n_ʵ�ʿ���� / n_ʵ�ʿ����;
      n_������ := Round(v_ԭ��.���۽�� * n_�����, n_С��);
    End If;
  
    If Nvl(v_ԭ��.ʵ������, 0) = 0 Then
      n_�ɱ��� := (v_ԭ��.���۽�� - n_������);
    Else
      n_�ɱ��� := (v_ԭ��.���۽�� - n_������) / v_ԭ��.ʵ������;
    End If;
    n_�ɱ���   := Nvl(n_�ɱ���, 0);
    n_�ɱ���� := Round(n_�ɱ��� * v_ԭ��.ʵ������, n_С��);
  
    Update ҩƷ�շ���¼ Set �ɱ��� = n_�ɱ���, �ɱ���� = n_�ɱ����, ��� = n_������ Where ID = v_ԭ��.Id;
  
    Update ҩƷ���
    Set ʵ������ = Nvl(ʵ������, 0) - Nvl(v_ԭ��.ʵ������, 0), ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) - Nvl(v_ԭ��.���۽��, 0), ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) - n_������
    Where �ⷿid = v_ԭ��.�ⷿid And ҩƷid = v_ԭ��.ҩƷid And Nvl(����, 0) = Nvl(v_ԭ��.����, 0) And ���� = 1;
  
    If Sql%NotFound Then
      Insert Into ҩƷ���
        (�ⷿid, ҩƷid, ����, ����, ��������, ʵ������, ʵ�ʽ��, ʵ�ʲ��, �ϴ�����, �ϴ���������, Ч��, �ϴβ���, ���Ч��, ��׼�ĺ�, �ϴβɹ���, ���ۼ�,ƽ���ɱ���)
      Values
        (v_ԭ��.�ⷿid, v_ԭ��.ҩƷid, Decode(v_ԭ��.����, Null, Null, 0, Null, v_ԭ��.����), 1,
         Decode(v_ԭ��.���ϵ��, 1, Nvl(v_ԭ��.ʵ������, 0), 0), v_ԭ��.ʵ������ * v_ԭ��.���ϵ��, v_ԭ��.���۽�� * v_ԭ��.���ϵ��, n_������ * v_ԭ��.���ϵ��,
         v_ԭ��.����, v_ԭ��.��������, v_ԭ��.Ч��, v_ԭ��.����, v_ԭ��.���Ч��, v_ԭ��.��׼�ĺ�, n_�ɱ���,
         Decode(n_ʵ������, 1, Decode(Nvl(v_ԭ��.����, 0), 0, Null, v_ԭ��.���ۼ�), Null),n_�ɱ���);
    End If;
    Delete From ҩƷ���
    Where �ⷿid = v_ԭ��.�ⷿid And ҩƷid = v_ԭ��.ҩƷid And Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0 And
          Nvl(ʵ�ʲ��, 0) = 0;
  
    Update ҩƷ�շ���¼
    Set �ɱ���� = Nvl(�ɱ����, 0) + n_�ɱ����
    Where NO = No_In And ��� = v_ԭ��.��� And ҩƷid = Nvl(v_ԭ��.���Ʋ���id, 0) And ���� = 16 And ��¼״̬ = 1 And ���ϵ�� = 1;
  End Loop;

  For v_���Ʋ��� In (Select ID, �ɱ����, ���ۼ�, ʵ������, ���۽��, ���, �ⷿid, ҩƷid, ����, �ɱ���, ����, Ч��, ����, ���Ч��, ��׼�ĺ�, ��������, ������id, ���ϵ��,
                        �Է�����id, ����id As ���Ʋ���id, Trunc(����) As ���
                 From ҩƷ�շ���¼
                 Where NO = No_In And ���� = 16 And ��¼״̬ = 1 And ���ϵ�� = 1
                 Order By ҩƷid, ����) Loop
    Select �Ƿ��� Into n_ʵ������ From �շ���ĿĿ¼ Where ID = v_���Ʋ���.ҩƷid;
  
    n_�ɱ���� := Nvl(v_���Ʋ���.�ɱ����, 0);
    If Nvl(v_���Ʋ���.ʵ������, 0) <> 0 Then
      n_�ɱ��� := n_�ɱ���� / Nvl(v_���Ʋ���.ʵ������, 0);
    Else
      n_�ɱ��� := n_�ɱ����;
    End If;
    n_������ := v_���Ʋ���.���۽�� - n_�ɱ����;
  
    Update ҩƷ�շ���¼ Set �ɱ��� = n_�ɱ���, �ɱ���� = n_�ɱ����, ��� = n_������ Where ID = v_���Ʋ���.Id;
    Update ҩƷ���
    Set �������� = Nvl(��������, 0) + Decode(v_���Ʋ���.���ϵ��, 1, Nvl(v_���Ʋ���.ʵ������, 0), 0), ʵ������ = Nvl(ʵ������, 0) + Nvl(v_���Ʋ���.ʵ������, 0),
        ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) + Nvl(v_���Ʋ���.���۽��, 0), ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + n_������,
        ���ۼ� = Decode(n_ʵ������, 1, Decode(Nvl(v_���Ʋ���.����, 0), 0, Null, Decode(Nvl(���ۼ�, 0), 0, v_���Ʋ���.���ۼ�, ���ۼ�)), Null)
    Where �ⷿid = v_���Ʋ���.�ⷿid And ҩƷid = v_���Ʋ���.ҩƷid And Nvl(����, 0) = Nvl(v_���Ʋ���.����, 0) And ���� = 1;
  
    If Sql%NotFound Then
      Insert Into ҩƷ���
        (�ⷿid, ҩƷid, ����, ����, ��������, ʵ������, ʵ�ʽ��, ʵ�ʲ��, �ϴ�����, �ϴ���������, Ч��, �ϴβ���, ���Ч��, ��׼�ĺ�, �ϴβɹ���, ���ۼ�,ƽ���ɱ���)
      Values
        (v_���Ʋ���.�ⷿid, v_���Ʋ���.ҩƷid, Decode(Nvl(v_���Ʋ���.����, 0), 0, Null, v_���Ʋ���.����), 1,
         Decode(v_���Ʋ���.���ϵ��, 1, Nvl(v_���Ʋ���.ʵ������, 0), 0), v_���Ʋ���.ʵ������ * v_���Ʋ���.���ϵ��, v_���Ʋ���.���۽�� * v_���Ʋ���.���ϵ��,
         n_������ * v_���Ʋ���.���ϵ��, v_���Ʋ���.����, v_���Ʋ���.��������, v_���Ʋ���.Ч��, v_���Ʋ���.����, v_���Ʋ���.���Ч��, v_���Ʋ���.��׼�ĺ�, n_�ɱ���,
         Decode(n_ʵ������, 1, Decode(Nvl(v_���Ʋ���.����, 0), 0, Null, v_���Ʋ���.���ۼ�), Null),n_�ɱ���);
    End If;
    Delete From ҩƷ���
    Where �ⷿid = v_���Ʋ���.�ⷿid And ҩƷid = v_���Ʋ���.ҩƷid And Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0 And
          Nvl(ʵ�ʲ��, 0) = 0;
  
    --���¸ò��ϵĳɱ���
    Update �������� Set �ɱ��� = n_�ɱ��� Where ����id = v_���Ʋ���.ҩƷid;
  
    --���¼�������е�ƽ���ɱ���
    Update ҩƷ���
    Set ƽ���ɱ��� = Decode(Nvl(����, 0), 0, decode((ʵ�ʽ�� - ʵ�ʲ��) / ʵ������,0,�ϴβɹ���,(ʵ�ʽ�� - ʵ�ʲ��) / ʵ������), �ϴβɹ���)
    Where ҩƷid = v_���Ʋ���.ҩƷid And Nvl(����, 0) = Nvl(v_���Ʋ���.����, 0) And �ⷿid = v_���Ʋ���.�ⷿid And ���� = 1 And Nvl(ʵ������, 0) <> 0;
    If Sql%NotFound Then
      Select �ɱ��� Into n_ƽ���ɱ��� From �������� Where ����id = v_���Ʋ���.ҩƷid;
      Update ҩƷ���
      Set ƽ���ɱ��� = n_ƽ���ɱ���
      Where ҩƷid = v_���Ʋ���.ҩƷid And �ⷿid = v_���Ʋ���.�ⷿid And Nvl(����, 0) = Nvl(v_���Ʋ���.����, 0) and ����=1;
    End If;
  
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���Ʋ������_Verify;
/

--117925:����,2017-12-07,��������������
Create Or Replace Procedure Zl_�����⹺_Verify
(
  No_In       In ҩƷ�շ���¼.No%Type := Null,
  �����_In   In ҩƷ�շ���¼.�����%Type := Null,
  �������_In In ҩƷ�շ���¼.�������%Type := Sysdate
) Is
  Err_Item Exception;
  v_Err_Msg Varchar2(500);

  n_��λid      ҩƷ�շ���¼.��ҩ��λid%Type;
  n_��Ʊ���    Ӧ����¼.��Ʊ���%Type;
  n_�����    ҩƷ���.ʵ�ʽ��%Type;
  n_�����    ҩƷ���.ʵ�ʲ��%Type;
  n_�������    ҩƷ���.ʵ������%Type;
  n_ʵ������    �շ���ĿĿ¼.�Ƿ���%Type;
  n_Batch_Count Integer; --ԭ���������ڷ����Ĳ��ϵ�����
  v_����ǰ׺    Varchar2(20);
  v_�ڲ�����    ҩƷ���.�ڲ�����%Type;
  v_�ƿ�no      ҩƷ�շ���¼.No%Type;
  v_�Է��ⷿid  ҩƷ�շ���¼.�ⷿid%Type := 0;
  v_�����id    ҩƷ�շ���¼.������id%Type := 0;
  v_�����id    ҩƷ�շ���¼.������id%Type := 0;
  n_ƽ���ɱ���  ҩƷ���.ƽ���ɱ���%Type;
  n_��������    ҩƷ�շ���¼.ʵ������%Type;
Begin
  v_����ǰ׺ := Nvl(zl_GetSysParameter(159), '');

  Update ҩƷ�շ���¼
  Set ����� = Nvl(�����_In, �����), ������� = �������_In
  Where NO = No_In And ���� = 15 And ��¼״̬ = 1 And ����� Is Null;

  If Sql%RowCount = 0 Then
    v_Err_Msg := '[ZLSOFT]�õ����Ѿ���������˻�ɾ�������ܽ�����ˣ�[ZLSOFT]';
    Raise Err_Item;
  End If;
  --��Ҫ���ԭ���������ڷ����Ĳ��ϣ����ܶ������
  Select Count(*)
  Into n_Batch_Count
  From ҩƷ�շ���¼ A, �������� B
  Where a.ҩƷid = b.����id And a.No = No_In And a.���� = 15 And a.��¼״̬ = 1 And Nvl(a.����, 0) = 0 And
        ((Nvl(b.�ⷿ����, 0) = 1 And
        a.�ⷿid Not In (Select ����id From ��������˵�� Where (�������� Like '���ϲ���') Or (�������� Like '�Ƽ���'))) Or Nvl(b.���÷���, 0) = 1);

  If n_Batch_Count > 0 Then
    v_Err_Msg := '[ZLSOFT]�õ����а�����һ��ԭ�������������ڷ����Ĳ��ϣ�������ˣ�[ZLSOFT]';
    Raise Err_Item;
  End If;

  --ԭ�����ֲ������Ĳ���,�����ʱ��Ҫ������
  Update ҩƷ�շ���¼
  Set ���� = 0
  Where ID In
        (Select ID
         From ҩƷ�շ���¼ A, �������� B
         Where a.ҩƷid = b.����id And a.No = No_In And a.���� = 15 And a.��¼״̬ = 1 And Nvl(a.����, 0) > 0 And
               (Nvl(b.�ⷿ����, 0) = 0 Or
               (Nvl(b.���÷���, 0) = 0 And
               a.�ⷿid In (Select ����id From ��������˵�� Where (�������� Like '���ϲ���') Or (�������� Like '�Ƽ���')))));

  For v_�շ� In (Select a.Id, a.ʵ������, a.��ҩ��ʽ, a.���ۼ�, a.���۽��, a.���, a.�ⷿid, a.ҩƷid, a.����, a.��ҩ��λid, a.�ɱ���, a.����, a.Ч��,
                      a.���Ч��, a.��������, a.����, a.������id, a.ע��֤��, a.����, a.��Ʒ����, a.�ڲ�����, Nvl(b.�Ƿ��������, 0) As �������, a.��׼�ĺ�,
                      Nvl(a.����id, 0) As ����id, ���
               From ҩƷ�շ���¼ A, �������� B
               Where a.ҩƷid = b.����id And a.No = No_In And a.���� = 15 And a.��¼״̬ = 1
               Order By a.ҩƷid, a.����) Loop
    v_�ڲ����� := Null;
    If v_�շ�.������� = 1 Then
      If v_�շ�.�ڲ����� Is Null Then
        If Not v_����ǰ׺ Is Null Then
          v_�ڲ����� := v_����ǰ׺ || Nextno(126);
        Else
          v_�ڲ����� := Nextno(126);
        End If;
      Else
        v_�ڲ����� := v_�շ�.�ڲ�����;
      End If;
	  --���������ӡ��������
      Insert Into ���������ӡ��¼
        (NO, ����, �ⷿid, ����id, ���, ��Ʒ����, �ڲ�����, �������, ��ӡ����, ���ʱ��)
      Values
        (No_In, 15, v_�շ�.�ⷿid, v_�շ�.ҩƷid, v_�շ�.���, v_�շ�.��Ʒ����, v_�ڲ�����, v_�շ�.ʵ������, 0, �������_In);
    End If;
  
    --���Ĳ��Ͽ������Ӧ����
    Select �Ƿ��� Into n_ʵ������ From �շ���ĿĿ¼ Where ID = v_�շ�.ҩƷid;
  
    If v_�շ�.����id = 2 Then
      n_�������� := Nvl(v_�շ�.ʵ������, 0);
    Else
      If v_�շ�.��ҩ��ʽ = 1 Then
        n_�������� := 0;
      Else
        n_�������� := Nvl(v_�շ�.ʵ������, 0);
      End If;
    End If;
  
    Update ҩƷ���
    Set �������� = Nvl(��������, 0) + n_��������, ʵ������ = Nvl(ʵ������, 0) + Nvl(v_�շ�.ʵ������, 0), ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) + Nvl(v_�շ�.���۽��, 0),
        ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + Nvl(v_�շ�.���, 0), �ϴι�Ӧ��id = Nvl(v_�շ�.��ҩ��λid, �ϴι�Ӧ��id), �ϴβɹ��� = Nvl(v_�շ�.�ɱ���, �ϴβɹ���),
        �ϴ����� = Nvl(v_�շ�.����, �ϴ�����), �ϴβ��� = Nvl(v_�շ�.����, �ϴβ���), ���Ч�� = Nvl(v_�շ�.���Ч��, ���Ч��),
        �ϴ��������� = Nvl(v_�շ�.��������, �ϴ���������), Ч�� = Nvl(v_�շ�.Ч��, Ч��),
        ���ۼ� = Decode(Nvl(v_�շ�.����, 0), 0, Null, Decode(n_ʵ������, 1, v_�շ�.���ۼ�, Null)), �ϴο��� = Nvl(v_�շ�.����, �ϴο���),
        ��Ʒ���� = v_�շ�.��Ʒ����, �ڲ����� = v_�ڲ�����, ��׼�ĺ� = v_�շ�.��׼�ĺ�
    Where �ⷿid = v_�շ�.�ⷿid And ҩƷid = v_�շ�.ҩƷid And Nvl(����, 0) = Nvl(v_�շ�.����, 0) And ���� = 1;
  
    If Sql%NotFound Then
      Insert Into ҩƷ���
        (�ⷿid, ҩƷid, ����, ����, ��������, ʵ������, ʵ�ʽ��, ʵ�ʲ��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴ���������, �ϴβ���, ���Ч��, Ч��, ���ۼ�, �ϴο���, ��Ʒ����,
         �ڲ�����, ƽ���ɱ���, ��׼�ĺ�)
      Values
        (v_�շ�.�ⷿid, v_�շ�.ҩƷid, v_�շ�.����, 1, n_��������, v_�շ�.ʵ������, v_�շ�.���۽��, v_�շ�.���, v_�շ�.��ҩ��λid, v_�շ�.�ɱ���, v_�շ�.����,
         v_�շ�.��������, v_�շ�.����, v_�շ�.���Ч��, v_�շ�.Ч��, Decode(Nvl(v_�շ�.����, 0), 0, Null, Decode(n_ʵ������, 1, v_�շ�.���ۼ�, Null)),
         v_�շ�.����, v_�շ�.��Ʒ����, v_�ڲ�����, v_�շ�.�ɱ���, v_�շ�.��׼�ĺ�);
    End If;
  
    If v_�շ�.�ڲ����� Is Null And Not v_�ڲ����� Is Null Then
      Update ҩƷ�շ���¼ Set �ڲ����� = v_�ڲ����� Where ID = v_�շ�.Id;
    End If;
  
    --����������Ϊ��ļ�¼
    Delete From ҩƷ���
    Where �ⷿid = v_�շ�.�ⷿid And ҩƷid = v_�շ�.ҩƷid And Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0 And
          Nvl(ʵ�ʲ��, 0) = 0;
  
    --���Ĳ����շ����ܱ����Ӧ����
    --���¸ò��ϵĳɱ���
    Begin
      Select Sum(Nvl(ʵ�ʽ��, 0)), Sum(Nvl(ʵ�ʲ��, 0)), Sum(Nvl(ʵ������, 0))
      Into n_�����, n_�����, n_�������
      From ҩƷ���
      Where ���� = 1 And ҩƷid = v_�շ�.ҩƷid;
    Exception
      When Others Then
        n_������� := 0;
    End;
  
    --���¸�ҩƷ�ĳɱ���
    Update ��������
    Set �ɱ��� = v_�շ�.�ɱ���, �ϴ��ۼ� = v_�շ�.���ۼ�, �ϴι�Ӧ��id = v_�շ�.��ҩ��λid, �ϴβ��� = v_�շ�.����
    Where ����id = v_�շ�.ҩƷid;
  
    --���Ĳ��������е�ע��֤��:������ֲ������Ա��е�ע��֤��û���ֱ�ӷ�д���������Ա��е�ע��֤��
    If Nvl(v_�շ�.ע��֤��, ' ') <> ' ' Then
      Update �������� Set ע��֤�� = v_�շ�.ע��֤�� Where ����id = v_�շ�.ҩƷid And ע��֤�� Is Null;
    End If;
  
    --���¼�������е�ƽ���ɱ���
    Update ҩƷ���
    Set ƽ���ɱ��� = Decode(Nvl(����, 0), 0, Decode((ʵ�ʽ�� - ʵ�ʲ��) / ʵ������, 0, �ϴβɹ���, (ʵ�ʽ�� - ʵ�ʲ��) / ʵ������), �ϴβɹ���)
    Where ҩƷid = v_�շ�.ҩƷid And Nvl(����, 0) = Nvl(v_�շ�.����, 0) And �ⷿid = v_�շ�.�ⷿid And ���� = 1 And Nvl(ʵ������, 0) <> 0;
    If Sql%NotFound Then
      Select �ɱ��� Into n_ƽ���ɱ��� From �������� Where ����id = v_�շ�.ҩƷid;
      Update ҩƷ���
      Set ƽ���ɱ��� = n_ƽ���ɱ���
      Where ҩƷid = v_�շ�.ҩƷid And �ⷿid = v_�շ�.�ⷿid And Nvl(����, 0) = Nvl(v_�շ�.����, 0) And ���� = 1;
    End If;
  End Loop;

  --��Ӧ��������д���
  --�˴���һ���飬��Ҫ�ǽ��û�ж�Ӧ��Ʊ�ŵļ�¼
  Begin
    Update Ӧ����¼
    Set ����� = �����_In, ������� = �������_In
    Where ��ⵥ�ݺ� = No_In And ϵͳ��ʶ = 5 And ��¼���� = 0 And ��¼״̬ = 1;
  
    Select b.��λid, Sum(��Ʊ���)
    Into n_��λid, n_��Ʊ���
    From ҩƷ�շ���¼ A, Ӧ����¼ B
    Where a.Id = b.�շ�id And a.No = No_In And a.���� = 15 And b.ϵͳ��ʶ = 5
    Group By b.��λid;
  
    If Nvl(n_��λid, 0) <> 0 Then
      Update Ӧ����� Set ��� = Nvl(���, 0) + Nvl(n_��Ʊ���, 0) Where ��λid = n_��λid And ���� = 1;
    
      If Sql%NotFound Then
        Insert Into Ӧ����� (��λid, ����, ���) Values (n_��λid, 1, n_��Ʊ���);
      End If;
    End If;
  Exception
    When No_Data_Found Then
      Null;
  End;

  --������Զ������ı���������ⵥ��������ƿⵥ
  For v_Data In (Select ID, ���, ʵ������, ��ҩ��ʽ, ���ۼ�, ���۽��, ���, �ⷿid, ҩƷid, ����, ��ҩ��λid, �ɱ���, �ɱ����, ����, Ч��, ���Ч��, ��������, ����,
                        ������id, ע��֤��, ����, ժҪ, ��Ʒ����, �ڲ�����, ����id, �����, �������
                 From ҩƷ�շ���¼
                 Where NO = No_In And ���� = 15 And ��¼״̬ = 1 And ������� Is Not Null And ����id > 0
                 Order By ҩƷid, ����,���) Loop
    If v_�Է��ⷿid = 0 Then
      Begin
        Select Distinct �ⷿid Into v_�Է��ⷿid From ҩƷ�շ���¼ Where ���� In (24, 25) And ����id = v_Data.����id;
      Exception
        When Others Then
          v_�Է��ⷿid := 0;
      End;
    End If;
  
    If v_�Է��ⷿid > 0 Then
      If v_�ƿ�no Is Null Then
        v_�ƿ�no := Nextno(72, v_Data.�ⷿid);
      End If;
    
      Zl_�����ƿ�_Insert(v_�ƿ�no, v_Data.��� * 2 - 1, v_Data.�ⷿid, v_�Է��ⷿid, v_Data.ҩƷid, v_Data.����, v_Data.ʵ������, v_Data.ʵ������,
                     v_Data.�ɱ���, v_Data.�ɱ����, v_Data.���ۼ�, v_Data.���۽��, v_Data.���, v_Data.�����, v_Data.����, v_Data.����,
                     v_Data.Ч��, v_Data.���Ч��, v_Data.ժҪ, v_Data.�������);
    End If;
  End Loop;

  --���²������ƿⵥ���б��Ϻ����
  If Not v_�ƿ�no Is Null Then
    Zl_�����ƿ�_Prepare(v_�ƿ�no, �����_In);
    Zl_�����ƿ�_Prepare(v_�ƿ�no);
  
    Select b.Id As ���id
    Into v_�����id
    From ҩƷ�������� A, ҩƷ������ B
    Where a.���id = b.Id And a.���� = 34 And ϵ�� = 1 And Rownum < 2;
  
    Select b.Id As ���id
    Into v_�����id
    From ҩƷ�������� A, ҩƷ������ B
    Where a.���id = b.Id And a.���� = 34 And ϵ�� = -1 And Rownum < 2;
  
    For v_Data In (Select ���, �ⷿid, �Է�����id, ҩƷid, ����, Nvl(����, 0) As ����, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ���۽��, ���, NO, ������, ����,
                          Ч��, ���Ч��, ��������
                   From ҩƷ�շ���¼
                   Where ���� = 19 And NO = v_�ƿ�no And ������� Is Null And ���ϵ�� = -1
                   Order By ҩƷid, ����,���) Loop
    
      Zl_�����ƿ�_Verify(v_Data.���, v_Data.�ⷿid, v_Data.�Է�����id, v_Data.ҩƷid, v_Data.����, v_Data.����, v_Data.��д����, v_Data.ʵ������,
                     v_Data.�ɱ���, v_Data.�ɱ����, v_Data.���۽��, v_Data.���, v_�����id, v_�����id, v_Data.No, v_Data.������, v_Data.����,
                     v_Data.Ч��, v_Data.���Ч��, v_Data.��������, 1, v_Data.���ۼ�);
    End Loop;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����⹺_Verify;
/

--117925:����,2017-12-07,��������������
Create Or Replace Procedure Zl_�����⹺_Delete(
                                           --ɾ��ҩƷ�շ���¼����Ӧ�ı�Ӧ����¼
                                           No_In In ҩƷ�շ���¼.No%Type) Is
  Merritem Exception;
  Merrmsg Varchar2(100);
Begin

  --�ָ���������
  For v_�շ� In (Select ʵ������, �ⷿid, ����, ҩƷid, �ɱ���, ����, ��������, ���Ч��, Ч��, ����, ��ҩ��λid, ��׼�ĺ�
               From ҩƷ�շ���¼
               Where NO = No_In And Nvl(��ҩ��ʽ, 0) = 1 And ���� = 15
               Order By ҩƷid,����,���) Loop
    Update ҩƷ���
    Set �������� = Nvl(��������, 0) + (-1 * v_�շ�.ʵ������)
    Where �ⷿid = v_�շ�.�ⷿid And ҩƷid = v_�շ�.ҩƷid And Nvl(����, 0) = Nvl(v_�շ�.����, 0) And ���� = 1;
  
    If Sql%NotFound Then
      Insert Into ҩƷ���
        (�ⷿid, ҩƷid, ����, ����, ��������, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴβ���, Ч��, ���Ч��, �ϴ���������, ��׼�ĺ�)
      Values
        (v_�շ�.�ⷿid, v_�շ�.ҩƷid, v_�շ�.����, 1, -1 * v_�շ�.ʵ������, v_�շ�.��ҩ��λid, v_�շ�.�ɱ���, v_�շ�.����, v_�շ�.����, v_�շ�.Ч��, v_�շ�.���Ч��,
         v_�շ�.��������, v_�շ�.��׼�ĺ�);
    End If;
  
    Delete From ҩƷ���
    Where �ⷿid = v_�շ�.�ⷿid And ҩƷid = v_�շ�.ҩƷid And Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0 And
          Nvl(ʵ�ʲ��, 0) = 0;
  
  End Loop;

  Delete Ӧ����¼ Where ϵͳ��ʶ = 5 And �շ�id In (Select ID From ҩƷ�շ���¼ Where NO = No_In And ���� = 15);
  --��ӦӦ����¼��ɾ��ͨ������ɾ��
  Delete --ɾ������
  From ҩƷ�շ���¼
  Where NO = No_In And ���� = 15 And ��¼״̬ = 1 And ����� Is Null;

  If Sql%RowCount = 0 Then
    Merrmsg := '[ZLSOFT]�õ����Ѿ�������ɾ�����ѱ�����ˣ�[ZLSOFT]';
    Raise Merritem;
  End If;
Exception
  When Merritem Then
    Raise_Application_Error(-20101, Merrmsg);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����⹺_Delete;
/

--117925:����,2017-12-07,��������������
Create Or Replace Procedure Zl_ҩƷ�̵�_Delete(
                                           --ɾ��ҩƷ�շ���¼���ָ���Ӧ�ı�ҩƷ���
                                           No_In In ҩƷ�շ���¼.No%Type) Is
  Err_Isverified Exception;

  Cursor c_ҩƷ�շ���¼ Is
    Select ID From ҩƷ�շ���¼ Where NO = No_In And ���� = 12 Order By ҩƷid,����;
Begin
  --ͨ��ѭ�����ָ��������ԭ���Ŀ���������
  --ʵ�������������������
  For v_ҩƷ�շ���¼ In c_ҩƷ�շ���¼ Loop
    Zl_ҩƷ���_Update(v_ҩƷ�շ���¼.Id, 1);
  End Loop;

  Delete From ҩƷ�շ���¼ Where NO = No_In And ���� = 12 And ��¼״̬ = 1 And ����� Is Null;

  If Sql%RowCount = 0 Then
    Raise Err_Isverified;
  End If;
Exception
  When Err_Isverified Then
    Raise_Application_Error(-20101, '[ZLSOFT]�õ����Ѿ�������ɾ�����ѱ�����ˣ�[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�̵�_Delete;
/

--117925:����,2017-12-07,��������������
Create Or Replace Procedure Zl_ҩƷ�̵�_Verify
(
  No_In     In ҩƷ�շ���¼.No%Type := Null,
  �����_In In ҩƷ�շ���¼.�����%Type := Null
) Is
  Err_Isverified Exception;

  Cursor c_ҩƷ�շ���¼ Is
    Select ID, ʵ������, ���۽��, ���, �ⷿid, ҩƷid, ����, ����, Ч��, ����, ������id, ���ϵ��, ��׼�ĺ�, ��ҩ��λid, ��������, ����
    From ҩƷ�շ���¼
    Where NO = No_In And ���� = 12 And ��¼״̬ = 1
    Order By ҩƷid,����;
Begin
  Update ҩƷ�շ���¼
  Set ����� = �����_In, ������� = Sysdate
  Where NO = No_In And ���� = 12 And ��¼״̬ = 1 And ����� Is Null;

  If Sql%RowCount = 0 Then
    Raise Err_Isverified;
  End If;

  For v_ҩƷ�շ���¼ In c_ҩƷ�շ���¼ Loop
    --����ҩƷ�������Ӧ����
    Zl_ҩƷ���_Update(v_ҩƷ�շ���¼.Id, 0);
  End Loop;
Exception
  When Err_Isverified Then
    Raise_Application_Error(-20101, '[ZLSOFT]�õ����Ѿ���������ˣ�[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�̵�_Verify;
/

--117925:����,2017-12-07,��������������
Create Or Replace Procedure Zl_ҩƷ�̵�_Strike
(
  No_In     In ҩƷ�շ���¼.No%Type,
  �����_In In ҩƷ�շ���¼.�����%Type
) Is
  Err_Isstriked Exception;

  Cursor c_ҩƷ�շ���¼ Is
    Select a.Id, a.ʵ������, a.���۽��, a.���, a.���ۼ�, Nvl(b.�Ƿ���, 0) As �Ƿ���, a.�ⷿid, a.ҩƷid, a.����, a.����, a.Ч��, a.����, a.������id,
           a.���ϵ��, a.����, a.��׼�ĺ�, a.��ҩ��λid, a.��������
    From ҩƷ�շ���¼ A, �շ���ĿĿ¼ B
    Where a.ҩƷid = b.Id And NO = No_In And ���� = 12 And ��¼״̬ = 2
    Order By ҩƷid,����;
Begin
  Update ҩƷ�շ���¼ Set ��¼״̬ = 3 Where NO = No_In And ���� = 12 And ��¼״̬ = 1;

  If Sql%RowCount = 0 Then
    Raise Err_Isstriked;
  End If;
  Insert Into ҩƷ�շ���¼
    (ID, ��¼״̬, ����, NO, ���, �ⷿid, ������id, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ��д����, ����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ���۽��, ���, ժҪ, ������,
     ��������, �����, �������, Ƶ��, ����, ��׼�ĺ�, ��ҩ��λid, ��������, �ⷿ��λ)
    Select ҩƷ�շ���¼_Id.Nextval, 2, ����, NO, ���, �ⷿid, ������id, ���ϵ��, a.ҩƷid,
           Decode(Nvl(a.����, 0), 0, Null, (Decode(Nvl(b.ҩ�����, 0), 0, Null, a.����))), a.����, ����, Ч��, ��д����, a.����, -ʵ������,
           a.�ɱ���, �ɱ����, ���ۼ�, -���۽��, -���, ժҪ, �����_In, Sysdate, �����_In, Sysdate, Ƶ��, ����, a.��׼�ĺ�, a.��ҩ��λid, a.��������, a.�ⷿ��λ
    From (Select * From ҩƷ�շ���¼ Where NO = No_In And ���� = 12 And ��¼״̬ = 3 Order By ҩƷid) A, ҩƷ��� B
    Where a.ҩƷid = b.ҩƷid;

  For v_ҩƷ�շ���¼ In c_ҩƷ�շ���¼ Loop
    --������
    Zl_ҩƷ���_Update(v_ҩƷ�շ���¼.Id, 0);
  
    --������ۺ����
    Zl_ҩƷ�շ���¼_��������(v_ҩƷ�շ���¼.Id);
  End Loop;
Exception
  When Err_Isstriked Then
    Raise_Application_Error(-20101, '[ZLSOFT]�õ����Ѿ������˳�����[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�̵�_Strike;
/

--117925:����,2017-12-07,��������������
Create Or Replace Procedure Zl_ҩƷ�ƿ�_Verify
(
  ���_In         In ҩƷ�շ���¼.���%Type,
  �ⷿid_In       In ҩƷ�շ���¼.�ⷿid%Type,
  �Է�����id_In   In ҩƷ�շ���¼.�Է�����id%Type,
  ҩƷid_In       In ҩƷ�շ���¼.ҩƷid%Type,
  ����_In         In ҩƷ�շ���¼.����%Type,
  ������_In       In ҩƷ�շ���¼.����%Type,
  ʵ������_In     In ҩƷ�շ���¼.ʵ������%Type,
  �ɱ���_In       In ҩƷ�շ���¼.�ɱ���%Type,
  �ɱ����_In     In ҩƷ�շ���¼.�ɱ����%Type,
  ���۽��_In     In ҩƷ�շ���¼.���۽��%Type,
  ���_In         In ҩƷ�շ���¼.���%Type,
  No_In           In ҩƷ�շ���¼.No%Type,
  �����_In       In ҩƷ�շ���¼.�����%Type,
  ����_In         In ҩƷ�շ���¼.����%Type := Null,
  Ч��_In         In ҩƷ�շ���¼.Ч��%Type := Null,
  �������_In     In ҩƷ�շ���¼.�������%Type := Null,
  �ϴι�Ӧ��id_In In ҩƷ�շ���¼.��ҩ��λid%Type := Null,
  ��׼�ĺ�_In     In ҩƷ�շ���¼.��׼�ĺ�%Type := Null,
  ���ۼ�_In       In ҩƷ�շ���¼.���ۼ�%Type := Null
) Is
  Err_Isverified Exception;
  Err_Isnonumber Exception;
  Err_Isbatch Exception;
  Err_Isprice Exception;
  v_Druginf  Varchar2(50); --ԭ���������ڷ�����ҩƷ��Ϣ 
  v_ʵ������ ҩƷ���.ʵ������%Type;
  v_����     �շ���ĿĿ¼.����%Type;
  Intdigit   Number;
  v_�ϴο��� ҩƷ���.�ϴο���%Type;
  Cursor c_ҩƷ�շ���¼ Is
    Select ID
    From ҩƷ�շ���¼
    Where NO = No_In And ���� = 6 And ҩƷid = ҩƷid_In And ��¼״̬ = 1 And ��� In (���_In, ���_In + 1) And ������� Is Not Null
	Order By ҩƷid,����;
Begin
  --��ȡ���С��λ�� 
  Select Nvl(����, 2) Into Intdigit From ҩƷ���ľ��� Where ���� = 0 And ��� = 1 And ���� = 4 And ��λ = 5;

  --��Ҫ���ԭ���������ڷ�����ҩƷ�����ܶ������ 
  --���������� 
  Begin
    Select Distinct '(' || i.���� || ')' || Nvl(n.����, i.����) As ҩƷ��Ϣ
    Into v_Druginf
    From ҩƷ�շ���¼ A, ҩƷ��� B, �շ���ĿĿ¼ I, �շ���Ŀ���� N
    Where a.ҩƷid = b.ҩƷid And a.ҩƷid = i.Id And a.ҩƷid = n.�շ�ϸĿid(+) And n.����(+) = 3 And a.No = No_In And a.���� = 6 And
          a.��¼״̬ = 1 And Nvl(a.����, 0) = 0 And a.ҩƷid + 0 = ҩƷid_In And a.��� = ���_In + 1 And
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

  Begin
    Select Nvl(ʵ������, 0), Nvl(�ϴο���, 100)
    Into v_ʵ������, v_�ϴο���
    From ҩƷ���
    Where ҩƷid = ҩƷid_In And Nvl(����, 0) = ������_In And �ⷿid = �ⷿid_In And ���� = 1 And Rownum = 1;
  Exception
    When Others Then
      v_ʵ������ := 0;
      v_�ϴο��� := 100;
  End;

  If ������_In > 0 Then
    If v_ʵ������ < ʵ������_In Then
      Raise Err_Isnonumber;
    End If;
  End If;

  Update ҩƷ�շ���¼
  Set ����� = Nvl(�����_In, �����), ������� = �������_In, ʵ������ = ʵ������_In, �ɱ��� = �ɱ���_In, �ɱ���� = �ɱ����_In, ���ۼ� = ���ۼ�_In, ���۽�� = ���۽��_In,
      ��� = ���_In, ���� = v_�ϴο���
  Where NO = No_In And ���� = 6 And ҩƷid = ҩƷid_In And ��¼״̬ = 1 And ��� In (���_In, ���_In + 1) And ����� Is Null;

  If Sql%RowCount = 0 Then
    Raise Err_Isverified;
  End If;

  --�����������
  For v_ҩƷ�շ���¼ In c_ҩƷ�շ���¼ Loop
    Zl_ҩƷ���_Update(v_ҩƷ�շ���¼.Id);
  End Loop;

Exception
  When Err_Isverified Then
    Raise_Application_Error(-20101, '[ZLSOFT]�õ����Ѿ���������ˣ�[ZLSOFT]');
  When Err_Isnonumber Then
    Select ���� Into v_���� From �շ���ĿĿ¼ Where ID = ҩƷid_In;
    Raise_Application_Error(-20101,
                            '[ZLSOFT]����Ϊ' || v_���� || ',����Ϊ' || ����_In || '��ҩ�����ҩƷ' || Chr(10) || Chr(13) ||
                             '���ÿ������������[ZLSOFT]');
  When Err_Isbatch Then
    Raise_Application_Error(-20102, '[ZLSOFT]�õ����а�����һ��ԭ�������������ڷ�����ҩƷ[' || v_Druginf || ']��������ˣ�[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�ƿ�_Verify;
/

--117925:����,2017-12-07,��������������
Create Or Replace Procedure Zl_ҩƷ�ƿ�_Delete
(
  No_In       In ҩƷ�շ���¼.No%Type,
  ��¼״̬_In In ҩƷ�շ���¼.��¼״̬%Type := 1
) Is
  Err_Isverified Exception;

  Cursor c_ҩƷ�շ���¼ Is
    Select ID From ҩƷ�շ���¼ Where NO = No_In And ���� = 6 Order By ҩƷid,����;

  Cursor c_���������¼ Is
    Select ID From ҩƷ�շ���¼ Where NO = No_In And ���� = 6 And ��¼״̬ = ��¼״̬_In Order By ҩƷid,����;
Begin

  If ��¼״̬_In = 1 Then
    --����δ����ƿⵥ��
    --ͨ��ѭ�����ָ�ԭ���Ŀ�������
    For v_ҩƷ�շ���¼ In c_ҩƷ�շ���¼ Loop
      Zl_ҩƷ���_Update(v_ҩƷ�շ���¼.Id, 1);
    End Loop;
  Else
    --�����ƿ������������
    --ͨ��ѭ�����ָ�ԭ���Ŀ�������
    For v_���������¼ In c_���������¼ Loop
      Zl_ҩƷ���_Update(v_���������¼.Id, 1);
    End Loop;
  End If;

  --ɾ��δ��˵���
  Delete From ҩƷ�շ���¼ Where NO = No_In And ���� = 6 And ��¼״̬ = ��¼״̬_In And ����� Is Null;

  If Sql%RowCount = 0 Then
    Raise Err_Isverified;
  End If;
Exception
  When Err_Isverified Then
    Raise_Application_Error(-20101, '[ZLSOFT]�õ����Ѿ�������ɾ�����ѱ�����ˣ�[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�ƿ�_Delete;
/

--117925:����,2017-12-07,��������������
Create Or Replace Procedure Zl_ҩƷ�ƿ�_Strike
(
  �д�_In       In Integer,
  ԭ��¼״̬_In In ҩƷ�շ���¼.��¼״̬%Type,
  No_In         In ҩƷ�շ���¼.No%Type,
  ���_In       In ҩƷ�շ���¼.���%Type,
  ҩƷid_In     In ҩƷ�շ���¼.ҩƷid%Type,
  ��������_In   In ҩƷ�շ���¼.ʵ������%Type,
  ������_In     In ҩƷ�շ���¼.������%Type,
  ��������_In   In ҩƷ�շ���¼.��������%Type,
  ժҪ_In       In ҩƷ�շ���¼.ժҪ%Type := Null,
  ������ʽ_In   In Integer := 0 --0������������ʽ��1�������������뵥�ݣ�2������Ѳ����ĳ������뵥��
) Is
  Err_Isstriked Exception;
  Err_Isoutstock Exception;
  Err_Isnonum Exception;
  Err_Isbatch Exception;
  v_Druginf      Varchar2(50); --ԭ���������ڷ�����ҩƷ��Ϣ
  v_�ⷿid       ҩƷ�շ���¼.�ⷿid%Type;
  v_����         ҩƷ�շ���¼.����%Type;
  v_�ɱ���       ҩƷ�շ���¼.�ɱ���%Type;
  v_�ɱ����     ҩƷ�շ���¼.�ɱ����%Type;
  v_���ۼ�       ҩƷ�շ���¼.���ۼ�%Type;
  v_���۽��     ҩƷ�շ���¼.���۽��%Type;
  v_���         ҩƷ�շ���¼.���%Type;
  v_ʣ������     ҩƷ�շ���¼.ʵ������%Type;
  v_ʣ��ɱ���� ҩƷ�շ���¼.�ɱ����%Type;
  v_ʣ�����۽�� ҩƷ�շ���¼.���۽��%Type;
  v_�շ�id       ҩƷ�շ���¼.Id%Type;
  v_��׼�ĺ�     ҩƷ�շ���¼.��׼�ĺ�%Type;

  v_ҩ����� Integer;
  v_ҩ������ Integer;
  Intdigit   Number;

  Cursor c_ҩƷ�շ���¼ Is
    Select a.Id, a.���, a.�ⷿid, a.�Է�����id, a.������id, a.���ϵ��, a.ҩƷid, a.����, a.����, a.����, a.Ч��, a.��ҩ��, a.��ҩ����, a.ժҪ, a.��ҩ��λid,
           a.��׼�ĺ�, a.��������, a.�ɱ���, a.���ۼ�, Nvl(b.�Ƿ���, 0) As ʱ��, a.����, a.����, a.Ƶ��
    From ҩƷ�շ���¼ A, �շ���ĿĿ¼ B
    Where a.ҩƷid = b.Id And a.No = No_In And a.���� = 6 And (a.��� >= ���_In And a.��� <= ���_In + 1) And
          (a.��¼״̬ = 1 Or Mod(a.��¼״̬, 3) = 0)
    Order By a.ҩƷid,a.����;

  Cursor c_���������¼ Is
    Select a.Id, a.���, a.�ⷿid, a.�Է�����id, a.������id, a.���ϵ��, a.ҩƷid, a.����, a.����, a.����, a.Ч��, a.��ҩ��, a.��ҩ����, a.ժҪ, a.��ҩ��λid,
           a.��׼�ĺ�, a.��������, a.�ɱ���, a.ʵ������, a.���۽��, a.���, a.���ۼ�, Nvl(b.�Ƿ���, 0) As ʱ��, a.����, a.����, a.Ƶ��
    From ҩƷ�շ���¼ A, �շ���ĿĿ¼ B
    Where a.ҩƷid = b.Id And a.No = No_In And a.���� = 6 And (a.��� >= ���_In And a.��� <= ���_In + 1) And
          (a.��¼״̬ = ԭ��¼״̬_In And Mod(a.��¼״̬, 3) = 2) And a.������� Is Null
    Order By a.ҩƷid,a.����;
Begin
  --��ȡ���С��λ��
  Select Nvl(����, 2) Into Intdigit From ҩƷ���ľ��� Where ���� = 0 And ��� = 1 And ���� = 4 And ��λ = 5;

  If ������ʽ_In = 1 Then
    --�����������뵥�ݣ�����д����ˡ�������ڣ������¿���¼
    If �д�_In = 1 Then
      Update ҩƷ�շ���¼
      Set ��¼״̬ = Decode(ԭ��¼״̬_In, 1, 3, ԭ��¼״̬_In + 3)
      Where NO = No_In And ���� = 6 And ��¼״̬ = ԭ��¼״̬_In;
      If Sql%RowCount = 0 Then
        Raise Err_Isstriked;
      End If;
    End If;
  
    --��Ҫ���ԭ���������ڷ�����ҩƷ�����ܶ������
    Begin
      Select Distinct '(' || i.���� || ')' || Nvl(n.����, i.����) As ҩƷ��Ϣ
      Into v_Druginf
      From ҩƷ�շ���¼ A, ҩƷ��� B, �շ���ĿĿ¼ I, �շ���Ŀ���� N
      Where a.ҩƷid = b.ҩƷid And a.ҩƷid = i.Id And a.ҩƷid = n.�շ�ϸĿid(+) And n.����(+) = 3 And a.No = No_In And a.���� = 6 And
            a.ҩƷid + 0 = ҩƷid_In And Mod(a.��¼״̬, 3) = 0 And Nvl(a.����, 0) = 0 And a.��� = ���_In And
            ((Nvl(b.ҩ�����, 0) = 1 And
            a.�ⷿid Not In (Select ����id From ��������˵�� Where (�������� Like '%ҩ��') Or (�������� Like '�Ƽ���'))) Or
            Nvl(b.ҩ������, 0) = 1) And Rownum = 1;
    Exception
      When Others Then
        v_Druginf := '';
    End;
  
    If v_Druginf Is Not Null Then
      Raise Err_Isbatch;
    End If;
  
    Select Sum(a.ʵ������) As ʣ������, Sum(a.�ɱ����) As ʣ��ɱ����, Sum(a.���۽��) As ʣ�����۽��, a.�ɱ���, a.���ۼ�, a.�Է�����id, Nvl(a.����, 0),
           b.ҩ�����, b.ҩ������, a.��׼�ĺ�
    Into v_ʣ������, v_ʣ��ɱ����, v_ʣ�����۽��, v_�ɱ���, v_���ۼ�, v_�ⷿid, v_����, v_ҩ�����, v_ҩ������, v_��׼�ĺ�
    From ҩƷ�շ���¼ A, ҩƷ��� B
    Where a.No = No_In And a.ҩƷid = b.ҩƷid And a.���� = 6 And a.ҩƷid = ҩƷid_In And a.��� = ���_In
    Group By a.�ɱ���, a.���ۼ�, a.�Է�����id, Nvl(a.����, 0), b.ҩ�����, b.ҩ������, a.��׼�ĺ�;
  
    --V_����:(ԭ����,�ַ���Ϊ����;����Ϊ��)ԭ������,�ַ���,�����̲�����
    --��Ϊ���ڶԷ��ⷿ���൱�ڳ��⣬�����ڵ�ǰ�ⷿ���൱����⣬���Ե�ǰ�ⷿ�����飬������˿�������¼
    Select Nvl(a.����, 0)
    Into v_����
    From ҩƷ�շ���¼ A
    Where a.No = No_In And a.���� = 6 And a.ҩƷid = ҩƷid_In And a.��� = ���_In + 1 And Mod(a.��¼״̬, 3) = 0;
  
    --������������ʣ��������������
    If v_ʣ������ < ��������_In Then
      Raise Err_Isnonum;
    End If;
  
    v_�ɱ���� := Round(��������_In / v_ʣ������ * v_ʣ��ɱ����, Intdigit);
    v_���۽�� := Round(��������_In / v_ʣ������ * v_ʣ�����۽��, Intdigit);
    v_���     := v_���۽�� - v_�ɱ����;
  
    For v_ҩƷ�շ���¼ In c_ҩƷ�շ���¼ Loop
    
      Select ҩƷ�շ���¼_Id.Nextval Into v_�շ�id From Dual;
      Insert Into ҩƷ�շ���¼
        (ID, ��¼״̬, ����, NO, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ���۽��, ���,
         ժҪ, ������, ��������, �����, �������, ��ҩ��, ��ҩ����, ��ҩ��λid, ��׼�ĺ�, ��������, ����, ����, Ƶ��)
      Values
        (v_�շ�id, Decode(ԭ��¼״̬_In, 1, 2, ԭ��¼״̬_In + 2), 6, No_In, v_ҩƷ�շ���¼.���, v_ҩƷ�շ���¼.�ⷿid, v_ҩƷ�շ���¼.�Է�����id,
         v_ҩƷ�շ���¼.������id, v_ҩƷ�շ���¼.���ϵ��, ҩƷid_In, v_ҩƷ�շ���¼.����, v_ҩƷ�շ���¼.����, v_ҩƷ�շ���¼.����, v_ҩƷ�շ���¼.Ч��, -��������_In, -��������_In,
         v_�ɱ���, -v_�ɱ����, v_���ۼ�, -v_���۽��, -v_���, ժҪ_In, ������_In, ��������_In, Null, Null, v_ҩƷ�շ���¼.��ҩ��, v_ҩƷ�շ���¼.��ҩ����,
         v_ҩƷ�շ���¼.��ҩ��λid, v_ҩƷ�շ���¼.��׼�ĺ�, v_ҩƷ�շ���¼.��������, v_ҩƷ�շ���¼.����, v_ҩƷ�շ���¼.����, v_ҩƷ�շ���¼.Ƶ��);
    
      --ԭ���ⷿ�ٹ�ѡ��¿��ʱҪ�¿��
      Zl_ҩƷ���_Update(v_�շ�id, 0, 1);
    End Loop;
  
  Elsif ������ʽ_In = 2 Then
    --����Ѳ����ĳ������뵥�ݣ���д����ˡ�������ڣ����¿���¼
    For v_ҩƷ�շ���¼ In c_���������¼ Loop
      --��д����ˡ��������
      Update ҩƷ�շ���¼
      Set ����� = ������_In, ������� = ��������_In
      Where NO = No_In And ���� = 6 And ID = v_ҩƷ�շ���¼.Id;
    
      --����ҩƷ�������Ӧ���ݣ�ע����ʱ������������Ǹ���
      --����Ϊ1��ʾ�������ʱ�¿�������������ԭ����ⷿ�����˿��������Ͳ����ٸ��¿���������
      Zl_ҩƷ���_Update(v_ҩƷ�շ���¼.Id, 0, 1);
    
      --������ۺ����
      Zl_ҩƷ�շ���¼_��������(v_ҩƷ�շ���¼.Id);
    End Loop;
  Else
    --����������ʽ������������¼����д����ˡ�������ڣ����¿���¼
    If �д�_In = 1 Then
      Update ҩƷ�շ���¼
      Set ��¼״̬ = Decode(ԭ��¼״̬_In, 1, 3, ԭ��¼״̬_In + 3)
      Where NO = No_In And ���� = 6 And ��¼״̬ = ԭ��¼״̬_In;
      If Sql%RowCount = 0 Then
        Raise Err_Isstriked;
      End If;
    End If;
  
    --��Ҫ���ԭ���������ڷ�����ҩƷ�����ܶ������
    Begin
      Select Distinct '(' || i.���� || ')' || Nvl(n.����, i.����) As ҩƷ��Ϣ
      Into v_Druginf
      From ҩƷ�շ���¼ A, ҩƷ��� B, �շ���ĿĿ¼ I, �շ���Ŀ���� N
      Where a.ҩƷid = b.ҩƷid And a.ҩƷid = i.Id And a.ҩƷid = n.�շ�ϸĿid(+) And n.����(+) = 3 And a.No = No_In And a.���� = 6 And
            a.ҩƷid + 0 = ҩƷid_In And Mod(a.��¼״̬, 3) = 0 And Nvl(a.����, 0) = 0 And a.��� = ���_In And
            ((Nvl(b.ҩ�����, 0) = 1 And
            a.�ⷿid Not In (Select ����id From ��������˵�� Where (�������� Like '%ҩ��') Or (�������� Like '�Ƽ���'))) Or
            Nvl(b.ҩ������, 0) = 1) And Rownum = 1;
    Exception
      When Others Then
        v_Druginf := '';
    End;
  
    If v_Druginf Is Not Null Then
      Raise Err_Isbatch;
    End If;
  
    Select Sum(a.ʵ������) As ʣ������, Sum(a.�ɱ����) As ʣ��ɱ����, Sum(a.���۽��) As ʣ�����۽��, a.�ɱ���, a.���ۼ�, a.�Է�����id, Nvl(a.����, 0),
           b.ҩ�����, b.ҩ������, a.��׼�ĺ�
    Into v_ʣ������, v_ʣ��ɱ����, v_ʣ�����۽��, v_�ɱ���, v_���ۼ�, v_�ⷿid, v_����, v_ҩ�����, v_ҩ������, v_��׼�ĺ�
    From ҩƷ�շ���¼ A, ҩƷ��� B
    Where a.No = No_In And a.ҩƷid = b.ҩƷid And a.���� = 6 And a.ҩƷid = ҩƷid_In And a.��� = ���_In
    Group By a.�ɱ���, a.���ۼ�, a.�Է�����id, Nvl(a.����, 0), b.ҩ�����, b.ҩ������, a.��׼�ĺ�;
  
    --V_����:(ԭ����,�ַ���Ϊ����;����Ϊ��)ԭ������,�ַ���,�����̲�����
    --��Ϊ���ڶԷ��ⷿ���൱�ڳ��⣬�����ڵ�ǰ�ⷿ���൱����⣬���Ե�ǰ�ⷿ�����飬������˿�������¼
    Select Nvl(a.����, 0)
    Into v_����
    From ҩƷ�շ���¼ A
    Where a.No = No_In And a.���� = 6 And a.ҩƷid = ҩƷid_In And a.��� = ���_In + 1 And Mod(a.��¼״̬, 3) = 0;
  
    --������������ʣ��������������
    If v_ʣ������ < ��������_In Then
      Raise Err_Isnonum;
    End If;
  
    v_�ɱ���� := Round(��������_In / v_ʣ������ * v_ʣ��ɱ����, Intdigit);
    v_���۽�� := Round(��������_In / v_ʣ������ * v_ʣ�����۽��, Intdigit);
    v_���     := v_���۽�� - v_�ɱ����;
  
    For v_ҩƷ�շ���¼ In c_ҩƷ�շ���¼ Loop
      Select ҩƷ�շ���¼_Id.Nextval Into v_�շ�id From Dual;
      Insert Into ҩƷ�շ���¼
        (ID, ��¼״̬, ����, NO, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ���۽��, ���,
         ժҪ, ������, ��������, �����, �������, ��ҩ��, ��ҩ����, ��ҩ��λid, ��׼�ĺ�, ��������, ����, ����, Ƶ��)
      Values
        (v_�շ�id, Decode(ԭ��¼״̬_In, 1, 2, ԭ��¼״̬_In + 2), 6, No_In, v_ҩƷ�շ���¼.���, v_ҩƷ�շ���¼.�ⷿid, v_ҩƷ�շ���¼.�Է�����id,
         v_ҩƷ�շ���¼.������id, v_ҩƷ�շ���¼.���ϵ��, ҩƷid_In, v_ҩƷ�շ���¼.����, v_ҩƷ�շ���¼.����, v_ҩƷ�շ���¼.����, v_ҩƷ�շ���¼.Ч��, -��������_In, -��������_In,
         v_�ɱ���, -v_�ɱ����, v_���ۼ�, -v_���۽��, -v_���, ժҪ_In, ������_In, ��������_In, ������_In, ��������_In, v_ҩƷ�շ���¼.��ҩ��, v_ҩƷ�շ���¼.��ҩ����,
         v_ҩƷ�շ���¼.��ҩ��λid, v_ҩƷ�շ���¼.��׼�ĺ�, v_ҩƷ�շ���¼.��������, v_ҩƷ�շ���¼.����, v_ҩƷ�շ���¼.����, v_ҩƷ�շ���¼.Ƶ��);
    
      --����ҩƷ�������Ӧ����
      Zl_ҩƷ���_Update(v_�շ�id, 0, 0);
    
      --������ۺ����
      Zl_ҩƷ�շ���¼_��������(v_�շ�id);
    End Loop;
  End If;
Exception
  When Err_Isstriked Then
    Raise_Application_Error(-20101, '[ZLSOFT]�õ����Ѿ������˳�����[ZLSOFT]');
  When Err_Isbatch Then
    Raise_Application_Error(-20102, '[ZLSOFT]�õ����а�����һ��ԭ�������������ڷ�����ҩƷ[' || v_Druginf || ']�����ܳ�����[ZLSOFT]');
  When Err_Isnonum Then
    Raise_Application_Error(-20103, '[ZLSOFT]�õ����е�' || Ceil(���_In / 2) || '�е�ҩƷ����������������ʣ������ݣ����ܳ�����[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�ƿ�_Strike;
/

--117925:����,2017-12-07,��������������
Create Or Replace Procedure Zl_ҩƷ��������_Delete(No_In In ҩƷ�շ���¼.No%Type) Is
  Err_Isverified Exception;

  Cursor c_ҩƷ�շ���¼ Is
    Select ID From ҩƷ�շ���¼ Where NO = No_In And ���� = 11 Order By ҩƷid,����;
Begin

  --����ϵͳ����������ʱ���˿�����������Ҫ�ָ�ԭ���Ŀ�������
  For v_ҩƷ�շ���¼ In c_ҩƷ�շ���¼ Loop
    Zl_ҩƷ���_Update(v_ҩƷ�շ���¼.Id, 1);
  End Loop;

  Delete From ҩƷ�շ���¼ Where NO = No_In And ���� = 11 And ��¼״̬ = 1 And ����� Is Null;

  If Sql%RowCount = 0 Then
    Raise Err_Isverified;
  End If;
Exception
  When Err_Isverified Then
    Raise_Application_Error(-20101, '[ZLSOFT]�õ����Ѿ�������ɾ�����ѱ�����ˣ�[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ��������_Delete;
/

--117925:����,2017-12-07,��������������
Create Or Replace Procedure Zl_ҩƷ����_Delete(
                                           --ɾ��ҩƷ�շ���¼���ָ���Ӧ�ı�ҩƷ���
                                           No_In In ҩƷ�շ���¼.No%Type) Is
  Err_Isverified Exception;

  Cursor c_ҩƷ�շ���¼ Is
    Select ID, ��д����, �ⷿid, ����, ҩƷid, ����, Ч��, ����, ��׼�ĺ�, �Է�����id, ��ҩ��ʽ
    From ҩƷ�շ���¼
    Where NO = No_In And ���� = 7
    Order By ҩƷid,����;
  v_������������ Varchar2(4000);
Begin

  Select Zl_Getsysparameter('������������', 1305) Into v_������������ From Dual;
  --ͨ��ѭ�����ָ�ԭ���Ŀ�������
  For v_ҩƷ�շ���¼ In c_ҩƷ�շ���¼ Loop
    Zl_ҩƷ���_Update(v_ҩƷ�շ���¼.Id, 1);
  
    If v_ҩƷ�շ���¼.��ҩ��ʽ = 1 Then
      Update ҩƷ����
      Set �������� = Nvl(��������, 0) + v_ҩƷ�շ���¼.��д����
      Where �ڼ� = To_Char(Sysdate, Decode(v_������������, '1', 'yyyymm', 'yyyy')) And ����id = v_ҩƷ�շ���¼.�Է�����id And
            �ⷿid = v_ҩƷ�շ���¼.�ⷿid And ҩƷid = v_ҩƷ�շ���¼.ҩƷid;
    End If;
  End Loop;

  Delete From ҩƷ�շ���¼ Where NO = No_In And ���� = 7 And ��¼״̬ = 1 And ����� Is Null;

  If Sql%RowCount = 0 Then
    Raise Err_Isverified;
  End If;
Exception
  When Err_Isverified Then
    Raise_Application_Error(-20101, '[ZLSOFT]�õ����Ѿ�������ɾ�����ѱ�����ˣ�[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ����_Delete;
/

--116996:������,2017-12-06,����ҩȡ����������
Create Or Replace Procedure Zl_����ҽ����¼_�ջ�
(
  --���ܣ���ָ��ҽ�����ڷ��Ͳ����ջء�����ϴη���û�в������ã�����ջ�ҽ�����ϴ�ִ��ʱ�䡣
  --������
  --      �ջ���_IN=����ҩ���г�ҩΪ��סԺ��λ���ջ���,����ҩΪ�ջظ���,������ҽ��Ϊ�ջ������������
  --      ҽ��ID_IN=ÿ��Ҫ�ջص�ҽ����¼��ID(��ϸ�洢��ID),�Գ�ҩ���䷽,��һ��������ҩ;�����÷��巨(����Ϊ������δ��ȡ)
  --      �ϴ�ʱ��_IN=ҽ�����ڷ��Ͳ����ջغ�Ӧ�û�ԭ���ϴ�ִ��ʱ��(�ϸ�Ƶ�ʼ������),Ϊ��ʱ��ʾ��ȫ���ջ��ˡ�
  --      NO_IN=���ջ�Ҫ�����������ü�¼ʱ��Ϊ�����ɼ�¼�ĵ��ݺ�(�����ü�ҩƷʹ��),��ǰ�����ֻ����NO��һ���ݡ�
  --            ��ΪҩƷ���ܷ���,��������ڴ���ʱȡ��
  --            ���ȫ�ǻ��۵�������ֵΪ���������۵������򲻲����������ݣ�ֱ���޸Ļ�ɾ�����۵�
  �ջ���_In     ����ҽ������.��������%Type,
  ҽ��id_In     ����ҽ����¼.Id%Type,
  �ϴ�ʱ��_In   ����ҽ����¼.�ϴ�ִ��ʱ��%Type,
  �ջ�ʱ��_In   Date,
  No_In         סԺ���ü�¼.No%Type := Null,
  ����Ա���_In ��Ա��.���%Type := Null,
  ����Ա����_In ��Ա��.����%Type := Null
) Is
  --�ջ�ҽ����Ӧ�ķ��ͷ�����ϸ��ʣ������,��������ķ������ջ�
  --ʣ������û���ſ���������������ݣ��ڲ���������ʱ����ԭ��������
  --��ҩƷ�����ģ���һ�����������ܴ���δִ�к���ִ�в��֣���ֱ���д�����¼������δִ������
  --ִ�б�־=0-δִ��,1-��ִ�У�ҩƷ���в���ִ�У����շ���¼�е���ϸ������Ϊ׼����ҩƷ��ֻ���ȴ���δִ�е�
  Cursor c_Detail Is
    Select *
    From (With ҽ�����ü�¼ As (Select Max(Decode(b.��¼״̬, 2, 0, b.Id)) As ����id, b.No, Nvl(b.�۸񸸺�, b.���) As ���, b.�շ�ϸĿid,
                                 b.���˲���id, Sum(Nvl(b.����, 1) * b.����) As ʣ������, b.�շ����, Max(Nvl(b.ִ��״̬, 0)) As ִ��״̬, d.��������,
                                 c.�������, c.ҽ������, c.��������, Max(b.��¼״̬) As ��¼״̬, Max(b.�Ǽ�ʱ��) As �Ǽ�ʱ��, Nvl(e.�շѷ�ʽ, 0) As �շѷ�ʽ
                          From ����ҽ������ A, סԺ���ü�¼ B, ����ҽ����¼ C, �������� D, ����ҽ���Ƽ� E
                          Where a.ҽ��id = ҽ��id_In And a.No = b.No And a.��¼���� = b.��¼���� And a.ҽ��id = b.ҽ����� And
                                b.�۸񸸺� Is Null And b.�շ�ϸĿid = d.����id(+) And c.Id = ҽ��id_In And e.ҽ��id(+) = b.ҽ����� And
                                e.�շ�ϸĿid(+) = b.�շ�ϸĿid
                          Group By b.No, b.��¼����, Nvl(b.�۸񸸺�, b.���), b.�շ�ϸĿid, b.���˲���id, b.�շ����, d.��������, c.�������, c.ҽ������,
                                   c.��������, e.�շѷ�ʽ
                          Having Sum(Nvl(b.����, 1) * b.����) > 0)
           Select ����id, NO, ���, �շ�ϸĿid, ���˲���id, �շ����, ��������, �������, ҽ������, ��������, ʣ������, Null As ��ִ����, Null As δִ����,
                  ִ��״̬ As ִ�б�־, ��¼״̬, �Ǽ�ʱ��, �շѷ�ʽ
           From ҽ�����ü�¼
           Where �շ���� Not In ('5', '6', '7') And Not (�շ���� = '4' And Nvl(��������, 0) = 1)
           Union All
           Select a.����id, a.No, a.���, a.�շ�ϸĿid, a.���˲���id, a.�շ����, a.��������, a.�������, a.ҽ������, a.��������, a.ʣ������, 0 As ��ִ����,
                  Sum(Nvl(b.����, 1) * b.ʵ������) As δִ����, 0 As ִ�б�־, a.��¼״̬, Max(a.�Ǽ�ʱ��) As �Ǽ�ʱ��, �շѷ�ʽ
           From ҽ�����ü�¼ A, ҩƷ�շ���¼ B
           Where (a.�շ���� In ('5', '6', '7') Or (a.�շ���� = '4' And Nvl(a.��������, 0) = 1)) And a.����id = b.����id And
                 a.No = b.No And b.���� In (9, 10, 25, 26) And Mod(b.��¼״̬, 3) = 1 And b.����� Is Null
           Group By a.����id, a.No, a.���, a.��¼״̬, a.�շ�ϸĿid, a.���˲���id, a.�շ����, a.��������, a.�������, a.ҽ������, a.��������, a.ʣ������, �շѷ�ʽ
           Having Sum(Nvl(b.����, 1) * b.ʵ������) > 0
           Union All
           Select a.����id, a.No, a.���, a.�շ�ϸĿid, a.���˲���id, a.�շ����, a.��������, a.�������, a.ҽ������, a.��������, a.ʣ������,
                  Sum(Nvl(b.����, 1) * b.ʵ������) As ��ִ����, 0 As δִ����, 1 As ִ�б�־, a.��¼״̬, Max(a.�Ǽ�ʱ��) As �Ǽ�ʱ��, �շѷ�ʽ
           From ҽ�����ü�¼ A, ҩƷ�շ���¼ B
           Where (a.�շ���� In ('5', '6', '7') Or (a.�շ���� = '4' And Nvl(a.��������, 0) = 1)) And a.����id = b.����id And
                 a.No = b.No And b.���� In (9, 10, 25, 26) And Not (Mod(b.��¼״̬, 3) = 1 And b.����� Is Null)
           Group By a.����id, a.No, a.���, a.��¼״̬, a.�շ�ϸĿid, a.���˲���id, a.�շ����, a.��������, a.�������, a.ҽ������, a.��������, a.ʣ������, �շѷ�ʽ
           Having Sum(Nvl(b.����, 1) * b.ʵ������) > 0)
           Order By Decode(�������, '5', 0, '6', 0, '7', 0, �շ�ϸĿid), ִ�б�־, �Ǽ�ʱ�� Desc;


  Cursor c_Applay(v_����ids Varchar2) Is
    Select a.����id, b.No, b.���, a.����, a.����ʱ��, a.�������
    From ���˷������� A, סԺ���ü�¼ B
    Where a.����id = b.Id And a.���벿��id = a.��˲���id And a.����ʱ�� = �ջ�ʱ��_In And
          a.����id In (Select * From Table(Cast(f_Num2list(v_����ids) As Zltools.t_Numlist)))
    Order By NO, ���;

  --����ָ��ҩƷ��������ʱ��������ط��ü�ҩƷ/���ļ�¼��Ϣ(���η����ж�����¼,���������ڽ����ֹ)
  --ҩƷҽ����д��"����ҽ������"��¼,��Ӧ�ĸ�ҩ;����һ����д�˵�(����Ϊ����),��NO��ͬ��
  --��ΪҪ�ջصĴ������ܰ����˶�η��͵�����,����Ҫ����η��͵��շ���¼��ȡ��������η���ʱ�����۵����ջأ��޸Ļ�ɾ����
  Cursor c_Drug Is
    Select a.����id, a.��ҳid, d.����, Nvl(x.����ϵ��, 1) As ����ϵ��, Nvl(x.סԺ��װ, 1) As סԺ��װ, x.���Ч��, Nvl(b.����, 1) * b.ʵ������ As ����,
           b.Id As �շ�id, b.����, b.ҩƷid, b.�Է�����id, b.�ⷿid, b.����id, Nvl(x.ҩ������, 0) As ����, b.����, b.����, b.Ч��, a.��¼״̬, a.No,
           a.���, a.�շ�ϸĿid, a.ִ��״̬ As ִ�б�־
    From סԺ���ü�¼ A, ҩƷ�շ���¼ B, ����ҽ������ C, ������Ϣ D, ҩƷ��� X
    Where c.ҽ��id = ҽ��id_In And a.No = c.No And a.��¼���� = c.��¼���� And a.��¼״̬ In (0, 1, 3) And a.ҽ����� + 0 = ҽ��id_In And
          a.No = b.No And a.Id = b.����id + 0 And b.���� In (9, 10, 25, 26) And (b.��¼״̬ = 1 Or Mod(b.��¼״̬, 3) = 0) And
          a.����id = d.����id And b.ҩƷid = x.ҩƷid(+)
    Order By a.��¼״̬, b.No Desc, b.Id Desc;

  --������ҩ����(����ҩ;��)����ʱ�������ķ���(����������ж�����¼)
  --�Է�ҩҽ��,ֱ���ջ�ָ����,���ܶ�η���(�����η��ͼ۸�ͬ,���ջصļ۸��������εģ���Ȼ��Ҫ���ݶ���������μ��ջ���)��
  --���ı������ۼ۵�λ������סԺ��λת��
  --��ҩ��������д�˷��ͼ�¼(�����˶���������ȼ�)
  --һ��ֻ��һ�λ�һ�η���ֻ��һ�ε���Ŀ��ʱ��֧�ָ�������
  Cursor c_Other Is
    With ҽ�����ü�¼ As
     (Select a.No, a.���, a.��¼״̬, a.�շ�ϸĿid, a.Id As ����id, a.���� As ʣ������, Nvl(a.ִ��״̬, 0) As ִ��״̬, a.ҽ�����, b.���ͺ�,
             c.���� As ��������, Nvl(c.�շѷ�ʽ, 0) As �շѷ�ʽ, a.�շ����
      From סԺ���ü�¼ A, ����ҽ������ B, ����ҽ���Ƽ� C
      Where a.No = b.No And a.��¼���� = b.��¼���� And a.ҽ����� + 0 = b.ҽ��id And b.ҽ��id = ҽ��id_In And a.ҽ����� = c.ҽ��id(+) And
            a.�շ�ϸĿid = c.�շ�ϸĿid(+))
    Select a.No, a.���, a.����id, a.ʣ������, a.�շ�ϸĿid, a.��¼״̬, a.ִ��״̬, a.��������, a.�շѷ�ʽ, a.�շ����
    From (Select a.No, a.���, a.��¼״̬, a.�շ�ϸĿid, a.����id, a.ʣ������, a.��������, a.ִ��״̬, a.ҽ�����, a.�շѷ�ʽ, a.�շ����
           From ҽ�����ü�¼ A
           Where a.��¼״̬ In (1, 3) And a.���ͺ� = (Select Max(���ͺ�) From ҽ�����ü�¼ Where ��¼״̬ In (1, 3))
           Union All
           Select a.No, a.���, a.��¼״̬, a.�շ�ϸĿid, a.����id, a.ʣ������, a.��������, a.ִ��״̬, a.ҽ�����, a.�շѷ�ʽ, a.�շ����
           From ҽ�����ü�¼ A
           Where a.��¼״̬ = 0) A
    Order By a.�շ�ϸĿid, a.���, a.��¼״̬;

  --�����������Ϊ�˲����¼�¼ʱ,��дͬһ�շ�ϸĿ�Ĳ�ͬ������Ŀ�ļ۸񸸺�

  --���α����ڴ��������ػ��ܱ�
  Cursor c_Money
  (
    v_Start סԺ���ü�¼.���%Type,
    v_End   סԺ���ü�¼.���%Type
  ) Is
    Select ����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, Sum(Nvl(Ӧ�ս��, 0)) As Ӧ�ս��, Sum(Nvl(ʵ�ս��, 0)) As ʵ�ս��
    From סԺ���ü�¼
    Where ��¼���� = 2 And ��¼״̬ = 1 And NO = No_In And ��� Between v_Start And v_End
    Group By ����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid;

  --ϵͳ����ָ��ִ�к���Ҫ�Զ���˵Ļ��۷��ã����ڷ�ҩҽ����������Ӧ��ҩƷ�����ķ���
  Cursor c_Verify
  (
    v_Start סԺ���ü�¼.���%Type,
    v_End   סԺ���ü�¼.���%Type
  ) Is
    Select NO, ���
    From סԺ���ü�¼
    Where ��¼���� = 2 And ��¼״̬ = 0 And NO = No_In And �۸񸸺� Is Null And ��� Between v_Start And v_End;

  Cursor c_Compound
  (
    ���id_In       ����ҽ����¼.���id%Type,
    ִ����ֹʱ��_In ����ҽ����¼.ִ����ֹʱ��%Type,
    ��ҩid_In       ��Һ��ҩ��¼.Id%Type
  ) Is
    Select b.����id, b.ҩƷid As �շ�ϸĿid, Sum(a.����) As ����, c.סԺ��װ, c.סԺ��λ, d.����, e.���˲���id, e.����״̬, e.Id As ��ҩid, f.No,
           Nvl(f.�۸񸸺�, f.���) As ���, f.��¼״̬ As ��¼״̬, f.ִ��״̬ As ִ�б�־
    From ��Һ��ҩ���� A, ҩƷ�շ���¼ B, ҩƷ��� C, �շ���ĿĿ¼ D, ��Һ��ҩ��¼ E, סԺ���ü�¼ F
    Where a.�շ�id = b.Id And b.ҩƷid = c.ҩƷid And c.ҩƷid = d.Id And e.Id = a.��¼id And f.No = b.No And f.Id = b.����id And
          e.ҽ��id = ���id_In And e.ִ��ʱ�� > ִ����ֹʱ��_In And e.Id = ��ҩid_In
    Group By b.����id, b.ҩƷid, c.סԺ��װ, c.סԺ��λ, d.����, e.���˲���id, e.����״̬, e.Id, f.No, f.�۸񸸺�, f.���, f.��¼״̬, f.ִ��״̬;

  v_Dec      Number;
  v_First    Number;
  v_������� Varchar2(255);

  v_������� ����ҽ����¼.�������%Type;
  v_�������� ����ҽ����¼.��������%Type;
  v_�������� ��������.��������%Type;

  v_������� סԺ���ü�¼.���%Type;
  v_�շ���� ҩƷ�շ���¼.���%Type;
  v_����id   סԺ���ü�¼.Id%Type;
  v_ʵ�ս�� סԺ���ü�¼.ʵ�ս��%Type;

  v_��ʼ��� סԺ���ü�¼.���%Type;
  v_������� סԺ���ü�¼.���%Type;

  v_ҽ��ִ�� ����ҽ������.ִ��״̬%Type;

  v_����ϵ�� ҩƷ���.����ϵ��%Type;
  v_סԺ��װ ҩƷ���.סԺ��װ%Type;
  v_ҽ������ ����ҽ����¼.ҽ������%Type;

  v_���ʲ���       Zlparameters.����ֵ%Type;
  v_��Һҩ�������� Zlparameters.����ֵ%Type;
  v_���ʽ��       סԺ���ü�¼.���ʽ��%Type;

  v_�շ�ϸĿid   סԺ���ü�¼.�շ�ϸĿid%Type;
  v_ʣ������     סԺ���ü�¼.����%Type;
  v_�ջ�����     סԺ���ü�¼.����%Type;
  v_��ǰ����     סԺ���ü�¼.����%Type;
  v_��ǰ����     סԺ���ü�¼.����%Type;
  v_����ids      Varchar2(4000);
  v_��id         ����ҽ����¼.Id%Type;
  v_��������     ����ҽ���Ƽ�.����%Type;

  v_Delno    Varchar2(4000);
  v_Temp     Varchar2(4000);
  v_�շ����� Varchar2(4000);
  v_No       סԺ���ü�¼.No%Type;
  v_��Ա��� סԺ���ü�¼.����Ա���%Type;
  v_��Ա���� סԺ���ü�¼.����Ա����%Type;

  n_���id       ����ҽ����¼.���id%Type;
  d_ִ����ֹʱ�� ����ҽ����¼.ִ����ֹʱ��%Type;
  n_ҩƷid       ����ҽ����¼.�շ�ϸĿid%Type;
  b_��Һ��ҩ��¼ Boolean;
  d_�ջ�ʱ��     ����ҽ����¼.ִ����ֹʱ��%Type;
  n_�������     ���˷�������.�������%Type;
  n_Count        Number;

  v_Error Varchar2(255);
  Err_Custom Exception;

  Procedure �����շ���¼_Insert
  (
    ����id_In     Number,
    ����_In       ҩƷ�շ���¼.����%Type,
    ����_In       ҩƷ���.ҩ������%Type,
    ����_In       ҩƷ�շ���¼.����%Type,
    Ч��_In       ҩƷ�շ���¼.Ч��%Type,
    ���Ч��_In   ҩƷ���.���Ч��%Type,
    �շ�id_In     ҩƷ�շ���¼.Id%Type,
    ����id_In     סԺ���ü�¼.����id%Type,
    ��ҳid_In     סԺ���ü�¼.��ҳid%Type,
    ҩƷid_In     ҩƷ�շ���¼.ҩƷid%Type,
    �ⷿid_In     ҩƷ�շ���¼.�ⷿid%Type,
    ����_In       ҩƷ�շ���¼.����%Type,
    ����_In       ������Ϣ.����%Type,
    �Է�����id_In ҩƷ�շ���¼.�Է�����id%Type,
    �շ����_In   סԺ���ü�¼.�շ����%Type,
    �������_In   Varchar
  ) Is
    v_����   ҩƷ�շ���¼.����%Type;
    v_Ч��   ҩƷ�շ���¼.Ч��%Type;
    v_����   ҩƷ�շ���¼.����%Type;
    v_���ȼ� ���.���ȼ�%Type;
  Begin
    --ȷ������
    If Nvl(����_In, 0) <> 0 And ����_In = 0 Then
      --ԭ����,�ֲ�����
      v_���� := Null;
      v_���� := ����_In;
      v_Ч�� := Ч��_In;
    Elsif Nvl(����_In, 0) = 0 And ����_In = 1 Then
      --ԭ������,�ַ���
      Select ҩƷ�շ���¼_Id.Nextval Into v_���� From Dual;
      Select To_Char(Sysdate, 'YYYYMMDD') Into v_���� From Dual;
      If ���Ч��_In Is Not Null Then
        v_Ч�� := Trunc(Sysdate + ���Ч��_In * 30);
      Else
        v_Ч�� := Null;
      End If;
    Else
      v_���� := ����_In;
      v_���� := ����_In;
      v_Ч�� := Ч��_In;
    End If;
  
    Insert Into ҩƷ�շ���¼
      (ID, ��¼״̬, ����, NO, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ����, ��д����, ʵ������, ���ۼ�, ���۽��, ժҪ, ������, ��������,
       ����id, ����, Ƶ��, �÷�, ��ҩ��λid, ��������, ��׼�ĺ�, ���Ч��)
      Select ҩƷ�շ���¼_Id.Nextval, 1, ����, No_In, v_�շ����, �ⷿid, �Է�����id, ������id, -1, ҩƷid, v_����, ����, v_����, v_Ч��, v_��ǰ����,
             -1 * v_��ǰ����, -1 * v_��ǰ����, ���ۼ�, Round(-1 * v_��ǰ���� * v_��ǰ���� * ���ۼ�, v_Dec), '���ڷ����ջ�', v_��Ա����, �ջ�ʱ��_In, ����id_In,
             ����, Ƶ��, �÷�, ��ҩ��λid, ��������, ��׼�ĺ�, ���Ч��
      From ҩƷ�շ���¼
      Where ID = �շ�id_In;
  
    --ҩƷ���
    Update ҩƷ���
    Set �������� = Nvl(��������, 0) - (-1 * v_��ǰ���� * v_��ǰ����)
    Where �ⷿid = �ⷿid_In And ҩƷid = ҩƷid_In And Nvl(����, 0) = Nvl(v_����, 0) And ���� = 1;
    If Sql%RowCount = 0 Then
      Insert Into ҩƷ���
        (�ⷿid, ҩƷid, ����, ��������, ����, Ч��)
      Values
        (�ⷿid_In, ҩƷid_In, 1, v_��ǰ���� * v_��ǰ����, v_����, v_Ч��);
    End If;
  
    --δ��ҩƷ��¼
    Update δ��ҩƷ��¼
    Set ����id = ����id_In, ��ҳid = ��ҳid_In, ���� = ����_In
    Where ���� = ����_In And NO = No_In And �ⷿid + 0 = �ⷿid_In;
  
    If Sql%RowCount = 0 Then
      --ȡ������ȼ�
      Begin
        Select b.���ȼ� Into v_���ȼ� From ������Ϣ A, ��� B Where a.��� = b.����(+) And a.����id = ����id_In;
      Exception
        When Others Then
          Null;
      End;
    
      Insert Into δ��ҩƷ��¼
        (����, NO, ����id, ��ҳid, ����, ���ȼ�, �Է�����id, �ⷿid, ��������, ���շ�, ��ӡ״̬)
      Values
        (����_In, No_In, ����id_In, ��ҳid_In, ����_In, v_���ȼ�, �Է�����id_In, �ⷿid_In, �ջ�ʱ��_In,
         Decode(Nvl(Instr(�������_In, Decode(�շ����_In, '4', '4', '5')), 0), 0, 1, 0), 0);
    End If;
  
    v_�շ���� := v_�շ���� + 1;
  End;
Begin
  --ȡ����Ա��Ϣ(����ID,��������;��ԱID,��Ա���,��Ա����)
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
  Select ҽ������ Into v_ҽ������ From ����ҽ����¼ Where ID = ҽ��id_In;
  Select Count(1)
  Into n_Count
  From ��Һ��ҩ��¼ A, ����ҽ����¼ B
  Where a.ҽ��id = b.Id And ҽ��id = ҽ��id_In And a.ִ��ʱ�� > b.ִ����ֹʱ�� And a.�Ƿ����� = 1;

  If n_Count > 0 Then
    v_Error := 'ҽ��"' || v_ҽ������ || '"����ҺҩƷ���Ѿ�����Һ�����������������ܳ����ջء�';
    Raise Err_Custom;
  End If;

  If Nvl(�ջ���_In, 0) > 0 Then
    --�ж��Ƿ�����Һ��ҩҩƷ(��Һ��������ҩƷͳһ����������)
    b_��Һ��ҩ��¼ := False;
    Select a.���id, a.ִ����ֹʱ��, Max(b.�շ�ϸĿid)
    Into n_���id, d_ִ����ֹʱ��, n_ҩƷid
    From ����ҽ����¼ A, סԺ���ü�¼ B
    Where a.Id = ҽ��id_In And a.Id = b.ҽ�����(+)
    Group By a.���id, a.ִ����ֹʱ��;
    If n_���id Is Not Null Then
      If d_ִ����ֹʱ�� Is Not Null Then
        d_�ջ�ʱ��       := �ջ�ʱ��_In;
        v_��Һҩ�������� := zl_GetSysParameter('��Һ��Һ����ҩ��������������', 1345);
        Select Count(1) Into n_Count From ��Һ��ҩ��¼ E Where e.ҽ��id = n_���id And e.ִ��ʱ�� > d_ִ����ֹʱ��;
        If n_Count > 0 Then
          b_��Һ��ҩ��¼ := True;
          For X In (Select e.Id As ��ҩid, e.����״̬
                    From ��Һ��ҩ��¼ E
                    Where e.ҽ��id = n_���id And e.ִ��ʱ�� > d_ִ����ֹʱ�� And Nvl(e.����״̬, 0) In (1, 2, 3, 4, 5, 6, 7, 8)) Loop
            If Not (x.����״̬ In (4, 5, 6, 7, 8) And Nvl(v_��Һҩ��������, '0') = '0') Then
              For r_Compound In c_Compound(n_���id, d_ִ����ֹʱ��, x.��ҩid) Loop
                If x.����״̬ = 1 Then
                  n_������� := 0;
                Else
                  n_������� := 1;
                End If;
                Select Count(1)
                Into n_Count
                From ���˷�������
                Where ����id = r_Compound.����id And �շ�ϸĿid = r_Compound.�շ�ϸĿid And
                      ����ʱ�� =
                      (Select Max(����ʱ��) From ��Һ��ҩ״̬ A Where a.��ҩid = r_Compound.��ҩid And a.�������� = 9);
                If n_Count = 0 Then
                  Zl_���˷�������_Insert(r_Compound.����id, r_Compound.�շ�ϸĿid, r_Compound.���˲���id, r_Compound.����, v_��Ա����, d_�ջ�ʱ��,
                                   n_�������, Null, r_Compound.��ҩid);
                  If x.����״̬ = 1 Then
                    --δ��ҩ�ģ��Զ���ˡ�
                    Zl_���˷�������_Audit(r_Compound.����id, d_�ջ�ʱ��, v_��Ա����, d_�ջ�ʱ��, 1, 1, n_�������);
                    Zl_סԺ���ʼ�¼_Delete(r_Compound.No, r_Compound.��� || ':' || r_Compound.���� || ':' || r_Compound.��ҩid,
                                     v_��Ա���, v_��Ա����, 2, Null, Null, d_�ջ�ʱ��);
                  End If;
                End If;
              End Loop;
              --���ڲ�ͬ���Σ�ִ��ʱ�䣩����ʱ������ʱ��ͷ���ID��ΨһԼ��������ͬʱ���ʶ������ʱ�����μ�һ��
              d_�ջ�ʱ�� := d_�ջ�ʱ�� + 1 / 24 / 60 / 60;
            End If;
          End Loop;
        End If;
      End If;
    End If;
    --a.���������ջ�ģʽ
    --��Һ��ҩ��¼������������
    If b_��Һ��ҩ��¼ = False Then
      If No_In Is Null Then
        v_���ʲ��� := zl_GetSysParameter(23);
        --�����ջ���������ԭʼ���ý��з�̯����
        For r_Detail In c_Detail Loop
          --ȷ�����շ�ϸĿID���ջ�������
          If Nvl(v_�շ�ϸĿid, 0) <> r_Detail.�շ�ϸĿid And (r_Detail.������� Not In ('5', '6', '7') Or Nvl(v_�շ�ϸĿid, 0) = 0) Then
            --����δ��̯���
            If v_�շ�ϸĿid Is Not Null And v_�ջ����� > 0 Then
              v_Error := 'ҽ��"' || v_ҽ������ || '"��Ӧ�ķ���ʣ�����������ջ����������ܴ����ֹ����ʻ��ѽ��ʵķ��ã�����Ӧ�Ļ��۵��ѱ�ɾ����';
              Raise Err_Custom;
            End If;
            --ҩƷ�ջ�������������͹��Ϊ׼����ģ��Դ˼�����ջ��ۼ�����
            Begin
              Select ����ϵ��, סԺ��װ Into v_����ϵ��, v_סԺ��װ From ҩƷ��� Where ҩƷid = r_Detail.�շ�ϸĿid;
            Exception
              When Others Then
                Null;
            End;
            --�����ջ�����������շѷ�ʽ����0������ȡ����ʹ������ķ������м���
            If r_Detail.�շѷ�ʽ = 0 Then
              If r_Detail.������� = '7' Then
                --��ҩ�䷽ҩƷ������*����
                v_�ջ����� := Round(�ջ���_In * r_Detail.�������� / Nvl(v_����ϵ��, 1), 5);
              Else
                If r_Detail.������� Not In ('5', '6') Then
                  Select Nvl(Max(����), 1)
                  Into v_��������
                  From ����ҽ���Ƽ�
                  Where ҽ��id = ҽ��id_In And �շ�ϸĿid = r_Detail.�շ�ϸĿid;
                Else
                  v_�������� := 1;
                End If;
                v_�ջ����� := Round(�ջ���_In * Nvl(v_סԺ��װ, 1), 5) * v_��������;
              End If;
            Else
              Select Nvl(Sum(����), 0)
              Into v_�ջ�����
              From ҽ��ִ�мƼ�
              Where ҽ��id = ҽ��id_In And �շ�ϸĿid = r_Detail.�շ�ϸĿid And
                    Ҫ��ʱ�� > Nvl(�ϴ�ʱ��_In, To_Date('1900-01-01', 'yyyy-MM-dd'));
            
              v_�ջ����� := Round(v_�ջ�����, 5);
            
            End If;
            v_ҽ������ := r_Detail.ҽ������;
          End If;
        
          --���շ�ϸĿ��ÿ��������ϸ��̯�ջ�
          If v_�ջ����� > 0 Then
            --����Ӧ�����Ƿ��ѽ��ʣ�����ֹʱ
            v_���ʽ�� := 0;
            If v_���ʲ��� = '2' And r_Detail.��¼״̬ <> 0 Then
              Select Sum(���ʽ��)
              Into v_���ʽ��
              From סԺ���ü�¼
              Where NO = r_Detail.No And ��¼���� In (2, 12) And Nvl(�۸񸸺�, ���) = r_Detail.���;
            End If;
          
            If Nvl(v_���ʽ��, 0) = 0 Then
              If r_Detail.�շ���� In ('5', '6', '7') Or r_Detail.�շ���� = '4' And r_Detail.�������� = 1 Then
                --ҩƷ�͸������õ�����
                If r_Detail.ִ�б�־ = 0 Then
                  v_ʣ������ := r_Detail.δִ����;
                Elsif r_Detail.ִ�б�־ = 1 Then
                  v_ʣ������ := r_Detail.��ִ����;
                End If;
              Else
                --��ͨ����
                v_ʣ������ := r_Detail.ʣ������;
              End If;
              If v_�ջ����� > v_ʣ������ Then
                v_��ǰ���� := v_ʣ������;
              Else
                v_��ǰ���� := v_�ջ�����;
              End If;
              v_�ջ����� := v_�ջ����� - v_��ǰ����;
              --ϵͳ��������ִ�к��Ƿ���˻��۵������ԣ���ִ�е���Ȼ�����ǻ��۵�
              If r_Detail.ִ�б�־ = 0 And r_Detail.��¼״̬ = 0 Then
                v_Delno := v_Delno || '|' || r_Detail.No || ',' || r_Detail.��� || ':' || v_��ǰ����;
              Else
                If Not (r_Detail.�շ���� = '7' And r_Detail.ִ�б�־ <> 0) Then
                  Zl_���˷�������_Insert(r_Detail.����id, r_Detail.�շ�ϸĿid, r_Detail.���˲���id, v_��ǰ����, v_��Ա����, �ջ�ʱ��_In,
                                   r_Detail.ִ�б�־);
                End If;
              End If;
              v_����ids := v_����ids || ',' || r_Detail.����id;
            End If;
          End If;
          v_�շ�ϸĿid := r_Detail.�շ�ϸĿid;
        End Loop;
      
        --����δ��̯���
        If v_�ջ����� > 0 Then
          v_Error := 'ҽ��"' || v_ҽ������ || '"��Ӧ�ķ���ʣ�����������ջ����������ܴ����ֹ����ʻ��ѽ��ʵķ��ã�����Ӧ�Ļ��۵��ѱ�ɾ����';
          Raise Err_Custom;
        End If;
        --���Ƶ����������Զ����
        If zl_GetSysParameter('�����ջط��ñ����Զ����', 1254) = '1' And v_����ids Is Not Null Then
          For r_Applay In c_Applay(Substr(v_����ids, 2)) Loop
            Zl_���˷�������_Audit(r_Applay.����id, r_Applay.����ʱ��, v_��Ա����, �ջ�ʱ��_In, 1, 1, r_Applay.�������);
            v_Delno := v_Delno || '|' || r_Applay.No || ',' || r_Applay.��� || ':' || r_Applay.����;
          End Loop;
        End If;
      Else
        ---b.�����ջ�ģʽ-------------------------------------------------------------------------------------------------------
        --���ȫ�ǻ��۵����Ͳ��ò���������������
        If No_In = '�������۵�' Then
          --δ��˵Ļ��۵����Ƚ����޸Ļ�ɾ�������ܶ�η���Ϊ��ͬ��NO,Ϊ�˼���ÿ�ε��ջ�������Ҫ���շ�ϸĿID����
          For r_Price In (Select c.�������, b.No, b.���, b.�շ�ϸĿid, Nvl(b.����, 1) * b.���� As ʣ������, c.��������, d.����ϵ��, d.סԺ��װ,
                                 c.ҽ������, Nvl(e.�շѷ�ʽ, 0) As �շѷ�ʽ
                          From ����ҽ������ A, סԺ���ü�¼ B, ����ҽ����¼ C, ҩƷ��� D, ����ҽ���Ƽ� E
                          Where a.ҽ��id = ҽ��id_In And a.No = b.No And a.��¼���� = b.��¼���� And a.ҽ��id = b.ҽ����� And
                                b.�۸񸸺� Is Null And b.�շ�ϸĿid = d.ҩƷid(+) And b.��¼״̬ = 0 And c.Id = a.ҽ��id And
                                b.ҽ����� = e.ҽ��id(+) And b.�շ�ϸĿid = e.�շ�ϸĿid(+)
                          Order By �շ�ϸĿid, NO Desc) Loop
            If Nvl(v_�շ�ϸĿid, 0) <> r_Price.�շ�ϸĿid Then
              --����δ��̯���
              If v_�շ�ϸĿid Is Not Null And v_�ջ����� > 0 Then
                v_Error := 'ҽ��"' || v_ҽ������ || '"��Ӧ�ķ���ʣ�����������ջ�������������ػ��۵��ѱ�ɾ������ˡ�';
                Raise Err_Custom;
              End If;
              --�����ջ�����������շѷ�ʽ����0������ȡ����ʹ������ķ������м���
              If r_Price.�շѷ�ʽ = 0 Then
                If r_Price.������� = '7' Then
                  --��ҩ�䷽ҩƷ������*����
                  v_�ջ����� := Round(�ջ���_In * r_Price.�������� / Nvl(r_Price.����ϵ��, 1), 5);
                Else
                  If r_Price.������� Not In ('5', '6') Then
                    Select Nvl(Max(����), 1)
                    Into v_��������
                    From ����ҽ���Ƽ�
                    Where ҽ��id = ҽ��id_In And �շ�ϸĿid = r_Price.�շ�ϸĿid;
                  Else
                    v_�������� := 1;
                  End If;
                  v_�ջ����� := Round(�ջ���_In * Nvl(r_Price.סԺ��װ, 1), 5) * v_��������;
                End If;
              Else
                Select Nvl(Sum(����), 0)
                Into v_�ջ�����
                From ҽ��ִ�мƼ�
                Where ҽ��id = ҽ��id_In And �շ�ϸĿid = r_Price.�շ�ϸĿid And
                      Ҫ��ʱ�� > Nvl(�ϴ�ʱ��_In, To_Date('1900-01-01', 'yyyy-MM-dd'));
              
                v_�ջ����� := Round(v_�ջ�����, 5);
              End If;
              v_ҽ������ := r_Price.ҽ������;
            End If;
            If v_�ջ����� > 0 Then
              If v_�ջ����� > r_Price.ʣ������ Then
                v_��ǰ���� := r_Price.ʣ������;
              Else
                v_��ǰ���� := v_�ջ�����;
              End If;
              v_�ջ����� := v_�ջ����� - v_��ǰ����;
              v_Delno    := v_Delno || '|' || r_Price.No || ',' || r_Price.��� || ':' || v_��ǰ����;
            End If;
            v_�շ�ϸĿid := r_Price.�շ�ϸĿid;
          End Loop;
          --����δ��̯���
          If v_�ջ����� > 0 Then
            v_Error := 'ҽ��"' || v_ҽ������ || '"��Ӧ�ķ���ʣ�����������ջ�������������ػ��۵��ѱ�ɾ������ˡ�';
            Raise Err_Custom;
          End If;
        Else
          --�������������ܴ��ڻ��۵�����ʵ���ϵ����
          --���С��λ��
          Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into v_Dec From Dual;
          --���ɻ��۵�ϵͳ����
          Select zl_GetSysParameter(80) Into v_������� From Dual;
          v_��ʼ��� := Null;
          v_������� := Null;
        
          Select a.�������, a.��������, b.��������
          Into v_�������, v_��������, v_��������
          From ����ҽ����¼ A, �������� B
          Where ID = ҽ��id_In And a.�շ�ϸĿid = b.����id(+);
        
          If v_������� In ('5', '6', '7') Or (v_������� = '4' And Nvl(v_��������, 0) = 1) Then
            --ҩƷ������
            -----------------------------------------------------------------------------------------------------
            v_�ջ����� := Null;
            Select Nvl(Max(���), 0) + 1
            Into v_�շ����
            From ҩƷ�շ���¼
            Where ���� In (9, 10, 25, 26) And ��¼״̬ = 1 And NO = No_In;
            Select Nvl(Max(���), 0) + 1
            Into v_�������
            From סԺ���ü�¼
            Where ��¼���� = 2 And ��¼״̬ In (0, 1) And NO = No_In;
          
            --һ��ҽ����ҩƷֻ��һ�У������ѭ����Ϊ�˴����η��͵����������ҩƷ�ڽ����ѽ��ø����ջ�
            For r_Drug In c_Drug Loop
              --��ʼ��Ҫ�ջص�������(��������)
              v_First := 0;
              If v_�ջ����� Is Null Then
                If v_������� = '7' Then
                  v_�ջ����� := Round(�ջ���_In * v_�������� / r_Drug.����ϵ��, 5);
                Else
                  If v_������� Not In ('5', '6') Then
                    Select Nvl(Max(����), 1)
                    Into v_��������
                    From ����ҽ���Ƽ�
                    Where ҽ��id = ҽ��id_In And �շ�ϸĿid = r_Drug.�շ�ϸĿid;
                  Else
                    v_�������� := 1;
                  End If;
                  v_�ջ����� := Round(�ջ���_In * r_Drug.סԺ��װ, 5) * v_��������;
                End If;
                v_First := 1;
              End If;
            
              --�����һ���������㹻���򰴸����������������ô���
              If v_�ջ����� > r_Drug.���� Then
                v_��ǰ���� := 1;
                v_��ǰ���� := r_Drug.����;
                v_�ջ����� := v_�ջ����� - r_Drug.����;
              Else
                If v_First = 1 And v_������� = '7' Then
                  v_��ǰ���� := �ջ���_In;
                  v_��ǰ���� := Round(v_�������� / r_Drug.����ϵ��, 5);
                Else
                  v_��ǰ���� := 1;
                  v_��ǰ���� := v_�ջ�����;
                End If;
                v_�ջ����� := 0;
              End If;
            
              If r_Drug.��¼״̬ = 0 Then
                v_Delno := v_Delno || '|' || r_Drug.No || ',' || r_Drug.��� || ':' || v_��ǰ���� * v_��ǰ����;
              Else
                If Not (v_������� = '7' And r_Drug.ִ�б�־ <> 0) Then
                
                  Select ���˷��ü�¼_Id.Nextval Into v_����id From Dual;
                  �����շ���¼_Insert(v_����id, r_Drug.����, r_Drug.����, r_Drug.����, r_Drug.Ч��, r_Drug.���Ч��, r_Drug.�շ�id,
                                r_Drug.����id, r_Drug.��ҳid, r_Drug.ҩƷid, r_Drug.�ⷿid, r_Drug.����, r_Drug.����, r_Drug.�Է�����id,
                                v_�������, v_�������);
                
                  --סԺ���ü�¼
                  -------------------------------------------------------------------------------------
                  --��¼��ŷ�Χ�Դ�����ܱ�
                  If v_��ʼ��� Is Null Then
                    v_��ʼ��� := v_�������;
                  End If;
                  v_������� := v_�������;
                
                  Insert Into סԺ���ü�¼
                    (ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, �����־, ����id, ��ҳid, ��ʶ��, ����, �Ա�, ����, ����, ���˲���id, ���˿���id,
                     �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��,
                     ͳ����, ���ʷ���, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ��״̬, ҽ�����, ������, ����Ա���, ����Ա����)
                    Select v_����id, 2, No_In, Decode(Nvl(Instr(v_�������, Decode(v_�������, '4', '4', '5')), 0), 0, 1, 0),
                           v_�������, Null, Null, �ಡ�˵�, 2, ����id, ��ҳid, ��ʶ��, ����, �Ա�, ����, ����, ���˲���id, ���˿���id, �ѱ�, �շ����,
                           �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, v_��ǰ����, -1 * v_��ǰ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����,
                           Round(-1 * v_��ǰ���� * v_��ǰ���� * ��׼����, v_Dec), Round(-1 * v_��ǰ���� * v_��ǰ���� * ��׼����, v_Dec), Null, 1,
                           ��������id, ������, �ջ�ʱ��_In, �ջ�ʱ��_In, ִ�в���id, 0, ҽ�����, v_��Ա����,
                           Decode(Nvl(Instr(v_�������, Decode(v_�������, '4', '4', '5')), 0), 0, v_��Ա���, Null),
                           Decode(Nvl(Instr(v_�������, Decode(v_�������, '4', '4', '5')), 0), 0, v_��Ա����, Null)
                    From סԺ���ü�¼
                    Where ID = r_Drug.����id;
                
                  Select Zl_Actualmoney(�ѱ�, �շ�ϸĿid, ������Ŀid, Ӧ�ս��, ����, ִ�в���id)
                  Into v_Temp
                  From סԺ���ü�¼
                  Where ID = v_����id;
                  v_ʵ�ս�� := Round(Substr(v_Temp, Instr(v_Temp, ':') + 1), v_Dec);
                  Update סԺ���ü�¼ A Set ʵ�ս�� = v_ʵ�ս�� Where ID = v_����id;
                
                  v_������� := v_������� + 1;
                End If;
                If v_�ջ����� <= 0 Then
                  Exit;
                End If;
              End If;
            End Loop;
          
            If v_�ջ����� <> 0 Then
              --û���ջ���������,�շ���¼����������(���¼��ȫ������Ϊ��)
              Null;
            End If;
          Else
            --������ҩҽ��(������ҩ;�������󶨵����ĵ�)
            -----------------------------------------------------------------------------------------------------
            Select Nvl(Max(���), 0) + 1
            Into v_�շ����
            From ҩƷ�շ���¼
            Where ���� In (9, 10, 25, 26) And ��¼״̬ = 1 And NO = No_In;
            --ȡ�������
            Select Nvl(Max(���), 0) + 1
            Into v_�������
            From סԺ���ü�¼
            Where ��¼���� = 2 And ��¼״̬ In (0, 1) And NO = No_In;
          
            For r_Other In c_Other Loop
              If Nvl(v_�շ�����, '0') <> r_Other.�շ�ϸĿid || ',' || r_Other.��� Then
                --�������һ�η��͵ķ��ü�¼������Ҫ�ջص�����ȫ���ջ�
                --�����ջ�����������շѷ�ʽ����0������ȡ����ʹ������ķ������м���
                If r_Other.�շѷ�ʽ = 0 Then
                  v_�ջ����� := �ջ���_In * Nvl(r_Other.��������, 1);
                Else
                  Select Nvl(Sum(����), 0)
                  Into v_�ջ�����
                  From ҽ��ִ�мƼ�
                  Where ҽ��id = ҽ��id_In And �շ�ϸĿid = r_Other.�շ�ϸĿid And
                        Ҫ��ʱ�� > Nvl(�ϴ�ʱ��_In, To_Date('1900-01-01', 'yyyy-MM-dd'));
                End If;
              End If;
            
              If v_�ջ����� > 0 Then
                If r_Other.��¼״̬ = 0 Then
                  If v_�ջ����� > r_Other.ʣ������ Then
                    v_��ǰ���� := r_Other.ʣ������;
                  Else
                    v_��ǰ���� := v_�ջ�����;
                  End If;
                Else
                  v_��ǰ���� := v_�ջ�����;
                End If;
                v_�ջ����� := v_�ջ����� - v_��ǰ����;
                v_��ǰ���� := 1;
              
                If r_Other.��¼״̬ = 0 Then
                  v_Delno := v_Delno || '|' || r_Other.No || ',' || r_Other.��� || ':' || v_��ǰ����;
                Else
                  --��¼��ŷ�Χ�Դ�����ܱ�
                  If v_��ʼ��� Is Null Then
                    v_��ʼ��� := v_�������;
                  End If;
                  v_������� := v_�������;
                
                  --סԺ���ü�¼:��������ջ����������ϴη�����,����ȷ
                  Select ���˷��ü�¼_Id.Nextval Into v_����id From Dual;
                  If r_Other.�շ���� In ('4', '5', '6', '7') Then
                    For r_Otherdrug In (Select a.����id, a.��ҳid, d.����, Nvl(x.����ϵ��, 1) As ����ϵ��, Nvl(x.סԺ��װ, 1) As סԺ��װ,
                                               x.���Ч��, Nvl(b.����, 1) * b.ʵ������ As ����, b.Id As �շ�id, b.����, b.ҩƷid, b.�Է�����id,
                                               b.�ⷿid, b.����id, Nvl(x.ҩ������, 0) As ����, b.����, b.����, b.Ч��, a.��¼״̬, a.No, a.���,
                                               a.�շ�ϸĿid
                                        From סԺ���ü�¼ A, ҩƷ�շ���¼ B, ������Ϣ D, ҩƷ��� X
                                        Where a.Id = r_Other.����id And a.��¼״̬ In (0, 1, 3) And a.No = b.No And
                                              a.Id = b.����id + 0 And b.���� In (9, 10, 25, 26) And
                                              (b.��¼״̬ = 1 Or Mod(b.��¼״̬, 3) = 0) And a.����id = d.����id And
                                              b.ҩƷid = x.ҩƷid(+)
                                        Order By a.��¼״̬, b.No Desc, b.Id Desc) Loop
                      �����շ���¼_Insert(v_����id, r_Otherdrug.����, r_Otherdrug.����, r_Otherdrug.����, r_Otherdrug.Ч��,
                                    r_Otherdrug.���Ч��, r_Otherdrug.�շ�id, r_Otherdrug.����id, r_Otherdrug.��ҳid,
                                    r_Otherdrug.ҩƷid, r_Otherdrug.�ⷿid, r_Otherdrug.����, r_Otherdrug.����,
                                    r_Otherdrug.�Է�����id, r_Other.�շ����, v_�������);
                    End Loop;
                  End If;
                  --ҽ����ִ�У��ջصķ���Ҳ��Ϊ��ִ�У�������ҩƷ�͸������õ����ģ���Ϊʵ�ʷ��ű�ʾִ��
                  Insert Into סԺ���ü�¼
                    (ID, ��¼����, NO, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, �����־, ����id, ��ҳid, ��ʶ��, ����, �Ա�, ����, ����, ���˲���id, ���˿���id,
                     �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��,
                     ͳ����, ���ʷ���, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ��״̬, ִ��ʱ��, ִ����, ҽ�����, ������, ����Ա���, ����Ա����)
                    Select v_����id, 2, No_In, Decode(Nvl(Instr(v_�������, r_Other.�շ����), 0), 0, 1, 0), v_�������, Null,
                           Decode(a.�۸񸸺�, Null, Null, v_������� + a.�۸񸸺� - a.���), a.�ಡ�˵�, 2, a.����id, a.��ҳid, a.��ʶ��, a.����,
                           a.�Ա�, a.����, a.����, a.���˲���id, a.���˿���id, a.�ѱ�, a.�շ����, a.�շ�ϸĿid, a.���㵥λ, a.������Ŀ��, a.���մ���id, 1,
                           -1 * v_��ǰ����, a.�Ӱ��־, a.���ӱ�־, a.Ӥ����, a.������Ŀid, a.�վݷ�Ŀ, a.��׼����,
                           Round(-1 * v_��ǰ���� * a.��׼����, v_Dec), Round(-1 * v_��ǰ���� * a.��׼����, v_Dec), Null, 1, a.��������id,
                           a.������, �ջ�ʱ��_In, �ջ�ʱ��_In, a.ִ�в���id,
                           Decode(r_Other.ִ��״̬, 1,
                                   Decode(a.�շ����, '4', Decode(b.��������, 1, 0, 1), Decode(Instr(',5,6,7,', a.�շ����), 0, 1, 0)),
                                   0),
                           Decode(r_Other.ִ��״̬, 1,
                                   Decode(a.�շ����, '4', Decode(b.��������, 1, Null, �ջ�ʱ��_In),
                                           Decode(Instr(',5,6,7,', a.�շ����), 0, �ջ�ʱ��_In, Null)), Null),
                           Decode(r_Other.ִ��״̬, 1,
                                   Decode(a.�շ����, '4', Decode(b.��������, 1, Null, v_��Ա����),
                                           Decode(Instr(',5,6,7,', a.�շ����), 0, v_��Ա����, Null)), Null), a.ҽ�����, v_��Ա����,
                           Decode(Nvl(Instr(v_�������, r_Other.�շ����), 0), 0, v_��Ա���, Null),
                           Decode(Nvl(Instr(v_�������, r_Other.�շ����), 0), 0, v_��Ա����, Null)
                    From סԺ���ü�¼ A, �������� B
                    Where a.Id = r_Other.����id And a.�շ�ϸĿid = b.����id(+);
                
                  Select Zl_Actualmoney(�ѱ�, �շ�ϸĿid, ������Ŀid, Ӧ�ս��, ����, ִ�в���id)
                  Into v_Temp
                  From סԺ���ü�¼
                  Where ID = v_����id;
                  v_ʵ�ս�� := Round(Substr(v_Temp, Instr(v_Temp, ':') + 1), v_Dec);
                  Update סԺ���ü�¼ A Set ʵ�ս�� = v_ʵ�ս�� Where ID = v_����id;
                
                  v_������� := v_������� + 1;
                  v_ҽ��ִ�� := r_Other.ִ��״̬; --����շ���Ŀ��ִ��״̬��һ����
                End If;
              
                v_�շ����� := r_Other.�շ�ϸĿid || ',' || r_Other.���;
              End If;
            End Loop;
          
            --���ҽ����ִ�У���ϵͳ����ִ�к��Զ���˷��ã�������ִ��ҽ����Ӧ��ҩƷ�����ķ��á�
            -----------------------------------------------------------------------------------------------------
            If Nvl(v_ҽ��ִ��, 0) = 1 And v_��ʼ��� Is Not Null And v_������� Is Not Null Then
              If zl_GetSysParameter(81) = '1' Then
                For r_Verify In c_Verify(v_��ʼ���, v_�������) Loop
                  Zl_סԺ���ʼ�¼_Verify(r_Verify.No, v_��Ա���, v_��Ա����, r_Verify.���, Null, �ջ�ʱ��_In);
                End Loop;
              End If;
            End If;
          End If;
        
          --������û��ܱ�
          -----------------------------------------------------------------------------------------------------
          If v_��ʼ��� Is Not Null And v_������� Is Not Null Then
            --���ͳһ���������ػ��ܱ�
            For r_Money In c_Money(v_��ʼ���, v_�������) Loop
              --�������
              Update �������
              Set ������� = Nvl(�������, 0) + r_Money.ʵ�ս��
              Where ����id = r_Money.����id And ���� = 1 And ���� = 2;
            
              If Sql%RowCount = 0 Then
                Insert Into �������
                  (����id, ����, ����, �������, Ԥ�����)
                Values
                  (r_Money.����id, 1, 2, r_Money.ʵ�ս��, 0);
              End If;
            
              --����δ�����
              Update ����δ�����
              Set ��� = Nvl(���, 0) + r_Money.ʵ�ս��
              Where ����id = r_Money.����id And ��ҳid = r_Money.��ҳid And Nvl(���˲���id, 0) = Nvl(r_Money.���˲���id, 0) And
                    Nvl(���˿���id, 0) = Nvl(r_Money.���˿���id, 0) And Nvl(��������id, 0) = Nvl(r_Money.��������id, 0) And
                    Nvl(ִ�в���id, 0) = Nvl(r_Money.ִ�в���id, 0) And ������Ŀid + 0 = r_Money.������Ŀid And ��Դ;�� + 0 = 2;
            
              If Sql%RowCount = 0 Then
                Insert Into ����δ�����
                  (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
                Values
                  (r_Money.����id, r_Money.��ҳid, r_Money.���˲���id, r_Money.���˿���id, r_Money.��������id, r_Money.ִ�в���id,
                   r_Money.������Ŀid, 2, r_Money.ʵ�ս��);
              End If;
            End Loop;
          End If;
        End If;
      End If;
    End If;
  End If;

  --����Zl_סԺ���ʼ�¼_Delete����֧��ÿ��ɾ��һ�е�ѭ����������������������һ������Ҫɾ�������һ���Դ���
  If Not v_Delno Is Null Then
    v_Temp := '';
    v_No   := '';
    For r_Price In (Select /*+ rule*/
                     C1 As NO, C2 As �������
                    From Table(f_Str2list2(Substr(v_Delno, 2), '|', ','))
                    Order By NO) Loop
      If v_No Is Not Null And v_No <> r_Price.No Then
        Zl_סԺ���ʼ�¼_Delete(v_No, v_Temp, v_��Ա���, v_��Ա����, 2);
        v_No := '';
      End If;
      If v_No Is Null Then
        v_No   := r_Price.No;
        v_Temp := r_Price.�������;
      Else
        v_Temp := v_Temp || ',' || r_Price.�������;
      End If;
    End Loop;
    If Not v_No Is Null Then
      Zl_סԺ���ʼ�¼_Delete(v_No, v_Temp, v_��Ա���, v_��Ա����, 2);
    End If;
  End If;

  --����ҽ�����ϴ�ִ��ʱ��:��ҩ;���ȿ�����Ϊδ���Ͷ�û�����ջع��̡�
  -----------------------------------------------------------------------------------------------------
  Select Nvl(���id, ID) Into v_��id From ����ҽ����¼ Where ID = ҽ��id_In;
  Update ����ҽ����¼ Set �ϴ�ִ��ʱ�� = �ϴ�ʱ��_In Where ID = v_��id Or ���id = v_��id;

  --ɾ��ҽ��ִ��ʱ��
  If �ϴ�ʱ��_In Is Null Then
    --ȫ���ջ�
    Delete From ҽ��ִ��ʱ�� Where ҽ��id = v_��id;
    Delete From ҽ��ִ�мƼ� Where ҽ��id = ҽ��id_In;
  Else
    --�����ջض�η��͵�����
    Delete From ҽ��ִ��ʱ�� Where ҽ��id = v_��id And Ҫ��ʱ�� > �ϴ�ʱ��_In;
    Delete From ҽ��ִ�мƼ� Where ҽ��id = ҽ��id_In And Ҫ��ʱ�� > �ϴ�ʱ��_In;
  End If;
  --������Һ��Һ��¼���������⣬ÿ��ҽ�������е��ã��ڹ�������ֻ��������Һ��Һ��ҽ��
  Zl_��Һ��ҩ��¼_���ε���(ҽ��id_In);

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ҽ����¼_�ջ�;
/

--116996:������,2017-12-06,����ҩȡ����������
--116388:����,2017-11-16,����ȡ���������ʺ�ָ��������͵ĸ�ֵ����
Create Or Replace Procedure Zl_���˷�������_Delete
(
  Ids_In    In Varchar2,
  ��ҩid_In In ��Һ��ҩ��¼.Id%Type := Null
) As
  n_Id  ���˷�������.����id%Type;
  v_Ids Varchar2(4000);

  n_ҽ��id   סԺ���ü�¼.Id%Type;
  v_No       סԺ���ü�¼.No%Type;
  v_������Ա ��Һ��ҩ��¼.������Ա%Type;
  d_����ʱ�� ��Һ��ҩ��¼.����ʱ��%Type;
  n_�������� ��Һ��ҩ��¼.����״̬%Type;
  n_����ʱ�� ��Һ��ҩ��¼.����ʱ��%Type;
Begin
  If ��ҩid_In Is Not Null Then
    Select ����ʱ��
    Into n_����ʱ��
    From (Select ������Ա, ����ʱ��, ��������
           From ��Һ��ҩ״̬
           Where ��ҩid = ��ҩid_In And �������� = 9
           Order By ����ʱ�� Desc)
    Where Rownum = 1;
  End If;
  
  v_Ids := Ids_In || ',';
  While v_Ids Is Not Null Loop
    n_Id  := To_Number(Substr(v_Ids, 1, Instr(v_Ids, ',') - 1));
    v_Ids := Substr(v_Ids, Instr(v_Ids, ',') + 1);
  
    If n_����ʱ�� Is Null Then
      Delete ���˷������� Where ����id = n_Id And ״̬ = 0;
      Select a.No, a.ҽ����� Into v_No, n_ҽ��id From סԺ���ü�¼ A Where a.Id = n_Id;
      If Not n_ҽ��id Is Null Then
        --��δ�ṩ����ҩ����ȡ���Ĺ��ܣ����������������һ��ȡ��
        For R In (Select d.Id
                  From ����ҽ����¼ A, ����ҽ������ B, ��Һ��ҩ��¼ D
                  Where a.Id = n_ҽ��id And a.Id = b.ҽ��id And b.No = v_No And a.���id = d.ҽ��id And b.���ͺ� = d.���ͺ� And
                        b.��¼���� = 2) Loop
          Select ������Ա, ����ʱ��, ��������
          Into v_������Ա, d_����ʱ��, n_��������
          From (Select ������Ա, ����ʱ��, ��������
                 From ��Һ��ҩ״̬
                 Where ��ҩid = r.Id And �������� <> 9
                 Order By ����ʱ�� Desc, �������� Desc)
          Where Rownum = 1;        
          Update ��Һ��ҩ��¼ Set ������Ա = v_������Ա, ����ʱ�� = d_����ʱ��, ����״̬ = n_�������� Where ID = r.Id;
        End Loop;
      End If;
    Else
      Delete ���˷������� Where ����id = n_Id And ״̬ = 0 And ����ʱ�� = n_����ʱ��;    
      Select ������Ա, ����ʱ��, ��������
      Into v_������Ա, d_����ʱ��, n_��������
      From (Select ������Ա, ����ʱ��, ��������
             From ��Һ��ҩ״̬
             Where ��ҩid = ��ҩid_In And �������� <> 9
             Order By ����ʱ�� Desc, �������� Desc)
      Where Rownum = 1;    
      Update ��Һ��ҩ��¼ Set ������Ա = v_������Ա, ����ʱ�� = d_����ʱ��, ����״̬ = n_�������� Where ID = ��ҩid_In;    
    End If;  
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˷�������_Delete;
/
--117925:����,2017-12-06,��������������
Create Or Replace Procedure Zl_ҩƷ�⹺_Delete(No_In In ҩƷ�շ���¼.No%Type) Is
  Err_Isverified Exception;

  Cursor c_ҩƷ�շ���¼ Is
    Select ID From ҩƷ�շ���¼ Where NO = No_In And ���� = 1 Order By ҩƷid,����,���;
Begin

  --ͨ��ѭ�����ָ�ԭ���Ŀ�������
  For v_ҩƷ�շ���¼ In c_ҩƷ�շ���¼ Loop
    --���ÿ����¹���
    Zl_ҩƷ���_Update(v_ҩƷ�շ���¼.Id, 1);
  End Loop;

  Delete Ӧ����¼ Where ϵͳ��ʶ = 1 And �շ�id In (Select ID From ҩƷ�շ���¼ Where NO = No_In And ���� = 1);

  --��ӦӦ����¼��ɾ��ͨ������ɾ��
  Delete --ɾ������
  From ҩƷ�շ���¼
  Where NO = No_In And ���� = 1 And ��¼״̬ = 1 And ����� Is Null;

  If Sql%RowCount = 0 Then
    Raise Err_Isverified;
  End If;
Exception
  When Err_Isverified Then
    Raise_Application_Error(-20101, '[ZLSOFT]�õ����Ѿ�������ɾ�����ѱ�����ˣ�[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�⹺_Delete;
/

--117925:����,2017-12-06,��������������
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
    Order By a.ҩƷid,a.����;
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
    If Oldno_In Is Null Then
      Zl_ҩƷ���_Update(v_ҩƷ�շ���¼.Id, 0);
    Else
      Zl_ҩƷ���_Update(v_ҩƷ�շ���¼.Id, 0, 0, 0, 1);
    End If;
  
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
      Select distinct �ɱ���, ���ۼ�
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

--117925:����,2017-12-06,��������������
Create Or Replace Procedure Zl_ҩƷ�������_Delete(No_In In ҩƷ�շ���¼.No%Type) Is
  Err_Isverified Exception;

  Cursor c_ҩƷ�շ���¼ Is
    Select ID From ҩƷ�շ���¼ Where NO = No_In And ���� = 4 Order By ҩƷid,����,���;
Begin
  --ͨ��ѭ�����ָ�ԭ���Ŀ�������
  For v_ҩƷ�շ���¼ In c_ҩƷ�շ���¼ Loop
    Zl_ҩƷ���_Update(v_ҩƷ�շ���¼.Id, 1);
  End Loop;

  --ɾ��ҩƷ�շ���¼
  Delete --ɾ������
  From ҩƷ�շ���¼
  Where NO = No_In And ���� = 4 And ��¼״̬ = 1 And ����� Is Null;

  If Sql%RowCount = 0 Then
    Raise Err_Isverified;
  End If;
Exception
  When Err_Isverified Then
    Raise_Application_Error(-20101, '[ZLSOFT]�õ����Ѿ�������ɾ�����ѱ�����ˣ�[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�������_Delete;
/

--117925:����,2017-12-06,��������������
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
    Order By a.ҩƷid,a.����;
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
    Set �ɱ��� = v_ҩƷ�շ���¼.�ɱ���, �ϴ��ۼ� = Decode(v_ҩƷ�շ���¼.ʱ��, 1, v_ҩƷ�շ���¼.���ۼ�, Null),
        �ϴι�Ӧ��id = Decode(v_ҩƷ�շ���¼.��ҩ��λid, Null, �ϴι�Ӧ��id, v_ҩƷ�շ���¼.��ҩ��λid), �ϴ����� = v_ҩƷ�շ���¼.����, �ϴ��������� = v_ҩƷ�շ���¼.��������,
        �ϴβ��� = v_ҩƷ�շ���¼.����, �ϴ���׼�ĺ� = v_ҩƷ�շ���¼.��׼�ĺ�
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

--117925:����,2017-12-06,��������������
Create Or Replace Procedure Zl_Э�����_Delete(No_In In ҩƷ�շ���¼.No%Type) Is
  Err_Isverified Exception;

  Cursor c_ҩƷ�շ���¼ Is
    Select ID From ҩƷ�շ���¼ Where NO = No_In And ���� = 3 Order By ҩƷid,����;
Begin
  --ͨ��ѭ�����ָ����й���Э��ҩԭ���Ŀ�������
  For v_ҩƷ�շ���¼ In c_ҩƷ�շ���¼ Loop
    Zl_ҩƷ���_Update(v_ҩƷ�շ���¼.Id, 1);
  End Loop;

  Delete --ɾ��������Ӧ�Ĺ���Э��ҩ
  From ҩƷ�շ���¼
  Where NO = No_In And ���� = 3 And ��¼״̬ = 1 And ����� Is Null;

  If Sql%RowCount = 0 Then
    Raise Err_Isverified;
  End If;
Exception
  When Err_Isverified Then
    Raise_Application_Error(-20101, '[ZLSOFT]�õ����Ѿ�������ɾ�����ѱ�����ˣ�[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Э�����_Delete;
/

--117925:����,2017-12-06,��������������
Create Or Replace Procedure Zl_Э�����_Verify
(
  No_In     In ҩƷ�շ���¼.No%Type := Null,
  �����_In In ҩƷ�շ���¼.�����%Type := Null
) Is
  Err_Isverified Exception;
  v_������ ҩƷ���.ʵ�ʲ��%Type;
  v_�ɱ���   ҩƷ�շ���¼.�ɱ���%Type;
  v_�ɱ���� ҩƷ�շ���¼.�ɱ����%Type;

  Cursor c_ҩƷ�շ���¼ Is
    Select ID, ��д����, ���ۼ�, ���۽��, ���, �ⷿid, ҩƷid, ����, �ɱ���, ����, ����, ������id, ���ϵ��, �Է�����id, ��ҩ��λid, ��������, ��׼�ĺ�, Ч��
    From ҩƷ�շ���¼
    Where NO = No_In And ���� = 3 And ��¼״̬ = 1
	Order By ҩƷid,����;
Begin
  Update ҩƷ�շ���¼
  Set ����� = Nvl(�����_In, �����), ������� = Sysdate
  Where NO = No_In And ���� = 3 And ��¼״̬ = 1 And ����� Is Null;

  If Sql%RowCount = 0 Then
    Raise Err_Isverified;
  End If;

  For v_ҩƷ�շ���¼ In c_ҩƷ�շ���¼ Loop
    --����ҩƷ�������Ӧ����
    If v_ҩƷ�շ���¼.���ϵ�� = -1 Then
      v_�ɱ���   := Zl_Fun_Getoutcost(v_ҩƷ�շ���¼.ҩƷid, Nvl(v_ҩƷ�շ���¼.����, 0), v_ҩƷ�շ���¼.�ⷿid);
      v_�ɱ���� := v_�ɱ��� * v_ҩƷ�շ���¼.��д����;
      v_������ := v_ҩƷ�շ���¼.���۽�� - v_�ɱ����;
    Else
      Begin
        Select Sum(�ɱ���)
        Into v_�ɱ���
        From (Select Decode(Nvl(c.ʵ�ʽ��, 0), 0,
                              Decode(Nvl(c.�ϴβɹ���, 0), 0, Decode(Nvl(b.�ɱ���, 0), 0, (d.�ּ� - d.�ּ� * (b.ָ������� / 100)), b.�ɱ���),
                                      c.�ϴβɹ���), (d.�ּ� - d.�ּ� * (c.ʵ�ʲ�� / c.ʵ�ʽ��))) * (a.���� / a.��ĸ) As �ɱ���               
               From Э��ҩƷ���� A,
                    (Select b.ҩƷid, b.�ɱ���, b.ָ�������
                      From �շ���ĿĿ¼ A, ҩƷ��� B
                      Where a.Id = b.ҩƷid And Nvl(�Ƿ���, 0) = 0) B,
                    (Select �ⷿid, ҩƷid, ʵ�ʽ��, ʵ�ʲ��, �ϴβɹ���
                      From ҩƷ���
                      Where ���� = 1 And �ⷿid = v_ҩƷ�շ���¼.�Է�����id) C,
                    (Select �շ�ϸĿid, �ּ�
                      From �շѼ�Ŀ
                      Where ((Sysdate Between ִ������ And ��ֹ����) Or (Sysdate >= ִ������ And ��ֹ���� Is Null))) D
               Where a.Э��ҩƷid = b.ҩƷid And b.ҩƷid = d.�շ�ϸĿid And b.ҩƷid = c.ҩƷid(+) And a.ҩƷid = v_ҩƷ�շ���¼.ҩƷid
               Union All
               Select Decode(Nvl(c.ʵ�ʽ��, 0), 0,
                              Decode(Nvl(c.�ϴβɹ���, 0), 0, Decode(Nvl(b.�ɱ���, 0), 0, (c.�ּ� - c.�ּ� * (b.ָ������� / 100)), b.�ɱ���),
                                      c.�ϴβɹ���), (c.�ּ� - c.�ּ� * (c.ʵ�ʲ�� / c.ʵ�ʽ��))) * (a.���� / a.��ĸ) As �ɱ���               
               From Э��ҩƷ���� A,
                    (Select b.ҩƷid, b.�ɱ���, b.ָ�������
                      From �շ���ĿĿ¼ A, ҩƷ��� B
                      Where a.Id = b.ҩƷid And Nvl(�Ƿ���, 0) = 1) B,
                    (Select �ⷿid, ҩƷid, ʵ�ʽ��, ʵ�ʲ��, ʵ�ʽ�� / ʵ������ As �ּ�, �ϴβɹ���
                      From ҩƷ���
                      Where ���� = 1 And �ⷿid = v_ҩƷ�շ���¼.�Է�����id And ʵ������ > 0) C
               Where a.Э��ҩƷid = b.ҩƷid And b.ҩƷid = c.ҩƷid And a.ҩƷid = v_ҩƷ�շ���¼.ҩƷid);
      Exception
        When Others Then
          v_�ɱ��� := 0;
      End;
    
      v_�ɱ���� := v_�ɱ��� * v_ҩƷ�շ���¼.��д����;
      v_������ := v_ҩƷ�շ���¼.���۽�� - v_�ɱ����;
    End If;
  
    Update ҩƷ�շ���¼ Set �ɱ��� = v_�ɱ���, �ɱ���� = v_�ɱ����, ��� = v_������ Where ID = v_ҩƷ�շ���¼.Id;
  
    --���¿��
    Zl_ҩƷ���_Update(v_ҩƷ�շ���¼.Id, 0);
  
  End Loop;
Exception
  When Err_Isverified Then
    Raise_Application_Error(-20101, '[ZLSOFT]�õ����Ѿ���������ˣ�[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Э�����_Verify;
/

--117925:����,2017-12-06,��������������
Create Or Replace Procedure Zl_Э�����_Strike
(
  No_In     In ҩƷ�շ���¼.No%Type,
  �����_In In ҩƷ�շ���¼.�����%Type
) Is
  Err_Isstriked Exception;

  Cursor c_ҩƷ�շ���¼ Is
    Select ID, �ⷿid, ������id, ���ϵ��, ҩƷid, ��д����, ����, ʵ������, �ɱ���, ���۽��, ���, ����, ����, Ч��, ��ҩ��λid, ��������, ��׼�ĺ�
    From ҩƷ�շ���¼ A
    Where NO = No_In And ���� = 3 And ��¼״̬ = 2
	Order By ҩƷid,����;
Begin
  Update ҩƷ�շ���¼ Set ��¼״̬ = 3 Where NO = No_In And ���� = 3 And ��¼״̬ = 1;

  If Sql%RowCount = 0 Then
    Raise Err_Isstriked;
  End If;

  Insert Into ҩƷ�շ���¼
    (ID, ��¼״̬, ����, NO, ���, �ⷿid, ������id, �Է�����id, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ���۽��, ���, ժҪ,
     ������, ��������, �����, �������, ����id, ����, ��ҩ��λid, ��������, ��׼�ĺ�)
    Select ҩƷ�շ���¼_Id.Nextval, 2, ����, No_In, ���, �ⷿid, ������id, �Է�����id, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, -��д����, -ʵ������, �ɱ���,
           -�ɱ����, ���ۼ�, -���۽��, -���, ժҪ, �����_In, Sysdate, �����_In, Sysdate, ����id, ����, ��ҩ��λid, ��������, ��׼�ĺ�
    From ҩƷ�շ���¼
    Where NO = No_In And ���� = 3 And ��¼״̬ = 3;

  For v_ҩƷ�շ���¼ In c_ҩƷ�շ���¼ Loop
    --����ҩƷ�������Ӧ����
    Zl_ҩƷ���_Update(v_ҩƷ�շ���¼.Id, 0);
  
    --������ۺ����
    Zl_ҩƷ�շ���¼_��������(v_ҩƷ�շ���¼.Id);
  End Loop;
Exception
  When Err_Isstriked Then
    Raise_Application_Error(-20101, '[ZLSOFT]�õ����Ѿ������˳�����[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Э�����_Strike;
/

--117925:����,2017-12-06,��������������
Create Or Replace Procedure Zl_�������_Delete(
                                           
                                           --ɾ��ҩƷ�շ���¼����Ӧ�ı�
                                           No_In In ҩƷ�շ���¼.No%Type) Is
  Err_Isverified Exception;

  Cursor c_ҩƷ�շ���¼ Is
    Select ID From ҩƷ�շ���¼ Where NO = No_In And ���� = 2 Order By ҩƷid,����;
Begin

  --ͨ��ѭ�����ָ����й���ԭ��ҩԭ���Ŀ�������
  For v_ҩƷ�շ���¼ In c_ҩƷ�շ���¼ Loop
    --���¿��
    Zl_ҩƷ���_Update(v_ҩƷ�շ���¼.Id, 1);
  End Loop;

  Delete --ɾ��������Ӧ�Ĺ���ԭ��ҩ
  From ҩƷ�շ���¼
  Where NO = No_In And ���� = 2 And ��¼״̬ = 1 And ����� Is Null;

  If Sql%RowCount = 0 Then
    Raise Err_Isverified;
  End If;
Exception
  When Err_Isverified Then
    Raise_Application_Error(-20101, '[ZLSOFT]�õ����Ѿ�������ɾ�����ѱ�����ˣ�[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�������_Delete;
/

--117925:����,2017-12-06,��������������
Create Or Replace Procedure Zl_�������_Verify
(
  No_In     In ҩƷ�շ���¼.No%Type := Null,
  �����_In In ҩƷ�շ���¼.�����%Type := Null
) Is
  Err_Isverified Exception;
  v_�����     Number;
  v_������   ҩƷ���.ʵ�ʲ��%Type;
  v_�ɱ���     ҩƷ�շ���¼.�ɱ���%Type;
  v_�ɱ����   ҩƷ�շ���¼.�ɱ����%Type;
  v_�ɱ��۷�ʽ Zlparameters.����ֵ%Type;
  Intdigit     Number;

  Cursor c_ҩƷ�շ���¼ Is
    Select ID, ʵ������, Nvl(���ۼ�, 0) As ���ۼ�, ���۽��, ���, �ⷿid, ҩƷid, ����, �ɱ���, ����, Ч��, ����, ������id, ���ϵ��, �Է�����id, ��ҩ��λid, ��������,
           ��׼�ĺ�
    From ҩƷ�շ���¼
    Where NO = No_In And ���� = 2 And ��¼״̬ = 1
    Order By ҩƷid,����;
Begin
  Select Zl_To_Number(Nvl(Zl_Getsysparameter(9), '2')) Into Intdigit From Dual;
  Select Nvl(����ֵ, 0)
  Into v_�ɱ��۷�ʽ
  From Zlparameters
  Where ������ = 'ҩƷ�������ɱ��ۼ��㷽ʽ' And ģ�� = 1301;

  Update ҩƷ�շ���¼
  Set ����� = Nvl(�����_In, �����), ������� = Sysdate
  Where NO = No_In And ���� = 2 And ��¼״̬ = 1 And ����� Is Null;

  If Sql%RowCount = 0 Then
    Raise Err_Isverified;
  End If;

  For v_ҩƷ�շ���¼ In c_ҩƷ�շ���¼ Loop
    --����ҩƷ�������Ӧ����
    If v_ҩƷ�շ���¼.���ϵ�� = -1 Then
      v_�ɱ���   := Zl_Fun_Getoutcost(v_ҩƷ�շ���¼.ҩƷid, Nvl(v_ҩƷ�շ���¼.����, 0), v_ҩƷ�շ���¼.�ⷿid);
      v_�ɱ���� := Round(v_�ɱ��� * v_ҩƷ�շ���¼.ʵ������, Intdigit);
      v_������ := Round(v_ҩƷ�շ���¼.���۽�� - v_�ɱ����, Intdigit);
    Else
      If v_�ɱ��۷�ʽ = 0 Then
        Begin
          Select Sum(�ɱ���)
          Into v_�ɱ���
          From (Select Decode(Nvl(c.ʵ�ʽ��, 0), 0,
                                Decode(Nvl(c.�ϴβɹ���, 0), 0, Decode(Nvl(b.�ɱ���, 0), 0, (d.�ּ� - d.�ּ� * (b.ָ������� / 100)), b.�ɱ���),
                                        c.�ϴβɹ���), (d.�ּ� - d.�ּ� * (c.ʵ�ʲ�� / c.ʵ�ʽ��))) * (a.���� / a.��ĸ) * (e.����ϵ�� / b.������ϵ��) As �ɱ���
                 From ����ҩƷ���� A,
                      (Select b.ҩƷid, b.�ɱ���, b.ָ�������, b.����ϵ�� As ������ϵ��
                        From �շ���ĿĿ¼ A, ҩƷ��� B
                        Where a.Id = b.ҩƷid And Nvl(�Ƿ���, 0) = 0) B,
                      (Select �ⷿid, ҩƷid, ʵ�ʽ��, ʵ�ʲ��, �ϴβɹ���
                        From ҩƷ���
                        Where ���� = 1 And �ⷿid = v_ҩƷ�շ���¼.�Է�����id) C,
                      (Select �շ�ϸĿid, �ּ�
                        From �շѼ�Ŀ
                        Where ((Sysdate Between ִ������ And ��ֹ����) Or (Sysdate >= ִ������ And ��ֹ���� Is Null))) D, ҩƷ��� E
                 Where a.ԭ��ҩƷid = b.ҩƷid And b.ҩƷid = d.�շ�ϸĿid And b.ҩƷid = c.ҩƷid(+) And e.ҩƷid = v_ҩƷ�շ���¼.ҩƷid And
                       a.����ҩƷid = v_ҩƷ�շ���¼.ҩƷid
                 Union All
                 Select Decode(Nvl(c.ʵ�ʽ��, 0), 0,
                                Decode(Nvl(c.�ϴβɹ���, 0), 0, Decode(Nvl(b.�ɱ���, 0), 0, (c.�ּ� - c.�ּ� * (b.ָ������� / 100)), b.�ɱ���),
                                        c.�ϴβɹ���), (c.�ּ� - c.�ּ� * (c.ʵ�ʲ�� / c.ʵ�ʽ��))) * (a.���� / a.��ĸ) * (e.����ϵ�� / b.������ϵ��) As �ɱ���
                 From ����ҩƷ���� A,
                      (Select b.ҩƷid, b.�ɱ���, b.ָ�������, b.����ϵ�� As ������ϵ��
                        From �շ���ĿĿ¼ A, ҩƷ��� B
                        Where a.Id = b.ҩƷid And Nvl(�Ƿ���, 0) = 1) B,
                      (Select �ⷿid, ҩƷid, ʵ�ʽ��, ʵ�ʲ��, ʵ�ʽ�� / ʵ������ As �ּ�, �ϴβɹ���
                        From ҩƷ���
                        Where ���� = 1 And �ⷿid = v_ҩƷ�շ���¼.�Է�����id And ʵ������ > 0) C, ҩƷ��� E
                 Where a.ԭ��ҩƷid = b.ҩƷid And b.ҩƷid = c.ҩƷid And e.ҩƷid = v_ҩƷ�շ���¼.ҩƷid And a.����ҩƷid = v_ҩƷ�շ���¼.ҩƷid);
        Exception
          When Others Then
            v_�ɱ��� := 0;
        End;
        v_�ɱ���� := Round(v_�ɱ��� * v_ҩƷ�շ���¼.ʵ������, Intdigit);
        v_������ := Round(v_ҩƷ�շ���¼.���۽�� - v_�ɱ����, Intdigit);
      Else
        v_�ɱ���   := v_ҩƷ�շ���¼.�ɱ���;
        v_�ɱ���� := Round(v_ҩƷ�շ���¼.�ɱ��� * v_ҩƷ�շ���¼.ʵ������, Intdigit);
        v_������ := Round(v_ҩƷ�շ���¼.���۽�� - v_�ɱ����, Intdigit);
      End If;
    End If;
  
    Update ҩƷ�շ���¼ Set �ɱ��� = v_�ɱ���, �ɱ���� = v_�ɱ����, ��� = v_������ Where ID = v_ҩƷ�շ���¼.Id;
  
    --���ÿ����¹���
    Zl_ҩƷ���_Update(v_ҩƷ�շ���¼.Id, 0);
  
    If v_ҩƷ�շ���¼.���ϵ�� = 1 Then
      --ֻ����ҵ��Ŵ���
      --���¸�ҩƷ�ĳɱ���
      Update ҩƷ��� Set �ɱ��� = v_�ɱ��� Where ҩƷid = v_ҩƷ�շ���¼.ҩƷid;
    End If;
  
  End Loop;
Exception
  When Err_Isverified Then
    Raise_Application_Error(-20101, '[ZLSOFT]�õ����Ѿ���������ˣ�[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�������_Verify;
/

--117925:����,2017-12-06,��������������
Create Or Replace Procedure Zl_�������_Strike
(
  No_In     In ҩƷ�շ���¼.No%Type,
  �����_In In ҩƷ�շ���¼.�����%Type
) Is
  Err_Isstriked Exception;

  v_������id ҩƷ�շ���¼.������id%Type;

  Cursor c_ҩƷ�շ���¼ Is
    Select ID, �ⷿid, ������id, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ��д����, ʵ������, �ɱ���, ���۽��, ���, ��ҩ��λid, ��������, ��׼�ĺ�
    From ҩƷ�շ���¼ A
    Where NO = No_In And ���� = 2 And ��¼״̬ = 2
    Order By ҩƷid,����;
Begin
  Update ҩƷ�շ���¼ Set ��¼״̬ = 3 Where NO = No_In And ���� = 2 And ��¼״̬ = 1;

  If Sql%RowCount = 0 Then
    Raise Err_Isstriked;
  End If;

  Insert Into ҩƷ�շ���¼
    (ID, ��¼״̬, ����, NO, ���, �ⷿid, ������id, �Է�����id, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ��д����, ʵ������, �ɱ���, �ɱ����, ���ۼ�, ���۽��, ���, ժҪ,
     ������, ��������, �����, �������, ����id, ����, ��ҩ��λid, ��������, ��׼�ĺ�)
    Select ҩƷ�շ���¼_Id.Nextval, 2, 2, No_In, ���, �ⷿid, ������id, �Է�����id, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, -��д����, -ʵ������, �ɱ���,
           -�ɱ����, ���ۼ�, -���۽��, -���, ժҪ, �����_In, Sysdate, �����_In, Sysdate, ����id, ����, ��ҩ��λid, ��������, ��׼�ĺ�
    From ҩƷ�շ���¼
    Where NO = No_In And ���� = 2 And ��¼״̬ = 3;

  For v_ҩƷ�շ���¼ In c_ҩƷ�շ���¼ Loop
    --����ҩƷ�������Ӧ����
    Zl_ҩƷ���_Update(v_ҩƷ�շ���¼.Id, 0);
  
    --������ۺ����
    Zl_ҩƷ�շ���¼_��������(v_ҩƷ�շ���¼.Id);
  End Loop;
Exception
  When Err_Isstriked Then
    Raise_Application_Error(-20101, '[ZLSOFT]�õ����Ѿ������˳�����[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�������_Strike;
/

--115026:������,2017-12-04,����Σ��ֵ
CREATE OR REPLACE Procedure Zl_������Ϣ_������Ϣ����_ҽ��
(
  ����id_In ������Ϣ�䶯.����id%Type,
  ����id_In Number,
  ����_In   ������Ϣ.����%Type,
  �Ա�_In   ������Ϣ.�Ա�%Type,
  ����_In   ������Ϣ.����%Type,
  ����_In   Number, --1-����;2-סԺ
  ˵��_Out  Out ������Ϣ�䶯.˵��%Type
) As
  ------------------------------------------------------------------------------------------
  --����:����ҽ�����ҵ�����ݵĲ��˻�����Ϣ
  --���:����id_In:����ID
  --     ����id_In:���ﲡ��Ϊ�Һ�ID;סԺ����Ϊ��ҳID;Ϊ0˵����������ǹҺž���Ĳ���(����id_InΪ��ʱ,���������ĸò��˵�����ҵ������)
  --     ����_In:��Ҫ���ĵĲ�������
  --     �Ա�_In:��Ҫ���ĵĲ����Ա�
  --     ����_In:��Ҫ���ĵĲ�������
  --     ����_In:1-����;2-סԺ
  --����:˵��_Out:������Ϣ�������˵����Ϣ��������ʾ����Ա������ز���
  ------------------------------------------------------------------------------------------
  Err_Custom Exception;
  V_Error Varchar2(2000);
  N_Count Number(3);
  V_No    ���˹Һż�¼.No%Type;
  V_Tmp   Varchar2(100);
Begin
  --������Ա��������
  If Nvl(����id_In, 0) = 0 Then
    Return;
  End If;
  --����ȡ�Һŵ�
  If Nvl(����_In, 0) = 1 Then
    Select NO Into V_No From ���˹Һż�¼ Where ID = ����id_In;
    If V_No Is Null Then
      V_Error := '���ҵ��ò��˵ĹҺż�¼,���ܸ��²��˻�����Ϣ.';
      Raise Err_Custom;
    End If;
    --����ҽ��ǩ��,�������޸Ĳ��˻�����Ϣ
    Select Nvl(Count(1), 0)
    Into N_Count
    From ����ҽ����¼
    Where ����id = ����id_In And �Һŵ� = V_No And �¿�ǩ��id Is Not Null And Rownum < 2;
    If N_Count <> 0 Then
      V_Error := '����ҽ���Ѿ�ǩ��,���ܸ��²��˻�����Ϣ.';
      Raise Err_Custom;
    End If;

    --���²��˱��ξ����ҽ���еĲ��˻�����Ϣ
    Update ����ҽ����¼
    Set ���� = Nvl(����_In, ����), �Ա� = Nvl(�Ա�_In, �Ա�), ���� = Nvl(����_In, ����)
    Where ����id = ����id_In And �Һŵ� = V_No;
    Return;
  End If;
  --סԺ����
  If Nvl(����_In, 0) = 2 Then
    --סԺҽ��ǩ��,�������޸Ĳ��˻�����Ϣ
    Select Nvl(Count(1), 0)
    Into N_Count
    From ����ҽ����¼
    Where ����id = ����id_In And ��ҳid = ����id_In And �¿�ǩ��id Is Not Null And Rownum < 2;

    If N_Count <> 0 Then
      V_Error := '�ò���ҽ���Ѿ�ǩ��,���ܸ��²��˻�����Ϣ.';
      Raise Err_Custom;
    End If;
    --סԺ��ҳ����ǩ���ģ��������޸Ĳ��˻�����Ϣ
    Select Nvl(Count(1), 0)
    Into N_Count
    From ������ҳ�ӱ�
    Where ����id = ����id_In And ��ҳid = ����id_In And ��Ϣ�� In ('סԺҽʦǩ��', '����ҽʦǩ��', '����ҽʦǩ��', '������ǩ��') And Rownum < 2;
    If N_Count <> 0 Then
      V_Error := '�ò���סԺ��ҳ�Ѿ�ǩ��,���ܸ��²��˻�����Ϣ.';
      Raise Err_Custom;
    End If;
    --������������״̬���������޸Ĳ��˻�����Ϣ
    Select Decode(����״̬, 1, '�ȴ������', 3, '���������', 5, '�Ѿ����鵵', 10, '���մ�����', Null)
    Into V_Tmp
    From ������ҳ
    Where ����id = ����id_In And ��ҳid = ����id_In;

    If Not V_Tmp Is Null Then
      V_Error := '�ò��˵Ĳ���' || V_Tmp || ',���ܸ��²��˻�����Ϣ.';
      Raise Err_Custom;
    End If;
    --�������ڱ�Ŀ״̬���������޸Ĳ��˻�����Ϣ
    Select Nvl(Count(1), 0)
    Into N_Count
    From ������ҳ
    Where ����id = ����id_In And ��ҳid = ����id_In And ��Ŀ���� Is Not Null;
    If N_Count <> 0 Then
      V_Error := '�ò��˵Ĳ����Ѿ���Ŀ,���ܸ��²��˻�����Ϣ.';
      Raise Err_Custom;
    End If;

    --�Ѿ���ӡ��ҽ���嵥����ʾ���´�ӡ
    Select Nvl(Count(1), 0)
    Into N_Count
    From ����ҽ����ӡ
    Where ����id = ����id_In And ��ҳid = ����id_In And Rownum < 2;

    If N_Count <> 0 Then
      If Not ˵��_Out Is Null Then
        ˵��_Out := ˵��_Out || Chr(13);
      End If;
      ˵��_Out := ˵��_Out || 'ҽ���嵥:�Ѿ���ӡ�����´�ӡ.';
    End If;

    --�Ѿ���ӡ����ҳ����ʾ���´�ӡ
    Select Nvl(Count(1), 0)
    Into N_Count
    From ���Ӳ�����ӡ
    Where ����id = ����id_In And ��ҳid = ����id_In And �ļ�id Is Null And ���� = 9 And Rownum < 2;
    If N_Count <> 0 Then
      If Not ˵��_Out Is Null Then
        ˵��_Out := ˵��_Out || Chr(13);
      End If;
      ˵��_Out := ˵��_Out || '������ҳ:�Ѿ���ӡ�����´�ӡ.';
    End If;

    Update ����ҽ����¼
    Set ���� = Nvl(����_In, ����), �Ա� = Nvl(�Ա�_In, �Ա�), ���� = Nvl(����_In, ����)
    Where ����id = ����id_In And ��ҳid = ����id_In;

    Update ��Һ��ҩ��¼
    Set ���� = Nvl(����_In, ����), �Ա� = Nvl(�Ա�_In, �Ա�), ���� = Nvl(����_In, ����)
    Where ҽ��id In (Select ID From ����ҽ����¼ Where ����id = ����id_In And ��ҳid = ����id_In);

	  ---���²���Σ��ֵ��¼
    Update ����Σ��ֵ��¼
    Set ���� = Nvl(����_In, ����), �Ա� = Nvl(�Ա�_In, �Ա�), ���� = Nvl(����_In, ����)
    Where ����id = ����id_In And ��ҳid = ����id_In;

    Return;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || V_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_������Ϣ_������Ϣ����_ҽ��;
/

--115026:������,2017-12-04,����Σ��ֵ
Create Or Replace Procedure Zl_����Σ��ֵҽ��_Update
(
  ����_In     In Number,
  Σ��ֵid_In In ����Σ��ֵҽ��.Σ��ֵid%Type,
  ҽ��id_In   In ����Σ��ֵҽ��.ҽ��id%Type
) Is
  --���ܣ�Σ��ֵҽ�����ù�ϵ
  --����������_In-1������Ӧ��ϵ��2-ɾ����Ӧ��ϵ��3-ҽ������ʱɾ����ϵ
  n_����id ����ҽ����¼.����id%Type;
  n_��ҳid ����ҽ����¼.��ҳid%Type;
  v_�Һŵ� ����ҽ����¼.�Һŵ�%Type;

  n_Cnt   Number;
  v_Error Varchar2(2000);
  Err_Custom Exception;
Begin
  If ����_In = 1 Then
  
    --ֻ�ܹ���ͬһ�ξ����ҽ��
    Select a.����id, a.��ҳid, a.�Һŵ�
    Into n_����id, n_��ҳid, v_�Һŵ�
    From ����ҽ����¼ A, ����Σ��ֵ��¼ B
    Where a.Id = b.ҽ��id And b.Id = Σ��ֵid_In;      
    If v_�Һŵ� Is Null Then
      Select Count(1)
      Into n_Cnt
      From ����ҽ����¼ A
      Where a.Id = ҽ��id_In And a.����id = n_����id And a.��ҳid = n_��ҳid;
    Else
      Select Count(1) Into n_Cnt From ����ҽ����¼ A Where a.Id = ҽ��id_In And a.�Һŵ� = v_�Һŵ�;
    End If;
    If n_Cnt = 0 Then
      v_Error := 'ֻ�ܹ������ξ����ҽ����';
      Raise Err_Custom;
    End If;
  
    Insert Into ����Σ��ֵҽ�� (Σ��ֵid, ҽ��id) Values (Σ��ֵid_In, ҽ��id_In);
  Elsif ����_In = 2 Then
    Delete ����Σ��ֵҽ�� A Where a.Σ��ֵid = Σ��ֵid_In And a.ҽ��id = ҽ��id_In;
  Elsif ����_In = 3 Then
    Delete ����Σ��ֵҽ�� A Where a.ҽ��id = ҽ��id_In;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Σ��ֵҽ��_Update;
/

--115026:������,2017-12-04,����Σ��ֵ
CREATE OR REPLACE Procedure Zl_����Σ��ֵ��¼_����
(
  Id_In         In ����Σ��ֵ��¼.Id%Type,
  �������_In   In ����Σ��ֵ��¼.�������%Type,
  ȷ��ʱ��_In   In ����Σ��ֵ��¼.ȷ��ʱ��%Type,
  ȷ����_In     In ����Σ��ֵ��¼.ȷ����%Type,
  ȷ�Ͽ���id_In In ����Σ��ֵ��¼.ȷ�Ͽ���id%Type,
  �Ƿ�Σ��ֵ_In In ����Σ��ֵ��¼.�Ƿ�Σ��ֵ%Type
) Is
  --����Σ��ֵ���������״̬����Ϊ2��ʾҽ���Ѵ���
Begin
  Update ����Σ��ֵ��¼
  Set ������� = �������_In, ȷ��ʱ�� = ȷ��ʱ��_In, ȷ���� = ȷ����_In, ȷ�Ͽ���id = ȷ�Ͽ���id_In, �Ƿ�Σ��ֵ = �Ƿ�Σ��ֵ_In, ״̬ = 2
  Where ID = Id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Σ��ֵ��¼_����;
/

--115026:������,2017-12-04,����Σ��ֵ
CREATE OR REPLACE Procedure Zl_����Σ��ֵ��¼_Insert
(
  Id_In         In ����Σ��ֵ��¼.Id%Type,
  ������Դ_In   In ����Σ��ֵ��¼.������Դ%Type,
  ����id_In     In ����Σ��ֵ��¼.����id%Type,
  ��ҳid_In     In ����Σ��ֵ��¼.��ҳid%Type,
  �Һŵ�_In     In ����Σ��ֵ��¼.�Һŵ�%Type,
  Ӥ��_In       In ����Σ��ֵ��¼.Ӥ��%Type,
  ����_In       In ����Σ��ֵ��¼.����%Type,
  �Ա�_In       In ����Σ��ֵ��¼.�Ա�%Type,
  ����_In       In ����Σ��ֵ��¼.����%Type,
  ҽ��id_In     In ����Σ��ֵ��¼.ҽ��id%Type,
  �걾id_In     In ����Σ��ֵ��¼.�걾id%Type,
  Σ��ֵ����_In In ����Σ��ֵ��¼.Σ��ֵ����%Type,
  ����ʱ��_In   In ����Σ��ֵ��¼.����ʱ��%Type,
  �������id_In In ����Σ��ֵ��¼.�������id%Type,
  ������_In     In ����Σ��ֵ��¼.������%Type
) Is
--���ܣ�Σ��ֵ�Ǽǣ�����ʱ ״̬ ȱʡΪ1
Begin
  Insert Into ����Σ��ֵ��¼
    (ID, ������Դ, ����id, ��ҳid, �Һŵ�, Ӥ��, ����, �Ա�, ����, ҽ��id, �걾id, Σ��ֵ����, ����ʱ��, �������id, ������, ״̬)
  Values
    (Id_In, ������Դ_In, ����id_In, ��ҳid_In, �Һŵ�_In, Ӥ��_In, ����_In, �Ա�_In, ����_In, ҽ��id_In, �걾id_In, Σ��ֵ����_In, ����ʱ��_In,
     �������id_In, ������_In, 1);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Σ��ֵ��¼_Insert;
/

--115026:������,2017-12-04,����Σ��ֵ
Create Or Replace Procedure Zl_����Σ��ֵ��¼_Update
(
  Id_In         In ����Σ��ֵ��¼.Id%Type,
  ������Դ_In   In ����Σ��ֵ��¼.������Դ%Type,
  ����id_In     In ����Σ��ֵ��¼.����id%Type,
  ��ҳid_In     In ����Σ��ֵ��¼.��ҳid%Type,
  �Һŵ�_In     In ����Σ��ֵ��¼.�Һŵ�%Type,
  Ӥ��_In       In ����Σ��ֵ��¼.Ӥ��%Type,
  ����_In       In ����Σ��ֵ��¼.����%Type,
  �Ա�_In       In ����Σ��ֵ��¼.�Ա�%Type,
  ����_In       In ����Σ��ֵ��¼.����%Type,
  ҽ��id_In     In ����Σ��ֵ��¼.ҽ��id%Type,
  �걾id_In     In ����Σ��ֵ��¼.�걾id%Type,
  Σ��ֵ����_In In ����Σ��ֵ��¼.Σ��ֵ����%Type,
  ����ʱ��_In   In ����Σ��ֵ��¼.����ʱ��%Type,
  �������id_In In ����Σ��ֵ��¼.�������id%Type,
  ������_In     In ����Σ��ֵ��¼.������%Type
) Is
  --���ܣ��޸�Σ��ֵ��¼
  n_״̬  ����Σ��ֵ��¼.״̬%Type;
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  Select ״̬ Into n_״̬ From ����Σ��ֵ��¼ Where ID = Id_In;
  If Nvl(n_״̬, 0) <> 1 Then
    v_Error := '��ǰΣ��ֵ��¼�ѱ�ҽ������ȷ�ϣ������޸ġ�';
    Raise Err_Custom;
  End If;
  Update ����Σ��ֵ��¼
  Set ������Դ = ������Դ_In, ����id = ����id_In, ��ҳid = ��ҳid_In, �Һŵ� = �Һŵ�_In, Ӥ�� = Ӥ��_In, ���� = ����_In, �Ա� = �Ա�_In, ���� = ����_In,
      ҽ��id = ҽ��id_In, �걾id = �걾id_In, Σ��ֵ���� = Σ��ֵ����_In, ����ʱ�� = ����ʱ��_In, �������id = �������id_In, ������ = ������_In
  Where ID = Id_In;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Σ��ֵ��¼_Update;
/

--115026:������,2017-12-04,����Σ��ֵ
Create Or Replace Procedure Zl_����Σ��ֵ��¼_Delete(Id_In In ����Σ��ֵ��¼.Id%Type) Is
  n_״̬  ����Σ��ֵ��¼.״̬%Type;
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  Select ״̬ Into n_״̬ From ����Σ��ֵ��¼ Where ID = Id_In;
  If Nvl(n_״̬, 0) <> 1 Then
    v_Error := '��ǰΣ��ֵ��¼�ѱ�ҽ������ȷ�ϣ�����ɾ����';
    Raise Err_Custom;
  End If;
  Delete ����Σ��ֵ��¼ Where ID = Id_In;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Σ��ֵ��¼_Delete;
/

--117527:Ƚ����,2017-12-01,Ԥ�������ʹ�����ѿ�֧�����������δ�˻����ѿ�
Create Or Replace Procedure Zl_���˿������¼_�˿�
(
  ԭԤ��id_In ����Ԥ����¼.Id%Type,
  Ԥ��id_In   ����Ԥ����¼.Id%Type,
  �˿���_In ����Ԥ����¼.��Ԥ��%Type
) Is
  --���ܣ��˻����ѿ�
  --˵������Ԥ����������
  Err_Item Exception;
  v_Err_Msg Varchar2(500);

  n_���ڿ�Ƭ Number(3);
  n_���     ���ѿ�Ŀ¼.���%Type;
  n_������ ���ѿ�Ŀ¼.���%Type;
  d_ͣ������ ���ѿ�Ŀ¼.ͣ������%Type;
  d_����ʱ�� ���ѿ�Ŀ¼.����ʱ��%Type;
  v_����     ���ѿ�Ŀ¼.����%Type;

  n_Id       ���˿������¼.Id%Type;
  n_������ ���˿������¼.������%Type;
  n_���ν�� ���˿������¼.������%Type;
Begin
  n_������ := Nvl(�˿���_In, 0);
  For c_���� In (Select Distinct Max(Decode(a.��¼״̬, 2, 0, a.Id)) As ID, a.���ѿ�id,
                               Decode(Max(a.��¼״̬), 1, 2, Max(a.��¼״̬) + 2) As ��¼״̬, Sum(a.������) As ������
               From ���˿������¼ A, ���˿������¼ B, ���˿�������� C
               Where a.�ӿڱ�� = b.�ӿڱ�� And a.���ѿ�id = b.���ѿ�id And a.��� = b.��� And b.Id = c.������id And c.Ԥ��id = ԭԤ��id_In
               Group By a.�ӿڱ��, a.���ѿ�id, a.���
               Having Nvl(Sum(a.������), 0) > 0) Loop
  
    If c_����.������ < n_������ Then
      n_���ν�� := c_����.������;
      n_������ := n_������ - c_����.������;
    Else
      n_���ν�� := n_������;
      n_������ := 0;
    End If;
  
    --��鵱ǰ�����Ƿ��Ѿ�ʹ�� 
    Begin
      Select ����, 1, ͣ������, (Select Max(���) From ���ѿ�Ŀ¼ B Where a.���� = b.���� And a.�ӿڱ�� = b.�ӿڱ��), ���, ����ʱ��
      Into v_����, n_���ڿ�Ƭ, d_ͣ������, n_������, n_���, d_����ʱ��
      From ���ѿ�Ŀ¼ A
      Where ID = c_����.���ѿ�id;
    Exception
      When Others Then
        n_���ڿ�Ƭ := 0;
    End;
  
    --ȡ��ͣ�� 
    If n_���ڿ�Ƭ = 0 Then
      v_Err_Msg := '����Ϊ"' || v_���� || '"�����ѿ����ܱ�����ɾ�����������˷ѵ��ÿ���';
      Raise Err_Item;
    End If;
    If Nvl(n_���, 0) < Nvl(n_������, 0) Then
      v_Err_Msg := '�������˷ѵ���ʷ���ſ�(����Ϊ"' || v_���� || '")��';
      Raise Err_Item;
    End If;
    If Nvl(d_ͣ������, To_Date('3000-01-01', 'yyyy-mm-dd')) < To_Date('3000-01-01', 'yyyy-mm-dd') Then
      v_Err_Msg := '����Ϊ"' || v_���� || '"�����ѿ��Ѿ�������ͣ�ã������ٽ����˷ѣ�';
      Raise Err_Item;
    End If;
    If Nvl(d_����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) < To_Date('3000-01-01', 'yyyy-mm-dd') Then
      v_Err_Msg := '����Ϊ"' || v_���� || '"�����ѿ��Ѿ����գ������ٽ����˷ѣ�';
      Raise Err_Item;
    End If;
  
    Update ���ѿ�Ŀ¼ Set ��� = Nvl(���, 0) + n_���ν�� Where ID = c_����.���ѿ�id;
  
    Select ���˿������¼_Id.Nextval Into n_Id From Dual;
    Insert Into ���˿������¼
      (ID, �ӿڱ��, ���ѿ�id, ���, ��¼״̬, ���㷽ʽ, ������, ����, ������ˮ��, ����ʱ��, ��ע, �����־)
      Select n_Id, �ӿڱ��, ���ѿ�id, ���, c_����.��¼״̬, ���㷽ʽ, -1 * n_���ν��, ����, ������ˮ��, ����ʱ��, ��ע, 1
      From ���˿������¼
      Where ID = c_����.Id;
  
    Insert Into ���˿�������� (Ԥ��id, ������id) Values (Ԥ��id_In, n_Id);
  
    Update ���˿������¼ Set ��¼״̬ = 3 Where ID = c_����.Id;
  
    If n_������ = 0 Then
      Exit;
    End If;
  End Loop;

  If n_������ > 0 Then
    v_Err_Msg := '���ѿ�ʣ����˽��(' || LTrim(To_Char(�˿���_In - n_������, '9999999990.00')) || ')���㱾���˿���(' ||
                 LTrim(To_Char(�˿���_In, '9999999990.00')) || ')�������˷ѣ�';
    Raise Err_Item;
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˿������¼_�˿�;
/

--117279:��˶,2017-12-01,ͣ����Աʱ�����˻�
Create Or Replace Procedure Zl_��Ա��_����
(
Id_In In ��Ա��.Id%Type
) Is
  v_User �ϻ���Ա��.�û���%Type;
Begin
  Update ��Ա�� Set ����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd'), ����ԭ�� = '' Where ID = Id_In;
  Select Max(�û���) Into v_User From �ϻ���Ա�� Where ��Աid = Id_In;
  If Not v_User Is Null Then
    Begin
      --�������ݿ��û�
      Execute Immediate 'Alter User ' || v_User || '  Account UnLock';
    Exception
      When Others Then
        Null;
        --1�������û����ܲ����ڣ��ϼ���Ա��Ϊ���Ǽ�¼���û������������ݿ�ʵ�ʴ����û� 
      --2��ϵͳ������û��Ȩ�ޣ���ǰϵͳ������û��ALter UserȨ��
      --��˲�ȡ��������
    End;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Ա��_����;
/

--117279:��˶,2017-12-01,ͣ����Աʱ�����˻�
Create Or Replace Procedure Zl_��Ա��_Delete(Id_In In ��Ա��.Id%Type) Is
  v_User �ϻ���Ա��.�û���%Type;
Begin
  Select Max(�û���) Into v_User From �ϻ���Ա�� Where ��Աid = Id_In;
  Delete From ��Ա�� Where ID = Id_In;
  If Not v_User Is Null Then
    Begin
      --ͣ�����ݿ��û�
      Execute Immediate 'Alter User ' || v_User || '  Account Lock';
    Exception
      When Others Then
        Null;
        --1�������û����ܲ����ڣ��ϼ���Ա��Ϊ���Ǽ�¼���û������������ݿ�ʵ�ʴ����û� 
      --2��ϵͳ������û��Ȩ�ޣ���ǰϵͳ������û��ALter UserȨ��
      --��˲�ȡ��������
    End;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Ա��_Delete;
/

--117279:��˶,2017-12-01,ͣ����Աʱ�����˻�
Create Or Replace Procedure Zl_��Ա��_ͣ��
(
  Id_In       In ��Ա��.Id%Type,
  ����ԭ��_In ��Ա��.����ԭ��%Type := Null
) Is
  v_User �ϻ���Ա��.�û���%Type;
Begin
  Update ��Ա�� Set ����ʱ�� = Sysdate, ����ԭ�� = ����ԭ��_In Where ID = Id_In;
  Select Max(�û���) Into v_User From �ϻ���Ա�� Where ��Աid = Id_In;
  If Not v_User Is Null Then
    Begin
      --ͣ�����ݿ��û�
      Execute Immediate 'Alter User ' || v_User || '  Account Lock';
    Exception
      When Others Then
        Null;
        --1�������û����ܲ����ڣ��ϼ���Ա��Ϊ���Ǽ�¼���û������������ݿ�ʵ�ʴ����û� 
      --2��ϵͳ������û��Ȩ�ޣ���ǰϵͳ������û��ALter UserȨ��
      --��˲�ȡ��������
    End;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Ա��_ͣ��;
/

--107618:������,2017-11-29,��������
Create Or Replace Procedure Zl_���˽���_Cancel
(
  ����id_In ������Ϣ.����id%Type,
  No_In     ���˹Һż�¼.No%Type
) As
  v_�����   ������Ϣ.�����%Type;
  v_�Һ�id   ���˹Һż�¼.Id%Type;
  v_���﷽ʽ �ҺŰ���.���﷽ʽ%Type;
  n_�Һ�ģʽ Number(3);

  n_ת��       Number(1);
  n_�������id ����ת���¼.�������id%Type;
  v_����ҽ��   ����ת���¼.����ҽ��%Type;
Begin
  n_�Һ�ģʽ := To_Number(Nvl(Substr(zl_GetSysParameter(256), 1, 1), 0));

  Select ����� Into v_����� From ������Ϣ Where ����id = ����id_In;
  Select ID Into v_�Һ�id From ���˹Һż�¼ Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 1;

  --ȷ��ԭ�Һźű��ҽ��,���ڻ�ԭ
  Begin
    If Nvl(n_�Һ�ģʽ, 0) = 0 Then
      Select Nvl(a.���﷽ʽ, 0) Into v_���﷽ʽ From �ҺŰ��� A, ���˹Һż�¼ B Where a.���� = b.�ű� And b.No = No_In;
    Else
      Select Nvl(a.���﷽ʽ, 0)
      Into v_���﷽ʽ
      From �ٴ������¼ A, ���˹Һż�¼ B
      Where a.Id = b.�����¼id And b.No = No_In;
    End If;
  Exception
    When Others Then
      Null;
  End;

  --�жϲ����Ƿ���ת�﷽ʽ(ǿ������/ת��),����Ǹûػֵ���ǰ�� ���Һ�ҽ�� Ȼ����ת��䶯��¼
  For R In (Select a.�Һ�id, a.�������id, a.����ҽ��, a.���տ���id, a.����ҽ��, a.����ʱ��
            From ����ת���¼ A
            Where a.No = No_In
            Order By a.����ʱ�� Desc) Loop
    n_�������id := r.�������id;
    v_����ҽ��   := r.����ҽ��;
    n_ת��       := 1;
    Delete ����ת���¼ Where �Һ�id = r.�Һ�id And ����ʱ�� = r.����ʱ��;
    Exit;
  End Loop;

  --����״̬
  Update ������Ϣ Set ����ʱ�� = Null, ����״̬ = 1 Where ����id = ����id_In;

  Update ������ü�¼
  Set ִ��״̬ = 0, ִ��ʱ�� = Null, ��ҩ���� = Decode(v_���﷽ʽ, 0, Null, ��ҩ����), ���� = Null
  Where NO = No_In And ��¼���� = 4 And ��¼״̬ In (1, 3);

  If n_ת�� = 1 Then
    Update ���˹Һż�¼
    Set ִ�в���id = n_�������id, ִ���� = v_����ҽ��, ִ��״̬ = 0, ִ��ʱ�� = Null, ���� = Decode(v_���﷽ʽ, 0, Null, ����), ժҪ = Null
    Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 1;
  Else
    Update ���˹Һż�¼
    Set ִ��״̬ = 0, ִ��ʱ�� = Null, ���� = Decode(v_���﷽ʽ, 0, Null, ����), ժҪ = Null
    Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 1;
  End If;
  
  --ɾ�������������Ϣ
  Zl_���˹�����¼_Delete(����id_In, v_�Һ�id);
  Zl_������ϼ�¼_Delete(����id_In, v_�Һ�id, Null, Null, '1,11');
  Update �ŶӽкŶ��� Set �Ŷ�״̬ = 0 Where ҵ������ = 0 And ҵ��id = v_�Һ�id;

  Delete From ����ҽ����¼ Where ����id = ����id_In And �Һŵ� = No_In And ҽ��״̬ = 1;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˽���_Cancel;
/

--97553:������,2017-11-29,����Ƥ�Խ���Ǽ�ʱ��
Create Or Replace Procedure Zl_����ҽ����¼_Ƥ��
(
  --���ܣ���дҽ��Ƥ�Խ��
  --˵����ͬʱ�����˹�����¼
  --��������ע_In=�������ԣ�"(+)",���ԣ�"(-)",���ԣ�"����"����
  --      ���_IN=0-����,1-���ԣ�NULL=����
  Id_In         ����ҽ����¼.Id%Type,
  ��ע_In       ����ҽ����¼.Ƥ�Խ��%Type,
  ���_In       ���˹�����¼.���%Type,
  ����Ա����_In Varchar2 := Null,
  Ƥ��ʱ��_In   ���˹�����¼.����ʱ��%Type := Null,
  ������Ӧ_In   ���˹���ҩ��.������Ӧ%Type := Null
) Is
  --���ù���������ص�����ҩƷ��Ϣ��Ŀ
  Cursor c_Data Is
    Select Distinct c.����id, Decode(c.�Һŵ�, Null, c.��ҳid, d.Id) As ��ҳid, a.��Ŀid, b.����
    From �����÷����� A, ������ĿĿ¼ B, ����ҽ����¼ C, ���˹Һż�¼ D
    Where Nvl(a.����, 0) = 0 And a.�÷�id = c.������Ŀid And a.��Ŀid = b.Id And b.��� In ('5', '6') And c.�Һŵ� = d.No(+) And
          c.Id = Id_In;

  v_�Һŵ�   ����ҽ����¼.�Һŵ�%Type;
  v_״̬     ����ҽ����¼.ҽ��״̬%Type;
  v_ҽ������ ����ҽ����¼.ҽ������%Type;

  v_Date     Date;
  d_����ʱ�� Date;
  v_Temp     Varchar2(255);
  v_��Ա��� ��Ա��.���%Type;
  v_��Ա���� ����ҽ��״̬.������Ա%Type;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --���ҽ��״̬�Ƿ���ȷ:��������
  Select �Һŵ�, ҽ��״̬, ҽ������ Into v_�Һŵ�, v_״̬, v_ҽ������ From ����ҽ����¼ Where ID = Id_In;
  If v_״̬ = 4 Then
    v_Error := 'ҽ��"' || v_ҽ������ || '"�Ѿ����ϣ����ܵǼǹ�����������';
    Raise Err_Custom;
  End If;
  If v_�Һŵ� Is Not Null And v_״̬ = 1 And Not ���_In Is Null Then
    v_Error := 'ҽ��"' || v_ҽ������ || '"��δ���ͣ����ܵǼǹ�����������';
    Raise Err_Custom;
  End If;

  --��ǰ������Ա
  If ����Ա����_In Is Null Then
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_��Ա��� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
    v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  Else
    v_��Ա���� := ����Ա����_In;
    Select ��� Into v_��Ա��� From ��Ա�� Where ���� = v_��Ա����;
  End If;

  If Ƥ��ʱ��_In Is Null Then
    Select Sysdate Into v_Date From Dual;
    d_����ʱ�� := v_Date;
  Else
    v_Date := Ƥ��ʱ��_In;
    Select Sysdate Into d_����ʱ�� From Dual;
  End If;

  --����ҽ����¼:�������һ����¼
  Update ����ҽ����¼ Set Ƥ�Խ�� = ��ע_In, �걾��λ = To_Char(v_Date, 'YYYY-MM-DD HH24:MI:SS') Where ID = Id_In;
  Insert Into ����ҽ��״̬ (ҽ��id, ��������, ������Ա, ����ʱ��) Values (Id_In, 10, v_��Ա����, d_����ʱ��);

  --�Ǽǲ��˹�����¼(��ʹ��ǰ��ͬ��ҩ�Ĺ�������Ǽ�)
  If Not ���_In Is Null Then
    For r_Data In c_Data Loop
      Insert Into ���˹�����¼
        (ID, ����id, ��ҳid, ��¼��Դ, ҩ��id, ҩ����, ���, ����ʱ��, ��¼ʱ��, ��¼��, ������Ӧ)
      Values
        (���˹�����¼_Id.Nextval, r_Data.����id, r_Data.��ҳid, 2, r_Data.��Ŀid, r_Data.����, ���_In, v_Date, d_����ʱ��, v_��Ա����, ������Ӧ_In);
      If ���_In = 1 Then
        Update ���˹���ҩ��
        Set ������Ӧ = ������Ӧ_In, ����ҩ��id = r_Data.��Ŀid
        Where ����id = r_Data.����id And ����ҩ�� = r_Data.����;
        If Sql%RowCount = 0 Then
          Insert Into ���˹���ҩ��
            (����id, ����ҩ��id, ����ҩ��, ������Ӧ)
          Values
            (r_Data.����id, r_Data.��Ŀid, r_Data.����, ������Ӧ_In);
        End If;
      Else
        --���û�й����ļ�¼��ɾ����ҩƷ�Ĺ�����¼
        Delete From ���˹���ҩ�� A
        Where a.����id = r_Data.����id And a.����ҩ�� = r_Data.���� And a.����ҩ��id = r_Data.��Ŀid And Not Exists
         (Select 1
               From ���˹�����¼ B
               Where b.����id = a.����id And b.ҩ��id = a.����ҩ��id And b.ҩ���� = a.����ҩ�� And ��� = 1);
      End If;
    End Loop;
    --���Ƥ�Խ��ʱ��ҽ���Զ���Ϊִ�����
    For X In (Select ִ��״̬, ���ͺ�, ִ�в���id From ����ҽ������ Where ҽ��id = Id_In) Loop
      If x.ִ��״̬ <> 1 Then
        Zl_����ҽ��ִ��_Finish(Id_In, x.���ͺ�, Null, 0, v_��Ա���, v_��Ա����, x.ִ�в���id);
      End If;
    End Loop;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ҽ����¼_Ƥ��;
/

--116388:����,2017-11-27,����ȡ����ҩ������ȷ�ָ�����״̬�Ĵ���
Create Or Replace Procedure Zl_��Һ��ҩ��¼_ȡ����ҩ(��ҩid_In In Varchar2 --ID��:ID1,ID2....
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
  n_Row Number(10);
  n_Out Number(1);

Begin
  v_Usercode := Zl_Identity;
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ';') + 1);
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ',') + 1);
  v_Usercode := Substr(v_Usercode, 1, Instr(v_Usercode, ',') - 1);
  n_Out      := Nvl(zl_GetSysParameter('��Ժ���˲������÷�', 1345), 0);

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
    
      If n_����״̬ != 4 Then
        v_Error := '�����ݵ�ǰ������ҩ״̬�����ܽ���ȡ����ҩ��';
        Raise Err_Custom;
      End If;
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
  
    --��[��Һ��ҩ״̬]���м�¼��ȡ����ҩ���Ĳ���
    Insert Into ��Һ��ҩ״̬
      (��ҩid, ��������, ������Ա, ����ʱ��, ����˵��)
    Values
      (v_Tansid, 2, v_������Ա, Sysdate, 'ȡ����ҩ');
  
    Select �Ƿ��� Into n_��� From ��Һ��ҩ��¼ Where ID = v_Tansid;
    If n_��� <> 1 Then
      For r_Item In (Select a.No, b.���
                     From ��Һ��ҩ���� A, סԺ���ü�¼ B
                     Where a.����id = b.����id And a.No = b.No And b.��¼״̬ = 1 And a.��ҩid = v_Tansid) Loop
        If r_Item.No Is Not Null Then
          Zl_סԺ���ʼ�¼_Delete(r_Item.No, r_Item.���, v_Usercode, Zl_Username);
        End If;
      End Loop;
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

--116388:����,2017-11-27,����ȡ����ҩ������ȷ�ָ�����״̬�Ĵ���
Create Or Replace Procedure Zl_��Һ��ҩ��¼_ȡ����ҩ(��ҩid_In In Varchar2 --ID��:��ҩID1,��ҩID2....
                                           ) Is
  v_Id       Varchar2(20);
  v_��ҩid   Varchar2(20);
  v_��ҩid   Varchar2(20);
  v_�շ�id   Varchar2(20);
  v_��ҩid   Varchar2(20);
  v_Tmp      Varchar2(4000);
  v_Date     Date;
  v_������Ա ��Һ��ҩ��¼.������Ա%Type;
  d_����ʱ�� ��Һ��ҩ��¼.����ʱ��%Type;
  n_����״̬ ��Һ��ҩ��¼.����״̬%Type;
  v_Error    Varchar2(255);
  Err_Custom Exception;

  Cursor c_��ҩ���� Is
    Select /*+ rule*/
    Distinct c.��¼id, a.Id As ��ҩid, c.�շ�id
    From ҩƷ�շ���¼ A, ҩƷ�շ���¼ B, ��Һ��ҩ���� C, Table(Cast(f_Num2list(��ҩid_In) As Zltools.t_Numlist)) D
    Where c.��¼id = d.Column_Value And b.Id = c.�շ�id And a.���� = b.���� And a.No = b.No And a.�ⷿid + 0 = b.�ⷿid And
          a.ҩƷid + 0 = b.ҩƷid And a.��� = b.��� And (a.��¼״̬ = 1 Or Mod(a.��¼״̬, 3) = 0);

  v_��ҩ���� c_��ҩ����%RowType;

  Cursor c_��ҩ��¼ Is
    Select a.Id, a.����, a.Ч��, a.����, b.���� As ��ҩ��
    From ҩƷ�շ���¼ A, ��Һ��ҩ���� B
    Where a.Id = v_��ҩid And a.����� Is Not Null And b.�շ�id = v_�շ�id And b.��¼id = v_��ҩid;

  v_��ҩ��¼ c_��ҩ��¼%RowType;

Begin
  If ��ҩid_In Is Null Then
    v_Tmp := Null;
  Else
    v_Tmp := ��ҩid_In || ',';
  End If;
  While v_Tmp Is Not Null Loop
    --�ֽⵥ��ID�� 
    v_Id  := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp := Replace(',' || v_Tmp, ',' || v_Id || ',');
  
    --��鵱ǰ��Һ����״̬�Ƿ�Ϊ����ҩ״̬ 
    Begin
      Select ����״̬ Into n_����״̬ From ��Һ��ҩ��¼ Where ID = v_Id;
    
      If n_����״̬ != 2 Then
        v_Error := '�������ѱ����������ܽ���ȡ����ҩ������';
        Raise Err_Custom;
      End If;
    Exception
      When Others Then
        v_Error := '�������ѱ�������';
        Raise Err_Custom;
    End;
  
    Select ������Ա, ����ʱ��
    Into v_������Ա, d_����ʱ��
    From ��Һ��ҩ״̬
    Where ��ҩid = v_Id And �������� = 1 And Rownum = 1;
  
    Update ��Һ��ҩ��¼ Set ����״̬ = 1, ������Ա = v_������Ա, ����ʱ�� = d_����ʱ�� Where ID = v_Id;
  
    --��[��Һ��ҩ״̬]���м�¼��ȡ����ҩ���Ĳ���
    Insert Into ��Һ��ҩ״̬
      (��ҩid, ��������, ������Ա, ����ʱ��, ����˵��)
    Values
      (v_Id, 1, v_������Ա, Sysdate, 'ȡ����ҩ');
  
  End Loop;

  Select Sysdate Into v_Date From Dual;

  For v_��ҩ���� In c_��ҩ���� Loop
    v_��ҩid := v_��ҩ����.��ҩid;
    v_�շ�id := v_��ҩ����.�շ�id;
    v_��ҩid := v_��ҩ����.��¼id;
    For v_��ҩ��¼ In c_��ҩ��¼ Loop
      --������ҩ 
      Zl_ҩƷ�շ���¼_������ҩ(v_��ҩ��¼.Id,
                     Zl_Username,
                     v_Date,
                     v_��ҩ��¼.����,
                     v_��ҩ��¼.Ч��,
                     v_��ҩ��¼.����,
                     v_��ҩ��¼.��ҩ��,
                     Null,
                     Zl_Username);
    
      Select Max(a.Id)
      Into v_��ҩid
      From ҩƷ�շ���¼ A, ҩƷ�շ���¼ B
      Where b.Id = v_��ҩid And a.���� = b.���� And a.No = b.No And a.�ⷿid + 0 = b.�ⷿid And a.ҩƷid + 0 = b.ҩƷid And
            a.��� = b.��� And Mod(a.��¼״̬, 3) = 1 And a.������� Is Null;
    
      --�滻��Һ��ҩ�����е��շ�ID 
      Update ��Һ��ҩ���� Set �շ�id = v_��ҩid Where ��¼id = v_��ҩid And �շ�id = v_��ҩ����.�շ�id;
    End Loop;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Һ��ҩ��¼_ȡ����ҩ;
/

--116388:����,2017-11-27,����ȡ����ҩ������ȷ�ָ�����״̬�Ĵ���
Create Or Replace Procedure Zl_��Һ��ҩ��¼_ȡ������(��ҩid_In In Varchar2 --ID��:ID1,ID2....
                                           ) Is
  v_Tansid   Varchar2(20);
  v_Tmp      Varchar2(4000);
  v_������Ա ��Һ��ҩ��¼.������Ա%Type;
  d_����ʱ�� ��Һ��ҩ��¼.����ʱ��%Type;
  n_���     Number(2);

  v_Error    Varchar2(255);
  n_����״̬ ��Һ��ҩ��¼.����״̬%Type;
  v_Usercode Varchar2(100);
  Err_Custom Exception;
Begin
  v_Usercode := Zl_Identity;
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ';') + 1);
  v_Usercode := Substr(v_Usercode, Instr(v_Usercode, ',') + 1);
  v_Usercode := Substr(v_Usercode, 1, Instr(v_Usercode, ',') - 1);

  If ��ҩid_In Is Null Then
    v_Tmp := Null;
  Else
    v_Tmp := ��ҩid_In || ',';
  End If;
  While v_Tmp Is Not Null Loop
    --�ֽⵥ��ID��
    v_Tansid := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
    v_Tmp    := Replace(',' || v_Tmp, ',' || v_Tansid || ',');
  
    --��鵱ǰ��Һ����״̬�Ƿ�Ϊ�ѷ���״̬
    Begin
      Select ����״̬ Into n_����״̬ From ��Һ��ҩ��¼ Where ID = v_Tansid;
    
      If n_����״̬ != 5 Then
        v_Error := '�������ѱ����������ܽ���ȡ�����Ͳ�����';
        Raise Err_Custom;
      End If;
    Exception
      When Others Then
        v_Error := '�������ѱ�������';
        Raise Err_Custom;
    End;
  
    Select ������Ա, ����ʱ��
    Into v_������Ա, d_����ʱ��
    From (Select ������Ա, ����ʱ�� From ��Һ��ҩ״̬ Where ��ҩid = v_Tansid And �������� = 4 Order By ����ʱ�� Desc)
    Where Rownum = 1;
  
    Update ��Һ��ҩ��¼ Set ����״̬ = 4, ������Ա = v_������Ա, ����ʱ�� = d_����ʱ�� Where ID = v_Tansid;
  
    --��[��Һ��ҩ״̬]���м�¼��ȡ�����͡��Ĳ���
    Insert Into ��Һ��ҩ״̬
      (��ҩid, ��������, ������Ա, ����ʱ��, ����˵��)
    Values
      (v_Tansid, 4, v_������Ա, Sysdate, 'ȡ������');
  
    Select �Ƿ��� Into n_��� From ��Һ��ҩ��¼ Where ID = v_Tansid;
    If n_��� <> 0 Then
      For r_Item In (Select a.No, b.���
                     From ��Һ��ҩ���� A, סԺ���ü�¼ B
                     Where a.����id = b.����id And a.No = b.No And b.��¼״̬ = 1 And a.��ҩid = v_Tansid) Loop
        If r_Item.No Is Not Null Then
          Zl_סԺ���ʼ�¼_Delete(r_Item.No, r_Item.���, v_Usercode, Zl_Username);
        End If;
      End Loop;
    End If;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Һ��ҩ��¼_ȡ������;
/

--92026:��͢��,2017-11-22,��������¼��ҽ������ִ��
CREATE OR REPLACE Procedure Zl_����ҽ��ִ��_Insert
( 
  ҽ��id_In       ����ҽ��ִ��.ҽ��id%Type, 
  ���ͺ�_In       ����ҽ��ִ��.���ͺ�%Type, 
  Ҫ��ʱ��_In     ����ҽ��ִ��.Ҫ��ʱ��%Type, 
  ��������_In     ����ҽ��ִ��.��������%Type, 
  ִ��ժҪ_In     ����ҽ��ִ��.ִ��ժҪ%Type, 
  ִ����_In       ����ҽ��ִ��.ִ����%Type, 
  ִ��ʱ��_In     ����ҽ��ִ��.ִ��ʱ��%Type, 
  ����ִ��_In     Number := 0, 
  �Զ����_In     Number := 0, 
  ִ�н��_In     ����ҽ��ִ��.ִ�н��%Type := 1, 
  δִ��ԭ��_In   ����ҽ��ִ��.˵��%Type := Null, 
  ����Ա���_In   ��Ա��.���%Type := Null, 
  ����Ա����_In   ��Ա��.����%Type := Null, 
  ִ�в���id_In   ������ü�¼.ִ�в���id%Type := 0, 
  ��Һ���_In     Number := 0, 
  ������Ŀ����_In Number := 0, 
  ��Һͨ��_In     ����ҽ��ִ��.��Һͨ��%Type := Null 
  --������ҽ��ID_IN=����ִ�е�ҽ��ID���������Ϊ��ʾ�ļ�����Ŀ��ID�� 
  --      ִ�н��_In=1- ���   =0  -δִ�� 
  --      �����̨ʽ������ ����Ա���_In ����Ա����_In �������������봫�� 
  --ִ�в���id_In=������ָ��ִ�в��ŵķ��ã���������0ʱ������ִ�в��� 
  --��Һ���_In=�ƶ�����վ����ʱ���Ƿ�����Һ��Ϣ�� 
  --������Ŀ����_In=����Ǽ�����Ŀʱ����Ҫ���ʵ������ҽ������״̬ 
) Is 
  --����Ҫִ�е�����¼,�������˸�������,��鲿λ�ļ�¼ 
  --��������,��ҩ�巨,�ɼ�������������,�������ֻ��д�ڵ�һ����Ŀ�ϣ���ִ��״̬��ͬ 
  v_��id     ����ҽ����¼.Id%Type; 
  v_������� ����ҽ����¼.�������%Type; 
  v_�Զ���� Number; 
  v_������Դ ����ҽ����¼.������Դ%Type; 
  v_�������� ����ҽ������.��¼����%Type; 
  v_�������� ������ĿĿ¼.��������%Type; 
  v_����id   ������ҳ.��ǰ����id%Type; 
  v_��Һ���� Varchar2(200); 
  v_Count    Number; 
  v_Temp     Varchar2(255); 
  v_��Ա��� ��Ա��.���%Type; 
  v_��Ա���� ��Ա��.����%Type; 
  n_��Ч     ����ҽ����¼.ҽ����Ч%Type; 
  n_������Ŀid ����ҽ����¼.������Ŀid%Type;
  v_����ִ��   Varchar2(5);
 
  n_ִ�д��� Number; 
  n_ʣ����� Number; 
  n_ִ��״̬ Number; 
  d_��ֹʱ�� Date; 
  d_��ʼʱ�� Date; 
  n_�������� Number;
  n_�Ǽ����� Number;
  n_�������� Number;
  d_Ҫ��ʱ�� Date;
 
  v_Date  Date; 
  v_Error Varchar2(255); 
  Err_Custom Exception; 
Begin 
  --������죬��ֹ��������ִ�м�¼ 
  Begin 
    Select (a.�������� - c.�ǼǴ���) As ʣ������, a.��������, Nvl(D.������Ŀid, 0)
    Into v_Count, n_��������, n_������Ŀid
    From ����ҽ������ A, 
         (Select ҽ��id_In As ҽ��id, ���ͺ�_In As ���ͺ�, Nvl(Sum(b.��������), 0) As �ǼǴ��� 
           From ����ҽ��ִ�� B 
           Where b.ҽ��id = ҽ��id_In And b.���ͺ� = ���ͺ�_In) C, ����ҽ����¼ D
    Where a.ҽ��id = c.ҽ��id And a.���ͺ� = c.���ͺ� And a.ҽ��id = ҽ��id_In And a.ҽ��id = d.Id And a.���ͺ� = ���ͺ�_In;
  Exception 
    When Others Then 
      v_Count := ��������_In; 
  End; 
  v_����ִ�� := zl_GetSysParameter(288);
  If ��������_In > v_Count And (Not (n_������Ŀid = 0 And v_����ִ�� = 1)) Then
    v_Error := '���ڲ������������Ѿ������˵Ǽǣ���ˢ�º����ԡ�'; 
    Raise Err_Custom; 
  End If; 
  --��ǰ������Ա 
  If ����Ա���_In Is Not Null And ����Ա����_In Is Not Null Then 
    v_��Ա��� := ����Ա���_In; 
    v_��Ա���� := ����Ա����_In; 
  Else 
    Begin 
      Select ����, ��� Into v_��Ա����, v_��Ա��� From ��Ա�� Where ���� = ִ����_In; 
    Exception 
      When Others Then 
        v_Temp     := Zl_Identity; 
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1); 
        v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1); 
        v_��Ա��� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1); 
        v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1); 
    End; 
  End If; 
  --��ҽ����ֹʱ����м�� 
  Select a.ִ����ֹʱ��, a.��ʼִ��ʱ��, a.ҽ����Ч 
  Into d_��ֹʱ��, d_��ʼʱ��, n_��Ч 
  From ����ҽ����¼ A 
  Where a.Id = ҽ��id_In; 
  If Not d_��ֹʱ�� Is Null And n_��Ч = 0 Then 
    If Ҫ��ʱ��_In > d_��ֹʱ�� Then 
      v_Error := 'Ҫ��ʱ�䳬����ҽ����ֹʱ�䣬��ȷ��ҽ���Ƿ���ǰֹͣ��'; 
      Raise Err_Custom; 
    End If; 
  End If; 
  If Not d_��ʼʱ�� Is Null Then 
    If ִ��ʱ��_In < d_��ʼʱ�� Then 
      v_Error := 'ִ��ʱ��������ҽ���Ŀ�ʼִ��ʱ��''' || To_Char(d_��ʼʱ��, 'yyyy-mm-dd HH24:mi:ss') || '''��'; 
      Raise Err_Custom; 
    End If; 
  End If; 
  Select Sysdate Into v_Date From Dual; 
  Select a.������Դ, ִ�п���id, Nvl(a.���id, a.Id), Nvl(a.�������, '*'), Nvl(b.��������, '0') ��������
  Into v_������Դ, v_����id, v_��id, v_�������, v_��������
  From ����ҽ����¼ A, ������ĿĿ¼ B
  Where a.Id = ҽ��id_In And a.������Ŀid = b.Id(+);

  If v_������Դ = 2 Then 
    Select Decode(��¼����, 1, 1, Decode(�������, 1, 1, 2)) 
    Into v_�������� 
    From ����ҽ������ 
    Where ���ͺ� = ���ͺ�_In And ҽ��id = ҽ��id_In; 
  Else 
    v_�������� := 1; 
  End If; 
 
  --�ƶ�ϵͳ��Һ��� 
  If ��Һ���_In = 1 Then 
    --��鵱ǰ�������������Ƿ������Һ�Ǽǹ��� 
    Select Nvl(Zl_Getsysparameter(184), '') Into v_��Һ���� From Dual; 
 
    If v_��Һ���� Is Not Null And ִ�н��_In <> 0 Then 
      If Instr(',' || v_��Һ���� || ',', ',' || v_����id || ',') > 0 Then 
        v_����id   := 0; 
        v_��Һ���� := 'Select 1 From ������Һ��¼ where ҽ��ID=:YZID AND ���ͺ�=:FSH AND Ҫ��ʱ��=:YQSJ'; 
        Begin 
          Execute Immediate v_��Һ���� 
            Into v_����id 
            Using ҽ��id_In, ���ͺ�_In, Ҫ��ʱ��_In; 
        Exception 
          When Others Then 
            Null; 
        End; 
        If v_����id = 0 Then 
          v_Error := '��ǰҽ����δ������Һ�����������ִ�еǼǣ�'; 
          Raise Err_Custom; 
        End If; 
      End If; 
    End If; 
    --��鵱ǰҽ���Ƿ�����Һ 
  End If; 
 
  --����ҽ��ִ�� 
  Select Count(1) 
  Into v_Count 
  From ����ҽ��ִ�� 
  Where ҽ��id = ҽ��id_In And ���ͺ� = ���ͺ�_In And ִ��ʱ�� = ִ��ʱ��_In; 
  If v_Count > 0 Then 
    v_Error := '��ָ����ִ��ʱ�䣬�Ѿ�ִ�й�����ҽ���������һ��ִ��ʱ�䡣'; 
    Raise Err_Custom; 
  End If; 
  Insert Into ����ҽ��ִ�� 
    (ҽ��id, ���ͺ�, Ҫ��ʱ��, ��������, ִ��ժҪ, ִ����, ִ��ʱ��, �Ǽ�ʱ��, �Ǽ���, ִ�н��, ˵��, ��Һͨ��) 
  Values 
    (ҽ��id_In, ���ͺ�_In, Ҫ��ʱ��_In, ��������_In, ִ��ժҪ_In, ִ����_In, ִ��ʱ��_In, v_Date, v_��Ա����, ִ�н��_In, δִ��ԭ��_In, ��Һͨ��_In); 
 
  --���ü�¼��ִ��״̬���и��� 
  Select Decode(a.ִ��״̬, 1, a.��������, c.�ǼǴ���), Decode(a.ִ��״̬, 1, 0, a.�������� - c.�ǼǴ���) ,c.�ǼǴ���
  Into n_ִ�д���, n_ʣ����� ,n_�Ǽ�����
  From ����ҽ������ A, 
       (Select ҽ��id_In ҽ��id, ���ͺ�_In ���ͺ�, Nvl(Sum(b.��������), 0) As �ǼǴ��� 
         From ����ҽ��ִ�� B 
         Where b.ҽ��id = ҽ��id_In And b.���ͺ� = ���ͺ�_In And Nvl(b.ִ�н��, 1) <> 0) C 
  Where a.ҽ��id = c.ҽ��id And a.���ͺ� = c.���ͺ� And a.ҽ��id = ҽ��id_In And a.���ͺ� = ���ͺ�_In; 
  --���ȫ��ִ����״̬Ϊ1��δִ��״̬Ϊ0������ִ��״̬Ϊ2 
  Select Decode(n_ʣ�����, 0, 1, Decode(n_ִ�д���, 0, 0, 2)) Into n_ִ��״̬ From Dual; 
 
  --��д��ִ��״̬��ͱ��Ϊ����ִ�� 
  If Nvl(����ִ��_In, 0) = 1 Then 
    Update ����ҽ������ 
    Set ִ��״̬ = Decode(n_ִ�д���, 0, 0, 3) 
    Where ִ��״̬ In (0, 3) And ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In; 
  Else 
    Update ����ҽ������ 
    Set ִ��״̬ = Decode(n_ִ�д���, 0, 0, 3) 
    Where ִ��״̬ In (0, 3) And ���ͺ� + 0 = ���ͺ�_In And 
          ҽ��id In (Select ID 
                   From ����ҽ����¼ 
                   Where ID = v_��id And Nvl(�������, '*') = v_�������
                   Union All
                   Select ID
                   From ����ҽ����¼
                   Where ���id = v_��id And Nvl(�������, '*') = v_�������);
  End If; 
 
  --���¶�Ӧ�ķ���ִ��״̬Ϊ��ִ��(������ִ��) 
  --��Ӧ�ô���ҩƷ�͸������õ����� 
  If ִ�н��_In = 1 Then 
    If v_�������� = 2 Then 
      If Nvl(����ִ��_In, 0) = 1 Then 
        Update סԺ���ü�¼ A 
        Set ִ��״̬ = n_ִ��״̬, ִ���� = Decode(n_ִ��״̬, 0, Null, ִ����_In), ִ��ʱ�� = Decode(n_ִ��״̬, 0, Null, ִ��ʱ��_In) 
        Where �շ���� Not In ('5', '6', '7') And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In) And Not Exists 
         (Select 1 From �������� Where ����id = a.�շ�ϸĿid And �������� = 1) And a.��¼״̬ In (0, 1, 3) And 
              (ҽ�����, NO, ��¼����) In 
              (Select ҽ��id, NO, ��¼���� 
               From ����ҽ������ 
               Where ִ��״̬ = Decode(n_ִ�д���, 0, 0, 3) And ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In); 
      Else 
        Update סԺ���ü�¼ A 
        Set ִ��״̬ = n_ִ��״̬, ִ���� = Decode(n_ִ��״̬, 0, Null, ִ����_In), ִ��ʱ�� = Decode(n_ִ��״̬, 0, Null, ִ��ʱ��_In) 
        Where �շ���� Not In ('5', '6', '7') And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In) And Not Exists 
         (Select 1 From �������� Where ����id = a.�շ�ϸĿid And �������� = 1) And a.��¼״̬ In (0, 1, 3) And 
              (ҽ�����, NO, ��¼����) In 
              (Select ҽ��id, NO, ��¼���� 
               From ����ҽ������ 
               Where ִ��״̬ = Decode(n_ִ�д���, 0, 0, 3) And ���ͺ� + 0 = ���ͺ�_In And 
                     ҽ��id In (Select ID 
                              From ����ҽ����¼ 
                              Where ID = v_��id And ������� = v_������� 
                              Union All 
                              Select ID From ����ҽ����¼ Where ���id = v_��id And ������� = v_�������)); 
      End If; 
    Else 
      If Nvl(����ִ��_In, 0) = 1 Then 
        --�������ﵥ��n_ִ��״̬����Ϊ0���Ǽ�ִ�������ѡ��ִ�н��Ϊδִ�У���������ж� 
        Update ������ü�¼ A 
        Set ִ��״̬ = n_ִ��״̬, ִ���� = Decode(n_ִ��״̬, 0, Null, ִ����_In), ִ��ʱ�� = Decode(n_ִ��״̬, 0, Null, ִ��ʱ��_In) 
        Where �շ���� Not In ('5', '6', '7') And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In) And Not Exists 
         (Select 1 From �������� Where ����id = a.�շ�ϸĿid And �������� = 1) And a.��¼״̬ In (0, 1, 3) And 
              (ҽ�����, NO, ��¼����) In 
              (Select ҽ��id, NO, ��¼���� 
               From ����ҽ������ 
               Where ִ��״̬ = Decode(n_ִ�д���, 0, 0, 3) And ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In); 
      Else 
        Update ������ü�¼ A 
        Set ִ��״̬ = n_ִ��״̬, ִ���� = Decode(n_ִ��״̬, 0, Null, ִ����_In), ִ��ʱ�� = Decode(n_ִ��״̬, 0, Null, ִ��ʱ��_In) 
        Where �շ���� Not In ('5', '6', '7') And (ִ�в���id_In = 0 Or a.ִ�в���id = ִ�в���id_In) And Not Exists 
         (Select 1 From �������� Where ����id = a.�շ�ϸĿid And �������� = 1) And a.��¼״̬ In (0, 1, 3) And 
              (ҽ�����, NO, ��¼����) In 
              (Select ҽ��id, NO, ��¼���� 
               From ����ҽ������ 
               Where ִ��״̬ = Decode(n_ִ�д���, 0, 0, 3) And ���ͺ� + 0 = ���ͺ�_In And 
                     ҽ��id In (Select ID 
                              From ����ҽ����¼ 
                              Where ID = v_��id And ������� = v_������� 
                              Union All 
                              Select ID From ����ҽ����¼ Where ���id = v_��id And ������� = v_�������)); 
      End If; 
    End If; 
    --�����Զ���ɲɼ� 
    If v_������� = 'E' And v_�������� = '6' Then 
      Update ����ҽ������ A 
      Set a.������ = ִ����_In, a.����ʱ�� = ִ��ʱ��_In 
      Where ҽ��id In (Select ID 
                     From ����ҽ����¼ 
                     Where ID = v_��id 
                     Union All 
                     Select ID From ����ҽ����¼ Where ���id = v_��id) And ���ͺ� = ���ͺ�_In; 
    End If; 
 
    --ִ�����δﵽ֮���Զ����ִ��(��Ҫ����PDA�Զ�ִ��)������������ƶ��ٴ�����ʿվ��PDAһ�¡� 
    v_�Զ���� := �Զ����_In; 
    If Nvl(v_�Զ����, 0) = 0 And v_������Դ = 2 And Instr('C,D', v_�������) = 0 Then 
      Begin 
        Execute Immediate 'Select Count(1) From ZLMBSYSTEMS' 
          Into v_Count; 
      Exception 
        When Others Then 
          Null; 
      End; 
      If v_Count > 0 Then 
        v_�Զ���� := 1; 
      End If; 
    End If; 
 
    If Nvl(v_�Զ����, 0) = 1 Or ������Ŀ����_In = 1 Then 
      Begin 
        Select Decode(Sign(Nvl(Sum(b.��������), 0) - a.��������), 1, 1, 0, 1, 0) 
        Into v_�Զ���� 
        From ����ҽ������ A, ����ҽ��ִ�� B 
        Where a.ҽ��id = b.ҽ��id And a.���ͺ� = b.���ͺ� And a.ִ��״̬ In (0, 3) And a.ҽ��id = ҽ��id_In And a.���ͺ� = ���ͺ�_In 
        Group By a.��������; 
      Exception 
        When Others Then 
          Null; 
      End; 
 
      If Nvl(v_�Զ����, 0) = 1 Or ������Ŀ����_In = 1 Then 
        Zl_����ҽ��ִ��_Finish(ҽ��id_In, ���ͺ�_In, Null, ����ִ��_In, v_��Ա���, v_��Ա����, ִ�в���id_In, ������Ŀ����_In); 
      End If; 
    End If; 
    --����ҽ��ִ�мƼ�.ִ��״̬
    If n_�������� > 0 Then
      Select Count(distinct Ҫ��ʱ��) Into v_Count From ҽ��ִ�мƼ� Where ҽ��ID = ҽ��ID_IN And ���ͺ� = ���ͺ�_IN;
      If v_Count > 0 Then
        n_�������� := n_�������� / v_Count;
        --��ִ������+�������� �ܹ��ܹ�ִ�ж��ٸ�ʱ���,ȡ�������
        v_Count := ceil((n_�Ǽ�����) / n_��������);
        --��ȡִ�н���Ҫ��ʱ�� 
        Select Ҫ��ʱ�� Into d_Ҫ��ʱ��
        From (Select Ҫ��ʱ��, Rownum As ����
               From (Select Distinct Ҫ��ʱ�� From ҽ��ִ�мƼ� Where ҽ��id = ҽ��ID_IN And ���ͺ� = ���ͺ�_IN Order By Ҫ��ʱ��))
        Where ���� = v_Count;
        
        If Not d_Ҫ��ʱ�� Is Null Then
          --�ȼ���Ƿ��Ѿ��˷�
          Select Max(NVL(ִ��״̬,0)) Into v_Count From ҽ��ִ�мƼ� Where ҽ��id = ҽ��ID_IN And ���ͺ� = ���ͺ�_IN And Ҫ��ʱ�� <= d_Ҫ��ʱ��;
          If v_Count = 2 Then
            v_Error := '��ָ����ִ��ʱ��ε�ҽ�������Ѿ����˷ѣ���������ִ�С�'; 
            Raise Err_Custom; 
          End If;
          --���½���Ҫ��ʱ��֮ǰ(��)�ļ�¼ִ��״̬��
          Update ҽ��ִ�мƼ� Set ִ��״̬ = 1 Where ҽ��id = ҽ��ID_IN And ���ͺ� = ���ͺ�_IN And Ҫ��ʱ�� <= d_Ҫ��ʱ�� And NVL(ִ��״̬,0) <> 2;
        End If;
      End If;
    End If;
  End If; 
Exception 
  When Err_Custom Then 
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]'); 
  When Others Then 
    Zl_Errorcenter(Sqlcode, Sqlerrm); 
End Zl_����ҽ��ִ��_Insert;
/

--114434:��ΰ��,2017-11-17,�������׺�����ҩ�����������
CREATE OR REPLACE Procedure Zl_����ҽ����¼_Insert
(
  Id_In           ����ҽ����¼.Id%Type,
  ���id_In       ����ҽ����¼.���id%Type,
  ���_In         ����ҽ����¼.���%Type,
  ������Դ_In     ����ҽ����¼.������Դ%Type,
  ����id_In       ����ҽ����¼.����id%Type,
  ��ҳid_In       ����ҽ����¼.��ҳid%Type,
  Ӥ��_In         ����ҽ����¼.Ӥ��%Type,
  ҽ��״̬_In     ����ҽ����¼.ҽ��״̬%Type,
  ҽ����Ч_In     ����ҽ����¼.ҽ����Ч%Type,
  �������_In     ����ҽ����¼.�������%Type,
  ������Ŀid_In   ����ҽ����¼.������Ŀid%Type,
  �շ�ϸĿid_In   ����ҽ����¼.�շ�ϸĿid%Type,
  ����_In         ����ҽ����¼.����%Type,
  ��������_In     ����ҽ����¼.��������%Type,
  �ܸ�����_In     ����ҽ����¼.�ܸ�����%Type,
  ҽ������_In     ����ҽ����¼.ҽ������%Type,
  ҽ������_In     ����ҽ����¼.ҽ������%Type,
  �걾��λ_In     ����ҽ����¼.�걾��λ%Type,
  ִ��Ƶ��_In     ����ҽ����¼.ִ��Ƶ��%Type,
  Ƶ�ʴ���_In     ����ҽ����¼.Ƶ�ʴ���%Type,
  Ƶ�ʼ��_In     ����ҽ����¼.Ƶ�ʼ��%Type,
  �����λ_In     ����ҽ����¼.�����λ%Type,
  ִ��ʱ�䷽��_In ����ҽ����¼.ִ��ʱ�䷽��%Type,
  �Ƽ�����_In     ����ҽ����¼.�Ƽ�����%Type,
  ִ�п���id_In   ����ҽ����¼.ִ�п���id%Type,
  ִ������_In     ����ҽ����¼.ִ������%Type,
  ������־_In     ����ҽ����¼.������־%Type,
  ��ʼִ��ʱ��_In ����ҽ����¼.��ʼִ��ʱ��%Type,
  ִ����ֹʱ��_In ����ҽ����¼.ִ����ֹʱ��%Type,
  ���˿���id_In   ����ҽ����¼.���˿���id%Type,
  ��������id_In   ����ҽ����¼.��������id%Type,
  ����ҽ��_In     ����ҽ����¼.����ҽ��%Type,
  ����ʱ��_In     ����ҽ����¼.����ʱ��%Type,
  �Һŵ�_In       ����ҽ����¼.�Һŵ�%Type := Null,
  ǰ��id_In       ����ҽ����¼.ǰ��id%Type := Null,
  ��鷽��_In     ����ҽ����¼.��鷽��%Type := Null,
  ִ�б��_In     ����ҽ����¼.ִ�б��%Type := Null,
  �ɷ����_In     ����ҽ����¼.�ɷ����%Type := Null,
  ժҪ_In         ����ҽ����¼.ժҪ%Type := Null,
  ����Ա����_In   ����ҽ��״̬.������Ա%Type := Null,
  ��Ѽ���_In     ����ҽ����¼.��Ѽ���%Type := Null,
  ��ҩĿ��_In     ����ҽ����¼.��ҩĿ��%Type := Null,
  ��ҩ����_In     ����ҽ����¼.��ҩ����%Type := Null,
  ���״̬_In     ����ҽ����¼.���״̬%Type := Null,
  �������_In     ����ҽ����¼.�������%Type := Null,
  ����˵��_In     ����ҽ����¼.����˵��%Type := Null,
  �״�����_In     ����ҽ����¼.�״�����%Type := Null,
  �䷽id_In       ����ҽ����¼.�䷽id%Type := Null,
  �������_In     ����ҽ����¼.�������%Type := Null,
  �����Ŀid_In   ����ҽ����¼.�����Ŀid%Type := Null,
  Ƥ�Խ��_In     ����ҽ����¼.Ƥ�Խ��%Type := Null,
  �������_In       ����ҽ����¼.�������%Type := Null
  --���ܣ�ҽ����ʿ�¿�,��¼ҽ��ʱ�²�����ҽ����¼�������������סԺ��
) Is
  v_Temp     Varchar2(255);
  v_��Ա���� ����ҽ��״̬.������Ա%Type;

  v_���� ������Ϣ.����%Type;
  v_�Ա� ������Ϣ.�Ա�%Type;
  v_���� ������Ϣ.����%Type;

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

  --����ҽ����¼
  Insert Into ����ҽ����¼
    (ID, ���id, ���, ������Դ, ����id, ��ҳid, ����, �Ա�, ����, Ӥ��, ҽ��״̬, ҽ����Ч, �������, ������Ŀid, �շ�ϸĿid, ����, ��������, �ܸ�����, ҽ������, ҽ������, �걾��λ,
     ��鷽��, ִ�б��, ִ��Ƶ��, Ƶ�ʴ���, Ƶ�ʼ��, �����λ, ִ��ʱ�䷽��, �Ƽ�����, ִ�п���id, ִ������, ������־, �ɷ����, ��ʼִ��ʱ��, ִ����ֹʱ��, ���˿���id, ��������id, ����ҽ��,
     ����ʱ��, �Һŵ�, ǰ��id, ժҪ, ��Ѽ���, ����ʱ��, ��ҩĿ��, ��ҩ����, ���״̬, �������, ����˵��, �״�����, �䷽id, �������, �����Ŀid, Ƥ�Խ��,�������)
  Values
    (Id_In, ���id_In, ���_In, ������Դ_In, ����id_In, ��ҳid_In, v_����, v_�Ա�, v_����, Ӥ��_In, ҽ��״̬_In, ҽ����Ч_In, �������_In, ������Ŀid_In,
     �շ�ϸĿid_In, ����_In, ��������_In, �ܸ�����_In, ҽ������_In, ҽ������_In, �걾��λ_In, ��鷽��_In, ִ�б��_In, ִ��Ƶ��_In, Ƶ�ʴ���_In, Ƶ�ʼ��_In, �����λ_In,
     ִ��ʱ�䷽��_In, �Ƽ�����_In, ִ�п���id_In, ִ������_In, ������־_In, �ɷ����_In, ��ʼִ��ʱ��_In, ִ����ֹʱ��_In, ���˿���id_In, ��������id_In, ����ҽ��_In,
     ����ʱ��_In, �Һŵ�_In, ǰ��id_In, ժҪ_In, ��Ѽ���_In,
     Decode(�������_In, 'F', To_Date(�걾��λ_In, 'yyyy-mm-dd hh24:mi:ss'), 'K', To_Date(�걾��λ_In, 'yyyy-mm-dd hh24:mi:ss'),
             Null), ��ҩĿ��_In, ��ҩ����_In, ���״̬_In, �������_In, ����˵��_In, �״�����_In, �䷽id_In, �������_In, �����Ŀid_In, Ƥ�Խ��_In,�������_In);

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

--111635:���Ʊ�,2017-11-16,�Զ������뵥
Create Or Replace Procedure Zl_���Ƶ���Ŀ¼_Edit
(
  ����_In  In Number, --1:����;2-�޸�;3-ɾ��
  Id_In    In �����ļ��б�.ID%Type,
  ���_In  In �����ļ��б�.���%Type := Null,
  ����_In  In �����ļ��б�.����%Type := Null,
  ˵��_In  In �����ļ��б�.˵��%Type := Null,
  ����_In  In �����ļ��б�.����%Type := Null,
  ͨ��_In  In �����ļ��б�.ͨ��%Type := Null,
  ����_In  In Varchar2 := Null, --��chr(10)�зָ���chr(9)�ֶηָ�
  ����1_In In Varchar2 := Null,
  ����2_In In Varchar2 := Null,
  ����_In  In �����ļ��б�.����%Type := Null,
  ��ʽ_In  In �����ļ��б�.��ʽ%Type := Null
) Is
  v_��� �����ļ��б�.���%Type; --ԭ���
  n_ϵͳ zlTools.zlReports.ϵͳ%Type;
  v_���� zlTools.zlReports.����%Type := 'Wait';

  --��õ�ǰϵͳ��
  Function f_Cur_Sys Return Number Is
    n_Sys_No zlTools.zlSystems.���%Type;
  Begin
    Select Min(���)
    Into n_Sys_No
    From zlTools.zlSystems
    Where ������ In
          (Select Owner From All_Objects Where Object_Name = Upper('Zl_���Ƶ���Ŀ¼_Edit') And Object_Type = 'PROCEDURE');
    Return n_Sys_No;
  End f_Cur_Sys;

  --��д���븽��
  Procedure p_Append_Items Is
    v_All    Varchar2(2000);
    v_Row    Varchar2(1000);
    v_Val    Varchar2(1000);
    v_��Ŀ   �������ݸ���.��Ŀ%Type;
    n_����   �������ݸ���.����%Type := 0;
    n_����   �������ݸ���.����%Type := 0;
    n_Ҫ��id �������ݸ���.Ҫ��id%Type := Null;
    v_����   �������ݸ���.����%Type;
    n_ֻ��   �������ݸ���.ֻ��%Type := 0;
  Begin
    Delete �������ݸ��� Where �ļ�id = Id_In;
    
    If ����_In Is Null Then
      Return;
    End If;
    
    v_All := ����_In || Chr(10);
    Loop
      v_Row  := Substr(v_All, 1, Instr(v_All, Chr(10)) - 1);
      v_��Ŀ := Substr(v_Row, 1, Instr(v_Row, Chr(9)) - 1);

      v_Val := Substr(v_Row, Instr(v_Row, Chr(9), 1, 1) + 1, Instr(v_Row, Chr(9), 1, 2) - Instr(v_Row, Chr(9), 1, 1) - 1);
      If v_Val Is Null Then
        n_���� := 0;
      Else
        n_���� := To_Number(v_Val);
      End If;
      
      v_Val := Substr(v_Row, Instr(v_Row, Chr(9), 1, 2) + 1, Instr(v_Row, Chr(9), 1, 3) - Instr(v_Row, Chr(9), 1, 2) - 1);
      If v_Val Is Null Then
        n_ֻ�� := 0;
      Else
        n_ֻ�� := To_Number(v_Val);
      End If;
            
      v_Val := Substr(v_Row, Instr(v_Row, Chr(9), 1, 3) + 1, Instr(v_Row, Chr(9), 1, 4) - Instr(v_Row, Chr(9), 1, 3) - 1);
      If v_Val Is Null Then
        n_Ҫ��id := Null;
      Else
        n_Ҫ��id := To_Number(v_Val);
      End If;
      
      v_Val  := Substr(v_Row, Instr(v_Row, Chr(9), 1, 4) + 1);
      v_���� := v_Val;

      n_���� := n_���� + 1;
      Insert Into �������ݸ���
        (�ļ�id, ��Ŀ, ����, ����, Ҫ��id, ����,ֻ��)
      Values
        (Id_In, v_��Ŀ, n_����, n_����, n_Ҫ��id, v_����,n_ֻ��);
      v_All := Substr(v_All, Instr(v_All, Chr(10)) + 1);
      Exit When v_All Is Null;
    End Loop;
    
    delete ��������ģ�� a where a.�����ļ�Id=Id_In and not exists(select 1 from �������ݸ��� where ��Ŀ=a.���ݸ���);
  End p_Append_Items;

  --�����Ƶ��ݱ���ģ����ӱ����ݶ�Ӧ����
  Procedure p_Add_Report(Form_In Number) Is
    --������form_In=1,����; form_In=2,����
    n_Mdl_Id zlTools.zlReports.ID%Type;
    n_Rpt_Id zlTools.zlReports.ID%Type;
    n_Dat_Id zlTools.zlRPTDatas.ID%Type;
    e_Mod_Lost Exception;
  Begin
    Begin
      Select ID Into n_Mdl_Id From zlReports Where ϵͳ = n_ϵͳ And Upper(���) = 'ZLEMRBILLMOLD1-' || Form_In;
    Exception
      When Others Then
        n_Mdl_Id := 0;
    End;
    If n_Mdl_Id = 0 Then
      Raise e_Mod_Lost;
    End If;
    -- 11698 �����ı���ȱ���룬����ֱ�����
    If Form_In = 1 Then
      v_���� := ����1_In;
    Elsif Form_In = 2 Then
      v_���� := ����2_In;
    End If;
    If v_���� Is Null Then
      v_���� := 'Wait...';
    End If;
    Select zlTools.Zlreports_Id.Nextval Into n_Rpt_Id From Dual;
    Insert Into zlTools.zlReports
      (ID, ���, ����, ˵��, ����, ��ֽ, ��ӡ��, Ʊ��, ϵͳ, ����id, ����, �޸�ʱ��, ����ʱ��)
      Select n_Rpt_Id, 'ZLCISBILL00' || ���_In || '-' || Form_In, ����_In, ˵��_In, v_����, ��ֽ, ��ӡ��, Ʊ��, ϵͳ, Null, Null,
             Sysdate, Null
      From zlTools.zlReports
      Where ID = n_Mdl_Id;
    For r_Rptdatas In (Select ID From zlTools.zlRPTDatas Where ����id = n_Mdl_Id) Loop
      Select zlTools.Zlrptdatas_Id.Nextval Into n_Dat_Id From Dual;
      Insert Into zlTools.zlRPTDatas
        (ID, ����id, ����, �ֶ�, ����, ����)
        Select n_Dat_Id, n_Rpt_Id, ����, �ֶ�, ����, ���� From zlTools.zlRPTDatas Where ID = r_Rptdatas.ID;
      Insert Into zlTools.zlRPTPars
        (Դid, ����, ���, ����, ����, ȱʡֵ, ��ʽ, ֵ�б�, ����sql, ��ϸsql, �����ֶ�, ��ϸ�ֶ�, ����)
        Select n_Dat_Id, ����, ���, ����, ����, ȱʡֵ, ��ʽ, ֵ�б�, ����sql, ��ϸsql, �����ֶ�, ��ϸ�ֶ�, ����
        From zlTools.zlRPTPars
        Where Դid = r_Rptdatas.ID;
      Insert Into zlTools.zlRPTSQLs
        (Դid, �к�, ����)
        Select n_Dat_Id, �к�, ���� From zlTools.zlRPTSQLs Where Դid = r_Rptdatas.ID;
    End Loop;
    Insert Into zlTools.zlRPTFMTs
      (����id, ���, ˵��, W, H, ֽ��, ֽ��, ��ֽ̬��, ͼ��)
      Select n_Rpt_Id, ���, ˵��, W, H, ֽ��, ֽ��, ��ֽ̬��, ͼ�� From zlTools.zlRPTFMTs Where ����id = n_Mdl_Id;
    For r_Rptitems In (Select ID From zlTools.zlRPTItems Where ����id = n_Mdl_Id Order By ID) Loop
      Insert Into zlTools.zlRPTItems
        (ID, ����id, ��ʽ��, ����, ����, �ϼ�id, ���, ����, ����, ����, ��ͷ, X, Y, W, H, �и�, ����, �Ե�, ����, �ֺ�, ����, б��, ����, ǰ��, ����, �߿�, ����,
         ��ʽ, ����, ����, ����, ϵͳ)
        Select zlTools.Zlrptitems_Id.Nextval, n_Rpt_Id, ��ʽ��, ����, ����, zlTools.Zlrptitems_Id.Nextval - (ID - �ϼ�id), ���, ����,
               ����, ����, ��ͷ, X, Y, W, H, �и�, ����, �Ե�, ����, �ֺ�, ����, б��, ����, ǰ��, ����, �߿�, ����, ��ʽ, ����, ����, ����, ϵͳ
        From zlTools.zlRPTItems
        Where ID = r_Rptitems.ID;
    End Loop;
    Update zlTools.zlRPTItems Set ���� = ����_In Where ����id = n_Rpt_Id And ���� = 'ZLBILLCAPTION';
  Exception
    When e_Mod_Lost Then
      Raise_Application_Error(-20101, '[ZLSOFT]���Ƶ���ģ�ⶪʧ������ϵϵͳ����Ա��[ZLSOFT]');
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End p_Add_Report;

  --������
Begin
  n_ϵͳ := f_Cur_Sys;
  If ����_In = 1 Then
    Insert Into ����ҳ���ʽ (����, ���, ����) Values (7, ���_In, ����_In);
    Insert Into �����ļ��б�
      (ID, ����, ���, ����, ˵��, ����, ͨ��, ҳ��,����, ��ʽ)
    Values
      (Id_In, 7, ���_In, ����_In, ˵��_In, ����_In, ͨ��_In, ���_In,����_In, ��ʽ_In);
    p_Append_Items;
    p_Add_Report(1);
    --If ͨ��_In = 2 Then
    -- 11547 �������Ƶ���ʱ��ֻҪ��ִ�к��б��桱��Ч�����������Ƶ�����Ҫ�ж�Ӧ���Զ��屨��
    p_Add_Report(2);
    --End If;

  Elsif ����_In = 2 Then
    Select ��� Into v_��� From �����ļ��б� Where ID = Id_In;
    Update ����ҳ���ʽ Set ��� = ���_In, ���� = ����_In Where ���� = 7 And ��� = v_���;
    Update �����ļ��б�
    Set ��� = ���_In, ���� = ����_In, ˵�� = ˵��_In, ͨ�� = ͨ��_In, ����=����_In, ��ʽ = ��ʽ_In
    Where ���� = 7 And ID = Id_In;
    p_Append_Items;

    Update zlTools.zlReports
    Set ��� = 'ZLCISBILL00' || ���_In || '-1', ���� = ����_In, ˵�� = ˵��_In, ���� = ����1_In
    Where ϵͳ = n_ϵͳ And ��� = 'ZLCISBILL00' || v_��� || '-1';
    If Sql%RowCount = 0 Then
      p_Add_Report(1);
    End If;
    --If ͨ��_In <> 2 Then
    -- 11323 �ı����Ƶ��ݵĸ�ʽ,��ɾ���Զ��屨��(2007-08-15 �¶�)
    -- Delete zlTools.zlReports Where ϵͳ = n_ϵͳ And ��� = 'ZLCISBILL00' || v_��� || '-2';
    --  Null;
    --Else
    Update zlTools.zlReports
    Set ��� = 'ZLCISBILL00' || ���_In || '-2', ���� = ����_In, ˵�� = ˵��_In, ���� = ����2_In
    Where ϵͳ = n_ϵͳ And ��� = 'ZLCISBILL00' || v_��� || '-2';
    If Sql%RowCount = 0 Then
      p_Add_Report(2);
    End If;
    --End If;

  Elsif ����_In = 3 Then
    Select ��� Into v_��� From �����ļ��б� Where ID = Id_In;
    Delete �����ļ��б� Where ID = Id_In;
    Delete ����ҳ���ʽ Where ���� = 7 And ��� = v_���;

    Delete zlTools.zlReports Where ϵͳ = n_ϵͳ And ��� Like 'ZLCISBILL00' || v_��� || '-_';
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���Ƶ���Ŀ¼_Edit;
/

--111635:���Ʊ�,2017-11-16,�Զ������뵥
Create Or Replace Function Zl_Lob_Read
(
  Tab_In     In Number,
  Key_In     In Varchar2,
  Pos_In     In Number,
  Moved_In   In Number := 0,
  Lobtype_In In Number := 0
  --����˵����
  --Tab_In������LOB�����ݱ�
  --        0-�������ͼ��;1-�����ļ���ʽ;2-�����ļ�ͼ��;3-�������ĸ�ʽ;4-��������ͼ��;
  --        5-���Ӳ�����ʽ;6-���Ӳ���ͼ��;7-����ҳ���ʽ(ͼ��);8-���Ӳ�������;9-�����ص����;
  --        10-�ٴ�·���ļ�;11-�ٴ�·��ͼ��;12-����ҳ���ʽ(ҳü�ļ�);13-����ҳ���ʽ(ҳ���ļ�);
  --        14-��Ա֤���¼;19-������չ��Ϣ;20-��Ա��չ��Ϣ;22-ҽ����������;23-��Ӧ����Ƭ;24-�Զ������뵥�ļ�;25-ҽ�����뵥�ļ�
  --Key_In�����ݼ�¼�Ĺؼ���
  --Pos_In����0��ʼ���϶�ȡ��ֱ������Ϊ��
  --Moved_In: 0������¼,1��ȡת���󱸱��¼
  --LobType_IN:0-BLOb,1-CLOB
) Return Varchar2 Is
  l_Blob   Blob;
  l_Clob   Clob;
  v_Buffer Varchar2(32767);
  n_Amount Number := 2000;
  n_Offset Number := 1;
  t_Key    t_Strlist;
Begin
  If Tab_In = 0 Then
    Select ͼ�� Into l_Blob From �������ͼ�� Where ���� = Key_In;
  Elsif Tab_In = 1 Then
    Select ���� Into l_Blob From �����ļ���ʽ Where �ļ�id = To_Number(Key_In);
  Elsif Tab_In = 2 Then
    Select ͼ�� Into l_Blob From �����ļ�ͼ�� Where ����id = To_Number(Key_In);
  Elsif Tab_In = 3 Then
    Select ���� Into l_Blob From �������ĸ�ʽ Where �ļ�id = To_Number(Key_In);
  Elsif Tab_In = 4 Then
    Select ͼ�� Into l_Blob From ��������ͼ�� Where ����id = To_Number(Key_In);
  Elsif Tab_In = 5 Then
    If Moved_In = 0 Then
      Select ���� Into l_Blob From ���Ӳ�����ʽ Where �ļ�id = To_Number(Key_In);
    Else
      Select ���� Into l_Blob From H���Ӳ�����ʽ Where �ļ�id = To_Number(Key_In);
    End If;
  Elsif Tab_In = 6 Then
    If Moved_In = 0 Then
      Select ͼ�� Into l_Blob From ���Ӳ���ͼ�� Where ����id = To_Number(Key_In);
    Else
      Select ͼ�� Into l_Blob From H���Ӳ���ͼ�� Where ����id = To_Number(Key_In);
    End If;
  Elsif Tab_In = 7 Then
    Select ͼ��
    Into l_Blob
    From ����ҳ���ʽ
    Where ���� = To_Number(Substr(Key_In, 1, 1)) And ��� = Substr(Key_In, 3);
  Elsif Tab_In = 8 Then
    If Moved_In = 0 Then
      Select ����
      Into l_Blob
      From ���Ӳ�������
      Where ����id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And ��� = Substr(Key_In, Instr(Key_In, ',') + 1);
    Else
      Select ����
      Into l_Blob
      From H���Ӳ�������
      Where ����id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And ��� = Substr(Key_In, Instr(Key_In, ',') + 1);
    End If;
  Elsif Tab_In = 9 Then
    Select ���ͼ�� Into l_Blob From �����ص���� Where ��� = To_Number(Key_In);
  Elsif Tab_In = 10 Then
    Select ����
    Into l_Blob
    From �ٴ�·���ļ�
    Where ·��id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And �ļ��� = Substr(Key_In, Instr(Key_In, ',') + 1);
  Elsif Tab_In = 11 Then
    Select ͼ�� Into l_Blob From �ٴ�·��ͼ�� Where ID = To_Number(Key_In);
  Elsif Tab_In = 12 Then
    Select ҳü�ļ�
    Into l_Blob
    From ����ҳ���ʽ
    Where ���� = To_Number(Substr(Key_In, 1, 1)) And ��� = Substr(Key_In, 3);
  Elsif Tab_In = 13 Then
    Select ҳ���ļ�
    Into l_Blob
    From ����ҳ���ʽ
    Where ���� = To_Number(Substr(Key_In, 1, 1)) And ��� = Substr(Key_In, 3);
  Elsif Tab_In = 14 Then
    Select Column_Value Bulk Collect Into t_Key From Table(f_Str2list(Key_In));
    Select ǩ����Ϣ Into l_Clob From ��Ա֤���¼ Where ��Աid = To_Number(t_Key(1)) And Certsn = t_Key(2);
  Elsif Tab_In = 19 Then
    Select Column_Value Bulk Collect Into t_Key From Table(f_Str2list(Key_In));
    Select ͼƬ Into l_Blob From ������չ��Ϣ Where ����id = To_Number(t_Key(1)) And ��Ŀ = t_Key(2);
  Elsif Tab_In = 20 Then
    Select Column_Value Bulk Collect Into t_Key From Table(f_Str2list(Key_In));
    Select ͼƬ Into l_Blob From ��Ա��չ��Ϣ Where ��Աid = To_Number(t_Key(1)) And ��Ŀ = t_Key(2);
  Elsif Tab_In = 22 Then
    Select Column_Value Bulk Collect Into t_Key From Table(f_Str2list(Key_In));
    Select ���� Into l_Blob From ҽ���������� Where ID = To_Number(Key_In);
  Elsif Tab_In = 23 Then
    If To_Number(Substr(Key_In, Instr(Key_In, ',') + 1))=0 Then
       Select ���֤����Ƭ Into l_Blob From ��Ӧ����Ƭ Where ��Ӧ��ID =To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1));
    Elsif  To_Number(Substr(Key_In, Instr(Key_In, ',') + 1))=1 Then
       Select ִ�պ���Ƭ Into l_Blob From ��Ӧ����Ƭ Where ��Ӧ��ID =To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1));
    Elsif To_Number(Substr(Key_In, Instr(Key_In, ',') + 1))=2 Then
       Select ��Ȩ����Ƭ Into l_Blob From ��Ӧ����Ƭ Where ��Ӧ��ID =To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1));
    End If;
  Elsif Tab_In = 24 Then
    Select ����
    Into l_Clob
    From �Զ������뵥�ļ�
    Where �ļ�id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And
          ��� = To_Number(Substr(Key_In, Instr(Key_In, ',') + 1)) ;
  Elsif Tab_In = 25 Then
    Select ����
    Into l_Clob
    From ҽ�����뵥�ļ�
    Where ҽ��id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And
          ��� = To_Number(Substr(Key_In, Instr(Key_In, ',') + 1)) ;
  End If;

  n_Offset := n_Offset + Pos_In * n_Amount;
  If Lobtype_In = 1 Then
    If l_Clob Is Null Then
      v_Buffer := Null;
    Else
      Dbms_Lob.Read(l_Clob, n_Amount, n_Offset, v_Buffer);
    End If;
  Else
    If l_Blob Is Null Then
      v_Buffer := Null;
    Else
      Dbms_Lob.Read(l_Blob, n_Amount, n_Offset, v_Buffer);
    End If;
  End If;
  Return v_Buffer;
Exception
  When No_Data_Found Then
    Return Null;
End Zl_Lob_Read;
/

--111635:���Ʊ�,2017-11-16,�Զ������뵥
Create Or Replace Procedure Zl_Lob_Append
(
  Tab_In     In Number,
  Key_In     In Varchar2,
  Txt_In     In Varchar2, --16���Ƶ��ļ�Ƭ�λ�����Ƭ��
  Cls_In     In Number := 0, --�Ƿ����ԭ�������ݣ���һƬ�δ���ʱΪ1���Ժ�Ϊ0
  Lobtype_In In Number := 0 --0-BLOB;1-CLOB
  --����˵����
  --Tab_In������LOB�����ݱ�
  --        0-�������ͼ��;1-�����ļ���ʽ;2-�����ļ�ͼ��;3-�������ĸ�ʽ;4-��������ͼ��;
  --        5-���Ӳ�����ʽ;6-���Ӳ���ͼ��;7-����ҳ���ʽ��8-���Ӳ�������;9-�����ص����
  --        10-�ٴ�·���ļ�,11-�ٴ�·��ͼ��;14-��Ա֤���¼;
  --        19-������չ��Ϣ;20-��Ա��չ��Ϣ;22-ҽ����������;23-��Ӧ����Ƭ;24-�Զ������뵥�ļ�;25-ҽ�����뵥�ļ�
  --Key_In�����ݼ�¼�Ĺؼ���
  --Txt_In��16���Ƶ��ļ�Ƭ�λ�����Ƭ��
  --Cls_In���Ƿ����ԭ�������ݣ���һƬ�δ���ʱΪ1���Ժ�Ϊ0
  --Lobtype_In:--0-BLOB;1-CLOB
) Is
  l_Blob Blob;
  l_Clob Clob;
  t_Key  t_Strlist;

Begin
  If Tab_In = 0 Then
    If Cls_In = 1 Then
      Update �������ͼ�� Set ͼ�� = Empty_Blob() Where ���� = Key_In;
    End If;
    Select ͼ�� Into l_Blob From �������ͼ�� Where ���� = Key_In For Update;
  Elsif Tab_In = 1 Then
    If Cls_In = 1 Then
      Update �����ļ���ʽ Set ���� = Empty_Blob() Where �ļ�id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into �����ļ���ʽ (�ļ�id, ����) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select ���� Into l_Blob From �����ļ���ʽ Where �ļ�id = To_Number(Key_In) For Update;
  Elsif Tab_In = 2 Then
    If Cls_In = 1 Then
      Update �����ļ�ͼ�� Set ͼ�� = Empty_Blob() Where ����id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into �����ļ�ͼ�� (����id, ͼ��) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select ͼ�� Into l_Blob From �����ļ�ͼ�� Where ����id = To_Number(Key_In) For Update;
  Elsif Tab_In = 3 Then
    If Cls_In = 1 Then
      Update �������ĸ�ʽ Set ���� = Empty_Blob() Where �ļ�id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into �������ĸ�ʽ (�ļ�id, ����) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select ���� Into l_Blob From �������ĸ�ʽ Where �ļ�id = To_Number(Key_In) For Update;
  Elsif Tab_In = 4 Then
    If Cls_In = 1 Then
      Update ��������ͼ�� Set ͼ�� = Empty_Blob() Where ����id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into ��������ͼ�� (����id, ͼ��) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select ͼ�� Into l_Blob From ��������ͼ�� Where ����id = To_Number(Key_In) For Update;
  Elsif Tab_In = 5 Then
    If Cls_In = 1 Then
      Update ���Ӳ�����ʽ Set ���� = Empty_Blob() Where �ļ�id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into ���Ӳ�����ʽ (�ļ�id, ����) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select ���� Into l_Blob From ���Ӳ�����ʽ Where �ļ�id = To_Number(Key_In) For Update;
  Elsif Tab_In = 6 Then
    If Cls_In = 1 Then
      Update ���Ӳ���ͼ�� Set ͼ�� = Empty_Blob() Where ����id = To_Number(Key_In);
      If Sql%RowCount = 0 Then
        Insert Into ���Ӳ���ͼ�� (����id, ͼ��) Values (To_Number(Key_In), Empty_Blob());
      End If;
    End If;
    Select ͼ�� Into l_Blob From ���Ӳ���ͼ�� Where ����id = To_Number(Key_In) For Update;
  Elsif Tab_In = 7 Then
    If Cls_In = 1 Then
      Update ����ҳ���ʽ
      Set ͼ�� = Empty_Blob()
      Where ���� = To_Number(Substr(Key_In, 1, 1)) And ��� = Substr(Key_In, 3);
    End If;
    Select ͼ��
    Into l_Blob
    From ����ҳ���ʽ
    Where ���� = To_Number(Substr(Key_In, 1, 1)) And ��� = Substr(Key_In, 3)
    For Update;
  Elsif Tab_In = 8 Then
    If Cls_In = 1 Then
      Update ���Ӳ�������
      Set ���� = Empty_Blob()
      Where ����id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And ��� = Substr(Key_In, Instr(Key_In, ',') + 1);
    End If;
    Select ����
    Into l_Blob
    From ���Ӳ�������
    Where ����id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And ��� = Substr(Key_In, Instr(Key_In, ',') + 1)
    For Update;
  Elsif Tab_In = 9 Then
    If Cls_In = 1 Then
      Update �����ص���� Set ���ͼ�� = Empty_Blob() Where ��� = To_Number(Key_In);
    End If;
    Select ���ͼ�� Into l_Blob From �����ص���� Where ��� = To_Number(Key_In) For Update;
  Elsif Tab_In = 10 Then
    If Cls_In = 1 Then
      Update �ٴ�·���ļ�
      Set ���� = Empty_Blob()
      Where ·��id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And
            �ļ��� = Substr(Key_In, Instr(Key_In, ',') + 1);
    End If;
    Select ����
    Into l_Blob
    From �ٴ�·���ļ�
    Where ·��id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And �ļ��� = Substr(Key_In, Instr(Key_In, ',') + 1)
    For Update;
  Elsif Tab_In = 11 Then
    If Cls_In = 1 Then
      Update �ٴ�·��ͼ�� Set ͼ�� = Empty_Blob() Where ID = To_Number(Key_In);
    End If;
    Select ͼ�� Into l_Blob From �ٴ�·��ͼ�� Where ID = To_Number(Key_In) For Update;
  Elsif Tab_In = 12 Then
    If Cls_In = 1 Then
      Update ����ҳ���ʽ
      Set ҳü�ļ� = Empty_Blob()
      Where ���� = To_Number(Substr(Key_In, 1, 1)) And ��� = Substr(Key_In, 3);
    End If;
    Select ҳü�ļ�
    Into l_Blob
    From ����ҳ���ʽ
    Where ���� = To_Number(Substr(Key_In, 1, 1)) And ��� = Substr(Key_In, 3)
    For Update;
  Elsif Tab_In = 13 Then
    If Cls_In = 1 Then
      Update ����ҳ���ʽ
      Set ҳ���ļ� = Empty_Blob()
      Where ���� = To_Number(Substr(Key_In, 1, 1)) And ��� = Substr(Key_In, 3);
    End If;
    Select ҳ���ļ�
    Into l_Blob
    From ����ҳ���ʽ
    Where ���� = To_Number(Substr(Key_In, 1, 1)) And ��� = Substr(Key_In, 3)
    For Update;
  Elsif Tab_In = 14 Then
    Select Column_Value Bulk Collect Into t_Key From Table(f_Str2list(Key_In));
    If Cls_In = 1 Then
      Update ��Ա֤���¼ Set ǩ����Ϣ = Empty_Clob() Where ��Աid = To_Number(t_Key(1)) And Certsn = t_Key(2);
    End If;
    Select ǩ����Ϣ Into l_Clob From ��Ա֤���¼ Where ��Աid = To_Number(t_Key(1)) And Certsn = t_Key(2) For Update;
  Elsif Tab_In = 19 Then
    Select Column_Value Bulk Collect Into t_Key From Table(f_Str2list(Key_In));
    If Cls_In = 1 Then
      Update ������չ��Ϣ Set ͼƬ = Empty_Blob() Where ����id = To_Number(t_Key(1)) And ��Ŀ = t_Key(2);
    End If;
    Select ͼƬ Into l_Blob From ������չ��Ϣ Where ����id = To_Number(t_Key(1)) And ��Ŀ = t_Key(2) For Update;
    Update ���ű� Set ����޸�ʱ�� = Sysdate Where ID = To_Number(t_Key(1));
  Elsif Tab_In = 20 Then
    Select Column_Value Bulk Collect Into t_Key From Table(f_Str2list(Key_In));
    If Cls_In = 1 Then
      Update ��Ա��չ��Ϣ Set ͼƬ = Empty_Blob() Where ��Աid = To_Number(t_Key(1)) And ��Ŀ = t_Key(2);
    End If;
    Select ͼƬ Into l_Blob From ��Ա��չ��Ϣ Where ��Աid = To_Number(t_Key(1)) And ��Ŀ = t_Key(2) For Update;
    Update ��Ա�� Set ����޸�ʱ�� = Sysdate Where ID = To_Number(t_Key(1));
  Elsif Tab_In = 22 Then
    Select Column_Value Bulk Collect Into t_Key From Table(f_Str2list(Key_In));
    If Cls_In = 1 Then
      Update ҽ���������� Set ���� = Empty_Blob() Where ID = To_Number(t_Key(1));
    End If;
    Select ���� Into l_Blob From ҽ���������� Where ID = To_Number(t_Key(1)) For Update;
  Elsif Tab_In = 23 Then
    If To_Number(Substr(Key_In, Instr(Key_In, ',') + 1))=0 Then
      If Cls_In = 1 Then
        Update ��Ӧ����Ƭ Set ���֤����Ƭ = Empty_Blob() Where ��Ӧ��ID = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1));
        If Sql%RowCount = 0 Then
          Insert Into ��Ӧ����Ƭ (��Ӧ��ID, ���֤����Ƭ,ִ�պ���Ƭ,��Ȩ����Ƭ) Values (To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)), Empty_Blob(), Empty_Blob(), Empty_Blob());
        End If;
      End If;
      Select ���֤����Ƭ Into l_Blob From ��Ӧ����Ƭ Where ��Ӧ��ID =To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) For Update;
    Elsif  To_Number(Substr(Key_In, Instr(Key_In, ',') + 1))=1 Then
      If Cls_In = 1 Then
        Update ��Ӧ����Ƭ Set ִ�պ���Ƭ = Empty_Blob() Where ��Ӧ��ID = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1));
        If Sql%RowCount = 0 Then
          Insert Into ��Ӧ����Ƭ (��Ӧ��ID, ���֤����Ƭ,ִ�պ���Ƭ,��Ȩ����Ƭ) Values (To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)), Empty_Blob(), Empty_Blob(), Empty_Blob());
        End If;
      End If;
      Select ִ�պ���Ƭ Into l_Blob From ��Ӧ����Ƭ Where ��Ӧ��ID =To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) For Update;
    Elsif To_Number(Substr(Key_In, Instr(Key_In, ',') + 1))=2 Then
     If Cls_In = 1 Then
        Update ��Ӧ����Ƭ Set ��Ȩ����Ƭ = Empty_Blob() Where ��Ӧ��ID = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1));
        If Sql%RowCount = 0 Then
          Insert Into ��Ӧ����Ƭ (��Ӧ��ID, ���֤����Ƭ,ִ�պ���Ƭ,��Ȩ����Ƭ) Values (To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)), Empty_Blob(), Empty_Blob(), Empty_Blob());
        End If;
      End If;
      Select ��Ȩ����Ƭ Into l_Blob From ��Ӧ����Ƭ Where ��Ӧ��ID =To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) For Update;
    End If;
  Elsif Tab_In = 24 Then
    If Cls_In = 1 Then
      Update �Զ������뵥�ļ�
      Set ���� = Empty_Clob()
      Where �ļ�id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And
            ��� = To_Number(Substr(Key_In, Instr(Key_In, ',') + 1)) ;
    End If;
    Select ����
    Into l_Clob
    From �Զ������뵥�ļ�
    Where �ļ�id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And ��� = To_Number(Substr(Key_In, Instr(Key_In, ',') + 1)) 
    For Update;
  ElsIf Tab_In = 25 Then
    If Cls_In = 1 Then
      Update ҽ�����뵥�ļ�
      Set ���� = Empty_Clob()
      Where ҽ��id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And 
            ��� = To_Number(Substr(Key_In, Instr(Key_In, ',') + 1))  ;
    End If;
    Select ����
    Into l_Clob
    From ҽ�����뵥�ļ�
    Where ҽ��id = To_Number(Substr(Key_In, 1, Instr(Key_In, ',') - 1)) And ��� = To_Number(Substr(Key_In, Instr(Key_In, ',') + 1)) 
    For Update;
  End If;

  If Lobtype_In = 1 Then
    Dbms_Lob.Writeappend(l_Clob, Length(Txt_In), Txt_In);
  Else
    Dbms_Lob.Writeappend(l_Blob, Length(Hextoraw(Txt_In)) / 2, Hextoraw(Txt_In));
  End If;

Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Lob_Append;
/

--115026:������,2017-12-04,����Σ��ֵ
--111635:���Ʊ�,2017-11-16,�Զ������뵥
CREATE OR REPLACE Procedure Zl_Retu_Clinic
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

  v_Dblink Varchar2(30);

  Type t_Tab_Col Is Table Of Varchar2(4000) Index By Varchar2(32);
  Arr_Tab_Col t_Tab_Col;

  ---------------------------------------------
  --���ܣ���ȡ����ֶ��ַ���
  Function Getfields(v_Table In Varchar2) Return Varchar2 As
    v_Colstr Varchar2(4000);
  Begin
    If Arr_Tab_Col.Exists(v_Table) Then
      v_Colstr := Arr_Tab_Col(v_Table);
    Else
      Select f_List2str(Cast(Collect(Column_Name) As t_Strlist)) As Colsstr
      Into v_Colstr
      From (Select Column_Name From User_Tab_Columns Where Table_Name = v_Table Order By Column_Id);

      Arr_Tab_Col(v_Table) := v_Colstr;
    End If;

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
              From Table(f_Str2list('���Ӳ�������,���Ӳ�����ʽ,���Ӳ�������,�����걨��¼,�������淴��,�����걨����'))) Loop
      v_Table := r.Column_Value;
      If v_Table = '���Ӳ�������' Then
        v_Field := '����id';
      Else
        v_Field := '�ļ�id';
      End If;
      v_Fields := Getfields(v_Table);

    --��LOB�ֶεı�(���Ӳ���ͼ��,���Ӳ�����ʽ,���Ӳ�������)����H������ʱ��������ֱ��ָ��dblink
      If v_Dblink Is Not Null And (v_Table = '���Ӳ�������' Or v_Table = '���Ӳ�����ʽ') Then
        v_Sql := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '��ת��', 'Null as ��ת��') ||
                 ' From ' || v_Table || '@' || v_Dblink || ' Where ' || v_Field || ' = :1';
      Else
        v_Sql := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' || Replace(v_Fields, '��ת��', 'Null as ��ת��') ||
                 ' From H' || v_Table || ' Where ' || v_Field || ' = :1';
      End If;
      Execute Immediate v_Sql
        Using n_Rec_Id;

      If v_Table = '���Ӳ�������' Then
        v_Fields := Getfields('���Ӳ���ͼ��');

        If v_Dblink Is Not Null Then
          v_Sql := 'Insert Into ���Ӳ���ͼ��(' || v_Fields || ') Select ' || Replace(v_Fields, '��ת��', 'Null as ��ת��') ||
                   ' From ���Ӳ���ͼ��@' || v_Dblink ||
                   ' a Where ����id In (Select ID From H���Ӳ������� Where �ļ�id = :1 And �������� = 5)';
        Else
          v_Sql := 'Insert Into ���Ӳ���ͼ��(' || v_Fields || ') Select ' || Replace(v_Fields, '��ת��', 'Null as ��ת��') ||
                   ' From H���Ӳ���ͼ�� Where ����id In (Select ID From H���Ӳ������� Where �ļ�id = :1 And �������� = 5)';
        End If;
        Execute Immediate v_Sql
          Using n_Rec_Id;

        If v_Dblink Is Not Null Then
          v_Sql := 'Delete ���Ӳ���ͼ��@' || v_Dblink ||
                   ' Where ����id In (Select ID From H���Ӳ������� Where �ļ�id = :1 And �������� = 5)';
          Execute Immediate v_Sql
            Using n_Rec_Id;
        Else
          Delete H���Ӳ���ͼ�� Where ����id In (Select ID From H���Ӳ������� Where �ļ�id = n_Rec_Id And �������� = 5);
        End If;
      End If;

      If v_Dblink Is Not Null And (v_Table = '���Ӳ�������' Or v_Table = '���Ӳ�����ʽ') Then
        v_Sql := 'Delete ' || v_Table || '@' || v_Dblink || ' Where ' || v_Field || ' = :1';
      Else
        v_Sql := 'Delete H' || v_Table || ' Where ' || v_Field || ' = :1';
      End If;
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
              From Table(f_Str2list('����ҽ���Ƽ�,����ҽ��״̬,����ҽ������,����ҽ������,����ҽ������,����ҽ��ִ��,����ҽ����ӡ,��Ѫ�����¼,��Ѫ������,��Ѫ������Ŀ,' ||
                                     'ҽ��ִ�д�ӡ,ҽ��ִ��ʱ��,ҽ��ִ�мƼ�,ִ�д�ӡ��¼,�������ҽ��,����·��ҽ��,����ҽ������,������ļ�¼,' ||
                                     'Ӱ�񱨸沵��,Ӱ�񱨸��¼,Ӱ�񱨸������¼,Ӱ�����¼,Ӱ�����뵥ͼ��,Ӱ���ղ�����,Ӱ��Σ��ֵ��¼,����걾��¼,�����Լ���¼,������ռ�¼,�������Լ�¼,ҽ�����뵥�ļ�,����Σ��ֵ��¼'))) Loop
      v_Table := p.Column_Value;
      If Instr('����·��ҽ��', v_Table) > 0 Then
        v_Field := '����ҽ��ID';
      Else
        v_Field := 'ҽ��ID';
      End If;

      v_Fields := Getfields(v_Table);
      v_Sql    := 'Insert Into ' || v_Table || '(' || v_Fields || ') Select ' ||
                  Replace(v_Fields, '��ת��', 'Null as ��ת��') || ' From H' || v_Table || ' Where ' || v_Field || ' = :1';
      If v_Table = '����ҽ��״̬' Or v_Table = '����ҽ������' Then
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
      Elsif v_Table = '����ҽ������' Then
        v_Fields := Getfields('ҽ����������');
        v_Sql    := 'Insert Into ҽ����������(' || v_Fields || ') Select ' || Replace(v_Fields, '��ת��', 'Null as ��ת��') ||
                    ' From Hҽ���������� Where ID In (Select ����id From H����ҽ������ Where ҽ��id = :1 And ����id Is Not Null)';
        Execute Immediate v_Sql
          Using n_Rec_Id;

        Delete Hҽ����������
        Where ID In (Select ����id From H����ҽ������ Where ҽ��id = n_Rec_Id And ����id Is Not Null);

        Execute Immediate v_Sqlchild
          Using n_Rec_Id;
      Elsif v_Table = '����Σ��ֵ��¼' Then
      
        v_Fields := Getfields('����Σ��ֵ����');
        v_Sql    := 'Insert Into ����Σ��ֵ����(' || v_Fields || ') Select ' || Replace(v_Fields, '��ת��', 'Null as ��ת��') ||
                    ' From H����Σ��ֵ���� Where Σ��ֵid In (Select ID From H����Σ��ֵ��¼ Where ҽ��id = :1)';
        Execute Immediate v_Sql
          Using n_Rec_Id;      
        Delete H����Σ��ֵ���� Where Σ��ֵid In (Select ID From H����Σ��ֵ��¼ Where ҽ��id = n_Rec_Id);

        v_Fields := Getfields('����Σ��ֵҽ��');
        v_Sql    := 'Insert Into ����Σ��ֵҽ��(' || v_Fields || ') Select ' || Replace(v_Fields, '��ת��', 'Null as ��ת��') ||
                    ' From H����Σ��ֵҽ�� Where Σ��ֵid In (Select ID From H����Σ��ֵ��¼ Where ҽ��id = :1)';
        Execute Immediate v_Sql
          Using n_Rec_Id;
      
        Delete H����Σ��ֵҽ�� Where Σ��ֵid In (Select ID From H����Σ��ֵ��¼ Where ҽ��id = n_Rec_Id);

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

  Select Max(Db����) Into v_Dblink From zlBakSpaces Where ϵͳ = 100 And ��ǰ = 1;

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

    For r_Epr In (Select b.Id
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

--111635:���Ʊ�,2017-07-14,XML�Զ������뵥
Create Or Replace Procedure Zl_�Զ������뵥�ļ�_Edit
(
  ģʽ_In   Number, --1-����/�޸�;2-ɾ��
  �ļ�id_In �Զ������뵥�ļ�.�ļ�id%Type,
  ���_In   �Զ������뵥�ļ�.���%Type,
  �ļ���_In �Զ������뵥�ļ�.�ļ���%Type := Null
) As
  v_Temp     Varchar(500);
  v_��Ա���� Varchar(100);
Begin
  v_Temp     := Zl_Identity;
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  If ģʽ_In = 1 Then
    Update �Զ������뵥�ļ�
    Set �ļ��� = �ļ���_In, ������ = v_��Ա����, ����ʱ�� = Sysdate
    Where �ļ�id = �ļ�id_In And ��� = ���_In;
    If Sql%RowCount = 0 Then
      Insert Into �Զ������뵥�ļ�
        (�ļ�id, �ļ���, ���, ������, ����ʱ��)
      Values
        (�ļ�id_In, �ļ���_In, ���_In, v_��Ա����, Sysdate);
    End If;
  Elsif ģʽ_In = 2 Then
    Delete From �Զ������뵥�ļ� Where �ļ�id = �ļ�id_In And ��� = ���_In;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_�Զ������뵥�ļ�_Edit;
/

--111635:���Ʊ�,2017-07-14,XML�Զ������뵥
CREATE OR REPLACE Procedure Zl_ҽ�����뵥�ļ�_Edit
(
  �ļ�id_In ҽ�����뵥�ļ�.�ļ�id%Type,
  �ļ���_IN ҽ�����뵥�ļ�.�ļ���%Type,
  ���_In   ҽ�����뵥�ļ�.���%Type,
  ҽ��ID_In   ҽ�����뵥�ļ�.ҽ��ID%Type 
) As

Begin
  Insert Into ҽ�����뵥�ļ�
      (�ļ�id,�ļ���, ҽ��ID, ���)
  Values
      (�ļ�id_In,�ļ���_IN, ҽ��ID_In, ���_In);
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_ҽ�����뵥�ļ�_Edit;
/

--115519:��С��,2017-11-16,���պ��زɱ걾����ҽ�����˺�����
Create Or Replace Procedure Zl_����Ԥ������_�ɼ����
(
  ҽ������_In Varchar2, --���ݰ������ҽ��IDʹ��","�ָ� 
  ��Ա���_In ��Ա��.���%Type := Null,
  ��Ա����_In ��Ա��.����%Type := Null, --Null=ȡ������Ϊ��ʱ��ɲɼ�
  ����_In     Number := 0, --0=��ɲɼ���1=ȡ���ɼ�
  ҽ�����_In Number := 0 --0=����ҽ��,1=��Ѫҽ�� 
) Is
  n_�Զ����� Number;
  --���ҵ�ǰ�걾��������� 
  Cursor c_Samplequest(v_ҽ��id In Varchar2) Is
    Select /*+ rule */
    Distinct ID As ҽ��id, ������Դ
    From ����ҽ����¼ A, ����ҽ������ B
    Where a.Id = b.ҽ��id And b.������ Is Null And Sign(Nvl(a.���id, 0)) = ҽ�����_In And
          a.Id In (Select * From Table(Cast(f_Num2list(v_ҽ��id) As Zltools.t_Numlist)));

  --δ��˵ķ�����(������ҩƷ) 
  Cursor c_Verify(v_ҽ��id In Varchar2) Is
    Select /*+ rule */
    Distinct ��¼����, NO, ���, ��¼״̬, �����־
    From סԺ���ü�¼
    Where �շ���� Not In ('5', '6', '7') And
          ҽ����� + 0 In (Select ID
                       From ����ҽ����¼
                       Where ID In (Select * From Table(Cast(f_Num2list(v_ҽ��id) As Zltools.t_Numlist))) And
                             Sign(Nvl(���id, 0)) = ҽ�����_In) And ���ʷ��� = 1 And ��¼״̬ = 0 And �۸񸸺� Is Null And
          (��¼����, NO) In (Select ��¼����, NO
                         From ����ҽ������
                         Where ҽ��id In (Select * From Table(Cast(f_Num2list(v_ҽ��id) As Zltools.t_Numlist)))
                         Union All
                         Select ��¼����, NO
                         From ����ҽ������
                         Where ҽ��id In (Select ID
                                        From ����ҽ����¼
                                        Where ID In (Select * From Table(Cast(f_Num2list(v_ҽ��id) As Zltools.t_Numlist))) And
                                              Sign(Nvl(���id, 0)) = ҽ�����_In) And ������ Is Null)
    Union All
    Select /*+ rule */
    Distinct ��¼����, NO, ���, ��¼״̬, �����־
    From ������ü�¼
    Where �շ���� Not In ('5', '6', '7') And
          ҽ����� + 0 In (Select ID
                       From ����ҽ����¼
                       Where ID In (Select * From Table(Cast(f_Num2list(v_ҽ��id) As Zltools.t_Numlist))) And
                             Sign(Nvl(���id, 0)) = ҽ�����_In) And ���ʷ��� = 1 And ��¼״̬ = 0 And �۸񸸺� Is Null And
          (��¼����, NO) In (Select ��¼����, NO
                         From ����ҽ������
                         Where ҽ��id In (Select * From Table(Cast(f_Num2list(v_ҽ��id) As Zltools.t_Numlist)))
                         Union All
                         Select ��¼����, NO
                         From ����ҽ������
                         Where ҽ��id In (Select ID
                                        From ����ҽ����¼
                                        Where ID In (Select * From Table(Cast(f_Num2list(v_ҽ��id) As Zltools.t_Numlist))) And
                                              Sign(Nvl(���id, 0)) = ҽ�����_In) And ������ Is Null)
    Order By ��¼����, NO, ���;

  v_����걾��¼ Number(18);
  v_ִ��״̬     Number(1);
  v_������       Varchar2(50);
  v_Error        Varchar2(100);
  v_No           ����ҽ������.No%Type;
  v_����         ����ҽ������.��¼����%Type;
  v_���         Varchar2(1000);

  v_�շ�ids Varchar2(4000);
  n_�ⷿid  Number;
  n_���Ϻ�  Number;

  Err_Custom Exception;
  n_Par Number;
Begin
  Select zl_GetSysParameter('�Զ���������', 1211) Into n_�Զ����� From Dual;
  If ��Ա����_In Is Not Null And ����_In = 0 Then
    --���걾�Ƿ񱻺��ջ���� 
    Begin
      Select /*+ rule */
       Nvl(c.Id, 0), b.ִ��״̬, b.������
      Into v_����걾��¼, v_ִ��״̬, v_������
      From ����ҽ����¼ A, ����ҽ������ B, ����걾��¼ C
      Where a.Id = b.ҽ��id And a.���id = c.ҽ��id(+) And
            a.Id In (Select * From Table(Cast(f_Num2list(ҽ������_In) As Zltools.t_Numlist)));
    Exception
      When Others Then
        v_����걾��¼ := 0;
    End;
  
    If v_����걾��¼ <> 0 Then
      v_Error := '�걾�ѱ�����ƺ��ղ�����ɲɼ�!';
      Raise Err_Custom;
    End If;
  
    If v_ִ��״̬ <> 2 And v_������ Is Not Null Then
      v_Error := '�걾�ѱ������ǩ�ղ�����ɲɼ�!';
      Raise Err_Custom;
    End If;
  
    --���ҽ���Ƿ��շ�
    n_Par := Zl_To_Number(Nvl(zl_GetSysParameter(163), '0'));
    If n_Par = 1 Then
      For r_Verify In c_Verify(ҽ������_In) Loop
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
    End If;
  
    Update /*+ rule */ ������ռ�¼
    Set �ز��� = ��Ա����_In, �ز�ʱ�� = Sysdate
    Where ҽ��id In (Select * From Table(Cast(f_Num2list(ҽ������_In) As Zltools.t_Numlist)));
  
    --���²ɼ���Ϣ(����Ͳɼ��� 
    Update /*+ rule */ ����ҽ������
    Set ������ = ��Ա����_In, ����ʱ�� = Sysdate, ִ��״̬ = Decode(ִ��״̬, 2, 0, ִ��״̬),
        �زɱ걾 = Decode(Nvl(�زɱ걾, 0), 0, Decode(ִ��״̬, 2, 1, 0), �زɱ걾), ִ��˵�� = Null
    Where ҽ��id In (Select * From Table(Cast(f_Num2list(ҽ������_In) As Zltools.t_Numlist)));
  
    --����ҽ���ͷ��ü�¼ 
    For r_Samplequest In c_Samplequest(ҽ������_In) Loop
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
               Where ҽ��id In (Select ID
                              From ����ҽ����¼ A, ����ҽ������ B
                              Where a.Id = b.ҽ��id And r_Samplequest.ҽ��id In (a.Id) And Sign(Nvl(a.���id, 0)) = ҽ�����_In And
                                    b.ִ��״̬ In (0, 2) And b.������ Is Null));
      Else
        --����
        If n_�Զ����� = 1 Then
          For c_Stuff In (Select a.��¼����, a.��¼״̬, b.Id, b.�ⷿid
                          From ������ü�¼ A, ҩƷ�շ���¼ B
                          Where a.Id = b.����id And a.�շ���� = '4' And b.����� Is Null And
                                (a.ҽ�����, a.��¼����, a.No) In
                                (Select ҽ��id, ��¼����, NO
                                 From ����ҽ������
                                 Where ҽ��id = r_Samplequest.ҽ��id
                                 Union All
                                 Select ҽ��id, ��¼����, NO
                                 From ����ҽ������
                                 Where ҽ��id In
                                       (Select ID
                                        From ����ҽ����¼ A, ����ҽ������ B
                                        Where a.Id = b.ҽ��id And r_Samplequest.ҽ��id In (a.Id) And
                                              Sign(Nvl(a.���id, 0)) = ҽ�����_In And b.ִ��״̬ In (0, 2) And b.������ Is Null))) Loop
            If Mod(Nvl(c_Stuff.��¼����, 0), 10) = 1 And Nvl(c_Stuff.��¼״̬, 0) = 1 Then
              If n_���Ϻ� Is Null Then
                n_���Ϻ� := Nextno(20);
              End If;
            
              If c_Stuff.�ⷿid <> Nvl(n_�ⷿid, 0) Then
                If Nvl(n_�ⷿid, 0) <> 0 And v_�շ�ids Is Not Null Then
                  v_�շ�ids := Substr(v_�շ�ids, 2);
                  Zl_ҩƷ�շ���¼_��������(v_�շ�ids, n_�ⷿid, ��Ա����_In, Sysdate, 1, ��Ա����_In, n_���Ϻ�, ��Ա����_In);
                End If;
              
                n_�ⷿid  := c_Stuff.�ⷿid;
                v_�շ�ids := Null;
              End If;
            
              v_�շ�ids := v_�շ�ids || '|' || c_Stuff.Id || ',0';
            End If;
          End Loop;
          If Nvl(n_�ⷿid, 0) <> 0 And v_�շ�ids Is Not Null Then
            v_�շ�ids := Substr(v_�շ�ids, 2);
            Zl_ҩƷ�շ���¼_��������(v_�շ�ids, n_�ⷿid, ��Ա����_In, Sysdate, 1, ��Ա����_In, n_���Ϻ�, ��Ա����_In);
          End If;
        End If;
      
        --2.����ִ�д��� 
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
               Where ҽ��id In (Select ID
                              From ����ҽ����¼ A, ����ҽ������ B
                              Where a.Id = b.ҽ��id And r_Samplequest.ҽ��id In (a.Id) And Sign(Nvl(a.���id, 0)) = ҽ�����_In And
                                    b.ִ��״̬ In (0, 2) And b.������ Is Null));
      End If;
    End Loop;
  
    --����ִ��״̬(ֻ���²ɼ��� 
    Update /*+ rule */ ����ҽ������
    Set ִ��״̬ = 1, ����� = ��Ա����_In, ���ʱ�� = Sysdate
    Where ҽ��id In (Select ID
                   From ����ҽ����¼
                   Where ID In (Select * From Table(Cast(f_Num2list(ҽ������_In) As Zltools.t_Numlist))) And
                         Sign(Nvl(���id, 0)) = ҽ�����_In);
    --3.�Զ���˼��� 
    For r_Verify In c_Verify(ҽ������_In) Loop
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
    End If;
  
  Else
    --���걾�Ƿ񱻺��ջ���� 
    Begin
      Select /*+ rule */
       Nvl(c.Id, 0), b.ִ��״̬, b.������
      Into v_����걾��¼, v_ִ��״̬, v_������
      From ����ҽ����¼ A, ����ҽ������ B, ����걾��¼ C
      Where a.Id = b.ҽ��id And a.���id = c.ҽ��id(+) And
            a.Id In (Select * From Table(Cast(f_Num2list(ҽ������_In) As Zltools.t_Numlist)));
    Exception
      When Others Then
        v_����걾��¼ := 0;
    End;
  
    If v_����걾��¼ <> 0 Then
      v_Error := '�걾�ѱ�����ƺ��ղ���ȡ����ɲɼ�!';
      Raise Err_Custom;
    End If;
  
    If v_ִ��״̬ <> 2 And v_������ Is Not Null Then
      v_Error := '�걾�ѱ������ǩ�ղ���ȡ����ɲɼ�!';
      Raise Err_Custom;
    End If;
  
    Update /*+ rule */ ����ҽ������
    Set ������ = Null, ����ʱ�� = Null, ִ��״̬ = 0, ִ��˵�� = Null, ����� = Null, ���ʱ�� = Null
    Where ҽ��id In (Select ID
                   From ����ҽ����¼
                   Where ID In (Select * From Table(Cast(f_Num2list(ҽ������_In) As Zltools.t_Numlist))));
  
    For r_Samplequest In c_Samplequest(ҽ������_In) Loop
    
      If r_Samplequest.������Դ = 2 Then
        --2.����ִ�д��� 
        Update סԺ���ü�¼
        Set ִ��״̬ = 0, ִ��ʱ�� = Null, ִ���� = ��Ա����_In
        Where �շ���� Not In ('5', '6', '7') And
              (ҽ�����, ��¼����, NO) In
              (Select ҽ��id, ��¼����, NO
               From ����ҽ������
               Where ҽ��id = r_Samplequest.ҽ��id
               Union All
               Select ҽ��id, ��¼����, NO
               From ����ҽ������
               Where ҽ��id In (Select ID
                              From ����ҽ����¼
                              Where r_Samplequest.ҽ��id In (ID) And Sign(Nvl(���id, 0)) = ҽ�����_In) And ִ��״̬ In (0, 2) And
                     ������ Is Null);
      Else
        --����
        If n_�Զ����� = 1 Then
          For c_Stuff In (Select b.Id, b.ʵ������
                          From ������ü�¼ A, ҩƷ�շ���¼ B
                          Where a.Id = b.����id And a.�շ���� = '4' And (b.��¼״̬ = 1 Or Mod(b.��¼״̬, 3) = 0) And
                                b.����� Is Not Null And (a.ҽ�����, a.��¼����, a.No) In
                                (Select ҽ��id, ��¼����, NO
                                                       From ����ҽ������
                                                       Where ҽ��id = r_Samplequest.ҽ��id
                                                       Union All
                                                       Select ҽ��id, ��¼����, NO
                                                       From ����ҽ������
                                                       Where ҽ��id In (Select ID
                                                                      From ����ҽ����¼
                                                                      Where r_Samplequest.ҽ��id In (ID) And
                                                                            Sign(Nvl(���id, 0)) = ҽ�����_In) And
                                                             ִ��״̬ In (0, 2) And ������ Is Null)) Loop
          
            Zl_�����շ���¼_��������(c_Stuff.Id, ��Ա����_In, Sysdate, Null, Null, Null, c_Stuff.ʵ������);
          End Loop;
        End If;
        --�˷�
        Update ������ü�¼
        Set ִ��״̬ = 0, ִ��ʱ�� = Null, ִ���� = ��Ա����_In
        Where �շ���� Not In ('5', '6', '7') And
              (ҽ�����, ��¼����, NO) In
              (Select ҽ��id, ��¼����, NO
               From ����ҽ������
               Where ҽ��id = r_Samplequest.ҽ��id
               Union All
               Select ҽ��id, ��¼����, NO
               From ����ҽ������
               Where ҽ��id In (Select ID
                              From ����ҽ����¼
                              Where r_Samplequest.ҽ��id In (ID) And Sign(Nvl(���id, 0)) = ҽ�����_In) And ִ��״̬ In (0, 2) And
                     ������ Is Null);
      End If;
    End Loop;
  End If;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Ԥ������_�ɼ����;
/

--116673:���˺�,2017-11-16,�������������Լ��λ����
Create Or Replace Procedure Zl_���ʷ��ü�¼_Unit
(
  Patientids_In Varchar2,
  ����id_In     ������ü�¼.����id%Type,
  ����ý���_In Number --���ν���ʱ�Ƿ��ſ����������ʵķ���,��������ʱ�����ſ�
) As
  Cursor c_Fee(v_����id ������ü�¼.����id%Type) Is
    Select A.ID, A.NO, A.���, A.��¼����, A.��¼״̬, A.ִ��״̬, Nvl(A.ʵ�ս��, 0) As δ����
    From ������ü�¼ A
    Where A.����id = v_����id And A.����id Is Null And A.��¼״̬ <> 0 And A.���ʷ��� = 1 And A.�����־ In (1, 4) And
          Not Exists
     (Select 1
           From ������ü�¼ B
           Where B.NO = A.NO And B.��¼���� = A.��¼���� And B.��� = A.���
           Group By B.NO, B.��¼����, B.���
           Having Nvl(Sum(B.ʵ�ս��), 0) = Decode(����ý���_In, 1, 1 + Nvl(Sum(B.ʵ�ս��), 0), 0))
    Union All
    Select 0 As ID, A.NO, A.���, Mod(A.��¼����, 10) As ��¼����, A.��¼״̬, A.ִ��״̬,
           Nvl(Sum(A.ʵ�ս��), 0) - Nvl(Sum(A.���ʽ��), 0) As δ����
    From (Select ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, ���ʵ�id, ����id, ҽ�����, �����־,
                  ���ʷ���, ����, �Ա�, ����, ��ʶ��, ���ʽ, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
                  ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������,
                  ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ����, ����Ա���,
                  ����Ա����, ����id, ���ʽ��, ���մ���id, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ,
                  �Ƿ���
           From ������ü�¼
           Union All
           Select ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, ���ʵ�id, ����id, ҽ�����, �����־,
                  ���ʷ���, ����, �Ա�, ����, ��ʶ��, ���ʽ, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����,
                  ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������,
                  ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ����, ����Ա���,
                  ����Ա����, ����id, ���ʽ��, ���մ���id, ������Ŀ��, ���ձ���, ��������, ͳ����, �Ƿ��ϴ�, ժҪ,
                  �Ƿ���
           From H������ü�¼) A
    Where A.����id = v_����id And A.����id Is Not Null And A.���ʷ��� = 1 And A.�����־ in( 1,4) And
          Nvl(A.ʵ�ս��, 0) <> Nvl(A.���ʽ��, 0)
    Group By A.NO, A.���, Mod(A.��¼����, 10), A.��¼״̬, A.ִ��״̬
    Having Nvl(Sum(A.ʵ�ս��), 0) - Nvl(Sum(A.���ʽ��), 0) <> 0;

  v_Patientids  Varchar2(4000);
  v_Patientid   Varchar2(4000);
  v_Banlanceids Varchar2(4000);

  Err_Custom Exception;
  v_Error Varchar2(255);
Begin

  v_Patientids := Patientids_In || ',';
  While v_Patientids Is Not Null Loop
    v_Patientid   := Substr(v_Patientids, 1, Instr(v_Patientids, ',') - 1);
    v_Patientids  := Substr(v_Patientids, Instr(v_Patientids, ',') + 1);
    v_Banlanceids := '';

    For r_Fee In c_Fee(v_Patientid) Loop
      If r_Fee.ID = 0 Then
        Zl_���ʷ��ü�¼_Insert(r_Fee.ID, r_Fee.NO, r_Fee.��¼����, r_Fee.��¼״̬, r_Fee.ִ��״̬, r_Fee.���,
                               r_Fee.δ����, ����id_In);
      Else
        v_Banlanceids := v_Banlanceids || ',' || r_Fee.ID;
        If Length(v_Banlanceids) > 3980 Then
          Zl_���ʷ��ü�¼_Batch(Substr(v_Banlanceids, 2), v_Patientid, ����id_In);
          v_Banlanceids := '';
        End If;
      End If;
    End Loop;

    If Not v_Banlanceids Is Null Then
      Zl_���ʷ��ü�¼_Batch(Substr(v_Banlanceids, 2), v_Patientid, ����id_In);
    End If;
  End Loop;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_���ʷ��ü�¼_Unit;
/

--115026:������,2017-12-04,����Σ��ֵ
--111635:���Ʊ�,2017-11-16,�Զ������뵥
--116697:������,2017-11-18,Ԥ�����˿�;��￨��������õ���ʷ����ת�����⴦��
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
  Update /*+ rule*/ סԺ���ü�¼ A
  Set ��ת�� = n_����
  Where ��ת�� Is Null And ����id Is Null And
        (����id, ��ҳid) In (Select ����id, ��ҳid
                         From ������ҳ C
                         Where ��Ժ���� < d_End And ��ת�� Is Null And ����ת�� Is Null And Not Exists
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
  Set ��ת�� = n_����
  Where ID In (Select c.����id
               From ����ҽ����¼ B, ����ҽ������ C
               Where c.ҽ��id = b.Id And Nvl(b.��ҳid, 0) = 0 And b.�Һŵ� Is Null And b.���id Is Null And b.��ת�� Is Null And
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

--117082:����,2017-11-21,����ѹ����Ĵ���
--115695:����,2017-11-09,�޸�����ʱ�ı����ַ�����
Create Or Replace Procedure Zl_���Ʒ���Ŀ¼_Update
(
  Id_In      ���Ʒ���Ŀ¼.Id%Type,
  �ϼ�id_In  ���Ʒ���Ŀ¼.�ϼ�id%Type,
  ����_In    ���Ʒ���Ŀ¼.����%Type,
  ����_In    ���Ʒ���Ŀ¼.����%Type,
  ����_In    ���Ʒ���Ŀ¼.����%Type,
  v_Brethren Number
  --�Ƿ��ͬ��������г��ȴ���,0-��,1-��
) As
  v_Oldcode Varchar2(20); --ԭ���ı���
  v_Parent  Varchar2(20); --�ϼ�����
  v_Extend  Number(20); --���䳤��(Ϊ����ʾѹ��)
  v_Kind    Number(1); --��ǰ��Ŀ������
  Err_Notfind Exception;

Begin
  Select RTrim(����), ���� Into v_Oldcode, v_Kind From ���Ʒ���Ŀ¼ Where ID = Id_In;
  If v_Oldcode Is Null Then
    Raise Err_Notfind;
  End If;
  --�޸���Ŀ����
  Update ���Ʒ���Ŀ¼
  Set �ϼ�id = Decode(�ϼ�id_In, 0, Null, �ϼ�id_In), ���� = ����_In, ���� = ����_In, ���� = ����_In
  Where ID = Id_In;
  --�޸ı�ϵ������������
  Update ���Ʒ���Ŀ¼
  Set ���� = ����_In || Substr(����, Length(v_Oldcode) + 1)
  Where ���� <> ����_In And ���� Like v_Oldcode || '_%' And ���� = v_Kind;
  --����ͬ������ĳ���
  If v_Brethren = 1 Then
    If Nvl(�ϼ�id_In, 0) <> 0 Then
      Select ���� Into v_Parent From ���Ʒ���Ŀ¼ Where ID = �ϼ�id_In;
    Else
      v_Parent := Null;
    End If;
    Begin
      Select Length(RTrim(����_In)) - Length(RTrim(����))
      Into v_Extend
      From ���Ʒ���Ŀ¼
      Where (�ϼ�id = �ϼ�id_In Or �ϼ�id Is Null And Nvl(�ϼ�id_In, 0) = 0) And ID <> Id_In And ���� = v_Kind And Rownum = 1;
    Exception
      When Others Then
        v_Extend := 0;
    End;
    If v_Extend > 0 Then
      --���䴦��
      If v_Parent Is Null Then
        Update ���Ʒ���Ŀ¼
        Set ���� = LPad('0', v_Extend, '0') || ����
        Where ���� = v_Kind And ID Not In (Select ID From ���Ʒ���Ŀ¼ Start With ID = Id_In Connect By Prior ID = �ϼ�id);
      Else
        Update ���Ʒ���Ŀ¼
        Set ���� = v_Parent || LPad('0', v_Extend, '0') || Substr(����, Length(v_Parent) + 1)
        Where ���� = v_Kind And ���� Like v_Parent || '_%' And
              ID Not In (Select ID From ���Ʒ���Ŀ¼ Start With ID = Id_In Connect By Prior ID = �ϼ�id);
      End If;
    End If;
    If v_Extend < 0 Then
      --ѹ������
      If v_Parent Is Null Then
        Update ���Ʒ���Ŀ¼
        Set ���� = Substr(����, 1 + Abs(v_Extend))
        Where ID Not In (Select ID
                         From ���Ʒ���Ŀ¼
                         Where ���� = v_Kind
                         Start With �ϼ�id = Id_In
                         Connect By Prior ID = �ϼ�id
                         Union All
                         Select ID From ���Ʒ���Ŀ¼ Where ���� = v_Kind And ID = Id_In) And ���� = v_Kind;
      Else
        Update ���Ʒ���Ŀ¼
        Set ���� = v_Parent || Substr(����, Length(v_Parent) + 1 + Abs(v_Extend))
        Where ���� Like v_Parent || '_%' And
              ID Not In (Select ID
                         From ���Ʒ���Ŀ¼
                         Where ���� = v_Kind
                         Start With �ϼ�id = Id_In
                         Connect By Prior ID = �ϼ�id
                         Union All
                         Select ID From ���Ʒ���Ŀ¼ Where ���� = v_Kind And ID = Id_In) And ���� = v_Kind;
      End If;
    End If;
  End If;

Exception
  When Err_Notfind Then
    Raise_Application_Error(-20101, '[ZLSOFT]����Ŀ�����ڣ������ѱ������û�ɾ����[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���Ʒ���Ŀ¼_Update;
/

--115695:����,2017-11-09,�޸��ϼ����롢���䳤�ȵĳ���
Create Or Replace Procedure Zl_���Ʒ���Ŀ¼_Insert
(
  Id_In      ���Ʒ���Ŀ¼.Id%Type,
  �ϼ�id_In  ���Ʒ���Ŀ¼.�ϼ�id%Type,
  ����_In    ���Ʒ���Ŀ¼.����%Type,
  ����_In    ���Ʒ���Ŀ¼.����%Type,
  ����_In    ���Ʒ���Ŀ¼.����%Type,
  ����_In    ���Ʒ���Ŀ¼.����%Type,
  v_Brethren Number
  --�Ƿ��ͬ��������г��ȴ���,0-��,1-��
) As
  v_Parent Varchar2(20); --�ϼ�����
  v_Extend Number(20); --���䳤��(Ϊ����ʾѹ��)
Begin
  --����ͬ������ĳ���
  If v_Brethren = 1 Then
    If Nvl(�ϼ�id_In, 0) <> 0 Then
      Select ���� Into v_Parent From ���Ʒ���Ŀ¼ Where ID = �ϼ�id_In;
    Else
      v_Parent := Null;
    End If;
    Begin
      Select Length(RTrim(����_In)) - Length(RTrim(����))
      Into v_Extend
      From ���Ʒ���Ŀ¼
      Where (�ϼ�id = �ϼ�id_In Or �ϼ�id Is Null And Nvl(�ϼ�id_In, 0) = 0) And ID <> Id_In And ���� = ����_In And Rownum = 1;
    Exception
      When Others Then
        v_Extend := 0;
    End;
    If v_Extend > 0 Then
      --���䴦��
      If v_Parent Is Null Then
        Update ���Ʒ���Ŀ¼ Set ���� = LPad('0', v_Extend, '0') || ���� Where ID <> Id_In And ���� = ����_In;
      Else
        Update ���Ʒ���Ŀ¼
        Set ���� = v_Parent || LPad('0', v_Extend, '0') || Substr(����, Length(v_Parent) + 1)
        Where ���� Like v_Parent || '_%' And ���� = ����_In;
      End If;
    End If;
    If v_Extend < 0 Then
      --ѹ������
      If v_Parent Is Null Then
        Update ���Ʒ���Ŀ¼ Set ���� = Substr(����, 1 + Abs(v_Extend)) Where ID <> Id_In And ���� = ����_In;
      Else
        Update ���Ʒ���Ŀ¼
        Set ���� = v_Parent || Substr(����, Length(v_Parent) + 1 + Abs(v_Extend))
        Where ���� Like v_Parent || '_%' And ���� = ����_In;
      End If;
    End If;
  End If;
  --��ӱ���¼
  Insert Into ���Ʒ���Ŀ¼
    (ID, �ϼ�id, ����, ����, ����, ����)
  Values
    (Id_In, Decode(�ϼ�id_In, 0, Null, �ϼ�id_In), ����_In, ����_In, ����_In, ����_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���Ʒ���Ŀ¼_Insert;
/

--116034:���Ʊ�,2017-11-03,·�����ɲ�������������
CREATE OR REPLACE Function Zl_Lob_ReadForPath
(
  Tab_In     In Number,
  Key_In     In Varchar2,
  Pos_In     In Number,
  Moved_In   In Number := 0,
  Lobtype_In In Number := 0
  --����˵����
  --Tab_In������LOB�����ݱ�
  --        5-���Ӳ�����ʽ
  --Key_In�����ݼ�¼�Ĺؼ���
  --Pos_In����0��ʼ���϶�ȡ��ֱ������Ϊ��
  --Moved_In: 0������¼,1��ȡת���󱸱��¼
  --LobType_IN:0-BLOb,1-CLOB
) Return Varchar2 Is
  l_Blob   Blob;
  l_Clob   Clob;
  v_Buffer Varchar2(32767);
  n_Amount Number := 2000;
  n_Offset Number := 1;
Begin
  If Tab_In = 5 Then
    If Moved_In = 0 Then
      Select ���� Into l_Blob From ���Ӳ�����ʽ Where �ļ�id = To_Number(Key_In);
    Else
      Select ���� Into l_Blob From H���Ӳ�����ʽ Where �ļ�id = To_Number(Key_In);
    End If;
  End If;

  n_Offset := n_Offset + Pos_In * n_Amount;
  If Lobtype_In = 1 Then
    If l_Clob Is Null Then
      v_Buffer := Null;
    Else
      Dbms_Lob.Read(l_Clob, n_Amount, n_Offset, v_Buffer);
    End If;
  Else
    If l_Blob Is Null Then
      v_Buffer := Null;
    Else
      Dbms_Lob.Read(l_Blob, n_Amount, n_Offset, v_Buffer);
    End If;
  End If;
  Return v_Buffer;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Lob_ReadForPath;
/

--114647:Ƚ����,2017-11-02,����ʵ������ʱ������ҩ���ڷ����仯��û�и���ҩƷ�շ���¼�ķ�ҩ����
Create Or Replace Procedure Zl_������ʼ�¼_Verify
(
  No_In         ������ü�¼.No%Type,
  ����Ա���_In ������ü�¼.����Ա���%Type,
  ����Ա����_In ������ü�¼.����Ա����%Type,
  ���_In       Varchar2 := Null,
  ���ʱ��_In   ������ü�¼.�Ǽ�ʱ��%Type := Null
) As
  --���ܣ����һ��������ʻ��۵�
  --������
  --    ���_IN����ʽ��"1,3,5,7,8",Ϊ�ձ�ʾ�������δ��˵���
  --    ���ʱ��_IN�����ڲ�����Ҫͳһ���ƻ򷵻�ʱ��ĵط�
  --ֻ��ȡָ����ŵ�,δ��˵Ĳ��ݽ��д���
  Cursor c_Bill Is
    Select a.Id, a.����id, a.ʵ�ս��, a.�����־, a.������Ŀid, a.ִ�в���id, a.��������id, a.���˿���id, a.��ҩ����, a.�շ����, Nvl(b.��������, 0) As ��������,
           a.ҽ�����
    From ������ü�¼ A, �������� B
    Where a.�շ�ϸĿid = b.����id(+) And a.��¼���� = 2 And a.��¼״̬ = 0 And a.No = No_In And
          (Instr(',' || ���_In || ',', ',' || Nvl(a.�۸񸸺�, a.���) || ',') > 0 Or ���_In Is Null)
    Order By a.���;

  --����а����������õ�δ������ʱ�����ݲ��������Ƿ��Զ�����
  Cursor c_Stuff Is
    Select ID, �ⷿid
    From ҩƷ�շ���¼ M
    Where NO = No_In And ���� = 25 And �ⷿid Is Not Null And ��¼״̬ = 1 And ����� Is Null And Exists
     (Select 1
           From ������ü�¼ A, �������� B
           Where a.Id = m.����id + 0 And a.��¼���� = 2 And a.��¼״̬ = 1 And a.No = No_In And
                 (Instr(',' || ���_In || ',', ',' || Nvl(a.�۸񸸺�, a.���) || ',') > 0 Or ���_In Is Null) And
                 a.�շ�ϸĿid = b.����id And b.�������� = 1)
    Order By �ⷿid, ҩƷid;

  --
  n_���Ϻ�   ҩƷ�շ���¼.���ܷ�ҩ��%Type;
  n_�ⷿid   ҩƷ�շ���¼.�ⷿid%Type;
  v_�շ�ids  Varchar2(4000);
  d_Date     Date;
  v_ҽ��ids  Varchar2(4000);
  v_��ҩ���� ҩƷ�շ���¼.��ҩ����%Type;

  Type t_Record Is Record(
    ҩ��id   Number(18),
    ��ҩ���� Varchar2(10));

  Type t_��ҩ���� Is Table Of t_Record;
  c_��ҩ���� t_��ҩ���� := t_��ҩ����();
  n_Step     Number(18);
  n_Havedata Number(2);
  n_Count    Number(18);

Begin
  If ���ʱ��_In Is Null Then
    Select Sysdate Into d_Date From Dual;
  Else
    d_Date := ���ʱ��_In;
  End If;

  For r_Bill In c_Bill Loop
  
    --����ҩ����
    If (r_Bill.�շ���� In ('5', '6', '7') Or r_Bill.�շ���� = '4' And r_Bill.�������� = 1) Then
      --ͬһ�ŵ���,����ͬһҩ��ͬһ����
      v_��ҩ���� := Null;
      n_Havedata := 0;
      For n_Step In 1 .. c_��ҩ����.Count Loop
        If c_��ҩ����(n_Step).ҩ��id = Nvl(r_Bill.ִ�в���id, 0) Then
          v_��ҩ���� := c_��ҩ����(n_Step).��ҩ����;
          n_Havedata := 1;
          Exit;
        End If;
      End Loop;
    
      If v_��ҩ���� Is Null Then
        --ͬһ��������ͨ�ŹҺ���Ч�Һ���������δ��ҩ�����ϰ��,�����һ�μ��˴���Ϊ׼
        n_Count := To_Number(Substr(Nvl(zl_GetSysParameter(21), '11') || '11', 1, 1));
        If n_Count = 0 Then
          n_Count := 1;
        End If;
      
        Begin
          Select ��ҩ����
          Into v_��ҩ����
          From (Select �Ǽ�ʱ��, ��ҩ����
                 From ������ü�¼ A
                 Where �շ���� In ('5', '6', '7', '4') And ����id = r_Bill.����id And �Ǽ�ʱ�� Between Sysdate - n_Count And Sysdate And
                       ��¼���� = 2 And ִ�в���id = r_Bill.ִ�в���id And ��ҩ���� Is Not Null And Exists
                  (Select 1
                        From δ��ҩƷ��¼
                        Where a.No = NO And ���� In (9, 25) And �ⷿid + 0 = r_Bill.ִ�в���id And ����id + 0 = r_Bill.����id) And
                       Exists
                  (Select 1
                        From ��ҩ����
                        Where Nvl(�ϰ��, 0) = 1 And ���� = a.��ҩ���� And Nvl(ר��, 0) = 0 And ҩ��id = r_Bill.ִ�в���id)
                 Order By �Ǽ�ʱ�� Desc)
          Where Rownum <= 1;
        Exception
          When Others Then
            v_��ҩ���� := Null;
        End;
        If v_��ҩ���� Is Null Then
          v_��ҩ���� := Zl_Get��ҩ����(r_Bill.ִ�в���id);
        End If;
      
      End If;
      If n_Havedata = 0 Then
        c_��ҩ����.Extend;
        c_��ҩ����(c_��ҩ����.Count).ҩ��id := r_Bill.ִ�в���id;
        c_��ҩ����(c_��ҩ����.Count).��ҩ���� := v_��ҩ����;
      End If;
    End If;
  
    Update ������ü�¼
    Set ��¼״̬ = 1, ��ҩ���� = v_��ҩ����, ����Ա��� = ����Ա���_In, ����Ա���� = ����Ա����_In, �Ǽ�ʱ�� = d_Date --�Ѳ�����ҩƷ��¼��ʱ�䲻�� 
    Where ID = r_Bill.Id;
  
    --ҩƷ�շ���¼.��������
    Update ҩƷ�շ���¼
    Set �������� = Decode(Sign(Nvl(�������, d_Date) - d_Date), -1, ��������, d_Date)
    Where NO = No_In And ���� In (9, 25) And ����id = r_Bill.Id;
  
    --�������
  
    Update �������
    Set ������� = Nvl(�������, 0) + Nvl(r_Bill.ʵ�ս��, 0)
    Where ����id = r_Bill.����id And ���� = 1 And ���� = 1;
  
    If Sql%RowCount = 0 Then
      Insert Into ������� (����id, ����, ����, �������, Ԥ�����) Values (r_Bill.����id, 1, 1, r_Bill.ʵ�ս��, 0);
    End If;
  
    --����δ�����
    Update ����δ�����
    Set ��� = Nvl(���, 0) + Nvl(r_Bill.ʵ�ս��, 0)
    Where ����id = r_Bill.����id And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���˿���id, 0) = Nvl(r_Bill.���˿���id, 0) And
          Nvl(��������id, 0) = Nvl(r_Bill.��������id, 0) And Nvl(ִ�в���id, 0) = Nvl(r_Bill.ִ�в���id, 0) And
          ������Ŀid + 0 = r_Bill.������Ŀid And ��Դ;�� + 0 = r_Bill.�����־;
  
    If Sql%RowCount = 0 Then
      Insert Into ����δ�����
        (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
      Values
        (r_Bill.����id, Null, Null, r_Bill.���˿���id, r_Bill.��������id, r_Bill.ִ�в���id, r_Bill.������Ŀid, r_Bill.�����־,
         Nvl(r_Bill.ʵ�ս��, 0));
    End If;
  
    If r_Bill.ҽ����� Is Not Null Then
      v_ҽ��ids := v_ҽ��ids || ',' || r_Bill.ҽ�����;
    End If;
  
  End Loop;

  --����ҽ�����ͼƷ�״̬
  If v_ҽ��ids Is Not Null Then
    v_ҽ��ids := Substr(v_ҽ��ids, 2);
    Zl_ҽ������_�Ʒ�״̬_Update(0, 2, 1, No_In, v_ҽ��ids);
  End If;
  --���·�ҩ����
  For n_Step In 1 .. c_��ҩ����.Count Loop
    Update ҩƷ�շ���¼
    Set ��ҩ���� = c_��ҩ����(n_Step).��ҩ����
    Where �ⷿid = c_��ҩ����(n_Step).ҩ��id And ���� In (9, 25) And NO = No_In And
          ����id + 0 In (Select ID
                       From ������ü�¼
                       Where ��¼���� = 2 And ��¼״̬ = 1 And NO = No_In And �շ���� In ('5', '6', '7'));
  
    Update δ��ҩƷ��¼
    Set ��ҩ���� = c_��ҩ����(n_Step).��ҩ����
    Where �ⷿid = c_��ҩ����(n_Step).ҩ��id And ���� In (9, 25) And NO = No_In;
  End Loop;

  --�ⷿ�е�ҩƷ��ȫ��������Ϊ���շ�
  Update δ��ҩƷ��¼
  Set ���շ� = 1, �������� = d_Date
  Where NO = No_In And ���� = 9 And Nvl(���շ�, 0) = 0 And
        Nvl(�ⷿid, 0) Not In
        (Select Distinct Nvl(ִ�в���id, 0)
         From ������ü�¼
         Where ��¼���� = 2 And NO = No_In And �շ���� In ('5', '6', '7') And ��¼״̬ = 0);

  Update δ��ҩƷ��¼
  Set ���շ� = 1, �������� = d_Date
  Where NO = No_In And ���� = 25 And Nvl(���շ�, 0) = 0 And
        Nvl(�ⷿid, 0) Not In (Select Distinct Nvl(ִ�в���id, 0)
                             From ������ü�¼
                             Where ��¼���� = 2 And NO = No_In And �շ���� = '4' And ��¼״̬ = 0);

  --����������������Զ�����
  If zl_GetSysParameter(92) = '1' Then
    For r_Stuff In c_Stuff Loop
      If n_���Ϻ� Is Null Then
        n_���Ϻ� := Nextno(20);
      End If;
    
      If r_Stuff.�ⷿid <> Nvl(n_�ⷿid, 0) Then
        If Nvl(n_�ⷿid, 0) <> 0 And v_�շ�ids Is Not Null Then
          v_�շ�ids := Substr(v_�շ�ids, 2);
          Zl_ҩƷ�շ���¼_��������(v_�շ�ids, n_�ⷿid, ����Ա����_In, Sysdate, 1, ����Ա����_In, n_���Ϻ�, ����Ա����_In);
        End If;
      
        n_�ⷿid  := r_Stuff.�ⷿid;
        v_�շ�ids := Null;
      End If;
    
      v_�շ�ids := v_�շ�ids || '|' || r_Stuff.Id || ',0';
    End Loop;
    If Nvl(n_�ⷿid, 0) <> 0 And v_�շ�ids Is Not Null Then
      v_�շ�ids := Substr(v_�շ�ids, 2);
      Zl_ҩƷ�շ���¼_��������(v_�շ�ids, n_�ⷿid, ����Ա����_In, Sysdate, 1, ����Ա����_In, n_���Ϻ�, ����Ա����_In);
    End If;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_������ʼ�¼_Verify;
/

--92026:��͢��,2017-11-01,����¼��ҽ��֧��ִ��
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

--112953:���Ʊ�,2017-09-11,ҩƷ˵����֪ʶ��
CREATE OR REPLACE Function Zl_Drugexplain_Readlob
(
  Key_In In Varchar2,
  Col_In In Varchar2,
  Pos_In     In Number
) Return Varchar2 Is
  l_Clob   Clob;
  v_Buffer Varchar2(32767);
  n_Amount Number := 2000;
  n_Offset Number := 1;
Begin

  If Col_In = '��ѧ����' Then
    Select t.��ѧ���� Into l_Clob From ҩƷ˵���� T Where ID = Key_In;
  Elsif Col_In = '��״' Then
    Select t.��״ Into l_Clob From ҩƷ˵���� T Where ID = Key_In;
  Elsif Col_In = 'ҩ����' Then
    Select t.ҩ���� Into l_Clob From ҩƷ˵���� T Where ID = Key_In;
  Elsif Col_In = 'ҩ������ѧ' Then
    Select t.ҩ������ѧ Into l_Clob From ҩƷ˵���� T Where ID = Key_In;
  Elsif Col_In = '��Ӧ֢' Then
    Select t.��Ӧ֢ Into l_Clob From ҩƷ˵���� T Where ID = Key_In;
  Elsif Col_In = '�÷�����' Then
    Select t.�÷����� Into l_Clob From ҩƷ˵���� T Where ID = Key_In;
  Elsif Col_In = '������Ӧ' Then
    Select t.������Ӧ Into l_Clob From ҩƷ˵���� T Where ID = Key_In;
  Elsif Col_In = '����֢' Then
    Select t.����֢ Into l_Clob From ҩƷ˵���� T Where ID = Key_In;
  Elsif Col_In = 'ע������' Then
    Select t.ע������ Into l_Clob From ҩƷ˵���� T Where ID = Key_In;
  Elsif Col_In = '�и���ҩ' Then
    Select t.�и���ҩ Into l_Clob From ҩƷ˵���� T Where ID = Key_In;
  Elsif Col_In = '��ͯ��ҩ' Then
    Select t.��ͯ��ҩ Into l_Clob From ҩƷ˵���� T Where ID = Key_In;
  Elsif Col_In = '�໥����' Then
    Select t.�໥���� Into l_Clob From ҩƷ˵���� T Where ID = Key_In;
  Elsif Col_In = 'ҩ�����' Then
    Select t.ҩ����� Into l_Clob From ҩƷ˵���� T Where ID = Key_In;
  Elsif Col_In = '��������' Then
    Select t.�������� Into l_Clob From ҩƷ˵���� T Where ID = Key_In;
  End If;

  n_Offset := n_Offset + Pos_In * n_Amount;
  If l_Clob Is Null Then
    v_Buffer := Null;
  Else
    Dbms_Lob.Read(l_Clob, n_Amount, n_Offset, v_Buffer);
  End If;
  Return v_Buffer;
Exception
  When No_Data_Found Then
    Return Null;
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Drugexplain_Readlob;
/

--109619:��ҵ��,2017-10-20,���۸������������
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
           a.�ڲ�����, b.�Ƿ���, a.����, a.Ƶ��, a.ժҪ, Nvl(a.����id, 0) As ����id
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
      Select Nvl(ʵ������, 0), ƽ���ɱ���
      Into n_�������, n_���ƽ����
      From ҩƷ���
      Where ���� = 1 And �ⷿid = v_Detail.�ⷿid And ҩƷid = v_Detail.ҩƷid And Nvl(����, 0) = v_Detail.����;
    Exception
      When Others Then
        n_������� := 0;
    End;
  
    If n_���ƽ���� Is Null Or n_���ƽ���� < 0 Then
      Select �ɱ��� Into n_���ƽ���� From ҩƷ��� Where ҩƷid = v_Detail.ҩƷid;
    
      If n_���ƽ���� Is Null Or n_���ƽ���� < 0 Then
        n_���ƽ���� := 0;
      End If;
    End If;
  
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
        
          Insert Into ҩƷ�����Ϣ
            (ҩƷid, �ⷿid, ����, �������)
            Select v_Detail.ҩƷid, v_Detail.�ⷿid, v_Detail.����, v_Detail.�������
            From Dual
            Where Not Exists (Select 1
                   From ҩƷ�����Ϣ
                   Where ҩƷid = v_Detail.ҩƷid And �ⷿid = v_Detail.�ⷿid And ���� = v_Detail.����);
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
          
            Insert Into ҩƷ�����Ϣ
              (ҩƷid, �ⷿid, ����, �������)
              Select v_Detail.ҩƷid, v_Detail.�ⷿid, v_Detail.����, v_Detail.�������
              From Dual
              Where Not Exists (Select 1
                     From ҩƷ�����Ϣ
                     Where ҩƷid = v_Detail.ҩƷid And �ⷿid = v_Detail.�ⷿid And ���� = v_Detail.����);
          End If;
        
          --����ҩƷ���Ŷ��ձ��еļ۸�
          If v_Detail.ժҪ = '�ɱ��۵���' Then
            Update ҩƷ���Ŷ��� Set �ɱ��� = n_�ɱ��� Where ҩƷid = v_Detail.ҩƷid And Nvl(����, 0) = v_Detail.����;
          End If;
        End If;
      Elsif v_Detail.���� = 13 Then
        --����=13 ���ۼ�������¼ ͬ�����µĽ��Ͳ�ۣ����Բ���Ҫ����ƽ���ɱ���
        If v_Detail.����id = 0 Then
          Update ҩƷ���
          Set ���ۼ� = Decode(n_ʱ�۷���, 1, n_���ۼ�, Null), ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) + n_���۽��, ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + n_���
          Where ���� = 1 And �ⷿid = v_Detail.�ⷿid And ҩƷid = v_Detail.ҩƷid And Nvl(����, 0) = v_Detail.����;
        Else
          --��������ʱ���������ۼ�
          Update ҩƷ���
          Set ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) + n_���۽��, ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + n_���
          Where ���� = 1 And �ⷿid = v_Detail.�ⷿid And ҩƷid = v_Detail.ҩƷid And Nvl(����, 0) = v_Detail.����;
        End If;
        If Sql%RowCount = 0 Then
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, ��������, ʵ������, ʵ�ʽ��, ʵ�ʲ��, ���ۼ�)
          Values
            (v_Detail.�ⷿid, v_Detail.ҩƷid, v_Detail.����, 1, 0, 0, n_���۽��, n_���۽��, Decode(n_ʱ�۷���, 1, n_���ۼ�, Null));
        
          Insert Into ҩƷ�����Ϣ
            (ҩƷid, �ⷿid, ����, �������)
            Select v_Detail.ҩƷid, v_Detail.�ⷿid, v_Detail.����, v_Detail.�������
            From Dual
            Where Not Exists (Select 1
                   From ҩƷ�����Ϣ
                   Where ҩƷid = v_Detail.ҩƷid And �ⷿid = v_Detail.�ⷿid And ���� = v_Detail.����);
        End If;
      
        --����ҩƷ���Ŷ��ձ��еļ۸�
        If v_Detail.ժҪ = 'ҩƷ����' Then
          Update ҩƷ���Ŷ��� Set �ۼ� = n_���ۼ� Where ҩƷid = v_Detail.ҩƷid And Nvl(����, 0) = v_Detail.����;
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
        
          --104843��34�汾���ʱ���������ۼۣ�������޿�����ͨ�������¿���¼�����˼۸�����Ѵ����˾Ͳ����£�
          /*          --�⹺��������������ʱ
          If (v_Detail.���� = 1 And v_Detail.��¼״̬ = 1 And �������_In = 0) Or (v_Detail.���� = 4 And v_Detail.��¼״̬ = 1) Then
            Update ҩƷ���
            Set ���ۼ� = Decode(n_ʱ�۷���, 1, n_���ۼ�, Null)
            Where �ⷿid = v_Detail.�ⷿid And ҩƷid = v_Detail.ҩƷid And Nvl(����, 0) = v_Detail.���� And ���� = 1;
          End If;*/
        
          --�����������Ҫ����ɱ���
          --�⹺�˻���������˺����г���ҵ�񲻸���ƽ���ɱ��ۣ����ֵ�ǰ�۸�
          If (v_Detail.���� = 1 And v_Detail.��ҩ��ʽ = 1) Or Mod(v_Detail.��¼״̬, 3) = 2 Or (v_Detail.���� = 1 And �������_In = 1) Then
            Null;
          Else
            --���ܽ��/��������ʽ����ƽ���ɱ��۶����ã����-��ۣ�/������Ϊ�����ݵ�׼ȷ��
            n_������ := (n_������� + n_ʵ������);
            If n_������ <> 0 And v_Detail.���� = 0 Then
              --104843���������Ĳ����㣬�����Ĳ�����
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
              ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + n_���, ƽ���ɱ��� = Decode(ƽ���ɱ���, Null, n_�ɱ���, ƽ���ɱ���),
              �ϴβɹ��� = Decode(�ϴβɹ���, Null, n_�ɱ���, �ϴβɹ���)
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
        
          Insert Into ҩƷ�����Ϣ
            (ҩƷid, �ⷿid, ����, �������)
            Select v_Detail.ҩƷid, v_Detail.�ⷿid, v_Detail.����, v_Detail.�������
            From Dual
            Where Not Exists (Select 1
                   From ҩƷ�����Ϣ
                   Where ҩƷid = v_Detail.ҩƷid And �ⷿid = v_Detail.�ⷿid And ���� = v_Detail.����);
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

--118817:����,2017-12-21,�������Ĳ�ѯ�쳣
--118364:����,2017-12-19,�����˷�ҩƷ����ҩ�ˡ��͡���ҩ���ڡ�����Ϊ�յ����
--82526:��ҵ��,2017-12-05,������ҩ�������ܷ�ҩ��
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
  δȡҩ_In        In ҩƷ�շ���¼.�Ƿ�δȡҩ%Type := Null,
  ���ܷ�ҩ��_In    In ҩƷ�շ���¼.���ܷ�ҩ��%Type := Null
) Is
  --סԺ����
  Cursor c_Modifybillin Is
    Select a.Id, a.ҩƷid, a.������id, a.���ϵ��, a.����id, Nvl(a.ʵ������, 0) * Nvl(a.����, 1) ����, Nvl(a.���۽��, 0) ���, Nvl(a.����, 0) ����,
           a.��ҩ��λid, a.�ɱ���, a.����, a.����, a.Ч��, a.��������, a.��׼�ĺ�, b.����id, b.���, Nvl(c.��������, Nvl(a.ע��֤��, 0)) ��������,
           Nvl(a.���ۼ�, 0) As ���ۼ�, a.��¼״̬
    From ҩƷ�շ���¼ A, סԺ���ü�¼ B, δ��ҩƷ��¼ C
    Where a.���� = c.���� And a.No = c.No And Nvl(a.�ⷿid, 0) = Nvl(c.�ⷿid, 0) And a.No = No_In And a.���� = Bill_In And
          (a.�ⷿid + 0 = Partid_In Or a.�ⷿid Is Null) And Nvl(a.ժҪ, 'С��') <> '�ܷ�' And a.����id = b.Id And
          Nvl(b.ִ��״̬, 0) <> 1 And Mod(a.��¼״̬, 3) = 1 And a.����� Is Null;

  --���ﲡ��
  Cursor c_Modifybillout Is
    Select a.Id, a.ҩƷid, a.������id, a.���ϵ��, a.����id, Nvl(a.ʵ������, 0) * Nvl(a.����, 1) ����, Nvl(a.���۽��, 0) ���, Nvl(a.����, 0) ����,
           a.��ҩ��λid, a.�ɱ���, a.����, a.����, a.Ч��, a.��������, a.��׼�ĺ�, b.����id, b.���, Nvl(c.��������, Nvl(a.ע��֤��, 0)) ��������,
           Nvl(a.���ۼ�, 0) As ���ۼ�, b.No, b.��¼����, a.��¼״̬
    From ҩƷ�շ���¼ A, ������ü�¼ B, δ��ҩƷ��¼ C
    Where a.���� = c.���� And a.No = c.No And Nvl(a.�ⷿid, 0) = Nvl(c.�ⷿid, 0) And a.No = No_In And a.���� = Bill_In And
          (a.�ⷿid + 0 = Partid_In Or a.�ⷿid Is Null) And Nvl(a.ժҪ, 'С��') <> '�ܷ�' And a.����id = b.Id And
          Nvl(b.ִ��״̬, 0) <> 1 And Mod(a.��¼״̬, 3) = 1 And a.����� Is Null
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
  v_��ҩ��          ҩƷ�շ���¼.��ҩ��%Type;
  v_��ҩ����        ҩƷ�շ���¼.��ҩ����%Type;
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
  Set ��ҩ�� = Decode(��ҩ��_In, Null, ��ҩ��, ��ҩ��_In), ��ҩ���� = Decode(��ҩ��_In, Null, ��ҩ����, Date����ʱ��), ���ܷ�ҩ�� = ���ܷ�ҩ��_In
  Where NO = No_In And ���� = Bill_In And (�ⷿid + 0 = Partid_In Or �ⷿid Is Null) And Mod(��¼״̬, 3) = 1 And ����� Is Not Null;

  --�����˷�ҩƷ����ҩ�ˡ��͡���ҩ���ڡ�����Ϊ�յ����
  Begin
    If ��ҩ��_In Is Null Then
      Select ��ҩ��, ��ҩ����
      Into v_��ҩ��, v_��ҩ����
      From ҩƷ�շ���¼
      Where NO = No_In And ���� = Bill_In And (�ⷿid + 0 = Partid_In Or �ⷿid Is Null) And Mod(��¼״̬, 3) = 1 And
            ��ҩ�� Is Not Null And Rownum = 1
      Order By ��¼״̬ Desc;
    
      Update ҩƷ�շ���¼
      Set ��ҩ�� = v_��ҩ��, ��ҩ���� = v_��ҩ����
      Where NO = No_In And ���� = Bill_In And (�ⷿid + 0 = Partid_In Or �ⷿid Is Null) And Mod(��¼״̬, 3) = 1 And ����� Is Null;
    End If;
  Exception
    When Others Then
      v_��ҩ��   := Null;
      v_��ҩ���� := Null;
  End;

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
      If v_Modifybillout.��¼״̬ = 1 Then
        --ԭʼ��ҩ��¼��ȡ���¼۸�
        n_ƽ���ɱ��� := Round(Zl_Fun_Getoutcost(v_Modifybillout.ҩƷid, v_Modifybillout.����, Partid_In), 5);
      Else
        --��ҩ�ٷ���¼��ȡԭʼ���ݼ۸�
        Select a.�ɱ���
        Into n_ƽ���ɱ���
        From ҩƷ�շ���¼ A, ҩƷ�շ���¼ B
        Where b.Id = v_Modifybillout.Id And a.���� = b.���� And a.No = b.No And a.�ⷿid = b.�ⷿid And a.ҩƷid + 0 = b.ҩƷid And
              a.��� = b.��� And Nvl(a.����, 0) = Nvl(b.����, 0) And (a.��¼״̬ = 1 Or Mod(a.��¼״̬, 3) = 0);
      End If;
    
      Dbl�ɱ���� := Round(n_ƽ���ɱ��� * Nvl(v_Modifybillout.����, 0), Intdigit_In);
      --���۽��
      Dblʵ�ʽ�� := Nvl(v_Modifybillout.���, 0);
      --���
      Dblʵ�ʲ�� := Round(Dblʵ�ʽ�� - Dbl�ɱ����, Intdigit_In);
    
      --����ҩƷ�շ���¼�����۽��ɱ�����ۡ�����˵���Ϣ
      Update ҩƷ�շ���¼
      Set �ɱ��� = n_ƽ���ɱ���, �ɱ���� = Dbl�ɱ����, ��� = Dblʵ�ʲ��, �ⷿid = Partid_In, ��ҩ�� = Decode(��ҩ��_In, Null, ��ҩ��, ��ҩ��_In),
          �˲��� = �˲���_In, �˲����� = v_�˲�����, ��ҩ���� = Decode(��ҩ��_In, Null, ��ҩ����, Date����ʱ��),
          ������ = Decode(У����_In, Null, ������, У����_In), ����� = Decode(People_In, Null, Zl_Username, People_In),
          ������� = Date����ʱ��, ��ҩ��ʽ = ��ҩ��ʽ_In, ע��֤�� = v_Modifybillout.��������, �Ƿ�δȡҩ = δȡҩ_In, ���ܷ�ҩ�� = ���ܷ�ҩ��_In
      Where ID = v_Modifybillout.Id;
    
      If Bln�շ��뷢ҩ���� = 1 Then
        Update ҩƷ���
        Set �������� = Nvl(��������, 0) - Nvl(v_Modifybillout.����, 0), ʵ������ = Nvl(ʵ������, 0) - Nvl(v_Modifybillout.����, 0),
            ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) - Nvl(v_Modifybillout.���, 0), ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) - Dblʵ�ʲ��,
            ƽ���ɱ��� = Decode(ƽ���ɱ���, Null, n_ƽ���ɱ���, ƽ���ɱ���), �ϴβɹ��� = Decode(�ϴβɹ���, Null, n_ƽ���ɱ���, �ϴβɹ���)
        Where �ⷿid + 0 = Partid_In And ҩƷid = v_Modifybillout.ҩƷid And ���� = 1 And Nvl(����, 0) = v_Modifybillout.����;
      Else
        Update ҩƷ���
        Set ʵ������ = Nvl(ʵ������, 0) - Nvl(v_Modifybillout.����, 0), ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) - Nvl(v_Modifybillout.���, 0),
            ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) - Dblʵ�ʲ��, ƽ���ɱ��� = Decode(ƽ���ɱ���, Null, n_ƽ���ɱ���, ƽ���ɱ���),
            �ϴβɹ��� = Decode(�ϴβɹ���, Null, n_ƽ���ɱ���, �ϴβɹ���)
        Where �ⷿid + 0 = Partid_In And ҩƷid = v_Modifybillout.ҩƷid And ���� = 1 And Nvl(����, 0) = v_Modifybillout.����;
      End If;
    
      If Sql%RowCount = 0 Then
        If Bln�շ��뷢ҩ���� = 1 Then
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, ��������, ʵ������, ʵ�ʽ��, ʵ�ʲ��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴβ���, Ч��, �ϴ���������, ��׼�ĺ�, ƽ���ɱ���)
          Values
            (Partid_In, v_Modifybillout.ҩƷid, v_Modifybillout.����, 1, 0 - Nvl(v_Modifybillout.����, 0),
             0 - Nvl(v_Modifybillout.����, 0), 0 - Nvl(v_Modifybillout.���, 0), 0 - Dblʵ�ʲ��, v_Modifybillout.��ҩ��λid,
             n_ƽ���ɱ���, v_Modifybillout.����, v_Modifybillout.����, v_Modifybillout.Ч��, v_Modifybillout.��������,
             v_Modifybillout.��׼�ĺ�, n_ƽ���ɱ���);
        Else
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, ʵ������, ʵ�ʽ��, ʵ�ʲ��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴβ���, Ч��, �ϴ���������, ��׼�ĺ�, ƽ���ɱ���)
          Values
            (Partid_In, v_Modifybillout.ҩƷid, v_Modifybillout.����, 1, 0 - Nvl(v_Modifybillout.����, 0),
             0 - Nvl(v_Modifybillout.���, 0), 0 - Dblʵ�ʲ��, v_Modifybillout.��ҩ��λid, n_ƽ���ɱ���, v_Modifybillout.����,
             v_Modifybillout.����, v_Modifybillout.Ч��, v_Modifybillout.��������, v_Modifybillout.��׼�ĺ�, n_ƽ���ɱ���);
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
      If v_Modifybillin.��¼״̬ = 1 Then
        --ԭʼ��ҩ��¼��ȡ���¼۸�
        n_ƽ���ɱ��� := Round(Zl_Fun_Getoutcost(v_Modifybillin.ҩƷid, v_Modifybillin.����, Partid_In), 5);
      Else
        --��ҩ�ٷ���¼��ȡԭʼ���ݼ۸�
        Select a.�ɱ���
        Into n_ƽ���ɱ���
        From ҩƷ�շ���¼ A, ҩƷ�շ���¼ B
        Where b.Id = v_Modifybillin.Id And a.���� = b.���� And a.No = b.No And a.�ⷿid = b.�ⷿid And a.ҩƷid + 0 = b.ҩƷid And
              a.��� = b.��� And Nvl(a.����, 0) = Nvl(b.����, 0) And (a.��¼״̬ = 1 Or Mod(a.��¼״̬, 3) = 0);
      End If;
    
      Dbl�ɱ���� := Round(n_ƽ���ɱ��� * Nvl(v_Modifybillin.����, 0), Intdigit_In);
      --���۽��
      Dblʵ�ʽ�� := Nvl(v_Modifybillin.���, 0);
      --���
      Dblʵ�ʲ�� := Round(Dblʵ�ʽ�� - Dbl�ɱ����, Intdigit_In);
    
      --����ҩƷ�շ���¼�����۽��ɱ�����ۡ�����˵���Ϣ
      Update ҩƷ�շ���¼
      Set �ɱ��� = n_ƽ���ɱ���, �ɱ���� = Dbl�ɱ����, ��� = Dblʵ�ʲ��, �ⷿid = Partid_In, ��ҩ�� = Decode(��ҩ��_In, Null, ��ҩ��, ��ҩ��_In),
          �˲��� = �˲���_In, �˲����� = v_�˲�����, ��ҩ���� = Decode(��ҩ��_In, Null, ��ҩ����, Date����ʱ��),
          ������ = Decode(У����_In, Null, ������, У����_In), ����� = Decode(People_In, Null, Zl_Username, People_In),
          ������� = Date����ʱ��, ��ҩ��ʽ = ��ҩ��ʽ_In, ע��֤�� = v_Modifybillin.��������, ���ܷ�ҩ�� = ���ܷ�ҩ��_In
      Where ID = v_Modifybillin.Id;
    
      If Bln�շ��뷢ҩ���� = 1 Then
        Update ҩƷ���
        Set �������� = Nvl(��������, 0) - Nvl(v_Modifybillin.����, 0), ʵ������ = Nvl(ʵ������, 0) - Nvl(v_Modifybillin.����, 0),
            ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) - Nvl(v_Modifybillin.���, 0), ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) - Dblʵ�ʲ��,
            ƽ���ɱ��� = Decode(ƽ���ɱ���, Null, n_ƽ���ɱ���, ƽ���ɱ���), �ϴβɹ��� = Decode(�ϴβɹ���, Null, n_ƽ���ɱ���, �ϴβɹ���)
        Where �ⷿid + 0 = Partid_In And ҩƷid = v_Modifybillin.ҩƷid And ���� = 1 And Nvl(����, 0) = v_Modifybillin.����;
      Else
        Update ҩƷ���
        Set ʵ������ = Nvl(ʵ������, 0) - Nvl(v_Modifybillin.����, 0), ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) - Nvl(v_Modifybillin.���, 0),
            ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) - Dblʵ�ʲ��, ƽ���ɱ��� = Decode(ƽ���ɱ���, Null, n_ƽ���ɱ���, ƽ���ɱ���),
            �ϴβɹ��� = Decode(�ϴβɹ���, Null, n_ƽ���ɱ���, �ϴβɹ���)
        Where �ⷿid + 0 = Partid_In And ҩƷid = v_Modifybillin.ҩƷid And ���� = 1 And Nvl(����, 0) = v_Modifybillin.����;
      End If;
    
      If Sql%RowCount = 0 Then
        If Bln�շ��뷢ҩ���� = 1 Then
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, ��������, ʵ������, ʵ�ʽ��, ʵ�ʲ��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴβ���, Ч��, �ϴ���������, ��׼�ĺ�, ƽ���ɱ���)
          Values
            (Partid_In, v_Modifybillin.ҩƷid, v_Modifybillin.����, 1, 0 - Nvl(v_Modifybillin.����, 0),
             0 - Nvl(v_Modifybillin.����, 0), 0 - Nvl(v_Modifybillin.���, 0), 0 - Dblʵ�ʲ��, v_Modifybillin.��ҩ��λid, n_ƽ���ɱ���,
             v_Modifybillin.����, v_Modifybillin.����, v_Modifybillin.Ч��, v_Modifybillin.��������, v_Modifybillin.��׼�ĺ�, n_ƽ���ɱ���);
        Else
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, ʵ������, ʵ�ʽ��, ʵ�ʲ��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴβ���, Ч��, �ϴ���������, ��׼�ĺ�, ƽ���ɱ���)
          Values
            (Partid_In, v_Modifybillin.ҩƷid, v_Modifybillin.����, 1, 0 - Nvl(v_Modifybillin.����, 0),
             0 - Nvl(v_Modifybillin.���, 0), 0 - Dblʵ�ʲ��, v_Modifybillin.��ҩ��λid, n_ƽ���ɱ���, v_Modifybillin.����,
             v_Modifybillin.����, v_Modifybillin.Ч��, v_Modifybillin.��������, v_Modifybillin.��׼�ĺ�, n_ƽ���ɱ���);
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

--104221:��ΰ��,2017-12-07,���֤�ż��
Create Or Replace Function Zl_Fun_Checkidcard
(
  Idcard_In   In varchar2,
  Calcdate_In In Date := Null
) Return varchar2 Is
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
  v_У��λ    varchar2(50);
  v_Pattern   varchar2(500);
  v_Err_Msg   varchar2(2000);
  v_�Ա�      varchar2(100);
  v_����      varchar2(100);
  d_Curr_Time Date;
  d_��������  Date;

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
        d_�������� := To_Date('19' || Substr(Idcard_In, 7, 6), 'yyyy-mm-dd');
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
        d_�������� := To_Date(Substr(Idcard_In, 7, 8), 'yyyy-mm-dd');
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


--118323:��ҵ��,2017-12-12,�޸�ɾ�����ۼ���ʱ��ֵ���Ŀ�洦����©
Create Or Replace Procedure Zl_���ﻮ�ۼ�¼_Delete
(
  No_In       ������ü�¼.No%Type,
  ���_In     Varchar2 := Null,
  �Զ����_In Number := 0
) As
  --���ܣ�ɾ��һ�����ﻮ�۵���
  --��Σ�
  --       ���_In����Ҫ��������ҽ��վ���ϵ���ҩƷ
  --      �Զ����_in���Ƿ��Զ�������۵� zl_���ﻮ�ۼ�¼_clear �ڵ���
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
    If Nvl(�Զ����_In, 0) = 1 Then
      --�Զ�������۵�����ʱ������ֱ���˳�
      Return;
    Else
      v_Err_Msg := 'Ҫɾ���ķ��ü�¼�����ڣ������Ѿ�ɾ�����Ѿ��շѡ�';
      Raise Err_Item;
    End If;
  End If;
  --�Ƿ��Ѿ�ִ��
  If Nvl(n_��ִ��_Count, 0) > 0 Then
    v_Err_Msg := 'Ҫɾ���ķ��ü�¼�а�����ִ�е����ݣ�';
    Raise Err_Item;
  End If;

  --ҽ�����ã��������ִ�е�ҽ��(ע����ִ�е������������,��Ϊ���� ���_IN ����������ý���������)
  --�Զ�������۵�����ʱ������ֻ�ᴫ��ҩƷ���ĵĶ�Ӧ��ţ����Բ��ü��ҽ����
  --������ҽ��������ͬһ��ҽ���м���ҩƷ��Ҳ��������Ŀ����������Ŀ����ִ�л���ִ��ʱ��ҩƷ���ۼ�¼��ɾ������
  If Nvl(�Զ����_In, 0) = 0 Then
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
    If Nvl(�Զ����_In, 0) = 1 Then
      --�Զ�������۵�����ʱ������ֱ���˳�
      Return;
    Else
      v_Err_Msg := 'Ҫɾ���ķ��ü�¼�����ڣ������Ѿ�ɾ�����Ѿ��շѡ�';
      Raise Err_Item;
    End If;
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


--118323:��ҵ��,2017-12-12,�޸�ɾ�����ۼ���ʱ��ֵ���Ŀ�洦����©
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
  For c_ҽ�� In (Select Distinct ҽ�����
               From ������ü�¼
               Where NO = No_In And ��¼���� = 2 And ��¼״̬ = 3 And ҽ����� Is Not Null) Loop
    Select Nvl(Count(*), 0)
    Into n_Count
    From (Select ���, Sum(����) As ʣ������
           From (Select ��¼״̬, Nvl(�۸񸸺�, ���) As ���, Avg(Nvl(����, 1) * ����) As ����
                  From ������ü�¼
                  Where ��¼���� = 2 And ҽ����� + 0 = c_ҽ��.ҽ����� And NO = No_In
                  Group By ��¼״̬, Nvl(�۸񸸺�, ���))
           Group By ���
           Having Sum(����) <> 0);
  
    If n_Count = 0 Then
      Delete From ����ҽ������ Where ҽ��id = c_ҽ��.ҽ����� And ��¼���� = 2 And NO = No_In;
    End If;
  End Loop;

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

---------------------------------------------------------------------------------------------------
--����ϵͳ�������İ汾��
-------------------------------------------------------------------------------------------------------
--113763:����һ,2017-12-07,DBA���������޸�
Insert Into Zltools.Zlfilesupgrade
  (���, ��������, ��װ·��, �ļ�����, �ļ���, �汾��, �޸�����, ����ϵͳ, ҵ�񲿼�, Md5, �ļ�˵��, �Զ�ע��, ǿ�Ƹ���, ���Ӱ�װ·��)
  Select ���, To_Date('2017-07-05 17:22:54', 'yyyy-mm-dd hh24:mi:ss'), '[APPSOFT]', 0, 'ZLDBATOOLS.EXE', Null, Null, Null,
         Null, Null, '��������:DBA�����ߵ���ִ���ļ�', 0, 0, Null
  From Dual A, (Select Nvl(Max(To_Number(���)), 0) + 1 ��� From zlFilesUpgrade) B
  Where Not Exists (Select 1 From Zltools.Zlfilesupgrade Where Upper(�ļ���) = 'ZLDBATOOLS.EXE');
--00000:��˶,2017-12-27,�ļ��嵥����
Update Zltools.Zlfilesupgrade
Set ��װ·�� = '[APPSOFT]'
Where Upper(�ļ���) = 'ZLRISDUMPTOOL.EXE' And Not Exists
 (Select 1 From Zltools.Zlfilesupgrade Where Upper(�ļ���) = 'ZLRISDUMPTOOL.EXE' And Upper(��װ·��) = '[APPSOFT]');
Delete Zltools.Zlfilesupgrade Where Upper(�ļ���) = 'ZLRISDUMPTOOL.EXE' And Upper(��װ·��) <> '[APPSOFT]';
Delete Zltools.Zlfilesupgrade Where Upper(�ļ���) In ('ZL9PEISDEVANALYSE', 'ZL9PEISINSTRUMENT');

--ϵͳ�汾��
Update zlSystems Set �汾��='10.34.140' Where ���=&n_System;
--�����汾��
Commit;

