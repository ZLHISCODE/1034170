--���ű�֧�ִ�ZLHIS+ v10.11.0 ������ v10.12.0
--����ϵͳ�����ߵ�¼PLSQL��ִ�����нű�
--�ű�ִ�к����ֹ�������������
-------------------------------------------------------------------------------
--�ṹ��������
-------------------------------------------------------------------------------
--6859
Alter Table ҩƷ�ⷿ��λ Modify ���� Varchar2(5);

--PACS
Alter Table Ӱ�����¼ Add ��ϵ�绰 Varchar2(20);
Alter Table Ӱ�����¼ Drop Constraint Ӱ�����¼_UQ_���� Cascade;
Alter Table HӰ�����¼ Drop Constraint HӰ�����¼_UQ_���� Cascade;
Create Index Ӱ�����¼_IX_���� on Ӱ�����¼ (����, Ӱ�����) PCTFREE 10 TABLESPACE zl9CisRec;
Create Index Ӱ����ʱ��¼_IX_���� on Ӱ����ʱ��¼ (����, Ӱ�����) PCTFREE 10 TABLESPACE zl9CisRec;

--6950
Alter Table ҩƷ�շ���¼ Add ��׼�ĺ� VARCHAR2(40);
Alter Table HҩƷ�շ���¼ Add ��׼�ĺ� VARCHAR2(40);
Alter Table ҩƷ��� Add ��׼�ĺ� VARCHAR2(40);
--7135
ALTER TABLE ���˹���ҩ�� DROP CONSTRAINT ���˹���ҩ��_UQ_����ҩ��ID;
--7303
Create Or Replace View ������ As 
Select ����id, ��ҳId,����id,������� As ������Ϣ,�������, 
   ��Ժ���,��ϴ���,�������, �Ƿ�δ��, �Ƿ�����
From ������ϼ�¼ Where ��¼��Դ=2;

--7163
Create Table ���˵�����¼(
    ����ID      NUMBER(18),
    ������      VARCHAR2(20),
    ������      NUMBER(16,5),
	��������    NUMBER(1),
    ����Ա���  VARCHAR2(6),
    ����Ա����  VARCHAR2(20),
    ����ʱ��    Date
    )
    TABLESPACE zl9Patient
    PCTFREE 10 PCTUSED 60 STORAGE (NEXT 8K PCTINCREASE 0 MAXEXTENTS UNLIMITED);

ALTER TABLE ���˵�����¼ ADD CONSTRAINT ���˵�����¼_PK PRIMARY KEY (����ID,����ʱ��) USING INDEX PCTFREE 5 TABLESPACE zl9Patient;
ALTER TABLE ���˵�����¼ ADD CONSTRAINT ���˵�����¼_FK_����ID FOREIGN KEY (����ID) REFERENCES ������Ϣ(����ID) ON DELETE CASCADE;

--7149
Alter Table ҩƷ�ɹ��ƻ� Modify �ڼ� Varchar(8);

--���� 2005-12-16 ����ҽ���������ݽṹ
ALTER TABLE ���ս����¼ ADD (������ˮ�� VARCHAR2(30),����ʱ�� DATE ,����վ VARCHAR2(50),�汾�� VARCHAR2(15));
ALTER TABLE ���ս����¼ ADD (ҽ����� VARCHAR2(3));
ALTER TABLE ���ս����¼ ADD (����ID NUMBER(18));
ALTER TABLE ���ս����¼ ADD (�������� VARCHAR2(100));
ALTER TABLE ���ս����¼ ADD (����֢ VARCHAR2(200));
ALTER TABLE ���ս����¼ MODIFY ��ע VARCHAR2(500);

--���˲���ID���������ͬʱ�޸Ĳ�����ݺϲ�����
CREATE TABLE ����ǼǼ�¼(
	���� NUMBER(18),
	����ID NUMBER(18),
	��ҳID NUMBER(18),
	����ʱ�� DATE ,
	״̬ NUMBER(2),		--1-������;0-δ����
	ҽ����� VARCHAR2(3),
	�ʻ���� NUMBER(16,5),
	����ID NUMBER(18),
	�������� VARCHAR2(100),
	����֢ VARCHAR2(200),
	IC����Ϣ VARCHAR2(200),
	HIS��ˮ�� VARCHAR2(30),
	YB��ˮ�� VARCHAR2(30),
	��¼ID NUMBER(18),	--����ID��������ô��ֶ���������סԺ����
	��ע VARCHAR2(200));
ALTER TABLE ����ǼǼ�¼ ADD CONSTRAINT ����ǼǼ�¼_PK PRIMARY KEY (����,����ID,����ʱ��);
ALTER TABLE ����ǼǼ�¼ ADD CONSTRAINT ����ǼǼ�¼_FK_���� FOREIGN KEY (����) REFERENCES �������(���);
ALTER TABLE ����ǼǼ�¼ ADD CONSTRAINT ����ǼǼ�¼_FK_����ID FOREIGN KEY (����ID) REFERENCES ������Ϣ(����ID);


--�¸��ݣ����
Alter Table �����Ͻ��� Add �Ƿ񼲲� Number(1);
Alter Table �����Ͻ��� Drop Column ����;
Alter Table �����Ͻ��� Drop Column ����id;

Alter Table �����Ա���� Drop Column ����id;

Update �����Ͻ��� Set �Ƿ񼲲�=1 Where ĩ��=1;
--}�¸���

--�������ļӳɷ���
CREATE TABLE ���ϼӳɷ���(
	���		number(18), 
	��ͼ�		number(16,5), 
	��߼�		number(16,5), 
	�ӳ���		number(16,5), 
	˵��		varchar2(50))
    TableSpace zl9BaseItem
    PCTFREE 5 PCTUSED 90 STORAGE (NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED);
ALTER TABLE ���ϼӳɷ��� ADD CONSTRAINT ���ϼӳɷ���_PK PRIMARY KEY (���) USING INDEX PCTFREE 5 TABLESPACE zl9BaseItem;

--7157(ZT)
Alter Table ����ҽ���Ƽ� Add ִ�п���ID Number(18);
Alter Table ����ҽ���Ƽ� 
    Add CONSTRAINT ����ҽ���Ƽ�_FK_ִ�п���ID
    Foreign Key (ִ�п���ID) 
    References ���ű�(ID);

--7108
Alter Table ������ĿĿ¼ Modify ���� VARCHAR2(20);
Alter Table �շ���ĿĿ¼ Modify ���� VARCHAR2(20);

--7058,7053
alter table ���˱䶯��¼ add ���� VARCHAR2(20);
alter table ���˱䶯��¼ add ����ҽʦ VARCHAR2(20);
alter table ���˱䶯��¼ add ����ҽʦ VARCHAR2(20);

--7038
Alter Table ҩƷ���� Add Ʒ��ҽ�� Number(1);

alter table �ѱ���ϸ add ���㷽�� number(1) default 0;

--ҽ�����ݹ�����:�������ƹ�ʽ�༭�ķ�ʽ�������´�ҽ��ʱ,ҽ�����ݵ����ɹ���
Create Table ҽ�����ݶ���(
    ������� VARCHAR2(1),
    ҽ������ VARCHAR2(500))
    TABLESPACE zl9BaseItem
    PCTFREE 10 PCTUSED 85 STORAGE (NEXT 255 PCTINCREASE 0 MAXEXTENTS UNLIMITED);
ALTER TABLE ҽ�����ݶ��� ADD CONSTRAINT ҽ�����ݶ���_PK PRIMARY KEY (�������) USING INDEX PCTFREE 5 TABLESPACE zl9BaseItem;

--����ǩ���ṹ����
CREATE SEQUENCE ��Ա֤���¼_ID START WITH 1;
CREATE TABLE ��Ա֤���¼(
	ID NUMBER(18),
	��ԱID NUMBER(18),
	CertDN VARCHAR2(300),
	CertSN VARCHAR2(100),
	SignCert VARCHAR2(2000),
	EncCert VARCHAR2(2000),
	ע��ʱ�� DATE)
    TABLESPACE zl9BaseItem
    PCTFREE 5 PCTUSED 90 STORAGE (NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED);
ALTER TABLE ��Ա֤���¼ ADD CONSTRAINT ��Ա֤���¼_PK PRIMARY KEY(ID) USING INDEX PCTFREE 5 TABLESPACE zl9BaseItem;
Alter Table ��Ա֤���¼ Add CONSTRAINT ��Ա֤���¼_FK_��ԱID Foreign Key (��ԱID) References ��Ա��(ID);

CREATE SEQUENCE ҽ��ǩ����¼_ID START WITH 1;
CREATE TABLE ҽ��ǩ����¼(
	ID NUMBER(18),
	ǩ������ NUMBER(2),--��1��ʼ˳����,���ڷ�ֹ���Ʒ������ҽ��Դ�ı����ɹ������仯
	ǩ����Ϣ VARCHAR2(2000),
	֤��ID	NUMBER(18),
	ǩ��ʱ�� DATE,
    ǩ���� VARCHAR2(20))
    TABLESPACE zl9CisRec
    PCTFREE 15 PCTUSED 70 
	STORAGE (NEXT 1K PCTINCREASE 0 MAXEXTENTS UNLIMITED);
Alter Table ҽ��ǩ����¼ Add CONSTRAINT ҽ��ǩ����¼_PK Primary Key (ID) USING INDEX PCTFREE 5 TABLESPACE zl9CisRec;
Alter Table ҽ��ǩ����¼ Add CONSTRAINT ҽ��ǩ����¼_FK_֤��ID Foreign Key (֤��ID) References ��Ա֤���¼(ID);
CREATE INDEX ҽ��ǩ����¼_IX_֤��ID ON ҽ��ǩ����¼(֤��ID) PCTFREE 10 TABLESPACE zl9CisRec
/

--ҽ�����ݽṹ����
Alter Table ����ҽ����¼ Add(���� VARCHAR2(20),�Ա� VARCHAR2(4),���� VARCHAR2(10));
Alter Table ����ҽ��״̬ Add ǩ��ID Number(18);
Alter Table ����ҽ��״̬ Add CONSTRAINT ����ҽ��״̬_FK_ǩ��ID Foreign Key (ǩ��ID) References ҽ��ǩ����¼(ID);
CREATE INDEX ����ҽ��״̬_IX_ǩ��ID ON ����ҽ��״̬(ǩ��ID) PCTFREE 10 TABLESPACE zl9CisRec
/

--��ʷ�����ݽṹ
Create Table Hҽ��ǩ����¼ Tablespace zl9History As Select * From ҽ��ǩ����¼ where 1=0;
Alter Table Hҽ��ǩ����¼ Add Constraint Hҽ��ǩ����¼_PK Primary Key (ID) Using Index Pctfree 5 Tablespace zl9History;
Alter Table Hҽ��ǩ����¼ Add CONSTRAINT Hҽ��ǩ����¼_FK_֤��ID Foreign Key (֤��ID) References ��Ա֤���¼(ID);
CREATE INDEX Hҽ��ǩ����¼_IX_֤��ID ON Hҽ��ǩ����¼(֤��ID) PCTFREE 10 TABLESPACE zl9History
/

Alter Table H����ҽ����¼ Add(���� VARCHAR2(20),�Ա� VARCHAR2(4),���� VARCHAR2(10));
Alter Table H����ҽ��״̬ Add ǩ��ID Number(18);
Alter Table H����ҽ��״̬ Add CONSTRAINT H����ҽ��״̬_FK_ǩ��ID Foreign Key (ǩ��ID) References Hҽ��ǩ����¼(ID);
CREATE INDEX H����ҽ��״̬_IX_ǩ��ID ON H����ҽ��״̬(ǩ��ID) PCTFREE 10 TABLESPACE zl9History
/

-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--7427
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1342,'����',user,'������Ʊ�','SELECT');

--�����ٴ��������ݰѱ���Ϊ1λ��3λ���ٴ���������Ϊ2λ��4λ����ǰ���0
alter table �ٴ����� drop CONSTRAINT �ٴ�����_FK_��������;
update �ٴ����� set ����='0'||���� where length(����)=1 or length(����)=3;
update �ٴ����� set ��������='0'||�������� where length(��������)=1 or length(��������)=3;
ALTER TABLE �ٴ����� ADD CONSTRAINT �ٴ�����_FK_�������� FOREIGN KEY (��������) REFERENCES �ٴ�����(����) ON DELETE CASCADE;
Insert into �ٴ�����(����,����,����,���) Values ('61','��֢�໤��(�ۺ�)','ZZJHS',163);
Insert into �ٴ�����(����,����,����,���) Values ('79','����','ZZJHS',164);
Insert into �ٴ�����(����,����,����,���) Values ('99','�������','ZZJHS',165);
Insert into �ٴ�����(����,����,����,���) Values ('9901','��Ⱦ(����)��','ZZJHS',166);

--LIS����
UPDATE ZLPrograms SET ����='������ʷ��¼��ѯ' WHERE ���=1210;

--�������ļӳɷ���
insert into ���ϼӳɷ��� (���, ��ͼ�, ��߼�, �ӳ���, ˵��) values (1, null, 500, 10, null);
insert into ���ϼӳɷ��� (���, ��ͼ�, ��߼�, �ӳ���, ˵��) values (2, 500, 2000, 8, null);
insert into ���ϼӳɷ��� (���, ��ͼ�, ��߼�, �ӳ���, ˵��) values (3, 2000, 5000, 5, null);
insert into ���ϼӳɷ��� (���, ��ͼ�, ��߼�, �ӳ���, ˵��) values (4, 5000, null, 2, null);

--�̶�������ҩ����
Insert Into ҩƷ����(����,����,����)
Select * From (
	Select zl_Incstr(Max(����)),'����','FJ' From ҩƷ����)
Where Not Exists(Select ���� From ҩƷ���� Where ����='����');

--7058
Update ���� Set ����='��' ,����='Z'  Where ����='��';

--7058 �������������ܴܺ�,ȱʡֻ������Ժ��������,�������������Ƿ�������Ժ��������
Update ������ҳ Set ��Ժ����='��' Where ��Ժ����='��' And ��Ժ���� IS Null;
Update ������ҳ Set ��ǰ����='��' Where ��ǰ����='��' And ��Ժ���� IS Null;

--�¸���
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'ZL_�����Ŀ����_EMPTY','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1861,'����',user,'zl_GetTextRows','EXECUTE');

--7058  ��ʼ��Ժ�����ڲ��˱䶯��¼�еĲ���
Create Or Replace Procedure zl_���˱䶯��¼_����
AS 
Begin
    For r_InPatient IN
        ( Select ����ID,��ҳID,��ǰ���� From ������ҳ Where ��Ժ���� IS Null)
    Loop
        Update ���˱䶯��¼ Set ����=r_InPatient.��ǰ���� Where ����ID=r_InPatient.����ID And ��ҳID=r_InPatient.��ҳID;
    End Loop;
EXCEPTION
    WHEN OTHERS THEN
        ZL_ERRORCENTER (SQLCODE, SQLERRM);
End zl_���˱䶯��¼_����;
/
Execute zl_���˱䶯��¼_����;
DROP Procedure zl_���˱䶯��¼_����;


----�⹺��⡢������⡢�ƿ�����в��ֳ�������ۼ۽��ɱ�����ۼ��������������7126����Ǽǣ�
--������ֻ�ڹ̶������·�������Ӱ���⹺��⡢������⡢�ƿ��������ϸ���ݺͿ������ݡ�
--����ֻӰ������ϸ�Ľ��û��Ӱ����������ֻ�ṩ�������ݵ���ϸ��������ű���
--���Ҫ��������ʱ��ҩƷ��ͨ����������������ҩƷ��ͨ���̵�������
--ִ�иýű����ܻ���ֵ����⣺����С���������⣬���ܽű��������������ݡ��û����Լ������Ƿ�ִ�и��������̡�

CREATE OR REPLACE PROCEDURE ZL_����������ϸ��¼
IS
  	INTDIGIT NUMBER;
    INT������¼ Number;
BEGIN
    --��ȡ���С��λ��
    SELECT NVL(����ֵ,ȱʡֵ) INTO INTDIGIT FROM ϵͳ������ WHERE ������='���ý���λ��';
    
    --����������ϸ
    Update ҩƷ�շ���¼ Set �ɱ����=round(ʵ������*�ɱ���,INTDIGIT),���۽��=round(ʵ������*���ۼ�,INTDIGIT),
    ���=round(ʵ������*���ۼ�,INTDIGIT)-round(ʵ������*�ɱ���,INTDIGIT)
    Where ���� In(1,4,6) And ���۽��<>round(���ۼ�*ʵ������,INTDIGIT) And Mod(��¼״̬,3)=2;
    
    --�����������ȫ������������޷���ȫ�������������Ѳ����µ����һ��������¼��ȥ
    For V_������� In(Select ����,No,ҩƷid,ʵ������,�ɱ����,���۽�� From 
        (Select ����,No,ҩƷid,Sum(ʵ������) ʵ������,Sum(�ɱ����) �ɱ����,Sum(���۽��) ���۽��  From ҩƷ�շ���¼ Where ���� In(1,4,6) 
        Group By ����,No ,ҩƷid Having Sum(ʵ������)=0) Where �ɱ����<>0 Or ���۽��<>0) Loop 
        
        --ȡ��������¼
        Select Max(��¼״̬) Into INT������¼ From ҩƷ�շ���¼ Where ����=V_�������.���� And No=V_�������.No And 
        ҩƷid=V_�������.ҩƷid And Mod(��¼״̬,3)=2;
        
        --�������
        Update ҩƷ�շ���¼ Set �ɱ����=�ɱ����-V_�������.�ɱ����,���۽��=���۽��-V_�������.���۽��,
        ���=���-( V_�������.���۽��-V_�������.�ɱ����)
        Where ����=V_�������.���� And No=V_�������.No And ҩƷid=V_�������.ҩƷid And ��¼״̬=INT������¼;
        
    End Loop;
    
    Commit;
EXCEPTION
    WHEN OTHERS THEN
        ZL_ERRORCENTER (SQLCODE, SQLERRM);
END ZL_����������ϸ��¼;
/
--������������Ƿ�ִ��
--Execute ZL_����������ϸ��¼;
--Drop Procedure  ZL_����������ϸ��¼;
----���ֳ�������ۼ۽��ɱ�����ۼ��������������7126����Ǽǣ�

--7049
Insert Into �������ʷ���(����,����,����,������,˵��)
Values('1','��ҽ��','ZYK',0,'��ָ������ҽѧ��������ϲ��������Ʒ������ٴ����š��ڱ�ϵͳ�о߱���ҽ�����Կ��ҵĲ��ˣ�������д��ҽ����ϣ�¼����ҽ�������ݡ�');


--���ݵ���
Update ϵͳ������ Set ������='����ҩ�������Ϻ���ҩ',����˵��='��ʾ����ҩƷҽ����Ҫ�����ϣ�Ȼ����ȥ��ҩ���˷�' Where ������=68;
Delete From ϵͳ������ Where ������=25 And ������<>'����ǩ����֤����';
Delete From ϵͳ������ Where ������=26 And ������<>'����ǩ��ʹ�ó���';
Insert Into ϵͳ������(������,������,����ֵ,ȱʡֵ,����˵��) 
	Select 25,'����ǩ����֤����','0','0','����ǩ����֤���ĵı��,��1��ʼ˳����,0��ʾ��ʹ�õ���ǩ��;��1Ϊ����CA��֤����' From Dual 
	Where Not Exists(Select 1 From ϵͳ������ Where ������=25);
Insert Into ϵͳ������(������,������,����ֵ,ȱʡֵ,����˵��) 
	Select 26,'����ǩ��ʹ�ó���','000','000','�Բ�ͬ�����Ƿ�ʹ�õ���ǩ�����п���,����λ���ֱ�Ϊ:����,סԺ,ҽ��,0-������,1-����' From Dual
	Where Not Exists(Select 1 From ϵͳ������ Where ������=26);

--���˺����
Insert Into ϵͳ������(������,������,����ֵ,ȱʡֵ,����˵��)
Select 120,'���ĸ���������㷽ʽ','0','0','�������ĸ�������,ȷ���ɱ��۵ļ��㷽ʽ��0-��ָ������ʼ���,1-����������(ȡ����Ŀ¼�еĳɱ��ۣ�'  From dual;

Insert Into ϵͳ������(������,������,����ֵ,ȱʡֵ,����˵��)
Select 121,'���ķֶμӳ���','0','0','����ʱ�����������ʱ,���ֶμӳ��ʼ���.'  From dual;

Insert Into ϵͳ������(������,������,����ֵ,ȱʡֵ,����˵��)
Select 123,'���ϸ����ָ���۸�','0','0','���ϸ��������ָ�����ۺ�ָ���ۼ�' From dual;



Insert Into ҽ�����ݶ���(�������,ҽ������) 
Select '5','[������]+iif([����]<>"","("+[����]+")","")+iif([���]<>""," "+[���],"")' From Dual Union All
Select '6','[������]+iif([����]<>"","("+[����]+")","")+iif([���]<>""," "+[���],"")' From Dual Union All
Select '8','"��ҩ"+[����]+"��,"+[����Ƶ��]+","+[�巨]+","+[�÷�]+":"+[�䷽���]' From Dual Union All
Select 'C','[������Ŀ]+iif([����걾]<>"","("+[����걾]+")","")' From Dual Union All
Select 'D','[�����Ŀ]+iif([��鲿λ]<>"","("+[��鲿λ]+")","")' From Dual Union All
Select 'E','[������Ŀ]' From Dual Union All
Select 'F','Format([��ʼʱ��],"MM��dd��HH:mm")+iif([������]<>""," �� "+[������]+" ���� "," �� ")+[��Ҫ����]+iif([��������]<>""," �� "+[��������],"")' From Dual Union All
Select 'H','[������Ŀ]' From Dual Union All
Select 'I','[������Ŀ]' From Dual Union All
Select 'K','[������Ŀ]' From Dual Union All
Select 'L','[������Ŀ]' From Dual Union All
Select 'M','[������Ŀ]' From Dual Union All
Select 'Z','[������Ŀ]' From Dual;

----������Ϊִ�н���,�ɲ�ִ��,��Ӱ���¿�ҽ����ǩ��
--Update ����ҽ����¼ A Set ����=(Select ���� From ������Ϣ B Where B.����ID=A.����ID);
--Update ����ҽ����¼ A Set ����=(Select ���� From ������Ϣ B Where B.����ID=A.����ID);
--Update ����ҽ����¼ A Set �Ա�=(Select �Ա� From ������Ϣ B Where B.����ID=A.����ID);
----������Ϊִ�н���,�ɲ�ִ��,��Ӱ���¿�ҽ����ǩ��
--Update H����ҽ����¼ A Set ����=(Select ���� From ������Ϣ B Where B.����ID=A.����ID);
--Update H����ҽ����¼ A Set ����=(Select ���� From ������Ϣ B Where B.����ID=A.����ID);
--Update H����ҽ����¼ A Set �Ա�=(Select �Ա� From ������Ϣ B Where B.����ID=A.����ID);

-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------
--7404
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1342,'����',user,'ҩƷ�����޶�','SELECT');
Delete From zlProgPrivs Where ϵͳ=100 And ���=1342 And ����='����' And ����='HҩƷ�շ���¼';
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1342,'����',user,'HҩƷ�շ���¼','SELECT');

--6888
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1023,'���пⷿ','��������пⷿҩƷ���д���ⷿ����');

--������Ϣ������
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1101,'����',user,'���˵�����¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1101,'������Ϣ',user,'zl_������Ϣ_UPDATE','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1101,'������Ϣ',user,'ְҵ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1101,'������Ϣ',user,'ҽ�Ƹ��ʽ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1101,'������Ϣ',user,'ѧ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1101,'������Ϣ',user,'�Ա�','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1101,'������Ϣ',user,'������Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1101,'������Ϣ',user,'�շ�ϸĿ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1101,'������Ϣ',user,'�շ��ض���Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1101,'������Ϣ',user,'�շѼ�Ŀ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1101,'������Ϣ',user,'���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1101,'������Ϣ',user,'����ϵ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1101,'������Ϣ',user,'����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1101,'������Ϣ',user,'���㷽ʽӦ��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1101,'������Ϣ',user,'���㷽ʽ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1101,'������Ϣ',user,'����״��','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1101,'������Ϣ',user,'��Լ��λ','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1101,'������Ϣ',user,'������Ʊ�','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1101,'������Ϣ',user,'����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1101,'������Ϣ',user,'����','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1101,'������Ϣ',user,'�ѱ�','SELECT');

--�ٴ�ҽ��������ҽ��վ��Ƭ
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1203,'��Ƭ����','�����Ƭ����վ������ص�Ӱ��');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1203,'����','Ӱ����ͼ��',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1203,'����','Ӱ��������',USER,'SELECT');

Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1204,'��Ƭ����','�����Ƭ����վ������ص�Ӱ��');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1204,'����','Ӱ����ͼ��',USER,'SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,����,������,Ȩ��) Values(100,1204,'����','Ӱ��������',USER,'SELECT');

--6986
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1309,'ҽ����ѯ','����鿴����ҽ����');
--6884
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1205,'�����ջص���','���ڷ����ջ�ҽ��ʱ�����ջ����Ĳ���Ȩ�ޡ�');

--6896,7074:��ʿȷ��ֹͣ,��ͣȨ�޶���
Delete From zlProgPrivs Where ϵͳ=100 And ���=1205 And ����='ҽ��ֹͣ' And Upper(����)='ZL_����ҽ����¼_ȷ��ֹͣ';
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1205,'ҽ��ȷ��ֹͣ','��ҽ����ֹͣ��ҽ������ȷ��ֹͣ�Ĳ���Ȩ�ޡ�');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1205,'ҽ��ȷ��ֹͣ',user,'ZL_����ҽ����¼_ȷ��ֹͣ','EXECUTE');

Delete From zlprogPrivs Where ϵͳ=100 And ���=1205 And ����='ҽ��ֹͣ' And Upper(����) IN('ZL_����ҽ����¼_��ͣ','ZL_����ҽ����¼_����');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1205,'ҽ����ͣ','���´��ҽ��������ͣ,���õĲ���Ȩ�ޡ�');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1205,'ҽ����ͣ',user,'ZL_����ҽ����¼_��ͣ','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1205,'ҽ����ͣ',user,'ZL_����ҽ����¼_����','EXECUTE');

--7080:Ƥ�Խ��Ȩ�޿���
Delete From zlProgPrivs Where ϵͳ=100 And ���=1203 And ����='ҽ���´�' And Upper(����)='ZL_����ҽ����¼_Ƥ��';
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1203,'Ƥ��ҽ�����','��Ƥ�Ե�ҽ������༭�Ĳ���Ȩ�ޡ�');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1203,'Ƥ��ҽ�����',user,'ZL_����ҽ����¼_Ƥ��','EXECUTE');
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1206,'��������','���ò����Ĳ���Ȩ�ޡ�');

--���� 2005-12-16
--Ȩ����������
--����/סԺ�����֤������ǼǼ�¼��zl_����ǼǼ�¼_UPDATE��zl_����ǼǼ�¼_DELETE
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1121,'����',user,'����ǼǼ�¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1121,'����',user,'zl_����ǼǼ�¼_UPDATE','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1121,'����',user,'zl_����ǼǼ�¼_DELETE','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1131,'����',user,'����ǼǼ�¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1131,'����',user,'zl_����ǼǼ�¼_UPDATE','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1131,'����',user,'zl_����ǼǼ�¼_DELETE','EXECUTE');
--������㣺zl_����ǼǼ�¼_����
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1121,'����',user,'zl_����ǼǼ�¼_����','EXECUTE');
--��Ժ�Ǽ�/��Ժ�ǼǼ�����������ǼǼ�¼��zl_����ǼǼ�¼_UPDATE��zl_����ǼǼ�¼_DELETE��zl_����ǼǼ�¼_����״̬
--�˴�Ȩ�޲��������֤��Ȩ��
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1131,'����',user,'zl_����ǼǼ�¼_����״̬','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1132,'����',user,'����ǼǼ�¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1132,'����',user,'zl_����ǼǼ�¼_UPDATE','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1132,'����',user,'zl_����ǼǼ�¼_DELETE','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1132,'����',user,'zl_����ǼǼ�¼_����״̬','EXECUTE');
--סԺ�������/���˷��ò�ѯ������ǼǼ�¼
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1137,'����',user,'����ǼǼ�¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1139,'����',user,'����ǼǼ�¼','SELECT');


--����ǩ��Ȩ�޵���
----ҽ�����ݶ���Ȩ��
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1203,'����',user,'ҽ�����ݶ���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1204,'����',user,'ҽ�����ݶ���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1205,'����',user,'ҽ�����ݶ���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1206,'����',user,'ҽ�����ݶ���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1207,'����',user,'ҽ�����ݶ���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1804,'����',user,'ҽ�����ݶ���','SELECT');
----������������
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1011,'����',user,'������Ŀ���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1011,'����',user,'ҽ�����ݶ���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1011,'����',user,'zl_ҽ�����ݶ���_Delete','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1011,'����',user,'zl_ҽ�����ݶ���_Insert','EXECUTE');
----��Ա����
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1002,'����֤��ע��','����Ա������֤�����ע���Ȩ��(���Ż����ע��)��');
--ҽ��ǩ��Ȩ����д��������
Delete From zlProgPrivs Where ϵͳ=100 And ���=1002 And ����='����' And ���� IN('ϵͳ����','ϵͳ������');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1002,'����',user,'ϵͳ������','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1002,'����',user,'��Ա֤���¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1002,'����֤��ע��',user,'zl_��Ա֤���¼_Insert','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1002,'����֤��ע��',user,'zl_��Ա֤���¼_Delete','EXECUTE');
----����ҺŹ���
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1111,'��ҽ�����˺�','����������ҽ���Ĳ��������˺š�');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1111,'�Һ�',user,'ZL_���￨��¼_INSERT','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1111,'�˺�',user,'ZL_���￨��¼_DELETE','EXECUTE');
----�����������
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1132,'����',user,'������ҳ�ӱ�','SELECT');
----ҽ�����Ҽ���
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1135,'���Ƹ�������','���Ƶ��ݸ������ʵĲ���Ȩ�ޡ�');
----����ҽ������վ
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1203,'����',user,'��Ա֤���¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1203,'����',user,'ҽ��ǩ����¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1203,'����',user,'Hҽ��ǩ����¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1203,'ҽ���´�',user,'ҽ��ǩ����¼_ID','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1203,'ҽ���´�',user,'zl_ҽ��ǩ����¼_Insert','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1203,'ҽ���´�',user,'zl_ҽ��ǩ����¼_Delete','EXECUTE');
----סԺҽ������վ
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1204,'����',user,'��Ա֤���¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1204,'����',user,'ҽ��ǩ����¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1204,'����',user,'Hҽ��ǩ����¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1204,'ҽ���´�',user,'ҽ��ǩ����¼_ID','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1204,'ҽ���´�',user,'zl_ҽ��ǩ����¼_Insert','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1204,'ҽ���´�',user,'zl_ҽ��ǩ����¼_Delete','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1204,'ҽ���´�',user,'zl_����ҽ����¼_����','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1204,'ҽ���´�',user,'zl_����ҽ����¼_��������','EXECUTE');
--סԺ��ʿ����վ:7031
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1205,'ҽ���´�',user,'ZL_����ҽ����¼_У��','EXECUTE');
----ҽ������վ
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1206,'����',user,'��Ա֤���¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1206,'����',user,'ҽ��ǩ����¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1206,'����',user,'Hҽ��ǩ����¼','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1206,'�����´�ҽ��',user,'ҽ��ǩ����¼_ID','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1206,'�����´�ҽ��',user,'zl_ҽ��ǩ����¼_Insert','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1206,'�����´�ҽ��',user,'zl_ҽ��ǩ����¼_Delete','EXECUTE');

----����Ȩ������
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1205,'ҽ���´�',user,'ZL_����ҽ����¼_�������','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1206,'�����´�ҽ��',user,'ZL_����ҽ����¼_�������','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values (100,1804,'����ҽ��',user,'ZL_����ҽ����¼_�������','EXECUTE');

--������ѯϵͳ
insert into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) values (100,1536,'����',user,'����ģ�����','SELECT');

--����Ŀ¼
Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(100,1711,'�ֶμӳ���','�����۷ֶ����üӳ��ʡ�');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1711,'����',user,'���ϼӳɷ���','SELECT');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1711,'�ֶμӳ���',user,'ZL_���ϼӳɷ���_DELETE','EXECUTE');
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1711,'�ֶμӳ���',user,'ZL_���ϼӳɷ���_INSERT','EXECUTE');
--�����⹺���
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1712,'����',user,'���ϼӳɷ���','SELECT');
--�����������
Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��) Values(100,1714,'����',user,'���ϼӳɷ���','SELECT');


-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--PACS
CREATE OR REPLACE Procedure ZL_Ӱ����_BEGIN(
  ִ�м�_IN   ����ҽ������.ִ�м�%Type,
  ����_IN   Ӱ�����¼.����%Type,
	ҽ��ID_IN		Ӱ�����¼.ҽ��ID%Type,
	���ͺ�_IN		Ӱ�����¼.���ͺ�%Type,
  Ӱ�����_IN Ӱ�����¼.Ӱ�����%Type,
  ����_IN     Ӱ�����¼.����%Type,
  Ӣ����_IN   Ӱ�����¼.Ӣ����%Type,
  �Ա�_IN     Ӱ�����¼.�Ա�%Type,
  ����_IN     Ӱ�����¼.����%Type,
  ��������_IN Ӱ�����¼.��������%Type,
  ���_IN     Ӱ�����¼.���%Type,
  ����_IN     Ӱ�����¼.����%Type,
  ������_IN Ӱ�����¼.������%Type,
  ���Ž�Ƭ_IN Ӱ�����¼.���Ž�Ƭ%Type,
  ����豸_IN Ӱ�����¼.����豸%Type,
  �޸�_IN     Number:=0,
  �绰_IN     Ӱ�����¼.��ϵ�绰%Type:=Null
) IS
  Cursor c_Advice IS
    Select A.ҽ��ID
    From ����ҽ������ A,����ҽ����¼ B
    Where (B.ID=ҽ��ID_IN Or (B.���ID=ҽ��ID_IN And B.������� IN('F','G','D')))
      And A.ҽ��ID=B.ID And A.���ͺ�+0=���ͺ�_IN;
  iRecCount Number;
Begin
	Select Count(*) Into iRecCount From Ӱ�����¼ Where ����=����_IN And Ӱ�����=Ӱ�����_IN;

	Update Ӱ�����¼
		Set Ӱ�����=Ӱ�����_IN,
   			����=����_IN,
				����=����_IN,
				Ӣ����=Ӣ����_IN,
				�Ա�=�Ա�_IN,
				����=����_IN,
				��������=��������_IN,
				���=���_IN,
				����=����_IN,
				������=������_IN,
				���Ž�Ƭ=���Ž�Ƭ_IN,
				����豸=����豸_IN,
				��ϵ�绰=�绰_IN
	Where ҽ��ID=ҽ��ID_IN And ���ͺ�=���ͺ�_IN;
	If SQl%RowCount=0 Then
    Insert Into Ӱ�����¼(ҽ��ID,���ͺ�,Ӱ�����,����,����,Ӣ����,�Ա�,����,��������,
      ���,����,������,���Ž�Ƭ,����豸,��ϵ�绰)
    Values(ҽ��ID_IN,���ͺ�_IN,Ӱ�����_IN,����_IN,����_IN,Ӣ����_IN,�Ա�_IN,����_IN,��������_IN,
      ���_IN,����_IN,������_IN,���Ž�Ƭ_IN,����豸_IN,�绰_IN);
  End if;
  
  If �޸�_IN=0 Then
    For r_Advice In c_Advice Loop
      Update ����ҽ������ Set �״�ʱ��=Sysdate,ĩ��ʱ��=Sysdate,ִ��״̬=3,ִ�й���=2,ִ�м�=ִ�м�_IN Where ҽ��ID=r_Advice.ҽ��ID And ���ͺ�=���ͺ�_IN;
    End Loop;
    If iRecCount=0 Then
      Update Ӱ������� Set ������=����_IN Where ����=Ӱ�����_IN;
    End If;
  Else
    For r_Advice In c_Advice Loop
      Update ����ҽ������ Set ִ�м�=ִ�м�_IN Where ҽ��ID=r_Advice.ҽ��ID And ���ͺ�=���ͺ�_IN;
    End Loop;
  End If;
Exception
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_Ӱ����_BEGIN;
/
--LIS
CREATE OR REPLACE Procedure ZL_����걾��¼_�������(
	ID_IN			����걾��¼.ID%Type
) IS

	--δ��˵ķ�����(������ҩƷ)
	Cursor c_Verify(v_ҽ��id In Number) is	Select Distinct ��¼����,NO,��� From ���˷��ü�¼
						Where �շ���� Not IN('5','6','7') And ҽ�����+0=v_ҽ��id
							And ���ʷ���=1 And ��¼״̬=0 And �۸񸸺� IS NULL
							And (��¼����,NO) IN(
								Select ��¼����,NO From ����ҽ������ Where ҽ��ID=v_ҽ��id
								Union ALL
								Select ��¼����,NO From ����ҽ������ Where ҽ��ID In (Select ID From ����ҽ����¼ Where v_ҽ��id In (ID,���id)))
						Order By ��¼����,NO,���;

	--��ȡ���˵������Ϣ
	CURSOR c_Advice2(v_ҽ��id In Number) IS	SELECT * FROM ����ҽ����¼ WHERE ID=v_ҽ��id;

	r_Advice2 c_Advice2%RowType;

	--�����ļ������Ԫ��
	CURSOR c_File(v_File number) IS	SELECT ����,����,�ı�ת��,�����ı�,������ʾ,��������,����λ��,��������,����λ��,Ƕ�뷽ʽ
					FROM �����ļ���� A,����Ԫ��Ŀ¼ B
					where A.����Ԫ��id=B.ID
					      AND A.�����ļ�id=v_File
					order by A.�������;

	--������Ŀ���
	CURSOR c_Result(v_ҽ��id In Number) IS	Select	B.������,B.ID,B.����,B.��λ,A.������,DECODE(A.�����־,1,'1-����',2,'2-ƫ��',3,'3-ƫ��',4,'4-����',5,'5-����','') AS �����־,A.����ο�,C.ID As �걾id
						From	������ͨ��� A,
							����������Ŀ B,
							����걾��¼ C,
							������Ŀ�ֲ� D
						Where	A.������Ŀid=B.ID
							AND A.����걾id=C.ID
							AND A.��¼����=C.������
							AND C.����״̬=2
							AND D.�걾id=A.����걾id
							AND D.��Ŀid=A.������Ŀid
							AND D.ҽ��id=v_ҽ��id;

	--���ҵ�ǰ�걾���������
	CURSOR c_SampleQuest(v_΢���� In Number) IS	Select Distinct ҽ��id From (
							Select ҽ��id From ������Ŀ�ֲ� Where 0=v_΢���� And �걾id=ID_IN
							Union
							Select ҽ��id From ����걾��¼ Where 1=v_΢���� And ID=ID_IN);

	Cursor c_Stuff(v_NO Varchar2,v_��ҳid Number) is
		Select NO,����,�ⷿID From δ��ҩƷ��¼
		Where NO=v_NO And ����=25 And �ⷿID IS Not Null
			And Not Exists(Select ����ֵ From ϵͳ������ Where ������=Decode(v_��ҳid,NULL,92,63) And ����ֵ='1')
			And Exists(
				Select A.��� From ���˷��ü�¼ A,�������� B
				Where A.��¼����=2 And A.��¼״̬=1 And A.NO=v_NO
					And A.�շ�ϸĿID=B.����ID And B.��������=1
				)
		Order BY �ⷿID;

	v_ִ��			Number(1);
	v_NO			����ҽ������.NO%Type;
	v_����			����ҽ������.��¼����%Type;
	v_���			Varchar2(1000);

	v_Temp			Varchar2(255);
	v_��Ա����ID		������Ա.����ID%Type;
	 v_��Ա���		��Ա��.���%Type;
	v_��Ա����		��Ա��.����%Type;
	v_Count			Number(18);


	v_��������id		number(18);
	v_����id		number(18);
	v_����id		number(18);
	v_�ļ�ID		number(18);
	v_Index			number(18);
	v_��������		�����ļ�Ŀ¼.����%Type;
	v_�������� 		�����ļ�Ŀ¼.����%Type;

	v_FLAG			Number(1);
	v_MaxIndex		Number(18);
	v_΢����걾		Number(1):=0;
	v_��ҳid	Number(18);
Begin

	--0.����Ա��Ϣ:����ID,��������;��ԱID,��Ա���,��Ա����
	v_Temp:=zl_Identity;
	v_��Ա����ID:=To_Number(Substr(v_Temp,1,Instr(v_Temp,',')-1));
	v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
	v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
	v_��Ա���:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
	v_��Ա����:=Substr(v_Temp,Instr(v_Temp,',')+1);

	--ִ�к��Զ���˶�Ӧ�ļ��ʻ��۵�(������ҩƷ)
	v_ִ��:=0;
	Begin
		Select To_Number(Nvl(����ֵ,'0')) Into v_ִ�� From ϵͳ������ Where ������=81;
	Exception
		When Others Then v_ִ��:=0;
	End;

	v_΢����걾:=0;
	Begin
		Select 1 Into v_΢����걾 From ����걾��¼ Where ΢����걾=1 And ID=ID_IN;
	Exception
		When Others Then v_΢����걾:=0;
	End;

	--1.�ñ��걾��״̬������˺�ʱ��
	UPDATE ����걾��¼ SET �����=v_��Ա����,���ʱ��=SYSDATE,����״̬=2 WHERE ID=ID_IN;

	--2.��鵱ǰ�걾��ص��������ر걾�Ƿ�������
	For r_SampleQuest In c_SampleQuest(v_΢����걾) Loop

		v_Count:=0;

		If v_΢����걾=0 Then
			Begin
				Select NVL(COUNT(1),0) INTO v_Count  From ����걾��¼ Where ����״̬<2 And ID In (Select �걾id From ������Ŀ�ֲ� Where ҽ��id=r_SampleQuest.ҽ��id);
			Exception
				When Others Then v_Count:=0;
			End;
		End If;

		--r_SampleQuest.ҽ��id�����Ѿ����,�����������
		If v_Count=0 Then

			--1.�����뵥��ִ��״̬
			Update ����ҽ������ Set ִ��״̬=1 Where ҽ��id In (Select ID From ����ҽ����¼ Where r_SampleQuest.ҽ��id In (ID,���id));

			--2.����ִ�д���
			Update ���˷��ü�¼ Set ִ��״̬=1,ִ��ʱ��=Sysdate,ִ����=v_��Ա����
			Where �շ���� Not IN('5','6','7')
				And (ҽ�����,��¼����,NO) IN (
					Select ҽ��ID,��¼����,NO From ����ҽ������ Where ҽ��ID=r_SampleQuest.ҽ��id
					Union ALL
					Select ҽ��ID,��¼����,NO From ����ҽ������ Where ҽ��ID In (Select ID From ����ҽ����¼ Where r_SampleQuest.ҽ��id In (ID,���ID)));

			--3.�Զ���˼���
			If v_ִ��=1 Then
				For r_Verify IN c_Verify(r_SampleQuest.ҽ��id) Loop
					IF 	r_Verify.NO||','||r_Verify.��¼����<>v_NO||','||v_���� Then
						If v_��� IS Not NULL Then
							If v_����=1 Then
								zl_������ʼ�¼_Verify(v_NO,v_��Ա���,v_��Ա����,Substr(v_���,2));
							Elsif v_����=2 Then
								zl_סԺ���ʼ�¼_Verify(v_NO,v_��Ա���,v_��Ա����,Substr(v_���,2));
							End IF;
						End IF;
						v_���:=NULL;
					End IF;
					v_NO:=r_Verify.NO;
					v_����:=r_Verify.��¼����;
					v_���:=v_���||','||r_Verify.���;
				End Loop;
				If v_��� IS Not NULL Then
					If v_����=1 Then
						zl_������ʼ�¼_Verify(v_NO,v_��Ա���,v_��Ա����,Substr(v_���,2));
					Elsif v_����=2 Then
						zl_סԺ���ʼ�¼_Verify(v_NO,v_��Ա���,v_��Ա����,Substr(v_���,2));
					End IF;
				End IF;
			End IF;

			--����Լ����ĵ�
			v_No:=NextNo(14);

			Update �����Լ���¼ Set No=v_No Where ҽ��id=r_SampleQuest.ҽ��id;

			If v_No Is Not Null Then

				ZL_�����Լ���¼_BILL(r_SampleQuest.ҽ��id,v_No);

				v_��ҳid:=Null;
				Select ��ҳid Into v_��ҳid From ����ҽ����¼ A WHERE ID=r_SampleQuest.ҽ��id;

				If v_��ҳid Is Null Then
					zl_������ʼ�¼_Verify(v_No,v_��Ա���,v_��Ա����);
				Else
					zl_סԺ���ʼ�¼_Verify(v_No,v_��Ա���,v_��Ա����);
				End If;

				--�������û���Զ�����,���Զ�����,���򲻴���
				For r_Stuff In c_Stuff(v_No,v_��ҳid) Loop
					zl_�����շ���¼_��������(r_Stuff.�ⷿID,25,v_No,v_��Ա����,v_��Ա����,v_��Ա����,1,Sysdate);
				End Loop;

			End If;

			--4.�Զ���д����,������ͨ��Ŀ,΢����
			If v_΢����걾>=0 Then
				v_����id:=0;
				v_����id:=0;
				begin
					Select Distinct nvl(����id,0) into v_����id from ����ҽ������  Where ҽ��id in (SELECT ID FROM ����ҽ����¼ WHERE r_SampleQuest.ҽ��id In (ID,���id));
				exception
					when others then v_����id:=0;
				end;

				If Nvl(v_����id,0)=0 then

					--�������˲�����¼
					Open c_Advice2(r_SampleQuest.ҽ��id);
					Fetch c_Advice2 Into r_Advice2;

					--���Ҫ��д�ı����ʽ���Ƿ��м���ר��ֽ,���û���򷵻�
					v_�ļ�ID:=0;
					begin
						Select U.ID,U.����,U.����
						Into v_�ļ�ID,v_��������,v_��������
						From �����ļ���� X,����Ԫ��Ŀ¼ Y,�����ļ�Ŀ¼ U
						Where X.�����ļ�id in (select A.�����ļ�id
									from ���Ƶ���Ӧ�� A,����ҽ����¼ B
									where A.������Ŀid=B.������Ŀid
										and B.���ID=r_SampleQuest.ҽ��id
										and A.Ӧ�ó���=r_Advice2.������Դ)
							AND X.����Ԫ��id=Y.ID
							AND U.ID=X.�����ļ�id
							AND Y.����=4 and Y.����='000009';
					exception
						when others then v_�ļ�ID:=0;
					end;

					--��LISר��ֽ,Ҫ��д
					If nvl(v_�ļ�ID,0)>0 then

						--�²�������id
						v_����id:=0;
						Select ���˲�����¼_ID.Nextval Into v_����id From Dual;

						ZL_���˲���_INSERT(v_����id,r_Advice2.����id,r_Advice2.��ҳid,r_Advice2.�Һŵ�,r_Advice2.Ӥ��,r_Advice2.���˿���id,v_��������,v_�ļ�ID,v_��������,v_��Ա����,r_SampleQuest.ҽ��id);

						--������������β����������Ԫ�ؼ�¼
						v_Index:=0;
						FOR r_File In c_File(v_�ļ�ID) LOOP
							v_Index:=v_Index+1;

							Select ���˲�������_ID.Nextval Into v_��������id From Dual;

							if r_File.����=4 and r_File.����='000009' then
								v_����id:=v_��������id;
							end if;

							ZL_���˲�������_INSERT(v_��������id,NULL,v_����id,v_Index,r_File.����,r_File.����,r_File.�ı�ת��,r_File.�����ı�,r_File.������ʾ,r_File.��������,r_File.����λ��,0,r_File.��������,r_File.����λ��,0,r_File.Ƕ�뷽ʽ);
						END LOOP;

					End if;
					Close c_Advice2;
				Else

					--���Ҫ����д�ı����ʽ���Ƿ��м���ר��ֽ,���ҳ�����ר��ֽ�ڱ����е�λ��,���û���򷵻�
					v_����id:=0;
					begin
						select nvl(id,0) into v_����id from ���˲������� where Ԫ������=4 and Ԫ�ر���='000009' and ������¼id=v_����id;
					exception
						when others then v_����id:=0;
					End;
				End If;

				--��Lisר��ֽ����,����д����������ר��ֽ��
				If v_����id>0 And v_����id>0 Then

					v_Index:=0;
          Delete from ���˲��������� where ������id+0 In
            (Select	D.��Ŀid From	����걾��¼ C,������Ŀ�ֲ� D
              Where	D.�걾id=C.ID AND C.����״̬=2 AND D.ҽ��id=r_SampleQuest.ҽ��id)
            And ����id In (Select ID From ���˲������� Where Ԫ������=4 and Ԫ�ر���='000009' and ������¼id=v_����id);
					FOR r_Result In c_Result(r_SampleQuest.ҽ��id) LOOP

						--Delete from ���˲��������� where ������id=r_Result.ID and ����id In (Select ID From ���˲������� Where Ԫ������=4 and Ԫ�ر���='000009' and ������¼id=v_����id);

						v_MaxIndex:=1;
						Begin
							Select Nvl(Max(A.�ؼ���),0)+1 Into v_MaxIndex From ���˲��������� A,���˲������� B Where A.����id=B.ID AND B.Ԫ������=4 and B.Ԫ�ر���='000009' and B.������¼id=v_����id;
						Exception
							When Others Then v_MaxIndex:=1;
						End;

						Insert Into ���˲���������(����ID,�ؼ���,�ؼ���,����,������ID,��ֵ����,������λ,��������)
						Select ID,v_MaxIndex,2,r_Result.������,r_Result.ID,r_Result.����,r_Result.��λ,r_Result.������||''''||r_Result.�����־||''''||r_Result.����ο�
						From ���˲������� Where Ԫ������=4 and Ԫ�ر���='000009' and ������¼id=v_����id;

					END LOOP;
				End If;
				--�޸Ĳ���ҽ�����ͼ�¼�ı���id��
				IF v_����id>0 THEN
					Update ����ҽ������ SET ����id=v_����id WHERE ҽ��id in (SELECT ID FROM ����ҽ����¼ WHERE r_SampleQuest.ҽ��id In (ID,���id));
				END IF;
			End If;
		End If;
	End Loop;
Exception
	When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_����걾��¼_�������;
/

-------------------------------------------------------------------------
--����˵���� �Զ�����һ�����˵ķ���
--��ڲ�����
--       PatiID  number    �������ID
--       PageID  number    ������ҳID������������ͬȷ����Ҫ����Ĳ���
--       ReCalcBDate  Date ���㿪ʼʱ��
--���ù�ϵ�� �ⲿӦ�ó�����ñ�����
-------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE Zl1_autocptpati (
    PatiID IN NUMBER,
    PageID IN NUMBER,
    ReCalcBDate IN ���˱䶯��¼.�ϴμ���ʱ��%Type:=Null
)
AS
    Modilast NUMBER (1);--�Ƿ����������Զ��ƷѲ���
    Period VARCHAR2 (6);--��Ҫ�������С�ڼ�
BEGIN
    SELECT �ڼ�
      INTO Period
      FROM �ڼ��
     WHERE TRUNC (SYSDATE) BETWEEN TRUNC (��ʼ����) AND TRUNC (��ֹ����);
    SELECT ����ֵ
      INTO Modilast
      FROM ϵͳ������
     WHERE ������ = 7;

    IF Modilast = 1 THEN
        Period :=
          TO_CHAR (ADD_MONTHS (TO_DATE (Period || '05', 'yyyymmdd'), -1),
              'yyyymm');
    END IF;
    
    IF ReCalcBDate IS NOT NULL THEN 
        Update ���˱䶯��¼ Set �ϴμ���ʱ��=Null Where ����ID=PatiID And ��ҳID=PageID And �ϴμ���ʱ��>=ReCalcBDate;
        COMMIT;
    END IF; 

    Zl1_autocptone (PatiID, PageID, Period);
    COMMIT;
END;
/
-------------------------------------------------------------------------
--����˵���� �Զ�����һ���������в��˷���
--��ڲ����� WardID    number  ����ID,ReCalcBDate  Date ���㿪ʼʱ��
--���ù�ϵ�� �ⲿӦ�ó�����ñ�����
-------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE Zl1_autocptward (
    WardID IN NUMBER,
    ReCalcBDate IN ���˱䶯��¼.�ϴμ���ʱ��%Type:=Null
)
AS
    Modilast NUMBER (1);--�Ƿ����������Զ��ƷѲ���
    Period VARCHAR2 (6);--��Ҫ�������С�ڼ�

    CURSOR Patitab
    IS
        SELECT Distinct ����ID, ��ҳID
          FROM �����Զ�����
         WHERE ����ID = WardID
            AND TRUNC (��ֹ����) >=
                                (SELECT MIN (��ʼ����)
                                    FROM �ڼ��
                                  WHERE �ڼ� >= Period);
BEGIN
    SELECT �ڼ�
      INTO Period
      FROM �ڼ��
     WHERE TRUNC (SYSDATE) - 1 BETWEEN TRUNC (��ʼ����) AND TRUNC (��ֹ����);
    SELECT ����ֵ
      INTO Modilast
      FROM ϵͳ������
     WHERE ������ = 7;

    IF Modilast = 1 THEN
        Period :=
          TO_CHAR (ADD_MONTHS (TO_DATE (Period || '05', 'yyyymmdd'), -1),
              'yyyymm');
    END IF;
    
    IF ReCalcBDate IS NOT NULL THEN 
        Update ���˱䶯��¼ Set �ϴμ���ʱ��=Null 
        Where (����ID,��ҳID) IN (Select ����ID,��ҳID From ������ҳ Where ��ǰ����ID=WardID And ��Ժ���� Is Null)  
                    And �ϴμ���ʱ��>=ReCalcBDate;
        COMMIT;
    END IF; 

    FOR Patifld IN Patitab LOOP
        IF      Patifld.����ID IS NOT NULL
            AND Patifld.��ҳID IS NOT NULL THEN
            Zl1_autocptone (Patifld.����ID, Patifld.��ҳID, Period);
            COMMIT;
        END IF;
    END LOOP;
END;
/

--���� 2005-12-16
--������������Ժ
CREATE OR REPLACE PROCEDURE zl_����ǼǼ�¼_UPDATE(
	����_IN				����ǼǼ�¼.����%TYPE,
	����ID_IN			����ǼǼ�¼.����ID%TYPE,
	��ҳID_IN			����ǼǼ�¼.��ҳID%TYPE,
	����ʱ��_IN			����ǼǼ�¼.����ʱ��%TYPE,
	״̬_IN				����ǼǼ�¼.״̬%TYPE:=0,
	ҽ�����_IN			����ǼǼ�¼.ҽ�����%TYPE:=NULL,
	�ʻ����_IN			����ǼǼ�¼.�ʻ����%TYPE:=0,
	����ID_IN			����ǼǼ�¼.����ID%TYPE:=NULL,
	��������_IN			����ǼǼ�¼.��������%TYPE:=NULL,
	����֢_IN			����ǼǼ�¼.����֢%TYPE:=NULL,
	IC����Ϣ_IN			����ǼǼ�¼.IC����Ϣ%TYPE:=NULL,
	HIS��ˮ��_IN		����ǼǼ�¼.HIS��ˮ��%TYPE:=NULL,
	YB��ˮ��_IN			����ǼǼ�¼.YB��ˮ��%TYPE:=NULL,
	��ע_IN				����ǼǼ�¼.��ע%TYPE:=NULL
)
AS 
BEGIN 
	UPDATE ����ǼǼ�¼
	SET ״̬=״̬_IN,
		ҽ�����=ҽ�����_IN,
		�ʻ����=�ʻ����_IN,
		����ID=����ID_IN,
		��������=��������_IN,
		����֢=����֢_IN,
		IC����Ϣ=IC����Ϣ_IN,
		HIS��ˮ��=HIS��ˮ��_IN,
		YB��ˮ��=YB��ˮ��_IN,
		��ע=��ע_IN
	WHERE ����=����_IN AND ����ID=����ID_IN 
	AND Nvl(��ҳID,0)=Nvl(��ҳID_IN,0) AND ����ʱ��=����ʱ��_IN;

	IF SQL%ROWCOUNT =0 THEN 
		INSERT INTO ����ǼǼ�¼
		(����,����ID,��ҳID,����ʱ��,״̬,ҽ�����,�ʻ����,
		����ID,��������,����֢,IC����Ϣ,HIS��ˮ��,YB��ˮ��,��ע)
		VALUES 
		(����_IN,����ID_IN,Nvl(��ҳID_IN,0),����ʱ��_IN,״̬_IN,ҽ�����_IN,�ʻ����_IN,
		����ID_IN,��������_IN,����֢_IN,IC����Ϣ_IN,HIS��ˮ��_IN,YB��ˮ��_IN,��ע_IN);
	END IF ;
EXCEPTION
    WHEN OTHERS THEN zl_ErrorCenter(SQLCODE,SQLERRM);
END zl_����ǼǼ�¼_UPDATE;
/


--����������ȡ������Ǽǡ�������Ժ���޷ѳ�Ժʱ��״̬����Ϊ��
CREATE OR REPLACE PROCEDURE zl_����ǼǼ�¼_DELETE(
	����_IN				����ǼǼ�¼.����%TYPE,
	����ID_IN			����ǼǼ�¼.����ID%TYPE,
	��ҳID_IN			����ǼǼ�¼.��ҳID%TYPE,
	����ʱ��_IN			����ǼǼ�¼.����ʱ��%TYPE
)
AS 
BEGIN 
	DELETE  ����ǼǼ�¼
	WHERE ����=����_IN AND ����ID=����ID_IN 
	AND NVL(��ҳID,0)=NVL(��ҳID_IN,0) AND ����ʱ��=����ʱ��_IN;
EXCEPTION
    WHEN OTHERS THEN zl_ErrorCenter(SQLCODE,SQLERRM);
END zl_����ǼǼ�¼_DELETE;
/


--�˹��̽��������Ժ������ͬ�����¾���ǼǼ�¼��״̬
CREATE OR REPLACE PROCEDURE zl_����ǼǼ�¼_����״̬(
	����_IN				����ǼǼ�¼.����%TYPE,
	����ID_IN			����ǼǼ�¼.����ID%TYPE,
	��ҳID_IN			����ǼǼ�¼.��ҳID%TYPE,
	״̬_IN 			����ǼǼ�¼.״̬%TYPE:=0
)
AS 
BEGIN 
	UPDATE ����ǼǼ�¼
	SET ״̬=״̬_IN
	WHERE ����=����_IN AND ����ID=����ID_IN AND Nvl(��ҳID,0)=Nvl(��ҳID_IN,0) ;

EXCEPTION 
    WHEN OTHERS THEN zl_ErrorCenter(SQLCODE,SQLERRM);
END zl_����ǼǼ�¼_����״̬;
/



--�˹��̽���������
CREATE OR REPLACE PROCEDURE zl_����ǼǼ�¼_����(
	����_IN				����ǼǼ�¼.����%TYPE,
	����ID_IN			����ǼǼ�¼.����ID%TYPE,
	��ҳID_IN			����ǼǼ�¼.��ҳID%TYPE,
	����ʱ��_IN			����ǼǼ�¼.����ʱ��%TYPE,
	��¼ID_IN			����ǼǼ�¼.��¼ID%TYPE
)
AS 
BEGIN 
	UPDATE ����ǼǼ�¼
	SET ��¼ID=��¼ID_IN,
		״̬=0
	WHERE ����=����_IN AND ����ID=����ID_IN AND Nvl(��ҳID,0)=Nvl(��ҳID_IN,0) AND ����ʱ��=����ʱ��_IN;

EXCEPTION 
    WHEN OTHERS THEN zl_ErrorCenter(SQLCODE,SQLERRM);
END zl_����ǼǼ�¼_����;
/



Create OR Replace Procedure zl_���ս����¼_Insert(
    ����_IN             ���ս����¼.����%Type,
    ��¼ID_IN           ���ս����¼.��¼ID%Type,
    ����_IN             ���ս����¼.����%Type,
    ����ID_IN           ���ս����¼.����ID%Type,
    ���_IN             ���ս����¼.���%Type,
    �ʻ��ۼ�����_IN     ���ս����¼.�ʻ��ۼ�����%Type,
    �ʻ��ۼ�֧��_IN     ���ս����¼.�ʻ��ۼ�֧��%Type,
    �ۼƽ���ͳ��_IN     ���ս����¼.�ۼƽ���ͳ��%Type,
    �ۼ�ͳ�ﱨ��_IN     ���ս����¼.�ۼ�ͳ�ﱨ��%Type,
    סԺ����_IN         ���ս����¼.סԺ����%Type,
    ����_IN           ���ս����¼.����%Type,
    �ⶥ��_IN           ���ս����¼.�ⶥ��%Type,
    ʵ������_IN       ���ս����¼.ʵ������%Type,
    �������ý��_IN     ���ս����¼.�������ý��%Type,
    ȫ�Ը����_IN       ���ս����¼.ȫ�Ը����%Type,
    �����Ը����_IN     ���ս����¼.�����Ը����%Type,
    ����ͳ����_IN     ���ս����¼.����ͳ����%Type,
    ͳ�ﱨ�����_IN     ���ս����¼.ͳ�ﱨ�����%Type,
    ���Ը����_IN     ���ս����¼.���Ը����%Type,
    �����Ը����_IN     ���ս����¼.�����Ը����%Type,
    �����ʻ�֧��_IN     ���ս����¼.�����ʻ�֧��%Type,
    ֧��˳���_IN       ���ս����¼.֧��˳���%Type,
    ��ҳID_IN           ���ս����¼.��ҳID%Type := null,
    ��;����_IN         ���ս����¼.��;����%Type := null,
    ��ע_IN             ���ս����¼.��ע%Type := null,
	У��_IN             ���ս����¼.У��%TYPE:=0,
	����վ_IN			���ս����¼.����վ%TYPE:=NULL,
	�汾��_IN			���ս����¼.�汾��%TYPE:=NULL,
	ҽ�����_IN			���ս����¼.ҽ�����%TYPE:=NULL,
	������ˮ��_IN		���ս����¼.������ˮ��%TYPE:=NULL,
	����ID_IN			���ս����¼.����ID%TYPE:=NULL,
	��������_IN			���ս����¼.��������%TYPE:=NULL,
	����֢_IN			���ս����¼.����֢%TYPE:=NULL,
	����ʱ��_IN			���ս����¼.����ʱ��%TYPE:=SYSDATE
)
AS
BEGIN
    Update ���ս����¼
        Set ����=����_IN,
            ��¼ID=��¼ID_IN,
            ����=����_IN,
            ����ID=����ID_IN,
            ���=���_IN,
            �ʻ��ۼ�����=�ʻ��ۼ�����_IN,
            �ʻ��ۼ�֧��=�ʻ��ۼ�֧��_IN,
            �ۼƽ���ͳ��=�ۼƽ���ͳ��_IN,
            �ۼ�ͳ�ﱨ��=�ۼ�ͳ�ﱨ��_IN,
            סԺ����=סԺ����_IN,
            ����=����_IN,
            �ⶥ��=�ⶥ��_IN,
            ʵ������=ʵ������_IN,
            �������ý��=�������ý��_IN,
            ȫ�Ը����=ȫ�Ը����_IN,
            �����Ը����=�����Ը����_IN,
            ����ͳ����=����ͳ����_IN,
            ͳ�ﱨ�����=ͳ�ﱨ�����_IN,
            ���Ը����=���Ը����_IN,
            �����Ը����=�����Ը����_IN,
            �����ʻ�֧��=�����ʻ�֧��_IN,
            ֧��˳���=֧��˳���_IN,
            ��ҳID=nvl(��ҳID_IN,��ҳID),
            ��;����=nvl(��;����_IN,��;����),
            ��ע=nvl(��ע_IN,��ע),
			У��=У��_IN,
			�汾��=�汾��_IN,
			ҽ�����=ҽ�����_IN,
			����ID=����ID_IN,
			��������=��������_IN,
			����֢=����֢_IN,
			����ʱ��=����ʱ��_IN ,
			����վ=����վ_IN,
			������ˮ��=������ˮ��_IN
    Where ��¼ID=��¼ID_IN And ����=����_IN;

    IF SQl%RowCount=0 Then
        Insert Into ���ս����¼(
            ����,��¼id,����,����id,���,�ʻ��ۼ�����,�ʻ��ۼ�֧��,�ۼƽ���ͳ��,�ۼ�ͳ�ﱨ��,סԺ����,
            ����,�ⶥ��,ʵ������,�������ý��,ȫ�Ը����,�����Ը����,����ͳ����,ͳ�ﱨ�����,
            ���Ը����,�����Ը����,�����ʻ�֧��,֧��˳���,��ҳID,��;����,��ע,У��,
			����վ,�汾��,ҽ�����,������ˮ��,����ID,��������,����֢,����ʱ��)
        Values(
            ����_IN,��¼id_IN,����_IN,����id_IN,���_IN,�ʻ��ۼ�����_IN,�ʻ��ۼ�֧��_IN,�ۼƽ���ͳ��_IN,
            �ۼ�ͳ�ﱨ��_IN,סԺ����_IN,����_IN,�ⶥ��_IN,ʵ������_IN,�������ý��_IN,ȫ�Ը����_IN,�����Ը����_IN,
            ����ͳ����_IN,ͳ�ﱨ�����_IN,���Ը����_IN,�����Ը����_IN,�����ʻ�֧��_IN,֧��˳���_IN,��ҳID_IN,��;����_IN,��ע_IN,У��_IN,
			����վ_IN,�汾��_IN,ҽ�����_IN,������ˮ��_IN,����ID_IN,��������_IN,����֢_IN,����ʱ��_IN);
    End if;
EXCEPTION
    WHEN OTHERS THEN zl_ErrorCenter(SQLCODE,SQLERRM);
END zl_���ս����¼_Insert;
/


-------------------------------------------------------
--ģ�飺������Ŀ����
--���ܣ���������Ϣ���浽ҽ��������ϸ�У�Ϊ�˱�������ǰ���ݣ���ȱʡ����µĶ�����Ϣͬ�����浽����֧����Ŀ��
----------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_����֧����Ŀ_Modify(
    �շ�ϸĿID_IN	IN ����֧����Ŀ.�շ�ϸĿID%TYPE,
    ����_IN         IN ����֧����Ŀ.����%TYPE,
    ����ID_IN		IN ����֧����Ŀ.����ID%TYPE,
    ��Ŀ����_IN		IN ����֧����Ŀ.��Ŀ����%TYPE,
    ��Ŀ����_IN		IN ����֧����Ŀ.��Ŀ����%TYPE,
    ��ע_IN         IN ����֧����Ŀ.��ע%TYPE,
    �Ƿ�ҽ��_IN     IN ����֧����Ŀ.�Ƿ�ҽ��%TYPE,
	���_IN			IN NUMBER := 0
)
IS 
BEGIN 
	--���ȸ���ҽ��������ϸ���ݣ�û�������
	IF ��Ŀ����_IN IS NOT NULL THEN 
		DELETE ҽ��������ϸ
		WHERE �շ�ϸĿID=�շ�ϸĿID_IN AND ����=����_IN AND ���=���_IN AND ��Ŀ����=��Ŀ����_IN;

		INSERT INTO ҽ��������ϸ(����,���,�շ�ϸĿID,��Ŀ����)
		VALUES (����_IN,���_IN,�շ�ϸĿID_IN,��Ŀ����_IN);
	END IF ;

	--������_IN=0���ʾ��ȱʡ���Ϊ�˱�������ǰ��ģʽ���ݣ�ֱ�Ӹ��±���֧����Ŀ�е�����
	IF ���_IN =0 THEN 
		--�����޸�
		UPDATE ����֧����Ŀ
		SET ����ID=����ID_IN,��Ŀ����=��Ŀ����_IN,��Ŀ����=��Ŀ����_IN,��ע=��ע_IN,�Ƿ�ҽ��=�Ƿ�ҽ��_IN
		WHERE �շ�ϸĿID=�շ�ϸĿID_IN AND ����=����_IN;
		
		IF SQL%NOTFOUND THEN 
			--�����ڣ���Ϊ����
			INSERT INTO ����֧����Ŀ(�շ�ϸĿID,����,����ID,��Ŀ����,��Ŀ����,��ע,�Ƿ�ҽ��)
			VALUES (�շ�ϸĿID_IN,����_IN,����ID_IN,��Ŀ����_IN,��Ŀ����_IN,��ע_IN,�Ƿ�ҽ��_IN);
		END IF;
	END IF ;
EXCEPTION 
    WHEN OTHERS THEN 
        zl_ErrorCenter (SQLCODE, SQLERRM); 
END ZL_����֧����Ŀ_Modify;
/



Create or Replace View �����Զ����� as
Select p.����id, p.��ҳid, i.����, i.�Ա�, i.����, i.סԺ��, a.�ѱ�, p.����id, p.����id, p.����, p.���Ӵ�λ, p.�շ�ϸĿid,
       p.������Ŀid, 1 As ��־, p.�ּ� As ��׼����, p.��ʼ����, p.��ֹ����, p.��ֹ���� - p.��ʼ���� As ����, p.����, p.����ҽʦ,
       p.���λ�ʿ, p.����Ա���, p.����Ա����
From ������Ϣ i, ������ҳ a,
     (Select b.����id, b.��ҳid, b.����id, b.����id, b.����, b.���Ӵ�λ, p.�շ�ϸĿid, p.������Ŀid, p.�ּ�, b.����ҽʦ, b.���λ�ʿ,
              b.����Ա���, b.����Ա����,
              Zl_Date_Half(Greatest(Least(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��), Nvl(b.��ֹʱ��, Sysdate), Sysdate), p.ִ������,
                                     Nvl(a.��������, Add_Months(Sysdate, -2)))) As ��ʼ����,
              Zl_Date_Half(Least(Nvl(b.��ֹʱ��, Sysdate), Nvl(p.��ֹ����, Sysdate + 30) + 1)) As ��ֹ����, b.����
       From �Զ��Ƽ���Ŀ a,
            (Select ����id, ��ҳid, ��ʼʱ��, ���Ӵ�λ, ����id, ����id, ����, ��λ�ȼ�id, 1 As ����, ���λ�ʿ, ����ҽʦ, ��ֹʱ��,
                     ����Ա���, ����Ա����, �ϴμ���ʱ��
              From ���˱䶯��¼
              Where ��ʼԭ�� <> 10
              Union All
              Select ����id, ��ҳid, ��ʼʱ��, ���Ӵ�λ, ����id, ����id, ����, i.����id As ��λ�ȼ�id, i.�������� As ����, ���λ�ʿ,
                     ����ҽʦ, ��ֹʱ��, ����Ա���, ����Ա����, �ϴμ���ʱ��
              From ���˱䶯��¼ b, �շѴ�����Ŀ i
              Where b.��λ�ȼ�id = i.����id And b.��ʼԭ�� <> 10 And i.���д��� > 0) b, �շѼ�Ŀ p
       Where a.����id = b.����id And Zl_Date_Half(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��)) <> Zl_Date_Half(Nvl(b.��ֹʱ��, Sysdate)) And
             p.�ּ� <> 0 And a.�����־ = 1 And b.��λ�ȼ�id = p.�շ�ϸĿid And
             Zl_Date_Half(Nvl(b.��ֹʱ��, Sysdate)) >= Zl_Date_Half(p.ִ������) And
             Zl_Date_Half(b.��ʼʱ��) <= Zl_Date_Half(Nvl(p.��ֹ����, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.��ֹʱ��, Sysdate), Nvl(p.��ֹ����, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.��������, Add_Months(Sysdate, -2)))
       Union All
       Select b.����id, b.��ҳid, b.����id, b.����id, b.����, b.���Ӵ�λ, p.�շ�ϸĿid, p.������Ŀid, p.�ּ�, b.����ҽʦ, b.���λ�ʿ,
              b.����Ա���, b.����Ա����,
              Zl_Date_Half(Greatest(Least(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��), Nvl(b.��ֹʱ��, Sysdate), Sysdate), p.ִ������,
                                     Nvl(a.��������, Add_Months(Sysdate, -2)))) As ��ʼ����,
              Zl_Date_Half(Least(Nvl(b.��ֹʱ��, Sysdate), Nvl(p.��ֹ����, Sysdate + 30) + 1)) As ��ֹ����, b.����
       From �Զ��Ƽ���Ŀ a,
            (Select ����id, ��ҳid, ��ʼʱ��, ���Ӵ�λ, ����id, ����id, ����, ����ȼ�id, 1 As ����, ���λ�ʿ, ����ҽʦ, ��ֹʱ��,
                     ����Ա���, ����Ա����, �ϴμ���ʱ��
              From ���˱䶯��¼
              Where ��ʼԭ�� <> 10
              Union All
              Select ����id, ��ҳid, ��ʼʱ��, ���Ӵ�λ, ����id, ����id, ����, i.����id As ����ȼ�id, i.�������� As ����, ���λ�ʿ,
                     ����ҽʦ, ��ֹʱ��, ����Ա���, ����Ա����, �ϴμ���ʱ��
              From ���˱䶯��¼ b, �շѴ�����Ŀ i
              Where b.����ȼ�id = i.����id And b.��ʼԭ�� <> 10 And i.���д��� > 0) b, �շѼ�Ŀ p
       Where a.����id = b.����id And b.���Ӵ�λ <> 1 And
             Zl_Date_Half(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��)) <> Zl_Date_Half(Nvl(b.��ֹʱ��, Sysdate)) And p.�ּ� <> 0 And
             a.�����־ = 2 And b.����ȼ�id = p.�շ�ϸĿid And Zl_Date_Half(Nvl(b.��ֹʱ��, Sysdate)) >= Zl_Date_Half(p.ִ������) And
             Zl_Date_Half(b.��ʼʱ��) <= Zl_Date_Half(Nvl(p.��ֹ����, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.��ֹʱ��, Sysdate), Nvl(p.��ֹ����, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.��������, Add_Months(Sysdate, -2)))
       Union All
       Select b.����id, b.��ҳid, b.����id, b.����id, b.����, b.���Ӵ�λ, p.�շ�ϸĿid, p.������Ŀid, p.�ּ�, b.����ҽʦ, b.���λ�ʿ,
              b.����Ա���, b.����Ա����,
              Zl_Date_Half(Greatest(Least(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��), Nvl(b.��ֹʱ��, Sysdate), Sysdate), p.ִ������,
                                     Nvl(a.��������, Add_Months(Sysdate, -2)))) As ��ʼ����,
              Zl_Date_Half(Least(Nvl(b.��ֹʱ��, Sysdate), Nvl(p.��ֹ����, Sysdate + 30) + 1)) As ��ֹ����, a.����
       From (Select ����id, �����־, �շ�ϸĿid, 1 As ����, ��������
              From �Զ��Ƽ���Ŀ
              Union All
              Select ����id, �����־, ����id, i.�������� As ����, ��������
              From �Զ��Ƽ���Ŀ a, �շѴ�����Ŀ i
              Where a.�շ�ϸĿid = i.����id And i.���д��� > 0) a, ���˱䶯��¼ b, �շѼ�Ŀ p
       Where a.����id = b.����id And b.���Ӵ�λ <> 1 And b.��ʼԭ�� <> 10 And
             Zl_Date_Half(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��)) <> Zl_Date_Half(Nvl(b.��ֹʱ��, Sysdate)) And p.�ּ� <> 0 And
             a.�շ�ϸĿid = p.�շ�ϸĿid And (a.�����־ = 6 And b.��λ�ȼ�id Is Not Null Or a.�����־ = 7) And
             Zl_Date_Half(Nvl(b.��ֹʱ��, Sysdate)) >= Zl_Date_Half(p.ִ������) And
             Zl_Date_Half(b.��ʼʱ��) <= Zl_Date_Half(Nvl(p.��ֹ����, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.��ֹʱ��, Sysdate), Nvl(p.��ֹ����, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.��������, Add_Months(Sysdate, -2)))) p
Where i.����id = p.����id And a.����id = p.����id And a.��ҳid = p.��ҳid;

-------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_�ѱ���ϸ_UPDATE (
    �ѱ�_IN IN �ѱ���ϸ.�ѱ�%TYPE,
    ������ĿID_IN IN �ѱ���ϸ.������ĿID%TYPE,
    ����_IN IN VARCHAR2,
    ���㷽��_IN In Number:=0
)
IS
-----------------------------------------------------------
--����_IN��������д��ʽ���£�  "1:0:34:100;2:35:101:50;3:102:100000:100;
-----------------------------------------------------------
    V_�κ� �ѱ���ϸ.�κ�%TYPE;
    V_Ӧ�ն���ֵ �ѱ���ϸ.Ӧ�ն���ֵ%TYPE;
    V_Ӧ�ն�βֵ �ѱ���ϸ.Ӧ�ն�βֵ%TYPE;
    V_ʵ�ձ��� �ѱ���ϸ.ʵ�ձ���%TYPE;
    Intpos PLS_INTEGER;
    Str���� VARCHAR2 (2000);
BEGIN
    --��ɾ�����е�
    Delete  FROM �ѱ���ϸ WHERE �ѱ� = �ѱ�_IN AND ������ĿID = ������ĿID_IN;

    --�����µķֶ�
    Str���� := ����_IN;

    WHILE Str���� IS NOT NULL LOOP
        Intpos := INSTR (Str����, ':');

        IF Intpos = 0 THEN
            Str���� := '';
        ELSE
            --�õ��κ�
            V_�κ� := TO_NUMBER (SUBSTR (Str����, 1, Intpos - 1));
            Str���� := SUBSTR (Str����, Intpos + 1);
            --�õ�Ӧ�ն���ֵ
            Intpos := INSTR (Str����, ':');
            V_Ӧ�ն���ֵ := TO_NUMBER (SUBSTR (Str����, 1, Intpos - 1));
            Str���� := SUBSTR (Str����, Intpos + 1);
            --�õ�Ӧ�ն�βֵ
            Intpos := INSTR (Str����, ':');
            V_Ӧ�ն�βֵ := TO_NUMBER (SUBSTR (Str����, 1, Intpos - 1));
            Str���� := SUBSTR (Str����, Intpos + 1);
            --�õ�ʵ�ձ���
            Intpos := INSTR (Str����, ';');
            V_ʵ�ձ��� := TO_NUMBER (SUBSTR (Str����, 1, Intpos - 1));
            Str���� := SUBSTR (Str����, Intpos + 1);

            Insert INTO �ѱ���ϸ(�ѱ�,������ĿID,���㷽��,�κ�,Ӧ�ն���ֵ,Ӧ�ն�βֵ,ʵ�ձ���)
                   VALUES (�ѱ�_IN,������ĿID_IN,���㷽��_IN,V_�κ�,V_Ӧ�ն���ֵ,V_Ӧ�ն�βֵ,V_ʵ�ձ���);
        END IF;
    END LOOP;
END zl_�ѱ���ϸ_UPDATE;
/

CREATE OR REPLACE PROCEDURE zl_������Ϣ_Insert (
    ����ID_IN		������Ϣ.����ID%TYPE,
    �����_IN       ������Ϣ.�����%TYPE,
    �ѱ�_IN         ������Ϣ.�ѱ�%TYPE,
    ҽ�Ƹ���_IN     ������Ϣ.ҽ�Ƹ��ʽ%TYPE,
    ����_IN         ������Ϣ.����%TYPE,
    �Ա�_IN         ������Ϣ.�Ա�%TYPE,
    ����_IN         ������Ϣ.����%TYPE,
    ��������_IN     ������Ϣ.��������%TYPE,
    �����ص�_IN     ������Ϣ.�����ص�%TYPE,
    ���֤��_IN     ������Ϣ.���֤��%TYPE,
    ���_IN         ������Ϣ.���%TYPE,
    ְҵ_IN         ������Ϣ.ְҵ%TYPE,
    ����_IN         ������Ϣ.����%TYPE,
    ����_IN         ������Ϣ.����%TYPE,
    ѧ��_IN         ������Ϣ.ѧ��%TYPE,
    ����_IN         ������Ϣ.����״��%TYPE,
    ��ͥ��ַ_IN     ������Ϣ.��ͥ��ַ%TYPE,
    ��ͥ�绰_IN     ������Ϣ.��ͥ�绰%TYPE,
    �����ʱ�_IN     ������Ϣ.�����ʱ�%TYPE,
    ��ϵ������_IN   ������Ϣ.��ϵ������%TYPE,
    ��ϵ�˹�ϵ_IN   ������Ϣ.��ϵ�˹�ϵ%TYPE,
    ��ϵ�˵�ַ_IN   ������Ϣ.��ϵ�˵�ַ%TYPE,
    ��ϵ�˵绰_IN   ������Ϣ.��ϵ�˵绰%TYPE,
    ��ͬ��λID_IN   ������Ϣ.��ͬ��λID%TYPE,
    ������λ_IN     ������Ϣ.������λ%TYPE,
    ��λ�绰_IN     ������Ϣ.��λ�绰%TYPE,
    ��λ�ʱ�_IN     ������Ϣ.��λ�ʱ�%TYPE,
    ��λ������_IN   ������Ϣ.��λ������%TYPE,
    ��λ�ʺ�_IN     ������Ϣ.��λ�ʺ�%TYPE,
    ������_IN       ������Ϣ.������%TYPE,
    ������_IN       ������Ϣ.������%TYPE,
    ����_IN         ������Ϣ.����%TYPE,
    �Ǽ�ʱ��_IN     ������Ϣ.�Ǽ�ʱ��%TYPE,
	����_IN			������Ϣ.����%Type:=NULL,
	��������_IN		������Ϣ.��������%Type:=NULL,
    ����Ա���_IN   ���˵�����¼.����Ա���%Type:=NULL,
    ����Ա����_IN   ���˵�����¼.����Ա����%Type:=NULL
)
AS
BEGIN
    Insert INTO ������Ϣ (
        ����ID,�����,�ѱ�,ҽ�Ƹ��ʽ,����,�Ա�,����,��������,�����ص�,
        ���֤��,���,ְҵ,����,����,����,ѧ��,����״��,��ͥ��ַ,��ͥ�绰,
        �����ʱ�,��ϵ������,��ϵ�˹�ϵ,��ϵ�˵�ַ,��ϵ�˵绰,��ͬ��λID,
        ������λ,��λ�绰,��λ�ʱ�,��λ������,��λ�ʺ�,������,������,��������,����,�Ǽ�ʱ��)
    VALUES (
        ����ID_IN,�����_IN,�ѱ�_IN,ҽ�Ƹ���_IN,����_IN,�Ա�_IN,����_IN,��������_IN,
        �����ص�_IN,���֤��_IN,���_IN,ְҵ_IN,����_IN,����_IN,����_IN,ѧ��_IN,����_IN,
        ��ͥ��ַ_IN,��ͥ�绰_IN,�����ʱ�_IN,��ϵ������_IN,��ϵ�˹�ϵ_IN,��ϵ�˵�ַ_IN,
        ��ϵ�˵绰_IN,DECODE (��ͬ��λID_IN, 0, NULL, ��ͬ��λID_IN),������λ_IN,
        ��λ�绰_IN,��λ�ʱ�_IN,��λ������_IN,��λ�ʺ�_IN,������_IN,
		DECODE (������_IN,0,NULL,������_IN),��������_IN,����_IN,�Ǽ�ʱ��_IN);
    
    IF �����_IN is Not NULL Then
        Insert Into ���ﲡ����¼(
            ����ID,������,��������,�������,�洢״̬,���λ��)
        Values(
            ����ID_IN,�����_IN,�Ǽ�ʱ��_IN,'һ��','����',NULL);
    End IF;
    
    If ������_IN Is Not Null Then 
      Insert Into ���˵�����¼(����Id,������,������,��������,����Ա���,����Ա����,����ʱ��)
         Values(����ID_IN,������_IN,������_IN,��������_IN,����Ա���_IN,����Ա����_IN,�Ǽ�ʱ��_IN);
    End If;
EXCEPTION
    WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_������Ϣ_Insert;
/

CREATE OR REPLACE PROCEDURE zl_������Ϣ_UPDATE (
    ����ID_IN			������Ϣ.����ID%TYPE,
    �����_IN           ������Ϣ.�����%TYPE,
    סԺ��_IN           ������Ϣ.סԺ��%TYPE,
    �ѱ�_IN             ������Ϣ.�ѱ�%TYPE,
    ҽ�Ƹ���_IN         ������Ϣ.ҽ�Ƹ��ʽ%TYPE,
    ����_IN             ������Ϣ.����%TYPE,
    �Ա�_IN             ������Ϣ.�Ա�%TYPE,
    ����_IN             ������Ϣ.����%TYPE,
    ��������_IN         ������Ϣ.��������%TYPE,
    �����ص�_IN         ������Ϣ.�����ص�%TYPE,
    ���֤��_IN         ������Ϣ.���֤��%TYPE,
    ���_IN             ������Ϣ.���%TYPE,
    ְҵ_IN             ������Ϣ.ְҵ%TYPE,
    ����_IN             ������Ϣ.����%TYPE,
    ����_IN             ������Ϣ.����%TYPE,
    ѧ��_IN             ������Ϣ.ѧ��%TYPE,
    ����_IN             ������Ϣ.����״��%TYPE,
    ��ͥ��ַ_IN         ������Ϣ.��ͥ��ַ%TYPE,
    ��ͥ�绰_IN         ������Ϣ.��ͥ�绰%TYPE,
    �����ʱ�_IN         ������Ϣ.�����ʱ�%TYPE,
    ��ϵ������_IN       ������Ϣ.��ϵ������%TYPE,
    ��ϵ�˹�ϵ_IN       ������Ϣ.��ϵ�˹�ϵ%TYPE,
    ��ϵ�˵�ַ_IN       ������Ϣ.��ϵ�˵�ַ%TYPE,
    ��ϵ�˵绰_IN       ������Ϣ.��ϵ�˵绰%TYPE,
    ��ͬ��λID_IN       ������Ϣ.��ͬ��λID%TYPE,
    ������λ_IN         ������Ϣ.������λ%TYPE,
    ��λ�绰_IN         ������Ϣ.��λ�绰%TYPE,
    ��λ�ʱ�_IN         ������Ϣ.��λ�ʱ�%TYPE,
    ��λ������_IN       ������Ϣ.��λ������%TYPE,
    ��λ�ʺ�_IN         ������Ϣ.��λ�ʺ�%TYPE,
    ������_IN           ������Ϣ.������%TYPE,
    ������_IN           ������Ϣ.������%TYPE,
    ����_IN             ������Ϣ.����%TYPE,
    סԺ�ѱ�_IN         Number:=0,--�Ƿ��޸ĵ��ǲ��˵�סԺ�ѱ�
	ҽ����_IN			�����ʻ�.ҽ����%Type:=NULL,
	����_IN				������Ϣ.����%Type:=NULL,
	��������_IN			������Ϣ.��������%Type:=NULL,
    ����Ա���_IN       ���˵�����¼.����Ա���%Type:=NULL,
    ����Ա����_IN       ���˵�����¼.����Ա����%Type:=NULL
)
AS
	v_��ҳID	        ������ҳ.��ҳID%Type;
    v_������            ������Ϣ.������%Type;
    v_������            ������Ϣ.������%Type;
    v_��������          ������Ϣ.��������%Type;
Begin    
    Select Nvl(������,'�����ܼ�'),Nvl(������,0),Nvl(��������,0) 
           Into v_������,v_������,v_�������� 
    From ������Ϣ Where ����Id=����ID_IN;
    If ������_IN<>v_������ Or (������_IN Is Null And v_������<>'�����ܼ�') Or ������_IN<>v_������ Or ��������_IN<>v_�������� Then 
       Insert Into ���˵�����¼(����Id,������,������,��������,����Ա���,����Ա����,����ʱ��)
       Values(����ID_IN,������_IN,������_IN,��������_IN,����Ա���_IN,����Ա����_IN,Sysdate);
    End If;

    UPDATE ������Ϣ
        SET �����=�����_IN,סԺ��=סԺ��_IN,ҽ�Ƹ��ʽ=ҽ�Ƹ���_IN,�ѱ�=Decode(Nvl(סԺ�ѱ�_IN,0),0,�ѱ�_IN,�ѱ�),
            ����=����_IN,�Ա�=�Ա�_IN,����=����_IN,��������=��������_IN,�����ص�=�����ص�_IN,
            ���֤��=���֤��_IN,���=���_IN,ְҵ=ְҵ_IN,
            ����=����_IN,����=����_IN,����=����_IN,ѧ��=ѧ��_IN,
            ����״��=����_IN,��ͥ��ַ=��ͥ��ַ_IN,��ͥ�绰=��ͥ�绰_IN,
            �����ʱ�=�����ʱ�_IN,��ϵ������=��ϵ������_IN,��ϵ�˹�ϵ=��ϵ�˹�ϵ_IN,
            ��ϵ�˵�ַ=��ϵ�˵�ַ_IN,��ϵ�˵绰=��ϵ�˵绰_IN,
            ��ͬ��λID=DECODE (��ͬ��λID_IN, 0, NULL, ��ͬ��λID_IN),
            ������λ=������λ_IN,��λ�绰=��λ�绰_IN,��λ�ʱ�=��λ�ʱ�_IN,
            ��λ������=��λ������_IN,��λ�ʺ�=��λ�ʺ�_IN,
            ������=������_IN,������=DECODE (������_IN,0,NULL,������_IN),
            ��������=��������_IN,����=����_IN
    WHERE ����ID=����ID_IN;
    
    IF �����_IN is Not NULL Then
        Update ���ﲡ����¼ Set ������=�����_IN Where ����ID=����ID_IN;
        IF SQL%RowCount=0 Then
            Insert Into ���ﲡ����¼(
                ����ID,������,��������,�������,�洢״̬,���λ��)
            Values(
                ����ID_IN,�����_IN,Sysdate,'һ��','����',NULL);
        End IF;
	Else
		Delete From ���ﲡ����¼ Where ����ID=����ID_IN;
    End IF;

    IF סԺ��_IN is Not NULL Then
        Update סԺ������¼ Set ������=סԺ��_IN Where ����ID=����ID_IN;
        IF SQL%RowCount=0 Then
            Insert Into סԺ������¼(
                ����ID,������,��������,�������,�洢״̬,���λ��)
            Values(
                ����ID_IN,סԺ��_IN,Sysdate,'һ��','��Ժ',NULL);
        End IF;
	Else
		Delete From סԺ������¼ Where ����ID=����ID_IN;
    End IF;
    
	Begin
		Select Max(��ҳID) Into v_��ҳID From ������ҳ Where ����ID=����ID_IN;
	Exception
		When Others Then NULL;
	End;
	If v_��ҳID IS Not NULL Then
		Update ������ҳ 
			Set �ѱ�=Decode(Nvl(סԺ�ѱ�_IN,0),1,�ѱ�_IN,�ѱ�),
				ҽ�Ƹ��ʽ=ҽ�Ƹ���_IN,
				����=Decode(����_IN,NULL,����,����_IN)
		Where ����ID=����ID_IN And ��ҳID=v_��ҳID;
		
		If ҽ����_IN IS Not NULL Then
			Update ������ҳ�ӱ� Set ��Ϣֵ=ҽ����_IN Where ����ID=����ID_IN And ��ҳID=v_��ҳID And ��Ϣ��='ҽ����';
			If SQL%RowCount=0 Then
				Insert Into ������ҳ�ӱ�(
					����ID,��ҳID,��Ϣ��,��Ϣֵ)
				Values(
					����ID_IN,v_��ҳID,'ҽ����',ҽ����_IN);
			End IF;
		Else
			Delete From ������ҳ�ӱ� Where ����ID=����ID_IN And ��ҳID=v_��ҳID And ��Ϣ��='ҽ����';
		End IF;
	End IF;
EXCEPTION
    WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_������Ϣ_UPDATE;
/

CREATE OR REPLACE Procedure zl_���˽��ʼ�¼_Delete(
    No_IN            ���˽��ʼ�¼.No%Type, 
    ����Ա���_IN    ���˽��ʼ�¼.����Ա���%Type, 
    ����Ա����_IN    ���˽��ʼ�¼.����Ա����%Type, 
    �����_IN      ����Ԥ����¼.��Ԥ��%Type :=0,        --ҽ����֧���˻�ʱ,ת�ֽ���������
    ���NO_IN        ���˷��ü�¼.No%Type:=Null,          --������ʱ�����,��������ʱ�����ʱ����ֵ
    �������Ͻ���_IN      Varchar2:=Null                   --���㷽ʽ|������|�������||......   
) 
AS 
    --���α�����Ԥ����¼�����Ϣ
    Cursor c_Deposit(v_ID ����Ԥ����¼.����ID%Type) is 
        Select * From ����Ԥ����¼ Where ����ID = v_ID; 
    r_DepositRow c_Deposit%RowType; 
 
    --���α����ڴ��������ػ��ܱ� 
    Cursor c_Money (v_ID ����Ԥ����¼.����ID%Type) is 
        Select * From ���˷��ü�¼ Where ����ID = v_ID; 
    r_MoneyRow c_Money%RowType; 
    
    --���α���������Ŀ�������Ϣ
    Cursor c_ErrItem is 
        Select 
            A.��� AS �շ����,A.ID AS �շ�ϸĿID,A.���㵥λ,C.ID AS ������ĿID,C.�վݷ�Ŀ 
        From �շ�ϸĿ A,�շѼ�Ŀ B,������Ŀ C,�շ��ض���Ŀ D
        Where D.�ض���Ŀ='�����' And D.�շ�ϸĿID=A.ID
            And A.ID=B.�շ�ϸĿID And B.������ĿID=C.ID
            And ((Sysdate Between B.ִ������ And B.��ֹ����) 
                Or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL));
    r_ErrItem c_ErrItem%RowType;

    --���α�������˵������Ϣ
    Cursor c_Pati(v_����Id ������Ϣ.����ID%Type) is
        Select A.����,A.�Ա�,A.����,A.סԺ��,A.�����,B.��ҳID,B.��Ժ����,
            B.��ǰ����ID,B.��Ժ����ID,Nvl(B.�ѱ�,A.�ѱ�) AS �ѱ�
        From ������Ϣ A,������ҳ B
        Where A.����ID=v_����Id And A.����ID=B.����ID(+) And Nvl(A.סԺ����,0)=B.��ҳID(+);
    r_Pati c_Pati%RowType;

    --���̱���
    v_��������		Varchar2(500);
    v_��ǰ����		Varchar2(50);
    v_���㷽ʽ		����Ԥ����¼.���㷽ʽ%Type;
    v_������		����Ԥ����¼.��Ԥ��%Type;
    v_�������		����Ԥ����¼.�������%Type;
    
    v_Temp          Varchar2(255);
    v_����Id        ������Ϣ.����ID%Type;
    v_��Ա����ID    ������Ա.����ID%Type;
 
    v_ԭID		      ���˽��ʼ�¼.ID%Type; 
    v_����ID        ���˽��ʼ�¼.ID%Type; 
    v_��ӡID        Ʊ�ݴ�ӡ����.ID%Type;    
    v_ʵ��Ʊ��      ����Ԥ����¼.ʵ��Ʊ��%Type; 

    v_���NO        ���˷��ü�¼.NO%Type;     
    v_Date		      Date;  
    Err_Custom      Exception; 
    v_Error         Varchar2 (255); 
Begin 
    Begin 
        Select ID,����Id,ʵ��Ʊ�� Into v_ԭID,v_����Id,v_ʵ��Ʊ�� From ���˽��ʼ�¼ Where ��¼״̬ = 1 And No = No_IN; 
        --���һ�δ�ӡ������
        Select Max(ID) Into v_��ӡID From Ʊ�ݴ�ӡ���� Where ��������=3 And NO=NO_IN;
    Exception 
        When Others Then 
        Begin 
            v_Error := 'û�з���Ҫ���ϵĽ��ʵ���,�����Ѿ����ϣ�'; 
            Raise Err_Custom; 
        End; 
    End; 
    Open c_Pati(v_����Id);
    Fetch c_Pati Into r_Pati;    --���ϵͳ���ô˹���,�������ʱû�в�����Ϣ
 
    Select Sysdate Into v_Date From Dual; 
    Select ���˽��ʼ�¼_ID.Nextval Into v_����ID From Dual;   
 
    --���˽��ʼ�¼ 
    Insert Into ���˽��ʼ�¼( 
        ID,No,ʵ��Ʊ��,��¼״̬,����ID,����Ա���,����Ա����,��ʼ����,��������,�շ�ʱ��) 
    Select 
        v_����ID,No,ʵ��Ʊ��,2,����ID,����Ա���_IN,����Ա����_IN,��ʼ����,��������,v_Date 
    From ���˽��ʼ�¼ Where ID = v_ԭID; 
 
    Update ���˽��ʼ�¼ Set ��¼״̬=3 Where ID=v_ԭID; 
 
    --�����ջ�Ʊ��(������ǰû��ʹ��Ʊ��,�޷��ջ�)
    IF v_��ӡID is NOT Null Then 
        Insert Into Ʊ��ʹ����ϸ(
            ID,Ʊ��,����,����,ԭ��,����ID,��ӡID,ʹ��ʱ��,ʹ����) 
        Select 
            Ʊ��ʹ����ϸ_ID.Nextval,Ʊ��,����,2,2,����ID,��ӡID,v_Date,����Ա����_IN
        From Ʊ��ʹ����ϸ
        Where ��ӡID=v_��ӡID And Ʊ��=3 And ����=1;
    End IF; 
 
    --����Ԥ����¼(��Ԥ�����ɿ�)     
    IF �������Ͻ���_IN Is Null Then 
        --��ҽ���������� 
        Insert Into ����Ԥ����¼( 
            ID,No,ʵ��Ʊ��,��¼����,��¼״̬,����ID,��ҳID,����ID,���,���㷽ʽ,�������,ժҪ, 
            �ɿλ,��λ������,��λ�ʺ�,�տ�ʱ��,����Ա����,����Ա���,��Ԥ��,����ID) 
        Select 
            ����Ԥ����¼_ID.Nextval,No,ʵ��Ʊ��,to_Number('1'||Substr(��¼����,Length(��¼����),1)), 
            ��¼״̬,����ID,��ҳID,����ID,Null,���㷽ʽ,�������,ժҪ,�ɿλ,��λ������,��λ�ʺ�, 
            �տ�ʱ��,����Ա����,����Ա���,-1*��Ԥ��,v_����ID 
        From ����Ԥ����¼ 
        Where ����ID=v_ԭID; 
    Else          
        --1.�ȴ����Ԥ������        
        Insert Into ����Ԥ����¼( 
            ID,No,ʵ��Ʊ��,��¼����,��¼״̬,����ID,��ҳID,����ID,���,���㷽ʽ,�������,ժҪ, 
            �ɿλ,��λ������,��λ�ʺ�,�տ�ʱ��,����Ա����,����Ա���,��Ԥ��,����ID) 
        Select 
            ����Ԥ����¼_ID.Nextval,No,ʵ��Ʊ��,to_Number('1'||Substr(��¼����,Length(��¼����),1)), 
            ��¼״̬,����ID,��ҳID,����ID,Null,���㷽ʽ,�������,ժҪ,�ɿλ,��λ������,��λ�ʺ�, 
            �տ�ʱ��,����Ա����,����Ա���,-1*��Ԥ��,v_����ID 
        From ����Ԥ����¼ 
        Where ����ID=v_ԭID And ��¼���� In(1,11); 
 
        --2.�ٴ�����ʽ���,����ҽ���ͷ�ҽ��         
        v_��������:=�������Ͻ���_IN||' ||';--�Կո�ֿ���|��β,û�н�������
        While v_�������� IS Not NULL Loop
            v_��ǰ����:=Substr(v_��������,1,Instr(v_��������,'||')-1);
            v_���㷽ʽ:=Substr(v_��ǰ����,1,Instr(v_��ǰ����,'|')-1);
            v_��ǰ����:=Substr(v_��ǰ����,Instr(v_��ǰ����,'|')+1);
            v_������:=To_Number(Substr(v_��ǰ����,1,Instr(v_��ǰ����,'|')-1));
            v_�������:=LTrim(Substr(v_��ǰ����,Instr(v_��ǰ����,'|')+1));
                        
            Insert Into ����Ԥ����¼( 
                ID,No,ʵ��Ʊ��,��¼����,��¼״̬,����ID,��ҳID,����ID,���,���㷽ʽ,�������,ժҪ, 
                �ɿλ,��λ������,��λ�ʺ�,�տ�ʱ��,����Ա����,����Ա���,��Ԥ��,����ID) 
            Values(
                ����Ԥ����¼_ID.Nextval,No_IN,v_ʵ��Ʊ��,12,1,v_����ID,r_Pati.��ҳID,r_Pati.��Ժ����ID,Null,v_���㷽ʽ,v_�������,'��������ҽ�������˷�',
                Null,Null,Null,v_Date,����Ա����_IN,����Ա���_IN,-1*v_������,v_����ID );
                            
            v_��������:=Substr(v_��������,Instr(v_��������,'||')+2);
        End Loop;
    End IF; 
 
    --���˷��ü�¼ 
    --��ȡԭ����ʱ������������,������,Ȼ�����ʼ�¼�����ڱ��ν���������    
    Begin
        Select No Into v_���No	From ���˷��ü�¼ Where ����ID=v_ԭID And Nvl(���ӱ�־,0)=9 And ��¼����=2 And ��¼״̬=1;
    Exception
        When Others Then NULL;
    End;
    If v_���NO IS Not NULL Then
    		--a.����ʱ�����       ����ʱ�������:
          --1.ԭ����(��ͨ�շѻ�ҽ��ȫ�����ʻ���)  :���ԭ����(ǰ��Ӹ���,����IDΪ�µ�),���¾ɽ���ID�����ļ�¼״̬,���ʴ����
          --2.ҽ��ֻ��������,��û�����,        :�����κδ���
          --3.ҽ��ֻ��������,�������,          :���������(ǰ��Ӹ���,����IDΪ�µ�),���¾ɽ���ID�����ļ�¼״̬,���ʴ����             
        If �������Ͻ���_IN Is Null Or �����_IN<>0 Then
            Insert Into ���˷��ü�¼(
                ID,NO,ʵ��Ʊ��,��¼����,��¼״̬,���,��������,�۸񸸺�,�ಡ�˵�,����ID,��ҳID,ҽ�����,
                �����־,����,�Ա�,����,��ʶ��,����,���˲���ID,���˿���ID,�ѱ�,�շ����,�շ�ϸĿID,���㵥λ,����,
                ��ҩ����,����,�Ӱ��־,���ӱ�־,Ӥ����,���ʷ���,������ĿID,�վݷ�Ŀ,��׼����,Ӧ�ս��,ʵ�ս��,
                ������,��������ID,������,����ʱ��,�Ǽ�ʱ��,ִ�в���ID,ִ��״̬,����Ա����,����Ա���,���ʽ��,����ID,�Ƿ��ϴ�)
            Select
                ���˷��ü�¼_ID.Nextval,No,NULL,��¼����,2,1,��������,�۸񸸺�,�ಡ�˵�,����ID,��ҳID,ҽ�����,
                �����־,����,�Ա�,����,��ʶ��,����,���˲���ID,���˿���ID,�ѱ�,�շ����,�շ�ϸĿID,���㵥λ,����,
                ��ҩ����,-1*����,�Ӱ��־,���ӱ�־,Ӥ����,���ʷ���,������ĿID,�վݷ�Ŀ,Decode(�������Ͻ���_IN,Null,Ӧ�ս��,�����_IN),
                Decode(�������Ͻ���_IN,Null,-1*Ӧ�ս��,-1*�����_IN),Decode(�������Ͻ���_IN,Null,-1*ʵ�ս��,-1*�����_IN),
                ������,��������ID,������,����ʱ��,�Ǽ�ʱ��,ִ�в���ID,ִ��״̬,����Ա����,����Ա���, 0,v_����ID,1
            From ���˷��ü�¼
            Where ����ID=v_ԭID And Nvl(���ӱ�־,0)=9 And ��¼����=2 And ��¼״̬=1;
            
            --����ɵ�����¼
            Update ���˷��ü�¼
                Set ��¼״̬=3,ִ��״̬=0
            Where ����ID=v_ԭID And Nvl(���ӱ�־,0)=9 And ��¼����=2 And ��¼״̬=1;        
            
            --���²�������������¼���н���
            Update ���˷��ü�¼ Set ���ʽ��=ʵ�ս��,����ID=v_����ID,�Ƿ��ϴ�=1
                Where NO=v_���NO And Nvl(���ӱ�־,0)=9 And ��¼����=2 And ��¼״̬=2;  
        End If;      
    Else           
        --b.����ʱ�²��������
        IF �����_IN<>0 Then
            v_Temp:=zl_Identity;
            v_��Ա����ID:=To_Number(Substr(v_Temp,1,Instr(v_Temp,',')-1));       
            
            Open c_ErrItem;
            Fetch c_ErrItem Into r_ErrItem;
            If c_ErrItem%RowCount=0 Then
                Close c_ErrItem;
                v_Error:='������ȷ��ȡ�������������Ŀ��Ϣ�����ȼ�����Ŀ�Ƿ���ȷ���á�';
                Raise Err_Custom;
            End IF;
            
            Insert Into ���˷��ü�¼(
                ID,NO,ʵ��Ʊ��,��¼����,��¼״̬,���,��������,�۸񸸺�,�ಡ�˵�,����ID,��ҳID,ҽ�����,
                �����־,����,�Ա�,����,��ʶ��,����,���˲���ID,���˿���ID,�ѱ�,�շ����,�շ�ϸĿID,���㵥λ,����,
                ��ҩ����,����,�Ӱ��־,���ӱ�־,Ӥ����,���ʷ���,������ĿID,�վݷ�Ŀ,��׼����,Ӧ�ս��,ʵ�ս��,
                ������,��������ID,������,����ʱ��,�Ǽ�ʱ��,ִ�в���ID,ִ��״̬,����Ա����,����Ա���,
                ���ʽ��,����ID,�Ƿ��ϴ�)
            Values(
                ���˷��ü�¼_ID.Nextval,���NO_IN,NULL,2,1,1,NULL,NULL,0,v_����Id,r_Pati.��ҳID,NULL,
                Decode(r_Pati.��ҳID,NULL,1,2),r_Pati.����,r_Pati.�Ա�,r_Pati.����,Decode(r_Pati.��ҳID,NULL,r_Pati.�����,r_Pati.סԺ��),
                r_Pati.��Ժ����,Nvl(r_Pati.��ǰ����ID,v_��Ա����ID),Nvl(r_Pati.��Ժ����ID,v_��Ա����ID),r_Pati.�ѱ�,
                r_ErrItem.�շ����,r_ErrItem.�շ�ϸĿID,r_ErrItem.���㵥λ,1,NULL,1,NULL,9,0,1,r_ErrItem.������ĿID,
                r_ErrItem.�վݷ�Ŀ,-1*�����_IN,-1*�����_IN,-1*�����_IN,����Ա����_IN,v_��Ա����ID,����Ա����_IN,v_Date,
                v_Date,v_��Ա����ID,0,����Ա����_IN,����Ա���_IN,-1*�����_IN,v_����Id,1);   
  
            --���ʴ����       
            Update ���˷��ü�¼ Set ���ʽ��=ʵ�ս��,����ID=v_����ID,�Ƿ��ϴ�=1
            Where NO=���NO_IN And Nvl(���ӱ�־,0)=9 And ��¼����=2 And ��¼״̬=1;  
            
            v_���NO:=���NO_IN;  --�����ں����ſ����Ļ��ܴ���
            Close c_ErrItem;
        End If;  
    End IF;
    
    --���Ͻ��ʶ�Ӧ�ķ��ü�¼:������ԭʼ���ʲ����������Ŀ
    Insert Into ���˷��ü�¼( 
        ID,No,ʵ��Ʊ��,��¼����,��¼״̬,���,��������,�۸񸸺�,�ಡ�˵�,���ʵ�ID,����ID, 
        ��ҳID,ҽ�����,�����־,����,�Ա�,����,��ʶ��,����,���˲���ID,���˿���ID,�ѱ�, 
        �շ����,�շ�ϸĿID,���㵥λ,����,��ҩ����,����,�Ӱ��־,���ӱ�־,Ӥ����,���ʷ���, 
        ������ĿID,�վݷ�Ŀ,��׼����,Ӧ�ս��,ʵ�ս��,������,��������ID,������,����ʱ��, 
        �Ǽ�ʱ��,ִ�в���ID,ִ��״̬,ִ����,ִ��ʱ��,����Ա����,����Ա���,���ʽ��,����ID,
        ������Ŀ��,���մ���ID,ͳ����,�Ƿ���,���ձ���,ժҪ) 
    Select 
        ���˷��ü�¼_ID.Nextval,No,ʵ��Ʊ��,to_Number('1'||Substr(��¼����,Length(��¼����),1)), 
        ��¼״̬,���,��������,�۸񸸺�,�ಡ�˵�,���ʵ�ID,����ID,��ҳID,ҽ�����,�����־,����, 
        �Ա�,����,��ʶ��,����,���˲���ID,���˿���ID,�ѱ�,�շ����,�շ�ϸĿID,���㵥λ,����,��ҩ����, 
        ����,�Ӱ��־,���ӱ�־,Ӥ����,���ʷ���,������ĿID,�վݷ�Ŀ,��׼����,Null,Null,������, 
        ��������ID,������,����ʱ��,�Ǽ�ʱ��,ִ�в���ID,ִ��״̬,ִ����,ִ��ʱ��,����Ա����,����Ա���, 
        -1 * ���ʽ��, v_����ID,
		������Ŀ��,���մ���ID,ͳ����,�Ƿ���,���ձ���,ժҪ
    From ���˷��ü�¼ 
    Where ����ID=v_ԭID And Nvl(���ӱ�־,0)<>9;
 
    --��ػ��ܱ���
    For r_DepositRow in c_Deposit(v_����ID) Loop 
        IF r_DepositRow.��¼���� in (1,11) Then 
            --�������(Ԥ��)     
            Update ������� 
                Set Ԥ�����=Nvl(Ԥ�����,0)-r_DepositRow.��Ԥ�� --ע:�µĽ���ID�������Ǹ������
             Where ����ID=r_DepositRow.����ID And ����=1; 
 
            IF SQL%RowCount = 0 Then 
                Insert Into �������( 
                    ����ID,����,Ԥ�����,�������) 
                Values( 
                    r_DepositRow.����ID,1,-1*r_DepositRow.��Ԥ��,0); 
            End IF; 
        Else 
           --��Ա�ɿ����,ҽ����֧�����ϵĽ��㷽ʽ���µ�Ԥ���������ѱ�����Ϊ�����ֽ�,
            --�˴��ü�,��ʾ�ջ��˸����˵��ֽ�(����ʱ,�˿��Ǹ�,����ʱ����)
            Update ��Ա�ɿ���� 
                Set ���=Nvl(���,0)+r_DepositRow.��Ԥ�� 
             Where �տ�Ա=����Ա����_IN And ���㷽ʽ = r_DepositRow.���㷽ʽ And ���� = 1; 

            IF SQL%RowCount = 0 Then 
                Insert Into ��Ա�ɿ����( 
                    �տ�Ա,���㷽ʽ,����,���) 
                Values( 
                    ����Ա����_IN,r_DepositRow.���㷽ʽ,1,r_DepositRow.��Ԥ��); 
            End IF; 
            Delete From ��Ա�ɿ���� Where �տ�Ա=����Ա����_IN And ���㷽ʽ=r_DepositRow.���㷽ʽ And ����=1 And Nvl(���,0)=0;  
        End IF; 
    End Loop; 
 
    For r_MoneyRow in c_Money(v_����ID) Loop 
        --������� ,������ѽ���,���Բ���Ҫ�������������ܱ�
        If Nvl(v_���NO,'sc')<>Nvl(r_MoneyRow.No,'sc') Then
            Update ������� 
                Set ������� = Nvl(�������,0)-r_MoneyRow.���ʽ��  --ע:�µĽ���ID�������Ǹ������
             Where ����ID=r_MoneyRow.����ID And ����=1; 
     
            IF SQL%RowCount = 0 Then 
                Insert Into �������( 
                    ����ID,����,Ԥ�����,�������) 
                Values( 
                    r_MoneyRow.����ID,1,0,-1*r_MoneyRow.���ʽ��); 
            End IF; 
     
            --����δ����� 
            Update ����δ����� 
                Set ��� = Nvl(���,0)-r_MoneyRow.���ʽ�� 
             Where ����ID=r_MoneyRow.����ID 
                And Nvl(��ҳID,0)=Nvl(r_MoneyRow.��ҳID,0) 
                And Nvl(���˲���ID,0)=Nvl(r_MoneyRow.���˲���ID,0) 
                And Nvl(���˿���ID,0)=Nvl(r_MoneyRow.���˿���ID,0) 
                And Nvl(��������ID,0)=Nvl(r_MoneyRow.��������ID,0) 
                And Nvl(ִ�в���ID,0)=Nvl(r_MoneyRow.ִ�в���ID,0) 
                And ������ĿID+0=r_MoneyRow.������ĿID 
                And ��Դ;��+0=r_MoneyRow.�����־; 
     
            IF SQL%RowCount = 0 Then 
                Insert Into ����δ�����( 
                    ����ID,��ҳID,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,��Դ;��,���) 
                Values( 
                    r_MoneyRow.����ID,r_MoneyRow.��ҳID,r_MoneyRow.���˲���ID,r_MoneyRow.���˿���ID,r_MoneyRow.��������ID, 
                    r_MoneyRow.ִ�в���ID,r_MoneyRow.������ĿID,r_MoneyRow.�����־,-1*r_MoneyRow.���ʽ��); 
            End IF; 
        End If;
 
        --���˷��û��� 
        Update ���˷��û��� 
            Set ���ʽ�� = Nvl(���ʽ��, 0) + r_MoneyRow.���ʽ�� 
         Where ���� = Trunc(v_Date) 
            And Nvl(���˲���ID,0) = Nvl(r_MoneyRow.���˲���ID,0) 
            And Nvl(���˿���ID,0) = Nvl(r_MoneyRow.���˿���ID,0) 
            And Nvl(��������ID,0) = Nvl(r_MoneyRow.��������ID,0) 
            And Nvl(ִ�в���ID,0) = Nvl(r_MoneyRow.ִ�в���ID,0) 
            And ������ĿID+0 = r_MoneyRow.������ĿID 
            And ��Դ;�� = r_MoneyRow.�����־ 
            And ���ʷ��� = r_MoneyRow.���ʷ���; 
 
        IF SQL%RowCount = 0 Then 
            Insert Into ���˷��û���( 
                ����,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,��Դ;��,���ʷ���,Ӧ�ս��,ʵ�ս��,���ʽ��) 
            Values( 
                Trunc (v_Date),r_MoneyRow.���˲���ID,r_MoneyRow.���˿���ID,r_MoneyRow.��������ID,r_MoneyRow.ִ�в���ID, 
                r_MoneyRow.������ĿID,r_MoneyRow.�����־,r_MoneyRow.���ʷ���,0,0,r_MoneyRow.���ʽ��); 
        End IF; 
    End Loop; 
   
    Close c_Pati;
    
Exception 
    When Err_Custom Then Raise_application_errOr (-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]'); 
    When Others Then zl_ErrOrCenter (SQLCODE, SQLERRM); 
End zl_���˽��ʼ�¼_Delete;
/

CREATE OR REPLACE PROCEDURE zl_���ʷ��ü�¼_Insert(
    ID_IN           ���˷��ü�¼.ID%TYPE,
    NO_IN           ���˷��ü�¼.No%TYPE,
    ��¼����_IN		���˷��ü�¼.��¼����%TYPE,
    ��¼״̬_IN     ���˷��ü�¼.��¼״̬%TYPE,
    ִ��״̬_IN     ���˷��ü�¼.ִ��״̬%TYPE,
    ���_IN         ���˷��ü�¼.���%TYPE,
    ���ʽ��_IN     ���˷��ü�¼.���ʽ��%TYPE,
    ����ID_IN       ���˷��ü�¼.����ID%TYPE
)
AS
    v_NextID        ���˷��ü�¼.ID%TYPE;
    v_����ID        ���˷��ü�¼.����ID%TYPE;
    v_��ҳID        ���˷��ü�¼.��ҳID%TYPE;
    v_���˲���ID    ���˷��ü�¼.���˲���ID%TYPE;
    v_���˿���ID    ���˷��ü�¼.���˿���ID%TYPE;
    v_��������ID    ���˷��ü�¼.��������ID%TYPE;
    v_ִ�в���ID    ���˷��ü�¼.ִ�в���ID%TYPE;
    v_������ĿID    ���˷��ü�¼.������ĿID%TYPE;
    v_�����־      ���˷��ü�¼.�����־%TYPE;
    v_���ʷ���      ���˷��ü�¼.���ʷ���%TYPE;
    
    v_���ʽ��      ���˷��ü�¼.���ʽ��%Type;
    v_ʵ�ս��      ���˷��ü�¼.ʵ�ս��%Type;

    Err_Custom      Exception;
    v_Error         Varchar2(255);
BEGIN
    IF ID_IN <> 0 THEN
        --��һ�ν���
        UPDATE ���˷��ü�¼
            SET ���ʽ��=���ʽ��_IN,
                 ����ID=����ID_IN
         WHERE ID=ID_IN And ����ID IS NULL;

        IF SQL%RowCount=0 Then
            v_Error:='�����Ѿ��������˽��ʵķ���,��ǰ���ʲ������ܼ�����';
            Raise Err_Custom;
        End IF;

        v_NextID:=ID_IN;
    ELSE
        --����ǰ������
        SELECT ���˷��ü�¼_ID.Nextval INTO v_NextID FROM Dual;

        Insert INTO ���˷��ü�¼(
            ID,No,ʵ��Ʊ��,��¼����,��¼״̬,���,��������,�۸񸸺�,�ಡ�˵�,���ʵ�ID,
            ����ID,��ҳID,ҽ�����,�����־,����,�Ա�,����,��ʶ��,����,
            ���˲���ID,���˿���ID,�ѱ�,�շ����,�շ�ϸĿID,���㵥λ,����,��ҩ����,����,
            �Ӱ��־,���ӱ�־,Ӥ����,���ʷ���,������ĿID,�վݷ�Ŀ,��׼����,Ӧ�ս��,ʵ�ս��,
            ������,��������ID,������,����ʱ��,�Ǽ�ʱ��,ִ�в���ID,ִ��״̬,ִ����,ִ��ʱ��,
            ����Ա����,����Ա���,���ʽ��,����ID,
			������Ŀ��,���մ���ID,ͳ����,���ձ���,�Ƿ���,ժҪ)
        SELECT v_NextID,NO,ʵ��Ʊ��,TO_NUMBER('1'||��¼����_IN),��¼״̬,
             ���,��������,�۸񸸺�,�ಡ�˵�,���ʵ�ID,����ID,��ҳID,
             ҽ�����,�����־,����,�Ա�,����,��ʶ��,����,
             ���˲���ID,���˿���ID,�ѱ�,�շ����,�շ�ϸĿID,���㵥λ,
             ����,��ҩ����,����,�Ӱ��־,���ӱ�־,Ӥ����,���ʷ���,
             ������ĿID,�վݷ�Ŀ,��׼����,NULL,NULL,������,
             ��������ID,������,����ʱ��,�Ǽ�ʱ��,ִ�в���ID,ִ��״̬,
             ִ����,ִ��ʱ��,����Ա����,����Ա���,���ʽ��_IN,����ID_IN,
			 ������Ŀ��,���մ���ID,ͳ����,���ձ���,�Ƿ���,ժҪ
         FROM ���˷��ü�¼
         WHERE No=No_IN AND ���=���_IN AND ��¼״̬=��¼״̬_IN And Nvl(ִ��״̬,0)=Nvl(ִ��״̬_IN,0)
            AND SUBSTR(��¼����,LENGTH(��¼����),1)=��¼����_IN AND ROWNUM=1;

        --����ν��ʺ���ʽ���Ƿ����ԭ���
        Select Nvl(Sum(ʵ�ս��),0),Nvl(Sum(���ʽ��),0) Into v_ʵ�ս��,v_���ʽ��
        From ���˷��ü�¼
        WHERE NO=NO_IN AND ���=���_IN AND ��¼״̬=��¼״̬_IN
            AND SUBSTR(��¼����,LENGTH(��¼����),1)=��¼����_IN And Nvl(ִ��״̬,0)=ִ��״̬_IN;
        If v_���ʽ��>v_ʵ�ս�� Then
            v_Error:='�����Ѿ��������˽��ʵķ���,��ǰ���ʲ������ܼ�����';
            Raise Err_Custom;
        End IF;
    END IF;

    Select 
        ����ID,��ҳID,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,�����־,���ʷ���
    Into 
        v_����ID,v_��ҳID,v_���˲���ID,v_���˿���ID,v_��������ID,v_ִ�в���ID,v_������ĿID,v_�����־,v_���ʷ���
    From ���˷��ü�¼ Where ID=v_NextID;

    --�������
    UPDATE ������� SET �������=NVL(�������,0)-���ʽ��_IN WHERE ����ID=v_����ID AND ����=1;
    IF SQL%ROWCOUNT=0 THEN
        Insert INTO �������(
            ����ID,����,Ԥ�����,�������) 
        VALUES(
            v_����ID,1,0,-1 * ���ʽ��_IN);
    END IF;
    DELETE FROM ������� WHERE NVL(Ԥ�����,0)=0 AND NVL(�������,0)=0 AND ����ID=v_����ID;

    --����δ�����
    UPDATE ����δ�����
        SET ���=NVL(���,0)-���ʽ��_IN
    WHERE ����ID=v_����ID
        AND NVL(��ҳID,0)=NVL(v_��ҳID,0)
        AND NVL(���˲���ID,0)=NVL(v_���˲���ID,0)
        AND NVL(���˿���ID,0)=NVL(v_���˿���ID,0)
        AND NVL(��������ID,0)=NVL(v_��������ID,0)
        AND NVL(ִ�в���ID,0)=NVL(v_ִ�в���ID,0)
        AND ������ĿID+0=v_������ĿID
        AND ��Դ;��+0=v_�����־;
    IF SQL%ROWCOUNT=0 THEN
        Insert INTO ����δ�����(
            ����ID,��ҳID,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,��Դ;��,���)
        VALUES(
            v_����ID,v_��ҳID,v_���˲���ID,v_���˿���ID,v_��������ID,v_ִ�в���ID,
            v_������ĿID,v_�����־,-1 * ���ʽ��_IN);
    END IF;
    DELETE FROM ����δ����� WHERE ����ID=v_����ID And Nvl(���,0)=0;

    --���˷��û���
    UPDATE ���˷��û���
        SET ���ʽ��=NVL(���ʽ��,0) + ���ʽ��_IN
    WHERE ����=TRUNC(SYSDATE)
        AND NVL(���˲���ID,0)=NVL(v_���˲���ID,0)
        AND NVL(���˿���ID,0)=NVL(v_���˿���ID,0)
        AND NVL(��������ID,0)=NVL(v_��������ID,0)
        AND NVL(ִ�в���ID,0)=NVL(v_ִ�в���ID,0)
        AND ������ĿID+0=v_������ĿID
        AND ��Դ;��=v_�����־
        AND ���ʷ���=v_���ʷ���;
    IF SQL%ROWCOUNT=0 THEN
        Insert INTO ���˷��û���(
            ����,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,��Դ;��,���ʷ���,Ӧ�ս��,ʵ�ս��,���ʽ��)
        VALUES(
            TRUNC(SYSDATE),v_���˲���ID,v_���˿���ID,v_��������ID,v_ִ�в���ID,v_������ĿID,
            v_�����־,v_���ʷ���,0,0,���ʽ��_IN);
    END IF;
EXCEPTION
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    WHEN OTHERS THEN zl_ErrorCenter(SQLCODE,SQLERRM);
END zl_���ʷ��ü�¼_Insert;
/

CREATE OR REPLACE PROCEDURE zl_���˽��ʼ�¼_Insert(
--���ܣ�����һ�����˽��ʼ�¼,ͬʱ������ý������
    ID_IN           ���˽��ʼ�¼.ID%TYPE,
    ���ݺ�_IN       ���˽��ʼ�¼.No%TYPE,
    ����ID_IN       ���˽��ʼ�¼.����ID%TYPE,
    ���NO_IN		���˷��ü�¼.NO%TYPE,
    �����_IN     ���˷��ü�¼.���ʽ��%TYPE,
    �շ�ʱ��_IN     ���˽��ʼ�¼.�շ�ʱ��%TYPE,
    ��ʼ����_IN     ���˽��ʼ�¼.��ʼ����%TYPE,
    ��������_IN     ���˽��ʼ�¼.��������%TYPE
) AS
    --���α���������Ŀ�������Ϣ
    Cursor c_ErrItem is 
        Select 
            A.��� AS �շ����,A.ID AS �շ�ϸĿID,A.���㵥λ,C.ID AS ������ĿID,C.�վݷ�Ŀ 
        From �շ�ϸĿ A,�շѼ�Ŀ B,������Ŀ C,�շ��ض���Ŀ D
        Where D.�ض���Ŀ='�����' And D.�շ�ϸĿID=A.ID
            And A.ID=B.�շ�ϸĿID And B.������ĿID=C.ID
            And ((Sysdate Between B.ִ������ And B.��ֹ����) 
                Or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL));
    r_ErrItem c_ErrItem%RowType;

    --���α�������˵������Ϣ
    Cursor c_Pati is
        Select A.����,A.�Ա�,A.����,A.סԺ��,A.�����,B.��ҳID,B.��Ժ����,
            B.��ǰ����ID,B.��Ժ����ID,Nvl(B.�ѱ�,A.�ѱ�) AS �ѱ�
        From ������Ϣ A,������ҳ B
        Where A.����ID=����ID_IN And A.����ID=B.����ID(+) And Nvl(A.סԺ����,0)=B.��ҳID(+);
    r_Pati c_Pati%RowType;

    v_Temp          Varchar2(255);
    v_��Ա����ID    ������Ա.����ID%Type;
    v_��Ա���		��Ա��.���%Type;
    v_��Ա����      ��Ա��.����%Type;

    v_Error			Varchar2(255);
    Err_Custom		Exception;
BEGIN
    --����Ա��Ϣ:����ID,��������;��ԱID,��Ա���,��Ա����
    v_Temp:=zl_Identity;
    v_��Ա����ID:=To_Number(Substr(v_Temp,1,Instr(v_Temp,',')-1));
    v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
    v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
    v_��Ա���:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
    v_��Ա����:=Substr(v_Temp,Instr(v_Temp,',')+1);
 
    --���˽��ʼ�¼
    Insert INTO ���˽��ʼ�¼(
        ID,No,ʵ��Ʊ��,��¼״̬,����ID,��ʼ����,��������,�շ�ʱ��,����Ա���,����Ա����)
    VALUES(
        ID_IN,���ݺ�_IN,NULL,1,����ID_IN,��ʼ����_IN,��������_IN,�շ�ʱ��_IN,v_��Ա���,v_��Ա����);

    --�������ʱ��������������
    If Nvl(�����_IN,0)<>0 Then
        Open c_ErrItem;
        Fetch c_ErrItem Into r_ErrItem;
        If c_ErrItem%RowCount=0 Then
            Close c_ErrItem;
            v_Error:='������ȷ��ȡ�������������Ŀ��Ϣ�����ȼ�����Ŀ�Ƿ���ȷ���á�';
            Raise Err_Custom;
        End IF;
        
        Open c_Pati;
        Fetch c_Pati Into r_Pati;        
        If c_Pati%RowCount=0 Then
            Close c_Pati;
            v_Error:='������ȷ��ȡ���ʲ�����Ϣ��';
            Raise Err_Custom;
        End IF;
        
        --���˷��ü�¼(����ͬʱ����):���ӱ�־=9
        Insert Into ���˷��ü�¼(
            ID,NO,ʵ��Ʊ��,��¼����,��¼״̬,���,��������,�۸񸸺�,�ಡ�˵�,����ID,��ҳID,ҽ�����,
            �����־,����,�Ա�,����,��ʶ��,����,���˲���ID,���˿���ID,�ѱ�,�շ����,�շ�ϸĿID,���㵥λ,����,
            ��ҩ����,����,�Ӱ��־,���ӱ�־,Ӥ����,���ʷ���,������ĿID,�վݷ�Ŀ,��׼����,Ӧ�ս��,ʵ�ս��,
            ������,��������ID,������,����ʱ��,�Ǽ�ʱ��,ִ�в���ID,ִ��״̬,����Ա����,����Ա���,
            ���ʽ��,����ID,�Ƿ��ϴ�)
        Values(
            ���˷��ü�¼_ID.Nextval,���NO_IN,NULL,2,1,1,NULL,NULL,0,����ID_IN,r_Pati.��ҳID,NULL,
            Decode(r_Pati.��ҳID,NULL,1,2),r_Pati.����,r_Pati.�Ա�,r_Pati.����,Decode(r_Pati.��ҳID,NULL,r_Pati.�����,r_Pati.סԺ��),
            r_Pati.��Ժ����,Nvl(r_Pati.��ǰ����ID,v_��Ա����ID),Nvl(r_Pati.��Ժ����ID,v_��Ա����ID),r_Pati.�ѱ�,
            r_ErrItem.�շ����,r_ErrItem.�շ�ϸĿID,r_ErrItem.���㵥λ,1,NULL,1,NULL,9,0,1,r_ErrItem.������ĿID,
            r_ErrItem.�վݷ�Ŀ,�����_IN,�����_IN,�����_IN,v_��Ա����,v_��Ա����ID,v_��Ա����,�շ�ʱ��_IN,
            �շ�ʱ��_IN,v_��Ա����ID,0,v_��Ա����,v_��Ա���,�����_IN,ID_IN,1);

        --�������,����δ�����(���=���+ʵ��-����,���Կ��Բ�����)

        --���˷��û���
        Update ���˷��û���
            Set Ӧ�ս��=Nvl(Ӧ�ս��,0)+�����_IN,
                ʵ�ս��=Nvl(ʵ�ս��,0)+�����_IN,
                ���ʽ��=Nvl(���ʽ��,0)+�����_IN
        Where ����=Trunc(�շ�ʱ��_IN)
            And Nvl(���˲���ID,0)=Nvl(r_Pati.��ǰ����ID,v_��Ա����ID)
            And Nvl(���˿���ID,0)=Nvl(r_Pati.��Ժ����ID,v_��Ա����ID)
            And Nvl(��������ID,0)=v_��Ա����ID
            And Nvl(ִ�в���ID,0)=v_��Ա����ID
            And ������ĿID+0=r_ErrItem.������ĿID
            And ��Դ;��=Decode(r_Pati.��ҳID,NULL,1,2)
            And ���ʷ���=1;
        IF SQL%RowCount=0 Then
            Insert Into ���˷��û���(
                ����,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,��Դ;��,���ʷ���,Ӧ�ս��,ʵ�ս��,���ʽ��)
            Values(
                Trunc(�շ�ʱ��_IN),Nvl(r_Pati.��ǰ����ID,v_��Ա����ID),Nvl(r_Pati.��Ժ����ID,v_��Ա����ID),v_��Ա����ID,
                v_��Ա����ID,r_ErrItem.������ĿID,Decode(r_Pati.��ҳID,NULL,1,2),1,�����_IN,�����_IN,�����_IN);
        End IF;
            
        Close c_Pati;
        Close c_ErrItem;
    End IF;
EXCEPTION
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_���˽��ʼ�¼_Insert;
/

Create Or Replace Procedure zl_סԺ�շѽ���_Update(
	   ����Id_IN			���˷��ü�¼.����ID%Type,
	   ���ʽ���_IN		    Varchar2,                   --���ʽ���_IN-��ҽ��ʱ:���㷽ʽ|������|�������||.....ҽ��ʱ:���㷽ʽ|������|�������,��������,�����ʺ�||.....
	   ��Ԥ��_IN		    Varchar2,                   --��Ԥ��_IN= ID|���ݺ�|���|��¼״̬||.....
	   �����_IN		    ���˷��ü�¼.ʵ�ս��%Type,
       ���NO_IN		    ���˷��ü�¼.NO%Type
 )
 As 
 --����:�������ʱ��ҽ����ʽ�����,��ؽ�����Ϣ�ĵ���
 --     ��Ϊ������ʺ�,���ɵ�ҽ���������ܶ��̯���ܻ�����ʽ����ʱ�в���,�����ṩ��У�Թ���,
 --		����Ա�ڽ���У��ʱ,���Ե�����ҽ�����㷽ʽ�ĸ��ֽ������ʽ,�������ɽ��㴮,���ҿ��ܲ��������.
 
--������Ϣ 
     Cursor c_Pati(v_����ID ������Ϣ.����ID%Type) is 
     Select A.����,A.�Ա�,A.����,A.סԺ��,A.�����,B.��ҳId,B.��Ժ����,
                B.��ǰ����ID,B.��Ժ����Id,Nvl(B.�ѱ�,A.�ѱ�) AS �ѱ�
            From ������Ϣ A,������ҳ B
            Where A.����ID=v_����Id And A.����ID=B.����ID(+) And Nvl(A.סԺ����,0)=B.��ҳID(+);
     r_pati c_Pati%RowType;
 
 --���̱���
    v_��������		Varchar2(4000);
    v_��ǰ����		Varchar2(100);
    v_���㷽ʽ		����Ԥ����¼.���㷽ʽ%Type;
    v_������		����Ԥ����¼.��Ԥ��%Type;
    v_�������		Varchar2(100);          --���ս����¼ʱ,����:�������,��������,�����ʺ�
	
    v_�շ����		���˷��ü�¼.�շ����%Type;
    v_�շ�ϸĿID	���˷��ü�¼.�շ�ϸĿID%Type;
    v_���㵥λ		���˷��ü�¼.���㵥λ%Type;
    v_������ĿID	���˷��ü�¼.������ĿID%Type;
    v_�վݷ�Ŀ		���˷��ü�¼.�վݷ�Ŀ%Type;
	
    v_No			����Ԥ����¼.No%Type;
    v_����Id		����Ԥ����¼.����Id%Type;   
    v_�տ�ʱ��		����Ԥ����¼.�տ�ʱ��%Type;
    v_����Ա���	����Ԥ����¼.����Ա���%Type;
    v_����Ա����	����Ԥ����¼.����Ա����%Type;
    v_��Ա����ID	������Ա.����ID%Type;
    v_Temp			Varchar2(500);    
	
    v_Ԥ�����		����Ԥ����¼.��Ԥ��%Type;	
    v_Ԥ��ID		����Ԥ����¼.Id%Type;
    v_��¼״̬		����Ԥ����¼.��¼״̬%Type;
	 
    v_�������		����Ԥ����¼.�ɿλ%Type;
    v_�����ʺ�		����Ԥ����¼.��λ������%Type;
    v_��������		����Ԥ����¼.��λ�ʺ�%Type;
   
    v_Error			VARCHAR2(255);
    Err_Custom		EXCEPTION;
 Begin
 
 --1.ȡԤ����¼�е���Ҫ�������Ϣ
    Select No,����Id,�շ�ʱ��,����Ա���,����Ա���� 
		   Into  v_No,v_����Id,v_�տ�ʱ��,v_����Ա���,v_����Ա����
    From ���˽��ʼ�¼ Where ID=����ID_IN;  
    
    Open c_Pati(v_����Id);
    Fetch c_Pati Into r_Pati;
    
    --��������Ϣ
    Begin
        Select A.���,A.ID,A.���㵥λ,C.ID,C.�վݷ�Ŀ 
        Into v_�շ����,v_�շ�ϸĿID,v_���㵥λ,v_������ĿID,v_�վݷ�Ŀ
        From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C,�շ��ض���Ŀ D
        Where D.�ض���Ŀ='�����' And D.�շ�ϸĿID=A.Id  And A.ID=B.�շ�ϸĿID And B.������ĿID=C.ID
            And Sysdate Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-01-01','YYYY-MM-DD')) ;
    Exception
        When Others Then
        Begin
            v_Error:='������ȷ��ȡ�շ���������Ϣ�����ȼ�����Ŀ�Ƿ�������ȷ��';
            Raise Err_Custom;
        End;
    End;	
	
 --2.ɾ���ɵļ�¼,���˻�������
    --������Ա�ɿ����,�������,
    For C_DEL In (SELECT * FROM ����Ԥ����¼ WHERE ����ID=����ID_IN And ��¼����=2) Loop
	      Update ��Ա�ɿ���� Set ���=Nvl(���,0)-Nvl(C_DEL.��Ԥ��,0) Where ���㷽ʽ=C_DEL.���㷽ʽ;	   
      	If SQL%RowCount=0 Then
               Insert Into ��Ա�ɿ����(�տ�Ա,���㷽ʽ,����,���) Values(C_DEL.����Ա����,C_DEL.���㷽ʽ,1,-1*C_DEL.��Ԥ��);
    		End If;
    End Loop;
	
    If v_����Id>0 Then
    	Begin
        	Select Sum(��Ԥ��) Into V_Ԥ����� From ����Ԥ����¼ Where ����Id=����id_IN And ��¼���� In (1,11);
        Exception
        	When Others Then NULL;
    	End;	
    	If v_Ԥ�����<>0 Then
        	Update ������� Set Ԥ�����=Nvl(Ԥ�����,0)+V_Ԥ����� Where ����ID=v_����Id And ����=1;
        	IF SQL%RowCount=0 Then
            	Insert Into �������(����ID,Ԥ�����,����) Values(v_����Id,V_Ԥ�����,1);
            End IF;
    	End If;
    End If;
    
	--���˲��˷��û���.         ����δ�����(��Ϊ������������,���Բ�����)  
	--ֻ���ܲ��������ı仯. �����ֻ���ܴ���һ��,��Ϊ�˱�������������α�
    For C_Error In (
        Select TRUNC(�Ǽ�ʱ��) as ����,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,�����־,Ӧ�ս��,ʵ�ս��,���ʽ��
        From ���˷��ü�¼
        Where ��¼����=2 And ��¼״̬=1 And ����Id=����Id_IN And ���ӱ�־=9
    ) Loop
        Update ���˷��û���
            Set Ӧ�ս��=Nvl(Ӧ�ս��,0)-C_Error.Ӧ�ս��,ʵ�ս��=Nvl(ʵ�ս��,0)-C_Error.ʵ�ս��,���ʽ��=Nvl(���ʽ��,0)-C_Error.���ʽ��
        Where ����=C_Error.����
            And Nvl(���˲���ID,0)=Nvl(C_Error.���˲���ID,0) And Nvl(���˿���ID,0)=Nvl(C_Error.���˿���ID,0)
            And Nvl(��������ID,0)=Nvl(C_Error.��������ID,0) And Nvl(ִ�в���ID,0)=Nvl(C_Error.ִ�в���ID,0)
            And ������ĿID+0=C_Error.������ĿId And ��Դ;��=C_Error.�����־ And ���ʷ���=1; 
        If SQL%RowCount=0 Then
            Insert Into ���˷��û���(
                ����,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,��Դ;��,���ʷ���,Ӧ�ս��,ʵ�ս��,���ʽ��)
            Values(
                C_Error.����,C_Error.���˲���ID,C_Error.���˿���ID,C_Error.��������ID,C_Error.ִ�в���ID,
                C_Error.������ĿID,C_Error.�����־,1,-1*C_Error.Ӧ�ս��,-1*C_Error.ʵ�ս��,-1*C_Error.���ʽ��);
        End If;
    End Loop; 
 
    --ɾ�����ʽɿ�,���ս����¼		     
    Delete ����Ԥ����¼ Where ����ID=����ID_IN And ��¼����=2; 
    --��һ�γ�Ԥ����,��ճ����
    Update ����Ԥ����¼ Set ��Ԥ��=Null,����Id=Null	Where ����Id=����ID_IN And ��¼����=1;
    --ɾ�������
    Delete ����Ԥ����¼ Where ����Id=����ID_IN And ��¼����=11;
    --ɾ������¼
    Delete ���˷��ü�¼ Where ����Id=����Id_IN And ���ӱ�־=9;	
 
 --3.�������˷��ü�¼������¼
    If �����_IN <>0 Then		    
        --����Ա��Ϣ:����ID,��������;��ԱID,��Ա���,��Ա����
        v_Temp:=zl_Identity;
        v_��Ա����ID:=To_Number(Substr(v_Temp,1,Instr(v_Temp,',')-1));             
        Insert Into ���˷��ü�¼(
            ID,NO,ʵ��Ʊ��,��¼����,��¼״̬,���,��������,�۸񸸺�,�ಡ�˵�,����ID,��ҳID,ҽ�����,
            �����־,����,�Ա�,����,��ʶ��,����,���˲���ID,���˿���ID,�ѱ�,�շ����,�շ�ϸĿID,���㵥λ,����,
            ��ҩ����,����,�Ӱ��־,���ӱ�־,Ӥ����,���ʷ���,������ĿID,�վݷ�Ŀ,��׼����,Ӧ�ս��,ʵ�ս��,
            ������,��������ID,������,����ʱ��,�Ǽ�ʱ��,ִ�в���ID,ִ��״̬,����Ա����,����Ա���,
            ���ʽ��,����ID,�Ƿ��ϴ�)
        Values(
            ���˷��ü�¼_ID.Nextval,���NO_IN,NULL,2,1,1,NULL,NULL,0,v_����Id,r_Pati.��ҳID,NULL,
            Decode(r_Pati.��ҳID,NULL,1,2),r_Pati.����,r_Pati.�Ա�,r_Pati.����,Decode(r_Pati.��ҳID,NULL,r_Pati.�����,r_Pati.סԺ��),
            r_Pati.��Ժ����,Nvl(r_Pati.��ǰ����ID,v_��Ա����ID),Nvl(r_Pati.��Ժ����ID,v_��Ա����ID),r_Pati.�ѱ�,
            v_�շ����,v_�շ�ϸĿID,v_���㵥λ,1,NULL,1,NULL,9,0,1,v_������ĿID,
            v_�վݷ�Ŀ,�����_IN,�����_IN,�����_IN,v_����Ա����,v_��Ա����ID,v_����Ա����,v_�տ�ʱ��,
            v_�տ�ʱ��,v_��Ա����ID,0,v_����Ա����,v_����Ա���,�����_IN,����ID_IN,1);
    End If;
  
 --4.�������ɲ���Ԥ����¼�������	
    --4.1.�������,���ս���
    If ���ʽ���_IN IS Not NULL Then
		v_��������:=���ʽ���_IN||'||';
        While v_�������� IS Not NULL Loop
      			v_��ǰ����:=Substr(v_��������,1,Instr(v_��������,'||')-1);
      			v_���㷽ʽ:=Substr(v_��ǰ����,1,Instr(v_��ǰ����,'|')-1);
      			v_��ǰ����:=Substr(v_��ǰ����,Instr(v_��ǰ����,'|')+1);
      			v_������:=To_Number(Substr(v_��ǰ����,1,Instr(v_��ǰ����,'|')-1));
      			v_�������:=LTrim(Substr(v_��ǰ����,Instr(v_��ǰ����,'|')+1));
            If Instr(v_�������,',')>0 Then   --ҽ������:�������,��������,�����ʺ�
                v_�������:=v_�������||',';
                  v_�������:=Substr(v_�������,1,Instr(v_�������,',')-1);
                v_�������:=Substr(v_�������,Instr(v_�������,',')+1);
                  v_��������:=Substr(v_�������,1,Instr(v_�������,',')-1);
                v_�������:=Substr(v_�������,Instr(v_�������,',')+1);
                  v_�����ʺ�:=Substr(v_�������,1,Instr(v_�������,',')-1);
                v_�������:=Null;
            Else
                v_�������:=Null;
                v_��������:=Null;
                v_�����ʺ�:=Null;
            End If;
			
  			If Nvl(v_������,0)<>0 Then
  				Insert Into ����Ԥ����¼(
                    ID,NO,ʵ��Ʊ��,��¼����,��¼״̬,����ID,��ҳID,����ID,���,���㷽ʽ,�������,ժҪ,�ɿλ,��λ������,��λ�ʺ�,
                    �տ�ʱ��,����Ա���,����Ա����,��Ԥ��,����ID)
                Values(
                    ����Ԥ����¼_ID.NextVal,v_No,Null,2,1,v_����Id,r_Pati.��ҳId,r_Pati.��Ժ����Id,
                    Null,v_���㷽ʽ,v_�������,'���ʽɿ�',v_�������,v_��������,v_�����ʺ�,
                    v_�տ�ʱ��,v_����Ա���,v_����Ա����,v_������,����ID_IN);
  			End IF;
            v_��������:=Substr(v_��������,Instr(v_��������,'||')+2);
        End Loop;
     End IF;

    --4.2.Ԥ������
    IF ��Ԥ��_IN Is Not Null Then
        v_��������:=��Ԥ��_IN||'||';
        V_Ԥ�����:=0;              --ǰ�����Ԥ�����ʱ�ù��˱���
        While v_�������� Is Not Null Loop
            v_��ǰ����:=Substr(v_��������,1,Instr(v_��������,'||')-1);
  		v_Ԥ��ID:=To_Number(Substr(v_��ǰ����,1,Instr(v_��ǰ����,'|')-1));      --�Ǽ�¼��Ԥ����ID
            v_��ǰ����:=Substr(v_��ǰ����,Instr(v_��ǰ����,'|')+1);
            v_�������:=Substr(v_��ǰ����,1,Instr(v_��ǰ����,'|')-1);               --�Ǽ�¼��Ԥ����NO��
            v_��ǰ����:=Substr(v_��ǰ����,Instr(v_��ǰ����,'|')+1);
  			v_������:=To_Number(Substr(v_��ǰ����,1,Instr(v_��ǰ����,'|')-1));
            v_��¼״̬:=To_Number(Substr(v_��ǰ����,Instr(v_��ǰ����,'|')+1));
                         
            IF v_Ԥ��Id <> 0 Then              --��һ�γ�Ԥ��
                UPDATE ����Ԥ����¼ SET ��Ԥ�� = v_������, ����ID = ����ID_IN WHERE ID = v_Ԥ��Id;
            Else                            --���ϴ�ʣ���
                Insert INTO ����Ԥ����¼(ID,No,ʵ��Ʊ��,��¼����,��¼״̬,����ID,��ҳID,����ID,���,���㷽ʽ,�������,ժҪ, 
                                        �ɿλ,��λ������,��λ�ʺ�,�տ�ʱ��,����Ա����,����Ա���,��Ԥ��,����Id)
                    SELECT ����Ԥ����¼_ID.Nextval, No, ʵ��Ʊ��, 11, v_��¼״̬, ����ID,��ҳID, ����ID, NULL, ���㷽ʽ, �������, ժҪ, 
                            �ɿλ,��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���,v_������, ����ID_IN
                    FROM ����Ԥ����¼
                    WHERE No = v_������� AND ��¼���� In(1,11) AND ��¼״̬ = v_��¼״̬ AND ROWNUM = 1;
            END IF;
            v_Ԥ�����:=v_Ԥ�����+v_������;
            v_��������:=Substr(v_��������,Instr(v_��������,'||')+2);
        End Loop;
        
        --���²������
        Update ������� Set Ԥ�����=Nvl(Ԥ�����,0)-v_Ԥ����� Where ����ID=v_����Id And ����=1;
        IF SQL%RowCount=0 Then
            Insert Into �������(����ID,Ԥ�����,����) Values(v_����Id,-1*v_Ԥ�����,1);
        End IF;
        Delete From ������� Where ����ID=v_����Id And ����=1 And Nvl(�������,0)=0 And Nvl(Ԥ�����,0)=0;
    End IF;
	
    --5.��ػ��ܱ�Ĵ���	
    --����"��Ա�ɿ����"
	--�ɿ����,���ս���
    IF ���ʽ���_IN IS Not NULL Then
        v_��������:=���ʽ���_IN||'||';
        While v_�������� IS Not NULL Loop
      			v_��ǰ����:=Substr(v_��������,1,Instr(v_��������,'||')-1);
      			v_���㷽ʽ:=Substr(v_��ǰ����,1,Instr(v_��ǰ����,'|')-1);
      			v_��ǰ����:=Substr(v_��ǰ����,Instr(v_��ǰ����,'|')+1);
      			v_������:=To_Number(Substr(v_��ǰ����,1,Instr(v_��ǰ����,'|')-1));
      			
      			If Nvl(v_������,0)<>0 Then
      				Update ��Ա�ɿ����	Set ���=Nvl(���,0)+Nvl(v_������,0)
      				    Where �տ�Ա=v_����Ա���� And ����=1 And ���㷽ʽ=v_���㷽ʽ;
      				If SQL%RowCount=0 Then
      					Insert Into ��Ա�ɿ����(�տ�Ա,���㷽ʽ,����,���)
      					Values(v_����Ա����,v_���㷽ʽ,1,Nvl(v_������,0));
      				End If;
      			End IF;
            v_��������:=Substr(v_��������,Instr(v_��������,'||')+2);
        End Loop;
    End IF;
    Delete From ��Ա�ɿ���� Where ����=1 And �տ�Ա=v_����Ա���� And Nvl(���,0)=0;

    --���˷��û���,ֻ���ػ������,��Ϊ��������,δ����ò���(�²�����������ѽ���),ֻ��һ������¼,��Ϊʹ�ñ�����������α�
    For r_MoneyRow In (
        Select TRUNC(�Ǽ�ʱ��) as ����,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,�����־,Ӧ�ս��,ʵ�ս��,���ʽ��
        From ���˷��ü�¼
        Where ��¼����=2 And ��¼״̬=1 And ����Id=����Id_IN And ���ӱ�־=9
	) Loop
        Update ���˷��û���
            Set Ӧ�ս��=Nvl(Ӧ�ս��,0)+r_MoneyRow.Ӧ�ս��,ʵ�ս��=Nvl(ʵ�ս��,0)+r_MoneyRow.ʵ�ս��,���ʽ��=Nvl(���ʽ��,0)+r_MoneyRow.���ʽ��
        Where ����=r_MoneyRow.����
            And Nvl(���˲���ID,0)=Nvl(r_MoneyRow.���˲���ID,0) And Nvl(���˿���ID,0)=Nvl(r_MoneyRow.���˿���ID,0)
            And Nvl(��������ID,0)=Nvl(r_MoneyRow.��������ID,0) And Nvl(ִ�в���ID,0)=Nvl(r_MoneyRow.ִ�в���ID,0)
            And ������ĿID+0=r_MoneyRow.������ĿId  And ��Դ;��=r_MoneyRow.�����־ And ���ʷ���=1;

        If SQL%RowCount=0 Then
            Insert Into ���˷��û���(
                ����,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,��Դ;��,���ʷ���,Ӧ�ս��,ʵ�ս��,���ʽ��)
            Values(
                r_MoneyRow.����,r_MoneyRow.���˲���ID,r_MoneyRow.���˿���ID,r_MoneyRow.��������ID,r_MoneyRow.ִ�в���ID,
                r_MoneyRow.������ĿID,r_MoneyRow.�����־,1,r_MoneyRow.Ӧ�ս��,r_MoneyRow.ʵ�ս��,r_MoneyRow.���ʽ��);
        End If;
    End Loop; 
 
 	--6.ҽ����ر�Ĵ���
    Delete ҽ���˶Ա� Where ����Id=����Id_IN;
    
    Close c_Pati;
 
EXCEPTION
    WHEN Err_Custom Then Raise_application_errOr(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    WHEN OTHERS Then zl_ErrOrCenter(SQLCODE,SQLERRM);
End zl_סԺ�շѽ���_Update;
/



-------------------------------------------------------
--ģ�飺���￨��¼.SQL
Create Or Replace Procedure zl_���￨��¼_Insert(
--��������������=0-����,1-����,2-����(�൱���ش�)
--      ����ʱ,���ݺ�_IN�������ԭ��/�����ĵ��ݺš�
--      ����/������,�ٻ���ʱ�������һ�ο���Ϊ׼��
    ��������_IN            Number, 
    ���ݺ�_IN            ���˷��ü�¼.No%Type,
    ����ID_IN            ���˷��ü�¼.����id%Type, 
    ��ҳID_IN            ���˷��ü�¼.��ҳid%Type,
    ��ʶ��_IN            ���˷��ü�¼.��ʶ��%Type, 
    �ѱ�_IN                ���˷��ü�¼.�ѱ�%Type,
    ԭ����_IN            ������Ϣ.���￨��%Type, 
    ����_IN                ������Ϣ.���￨��%Type,
    ����_IN                ������Ϣ.����֤��%Type, 
    ����_IN                ���˷��ü�¼.����%Type,
    �Ա�_IN                ���˷��ü�¼.�Ա�%Type, 
    ����_IN                ���˷��ü�¼.����%Type,
    ���˲���ID_IN        ���˷��ü�¼.���˲���id%Type,
    ���˿���ID_IN        ���˷��ü�¼.���˿���id%Type,
    �շ�ϸĿID_IN        ���˷��ü�¼.�շ�ϸĿid%Type,
    �շ����_IN            ���˷��ü�¼.�շ����%Type,
    ���㵥λ_IN            ���˷��ü�¼.���㵥λ%Type,
    ������ĿID_IN        ���˷��ü�¼.������Ŀid%Type,
    �վݷ�Ŀ_IN            ���˷��ü�¼.�վݷ�Ŀ%Type,
    ���_IN                ���˷��ü�¼.ʵ�ս��%Type,
    ִ�в���ID_IN        ���˷��ü�¼.ִ�в���id%Type,
    ��������ID_IN        ���˷��ü�¼.��������id%Type,
    ����Ա���_IN        ���˷��ü�¼.����Ա���%Type,
    ����Ա����_IN        ���˷��ü�¼.����Ա����%Type,
    �Ӱ��־_IN            ���˷��ü�¼.�Ӱ��־%Type,
    ����ʱ��_IN            ���˷��ü�¼.�Ǽ�ʱ��%Type,
    ���㷽ʽ_IN            ����Ԥ����¼.���㷽ʽ%Type,
    ����ID_IN            Ʊ��ʹ����ϸ.����id%Type
) As
    Cursor c_PreCard Is
        Select Id As ����ID From ���˷��ü�¼ 
        Where ��¼���� = 5 And ʵ��Ʊ��=ԭ����_IN And ����id = ����ID_IN;
    r_CardRow c_PreCard%Rowtype;
    

    v_����id    ���˷��ü�¼.Id%Type;
    v_����id    ���˷��ü�¼.����id%Type;
    v_�ջ�ID    Ʊ�ݴ�ӡ����.ID%Type;
    v_��ӡID    Ʊ�ݴ�ӡ����.ID%Type;

    Err_NoPreCard Exception;
Begin
    If Not ���㷽ʽ_IN Is Null Then
        Select ���˽��ʼ�¼_Id.Nextval Into v_����id From Dual;
    End If;

    If ��������_IN <> 2 Then
        --���￨���ü�¼
        Select ���˷��ü�¼_Id.Nextval Into v_����id From Dual;

        Insert Into ���˷��ü�¼
            (Id, ��¼����, ��¼״̬, No,ʵ��Ʊ��,���,����id, ��ҳid, ���˲���id, ���˿���id, ��ʶ��, ����, �Ա�, ����, �ѱ�,
             ���ʷ���, �����־, �Ӱ��־, ��������id,������, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, �շ�ϸĿid,
             �շ����, ���㵥λ, ����, ����, ��ҩ����,���ӱ�־, ִ�в���id, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��,
             ʵ�ս��, ����id, ���ʽ��)
        Values
            (v_����id,5,1,���ݺ�_IN,����_IN,1,����ID_IN, Decode(��ҳID_IN, 0, Null, ��ҳID_IN),
             Decode(���˲���ID_IN, 0, Null, ���˲���ID_IN), Decode(���˿���ID_IN, 0, Null, ���˿���ID_IN),
             Decode(��ʶ��_IN, 0, Null, ��ʶ��_IN), ����_IN, �Ա�_IN, ����_IN, �ѱ�_IN, Decode(���㷽ʽ_IN, Null, 1, 0), 3,
             �Ӱ��־_IN, ��������ID_IN, ����Ա����_IN, ����Ա���_IN, ����Ա����_IN, ����ʱ��_IN, ����ʱ��_IN, �շ�ϸĿID_IN,
             �շ����_IN, ���㵥λ_IN, 1, 1, ����_IN, ��������_IN, ִ�в���ID_IN, ������ĿID_IN, �վݷ�Ŀ_IN, ���_IN, ���_IN,
             ���_IN, v_����id, Decode(���㷽ʽ_IN, Null, Null, ���_IN));
    
        --��������վ��￨���ã��򽫽������벡��Ԥ����¼
        If Not ���㷽ʽ_IN Is Null Then
            Insert Into ����Ԥ����¼
                (Id, No, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id,
                 ժҪ)
            Values
                (����Ԥ����¼_Id.Nextval, ���ݺ�_IN, 5, 1, ����ID_IN, Decode(��ҳID_IN, 0, Null, ��ҳID_IN),
                 Decode(���˿���ID_IN, 0, Null, ���˿���ID_IN), ���㷽ʽ_IN, ����ʱ��_IN, ����Ա���_IN, ����Ա����_IN, ���_IN,
                 v_����id, '���￨����');
        End If;
    
        IF Not ����ID_IN Is Null then
            --����ʹ��Ʊ��
            Select Ʊ�ݴ�ӡ����_ID.Nextval Into v_��ӡID From Dual;
            Insert Into Ʊ�ݴ�ӡ����(
                ID,��������,NO)
            Values(
                v_��ӡID,5,���ݺ�_IN);

            Insert Into Ʊ��ʹ����ϸ(
                ID,Ʊ��,����,����,ԭ��,����id,��ӡid,ʹ��ʱ��,ʹ����)
            Values(
                Ʊ��ʹ����ϸ_ID.Nextval,5,����_IN,1,1,����ID_IN,v_��ӡID,����ʱ��_IN,����Ա����_IN);
        
            --��������״̬�仯
            Update Ʊ�����ü�¼
                Set ��ǰ����=����_IN,
                    ʣ������=Decode(Sign(ʣ������-1),-1,0,ʣ������-1)
            Where Id=Nvl(����ID_IN,0);
        End IF;
    
        --��ػ��ܱ�Ĵ���
        If ���㷽ʽ_IN Is Null Then
            --����'�������'
            Update ������� Set ������� = Nvl(�������, 0) + ���_IN Where ���� = 1 And ����id = ����ID_IN;
        
            If Sql%Rowcount = 0 Then
                Insert Into ������� (����id, ����, Ԥ�����, �������) Values (����ID_IN, 1, 0, ���_IN);
            End If;
        
            Delete From ������� Where ����id = ����ID_IN And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
        
            --����'����δ�����'
            Update ����δ�����
            Set ��� = Nvl(���, 0) + ���_IN
            Where ����id = ����ID_IN And Nvl(��ҳid, 0) = Nvl(��ҳID_IN, 0) And Nvl(���˲���id, 0) = Nvl(���˲���ID_IN, 0) And
                        Nvl(���˿���id, 0) = Nvl(���˿���ID_IN, 0) And Nvl(��������id, 0) = Nvl(��������ID_IN, 0) And
                        Nvl(ִ�в���id, 0) = Nvl(ִ�в���ID_IN, 0) And ������Ŀid+0 = ������ĿID_IN And ��Դ;�� = 3;
        
            If Sql%Rowcount = 0 Then
                Insert Into ����δ�����
                    (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
                Values
                    (����ID_IN, Decode(��ҳID_IN, 0, Null, ��ҳID_IN), Decode(���˲���ID_IN, 0, Null, ���˲���ID_IN),
                     Decode(���˿���ID_IN, 0, Null, ���˿���ID_IN), ��������ID_IN, ִ�в���ID_IN, ������ĿID_IN, 3, ���_IN);
            End If;
        Else
            --����"��Ա�ɿ����"
            Update ��Ա�ɿ����
            Set ��� = Nvl(���, 0) + ���_IN
            Where �տ�Ա = ����Ա����_IN And ���� = 1 And ���㷽ʽ = ���㷽ʽ_IN;
        
            If Sql%Rowcount = 0 Then
                Insert Into ��Ա�ɿ���� (�տ�Ա, ���㷽ʽ, ����, ���) Values (����Ա����_IN, ���㷽ʽ_IN, 1, ���_IN);
            End If;
        
            Delete From ��Ա�ɿ����
            Where ���� = 1 And �տ�Ա = ����Ա����_IN And ���㷽ʽ = ���㷽ʽ_IN And Nvl(���, 0) = 0;
        End If;
    
        --����'���˷��û���'
        Update ���˷��û���
        Set Ӧ�ս�� = Nvl(Ӧ�ս��, 0) + ���_IN, ʵ�ս�� = Nvl(ʵ�ս��, 0) + ���_IN,
                ���ʽ�� = Nvl(���ʽ��, 0) + Decode(���㷽ʽ_IN, Null, 0, ���_IN)
        Where ���� = Trunc(����ʱ��_IN) And Nvl(���˲���id, 0) = Nvl(���˲���ID_IN,0) 
                    And Nvl(���˿���id, 0) = Nvl(���˿���ID_IN,0) 
                    And ��������id = ��������ID_IN And ִ�в���id = ִ�в���ID_IN 
                    And ������Ŀid+0 = ������ĿID_IN And ��Դ;�� = 3 
                    And ���ʷ��� = Decode(���㷽ʽ_IN, Null, 1, 0);
    
        If Sql%Rowcount = 0 Then
            Insert Into ���˷��û���
                (����, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���ʷ���, Ӧ�ս��, ʵ�ս��,
                 ���ʽ��)
            Values
                (Trunc(����ʱ��_IN), Decode(���˲���ID_IN, 0, Null, ���˲���ID_IN),
                 Decode(���˿���ID_IN, 0, Null, ���˿���ID_IN), ��������ID_IN, ִ�в���ID_IN, ������ĿID_IN, 3,
                 Decode(���㷽ʽ_IN, Null, 1, 0), ���_IN, ���_IN, Decode(���㷽ʽ_IN, Null, 0, ���_IN));
        End If;
    Else
        --��������ʽ
        --���Ȳ�����Ҫ������ԭ���￨���ü�¼
        Open c_PreCard;
        Fetch c_PreCard Into r_CardRow;
    
        If c_PreCard%Rowcount = 0 Then
            Close c_PreCard;
            Raise Err_NoPreCard;
        Else
            --������ԭ���ü�¼ʱ�Ŵ���
            --�ش��ջ�Ʊ��
            Begin
                Select Max(ID) Into v_�ջ�ID From Ʊ�ݴ�ӡ���� Where ��������=5 And NO=���ݺ�_IN;
            Exception
                When Others Then NULL;
            End;
            If v_�ջ�ID Is Not Null Then
                Insert Into Ʊ��ʹ����ϸ(
                    ID,Ʊ��,����,����,ԭ��,����id,��ӡid,ʹ��ʱ��,ʹ����)
                Select 
                    Ʊ��ʹ����ϸ_ID.Nextval,Ʊ��,����,2,4,����ID,��ӡID,����ʱ��_IN,����Ա����_IN
                From Ʊ��ʹ����ϸ
                Where ��ӡID=v_�ջ�ID And Ʊ��=5 And ����=1;
            End If;
            
            --�ش򷢳�Ʊ��
            Select Ʊ�ݴ�ӡ����_ID.Nextval Into v_��ӡID From Dual;

            Insert Into Ʊ�ݴ�ӡ����(
                ID,��������,NO)
            Values(
                v_��ӡID,5,���ݺ�_IN);

            Insert Into Ʊ��ʹ����ϸ(
                ID,Ʊ��,����,����,ԭ��,����id,��ӡid,ʹ��ʱ��,ʹ����)
            Values(
                Ʊ��ʹ����ϸ_ID.Nextval,5,����_IN,1,Decode(v_�ջ�ID,NULL,1,3),����ID_IN,v_��ӡID,����ʱ��_IN,����Ա����_IN);
        
            --����״̬�仯
            Update Ʊ�����ü�¼
                Set ��ǰ����=����_IN, 
                    ʣ������=Decode(Sign(ʣ������-1),-1,0,ʣ������-1)
            Where Id=Nvl(����ID_IN,0);

            --����ԭ������¼״̬
            Update ���˷��ü�¼ 
                Set ʵ��Ʊ��=����_IN,
                    ��ҩ����=����_IN,
                    ���ӱ�־=2
            Where Id=r_CardRow.����id;
        
            Close c_PreCard;
        End If;
    End If;

    --���˾��￨��Ϣ�仯
    Update ������Ϣ Set ���￨��=����_IN, ����֤��=����_IN Where ����id=����ID_IN;
Exception
    When Err_NoPreCard Then Raise_Application_Error(-20101, '[ZLSOFT]û�з���ԭ���￨���ż�¼,��������ʧ�ܣ�[ZLSOFT]');
    When Others Then Zl_Errorcenter(Sqlcode, Sqlerrm);
End zl_���￨��¼_Insert;
/

Create Or Replace Procedure zl_���￨��¼_Delete(
    ���ݺ�_In        ���˷��ü�¼.No%Type,
    ����Ա���_In    ���˷��ü�¼.����Ա���%Type,
    ����Ա����_In    ���˷��ü�¼.����Ա����%Type
) As
    Cursor c_Cardinfo Is
        Select a.Id As ����id, Nvl(a.���ʷ���, 0) As ����, a.����id, a.ʵ��Ʊ��, a.����id, Nvl(a.��ҳid, 0) As ��ҳid,
             Nvl(a.���˲���id, 0) As ���˲���id, Nvl(a.���˿���id, 0) As ���˿���id, Nvl(a.��������id, 0) As ��������id,
             Nvl(a.ִ�в���id, 0) As ִ�в���id, a.������Ŀid, a.ʵ�ս��, b.���㷽ʽ, b.��Ԥ��
        From ���˷��ü�¼ a, ����Ԥ����¼ b
        Where a.��¼���� = 5 And a.��¼״̬ = 1 And a.No = ���ݺ�_In And a.����id = b.����id(+);
    r_Cardrow c_Cardinfo%Rowtype;
    
    v_����id    ���˷��ü�¼.Id%Type;
    v_����id    ���˷��ü�¼.����id%Type;
    v_��ӡID    Ʊ�ݴ�ӡ����.ID%Type;

    v_Date Date;
    Err_Custom Exception;
Begin
    Open c_Cardinfo;
    Fetch c_Cardinfo Into r_Cardrow;

    --�����ж�Ҫ�˿��ļ�¼�Ƿ����
    If c_Cardinfo%Rowcount = 0 Then
        Close c_Cardinfo;
        Raise Err_Custom;
    Else
        Select Sysdate Into v_Date From Dual;
        Select ���˷��ü�¼_Id.Nextval Into v_����id From Dual;
    
        If r_Cardrow.���� = 0 Then
            Select ���˽��ʼ�¼_Id.Nextval Into v_����id From Dual;
        End If;
    
        --�˳����￨���ü�¼
        Insert Into ���˷��ü�¼
            (Id, No,ʵ��Ʊ��,��¼����, ��¼״̬, ���, ����id, ��ҳid, ���˲���id, ���˿���id, �����־, ����, �Ա�, ����, �ѱ�,
             �շ����, �շ�ϸĿid, ���㵥λ, ����, ����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���,
             ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ����id,
             ���ʽ��)
            Select v_����id, No,ʵ��Ʊ��,��¼����, 2, ���, ����id, ��ҳid, ���˲���id, ���˿���id, �����־, ����, �Ա�, ����, �ѱ�,
                 �շ����, �շ�ϸĿid, ���㵥λ, ����, -����, �Ӱ��־, ���ӱ�־, ��ҩ����, ������Ŀid, �վݷ�Ŀ, ���ʷ���,
                 ��׼����, -Ӧ�ս��, -ʵ�ս��, ��������id, ������, ִ�в���id, ����Ա���_In, ����Ա����_In, ����ʱ��,
                 v_Date, v_����id, Decode(v_����id, Null, Null, -���ʽ��)
            From ���˷��ü�¼
            Where Id = r_Cardrow.����id;
    
        Update ���˷��ü�¼ Set ��¼״̬ = 3 Where Id = r_Cardrow.����id;
    
        --Ԥ���������յĽ�����
        If r_Cardrow.���� = 0 Then
            Insert Into ����Ԥ����¼(
                Id, No, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��,����id)
            Select ����Ԥ����¼_Id.Nextval, No, ��¼����, 2, ����id, ��ҳid, ����id, ժҪ, 
                ���㷽ʽ, v_Date,����Ա���_In, ����Ա����_In, -��Ԥ��, v_����id
            From ����Ԥ����¼
            Where ��¼���� = 5 And ��¼״̬ = 1 And ����id = r_Cardrow.����id;
        
            Update ����Ԥ����¼ Set ��¼״̬ = 3 Where ��¼���� = 5 And ��¼״̬ = 1 And ����id = r_Cardrow.����id;
        End If;
    
        --�˿��ջ�Ʊ��
        Begin
            Select Max(ID) Into v_��ӡID From Ʊ�ݴ�ӡ���� Where ��������=5 And NO=���ݺ�_IN;
        Exception
            When Others Then NULL;
        End;
        If v_��ӡID Is Not Null Then
            Insert Into Ʊ��ʹ����ϸ(
                ID,Ʊ��,����,����,ԭ��,����id,��ӡid,ʹ��ʱ��,ʹ����)
            Select
                Ʊ��ʹ����ϸ_ID.Nextval,Ʊ��,����,2,2,����id,��ӡID,v_Date,����Ա����_In
            From Ʊ��ʹ����ϸ
            Where ��ӡID=v_��ӡID And Ʊ��=5 And ����=1;
        End If;
    
        --���²�����Ϣ
        Update ������Ϣ Set ���￨��=Null, ����֤��=Null Where ���￨��=r_Cardrow.ʵ��Ʊ��;
    
        --��ػ��ܱ�Ĵ���
        If r_Cardrow.���� = 1 Then
            --����'�������'
            Update �������
                Set ������� = Nvl(�������, 0) + (-1 * r_Cardrow.ʵ�ս��)
            Where ���� = 1 And ����id = r_Cardrow.����id;
        
            If Sql%Rowcount = 0 Then
                Insert Into �������
                    (����id, ����, Ԥ�����, �������)
                Values
                    (r_Cardrow.����id, 1, 0, -1 * r_Cardrow.ʵ�ս��);
            End If;
        
            --����'����δ�����'
            Update ����δ�����
                Set ��� = Nvl(���, 0) + (-1 * r_Cardrow.ʵ�ս��)
            Where ����id = r_Cardrow.����id And Nvl(��ҳid, 0) = r_Cardrow.��ҳid And
                Nvl(���˲���id, 0) = r_Cardrow.���˲���id And Nvl(���˿���id, 0) = r_Cardrow.���˿���id And
                Nvl(��������id, 0) = r_Cardrow.��������id And Nvl(ִ�в���id, 0) = r_Cardrow.ִ�в���id And
                ������Ŀid+0 = r_Cardrow.������Ŀid And ��Դ;�� = 3;
        
            If Sql%Rowcount = 0 Then
                Insert Into ����δ�����
                    (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
                Values
                    (r_Cardrow.����id, Decode(r_Cardrow.��ҳid, 0, Null, r_Cardrow.��ҳid),
                     Decode(r_Cardrow.���˲���id, 0, Null, r_Cardrow.���˲���id),
                     Decode(r_Cardrow.���˿���id, 0, Null, r_Cardrow.���˿���id),
                     Decode(r_Cardrow.��������id, 0, Null, r_Cardrow.��������id),
                     Decode(r_Cardrow.ִ�в���id, 0, Null, r_Cardrow.ִ�в���id), r_Cardrow.������Ŀid, 3, -1 * r_Cardrow.ʵ�ս��);
            End If;
        Elsif r_Cardrow.���㷽ʽ Is Not Null Then
            --����"��Ա�ɿ����"
            Update ��Ա�ɿ����
                Set ��� = Nvl(���, 0) + (-1 * r_Cardrow.��Ԥ��)
            Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = r_Cardrow.���㷽ʽ;
        
            If Sql%Rowcount = 0 Then
                Insert Into ��Ա�ɿ����
                    (�տ�Ա, ���㷽ʽ, ����, ���)
                Values
                    (����Ա����_In, r_Cardrow.���㷽ʽ, 1, -1 * r_Cardrow.��Ԥ��);
            End If;
        
            Delete From ��Ա�ɿ����
            Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = r_Cardrow.���㷽ʽ And Nvl(���, 0) = 0;
        End If;
    
        --����'���˷��û���'
        Update ���˷��û���
            Set Ӧ�ս�� = Nvl(Ӧ�ս��, 0) + (-1 * r_Cardrow.ʵ�ս��), ʵ�ս�� = Nvl(ʵ�ս��, 0) + (-1 * r_Cardrow.ʵ�ս��),
                ���ʽ�� = Nvl(���ʽ��, 0) + Decode(r_Cardrow.����, 0, -1 * r_Cardrow.ʵ�ս��, Null)
        Where ���� = Trunc(v_Date) And Nvl(���˲���id, 0) = r_Cardrow.���˲���id And
                    Nvl(���˿���id, 0) = r_Cardrow.���˿���id And Nvl(��������id, 0) = r_Cardrow.��������id And
                    Nvl(ִ�в���id, 0) = r_Cardrow.ִ�в���id And ������Ŀid+0 = r_Cardrow.������Ŀid And ��Դ;�� = 3 And
                    ���ʷ��� = r_Cardrow.����;
    
        If Sql%Rowcount = 0 Then
            Insert Into ���˷��û���
                (����, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���ʷ���, Ӧ�ս��, ʵ�ս��,
                 ���ʽ��)
            Values
                (Trunc(v_Date), Decode(r_Cardrow.���˲���id, 0, Null, r_Cardrow.���˲���id),
                 Decode(r_Cardrow.���˿���id, 0, Null, r_Cardrow.���˿���id),
                 Decode(r_Cardrow.��������id, 0, Null, r_Cardrow.��������id),
                 Decode(r_Cardrow.ִ�в���id, 0, Null, r_Cardrow.ִ�в���id), r_Cardrow.������Ŀid, 3, r_Cardrow.����,
                 -1 * r_Cardrow.ʵ�ս��, -1 * r_Cardrow.ʵ�ս��, Decode(r_Cardrow.����, 0, -1 * r_Cardrow.ʵ�ս��, Null));
        End If;
    
        Close c_Cardinfo;
    End If;
Exception
    When Err_Custom Then Raise_Application_Error(-20999, 'û�з���Ҫ�˿��ļ�¼,�ü�¼�����Ѿ��˳���');
    When Others Then Zl_ErrorCenter(Sqlcode, Sqlerrm);
End zl_���￨��¼_Delete;
/

-------------------------------------------------------
--ģ�飺������ʼ�¼.SQL
Create Or Replace Procedure zl_������ʼ�¼_Delete(
    NO_IN            ���˷��ü�¼.NO%Type,
    ���_IN          Varchar2,
    ����Ա���_IN    ���˷��ü�¼.����Ա���%Type,
    ����Ա����_IN    ���˷��ü�¼.����Ա����%Type
)
AS
	--���ܣ�����һ��������ʵ�����ָ�������
	--��ţ���ʽ��"1,3,5,7,8",Ϊ�ձ�ʾ�������пɳ�����
    --�ù����������ָ��������

    --���α�ΪҪ�˷ѵ��ݵ�����ԭʼ��¼
    Cursor c_Bill is
        Select * From ���˷��ü�¼
        Where NO=NO_IN And ��¼����=2 And ��¼״̬ IN(0,1,3) And �����־=1
        Order by �շ�ϸĿID,���;

    --���α����ڴ���ҩƷ����������
    --��Ҫ�ܷ��õ�ִ��״̬,��Ϊ���ڴ˲�����
    Cursor c_Stock is
        Select * From ҩƷ�շ���¼
        Where NO=NO_IN And ���� IN(9,25) And Mod(��¼״̬,3)=1 And ����� IS NULL
            And ����ID IN(
                Select ID From ���˷��ü�¼ 
                Where NO=NO_IN And ��¼����=2 And ��¼״̬ IN(0,1,3) 
                    And �շ���� IN('4','5','6','7') And �����־=1
                    And (INSTR(','||���_IN||',',','||���||',')>0 Or ���_IN Is Null)
                )
        Order BY ҩƷID;
    
    --���α����ڴ���δ��ҩƷ��¼
    Cursor c_Spare is
        Select * From δ��ҩƷ��¼ Where NO=NO_IN And ���� IN(9,25);

    --���α����ڴ�����ü�¼���
    Cursor c_Serial is
        Select ���,�۸񸸺� From ���˷��ü�¼ Where NO=NO_IN And ��¼����=2 And ��¼״̬ IN(0,1,3) Order BY ���;

	v_ҽ��ID		����ҽ����¼.ID%Type;
	v_����			Number;
	v_����			���˷��ü�¼.�۸񸸺�%Type;

    --�����˷Ѽ������
    v_ʣ������		Number;
    v_ʣ��Ӧ��		Number;
    v_ʣ��ʵ��		Number;
    v_ʣ��ͳ��		Number;

    v_׼������		Number;
    v_�˷Ѵ���		Number;

    v_Ӧ�ս��		Number;
    v_ʵ�ս��		Number;
    v_ͳ����		Number;

    v_Dec			Number;
	
    v_Count			Number;
    v_CurDate		Date;
    Err_Custom		Exception;
    v_Error			Varchar2(255);
Begin
    --�Ƿ��Ѿ�ȫ����ȫִ��(ֻ�����ŵ��ݵļ��)
    Select Nvl(Count(*),0) Into v_Count 
    From ���˷��ü�¼ 
    Where NO=NO_IN And ��¼����=2 And ��¼״̬ IN(0,1,3) And Nvl(ִ��״̬,0)<>1;
    IF v_Count = 0 Then
        v_Error := '�õ����е���Ŀ�Ѿ�ȫ����ȫִ�У�';
        Raise Err_Custom;
    End IF;

    --δ��ȫִ�е���Ŀ�Ƿ���ʣ������(ֻ�����ŵ��ݵļ��)
    Select Nvl(Count(*),0) Into v_Count
    From (
        Select ���,Sum(����) as ʣ������
        From (
            Select ��¼״̬,Nvl(�۸񸸺�,���) as ���,
                Avg(Nvl(����,1)*����) as ���� 
            From ���˷��ü�¼
            Where NO=NO_IN And ��¼����=2 And �����־=1
                And Nvl(�۸񸸺�,���) IN (
                        Select Nvl(�۸񸸺�,���) 
                        From ���˷��ü�¼ 
                        Where NO=NO_IN And ��¼����=2 And �����־=1
                            And ��¼״̬ IN(0,1,3) And Nvl(ִ��״̬,0)<>1)
            Group by ��¼״̬,Nvl(�۸񸸺�,���)
            )
        Group by ��� Having Sum(����)<>0);
    IF v_Count = 0 Then
        v_Error := '�õ�����δ��ȫִ�в�����Ŀʣ������Ϊ��,û�п������ʵķ��ã�';
        Raise Err_Custom;
    End IF;
    
    ---------------------------------------------------------------------------------
    --���ñ���
    Select Sysdate Into v_CurDate From Dual;

    --���С��λ��
    Begin
        Select To_Number(����ֵ) Into v_Dec From ϵͳ������ Where ������=9;
    Exception
        When Others Then v_Dec:=2;
    End;

    --ѭ������ÿ�з���(������Ŀ��)
	For r_Bill IN c_Bill Loop
		IF INSTR(','||���_IN||',',','||Nvl(r_Bill.�۸񸸺�,r_Bill.���)||',') >0 Or ���_IN Is Null Then
			Select Decode(��¼״̬,0,1,0) Into v_���� From ���˷��ü�¼ Where ID=r_Bill.ID;
			If v_����=0 Then
				IF Nvl(r_Bill.ִ��״̬,0)<>1 Then
					--��ʣ������,ʣ��Ӧ��,ʣ��ʵ��
					Select 
						Sum(Nvl(����,1)*����),Sum(Ӧ�ս��),Sum(ʵ�ս��),Sum(ͳ����)
						Into v_ʣ������,v_ʣ��Ӧ��,v_ʣ��ʵ��,v_ʣ��ͳ��
					From ���˷��ü�¼ 
					Where NO=NO_IN And ��¼����=2 And ���=r_Bill.���;

					IF v_ʣ������=0 Then
						IF ���_IN IS Not NULL Then 
							v_Error := '�����е�'||Nvl(r_Bill.�۸񸸺�,r_Bill.���)||'�з����Ѿ�ȫ�����ʣ�';
							Raise Err_Custom;
						End IF;
						--�����δ�޶��к�,ԭʼ�����еĸñ��Ѿ�ȫ������(ִ��״̬=0��һ�ֿ���)
					Else
						--׼������(��ҩƷ��ĿΪʣ������,ԭʼ����)
						IF Instr(',4,5,6,7,',r_Bill.�շ����)=0 Then
							v_׼������:=v_ʣ������;
						Else
							Select Sum(Nvl(����,1)*ʵ������) Into v_׼������
							From ҩƷ�շ���¼
							Where NO=NO_IN And ���� IN(9,25) And MOD(��¼״̬,3)=1 
								And ����� is NULL And ����ID=r_Bill.ID;

							--���������õ���������
							If r_Bill.�շ����='4' And Nvl(v_׼������,0)=0 Then
								v_׼������:=v_ʣ������;
							End IF;
						End if;

						--�����˷��ü�¼
						
						--�ñ���Ŀ�ڼ�������
						Select Nvl(Max(Abs(ִ��״̬)),0)+1 Into v_�˷Ѵ���
						From ���˷��ü�¼ 
						Where NO=NO_IN And ��¼����=2 And ��¼״̬=2 And ���=r_Bill.���;
						
						--���=ʣ����*(׼����/ʣ����)
						v_Ӧ�ս��:=Round(v_ʣ��Ӧ��*(v_׼������/v_ʣ������),v_Dec);
						v_ʵ�ս��:=Round(v_ʣ��ʵ��*(v_׼������/v_ʣ������),v_Dec);
						v_ͳ����:=Round(v_ʣ��ͳ��*(v_׼������/v_ʣ������),v_Dec);

						--�����˷Ѽ�¼
						Insert Into ���˷��ü�¼(
							ID,NO,��¼����,��¼״̬,���,��������,�۸񸸺�,����ID,ҽ�����,�����־,�ಡ�˵�,Ӥ����,����,
							�Ա�,����,��ʶ��,����,�ѱ�,���˲���ID,���˿���ID,�շ����,�շ�ϸĿID,���㵥λ,����,��ҩ����,
							����,�Ӱ��־,���ӱ�־,������ĿID,�վݷ�Ŀ,���ʷ���,��׼����,Ӧ�ս��,ʵ�ս��,��������ID,
							������,ִ�в���ID,������,ִ����,ִ��״̬,ִ��ʱ��,����Ա���,����Ա����,����ʱ��,�Ǽ�ʱ��,
							������Ŀ��,���մ���ID,ͳ����,���ʵ�ID,ժҪ)
						Select ���˷��ü�¼_ID.Nextval,NO,��¼����,2,���,��������,�۸񸸺�,����ID,ҽ�����,�����־,�ಡ�˵�,
							Ӥ����,����,�Ա�,����,��ʶ��,����,�ѱ�,���˲���ID,���˿���ID,�շ����,�շ�ϸĿID,���㵥λ,
							Decode(Sign(v_׼������-Nvl(����,1)*����),0,����,1),��ҩ����,
							Decode(Sign(v_׼������-Nvl(����,1)*����),0,-1*����,-1*v_׼������),�Ӱ��־,���ӱ�־,
							������ĿID,�վݷ�Ŀ,���ʷ���,��׼����,-1*v_Ӧ�ս��,-1*v_ʵ�ս��,��������ID,������,ִ�в���ID,
							������,ִ����,-1*v_�˷Ѵ���,ִ��ʱ��,����Ա���_IN,����Ա����_IN,����ʱ��,v_CurDate,
							������Ŀ��,���մ���ID,-1*v_ͳ����,���ʵ�ID,ժҪ
						From ���˷��ü�¼ Where ID=r_Bill.ID;

						--��¼����ҽ�����Ѷ�Ӧ��ҽ��ID(����������)
						If v_ҽ��ID IS Null And r_Bill.ҽ����� IS Not Null Then
							v_ҽ��ID:=r_Bill.ҽ�����;
						End IF;

						--�������
						Update �������
							Set �������=Nvl(�������,0) - v_ʵ�ս��
						 Where ����ID=r_Bill.����ID And ����=1;
						IF SQL%RowCount=0 Then
							Insert Into �������(
								����ID,����,�������,Ԥ�����)
							Values(
								r_Bill.����ID,1,-1*v_ʵ�ս��,0);
						End IF;
						
						--����δ�����
						Update ����δ�����
							Set ���=Nvl(���,0) - v_ʵ�ս��
						 Where ����ID=r_Bill.����ID
							And Nvl(��ҳID,0)=Nvl(r_Bill.��ҳID,0)
							And Nvl(���˲���ID,0)=Nvl(r_Bill.���˲���ID,0)
							And Nvl(���˿���ID,0)=Nvl(r_Bill.���˿���ID,0)
							And Nvl(��������ID,0)=Nvl(r_Bill.��������ID,0)
							And Nvl(ִ�в���ID,0)=Nvl(r_Bill.ִ�в���ID,0)
							And ������ĿID+0=r_Bill.������ĿID And ��Դ;��+0=1;
						IF SQL%RowCount=0 Then
							Insert Into ����δ�����(
								����ID,��ҳID,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,��Դ;��,���)
							Values(
								r_Bill.����ID,r_Bill.��ҳID,r_Bill.���˲���ID,r_Bill.���˿���ID,
								r_Bill.��������ID,r_Bill.ִ�в���ID,r_Bill.������ĿID,1,-1*v_ʵ�ս��);
						End IF;

						--�����˷��û���
						Update ���˷��û���
							Set Ӧ�ս��=Nvl(Ӧ�ս��,0) - v_Ӧ�ս��,
								ʵ�ս��=Nvl(ʵ�ս��,0) - v_ʵ�ս��
						 Where ����=Trunc(v_CurDate)
							And Nvl(���˲���ID,0)=Nvl(r_Bill.���˲���ID,0)
							And Nvl(���˿���ID,0)=Nvl(r_Bill.���˿���ID,0)
							And Nvl(��������ID,0)=Nvl(r_Bill.��������ID,0)
							And Nvl(ִ�в���ID,0)=Nvl(r_Bill.ִ�в���ID,0)
							And ������ĿID+0=r_Bill.������ĿID
							And ��Դ;��=1 And ���ʷ���=1;
						IF SQL%RowCount=0 Then
							Insert Into ���˷��û���(
								����,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,��Դ;��,���ʷ���,Ӧ�ս��,ʵ�ս��,���ʽ��)
							Values(
								Trunc(v_CurDate),r_Bill.���˲���ID,r_Bill.���˿���ID,r_Bill.��������ID,r_Bill.ִ�в���ID,
								r_Bill.������ĿID,1,1,-1 * v_Ӧ�ս��,-1 * v_ʵ�ս��,0);
						End IF;
						
						--���ԭ���ü�¼
						--ִ��״̬:ȫ������(׼����=ʣ����)���Ϊ0,������Ϊ1
						Update ���˷��ü�¼ 
							Set ��¼״̬=3,
								ִ��״̬=Decode(Sign(v_׼������-v_ʣ������),0,0,1) 
						Where ID=r_Bill.ID;
					End IF;
				Else
					IF ���_IN Is Not Null Then
						v_Error := '�����е�'||Nvl(r_Bill.�۸񸸺�,r_Bill.���)||'�з����Ѿ���ȫִ��,�������ʣ�';
						Raise Err_Custom;
					End IF;
					--���:û�޶��к�,ԭʼ�����а����Ѿ���ȫִ�е�
				End IF;
			End IF;
		End IF;
	End Loop;
    
    ---------------------------------------------------------------------------------
    --ҩƷ�������
    For r_Stock in c_Stock Loop
        --����ҩƷ���
        If r_Stock.�ⷿID IS Not Null then
            Update ҩƷ���
                Set ��������=Nvl(��������,0)+Nvl(r_Stock.����,1)*Nvl(r_Stock.ʵ������,0)
             Where �ⷿID=r_Stock.�ⷿID And ҩƷID=r_Stock.ҩƷID
                And Nvl(����,0)=Nvl(r_Stock.����,0) And ����=1;
            IF SQL%RowCount=0 Then
                Insert Into ҩƷ���(
                    �ⷿID,ҩƷID,����,����,Ч��,��������,�ϴ�����,�ϴβ���,���Ч��)
                Values(
                    r_Stock.�ⷿID,r_Stock.ҩƷID,1,r_Stock.����,r_Stock.Ч��,
                    Nvl(r_Stock.����,1)*Nvl(r_Stock.ʵ������,0),r_Stock.����,r_Stock.����,r_Stock.���Ч��);
            End IF;
        End IF;

        --ɾ��ҩƷ�շ���¼
        Delete From ҩƷ�շ���¼ Where ID=r_Stock.ID;
    End Loop;

    --δ��ҩƷ��¼
    For r_Spare IN c_Spare Loop
        Select Nvl(Count(*),0) Into v_Count
        From ҩƷ�շ���¼ 
        Where NO=NO_IN And ����=r_Spare.���� And Mod(��¼״̬,3)=1 
            And ����� is NULL And Nvl(�ⷿID,0)=Nvl(r_Spare.�ⷿID,0);
        If v_Count=0 Then
            Delete From δ��ҩƷ��¼ Where ����=r_Spare.���� And NO=NO_IN And Nvl(�ⷿID,0)=Nvl(r_Spare.�ⷿID,0);
        End IF;
    End Loop;
	
	---------------------------------------------------------------------------------
	--����ǻ���,ֱ��ɾ�����ü�¼(ҩƷ�����)
	v_Count:=0;
	For r_Bill IN c_Bill Loop
		IF INSTR(','||���_IN||',',','||Nvl(r_Bill.�۸񸸺�,r_Bill.���)||',') >0 Or ���_IN Is Null Then
			Select Decode(��¼״̬,0,1,0) Into v_���� From ���˷��ü�¼ Where ID=r_Bill.ID;
			If v_����=1 Then
				IF Nvl(r_Bill.ִ��״̬,0)<>1 Then
					Delete From ���˷��ü�¼ Where ID=r_Bill.ID;
					v_Count:=v_Count+1;--��¼�Ƿ���ɾ����

					--��¼����ҽ�����Ѷ�Ӧ��ҽ��ID(����������)
					If v_ҽ��ID IS Null And r_Bill.ҽ����� IS Not Null Then
						v_ҽ��ID:=r_Bill.ҽ�����;
					End IF;
				Else
					IF ���_IN Is Not Null Then
						v_Error := '�����е�'||Nvl(r_Bill.�۸񸸺�,r_Bill.���)||'�з����Ѿ���ȫִ��,�������ʣ�';
						Raise Err_Custom;
					End IF;
					--���:û�޶��к�,ԭʼ�����а����Ѿ���ȫִ�е�
				End IF;
			End IF;
		End IF;
	End Loop;

	--ɾ��֮����ͳһ�������
	If v_Count>0 Then
		v_Count:=1;
		For r_Serial In c_Serial Loop
			If r_Serial.�۸񸸺� IS NULL Then 
				v_����:=v_Count;
			End IF;

			Update ���˷��ü�¼ 
				Set ���=v_Count,
					�۸񸸺�=Decode(�۸񸸺�,NULL,NULL,v_����)
			Where NO=NO_IN And ��¼����=2 And ���=r_Serial.���;
			
			Update ���˷��ü�¼
				Set ��������=v_Count
			Where NO=NO_IN And ��¼����=2 And ��������=r_Serial.���;

			v_Count:=v_Count+1;
		End Loop;
	End IF;

	--���ŵ���ȫ������ʱ��ɾ������ҽ������
	If ���_IN IS NULL And v_ҽ��ID IS Not NULL Then
		Select Nvl(Count(*),0) Into v_Count
		From (
			Select ���,Sum(����) as ʣ������
			From (
				Select ��¼״̬,Nvl(�۸񸸺�,���) as ���,
					Avg(Nvl(����,1)*����) as ���� 
				From ���˷��ü�¼
				Where NO=NO_IN And ��¼����=2 And ҽ�����+0=v_ҽ��ID
				Group by ��¼״̬,Nvl(�۸񸸺�,���)
				)
			Group by ��� Having Nvl(Sum(����),0)<>0);
		IF v_Count = 0 Then
			Delete From ����ҽ������ Where ҽ��ID=v_ҽ��ID And ��¼����=2 And NO=NO_IN;
		End IF;
	End IF;
Exception
    When Err_Custom Then Raise_Application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When Others Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_������ʼ�¼_Delete;
/

CREATE OR REPLACE PROCEDURE zl_������ʼ�¼_Verify (
    NO_IN			���˷��ü�¼.NO%TYPE,
    ����Ա���_IN   ���˷��ü�¼.����Ա���%TYPE,
    ����Ա����_IN   ���˷��ü�¼.����Ա����%TYPE,
	���_IN			Varchar2:=NULL,
	���ʱ��_IN		���˷��ü�¼.�Ǽ�ʱ��%Type:=NULL
) AS
--���ܣ����һ��������ʻ��۵�
--������
--		���_IN����ʽ��"1,3,5,7,8",Ϊ�ձ�ʾ�������δ��˵���
--		���ʱ��_IN�����ڲ�����Ҫͳһ���ƻ򷵻�ʱ��ĵط�
	--ֻ��ȡָ����ŵ�,δ��˵Ĳ��ݽ��д���
	Cursor c_Bill is
		Select * From ���˷��ü�¼ 
		Where ��¼����=2 And ��¼״̬=0 And NO=NO_IN
			And (Instr(','||���_IN||',',','||Nvl(�۸񸸺�,���)||',')>0 Or ���_IN Is Null)
		Order BY ���;

	--����а����������õ�δ������ʱ�����ݲ��������Ƿ��Զ�����  25-���ʵ���������
	Cursor c_Stuff is
		Select NO,����,�ⷿID From δ��ҩƷ��¼
		Where NO=NO_IN And ����=25 And �ⷿID IS Not Null
			And Exists(Select ����ֵ From ϵͳ������ Where ������=92 And ����ֵ='1')
			And Exists(
				Select A.��� From ���˷��ü�¼ A,�������� B
				Where A.��¼����=2 And A.��¼״̬=1 And A.NO=NO_IN
					And (Instr(','||���_IN||',',','||Nvl(A.�۸񸸺�,A.���)||',')>0 Or ���_IN Is Null)					
					And A.�շ�ϸĿID=B.����ID And B.��������=1
				)
		Order BY �ⷿID;

	v_Date	Date;
BEGIN
	If ���ʱ��_IN IS Null Then
		Select Sysdate Into v_Date From Dual;
	Else
		v_Date:=���ʱ��_IN;
	End IF;

	For r_Bill IN c_Bill Loop
		Update ���˷��ü�¼
			Set ��¼״̬=1,
				����Ա���=����Ա���_IN,
				����Ա����=����Ա����_IN,
				�Ǽ�ʱ��=v_Date --�Ѳ�����ҩƷ��¼��ʱ�䲻��
		Where ID=r_Bill.ID;

		--ҩƷ�շ���¼.��������
		Update ҩƷ�շ���¼
			Set ��������=Decode(Sign(Nvl(�������,v_Date)-v_Date),-1,��������,v_Date)  
		Where NO=NO_IN AND ���� IN(9,25)  AND ����ID=r_Bill.ID;

		--�������
		Update �������
			Set �������=Nvl(�������,0)+Nvl(r_Bill.ʵ�ս��,0)
		Where ����ID=r_Bill.����ID And ����=1;

		IF SQL%RowCount=0 Then
			Insert Into �������(
				����ID,����,�������,Ԥ�����)
			Values(
				r_Bill.����ID,1,r_Bill.ʵ�ս��,0);
		End IF;

		--����δ�����
		Update ����δ�����
			Set ���=Nvl(���,0)+Nvl(r_Bill.ʵ�ս��,0)
		 Where ����ID=r_Bill.����ID
			And Nvl(��ҳID,0)=Nvl(r_Bill.��ҳID,0)
			And Nvl(���˲���ID,0)=Nvl(r_Bill.���˲���ID,0)
			And Nvl(���˿���ID,0)=Nvl(r_Bill.���˿���ID,0)
			And Nvl(��������ID,0)=Nvl(r_Bill.��������ID,0)
			And Nvl(ִ�в���ID,0)=Nvl(r_Bill.ִ�в���ID,0)
			And ������ĿID+0=r_Bill.������ĿID
			And ��Դ;��+0=r_Bill.�����־;

		IF SQL%RowCount=0 Then
			Insert Into ����δ�����(
				����ID,��ҳID,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,��Դ;��,���)
			Values(
				r_Bill.����ID,r_Bill.��ҳID,r_Bill.���˲���ID,r_Bill.���˿���ID,r_Bill.��������ID,r_Bill.ִ�в���ID,r_Bill.������ĿID,r_Bill.�����־,Nvl(r_Bill.ʵ�ս��,0));
		End IF;

		--���˷��û���
		Update ���˷��û���
			Set Ӧ�ս��=Nvl(Ӧ�ս��,0)+Nvl(r_Bill.Ӧ�ս��,0),
				ʵ�ս��=Nvl(ʵ�ս��,0)+Nvl(r_Bill.ʵ�ս��,0)
		 Where ����=Trunc(v_Date)
			And Nvl(���˲���ID,0)=Nvl(r_Bill.���˲���ID,0)
			And Nvl(���˿���ID,0)=Nvl(r_Bill.���˿���ID,0)
			And Nvl(��������ID,0)=Nvl(r_Bill.��������ID,0)
			And Nvl(ִ�в���ID,0)=Nvl(r_Bill.ִ�в���ID,0)
			And ������ĿID+0=r_Bill.������ĿID
			And ��Դ;��=r_Bill.�����־ And ���ʷ���=1;

		IF SQL%RowCount=0 Then
			Insert Into ���˷��û���(
				����,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,
				��Դ;��,���ʷ���,Ӧ�ս��,ʵ�ս��,���ʽ��)
			Values(
				Trunc(v_Date),r_Bill.���˲���ID,r_Bill.���˿���ID,r_Bill.��������ID,
				r_Bill.ִ�в���ID,r_Bill.������ĿID,r_Bill.�����־,1,r_Bill.Ӧ�ս��,r_Bill.ʵ�ս��,0);
		End IF;
	End Loop;

	--�ⷿ�е�ҩƷ��ȫ��������Ϊ���շ�
	Update δ��ҩƷ��¼ Set ���շ�=1,��������=v_Date
	Where NO=NO_IN And ����=9 And Nvl(���շ�,0)=0 
		And Nvl(�ⷿID,0) Not IN(
			Select Distinct Nvl(ִ�в���ID,0) From ���˷��ü�¼ 
				Where ��¼����=2 And NO=NO_IN And �շ���� IN('5','6','7') And ��¼״̬=0);

	Update δ��ҩƷ��¼ Set ���շ�=1,��������=v_Date
	Where NO=NO_IN And ����=25 And Nvl(���շ�,0)=0 
		And Nvl(�ⷿID,0) Not IN(
			Select Distinct Nvl(ִ�в���ID,0) From ���˷��ü�¼ 
				Where ��¼����=2 And NO=NO_IN And �շ����='4' And ��¼״̬=0);

	--���������Զ�����
	For r_Stuff In c_Stuff Loop
		zl_�����շ���¼_��������(r_Stuff.�ⷿID,r_Stuff.����,r_Stuff.NO,����Ա����_IN,����Ա����_IN,����Ա����_IN,1,Sysdate);
	End Loop;
EXCEPTION
    WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_������ʼ�¼_Verify;
/

Create Or Replace Procedure zl_������ʼ�¼_Insert(
    NO_IN				���˷��ü�¼.NO%Type,
    ���_IN				���˷��ü�¼.���%Type,
    ����ID_IN			���˷��ü�¼.����ID%Type,
    ��ʶ��_IN			���˷��ü�¼.��ʶ��%Type,
    ����_IN				���˷��ü�¼.����%Type,
    �Ա�_IN				���˷��ü�¼.�Ա�%Type,
    ����_IN				���˷��ü�¼.����%Type,
    �ѱ�_IN				���˷��ü�¼.�ѱ�%Type,
    �Ӱ��־_IN			���˷��ü�¼.�Ӱ��־%Type,
    Ӥ����_IN			���˷��ü�¼.Ӥ����%Type,
	���˲���ID_IN		���˷��ü�¼.���˲���ID%Type,
	���˿���ID_IN		���˷��ü�¼.���˿���ID%Type,
    ��������ID_IN		���˷��ü�¼.��������ID%Type,
    ������_IN			���˷��ü�¼.������%Type,
    ��������_IN			���˷��ü�¼.��������%Type,
    �շ�ϸĿID_IN		���˷��ü�¼.�շ�ϸĿID%Type,
    �շ����_IN			���˷��ü�¼.�շ����%Type,
    ���㵥λ_IN			���˷��ü�¼.���㵥λ%Type,
    ����_IN				���˷��ü�¼.����%Type,
    ����_IN				���˷��ü�¼.����%Type,
    ���ӱ�־_IN			���˷��ü�¼.���ӱ�־%Type,
    ִ�в���ID_IN		���˷��ü�¼.ִ�в���ID%Type,
    �۸񸸺�_IN			���˷��ü�¼.�۸񸸺�%Type,
    ������ĿID_IN		���˷��ü�¼.������ĿID%Type,
    �վݷ�Ŀ_IN			���˷��ü�¼.�վݷ�Ŀ%Type,
    ��׼����_IN			���˷��ü�¼.��׼����%Type,
    Ӧ�ս��_IN			���˷��ü�¼.Ӧ�ս��%Type,
    ʵ�ս��_IN			���˷��ü�¼.ʵ�ս��%Type,
    ����ʱ��_IN			���˷��ü�¼.����ʱ��%Type,
    �Ǽ�ʱ��_IN			���˷��ü�¼.�Ǽ�ʱ��%Type,
    ҩƷժҪ_IN			ҩƷ�շ���¼.ժҪ%Type,
    ����_IN				Number,
    ����Ա���_IN		���˷��ü�¼.����Ա���%Type,
    ����Ա����_IN		���˷��ü�¼.����Ա����%Type,
    ���ID_IN			ҩƷ��������.���ID%Type:=Null,
    ���ʵ�ID_IN			���˷��ü�¼.���ʵ�ID%Type:=Null,
    ����ժҪ_IN			���˷��ü�¼.ժҪ%Type:=Null,
    ҽ�����_IN			���˷��ü�¼.ҽ�����%TYPE:=NULL,
    Ƶ��_IN				ҩƷ�շ���¼.Ƶ��%Type:=NULL,
    ����_IN				ҩƷ�շ���¼.����%Type:=NULL,
    �÷�_IN				ҩƷ�շ���¼.�÷�%Type:=NULL,--�÷�[|�巨]
    ��Ч_IN				ҩƷ�շ���¼.����%Type:=NULL,
    �Ƽ�����_IN			ҩƷ�շ���¼.����%Type:=NULL
)
AS
    --���ܣ�����һ��������ʵ���
    --������
    --   ҩƷժҪ_IN:�޸ı����µ���ʱ�á�Ŀǰ�����ڴ����ҩƷ�շ���¼��ժҪ�С�
    --         ԭ����(��¼״̬=2)��¼�޸Ĳ������µ��ݺš�
    --         �µ���(��¼״̬=1)��¼���޸ĵ�ԭ���ݺš�
    v_����ID ���˷��ü�¼.ID%Type;
    v_���ȼ� δ��ҩƷ��¼.���ȼ�%Type;

    --ҩ��������ʱ��ҩƷ--
    ------------------------------------------------------------
    --���α����ڷ���ҩƷ�����ֽ�
    Cursor c_Stock is
        Select * From ҩƷ��� 
        Where ҩƷID=�շ�ϸĿID_IN And �ⷿID=ִ�в���ID_IN
            And ����=1 And(Nvl(����,0)=0 Or Ч�� is Null Or Ч��>Trunc(Sysdate))
            And Nvl(��������,0)<>0
        Order By Nvl(����,0);
    r_Stock c_Stock%RowType;
    
    --����
    v_����			ҩƷ���.ҩ������%Type;
    v_ʱ��			�շ���ĿĿ¼.�Ƿ���%Type;
    v_����			�շ���ĿĿ¼.����%Type;
    --��ʱ����
    v_������		Number;
    v_��ǰ����		Number;
    v_�ܽ��		Number;
    v_��ǰ����		Number;
    --ҩƷ�շ���¼
    v_����			ҩƷ�շ���¼.����%Type;
    v_����			ҩƷ�շ���¼.����%Type;
    v_����			ҩƷ�շ���¼.����%Type;
    v_Ч��			ҩƷ�շ���¼.Ч��%Type;
    v_���			ҩƷ�շ���¼.���%Type;
    v_����			ҩƷ�շ���¼.����%Type;
	v_���Ч��		ҩƷ�շ���¼.���Ч��%Type;
	v_�������		ҩƷ�շ���¼.�������%Type;
    ------------------------------------------------------------
	v_�÷�			ҩƷ�շ���¼.�÷�%Type;
	v_�巨			ҩƷ�շ���¼.���%Type;

    v_Dec			Number;
	v_Count			Number;
	v_Error			Varchar2(255);
    Err_Custom		Exception;
Begin
	--ҩƷ�÷��巨�ֽ�
	IF �÷�_IN IS Not NULL Then
		IF Instr(�÷�_IN,'|')>0 Then
			v_�÷�:=Substr(�÷�_IN,1,Instr(�÷�_IN,'|')-1);
			v_�巨:=Substr(�÷�_IN,Instr(�÷�_IN,'|')+1);
		Else
			v_�÷�:=�÷�_IN;
		End IF;
	End IF;

    --���С��λ��
    Begin
        Select To_Number(����ֵ) Into v_Dec From ϵͳ������ Where ������=9;
    Exception
        When Others Then v_Dec:=2;
    End;

    --���˷��ü�¼
    Select ���˷��ü�¼_ID.Nextval Into v_����ID From Dual;

    Insert Into ���˷��ü�¼(
        ID,��¼����,NO,��¼״̬,���,��������,�۸񸸺�,�����־,����ID,��ʶ��,
        ����,�Ա�,����,���˲���ID,���˿���ID,�ѱ�,�շ����,�շ�ϸĿID,���㵥λ,
        ����,����,�Ӱ��־,���ӱ�־,������ĿID,�վݷ�Ŀ,��׼����,Ӧ�ս��,ʵ�ս��,
        ���ʷ���,������,��������ID,������,����ʱ��,�Ǽ�ʱ��,ִ�в���ID,ִ��״̬,
        ����Ա���,����Ա����,Ӥ����,���ʵ�ID,ժҪ,ҽ�����)
    Values(
        v_����ID,2,NO_IN,Decode(����_IN,1,0,1),���_IN,Decode(��������_IN,0,Null,��������_IN),
        Decode(�۸񸸺�_IN,0,Null,�۸񸸺�_IN),1,����ID_IN,
        Decode(��ʶ��_IN,0,Null,��ʶ��_IN),����_IN,�Ա�_IN,����_IN,���˲���ID_IN,
        ���˿���ID_IN,�ѱ�_IN,�շ����_IN,�շ�ϸĿID_IN,���㵥λ_IN,����_IN,����_IN,
        �Ӱ��־_IN,���ӱ�־_IN,������ĿID_IN,�վݷ�Ŀ_IN,��׼����_IN,Ӧ�ս��_IN,
        ʵ�ս��_IN,1,����Ա����_IN,��������ID_IN,������_IN,����ʱ��_IN,�Ǽ�ʱ��_IN,
        ִ�в���ID_IN,0,Decode(����_IN,1,Null,����Ա���_IN),
        Decode(����_IN,1,Null,����Ա����_IN),Ӥ����_IN,���ʵ�ID_IN,����ժҪ_IN,ҽ�����_IN);

    --��ػ��ܱ�Ĵ���
	If Nvl(����_IN,0)=0 Then
		--�������
		Update �������
			Set �������=Nvl(�������,0)+ʵ�ս��_IN
		Where ����ID=����ID_IN And ����=1;

		IF SQL%RowCount=0 Then
			Insert Into �������(
				����ID,����,�������,Ԥ�����)
			Values(
				����ID_IN,1,ʵ�ս��_IN,0);
		End IF;

		--����δ�����
		Update ����δ�����
			Set ���=Nvl(���,0)+ʵ�ս��_IN
		 Where ����ID=����ID_IN
			And Nvl(��ҳID,0)=0
			And Nvl(���˲���ID,0)=Nvl(���˲���ID_IN,0)
			And Nvl(���˿���ID,0)=Nvl(���˿���ID_IN,0)
			And Nvl(��������ID,0)=Nvl(��������ID_IN,0)
			And Nvl(ִ�в���ID,0)=Nvl(ִ�в���ID_IN,0)
			And ������ĿID+0=������ĿID_IN
			And ��Դ;��+0=1;

		IF SQL%RowCount=0 Then
			Insert Into ����δ�����(
				����ID,��ҳID,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,��Դ;��,���)
			Values(
				����ID_IN,Null,���˲���ID_IN,���˿���ID_IN,��������ID_IN,ִ�в���ID_IN,������ĿID_IN,1,ʵ�ս��_IN);
		End IF;

		--���˷��û���
		Update ���˷��û���
			Set Ӧ�ս��=Nvl(Ӧ�ս��,0)+Ӧ�ս��_IN,
				ʵ�ս��=Nvl(ʵ�ս��,0)+ʵ�ս��_IN
		 Where ����=Trunc(�Ǽ�ʱ��_IN)
			And Nvl(���˲���ID,0)=Nvl(���˲���ID_IN,0)
			And Nvl(���˿���ID,0)=Nvl(���˿���ID_IN,0)
			And Nvl(��������ID,0)=Nvl(��������ID_IN,0)
			And Nvl(ִ�в���ID,0)=Nvl(ִ�в���ID_IN,0)
			And ������ĿID+0=������ĿID_IN
			And ��Դ;��=1 And ���ʷ���=1;

		IF SQL%RowCount=0 Then
			Insert Into ���˷��û���(
				����,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,
				��Դ;��,���ʷ���,Ӧ�ս��,ʵ�ս��,���ʽ��)
			Values(
				Trunc(�Ǽ�ʱ��_IN),���˲���ID_IN,���˿���ID_IN,��������ID_IN,
				ִ�в���ID_IN,������ĿID_IN,1,1,Ӧ�ս��_IN,ʵ�ս��_IN,0);
		End IF;
	End IF;

    --ҩƷ���������ϲ���
	v_Count:=0;
	If �շ����_IN='4' Then--�������õ����ĲŴ���
		Select �������� Into v_Count From �������� Where ����ID=�շ�ϸĿID_IN;
	End IF;
    IF �շ����_IN in('5','6','7') Or (�շ����_IN='4' And Nvl(v_Count,0)=1) Then
		If �շ����_IN='4' Then
			Select Nvl(A.���÷���,0),Nvl(B.�Ƿ���,0),B.���� 
				Into v_����,v_ʱ��,v_����
			From �������� A,�շ���ĿĿ¼ B
			Where A.����ID=B.ID And B.ID=�շ�ϸĿID_IN;
		Else
			Select Nvl(A.ҩ������,0),Nvl(B.�Ƿ���,0),B.���� 
				Into v_����,v_ʱ��,v_����
			From ҩƷ��� A,�շ���ĿĿ¼ B
			Where A.ҩƷID=B.ID And B.ID=�շ�ϸĿID_IN;
		End IF;

        v_������:=����_IN*����_IN;
        v_�ܽ��:=0;
        Open c_Stock;

        While v_������<>0 Loop
            Fetch c_Stock Into r_Stock;
            IF c_Stock%NotFound Then
                --��һ�ξ�û�п��,������ʱ�۶�������
                --����ҩƷ�����ֽⲻ��,Ҳ���ǿ�治�㡣
                IF v_����=1 Or v_ʱ��=1 Then
                    Close c_Stock;
					If ҽ�����_IN IS NULL Then
						If �շ����_IN='4' Then
							v_Error:='�� '||���_IN||' �еķ�����ʱ����������"'||v_����||'"û���㹻�Ŀ�棡';
						Else
							v_Error:='�� '||���_IN||' �еķ�����ʱ��ҩƷ"'||v_����||'"û���㹻�Ŀ�棡';
						End IF;
					Else
						If �շ����_IN='4' Then
							v_Error:='�ڴ�����"'||����_IN||'"ʱ���ַ�����ʱ����������"'||v_����||'"û���㹻�Ŀ�棡';
						Else
							v_Error:='�ڴ�����"'||����_IN||'"ʱ���ַ�����ʱ��ҩƷ"'||v_����||'"û���㹻�Ŀ�棡';
						End IF;
					End IF;
                    Raise Err_Custom;
                End IF;
            ElsIF(v_����=1 And Nvl(r_Stock.����,0)=0) Or(v_����=0 And Nvl(r_Stock.����,0)<>0) Then 
                Close c_Stock;
                If ҽ�����_IN IS NULL Then
					If �շ����_IN='4' Then
						v_Error:='�� '||���_IN||' ����������"'||v_����||'"�ķ������������¼�����,����������ݵ���ȷ�ԣ�';
					Else
	                    v_Error:='�� '||���_IN||' ��ҩƷ"'||v_����||'"�ķ������������¼�����,����ҩƷ���ݵ���ȷ�ԣ�';
					End IF;
                Else
					If �շ����_IN='4' Then
						v_Error:='�ڴ�����"'||����_IN||'"ʱ������������"'||v_����||'"�ķ������������¼�����,����������ݵ���ȷ�ԣ�';
					Else
	                    v_Error:='�ڴ�����"'||����_IN||'"ʱ����ҩƷ"'||v_����||'"�ķ������������¼�����,����ҩƷ���ݵ���ȷ�ԣ�';
					End IF;
                End IF;
                Raise Err_Custom;
            End IF;

            --ȷ�����ηֽ�����
            IF v_����=1 Or v_ʱ��=1 Then
                --���ڲ�������ʱ��ֻ���ֽܷ�һ��,�ֽⲻ�������ж���.���ֽ���Ϊ�˼��㵥��.
                --ÿ�ηֽ�ȡС��,��治���ֽⲻ���������ж�.
                IF v_������<=Nvl(r_Stock.��������,0) Then
                    v_��ǰ����:=v_������;
                Else
                    v_��ǰ����:=Nvl(r_Stock.��������,0);
                End if;
                IF v_ʱ��=1 Then 
                    If r_Stock.ʵ������=0 Then
                        v_��ǰ����:=0;
                    Else
                        v_��ǰ����:=Round(Nvl(r_Stock.ʵ�ʽ��/r_Stock.ʵ������,0),5);
                    End IF;
                ElsIf v_����=1 Then
                    v_��ǰ����:=��׼����_IN;
                End IF;
            Else
                --��ͨҩƷ
                --���ܹ�����,�������Ѹ��ݲ����ж�
                v_��ǰ����:=v_������;
                v_��ǰ����:=��׼����_IN;
            End IF;

            --ҩƷ���(��ͨ�������û�м�¼)
            IF c_Stock%Found Then
                Update ҩƷ���
                    Set ��������=Nvl(��������,0)-v_��ǰ����
                Where �ⷿID=ִ�в���ID_IN And ҩƷID=�շ�ϸĿID_IN
                    And Nvl(����,0)=Nvl(r_Stock.����,0) And ����=1;
            ElsIf ִ�в���ID_IN IS Not NULL Then
                --ֻ�в�������ʱ��ҩƷ���ܿ�治�����
                Update ҩƷ���
                    Set ��������=Nvl(��������,0)-v_��ǰ����
                Where �ⷿID=ִ�в���ID_IN And ҩƷID=�շ�ϸĿID_IN
                    And Nvl(����,0)=0 And ����=1;
                IF SQL%RowCount=0 Then
                    Insert Into ҩƷ���(
                        �ⷿID,ҩƷID,����,��������)
                    Values(
                        ִ�в���ID_IN,�շ�ϸĿID_IN,1,-1*v_��ǰ����);
                End IF;
            End IF;

            --ҩƷ�շ���¼
			v_����:=Null;v_����:=Null;
			v_Ч��:=Null;v_����:=Null;
			v_���Ч��:=Null;v_�������:=Null;
            IF c_Stock%Found Then
                v_����:=r_Stock.����;
                v_����:=r_Stock.�ϴ�����;
                v_Ч��:=r_Stock.Ч��;
                v_����:=r_Stock.�ϴβ���;

				--�������Ч��:һ���Բ�������Ч��
				IF �շ����_IN='4' Then
					v_Count:=0;
					Begin
						Select ���Ч�� Into v_Count From �������� Where Nvl(һ���Բ���,0)=1 And ����ID=�շ�ϸĿID_IN;
					Exception
						When Others Then Null;
					End;
					IF Nvl(v_Count,0)>0 Then
						v_���Ч��:=r_Stock.���Ч��;	
						v_�������:=v_���Ч��-v_Count*30;
					End IF;
				End IF;
            End IF;

            Select Nvl(Max(���),0)+1 Into v_��� From ҩƷ�շ���¼ 
				Where ����=Decode(�շ����_IN,'4',25,9) And ��¼״̬=1 And NO=NO_IN;

            --�޸ĵ�ԭ���ݺŴ����ժҪ��
			v_����:=NULL;
            If ��Ч_IN IS Not NULL Or �Ƽ�����_IN IS Not NULL THEN 
                v_����:=Nvl(��Ч_IN,0)||Nvl(�Ƽ�����_IN,0);
            End IF;
            Insert Into ҩƷ�շ���¼(
                ID,��¼״̬,����,NO,���,�ⷿID,�Է�����ID,������ID,���ϵ��,
                ҩƷID,����,����,����,Ч��,����,��д����,ʵ������,���ۼ�,���۽��,
                ժҪ,������,��������,����ID,Ƶ��,����,�÷�,���,����,���Ч��,�������)
            Values(
                ҩƷ�շ���¼_ID.Nextval,1,Decode(�շ����_IN,'4',25,9),NO_IN,v_���,ִ�в���ID_IN,��������ID_IN,
                ���ID_IN,-1,�շ�ϸĿID_IN,v_����,v_����,v_����,v_Ч��,Decode(v_����,1,1,����_IN),
                Decode(v_����,1,v_��ǰ����,v_��ǰ����/����_IN),Decode(v_����,1,v_��ǰ����,v_��ǰ����/����_IN),
                v_��ǰ����,Round(v_��ǰ����*v_��ǰ����,v_Dec),ҩƷժҪ_IN,����Ա����_IN,�Ǽ�ʱ��_IN,
                v_����ID,Ƶ��_IN,����_IN,v_�÷�,v_�巨,v_����,v_���Ч��,v_�������);

            --δ��ҩƷ��¼
            Update δ��ҩƷ��¼
                Set ����ID=����ID_IN,����=����_IN
             Where ����=Decode(�շ����_IN,'4',25,9) And NO=NO_IN 
				And Nvl(�ⷿID,0)=Nvl(ִ�в���ID_IN,0);

            IF SQL%RowCount=0 Then
                --ȡ������ȼ�
                Begin
                    Select B.���ȼ� Into v_���ȼ� From ������Ϣ A,��� B
                     Where A.���=B.����(+) And A.����ID=����ID_IN;
                Exception
                    When Others Then Null;
                End;

                Insert Into δ��ҩƷ��¼(
                    ����,NO,����ID,����,���ȼ�,�Է�����ID,�ⷿID,��������,���շ�,��ӡ״̬)
                Values(
                    Decode(�շ����_IN,'4',25,9),NO_IN,����ID_IN,����_IN,v_���ȼ�,
					��������ID_IN,ִ�в���ID_IN,�Ǽ�ʱ��_IN,Decode(����_IN,1,0,1),0);
            End IF;

            v_������:=v_������-v_��ǰ����;
            v_�ܽ��:=v_�ܽ��+Round(v_��ǰ����*v_��ǰ����,v_Dec);
        End Loop;
        
        --���ܷ���ʱ��ҩƷ�ֽ�����α���
        IF v_ʱ��=1 Then
            IF Round(v_�ܽ��/(����_IN*����_IN),5)<>��׼����_IN Then 
                Close c_Stock;    
                If ҽ�����_IN IS NULL Then
					If �շ����_IN='4' Then
						v_Error:='�� '||���_IN||' �е�ʱ����������"'||v_����||'"��ǰ���㵥�۲�һ��,�����������������㣡';
					Else
	                    v_Error:='�� '||���_IN||' �е�ʱ��ҩƷ"'||v_����||'"��ǰ���㵥�۲�һ��,�����������������㣡';
					End IF;
                Else
					If �շ����_IN='4' Then
	                    v_Error:='�ڴ�����"'||����_IN||'"ʱ����ʱ����������"'||v_����||'"��ǰ����ĵ��۷����仯��'||CHR(13)||CHR(10)||'����ò����Ƿ�ͬʱʹ����������ͬ��"'||v_����||'"��';
					Else
						v_Error:='�ڴ�����"'||����_IN||'"ʱ����ʱ��ҩƷ"'||v_����||'"��ǰ����ĵ��۷����仯��'||CHR(13)||CHR(10)||'����ò����Ƿ�ͬʱʹ����������ͬ��"'||v_����||'"��';
					End IF;
                End IF;
                Raise Err_Custom;
            End IF;
        End IF;
        
        Close c_Stock;
    End IF;
Exception
    When Err_Custom Then Raise_Application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When Others Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_������ʼ�¼_Insert;
/

-------------------------------------------------------
--ģ�飺�����շѼ�¼.SQL
Create Or Replace Procedure zl_�����շѼ�¼_Insert(
    NO_IN				���˷��ü�¼.NO%Type,
    ���_IN             ���˷��ü�¼.���%Type,
    ����ID_IN           ���˷��ü�¼.����ID%Type,
    ��ҳID_IN           ���˷��ü�¼.��ҳID%Type,
    ��ʶ��_IN           ���˷��ü�¼.��ʶ��%Type,
    ����_IN             ���˷��ü�¼.����%Type,
    ����_IN             ���˷��ü�¼.����%Type,
    �Ա�_IN             ���˷��ü�¼.�Ա�%Type,
    ����_IN             ���˷��ü�¼.����%Type,
    �ѱ�_IN             ���˷��ü�¼.�ѱ�%Type,
    �Ӱ��־_IN         ���˷��ü�¼.�Ӱ��־%Type,
    ���˲���ID_IN       ���˷��ü�¼.���˲���ID%Type,
    ���˿���ID_IN       ���˷��ü�¼.���˿���ID%Type,
    ��������ID_IN       ���˷��ü�¼.��������ID%Type,
    ������_IN           ���˷��ü�¼.������%Type,
    ��������_IN         ���˷��ü�¼.��������%Type,
    �շ�ϸĿID_IN       ���˷��ü�¼.�շ�ϸĿID%Type,
    �շ����_IN         ���˷��ü�¼.�շ����%Type,
    ���㵥λ_IN         ���˷��ü�¼.���㵥λ%Type,
    ������Ŀ��_IN       ���˷��ü�¼.������Ŀ��%Type,
    ���մ���ID_IN       ���˷��ü�¼.���մ���ID%Type,
    ��ҩ����_IN         ���˷��ü�¼.��ҩ����%Type,
    ����_IN             ���˷��ü�¼.����%Type,
    ����_IN             ���˷��ü�¼.����%Type,
    ���ӱ�־_IN         ���˷��ü�¼.���ӱ�־%Type,
    ִ�в���ID_IN       ���˷��ü�¼.ִ�в���ID%Type,
    �۸񸸺�_IN         ���˷��ü�¼.�۸񸸺�%Type,
    ������ĿID_IN       ���˷��ü�¼.������ĿID%Type,
    �վݷ�Ŀ_IN         ���˷��ü�¼.�վݷ�Ŀ%Type,
    ��׼����_IN         ���˷��ü�¼.��׼����%Type,
    Ӧ�ս��_IN         ���˷��ü�¼.Ӧ�ս��%Type,
    ʵ�ս��_IN         ���˷��ü�¼.ʵ�ս��%Type,
    ͳ����_IN         ���˷��ü�¼.ͳ����%Type,
    ����ʱ��_IN         ���˷��ü�¼.����ʱ��%Type,
    �Ǽ�ʱ��_IN         ���˷��ü�¼.�Ǽ�ʱ��%Type,
    ԭNO_IN             ���˷��ü�¼.NO%Type,
    ����ID_IN           ���˷��ü�¼.����ID%Type,
	�շѽ���_IN			Varchar2,
    ��Ԥ����_IN         ����Ԥ����¼.��Ԥ��%Type,
    ���ս���_IN         Varchar2,
    ����Ա���_IN       ���˷��ü�¼.����Ա���%Type,
    ����Ա����_IN       ���˷��ü�¼.����Ա����%Type,
    ���ID_IN           ҩƷ��������.���ID%Type:=Null,
    ժҪ_IN             ���˷��ü�¼.ժҪ%Type:=Null,
    �Ƿ���_IN         ���˷��ü�¼.�Ƿ���%Type:=0,
    �÷�_IN                 ҩƷ�շ���¼.�÷�%Type:=NULL--�÷�[|�巨]
)
AS
    --���ܣ�����һ�������շѵ���
    --������
    --  ��ҳID_IN:סԺ�����շ�ʱ�á�
    --  ԭNO_IN:�޸ı����µ���ʱ�á�Ŀǰ���ڴ����ҩƷ�շ���¼��ժҪ�С�
    --         ԭ����(��¼״̬=2)��¼�޸Ĳ������µ��ݺš�
    --         �µ���(��¼״̬=1)��¼���޸ĵ�ԭ���ݺš�
	--	�շѽ���_IN:��ʽ="���㷽ʽ|������|�������||.....",ע���޽������Ҫ�ÿո����
	--	���ս���_IN:��ʽ="���㷽ʽ|������||....."

    --���α������շѳ�Ԥ���Ŀ���Ԥ���б�(��SQL�ο�סԺ����)
    --��ID�������ȳ��ϴ�δ����ġ�
	--���������㷽ʽΪ���տ����Ԥ���
    Cursor c_Deposit(v_����ID ������Ϣ.����ID%Type) is
    Select * From(
        Select A.ID,A.��¼״̬,A.NO,Nvl(A.���,0) as ���
        From ����Ԥ����¼ A,(
                Select NO,Sum(Nvl(A.���,0)) as ��� 
                From ����Ԥ����¼ A
				Where A.����ID is Null And Nvl(A.���,0)<>0 
					And A.����ID=v_����ID
				Group by NO Having Sum(Nvl(A.���,0))<>0
                ) B
        Where A.����ID is Null And Nvl(A.���,0)<>0 
			And A.���㷽ʽ Not IN(Select ���� From ���㷽ʽ Where ����=5)
			And A.NO=B.NO And A.����ID=v_����ID
        Union All
        Select 0 as ID,��¼״̬,NO,Sum(Nvl(���,0)-Nvl(��Ԥ��,0)) as ���
        From ����Ԥ����¼
        Where ��¼���� IN(1,11) And ����ID is Not NULL 
			And Nvl(���,0)<>Nvl(��Ԥ��,0) And ����ID=v_����ID
        Having Sum(Nvl(���,0)-Nvl(��Ԥ��,0))<>0
        Group by ��¼״̬,NO)
    Order by ID,NO;

    v_����ID		���˷��ü�¼.ID%Type;
    v_���ȼ�		δ��ҩƷ��¼.���ȼ�%Type;
    v_Ԥ�����		����Ԥ����¼.��Ԥ��%Type;

    --ҩ��������ʱ��ҩƷ--
    ------------------------------------------------------------
    --���α����ڷ���ҩƷ�����ֽ�
    Cursor c_Stock is
        Select * From ҩƷ��� 
        Where ҩƷID=�շ�ϸĿID_IN And �ⷿID=ִ�в���ID_IN
            And ����=1 And(Nvl(����,0)=0 Or Ч�� is Null Or Ч��>Trunc(Sysdate))
            And Nvl(��������,0)<>0
        Order By Nvl(����,0);
    r_Stock c_Stock%RowType;
    
    --����
    v_����			ҩƷ���.ҩ������%Type;
    v_ʱ��			�շ���ĿĿ¼.�Ƿ���%Type;
    v_����			�շ���ĿĿ¼.����%Type;
    --��ʱ����
    v_������		Number;
    v_��ǰ����		Number;
    v_�ܽ��		Number;
    v_��ǰ����		Number;
    --ҩƷ�շ���¼
    v_����			ҩƷ�շ���¼.����%Type;
    v_����			ҩƷ�շ���¼.����%Type;
    v_����			ҩƷ�շ���¼.����%Type;
    v_Ч��			ҩƷ�շ���¼.Ч��%Type;
    v_���			ҩƷ�շ���¼.���%Type;
    v_���Ч��		ҩƷ�շ���¼.���Ч��%Type;
    v_�������		ҩƷ�շ���¼.�������%Type;
    v_�巨			ҩƷ�շ���¼.���%Type;
    ------------------------------------------------------------
	--���㷽ʽ��
	v_��������	Varchar2(500);
	v_��ǰ����	Varchar2(50);
	v_���㷽ʽ	����Ԥ����¼.���㷽ʽ%Type;
	v_������	����Ԥ����¼.��Ԥ��%Type;
	v_�������	����Ԥ����¼.�������%Type;

    v_Dec			Number;

    --��ʱ����
    v_Count			Number;
    Err_Custom		Exception;
    v_Error			Varchar2(255);
Begin
   	--ҩƷ�÷��巨�ֽ�
	IF �÷�_IN IS Not NULL Then
		IF Instr(�÷�_IN,'|')>0 Then			
			v_�巨:=Substr(�÷�_IN,Instr(�÷�_IN,'|')+1);		
		End IF;
	End IF;

    --���С��λ��
    Begin
        Select To_Number(����ֵ) Into v_Dec From ϵͳ������ Where ������=9;
    Exception
        When Others Then v_Dec:=2;
    End;

    --���˷��ü�¼
    Select ���˷��ü�¼_ID.Nextval Into v_����ID From Dual;
    Insert Into ���˷��ü�¼(
        ID,��¼����,NO,��¼״̬,���,��������,�۸񸸺�,�����־,����ID,��ҳID,��ʶ��,����,����,�Ա�,
        ����,���˲���ID,���˿���ID,�ѱ�,�շ����,�շ�ϸĿID,���㵥λ,������Ŀ��,���մ���ID,����,����,��ҩ����,�Ӱ��־,���ӱ�־,
        ������ĿID,�վݷ�Ŀ,��׼����,Ӧ�ս��,ʵ�ս��,ͳ����,���ʷ���,������,��������ID,������,����ʱ��,�Ǽ�ʱ��,
        ִ�в���ID,ִ��״̬,����ID,���ʽ��,����Ա���,����Ա����,ժҪ,�Ƿ���)
    Values(
        v_����ID,1,NO_IN,1,���_IN,Decode(��������_IN,0,Null,��������_IN),Decode(�۸񸸺�_IN,0,Null,�۸񸸺�_IN),Decode(��ҳID_IN,NULL,1,2),
        Decode(����ID_IN,0,Null,����ID_IN),��ҳID_IN,Decode(��ʶ��_IN,0,Null,��ʶ��_IN),����_IN,����_IN,�Ա�_IN,����_IN,���˲���ID_IN,
        ���˿���ID_IN,�ѱ�_IN,�շ����_IN,�շ�ϸĿID_IN,���㵥λ_IN,������Ŀ��_IN,���մ���ID_IN,����_IN,����_IN,��ҩ����_IN,�Ӱ��־_IN,
        ���ӱ�־_IN,������ĿID_IN,�վݷ�Ŀ_IN,��׼����_IN,Ӧ�ս��_IN,ʵ�ս��_IN,ͳ����_IN,0,����Ա����_IN,��������ID_IN,������_IN,
        ����ʱ��_IN,�Ǽ�ʱ��_IN,ִ�в���ID_IN,0,����ID_IN,ʵ�ս��_IN,����Ա���_IN,����Ա����_IN,ժҪ_IN,�Ƿ���_IN);

    IF ���_IN=1 Then
        --����Ԥ����¼(��һ��ʱ����)
		--��������
		IF �շѽ���_IN is Not Null Then
			--�����շѽ���	
			v_��������:=�շѽ���_IN||'||';
			While v_�������� IS Not NULL Loop
				v_��ǰ����:=Substr(v_��������,1,Instr(v_��������,'||')-1);
				v_���㷽ʽ:=Substr(v_��ǰ����,1,Instr(v_��ǰ����,'|')-1);
				v_��ǰ����:=Substr(v_��ǰ����,Instr(v_��ǰ����,'|')+1);
				v_������:=To_Number(Substr(v_��ǰ����,1,Instr(v_��ǰ����,'|')-1));
				v_�������:=LTrim(Substr(v_��ǰ����,Instr(v_��ǰ����,'|')+1));
				
				If Nvl(v_������,0)<>0 Then
					Insert Into ����Ԥ����¼(
						ID,��¼����,NO,��¼״̬,����ID,��ҳID,ժҪ,���㷽ʽ,�������,�տ�ʱ��,����Ա���,����Ա����,��Ԥ��,����ID)
					Values(
						����Ԥ����¼_ID.Nextval,3,NO_IN,1,Decode(����ID_IN,0,Null,����ID_IN),��ҳID_IN,'�շѽ���',
						v_���㷽ʽ,v_�������,�Ǽ�ʱ��_IN,����Ա���_IN,����Ա����_IN,v_������,����ID_IN);
				End IF;

				v_��������:=Substr(v_��������,Instr(v_��������,'||')+2);
			End Loop;
		End IF;

       --�������ս���    
	IF ���ս���_IN IS NOT NULL Then 
		v_��������:=���ս���_IN||'||';
		While v_�������� IS Not NULL Loop
			v_��ǰ����:=Substr(v_��������,1,Instr(v_��������,'||')-1);

			v_���㷽ʽ:=Substr(v_��ǰ����,1,Instr(v_��ǰ����,'|')-1);
			v_������:=To_Number(Substr(v_��ǰ����,Instr(v_��ǰ����,'|')+1));
					
			If Nvl(v_������,0)<>0 Then
				Insert Into ����Ԥ����¼(
					ID,��¼����,NO,��¼״̬,����ID,ժҪ,���㷽ʽ,�տ�ʱ��,����Ա���,����Ա����,��Ԥ��,����ID)
				Values(
					����Ԥ����¼_ID.Nextval,3,NO_IN,1,Decode(����ID_IN,0,Null,����ID_IN),'���ս���',
					v_���㷽ʽ,�Ǽ�ʱ��_IN,����Ա���_IN,����Ա����_IN,v_������,����ID_IN);
			End IF;
			v_��������:=Substr(v_��������,Instr(v_��������,'||')+2);
		End Loop;        
	END if;

        --Ԥ������
        IF Nvl(��Ԥ����_IN,0)<>0 Then
            IF Nvl(����ID_IN,0)=0 Then
                v_Error:='����ȷ�����˲���ID,�շ�ʹ��Ԥ�������ʧ�ܣ�';
                Raise Err_Custom;
            End if;

            v_Ԥ�����:=��Ԥ����_IN;
            For r_Deposit IN c_Deposit(����ID_IN) Loop
                IF r_Deposit.ID<>0 Then
                    --��һ�γ�Ԥ��
                    Update ����Ԥ����¼ 
                        Set ��Ԥ��=Decode(Sign(r_Deposit.���-v_Ԥ�����),-1,r_Deposit.���,v_Ԥ�����),
                            ����ID=����ID_IN
                    Where ID=r_Deposit.ID;
                Else
                    --���ϴ�ʣ���
                    INSERT Into ����Ԥ����¼(
                        ID,NO,ʵ��Ʊ��,��¼����,��¼״̬,����ID,��ҳID,����ID,���,
                        ���㷽ʽ,�������,ժҪ,�ɿλ,��λ������,��λ�ʺ�,�տ�ʱ��,
                        ����Ա����,����Ա���,��Ԥ��,����ID)
                    Select ����Ԥ����¼_ID.Nextval,NO,ʵ��Ʊ��,11,��¼״̬,����ID,
                         ��ҳID,����ID,NULL,���㷽ʽ,�������,ժҪ,�ɿλ,
                         ��λ������,��λ�ʺ�,�տ�ʱ��,����Ա����,����Ա���,
                         Decode(Sign(r_Deposit.���-v_Ԥ�����),-1,r_Deposit.���,v_Ԥ�����),����ID_IN
                    From ����Ԥ����¼
                    Where NO=r_Deposit.NO And ��¼״̬=r_Deposit.��¼״̬
                        And ��¼���� IN(1,11) And RowNum=1;
                End IF;
                --����Ƿ��Ѿ�������
                IF r_Deposit.���<v_Ԥ����� Then
                    v_Ԥ�����:=v_Ԥ�����-r_Deposit.���;
                Else
                    v_Ԥ�����:=0;
                End IF;
                IF v_Ԥ�����=0 Then 
                    Exit;
                End IF;
            End Loop;
            --������Ƿ��㹻
            IF v_Ԥ�����>0 Then
                v_Error:='���˵ĵ�ǰԤ�������� '||Ltrim(To_Char(��Ԥ����_IN,'9999999990.00'))||' ��';
                Raise Err_Custom;
            End IF;

            --���²���Ԥ�����
            Update ������� Set Ԥ�����=Nvl(Ԥ�����,0)-��Ԥ����_IN Where ����ID=����ID_IN And ����=1;
            IF SQL%RowCount=0 Then
                Insert Into �������(����ID,Ԥ�����,����) Values(����ID_IN,-��Ԥ����_IN,1);
            End IF;
            Delete From ������� Where ����ID=����ID_IN And ����=1 And Nvl(�������,0)=0 And Nvl(Ԥ�����,0)=0;
        End IF;
    End IF;

    --��ػ��ܱ�Ĵ���
    --����"��Ա�ɿ����"(ע��Ҫ��������ʻ��Ľ���)
    IF ���_IN=1 Then
		--�����շѽ���	
		IF �շѽ���_IN IS Not NULL Then 
			v_��������:=�շѽ���_IN||'||';
			While v_�������� IS Not NULL Loop
				v_��ǰ����:=Substr(v_��������,1,Instr(v_��������,'||')-1);
				v_���㷽ʽ:=Substr(v_��ǰ����,1,Instr(v_��ǰ����,'|')-1);
				v_��ǰ����:=Substr(v_��ǰ����,Instr(v_��ǰ����,'|')+1);
				v_������:=To_Number(Substr(v_��ǰ����,1,Instr(v_��ǰ����,'|')-1));
				
				If Nvl(v_������,0)<>0 Then
					Update ��Ա�ɿ����
						Set ���=Nvl(���,0)+Nvl(v_������,0)
					 Where �տ�Ա=����Ա����_IN And ����=1
						And ���㷽ʽ=v_���㷽ʽ;

					IF SQL%RowCount=0 Then
						Insert Into ��Ա�ɿ����(
							�տ�Ա,���㷽ʽ,����,���)
						Values(
							����Ա����_IN,v_���㷽ʽ,1,Nvl(v_������,0));
					End IF;
				End IF;

				v_��������:=Substr(v_��������,Instr(v_��������,'||')+2);
			End Loop;
		End IF;

        --�������ս���    
        IF ���ս���_IN IS Not NULL Then 
            v_��������:=���ս���_IN||'||';
            While v_�������� IS Not NULL Loop
                v_��ǰ����:=Substr(v_��������,1,Instr(v_��������,'||')-1);

                v_���㷽ʽ:=Substr(v_��ǰ����,1,Instr(v_��ǰ����,'|')-1);
                v_������:=To_Number(Substr(v_��ǰ����,Instr(v_��ǰ����,'|')+1));
				
				If Nvl(v_������,0)<>0 Then
					Update ��Ա�ɿ����
						Set ���=Nvl(���,0)+Nvl(v_������,0)
					 Where �տ�Ա=����Ա����_IN And ����=1
						And ���㷽ʽ=v_���㷽ʽ;

					IF SQL%RowCount=0 Then
						Insert Into ��Ա�ɿ����(
							�տ�Ա,���㷽ʽ,����,���)
						Values(
							����Ա����_IN,v_���㷽ʽ,1,Nvl(v_������,0));
					End IF;
				End IF;                

                v_��������:=Substr(v_��������,Instr(v_��������,'||')+2);
            End Loop;
        End IF;
        Delete From ��Ա�ɿ���� Where ����=1 And �տ�Ա=����Ա����_IN And Nvl(���,0)=0;
    End IF;

    --���˷��û���
    Update ���˷��û���
        Set Ӧ�ս��=Nvl(Ӧ�ս��,0)+Ӧ�ս��_IN,
            ʵ�ս��=Nvl(ʵ�ս��,0)+ʵ�ս��_IN,
            ���ʽ��=Nvl(���ʽ��,0)+ʵ�ս��_IN
    Where ����=Trunc(�Ǽ�ʱ��_IN)
        And Nvl(���˲���ID,0)=Nvl(���˲���ID_IN,0)
        And Nvl(���˿���ID,0)=Nvl(���˿���ID_IN,0)
        And Nvl(��������ID,0)=Nvl(��������ID_IN,0)
        And Nvl(ִ�в���ID,0)=Nvl(ִ�в���ID_IN,0)
        And ������ĿID+0=������ĿID_IN
        And ��Դ;��=Decode(��ҳID_IN,NULL,1,2) And ���ʷ���=0;

    IF SQL%RowCount=0 Then
        Insert Into ���˷��û���(
            ����,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,��Դ;��,���ʷ���,Ӧ�ս��,ʵ�ս��,���ʽ��)
        Values(
            Trunc(�Ǽ�ʱ��_IN),���˲���ID_IN,���˿���ID_IN,��������ID_IN,ִ�в���ID_IN,������ĿID_IN,
            Decode(��ҳID_IN,NULL,1,2),0,Ӧ�ս��_IN,ʵ�ս��_IN,ʵ�ս��_IN);
    End IF;

    --ҩƷ���������ϲ���
	v_Count:=0;
	If �շ����_IN='4' Then--�������õ����ĲŴ���
		Select �������� Into v_Count From �������� Where ����ID=�շ�ϸĿID_IN;
	End IF;
    IF �շ����_IN in('5','6','7') Or (�շ����_IN='4' And Nvl(v_Count,0)=1) Then
		If �շ����_IN='4' Then
			Select Nvl(A.���÷���,0),Nvl(B.�Ƿ���,0),B.���� 
				Into v_����,v_ʱ��,v_����
			From �������� A,�շ���ĿĿ¼ B
			Where A.����ID=B.ID And B.ID=�շ�ϸĿID_IN;
		Else
			Select Nvl(A.ҩ������,0),Nvl(B.�Ƿ���,0),B.���� 
				Into v_����,v_ʱ��,v_����
			From ҩƷ��� A,�շ���ĿĿ¼ B
			Where A.ҩƷID=B.ID And B.ID=�շ�ϸĿID_IN;
		End IF;

        v_������:=����_IN*����_IN;
        v_�ܽ��:=0;
        Open c_Stock;

        While v_������<>0 Loop
            Fetch c_Stock Into r_Stock;
            IF c_Stock%NotFound Then
                --��һ�ξ�û�п��,������ʱ�۶�������
                --����ҩƷ�����ֽⲻ��,Ҳ���ǿ�治�㡣
                IF v_����=1 Or v_ʱ��=1 Then
                    Close c_Stock;
					If �շ����_IN='4' Then
						v_Error:='�� '||���_IN||' �еķ�����ʱ����������"'||v_����||'"û�п��õĿ�棡';
					Else
	                    v_Error:='�� '||���_IN||' �еķ�����ʱ��ҩƷ"'||v_����||'"û�п��õ�ҩƷ��棡';
					End IF;
                    Raise Err_Custom;
                End IF;
            ElsIF(v_����=1 And Nvl(r_Stock.����,0)=0) Or(v_����=0 And Nvl(r_Stock.����,0)<>0) Then 
                Close c_Stock;
				If �շ����_IN='4' Then
					v_Error:='�� '||���_IN||' ����������"'||v_����||'"�����÷������������¼�����,����������ݵ���ȷ�ԣ�';
				Else
	                v_Error:='�� '||���_IN||' ��ҩƷ"'||v_����||'"�ķ������������¼�����,����ҩƷ���ݵ���ȷ�ԣ�';
				End IF;
                Raise Err_Custom;
            End IF;

            --ȷ�����ηֽ�����
            IF v_����=1 Or v_ʱ��=1 Then
                --���ڲ�������ʱ��ֻ���ֽܷ�һ��,�ֽⲻ�������ж���.���ֽ���Ϊ�˼��㵥��.
                --ÿ�ηֽ�ȡС��,��治���ֽⲻ���������ж�.
                IF v_������<=Nvl(r_Stock.��������,0) Then
                    v_��ǰ����:=v_������;
                Else
                    v_��ǰ����:=Nvl(r_Stock.��������,0);
                End if;
                IF v_ʱ��=1 Then 
                    If r_Stock.ʵ������=0 Then
                        v_��ǰ����:=0;
                    Else
                        v_��ǰ����:=Round(Nvl(r_Stock.ʵ�ʽ��/r_Stock.ʵ������,0),5);
                    End IF;
                ElsIf v_����=1 Then
                    v_��ǰ����:=��׼����_IN;
                End IF;
            Else
                --��ͨҩƷ
                --���ܹ�����,�������Ѹ��ݲ����ж�
                v_��ǰ����:=v_������;
                v_��ǰ����:=��׼����_IN;
            End IF;

            --ҩƷ���(��ͨ�������û�м�¼)
            IF c_Stock%Found Then
                Update ҩƷ���
                    Set ��������=Nvl(��������,0)-v_��ǰ����
                Where �ⷿID=ִ�в���ID_IN And ҩƷID=�շ�ϸĿID_IN
                    And Nvl(����,0)=Nvl(r_Stock.����,0) And ����=1;
            Elsif ִ�в���ID_IN IS Not NULL Then
                --ֻ�в�������ʱ��ҩƷ���ܿ�治�����
                Update ҩƷ���
                    Set ��������=Nvl(��������,0)-v_��ǰ����
                Where �ⷿID=ִ�в���ID_IN And ҩƷID=�շ�ϸĿID_IN
                    And Nvl(����,0)=0 And ����=1;
                IF SQL%RowCount=0 Then
                    Insert Into ҩƷ���(
                        �ⷿID,ҩƷID,����,��������)
                    Values(
                        ִ�в���ID_IN,�շ�ϸĿID_IN,1,-1*v_��ǰ����);
                End IF;
            End IF;

            --ҩƷ�շ���¼
			v_����:=Null;v_����:=Null;
			v_Ч��:=Null;v_����:=Null;
			v_���Ч��:=Null;v_�������:=Null;
            IF c_Stock%Found Then
                v_����:=r_Stock.����;
                v_����:=r_Stock.�ϴ�����;
                v_Ч��:=r_Stock.Ч��;
                v_����:=r_Stock.�ϴβ���;
				
				--�������Ч��:һ���Բ�������Ч��
				IF �շ����_IN='4' Then
					v_Count:=0;
					Begin
						Select ���Ч�� Into v_Count From �������� Where Nvl(һ���Բ���,0)=1 And ����ID=�շ�ϸĿID_IN;
					Exception
						When Others Then Null;
					End;
					IF Nvl(v_Count,0)>0 Then
						v_���Ч��:=r_Stock.���Ч��;	
						v_�������:=v_���Ч��-v_Count*30;
					End IF;
				End IF;
            End IF;

            Select Nvl(Max(���),0)+1 Into v_��� From ҩƷ�շ���¼ 
				Where ����=Decode(�շ����_IN,'4',24,8) And ��¼״̬=1 And NO=NO_IN;

            --�޸ĵ�ԭ���ݺŴ����ժҪ��
			--ע�����ĵ�����ҩƷ���ݲ�ͬ
            Insert Into ҩƷ�շ���¼(
                ID,��¼״̬,����,NO,���,�ⷿID,�Է�����ID,������ID,���ϵ��,
                ҩƷID,����,����,����,Ч��,����,��д����,ʵ������,���ۼ�,���۽��,
                ժҪ,������,��������,����ID,��ҩ����,���Ч��,�������,���)
            Values(
                ҩƷ�շ���¼_ID.Nextval,1,Decode(�շ����_IN,'4',24,8),NO_IN,v_���,ִ�в���ID_IN,��������ID_IN,
                ���ID_IN,-1,�շ�ϸĿID_IN,v_����,v_����,v_����,v_Ч��,Decode(v_����,1,1,����_IN),
                Decode(v_����,1,v_��ǰ����,v_��ǰ����/����_IN),Decode(v_����,1,v_��ǰ����,v_��ǰ����/����_IN),
                v_��ǰ����,Round(v_��ǰ����*v_��ǰ����,v_Dec),ԭNO_IN,����Ա����_IN,�Ǽ�ʱ��_IN,v_����ID,
                ��ҩ����_IN,v_���Ч��,v_�������,v_�巨);

            --δ��ҩƷ��¼:����ͬһ���ⷿ,��һ��ΪҩƷ,һ��Ϊ����,����������¼��
            Update δ��ҩƷ��¼
                Set ����ID=Decode(����ID_IN,0,Null,����ID_IN),
                    ��ҳID=��ҳID_IN,����=����_IN,
					��ҩ����=Nvl(��ҩ����_IN,��ҩ����)--����ҩƷ�Ͳ�����ͬһ���ⷿ,�������޷�ҩ����
             Where ����=Decode(�շ����_IN,'4',24,8) 
				And NO=NO_IN And Nvl(�ⷿID,0)=Nvl(ִ�в���ID_IN,0);

            IF SQL%RowCount=0 Then
                --ȡ������ȼ�
                IF Nvl(����ID_IN,0)<>0 And v_���ȼ� is Null Then
                    Begin
                        Select B.���ȼ� Into v_���ȼ� From ������Ϣ A,��� B
                         Where A.���=B.����(+) And A.����ID=����ID_IN;
                    Exception
                        When Others Then Null;
                    End;
                End IF;

                Insert Into δ��ҩƷ��¼(
                    ����,NO,����ID,��ҳID,����,���ȼ�,�Է�����ID,�ⷿID,��ҩ����,��������,���շ�,��ӡ״̬)
                Values(
                    Decode(�շ����_IN,'4',24,8),NO_IN,Decode(����ID_IN,0,Null,����ID_IN),��ҳID_IN,����_IN,
					v_���ȼ�,��������ID_IN,ִ�в���ID_IN,��ҩ����_IN,�Ǽ�ʱ��_IN,1,0);
            End IF;

            v_������:=v_������-v_��ǰ����;
            v_�ܽ��:=v_�ܽ��+Round(v_��ǰ����*v_��ǰ����,v_Dec);
        End Loop;
        
        --���ܷ���ʱ��ҩƷ�ֽ�����α���
        IF v_ʱ��=1 Then
            IF Round(v_�ܽ��/(����_IN*����_IN),5)<>��׼����_IN Then 
                Close c_Stock;  
				If �շ����_IN='4' Then
					v_Error:='�� '||���_IN||' �е�ʱ����������"'||v_����||'"��ǰ���㵥�۲�һ��,�����������������㣡';
				Else
	                v_Error:='�� '||���_IN||' �е�ʱ��ҩƷ"'||v_����||'"��ǰ���㵥�۲�һ��,�����������������㣡';
				End IF;
                Raise Err_Custom;
            End IF;
        End IF;

        Close c_Stock;
    End IF;

	--���²��ݲ�����Ϣ
	If ���_IN=1 And ����ID_IN IS Not NULL Then
		Update ������Ϣ 
			Set �Ա�=Nvl(�Ա�_IN,�Ա�),
				����=Nvl(����_IN,����)				
		Where ����ID=����ID_IN;
		UPDATE ������Ϣ SET �ѱ�=Nvl(�ѱ�_IN,�ѱ�) WHERE ����ID=����ID_IN AND NOT Exists (SELECT 'X' FROM �ѱ� WHERE ����=�ѱ�_IN AND ����=2);
		If ����_IN IS Not NULL  And ��ҳID_IN IS NULL  Then
			Update ������Ϣ Set ҽ�Ƹ��ʽ=(Select ���� From ҽ�Ƹ��ʽ Where ����=����_IN) Where ����ID=����ID_IN;
		End IF;
	End IF;
Exception
    When Err_Custom Then Raise_Application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When Others Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_�����շѼ�¼_Insert;
/

Create Or Replace Procedure zl_�����շѼ�¼_Insert(
    NO_IN				���˷��ü�¼.NO%Type,
    ����ID_IN			���˷��ü�¼.����ID%Type,
    ��ҳID_IN			���˷��ü�¼.��ҳID%Type,
    ����_IN				���˷��ü�¼.����%Type,
    ����_IN				���˷��ü�¼.����%Type,
    �Ա�_IN				���˷��ü�¼.�Ա�%Type,
    ����_IN				���˷��ü�¼.����%Type,
    ���˲���ID_IN		���˷��ü�¼.���˲���ID%Type,
    ���˿���ID_IN		���˷��ü�¼.���˿���ID%Type,
    ��������ID_IN		���˷��ü�¼.��������ID%Type,
    ������_IN			���˷��ü�¼.������%Type,
    �շѽ���_IN			Varchar2,
    ��Ԥ����_IN			����Ԥ����¼.��Ԥ��%Type,
    ���ս���_IN			Varchar2,
    ����ID_IN			���˷��ü�¼.����ID%Type,
    ����ʱ��_IN			���˷��ü�¼.����ʱ��%Type,
    ����Ա���_IN		���˷��ü�¼.����Ա���%Type,
    ����Ա����_IN		���˷��ü�¼.����Ա����%Type,
    ��ҩ����_IN			���˷��ü�¼.��ҩ����%Type:=Null,
    �Ƿ���_IN			���˷��ü�¼.�Ƿ���%Type:=0
) AS
     --���ܣ������շ�ʱ��ȡ���۵�����    
     --������
     --        �շѽ���_IN:��ʽ="���㷽ʽ|������|�������||.....",ע���޽������Ҫ�ÿո����
     --        ���ս���_IN:��ʽ="���㷽ʽ|������||....."
     --˵����
     --        1.��ȡ���۷���ʱ,�ż��������ػ���,�ڻ���ʱ������;��ҩƷ��ػ���(��������)����ʱ�Ѿ����㡣
     --        2.��ȡ���۷���ʱ,Ŀǰ���漰������δ������չ�����,�ɻ���ʱֱ�Ӵ���
    --���α�Ϊ����ԭ��������
    Cursor c_Price is
        Select * From ���˷��ü�¼
        Where NO=NO_IN And ��¼����=1 And ��¼״̬=0 And ����Ա���� is Null
        Order by ���;
    r_PriceRow c_Price%RowType;

    --���α������շѳ�Ԥ���Ŀ���Ԥ���б�(��SQL�ο�סԺ����)
    --��ID�������ȳ��ϴ�δ����ġ�
    Cursor c_Deposit(v_����ID ������Ϣ.����ID%Type) is
        Select * From(
            Select A.ID,A.��¼״̬,A.NO,Nvl(A.���,0) as ���
            From ����Ԥ����¼ A,(
                    Select NO,Sum(Nvl(A.���,0)) as ��� 
                    From ����Ԥ����¼ A
                Where A.����ID Is Null And Nvl(A.���,0)<>0 And A.����ID=v_����ID
                  Group by NO Having Sum(Nvl(A.���,0))<>0
                    ) B
            Where A.����ID Is Null And Nvl(A.���,0)<>0 And A.NO=B.NO And A.����ID=v_����ID
            Union All
            Select 0 as ID,��¼״̬,NO,Sum(Nvl(���,0)-Nvl(��Ԥ��,0)) as ���
            From ����Ԥ����¼
            Where ��¼���� IN(1,11) And ����ID is Not NULL And Nvl(���,0)<>Nvl(��Ԥ��,0) And ����ID=v_����ID
            Having Sum(Nvl(���,0)-Nvl(��Ԥ��,0))<>0
            Group by ��¼״̬,NO)
        Order by ID,NO;

    --���α����ڲ��˷��û��ܴ���
    Cursor c_Money is
        Select TRUNC(�Ǽ�ʱ��) as ����,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,�����־,
             Sum(Ӧ�ս��) as Ӧ�ս��,Sum(ʵ�ս��) as ʵ�ս��,Sum(���ʽ��) as ���ʽ��
        From ���˷��ü�¼
        Where ��¼����=1 And ��¼״̬=1 And NO=NO_IN
        Group by TRUNC(�Ǽ�ʱ��),���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,�����־;
    r_MoneyRow c_Money%RowType;

    --Ԥ���������ر���
    v_Ԥ�����		����Ԥ����¼.��Ԥ��%Type;
    v_��������		Varchar2(500);
    v_��ǰ����		Varchar2(50);
    v_���㷽ʽ		����Ԥ����¼.���㷽ʽ%Type;
    v_������		����Ԥ����¼.��Ԥ��%Type;
    v_�������		����Ԥ����¼.�������%Type;

    v_��ʶ��		���˷��ü�¼.��ʶ��%Type;

    --��ʱ����
    v_Count			NUMBER;
    v_Date			DATE;
    v_Error			VARCHAR2(255);
    Err_Custom		EXCEPTION;
BEGIN
    Select Count(ID) Into v_Count From ���˷��ü�¼ Where ��¼����=1 And ��¼״̬=0 And NO=NO_IN And ����Ա���� is Null;
    If v_Count=0 Then
        v_Error:='���ܶ�ȡ���۵�����,�õ��ݿ����Ѿ�ɾ�����Ѿ��շѣ�';
        Raise Err_Custom;
    End If;
   
    Select Sysdate Into v_Date From Dual;
	
  	If Nvl(����ID_IN,0)<>0 Then
  		Select 
  			Decode(��ǰ����ID,NULL,�����,סԺ��) Into v_��ʶ��
  		From ������Ϣ
  		Where ����ID=����ID_IN;
  	End IF;

    --ѭ�������˷��ü�¼
    For r_PriceRow In c_Price Loop
        --ִ��״̬����ֶβ�����,�ڻ���ʱ����;��Ϊ����δ�շѷ�ҩ,������ִ�еĻ��۵��������շѲ����ġ�
        --Ϊ��֤��Ԥ�������¼��ʱ����ͬ,������д�Ǽ�ʱ��,��ҩƷ���ֲ��䶯��
        Update ���˷��ü�¼
            Set ��¼״̬=1,
                ����ID=Decode(����ID_IN,0,Null,����ID_IN),
                ��ҳID=��ҳID_IN,
				��ʶ��=v_��ʶ��,
                ����=����_IN,
                ����=����_IN,
                ����=����_IN,
                �Ա�=�Ա�_IN,
                ���˲���ID=Nvl(���˲���ID_IN,���˲���ID),--���ܱ���ҽ�����͵�����
                ���˿���ID=Nvl(���˿���ID_IN,���˿���ID),
                ��������ID=Nvl(��������ID_IN,��������ID),
                ������=Nvl(������_IN,������),
                ���ʽ��=ʵ�ս��,
                ����ID=����ID_IN,
                ����ʱ��=����ʱ��_IN,
                �Ǽ�ʱ��=v_Date,
                ����Ա���=����Ա���_IN,
                ����Ա����=����Ա����_IN,
                ��ҩ����=Nvl(��ҩ����_IN,��ҩ����),
                �Ƿ���=�Ƿ���_IN
        Where ID=r_Pricerow.ID;
    End Loop;

    --Ԥ������ؽ���
    --�շѽ���
    If �շѽ���_IN IS Not NULL Then
        v_��������:=�շѽ���_IN||'||';
        While v_�������� IS Not NULL Loop
			v_��ǰ����:=Substr(v_��������,1,Instr(v_��������,'||')-1);
			v_���㷽ʽ:=Substr(v_��ǰ����,1,Instr(v_��ǰ����,'|')-1);
			v_��ǰ����:=Substr(v_��ǰ����,Instr(v_��ǰ����,'|')+1);
			v_������:=To_Number(Substr(v_��ǰ����,1,Instr(v_��ǰ����,'|')-1));
			v_�������:=LTrim(Substr(v_��ǰ����,Instr(v_��ǰ����,'|')+1));
			
			If Nvl(v_������,0)<>0 Then
				Insert Into ����Ԥ����¼(
					ID,��¼����,NO,��¼״̬,����ID,��ҳID,ժҪ,���㷽ʽ,�������,�տ�ʱ��,����Ա���,����Ա����,��Ԥ��,����ID)
				Values(
					����Ԥ����¼_ID.Nextval,3,NO_IN,1,Decode(����ID_IN,0,Null,����ID_IN),��ҳID_IN,
					'�շѽ���',v_���㷽ʽ,v_�������,v_Date,����Ա���_IN,����Ա����_IN,v_������,����ID_IN);
			End IF;

            v_��������:=Substr(v_��������,Instr(v_��������,'||')+2);
        End Loop;
     End IF;

     IF ���ս���_IN IS NOT NULL Then 
        --�������ս���    
        v_��������:=���ս���_IN||'||';
        While v_�������� IS Not NULL Loop
            v_��ǰ����:=Substr(v_��������,1,Instr(v_��������,'||')-1);

            v_���㷽ʽ:=Substr(v_��ǰ����,1,Instr(v_��ǰ����,'|')-1);
            v_������:=To_Number(Substr(v_��ǰ����,Instr(v_��ǰ����,'|')+1));
			
			If Nvl(v_������,0)<>0 Then
				Insert Into ����Ԥ����¼(
					ID,��¼����,NO,��¼״̬,����ID,ժҪ,���㷽ʽ,�տ�ʱ��,����Ա���,����Ա����,��Ԥ��,����ID)
				Values(
					����Ԥ����¼_ID.Nextval,3,NO_IN,1,Decode(����ID_IN,0,Null,����ID_IN),'���ս���',
					v_���㷽ʽ,v_Date,����Ա���_IN,����Ա����_IN,v_������,����ID_IN);
			End IF;

            v_��������:=Substr(v_��������,Instr(v_��������,'||')+2);
        End Loop;
    End IF;

    --Ԥ������
    IF Nvl(��Ԥ����_IN,0)<>0 THEN
        IF Nvl(����ID_IN,0)=0 Then
            v_Error:='����ȷ�����˲���ID,�շ�ʹ��Ԥ�������ʧ�ܣ�';
            Raise Err_Custom;
        End if;

        v_Ԥ�����:=��Ԥ����_IN;
        For r_Deposit IN c_Deposit(����ID_IN) Loop
            IF r_Deposit.ID<>0 Then
                --��һ�γ�Ԥ��
                Update ����Ԥ����¼ 
                    Set ��Ԥ��=Decode(Sign(r_Deposit.���-v_Ԥ�����),-1,r_Deposit.���,v_Ԥ�����),
                        ����ID=����ID_IN
                Where ID=r_Deposit.ID;
            Else
                --���ϴ�ʣ���
                INSERT Into ����Ԥ����¼(
                    ID,NO,ʵ��Ʊ��,��¼����,��¼״̬,����ID,��ҳID,����ID,���,
                    ���㷽ʽ,�������,ժҪ,�ɿλ,��λ������,��λ�ʺ�,�տ�ʱ��,
                    ����Ա����,����Ա���,��Ԥ��,����ID)
                Select ����Ԥ����¼_ID.Nextval,NO,ʵ��Ʊ��,11,��¼״̬,����ID,
                     ��ҳID,����ID,NULL,���㷽ʽ,�������,ժҪ,�ɿλ,
                     ��λ������,��λ�ʺ�,�տ�ʱ��,����Ա����,����Ա���,
                     Decode(Sign(r_Deposit.���-v_Ԥ�����),-1,r_Deposit.���,v_Ԥ�����),����ID_IN
                From ����Ԥ����¼
                Where NO=r_Deposit.NO And ��¼״̬=r_Deposit.��¼״̬
                    AND ��¼���� IN(1,11) And RowNum=1;
            End IF;
            --����Ƿ��Ѿ�������
            IF r_Deposit.���<v_Ԥ����� Then
                v_Ԥ�����:=v_Ԥ�����-r_Deposit.���;
            Else
                v_Ԥ�����:=0;
            End IF;
            IF v_Ԥ�����=0 Then 
                Exit;
            End IF;
        End Loop;
        --������Ƿ��㹻
        IF v_Ԥ�����>0 Then
            v_Error:='���˵ĵ�ǰԤ�������� '||Ltrim(To_Char(��Ԥ����_IN,'9999999990.00'))||' ��';
            Raise Err_Custom;
        End IF;

        --���²���Ԥ�����
        Update ������� Set Ԥ�����=Nvl(Ԥ�����,0)-��Ԥ����_IN Where ����ID=����ID_IN And ����=1;
        IF SQL%RowCount=0 Then
            Insert Into �������(����ID,Ԥ�����,����) Values(����ID_IN,-��Ԥ����_IN,1);
        End IF;
        Delete From ������� Where ����ID=����ID_IN And ����=1 And Nvl(�������,0)=0 And Nvl(Ԥ�����,0)=0;
    End IF;

    --��ػ��ܱ�Ĵ���

    --����"��Ա�ɿ����"
	--�շѽ���
    IF �շѽ���_IN IS Not NULL Then
        v_��������:=�շѽ���_IN||'||';
        While v_�������� IS Not NULL Loop
			v_��ǰ����:=Substr(v_��������,1,Instr(v_��������,'||')-1);
			v_���㷽ʽ:=Substr(v_��ǰ����,1,Instr(v_��ǰ����,'|')-1);
			v_��ǰ����:=Substr(v_��ǰ����,Instr(v_��ǰ����,'|')+1);
			v_������:=To_Number(Substr(v_��ǰ����,1,Instr(v_��ǰ����,'|')-1));
			
			If Nvl(v_������,0)<>0 Then
				Update ��Ա�ɿ����
					Set ���=Nvl(���,0)+Nvl(v_������,0)
				 Where �տ�Ա=����Ա����_IN
					And ����=1 And ���㷽ʽ=v_���㷽ʽ;
				If SQL%RowCount=0 Then
					Insert Into ��Ա�ɿ����(
						�տ�Ա,���㷽ʽ,����,���)
					Values(
						����Ա����_IN,v_���㷽ʽ,1,Nvl(v_������,0));
				End If;
			End IF;

            v_��������:=Substr(v_��������,Instr(v_��������,'||')+2);
        End Loop;
    End IF;

    --�������ս���
    IF ���ս���_IN IS Not NULL Then
        v_��������:=���ս���_IN||'||';
        While v_�������� IS Not NULL Loop
            v_��ǰ����:=Substr(v_��������,1,Instr(v_��������,'||')-1);

            v_���㷽ʽ:=Substr(v_��ǰ����,1,Instr(v_��ǰ����,'|')-1);
            v_������:=To_Number(Substr(v_��ǰ����,Instr(v_��ǰ����,'|')+1));
			
			If Nvl(v_������,0)<>0 Then
				Update ��Ա�ɿ����
					Set ���=Nvl(���,0)+Nvl(v_������,0)
				 Where �տ�Ա=����Ա����_IN
					And ����=1 And ���㷽ʽ=v_���㷽ʽ;
				If SQL%RowCount=0 Then
					Insert Into ��Ա�ɿ����(
						�տ�Ա,���㷽ʽ,����,���)
					Values(
						����Ա����_IN,v_���㷽ʽ,1,Nvl(v_������,0));
				End If;
			End IF;

            v_��������:=Substr(v_��������,Instr(v_��������,'||')+2);
        End Loop;
    End IF;
    Delete From ��Ա�ɿ���� Where ����=1 And �տ�Ա=����Ա����_IN And Nvl(���,0)=0;

    --���˷��û���
    For r_MoneyRow In c_Money Loop
        Update ���˷��û���
            Set Ӧ�ս��=Nvl(Ӧ�ս��,0)+r_MoneyRow.Ӧ�ս��,
                ʵ�ս��=Nvl(ʵ�ս��,0)+r_MoneyRow.ʵ�ս��,
                ���ʽ��=Nvl(���ʽ��,0)+r_MoneyRow.���ʽ��
         Where ����=r_MoneyRow.����
            And Nvl(���˲���ID,0)=Nvl(r_MoneyRow.���˲���ID,0)
            And Nvl(���˿���ID,0)=Nvl(r_MoneyRow.���˿���ID,0)
            And Nvl(��������ID,0)=Nvl(r_MoneyRow.��������ID,0)
            And Nvl(ִ�в���ID,0)=Nvl(r_MoneyRow.ִ�в���ID,0)
            And ������ĿID+0=r_MoneyRow.������ĿID
            And ��Դ;��=r_MoneyRow.�����־ And ���ʷ���=0;

        If SQL%RowCount=0 Then
            Insert Into ���˷��û���(
                ����,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,��Դ;��,���ʷ���,Ӧ�ս��,ʵ�ս��,���ʽ��)
            Values(
                r_MoneyRow.����,r_MoneyRow.���˲���ID,r_MoneyRow.���˿���ID,r_MoneyRow.��������ID,r_MoneyRow.ִ�в���ID,
                r_MoneyRow.������ĿID,r_MoneyRow.�����־,0,r_MoneyRow.Ӧ�ս��,r_MoneyRow.ʵ�ս��,r_MoneyRow.���ʽ��);
        End If;
    End Loop;

    --ҩƷ���ַǷ�����Ϣ���޸�
    --ҩƷδ����¼(����ѷ�ҩ���޸Ĳ���),���뷢ҩʱ�޿ⷿID
	--���ܴ��ڲ��Ϻ�ҩƷ�ⷿ��ͬ���������޷�ҩ����
    Update δ��ҩƷ��¼
        Set ����ID=Decode(����ID_IN,0,Null,����ID_IN),
            ��ҳID=��ҳID_IN,����=����_IN,
            �Է�����ID=��������ID_IN,���շ�=1,��������=v_Date
     Where ����=24 And NO=NO_IN And Nvl(�ⷿID,0) IN(
		Select Distinct Nvl(ִ�в���ID,0) From ���˷��ü�¼ 
			Where ��¼����=1 And ��¼״̬=1 And NO=NO_IN And �շ����='4');
	
	Update δ��ҩƷ��¼
        Set ����ID=Decode(����ID_IN,0,Null,����ID_IN),
            ��ҳID=��ҳID_IN,����=����_IN,
            �Է�����ID=��������ID_IN,���շ�=1,
            ��ҩ����=Nvl(��ҩ����_IN,��ҩ����),��������=v_Date
     Where ����=8 And NO=NO_IN And Nvl(�ⷿID,0) IN(
		Select Distinct Nvl(ִ�в���ID,0) From ���˷��ü�¼ 
			Where ��¼����=1 And ��¼״̬=1 And NO=NO_IN And �շ���� IN('5','6','7'));

    --ҩƷ�շ���¼(�����Ѿ���ҩ��ȡ����ҩ,���м�¼����)
    Update ҩƷ�շ���¼
        Set �Է�����ID=��������ID_IN,��������=Decode(Sign(Nvl(�������,v_Date)-v_Date),-1,��������,v_Date)
     Where ����=24 And NO=NO_IN And ����ID+0 IN(
		Select ID From ���˷��ü�¼ 
			Where ��¼����=1 And ��¼״̬=1 And NO=NO_IN And �շ����='4');

	Update ҩƷ�շ���¼
        Set �Է�����ID=��������ID_IN,��ҩ����=Nvl(��ҩ����_IN,��ҩ����),��������=Decode(Sign(Nvl(�������,v_Date)-v_Date),-1,��������,v_Date)
     Where ����=8 And NO=NO_IN And ����ID+0 IN(
		Select ID From ���˷��ü�¼ 
			Where ��¼����=1 And ��¼״̬=1 And NO=NO_IN And �շ���� IN('5','6','7'));

	--���²��ݲ�����Ϣ
	If ����ID_IN IS Not NULL Then
		Update ������Ϣ 
			Set �Ա�=Nvl(�Ա�_IN,�Ա�),
				����=Nvl(����_IN,����)
		Where ����ID=����ID_IN;
		If ����_IN IS Not NULL  And ��ҳID_IN IS NULL Then
			Update ������Ϣ Set ҽ�Ƹ��ʽ=(Select ���� From ҽ�Ƹ��ʽ Where ����=����_IN) Where ����ID=����ID_IN;
		End IF;
	End IF;
EXCEPTION
    WHEN Err_Custom Then Raise_application_errOr(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    WHEN OTHERS Then zl_ErrOrCenter(SQLCODE,SQLERRM);
End zl_�����շѼ�¼_Insert;
/

Create Or Replace Procedure zl_�����շ����_Insert(
--���ܣ���д�����շ�ʱ��������������,�Ա�֤�������һ�¡�
--      ���������շ�,������ȡ���۵�����,�����˷ѡ�
--���������_IN=Sum(������)-Sum(���ʽ��)
--      �˷�_IN=�����Ƿ������˷�ʱ����(�����ŵ��ݲ����˷�ʱ�ŵ���)
    NO_IN			���˷��ü�¼.NO%Type,
    ���_IN         ���˷��ü�¼.ʵ�ս��%Type,
    �˷�_IN         Number:=0
) AS
    v_�շ����      ���˷��ü�¼.�շ����%Type;
    v_�շ�ϸĿID    ���˷��ü�¼.�շ�ϸĿID%Type;
    v_���㵥λ      ���˷��ü�¼.���㵥λ%Type;
    v_������ĿID    ���˷��ü�¼.������ĿID%Type;
    v_�վݷ�Ŀ      ���˷��ü�¼.�վݷ�Ŀ%Type;
    
    v_���˲���ID    ���˷��ü�¼.���˲���ID%Type;
    v_���˿���ID    ���˷��ü�¼.���˿���ID%Type;
    v_��������ID    ���˷��ü�¼.��������ID%Type;
    v_ִ�в���ID    ���˷��ü�¼.ִ�в���ID%Type;
    v_�Ǽ�ʱ��      ���˷��ü�¼.�Ǽ�ʱ��%Type;
    v_������Դ		���˷��ü�¼.�����־%Type;

    v_����ID		���˷��ü�¼.ID%Type;
    v_���			���˷��ü�¼.���%Type;
    v_����ID		���˷��ü�¼.����ID%Type;
    v_ִ��״̬		���˷��ü�¼.ִ��״̬%Type;

    v_Temp			Varchar2(255);
    v_��Ա���		���˷��ü�¼.����Ա���%Type;
    v_��Ա����		���˷��ü�¼.����Ա����%Type;

    v_Sign			Number;
    Err_Custom		Exception;
    v_Error			Varchar2(255);
Begin
    If Nvl(���_IN,0)=0 THEN 
        Return;
    End IF;

    --����Ա��Ϣ:����ID,��������;��ԱID,��Ա���,��Ա����
    v_Temp:=zl_Identity;
    v_ִ�в���ID:=To_Number(Substr(v_Temp,1,Instr(v_Temp,',')-1));
    v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
    v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
    v_��Ա���:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
    v_��Ա����:=Substr(v_Temp,Instr(v_Temp,',')+1);

    --�����Ŀ����
    Begin
        Select A.���,A.ID,A.���㵥λ,C.ID,C.�վݷ�Ŀ 
            Into v_�շ����,v_�շ�ϸĿID,v_���㵥λ,v_������ĿID,v_�վݷ�Ŀ
        From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C,�շ��ض���Ŀ D
        Where D.�ض���Ŀ='�����' And D.�շ�ϸĿID=A.ID
            And A.ID=B.�շ�ϸĿID And B.������ĿID=C.ID
            And ((Sysdate Between B.ִ������ And B.��ֹ����) 
                Or (Sysdate>=B.ִ������ And B.��ֹ���� is NULL));
    Exception
        When Others Then
        Begin
            v_Error:='������ȷ��ȡ�������������Ŀ��Ϣ�����ȼ�����Ŀ�Ƿ���ȷ���á�';
            Raise Err_Custom;
        End;
    End;
    
    If Nvl(�˷�_IN,0)<>0 Then
        --�˷Ѵ������ʱ,���շѵ�����¼�ϴ��������,��ֱ�������˷Ѽ�¼
        Begin
            Select ��� Into v_��� From ���˷��ü�¼ Where NO=NO_IN And ��¼����=1 And ��¼״̬ IN(1,3) And ���ӱ�־=9;
        Exception
            When Others Then NULL;
        End;
    End IF;
    If v_��� IS NULL Then
        Select Max(���)+1 Into v_��� From ���˷��ü�¼ Where NO=NO_IN And ��¼����=1;    
    End IF;

    v_Sign:=1;v_ִ��״̬:=0;
    --�ñ���Ŀ�ڼ����˷�(�˷�ʱ)
    If Nvl(�˷�_IN,0)<>0 Then
        v_Sign:=-1;
        Select -1*(Nvl(Max(Abs(ִ��״̬)),0)+1) Into v_ִ��״̬
        From ���˷��ü�¼ 
        Where NO=NO_IN And ��¼����=1 And ��¼״̬=2 And ���=v_���;
    End IF;

    --ȡ����շѻ��˷ѵĽ���ID(��Ҫ��Ϊ��ȷ�����һ���˷�)
    Select Max(����ID) Into v_����ID From ���˷��ü�¼ Where NO=NO_IN And ��¼����=1;
    Select ���˷��ü�¼_ID.Nextval Into v_����ID From Dual;

    --���˷��ü�¼:���ӱ�־=9
    Insert Into ���˷��ü�¼(
        ID,��¼����,NO,ʵ��Ʊ��,��¼״̬,���,��������,�۸񸸺�,�����־,����ID,��ʶ��,����,����,�Ա�,
        ����,���˲���ID,���˿���ID,�ѱ�,�շ����,�շ�ϸĿID,���㵥λ,��ҩ����,����,����,�Ӱ��־,���ӱ�־,
        ������ĿID,�վݷ�Ŀ,��׼����,Ӧ�ս��,ʵ�ս��,���ʷ���,������,��������ID,������,����ʱ��,�Ǽ�ʱ��,
        ִ�в���ID,ִ��״̬,����ID,���ʽ��,����Ա���,����Ա����,�Ƿ��ϴ�)
    Select
        v_����ID,��¼����,NO,ʵ��Ʊ��,��¼״̬,v_���,NULL,NULL,�����־,����ID,��ʶ��,����,����,�Ա�,����,
        ���˲���ID,���˿���ID,�ѱ�,v_�շ����,v_�շ�ϸĿID,v_���㵥λ,��ҩ����,1,v_Sign*1,�Ӱ��־,9,
        v_������ĿID,v_�վݷ�Ŀ,���_IN,v_Sign*���_IN,v_Sign*���_IN,���ʷ���,������,��������ID,������,����ʱ��,�Ǽ�ʱ��,
        v_ִ�в���ID,v_ִ��״̬,����ID,v_Sign*���_IN,v_��Ա���,v_��Ա����,1
    From ���˷��ü�¼
    Where NO=NO_IN And ��¼����=1 And ��¼״̬=Decode(Nvl(�˷�_IN,0),0,1,2) And ����ID=v_����ID And Rownum=1;

    If Nvl(�˷�_IN,0)<>0 Then
        Update ���˷��ü�¼ Set ��¼״̬=3 Where ��¼����=1 And NO=NO_IN And ��¼״̬=1 And ���=v_���;
    End IF;

    --���˷��û���
    Select ���˲���ID,���˿���ID,��������ID,�Ǽ�ʱ��,�����־
        Into v_���˲���ID,v_���˿���ID,v_��������ID,v_�Ǽ�ʱ��,v_������Դ
    From ���˷��ü�¼ Where ID=v_����ID;
    Update ���˷��û���
        Set Ӧ�ս��=Nvl(Ӧ�ս��,0)+v_Sign*���_IN,
            ʵ�ս��=Nvl(ʵ�ս��,0)+v_Sign*���_IN,
            ���ʽ��=Nvl(���ʽ��,0)+v_Sign*���_IN
    Where ����=Trunc(v_�Ǽ�ʱ��)
        And Nvl(���˲���ID,0)=Nvl(v_���˲���ID,0)
        And Nvl(���˿���ID,0)=Nvl(v_���˿���ID,0)
        And Nvl(��������ID,0)=Nvl(v_��������ID,0)
        And Nvl(ִ�в���ID,0)=Nvl(v_ִ�в���ID,0)
        And ������ĿID+0=v_������ĿID
        And ��Դ;��=v_������Դ And ���ʷ���=0;
    IF SQL%RowCount=0 Then
        Insert Into ���˷��û���(
            ����,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,��Դ;��,���ʷ���,Ӧ�ս��,ʵ�ս��,���ʽ��)
        Values(
            Trunc(v_�Ǽ�ʱ��),v_���˲���ID,v_���˿���ID,v_��������ID,v_ִ�в���ID,v_������ĿID,v_������Դ,0,v_Sign*���_IN,v_Sign*���_IN,v_Sign*���_IN);
    End IF;
Exception
    When Err_Custom Then Raise_Application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When Others Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_�����շ����_Insert;
/

Create Or Replace Procedure zl_�����շѽ���_Update(
	   ����Id_IN			���˷��ü�¼.����ID%Type,
	   �շѽ���_IN			Varchar2,
	   ��Ԥ����_IN			����Ԥ����¼.��Ԥ��%Type,
	   ���ս���_IN			Varchar2,
	   ���_IN				���˷��ü�¼.ʵ�ս��%Type
 )
 As 
 --����:�����շ�ʱ��ҽ����ʽ�����,��ؽ�����Ϣ�ĵ���
 --     ��ΪԤ�����,���ɵ�ҽ���������ܶ��̯���ܻ�����ʽ����ʱ�в���,�����ṩ��У�Թ���,
 --		����Ա�ڽ���У��ʱ,���Ե�����ҽ�����㷽ʽ�ĸ��ֽ������ʽ,�������ɽ��㴮,���ҿ��ܲ��������.
  
  --���α������շѳ�Ԥ���Ŀ���Ԥ���б�(��SQL�ο�סԺ����)
    --��ID�������ȳ��ϴ�δ����ġ�
    Cursor c_Deposit(v_����ID ������Ϣ.����ID%Type) is
      Select * From(
          Select A.ID,A.��¼״̬,A.NO,Nvl(A.���,0) as ���
          From ����Ԥ����¼ A,(
                  Select NO,Sum(Nvl(A.���,0)) as ��� 
                  From ����Ԥ����¼ A
              Where A.����ID Is Null And Nvl(A.���,0)<>0 And A.����ID=v_����ID
                Group by NO Having Sum(Nvl(A.���,0))<>0
                  ) B
          Where A.����ID Is Null And Nvl(A.���,0)<>0 And A.NO=B.NO And A.����ID=v_����ID
          Union All
          Select 0 as ID,��¼״̬,NO,Sum(Nvl(���,0)-Nvl(��Ԥ��,0)) as ���
          From ����Ԥ����¼
          Where ��¼���� IN(1,11) And ����ID is Not NULL And Nvl(���,0)<>Nvl(��Ԥ��,0) And ����ID=v_����ID
          Having Sum(Nvl(���,0)-Nvl(��Ԥ��,0))<>0
          Group by ��¼״̬,NO)
      Order by ID,NO;
 
 --���̱���
    v_��������		Varchar2(500);
    v_��ǰ����		Varchar2(50);
    v_���㷽ʽ		����Ԥ����¼.���㷽ʽ%Type;
    v_������		����Ԥ����¼.��Ԥ��%Type;
    v_�������		����Ԥ����¼.�������%Type;
	
    v_����ID		���˷��ü�¼.ID%Type;
    v_���			���˷��ü�¼.���%Type;
	
    v_�շ����		���˷��ü�¼.�շ����%Type;
    v_�շ�ϸĿID	���˷��ü�¼.�շ�ϸĿID%Type;
    v_���㵥λ		���˷��ü�¼.���㵥λ%Type;
    v_������ĿID	���˷��ü�¼.������ĿID%Type;
    v_�վݷ�Ŀ		���˷��ü�¼.�վݷ�Ŀ%Type;
    v_ִ�в���ID	���˷��ü�¼.ִ�в���ID%Type;
    v_Temp			Varchar2(500);
	
    v_No			����Ԥ����¼.No%Type;
    v_����Id		����Ԥ����¼.����Id%Type;
    v_��ҳId		����Ԥ����¼.��ҳId%Type;
    v_�տ�ʱ��		����Ԥ����¼.�տ�ʱ��%Type;
    v_����Ա���	����Ԥ����¼.����Ա���%Type;
    v_����Ա����	����Ԥ����¼.����Ա����%Type;
	
    v_Ԥ�����		����Ԥ����¼.��Ԥ��%Type;	
	 
    v_Error			VARCHAR2(255);
    Err_Custom		EXCEPTION;
 Begin
 
 --1.ȡԤ����¼��Ҫ�������Ϣ
    Select No,����Id,��ҳId,�Ǽ�ʱ��,����Ա���,����Ա���� 
		   Into  v_No,v_����Id,v_��ҳId,v_�տ�ʱ��,v_����Ա���,v_����Ա����
    From ���˷��ü�¼ Where ����ID=����ID_IN And Rownum=1 And ��¼����=1;		
    --��������Ϣ
    Begin
        Select A.���,A.ID,A.���㵥λ,C.ID,C.�վݷ�Ŀ 
        Into v_�շ����,v_�շ�ϸĿID,v_���㵥λ,v_������ĿID,v_�վݷ�Ŀ
        From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C,�շ��ض���Ŀ D
        Where D.�ض���Ŀ='�����' And D.�շ�ϸĿID=A.Id  And A.ID=B.�շ�ϸĿID And B.������ĿID=C.ID
            And Sysdate Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-01-01','YYYY-MM-DD')) ;
    Exception
        When Others Then
        Begin
            v_Error:='������ȷ��ȡ�շ���������Ϣ�����ȼ�����Ŀ�Ƿ�������ȷ��';
            Raise Err_Custom;
        End;
    End;	
	
 --2.ɾ���ɵļ�¼,���˻�������
    --������Ա�ɿ����,�������,
    For C_DEL In (SELECT * FROM ����Ԥ����¼ WHERE ����ID=����ID_IN And ��¼����=3) Loop
	    Update ��Ա�ɿ���� Set ���=Nvl(���,0)-Nvl(C_DEL.��Ԥ��,0) Where ���㷽ʽ=C_DEL.���㷽ʽ;	   
      	If SQL%RowCount=0 Then
           Insert Into ��Ա�ɿ����(�տ�Ա,���㷽ʽ,����,���) Values(C_DEL.����Ա����,C_DEL.���㷽ʽ,1,-1*C_DEL.��Ԥ��);
		End If;
    End Loop;
	
    If v_����Id>0 Then
    	Begin
        	Select Sum(��Ԥ��) Into V_Ԥ����� From ����Ԥ����¼ Where ����Id=����id_IN And ��¼���� In (1,11);
        Exception
        	When Others Then NULL;
    	End;	
    	If v_Ԥ�����<>0 Then
        	Update ������� Set Ԥ�����=Nvl(Ԥ�����,0)+V_Ԥ����� Where ����ID=v_����Id And ����=1;
        	IF SQL%RowCount=0 Then
            	Insert Into �������(����ID,Ԥ�����,����) Values(v_����Id,V_Ԥ�����,1);
            End IF;
    	End If;
    End If;
	
	--���˲��˷��û���
	--ֻ���ܲ��������ı仯,�����Ĳ����,����ֻ���ܴ���һ��,��Ϊ�˱�������������α�
    For C_Error In (
        Select TRUNC(�Ǽ�ʱ��) as ����,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,�����־,Ӧ�ս��,ʵ�ս��,���ʽ��
        From ���˷��ü�¼
        Where ��¼����=1 And ��¼״̬=1 And ����Id=����Id_IN And ���ӱ�־=9
    ) Loop
        Update ���˷��û���
            Set Ӧ�ս��=Nvl(Ӧ�ս��,0)-C_Error.Ӧ�ս��,ʵ�ս��=Nvl(ʵ�ս��,0)-C_Error.ʵ�ս��,���ʽ��=Nvl(���ʽ��,0)-C_Error.���ʽ��
        Where ����=C_Error.����
            And Nvl(���˲���ID,0)=Nvl(C_Error.���˲���ID,0) And Nvl(���˿���ID,0)=Nvl(C_Error.���˿���ID,0)
            And Nvl(��������ID,0)=Nvl(C_Error.��������ID,0) And Nvl(ִ�в���ID,0)=Nvl(C_Error.ִ�в���ID,0)
            And ������ĿID+0=C_Error.������ĿId And ��Դ;��=C_Error.�����־ And ���ʷ���=0; 
        If SQL%RowCount=0 Then
            Insert Into ���˷��û���(
                ����,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,��Դ;��,���ʷ���,Ӧ�ս��,ʵ�ս��,���ʽ��)
            Values(
                C_Error.����,C_Error.���˲���ID,C_Error.���˿���ID,C_Error.��������ID,C_Error.ִ�в���ID,
                C_Error.������ĿID,C_Error.�����־,0,-1*C_Error.Ӧ�ս��,-1*C_Error.ʵ�ս��,-1*C_Error.���ʽ��);
        End If;
    End Loop; 
 
    --ɾ���շѽ���,���ս����¼		     
    Delete ����Ԥ����¼ Where ����ID=����ID_IN And ��¼����=3; 
    --��һ�γ�Ԥ����,��ճ����
    Update ����Ԥ����¼ Set ��Ԥ��=Null,����Id=Null	Where ����Id=����ID_IN And ��¼����=1;
    --ɾ�������
    Delete ����Ԥ����¼ Where ����Id=����ID_IN And ��¼����=11;
    --ɾ������¼
    Delete ���˷��ü�¼ Where ����Id=����Id_IN And ���ӱ�־=9;	
 
 --3.�������˷��ü�¼������¼
    If ���_IN <>0 Then 
        Select ���˷��ü�¼_ID.Nextval Into v_����ID From Dual;
        Select Max(���)+1 Into v_��� From ���˷��ü�¼ Where ����ID=����ID_IN And ��¼����=1;
	 v_Temp:=zl_Identity;
	 v_ִ�в���ID:=To_Number(Substr(v_Temp,1,Instr(v_Temp,',')-1));
			
        Insert Into ���˷��ü�¼(
            ID,��¼����,NO,ʵ��Ʊ��,��¼״̬,���,��������,�۸񸸺�,�����־,����ID,��ʶ��,����,����,�Ա�,
            ����,���˲���ID,���˿���ID,�ѱ�,�շ����,�շ�ϸĿID,���㵥λ,��ҩ����,����,����,�Ӱ��־,���ӱ�־,
            ������ĿID,�վݷ�Ŀ,��׼����,Ӧ�ս��,ʵ�ս��,���ʷ���,������,��������ID,������,����ʱ��,�Ǽ�ʱ��,
            ִ�в���ID,ִ��״̬,����ID,���ʽ��,����Ա���,����Ա����,�Ƿ��ϴ�)
        Select
            v_����ID,��¼����,NO,ʵ��Ʊ��,��¼״̬,v_���,NULL,NULL,�����־,����ID,��ʶ��,����,����,�Ա�,����,
            ���˲���ID,���˿���ID,�ѱ�,v_�շ����,v_�շ�ϸĿID,v_���㵥λ,��ҩ����,1,1,�Ӱ��־,9,
            v_������ĿID,v_�վݷ�Ŀ,���_IN,���_IN,���_IN,���ʷ���,������,��������ID,������,����ʱ��,�Ǽ�ʱ��,
            v_ִ�в���ID,ִ��״̬,����ID,���_IN,����Ա���,����Ա����,1
        From ���˷��ü�¼
        Where ��¼����=1 And ��¼״̬=1 And ����ID=����ID_IN And Rownum=1;
    End If;
  
 --4.�������ɲ���Ԥ����¼�������	
    --4.1.�շѽ���
    If �շѽ���_IN IS Not NULL Then
		v_��������:=�շѽ���_IN||'||';
        While v_�������� IS Not NULL Loop
			v_��ǰ����:=Substr(v_��������,1,Instr(v_��������,'||')-1);
			v_���㷽ʽ:=Substr(v_��ǰ����,1,Instr(v_��ǰ����,'|')-1);
			v_��ǰ����:=Substr(v_��ǰ����,Instr(v_��ǰ����,'|')+1);
			v_������:=To_Number(Substr(v_��ǰ����,1,Instr(v_��ǰ����,'|')-1));
			v_�������:=LTrim(Substr(v_��ǰ����,Instr(v_��ǰ����,'|')+1));
			
			If Nvl(v_������,0)<>0 Then
				Insert Into ����Ԥ����¼(
					ID,��¼����,NO,��¼״̬,����ID,��ҳID,ժҪ,���㷽ʽ,�������,�տ�ʱ��,����Ա���,����Ա����,��Ԥ��,����ID)
				Values(
					����Ԥ����¼_ID.Nextval,3,v_No,1,v_����Id,v_��ҳId,'�շѽ���',
					v_���㷽ʽ,v_�������,v_�տ�ʱ��,v_����Ա���,v_����Ա����,v_������,����ID_IN);
			End IF;
            v_��������:=Substr(v_��������,Instr(v_��������,'||')+2);
        End Loop;
     End IF;
	 
    --4.2.���ս���
    If ���ս���_IN IS Not NULL Then  
		v_��������:=���ս���_IN||'||';
        While v_�������� IS Not NULL Loop
            v_��ǰ����:=Substr(v_��������,1,Instr(v_��������,'||')-1);
            v_���㷽ʽ:=Substr(v_��ǰ����,1,Instr(v_��ǰ����,'|')-1);
            v_������:=To_Number(Substr(v_��ǰ����,Instr(v_��ǰ����,'|')+1));
			
			If Nvl(v_������,0)<>0 Then
				Insert Into ����Ԥ����¼(
					ID,��¼����,NO,��¼״̬,����ID,��ҳID,ժҪ,���㷽ʽ,�տ�ʱ��,����Ա���,����Ա����,��Ԥ��,����ID)
				Values(
					����Ԥ����¼_ID.Nextval,3,v_No,1,v_����Id,v_��ҳId,'���ս���',
					v_���㷽ʽ,v_�տ�ʱ��,v_����Ա���,v_����Ա����,v_������,����ID_IN);
			End IF;
            v_��������:=Substr(v_��������,Instr(v_��������,'||')+2);
        End Loop;
    End IF;

    --4.3.Ԥ������
    IF Nvl(��Ԥ����_IN,0)<>0 THEN
        IF Nvl(v_����Id,0)=0 Then
            v_Error:='����ȷ�����˵Ĳ���ID,�շѲ���ʹ��Ԥ�������,�������ʧ�ܣ�';
            Raise Err_Custom;
        End if;
		
        v_Ԥ�����:=��Ԥ����_IN;
        For r_Deposit IN c_Deposit(v_����Id) Loop
            IF r_Deposit.ID<>0 Then
                --��һ�γ�Ԥ��
                Update ����Ԥ����¼ 
                    Set ��Ԥ��=Decode(Sign(r_Deposit.���-v_Ԥ�����),-1,r_Deposit.���,v_Ԥ�����),����ID=����ID_IN
                Where ID=r_Deposit.ID;
            Else
                --���ϴ�ʣ���
                INSERT Into ����Ԥ����¼(
                    ID,NO,ʵ��Ʊ��,��¼����,��¼״̬,����ID,��ҳID,����ID,���,
                    ���㷽ʽ,�������,ժҪ,�ɿλ,��λ������,��λ�ʺ�,�տ�ʱ��,
                    ����Ա����,����Ա���,��Ԥ��,����ID)
                Select ����Ԥ����¼_ID.Nextval,NO,ʵ��Ʊ��,11,��¼״̬,����ID,
                     ��ҳID,����ID,NULL,���㷽ʽ,�������,ժҪ,�ɿλ,
                     ��λ������,��λ�ʺ�,�տ�ʱ��,����Ա����,����Ա���,
                     Decode(Sign(r_Deposit.���-v_Ԥ�����),-1,r_Deposit.���,v_Ԥ�����),����ID_IN
                From ����Ԥ����¼
                Where NO=r_Deposit.NO And ��¼״̬=r_Deposit.��¼״̬ AND ��¼���� IN(1,11) And RowNum=1;
            End IF;
            --����Ƿ��Ѿ�������
            IF r_Deposit.���<v_Ԥ����� Then
                v_Ԥ�����:=v_Ԥ�����-r_Deposit.���;
            Else
                v_Ԥ�����:=0;
            End IF;
            IF v_Ԥ�����=0 Then 
                Exit;
            End IF;
        End Loop;
        --������Ƿ��㹻
        IF v_Ԥ�����>0 Then
            v_Error:='���˵ĵ�ǰԤ�������� '||Ltrim(To_Char(��Ԥ����_IN,'9999999990.00'))||' ��';
            Raise Err_Custom;
        End IF;

        --���²���Ԥ�����
        Update ������� Set Ԥ�����=Nvl(Ԥ�����,0)-��Ԥ����_IN Where ����ID=v_����Id And ����=1;
        IF SQL%RowCount=0 Then
            Insert Into �������(����ID,Ԥ�����,����) Values(v_����Id,-1*��Ԥ����_IN,1);
        End IF;
        Delete From ������� Where ����ID=v_����Id And ����=1 And Nvl(�������,0)=0 And Nvl(Ԥ�����,0)=0;
    End IF;
	
    --5.��ػ��ܱ�Ĵ���	
    --����"��Ա�ɿ����"
	--�շѽ���
    IF �շѽ���_IN IS Not NULL Then
        v_��������:=�շѽ���_IN||'||';
        While v_�������� IS Not NULL Loop
			v_��ǰ����:=Substr(v_��������,1,Instr(v_��������,'||')-1);
			v_���㷽ʽ:=Substr(v_��ǰ����,1,Instr(v_��ǰ����,'|')-1);
			v_��ǰ����:=Substr(v_��ǰ����,Instr(v_��ǰ����,'|')+1);
			v_������:=To_Number(Substr(v_��ǰ����,1,Instr(v_��ǰ����,'|')-1));
			
			If Nvl(v_������,0)<>0 Then
				Update ��Ա�ɿ����
					Set ���=Nvl(���,0)+Nvl(v_������,0)
				 Where �տ�Ա=v_����Ա����
					And ����=1 And ���㷽ʽ=v_���㷽ʽ;
				If SQL%RowCount=0 Then
					Insert Into ��Ա�ɿ����(
						�տ�Ա,���㷽ʽ,����,���)
					Values(
						v_����Ա����,v_���㷽ʽ,1,Nvl(v_������,0));
				End If;
			End IF;
            v_��������:=Substr(v_��������,Instr(v_��������,'||')+2);
        End Loop;
    End IF;

    --�������ս���
    IF ���ս���_IN IS Not NULL Then
        v_��������:=���ս���_IN||'||';
        While v_�������� IS Not NULL Loop
            v_��ǰ����:=Substr(v_��������,1,Instr(v_��������,'||')-1);
            v_���㷽ʽ:=Substr(v_��ǰ����,1,Instr(v_��ǰ����,'|')-1);
            v_������:=To_Number(Substr(v_��ǰ����,Instr(v_��ǰ����,'|')+1));
			
			If Nvl(v_������,0)<>0 Then
				Update ��Ա�ɿ����
					Set ���=Nvl(���,0)+Nvl(v_������,0)
				 Where �տ�Ա=v_����Ա����
					And ����=1 And ���㷽ʽ=v_���㷽ʽ;
				If SQL%RowCount=0 Then
					Insert Into ��Ա�ɿ����(
						�տ�Ա,���㷽ʽ,����,���)
					Values(
						v_����Ա����,v_���㷽ʽ,1,Nvl(v_������,0));
				End If;
			End IF;
            v_��������:=Substr(v_��������,Instr(v_��������,'||')+2);
        End Loop;
    End IF;
    Delete From ��Ա�ɿ���� Where ����=1 And �տ�Ա=v_����Ա���� And Nvl(���,0)=0;

    --���˷��û���,ֻ���ػ������,��Ϊ��������
    For r_MoneyRow In (
        Select TRUNC(�Ǽ�ʱ��) as ����,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,�����־,Ӧ�ս��,ʵ�ս��,���ʽ��
        From ���˷��ü�¼
        Where ��¼����=1 And ��¼״̬=1 And ����Id=����Id_IN And ���ӱ�־=9
	) Loop
        Update ���˷��û���
            Set Ӧ�ս��=Nvl(Ӧ�ս��,0)+r_MoneyRow.Ӧ�ս��,ʵ�ս��=Nvl(ʵ�ս��,0)+r_MoneyRow.ʵ�ս��,���ʽ��=Nvl(���ʽ��,0)+r_MoneyRow.���ʽ��
        Where ����=r_MoneyRow.����
            And Nvl(���˲���ID,0)=Nvl(r_MoneyRow.���˲���ID,0) And Nvl(���˿���ID,0)=Nvl(r_MoneyRow.���˿���ID,0)
            And Nvl(��������ID,0)=Nvl(r_MoneyRow.��������ID,0) And Nvl(ִ�в���ID,0)=Nvl(r_MoneyRow.ִ�в���ID,0)
            And ������ĿID+0=r_MoneyRow.������ĿId  And ��Դ;��=r_MoneyRow.�����־ And ���ʷ���=0;

        If SQL%RowCount=0 Then
            Insert Into ���˷��û���(
                ����,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,��Դ;��,���ʷ���,Ӧ�ս��,ʵ�ս��,���ʽ��)
            Values(
                r_MoneyRow.����,r_MoneyRow.���˲���ID,r_MoneyRow.���˿���ID,r_MoneyRow.��������ID,r_MoneyRow.ִ�в���ID,
                r_MoneyRow.������ĿID,r_MoneyRow.�����־,0,r_MoneyRow.Ӧ�ս��,r_MoneyRow.ʵ�ս��,r_MoneyRow.���ʽ��);
        End If;
    End Loop; 
 
 	--6.ҽ����ر�Ĵ���
    Delete ҽ���˶Ա� Where ����Id=����Id_IN;
 
EXCEPTION
    WHEN Err_Custom Then Raise_application_errOr(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    WHEN OTHERS Then zl_ErrOrCenter(SQLCODE,SQLERRM);
End zl_�����շѽ���_Update;
/

CREATE OR REPLACE Procedure zl_�����շѼ�¼_Delete(
    NO_IN				���˷��ü�¼.NO%Type,
    ����Ա���_IN		���˷��ü�¼.����Ա���%Type,
    ����Ա����_IN		���˷��ü�¼.����Ա����%Type,
    ҽ�����㷽ʽ_IN		Varchar2:=Null,
    ���_IN				Varchar2:=NULL,
    ���㷽ʽ_IN			����Ԥ����¼.���㷽ʽ%Type:=NULL,
    Ʊ�ݺ�_IN			Ʊ��ʹ����ϸ.����%Type:=NULL,
    ����ID_IN			Ʊ��ʹ����ϸ.����ID%Type:=NULL,
    ���_IN				���˷��ü�¼.ʵ�ս��%Type:=0,
	�˷�ʱ��_IN			���˷��ü�¼.�Ǽ�ʱ��%Type:=NULL,
	Ʊ�ݴ���_IN			Number:=0
)
AS
--���ܣ�ɾ��һ�������շѵ���
--������
--        ���_IN:Ҫ�˷ѵ���Ŀ���,��ʽΪ"1,3,5,6...",ȱʡNULL��ʾ��"δ�˵�"�����С�
--        ���㷽ʽ_IN:��Ϊ�����˷�ʱ,�˷ѽ��Ľ��㷽ʽ��
--		  ���_IN=�����˷ѻ�ҽ��ȫ�˵�ĳ�ֽ������ֽ�ʱ����,���������������������
--		  ҽ�����㷽ʽ_IN=�����ʻ�,���ݲ���
--          ҽ���˷�ʱ,��֧�ֽ������ϵĽ��㷽ʽ,���Ϊ�ձ�ʾ��ҽ���˷ѻ�ҽ���˷�ȫ�������������ϡ�
--        Ʊ�ݺ�_IN,����ID_IN:��Ϊ�����˷�ʱ,��Ҫ�ش��վݵ�(��ʼ)Ʊ�ݺź���������ID��
--        Ʊ�ݴ���_IN=0-���ڵ��ŵ����˷�,��������ʽ����
--                    1-���ڶ��ŵ����˷�,����ȫ����,ֻ�ջ�Ʊ��(ע��ֻ���ջ�һ��)��
--                    2-���ڶ��ŵ����˷�,���ǲ�����,���ﲻ����Ʊ��,ȫ���˷Ѻ󵥶�����

    --���α�ΪҪ�˷ѵ��ݵ�����ԭʼ��¼
      --ҽ��ȫ�˵�ĳ�ֽ������ֽ�Ӷ��������µ����ʱ,�ſ��˴�������,ִ���걾���̺�,��������е������������    
    Cursor c_Bill is
        Select * From ���˷��ü�¼
        Where NO=NO_IN And ��¼����=1 And ��¼״̬ IN(1,3) And 
              NVL(���ӱ�־,0)<>Decode(ҽ�����㷽ʽ_IN,Null,999,Decode(sign(���_IN),0,999,9))
        Order by �շ�ϸĿID,���;

    --���α����ڴ���ҩƷ����������
    --��Ҫ�ܷ��õ�ִ��״̬,��Ϊ���ڴ˲�����
    Cursor c_Stock is
        Select * From ҩƷ�շ���¼
        Where NO=NO_IN And ���� IN(8,24) --@@@
			And Mod(��¼״̬,3)=1 And ����� IS NULL
            And ����ID IN(
                Select ID From ���˷��ü�¼
                Where NO=NO_IN And ��¼����=1 And ��¼״̬ IN(1,3)
                    And �շ���� IN('4','5','6','7')--@@@
                    And (INSTR(','||���_IN||',',','||���||',')>0 Or ���_IN Is Null)
                )
        Order BY ҩƷID;

    --���α����ڴ���δ��ҩƷ��¼
    Cursor c_Spare is
        Select * From δ��ҩƷ��¼ Where NO=NO_IN And ���� IN(8,24);--@@@

    --�ù�����ڴ�����Ա�ɿ�������˵Ĳ�ͬ���㷽ʽ�Ľ��
    Cursor c_Money(v����ID ����Ԥ����¼.����ID%Type) is
        Select ���㷽ʽ,��Ԥ��
        From ����Ԥ����¼
        Where ��¼����=3 And ��¼״̬=2 And ����ID=v����ID
            And ���㷽ʽ is Not Null And Nvl(��Ԥ��,0)<>0;

	v_ҽ��ID		����ҽ����¼.ID%Type;

    v_����ID		������Ϣ.����ID%Type;
    v_����ID		���˷��ü�¼.����ID%Type;
    v_��ӡID		Ʊ�ݴ�ӡ����.ID%Type;
    v_���˽��		����Ԥ����¼.��Ԥ��%Type;
    v_Ԥ�����		����Ԥ����¼.��Ԥ��%Type;

    --�����˷Ѽ������
    v_ʣ������		Number;
    v_ʣ��Ӧ��		Number;
    v_ʣ��ʵ��		Number;
    v_ʣ��ͳ��		Number;

    v_׼������		Number;
    v_�˷Ѵ���		Number;

    v_Ӧ�ս��		Number;
    v_ʵ�ս��		Number;
    v_ͳ����		Number;
    v_�ܽ��		Number;

	v_�״�ȫ��		Number;--�Ƿ��һ���˷�����ȫ��,�������_IN=NULLʱ�ж��Ƿ�Ҳ�����
    v_�����˷�		Number;--�Ƿ��һ���˷���ȫ���˷�,��ÿ���˷ѹ������жϵõ���
    v_ȫ������		Number;--�����˷��Ƿ�ʣ�ಿ��ȫ��������,�˷���ɺ��SQL�õ���

    v_�˷ѽ���		���㷽ʽ.����%Type;

    v_��������      Varchar2(500);

    v_Dec			Number;

	v_Date			Date;
    v_Count			Number;
    Err_Custom		Exception;
    v_Error			Varchar2(255);
Begin
    --�Ƿ��Ѿ�ȫ����ȫִ��(ֻ�Ǹõ������ŵ��ݵļ��)
    Select Nvl(Count(*),0) Into v_Count From ���˷��ü�¼
    Where NO=NO_IN And ��¼����=1 And ��¼״̬ IN(1,3) And Nvl(ִ��״̬,0)<>1;
    IF v_Count = 0 Then
        v_Error := '�õ����е���Ŀ�Ѿ�ȫ����ȫִ�У�';
        Raise Err_Custom;
    End IF;

    --δ��ȫִ�е���Ŀ�Ƿ���ʣ������(ֻ�����ŵ��ݵļ��)
    --ִ��״̬��ԭʼ��¼���ж�
    Select Nvl(Count(*),0) Into v_Count
    From (
        Select ���,Sum(����) as ʣ������
        From (
            Select ��¼״̬,Nvl(�۸񸸺�,���) as ���,
                Avg(Nvl(����,1)*����) as ����
            From ���˷��ü�¼
            Where NO=NO_IN And ��¼����=1 And Nvl(���ӱ�־,0)<>9
                And Nvl(�۸񸸺�,���) IN (
                        Select Nvl(�۸񸸺�,���)
                        From ���˷��ü�¼
                        Where NO=NO_IN And ��¼����=1
                            And ��¼״̬ IN(1,3) And Nvl(ִ��״̬,0)<>1)
            Group by ��¼״̬,Nvl(�۸񸸺�,���)
            )
        Group by ��� Having Sum(����)<>0);
    IF v_Count = 0 Then
        v_Error := '�õ�����δ��ȫִ�в�����Ŀʣ������Ϊ��,û�п����˷ѵķ��ã�';
        Raise Err_Custom;
    End IF;

    ---------------------------------------------------------------------------------
    --���ñ���
	If �˷�ʱ��_IN IS Not NULL Then
		v_Date:=�˷�ʱ��_IN;
	Else
		Select Sysdate Into v_Date From Dual;
	End IF;
    Select ���˽��ʼ�¼_ID.Nextval Into v_����ID From Dual;

    --���С��λ��
    Begin
        Select To_Number(����ֵ) Into v_Dec From ϵͳ������ Where ������=9;
    Exception
        When Others Then v_Dec:=2;
    End;

    --��ȡ���㷽ʽ����
    v_�˷ѽ���:=���㷽ʽ_IN;
    If v_�˷ѽ��� IS NULL Then
        Begin
            Select ���� Into v_�˷ѽ��� From ���㷽ʽ Where ����=1;
        Exception
            When Others Then v_�˷ѽ���:='�ֽ�';
        End;
    End IF;

	--�Ƿ��״�ȫ��:��һ����������ʣ������Ϊ��Ϊ�ж�����׼����=ʣ����
	Select Decode(Count(*),0,1,0) Into v_�״�ȫ�� From ���˷��ü�¼ Where ��¼����=1 And ��¼״̬=2 And NO=NO_IN;
	If v_�״�ȫ��=1 Then
		Select
			Decode(Count(A.���),0,1,0) Into v_�״�ȫ��
		From (
			Select --ÿ��ʣ��������׼������
				A.���,A.ʣ������,Decode(A.ִ��״̬,1,0,
					Decode(Instr(',4,5,6,7,',A.�շ����),0,A.ʣ������,Nvl(B.׼������,A.ʣ������))) As ׼������--@@@
			From (
				Select --��ʣ��������ÿ��ʣ��������ԭʼ����ID��ִ��״̬
					Sum(A.ID) As ID,Sum(A.ִ��״̬) As ִ��״̬,A.���,A.�շ����,Sum(����) As ʣ������
				From (
					Select
						Decode(A.��¼״̬,2,0,A.ID) As ID,
						Decode(A.��¼״̬,2,0,Nvl(A.ִ��״̬,0)) As ִ��״̬,
						A.���,A.�շ����,Nvl(A.����,1)*A.���� As ����
					From ���˷��ü�¼ a
					Where A.�۸񸸺� Is Null And Nvl(A.���ӱ�־,0)<>9
						And A.��¼����=1 And A.NO=NO_IN
					) A
				Group By A.���,A.�շ���� Having Nvl(Sum(����),0)<>0
				) A,(
				Select --ҩƷ׼������
					����ID,Sum(Nvl(����,1)*ʵ������) As ׼������
				From ҩƷ�շ���¼
				Where NO=NO_IN And Mod(��¼״̬,3)=1
					And ����� Is Null And ���� IN(8,24)--@@@
				Group By ����ID
				) B
			Where A.ID=B.����ID(+)) A
		Where Nvl(A.׼������,0)<>A.ʣ������;
	End IF;

    --ѭ������ÿ�з���(������Ŀ��)
    v_�ܽ��:=0;
    v_�����˷�:=1;
    For r_Bill IN c_Bill Loop
        IF INSTR(','||���_IN||',',','||Nvl(r_Bill.�۸񸸺�,r_Bill.���)||',') >0 Or ���_IN Is Null Then
			If ���_IN IS NULL And Nvl(r_Bill.���ӱ�־,0)=9 And v_�״�ȫ��=0 Then
				--���ǵ�һ���˷ѵ�ȫ��ʱ,������������
				v_�����˷�:=0;--Ӧ���ǲ����˷�
            ElsIF Nvl(r_Bill.ִ��״̬,0)<>1 Then
                --��ʣ������,ʣ��Ӧ��,ʣ��ʵ��
                Select
                    Sum(Nvl(����,1)*����),Sum(Ӧ�ս��),Sum(ʵ�ս��),Sum(ͳ����)
                    Into v_ʣ������,v_ʣ��Ӧ��,v_ʣ��ʵ��,v_ʣ��ͳ��
                From ���˷��ü�¼
                Where NO=NO_IN And ��¼����=1 And ���=r_Bill.���;

                IF v_ʣ������=0 Then
                    IF ���_IN IS Not NULL Then
                        v_Error := '�����е�'||Nvl(r_Bill.�۸񸸺�,r_Bill.���)||'�з����Ѿ�ȫ���˷ѣ�';
                        Raise Err_Custom;
                    End IF;
                    --�����δ�޶��к�,ԭʼ�����еĸñ��Ѿ�ȫ���˷�(ִ��״̬=0��һ�ֿ���)
                    v_�����˷�:=0;
                Else
                    --׼������(��ҩƷ��ĿΪʣ������,ԭʼ����)
                    IF Instr(',4,5,6,7,',r_Bill.�շ����)=0 Then--@@@
                        v_׼������:=v_ʣ������;
                    Else
                        Select Sum(Nvl(����,1)*ʵ������) Into v_׼������
                        From ҩƷ�շ���¼
                        Where NO=NO_IN And ���� IN(8,24) And Mod(��¼״̬,3)=1 --@@@
                            And ����� is NULL And ����ID=r_Bill.ID;

						--���������õ���������@@@
						--��ʣ��������׼�������������������
							--1.�޶�Ӧ���շ���¼(����ͨ���û򲻸������õ�����),��ʱӦ��ʣ������
							--2.�շ���¼����ȫ������,����ȫ��ִ��,���ų��������
						If r_Bill.�շ����='4' And Nvl(v_׼������,0)=0 Then
							v_׼������:=v_ʣ������;
						End IF;
                    End if;
                    --�Ƿ񲿷��˷�
                    If r_Bill.ִ��״̬=2 Or v_׼������<>Nvl(r_Bill.����,1)*r_Bill.���� Then
                        v_�����˷�:=0;
                    End IF;

                    --�����˷��ü�¼

                    --�ñ���Ŀ�ڼ����˷�
                    Select Nvl(Max(Abs(ִ��״̬)),0)+1 Into v_�˷Ѵ���
                    From ���˷��ü�¼
                    Where NO=NO_IN And ��¼����=1 And ��¼״̬=2 And ���=r_Bill.���;

                    --���=ʣ����*(׼����/ʣ����)
                    v_Ӧ�ս��:=Round(v_ʣ��Ӧ��*(v_׼������/v_ʣ������),v_Dec);
                    v_ʵ�ս��:=Round(v_ʣ��ʵ��*(v_׼������/v_ʣ������),v_Dec);
                    v_ͳ����:=Round(v_ʣ��ͳ��*(v_׼������/v_ʣ������),v_Dec);
                    v_�ܽ��:=v_�ܽ��+v_ʵ�ս��;

                    --�����˷Ѽ�¼
                    Insert Into ���˷��ü�¼(
                        ID,NO,ʵ��Ʊ��,��¼����,��¼״̬,���,��������,�۸񸸺�,����ID,��ҳID,ҽ�����,�����־,����,
                        �Ա�,����,��ʶ��,����,�ѱ�,���˲���ID,���˿���ID,�շ����,�շ�ϸĿID,���㵥λ,����,��ҩ����,
                        ����,�Ӱ��־,���ӱ�־,������ĿID,�վݷ�Ŀ,���ʷ���,��׼����,Ӧ�ս��,ʵ�ս��,��������ID,
                        ������,ִ�в���ID,������,ִ����,ִ��״̬,ִ��ʱ��,����Ա���,����Ա����,����ʱ��,�Ǽ�ʱ��,
                        ����ID,���ʽ��,������Ŀ��,���մ���ID,ͳ����,ժҪ,�Ƿ��ϴ�)
                    Select ���˷��ü�¼_ID.Nextval,NO,ʵ��Ʊ��,��¼����,2,���,��������,�۸񸸺�,����ID,��ҳID,
                        ҽ�����,�����־,����,�Ա�,����,��ʶ��,����,�ѱ�,���˲���ID,���˿���ID,�շ����,
                        �շ�ϸĿID,���㵥λ,Decode(Sign(v_׼������-Nvl(����,1)*����),0,����,1),��ҩ����,
                        Decode(Sign(v_׼������-Nvl(����,1)*����),0,-1*����,-1*v_׼������),�Ӱ��־,���ӱ�־,
                        ������ĿID,�վݷ�Ŀ,���ʷ���,��׼����,-1*v_Ӧ�ս��,-1*v_ʵ�ս��,��������ID,������,ִ�в���ID,
                        ������,ִ����,-1*v_�˷Ѵ���,ִ��ʱ��,����Ա���_IN,����Ա����_IN,����ʱ��,v_Date,v_����ID,
                        -1*v_ʵ�ս��,������Ŀ��,���մ���ID,-1*v_ͳ����,ժҪ,Decode(Nvl(���ӱ�־,0),9,1,0)
                    From ���˷��ü�¼ Where ID=r_Bill.ID;

                    --�����˷��û���
                    Update ���˷��û���
                        Set Ӧ�ս��=Nvl(Ӧ�ս��,0) - v_Ӧ�ս��,
                            ʵ�ս��=Nvl(ʵ�ս��,0) - v_ʵ�ս��,
                            ���ʽ��=Nvl(���ʽ��,0) - v_ʵ�ս��
                     Where ����=Trunc(v_Date)
                        And Nvl(���˲���ID,0)=Nvl(r_Bill.���˲���ID,0)
                        And Nvl(���˿���ID,0)=Nvl(r_Bill.���˿���ID,0)
                        And Nvl(��������ID,0)=Nvl(r_Bill.��������ID,0)
                        And Nvl(ִ�в���ID,0)=Nvl(r_Bill.ִ�в���ID,0)
                        And ������ĿID+0=r_Bill.������ĿID
                        And ��Դ;��=r_Bill.�����־ And ���ʷ���=0;
                    IF SQL%RowCount=0 Then
                        Insert Into ���˷��û���(
                            ����,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,��Դ;��,���ʷ���,Ӧ�ս��,ʵ�ս��,���ʽ��)
                        Values(
                            Trunc(v_Date),r_Bill.���˲���ID,r_Bill.���˿���ID,r_Bill.��������ID,r_Bill.ִ�в���ID,
                            r_Bill.������ĿID,r_Bill.�����־,0,-1 * v_Ӧ�ս��,-1 * v_ʵ�ս��,-1 * v_ʵ�ս��);
                    End IF;

                    --���ԭ���ü�¼
                    --ִ��״̬:ȫ������(׼����=ʣ����)���Ϊ0,������Ϊ1
                    Update ���˷��ü�¼
                        Set ��¼״̬=3,
                            ִ��״̬=Decode(Sign(v_׼������-v_ʣ������),0,0,1)
                    Where ID=r_Bill.ID;
                End IF;
            Else
                IF ���_IN Is Not Null Then
                    v_Error := '�����е�'||Nvl(r_Bill.�۸񸸺�,r_Bill.���)||'�з����Ѿ���ȫִ��,�����˷ѣ�';
                    Raise Err_Custom;
                End IF;
                --���:û�޶��к�,ԭʼ�����а����Ѿ���ȫִ�е�
                v_�����˷�:=0;
            End IF;
        Else
            v_�����˷�:=0;--δָ���ñ�,���ڲ����˷�
        End IF;
    End Loop;

    ---------------------------------------------------------------------------------
    --������Ԥ����¼

    --ԭ���ݵĽ���ID
    Select ����ID Into v_Count From ���˷��ü�¼ Where NO=NO_IN And ��¼����=1 And ��¼״̬ IN(1,3) And Rownum=1;

    IF v_�����˷�=1 THEN --���ݵ�һ���˷���ȫ������

		--��Ԥ�����ּ�¼
        Insert Into ����Ԥ����¼(
            ID,NO,ʵ��Ʊ��,��¼����,��¼״̬,����ID,��ҳID,����ID,���,���㷽ʽ,�������,ժҪ,
            �ɿλ,��λ������,��λ�ʺ�,�տ�ʱ��,����Ա����,����Ա���,��Ԥ��,����ID)
        Select
            ����Ԥ����¼_ID.Nextval,NO,ʵ��Ʊ��,11,��¼״̬,����ID,��ҳID,����ID,Null,
            ���㷽ʽ,�������,ժҪ,�ɿλ,��λ������,��λ�ʺ�,�տ�ʱ��,����Ա����,
            ����Ա���,-1*��Ԥ��,v_����ID
        From ����Ԥ����¼
        Where ��¼���� IN(1,11) And ����ID=v_Count;

        --������Ԥ�����
        Begin
            Select ����ID,Sum(Nvl(��Ԥ��,0)) Into v_����ID,v_Ԥ����� From ����Ԥ����¼
            Where ��¼���� IN(1,11) And ����ID=v_Count
            Group by ����ID;
        Exception
            When Others Then NULL;
        End;
        IF Nvl(v_����ID,0)<>0 And Nvl(v_Ԥ�����,0)<>0 Then
            Update ������� Set Ԥ�����=Nvl(Ԥ�����,0)+v_Ԥ����� Where ����ID=v_����ID And ����=1;
            IF SQL%RowCount=0 Then
                Insert Into �������(����ID,Ԥ�����,����) Values(v_����ID,v_Ԥ�����,1);
            End IF;
            Delete From ������� Where ����ID=v_����ID And ����=1 And Nvl(Ԥ�����,0)=0 And Nvl(�������,0)=0;
        End IF;

        --��ҽ��ȫ��,��ҽ�����н��㷽ʽ���������,ԭ���˻�(��Ԥ����ǰ���Ѵ���)
        IF ҽ�����㷽ʽ_IN Is Null Then
            Insert Into ����Ԥ����¼(
                ID,��¼����,NO,��¼״̬,����ID,��ҳID,ժҪ,���㷽ʽ,�������,�տ�ʱ��,�ɿλ,��λ������,��λ�ʺ�,����Ա���,����Ա����,��Ԥ��,����ID)
            Select
                ����Ԥ����¼_ID.Nextval,��¼����,NO,2,����ID,��ҳID,ժҪ,���㷽ʽ,�������,v_Date,�ɿλ,��λ������,��λ�ʺ�,
                ����Ա���_IN,����Ա����_IN,-1*��Ԥ��,v_����ID
            From ����Ԥ����¼
            Where ��¼����=3 And ��¼״̬=1 And ����ID=v_Count;

        --ҽ�����������ϵĽ��㷽ʽ��,�������,�˵�ָ���Ľ��㷽ʽ��
        Else
            --a.ԭ���˻�
            v_��������:=','||ҽ�����㷽ʽ_IN ||','||v_�˷ѽ���||',' ;           
            Insert Into ����Ԥ����¼(
                ID,��¼����,NO,��¼״̬,����ID,��ҳID,ժҪ,���㷽ʽ,�������,�տ�ʱ��,����Ա���,����Ա����,��Ԥ��,����ID)
            Select
                ����Ԥ����¼_ID.Nextval,��¼����,NO,2,����ID,��ҳID,ժҪ,���㷽ʽ,�������,v_Date,����Ա���_IN,����Ա����_IN,-1*��Ԥ��,v_����ID
            From ����Ԥ����¼
            Where ��¼����=3 And ��¼״̬=1 And ����ID=v_Count And instr(v_��������,','||���㷽ʽ||',')=0;
               
            --b.���µľ���ҽ�����������ϵĽ��㷽ʽ,���ϵ�ָ���Ľ��㷽ʽ��,�������(��Ϊ������������֮�������)          
            Begin
                Select -1*Nvl(Sum(��Ԥ��),0) Into v_���˽�� From ����Ԥ����¼ Where ����ID=v_����Id;
            Exception
                When Others Then  v_���˽��:=0;
            End;
            IF (v_�ܽ��-v_���˽��)<>0 Then             --��ʱ���ܽ�û�а������,��Ϊ����������ڵ��ñ����̺�Ų��������ü�¼
                Insert Into ����Ԥ����¼(
                    ID,��¼����,NO,��¼״̬,����ID,��ҳID,ժҪ,���㷽ʽ,�տ�ʱ��,����Ա���,����Ա����,��Ԥ��,����ID)
                Select
                    ����Ԥ����¼_ID.Nextval,3,NO,2,����ID,��ҳID,'����ҽ�������˷�',v_�˷ѽ���,v_Date,����Ա���_IN,
                    ����Ա����_IN,-1*(v_�ܽ��-v_���˽��+Nvl(���_IN,0)),v_����ID
                From ����Ԥ����¼
                Where ��¼����=3 And ��¼״̬=1 And ����ID=v_Count And Rownum=1;
            End IF;
        End IF;      
    Else
        -------------------------------------------------
        --�����˷�ֱ����Ϊָ�����㷽ʽ
        Insert Into ����Ԥ����¼(
            ID,��¼����,NO,��¼״̬,����ID,��ҳID,ժҪ,���㷽ʽ,�տ�ʱ��,�ɿλ,��λ������,��λ�ʺ�,����Ա���,����Ա����,��Ԥ��,����ID)
        Select
            ����Ԥ����¼_ID.Nextval,3,NO_IN,2,����ID,��ҳID,'�����˷ѽ���',v_�˷ѽ���,v_Date,NULL,NULL,NULL,
            ����Ա���_IN,����Ա����_IN,-1*(v_�ܽ��+Nvl(���_IN,0)),v_����ID
        From ����Ԥ����¼
        Where ��¼����=3 And ��¼״̬ IN(1,3) And ����ID=v_Count And Rownum=1;
    End IF;

    --����ԭ��¼
    Update ����Ԥ����¼ Set ��¼״̬=3 Where ��¼����=3 And ��¼״̬ IN(1,3) And ����ID=v_Count;

    ---------------------------------------------------------------------------------
    --��Ա�ɿ����(ע����Ԥ����¼�����Ŵ������������ʻ��ȵĽ�����,�����˳�Ԥ����)
    For r_MoneyRow in c_Money(v_����ID) Loop
        Update ��Ա�ɿ����
            Set ���=Nvl(���,0)+r_MoneyRow.��Ԥ��
         Where �տ�Ա=����Ա����_IN And ����=1 And ���㷽ʽ=r_MoneyRow.���㷽ʽ;
        IF SQL%RowCount=0 Then
            Insert Into ��Ա�ɿ����(
                �տ�Ա,���㷽ʽ,����,���)
            Values(
                ����Ա����_IN,r_MoneyRow.���㷽ʽ,1,r_MoneyRow.��Ԥ��);
        End IF;
        Delete From ��Ա�ɿ���� Where �տ�Ա=����Ա����_IN And ����=1 And ���㷽ʽ=r_MoneyRow.���㷽ʽ And Nvl(���,0)=0;
    End Loop;

    ---------------------------------------------------------------------------------
    --�շ�Ʊ�ݴ���
    --��ȡ�������һ�εĴ�ӡID(�����Ƕ��ŵ����շѴ�ӡ)
    --������ǰû�д�ӡ����,��ʱ�����ȫ������,�����ջأ���δȫ������ʱ��Ҫ�ش�
    Begin
        Select Max(ID) Into v_��ӡID From Ʊ�ݴ�ӡ���� Where ��������=1 And NO=NO_IN;
    Exception
        When Others Then NULL;
    End;

	If Nvl(Ʊ�ݴ���_IN,0)=0 Then
		--�ж���ȫ��������,�Ծ����Ƿ��ش�Ʊ�ݡ�
		--����ִ��״̬=1�����(�����β����˲���),���ŵ��ݿ�
		--�൥���շ�ʱ,�������е��ݶ������˲Ų��ش�
		Select Nvl(Count(*),0) Into v_Count
		From (
			Select NO,���,Sum(����) as ʣ������
			From (
				Select NO,��¼״̬,Nvl(�۸񸸺�,���) as ���,
					Avg(Nvl(����,1)*����) as ����
				From ���˷��ü�¼
				Where ��¼����=1 And Nvl(���ӱ�־,0)<>9
					And NO IN(
						Select NO From Ʊ�ݴ�ӡ���� Where ID=v_��ӡID And ��������=1
						Union ALL
						Select NO_IN From Dual)
				Group by NO,��¼״̬,Nvl(�۸񸸺�,���)
				)
			Group by NO,��� Having Sum(����)<>0);
		If v_Count=0 Then
			v_ȫ������:=1;
		Else
			v_ȫ������:=0;
		End IF;

		IF v_ȫ������=1 Then
			--ȫ�������ջض���Ʊ��(������ǰû�д�ӡ,���ջ�)
			If v_��ӡID IS Not NULL Then
				Insert Into Ʊ��ʹ����ϸ(
					ID,Ʊ��,����,����,ԭ��,����ID,��ӡID,ʹ��ʱ��,ʹ����)
				Select
					Ʊ��ʹ����ϸ_ID.Nextval,Ʊ��,����,2,2,����ID,��ӡID,v_Date,����Ա����_IN
				From Ʊ��ʹ����ϸ
				Where ��ӡID=v_��ӡID And Ʊ��=1 And ����=1;
			End IF;
		Else
			--�ǲ����˷�,�ش�Ʊ��
			If Ʊ�ݺ�_IN is Not NULL Then
				--���ڿ��Դ�,����ǰ���û������ھ����ջ�
				zl_�����շѼ�¼_RePrint(NO_IN,Ʊ�ݺ�_IN,����ID_IN,����Ա����_IN,v_Date,1);
			ElsIf v_��ӡID IS Not NULL Then
				--���ڲ���ӡ,��Ҫ�ջ���ǰ��
				Insert Into Ʊ��ʹ����ϸ(
					ID,Ʊ��,����,����,ԭ��,����ID,��ӡID,ʹ��ʱ��,ʹ����)
				Select
					Ʊ��ʹ����ϸ_ID.Nextval,Ʊ��,����,2,2,����ID,��ӡID,v_Date,����Ա����_IN
				From Ʊ��ʹ����ϸ
				Where ��ӡID=v_��ӡID And Ʊ��=1 And ����=1;
			End IF;
		End If;
	ElsIf Nvl(Ʊ�ݴ���_IN,0)=1 Then
		--���ŵ���ȫ�ˣ�������Ʊ��(���ŵ���ֻ���ջ�һ��)
		If v_��ӡID IS Not NULL Then
			Select Count(*) Into v_Count From Ʊ��ʹ����ϸ Where Ʊ��=1 And ����=2 And ��ӡID=v_��ӡID;
			If v_Count=0 Then
				Insert Into Ʊ��ʹ����ϸ(
					ID,Ʊ��,����,����,ԭ��,����ID,��ӡID,ʹ��ʱ��,ʹ����)
				Select
					Ʊ��ʹ����ϸ_ID.Nextval,Ʊ��,����,2,2,����ID,��ӡID,v_Date,����Ա����_IN
				From Ʊ��ʹ����ϸ
				Where ��ӡID=v_��ӡID And Ʊ��=1 And ����=1;
			End IF;
		End IF;
	ElsIf Nvl(Ʊ�ݴ���_IN,0)=2 Then
		NULL;--���ŵ��ݲ����ˣ���ʱ������Ʊ��
	End IF;

    ---------------------------------------------------------------------------------
    --ҩƷ�������
    For r_Stock in c_Stock Loop
        --����ҩƷ���
        If r_Stock.�ⷿID IS Not NULL Then
            Update ҩƷ���
                Set ��������=Nvl(��������,0)+Nvl(r_Stock.����,1)*Nvl(r_Stock.ʵ������,0)
             Where �ⷿID=r_Stock.�ⷿID And ҩƷID=r_Stock.ҩƷID
                And Nvl(����,0)=Nvl(r_Stock.����,0) And ����=1;
            IF SQL%RowCount=0 Then
                Insert Into ҩƷ���(
                    �ⷿID,ҩƷID,����,����,Ч��,��������,�ϴ�����,�ϴβ���,���Ч��)--@@@
                Values(
                    r_Stock.�ⷿID,r_Stock.ҩƷID,1,r_Stock.����,r_Stock.Ч��,
                    Nvl(r_Stock.����,1)*Nvl(r_Stock.ʵ������,0),r_Stock.����,r_Stock.����,r_Stock.���Ч��);
            End IF;
        End IF;

        --ɾ��ҩƷ�շ���¼
        Delete From ҩƷ�շ���¼ Where ID=r_Stock.ID;
    End Loop;

    --δ��ҩƷ��¼
    For r_Spare IN c_Spare Loop
        Select Nvl(Count(*),0) Into v_Count
        From ҩƷ�շ���¼
        Where NO=NO_IN And ����=r_Spare.���� --@@@
			And Mod(��¼״̬,3)=1 And ����� is NULL
			And Nvl(�ⷿID,0)=Nvl(r_Spare.�ⷿID,0);

        If v_Count=0 Then
            Delete From δ��ҩƷ��¼ Where ����=r_Spare.���� --@@@
				And NO=NO_IN And Nvl(�ⷿID,0)=Nvl(r_Spare.�ⷿID,0);
        End IF;
    End Loop;

	--���ŵ���ȫ������ʱ��ɾ������ҽ������
	If ���_IN IS NULL Then
		Begin
			Select ҽ����� Into v_ҽ��ID From ���˷��ü�¼ Where ��¼����=1 And ��¼״̬=3 And NO=NO_IN And Rownum=1;
		Exception
			When Others Then NULL;
		End;
		If v_ҽ��ID IS Not NULL Then
			Select Nvl(Count(*),0) Into v_Count
			From (
				Select ���,Sum(����) as ʣ������
				From (
					Select ��¼״̬,Nvl(�۸񸸺�,���) as ���,
						Avg(Nvl(����,1)*����) as ����
					From ���˷��ü�¼
					Where ��¼����=1 And Nvl(���ӱ�־,0)<>9
						And ҽ�����+0=v_ҽ��ID And NO=NO_IN
					Group by ��¼״̬,Nvl(�۸񸸺�,���)
					)
				Group by ��� Having Sum(����)<>0);
			IF v_Count = 0 Then
				Delete From ����ҽ������ Where ҽ��ID=v_ҽ��ID And ��¼����=1 And NO=NO_IN;
			End IF;
		End IF;
	End IF;
Exception
    When Err_Custom Then Raise_Application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When Others Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_�����շѼ�¼_Delete;
/

-------------------------------------------------------
--ģ�飺סԺ���ʼ�¼.SQL
Create Or Replace Procedure zl_סԺ���ʼ�¼_Insert(
    NO_IN			���˷��ü�¼.NO%Type,
    ���_IN         ���˷��ü�¼.���%Type,
    ����ID_IN       ���˷��ü�¼.����ID%Type,
    ��ҳID_IN       ���˷��ü�¼.��ҳID%Type,
    ��ʶ��_IN       ���˷��ü�¼.��ʶ��%Type,
    ����_IN         ���˷��ü�¼.����%Type,
    �Ա�_IN         ���˷��ü�¼.�Ա�%Type,
    ����_IN         ���˷��ü�¼.����%Type,
    ����_IN         ���˷��ü�¼.����%Type,
    �ѱ�_IN         ���˷��ü�¼.�ѱ�%Type,
    ����ID_IN       ���˷��ü�¼.���˲���ID%Type,
    ����ID_IN       ���˷��ü�¼.���˿���ID%Type,
    �Ӱ��־_IN     ���˷��ü�¼.�Ӱ��־%Type,
    Ӥ����_IN       ���˷��ü�¼.Ӥ����%Type,
    ��������ID_IN   ���˷��ü�¼.��������ID%Type,
    ������_IN       ���˷��ü�¼.������%Type,
    ��������_IN     ���˷��ü�¼.��������%Type,
    �շ�ϸĿID_IN   ���˷��ü�¼.�շ�ϸĿID%Type,
    �շ����_IN     ���˷��ü�¼.�շ����%Type,
    ���㵥λ_IN     ���˷��ü�¼.���㵥λ%Type,
    ������Ŀ��_IN   ���˷��ü�¼.������Ŀ��%Type,
    ���մ���ID_IN   ���˷��ü�¼.���մ���ID%Type,
    ���ձ���_IN     ���˷��ü�¼.���ձ���%Type,
    ����_IN         ���˷��ü�¼.����%Type,
    ����_IN         ���˷��ü�¼.����%Type,
    ���ӱ�־_IN     ���˷��ü�¼.���ӱ�־%Type,
    ִ�в���ID_IN   ���˷��ü�¼.ִ�в���ID%Type,
    �۸񸸺�_IN     ���˷��ü�¼.�۸񸸺�%Type,
    ������ĿID_IN   ���˷��ü�¼.������ĿID%Type,
    �վݷ�Ŀ_IN     ���˷��ü�¼.�վݷ�Ŀ%Type,
    ��׼����_IN     ���˷��ü�¼.��׼����%Type,
    Ӧ�ս��_IN     ���˷��ü�¼.Ӧ�ս��%Type,
    ʵ�ս��_IN     ���˷��ü�¼.ʵ�ս��%Type,
    ͳ����_IN     ���˷��ü�¼.ͳ����%Type,
    ����ʱ��_IN     ���˷��ü�¼.����ʱ��%Type,
    �Ǽ�ʱ��_IN     ���˷��ü�¼.�Ǽ�ʱ��%Type,
    ҩƷժҪ_IN     ҩƷ�շ���¼.ժҪ%Type,
    ����_IN         Number,
    ����Ա���_IN   ���˷��ü�¼.����Ա���%Type,
    ����Ա����_IN   ���˷��ü�¼.����Ա����%Type,
    �ಡ�˵�_IN     Number := 0,
    ���ID_IN       ҩƷ��������.���ID%Type:=Null,
    ���ʵ�ID_IN     ���˷��ü�¼.���ʵ�ID%Type:=Null,
    ����ժҪ_IN     ���˷��ü�¼.ժҪ%Type:=Null,
    �Ƿ���_IN     ���˷��ü�¼.�Ƿ���%Type:=0,
    ҽ�����_IN     ���˷��ü�¼.ҽ�����%TYPE:=NULL,
    Ƶ��_IN         ҩƷ�շ���¼.Ƶ��%Type:=NULL,
    ����_IN         ҩƷ�շ���¼.����%Type:=NULL,
    �÷�_IN         ҩƷ�շ���¼.�÷�%Type:=NULL,--�÷�[|�巨]
    ��Ч_IN         ҩƷ�շ���¼.����%Type:=NULL,
    �Ƽ�����_IN     ҩƷ�շ���¼.����%Type:=NULL,
    �򵥼���_IN     Number:=0
)
AS
    --���ܣ�����һ��סԺ���ʵ���
    --������
    --   ҩƷժҪ_IN:���ҽ���еĸ���˵�����޸ı����µ���ʱ�á�Ŀǰ�����ڴ����ҩƷ�շ���¼��ժҪ�С�
    --         ԭ����(��¼״̬=2)��¼�޸Ĳ������µ��ݺš�
    --         �µ���(��¼״̬=1)��¼���޸ĵ�ԭ���ݺš�
    --   ����-�Ƿ�����סԺ���ۡ�
    v_����ID ���˷��ü�¼.ID%Type;
    v_���ȼ� δ��ҩƷ��¼.���ȼ�%Type;

    --ҩ��������ʱ��ҩƷ--
    ------------------------------------------------------------
    --���α����ڷ���ҩƷ�����ֽ�
    Cursor c_Stock is
        Select * From ҩƷ��� 
        Where ҩƷID=�շ�ϸĿID_IN And �ⷿID=ִ�в���ID_IN
            And ����=1 And(Nvl(����,0)=0 Or Ч�� is Null Or Ч��>Trunc(Sysdate))
            And Nvl(��������,0)<>0
        Order By Nvl(����,0);
    r_Stock c_Stock%RowType;
    
    --����
    v_����			ҩƷ���.ҩ������%Type;
    v_ʱ��			�շ���ĿĿ¼.�Ƿ���%Type;
    v_����			�շ���ĿĿ¼.����%Type;
    --��ʱ����
    v_������		Number;
    v_��ǰ����		Number;
    v_�ܽ��		Number;
    v_��ǰ����		Number;
    --ҩƷ�շ���¼
    v_����			ҩƷ�շ���¼.����%Type;
    v_����			ҩƷ�շ���¼.����%Type;
    v_����			ҩƷ�շ���¼.����%Type;
    v_Ч��			ҩƷ�շ���¼.Ч��%Type;
    v_���			ҩƷ�շ���¼.���%Type;
    v_����			ҩƷ�շ���¼.����%Type;
	v_���Ч��		ҩƷ�շ���¼.���Ч��%Type;
	v_�������		ҩƷ�շ���¼.�������%Type;
    ------------------------------------------------------------
	v_�÷�			ҩƷ�շ���¼.�÷�%Type;
	v_�巨			ҩƷ�շ���¼.���%Type;

	v_Dec			Number;
	v_Count			Number;
	v_Error			Varchar2(255);
    Err_custom		Exception;
Begin
	--ҩƷ�÷��巨�ֽ�
	IF �÷�_IN IS Not NULL Then
		IF Instr(�÷�_IN,'|')>0 Then
			v_�÷�:=Substr(�÷�_IN,1,Instr(�÷�_IN,'|')-1);
			v_�巨:=Substr(�÷�_IN,Instr(�÷�_IN,'|')+1);
		Else
			v_�÷�:=�÷�_IN;
		End IF;
	End IF;

    --���С��λ��
    Begin
        Select To_Number(����ֵ) Into v_Dec From ϵͳ������ Where ������=9;
    Exception
        When Others Then v_Dec:=2;
    End;

    --���˷��ü�¼
    Select ���˷��ü�¼_ID.Nextval Into v_����ID From Dual;

    Insert Into ���˷��ü�¼(
        ID,��¼����,NO,��¼״̬,���,��������,�۸񸸺�,�ಡ�˵�,�����־,����ID,��ҳID,
        ��ʶ��,����,�Ա�,����,����,���˲���ID,���˿���ID,�ѱ�,�շ����,�շ�ϸĿID,���㵥λ,
        ������Ŀ��,���մ���ID,���ձ���,��ҩ����,����,����,�Ӱ��־,���ӱ�־,Ӥ����,������ĿID,�վݷ�Ŀ,
        ��׼����,Ӧ�ս��,ʵ�ս��,ͳ����,���ʷ���,��������ID,������,����ʱ��,�Ǽ�ʱ��,
        ִ�в���ID,ִ��״̬,������,����Ա���,����Ա����,���ʵ�ID,ժҪ,�Ƿ���,ҽ�����)
    Values(
        v_����ID,2,NO_IN,Decode(����_IN,1,0,1),���_IN,Decode(��������_IN,0,Null,��������_IN),
        Decode(�۸񸸺�_IN,0,Null,�۸񸸺�_IN),�ಡ�˵�_IN,2,����ID_IN,��ҳID_IN,
        Decode(��ʶ��_IN,0,Null,��ʶ��_IN),����_IN,�Ա�_IN,����_IN,
        Decode(����_IN,0,Null,����_IN),Decode(����ID_IN,0,Null,����ID_IN),
        Decode(����ID_IN,0,Null,����ID_IN),�ѱ�_IN,�շ����_IN,�շ�ϸĿID_IN,
        ���㵥λ_IN,������Ŀ��_IN,���մ���ID_IN,���ձ���_IN,Decode(Nvl(�򵥼���_IN,0),0,NULL,�շ����_IN),
        ����_IN,����_IN,�Ӱ��־_IN,���ӱ�־_IN,Ӥ����_IN,������ĿID_IN,�վݷ�Ŀ_IN,��׼����_IN,Ӧ�ս��_IN,
        ʵ�ս��_IN,ͳ����_IN,1,��������ID_IN,������_IN,����ʱ��_IN,�Ǽ�ʱ��_IN,
        ִ�в���ID_IN,0,����Ա����_IN,Decode(����_IN,1,Null,����Ա���_IN),
        Decode(����_IN,1,Null,����Ա����_IN),���ʵ�ID_IN,����ժҪ_IN,�Ƿ���_IN,ҽ�����_IN);

    --��ػ��ܱ�Ĵ���
	If Nvl(����_IN,0)=0 Then
		--�������
		Update �������
			Set �������=Nvl(�������,0)+ʵ�ս��_IN
		 Where ����ID=����ID_IN And ����=1;
		IF SQL%RowCount=0 Then
			Insert Into �������(
				����ID,����,�������,Ԥ�����)
			Values(
				����ID_IN,1,ʵ�ս��_IN,0);
		End IF;

		--����δ�����
		Update ����δ�����
			Set ���=Nvl(���,0)+ʵ�ս��_IN
		 Where ����ID=����ID_IN
			And Nvl(��ҳID,0)=Nvl(��ҳID_IN,0)
			And Nvl(���˲���ID,0)=Nvl(����ID_IN,0)
			And Nvl(���˿���ID,0)=Nvl(����ID_IN,0)
			And Nvl(��������ID,0)=Nvl(��������ID_IN,0)
			And Nvl(ִ�в���ID,0)=Nvl(ִ�в���ID_IN,0)
			And ������ĿID+0=������ĿID_IN
			And ��Դ;��+0=2;
		IF SQL%RowCount=0 Then
			Insert Into ����δ�����(
				����ID,��ҳID,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,��Դ;��,���)
			Values(
				����ID_IN,��ҳID_IN,����ID_IN,����ID_IN,��������ID_IN,ִ�в���ID_IN,������ĿID_IN,2,ʵ�ս��_IN);
		End IF;

		--���˷��û���
		Update ���˷��û���
			Set Ӧ�ս��=Nvl(Ӧ�ս��,0)+Ӧ�ս��_IN,
				 ʵ�ս��=Nvl(ʵ�ս��,0)+ʵ�ս��_IN
		 Where ����=Trunc(�Ǽ�ʱ��_IN)
			And Nvl(���˲���ID,0)=Nvl(����ID_IN,0)
			And Nvl(���˿���ID,0)=Nvl(����ID_IN,0)
			And Nvl(��������ID,0)=Nvl(��������ID_IN,0)
			And Nvl(ִ�в���ID,0)=Nvl(ִ�в���ID_IN,0)
			And ������ĿID+0=������ĿID_IN
			And ��Դ;��=2 And ���ʷ���=1;
		IF SQL%RowCount=0 Then
			Insert Into ���˷��û���(
				����,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,��Դ;��,���ʷ���,Ӧ�ս��,ʵ�ս��,���ʽ��)
			Values(
				Trunc(�Ǽ�ʱ��_IN),����ID_IN,����ID_IN,��������ID_IN,ִ�в���ID_IN,������ĿID_IN,2,1,Ӧ�ս��_IN,ʵ�ս��_IN,0);
		End IF;
	End IF;

    --ҩƷ���������ϲ���
	v_Count:=0;--@@@
	If �շ����_IN='4' Then--�������õ����ĲŴ���
		Select �������� Into v_Count From �������� Where ����ID=�շ�ϸĿID_IN;
	End IF;
    IF �շ����_IN in('5','6','7') Or (�շ����_IN='4' And Nvl(v_Count,0)=1) Then
		If �շ����_IN='4' Then
			Select Nvl(A.���÷���,0),Nvl(B.�Ƿ���,0),B.���� 
				Into v_����,v_ʱ��,v_����
			From �������� A,�շ���ĿĿ¼ B
			Where A.����ID=B.ID And B.ID=�շ�ϸĿID_IN;
		Else
			Select Nvl(A.ҩ������,0),Nvl(B.�Ƿ���,0),B.���� 
				Into v_����,v_ʱ��,v_����
			From ҩƷ��� A,�շ���ĿĿ¼ B
			Where A.ҩƷID=B.ID And B.ID=�շ�ϸĿID_IN;
		End IF;

        v_������:=����_IN*����_IN;
        v_�ܽ��:=0;
        Open c_Stock;

        While v_������<>0 Loop
            Fetch c_Stock Into r_Stock;
            IF c_Stock%NotFound Then
                --��һ�ξ�û�п��,������ʱ�۶�������
                --����ҩƷ�����ֽⲻ��,Ҳ���ǿ�治�㡣
                IF v_����=1 Or v_ʱ��=1 Then
                    Close c_Stock;
                    If ҽ�����_IN IS NULL Then
						If �շ����_IN='4' Then
							v_Error:='�� '||���_IN||' �еķ�����ʱ����������"'||v_����||'"û���㹻�Ĳ��Ͽ�棡';
						Else
	                        v_Error:='�� '||���_IN||' �еķ�����ʱ��ҩƷ"'||v_����||'"û���㹻��ҩƷ��棡';
						End IF;
                    Else
						If �շ����_IN='4' Then
							v_Error:='�ڴ�����"'||����_IN||'"ʱ���ַ�����ʱ����������"'||v_����||'"û���㹻�Ĳ��Ͽ�棡';
						Else
	                        v_Error:='�ڴ�����"'||����_IN||'"ʱ���ַ�����ʱ��ҩƷ"'||v_����||'"û���㹻��ҩƷ��棡';
						End IF;
                    End IF;
                    Raise Err_Custom;
                End IF;
            ElsIF(v_����=1 And Nvl(r_Stock.����,0)=0) Or(v_����=0 And Nvl(r_Stock.����,0)<>0) Then 
                Close c_Stock;
                If ҽ�����_IN IS NULL Then
					If �շ����_IN='4' Then
						v_Error:='�� '||���_IN||' ����������"'||v_����||'"�ķ������������¼�����,����������ݵ���ȷ�ԣ�';
					Else
	                    v_Error:='�� '||���_IN||' ��ҩƷ"'||v_����||'"�ķ������������¼�����,����ҩƷ���ݵ���ȷ�ԣ�';
					End IF;
                Else
					If �շ����_IN='4' Then
						v_Error:='�ڴ�����"'||����_IN||'"ʱ������������"'||v_����||'"�ķ������������¼�����,����������ݵ���ȷ�ԣ�';
					Else
	                    v_Error:='�ڴ�����"'||����_IN||'"ʱ����ҩƷ"'||v_����||'"�ķ������������¼�����,����ҩƷ���ݵ���ȷ�ԣ�';
					End IF;
                End IF;
                Raise Err_Custom;
            End IF;

            --ȷ�����ηֽ�����
            IF v_����=1 Or v_ʱ��=1 Then
                --���ڲ�������ʱ��ֻ���ֽܷ�һ��,�ֽⲻ�������ж���.���ֽ���Ϊ�˼��㵥��.
                --ÿ�ηֽ�ȡС��,��治���ֽⲻ���������ж�.
                IF v_������<=Nvl(r_Stock.��������,0) Then
                    v_��ǰ����:=v_������;
                Else
                    v_��ǰ����:=Nvl(r_Stock.��������,0);
                End if;
                IF v_ʱ��=1 Then 
                    If r_Stock.ʵ������=0 Then
                        v_��ǰ����:=0;
                    Else
                        v_��ǰ����:=Round(Nvl(r_Stock.ʵ�ʽ��/r_Stock.ʵ������,0),5);
                    End IF;
                ElsIf v_����=1 Then
                    v_��ǰ����:=��׼����_IN;
                End IF;
            Else
                --��ͨҩƷ
                --���ܹ�����,�������Ѹ��ݲ����ж�
                v_��ǰ����:=v_������;
                v_��ǰ����:=��׼����_IN;
            End IF;

            --ҩƷ���(��ͨ�������û�м�¼)
            IF c_Stock%Found Then
                Update ҩƷ���
                    Set ��������=Nvl(��������,0)-v_��ǰ����
                Where �ⷿID=ִ�в���ID_IN And ҩƷID=�շ�ϸĿID_IN
                    And Nvl(����,0)=Nvl(r_Stock.����,0) And ����=1;
            ElsIf ִ�в���ID_IN IS Not NULL Then
                --ֻ�в�������ʱ��ҩƷ���ܿ�治�����
                Update ҩƷ���
                    Set ��������=Nvl(��������,0)-v_��ǰ����
                Where �ⷿID=ִ�в���ID_IN And ҩƷID=�շ�ϸĿID_IN
                    And Nvl(����,0)=0 And ����=1;
                IF SQL%RowCount=0 Then
                    Insert Into ҩƷ���(
                        �ⷿID,ҩƷID,����,��������)
                    Values(
                        ִ�в���ID_IN,�շ�ϸĿID_IN,1,-1*v_��ǰ����);
                End IF;
            End IF;

            --ҩƷ�շ���¼
			v_����:=Null;v_����:=Null;
			v_Ч��:=Null;v_����:=Null;
			v_���Ч��:=Null;v_�������:=Null;
            IF c_Stock%Found Then
                v_����:=r_Stock.����;
                v_����:=r_Stock.�ϴ�����;
                v_Ч��:=r_Stock.Ч��;
                v_����:=r_Stock.�ϴβ���;

				--�������Ч��:һ���Բ�������Ч��
				IF �շ����_IN='4' Then
					v_Count:=0;
					Begin
						Select ���Ч�� Into v_Count From �������� Where Nvl(һ���Բ���,0)=1 And ����ID=�շ�ϸĿID_IN;
					Exception
						When Others Then Null;
					End;
					IF Nvl(v_Count,0)>0 Then
						v_���Ч��:=r_Stock.���Ч��;	
						v_�������:=v_���Ч��-v_Count*30;
					End IF;
				End IF;
            End IF;

            Select Nvl(Max(���),0)+1 Into v_��� From ҩƷ�շ���¼ 
				Where NO=NO_IN And ��¼״̬=1 And ����=Decode(�ಡ�˵�_IN,1,10,9)+Decode(�շ����_IN,'4',16,0);
					

            --�޸ĵ�ԭ���ݺŴ����ժҪ��
			v_����:=Null;
            If ��Ч_IN IS Not NULL Or �Ƽ�����_IN IS Not NULL THEN 
                v_����:=Nvl(��Ч_IN,0)||Nvl(�Ƽ�����_IN,0);
            End IF;
            Insert Into ҩƷ�շ���¼(
                ID,��¼״̬,����,NO,���,�ⷿID,�Է�����ID,������ID,���ϵ��,
                ҩƷID,����,����,����,Ч��,����,��д����,ʵ������,���ۼ�,���۽��,
                ժҪ,������,��������,����ID,Ƶ��,����,�÷�,���,����,���Ч��,�������)
            Values(
                ҩƷ�շ���¼_ID.Nextval,1,Decode(�ಡ�˵�_IN,1,10,9)+Decode(�շ����_IN,'4',16,0),
				NO_IN,v_���,ִ�в���ID_IN,��������ID_IN,���ID_IN,-1,�շ�ϸĿID_IN,v_����,v_����,v_����,v_Ч��,
				Decode(v_����,1,1,����_IN),Decode(v_����,1,v_��ǰ����,v_��ǰ����/����_IN),Decode(v_����,1,v_��ǰ����,v_��ǰ����/����_IN),
                v_��ǰ����,Round(v_��ǰ����*v_��ǰ����,v_Dec),ҩƷժҪ_IN,����Ա����_IN,�Ǽ�ʱ��_IN,v_����ID,
                Ƶ��_IN,����_IN,v_�÷�,v_�巨,v_����,v_���Ч��,v_�������);

            --δ��ҩƷ��¼
            Update δ��ҩƷ��¼
                Set ����ID=����ID_IN,
                    ��ҳID=��ҳID_IN,
                    ����=����_IN
            Where ����=Decode(�ಡ�˵�_IN,1,10,9)+Decode(�շ����_IN,'4',16,0)
				And NO=NO_IN And Nvl(�ⷿID,0)=Nvl(ִ�в���ID_IN,0);

            IF SQL%RowCount=0 Then
                --ȡ������ȼ�
                Begin
                    Select B.���ȼ� Into v_���ȼ� From ������Ϣ A,��� B
                     Where A.���=B.����(+) And A.����ID=����ID_IN;
                Exception
                    When Others Then Null;
                End;

                Insert Into δ��ҩƷ��¼(
                    ����,NO,����ID,��ҳID,����,���ȼ�,�Է�����ID,�ⷿID,��������,���շ�,��ӡ״̬)
                Values(
                    Decode(�ಡ�˵�_IN,1,10,9)+Decode(�շ����_IN,'4',16,0),NO_IN,����ID_IN,
					��ҳID_IN,����_IN,v_���ȼ�,��������ID_IN,ִ�в���ID_IN,�Ǽ�ʱ��_IN,Decode(����_IN,1,0,1),0);
            End IF;

            v_������:=v_������-v_��ǰ����;
            v_�ܽ��:=v_�ܽ��+Round(v_��ǰ����*v_��ǰ����,v_Dec);
        End Loop;
        
        --����ʱ��ҩƷ�Ŀ����������仯��
        IF v_ʱ��=1 Then
            IF Round(v_�ܽ��/(����_IN*����_IN),5)<>��׼����_IN Then 
                Close c_Stock;    
                If ҽ�����_IN IS NULL Then
					If �շ����_IN='4' Then
						v_Error:='�� '||���_IN||' �е�ʱ����������"'||v_����||'"��ǰ���㵥�۲�һ��,�����������������㣡';
					Else
						v_Error:='�� '||���_IN||' �е�ʱ��ҩƷ"'||v_����||'"��ǰ���㵥�۲�һ��,�����������������㣡';
					End IF;
                Else
                    --ҽ����ҩʱ�ǰ����˷ִμ��㲢�ύ���ݿ�,��˲�ͬ����ʹ����ͬʵ��ҩƷû�����⡣
                    --��ͬһ����ͬʱʹ������������ͬʵ��ҩƷ��������⡣
					IF �շ����_IN='4' Then
						v_Error:='�ڴ�����"'||����_IN||'"ʱ����ʱ����������"'||v_����||'"��ǰ����ĵ��۷����仯��'||CHR(13)||CHR(10)||'����ò����Ƿ�ͬʱʹ����������ͬ��"'||v_����||'"��';
					Else
	                    v_Error:='�ڴ�����"'||����_IN||'"ʱ����ʱ��ҩƷ"'||v_����||'"��ǰ����ĵ��۷����仯��'||CHR(13)||CHR(10)||'����ò����Ƿ�ͬʱʹ����������ͬ��"'||v_����||'"��';
					End IF;
                End IF;
                Raise Err_Custom;
            End IF;
        End IF;

        Close c_Stock;
    End IF;
Exception
    When Err_custom Then Raise_Application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When Others Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_סԺ���ʼ�¼_Insert;
/

CREATE OR REPLACE PROCEDURE zl_סԺ���ʼ�¼_Verify (
    NO_IN           ���˷��ü�¼.NO%TYPE,
    ����Ա���_IN   ���˷��ü�¼.����Ա���%TYPE,
    ����Ա����_IN	���˷��ü�¼.����Ա����%TYPE,
	���_IN			Varchar2:=NULL,
	����ID_IN		���˷��ü�¼.����ID%Type:=NULL,
	���ʱ��_IN		���˷��ü�¼.�Ǽ�ʱ��%Type:=NULL
) AS
--���ܣ����һ��סԺ���ʻ��۵�
--������
--		���_IN����ʽ��"1,3,5,7,8",Ϊ�ձ�ʾ�������δ��˵���
--		����ID_IN��ֻ���ָ������,���ڰ�������˼��ʱ�
--		���ʱ��_IN�����ڲ�����Ҫͳһ���ƻ򷵻�ʱ��ĵط�
	--ֻ��ȡָ����ŵ�,δ��˵Ĳ��ݽ��д���
	Cursor c_Bill is
		Select * From ���˷��ü�¼ 
		Where ��¼����=2 And ��¼״̬=0 And NO=NO_IN
			And (Instr(','||���_IN||',',','||Nvl(�۸񸸺�,���)||',')>0 Or ���_IN Is Null)
			And (����ID+0=����ID_IN Or ����ID_IN IS NULL)
		Order BY ���;
	
	--����а����������õ�δ������ʱ�����ݲ��������Ƿ��Զ�����
	Cursor c_Stuff is
		Select NO,����,�ⷿID From δ��ҩƷ��¼
		Where NO=NO_IN And ���� IN(25,26) And �ⷿID IS Not Null
			And Exists(Select ����ֵ From ϵͳ������ Where ������=63 And ����ֵ='1')
			And Exists(
				Select A.��� From ���˷��ü�¼ A,�������� B
				Where A.��¼����=2 And A.��¼״̬=1 And A.NO=NO_IN
					And (Instr(','||���_IN||',',','||Nvl(A.�۸񸸺�,A.���)||',')>0 Or ���_IN Is Null)
					And (A.����ID+0=����ID_IN Or ����ID_IN IS NULL)
					And A.�շ�ϸĿID=B.����ID And B.��������=1
				)
		Order BY �ⷿID;
	
	v_Date	Date;
BEGIN
	If ���ʱ��_IN IS Null Then
		Select Sysdate Into v_Date From Dual;
	Else
		v_Date:=���ʱ��_IN;
	End IF;

	For r_Bill IN c_Bill Loop
		Update ���˷��ü�¼
			Set ��¼״̬=1,
				����Ա���=����Ա���_IN,
				����Ա����=����Ա����_IN,
				�Ǽ�ʱ��=v_Date --�Ѳ�����ҩƷ��¼��ʱ�䲻��
		Where ID=r_Bill.ID;

		--ҩƷ�շ���¼.��������
		Update ҩƷ�շ���¼
			Set ��������=Decode(Sign(Nvl(�������,v_Date)-v_Date),-1,��������,v_Date)  
		Where NO=NO_IN AND ���� IN(9,10,25,26)  AND ����ID=r_Bill.ID;

		--�������
		Update �������
			Set �������=Nvl(�������,0)+Nvl(r_Bill.ʵ�ս��,0)
		Where ����ID=r_Bill.����ID And ����=1;

		IF SQL%RowCount=0 Then
			Insert Into �������(
				����ID,����,�������,Ԥ�����)
			Values(
				r_Bill.����ID,1,r_Bill.ʵ�ս��,0);
		End IF;

		--����δ�����
		Update ����δ�����
			Set ���=Nvl(���,0)+Nvl(r_Bill.ʵ�ս��,0)
		 Where ����ID=r_Bill.����ID
			And Nvl(��ҳID,0)=Nvl(r_Bill.��ҳID,0)
			And Nvl(���˲���ID,0)=Nvl(r_Bill.���˲���ID,0)
			And Nvl(���˿���ID,0)=Nvl(r_Bill.���˿���ID,0)
			And Nvl(��������ID,0)=Nvl(r_Bill.��������ID,0)
			And Nvl(ִ�в���ID,0)=Nvl(r_Bill.ִ�в���ID,0)
			And ������ĿID+0=r_Bill.������ĿID
			And ��Դ;��+0=r_Bill.�����־;

		IF SQL%RowCount=0 Then
			Insert Into ����δ�����(
				����ID,��ҳID,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,��Դ;��,���)
			Values(
				r_Bill.����ID,r_Bill.��ҳID,r_Bill.���˲���ID,r_Bill.���˿���ID,r_Bill.��������ID,r_Bill.ִ�в���ID,r_Bill.������ĿID,r_Bill.�����־,Nvl(r_Bill.ʵ�ս��,0));
		End IF;

		--���˷��û���
		Update ���˷��û���
			Set Ӧ�ս��=Nvl(Ӧ�ս��,0)+Nvl(r_Bill.Ӧ�ս��,0),
				ʵ�ս��=Nvl(ʵ�ս��,0)+Nvl(r_Bill.ʵ�ս��,0)
		 Where ����=Trunc(v_Date)
			And Nvl(���˲���ID,0)=Nvl(r_Bill.���˲���ID,0)
			And Nvl(���˿���ID,0)=Nvl(r_Bill.���˿���ID,0)
			And Nvl(��������ID,0)=Nvl(r_Bill.��������ID,0)
			And Nvl(ִ�в���ID,0)=Nvl(r_Bill.ִ�в���ID,0)
			And ������ĿID+0=r_Bill.������ĿID
			And ��Դ;��=r_Bill.�����־ And ���ʷ���=1;

		IF SQL%RowCount=0 Then
			Insert Into ���˷��û���(
				����,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,
				��Դ;��,���ʷ���,Ӧ�ս��,ʵ�ս��,���ʽ��)
			Values(
				Trunc(v_Date),r_Bill.���˲���ID,r_Bill.���˿���ID,r_Bill.��������ID,
				r_Bill.ִ�в���ID,r_Bill.������ĿID,r_Bill.�����־,1,r_Bill.Ӧ�ս��,r_Bill.ʵ�ս��,0);
		End IF;
	End Loop;

	--�ⷿ�е�ҩƷ��ȫ��������Ϊ���շ�
	Update δ��ҩƷ��¼ Set ���շ�=1, ��������=v_Date
	Where NO=NO_IN And ���� IN(9,10) And Nvl(���շ�,0)=0 
		And Nvl(�ⷿID,0) Not IN(
			Select Distinct Nvl(ִ�в���ID,0) From ���˷��ü�¼ 
				Where ��¼����=2 And NO=NO_IN And �շ���� IN('5','6','7') And ��¼״̬=0);

	Update δ��ҩƷ��¼ Set ���շ�=1, ��������=v_Date
	Where NO=NO_IN And ���� IN(25,26) And Nvl(���շ�,0)=0 
		And Nvl(�ⷿID,0) Not IN(
			Select Distinct Nvl(ִ�в���ID,0) From ���˷��ü�¼ 
				Where ��¼����=2 And NO=NO_IN And �շ����='4' And ��¼״̬=0);

	--���������Զ�����
	For r_Stuff In c_Stuff Loop
		zl_�����շ���¼_��������(r_Stuff.�ⷿID,r_Stuff.����,r_Stuff.NO,����Ա����_IN,����Ա����_IN,����Ա����_IN,1,Sysdate);
	End Loop;
EXCEPTION
    WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_סԺ���ʼ�¼_Verify;
/



Create OR Replace Procedure zl_������Ϣ_Merge(
	A����ID_IN		������Ϣ.����ID%Type,--Ҫ�ϲ��Ĳ�����Ϣ
	B����ID_IN		������Ϣ.����ID%Type --Ҫ�����Ĳ�����Ϣ
--�漰��
--������Ϣ,������ҳ,������ҳ�ӱ�,���˱䶯��¼,���ⲡ��
--���ﲡ����¼,סԺ������¼,��鲡����¼,��λ״����¼
--�����ʻ�,����ģ�����,���ս����¼,�ʻ������Ϣ
--�������,����δ�����,���˷��ü�¼,����Ԥ����¼,���˽��ʼ�¼,δ��ҩƷ��¼
--���˹Һż�¼,���˹���ҩ��,���˹�����¼,������ϼ�¼,������ϼ�¼
--����ҽ����¼,���˲�����¼,����������¼,���������¼
--�󱸱�
--H���˽��ʼ�¼,H����Ԥ����¼,H���˷��ü�¼
--H����ҽ����¼,H������ϼ�¼,H���˹�����¼
--H���˲�����¼,H���������¼,H����������¼
) AS
	--���ϲ��Ĳ���
	Cursor c_InfoA IS
		Select A.*,B.��ҳID,B.��Ժ����,B.��Ժ����
		From ������Ϣ A,������ҳ B
		Where A.����ID=B.����ID(+) And A.����ID=A����ID_IN
		Order by ��ҳID;
	r_InfoA c_InfoA%RowType;

	--Ҫ�����Ĳ��� 
	Cursor c_InfoB IS
		Select A.*,B.��ҳID,B.��Ժ����,B.��Ժ����
		From ������Ϣ A,������ҳ B
		Where A.����ID=B.����ID(+) And A.����ID=B����ID_IN
		Order by ��ҳID;
	r_InfoB c_InfoB%RowType;

	--�ϲ������Ϣ
	Cursor c_Info(v_����ID ������Ϣ.����ID%Type) is
		Select * From ������ҳ
		Where ��ҳID=(Select Max(��ҳID) From ������ҳ Where ����ID=v_����ID)
			And ����ID=v_����ID;
	r_Info c_Info%RowType;

	--�ϲ�����סԺ����
	Cursor c_MergePati IS
		Select A.����,A.�����,A.סԺ��,B.* 
		From ������Ϣ A,������ҳ B
		Where A.����ID=B.����ID And A.����ID IN(A����ID_IN,B����ID_IN)
		Order by B.��Ժ���� Desc,NVL(B.��Ժ����,SYSDATE) Desc;
	
	v_����ID	������Ϣ.����ID%Type;
	v_�ϲ�ID	������Ϣ.����ID%Type;
	v_�����	������Ϣ.�����%Type;
	v_סԺ��	������Ϣ.סԺ��%Type;

	--����δ�����(���ﲿ��)
	Cursor c_Owe(v_����ID ������Ϣ.����ID%Type) IS
		Select 
			���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,��Դ;��,Sum(���) AS ���
		From ����δ����� 
		Where ��ҳID IS NULL And ����ID=v_����ID
		Group BY ���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,��Դ;��;

	--�������
	Cursor c_Spare(v_����ID ������Ϣ.����ID%Type) IS
		Select ����,Ԥ�����,�������
		From ������� Where ����ID=v_����ID;
	
	--��鲡����¼(ZLHIS+)
	Cursor c_Check(v_����ID ������Ϣ.����ID%Type) is
		Select * From ��鲡����¼ Where ����ID=v_����ID Order BY ������;

	--�����ʻ�
	Cursor c_Insure(v_����ID ������Ϣ.����ID%Type) IS
		Select * From �����ʻ� Where ����ID=v_����ID Order BY ����;
	
	--Ҫ�����ı����ʻ�
	Cursor c_KeepInsure(v_����ID ������Ϣ.����ID%Type,v_���� �����ʻ�.����%Type) IS
		Select * From �����ʻ� Where ����ID=v_����ID And ����=v_����;
	r_KeepInsure c_KeepInsure%RowType;
	
	Cursor c_Year(v_����ID ������Ϣ.����ID%Type,v_���� �����ʻ�.����%Type) is
		Select * From �ʻ������Ϣ Where ����ID=v_����ID And ����=v_����;

	v_����		�����ʻ�.����%Type;
	v_ҽ����	�����ʻ�.ҽ����%Type;
	
	v_Count		NUMBER;
	v_Error		VARCHAR2(255);
	Err_Custom	EXCEPTION;

	--�ַ����任����
	Function strSwitch(v_InStr In Varchar2,v_Mask In Number)
		Return Varchar2 Is
		v_OutStr	Varchar2(1000);
	Begin
		For v_Bit In 1 .. Length(v_InStr) Loop
			v_OutStr:=v_OutStr||Chr(Ascii(Substr(v_InStr,v_Bit,1))+v_Mask);
		End Loop;
		Return(v_OutStr);
	End strSwitch;
BEGIN
	--�������Ѽ�飺
	--1.ѡ����ͬһ������
	--2.����סԺ��������Ժ��ȴ��Ժ(������������Ժ)��
	--3.����סԺ���˵�סԺ�ڼ���ڽ�������

	OPEN c_InfoA;
	FETCH c_InfoA Into r_InfoA;
	IF c_InfoA%Rowcount=0 THEN
		Close c_InfoA;
		v_Error:='û�з��ֱ��ϲ��Ĳ�����Ϣ��';
		RAISE Err_Custom;
	END IF;

	OPEN c_InfoB;
	FETCH c_InfoB Into r_InfoB;
	IF c_InfoB%Rowcount=0 THEN
		Close c_InfoB;
		v_Error:='û�з���Ҫ�����Ĳ�����Ϣ��';
		RAISE Err_Custom;
	END IF;

	--����סԺ���ȵǼǵĲ���ID��Ϊʵ����Ҫ�����Ĳ���ID
	Select ����ID Into v_����ID
	From (
		Select A.����ID
		From ������Ϣ A,������ҳ B
		Where A.����ID=B.����ID(+) 
			And A.����ID IN(A����ID_IN,B����ID_IN)
		Order by Nvl(B.��Ժ����,To_Date('3000-01-01','YYYY-MM-DD')),NVL(B.��Ժ����,To_Date('3000-01-01','YYYY-MM-DD')),A.�Ǽ�ʱ��,A.����ID --סԺ��������
		)
	Where Rownum=1;

	--����һ������ʵ�����Ҫɾ���Ĳ���ID
	If v_����ID=A����ID_IN Then
		v_�ϲ�ID:=B����ID_IN;
	Else
		v_�ϲ�ID:=A����ID_IN;
	End IF;
	v_�����:=Nvl(r_InfoB.�����,r_InfoA.�����);
	v_סԺ��:=Nvl(r_InfoB.סԺ��,r_InfoA.סԺ��);
	
	IF r_InfoA.��ҳID IS Not NULL And r_InfoB.��ҳID IS Not NULL THEN
		--�����������ܹ���סԺ����
		Select Count(*) Into v_Count From ������ҳ Where ����ID IN(A����ID_IN,B����ID_IN);
		
		FOR r_Merge IN c_MergePati LOOP
			--��������ҳ����(�漰����ID,��ҳID�ֶεı�)
			If Not (r_Merge.����ID=v_����ID And r_Merge.��ҳID=v_Count) Then
				--�ò�����ҳҪɾ��ʱ,�������ѱ�Ŀ�˵ġ�
				If r_Merge.��Ŀ���� IS Not NULL Then
					Close c_InfoA; Close c_InfoB;
					If r_Merge.סԺ�� IS NULL Then
						v_Error:='����'||r_Merge.����||'(����ID='||r_Merge.����ID||')�����ѱ�Ŀ�Ĳ���,������ϲ��ò��ˡ�';
					Else
						v_Error:='����'||r_Merge.����||'(����ID='||r_Merge.����ID||',סԺ��='||r_Merge.סԺ��||')�����ѱ�Ŀ�Ĳ���,������ϲ��ò��ˡ�';
					End IF;
					Raise Err_Custom;
				End IF;

				Insert Into ������ҳ(
					����ID,��ҳID,��������,ҽ�Ƹ��ʽ,�ѱ�,��Ժ����ID,��Ժ����ID,��Ժ����,��Ժ����,
					��Ժ��ʽ,����Ժת��,סԺĿ��,��Ժ����,�Ƿ����,��ǰ����,��ǰ����ID,����ȼ�ID,
					��Ժ����ID,��Ժ����,��Ժ����,סԺ����,��Ժ��ʽ,�Ƿ�ȷ��,ȷ������,�·�����,Ѫ��,
					���ȴ���,�ɹ�����,�����־,��������,ʬ���־,����ҽʦ,���λ�ʿ,סԺҽʦ,
					��ĿԱ���,��ĿԱ����,��Ŀ����,״̬,���ú�,����,����״��,ְҵ,����,ѧ��,��λ�绰,
					��λ�ʱ�,��λ��ַ,����,��ͥ��ַ,��ͥ�绰,�����ʱ�,��ϵ������,��ϵ�˹�ϵ,��ϵ�˵�ַ,
					��ϵ�˵绰,��ҽ�������,�Ǽ���,�Ǽ�ʱ��,����,�Ƿ��ϴ�,��ע,����ת��)
				Values(
					v_����ID,v_Count,r_Merge.��������,r_Merge.ҽ�Ƹ��ʽ,r_Merge.�ѱ�,r_Merge.��Ժ����ID,
					r_Merge.��Ժ����ID,r_Merge.��Ժ����,r_Merge.��Ժ����,r_Merge.��Ժ��ʽ,r_Merge.����Ժת��,
					r_Merge.סԺĿ��,r_Merge.��Ժ����,r_Merge.�Ƿ����,r_Merge.��ǰ����,r_Merge.��ǰ����ID,
					r_Merge.����ȼ�ID,r_Merge.��Ժ����ID,r_Merge.��Ժ����,r_Merge.��Ժ����,r_Merge.סԺ����,
					r_Merge.��Ժ��ʽ,r_Merge.�Ƿ�ȷ��,r_Merge.ȷ������,r_Merge.�·�����,r_Merge.Ѫ��,
					r_Merge.���ȴ���,r_Merge.�ɹ�����,r_Merge.�����־,r_Merge.��������,r_Merge.ʬ���־,
					r_Merge.����ҽʦ,r_Merge.���λ�ʿ,r_Merge.סԺҽʦ,r_Merge.��ĿԱ���,r_Merge.��ĿԱ����,
					r_Merge.��Ŀ����,r_Merge.״̬,r_Merge.���ú�,r_Merge.����,r_Merge.����״��,r_Merge.ְҵ,
					r_Merge.����,r_Merge.ѧ��,r_Merge.��λ�绰,r_Merge.��λ�ʱ�,r_Merge.��λ��ַ,r_Merge.����,
					r_Merge.��ͥ��ַ,r_Merge.��ͥ�绰,r_Merge.�����ʱ�,r_Merge.��ϵ������,r_Merge.��ϵ�˹�ϵ,
					r_Merge.��ϵ�˵�ַ,r_Merge.��ϵ�˵绰,r_Merge.��ҽ�������,r_Merge.�Ǽ���,r_Merge.�Ǽ�ʱ��,
					r_Merge.����,r_Merge.�Ƿ��ϴ�,r_Merge.��ע,r_Merge.����ת��);
				
				--���²�����ر�Ĳ���ָ��
				---------------------------------------------------------------
				--���˱䶯��¼
				Update ���˱䶯��¼
					Set ����ID=v_����ID,��ҳID=v_Count
				Where ����ID=r_Merge.����ID And ��ҳID=r_Merge.��ҳID;

				--������ҳ�ӱ�
				Update ������ҳ�ӱ�
					Set ����ID=v_����ID,��ҳID=v_Count
				Where ����ID=r_Merge.����ID And ��ҳID=r_Merge.��ҳID;
				
				--���˷��ü�¼
				Update ���˷��ü�¼
					Set ����ID=v_����ID,��ҳID=v_Count,
						��ʶ��=Nvl(Decode(�����־,1,v_�����,v_סԺ��),��ʶ��)
				Where ����ID=r_Merge.����ID And ��ҳID=r_Merge.��ҳID;
				Update H���˷��ü�¼
					Set ����ID=v_����ID,��ҳID=v_Count,
						��ʶ��=Nvl(Decode(�����־,1,v_�����,v_סԺ��),��ʶ��)
				Where ����ID=r_Merge.����ID And ��ҳID=r_Merge.��ҳID;

				--����Ԥ����¼
				Update ����Ԥ����¼
					Set ����ID=v_����ID,��ҳID=v_Count
				Where ����ID=r_Merge.����ID And ��ҳID=r_Merge.��ҳID;
				Update H����Ԥ����¼
					Set ����ID=v_����ID,��ҳID=v_Count
				Where ����ID=r_Merge.����ID And ��ҳID=r_Merge.��ҳID;

				--����δ�����
				Update ����δ�����
					Set ����ID=v_����ID,��ҳID=v_Count
				Where ����ID=r_Merge.����ID And ��ҳID=r_Merge.��ҳID;
				
				--δ��ҩƷ��¼
				Update δ��ҩƷ��¼
					Set ����ID=v_����ID,��ҳID=v_Count
				Where ����ID=r_Merge.����ID And ��ҳID=r_Merge.��ҳID;
				
				--������ϼ�¼
				Update ������ϼ�¼
					Set ����ID=v_����ID,��ҳID=v_Count
				Where ����ID=r_Merge.����ID And ��ҳID=r_Merge.��ҳID;
				
				--���ս����¼(�Ƚ�����,��ҳID�����,���������Ӧ�����ʻ�,���������ҳID,����ID�����)
				Update ���ս����¼
					Set ��ҳID=v_Count
				Where ����ID=r_Merge.����ID And ��ҳID=r_Merge.��ҳID;

				--����ģ�����
				Update ����ģ�����
					Set ����ID=v_����ID,��ҳID=v_Count
				Where ����ID=r_Merge.����ID And ��ҳID=r_Merge.��ҳID;

				--����ҽ����¼(ZLHIS+)
				Update ����ҽ����¼
					Set ����ID=v_����ID,��ҳID=v_Count
				Where ����ID=r_Merge.����ID And ��ҳID=r_Merge.��ҳID;
				Update H����ҽ����¼
					Set ����ID=v_����ID,��ҳID=v_Count
				Where ����ID=r_Merge.����ID And ��ҳID=r_Merge.��ҳID;
				
				--���˹�����¼(ZLHIS+)
				Update ���˹�����¼
					Set ����ID=v_����ID,��ҳID=v_Count
				Where ����ID=r_Merge.����ID And ��ҳID=r_Merge.��ҳID;
				Update H���˹�����¼
					Set ����ID=v_����ID,��ҳID=v_Count
				Where ����ID=r_Merge.����ID And ��ҳID=r_Merge.��ҳID;

				--������ϼ�¼(ZLHIS+)
				Update ������ϼ�¼
					Set ����ID=v_����ID,��ҳID=v_Count
				Where ����ID=r_Merge.����ID And ��ҳID=r_Merge.��ҳID;
				Update H������ϼ�¼
					Set ����ID=v_����ID,��ҳID=v_Count
				Where ����ID=r_Merge.����ID And ��ҳID=r_Merge.��ҳID;

				--���˲�����¼(ZLHIS+)
				Update ���˲�����¼
					Set ����ID=v_����ID,��ҳID=v_Count
				Where ����ID=r_Merge.����ID And ��ҳID=r_Merge.��ҳID;
				Update H���˲�����¼
					Set ����ID=v_����ID,��ҳID=v_Count
				Where ����ID=r_Merge.����ID And ��ҳID=r_Merge.��ҳID;

				--���������¼(ZLHIS+)
				Update ���������¼
					Set ����ID=v_����ID,��ҳID=v_Count
				Where ����ID=r_Merge.����ID And ��ҳID=r_Merge.��ҳID;
				Update H���������¼
					Set ����ID=v_����ID,��ҳID=v_Count
				Where ����ID=r_Merge.����ID And ��ҳID=r_Merge.��ҳID;
				
				--����������¼(ZLHIS+)
				Update ����������¼
					Set ����ID=v_����ID,��ҳID=v_Count
				Where ����ID=r_Merge.����ID And ��ҳID=r_Merge.��ҳID;
				Update H����������¼
					Set ����ID=v_����ID,��ҳID=v_Count
				Where ����ID=r_Merge.����ID And ��ҳID=r_Merge.��ҳID;

				--ɾ���ѵ�����Ĳ�����ҳ
				Delete From ������ҳ Where ����ID=r_Merge.����ID And ��ҳID=r_Merge.��ҳID;
			End IF;
			v_Count:=v_Count-1;
		End Loop;
	End IF;

	--���漰��ҳID���ݵĸ���(��סԺ���˻�סԺ����סԺǰ������)
	---------------------------------------------------------------
	--���˷��ü�¼
	Update ���˷��ü�¼
		Set ����ID=v_����ID,
			��ʶ��=Nvl(Decode(�����־,2,v_סԺ��,v_�����),��ʶ��)
	Where ����ID=v_�ϲ�ID And ��ҳID IS NULL;
	Update H���˷��ü�¼
		Set ����ID=v_����ID,
			��ʶ��=Nvl(Decode(�����־,2,v_סԺ��,v_�����),��ʶ��)
	Where ����ID=v_�ϲ�ID And ��ҳID IS NULL;

	--����Ԥ����¼
	Update ����Ԥ����¼
		Set ����ID=v_����ID
	Where ����ID=v_�ϲ�ID And ��ҳID IS NULL;
	Update H����Ԥ����¼
		Set ����ID=v_����ID
	Where ����ID=v_�ϲ�ID And ��ҳID IS NULL;

	--δ��ҩƷ��¼
	Update δ��ҩƷ��¼
		Set ����ID=v_����ID
	Where ����ID=v_�ϲ�ID And ��ҳID IS NULL;

	--������ϼ�¼
	Update ������ϼ�¼
		Set ����ID=v_����ID
	Where ����ID=v_�ϲ�ID And ��ҳID IS NULL;

	--����ҽ����¼(ZLHIS+)
	Update ����ҽ����¼
		Set ����ID=v_����ID
	Where ����ID=v_�ϲ�ID And ��ҳID IS NULL;
	Update H����ҽ����¼
		Set ����ID=v_����ID
	Where ����ID=v_�ϲ�ID And ��ҳID IS NULL;

	--���˹�����¼(ZLHIS+)
	Update ���˹�����¼
		Set ����ID=v_����ID
	Where ����ID=v_�ϲ�ID And ��ҳID IS NULL;
	Update H���˹�����¼
		Set ����ID=v_����ID
	Where ����ID=v_�ϲ�ID And ��ҳID IS NULL;

	--������ϼ�¼(ZLHIS+)
	Update ������ϼ�¼
		Set ����ID=v_����ID
	Where ����ID=v_�ϲ�ID And ��ҳID IS NULL;
	Update H������ϼ�¼
		Set ����ID=v_����ID
	Where ����ID=v_�ϲ�ID And ��ҳID IS NULL;

	--���˲�����¼(ZLHIS+)
	Update ���˲�����¼
		Set ����ID=v_����ID
	Where ����ID=v_�ϲ�ID And ��ҳID IS NULL;
	Update H���˲�����¼
		Set ����ID=v_����ID
	Where ����ID=v_�ϲ�ID And ��ҳID IS NULL;

	--���������¼(ZLHIS+)
	Update ���������¼
		Set ����ID=v_����ID
	Where ����ID=v_�ϲ�ID And ��ҳID IS NULL;
	Update H���������¼
		Set ����ID=v_����ID
	Where ����ID=v_�ϲ�ID And ��ҳID IS NULL;

	--����������¼(ZLHIS+)
	Update ����������¼
		Set ����ID=v_����ID
	Where ����ID=v_�ϲ�ID And ��ҳID IS NULL;
	Update H����������¼
		Set ����ID=v_����ID
	Where ����ID=v_�ϲ�ID And ��ҳID IS NULL;

	--���˹Һż�¼(ZLHIS+)
	Update ���˹Һż�¼ 
		Set ����ID=v_����ID,
			�����=Nvl(v_�����,�����)
	Where ����ID=v_�ϲ�ID;

	--���˽��ʼ�¼
	Update ���˽��ʼ�¼ Set ����ID=v_����ID Where ����ID=v_�ϲ�ID;
	Update H���˽��ʼ�¼ Set ����ID=v_����ID Where ����ID=v_�ϲ�ID;
	
	--��λ״����¼
	Update ��λ״����¼ Set ����ID=v_����ID Where ����ID=v_�ϲ�ID;

	--���ⲡ��
	Select Count(*) Into v_Count From ���ⲡ�� Where ����ID=v_����ID;
	If v_Count=0 Then
		Update ���ⲡ�� Set ����ID=v_����ID Where ����ID=v_�ϲ�ID;	
	Else
		Delete From ���ⲡ�� Where ����ID=v_�ϲ�ID;			
	End IF;
	
	--����δ�����
	For r_Owe In c_Owe(v_�ϲ�ID) Loop
		Update ����δ�����
			Set ���=Nvl(���,0)+Nvl(r_Owe.���,0)
		Where ��ҳID IS NULL And ����ID=v_����ID
			And Nvl(���˲���ID,0)=Nvl(r_Owe.���˲���ID,0)
			And Nvl(���˿���ID,0)=Nvl(r_Owe.���˿���ID,0)
			And Nvl(��������ID,0)=Nvl(r_Owe.��������ID,0)
			And Nvl(ִ�в���ID,0)=Nvl(r_Owe.ִ�в���ID,0)
			And Nvl(������ĿID,0)=Nvl(r_Owe.������ĿID,0)
			And Nvl(��Դ;��,0)=Nvl(r_Owe.��Դ;��,0);
		If SQl%RowCount=0 Then
			Insert Into ����δ�����(
				����ID,��ҳID,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,��Դ;��,���)
			Values(
				v_����ID,NULL,r_Owe.���˲���ID,r_Owe.���˿���ID,r_Owe.��������ID,r_Owe.ִ�в���ID,r_Owe.������ĿID,r_Owe.��Դ;��,r_Owe.���);
		End IF;
	End Loop;
	Delete From ����δ����� Where ����ID=v_�ϲ�ID;
	Delete From ����δ����� Where ����ID=v_����ID And Nvl(���,0)=0;

	--�������
	For r_Spare In c_Spare(v_�ϲ�ID) Loop
		Update �������
			Set Ԥ�����=Nvl(Ԥ�����,0)+Nvl(r_Spare.Ԥ�����,0),
				�������=Nvl(�������,0)+Nvl(r_Spare.�������,0)
		Where Nvl(����,0)=Nvl(r_Spare.����,0) And ����ID=v_����ID;
		If SQL%RowCount=0 Then
			Insert Into �������(
				����ID,����,Ԥ�����,�������)
			Values(
				v_����ID,r_Spare.����,r_Spare.Ԥ�����,r_Spare.�������);
		End If;
	End Loop;
	Delete From ������� Where ����ID=v_�ϲ�ID;
	Delete From ������� Where ����ID=v_����ID And Nvl(Ԥ�����,0)=0 And Nvl(�������,0)=0 And ����=1;

	--���˹���ҩ��
	Insert Into ���˹���ҩ��(
		����ID,����ҩ��ID,����ҩ��)
	Select 
		v_����ID,����ҩ��ID,����ҩ��
	From ���˹���ҩ��
	Where ����ID=v_�ϲ�ID 
		And ����ҩ��ID Not IN(Select ����ҩ��ID From ���˹���ҩ�� Where ����ID=v_����ID);

	Delete From ���˹���ҩ�� Where ����ID=v_�ϲ�ID;

	--���ﲡ����¼
	Select Count(*) Into v_Count From ���ﲡ����¼ Where ����ID=v_����ID;
	If v_Count=0 Then
		Select Count(*) Into v_Count From ���ﲡ����¼ Where ����ID=v_�ϲ�ID;
		If v_Count>0 Then
			Update ���ﲡ����¼ Set ����ID=v_����ID,������=v_����� Where ����ID=v_�ϲ�ID;
		End IF;
	Else
		Delete From ���ﲡ����¼ Where ����ID=v_�ϲ�ID;
		Update ���ﲡ����¼ Set ������=v_����� Where ����ID=v_����ID;
	End IF;

	--סԺ������¼
	Select Count(*) Into v_Count From סԺ������¼ Where ����ID=v_����ID;
	If v_Count=0 Then
		Select Count(*) Into v_Count From סԺ������¼ Where ����ID=v_�ϲ�ID;
		If v_Count>0 Then
			Update סԺ������¼ Set ����ID=v_����ID,������=v_סԺ�� Where ����ID=v_�ϲ�ID;
		End IF;
	Else
		Delete From סԺ������¼ Where ����ID=v_�ϲ�ID;
		Update סԺ������¼ Set ������=v_סԺ�� Where ����ID=v_����ID;
	End IF;
	
	--��鲡����¼(ZLHIS+)
	For r_Check In c_Check(v_�ϲ�ID) Loop
		Select Count(*) Into v_Count From ��鲡����¼ Where ����ID=v_����ID And ������=r_Check.������;
		If v_Count=0 Then
			Update ��鲡����¼
				Set ����ID=v_����ID
			Where ����ID=v_�ϲ�ID And ������=r_Check.������;
		Else
			Delete From ��鲡����¼ Where ����ID=v_�ϲ�ID And ������=r_Check.������;
		End IF;
	End Loop;

	--�����ʻ�
	For r_Insure In c_Insure(v_�ϲ�ID) Loop
		Select Count(*) Into v_Count From �����ʻ� Where ����ID=v_����ID And ����=r_Insure.����;
		If v_Count>0 Then
			--�������˾�����ͬ������ʻ�
			--ת���ʻ������Ϣ
			For r_Year In c_Year(v_�ϲ�ID,r_Insure.����) Loop
				Update �ʻ������Ϣ
					Set �ʻ������ۼ�=Nvl(�ʻ������ۼ�,0)+Nvl(r_Year.�ʻ������ۼ�,0),
						�ʻ�֧���ۼ�=Nvl(�ʻ�֧���ۼ�,0)+Nvl(r_Year.�ʻ�֧���ۼ�,0),
						����ͳ���ۼ�=Nvl(����ͳ���ۼ�,0)+Nvl(r_Year.����ͳ���ۼ�,0),
						ͳ�ﱨ���ۼ�=Nvl(ͳ�ﱨ���ۼ�,0)+Nvl(r_Year.ͳ�ﱨ���ۼ�,0),
						סԺ�����ۼ�=Nvl(סԺ�����ۼ�,0)+Nvl(r_Year.סԺ�����ۼ�,0),
						���ͳ���ۼ�=Nvl(���ͳ���ۼ�,0)+Nvl(r_Year.���ͳ���ۼ�,0),
						�����ۼ�=Nvl(�����ۼ�,0)+Nvl(r_Year.�����ۼ�,0),
						��������=Nvl(��������,r_Year.��������),
						����ͳ���޶�=Nvl(����ͳ���޶�,r_Year.����ͳ���޶�),
						���ͳ���޶�=Nvl(���ͳ���޶�,r_Year.���ͳ���޶�),
						������Ϣ=Nvl(������Ϣ,r_Year.������Ϣ)
				Where ����ID=v_����ID And ����=r_Insure.���� And ���=r_Year.���;
				If SQL%RowCount=0 Then
					Insert Into �ʻ������Ϣ(
						����ID,����,���,�ʻ������ۼ�,�ʻ�֧���ۼ�,����ͳ���ۼ�,ͳ�ﱨ���ۼ�,
						סԺ�����ۼ�,��������,����ͳ���޶�,���ͳ���޶�,�����ۼ�,���ͳ���ۼ�,������Ϣ)
					Values(
						v_����ID,r_Insure.����,r_Year.���,r_Year.�ʻ������ۼ�,r_Year.�ʻ�֧���ۼ�,r_Year.����ͳ���ۼ�,r_Year.ͳ�ﱨ���ۼ�,
						r_Year.סԺ�����ۼ�,r_Year.��������,r_Year.����ͳ���޶�,r_Year.���ͳ���޶�,r_Year.�����ۼ�,r_Year.���ͳ���ۼ�,r_Year.������Ϣ);
				End IF;
			End Loop;

			--ת�Ʊ��ս����¼
			Update ���ս����¼ 
				Set ����ID=v_����ID
			Where ����ID=v_�ϲ�ID And ����=r_Insure.����;
			
			--��ȡ�û�ָ��Ҫ�������˵��ʻ���Ϣ
			If v_�ϲ�ID=B����ID_IN Then
				Open c_KeepInsure(B����ID_IN,r_Insure.����);
				Fetch c_KeepInsure Into r_KeepInsure;
			End IF;

			Delete From �����ʻ� Where ����ID=v_�ϲ�ID And ����=r_Insure.����;	
			
			--�����û�ָ��Ҫ�������˵��ʻ���Ϣ
			If v_�ϲ�ID=B����ID_IN Then
				If c_KeepInsure%RowCount>0 Then
					Update �����ʻ�
						Set ����=r_KeepInsure.����,
							����=r_KeepInsure.����,
							ҽ����=r_KeepInsure.ҽ����,
							����=r_KeepInsure.����,
							��Ա���=r_KeepInsure.��Ա���,
							��λ����=r_KeepInsure.��λ����,
							˳���=r_KeepInsure.˳���,
							����֤��=r_KeepInsure.����֤��,
							�ʻ����=r_KeepInsure.�ʻ����,
							��ǰ״̬=r_KeepInsure.��ǰ״̬,
							����ID=r_KeepInsure.����ID,
							��ְ=r_KeepInsure.��ְ,
							�����=r_KeepInsure.�����,
							�Ҷȼ�=r_KeepInsure.�Ҷȼ�,
							����ʱ��=r_KeepInsure.����ʱ��
					Where ����=r_Insure.���� And ����ID=v_����ID;
				End IF;
				Close c_KeepInsure;
			End IF;
		Else
			--�������˾��в�ͬ������ʻ�(��������û��)

			--�������˷ֱ����ڲ�ͬ����ʱ������ϲ�
			Select Count(*) Into v_Count From �����ʻ� Where ����ID=v_����ID;
			IF v_Count>0 Then
				Close c_InfoA; Close c_InfoB;
				v_Error:='�������˷ֱ����ڲ�ͬ�ı�����𣬲�����ϲ���';
				Raise Err_Custom;
			End IF;
			
			--Ϊ��������ظ�,����ҽ���źͿ���
			v_����:=strSwitch(r_Insure.����,-31);
			v_ҽ����:=strSwitch(r_Insure.ҽ����,-31);
			Insert Into �����ʻ�(
				����ID,����,����,����,ҽ����,����,��Ա���,
				��λ����,˳���,����֤��,�ʻ����,��ǰ״̬,
				����ID,��ְ,�����,�Ҷȼ�,����ʱ��)
			Values(
				v_����ID,r_Insure.����,r_Insure.����,v_����,v_ҽ����,
				r_Insure.����,r_Insure.��Ա���,r_Insure.��λ����,r_Insure.˳���,r_Insure.����֤��,
				r_Insure.�ʻ����,r_Insure.��ǰ״̬,r_Insure.����ID,r_Insure.��ְ,r_Insure.�����,
				r_Insure.�Ҷȼ�,r_Insure.����ʱ��);

			--ת���ʻ������Ϣ
			Update �ʻ������Ϣ 
				Set ����ID=v_����ID
			Where ����ID=v_�ϲ�ID And ����=r_Insure.����;

			--ת�Ʊ��ս����¼
			Update ���ս����¼ 
				Set ����ID=v_����ID
			Where ����ID=v_�ϲ�ID And ����=r_Insure.����;

			Delete From �����ʻ� Where ����ID=v_�ϲ�ID And ����=r_Insure.����;
			
			--��ԭҽ���źͿ���
			v_����:=strSwitch(v_����,31);
			v_ҽ����:=strSwitch(v_ҽ����,31);
			Update �����ʻ�
				Set ����=v_����,ҽ����=v_ҽ����
			Where ����ID=v_����ID And ����=r_Insure.����;
		End IF;
	End Loop;

	--ɾ��ʵ�ʲ������Ĳ�����Ϣ
	Delete From ������Ϣ Where ����ID=v_�ϲ�ID;

	--���ݽ���ѡ����������Ϣ
	Update ������Ϣ
		Set ����=Nvl(r_InfoB.����,r_InfoA.����),
			�Ա�=Nvl(r_InfoB.�Ա�,r_InfoA.�Ա�),
			����=Nvl(r_InfoB.����,r_InfoA.����),
			�����=Nvl(r_InfoB.�����,r_InfoA.�����),
			סԺ��=Nvl(r_InfoB.סԺ��,r_InfoA.סԺ��),
			���￨��=Nvl(r_InfoB.���￨��,r_InfoA.���￨��),
			����֤��=Decode(r_InfoB.���￨��,NULL,r_InfoA.����֤��,r_InfoB.����֤��),
			�ѱ�=Nvl(r_InfoB.�ѱ�,r_InfoA.�ѱ�),
			ҽ�Ƹ��ʽ=Nvl(r_InfoB.ҽ�Ƹ��ʽ,r_InfoA.ҽ�Ƹ��ʽ),
			��������=Nvl(r_InfoB.��������,r_InfoA.��������),
			�����ص�=Nvl(r_InfoB.�����ص�,r_InfoA.�����ص�),
			���֤��=Nvl(r_InfoB.���֤��,r_InfoA.���֤��),
			���=Nvl(r_InfoB.���,r_InfoA.���),
			ְҵ=Nvl(r_InfoB.ְҵ,r_InfoA.ְҵ),
			����=Nvl(r_InfoB.����,r_InfoA.����),
			����=Nvl(r_InfoB.����,r_InfoA.����),
			ѧ��=Nvl(r_InfoB.ѧ��,r_InfoA.ѧ��),
			����״��=Nvl(r_InfoB.����״��,r_InfoA.����״��),
			��ͥ��ַ=Nvl(r_InfoB.��ͥ��ַ,r_InfoA.��ͥ��ַ),
			��ͥ�绰=Nvl(r_InfoB.��ͥ�绰,r_InfoA.��ͥ�绰),
			�����ʱ�=Nvl(r_InfoB.�����ʱ�,r_InfoA.�����ʱ�),
			��ϵ������=Nvl(r_InfoB.��ϵ������,r_InfoA.��ϵ������),
			��ϵ�˹�ϵ=Nvl(r_InfoB.��ϵ�˹�ϵ,r_InfoA.��ϵ�˹�ϵ),
			��ϵ�˵�ַ=Nvl(r_InfoB.��ϵ�˵�ַ,r_InfoA.��ϵ�˵�ַ),
			��ϵ�˵绰=Nvl(r_InfoB.��ϵ�˵绰,r_InfoA.��ϵ�˵绰),
			��ͬ��λid=Nvl(r_InfoB.��ͬ��λid,r_InfoA.��ͬ��λid),
			������λ=Nvl(r_InfoB.������λ,r_InfoA.������λ),
			��λ�绰=Nvl(r_InfoB.��λ�绰,r_InfoA.��λ�绰),
			��λ�ʱ�=Nvl(r_InfoB.��λ�ʱ�,r_InfoA.��λ�ʱ�),
			��λ������=Nvl(r_InfoB.��λ������,r_InfoA.��λ������),
			��λ�ʺ�=Nvl(r_InfoB.��λ�ʺ�,r_InfoA.��λ�ʺ�),
			������=Nvl(r_InfoB.������,r_InfoA.������),
			������=Decode(r_InfoB.������,NULL,r_InfoA.������,r_InfoB.������),
			��������=Decode(r_InfoB.������,NULL,r_InfoA.��������,r_InfoB.��������),
			����ʱ��=Nvl(r_InfoB.����ʱ��,r_InfoA.����ʱ��),
			����״̬=Nvl(r_InfoB.����״̬,r_InfoA.����״̬),
			��������=Nvl(r_InfoB.��������,r_InfoA.��������),
			����=Nvl(r_InfoB.����,r_InfoA.����),
			�Ǽ�ʱ��=Nvl(r_InfoB.�Ǽ�ʱ��,r_InfoA.�Ǽ�ʱ��),
			סԺ����=NULL,��ǰ����=NULL,
			��ǰ����ID=NULL,��ǰ����ID=NULL,
			��Ժʱ��=NULL,��Ժʱ��=NULL
	Where ����ID=v_����ID;

	OPEN c_Info(v_����ID);
	FETCH c_Info Into r_Info;
	IF c_Info%Rowcount>0 THEN
		Update ������Ϣ
			Set סԺ����=r_Info.��ҳID,
				��ǰ����=Decode(r_Info.��Ժ����,NULL,r_Info.��Ժ����,NULL),
				��ǰ����ID=Decode(r_Info.��Ժ����,NULL,r_Info.��ǰ����ID,NULL),
				��ǰ����ID=Decode(r_Info.��Ժ����,NULL,r_Info.��Ժ����ID,NULL),
				��Ժʱ��=r_Info.��Ժ����,��Ժʱ��=r_Info.��Ժ����
		Where ����ID=v_����ID;
	End IF;
	Close c_Info;

	Close c_InfoA;
	Close c_InfoB;
EXCEPTION
	WHEN Err_Custom THEN Raise_application_error(-20101,'[ZLSOFT]' || v_Error || '[ZLSOFT]');
	WHEN OTHERS THEN zl_ErrorCenter(SQLCODE,SQLERRM);
END zl_������Ϣ_Merge;
/


CREATE OR REPLACE Procedure ZL_���˱䶯��¼_תסԺ(
--���ܣ���סԺ���۲���תΪסԺ����
    ����ID_IN    ������ҳ.����ID%Type,
    ��ҳID_IN    ������ҳ.��ҳID%Type,
    סԺ��_IN    ������Ϣ.סԺ��%Type
) IS
    Cursor c_Info IS
        Select * From ���˱䶯��¼
        Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��ֹʱ�� is Null;
    r_Info    c_Info%RowType;

    v_Count        Number;
    v_Date        Date;
    v_Temp        Varchar2(255);
    v_��Ա���    ���˷��ü�¼.����Ա���%Type;
    v_��Ա����    ���˷��ü�¼.����Ա����%Type;

    v_Error        Varchar2(255);
    Err_Custom    Exception;
Begin
    --�����������
    Select Nvl(״̬,0) Into v_Count From ������ҳ Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��������=2;
    if v_Count=1 Then
        v_Error:='���˵�ǰ��δ���,����תΪסԺ���ˡ����Ƚ�������ƺ����ԡ�';
        Raise Err_Custom;
    ElsIf v_Count=2 Then
        v_Error:='���˵�ǰ����ת��,����תΪסԺ���ˡ����Ƚ�����ת�ƻ�ȡ��ת�ƺ����ԡ�';
        Raise Err_Custom;
    End IF;

    Select Sysdate Into v_Date From Dual;
    v_Temp:=zl_Identity;
    v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
    v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
    v_��Ա���:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
    v_��Ա����:=Substr(v_Temp,Instr(v_Temp,',')+1);

    Open c_Info;--�����ȴ�

    --ȡ���ϴα䶯
    Update ���˱䶯��¼
        Set ��ֹʱ��=v_Date,��ֹԭ��=9,��ֹ��Ա=v_��Ա����
    Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��ֹʱ�� is Null;

    --�����±䶯
    Fetch c_Info Into r_Info;
    if c_Info%RowCount=0 Then
        Close c_Info;
        v_Error:='δ���ָò��˵�ǰ��Ч�ı䶯��¼��';
        Raise Err_Custom;
    End IF;

    --�����䶯��¼
    While c_Info%Found Loop
        Insert Into ���˱䶯��¼(
            ����ID,��ҳID,��ʼʱ��,��ʼԭ��,���Ӵ�λ,����ID,����ID,
            ����ȼ�ID,��λ�ȼ�ID,����,���λ�ʿ,����ҽʦ,����ҽʦ,����ҽʦ,����,����Ա���,����Ա����)
        Values(
            ����ID_IN,��ҳID_IN,v_Date,9,r_Info.���Ӵ�λ,r_Info.����ID,
            r_Info.����ID,r_Info.����ȼ�ID,r_Info.��λ�ȼ�ID,r_Info.����,
            r_Info.���λ�ʿ,r_Info.����ҽʦ,r_Info.����ҽʦ,r_Info.����ҽʦ,r_Info.����,v_��Ա���,v_��Ա����);
        Fetch c_Info Into r_Info;
    End Loop;

    Close c_Info;

    Update ������ҳ Set ��������=0 Where ����ID=����ID_IN And ��ҳID=��ҳID_IN;    
    Update ������Ϣ Set סԺ��=סԺ��_IN Where ����ID=����ID_IN;

    --�����������
    Select Count(*) Into v_Count 
    From ���˱䶯��¼ 
    Where ����ID=����ID_IN And ��ҳID=��ҳID_IN
        AND NVL(���Ӵ�λ,0)=0 And ��ʼʱ�� IS NOT Null
        AND ��ֹʱ�� is Null;

    if v_Count > 1 Then
        v_Error:='���ֲ��˴��ڷǷ��ı䶯��¼,��ǰ�������ܼ�����'||Chr(13)||Chr(10)||'��������������粢�����������,��ˢ�²���״̬�����ԣ�';
        Raise Err_Custom;
    End IF;
Exception
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_���˱䶯��¼_תסԺ;
/

Create Or Replace Procedure zl_���˱䶯��¼_Nurse(
    ����ID_IN        ������ҳ.����ID%Type,
    ��ҳID_IN        ������ҳ.��ҳID%Type,
    ����ID_IN        ���˱䶯��¼.����ȼ�ID%Type,
    ��Чʱ��_IN        ���˱䶯��¼.��ʼʱ��%Type,
    ����Ա���_IN    ���˱䶯��¼.����Ա���%Type,
    ����Ա����_IN    ���˱䶯��¼.����Ա����%Type
)
AS
-----------------------------------------------------------
--˵�������Ĳ��˻���ȼ�
-----------------------------------------------------------
    Cursor c_OldInfo IS
        Select * From ���˱䶯��¼
        Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��ֹʱ�� is Null;

    r_OldInfo    c_OldInfo%RowType;
    v_Count        Number;
    v_Error        Varchar2(255);
    Err_Custom    Exception;
Begin
    Open c_OldInfo;--�����ȴ�

    --ȡ���ϴα䶯
    Update ���˱䶯��¼
        Set ��ֹʱ��=��Чʱ��_IN,��ֹԭ��=6,��ֹ��Ա=����Ա����_IN
    Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��ֹʱ�� is Null;

    --�����±䶯
    Fetch c_OldInfo Into r_OldInfo;

    if c_OldInfo%RowCount=0 Then
        Close c_OldInfo;
        v_Error:='δ���ָò��˵�ǰ��Ч�ı䶯��¼��';
        Raise Err_Custom;
    End IF;

    While c_OldInfo%Found Loop
        Insert Into ���˱䶯��¼(
            ����ID,��ҳID,��ʼʱ��,��ʼԭ��,���Ӵ�λ,����ID,����ID,
            ����ȼ�ID,��λ�ȼ�ID,����,���λ�ʿ,����ҽʦ,����ҽʦ,����ҽʦ,����,����Ա���,����Ա����)
        Values(
            ����ID_IN,��ҳID_IN,��Чʱ��_IN,6,r_OldInfo.���Ӵ�λ,r_OldInfo.����ID,
            r_OldInfo.����ID,����ID_IN,r_OldInfo.��λ�ȼ�ID,r_OldInfo.����,
            r_OldInfo.���λ�ʿ,r_OldInfo.����ҽʦ,r_OldInfo.����ҽʦ,r_OldInfo.����ҽʦ,r_OldInfo.����,����Ա���_IN,����Ա����_IN);
        Fetch c_OldInfo Into r_OldInfo;
    End Loop;

    Close c_OldInfo;

    Update ������ҳ Set ����ȼ�ID=����ID_IN
    Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��Ժ���� is Null;

    --�����������
    Select Count(*) Into v_Count
    From ���˱䶯��¼
    Where ����ID=����ID_IN And ��ҳID=��ҳID_IN
        AND NVL(���Ӵ�λ,0)=0 And ��ʼʱ�� IS NOT Null
        AND ��ֹʱ�� is Null;

    if v_Count > 1 Then
        v_Error:='���ֲ��˴��ڷǷ��ı䶯��¼,��ǰ�������ܼ�����'||Chr(13)||Chr(10)||'��������������粢�����������,��ˢ�²���״̬�����ԣ�';
        Raise Err_Custom;
    End IF;
Exception
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_���˱䶯��¼_Nurse;
/

Create Or Replace Procedure zl_���˱䶯��¼_BedLevel(
    ����ID_IN        ������ҳ.����ID%Type,
    ��ҳID_IN        ������ҳ.��ҳID%Type,
    ����_IN            ���˱䶯��¼.����%Type,
    �ȼ�ID_IN        ���˱䶯��¼.��λ�ȼ�ID%Type,
    ��Чʱ��_IN        ���˱䶯��¼.��ʼʱ��%Type,
    ����Ա���_IN    ���˱䶯��¼.����Ա���%Type,
    ����Ա����_IN    ���˱䶯��¼.����Ա����%Type
)
AS
-----------------------------------------------------------
--˵�������Ĵ�λ�ȼ�
-----------------------------------------------------------
    Cursor c_OldInfo IS
        Select * From ���˱䶯��¼
        Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��ֹʱ�� is Null;

    r_OldInfo    c_OldInfo%RowType;
    v_Count        Number;
    v_Error        Varchar2(255);
    Err_Custom    Exception;
Begin
    Open c_OldInfo;--�����ȴ�

    --ȡ���ϴα䶯
    Update ���˱䶯��¼ 
        Set ��ֹʱ��=��Чʱ��_IN,��ֹԭ��=5,��ֹ��Ա=����Ա����_IN
    Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��ֹʱ�� is Null;

    --�����±䶯
    Fetch c_OldInfo Into r_OldInfo;

    if c_OldInfo%RowCount=0 Then
        Close c_OldInfo;
        v_Error:='δ���ָò��˵�ǰ��Ч�ı䶯��¼��';
        Raise Err_Custom;
    End IF;

    While c_OldInfo%Found Loop
        if r_OldInfo.����=����_IN Then
            Update ��λ״����¼ Set �ȼ�ID=�ȼ�ID_IN
            Where ����ID=r_OldInfo.����ID And ����=����_IN;

            Insert Into ���˱䶯��¼(
                ����ID,��ҳID,��ʼʱ��,��ʼԭ��,���Ӵ�λ,����ID,����ID,
                ����ȼ�ID,��λ�ȼ�ID,����,���λ�ʿ,����ҽʦ,����ҽʦ,����ҽʦ,����,����Ա���,����Ա����)
            Values(
                ����ID_IN,��ҳID_IN,��Чʱ��_IN,5,r_OldInfo.���Ӵ�λ,r_OldInfo.����ID,
                r_OldInfo.����ID,r_OldInfo.����ȼ�ID,�ȼ�ID_IN,����_IN,r_OldInfo.���λ�ʿ,
                r_OldInfo.����ҽʦ,r_OldInfo.����ҽʦ,r_OldInfo.����ҽʦ,r_OldInfo.����,����Ա���_IN,����Ա����_IN);
        ELSE
            Insert Into ���˱䶯��¼(
                ����ID,��ҳID,��ʼʱ��,��ʼԭ��,���Ӵ�λ,����ID,����ID,����ȼ�ID,
                ��λ�ȼ�ID,����,���λ�ʿ,����ҽʦ,����ҽʦ,����ҽʦ,����,����Ա���,����Ա����)
            Values(
                ����ID_IN,��ҳID_IN,��Чʱ��_IN,5,r_OldInfo.���Ӵ�λ,r_OldInfo.����ID,
                r_OldInfo.����ID,r_OldInfo.����ȼ�ID,r_OldInfo.��λ�ȼ�ID,r_OldInfo.����,
                r_OldInfo.���λ�ʿ,r_OldInfo.����ҽʦ,r_OldInfo.����ҽʦ,r_OldInfo.����ҽʦ,r_OldInfo.����,����Ա���_IN,����Ա����_IN);
        End IF;

        Fetch c_OldInfo Into r_OldInfo;
    End Loop;

    Close c_OldInfo;
    --�����������
    Select Count(*) Into v_Count From ���˱䶯��¼
    Where ����ID=����ID_IN And ��ҳID=��ҳID_IN
        AND NVL(���Ӵ�λ,0)=0 And ��ʼʱ�� IS NOT Null
        AND ��ֹʱ�� is Null;

    if v_Count > 1 Then
        v_Error:='���ֲ��˴��ڷǷ��ı䶯��¼,��ǰ�������ܼ�����'||Chr(13)||Chr(10) ||'��������������粢�����������,��ˢ�²���״̬�����ԣ�';
        Raise Err_Custom;
    End IF;
Exception
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_���˱䶯��¼_BedLevel;
/

Create Or Replace Procedure zl_���˱䶯��¼_Move(
    ����ID_IN        ������ҳ.����ID%Type,
    ��ҳID_IN        ������ҳ.��ҳID%Type,
    ����ʱ��_IN        ���˱䶯��¼.��ʼʱ��%Type,
    ����_IN            Varchar2,
    ����Ա���_IN    ���˱䶯��¼.����Ա���%Type,
    ����Ա����_IN    ���˱䶯��¼.����Ա����%Type,
    ����ID_IN    ���˱䶯��¼.����ID%Type:=Null
)
AS
-----------------------------------------------------------
--˵�������˻���
--������
--       ����=Null:��ͥ����;"����1,����2,....����n"
-----------------------------------------------------------
    Cursor c_BedInfo IS
        Select ����ID,���� From ��λ״����¼ Where ����ID=����ID_IN;

    Cursor c_OldInfo IS
        Select * From ���˱䶯��¼ Where ����ID=����ID_IN
            AND ��ҳID=��ҳID_IN And ��ֹʱ�� is Null And NVL(���Ӵ�λ,0)=0;

    r_OldInfo    c_OldInfo%RowType;
    v_���Ŵ�    Varchar2(255);
    v_����        ���˱䶯��¼.����%Type;
    v_�ȼ�ID    ��λ״����¼.�ȼ�ID%Type;
    v_�������Ҷ��� Number(1);
    v_Tmp        Number;
    v_Count        Number;
    v_Error        Varchar2(255);
    Err_Custom    Exception;
Begin
    --�����������
    Select Count(*) Into v_Count From ������ҳ
    Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And NVL(״̬,0)=0;

    if v_Count=0 Then
        v_Error:='���˵�ǰ����������סԺ״̬,������δ���,�������ܼ�����'||Chr(13) ||Chr(10) ||'��������������粢�����������,��ˢ�²���״̬�����ԣ�';
        Raise Err_Custom;
    End IF;

    v_�������Ҷ���:=To_Number(Nvl(ZL_GetSysParameter(99),0));    --���˻���ʱ���Ի�����

    --�˳�����ԭ��λ
    For r_bedrow IN c_BedInfo Loop
        Update ��λ״����¼ 
            Set ״̬='�մ�',
                ����ID=Null,
                ����ID=Decode(����,1,NULL,����ID)
        Where ����ID=r_bedrow.����ID And ����=r_bedrow.����;
    End Loop;    

    Open c_OldInfo;
    Fetch c_OldInfo Into r_OldInfo;

    --ȡ�������䶯��¼
    Update ���˱䶯��¼
        Set ��ֹʱ��=����ʱ��_IN,��ֹԭ��=4,��ֹ��Ա=����Ա����_IN
    Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��ֹʱ�� is Null;

    --�������˴�λ
    if ����_IN is Null Then
        --��ͥ����
        Insert Into ���˱䶯��¼(
            ����ID,��ҳID,��ʼʱ��,��ʼԭ��,���Ӵ�λ,����ID,����ID,
            ����ȼ�ID,��λ�ȼ�ID,����,���λ�ʿ,����ҽʦ,����ҽʦ,����ҽʦ,����,����Ա���,����Ա����)
        Values(
            ����ID_IN,��ҳID_IN,����ʱ��_IN,4,0,r_OldInfo.����ID,r_OldInfo.����ID,
            r_OldInfo.����ȼ�ID,Null,Null,r_OldInfo.���λ�ʿ,r_OldInfo.����ҽʦ,r_OldInfo.����ҽʦ,r_OldInfo.����ҽʦ,r_OldInfo.����,����Ա���_IN,����Ա����_IN);
    ELSE
        --��סһ�Ż���Ų�������
        v_Count:=0;
        v_���Ŵ�:=����_IN||',';

        While v_���Ŵ� IS NOT Null Loop
            v_����:=to_Number(Substr(v_���Ŵ�,1,Instr(v_���Ŵ�,',')-1));
            
            Select �ȼ�ID Into v_�ȼ�ID From ��λ״����¼ Where ����ID=Decode(v_�������Ҷ���,1,����ID_IN,r_OldInfo.����ID) And ����=v_����;

            Insert Into ���˱䶯��¼(
                ����ID,��ҳID,��ʼʱ��,��ʼԭ��,���Ӵ�λ,����ID,����ID,����ȼ�ID,
                ��λ�ȼ�ID,����,���λ�ʿ,����ҽʦ,����ҽʦ,����ҽʦ,����,����Ա���,����Ա����)
            Values(
                ����ID_IN,��ҳID_IN,����ʱ��_IN,4,Decode(v_Count,0,0,1),
                Decode(v_�������Ҷ���,1,����ID_IN,r_OldInfo.����ID),r_OldInfo.����ID,r_OldInfo.����ȼ�ID,v_�ȼ�ID,
                v_����,r_OldInfo.���λ�ʿ,r_OldInfo.����ҽʦ,r_OldInfo.����ҽʦ,r_OldInfo.����ҽʦ,r_OldInfo.����,����Ա���_IN,����Ա����_IN);

            Select Count(*) Into v_Tmp From ��λ״����¼ Where ����ID=Decode(v_�������Ҷ���,1,����ID_IN,r_OldInfo.����ID) And ����=v_���� And ״̬='�մ�';
            if v_Tmp=0 Then
                v_Error:='����ʧ��,��λ '||v_����||' ���ǿմ���';
                Raise Err_Custom;
            End IF;

            Update ��λ״����¼ 
                Set ״̬='ռ��',
                    ����ID=����ID_IN ,
                    ����ID=Decode(����,1,r_OldInfo.����ID,����ID)
            Where ����ID=Decode(v_�������Ҷ���,1,����ID_IN,r_OldInfo.����ID) And ����=v_����;

            v_���Ŵ�:=Substr(v_���Ŵ�,Instr(v_���Ŵ�,',')+1);
            v_Count:=v_Count+1;
        End Loop;
    End IF;

    Close c_OldInfo;
    --������Ϣ��������ҳ(��¼��һ�Ŵ�λ)
    v_���Ŵ�:=����_IN||',';
    v_����:=to_Number(Substr(v_���Ŵ�,1,Instr(v_���Ŵ�,',')-1));

    IF v_�������Ҷ���=1 THEN 
        Update ������ҳ Set ��Ժ����=v_����,��ǰ����ID=����ID_IN Where ����ID=����ID_IN And ��ҳID=��ҳID_IN;
        Update ������Ϣ Set ��ǰ����=v_����,��ǰ����ID=����ID_IN Where ����ID=����ID_IN;
    ELSE 
        Update ������ҳ Set ��Ժ����=v_���� Where ����ID=����ID_IN And ��ҳID=��ҳID_IN;
        Update ������Ϣ Set ��ǰ����=v_���� Where ����ID=����ID_IN;
    END IF;

    --�����������
    Select Count(*) Into v_Count From ���˱䶯��¼
    Where ����ID=����ID_IN And ��ҳID=��ҳID_IN
        AND NVL(���Ӵ�λ,0)=0 And ��ʼʱ�� IS NOT Null
        AND ��ֹʱ�� is Null;

    if v_Count > 1 Then
        v_Error:='���ֲ��˴��ڷǷ��ı䶯��¼,��ǰ�������ܼ�����'||Chr(13)||Chr(10) ||'��������������粢�����������,��ˢ�²���״̬�����ԣ�';
        Raise Err_Custom;
    End IF;
Exception
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_���˱䶯��¼_Move;
/

Create Or Replace Procedure zl_���˱䶯��¼_InDept(
    ����ID_IN        ������ҳ.����ID%Type,
    ��ҳID_IN        ������ҳ.��ҳID%Type,
    ����_IN            Varchar2,
    ����ID_IN        ������ҳ.��ǰ����ID%Type,
    ����ID_IN        ������ҳ.��Ժ����ID%Type,
    ����ȼ�ID_IN    ������ҳ.����ȼ�ID%Type,
    ��ǰ����_IN        ������ҳ.��ǰ����%Type,
    ���λ�ʿ_IN        ������ҳ.���λ�ʿ%Type,
    ����ҽʦ_IN        ������ҳ.����ҽʦ%Type,
    סԺҽʦ_IN        ������ҳ.סԺҽʦ%Type,
    �Ƿ����_IN        ������ҳ.�Ƿ����%Type,
    ���ʱ��_IN        ���˱䶯��¼.��ʼʱ��%Type,
    ����Ա���_IN    ��Ա��.���%Type,
    ����Ա����_IN    ��Ա��.����%Type,
    ��Ժ_IN            Number,
    ����ҽʦ_IN        ������ҳ.סԺҽʦ%Type:=Null,
    ����ҽʦ_IN        ������ҳ.סԺҽʦ%Type:=Null
)
AS
-----------------------------------------------------------
--˵������ɲ�����Ժ��ת����ƴ���
--������
--       ��Ժ_IN:��������Ժ����ת����ơ�
--       ����_IN:Ϊ�ձ�ʾ��ͥ����,����Ϊ"����1,����2,...����n",�������ʱ,��ʾ������
-----------------------------------------------------------
    Cursor c_BedInfo IS
        Select ����ID,���� From ��λ״����¼ Where ����ID=����ID_IN;

    v_����        Varchar2(255);
    v_��ǰ����		��λ״����¼.����%Type;
    v_�ȼ�ID		��λ״����¼.�ȼ�ID%Type;
    v_����ID		������ҳ.��ǰ����ID%Type;
    v_��ֹ��Ա		���˱䶯��¼.��ֹ��Ա%Type;
    v_Count        Number;
    Err_Custom    Exception;
    v_Error        Varchar2(255);
Begin
    --����ʱ,����ֻȡһ����д��Ժ������
    v_����:=����_IN||',';
    v_����:=Substr(v_����,1,Instr(v_����,',')-1);

    --��Ҫ���²�����Ϣ
    Update ������Ϣ
        Set ��ǰ����ID=����ID_IN,
            ��ǰ����ID=����ID_IN,
            ��ǰ����=to_Number(v_����)
    Where ����ID=����ID_IN;

    if ��Ժ_IN=1 Then
        --��Ժ���
        Select Count(*) Into v_Count From ������ҳ Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ״̬=1;

        if v_Count=0 Then
            v_Error:='���˵�ǰ��������Ժ״̬,�����Ѿ�����Ժ���������ܼ�����'||Chr(13) ||Chr(10) ||'��������������粢����������ģ���ˢ�²���״̬�����ԣ�';
            Raise Err_Custom;
        End IF;

        Select ��Ժ����ID Into v_����ID From ������ҳ Where ����ID=����ID_IN And ��ҳID=��ҳID_IN;

        if v_����ID <> ����ID_IN Then
            v_Error:='��ǰ��ס�����벡�˵Ǽǲ�����һ��,����״̬�Ѿ�����,�������ܼ�����'||Chr(13) ||Chr(10) ||'��������������粢����������ģ���ˢ�²���״̬�����ԣ�';
            Raise Err_Custom;
        End IF;

        --������ҳ
        --ͬʱ��������Ժ�Ǽ�ʱ�Ŀ���,����,����
        Update ������ҳ
            Set ��Ժ����ID=����ID_IN,
                ��Ժ����ID=����ID_IN,
                ��Ժ����=��ǰ����_IN,
                ״̬=0,
                ��Ժ����=to_Number(v_����),
                ��Ժ����=to_Number(v_����),
                ��ǰ����ID=����ID_IN,
                ��Ժ����ID=����ID_IN,
                ����ȼ�ID=Decode(����ȼ�ID_IN,0,Null,����ȼ�ID_IN),
                ��ǰ����=��ǰ����_IN,
                ���λ�ʿ=���λ�ʿ_IN,
                ����ҽʦ=����ҽʦ_IN,
                סԺҽʦ=סԺҽʦ_IN,
                �Ƿ����=�Ƿ����_IN
         Where ����ID=����ID_IN And ��ҳID=��ҳID_IN;
         Insert Into ������ҳ�ӱ� (����ID,��ҳID,��Ϣ��,��Ϣֵ) 
                Values (����ID_IN,��ҳID_IN,'����ҽʦ',����ҽʦ_IN);
         Insert Into ������ҳ�ӱ� (����ID,��ҳID,��Ϣ��,��Ϣֵ) 
                Values (����ID_IN,��ҳID_IN,'����ҽʦ',����ҽʦ_IN);

        --��¼��һ������ֹ������Ա
        Update ���˱䶯��¼ 
            Set ��ֹʱ��=���ʱ��_IN,��ֹԭ��=2,��ֹ��Ա=����Ա����_IN
        Where ����ID=����ID_IN And ��ҳID=��ҳID_IN
            AND ��ʼʱ�� IS Not NULL And ��ֹʱ�� is Null;

        if ����_IN is Null Then
            --����ͥ����
            Insert Into ���˱䶯��¼(
                ����ID,��ҳID,��ʼʱ��,��ʼԭ��,���Ӵ�λ,����ID,����ID,
                ����ȼ�ID,��λ�ȼ�ID,����,���λ�ʿ,����ҽʦ,����ҽʦ,����ҽʦ,����,����Ա���,
                ����Ա����)
            Values(
                ����ID_IN,��ҳID_IN,���ʱ��_IN,2,0,����ID_IN,����ID_IN,
                Decode(����ȼ�ID_IN,0,Null,����ȼ�ID_IN),Null,Null,
                ���λ�ʿ_IN,סԺҽʦ_IN,����ҽʦ_IN,����ҽʦ_IN,��ǰ����_IN,����Ա���_IN,����Ա����_IN);
        ELSE
            --���Ŵ�λ
            v_Count:=0;
            v_����:=����_IN||',';

            While v_���� IS NOT Null Loop
                v_��ǰ����:=to_Number(Substr(v_����,1,Instr(v_����,',')-1));
                Select �ȼ�ID Into v_�ȼ�ID From ��λ״����¼ Where ����ID=����ID_IN And ����=v_��ǰ����;

                Insert Into ���˱䶯��¼(
                    ����ID,��ҳID,��ʼʱ��,��ʼԭ��,���Ӵ�λ,����ID,����ID,����ȼ�ID,
                    ��λ�ȼ�ID,����,���λ�ʿ,����ҽʦ,����ҽʦ,����ҽʦ,����,����Ա���,����Ա����)
                Values(
                    ����ID_IN,��ҳID_IN,���ʱ��_IN,2,Decode(v_Count,0,0,1),
                    ����ID_IN,����ID_IN,Decode(����ȼ�ID_IN,0,Null,����ȼ�ID_IN),
                    v_�ȼ�ID,v_��ǰ����,���λ�ʿ_IN,סԺҽʦ_IN,����ҽʦ_IN,����ҽʦ_IN,��ǰ����_IN,����Ա���_IN,����Ա����_IN);

                Select Count(*) Into v_Count From ��λ״����¼ Where ����ID=����ID_IN And ����=v_��ǰ���� And ״̬='�մ�';

                if v_Count=0 Then
                    v_Error:='����ʧ��,��λ '||v_��ǰ����||' ���ǿմ���';
                    Raise Err_Custom;
                End IF;

                Update ��λ״����¼ Set ״̬='ռ��',����ID=����ID_IN,����ID=Decode(����,1,����ID_IN,����ID) Where ����ID=����ID_IN And ����=v_��ǰ����;

                v_����:=Substr(v_����,Instr(v_����,',')+1);
                v_Count:=v_Count+1;
            End Loop;
        End IF;
    ELSE
        --ת�����
        Select Count(*) Into v_Count From ������ҳ Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ״̬=2;

        if v_Count=0 Then
            v_Error:='���˵�ǰ������ת��״̬,�����Ѿ���ת�ƣ��������ܼ�����'||Chr(13) ||Chr(10) ||'��������������粢����������ģ���ˢ�²���״̬�����ԣ�';
            Raise Err_Custom;
        End IF;

        Select ����ID,����Ա���� Into v_����ID,v_��ֹ��Ա From ���˱䶯��¼       --��������Ҷ���ʱ,��ʱ��¼��û����,ΪNull
        Where ����ID=����ID_IN And ��ҳID=��ҳID_IN
            AND ��ʼʱ�� is Null And ��ֹʱ�� is Null;

        if v_����ID <> ����ID_IN Then                         --��������Ҷ���ʱ,������Null�ж�,���Բ�����ʾ
            v_Error:='��ǰ��ס�����벡�˵Ǽǲ�����һ��,����״̬�Ѿ�����,�������ܼ�����' ||Chr(13) ||Chr(10) ||'��������������粢����������ģ���ˢ�²���״̬�����ԣ�';
            Raise Err_Custom;
        End IF;

        --������ҳ
        Update ������ҳ
            Set ״̬=0,
                ��Ժ����=to_Number(v_����),
                ��ǰ����ID=����ID_IN,
                ��Ժ����ID=����ID_IN,
                ����ȼ�ID=Decode(����ȼ�ID_IN,0,Null,����ȼ�ID_IN),
                ��ǰ����=��ǰ����_IN,
                ���λ�ʿ=���λ�ʿ_IN,
                ����ҽʦ=����ҽʦ_IN,
                סԺҽʦ=סԺҽʦ_IN,
                �Ƿ����=�Ƿ����_IN
        Where ����ID=����ID_IN And ��ҳID=��ҳID_IN;

         Update ������ҳ�ӱ� Set ��Ϣֵ=����ҽʦ_IN Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��Ϣ��='����ҽʦ';
         IF SQL%RowCount=0 Then
             Insert Into ������ҳ�ӱ� (����ID,��ҳID,��Ϣ��,��Ϣֵ) Values (����ID_IN,��ҳID_IN,'����ҽʦ',����ҽʦ_IN);
         End IF;
         Update ������ҳ�ӱ� Set ��Ϣֵ=����ҽʦ_IN Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��Ϣ��='����ҽʦ';
         IF SQL%RowCount=0 Then
             Insert Into ������ҳ�ӱ� (����ID,��ҳID,��Ϣ��,��Ϣֵ) Values (����ID_IN,��ҳID_IN,'����ҽʦ',����ҽʦ_IN);
         End IF;


        --�˳����˵�ǰ��λ
        For r_bedrow IN c_BedInfo Loop
            Update ��λ״����¼ Set ״̬='�մ�',����ID=Null,����ID=Decode(����,1,NULL,����ID) Where ����ID=r_bedrow.����ID And ����=r_bedrow.����;
        End Loop;

        --��¼��һ������ֹ������Ա
        Update ���˱䶯��¼ 
            Set ��ֹʱ��=���ʱ��_IN,��ֹԭ��=3,��ֹ��Ա=v_��ֹ��Ա
        Where ����ID=����ID_IN And ��ҳID=��ҳID_IN
            AND ��ʼʱ�� IS NOT Null And ��ֹʱ�� is Null;

        Delete From ���˱䶯��¼
        Where ����ID=����ID_IN And ��ҳID=��ҳID_IN
            AND ��ʼԭ��=3 And ��ʼʱ�� is Null And ��ֹʱ�� is Null;

        --�µĴ�λ��¼
        if ����_IN is Null Then
            --����ͥ����
            Insert Into ���˱䶯��¼(
                ����ID,��ҳID,��ʼʱ��,��ʼԭ��,���Ӵ�λ,����ID,����ID,
                ����ȼ�ID,��λ�ȼ�ID,����,���λ�ʿ,����ҽʦ,����ҽʦ,����ҽʦ,����,����Ա���,����Ա����)
            Values(
                ����ID_IN,��ҳID_IN,���ʱ��_IN,3,0,����ID_IN,����ID_IN,
                Decode(����ȼ�ID_IN,0,Null,����ȼ�ID_IN),Null,
                Null,���λ�ʿ_IN,סԺҽʦ_IN,����ҽʦ_IN,����ҽʦ_IN,��ǰ����_IN,����Ա���_IN,����Ա����_IN);
        ELSE
            v_Count:=0;
            v_����:=����_IN||',';

            While v_���� IS NOT Null Loop
                v_��ǰ����:=to_Number(Substr(v_����,1,Instr(v_����,',')-1));
                Select �ȼ�ID Into v_�ȼ�ID From ��λ״����¼ Where ����ID=����ID_IN And ����=v_��ǰ����;

                Insert Into ���˱䶯��¼(
                    ����ID,��ҳID,��ʼʱ��,��ʼԭ��,���Ӵ�λ,����ID,����ID,
                    ����ȼ�ID,��λ�ȼ�ID,����,���λ�ʿ,����ҽʦ,����ҽʦ,����ҽʦ,����,����Ա���,����Ա����)
                Values(
                    ����ID_IN,��ҳID_IN,���ʱ��_IN,3,Decode(v_Count,0,0,1),
                    ����ID_IN,����ID_IN,Decode(����ȼ�ID_IN,0,Null,����ȼ�ID_IN),
                    v_�ȼ�ID,v_��ǰ����,���λ�ʿ_IN,סԺҽʦ_IN,����ҽʦ_IN,����ҽʦ_IN,��ǰ����_IN,����Ա���_IN,����Ա����_IN);

                Select Count(*) Into v_Count From ��λ״����¼ Where ����ID=����ID_IN And ����=v_��ǰ���� And ״̬='�մ�';

                if v_Count=0 Then
                    v_Error:='����ʧ��,��λ '||v_��ǰ����||' ���ǿմ���';
                    Raise Err_Custom;
                End IF;

                Update ��λ״����¼ Set ״̬='ռ��',����ID=����ID_IN,����ID=Decode(����,1,����ID_IN,����ID) Where ����ID=����ID_IN And ����=v_��ǰ����;

                v_����:=Substr(v_����,Instr(v_����,',')+1);
                v_Count:=v_Count+1;
            End Loop;
        End IF;
    End IF;

    --�����������
    Select Count(*) Into v_Count From ���˱䶯��¼
    Where ����ID=����ID_IN And ��ҳID=��ҳID_IN
        AND NVL(���Ӵ�λ,0)=0 And ��ʼʱ�� IS NOT Null
        AND ��ֹʱ�� is Null;

    if v_Count > 1 Then
        v_Error:='���ֲ��˴��ڷǷ��ı䶯��¼,��ǰ�������ܼ�����'||Chr(13)||Chr(10) ||'��������������粢�����������,��ˢ�²���״̬�����ԣ�';
        Raise Err_Custom;
    End IF;
Exception
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_���˱䶯��¼_InDept;
/

Create Or Replace Procedure zl_���˱䶯��¼_Undo(
    ����ID_IN        ������ҳ.����ID%Type,
    ��ҳID_IN        ������ҳ.��ҳID%Type,
    ����Ա���_IN    ���˱䶯��¼.����Ա���%Type,
    ����Ա����_IN    ���˱䶯��¼.����Ա����%Type,
    ����_IN            Varchar2:=NULL--��������,��һ���õõ�
)
AS
 -----------------------------------------------------------
 --˵����1.�����������һ�εı䶯
 --        2.ǰ�᣺�����˰���ʱ,������һ�Ŵ�λ���䶯,�����д�λ��Ӧ�����䶯
 -----------------------------------------------------------
    --Ҫ�����ı䶯��¼(�������,���ܶ���)
    Cursor c_CurLog IS
        Select * From ���˱䶯��¼ 
        Where ����ID=����ID_IN And ��ҳID=��ҳID_IN
            AND(��ֹʱ�� is Null Or ��ֹԭ��=1)
        Order by ��ֹʱ�� DESC,��ʼʱ�� DESC;
    r_CurLogRow c_CurLog%RowType;

    --������Ҫ�ָ��ı䶯��¼(�������,���ܶ���)
    Cursor c_PreLog(
        v_��ֹʱ�� ���˱䶯��¼.��ֹʱ��%Type,
        v_��ֹԭ�� ���˱䶯��¼.��ֹԭ��%Type) IS
        Select * From ���˱䶯��¼
        Where ����ID=����ID_IN
            AND ��ҳID=��ҳID_IN
            AND ��ֹʱ��=v_��ֹʱ��
            AND ��ֹԭ��=v_��ֹԭ��
        Order by ��ֹʱ�� DESC,��ʼʱ�� DESC;
    r_PreLogRow        c_PreLog%RowType;

    v_��ʼʱ��        ���˱䶯��¼.��ʼʱ��%Type;
    v_��ʼԭ��        ���˱䶯��¼.��ʼԭ��%Type;       
    v_��ֹ��Ա		���˱䶯��¼.��ֹ��Ա%Type; 

    v_�������Ҷ��� Number(1);
    v_Count            Number;
    Err_Custom        Exception;
    v_Error            Varchar2(255);
Begin
    Open c_CurLog;
    Fetch c_CurLog Into r_CurLogRow;
    If c_CurLog%RowCount=0 Then
        v_Error:='[ZLSOFT]���˵�ǰû�п��Գ����Ĳ�����[ZLSOFT]';
        Close c_CurLog;
        Raise Err_Custom;
    End IF;

    if r_CurLogRow.��ֹʱ�� is Null And r_CurLogRow.��ʼʱ�� is Null And r_CurLogRow.��ʼԭ��=3 Then
        --����ת��(��־)
        Delete From ���˱䶯��¼
        Where ����ID=����ID_IN And ��ҳID=��ҳID_IN
            AND ��ʼʱ�� is Null And ��ֹʱ�� is Null And ��ʼԭ��=3;

        Update ������ҳ Set ״̬=0 Where ����ID=����ID_IN And ��ҳID=��ҳID_IN;

        Close c_CurLog;
    Elsif r_CurLogRow.��ֹʱ�� IS NOT Null And r_CurLogRow.��ֹԭ��=1 Then
        --������Ժ
        Close c_CurLog;

        --����λ
        For r_CurLogRow IN c_CurLog Loop
            if r_CurLogRow.���� IS NOT Null Then
                Select Count(*) Into v_Count From ��λ״����¼ Where ����ID=r_CurLogRow.����ID AND ����=r_CurLogRow.���� And ״̬='�մ�';
                if v_Count=0 Then
                    v_Error:='[ZLSOFT]�ò��˳�Ժǰ����ס�Ĳ��� '||r_CurLogRow.���� ||' ��ǰ�ǿմ����Ѿ�������[ZLSOFT]';
                    Raise Err_Custom;
                End IF;

                --����ռ�ô�λ(��λ��Ϣ����б䶯,��ǿ�лָ�)
                Update ��λ״����¼
                    Set ״̬='ռ��',
                        ����ID=����ID_IN,
                        �ȼ�ID=r_CurLogRow.��λ�ȼ�ID,
                        ����ID=r_CurLogRow.����ID--ǿ�лָ���ǰ�Ŀ���,���ô�Ҳ���ô����ˡ�
                Where ����ID=r_CurLogRow.����ID And ����=r_CurLogRow.����;
            End IF;

            if NVL(r_CurLogRow.���Ӵ�λ,0)=0 Then
                Update ������Ϣ
                    Set ��Ժʱ��=Null,
                        ��ǰ����ID=r_CurLogRow.����ID,
                        ��ǰ����ID=r_CurLogRow.����ID,
                        ��ǰ����=r_CurLogRow.����
                Where ����ID=����ID_IN;
            End IF;
        End Loop;

        --�ָ���Ժ
        Update ���˱䶯��¼
            Set ��ֹʱ��=Null,��ֹԭ��=Null,��ֹ��Ա=null,
                �ϴμ���ʱ��=Null
        Where ����ID=����ID_IN And ��ҳID=��ҳID_IN
            AND ��ֹʱ�� IS NOT Null And ��ֹԭ��=1;

        Select ��ʼԭ�� Into v_��ʼԭ�� From ���˱䶯��¼
        Where ����ID=����ID_IN And ��ҳID=��ҳID_IN
            And ��ֹʱ�� is Null And Nvl(���Ӵ�λ,0)=0;

        Update ������ҳ
            Set ״̬=Decode(v_��ʼԭ��,10,3,״̬),
                ��Ժ����=Null,��Ժ��ʽ=Null,
                �����־=Null,��������=Null
         Where ����ID=����ID_IN And ��ҳID=��ҳID_IN;

        --ɾ����Ժ���
        Delete From ������ϼ�¼ Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And �������=3 AND ��¼��Դ=2;
    Elsif r_CurLogRow.��ʼԭ��=1 Then
        --�������(��Ժͬʱ���)
        v_��ʼʱ��:=r_CurLogRow.��ʼʱ��;
        Close c_CurLog;

        --�˳���ǰ��λ
        For r_CurLogRow IN c_CurLog Loop
            if r_CurLogRow.���� IS NOT Null Then
                Update ��λ״����¼ 
                    Set ״̬='�մ�',
                    ����ID=Null,
                    ����ID=Decode(����,1,NULL,����ID)
                Where ����ID=r_CurLogRow.����ID
                    AND ����=r_CurLogRow.����;
            End IF;
        End Loop;

        --�����Ϣ��ԭ
        Update ������ҳ
            Set ��Ժ����=Null,��Ժ����=Null,״̬=1
        Where ����ID=����ID_IN And ��ҳID=��ҳID_IN;

        Update ������Ϣ Set ��ǰ����=Null Where ����ID=����ID_IN;

        --�ָ��䶯(��Ժͬʱ��Ʋ����а���)
        --��Ϊ��ͬһ����¼�еĳ���,���Բ�������Ա
        Update ���˱䶯��¼
            Set ��λ�ȼ�ID=Null,����=Null,
                ���λ�ʿ=Null,����ҽʦ=Null,
                �ϴμ���ʱ��=Null
        Where ����ID=����ID_IN AND ��ҳID=��ҳID_IN
            AND ��ʼԭ��=1 AND ��ֹʱ�� is Null;
    Elsif r_CurLogRow.��ʼԭ��=2 Then
        --������Ժ���
        v_��ʼʱ��:=r_CurLogRow.��ʼʱ��;
        Close c_CurLog;

        --�˳���ǰ��λ
        For r_CurLogRow IN c_CurLog Loop
            if r_CurLogRow.���� IS NOT Null Then
                Update ��λ״����¼ 
                    Set ״̬='�մ�',
                        ����ID=Null,
                        ����ID=Decode(����,1,NULL,����ID)
                Where ����ID=r_CurLogRow.����ID 
                    And ����=r_CurLogRow.����;
            End IF;
        End Loop;

        --�����Ϣ��ԭ
        Open c_PreLog(v_��ʼʱ��,2);
        Fetch c_PreLog Into r_PreLogRow;
        Update ������ҳ Set ��Ժ����=Null,��Ժ����=Null,״̬=1,
                        ��ǰ����=r_PreLogRow.����,��Ժ����=r_PreLogRow.����,����ȼ�ID=r_PreLogRow.����ȼ�ID 
                    Where ����ID=����ID_IN And ��ҳID=��ҳID_IN;
        Close c_PreLog;
        
        Update ������Ϣ Set ��ǰ����=Null Where ����ID=����ID_IN;
        Delete ������ҳ�ӱ� Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And (��Ϣ��='����ҽʦ' Or ��Ϣ��='����ҽʦ');

        --�ָ��䶯
        Delete From ���˱䶯��¼
        Where ����ID=����ID_IN And ��ҳID=��ҳID_IN
            AND ��ʼԭ��=2 And ��ֹʱ�� is Null;

        Update ���˱䶯��¼
            Set ��ֹʱ��=Null,��ֹԭ��=Null,��ֹ��Ա=null,
                �ϴμ���ʱ��=Null
        Where ����ID=����ID_IN And ��ҳID=��ҳID_IN
            AND ��ֹԭ��=2 And ��ֹʱ��=v_��ʼʱ��;
    Elsif r_CurLogRow.��ʼԭ��=3 Then
        --����ת�����
        v_��ʼʱ��:=r_CurLogRow.��ʼʱ��;
        Close c_CurLog;

        --�˳���ǰ��λ
        For r_CurLogRow IN c_CurLog Loop
            if r_CurLogRow.���� IS NOT Null Then
                Update ��λ״����¼ 
                    Set ״̬='�մ�',
                        ����ID=Null,
                        ����ID=Decode(����,1,NULL,����ID)
                Where ����ID=r_CurLogRow.����ID 
                    And ����=r_CurLogRow.����;
            End IF;
        End Loop;

        --��鼰��ԭԭ��λ
        For r_PreLogRow IN c_PreLog(v_��ʼʱ��,3) Loop
            if r_PreLogRow.���� IS NOT Null Then
                Select Count(*) Into v_Count From ��λ״����¼ Where ����ID=r_PreLogRow.����ID AND ����=r_PreLogRow.���� And ״̬='�մ�';
                if v_Count=0 Then
                    v_Error:='[ZLSOFT]����ת��ÿ���ǰ�Ĵ�λ '||r_PreLogRow.���� ||' ��ǰ�ǿմ����Ѿ�������[ZLSOFT]';
                    Raise Err_Custom;
                End IF;

                Update ��λ״����¼
                    Set ״̬='ռ��',
                        ����ID=����ID_IN,
                        ����ID=r_PreLogRow.����ID--ǿ�лָ���ǰ�Ŀ���,���ô�Ҳ���ô����ˡ�
                Where ����ID=r_PreLogRow.����ID 
                    And ����=r_PreLogRow.����;
            End IF;

            --�����Ϣ��ԭ
            if NVL(r_PreLogRow.���Ӵ�λ,0)=0 Then
                Update ������ҳ
                    Set ״̬=2,
                        ��ǰ����ID=r_PreLogRow.����ID,
                        ��Ժ����ID=r_PreLogRow.����ID,
                        ��Ժ����=r_PreLogRow.����,
                        ����ȼ�ID=r_PreLogRow.����ȼ�ID,
                        ���λ�ʿ=r_PreLogRow.���λ�ʿ,
                        סԺҽʦ=r_PreLogRow.����ҽʦ,
                        ��ǰ����=r_CurLogRow.����
                Where ����ID=����ID_IN And ��ҳID=��ҳID_IN;

                Update ������ҳ�ӱ�
                    SET ��Ϣֵ=r_PreLogRow.����ҽʦ
                Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��Ϣ��='����ҽʦ';
                Update ������ҳ�ӱ�
                    SET ��Ϣֵ=r_PreLogRow.����ҽʦ
                Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��Ϣ��='����ҽʦ';

                Update ������Ϣ
                    Set ��ǰ����ID=r_PreLogRow.����ID,
                        ��ǰ����ID=r_PreLogRow.����ID,
                        ��ǰ����=r_PreLogRow.����
                Where ����ID=����ID_IN;
            End IF;
        End Loop;

        --�ָ��䶯(�ָ�����ʱת�Ʊ��״̬)
        Delete From ���˱䶯��¼
        Where ���Ӵ�λ=1 And ����ID=����ID_IN
            AND ��ҳID=��ҳID_IN And ��ʼԭ��=3 And ��ֹʱ�� is Null;
        	
		   Select ��ֹ��Ա Into v_��ֹ��Ա From ���˱䶯��¼
        Where ����ID=����ID_IN And ��ҳID=��ҳID_IN 
			  And ��ֹԭ��=3 And ��ֹʱ��=v_��ʼʱ�� And Nvl(���Ӵ�λ,0)=0;

        v_�������Ҷ���:=To_Number(Nvl(ZL_GetSysParameter(99),0)); 
        
        --��ʱ��¼�Ĳ���Ա��Ϣ��¼������ֹ��Ա,��Ϊû�м�¼��ֹ��Ա���,�Ͳ��ָ�
        IF v_�������Ҷ���=1 THEN 
             Update ���˱䶯��¼
                Set ��ʼʱ��=Null,����ȼ�ID=Null,
                    ��λ�ȼ�ID=Null,����=Null,
                    ���λ�ʿ=Null,����ҽʦ=Null,
                    ����Ա���=Null,����Ա����=v_��ֹ��Ա,
                    �ϴμ���ʱ��=Null,����ID=Null
            Where ����ID=����ID_IN And ��ҳID=��ҳID_IN
                AND ��ʼԭ��=3 And ��ֹʱ�� is Null;
        ELSE 
            Update ���˱䶯��¼
                Set ��ʼʱ��=Null,����ȼ�ID=Null,
                    ��λ�ȼ�ID=Null,����=Null,
                    ���λ�ʿ=Null,����ҽʦ=Null,
                    ����Ա���=Null,����Ա����=v_��ֹ��Ա,
                    �ϴμ���ʱ��=Null
            Where ����ID=����ID_IN And ��ҳID=��ҳID_IN
                AND ��ʼԭ��=3 And ��ֹʱ�� is Null;
        END if;

        Update ���˱䶯��¼
            Set ��ֹʱ��=Null,��ֹԭ��=Null,��ֹ��Ա=Null,
                �ϴμ���ʱ��=Null
        Where ����ID=����ID_IN And ��ҳID=��ҳID_IN
            AND ��ֹԭ��=3 And ��ֹʱ��=v_��ʼʱ��;
        
    Elsif r_CurLogRow.��ʼԭ��=4 Then
        --��������
        v_��ʼʱ��:=r_CurLogRow.��ʼʱ��;
        Close c_CurLog;

        --�˳���ǰ��λ
        For r_CurLogRow IN c_CurLog Loop
            if r_CurLogRow.���� IS NOT Null Then
                Update ��λ״����¼ 
                    Set ״̬='�մ�',
                        ����ID=Null,
                        ����ID=Decode(����,1,NULL,����ID)
                Where ����ID=r_CurLogRow.����ID 
                    And ����=r_CurLogRow.����;
            End IF;
        End Loop;

        --��鼰��ԭԭ��λ
        For r_PreLogRow IN c_PreLog(v_��ʼʱ��,4) Loop
            if r_PreLogRow.���� IS NOT Null Then
                Select Count(*) Into v_Count From ��λ״����¼ Where ����ID=r_PreLogRow.����ID And ����=r_PreLogRow.���� And ״̬='�մ�';
                if v_Count=0 Then
                    v_Error:='[ZLSOFT]�������һ�λ���ǰ����ס�Ĵ�λ '||r_PreLogRow.���� ||' ��ǰ�ǿմ����Ѿ�������[ZLSOFT]';
                    Raise Err_Custom;
                End IF;

                Update ��λ״����¼ 
                    Set ״̬='ռ��',
                        ����ID=����ID_IN,
                        ����ID=Decode(����,1,r_PreLogRow.����ID,����ID)
                Where ����ID=r_PreLogRow.����ID 
                    And ����=r_PreLogRow.����;
            End IF;

            --������Ϣ��������ҳ,������������Ҷ���ʱ,�����ſ��Ի�����,�˴�Ϊ���ж�,ͳһ��ԭ����
            if NVL(r_PreLogRow.���Ӵ�λ,0)=0 Then
                Update ������ҳ Set ��Ժ����=r_PreLogRow.����,��ǰ����ID=r_PreLogRow.����ID      
                Where ����ID=����ID_IN And ��ҳID=��ҳID_IN;

                Update ������Ϣ Set ��ǰ����=r_PreLogRow.����,��ǰ����ID=r_PreLogRow.����ID Where ����ID=����ID_IN;
            End IF;
        End Loop;

        --�ָ��䶯
        Delete From ���˱䶯��¼
        Where ����ID=����ID_IN And ��ҳID=��ҳID_IN
            AND ��ʼԭ��=4 And ��ֹʱ�� is Null;

        Update ���˱䶯��¼
            Set ��ֹʱ��=Null,��ֹԭ��=Null,��ֹ��Ա=Null,�ϴμ���ʱ��=Null
         Where ����ID=����ID_IN And ��ҳID=��ҳID_IN
            AND ��ֹԭ��=4 And ��ֹʱ��=v_��ʼʱ��;
    Elsif r_CurLogRow.��ʼԭ��=5 Then
        --������λ�ȼ��䶯
        --��ԭԭ��λ�ĵȼ�
        For r_PreLogRow IN c_PreLog(r_CurLogRow.��ʼʱ��,5) Loop
            if r_PreLogRow.���� IS NOT Null Then
                Update ��λ״����¼ Set �ȼ�ID=r_PreLogRow.��λ�ȼ�ID Where ����ID=r_PreLogRow.����ID And ����=r_PreLogRow.����;
            End IF;
        End Loop;
        --�ָ��䶯
        Delete From ���˱䶯��¼ Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��ʼԭ��=5 And ��ֹʱ�� is Null;

        Update ���˱䶯��¼ 
            Set ��ֹʱ��=Null,��ֹԭ��=Null,��ֹ��Ա=Null,�ϴμ���ʱ��=Null
        Where ����ID=����ID_IN And ��ҳID=��ҳID_IN
            AND ��ֹԭ��=5 And ��ֹʱ��=r_CurLogRow.��ʼʱ��;

        Close c_CurLog;
    Elsif r_CurLogRow.��ʼԭ��=6 Then
        --��������ȼ��䶯
        Open c_PreLog(r_CurLogRow.��ʼʱ��,6);
        Fetch c_PreLog Into r_PreLogRow;
        --�ָ�ԭ����ȼ�
        Update ������ҳ
            Set ����ȼ�ID=r_PreLogRow.����ȼ�ID
        Where ����ID=����ID_IN And ��ҳID=��ҳID_IN;
        --�ָ��䶯
        Delete From ���˱䶯��¼
        Where ����ID=����ID_IN And ��ҳID=��ҳID_IN
            AND ��ʼԭ��=6 And ��ֹʱ�� is Null;

        Update ���˱䶯��¼
            Set ��ֹʱ��=Null,��ֹԭ��=Null,��ֹ��Ա=Null,�ϴμ���ʱ��=Null
        Where ����ID=����ID_IN And ��ҳID=��ҳID_IN
            AND ��ֹԭ��=6 And ��ֹʱ��=r_CurLogRow.��ʼʱ��;

        Close c_PreLog;
        Close c_CurLog;
    Elsif r_CurLogRow.��ʼԭ��=7 Then
        --��������ҽʦ�ı�
        Open c_PreLog(r_CurLogRow.��ʼʱ��,7);
        Fetch c_PreLog Into r_PreLogRow;
        --�ָ�ԭҽʦ
        Update ������ҳ Set סԺҽʦ=r_PreLogRow.����ҽʦ Where ����ID=����ID_IN And ��ҳID=��ҳID_IN;
        --�ָ��䶯
        Delete From ���˱䶯��¼ Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��ʼԭ��=7 And ��ֹʱ�� is Null;

        Update ���˱䶯��¼ 
            Set ��ֹʱ��=Null,��ֹԭ��=Null,��ֹ��Ա=Null,�ϴμ���ʱ��=Null
        Where ����ID=����ID_IN And ��ҳID=��ҳID_IN 
            And ��ֹԭ��=7 And ��ֹʱ��=r_CurLogRow.��ʼʱ��;

        Close c_PreLog;
        Close c_CurLog;
    Elsif r_CurLogRow.��ʼԭ��=8 Then
        --�������λ�ʿ�ı�
        Open c_PreLog(r_CurLogRow.��ʼʱ��,8);
        Fetch c_PreLog Into r_PreLogRow;

        --�ָ�ԭ���λ�ʿ
        Update ������ҳ Set ���λ�ʿ=r_PreLogRow.���λ�ʿ
        Where ����ID=����ID_IN And ��ҳID=��ҳID_IN;
        --�ָ��䶯
        Delete From ���˱䶯��¼
        Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��ʼԭ��=8 And ��ֹʱ�� is Null;

        Update ���˱䶯��¼ 
            Set ��ֹʱ��=Null,��ֹԭ��=Null,��ֹ��Ա=Null,�ϴμ���ʱ��=Null
        Where ����ID=����ID_IN And ��ҳID=��ҳID_IN
            AND ��ֹԭ��=8 And ��ֹʱ��=r_CurLogRow.��ʼʱ��;

        Close c_PreLog;
        Close c_CurLog;
    Elsif r_CurLogRow.��ʼԭ��=9 Then
        --����תΪסԺ����

        --�ָ�ԭ���λ�ʿ
        Update ������ҳ Set ��������=2 Where ����ID=����ID_IN And ��ҳID=��ҳID_IN;
        --�ָ��䶯
        Delete From ���˱䶯��¼ Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��ʼԭ��=9 And ��ֹʱ�� is Null;

        Update ���˱䶯��¼ 
            Set ��ֹʱ��=Null,��ֹԭ��=Null,��ֹ��Ա=Null,�ϴμ���ʱ��=Null
        Where ����ID=����ID_IN And ��ҳID=��ҳID_IN
            AND ��ֹԭ��=9 And ��ֹʱ��=r_CurLogRow.��ʼʱ��;
        
        If ����_IN IS Not NULL Then 
            Update ������Ϣ Set סԺ��=NULL Where ����ID=����ID_IN;
        END if;

        Close c_CurLog;
    Elsif r_CurLogRow.��ʼԭ��=10 Then
        --����Ԥ��Ժ

        --�ָ�סԺ״̬
        Update ������ҳ Set ״̬=0 Where ����ID=����ID_IN And ��ҳID=��ҳID_IN;
        --�ָ��䶯
        Delete From ���˱䶯��¼ Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��ʼԭ��=10 And ��ֹʱ�� is Null;

        Update ���˱䶯��¼ 
            Set ��ֹʱ��=Null,��ֹԭ��=Null,�ϴμ���ʱ��=Null,��ֹ��Ա=Null
        Where ����ID=����ID_IN And ��ҳID=��ҳID_IN 
            And ��ֹԭ��=10 And ��ֹʱ��=r_CurLogRow.��ʼʱ��;

        Close c_CurLog;
    Elsif r_CurLogRow.��ʼԭ��=11 Then
        --��������ҽʦ�ı�
        Open c_PreLog(r_CurLogRow.��ʼʱ��,11);
        Fetch c_PreLog Into r_PreLogRow;
        --�ָ�ԭ����ҽʦ
        Update ������ҳ�ӱ� Set ��Ϣֵ=r_PreLogRow.����ҽʦ Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��Ϣ��='����ҽʦ';
        --�ָ��䶯
        Delete From ���˱䶯��¼ Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��ʼԭ��=11 And ��ֹʱ�� is Null;

        Update ���˱䶯��¼ 
            Set ��ֹʱ��=Null,��ֹԭ��=Null,��ֹ��Ա=Null,�ϴμ���ʱ��=Null
        Where ����ID=����ID_IN And ��ҳID=��ҳID_IN 
            And ��ֹԭ��=11 And ��ֹʱ��=r_CurLogRow.��ʼʱ��;

        Close c_PreLog;
        Close c_CurLog;
    Elsif r_CurLogRow.��ʼԭ��=12 Then
        --��������ҽʦ�ı�
        Open c_PreLog(r_CurLogRow.��ʼʱ��,12);
        Fetch c_PreLog Into r_PreLogRow;
        --�ָ�ԭ����ҽʦ
        Update ������ҳ�ӱ� Set ��Ϣֵ=r_PreLogRow.����ҽʦ Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��Ϣ��='����ҽʦ';
        --�ָ��䶯
        Delete From ���˱䶯��¼ Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��ʼԭ��=12 And ��ֹʱ�� is Null;

        Update ���˱䶯��¼ 
            Set ��ֹʱ��=Null,��ֹԭ��=Null,��ֹ��Ա=Null,�ϴμ���ʱ��=Null
        Where ����ID=����ID_IN And ��ҳID=��ҳID_IN 
            And ��ֹԭ��=12 And ��ֹʱ��=r_CurLogRow.��ʼʱ��;

        Close c_PreLog;
        Close c_CurLog;
    Elsif r_CurLogRow.��ʼԭ��=13 Then
        --��������ı�
        Open c_PreLog(r_CurLogRow.��ʼʱ��,13);
        Fetch c_PreLog Into r_PreLogRow;
        --�ָ�ԭ����
        Update ������ҳ Set ��ǰ����=r_PreLogRow.���� Where ����ID=����ID_IN And ��ҳID=��ҳID_IN ;
        --�ָ��䶯
        Delete From ���˱䶯��¼ Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��ʼԭ��=13 And ��ֹʱ�� is Null;

        Update ���˱䶯��¼ 
            Set ��ֹʱ��=Null,��ֹԭ��=Null,��ֹ��Ա=Null,�ϴμ���ʱ��=Null
        Where ����ID=����ID_IN And ��ҳID=��ҳID_IN 
            And ��ֹԭ��=13 And ��ֹʱ��=r_CurLogRow.��ʼʱ��;

        Close c_PreLog;
        Close c_CurLog;
    End IF;
Exception
    When Err_Custom Then Raise_application_error(-20101,v_Error);
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_���˱䶯��¼_Undo;
/

Create Or Replace Procedure ZL_���˱䶯��¼_PreOut(
    ����ID_IN		������ҳ.����ID%Type,
    ��ҳID_IN		������ҳ.��ҳID%Type,
	����ʱ��_IN		���˱䶯��¼.��ʼʱ��%Type	
) AS
-----------------------------------------------------------
--���ܣ������˱�ΪԤ��Ժ״̬��������һ���䶯
-----------------------------------------------------------
    Cursor c_OldInfo IS
        Select * From ���˱䶯��¼
        Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��ֹʱ�� is Null;
    r_OldInfo    c_OldInfo%RowType;
    
    v_Temp        Varchar2(255);
    v_��Ա���    ���˷��ü�¼.����Ա���%Type;
    v_��Ա����    ���˷��ü�¼.����Ա����%Type;

    v_Count        Number;
    v_Error        Varchar2(255);
    Err_Custom    Exception;
Begin
    --�����������
    Select Nvl(״̬,0) Into v_Count From ������ҳ Where ����ID=����ID_IN And ��ҳID=��ҳID_IN;
    If v_Count<>0 Then
        v_Error:='�ò��˵�ǰ����ת�ƻ���δ��ƣ�����ִ��Ԥ��Ժ��';
        Raise Err_Custom;
    End IF;
    
    --����Ա��Ϣ
    v_Temp:=zl_Identity;
    v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
    v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
    v_��Ա���:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
    v_��Ա����:=Substr(v_Temp,Instr(v_Temp,',')+1);
    
    Open c_OldInfo;--�����ڴ���֮ǰ�ȴ�

    --ȡ���ϴα䶯
    Update ���˱䶯��¼
        Set ��ֹʱ��=����ʱ��_IN,��ֹԭ��=10,��ֹ��Ա=v_��Ա����
    Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��ֹʱ�� is Null;

    --�����±䶯
    Fetch c_OldInfo Into r_OldInfo;
    If c_OldInfo%RowCount=0 Then
        Close c_OldInfo;
        v_Error:='δ���ָò��˵�ǰ��Ч�ı䶯��¼��';
        Raise Err_Custom;
    End IF;

    While c_OldInfo%Found Loop
        Insert Into ���˱䶯��¼(
            ����ID,��ҳID,��ʼʱ��,��ʼԭ��,���Ӵ�λ,����ID,����ID,
            ����ȼ�ID,��λ�ȼ�ID,����,���λ�ʿ,����ҽʦ,����ҽʦ,����ҽʦ,����,����Ա���,����Ա����)
        Values(
            ����ID_IN,��ҳID_IN,����ʱ��_IN,10,r_OldInfo.���Ӵ�λ,r_OldInfo.����ID,
            r_OldInfo.����ID,r_OldInfo.����ȼ�ID,r_OldInfo.��λ�ȼ�ID,r_OldInfo.����,
            r_OldInfo.���λ�ʿ,r_OldInfo.����ҽʦ,r_OldInfo.����ҽʦ,r_OldInfo.����ҽʦ,r_OldInfo.����,v_��Ա���,v_��Ա����);
        Fetch c_OldInfo Into r_OldInfo;
    End Loop;

    Close c_OldInfo;

    Update ������ҳ Set ״̬=3 Where ����ID=����ID_IN And ��ҳID=��ҳID_IN;

    --�����������
    Select Count(*) Into v_Count
    From ���˱䶯��¼
    Where ����ID=����ID_IN And ��ҳID=��ҳID_IN
        AND NVL(���Ӵ�λ,0)=0 And ��ʼʱ�� IS NOT Null
        AND ��ֹʱ�� is Null;
    If v_Count > 1 Then
        v_Error:='���ֲ��˴��ڷǷ��ı䶯��¼,��ǰ�������ܼ�����'||Chr(13)||Chr(10)||'��������������粢�����������,��ˢ�²���״̬�����ԣ�';
        Raise Err_Custom;
    End IF;
Exception
    When Err_Custom Then Raise_Application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_���˱䶯��¼_PreOut;
/


Create Or Replace Procedure zl_��Ժ������ҳ_Insert(
	�Ǽ�ģʽ_IN			Number,
    ��������_IN			������ҳ.��������%Type,
    ����ID_IN           ������Ϣ.����ID%Type,
    סԺ��_IN           ������Ϣ.סԺ��%Type,
	ҽ����_IN			�����ʻ�.ҽ����%Type,
    ����_IN             ������Ϣ.����%Type,
    �Ա�_IN             ������Ϣ.�Ա�%Type,
    ����_IN             ������Ϣ.����%Type,
    �ѱ�_IN             ������Ϣ.�ѱ�%Type,
    ��������_IN         ������Ϣ.��������%Type,
    ����_IN             ������Ϣ.����%Type,
    ����_IN             ������Ϣ.����%Type,
    ѧ��_IN             ������Ϣ.ѧ��%Type,
    ����״��_IN         ������Ϣ.����״��%Type,
    ְҵ_IN             ������Ϣ.ְҵ%Type,
    ���_IN             ������Ϣ.���%Type,
    ���֤��_IN         ������Ϣ.���֤��%Type,
    �����ص�_IN         ������Ϣ.�����ص�%Type,
    ��ͥ��ַ_IN         ������Ϣ.��ͥ��ַ%Type,
    �����ʱ�_IN         ������Ϣ.�����ʱ�%Type,
    ��ͥ�绰_IN         ������Ϣ.��ͥ�绰%Type,
    ��ϵ������_IN       ������Ϣ.��ϵ������%Type,
    ��ϵ�˹�ϵ_IN       ������Ϣ.��ϵ�˹�ϵ%Type,
    ��ϵ�˵�ַ_IN       ������Ϣ.��ϵ�˵�ַ%Type,
    ��ϵ�˵绰_IN       ������Ϣ.��ϵ�˵绰%Type,
    ������λ_IN         ������Ϣ.������λ%Type,
    ��ͬ��λID_IN       ������Ϣ.��ͬ��λID%Type,
    ��λ�绰_IN         ������Ϣ.��λ�绰%Type,
    ��λ�ʱ�_IN         ������Ϣ.��λ�ʱ�%Type,
    ��λ������_IN       ������Ϣ.��λ������%Type,
    ��λ�ʺ�_IN         ������Ϣ.��λ�ʺ�%Type,
    ������_IN           ������Ϣ.������%Type,
    ������_IN           ������Ϣ.������%Type,
	��������_IN			������Ϣ.��������%Type,
    ��Ժ����ID_IN       ������ҳ.��Ժ����ID%Type,
    ����ȼ�ID_IN       ������ҳ.����ȼ�ID%Type,
    ��Ժ����_IN         ������ҳ.��Ժ����%Type,
    ��Ժ��ʽ_IN         ������ҳ.��Ժ��ʽ%Type,
    סԺĿ��_IN         ������ҳ.סԺĿ��%Type,
    ����Ժת��_IN       ������ҳ.����Ժת��%Type,
    ����ҽʦ_IN         ������ҳ.����ҽʦ%Type,
    ����_IN             ������ҳ.����%Type,
    ��Ժʱ��_IN         ������ҳ.��Ժ����%Type,
    �Ƿ����_IN         ������ҳ.�Ƿ����%Type,
    ����_IN             ������ҳ.��Ժ����%Type,
    ���ʽ_IN         ������ҳ.ҽ�Ƹ��ʽ%Type,
    ����ID_IN           ������ϼ�¼.����ID%Type,
    �������_IN         ������ϼ�¼.�������%Type,
    ��ҽ����ID_IN       ������ϼ�¼.����ID%Type,
    ��ҽ���_IN			������ϼ�¼.�������%Type,
    ����_IN             ������ҳ.����%Type,
    ����Ա���_IN       ������ҳ.��ĿԱ���%Type,
    ����Ա����_IN       ������ҳ.��ĿԱ����%Type,
    �²���_IN           Number:=1,
	��ע_IN				������ҳ.��ע%Type:=Null,
    ��Ժ����ID_IN       ������ҳ.��Ժ����ID%Type:=Null
) AS
-----------------------------------------------------------
--���ܣ�����Ժ��������һ�Ų�����ҳ��ͬʱ���ܴ�����ơ�
--������
--      �Ǽ�ģʽ_IN=0-�����Ǽ�,1-ԤԼ�Ǽ�,2-����ԤԼ(�²���_IN=0)
--      ��������_IN=��Ӧ"������ҳ.��������"
--      ����_IN=Null:��ͬʱ���;0:�����ͥ����,��Ϊ��;����:������崲λ��
--      �²���_IN=��������е����Ĳ�����Ժ,��ò���Ϊ0��ȱʡΪ�²���
--      ��Ժ����ID_IN=ֻ�е�ʹ��[����������]ģʽ(������99)ʱ,������Ժͬʱ��Ʒִ�ʱ,����ֵ
-----------------------------------------------------------
    v_��ҳID	                ������ҳ.��ҳID%Type;
    v_����ID	                ������ҳ.��Ժ����ID%Type;
    v_�ȼ�ID                  ��λ״����¼.�ȼ�ID%Type;
   	v_�������Ҷ���     Number(1);

    v_Count     Number;
    v_Date      Date;
    v_Error     Varchar2(255);
    Err_Custom  Exception;
Begin
    Select Sysdate Into v_Date From Dual;

    --���˻�����Ϣ
    IF �²���_IN=1 Then
        Insert Into ������Ϣ(
            ����ID,סԺ��,����,�Ա�,����,�ѱ�,ҽ�Ƹ��ʽ,��������,����,����,����,ѧ��,
            ����״��,ְҵ,���,���֤��,�����ص�,��ͥ��ַ,�����ʱ�,��ͥ�绰,��ϵ������,
            ��ϵ�˹�ϵ,��ϵ�˵�ַ,��ϵ�˵绰,������λ,��ͬ��λID,��λ�绰,��λ�ʱ�,
            ��λ������,��λ�ʺ�,������,������,��������,����,�Ǽ�ʱ��)
        Values(
            ����ID_IN,סԺ��_IN,����_IN,�Ա�_IN,����_IN,�ѱ�_IN,���ʽ_IN,��������_IN,
            ����_IN,����_IN,����_IN,ѧ��_IN,����״��_IN,ְҵ_IN,���_IN,���֤��_IN,�����ص�_IN,
            ��ͥ��ַ_IN,�����ʱ�_IN,��ͥ�绰_IN,��ϵ������_IN,��ϵ�˹�ϵ_IN,��ϵ�˵�ַ_IN,
            ��ϵ�˵绰_IN,������λ_IN,Decode(��ͬ��λID_IN,0,Null,��ͬ��λID_IN),��λ�绰_IN,
            ��λ�ʱ�_IN,��λ������_IN,��λ�ʺ�_IN,������_IN,Decode(������_IN,0,Null,������_IN),
            ��������_IN,����_IN,��Ժʱ��_IN);
    Else
        --�ϲ��˵�����ѱ𲻱�,�������������۲���
        Update ������Ϣ
            Set סԺ��=סԺ��_IN,����=����_IN,
                �Ա�=�Ա�_IN,����=����_IN,
				�ѱ�=Decode(��������_IN,1,�ѱ�_IN,�ѱ�),
                ҽ�Ƹ��ʽ=���ʽ_IN,
                ��������=��������_IN,����=����_IN,
                ����=����_IN,����=����_IN,ѧ��=ѧ��_IN,
                ����״��=����״��_IN,ְҵ=ְҵ_IN,
                ���=���_IN,���֤��=���֤��_IN,
                �����ص�=�����ص�_IN,��ͥ��ַ=��ͥ��ַ_IN,
                �����ʱ�=�����ʱ�_IN,��ͥ�绰=��ͥ�绰_IN,
                ��ϵ������=��ϵ������_IN,��ϵ�˹�ϵ=��ϵ�˹�ϵ_IN,
                ��ϵ�˵�ַ=��ϵ�˵�ַ_IN,��ϵ�˵绰=��ϵ�˵绰_IN,
                ������λ=������λ_IN,��ͬ��λID=Decode(��ͬ��λID_IN,0,Null,��ͬ��λID_IN),
                ��λ�绰=��λ�绰_IN,��λ�ʱ�=��λ�ʱ�_IN,
                ��λ������=��λ������_IN,��λ�ʺ�=��λ�ʺ�_IN,
                ������=������_IN,������=Decode(������_IN,0,Null,������_IN),
                ��������=��������_IN,����=����_IN
        Where ����ID=����ID_IN;
    End if;

    --סԺ������¼:ԤԼʱ������,����ʱ�Ų���,��������ԤԼסԺ��
	If �Ǽ�ģʽ_IN<>1 Then
		If סԺ��_IN IS Not NULL Then
			Update סԺ������¼
				Set ������=סԺ��_IN,�������='һ��',�洢״̬='��Ժ'
			Where ����ID=����ID_IN;
			If SQL%RowCount=0 Then
				Insert Into סԺ������¼(
					����ID,������,�������,�洢״̬,��������)
				Values(
					����ID_IN,סԺ��_IN,'һ��','��Ժ',��Ժʱ��_IN);
			End IF;
		Else
			Delete From סԺ������¼ Where ����ID=����ID_IN;
		End IF;
	End IF;

    --������Ϣ
    Begin
        If �Ǽ�ģʽ_IN=1 Then
            v_��ҳID:=0;--ԤԼ�ǼǼ�¼����ҳID=0
        Else
            Select Nvl(Max(��ҳID),0)+1 Into v_��ҳID From ������ҳ Where ����ID=����ID_IN And Nvl(��ҳID,0)<>0;
        End IF;
        
         --�������Ҷ���ģʽ,������Ժͬʱ����䴲,��Ҫ�ڽ���ѡ����,���û�зִ�,�����
        v_�������Ҷ���:=To_Number(Nvl(ZL_GetSysParameter(99),0)); 
        IF v_�������Ҷ���=1 Then 
            v_����ID :=��Ժ����ID_IN;
        Else        
           Select DISTINCT ����ID Into v_����ID From ��λ״����¼ Where ����ID=��Ժ����ID_IN;        
        End If;
    Exception
        When OTHERS Then Null;
    End;
	
	If �Ǽ�ģʽ_IN<>1 Then
		Update ������Ϣ
			Set סԺ����=v_��ҳID,��ǰ����ID=v_����ID,
				��ǰ����ID=��Ժ����ID_IN,��ǰ����=Decode(����_IN,Null,Null,0,Null,����_IN),
				��Ժʱ��=��Ժʱ��_IN,��Ժʱ��=Null
		Where ����ID=����ID_IN;
	End IF;

    --״̬��0-������Ժ,1-�ȴ����,2-�ȴ�ת��
	IF �Ǽ�ģʽ_IN=2 Then
		--����ԤԼ
		Update ������ҳ
			Set ��ҳID=v_��ҳID,��������=��������_IN,--��ҳID���,�������ʿ��ܱ��
				�ѱ�=�ѱ�_IN,��Ժ����ID=v_����ID,
				��Ժ����ID=��Ժ����ID_IN,��Ժ����=��Ժʱ��_IN,
				��Ժ����=��Ժ����_IN,��Ժ��ʽ=��Ժ��ʽ_IN,
				����Ժת��=����Ժת��_IN,סԺĿ��=סԺĿ��_IN,
				��Ժ����=Decode(����_IN,Null,Null,0,Null,����_IN),
				�Ƿ����=�Ƿ����_IN,
				��ǰ����=��Ժ����_IN,��ǰ����ID=v_����ID,
				����ȼ�ID=Decode(����ȼ�ID_IN,0,Null,����ȼ�ID_IN),
				��Ժ����ID=��Ժ����ID_IN,
				��Ժ����=Decode(����_IN,Null,Null,0,Null,����_IN),
				����ҽʦ=����ҽʦ_IN,
				��ĿԱ���=����Ա���_IN,��ĿԱ����=����Ա����_IN,
				����=����_IN,����״��=����״��_IN,
				ְҵ=ְҵ_IN,����=����_IN,
				ѧ��=ѧ��_IN,��λ�绰=��λ�绰_IN,
				��λ�ʱ�=��λ�ʱ�_IN,��λ��ַ=������λ_IN,
				����=����_IN,��ͥ��ַ=��ͥ��ַ_IN,
				��ͥ�绰=��ͥ�绰_IN,�����ʱ�=�����ʱ�_IN,
				��ϵ������=��ϵ������_IN,��ϵ�˹�ϵ=��ϵ�˹�ϵ_IN,
				��ϵ�˵�ַ=��ϵ�˵�ַ_IN,��ϵ�˵绰=��ϵ�˵绰_IN,
				ҽ�Ƹ��ʽ=���ʽ_IN,��ע=��ע_IN,
				����=����_IN,״̬=Decode(����_IN,Null,1,0),
				�Ǽ���=����Ա����_IN,�Ǽ�ʱ��=v_Date
		Where ����ID=����ID_IN And Nvl(��ҳID,0)=0;
	Else
		--��Ժ�Ǽǻ�ԤԼ�Ǽ�
		Insert Into ������ҳ(
			��������,����ID,��ҳID,�ѱ�,��Ժ����ID,��Ժ����ID,��Ժ����,��Ժ����,��Ժ��ʽ,����Ժת��,סԺĿ��,
			��Ժ����,�Ƿ����,��ǰ����,��ǰ����ID,����ȼ�ID,��Ժ����ID,��Ժ����,����ҽʦ,��ĿԱ���,
			��ĿԱ����,״̬,����,����״��,ְҵ,����,ѧ��,��λ�绰,��λ�ʱ�,��λ��ַ,����,��ͥ��ַ,
			��ͥ�绰,�����ʱ�,��ϵ������,��ϵ�˹�ϵ,��ϵ�˵�ַ,��ϵ�˵绰,ҽ�Ƹ��ʽ,����,��ע,�Ǽ���,�Ǽ�ʱ��)
		Values(
			��������_IN,����ID_IN,v_��ҳID,�ѱ�_IN,v_����ID,��Ժ����ID_IN,��Ժʱ��_IN,��Ժ����_IN,��Ժ��ʽ_IN,
			����Ժת��_IN,סԺĿ��_IN,Decode(����_IN,Null,Null,0,Null,����_IN),�Ƿ����_IN,
			��Ժ����_IN,v_����ID,Decode(����ȼ�ID_IN,0,Null,����ȼ�ID_IN),��Ժ����ID_IN,
			Decode(����_IN,Null,Null,0,Null,����_IN),����ҽʦ_IN,����Ա���_IN,����Ա����_IN,
			Decode(����_IN,Null,1,0),����_IN,����״��_IN,ְҵ_IN,����_IN,ѧ��_IN,��λ�绰_IN,
			��λ�ʱ�_IN,������λ_IN,����_IN,��ͥ��ַ_IN,��ͥ�绰_IN,�����ʱ�_IN,��ϵ������_IN,
			��ϵ�˹�ϵ_IN,��ϵ�˵�ַ_IN,��ϵ�˵绰_IN,���ʽ_IN,����_IN,��ע_IN,����Ա����_IN,v_Date);
	End If;
	
	--ҽ����
	If �Ǽ�ģʽ_IN<>1 Then
		If ҽ����_IN IS Not Null Then
			Insert Into ������ҳ�ӱ�(
				����ID,��ҳID,��Ϣ��,��Ϣֵ)
			Values(
				����ID_IN,v_��ҳID,'ҽ����',ҽ����_IN);
		End IF;

		--���˱䶯��¼
		--ͬʱ����ҷǼ�ͥ����ʱ�еȼ�
		if Nvl(����_IN,0) <> 0 Then
			Select �ȼ�ID Into v_�ȼ�ID From ��λ״����¼ Where ����ID=v_����ID And ����=����_IN;
		End IF;

		--���ͬʱ���,����Ժ�������д��һ����Ժ�䶯
		Insert Into ���˱䶯��¼(
			����ID,��ҳID,��ʼʱ��,��ʼԭ��,���Ӵ�λ,����ID,����ID,����ȼ�ID,��λ�ȼ�ID,����,����,����Ա���,����Ա����)
		Values(
			����ID_IN,v_��ҳID,��Ժʱ��_IN,1,0,v_����ID,��Ժ����ID_IN,Decode(����ȼ�ID_IN,0,Null,����ȼ�ID_IN),
			v_�ȼ�ID,Decode(����_IN,0,Null,����_IN),��Ժ����_IN,����Ա���_IN,����Ա����_IN);

		--ͬʱ����ҷǼ�ͥ����ʱ��λ��ռ��
		If Nvl(����_IN,0) <> 0 Then
			Select Count(*) Into v_Count From ��λ״����¼ Where ����ID=v_����ID And ����=����_IN And ״̬='�մ�';

			if v_Count=0 Then
				v_Error:='����ʧ��,��λ '||����_IN||' ���ǿմ���';
				Raise Err_Custom;
			End IF;

			Update ��λ״����¼ Set ״̬='ռ��',����ID=����ID_IN,����ID=Decode(����,1,��Ժ����ID_IN,����ID) Where ����ID=v_����ID And ����=����_IN;
		End IF;

		--������ϼ�¼
		If �������_IN IS Not Null Or ����ID_IN IS Not NULL Then
			Insert Into ������ϼ�¼(
				ID,����ID,��ҳID,��¼��Դ,�������,��ϴ���,����ID,�������,��¼����,��¼��) 
			Values(
				������ϼ�¼_ID.Nextval,����ID_IN,v_��ҳID,2,1,1,����ID_IN,�������_IN,sysdate,����Ա����_IN);
		End IF;        
		If ��ҽ���_IN IS Not Null Or ��ҽ����ID_IN IS Not NULL Then
			Insert Into ������ϼ�¼(
				ID,����ID,��ҳID,��¼��Դ,�������,��ϴ���,����ID,�������,��¼����,��¼��) 
			Values(
				������ϼ�¼_ID.Nextval,����ID_IN,v_��ҳID,2,11,1,��ҽ����ID_IN,��ҽ���_IN,sysdate,����Ա����_IN);
		End IF;        

		--�����������
		Select Count(*) Into v_Count From ������ҳ Where ����ID=����ID_IN And ��Ժ���� is Null;
		If v_Count>1 Then
			v_Error:='���ֲ��˴��ڷǷ��Ĳ�����¼,��ǰ�������ܼ�����'||Chr(13)||Chr(10) ||'��������������粢�����������,��ˢ�²���״̬�����ԣ�';
			Raise Err_Custom;
		End IF;

		Select Count(*) Into v_Count
			From ���˱䶯��¼
			Where ����ID=����ID_IN And ��ҳID=v_��ҳID And Nvl(���Ӵ�λ,0)=0
				And ��ʼʱ�� IS Not Null And ��ֹʱ�� is Null;
		If v_Count>1 Then
			v_Error:='���ֲ��˴��ڷǷ��ı䶯��¼,��ǰ�������ܼ�����'||Chr(13)||Chr(10) ||'��������������粢�����������,��ˢ�²���״̬�����ԣ�';
			Raise Err_Custom;
		End IF;
	End IF;
Exception
    When Err_Custom Then Raise_application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_��Ժ������ҳ_Insert;
/


Create Or Replace Procedure zl_��Ժ������ҳ_Update(
	�Ǽ�ģʽ_IN		Number,
	����ID_IN		������Ϣ.����ID%Type,
    סԺ��_IN       ������Ϣ.סԺ��%Type,
	ҽ����_IN		�����ʻ�.ҽ����%Type,
    ����_IN         ������Ϣ.����%Type,
    �Ա�_IN         ������Ϣ.�Ա�%Type,
    ����_IN         ������Ϣ.����%Type,
    �ѱ�_IN         ������Ϣ.�ѱ�%Type,
    ��������_IN     ������Ϣ.��������%Type,
    ����_IN         ������Ϣ.����%Type,
    ����_IN         ������Ϣ.����%Type,
    ѧ��_IN         ������Ϣ.ѧ��%Type,
    ����״��_IN     ������Ϣ.����״��%Type,
    ְҵ_IN         ������Ϣ.ְҵ%Type,
    ���_IN         ������Ϣ.���%Type,
    ���֤��_IN     ������Ϣ.���֤��%Type,
    �����ص�_IN     ������Ϣ.�����ص�%Type,
    ��ͥ��ַ_IN     ������Ϣ.��ͥ��ַ%Type,
    �����ʱ�_IN     ������Ϣ.�����ʱ�%Type,
    ��ͥ�绰_IN     ������Ϣ.��ͥ�绰%Type,
    ��ϵ������_IN   ������Ϣ.��ϵ������%Type,
    ��ϵ�˹�ϵ_IN   ������Ϣ.��ϵ�˹�ϵ%Type,
    ��ϵ�˵�ַ_IN   ������Ϣ.��ϵ�˵�ַ%Type,
    ��ϵ�˵绰_IN   ������Ϣ.��ϵ�˵绰%Type,
    ������λ_IN     ������Ϣ.������λ%Type,
    ��ͬ��λID_IN   ������Ϣ.��ͬ��λID%Type,
    ��λ�绰_IN     ������Ϣ.��λ�绰%Type,
    ��λ�ʱ�_IN     ������Ϣ.��λ�ʱ�%Type,
    ��λ������_IN   ������Ϣ.��λ������%Type,
    ��λ�ʺ�_IN     ������Ϣ.��λ�ʺ�%Type,
    ������_IN       ������Ϣ.������%Type,
    ������_IN       ������Ϣ.������%Type,
	��������_IN		������Ϣ.��������%Type,
    ��ҳID_IN       ������ҳ.��ҳID%Type,
    ��Ժ����ID_IN   ������ҳ.��Ժ����ID%Type,
    ����ȼ�ID_IN   ������ҳ.����ȼ�ID%Type,
    ��Ժ����_IN     ������ҳ.��Ժ����%Type,
    ��Ժ��ʽ_IN     ������ҳ.��Ժ��ʽ%Type,
    סԺĿ��_IN     ������ҳ.סԺĿ��%Type,
    ����Ժת��_IN   ������ҳ.����Ժת��%Type,
    ����ҽʦ_IN     ������ҳ.����ҽʦ%Type,
    ����_IN         ������ҳ.����%Type,
    ��Ժʱ��_IN     ������ҳ.��Ժ����%Type,
    ���ʽ_IN     ������ҳ.ҽ�Ƹ��ʽ%Type,
    ����ID_IN       ������ϼ�¼.����ID%Type,
    �������_IN     ������ϼ�¼.�������%Type,
    ��ҽ����ID_IN   ������ϼ�¼.����ID%Type,
    ��ҽ���_IN     ������ϼ�¼.�������%Type,
    ����Ա���_IN   ������ҳ.��ĿԱ���%Type,
    ����Ա����_IN   ������ҳ.��ĿԱ����%Type,
	��ע_IN			������ҳ.��ע%Type:=Null,
    ����ID_IN       ������ҳ.��Ժ����Id%Type:=Null
) AS
-----------------------------------------------------------
--˵������������������Ժδ��ƵǼǲ�����Ϣ���޸�
--      �Ǽ�ģʽ_IN=0-�����Ǽ�,1-ԤԼ�Ǽ�
--      ����ID_IN=ֻ�е�����������ģʽ��,��Ժʱ���ʱ,�Ż���ֵ
-----------------------------------------------------------    
	v_����ID				    ������ҳ.��Ժ����ID%Type;
	v_�ȼ�ID				    ��λ״����¼.�ȼ�ID%Type;
	v_��������			    ������ҳ.��������%Type;
	v_�������Ҷ���     Number(1);
    
    v_Count			Number;
    v_Date			Date;
    v_Error			Varchar2(255);
    Err_Custom		Exception;
Begin
    --�жϲ����Ƿ�δ��Ժ
    Select Count(*) Into v_Count From ������ҳ Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ״̬=1;
    if v_Count=0 Then
        v_Error:='���˵�ǰ�����ڵȴ����״̬���������ܼ�����'||Chr(13)||Chr(10)||'���ܸò����Ѿ�����������Աȡ���Ǽǻ���䴲λ��';
        Raise Err_Custom;
    End IF;
	
    Select Sysdate Into v_Date From Dual;
	Select �������� Into v_�������� From ������ҳ Where ����ID=����ID_IN And ��ҳID=��ҳID_IN;
	
    --�������Ҷ���ģʽ,������Ժͬʱ����䴲,��Ҫ�ڽ���ѡ����,���û�зִ�,�����
    v_�������Ҷ���:=To_Number(Nvl(ZL_GetSysParameter(99),0));    
    If v_�������Ҷ���=1 Then
        v_����ID:=����ID_IN;
    Else        
        Begin
	        Select Distinct ����ID Into v_����ID From ��λ״����¼ Where ����ID=��Ժ����ID_IN;
        Exception
             When OTHERS Then Null;
        End;
    End If;

    --���˻�����Ϣ
    --�ǵ�һ����Ժʱ,����ѱ𱣳ֲ���,�������������۲���
	Update ������Ϣ
        Set סԺ��=סԺ��_IN,����=����_IN,�Ա�=�Ա�_IN,
            ����=����_IN,ҽ�Ƹ��ʽ=���ʽ_IN,
            �ѱ�=Decode(v_��������,1,�ѱ�_IN,�ѱ�),
            ��������=��������_IN,����=����_IN,����=����_IN,
            ����=����_IN,ѧ��=ѧ��_IN,����״��=����״��_IN,ְҵ=ְҵ_IN,
            ���=���_IN,���֤��=���֤��_IN,�����ص�=�����ص�_IN,
            ��ͥ��ַ=��ͥ��ַ_IN,�����ʱ�=�����ʱ�_IN,��ͥ�绰=��ͥ�绰_IN,
            ��ϵ������=��ϵ������_IN,��ϵ�˹�ϵ=��ϵ�˹�ϵ_IN,
            ��ϵ�˵�ַ=��ϵ�˵�ַ_IN,��ϵ�˵绰=��ϵ�˵绰_IN,
            ������λ=������λ_IN,��ͬ��λID=Decode(��ͬ��λID_IN,0,Null,��ͬ��λID_IN),
            ��λ�绰=��λ�绰_IN,��λ�ʱ�=��λ�ʱ�_IN,��λ������=��λ������_IN,
            ��λ�ʺ�=��λ�ʺ�_IN,������=������_IN,
			������=Decode(������_IN,0,Null,������_IN),��������=��������_IN
	Where ����ID=����ID_IN;

	If �Ǽ�ģʽ_IN=0 Then
		--סԺ������¼
		If סԺ��_IN IS Not NULL Then
			Update סԺ������¼ Set ������=סԺ��_IN Where ����ID=����ID_IN;
		Else
			Delete From סԺ������¼ Where ����ID=����ID_IN;
		End IF;

		--������Ϣ
		Update ������Ϣ
			Set ��ǰ����ID=v_����ID,��ǰ����ID=��Ժ����ID_IN,
				��Ժʱ��=��Ժʱ��_IN,��Ժʱ��=Null
		Where ����ID=����ID_IN;
	End If;

    --�޸Ĳ�����ҳ
    Update ������ҳ
        Set �ѱ�=�ѱ�_IN,��Ժ����ID=v_����ID,
            ��Ժ����ID=��Ժ����ID_IN,��Ժ����=��Ժʱ��_IN,
            ��Ժ����=��Ժ����_IN,��Ժ��ʽ=��Ժ��ʽ_IN,
            ����Ժת��=����Ժת��_IN,סԺĿ��=סԺĿ��_IN,
            ��ǰ����=��Ժ����_IN,��ǰ����ID=v_����ID,
            ����ȼ�ID=Decode(����ȼ�ID_IN,0,Null,����ȼ�ID_IN),
            ��Ժ����ID=��Ժ����ID_IN,����ҽʦ=����ҽʦ_IN,
            ��ĿԱ���=����Ա���_IN,��ĿԱ����=����Ա����_IN,
            ����=����_IN,����״��=����״��_IN,
            ְҵ=ְҵ_IN,����=����_IN,
            ѧ��=ѧ��_IN,��λ�绰=��λ�绰_IN,
            ��λ�ʱ�=��λ�ʱ�_IN,��λ��ַ=������λ_IN,
            ����=����_IN,��ͥ��ַ=��ͥ��ַ_IN,
            ��ͥ�绰=��ͥ�绰_IN,�����ʱ�=�����ʱ�_IN,
            ��ϵ������=��ϵ������_IN,��ϵ�˹�ϵ=��ϵ�˹�ϵ_IN,
            ��ϵ�˵�ַ=��ϵ�˵�ַ_IN,��ϵ�˵绰=��ϵ�˵绰_IN,
            ҽ�Ƹ��ʽ=���ʽ_IN,��ע=��ע_IN,
            �Ǽ���=����Ա����_IN,�Ǽ�ʱ��=v_Date
    Where ����ID=����ID_IN And Nvl(��ҳID,0)=Nvl(��ҳID_IN,0);
	
	If �Ǽ�ģʽ_IN=0 Then
		--ҽ����
		If ҽ����_IN IS Not NULL Then
			Update ������ҳ�ӱ� Set ��Ϣֵ=ҽ����_IN Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��Ϣ��='ҽ����';
			If SQL%RowCount=0 Then
				Insert Into ������ҳ�ӱ�(
					����ID,��ҳID,��Ϣ��,��Ϣֵ)
				Values(
					����ID_IN,��ҳID_IN,'ҽ����',ҽ����_IN);
			End IF;
		Else
			Delete From ������ҳ�ӱ� Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��Ϣ��='ҽ����';
		End IF;

		--�޸Ĳ��˱䶯��¼(�϶�Ϊ��Ժ�䶯;������ƵĲ�׼�޸�,��Ժͬʱ��ƵĲ��˽�������ֹ�޸�)
		Update ���˱䶯��¼
			Set ��ʼʱ��=��Ժʱ��_IN,����ID=v_����ID,����ID=��Ժ����ID_IN,
				����ȼ�ID=Decode(����ȼ�ID_IN,0,Null,����ȼ�ID_IN),����=��Ժ����_IN,
				����Ա���=����Ա���_IN,����Ա����=����Ա����_IN,�ϴμ���ʱ��=Null
		Where ����ID=����ID_IN And ��ҳID=��ҳID_IN
			And ��ֹʱ�� is Null And ��ʼԭ��=1;
		
		--�����������
		If �������_IN is Null AND ����ID_IN IS NULL Then
			Delete From ������ϼ�¼ Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And �������=1 And ��¼��Դ=2;
		Else
			Update ������ϼ�¼ Set ����ID=����ID_IN,�������=�������_IN,��¼����=sysdate,��¼��=����Ա����_IN
            Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And �������=1  And ��¼��Դ=2;
			IF SQL%RowCount=0 Then
				Insert Into ������ϼ�¼(
					ID,����ID,��ҳID,��¼��Դ,�������,��ϴ���,����ID,�������,��¼����,��¼��) 
				Values(
					������ϼ�¼_ID.Nextval ,����ID_IN,��ҳID_IN,2,1,1,����ID_IN,�������_IN,sysdate,����Ա����_IN);
			End IF;
		End IF;

		--������ҽ���
		If ��ҽ���_IN is Null AND ��ҽ����ID_IN IS NULL Then
			Delete From ������ϼ�¼ Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And �������=11  And ��¼��Դ=2;
		Else
			Update ������ϼ�¼ Set ����ID=��ҽ����ID_IN,�������=��ҽ���_IN ,��¼����=sysdate,��¼��=����Ա����_IN
            Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And �������=11  And ��¼��Դ=2;
			IF SQL%RowCount=0 Then
				Insert Into ������ϼ�¼(
					ID,����ID,��ҳID,��¼��Դ,�������,��ϴ���,����ID,�������,��¼����,��¼��) 
				Values(
					������ϼ�¼_ID.Nextval,����ID_IN,��ҳID_IN,2,11,1,��ҽ����ID_IN,��ҽ���_IN,sysdate,����Ա����_IN);
			End IF;
		End IF;
	End IF;
Exception
    When Err_Custom Then Raise_Application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_��Ժ������ҳ_UpDate;
/

Create Or Replace Procedure zl_סԺ������ҳ_Update(
    ����ID_IN		������ҳ.����ID%Type,
    ��ҳID_IN       ������ҳ.��ҳID%Type,
    ����_IN         ������ҳ.����%Type,
    �ѱ�_IN         ������ҳ.�ѱ�%Type,
    ����״��_IN     ������ҳ.����״��%Type,
    ѧ��_IN         ������ҳ.ѧ��%Type,
    ְҵ_IN         ������ҳ.ְҵ%Type,
    ��ǰ����_IN     ������ҳ.��ǰ����%Type,
    ��λ��ַ_IN     ������ҳ.��λ��ַ%Type,
    ��ͬ��λID_IN   ������Ϣ.��ͬ��λID%Type,
    ��λ�绰_IN     ������ҳ.��λ�绰%Type,
    ��λ�ʱ�_IN     ������ҳ.��λ�ʱ�%Type,
    ��ͥ��ַ_IN     ������ҳ.��ͥ��ַ%Type,
    ��ͥ�绰_IN     ������ҳ.��ͥ�绰%Type,
    �����ʱ�_IN     ������ҳ.�����ʱ�%Type,
    ��ϵ������_IN   ������ҳ.��ϵ������%Type,
    ��ϵ�˹�ϵ_IN   ������ҳ.��ϵ�˹�ϵ%Type,
    ��ϵ�˵绰_IN   ������ҳ.��ϵ�˵绰%Type,
    ��ϵ�˵�ַ_IN   ������ҳ.��ϵ�˵�ַ%Type,
    ���λ�ʿ_IN     ������ҳ.���λ�ʿ%Type,
    ����ҽʦ_IN     ������ҳ.����ҽʦ%Type,
    סԺҽʦ_IN     ������ҳ.סԺҽʦ%Type,
    ����ID_IN       ������ϼ�¼.����ID%Type,
    ��Ժ���_IN     ������ϼ�¼.�������%Type,
    ��ҽ����ID_IN   ������ϼ�¼.����ID%Type,
    ��ҽ���_IN     ������ϼ�¼.�������%Type,
    ����Ա���_IN   ������ҳ.��ĿԱ���%Type,
    ����Ա����_IN   ������ҳ.��ĿԱ����%Type,
    ����ҽʦ_IN       ������ҳ.סԺҽʦ%Type,
    ����ҽʦ_IN       ������ҳ.סԺҽʦ%Type
)
AS
    Cursor c_OldInfo IS
        Select * From ���˱䶯��¼ Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��ֹʱ�� is Null;
    r_OldInfo c_OldInfo%RowType;

    v_���λ�ʿ		������ҳ.���λ�ʿ%Type;
    v_סԺҽʦ		������ҳ.סԺҽʦ%Type;
    v_����ҽʦ		������ҳ.סԺҽʦ%Type;
    v_����ҽʦ		������ҳ.סԺҽʦ%Type;
	v_��������		������ҳ.��������%Type;
    v_��ǰ����      ������ҳ.��ǰ����%Type;
    v_ԭ��			���˱䶯��¼.��ʼԭ��%Type;
    v_Count         Number;
    v_CurDate		Date;
    v_Error			Varchar2(255);
    Err_Custom		Exception;
Begin
    --ȡ����ǰ������(��NoneData���µıȽ�)
    Select 
		��������,Nvl(���λ�ʿ,'NoneData'),Nvl(סԺҽʦ,'NoneData'),Nvl(��ǰ����,'NoneData') Into v_��������,v_���λ�ʿ,v_סԺҽʦ,v_��ǰ����
    From ������ҳ 
	Where ����ID=����ID_IN And ��ҳID=��ҳID_IN;
    Begin 
        Select Nvl(��Ϣֵ,'NoneData') Into v_����ҽʦ From ������ҳ�ӱ� Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��Ϣ��='����ҽʦ';
    Exception
        When Others Then v_����ҽʦ:='NoneData';
    End;
    Begin
        Select Nvl(��Ϣֵ,'NoneData') Into v_����ҽʦ From ������ҳ�ӱ� Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��Ϣ��='����ҽʦ';
    Exception
        When Others Then v_����ҽʦ:='NoneData';
    End;

    Update ������ҳ
    Set ����=����_IN,�ѱ�=�ѱ�_IN,
        ����״��=����״��_IN,ѧ��=ѧ��_IN,
        ��ǰ����=��ǰ����_IN,ְҵ=ְҵ_IN,
        ��λ��ַ=��λ��ַ_IN,��λ�绰=��λ�绰_IN,
        ��λ�ʱ�=��λ�ʱ�_IN,��ͥ��ַ=��ͥ��ַ_IN,
        ��ͥ�绰=��ͥ�绰_IN,�����ʱ�=�����ʱ�_IN,
        ��ϵ������=��ϵ������_IN,��ϵ�˹�ϵ=��ϵ�˹�ϵ_IN,
        ��ϵ�˵�ַ=��ϵ�˵�ַ_IN,��ϵ�˵绰=��ϵ�˵绰_IN,
        ���λ�ʿ=���λ�ʿ_IN,����ҽʦ=����ҽʦ_IN,
        סԺҽʦ=סԺҽʦ_IN,��ĿԱ���=����Ա���_IN,
        ��ĿԱ����=����Ա����_IN
    Where ����ID=����ID_IN And ��ҳID=��ҳID_IN;
    
    --���޸�סԺ�ѱ�,�������������۲���
    Update ������Ϣ
    Set ����=����_IN,�ѱ�=Decode(v_��������,1,�ѱ�_IN,�ѱ�),
        ����״��=����״��_IN,ѧ��=ѧ��_IN,
        ְҵ=ְҵ_IN,������λ=��λ��ַ_IN,
        ��ͬ��λID=Decode(��ͬ��λID_IN,0,��ͬ��λID,��ͬ��λID_IN),
        ��λ�绰=��λ�绰_IN,��λ�ʱ�=��λ�ʱ�_IN,
        ��ͥ��ַ=��ͥ��ַ_IN,��ͥ�绰=��ͥ�绰_IN,
        �����ʱ�=�����ʱ�_IN,��ϵ������=��ϵ������_IN,
        ��ϵ�˹�ϵ=��ϵ�˹�ϵ_IN,��ϵ�˵�ַ=��ϵ�˵�ַ_IN,
        ��ϵ�˵绰=��ϵ�˵绰_IN
    Where ����ID=����ID_IN;

    --�����䶯��¼
    if v_סԺҽʦ <> סԺҽʦ_IN Or v_���λ�ʿ <> ���λ�ʿ_IN Or v_����ҽʦ <> ����ҽʦ_IN Or v_����ҽʦ <> ����ҽʦ_IN Or v_��ǰ����<>��ǰ����_IN Then
        Select Sysdate Into v_CurDate From Dual;
        Open c_OldInfo;
        Fetch c_OldInfo Into r_OldInfo;
        if c_OldInfo%RowCount=0 Then
            Close c_OldInfo;
            v_Error:='δ���ָò��˵�ǰ��Ч�ı䶯��¼��';
            Raise Err_Custom;
        End IF;

        if v_סԺҽʦ <> סԺҽʦ_IN Then
            v_ԭ��:=7;    
            Update ���˱䶯��¼ 
                Set ��ֹʱ��=v_CurDate,��ֹԭ��=v_ԭ��,
                    ����Ա���=����Ա���_IN,����Ա����=����Ա����_IN
            Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��ֹʱ�� is Null;    
            While c_OldInfo%Found Loop                  --ע��:�и��Ӵ�λʱ�ж�����¼                                                                    
                Insert Into ���˱䶯��¼(
                    ����ID,��ҳID,��ʼʱ��,��ʼԭ��,���Ӵ�λ,����ID,����ID,����ȼ�ID,
                    ��λ�ȼ�ID,����,���λ�ʿ,����ҽʦ,����ҽʦ,����ҽʦ,����,����Ա���,����Ա����)
                Values(
                    ����ID_IN,��ҳID_IN,v_CurDate,v_ԭ��,r_OldInfo.���Ӵ�λ,r_OldInfo.����ID,
                    r_OldInfo.����ID,r_OldInfo.����ȼ�ID,r_OldInfo.��λ�ȼ�ID,r_OldInfo.����,
                    r_OldInfo.���λ�ʿ,סԺҽʦ_IN,r_OldInfo.����ҽʦ,r_OldInfo.����ҽʦ,r_OldInfo.����,����Ա���_IN,����Ա����_IN); 
                Fetch c_OldInfo Into r_OldInfo; 
            End Loop;
            Close c_OldInfo;
            Open c_OldInfo;    --���´�,�Ա�ȡ������Ϣ
            Fetch c_OldInfo Into r_OldInfo; 
        End IF;
    
       if v_���λ�ʿ <> ���λ�ʿ_IN Then
            v_ԭ��:=8;    
            Update ���˱䶯��¼ 
                Set ��ֹʱ��=v_CurDate,��ֹԭ��=v_ԭ��,
                    ����Ա���=����Ա���_IN,����Ա����=����Ա����_IN
            Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��ֹʱ�� is Null;             
            While c_OldInfo%Found Loop            
                Insert Into ���˱䶯��¼(
                    ����ID,��ҳID,��ʼʱ��,��ʼԭ��,���Ӵ�λ,����ID,����ID,����ȼ�ID,
                    ��λ�ȼ�ID,����,���λ�ʿ,����ҽʦ,����ҽʦ,����ҽʦ,����,����Ա���,����Ա����)
                Values(
                    ����ID_IN,��ҳID_IN,v_CurDate,v_ԭ��,r_OldInfo.���Ӵ�λ,r_OldInfo.����ID,
                    r_OldInfo.����ID,r_OldInfo.����ȼ�ID,r_OldInfo.��λ�ȼ�ID,r_OldInfo.����,
                    ���λ�ʿ_IN,r_OldInfo.����ҽʦ,r_OldInfo.����ҽʦ,r_OldInfo.����ҽʦ,r_OldInfo.����,����Ա���_IN,����Ա����_IN);                
                Fetch c_OldInfo Into r_OldInfo; 
            End Loop;
            Close c_OldInfo;
            Open c_OldInfo;
            Fetch c_OldInfo Into r_OldInfo; 
        End IF;

        if v_����ҽʦ <> ����ҽʦ_IN Then
            Update ������ҳ�ӱ� Set ��Ϣֵ=����ҽʦ_IN Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��Ϣ��='����ҽʦ';
            IF SQL%RowCount=0 Then
                Insert Into ������ҳ�ӱ�(����ID,��ҳID,��Ϣ��,��Ϣֵ) Values (����ID_IN,��ҳID_IN,'����ҽʦ',����ҽʦ_IN);
            End IF;

            v_ԭ��:=11;    
            Update ���˱䶯��¼ 
                Set ��ֹʱ��=v_CurDate,��ֹԭ��=v_ԭ��,
                    ����Ա���=����Ա���_IN,����Ա����=����Ա����_IN
            Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��ֹʱ�� is Null;    
            While c_OldInfo%Found Loop                                                                     
                Insert Into ���˱䶯��¼(
                    ����ID,��ҳID,��ʼʱ��,��ʼԭ��,���Ӵ�λ,����ID,����ID,����ȼ�ID,
                    ��λ�ȼ�ID,����,���λ�ʿ,����ҽʦ,����ҽʦ,����ҽʦ,����,����Ա���,����Ա����)
                Values(
                    ����ID_IN,��ҳID_IN,v_CurDate,v_ԭ��,r_OldInfo.���Ӵ�λ,r_OldInfo.����ID,
                    r_OldInfo.����ID,r_OldInfo.����ȼ�ID,r_OldInfo.��λ�ȼ�ID,r_OldInfo.����,
                    r_OldInfo.���λ�ʿ,r_OldInfo.����ҽʦ,����ҽʦ_IN,r_OldInfo.����ҽʦ,r_OldInfo.����,����Ա���_IN,����Ա����_IN);                
                Fetch c_OldInfo Into r_OldInfo; 
            End Loop;
            Close c_OldInfo;
            Open c_OldInfo; 
            Fetch c_OldInfo Into r_OldInfo; 
        End IF;

        if v_����ҽʦ <> ����ҽʦ_IN Then
            Update ������ҳ�ӱ� Set ��Ϣֵ=����ҽʦ_IN Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��Ϣ��='����ҽʦ';
            IF SQL%RowCount=0 Then
                Insert Into ������ҳ�ӱ�(����ID,��ҳID,��Ϣ��,��Ϣֵ) Values (����ID_IN,��ҳID_IN,'����ҽʦ',����ҽʦ_IN);
            End IF;  

            v_ԭ��:=12;    
            Update ���˱䶯��¼ 
                Set ��ֹʱ��=v_CurDate,��ֹԭ��=v_ԭ��,
                    ����Ա���=����Ա���_IN,����Ա����=����Ա����_IN
            Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��ֹʱ�� is Null;    
            While c_OldInfo%Found Loop
                Insert Into ���˱䶯��¼(
                    ����ID,��ҳID,��ʼʱ��,��ʼԭ��,���Ӵ�λ,����ID,����ID,����ȼ�ID,
                    ��λ�ȼ�ID,����,���λ�ʿ,����ҽʦ,����ҽʦ,����ҽʦ,����,����Ա���,����Ա����)
                Values(
                    ����ID_IN,��ҳID_IN,v_CurDate,v_ԭ��,r_OldInfo.���Ӵ�λ,r_OldInfo.����ID,
                    r_OldInfo.����ID,r_OldInfo.����ȼ�ID,r_OldInfo.��λ�ȼ�ID,r_OldInfo.����,
                    r_OldInfo.���λ�ʿ,r_OldInfo.����ҽʦ,r_OldInfo.����ҽʦ,����ҽʦ_IN,r_OldInfo.����,����Ա���_IN,����Ա����_IN);                
                Fetch c_OldInfo Into r_OldInfo; 
            End Loop;
            Close c_OldInfo;
            Open c_OldInfo;
            Fetch c_OldInfo Into r_OldInfo; 
        End IF;

        if v_��ǰ����<>��ǰ����_IN Then
            v_ԭ��:=13;    
            Update ���˱䶯��¼ 
                Set ��ֹʱ��=v_CurDate,��ֹԭ��=v_ԭ��,
                    ����Ա���=����Ա���_IN,����Ա����=����Ա����_IN
            Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And ��ֹʱ�� is Null;    
            While c_OldInfo%Found Loop                                                          
                Insert Into ���˱䶯��¼(
                    ����ID,��ҳID,��ʼʱ��,��ʼԭ��,���Ӵ�λ,����ID,����ID,����ȼ�ID,
                    ��λ�ȼ�ID,����,���λ�ʿ,����ҽʦ,����ҽʦ,����ҽʦ,����,����Ա���,����Ա����)
                Values(
                    ����ID_IN,��ҳID_IN,v_CurDate,v_ԭ��,r_OldInfo.���Ӵ�λ,r_OldInfo.����ID,
                    r_OldInfo.����ID,r_OldInfo.����ȼ�ID,r_OldInfo.��λ�ȼ�ID,r_OldInfo.����,
                    r_OldInfo.���λ�ʿ,r_OldInfo.����ҽʦ,r_OldInfo.����ҽʦ,r_OldInfo.����ҽʦ,��ǰ����_IN,����Ա���_IN,����Ա����_IN);                
                Fetch c_OldInfo Into r_OldInfo; 
            End Loop;
            Close c_OldInfo;
            Open c_OldInfo;
            Fetch c_OldInfo Into r_OldInfo; 
        End IF; 
        Close c_OldInfo;
    End IF;

	--������Ժ���
    IF ��Ժ���_IN is Null AND ����ID_IN IS NULL Then
        Delete From ������ϼ�¼ Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And �������=2 AND ��¼��Դ=2;
    Else
        Update ������ϼ�¼ Set ����ID=����ID_IN,�������=��Ժ���_IN,��¼����=sysdate,��¼��=����Ա����_IN
        Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And �������=2 AND ��¼��Դ=2;
        IF SQL%RowCount=0 Then
            Insert Into ������ϼ�¼(
                ID,����ID,��ҳID,��¼��Դ,�������,��ϴ���,����ID,�������,��¼����,��¼��) 
            Values(
                ������ϼ�¼_ID.Nextval,����ID_IN,��ҳID_IN,2,2,1,����ID_IN,��Ժ���_IN,sysdate,����Ա����_IN);
        End IF;
    End IF;

	--������ҽ��Ժ���
    IF ��ҽ���_IN is Null AND ��ҽ����ID_IN IS NULL Then
        Delete From ������ϼ�¼ Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And �������=12 AND ��¼��Դ=2;
    Else
        Update ������ϼ�¼ Set ����ID=��ҽ����ID_IN,�������=��ҽ���_IN,��¼����=sysdate,��¼��=����Ա����_IN
        Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And �������=12 AND ��¼��Դ=2;
        IF SQL%RowCount=0 Then
            Insert Into ������ϼ�¼(
                ID,����ID,��ҳID,��¼��Դ,�������,��ϴ���,����ID,�������,��¼����,��¼��) 
            Values(
                ������ϼ�¼_ID.Nextval,����ID_IN,��ҳID_IN,2,12,1,��ҽ����ID_IN,��ҽ���_IN,sysdate,����Ա����_IN);
        End IF;
    End IF;

    Select Count(*) Into v_Count
        From ���˱䶯��¼
        Where ����ID=����ID_IN And ��ҳID=��ҳID_IN And Nvl(���Ӵ�λ,0)=0
            And ��ʼʱ�� IS Not Null And ��ֹʱ�� is Null;
    If v_Count>1 Then
        v_Error:='���ֲ��˴��ڷǷ��ı䶯��¼,��ǰ�������ܼ�����'||Chr(13)||Chr(10) ||'��������������粢�����������,��ˢ�²���״̬�����ԣ�';
        Raise Err_Custom;
    End IF;

Exception
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_סԺ������ҳ_UpDate;
/

CREATE OR REPLACE PROCEDURE ZL_���˽����¼_UPDATE(
	����ID_IN			����Ԥ����¼.����ID%TYPE,
	���ս���_IN			VARCHAR2,--"���㷽ʽ|������||....."
	����_IN				NUMBER:=0,
	ȱʡ���㷽ʽ_IN VARCHAR2:=NULL
) AS
	--���α�ΪҪɾ�����ɷ��ü�¼�����Ľ����¼
	CURSOR C_DEL IS
		SELECT A.*,B.���� FROM ����Ԥ����¼ A,���㷽ʽ B 
		WHERE A.���㷽ʽ=B.���� AND A.����ID=����ID_IN;

	--�����Ϣ
	V_NO			����Ԥ����¼.NO%TYPE;
	V_����ID		���˷��ü�¼.����ID%TYPE;
	V_��ҳID		���˷��ü�¼.��ҳID%TYPE;
	V_�Ǽ�ʱ��		���˷��ü�¼.�Ǽ�ʱ��%TYPE;
	V_����Ա���	���˷��ü�¼.����Ա���%TYPE;
	V_����Ա����	���˷��ü�¼.����Ա����%TYPE;
	
	--���ν������
	V_���ϼ�	    ����Ԥ����¼.��Ԥ��%TYPE;
	V_��Ԥ����	    ����Ԥ����¼.��Ԥ��%TYPE;
    
	--���ս���
	V_���ս���	VARCHAR2(255);
	V_��ǰ����	VARCHAR2(50);
	V_���㷽ʽ	����Ԥ����¼.���㷽ʽ%TYPE;
	V_������	����Ԥ����¼.��Ԥ��%TYPE;

	v_��¼����	����Ԥ����¼.��¼����%Type;
	v_ȱʡ ����Ԥ����¼.���㷽ʽ%TYPE;
        
    --�ֱҴ���������
    v_����          �����ʻ�.����%TYPE;    
    v_CentMode      Number;
    v_�ֽ���	    ����Ԥ����¼.��Ԥ��%TYPE;
    v_CashCented    ����Ԥ����¼.��Ԥ��%TYPE;
    v_�����      ����Ԥ����¼.��Ԥ��%TYPE;
    v_����ID		���˷��ü�¼.ID%Type;
    v_���			���˷��ü�¼.���%Type;
    v_�շ����		���˷��ü�¼.�շ����%Type;
    v_�շ�ϸĿID	���˷��ü�¼.�շ�ϸĿID%Type;    
    v_������ĿID	���˷��ü�¼.������ĿID%Type;
    v_�վݷ�Ŀ		���˷��ü�¼.�վݷ�Ŀ%Type;
    
	--��ʱ����    
	ERR_CUSTOM	EXCEPTION;
	V_ERROR		VARCHAR2(255);
BEGIN
	--���ȱʡ���㷽ʽΪ�գ���ȡ�ֽ���㷽ʽ
	IF ȱʡ���㷽ʽ_IN IS NULL THEN 
		BEGIN 
			SELECT ���� INTO v_ȱʡ FROM ���㷽ʽ WHERE ����=1 AND ROWNUM<2;
		EXCEPTION 
			WHEN OTHERS THEN v_ȱʡ:='�ֽ�';
		END ;
	ELSE 
		v_ȱʡ:=ȱʡ���㷽ʽ_IN;
	END IF ;

	--ȡ�ñ��ν���������Ϣ
	IF NVL(����_IN,0)=1 THEN
		SELECT NO,����ID,�շ�ʱ��,����Ա���,����Ա����
			INTO V_NO,V_����ID,V_�Ǽ�ʱ��,V_����Ա���,V_����Ա����
		FROM ���˽��ʼ�¼ WHERE ID=����ID_IN;
	ELSE
		SELECT NO,����ID,�Ǽ�ʱ��,����Ա���,����Ա����
			INTO V_NO,V_����ID,V_�Ǽ�ʱ��,V_����Ա���,V_����Ա����
		FROM ���˷��ü�¼ WHERE ����ID=����ID_IN AND ROWNUM=1;

		Begin --20071027 �¶�
			Select ��¼���� Into v_��¼����
			From ����Ԥ����¼ Where ����ID=����ID_IN And Rownum=1;
		Exception --20071027 �¶�
			WHEN OTHERS Then v_��¼����:=-1; --20071027 �¶�
		End; --20071027 �¶�

	END IF;
	IF NVL(V_����ID,0)<>0 THEN
		SELECT סԺ���� INTO V_��ҳID FROM ������Ϣ WHERE ����ID=V_����ID;
	END IF;
	
	--���˽ɿ�,Ԥ������,��Ϊû�иĳ�Ԥ����
	V_���ϼ�:=0;V_��Ԥ����:=0;
	FOR R_DEL IN C_DEL LOOP
		IF r_Del.��¼���� Not IN(1,11) THEN 
			UPDATE ��Ա�ɿ����
				SET ���=NVL(���,0)-R_DEL.��Ԥ��
			 WHERE �տ�Ա=V_����Ա���� AND ����=1
				AND ���㷽ʽ=R_DEL.���㷽ʽ;
			IF SQL%ROWCOUNT=0 THEN
				INSERT INTO ��Ա�ɿ����(
					�տ�Ա,���㷽ʽ,����,���)
				VALUES(
					V_����Ա����,R_DEL.���㷽ʽ,1,-1*R_DEL.��Ԥ��);
			END IF;
			
			V_���ϼ�:=V_���ϼ�+R_DEL.��Ԥ��;

			DELETE FROM ����Ԥ����¼ WHERE ID=R_DEL.ID;
		END IF ;
	END LOOP;
	
	--------------------------------------------------------------------------------------------------------------
	--------------------------------------------------------------------------------------------------------------
	--����ҽ��֧������
	IF ���ս���_IN IS NOT NULL THEN 
		--�������ս���	
		V_���ս���:=���ս���_IN||'||';
		WHILE V_���ս��� IS NOT NULL LOOP
			V_��ǰ����:=SUBSTR(V_���ս���,1,INSTR(V_���ս���,'||')-1);

			V_���㷽ʽ:=SUBSTR(V_��ǰ����,1,INSTR(V_��ǰ����,'|')-1);
			V_������:=TO_NUMBER(SUBSTR(V_��ǰ����,INSTR(V_��ǰ����,'|')+1));

			INSERT INTO ����Ԥ����¼(
				ID,��¼����,NO,��¼״̬,����ID,��ҳID,ժҪ,���㷽ʽ,�տ�ʱ��,����Ա���,����Ա����,��Ԥ��,����ID)
			VALUES(
				����Ԥ����¼_ID.NEXTVAL,DECODE(����_IN,1,2,v_��¼����),V_NO,1,V_����ID,V_��ҳID,'���ղ���',
				V_���㷽ʽ,V_�Ǽ�ʱ��,V_����Ա���,V_����Ա����,V_������,����ID_IN);
			
			V_���ϼ�:=V_���ϼ�-V_������;

			V_���ս���:=SUBSTR(V_���ս���,INSTR(V_���ս���,'||')+2);
		END LOOP;
	END IF;

	--ʣ�ಿ��ȫ����ȱʡ���㷽ʽ���㣬(С����Ҳ�����ж��⴦��)
	IF V_���ϼ�<>0 Then             
		UPDATE ����Ԥ����¼
			SET ��Ԥ��=��Ԥ��+V_���ϼ�
		WHERE ����ID=����ID_IN AND ���㷽ʽ=v_ȱʡ AND ��¼����=DECODE(����_IN,1,2,v_��¼����);
		IF SQL%ROWCOUNT=0 THEN 
			INSERT INTO ����Ԥ����¼(
				ID,��¼����,NO,��¼״̬,����ID,��ҳID,ժҪ,���㷽ʽ,�տ�ʱ��,����Ա���,����Ա����,��Ԥ��,����ID)
			VALUES(
				����Ԥ����¼_ID.NEXTVAL,DECODE(����_IN,1,2,v_��¼����),V_NO,1,V_����ID,V_��ҳID,'�ֽ𲿷�',v_ȱʡ,
				V_�Ǽ�ʱ��,V_����Ա���,V_����Ա����,V_���ϼ�,����ID_IN);
		END IF ;
        
        --�ҺŽ���,�ֱҴ���(���ڹҺŽ���û��Ԥ����,�����ڴ˹����и��ݷֱҴ������������)
        If v_��¼����=4 Then
           Begin
               Select ���� Into v_���� From �����ʻ� Where ����Id=V_����Id;                              
           Exception
               When Others Then v_����:=0;
           End;
           If v_����=413 Then --�Ϻ�ҽ��,�Һ�֧�ֱַҴ���
               Begin
                   Select A.��Ԥ�� Into V_�ֽ��� From ����Ԥ����¼ A,���㷽ʽ B
                   Where A.���㷽ʽ=B.���� And B.����=1 And A.����Id=����ID_IN;
               Exception
                   When Others Then V_�ֽ���:=0;
               End;           
               If FLOOR(V_�ֽ���*10)<>V_�ֽ���*10 Then    
                   --v_CentMode:=Nvl(zl_GetSysParameter(14),0);              
                   --If v_CentMode=1 Then                                                                      
                   --   v_CashCented:=round(V_�ֽ���,1);                      --1.�������뷨,eg:0.51=0.50;0.56=0.60  
                   --Elsif  v_CentMode=2 Then
                   --   v_CashCented:=CEIL(V_�ֽ���*10)/10;                   --2.�����շ�,eg:0.51=0.60,0.56=0.60
                   --Elsif  v_CentMode=3 Then
                      v_CashCented:=FLOOR(V_�ֽ���*10)/10;                  --3.����շ�,eg:0.51=0.50,0.56=0.50
                   --Else
                   --   v_CashCented:=V_�ֽ���;
                   --End If;        
                   v_�����:=v_CashCented-V_�ֽ���;
                   If v_�����<>0 Then                 
                      --1.����Ԥ����¼(һ�����ڼ�¼)
                      Update ����Ԥ����¼ Set ��Ԥ��=v_CashCented
                      Where ���㷽ʽ=(Select ���� From ���㷽ʽ Where ����=1 And Rownum=1) And ����Id=����ID_IN;
                      
                      --2.���������ü�¼(ע:���㵥λ��¼���Ǻű�,���Բ�ȡ������)
                      Begin
                          Select A.���,A.ID,C.ID,C.�վݷ�Ŀ 
                          Into v_�շ����,v_�շ�ϸĿID,v_������ĿID,v_�վݷ�Ŀ
                          From �շ���ĿĿ¼ A,�շѼ�Ŀ B,������Ŀ C,�շ��ض���Ŀ D
                          Where D.�ض���Ŀ='�����' And D.�շ�ϸĿID=A.Id  And A.ID=B.�շ�ϸĿID And B.������ĿID=C.ID
                              And Sysdate Between B.ִ������ And Nvl(B.��ֹ����,To_Date('3000-01-01','YYYY-MM-DD')) ;
                      Exception
                          When Others Then                        
                          v_Error:='������ȷ��ȡ�շ���������Ϣ�����ȼ�����Ŀ�Ƿ�������ȷ��';
                          Raise Err_Custom;
                      End; 
                      Select ���˷��ü�¼_ID.Nextval Into v_����ID From Dual;
                      Select Max(���)+1 Into v_��� From ���˷��ü�¼ Where ����ID=����ID_IN;
                                           
                      Insert Into ���˷��ü�¼(
                          ID,��¼����,NO,ʵ��Ʊ��,��¼״̬,���,��������,�۸񸸺�,�����־,����ID,��ʶ��,����,����,�Ա�,
                          ����,���˲���ID,���˿���ID,�ѱ�,�շ����,�շ�ϸĿID,���㵥λ,��ҩ����,����,����,�Ӱ��־,���ӱ�־,
                          ������ĿID,�վݷ�Ŀ,��׼����,Ӧ�ս��,ʵ�ս��,���ʷ���,������,��������ID,������,����ʱ��,�Ǽ�ʱ��,
                          ִ�в���ID,ִ����,ִ��״̬,����ID,���ʽ��,����Ա���,����Ա����,�Ƿ��ϴ�)
                      Select
                          v_����ID,��¼����,NO,ʵ��Ʊ��,��¼״̬,v_���,NULL,NULL,�����־,����ID,��ʶ��,����,����,�Ա�,����,
                          ���˲���ID,���˿���ID,�ѱ�,v_�շ����,v_�շ�ϸĿID,���㵥λ,��ҩ����,1,1,�Ӱ��־,9,
                          v_������ĿID,v_�վݷ�Ŀ,v_�����,v_�����,v_�����,���ʷ���,������,��������ID,������,����ʱ��,�Ǽ�ʱ��,
                          ִ�в���ID,ִ����,ִ��״̬,����ID_IN,v_�����,����Ա���,����Ա����,1
                      From ���˷��ü�¼
                      Where ����ID=����ID_IN And Rownum=1; 
                                         
                      --3.����"���˷��û���"  
                      --ֻ���ܲ��������ı仯.��Ϊ�˱�������������α�
                      For C_Error In (
                          Select TRUNC(�Ǽ�ʱ��) as ����,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,�����־,Ӧ�ս��,ʵ�ս��,���ʽ��
                          From ���˷��ü�¼
                          Where ����Id=����Id_IN And ���ӱ�־=9
                      ) Loop
                          Update ���˷��û���
                              Set Ӧ�ս��=Nvl(Ӧ�ս��,0)+C_Error.Ӧ�ս��,ʵ�ս��=Nvl(ʵ�ս��,0)+C_Error.ʵ�ս��,���ʽ��=Nvl(���ʽ��,0)+C_Error.���ʽ��
                          Where ����=C_Error.����
                              And Nvl(���˲���ID,0)=Nvl(C_Error.���˲���ID,0) And Nvl(���˿���ID,0)=Nvl(C_Error.���˿���ID,0)
                              And Nvl(��������ID,0)=Nvl(C_Error.��������ID,0) And Nvl(ִ�в���ID,0)=Nvl(C_Error.ִ�в���ID,0)
                              And ������ĿID+0=C_Error.������ĿId And ��Դ;��=C_Error.�����־ And ���ʷ���=0; 
                          If SQL%RowCount=0 Then
                              Insert Into ���˷��û���(
                                  ����,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,��Դ;��,���ʷ���,Ӧ�ս��,ʵ�ս��,���ʽ��)
                              Values(
                                  C_Error.����,C_Error.���˲���ID,C_Error.���˿���ID,C_Error.��������ID,C_Error.ִ�в���ID,
                                  C_Error.������ĿID,C_Error.�����־,0,C_Error.Ӧ�ս��,C_Error.ʵ�ս��,C_Error.���ʽ��);
                          End If;
                      End Loop;                   
                   End If;
               End If;
           End If;
        End If;        
	END IF;
	
	--����ٴ���"��Ա�ɿ����"(û�ж���Ԥ���ǲ���,����"�������"��Ԥ�����ø���)
	FOR R_DEL IN C_DEL LOOP
		IF r_Del.��¼���� Not IN(1,11) THEN 
			UPDATE ��Ա�ɿ����
				SET ���=NVL(���,0)+R_DEL.��Ԥ��
			 WHERE �տ�Ա=V_����Ա���� AND ����=1
				AND ���㷽ʽ=R_DEL.���㷽ʽ;
			IF SQL%ROWCOUNT=0 THEN
				INSERT INTO ��Ա�ɿ����(
					�տ�Ա,���㷽ʽ,����,���)
				VALUES(
					V_����Ա����,R_DEL.���㷽ʽ,1,R_DEL.��Ԥ��);
			END IF;
		END IF ;
	END LOOP;
	DELETE FROM ��Ա�ɿ���� WHERE ����=1 AND �տ�Ա=V_����Ա���� AND NVL(���,0)=0;
EXCEPTION
	WHEN ERR_CUSTOM THEN RAISE_APPLICATION_ERROR(-20101,'[ZLSOFT]'||V_ERROR||'[ZLSOFT]');
	WHEN OTHERS THEN ZL_ERRORCENTER(SQLCODE,SQLERRM);
END ZL_���˽����¼_UPDATE;
/

Create Or Replace Procedure zl_סԺ���ʼ�¼_Delete(
    NO_IN            ���˷��ü�¼.NO%Type,
    ���_IN            Varchar2,
    ����Ա���_IN    ���˷��ü�¼.����Ա���%Type,
    ����Ա����_IN    ���˷��ü�¼.����Ա����%Type,
    ��¼����_IN        ���˷��ü�¼.��¼����%Type:=2
)
AS
     --���ܣ�����һ��סԺ���ʵ�����ָ�������
     --��ţ���ʽ��"1,3,5,7,8",Ϊ�ձ�ʾ�������пɳ�����
     --��¼����:    2-�˹����ʵ�,3-�Զ����ʵ�
    --�ù����������ָ��������

    --���α�ΪҪ�˷ѵ��ݵ�����ԭʼ��¼
    Cursor c_Bill is
        Select * From ���˷��ü�¼
        Where NO=NO_IN And ��¼����=��¼����_IN And ��¼״̬ IN(0,1,3) And �����־=2
        Order by �շ�ϸĿID,���;

    --���α����ڴ���ҩƷ����������
    --��Ҫ�ܷ��õ�ִ��״̬,��Ϊ���ڴ˲�����
    Cursor c_Stock is
        Select * From ҩƷ�շ���¼
        Where NO=NO_IN And ���� IN(9,10,25,26) And Mod(��¼״̬,3)=1 And ����� IS NULL
            And ����ID IN(
                Select ID From ���˷��ü�¼ 
                Where NO=NO_IN And ��¼����=��¼����_IN And ��¼״̬ IN(0,1,3) 
                    And �շ���� IN('4','5','6','7') And �����־=2
                    And (INSTR(','||���_IN||',',','||���||',')>0 Or ���_IN Is Null)
                )
        Order BY ҩƷID;
	r_Stock c_Stock%RowType;
    
    --���α����ڴ���δ��ҩƷ��¼
    Cursor c_Spare is
        Select * From δ��ҩƷ��¼ Where NO=NO_IN And ���� IN(9,10,25,26);

    --���α����ڴ�����ü�¼���
    Cursor c_Serial is
        Select ���,�۸񸸺� From ���˷��ü�¼ Where NO=NO_IN And ��¼����=��¼����_IN And ��¼״̬ IN(0,1,3) Order BY ���;

	v_ҽ��ID		����ҽ����¼.ID%Type;
	v_����			Number;
	v_����			���˷��ü�¼.�۸񸸺�%Type;
	
    --�����˷Ѽ������
    v_ʣ������		Number;
    v_ʣ��Ӧ��		Number;
    v_ʣ��ʵ��		Number;
    v_ʣ��ͳ��		Number;

    v_׼������		Number;
    v_�˷Ѵ���		Number;

    v_Ӧ�ս��		Number;
    v_ʵ�ս��		Number;
    v_ͳ����		Number;

    v_Dec			Number;    
    v_Count			Number;
    v_CurDate		Date;
    Err_Custom		Exception;
    v_Error			Varchar2(255);
Begin
    --�Ƿ��Ѿ�ȫ����ȫִ��(ֻ�����ŵ��ݵļ��)
    Select Nvl(Count(*),0) Into v_Count 
    From ���˷��ü�¼ 
    Where NO=NO_IN And ��¼����=��¼����_IN And ��¼״̬ IN(0,1,3) And Nvl(ִ��״̬,0)<>1 And �����־=2;
    IF v_Count = 0 Then
        v_Error := '�õ����е���Ŀ�Ѿ�ȫ����ȫִ�У�';
        Raise Err_Custom;
    End IF;

    --δ��ȫִ�е���Ŀ�Ƿ���ʣ������(ֻ�����ŵ��ݵļ��)
    Select Nvl(Count(*),0) Into v_Count
    From (
        Select ���,Sum(����) as ʣ������
        From (
            Select ��¼״̬,Nvl(�۸񸸺�,���) as ���,
                Avg(Nvl(����,1)*����) as ���� 
            From ���˷��ü�¼
            Where NO=NO_IN And ��¼����=��¼����_IN And �����־=2
                And Nvl(�۸񸸺�,���) IN (
                        Select Nvl(�۸񸸺�,���) 
                        From ���˷��ü�¼ 
                        Where NO=NO_IN And ��¼����=��¼����_IN 
                            And ��¼״̬ IN(0,1,3) And Nvl(ִ��״̬,0)<>1)
            Group by ��¼״̬,Nvl(�۸񸸺�,���)
            )
        Group by ��� Having Sum(����)<>0);
    IF v_Count = 0 Then
        v_Error := '�õ�����δ��ȫִ�в�����Ŀʣ������Ϊ��,û�п������ʵķ��ã�';
        Raise Err_Custom;
    End IF;
    
    ---------------------------------------------------------------------------------
	--�ȴ�ҩƷ��Ӧ���ݼ�,��ȷ����ǰ������������,Ϊ�˴������ж�
	--�������α�������ȡ��"����� is Null"��������Ϊ�����ҩ���ܲ������ѷ�
	Open c_Stock;

    --���ñ���
    Select Sysdate Into v_CurDate From Dual;
    --���С��λ��
    Begin
        Select To_Number(����ֵ) Into v_Dec From ϵͳ������ Where ������=9;
    Exception
        When Others Then v_Dec:=2;
    End;
    
    --ѭ������ÿ�з���(������Ŀ��)
	For r_Bill IN c_Bill Loop
		IF INSTR(','||���_IN||',',','||Nvl(r_Bill.�۸񸸺�,r_Bill.���)||',') >0 Or ���_IN Is Null Then
			Select Decode(��¼״̬,0,1,0) Into v_���� From ���˷��ü�¼ Where ID=r_Bill.ID;
			If v_����=0 Then
				IF Nvl(r_Bill.ִ��״̬,0)<>1 Then
					--��ʣ������,ʣ��Ӧ��,ʣ��ʵ��
					Select 
						Sum(Nvl(����,1)*����),Sum(Ӧ�ս��),Sum(ʵ�ս��),Sum(ͳ����)
						Into v_ʣ������,v_ʣ��Ӧ��,v_ʣ��ʵ��,v_ʣ��ͳ��
					From ���˷��ü�¼ 
					Where NO=NO_IN And ��¼����=��¼����_IN And ���=r_Bill.���;

					IF v_ʣ������=0 Then
						IF ���_IN IS Not NULL Then 
							v_Error := '�����е�'||Nvl(r_Bill.�۸񸸺�,r_Bill.���)||'�з����Ѿ�ȫ�����ʣ�';
							Raise Err_Custom;
						End IF;
						--�����δ�޶��к�,ԭʼ�����еĸñ��Ѿ�ȫ������(ִ��״̬=0��һ�ֿ���)
					Else
						--׼������(��ҩƷ��ĿΪʣ������,ԭʼ����)
						IF Instr(',4,5,6,7,',r_Bill.�շ����)=0 Then
							v_׼������:=v_ʣ������;
						Else
							Select Sum(Nvl(����,1)*ʵ������) Into v_׼������
							From ҩƷ�շ���¼
							Where NO=NO_IN And ���� IN(9,10,25,26) And MOD(��¼״̬,3)=1 
								And ����� is NULL And ����ID=r_Bill.ID;

							--���������õ���������
							If r_Bill.�շ����='4' And Nvl(v_׼������,0)=0 Then
								v_׼������:=v_ʣ������;
							End IF;
						End if;

						--�����˷��ü�¼
						
						--�ñ���Ŀ�ڼ�������
						Select Nvl(Max(Abs(ִ��״̬)),0)+1 Into v_�˷Ѵ���
						From ���˷��ü�¼ 
						Where NO=NO_IN And ��¼����=��¼����_IN And ��¼״̬=2 And ���=r_Bill.��� And �����־=2;
						
						--���=ʣ����*(׼����/ʣ����)
						v_Ӧ�ս��:=Round(v_ʣ��Ӧ��*(v_׼������/v_ʣ������),v_Dec);
						v_ʵ�ս��:=Round(v_ʣ��ʵ��*(v_׼������/v_ʣ������),v_Dec);
						v_ͳ����:=Round(v_ʣ��ͳ��*(v_׼������/v_ʣ������),v_Dec);

						--�����˷Ѽ�¼
						Insert Into ���˷��ü�¼(
							ID,NO,��¼����,��¼״̬,���,��������,�۸񸸺�,��ҳID,����ID,ҽ�����,�����־,�ಡ�˵�,Ӥ����,����,
							�Ա�,����,��ʶ��,����,�ѱ�,���˲���ID,���˿���ID,�շ����,�շ�ϸĿID,���㵥λ,����,��ҩ����,
							����,�Ӱ��־,���ӱ�־,������ĿID,�վݷ�Ŀ,���ʷ���,��׼����,Ӧ�ս��,ʵ�ս��,��������ID,
							������,ִ�в���ID,������,ִ����,ִ��״̬,ִ��ʱ��,����Ա���,����Ա����,����ʱ��,�Ǽ�ʱ��,
							������Ŀ��,���մ���ID,ͳ����,���ձ���,���ʵ�ID,ժҪ)
						Select ���˷��ü�¼_ID.Nextval,NO,��¼����,2,���,��������,�۸񸸺�,��ҳID,����ID,ҽ�����,�����־,�ಡ�˵�,
							Ӥ����,����,�Ա�,����,��ʶ��,����,�ѱ�,���˲���ID,���˿���ID,�շ����,�շ�ϸĿID,���㵥λ,
							Decode(Sign(v_׼������-Nvl(����,1)*����),0,����,1),��ҩ����,
							Decode(Sign(v_׼������-Nvl(����,1)*����),0,-1*����,-1*v_׼������),�Ӱ��־,���ӱ�־,
							������ĿID,�վݷ�Ŀ,���ʷ���,��׼����,-1*v_Ӧ�ս��,-1*v_ʵ�ս��,��������ID,������,ִ�в���ID,
							������,ִ����,-1*v_�˷Ѵ���,ִ��ʱ��,����Ա���_IN,����Ա����_IN,����ʱ��,v_CurDate,
							������Ŀ��,���մ���ID,-1*v_ͳ����,���ձ���,���ʵ�ID,ժҪ
						From ���˷��ü�¼ Where ID=r_Bill.ID;
						
						--��¼����ҽ�����Ѷ�Ӧ��ҽ��ID(����������)
						If v_ҽ��ID IS Null And r_Bill.ҽ����� IS Not Null Then
							v_ҽ��ID:=r_Bill.ҽ�����;
						End IF;

						--�������
						Update �������
							Set �������=Nvl(�������,0) - v_ʵ�ս��
						 Where ����ID=r_Bill.����ID And ����=1;
						IF SQL%RowCount=0 Then
							Insert Into �������(
								����ID,����,�������,Ԥ�����)
							Values(
								r_Bill.����ID,1,-1*v_ʵ�ս��,0);
						End IF;
						
						--����δ�����
						Update ����δ�����
							Set ���=Nvl(���,0) - v_ʵ�ս��
						 Where ����ID=r_Bill.����ID
							And Nvl(��ҳID,0)=Nvl(r_Bill.��ҳID,0)
							And Nvl(���˲���ID,0)=Nvl(r_Bill.���˲���ID,0)
							And Nvl(���˿���ID,0)=Nvl(r_Bill.���˿���ID,0)
							And Nvl(��������ID,0)=Nvl(r_Bill.��������ID,0)
							And Nvl(ִ�в���ID,0)=Nvl(r_Bill.ִ�в���ID,0)
							And ������ĿID+0=r_Bill.������ĿID And ��Դ;��+0=2;
						IF SQL%RowCount=0 Then
							Insert Into ����δ�����(
								����ID,��ҳID,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,��Դ;��,���)
							Values(
								r_Bill.����ID,r_Bill.��ҳID,r_Bill.���˲���ID,r_Bill.���˿���ID,
								r_Bill.��������ID,r_Bill.ִ�в���ID,r_Bill.������ĿID,2,-1*v_ʵ�ս��);
						End IF;

						--�����˷��û���
						Update ���˷��û���
							Set Ӧ�ս��=Nvl(Ӧ�ս��,0) - v_Ӧ�ս��,
								ʵ�ս��=Nvl(ʵ�ս��,0) - v_ʵ�ս��
						 Where ����=Trunc(v_CurDate)
							And Nvl(���˲���ID,0)=Nvl(r_Bill.���˲���ID,0)
							And Nvl(���˿���ID,0)=Nvl(r_Bill.���˿���ID,0)
							And Nvl(��������ID,0)=Nvl(r_Bill.��������ID,0)
							And Nvl(ִ�в���ID,0)=Nvl(r_Bill.ִ�в���ID,0)
							And ������ĿID+0=r_Bill.������ĿID
							And ��Դ;��=2 And ���ʷ���=1;
						IF SQL%RowCount=0 Then
							Insert Into ���˷��û���(
								����,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,��Դ;��,���ʷ���,Ӧ�ս��,ʵ�ս��,���ʽ��)
							Values(
								Trunc(v_CurDate),r_Bill.���˲���ID,r_Bill.���˿���ID,r_Bill.��������ID,r_Bill.ִ�в���ID,
								r_Bill.������ĿID,2,1,-1 * v_Ӧ�ս��,-1 * v_ʵ�ս��,0);
						End IF;
						
						--���ԭ���ü�¼
						--ִ��״̬:ȫ������(׼����=ʣ����)���Ϊ0,������Ϊ1
						Update ���˷��ü�¼ 
							Set ��¼״̬=3,
								ִ��״̬=Decode(Sign(v_׼������-v_ʣ������),0,0,1) 
						Where ID=r_Bill.ID;
					End IF;
				Else
					IF ���_IN Is Not Null Then
						v_Error := '�����е�'||Nvl(r_Bill.�۸񸸺�,r_Bill.���)||'�з����Ѿ���ȫִ��,�������ʣ�';
						Raise Err_Custom;
					End IF;
					--���:û�޶��к�,ԭʼ�����а����Ѿ���ȫִ�е�
				End IF;
			End IF;
		End IF;
	End Loop;
	
    ---------------------------------------------------------------------------------
    --ҩƷ�������
	Fetch c_Stock Into r_Stock;
    While c_Stock%Found Loop
        --����ҩƷ���
        If r_Stock.�ⷿID IS Not NULL Then
            Update ҩƷ���
                Set ��������=Nvl(��������,0)+Nvl(r_Stock.����,1)*Nvl(r_Stock.ʵ������,0)
             Where �ⷿID=r_Stock.�ⷿID And ҩƷID=r_Stock.ҩƷID
                And Nvl(����,0)=Nvl(r_Stock.����,0) And ����=1;
            IF SQL%RowCount=0 Then
                Insert Into ҩƷ���(
                    �ⷿID,ҩƷID,����,����,Ч��,��������,�ϴ�����,�ϴβ���,���Ч��)
                Values(
                    r_Stock.�ⷿID,r_Stock.ҩƷID,1,r_Stock.����,r_Stock.Ч��,
                    Nvl(r_Stock.����,1)*Nvl(r_Stock.ʵ������,0),r_Stock.����,r_Stock.����,r_Stock.���Ч��);
            End IF;
        End IF;

        --ɾ��ҩƷ�շ���¼(���ϲ����������:����� Is Null)
        Delete From ҩƷ�շ���¼ Where ID=r_Stock.ID And ����� Is Null;
		IF SQL%RowCount=0 Then
			If r_Stock.���� IN(9,10) Then
				v_Error:='Ҫ���ʵķ����д����ѷ�ҩ��ҩƷ�����ѱ����������ʣ�������ǲ�����������ġ�';
			Else
				v_Error:='Ҫ���ʵķ����д����ѷ��ϵ����ģ����ѱ����������ʣ�������ǲ�����������ġ�';
			End IF;
			Raise Err_Custom;
		End IF;

		Fetch c_Stock Into r_Stock;
    End Loop;
	Close c_Stock;

    --δ��ҩƷ��¼
    For r_Spare IN c_Spare Loop
        Select Nvl(Count(*),0) Into v_Count
        From ҩƷ�շ���¼ 
        Where NO=NO_IN And ����=r_Spare.���� And Mod(��¼״̬,3)=1 
            And ����� is NULL And Nvl(�ⷿID,0)=Nvl(r_Spare.�ⷿID,0);
        If v_Count=0 Then
            Delete From δ��ҩƷ��¼ Where ����=r_Spare.���� And NO=NO_IN And Nvl(�ⷿID,0)=Nvl(r_Spare.�ⷿID,0);
        End IF;
    End Loop;

	---------------------------------------------------------------------------------
	--����ǻ���,ֱ��ɾ�����ü�¼(ҩƷ�����)
	v_Count:=0;
	For r_Bill IN c_Bill Loop
		IF INSTR(','||���_IN||',',','||Nvl(r_Bill.�۸񸸺�,r_Bill.���)||',') >0 Or ���_IN Is Null Then
			Select Decode(��¼״̬,0,1,0) Into v_���� From ���˷��ü�¼ Where ID=r_Bill.ID;
			If v_����=1 Then
				IF Nvl(r_Bill.ִ��״̬,0)<>1 Then
					Delete From ���˷��ü�¼ Where ID=r_Bill.ID;
					v_Count:=v_Count+1;--��¼�Ƿ���ɾ����

					--��¼����ҽ�����Ѷ�Ӧ��ҽ��ID(����������)
					If v_ҽ��ID IS Null And r_Bill.ҽ����� IS Not Null Then
						v_ҽ��ID:=r_Bill.ҽ�����;
					End IF;
				Else
					IF ���_IN Is Not Null Then
						v_Error := '�����е�'||Nvl(r_Bill.�۸񸸺�,r_Bill.���)||'�з����Ѿ���ȫִ��,�������ʣ�';
						Raise Err_Custom;
					End IF;
					--���:û�޶��к�,ԭʼ�����а����Ѿ���ȫִ�е�
				End IF;
			End IF;
		End IF;
	End Loop;

	--ɾ��֮����ͳһ�������
	If v_Count>0 Then
		v_Count:=1;
		For r_Serial In c_Serial Loop
			If r_Serial.�۸񸸺� IS NULL Then 
				v_����:=v_Count;
			End IF;

			Update ���˷��ü�¼ 
				Set ���=v_Count,
					�۸񸸺�=Decode(�۸񸸺�,NULL,NULL,v_����)
			Where NO=NO_IN And ��¼����=��¼����_IN And ���=r_Serial.���;
			
			Update ���˷��ü�¼
				Set ��������=v_Count
			Where NO=NO_IN And ��¼����=��¼����_IN And ��������=r_Serial.���;

			v_Count:=v_Count+1;
		End Loop;
	End IF;

	--���ŵ���ȫ������ʱ��ɾ������ҽ������
	If ���_IN IS NULL And v_ҽ��ID IS Not NULL Then
		Select Nvl(Count(*),0) Into v_Count
		From (
			Select ���,Sum(����) as ʣ������
			From (
				Select ��¼״̬,Nvl(�۸񸸺�,���) as ���,
					Avg(Nvl(����,1)*����) as ���� 
				From ���˷��ü�¼
				Where NO=NO_IN And ��¼����=2 And ҽ�����+0=v_ҽ��ID
				Group by ��¼״̬,Nvl(�۸񸸺�,���)
				)
			Group by ��� Having Nvl(Sum(����),0)<>0);
		IF v_Count = 0 Then
			Delete From ����ҽ������ Where ҽ��ID=v_ҽ��ID And ��¼����=2 And NO=NO_IN;
		End IF;
	End IF;
Exception
    When Err_Custom Then Raise_Application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When Others Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_סԺ���ʼ�¼_Delete;
/

CREATE OR REPLACE Procedure zl_סԺһ�η���_Delete(
    ����ID_IN        ���˷��ü�¼.����ID%Type,
    ��ҳID_IN        ���˷��ü�¼.��ҳID%Type
)
AS
    --���ܣ�ɾ��סԺ���˼����һ���Է��á�
    Cursor c_Money is
        Select ���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,
            Nvl(Sum(Ӧ�ս��),0) AS Ӧ�ս��,Nvl(Sum(ʵ�ս��),0) AS ʵ�ս��
        From ���˷��ü�¼
        Where ��¼����=3 And ��¼״̬=1 And ���ӱ�־=8 And ����ID=����ID_IN And ��ҳID=��ҳID_IN
        Group BY ���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID;

    v_��Ա���        ���˷��ü�¼.����Ա���%Type;
    v_��Ա����        ���˷��ü�¼.����Ա����%Type;
    
    v_Date        Date;
    v_Temp        Varchar2(255);
    v_Error        Varchar2(255);
    Err_Custom    Exception;
Begin
    --ȡ����Ա��Ϣ(����ID,��������;��ԱID,��Ա���,��Ա����)
    v_Temp:=zl_Identity;
    v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
    v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
    v_��Ա���:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
    v_��Ա����:=Substr(v_Temp,Instr(v_Temp,',')+1);

    Select Sysdate Into v_Date From Dual;
    
    --�������ϼ�¼
    Insert Into ���˷��ü�¼(
        ID,��¼����,NO,��¼״̬,���,�۸񸸺�,����ID,��ҳID,�����־,���ʷ���,����,�Ա�,����,��ʶ��,
        ����,���˲���ID,���˿���ID,�ѱ�,�շ����,�շ�ϸĿID,���㵥λ,����,����,���ӱ�־,������ĿID,
        �վݷ�Ŀ,��׼����,Ӧ�ս��,ʵ�ս��,������,��������ID,������,����ʱ��,�Ǽ�ʱ��,ִ�в���ID,
        ����Ա���,����Ա����)
    Select
        ���˷��ü�¼_ID.Nextval,��¼����,NO,2,���,�۸񸸺�,����ID,��ҳID,�����־,���ʷ���,����,
        �Ա�,����,��ʶ��,����,���˲���ID,���˿���ID,�ѱ�,�շ����,�շ�ϸĿID,���㵥λ,����,-1*����,
        ���ӱ�־,������ĿID,�վݷ�Ŀ,��׼����,-1*Ӧ�ս��,-1*ʵ�ս��,������,��������ID,������,����ʱ��,
        �Ǽ�ʱ��,ִ�в���ID,v_��Ա���,v_��Ա����
    From ���˷��ü�¼
    Where ��¼����=3 And ��¼״̬=1 And ���ӱ�־=8 And ����ID=����ID_IN And ��ҳID=��ҳID_IN;
    
    --������ܱ�
    For r_Money IN c_Money Loop
        --�������
        Update �������
            Set �������=Nvl(�������,0)-r_Money.ʵ�ս��
         Where ����ID=����ID_IN And ����=1;
        IF SQL%RowCount=0 Then
            Insert Into �������(
                ����ID,����,�������,Ԥ�����)
            Values(
                ����ID_IN,1,-1*r_Money.ʵ�ս��,0);
        End IF;

        --����δ�����
        Update ����δ�����
            Set ���=Nvl(���,0)-r_Money.ʵ�ս��
         Where ����ID=����ID_IN
            And Nvl(��ҳID,0)=Nvl(��ҳID_IN,0)
            And Nvl(���˲���ID,0)=Nvl(r_Money.���˲���ID,0)
            And Nvl(���˿���ID,0)=Nvl(r_Money.���˿���ID,0)
            And Nvl(��������ID,0)=r_Money.��������ID
            And Nvl(ִ�в���ID,0)=r_Money.ִ�в���ID
            And ������ĿID+0=r_Money.������ĿID
            And ��Դ;��=2;

        IF SQL%RowCount=0 Then
            Insert Into ����δ�����(
                ����ID,��ҳID,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,��Դ;��,���)
            Values(
                ����ID_IN,��ҳID_IN,r_Money.���˲���ID,r_Money.���˿���ID,r_Money.��������ID,r_Money.ִ�в���ID,r_Money.������ĿID,2,-1*r_Money.ʵ�ս��);
        End IF;

        --���˷��û���
        Update ���˷��û���
            Set Ӧ�ս��=Nvl(Ӧ�ս��,0)-r_Money.Ӧ�ս��,
                 ʵ�ս��=Nvl(ʵ�ս��,0)-r_Money.ʵ�ս��
         Where ����=Trunc(v_Date)
            And Nvl(���˲���ID,0)=Nvl(r_Money.���˲���ID,0)
            And Nvl(���˿���ID,0)=Nvl(r_Money.���˿���ID,0)
            And Nvl(��������ID,0)=r_Money.��������ID
            And Nvl(ִ�в���ID,0)=r_Money.ִ�в���ID
            And ������ĿID+0=r_Money.������ĿID
            And ��Դ;��=2 And ���ʷ���=1;
        IF SQL%RowCount=0 Then
            Insert Into ���˷��û���(
                ����,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,��Դ;��,���ʷ���,Ӧ�ս��,ʵ�ս��,���ʽ��)
            Values(
                Trunc(v_Date),r_Money.���˲���ID,r_Money.���˿���ID,r_Money.��������ID,r_Money.ִ�в���ID,r_Money.������ĿID,2,1,-1*r_Money.Ӧ�ս��,-1*r_Money.ʵ�ս��,0);
        End IF;
    End Loop;

    --����ԭʼ��¼
    Update ���˷��ü�¼ Set ��¼״̬=3 Where ��¼����=3 And ��¼״̬=1 And ���ӱ�־=8 And ����ID=����ID_IN And ��ҳID=��ҳID_IN;
Exception
    When Err_custom Then Raise_Application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When Others Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_סԺһ�η���_Delete;
/

CREATE OR REPLACE Procedure zl_סԺһ�η���_Insert(
    ����ID_IN        ���˷��ü�¼.����ID%Type,
    ��ҳID_IN        ���˷��ü�¼.��ҳID%Type
)
AS
    Cursor c_Money is
        Select 
            E.����,E.�Ա�,E.����,E.סԺ��,D.��Ժ����,D.��Ժ����ID,D.��Ժ����ID,D.�ѱ�,
            A.���,C.�շ�ϸĿID,A.���㵥λ,B.������ĿID,F.�վݷ�Ŀ,B.�ּ�,
            D.��Ժ����,Nvl(A.ִ�п���,0) AS ִ�п���,Nvl(A.���ηѱ�,0) AS ���ηѱ�
        From �շ�ϸĿ A,�շѼ�Ŀ B,�Զ��Ƽ���Ŀ C,������ҳ D,������Ϣ E,������Ŀ F
        Where A.ID=B.�շ�ϸĿID And A.ID=C.�շ�ϸĿID 
            And C.����ID=D.��Ժ����ID And C.�����־=8 And D.��Ժ����>=Nvl(C.��������,To_Date('3000-01-01','YYYY-MM-DD'))
            And D.��ҳID=��ҳID_IN And D.����ID=����ID_IN 
            And E.����ID=D.����ID And B.������ĿID=F.ID
            And ((D.��Ժ���� Between B.ִ������ and B.��ֹ����) or (D.��Ժ����>=B.ִ������ And B.��ֹ���� is NULL))
        Order BY A.ID,B.������ĿID;

    --���ܣ���סԺ���˼���һ���Է��á�
    v_BillNO        ���˷��ü�¼.NO%Type;
    v_ִ�в���ID	���˷��ü�¼.ִ�в���ID%Type;
    v_ʵ�ս��      ���˷��ü�¼.ʵ�ս��%Type;
    v_�۸񸸺�      ���˷��ü�¼.�۸񸸺�%Type;
    v_��ĿID        �շ���ĿĿ¼.ID%Type;

    v_��Ա���      ���˷��ü�¼.����Ա���%Type;
    v_��Ա����      ���˷��ü�¼.����Ա����%Type;
    v_��Ա����ID    ���ű�.ID%Type;

    v_Dec			Number;        
    v_Date			Date;
    v_Count			Number;
    v_Temp			Varchar2(255);
    v_Error			Varchar2(255);
    Err_Custom		Exception;
Begin
    --����Ƿ��Ǽ���һ�ε���Ŀ
    Select Count(*) Into v_Count
    From ������ҳ A,�Զ��Ƽ���Ŀ B
    Where A.����ID=����ID_IN And A.��ҳID=��ҳID_IN
        And A.��Ժ����ID=B.����ID And B.�����־=8 And A.��Ժ����>=Nvl(B.��������,To_Date('3000-01-01','YYYY-MM-DD'));
    If v_Count=0 Then 
        Return;
    End IF;

    --���ò��˱���סԺ�Ƿ��Ѿ������
    Select 
        Count(*) Into v_Count 
    From ���˷��ü�¼ 
    Where ����ID=����ID_IN And ��ҳID=��ҳID_IN
        And ��¼����=3 And ��¼״̬=1 And ���ӱ�־=8;
    If v_Count>0 Then 
        Return;
    End IF;
    
    --ȡ����Ա��Ϣ(����ID,��������;��ԱID,��Ա���,��Ա����)
    v_Temp:=zl_Identity;
    v_��Ա����ID:=To_Number(Substr(v_Temp,1,Instr(v_Temp,',')-1));
    v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
    v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
    v_��Ա���:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
    v_��Ա����:=Substr(v_Temp,Instr(v_Temp,',')+1);

    --ȡ���ݺ�
    v_BillNO:=zl1_BillAutoNO;
    Update ������Ʊ� Set ������=v_BillNO Where ��Ŀ����='�Զ����ʺ�';
    
    --ȡʱ��
    Select Sysdate Into v_Date From Dual;

    --���С��λ��
    Begin
        Select To_Number(����ֵ) Into v_Dec From ϵͳ������ Where ������=9;
    Exception
        When Others Then v_Dec:=2;
    End;

    --����������Ϣ
    v_��ĿID:=NULL;
    v_Count:=1;--���
    For r_Money In c_Money Loop
        If Nvl(v_��ĿID,0)<>r_Money.�շ�ϸĿID Then
            --��ִ�в���
            IF r_Money.ִ�п���=2 Then
                --��ס����    
                v_ִ�в���ID:=r_Money.��Ժ����ID;
            ElsIF r_Money.ִ�п���=1 Then
                --ָ������
                Begin
                    Select ִ�в���ID Into v_ִ�в���ID From �շ�ִ�в��� Where �շ�ϸĿID=r_Money.�շ�ϸĿID And Rownum<2;
                Exception
                    When Others Then v_ִ�в���ID:=v_��Ա����ID;
                End;
            Else
                --δָ�������Ա����
                v_ִ�в���ID:=v_��Ա����ID;
            End IF;
            --����Ŀ������������Ŀ�ļ۸񸸺�
            v_�۸񸸺�:=v_Count;

        End IF;

        --��ʵ�ս��
        IF r_Money.���ηѱ�=1 Then
            v_ʵ�ս��:=Round(r_Money.�ּ�,v_Dec);
        Else
            Begin
                Select 
                    Round(Round(r_Money.�ּ�,5)*ʵ�ձ���/100,v_Dec) Into v_ʵ�ս��
                From �ѱ���ϸ 
                Where ������ĿID=r_Money.������ĿID And �ѱ�=r_Money.�ѱ� 
                    And Round(r_Money.�ּ�,5) Between Ӧ�ն���ֵ and Ӧ�ն�βֵ;
            Exception
                When Others Then v_ʵ�ս��:=Round(r_Money.�ּ�,v_Dec);
            End;
        End IF;
        
        --������ü�¼(���ӱ�־=8,��¼����=3)
        Insert Into ���˷��ü�¼(
            ID,��¼����,NO,��¼״̬,���,�۸񸸺�,����ID,��ҳID,�����־,���ʷ���,����,�Ա�,����,��ʶ��,
            ����,���˲���ID,���˿���ID,�ѱ�,�շ����,�շ�ϸĿID,���㵥λ,����,����,���ӱ�־,������ĿID,
            �վݷ�Ŀ,��׼����,Ӧ�ս��,ʵ�ս��,������,��������ID,������,����ʱ��,�Ǽ�ʱ��,ִ�в���ID,
            ����Ա���,����Ա����)
        Values(
            ���˷��ü�¼_ID.Nextval,3,v_BillNo,1,v_Count,Decode(Sign(Nvl(v_��ĿID,0)-r_Money.�շ�ϸĿID),0,v_�۸񸸺�,NULL),
            ����ID_IN,��ҳID_IN,2,1,r_Money.����,r_Money.�Ա�,r_Money.����,r_Money.סԺ��,r_Money.��Ժ����,
            r_Money.��Ժ����ID,r_Money.��Ժ����ID,r_Money.�ѱ�,r_Money.���,r_Money.�շ�ϸĿID,r_Money.���㵥λ,
            1,1,8,r_Money.������ĿID,r_Money.�վݷ�Ŀ,Round(r_Money.�ּ�,5),Round(r_Money.�ּ�,v_Dec),v_ʵ�ս��,
            v_��Ա����,v_��Ա����ID,v_��Ա����,r_Money.��Ժ����,v_Date,v_ִ�в���ID,v_��Ա���,v_��Ա����);

        v_Count:=v_Count+1;
        v_��ĿID:=r_Money.�շ�ϸĿID;--��¼�ϴδ����е���Ŀ

        --��ػ��ܱ�Ĵ���
        --�������
        Update �������
            Set �������=Nvl(�������,0)+v_ʵ�ս��
         Where ����ID=����ID_IN And ����=1;
        IF SQL%RowCount=0 Then
            Insert Into �������(
                ����ID,����,�������,Ԥ�����)
            Values(
                ����ID_IN,1,v_ʵ�ս��,0);
        End IF;

        --����δ�����
        Update ����δ�����
            Set ���=Nvl(���,0)+v_ʵ�ս��
         Where ����ID=����ID_IN
            And Nvl(��ҳID,0)=Nvl(��ҳID_IN,0)
            And Nvl(���˲���ID,0)=Nvl(r_Money.��Ժ����ID,0)
            And Nvl(���˿���ID,0)=Nvl(r_Money.��Ժ����ID,0)
            And Nvl(��������ID,0)=v_��Ա����ID
            And Nvl(ִ�в���ID,0)=v_ִ�в���ID
            And ������ĿID+0=r_Money.������ĿID
            And ��Դ;��=2;

        IF SQL%RowCount=0 Then
            Insert Into ����δ�����(
                ����ID,��ҳID,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,��Դ;��,���)
            Values(
                ����ID_IN,��ҳID_IN,r_Money.��Ժ����ID,r_Money.��Ժ����ID,v_��Ա����ID,v_ִ�в���ID,r_Money.������ĿID,2,v_ʵ�ս��);
        End IF;

        --���˷��û���
        Update ���˷��û���
            Set Ӧ�ս��=Nvl(Ӧ�ս��,0)+Round(r_Money.�ּ�,v_Dec),
                 ʵ�ս��=Nvl(ʵ�ս��,0)+v_ʵ�ս��
         Where ����=Trunc(v_Date)
            And Nvl(���˲���ID,0)=Nvl(r_Money.��Ժ����ID,0)
            And Nvl(���˿���ID,0)=Nvl(r_Money.��Ժ����ID,0)
            And Nvl(��������ID,0)=v_��Ա����ID
            And Nvl(ִ�в���ID,0)=v_ִ�в���ID
            And ������ĿID+0=r_Money.������ĿID
            And ��Դ;��=2 And ���ʷ���=1;
        IF SQL%RowCount=0 Then
            Insert Into ���˷��û���(
                ����,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,��Դ;��,���ʷ���,Ӧ�ս��,ʵ�ս��,���ʽ��)
            Values(
                Trunc(v_Date),r_Money.��Ժ����ID,r_Money.��Ժ����ID,v_��Ա����ID,v_ִ�в���ID,r_Money.������ĿID,2,1,Round(r_Money.�ּ�,v_Dec),v_ʵ�ս��,0);
        End IF;
    End Loop;
Exception
    When Err_custom Then Raise_Application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When Others Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_סԺһ�η���_Insert;
/

-------------------------------------------------------
--ģ�飺���˹Һż�¼.SQL
Create Or Replace Procedure zl_���˹Һż�¼_Delete(
    ���ݺ�_IN                   ���˷��ü�¼.NO%Type,
    ����Ա���_IN           ���˷��ü�¼.����Ա���%Type,
    ����Ա����_IN           ���˷��ü�¼.����Ա����%Type,
    ɾ�������_IN           Number:=0,
    ҽ���˷ѽ���_IN		Varchar2:=Null                     --ҽ����������˷ѽ��㷽ʽ,�ձ�ʾȫ������
) AS
    --���α������ж��Ƿ񵥶��ղ�����,���ҺŻ��ܱ���
    Cursor c_RegistInfo(v_״̬ ���˷��ü�¼.��¼״̬%Type) IS
        Select A.����ʱ��,A.�Ǽ�ʱ��,B.��ĿID,B.����ID,B.ҽ������,B.ҽ��ID
        From ���˷��ü�¼ A,�ҺŰ��� B
        Where A.��¼����=4 And A.��¼״̬=v_״̬
            And A.NO=���ݺ�_IN
            And Nvl(A.���㵥λ,'�ű�')=B.����
            And ROWNUM=1;
    r_RegistRow c_RegistInfo%ROWTYPE;

    --���α������жϼ�¼�Ƿ����,�����û��ܱ���
    Cursor c_MoneyInfo(v_״̬ ���˷��ü�¼.��¼״̬%Type) IS
        Select ���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,
                Nvl(Sum(Ӧ�ս��),0) AS Ӧ��,
                Nvl(Sum(ʵ�ս��),0) AS ʵ��,
                Nvl(Sum(���ʽ��),0) AS ����
        From ���˷��ü�¼
        Where ��¼����=4 And ��¼״̬=v_״̬ And NO=���ݺ�_IN
        Group BY ���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID;
    r_MoneyRow c_MoneyInfo%ROWTYPE;
	
    --�ù�����ڴ�����Ա�ɿ�������˵Ĳ�ͬ���㷽ʽ�Ľ��
    Cursor c_OperMoney is
		Select DISTINCT B.���㷽ʽ,B.��Ԥ�� 
		From ���˷��ü�¼ A,����Ԥ����¼ B
		Where A.����ID=B.����ID And A.NO=���ݺ�_IN
			And A.��¼����=4 And A.��¼״̬=3
			And B.��¼����=4 And B.��¼״̬=3 And Nvl(B.��Ԥ��,0)<>0;

	v_��ӡID    Ʊ�ݴ�ӡ����.ID%Type;
	v_����ID    ���˷��ü�¼.����ID%Type;
	v_����ID    ���˷��ü�¼.ID%Type;--���������ջ�Ʊ��
	v_����ID    ������Ϣ.����ID%Type;
	v_Ԥ�����	����Ԥ����¼.��Ԥ��%Type;
  V_�˷ѽ��  ����Ԥ����¼.��Ԥ��%Type;
	
	v_ԤԼ�Һ�	Number;

	v_Count		    Number;
	v_Date		    Date;
	v_Error		    Varchar(255);
	Err_Custom	Exception;
Begin
	--�����ж�Ҫ�˺�/ȡ��ԤԼ�ļ�¼�Ƿ����
    Open c_MoneyInfo(1);
    Fetch c_MoneyInfo Into r_MoneyRow;
	IF c_MoneyInfo%RowCount=0 Then
		Close c_MoneyInfo;
		Open c_MoneyInfo(0);
		Fetch c_MoneyInfo Into r_MoneyRow;
		IF c_MoneyInfo%RowCount=0 Then
			v_Error:='Ҫ����ĵ��ݲ����ڡ�';
			Raise Err_Custom;
		End IF;
		v_ԤԼ�Һ�:=1;
	End IF;        
	Close c_MoneyInfo;

	If Nvl(v_ԤԼ�Һ�,0)=1 Then
		--������Լ��
		Open c_RegistInfo(0);
		Fetch c_RegistInfo Into r_RegistRow;
		Update ���˹ҺŻ���
			Set ��Լ��=Nvl(��Լ��,0) - 1
		Where ����=Trunc(r_RegistRow.����ʱ��)
			And ����ID=r_RegistRow.����ID
			And ��ĿID=r_RegistRow.��ĿID
			And Nvl(ҽ������,'ҽ��')=Nvl(r_RegistRow.ҽ������,'ҽ��')
			And Nvl(ҽ��ID,0)=Nvl(r_RegistRow.ҽ��ID,0);

		IF SQL%RowCount=0 Then
			Insert Into ���˹ҺŻ���(
				����,����ID,��ĿID,ҽ������,ҽ��ID,��Լ��)
			Values(
				Trunc(r_RegistRow.����ʱ��),r_RegistRow.����ID,r_RegistRow.��ĿID,r_RegistRow.ҽ������,
				Decode(r_RegistRow.ҽ��ID,0,Null,r_RegistRow.ҽ��ID),-1);
		End If;
		Close c_RegistInfo;

		--ɾ�����˷��ü�¼
		Delete From ���˷��ü�¼ Where NO=���ݺ�_IN And ��¼����=4 And ��¼״̬=0;
	Else
		Select Sysdate,���˽��ʼ�¼_ID.Nextval Into v_Date,v_����ID From Dual;
		
		--���˾���״̬
		Select ����ID Into v_����ID From ���˷��ü�¼ Where ��¼����=4 And ��¼״̬=1 And NO=���ݺ�_IN And ���=1;
		IF v_����ID IS Not NULL Then
			Update ������Ϣ Set ����״̬=0,��������=NULL Where ����ID=v_����ID;      		
			--ɾ���������ش���,ֻ�е�ֻ��һ���Һż�¼���Ҳ��˽���������Һ����ڽ���ʱ�Żᴦ��
			If ɾ�������_IN=1 Then
				Delete ���ﲡ����¼ Where ����Id=v_����ID;
				Update ������Ϣ Set �����=Null Where ����Id=v_����ID;
				--���ü�¼�����Һż����������￨����,�Լ����˽��Ѻ��˷ѻ����ʵķ���,�Һż�¼�������
				Update ���˷��ü�¼ Set ��ʶ��=Null Where �����־=1 And ����Id=v_����ID;	
			End If;
		End IF;

		--�����ʱ���˾��￨��,�˷�ʱ������￨��
		v_����ID:=Null;
		Begin
			Select ����ID Into v_����ID From ���˷��ü�¼ Where ��¼����=4 And ��¼״̬=1 And NO=���ݺ�_IN And ���ӱ�־=2 And Rownum<2;			
		Exception
			When Others Then NULL;
		End;
		IF v_����ID IS Not NULL Then
			Update ������Ϣ Set ���￨��=NULL,����֤��=NULL Where ����ID=v_����ID;
		End IF;

		--���˷��ü�¼
		Insert Into ���˷��ü�¼(
			ID,NO,ʵ��Ʊ��,��¼����,��¼״̬,���,�۸񸸺�,��������,����ID,��ҳID,���˲���ID,
			���˿���ID,�����־,��ʶ��,����,�Ա�,����,�ѱ�,�շ����,�շ�ϸĿID,���㵥λ,
			����,����,�Ӱ��־,���ӱ�־,��ҩ����,������ĿID,�վݷ�Ŀ,���ʷ���,��׼����,
			Ӧ�ս��,ʵ�ս��,��������ID,������,ִ�в���ID,ִ����,����Ա���,����Ա����,����ʱ��,
			�Ǽ�ʱ��,����ID,���ʽ��)
		Select ���˷��ü�¼_ID.Nextval,NO,ʵ��Ʊ��,��¼����,2,���,�۸񸸺�,��������,����ID,��ҳID,
			���˲���ID,���˿���ID,�����־,��ʶ��,����,�Ա�,����,�ѱ�,�շ����,�շ�ϸĿID,
			���㵥λ,����,-����,�Ӱ��־,���ӱ�־,��ҩ����,������ĿID,�վݷ�Ŀ,���ʷ���,
			��׼����,-Ӧ�ս��,-ʵ�ս��,��������ID,������,ִ�в���ID,ִ����,����Ա���_IN,
			����Ա����_IN,����ʱ��,v_Date,v_����ID,-���ʽ��
		From ���˷��ü�¼
		Where ��¼����=4 And ��¼״̬=1 And NO=���ݺ�_IN;

		Update ���˷��ü�¼ Set ��¼״̬=3 Where ��¼����=4 And ��¼״̬=1 And NO=���ݺ�_IN;

		Select ����ID Into v_Count From ���˷��ü�¼ Where ��¼����=4 And ��¼״̬=3 And NO=���ݺ�_IN And Rownum=1;
		--���˹ҺŽ���:�ֽ�͸����ʻ�����
		IF ҽ���˷ѽ���_IN IS NOT NULL THEN         
            --a.����Ľ��㷽ʽԭ����
            Insert Into ����Ԥ����¼(
              ID,NO,ʵ��Ʊ��,��¼����,��¼״̬,����ID,��ҳID,����ID,ժҪ,
              ���㷽ʽ,�տ�ʱ��,����Ա���,����Ա����,��Ԥ��,����ID)
            Select ����Ԥ����¼_ID.Nextval,NO,ʵ��Ʊ��,��¼����,2,����ID,
              ��ҳID,����ID,ժҪ,���㷽ʽ,v_Date,����Ա���_IN,
              ����Ա����_IN,-��Ԥ��,v_����ID
            From ����Ԥ����¼
            Where ��¼����=4 And ��¼״̬=1 And ����ID=v_Count And ���㷽ʽ<>ҽ���˷ѽ���_IN;
            
            --b.����������ֽ�
            Begin
              Select ��Ԥ�� Into V_�˷ѽ�� From ����Ԥ����¼
                   Where ��¼����=4 And ��¼״̬=1 And ����ID=v_Count And ���㷽ʽ=ҽ���˷ѽ���_IN;
            Exception
                   When Others Then V_�˷ѽ��:=0;
            End;
            If V_�˷ѽ��<>0 Then
               Update ����Ԥ����¼ Set ��Ԥ��=��Ԥ��-V_�˷ѽ�� Where ��¼����=4 And ��¼״̬=2 And ����ID=v_����ID 
                      And ���㷽ʽ=(Select ���� From ���㷽ʽ Where ����=1);
                      
               If Sql%Rowcount=0 Then
                  Insert Into ����Ԥ����¼(
                    ID,NO,ʵ��Ʊ��,��¼����,��¼״̬,����ID,��ҳID,����ID,ժҪ,
                    ���㷽ʽ,�տ�ʱ��,����Ա���,����Ա����,��Ԥ��,����ID)
                  Select ����Ԥ����¼_ID.Nextval,A.NO,A.ʵ��Ʊ��,A.��¼����,2,A.����ID,
                    A.��ҳID,A.����ID,A.ժҪ,B.����,v_Date,����Ա���_IN,
                    ����Ա����_IN,-1*V_�˷ѽ��,v_����ID
                  From ����Ԥ����¼ A,���㷽ʽ B
                  Where B.����=1 And A.��¼����=4 And A.��¼״̬=1 And A.����ID=v_Count And Rownum=1;
               End If;        
            End If;        
        Else
          Insert Into ����Ԥ����¼(
                ID,NO,ʵ��Ʊ��,��¼����,��¼״̬,����ID,��ҳID,����ID,ժҪ,
                ���㷽ʽ,�տ�ʱ��,����Ա���,����Ա����,��Ԥ��,����ID)
            Select ����Ԥ����¼_ID.Nextval,NO,ʵ��Ʊ��,��¼����,2,����ID,
                ��ҳID,����ID,ժҪ,���㷽ʽ,v_Date,����Ա���_IN,
                ����Ա����_IN,-��Ԥ��,v_����ID
            From ����Ԥ����¼
            Where ��¼����=4 And ��¼״̬=1 And ����ID=v_Count;
        End If;

		Update ����Ԥ����¼ Set ��¼״̬=3 Where ��¼����=4 And ��¼״̬=1 And ����ID=v_Count;

		--���˹ҺŽ���:��Ԥ�����
		Insert Into ����Ԥ����¼(
		    ID,NO,ʵ��Ʊ��,��¼����,��¼״̬,����ID,��ҳID,����ID,���,���㷽ʽ,�������,ժҪ,
		    �ɿλ,��λ������,��λ�ʺ�,�տ�ʱ��,����Ա����,����Ա���,��Ԥ��,����ID)
		Select
		    ����Ԥ����¼_ID.Nextval,NO,ʵ��Ʊ��,11,��¼״̬,����ID,��ҳID,����ID,Null,
		    ���㷽ʽ,�������,ժҪ,�ɿλ,��λ������,��λ�ʺ�,�տ�ʱ��,����Ա����,
		    ����Ա���,-1*��Ԥ��,v_����ID
		From ����Ԥ����¼
		Where ��¼���� IN(1,11) And ����ID=v_Count;

		--������Ԥ�����
		Begin
		    Select ����ID,Sum(Nvl(��Ԥ��,0)) Into v_����ID,v_Ԥ����� From ����Ԥ����¼ 
		    Where ��¼���� IN(1,11) And ����ID=v_Count 
		    Group by ����ID;
		Exception
		    When Others Then NULL;
		End;
		IF Nvl(v_����ID,0)<>0 And Nvl(v_Ԥ�����,0)<>0 Then
		    Update ������� Set Ԥ�����=Nvl(Ԥ�����,0)+v_Ԥ����� Where ����ID=v_����ID And ����=1;
		    IF SQL%RowCount=0 Then
			     Insert Into �������(����ID,Ԥ�����,����) Values(v_����ID,v_Ԥ�����,1);
		    End IF;
		    Delete From ������� Where ����ID=v_����ID And ����=1 And Nvl(Ԥ�����,0)=0 And Nvl(�������,0)=0;
		End IF;

		--�˿��ջ�Ʊ��(�����ϴιҺ�ʹ��Ʊ��,�����ջ�)
		Begin
			--�����һ�εĴ�ӡ������ȡ
			Select Max(ID) Into v_��ӡID From Ʊ�ݴ�ӡ���� Where ��������=4 And NO=���ݺ�_IN;
		Exception
			When Others Then NULL;
		End;
		IF v_��ӡID IS Not NULL Then
			Insert Into Ʊ��ʹ����ϸ(
				ID,Ʊ��,����,����,ԭ��,����ID,��ӡID,ʹ��ʱ��,ʹ����)
			Select
				Ʊ��ʹ����ϸ_ID.Nextval,Ʊ��,����,2,2,����ID,��ӡID,v_Date,����Ա����_IN
			From Ʊ��ʹ����ϸ
			Where ��ӡID=v_��ӡID And ����=1;
		End If;

		--��ػ��ܱ�Ĵ���

		--���˹ҺŻ���
		Open c_RegistInfo(3);
		Fetch c_RegistInfo Into r_RegistRow;
		IF c_RegistInfo%RowCount=0 Then
			--ֻ�ղ�����ʱ�޺ű�,������
			Close c_RegistInfo;
		Else
			Update ���˹ҺŻ���
				Set �ѹ���=Nvl(�ѹ���,0) - 1
			Where ����=Trunc(r_RegistRow.�Ǽ�ʱ��)
				And ����ID=r_RegistRow.����ID
				And ��ĿID=r_RegistRow.��ĿID
				And Nvl(ҽ������,'ҽ��')=Nvl(r_RegistRow.ҽ������,'ҽ��')
				And Nvl(ҽ��ID,0)=Nvl(r_RegistRow.ҽ��ID,0);

			IF SQL%RowCount=0 Then
				Insert Into ���˹ҺŻ���(
					����,����ID,��ĿID,ҽ������,ҽ��ID,�ѹ���)
				Values(
					Trunc(r_RegistRow.�Ǽ�ʱ��),r_RegistRow.����ID,r_RegistRow.��ĿID,r_RegistRow.ҽ������,
					Decode(r_RegistRow.ҽ��ID,0,Null,r_RegistRow.ҽ��ID),-1);
			End If;
			Close c_RegistInfo;
		End If;

		--���˷��û���:һ�ŹҺŵ��������(��������)
		For r_MoneyRow IN c_MoneyInfo(3) Loop
			Update ���˷��û���
				Set Ӧ�ս��=Nvl(Ӧ�ս��,0) +(-1 * r_MoneyRow.Ӧ��),
					ʵ�ս��=Nvl(ʵ�ս��,0) +(-1 * r_MoneyRow.ʵ��),
					���ʽ��=Nvl(���ʽ��,0) +(-1 * r_MoneyRow.����)
			Where ����=Trunc(v_Date)
				And Nvl(���˲���ID,0)=Nvl(r_MoneyRow.���˲���ID,0)
				And Nvl(���˿���ID,0)=Nvl(r_MoneyRow.���˿���ID,0)
				And Nvl(��������ID,0)=Nvl(r_MoneyRow.��������ID,0)
				And Nvl(ִ�в���ID,0)=Nvl(r_MoneyRow.ִ�в���ID,0)
				And ������ĿID+0=r_MoneyRow.������ĿID
				And ��Դ;��=1
				And ���ʷ���=0;

			IF SQL%RowCount=0 Then
				Insert Into ���˷��û���(
					����,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,
					��Դ;��,���ʷ���,Ӧ�ս��,ʵ�ս��,���ʽ��)
				Values(
					Trunc(v_Date),
					Decode(r_MoneyRow.���˲���ID,0,Null,r_MoneyRow.���˲���ID),
					Decode(r_MoneyRow.���˿���ID,0,Null,r_MoneyRow.���˿���ID),
					Decode(r_MoneyRow.��������ID,0,Null,r_MoneyRow.��������ID),
					Decode(r_MoneyRow.ִ�в���ID,0,Null,r_MoneyRow.ִ�в���ID),
					r_MoneyRow.������ĿID,1,0,
					-1 * r_MoneyRow.Ӧ��,-1 * r_MoneyRow.ʵ��,-1 * r_MoneyRow.����);
			End If;
		End Loop;

		--��Ա�ɿ����(���������ʻ��ȵĽ�����,�����˳�Ԥ����)
		For r_OperMoney in c_OperMoney Loop
			Update ��Ա�ɿ����
				Set ���=Nvl(���,0) +(-1*r_OperMoney.��Ԥ��)
			 Where �տ�Ա=����Ա����_IN And ���㷽ʽ=r_OperMoney.���㷽ʽ And ����=1;
			IF SQL%RowCount=0 Then
				Insert Into ��Ա�ɿ����(
					�տ�Ա,���㷽ʽ,����,���)
				Values(
					����Ա����_IN,r_OperMoney.���㷽ʽ,1,-1*r_OperMoney.��Ԥ��);
			End IF;
		End Loop;
		Delete From ��Ա�ɿ���� Where �տ�Ա=����Ա����_IN And ����=1 And Nvl(���,0)=0;

		--���˹Һż�¼
		Delete From ���˹Һż�¼ Where NO=���ݺ�_IN;
	End IF;
Exception
    When Err_Custom Then Raise_application_errOr(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]'); 
    When Others Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_���˹Һż�¼_Delete;
/

--���ܣ��������ﲡ�˹Һź������Һš�
CREATE OR REPLACE Procedure zl_���˹Һż�¼_Insert(
    ����ID_IN			���˷��ü�¼.����ID%Type, 
    �����_IN			���˷��ü�¼.��ʶ��%Type, 
    ����_IN				���˷��ü�¼.����%Type, 
    �Ա�_IN				���˷��ü�¼.�Ա�%Type, 
    ����_IN				���˷��ü�¼.����%Type, 
	����_IN				���˷��ü�¼.����%Type,--���ڴ�Ų��˵�ҽ�Ƹ��ʽ���
    �ѱ�_IN				���˷��ü�¼.�ѱ�%Type, 
    ���ݺ�_IN			���˷��ü�¼.NO%Type, 
    Ʊ�ݺ�_IN			���˷��ü�¼.ʵ��Ʊ��%Type, 
    ���_IN				���˷��ü�¼.���%Type, 
    �۸񸸺�_IN			���˷��ü�¼.�۸񸸺�%Type, 
	��������_IN			���˷��ü�¼.��������%Type,
    �շ����_IN			���˷��ü�¼.�շ����%Type, 
    �շ�ϸĿID_IN		���˷��ü�¼.�շ�ϸĿID%Type, 
	����_IN				���˷��ü�¼.����%Type,
	��׼����_IN			���˷��ü�¼.��׼����%Type,
    ������ĿID_IN		���˷��ü�¼.������ĿID%Type, 
    �վݷ�Ŀ_IN			���˷��ü�¼.�վݷ�Ŀ%Type, 
    ���㷽ʽ_IN			����Ԥ����¼.���㷽ʽ%Type,--�ֽ�Ľ�������
    Ӧ�ս��_IN			���˷��ü�¼.Ӧ�ս��%Type, 
    ʵ�ս��_IN			���˷��ü�¼.ʵ�ս��%Type, 
    ���˿���ID_IN		���˷��ü�¼.���˿���ID%Type, 
    ��������ID_IN		���˷��ü�¼.��������ID%Type, 
    ִ�в���ID_IN		���˷��ü�¼.ִ�в���ID%Type, 
    ����Ա���_IN		���˷��ü�¼.����Ա���%Type, 
    ����Ա����_IN		���˷��ü�¼.����Ա����%Type, 
    ����ʱ��_IN			���˷��ü�¼.����ʱ��%Type, 
    �Ǽ�ʱ��_IN			���˷��ü�¼.�Ǽ�ʱ��%Type, 
    ҽ������_IN			�ҺŰ���.ҽ������%Type,
    ҽ��ID_IN			�ҺŰ���.ҽ��ID%Type, 
    ������_IN			Number,--������¼�Ƿ���������
    ����_IN				Number,
    �ű�_IN				�ҺŰ���.����%Type, 
    ����_IN				���˷��ü�¼.��ҩ����%Type, 
    ����ID_IN			���˷��ü�¼.����ID%Type, 
    ����ID_IN			Ʊ��ʹ����ϸ.����ID%Type,
    Ԥ��֧��_IN			����Ԥ����¼.��Ԥ��%Type,--ˢ���Һ�ʱʹ�õ�Ԥ�����,���Ϊ1����.
    �ֽ�֧��_IN			����Ԥ����¼.��Ԥ��%Type,--�Һ�ʱ�ֽ�֧�����ݽ��,���Ϊ1����.
    ����֧��_IN			����Ԥ����¼.��Ԥ��%Type,--�Һ�ʱ�����ʻ�֧�����,,���Ϊ1����.
    ���մ���ID_IN		���˷��ü�¼.���մ���ID%Type,
    ������Ŀ��_IN		���˷��ü�¼.������Ŀ��%Type,
    ͳ����_IN			���˷��ü�¼.ͳ����%Type,
	ժҪ_IN				���˷��ü�¼.ժҪ%Type,--ԤԼ�Һ�ժҪ��Ϣ
	ԤԼ�Һ�_IN			Number:=0, --ԤԼ�Һ�ʱ��(��¼״̬=0,����ʱ��ΪԤԼʱ��),��ʱ����Ҫ���������ز���
	�շ�Ʊ��_IN			Number:=0, --�Һ��Ƿ�ʹ���շ�Ʊ��
    ���ձ���_IN         ���˷��ü�¼.���ձ���%Type
) AS 
	--���α������շѳ�Ԥ���Ŀ���Ԥ���б�
    --��ID�������ȳ��ϴ�δ����ġ� 
    Cursor c_Deposit(v_����ID ������Ϣ.����ID%Type) is 
        Select * From( 
            Select A.ID,A.��¼״̬,A.NO,Nvl(A.���,0) as ��� 
            From ����Ԥ����¼ A,( 
                Select NO,Sum(Nvl(A.���,0)) as ��� 
                From ����Ԥ����¼ A 
                Where A.����ID is Null And Nvl(A.���,0)<>0 
					And A.����ID=v_����ID 
				Group by NO Having Sum(Nvl(A.���,0))<>0 
				) B
        Where A.����ID is Null And Nvl(A.���,0)<>0 
			And A.���㷽ʽ Not IN(Select ���� From ���㷽ʽ Where ����=5)
			And A.NO=B.NO And A.����ID=v_����ID 
        Union All 
        Select 0 as ID,��¼״̬,NO,Sum(Nvl(���,0)-Nvl(��Ԥ��,0)) as ��� 
        From ����Ԥ����¼ 
        Where ��¼���� IN(1,11) And ����ID is Not NULL
			And Nvl(���,0)<>Nvl(��Ԥ��,0) And ����ID=v_����ID 
        Having Sum(Nvl(���,0)-Nvl(��Ԥ��,0))<>0 
        Group by ��¼״̬,NO) 
        Order by ID,NO; 
    --���ܣ�����һ�в��˹Һŷ��ã�����������ܵ�����Ԥ����¼ 
    --       ͬʱ������صĻ��ܱ�(���˹ҺŻ��ܡ����˷��û���) 
    --       ��һ�з��ô���Ʊ��ʹ�����(����ID_IN>0) 
    
    v_��ӡID    Ʊ�ݴ�ӡ����.ID%Type;
    v_����ID    ���˷��ü�¼.ID%Type;
    v_Ԥ�����    ����Ԥ����¼.���%Type;

    v_�ֽ�        ���㷽ʽ.����%Type;
    v_�����ʻ�    ���㷽ʽ.����%Type;

    v_Count        Number;  
    v_Error        Varchar2(255); 
    Err_Custom    Exception; 
Begin 
	If Nvl(ԤԼ�Һ�_IN,0)=0 Then
		--��Ϊ�����а��ձ൥�ݺŹ���,�չҺ������ܳ���10000��,����Ҫ���ΨһԼ����
		Select Count(*) Into v_Count From ���˷��ü�¼ Where ��¼����=4 And ��¼״̬ IN(1,3) And ���=���_IN And NO=���ݺ�_IN;
		If v_Count<>0 THEN 
			v_Error:='�Һŵ��ݺ��ظ�,���ܱ��棡'||CHR(13)||'���ʹ���˰���˳����,���չҺ������ܳ���10000�˴Ρ�';
			Raise Err_Custom;
		End IF;

		--��ȡ���㷽ʽ����
		Begin
			Select ���� Into v_�ֽ� From ���㷽ʽ Where ����=1;
		Exception
			When Others Then v_�ֽ�:='�ֽ�';
		End;
		Begin
			Select ���� Into v_�����ʻ� From ���㷽ʽ Where ����=3;
		Exception
			When Others Then v_�����ʻ�:='�����ʻ�';
		End;
	End IF;

    --�������˹Һŷ���(���ܵ����ǻ������������) 
    Select ���˷��ü�¼_ID.Nextval Into v_����ID From Dual; --Ӧ��ͨ������õ�
    Insert Into ���˷��ü�¼( 
        ID,��¼����,��¼״̬,���,�۸񸸺�,��������,NO,ʵ��Ʊ��,�����־,�Ӱ��־,���ӱ�־,��ҩ����,����ID,
		��ʶ��, ����,����,�Ա�,����,�ѱ�,���˲���ID,���˿���ID,�շ����,���㵥λ,�շ�ϸĿID,������ĿID, 
        �վݷ�Ŀ,����,����,��׼����,Ӧ�ս��,ʵ�ս��,���ʽ��,����ID,���ʷ���,��������ID,������,
		ִ�в���ID,ִ����,����Ա���,����Ա����,����ʱ��,�Ǽ�ʱ��,���մ���ID,������Ŀ��,���ձ���,ͳ����,ժҪ) 
    Values( 
        v_����ID,4,Decode(ԤԼ�Һ�_IN,1,0,1),���_IN,Decode(�۸񸸺�_IN,0,Null,�۸񸸺�_IN),��������_IN,
		���ݺ�_IN,Ʊ�ݺ�_IN,1,����_IN,������_IN,����_IN,Decode(����ID_IN,0,Null,����ID_IN), 
        Decode(�����_IN,0,Null,�����_IN),����_IN,����_IN,Decode(����_IN,Null,Null,�Ա�_IN), 
        Decode(����_IN,Null,Null,����_IN),�ѱ�_IN,���˿���ID_IN,���˿���ID_IN,�շ����_IN, 
        �ű�_IN,�շ�ϸĿID_IN,������ĿID_IN,�վݷ�Ŀ_IN,1,����_IN,��׼����_IN,Ӧ�ս��_IN,ʵ�ս��_IN, 
        Decode(ԤԼ�Һ�_IN,1,NULL,ʵ�ս��_IN),Decode(ԤԼ�Һ�_IN,1,NULL,����ID_IN),0,��������ID_IN,����Ա����_IN,
		ִ�в���ID_IN,ҽ������_IN,����Ա���_IN, ����Ա����_IN,����ʱ��_IN,�Ǽ�ʱ��_IN,���մ���ID_IN,������Ŀ��_IN,���ձ���_IN,ͳ����_IN,ժҪ_IN); 
 
    --���ܽ��㵽����Ԥ����¼
	If Nvl(ԤԼ�Һ�_IN,0)=0 Then
		IF Nvl(�ֽ�֧��_IN,0)<> 0 And ���_IN=1 THEN     
			Insert Into ����Ԥ����¼( 
				ID,��¼����,��¼״̬,NO,����ID,���㷽ʽ,��Ԥ��, 
				�տ�ʱ��,����Ա���,����Ա����,����ID,ժҪ) 
			Values( 
				����Ԥ����¼_ID.Nextval,4,1,���ݺ�_IN,Decode(����ID_IN,0,Null,����ID_IN), 
				Nvl(���㷽ʽ_IN,v_�ֽ�),�ֽ�֧��_IN,�Ǽ�ʱ��_IN,����Ա���_IN,����Ա����_IN,����ID_IN,'�Һ��շ�'); 
		END IF;
		
		--����ҽ���Һ�
		IF Nvl(����֧��_IN,0)<> 0 And ���_IN=1 THEN
			Insert Into ����Ԥ����¼( 
				ID,��¼����,��¼״̬,NO,����ID,���㷽ʽ,��Ԥ��, 
				�տ�ʱ��,����Ա���,����Ա����,����ID,ժҪ) 
			Values( 
				����Ԥ����¼_ID.Nextval,4,1,���ݺ�_IN,Decode(����ID_IN,0,Null,����ID_IN), 
				v_�����ʻ�,����֧��_IN,�Ǽ�ʱ��_IN,����Ա���_IN,����Ա����_IN,����ID_IN,'ҽ���Һ�');
		END IF;
	  
		--���ھ��￨ͨ��Ԥ����Һ� 
		IF Nvl(Ԥ��֧��_IN,0)<> 0 And ���_IN=1 THEN
			v_Ԥ�����:=Ԥ��֧��_IN; 
			For r_Deposit IN c_Deposit(����ID_IN) Loop 
				IF r_Deposit.ID <> 0 Then 
					--��һ�γ�Ԥ�� 
					Update ����Ԥ����¼ 
						Set ��Ԥ��=Decode(Sign(r_Deposit.���-v_Ԥ�����),-1,r_Deposit.���,v_Ԥ�����), 
							����ID=����ID_IN 
					Where ID=r_Deposit.ID; 
				Else 
					--���ϴ�ʣ��� 
					INSERT Into ����Ԥ����¼( 
						ID,NO,ʵ��Ʊ��,��¼����,��¼״̬,����ID,��ҳID,����ID,���, 
						���㷽ʽ,�������,ժҪ,�ɿλ,��λ������,��λ�ʺ�,�տ�ʱ��, 
						����Ա����,����Ա���,��Ԥ��,����ID) 
					Select ����Ԥ����¼_ID.Nextval,NO,ʵ��Ʊ��,11,��¼״̬,����ID, 
						��ҳID,����ID,NULL,���㷽ʽ,�������,ժҪ,�ɿλ, 
						��λ������,��λ�ʺ�,�տ�ʱ��,����Ա����,����Ա���, 
						Decode(Sign(r_Deposit.���-v_Ԥ�����),-1,r_Deposit.���,v_Ԥ�����),����ID_IN 
					From ����Ԥ����¼ 
					Where NO=r_Deposit.NO And ��¼״̬=r_Deposit.��¼״̬ 
						And ��¼���� IN(1,11) And RowNum=1; 
				End IF; 

				--����Ƿ��Ѿ������� 
				IF r_Deposit.���<v_Ԥ����� Then 
					v_Ԥ�����:=v_Ԥ�����-r_Deposit.���; 
				Else 
					v_Ԥ�����:=0; 
				End IF; 

				IF v_Ԥ�����=0 Then 
					Exit; 
				End IF; 
			End Loop; 
		
			--���²���Ԥ����� 
			Update ������� Set Ԥ�����=Nvl(Ԥ�����,0)-Ԥ��֧��_IN Where ����ID=����ID_IN And ����=1; 
			Delete From ������� Where ����ID=����ID_IN And ����=1 And Nvl(�������,0)=0 And Nvl(Ԥ�����,0)=0; 
		End IF; 
		
		--��ػ��ܱ�Ĵ��� 
		--��Ա�ɿ���� 
		IF ���_IN=1 THEN 
			IF Nvl(�ֽ�֧��_IN,0)<> 0 THEN
				Update ��Ա�ɿ���� Set ���=Nvl(���,0)+�ֽ�֧��_IN 
				Where ����=1 And �տ�Ա=����Ա����_IN And ���㷽ʽ=Nvl(���㷽ʽ_IN,v_�ֽ�); 

				IF SQL%RowCount=0 Then 
					Insert Into ��Ա�ɿ����( 
						�տ�Ա,���㷽ʽ,����,���) 
					Values( 
						����Ա����_IN,Nvl(���㷽ʽ_IN,v_�ֽ�),1,�ֽ�֧��_IN); 
				End If; 
			END IF;

			IF Nvl(����֧��_IN,0)<> 0 THEN
				Update ��Ա�ɿ���� Set ���=Nvl(���,0)+����֧��_IN 
				Where ����=1 And �տ�Ա=����Ա����_IN And ���㷽ʽ=v_�����ʻ�; 

				IF SQL%RowCount=0 Then 
					Insert Into ��Ա�ɿ����( 
						�տ�Ա,���㷽ʽ,����,���) 
					Values( 
						����Ա����_IN,v_�����ʻ�,1,����֧��_IN); 
				End If; 
			END IF;
			Delete From ��Ա�ɿ���� Where �տ�Ա=����Ա����_IN And ����=1 And Nvl(���,0)=0; 
		END if;
	End IF;

    --���˹ҺŻ���(ֻ����һ��,�ҵ�����ȡ�����Ѳ�����) 
    IF �ű�_IN is Not Null And ���_IN=1 Then 
		If Nvl(ԤԼ�Һ�_IN,0)=0 Then
			Update ���˹ҺŻ��� 
				Set �ѹ���=Nvl(�ѹ���,0)+1 
			Where ����=Trunc(�Ǽ�ʱ��_IN) 
				And Nvl(����ID,0)=ִ�в���ID_IN 
				And Nvl(��ĿID,0)=�շ�ϸĿID_IN 
				And Nvl(ҽ������,'ҽ��')=Nvl(ҽ������_IN,'ҽ��') 
				And Nvl(ҽ��ID,0)=Nvl(ҽ��ID_IN,0); 

			IF SQL%RowCount=0 Then 
				Insert Into ���˹ҺŻ���( 
					����,����ID,��ĿID,ҽ������,ҽ��ID,�ѹ���) 
				Values( 
					Trunc(�Ǽ�ʱ��_IN),ִ�в���ID_IN,�շ�ϸĿID_IN,ҽ������_IN,Decode(ҽ��ID_IN,0,Null,ҽ��ID_IN),1); 
			End If; 
		Else
			Update ���˹ҺŻ��� 
				Set ��Լ��=Nvl(��Լ��,0)+1 
			Where ����=Trunc(����ʱ��_IN) 
				And Nvl(����ID,0)=ִ�в���ID_IN 
				And Nvl(��ĿID,0)=�շ�ϸĿID_IN 
				And Nvl(ҽ������,'ҽ��')=Nvl(ҽ������_IN,'ҽ��') 
				And Nvl(ҽ��ID,0)=Nvl(ҽ��ID_IN,0); 

			IF SQL%RowCount=0 Then 
				Insert Into ���˹ҺŻ���( 
					����,����ID,��ĿID,ҽ������,ҽ��ID,��Լ��) 
				Values( 
					Trunc(����ʱ��_IN),ִ�в���ID_IN,�շ�ϸĿID_IN,ҽ������_IN,Decode(ҽ��ID_IN,0,Null,ҽ��ID_IN),1); 
			End If; 
		End IF;
    End If; 
    
    --���˷��û��� 
	If Nvl(ԤԼ�Һ�_IN,0)=0 Then
		Update ���˷��û��� 
			Set Ӧ�ս��=Nvl(Ӧ�ս��,0)+Ӧ�ս��_IN, 
				ʵ�ս��=Nvl(ʵ�ս��,0)+ʵ�ս��_IN, 
				���ʽ��=Nvl(���ʽ��,0)+ʵ�ս��_IN 
			Where ����=Trunc(�Ǽ�ʱ��_IN) 
				And Nvl(���˲���ID,0)=���˿���ID_IN 
				And Nvl(���˿���ID,0)=���˿���ID_IN 
				And Nvl(��������ID,0)=��������ID_IN 
				And Nvl(ִ�в���ID,0)=ִ�в���ID_IN 
				And ������ĿID+0=������ĿID_IN 
				And ��Դ;��=1 And ���ʷ���=0; 
		IF SQL%RowCount=0 Then 
			Insert Into ���˷��û���( 
				����,���˲���ID,���˿���ID,��������ID,ִ�в���ID, 
				������ĿID,��Դ;��,���ʷ���,Ӧ�ս��,ʵ�ս��,���ʽ��) 
			Values( 
				Trunc(�Ǽ�ʱ��_IN),���˿���ID_IN,���˿���ID_IN,��������ID_IN, 
				ִ�в���ID_IN,������ĿID_IN,1,0,Ӧ�ս��_IN,ʵ�ս��_IN,ʵ�ս��_IN); 
		End If; 
		
		--����Ʊ��ʹ�����
		IF ���_IN=1 And Ʊ�ݺ�_IN is Not Null Then 
			Select Ʊ�ݴ�ӡ����_ID.Nextval Into v_��ӡID From Dual;

			--����Ʊ�� 
			Insert Into Ʊ�ݴ�ӡ����(
				ID,��������,NO)
			Values(
				v_��ӡID,4,���ݺ�_IN);

			Insert Into Ʊ��ʹ����ϸ( 
				ID,Ʊ��,����,����,ԭ��,����ID,��ӡID,ʹ��ʱ��,ʹ����) 
			Values( 
				Ʊ��ʹ����ϸ_ID.Nextval,Decode(�շ�Ʊ��_IN,1,1,4),Ʊ�ݺ�_IN,1,1,����ID_IN,v_��ӡID,�Ǽ�ʱ��_IN,����Ա����_IN); 

			--״̬�Ķ� 
			Update Ʊ�����ü�¼ 
				Set ��ǰ����=Ʊ�ݺ�_IN,ʣ������=Decode(SIGN(ʣ������-1),-1,0,ʣ������-1) 
			Where ID=Nvl(����ID_IN,0); 
		End If; 
	 
		--���˱��ξ���(�Է���ʱ��Ϊ׼) 
		IF Nvl(����ID_IN,0)<>0 And ���_IN=1 Then 
			Update ������Ϣ 
				Set ����ʱ��=�Ǽ�ʱ��_IN, 
					����״̬=1, 
					��������=����_IN 
			Where ����ID=����ID_IN; 
		End If; 
	End IF;

	--���˹Һż�¼
	IF �ű�_IN is Not Null And ���_IN=1 And Nvl(ԤԼ�Һ�_IN,0)=0 Then 
		Insert Into ���˹Һż�¼(
			ID,NO,����ID,�����,����,�Ա�,����,�ű�,����,����,���ӱ�־,
			ִ�в���ID,ִ����,ִ��״̬,ִ��ʱ��,�Ǽ�ʱ��,����Ա���,����Ա����,ժҪ)
		Values(
			���˹Һż�¼_ID.Nextval,���ݺ�_IN,����ID_IN,�����_IN,����_IN,�Ա�_IN,
			����_IN,�ű�_IN,����_IN,����_IN,NULL,ִ�в���ID_IN,ҽ������_IN,0,NULL,
			�Ǽ�ʱ��_IN,����Ա���_IN,����Ա����_IN,ժҪ_IN);
	End IF;
Exception 
    When Err_Custom Then Raise_application_errOr(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]'); 
    When Others Then zl_ErrOrCenter(SQLCODE,SQLERRM); 
End zl_���˹Һż�¼_Insert; 
/

CREATE OR REPLACE Procedure ZL_ԤԼ�ҺŽ���_INSERT(
    NO_IN				���˷��ü�¼.NO%Type, 
    Ʊ�ݺ�_IN			���˷��ü�¼.ʵ��Ʊ��%Type, 
    ����ID_IN			Ʊ��ʹ����ϸ.����ID%Type,
    ����ID_IN			���˷��ü�¼.����ID%Type, 
    ����_IN				���˷��ü�¼.��ҩ����%Type, 
	����ID_IN			���˷��ü�¼.����ID%Type, 
    �����_IN			���˷��ü�¼.��ʶ��%Type, 
    ����_IN				���˷��ü�¼.����%Type, 
    �Ա�_IN				���˷��ü�¼.�Ա�%Type, 
    ����_IN				���˷��ü�¼.����%Type, 
	����_IN				���˷��ü�¼.����%Type,--���ڴ�Ų��˵�ҽ�Ƹ��ʽ���
    �ѱ�_IN				���˷��ü�¼.�ѱ�%Type, 
    ���㷽ʽ_IN			����Ԥ����¼.���㷽ʽ%Type,--�ֽ�Ľ�������
    �ֽ�֧��_IN			����Ԥ����¼.��Ԥ��%Type,--�Һ�ʱ�ֽ�֧�����ݽ��
	Ԥ��֧��_IN			����Ԥ����¼.��Ԥ��%Type,--�Һ�ʱʹ�õ�Ԥ�����
    ����֧��_IN			����Ԥ����¼.��Ԥ��%Type,--�Һ�ʱ�����ʻ�֧�����
    ����ʱ��_IN			���˷��ü�¼.����ʱ��%Type 
) AS 
	--���α������շѳ�Ԥ���Ŀ���Ԥ���б�
    --��ID�������ȳ��ϴ�δ����ġ� 
    Cursor c_Deposit(v_����ID ������Ϣ.����ID%Type) is 
        Select * From( 
            Select A.ID,A.��¼״̬,A.NO,Nvl(A.���,0) as ��� 
            From ����Ԥ����¼ A,( 
                Select NO,Sum(Nvl(A.���,0)) as ��� 
                From ����Ԥ����¼ A 
                Where A.����ID is Null And Nvl(A.���,0)<>0 
					And A.����ID=v_����ID 
				Group by NO Having Sum(Nvl(A.���,0))<>0 
				) B
        Where A.����ID is Null And Nvl(A.���,0)<>0 
			And A.���㷽ʽ Not IN(Select ���� From ���㷽ʽ Where ����=5)
			And A.NO=B.NO And A.����ID=v_����ID 
        Union All 
        Select 0 as ID,��¼״̬,NO,Sum(Nvl(���,0)-Nvl(��Ԥ��,0)) as ��� 
        From ����Ԥ����¼ 
        Where ��¼���� IN(1,11) And ����ID is Not NULL
			And Nvl(���,0)<>Nvl(��Ԥ��,0) And ����ID=v_����ID 
        Having Sum(Nvl(���,0)-Nvl(��Ԥ��,0))<>0 
        Group by ��¼״̬,NO) 
        Order by ID,NO; 

	--�����˷��û���
	Cursor c_Money is
		Select ���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,
			Sum(Ӧ�ս��) AS Ӧ�ս��,Sum(ʵ�ս��) AS ʵ�ս��,Sum(���ʽ��) AS ���ʽ��
		From ���˷��ü�¼
		Where ��¼����=4 And ��¼״̬=1 And NO=NO_IN
		Group By ���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID;

	--�ű���Ϣ
	Cursor c_Regist is
		Select B.����ID,B.��ĿID,B.ҽ��ID,B.ҽ������
		From ���˷��ü�¼ A,�ҺŰ��� B
		Where A.��¼����=4 And A.��¼״̬=1 And A.NO=NO_IN
			And A.���=1 And A.���㵥λ=B.����;
    r_Regist c_Regist%RowType;	

    v_�ֽ�			���㷽ʽ.����%Type;
    v_�����ʻ�		���㷽ʽ.����%Type;
    v_��ӡID		Ʊ�ݴ�ӡ����.ID%Type;
	v_Ԥ�����		����Ԥ����¼.���%Type;
	
    v_Temp			Varchar2(255);
    v_��Ա���		���˷��ü�¼.����Ա���%Type;
    v_��Ա����		���˷��ü�¼.����Ա����%Type;

	v_Date			Date;
    v_Error			Varchar2(255); 
    Err_Custom		Exception; 
Begin 
	--��ȡ���㷽ʽ����
	Begin
		Select ���� Into v_�ֽ� From ���㷽ʽ Where ����=1;
	Exception
		When Others Then v_�ֽ�:='�ֽ�';
	End;
	Begin
		Select ���� Into v_�����ʻ� From ���㷽ʽ Where ����=3;
	Exception
		When Others Then v_�����ʻ�:='�����ʻ�';
	End;
	Select Sysdate Into v_Date From Dual;
    v_Temp:=zl_Identity;
    v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
    v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
    v_��Ա���:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
    v_��Ա����:=Substr(v_Temp,Instr(v_Temp,',')+1);

	--���²��˷��ü�¼
	Update ���˷��ü�¼
		Set ��¼״̬=1,
			ʵ��Ʊ��=Ʊ�ݺ�_IN,
			����ID=����ID_IN,
			���ʽ��=ʵ�ս��,
			��ҩ����=����_IN,
			����ID=����ID_IN,
			��ʶ��=�����_IN,
			����=����_IN,
			����=����_IN,
			�Ա�=�Ա�_IN,
			����=����_IN,
			�ѱ�=�ѱ�_IN,
			����ʱ��=����ʱ��_IN,
			�Ǽ�ʱ��=v_Date,
			����Ա���=v_��Ա���,
			����Ա����=v_��Ա����
	Where ��¼����=4 And ��¼״̬=0 And NO=NO_IN;
	 
	--���˹Һż�¼
	Insert Into ���˹Һż�¼(
		ID,NO,����ID,�����,����,�Ա�,����,�ű�,����,����,���ӱ�־,
		ִ�в���ID,ִ����,ִ��״̬,ִ��ʱ��,�Ǽ�ʱ��,����Ա���,����Ա����,ժҪ)
	Select
		���˹Һż�¼_ID.Nextval,NO_IN,����ID_IN,�����_IN,����_IN,�Ա�_IN,
		����_IN,���㵥λ,�Ӱ��־,����_IN,NULL,ִ�в���ID,ִ����,0,NULL,
		v_Date,v_��Ա���,v_��Ա����,Nvl(ժҪ,����)
	From ���˷��ü�¼
	Where ��¼����=4 And ��¼״̬=1 And ���=1 And NO=NO_IN;

    --���ܽ��㵽����Ԥ����¼
	IF Nvl(�ֽ�֧��_IN,0)<> 0  THEN     
		Insert Into ����Ԥ����¼( 
			ID,��¼����,��¼״̬,NO,����ID,���㷽ʽ,��Ԥ��, 
			�տ�ʱ��,����Ա���,����Ա����,����ID,ժҪ) 
		Values( 
			����Ԥ����¼_ID.Nextval,4,1,NO_IN,����ID_IN, 
			Nvl(���㷽ʽ_IN,v_�ֽ�),�ֽ�֧��_IN,v_Date,v_��Ա���,v_��Ա����,����ID_IN,'�Һ��շ�'); 
	END IF;
	
	--���ھ��￨ͨ��Ԥ����Һ� 
	IF Nvl(Ԥ��֧��_IN,0)<> 0 THEN
		v_Ԥ�����:=Ԥ��֧��_IN; 
		For r_Deposit IN c_Deposit(����ID_IN) Loop 
			IF r_Deposit.ID <> 0 Then 
				--��һ�γ�Ԥ�� 
				Update ����Ԥ����¼ 
					Set ��Ԥ��=Decode(Sign(r_Deposit.���-v_Ԥ�����),-1,r_Deposit.���,v_Ԥ�����), 
						����ID=����ID_IN 
				Where ID=r_Deposit.ID; 
			Else 
				--���ϴ�ʣ��� 
				INSERT Into ����Ԥ����¼( 
					ID,NO,ʵ��Ʊ��,��¼����,��¼״̬,����ID,��ҳID,����ID,���, 
					���㷽ʽ,�������,ժҪ,�ɿλ,��λ������,��λ�ʺ�,�տ�ʱ��, 
					����Ա����,����Ա���,��Ԥ��,����ID) 
				Select ����Ԥ����¼_ID.Nextval,NO,ʵ��Ʊ��,11,��¼״̬,����ID, 
					��ҳID,����ID,NULL,���㷽ʽ,�������,ժҪ,�ɿλ, 
					��λ������,��λ�ʺ�,�տ�ʱ��,����Ա����,����Ա���, 
					Decode(Sign(r_Deposit.���-v_Ԥ�����),-1,r_Deposit.���,v_Ԥ�����),����ID_IN 
				From ����Ԥ����¼ 
				Where NO=r_Deposit.NO And ��¼״̬=r_Deposit.��¼״̬ 
					And ��¼���� IN(1,11) And RowNum=1; 
			End IF; 

			--����Ƿ��Ѿ������� 
			IF r_Deposit.���<v_Ԥ����� Then 
				v_Ԥ�����:=v_Ԥ�����-r_Deposit.���; 
			Else 
				v_Ԥ�����:=0; 
			End IF; 

			IF v_Ԥ�����=0 Then 
				Exit; 
			End IF; 
		End Loop; 
	
		--���²���Ԥ����� 
		Update ������� Set Ԥ�����=Nvl(Ԥ�����,0)-Ԥ��֧��_IN Where ����ID=����ID_IN And ����=1; 
		Delete From ������� Where ����ID=����ID_IN And ����=1 And Nvl(�������,0)=0 And Nvl(Ԥ�����,0)=0; 
	End IF; 

	--����ҽ���Һ�
	IF Nvl(����֧��_IN,0)<> 0 THEN
		Insert Into ����Ԥ����¼( 
			ID,��¼����,��¼״̬,NO,����ID,���㷽ʽ,��Ԥ��, 
			�տ�ʱ��,����Ա���,����Ա����,����ID,ժҪ) 
		Values( 
			����Ԥ����¼_ID.Nextval,4,1,NO_IN,����ID_IN, v_�����ʻ�,����֧��_IN,
			v_Date,v_��Ա���,v_��Ա����,����ID_IN,'ҽ���Һ�');
	END IF;
  
	--��ػ��ܱ�Ĵ��� 
	--��Ա�ɿ���� 
	IF Nvl(�ֽ�֧��_IN,0)<> 0 THEN
		Update ��Ա�ɿ���� 
			Set ���=Nvl(���,0)+�ֽ�֧��_IN 
		Where ����=1 And �տ�Ա=v_��Ա���� And ���㷽ʽ=Nvl(���㷽ʽ_IN,v_�ֽ�); 

		IF SQL%RowCount=0 Then 
			Insert Into ��Ա�ɿ����( 
				�տ�Ա,���㷽ʽ,����,���) 
			Values( 
				v_��Ա����,Nvl(���㷽ʽ_IN,v_�ֽ�),1,�ֽ�֧��_IN); 
		End If; 
	END IF;

	IF Nvl(����֧��_IN,0)<> 0 THEN
		Update ��Ա�ɿ���� Set ���=Nvl(���,0)+����֧��_IN 
		Where ����=1 And �տ�Ա=v_��Ա���� And ���㷽ʽ=v_�����ʻ�; 

		IF SQL%RowCount=0 Then 
			Insert Into ��Ա�ɿ����( 
				�տ�Ա,���㷽ʽ,����,���) 
			Values( 
				v_��Ա����,v_�����ʻ�,1,����֧��_IN); 
		End If; 
	END IF;
	Delete From ��Ա�ɿ���� Where �տ�Ա=v_��Ա���� And ����=1 And Nvl(���,0)=0; 

    --���˹ҺŻ���(ֻ����һ��,�ҵ�����ȡ�����Ѳ�����) 
	Open c_Regist;
	Fetch c_Regist Into r_Regist;
	Update ���˹ҺŻ��� 
		Set �ѹ���=Nvl(�ѹ���,0)+1 
	Where ����=Trunc(v_Date) 
		And Nvl(����ID,0)=Nvl(r_Regist.����ID,0)
		And Nvl(��ĿID,0)=Nvl(r_Regist.��ĿID,0)
		And Nvl(ҽ������,'ҽ��')=Nvl(r_Regist.ҽ������,'ҽ��') 
		And Nvl(ҽ��ID,0)=Nvl(r_Regist.ҽ��ID,0); 
	IF SQL%RowCount=0 Then 
		Insert Into ���˹ҺŻ���( 
			����,����ID,��ĿID,ҽ������,ҽ��ID,�ѹ���) 
		Values(
			Trunc(v_Date),r_Regist.����ID,r_Regist.��ĿID,r_Regist.ҽ������,r_Regist.ҽ��ID,1); 
	End If;
	Close c_Regist;
    
    --���˷��û��� 
	For r_Money In c_Money Loop
		Update ���˷��û��� 
			Set Ӧ�ս��=Nvl(Ӧ�ս��,0)+Nvl(r_Money.Ӧ�ս��,0), 
				ʵ�ս��=Nvl(ʵ�ս��,0)+Nvl(r_Money.ʵ�ս��,0), 
				���ʽ��=Nvl(���ʽ��,0)+Nvl(r_Money.���ʽ��,0) 
			Where ����=Trunc(v_Date) 
				And Nvl(���˲���ID,0)=Nvl(r_Money.���˲���ID,0)
				And Nvl(���˿���ID,0)=Nvl(r_Money.���˿���ID,0)
				And Nvl(��������ID,0)=Nvl(r_Money.��������ID,0)
				And Nvl(ִ�в���ID,0)=Nvl(r_Money.ִ�в���ID,0)
				And ������ĿID+0=r_Money.������ĿID
				And ��Դ;��=1 And ���ʷ���=0;
		IF SQL%RowCount=0 Then 
			Insert Into ���˷��û���( 
				����,���˲���ID,���˿���ID,��������ID,ִ�в���ID, 
				������ĿID,��Դ;��,���ʷ���,Ӧ�ս��,ʵ�ս��,���ʽ��) 
			Values( 
				Trunc(v_Date),r_Money.���˲���ID,r_Money.���˿���ID,r_Money.��������ID, 
				r_Money.ִ�в���ID,r_Money.������ĿID,1,0,r_Money.Ӧ�ս��,r_Money.ʵ�ս��,r_Money.���ʽ��); 
		End If; 
	End Loop;

	--����Ʊ��ʹ�����
	IF Ʊ�ݺ�_IN is Not Null Then 
		Select Ʊ�ݴ�ӡ����_ID.Nextval Into v_��ӡID From Dual;

		--����Ʊ�� 
		Insert Into Ʊ�ݴ�ӡ����(
			ID,��������,NO)
		Values(
			v_��ӡID,4,NO_IN);

		Insert Into Ʊ��ʹ����ϸ( 
			ID,Ʊ��,����,����,ԭ��,����ID,��ӡID,ʹ��ʱ��,ʹ����) 
		Values( 
			Ʊ��ʹ����ϸ_ID.Nextval,4,Ʊ�ݺ�_IN,1,1,����ID_IN,v_��ӡID,v_Date,v_��Ա����); 

		--״̬�Ķ� 
		Update Ʊ�����ü�¼ 
			Set ��ǰ����=Ʊ�ݺ�_IN,ʣ������=Decode(SIGN(ʣ������-1),-1,0,ʣ������-1) 
		Where ID=Nvl(����ID_IN,0); 
	End If; 
 
	--���˱��ξ���(�Է���ʱ��Ϊ׼) 
	IF Nvl(����ID_IN,0)<>0 Then 
		Update ������Ϣ 
			Set ����ʱ��=v_Date, 
				����״̬=1, 
				��������=����_IN 
		Where ����ID=����ID_IN; 
	End If; 
Exception 
    When Err_Custom Then Raise_application_errOr(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]'); 
    When Others Then zl_ErrOrCenter(SQLCODE,SQLERRM); 
End ZL_ԤԼ�ҺŽ���_INSERT; 
/

Create Or Replace Procedure ZL_���˹Һż�¼_����(
--���ܣ���ɲ��˻��Ź��ܣ��ڹҺ���ĿID��ͬ������¡�
    NO_IN			���˹Һż�¼.NO%Type,
    �ű�_IN			���˹Һż�¼.�ű�%Type,
    ����_IN			���˹Һż�¼.����%Type,
    ����ID_IN		���˹Һż�¼.ִ�в���ID%Type,
    ԭҽ��_IN		���˹Һż�¼.ִ����%Type,
    ԭҽ��ID_IN		���˹ҺŻ���.ҽ��ID%Type,
    ��ҽ��_IN		���˹Һż�¼.ִ����%Type,
    ��ҽ��ID_IN		���˹ҺŻ���.ҽ��ID%Type
) AS
    Cursor c_Bill  IS
	    Select * From ���˷��ü�¼ Where ��¼����=4 And ��¼״̬=1 And NO=NO_IN Order BY ���;

	v_����ID	���˷��ü�¼.ID%Type;
    v_Error		Varchar2(255);
    Err_Custom	Exception;
Begin
	v_����ID:=0;
    Begin
      Select ����ID Into v_����ID From ���˹Һż�¼ Where NO=NO_IN;
    Exception
        When OTHERS Then Null;
    End;
	If v_����ID=0 Then
        v_Error:='û���ҵ����˵ĹҺ���Ϣ��';
        Raise Err_Custom;
    ElsIf v_����ID IS Null Then 
        v_Error:='û���ҵ�������Ϣ��';
        Raise Err_Custom;
    End If;

    ---�ȸ��²�����Ϣ�ľ������Һ�״̬
    Update ������Ϣ Set ��������=����_IN,����״̬=1 Where ����ID=v_����ID And ����״̬ IN(1,2);
     
    For r_Bill IN c_Bill  Loop 
		--�ָ���ǰ�ĹҺŻ���
		Update ���˹ҺŻ��� 
			Set �ѹ���=Nvl(�ѹ���,0)-1 
		Where ����=Trunc(r_Bill.�Ǽ�ʱ��)
			And Nvl(����ID,0)=Nvl(r_Bill.ִ�в���ID,0)
			And Nvl(��ĿID,0)=Nvl(r_Bill.�շ�ϸĿID,0)
			And Nvl(ҽ������,'ҽ��')=Nvl(ԭҽ��_IN,'ҽ��') 
			And Nvl(ҽ��ID,0)=Nvl(ԭҽ��ID_IN,0); 
		If SQL%RowCount=0 Then 
			Insert Into ���˹ҺŻ���(
				����,����ID,��ĿID,ҽ������,ҽ��ID,�ѹ���) 
			Values(
				Trunc(r_Bill.�Ǽ�ʱ��),r_Bill.ִ�в���ID,r_Bill.�շ�ϸĿID,ԭҽ��_IN,Decode(ԭҽ��ID_IN,0,Null,ԭҽ��ID_IN),-1); 
		End If;     
			
		--�ָ���ǰ�ķ��û���
		Update ���˷��û��� 
			Set Ӧ�ս��=Nvl(Ӧ�ս��,0)-Nvl(r_Bill.Ӧ�ս��,0), 
				ʵ�ս��=Nvl(ʵ�ս��,0)-Nvl(r_Bill.ʵ�ս��,0), 
				���ʽ��=Nvl(���ʽ��,0)-Nvl(r_Bill.���ʽ��,0) 
		Where ����=Trunc(r_Bill.�Ǽ�ʱ��)
			And Nvl(���˲���ID,0)=Nvl(r_Bill.���˿���ID,0)
			And Nvl(���˿���ID,0)=Nvl(r_Bill.���˿���ID,0)
			And Nvl(��������ID,0)=Nvl(r_Bill.��������ID,0)
			And Nvl(ִ�в���ID,0)=Nvl(r_Bill.ִ�в���ID,0)
			And ������ĿID+0=r_Bill.������ĿID
			And ��Դ;��=1 And ���ʷ���=0; 
		If SQL%RowCount=0 Then 
			Insert Into ���˷��û���( 
				����,���˲���ID,���˿���ID,��������ID,ִ�в���ID, 
				������ĿID,��Դ;��,���ʷ���,Ӧ�ս��,ʵ�ս��,���ʽ��) 
			Values( 
				Trunc(r_Bill.�Ǽ�ʱ��),r_Bill.���˿���ID,r_Bill.���˿���ID,r_Bill.��������ID, 
				r_Bill.ִ�в���ID,r_Bill.������ĿID,1,0,-1*r_Bill.Ӧ�ս��,-1*r_Bill.ʵ�ս��,-1*r_Bill.���ʽ��); 
		End If; 

		----Ȼ���ٸ��¹ҺŻ���
		Update ���˹ҺŻ��� 
			Set �ѹ���=Nvl(�ѹ���,0)+1 
		Where ����=Trunc(r_Bill.�Ǽ�ʱ��)
			And Nvl(����ID,0)=����ID_IN
			And Nvl(��ĿID,0)=Nvl(r_Bill.�շ�ϸĿID,0)
			And Nvl(ҽ������,'ҽ��')=Nvl(��ҽ��_IN,'ҽ��') 
			And Nvl(ҽ��ID,0)=Nvl(��ҽ��ID_IN,0); 
		If SQL%RowCount=0 Then 
			Insert Into ���˹ҺŻ���(
				����,����ID,��ĿID,ҽ������,ҽ��ID,�ѹ���) 
			Values(
				Trunc(r_Bill.�Ǽ�ʱ��),����ID_IN,r_Bill.�շ�ϸĿID,��ҽ��_IN,Decode(��ҽ��ID_IN,0,Null,��ҽ��ID_IN),1); 
		End If;     
			
		-----Ȼ���ٸ��¹Һŷ��û���
		Update ���˷��û��� 
			Set Ӧ�ս��=Nvl(Ӧ�ս��,0)+Nvl(r_Bill.Ӧ�ս��,0), 
				ʵ�ս��=Nvl(ʵ�ս��,0)+Nvl(r_Bill.ʵ�ս��,0), 
				���ʽ��=Nvl(���ʽ��,0)+Nvl(r_Bill.���ʽ��,0) 
		Where ����=Trunc(r_Bill.�Ǽ�ʱ��)
			And Nvl(���˲���ID,0)=Nvl(r_Bill.���˿���ID,0)
			And Nvl(���˿���ID,0)=Nvl(r_Bill.���˿���ID,0)
			And Nvl(��������ID,0)=Nvl(r_Bill.��������ID,0)
			And Nvl(ִ�в���ID,0)=����ID_IN
			And ������ĿID+0=r_Bill.������ĿID
			And ��Դ;��=1 And ���ʷ���=0; 
		If SQL%RowCount=0 Then 
			Insert Into ���˷��û���( 
				����,���˲���ID,���˿���ID,��������ID,ִ�в���ID, 
				������ĿID,��Դ;��,���ʷ���,Ӧ�ս��,ʵ�ս��,���ʽ��) 
			Values( 
				Trunc(r_Bill.�Ǽ�ʱ��),r_Bill.���˿���ID,r_Bill.���˿���ID,r_Bill.��������ID, 
				����ID_IN,r_Bill.������ĿID,1,0,r_Bill.Ӧ�ս��,r_Bill.ʵ�ս��,r_Bill.���ʽ��); 
		End If; 
		  
		---���¹Һż�¼
		Update ���˷��ü�¼
			Set ִ�в���ID=����ID_IN,
				���˿���ID=����ID_IN,
				���˲���ID=����ID_IN,
				���㵥λ=�ű�_IN,
				��ҩ����=����_IN,
				ִ����=��ҽ��_IN,
				ִ��״̬=0,ִ��ʱ��=Null
		Where ID=r_Bill.ID;

		--���²��˹Һż�¼
		If r_Bill.���=1 Then
			Update ���˹Һż�¼
				Set ִ�в���ID=����ID_IN,
					�ű�=�ű�_IN,
					����=����_IN,
					ִ����=��ҽ��_IN,
					ִ��״̬=0,
					ִ��ʱ��=Null
			Where NO=r_Bill.NO;
		End If;
    End Loop;
Exception
    When Err_Custom Then RAISE_APPLICATION_ERROR (-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
    When OTHERS Then ZL_ERRORCENTER (SQLCODE, SQLERRM);
End ZL_���˹Һż�¼_����;
/

CREATE OR REPLACE Procedure ZL_����ҽ����¼_�ջ�(
--���ܣ���ָ��ҽ�����ڷ��Ͳ����ջء�����ϴη���û�в������ã�����ջ�ҽ�����ϴ�ִ��ʱ�䡣
--������NO_IN=�����ջز����������ݵ��µ��ݺ�(�����ü�ҩƷʹ��),��ǰ�����ֻ����NO��һ���ݡ�
--            ��ΪҩƷ���ܷ���,��������ڴ���ʱȡ��
--      �ջ���_IN=��ҩƷΪ��סԺ��λ���ջ���,������ҽ��Ϊ�ջ������������
--      ҽ��ID_IN=ÿ��Ҫ�ջص�ҽ����¼��ID(��ϸ�洢��ID),�Գ�ҩ���䷽,��һ��������ҩ;�����÷��巨(����Ϊ������δ��ȡ)
--      �ϴ�ʱ��_IN=ҽ�����ڷ��Ͳ����ջغ�Ӧ�û�ԭ���ϴ�ִ��ʱ��(�ϸ�Ƶ�ʼ������),Ϊ��ʱ��ʾ��ȫ���ջ��ˡ�
    NO_IN				���˷��ü�¼.NO%Type,
    �ջ���_IN			����ҽ������.��������%Type,
    ҽ��ID_IN			����ҽ����¼.ID%Type,
    �������_IN			����ҽ����¼.�������%Type,
    �ϴ�ʱ��_IN			����ҽ����¼.�ϴ�ִ��ʱ��%Type
) IS
    --����ָ����ҩ��������ʱ��������ط��ü�ҩƷ��¼��Ϣ(���η��ͻ�������ж�����¼)
    --ҩƷҽ����д��"����ҽ������"��¼,��Ӧ�ĸ�ҩ;����һ����д�˵�(����Ϊ����),��NO��ͬ��
    --��ΪҪ�ջصĴ������ܰ����˶�η��͵�����,����Ҫ����η��͵��շ���¼��ȡ����
    Cursor c_Drug is
        Select A.����ID,A.��ҳID,D.����,
            X.סԺ��װ,X.���Ч��,Nvl(B.����,1)*B.ʵ������ AS ����,
            B.ID AS �շ�ID,B.����,B.ҩƷID,B.�Է�����ID,B.�ⷿID,B.����ID,
            Nvl(X.ҩ������,0) AS ����,B.����,B.����,B.Ч��
        From ���˷��ü�¼ A,ҩƷ�շ���¼ B,����ҽ������ C,������Ϣ D,ҩƷ��� X
        Where C.ҽ��ID=ҽ��ID_IN And A.NO=C.NO And A.��¼����=C.��¼���� And A.��¼״̬ IN(0,1,3)
            And A.ҽ�����+0=ҽ��ID_IN And A.NO=B.NO And A.ID=B.����ID+0
            And B.���� IN(9,10) And (B.��¼״̬=1 Or Mod(B.��¼״̬,3)=0)
            And A.����ID=D.����ID And B.ҩƷID=X.ҩƷID
        Order BY B.NO Desc,B.ID Desc;
    
    --������ҩ����(����ҩ;��)����ʱ�������ķ���(����������ж�����¼)
    --�Է�ҩҽ��,ֱ���ջ�ָ����,���ܶ�η���(�����η��ͼ۸�ͬ,���ջصļ۸��������εģ���Ȼ��Ҫ���ݶ���������μ��ջ���)��
    --��ҩ����Ӧ�ö���д�˷��ͼ�¼(�����˶���������ȼ�)
    Cursor c_Other is
        Select A.ID AS ����ID,Nvl(A.����,1)*A.���� AS ����
        From ���˷��ü�¼ A,����ҽ������ B
        Where A.NO=B.NO And A.��¼����=B.��¼���� And A.��¼״̬ IN(0,1,3)
            And A.ҽ�����+0=ҽ��ID_IN And B.ҽ��ID=ҽ��ID_IN
            And B.���ͺ�=(Select Max(���ͺ�) From ����ҽ������ Where ҽ��ID=ҽ��ID_IN)
        Order BY A.�շ�ϸĿID,A.���;

    --���α����ڴ��������ػ��ܱ�
    Cursor c_Money(v_Start ���˷��ü�¼.���%Type,v_End ���˷��ü�¼.���%Type) IS
        Select ����ID,��ҳID,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,
            Sum(Nvl(Ӧ�ս��,0)) AS Ӧ�ս��,Sum(Nvl(ʵ�ս��,0)) AS ʵ�ս��
        From ���˷��ü�¼ 
        Where ��¼����=2 And ��¼״̬=1 And NO=NO_IN And ��� Between v_Start And v_End
        Group BY ����ID,��ҳID,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID;

    v_Date          Date;
    v_�������      ���˷��ü�¼.���%Type;
    v_�շ����      ҩƷ�շ���¼.���%Type;
    v_����ID        ���˷��ü�¼.ID%Type;
    v_����          ҩƷ�շ���¼.����%Type;
    v_Ч��          ҩƷ�շ���¼.Ч��%Type;
    v_����          ҩƷ�շ���¼.����%Type;
    v_���ȼ�        ���.���ȼ�%Type;
    v_ʵ�ս��      ���˷��ü�¼.ʵ�ս��%Type;
    
    v_��ʼ���      ���˷��ü�¼.���%Type;
    v_�������      ���˷��ü�¼.���%Type;

    v_ʣ����        ҩƷ�շ���¼.ʵ������%Type;
    v_��ǰ��        ҩƷ�շ���¼.ʵ������%Type;
    
	v_��ID			����ҽ����¼.ID%Type;

    v_Temp          Varchar2(255);
    v_��Ա���      ���˷��ü�¼.����Ա���%Type;
    v_��Ա����      ���˷��ü�¼.����Ա����%Type;

	v_ҩƷ����		Number;
	v_��������		Number;

    v_Dec			Number;
    v_Error         Varchar2(255);
    Err_Custom      Exception;
Begin
    --���С��λ��
    Begin
        Select To_Number(����ֵ) Into v_Dec From ϵͳ������ Where ������=9;
    Exception
        When Others Then v_Dec:=2;
    End;

    --ȡ����Ա��Ϣ(����ID,��������;��ԱID,��Ա���,��Ա����)
    v_Temp:=zl_Identity;
    v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
    v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
    v_��Ա���:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
    v_��Ա����:=Substr(v_Temp,Instr(v_Temp,',')+1);

	--���ɻ��۵�ϵͳ����
	Begin
		Select To_Number(Nvl(����ֵ,'0')) Into v_ҩƷ���� From ϵͳ������ Where ������=79;
	Exception
		When Others Then v_ҩƷ����:=0;
	End;
	Begin
		Select To_Number(Nvl(����ֵ,'0')) Into v_�������� From ϵͳ������ Where ������=80;
	Exception
		When Others Then v_��������:=0;
	End;

    Select Sysdate Into v_Date From Dual;
    v_��ʼ���:=NULL;v_�������:=NULL;

    If Nvl(�ջ���_IN,0)<>0 Then
        If �������_IN IN('5','6') Then
            --�У�����ҩ
            -----------------------------------------------------------------------------------------------------
            v_ʣ����:=NULL;
            Select Nvl(Max(���),0)+1 Into v_�շ���� From ҩƷ�շ���¼ Where ���� IN(9,10) And ��¼״̬=1 And NO=NO_IN;
            Select Nvl(Max(���),0)+1 Into v_������� From ���˷��ü�¼ Where ��¼����=2 And ��¼״̬ IN(0,1) And NO=NO_IN;
            For r_Drug In c_Drug Loop
                --��ʼ��Ҫ�ջص�������(��������)
                IF v_ʣ���� IS NULL Then
                    v_ʣ����:=Round(�ջ���_IN*r_Drug.סԺ��װ,5);
                End IF;
                
                If v_ʣ����>=r_Drug.���� Then
                    v_��ǰ��:=r_Drug.����;
                Else
                    v_��ǰ��:=v_ʣ����;
                End IF;
                v_ʣ����:=v_ʣ����-v_��ǰ��;

                Select ���˷��ü�¼_ID.Nextval Into v_����ID From Dual;
                
                --ȷ������
                If Nvl(r_Drug.����,0)<>0 And r_Drug.����=0 Then
                    --ԭ����,�ֲ�����
                    v_����:=NULL;
                    v_����:=r_Drug.����;
                    v_Ч��:=r_Drug.Ч��;
                ElsIf Nvl(r_Drug.����,0)=0 And r_Drug.����=1 Then
                    --ԭ������,�ַ���
                    Select ҩƷ�շ���¼_ID.Nextval Into v_���� From Dual;
                    Select To_Char(Sysdate,'YYYYMMDD') Into v_���� From Dual;
                    If r_Drug.���Ч�� is Not Null Then
                        v_Ч��:=Trunc(Sysdate+r_Drug.���Ч��*30);
                    Else
                        v_Ч��:=NULL;
                    End IF;
                Else
                    v_����:=r_Drug.����;
                    v_����:=r_Drug.����;
                    v_Ч��:=r_Drug.Ч��;
                End IF;

                Insert Into ҩƷ�շ���¼(
                    ID,��¼״̬,����,NO,���,�ⷿID,�Է�����ID,������ID,���ϵ��,
                    ҩƷID,����,����,����,Ч��,����,��д����,ʵ������,���ۼ�,���۽��,
                    ժҪ,������,��������,����ID,����,Ƶ��,�÷�)
                Select
                    ҩƷ�շ���¼_ID.Nextval,1,����,NO_IN,v_�շ����,�ⷿID,�Է�����ID,
                    ������ID,-1,ҩƷID,v_����,����,v_����,v_Ч��,1,-1*v_��ǰ��,-1*v_��ǰ��,
                    ���ۼ�,Round(-1*v_��ǰ��*���ۼ�,v_Dec),'���ڷ����ջ�',v_��Ա����,v_Date,v_����ID,
                    ����,Ƶ��,�÷�
                From ҩƷ�շ���¼ Where ID=r_Drug.�շ�ID;

                --ҩƷ���
                Update ҩƷ���
                    Set ��������=Nvl(��������,0)-(-1*v_��ǰ��)
                Where �ⷿID=r_Drug.�ⷿID And ҩƷID=r_Drug.ҩƷID
                    And Nvl(����,0)=Nvl(v_����,0) And ����=1;
                IF SQL%RowCount=0 Then
                    Insert Into ҩƷ���(
                        �ⷿID,ҩƷID,����,��������,����,Ч��)
                    Values(
                        r_Drug.�ⷿID,r_Drug.ҩƷID,1,v_��ǰ��,v_����,v_Ч��);
                End IF;

                --δ��ҩƷ��¼
                Update δ��ҩƷ��¼
                    Set ����ID=r_Drug.����ID,
                        ��ҳID=r_Drug.��ҳID,
                        ����=r_Drug.����
                 Where ����=r_Drug.���� And NO=NO_IN And �ⷿID+0=r_Drug.�ⷿID;

                IF SQL%RowCount=0 Then
                    --ȡ������ȼ�
                    Begin
                        Select B.���ȼ� Into v_���ȼ� From ������Ϣ A,��� B
                         Where A.���=B.����(+) And A.����ID=r_Drug.����ID;
                    Exception
                        When Others Then Null;
                    End;

                    Insert Into δ��ҩƷ��¼(
                        ����,NO,����ID,��ҳID,����,���ȼ�,�Է�����ID,�ⷿID,��������,���շ�,��ӡ״̬)
                    Values(
                        r_Drug.����,NO_IN,r_Drug.����ID,r_Drug.��ҳID,r_Drug.����,v_���ȼ�,r_Drug.�Է�����ID,r_Drug.�ⷿID,v_Date,Decode(v_ҩƷ����,1,0,1),0);
                End IF;
                
                v_�շ����:=v_�շ����+1;

                --���˷��ü�¼
                -------------------------------------------------------------------------------------
                --��¼��ŷ�Χ�Դ�����ܱ�
                IF v_��ʼ��� IS NULL Then
                    v_��ʼ���:=v_�������;    
                End IF;
                v_�������:=v_�������;

                Insert Into ���˷��ü�¼(
                    ID,��¼����,NO,��¼״̬,���,��������,�۸񸸺�,�ಡ�˵�,�����־,����ID,��ҳID,
                    ��ʶ��,����,�Ա�,����,����,���˲���ID,���˿���ID,�ѱ�,�շ����,�շ�ϸĿID,���㵥λ,
                    ������Ŀ��,���մ���ID,����,����,�Ӱ��־,���ӱ�־,Ӥ����,������ĿID,�վݷ�Ŀ,��׼����,
                    Ӧ�ս��,ʵ�ս��,ͳ����,���ʷ���,��������ID,������,����ʱ��,�Ǽ�ʱ��,ִ�в���ID,
                    ִ��״̬,ҽ�����,������,����Ա���,����Ա����)
                Select 
                    v_����ID,2,NO_IN,Decode(v_ҩƷ����,1,0,1),v_�������,NULL,NULL,�ಡ�˵�,2,����ID,��ҳID,��ʶ��,����,�Ա�,����,
                    ����,���˲���ID,���˿���ID,�ѱ�,�շ����,�շ�ϸĿID,���㵥λ,������Ŀ��,���մ���ID,
                    1,-1*v_��ǰ��,�Ӱ��־,���ӱ�־,Ӥ����,������ĿID,�վݷ�Ŀ,��׼����,
                    Round(-1*v_��ǰ��*��׼����,v_Dec),Round(-1*v_��ǰ��*��׼����,v_Dec),NULL,1,��������ID,������,
                    v_Date,v_Date,ִ�в���ID,0,ҽ�����,v_��Ա����,Decode(v_ҩƷ����,1,NULL,v_��Ա���),Decode(v_ҩƷ����,1,NULL,v_��Ա����)
                From ���˷��ü�¼ Where ID=r_Drug.����ID;

                Begin
                    Select Round(B.Ӧ�ս��*A.ʵ�ձ���/100,v_Dec) Into v_ʵ�ս��
                    From �ѱ���ϸ A,���˷��ü�¼ B
                    Where B.ID=v_����ID And A.������ĿID=B.������ĿID And A.�ѱ�=B.�ѱ� 
                        And Abs(B.Ӧ�ս��) Between A.Ӧ�ն���ֵ And A.Ӧ�ն�βֵ;

                    Update ���˷��ü�¼ A Set ʵ�ս��=v_ʵ�ս�� Where ID=v_����ID;
                Exception
                    When Others Then NULL;
                End;
                
                v_�������:=v_�������+1;

                If v_ʣ����<=0 Then 
                    Exit;
                End IF;
            End Loop;

            If v_ʣ����<>0 Then
                --û���ջ���������,�շ���¼����������(���¼��ȫ������Ϊ��)
                NULL;
            End IF;
        Else
            --������ҩҽ��(������ҩ;��)
            -----------------------------------------------------------------------------------------------------
            Select Nvl(Max(���),0)+1 Into v_������� From ���˷��ü�¼ Where ��¼����=2 And ��¼״̬ IN(0,1) And NO=NO_IN;
            For r_Other In c_Other Loop
                --��¼��ŷ�Χ�Դ�����ܱ�
                IF v_��ʼ��� IS NULL Then
                    v_��ʼ���:=v_�������;    
                End IF;
                v_�������:=v_�������;
                
                --���˷��ü�¼:��������ջ����������ϴη�����,����ȷ
                Select ���˷��ü�¼_ID.Nextval Into v_����ID From Dual;
                Insert Into ���˷��ü�¼(
                    ID,��¼����,NO,��¼״̬,���,��������,�۸񸸺�,�ಡ�˵�,�����־,����ID,��ҳID,
                    ��ʶ��,����,�Ա�,����,����,���˲���ID,���˿���ID,�ѱ�,�շ����,�շ�ϸĿID,���㵥λ,
                    ������Ŀ��,���մ���ID,����,����,�Ӱ��־,���ӱ�־,Ӥ����,������ĿID,�վݷ�Ŀ,��׼����,
                    Ӧ�ս��,ʵ�ս��,ͳ����,���ʷ���,��������ID,������,����ʱ��,�Ǽ�ʱ��,ִ�в���ID,
                    ִ��״̬,ҽ�����,������,����Ա���,����Ա����)
                Select 
                    v_����ID,2,NO_IN,Decode(v_��������,1,0,1),v_�������,NULL,Decode(�۸񸸺�,NULL,NULL,v_�������+�۸񸸺�-���),�ಡ�˵�,2,
                    ����ID,��ҳID,��ʶ��,����,�Ա�,����,����,���˲���ID,���˿���ID,�ѱ�,�շ����,�շ�ϸĿID,���㵥λ,
                    ������Ŀ��,���մ���ID,1,-1*�ջ���_IN,�Ӱ��־,���ӱ�־,Ӥ����,������ĿID,�վݷ�Ŀ,��׼����,
                    Round(-1*�ջ���_IN*��׼����,v_Dec),Round(-1*�ջ���_IN*��׼����,v_Dec),NULL,1,��������ID,������,
                    v_Date,v_Date,ִ�в���ID,0,ҽ�����,v_��Ա����,Decode(v_��������,1,NULL,v_��Ա���),Decode(v_��������,1,NULL,v_��Ա����)
                From ���˷��ü�¼ Where ID=r_Other.����ID;
                
                Begin
                    Select Round(B.Ӧ�ս��*A.ʵ�ձ���/100,v_Dec) Into v_ʵ�ս��
                    From �ѱ���ϸ A,���˷��ü�¼ B
                    Where B.ID=v_����ID And A.������ĿID=B.������ĿID And A.�ѱ�=B.�ѱ� 
                        And Abs(B.Ӧ�ս��) Between A.Ӧ�ն���ֵ And A.Ӧ�ն�βֵ;

                    Update ���˷��ü�¼ A Set ʵ�ս��=v_ʵ�ս�� Where ID=v_����ID;
                Exception
                    When Others Then NULL;
                End;

                v_�������:=v_�������+1;
            End Loop;
        End IF;
    End IF;

    --������û��ܱ�
    -----------------------------------------------------------------------------------------------------
    If v_��ʼ��� IS Not NULL And v_������� IS Not NULL Then
        --���ͳһ���������ػ��ܱ�
        For r_Money IN c_Money(v_��ʼ���,v_�������) Loop
            --�������
            Update �������
                Set �������=Nvl(�������,0)+r_Money.ʵ�ս��
            Where ����ID=r_Money.����ID And ����=1;

            IF SQL%RowCount=0 Then
                Insert Into �������(
                    ����ID,����,�������,Ԥ�����)
                Values(
                    r_Money.����ID,1,r_Money.ʵ�ս��,0);
            End IF;

            --����δ�����
            Update ����δ�����
                Set ���=Nvl(���,0)+r_Money.ʵ�ս��
            Where ����ID=r_Money.����ID
                And ��ҳID=r_Money.��ҳID
                And Nvl(���˲���ID,0)=Nvl(r_Money.���˲���ID,0)
                And Nvl(���˿���ID,0)=Nvl(r_Money.���˿���ID,0)
                And ��������ID+0=r_Money.��������ID
                And ִ�в���ID+0=r_Money.ִ�в���ID
                And ������ĿID+0=r_Money.������ĿID
                And ��Դ;��+0=2;

            IF SQL%RowCount=0 Then
                Insert Into ����δ�����(
                    ����ID,��ҳID,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,��Դ;��,���)
                Values(
                    r_Money.����ID,r_Money.��ҳID,r_Money.���˲���ID,r_Money.���˿���ID,r_Money.��������ID,
                    r_Money.ִ�в���ID,r_Money.������ĿID,2,r_Money.ʵ�ս��);
            End IF;

            --���˷��û���
            Update ���˷��û���
                Set Ӧ�ս��=Nvl(Ӧ�ս��,0)+r_Money.Ӧ�ս��,
                    ʵ�ս��=Nvl(ʵ�ս��,0)+r_Money.ʵ�ս��
            Where ����=Trunc(v_Date)
                And Nvl(���˲���ID,0)=Nvl(r_Money.���˲���ID,0)
                And Nvl(���˿���ID,0)=Nvl(r_Money.���˿���ID,0)
                And ��������ID+0=r_Money.��������ID
                And ִ�в���ID+0=r_Money.ִ�в���ID
                And ������ĿID+0=r_Money.������ĿID
                And ��Դ;��=2
                And ���ʷ���=1;

            IF SQL%RowCount=0 Then
                Insert Into ���˷��û���(
                    ����,���˲���ID,���˿���ID,��������ID,ִ�в���ID,������ĿID,��Դ;��,���ʷ���,Ӧ�ս��,ʵ�ս��,���ʽ��)
                Values(
                    Trunc(v_Date),r_Money.���˲���ID,r_Money.���˿���ID,r_Money.��������ID,r_Money.ִ�в���ID,
                    r_Money.������ĿID,2,1,r_Money.Ӧ�ս��,r_Money.ʵ�ս��,0);
            End IF;
        End Loop;                            
    End IF;

    --����ҽ�����ϴ�ִ��ʱ��:��ҩ;���ȿ�����Ϊδ���Ͷ�û�����ջع��̡�
    -----------------------------------------------------------------------------------------------------
	Select Nvl(���ID,ID) Into v_��ID From ����ҽ����¼ Where ID=ҽ��ID_IN;
    Update ����ҽ����¼ Set �ϴ�ִ��ʱ��=�ϴ�ʱ��_IN Where ID=v_��ID Or ���ID=v_��ID;
Exception
    When Err_Custom Then Raise_Application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_����ҽ����¼_�ջ�;
/

CREATE OR REPLACE Procedure ZL_������ϼ�¼_Insert(
--���ܣ����벡����ϼ�¼
    ����ID_IN		������ϼ�¼.����ID%Type,
    ��ҳID_IN		������ϼ�¼.��ҳID%Type,
    ��¼��Դ_IN		������ϼ�¼.��¼��Դ%Type,
    ����ID_IN		������ϼ�¼.����ID%Type,
    �������_IN		������ϼ�¼.�������%Type,
    ����ID_IN		������ϼ�¼.����ID%Type,
    ���ID_IN		������ϼ�¼.���ID%Type,
    ֤��ID_IN		������ϼ�¼.֤��ID%Type,
    �������_IN		������ϼ�¼.�������%Type,
    ��Ժ���_IN		������ϼ�¼.��Ժ���%Type,
    �Ƿ�δ��_IN		������ϼ�¼.�Ƿ�δ��%Type,
    �Ƿ�����_IN		������ϼ�¼.�Ƿ�����%Type,
    ��¼����_IN		������ϼ�¼.��¼����%Type,
    ҽ��ID_IN		������ϼ�¼.ҽ��ID%Type:=NULL,
	��ϴ���_IN		������ϼ�¼.��ϴ���%Type:=1
) IS
    v_Temp            Varchar2(255);
    v_��Ա���        ��Ա��.���%Type;
    v_��Ա����        ��Ա��.����%Type;
Begin
    --��ǰ������Ա
    v_Temp:=zl_Identity;
    v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
    v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
    v_��Ա���:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
    v_��Ա����:=Substr(v_Temp,Instr(v_Temp,',')+1);
    
    Insert Into ������ϼ�¼(
        ID,����ID,��ҳID,��¼��Դ,����ID,�������,��ϴ���,����ID,���ID,֤��ID,�������,��Ժ���,�Ƿ�δ��,�Ƿ�����,��¼����,��¼��,ҽ��id)
    Values(
        ������ϼ�¼_ID.Nextval,����ID_IN,��ҳID_IN,��¼��Դ_IN,����ID_IN,�������_IN,��ϴ���_IN,����ID_IN,
        ���ID_IN,֤��ID_IN,�������_IN,��Ժ���_IN,�Ƿ�δ��_IN,�Ƿ�����_IN,��¼����_IN,v_��Ա����,ҽ��id_IN);
Exception
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_������ϼ�¼_Insert;
/

CREATE OR REPLACE Procedure zl_ҽ�����ݶ���_Delete
IS
Begin
	Delete From ҽ�����ݶ���;
Exception
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_ҽ�����ݶ���_Delete;
/

CREATE OR REPLACE Procedure zl_ҽ�����ݶ���_Insert(
	�������_IN		ҽ�����ݶ���.�������%Type,
	ҽ������_IN		ҽ�����ݶ���.ҽ������%Type
) IS
Begin
	Update ҽ�����ݶ��� Set ҽ������=ҽ������_IN Where �������=�������_IN;
	If SQL%RowCount=0 Then
		Insert Into ҽ�����ݶ���(
			�������,ҽ������)
		Values(
			�������_IN,ҽ������_IN);
	End IF;
Exception
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_ҽ�����ݶ���_Insert;
/

CREATE OR REPLACE Procedure zl_��Ա֤���¼_Insert(
	��ԱID_IN		��Ա֤���¼.��ԱID%Type,
	CertDN_IN		��Ա֤���¼.CertDN%Type,
	CertSN_IN		��Ա֤���¼.CertSN%Type,
	SignCert_IN		��Ա֤���¼.SignCert%Type,
	EncCert_IN		��Ա֤���¼.EncCert%Type
) IS
	v_����			��Ա��.����%Type;
    v_Error			Varchar2(255);
    Err_Custom		Exception;
Begin
	Begin
		Select A.���� Into v_���� From ��Ա�� A,��Ա֤���¼ B Where A.ID=B.��ԱID And B.CertSN=CertSN_IN;
	Exception
		When Others Then Null;
	End;
	If v_���� Is Not Null Then
		v_Error:='������֤���Ѿ�ע���"'||v_����||'"�������ظ�ע�ᡣ';
		Raise Err_Custom;
	End IF;

	Insert Into ��Ա֤���¼(
		ID,��ԱID,CertDN,CertSN,SignCert,EncCert,ע��ʱ��)
	Values(
		��Ա֤���¼_ID.Nextval,��ԱID_IN,CertDN_IN,CertSN_IN,SignCert_IN,EncCert_IN,Sysdate);
Exception
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_��Ա֤���¼_Insert;
/

CREATE OR REPLACE Procedure zl_��Ա֤���¼_Delete(
	֤��ID_IN		��Ա֤���¼.ID%Type
) IS
Begin
	Delete From ��Ա֤���¼ Where ID=֤��ID_IN;
Exception
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_��Ա֤���¼_Delete;
/

CREATE OR REPLACE Procedure zl_ҽ��ǩ����¼_Insert(
	ǩ��ID_IN		ҽ��ǩ����¼.ID%Type,
	ǩ������_IN		����ҽ��״̬.��������%Type,--��ӦΪ1-�¿�,4-����,8-ֹͣ
	ǩ������_IN		ҽ��ǩ����¼.ǩ������%Type,
	ǩ����Ϣ_IN		ҽ��ǩ����¼.ǩ����Ϣ%Type,
	֤��ID_IN		ҽ��ǩ����¼.֤��ID%Type,
	ҽ��IDs_IN		Varchar2 --����ǩ����ҽ��ID����,��ʽΪ'1,2,3,...'
) IS
	v_ҽ��IDs		Varchar2(4000);
	v_��ǰID		����ҽ����¼.ID%Type;

    v_Temp			Varchar2(255);
    v_��Ա����		���˷��ü�¼.����Ա����%Type;

    v_Error			Varchar2(255);
    Err_Custom		Exception;
Begin
    --��ǰ������Ա
    v_Temp:=zl_Identity;
    v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
    v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
    v_��Ա����:=Substr(v_Temp,Instr(v_Temp,',')+1);
	
	Insert Into ҽ��ǩ����¼(
		ID,ǩ������,ǩ����Ϣ,֤��ID,ǩ��ʱ��,ǩ����)
	Values(
		ǩ��ID_IN,ǩ������_IN,ǩ����Ϣ_IN,֤��ID_IN,Sysdate,v_��Ա����);
	
	--����ǩ����Ӧ��ҽ��
    v_ҽ��IDs:=ҽ��IDs_IN||',';
	While v_ҽ��IDs Is Not Null Loop
		v_��ǰID:=to_Number(Substr(v_ҽ��IDs,1,Instr(v_ҽ��IDs,',')-1));
		
		--��Ϊ�⼸�������������ظ�����˲���ʱ����ж�Ҳ���Բ�Ҫ��
		Update ����ҽ��״̬ 
			Set ǩ��ID=ǩ��ID_IN 
		Where ҽ��ID=v_��ǰID And ��������=ǩ������_IN And ǩ��ID Is Null
			And ����ʱ��=(Select Max(����ʱ��) From ����ҽ��״̬ Where ҽ��ID=v_��ǰID And ��������=ǩ������_IN);
		If SQL%RowCount=0 Then
			v_Error:='û���ҵ�Ҫǩ����ҽ����������Ч��ɵ���ǩ����';
			Raise Err_Custom;
		End IF;

		v_ҽ��IDs:=Substr(v_ҽ��IDs,Instr(v_ҽ��IDs,',')+1);
	End Loop;
Exception
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_ҽ��ǩ����¼_Insert;
/

CREATE OR REPLACE Procedure zl_ҽ��ǩ����¼_Delete(
	ǩ��ID_IN		ҽ��ǩ����¼.ID%Type
) IS
	v_Count			Number;
    v_Error			Varchar2(255);
    Err_Custom		Exception;
Begin
	--ȡ���¿�ҽ��ǩ���Ĳ������
	Select 
		Count(A.ID) Into v_Count
	From ����ҽ����¼ A,����ҽ��״̬ B 
	Where A.ҽ��״̬ Not IN(1,2) And A.ID=B.ҽ��ID 
		And B.��������=1 And B.ǩ��ID=ǩ��ID_IN;
	If Nvl(v_Count,0)>0 Then
		v_Error:='���ҽ���Ѿ�У�Ի��ͣ�����ȡ������ǩ����';
		Raise Err_Custom;
	End IF;
	
	Update ����ҽ��״̬ Set ǩ��ID=NULL Where ǩ��ID=ǩ��ID_IN;
	Delete From ҽ��ǩ����¼ Where ID=ǩ��ID_IN;
	If SQL%RowCount=0 Then
		v_Error:='û���ҵ�ҽ����ǩ����¼���޷�ȡ������ǩ����';
		Raise Err_Custom;
	End IF;
Exception
	When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End zl_ҽ��ǩ����¼_Delete;
/

CREATE OR REPLACE Procedure ZL_����ҽ����¼_Insert(
--���ܣ�ҽ����ʿ�¿�,��¼ҽ��ʱ�²�����ҽ����¼�������������סԺ��
    ID_IN				����ҽ����¼.ID%TYPE,
    ���ID_IN           ����ҽ����¼.���ID%TYPE,
    ���_IN				����ҽ����¼.���%TYPE,
    ������Դ_IN         ����ҽ����¼.������Դ%TYPE,
    ����ID_IN           ����ҽ����¼.����ID%TYPE,
    ��ҳID_IN           ����ҽ����¼.��ҳID%TYPE,
    Ӥ��_IN             ����ҽ����¼.Ӥ��%TYPE,
    ҽ��״̬_IN         ����ҽ����¼.ҽ��״̬%TYPE,
    ҽ����Ч_IN         ����ҽ����¼.ҽ����Ч%TYPE,
    �������_IN         ����ҽ����¼.�������%TYPE,
    ������ĿID_IN       ����ҽ����¼.������ĿID%TYPE,
    �շ�ϸĿID_IN       ����ҽ����¼.�շ�ϸĿID%TYPE,
	����_IN				����ҽ����¼.����%TYPE,
    ��������_IN         ����ҽ����¼.��������%TYPE,
    �ܸ�����_IN         ����ҽ����¼.�ܸ�����%TYPE,
    ҽ������_IN         ����ҽ����¼.ҽ������%TYPE,
    ҽ������_IN         ����ҽ����¼.ҽ������%TYPE,
    �걾��λ_IN         ����ҽ����¼.�걾��λ%TYPE,
    ִ��Ƶ��_IN         ����ҽ����¼.ִ��Ƶ��%TYPE,
    Ƶ�ʴ���_IN         ����ҽ����¼.Ƶ�ʴ���%TYPE,
    Ƶ�ʼ��_IN         ����ҽ����¼.Ƶ�ʼ��%TYPE,
    �����λ_IN         ����ҽ����¼.�����λ%TYPE,
    ִ��ʱ�䷽��_IN		����ҽ����¼.ִ��ʱ�䷽��%TYPE,
    �Ƽ�����_IN         ����ҽ����¼.�Ƽ�����%TYPE,
    ִ�п���ID_IN       ����ҽ����¼.ִ�п���ID%TYPE,
    ִ������_IN         ����ҽ����¼.ִ������%TYPE,
    ������־_IN         ����ҽ����¼.������־%TYPE,
    ��ʼִ��ʱ��_IN     ����ҽ����¼.��ʼִ��ʱ��%TYPE,
    ִ����ֹʱ��_IN     ����ҽ����¼.ִ����ֹʱ��%TYPE,
    ���˿���ID_IN       ����ҽ����¼.���˿���ID%TYPE,
    ��������ID_IN       ����ҽ����¼.��������ID%TYPE,
    ����ҽ��_IN         ����ҽ����¼.����ҽ��%TYPE,
    ����ʱ��_IN         ����ҽ����¼.����ʱ��%TYPE,
    �Һŵ�_IN           ����ҽ����¼.�Һŵ�%TYPE:=NULL,
	ǰ��ID_IN			����ҽ����¼.ǰ��ID%Type:=NULL
) IS
    v_Temp			Varchar2(255);
    v_��Ա���		���˷��ü�¼.����Ա���%Type;
    v_��Ա����		���˷��ü�¼.����Ա����%Type;
	
	v_����			����ҽ����¼.����%Type;
	v_�Ա�			����ҽ����¼.�Ա�%Type;
	v_����			����ҽ����¼.����%Type;

    v_Error			Varchar2(255);
    Err_Custom		Exception;
Begin
    --��ǰ������Ա
    v_Temp:=zl_Identity;
    v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
    v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
    v_��Ա���:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
    v_��Ա����:=Substr(v_Temp,Instr(v_Temp,',')+1);
	
	Select ����,�Ա�,���� Into v_����,v_�Ա�,v_���� From ������Ϣ Where ����ID=����ID_IN;

    --����ҽ����¼
    Insert Into ����ҽ����¼(
        ID,���ID,���,������Դ,����ID,��ҳID,����,�Ա�,����,Ӥ��,ҽ��״̬,ҽ����Ч,�������,������ĿID,�շ�ϸĿID,
        ����,��������,�ܸ�����,ҽ������,ҽ������,�걾��λ,ִ��Ƶ��,Ƶ�ʴ���,Ƶ�ʼ��,�����λ,ִ��ʱ�䷽��,
        �Ƽ�����,ִ�п���ID,ִ������,������־,��ʼִ��ʱ��,ִ����ֹʱ��,���˿���ID,��������ID,����ҽ��,
		����ʱ��,�Һŵ�,ǰ��ID)
    VALUES(
        ID_IN,���ID_IN,���_IN,������Դ_IN,����ID_IN,��ҳID_IN,v_����,v_�Ա�,v_����,Ӥ��_IN,ҽ��״̬_IN,ҽ����Ч_IN,
        �������_IN,������ĿID_IN,�շ�ϸĿID_IN,����_IN,��������_IN,�ܸ�����_IN,ҽ������_IN,ҽ������_IN,
        �걾��λ_IN,ִ��Ƶ��_IN,Ƶ�ʴ���_IN,Ƶ�ʼ��_IN,�����λ_IN,ִ��ʱ�䷽��_IN,�Ƽ�����_IN,
        ִ�п���ID_IN,ִ������_IN,������־_IN,��ʼִ��ʱ��_IN,ִ����ֹʱ��_IN,
        ���˿���ID_IN,��������ID_IN,����ҽ��_IN,����ʱ��_IN,�Һŵ�_IN,ǰ��ID_IN);

    --����ҽ��״̬
	Delete From ����ҽ��״̬ Where ҽ��ID=ID_IN And ��������=1;
	If SQL%RowCount<>0 Then
		v_Error:='��ͬID���¿�ҽ���Ѿ����ڡ�';
		Raise Err_Custom;
	End IF;
	--��Ϊ����ͬʱ���¿�->�Զ�У��->�����Զ�ֹͣ,��˷ֱ�-2,-1��
    Insert Into ����ҽ��״̬(
        ҽ��ID,��������,������Ա,����ʱ��)
    Values(
        ID_IN,1,v_��Ա����,Sysdate-2/60/60/24);
Exception
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_����ҽ����¼_Insert;
/

CREATE OR REPLACE Procedure ZL_����ҽ����¼_Update(
--���ܣ���ҽ����ʿ�޸��˲������ݵ�ҽ����¼�������������סԺ��
--˵����Updateʱ֮�����漰������ĿID,�Ƽ����Ա仯,����Ϊ��ҩ;��,�÷��ı仯
--      Updateʱ֮�����漰��Ч�仯,����Ϊ����¼��ҽ��������ı���Ч
    ID_IN               ����ҽ����¼.ID%TYPE,
    ���ID_IN           ����ҽ����¼.���ID%TYPE,
    ���_IN             ����ҽ����¼.���%TYPE,
    ҽ��״̬_IN         ����ҽ����¼.ҽ��״̬%TYPE,
	ҽ����Ч_IN			����ҽ����¼.ҽ����Ч%TYPE,
    ������ĿID_IN       ����ҽ����¼.������ĿID%TYPE,
	����_IN				����ҽ����¼.����%TYPE,
    ��������_IN         ����ҽ����¼.��������%TYPE,
    �ܸ�����_IN         ����ҽ����¼.�ܸ�����%TYPE,
    ҽ������_IN         ����ҽ����¼.ҽ������%TYPE,
    ҽ������_IN         ����ҽ����¼.ҽ������%TYPE,
    �걾��λ_IN         ����ҽ����¼.�걾��λ%TYPE,
    ִ��Ƶ��_IN         ����ҽ����¼.ִ��Ƶ��%TYPE,
    Ƶ�ʴ���_IN         ����ҽ����¼.Ƶ�ʴ���%TYPE,
    Ƶ�ʼ��_IN         ����ҽ����¼.Ƶ�ʼ��%TYPE,
    �����λ_IN         ����ҽ����¼.�����λ%TYPE,
    ִ��ʱ�䷽��_IN     ����ҽ����¼.ִ��ʱ�䷽��%TYPE,
    �Ƽ�����_IN         ����ҽ����¼.�Ƽ�����%TYPE,
    ִ�п���ID_IN       ����ҽ����¼.ִ�п���ID%TYPE,
    ִ������_IN         ����ҽ����¼.ִ������%TYPE,
    ������־_IN         ����ҽ����¼.������־%TYPE,
    ��ʼִ��ʱ��_IN     ����ҽ����¼.��ʼִ��ʱ��%TYPE,
    ִ����ֹʱ��_IN     ����ҽ����¼.ִ����ֹʱ��%TYPE,
    ���˿���ID_IN       ����ҽ����¼.���˿���ID%TYPE,
    ��������ID_IN       ����ҽ����¼.��������ID%TYPE,
    ����ҽ��_IN         ����ҽ����¼.����ҽ��%TYPE,
    ����ʱ��_IN         ����ҽ����¼.����ʱ��%TYPE
) IS
    v_Count			Number;

    v_Temp			Varchar2(255);
    v_��Ա���		���˷��ü�¼.����Ա���%Type;
    v_��Ա����		���˷��ü�¼.����Ա����%Type;

    v_Error        Varchar2(255);
    Err_Custom    Exception;
Begin
    --����ҽ��״̬:��������
    Begin
        Select ҽ��״̬ Into v_Count From ����ҽ����¼ Where ID=ID_IN;
    Exception
        When Others Then
        Begin
            v_Error:='ҽ��"'||ҽ������_IN||'"�Ѿ�������,�����ѱ�������ɾ����';
            Raise Err_Custom;
        End;
    End;
    If v_Count Not IN(1,2) Then
        v_Error:='ҽ��"'||ҽ������_IN||'"�Ѿ�У�Ի���,�������޸ġ�';
        Raise Err_Custom;
    End IF;

	Select Count(*) Into v_Count From ����ҽ��״̬ Where ҽ��ID=ID_IN And ��������=1 And ǩ��ID Is Not Null;
	If Nvl(v_Count,0)>0 Then
        v_Error:='ҽ��"'||ҽ������_IN||'"�Ѿ�����ǩ��,�������޸ġ�';
        Raise Err_Custom;
	End IF;

    --��ǰ������Ա    
    v_Temp:=zl_Identity;
    v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
    v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
    v_��Ա���:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
    v_��Ա����:=Substr(v_Temp,Instr(v_Temp,',')+1);

    --����ҽ����¼
    Update ����ҽ����¼ 
        Set ���ID=���ID_IN,--����һ����ҩ���������ü�鲿λ����������ID�仯
            ���=���_IN,
            ҽ��״̬=ҽ��״̬_IN,--!��Ϊֻ���޸�δУ��ҽ��������Ӧ��Ϊ�¿���У�����ʵ�ҽ���޸ĺ�Ϊ�¿�
			ҽ����Ч=ҽ����Ч_IN,
            ������ĿID=������ĿID_IN,
			����=����_IN,
            ��������=��������_IN,
            �ܸ�����=�ܸ�����_IN,
            ҽ������=ҽ������_IN,
            ҽ������=ҽ������_IN,
            �걾��λ=�걾��λ_IN,
            ִ��Ƶ��=ִ��Ƶ��_IN,
            Ƶ�ʴ���=Ƶ�ʴ���_IN,
            Ƶ�ʼ��=Ƶ�ʼ��_IN,
            �����λ=�����λ_IN,
            ִ��ʱ�䷽��=ִ��ʱ�䷽��_IN,
            �Ƽ�����=�Ƽ�����_IN,
            ִ�п���ID=ִ�п���ID_IN,
            ִ������=ִ������_IN,--ҩƷ�����⹺ҩ,��Ժ��ҩ�ĵ���ʱ�ᷢ���仯
            ������־=������־_IN,
            ��ʼִ��ʱ��=��ʼִ��ʱ��_IN,
            ִ����ֹʱ��=ִ����ֹʱ��_IN,--!��������ֹʱ������޸�,����Ӧ��Ϊ��
            ���˿���ID=���˿���ID_IN,--�޸�ʱ����Ϊ���˵ĵ�ǰ����
            ��������ID=��������ID_IN,--�޸ĺ����ݵ�ǰ���ұ仯
            ����ҽ��=����ҽ��_IN,--��ʿ��ҽ��ʱ���Ը���
            ����ʱ��=����ʱ��_IN--��¼�Ŀ����޸�
    Where ID=ID_IN;
    
    --����ҽ��״̬:����ҽ���¿�����
    Update ����ҽ��״̬
        Set ������Ա=v_��Ա����,
            ����ʱ��=Sysdate
    Where ҽ��ID=ID_IN And ��������=1;--�¿�����ʼ����,У�����ʱ�����Ϊ��ʷ��¼
Exception
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_����ҽ����¼_Update;
/

CREATE OR REPLACE Procedure ZL_����ҽ����¼_Delete(
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
    ҽ��ID_IN		����ҽ����¼.ID%TYPE,
    ɾ���_IN       Number:=0
) IS
	Cursor c_Case is
		Select ����ID From ����ҽ����¼ Where ����ID IS Not NULL And (ID=ҽ��ID_IN Or ���ID=ҽ��ID_IN); 
	r_Case	c_Case%RowType;

    v_״̬			����ҽ����¼.ҽ��״̬%Type;
    v_���ID		����ҽ����¼.���ID%Type;
    v_����ID		����ҽ����¼.����ID%Type;
    v_�Һŵ�		����ҽ����¼.�Һŵ�%Type;
    v_��ҳID		����ҽ����¼.��ҳID%Type;
    v_Ӥ��			����ҽ����¼.Ӥ��%Type;
    v_���			����ҽ����¼.���%Type;
    v_����			����ҽ����¼.ҽ������%Type;
    v_Count			Number(5);

    v_Error			Varchar2(255);
    Err_Custom		Exception;
Begin
    --���ҽ��״̬:��������
    Begin
        Select ����ID,�Һŵ�,��ҳID,Ӥ��,ҽ��״̬,���ID,ҽ������
            Into v_����ID,v_�Һŵ�,v_��ҳID,v_Ӥ��,v_״̬,v_���ID,v_����
        From ����ҽ����¼ Where ID=ҽ��ID_IN;
    Exception
        When Others Then
        Begin
            v_Error:='δ����Ҫɾ����ҽ����¼�������ѱ�������ɾ����';
            Raise Err_Custom;
        End;
    End;
    If v_�Һŵ� IS NULL Then
        IF Not v_״̬ IN(1,2) Then
            v_Error:='ҽ��"'||v_����||'"�Ѿ���У�ԣ�������ɾ����';
            Raise Err_Custom;
        End IF;
    Else
        IF v_״̬<>1 Then
            v_Error:='ҽ��"'||v_����||'"�Ѿ������ͻ����ϣ�����ɾ����';
            Raise Err_Custom;
        End IF;
    End IF;

	Select Count(*) Into v_Count From ����ҽ��״̬ Where ҽ��ID=ҽ��ID_IN And ��������=1 And ǩ��ID Is Not Null;
	If Nvl(v_Count,0)>0 Then
        v_Error:='ҽ��"'||v_����||'"�Ѿ�����ǩ��,����ɾ����';
        Raise Err_Custom;
	End IF;

	IF Nvl(ɾ���_IN,0)=0 then
		Begin
			Select ����ID Into v_Count From ����ҽ����¼ Where ID=ҽ��ID_IN;
		Exception
			When Others Then v_Count:=NULL;
		End;

		--ɾ��ҽ��
		Delete From ����ҽ����¼ Where ID=ҽ��ID_IN;
		
		--ɾ����Ӧ�����뵥
		IF v_Count IS Not NULL Then
			Delete From ���˲�����¼ Where ID=v_Count;
		End IF;
    Else
		If v_���ID IS NULL Then
            --������,����������,��ҩ�䷽,�������,�Լ�����ҽ��
            Select Max(���),Count(*) Into v_���,v_Count From ����ҽ����¼ Where ID=ҽ��ID_IN Or ���ID=ҽ��ID_IN;
			
			Open c_Case;--�����ȴ�

			--ɾ��ҽ��
			Delete From ����ҽ����¼ Where ID=ҽ��ID_IN Or ���ID=ҽ��ID_IN;

			--ɾ����Ӧ�����뵥
			Fetch c_Case Into r_Case;
			While c_Case%Found Loop
				Delete From ���˲�����¼ Where ID=r_Case.����ID;
				Fetch c_Case Into r_Case;
			End Loop;
			Close c_Case;
        Else
            --��ҩһ����ҩ�����(������)
            --���ж��Ƿ�һ����ҩ
            Select Count(*) Into v_Count From ����ҽ����¼ Where ���ID=v_���ID;
            
            If v_Count=1 Then
                --������ҩ:ͬʱɾ�����ҩ;��
                Select Max(���),Count(*) Into v_���,v_Count From ����ҽ����¼ Where ID=ҽ��ID_IN Or ID=v_���ID;
                Delete From ����ҽ����¼ Where ID=ҽ��ID_IN Or ID=v_���ID;
            Else
                --һ����ҩ:ֻɾ����ǰҩƷ
                v_Count:=1;
                Select ��� Into v_��� From ����ҽ����¼ Where ID=ҽ��ID_IN;
                Delete From ����ҽ����¼ Where ID=ҽ��ID_IN;
            End If;
        End if;

        --�������
        Update ����ҽ����¼ 
            Set ���=���-v_Count
        Where ����ID=v_����ID 
            And Nvl(��ҳID,0)=Nvl(v_��ҳID,0) 
            And Nvl(�Һŵ�,'��')=Nvl(v_�Һŵ�,'��')
            And Nvl(Ӥ��,0)=Nvl(v_Ӥ��,0) 
            And ���>v_���;
    End if;
Exception
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_����ҽ����¼_Delete;
/

CREATE OR REPLACE Procedure ZL_����ҽ����¼_У��(
--���ܣ�У��ָ����ҽ��
--������ҽ��ID_IN=Nvl(���ID,ID)
--      ״̬_IN=У��ͨ��3��У������2
--      �Զ�У��_IN=����֮������Զ�У��,�Զ���д�Ƽ�����
--˵����һ��ҽ��ֻ�ܵ���һ��,����ͬʱ��ɴ���һ��ҽ����У��
    ҽ��ID_IN		����ҽ����¼.ID%TYPE,
    ״̬_IN			����ҽ����¼.ҽ��״̬%TYPE,
    У��ʱ��_IN		����ҽ��״̬.����ʱ��%TYPE,
	�Զ�У��_IN		Number:=Null
) IS
    --����ҽ�����
    v_״̬			����ҽ����¼.ҽ��״̬%Type;
	v_��Ч			����ҽ����¼.ҽ����Ч%Type;
    v_����ID        ����ҽ����¼.����ID%Type;
    v_��ҳID        ����ҽ����¼.��ҳID%Type;
    v_Ӥ��			����ҽ����¼.Ӥ��%Type;
    v_ҽ������		����ҽ����¼.ҽ������%Type;
	v_����ʱ��		����ҽ����¼.����ʱ��%Type;
	v_��ʼʱ��		����ҽ����¼.��ʼִ��ʱ��%Type;
	v_����ҽ��		����ҽ����¼.����ҽ��%Type;
	v_ǰ��ID		����ҽ����¼.ǰ��ID%Type;

    --���ڱ������ȼ�
    v_�������		����ҽ����¼.�������%TYPE;
    v_������ĿID    ����ҽ����¼.������ĿID%TYPE;
    v_��������      ������ĿĿ¼.��������%TYPE;
    v_����ȼ�ID    ������ҳ.����ȼ�ID%TYPE;

    --�����Ŀͬһ�Զ�ֹͣ���������Ŀ:����Ӧ�ö��ǳ���(������ǰҽ��),����Ӧ�Ѽ�顣
    --ע��Ӧ��Ӥ������,ͬʱҲӦֹͣ����ǰҽ�����������ͬ������Ŀ��ҽ����
    Cursor c_Exclude IS
        Select Distinct B.ID AS ҽ��ID,B.��ʼִ��ʱ��,B.ִ����ֹʱ��,B.�ϴ�ִ��ʱ��,B.����ҽ��,
            B.ִ��ʱ�䷽��,B.Ƶ�ʼ��,B.Ƶ�ʴ���,B.�����λ
        From ���ƻ�����Ŀ A,����ҽ����¼ B
        Where A.����=3 And A.��ĿID=B.������ĿID And B.ID<>ҽ��ID_IN
            And Nvl(B.ҽ����Ч,0)=0 And B.ҽ��״̬ IN(3,5,6,7)
            And B.����ID=v_����ID And Nvl(B.��ҳID,0)=Nvl(v_��ҳID,0) And Nvl(B.Ӥ��,0)=Nvl(v_Ӥ��,0)
            And A.���� IN(Select Distinct ���� From ���ƻ�����Ŀ Where ����=3 And ��ĿID=v_������ĿID)
            Order by B.ID;
    v_��ֹʱ�� ����ҽ����¼.ִ����ֹʱ��%TYPE;

    Cursor c_Nurse IS
        Select A.ID AS ҽ��ID,A.��ʼִ��ʱ��,A.ִ����ֹʱ��,A.�ϴ�ִ��ʱ��,A.����ҽ��
        From ����ҽ����¼ A,������ĿĿ¼ B
        Where A.������ĿID=B.ID And A.�������='H' And B.��������='1'
            And A.����ID=v_����ID And Nvl(A.��ҳID,0)=Nvl(v_��ҳID,0) And Nvl(A.Ӥ��,0)=Nvl(v_Ӥ��,0)
            And Nvl(A.ҽ����Ч,0)=0 And A.ҽ��״̬ IN(3,5,6,7) And A.ID<>ҽ��ID_IN;

    --��������(Ӥ��)������δͣ����(���䷽����)
    Cursor c_NeedStop(
        v_����ID    ����ҽ����¼.����ID%Type,
        v_��ҳID    ����ҽ����¼.��ҳID%Type,
        v_Ӥ��		����ҽ����¼.Ӥ��%Type,
        v_StopTime	Date) is
        Select ID From ����ҽ����¼ 
        Where ����ID=v_����ID And ��ҳID=v_��ҳID And Nvl(Ӥ��,0)=Nvl(v_Ӥ��,0) 
            And Nvl(ҽ����Ч,0)=0 And ҽ��״̬ Not IN(1,2,4,8,9)
            And ��ʼִ��ʱ��<v_StopTime
        Order BY ���;
	--��������(Ӥ��)����ͣ��δȷ�ϵĳ���,��ִֹ��ʱ����ָ��ʱ��֮��
    Cursor c_HaveStop(
        v_����ID    ����ҽ����¼.����ID%Type,
        v_��ҳID    ����ҽ����¼.��ҳID%Type,
        v_Ӥ��		����ҽ����¼.Ӥ��%Type,
        v_StopTime	Date) is
        Select ID From ����ҽ����¼ 
        Where ����ID=v_����ID And ��ҳID=v_��ҳID And Nvl(Ӥ��,0)=Nvl(v_Ӥ��,0) 
            And Nvl(ҽ����Ч,0)=0 And ҽ��״̬=8 And ִ����ֹʱ��>v_StopTime
			And ��ʼִ��ʱ��<v_StopTime
        Order BY ���;
	
	--ȡһ��ҽ���ļƼ�����
	Cursor c_Price Is
		Select A.ID,B.�շ���ĿID,B.�շ�����,B.������Ŀ,
			Sum(Decode(Nvl(C.�Ƿ���,0),1,D.ԭ��,Null)) as ����
		From ����ҽ����¼ A,�����շѹ�ϵ B,�շ���ĿĿ¼ C,�շѼ�Ŀ D
		Where A.������ĿID=B.������ĿID And B.�շ���ĿID=C.ID
			And C.ID=D.�շ�ϸĿID And A.������� Not IN('5','6','7') 
			And Nvl(A.�Ƽ�����,0)=0 And Nvl(A.ִ������,0) Not IN(0,5)
			And C.������� IN(1,3) And (C.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or C.����ʱ�� is NULL)
			And Sysdate Between D.ִ������ And Nvl(D.��ֹ����,To_Date('3000-01-01','YYYY-MM-DD'))
			And Nvl(B.�շ�����,0)<>0 And Not(Nvl(C.�Ƿ���,0)=1 And Nvl(D.ԭ��,0)=0)
			And (A.ID=ҽ��ID_IN Or A.���ID=ҽ��ID_IN)
		Group by A.ID,B.�շ���ĿID,B.�շ�����,B.������Ŀ;

    --������ʱ����
	v_����ֵ		ϵͳ������.����ֵ%Type;
	v_Count			Number;
    v_Date			Date;
    v_Temp			Varchar2(255);
    v_��Ա���		���˷��ü�¼.����Ա���%Type;
    v_��Ա����		���˷��ü�¼.����Ա����%Type;

    v_Error			Varchar2(255);
    Err_Custom		Exception;

	Function GetAdviceText(v_ҽ��ID ����ҽ����¼.ID%Type)
		Return Varchar2 Is
		v_Text	����ҽ����¼.ҽ������%Type;
		v_���	����ҽ����¼.�������%Type;
		v_�䷽	Number;
	Begin
		Select �������,ҽ������ Into v_���,v_Text From ����ҽ����¼ Where ID=v_ҽ��ID;
		If v_���='E' Then
			--��ҩ���г�ҩ��ҽ������
			Begin
				Select �������,Decode(�������,'7',v_Text,ҽ������) 
					Into v_���,v_Text 
				From ����ҽ����¼ Where ���ID=v_ҽ��ID And ������� IN('5','6','7') And Rownum=1;
			Exception
				When Others Then Null;
			End;
			If v_���='7' Then
				v_�䷽:=1;
			End IF;
		End IF;
		If Length(v_Text)>30 Then 
			v_Text:=Substr(v_Text,1,30)||'...';
		End IF;
		If Length(v_Text)>20 Then 
			v_Text:='"'||v_Text||'"'||CHR(13)||CHR(10);
		Else
			v_Text:='"'||v_Text||'"';
		End IF;
		If v_�䷽=1 Then
			v_Text:='��ҩ�䷽'||v_Text;
		End IF;
		Return(v_Text);
	End;
Begin
    --���ҽ��״̬�Ƿ���ȷ:��������
	Begin
		Select A.ҽ����Ч,A.ҽ��״̬,A.����ʱ��,A.����ҽ��,A.��ʼִ��ʱ��,A.����ID,A.��ҳID,A.Ӥ��,A.ҽ������,A.�������,A.������ĿID,A.ǰ��ID,Nvl(B.��������,'0')
			Into v_��Ч,v_״̬,v_����ʱ��,v_����ҽ��,v_��ʼʱ��,v_����ID,v_��ҳID,v_Ӥ��,v_ҽ������,v_�������,v_������ĿID,v_ǰ��ID,v_��������
		From ����ҽ����¼ A,������ĿĿ¼ B
		Where A.������ĿID=B.ID(+) And A.ID=ҽ��ID_IN;
	Exception
		When Others Then
		Begin
			v_Error:='ҽ���ѱ�ɾ�������ܽ���У�ԡ�'||CHR(13)||CHR(10)||'������ǲ�����������ģ������¶�ȡУ�����ݡ�';
			Raise Err_Custom;
		End;
	End;
	IF v_״̬<>1 Then
		v_Error:='ҽ��"'||GetAdviceText(ҽ��ID_IN)||'"�����¿���ҽ��������ͨ��У�ԡ�'||CHR(13)||CHR(10)||'������ǲ�����������ģ������¶�ȡУ�����ݡ�';
		Raise Err_Custom;
	End IF;
	--�ٴμ��У��ʱ�����Ч��:��������
	If To_Char(v_����ʱ��,'YYYY-MM-DD HH24:MI') <= To_Char(v_��ʼʱ��,'YYYY-MM-DD HH24:MI') Then
		If To_Char(У��ʱ��_IN,'YYYY-MM-DD HH24:MI') < To_Char(v_����ʱ��,'YYYY-MM-DD HH24:MI') Then
			v_Error:='ҽ��"'||GetAdviceText(ҽ��ID_IN)||'"��У��ʱ�䲻��С�ڿ���ʱ�� '||To_Char(v_����ʱ��,'YYYY-MM-DD HH24:MI')||'��'||CHR(13)||CHR(10)||'������ǲ�����������ģ������¶�ȡУ�����ݡ�';
			Raise Err_Custom;
		End If;
	Else
		If To_Char(У��ʱ��_IN,'YYYY-MM-DD HH24:MI') < To_Char(v_��ʼʱ��,'YYYY-MM-DD HH24:MI') Then
			v_Error:='ҽ��"'||GetAdviceText(ҽ��ID_IN)||'"��У��ʱ�䲻��С�ڿ�ʼִ��ʱ�� '||To_Char(v_��ʼʱ��,'YYYY-MM-DD HH24:MI')||'��'||CHR(13)||CHR(10)||'������ǲ�����������ģ������¶�ȡУ�����ݡ�';
			Raise Err_Custom;
		End If;
	End If;
	
	
	--���Ҫ��ǩ�������У��ʱ�Ƿ���ǩ��(����ȡ��ǩ��)
	If ״̬_IN=3 Then
		Begin
			Select ����ֵ Into v_����ֵ From ϵͳ������ Where ������=25;
		Exception
			When Others Then Null;
		End;
		If Nvl(v_����ֵ,'0')<>'0' Then
			v_����ֵ:=Null;
			Begin
				Select ����ֵ Into v_����ֵ From ϵͳ������ Where ������=26;
			Exception
				When Others Then Null;
			End;
			If Nvl(Substr(v_����ֵ,2,1),'0')='1' And v_ǰ��ID Is Null 
				Or Nvl(Substr(v_����ֵ,3,1),'0')='1' And v_ǰ��ID Is Not Null Then
				Select Count(*) Into v_Count From ����ҽ��״̬ Where ��������=1 And ǩ��ID Is Not Null And ҽ��ID=ҽ��ID_IN;
				If Nvl(v_Count,0)=0 Then
					v_Error:='ҽ��"'||GetAdviceText(ҽ��ID_IN)||'"��û�е���ǩ��������ͨ��У�ԡ�';
					Raise Err_Custom;
				End If;
			End If;
		End IF;
	End IF;

    --��ǰ������Ա    
    v_Temp:=zl_Identity;
    v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
    v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
    v_��Ա���:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
    v_��Ա����:=Substr(v_Temp,Instr(v_Temp,',')+1);
    
    --��Ϊ����ͬʱ���¿�->�Զ�У��->�����Զ�ֹͣ,��˷ֱ�-2,-1��
    Select Sysdate-1/60/60/24 Into v_Date From Dual;

	Update ����ҽ����¼ 
        Set ҽ��״̬=״̬_IN,
            У�Ի�ʿ=v_��Ա����,
            У��ʱ��=У��ʱ��_IN
    Where ID=ҽ��ID_IN Or ���ID=ҽ��ID_IN;

    Insert Into ����ҽ��״̬(
        ҽ��ID,��������,������Ա,����ʱ��)
    Select 
        ID,״̬_IN,v_��Ա����,v_Date
    From ����ҽ����¼
    Where ID=ҽ��ID_IN Or ���ID=ҽ��ID_IN;
    
    --У��ͨ��ʱ����������
    If ״̬_IN=3 Then
		--�Զ�У��ʱ���Զ���дȱʡ�ļƼ�����
		If Nvl(�Զ�У��_IN,0)=1 Then
			--1.��۵ļƼ���Ŀ,�������޼۲�Ϊ0,��ȱʡΪ����޼�,���򲻼���;�����ֹ��Ƽ�.
			--2.���ڷ�ҩ��ҩƷ����������δ��ִ�п���,����ʱ��ȡȱʡ��,�����ֹ����á�
			For r_Price In c_Price Loop
				Insert Into ����ҽ���Ƽ�(
					ҽ��ID,�շ�ϸĿID,����,����,����,ִ�п���ID)
				Values(
					r_Price.ID,r_Price.�շ���ĿID,r_Price.�շ�����,r_Price.����,r_Price.������Ŀ,NULL);
			End Loop;
		End IF;
		
		--����¼�������ҽ�����Ϊֹͣ
		If Nvl(v_��Ч,0)=1 And v_������ĿID Is Null Then
            Update ����ҽ����¼ 
                Set ҽ��״̬=8,
                    ͣ��ʱ��=У��ʱ��_IN,
                    ͣ��ҽ��=v_����ҽ��
            Where ID=ҽ��ID_IN Or ���ID=ҽ��ID_IN;
            
            Insert Into ����ҽ��״̬(
                ҽ��ID,��������,������Ա,����ʱ��) 
            Select 
                ID,8,v_��Ա����,У��ʱ��_IN 
            From ����ҽ����¼ 
            Where ID=ҽ��ID_IN Or ���ID=ҽ��ID_IN;
		End IF;

        --��ͬһ�Զ�ֹͣ�������еĲ�������ҽ��ֹͣ(�����δֹͣ)
        For r_Exclude In c_Exclude Loop
			Select Decode(Sign(r_Exclude.��ʼִ��ʱ��-v_��ʼʱ��),1,r_Exclude.��ʼִ��ʱ��,v_��ʼʱ��) Into v_��ֹʱ�� From Dual;
            ZL_����ҽ����¼_ֹͣ(r_Exclude.ҽ��ID,v_��ֹʱ��,v_����ҽ��);
        End Loop;

        --��һЩ����ҽ���Ĵ���
        If v_�������='H' And v_��������='1' And Nvl(v_Ӥ��,0)=0 Then
            --У�Ի���ȼ�ʱ,ͬ�����Ĳ��˻���ȼ�
			
			--���˵�ǰӦ��������סԺ״̬
			v_Temp:=Null;
			Begin
				Select Decode(״̬,1,'�ȴ����',2,'����ת��',3,'��Ԥ��Ժ',Null) Into v_Temp From ������ҳ Where ����ID=v_����ID And ��ҳID=v_��ҳID;
			Exception
				When Others Then Null;
			End;
			If v_Temp IS Not Null Then
				v_Error:='���˵�ǰ����'||v_Temp||'״̬,ҽ��"'||v_ҽ������||'"����ͨ��У�ԡ�';
				Raise Err_Custom;
			End If;

            Begin
                --δ����ʱ,������,�ж��ʱ,ֻȡһ����
                Select �շ���ĿID Into v_����ȼ�ID From �����շѹ�ϵ Where ������ĿID=v_������ĿID And Rownum=1;
            Exception
                When Others Then NULL;
            End;
            IF v_����ȼ�ID IS Not NULL Then
                zl_���˱䶯��¼_Nurse(v_����ID,v_��ҳID,v_����ȼ�ID,v_Date,v_��Ա���,v_��Ա����);
            End IF;
            
            --��ֹͣ��������ȼ�ҽ��(����ȼ�Ӧ�ö�Ϊ"������"����,��ֻ��һ��δͣ)
            For r_Nurse In c_Nurse Loop
				Select Decode(Sign(r_Nurse.��ʼִ��ʱ��-v_��ʼʱ��),1,r_Nurse.��ʼִ��ʱ��,v_��ʼʱ��) Into v_��ֹʱ�� From Dual;
                ZL_����ҽ����¼_ֹͣ(r_Nurse.ҽ��ID,v_��ֹʱ��,v_����ҽ��);
            End Loop;
		ElsIf v_�������='Z' And v_��������='4' Then
			--����ҽ��У��ʱֹͣǰ��ĳ���,������ʼʱ��ֹ
			For r_NeedStop IN c_NeedStop(v_����ID,v_��ҳID,v_Ӥ��,v_��ʼʱ��) Loop
				Update ����ҽ����¼
					Set ҽ��״̬=8,
						ִ����ֹʱ��=Decode(Sign(��ʼִ��ʱ��-v_��ʼʱ��),1,��ʼִ��ʱ��,v_��ʼʱ��),
						ͣ��ʱ��=У��ʱ��_IN,
						ͣ��ҽ��=v_����ҽ��
				Where ID=r_NeedStop.ID;

				Insert Into ����ҽ��״̬(
					ҽ��ID,��������,������Ա,����ʱ��) 
				Select 
					ID,8,v_��Ա����,У��ʱ��_IN
				From ����ҽ����¼ 
				Where ID=r_NeedStop.ID;
			End Loop;
			--��ֹͣδȷ�ϵĳ���,��ֹʱ��������ʼ���,��ǰ����ֹʱ��(ͬʱ������������ҽ�������)
			For r_HaveStop IN c_HaveStop(v_����ID,v_��ҳID,v_Ӥ��,v_��ʼʱ��) Loop
				Update ����ҽ����¼
					Set ִ����ֹʱ��=Decode(Sign(��ʼִ��ʱ��-v_��ʼʱ��),1,��ʼִ��ʱ��,v_��ʼʱ��),
						ͣ��ʱ��=У��ʱ��_IN,
						ͣ��ҽ��=v_����ҽ��
				Where ID=r_HaveStop.ID;
				
				Update ����ҽ��״̬
					Set ����ʱ��=У��ʱ��_IN,
						������Ա=v_��Ա����
				Where ҽ��ID=r_HaveStop.ID And ��������=8;
			End Loop;
        End IF;
    End IF;
Exception
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_����ҽ����¼_У��;
/

CREATE OR REPLACE Procedure ZL_����ҽ������_Insert(
--���ܣ���д����ҽ�����ͼ�¼
    ҽ��ID_IN		����ҽ������.ҽ��ID%Type,
    ���ͺ�_IN       ����ҽ������.���ͺ�%Type,
    ��¼����_IN     ����ҽ������.��¼����%Type,
    NO_IN           ����ҽ������.NO%Type,
    ��¼���_IN     ����ҽ������.��¼���%Type,
    ��������_IN     ����ҽ������.��������%Type,
    �״�ʱ��_IN     ����ҽ������.�״�ʱ��%Type,
    ĩ��ʱ��_IN     ����ҽ������.ĩ��ʱ��%Type,
    ����ʱ��_IN     ����ҽ������.����ʱ��%Type,
    ִ��״̬_IN     ����ҽ������.ִ��״̬%Type,
    ִ�в���ID_IN   ����ҽ������.ִ�в���ID%Type,
    �Ʒ�״̬_IN     ����ҽ������.�Ʒ�״̬%Type,
    First_IN        Number:=0
--������First_IN=��ʾ�Ƿ�һ��ҽ���ĵ�һҽ����,�Ա㴦��ҽ���������(���ҩ,�䷽�ĵ�һ��,��Ϊ��ҩ;��,�䷽�巨,�÷�����Ϊ����������)
--      ��������_IN,�״�ʱ��_IN,ĩ��ʱ��_IN:��"������"����,����д��������,����д��ĩ��ʱ��(���ڻ���)��
) IS
    --�������˼�ҽ��(һ��ҽ���е�һ��)�����Ϣ���α�
    Cursor c_Advice is
        Select 
            Nvl(A.���ID,A.ID) AS ��ID,A.���,A.����ID,A.��ҳID,A.Ӥ��,B.����,B.��ǰ����ID,C.��������,
            A.�������,A.ҽ����Ч,A.ҽ��״̬,A.ҽ������,A.����ҽ��,A.����ʱ��,A.��ʼִ��ʱ��,A.�ϴ�ִ��ʱ��,A.ִ����ֹʱ��,
            A.ִ��ʱ�䷽��,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ
        From ����ҽ����¼ A,������Ϣ B,������ĿĿ¼ C
        Where A.����ID=B.����ID And A.������ĿID=C.ID And A.ID=ҽ��ID_IN
        Group BY Nvl(A.���ID,A.ID),A.���,A.����ID,A.��ҳID,A.Ӥ��,B.����,B.��ǰ����ID,C.��������,A.�������,A.ҽ����Ч,
            A.ҽ��״̬,A.ҽ������,A.����ҽ��,A.����ʱ��,A.��ʼִ��ʱ��,A.�ϴ�ִ��ʱ��,A.ִ����ֹʱ��,
            A.ִ��ʱ�䷽��,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ;
    r_Advice c_Advice%RowType;

    --��������(Ӥ��)������δͣ����(���䷽����)
    Cursor c_NeedStop(
        v_����ID    ����ҽ����¼.����ID%Type,
        v_��ҳID    ����ҽ����¼.��ҳID%Type,
        v_Ӥ��      ����ҽ����¼.Ӥ��%Type,
        v_StopTime  Date) is
        Select ID From ����ҽ����¼ 
        Where ����ID=v_����ID And ��ҳID=v_��ҳID And Nvl(Ӥ��,0)=Nvl(v_Ӥ��,0) 
            And Nvl(ҽ����Ч,0)=0 And ҽ��״̬ Not IN(1,2,4,8,9)
			And ��ʼִ��ʱ��<v_StopTime
        Order BY ���;
	--��������(Ӥ��)����ͣ��δȷ�ϵĳ���,��ִֹ��ʱ����ָ��ʱ��֮��
    Cursor c_HaveStop(
        v_����ID    ����ҽ����¼.����ID%Type,
        v_��ҳID    ����ҽ����¼.��ҳID%Type,
        v_Ӥ��		����ҽ����¼.Ӥ��%Type,
        v_StopTime	Date) is
        Select ID From ����ҽ����¼ 
        Where ����ID=v_����ID And ��ҳID=v_��ҳID And Nvl(Ӥ��,0)=Nvl(v_Ӥ��,0) 
            And Nvl(ҽ����Ч,0)=0 And ҽ��״̬=8 And ִ����ֹʱ��>v_StopTime
			And ��ʼִ��ʱ��<v_StopTime
        Order BY ���;

    --������ʱ����
    v_������        Number(1);--�Ƿ�����Գ���
    v_AutoStop		Number(1);
    v_Temp			Varchar2(255);
    v_��Ա���		���˷��ü�¼.����Ա���%Type;
    v_��Ա����		���˷��ü�¼.����Ա����%Type;

    v_Error			Varchar2(255);
    Err_Custom		Exception;
Begin
    --��ǰ������Ա    
    v_Temp:=zl_Identity;
    v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
    v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
    v_��Ա���:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
    v_��Ա����:=Substr(v_Temp,Instr(v_Temp,',')+1);

    --��һ��ҽ���ĵ�һ��ʱ����ҽ������
    If Nvl(First_IN,0)=1 Then
        Open c_Advice;
        Fetch c_Advice Into r_Advice;
        
        --�����������
        ---------------------------------------------------------------------------------------
        IF Nvl(r_Advice.ҽ��״̬,0)=4 Then
            --���Ҫ���͵�ҽ���Ƿ�����
            v_Error:='"'||r_Advice.����||'"��ҽ��"'||r_Advice.ҽ������||'"�Ѿ������������ϡ�'
                ||CHR(13)||CHR(10)||'�ò��˵�ҽ������ʧ�ܡ������¶�ȡ�����嵥���ԡ�';
            Raise Err_Custom;
        End IF;

        If Nvl(r_Advice.ҽ����Ч,0)=0 And r_Advice.�������<>'7' Then
            --����������ҩ����,��ҩ"��ѡƵ��"����,��ҩ"������"����
            
            --��鳤���Ƿ��ѱ�����
            If r_Advice.�ϴ�ִ��ʱ�� IS Not NULL Then
                If r_Advice.�ϴ�ִ��ʱ��>=�״�ʱ��_IN Then
                    v_Error:='"'||r_Advice.����||'"��ҽ��"'||r_Advice.ҽ������||'"�Ѿ��������˷��͡�'
                        ||CHR(13)||CHR(10)||'�ò��˵�ҽ������ʧ�ܡ������¶�ȡ�����嵥���ԡ�';
                    Raise Err_Custom;
                End IF;
            End IF;

            --��鳤������ǰ�Ƿ��ѱ��Զ�ֹͣ(������)
            If r_Advice.ִ����ֹʱ�� Is Not NULL Then
                If �״�ʱ��_IN>r_Advice.ִ����ֹʱ�� Then
                    v_Error:='"'||r_Advice.����||'"��ҽ��"'||r_Advice.ҽ������||'"�Ѿ���ֹͣ��'
                        ||CHR(13)||CHR(10)||'�ò��˵�ҽ������ʧ�ܡ������¶�ȡ�����嵥���ԡ�';
                    Raise Err_Custom;
                End IF;
            End IF;
        ElsIF Nvl(r_Advice.ҽ��״̬,0) IN(8,9) Then
            --���������䷽����

            --����Ƿ��ѱ�����(��������ԭ���Զ�ֹͣ)
            v_Error:='"'||r_Advice.����||'"��ҽ��"'||r_Advice.ҽ������||'"�Ѿ��������˷��͡�'
                ||CHR(13)||CHR(10)||'�ò��˵�ҽ������ʧ�ܡ������¶�ȡ�����嵥���ԡ�';
            Raise Err_Custom;
        End IF;
        
        --���ͺ��ҽ������
        ---------------------------------------------------------------------------------------
        If Nvl(r_Advice.ҽ����Ч,0)=0 And r_Advice.�������<>'7' Then
            --����ҽ��:�����ϴ�ִ��ʱ��
            Update ����ҽ����¼ 
                Set �ϴ�ִ��ʱ��=ĩ��ʱ��_IN 
            Where ID=r_Advice.��ID Or ���ID=r_Advice.��ID;
            
            --�ж��Ƿ�����Գ���
            v_������:=0;
            If r_Advice.ִ��ʱ�䷽�� IS NULL 
                And (Nvl(r_Advice.Ƶ�ʴ���,0)=0 Or Nvl(r_Advice.Ƶ�ʼ��,0)=0 Or r_Advice.�����λ IS NULL) Then
                v_������:=1;
            End IF;

            --Ԥ������ֹʱ����δֹͣ���Զ�ֹͣ
            IF r_Advice.ִ����ֹʱ�� IS Not NULL And Nvl(r_Advice.ҽ��״̬,0) Not IN(8,9) Then
                v_AutoStop:=0;
                If v_������=1 Then
                    --��ҩ"������"����
                    If Trunc(ĩ��ʱ��_IN)=Trunc(r_Advice.ִ����ֹʱ��-1) Then
                        v_AutoStop:=1; --��ֹ���첻ִ��
                    End IF;
                ElsIf zl_AdviceNextTime(ҽ��ID_IN)>r_Advice.ִ����ֹʱ�� Then
                    --��ҩ�������ҩ"��ѡƵ��"����
                    v_AutoStop:=1; --����ǵ���,������ִ��һ��
                End IF;

                If v_AutoStop=1 Then
                    Update ����ҽ����¼ 
                        Set ҽ��״̬=8,
                            ͣ��ʱ��=ĩ��ʱ��_IN,
                            ͣ��ҽ��=r_Advice.����ҽ��
                    Where ID=r_Advice.��ID Or ���ID=r_Advice.��ID;
                    
                    Insert Into ����ҽ��״̬(
                        ҽ��ID,��������,������Ա,����ʱ��) 
                    Select 
                        ID,8,v_��Ա����,����ʱ��_IN 
                    From ����ҽ����¼ 
                    Where ID=r_Advice.��ID Or ���ID=r_Advice.��ID;
                End IF;
            End IF;
        Else
            --����(���䷽����):ֹͣ��(Ϊ����һ,�䷽����Ҳ����ִ����ֹʱ��)
            Update ����ҽ����¼ 
                Set ҽ��״̬=8,
                    ִ����ֹʱ��=ĩ��ʱ��_IN,--Ϊһ��������ʱû��
                    �ϴ�ִ��ʱ��=ĩ��ʱ��_IN,--Ϊһ��������ʱû��
                    ͣ��ʱ��=����ʱ��_IN,
                    ͣ��ҽ��=r_Advice.����ҽ��
            Where ID=r_Advice.��ID Or ���ID=r_Advice.��ID;
            
            Insert Into ����ҽ��״̬(
                ҽ��ID,��������,������Ա,����ʱ��) 
            Select 
                ID,8,v_��Ա����,����ʱ��_IN 
            From ����ҽ����¼ 
            Where ID=r_Advice.��ID Or ���ID=r_Advice.��ID;
        End IF;

        --����ҽ���Ĵ���
        ---------------------------------------------------------------------------------------
        If r_Advice.�������='Z' And Nvl(r_Advice.��������,'0')<>'0' Then
            --(1-����;2-סԺ;)3-ת��;4-����(������);5-��Ժ;6-תԺ,7-����

            --��������ҽ��Ҫ�Զ�ֹͣ���˸�ҽ��֮ǰ(��ʱ����)����δͣ�ĳ���
            If r_Advice.�������� IN('3','5','6') Then
                For r_NeedStop IN c_NeedStop(r_Advice.����ID,r_Advice.��ҳID,r_Advice.Ӥ��,r_Advice.��ʼִ��ʱ��) Loop
                    Update ����ҽ����¼
                        Set ҽ��״̬=8,
							ִ����ֹʱ��=Decode(Sign(��ʼִ��ʱ��-r_Advice.��ʼִ��ʱ��),1,��ʼִ��ʱ��,r_Advice.��ʼִ��ʱ��),
                            ͣ��ʱ��=����ʱ��_IN,
                            ͣ��ҽ��=r_Advice.����ҽ��
                    Where ID=r_NeedStop.ID;

                    Insert Into ����ҽ��״̬(
                        ҽ��ID,��������,������Ա,����ʱ��) 
                    Select 
                        ID,8,v_��Ա����,����ʱ��_IN 
                    From ����ҽ����¼ 
                    Where ID=r_NeedStop.ID;
                End Loop;
				--��ֹͣδȷ�ϵĳ���,��ֹʱ��������ʼ���,��ǰ����ֹʱ��(ͬʱ������������ҽ�������)
				For r_HaveStop IN c_HaveStop(r_Advice.����ID,r_Advice.��ҳID,r_Advice.Ӥ��,r_Advice.��ʼִ��ʱ��) Loop
					Update ����ҽ����¼
						Set ִ����ֹʱ��=Decode(Sign(��ʼִ��ʱ��-r_Advice.��ʼִ��ʱ��),1,��ʼִ��ʱ��,r_Advice.��ʼִ��ʱ��),
							ͣ��ʱ��=����ʱ��_IN,
							ͣ��ҽ��=r_Advice.����ҽ��
					Where ID=r_HaveStop.ID;
					
					Update ����ҽ��״̬
						Set ����ʱ��=����ʱ��_IN,
							������Ա=v_��Ա����
					Where ҽ��ID=r_HaveStop.ID And ��������=8;
				End Loop;
            End IF;

            --��������⴦��
			If Nvl(r_Advice.Ӥ��,0)=0 Then
				If r_Advice.��������='3' And ִ�в���ID_IN IS Not NULL 
					And r_Advice.��ǰ����ID IS Not NULL And Nvl(r_Advice.��ǰ����ID,0)<>Nvl(ִ�в���ID_IN,0) Then
					--ת��ҽ��,�����˵Ǽ�ת�Ƶ�"ִ�п���ID"(��Ժ�����ҵ�ǰ������ת����Ҳ�ͬ�Ŵ���)
					zl_���˱䶯��¼_Change(r_Advice.����ID,r_Advice.��ҳID,ִ�в���ID_IN,v_��Ա���,v_��Ա����);
				ElsIf r_Advice.��������='5' Then
					--��Ժҽ��,�����˱��ΪԤ��Ժ
					ZL_���˱䶯��¼_PreOut(r_Advice.����ID,r_Advice.��ҳID,r_Advice.��ʼִ��ʱ��);
				ElsIf r_Advice.��������='6' Then
					--תԺҽ��,�����˱��ΪԤ��Ժ
					ZL_���˱䶯��¼_PreOut(r_Advice.����ID,r_Advice.��ҳID,r_Advice.��ʼִ��ʱ��);
				End IF;
			End IF;
        End IF;

        Close c_Advice;
    End IF;
    
    --��д���ͼ�¼
    ---------------------------------------------------------------------------------------
    Insert Into ����ҽ������(
        ҽ��ID,���ͺ�,��¼����,NO,��¼���,��������,������,����ʱ��,ִ��״̬,ִ�в���ID,�Ʒ�״̬,�״�ʱ��,ĩ��ʱ��)
    Values(
        ҽ��ID_IN,���ͺ�_IN,��¼����_IN,NO_IN,��¼���_IN,��������_IN,
        v_��Ա����,����ʱ��_IN,ִ��״̬_IN,ִ�в���ID_IN,�Ʒ�״̬_IN,
        �״�ʱ��_IN,ĩ��ʱ��_IN);

	--�Զ���Ϊ��ִ��ʱ����Ҫͬ���������ִ��״̬����˻���״̬
	If ִ��״̬_IN=1 Then
		ZL_����ҽ��ִ��_Finish(ҽ��ID_IN,���ͺ�_IN);
	End IF;
Exception
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_����ҽ������_Insert;
/

--����ת�ƹ��̵���(ǩ���������Ӧ��ҽ������ͬ��ת��)
Create Or Replace Procedure Zl1_Datamoveout1(d_Demoded In Number) As
   --------------------------------------------
   --����:d_Demoded,ת�����ݱ����Ƕ�������ǰ������
   --------------------------------------------
   d_Current Date;
   v_Version Varchar2(128);

   --------------------------------------------
   --ת��ָ��ID�Ĳ���Ԥ����¼�ӹ��̣�
   --------------------------------------------
   Procedure Zl_Move_Prepay(n_Settle_Id ����Ԥ����¼.����id%Type) As
   Begin
      For r_Rec In (Select * From ����Ԥ����¼ Where ����id = n_Settle_Id) Loop
         Update ��Ա�ɿ����
         Set ��� = Nvl(���, 0) - Nvl(Decode(r_Rec.��¼����, 1, r_Rec.���, 11, r_Rec.���, r_Rec.��Ԥ��), 0)
         Where �տ�Ա = r_Rec.����Ա���� And ���㷽ʽ = r_Rec.���㷽ʽ And ���� = 0;
         If Sql%Rowcount = 0 Then
            Insert Into ��Ա�ɿ����
               (�տ�Ա, ���㷽ʽ, ����, ���)
            Values
               (r_Rec.����Ա����, r_Rec.���㷽ʽ, 0,
                -1 * Nvl(Decode(r_Rec.��¼����, 1, r_Rec.���, 11, r_Rec.���, r_Rec.��Ԥ��), 0));
         End If;
      End Loop;
      Insert Into H����Ԥ����¼
         (Id, ��¼����, No, ʵ��Ʊ��, ��¼״̬, ����id, ��ҳid, ����id, �ɿλ, ��λ������, ��λ�ʺ�, ժҪ, ���, ���㷽ʽ,
          �������, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id)
         Select Id, ��¼����, No, ʵ��Ʊ��, ��¼״̬, ����id, ��ҳid, ����id, �ɿλ, ��λ������, ��λ�ʺ�, ժҪ, ���, ���㷽ʽ,
                �������, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id
         From ����Ԥ����¼
         Where ����id = n_Settle_Id;
      Delete ����Ԥ����¼ Where ����id = n_Settle_Id;
   End Zl_Move_Prepay;

   --------------------------------------------
   --ת��ָ��ID�Ĳ��˷��ü�¼�ӹ��̣�
   --------------------------------------------
   Procedure Zl_Move_Fee(n_Settle_Id ���˷��ü�¼.����id%Type) As
   Begin
      Insert Into H���˷��ü�¼
         (Id, ��¼����, No, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, ���ʵ�id, ����id, ��ҳid, ҽ�����, �����־,
          ���ʷ���, ����, �Ա�, ����, ��ʶ��, ����, ���˲���id, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����,
          ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������, ��������id, ������,
          ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ����, ����Ա���, ����Ա����, ����id, ���ʽ��, ���մ���id,
          ������Ŀ��, ���ձ���, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���)
         Select Id, ��¼����, No, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, �ಡ�˵�, ���ʵ�id, ����id, ��ҳid, ҽ�����,
                �����־, ���ʷ���, ����, �Ա�, ����, ��ʶ��, ����, ���˲���id, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ,
                ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, Ӥ����, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ������,
                ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, ����, ����Ա���, ����Ա����,
                ����id, ���ʽ��, ���մ���id, ������Ŀ��, ���ձ���, ͳ����, �Ƿ��ϴ�, ժҪ, �Ƿ���
         From ���˷��ü�¼
         Where ����id = n_Settle_Id;
      Delete ���˷��ü�¼ Where ����id = n_Settle_Id;
   End Zl_Move_Fee;

   --------------------------------------------
   --ת��ָ��ID��ҩƷ�շ���¼�ӹ��̣�
   --------------------------------------------
   Procedure Zl_Move_Medilist(n_Rec_Id ҩƷ�շ���¼.Id%Type) As
      r_Rec ҩƷ�շ���¼%Rowtype;
   Begin
      Select * Into r_Rec From ҩƷ�շ���¼ Where Id = n_Rec_Id;
      Update ҩƷ���
      Set �������� = Nvl(��������, 0) - r_Rec.���ϵ�� * Nvl(r_Rec.ʵ������ * r_Rec.����, 0),
          ʵ������ = Nvl(ʵ������, 0) - r_Rec.���ϵ�� * Nvl(r_Rec.ʵ������ * r_Rec.����, 0),
          ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) - r_Rec.���ϵ�� * Nvl(r_Rec.���۽��, 0),
          ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) - r_Rec.���ϵ�� * Nvl(r_Rec.���, 0), �ϴι�Ӧ��id = Nvl(�ϴι�Ӧ��id, r_Rec.��ҩ��λid),
          �ϴβɹ��� = Nvl(�ϴβɹ���, r_Rec.�ɱ���), �ϴ����� = Nvl(�ϴ�����, r_Rec.����), �ϴβ��� = Nvl(�ϴβ���, r_Rec.����),�ϴ���������=nvl(�ϴ���������,r_rec.��������)
      Where �ⷿid = r_Rec.�ⷿid And ҩƷid = r_Rec.ҩƷid And Nvl(����, 0) = Nvl(r_Rec.����, 0) And ���� = 0;
      If Sql%Notfound Then
         Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, ��������, ʵ������, ʵ�ʽ��, ʵ�ʲ��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, Ч��,
             �ϴβ���,�ϴ���������)
         Values
            (r_Rec.�ⷿid, r_Rec.ҩƷid, r_Rec.����, 0, -r_Rec.���ϵ�� * Nvl(r_Rec.ʵ������ * r_Rec.����, 0),
             -r_Rec.���ϵ�� * Nvl(r_Rec.ʵ������ * r_Rec.����, 0), -r_Rec.���ϵ�� * Nvl(r_Rec.���۽��, 0),
             -r_Rec.���ϵ�� * Nvl(r_Rec.���, 0), r_Rec.��ҩ��λid, r_Rec.�ɱ���, r_Rec.����, r_Rec.Ч��, r_Rec.����, r_Rec.��������);
      End If;
      Insert Into HҩƷ�շ���¼
         (Id, ��¼״̬, ����, No, ���, �ⷿid, ��ҩ��λid, ������id, �Է�����id, ���ϵ��, ҩƷid, ����,��������, ����, ����, Ч��,
          ����, ��д����, ʵ������, �ɱ���, �ɱ����, ����, ���ۼ�, ���۽��, ���, ժҪ, ������, ��������, ��ҩ��, �����,
          �������, �۸�id, ����id, ����, Ƶ��, �÷�, ��ҩ��ʽ, ��ҩ����, ��ҩ����, ���, ��Ʒ�ϸ�֤, �������, ���Ч��,������)
         Select Id, ��¼״̬, ����, No, ���, �ⷿid, ��ҩ��λid, ������id, �Է�����id, ���ϵ��, ҩƷid, ����,��������, ����, ����, Ч��,
                ����, ��д����, ʵ������, �ɱ���, �ɱ����, ����, ���ۼ�, ���۽��, ���, ժҪ, ������, ��������, ��ҩ��, �����,
                �������, �۸�id, ����id, ����, Ƶ��, �÷�, ��ҩ��ʽ, ��ҩ����, ��ҩ����, ���, ��Ʒ�ϸ�֤, �������, ���Ч��,������
         From ҩƷ�շ���¼
         Where Id = r_Rec.Id;
      Delete ҩƷ�շ���¼ Where Id = r_Rec.Id;
   End Zl_Move_Medilist;

   --------------------------------------------
   --ת��ָ��ID�Ĳ��˲�����¼�ӹ��̣�
   --------------------------------------------
   Procedure Zl_Move_Cpr(n_Rec_Id ���˲�����¼.Id%Type) As
   Begin
      Insert Into H���˲�����¼
         (Id, ����id, ��ҳid, �Һŵ�, Ӥ��, ����id, ��������, �ļ�id, ��������, ��д��, ��д����, ������, ��������, �鵵��,
          �鵵����, ������, ��������, ҽ��id)
         Select Id, ����id, ��ҳid, �Һŵ�, Ӥ��, ����id, ��������, �ļ�id, ��������, ��д��, ��д����, ������, ��������, �鵵��,
                �鵵����, ������, ��������, ҽ��id
         From ���˲�����¼
         Where Id = n_Rec_Id;
   
      Insert Into H������ӡ��¼
         (Id, ������¼id, ��ʼҳ��, ����ҳ��, ��ʼλ��, ����λ��, ��ӡʱ��, ��ӡ��)
         Select Id, ������¼id, ��ʼҳ��, ����ҳ��, ��ʼλ��, ����λ��, ��ӡʱ��, ��ӡ��
         From ������ӡ��¼
         Where ������¼id = n_Rec_Id;
   
      Insert Into H���˲����޶���¼
         (Id, ������¼id, ��д��, ��д����, �汾���)
         Select Id, ������¼id, ��д��, ��д����, �汾��� From ���˲����޶���¼ Where ������¼id = n_Rec_Id;
   
      Insert Into H���˲�������
         (Id, ����ʾ��id, ������¼id, �������, Ԫ������, Ԫ�ر���, ��дʱ��, �����ı�, �ı�ת��, ������ʾ, ��������, ������ɫ,
          ����λ��, ��������, ������ɫ, ����λ��, Ƕ�뷽ʽ, �����޶�id)
         Select Id, ����ʾ��id, ������¼id, �������, Ԫ������, Ԫ�ر���, ��дʱ��, �����ı�, �ı�ת��, ������ʾ, ��������,
                ������ɫ, ����λ��, ��������, ������ɫ, ����λ��, Ƕ�뷽ʽ, �����޶�id
         From ���˲�������
         Where ������¼id = n_Rec_Id;
   
      Insert Into H���˲����ı���
         (����id, �к�, ����)
         Select ����id, �к�, ���� From ���˲����ı��� Where ����id In (Select Id From ���˲������� Where ������¼id = n_Rec_Id);
   
      Insert Into H���˲���������
         (����id, �ؼ���, �ؼ���, ����, �̶���, �̶���, ��, ��, ��, ��, ����, �ϲ���, ����д, ������, ������id, ��ֵ����,
          ��������, ������λ)
         Select ����id, �ؼ���, �ؼ���, ����, �̶���, �̶���, ��, ��, ��, ��, ����, �ϲ���, ����д, ������, ������id, ��ֵ����,
                ��������, ������λ
         From ���˲���������
         Where ����id In (Select Id From ���˲������� Where ������¼id = n_Rec_Id);
   
      Insert Into H���˲������ͼ
         (����id, ����, ����, ����, �㼯, X1, Y1, X2, Y2, ���ɫ, ��䷽ʽ, ����ɫ, ����, �߿�)
         Select ����id, ����, ����, ����, �㼯, X1, Y1, X2, Y2, ���ɫ, ��䷽ʽ, ����ɫ, ����, �߿�
         From ���˲������ͼ
         Where ����id In (Select Id From ���˲������� Where ������¼id = n_Rec_Id);
   
      Insert Into H���˲����ⲿͼ
         (����id, ���, ͼ������, ͼ��·��, ͼ���ļ�)
         Select ����id, ���, ͼ������, ͼ��·��, ͼ���ļ�
         From ���˲����ⲿͼ
         Where ����id In (Select Id From ���˲������� Where ������¼id = n_Rec_Id);
   
      Insert Into H���˹�����¼
         (Id, ����id, ��ҳid, ��¼��Դ, ����id, ҩ��id, ҩ����, ��¼ʱ��, ��¼��, ���)
         Select Id, ����id, ��ҳid, ��¼��Դ, ����id, ҩ��id, ҩ����, ��¼ʱ��, ��¼��, ���
         From ���˹�����¼
         Where ����id In (Select Id From ���˲������� Where ������¼id = n_Rec_Id);
   
      Insert Into H������ϼ�¼
         (Id, ����id, ��ҳid, ��¼��Դ, ����id, �������, ����id, ���id, ֤��id, �������, ��Ժ���, �Ƿ�δ��, �Ƿ�����,
          ��¼����, ��¼��, ȡ��ʱ��, ȡ����, ҽ��id, ��ϴ���, �������)
         Select Id, ����id, ��ҳid, ��¼��Դ, ����id, �������, ����id, ���id, ֤��id, �������, ��Ժ���, �Ƿ�δ��, �Ƿ�����,
                ��¼����, ��¼��, ȡ��ʱ��, ȡ����, ҽ��id, ��ϴ���, �������
         From ������ϼ�¼
         Where ����id In (Select Id From ���˲������� Where ������¼id = n_Rec_Id);
   
      Insert Into H���˼�ϵͼ
         (����id, ���, ����, ĸ��, ����, ��ν, �Ա�, ״̬, ˵��, ����, ������ϵ, ������ϵ)
         Select ����id, ���, ����, ĸ��, ����, ��ν, �Ա�, ״̬, ˵��, ����, ������ϵ, ������ϵ
         From ���˼�ϵͼ
         Where ����id In (Select Id From ���˲������� Where ������¼id = n_Rec_Id);
   
      Insert Into H���������¼
         (Id, ����id, ��ҳid, ��¼��Դ, ����id, ��������, ������ʼʱ��, ��������ʱ��, ��������, ��������id, ������Ŀid, ��������,
          ����ҽʦ, ��һ����, �ڶ�����, ������ʿ, ����ʼʱ��, �������ʱ��, ����ʽ, ��������, ��������, ��Һ����, ����ҽʦ,
          ������ʼʱ��, ��������ʱ��, �п�, ����, ��¼����, ��¼��, ȡ��ʱ��, ȡ����)
         Select Id, ����id, ��ҳid, ��¼��Դ, ����id, ��������, ������ʼʱ��, ��������ʱ��, ��������, ��������id, ������Ŀid,
                ��������, ����ҽʦ, ��һ����, �ڶ�����, ������ʿ, ����ʼʱ��, �������ʱ��, ����ʽ, ��������, ��������,
                ��Һ����, ����ҽʦ, ������ʼʱ��, ��������ʱ��, �п�, ����, ��¼����, ��¼��, ȡ��ʱ��, ȡ����
         From ���������¼
         Where ����id In (Select Id From ���˲������� Where ������¼id = n_Rec_Id);
   
      Insert Into H����������ҩ
         (��¼id, ����, ���, ҩ��id, ҩ����, ����, ��ʽ)
         Select ��¼id, ����, ���, ҩ��id, ҩ����, ����, ��ʽ
         From ����������ҩ
         Where ��¼id In (Select Id From ���������¼ Where ����id In (Select Id From ���˲������� Where ������¼id = n_Rec_Id));
      Delete ���˲�����¼ Where Id = n_Rec_Id;
   End Zl_Move_Cpr;

   --------------------------------------------
   --ת��ָ��ID�Ĳ���ҽ����¼�ӹ��̣�
   --------------------------------------------
   Procedure Zl_Move_Order(n_Rec_Id ����ҽ����¼.Id%Type) As
   Begin
      Insert Into H����ҽ����¼
         (Id, ���id, ������Դ, ����id, ��ҳid, ����, �Ա�, ����, �Һŵ�, Ӥ��, ���˿���id, ���, ҽ��״̬, ҽ����Ч, �������, ������Ŀid,
          �걾��λ, �շ�ϸĿid, ��������, �ܸ�����, ҽ������, ҽ������, ִ�п���id, Ƥ�Խ��, ִ��Ƶ��, Ƶ�ʴ���, Ƶ�ʼ��,
          �����λ, ִ��ʱ�䷽��, �Ƽ�����, ִ������, ������־, ��ʼִ��ʱ��, ִ����ֹʱ��, �ϴ�ִ��ʱ��, ��������id, ����ҽ��,
          ����ʱ��, У�Ի�ʿ, У��ʱ��, ͣ��ҽ��, ͣ��ʱ��, ȷ��ͣ��ʱ��, ����id, ǰ��id, �Ƿ��ϴ�, ����,�����)
         Select Id, ���id, ������Դ, ����id, ��ҳid, ����, �Ա�, ����, �Һŵ�, Ӥ��, ���˿���id, ���, ҽ��״̬, ҽ����Ч, �������, ������Ŀid,
                �걾��λ, �շ�ϸĿid, ��������, �ܸ�����, ҽ������, ҽ������, ִ�п���id, Ƥ�Խ��, ִ��Ƶ��, Ƶ�ʴ���, Ƶ�ʼ��,
                �����λ, ִ��ʱ�䷽��, �Ƽ�����, ִ������, ������־, ��ʼִ��ʱ��, ִ����ֹʱ��, �ϴ�ִ��ʱ��, ��������id,
                ����ҽ��, ����ʱ��, У�Ի�ʿ, У��ʱ��, ͣ��ҽ��, ͣ��ʱ��, ȷ��ͣ��ʱ��, ����id, ǰ��id, �Ƿ��ϴ�, ����,�����
         From ����ҽ����¼
         Where Id = n_Rec_Id;
   
      Insert Into H����ҽ���Ƽ�
         (ҽ��id, �շ�ϸĿid, ����, ����, ����)
         Select ҽ��id, �շ�ϸĿid, ����, ����, ���� From ����ҽ���Ƽ� Where ҽ��id = n_Rec_Id;
      Delete ����ҽ���Ƽ� Where ҽ��id = n_Rec_Id;
      
	  --ת��ҽ��ǩ������(��Ϊ����ҽ������һ��ǩ������������ת������Ҫ�Ƚ���"FK_ǩ��ID"���)
      Insert Into Hҽ��ǩ����¼(
		ID, ǩ������, ǩ����Ϣ, ֤��ID, ǩ��ʱ��, ǩ����)
      Select 
	    ID, ǩ������, ǩ����Ϣ, ֤��ID, ǩ��ʱ��, ǩ���� 
	  From ҽ��ǩ����¼ 
	  Where ID IN (Select ǩ��ID From ����ҽ��״̬ Where ҽ��id = n_Rec_Id);
	  IF SQL%RowCount<>0 Then--����ת����ҽ��ʱ�Ѿ�ɾ��
          Delete ҽ��ǩ����¼ Where ID IN (Select ǩ��ID From ����ҽ��״̬ Where ҽ��id = n_Rec_Id);
      End IF;
   
      Insert Into H����ҽ��״̬
         (ҽ��id, ��������, ������Ա, ����ʱ��, ǩ��ID)
         Select ҽ��id, ��������, ������Ա, ����ʱ��, ǩ��ID From ����ҽ��״̬ Where ҽ��id = n_Rec_Id;
      Delete ����ҽ��״̬ Where ҽ��id = n_Rec_Id;
   
      --ҽ�����Ͳ���   
      Insert Into H����ҽ������
         (ҽ��id, ���ͺ�, ��¼����, No, ��¼���, ��������, ������, ����ʱ��, �״�ʱ��, ĩ��ʱ��, ִ��״̬, ִ�в���id, �Ʒ�״̬,
          ����id, ִ�м�, ִ�й���, ������, ����ʱ��, ��������)
         Select ҽ��id, ���ͺ�, ��¼����, No, ��¼���, ��������, ������, ����ʱ��, �״�ʱ��, ĩ��ʱ��, ִ��״̬, ִ�в���id,
                �Ʒ�״̬, ����id, ִ�м�, ִ�й���, ������, ����ʱ��, ��������
         From ����ҽ������
         Where ҽ��id = n_Rec_Id;
   
      Insert Into H����ҽ������
         (ҽ��id, ���ͺ�, ��¼����, No)
         Select ҽ��id, ���ͺ�, ��¼����, No From ����ҽ������ Where ҽ��id = n_Rec_Id;
      Delete ����ҽ������ Where ҽ��id = n_Rec_Id;
   
      Insert Into H����ҽ��ִ��
         (ҽ��id, ���ͺ�, Ҫ��ʱ��, ��������, ִ��ժҪ, ִ����, ִ��ʱ��, �Ǽ���, �Ǽ�ʱ��)
         Select ҽ��id, ���ͺ�, Ҫ��ʱ��, ��������, ִ��ժҪ, ִ����, ִ��ʱ��, �Ǽ���, �Ǽ�ʱ��
         From ����ҽ��ִ��
         Where ҽ��id = n_Rec_Id;
      Delete ����ҽ��ִ�� Where ҽ��id = n_Rec_Id;
   
      Insert Into HӰ�����¼
         (ҽ��id, ���ͺ�, Ӱ�����, ����, ����, Ӣ����, �Ա�, ����, ��������, ���, ����, ������, ���Ž�Ƭ, ���uid, λ��һ,
          λ�ö�, λ����, ����豸, ����ͼ��, ��������)
         Select ҽ��id, ���ͺ�, Ӱ�����, ����, ����, Ӣ����, �Ա�, ����, ��������, ���, ����, ������, ���Ž�Ƭ, ���uid,
                λ��һ, λ�ö�, λ����, ����豸, ����ͼ��, ��������
         From Ӱ�����¼
         Where ҽ��id = n_Rec_Id;
      For r_Ris In (Select ���uid
                    From Ӱ�����¼ r, ����ҽ������ s
                    Where r.ҽ��id = s.ҽ��id And r.���ͺ� = s.���ͺ� And s.ҽ��id = n_Rec_Id) Loop
         Insert Into HӰ��������
            (����uid, ���uid, ���к�, ��������, �ɼ�ʱ��)
            Select ����uid, ���uid, ���к�, ��������, �ɼ�ʱ�� From Ӱ�������� Where ���uid = r_Ris.���uid;
         For r_Seq In (Select ����uid From Ӱ�������� Where ���uid = r_Ris.���uid) Loop
            Insert Into HӰ����ͼ��
               (ͼ��uid, ����uid, ͼ���, ͼ������)
               Select ͼ��uid, ����uid, ͼ���, ͼ������ From Ӱ����ͼ�� Where ����uid = r_Seq.����uid;
            Delete Ӱ����ͼ�� Where ����uid = r_Seq.����uid;
         End Loop;
         Delete Ӱ�������� Where ���uid = r_Ris.���uid;
      End Loop;
      Delete Ӱ�����¼ Where ҽ��id = n_Rec_Id;
   
      Delete ����ҽ������ Where ҽ��id = n_Rec_Id;
   
      --������¼����
      Insert Into H����������¼
         (Id, ҽ��id, ����id, ����id, ��ҳid, ��¼��Դ, ��������, ������ʼʱ��, ��������ʱ��, ������ģ, ����ʼʱ��,
          �������ʱ��, ����ʽ, ��������, ��������, ��Һ����, ������ʼʱ��, ��������ʱ��, �п�, ����, �޾�����, ������, ������id,
          ������id, ��¼����, ��¼��, ȡ��ʱ��, ȡ����)
         Select Id, ҽ��id, ����id, ����id, ��ҳid, ��¼��Դ, ��������, ������ʼʱ��, ��������ʱ��, ������ģ, ����ʼʱ��,
                �������ʱ��, ����ʽ, ��������, ��������, ��Һ����, ������ʼʱ��, ��������ʱ��, �п�, ����, �޾�����, ������,
                ������id, ������id, ��¼����, ��¼��, ȡ��ʱ��, ȡ����
         From ����������¼
         Where ҽ��id = n_Rec_Id;
      For r_Ops In (Select Id From ����������¼ Where ҽ��id = n_Rec_Id) Loop
         Insert Into H�����������
            (��¼id, ����, ȱʡ, ��������, ��������id, ������Ŀid)
            Select ��¼id, ����, ȱʡ, ��������, ��������id, ������Ŀid From ����������� Where ��¼id = r_Ops.Id;
         Delete ����������� Where ��¼id = r_Ops.Id;
      
         Insert Into H����������ҩ
            (��¼id, ����, ���, ҩƷid, ҩ��id, ҩƷ����, ׼������, ʹ������, ��ҩ��ʽ, ִ�п���id)
            Select ��¼id, ����, ���, ҩƷid, ҩ��id, ҩƷ����, ׼������, ʹ������, ��ҩ��ʽ, ִ�п���id
            From ����������ҩ
            Where ��¼id = r_Ops.Id;
         Delete ����������ҩ Where ��¼id = r_Ops.Id;
      
         Insert Into H����������Ա
            (��¼id, ����id, ��λ, ��Աid, ����, ����)
            Select ��¼id, ����id, ��λ, ��Աid, ����, ���� From ����������Ա Where ��¼id = r_Ops.Id;
         Delete ����������Ա Where ��¼id = r_Ops.Id;
      
         Insert Into H������������
            (��¼id, ���, ����id, ׼������, ʵ������, �����, ִ�п���id, ����˵��)
            Select ��¼id, ���, ����id, ׼������, ʵ������, �����, ִ�п���id, ����˵��
            From ������������
            Where ��¼id = r_Ops.Id;
         Delete ������������ Where ��¼id = r_Ops.Id;
      
         Insert Into H���������Ƽ�
            (��¼id, ���, ����, ����, �շ�ϸĿid, ִ�п���id)
            Select ��¼id, ���, ����, ����, �շ�ϸĿid, ִ�п���id From ���������Ƽ� Where ��¼id = r_Ops.Id;
         Delete ���������Ƽ� Where ��¼id = r_Ops.Id;
      
      End Loop;
      Delete ����������¼ Where ҽ��id = n_Rec_Id;
   
      Insert Into H��������״̬
         (ҽ��id, ���, �ϴ����, ��������, ��¼״̬, ������, ����ʱ��, ����˵��, ��ǰ״̬, ���ݺ�)
         Select ҽ��id, ���, �ϴ����, ��������, ��¼״̬, ������, ����ʱ��, ����˵��, ��ǰ״̬, ���ݺ�
         From ��������״̬
         Where ҽ��id = n_Rec_Id;
      Delete ��������״̬ Where ҽ��id = n_Rec_Id;
   
      --���鲿��
      Insert Into H����걾��¼
         (Id, ҽ��id, �걾���, ����ʱ��, ������, �걾����, ������, ����ʱ��, ����״̬, ������, ����ʱ��, �����, ���ʱ��,
          �ϲ������, ��ӡ����, ��������, ����id, ��������, ������, ��ע, δͨ�����ԭ��, ����ʱ��, �걾��̬, �Ƿ��ʿ�Ʒ,
          ִ�п���id)
         Select Id, ҽ��id, �걾���, ����ʱ��, ������, �걾����, ������, ����ʱ��, ����״̬, ������, ����ʱ��, �����, ���ʱ��,
                �ϲ������, ��ӡ����, ��������, ����id, ��������, ������, ��ע, δͨ�����ԭ��, ����ʱ��, �걾��̬, �Ƿ��ʿ�Ʒ,
                ִ�п���id
         From ����걾��¼
         Where ҽ��id = n_Rec_Id;
      For r_Retu In (Select Id From ������ͨ��� Where ����걾id In (Select Id From ����걾��¼ Where ҽ��id = n_Rec_Id)) Loop
         Insert Into H������ͨ���
            (Id, ����걾id, ������Ŀid, ������, �����־, ����ο�, �޸���, �޸�ʱ��, ��¼����, ԭʼ���, ԭʼ��¼ʱ��, ��¼��,
             �Ƿ����, �޸�ԭ��, ϸ��id, ����id, ��������)
            Select Id, ����걾id, ������Ŀid, ������, �����־, ����ο�, �޸���, �޸�ʱ��, ��¼����, ԭʼ���, ԭʼ��¼ʱ��,
                   ��¼��, �Ƿ����, �޸�ԭ��, ϸ��id, ����id, ��������
            From ������ͨ���
            Where Id = r_Retu.Id;
         Insert Into H����ҩ�����
            (ϸ�����id, ������id, �޸���, �޸�ʱ��, ���, �������, ��¼����, ����id)
            Select ϸ�����id, ������id, �޸���, �޸�ʱ��, ���, �������, ��¼����, ����id
            From ����ҩ�����
            Where ϸ�����id = r_Retu.Id;

		 Delete ����ҩ����� Where ϸ�����id = r_Retu.Id;
         Delete ������ͨ��� Where Id = r_Retu.Id;
      End Loop;
      Delete ����걾��¼ Where ҽ��id = n_Rec_Id;
   
      Delete ����ҽ����¼ Where Id = n_Rec_Id;
   End Zl_Move_Order;

   --------------------------------------------
   --����Ϊ��������
   --------------------------------------------
Begin
   --��ֹ���ݲ�һ�£��Ƚ���һЩԼ��
   Begin
      Execute Immediate 'Alter Table ����ҽ����¼ Modify Constraint ����ҽ����¼_FK_����ID Disable';
      Execute Immediate 'Alter Table ����ҽ������ Modify Constraint ����ҽ������_FK_����ID Disable';
	  Execute Immediate 'Alter Table ����ҽ��״̬ Modify Constraint ����ҽ��״̬_FK_ǩ��ID Disable';
   Exception
      When Others Then Null;
   End;
   
   Select Trunc(Sysdate) Into d_Current From Dual;

   --------------------------------------------
   --1��ָ��ʱ��ǰ��ת����û�г�Ԥ���շѺͶ�Ӧ��ҩ��¼ת��
   For r_Settle In (Select l.����id
                    From ����Ԥ����¼ l,
                         (Select ����id From ����Ԥ����¼ Where �տ�ʱ�� < d_Current - d_Demoded And ��¼���� In (3, 4, 5)) c
                    Where l.����id = c.����id
                    Group By l.����id
                    Having Sum(Decode(��¼����, 1, 1, 11, 1, 0)) = 0
					Union
					Select ����id From ���˷��ü�¼ 
					Where �Ǽ�ʱ�� < d_Current - d_Demoded And ��¼����=4 
						And Nvl(Ӧ�ս��,0)=0 And Nvl(ʵ�ս��,0)=0
                    Minus
                    Select Distinct d.����id
                    From ҩƷ�շ���¼ l,
                         (Select d.No, d.Id, d.��¼����, d.����id
                           From ���˷��ü�¼ d
                           Where d.�Ǽ�ʱ�� < d_Current - d_Demoded And d.��¼���� = 1 And d.�շ���� In ('4', '5', '6', '7')) d
                    Where l.No = d.No And l.����id = d.Id And Nvl(��ҩ��ʽ, 0) <> -1 And
                          (l.������� >= d_Current - d_Demoded Or l.������� Is Null) And l.���� In (8, 24)) Loop
   
      Zl_Move_Prepay(r_Settle.����id);
      For r_Rxlist In (Select m.Id
                       From ҩƷ�շ���¼ m,
                            (Select Id, No, ���, ��¼����
                              From ���˷��ü�¼
                              Where ����id = r_Settle.����id And �շ���� In ('4', '5', '6', '7') And ��¼���� In (1, 2)) e
                       Where m.No = e.No And m.����id = e.Id And
                             (e.��¼���� = 1 And m.���� In (8, 24) Or e.��¼���� <> 1 And m.���� In (9, 10, 25, 26))) Loop
         Zl_Move_Medilist(r_Rxlist.Id);
      End Loop;
      Zl_Move_Fee(r_Settle.����id);
   
      Commit;
   End Loop;

   --------------------------------------------
   --2��ָ��ʱ��ǰ��ת���Ľ��ʼ�¼������Ԥ����¼�����˼��ʷ��úͶ�Ӧ���ʷ�ҩ��¼ת��
   For r_Settle In (Select ����id
                    From ���˷��ü�¼
                    Where �Ǽ�ʱ�� < d_Current - d_Demoded And ��¼���� In (1, 4, 5) And Nvl(���ʷ���,0) <> 1
                    Union
                    Select Id As ����id
                    From ���˽��ʼ�¼ l
                    Where l.�շ�ʱ�� < d_Current - d_Demoded
                    Minus
                    Select Distinct d.����id
                    From ����Ԥ����¼ d,
                         (Select d.No
                           From ����Ԥ����¼ d,
                                (Select ����id
                                  From ���˷��ü�¼
                                  Where �Ǽ�ʱ�� < d_Current - d_Demoded And ��¼���� In (1, 4, 5) And Nvl(���ʷ���,0) <> 1
                                  Union
                                  Select Id As ����id From ���˽��ʼ�¼ Where �շ�ʱ�� < d_Current - d_Demoded) l
                           Where d.����id = l.����id And d.��¼���� In (1, 11)
                           Group By d.No
                           Having d.No Is Not Null And Sum(d.���) - Sum(d.��Ԥ��) <> 0) n
                    Where d.No = n.No And d.��¼���� In (1, 11)
                    Minus
                    Select Distinct d.����id
                    From ���˷��ü�¼ d,
                         (Select d.No, d.���, Decode(d.��¼����, 12, 2, 13, 3, 15, 5, d.��¼����) As ��¼����
                           From ���˷��ü�¼ d, ���˽��ʼ�¼ l
                           Where d.����id = l.Id And l.�շ�ʱ�� < d_Current - d_Demoded And d.��¼���� In (2, 12, 3, 13, 5, 15) And
                                 d.���ʷ��� = 1
                           Group By d.No, d.���, Decode(d.��¼����, 12, 2, 13, 3, 15, 5, d.��¼����)
                           Having d.No Is Not Null And d.��� Is Not Null And Nvl(Sum(d.ʵ�ս��),0) - Nvl(Sum(d.���ʽ��),0) <> 0) n
                    Where d.No = n.No And d.��� = n.��� And Decode(d.��¼����, 12, 2, 13, 3, 15, 5, d.��¼����) = n.��¼����
                    Minus
                    Select Distinct d.����id
                    From ҩƷ�շ���¼ l,
                         (Select d.No, d.Id, d.��¼����, d.����id
                           From ���˷��ü�¼ d
                           Where d.�Ǽ�ʱ�� < d_Current - d_Demoded And d.��¼���� In (1, 2) And d.�շ���� In ('4', '5', '6', '7')) d
                    Where l.No = d.No And l.����id = d.Id And Nvl(��ҩ��ʽ, 0) <> -1 And
                          (l.������� >= d_Current - d_Demoded Or l.������� Is Null) And
                          (d.��¼���� = 1 And l.���� In (8, 24) Or d.��¼���� <> 1 And l.���� In (9, 10, 25, 26))) Loop
   
      Insert Into H���˽��ʼ�¼
         (Id, No, ʵ��Ʊ��, ��¼״̬, ����id, ����Ա���, ����Ա����, �շ�ʱ��, ��ʼ����, ��������)
         Select Id, No, ʵ��Ʊ��, ��¼״̬, ����id, ����Ա���, ����Ա����, �շ�ʱ��, ��ʼ����, ��������
         From ���˽��ʼ�¼
         Where Id = r_Settle.����id;
   
      Zl_Move_Prepay(r_Settle.����id);
      For r_Rxlist In (Select m.Id
                       From ҩƷ�շ���¼ m,
                            (Select Id, No, ���, ��¼����
                              From ���˷��ü�¼
                              Where ����id = r_Settle.����id And �շ���� In ('4', '5', '6', '7') And ��¼���� In (1, 2)) e
                       Where m.No = e.No And m.����id = e.Id And
                             (e.��¼���� = 1 And m.���� In (8, 24) Or e.��¼���� <> 1 And m.���� In (9, 10, 25, 26))) Loop
         Zl_Move_Medilist(r_Rxlist.Id);
      End Loop;
      Zl_Move_Fee(r_Settle.����id);
   
      Delete ���˽��ʼ�¼ Where Id = r_Settle.����id;
   
      Commit;
   End Loop;

   --------------------------------------------
   --3��ָ��ʱ��ǰ������ﲡ��ҽ����������(ǰ���ǲ��˱��ξ�������Ѿ�����ת��)
   For r_Regist In (Select r.Id, r.����id, r.No
                    From ���˹Һż�¼ r, (Select No From ���˷��ü�¼ Where �Ǽ�ʱ�� < d_Current - d_Demoded And ��¼���� = 4) d,
                         (Select r.Id As �Һ�id
                           From ����ҽ����¼ a, ���˹Һż�¼ r
                           Where a.�Һŵ� = r.No And r.�Ǽ�ʱ�� < d_Current - d_Demoded
                           Group By r.Id
                           Having Max(a.ͣ��ʱ��) >= d_Current - d_Demoded) a,
                         (Select a.�Һ�id
                           From ���˷��ü�¼ e,
                                (Select a.Id, r.Id As �Һ�id
                                  From ����ҽ����¼ a, ���˹Һż�¼ r
                                  Where a.�Һŵ� = r.No And r.�Ǽ�ʱ�� < d_Current - d_Demoded) a
                           Where e.ҽ����� = a.Id) e
                    Where r.No = d.No(+) And r.Id = a.�Һ�id(+) And r.Id = e.�Һ�id(+) 
						And r.ִ��״̬<>2 And r.�Ǽ�ʱ�� < d_Current - d_Demoded
                    Group By r.Id, r.����id, r.No
                    Having Count(d.No) = 0 And Count(a.�Һ�id) = 0 And Count(e.�Һ�id) = 0) Loop
   
      Insert Into H���˹Һż�¼
         (Id, No, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��,
          ����Ա���, ����Ա����, ժҪ)
         Select Id, No, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��,
                �Ǽ�ʱ��, ����Ա���, ����Ա����, ժҪ
         From ���˹Һż�¼
         Where Id = r_Regist.Id;
   
      For r_Cpr In (Select Id From ���˲�����¼ Where ����id = r_Regist.����id And �Һŵ� = r_Regist.No) Loop
         Zl_Move_Cpr(r_Cpr.Id);
      End Loop;
   
      For r_Order In (Select Id From ����ҽ����¼ Where ����id = r_Regist.����id And �Һŵ� = r_Regist.No) Loop
         Zl_Move_Order(r_Order.Id);
      End Loop;
   
      Delete ���˹Һż�¼ Where Id = r_Regist.Id;
   
      Commit;
   End Loop;

   --------------------------------------------
   --4��ָ��ʱ��ǰ��Ժ��סԺ���ﲡ��ҽ����������(ǰ���ǲ��˱��ξ�������Ѿ�����ת��)
   For r_Page In (Select ����id, ��ҳid
                  From ������ҳ p
                  Where ��Ժ���� < d_Current - d_Demoded And Nvl(����ת��, 0) <> 1 And Not Exists
                   (Select 1 From ���˷��ü�¼ Where ����id = p.����id And ��ҳid = p.��ҳid)) Loop
   
      For r_Cpr In (Select Id From ���˲�����¼ Where ����id = r_Page.����id And ��ҳid = r_Page.��ҳid) Loop
         Zl_Move_Cpr(r_Cpr.Id);
      End Loop;
   
      For r_Order In (Select Id From ����ҽ����¼ Where ����id = r_Page.����id And ��ҳid = r_Page.��ҳid) Loop
         Zl_Move_Order(r_Order.Id);
      End Loop;
   
      Update ������ҳ Set ����ת�� = 1 Where ����id = r_Page.����id And ��ҳid = r_Page.��ҳid;
   
      Commit;
   End Loop;

   --------------------------------------------
   --5��ָ��ʱ��ǰ��Ա�ɿ��¼��ת��
   For r_Hand In (Select * From ��Ա�ɿ��¼ Where �Ǽ�ʱ�� < d_Current - d_Demoded) Loop
      Insert Into H��Ա�ɿ��¼
         (Id, ����id, �տ�Ա, ���㷽ʽ, �����, ���, ժҪ, ��ֹʱ��, �Ǽ�ʱ��, �Ǽ���)
         Select Id, ����id, �տ�Ա, ���㷽ʽ, �����, ���, ժҪ, ��ֹʱ��, �Ǽ�ʱ��, �Ǽ��� 
         From ��Ա�ɿ��¼
         Where Id = r_Hand.Id;
   
      Update ��Ա�ɿ����
      Set ��� = Nvl(���, 0) + r_Hand.���
      Where �տ�Ա = r_Hand.�տ�Ա And ���㷽ʽ = r_Hand.���㷽ʽ And ���� = 0;
      If Sql%Rowcount = 0 Then
         Insert Into ��Ա�ɿ���� (�տ�Ա, ���㷽ʽ, ����, ���) Values (r_Hand.�տ�Ա, r_Hand.���㷽ʽ, 0, r_Hand.���);
      End If;
   
      If r_Hand.Id = r_Hand.����id Then
         Insert Into H��Ա�ɿ����
            (����id, ����, ��¼id)
            Select ����id, ����, ��¼id From ��Ա�ɿ���� Where ����id = r_Hand.����id;
         Delete From ��Ա�ɿ���� Where ����id = r_Hand.����id;
      End If;
   
      Delete ��Ա�ɿ��¼ Where Id = r_Hand.Id;
      Commit;
   End Loop;

   --------------------------------------------
   --6��ָ��ʱ��ǰ�����Ʊ�����ݵ�ת��
   For r_Bill In (Select d.����id As Id
                  From Ʊ��ʹ����ϸ d, (Select Id From Ʊ�����ü�¼ Where �Ǽ�ʱ�� < d_Current - d_Demoded And ʣ������ = 0) l
                  Where d.����id = l.Id
                  Group By d.����id
                  Having Max(d.ʹ��ʱ��) < d_Current - d_Demoded) Loop
      Insert Into HƱ�����ü�¼
         (Id, Ʊ��, ������, ǰ׺�ı�, ��ʼ����, ��ֹ����, ʹ�÷�ʽ, �Ǽ�ʱ��, �Ǽ���, ��ǰ����, ʣ������)
         Select Id, Ʊ��, ������, ǰ׺�ı�, ��ʼ����, ��ֹ����, ʹ�÷�ʽ, �Ǽ�ʱ��, �Ǽ���, ��ǰ����, ʣ������
         From Ʊ�����ü�¼
         Where Id = r_Bill.Id;
   
      For r_Used In (Select Id, ��ӡid From Ʊ��ʹ����ϸ Where ����id = r_Bill.Id) Loop
         Insert Into HƱ��ʹ����ϸ
            (Id, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����)
            Select Id, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ���� From Ʊ��ʹ����ϸ Where Id = r_Used.Id;
      
         Insert Into HƱ�ݴ�ӡ����
            (Id, ��������, No)
            Select Id, ��������, No From Ʊ�ݴ�ӡ���� Where Id = r_Used.��ӡid;
         Delete Ʊ�ݴ�ӡ���� Where Id = r_Used.��ӡid;
      
         Delete Ʊ��ʹ����ϸ Where Id = r_Used.Id;
      End Loop;
   
      Delete Ʊ�����ü�¼ Where Id = r_Bill.Id;
      Commit;
   End Loop;

   --   --------------------------------------------
   --   --7��ָ��ʱ��ǰ����˵�ҩƷ�����¼
   --   For r_Flow In (Select Id From ҩƷ�շ���¼ Where ������� < d_Current - d_Demoded And ���� Not In (8, 9, 10, 24, 25, 26)) Loop
   --      Zl_Move_Medilist(r_Flow.Id);
   --      Commit;
   --   End Loop;
   --
   --   --------------------------------------------
   --   --8��ָ��ʱ��ǰ�Ѿ���ȷ���Խ����Ӧ����͸����¼ת��
   --   For r_Clear In (Select s.�������
   --                   From �����¼ p, Ӧ����¼ m, (Select Distinct ������� From �����¼ Where ������� < d_Current - d_Demoded) s
   --                   Where p.������� = s.������� And m.������� = s.�������
   --                   Group By s.�������
   --                   Having Max(p.�������) < d_Current - d_Demoded And Max(m.�������) < d_Current - d_Demoded) Loop
   --
   --      Insert Into HӦ����¼
   --         (Id, ��¼����, ��¼״̬, No, �շ�id, ��λid, Ʒ��, ���, ����, ����, ������λ, ��ⵥ�ݺ�, ���ݽ��, ����, �ɹ���,
   --          �ɹ����, ��Ʊ��, ��Ʊ����, ��Ʊ���, �ƶ�����, �ƻ����, �ƻ���, �ƻ�����, ������, ��������, �����, �������, ժҪ,
   --          �������, �ƻ����, ϵͳ��ʶ)
   --         Select Id, ��¼����, ��¼״̬, No, �շ�id, ��λid, Ʒ��, ���, ����, ����, ������λ, ��ⵥ�ݺ�, ���ݽ��, ����, �ɹ���,
   --                �ɹ����, ��Ʊ��, ��Ʊ����, ��Ʊ���, �ƶ�����, �ƻ����, �ƻ���, �ƻ�����, ������, ��������, �����, �������,
   --                ժҪ, �������, �ƻ����, ϵͳ��ʶ
   --         From Ӧ����¼
   --         Where ������� = r_Clear.�������;
   --      Delete Ӧ����¼ Where ������� = r_Clear.�������;
   --
   --      Insert Into H�����¼
   --         (Id, ��¼״̬, No, ���, Ԥ����, ��λid, ���, ���㷽ʽ, �������, ժҪ, ������, ��������, �����, �������, �������)
   --         Select Id, ��¼״̬, No, ���, Ԥ����, ��λid, ���, ���㷽ʽ, �������, ժҪ, ������, ��������, �����, �������,
   --                �������
   --         From �����¼
   --         Where ������� = r_Clear.�������;
   --      Delete �����¼ Where ������� = r_Clear.�������;
   --
   --      Commit;
   --   End Loop;

   --------------------------------------------
   --9���������ݱ������ؽ�:Oracle 8.1.6/8.1.7�汾��Ҫ�ֹ�ִ���ؽ�����
   --����Լ��
   Begin
      Execute Immediate 'Alter Table ����ҽ����¼ Modify Constraint ����ҽ����¼_FK_����ID Enable';
      Execute Immediate 'Alter Table ����ҽ������ Modify Constraint ����ҽ������_FK_����ID Enable';
	  Execute Immediate 'Alter Table ����ҽ��״̬ Modify Constraint ����ҽ��״̬_FK_ǩ��ID Enable';
   Exception
      When Others Then Null;
   End;

   Select Version Into v_Version From Product_Component_Version Where Upper(Substr(PRODUCT,1,6))=Upper('Oracle');
   If v_Version>='9' Then
	   For r_Sql In (Select 'alter index ' || Index_Name || ' rebuild online nologging' As Sqltext
					 From User_Indexes
					 Where Table_Name In (Select Substr(Table_Name, 2) From User_Tables Where Table_Name Like 'H%')) Loop
		  Execute Immediate r_Sql.Sqltext;
	   End Loop;
   End If;

   --------------------------------------------
   Update Zldatamove
   Set �ϴ����� = Greatest(d_Current - d_Demoded, Nvl(�ϴ�����, d_Current - d_Demoded))
   Where ϵͳ In (Select ��� From Zlsystems Where Upper(������) = Zl_Owner And ��� Like '1%') And ��� = 1;
   Commit;
End Zl1_Datamoveout1;
/

--------------------------------------------
--���ݷ��ع���2����ѡ���ز���ĳ������סԺҽ������
--------------------------------------------
CREATE OR REPLACE Procedure Zl_Retu_Clinic(n_Patiid In Number, v_Times In Varchar2, n_Flag In Number) As
   --------------------------------------------
   --����:n_Patiid,����id
   --     v_Times,�Һŵ��Ż�סԺ��ҳid
   --     n_Flag,�����סԺ��־:0-����,1-סԺ
   --------------------------------------------
   --------------------------------------------
   --����ָ��ID�Ĳ��˲�����¼�ӹ��̣�
   --------------------------------------------
   Procedure Zl_Retu_Cpr(n_Rec_Id H���˲�����¼.Id%Type) As
   Begin
      Insert Into ���˲�����¼
         (Id, ����id, ��ҳid, �Һŵ�, Ӥ��, ����id, ��������, �ļ�id, ��������, ��д��, ��д����, ������, ��������, �鵵��,
          �鵵����, ������, ��������, ҽ��id)
         Select Id, ����id, ��ҳid, �Һŵ�, Ӥ��, ����id, ��������, �ļ�id, ��������, ��д��, ��д����, ������, ��������, �鵵��,
                �鵵����, ������, ��������, ҽ��id
         From H���˲�����¼
         Where Id = n_Rec_Id;
   
      Insert Into ������ӡ��¼
         (Id, ������¼id, ��ʼҳ��, ����ҳ��, ��ʼλ��, ����λ��, ��ӡʱ��, ��ӡ��)
         Select Id, ������¼id, ��ʼҳ��, ����ҳ��, ��ʼλ��, ����λ��, ��ӡʱ��, ��ӡ��
         From H������ӡ��¼
         Where ������¼id = n_Rec_Id;
   
      Insert Into ���˲����޶���¼
         (Id, ������¼id, ��д��, ��д����, �汾���)
         Select Id, ������¼id, ��д��, ��д����, �汾��� 
		 From H���˲����޶���¼ Where ������¼id = n_Rec_Id;
   
      Insert Into ���˲�������
         (Id, ����ʾ��id, ������¼id, �������, Ԫ������, Ԫ�ر���, ��дʱ��, �����ı�, �ı�ת��, ������ʾ, ��������, ������ɫ,
          ����λ��, ��������, ������ɫ, ����λ��, Ƕ�뷽ʽ, �����޶�id)
         Select Id, ����ʾ��id, ������¼id, �������, Ԫ������, Ԫ�ر���, ��дʱ��, �����ı�, �ı�ת��, ������ʾ, ��������,
                ������ɫ, ����λ��, ��������, ������ɫ, ����λ��, Ƕ�뷽ʽ, �����޶�id
         From H���˲�������
         Where ������¼id = n_Rec_Id;
   
      Insert Into ���˲����ı���
         (����id, �к�, ����)
         Select ����id, �к�, ����
         From H���˲����ı���
         Where ����id In (Select Id From H���˲������� Where ������¼id = n_Rec_Id);
   
      Insert Into ���˲���������
         (����id, �ؼ���, �ؼ���, ����, �̶���, �̶���, ��, ��, ��, ��, ����, �ϲ���, ����д, ������, ������id, ��ֵ����,
          ��������, ������λ)
         Select ����id, �ؼ���, �ؼ���, ����, �̶���, �̶���, ��, ��, ��, ��, ����, �ϲ���, ����д, ������, ������id, ��ֵ����,
                ��������, ������λ
         From H���˲���������
         Where ����id In (Select Id From H���˲������� Where ������¼id = n_Rec_Id);
   
      Insert Into ���˲������ͼ
         (����id, ����, ����, ����, �㼯, X1, Y1, X2, Y2, ���ɫ, ��䷽ʽ, ����ɫ, ����, �߿�)
         Select ����id, ����, ����, ����, �㼯, X1, Y1, X2, Y2, ���ɫ, ��䷽ʽ, ����ɫ, ����, �߿�
         From H���˲������ͼ
         Where ����id In (Select Id From H���˲������� Where ������¼id = n_Rec_Id);
   
      Insert Into ���˲����ⲿͼ
         (����id, ���, ͼ������, ͼ��·��, ͼ���ļ�)
         Select ����id, ���, ͼ������, ͼ��·��, ͼ���ļ�
         From H���˲����ⲿͼ
         Where ����id In (Select Id From H���˲������� Where ������¼id = n_Rec_Id);
   
      Insert Into ���˹�����¼
         (Id, ����id, ��ҳid, ��¼��Դ, ����id, ҩ��id, ҩ����, ��¼ʱ��, ��¼��, ���)
         Select Id, ����id, ��ҳid, ��¼��Դ, ����id, ҩ��id, ҩ����, ��¼ʱ��, ��¼��, ���
         From H���˹�����¼
         Where ����id In (Select Id From H���˲������� Where ������¼id = n_Rec_Id);
   
      Insert Into ������ϼ�¼
         (Id, ����id, ��ҳid, ��¼��Դ, ����id, �������, ����id, ���id, ֤��id, �������, ��Ժ���, �Ƿ�δ��, �Ƿ�����,
          ��¼����, ��¼��, ȡ��ʱ��, ȡ����, ҽ��id, ��ϴ���, �������)
         Select Id, ����id, ��ҳid, ��¼��Դ, ����id, �������, ����id, ���id, ֤��id, �������, ��Ժ���, �Ƿ�δ��, �Ƿ�����,
                ��¼����, ��¼��, ȡ��ʱ��, ȡ����, ҽ��id, ��ϴ���, �������
         From H������ϼ�¼
         Where ����id In (Select Id From H���˲������� Where ������¼id = n_Rec_Id);
   
      Insert Into ���˼�ϵͼ
         (����id, ���, ����, ĸ��, ����, ��ν, �Ա�, ״̬, ˵��, ����, ������ϵ, ������ϵ)
         Select ����id, ���, ����, ĸ��, ����, ��ν, �Ա�, ״̬, ˵��, ����, ������ϵ, ������ϵ
         From H���˼�ϵͼ
         Where ����id In (Select Id From H���˲������� Where ������¼id = n_Rec_Id);
   
      Insert Into ���������¼
         (Id, ����id, ��ҳid, ��¼��Դ, ����id, ��������, ������ʼʱ��, ��������ʱ��, ��������, ��������id, ������Ŀid, ��������,
          ����ҽʦ, ��һ����, �ڶ�����, ������ʿ, ����ʼʱ��, �������ʱ��, ����ʽ, ��������, ��������, ��Һ����, ����ҽʦ,
          ������ʼʱ��, ��������ʱ��, �п�, ����, ��¼����, ��¼��, ȡ��ʱ��, ȡ����)
         Select Id, ����id, ��ҳid, ��¼��Դ, ����id, ��������, ������ʼʱ��, ��������ʱ��, ��������, ��������id, ������Ŀid,
                ��������, ����ҽʦ, ��һ����, �ڶ�����, ������ʿ, ����ʼʱ��, �������ʱ��, ����ʽ, ��������, ��������,
                ��Һ����, ����ҽʦ, ������ʼʱ��, ��������ʱ��, �п�, ����, ��¼����, ��¼��, ȡ��ʱ��, ȡ����
         From H���������¼
         Where ����id In (Select Id From H���˲������� Where ������¼id = n_Rec_Id);
   
      Insert Into ����������ҩ
         (��¼id, ����, ���, ҩ��id, ҩ����, ����, ��ʽ)
         Select ��¼id, ����, ���, ҩ��id, ҩ����, ����, ��ʽ
         From H����������ҩ
         Where ��¼id In
               (Select Id From H���������¼ Where ����id In (Select Id From H���˲������� Where ������¼id = n_Rec_Id));
   
      Delete H���˲�����¼ Where Id = n_Rec_Id;
   End Zl_Retu_Cpr;

   --------------------------------------------
   --����ָ��ID�Ĳ���ҽ����¼�ӹ��̣�
   --------------------------------------------
   Procedure Zl_Retu_Order(n_Rec_Id H����ҽ����¼.Id%Type) As
   Begin
      Insert Into ����ҽ����¼
         (Id, ���id, ������Դ, ����id, ��ҳid, ����, �Ա�, ����, �Һŵ�, Ӥ��, ���˿���id, ���, ҽ��״̬, ҽ����Ч, �������, ������Ŀid,
          �걾��λ, �շ�ϸĿid, ��������, �ܸ�����, ҽ������, ҽ������, ִ�п���id, Ƥ�Խ��, ִ��Ƶ��, Ƶ�ʴ���, Ƶ�ʼ��,
          �����λ, ִ��ʱ�䷽��, �Ƽ�����, ִ������, ������־, ��ʼִ��ʱ��, ִ����ֹʱ��, �ϴ�ִ��ʱ��, ��������id, ����ҽ��,
          ����ʱ��, У�Ի�ʿ, У��ʱ��, ͣ��ҽ��, ͣ��ʱ��, ȷ��ͣ��ʱ��, ����id, ǰ��id, �Ƿ��ϴ�, ����,�����)
         Select Id, ���id, ������Դ, ����id, ��ҳid, ����, �Ա�, ����, �Һŵ�, Ӥ��, ���˿���id, ���, ҽ��״̬, ҽ����Ч, �������, ������Ŀid,
                �걾��λ, �շ�ϸĿid, ��������, �ܸ�����, ҽ������, ҽ������, ִ�п���id, Ƥ�Խ��, ִ��Ƶ��, Ƶ�ʴ���, Ƶ�ʼ��,
                �����λ, ִ��ʱ�䷽��, �Ƽ�����, ִ������, ������־, ��ʼִ��ʱ��, ִ����ֹʱ��, �ϴ�ִ��ʱ��, ��������id,
                ����ҽ��, ����ʱ��, У�Ի�ʿ, У��ʱ��, ͣ��ҽ��, ͣ��ʱ��, ȷ��ͣ��ʱ��, ����id, ǰ��id, �Ƿ��ϴ�, ����,�����
         From H����ҽ����¼
         Where Id = n_Rec_Id;
   
      Insert Into ����ҽ���Ƽ�
         (ҽ��id, �շ�ϸĿid, ����, ����, ����)
         Select ҽ��id, �շ�ϸĿid, ����, ����, ���� From H����ҽ���Ƽ� Where ҽ��id = n_Rec_Id;
      Delete H����ҽ���Ƽ� Where ҽ��id = n_Rec_Id;
   
      --����ҽ��ǩ������
      Insert Into ҽ��ǩ����¼(
		ID, ǩ������, ǩ����Ϣ, ֤��ID, ǩ��ʱ��, ǩ����)
      Select 
	    ID, ǩ������, ǩ����Ϣ, ֤��ID, ǩ��ʱ��, ǩ���� 
	  From Hҽ��ǩ����¼ 
	  Where ID IN (Select ǩ��ID From H����ҽ��״̬ Where ҽ��id = n_Rec_Id);
	  IF SQL%RowCount<>0 Then--����ת����ҽ��ʱ�Ѿ�ɾ��
          Delete Hҽ��ǩ����¼ Where ID IN (Select ǩ��ID From H����ҽ��״̬ Where ҽ��id = n_Rec_Id);
      End IF;
	  
      Insert Into ����ҽ��״̬
         (ҽ��id, ��������, ������Ա, ����ʱ��, ǩ��ID)
         Select ҽ��id, ��������, ������Ա, ����ʱ��, ǩ��ID From H����ҽ��״̬ Where ҽ��id = n_Rec_Id;
      Delete H����ҽ��״̬ Where ҽ��id = n_Rec_Id;
   
      --ҽ�����Ͳ���
      Insert Into ����ҽ������
         (ҽ��id, ���ͺ�, ��¼����, No, ��¼���, ��������, ������, ����ʱ��, �״�ʱ��, ĩ��ʱ��, ִ��״̬, ִ�в���id, �Ʒ�״̬,
          ����id, ִ�м�, ִ�й���, ������, ����ʱ��, ��������)
         Select ҽ��id, ���ͺ�, ��¼����, No, ��¼���, ��������, ������, ����ʱ��, �״�ʱ��, ĩ��ʱ��, ִ��״̬, ִ�в���id,
                �Ʒ�״̬, ����id, ִ�м�, ִ�й���, ������, ����ʱ��, ��������
         From H����ҽ������
         Where ҽ��id = n_Rec_Id;
   
      Insert Into ����ҽ������
         (ҽ��id, ���ͺ�, ��¼����, No)
         Select ҽ��id, ���ͺ�, ��¼����, No From H����ҽ������ Where ҽ��id = n_Rec_Id;
      Delete H����ҽ������ Where ҽ��id = n_Rec_Id;
   
      Insert Into ����ҽ��ִ��
         (ҽ��id, ���ͺ�, Ҫ��ʱ��, ��������, ִ��ժҪ, ִ����, ִ��ʱ��, �Ǽ���, �Ǽ�ʱ��)
         Select ҽ��id, ���ͺ�, Ҫ��ʱ��, ��������, ִ��ժҪ, ִ����, ִ��ʱ��, �Ǽ���, �Ǽ�ʱ��
         From H����ҽ��ִ��
         Where ҽ��id = n_Rec_Id;
      Delete H����ҽ��ִ�� Where ҽ��id = n_Rec_Id;
   
      Insert Into Ӱ�����¼
         (ҽ��id, ���ͺ�, Ӱ�����, ����, ����, Ӣ����, �Ա�, ����, ��������, ���, ����, ������, ���Ž�Ƭ, ���uid, λ��һ,
          λ�ö�, λ����, ����豸, ����ͼ��, ��������)
         Select ҽ��id, ���ͺ�, Ӱ�����, ����, ����, Ӣ����, �Ա�, ����, ��������, ���, ����, ������, ���Ž�Ƭ, ���uid,
                λ��һ, λ�ö�, λ����, ����豸, ����ͼ��, ��������
         From HӰ�����¼
         Where ҽ��id = n_Rec_Id;
      For r_Ris In (Select ���uid
                    From HӰ�����¼ r, H����ҽ������ s
                    Where r.ҽ��id = s.ҽ��id And r.���ͺ� = s.���ͺ� And s.ҽ��id = n_Rec_Id) Loop
         Insert Into Ӱ��������
            (����uid, ���uid, ���к�, ��������, �ɼ�ʱ��)
            Select ����uid, ���uid, ���к�, ��������, �ɼ�ʱ�� From HӰ�������� Where ���uid = r_Ris.���uid;
         For r_Seq In (Select ����uid From HӰ�������� Where ���uid = r_Ris.���uid) Loop
            Insert Into Ӱ����ͼ��
               (ͼ��uid, ����uid, ͼ���, ͼ������)
               Select ͼ��uid, ����uid, ͼ���, ͼ������ From HӰ����ͼ�� Where ����uid = r_Seq.����uid;
            Delete HӰ����ͼ�� Where ����uid = r_Seq.����uid;
         End Loop;
         Delete HӰ�������� Where ���uid = r_Ris.���uid;
      End Loop;
      Delete HӰ�����¼ Where ҽ��id = n_Rec_Id;
   
      Delete H����ҽ������ Where ҽ��id = n_Rec_Id;
   
      --������¼����
      Insert Into ����������¼
         (Id, ҽ��id, ����id, ����id, ��ҳid, ��¼��Դ, ��������, ������ʼʱ��, ��������ʱ��, ������ģ, ����ʼʱ��,
          �������ʱ��, ����ʽ, ��������, ��������, ��Һ����, ������ʼʱ��, ��������ʱ��, �п�, ����, �޾�����, ������, ������id,
          ������id, ��¼����, ��¼��, ȡ��ʱ��, ȡ����)
         Select Id, ҽ��id, ����id, ����id, ��ҳid, ��¼��Դ, ��������, ������ʼʱ��, ��������ʱ��, ������ģ, ����ʼʱ��,
                �������ʱ��, ����ʽ, ��������, ��������, ��Һ����, ������ʼʱ��, ��������ʱ��, �п�, ����, �޾�����, ������,
                ������id, ������id, ��¼����, ��¼��, ȡ��ʱ��, ȡ����
         From H����������¼
         Where ҽ��id = n_Rec_Id;
      For r_Ops In (Select Id From H����������¼ Where ҽ��id = n_Rec_Id) Loop
      
         Insert Into �����������
            (��¼id, ����, ȱʡ, ��������, ��������id, ������Ŀid)
            Select ��¼id, ����, ȱʡ, ��������, ��������id, ������Ŀid From H����������� Where ��¼id = r_Ops.Id;
         Delete H����������� Where ��¼id = r_Ops.Id;
      
         Insert Into ����������ҩ
            (��¼id, ����, ���, ҩƷid, ҩ��id, ҩƷ����, ׼������, ʹ������, ��ҩ��ʽ, ִ�п���id)
            Select ��¼id, ����, ���, ҩƷid, ҩ��id, ҩƷ����, ׼������, ʹ������, ��ҩ��ʽ, ִ�п���id
            From H����������ҩ
            Where ��¼id = r_Ops.Id;
         Delete H����������ҩ Where ��¼id = r_Ops.Id;
      
         Insert Into ����������Ա
            (��¼id, ����id, ��λ, ��Աid, ����, ����)
            Select ��¼id, ����id, ��λ, ��Աid, ����, ���� From H����������Ա Where ��¼id = r_Ops.Id;
         Delete H����������Ա Where ��¼id = r_Ops.Id;
      
         Insert Into ������������
            (��¼id, ���, ����id, ׼������, ʵ������, �����, ִ�п���id, ����˵��)
            Select ��¼id, ���, ����id, ׼������, ʵ������, �����, ִ�п���id, ����˵��
            From H������������
            Where ��¼id = r_Ops.Id;
         Delete H������������ Where ��¼id = r_Ops.Id;
      
         Insert Into ���������Ƽ�
            (��¼id, ���, ����, ����, �շ�ϸĿid, ִ�п���id)
            Select ��¼id, ���, ����, ����, �շ�ϸĿid, ִ�п���id From H���������Ƽ� Where ��¼id = r_Ops.Id;
         Delete H���������Ƽ� Where ��¼id = r_Ops.Id;
      End Loop;
      Delete H����������¼ Where ҽ��id = n_Rec_Id;
   
      Insert Into ��������״̬
         (ҽ��id, ���, �ϴ����, ��������, ��¼״̬, ������, ����ʱ��, ����˵��, ��ǰ״̬, ���ݺ�)
         Select ҽ��id, ���, �ϴ����, ��������, ��¼״̬, ������, ����ʱ��, ����˵��, ��ǰ״̬, ���ݺ�
         From H��������״̬
         Where ҽ��id = n_Rec_Id;
      Delete H��������״̬ Where ҽ��id = n_Rec_Id;
   
      --���鲿��
      Insert Into ����걾��¼
         (Id, ҽ��id, �걾���, ����ʱ��, ������, �걾����, ������, ����ʱ��, ����״̬, ������, ����ʱ��, �����, ���ʱ��,
          �ϲ������, ��ӡ����, ��������, ����id, ��������, ������, ��ע, δͨ�����ԭ��, ����ʱ��, �걾��̬, �Ƿ��ʿ�Ʒ,
          ִ�п���id)
         Select Id, ҽ��id, �걾���, ����ʱ��, ������, �걾����, ������, ����ʱ��, ����״̬, ������, ����ʱ��, �����, ���ʱ��,
                �ϲ������, ��ӡ����, ��������, ����id, ��������, ������, ��ע, δͨ�����ԭ��, ����ʱ��, �걾��̬, �Ƿ��ʿ�Ʒ,
                ִ�п���id
         From H����걾��¼
         Where ҽ��id = n_Rec_Id;
      For r_Retu In (Select Id From H������ͨ��� Where ����걾id In (Select Id From H����걾��¼ Where ҽ��id = n_Rec_Id)) Loop
         Insert Into ������ͨ���
            (Id, ����걾id, ������Ŀid, ������, �����־, ����ο�, �޸���, �޸�ʱ��, ��¼����, ԭʼ���, ԭʼ��¼ʱ��, ��¼��,
             �Ƿ����, �޸�ԭ��, ϸ��id, ����id, ��������)
            Select Id, ����걾id, ������Ŀid, ������, �����־, ����ο�, �޸���, �޸�ʱ��, ��¼����, ԭʼ���, ԭʼ��¼ʱ��,
                   ��¼��, �Ƿ����, �޸�ԭ��, ϸ��id, ����id, ��������
            From H������ͨ���
            Where Id = r_Retu.Id;
         Insert Into ����ҩ�����
            (ϸ�����id, ������id, �޸���, �޸�ʱ��, ���, �������, ��¼����, ����id)
            Select ϸ�����id, ������id, �޸���, �޸�ʱ��, ���, �������, ��¼����, ����id
            From H����ҩ�����
            Where ϸ�����id = r_Retu.Id;
         Delete H����ҩ����� Where ϸ�����id = r_Retu.Id;
         Delete H������ͨ��� Where Id = r_Retu.Id;
      End Loop;
      Delete H����걾��¼ Where ҽ��id = n_Rec_Id;
   
      Delete H����ҽ����¼ Where Id = n_Rec_Id;
   End Zl_Retu_Order;

   --------------------------------------------
   --����Ϊ��������
   --------------------------------------------
Begin
   --��ֹ���ݲ�һ�£��Ƚ���һЩԼ��
   Begin
	  Execute Immediate 'Alter Table H����ҽ��״̬ Modify Constraint H����ҽ��״̬_FK_ǩ��ID Disable';
   Exception
      When Others Then Null;
   End;

   If n_Flag = 0 Then
      Insert Into ���˹Һż�¼
         (Id, No, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��,
          ����Ա���, ����Ա����, ժҪ)
         Select Id, No, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��,
                �Ǽ�ʱ��, ����Ա���, ����Ա����, ժҪ
         From H���˹Һż�¼
         Where No = v_Times;
      For r_Cpr In (Select Id From H���˲�����¼ Where ����id = n_Patiid And �Һŵ� = v_Times) Loop
         Zl_Retu_Cpr(r_Cpr.Id);
      End Loop;
      For r_Order In (Select Id From H����ҽ����¼ Where ����id = n_Patiid And �Һŵ� = v_Times) Loop
         Zl_Retu_Order(r_Order.Id);
      End Loop;
      Delete H���˹Һż�¼ Where No = v_Times;
   Else
      For r_Cpr In (Select Id From H���˲�����¼ Where ����id = n_Patiid And ��ҳid = To_Number(v_Times)) Loop
         Zl_Retu_Cpr(r_Cpr.Id);
      End Loop;
      For r_Order In (Select Id From H����ҽ����¼ Where ����id = n_Patiid And ��ҳid = To_Number(v_Times)) Loop
         Zl_Retu_Order(r_Order.Id);
      End Loop;
      Update ������ҳ Set ����ת�� = 0 Where ����id = n_Patiid And ��ҳid = To_Number(v_Times);
   End If;

   --����Լ��
   Begin
	  Execute Immediate 'Alter Table H����ҽ��״̬ Modify Constraint H����ҽ��״̬_FK_ǩ��ID Enable';
   Exception
      When Others Then Null;
   End;

   Commit;
End Zl_Retu_Clinic;
/

CREATE OR REPLACE Procedure ZL_����ҽ����¼_����(
--���ܣ�����ָ����ҽ��(δ���͵ĳ���������)
--˵����һ����ҩ��ֻ�ܵ���һ��(������ʾ�ж���)
--������ID_IN=���IDΪNULL��ҽ����ID(��ҩ;��,��ҩ�÷�,�����Ŀ,��Ҫ����,������ҽ��)
    ID_IN    ����ҽ����¼.ID%TYPE
) IS
    --����ҽ�������Ϣ
    Cursor c_Advice is
        Select A.����ID,A.�Һŵ�,A.��ҳID,A.Ӥ��,A.ҽ��״̬,
			A.�ϴ�ִ��ʱ��,A.ҽ������,A.�������,B.��������,A.ִ�п���ID
        From ����ҽ����¼ A,������ĿĿ¼ B
        Where A.������ĿID=B.ID And A.ID=ID_IN;
    r_Advice c_Advice%RowType;
	
	--����ҽ�����ϣ�
    --����ҽ��������NO������λ���Ҫ���ʻ�ɾ��(���ﻮ�۵�)�ķ��ü�¼
    --һ��ҽ�������Ƕ���д�˷��ͼ�¼,�ҿ���NO��ͬ(ҩƷ��,�÷��巨��һ����)
    --���ܷ��ͼ�¼�ļƷ�״̬(��������Ʒ�),�з��ü�¼��Ȼ��������
    --����ֻ��۸񸸺�Ϊ�յ�,�Ա�ȡ�������(������ʵ�)
    --ֻ�ܼ�¼״̬Ϊ1�ļ�¼,����Ѿ����ʻ򲿷����ʵļ�¼,���ٴ���
    --ZYL:�����ҩƷ����ʹ�Ѿ��շѺͷ�ҩ,��Ȼ��������
    Cursor c_RollMoney(v_���ͺ� ����ҽ������.���ͺ�%Type) is
        Select A.��¼����,A.��¼״̬,A.NO,A.���,
			A.ִ��״̬ as ����ִ��,C.ִ��״̬ as ҽ��ִ��
        From ���˷��ü�¼ A,����ҽ����¼ B,����ҽ������ C,������ĿĿ¼ I
        Where C.ҽ��ID=B.ID And C.���ͺ�=v_���ͺ�
            And (B.ID=ID_IN Or B.���ID=ID_IN)
            And A.ҽ�����=B.ID And A.��¼״̬ IN(0,1)
            And A.NO=C.NO And A.��¼����=C.��¼����
            And B.������ĿID=I.ID And A.�۸񸸺� IS NULL
            And (
				A.�շ���� Not In ('5','6','7','E')
                Or A.�շ����='E' And I.�������� Not In ('2','3','4')
                Or A.�շ���� In ('5','6','7') And Nvl(A.ִ��״̬,0)=0 And Not(A.��¼����=1 And A.��¼״̬<>0)
				Or Exists(Select ����ֵ From ϵͳ������ Where ������=68 And Nvl(����ֵ,0)=0)
				)
        Order BY A.�շ�ϸĿID;

	--����ɾ�������¼
	Cursor c_Case is
		Select ����ID From ����ҽ������ 
		Where ����ID IS Not NULL And ҽ��ID IN(
			Select ID From ����ҽ����¼ Where ID=ID_IN OR ���ID=ID_IN);
	r_Case c_Case%RowType;

    v_���ͺ�        ����ҽ������.���ͺ�%Type;
    v_����NO        ���˷��ü�¼.NO%Type;
    v_��¼����      ���˷��ü�¼.��¼����%Type;
    v_�������      Varchar2(255);

    v_Date          Date;
	v_Count			Number;
    v_Temp          Varchar2(255);
    v_��Ա���      ���˷��ü�¼.����Ա���%Type;
    v_��Ա����      ���˷��ü�¼.����Ա����%Type;

    v_Error			Varchar2(255);
    Err_Custom      Exception;
Begin
    --���ҽ��״̬�Ƿ���ȷ:��������
    Open c_Advice;
    Fetch c_Advice Into r_Advice;

    If r_Advice.�Һŵ� IS NULL Then
        IF r_Advice.ҽ��״̬ IN(4,8,9) Then
            v_Error:='ҽ��"'||r_Advice.ҽ������||'"�Ѿ������ϻ�ֹͣ�����������ϡ�';
            Raise Err_Custom;
        Elsif r_Advice.�ϴ�ִ��ʱ�� IS Not NULL Then
            v_Error:='ҽ��"'||r_Advice.ҽ������||'"�Ѿ����ͣ����ܱ����ϡ�';
            Raise Err_Custom;
        End IF;
    Else
        IF r_Advice.ҽ��״̬<>8 Then
            v_Error:='ҽ��"'||r_Advice.ҽ������||'"��δ���ͻ��Ѿ����ϡ�';
            Raise Err_Custom;
        End IF;

		--����ҽ��ֻ���ܷ���һ��
		Select Count(*) Into v_Count 
		From ����ҽ������ 
		Where ִ��״̬ IN(1,3) And ҽ��ID IN(
			Select ID From ����ҽ����¼ Where ID=ID_IN Or ���ID=ID_IN);
		If v_Count>0 Then
			v_Error:='��ҽ���Ѿ�ִ�л�����ִ�У��������ϡ�';
			Raise Err_Custom;
		End IF;
    End IF;

    --��ǰ������Ա
    v_Temp:=zl_Identity;
    v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
    v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
    v_��Ա���:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
    v_��Ա����:=Substr(v_Temp,Instr(v_Temp,',')+1);
    Select Sysdate Into v_Date From Dual;

    Update ����ҽ����¼
        Set ҽ��״̬=4
    Where ID=ID_IN Or ���ID=ID_IN;

    Insert Into ����ҽ��״̬(
        ҽ��ID,��������,������Ա,����ʱ��)
    Select
        ID,4,v_��Ա����,v_Date
    From ����ҽ����¼
    Where ID=ID_IN Or ���ID=ID_IN;

    --��������
    ---------------------------------------------------------------------------------------
    --����/סԺҽ������ʱ�Ѷ�Ӧ�����뵥����
	Update ���˲�����¼
		Set ������=v_��Ա����,��������=v_Date
	Where ������ IS NULL And ID IN(
		Select ����ID From ����ҽ����¼ Where ID=ID_IN Or ���ID=ID_IN);

    If r_Advice.�Һŵ� IS Not NULL Then
        --����ҽ��(����)����ʱ����Ҫ�����������:ֻ��һ�η���
        --���˻��ۻ���ʷ���
        Begin
            --ҽ��IDΪ����ֵ������ҽ����һ�������˵�,�����޷��͡�
            Select Distinct ���ͺ� Into v_���ͺ� From ����ҽ������
            Where ҽ��ID IN(Select ID From ����ҽ����¼ Where ID=ID_IN Or ���ID=ID_IN);
        Exception
            When Others Then v_���ͺ�:=NULL;
        End;
        If v_���ͺ� IS Not NULL Then
            --������ҽ���ķ���ɾ��������(��һ��ҽ�������в�ͬNO����)
            --������ʣ����ԭʼ�����ѱ�����(�򲿷�����),���ù��������ж�
            --���ﻮ�ۣ�������շѣ�������ɾ��
            v_����NO:=NULL;v_�������:=NULL;
            For r_RollMoney In c_RollMoney(v_���ͺ�) Loop
                If Nvl(v_����NO,'��')<>r_RollMoney.NO Then
                    If v_������� IS Not NULL And v_����NO IS Not NULL Then
                        v_�������:=Substr(v_�������,2);
                        IF v_��¼����=1 Then
                            zl_���ﻮ�ۼ�¼_Delete(v_����NO,v_�������);
                        Elsif v_��¼����=2 Then
                            zl_������ʼ�¼_Delete(v_����NO,v_�������,v_��Ա���,v_��Ա����);
                        End If;
                    End IF;
                    v_�������:=NULL;
                End IF;
                v_��¼����:=r_RollMoney.��¼����;
                v_����NO:=r_RollMoney.NO;
                v_�������:=v_�������||','||r_RollMoney.���;

                If Nvl(r_RollMoney.ҽ��ִ��,0) IN(1,3) Then --1-��ȫִ��;3-����ִ��
                    v_Error:='ҽ��"'||r_Advice.ҽ������||'"�Ѿ�ִ�л�����ִ�У��������ϡ�';
                    Raise Err_Custom;
                End IF;
                If Nvl(r_RollMoney.����ִ��,0) IN(1,2) Then --1-��ȫִ��;2-����ִ��
                    v_Error:='ҽ�����õ���"'||r_RollMoney.NO||'"�е������Ѿ�ȫ���򲿷�ִ�У��������ϡ�';
                    Raise Err_Custom;
                End IF;
                If r_RollMoney.��¼����=1 And r_RollMoney.��¼״̬<>0 Then
                    v_Error:='ҽ�����õ���"'||r_RollMoney.NO||'"�Ѿ��շѣ��������ϡ�';
                    Raise Err_Custom;
                End IF;
            End Loop;
            If v_������� IS Not NULL And v_����NO IS Not NULL Then
                v_�������:=Substr(v_�������,2);
                IF v_��¼����=1 Then
                    zl_���ﻮ�ۼ�¼_Delete(v_����NO,v_�������);
                Elsif v_��¼����=2 Then
                    zl_������ʼ�¼_Delete(v_����NO,v_�������,v_��Ա���,v_��Ա����);
                End If;
            End IF;
			
			Open c_Case;--�����ȴ�

            --����ҽ�����ͼ�¼
            Delete From ����ҽ������ Where ҽ��ID IN(
				Select ID From ����ҽ����¼ Where ID=ID_IN Or ���ID=ID_IN);

            --ɾ����Ӧ�ı��浥
			Fetch c_Case Into r_Case;
			While c_Case%Found Loop
	            Delete From ���˲�����¼ Where ID=r_Case.����ID;
				Fetch c_Case Into r_Case;
			End Loop;
			Close c_Case;

            --��������ҽ���Ĵ���
            If r_Advice.�������='Z' And Nvl(r_Advice.��������,'0')<>'0' And Nvl(r_Advice.Ӥ��,0)=0 Then
                If r_Advice.��������='1' And r_Advice.ִ�п���ID IS Not NULL Then
                    --����ҽ��
					Select Count(*) Into v_Count From ������ҳ Where ����ID=r_Advice.����ID And Nvl(��ҳID,0)=0 And ��Ժ����ID=r_Advice.ִ�п���ID And �������� IN(1,2);
					If v_Count=1 Then
						zl_��Ժ������ҳ_Delete(r_Advice.����ID,0);
					End IF;
                ElsIf r_Advice.��������='2' And r_Advice.ִ�п���ID IS Not NULL Then
                    --סԺҽ��
					Select Count(*) Into v_Count From ������ҳ Where ����ID=r_Advice.����ID And Nvl(��ҳID,0)=0 And ��Ժ����ID=r_Advice.ִ�п���ID And Nvl(��������,0)=0;
					If v_Count=1 Then
						zl_��Ժ������ҳ_Delete(r_Advice.����ID,0);
					End IF;
                End IF;
            End IF;
        End IF;
    End IF;
	
	--ɾ�������ǼǼ�¼
	If r_Advice.�������='E' And r_Advice.��������='1'  Then

		--Update ����ҽ����¼ Set Ƥ�Խ��=Null Where ID=ID_IN; --��������Ƥ�Խ��

		For r_Test IN(Select ����ʱ�� From ����ҽ��״̬ Where ҽ��ID=ID_IN And ��������=10) Loop
			Delete From ���˹�����¼ 
			Where ����ID=r_Advice.����ID And ��¼��Դ=2
				And Nvl(��ҳID,0)=Nvl(r_Advice.��ҳID,0)
				And ��¼ʱ��=r_Test.����ʱ��;
		End Loop;
	End IF;

    Close c_Advice;
Exception
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_����ҽ����¼_����;
/

-------------------------------------------------------------------------
--���˺�:�Գ�����¼���з�Ʊ��Ϣ�޸ģ������޸ķ�Ʊ���.

-- �����⹺��ⷢƱ��Ϣ���޸�
CREATE OR REPLACE PROCEDURE zl_�����⹺��Ʊ��Ϣ_UPDATE (
    NO_IN		IN ҩƷ�շ���¼.NO%TYPE := NULL,
    ��¼״̬_IN		IN ҩƷ�շ���¼.��¼״̬%type:=NULL,
    ���_IN		IN ҩƷ�շ���¼.���%TYPE:=NULL,
    ��Ʊ��_IN		IN Ӧ����¼.��Ʊ��%TYPE := NULL,
    ��Ʊ����_IN		IN Ӧ����¼.��Ʊ����%TYPE := NULL,
    ��Ʊ���_IN		IN Ӧ����¼.��Ʊ���%TYPE := NULL,
    ��ҩ��λ_IN		in Ӧ����¼.��λID%TYPE:=0
)
IS
    mErrMsg		varchar2(255);
    mErrItem		exception;

    V_NO		Ӧ����¼.NO%TYPE;
    V_Ӧ��ID		Ӧ����¼.ID%TYPE;
    V_�շ�ID		Ӧ����¼.�շ�ID%TYPE;
    V_�������		Ӧ����¼.�������%TYPE;
    V_��Ʊ���		Ӧ����¼.��Ʊ���%TYPE;--�ɷ�Ʊ���
    V_��ҩ��λID	Ӧ����¼.��λID%TYPE;
BEGIN
    --ȡ�Ƿ񸶿�ܶ�
    BEGIN 
        Select max(�������),sum(nvl(��Ʊ���,0)) INTO v_�������,v_��Ʊ��� 
        FROM Ӧ����¼ 
        WHERE �շ�id=(Select ID From ҩƷ�շ���¼ Where NO=NO_IN And ���=���_IN And ����=15) 
		AND ϵͳ��ʶ=5 And ��¼����=-1;
    EXCEPTION 
        WHEN OTHERS THEN 
        v_��Ʊ���:=0;
    END ;


    v_�������:=nvl(v_�������,0);
    
    IF v_�������<>0 then
       mErrMsg:='[ZLSOFT]�õ����Ѿ������˿�������޸ķ�Ʊ��Ϣ[ZLSOFT]';
       RAISE mErrItem;
    END IF ;

    if ��Ʊ���_IN>v_��Ʊ��� And v_��Ʊ���<>0 then 
        mErrMsg:='[ZLSOFT]��Ʊ����С�ڼƻ�������[ZLSOFT]';
        raise mErrItem;
    end if ;

    
    --�ж��Ƿ������ļ�¼
    IF ��¼״̬_IN<>1 THEN 
	IF nvl(��Ʊ��_IN,' ') <>' ' AND ��Ʊ���_IN=0  THEN 
	       mErrMsg:='[ZLSOFT]���ܶԷ�Ʊ���Ϊ��ķ�Ʊ��Ϣ�����޸ġ�[ZLSOFT]';
	       RAISE mErrItem;
	END IF ;

	IF nvl(��Ʊ��_IN,' ') =' ' AND ��Ʊ���_IN<>0  THEN 
	       mErrMsg:='[ZLSOFT]���ܶԳ����򱻳�����¼�ķ�Ʊ�Ÿ�Ϊ��,���ܱ��棡[ZLSOFT]';
	       RAISE mErrItem;
	END IF ;

	--������صķ�Ʊ��Ϣ,ֻ���ķ�Ʊ�ţ���Ʊ����
	FOR V_�շ� IN (Select ID From ҩƷ�շ���¼ WHERE ����=15 AND NO=NO_IN AND ���=���_IN )
	LOOP 
	    UPDATE Ӧ����¼
	    SET ��Ʊ�� = ��Ʊ��_IN,
		��Ʊ���� = ��Ʊ����_IN
	    WHERE �շ�ID = V_�շ�.ID And ϵͳ��ʶ=5 AND ��¼����=0;
	END LOOP ;
	RETURN ;
	
    END IF ;

    SELECT A.ID,nvl(B.��Ʊ���,0),A.��ҩ��λID    INTO V_�շ�ID,V_��Ʊ���,V_��ҩ��λID
    FROM ҩƷ�շ���¼ A,(Select * From Ӧ����¼ Where ϵͳ��ʶ=5  AND ��¼����<>-1 And ������� Is NULL) B
    WHERE A.ID = B.�շ�ID(+)
        AND A.NO = NO_IN
        AND A.���� = 15
        AND A.��¼״̬ = 1
        AND A.��� = ���_IN; 
    
    UPDATE Ӧ����¼
    SET ��Ʊ�� = ��Ʊ��_IN,
        ��Ʊ���� = ��Ʊ����_IN,
        ��Ʊ��� = ��Ʊ���_IN,
        ��λID=��ҩ��λ_IN
    WHERE �շ�ID = V_�շ�ID And ϵͳ��ʶ=5 And ��¼״̬=1 And ��¼����=0;


    if sql%rowcount=0 then 
        IF ��Ʊ��_IN IS NOT NULL THEN
            --����ǵ�һ����ϸ,�����Ӧ����¼��NO
            BEGIN 
                SELECT NO INTO V_NO FROM Ӧ����¼ 
                WHERE ϵͳ��ʶ=5 AND ��¼����=0 AND ��¼״̬=1 
                    AND ��ⵥ�ݺ�=NO_IN AND ROWNUM<2;
            EXCEPTION
                WHEN OTHERS THEN V_NO:=NEXTNO(69);
            END ;

            SELECT Ӧ����¼_ID.NEXTVAL INTO V_Ӧ��ID FROM DUAL;
            
            INSERT INTO Ӧ����¼
            (ID,��¼����,��¼״̬,��Ŀid,���,��λID,NO,ϵͳ��ʶ,�շ�ID,��ⵥ�ݺ�,���ݽ��,��Ʊ��,��Ʊ����,��Ʊ���,Ʒ��,
            ���,����,����,������λ,����,�ɹ���,�ɹ����,������,��������,�����,�������,ժҪ)
            select V_Ӧ��ID,0,1,A.ҩƷid,A.���,��ҩ��λ_IN,V_NO,5,V_�շ�ID,A.NO,A.���۽��,��Ʊ��_IN,��Ʊ����_IN,��Ʊ���_IN,B.����,
            B.���,B.����,A.����,B.���㵥λ,A.ʵ������,A.�ɱ���,A.�ɱ����,A.������,A.��������,A.�����,A.�������,A.ժҪ
            from ҩƷ�շ���¼ A,�շ���ĿĿ¼ B
            Where A.����=15 And A.NO=NO_in And A.���=���_IN And A.ҩƷID=B.ID;
        END IF;
    END IF;

    UPDATE Ӧ����� SET ��� = NVL (���,0) - V_��Ʊ���
    WHERE ��λID = V_��ҩ��λID AND ���� = 1;
    IF SQL%NOTFOUND THEN
        INSERT INTO Ӧ�����(��λID,����,���) VALUES (V_��ҩ��λID,1,-V_��Ʊ���);
    END IF; 
    UPDATE Ӧ����� SET ���=NVL(���,0)+��Ʊ���_IN
    WHERE ��λID=��ҩ��λ_IN AND ����=1;

    IF SQL%NOTFOUND THEN
        INSERT INTO Ӧ�����(��λID,����,���) VALUES (��ҩ��λ_IN,1,��Ʊ���_IN);
    END IF; 

    --����ҩƷ�շ���¼�еĹ�ҩ��λ
    UPDATE ҩƷ�շ���¼ SET ��ҩ��λID=��ҩ��λ_IN WHERE NO=NO_IN AND ����=15 And ���=���_IN;

    --����ҩƷ�������ϴι�Ӧ��
    UPDATE ҩƷ��� SET �ϴι�Ӧ��ID=��ҩ��λ_IN WHERE ����=1 and  (�ⷿID,ҩƷID) IN (SELECT �ⷿID,ҩƷID FROM ҩƷ�շ���¼ WHERE NO=NO_IN AND ����=15);
EXCEPTION
    when mErrItem then Raise_application_error (-20101,mErrMsg); 
    WHEN NO_data_found THEN
        Raise_application_error (-20101,'[ZLSOFT]�õ����Ѿ������˳������Ѿ������[ZLSOFT]'); 
    WHEN OTHERS THEN
    zl_ErrorCenter (SQLCODE,SQLERRM);
END zl_�����⹺��Ʊ��Ϣ_UPDATE;
/



CREATE OR REPLACE PROCEDURE zl_��������ԭ�ϳ���_Insert (
		NO_IN		IN ҩƷ�շ���¼.NO%TYPE,
		�Է�����ID_IN	IN ҩƷ�շ���¼.�Է�����ID%TYPE)
AS
		mErrMsg         varchar2(100);
		mErrItem        EXCEPTION ;

		v_����		ҩƷ�շ���¼.ʵ������%type;
		v_�ɱ���	ҩƷ�շ���¼.�ɱ���%type;
		v_�ɱ����	ҩƷ�շ���¼.�ɱ����%type;
		v_���		ҩƷ�շ���¼.���%type;

		v_�ۼ�		ҩƷ�շ���¼.���ۼ�%type;
		v_���۽��	ҩƷ�շ���¼.���۽��%type;

		v_�����	ҩƷ���.ʵ�ʽ��%type;
		v_�����	ҩƷ���.ʵ�ʲ��%type;
		v_��������	ҩƷ���.��������%type;
		v_ʵ������	ҩƷ���.ʵ������%type;
		v_�ϴβ���	ҩƷ���.�ϴβ���%type;
		v_���ɱ�����	ϵͳ������.����ֵ%type;
		V_�������ID	ҩƷ�շ���¼.������ID%TYPE;--������ID
		V_maxserial	ҩƷ�շ���¼.���%TYPE;
BEGIN
	SELECT B.ID INTO V_�������ID    
	FROM ҩƷ�������� A,ҩƷ������ B
	WHERE A.���ID = B.ID AND A.���� = 31 AND B.ϵ�� = -1
		AND ROWNUM < 2;
	
	SELECT ����ֵ INTO v_���ɱ����� FROM ϵͳ������ WHERE ������=120;

	SELECT MAX (���) INTO V_maxserial
	FROM ҩƷ�շ���¼
	WHERE NO = NO_IN AND ���� = 16 AND ���ϵ�� = 1;


	FOR v_���� IN (SELECT * FROM ҩƷ�շ���¼ WHERE NO = NO_IN AND ���� = 16 AND ���ϵ�� = 1)
	LOOP 
		FOR v_��� IN (	Select a.*,b.�Ƿ���,c.ָ�������,c.�ɱ���
				From ���Ʋ��Ϲ��� a,�շ���ĿĿ¼ b,�������� c
				WHERE  a.ԭ�ϲ���id=b.id  AND  a.���Ʋ���ID=v_����.ҩƷid AND a.ԭ�ϲ���id=c.����id
			)
		LOOP 
			BEGIN 
				SELECT ��������,ʵ������,ʵ�ʲ��,ʵ�ʽ��,�ϴβ��� 
					INTO v_��������,v_ʵ������,v_�����,v_�����,v_�ϴβ���
				FROM ҩƷ���
				WHERE ҩƷid=v_���.ԭ�ϲ���id AND ���� = 1 AND �ⷿID=�Է�����ID_IN;
			EXCEPTION 
				WHEN OTHERS THEN 
				     v_��������:=0;
				     v_ʵ������:=0;
				     v_�����:=0;
				     v_�����:=0;
			END ;
			IF nvl(v_���.�Ƿ���,0)=1 THEN 
				--ʵ��
				IF nvl(v_ʵ������,0)> 0 THEN 
					v_�ۼ�:=nvl(v_�����,0)/v_ʵ������;
				ELSE	
					--�޿���:����ʾ
					mErrMsg:='[ZLSOFT]�õ����д���һ������ԭ�ϵ�ʵ����������[ZLSOFT]';
					RAISE mErrItem;
				END IF ;
			ELSE 
				--����,���ּ�Ϊ׼
				BEGIN 
					SELECT nvl(�ּ�,0) INTO v_�ۼ�  
					FROM �շѼ�Ŀ 
					WHERE  �շ�ϸĿID=v_���.ԭ�ϲ���id AND ((SYSDATE BETWEEN ִ������ AND ��ֹ����) OR (SYSDATE >= ִ������ AND ��ֹ���� IS NULL));

				EXCEPTION 
					WHEN OTHERS THEN mErrMsg:='Err';
				END ;
				IF mErrMsg='Err' THEN 
					mErrMsg:='[ZLSOFT]�õ����д���һ������ԭ�ϻ�δ���ж��ۣ�[ZLSOFT]';
					RAISE mErrItem;
				END IF ;
			END IF ;
			v_����:=nvl(v_����.ʵ������,0)*v_���.����/v_���.��ĸ;
			
			If v_����=0 Then 
				mErrMsg:='[ZLSOFT]�õ����д���һ������ԭ�ϵ�����Ϊ���ˣ�[ZLSOFT]';
				RAISE mErrItem;
			End If ;
			v_���۽��:=v_����*v_�ۼ�;

			--��ɱ���
			IF nvl(v_�����,0)<=0 THEN 
				IF v_���ɱ�����='1' AND nvl(v_���.�ɱ���,0)>0 THEN 
					v_�ɱ���:=v_���.�ɱ���;
					v_���:=v_���۽��-v_����*v_�ɱ���;
				ELSE 
					v_���:=v_���۽��*v_���.ָ�������/100;
					v_�ɱ���:=(v_���۽��-v_���)/v_���� ;
				END if;
			ELSE 
				v_��� := v_���۽�� * (v_����� / v_�����);
				v_�ɱ��� := (v_���۽�� - v_���) / v_����;
			END IF ;
			v_�ɱ���:=nvl(v_�ɱ���,0);
			v_�ɱ����:=v_�ɱ���*v_����;
			V_maxserial:=V_maxserial+1;

			Insert INTO ҩƷ�շ���¼
			    (ID,��¼״̬,����,NO,���,�ⷿID,�Է�����ID,������ID,���ϵ��,ҩƷID,����,��д����,ʵ������,�ɱ���,�ɱ����,
			    ���ۼ�,���۽��,���,ժҪ,������,��������,����ID,����)
			VALUES (
			    ҩƷ�շ���¼_ID.Nextval,1,16,NO_IN,V_maxserial,v_����.�Է�����ID,v_����.�ⷿID,
			    V_�������ID,-1,v_���.ԭ�ϲ���id,v_�ϴβ���,v_����,v_����,
			    v_�ɱ���,v_�ɱ����,v_�ۼ�,v_���۽��,v_���,v_����.ժҪ,
			    v_����.������,v_����.��������,v_����.ҩƷID,v_����.���);

			--IF v_��������<0 then
			--    mErrMsg:='[ZLSOFT]�õ����д���һ������ԭ�ϵĿ�����������[ZLSOFT]';
			--    RAISE mErrItem;
			--END IF ;

			UPDATE ҩƷ���
			SET �������� = NVL (��������,0) - v_����
			WHERE �ⷿID = v_����.�Է�����ID AND ҩƷID = v_���.ԭ�ϲ���id AND ���� = 1;

			IF SQL%NOTFOUND THEN
			    Insert INTO ҩƷ��� (�ⷿID,ҩƷID,����,��������)
			    VALUES (v_����.�Է�����ID,v_���.ԭ�ϲ���id,1,-v_����);
			END IF;

			DELETE
			FROM ҩƷ���
			WHERE �ⷿID=v_����.�Է�����ID And ҩƷID=v_���.ԭ�ϲ���id
				And nvl(��������,0)=0 And nvl(ʵ������,0)=0
				And nvl(ʵ�ʽ��,0)=0 And nvl(ʵ�ʲ��,0)=0;

		END LOOP ;
	END LOOP ;
EXCEPTION
    WHEN mErrItem  THEN
        Raise_application_error (-20101,mErrMsg    );
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE,SQLERRM);
END zl_��������ԭ�ϳ���_Insert;
/



-----------------------------------------------------------
-- ��������������˴���
--˵�������ȶ�ҩƷ�շ���¼���е�����˺����ʱ����д���
--���Ŷ�ҩƷ����ҩƷ�շ����ܱ��е���Ӧ�����ͽ����д���
--�ر�˵������ҩƷ���Ĵ����Ƿֿ��ģ�������ҩƷ�Ĵ������������һ����
--��ԭ�ϵĴ���ֻ��ʵ��������ʵ�ʽ�ʵ�ʲ�۽��м��ٵĴ������Կ����������д�����Ϊ��ǰ����ʱ�Ѵ���
--�������
--���룺NO_IN,�����_IN
------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_���Ʋ������_verify (
    NO_IN        IN ҩƷ�շ���¼.NO%TYPE := NULL,
    �����_IN    IN ҩƷ�շ���¼.�����%TYPE := NULL
)
IS
	mErrItem		EXCEPTION;
	mErrMsg			varchar2(100);

	V_ʵ�ʿ����		ҩƷ���.ʵ�ʽ��%TYPE;
	V_ʵ�ʿ����		ҩƷ���.ʵ�ʲ��%TYPE;
	V_������		ҩƷ���.ʵ�ʲ��%TYPE;
	V_�����		number(18,8);
	V_�ɱ���		ҩƷ�շ���¼.�ɱ���%TYPE;
	V_�ɱ����		ҩƷ�շ���¼.�ɱ����%TYPE;
	v_���ɱ�����		ϵͳ������.����ֵ%type;
	V_С��			number(2);

BEGIN

    	SELECT ����ֵ INTO v_���ɱ����� FROM ϵͳ������ WHERE ������=120;


	UPDATE ҩƷ�շ���¼
	SET ����� = NVL (�����_IN,�����),������� = SYSDATE
	WHERE NO = NO_IN AND ���� = 16    AND ��¼״̬ = 1     AND ����� IS NULL;

	IF SQL%ROWCOUNT = 0 THEN
		mErrMsg:='[ZLSOFT]�õ����Ѿ���������ˣ�[ZLSOFT]';    
		RAISE mErrItem;
	END IF;

	BEGIN 
	    SELECT  to_number(Nvl(����ֵ,ȱʡֵ),'99999')  INTO v_С�� 
	    From ϵͳ������ where ������='���ý���λ��';
	EXCEPTION 
	WHEN OTHERS THEN 
		v_С��:=2;		
	END;


	FOR V_ҩƷ�շ���¼ IN (
			SELECT ID,ʵ������,���۽��,���,�ⷿID,ҩƷID,����,�ɱ���,����,Ч��,����,������ID,���ϵ��,�Է�����ID
			FROM ҩƷ�շ���¼
			WHERE NO = NO_IN AND ���� = 16 AND ��¼״̬ = 1    
			ORDER BY ҩƷID    ) 
	LOOP
		--����ҩƷ�������Ӧ����
		IF V_ҩƷ�շ���¼.���ϵ�� = -1 THEN
		    BEGIN
			SELECT nvl(ʵ�ʽ��,0),nvl(ʵ�ʲ��,0) INTO V_ʵ�ʿ����,V_ʵ�ʿ����
			FROM ҩƷ���
			WHERE ҩƷID = V_ҩƷ�շ���¼.ҩƷID
			    AND NVL (����,0) = NVL (V_ҩƷ�շ���¼.����,0)
			    AND �ⷿID = V_ҩƷ�շ���¼.�ⷿID
			    AND ���� = 1
			    AND ROWNUM = 1;
		    EXCEPTION
			WHEN OTHERS THEN
			    V_ʵ�ʿ���� := 0;
		    END;

		    IF V_ʵ�ʿ���� <= 0 THEN
			BEGIN
			    SELECT ָ������� / 100 INTO V_�����
			    FROM ��������
			    WHERE ����ID = V_ҩƷ�շ���¼.ҩƷID;
			EXCEPTION
			    WHEN OTHERS THEN
			    V_����� := 0;
			END;
			IF v_���ɱ����� ='1' THEN 
				BEGIN 
					SELECT nvl(�ɱ���,0) INTO v_�ɱ��� FROM �������� WHERE ����id=v_ҩƷ�շ���¼.ҩƷid;
				EXCEPTION 
					WHEN OTHERS THEN v_�ɱ���:=0;
				END ;
				IF v_�ɱ���=0 THEN 
					V_������ := round(V_ҩƷ�շ���¼.���۽�� * V_�����,4);
				ELSE 
					V_������ :=round(v_ҩƷ�շ���¼.���۽��-v_ҩƷ�շ���¼.ʵ������*v_�ɱ���,4);
				END IF ;
			ELSE 
				V_������ :=round( V_ҩƷ�շ���¼.���۽�� * V_�����,V_С��);
			END IF ;
		    ELSE
			V_����� := V_ʵ�ʿ���� / V_ʵ�ʿ����;
			V_������ :=round( V_ҩƷ�շ���¼.���۽�� * V_�����,V_С��);
		    END IF;


		    IF NVL (V_ҩƷ�շ���¼.ʵ������,0) = 0 THEN
			V_�ɱ��� := (V_ҩƷ�շ���¼.���۽�� - V_������);
		    ELSE
			V_�ɱ��� :=(V_ҩƷ�շ���¼.���۽�� - V_������) / V_ҩƷ�շ���¼.ʵ������;
		    END IF;
  
		    v_�ɱ���:=nvl(v_�ɱ���,0);
		    V_�ɱ���� := round(V_�ɱ��� * V_ҩƷ�շ���¼.ʵ������,V_С��);
		ELSE
		    BEGIN
			SELECT SUM(�ɱ���) INTO V_�ɱ���
			FROM (
			    SELECT DECODE(SIGN(NVL(C.ʵ�ʽ��,0)),1,(D.�ּ�-D.�ּ�*(C.ʵ�ʲ��/C.ʵ�ʽ��)),(D.�ּ�-D.�ּ�*(B.ָ�������/100)))
				*(A.����/A.��ĸ) AS �ɱ���
			    FROM ���Ʋ��Ϲ��� A,
				(SELECT j.����id AS ҩƷid,j.ָ������� FROM �������� j,�շ���ĿĿ¼ q  WHERE j.����id=q.id and  nvl(q.�Ƿ���,0)=0) B,
				(SELECT �ⷿID,ҩƷID,ʵ�ʽ��,ʵ�ʲ�� FROM ҩƷ��� WHERE ���� = 1 AND �ⷿID = V_ҩƷ�շ���¼.�Է�����ID ) C,
				(SELECT �շ�ϸĿID,�ּ� FROM �շѼ�Ŀ WHERE ((SYSDATE BETWEEN ִ������ AND ��ֹ����) OR (SYSDATE >= ִ������ AND ��ֹ���� IS NULL))
				) D
			    WHERE A.ԭ�ϲ���ID = B.ҩƷID AND B.ҩƷID = D.�շ�ϸĿID AND B.ҩƷID = C.ҩƷID (+)
				AND A.���Ʋ���ID = V_ҩƷ�շ���¼.ҩƷID
			    UNION all
			    SELECT DECODE(SIGN(NVL(C.ʵ�ʽ��,0)),1,(C.�ּ�-C.�ּ�*(C.ʵ�ʲ��/C.ʵ�ʽ��)),(C.�ּ�-C.�ּ�*(B.ָ�������/100)))*(A.����/A.��ĸ) AS �ɱ���
			    FROM ���Ʋ��Ϲ��� A,
				(SELECT j.����id AS ҩƷid,j.ָ������� FROM �������� j,�շ���ĿĿ¼ q  WHERE j.����id=q.id and  nvl(q.�Ƿ���,0)=1) B,
				(SELECT �ⷿID,ҩƷID,ʵ�ʽ��,ʵ�ʲ��,ʵ�ʽ��/ʵ������ AS �ּ� FROM ҩƷ��� WHERE ���� = 1 AND �ⷿID = V_ҩƷ�շ���¼.�Է�����ID AND ʵ������>0 ) C 
			    WHERE A.ԭ�ϲ���ID = B.ҩƷID AND B.ҩƷID = C.ҩƷID AND A.���Ʋ���ID = V_ҩƷ�շ���¼.ҩƷID);
		    EXCEPTION
			WHEN OTHERS THEN
			    V_�ɱ��� := 0;
		    END;

		    V_�ɱ���� := V_�ɱ��� * V_ҩƷ�շ���¼.ʵ������;
		    V_������ := V_ҩƷ�շ���¼.���۽�� - V_�ɱ����;

			--���¸ò��ϵĳɱ���
			UPDATE ��������
			SET �ɱ���=V_�ɱ��� 
			WHERE ����ID=V_ҩƷ�շ���¼.ҩƷID;
		    
		END IF;

		UPDATE ҩƷ�շ���¼
		SET �ɱ��� = V_�ɱ���,�ɱ���� = V_�ɱ����,��� = V_������
		WHERE ID = V_ҩƷ�շ���¼.ID;

		UPDATE ҩƷ���
		SET ��������=NVL(��������,0)+DECODE(V_ҩƷ�շ���¼.���ϵ��,1,NVL(V_ҩƷ�շ���¼.ʵ������,0),0),
		    ʵ������=NVL(ʵ������,0)+NVL(V_ҩƷ�շ���¼.ʵ������,0)*V_ҩƷ�շ���¼.���ϵ��,
		    ʵ�ʽ��=NVL(ʵ�ʽ��,0)+NVL(V_ҩƷ�շ���¼.���۽��,0)*V_ҩƷ�շ���¼.���ϵ��,
		    ʵ�ʲ��=NVL(ʵ�ʲ��,0)+V_������*V_ҩƷ�շ���¼.���ϵ��
		WHERE �ⷿID = V_ҩƷ�շ���¼.�ⷿID
		    AND ҩƷID = V_ҩƷ�շ���¼.ҩƷID
		    AND NVL (����,0) = 0
		    AND ���� = 1;

		IF SQL%NOTFOUND THEN
		Insert INTO ҩƷ���
		    (�ⷿID,ҩƷID,����,��������,ʵ������,ʵ�ʽ��,ʵ�ʲ��)
		VALUES (V_ҩƷ�շ���¼.�ⷿID,V_ҩƷ�շ���¼.ҩƷID,1,
		    DECODE (V_ҩƷ�շ���¼.���ϵ��,1,NVL (V_ҩƷ�շ���¼.ʵ������,0),0),
		    V_ҩƷ�շ���¼.ʵ������ * V_ҩƷ�շ���¼.���ϵ��,
		    V_ҩƷ�շ���¼.���۽�� * V_ҩƷ�շ���¼.���ϵ��,
		    V_������ * V_ҩƷ�շ���¼.���ϵ��);
		END IF;

		DELETE
		FROM ҩƷ���
		WHERE �ⷿID = V_ҩƷ�շ���¼.�ⷿID
		    AND ҩƷID = V_ҩƷ�շ���¼.ҩƷID
		    AND NVL (��������,0) = 0
		    AND NVL (ʵ������,0) = 0
		    AND NVL (ʵ�ʽ��,0) = 0
		    AND NVL (ʵ�ʲ��,0) = 0;

		--����ҩƷ�շ����ܱ����Ӧ����
		UPDATE ҩƷ�շ�����
		SET ���� =NVL(����,0)+NVL(V_ҩƷ�շ���¼.ʵ������,0)*V_ҩƷ�շ���¼.���ϵ��,
		    ��� =NVL(���,0)+NVL (V_ҩƷ�շ���¼.���۽��,0) * V_ҩƷ�շ���¼.���ϵ��,
		    ��� =NVL (���,0) + V_������ * V_ҩƷ�շ���¼.���ϵ��
		WHERE ���� = TRUNC (SYSDATE)
		    AND �ⷿID = V_ҩƷ�շ���¼.�ⷿID
		    AND ҩƷID = V_ҩƷ�շ���¼.ҩƷID
		    AND ���ID = V_ҩƷ�շ���¼.������ID
		    AND ���� = 16;

		IF SQL%NOTFOUND THEN
		    Insert INTO ҩƷ�շ�����
			(����,�ⷿID,ҩƷID,���ID,����,����,���,���)
		    VALUES (
			TRUNC (SYSDATE),
			V_ҩƷ�շ���¼.�ⷿID,
			V_ҩƷ�շ���¼.ҩƷID,
			V_ҩƷ�շ���¼.������ID,
			16,
			V_ҩƷ�շ���¼.ʵ������ * V_ҩƷ�շ���¼.���ϵ��,
			V_ҩƷ�շ���¼.���۽�� * V_ҩƷ�շ���¼.���ϵ��,
			V_������ * V_ҩƷ�շ���¼.���ϵ��
			);
		END IF;
    END LOOP;
EXCEPTION
    WHEN mErrItem THEN
        Raise_application_error (-20101,mErrMsg);
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE,SQLERRM);
END zl_���Ʋ������_verify;
/





-----------------------------------------------------------
--�������õ���˴���
--˵�������ȶ�ҩƷ�շ���¼���е�����˺����ʱ�估ʵ���������д���
--���Ŷ�ҩƷ���Ͳ����շ����ܱ��е���Ӧ�����ͽ����д���
--�ر�˵������ҩƷ�����շ����ܱ�Ĵ����Ƿֿ��ģ�
--
--�������
--���룺NO_IN,�����_IN
------------------------------------------------------------

CREATE OR REPLACE PROCEDURE zl_��������_verify (
	���_IN		IN ҩƷ�շ���¼.���%TYPE,
	NO_IN		IN ҩƷ�շ���¼.NO%TYPE,
	�ⷿID_IN	IN ҩƷ�շ���¼.�ⷿID%TYPE,
	�Է�����ID_IN   IN ҩƷ�շ���¼.�Է�����ID%TYPE,
	����ID_IN	IN ҩƷ�շ���¼.ҩƷID%TYPE,
	����_IN		IN ҩƷ�շ���¼.����%TYPE,
	����_IN		IN ҩƷ�շ���¼.����%TYPE,
	��д����_IN	IN ҩƷ�շ���¼.��д����%TYPE,
	ʵ������_IN	IN ҩƷ�շ���¼.ʵ������%TYPE,
	�ɱ���_IN	IN ҩƷ�շ���¼.�ɱ���%TYPE,
	�ɱ����_IN	IN ҩƷ�շ���¼.�ɱ����%TYPE,
	���۽��_IN	IN ҩƷ�շ���¼.���۽��%TYPE,
	���_IN		IN ҩƷ�շ���¼.���%TYPE,
	������ID_IN   IN ҩƷ�շ���¼.������ID%TYPE,
	�����_IN	IN ҩƷ�շ���¼.�����%TYPE,
	�������_IN	IN ҩƷ�շ���¼.�������%TYPE,
	����_IN		IN ҩƷ�շ���¼.����%TYPE := NULL,
	Ч��_IN		IN ҩƷ�շ���¼.Ч��%TYPE := NULL
)
IS
	mErrMsg		varchar2(100);
	mErrItem		EXCEPTION;
	V_��������		ҩƷ���.��������%TYPE;
	V_����		�շ���ĿĿ¼.����%TYPE;
	V_ʵ�ʿ����      ҩƷ���.ʵ�ʽ��%TYPE;
	V_ʵ�ʿ����      ҩƷ���.ʵ�ʲ��%TYPE;
	V_�����            number(18,8);
	V_������		ҩƷ���.ʵ�ʲ��%TYPE;
	V_�ɱ���            ҩƷ�շ���¼.�ɱ���%TYPE;
	V_�ɱ����		ҩƷ�շ���¼.�ɱ����%TYPE;
	V_С��		number(2);
	v_���ɱ�����		ϵͳ������.����ֵ%type;
BEGIN
    	SELECT ����ֵ INTO v_���ɱ����� FROM ϵͳ������ WHERE ������=120;

	BEGIN 
	    SELECT  to_number(Nvl(����ֵ,ȱʡֵ),'99999')  INTO v_С�� 
	    From ϵͳ������ where ������='���ý���λ��';
	EXCEPTION 
		WHEN OTHERS THEN 
			v_С��:=2;		
	END;

	--�������ô������������ʱ�ı�ʵ��������
	--�������ȶ�ʵ��������������Ӧ���ֶν��и��¡�

	BEGIN
		SELECT nvl(ʵ�ʽ��,0),nvl(ʵ�ʲ��,0),nvl(��������,0)    INTO V_ʵ�ʿ����,V_ʵ�ʿ����,V_��������
		FROM ҩƷ���
		WHERE ҩƷID = ����ID_IN     AND NVL(����,0) = ����_IN AND �ⷿID = �ⷿID_IN AND ���� = 1 AND ROWNUM = 1;

	EXCEPTION
		WHEN OTHERS THEN
		    V_ʵ�ʿ���� := 0;
		    V_�������� := 0;
	END;

	IF V_ʵ�ʿ���� <= 0 THEN
		BEGIN
		    SELECT ָ������� / 100,nvl(�ɱ���,0)    INTO V_�����,v_�ɱ���
		    FROM ��������
		    WHERE ����ID = ����ID_IN;
		EXCEPTION
		    WHEN OTHERS THEN
			V_����� := 0;
		END;
		IF v_���ɱ����� ='1' THEN 
			IF v_�ɱ���=0 THEN 
				V_����� := V_ʵ�ʿ���� / V_ʵ�ʿ����;
				V_������ :=round( ���۽��_IN * V_�����,v_С��);
			ELSE 
				V_������ :=round( ���۽��_IN-ʵ������_IN * V_�ɱ���,v_С��);
			END IF ;
		ELSE 
			V_������ :=round( ���۽��_IN * V_�����,v_С��);
		END IF ;
	ELSE
		V_����� := V_ʵ�ʿ���� / V_ʵ�ʿ����;
		V_������ :=round( ���۽��_IN * V_�����,v_С��);
	END IF;

	IF ʵ������_IN=0 THEN 
		V_�ɱ��� :=�ɱ���_IN; 
	ELSE 
		V_�ɱ��� := (���۽��_IN - V_������) / ʵ������_IN; 
	END IF; 

	V_�ɱ���� := round(V_�ɱ��� * ʵ������_IN,v_С��);

	UPDATE ҩƷ�շ���¼
	SET ����� = NVL (�����_IN,�����),
		������� = �������_IN,
		ʵ������ = ʵ������_IN,
		�ɱ��� = V_�ɱ���,
		�ɱ���� = V_�ɱ����,
		���۽�� = ���۽��_IN,
		��� = V_������
	WHERE NO = NO_IN
		AND ���� =20
		AND ҩƷID = ����ID_IN
		AND ��� = ���_IN
		AND ��¼״̬ = 1
		AND ����� IS NULL;

	--����ҩƷ������Ӧ����
	IF SQL%ROWCOUNT = 0 THEN
		mErrMsg:='[ZLSOFT]�õ����Ѿ���������ˣ�[ZLSOFT]';
		RAISE mErrItem;
	END IF;

	IF ����_IN > 0 AND (V_�������� + ��д����_IN - ʵ������_IN) < 0 THEN
		SELECT ���� INTO V_���� FROM �շ���ĿĿ¼  WHERE ID = ����ID_IN;
		mErrMsg:='[ZLSOFT]����Ϊ' || V_���� || ',����Ϊ' || ����_IN ||'�ķ����������' || CHR (10) ||CHR (13) ||'���ÿ������������[ZLSOFT]';
		RAISE mErrItem;
	END IF;

	UPDATE ҩƷ���
	SET �������� = NVL (��������,0) + ��д����_IN - ʵ������_IN,
		ʵ������ = NVL (ʵ������,0) - ʵ������_IN,
		ʵ�ʽ�� = NVL (ʵ�ʽ��,0) - ���۽��_IN,
		ʵ�ʲ�� = NVL (ʵ�ʲ��,0) - V_������
	WHERE �ⷿID = �ⷿID_IN
		AND ҩƷID = ����ID_IN
		AND NVL (����,0) = NVL (����_IN,0)
		AND ���� = 1;

	IF SQL%NOTFOUND THEN
		Insert INTO ҩƷ���(�ⷿID,ҩƷID,����,����,��������,ʵ������,ʵ�ʽ��,ʵ�ʲ��)
		VALUES (�ⷿID_IN,����ID_IN,����_IN,1,-ʵ������_IN,-ʵ������_IN,-���۽��_IN,-V_������);
	END IF;

	DELETE
	FROM ҩƷ���
	WHERE �ⷿID = �ⷿID_IN
		AND ҩƷID = ����ID_IN
		AND NVL (��������,0) = 0
		AND NVL (ʵ������,0) = 0
		AND NVL (ʵ�ʽ��,0) = 0
		AND NVL (ʵ�ʲ��,0) = 0;

	--�������շ����ܱ����Ӧ����
	UPDATE ҩƷ�շ�����
	SET ���� = NVL (����,0) - ʵ������_IN,
		��� = NVL (���,0) - ���۽��_IN,
		��� = NVL (���,0) - V_������
	WHERE ���� = TRUNC (SYSDATE)
		AND �ⷿID = �ⷿID_IN
		AND ҩƷID = ����ID_IN
		AND ���ID = ������ID_IN
		AND ���� = 20;

	IF SQL%NOTFOUND THEN
		Insert INTO ҩƷ�շ�����
		    (����,�ⷿID,ҩƷID,���ID,����,����,���,���)
		VALUES (TRUNC (SYSDATE),�ⷿID_IN,����ID_IN,������ID_IN,20,-ʵ������_IN,-���۽��_IN,-V_������    );
	END IF;
EXCEPTION
    WHEN mErrItem THEN
        Raise_application_error (-20101,mErrMsg);
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE,SQLERRM);
END zl_��������_verify;
/




-----------------------------------------------------------
--���������������˴���
--˵�������ȶ�ҩƷ�շ���¼���е�����˺����ʱ�估ʵ���������д���
--���Ŷ�ҩƷ����ҩƷ�շ����ܱ��е���Ӧ�����ͽ����д���
--�ر�˵������ҩƷ�����շ����ܱ�Ĵ����Ƿֿ��ģ�
--
--�������
--���룺NO_IN,�����_IN
------------------------------------------------------------

CREATE OR REPLACE PROCEDURE zl_������������_verify (
	���_IN		IN ҩƷ�շ���¼.���%TYPE,
	NO_IN		IN ҩƷ�շ���¼.NO%TYPE,
	�ⷿID_IN	IN ҩƷ�շ���¼.�ⷿID%TYPE,
	����ID_IN	IN ҩƷ�շ���¼.ҩƷID%TYPE,
	����_IN         IN ҩƷ�շ���¼.����%TYPE,
	ʵ������_IN     IN ҩƷ�շ���¼.ʵ������%TYPE,
	�ɱ���_IN	IN ҩƷ�շ���¼.�ɱ���%TYPE,
	�ɱ����_IN     IN ҩƷ�շ���¼.�ɱ����%TYPE,
	���۽��_IN     IN ҩƷ�շ���¼.���۽��%TYPE,
	���_IN         IN ҩƷ�շ���¼.���%TYPE,
	������ID_IN   IN ҩƷ�շ���¼.������ID%TYPE,
	�����_IN	IN ҩƷ�շ���¼.�����%TYPE,
	�������_IN	IN ҩƷ�շ���¼.�������%TYPE
)
IS
	mErrMsg            varchar2(100);
	mErrItem        EXCEPTION;

	V_��������        ҩƷ���.��������%TYPE;
	V_����            �շ���ĿĿ¼.����%TYPE;
	V_ʵ�ʿ����        ҩƷ���.ʵ�ʽ��%TYPE;
	V_ʵ�ʿ����        ҩƷ���.ʵ�ʲ��%TYPE;
	V_�����            number(18,8);
	V_������        ҩƷ���.ʵ�ʲ��%TYPE;
	V_�ɱ���            ҩƷ�շ���¼.�ɱ���%TYPE;
	V_�ɱ����        ҩƷ�շ���¼.�ɱ����%TYPE;
	V_С��		number(2);
	v_���ɱ�����		ϵͳ������.����ֵ%type;

BEGIN
	BEGIN 
		SELECT  to_number(Nvl(����ֵ,ȱʡֵ),'99999')  INTO v_С�� 
		From ϵͳ������ where ������='���ý���λ��';
	EXCEPTION 
		WHEN OTHERS THEN 
			v_С��:=2;		
	END;
    	SELECT ����ֵ INTO v_���ɱ����� FROM ϵͳ������ WHERE ������=120;



	--�������ô������������ʱ�ı�ʵ��������
	--�������ȶ�ʵ��������������Ӧ���ֶν��и��¡�
	BEGIN
		SELECT nvl(ʵ�ʽ��,0),nvl(ʵ�ʲ��,0)    INTO V_ʵ�ʿ����,V_ʵ�ʿ����
		FROM ҩƷ���
		WHERE ҩƷID = ����ID_IN
		    AND NVL (����,0) = ����_IN
		    AND �ⷿID = �ⷿID_IN
		    AND ���� = 1
		    AND ROWNUM = 1;
	EXCEPTION
		WHEN OTHERS THEN
		    V_ʵ�ʿ���� := 0;
	END;
	IF V_ʵ�ʿ���� <= 0 THEN
		BEGIN
		    SELECT ָ������� / 100,nvl(�ɱ���,0)    INTO V_�����,v_�ɱ���
		    FROM ��������
		    WHERE ����ID = ����ID_IN;
		EXCEPTION
		    WHEN OTHERS THEN
			V_����� := 0;
		END;
		IF v_���ɱ����� ='1' THEN 
			IF v_�ɱ���=0 THEN 
				V_����� := V_ʵ�ʿ���� / V_ʵ�ʿ����;
				V_������ :=round( ���۽��_IN * V_�����,v_С��);
			ELSE 
				V_������ :=round( ���۽��_IN-ʵ������_IN * V_�ɱ���,v_С��);
			END IF ;
		ELSE 
			V_������ :=round( ���۽��_IN * V_�����,v_С��);
		END IF ;
	ELSE
		V_����� := V_ʵ�ʿ���� / V_ʵ�ʿ����;
		V_������ :=round( ���۽��_IN * V_�����,v_С��);
	END IF;
	IF ʵ������_IN<=0 THEN 
		V_�ɱ��� := �ɱ���_IN;
	ELSE 
		V_�ɱ��� := (���۽��_IN - V_������) / ʵ������_IN;
	END IF ;
	V_�ɱ���� :=round( V_�ɱ��� * ʵ������_IN,v_С��);

	UPDATE ҩƷ�շ���¼
	SET ����� = NVL (�����_IN,�����),
		������� = �������_IN,
		�ɱ��� = V_�ɱ���,
		�ɱ���� = V_�ɱ����,
		��� = V_������
	WHERE NO = NO_IN
		AND ���� = 21
		AND ҩƷID = ����ID_IN
		AND ��� = ���_IN
		AND ��¼״̬ = 1
		AND ����� IS NULL;

	IF SQL%ROWCOUNT = 0 THEN
		mErrMsg:='[ZLSOFT]�õ����Ѿ���������ˣ�[ZLSOFT]';
		RAISE mErrItem;
	END IF;

	--����ҩƷ������Ӧ����
	UPDATE ҩƷ���
	SET ʵ������ = NVL (ʵ������,0) - ʵ������_IN,
		ʵ�ʽ�� = NVL (ʵ�ʽ��,0) - ���۽��_IN,
		ʵ�ʲ�� = NVL (ʵ�ʲ��,0) - V_������
	WHERE �ⷿID = �ⷿID_IN
		AND ҩƷID = ����ID_IN
		AND NVL (����,0) = NVL (����_IN,0)
		AND ���� = 1;

	IF SQL%NOTFOUND THEN
		Insert INTO ҩƷ���
		    (�ⷿID,ҩƷID,����,����,��������,ʵ������,ʵ�ʽ��,ʵ�ʲ��)
		VALUES (�ⷿID_IN,����ID_IN,����_IN,1,-ʵ������_IN,-ʵ������_IN,-���۽��_IN,-V_������);
	END IF;

	DELETE
	FROM ҩƷ���
	WHERE �ⷿID = �ⷿID_IN
		AND ҩƷID = ����ID_IN
		AND NVL (��������,0) = 0
		AND NVL (ʵ������,0) = 0
		AND NVL (ʵ�ʽ��,0) = 0
		AND NVL (ʵ�ʲ��,0) = 0;

	--��ҩƷ�շ����ܱ����Ӧ����
	UPDATE ҩƷ�շ�����
	SET ���� = NVL (����,0) - ʵ������_IN,
		��� = NVL (���,0) - ���۽��_IN,
		��� = NVL (���,0) - V_������
	WHERE ���� = TRUNC (SYSDATE)
		AND �ⷿID = �ⷿID_IN
		AND ҩƷID = ����ID_IN
		AND ���ID = ������ID_IN
		AND ���� = 21;

	IF SQL%NOTFOUND THEN
		Insert INTO ҩƷ�շ�����
		    (����,�ⷿID,ҩƷid ,���ID,����,����,���,���)
		VALUES (TRUNC (SYSDATE),�ⷿID_IN,����ID_IN,������ID_IN,21,-ʵ������_IN,-���۽��_IN,-V_������);
	END IF;
EXCEPTION
    WHEN mErrItem THEN
        Raise_application_error (-20101,mErrMsg);
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE,SQLERRM);
END zl_������������_verify;
/




-----------------------------------------------------------
--�����ƿ����˴���
--˵�������ȶ�ҩƷ�շ���¼���е�����˺����ʱ�估ʵ���������д���
--���Ŷ�ҩƷ����ҩƷ�շ����ܱ��е���Ӧ�����ͽ����д���
--�ر�˵������ҩƷ�����շ����ܱ�Ĵ����Ƿֿ��ģ�
------------------------------------------------------------

    
CREATE OR REPLACE PROCEDURE ZL_�����ƿ�_VERIFY (
    ���_IN		IN ҩƷ�շ���¼.���%TYPE,
    �ⷿID_IN		IN ҩƷ�շ���¼.�ⷿID%TYPE,
    �Է�����ID_IN	IN ҩƷ�շ���¼.�Է�����ID%TYPE,
    ����ID_IN		IN ҩƷ�շ���¼.ҩƷID%TYPE,
    ����_IN		IN ҩƷ�շ���¼.����%TYPE,
    ������_IN		IN ҩƷ�շ���¼.����%TYPE,
    ��д����_IN		IN ҩƷ�շ���¼.��д����%TYPE,
    ʵ������_IN		IN ҩƷ�շ���¼.ʵ������%TYPE,
    �ɱ���_IN		IN ҩƷ�շ���¼.�ɱ���%TYPE,
    �ɱ����_IN		IN ҩƷ�շ���¼.�ɱ����%TYPE,
    ���۽��_IN		IN ҩƷ�շ���¼.���۽��%TYPE,
    ���_IN		IN ҩƷ�շ���¼.���%TYPE,
    �����ID_IN		IN ҩƷ�շ���¼.������ID%TYPE,
    �����ID_IN		IN ҩƷ�շ���¼.������ID%TYPE,
    NO_IN		IN ҩƷ�շ���¼.NO%TYPE,
    �����_IN		IN ҩƷ�շ���¼.�����%TYPE,
    ����_IN		IN ҩƷ�շ���¼.����%TYPE := NULL,
    Ч��_IN		IN ҩƷ�շ���¼.Ч��%TYPE := NULL,
    ���Ч��_IN        IN ҩƷ�շ���¼.���Ч��%type:=NULL,
    �������_IN		IN ҩƷ�շ���¼.�������%TYPE := NULL,
    �ƿⵥ_IN		IN NUMBER:=1)
IS
	mErrMsg		varchar2(500);
	mErrItem	EXCEPTION;

	V_������	ҩƷ�շ���¼.����%TYPE := NULL;
	V_ʵ�ʿ����	ҩƷ���.ʵ�ʽ��%TYPE;
	V_ʵ�ʿ����	ҩƷ���.ʵ�ʲ��%TYPE;
	V_�����	NUMBER(18,8);
	V_������	ҩƷ���.ʵ�ʲ��%TYPE;
	V_�ɱ���	ҩƷ�շ���¼.�ɱ���%TYPE;
	V_�ɱ����	ҩƷ�շ���¼.�ɱ����%TYPE;
	V_ʵ������	ҩƷ���.ʵ������%TYPE;
	V_����		�շ���ĿĿ¼.����%TYPE;
	v_С��		NUMBER ;
	v_���ɱ�����		ϵͳ������.����ֵ%type;

BEGIN
	--��ȡ���С��λ��
	SELECT to_number(Nvl(����ֵ,ȱʡֵ),'99999') INTO v_С�� FROM ϵͳ������ WHERE ������='���ý���λ��';
    	SELECT ����ֵ INTO v_���ɱ����� FROM ϵͳ������ WHERE ������=120;

	--�����ƿ⴦�����������ʱ�ı�ʵ��������
        --�������ȶ�ʵ��������������Ӧ���ֶν��и��¡�
	BEGIN
		SELECT NVL(ʵ�ʽ��,0), NVL(ʵ�ʲ��,0), NVL(ʵ������,0) INTO V_ʵ�ʿ����, V_ʵ�ʿ����, V_ʵ������
		FROM ҩƷ���
		WHERE ҩƷID = ����ID_IN
			AND NVL (����, 0) = ������_IN
			AND �ⷿID = �ⷿID_IN
			AND ���� = 1
			AND ROWNUM = 1;
	EXCEPTION
		WHEN OTHERS THEN
		    V_ʵ�ʿ���� := 0;
		    V_ʵ������ := 0;
	END;

	IF V_ʵ�ʿ���� <= 0 THEN
		BEGIN
		    SELECT ָ������� / 100,nvl(�ɱ���,0)    INTO V_�����,v_�ɱ���
		    FROM ��������
		    WHERE ����ID = ����ID_IN;
		EXCEPTION
		    WHEN OTHERS THEN
			V_����� := 0;
		END;
		IF v_���ɱ����� ='1' THEN 
			IF v_�ɱ���=0 THEN 
				V_����� := V_ʵ�ʿ���� / V_ʵ�ʿ����;
				V_������ :=round( ���۽��_IN * V_�����,v_С��);
			ELSE 
				V_������ :=round( ���۽��_IN-ʵ������_IN * V_�ɱ���,v_С��);
			END IF ;
		ELSE 
			V_������ :=round( ���۽��_IN * V_�����,v_С��);
		END IF ;
	ELSE
		V_����� := V_ʵ�ʿ���� / V_ʵ�ʿ����;
		V_������ :=round( ���۽��_IN * V_�����,v_С��);
	END IF;


	IF ʵ������_IN=0 THEN
		V_�ɱ��� :=�ɱ���_IN;
	ELSE
		V_�ɱ��� := (���۽��_IN - V_������) / ʵ������_IN; 
	END IF; 
	
	V_�ɱ���� := ROUND(V_�ɱ��� * ʵ������_IN,v_С��);

	UPDATE ҩƷ�շ���¼
	SET ����� = NVL (�����_IN, �����),
	     ������� = �������_IN,
	     ʵ������ = ʵ������_IN,
	     �ɱ��� = V_�ɱ���,
	     �ɱ���� = V_�ɱ����,
	     ���۽�� = ���۽��_IN,
	     ��� = V_������
	WHERE NO = NO_IN
	AND ���� = 19
	AND ҩƷID = ����ID_IN
	AND ��¼״̬ = 1
	AND ��� IN (���_IN, ���_IN + 1)
	AND ����� IS NULL;

	IF SQL%ROWCOUNT = 0 THEN
		mErrMsg:='[ZLSOFT]�õ����Ѿ���������ˣ�[ZLSOFT]';
		RAISE mErrItem;
	END IF;

	IF ������_IN > 0 THEN
		IF V_ʵ������ < ʵ������_IN THEN
			SELECT ���� INTO V_���� FROM �շ���ĿĿ¼ WHERE ID = ����ID_IN;
			mErrMsg:= '[ZLSOFT]����Ϊ' || V_���� || ',����Ϊ' || ����_IN ||'�Ŀⷿ��������' ||
				CHR(10) || CHR(13) || '���ÿ������������[ZLSOFT]';
			RAISE mErrItem;
		END IF;
	END IF;

	--ȡ����������
	SELECT ���� INTO V_������ FROM ҩƷ�շ���¼ WHERE NO = NO_IN AND ���� = 19 AND ��¼״̬ = 1 AND ��� = ���_IN+1;
        
	--���������Ĳ��Ͽ�����Ӧ����

	UPDATE ҩƷ���
	SET �������� = NVL (��������, 0) + ʵ������_IN,
		ʵ������ = NVL (ʵ������, 0) + ʵ������_IN,
		ʵ�ʽ�� = NVL (ʵ�ʽ��, 0) + ���۽��_IN,
		ʵ�ʲ�� = NVL (ʵ�ʲ��, 0) + V_������,
		�ϴβɹ��� = V_�ɱ���,
		�ϴ����� = NVL (����_IN, �ϴ�����),
		�ϴβ��� = NVL (����_IN, �ϴβ���),
		���Ч��=nvl(���Ч��_IN,���Ч��)
	WHERE �ⷿID = �Է�����ID_IN AND ҩƷID = ����ID_IN AND NVL (����, 0) = NVL (V_������, 0) AND ���� = 1;

	IF SQL%NOTFOUND THEN
		INSERT INTO ҩƷ���
			(�ⷿID,ҩƷID,����,����,��������,ʵ������,ʵ�ʽ��,ʵ�ʲ��,�ϴβɹ���,�ϴ�����,�ϴβ���,Ч��,���Ч��)
		VALUES 
			(�Է�����ID_IN,����ID_IN,V_������,1,ʵ������_IN,ʵ������_IN,���۽��_IN,V_������,V_�ɱ���,����_IN,����_IN,Ч��_IN,���Ч��_IN);
	END IF;

	--���ĳ����Ĳ��Ͽ�����Ӧ����

	UPDATE ҩƷ���
	SET	ʵ������ = NVL (ʵ������, 0) - ʵ������_IN,
		ʵ�ʽ�� = NVL (ʵ�ʽ��, 0) - ���۽��_IN,
		ʵ�ʲ�� = NVL (ʵ�ʲ��, 0) - V_������
	WHERE �ⷿID = �ⷿID_IN AND ҩƷID = ����ID_IN
		AND NVL (����, 0) = NVL (������_IN, 0)
		AND ���� = 1;

	IF SQL%NOTFOUND THEN
		INSERT INTO ҩƷ���(�ⷿID,ҩƷID,����,����,��������,ʵ������,ʵ�ʽ��,ʵ�ʲ��,�ϴ�����,Ч��,���Ч��)
		VALUES (�ⷿID_IN,����ID_IN,������_IN,1,0,-ʵ������_IN,-���۽��_IN,-V_������,����_IN,Ч��_IN,���Ч��_IN);
	END IF;

	DELETE
	FROM ҩƷ���
	WHERE �ⷿID = �ⷿID_IN
		AND ҩƷID = ����ID_IN
		AND NVL (��������, 0) = 0
		AND NVL (ʵ������, 0) = 0
		AND NVL (ʵ�ʽ��, 0) = 0
		AND NVL (ʵ�ʲ��, 0) = 0;

	--����������ҩƷ�շ����ܱ����Ӧ����
	UPDATE ҩƷ�շ�����
	SET	���� = NVL (����, 0) + ʵ������_IN,
		��� = NVL (���, 0) + ���۽��_IN,
		��� = NVL (���, 0) + V_������
	WHERE ���� = TRUNC (SYSDATE)
		AND �ⷿID = �Է�����ID_IN
		AND ҩƷID = ����ID_IN
		AND ���ID = �����ID_IN
		AND ���� = 19;

	IF SQL%NOTFOUND THEN
		INSERT INTO ҩƷ�շ�����
			(����, �ⷿID, ҩƷID, ���ID, ����, ����, ���, ���)
		VALUES(
			TRUNC(SYSDATE),�Է�����ID_IN,����ID_IN,�����ID_IN,19,ʵ������_IN,���۽��_IN,V_������);
	END IF;

	--���ĳ�����ҩƷ�շ����ܱ����Ӧ����
	UPDATE ҩƷ�շ�����
	SET ���� = NVL (����, 0) - ʵ������_IN,
		��� = NVL (���, 0) - ���۽��_IN,
		��� = NVL (���, 0) - V_������
	WHERE ���� = TRUNC (SYSDATE)
		AND �ⷿID = �ⷿID_IN
		AND ҩƷID = ����ID_IN
		AND ���ID = �����ID_IN
		AND ���� = 19;

	IF SQL%NOTFOUND THEN
		INSERT INTO ҩƷ�շ�����
				(����, �ⷿID, ҩƷID, ���ID, ����, ����, ���, ���)
			VALUES (
				TRUNC(SYSDATE),�ⷿID_IN,����ID_IN,�����ID_IN,19,-ʵ������_IN,-���۽��_IN,-V_������);
	END IF;
EXCEPTION
    WHEN mErrItem THEN
        RAISE_APPLICATION_ERROR (-20101, mErrMsg);
    WHEN OTHERS THEN
        ZL_ERRORCENTER (SQLCODE, SQLERRM);
END ZL_�����ƿ�_VERIFY;
/

CREATE OR REPLACE PROCEDURE zl_��ҩƷ��_INSERT(
    ���_IN IN ������ĿĿ¼.���%TYPE := NULL,
    ����ID_IN IN ������ĿĿ¼.����ID%TYPE := NULL,
    ID_IN IN ������ĿĿ¼.ID%TYPE,
    ����_IN IN ������ĿĿ¼.����%TYPE := NULL,
    ����_IN IN ������ĿĿ¼.����%TYPE := NULL,
    ƴ��_IN IN ������Ŀ����.����%TYPE := NULL,
    ���_IN IN ������Ŀ����.����%TYPE := NULL,
    Ӣ��_IN IN ������Ŀ����.����%TYPE := NULL,
    ��λ_IN IN ������ĿĿ¼.���㵥λ%TYPE := NULL,
    ҩƷ����_IN IN ҩƷ����.ҩƷ����%TYPE := NULL,
    �������_IN IN ҩƷ����.�������%TYPE := NULL,
    ��ֵ����_IN IN ҩƷ����.��ֵ����%TYPE := NULL,
    ��Դ���_IN IN ҩƷ����.��Դ���%TYPE := NULL,
    ��ҩ�ݴ�_IN IN ҩƷ����.��ҩ�ݴ�%TYPE := NULL,
    ҩƷ����_IN IN ҩƷ����.ҩƷ����%TYPE := NULL,
    ����ְ��_IN IN ҩƷ����.����ְ��%TYPE := '00',
    ��������_IN IN ҩƷ����.��������%TYPE := NULL,
    ����ҩ��_IN IN ҩƷ����.����ҩ��%TYPE := 0,
    �Ƿ���ҩ_IN IN ҩƷ����.�Ƿ���ҩ%TYPE := 0,
    �Ƿ�ԭ��_IN IN ҩƷ����.�Ƿ�ԭ��%TYPE := 0,
    �Ƿ�Ƥ��_IN IN ҩƷ����.�Ƿ�Ƥ��%TYPE := 0,
    �ο�Ŀ¼Id_IN In ������ĿĿ¼.�ο�Ŀ¼Id%Type:=Null,
    Ʒ��ҽ��_IN In ҩƷ����.Ʒ��ҽ��%TYPE := 0,
    ��������_IN IN Varchar2      --��"|"�ָ��ı�����¼��ÿ����¼��"����^ƴ��^���"��֯

) IS
    v_Records VARCHAR2(4000);   --��ʱ��¼�������ݵ��ַ���
    v_CurrRec VARCHAR2(1000);   --�����ڱ�����¼�е�һ������
    v_Fields  VARCHAR2(1000);   --��ʱ��¼һ���������ַ���
    v_���� ������ĿĿ¼.����%TYPE;
    v_ƴ�� ������Ŀ����.����%TYPE;
    v_��� ������Ŀ����.����%TYPE;
	v_������ĿID number(18);
BEGIN
    INSERT INTO ������ĿĿ¼(���,����ID,ID,����,����,���㵥λ,
        ���㷽ʽ,ִ��Ƶ��,�����Ա�,����Ӧ��,�����Ŀ,ִ�а���,�Ƽ�����,�������,����ʱ��,����ʱ��,�ο�Ŀ¼Id)
    VALUES (���_IN,����ID_IN,ID_IN,����_IN,����_IN,��λ_IN,
        1,0,0,1,0,0,0,3,sysdate,to_date('3000-01-01','YYYY-MM-DD'),�ο�Ŀ¼Id_IN);

    INSERT INTO ҩƷ����(ҩ��ID,ҩƷ����,�������,��ֵ����,��Դ���,��ҩ�ݴ�,
        ҩƷ����,����ְ��,��������,����ҩ��,�Ƿ���ҩ,�Ƿ�ԭ��,�Ƿ�Ƥ��,Ʒ��ҽ��)
    VALUES (ID_IN,ҩƷ����_IN,�������_IN,��ֵ����_IN,��Դ���_IN,��ҩ�ݴ�_IN,
        ҩƷ����_IN,����ְ��_IN,��������_IN,����ҩ��_IN,�Ƿ���ҩ_IN,�Ƿ�ԭ��_IN,�Ƿ�Ƥ��_IN,Ʒ��ҽ��_IN);

    IF ƴ��_IN IS NOT NULL THEN
        INSERT INTO ������Ŀ����(������ĿID,����,����,����,����) VALUES (ID_IN,����_IN,1,ƴ��_IN,1);
    END IF;
    IF ���_IN IS NOT NULL THEN
        INSERT INTO ������Ŀ����(������ĿID,����,����,����,����) VALUES (ID_IN,����_IN,1,���_IN,2);
    END IF;
    IF Ӣ��_IN IS NOT NULL THEN
        INSERT INTO ������Ŀ����(������ĿID,����,����,����,����) VALUES (ID_IN,Ӣ��_IN,2,null,0);
    END IF;

    IF ��������_IN IS NULL THEN
        v_Records := null;
    ELSE
        v_Records := ��������_IN||'|';
    END IF;
    WHILE v_Records IS NOT NULL LOOP
        v_CurrRec:=substr(v_Records,1,instr(v_Records,'|')-1);
        v_Fields:=v_CurrRec;
        v_����:=substr(v_Fields,1,instr(v_Fields,'^')-1);
        v_Fields:=substr(v_Fields,instr(v_Fields,'^')+1);
        v_ƴ��:=substr(v_Fields,1,instr(v_Fields,'^')-1);
        v_Fields:=substr(v_Fields,instr(v_Fields,'^')+1);
        v_���:=v_Fields;
        IF V_ƴ�� IS NOT NULL THEN
            INSERT INTO ������Ŀ����(������ĿID,����,����,����,����) VALUES (ID_IN,v_����,9,v_ƴ��,1);
        END IF;
        IF v_��� IS NOT NULL THEN
            INSERT INTO ������Ŀ����(������ĿID,����,����,����,����) VALUES (ID_IN,v_����,9,v_���,2);
        END IF;
        v_Records:=replace('|'||v_Records,'|'||v_CurrRec||'|');
    END LOOP;
    --���ȱʡ�Ķ�Ӧ�������
    INSERT INTO ���Ƶ���Ӧ��(�����ļ�id,Ӧ�ó���,������Ŀid)
    SELECT A.�����ļ�id,1,ID_IN
    FROM ���Ƶ���Ӧ�� A,������ĿĿ¼ I
    Where A.������Ŀid=I.Id And I.���=���_IN And Ӧ�ó���=1 And Rownum<2;
    INSERT INTO ���Ƶ���Ӧ��(�����ļ�id,Ӧ�ó���,������Ŀid)
    SELECT A.�����ļ�id,2,ID_IN
    FROM ���Ƶ���Ӧ�� A,������ĿĿ¼ I
    Where A.������Ŀid=I.Id And I.���=���_IN And Ӧ�ó���=2 And Rownum<2;
	
EXCEPTION
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_��ҩƷ��_INSERT;
/

CREATE OR REPLACE PROCEDURE zl_��ҩƷ��_UPDATE(
    ����ID_IN IN ������ĿĿ¼.����ID%TYPE := NULL,
    ID_IN IN ������ĿĿ¼.ID%TYPE,
    ����_IN IN ������ĿĿ¼.����%TYPE := NULL,
    ����_IN IN ������ĿĿ¼.����%TYPE := NULL,
    ƴ��_IN IN ������Ŀ����.����%TYPE := NULL,
    ���_IN IN ������Ŀ����.����%TYPE := NULL,
    Ӣ��_IN IN ������Ŀ����.����%TYPE := NULL,
    ��λ_IN IN ������ĿĿ¼.���㵥λ%TYPE := NULL,
    ҩƷ����_IN IN ҩƷ����.ҩƷ����%TYPE := NULL,
    �������_IN IN ҩƷ����.�������%TYPE := NULL,
    ��ֵ����_IN IN ҩƷ����.��ֵ����%TYPE := NULL,
    ��Դ���_IN IN ҩƷ����.��Դ���%TYPE := NULL,
    ��ҩ�ݴ�_IN IN ҩƷ����.��ҩ�ݴ�%TYPE := NULL,
    ҩƷ����_IN IN ҩƷ����.ҩƷ����%TYPE := NULL,
    ����ְ��_IN IN ҩƷ����.����ְ��%TYPE := '00',
    ��������_IN IN ҩƷ����.��������%TYPE := NULL,
    ����ҩ��_IN IN ҩƷ����.����ҩ��%TYPE := 0,
    �Ƿ���ҩ_IN IN ҩƷ����.�Ƿ���ҩ%TYPE := 0,
    �Ƿ�ԭ��_IN IN ҩƷ����.�Ƿ�ԭ��%TYPE := 0,
    �Ƿ�Ƥ��_IN IN ҩƷ����.�Ƿ�Ƥ��%TYPE := 0,
    �ο�Ŀ¼Id_IN In ������ĿĿ¼.�ο�Ŀ¼Id%Type:=Null,
    Ʒ��ҽ��_IN In ҩƷ����.Ʒ��ҽ��%TYPE := 0,
    ��������_IN IN Varchar2      --��"|"�ָ��ı�����¼��ÿ����¼��"����^ƴ��^���"��֯
) IS
    v_Records VARCHAR2(4000);   --��ʱ��¼�������ݵ��ַ���
    v_CurrRec VARCHAR2(1000);   --�����ڱ�����¼�е�һ������
    v_Fields  VARCHAR2(1000);   --��ʱ��¼һ���������ַ���
    v_���� ������ĿĿ¼.����%TYPE;
    v_ƴ�� ������Ŀ����.����%TYPE;
    v_��� ������Ŀ����.����%TYPE;
    Err_NotFind  EXCEPTION;
BEGIN
    UPDATE ������ĿĿ¼
    SET ����ID=����ID_IN,����=����_IN,����=����_IN,���㵥λ=��λ_IN,�ο�Ŀ¼Id=�ο�Ŀ¼Id_IN
    WHERE ID=ID_IN;
    IF SQL%ROWCOUNT=0 THEN
        RAISE Err_NotFind;
    END IF;

    UPDATE ҩƷ����
    SET ҩƷ����=ҩƷ����_IN,�������=�������_IN,��ֵ����=��ֵ����_IN,��Դ���=��Դ���_IN,��ҩ�ݴ�=��ҩ�ݴ�_IN,
        ҩƷ����=ҩƷ����_IN,����ְ��=����ְ��_IN,��������=��������_IN,
        ����ҩ��=����ҩ��_IN,�Ƿ���ҩ=�Ƿ���ҩ_IN,�Ƿ�ԭ��=�Ƿ�ԭ��_IN,�Ƿ�Ƥ��=�Ƿ�Ƥ��_IN,Ʒ��ҽ��=Ʒ��ҽ��_IN
    WHERE ҩ��ID=ID_IN;

    update �շ���ĿĿ¼
    set ����=����_IN
    where ID in (select ҩƷid from ҩƷ��� where ҩ��ID=ID_IN);

    IF ƴ��_IN IS NULL THEN
        DELETE FROM ������Ŀ���� WHERE ������ĿID=ID_IN AND ����=1 AND ����=1;
        DELETE FROM �շ���Ŀ���� 
        WHERE �շ�ϸĿid in (select ҩƷid from ҩƷ��� where ҩ��id=ID_IN) AND ����=1 AND ����=1;
    ELSE
        UPDATE ������Ŀ���� SET ����=����_IN, ����=ƴ��_IN WHERE ������ĿID=ID_IN AND ����=1 AND ����=1;
        IF SQL%ROWCOUNT=0 THEN
            INSERT INTO ������Ŀ����(������ĿID,����,����,����,����) VALUES (ID_IN,����_IN,1,ƴ��_IN,1);
        END IF;
        for r_Spec in (
            select ҩƷid from ҩƷ��� where ҩ��id=ID_IN)
        loop
            update �շ���Ŀ���� SET ����=����_IN, ����=ƴ��_IN WHERE �շ�ϸĿID=r_Spec.ҩƷid AND ����=1 AND ����=1;
            if sql%rowcount=0 then
               insert into �շ���Ŀ����(�շ�ϸĿID,����,����,����,����) values(r_Spec.ҩƷid,����_IN,1,ƴ��_IN,1);
            end if;
        end loop;
    END IF;
    IF ���_IN IS NULL THEN
        DELETE FROM ������Ŀ���� WHERE ������ĿID=ID_IN AND ����=1 AND ����=2;
        DELETE FROM �շ���Ŀ���� 
        WHERE �շ�ϸĿid in (select ҩƷid from ҩƷ��� where ҩ��id=ID_IN) AND ����=1 AND ����=2;
    ELSE
        UPDATE ������Ŀ���� SET ����=����_IN, ����=���_IN WHERE ������ĿID=ID_IN AND ����=1 AND ����=2;
        IF SQL%ROWCOUNT=0 THEN
            INSERT INTO ������Ŀ����(������ĿID,����,����,����,����) VALUES (ID_IN,����_IN,1,���_IN,2);
        END IF;
        for r_Spec in (
            select ҩƷid from ҩƷ��� where ҩ��id=ID_IN)
        loop
            update �շ���Ŀ���� SET ����=����_IN, ����=���_IN WHERE �շ�ϸĿID=r_Spec.ҩƷid AND ����=1 AND ����=2;
            if sql%rowcount=0 then
               insert into �շ���Ŀ����(�շ�ϸĿID,����,����,����,����) values(r_Spec.ҩƷid,����_IN,1,���_IN,2);
            end if;
        end loop;
    END IF;
    IF Ӣ��_IN IS NULL THEN
        DELETE FROM ������Ŀ���� WHERE ������ĿID=ID_IN AND ����=2;
        DELETE FROM �շ���Ŀ���� 
        WHERE �շ�ϸĿid in (select ҩƷid from ҩƷ��� where ҩ��id=ID_IN) AND ����=2;
    ELSE
        UPDATE ������Ŀ���� SET ����=Ӣ��_IN WHERE ������ĿID=ID_IN AND ����=2;
        IF SQL%ROWCOUNT=0 THEN
            INSERT INTO ������Ŀ����(������ĿID,����,����,����,����) VALUES (ID_IN,Ӣ��_IN,2,null,0);
        END IF;
        for r_Spec in (
            select ҩƷid from ҩƷ��� where ҩ��id=ID_IN)
        loop
            update �շ���Ŀ���� SET ����=Ӣ��_IN WHERE �շ�ϸĿID=r_Spec.ҩƷid AND ����=2;
            if sql%rowcount=0 then
               insert into �շ���Ŀ����(�շ�ϸĿID,����,����,����,����) values(r_Spec.ҩƷid,Ӣ��_IN,2,null,0);
            end if;
        end loop;
    END IF;

    DELETE FROM ������Ŀ���� WHERE ������ĿID=ID_IN AND ����=9;
    DELETE FROM �շ���Ŀ���� 
    WHERE �շ�ϸĿid in (select ҩƷid from ҩƷ��� where ҩ��id=ID_IN) AND ����=9;
    IF ��������_IN IS NULL THEN
        v_Records := null;
    ELSE
        v_Records := ��������_IN||'|';
    END IF;
    WHILE v_Records IS NOT NULL LOOP
        v_CurrRec:=substr(v_Records,1,instr(v_Records,'|')-1);
        v_Fields:=v_CurrRec;
        v_����:=substr(v_Fields,1,instr(v_Fields,'^')-1);
        v_Fields:=substr(v_Fields,instr(v_Fields,'^')+1);
        v_ƴ��:=substr(v_Fields,1,instr(v_Fields,'^')-1);
        v_Fields:=substr(v_Fields,instr(v_Fields,'^')+1);
        v_���:=v_Fields;
        IF V_ƴ�� IS NOT NULL THEN
            INSERT INTO ������Ŀ����(������ĿID,����,����,����,����) VALUES (ID_IN,v_����,9,v_ƴ��,1);
            insert into �շ���Ŀ����(�շ�ϸĿID,����,����,����,����) 
            select ҩƷid,v_����,9,v_ƴ��,1 from ҩƷ��� where ҩ��id=ID_IN;
        END IF;
        IF v_��� IS NOT NULL THEN
            INSERT INTO ������Ŀ����(������ĿID,����,����,����,����) VALUES (ID_IN,v_����,9,v_���,2);
            insert into �շ���Ŀ����(�շ�ϸĿID,����,����,����,����) 
            select ҩƷid,v_����,9,v_���,2 from ҩƷ��� where ҩ��id=ID_IN;
        END IF;
        v_Records:=replace('|'||v_Records,'|'||v_CurrRec||'|');
    END LOOP;

EXCEPTION
    WHEN Err_NotFind THEN
        Raise_application_error (-20101, '[ZLSOFT]��Ʒ�ֲ����ڣ������ѱ������û�ɾ����[ZLSOFT]');
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_��ҩƷ��_UPDATE;
/


-----------------------------------------------------------
-- �����⹺������˴���
--˵�������ȶԲ����շ���¼���е�����˺����ʱ����д���
--���ŶԲ��Ͽ��Ͳ����շ����ܱ��е���Ӧ�����ͽ����д���
--�������
--���룺NO_IN,�����_IN
------------------------------------------------------------

CREATE OR REPLACE PROCEDURE ZL_�����⹺_VERIFY (
    NO_IN        IN ҩƷ�շ���¼.NO%TYPE := NULL,
    �����_IN    IN ҩƷ�շ���¼.�����%TYPE := NULL
)
IS
    mErrItem        EXCEPTION;
    mErrMsg        varchar2(100);

    V_BATCHCOUNT    INTEGER;        --ԭ���������ڷ����Ĳ��ϵ�����
    V_��λID        ҩƷ�շ���¼.��ҩ��λID%TYPE;

    V_��Ʊ���    Ӧ����¼.��Ʊ���%TYPE;
    V_�����    NUMBER(16,5);
    V_�����    NUMBER(16,5);
    V_�������    NUMBER(16,5);
    V_�ɱ���        NUMBER(16,5);

    CURSOR C_ҩƷ�շ���¼    IS
    SELECT ID,ʵ������,���۽��,���,�ⷿID,ҩƷID,����,��ҩ��λID,�ɱ���,����,Ч��,���Ч��,��������,����,������ID
    FROM ҩƷ�շ���¼
    WHERE NO = NO_IN AND ���� = 15    AND ��¼״̬ = 1
    ORDER BY ҩƷID;
BEGIN

    UPDATE ҩƷ�շ���¼
    SET ����� = NVL (�����_IN,�����),������� = SYSDATE
    WHERE NO = NO_IN AND ���� = 15 AND ��¼״̬ = 1 AND ����� IS NULL;

    IF SQL%ROWCOUNT = 0 THEN
        mErrMsg:='[ZLSOFT]�õ����Ѿ���������˻�ɾ�������ܽ�����ˣ�[ZLSOFT]';
        RAISE mErrItem;
    END IF;

    --��Ҫ���ԭ���������ڷ����Ĳ��ϣ����ܶ������
    SELECT COUNT(*) INTO V_BATCHCOUNT 
    FROM    ҩƷ�շ���¼ A,�������� B
    WHERE    A.ҩƷID=B.����ID AND A.NO=NO_IN     AND A.����=15 AND A.��¼״̬=1    AND NVL(A.����,0)=0
        AND ((NVL(B.�ⷿ����,0)=1 AND 
        A.�ⷿID NOT IN (SELECT ����ID FROM  ��������˵�� 
                 WHERE (�������� LIKE '���ϲ���') OR (�������� LIKE '�Ƽ���'))) OR NVL(B.���÷���,0)=1);

    IF V_BATCHCOUNT>0 THEN
        mErrMsg:='[ZLSOFT]�õ����а�����һ��ԭ�������������ڷ����Ĳ��ϣ�������ˣ�[ZLSOFT]';
        RAISE mErrItem;
    END IF;

    --ԭ�����ֲ������Ĳ���,�����ʱ��Ҫ������
    UPDATE ҩƷ�շ���¼ SET ����=0
    WHERE    ID IN (    SELECT ID FROM ҩƷ�շ���¼ A,�������� B
            WHERE A.ҩƷID=B.����ID    AND A.NO=NO_IN    AND A.���� = 15    AND A.��¼״̬ = 1
                AND NVL(A.����,0)>0 AND (NVL(B.�ⷿ����,0)=0 OR    (NVL(B.���÷���,0)=0 AND
                A.�ⷿID IN (SELECT ����ID FROM  ��������˵�� WHERE (�������� LIKE '���ϲ���') OR (�������� LIKE '�Ƽ���'))))
            );

    FOR V_ҩƷ�շ���¼ IN C_ҩƷ�շ���¼ LOOP
        --���Ĳ��Ͽ������Ӧ����

        UPDATE ҩƷ���
        SET    �������� = NVL (��������,0) + NVL (V_ҩƷ�շ���¼.ʵ������,0),
            ʵ������ = NVL (ʵ������,0) + NVL (V_ҩƷ�շ���¼.ʵ������,0),
            ʵ�ʽ�� = NVL (ʵ�ʽ��,0) + NVL (V_ҩƷ�շ���¼.���۽��,0),
            ʵ�ʲ�� = NVL (ʵ�ʲ��,0) + NVL (V_ҩƷ�շ���¼.���,0),
            �ϴι�Ӧ��ID = NVL (V_ҩƷ�շ���¼.��ҩ��λID,�ϴι�Ӧ��ID),
            �ϴβɹ��� = NVL (V_ҩƷ�շ���¼.�ɱ���,�ϴβɹ���),
            �ϴ����� = NVL (V_ҩƷ�շ���¼.����,�ϴ�����),
            �ϴβ��� = NVL (V_ҩƷ�շ���¼.����,�ϴβ���),
            ���Ч��=NVL (V_ҩƷ�շ���¼.���Ч��,���Ч��),
            �ϴ��������� = NVL (V_ҩƷ�շ���¼.��������,�ϴ���������),
            Ч�� = NVL (V_ҩƷ�շ���¼.Ч��,Ч��)
        WHERE �ⷿID = V_ҩƷ�շ���¼.�ⷿID
            AND ҩƷID = V_ҩƷ�շ���¼.ҩƷID
            AND NVL (����,0) = NVL (V_ҩƷ�շ���¼.����,0)
            AND ���� = 1;

        IF SQL%NOTFOUND THEN
            INSERT INTO ҩƷ���
                (�ⷿID,ҩƷID,����,   ����,��������,ʵ������,ʵ�ʽ��,ʵ�ʲ��,�ϴι�Ӧ��ID,�ϴβɹ���,�ϴ�����,�ϴ���������,�ϴβ���,���Ч��,Ч��)
            VALUES (
                V_ҩƷ�շ���¼.�ⷿID,
                V_ҩƷ�շ���¼.ҩƷID,
                V_ҩƷ�շ���¼.����,
                1,
                V_ҩƷ�շ���¼.ʵ������,
                V_ҩƷ�շ���¼.ʵ������,
                V_ҩƷ�շ���¼.���۽��,
                V_ҩƷ�շ���¼.���,
                V_ҩƷ�շ���¼.��ҩ��λID,
                V_ҩƷ�շ���¼.�ɱ���,
                V_ҩƷ�շ���¼.����,
                V_ҩƷ�շ���¼.��������,
                V_ҩƷ�շ���¼.����,
                V_ҩƷ�շ���¼.���Ч��,
                V_ҩƷ�շ���¼.Ч��);
        END IF;

        --����������Ϊ��ļ�¼
        DELETE
        FROM ҩƷ���
        WHERE �ⷿID = V_ҩƷ�շ���¼.�ⷿID
            AND ҩƷID = V_ҩƷ�շ���¼.ҩƷID
            AND NVL(��������,0) = 0
            AND NVL(ʵ������,0) = 0
            AND NVL(ʵ�ʽ��,0) = 0
            AND NVL(ʵ�ʲ��,0) = 0;

        --���Ĳ����շ����ܱ����Ӧ����

        UPDATE ҩƷ�շ����� SET 
            ���� = NVL (����,0) + NVL (V_ҩƷ�շ���¼.ʵ������,0),
            ��� = NVL (���,0) + NVL (V_ҩƷ�շ���¼.���۽��,0),
            ��� = NVL (���,0) + NVL (V_ҩƷ�շ���¼.���,0)
        WHERE ���� = TRUNC (SYSDATE)
            AND �ⷿID = V_ҩƷ�շ���¼.�ⷿID
            AND ҩƷID = V_ҩƷ�շ���¼.ҩƷID
            AND ���ID = V_ҩƷ�շ���¼.������ID
            AND ���� = 15;

        IF SQL%NOTFOUND THEN
            INSERT INTO ҩƷ�շ�����
                (����,�ⷿID,ҩƷID,���ID,����,����,���,���)
            VALUES (
                  TRUNC (SYSDATE),
                  V_ҩƷ�շ���¼.�ⷿID,
                  V_ҩƷ�շ���¼.ҩƷID,
                  V_ҩƷ�շ���¼.������ID,
                  15,
                  V_ҩƷ�շ���¼.ʵ������,
                  V_ҩƷ�շ���¼.���۽��,
                  V_ҩƷ�շ���¼.���
                  );
        END IF;

        --���¸ò��ϵĳɱ���
        BEGIN 
            SELECT SUM(NVL(ʵ�ʽ��,0)),SUM(NVL(ʵ�ʲ��,0)),SUM(NVL(ʵ������,0))
                INTO V_�����,V_�����,V_�������
            FROM ҩƷ���
            WHERE ����=1 and ҩƷID=V_ҩƷ�շ���¼.ҩƷID;
        EXCEPTION 
            WHEN OTHERS THEN V_�������:=0;
        END ;

	--���¸�ҩƷ�ĳɱ���
	UPDATE ��������
	SET �ɱ���=V_ҩƷ�շ���¼.�ɱ��� 
	WHERE ����ID=V_ҩƷ�շ���¼.ҩƷID;

    END LOOP;


    --��Ӧ��������д���
    --�˴���һ���飬��Ҫ�ǽ��û�ж�Ӧ��Ʊ�ŵļ�¼
    BEGIN
        UPDATE Ӧ����¼
        SET �����=�����_IN,�������=SYSDATE
        WHERE ��ⵥ�ݺ�=NO_IN AND ϵͳ��ʶ=5 and ��¼����=0 And ��¼״̬=1;

        SELECT B.��λID,SUM (��Ʊ���)    INTO V_��λID,V_��Ʊ���
        FROM ҩƷ�շ���¼ A,Ӧ����¼ B
        WHERE A.ID = B.�շ�ID
            AND A.NO = NO_IN
            AND A.���� = 15 AND B.ϵͳ��ʶ=5
        GROUP BY B.��λID;

        IF NVL (V_��λID,0) <> 0 THEN
            UPDATE Ӧ�����    SET 
                ��� = NVL (���,0) + NVL (V_��Ʊ���,0)
            WHERE ��λID = V_��λID    AND ���� = 1;

            IF SQL%NOTFOUND THEN
                INSERT INTO Ӧ����� (��λID,����,���)
                VALUES (V_��λID,1,V_��Ʊ���);
            END IF;
        END IF;
    EXCEPTION
        WHEN NO_DATA_FOUND THEN
            NULL;
    END;

EXCEPTION
    WHEN mErrItem THEN
        RAISE_APPLICATION_ERROR (-20101,mErrMsg);
    WHEN OTHERS THEN
        ZL_ERRORCENTER (SQLCODE,SQLERRM);
END ZL_�����⹺_VERIFY;
/
-----------------------------------------------------------
-- �����⹺���ĳ�������
--˵�������ȸ�ԭ���ݵļ�¼״̬Ϊ3��+3;
--������һ�ŵ��ݺ���ͬ����¼״̬Ϊ2��+3�������ͽ��Ϊ���ĳ�������;
--ͬʱ������Ӧ��Ӧ����¼;
--������ҩƷ�����ҩƷ�շ����ܱ����Ӧ�����ͽ�
--�������
--���룺NO_IN,�����_IN
------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_�����⹺_STRIKE (
    �д�_IN        IN INTEGER,
    ԭ��¼״̬_IN    IN ҩƷ�շ���¼.��¼״̬%TYPE,
    NO_IN        IN ҩƷ�շ���¼.NO%TYPE,
    ���_IN        IN ҩƷ�շ���¼.���%TYPE,
    ����id_IN    IN ҩƷ�շ���¼.ҩƷID%TYPE,
    ��������_IN    IN ҩƷ�շ���¼.ʵ������%TYPE,
    ������_IN    IN ҩƷ�շ���¼.������%TYPE ,
    ��������_IN    IN ҩƷ�շ���¼.��������%TYPE ,
    ��Ʊ��_IN    IN Ӧ����¼.��Ʊ��%TYPE := NULL,
    ��Ʊ����_IN    IN Ӧ����¼.��Ʊ����%TYPE := NULL,
    ��Ʊ���_IN    IN Ӧ����¼.��Ʊ���%TYPE := NULL,
    ȫ������_IN    IN ҩƷ�շ���¼.ʵ������%TYPE := 0    --���ڲ������
) 
IS
    mErrItem    EXCEPTION ;
    mErrMsg        varchar2(100);

    V_BATCHCOUNT    INTEGER;    --ԭ���������ڷ����Ĳ��ϵ����� 
    v_Ӧ��ID        Ӧ����¼.ID%TYPE;
    V_�ⷿID        ҩƷ�շ���¼.�ⷿID%TYPE; 
    V_��ҩ��λID    ҩƷ�շ���¼.��ҩ��λID%TYPE; 
    V_������ID    ҩƷ�շ���¼.������ID%TYPE ;
    V_����        ҩƷ�շ���¼.����%TYPE ; 
    V_����        ҩƷ�շ���¼.����%TYPE ; 
    V_����        ҩƷ�շ���¼.����%TYPE ; 
    V_Ч��        ҩƷ�շ���¼.Ч��%TYPE ; 
    V_�ɱ���        ҩƷ�շ���¼.�ɱ���%TYPE ; 
    V_�ɱ����    ҩƷ�շ���¼.�ɱ����%TYPE ; 
    V_����        ҩƷ�շ���¼.����%TYPE ; 
    V_���ۼ�        ҩƷ�շ���¼.���ۼ�%TYPE ; 
    V_���۽��    ҩƷ�շ���¼.���۽��%TYPE ; 
    V_���        ҩƷ�շ���¼.���%TYPE ; 
    V_ժҪ        ҩƷ�շ���¼.ժҪ%TYPE ; 
    V_ʣ������    ҩƷ�շ���¼.ʵ������%TYPE; 
    V_ʣ��ɱ���� ҩƷ�շ���¼.�ɱ����%Type;
    V_ʣ�����۽�� ҩƷ�շ���¼.���۽��%Type;

    V_���ϵ��    ҩƷ�շ���¼.���ϵ��%TYPE; 
    V_��������    ҩƷ�շ���¼.ʵ������%TYPE;
    V_���Ч��    ҩƷ�շ���¼.���Ч��%TYPE; 
    v_�������    ҩƷ�շ���¼.�������%TYPE; 
    v_��������    ҩƷ�շ���¼.��������%TYPE; 

    V_��¼�� NUMBER; 
    V_�շ�ID        ҩƷ�շ���¼.ID%TYPE; 

    --�Գ����������м��
    V_�����        ҩƷ���.ʵ������%TYPE;
    V_�ⷿ����    INTEGER;
    V_���÷���    INTEGER;
    V_��������    INTEGER;
    v_�ⷿ        INTEGER;
    v_��¼״̬	  ҩƷ�շ���¼.��¼״̬%type;
    V_����        NUMBER;
    V_С��		number(2);
    V_��Ʊ���	  NUMBER(16,5);

BEGIN
    BEGIN 
	    SELECT  to_number(Nvl(����ֵ,ȱʡֵ),'99999')  INTO v_С�� 
	    From ϵͳ������ where ������='���ý���λ��';
    EXCEPTION 
	WHEN OTHERS THEN 
		v_С��:=2;		
    END;

    V_��������:=��������_IN;
    IF �д�_IN =1 THEN
        UPDATE ҩƷ�շ���¼ 
        SET ��¼״̬ = DECODE(ԭ��¼״̬_IN,1,3,ԭ��¼״̬_IN+3) 
        WHERE NO = NO_IN 
            AND ���� = 15 
            AND ��¼״̬ =ԭ��¼״̬_IN ; 

        IF SQL%ROWCOUNT = 0 THEN 
            mErrMsg:='[ZLSOFT]�õ����Ѿ������˳�����[ZLSOFT]';
            RAISE mErrItem; 
        END IF; 
    END IF;
    
    --��Ҫ���ԭ���������ڷ����Ĳ��ϣ����ܶ������ 
    SELECT COUNT(*) INTO V_BATCHCOUNT 
    FROM     ҩƷ�շ���¼ A,�������� b
    WHERE A.ҩƷID=B.����id    AND A.NO=NO_IN     AND A.����=15 AND MOD(A.��¼״̬,3)=0
        AND NVL(A.����,0)=0 AND A.ҩƷID+0=����id_IN
        AND ((NVL(B.���÷���,0)=1 AND 
        A.�ⷿID NOT IN (SELECT ����ID FROM  ��������˵�� WHERE (�������� LIKE '���ϲ���') OR (�������� LIKE '�Ƽ���')))    or nvl(b.���÷���,0)=1); 
    
    IF V_BATCHCOUNT>0 THEN 
        mErrMsg:='[ZLSOFT]�õ����е�' || ���_IN || '�еĲ���ԭ�������������ڷ����Ĳ��ϣ����ܳ�����[ZLSOFT]';
        RAISE mErrItem; 
    END IF;
        
    SELECT SUM(A.ʵ������) AS ʣ������,SUM(A.�ɱ����) AS ʣ��ɱ����,SUM(A.���۽��) AS ʣ�����۽��,A.�ⷿID,A.��ҩ��λID,A.������ID,A.���ϵ��,NVL(A.����,0),A.����,A.����,A.Ч��,A.���Ч��,A.�������,A.��������,A.�ɱ���,A.����,A.���ۼ�,A.ժҪ,B.�ⷿ����,B.���÷���
        INTO  V_ʣ������,v_ʣ��ɱ����,V_ʣ�����۽��,V_�ⷿID,V_��ҩ��λID,V_������ID,V_���ϵ��,V_����,V_����,V_����,V_Ч��,v_���Ч��,v_�������,v_��������,V_�ɱ���,V_����,V_���ۼ�,V_ժҪ,V_�ⷿ����,V_���÷���
    FROM ҩƷ�շ���¼ A,�������� B
    WHERE A.NO=NO_IN And A.ҩƷID=B.����ID AND A.����=15 AND A.ҩƷID+0=����id_IN AND A.���=���_IN
    GROUP BY A.�ⷿID,A.��ҩ��λID,A.������ID,A.���ϵ��,NVL(A.����,0),A.����,A.����,A.Ч��,A.���Ч��,A.�������,A.��������,A.�ɱ���,A.����,A.���ۼ�,A.ժҪ,B.�ⷿ����,B.���÷���;

    --�жϸò����ǿⷿ���Ƿ��ϲ���
    BEGIN
        SELECT DISTINCT 0 INTO v_�ⷿ
        FROM ��������˵��
        WHERE (�������� ='���ϲ���' OR �������� = '�Ƽ���')
        AND ����ID = V_�ⷿID;
    EXCEPTION 
        WHEN OTHERS THEN v_�ⷿ:=1;
    END ;
    
    --���ݲ�������,�жϷ�������
    IF v_�ⷿ=0 THEN 
        v_��������:=V_���÷���;
    ELSE
        V_��������:=V_�ⷿ����;
    END IF ;

    --V_����:(ԭ����,�ַ���Ϊ����;����Ϊ��)ԭ������,�ַ���,�����̲�����
    V_����:=0;
    IF V_��������=1 AND V_����<>0 THEN 
        V_����:=V_����;
    END IF ;
    
    --ȡ�����
    BEGIN
        SELECT Nvl(ʵ������,0) INTO V_����� 
        FROM ҩƷ��� 
        WHERE �ⷿID=V_�ⷿID AND ҩƷID=����id_IN AND Nvl(����,0)=V_���� And ����=1;
    EXCEPTION 
    WHEN OTHERS THEN V_�����:=0;
    END ;

    IF nvl(V_ʣ������,0)=0 THEN 
            mErrMsg:='[ZLSOFT]�õ����е�' || ���_IN || '�еĲ����Ѿ����������,�����ٳ壡[ZLSOFT]';
            RAISE mErrItem; 
    END IF ;

    --������������ʣ������,ȡʣ������;����ȡ�����
    IF V_�����<V_ʣ������ THEN 
        if ȫ������_IN=1 then 
            --������
            mErrMsg:='[ZLSOFT]�õ����е�' || ���_IN || '�еĲ��ϳ���������������ʣ������ݣ����ܳ�����[ZLSOFT]';
            RAISE mErrItem; 
        else
            v_ʣ��ɱ����:=V_�����/V_ʣ������*v_ʣ��ɱ����;
            V_ʣ�����۽��:=V_�����/V_ʣ������*V_ʣ�����۽��;
            V_ʣ������:=V_�����;
 
        end if ;
    END IF ;
    
    IF ȫ������_IN=1  THEN 
        V_��������:=V_ʣ������;
    END IF;

    --������������ʣ��������������
    IF V_ʣ������<V_��������  THEN
        mErrMsg:='[ZLSOFT]�õ����е�' || ���_IN || '�еĲ��ϳ���������������ʣ������ݣ����ܳ�����[ZLSOFT]';
        RAISE mErrItem; 
    END IF;

    V_�ɱ����:= ROUND(V_��������/v_ʣ������*v_ʣ��ɱ����,v_С��);
    V_���۽��:= ROUND(V_��������/v_ʣ������*V_ʣ�����۽��,v_С��);
    V_���:=round(V_���۽��-V_�ɱ����,v_С��);

    SELECT ҩƷ�շ���¼_ID.NEXTVAL INTO V_�շ�ID  FROM DUAL;
         
    INSERT INTO ҩƷ�շ���¼ 
    ( ID,��¼״̬,����,NO,���,�ⷿID,��ҩ��λID,������ID,���ϵ��,ҩƷID,����,����,����,��������,Ч��,���Ч��,�������,
    ��д����,ʵ������,�ɱ���,�ɱ����,����,���ۼ�,���۽��,���,ժҪ,������,��������,�����,������� 
    ) 
    VALUES (V_�շ�ID,DECODE(ԭ��¼״̬_IN,1,2,ԭ��¼״̬_IN+2),15,NO_IN,���_IN,V_�ⷿID,V_��ҩ��λID,
    V_������ID,1,����id_IN,V_����,V_����,V_����,v_��������,V_Ч��,v_���Ч��,v_�������,-V_��������,-V_��������,V_�ɱ���,-V_�ɱ����,
    V_����,V_���ۼ�,-V_���۽��,-V_���,V_ժҪ,������_IN,��������_IN,������_IN,��������_IN); 


    --���ڳ����ĵ���ҲӦ�ö�Ӧ��������д��� 
    --ֻ�����˷�Ʊ�ŵļ�¼���д��� 
    V_��Ʊ���:=NVL (��Ʊ���_IN,0);

    IF NVL (��Ʊ��_IN,' ') <> ' ' AND NVL (V_��Ʊ���,0)<>0 THEN 
	--���ڲ�����˵ģ�Ҫ��ʣ��ķ�Ʊ���ȫ������
	IF ȫ������_IN=1 THEN
		SELECT SUM(B.��Ʊ���) INTO v_��Ʊ���
		FROM
			(SELECT ID
			FROM ҩƷ�շ���¼
			WHERE ����=15 AND NO=NO_IN AND ���=���_IN) A,Ӧ����¼ B
		WHERE A.ID=B.�շ�ID AND B.ϵͳ��ʶ=5 And B.��¼����<>-1;
	END IF;

	UPDATE Ӧ�����    SET ��� = NVL (���,0) - NVL (v_��Ʊ���,0)
	WHERE ��λID = V_��ҩ��λID  AND ���� = 1; 

        IF SQL%NOTFOUND THEN 
            INSERT INTO Ӧ����� (��λID,����,���) VALUES (V_��ҩ��λID,1,-NVL (v_��Ʊ���,0)); 
        END IF; 
    END IF; 
    
    UPDATE ҩƷ��� 
    SET �������� = NVL (��������,0) -V_��������,
         ʵ������ = NVL (ʵ������,0) - V_��������,
         ʵ�ʽ�� = NVL (ʵ�ʽ��,0) - V_���۽��,
         ʵ�ʲ�� = NVL (ʵ�ʲ��,0) -V_���,
         �ϴι�Ӧ��ID = V_��ҩ��λID,
         �ϴβɹ��� = V_�ɱ���,
         �ϴ����� = V_����,
         �ϴβ��� = V_����,
         ���Ч��=v_���Ч��,
	       �ϴ���������=v_��������,
         Ч�� = V_Ч�� 
    WHERE �ⷿID = V_�ⷿID 
        AND ҩƷID = ����id_IN 
        AND NVL (����,0) = NVL(V_����,0) 
        AND ���� = 1; 
 
    IF SQL%NOTFOUND THEN 
        INSERT INTO ҩƷ��� 
            (�ⷿID,ҩƷID,����,����,��������,ʵ������,ʵ�ʽ��,ʵ�ʲ��,�ϴι�Ӧ��ID,�ϴβɹ���,�ϴ�����,�ϴβ���,Ч��,���Ч��,�ϴ���������) 
        VALUES 
            (V_�ⷿID,����id_IN,V_����,1,-V_��������,-V_��������,-V_���۽��,-V_���,V_��ҩ��λID,V_�ɱ���,V_����,V_����,V_Ч��,v_���Ч��,v_��������) ;
    END IF; 
 
    --����������Ϊ��ļ�¼ 
    DELETE     FROM ҩƷ��� 
    WHERE �ⷿID = V_�ⷿID 
        AND ҩƷID = ����id_IN 
        AND NVL (��������,0) = 0 
        AND NVL (ʵ������,0) = 0 
        AND NVL (ʵ�ʽ��,0) = 0 
        AND NVL (ʵ�ʲ��,0) = 0; 
 
    --����ҩƷ�շ����ܱ����Ӧ���� 
    UPDATE ҩƷ�շ����� 
    SET ���� =NVL(����,0) - V_��������,
        ��� =NVL (���,0) -V_���۽��,
        ��� =NVL (���,0) -V_��� 
    WHERE ���� = TRUNC (SYSDATE) 
        AND �ⷿID = V_�ⷿID 
        AND ҩƷID = ����id_IN 
        AND ���ID = V_������ID 
        AND ���� = 15; 

    IF SQL%NOTFOUND THEN 
        INSERT INTO ҩƷ�շ����� 
            (����,�ⷿID,ҩƷID,���ID,����,����,���,���) 
        VALUES ( 
            TRUNC (SYSDATE),V_�ⷿID,����id_IN,V_������ID,15,-V_��������,-V_���۽��,-V_���); 
    END IF; 

    --����Ӧ����¼�ĳ�����¼(���ж�Ӧ����¼���Ƿ��Ѵ��ڸü�¼��Ӧ�ĳ�����¼,�������;��������)
    SELECT Ӧ����¼_ID.NEXTVAL INTO V_Ӧ��ID FROM DUAL;

	begin 
		select max(��¼״̬)+3 into v_��¼״̬
		from Ӧ����¼ 
		where (ϵͳ��ʶ,��¼����,no,��Ŀid,���) in (	select ϵͳ��ʶ,��¼����,NO,��Ŀid,���
								from Ӧ����¼
								where �շ�ID=(SELECT ID FROM ҩƷ�շ���¼ WHERE NO=NO_IN AND ����=15 AND ���=���_IN And Mod(��¼״̬,3)=0) AND ϵͳ��ʶ=5 and  ��¼����=0)
		       and ��¼״̬<>1 and mod(��¼״̬,3)<>0;

	exception 
		when others then 
			v_��¼״̬:=2;
	end ;
	if v_��¼״̬ is null then 
		v_��¼״̬:=2;
	end if ;
	if mod(v_��¼״̬,3)<>2 then 
		v_��¼״̬:=v_��¼״̬+1;
	end if ;
	if mod(v_��¼״̬,3)<>2 then 
		v_��¼״̬:=v_��¼״̬+1;
	end if ;
	
    INSERT INTO Ӧ����¼
    (ID,��¼����,��¼״̬,��ĿID,���,��λID,NO,ϵͳ��ʶ,�շ�ID,��ⵥ�ݺ�,���ݽ��,��Ʊ��,��Ʊ����,��Ʊ���,Ʒ��,
		���,����,����,������λ,����,�ɹ���,�ɹ����,������,��������,�����,�������,ժҪ)
    SELECT V_Ӧ��ID,��¼����,v_��¼״̬,����ID_In,���_IN,��λID,NO,5,V_�շ�ID,��ⵥ�ݺ�,-1* V_���۽��,��Ʊ��,��Ʊ����,-v_��Ʊ���,Ʒ��,
    ���,����,����,������λ,-1*V_��������,�ɹ���,-1* �ɹ���*V_��������,������_IN,��������_IN,������_IN,��������_IN,ժҪ
    FROM Ӧ����¼
    WHERE �շ�ID=(SELECT ID FROM ҩƷ�շ���¼ WHERE NO=NO_IN AND ����=15 AND ���=���_IN And Mod(��¼״̬,3)=0) AND ϵͳ��ʶ=5 AND ��¼����=0;

    update Ӧ����¼    set ��¼״̬=3
    WHERE �շ�ID=(SELECT ID FROM ҩƷ�շ���¼ WHERE NO=NO_IN AND ����=15 AND ���=���_IN And Mod(��¼״̬,3)=0) AND ϵͳ��ʶ=5 AND ��¼����=0;
    
EXCEPTION 
    WHEN mErrItem THEN 
        RAISE_APPLICATION_ERROR ( -20101,mErrMsg); 
    WHEN OTHERS THEN 
        ZL_ERRORCENTER (SQLCODE,SQLERRM); 
END ZL_�����⹺_STRIKE; 
/

-----------------------------------------------------------
--�����������ĳ�������
--˵�������ȸ�ԭ���ݵļ�¼״̬Ϊ3;
--������һ�ŵ��ݺ���ͬ����¼״̬Ϊ2�������ͽ��Ϊ���ĳ�������;
--������ҩƷ�����ҩƷ�շ����ܱ����Ӧ�����ͽ�
--�������
--���룺NO_IN,�����_IN
------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_�����������_strike (
    �д�_IN        IN INTEGER,
    ԭ��¼״̬_IN    IN ҩƷ�շ���¼.��¼״̬%TYPE,
    NO_IN        IN ҩƷ�շ���¼.NO%TYPE,
    ���_IN        IN ҩƷ�շ���¼.���%TYPE,
    ����ID_IN    IN ҩƷ�շ���¼.ҩƷID%TYPE,
    ��������_IN    IN ҩƷ�շ���¼.ʵ������%TYPE,
    ������_IN    IN ҩƷ�շ���¼.������%TYPE,
    ��������_IN    IN ҩƷ�շ���¼.��������%TYPE
)
IS
    mErrMsg            varchar2(100);
    mErrItem        EXCEPTION;

    v_BatchCount        INTEGER;    --ԭ���������ڷ����Ĳ��ϵ�����

    V_�ⷿID            ҩƷ�շ���¼.�ⷿID%TYPE; 
    V_������ID        ҩƷ�շ���¼.������ID%TYPE ;
    V_����            ҩƷ�շ���¼.����%TYPE ; 
    V_����            ҩƷ�շ���¼.����%TYPE ; 
    V_����            ҩƷ�շ���¼.����%TYPE ; 
    V_��������        ҩƷ�շ���¼.��������%TYPE ; 
    V_Ч��            ҩƷ�շ���¼.Ч��%TYPE ; 
    V_�ɱ���            ҩƷ�շ���¼.�ɱ���%TYPE ; 
    V_�ɱ����        ҩƷ�շ���¼.�ɱ����%TYPE ; 
    V_����            ҩƷ�շ���¼.����%TYPE ; 
    V_���ۼ�            ҩƷ�շ���¼.���ۼ�%TYPE ; 
    V_���۽��        ҩƷ�շ���¼.���۽��%TYPE ; 
    V_���            ҩƷ�շ���¼.���%TYPE ; 
    V_ժҪ            ҩƷ�շ���¼.ժҪ%TYPE ; 
    V_���ϵ��        ҩƷ�շ���¼.���ϵ��%TYPE; 
    V_���Ч��        ҩƷ�շ���¼.���Ч��%type;
    v_�������        ҩƷ�շ���¼.�������%type;
    V_��¼��            NUMBER; 
    V_�շ�ID            ҩƷ�շ���¼.ID%TYPE; 
    V_ʣ������		ҩƷ�շ���¼.ʵ������%TYPE; 
    V_ʣ��ɱ����	ҩƷ�շ���¼.�ɱ����%Type;
    V_ʣ�����۽��	ҩƷ�շ���¼.���۽��%Type;

    --�Գ����������м��
    V_�����            ҩƷ���.ʵ������%TYPE;
    V_�ⷿ����        INTEGER;
    V_���÷���        INTEGER;
    V_��������        INTEGER;
    v_�ⷿ            INTEGER;
    V_����            NUMBER;
    V_С��		number(2);
BEGIN
    BEGIN 
	    SELECT  to_number(Nvl(����ֵ,ȱʡֵ),'99999')  INTO v_С�� 
	    From ϵͳ������ where ������='���ý���λ��';
    EXCEPTION 
	WHEN OTHERS THEN 
		v_С��:=2;		
    END;

    IF �д�_IN =1 THEN
        UPDATE ҩƷ�շ���¼ 
        SET ��¼״̬ = DECODE(ԭ��¼״̬_IN,1,3,ԭ��¼״̬_IN+3) 
        WHERE NO = NO_IN AND ���� = 17 AND ��¼״̬ =ԭ��¼״̬_IN ; 

        IF SQL%ROWCOUNT = 0 THEN 
            mErrMsg:='[ZLSOFT]�õ����Ѿ������˳�����[ZLSOFT]';
            RAISE mErrItem; 
        END IF; 
    END IF;
    
    --��Ҫ���ԭ���������ڷ����Ĳ��ϣ����ܶ������
    SELECT COUNT(*) INTO v_BatchCount 
    FROM ҩƷ�շ���¼ a,�������� b
    WHERE a.ҩƷid=b.����id    AND a.no=NO_IN     AND a.����=17 AND MOD(a.��¼״̬,3)=0 AND a.ҩƷID+0=����ID_IN
        AND nvl(a.����,0)=0
        AND ((nvl(b.�ⷿ����,0)=1 AND a.�ⷿid not in (select ����id from  ��������˵�� where (�������� LIKE '���ϲ���') or (�������� LIKE '�Ƽ���')))
        or nvl(b.���÷���,0)=1);

    IF v_batchcount>0 THEN
        mErrMsg:='[ZLSOFT]�õ����а�����һ��ԭ�������������ڷ����Ĳ��ϣ����ܳ�����[ZLSOFT]';
        RAISE mErrItem; 
    END IF;  

    SELECT SUM(A.ʵ������) AS ʣ������,SUM(A.�ɱ����) AS ʣ��ɱ����,SUM(A.���۽��) AS ʣ�����۽��,A.�ⷿID,A.������ID,A.���ϵ��,Nvl(A.����,0),A.����,A.����,A.��������,A.Ч��,A.���Ч��,A.�������,A.�ɱ���,A.����,A.���ۼ�,
        A.ժҪ,B.�ⷿ����,B.���÷���    INTO  V_ʣ������,v_ʣ��ɱ����,V_ʣ�����۽��,V_�ⷿID,V_������ID,V_���ϵ��,V_����,V_����,V_����,V_��������,V_Ч��,V_���Ч��,v_�������,
        V_�ɱ���,V_����,V_���ۼ�,V_ժҪ,V_�ⷿ����,V_���÷���
    FROM ҩƷ�շ���¼ A,�������� B
    WHERE A.NO=NO_IN AND A.����=17 AND A.ҩƷID=B.����ID AND A.ҩƷID+0=����ID_IN AND A.���=���_IN
    GROUP BY A.�ⷿID,A.������ID,A.���ϵ��,NVL(A.����,0),A.����,A.����,A.��������,A.Ч��,A.���Ч��,a.�������,A.�ɱ���,A.����,A.���ۼ�,A.ժҪ,B.�ⷿ����,B.���÷���;

    --�жϸò����ǿⷿ���Ƿ��ϲ���
    BEGIN
        SELECT DISTINCT 0 INTO v_�ⷿ 
        FROM ��������˵��
        WHERE ((�������� LIKE '���ϲ���') OR (�������� LIKE '�Ƽ���')) AND ����ID = V_�ⷿID;
    EXCEPTION 
        WHEN OTHERS THEN v_�ⷿ:=1;
    END ;
    
    --���ݲ�������,�жϷ�������
    IF v_�ⷿ=0 THEN 
        v_��������:=V_���÷���;
    ELSE
        V_��������:=V_�ⷿ����;
    END IF ;

    --V_����:(ԭ����,�ַ���Ϊ����;����Ϊ��)ԭ������,�ַ���,�����̲�����
    V_����:=0;
    IF V_��������=1 AND V_����<>0 THEN 
        V_����:=V_����;
    END IF ;
    
    --ȡ�����
    BEGIN
        SELECT Nvl(ʵ������,0) INTO V_����� FROM ҩƷ��� 
        WHERE �ⷿID=V_�ⷿID AND ҩƷID=����ID_IN AND Nvl(����,0)=V_���� And ����=1;
    EXCEPTION 
        WHEN OTHERS THEN V_�����:=0;
    END ;

    IF nvl(V_ʣ������,0)=0 THEN 
            mErrMsg:='[ZLSOFT]�õ����е�' || ���_IN || '�еĲ����Ѿ����������,�����ٳ壡[ZLSOFT]';
            RAISE mErrItem; 
    END IF ;

  
    --������������ʣ������,ȡʣ������;����ȡ�����
    IF V_�����<V_ʣ������ THEN 
	v_ʣ��ɱ����:=V_�����/V_ʣ������*v_ʣ��ɱ����;
	V_ʣ�����۽��:=V_�����/V_ʣ������*V_ʣ�����۽��;
	V_ʣ������:=V_�����;
    END IF ;

    --������������ʣ��������������
    IF V_ʣ������<��������_IN THEN
        mErrMsg:='[ZLSOFT]�õ����е�' || ���_IN || '�еĲ��ϳ���������������ʣ������ݣ����ܳ�����[ZLSOFT]';
        RAISE mErrItem; 
    END IF;

    V_�ɱ����:= ROUND(��������_IN/v_ʣ������*v_ʣ��ɱ����,v_С��);
    V_���۽��:= ROUND(��������_IN/v_ʣ������*V_ʣ�����۽��,v_С��);
    V_���:=round(V_���۽��-V_�ɱ����,v_С��);


    SELECT ҩƷ�շ���¼_ID.NEXTVAL INTO V_�շ�ID  FROM DUAL;

    Insert INTO ҩƷ�շ���¼
        (ID,��¼״̬,����,NO,���,�ⷿID,������ID,���ϵ��,ҩƷID,����,����,����,��������,
        Ч��,�������,���Ч��,��д����,ʵ������,�ɱ���,�ɱ����,���ۼ�,���۽��,���,ժҪ,������,��������,�����,�������)
    VALUES 
        (V_�շ�ID,DECODE(ԭ��¼״̬_IN,1,2,ԭ��¼״̬_IN+2),17,NO_IN,���_IN,V_�ⷿID,V_������ID,
        V_���ϵ��,����ID_IN,V_����,V_����,V_����,V_��������,V_Ч��,v_�������,v_���Ч��,-��������_IN,-��������_IN,V_�ɱ���,-V_�ɱ����,
        V_���ۼ�,-V_���۽��,-V_���,V_ժҪ,������_IN,��������_IN,������_IN,��������_IN);

    --����ҩƷ�������Ӧ����

    UPDATE ҩƷ���
    SET �������� = NVL (��������,0) - NVL (��������_IN,0),
        ʵ������ = NVL (ʵ������,0) - NVL (��������_IN,0),
        ʵ�ʽ�� = NVL (ʵ�ʽ��,0) - NVL (V_���۽��,0),
        ʵ�ʲ�� = NVL (ʵ�ʲ��,0) - NVL (V_���,0),
        �ϴβɹ��� = NVL (V_�ɱ���,�ϴβɹ���),
        �ϴ����� = NVL (V_����,�ϴ�����),
        �ϴβ��� = NVL (V_����,�ϴβ���),
	�ϴ���������=nvl(V_��������,�ϴ���������),
        ���Ч��=nvl(v_���Ч��,���Ч��),
        Ч�� = NVL (V_Ч��,Ч��)
    WHERE �ⷿID = V_�ⷿID
        AND ҩƷID = ����ID_IN
        AND NVL (����,0) = NVL (V_����,0)
        AND ���� = 1;

    IF SQL%NOTFOUND THEN
        Insert INTO ҩƷ���
            (�ⷿID,ҩƷID,����,����,��������,ʵ������,ʵ�ʽ��,
            ʵ�ʲ��,�ϴβɹ���,�ϴ�����,�ϴ���������,�ϴβ���,Ч��,���Ч��)
        VALUES (
            V_�ⷿID,����ID_IN,V_����,1,-��������_IN,-��������_IN,-V_���۽��,
            -V_���,V_�ɱ���,V_����,V_��������,V_����,V_Ч��,v_���Ч��);
    END IF;

    DELETE
    FROM ҩƷ���
    WHERE �ⷿID = V_�ⷿID
        AND ҩƷID = ����ID_IN
        AND nvl(��������,0) = 0
        AND nvl(ʵ������,0) = 0
        AND nvl(ʵ�ʽ��,0) = 0
        AND nvl(ʵ�ʲ��,0) = 0;

    --����ҩƷ�շ����ܱ����Ӧ����

    UPDATE ҩƷ�շ�����
    SET ���� = NVL (����,0) - NVL (��������_IN,0),
        ��� = NVL (���,0) - NVL (V_���۽��,0),
        ��� = NVL (���,0) - NVL (V_���,0)
    WHERE ���� = TRUNC (��������_IN)
        AND �ⷿID = V_�ⷿID
        AND ҩƷID = ����ID_IN
        AND ���ID = V_������ID
        AND ���� = 17;

    IF SQL%NOTFOUND THEN
        Insert INTO ҩƷ�շ�����
            (����,�ⷿID,ҩƷID,���ID,����,����,���,���)
        VALUES (
            TRUNC (��������_IN),V_�ⷿID,����ID_IN,
            V_������ID,17,-��������_IN,-V_���۽��,-V_���);
    END IF;
EXCEPTION
    WHEN mErrItem THEN
        Raise_application_error (-20101,mErrMsg);
    WHEN OTHERS THEN 
        zl_ErrorCenter (SQLCODE,SQLERRM);
END zl_�����������_strike;
/


-----------------------------------------------------------
--���Ͽ���۵����ĳ�������
--˵�������ȸ�ԭ���ݵļ�¼״̬Ϊ3;
--������һ�ŵ��ݺ���ͬ����¼״̬Ϊ2�������ͽ��Ϊ���ĳ�������;
--������ҩƷ�����ҩƷ�շ����ܱ����Ӧ�����ͽ�
--�������
--���룺NO_IN,�����_IN
------------------------------------------------------------

CREATE OR REPLACE PROCEDURE zl_���Ͽ���۵���_strike (
    NO_IN        IN ҩƷ�շ���¼.NO%TYPE,
    �����_IN    IN ҩƷ�շ���¼.�����%TYPE
)
IS
    mErrMsg            varchar2(100);
    mErrItem        EXCEPTION;
    v_BatchCount        integer;    --ԭ���������ڷ����Ĳ��ϵ�����
    V_COUNT            INTEGER;    --ԭ�����ֲ�����
    V_����            ҩƷ�շ���¼.����%TYPE;

    CURSOR C_ҩƷ�շ���¼    IS
    SELECT ������ID,�ⷿID,ҩƷid ����ID,����,���,����,Ч��,���Ч��,�������
    FROM ҩƷ�շ���¼ A
    WHERE NO = NO_IN
        AND ���� = 18
        AND ��¼״̬ = 2
    ORDER BY ҩƷID;
BEGIN
    UPDATE ҩƷ�շ���¼
    SET ��¼״̬ = 3
    WHERE NO = NO_IN AND ���� = 18    AND ��¼״̬ = 1;

    IF SQL%ROWCOUNT = 0 THEN
        mErrMsg:='[ZLSOFT]�õ����Ѿ������˳�����[ZLSOFT]';
        RAISE mErrItem; 
    END IF;
    
    --��Ҫ���ԭ���������ڷ����Ĳ��ϣ����ܶ������
    SELECT COUNT(*) INTO v_BatchCount 
    FROM  ҩƷ�շ���¼ a,�������� b
    WHERE a.ҩƷid=b.����id
        AND a.no=NO_IN  AND a.����=18 AND a.��¼״̬=3 AND nvl(a.����,0)=0
        AND ((NVL(B.�ⷿ����,0)=1 AND  
        A.�ⷿID NOT IN (SELECT ����ID FROM  ��������˵�� WHERE (�������� LIKE '���ϲ���') OR (�������� LIKE '�Ƽ���')))
        OR NVL(B.���÷���,0)=1);
    IF v_batchcount>0 THEN
        mErrMsg:='[ZLSOFT]�õ����а�����һ��ԭ�������������ڷ����Ĳ��ϣ����ܳ�����[ZLSOFT]';
        RAISE mErrItem; 
    END IF;  

    Insert INTO ҩƷ�շ���¼
        (ID,��¼״̬,����,NO,���,�ⷿID,������ID,���ϵ��,ҩƷID,����,
        ����,����,Ч��,���Ч��,�������,���۽��,���,ժҪ,������,��������,�����,�������)
        SELECT ҩƷ�շ���¼_ID.Nextval,2,����,NO_IN,���,�ⷿID,
            ������ID,���ϵ��,a.ҩƷID,
            DECODE(NVL(a.����,0),0,NULL,(DECODE(NVL(b.�ⷿ����,0),0,NULL,a.����))),
            a.����,a.����,a.Ч��,a.���Ч��,a.�������,a.���۽��,-a.���,a.ժҪ,
            �����_IN,SYSDATE,�����_IN,SYSDATE
        FROM ҩƷ�շ���¼ a,�������� b
        WHERE NO = NO_IN
            AND a.ҩƷid=b.����id
            AND ���� = 18
            AND ��¼״̬ = 3;

    FOR V_ҩƷ�շ���¼ IN C_ҩƷ�շ���¼ LOOP
        --ԭ�����ֲ������Ĳ���,��C����ʱ��Ҫ������
        BEGIN 
            SELECT COUNT(*) INTO V_COUNT
            FROM ҩƷ�շ���¼ A,�������� B
            WHERE a.ҩƷID =b.����id AND a.ҩƷID+0=V_ҩƷ�շ���¼.����ID
                AND A.NO=NO_IN AND A.���� = 18  and a.�ⷿid+0=V_ҩƷ�շ���¼.�ⷿid
                AND A.��¼״̬ = 3 AND NVL(A.����,0)>0
                AND (NVL(B.�ⷿ����,0)=0 OR (NVL(B.���÷���,0)=0 AND
                A.�ⷿID IN (SELECT ����ID FROM  ��������˵�� WHERE (�������� LIKE '���ϲ���') OR (�������� LIKE '�Ƽ���'))));
        EXCEPTION 
            WHEN OTHERS THEN
                V_COUNT:=0;
        END;
        IF V_COUNT>0 THEN
            V_����:=0;
        ELSE
            V_����:=NVL (V_ҩƷ�շ���¼.����,0);
        END IF;

        --����ҩƷ�������Ӧ����

        UPDATE ҩƷ���
        SET ʵ�ʲ�� = NVL (ʵ�ʲ��,0) + NVL (V_ҩƷ�շ���¼.���,0)
        WHERE �ⷿID = V_ҩƷ�շ���¼.�ⷿID
            AND ҩƷID = V_ҩƷ�շ���¼.����ID
            AND NVL (����,0) = V_����
            AND ���� = 1;

        IF SQL%NOTFOUND THEN
            Insert INTO ҩƷ���
                (�ⷿID,ҩƷID,����,����,ʵ�ʲ��,�ϴ�����,Ч��,���Ч��)
            VALUES (
            V_ҩƷ�շ���¼.�ⷿID,
            V_ҩƷ�շ���¼.����ID,
            V_����,
            18,
            V_ҩƷ�շ���¼.���,
            V_ҩƷ�շ���¼.����,
            V_ҩƷ�շ���¼.Ч��,
            v_ҩƷ�շ���¼.���Ч��
            );
        END IF;

        DELETE
        FROM ҩƷ���
        WHERE �ⷿID = V_ҩƷ�շ���¼.�ⷿID
            AND ҩƷID = V_ҩƷ�շ���¼.����ID
            AND nvl(��������,0) = 0 AND nvl(ʵ������,0) = 0 AND nvl(ʵ�ʽ��,0) = 0 AND nvl(ʵ�ʲ��,0) = 0;

        --����ҩƷ�շ����ܱ����Ӧ����

        UPDATE ҩƷ�շ�����
        SET ��� = NVL (���,0) + NVL (V_ҩƷ�շ���¼.���,0)
        WHERE ���� = TRUNC (SYSDATE)
            AND �ⷿID = V_ҩƷ�շ���¼.�ⷿID
            AND ҩƷID = V_ҩƷ�շ���¼.����ID
            AND ���ID = V_ҩƷ�շ���¼.������ID
            AND ���� = 18;

        IF SQL%NOTFOUND THEN
            Insert INTO ҩƷ�շ�����
                (����,�ⷿID,ҩƷID,���ID,����,���)
            VALUES (
                TRUNC (SYSDATE),
                V_ҩƷ�շ���¼.�ⷿID,
                V_ҩƷ�շ���¼.����ID,
                V_ҩƷ�շ���¼.������ID,
                18,
                V_ҩƷ�շ���¼.���);
        END IF;
    END LOOP;
EXCEPTION
    WHEN mErrItem THEN
        Raise_application_error (-20101,mErrMsg);
    WHEN OTHERS THEN 
        zl_ErrorCenter (SQLCODE,SQLERRM);
END zl_���Ͽ���۵���_strike;
/

CREATE OR REPLACE PROCEDURE zl_�����ƿ�_Insert (
	NO_IN		IN ҩƷ�շ���¼.NO%TYPE,
	���_IN		IN ҩƷ�շ���¼.���%TYPE,
	�ⷿID_IN	IN ҩƷ�շ���¼.�ⷿID%TYPE,
	�Է�����ID_IN   IN ҩƷ�շ���¼.�Է�����ID%TYPE,
	����ID_IN	IN ҩƷ�շ���¼.ҩƷID%TYPE,
	����_IN		IN ҩƷ�շ���¼.����%TYPE,
	��д����_IN	IN ҩƷ�շ���¼.��д����%TYPE,
	ʵ������_IN	IN ҩƷ�շ���¼.ʵ������%TYPE,
	�ɱ���_IN	IN ҩƷ�շ���¼.�ɱ���%TYPE,
	�ɱ����_IN	IN ҩƷ�շ���¼.�ɱ����%TYPE,
	���ۼ�_IN	IN ҩƷ�շ���¼.���ۼ�%TYPE,
	���۽��_IN	IN ҩƷ�շ���¼.���۽��%TYPE,
	���_IN		IN ҩƷ�շ���¼.���%TYPE,
	������_IN	IN ҩƷ�շ���¼.������%TYPE,
	����_IN		IN ҩƷ�շ���¼.����%TYPE := NULL,
	����_IN		IN ҩƷ�շ���¼.����%TYPE := NULL,
	Ч��_IN		IN ҩƷ�շ���¼.Ч��%TYPE := NULL,
	���Ч��_IN	IN ҩƷ�շ���¼.���Ч��%TYPE := NULL,
	ժҪ_IN		IN ҩƷ�շ���¼.ժҪ%TYPE := NULL,
	��������_IN	IN ҩƷ�շ���¼.��������%TYPE := NULL
	)
IS
	mErrItem	EXCEPTION;
	mErrMsg		varchar2(100);

	v_�¿��	ϵͳ������.����ֵ%type;
	V_����		�շ���ĿĿ¼.����%TYPE;
	V_��������	ҩƷ���.��������%TYPE;
	V_lngID		ҩƷ�շ���¼.ID%TYPE;--�շ�ID
	V_������ID    ҩƷ�շ���¼.������ID%TYPE;--������ID
	V_�������ID    ҩƷ�շ���¼.������ID%TYPE;--������ID
	V_����		ҩƷ�շ���¼.����%TYPE := NULL;--��Ҫ��������ʵ�з�������Ĳ���
	V_�Ƿ����	INTEGER;--�ж�����Ƿ��������   1:������0��������
	V_�ⷿ����	INTEGER;--�ж�����Ƿ��������   1:������0��������
	V_���÷���	INTEGER;--�ж�����Ƿ��������   1:������0��������
	intRecords	NUMBER ;

BEGIN
	BEGIN
		SELECT nvl(����ֵ,'0') INTO v_�¿�� FROM ϵͳ������ WHERE ������=95;
	EXCEPTION 
		WHEN OTHERS THEN v_�¿��:='-99';
	END;

	IF v_�¿��='-99' THEN 
		mErrMsg:='[ZLSOFT]��ϵͳ��������"���������¿��ÿ��"����,����ϵͳ��Ա��ϵ![ZLSOFT]';
		RAISE mErrItem;
	END IF ;

	IF ����_IN > 0 THEN
		BEGIN
		    SELECT ��������    INTO V_��������
		    FROM ҩƷ���
		    WHERE ҩƷID = ����ID_IN
			AND NVL (����,0) = ����_IN
			AND �ⷿID = �ⷿID_IN
			AND ���� = 1
			AND ROWNUM = 1;
		EXCEPTION
		    WHEN OTHERS THEN
			V_�������� := 0;
		END;

		IF V_�������� - ʵ������_IN < 0 THEN
		    SELECT ���� INTO V_����    FROM �շ���ĿĿ¼ WHERE ID = ����ID_IN;
		    mErrMsg:='[ZLSOFT]����Ϊ' || V_���� || ',����Ϊ' || ����_IN ||'�ķ����������' || CHR (10) || CHR (13) || '���ÿ������������[ZLSOFT]';
		    RAISE mErrItem;
		END IF;
	END IF;

	--�����ҳ���ͳ������ID
	SELECT B.ID INTO V_������ID 
	FROM ҩƷ�������� A,ҩƷ������ B 
	WHERE A.���ID = B.ID AND A.���� = 34 AND B.ϵ�� = 1 AND ROWNUM < 2;

	SELECT B.ID INTO V_�������ID
	FROM ҩƷ�������� A,ҩƷ������ B
	WHERE A.���ID = B.ID AND A.���� = 34 AND B.ϵ�� = -1 AND ROWNUM < 2;

	--�������Ϊ������һ��
	Insert INTO ҩƷ�շ���¼
		(ID,��¼״̬,����,NO,���,�ⷿID,�Է�����ID,������ID,���ϵ��,ҩƷID,����,����,����,Ч��,���Ч��,
		��д����,ʵ������,�ɱ���,�ɱ����,���ۼ�,���۽��,���,ժҪ,������,��������)
	VALUES (ҩƷ�շ���¼_ID.Nextval,1,19,NO_IN,���_IN,�ⷿID_IN,�Է�����ID_IN,
		V_�������ID,-1,����ID_IN,����_IN,����_IN,����_IN,Ч��_IN,���Ч��_IN,
		��д����_IN,ʵ������_IN,�ɱ���_IN,�ɱ����_IN,���ۼ�_IN,���۽��_IN,
		���_IN,ժҪ_IN,������_IN,��������_IN);


	IF to_number(v_�¿��,'9999')=1 THEN 
		

		UPDATE ҩƷ���
		SET �������� = NVL(��������, 0) - ʵ������_IN
		WHERE �ⷿID = �ⷿID_IN AND ҩƷID = ����ID_IN 
			AND NVL(����, 0) = NVL(����_IN, 0) AND ���� = 1;

		IF SQL%ROWCOUNT = 0 THEN
			INSERT INTO ҩƷ���(�ⷿID, ҩƷID, ����, ����, ��������)
			VALUES(�ⷿID_IN, ����ID_IN, NVL(����_IN, 0), 1, -ʵ������_IN);
		END IF;
			
		--ͬʱ���¿����
		DELETE
		FROM ҩƷ���
		WHERE �ⷿID = �ⷿID_IN
			AND ҩƷID = ����ID_IN
			AND nvl(��������,0) = 0
			AND nvl(ʵ������,0) = 0
			AND nvl(ʵ�ʽ��,0) = 0
			AND nvl(ʵ�ʲ��,0) = 0;
	END IF ;

	--�������ж����Ĳ����Ƿ��Ƿ����������
	SELECT NVL (�ⷿ����,0),nvl(���÷���,0) INTO V_�ⷿ����,v_���÷���
	FROM ��������
	WHERE ����ID = ����ID_IN;

	V_�Ƿ���� := 0;

	IF v_���÷���=0 then
		IF V_�ⷿ���� = 1 THEN
		    BEGIN
			SELECT DISTINCT 0 INTO V_�Ƿ����
			FROM ��������˵��
			WHERE ( (�������� LIKE '���ϲ���')    OR (�������� LIKE '�Ƽ���')) AND ����ID = �Է�����ID_IN;
		    EXCEPTION
			WHEN OTHERS THEN
			V_�Ƿ���� := 1;
		    END;
		END IF;
	ELSE 
		V_�Ƿ���� := 1;
	END if;

	SELECT ҩƷ�շ���¼_ID.Nextval INTO V_lngID FROM Dual;

	IF V_�Ƿ���� = 1 AND NVL (����_IN,0) = 0 THEN--�������ҳ��ⲻ����
		V_���� := V_lngID;
	ELSIF    V_�Ƿ���� = 0 THEN--��ⲻ����
		V_���� := 0;
	ELSIF NVL (����_IN,0) <> 0 THEN--�������ҳ���Ҳ����
		V_���� := ����_IN;
	END IF;

	--�������Ϊ�����һ��
	Insert INTO ҩƷ�շ���¼
		(ID,��¼״̬,����,NO,���,�ⷿID,�Է�����ID,������ID,���ϵ��,ҩƷID,����,����,����,Ч��,���Ч��,
		��д����,ʵ������,�ɱ���,�ɱ����,���ۼ�,���۽��,���,ժҪ,������,��������)
	VALUES (V_lngID,1,19,NO_IN,���_IN + 1,�Է�����ID_IN,�ⷿID_IN,V_������ID,
		1,����ID_IN,V_����,����_IN,����_IN,Ч��_IN,���Ч��_IN,��д����_IN,ʵ������_IN,
		�ɱ���_IN,�ɱ����_IN,���ۼ�_IN,���۽��_IN,���_IN,ժҪ_IN,
		������_IN,��������_IN
		);

	--����Ƿ������ͬ������ͬ���ε����ݣ�������ڲ�������
	SELECT COUNT(*) INTO intRecords 
	FROM ҩƷ�շ���¼
	WHERE ����=19 AND NO=NO_IN AND ���ϵ��=-1 AND ҩƷID+0=����ID_IN AND Nvl(����,0)=NVL(����_IN,0);

	IF intRecords>1 THEN
		SELECT ���� INTO V_���� FROM �շ���ĿĿ¼ WHERE ID = ����ID_IN;
		mErrMsg:='[ZLSOFT]����Ϊ'||V_����||'��ҩƷ�����ڶ����ظ��ļ�¼����ϲ�Ϊһ����¼��[ZLSOFT]';
		RAISE mErrItem;
	END IF ;

EXCEPTION
    WHEN mErrItem THEN
        Raise_application_error (-20101,mErrMsg  );
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE,SQLERRM);
END zl_�����ƿ�_Insert;
/


-----------------------------------------------------------
-- �����ƿ�ĳ�������
--˵�������ȸ�ԭ���ݵļ�¼״̬Ϊ3;
--������һ�ŵ��ݺ���ͬ����¼״̬Ϊ2�������ͽ��Ϊ���ĳ�������;
-- �Կ���Ĵ���Ҫ�ֿ��������������мӣ������ҩ���м���
--������ҩƷ�շ����ܱ����Ӧ�����ͽ�
--�������
--���룺NO_IN,�����_IN
------------------------------------------------------------

CREATE OR REPLACE PROCEDURE ZL_�����ƿ�_STRIKE (
	�д�_IN			IN INTEGER,
	ԭ��¼״̬_IN		IN ҩƷ�շ���¼.��¼״̬%TYPE,
	NO_IN			IN ҩƷ�շ���¼.NO%TYPE,
	���_IN			IN ҩƷ�շ���¼.���%TYPE,
	����id_IN		IN ҩƷ�շ���¼.ҩƷID%TYPE,
	��������_IN		IN ҩƷ�շ���¼.ʵ������%TYPE,
	������_IN		IN ҩƷ�շ���¼.������%TYPE,
	��������_IN		IN ҩƷ�շ���¼.��������%TYPE
)
IS
	mErrMsg		varchar2(500);
	mErrItem	EXCEPTION;

	V_BATCHCOUNT INTEGER;    --ԭ���������ڷ�����ҩƷ������
	V_���		ҩƷ�շ���¼.���%TYPE;
	V_�ⷿID	ҩƷ�շ���¼.�ⷿID%TYPE;
	V_�Է�����ID	ҩƷ�շ���¼.�Է�����ID%TYPE;
	V_������ID	ҩƷ�շ���¼.������ID%TYPE ;
	V_����		ҩƷ�շ���¼.����%TYPE ;
	V_����		ҩƷ�շ���¼.����%TYPE ;
	V_����		ҩƷ�շ���¼.����%TYPE ;
	V_Ч��		ҩƷ�շ���¼.Ч��%TYPE ;
	V_�ɱ���	ҩƷ�շ���¼.�ɱ���%TYPE ;
	V_�ɱ����	ҩƷ�շ���¼.�ɱ����%TYPE ;
	V_����		ҩƷ�շ���¼.����%TYPE ;
	V_���ۼ�	ҩƷ�շ���¼.���ۼ�%TYPE ;
	V_���۽��	ҩƷ�շ���¼.���۽��%TYPE ;
	V_���		ҩƷ�շ���¼.���%TYPE ;
	V_ժҪ		ҩƷ�շ���¼.ժҪ%TYPE ;
	V_ʣ������	ҩƷ�շ���¼.ʵ������%TYPE; 
	V_ʣ��ɱ����	ҩƷ�շ���¼.�ɱ����%Type;
	V_ʣ�����۽��	ҩƷ�շ���¼.���۽��%Type;
	V_���ϵ��	ҩƷ�շ���¼.���ϵ��%TYPE;
	V_�������	ҩƷ�շ���¼.�������%TYPE;
	V_���Ч��	ҩƷ�շ���¼.���Ч��%TYPE;
	V_��¼��	NUMBER;
	V_�շ�ID	ҩƷ�շ���¼.ID%TYPE;
	V_��ҩ��	ҩƷ�շ���¼.��ҩ��%TYPE;
	V_��������	ҩƷ�շ���¼.��ҩ����%TYPE;
 
	--�Գ����������м��
	V_�����	ҩƷ���.ʵ������%TYPE;
	V_�ⷿ����	INTEGER;
	V_���÷���	INTEGER;
	V_��������	INTEGER;
	V_���Ŀ�		INTEGER;
	V_����		NUMBER;
	INTDIGIT	NUMBER;
     
BEGIN
         --��ȡ���С��λ��
    SELECT NVL(����ֵ,ȱʡֵ) INTO INTDIGIT FROM ϵͳ������ WHERE ������='���ý���λ��';
 

    IF �д�_IN =1 THEN
	UPDATE ҩƷ�շ���¼
	SET ��¼״̬ = DECODE(ԭ��¼״̬_IN,1,3,ԭ��¼״̬_IN+3)
	WHERE NO = NO_IN AND ���� = 19 AND ��¼״̬ =ԭ��¼״̬_IN ;

	IF SQL%ROWCOUNT = 0 THEN
		mErrMsg:='[ZLSOFT]�õ����Ѿ������˳�����[ZLSOFT]';
		RAISE mErrItem;
	END IF;
    END IF;
 

    SELECT COUNT(*) INTO V_BATCHCOUNT FROM
            ҩƷ�շ���¼ A,�������� B
    WHERE A.ҩƷID=B.����ID
        AND A.NO=NO_IN
        AND A.����=19
        AND A.ҩƷID+0=����id_IN
        AND MOD(A.��¼״̬,3)=0
        AND NVL(A.����,0)=0
        AND ((NVL(B.�ⷿ����,0)=1 AND A.�ⷿID NOT IN (SELECT ����ID FROM  ��������˵�� WHERE (�������� LIKE '%���ϲ���') OR (�������� LIKE '�Ƽ���')))
            OR NVL(B.���÷���,0)=1);
 

    IF V_BATCHCOUNT>0 THEN
	mErrMsg:='[ZLSOFT]�õ����а�����һ��ԭ�������������ڷ������������ϣ����ܳ�����[ZLSOFT]';
        RAISE mErrItem;
    END IF;
 
    SELECT SUM(A.ʵ������) AS ʣ������,SUM(A.�ɱ����) AS ʣ��ɱ����,SUM(A.���۽��) AS ʣ�����۽��,A.�ɱ���,A.���ۼ�,A.�Է�����ID,NVL(A.����,0),B.�ⷿ����,B.���÷���
    INTO  V_ʣ������,v_ʣ��ɱ����,V_ʣ�����۽��,V_�ɱ���,V_���ۼ�,V_�ⷿID,V_����,V_�ⷿ����,V_���÷���
    FROM ҩƷ�շ���¼ A,�������� B
    WHERE A.NO=NO_IN AND A.ҩƷID=B.����ID AND A.����=19 AND A.ҩƷID+0=����id_IN AND A.���=���_IN
    GROUP BY A.�ɱ���,A.���ۼ�,A.�Է�����ID,NVL(A.����,0),B.�ⷿ����,B.���÷���;
 
    --�жϸò����ǿⷿ���Ƿ��ϲ���
    BEGIN
        SELECT DISTINCT 0   INTO V_���Ŀ�
        FROM ��������˵��
        WHERE ((�������� LIKE '%���Ŀ�')
		OR (�������� LIKE '�Ƽ���'))
		AND ����ID = V_�ⷿID;
    EXCEPTION
        WHEN OTHERS THEN V_���Ŀ�:=1;
    END ;
 
    --���ݲ�������,�жϷ�������
    IF V_���Ŀ�=0 THEN
        V_��������:=V_���÷���;
    ELSE
        V_��������:=V_�ⷿ����;
    END IF ;
 
    --V_����:(ԭ����,�ַ���Ϊ����;����Ϊ��)ԭ������,�ַ���,�����̲�����
    --��Ϊ���ڶԷ��ⷿ���൱�ڳ��⣬�����ڵ�ǰ�ⷿ���൱����⣬���Ե�ǰ�ⷿ�����飬������˿�������¼
    SELECT NVL(A.����,0) INTO V_����
    FROM ҩƷ�շ���¼ A
    WHERE A.NO=NO_IN AND A.����=19 AND A.ҩƷID+0=����id_IN AND A.���=���_IN+1 AND MOD(A.��¼״̬,3)=0;
 
    --ȡ�����
    BEGIN
        SELECT NVL(ʵ������,0) INTO V_����� FROM ҩƷ���
        WHERE �ⷿID=V_�ⷿID AND ҩƷID=����id_IN AND NVL(����,0)=V_���� AND ����=1;
    EXCEPTION
        WHEN OTHERS THEN V_�����:=0;
    END ;
    
    IF nvl(V_ʣ������,0)=0 THEN 
            mErrMsg:='[ZLSOFT]�õ����е�' || ceil(���_IN/2) || '�еĲ����Ѿ����������,�����ٳ壡[ZLSOFT]';
            RAISE mErrItem; 
    END IF ;

    --������������ʣ������,ȡʣ������;����ȡ�����
    IF V_�����<V_ʣ������ THEN
            v_ʣ��ɱ����:=V_�����/V_ʣ������*v_ʣ��ɱ����;
            V_ʣ�����۽��:=V_�����/V_ʣ������*V_ʣ�����۽��;
            V_ʣ������:=V_�����;
    END IF ;
 
    --������������ʣ��������������
    IF V_ʣ������<��������_IN THEN
	mErrMsg:='[ZLSOFT]�õ����е�' || ceil(���_IN/2) || '�е��������ϳ���������������ʣ������ݣ����ܳ�����[ZLSOFT]';
        RAISE mErrItem;
    END IF;
 
    V_�ɱ����:= ROUND(��������_IN/v_ʣ������*v_ʣ��ɱ����,INTDIGIT);
    V_���۽��:= ROUND(��������_IN/v_ʣ������*V_ʣ�����۽��,INTDIGIT);
    V_���:=round(V_���۽��-V_�ɱ����,INTDIGIT);





    FOR v_���� IN (SELECT ���,�ⷿID,�Է�����ID,������ID,���ϵ��,ҩƷID,����,����,����,Ч��,�������,���Ч��,��ҩ��,��ҩ����,ժҪ
			   FROM ҩƷ�շ���¼
			   WHERE NO = NO_IN AND ���� = 19 AND (���>=���_IN AND ���<=���_IN+1) AND (��¼״̬=1 OR MOD(��¼״̬,3)=0)
		           ORDER BY ҩƷID)
    LOOP
        V_���:=v_����.���;
        V_�ⷿID:=v_����.�ⷿID;
        V_�Է�����ID:=v_����.�Է�����ID;
        V_������ID:=v_����.������ID;
        V_���ϵ��:=v_����.���ϵ��;
        V_����:=v_����.����;
        V_����:=v_����.����;
        V_����:=v_����.����;
        V_Ч��:=v_����.Ч��;
        v_ժҪ:=v_����.ժҪ;
        V_��ҩ��:=v_����.��ҩ��;
        V_��������:=v_����.��ҩ����;
	v_�������:=v_����.�������;
	v_���Ч��:=v_����.���Ч��;

        SELECT ҩƷ�շ���¼_ID.NEXTVAL INTO V_�շ�ID  FROM DUAL;

        INSERT INTO ҩƷ�շ���¼
        (	ID,��¼״̬,����,NO,���,�ⷿID,�Է�����ID,������ID,���ϵ��,
		ҩƷID,����,����,����,Ч��,�������,���Ч��,��д����,ʵ������,�ɱ���,
		�ɱ����,���ۼ�,���۽��,���,ժҪ,������,��������,�����,�������,��ҩ��,��ҩ����)
        VALUES
        (V_�շ�ID,DECODE(ԭ��¼״̬_IN,1,2,ԭ��¼״̬_IN+2),19,NO_IN,V_���,V_�ⷿID,V_�Է�����ID,V_������ID,V_���ϵ��,
        ����id_IN,V_����,V_����,V_����,V_Ч��,v_�������,v_���Ч��,-��������_IN,-��������_IN,V_�ɱ���,-V_�ɱ����,V_���ۼ�,-V_���۽��,
        -V_���,V_ժҪ,������_IN,��������_IN,������_IN,��������_IN,V_��ҩ��,V_��������);
 
        --ԭ�����ֲ������Ĳ���,�ڳ���ʱ��Ҫ������
        BEGIN
            SELECT COUNT(*) INTO V_��¼��
            FROM ҩƷ�շ���¼ A, �������� B
            WHERE B.����ID=A.ҩƷID
            AND A.ҩƷID=����id_IN
            AND A.NO=NO_IN
            AND A.���� = 19
            AND A.�ⷿID=V_�ⷿID
            AND MOD(A.��¼״̬,3)=0
            AND NVL(A.����,0)>0
            AND (NVL(B.�ⷿ����,0)=0 OR
                (NVL(B.���÷���,0)=0 AND A.�ⷿID IN (SELECT ����ID FROM  ��������˵�� WHERE (�������� LIKE '%���ϲ���') OR (�������� LIKE '�Ƽ���'))))
            ;
        EXCEPTION
            WHEN OTHERS THEN V_��¼��:=0;
        END;
        IF V_��¼��>0 THEN
            V_����:=0;
        ELSE
            V_����:=NVL (V_����, 0);
        END IF;
 
        --����ҩƷ�������Ӧ����
        UPDATE ҩƷ���
            SET ��������=NVL(��������,0)-NVL(��������_IN,0)*V_���ϵ��,
                ʵ������=NVL(ʵ������,0)-NVL(��������_IN,0)*V_���ϵ��,
                ʵ�ʽ��=NVL(ʵ�ʽ��,0)-NVL(V_���۽��,0)*V_���ϵ��,
                ʵ�ʲ��=NVL(ʵ�ʲ��,0)-NVL(V_���,0)*V_���ϵ��,
                �ϴβɹ���=NVL(V_�ɱ���,�ϴβɹ���),
                �ϴ�����=NVL(V_����,�ϴ�����),
                �ϴβ���=NVL(V_����,�ϴβ���),
                Ч��=NVL(V_Ч��,Ч��)
          WHERE �ⷿID = V_�ⷿID
            AND ҩƷID = ����id_IN
            AND NVL (����, 0) = V_����
            AND ���� = 1;
 
        IF SQL%NOTFOUND THEN
            INSERT INTO ҩƷ���
            (�ⷿID,ҩƷID,����,����,��������,ʵ������,ʵ�ʽ��, ʵ�ʲ��,�ϴβɹ���,�ϴ�����,�ϴβ���,Ч��,���Ч��)
            VALUES
            (V_�ⷿID,����id_IN,V_����,1,-��������_IN*V_���ϵ��,-��������_IN*V_���ϵ��,
            -V_���۽��*V_���ϵ��,-V_���*V_���ϵ��,V_�ɱ���,V_����,V_����,V_Ч��,v_���Ч��);
        END IF;
 
        DELETE
          FROM ҩƷ���
         WHERE �ⷿID = V_�ⷿID
           AND ҩƷID = ����id_IN
           AND NVL(��������,0)=0
           AND NVL(ʵ������,0)=0
           AND NVL(ʵ�ʽ��,0)=0
           AND NVL(ʵ�ʲ��,0)=0;
 
        --����ҩƷ�շ����ܱ����Ӧ����
        UPDATE ҩƷ�շ�����
         SET ���� =    NVL (����,0)  - NVL (��������_IN,0)*V_���ϵ��,
             ��� = NVL (���, 0) - NVL (V_���۽��, 0)*V_���ϵ��,
             ��� = NVL (���, 0) - NVL (V_���, 0)*V_���ϵ��
        WHERE ���� = TRUNC (��������_IN)
         AND �ⷿID = V_�ⷿID
         AND ҩƷID = ����id_IN
         AND ���ID = V_������ID
         AND ���� = 19;
        IF SQL%NOTFOUND THEN
            INSERT INTO ҩƷ�շ�����
            (����, �ⷿID, ҩƷID, ���ID, ����, ����, ���, ���)
            VALUES
            (TRUNC (��������_IN),V_�ⷿID,����id_IN,V_������ID,
            19,-��������_IN*V_���ϵ��,-V_���۽��*V_���ϵ��,-V_���*V_���ϵ��);
        END IF;
    END LOOP;
EXCEPTION
    WHEN mErrItem THEN
        RAISE_APPLICATION_ERROR (-20101, mErrMsg);
    WHEN OTHERS THEN ZL_ERRORCENTER (SQLCODE, SQLERRM);
END ZL_�����ƿ�_STRIKE;
/

-----------------------------------------------------------
-- �������õĳ�������
--˵�������ȸ�ԭ���ݵļ�¼״̬Ϊ3;
--������һ�ŵ��ݺ���ͬ����¼״̬Ϊ2�������ͽ��Ϊ���ĳ�������;
--������ҩƷ�շ����ܱ����Ӧ�����ͽ�
--�������
--���룺NO_IN,�����_IN
------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_��������_STRIKE (
    �д�_IN        IN INTEGER,
    ԭ��¼״̬_IN    IN ҩƷ�շ���¼.��¼״̬%TYPE,
    NO_IN        IN ҩƷ�շ���¼.NO%TYPE,
    ���_IN        IN ҩƷ�շ���¼.���%TYPE,
    ����ID_IN    IN ҩƷ�շ���¼.ҩƷID%TYPE,
    ��������_IN    IN ҩƷ�շ���¼.ʵ������%TYPE,
    ������_IN    IN ҩƷ�շ���¼.������%TYPE,
    ��������_IN    IN ҩƷ�շ���¼.��������%TYPE
)
IS
	mErrMsg         varchar2(100);
	mErrItem        EXCEPTION;
	v_BatchCount    INTEGER;    --ԭ���������ڷ����Ĳ��ϵ�����

	V_�ⷿID        ҩƷ�շ���¼.�ⷿID%TYPE; 
	V_�Է�����ID    ҩƷ�շ���¼.�Է�����ID%TYPE;
	V_������ID    ҩƷ�շ���¼.������ID%TYPE ;
	V_����          ҩƷ�շ���¼.����%TYPE ; 
	V_����          ҩƷ�շ���¼.����%TYPE ; 
	V_����          ҩƷ�շ���¼.����%TYPE ; 
	V_Ч��          ҩƷ�շ���¼.Ч��%TYPE ; 
	V_�ɱ���        ҩƷ�շ���¼.�ɱ���%TYPE ; 
	V_�ɱ����      ҩƷ�շ���¼.�ɱ����%TYPE ; 
	V_����          ҩƷ�շ���¼.����%TYPE ; 
	V_���ۼ�        ҩƷ�շ���¼.���ۼ�%TYPE ; 
	V_���۽��      ҩƷ�շ���¼.���۽��%TYPE ; 
	V_���          ҩƷ�շ���¼.���%TYPE ; 
	V_ժҪ          ҩƷ�շ���¼.ժҪ%TYPE ; 
	V_ʣ������	ҩƷ�շ���¼.ʵ������%TYPE; 
	V_ʣ��ɱ����	ҩƷ�շ���¼.�ɱ����%Type;
	V_ʣ�����۽��	ҩƷ�շ���¼.���۽��%Type;
	V_���ϵ��      ҩƷ�շ���¼.���ϵ��%TYPE; 
	V_�շ�ID        ҩƷ�շ���¼.ID%TYPE; 
	v_�������        ҩƷ�շ���¼.�������%TYPE; 
	v_���Ч��        ҩƷ�շ���¼.���Ч��%TYPE; 
	V_��¼��            NUMBER; 
	V_С��		number(2);
BEGIN
    BEGIN 
	    SELECT  to_number(Nvl(����ֵ,ȱʡֵ),'99999')  INTO v_С�� 
	    From ϵͳ������ where ������='���ý���λ��';
    EXCEPTION 
	WHEN OTHERS THEN 
		v_С��:=2;		
    END;

    IF �д�_IN =1 THEN
        UPDATE ҩƷ�շ���¼ 
        SET ��¼״̬ = DECODE(ԭ��¼״̬_IN,1,3,ԭ��¼״̬_IN+3) 
        WHERE NO = NO_IN AND ���� = 20 AND ��¼״̬ =ԭ��¼״̬_IN ; 

        IF SQL%ROWCOUNT = 0 THEN 
            mErrMsg:='[ZLSOFT]�õ����Ѿ������˳�����[ZLSOFT]';
            RAISE mErrItem;
        END IF; 
    END IF;

    --��Ҫ���ԭ���������ڷ����Ĳ��ϣ����ܶ������
    SELECT COUNT(*) INTO V_BATCHCOUNT 
    FROM  ҩƷ�շ���¼ A,�������� B
    WHERE A.ҩƷID=B.����ID
        AND A.NO=NO_IN     AND A.����=20 AND Mod(A.��¼״̬,3)=0 AND NVL(A.����,0)=0
        AND ((NVL(B.�ⷿ����,0)=1 AND A.�ⷿID NOT IN (SELECT ����ID FROM  ��������˵�� WHERE (�������� LIKE '���ϲ���') OR (�������� LIKE '�Ƽ���')))
        OR NVL(B.���÷���,0)=1);
        
    IF V_BATCHCOUNT>0 THEN
        mErrMsg:='[ZLSOFT]�õ����а�����һ��ԭ�������������ڷ����Ĳ��ϣ����ܳ�����[ZLSOFT]';
        RAISE mErrItem;
    END IF;  

    SELECT SUM(ʵ������) AS ʣ������,SUM(�ɱ����) AS ʣ��ɱ����,SUM(���۽��) AS ʣ�����۽��,�ⷿID,�Է�����ID,������ID,���ϵ��,����,����,����,Ч��,�������,���Ч��,�ɱ���,����,���ۼ�,ժҪ
        INTO  V_ʣ������,v_ʣ��ɱ����,V_ʣ�����۽��,V_�ⷿID,V_�Է�����ID,V_������ID,V_���ϵ��,V_����,V_����,V_����,V_Ч��,v_�������,
        v_���Ч��,V_�ɱ���,V_����,V_���ۼ�,V_ժҪ
    FROM ҩƷ�շ���¼ 
    WHERE NO=NO_IN 
        AND ����=20
        AND ҩƷID+0=����ID_IN 
        AND ���=���_IN
    GROUP BY �ⷿID,�Է�����ID,������ID,���ϵ��,����,����,����,Ч��,�������,���Ч��,�ɱ���,����,���ۼ�,ժҪ;

    IF nvl(V_ʣ������,0)=0 THEN 
            mErrMsg:='[ZLSOFT]�õ����е�' || ���_IN || '�еĲ����Ѿ����������,�����ٳ壡[ZLSOFT]';
            RAISE mErrItem; 
    END IF ;

    --������������ʣ��������������
    IF V_ʣ������<��������_IN THEN
        mErrMsg:='[ZLSOFT]ʣ�����ݲ����ڱ���������,���ܳ�����[ZLSOFT]';
        RAISE mErrItem;
    END IF;

    V_�ɱ����:= ROUND(��������_IN/v_ʣ������*v_ʣ��ɱ����,v_С��);
    V_���۽��:= ROUND(��������_IN/v_ʣ������*V_ʣ�����۽��,v_С��);
    V_���:=round(V_���۽��-V_�ɱ����,v_С��);

    SELECT ҩƷ�շ���¼_ID.NEXTVAL INTO V_�շ�ID  FROM DUAL;

    Insert INTO ҩƷ�շ���¼
        (ID,��¼״̬,����,NO,���,�ⷿID,�Է�����ID,������ID,���ϵ��,
        ҩƷID,����,����,����,Ч��,�������,���Ч��,��д����,ʵ������,�ɱ���,
        �ɱ����,���ۼ�,���۽��,���,ժҪ,������,��������,�����,�������)
    VALUES 
        (V_�շ�ID,DECODE(ԭ��¼״̬_IN,1,2,ԭ��¼״̬_IN+2),20,NO_IN,���_IN,V_�ⷿID,V_�Է�����ID,V_������ID,V_���ϵ��,
        ����ID_IN,V_����,V_����,V_����,V_Ч��,V_�������,V_���Ч��,-��������_IN,-��������_IN,V_�ɱ���,-V_�ɱ����,V_���ۼ�,-V_���۽��,
        -V_���,V_ժҪ,������_IN,��������_IN,������_IN,��������_IN);

    --ԭ�����ֲ������Ĳ���,��C����ʱ��Ҫ������
    BEGIN 
        SELECT COUNT(*) INTO V_��¼��
        FROM ҩƷ�շ���¼ A,�������� B
        WHERE A.ҩƷid=b.����id AND  B.����ID+0=����ID_IN     AND A.NO=NO_IN    AND A.���� = 20     AND Mod(A.��¼״̬,3)=0    AND NVL(A.����,0)>0
        AND (NVL(B.�ⷿ����,0)=0 OR 
            (NVL(B.���÷���,0)=0 AND A.�ⷿID IN (SELECT ����ID FROM  ��������˵�� WHERE (�������� LIKE '���ϲ���') OR (�������� LIKE '�Ƽ���'))));
    EXCEPTION 
        WHEN OTHERS THEN V_��¼��:=0;
    END;

    IF V_��¼��>0 THEN
        V_����:=0;
    ELSE
        V_����:=NVL (V_����,0);
    END IF;

    --����ҩƷ�������Ӧ����
    UPDATE ҩƷ���
    SET �������� = NVL (��������,0) + NVL (��������_IN,0),
        ʵ������ = NVL (ʵ������,0) + NVL (��������_IN,0),
        ʵ�ʽ�� = NVL (ʵ�ʽ��,0) + NVL (V_���۽��,0),
        ʵ�ʲ�� = NVL (ʵ�ʲ��,0) + NVL (V_���,0)
    WHERE �ⷿID = V_�ⷿID
        AND ҩƷID = ����ID_IN
        AND NVL (����,0) = V_����
        AND ���� = 1;

    IF SQL%NOTFOUND THEN
        INSERT INTO ҩƷ���
            (�ⷿID,ҩƷID,����,����,��������,ʵ������,
            ʵ�ʽ��,ʵ�ʲ��,�ϴ�����,Ч��,���Ч��)
        VALUES 
            (V_�ⷿID,����ID_IN,V_����,1,��������_IN,��������_IN,
            V_���۽��,V_���,V_����,V_Ч��,V_���Ч��);
    END IF;

    DELETE ҩƷ���
    WHERE �ⷿID = V_�ⷿID
        AND ҩƷID = ����ID_IN
        AND NVL (��������,0) = 0
        AND NVL (ʵ������,0) = 0
        AND NVL (ʵ�ʽ��,0) = 0
        AND NVL (ʵ�ʲ��,0) = 0;

    --����ҩƷ�շ����ܱ����Ӧ����
    UPDATE ҩƷ�շ�����
    SET ���� = NVL (����,0)  +NVL(��������_IN,0),
        ��� = NVL (���,0) +NVL(V_���۽��,0),
        ��� = NVL (���,0) +NVL(V_���,0)
    WHERE ���� = TRUNC (��������_IN)
        AND �ⷿID = V_�ⷿID
        AND ҩƷID = ����ID_IN
        AND ���ID = V_������ID
        AND ���� = 20;

    IF SQL%NOTFOUND THEN
        INSERT INTO ҩƷ�շ�����
            (����,�ⷿID,ҩƷID,���ID,����,����,���,���)
        VALUES 
            (TRUNC (��������_IN),V_�ⷿID,����ID_IN,V_������ID,20,��������_IN,V_���۽��,V_���);
    END IF;
EXCEPTION
    WHEN mErrItem THEN
        RAISE_APPLICATION_ERROR (-20101,mErrMsg);
    WHEN OTHERS THEN ZL_ERRORCENTER (SQLCODE,SQLERRM);
END ZL_��������_STRIKE;
/



-----------------------------------------------------------
-- ��������������˴���
--˵�������ȶ�ҩƷ�շ���¼���е�����˺����ʱ����д���
--���Ŷ�ҩƷ����ҩƷ�շ����ܱ��е���Ӧ�����ͽ����д���
--�������
--���룺NO_IN,�����_IN
------------------------------------------------------------

CREATE OR REPLACE PROCEDURE zl_�����������_verify (
    NO_IN        IN ҩƷ�շ���¼.NO%TYPE := NULL,
    �����_IN    IN ҩƷ�շ���¼.�����%TYPE := NULL
)
IS
    mErrMsg            varchar2(100);
    mErrItem        EXCEPTION;

    v_BatchCount integer;    --ԭ���������ڷ����Ĳ��ϵ�����
 

    CURSOR C_ҩƷ�շ���¼    IS
    SELECT ID,ʵ������,���۽��,���,�ⷿID,ҩƷID,����,�ɱ���,����,Ч��,���Ч��,�������,����,������ID,��������
    FROM ҩƷ�շ���¼
    WHERE NO = NO_IN AND ���� = 17    AND ��¼״̬ = 1
    ORDER BY ҩƷID;
BEGIN
    UPDATE ҩƷ�շ���¼
    SET ����� = �����_IN,������� = SYSDATE
    WHERE NO = NO_IN AND ���� = 17 AND ��¼״̬ = 1 AND ����� IS NULL;

    IF SQL%ROWCOUNT = 0 THEN
        mErrMsg:='[ZLSOFT]�õ����Ѿ���������ˣ�[ZLSOFT]';
        RAISE mErrItem;
    END IF;
    
    --��Ҫ���ԭ���������ڷ����Ĳ��ϣ����ܶ������
    SELECT COUNT(*) INTO v_BatchCount 
    FROM   ҩƷ�շ���¼ a,�������� b
    WHERE a.ҩƷid=b.����id AND a.no=NO_IN  AND a.����=17 AND a.��¼״̬=1 AND nvl(a.����,0)=0
        AND ((nvl(b.�ⷿ����,0)=1 AND a.�ⷿid not in (select ����id from  ��������˵�� where (�������� LIKE '���ϲ���') or (�������� LIKE '�Ƽ���')))
            or nvl(b.���÷���,0)=1);
        
    IF v_batchcount>0 THEN
        mErrMsg:='[ZLSOFT]�õ����а�����һ��ԭ�������������ڷ����Ĳ��ϣ�������ˣ�[ZLSOFT';
        RAISE mErrItem;
    END IF;  
    
    --ԭ�����ֲ������Ĳ���,�����ʱ��Ҫ������
    UPDATE ҩƷ�շ���¼ SET ����=0
    WHERE  id=(    SELECT id FROM ҩƷ�շ���¼ a,�������� b 
            WHERE b.����id=a.ҩƷID AND a.no=no_in AND a.���� = 17
                AND a.��¼״̬ = 1 AND nvl(a.����,0)>0 AND (nvl(b.�ⷿ����,0)=0 or 
                (nvl(b.���÷���,0)=0 and a.�ⷿid in (select ����id from  ��������˵�� where (�������� LIKE '���ϲ���') or 
                (�������� LIKE '�Ƽ���'))))
            );

    FOR V_ҩƷ�շ���¼ IN C_ҩƷ�շ���¼ LOOP
        --����ҩƷ�������Ӧ����

        UPDATE ҩƷ���
        SET �������� = NVL (��������,0) + NVL (V_ҩƷ�շ���¼.ʵ������,0),
            ʵ������ = NVL (ʵ������,0) + NVL (V_ҩƷ�շ���¼.ʵ������,0),
            ʵ�ʽ�� = NVL (ʵ�ʽ��,0) + NVL (V_ҩƷ�շ���¼.���۽��,0),
            ʵ�ʲ�� = NVL (ʵ�ʲ��,0) + NVL (V_ҩƷ�շ���¼.���,0),
            �ϴβɹ��� = NVL (V_ҩƷ�շ���¼.�ɱ���,�ϴβɹ���),
            �ϴ����� = NVL (V_ҩƷ�շ���¼.����,�ϴ�����),
            �ϴ��������� = NVL (V_ҩƷ�շ���¼.��������,�ϴ���������),
            �ϴβ��� = NVL (V_ҩƷ�շ���¼.����,�ϴβ���),
            Ч�� = NVL (V_ҩƷ�շ���¼.Ч��,Ч��),
            ���Ч�� = NVL (V_ҩƷ�շ���¼.���Ч��,���Ч��)
        WHERE �ⷿID = V_ҩƷ�շ���¼.�ⷿID
            AND ҩƷID = V_ҩƷ�շ���¼.ҩƷID
            AND NVL (����,0) = NVL (V_ҩƷ�շ���¼.����,0)
            AND ���� = 1;

        IF SQL%NOTFOUND THEN
            Insert INTO ҩƷ���(�ⷿID,ҩƷID,����,����,��������,ʵ������,ʵ�ʽ��,ʵ�ʲ��,�ϴβɹ���,�ϴ�����,�ϴ���������,�ϴβ���,Ч��,���Ч��)
            VALUES (
                V_ҩƷ�շ���¼.�ⷿID,
                V_ҩƷ�շ���¼.ҩƷID,
                V_ҩƷ�շ���¼.����,
                1,
                V_ҩƷ�շ���¼.ʵ������,
                V_ҩƷ�շ���¼.ʵ������,
                V_ҩƷ�շ���¼.���۽��,
                V_ҩƷ�շ���¼.���,
                V_ҩƷ�շ���¼.�ɱ���,
                V_ҩƷ�շ���¼.����,
                V_ҩƷ�շ���¼.��������,
                V_ҩƷ�շ���¼.����,
                V_ҩƷ�շ���¼.Ч��,
                V_ҩƷ�շ���¼.���Ч��
                );
        END IF;

        DELETE
        FROM ҩƷ���
        WHERE �ⷿID = V_ҩƷ�շ���¼.�ⷿID
            AND ҩƷID = V_ҩƷ�շ���¼.ҩƷID
            AND nvl(��������,0) = 0 AND nvl(ʵ������,0) = 0 AND nvl(ʵ�ʽ��,0) = 0 AND nvl(ʵ�ʲ��,0) = 0;


        --����ҩƷ�շ����ܱ����Ӧ����
        UPDATE ҩƷ�շ�����
        SET ���� = NVL (����,0) + NVL (V_ҩƷ�շ���¼.ʵ������,0),
            ��� = NVL (���,0) + NVL (V_ҩƷ�շ���¼.���۽��,0),
            ��� = NVL (���,0) + NVL (V_ҩƷ�շ���¼.���,0)
        WHERE ���� = TRUNC (SYSDATE)
            AND �ⷿID = V_ҩƷ�շ���¼.�ⷿID
            AND ҩƷID = V_ҩƷ�շ���¼.ҩƷID
            AND ���ID = V_ҩƷ�շ���¼.������ID
            AND ���� = 17;

        IF SQL%NOTFOUND THEN
            Insert INTO ҩƷ�շ�����
                (����,�ⷿID,ҩƷID,���ID,����,����,���,���)
            VALUES (
                TRUNC (SYSDATE),
                V_ҩƷ�շ���¼.�ⷿID,
                V_ҩƷ�շ���¼.ҩƷID,
                V_ҩƷ�շ���¼.������ID,
                17,
                V_ҩƷ�շ���¼.ʵ������,
                V_ҩƷ�շ���¼.���۽��,
                V_ҩƷ�շ���¼.���
                );
        END IF;
	--���¸ò��ϵĳɱ���
	UPDATE ��������
	SET �ɱ���=V_ҩƷ�շ���¼.�ɱ��� 
	WHERE ����ID=V_ҩƷ�շ���¼.ҩƷID;

    END LOOP;
EXCEPTION
    WHEN mErrItem THEN
        Raise_application_error (-20101,mErrMsg);
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE,SQLERRM);
END zl_�����������_verify;
/

-----------------------------------------------------------
-- ������������ĳ�������
--˵�������ȸ�ԭ���ݵļ�¼״̬Ϊ3;
--������һ�ŵ��ݺ���ͬ����¼״̬Ϊ2�������ͽ��Ϊ���ĳ�������;
--�����Ĳ����շ����ܱ����Ӧ�����ͽ�
--�������
--���룺NO_IN,�����_IN
------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_������������_strike (
    �д�_IN            IN INTEGER,
    ԭ��¼״̬_IN        IN ҩƷ�շ���¼.��¼״̬%TYPE,
    NO_IN            IN ҩƷ�շ���¼.NO%TYPE,
    ���_IN            IN ҩƷ�շ���¼.���%TYPE,
    ����ID_IN        IN ҩƷ�շ���¼.ҩƷID%TYPE,
    ��������_IN        IN ҩƷ�շ���¼.ʵ������%TYPE,
    ������_IN        IN ҩƷ�շ���¼.������%TYPE,
    ��������_IN        IN ҩƷ�շ���¼.��������%TYPE
)
IS
    mErrMsg            varchar2(100);
    mErrItem        EXCEPTION;
    v_BatchCount        INTEGER;    --ԭ���������ڷ����Ĳ��ϵ�����

    V_�ⷿID            ҩƷ�շ���¼.�ⷿID%TYPE; 
    V_������ID        ҩƷ�շ���¼.������ID%TYPE ;
    V_����            ҩƷ�շ���¼.����%TYPE ; 
    V_����            ҩƷ�շ���¼.����%TYPE ; 
    V_����            ҩƷ�շ���¼.����%TYPE ; 
    V_Ч��            ҩƷ�շ���¼.Ч��%TYPE ; 
    V_�ɱ���            ҩƷ�շ���¼.�ɱ���%TYPE ; 
    V_�ɱ����        ҩƷ�շ���¼.�ɱ����%TYPE ; 
    V_����            ҩƷ�շ���¼.����%TYPE ; 
    V_���ۼ�            ҩƷ�շ���¼.���ۼ�%TYPE ; 
    V_���۽��        ҩƷ�շ���¼.���۽��%TYPE ; 
    V_���            ҩƷ�շ���¼.���%TYPE ; 
    V_ժҪ            ҩƷ�շ���¼.ժҪ%TYPE ; 
    V_ʣ������	ҩƷ�շ���¼.ʵ������%TYPE; 
    V_ʣ��ɱ����	ҩƷ�շ���¼.�ɱ����%Type;
    V_ʣ�����۽��	ҩƷ�շ���¼.���۽��%Type;

    V_���ϵ��        ҩƷ�շ���¼.���ϵ��%TYPE; 

    v_���Ч��        ҩƷ�շ���¼.���Ч��%TYPE; 
    V_��¼��            NUMBER; 
    V_�շ�ID            ҩƷ�շ���¼.ID%TYPE; 
    V_С��		number(2);
BEGIN
    BEGIN 
	    SELECT  to_number(Nvl(����ֵ,ȱʡֵ),'99999')  INTO v_С�� 
	    From ϵͳ������ where ������='���ý���λ��';
    EXCEPTION 
	WHEN OTHERS THEN 
		v_С��:=2;		
    END;
    
    IF �д�_IN =1 THEN
        UPDATE ҩƷ�շ���¼ 
        SET ��¼״̬ = DECODE(ԭ��¼״̬_IN,1,3,ԭ��¼״̬_IN+3) 
        WHERE NO = NO_IN AND ���� = 21 AND ��¼״̬ =ԭ��¼״̬_IN ; 
        IF SQL%ROWCOUNT = 0 THEN 
            mErrMsg:='[ZLSOFT]�õ����Ѿ������˳�����[ZLSOFT]';
            RAISE mErrItem;
        END IF; 
    END IF;
    
    --��Ҫ���ԭ���������ڷ����Ĳ��ϣ����ܶ������
    SELECT COUNT(*) INTO v_BatchCount 
    FROM ҩƷ�շ���¼ a,�������� b
    WHERE a.ҩƷid=b.����id    AND a.no=NO_IN     AND a.����=21    AND A.ҩƷID+0=����ID_IN    AND MOD(a.��¼״̬,3)=0    AND nvl(a.����,0)=0
        AND ((NVL(B.�ⷿ����,0)=1 AND A.�ⷿID NOT IN (SELECT ����ID FROM  ��������˵�� WHERE (�������� LIKE '���ϲ���') OR (�������� LIKE '�Ƽ���')))
            OR NVL(B.���÷���,0)=1);
    
    IF v_batchcount>0 THEN
        mErrMsg:='[ZLSOFT]�õ����а�����һ��ԭ�������������ڷ����Ĳ��ϣ����ܳ�����[ZLSOFT]';
        RAISE mErrItem;
    END IF;  

    SELECT SUM(ʵ������) AS ʣ������,SUM(�ɱ����) AS ʣ��ɱ����,SUM(���۽��) AS ʣ�����۽��,�ⷿID,������ID,���ϵ��,����,����,����,Ч��,���Ч��,�ɱ���,����,���ۼ�,ժҪ
        INTO  V_ʣ������,v_ʣ��ɱ����,V_ʣ�����۽��,V_�ⷿID,V_������ID,V_���ϵ��,V_����,V_����,V_����,V_Ч��,v_���Ч��,V_�ɱ���,V_����,V_���ۼ�,V_ժҪ
    FROM ҩƷ�շ���¼ 
    WHERE NO=NO_IN AND ����=21 AND ҩƷID+0=����ID_IN AND ���=���_IN
    GROUP BY �ⷿID,������ID,���ϵ��,����,����,����,Ч��,���Ч��,�ɱ���,����,���ۼ�,ժҪ;

    IF nvl(V_ʣ������,0)=0 THEN 
            mErrMsg:='[ZLSOFT]�õ����а���һ���Ĳ����Ѿ����������,�����ٳ壡[ZLSOFT]';
            RAISE mErrItem; 
    END IF ;

    --������������ʣ��������������
    IF V_ʣ������<��������_IN THEN
        mErrMsg:='[ZLSOFT]�õ����а�����һ��ԭ�������������ڷ����Ĳ��ϣ����ܳ�����[ZLSOFT]';
        RAISE mErrItem;
    END IF;

    V_�ɱ����:= ROUND(��������_IN/v_ʣ������*v_ʣ��ɱ����,v_С��);
    V_���۽��:= ROUND(��������_IN/v_ʣ������*V_ʣ�����۽��,v_С��);
    V_���:=round(V_���۽��-V_�ɱ����,v_С��);

    SELECT ҩƷ�շ���¼_ID.NEXTVAL INTO V_�շ�ID  FROM DUAL;

    Insert INTO ҩƷ�շ���¼
        (ID,��¼״̬,����,NO,���,�ⷿID,������ID,���ϵ��,
        ҩƷID,����,����,����,Ч��,���Ч��,��д����,ʵ������,�ɱ���,
        �ɱ����,���ۼ�,���۽��,���,ժҪ,������,��������,�����,�������)
    VALUES 
        (V_�շ�ID,DECODE(ԭ��¼״̬_IN,1,2,ԭ��¼״̬_IN+2),21,NO_IN,���_IN,V_�ⷿID,V_������ID,V_���ϵ��,
        ����ID_IN,V_����,V_����,V_����,V_Ч��,v_���Ч��,-��������_IN,-��������_IN,V_�ɱ���,-V_�ɱ����,V_���ۼ�,-V_���۽��,
        -V_���,V_ժҪ,������_IN,��������_IN,������_IN,��������_IN);

    --ԭ�����ֲ������Ĳ���,�ڳ���ʱ��Ҫ������
    BEGIN 
        SELECT COUNT(*) INTO V_��¼��
        FROM ҩƷ�շ���¼ A,�������� B
        WHERE A.ҩƷID=B.����ID     AND B.����ID+0=����ID_IN    AND A.NO=NO_IN    AND A.���� = 21    
            AND MOD(A.��¼״̬,3)=0    AND NVL(A.����,0)>0
            AND (NVL(B.�ⷿ����,0)=0 OR 
            (NVL(B.���÷���,0)=0 AND A.�ⷿID IN (SELECT ����ID FROM  ��������˵�� WHERE (�������� LIKE '���ϲ���') OR (�������� LIKE '�Ƽ���'))));
    EXCEPTION 
        WHEN OTHERS THEN
            V_��¼��:=0;
    END;
    IF V_��¼��>0 THEN
        V_����:=0;
    ELSE
        V_����:=NVL (V_����,0);
    END IF;
    --����ҩƷ�������Ӧ����
    UPDATE ҩƷ���
    SET �������� = NVL (��������,0) + NVL (��������_IN,0),
        ʵ������ = NVL (ʵ������,0) + NVL (��������_IN,0),
        ʵ�ʽ�� = NVL (ʵ�ʽ��,0) + NVL (V_���۽��,0),
        ʵ�ʲ�� = NVL (ʵ�ʲ��,0) + NVL (V_���,0)
    WHERE �ⷿID = V_�ⷿID
        AND ҩƷID = ����ID_IN
        AND NVL (����,0) = v_����
        AND ���� = 1;

    IF SQL%NOTFOUND THEN
        Insert INTO ҩƷ���
            (�ⷿID,ҩƷID,���Ч��,����,����,��������,ʵ������,ʵ�ʽ��,ʵ�ʲ��,�ϴ�����,Ч��)
        VALUES 
            (V_�ⷿID,����ID_IN,v_���Ч��,V_����,1,��������_IN,��������_IN,V_���۽��,V_���,V_����,V_Ч��);
    END IF;

    DELETE
    FROM ҩƷ���
    WHERE �ⷿID = V_�ⷿID
        AND ҩƷID = ����ID_IN
        AND NVL (��������,0) = 0
        AND NVL (ʵ������,0) = 0
        AND NVL (ʵ�ʽ��,0) = 0
        AND NVL (ʵ�ʲ��,0) = 0;

    --����ҩƷ�շ����ܱ����Ӧ����

    UPDATE ҩƷ�շ�����
    SET ���� = NVL (����,0) + NVL (��������_IN,0),
         ��� = NVL (���,0) + NVL (V_���۽��,0),
         ��� = NVL (���,0) + NVL (V_���,0)
    WHERE ���� = TRUNC (��������_IN)
        AND �ⷿID = V_�ⷿID
        AND ҩƷID = ����ID_IN
        AND ���ID = V_������ID
        AND ���� = 21;

    IF SQL%NOTFOUND THEN
        Insert INTO ҩƷ�շ�����
            (����,�ⷿID,ҩƷID,���ID,����,����,���,���)
        VALUES 
            (TRUNC (��������_IN),V_�ⷿID,����ID_IN,V_������ID,21,��������_IN,V_���۽��,V_���);
    END IF;
EXCEPTION
    WHEN mErrItem THEN
        RAISE_APPLICATION_ERROR (-20101,mErrMsg);
    WHEN OTHERS THEN ZL_ERRORCENTER (SQLCODE,SQLERRM);
END zl_������������_strike;
/

-----------------------------------------------------------
-- �����̵�ĳ�������
--˵�������ȸ�ԭ���ݵļ�¼״̬Ϊ3;
--������һ�ŵ��ݺ���ͬ����¼״̬Ϊ2�������ͽ��Ϊ���ĳ�������;
--������ҩƷ�շ����ܱ����Ӧ�����ͽ�
--�������
--���룺NO_IN,�����_IN
------------------------------------------------------------

CREATE OR REPLACE PROCEDURE zl_�����̵�_strike (
    NO_IN        IN ҩƷ�շ���¼.NO%TYPE,
    �����_IN    IN ҩƷ�շ���¼.�����%TYPE)
IS
    mErrMsg        varchar2(100);
    mErrItem    EXCEPTION;
    v_BatchCount    integer;    --ԭ���������ڷ����Ĳ��ϵ�����
    V_COUNT        INTEGER;    --ԭ�����ֲ�����
    V_����        ҩƷ�շ���¼.����%TYPE;

    CURSOR C_ҩƷ�շ���¼  IS
    SELECT ID,ʵ������,���۽��,���,�ⷿID,ҩƷID ����id,����,����,Ч��,����,������ID,���ϵ��
    FROM ҩƷ�շ���¼
    WHERE NO = NO_IN AND ���� = 22 AND ��¼״̬ = 2
    ORDER BY ҩƷID;
BEGIN
    UPDATE ҩƷ�շ���¼
    SET ��¼״̬ = 3
    WHERE NO = NO_IN AND ���� = 22     AND ��¼״̬ = 1;

    IF SQL%ROWCOUNT = 0 THEN
        mErrMsg:='[ZLSOFT]�õ����Ѿ������˳�����[ZLSOFT]';
        RAISE mErrItem;
    END IF;
   
    --��Ҫ���ԭ���������ڷ����Ĳ��ϣ����ܶ������
    SELECT COUNT(*) INTO v_BatchCount 
    FROM  ҩƷ�շ���¼ a,�������� b
    WHERE a.ҩƷid=b.����id AND a.no=NO_IN  AND a.����=22 AND a.��¼״̬=3 AND nvl(a.����,0)=0
        AND ((NVL(B.�ⷿ����,0)=1 
        AND  A.�ⷿID NOT IN (SELECT ����ID FROM  ��������˵�� WHERE (�������� LIKE '���ϲ���') OR (�������� LIKE '�Ƽ���'))) OR 
                    NVL(B.���÷���,0)=1);
        IF v_batchcount>0 THEN
        mErrMsg:='[ZLSOFT]�õ����а�����һ��ԭ�������������ڷ����Ĳ��ϣ����ܳ�����[ZLSOFT]';
        RAISE mErrItem;
    END IF;  

    Insert INTO ҩƷ�շ���¼
        (ID,��¼״̬,����,NO,���,�ⷿID,������ID,���ϵ��,
        ҩƷID,����,����,����,Ч��,���Ч��,��д����,����,ʵ������,
        �ɱ���,�ɱ����,���ۼ�,���۽��,���,ժҪ,
        ������,��������,�����,�������,Ƶ��)
        SELECT ҩƷ�շ���¼_ID.Nextval,2,����,NO,���,�ⷿID,������ID,���ϵ��,a.ҩƷID,
            DECODE(NVL(a.����,0),0,NULL,(DECODE(NVL(b.�ⷿ����,0),0,NULL,a.����))),
            a.����,����,a.Ч��,a.���Ч��,��д����,a.����,
            -ʵ������,a.�ɱ���,�ɱ����,���ۼ�,-���۽��,-���,ժҪ,
            �����_IN,SYSDATE,�����_IN,SYSDATE,Ƶ��
        FROM ҩƷ�շ���¼ a,�������� b
        WHERE NO = NO_IN AND a.ҩƷid=b.����id AND ���� = 22 AND ��¼״̬ = 3;

    FOR V_ҩƷ�շ���¼ IN C_ҩƷ�շ���¼ LOOP
        --ԭ�����ֲ������Ĳ���,��C����ʱ��Ҫ������
        BEGIN 
            SELECT COUNT(*) INTO V_COUNT
            FROM ҩƷ�շ���¼ A,�������� B
            WHERE B.����ID+0=V_ҩƷ�շ���¼.����ID
                AND A.NO=NO_IN AND a.ҩƷid=b.����id
                AND A.���� = 22
                and a.�ⷿid+0=V_ҩƷ�շ���¼.�ⷿid
                AND A.��¼״̬ = 3 
                AND NVL(A.����,0)>0
                AND (NVL(B.�ⷿ����,0)=0 OR 
                (NVL(B.���÷���,0)=0 AND A.�ⷿID IN (SELECT ����ID FROM  ��������˵�� WHERE (�������� LIKE '���ϲ���') OR (�������� LIKE '�Ƽ���'))));
        EXCEPTION 
            WHEN OTHERS THEN
            V_COUNT:=0;
        END;
        IF V_COUNT>0 THEN
            V_����:=0;
        ELSE
            V_����:=NVL (V_ҩƷ�շ���¼.����,0);
        END IF;

        --����ҩƷ�������Ӧ����
        UPDATE ҩƷ���
        SET ��������=NVL(��������,0)+NVL(V_ҩƷ�շ���¼.ʵ������,0)*V_ҩƷ�շ���¼.���ϵ��,
            ʵ������=NVL(ʵ������,0)+NVL(V_ҩƷ�շ���¼.ʵ������,0)*V_ҩƷ�շ���¼.���ϵ��,
            ʵ�ʽ��=NVL(ʵ�ʽ��,0)+NVL(V_ҩƷ�շ���¼.���۽��,0)*V_ҩƷ�շ���¼.���ϵ��,
            ʵ�ʲ��=NVL(ʵ�ʲ��,0)+NVL(V_ҩƷ�շ���¼.���,0)*V_ҩƷ�շ���¼.���ϵ��
        WHERE �ⷿID = V_ҩƷ�շ���¼.�ⷿID
            AND ҩƷID = V_ҩƷ�շ���¼.����ID
            AND NVL (����,0) = V_����
            AND ���� = 1;

        IF SQL%NOTFOUND THEN
            Insert INTO ҩƷ���
                (�ⷿID,ҩƷID,����,����,��������,ʵ������,
                ʵ�ʽ��,ʵ�ʲ��,�ϴ�����,�ϴβ���,Ч��)
            VALUES (
                V_ҩƷ�շ���¼.�ⷿID,
                V_ҩƷ�շ���¼.����ID,
                V_����,
                1,
                V_ҩƷ�շ���¼.ʵ������*V_ҩƷ�շ���¼.���ϵ��,
                V_ҩƷ�շ���¼.ʵ������*V_ҩƷ�շ���¼.���ϵ��,
                V_ҩƷ�շ���¼.���۽��*V_ҩƷ�շ���¼.���ϵ��,
                V_ҩƷ�շ���¼.���*V_ҩƷ�շ���¼.���ϵ��,
                V_ҩƷ�շ���¼.����,
                V_ҩƷ�շ���¼.����,
                V_ҩƷ�շ���¼.Ч��);
        END IF;


        DELETE 
        FROM ҩƷ���
        WHERE �ⷿID = V_ҩƷ�շ���¼.�ⷿID
            AND ҩƷID = V_ҩƷ�շ���¼.����ID
            AND nvl(��������,0)=0 
            And nvl(ʵ������,0)=0 
            And nvl(ʵ�ʽ��,0)=0 
            And nvl(ʵ�ʲ��,0)=0;

        --����ҩƷ�շ����ܱ����Ӧ����
        UPDATE ҩƷ�շ�����
        SET ����=NVL(����,0)+NVL(V_ҩƷ�շ���¼.ʵ������,0)*V_ҩƷ�շ���¼.���ϵ��,
            ���=NVL(���,0)+NVL(V_ҩƷ�շ���¼.���۽��,0)*V_ҩƷ�շ���¼.���ϵ��,
            ���=NVL(���,0)+NVL(V_ҩƷ�շ���¼.���,0)*V_ҩƷ�շ���¼.���ϵ��
        WHERE ���� = TRUNC (SYSDATE)
            AND �ⷿID = V_ҩƷ�շ���¼.�ⷿID
            AND ҩƷID = V_ҩƷ�շ���¼.����ID
            AND ���ID = V_ҩƷ�շ���¼.������ID
            AND ���� = 22;

        IF SQL%NOTFOUND THEN
            Insert INTO ҩƷ�շ�����
                (����,�ⷿID,ҩƷID,���ID,����,����,���,���)
            VALUES (TRUNC (SYSDATE),V_ҩƷ�շ���¼.�ⷿID,V_ҩƷ�շ���¼.����ID,V_ҩƷ�շ���¼.������ID,22,
            V_ҩƷ�շ���¼.ʵ������*V_ҩƷ�շ���¼.���ϵ��,V_ҩƷ�շ���¼.���۽��*V_ҩƷ�շ���¼.���ϵ��,
            V_ҩƷ�շ���¼.���*V_ҩƷ�շ���¼.���ϵ��);
        END IF;
    END LOOP;
EXCEPTION
    WHEN mErrItem THEN
        Raise_application_error (-20101,mErrMsg);
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE,SQLERRM);
END zl_�����̵�_strike;
/
-------------------------------------------------
--���ܲ��ϵ����λ�ʱ�����ԣ�ֱ��ԭ������
-------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_��������_Insert (
	NO_IN		IN ҩƷ�շ���¼.NO%TYPE,
	���_IN		IN ҩƷ�շ���¼.���%TYPE,
	�ⷿID_IN	IN ҩƷ�շ���¼.�ⷿID%TYPE,
	�Է�����ID_IN   IN ҩƷ�շ���¼.�Է�����ID%TYPE,
	����ID_IN	IN ҩƷ�շ���¼.ҩƷID%TYPE,
	����_IN		IN ҩƷ�շ���¼.����%TYPE,
	��д����_IN	IN ҩƷ�շ���¼.��д����%TYPE,
	ʵ������_IN	IN ҩƷ�շ���¼.ʵ������%TYPE,
	�ɱ���_IN	IN ҩƷ�շ���¼.�ɱ���%TYPE,
	�ɱ����_IN	IN ҩƷ�շ���¼.�ɱ����%TYPE,
	���ۼ�_IN	IN ҩƷ�շ���¼.���ۼ�%TYPE,
	���۽��_IN	IN ҩƷ�շ���¼.���۽��%TYPE,
	���_IN		IN ҩƷ�շ���¼.���%TYPE,
	������_IN	IN ҩƷ�շ���¼.������%TYPE,
	����_IN		IN ҩƷ�շ���¼.����%TYPE := NULL,
	����_IN		IN ҩƷ�շ���¼.����%TYPE := NULL,
	Ч��_IN		IN ҩƷ�շ���¼.Ч��%TYPE := NULL,
	���Ч��_IN	IN ҩƷ�շ���¼.���Ч��%TYPE := NULL,
	ժҪ_IN		IN ҩƷ�շ���¼.ժҪ%TYPE := NULL,
	��������_IN	IN ҩƷ�շ���¼.��������%TYPE := NULL
)
IS
	mErrItem	EXCEPTION ;
	mErrMsg		varchar2(100);

	V_lngID		ҩƷ�շ���¼.ID%TYPE;--�շ�ID
	V_������ID    ҩƷ�շ���¼.������ID%TYPE;--������ID
	V_�������ID    ҩƷ�շ���¼.������ID%TYPE;--������ID
	v_�¿��	ϵͳ������.����ֵ%type;
	v_��ȷ����	ϵͳ������.����ֵ%type;

	V_����		�շ���ĿĿ¼.����%TYPE;
	V_��������	ҩƷ���.��������%TYPE;
	V_����		ҩƷ�շ���¼.����%TYPE := NULL;--��Ҫ��������ʵ�з�������Ĳ���
	V_�Ƿ����	INTEGER;--�ж�����Ƿ��������   1:������0��������
	V_�ⷿ����	INTEGER;--�ж�����Ƿ��������   1:������0��������
	V_���÷���	INTEGER;--�ж�����Ƿ��������   1:������0��������
	intRecords	NUMBER ;
BEGIN
	BEGIN
		SELECT nvl(����ֵ,'0') INTO v_�¿�� FROM ϵͳ������ WHERE ������=95;
	EXCEPTION 
		WHEN OTHERS THEN v_�¿��:='-99';
	END;

	IF v_�¿��='-99' THEN 
		mErrMsg:='[ZLSOFT]��ϵͳ��������"������¿��ÿ��"����,����ϵͳ��Ա��ϵ![ZLSOFT]';
		RAISE mErrItem;
	END IF ;

	--ֻ������ȷ���ε�����²����¿��ÿ��
	BEGIN
		SELECT nvl(����ֵ,'0') INTO v_��ȷ���� FROM ϵͳ������ WHERE ������=83;
	EXCEPTION 
		WHEN OTHERS THEN v_��ȷ����:='-99';
	END;

	
	IF v_��ȷ����='-99' THEN 
		mErrMsg:='[ZLSOFT]��ϵͳ��������"������������������"����,����ϵͳ��Ա��ϵ![ZLSOFT]';
		RAISE mErrItem;
	END IF ;

	--�����ҳ���ͳ������ID
	SELECT B.ID INTO V_������ID
	FROM ҩƷ�������� A,ҩƷ������ B
	WHERE A.���ID = B.ID AND A.���� = 34 AND B.ϵ�� = 1 AND ROWNUM < 2;

	SELECT B.ID INTO V_�������ID
	FROM ҩƷ�������� A,ҩƷ������ B
	WHERE A.���ID = B.ID AND A.���� = 34 AND B.ϵ�� = -1 AND ROWNUM < 2;


	SELECT ҩƷ�շ���¼_ID.Nextval INTO V_lngID FROM Dual;

	--�������Ϊ������һ��
	Insert INTO ҩƷ�շ���¼
		(ID,��¼״̬,����,NO,���,�ⷿID,�Է�����ID,������ID,���ϵ��,ҩƷID,����,����,����,Ч��,���Ч��,
		��д����,ʵ������,�ɱ���,�ɱ����,���ۼ�,���۽��,���,ժҪ,������,��������,��ҩ��ʽ)
	VALUES (ҩƷ�շ���¼_ID.Nextval,1,19,NO_IN,���_IN,�ⷿID_IN,�Է�����ID_IN,
		V_�������ID,-1,����ID_IN,����_IN,����_IN,����_IN,Ч��_IN,���Ч��_IN,
		��д����_IN,ʵ������_IN,�ɱ���_IN,�ɱ����_IN,���ۼ�_IN,���۽��_IN,
		���_IN,ժҪ_IN,������_IN,��������_IN,1);

	
	IF to_number(v_�¿��,'99999') =1 AND to_number(v_��ȷ����,'99999')=1 THEN 

		--��Ҫ�¿��ÿ��

		--�ж��Ƿ��п��ÿ��
		IF ����_IN > 0 THEN
			BEGIN
			    SELECT ��������    INTO V_��������
			    FROM ҩƷ���
			    WHERE ҩƷID = ����ID_IN
				AND NVL (����,0) = ����_IN
				AND �ⷿID = �ⷿID_IN
				AND ���� = 1
				AND ROWNUM = 1;
			EXCEPTION
			    WHEN OTHERS THEN
				V_�������� := 0;
			END;

			IF V_�������� - ʵ������_IN < 0 THEN
			    SELECT ���� INTO V_����    FROM �շ���ĿĿ¼ WHERE ID = ����ID_IN;
			    mErrMsg:='[ZLSOFT]����Ϊ' || V_���� || ',����Ϊ' || ����_IN ||'�ķ����������' || CHR (10) || CHR (13) || '���ÿ������������[ZLSOFT]';
			    RAISE mErrItem;
			END IF;
		END IF;

		UPDATE ҩƷ���
		SET �������� = NVL(��������, 0) - ʵ������_IN
		WHERE �ⷿID = �ⷿID_IN AND ҩƷID = ����ID_IN 
			AND NVL(����, 0) = NVL(����_IN, 0) AND ���� = 1;

		IF SQL%ROWCOUNT = 0 THEN
			INSERT INTO ҩƷ���(�ⷿID, ҩƷID, ����, ����, ��������)
			VALUES(�ⷿID_IN, ����ID_IN, NVL(����_IN, 0), 1, -ʵ������_IN);
		END IF;

		--ͬʱ���¿����
		DELETE
		FROM ҩƷ���
		WHERE �ⷿID = �ⷿID_IN
			AND ҩƷID = ����ID_IN
			AND nvl(��������,0) = 0
			AND nvl(ʵ������,0) = 0
			AND nvl(ʵ�ʽ��,0) = 0
			AND nvl(ʵ�ʲ��,0) = 0;

		--�������ж����Ĳ����Ƿ��Ƿ����������
		SELECT NVL (�ⷿ����,0),nvl(���÷���,0) INTO V_�ⷿ����,v_���÷���
		FROM ��������
		WHERE ����ID = ����ID_IN;

		V_�Ƿ���� := 0;
		IF v_���÷���=0 then
			IF V_�ⷿ���� = 1 THEN
			    BEGIN
				SELECT DISTINCT 0 INTO V_�Ƿ����
				FROM ��������˵��
				WHERE ( (�������� LIKE '���ϲ���')    OR (�������� LIKE '�Ƽ���')) AND ����ID = �Է�����ID_IN;
			    EXCEPTION
				WHEN OTHERS THEN V_�Ƿ���� := 1;
			    END;
			END IF;
		ELSE 
			V_�Ƿ���� := 1;
		END if;


		IF V_�Ƿ���� = 1 AND NVL (����_IN,0) = 0 THEN--�������ҳ��ⲻ����
			V_���� := V_lngID;
		ELSIF    V_�Ƿ���� = 0 THEN--��ⲻ����
			V_���� := 0;
		ELSIF NVL (����_IN,0) <> 0 THEN--�������ҳ���Ҳ����
			V_���� := ����_IN;
		END IF;


		--�������Ϊ�����һ��
		Insert INTO ҩƷ�շ���¼
		(ID,��¼״̬,����,NO,���,�ⷿID,�Է�����ID,������ID,���ϵ��,ҩƷID,����,����,����,Ч��,���Ч��,
		��д����,ʵ������,�ɱ���,�ɱ����,���ۼ�,���۽��,���,ժҪ,������,��������,��ҩ��ʽ)
		VALUES (V_lngID,1,19,NO_IN,���_IN + 1,�Է�����ID_IN,�ⷿID_IN,V_������ID,
		1,����ID_IN,V_����,����_IN,����_IN,Ч��_IN,���Ч��_IN,��д����_IN,ʵ������_IN,
		�ɱ���_IN,�ɱ����_IN,���ۼ�_IN,���۽��_IN,���_IN,ժҪ_IN,
		������_IN,��������_IN,1);

		--����Ƿ������ͬ������ͬ���ε����ݣ�������ڲ�������
		SELECT COUNT(*) INTO intRecords 
		FROM ҩƷ�շ���¼
		WHERE ����=19 AND NO=NO_IN AND ���ϵ��=-1 AND ҩƷID+0=����ID_IN AND Nvl(����,0)=NVL(����_IN,0);

		IF intRecords>1 THEN
			SELECT ���� INTO V_���� FROM �շ���ĿĿ¼ WHERE ID = ����ID_IN;
			mErrMsg:='[ZLSOFT]����Ϊ'||V_����||'��ҩƷ�����ڶ����ظ��ļ�¼����ϲ�Ϊһ����¼��[ZLSOFT]';
			RAISE mErrItem;
		END IF;
	ELSE 
		--�������Ϊ�����һ��
		Insert INTO ҩƷ�շ���¼
		(ID,��¼״̬,����,NO,���,�ⷿID,�Է�����ID,������ID,���ϵ��,ҩƷID,����,����,����,Ч��,���Ч��,
		��д����,ʵ������,�ɱ���,�ɱ����,���ۼ�,���۽��,���,ժҪ,������,��������,��ҩ��ʽ)
		VALUES (V_lngID,1,19,NO_IN,���_IN + 1,�Է�����ID_IN,�ⷿID_IN,V_������ID,
		1,����ID_IN,����_IN,����_IN,����_IN,Ч��_IN,���Ч��_IN,��д����_IN,ʵ������_IN,
		�ɱ���_IN,�ɱ����_IN,���ۼ�_IN,���۽��_IN,���_IN,ժҪ_IN,
		������_IN,��������_IN,1);

	END IF ;
        
EXCEPTION
	WHEN mErrItem THEN 
		raise_application_error(-20101,mErrMsg);
	WHEN OTHERS THEN
	        zl_ErrorCenter (SQLCODE,SQLERRM);
END zl_��������_Insert;
/

CREATE OR REPLACE PROCEDURE zl_��ҩҩƷ_INSERT(
    ����ID_IN IN ������ĿĿ¼.����ID%TYPE := NULL,
    ID_IN IN ������ĿĿ¼.ID%TYPE,
    ����_IN IN ������ĿĿ¼.����%TYPE := NULL,
    ��ʶ��_IN IN ҩƷ���.��ʶ��%TYPE := NULL,
    ����_IN IN ������ĿĿ¼.����%TYPE := NULL,
    ƴ��_IN IN ������Ŀ����.����%TYPE := NULL,
    ���_IN IN ������Ŀ����.����%TYPE := NULL,
    ����_IN IN �շ���ĿĿ¼.����%TYPE := NULL,
    ��λ_IN IN ������ĿĿ¼.���㵥λ%TYPE := NULL,
    ���_IN IN �շ���ĿĿ¼.���%Type := NULL,           --������
    �ۼ۵�λ_IN IN �շ���ĿĿ¼.���㵥λ%TYPE := NULL,   --������
  	����ϵ��_IN IN ҩƷ���.����ϵ��%TYPE := NULL,       --������
  	���ﵥλ_IN IN ҩƷ���.���ﵥλ%TYPE := NULL,       --������
  	�����װ_IN IN ҩƷ���.�����װ%TYPE := NULL,       --������
  	סԺ��λ_IN IN ҩƷ���.סԺ��λ%TYPE := NULL,       --������
  	סԺ��װ_IN IN ҩƷ���.סԺ��װ%TYPE := NULL,       --������
    ҩ�ⵥλ_IN IN ҩƷ���.ҩ�ⵥλ%TYPE := NULL,       
    ҩ���װ_IN IN ҩƷ���.ҩ���װ%TYPE := NULL,
	  ���쵥λ_IN IN ҩƷ���.���쵥λ%TYPE := 1,
	  ���췧ֵ_IN IN ҩƷ���.���췧ֵ%TYPE := NULL,
    �������_IN IN ҩƷ����.�������%TYPE := NULL,
    ��ֵ����_IN IN ҩƷ����.��ֵ����%TYPE := NULL,
    ��Դ���_IN IN ҩƷ����.��Դ���%TYPE := NULL,
    ��ҩ�ݴ�_IN IN ҩƷ����.��ҩ�ݴ�%TYPE := NULL,
    ҩƷ����_IN IN ҩƷ����.ҩƷ����%TYPE := NULL,
    ����ְ��_IN IN ҩƷ����.����ְ��%TYPE := '00',
    ��������_IN IN ҩƷ����.��������%TYPE := NULL,
    ����Ӧ��_IN IN ������ĿĿ¼.����Ӧ��%TYPE := NULL,
    �Ƿ�ԭ��_IN IN ҩƷ����.�Ƿ�ԭ��%TYPE := 0,
    �Ƿ���_IN IN �շ���ĿĿ¼.�Ƿ���%TYPE := NULL,
    ָ��������_IN IN ҩƷ���.ָ��������%TYPE := NULL,
    ����_IN IN ҩƷ���.����%TYPE := 95,
    ָ�����ۼ�_IN IN ҩƷ���.ָ�����ۼ�%TYPE := NULL,
    ָ�������_IN IN ҩƷ���.ָ�������%TYPE := NULL,
	  ����ѱ���_IN IN ҩƷ���.����ѱ���%TYPE := NULL,
    ҩ�ۼ���_IN IN ҩƷ���.ҩ�ۼ���%TYPE := NULL,
    ��������_IN IN �շ���ĿĿ¼.��������%TYPE := NULL,
    �������_IN IN ������ĿĿ¼.�������%TYPE := NULL,
    ���ηѱ�_IN IN �շ���ĿĿ¼.���ηѱ�%TYPE := 0,
    ҩ�����_IN IN ҩƷ���.ҩ�����%TYPE := NULL,
    ҩ������_IN IN ҩƷ���.ҩ������%TYPE := NULL,
    �ο�Ŀ¼Id_IN In ������ĿĿ¼.�ο�Ŀ¼Id%Type:=Null,
    ��������_IN IN VARCHAR2 :=NULL,      --��"|"�ָ��ı�����¼��ÿ����¼��"����^ƴ��^���"��֯
	  �ɱ���_IN in ҩƷ���.�ɱ���%TYPE := 0,
    ��ǰ�ۼ�_IN IN �շѼ�Ŀ.�ּ�%TYPE := 0,
    ����ID_IN IN �շѼ�Ŀ.������ĿID%TYPE := NULL,
    ��ͬ��λid_IN IN ҩƷ���.��ͬ��λid%TYPE := Null,
    ˵��_IN In �շ���ĿĿ¼.˵��%Type:=Null,
    �ɷ����_IN In ҩƷ���.�ɷ����%Type := NULL
) IS
    v_ҩƷID NUMBER;             --�շ���ĿID������������ȡ����
    v_Records VARCHAR2(4000);    --��ʱ��¼�������ݵ��ַ���
    v_CurrRec VARCHAR2(1000);    --�����ڱ�����¼�е�һ������
    v_Fields  VARCHAR2(1000);    --��ʱ��¼һ���������ַ���
    v_���� ������ĿĿ¼.����%TYPE;
    v_ƴ�� ������Ŀ����.����%TYPE;
    v_��� ������Ŀ����.����%TYPE;
    v_������ĿID number;
	
	--���ҿⷿid
	Cursor c_StorageID Is 
		Select DISTINCT ����id 
		From ��������˵��
		Where �������� Like '��ҩ%' Or �������� = '�Ƽ���';
	r_StorageID c_StorageID%RowType;
BEGIN
    INSERT INTO ������ĿĿ¼(���,����ID,ID,����,����,���㵥λ,
        ���㷽ʽ,ִ��Ƶ��,�����Ա�,�������,����Ӧ��,�����Ŀ,ִ�а���,�Ƽ�����,����ʱ��,����ʱ��,�ο�Ŀ¼Id)
    VALUES ('7',����ID_IN,ID_IN,����_IN,����_IN,��λ_IN,
        1,0,0,�������_IN,����Ӧ��_IN,0,0,0,sysdate,to_date('3000-01-01','YYYY-MM-DD'),�ο�Ŀ¼Id_IN);
    INSERT INTO ҩƷ����(ҩ��ID,ҩƷ����,�������,��ֵ����,��Դ���,��ҩ�ݴ�,
        ҩƷ����,����ְ��,��������,����ҩ��,�Ƿ���ҩ,�Ƿ�ԭ��,�Ƿ�Ƥ��)
    VALUES (ID_IN,'����',�������_IN,��ֵ����_IN,��Դ���_IN,��ҩ�ݴ�_IN,
        ҩƷ����_IN,����ְ��_IN,��������_IN,0,0,�Ƿ�ԭ��_IN,0);

    select �շ���ĿĿ¼_ID.NEXTVAL into v_ҩƷID from dual;
    insert into �շ���ĿĿ¼(���,ID,����,����,���,����,���㵥λ,
           ��������,�������,���ηѱ�,�Ƿ���,����ʱ��,����ʱ��,˵��)
    values ('7',v_ҩƷID,����_IN,����_IN,���_IN,����_IN,�ۼ۵�λ_IN,
           ��������_IN,�������_IN,���ηѱ�_IN,�Ƿ���_IN,sysdate,to_date('3000-01-01','YYYY-MM-DD'),˵��_IN);
    Insert INTO ҩƷ���(ҩ��ID,ҩƷID,��ʶ��,ҩƷ��Դ,����ϵ��,
           ���ﵥλ,�����װ,סԺ��λ,סԺ��װ,ҩ�ⵥλ,ҩ���װ,���쵥λ,���췧ֵ,
           ָ��������,����,ָ�����ۼ�,ָ�������,����ѱ���,ҩ�ۼ���,�ɱ���,
           �ɷ����,ҩ�����,ҩ������,���Ч��,��ͬ��λid)
    VALUES (ID_IN,v_ҩƷID,��ʶ��_IN,'����',����ϵ��_IN,
           ���ﵥλ_IN,�����װ_IN,סԺ��λ_IN,סԺ��װ_IN,
           ҩ�ⵥλ_IN,ҩ���װ_IN,���쵥λ_IN,���췧ֵ_IN,
           ָ��������_IN,����_IN,ָ�����ۼ�_IN,ָ�������_IN,����ѱ���_IN,ҩ�ۼ���_IN,�ɱ���_IN,
           �ɷ����_IN,ҩ�����_IN,ҩ������_IN,0,��ͬ��λid_IN);

    IF ƴ��_IN IS NOT NULL THEN
        INSERT INTO ������Ŀ����(������ĿID,����,����,����,����) VALUES (ID_IN,����_IN,1,ƴ��_IN,1);
        INSERT INTO �շ���Ŀ����(�շ�ϸĿid,����,����,����,����) VALUES (v_ҩƷID,����_IN,1,ƴ��_IN,1);
    END IF;
    IF ���_IN IS NOT NULL THEN
        INSERT INTO ������Ŀ����(������ĿID,����,����,����,����) VALUES (ID_IN,����_IN,1,���_IN,2);
        INSERT INTO �շ���Ŀ����(�շ�ϸĿid,����,����,����,����) VALUES (v_ҩƷID,����_IN,1,���_IN,2);
    END IF;
    IF ��������_IN IS NULL THEN
        v_Records := null;
    ELSE
        v_Records := ��������_IN||'|';
    END IF;
    WHILE v_Records IS NOT NULL LOOP
        v_CurrRec:=substr(v_Records,1,instr(v_Records,'|')-1);
        v_Fields:=v_CurrRec;
        v_����:=substr(v_Fields,1,instr(v_Fields,'^')-1);
        v_Fields:=substr(v_Fields,instr(v_Fields,'^')+1);
        v_ƴ��:=substr(v_Fields,1,instr(v_Fields,'^')-1);
        v_Fields:=substr(v_Fields,instr(v_Fields,'^')+1);
        v_���:=v_Fields;
        IF V_ƴ�� IS NOT NULL THEN
            INSERT INTO ������Ŀ����(������ĿID,����,����,����,����) VALUES (ID_IN,v_����,9,v_ƴ��,1);
            INSERT INTO �շ���Ŀ����(�շ�ϸĿid,����,����,����,����) VALUES (v_ҩƷID,v_����,9,v_ƴ��,1);
        END IF;
        IF v_��� IS NOT NULL THEN
            INSERT INTO ������Ŀ����(������ĿID,����,����,����,����) VALUES (ID_IN,v_����,9,v_���,2);
            INSERT INTO �շ���Ŀ����(�շ�ϸĿid,����,����,����,����) VALUES (v_ҩƷID,v_����,9,v_���,2);
        END IF;
        v_Records:=replace('|'||v_Records,'|'||v_CurrRec||'|');
    END LOOP;
    --������Ϣ
    if ����ID_IN is not null then
       insert into �շѼ�Ŀ(ID,ԭ��ID,�շ�ϸĿID,ԭ��,�ּ�,������ĿID,�䶯ԭ��,����˵��,������,ִ������,��ֹ����)
       values (�շѼ�Ŀ_ID.Nextval,null,v_ҩƷID,0,��ǰ�ۼ�_IN,����ID_IN,1,'��������',user,sysdate,to_date('3000-01-01','YYYY-MM-DD'));
    end if;

    --���ȱʡ�Ķ�Ӧ�������
    INSERT INTO ���Ƶ���Ӧ��(�����ļ�id,Ӧ�ó���,������Ŀid)
    SELECT A.�����ļ�id,1,ID_IN
    FROM ���Ƶ���Ӧ�� A,������ĿĿ¼ I
    Where A.������Ŀid=I.Id And I.���='7' And Ӧ�ó���=1 And Rownum<2;
    INSERT INTO ���Ƶ���Ӧ��(�����ļ�id,Ӧ�ó���,������Ŀid)
    SELECT A.�����ļ�id,2,ID_IN
    FROM ���Ƶ���Ӧ�� A,������ĿĿ¼ I
    Where A.������Ŀid=I.Id And I.���='7' And Ӧ�ó���=2 And Rownum<2;

	--�����̵�����
	For r_StorageID In c_StorageID Loop
		Insert Into ҩƷ�����޶�(�ⷿID, ҩƷID, ����, ����, �̵�����, �ⷿ��λ) Values(r_StorageID.����id, v_ҩƷID, 0, 0, '1111', null);
	End Loop;
EXCEPTION
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_��ҩҩƷ_INSERT;
/

CREATE OR REPLACE PROCEDURE zl_��ҩҩƷ_Update(
    ����ID_IN IN ������ĿĿ¼.����ID%TYPE := NULL,
    ID_IN IN ������ĿĿ¼.ID%TYPE,
    ����_IN IN ������ĿĿ¼.����%TYPE := NULL,
    ��ʶ��_IN IN ҩƷ���.��ʶ��%TYPE := NULL,
    ����_IN IN ������ĿĿ¼.����%TYPE := NULL,
    ƴ��_IN IN ������Ŀ����.����%TYPE := NULL,
    ���_IN IN ������Ŀ����.����%TYPE := NULL,
    ����_IN IN �շ���ĿĿ¼.����%TYPE := NULL,
    ��λ_IN IN ������ĿĿ¼.���㵥λ%TYPE := NULL,
    ���_IN IN �շ���ĿĿ¼.���%Type := NULL,           --������
    �ۼ۵�λ_IN IN �շ���ĿĿ¼.���㵥λ%TYPE := NULL,   --������
  	����ϵ��_IN IN ҩƷ���.����ϵ��%TYPE := NULL,       --������
  	���ﵥλ_IN IN ҩƷ���.���ﵥλ%TYPE := NULL,       --������
  	�����װ_IN IN ҩƷ���.�����װ%TYPE := NULL,       --������
  	סԺ��λ_IN IN ҩƷ���.סԺ��λ%TYPE := NULL,       --������
  	סԺ��װ_IN IN ҩƷ���.סԺ��װ%TYPE := NULL,       --������
    ҩ�ⵥλ_IN IN ҩƷ���.ҩ�ⵥλ%TYPE := NULL,
    ҩ���װ_IN IN ҩƷ���.ҩ���װ%TYPE := NULL,
	  ���쵥λ_IN IN ҩƷ���.���쵥λ%TYPE := 1,
	  ���췧ֵ_IN IN ҩƷ���.���췧ֵ%TYPE := NULL,
    �������_IN IN ҩƷ����.�������%TYPE := NULL,
    ��ֵ����_IN IN ҩƷ����.��ֵ����%TYPE := NULL,
    ��Դ���_IN IN ҩƷ����.��Դ���%TYPE := NULL,
    ��ҩ�ݴ�_IN IN ҩƷ����.��ҩ�ݴ�%TYPE := NULL,
    ҩƷ����_IN IN ҩƷ����.ҩƷ����%TYPE := NULL,
    ����ְ��_IN IN ҩƷ����.����ְ��%TYPE := '00',
    ��������_IN IN ҩƷ����.��������%TYPE := NULL,
    ����Ӧ��_IN IN ������ĿĿ¼.����Ӧ��%TYPE := NULL,
    �Ƿ�ԭ��_IN IN ҩƷ����.�Ƿ�ԭ��%TYPE := 0,
    �Ƿ���_IN IN �շ���ĿĿ¼.�Ƿ���%TYPE := NULL,
    ָ��������_IN IN ҩƷ���.ָ��������%TYPE := NULL,
    ����_IN IN ҩƷ���.����%TYPE := 95,
    ָ�����ۼ�_IN IN ҩƷ���.ָ�����ۼ�%TYPE := NULL,
    ָ�������_IN IN ҩƷ���.ָ�������%TYPE := NULL,
	  ����ѱ���_IN IN ҩƷ���.����ѱ���%TYPE := NULL,
    ҩ�ۼ���_IN IN ҩƷ���.ҩ�ۼ���%TYPE := NULL,
    ��������_IN IN �շ���ĿĿ¼.��������%TYPE := NULL,
    �������_IN IN ������ĿĿ¼.�������%TYPE := NULL,
    ���ηѱ�_IN IN �շ���ĿĿ¼.���ηѱ�%TYPE := 0,
    ҩ�����_IN IN ҩƷ���.ҩ�����%TYPE := NULL,
    ҩ������_IN IN ҩƷ���.ҩ������%TYPE := NULL,
    �ο�Ŀ¼Id_IN In ������ĿĿ¼.�ο�Ŀ¼Id%Type:=Null,
    ��������_IN IN VARCHAR2,      --��"|"�ָ��ı�����¼��ÿ����¼��"����^ƴ��^���"��֯
	  �ɱ���_IN in ҩƷ���.�ɱ���%TYPE := 0,
    ��ǰ�ۼ�_IN IN �շѼ�Ŀ.�ּ�%TYPE := 0,
    ����ID_IN IN �շѼ�Ŀ.������ĿID%TYPE := NULL,
    ��ͬ��λid_IN IN ҩƷ���.��ͬ��λid%TYPE := Null,
    ˵��_IN In �շ���ĿĿ¼.˵��%Type:=Null,
    �ɷ����_IN In ҩƷ���.�ɷ����%Type := NULL
) IS
    v_ҩƷID NUMBER;             --�շ���ĿID������������ȡ����
    v_Records VARCHAR2(4000);    --��ʱ��¼�������ݵ��ַ���
    v_CurrRec VARCHAR2(1000);    --�����ڱ�����¼�е�һ������
    v_Fields  VARCHAR2(1000);    --��ʱ��¼һ���������ַ���
    v_���� ������ĿĿ¼.����%TYPE;
    v_ƴ�� ������Ŀ����.����%TYPE;
    v_��� ������Ŀ����.����%TYPE;
    v_���� Number(2);
    Err_NotFind  EXCEPTION;
BEGIN
    UPDATE ������ĿĿ¼
    SET ����ID=����ID_IN,����=����_IN,����=����_IN,���㵥λ=��λ_IN,�������=�������_IN,����Ӧ��=����Ӧ��_IN,�ο�Ŀ¼Id=�ο�Ŀ¼Id_IN
    WHERE ID=ID_IN;
    IF SQL%ROWCOUNT=0 THEN
        RAISE Err_NotFind;
    END IF;
    UPDATE ҩƷ����
    SET �������=�������_IN,��ֵ����=��ֵ����_IN,��Դ���=��Դ���_IN,��ҩ�ݴ�=��ҩ�ݴ�_IN,
        ҩƷ����=ҩƷ����_IN,����ְ��=����ְ��_IN,��������=��������_IN,
        �Ƿ�ԭ��=�Ƿ�ԭ��_IN
    WHERE ҩ��ID=ID_IN;

    select ҩƷid into v_ҩƷID from ҩƷ��� where ҩ��id=ID_IN and rownum<2;
    update �շ���ĿĿ¼
    set ����=����_IN,����=����_IN,����=����_IN,���㵥λ=�ۼ۵�λ_IN,���=���_IN,
        ��������=��������_IN,�������=�������_IN,���ηѱ�=���ηѱ�_IN,�Ƿ���=�Ƿ���_IN,˵��=˵��_IN
    where ID=v_ҩƷID;

    update ҩƷ���
    set ��ʶ��=��ʶ��_IN,����ϵ��=����ϵ��_IN,
        ���ﵥλ=���ﵥλ_IN,�����װ=�����װ_IN,
        סԺ��λ=סԺ��λ_IN,סԺ��װ=סԺ��װ_IN,
        ҩ�ⵥλ=ҩ�ⵥλ_IN,ҩ���װ=ҩ���װ_IN,���쵥λ=���쵥λ_IN,���췧ֵ=���췧ֵ_IN,
        ָ��������=ָ��������_IN,����=����_IN,ָ�����ۼ�=ָ�����ۼ�_IN,ָ�������=ָ�������_IN,
		����ѱ���=����ѱ���_IN,ҩ�ۼ���=ҩ�ۼ���_IN,ҩ�����=ҩ�����_IN,ҩ������=ҩ������_IN,��ͬ��λid=��ͬ��λid_IN,
    �ɷ����=�ɷ����_IN
    where ҩƷID=v_ҩƷID;

    IF ƴ��_IN IS NULL THEN
        DELETE FROM ������Ŀ���� WHERE ������ĿID=ID_IN AND ����=1 AND ����=1;
        DELETE FROM �շ���Ŀ���� WHERE �շ�ϸĿid=v_ҩƷID AND ����=1 AND ����=1;
    ELSE
        UPDATE ������Ŀ���� SET ����=����_IN, ����=ƴ��_IN WHERE ������ĿID=ID_IN AND ����=1 AND ����=1;
        IF SQL%ROWCOUNT=0 THEN
            INSERT INTO ������Ŀ����(������ĿID,����,����,����,����) VALUES (ID_IN,����_IN,1,ƴ��_IN,1);
        END IF;
        update �շ���Ŀ���� SET ����=����_IN, ����=ƴ��_IN WHERE �շ�ϸĿID=v_ҩƷID AND ����=1 AND ����=1;
        if sql%rowcount=0 then
           insert into �շ���Ŀ����(�շ�ϸĿID,����,����,����,����) values(v_ҩƷID,����_IN,1,ƴ��_IN,1);
        end if;
    END IF;
    IF ���_IN IS NULL THEN
        DELETE FROM ������Ŀ���� WHERE ������ĿID=ID_IN AND ����=1 AND ����=2;
        DELETE FROM �շ���Ŀ���� WHERE �շ�ϸĿid=v_ҩƷID AND ����=1 AND ����=2;
    ELSE
        UPDATE ������Ŀ���� SET ����=����_IN, ����=���_IN WHERE ������ĿID=ID_IN AND ����=1 AND ����=2;
        IF SQL%ROWCOUNT=0 THEN
            INSERT INTO ������Ŀ����(������ĿID,����,����,����,����) VALUES (ID_IN,����_IN,1,���_IN,2);
        END IF;
        update �շ���Ŀ���� SET ����=����_IN, ����=���_IN WHERE �շ�ϸĿID=v_ҩƷID AND ����=1 AND ����=2;
        if sql%rowcount=0 then
           insert into �շ���Ŀ����(�շ�ϸĿID,����,����,����,����) values(v_ҩƷID,����_IN,1,���_IN,2);
        end if;
    END IF;

    DELETE FROM ������Ŀ���� WHERE ������ĿID=ID_IN AND ����=9;
    DELETE FROM �շ���Ŀ���� WHERE �շ�ϸĿid=v_ҩƷid AND ����=9;
    IF ��������_IN IS NULL THEN
        v_Records := null;
    ELSE
        v_Records := ��������_IN||'|';
    END IF;
    WHILE v_Records IS NOT NULL LOOP
        v_CurrRec:=substr(v_Records,1,instr(v_Records,'|')-1);
        v_Fields:=v_CurrRec;
        v_����:=substr(v_Fields,1,instr(v_Fields,'^')-1);
        v_Fields:=substr(v_Fields,instr(v_Fields,'^')+1);
        v_ƴ��:=substr(v_Fields,1,instr(v_Fields,'^')-1);
        v_Fields:=substr(v_Fields,instr(v_Fields,'^')+1);
        v_���:=v_Fields;
        IF V_ƴ�� IS NOT NULL THEN
            INSERT INTO ������Ŀ����(������ĿID,����,����,����,����) VALUES (ID_IN,v_����,9,v_ƴ��,1);
            INSERT INTO �շ���Ŀ����(�շ�ϸĿid,����,����,����,����) VALUES (v_ҩƷID,v_����,9,v_ƴ��,1);
        END IF;
        IF v_��� IS NOT NULL THEN
            INSERT INTO ������Ŀ����(������ĿID,����,����,����,����) VALUES (ID_IN,v_����,9,v_���,2);
            INSERT INTO �շ���Ŀ����(�շ�ϸĿid,����,����,����,����) VALUES (v_ҩƷID,v_����,9,v_���,2);
        END IF;
        v_Records:=replace('|'||v_Records,'|'||v_CurrRec||'|');
    END LOOP;
    --������Ϣ������Ѿ��з�����������ֱ�Ӹ�����Щ��Ϣ
    Select nvl(Count(*),0) Into v_���� From ҩƷ�շ���¼ Where ҩƷid=v_ҩƷID And rownum<2;
    If v_����=0 Then
        update �շ���ĿĿ¼ set �Ƿ���=�Ƿ���_IN where ID=v_ҩƷID;
        update ҩƷ��� set �ɱ���=�ɱ���_IN where ҩƷID=v_ҩƷID;
        if ����ID_IN is not null Then
           Update �շѼ�Ŀ
           Set �ּ�=��ǰ�ۼ�_IN,������ĿID=����ID_IN,�䶯ԭ��=1,����˵��='�޸Ķ���',������=User
           Where �շ�ϸĿID=v_ҩƷID
                 And (Sysdate Between ִ������ And ��ֹ���� Or Sysdate >= ִ������ And ��ֹ���� Is Null) And �䶯ԭ��=1;
           If Sql%Rowcount=0 Then
              insert into �շѼ�Ŀ(ID,ԭ��ID,�շ�ϸĿID,ԭ��,�ּ�,������ĿID,�䶯ԭ��,����˵��,������,ִ������,��ֹ����)
              values (�շѼ�Ŀ_ID.Nextval,null,v_ҩƷID,0,��ǰ�ۼ�_IN,����ID_IN,1,'��������',user,sysdate,to_date('3000-01-01','YYYY-MM-DD'));
           End If;
        end if;
    End If;
EXCEPTION
    WHEN Err_NotFind THEN
        Raise_application_error (-20101, '[ZLSOFT]��ҩƷ�����ڣ������ѱ������û�ɾ����[ZLSOFT]');
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_��ҩҩƷ_Update;
/

--by �¸��� {
--����һ������ռ������
CREATE OR REPLACE Function zl_GetTextRows(�ı�_IN Varchar2,����_IN Number) Return Number
As
	v_Rows		Number(18);
	v_Len		Number(18);
	v_Pos		Number(18);
	v_Tmp		Varchar2(4000);
	v_TmpLine	Varchar2(4000);
Begin

	v_Rows:=0;

	v_Tmp:=�ı�_IN;
	v_Pos:=Instrb(v_Tmp,chr(10));
	While v_Pos>0	loop
		v_TmpLine:=SubStrb(v_Tmp,1,v_Pos-1);

		v_Len:=Lengthb(v_TmpLine);

		v_Rows:=v_Rows+Ceil(v_Len/����_IN);

		v_Tmp:=SubStrb(v_Tmp,v_Pos+1);
		v_Pos:=Instrb(v_Tmp,chr(10));

	END loop;
	v_Len:=Lengthb(v_Tmp);
	v_Rows:=v_Rows+Ceil(v_Len/����_IN);

	If v_Rows<1 Then
		v_Rows:=1;
	End If;

	Return(v_Rows);
End;
/
----------------------------------------------------------------------------
---  INSERT   for   �����Ͻ���
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_�����Ͻ���_INSERT(
	���_IN IN �����Ͻ���.���%TYPE,
	����_IN IN �����Ͻ���.����%TYPE,
	����_IN IN �����Ͻ���.����%TYPE,
	����_IN IN �����Ͻ���.����%TYPE,
	�Ƿ񼲲�_IN IN �����Ͻ���.�Ƿ񼲲�%TYPE,
	��Ͻ���_IN IN �����Ͻ���.��Ͻ���%TYPE,
	�ϼ����_IN IN �����Ͻ���.�ϼ����%TYPE:=NULL,
	ĩ��_IN IN �������.ĩ��%TYPE:=1,
	ͬ������_IN  NUMBER:=0
)
IS
	v_Extend number(18);
	v_Parent varchar2(30);
BEGIN		
	IF ĩ��_IN=0 THEN
		IF ͬ������_IN=1 THEN
			    --����ͬ������ĳ���
			IF NVL(�ϼ����_IN,0)<>0 THEN
			    SELECT ���� INTO v_Parent FROM �����Ͻ��� WHERE ���=�ϼ����_IN;
			ELSE
			    v_Parent:=NULL;
			END IF;

			BEGIN
			    SELECT length(rtrim(����_IN))-length(rtrim(����)) INTO v_Extend
			    FROM �����Ͻ���
			    WHERE ĩ��=0 AND (�ϼ����=�ϼ����_IN OR �ϼ���� IS NULL AND NVL(�ϼ����_IN,0)=0) AND Rownum=1;
			EXCEPTION
			    WHEN OTHERS THEN v_Extend:=0;
			END;

			IF v_Extend>0 THEN
			    --���䴦��
			    IF v_Parent IS null THEN
				UPDATE �����Ͻ��� SET ����=lpad('0',v_Extend,'0')||���� WHERE ���<>���_IN AND ĩ��=0;
			    ELSE
				UPDATE �����Ͻ��� SET ����=v_Parent||lpad('0',v_Extend,'0')||substr(����,length(v_Parent)+1) WHERE ���� LIKE v_Parent||'_%' AND ĩ��=0;
			    END IF;
			END IF;

			IF v_Extend<0 THEN
			    --ѹ������
			    IF v_Parent IS null THEN
				UPDATE �����Ͻ��� SET ����=substr(����,1+abs(v_Extend)) WHERE ���<>���_IN AND ĩ��=0;
			    ELSE
				UPDATE �����Ͻ��� SET ����=v_Parent||substr(����,length(v_Parent)+1+abs(v_Extend)) WHERE ���� LIKE v_Parent||'_%' AND ĩ��=0;
			    END IF;
			END IF;

		END IF;
	END IF;
	Insert Into �����Ͻ���(���,�ϼ����,ĩ��,����,����,����,�Ƿ񼲲�,��Ͻ���) VALUES(���_IN,DECODE(�ϼ����_IN,0,NULL,�ϼ����_IN),ĩ��_IN,����_IN,����_IN,����_IN,�Ƿ񼲲�_IN,��Ͻ���_IN);
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_�����Ͻ���_INSERT;
/
----------------------------------------------------------------------------
---  UPDATE   for   �����Ͻ���
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_�����Ͻ���_UPDATE(
	���_IN IN �����Ͻ���.���%TYPE,
	����_IN IN �����Ͻ���.����%TYPE,
	����_IN IN �����Ͻ���.����%TYPE,
	����_IN IN �����Ͻ���.����%TYPE,
	�Ƿ񼲲�_IN IN �����Ͻ���.�Ƿ񼲲�%TYPE,
	��Ͻ���_IN IN �����Ͻ���.��Ͻ���%TYPE,
	�ϼ����_IN IN �����Ͻ���.�ϼ����%TYPE:=NULL,
	ͬ������_IN  NUMBER:=0
)
IS
	v_OldCode  VARCHAR2(30);  --ԭ���ı���
	v_Parent  VARCHAR2(30);  --�ϼ�����
	v_Extend  NUMBER(18);    --���䳤��(Ϊ����ʾѹ��)
	Err_NotFind  EXCEPTION;
BEGIN
	
	SELECT rtrim(����) INTO v_OldCode FROM �����Ͻ��� WHERE ���=���_IN;
	IF v_OldCode is null THEN
		RAISE Err_NotFind;
	END IF;

	--�޸���Ŀ����
	Update �����Ͻ���
		Set ����=����_IN,
		    ����=����_IN,
		    ����=����_IN,
		    �Ƿ񼲲�=�Ƿ񼲲�_IN,
		    ��Ͻ���=��Ͻ���_IN,
		    �ϼ����=DECODE(�ϼ����_IN,0,NULL,�ϼ����_IN)
	WHERE ���=���_IN;    

	--�޸ı�ϵ������������

	UPDATE �����Ͻ��� SET ����=����_IN||substr(����,length(v_OldCode)+1) WHERE ����<>����_IN And ���� LIKE v_OldCode||'_%' And ĩ��=0;

	--����ͬ������ĳ���
	IF ͬ������_IN=1 THEN
		IF NVL(�ϼ����_IN,0)<>0 THEN
		    SELECT ���� INTO v_Parent FROM �����Ͻ��� WHERE ���=�ϼ����_IN;
		ELSE
		    v_Parent:=NULL;
		END IF;

		BEGIN
		    SELECT length(rtrim(����_IN))-length(rtrim(����)) INTO v_Extend FROM �����Ͻ��� WHERE ĩ��=0 AND (�ϼ����=�ϼ����_IN OR �ϼ���� IS NULL AND nvl(�ϼ����_IN,0)=0) AND ���<>���_IN AND Rownum=1;
		EXCEPTION
		    WHEN OTHERS THEN v_Extend:=0;
		END;

		IF v_Extend>0 THEN
		    --���䴦��
		    IF v_Parent IS null THEN
			UPDATE �����Ͻ��� SET ����=lpad('0',v_Extend,'0')||����  WHERE ĩ��=0 and ��� not in (select ��� from �����Ͻ��� WHERE ĩ��=0 start with ���=���_IN connect by prior ���=�ϼ����);
		    ELSE
			UPDATE �����Ͻ���	SET ����=v_Parent||lpad('0',v_Extend,'0')||substr(����,length(v_Parent)+1) WHERE ĩ��=0 AND ���� LIKE v_Parent||'_%' and ��� not in (select ��� from �����Ͻ��� where ĩ��=0 start with ���=���_IN connect by prior ���=�ϼ����);
		    END IF;
		END IF;

		IF v_Extend<0 THEN
		    --ѹ������
		    IF v_Parent IS null THEN
			UPDATE �����Ͻ��� SET ����=substr(����,1+abs(v_Extend)) WHERE ���<>���_IN AND ĩ��=0;
		    ELSE
			UPDATE �����Ͻ��� SET ����=v_Parent||substr(����,length(v_Parent)+1+abs(v_Extend)) WHERE ���� LIKE v_Parent||'_%' AND ���<>���_IN AND ĩ��=0;
		    END IF;
		END IF;
	END IF;
EXCEPTION
	WHEN Err_NotFind THEN Raise_application_error (-20101, '[ZLSOFT]����Ŀ�����ڣ������ѱ������û�ɾ����[ZLSOFT]');
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_�����Ͻ���_UPDATE;
/

CREATE OR REPLACE PROCEDURE ZL_�����Ա����_INSERT(
	����id_IN IN �����Ա����.����id%TYPE,
	��ҳid_IN IN �����Ա����.��ҳid%TYPE,
	����id_IN IN �����Ա����.����id%TYPE,
	��¼����_IN IN �����Ա����.��¼����%TYPE,
	��¼���_IN IN �����Ա����.��¼���%TYPE,
	��������_IN IN �����Ա����.��������%TYPE,
	�ο�����_IN IN �����Ա����.�ο�����%TYPE,
	����id_IN IN �����Ա����.����id%TYPE,
	�Ƿ񼲲�_IN IN �����Ա����.�Ƿ񼲲�%TYPE:=0,
	��Ͻ���_IN IN �����Ա����.��Ͻ���%TYPE:=Null

)
IS
BEGIN
	INSERT INTO �����Ա����(����id,��ҳid,����id,��¼����,��¼���,��������,�ο�����,����id,�Ƿ񼲲�,��Ͻ���) 
	VALUES (����id_IN,DECODE(��ҳid_IN,0,NULL,��ҳid_IN),����id_IN,��¼����_IN,��¼���_IN,��������_IN,�ο�����_IN,DECODE(����id_IN,0,NULL,����id_IN),�Ƿ񼲲�_IN,��Ͻ���_IN);
EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_�����Ա����_INSERT;
/

--�����յĲ��˲�������д�˼���д����Ϊ��
CREATE OR REPLACE PROCEDURE ZL_�����Ŀ����_EMPTY(
	����id_IN	IN	�����Ա����.����id%TYPE,
	�嵥ID_IN	IN	NUMBER
)
IS
	--ҽ��
	CURSOR c_Advice(v_ҽ��id In Number) IS	SELECT * FROM ����ҽ����¼ WHERE ID=v_ҽ��id;
	r_Advice c_Advice%RowType;
	
	--�����ļ�
	CURSOR c_File(v_File number) IS	SELECT * FROM �����ļ�Ŀ¼ A where A.ID=v_File;
	r_File c_File%RowType;

	--���Ԫ��
	CURSOR c_Element(v_File number) IS	
		SELECT ����,����,B.ID,�ı�ת��,�����ı�,������ʾ,��������,����λ��,��������,����λ��,Ƕ�뷽ʽ,B.����
		FROM �����ļ���� A,����Ԫ��Ŀ¼ B
		Where A.����Ԫ��id=B.ID
		      And A.�����ļ�id=v_File
		Order By A.�������;
	
	--������Ԫ��
	CURSOR c_ElementPaper(v_Element number) IS
		Select * From ���������� Where Ԫ��ID=v_Element Order By �ؼ���;
	
	--��������
	CURSOR c_ElementVerfy(v_ҽ��id Number) IS
		SELECT DISTINCT C.������ĿID,
                               G.������,
                               zlGetReference(C.������ĿID,A.�걾��λ,DECODE(E.�Ա�,'��',1,'Ů',2,0),E.��������) AS ����ο�,
                               D.�������,
                               B.���㵥λ,
                               D.���㹫ʽ,C.�������
                        FROM ����ҽ����¼ A,
                             ������ĿĿ¼ B,
                             ���鱨����Ŀ C,
                             ������Ŀ D,
                             ������Ϣ E,
                             ����������Ŀ G
                        Where A.���ID =v_ҽ��id
                              AND E.����ID=A.����id
                              AND A.������ĿID=B.ID
                              AND C.������ĿID=B.ID
                              AND D.������ĿID=C.������ĿID
                              AND G.ID=C.������ĿID Order By C.�������;


	v_Index			Number(18);
	v_VerIndex		Number(18);
	v_����id		Number(18);
	v_��������id		Number(18);
	v_�����ļ�id		Number(18);
	v_ҽ��id		NUMBER(18);
BEGIN
	v_�����ļ�id:=0;
	v_ҽ��id:=0;
	Begin 
		Select B.�����ļ�id,C.ҽ��id Into v_�����ļ�id,v_ҽ��id From �����Ŀ�嵥 A,���Ƶ���Ӧ�� B,�����Ŀҽ�� C Where C.�嵥id=A.ID AND C.����id=����id_IN AND A.ID=�嵥ID_IN AND A.������Ŀid=B.������Ŀid AND B.Ӧ�ó���=4;
	Exception
		When Others Then v_�����ļ�id:=0;				
	End;
	
	If v_�����ļ�id>0 Then 
		Open c_Advice(v_ҽ��id);
		Fetch c_Advice Into r_Advice;

		Open c_File(v_�����ļ�id);
		Fetch c_File Into r_File;

		Select ���˲�����¼_ID.Nextval Into v_����id From Dual;
		
		ZL_���˲���_INSERT(v_����id,r_Advice.����id,r_Advice.��ҳid,r_Advice.�Һŵ�,r_Advice.Ӥ��,r_Advice.���˿���id,r_File.����,r_File.ID,r_File.����,NULL,v_ҽ��id);
		UPDATE ���˲�����¼ SET ��д��=NULL,��д����=NULL WHERE ID=v_����id;
	
		v_Index:=0;
		FOR r_Element In c_Element(v_�����ļ�id) LOOP
			v_Index:=v_Index+1;

			Select ���˲�������_ID.Nextval Into v_��������id From Dual;
			ZL_���˲�������_INSERT(v_��������id,NULL,v_����id,v_Index,r_Element.����,r_Element.����,r_Element.�ı�ת��,r_Element.�����ı�,r_Element.������ʾ,r_Element.��������,r_Element.����λ��,0,r_Element.��������,r_Element.����λ��,0,r_Element.Ƕ�뷽ʽ);
			
			--0-�ı��Σ�1-���ӱ�2-��������3-���ͼ��4-ר��ֽ
			If r_Element.����=4 Then
				--ר��ֽ
				If UPPER(r_Element.����)='ZL9CISCORE.USRVERIFYREPORT' then				
					--����ר��ֽ
					v_VerIndex:=0;
					FOR r_ElementVerfy In c_ElementVerfy(v_ҽ��id) LOOP
						
						v_VerIndex:=v_VerIndex+1;

						ZL_���˲���������_SAVE(v_��������id,v_VerIndex,2,r_ElementVerfy.������,NULL,NULL,NULL,NULL,
									NULL,NULL,NULL,r_ElementVerfy.������ĿID,r_ElementVerfy.�������,r_ElementVerfy.���㵥λ,''||''''''||r_ElementVerfy.����ο�);
					END LOOP;
				End If;

				If UPPER(r_Element.����)='ZL9CISCORE.USRMEDICALGROUP' then				
					--���С��ר��ֽ
					NULL;
				End If;

				If UPPER(r_Element.����)='ZL9CISCORE.USRMEDICALSUM' then				
					--����ܽ�ר��ֽ
					NULL;
				End If;
			End If;

			If r_Element.����=2 Then
				--������
				FOR r_ElementPaper In c_ElementPaper(r_Element.ID) LOOP
					
					ZL_���˲���������_SAVE(v_��������id,r_ElementPaper.�ؼ���,r_ElementPaper.�ؼ���,r_ElementPaper.����,r_ElementPaper.��,r_ElementPaper.��,r_ElementPaper.��,r_ElementPaper.��,
								r_ElementPaper.����,r_ElementPaper.����д,r_ElementPaper.������,r_ElementPaper.������ID,r_ElementPaper.��ֵ����,r_ElementPaper.������λ,NULL);

				END LOOP;
			End If;

		END LOOP;

		Close c_File;
		Close c_Advice;

		Update ����ҽ������ Set ����id=v_����id Where ҽ��id In (Select ID From ����ҽ����¼ Where ID=v_ҽ��id Union All Select ���ID As ID From ����ҽ����¼ Where ���ID=v_ҽ��id);
	End If;


EXCEPTION
	WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_�����Ŀ����_EMPTY;
/


CREATE OR REPLACE Procedure ZL_���ǼǼ�¼_������д(
	����id_IN		���˲���������.����id%TYPE,
	������id_IN		���˲���������.������id%TYPE,
	��������_IN		���˲���������.��������%TYPE
) IS
	v_Temp			Varchar2(255);
	v_��Ա���		��Ա��.���%Type;
	v_��Ա����		��Ա��.����%Type;
Begin
	
	--��ǰ������Ա
	v_Temp:=zl_Identity;
	v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
	v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
	v_��Ա���:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
	v_��Ա����:=Substr(v_Temp,Instr(v_Temp,',')+1);

	UPDATE ���˲��������� SET ��������=��������_IN WHERE ����id=����id_IN AND ������id=������id_IN;	
	UPDATE ���˲�����¼ SET ��д��=v_��Ա����,��д����=SYSDATE WHERE ��д�� IS NULL AND ID=(SELECT A.������¼id FROM ���˲������� A,���˲��������� B WHERE A.ID=B.����id AND A.ID=����id_IN AND ROWNUM<2);
Exception
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_���ǼǼ�¼_������д;
/
--}�¸���

--�������� By:��ͮ��
CREATE OR REPLACE PROCEDURE ZL_���˲����޶�_INSERT(
  RecID ���˲����޶���¼.������¼ID%Type,
  Writer ���˲����޶���¼.��д��%Type
)
AS
  nVersion number;
  vLastWriter varchar2(50);
  dLastDate date;
begin
  Select Max(�汾���)+1 Into nVersion From ���˲����޶���¼ Where ������¼ID=Recid;
  If nVersion Is Null Then
    nVersion:=1;
  End If;
  Select ������,�������� Into vLastWriter,dLastDate From ���˲�����¼ Where ID=RecID;

  Insert Into ���˲����޶���¼(ID,������¼ID,��д��,��д����,�汾���) Values(
    ���˲����޶���¼_ID.Nextval,RecID,vLastWriter,dLastDate,nVersion);
  Update ���˲�����¼ Set ��д��=Decode(��д��,Null,Writer,��д��),��д����=Decode(��д����,Null,SYSDATE,��д����),������=Writer,��������=Sysdate Where ID=RecID;

  Update ���˲������� Set ������¼ID=null,�����޶�ID=���˲����޶���¼_ID.CurrVal
    Where ������¼ID=RecID;
  EXCEPTION
    WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
end ZL_���˲����޶�_INSERT;
/



----------------------------------------------------------------------------
---  DELETE   for   ���ϼӳɷ���
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_���ϼӳɷ���_DELETE
IS
BEGIN
	Delete From ���ϼӳɷ���;
EXCEPTION
	WHEN OTHERS THEN
		zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_���ϼӳɷ���_DELETE;
/


----------------------------------------------------------------------------
---  INSERT   for   ���ϼӳɷ���
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_���ϼӳɷ���_INSERT(
	���_IN		IN ���ϼӳɷ���.���%TYPE,
	��ͼ�_IN	IN ���ϼӳɷ���.��ͼ�%TYPE,
	��߼�_IN	IN ���ϼӳɷ���.��߼�%TYPE,
	�ӳ���_IN	IN ���ϼӳɷ���.�ӳ���%TYPE,
	˵��_IN		IN ���ϼӳɷ���.˵��%TYPE
)
IS
BEGIN
	Insert Into ���ϼӳɷ���
		(���,��ͼ�,��߼�,�ӳ���,˵��)
		VALUES
		(���_IN,��ͼ�_IN,��߼�_IN,�ӳ���_IN,˵��_IN);
EXCEPTION
	WHEN OTHERS THEN
		zl_ErrorCenter (SQLCODE, SQLERRM);
END ZL_���ϼӳɷ���_INSERT;
/


CREATE OR REPLACE Procedure ZL_����ҽ���Ƽ�_Insert(
--���ܣ�ָ��ҽ���ļƼ�
    ҽ��ID_IN		����ҽ���Ƽ�.ҽ��ID%TYPE,
    �շ�ϸĿID_IN	����ҽ���Ƽ�.�շ�ϸĿID%TYPE,
    ����_IN			����ҽ���Ƽ�.����%TYPE,
    ����_IN			����ҽ���Ƽ�.����%TYPE,
	����_IN			����ҽ���Ƽ�.����%TYPE:=Null,
	ִ�п���ID_IN	����ҽ���Ƽ�.ִ�п���ID%Type:=Null
) IS
Begin
    Insert Into ����ҽ���Ƽ�(
        ҽ��ID,�շ�ϸĿID,����,����,����,ִ�п���ID)
    Values(
        ҽ��ID_IN,�շ�ϸĿID_IN,����_IN,����_IN,����_IN,ִ�п���ID_IN);
Exception
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_����ҽ���Ƽ�_Insert;
/


CREATE OR REPLACE PROCEDURE zl_��������_Update (
    ����ID_IN        IN ��������.����ID%TYPE,
    ����ID_IN        IN ���Ʒ���Ŀ¼.id%type,
    ����_IN            IN �շ���ĿĿ¼.����%TYPE,
    Ʒ��_IN            IN �շ���Ŀ����.����%TYPE := NULL,
    ���_IN            IN �շ���ĿĿ¼.���%TYPE,
    ����_IN            IN �շ���ĿĿ¼.����%TYPE := NULL,
    ƴ��_IN            IN �շ���Ŀ����.����%TYPE := NULL,
    ���_IN            IN �շ���Ŀ����.����%TYPE := NULL,
    ��ʶ����_IN		IN �շ���ĿĿ¼.��ʶ����%TYPE := NULL,
    ��ʶ����_IN		IN �շ���ĿĿ¼.��ʶ����%TYPE := NULL,
    ������Դ_IN        IN ��������.������Դ%TYPE := NULL,
    ��Դ���_IN        IN ��������.��Դ���%TYPE := NULL,
    ɢװ��λ_IN        IN �շ���ĿĿ¼.���㵥λ%TYPE := NULL,
    ��װ��λ_IN        IN ��������.��װ��λ%TYPE := NULL,
    ����ϵ��_IN        IN ��������.����ϵ��%TYPE := NULL,
    �Ƿ���_IN        IN �շ���ĿĿ¼.�Ƿ���%TYPE := NULL,
    ָ��������_IN        IN ��������.ָ��������%TYPE := NULL,
    ����_IN            IN ��������.����%TYPE := 95,
    ָ�����ۼ�_IN        IN ��������.ָ�����ۼ�%TYPE := NULL,
    ָ�������_IN        IN ��������.ָ�������%TYPE := NULL,
    ��������_IN        IN �շ���ĿĿ¼.��������%TYPE := NULL,
    �������_IN        IN �շ���ĿĿ¼.�������%TYPE := NULL,
    ���ηѱ�_IN        IN �շ���ĿĿ¼.���ηѱ�%TYPE := 0,
    �ⷿ����_IN        IN ��������.�ⷿ����%TYPE := NULL,
    ���÷���_IN        IN ��������.���÷���%TYPE := NULL,
    ���Ч��_IN        IN ��������.���Ч��%TYPE := NULL,
    ���Ч��_IN        IN ��������.���Ч��%TYPE := NULL,
    �޾��Բ���_IN        IN ��������.�޾��Բ���%TYPE := NULL,
    һ���Բ���_IN        IN ��������.һ���Բ���%TYPE := NULL,
    ԭ����_IN        IN ��������.ԭ����%TYPE := NULL,
    ���������_IN        IN ��������.���������%TYPE := 0,
    �ɱ���_IN        IN ��������.�ɱ���%TYPE := 0,
    ��������_IN        IN ��������.��������%TYPE := NULL,
    ��ǰ�ۼ�_IN        IN �շѼ�Ŀ.�ּ�%TYPE := 0,
    ����ID_IN        IN �շѼ�Ŀ.������ĿID%TYPE := NULL,
    ��׼�ĺ�_IN  IN ��������.��׼�ĺ�%TYPE := NULL,
    ע���̱�_IN  IN ��������.ע���̱�%TYPE := NULL
) IS
    m����ID        ������ĿĿ¼.ID%type;
    mErrMsg        varchar2(200);
    mErrItem    EXCEPTION;
    m����        integer;
    m��������    INTEGER;
    mCount         INTEGER ;
BEGIN
    mErrMsg:='��';
    --�޸�������Ŀ
    BEGIN 
        SELECT ����ID,��������  INTO  m����ID,m�������� FROM �������� WHERE ����id=����id_IN;
    EXCEPTION 
        WHEN OTHERS THEN 
        mErrMsg:='[ZLSOFT]������������Ŀ,���ܱ������û�ɾ����,����![ZLSOFT]';
    END;
    IF mErrMsg<>'��' THEN 
        RAISE mErrItem;
    END IF ;
    
    --�������ǰ�Ĳ���Ϊ��������,�����Ϊ�˲����������жϿ��
    IF m��������=1 AND ��������_IN<>1 THEN 
        BEGIN 
            SELECT count(*) INTO mCount  FROM ҩƷ��� 
            WHERE ҩƷid=����id_In AND ( nvl(��������,0)<>0 or nvl(ʵ������,0)<>0 or 
                nvl(ʵ�ʽ��,0)<>0 or nvl(ʵ�ʲ��,0)<>0);
            IF mcount <>0 THEN 
                mErrMsg:='[ZLSOFT]���������ϴ��ڿ��,����ȡ��������������,����![ZLSOFT]';
            END IF ;
        EXCEPTION 
            WHEN OTHERS THEN 
            null;
        END ;

    END IF ;
    IF mErrMsg<>'��' THEN 
        RAISE mErrItem;
    END IF ;

    UPDATE ������ĿĿ¼ 
    SET    ����id = ����id_IN,
        ���� = ����_IN,
        ���� = substr(Ʒ��_IN||' '||���_IN,1,60),
        ���㵥λ = ɢװ��λ_IN,
        ������� = �������_IN
    WHERE id=m����id;

    
    IF ƴ��_IN IS NULL THEN 
        DELETE ������Ŀ���� WHERE ������Ŀid = m����ID AND ����=1 AND ����= Ʒ��_IN  AND ����=1; 
    ELSE 
        UPDATE ������Ŀ����
        SET ���� = Ʒ��_IN,
            ���� = ƴ��_IN
        where ������Ŀid = m����ID AND ����=1 AND ����=1;
    END IF ;

    IF ���_IN IS NULL THEN 
        DELETE ������Ŀ���� WHERE ������Ŀid = m����ID AND ����=2 AND ����= Ʒ��_IN  AND ����=1; 
    ELSE 
        UPDATE ������Ŀ����
        SET ���� = Ʒ��_IN,
            ���� = ���_IN
        where ������Ŀid = m����ID AND ����=1 AND ����=2;
    END IF ;

    --�����Ϣ
    update �շ���ĿĿ¼
        set ����=����_IN,����=Ʒ��_IN,���=���_IN,��ʶ����=��ʶ����_IN,��ʶ����=��ʶ����_IN,����=����_IN,�Ƿ���=�Ƿ���_IN,���㵥λ=ɢװ��λ_IN,
        ��������=��������_IN,�������=�������_IN,���ηѱ�=���ηѱ�_IN
    where ID=����ID_IN;

    IF SQL%ROWCOUNT=0 THEN
        mErrMsg:='[ZLSOFT]���������Ͽ��ܱ������û�ɾ����,����![ZLSOFT]';
        RAISE mErrItem;
    END IF;

    --��������
    UPDATE  ��������
        SET ���Ч��=���Ч��_IN,
            ���Ч��=���Ч��_IN,
            �޾��Բ���=�޾��Բ���_IN,
            һ���Բ���=һ���Բ���_IN,
            ԭ����=ԭ����_IN,
            ��Դ���=��Դ���_IN,
            ��װ��λ=��װ��λ_IN,
            ����ϵ��=����ϵ��_IN,
            ָ��������=ָ��������_IN,
            ָ�����ۼ�=ָ�����ۼ�_IN,
            ָ�������=ָ�������_IN,
            ����=����_IN,
            �ⷿ����=�ⷿ����_IN,
            ���÷���=���÷���_IN,
            ������Դ=������Դ_IN,
            ���������=���������_IN,
            �ɱ���=�ɱ���_IN,
            ��������=��������_IN,
	    ��׼�ĺ�=��׼�ĺ�_IN,
	    ע���̱�=ע���̱�_In
    WHERE ����id=����id_IN;

    IF ƴ��_IN IS NULL THEN 
        DELETE �շ���Ŀ���� WHERE �շ�ϸĿid = ����ID_IN AND ����=1 AND ����= Ʒ��_IN  AND ����=1; 
    ELSE 
        UPDATE �շ���Ŀ����
        SET ���� = Ʒ��_IN,
            ���� = ƴ��_IN
        where �շ�ϸĿid = ����ID_IN AND ����=1 AND ����=1;
            if sql%rowcount=0 then
               INSERT INTO �շ���Ŀ����(�շ�ϸĿid,����,����,����,����) VALUES (����ID_IN,Ʒ��_IN,1,ƴ��_IN,1);
            end if;
    END IF ;

    IF ���_IN IS NULL THEN 
        DELETE �շ���Ŀ���� WHERE �շ�ϸĿid = ����ID_IN AND ����=2 AND ����= Ʒ��_IN  AND ����=1; 
    ELSE 
        UPDATE �շ���Ŀ����
        SET ���� = Ʒ��_IN,
            ���� = ���_IN
        WHERE �շ�ϸĿid = ����ID_IN AND ����=1 AND ����=2;
        if sql%rowcount=0 then
            INSERT INTO �շ���Ŀ����(�շ�ϸĿid,����,����,����,����) VALUES (����ID_IN,Ʒ��_IN,1,ƴ��_IN,2);
        end if;
    END IF ;

    IF Ʒ��_IN is null then
        delete �շ���Ŀ���� where �շ�ϸĿid=����ID_IN and ����=3;
    else
        if ƴ��_IN IS NULL THEN
            delete �շ���Ŀ���� where �շ�ϸĿid=����ID_IN and ����=1 and ����=1;
        else
            update �շ���Ŀ���� set ����=Ʒ��_IN,����=ƴ��_IN where �շ�ϸĿid=����ID_IN and ����=1 and ����=1;
            if sql%rowcount=0 then
               INSERT INTO �շ���Ŀ����(�շ�ϸĿid,����,����,����,����) VALUES (����ID_IN,Ʒ��_IN,1,ƴ��_IN,1);
            end if;
        end if;
        if ���_IN IS NULL THEN
           delete �շ���Ŀ���� where �շ�ϸĿid=����ID_IN and ����=1 and ����=2;
        else
            update �շ���Ŀ���� set ����=Ʒ��_IN,����=���_IN where �շ�ϸĿid=����ID_IN and ����=1 and ����=2;
            if sql%rowcount=0 then
               INSERT INTO �շ���Ŀ����(�շ�ϸĿid,����,����,����,����) VALUES (����ID_IN,Ʒ��_IN,1,���_IN,2);
            end if;
        end if;
    END IF;

    --������Ϣ������Ѿ��з�����������ֱ�Ӹ�����Щ��Ϣ
    Select nvl(Count(*),0) Into m���� From ҩƷ�շ���¼ Where ҩƷid=����ID_IN And rownum<2;

    If m����=0 Then 
        Update �շ���ĿĿ¼ set �Ƿ���=�Ƿ���_IN where ID=����ID_IN;
        Update �������� Set �ɱ���=�ɱ���_IN Where ����ID=����ID_IN;

        if ����ID_IN is not null Then
           Update �շѼ�Ŀ
           Set �ּ�=��ǰ�ۼ�_IN,������ĿID=����ID_IN,�䶯ԭ��=1,����˵��='�޸Ķ���',������=User
           Where �շ�ϸĿID=����ID_IN
             --And (��ֹ���� Is Null Or ��ֹ����=to_date('3000-01-01','YYYY-MM-DD'));
              And (Sysdate Between ִ������ And ��ֹ���� Or Sysdate >= ִ������ And ��ֹ���� Is Null) And �䶯ԭ��=1;
 
           If Sql%Rowcount=0 Then
              insert into �շѼ�Ŀ(ID,ԭ��ID,�շ�ϸĿID,ԭ��,�ּ�,������ĿID,�䶯ԭ��,����˵��,������,ִ������,��ֹ����)
              values (�շѼ�Ŀ_ID.Nextval,null,����ID_IN,0,��ǰ�ۼ�_IN,����ID_IN,1,'��������',user,sysdate,to_date('3000-01-01','YYYY-MM-DD'));
           End If;
        end if;
    End If;

    --���������̱Ƚ�����
    if ����_IN is not null then
        update ���������� set ����=����_IN where ����=����_IN;
        if sql%rowcount=0 then
              Insert INTO ����������(����,����,����)
              select nvl(max(to_number(����)),0)+1,����_IN,ZLSpellCode(����_IN) from ����������;
        end if;
    end if;

EXCEPTION
    WHEN mErrItem THEN 
        raise_application_error(-20101,mErrMsg);
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE,SQLERRM);
END zl_��������_Update;
/


CREATE OR REPLACE PROCEDURE zl_ҩƷ�⹺_Insert (
    NO_IN IN ҩƷ�շ���¼.NO%TYPE,
    ���_IN IN ҩƷ�շ���¼.���%TYPE,
    �ⷿID_IN IN ҩƷ�շ���¼.�ⷿID%TYPE,
    ��ҩ��λID_IN IN ҩƷ�շ���¼.��ҩ��λID%TYPE,
    ҩƷID_IN IN ҩƷ�շ���¼.ҩƷID%TYPE,
    ����_IN IN ҩƷ�շ���¼.����%TYPE := NULL,
    ����_IN IN ҩƷ�շ���¼.����%TYPE := NULL,
    Ч��_IN IN ҩƷ�շ���¼.Ч��%TYPE := NULL,
    ʵ������_IN IN ҩƷ�շ���¼.ʵ������%TYPE := NULL,
    �ɱ���_IN IN ҩƷ�շ���¼.�ɱ���%TYPE := NULL,
    �ɱ����_IN IN ҩƷ�շ���¼.�ɱ����%TYPE := NULL,
    ����_IN IN ҩƷ�շ���¼.����%TYPE := NULL,
    ���ۼ�_IN IN ҩƷ�շ���¼.���ۼ�%TYPE := NULL,
    ���۽��_IN IN ҩƷ�շ���¼.���۽��%TYPE := NULL,
    ���_IN IN ҩƷ�շ���¼.���%TYPE := NULL,
    ժҪ_IN IN ҩƷ�շ���¼.ժҪ%TYPE := NULL,	
    ������_IN IN ҩƷ�շ���¼.������%TYPE := NULL,
    ��Ʊ��_IN IN Ӧ����¼.��Ʊ��%TYPE := NULL,
    ��Ʊ����_IN IN Ӧ����¼.��Ʊ����%TYPE := NULL,
    ��Ʊ���_IN IN Ӧ����¼.��Ʊ���%TYPE := NULL,
    ��������_IN IN ҩƷ�շ���¼.��������%TYPE := NULL,
	���_IN IN ҩƷ�շ���¼.���%TYPE:=NULL,
	��Ʒ�ϸ�֤_IN IN ҩƷ�շ���¼.��Ʒ�ϸ�֤%TYPE:=NULL,
	�˲���_IN IN ҩƷ�շ���¼.��ҩ��%TYPE := NULL,
	�˲�����_IN IN ҩƷ�շ���¼.��ҩ����%TYPE :=NULL,
	����_IN IN ҩƷ�շ���¼.����%TYPE:=0,
	�˻�_IN IN NUMBER:=1,
  ��������_IN In ҩƷ�շ���¼.��������%TYPE := Null,
  ��׼�ĺ�_IN In ҩƷ�շ���¼.��׼�ĺ�%TYPE := Null
)
IS
    v_NO Ӧ����¼.NO%TYPE;		 --Ӧ����¼��NO
	v_��Ʒ�� �շ���ĿĿ¼.����%TYPE;--ͨ������
	v_��� �շ���ĿĿ¼.���%TYPE;
	v_���� �շ���ĿĿ¼.���%TYPE;
	v_��λ �շ���ĿĿ¼.���㵥λ%TYPE;
    V_lngID ҩƷ�շ���¼.ID%TYPE;--�շ�ID
	V_Ӧ��ID Ӧ����¼.ID%TYPE;	 --Ӧ����¼��ID
    V_������ID ҩƷ�շ���¼.������ID%TYPE;--������ID
    V_���ϵ�� ҩƷ�շ���¼.���ϵ��%TYPE;--���ϵ��
    V_���� ҩƷ�շ���¼.����%TYPE := NULL;--����
    v_ҩ����� INTEGER;--�Ƿ�ҩ�����    1:������0��������
    v_ҩ������ integer;--�Ƿ�ҩ������       1:������0��������
	v_ָ������ ҩƷ���.ָ��������%TYPE;

	v_������� ҩƷ���.ʵ������%TYPE;
	v_�˻����� ҩƷ���.ʵ������%TYPE;
	err_MSG VARCHAR2(255);
	ERR_NOENOUGH EXCEPTION ;
BEGIN
	--ȡ��ҩƷ����Ʒ��
	v_����:='';
	SELECT ����,���,���㵥λ INTO V_��Ʒ��,v_���,v_��λ FROM �շ���ĿĿ¼ WHERE ID=ҩƷID_IN;
	IF V_��� IS NOT NULL THEN
		IF INSTR(V_���,'|')<>0 THEN
			V_����:=SUBSTR(V_���,INSTR(V_���,'|'));
			V_���:=SUBSTR(V_���,INSTR(V_���,'|')-1);
		END IF ;
	END IF ;

    SELECT ҩƷ�շ���¼_ID.Nextval
      INTO V_lngID
      FROM Dual;
    SELECT NVL (ҩ�����, 0),NVL (ҩ������, 0),NVL(ָ��������,0) 
      INTO v_ҩ�����,v_ҩ������,v_ָ������
      FROM ҩƷ���
     WHERE ҩƷID = ҩƷID_IN;

    IF v_ҩ������=0 then
        IF v_ҩ����� = 1 THEN
            BEGIN
                SELECT DISTINCT 0 INTO v_ҩ�����
                  FROM ��������˵��
                 WHERE ((�������� LIKE '%ҩ��') OR (�������� LIKE '�Ƽ���'))
                    AND ����ID = �ⷿID_IN;
            EXCEPTION
                WHEN OTHERS THEN
                    v_ҩ����� := 1;
            END;

            IF v_ҩ����� = 1 THEN
                V_���� := V_lngID;
            END IF;
        END IF;
    else
        V_���� := V_lngID;
    END if;

    SELECT B.ID, B.ϵ��
      INTO V_������ID, V_���ϵ��
      FROM ҩƷ�������� A, ҩƷ������ B
     WHERE A.���ID = B.ID
        AND A.���� = 1
        AND ROWNUM < 2;

    Insert INTO ҩƷ�շ���¼
        (ID,��¼״̬,����,NO,���,�ⷿID,��ҩ��λID,������ID,���ϵ��,ҩƷID,
        ����,����,����,Ч��,��д����,ʵ������,�ɱ���,�ɱ����,����,���ۼ�,���۽��,
        ���,ժҪ,������,��������,��ҩ��,��ҩ����,��ҩ��ʽ,����,���,��Ʒ�ϸ�֤,��������,��׼�ĺ�)
	VALUES (V_lngID,1,1,NO_IN,���_IN,�ⷿID_IN,��ҩ��λID_IN,V_������ID,
		V_���ϵ��,ҩƷID_IN,Decode(�˻�_IN,-1,����_IN,V_����),����_IN,����_IN,Ч��_IN,�˻�_IN*ʵ������_IN,
		�˻�_IN*ʵ������_IN,�ɱ���_IN,�˻�_IN*�ɱ����_IN,����_IN,���ۼ�_IN,�˻�_IN*���۽��_IN,
		�˻�_IN*���_IN,ժҪ_IN,������_IN,��������_IN,�˲���_IN,�˲�����_IN,DECODE(�˻�_IN,-1,1,0),v_ָ������,���_IN,��Ʒ�ϸ�֤_IN,��������_IN,��׼�ĺ�_IN);

    IF ��Ʊ��_IN IS NOT NULL Then
      --����ǵ�һ����ϸ,�����Ӧ����¼��NO
    	BEGIN
    		SELECT NO INTO V_NO FROM Ӧ����¼
    		WHERE ϵͳ��ʶ=1 AND ��¼����=0 AND ��¼״̬=1
    			AND ��ⵥ�ݺ�=NO_IN AND ROWNUM<2;
    	EXCEPTION
    		WHEN OTHERS THEN V_NO:=NEXTNO(67);
    	END ; 
    	SELECT Ӧ����¼_ID.NEXTVAL INTO V_Ӧ��ID FROM DUAL;
        INSERT INTO Ӧ����¼
		(ID,��¼����,��¼״̬,��λID,NO,ϵͳ��ʶ,�շ�ID,��ⵥ�ݺ�,���ݽ��,��Ʊ��,��Ʊ����,��Ʊ���,Ʒ��,
		���,����,����,������λ,����,�ɹ���,�ɹ����,������,��������,�����,�������,ժҪ,��ĿID,���)
        VALUES (V_Ӧ��ID,0,1,��ҩ��λID_IN,V_NO,1,V_LNGID,NO_IN,�˻�_IN*���۽��_IN,��Ʊ��_IN,��Ʊ����_IN,�˻�_IN*��Ʊ���_IN,V_��Ʒ��,
		V_���,V_����,����_IN,V_��λ,�˻�_IN*ʵ������_IN,�ɱ���_IN,�˻�_IN*�ɱ����_IN,������_IN,��������_IN,NULL,NULL,ժҪ_IN,ҩƷID_IN,���_IN);
    END IF;
EXCEPTION
    WHEN ERR_NOENOUGH THEN
        RAISE_APPLICATION_ERROR (
            -20101, '[ZLSOFT]'||err_MSG||'[ZLSOFT]');
    WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_ҩƷ�⹺_Insert;
/

CREATE OR REPLACE PROCEDURE ZL_ҩƷ�⹺_VERIFY (
    NO_IN IN ҩƷ�շ���¼.NO%TYPE := NULL,
    �����_IN IN ҩƷ�շ���¼.�����%TYPE := NULL
)
IS
    ERR_ISVERIFIED EXCEPTION;
    ERR_ISBATCH EXCEPTION;
    V_BATCHCOUNT INTEGER;    --ԭ���������ڷ�����ҩƷ������
    V_��ҩ��λID ҩƷ�շ���¼.��ҩ��λID%TYPE;
    V_��Ʊ��� Ӧ����¼.��Ʊ���%TYPE;
	V_����� ҩƷ���.ʵ�ʽ��%TYPE;
	V_����� ҩƷ���.ʵ�ʲ��%TYPE;
	V_������� ҩƷ���.ʵ������%TYPE;
	V_�ɱ���   ҩƷ���.�ϴβɹ���%TYPE;

    CURSOR C_ҩƷ�շ���¼
    IS
        SELECT ID, ʵ������, ���۽��, ���, �ⷿID, ҩƷID, ����, ��ҩ��λID,
                 �ɱ���, ����, Ч��, ����, ������ID,��������,��׼�ĺ�
          FROM ҩƷ�շ���¼
         WHERE NO = NO_IN
            AND ���� = 1
            AND ��¼״̬ = 1
         ORDER BY ҩƷID;
BEGIN
    UPDATE ҩƷ�շ���¼
        SET ����� = NVL (�����_IN, �����),
             ������� = SYSDATE
     WHERE NO = NO_IN
        AND ���� = 1
        AND ��¼״̬ = 1
        AND ����� IS NULL;

    IF SQL%ROWCOUNT = 0 THEN
        RAISE ERR_ISVERIFIED;
    END IF;

    --��Ҫ���ԭ���������ڷ�����ҩƷ�����ܶ������
    SELECT COUNT(*) INTO V_BATCHCOUNT FROM
            ҩƷ�շ���¼ A,ҩƷ��� B
    WHERE A.ҩƷID=B.ҩƷID
        AND A.NO=NO_IN
        AND A.����=1
        AND A.��¼״̬=1
        AND NVL(A.����,0)=0
        AND ((NVL(B.ҩ�����,0)=1 AND A.�ⷿID NOT IN (SELECT ����ID FROM  ��������˵�� WHERE (�������� LIKE '%ҩ��') OR (�������� LIKE '�Ƽ���')))
            OR NVL(B.ҩ������,0)=1);
    IF V_BATCHCOUNT>0 THEN
        RAISE ERR_ISBATCH;
    END IF;

    --ԭ�����ֲ�������ҩƷ,�����ʱ��Ҫ������
    UPDATE ҩƷ�շ���¼ SET ����=0
    WHERE
    ID in 
    (SELECT ID
        FROM ҩƷ�շ���¼ A, ҩƷ��� B
        WHERE B.ҩƷID=A.ҩƷID
        AND A.NO=NO_IN
        AND A.���� = 1
            AND A.��¼״̬ = 1
            AND NVL(A.����,0)>0
            AND (NVL(B.ҩ�����,0)=0 OR
                (NVL(B.ҩ������,0)=0 AND A.�ⷿID IN (SELECT ����ID FROM  ��������˵�� WHERE (�������� LIKE '%ҩ��') OR (�������� LIKE '�Ƽ���'))))
            );

    FOR V_ҩƷ�շ���¼ IN C_ҩƷ�շ���¼ LOOP
        --����ҩƷ�������Ӧ����

        UPDATE ҩƷ���
            SET �������� = NVL (��������, 0) + NVL (V_ҩƷ�շ���¼.ʵ������, 0),
                 ʵ������ = NVL (ʵ������, 0) + NVL (V_ҩƷ�շ���¼.ʵ������, 0),
                 ʵ�ʽ�� = NVL (ʵ�ʽ��, 0) + NVL (V_ҩƷ�շ���¼.���۽��, 0),
                 ʵ�ʲ�� = NVL (ʵ�ʲ��, 0) + NVL (V_ҩƷ�շ���¼.���, 0),
                 �ϴι�Ӧ��ID = NVL (V_ҩƷ�շ���¼.��ҩ��λID, �ϴι�Ӧ��ID),
                 �ϴβɹ��� = NVL (V_ҩƷ�շ���¼.�ɱ���, �ϴβɹ���),
                 �ϴ����� = NVL (V_ҩƷ�շ���¼.����, �ϴ�����),
                 �ϴβ��� = NVL (V_ҩƷ�շ���¼.����, �ϴβ���),
                 Ч�� = NVL (V_ҩƷ�շ���¼.Ч��, Ч��),
                 �ϴ��������� = NVL (V_ҩƷ�շ���¼.��������, �ϴ���������),
                 ��׼�ĺ� = NVL (V_ҩƷ�շ���¼.��׼�ĺ�, ��׼�ĺ�)
         WHERE �ⷿID = V_ҩƷ�շ���¼.�ⷿID
            AND ҩƷID = V_ҩƷ�շ���¼.ҩƷID
            AND NVL (����, 0) = NVL (V_ҩƷ�շ���¼.����, 0)
            AND ���� = 1;

        IF SQL%NOTFOUND THEN
            INSERT INTO ҩƷ���
                            (
                                �ⷿID,
                                ҩƷID,
                                ����,
                                ����,
                                ��������,
                                ʵ������,
                                ʵ�ʽ��,
                                ʵ�ʲ��,
                                �ϴι�Ӧ��ID,
                                �ϴβɹ���,
                                �ϴ�����,
                                �ϴβ���,
                                Ч��,
                                �ϴ���������,
                                ��׼�ĺ�
                            )
                  VALUES (
                      V_ҩƷ�շ���¼.�ⷿID,
                      V_ҩƷ�շ���¼.ҩƷID,
                      V_ҩƷ�շ���¼.����,
                      1,
                      V_ҩƷ�շ���¼.ʵ������,
                      V_ҩƷ�շ���¼.ʵ������,
                      V_ҩƷ�շ���¼.���۽��,
                      V_ҩƷ�շ���¼.���,
                      V_ҩƷ�շ���¼.��ҩ��λID,
                      V_ҩƷ�շ���¼.�ɱ���,
                      V_ҩƷ�շ���¼.����,
                      V_ҩƷ�շ���¼.����,
                      V_ҩƷ�շ���¼.Ч��,
                      V_ҩƷ�շ���¼.��������,
                      V_ҩƷ�շ���¼.��׼�ĺ�
                  );
        END IF;

        --����������Ϊ��ļ�¼
        DELETE
          FROM ҩƷ���
         WHERE �ⷿID = V_ҩƷ�շ���¼.�ⷿID
            AND ҩƷID = V_ҩƷ�շ���¼.ҩƷID
            AND NVL(��������,0) = 0
            AND NVL(ʵ������,0) = 0
            AND NVL(ʵ�ʽ��,0) = 0
            AND NVL(ʵ�ʲ��,0) = 0;

        --����ҩƷ�շ����ܱ����Ӧ����

        UPDATE ҩƷ�շ�����
            SET ���� = NVL (����, 0) + NVL (V_ҩƷ�շ���¼.ʵ������, 0),
                 ��� = NVL (���, 0) + NVL (V_ҩƷ�շ���¼.���۽��, 0),
                 ��� = NVL (���, 0) + NVL (V_ҩƷ�շ���¼.���, 0)
         WHERE ���� = TRUNC (SYSDATE)
            AND �ⷿID = V_ҩƷ�շ���¼.�ⷿID
            AND ҩƷID = V_ҩƷ�շ���¼.ҩƷID
            AND ���ID = V_ҩƷ�շ���¼.������ID
            AND ���� = 1;

        IF SQL%NOTFOUND THEN
            INSERT INTO ҩƷ�շ�����
                            (����, �ⷿID, ҩƷID, ���ID, ����, ����, ���, ���)
                  VALUES (
                      TRUNC (SYSDATE),
                      V_ҩƷ�շ���¼.�ⷿID,
                      V_ҩƷ�շ���¼.ҩƷID,
                      V_ҩƷ�շ���¼.������ID,
                      1,
                      V_ҩƷ�շ���¼.ʵ������,
                      V_ҩƷ�շ���¼.���۽��,
                      V_ҩƷ�շ���¼.���
                  );
        END IF;

		--���¸�ҩƷ�ĳɱ���
		UPDATE ҩƷ���
		SET �ɱ���=V_ҩƷ�շ���¼.�ɱ��� 
		WHERE ҩƷID=V_ҩƷ�շ���¼.ҩƷID;
    END LOOP;

    --��Ӧ��������д���
      --�˴���һ���飬��Ҫ�ǽ��û�ж�Ӧ��Ʊ�ŵļ�¼
    BEGIN
		UPDATE Ӧ����¼
		SET �����=�����_IN,�������=SYSDATE
		WHERE ��ⵥ�ݺ�=NO_IN AND ϵͳ��ʶ=1 and ��¼����=0 And ��¼״̬=1;

        SELECT B.��λID, SUM (��Ʊ���)
          INTO V_��ҩ��λID, V_��Ʊ���
          FROM ҩƷ�շ���¼ A, Ӧ����¼ B
         WHERE A.ID = B.�շ�ID
            AND A.NO = NO_IN
            AND A.���� = 1 AND B.ϵͳ��ʶ=1
         GROUP BY B.��λID;

        IF NVL (V_��ҩ��λID, 0) <> 0 THEN
            UPDATE Ӧ�����
                SET ��� = NVL (���, 0) + NVL (V_��Ʊ���, 0)
             WHERE ��λID = V_��ҩ��λID
                AND ���� = 1;

            IF SQL%NOTFOUND THEN
                INSERT INTO Ӧ�����
                                (��λID, ����, ���)
                      VALUES (V_��ҩ��λID, 1, V_��Ʊ���);
            END IF;
        END IF;
    EXCEPTION
        WHEN NO_DATA_FOUND THEN
            NULL;
    END;
EXCEPTION
    WHEN ERR_ISVERIFIED THEN
        RAISE_APPLICATION_ERROR (
            -20101, '[ZLSOFT]�õ����Ѿ���������ˣ�[ZLSOFT]'
        );
    WHEN ERR_ISBATCH THEN
        RAISE_APPLICATION_ERROR (
            -20102, '[ZLSOFT]�õ����а�����һ��ԭ�������������ڷ�����ҩƷ��������ˣ�[ZLSOFT]'
        );
    WHEN OTHERS THEN
        ZL_ERRORCENTER (SQLCODE, SQLERRM);
END ZL_ҩƷ�⹺_VERIFY;
/


CREATE OR REPLACE PROCEDURE ZL_ҩƷ�⹺_STRIKE (
	�д�_IN IN INTEGER,
    ԭ��¼״̬_IN IN ҩƷ�շ���¼.��¼״̬%TYPE,
    NO_IN IN ҩƷ�շ���¼.NO%TYPE,
    ���_IN IN ҩƷ�շ���¼.���%TYPE,
    ҩƷID_IN IN ҩƷ�շ���¼.ҩƷID%TYPE,
    ��������_IN Number,
    ������_IN IN ҩƷ�շ���¼.������%TYPE ,
    ��������_IN IN ҩƷ�շ���¼.��������%TYPE ,
    ��Ʊ��_IN IN Ӧ����¼.��Ʊ��%TYPE := NULL,
    ��Ʊ����_IN IN Ӧ����¼.��Ʊ����%TYPE := NULL,
    ��Ʊ���_IN IN Ӧ����¼.��Ʊ���%TYPE := NULL,
    ȫ������_IN IN ҩƷ�շ���¼.ʵ������%TYPE := 0,    --���ڲ������
    �������_IN In Number :=0  --������˱�־
)
IS
    ERR_ISSTRIKED EXCEPTION;
    ERR_ISOUTSTOCK EXCEPTION;
    ERR_ISBATCH EXCEPTION;
    ERR_ISNONUM EXCEPTION;
    V_BATCHCOUNT INTEGER;    --ԭ���������ڷ�����ҩƷ������

	  v_Ӧ��ID Ӧ����¼.ID%TYPE;
    V_�ⷿID ҩƷ�շ���¼.�ⷿID%TYPE;
    V_��ҩ��λID ҩƷ�շ���¼.��ҩ��λID%TYPE;
    V_������ID ҩƷ�շ���¼.������ID%TYPE ;
    V_���� ҩƷ�շ���¼.����%TYPE ;
    V_���� ҩƷ�շ���¼.����%TYPE ;
    V_���� ҩƷ�շ���¼.����%TYPE ;
    V_Ч�� ҩƷ�շ���¼.Ч��%TYPE ;
    V_�ɱ��� ҩƷ�շ���¼.�ɱ���%TYPE ;
    V_�ɱ���� ҩƷ�շ���¼.�ɱ����%TYPE ;
    V_���� ҩƷ�շ���¼.����%TYPE ;
    V_���ۼ� ҩƷ�շ���¼.���ۼ�%TYPE ;
    V_���۽�� ҩƷ�շ���¼.���۽��%TYPE ;
    V_��� ҩƷ�շ���¼.���%TYPE ;
    V_ժҪ ҩƷ�շ���¼.ժҪ%TYPE ;
    V_ʣ������ Number;
    V_ʣ��ɱ���� ҩƷ�շ���¼.�ɱ����%Type;
    V_ʣ�����۽�� ҩƷ�շ���¼.���۽��%Type;
    V_���ϵ�� ҩƷ�շ���¼.���ϵ��%TYPE;
    V_�������� Number;
	  v_���� ҩƷ�շ���¼.����%TYPE;
    V_�������� ҩƷ�շ���¼.��������%TYPE;
    V_��׼�ĺ� ҩƷ�շ���¼.��׼�ĺ�%TYPE;

    v_�˲��� ҩƷ�շ���¼.��ҩ��%TYPE;
    v_�˲����� ҩƷ�շ���¼.��ҩ����%TYPE;
    V_��¼�� NUMBER;
	  V_��ҩ��ʽ NUMBER ;
    V_�շ�ID ҩƷ�շ���¼.ID%TYPE;

    --�Գ����������м��
    V_����� ҩƷ���.ʵ������%TYPE;
    V_ҩ����� INTEGER;
    V_ҩ������ INTEGER;
    V_�������� INTEGER;
    V_ҩ�� INTEGER;
    V_���� NUMBER;

    V_����� NUMBER(16,5);
	  V_����� NUMBER(16,5);
	  V_������� NUMBER(16,5);
  	INTDIGIT NUMBER;
	  V_��Ʊ��� NUMBER ;

BEGIN
	--��ȡ���С��λ��
    SELECT NVL(����ֵ,ȱʡֵ) INTO INTDIGIT FROM ϵͳ������ WHERE ������='���ý���λ��';

	--ȡ�˲���
	SELECT min(��ҩ��) ��ҩ��,min(��ҩ����) ��ҩ����,sum(ʵ������) ʵ������
	INTO v_�˲���,v_�˲�����,V_��������
	FROM ҩƷ�շ���¼
	WHERE NO=NO_IN AND ����=1 AND ���=���_IN Group By ��ҩ��,��ҩ���� ;

  If �������_IN=0 Then
    V_��������:=��������_IN;
  End If;

    IF �д�_IN =1 THEN
        UPDATE ҩƷ�շ���¼
            SET ��¼״̬ = DECODE(ԭ��¼״̬_IN,1,3,ԭ��¼״̬_IN+3)
         WHERE NO = NO_IN
            AND ���� = 1
            AND ��¼״̬ =ԭ��¼״̬_IN ;

        IF SQL%ROWCOUNT = 0 THEN
            RAISE ERR_ISSTRIKED;
        END IF;
    END IF;

    --��Ҫ���ԭ���������ڷ�����ҩƷ�����ܶ������
    SELECT COUNT(*) INTO V_BATCHCOUNT FROM
            ҩƷ�շ���¼ A,ҩƷ��� B
    WHERE A.ҩƷID=B.ҩƷID
        AND A.NO=NO_IN
        AND A.����=1
        AND MOD(A.��¼״̬,3)=0
        AND NVL(A.����,0)=0
        AND A.ҩƷID+0=ҩƷID_IN
        AND ((NVL(B.ҩ�����,0)=1 AND A.�ⷿID NOT IN (SELECT ����ID FROM  ��������˵�� WHERE (�������� LIKE '%ҩ��') OR (�������� LIKE '�Ƽ���')))
            or nvl(b.ҩ������,0)=1);
    IF V_BATCHCOUNT>0 THEN
        RAISE ERR_ISBATCH;
    END IF;

    SELECT SUM(A.ʵ������) AS ʣ������,SUM(A.�ɱ����) AS ʣ��ɱ����,SUM(A.���۽��) AS ʣ�����۽��,A.�ⷿID,A.��ҩ��λID,A.������ID,A.���ϵ��,NVL(A.����,0),A.����,A.����,A.Ч��,A.�ɱ���,A.����,A.���ۼ�,A.ժҪ,B.ҩ�����,B.ҩ������,A.��ҩ��ʽ,A.����,A.��������,A.��׼�ĺ�
    INTO  V_ʣ������,v_ʣ��ɱ����,V_ʣ�����۽��,V_�ⷿID,V_��ҩ��λID,V_������ID,V_���ϵ��,V_����,V_����,V_����,V_Ч��,V_�ɱ���,V_����,V_���ۼ�,V_ժҪ,V_ҩ�����,V_ҩ������,V_��ҩ��ʽ,V_����,V_��������,V_��׼�ĺ�
    FROM ҩƷ�շ���¼ A,ҩƷ��� B
    WHERE A.NO=NO_IN And A.ҩƷID=B.ҩƷID AND A.����=1 AND A.ҩƷID=ҩƷID_IN AND A.���=���_IN
    GROUP BY A.�ⷿID,A.��ҩ��λID,A.������ID,A.���ϵ��,NVL(A.����,0),A.����,A.����,A.Ч��,A.�ɱ���,A.����,A.���ۼ�,A.ժҪ,B.ҩ�����,B.ҩ������,A.��ҩ��ʽ,A.����,A.��������,A.��׼�ĺ�;

    --�жϸò�����ҩ�⻹��ҩ��
    BEGIN
        SELECT DISTINCT 0
        INTO v_ҩ��
        FROM ��������˵��
        WHERE (   (�������� LIKE '%ҩ��')
              OR (�������� LIKE '�Ƽ���'))
        AND ����ID = V_�ⷿID;
    EXCEPTION
        WHEN OTHERS THEN V_ҩ��:=1;
    END ;

    --���ݲ�������,�жϷ�������
    IF V_ҩ��=0 THEN
        v_��������:=V_ҩ������;
    ELSE
        V_��������:=V_ҩ�����;
    END IF ;

    --V_����:(ԭ����,�ַ���Ϊ����;����Ϊ��)ԭ������,�ַ���,�����̲�����
    V_����:=0;
    IF V_��������=1 AND V_����<>0 THEN
        V_����:=V_����;
    END IF ;

    --ȡ�����
    BEGIN
        SELECT Nvl(ʵ������,0) INTO V_����� FROM ҩƷ���
        WHERE �ⷿID=V_�ⷿID AND ҩƷID=ҩƷID_IN AND Nvl(����,0)=V_���� And ����=1;
    EXCEPTION
        WHEN OTHERS THEN V_�����:=0;
    END ;
   
    --������������ʣ������,ȡʣ������;����ȡ�����
    IF V_�����<V_ʣ������ And �������_IN=0 THEN
        if ȫ������_IN=1  then
            --������
            raise ERR_ISNONUM;
        Else
            v_ʣ��ɱ����:=V_�����/V_ʣ������*v_ʣ��ɱ����;
            V_ʣ�����۽��:=V_�����/V_ʣ������*V_ʣ�����۽��;
            V_ʣ������:=V_�����;
        end if ;
    END IF ;

    IF ȫ������_IN=1 And �������_IN=0 THEN
        V_��������:=V_ʣ������;
    END IF;

    --������������ʣ��������������(������˳���)
    IF ABS(V_ʣ������)<ABS(V_��������) And �������_IN=0 THEN
        RAISE ERR_ISNONUM;
    END IF;

    If �������_IN=0 Then
       V_�ɱ����:= ROUND(V_��������/v_ʣ������*v_ʣ��ɱ����,INTDIGIT);
       V_���۽��:= ROUND(V_��������/v_ʣ������*V_ʣ�����۽��,INTDIGIT);
    Else
       V_�ɱ����:= ROUND(v_ʣ��ɱ����,INTDIGIT);
       V_���۽��:= ROUND(V_ʣ�����۽��,INTDIGIT);
    End If;
    V_���:=V_���۽��-V_�ɱ����;

    SELECT ҩƷ�շ���¼_ID.NEXTVAL INTO V_�շ�ID  FROM DUAL;

    INSERT INTO ҩƷ�շ���¼
        ( ID,��¼״̬,����,NO,���,�ⷿID,��ҩ��λID,������ID,���ϵ��,ҩƷID,����,����,����,Ч��,
        ��д����,ʵ������,�ɱ���,�ɱ����,����,���ۼ�,���۽��,���,ժҪ,������,��������,��ҩ��,��ҩ����,�����,�������,��ҩ��ʽ,����,��������,��׼�ĺ�
        )
    VALUES (V_�շ�ID,DECODE(ԭ��¼״̬_IN,1,2,ԭ��¼״̬_IN+2),1,NO_IN,���_IN,V_�ⷿID,V_��ҩ��λID,
        V_������ID,1,ҩƷID_IN,V_����,V_����,V_����,V_Ч��,-V_��������, -V_��������,V_�ɱ���,-V_�ɱ����,
        V_����,V_���ۼ�,-V_���۽��,-V_���,V_ժҪ,������_IN, ��������_IN,v_�˲���,v_�˲�����, ������_IN, ��������_IN,V_��ҩ��ʽ,v_����,V_��������,V_��׼�ĺ�);


    --���ڳ����ĵ���ҲӦ�ö�Ӧ��������д���
    --ֻ�����˷�Ʊ�ŵļ�¼���д���
	v_��Ʊ���:=Nvl(��Ʊ���_IN,0);
    IF NVL (��Ʊ��_IN,' ') <> ' ' AND NVL (��Ʊ���_IN, 0)<>0 THEN
		--���ڲ�����˵ģ�Ҫ��ʣ��ķ�Ʊ���ȫ������
		IF ȫ������_IN=1 THEN
			SELECT SUM(B.��Ʊ���) INTO v_��Ʊ���
			FROM
				(SELECT ID
				FROM ҩƷ�շ���¼
				WHERE ����=1 AND NO=NO_IN AND ���=���_IN) A,Ӧ����¼ B
			WHERE A.ID=B.�շ�ID AND B.ϵͳ��ʶ=1 And B.��¼����<>-1;
		END IF;

		UPDATE Ӧ�����
            SET ��� = NVL (���, 0) - NVL (v_��Ʊ���, 0)
         WHERE ��λID = V_��ҩ��λID
            AND ���� = 1;

        IF SQL%NOTFOUND THEN
            INSERT INTO Ӧ�����
            (��λID, ����, ���)
            VALUES (V_��ҩ��λID, 1, -NVL (v_��Ʊ���, 0));
        END IF;
    END IF;

    UPDATE ҩƷ���
    SET �������� = NVL (��������, 0) -V_��������,
             ʵ������ = NVL (ʵ������, 0) - V_��������,
             ʵ�ʽ�� = NVL (ʵ�ʽ��, 0) - V_���۽��,
             ʵ�ʲ�� = NVL (ʵ�ʲ��, 0) -V_���,
             �ϴι�Ӧ��ID = V_��ҩ��λID,
             �ϴβɹ��� = V_�ɱ���,
             �ϴ����� = V_����,
             �ϴβ��� = V_����,
             Ч�� = V_Ч��,
             �ϴ���������=V_��������,
             ��׼�ĺ�=V_��׼�ĺ�
     WHERE �ⷿID = V_�ⷿID
        AND ҩƷID = ҩƷID_IN
        AND NVL (����, 0) = NVL(V_����,0)
        AND ���� = 1;

    IF SQL%NOTFOUND THEN
        INSERT INTO ҩƷ���
        (�ⷿID,ҩƷID,����,����,��������,ʵ������,ʵ�ʽ��,ʵ�ʲ��,�ϴι�Ӧ��ID,�ϴβɹ���,�ϴ�����,�ϴβ���,Ч��,�ϴ���������,��׼�ĺ�)
        VALUES
        (V_�ⷿID,ҩƷID_IN,V_����,1,-V_��������,-V_��������,-V_���۽��,-V_���,V_��ҩ��λID,V_�ɱ���,V_����,V_����,V_Ч��,V_��������,V_��׼�ĺ�) ;

    END IF;

    --����������Ϊ��ļ�¼
    DELETE
    FROM ҩƷ���
    WHERE �ⷿID = V_�ⷿID
    AND ҩƷID = ҩƷID_IN
    AND NVL (��������, 0) = 0
    AND NVL (ʵ������, 0) = 0
    AND NVL (ʵ�ʽ��, 0) = 0
    AND NVL (ʵ�ʲ��, 0) = 0;

    --����ҩƷ�շ����ܱ����Ӧ����

    UPDATE ҩƷ�շ�����
        SET ���� =NVL(����,0) - V_��������,
            ��� =NVL (���, 0) -V_���۽��,
            ��� =NVL (���, 0) -V_���
     WHERE ���� = TRUNC (SYSDATE)
        AND �ⷿID = V_�ⷿID
        AND ҩƷID = ҩƷID_IN
        AND ���ID = V_������ID
        AND ���� = 1;

    IF SQL%NOTFOUND THEN
		INSERT INTO ҩƷ�շ�����
        (����, �ⷿID, ҩƷID, ���ID, ����, ����, ���, ���)
        VALUES (
            TRUNC (SYSDATE),V_�ⷿID,ҩƷID_IN,V_������ID,1,-V_��������,-V_���۽��,-V_���);
    END IF;

	--����Ӧ����¼�ĳ�����¼(���ж�Ӧ����¼���Ƿ��Ѵ��ڸü�¼��Ӧ�ĳ�����¼,�������;��������)
	SELECT Ӧ����¼_ID.NEXTVAL INTO V_Ӧ��ID FROM DUAL;
	INSERT INTO Ӧ����¼
	(ID,��¼����,��¼״̬,��λID,NO,ϵͳ��ʶ,�շ�ID,��ⵥ�ݺ�,���ݽ��,��Ʊ��,��Ʊ����,��Ʊ���,Ʒ��,
	���,����,����,������λ,����,�ɹ���,�ɹ����,������,��������,�����,�������,ժҪ,��ĿID,���)
	SELECT V_Ӧ��ID,��¼����,DECODE(ԭ��¼״̬_IN,1,2,ԭ��¼״̬_IN+2),��λID,NO,1,V_�շ�ID,��ⵥ�ݺ�,-V_���۽��,��Ʊ��,��Ʊ����,-v_��Ʊ���,Ʒ��,
	���,����,����,������λ,-V_��������,�ɹ���,-�ɹ���*V_��������,������_IN,��������_IN,������_IN,��������_IN,ժҪ,��ĿID,���
	FROM Ӧ����¼
	WHERE �շ�ID=(SELECT ID FROM ҩƷ�շ���¼ WHERE NO=NO_IN AND ����=1 AND ���=���_IN And Mod(��¼״̬,3)=0) AND ϵͳ��ʶ=1 AND ��¼����=0;

	update Ӧ����¼
	set ��¼״̬=3
	WHERE �շ�ID=(SELECT ID FROM ҩƷ�շ���¼ WHERE NO=NO_IN AND ����=1 AND ���=���_IN And Mod(��¼״̬,3)=0) AND ϵͳ��ʶ=1 AND ��¼����=0;

EXCEPTION
    WHEN ERR_ISSTRIKED THEN
        RAISE_APPLICATION_ERROR (
            -20101, '[ZLSOFT]�õ����Ѿ������˳�����[ZLSOFT]'
        );
    WHEN ERR_ISOUTSTOCK THEN
        RAISE_APPLICATION_ERROR (
            -20102, '[ZLSOFT]�õ����е�' || ���_IN || '�е�ҩƷ�ѳ��⣬���ܳ�����[ZLSOFT]'
        );
    WHEN ERR_ISBATCH THEN
        RAISE_APPLICATION_ERROR (
            -20102, '[ZLSOFT]�õ����е�' || ���_IN || '�е�ҩƷԭ�������������ڷ�����ҩƷ�����ܳ�����[ZLSOFT]'
        );
    WHEN ERR_ISNONUM THEN
        RAISE_APPLICATION_ERROR (
            -20102, '[ZLSOFT]�õ����е�' || ���_IN || '�е�ҩƷ����������������ʣ������ݣ����ܳ�����[ZLSOFT]'
        );
    WHEN OTHERS THEN
        ZL_ERRORCENTER (SQLCODE, SQLERRM);
END ZL_ҩƷ�⹺_STRIKE;
/

CREATE OR REPLACE PROCEDURE zl_ҩƷ�������_Insert (
    NO_IN IN ҩƷ�շ���¼.NO%TYPE,
    ���_IN IN ҩƷ�շ���¼.���%TYPE,
    �ⷿID_IN IN ҩƷ�շ���¼.�ⷿID%TYPE,
    ������ID_IN IN ҩƷ�շ���¼.������ID%TYPE,
    ҩƷID_IN IN ҩƷ�շ���¼.ҩƷID%TYPE,
    ʵ������_IN IN ҩƷ�շ���¼.ʵ������%TYPE,
    �ɱ���_IN IN ҩƷ�շ���¼.�ɱ���%TYPE,
    �ɱ����_IN IN ҩƷ�շ���¼.�ɱ����%TYPE,
    ���ۼ�_IN IN ҩƷ�շ���¼.���ۼ�%TYPE,
    ���۽��_IN IN ҩƷ�շ���¼.���۽��%TYPE,
    ���_IN IN ҩƷ�շ���¼.���%TYPE,
    ������_IN IN ҩƷ�շ���¼.������%TYPE,
    ��������_IN IN ҩƷ�շ���¼.��������%TYPE,
    ժҪ_IN IN ҩƷ�շ���¼.ժҪ%TYPE := NULL,
    ����_IN IN ҩƷ�շ���¼.����%TYPE := NULL,
    ����_IN IN ҩƷ�շ���¼.����%TYPE := NULL,
    Ч��_IN IN ҩƷ�շ���¼.Ч��%TYPE := NULL,
    ��������_IN In ҩƷ�շ���¼.��������%TYPE := Null,
    ��׼�ĺ�_IN In ҩƷ�շ���¼.��׼�ĺ�%TYPE := Null
)
Is      
    V_lngID ҩƷ�շ���¼.ID%TYPE;--�շ�ID
    V_���ϵ�� ҩƷ�շ���¼.���ϵ��%TYPE;
    V_���� ҩƷ�շ���¼.����%TYPE := NULL;--����
    v_ҩ����� INTEGER;--�Ƿ�ҩ�����    1:����;0��������
    v_ҩ������ INTEGER;--�Ƿ�ҩ�����    1:����;0��������
BEGIN
    V_���ϵ�� := 1;
    SELECT ҩƷ�շ���¼_ID.Nextval
      INTO V_lngID
      FROM Dual;
    SELECT NVL (ҩ�����, 0),nvl(ҩ������,0)
      INTO v_ҩ�����,v_ҩ������
      FROM ҩƷ���
     WHERE ҩƷID = ҩƷID_IN;

    IF v_ҩ������=0 then
        IF v_ҩ����� = 1 THEN
            BEGIN
                SELECT DISTINCT 0
                  INTO v_ҩ�����
                  FROM ��������˵��
                 WHERE (   (�������� LIKE '%ҩ��')
                          OR (�������� LIKE '�Ƽ���'))
                    AND ����ID = �ⷿID_IN;
            EXCEPTION
                WHEN OTHERS THEN
                    v_ҩ����� := 1;
            END;

            IF v_ҩ����� = 1 THEN
                V_���� := V_lngID;
            END IF;
        END IF;
    else
        V_���� := V_lngID;
    END if;

    Insert INTO ҩƷ�շ���¼
        (ID,��¼״̬,����,NO,���,�ⷿID,������ID,���ϵ��,ҩƷID,
        ����,����,����,Ч��,��д����,ʵ������,�ɱ���,�ɱ����,���ۼ�,
        ���۽��,���,ժҪ,������,��������,��������,��׼�ĺ�)
    VALUES (V_lngID,1,4,NO_IN,���_IN,�ⷿID_IN,������ID_IN,
        V_���ϵ��,ҩƷID_IN,V_����,����_IN,����_IN,Ч��_IN,
        ʵ������_IN,ʵ������_IN,�ɱ���_IN,�ɱ����_IN,���ۼ�_IN,
        ���۽��_IN,���_IN,ժҪ_IN,������_IN,��������_IN,��������_IN,��׼�ĺ�_IN
        );

Exception
   WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_ҩƷ�������_Insert;
/


CREATE OR REPLACE PROCEDURE zl_ҩƷ�������_verify (
    NO_IN IN ҩƷ�շ���¼.NO%TYPE := NULL,
    �����_IN IN ҩƷ�շ���¼.�����%TYPE := NULL
)
IS
    Err_isverified EXCEPTION;
    Err_isBatch exception;
    v_BatchCount integer;    --ԭ���������ڷ�����ҩƷ������

	V_����� ҩƷ���.ʵ�ʽ��%TYPE;
	V_����� ҩƷ���.ʵ�ʲ��%TYPE;
	V_������� ҩƷ���.ʵ������%TYPE;
	V_�ɱ���   ҩƷ���.�ϴβɹ���%TYPE;

    CURSOR C_ҩƷ�շ���¼
    IS
        SELECT ID, ʵ������, ���۽��, ���, �ⷿID, ҩƷID, ����, �ɱ���, ����,
                 Ч��, ����, ������ID,��������,��׼�ĺ�
          FROM ҩƷ�շ���¼
         WHERE NO = NO_IN
            AND ���� = 4
            AND ��¼״̬ = 1
         ORDER BY ҩƷID;
BEGIN
    UPDATE ҩƷ�շ���¼
        SET ����� = �����_IN,
             ������� = SYSDATE
     WHERE NO = NO_IN
        AND ���� = 4
        AND ��¼״̬ = 1
        AND ����� IS NULL;

    IF SQL%ROWCOUNT = 0 THEN
        RAISE Err_isverified;
    END IF;
    
    --��Ҫ���ԭ���������ڷ�����ҩƷ�����ܶ������
    SELECT COUNT(*) INTO v_BatchCount FROM        
            ҩƷ�շ���¼ a,ҩƷ��� b
    WHERE a.ҩƷid=b.ҩƷid
        AND a.no=NO_IN 
        AND a.����=4
        AND a.��¼״̬=1
        AND nvl(a.����,0)=0
        AND ((nvl(b.ҩ�����,0)=1 AND a.�ⷿid not in (select ����id from  ��������˵�� where (�������� LIKE '%ҩ��') or (�������� LIKE '�Ƽ���')))
            or nvl(b.ҩ������,0)=1);
        
    IF v_batchcount>0 THEN
        RAISE Err_isBatch;
    END IF;  
    
    --ԭ�����ֲ�������ҩƷ,�����ʱ��Ҫ������
    UPDATE ҩƷ�շ���¼ SET ����=0
    WHERE 
    id=
    (SELECT id
        FROM ҩƷ�շ���¼ a, ҩƷ��� b
        WHERE b.ҩƷid=a.ҩƷID
        AND a.no=no_in
        AND a.���� = 4 
            AND a.��¼״̬ = 1 
            AND nvl(a.����,0)>0
            AND (nvl(b.ҩ�����,0)=0 or 
                (nvl(b.ҩ������,0)=0 and a.�ⷿid in (select ����id from  ��������˵�� where (�������� LIKE '%ҩ��') or (�������� LIKE '�Ƽ���'))))
            );

    FOR V_ҩƷ�շ���¼ IN C_ҩƷ�շ���¼ LOOP
        --����ҩƷ�������Ӧ����

        UPDATE ҩƷ���
            SET �������� = NVL (��������, 0) + NVL (V_ҩƷ�շ���¼.ʵ������, 0),
                 ʵ������ = NVL (ʵ������, 0) + NVL (V_ҩƷ�շ���¼.ʵ������, 0),
                 ʵ�ʽ�� = NVL (ʵ�ʽ��, 0) + NVL (V_ҩƷ�շ���¼.���۽��, 0),
                 ʵ�ʲ�� = NVL (ʵ�ʲ��, 0) + NVL (V_ҩƷ�շ���¼.���, 0),
                 �ϴβɹ��� = NVL (V_ҩƷ�շ���¼.�ɱ���, �ϴβɹ���),
                 �ϴ����� = NVL (V_ҩƷ�շ���¼.����, �ϴ�����),
                 �ϴβ��� = NVL (V_ҩƷ�շ���¼.����, �ϴβ���),
                 Ч�� = NVL (V_ҩƷ�շ���¼.Ч��, Ч��),
                 �ϴ��������� = NVL (V_ҩƷ�շ���¼.��������, �ϴ���������),
                 ��׼�ĺ�= NVL (V_ҩƷ�շ���¼.��׼�ĺ�, ��׼�ĺ�)
         WHERE �ⷿID = V_ҩƷ�շ���¼.�ⷿID
            AND ҩƷID = V_ҩƷ�շ���¼.ҩƷID
            AND NVL (����, 0) = NVL (V_ҩƷ�շ���¼.����, 0)
            AND ���� = 1;

        IF SQL%NOTFOUND THEN
            Insert INTO ҩƷ���
                            (
                                �ⷿID,
                                ҩƷID,
                                ����,
                                ����,
                                ��������,
                                ʵ������,
                                ʵ�ʽ��,
                                ʵ�ʲ��,
                                �ϴβɹ���,
                                �ϴ�����,
                                �ϴβ���,
                                Ч��,
                                �ϴ���������,
                                ��׼�ĺ�
                            )
                  VALUES (
                      V_ҩƷ�շ���¼.�ⷿID,
                      V_ҩƷ�շ���¼.ҩƷID,
                      V_ҩƷ�շ���¼.����,
                      1,
                      V_ҩƷ�շ���¼.ʵ������,
                      V_ҩƷ�շ���¼.ʵ������,
                      V_ҩƷ�շ���¼.���۽��,
                      V_ҩƷ�շ���¼.���,
                      V_ҩƷ�շ���¼.�ɱ���,
                      V_ҩƷ�շ���¼.����,
                      V_ҩƷ�շ���¼.����,
                      V_ҩƷ�շ���¼.Ч��,
                      V_ҩƷ�շ���¼.��������,
                      V_ҩƷ�շ���¼.��׼�ĺ�
                  );
        END IF;

        DELETE
          FROM ҩƷ���
         WHERE �ⷿID = V_ҩƷ�շ���¼.�ⷿID
            AND ҩƷID = V_ҩƷ�շ���¼.ҩƷID
            AND nvl(��������,0) = 0 AND nvl(ʵ������,0) = 0 AND nvl(ʵ�ʽ��,0) = 0 AND nvl(ʵ�ʲ��,0) = 0;

        --����ҩƷ�շ����ܱ����Ӧ����

        UPDATE ҩƷ�շ�����
            SET ���� = NVL (����, 0) + NVL (V_ҩƷ�շ���¼.ʵ������, 0),
                 ��� = NVL (���, 0) + NVL (V_ҩƷ�շ���¼.���۽��, 0),
                 ��� = NVL (���, 0) + NVL (V_ҩƷ�շ���¼.���, 0)
         WHERE ���� = TRUNC (SYSDATE)
            AND �ⷿID = V_ҩƷ�շ���¼.�ⷿID
            AND ҩƷID = V_ҩƷ�շ���¼.ҩƷID
            AND ���ID = V_ҩƷ�շ���¼.������ID
            AND ���� = 4;

        IF SQL%NOTFOUND THEN
            Insert INTO ҩƷ�շ�����
                            (����, �ⷿID, ҩƷID, ���ID, ����, ����, ���, ���)
                  VALUES (
                      TRUNC (SYSDATE),
                      V_ҩƷ�շ���¼.�ⷿID,
                      V_ҩƷ�շ���¼.ҩƷID,
                      V_ҩƷ�շ���¼.������ID,
                      4,
                      V_ҩƷ�շ���¼.ʵ������,
                      V_ҩƷ�շ���¼.���۽��,
                      V_ҩƷ�շ���¼.���
                  );
        END IF;

		UPDATE ҩƷ���
		SET �ɱ���=V_ҩƷ�շ���¼.�ɱ��� 
		WHERE ҩƷID=V_ҩƷ�շ���¼.ҩƷID;
    END LOOP;
EXCEPTION
    WHEN Err_isverified THEN
        Raise_application_error (
            -20101, '[ZLSOFT]�õ����Ѿ���������ˣ�[ZLSOFT]'
        );
    when Err_isBatch then
        Raise_application_error ( 
            -20102, '[ZLSOFT]�õ����а�����һ��ԭ�������������ڷ�����ҩƷ��������ˣ�[ZLSOFT]' 
        );  
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_ҩƷ�������_verify;
/



CREATE OR REPLACE PROCEDURE zl_ҩƷ�������_strike (
    �д�_IN IN INTEGER,
    ԭ��¼״̬_IN IN ҩƷ�շ���¼.��¼״̬%TYPE,
    NO_IN IN ҩƷ�շ���¼.NO%TYPE,
    ���_IN IN ҩƷ�շ���¼.���%TYPE,
    ҩƷID_IN IN ҩƷ�շ���¼.ҩƷID%TYPE,
    ��������_IN IN ҩƷ�շ���¼.ʵ������%TYPE,
    ������_IN IN ҩƷ�շ���¼.������%TYPE,
    ��������_IN IN ҩƷ�շ���¼.��������%TYPE
)
IS
    Err_isstriked EXCEPTION;
    Err_isoutstock EXCEPTION;
    ERR_ISNONUM EXCEPTION;
    Err_isBatch EXCEPTION;
    v_BatchCount INTEGER;    --ԭ���������ڷ�����ҩƷ������

    V_�ⷿID ҩƷ�շ���¼.�ⷿID%TYPE; 
    V_������ID ҩƷ�շ���¼.������ID%TYPE ;
    V_���� ҩƷ�շ���¼.����%TYPE ; 
    V_���� ҩƷ�շ���¼.����%TYPE ; 
    V_���� ҩƷ�շ���¼.����%TYPE ; 
    V_Ч�� ҩƷ�շ���¼.Ч��%TYPE ; 
    V_�ɱ��� ҩƷ�շ���¼.�ɱ���%TYPE ; 
    V_�ɱ���� ҩƷ�շ���¼.�ɱ����%TYPE ; 
    V_���� ҩƷ�շ���¼.����%TYPE ; 
    V_���ۼ� ҩƷ�շ���¼.���ۼ�%TYPE ; 
    V_���۽�� ҩƷ�շ���¼.���۽��%TYPE ; 
    V_��� ҩƷ�շ���¼.���%TYPE ; 
    V_ժҪ ҩƷ�շ���¼.ժҪ%TYPE ; 
    V_ʣ������ ҩƷ�շ���¼.ʵ������%TYPE;
    V_ʣ��ɱ���� ҩƷ�շ���¼.�ɱ����%Type;
    V_ʣ�����۽�� ҩƷ�շ���¼.���۽��%Type; 
    V_���ϵ�� ҩƷ�շ���¼.���ϵ��%TYPE; 
    V_�������� ҩƷ�շ���¼.��������%TYPE;
    V_��׼�ĺ� ҩƷ�շ���¼.��׼�ĺ�%TYPE;
    
	  V_����� NUMBER(16,5);
	  V_����� NUMBER(16,5);
	  V_������� NUMBER(16,5);

    V_��¼�� NUMBER; 
    V_�շ�ID ҩƷ�շ���¼.ID%TYPE; 

    --�Գ����������м��
    V_����� ҩƷ���.ʵ������%TYPE;
    V_ҩ����� INTEGER;
    V_ҩ������ INTEGER;
    V_�������� INTEGER;
    V_ҩ�� INTEGER;
    V_���� NUMBER;
	INTDIGIT NUMBER;
BEGIN
	--��ȡ���С��λ��
    SELECT NVL(����ֵ,ȱʡֵ) INTO INTDIGIT FROM ϵͳ������ WHERE ������='���ý���λ��';

    IF �д�_IN =1 THEN
        UPDATE ҩƷ�շ���¼ 
        SET ��¼״̬ = DECODE(ԭ��¼״̬_IN,1,3,ԭ��¼״̬_IN+3) 
        WHERE NO = NO_IN AND ���� = 4 AND ��¼״̬ =ԭ��¼״̬_IN ; 

        IF SQL%ROWCOUNT = 0 THEN 
            RAISE ERR_ISSTRIKED; 
        END IF; 
    END IF;
    
    --��Ҫ���ԭ���������ڷ�����ҩƷ�����ܶ������
    SELECT COUNT(*) INTO v_BatchCount 
    FROM ҩƷ�շ���¼ a,ҩƷ��� b
    WHERE a.ҩƷid=b.ҩƷid
        AND a.no=NO_IN 
        AND a.����=4
        AND MOD(a.��¼״̬,3)=0
        AND a.ҩƷID+0=ҩƷID_IN
        AND nvl(a.����,0)=0
        AND ((nvl(b.ҩ�����,0)=1 AND a.�ⷿid not in (select ����id from  ��������˵�� where (�������� LIKE '%ҩ��') or (�������� LIKE '�Ƽ���')))
            or nvl(b.ҩ������,0)=1);

    IF v_batchcount>0 THEN
        RAISE Err_isBatch;
    END IF;  
    
    SELECT SUM(A.ʵ������) AS ʣ������,SUM(A.�ɱ����) AS ʣ��ɱ����,SUM(A.���۽��) AS ʣ�����۽��,A.�ⷿID,A.������ID,A.���ϵ��,Nvl(A.����,0),A.����,A.����,A.Ч��,A.�ɱ���,A.����,A.���ۼ�,A.ժҪ,B.ҩ�����,B.ҩ������,A.��������,A.��׼�ĺ�
    INTO  V_ʣ������,v_ʣ��ɱ����,V_ʣ�����۽��,V_�ⷿID,V_������ID,V_���ϵ��,V_����,V_����,V_����,V_Ч��,V_�ɱ���,V_����,V_���ۼ�,V_ժҪ,V_ҩ�����,V_ҩ������,V_��������,V_��׼�ĺ�
    FROM ҩƷ�շ���¼ A,ҩƷ��� B
    WHERE A.NO=NO_IN AND A.����=4 AND A.ҩƷID=B.ҩƷID AND A.ҩƷID=ҩƷID_IN AND A.���=���_IN
    GROUP BY A.�ⷿID,A.������ID,A.���ϵ��,NVL(A.����,0),A.����,A.����,A.Ч��,A.�ɱ���,A.����,A.���ۼ�,A.ժҪ,B.ҩ�����,B.ҩ������,A.��������,A.��׼�ĺ�;

    --�жϸò�����ҩ�⻹��ҩ��
    BEGIN
        SELECT DISTINCT 0
        INTO v_ҩ��
        FROM ��������˵��
        WHERE (   (�������� LIKE '%ҩ��')
              OR (�������� LIKE '�Ƽ���'))
        AND ����ID = V_�ⷿID;
    EXCEPTION 
        WHEN OTHERS THEN V_ҩ��:=1;
    END ;
    
    --���ݲ�������,�жϷ�������
    IF V_ҩ��=0 THEN 
        v_��������:=V_ҩ������;
    ELSE
        V_��������:=V_ҩ�����;
    END IF ;

    --V_����:(ԭ����,�ַ���Ϊ����;����Ϊ��)ԭ������,�ַ���,�����̲�����
    V_����:=0;
    IF V_��������=1 AND V_����<>0 THEN 
        V_����:=V_����;
    END IF ;
    
    --ȡ�����
    BEGIN
        SELECT Nvl(ʵ������,0) INTO V_����� FROM ҩƷ��� 
        WHERE �ⷿID=V_�ⷿID AND ҩƷID=ҩƷID_IN AND Nvl(����,0)=V_���� And ����=1;
    EXCEPTION 
        WHEN OTHERS THEN V_�����:=0;
    END ;

    --������������ʣ������,ȡʣ������;����ȡ�����
    IF V_�����<V_ʣ������ Then
       v_ʣ��ɱ����:=V_�����/V_ʣ������*v_ʣ��ɱ����;
       V_ʣ�����۽��:=V_�����/V_ʣ������*V_ʣ�����۽��;
       V_ʣ������:=V_�����; 
    END IF ;

    --������������ʣ��������������
    IF abs(V_ʣ������)<abs(��������_IN) THEN
        RAISE ERR_ISNONUM; 
    END IF;

    V_�ɱ����:= ROUND(��������_IN/v_ʣ������*v_ʣ��ɱ����,INTDIGIT);
    V_���۽��:= ROUND(��������_IN/v_ʣ������*V_ʣ�����۽��,INTDIGIT);
    V_���:=V_���۽��-V_�ɱ����;

    SELECT ҩƷ�շ���¼_ID.NEXTVAL INTO V_�շ�ID  FROM DUAL;

    Insert INTO ҩƷ�շ���¼
    (ID,��¼״̬,����,NO,���,�ⷿID,������ID,���ϵ��,ҩƷID,����,����,����,
    Ч��,��д����,ʵ������,�ɱ���,�ɱ����,���ۼ�,���۽��,���,ժҪ,������,��������,�����,�������,��������,��׼�ĺ�)
    VALUES 
    (V_�շ�ID,DECODE(ԭ��¼״̬_IN,1,2,ԭ��¼״̬_IN+2), 4, NO_IN, ���_IN, V_�ⷿID, V_������ID,
    V_���ϵ��, ҩƷID_IN,V_����,V_����, V_����, V_Ч��, -��������_IN, -��������_IN,V_�ɱ���, -V_�ɱ����, 
    V_���ۼ�, -V_���۽��, -V_���, V_ժҪ, ������_IN,��������_IN,������_IN,��������_IN,V_��������,V_��׼�ĺ�);

    --����ҩƷ�������Ӧ����

    UPDATE ҩƷ���
        SET �������� = NVL (��������, 0) - NVL (��������_IN, 0),
             ʵ������ = NVL (ʵ������, 0) - NVL (��������_IN, 0),
             ʵ�ʽ�� = NVL (ʵ�ʽ��, 0) - NVL (V_���۽��, 0),
             ʵ�ʲ�� = NVL (ʵ�ʲ��, 0) - NVL (V_���, 0),
             �ϴβɹ��� = NVL (V_�ɱ���, �ϴβɹ���),
             �ϴ����� = NVL (V_����, �ϴ�����),
             �ϴβ��� = NVL (V_����, �ϴβ���),
             Ч�� = NVL (V_Ч��, Ч��),
             �ϴ���������=NVL(V_��������,�ϴ���������),
             ��׼�ĺ�=NVL(V_��׼�ĺ�,��׼�ĺ�)
     WHERE �ⷿID = V_�ⷿID
        AND ҩƷID = ҩƷID_IN
        AND NVL (����,0) = NVL (V_����, 0)
        AND ���� = 1;

    IF SQL%NOTFOUND THEN
        Insert INTO ҩƷ���
        (�ⷿID,ҩƷID,����,����,��������,ʵ������,ʵ�ʽ��,
        ʵ�ʲ��,�ϴβɹ���,�ϴ�����,�ϴβ���,Ч��,�ϴ���������,��׼�ĺ�)
        VALUES (
        V_�ⷿID,ҩƷID_IN,V_����,1,-��������_IN,-��������_IN,-V_���۽��,
        -V_���,V_�ɱ���,V_����,V_����,V_Ч��,V_��������,V_��׼�ĺ�);
    END IF;

    DELETE
      FROM ҩƷ���
     WHERE �ⷿID = V_�ⷿID
        AND ҩƷID = ҩƷID_IN
        AND nvl(��������,0) = 0
        AND nvl(ʵ������,0) = 0
        AND nvl(ʵ�ʽ��,0) = 0
        AND nvl(ʵ�ʲ��,0) = 0;

    --����ҩƷ�շ����ܱ����Ӧ����

    UPDATE ҩƷ�շ�����
        SET ���� = NVL (����, 0) - NVL (��������_IN, 0),
             ��� = NVL (���, 0) - NVL (V_���۽��, 0),
             ��� = NVL (���, 0) - NVL (V_���, 0)
     WHERE ���� = TRUNC (��������_IN)
        AND �ⷿID = V_�ⷿID
        AND ҩƷID = ҩƷID_IN
        AND ���ID = V_������ID
        AND ���� = 4;

    IF SQL%NOTFOUND THEN
        Insert INTO ҩƷ�շ�����
        (����, �ⷿID, ҩƷID, ���ID, ����, ����, ���, ���)
        VALUES (
        TRUNC (��������_IN),V_�ⷿID,ҩƷID_IN,
        V_������ID,4,-��������_IN,-V_���۽��,-V_���);
    END IF;

EXCEPTION
    WHEN Err_isstriked THEN
        Raise_application_error (-20101, '[ZLSOFT]�õ����Ѿ������˳�����[ZLSOFT]');
    WHEN Err_isoutstock THEN
        Raise_application_error (-20102, '[ZLSOFT]�õ�������һ�ʷ���ҩƷ�ѳ��⣬���ܳ�����[ZLSOFT]');
    when Err_isBatch then
        Raise_application_error (-20102, '[ZLSOFT]�õ����а�����һ��ԭ�������������ڷ�����ҩƷ�����ܳ�����[ZLSOFT]' );  
    WHEN ERR_ISNONUM THEN 
        RAISE_APPLICATION_ERROR (-20102, '[ZLSOFT]�õ����е�' || ���_IN || '�е�ҩƷ����������������ʣ������ݣ����ܳ�����[ZLSOFT]' ); 
    WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_ҩƷ�������_strike;
/

CREATE OR REPLACE PROCEDURE zl_ҩƷ��������_Insert (
    ������ID_IN IN ҩƷ�շ���¼.������ID%TYPE,
    NO_IN IN ҩƷ�շ���¼.NO%TYPE,
    ���_IN IN ҩƷ�շ���¼.���%TYPE,
    �ⷿID_IN IN ҩƷ�շ���¼.�ⷿID%TYPE,
    ҩƷID_IN IN ҩƷ�շ���¼.ҩƷID%TYPE,
    ����_IN IN ҩƷ�շ���¼.����%TYPE,
    ��д����_IN IN ҩƷ�շ���¼.��д����%TYPE,
    �ɱ���_IN IN ҩƷ�շ���¼.�ɱ���%TYPE,
    �ɱ����_IN IN ҩƷ�շ���¼.�ɱ����%TYPE,
    ���ۼ�_IN IN ҩƷ�շ���¼.���ۼ�%TYPE,
    ���۽��_IN IN ҩƷ�շ���¼.���۽��%TYPE,
    ���_IN IN ҩƷ�շ���¼.���%TYPE,
	�����_IN IN ҩƷ�շ���¼.����%TYPE,
	�����λ_IN IN ҩƷ�շ���¼.��ҩ����%TYPE,
    ������_IN IN ҩƷ�շ���¼.������%TYPE,
    ��������_IN IN ҩƷ�շ���¼.��������%TYPE,
    ����_IN IN ҩƷ�շ���¼.����%TYPE := NULL,
    ����_IN IN ҩƷ�շ���¼.����%TYPE := NULL,
    Ч��_IN IN ҩƷ�շ���¼.Ч��%TYPE := NULL,
    ժҪ_IN IN ҩƷ�շ���¼.ժҪ%TYPE := NULL,
    ��׼�ĺ�_IN In ҩƷ�շ���¼.��׼�ĺ�%TYPE := Null
)
Is
    ERR_MutilROW EXCEPTION ;
    Err_isNOnumber EXCEPTION;
    intRecords NUMBER ;
    V_���� �շ���ĿĿ¼.����%TYPE;
    V_�������� ҩƷ���.��������%TYPE;
    V_���ϵ�� ҩƷ�շ���¼.���ϵ��%TYPE;--�շ�ID
BEGIN
    V_���ϵ�� := -1;

    IF ����_IN > 0 THEN
        BEGIN
            SELECT ��������
              INTO V_��������
              FROM ҩƷ���
             WHERE ҩƷID = ҩƷID_IN
                AND NVL (����, 0) = ����_IN
                AND �ⷿID = �ⷿID_IN
                AND ���� = 1
                AND ROWNUM = 1;
        EXCEPTION
            WHEN OTHERS THEN
                V_�������� := 0;
        END;

        IF V_�������� - ��д����_IN < 0 THEN
            RAISE Err_isNOnumber;
        END IF;
    END IF;

    Insert INTO ҩƷ�շ���¼
    (ID,��¼״̬,����,NO,���,�ⷿID,������ID,���ϵ��,ҩƷID,����,����,����,Ч��,
    ��д����,ʵ������,�ɱ���,�ɱ����,���ۼ�,���۽��,���,ժҪ,������,��������,����,��ҩ����,��׼�ĺ�)
	VALUES (
	ҩƷ�շ���¼_ID.Nextval,1,11,NO_IN,���_IN,�ⷿID_IN,������ID_IN,V_���ϵ��,ҩƷID_IN,����_IN,
	����_IN,����_IN,Ч��_IN,��д����_IN,��д����_IN,�ɱ���_IN,�ɱ����_IN,���ۼ�_IN,���۽��_IN,���_IN,
	ժҪ_IN,������_IN,��������_IN,�����_IN,�����λ_IN,��׼�ĺ�_IN);
  
  --����Ƿ������ͬҩƷ��ͬ���ε����ݣ�������ڲ�������
	SELECT COUNT(*) INTO intRecords
	FROM ҩƷ�շ���¼
	WHERE ����=11 AND NO=NO_IN AND ���ϵ��=-1 AND ҩƷID+0=ҩƷID_IN AND Nvl(����,0)=NVL(����_IN,0);
	IF intRecords>1 THEN
		RAISE ERR_MutilROW;
	END IF ;
  
    --ͬʱ���¿����
	UPDATE ҩƷ���
	SET �������� = NVL (��������, 0) - ��д����_IN
	WHERE �ⷿID = �ⷿID_IN
	AND ҩƷID = ҩƷID_IN
	AND NVL (����, 0) = NVL (����_IN, 0)
	AND ���� = 1;

    --��������������Ϊ����ҩƷ��������׼����
    IF SQL%NOTFOUND THEN
        Insert INTO ҩƷ���
		(�ⷿID, ҩƷID, ����, ��������,�ϴ�����,Ч��,�ϴβ���,��׼�ĺ�)
		VALUES 
		(�ⷿID_IN, ҩƷID_IN, 1, -��д����_IN,����_IN,Ч��_IN,����_IN,��׼�ĺ�_IN);
    END IF;

    DELETE
      FROM ҩƷ���
     WHERE �ⷿID = �ⷿID_IN
        AND ҩƷID = ҩƷID_IN
        AND NVL (��������, 0) = 0
        AND NVL (ʵ������, 0) = 0
        AND NVL (ʵ�ʽ��, 0) = 0
        AND NVL (ʵ�ʲ��, 0) = 0;
Exception
    WHEN ERR_MutilROW THEN
		SELECT ���� INTO V_���� FROM �շ���ĿĿ¼ WHERE ID = ҩƷID_IN;
		RAISE_APPLICATION_ERROR(-20101,'[ZLSOFT]����Ϊ'||V_����||'��ҩƷ�����ڶ����ظ��ļ�¼����ϲ�Ϊһ����¼��[ZLSOFT]');
    
    WHEN Err_isNOnumber THEN
        SELECT ����
          INTO V_����
          FROM �շ���ĿĿ¼
         WHERE ID = ҩƷID_IN;
        Raise_application_error (
            -20101, '[ZLSOFT]����Ϊ' || V_���� || ',����Ϊ' || ����_IN ||
                          '��ҩ�����ҩƷ' ||
                          CHR (10) ||
                          CHR (13) ||
                          '���ÿ������������[ZLSOFT]'
        );
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_ҩƷ��������_Insert;
/


CREATE OR REPLACE PROCEDURE zl_ҩƷ��������_DELETE (
    
    --ɾ��ҩƷ�շ���¼���ָ���Ӧ�ı�ҩƷ���
    NO_IN IN ҩƷ�շ���¼.NO%TYPE
)
IS
    Err_isverified EXCEPTION;

    CURSOR C_ҩƷ�շ���¼
    IS
        SELECT ��д����, �ⷿID, ����, ҩƷID,����,Ч��,����,��׼�ĺ�
          FROM ҩƷ�շ���¼
         WHERE NO = NO_IN
            AND ���� = 11
         ORDER BY ҩƷID;
BEGIN
    --ͨ��ѭ�����ָ�ԭ���Ŀ�������
    FOR V_ҩƷ�շ���¼ IN C_ҩƷ�շ���¼ LOOP
        UPDATE ҩƷ���
            SET �������� = NVL (��������, 0) + V_ҩƷ�շ���¼.��д����
         WHERE �ⷿID = V_ҩƷ�շ���¼.�ⷿID
            AND ҩƷID = V_ҩƷ�շ���¼.ҩƷID
            AND NVL (����, 0) = NVL (V_ҩƷ�շ���¼.����, 0)
            AND ���� = 1;

        IF SQL%NOTFOUND THEN
            Insert INTO ҩƷ���
                            (�ⷿID, ҩƷID, ����, ����, ��������,�ϴ�����,Ч��,�ϴβ���,��׼�ĺ�)
                  VALUES (
                      V_ҩƷ�շ���¼.�ⷿID,V_ҩƷ�շ���¼.ҩƷID,V_ҩƷ�շ���¼.����,1,
                      V_ҩƷ�շ���¼.��д����,V_ҩƷ�շ���¼.����,V_ҩƷ�շ���¼.Ч��,V_ҩƷ�շ���¼.����,V_ҩƷ�շ���¼.��׼�ĺ�
                  );
        END IF;

        DELETE
          FROM ҩƷ���
         WHERE �ⷿID = V_ҩƷ�շ���¼.�ⷿID
            AND ҩƷID = V_ҩƷ�շ���¼.ҩƷID
            AND NVL (��������, 0) = 0
            AND NVL (ʵ������, 0) = 0
            AND NVL (ʵ�ʽ��, 0) = 0
            AND NVL (ʵ�ʲ��, 0) = 0;
    END LOOP;

    DELETE
      FROM ҩƷ�շ���¼
     WHERE NO = NO_IN
        AND ���� = 11
        AND ��¼״̬ = 1
        AND ����� IS NULL;

    IF SQL%ROWCOUNT = 0 THEN
        RAISE Err_isverified;
    END IF;
EXCEPTION
    WHEN Err_isverified THEN
        Raise_application_error (
            -20101, '[ZLSOFT]�õ����Ѿ�������ɾ�����ѱ�����ˣ�[ZLSOFT]'
        );
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_ҩƷ��������_DELETE;
/


CREATE OR REPLACE PROCEDURE zl_ҩƷ��������_verify (
    ���_IN IN ҩƷ�շ���¼.���%TYPE,
    NO_IN IN ҩƷ�շ���¼.NO%TYPE,
    �ⷿID_IN IN ҩƷ�շ���¼.�ⷿID%TYPE,
    ҩƷID_IN IN ҩƷ�շ���¼.ҩƷID%TYPE,
    ����_IN IN ҩƷ�շ���¼.����%TYPE,
    ʵ������_IN IN ҩƷ�շ���¼.ʵ������%TYPE,
    �ɱ���_IN IN ҩƷ�շ���¼.�ɱ���%TYPE,
    �ɱ����_IN IN ҩƷ�շ���¼.�ɱ����%TYPE,
    ���۽��_IN IN ҩƷ�շ���¼.���۽��%TYPE,
    ���_IN IN ҩƷ�շ���¼.���%TYPE,
    ������ID_IN IN ҩƷ�շ���¼.������ID%TYPE,
    �����_IN IN ҩƷ�շ���¼.�����%TYPE,
    �������_IN IN ҩƷ�շ���¼.�������%Type
)
IS
    Err_isverified EXCEPTION;
    V_ʵ�ʿ���� ҩƷ���.ʵ�ʽ��%TYPE;
    V_ʵ�ʿ���� ҩƷ���.ʵ�ʲ��%TYPE;
    V_����� number(18,8);
    V_������ ҩƷ���.ʵ�ʲ��%TYPE;
    V_�ɱ��� ҩƷ�շ���¼.�ɱ���%TYPE;
    V_�ɱ���� ҩƷ�շ���¼.�ɱ����%TYPE;
	v_���� ҩƷ�շ���¼.����%TYPE;
	v_Ч�� ҩƷ�շ���¼.Ч��%TYPE;
	v_���� ҩƷ�շ���¼.����%TYPE;
  v_��׼�ĺ� ҩƷ�շ���¼.��׼�ĺ�%TYPE;
	INTDIGIT NUMBER ;
BEGIN
	--��ȡ���С��λ��
    SELECT NVL(����ֵ,ȱʡֵ) INTO INTDIGIT FROM ϵͳ������ WHERE ������='���ý���λ��';

    --�������ô������������ʱ�ı�ʵ��������
      --�������ȶ�ʵ��������������Ӧ���ֶν��и��¡�
    BEGIN
        SELECT nvl(ʵ�ʽ��,0), nvl(ʵ�ʲ��,0)
          INTO V_ʵ�ʿ����, V_ʵ�ʿ����
          FROM ҩƷ���
         WHERE ҩƷID = ҩƷID_IN
            AND NVL (����, 0) = ����_IN
            AND �ⷿID = �ⷿID_IN
            AND ���� = 1
            AND ROWNUM = 1;
    EXCEPTION
        WHEN OTHERS THEN
            V_ʵ�ʿ���� := 0;
    END;

    IF V_ʵ�ʿ���� <= 0 THEN
        BEGIN
            SELECT ָ������� / 100
              INTO V_�����
              FROM ҩƷ���
             WHERE ҩƷID = ҩƷID_IN;
        EXCEPTION
            WHEN OTHERS THEN
                V_����� := 0;
        END;
    ELSE
        V_����� := V_ʵ�ʿ���� / V_ʵ�ʿ����;
    END IF;

    V_������ := round(���۽��_IN * V_�����,INTDIGIT);
    IF ʵ������_IN<=0 THEN 
        V_�ɱ��� := �ɱ���_IN;
    ELSE 
        V_�ɱ��� := (���۽��_IN - V_������) / ʵ������_IN;
    END IF ;
    V_�ɱ���� := round(V_�ɱ��� * ʵ������_IN,INTDIGIT);
	
	--��ȡҩƷ�������ⵥָ����ϸ������,Ч���������Ϣ
	SELECT ����,Ч��,����,��׼�ĺ� INTO v_����,v_Ч��,v_����,v_��׼�ĺ�
	FROM ҩƷ�շ���¼
	WHERE ����=11 AND NO=NO_IN AND ���=���_IN;

    UPDATE ҩƷ�շ���¼
        SET ����� = NVL (�����_IN, �����),
             ������� = �������_IN,
             �ɱ��� = V_�ɱ���,
             �ɱ���� = V_�ɱ����,
             ��� = V_������
     WHERE NO = NO_IN
        AND ���� = 11
        AND ҩƷID = ҩƷID_IN
        AND ��� = ���_IN
        AND ��¼״̬ = 1
        AND ����� IS NULL;

    IF SQL%ROWCOUNT = 0 THEN
        RAISE Err_isverified;
    END IF;

    --����ҩƷ������Ӧ����

    UPDATE ҩƷ���
        SET ʵ������ = NVL (ʵ������, 0) - ʵ������_IN,
             ʵ�ʽ�� = NVL (ʵ�ʽ��, 0) - ���۽��_IN,
             ʵ�ʲ�� = NVL (ʵ�ʲ��, 0) - V_������
     WHERE �ⷿID = �ⷿID_IN
        AND ҩƷID = ҩƷID_IN
        AND NVL (����, 0) = NVL (����_IN, 0)
        AND ���� = 1;

    IF SQL%NOTFOUND THEN
        Insert INTO ҩƷ���
                        (
                            �ⷿID,
                            ҩƷID,
                            ����,
                            ����,
                            ��������,
                            ʵ������,
                            ʵ�ʽ��,
                            ʵ�ʲ��,
							�ϴ�����,Ч��,�ϴβ���,��׼�ĺ�
                        )
              VALUES (
                  �ⷿID_IN,
                  ҩƷID_IN,
                  ����_IN,
                  1,
                  -ʵ������_IN,
                  -ʵ������_IN,
                  -���۽��_IN,
                  -V_������,
				  v_����,v_Ч��,v_����,v_��׼�ĺ�
              );
    END IF;

    DELETE
      FROM ҩƷ���
     WHERE �ⷿID = �ⷿID_IN
        AND ҩƷID = ҩƷID_IN
        AND NVL (��������, 0) = 0
        AND NVL (ʵ������, 0) = 0
        AND NVL (ʵ�ʽ��, 0) = 0
        AND NVL (ʵ�ʲ��, 0) = 0;

    --��ҩƷ�շ����ܱ����Ӧ����
    UPDATE ҩƷ�շ�����
        SET ���� = NVL (����, 0) - ʵ������_IN,
             ��� = NVL (���, 0) - ���۽��_IN,
             ��� = NVL (���, 0) - V_������
     WHERE ���� = TRUNC (SYSDATE)
        AND �ⷿID = �ⷿID_IN
        AND ҩƷID = ҩƷID_IN
        AND ���ID = ������ID_IN
        AND ���� = 11;

    IF SQL%NOTFOUND THEN
        Insert INTO ҩƷ�շ�����
                        (����, �ⷿID, ҩƷID, ���ID, ����, ����, ���, ���)
              VALUES (
                  TRUNC (SYSDATE),
                  �ⷿID_IN,
                  ҩƷID_IN,
                  ������ID_IN,
                  11,
                  -ʵ������_IN,
                  -���۽��_IN,
                  -V_������
              );
    END IF;
EXCEPTION
    WHEN Err_isverified THEN
        Raise_application_error (
            -20101, '[ZLSOFT]�õ����Ѿ���������ˣ�[ZLSOFT]'
        );
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_ҩƷ��������_verify;
/


CREATE OR REPLACE PROCEDURE zl_ҩƷ��������_strike (
    �д�_IN IN INTEGER,
    ԭ��¼״̬_IN IN ҩƷ�շ���¼.��¼״̬%TYPE,
    NO_IN IN ҩƷ�շ���¼.NO%TYPE,
    ���_IN IN ҩƷ�շ���¼.���%TYPE,
    ҩƷID_IN IN ҩƷ�շ���¼.ҩƷID%TYPE,
    ��������_IN IN ҩƷ�շ���¼.ʵ������%TYPE,
    ������_IN IN ҩƷ�շ���¼.������%TYPE,
    ��������_IN IN ҩƷ�շ���¼.��������%TYPE
)
IS
    Err_isstriked EXCEPTION;
    Err_isoutstock EXCEPTION;
    ERR_ISNONUM EXCEPTION;
    Err_isBatch EXCEPTION;
    v_BatchCount INTEGER;    --ԭ���������ڷ�����ҩƷ������

    V_�ⷿID ҩƷ�շ���¼.�ⷿID%TYPE; 
    V_������ID ҩƷ�շ���¼.������ID%TYPE ;
    V_���� ҩƷ�շ���¼.����%TYPE ; 
    V_���� ҩƷ�շ���¼.����%TYPE ; 
    V_���� ҩƷ�շ���¼.����%TYPE ; 
    V_Ч�� ҩƷ�շ���¼.Ч��%TYPE ; 
    V_�ɱ��� ҩƷ�շ���¼.�ɱ���%TYPE ; 
    V_�ɱ���� ҩƷ�շ���¼.�ɱ����%TYPE ; 
    V_���� ҩƷ�շ���¼.����%TYPE ; 
    V_���ۼ� ҩƷ�շ���¼.���ۼ�%TYPE ; 
    V_���۽�� ҩƷ�շ���¼.���۽��%TYPE ; 
    V_��� ҩƷ�շ���¼.���%TYPE ; 
    V_ժҪ ҩƷ�շ���¼.ժҪ%TYPE ; 
    V_ʣ������ ҩƷ�շ���¼.ʵ������%TYPE;
    V_ʣ��ɱ���� ҩƷ�շ���¼.�ɱ����%Type;
    V_ʣ�����۽�� ҩƷ�շ���¼.���۽��%Type; 
    V_���ϵ�� ҩƷ�շ���¼.���ϵ��%TYPE; 
	V_����� ҩƷ�շ���¼.����%TYPE;
	V_�����λ ҩƷ�շ���¼.��ҩ����%TYPE;
    v_��׼�ĺ� ҩƷ�շ���¼.��׼�ĺ�%TYPE;

    V_��¼�� NUMBER; 
    V_�շ�ID ҩƷ�շ���¼.ID%TYPE; 
	INTDIGIT NUMBER;
BEGIN
	--��ȡ���С��λ��
    SELECT NVL(����ֵ,ȱʡֵ) INTO INTDIGIT FROM ϵͳ������ WHERE ������='���ý���λ��';

    IF �д�_IN =1 THEN
        UPDATE ҩƷ�շ���¼ 
        SET ��¼״̬ = DECODE(ԭ��¼״̬_IN,1,3,ԭ��¼״̬_IN+3) 
        WHERE NO = NO_IN AND ���� = 11 AND ��¼״̬ =ԭ��¼״̬_IN ; 
        IF SQL%ROWCOUNT = 0 THEN 
            RAISE ERR_ISSTRIKED; 
        END IF; 
    END IF;
    
    --��Ҫ���ԭ���������ڷ�����ҩƷ�����ܶ������
    SELECT COUNT(*) INTO v_BatchCount 
    FROM ҩƷ�շ���¼ a,ҩƷ��� b
    WHERE a.ҩƷid=b.ҩƷid
        AND a.no=NO_IN 
        AND a.����=11
        AND A.ҩƷID+0=ҩƷID_IN
        AND MOD(a.��¼״̬,3)=0
        AND nvl(a.����,0)=0
        AND ((NVL(B.ҩ�����,0)=1 AND A.�ⷿID NOT IN (SELECT ����ID FROM  ��������˵�� WHERE (�������� LIKE '%ҩ��') OR (�������� LIKE '�Ƽ���')))
            OR NVL(B.ҩ������,0)=1);
    
    IF v_batchcount>0 THEN
        RAISE Err_isBatch;
    END IF;  
    
    SELECT SUM(ʵ������) AS ʣ������,SUM(�ɱ����) AS ʣ��ɱ����,SUM(���۽��) AS ʣ�����۽��,�ⷿID,������ID,���ϵ��,����,����,����,Ч��,�ɱ���,����,���ۼ�,ժҪ,����,��ҩ����,��׼�ĺ�
    INTO  V_ʣ������,V_ʣ��ɱ����,V_ʣ�����۽��,V_�ⷿID,V_������ID,V_���ϵ��,V_����,V_����,V_����,V_Ч��,V_�ɱ���,V_����,V_���ۼ�,V_ժҪ,V_�����,V_�����λ,v_��׼�ĺ�
    FROM ҩƷ�շ���¼ 
    WHERE NO=NO_IN 
    AND ����=11 
    AND ҩƷID=ҩƷID_IN 
    AND ���=���_IN
    GROUP BY �ⷿID,������ID,���ϵ��,����,����,����,Ч��,�ɱ���,����,���ۼ�,ժҪ,����,��ҩ����,��׼�ĺ�;

    --������������ʣ��������������
    IF abs(V_ʣ������)<abs(��������_IN) THEN
        RAISE ERR_ISNONUM; 
    END IF;

    V_�ɱ����:= ROUND(��������_IN/v_ʣ������*V_ʣ��ɱ����,INTDIGIT);
    V_���۽��:= ROUND(��������_IN/v_ʣ������*V_ʣ�����۽��,INTDIGIT);
    V_���:=V_���۽��-V_�ɱ����;

    SELECT ҩƷ�շ���¼_ID.NEXTVAL INTO V_�շ�ID  FROM DUAL;
    Insert INTO ҩƷ�շ���¼
    (ID,��¼״̬,����,NO,���,�ⷿID,������ID,���ϵ��,
    ҩƷID,����,����,����,Ч��,��д����,ʵ������,�ɱ���,
    �ɱ����,���ۼ�,���۽��,���,����,��ҩ����,
	ժҪ,������,��������,�����,�������,��׼�ĺ�)
    VALUES 
    (V_�շ�ID,DECODE(ԭ��¼״̬_IN,1,2,ԭ��¼״̬_IN+2),11,NO_IN,���_IN,V_�ⷿID,V_������ID,V_���ϵ��,
    ҩƷID_IN,V_����,V_����,V_����,V_Ч��,-��������_IN,-��������_IN,V_�ɱ���,-V_�ɱ����,V_���ۼ�,-V_���۽��,
    -V_���,V_�����,V_�����λ,V_ժҪ,������_IN,��������_IN,������_IN,��������_IN,v_��׼�ĺ�);

    --ԭ�����ֲ�������ҩƷ,�ڳ���ʱ��Ҫ������
    BEGIN 
        SELECT COUNT(*) INTO V_��¼��
        FROM ҩƷ�շ���¼ A, ҩƷ��� B
        WHERE A.ҩƷID+0=B.ҩƷID 
        AND B.ҩƷID=ҩƷID_IN
        AND A.NO=NO_IN
        AND A.���� = 11 
        AND MOD(A.��¼״̬,3)=0
        AND NVL(A.����,0)>0
        AND (NVL(B.ҩ�����,0)=0 OR 
            (NVL(B.ҩ������,0)=0 AND A.�ⷿID IN (SELECT ����ID FROM  ��������˵�� WHERE (�������� LIKE '%ҩ��') OR (�������� LIKE '�Ƽ���'))));
    EXCEPTION 
        WHEN OTHERS THEN
            V_��¼��:=0;
    END;
    IF V_��¼��>0 THEN
        V_����:=0;
    ELSE
        V_����:=NVL (V_����, 0);
    END IF;
    --����ҩƷ�������Ӧ����
    UPDATE ҩƷ���
        SET �������� = NVL (��������, 0) + NVL (��������_IN, 0),
             ʵ������ = NVL (ʵ������, 0) + NVL (��������_IN, 0),
             ʵ�ʽ�� = NVL (ʵ�ʽ��, 0) + NVL (V_���۽��, 0),
             ʵ�ʲ�� = NVL (ʵ�ʲ��, 0) + NVL (V_���, 0)
     WHERE �ⷿID = V_�ⷿID
        AND ҩƷID = ҩƷID_IN
        AND NVL (����, 0) = v_����
        AND ���� = 1;

    IF SQL%NOTFOUND THEN
        Insert INTO ҩƷ���
        (�ⷿID,ҩƷID,����,����,��������,ʵ������,ʵ�ʽ��,ʵ�ʲ��,�ϴ�����,Ч��,�ϴβ���,��׼�ĺ�)
        VALUES 
        (V_�ⷿID,ҩƷID_IN,V_����,1,��������_IN,��������_IN,V_���۽��,V_���,V_����,V_Ч��,v_����,v_��׼�ĺ�);
    END IF;

    DELETE
      FROM ҩƷ���
     WHERE �ⷿID = V_�ⷿID
        AND ҩƷID = ҩƷID_IN
        AND NVL (��������, 0) = 0
        AND NVL (ʵ������, 0) = 0
        AND NVL (ʵ�ʽ��, 0) = 0
        AND NVL (ʵ�ʲ��, 0) = 0;

    --����ҩƷ�շ����ܱ����Ӧ����

    UPDATE ҩƷ�շ�����
        SET ���� = NVL (����, 0) + NVL (��������_IN, 0),
             ��� = NVL (���, 0) + NVL (V_���۽��, 0),
             ��� = NVL (���, 0) + NVL (V_���, 0)
     WHERE ���� = TRUNC (��������_IN)
        AND �ⷿID = V_�ⷿID
        AND ҩƷID = ҩƷID_IN
        AND ���ID = V_������ID
        AND ���� = 11;

    IF SQL%NOTFOUND THEN
        Insert INTO ҩƷ�շ�����
        (����, �ⷿID, ҩƷID, ���ID, ����, ����, ���, ���)
        VALUES 
        (TRUNC (��������_IN),V_�ⷿID,ҩƷID_IN,V_������ID,11,��������_IN,V_���۽��,V_���);
    END IF;
EXCEPTION
    WHEN Err_isstriked THEN
        Raise_application_error (-20101, '[ZLSOFT]�õ����Ѿ������˳�����[ZLSOFT]');
    WHEN Err_isBatch THEN
        Raise_application_error (-20102, '[ZLSOFT]�õ����а�����һ��ԭ�������������ڷ�����ҩƷ�����ܳ�����[ZLSOFT]'); 
    WHEN OTHERS THEN zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_ҩƷ��������_strike;
/

CREATE OR REPLACE PROCEDURE zl_ҩƷ�̵�_Insert (
    NO_IN IN ҩƷ�շ���¼.NO%TYPE,
    ���_IN IN ҩƷ�շ���¼.���%TYPE,
    �ⷿID_IN IN ҩƷ�շ���¼.�ⷿID%TYPE,
    ����_IN IN ҩƷ�շ���¼.����%TYPE,
    ������ID_IN IN ҩƷ�շ���¼.������ID%TYPE,
    --ÿ�ζ�ʹ�ã����������洫��ȽϺ�
    ���ϵ��_IN IN ҩƷ�շ���¼.���ϵ��%TYPE,
    ҩƷID_IN IN ҩƷ�շ���¼.ҩƷID%TYPE,
    ��������_IN IN ҩƷ�շ���¼.��д����%TYPE,
    ʵ������_IN IN ҩƷ�շ���¼.����%TYPE,
    ������_IN IN ҩƷ�շ���¼.ʵ������%TYPE,
    �ۼ�_IN IN ҩƷ�շ���¼.���ۼ�%TYPE,
    ����_IN IN ҩƷ�շ���¼.���۽��%TYPE,
    ��۲�_IN IN ҩƷ�շ���¼.���%TYPE,
    ������_IN IN ҩƷ�շ���¼.������%TYPE,
    ��������_IN IN ҩƷ�շ���¼.��������%TYPE,
    ժҪ_IN IN ҩƷ�շ���¼.ժҪ%TYPE := NULL,
    ����_IN IN ҩƷ�շ���¼.����%TYPE := NULL,
    ����_IN IN ҩƷ�շ���¼.����%TYPE := NULL,
    Ч��_IN IN ҩƷ�շ���¼.Ч��%TYPE := NULL,
    �̵�ʱ��_IN IN ҩƷ�շ���¼.Ƶ��%TYPE := NULL,
    �����_IN IN ҩƷ�շ���¼.�ɱ���%TYPE := NULL,
    �����_IN IN ҩƷ�շ���¼.�ɱ����%TYPE := NULL,
    �ɱ���_IN In ҩƷ�շ���¼.����%TYPE := NULL,
    ��׼�ĺ�_IN In ҩƷ�շ���¼.��׼�ĺ�%TYPE := Null
)
IS
    V_lngID NUMBER(18);
BEGIN
    --�������_INΪ-1,���ʾ�²���һ������ҩƷ
    SELECT ҩƷ�շ���¼_ID.Nextval
    INTO V_lngID
    FROM Dual;

    Insert INTO ҩƷ�շ���¼
        (ID,��¼״̬,����,NO,���,�ⷿID,������ID,���ϵ��,ҩƷID,����,
        ����,����,Ч��,��д����,����,ʵ������,���ۼ�,���۽��,���,
        ժҪ,������,��������,Ƶ��,�ɱ���,�ɱ����,����,��׼�ĺ�)
    VALUES (
        V_lngID,1,12,NO_IN,���_IN,�ⷿID_IN,������ID_IN,���ϵ��_IN,ҩƷID_IN,
        DECODE(����_IN,-1,V_lngID,����_IN),����_IN,����_IN,Ч��_IN,��������_IN,
        ʵ������_IN,������_IN,�ۼ�_IN,����_IN,��۲�_IN,ժҪ_IN,������_IN,
        ��������_IN,�̵�ʱ��_IN,�����_IN,�����_IN,�ɱ���_IN,��׼�ĺ�_IN);

    IF ���ϵ��_IN = -1 THEN
        --ͬʱ���¿����
        UPDATE ҩƷ���
        SET �������� = NVL (��������, 0) - ������_IN
        WHERE �ⷿID = �ⷿID_IN
        AND ҩƷID = ҩƷID_IN
        AND NVL (����, 0) = NVL (����_IN, 0)
        AND ���� = 1;

        IF SQL%NOTFOUND THEN
            Insert INTO ҩƷ���
            (�ⷿID, ҩƷID, ����, ����, ��������,�ϴ�����,Ч��,�ϴβ���,��׼�ĺ�)
            VALUES 
            (�ⷿID_IN, ҩƷID_IN, 1, ����_IN, -������_IN,����_IN,Ч��_IN,����_IN,��׼�ĺ�_IN);
        END IF;

        DELETE
        FROM ҩƷ���
        WHERE �ⷿID = �ⷿID_IN
        AND ҩƷID = ҩƷID_IN
        AND NVL (��������, 0) = 0
        AND NVL (ʵ������, 0) = 0
        AND NVL (ʵ�ʽ��, 0) = 0
        AND NVL (ʵ�ʲ��, 0) = 0;
    END IF;
EXCEPTION
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_ҩƷ�̵�_Insert;
/




CREATE OR REPLACE PROCEDURE zl_ҩƷ�̵�_DELETE (
    
    --ɾ��ҩƷ�շ���¼���ָ���Ӧ�ı�ҩƷ���
    NO_IN IN ҩƷ�շ���¼.NO%TYPE
)
IS
    Err_isverified EXCEPTION;

    CURSOR C_ҩƷ�շ���¼
    IS
        SELECT ʵ������, �ⷿID, ����, ҩƷID,����,Ч��,����,��׼�ĺ�
          FROM ҩƷ�շ���¼
         WHERE NO = NO_IN
            AND ���� = 12
            AND ���ϵ�� = -1
         ORDER BY ҩƷID;
BEGIN
    --ͨ��ѭ�����ָ��������ԭ���Ŀ���������
     --ʵ�������������������
    FOR V_ҩƷ�շ���¼ IN C_ҩƷ�շ���¼ LOOP
        UPDATE ҩƷ���
            SET �������� = NVL (��������, 0) + V_ҩƷ�շ���¼.ʵ������
         WHERE �ⷿID = V_ҩƷ�շ���¼.�ⷿID
            AND ҩƷID = V_ҩƷ�շ���¼.ҩƷID
            AND NVL (����, 0) = NVL (V_ҩƷ�շ���¼.����, 0)
            AND ���� = 1;

        IF SQL%NOTFOUND THEN
            Insert INTO ҩƷ���
                            (�ⷿID, ҩƷID, ����, ����, ��������,�ϴ�����,Ч��,�ϴβ���,��׼�ĺ�)
                  VALUES (
                      V_ҩƷ�շ���¼.�ⷿID,V_ҩƷ�շ���¼.ҩƷID,V_ҩƷ�շ���¼.����,1,
                      V_ҩƷ�շ���¼.ʵ������,V_ҩƷ�շ���¼.����,V_ҩƷ�շ���¼.Ч��,V_ҩƷ�շ���¼.����,V_ҩƷ�շ���¼.��׼�ĺ�
                  );
        END IF;

        DELETE
          FROM ҩƷ���
         WHERE �ⷿID = V_ҩƷ�շ���¼.�ⷿID
            AND ҩƷID = V_ҩƷ�շ���¼.ҩƷID
            AND NVL (��������, 0) = 0
            AND NVL (ʵ������, 0) = 0
            AND NVL (ʵ�ʽ��, 0) = 0
            AND NVL (ʵ�ʲ��, 0) = 0;
    END LOOP;

    DELETE
      FROM ҩƷ�շ���¼
     WHERE NO = NO_IN
        AND ���� = 12
        AND ��¼״̬ = 1
        AND ����� IS NULL;

    IF SQL%ROWCOUNT = 0 THEN
        RAISE Err_isverified;
    END IF;
EXCEPTION
    WHEN Err_isverified THEN
        Raise_application_error (
            -20101, '[ZLSOFT]�õ����Ѿ�������ɾ�����ѱ�����ˣ�[ZLSOFT]'
        );
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_ҩƷ�̵�_DELETE;
/


CREATE OR REPLACE PROCEDURE zl_ҩƷ�̵�_verify (
    NO_IN IN ҩƷ�շ���¼.NO%TYPE := NULL,
    �����_IN IN ҩƷ�շ���¼.�����%TYPE := NULL
)
IS
    Err_isverified EXCEPTION;
    Err_isBatch exception;
    v_BatchCount integer;    --ԭ���������ڷ�����ҩƷ������

    CURSOR C_ҩƷ�շ���¼
    IS
        SELECT ID, ʵ������, ���۽��, ���, �ⷿID, ҩƷID, ����, ����, Ч��,
                 ����, ������ID, ���ϵ��,��׼�ĺ�
          FROM ҩƷ�շ���¼
         WHERE NO = NO_IN
            AND ���� = 12
            AND ��¼״̬ = 1
         ORDER BY ҩƷID;
BEGIN
    UPDATE ҩƷ�շ���¼
        SET ����� = �����_IN,
             ������� = SYSDATE
     WHERE NO = NO_IN
        AND ���� = 12
        AND ��¼״̬ = 1
        AND ����� IS NULL;

    IF SQL%ROWCOUNT = 0 THEN
        RAISE Err_isverified;
    END IF;

    --��Ҫ���ԭ���������ڷ�����ҩƷ�����ܶ������
    SELECT count(*) INTO v_BatchCount FROM        
            ҩƷ�շ���¼ a,ҩƷ��� b
    WHERE a.ҩƷid=b.ҩƷid
        AND a.no=NO_IN 
        AND a.����=12
        AND a.��¼״̬=1
        AND nvl(a.����,0)=0
        AND nvl(b.ҩ�����,0)=1
        AND a.�ⷿid not in
        (select ����id from  ��������˵�� where (�������� LIKE '%ҩ��') or (�������� LIKE '�Ƽ���'));
    IF v_batchcount>0 THEN
        raise Err_isBatch;
    END IF;  
    
    --ԭ�����ֲ�������ҩƷ,�����ʱ��Ҫ������
    UPDATE ҩƷ�շ���¼ SET ����=0
    WHERE 
    id=
    (SELECT id
        FROM ҩƷ�շ���¼ a, ҩƷ��� b
        WHERE b.ҩƷid=a.ҩƷID
        AND a.no=no_in
        AND a.���� = 12 
            AND a.��¼״̬ = 1 
            AND nvl(a.����,0)>0
            AND nvl(b.ҩ�����,0)=0
            );
    
    FOR V_ҩƷ�շ���¼ IN C_ҩƷ�շ���¼ LOOP
        --����ҩƷ�������Ӧ����
        UPDATE ҩƷ���
            SET �������� =
                     NVL (��������, 0) +
                         DECODE (
                             V_ҩƷ�շ���¼.���ϵ��, 1,
                             NVL (V_ҩƷ�շ���¼.ʵ������, 0), 0
                         ),
                 ʵ������ =
                     NVL (ʵ������, 0) +
                         NVL (V_ҩƷ�շ���¼.ʵ������, 0) * V_ҩƷ�շ���¼.���ϵ��,
                 ʵ�ʽ�� =
                     NVL (ʵ�ʽ��, 0) +
                         NVL (V_ҩƷ�շ���¼.���۽��, 0) * V_ҩƷ�շ���¼.���ϵ��,
                 ʵ�ʲ�� =
                     NVL (ʵ�ʲ��, 0) +
                         NVL (V_ҩƷ�շ���¼.���, 0) * V_ҩƷ�շ���¼.���ϵ��
         WHERE �ⷿID = V_ҩƷ�շ���¼.�ⷿID
            AND ҩƷID = V_ҩƷ�շ���¼.ҩƷID
            AND NVL (����, 0) = NVL (V_ҩƷ�շ���¼.����, 0)
            AND ���� = 1;

        IF SQL%NOTFOUND THEN
            Insert INTO ҩƷ���
                            (
                                �ⷿID,
                                ҩƷID,
                                ����,
                                ����,
                                ��������,
                                ʵ������,
                                ʵ�ʽ��,
                                ʵ�ʲ��,
                                �ϴ�����,
                                �ϴβ���,
                                Ч��,
                                ��׼�ĺ�
                            )
                  VALUES (
                      V_ҩƷ�շ���¼.�ⷿID,
                      V_ҩƷ�շ���¼.ҩƷID,
                      V_ҩƷ�շ���¼.����,
                      1,
                      DECODE (
                          V_ҩƷ�շ���¼.���ϵ��, 1, NVL (
                                                                    V_ҩƷ�շ���¼.ʵ������, 0
                                                                ), 0
                      ),
                      V_ҩƷ�շ���¼.ʵ������ * V_ҩƷ�շ���¼.���ϵ��,
                      V_ҩƷ�շ���¼.���۽�� * V_ҩƷ�շ���¼.���ϵ��,
                      V_ҩƷ�շ���¼.��� * V_ҩƷ�շ���¼.���ϵ��,
                      V_ҩƷ�շ���¼.����,
                      V_ҩƷ�շ���¼.����,
                      V_ҩƷ�շ���¼.Ч��,
                      V_ҩƷ�շ���¼.��׼�ĺ�
                  );
        END IF;

        DELETE
          FROM ҩƷ���
         WHERE �ⷿID = V_ҩƷ�շ���¼.�ⷿID
            AND ҩƷID = V_ҩƷ�շ���¼.ҩƷID
            AND NVL (��������, 0) = 0
            AND NVL (ʵ������, 0) = 0
            AND NVL (ʵ�ʽ��, 0) = 0
            AND NVL (ʵ�ʲ��, 0) = 0;

        --����ҩƷ�շ����ܱ����Ӧ����

        UPDATE ҩƷ�շ�����
            SET ���� =
                     NVL (����, 0) +
                         NVL (V_ҩƷ�շ���¼.ʵ������, 0) * V_ҩƷ�շ���¼.���ϵ��,
                 ��� =
                     NVL (���, 0) +
                         NVL (V_ҩƷ�շ���¼.���۽��, 0) * V_ҩƷ�շ���¼.���ϵ��,
                 ��� =
                     NVL (���, 0) +
                         NVL (V_ҩƷ�շ���¼.���, 0) * V_ҩƷ�շ���¼.���ϵ��
         WHERE ���� = TRUNC (SYSDATE)
            AND �ⷿID = V_ҩƷ�շ���¼.�ⷿID
            AND ҩƷID = V_ҩƷ�շ���¼.ҩƷID
            AND ���ID = V_ҩƷ�շ���¼.������ID
            AND ���� = 12;

        IF SQL%NOTFOUND THEN
            Insert INTO ҩƷ�շ�����
                (����, �ⷿID, ҩƷID, ���ID, ����, ����, ���, ���)
                  VALUES (
                      TRUNC (SYSDATE),
                      V_ҩƷ�շ���¼.�ⷿID,
                      V_ҩƷ�շ���¼.ҩƷID,
                      V_ҩƷ�շ���¼.������ID,
                      12,
                      V_ҩƷ�շ���¼.ʵ������ * V_ҩƷ�շ���¼.���ϵ��,
                      V_ҩƷ�շ���¼.���۽�� * V_ҩƷ�շ���¼.���ϵ��,
                      V_ҩƷ�շ���¼.��� * V_ҩƷ�շ���¼.���ϵ��
                  );
        END IF;
    END LOOP;
EXCEPTION
    WHEN Err_isverified THEN
        Raise_application_error (
            -20101, '[ZLSOFT]�õ����Ѿ���������ˣ�[ZLSOFT]'
        );
    WHEN Err_isBatch THEN
        Raise_application_error ( 
            -20102, '[ZLSOFT]�õ����а�����һ��ԭ�������������ڷ�����ҩƷ��������ˣ�[ZLSOFT]' 
        ); 
        
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_ҩƷ�̵�_verify;
/


CREATE OR REPLACE PROCEDURE zl_ҩƷ�̵�_strike (
   NO_IN IN ҩƷ�շ���¼.NO%TYPE,
   �����_IN IN ҩƷ�շ���¼.�����%TYPE
)
IS
   Err_isstriked EXCEPTION;
   Err_isBatch exception;
   v_BatchCount integer;    --ԭ���������ڷ�����ҩƷ������
   V_COUNT INTEGER;    --ԭ�����ֲ�����
   V_���� ҩƷ�շ���¼.����%TYPE;

   CURSOR C_ҩƷ�շ���¼
   IS
      SELECT ID, ʵ������, ���۽��, ���, �ⷿID, ҩƷID, ����, ����,Ч��, ����,
             ������ID, ���ϵ��,����,��׼�ĺ�
        FROM ҩƷ�շ���¼
       WHERE NO = NO_IN
         AND ���� = 12
         AND ��¼״̬ = 2
       ORDER BY ҩƷID;
BEGIN
   UPDATE ҩƷ�շ���¼
      SET ��¼״̬ = 3
    WHERE NO = NO_IN
      AND ���� = 12
      AND ��¼״̬ = 1;

   IF SQL%ROWCOUNT = 0 THEN
      RAISE Err_isstriked;
   END IF;
   
   --��Ҫ���ԭ���������ڷ�����ҩƷ�����ܶ������
    SELECT COUNT(*) INTO v_BatchCount FROM        
            ҩƷ�շ���¼ a,ҩƷ��� b
    WHERE a.ҩƷid=b.ҩƷid
        AND a.no=NO_IN 
        AND a.����=12
        AND a.��¼״̬=3
        AND nvl(a.����,0)=0
        AND ((NVL(B.ҩ�����,0)=1 AND A.�ⷿID NOT IN (SELECT ����ID FROM  ��������˵�� WHERE (�������� LIKE '%ҩ��') OR (�������� LIKE '�Ƽ���')))
            OR NVL(B.ҩ������,0)=1);
        
    IF v_batchcount>0 THEN
        RAISE Err_isBatch;
    END IF;  

   Insert INTO ҩƷ�շ���¼
        (ID,��¼״̬,����,NO,���,�ⷿID,������ID,���ϵ��,
         ҩƷID,����,����,����,Ч��,��д����,����,ʵ������,
         �ɱ���,�ɱ����,���ۼ�,���۽��,���,ժҪ,
         ������,��������,�����,�������,Ƶ��,����,��׼�ĺ�
         )
      SELECT ҩƷ�շ���¼_ID.Nextval, 2, ����, NO, ���, �ⷿID, ������ID,
             ���ϵ��, a.ҩƷID, 
             DECODE(NVL(a.����,0),0,NULL,(DECODE(NVL(b.ҩ�����,0),0,NULL,a.����))), 
             a.����, ����, Ч��, ��д����, a.����,
             -ʵ������, a.�ɱ���, �ɱ����, ���ۼ�, -���۽��, -���, ժҪ,
             �����_IN, SYSDATE, �����_IN, SYSDATE, Ƶ��,����,a.��׼�ĺ�
        FROM ҩƷ�շ���¼ a,ҩƷ��� b
       WHERE NO = NO_IN
         AND a.ҩƷid=b.ҩƷid
         AND ���� = 12
         AND ��¼״̬ = 3;

   FOR V_ҩƷ�շ���¼ IN C_ҩƷ�շ���¼ LOOP
    --ԭ�����ֲ�������ҩƷ,��C����ʱ��Ҫ������
    BEGIN 
        SELECT COUNT(*) INTO V_COUNT
        FROM ҩƷ�շ���¼ A, ҩƷ��� B
        WHERE B.ҩƷID=V_ҩƷ�շ���¼.ҩƷID
        AND A.NO=NO_IN
        AND A.���� = 12
        and a.�ⷿid+0=V_ҩƷ�շ���¼.�ⷿid
        AND A.��¼״̬ = 3 
        AND NVL(A.����,0)>0
        AND (NVL(B.ҩ�����,0)=0 OR 
            (NVL(B.ҩ������,0)=0 AND A.�ⷿID IN (SELECT ����ID FROM  ��������˵�� WHERE (�������� LIKE '%ҩ��') OR (�������� LIKE '�Ƽ���'))))
        ;
    EXCEPTION 
        WHEN OTHERS THEN
            V_COUNT:=0;
    END;
    IF V_COUNT>0 THEN
        V_����:=0;
    ELSE
        V_����:=NVL (V_ҩƷ�շ���¼.����, 0);
    END IF;
      --����ҩƷ�������Ӧ����
      
        UPDATE ҩƷ���
            SET ��������=NVL(��������,0)+NVL(V_ҩƷ�շ���¼.ʵ������,0)*V_ҩƷ�շ���¼.���ϵ��,
                ʵ������=NVL(ʵ������,0)+NVL(V_ҩƷ�շ���¼.ʵ������,0)*V_ҩƷ�շ���¼.���ϵ��,
                ʵ�ʽ��=NVL(ʵ�ʽ��,0)+NVL(V_ҩƷ�շ���¼.���۽��,0)*V_ҩƷ�շ���¼.���ϵ��,
                ʵ�ʲ��=NVL(ʵ�ʲ��,0)+NVL(V_ҩƷ�շ���¼.���,0)*V_ҩƷ�շ���¼.���ϵ��
          WHERE �ⷿID = V_ҩƷ�շ���¼.�ⷿID
            AND ҩƷID = V_ҩƷ�շ���¼.ҩƷID
            AND NVL (����, 0) = V_����
            AND ���� = 1;

        IF SQL%NOTFOUND THEN
            Insert INTO ҩƷ���
                (�ⷿID,ҩƷID,����,����,��������,ʵ������,
                 ʵ�ʽ��,ʵ�ʲ��,�ϴ�����,�ϴβ���,Ч��,�ϴβɹ���,��׼�ĺ�
                 )
                VALUES (
                    V_ҩƷ�շ���¼.�ⷿID,
                    V_ҩƷ�շ���¼.ҩƷID,
                    V_����,
                    1,
                    V_ҩƷ�շ���¼.ʵ������*V_ҩƷ�շ���¼.���ϵ��,
                    V_ҩƷ�շ���¼.ʵ������*V_ҩƷ�շ���¼.���ϵ��,
                    V_ҩƷ�շ���¼.���۽��*V_ҩƷ�շ���¼.���ϵ��,
                    V_ҩƷ�շ���¼.���*V_ҩƷ�շ���¼.���ϵ��,
                    V_ҩƷ�շ���¼.����,
                    V_ҩƷ�շ���¼.����,
                    V_ҩƷ�շ���¼.Ч��,
                    V_ҩƷ�շ���¼.����,
                    V_ҩƷ�շ���¼.��׼�ĺ�
                 );
        END IF;
      

        DELETE 
          FROM ҩƷ���
         WHERE �ⷿID = V_ҩƷ�շ���¼.�ⷿID
           AND ҩƷID = V_ҩƷ�շ���¼.ҩƷID
           AND nvl(��������,0)=0 
           And nvl(ʵ������,0)=0 
           And nvl(ʵ�ʽ��,0)=0 
           And nvl(ʵ�ʲ��,0)=0;

      --����ҩƷ�շ����ܱ����Ӧ����

       UPDATE ҩƷ�շ�����
          SET ����=NVL(����,0)+NVL(V_ҩƷ�շ���¼.ʵ������,0)*V_ҩƷ�շ���¼.���ϵ��,
              ���=NVL(���,0)+NVL(V_ҩƷ�շ���¼.���۽��,0)*V_ҩƷ�շ���¼.���ϵ��,
              ���=NVL(���,0)+NVL(V_ҩƷ�շ���¼.���,0)*V_ҩƷ�շ���¼.���ϵ��
        WHERE ���� = TRUNC (SYSDATE)
          AND �ⷿID = V_ҩƷ�շ���¼.�ⷿID
          AND ҩƷID = V_ҩƷ�շ���¼.ҩƷID
          AND ���ID = V_ҩƷ�շ���¼.������ID
          AND ���� = 12;

      IF SQL%NOTFOUND THEN
         Insert INTO ҩƷ�շ�����
            (����, �ⷿID, ҩƷID, ���ID, ����, ����, ���, ���)
              VALUES (
                 TRUNC (SYSDATE),
                 V_ҩƷ�շ���¼.�ⷿID,
                 V_ҩƷ�շ���¼.ҩƷID,
                 V_ҩƷ�շ���¼.������ID,
                 12,
                 V_ҩƷ�շ���¼.ʵ������*V_ҩƷ�շ���¼.���ϵ��,
                 V_ҩƷ�շ���¼.���۽��*V_ҩƷ�շ���¼.���ϵ��,
                 V_ҩƷ�շ���¼.���*V_ҩƷ�շ���¼.���ϵ��
              );
      END IF;
   END LOOP;
EXCEPTION
   WHEN Err_isstriked THEN
      Raise_application_error (
         -20101, '[ZLSOFT]�õ����Ѿ������˳�����[ZLSOFT]'
      );
   WHEN Err_isBatch THEN
      Raise_application_error ( 
            -20102, '[ZLSOFT]�õ����а�����һ��ԭ�������������ڷ�����ҩƷ�����ܳ�����[ZLSOFT]' 
        ); 
   WHEN OTHERS THEN
      zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_ҩƷ�̵�_strike;
/

CREATE OR REPLACE PROCEDURE zl_ҩƷ�̵��¼��_Insert (
    NO_IN IN ҩƷ�շ���¼.NO%TYPE,
    ���_IN IN ҩƷ�շ���¼.���%TYPE,
    �ⷿID_IN IN ҩƷ�շ���¼.�ⷿID%TYPE,
    ����_IN IN ҩƷ�շ���¼.����%TYPE,
    ������ID_IN IN ҩƷ�շ���¼.������ID%TYPE,
    --ÿ�ζ�ʹ�ã����������洫��ȽϺ�
    ���ϵ��_IN IN ҩƷ�շ���¼.���ϵ��%TYPE,
    ҩƷID_IN IN ҩƷ�շ���¼.ҩƷID%TYPE,
    ��������_IN IN ҩƷ�շ���¼.��д����%TYPE,
    ʵ������_IN IN ҩƷ�շ���¼.����%TYPE,
    ������_IN IN ҩƷ�շ���¼.ʵ������%TYPE,
    �ۼ�_IN IN ҩƷ�շ���¼.���ۼ�%TYPE,
    ����_IN IN ҩƷ�շ���¼.���۽��%TYPE,
    ��۲�_IN IN ҩƷ�շ���¼.���%TYPE,
    ������_IN IN ҩƷ�շ���¼.������%TYPE,
    ��������_IN IN ҩƷ�շ���¼.��������%TYPE,
    ժҪ_IN IN ҩƷ�շ���¼.ժҪ%TYPE := NULL,
    ����_IN IN ҩƷ�շ���¼.����%TYPE := NULL,
    ����_IN IN ҩƷ�շ���¼.����%TYPE := NULL,
    Ч��_IN IN ҩƷ�շ���¼.Ч��%TYPE := NULL,
    �̵�ʱ��_IN IN ҩƷ�շ���¼.Ƶ��%TYPE := NULL,
    �����_IN IN ҩƷ�շ���¼.�ɱ���%TYPE := NULL,
    �����_IN IN ҩƷ�շ���¼.�ɱ����%TYPE := NULL,
    ��׼�ĺ�_IN In ҩƷ�շ���¼.��׼�ĺ�%TYPE := Null
)
IS
BEGIN
    Insert INTO ҩƷ�շ���¼
        (ID,��¼״̬,����,NO,���,�ⷿID,������ID,���ϵ��,ҩƷID,����,����,����,Ч��,
        ��д����,����,ʵ������,���ۼ�,���۽��,���,ժҪ,������,��������,Ƶ��,�ɱ���,�ɱ����,��׼�ĺ�)
    VALUES (
        ҩƷ�շ���¼_ID.Nextval,1,14,NO_IN,���_IN,�ⷿID_IN,������ID_IN,���ϵ��_IN,ҩƷID_IN,
        ����_IN,����_IN,����_IN,Ч��_IN,��������_IN,ʵ������_IN,������_IN,�ۼ�_IN,����_IN,
        ��۲�_IN,ժҪ_IN,������_IN,��������_IN,�̵�ʱ��_IN,�����_IN,�����_IN,��׼�ĺ�_IN);
EXCEPTION
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_ҩƷ�̵��¼��_Insert;
/

CREATE OR REPLACE PROCEDURE zl_ҩƷ����_Insert (
    NO_IN IN ҩƷ�շ���¼.NO%TYPE,
    ���_IN IN ҩƷ�շ���¼.���%TYPE,
    �ⷿID_IN IN ҩƷ�շ���¼.�ⷿID%TYPE,
    �Է�����ID_IN IN ҩƷ�շ���¼.�Է�����ID%TYPE,
    ҩƷID_IN IN ҩƷ�շ���¼.ҩƷID%TYPE,
    ����_IN IN ҩƷ�շ���¼.����%TYPE,
    ��д����_IN IN ҩƷ�շ���¼.��д����%TYPE,
    ʵ������_IN in ҩƷ�շ���¼.ʵ������%TYPE,
    �ɱ���_IN IN ҩƷ�շ���¼.�ɱ���%TYPE,
    �ɱ����_IN IN ҩƷ�շ���¼.�ɱ����%TYPE,
    ���ۼ�_IN IN ҩƷ�շ���¼.���ۼ�%TYPE,
    ���۽��_IN IN ҩƷ�շ���¼.���۽��%TYPE,
    ���_IN IN ҩƷ�շ���¼.���%TYPE,
    ������_IN IN ҩƷ�շ���¼.������%TYPE,
    ����_IN IN ҩƷ�շ���¼.����%TYPE := NULL,
    ����_IN IN ҩƷ�շ���¼.����%TYPE := NULL,
    Ч��_IN IN ҩƷ�շ���¼.Ч��%TYPE := NULL,
    ժҪ_IN IN ҩƷ�շ���¼.ժҪ%TYPE := NULL,
    ��������_IN IN ҩƷ�շ���¼.��������%TYPE := NULL,
    �ϴι�Ӧ��ID_IN In ҩƷ�շ���¼.��ҩ��λID%TYPE := Null,
    ��׼�ĺ�_IN In ҩƷ�շ���¼.��׼�ĺ�%TYPE := Null
)
IS
    V_lngID ҩƷ�շ���¼.ID%TYPE;--�շ�ID
    V_������ID ҩƷ�շ���¼.������ID%TYPE;--������ID
    V_�������ID ҩƷ�շ���¼.������ID%TYPE;--������ID
    V_���� �շ���ĿĿ¼.����%TYPE;
    V_�¿��ÿ�� ϵͳ������.����ֵ%Type;

	ERR_MutilROW EXCEPTION ;
	intRecords NUMBER ;
BEGIN
    --�����ҳ���ͳ������ID
    SELECT B.ID INTO V_������ID
    FROM ҩƷ�������� A, ҩƷ������ B
    WHERE A.���ID = B.ID AND A.���� = 6 AND B.ϵ�� = 1 AND ROWNUM < 2;
    SELECT B.ID INTO V_�������ID
    FROM ҩƷ�������� A, ҩƷ������ B
    WHERE A.���ID = B.ID AND A.���� = 6 AND B.ϵ�� = -1 AND ROWNUM < 2;

	SELECT ҩƷ�շ���¼_ID.Nextval INTO V_lngID FROM Dual;

    --�������Ϊ������һ��
    Insert INTO ҩƷ�շ���¼
        (ID,��¼״̬,����,NO,���,�ⷿID,�Է�����ID,������ID,���ϵ��,ҩƷID,����,����,����,Ч��,
        ��д����,ʵ������,�ɱ���,�ɱ����,���ۼ�,���۽��,���,ժҪ,������,��������,��ҩ��ʽ,��ҩ��λID,��׼�ĺ�)
    VALUES (ҩƷ�շ���¼_ID.Nextval,1,6,NO_IN,���_IN,�ⷿID_IN,�Է�����ID_IN,
		V_�������ID,-1,ҩƷID_IN,����_IN,����_IN,����_IN,Ч��_IN,
		��д����_IN,ʵ������_IN,�ɱ���_IN,�ɱ����_IN,���ۼ�_IN,���۽��_IN,
		���_IN,ժҪ_IN,������_IN,��������_IN,1,�ϴι�Ӧ��ID_IN,��׼�ĺ�_IN);

    --�������Ϊ�����һ��
    Insert INTO ҩƷ�շ���¼
        (ID,��¼״̬,����,NO,���,�ⷿID,�Է�����ID,������ID,���ϵ��,ҩƷID,����,����,����,Ч��,
        ��д����,ʵ������,�ɱ���,�ɱ����,���ۼ�,���۽��,���,ժҪ,������,��������,��ҩ��ʽ,��ҩ��λID,��׼�ĺ�)
          VALUES (V_lngID,1,6,NO_IN,���_IN + 1,�Է�����ID_IN,�ⷿID_IN,V_������ID,
              1,ҩƷID_IN,����_IN,����_IN,����_IN,Ч��_IN,��д����_IN,ʵ������_IN,
              �ɱ���_IN,�ɱ����_IN,���ۼ�_IN,���۽��_IN,���_IN,ժҪ_IN,
              ������_IN,��������_IN,1,�ϴι�Ӧ��ID_IN,��׼�ĺ�_IN);

	--����Ƿ������ͬҩƷ��ͬ���ε����ݣ�������ڲ�������
	SELECT COUNT(*) INTO intRecords 
	FROM ҩƷ�շ���¼
	WHERE ����=6 AND NO=NO_IN AND ���ϵ��=-1 AND ҩƷID+0=ҩƷID_IN AND Nvl(����,0)=NVL(����_IN,0);
	IF intRecords>1 THEN 
		RAISE ERR_MutilROW;
	END IF ;
  
	--���ݲ��������Ƿ��·�ҩ�ⷿ�Ŀ��ÿ��
	Select ����ֵ Into v_�¿��ÿ�� From ϵͳ������ Where ������ = 96;
  
  --����Ϊ1��ʾ���ʱ�¿�������
	If v_�¿��ÿ�� = '1' Then
      UPDATE ҩƷ���
			SET ��������=NVL(��������,0)-ʵ������_IN
			WHERE �ⷿID=�ⷿID_IN AND ҩƷID=ҩƷID_IN AND NVL(����,0)=����_IN AND ����=1;
      
			IF SQL%ROWCOUNT=0 THEN 
				INSERT INTO ҩƷ���(�ⷿID,ҩƷID,����,����,��������,�ϴ�����,Ч��,�ϴβ���)
				VALUES (�ⷿID_IN,ҩƷID_IN,����_IN,1,-1*ʵ������_IN,
					����_IN,Ч��_IN,����_IN);
			END IF ;

			DELETE
			FROM ҩƷ���
			WHERE �ⷿID = �ⷿID_IN
			AND ҩƷID = ҩƷID_IN
			AND NVL(��������,0)=0
			AND NVL(ʵ������,0)=0
			AND NVL(ʵ�ʽ��,0)=0
			AND NVL(ʵ�ʲ��,0)=0;
  End If;
  
EXCEPTION
	WHEN ERR_MutilROW THEN 
		SELECT ���� INTO V_���� FROM �շ���ĿĿ¼ WHERE ID = ҩƷID_IN;
		RAISE_APPLICATION_ERROR(-20101,'[ZLSOFT]����Ϊ'||V_����||'��ҩƷ�����ڶ����ظ��ļ�¼����ϲ�Ϊһ����¼��[ZLSOFT]');
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_ҩƷ����_Insert;
/


CREATE OR REPLACE PROCEDURE zl_ҩƷ����_DELETE (
    NO_IN IN ҩƷ�շ���¼.NO%TYPE
)
Is
    V_�¿��ÿ�� ϵͳ������.����ֵ%Type;
    
    CURSOR C_ҩƷ�շ���¼
    IS
		SELECT ʵ������,�ⷿID,����,ҩƷID,����,Ч��,����,��ҩ��λID,��׼�ĺ�
		FROM ҩƷ�շ���¼
		WHERE NO = NO_IN
		AND ���� = 6
		AND ���ϵ�� = -1
		ORDER BY ҩƷID;
BEGIN
    --���ݲ��������Ƿ�ָ���ҩ�ⷿ�Ŀ��ÿ��
    Select ����ֵ Into v_�¿��ÿ�� From ϵͳ������ Where ������ = 96;
    
    --����Ϊ1��ʾ��Ҫ�ָ���������
    IF  v_�¿��ÿ��='1' THEN
      FOR V_ҩƷ�շ���¼ IN C_ҩƷ�շ���¼ LOOP
  			UPDATE ҩƷ���
  			SET �������� = NVL (��������, 0) + V_ҩƷ�շ���¼.ʵ������
  			WHERE �ⷿID = V_ҩƷ�շ���¼.�ⷿID
  			AND ҩƷID = V_ҩƷ�շ���¼.ҩƷID
  			AND NVL (����, 0) = NVL (V_ҩƷ�շ���¼.����, 0)
  			AND ���� = 1;
  
  			IF SQL%NOTFOUND THEN
  				INSERT INTO ҩƷ���
  					(�ⷿID, ҩƷID, ����, ����, ��������,�ϴ�����,Ч��,�ϴβ���,�ϴι�Ӧ��ID,��׼�ĺ�)
  				VALUES (
  					V_ҩƷ�շ���¼.�ⷿID,V_ҩƷ�շ���¼.ҩƷID,V_ҩƷ�շ���¼.����,1,
  					V_ҩƷ�շ���¼.ʵ������,V_ҩƷ�շ���¼.����,V_ҩƷ�շ���¼.Ч��,V_ҩƷ�շ���¼.����,V_ҩƷ�շ���¼.��ҩ��λID,V_ҩƷ�շ���¼.��׼�ĺ�
  					);
  			END IF;
  
  			DELETE
  			FROM ҩƷ���
  			WHERE �ⷿID = V_ҩƷ�շ���¼.�ⷿID
  			AND ҩƷID = V_ҩƷ�շ���¼.ҩƷID
  			AND NVL(��������,0) = 0
  			AND NVL(ʵ������,0) = 0
  			AND NVL(ʵ�ʽ��,0) = 0
  			AND NVL(ʵ�ʲ��,0) = 0;
  		END LOOP;
	  END IF ;
    
    DELETE ҩƷ�շ���¼
    WHERE NO = NO_IN AND ���� = 6 AND ��¼״̬ = 1 AND ����� IS NULL;
    
EXCEPTION
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_ҩƷ����_DELETE;
/

CREATE OR REPLACE PROCEDURE zl_ҩƷ����_Insert (
    ������ID_IN IN ҩƷ�շ���¼.������ID%TYPE,
    NO_IN IN ҩƷ�շ���¼.NO%TYPE,
    ���_IN IN ҩƷ�շ���¼.���%TYPE,
    �ⷿID_IN IN ҩƷ�շ���¼.�ⷿID%TYPE,
    �Է�����ID_IN IN ҩƷ�շ���¼.�Է�����ID%TYPE,
    ҩƷID_IN IN ҩƷ�շ���¼.ҩƷID%TYPE,
    ����_IN IN ҩƷ�շ���¼.����%TYPE,
    ��д����_IN IN ҩƷ�շ���¼.��д����%TYPE,
    �ɱ���_IN IN ҩƷ�շ���¼.�ɱ���%TYPE,
    �ɱ����_IN IN ҩƷ�շ���¼.�ɱ����%TYPE,
    ���ۼ�_IN IN ҩƷ�շ���¼.���ۼ�%TYPE,
    ���۽��_IN IN ҩƷ�շ���¼.���۽��%TYPE,
    ���_IN IN ҩƷ�շ���¼.���%TYPE,
    ������_IN IN ҩƷ�շ���¼.������%TYPE,
    ��������_IN IN ҩƷ�շ���¼.��������%TYPE,
    ����_IN IN ҩƷ�շ���¼.����%TYPE := NULL,
    ����_IN IN ҩƷ�շ���¼.����%TYPE := NULL,
    Ч��_IN IN ҩƷ�շ���¼.Ч��%TYPE := NULL,
    ժҪ_IN IN ҩƷ�շ���¼.ժҪ%TYPE := Null,
    ������_IN IN ҩƷ�շ���¼.������%TYPE := Null,
    ��׼�ĺ�_IN In ҩƷ�շ���¼.��׼�ĺ�%TYPE := Null
)
IS
    V_���ϵ�� ҩƷ�շ���¼.���ϵ��%TYPE;--�շ�ID
    Err_isNOnumber EXCEPTION;
    V_���� �շ���ĿĿ¼.����%TYPE;
    V_�������� ҩƷ���.��������%TYPE;
BEGIN
    V_���ϵ�� := -1;

    IF ����_IN > 0 THEN
        BEGIN
            SELECT ��������
              INTO V_��������
              FROM ҩƷ���
             WHERE ҩƷID = ҩƷID_IN
                AND NVL (����, 0) = ����_IN
                AND �ⷿID = �ⷿID_IN
                AND ���� = 1
                AND ROWNUM = 1;
        EXCEPTION
            WHEN OTHERS THEN
                V_�������� := 0;
        END;

        IF V_�������� - ��д����_IN < 0 THEN
            RAISE Err_isNOnumber;
        END IF;
    END IF;

    --�������Ϊ������һ��
    Insert INTO ҩƷ�շ���¼
                    (
                        ID,
                        ��¼״̬,
                        ����,
                        NO,
                        ���,
                        �ⷿID,
                        �Է�����ID,
                        ������ID,
                        ���ϵ��,
                        ҩƷID,
                        ����,
                        ����,
                        ����,
                        Ч��,
                        ��д����,
                        ʵ������,
                        �ɱ���,
                        �ɱ����,
                        ���ۼ�,
                        ���۽��,
                        ���,
                        ժҪ,
                        ������,
                        ��������,
                        ������,
                        ��׼�ĺ�
                    )
          VALUES (
              ҩƷ�շ���¼_ID.Nextval,
              1,
              7,
              NO_IN,
              ���_IN,
              �ⷿID_IN,
              �Է�����ID_IN,
              ������ID_IN,
              V_���ϵ��,
              ҩƷID_IN,
              ����_IN,
              ����_IN,
              ����_IN,
              Ч��_IN,
              ��д����_IN,
              ��д����_IN,
              �ɱ���_IN,
              �ɱ����_IN,
              ���ۼ�_IN,
              ���۽��_IN,
              ���_IN,
              ժҪ_IN,
              ������_IN,
              ��������_IN,
              ������_IN,
              ��׼�ĺ�_IN
          );

    --ͬʱ���¿����
    UPDATE ҩƷ���
        SET �������� = NVL (��������, 0) - ��д����_IN
     WHERE �ⷿID = �ⷿID_IN
        AND ҩƷID = ҩƷID_IN
        AND NVL (����, 0) = NVL (����_IN, 0)
        AND ���� = 1;

    --��������������Ϊ����ҩƷ��������׼����
    IF SQL%NOTFOUND THEN
        Insert INTO ҩƷ���
                        (�ⷿID, ҩƷID, ����, ��������,�ϴ�����,Ч��,�ϴβ���,��׼�ĺ�)
              VALUES (�ⷿID_IN, ҩƷID_IN, 1, -��д����_IN,����_IN,Ч��_IN,����_IN,��׼�ĺ�_IN);
    END IF;

    DELETE
      FROM ҩƷ���
     WHERE �ⷿID = �ⷿID_IN
        AND ҩƷID = ҩƷID_IN
        AND nvl(��������,0) = 0
        AND nvl(ʵ������,0) = 0
        AND nvl(ʵ�ʽ��,0) = 0
        AND nvl(ʵ�ʲ��,0) = 0;
EXCEPTION
    WHEN Err_isNOnumber THEN
        SELECT ����
          INTO V_����
          FROM �շ���ĿĿ¼
         WHERE ID = ҩƷID_IN;
        Raise_application_error (
            -20101, '[ZLSOFT]����Ϊ' || V_���� || ',����Ϊ' || ����_IN ||
                          '�ķ�������ҩƷ' ||
                          CHR (10) ||
                          CHR (13) ||
                          '���ÿ������������[ZLSOFT]'
        );
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_ҩƷ����_Insert;
/


CREATE OR REPLACE PROCEDURE zl_ҩƷ����_DELETE (
    
    --ɾ��ҩƷ�շ���¼���ָ���Ӧ�ı�ҩƷ���
    NO_IN IN ҩƷ�շ���¼.NO%TYPE
)
IS
    Err_isverified EXCEPTION;

    CURSOR C_ҩƷ�շ���¼
    IS
        SELECT ��д����, �ⷿID, ����, ҩƷID,����,Ч��,����,��׼�ĺ�
          FROM ҩƷ�շ���¼
         WHERE NO = NO_IN
            AND ���� = 7
            AND ���ϵ�� = -1
         ORDER BY ҩƷID;
BEGIN
    --ͨ��ѭ�����ָ�ԭ���Ŀ�������
    FOR V_ҩƷ�շ���¼ IN C_ҩƷ�շ���¼ LOOP
        UPDATE ҩƷ���
            SET �������� = NVL (��������, 0) + V_ҩƷ�շ���¼.��д����
         WHERE �ⷿID = V_ҩƷ�շ���¼.�ⷿID
            AND ҩƷID = V_ҩƷ�շ���¼.ҩƷID
            AND NVL (����, 0) = NVL (V_ҩƷ�շ���¼.����, 0)
            AND ���� = 1;

        IF SQL%NOTFOUND THEN
            Insert INTO ҩƷ���
                            (�ⷿID, ҩƷID, ����, ����, ��������,�ϴ�����,Ч��,�ϴβ���,��׼�ĺ�)
                  VALUES (
                      V_ҩƷ�շ���¼.�ⷿID,V_ҩƷ�շ���¼.ҩƷID,V_ҩƷ�շ���¼.����,1,
                      V_ҩƷ�շ���¼.��д����,V_ҩƷ�շ���¼.����,V_ҩƷ�շ���¼.Ч��,V_ҩƷ�շ���¼.����,V_ҩƷ�շ���¼.��׼�ĺ�
                  );
        END IF;

        DELETE
          FROM ҩƷ���
         WHERE �ⷿID = V_ҩƷ�շ���¼.�ⷿID
            AND ҩƷID = V_ҩƷ�շ���¼.ҩƷID
            AND NVL (��������, 0) = 0
            AND NVL (ʵ������, 0) = 0
            AND NVL (ʵ�ʽ��, 0) = 0
            AND NVL (ʵ�ʲ��, 0) = 0;
    END LOOP;

    DELETE
      FROM ҩƷ�շ���¼
     WHERE NO = NO_IN
        AND ���� = 7
        AND ��¼״̬ = 1
        AND ����� IS NULL;

    IF SQL%ROWCOUNT = 0 THEN
        RAISE Err_isverified;
    END IF;
EXCEPTION
    WHEN Err_isverified THEN
        Raise_application_error (
            -20101, '[ZLSOFT]�õ����Ѿ�������ɾ�����ѱ�����ˣ�[ZLSOFT]'
        );
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_ҩƷ����_DELETE;
/


CREATE OR REPLACE PROCEDURE zl_ҩƷ����_verify (
    ���_IN IN ҩƷ�շ���¼.���%TYPE,
    NO_IN IN ҩƷ�շ���¼.NO%TYPE,
    �ⷿID_IN IN ҩƷ�շ���¼.�ⷿID%TYPE,
    �Է�����ID_IN IN ҩƷ�շ���¼.�Է�����ID%TYPE,
    ҩƷID_IN IN ҩƷ�շ���¼.ҩƷID%TYPE,
    ����_IN IN ҩƷ�շ���¼.����%TYPE,
    ����_IN IN ҩƷ�շ���¼.����%TYPE,
    ��д����_IN IN ҩƷ�շ���¼.��д����%TYPE,
    ʵ������_IN IN ҩƷ�շ���¼.ʵ������%TYPE,
    �ɱ���_IN IN ҩƷ�շ���¼.�ɱ���%TYPE,
    �ɱ����_IN IN ҩƷ�շ���¼.�ɱ����%TYPE,
    ���۽��_IN IN ҩƷ�շ���¼.���۽��%TYPE,
    ���_IN IN ҩƷ�շ���¼.���%TYPE,
    ������ID_IN IN ҩƷ�շ���¼.������ID%TYPE,
    �����_IN IN ҩƷ�շ���¼.�����%TYPE,
    �������_IN IN ҩƷ�շ���¼.�������%TYPE,
    ����_IN IN ҩƷ�շ���¼.����%TYPE := NULL,
    Ч��_IN IN ҩƷ�շ���¼.Ч��%TYPE := NULL,
    ��׼�ĺ�_IN In ҩƷ�շ���¼.��׼�ĺ�%TYPE := Null
)
IS
    Err_isverified EXCEPTION;
    Err_isNOnumber EXCEPTION;
    V_�������� ҩƷ���.��������%TYPE;
    V_���� �շ���ĿĿ¼.����%TYPE;
    V_ʵ�ʿ���� ҩƷ���.ʵ�ʽ��%TYPE;
    V_ʵ�ʿ���� ҩƷ���.ʵ�ʲ��%TYPE;
    V_����� number(18,8);
    V_������ ҩƷ���.ʵ�ʲ��%TYPE;
    V_�ɱ��� ҩƷ�շ���¼.�ɱ���%TYPE;
    V_�ɱ���� ҩƷ�շ���¼.�ɱ����%TYPE;
	INTDIGIT NUMBER ;
BEGIN
	--��ȡ���С��λ��
    SELECT NVL(����ֵ,ȱʡֵ) INTO INTDIGIT FROM ϵͳ������ WHERE ������='���ý���λ��';

    --�������ô������������ʱ�ı�ʵ��������
      --�������ȶ�ʵ��������������Ӧ���ֶν��и��¡�
    BEGIN
        SELECT nvl(ʵ�ʽ��,0), nvl(ʵ�ʲ��,0), nvl(��������,0)
          INTO V_ʵ�ʿ����, V_ʵ�ʿ����, V_��������
          FROM ҩƷ���
         WHERE ҩƷID = ҩƷID_IN
            AND NVL (����, 0) = ����_IN
            AND �ⷿID = �ⷿID_IN
            AND ���� = 1
            AND ROWNUM = 1;
    EXCEPTION
        WHEN OTHERS THEN
            V_ʵ�ʿ���� := 0;
            V_�������� := 0;
    END;

    IF V_ʵ�ʿ���� <= 0 THEN
        BEGIN
            SELECT ָ������� / 100
              INTO V_�����
              FROM ҩƷ���
             WHERE ҩƷID = ҩƷID_IN;
        EXCEPTION
            WHEN OTHERS THEN
                V_����� := 0;
        END;
    ELSE
        V_����� := V_ʵ�ʿ���� / V_ʵ�ʿ����;
    END IF;

    V_������ := round(���۽��_IN * V_�����,INTDIGIT);
    IF ʵ������_IN=0 THEN 
        V_�ɱ��� :=�ɱ���_IN; 
    ELSE 
        V_�ɱ��� := (���۽��_IN - V_������) / ʵ������_IN; 
    END IF; 
    V_�ɱ���� := round(V_�ɱ��� * ʵ������_IN,INTDIGIT);

    UPDATE ҩƷ�շ���¼
        SET ����� = NVL (�����_IN, �����),
             ������� = �������_IN,
             ʵ������ = ʵ������_IN,
             �ɱ��� = V_�ɱ���,
             �ɱ���� = V_�ɱ����,
             ���۽�� = ���۽��_IN,
             ��� = V_������
     WHERE NO = NO_IN
        AND ���� = 7
        AND ҩƷID = ҩƷID_IN
        AND ��� = ���_IN
        AND ��¼״̬ = 1
        AND ����� IS NULL;

    --����ҩƷ������Ӧ����
    IF SQL%ROWCOUNT = 0 THEN
        RAISE Err_isverified;
    END IF;

    IF ����_IN > 0 THEN
        IF V_�������� + ��д����_IN - ʵ������_IN < 0 THEN
            RAISE Err_isNOnumber;
        END IF;
    END IF;

    UPDATE ҩƷ���
        SET �������� = NVL (��������, 0) + ��д����_IN - ʵ������_IN,
             ʵ������ = NVL (ʵ������, 0) - ʵ������_IN,
             ʵ�ʽ�� = NVL (ʵ�ʽ��, 0) - ���۽��_IN,
             ʵ�ʲ�� = NVL (ʵ�ʲ��, 0) - V_������
     WHERE �ⷿID = �ⷿID_IN
        AND ҩƷID = ҩƷID_IN
        AND NVL (����, 0) = NVL (����_IN, 0)
        AND ���� = 1;

    IF SQL%NOTFOUND THEN
        Insert INTO ҩƷ���
                        (
                            �ⷿID,
                            ҩƷID,
                            ����,
                            ����,
                            ��������,
                            ʵ������,
                            ʵ�ʽ��,
                            ʵ�ʲ��,
							�ϴ�����,Ч��,�ϴβ���,��׼�ĺ�
                        )
              VALUES (
                  �ⷿID_IN,
                  ҩƷID_IN,
                  ����_IN,
                  1,
                  -ʵ������_IN,
                  -ʵ������_IN,
                  -���۽��_IN,
                  -V_������,
				  ����_IN,Ч��_IN,����_IN,��׼�ĺ�_IN
              );
    END IF;

    DELETE
      FROM ҩƷ���
     WHERE �ⷿID = �ⷿID_IN
        AND ҩƷID = ҩƷID_IN
        AND NVL (��������, 0) = 0
        AND NVL (ʵ������, 0) = 0
        AND NVL (ʵ�ʽ��, 0) = 0
        AND NVL (ʵ�ʲ��, 0) = 0;

    --��ҩƷ�շ����ܱ����Ӧ����
    UPDATE ҩƷ�շ�����
        SET ���� = NVL (����, 0) - ʵ������_IN,
             ��� = NVL (���, 0) - ���۽��_IN,
             ��� = NVL (���, 0) - V_������
     WHERE ���� = TRUNC (SYSDATE)
        AND �ⷿID = �ⷿID_IN
        AND ҩƷID = ҩƷID_IN
        AND ���ID = ������ID_IN
        AND ���� = 7;

    IF SQL%NOTFOUND THEN
        Insert INTO ҩƷ�շ�����
                        (����, �ⷿID, ҩƷID, ���ID, ����, ����, ���, ���)
              VALUES (
                  TRUNC (SYSDATE),
                  �ⷿID_IN,
                  ҩƷID_IN,
                  ������ID_IN,
                  7,
                  -ʵ������_IN,
                  -���۽��_IN,
                  -V_������
              );
    END IF;
EXCEPTION
    WHEN Err_isverified THEN
        Raise_application_error (
            -20101, '[ZLSOFT]�õ����Ѿ���������ˣ�[ZLSOFT]'
        );
    WHEN Err_isNOnumber THEN
        SELECT ����
          INTO V_����
          FROM �շ���ĿĿ¼
         WHERE ID = ҩƷID_IN;
        Raise_application_error (
            -20101, '[ZLSOFT]����Ϊ' || V_���� || ',����Ϊ' || ����_IN ||
                          '�ķ�������ҩƷ' ||
                          CHR (10) ||
                          CHR (13) ||
                          '���ÿ������������[ZLSOFT]'
        );
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_ҩƷ����_verify;
/

CREATE OR REPLACE PROCEDURE ZL_ҩƷ����_STRIKE (
    �д�_IN IN INTEGER,
    ԭ��¼״̬_IN IN ҩƷ�շ���¼.��¼״̬%TYPE,
    NO_IN IN ҩƷ�շ���¼.NO%TYPE,
    ���_IN IN ҩƷ�շ���¼.���%TYPE,
    ҩƷID_IN IN ҩƷ�շ���¼.ҩƷID%TYPE,
    ��������_IN IN ҩƷ�շ���¼.ʵ������%TYPE,
    ������_IN IN ҩƷ�շ���¼.������%TYPE,
    ��������_IN IN ҩƷ�շ���¼.��������%TYPE
)
IS
    Err_isstriked EXCEPTION;
    Err_isoutstock EXCEPTION;
    ERR_ISNONUM EXCEPTION;
    Err_isBatch EXCEPTION;
    v_BatchCount INTEGER;    --ԭ���������ڷ�����ҩƷ������

    V_�ⷿID ҩƷ�շ���¼.�ⷿID%TYPE; 
    V_�Է�����ID ҩƷ�շ���¼.�Է�����ID%TYPE;
    V_������ID ҩƷ�շ���¼.������ID%TYPE ;
    V_���� ҩƷ�շ���¼.����%TYPE ; 
    V_���� ҩƷ�շ���¼.����%TYPE ; 
    V_���� ҩƷ�շ���¼.����%TYPE ; 
    V_Ч�� ҩƷ�շ���¼.Ч��%TYPE ; 
    V_�ɱ��� ҩƷ�շ���¼.�ɱ���%TYPE ; 
    V_�ɱ���� ҩƷ�շ���¼.�ɱ����%TYPE ; 
    V_���� ҩƷ�շ���¼.����%TYPE ; 
    V_���ۼ� ҩƷ�շ���¼.���ۼ�%TYPE ; 
    V_���۽�� ҩƷ�շ���¼.���۽��%TYPE ; 
    V_��� ҩƷ�շ���¼.���%TYPE ; 
    V_ժҪ ҩƷ�շ���¼.ժҪ%TYPE ; 
    V_ʣ������ ҩƷ�շ���¼.ʵ������%TYPE; 
    V_ʣ��ɱ���� ҩƷ�շ���¼.�ɱ����%Type;
    V_ʣ�����۽�� ҩƷ�շ���¼.���۽��%Type;
    V_���ϵ�� ҩƷ�շ���¼.���ϵ��%TYPE; 

    V_��¼�� NUMBER; 
    V_�շ�ID ҩƷ�շ���¼.ID%TYPE; 
    V_������ ҩƷ�շ���¼.������%Type;
    V_��׼�ĺ� ҩƷ�շ���¼.��׼�ĺ�%TYPE;
	INTDIGIT NUMBER;
BEGIN
	--��ȡ���С��λ��
    SELECT NVL(����ֵ,ȱʡֵ) INTO INTDIGIT FROM ϵͳ������ WHERE ������='���ý���λ��';

	IF �д�_IN =1 THEN
        UPDATE ҩƷ�շ���¼ 
        SET ��¼״̬ = DECODE(ԭ��¼״̬_IN,1,3,ԭ��¼״̬_IN+3) 
        WHERE NO = NO_IN AND ���� = 7 AND ��¼״̬ =ԭ��¼״̬_IN ; 
        IF SQL%ROWCOUNT = 0 THEN 
            RAISE ERR_ISSTRIKED; 
        END IF; 
    END IF;
    
    --��Ҫ���ԭ���������ڷ�����ҩƷ�����ܶ������
    SELECT COUNT(*) INTO V_BATCHCOUNT FROM        
            ҩƷ�շ���¼ A,ҩƷ��� B
    WHERE A.ҩƷID=B.ҩƷID
        AND A.NO=NO_IN 
        AND A.����=7
        AND Mod(A.��¼״̬,3)=0
        AND NVL(A.����,0)=0
        AND ((NVL(B.ҩ�����,0)=1 AND A.�ⷿID NOT IN (SELECT ����ID FROM  ��������˵�� WHERE (�������� LIKE '%ҩ��') OR (�������� LIKE '�Ƽ���')))
            OR NVL(B.ҩ������,0)=1);
        
    IF V_BATCHCOUNT>0 THEN
        RAISE ERR_ISBATCH;
    END IF;  
    
    SELECT SUM(ʵ������) AS ʣ������,SUM(�ɱ����) AS ʣ��ɱ����,SUM(���۽��) AS ʣ�����۽��,�ⷿID,�Է�����ID,������ID,���ϵ��,����,����,����,Ч��,�ɱ���,����,���ۼ�,ժҪ,������,��׼�ĺ�
    INTO  V_ʣ������,V_ʣ��ɱ����,V_ʣ�����۽��,V_�ⷿID,V_�Է�����ID,V_������ID,V_���ϵ��,V_����,V_����,V_����,V_Ч��,V_�ɱ���,V_����,V_���ۼ�,V_ժҪ,V_������,V_��׼�ĺ�
    FROM ҩƷ�շ���¼ 
    WHERE NO=NO_IN 
    AND ����=7
    AND ҩƷID=ҩƷID_IN 
    AND ���=���_IN
    GROUP BY �ⷿID,�Է�����ID,������ID,���ϵ��,����,����,����,Ч��,�ɱ���,����,���ۼ�,ժҪ,������,��׼�ĺ�;

    --������������ʣ��������������
    IF V_ʣ������<��������_IN THEN
        RAISE ERR_ISNONUM; 
    END IF;

    V_�ɱ����:= ROUND(��������_IN/v_ʣ������*V_ʣ��ɱ����,INTDIGIT);
    V_���۽��:= ROUND(��������_IN/v_ʣ������*V_ʣ�����۽��,INTDIGIT);
    V_���:=V_���۽��-V_�ɱ����;

    SELECT ҩƷ�շ���¼_ID.NEXTVAL INTO V_�շ�ID  FROM DUAL;
    Insert INTO ҩƷ�շ���¼
    (ID,��¼״̬,����,NO,���,�ⷿID,�Է�����ID,������ID,���ϵ��,
    ҩƷID,����,����,����,Ч��,��д����,ʵ������,�ɱ���,
    �ɱ����,���ۼ�,���۽��,���,ժҪ,������,��������,�����,�������,������,��׼�ĺ�)
    VALUES 
    (V_�շ�ID,DECODE(ԭ��¼״̬_IN,1,2,ԭ��¼״̬_IN+2),7,NO_IN,���_IN,V_�ⷿID,V_�Է�����ID,V_������ID,V_���ϵ��,
    ҩƷID_IN,V_����,V_����,V_����,V_Ч��,-��������_IN,-��������_IN,V_�ɱ���,-V_�ɱ����,V_���ۼ�,-V_���۽��,
    -V_���,V_ժҪ,������_IN,��������_IN,������_IN,��������_IN,V_������,V_��׼�ĺ�);

    --ԭ�����ֲ�������ҩƷ,��C����ʱ��Ҫ������
    BEGIN 
        SELECT COUNT(*) INTO V_��¼��
        FROM ҩƷ�շ���¼ A, ҩƷ��� B
        WHERE B.ҩƷID=ҩƷID_IN
        AND A.NO=NO_IN
        AND A.���� = 7 
        AND Mod(A.��¼״̬,3)=0
        AND NVL(A.����,0)>0
        AND (NVL(B.ҩ�����,0)=0 OR 
            (NVL(B.ҩ������,0)=0 AND A.�ⷿID IN (SELECT ����ID FROM  ��������˵�� WHERE (�������� LIKE '%ҩ��') OR (�������� LIKE '�Ƽ���'))))
        ;
    EXCEPTION 
        WHEN OTHERS THEN V_��¼��:=0;
    END;
    IF V_��¼��>0 THEN
        V_����:=0;
    ELSE
        V_����:=NVL (V_����, 0);
    END IF;

    --����ҩƷ�������Ӧ����
    UPDATE ҩƷ���
       SET �������� = NVL (��������, 0) + NVL (��������_IN, 0),
           ʵ������ = NVL (ʵ������, 0) + NVL (��������_IN, 0),
           ʵ�ʽ�� = NVL (ʵ�ʽ��, 0) + NVL (V_���۽��, 0),
           ʵ�ʲ�� = NVL (ʵ�ʲ��, 0) + NVL (V_���, 0)
     WHERE �ⷿID = V_�ⷿID
       AND ҩƷID = ҩƷID_IN
       AND NVL (����, 0) = V_����
       AND ���� = 1;

    IF SQL%NOTFOUND THEN
        INSERT INTO ҩƷ���
        (�ⷿID,ҩƷID,����,����,��������,ʵ������,
            ʵ�ʽ��,ʵ�ʲ��,�ϴ�����,Ч��,�ϴβ���,��׼�ĺ�)
        VALUES 
        (V_�ⷿID,ҩƷID_IN,V_����,1,��������_IN,��������_IN,
        V_���۽��,V_���,V_����,V_Ч��,v_����,V_��׼�ĺ�);
    END IF;

    DELETE ҩƷ���
     WHERE �ⷿID = V_�ⷿID
        AND ҩƷID = ҩƷID_IN
        AND NVL (��������, 0) = 0
        AND NVL (ʵ������, 0) = 0
        AND NVL (ʵ�ʽ��, 0) = 0
        AND NVL (ʵ�ʲ��, 0) = 0;

    --����ҩƷ�շ����ܱ����Ӧ����

    UPDATE ҩƷ�շ�����
        SET ���� = NVL (����, 0)  +NVL(��������_IN, 0),
             ��� = NVL (���, 0) +NVL(V_���۽��, 0),
             ��� = NVL (���, 0) +NVL(V_���, 0)
     WHERE ���� = TRUNC (��������_IN)
        AND �ⷿID = V_�ⷿID
        AND ҩƷID = ҩƷID_IN
        AND ���ID = V_������ID
        AND ���� = 7;

    IF SQL%NOTFOUND THEN
        INSERT INTO ҩƷ�շ�����
        (����, �ⷿID, ҩƷID, ���ID, ����, ����, ���, ���)
        VALUES 
        (TRUNC (��������_IN),V_�ⷿID,ҩƷID_IN,V_������ID,7,��������_IN,V_���۽��,V_���);
    END IF;
EXCEPTION
    WHEN ERR_ISSTRIKED THEN
        RAISE_APPLICATION_ERROR (-20101, '[ZLSOFT]�õ����Ѿ������˳�����[ZLSOFT]');
    WHEN ERR_ISBATCH THEN
        RAISE_APPLICATION_ERROR (-20102, '[ZLSOFT]�õ����а�����һ��ԭ�������������ڷ�����ҩƷ�����ܳ�����[ZLSOFT]');  
    WHEN OTHERS THEN ZL_ERRORCENTER (SQLCODE, SQLERRM);
END ZL_ҩƷ����_STRIKE;
/

CREATE OR REPLACE PROCEDURE ZL_ҩƷ�ƿ�_INSERT(
	NO_IN         IN ҩƷ�շ���¼.NO%TYPE,
	���_IN       IN ҩƷ�շ���¼.���%TYPE,
	�ⷿID_IN     IN ҩƷ�շ���¼.�ⷿID%TYPE,
	�Է�����ID_IN IN ҩƷ�շ���¼.�Է�����ID%TYPE,
	ҩƷID_IN     IN ҩƷ�շ���¼.ҩƷID%TYPE,
	����_IN       IN ҩƷ�շ���¼.����%TYPE,
	��д����_IN   IN ҩƷ�շ���¼.��д����%TYPE,
	ʵ������_IN   IN ҩƷ�շ���¼.ʵ������%TYPE,
	�ɱ���_IN     IN ҩƷ�շ���¼.�ɱ���%TYPE,
	�ɱ����_IN   IN ҩƷ�շ���¼.�ɱ����%TYPE,
	���ۼ�_IN     IN ҩƷ�շ���¼.���ۼ�%TYPE,
	���۽��_IN   IN ҩƷ�շ���¼.���۽��%TYPE,
	���_IN       IN ҩƷ�շ���¼.���%TYPE,
	������_IN     IN ҩƷ�շ���¼.������%TYPE,
	����_IN       IN ҩƷ�շ���¼.����%TYPE := NULL,
	����_IN       IN ҩƷ�շ���¼.����%TYPE := NULL,
	Ч��_IN       IN ҩƷ�շ���¼.Ч��%TYPE := NULL,
	ժҪ_IN       IN ҩƷ�շ���¼.ժҪ%TYPE := NULL,
	��������_IN   IN ҩƷ�շ���¼.��������%TYPE := NULL,
  �ϴι�Ӧ��ID_IN In ҩƷ�շ���¼.��ҩ��λID%TYPE := Null,
  ��׼�ĺ�_IN In ҩƷ�շ���¼.��׼�ĺ�%TYPE := Null
	)
IS
	ERR_MutilROW EXCEPTION ;
	ERR_ISNONUMBER EXCEPTION;
	V_����       �շ���ĿĿ¼.����%TYPE;
  V_LNGID      ҩƷ�շ���¼.ID%TYPE; --�շ�ID
	V_������ID ҩƷ�շ���¼.������ID%TYPE; --������ID
	V_�������ID ҩƷ�շ���¼.������ID%TYPE; --������ID
	V_����       ҩƷ�շ���¼.����%TYPE := NULL; --��Ҫ��������ʵ��ҩ�������ҩƷ
	V_�Ƿ����   INTEGER; --�ж�����Ƿ�ҩ�����   1:������0��������
	V_ҩ�����   INTEGER; --�ж�����Ƿ�ҩ�����   1:������0��������
	V_ҩ������   INTEGER; --�ж�����Ƿ�ҩ�����   1:������0��������
	intRecords NUMBER ;
  V_�¿��ÿ�� ϵͳ������.����ֵ%Type;
BEGIN
	SELECT B.ID
	INTO V_������ID
	FROM ҩƷ�������� A, ҩƷ������ B
	WHERE A.���ID = B.ID AND A.���� = 6 AND B.ϵ�� = 1 AND ROWNUM < 2;
	SELECT B.ID
	INTO V_�������ID
	FROM ҩƷ�������� A, ҩƷ������ B
	WHERE A.���ID = B.ID AND A.���� = 6 AND B.ϵ�� = -1 AND ROWNUM < 2;

  INSERT INTO ҩƷ�շ���¼
      (ID,��¼״̬,����,NO,���,�ⷿID,�Է�����ID,������ID,���ϵ��,
       ҩƷID,����,����,����,Ч��,��д����,ʵ������,
    �ɱ���,�ɱ����,���ۼ�,���۽��,���,
       ժҪ,������,��������,��ҩ��λID,��׼�ĺ�)
  VALUES
      (ҩƷ�շ���¼_ID.NEXTVAL,1,6,NO_IN,���_IN,�ⷿID_IN,�Է�����ID_IN,V_�������ID,-1,
       ҩƷID_IN,����_IN,����_IN,����_IN,Ч��_IN,��д����_IN,ʵ������_IN,
 �ɱ���_IN,�ɱ����_IN,���ۼ�_IN,���۽��_IN,���_IN,
       ժҪ_IN,������_IN,��������_IN,�ϴι�Ӧ��ID_IN,��׼�ĺ�_IN);

  SELECT NVL(ҩ�����, 0), NVL(ҩ������, 0)
	INTO V_ҩ�����, V_ҩ������
	FROM ҩƷ���
	WHERE ҩƷID = ҩƷID_IN;

  V_�Ƿ���� := 0;
  IF V_ҩ������ = 0 THEN
      IF V_ҩ����� = 1 THEN
	BEGIN
		SELECT DISTINCT 0
		INTO V_�Ƿ����
		FROM ��������˵��
		WHERE ((�������� LIKE '%ҩ��') OR (�������� LIKE '�Ƽ���'))
		AND ����ID = �Է�����ID_IN;
	EXCEPTION
		WHEN OTHERS THEN V_�Ƿ���� := 1;
	END;
      END IF;
  ELSE
      V_�Ƿ���� := 1;
  END IF;

  SELECT ҩƷ�շ���¼_ID.NEXTVAL INTO V_LNGID FROM DUAL;

  IF V_�Ƿ���� = 1 AND NVL(����_IN, 0) = 0 THEN
      --�������ҳ��ⲻ����
      V_���� := V_LNGID;
  ELSIF V_�Ƿ���� = 0 THEN
      --��ⲻ����
      V_���� := 0;
  ELSIF NVL(����_IN, 0) <> 0 THEN
      --�������ҳ���Ҳ����
      V_���� := ����_IN;
  END IF;

  INSERT INTO ҩƷ�շ���¼
      (ID,��¼״̬,����,NO,���,�ⷿID,�Է�����ID,������ID,���ϵ��,
       ҩƷID,����,����,����,Ч��,��д����,ʵ������,
   �ɱ���,�ɱ����,���ۼ�,���۽��,���,
       ժҪ,������,��������,��ҩ��λID,��׼�ĺ�)
  VALUES
      (V_LNGID,1,6,NO_IN,���_IN + 1,�Է�����ID_IN,�ⷿID_IN,V_������ID,1,
       ҩƷID_IN,V_����,����_IN,����_IN,Ч��_IN,��д����_IN,ʵ������_IN,
       �ɱ���_IN,�ɱ����_IN,���ۼ�_IN,���۽��_IN,���_IN,
       ժҪ_IN,������_IN,��������_IN,�ϴι�Ӧ��ID_IN,��׼�ĺ�_IN);

	--����Ƿ������ͬҩƷ��ͬ���ε����ݣ�������ڲ�������
	SELECT COUNT(*) INTO intRecords
	FROM ҩƷ�շ���¼
	WHERE ����=6 AND NO=NO_IN AND ���ϵ��=-1 AND ҩƷID+0=ҩƷID_IN AND Nvl(����,0)=NVL(����_IN,0);
	IF intRecords>1 THEN
		RAISE ERR_MutilROW;
	END IF ;
  
  --���ݲ��������Ƿ��·�ҩ�ⷿ�Ŀ��ÿ��
	Select ����ֵ Into v_�¿��ÿ�� From ϵͳ������ Where ������ = 96;
  
  --����Ϊ1��ʾ���ʱ�¿�������
	If v_�¿��ÿ�� = '1' Then
      UPDATE ҩƷ���
			SET ��������=NVL(��������,0)-ʵ������_IN
			WHERE �ⷿID=�ⷿID_IN AND ҩƷID=ҩƷID_IN AND NVL(����,0)=����_IN AND ����=1;
      
			IF SQL%ROWCOUNT=0 THEN 
				INSERT INTO ҩƷ���(�ⷿID,ҩƷID,����,����,��������,�ϴ�����,Ч��,�ϴβ���)
				VALUES (�ⷿID_IN,ҩƷID_IN,����_IN,1,-1*ʵ������_IN,
					����_IN,Ч��_IN,����_IN);
			END IF ;

			DELETE
			FROM ҩƷ���
			WHERE �ⷿID = �ⷿID_IN
			AND ҩƷID = ҩƷID_IN
			AND NVL(��������,0)=0
			AND NVL(ʵ������,0)=0
			AND NVL(ʵ�ʽ��,0)=0
			AND NVL(ʵ�ʲ��,0)=0;
  End If;
EXCEPTION
    WHEN ERR_ISNONUMBER THEN
        SELECT ���� INTO V_���� FROM �շ���ĿĿ¼ WHERE ID = ҩƷID_IN;
        RAISE_APPLICATION_ERROR(-20101,'[ZLSOFT]����Ϊ'||V_����||',����Ϊ'||
			����_IN||'��ҩ�����ҩƷ'||CHR(10) ||CHR(13)||'���ÿ������������[ZLSOFT]');
	WHEN ERR_MutilROW THEN
		SELECT ���� INTO V_���� FROM �շ���ĿĿ¼ WHERE ID = ҩƷID_IN;
		RAISE_APPLICATION_ERROR(-20101,'[ZLSOFT]����Ϊ'||V_����||'��ҩƷ�����ڶ����ظ��ļ�¼����ϲ�Ϊһ����¼��[ZLSOFT]');
    WHEN OTHERS THEN
        ZL_ERRORCENTER(SQLCODE, SQLERRM);
END ZL_ҩƷ�ƿ�_INSERT;
/


CREATE OR REPLACE PROCEDURE ZL_ҩƷ�ƿ�_DELETE (
    --ɾ��ҩƷ�շ���¼����Ӧ�ı�ҩƷ���
    NO_IN IN ҩƷ�շ���¼.NO%TYPE
)
IS
	V_���� ҩƷ�շ���¼.��ҩ����%TYPE;
	ERR_ISVERIFIED EXCEPTION;
  V_�¿��ÿ�� ϵͳ������.����ֵ%Type;

    CURSOR C_ҩƷ�շ���¼
    IS
		SELECT ʵ������, �ⷿID, ����, ҩƷID,����,Ч��,����,��ҩ��λid,��׼�ĺ�
		FROM ҩƷ�շ���¼
		WHERE NO = NO_IN
		AND ���� = 6
		AND ���ϵ�� = -1
		ORDER BY ҩƷID;
BEGIN
	--����Ƿ��ѷ��ͣ��ѷ��͵ĵ�����Ҫ��ԭ��������
	SELECT ��ҩ���� INTO V_����
	FROM ҩƷ�շ���¼
	WHERE ����=6 AND NO=NO_IN AND ROWNUM<2;

	Select ����ֵ Into v_�¿��ÿ�� From ϵͳ������ Where ������ = 96;  
  
  --�������ֵΪ1ҲҪ�ָ�ԭ���Ŀ�������
	IF V_���� IS NOT NULL Or v_�¿��ÿ��='1' THEN
		--ͨ��ѭ�����ָ�ԭ���Ŀ�������
		FOR V_ҩƷ�շ���¼ IN C_ҩƷ�շ���¼ LOOP
			UPDATE ҩƷ���
			SET �������� = NVL (��������, 0) + V_ҩƷ�շ���¼.ʵ������
			WHERE �ⷿID = V_ҩƷ�շ���¼.�ⷿID
			AND ҩƷID = V_ҩƷ�շ���¼.ҩƷID
			AND NVL (����, 0) = NVL (V_ҩƷ�շ���¼.����, 0)
			AND ���� = 1;

			IF SQL%NOTFOUND THEN
				INSERT INTO ҩƷ���
					(�ⷿID, ҩƷID, ����, ����, ��������,�ϴ�����,Ч��,�ϴβ���,�ϴι�Ӧ��Id,��׼�ĺ�)
				VALUES (
					V_ҩƷ�շ���¼.�ⷿID,V_ҩƷ�շ���¼.ҩƷID,V_ҩƷ�շ���¼.����,1,
					V_ҩƷ�շ���¼.ʵ������,V_ҩƷ�շ���¼.����,V_ҩƷ�շ���¼.Ч��,V_ҩƷ�շ���¼.����,V_ҩƷ�շ���¼.��ҩ��λid,V_ҩƷ�շ���¼.��׼�ĺ�
					);
			END IF;

			DELETE
			FROM ҩƷ���
			WHERE �ⷿID = V_ҩƷ�շ���¼.�ⷿID
			AND ҩƷID = V_ҩƷ�շ���¼.ҩƷID
			AND NVL(��������,0) = 0
			AND NVL(ʵ������,0) = 0
			AND NVL(ʵ�ʽ��,0) = 0
			AND NVL(ʵ�ʲ��,0) = 0;
		END LOOP;
	END IF ;

    DELETE--����ͳ����������ƿⵥ��ɾ��
	FROM ҩƷ�շ���¼
	WHERE NO = NO_IN
	AND ���� = 6
	AND ��¼״̬ = 1
	AND ����� IS NULL;

    IF SQL%ROWCOUNT = 0 THEN
        RAISE ERR_ISVERIFIED;
    END IF;
EXCEPTION
    WHEN ERR_ISVERIFIED THEN
        RAISE_APPLICATION_ERROR (
            -20101, '[ZLSOFT]�õ����Ѿ�������ɾ�����ѱ�����ˣ�[ZLSOFT]');
    WHEN OTHERS THEN
        ZL_ERRORCENTER (SQLCODE, SQLERRM);
END ZL_ҩƷ�ƿ�_DELETE;
/


CREATE OR REPLACE PROCEDURE ZL_ҩƷ�ƿ�_VERIFY (
    ���_IN IN ҩƷ�շ���¼.���%TYPE,
    �ⷿID_IN IN ҩƷ�շ���¼.�ⷿID%TYPE,
    �Է�����ID_IN IN ҩƷ�շ���¼.�Է�����ID%TYPE,
    ҩƷID_IN IN ҩƷ�շ���¼.ҩƷID%TYPE,
    ����_IN IN ҩƷ�շ���¼.����%TYPE,
    ������_IN IN ҩƷ�շ���¼.����%TYPE,
    ��д����_IN IN ҩƷ�շ���¼.��д����%TYPE,
    ʵ������_IN IN ҩƷ�շ���¼.ʵ������%TYPE,
    �ɱ���_IN IN ҩƷ�շ���¼.�ɱ���%TYPE,
    �ɱ����_IN IN ҩƷ�շ���¼.�ɱ����%TYPE,
    ���۽��_IN IN ҩƷ�շ���¼.���۽��%TYPE,
    ���_IN IN ҩƷ�շ���¼.���%TYPE,
    �����ID_IN IN ҩƷ�շ���¼.������ID%TYPE,
    �����ID_IN IN ҩƷ�շ���¼.������ID%TYPE,
    NO_IN IN ҩƷ�շ���¼.NO%TYPE,
    �����_IN IN ҩƷ�շ���¼.�����%TYPE,
    ����_IN IN ҩƷ�շ���¼.����%TYPE := NULL,
    Ч��_IN IN ҩƷ�շ���¼.Ч��%TYPE := NULL,
    �������_IN IN ҩƷ�շ���¼.�������%TYPE := NULL,
    �ƿⵥ_IN IN NUMBER:=1,
    �ϴι�Ӧ��ID_IN In ҩƷ�շ���¼.��ҩ��λID%TYPE := Null,
    ��׼�ĺ�_IN In ҩƷ�շ���¼.��׼�ĺ�%TYPE := Null
)
IS
    ERR_ISVERIFIED EXCEPTION;
    ERR_ISNONUMBER EXCEPTION;
    V_������ ҩƷ�շ���¼.����%TYPE := NULL;
    V_ʵ�ʿ���� ҩƷ���.ʵ�ʽ��%TYPE;
    V_ʵ�ʿ���� ҩƷ���.ʵ�ʲ��%TYPE;
    V_����� NUMBER(18,8);
    V_������ ҩƷ���.ʵ�ʲ��%TYPE;
    V_�ɱ��� ҩƷ�շ���¼.�ɱ���%TYPE;
    V_�ɱ���� ҩƷ�շ���¼.�ɱ����%TYPE;
    V_ʵ������ ҩƷ���.ʵ������%TYPE;
    V_���� �շ���ĿĿ¼.����%TYPE;
	INTDIGIT NUMBER ;
BEGIN
	--��ȡ���С��λ��
    SELECT NVL(����ֵ,ȱʡֵ) INTO INTDIGIT FROM ϵͳ������ WHERE ������='���ý���λ��';

	--�����ƿ⴦�����������ʱ�ı�ʵ��������
      --�������ȶ�ʵ��������������Ӧ���ֶν��и��¡�
    BEGIN
		SELECT NVL(ʵ�ʽ��,0), NVL(ʵ�ʲ��,0), NVL(ʵ������,0)
		INTO V_ʵ�ʿ����, V_ʵ�ʿ����, V_ʵ������
		FROM ҩƷ���
		WHERE ҩƷID = ҩƷID_IN
		AND NVL (����, 0) = ������_IN
		AND �ⷿID = �ⷿID_IN
		AND ���� = 1
		AND ROWNUM = 1;
    EXCEPTION
        WHEN OTHERS THEN
            V_ʵ�ʿ���� := 0;
            V_ʵ������ := 0;
    END;

    IF V_ʵ�ʿ���� <= 0 THEN
        BEGIN
			SELECT ָ������� / 100
			INTO V_�����
			FROM ҩƷ���
			WHERE ҩƷID = ҩƷID_IN;
        EXCEPTION
            WHEN OTHERS THEN
                V_����� := 0;
        END;
    ELSE
        V_����� := V_ʵ�ʿ���� / V_ʵ�ʿ����;
    END IF;

    V_������ := ROUND(���۽��_IN * V_�����,INTDIGIT);
    IF ʵ������_IN=0 THEN
        V_�ɱ��� :=�ɱ���_IN;
    ELSE
        V_�ɱ��� := (���۽��_IN - V_������) / ʵ������_IN; 
    END IF; 
    V_�ɱ���� := ROUND(V_�ɱ��� * ʵ������_IN,INTDIGIT);

    UPDATE ҩƷ�շ���¼
        SET ����� = NVL (�����_IN, �����),
             ������� = �������_IN,
             ʵ������ = ʵ������_IN,
             �ɱ��� = V_�ɱ���,
             �ɱ���� = V_�ɱ����,
             ���۽�� = ���۽��_IN,
             ��� = V_������
     WHERE NO = NO_IN
        AND ���� = 6
        AND ҩƷID = ҩƷID_IN
        AND ��¼״̬ = 1
        AND ��� IN (���_IN, ���_IN + 1)
        AND ����� IS NULL;

    IF SQL%ROWCOUNT = 0 THEN
        RAISE ERR_ISVERIFIED;
    END IF;

    IF ������_IN > 0 THEN
        IF V_ʵ������ < ʵ������_IN THEN
            RAISE ERR_ISNONUMBER;
        END IF;
    END IF;

    --ȡ����������
	SELECT ����
	INTO V_������
	FROM ҩƷ�շ���¼
	WHERE NO = NO_IN
	AND ���� = 6
	AND ��¼״̬ = 1
	AND ��� = ���_IN+1;
        
    --����������ҩƷ������Ӧ����

	UPDATE ҩƷ���
	SET �������� = NVL (��������, 0) + ʵ������_IN,
		ʵ������ = NVL (ʵ������, 0) + ʵ������_IN,
		ʵ�ʽ�� = NVL (ʵ�ʽ��, 0) + ���۽��_IN,
		ʵ�ʲ�� = NVL (ʵ�ʲ��, 0) + V_������,
		�ϴβɹ��� = V_�ɱ���,
		�ϴ����� = NVL (����_IN, �ϴ�����),
		�ϴβ��� = NVL (����_IN, �ϴβ���),
		Ч�� = NVL (Ч��_IN, Ч��),
    �ϴι�Ӧ��ID=�ϴι�Ӧ��ID_IN,
    ��׼�ĺ�=��׼�ĺ�_IN
	WHERE �ⷿID = �Է�����ID_IN
	AND ҩƷID = ҩƷID_IN
	AND NVL (����, 0) = NVL (V_������, 0)
	AND ���� = 1;

    IF SQL%NOTFOUND THEN
        INSERT INTO ҩƷ���
			(�ⷿID,ҩƷID,����,����,��������,ʵ������,ʵ�ʽ��,ʵ�ʲ��,�ϴβɹ���,�ϴ�����,�ϴβ���,Ч��,�ϴι�Ӧ��ID,��׼�ĺ�)
        VALUES (
			�Է�����ID_IN,ҩƷID_IN,V_������,1,ʵ������_IN,ʵ������_IN,���۽��_IN,
            V_������,V_�ɱ���,����_IN,����_IN,Ч��_IN,�ϴι�Ӧ��ID_IN,��׼�ĺ�_IN);
    END IF;

    --���ĳ�����ҩƷ������Ӧ����

    UPDATE ҩƷ���
	SET 
		ʵ������ = NVL (ʵ������, 0) - ʵ������_IN,
		ʵ�ʽ�� = NVL (ʵ�ʽ��, 0) - ���۽��_IN,
		ʵ�ʲ�� = NVL (ʵ�ʲ��, 0) - V_������
	WHERE �ⷿID = �ⷿID_IN
	AND ҩƷID = ҩƷID_IN
	AND NVL (����, 0) = NVL (������_IN, 0)
	AND ���� = 1;

    IF SQL%NOTFOUND THEN
        INSERT INTO ҩƷ���(�ⷿID,ҩƷID,����,����,��������,ʵ������,ʵ�ʽ��,ʵ�ʲ��,�ϴ�����,Ч��,�ϴι�Ӧ��ID,��׼�ĺ�)
        VALUES (�ⷿID_IN,ҩƷID_IN,������_IN,1,0,-ʵ������_IN,-���۽��_IN,-V_������,����_IN,Ч��_IN,�ϴι�Ӧ��ID_IN,��׼�ĺ�_IN);
    END IF;

	DELETE
	FROM ҩƷ���
	WHERE �ⷿID = �ⷿID_IN
	AND ҩƷID = ҩƷID_IN
	AND NVL (��������, 0) = 0
	AND NVL (ʵ������, 0) = 0
	AND NVL (ʵ�ʽ��, 0) = 0
	AND NVL (ʵ�ʲ��, 0) = 0;

    --����������ҩƷ�շ����ܱ����Ӧ����
	UPDATE ҩƷ�շ�����
	SET ���� = NVL (����, 0) + ʵ������_IN,
		��� = NVL (���, 0) + ���۽��_IN,
		��� = NVL (���, 0) + V_������
	WHERE ���� = TRUNC (SYSDATE)
	AND �ⷿID = �Է�����ID_IN
	AND ҩƷID = ҩƷID_IN
	AND ���ID = �����ID_IN
	AND ���� = 6;

    IF SQL%NOTFOUND THEN
		INSERT INTO ҩƷ�շ�����
			(����, �ⷿID, ҩƷID, ���ID, ����, ����, ���, ���)
		VALUES(
			TRUNC(SYSDATE),�Է�����ID_IN,ҩƷID_IN,�����ID_IN,6,ʵ������_IN,���۽��_IN,V_������);
    END IF;

    --���ĳ�����ҩƷ�շ����ܱ����Ӧ����
    UPDATE ҩƷ�շ�����
	SET ���� = NVL (����, 0) - ʵ������_IN,
		��� = NVL (���, 0) - ���۽��_IN,
		��� = NVL (���, 0) - V_������
	WHERE ���� = TRUNC (SYSDATE)
	AND �ⷿID = �ⷿID_IN
	AND ҩƷID = ҩƷID_IN
	AND ���ID = �����ID_IN
	AND ���� = 6;

    IF SQL%NOTFOUND THEN
        INSERT INTO ҩƷ�շ�����
			(����, �ⷿID, ҩƷID, ���ID, ����, ����, ���, ���)
		VALUES (
			TRUNC(SYSDATE),�ⷿID_IN,ҩƷID_IN,�����ID_IN,6,-ʵ������_IN,-���۽��_IN,-V_������);
    END IF;
EXCEPTION
    WHEN ERR_ISVERIFIED THEN
        RAISE_APPLICATION_ERROR (-20101, '[ZLSOFT]�õ����Ѿ���������ˣ�[ZLSOFT]');
    WHEN ERR_ISNONUMBER THEN
		SELECT ����
		INTO V_����
		FROM �շ���ĿĿ¼
		WHERE ID = ҩƷID_IN;
        RAISE_APPLICATION_ERROR (
            -20101, '[ZLSOFT]����Ϊ' || V_���� || ',����Ϊ' || ����_IN ||
                          '��ҩ�����ҩƷ' || CHR(10) || CHR(13) || '���ÿ������������[ZLSOFT]');
    WHEN OTHERS THEN
        ZL_ERRORCENTER (SQLCODE, SQLERRM);
END ZL_ҩƷ�ƿ�_VERIFY;
/


CREATE OR REPLACE PROCEDURE ZL_ҩƷ�ƿ�_STRIKE (
    �д�_IN IN INTEGER,
    ԭ��¼״̬_IN IN ҩƷ�շ���¼.��¼״̬%TYPE,
    NO_IN IN ҩƷ�շ���¼.NO%TYPE,
    ���_IN IN ҩƷ�շ���¼.���%TYPE,
    ҩƷID_IN IN ҩƷ�շ���¼.ҩƷID%TYPE,
    ��������_IN IN ҩƷ�շ���¼.ʵ������%TYPE,
    ������_IN IN ҩƷ�շ���¼.������%TYPE,
    ��������_IN IN ҩƷ�շ���¼.��������%TYPE
)
IS
    ERR_ISSTRIKED EXCEPTION;
    ERR_ISOUTSTOCK EXCEPTION;
    ERR_ISNONUM EXCEPTION;
    ERR_ISBATCH EXCEPTION;
    V_BATCHCOUNT INTEGER;    --ԭ���������ڷ�����ҩƷ������
    V_��� ҩƷ�շ���¼.���%TYPE;
    V_�ⷿID ҩƷ�շ���¼.�ⷿID%TYPE;
    V_�Է�����ID ҩƷ�շ���¼.�Է�����ID%TYPE;
    V_������ID ҩƷ�շ���¼.������ID%TYPE ;
    V_���� ҩƷ�շ���¼.����%TYPE ;
    V_���� ҩƷ�շ���¼.����%TYPE ;
    V_���� ҩƷ�շ���¼.����%TYPE ;
    V_Ч�� ҩƷ�շ���¼.Ч��%TYPE ;
    V_�ɱ��� ҩƷ�շ���¼.�ɱ���%TYPE ;
    V_�ɱ���� ҩƷ�շ���¼.�ɱ����%TYPE ;
    V_���� ҩƷ�շ���¼.����%TYPE ;
    V_���ۼ� ҩƷ�շ���¼.���ۼ�%TYPE ;
    V_���۽�� ҩƷ�շ���¼.���۽��%TYPE ;
    V_��� ҩƷ�շ���¼.���%TYPE ;
    V_ժҪ ҩƷ�շ���¼.ժҪ%TYPE ;
    V_ʣ������ ҩƷ�շ���¼.ʵ������%TYPE;
    V_ʣ��ɱ���� ҩƷ�շ���¼.�ɱ����%Type;
    V_ʣ�����۽�� ҩƷ�շ���¼.���۽��%Type;
    V_���ϵ�� ҩƷ�շ���¼.���ϵ��%TYPE;
    V_��¼�� NUMBER;
    V_�շ�ID ҩƷ�շ���¼.ID%TYPE;
    V_��ҩ�� ҩƷ�շ���¼.��ҩ��%TYPE;
    V_�������� ҩƷ�շ���¼.��ҩ����%TYPE;
    V_�ϴι�Ӧ��ID ҩƷ�շ���¼.��ҩ��λid%Type;
    V_��׼�ĺ� ҩƷ�շ���¼.��׼�ĺ�%TYPE;
 
    --�Գ����������м��
    V_����� ҩƷ���.ʵ������%TYPE;
    V_ҩ����� INTEGER;
    V_ҩ������ INTEGER;
    V_�������� INTEGER;
    V_ҩ�� INTEGER;
    V_���� NUMBER;
	INTDIGIT NUMBER;
 
    CURSOR C_ҩƷ�շ���¼
    IS
    SELECT ���,�ⷿID,�Է�����ID,������ID,���ϵ��,ҩƷID,����,����,����,Ч��,��ҩ��,��ҩ����,ժҪ,��ҩ��λID,��׼�ĺ�
    FROM ҩƷ�շ���¼
    WHERE NO = NO_IN AND ���� = 6 AND (���>=���_IN AND ���<=���_IN+1) AND (��¼״̬=1 OR MOD(��¼״̬,3)=0)
    ORDER BY ҩƷID;
BEGIN
         --��ȡ���С��λ��
    SELECT NVL(����ֵ,ȱʡֵ) INTO INTDIGIT FROM ϵͳ������ WHERE ������='���ý���λ��';
 
    IF �д�_IN =1 THEN
        UPDATE ҩƷ�շ���¼
        SET ��¼״̬ = DECODE(ԭ��¼״̬_IN,1,3,ԭ��¼״̬_IN+3)
        WHERE NO = NO_IN AND ���� = 6 AND ��¼״̬ =ԭ��¼״̬_IN ;
        IF SQL%ROWCOUNT = 0 THEN
            RAISE ERR_ISSTRIKED;
        END IF;
    END IF;
 
    SELECT COUNT(*) INTO V_BATCHCOUNT FROM
            ҩƷ�շ���¼ A,ҩƷ��� B
    WHERE A.ҩƷID=B.ҩƷID
        AND A.NO=NO_IN
        AND A.����=6
        AND A.ҩƷID+0=ҩƷID_IN
        AND MOD(A.��¼״̬,3)=0
        AND NVL(A.����,0)=0
        AND ((NVL(B.ҩ�����,0)=1 AND A.�ⷿID NOT IN (SELECT ����ID FROM  ��������˵�� WHERE (�������� LIKE '%ҩ��') OR (�������� LIKE '�Ƽ���')))
            OR NVL(B.ҩ������,0)=1);
 
    IF V_BATCHCOUNT>0 THEN
        RAISE ERR_ISBATCH;
    END IF;
    
    SELECT SUM(A.ʵ������) AS ʣ������,SUM(A.�ɱ����) AS ʣ��ɱ����,SUM(A.���۽��) AS ʣ�����۽��,A.�ɱ���,A.���ۼ�,A.�Է�����ID,NVL(A.����,0),B.ҩ�����,B.ҩ������,A.��׼�ĺ�
    INTO  V_ʣ������,V_ʣ��ɱ����,V_ʣ�����۽��,V_�ɱ���,V_���ۼ�,V_�ⷿID,V_����,V_ҩ�����,V_ҩ������,V_��׼�ĺ�
    FROM ҩƷ�շ���¼ A,ҩƷ��� B
    WHERE A.NO=NO_IN AND A.ҩƷID=B.ҩƷID AND A.����=6 AND A.ҩƷID=ҩƷID_IN AND A.���=���_IN
    GROUP BY A.�ɱ���,A.���ۼ�,A.�Է�����ID,NVL(A.����,0),B.ҩ�����,B.ҩ������,A.��׼�ĺ�;
 
    --�жϸò�����ҩ�⻹��ҩ��
    BEGIN
        SELECT DISTINCT 0
        INTO V_ҩ��
        FROM ��������˵��
        WHERE (   (�������� LIKE '%ҩ��')
              OR (�������� LIKE '�Ƽ���'))
        AND ����ID = V_�ⷿID;
    EXCEPTION
        WHEN OTHERS THEN V_ҩ��:=1;
    END ;
 
    --���ݲ�������,�жϷ�������
    IF V_ҩ��=0 THEN
        V_��������:=V_ҩ������;
    ELSE
        V_��������:=V_ҩ�����;
    END IF ;
 
    --V_����:(ԭ����,�ַ���Ϊ����;����Ϊ��)ԭ������,�ַ���,�����̲�����
    --��Ϊ���ڶԷ��ⷿ���൱�ڳ��⣬�����ڵ�ǰ�ⷿ���൱����⣬���Ե�ǰ�ⷿ�����飬������˿�������¼
    SELECT NVL(A.����,0) INTO V_����
    FROM ҩƷ�շ���¼ A
    WHERE A.NO=NO_IN AND A.����=6 AND A.ҩƷID=ҩƷID_IN AND A.���=���_IN+1 AND MOD(A.��¼״̬,3)=0;
 
    --ȡ�����
    BEGIN
        SELECT NVL(ʵ������,0) INTO V_����� FROM ҩƷ���
        WHERE �ⷿID=V_�ⷿID AND ҩƷID=ҩƷID_IN AND NVL(����,0)=V_���� AND ����=1;
    EXCEPTION
        WHEN OTHERS THEN V_�����:=0;
    END ;
 
    --������������ʣ������,ȡʣ������;����ȡ�����
    IF V_�����<V_ʣ������ Then
       v_ʣ��ɱ����:=V_�����/V_ʣ������*v_ʣ��ɱ����;
       V_ʣ�����۽��:=V_�����/V_ʣ������*V_ʣ�����۽��;
       V_ʣ������:=V_�����;
    END IF ;
 
    --������������ʣ��������������
    IF V_ʣ������<��������_IN THEN
        RAISE ERR_ISNONUM;
    END IF;
 
    V_�ɱ����:= ROUND(��������_IN/v_ʣ������*V_ʣ��ɱ����,INTDIGIT);
    V_���۽��:= ROUND(��������_IN/v_ʣ������*V_ʣ�����۽��,INTDIGIT);
    V_���:=V_���۽��-V_�ɱ����;
 
    FOR V_ҩƷ�շ���¼ IN C_ҩƷ�շ���¼ LOOP
        V_���:=V_ҩƷ�շ���¼.���;
        V_�ⷿID:=V_ҩƷ�շ���¼.�ⷿID;
        V_�Է�����ID:=V_ҩƷ�շ���¼.�Է�����ID;
        V_������ID:=V_ҩƷ�շ���¼.������ID;
        V_���ϵ��:=V_ҩƷ�շ���¼.���ϵ��;
        V_����:=V_ҩƷ�շ���¼.����;
        V_����:=V_ҩƷ�շ���¼.����;
        V_����:=V_ҩƷ�շ���¼.����;
        V_Ч��:=V_ҩƷ�շ���¼.Ч��;
         v_ժҪ:=v_ҩƷ�շ���¼.ժҪ;
         V_��ҩ��:=V_ҩƷ�շ���¼.��ҩ��;
         V_��������:=V_ҩƷ�շ���¼.��ҩ����;
         V_�ϴι�Ӧ��ID:=V_ҩƷ�շ���¼.��ҩ��λID;
         V_��׼�ĺ�:=V_ҩƷ�շ���¼.��׼�ĺ�;
 
        SELECT ҩƷ�շ���¼_ID.NEXTVAL INTO V_�շ�ID  FROM DUAL;
        INSERT INTO ҩƷ�շ���¼
        (ID,��¼״̬,����,NO,���,�ⷿID,�Է�����ID,������ID,���ϵ��,
        ҩƷID,����,����,����,Ч��,��д����,ʵ������,�ɱ���,
        �ɱ����,���ۼ�,���۽��,���,ժҪ,������,��������,�����,�������,��ҩ��,��ҩ����,��ҩ��λID,��׼�ĺ�)
        VALUES
        (V_�շ�ID,DECODE(ԭ��¼״̬_IN,1,2,ԭ��¼״̬_IN+2),6,NO_IN,V_���,V_�ⷿID,V_�Է�����ID,V_������ID,V_���ϵ��,
        ҩƷID_IN,V_����,V_����,V_����,V_Ч��,-��������_IN,-��������_IN,V_�ɱ���,-V_�ɱ����,V_���ۼ�,-V_���۽��,
        -V_���,V_ժҪ,������_IN,��������_IN,������_IN,��������_IN,V_��ҩ��,V_��������,V_�ϴι�Ӧ��ID,V_��׼�ĺ�);
 
        --ԭ�����ֲ�������ҩƷ,��C����ʱ��Ҫ������
        BEGIN
            SELECT COUNT(*) INTO V_��¼��
            FROM ҩƷ�շ���¼ A, ҩƷ��� B
            WHERE B.ҩƷID=A.ҩƷID
            AND A.ҩƷID+0=ҩƷID_IN
            AND A.NO=NO_IN
            AND A.���� = 6
            AND A.�ⷿID+0=V_�ⷿID
            AND MOD(A.��¼״̬,3)=0
            AND NVL(A.����,0)>0
            AND (NVL(B.ҩ�����,0)=0 OR
                (NVL(B.ҩ������,0)=0 AND A.�ⷿID IN (SELECT ����ID FROM  ��������˵�� WHERE (�������� LIKE '%ҩ��') OR (�������� LIKE '�Ƽ���'))))
            ;
        EXCEPTION
            WHEN OTHERS THEN V_��¼��:=0;
        END;
        IF V_��¼��>0 THEN
            V_����:=0;
        ELSE
            V_����:=NVL (V_����, 0);
        END IF;
 
        --����ҩƷ�������Ӧ����
        UPDATE ҩƷ���
            SET ��������=NVL(��������,0)-NVL(��������_IN,0)*V_���ϵ��,
                ʵ������=NVL(ʵ������,0)-NVL(��������_IN,0)*V_���ϵ��,
                ʵ�ʽ��=NVL(ʵ�ʽ��,0)-NVL(V_���۽��,0)*V_���ϵ��,
                ʵ�ʲ��=NVL(ʵ�ʲ��,0)-NVL(V_���,0)*V_���ϵ��,
                �ϴβɹ���=NVL(V_�ɱ���,�ϴβɹ���),
                �ϴ�����=NVL(V_����,�ϴ�����),
                �ϴβ���=NVL(V_����,�ϴβ���),
                Ч��=NVL(V_Ч��,Ч��)
          WHERE �ⷿID = V_�ⷿID
            AND ҩƷID = ҩƷID_IN
            AND NVL (����, 0) = V_����
            AND ���� = 1;
 
        IF SQL%NOTFOUND THEN
            INSERT INTO ҩƷ���
            (�ⷿID,ҩƷID,����,����,��������,ʵ������,ʵ�ʽ��,
            ʵ�ʲ��,�ϴβɹ���,�ϴ�����,�ϴβ���,Ч��,�ϴι�Ӧ��id,��׼�ĺ�)
            VALUES
            (V_�ⷿID,ҩƷID_IN,V_����,1,-��������_IN*V_���ϵ��,-��������_IN*V_���ϵ��,
            -V_���۽��*V_���ϵ��,-V_���*V_���ϵ��,V_�ɱ���,V_����,V_����,V_Ч��,V_�ϴι�Ӧ��ID,V_��׼�ĺ�);
        END IF;
 
        DELETE
          FROM ҩƷ���
         WHERE �ⷿID = V_�ⷿID
           AND ҩƷID = ҩƷID_IN
           AND NVL(��������,0)=0
           AND NVL(ʵ������,0)=0
           AND NVL(ʵ�ʽ��,0)=0
           AND NVL(ʵ�ʲ��,0)=0;
 
        --����ҩƷ�շ����ܱ����Ӧ����
        UPDATE ҩƷ�շ�����
         SET ���� =    NVL (����,0)  - NVL (��������_IN,0)*V_���ϵ��,
             ��� = NVL (���, 0) - NVL (V_���۽��, 0)*V_���ϵ��,
             ��� = NVL (���, 0) - NVL (V_���, 0)*V_���ϵ��
        WHERE ���� = TRUNC (��������_IN)
         AND �ⷿID = V_�ⷿID
         AND ҩƷID = ҩƷID_IN
         AND ���ID = V_������ID
         AND ���� = 6;
        IF SQL%NOTFOUND THEN
            INSERT INTO ҩƷ�շ�����
            (����, �ⷿID, ҩƷID, ���ID, ����, ����, ���, ���)
            VALUES
            (TRUNC (��������_IN),V_�ⷿID,ҩƷID_IN,V_������ID,
            6,-��������_IN*V_���ϵ��,-V_���۽��*V_���ϵ��,-V_���*V_���ϵ��);
        END IF;
    END LOOP;
EXCEPTION
    WHEN ERR_ISSTRIKED THEN
        RAISE_APPLICATION_ERROR (-20101, '[ZLSOFT]�õ����Ѿ������˳�����[ZLSOFT]');
    WHEN ERR_ISBATCH THEN
        RAISE_APPLICATION_ERROR (-20102, '[ZLSOFT]�õ����а�����һ��ԭ�������������ڷ�����ҩƷ�����ܳ�����[ZLSOFT]');
    WHEN ERR_ISNONUM THEN
        RAISE_APPLICATION_ERROR (-20103, '[ZLSOFT]�õ����е�' || ceil(���_IN/2) || '�е�ҩƷ����������������ʣ������ݣ����ܳ�����[ZLSOFT]' );
    WHEN OTHERS THEN ZL_ERRORCENTER (SQLCODE, SQLERRM);
END ZL_ҩƷ�ƿ�_STRIKE;
/

CREATE OR REPLACE PROCEDURE ZL_ҩƷ�ƿ�_PREPARE(
    NO_IN IN ҩƷ�շ���¼.NO%TYPE,
    ����Ա_IN VARCHAR2:=NULL
)
IS
  str���� VARCHAR2(20);
  V_�¿��ÿ�� ϵͳ������.����ֵ%Type;

    CURSOR C_ҩƷ�շ���¼
    IS
		SELECT ʵ������,�ⷿID,����,ҩƷID,����,Ч��,����
		FROM ҩƷ�շ���¼
		WHERE NO = NO_IN
		AND ���� = 6
		AND ���ϵ�� = -1
		ORDER BY ҩƷID;
BEGIN
    IF ����Ա_IN IS NOT NULL THEN
		UPDATE ҩƷ�շ���¼
		SET ��ҩ��=����Ա_IN,
			���=to_char(SYSDATE,'yyyy-MM-dd hh24:mi:ss')
		WHERE ����=6 AND NO=NO_IN;
	ELSE
		SELECT to_char(SYSDATE,'yyyy-MM-dd hh24:mi:ss') INTO str���� FROM dual;

		UPDATE ҩƷ�շ���¼
		SET ��ҩ����=to_date(str����,'yyyy-MM-dd hh24:mi:ss')
		WHERE ����=6 AND NO=NO_IN;
    
    --���ݲ��������Ƿ��·�ҩ�ⷿ�Ŀ��ÿ��
    Select ����ֵ Into v_�¿��ÿ�� From ϵͳ������ Where ������ = 96;
    
    --����Ϊ0��ʾ�ڷ���ʱ���¿�������
    If v_�¿��ÿ��='0' Then
  		FOR v_ҩƷ�շ���¼ IN c_ҩƷ�շ���¼ LOOP
  			UPDATE ҩƷ���
  			SET ��������=NVL(��������,0)-v_ҩƷ�շ���¼.ʵ������
  			WHERE �ⷿID=v_ҩƷ�շ���¼.�ⷿID AND ҩƷID=v_ҩƷ�շ���¼.ҩƷID AND NVL(����,0)=NVL(v_ҩƷ�շ���¼.����,0) AND ����=1;
  			IF SQL%ROWCOUNT=0 THEN
  				INSERT INTO ҩƷ���(�ⷿID,ҩƷID,����,����,��������,�ϴ�����,Ч��,�ϴβ���)
  				VALUES (v_ҩƷ�շ���¼.�ⷿID,v_ҩƷ�շ���¼.ҩƷID,NVL(v_ҩƷ�շ���¼.����,0),1,-1*v_ҩƷ�շ���¼.ʵ������,
  					v_ҩƷ�շ���¼.����,v_ҩƷ�շ���¼.Ч��,v_ҩƷ�շ���¼.����);
  			END IF ;
  
  			DELETE
  			FROM ҩƷ���
  			WHERE �ⷿID = v_ҩƷ�շ���¼.�ⷿID
  			AND ҩƷID = v_ҩƷ�շ���¼.ҩƷID
  			AND NVL(��������,0)=0
  			AND NVL(ʵ������,0)=0
  			AND NVL(ʵ�ʽ��,0)=0
  			AND NVL(ʵ�ʲ��,0)=0;
  		END LOOP ;
    End If;
	END IF ;
EXCEPTION
    WHEN OTHERS THEN ZL_ERRORCENTER (SQLCODE, SQLERRM);
END ZL_ҩƷ�ƿ�_PREPARE;
/


CREATE OR REPLACE PROCEDURE ZL_ҩƷ�ƿ�_BACK(
    NO_IN IN ҩƷ�շ���¼.NO%TYPE
)
IS
	str���� ҩƷ�շ���¼.��ҩ����%TYPE;
	str��ҩ ҩƷ�շ���¼.��ҩ��%TYPE;
	str��� ҩƷ�շ���¼.�����%TYPE;

	Err_Note VARCHAR2(255);
	Err_Custom EXCEPTION ;
  
  V_�¿��ÿ�� ϵͳ������.����ֵ%Type;
  
    CURSOR C_ҩƷ�շ���¼
    IS
		SELECT ʵ������,�ⷿID,����,ҩƷID,����,Ч��,����,��ҩ��λID,��׼�ĺ�
		FROM ҩƷ�շ���¼
		WHERE NO = NO_IN
		AND ���� = 6
		AND ���ϵ�� = -1
		ORDER BY ҩƷID;
BEGIN
    SELECT ��ҩ��,��ҩ����,����� INTO str��ҩ,str����,str���
	FROM ҩƷ�շ���¼
	WHERE ����=6 AND NO=NO_IN AND ROWNUM<2;

	IF str��� IS NOT NULL THEN
		Err_Note:='�õ����ѱ����տⷿ���գ�����������ˣ�';
		RAISE Err_Custom;
	END IF ;
	IF str��ҩ IS NULL THEN
		RETURN ;
	END IF ;
	IF str���� IS NULL THEN
		--��������ҩ��Ϊ�ռ���
		UPDATE ҩƷ�շ���¼
		SET ��ҩ��=NULL,
			���=NULL
		WHERE ����=6 AND NO=NO_IN;
	ELSE
		--��Ҫ�ָ�����ⷿ�Ŀ�������
		UPDATE ҩƷ�շ���¼
		SET ��ҩ����=NULL
		WHERE ����=6 AND NO=NO_IN;

    --���ݲ��������Ƿ�ָ���ҩ�ⷿ�Ŀ��ÿ��
  	Select ����ֵ Into v_�¿��ÿ�� From ϵͳ������ Where ������ = 96;
    
    --����Ϊ0��ʾ����ʱҪ�ָ���������
  	If v_�¿��ÿ�� = '0' Then  
      FOR v_ҩƷ�շ���¼ IN c_ҩƷ�շ���¼ LOOP
  			UPDATE ҩƷ���
  			SET ��������=NVL(��������,0)+v_ҩƷ�շ���¼.ʵ������
  			WHERE �ⷿID=v_ҩƷ�շ���¼.�ⷿID AND ҩƷID=v_ҩƷ�շ���¼.ҩƷID AND NVL(����,0)=NVL(v_ҩƷ�շ���¼.����,0) AND ����=1;
  			IF SQL%ROWCOUNT=0 THEN
  				INSERT INTO ҩƷ���(�ⷿID,ҩƷID,����,����,��������,�ϴ�����,Ч��,�ϴβ���,�ϴι�Ӧ��ID,��׼�ĺ�)
  				VALUES (v_ҩƷ�շ���¼.�ⷿID,v_ҩƷ�շ���¼.ҩƷID,NVL(v_ҩƷ�շ���¼.����,0),1,v_ҩƷ�շ���¼.ʵ������,
  					v_ҩƷ�շ���¼.����,v_ҩƷ�շ���¼.Ч��,v_ҩƷ�շ���¼.����,v_ҩƷ�շ���¼.��ҩ��λID,v_ҩƷ�շ���¼.��׼�ĺ�);
  			END IF ;
  
  			DELETE
  			FROM ҩƷ���
  			WHERE �ⷿID = v_ҩƷ�շ���¼.�ⷿID
  			AND ҩƷID = v_ҩƷ�շ���¼.ҩƷID
  			AND NVL(��������,0)=0
  			AND NVL(ʵ������,0)=0
  			AND NVL(ʵ�ʽ��,0)=0
  			AND NVL(ʵ�ʲ��,0)=0;
  		END LOOP ;
  	END IF ;
  End If;
EXCEPTION
    WHEN Err_Custom THEN
        RAISE_APPLICATION_ERROR (-20101, '[ZLSOFT]'||Err_Note||'[ZLSOFT]');
    WHEN OTHERS THEN ZL_ERRORCENTER (SQLCODE, SQLERRM);
END ZL_ҩƷ�ƿ�_BACK;
/

CREATE OR REPLACE PROCEDURE zl_ҩƷ����۵���_strike (
    NO_IN IN ҩƷ�շ���¼.NO%TYPE,
    �����_IN IN ҩƷ�շ���¼.�����%TYPE
)
IS
    Err_isstriked EXCEPTION;
    Err_isBatch exception;
    v_BatchCount integer;    --ԭ���������ڷ�����ҩƷ������
    V_COUNT INTEGER;    --ԭ�����ֲ�����
    V_���� ҩƷ�շ���¼.����%TYPE;

    CURSOR C_ҩƷ�շ���¼
    IS
        SELECT ������ID, �ⷿID, ҩƷID, ����, ���,����,Ч��,����
          FROM ҩƷ�շ���¼ A
         WHERE NO = NO_IN
            AND ���� = 5
            AND ��¼״̬ = 2
         ORDER BY ҩƷID;
BEGIN
    UPDATE ҩƷ�շ���¼
        SET ��¼״̬ = 3
     WHERE NO = NO_IN
        AND ���� = 5
        AND ��¼״̬ = 1;

    IF SQL%ROWCOUNT = 0 THEN
        RAISE Err_isstriked;
    END IF;
    
    --��Ҫ���ԭ���������ڷ�����ҩƷ�����ܶ������
    SELECT COUNT(*) INTO v_BatchCount FROM        
            ҩƷ�շ���¼ a,ҩƷ��� b
    WHERE a.ҩƷid=b.ҩƷid
        AND a.no=NO_IN 
        AND a.����=5
        AND a.��¼״̬=3
        AND nvl(a.����,0)=0
        AND ((NVL(B.ҩ�����,0)=1 AND A.�ⷿID NOT IN (SELECT ����ID FROM  ��������˵�� WHERE (�������� LIKE '%ҩ��') OR (�������� LIKE '�Ƽ���')))
            OR NVL(B.ҩ������,0)=1);
    
    IF v_batchcount>0 THEN
        RAISE Err_isBatch;
    END IF;  

    Insert INTO ҩƷ�շ���¼
                    (
                        ID,
                        ��¼״̬,
                        ����,
                        NO,
                        ���,
                        �ⷿID,
                        ������ID,
                        ���ϵ��,
                        ҩƷID,
                        ����,
                        ����,
                        ����,
                        Ч��,
                        ���۽��,
                        ���,
                        ժҪ,
                        ������,
                        ��������,
                        �����,
                        �������
                    )
        SELECT ҩƷ�շ���¼_ID.Nextval, 2, ����, NO_IN, ���, �ⷿID,
                 ������ID,
                 ���ϵ��, a.ҩƷID, 
                 DECODE(NVL(a.����,0),0,NULL,(DECODE(NVL(b.ҩ�����,0),0,NULL,a.����))), 
                 a.����, ����, Ч��, ���۽��, -���, ժҪ,
                 �����_IN, SYSDATE, �����_IN, SYSDATE
          FROM ҩƷ�շ���¼ a,ҩƷ��� b
         WHERE NO = NO_IN
            AND a.ҩƷid=b.ҩƷid
            AND ���� = 5
            AND ��¼״̬ = 3;

    FOR V_ҩƷ�շ���¼ IN C_ҩƷ�շ���¼ LOOP
        --ԭ�����ֲ�������ҩƷ,��C����ʱ��Ҫ������
        BEGIN 
            SELECT COUNT(*) INTO V_COUNT
            FROM ҩƷ�շ���¼ A, ҩƷ��� B
            WHERE B.ҩƷID=V_ҩƷ�շ���¼.ҩƷID
            AND A.NO=NO_IN
            AND A.���� = 5 
            and a.�ⷿid+0=V_ҩƷ�շ���¼.�ⷿid
            AND A.��¼״̬ = 3 
            AND NVL(A.����,0)>0
            AND (NVL(B.ҩ�����,0)=0 OR 
                (NVL(B.ҩ������,0)=0 AND A.�ⷿID IN (SELECT ����ID FROM  ��������˵�� WHERE (�������� LIKE '%ҩ��') OR (�������� LIKE '�Ƽ���'))))
            ;
        EXCEPTION 
            WHEN OTHERS THEN
                V_COUNT:=0;
        END;
        IF V_COUNT>0 THEN
            V_����:=0;
        ELSE
            V_����:=NVL (V_ҩƷ�շ���¼.����, 0);
        END IF;

        --����ҩƷ�������Ӧ����

        UPDATE ҩƷ���
            SET ʵ�ʲ�� = NVL (ʵ�ʲ��, 0) + NVL (V_ҩƷ�շ���¼.���, 0)
         WHERE �ⷿID = V_ҩƷ�շ���¼.�ⷿID
            AND ҩƷID = V_ҩƷ�շ���¼.ҩƷID
            AND NVL (����, 0) = V_����
            AND ���� = 1;

        IF SQL%NOTFOUND THEN
            Insert INTO ҩƷ���
                (�ⷿID, ҩƷID, ����, ����, ʵ�ʲ��,�ϴ�����,Ч��,�ϴβ���)
                  VALUES (
                      V_ҩƷ�շ���¼.�ⷿID,
                      V_ҩƷ�շ���¼.ҩƷID,
                      V_����,
                      1,
                      V_ҩƷ�շ���¼.���,
                      V_ҩƷ�շ���¼.����,
                      V_ҩƷ�շ���¼.Ч��,
					  v_ҩƷ�շ���¼.����
                  );
        END IF;

        DELETE
          FROM ҩƷ���
         WHERE �ⷿID = V_ҩƷ�շ���¼.�ⷿID
            AND ҩƷID = V_ҩƷ�շ���¼.ҩƷID
            AND nvl(��������,0) = 0 AND nvl(ʵ������,0) = 0 AND nvl(ʵ�ʽ��,0) = 0 AND nvl(ʵ�ʲ��,0) = 0;

        --����ҩƷ�շ����ܱ����Ӧ����

        UPDATE ҩƷ�շ�����
            SET ��� = NVL (���, 0) + NVL (V_ҩƷ�շ���¼.���, 0)
         WHERE ���� = TRUNC (SYSDATE)
            AND �ⷿID = V_ҩƷ�շ���¼.�ⷿID
            AND ҩƷID = V_ҩƷ�շ���¼.ҩƷID
            AND ���ID = V_ҩƷ�շ���¼.������ID
            AND ���� = 5;

        IF SQL%NOTFOUND THEN
            Insert INTO ҩƷ�շ�����
                            (����, �ⷿID, ҩƷID, ���ID, ����, ���)
                  VALUES (
                      TRUNC (SYSDATE),
                      V_ҩƷ�շ���¼.�ⷿID,
                      V_ҩƷ�շ���¼.ҩƷID,
                      V_ҩƷ�շ���¼.������ID,
                      5,
                      V_ҩƷ�շ���¼.���
                  );
        END IF;
    END LOOP;
EXCEPTION
    WHEN Err_isstriked THEN
        Raise_application_error (
            -20101, '[ZLSOFT]�õ����Ѿ������˳�����[ZLSOFT]'
        );
    WHEN Err_isBatch THEN
        Raise_application_error ( 
            -20102, '[ZLSOFT]�õ����а�����һ��ԭ�������������ڷ�����ҩƷ�����ܳ�����[ZLSOFT]' 
        );  
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_ҩƷ����۵���_strike;
/

CREATE OR REPLACE PROCEDURE zl_ҩƷ�⹺��Ʊ��Ϣ_UPDATE (
    NO_IN IN ҩƷ�շ���¼.NO%TYPE := NULL,
    ���_IN IN ҩƷ�շ���¼.���%TYPE,
    ��Ʊ��_IN IN Ӧ����¼.��Ʊ��%TYPE := NULL,
    ��Ʊ����_IN IN Ӧ����¼.��Ʊ����%TYPE := NULL,
    ��Ʊ���_IN IN Ӧ����¼.��Ʊ���%TYPE := NULL,
    ��ҩ��λ_IN in Ӧ����¼.��λID%TYPE:=0,
    ������־_IN Number                        --1��δ���������޸ķ�Ʊ��Ϣ; 2�����ֳ��������޸ķ�Ʊ��Ϣ
)
IS
	ErrInfor varchar2(255);
	ErrItem exception;

	V_NO Ӧ����¼.NO%TYPE;
	V_Ӧ��ID Ӧ����¼.ID%TYPE;
    V_�շ�ID Ӧ����¼.�շ�ID%TYPE;
	V_������� Ӧ����¼.�������%TYPE;
    V_��Ʊ��� Ӧ����¼.��Ʊ���%TYPE;--�ɷ�Ʊ���
    V_��ҩ��λID Ӧ����¼.��λID%TYPE;
BEGIN
	If ������־_IN=1 Then    --δ��������
    --ȡ�Ƿ񸶿�ܶ�
    Begin
  		Select max(�������),sum(nvl(��Ʊ���,0)) INTO v_�������,v_��Ʊ���
  		FROM Ӧ����¼
  		WHERE �շ�id=(Select ID From ҩƷ�շ���¼ Where NO=NO_IN And ���=���_IN And ����=1) AND ϵͳ��ʶ=1 And ��¼����=-1;
  	EXCEPTION
  		WHEN OTHERS THEN
  		v_��Ʊ���:=0;
  		NULL;
  	END ;
  	v_�������:=nvl(v_�������,0);
  	IF v_�������<>0 then
  	   ErrInfor:='[ZLSOFT]�õ����Ѿ������˿�������޸ķ�Ʊ��Ϣ[ZLSOFT]';
  	   RAISE ErrItem;
  	END IF ;
  	if ��Ʊ���_IN>v_��Ʊ��� And v_��Ʊ���<>0 then
  		ErrInfor:='[ZLSOFT]��Ʊ���ܴ��ڼƻ�������[ZLSOFT]';
  		raise ErrItem;
  	end if ;

  	SELECT A.ID, nvl(B.��Ʊ���,0), A.��ҩ��λID
        INTO V_�շ�ID, V_��Ʊ���, V_��ҩ��λID
        FROM ҩƷ�շ���¼ A, (Select * From Ӧ����¼ Where ϵͳ��ʶ=1 And ��¼����=0 AND ��¼״̬=1 And ������� Is NULL) B
       WHERE A.ID = B.�շ�ID(+)
          AND A.NO = NO_IN
          AND A.���� = 1
          AND A.��¼״̬ = 1
          AND A.��� = ���_IN;

  	UPDATE Ӧ����¼
  	SET ��Ʊ�� = ��Ʊ��_IN,
  		��Ʊ���� = ��Ʊ����_IN,
  		��Ʊ��� = ��Ʊ���_IN,
  		��λID=��ҩ��λ_IN
  	WHERE �շ�ID = V_�շ�ID And ϵͳ��ʶ=1 And ��¼״̬=1 And ��¼����=0;

  	if sql%rowcount=0 then
  		IF ��Ʊ��_IN IS NOT NULL THEN
  			--����ǵ�һ����ϸ,�����Ӧ����¼��NO
  			BEGIN
  				SELECT NO INTO V_NO FROM Ӧ����¼
  				WHERE ϵͳ��ʶ=1 AND ��¼����=0 AND ��¼״̬=1
  					AND ��ⵥ�ݺ�=NO_IN AND ROWNUM<2;
  			EXCEPTION
  				WHEN OTHERS THEN V_NO:=NEXTNO(67);
  			END ;

  			SELECT Ӧ����¼_ID.NEXTVAL INTO V_Ӧ��ID FROM DUAL;
  			INSERT INTO Ӧ����¼
  			(ID,��¼����,��¼״̬,��λID,NO,ϵͳ��ʶ,�շ�ID,��ⵥ�ݺ�,���ݽ��,��Ʊ��,��Ʊ����,��Ʊ���,Ʒ��,
  			���,����,����,������λ,����,�ɹ���,�ɹ����,������,��������,�����,�������,ժҪ,��ĿID,���)
  			select V_Ӧ��ID,0,1,��ҩ��λ_IN,V_NO,1,V_�շ�ID,A.NO,A.���۽��,��Ʊ��_IN,��Ʊ����_IN,��Ʊ���_IN,B.����,
  			B.���,B.����,A.����,B.���㵥λ,A.ʵ������,A.�ɱ���,A.�ɱ����,A.������,A.��������,A.�����,A.�������,A.ժҪ,A.ҩƷID,A.���
  			from ҩƷ�շ���¼ A,�շ���ĿĿ¼ B
  			Where A.����=1 And A.NO=NO_in And A.���=���_IN And A.ҩƷID=B.ID;
  		END IF;
  	END IF;

      UPDATE Ӧ�����
          SET ��� = NVL (���, 0) - V_��Ʊ���
      WHERE ��λID = V_��ҩ��λID AND ���� = 1;
      IF SQL%NOTFOUND THEN
          INSERT INTO Ӧ�����(��λID, ����, ���)
          VALUES (V_��ҩ��λID, 1,-V_��Ʊ���);
      END IF;
  	UPDATE Ӧ�����
          SET ���=NVL(���,0)+��Ʊ���_IN
      WHERE ��λID=��ҩ��λ_IN AND ����=1;
      IF SQL%NOTFOUND THEN
          INSERT INTO Ӧ�����(��λID, ����, ���)
          VALUES (��ҩ��λ_IN, 1,��Ʊ���_IN);
      END IF;

      --����ҩƷ�շ���¼�еĹ�ҩ��λ
      UPDATE ҩƷ�շ���¼ SET ��ҩ��λID=��ҩ��λ_IN WHERE NO=NO_IN AND ����=1 And ���=���_IN;

      --����ҩƷ�������ϴι�Ӧ��
      UPDATE ҩƷ��� SET �ϴι�Ӧ��ID=��ҩ��λ_IN WHERE (�ⷿID,ҩƷID) IN (SELECT �ⷿID,ҩƷID FROM ҩƷ�շ���¼ WHERE NO=NO_IN AND ����=1) AND ����=1;
   Else      --���ֳ������ݣ�ֻ���·�Ʊ��
      Update Ӧ����¼ Set ��Ʊ��=��Ʊ��_IN Where ��ⵥ�ݺ�=NO_IN And ���=���_IN;
   End If;
EXCEPTION
	when ErrItem then Raise_application_error (-20101, ErrInfor);
    WHEN NO_data_found THEN
        Raise_application_error (-20101, '[ZLSOFT]�õ����Ѿ������˳������Ѿ������[ZLSOFT]');
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_ҩƷ�⹺��Ʊ��Ϣ_UPDATE;
/

CREATE OR REPLACE PROCEDURE zl_BillCopy (
    ����_IN IN ҩƷ�շ���¼.����%TYPE,
    NO_IN IN ҩƷ�շ���¼.NO%TYPE,
    NewNO_IN IN ҩƷ�շ���¼.NO%TYPE
)
IS
	V_NO Ӧ����¼.NO%TYPE;
	V_Ӧ��ID Ӧ����¼.ID%TYPE;
BEGIN
    --���Ʋ����µ���(���ܴ��ڶ��Ѳ��ֳ����ĵ��ݽ��в������,��Ҫ��������,���ɹ��ۡ��ɹ���������������̸���)
    INSERT INTO ҩƷ�շ���¼
    (ID,��¼״̬,����,NO,���,�ⷿID,��ҩ��λID,������ID,���ϵ��,ҩƷID,
    ����,����,����,Ч��,��д����,ʵ������,�ɱ���,�ɱ����,����,���ۼ�,���۽��,
    ���,ժҪ,������,��������,��ҩ��,��ҩ����,����,���,��Ʒ�ϸ�֤,Ƶ��,��������,��׼�ĺ�)
    SELECT ҩƷ�շ���¼_ID.NEXTVAL ID,1 ��¼״̬,A.����,A.NO,A.���,A.�ⷿID,A.��ҩ��λID,A.������ID,A.���ϵ��,A.ҩƷID,
    A.����,A.����,A.����,A.Ч��,B.ʵ������,B.ʵ������,A.�ɱ���,B.�ɱ����,A.����,A.���ۼ�,B.���۽��,
    B.���,A.ժҪ,A.������,A.��������,��ҩ��,��ҩ����,����,���,��Ʒ�ϸ�֤,to_char(Sysdate,'YYYY-MM-DD HH24:MI:SS'),A.��������,A.��׼�ĺ�
    FROM 
        (SELECT ����,NEWNO_IN NO,���,�ⷿID,��ҩ��λID,������ID,���ϵ��,ҩƷID,
        ����,����,����,Ч��,�ɱ���,�ɱ����,����,���ۼ�,���۽��,
        ���,ժҪ,������,��������,��ҩ��,��ҩ����,����,���,��Ʒ�ϸ�֤,��������,��׼�ĺ�
        FROM ҩƷ�շ���¼
        WHERE ����=����_IN AND NO=NO_IN AND (��¼״̬=1 OR MOD(��¼״̬,3)=0)) A,
        (SELECT ���,SUM(ʵ������) ʵ������,Sum(���۽��) ���۽��,Sum(���) ���,Sum(�ɱ����) �ɱ����
        FROM ҩƷ�շ���¼
        WHERE ����=����_IN AND NO=NO_IN 
        GROUP BY ���) B
    WHERE A.���=B.���;

    IF ����_IN=1 THEN 
		V_NO:=NEXTNO(67);
		
		INSERT INTO Ӧ����¼
			(ID,��¼����,��¼״̬,��λID,NO,ϵͳ��ʶ,�շ�ID,��ⵥ�ݺ�,���ݽ��,��Ʊ��,��Ʊ����,��Ʊ���,Ʒ��,
			���,����,����,������λ,����,�ɹ���,�ɹ����,������,��������,�����,�������,ժҪ,��ĿID,���)
		SELECT Ӧ����¼_ID.NEXTVAL,0,1,C.��λID,V_NO,1,A.�շ�ID,NewNO_IN,B.���۽��,C.��Ʊ��,C.��Ʊ����,C.��Ʊ���,B.����,
			B.���,B.����,B.����,B.�ۼ۵�λ, B.ʵ������,B.�ɱ���,B.�ɱ����,B.������,B.��������,B.�����,B.�������,B.ժҪ,B.ҩƷID,B.���
   FROM (SELECT Id �շ�ID,��� FROM ҩƷ�շ���¼ WHERE ����=����_IN AND NO=NEWNO_IN) A,
			(SELECT A.No,A.ҩƷID,A.���,sum(A.���۽��) ���۽��,A.����,sum(A.ʵ������) ʵ������,A.�ɱ���,sum(A.�ɱ����) �ɱ����,
              A.������,min(A.��������) ��������,A.�����,min(A.�������) �������,A.ժҪ ,C.���,C.����,C.����,C.���㵥λ AS �ۼ۵�λ
			FROM ҩƷ�շ���¼ A,�շ���ĿĿ¼ C 
			WHERE A.����=����_IN AND A.NO=NO_IN AND A.ҩƷID=C.Id
      Group By A.No,A.ҩƷID,A.���,A.����,A.�ɱ���,A.������,A.�����,A.ժҪ ,C.���,C.����,C.����,C.���㵥λ) B,
			(Select ��ⵥ�ݺ�,��λID,��Ʊ��,��Ʊ����,SUM(��Ʊ���) ��Ʊ��� ,���
			From Ӧ����¼ 
			Where ϵͳ��ʶ=1 And ��¼����=0
			GROUP BY ��ⵥ�ݺ�,��λID,��Ʊ��,��Ʊ����,���) C
		WHERE A.���=B.��� AND B.No=C.��ⵥ�ݺ� And b.���=c.���;
    
 END IF ;
EXCEPTION
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
        NULL;
END zl_BillCopy;
/

CREATE OR REPLACE PROCEDURE zl_ҩƷ�շ���¼_������ҩ (
    BillID_IN IN ҩƷ�շ���¼.ID%TYPE,
    People_IN IN ҩƷ�շ���¼.�����%TYPE,
    Date_IN IN ҩƷ�շ���¼.�������%TYPE,
    ����_IN IN ҩƷ���.�ϴ�����%TYPE:=NULL,
    Ч��_IN IN ҩƷ���.Ч��%TYPE:=NULL,
    ����_IN IN ҩƷ���.�ϴβ���%TYPE:=NULL,
    ��ҩ����_IN IN ҩƷ�շ���¼.ʵ������%TYPE:=NULL,
    ��ҩ�ⷿ_IN IN ҩƷ�շ���¼.�ⷿID%TYPE:=Null,
    ��ҩ��_IN In ҩƷ�շ���¼.������%TYPE:=Null
)
IS
    --ֻ������
    int��¼״̬ ҩƷ�շ���¼.��¼״̬%TYPE;
    intִ��״̬ ���˷��ü�¼.ִ��״̬%TYPE;
    bln������ҩ NUMBER;
    lng������ID NUMBER (18);
    strNO ҩƷ�շ���¼.NO%TYPE;
    int���� ҩƷ�շ���¼.����%TYPE;
    lng�ⷿID ҩƷ�շ���¼.�ⷿID%TYPE;
    lngҩƷID ҩƷ�շ���¼.ҩƷID%TYPE;
    Dblʵ������ ҩƷ�շ���¼.ʵ������%TYPE;
    Dblʵ�ʽ�� ҩƷ�շ���¼.���۽��%TYPE;
    Dblʵ�ʳɱ� ҩƷ�շ���¼.�ɱ����%TYPE;
    Dblʵ�ʲ�� ҩƷ�շ���¼.���%TYPE;
    lng����ID ҩƷ�շ���¼.����ID%TYPE;
    BillNO NUMBER (8);        --���۵���
    dblԭ�� ҩƷ�շ���¼.���ۼ�%TYPE;
    dbl�ּ� ҩƷ�շ���¼.���ۼ�%TYPE;

    --20020731 Modified by zyb
    --������ҩʱ�������������ʸı��Ĵ���
    lng������ ҩƷ�շ���¼.����%TYPE;
    lng���� ҩƷ���.ҩ������%TYPE;
    lng���� ҩƷ�շ���¼.����%TYPE;			--ԭ����
	str���� ҩƷ�շ���¼.����%TYPE;			--ԭ����
	dateЧ�� ҩƷ�շ���¼.Ч��%TYPE;		--ԭЧ��

    intDigit number(1);
    Err_custom Exception;
    v_Error Varchar2(255);
BEGIN
    --��ȡ���С��λ��
    SELECT NVL(����ֵ,ȱʡֵ) INTO INTDIGIT FROM ϵͳ������ WHERE ������='���ý���λ��';

	IF ��ҩ����_IN IS NOT NULL THEN
        IF ��ҩ����_IN=0 THEN
            RETURN;
        END IF ;
    END IF ;
    --��ȡ���շ���¼�ĵ��ݡ�ҩƷID���ⷿID
    SELECT ����,NO,�ⷿID,ҩƷID,����ID,������ID,��¼״̬,Nvl(����,0),����,Ч��
    INTO int����,strNO,lng�ⷿID,lngҩƷID,lng����ID,lng������ID,int��¼״̬,lng����,str����,dateЧ��
    FROM ҩƷ�շ���¼
    WHERE ID = BillID_IN;
    --��ȡ�ñʼ�¼ʣ��δ�������������
    --������������δ���������
    SELECT SUM(NVL(ʵ������,0)*NVL(����,1)),SUM(NVL(���۽��,0)),SUM(NVL(�ɱ����,0)),SUM(NVL(���,0))
    INTO Dblʵ������,Dblʵ�ʽ��,Dblʵ�ʳɱ�,Dblʵ�ʲ��
    FROM ҩƷ�շ���¼
    WHERE ����� IS NOT NULL AND NO=strNO AND ����=int����
    AND ���=(SELECT ��� FROM ҩƷ�շ���¼ WHERE ID=BillID_IN);

    --���������ҩ��Ϊ�㣬��ʾ����ҩ
    IF Dblʵ������=0 THEN
        v_Error:='�õ����ѱ���������Ա��ҩ����ˢ�º����ԣ�';
        RAISE Err_custom;
    END IF ;
    IF NVL(��ҩ����_IN,0)>Dblʵ������ THEN
        v_Error:='�õ����ѱ���������Ա������ҩ����ˢ�º����ԣ�';
        RAISE Err_custom;
    END IF ;

    --��ȡ��ҩƷ��ǰ�Ƿ��������Ϣ
    SELECT Nvl(ҩ������,0)  INTO lng����
    FROM ҩƷ���
    WHERE ҩƷID=lngҩƷID;
    --����ǲ�����ҩ�������¼������۽����
    bln������ҩ:=0;
    IF NOT (��ҩ����_IN IS NULL OR NVL(��ҩ����_IN,0)=Dblʵ������) THEN
        bln������ҩ:=1;
    END IF ;
    IF bln������ҩ=1 THEN
        Dblʵ�ʽ��:=ROUND(Dblʵ�ʽ��*��ҩ����_IN/Dblʵ������,INTDIGIT);
        Dblʵ�ʳɱ�:=ROUND(Dblʵ�ʳɱ�*��ҩ����_IN/Dblʵ������,INTDIGIT);
        Dblʵ�ʲ��:=ROUND(Dblʵ�ʲ��*��ҩ����_IN/Dblʵ������,INTDIGIT);
        Dblʵ������:=��ҩ����_IN;
    END IF ;

    --lng����:0-������;1-����;2-ԭ�������ֲ�������������������;3-ԭ���������ַ���������������
    IF lng����=0 AND lng����<>0 THEN
        --ԭ�������ֲ�������������������
        lng����:=2 ;
    ELSIF lng����<>0 AND lng����=0 THEN
        --ԭ������,�ַ���,�����µ����Σ������²����ķ�ҩ��¼��ʹ��
        lng����:=3;
    ELSE
        IF lng����=0 THEN
            lng����:=0;
        ELSE
            lng����:=1;
        END IF ;
    END IF ;

    --��¼״̬�ĺ��������仯
    --�����ļ�¼״̬        :iif(int��¼״̬=1,0,1)+1
    --�������ļ�¼״̬        :iif(int��¼״̬=1,0,1)+2
    --�ȴ���ҩ�ļ�¼״̬    :iif(int��¼״̬=1,0,1)+3

    --����������¼
    Insert INTO ҩƷ�շ���¼
    (ID,��¼״̬,����,NO,���,�ⷿID,�Է�����ID,������ID,���ϵ��,ҩƷID,
    ����,����,����,Ч��,����,��д����,ʵ������,�ɱ���,�ɱ����,����,���ۼ�,
    ���۽��,���,ժҪ,������,��������,��ҩ��,�����,�������,����ID,����,Ƶ��,�÷�,��ҩ����,���,������)
    SELECT ҩƷ�շ���¼_ID.Nextval, int��¼״̬+DECODE(int��¼״̬,1,0,1)+1, int����, strNO, ���,
        �ⷿID,�Է�����ID, ������ID, ���ϵ��, ҩƷID, ����, ����, ����, Ч��,
        1, -Dblʵ������, -Dblʵ������, �ɱ���, -Dblʵ�ʳɱ�, ����, ���ۼ�,
        -Dblʵ�ʽ��, -Dblʵ�ʲ��, ժҪ, People_IN, Date_IN, ��ҩ��, People_IN,
        Date_IN,����ID,����,Ƶ��,�÷�,��ҩ����,��ҩ�ⷿ_IN,��ҩ��_IN
    FROM ҩƷ�շ���¼
    WHERE ID = BillID_IN;

    --����ǲ��ֳ�����������Ϊ1��ʵ������Ϊ������ʵ�������Ļ�
    --����������¼�Թ�������ҩ
    SELECT ҩƷ�շ���¼_ID.Nextval INTO lng������ FROM dual;
    Insert INTO ҩƷ�շ���¼
    (ID,��¼״̬,����,NO,���,�ⷿID,�Է�����ID,������ID,���ϵ��,ҩƷID,
    ����,����,����,Ч��,����,��д����,ʵ������,�ɱ���,�ɱ����,����,���ۼ�,
    ���۽��,���,ժҪ,������,��������,��ҩ��,�����,�������,����ID,����,Ƶ��,�÷�,��ҩ����)
    SELECT lng������, int��¼״̬+DECODE(int��¼״̬,1,0,1)+3, int����, strNO, ���,
         �ⷿID, �Է�����ID, ������ID, ���ϵ��, ҩƷID, DECODE(lng����,1,����,3,lng������,NULL), DECODE(lng����,3,����_IN,1,����,����),
         DECODE(lng����,3,����_IN,1,����,NULL),DECODE(lng����,3,Ч��_IN,1,Ч��,NULL),1,
         Dblʵ������,Dblʵ������,�ɱ���,Dblʵ�ʳɱ�,����,���ۼ�,Dblʵ�ʽ��,
         Dblʵ�ʲ��,ժҪ,������,��������,NULL,NULL,NULL,����ID,����,Ƶ��,�÷�,��ҩ����
    FROM ҩƷ�շ���¼
    WHERE ID = BillID_IN;

    --���²��˷��ü�¼��ִ��״̬(0-δִ��;1-��ȫִ��;2-����ִ��)
    SELECT DECODE(SUM(NVL(����,1)*ʵ������),NULL,0,0,0,2) INTO intִ��״̬
    FROM ҩƷ�շ���¼
    WHERE ����=int���� AND No=strNO AND ����id=lng����ID AND ����� IS NOT NULL;
    UPDATE ���˷��ü�¼
    SET ִ��״̬ = intִ��״̬
    WHERE ID =lng����ID;

    --����δ��ҩƷ��¼
    BEGIN
        Insert INTO δ��ҩƷ��¼
        (����,NO,����ID,��ҳID,����,���ȼ�,�Է�����ID,
        �ⷿID,��ҩ����,��������,���շ�,��ҩ��,��ӡ״̬,δ����)
        SELECT A.����, A.NO, A.����ID, A.��ҳID, A.����,NVL (B.���ȼ�, 0) ���ȼ�,
            A.�Է�����ID, A.�ⷿID, A.��ҩ����, A.��������, A.���շ�,NULL, 1, 1
        FROM (
            SELECT B.����, B.NO, A.����ID, A.��ҳID, A.����,
            DECODE (A.��¼״̬,0,0,1) ���շ�,
            B.�Է�����ID, B.�ⷿID, B.��ҩ����, B.��������,C.���
            FROM ���˷��ü�¼ A, ҩƷ�շ���¼ B,������Ϣ C
            WHERE B.ID = BillID_IN
            AND A.ID = B.����ID+0 And A.����ID=C.����ID(+)) A,��� B
            Where B.����(+)=A.���;
    EXCEPTION
        WHEN OTHERS THEN NULL;
    END;

    --�޸�ԭ��¼Ϊ��������¼
    UPDATE ҩƷ�շ���¼
    SET ��¼״̬ = int��¼״̬+DECODE(int��¼״̬,1,0,1)+2
    WHERE ID = BillID_IN;

    --�޸�ҩƷ���(������)
    IF lng����<>3 THEN
        UPDATE ҩƷ���
        SET ʵ������ = NVL (ʵ������, 0) + Dblʵ������,
            ʵ�ʽ�� = NVL (ʵ�ʽ��, 0) + Dblʵ�ʽ��,
            ʵ�ʲ�� = NVL (ʵ�ʲ��, 0) + Dblʵ�ʲ��
        WHERE �ⷿID+0 = lng�ⷿID AND ҩƷID = lngҩƷID AND ���� = 1 AND NVL(����,0)=lng����;

        IF SQL%ROWCOUNT = 0 THEN
            INSERT INTO ҩƷ���
            (�ⷿID,ҩƷID,����,����,ʵ������,ʵ�ʽ��,ʵ�ʲ��,�ϴ�����,Ч��)
            VALUES
            (lng�ⷿID,lngҩƷID,DECODE(lng����,2,NULL,lng����),1,Dblʵ������,Dblʵ�ʽ��,Dblʵ�ʲ��,DECODE(lng����,1,str����,NULL),DECODE(lng����,1,dateЧ��,NULL));
        END IF;
    ELSE
        INSERT INTO ҩƷ���
        (�ⷿID,ҩƷID,����,Ч��,����,ʵ������,ʵ�ʽ��,ʵ�ʲ��,�ϴ�����,�ϴβ���)
        VALUES
        (lng�ⷿID,lngҩƷID,lng������,Ч��_IN,1,dblʵ������,Dblʵ�ʽ��,dblʵ�ʲ��,����_IN,����_IN);
    END IF ;

    DELETE ҩƷ���
    WHERE �ⷿID+0 = lng�ⷿID AND ҩƷID = lngҩƷID AND ����=1
    AND NVL(��������,0) = 0 AND NVL(ʵ������,0) = 0 AND NVL(ʵ�ʽ��,0) = 0 AND NVL(ʵ�ʲ��,0) = 0;

    --����ҩƷ�շ�����
    UPDATE ҩƷ�շ�����
    SET ���� = NVL (����, 0) + Dblʵ������ ,
        ��� = NVL (���, 0) + Dblʵ�ʽ�� ,
        ��� = NVL (���, 0) + Dblʵ�ʲ��
    WHERE �ⷿID+0 = lng�ⷿID AND ҩƷID+0 = lngҩƷID AND ���ID+0 = lng������ID
    AND ���� = TRUNC (Date_IN) AND ���� = int����;

    IF SQL%ROWCOUNT = 0 THEN
        Insert INTO ҩƷ�շ�����
        (����, �ⷿID, ҩƷID, ����, ���ID, ����, ���, ���)
        VALUES
        (TRUNC (Date_IN),lng�ⷿID,lngҩƷID,int����,lng������ID,Dblʵ������ ,Dblʵ�ʽ�� ,Dblʵ�ʲ�� );
    END IF;

    DELETE ҩƷ�շ�����
    WHERE �ⷿID+0 = lng�ⷿID AND ҩƷID+0 = lngҩƷID AND ���ID+0 = lng������ID
    AND ���� = TRUNC (Date_IN) AND ���� = int����
    AND Nvl(����,0)=0 AND Nvl(���,0)=0 And Nvl(���,0)=0;
    
    --���������ҩ
    select nvl(���ۼ�,0) ԭ��,�ּ� Into dblԭ��,dbl�ּ�
    from ҩƷ�շ���¼ a,�շѼ�Ŀ b
    where a.ҩƷid=b.�շ�ϸĿid  And (SYSDATE BETWEEN b.ִ������ AND b.��ֹ���� Or  SYSDATE >= b.ִ������ AND b.��ֹ���� IS Null) 
    And a.id=BillID_IN;
        
    If dblԭ��<>dbl�ּ� Then
       SELECT ҩƷ�շ���¼_ID.Nextval INTO BillNO FROM Dual; 
       
       SELECT ���ID INTO lng������ID FROM ҩƷ�������� WHERE ���� = 13;
       
       Insert INTO ҩƷ�շ���¼ ( ID,��¼״̬,����,NO,���,������ID,ҩƷID,����,����,Ч��,����,����,��д����,ʵ������,�ɱ���,�ɱ����, 
                                   ���ۼ�,����,���۽��,���,ժҪ,������,��������,�ⷿID,���ϵ��,�۸�ID,�����,�������,����ID) 
       Select  ҩƷ�շ���¼_ID.Nextval,��¼״̬,13,BillNO,���,lng������ID,ҩƷID,����,����,Ч��,����,����,Dblʵ������,0,�ɱ���,0, 
                                   dbl�ּ�-dblԭ��,����,(dbl�ּ�-dblԭ��)*Dblʵ������,(dbl�ּ�-dblԭ��)*Dblʵ������,'������ҩ',People_IN,Sysdate,�ⷿID,1,�۸�ID,People_IN,Sysdate,BillID_IN
       From ҩƷ�շ���¼ Where id=BillID_IN;
    End If;
       
EXCEPTION
    When Err_custom Then Raise_Application_Error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When Others Then zl_ErrorCenter(SQLCODE,SQLERRM);
END zl_ҩƷ�շ���¼_������ҩ;
/

CREATE OR REPLACE PROCEDURE zl_ҩƷ�����޶�_Update (
    �ⷿID_IN In ҩƷ�����޶�.�ⷿID%Type,
    ҩƷID_IN IN ҩƷ�����޶�.ҩƷID%Type,
    ����_IN In ҩƷ�����޶�.����%TYPE := 0,
    ����_IN In ҩƷ�����޶�.����%TYPE := 0,
    �̵�_IN In ҩƷ�����޶�.�̵�����%TYPE := '0000',
    ��λ_IN In ҩƷ�����޶�.�ⷿ��λ%TYPE := NULL 
) IS
    v_��� �շ���ĿĿ¼.���%TYPE;
BEGIN
    IF ����_IN<>0 or ����_IN<>0 or �̵�_IN<>'0000' and �̵�_IN is not null or ��λ_IN is not null THEN
       update ҩƷ�����޶�
       set ����=����_IN,����=����_IN,�̵�����=nvl(�̵�_IN,'0000'),�ⷿ��λ=��λ_IN
       where �ⷿID=�ⷿID_IN and ҩƷID=ҩƷID_IN;
       if sql%rowcount=0 then
          INSERT INTO ҩƷ�����޶� (�ⷿID,ҩƷID,����,����,�̵�����,�ⷿ��λ)
          values(�ⷿID_IN,ҩƷID_IN,����_IN,����_IN,nvl(�̵�_IN,'0000'),��λ_IN);
       end if;
    Else
       Delete ҩƷ�����޶�
       where �ⷿID=�ⷿID_IN and ҩƷID=ҩƷID_IN;
    END IF;
    --ɾ���Ѿ��޸����ʵķǿⷿ�޶�
    select ��� into v_��� from �շ���ĿĿ¼ where ID=ҩƷID_IN;
    if v_���='5' then
       delete ҩƷ�����޶�
       where ҩƷID=ҩƷID_IN
             and �ⷿid not in (
                 select distinct ����id from ��������˵�� where �������� like '��ҩ%' or ��������='�Ƽ���');
    end if;
    if v_���='6' then
       delete ҩƷ�����޶�
       where ҩƷID=ҩƷID_IN
             and �ⷿid not in (
                 select distinct ����id from ��������˵�� where �������� like '��ҩ%' or ��������='�Ƽ���');
    end if;
    if v_���='7' then
       delete ҩƷ�����޶�
       where ҩƷID=ҩƷID_IN
             and �ⷿid not in (
                 select distinct ����id from ��������˵�� where �������� like '��ҩ%' or ��������='�Ƽ���');
    end if;

      --�����¿ⷿ��λ  
	If ��λ_In Is Not Null Then
		Update ҩƷ�ⷿ��λ Set ���� = ��λ_In Where ���� = ��λ_In;
		If Sql%Rowcount = 0 Then
			Insert Into ҩƷ�ⷿ��λ
				(����, ����, ����)
				Select trim(to_char(Nvl(Max(To_Number(����)), 0) + 1,'00000')), ��λ_In, Zlspellcode(��λ_In) From ҩƷ�ⷿ��λ;
		End If;
	End If;

EXCEPTION
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_ҩƷ�����޶�_Update;
/

CREATE OR REPLACE Procedure ZL_����ҽ��ִ��_Cancel(
	ҽ��ID_IN		����ҽ��ִ��.ҽ��ID%Type,
	���ͺ�_IN		����ҽ��ִ��.���ͺ�%Type,
	ȡ��Ƥ��_IN		Number:=Null
--������ҽ��ID_IN=����ִ�е�ҽ��ID���������Ϊ��ʾ�ļ�����Ŀ��ID��
) IS
	Cursor c_Advice is
		Select A.ID,A.���ID,A.����ID,A.��ҳID,A.�������,B.�������� 
			From ����ҽ����¼ A,������ĿĿ¼ B Where A.������ĿID=B.ID And A.ID=ҽ��ID_IN;
	r_Advice c_Advice%RowType;

    v_Temp            Varchar2(255);
    v_��Ա���        ���˷��ü�¼.����Ա���%Type;
    v_��Ա����        ���˷��ü�¼.����Ա����%Type;
	
	v_Date		Date;
	v_Count		Number;
Begin
	--�Լ������,ִ�����ֻ��д���˵�һ��������Ŀ��
	Select Count(*) Into v_Count From ����ҽ��ִ�� Where ҽ��ID=ҽ��ID_IN And ���ͺ�+0=���ͺ�_IN;

	Open c_Advice;
	Fetch c_Advice Into r_Advice;

	If r_Advice.�������='C' And r_Advice.���ID IS Not NULL Then
		--����һ���ɼ������м�����Ŀ
		Update ����ҽ������
			Set ִ��״̬=Decode(v_Count,0,0,3) 
		Where ���ͺ�+0=���ͺ�_IN And ҽ��ID IN(
			Select ID From ����ҽ����¼ Where ���ID=r_Advice.���ID);
	Else
		--������������,���鲿λ,�Լ���������ҽ��;�������ҩ�巨�ǵ�������
		Update ����ҽ������
			Set ִ��״̬=Decode(v_Count,0,0,3)
		Where ���ͺ�+0=���ͺ�_IN And ҽ��ID IN(
			Select ID From ����ҽ����¼ Where ID=ҽ��ID_IN
			Union ALL
			Select ID From ����ҽ����¼ Where ���ID=ҽ��ID_IN And ������� IN('F','D'));
	End If;

	--�����ÿ�����Ҫ����ҽ�����
	Update ���˷��ü�¼ 
		Set ִ��״̬=0,ִ��ʱ��=NULL,ִ����=NULL
	Where �շ���� Not IN('5','6','7') And ҽ�����+0=ҽ��ID_IN
		And (��¼����,NO) IN(
			Select ��¼����,NO From ����ҽ������ Where ҽ��ID=ҽ��ID_IN And ���ͺ�+0=���ͺ�_IN
			Union ALL
			Select ��¼����,NO From ����ҽ������ Where ҽ��ID=ҽ��ID_IN And ���ͺ�+0=���ͺ�_IN);

	--ɾ�������ǼǼ�¼(��ǰ��Ա�Ǽǵ�)
	If r_Advice.�������='E' And r_Advice.��������='1' Then
		v_Temp:=zl_Identity;
		v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
		v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
		v_��Ա���:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
		v_��Ա����:=Substr(v_Temp,Instr(v_Temp,',')+1);
		
		Begin
			Select Max(����ʱ��) Into v_Date From ����ҽ��״̬ 
			Where ҽ��ID=ҽ��ID_IN And ��������=10 
				And (������Ա=v_��Ա���� Or Nvl(ȡ��Ƥ��_IN,0)=1);
		Exception
			When Others Then Null;
		End;
		If v_Date IS Not Null Then
			Delete From ���˹�����¼ 
			Where ����ID=r_Advice.����ID And ��¼��Դ=2
				And Nvl(��ҳID,0)=Nvl(r_Advice.��ҳID,0)
				And ��¼ʱ��=v_Date And (��¼��=v_��Ա���� Or Nvl(ȡ��Ƥ��_IN,0)=1);
			If SQL%RowCount>0 Then
				Update ����ҽ����¼ Set Ƥ�Խ��=Null Where ID=ҽ��ID_IN;
			End IF;
		End If;
	End IF;

	Close c_Advice;			
Exception
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_����ҽ��ִ��_Cancel;
/

CREATE OR REPLACE Procedure ZL_����ҽ����¼_����(
--���ܣ�����סԺҽ����״̬�������Ͳ���
    ҽ��ID_IN       ����ҽ����¼.ID%Type,
    FLAG_IN			Number:=0,
	ҽ������_IN		����ҽ����¼.ҽ������%Type:=Null,
	��������_IN		����ҽ��״̬.��������%Type:=Null
--������ҽ��ID_IN=���IDΪ�յ�ҽ����ID(��ҩ;��,��ҩ�÷�,��Ҫ����,�����Ŀ,������ҽ��),�൱��ҽ����ID
--      FLAG_IN=�������ݡ�����ֹͣ��0=���ִ����ֹʱ��,1=�������е�ִ����ֹʱ�䡣
--      ҽ������_IN=�ù��̱��������˵���ʱ���ã����ڴ�����ʾ��
--      ��������_IN=�ù��̱��������˵���ʱ���ã����ں˶Ի������ݡ�0-���˷���,n=���˾���ҽ������
) IS
    --����ָ��ҽ���Ĳ�����¼,��һ��ΪҪ���˵�����(״̬��������)
	--���������˷��ͺ���Զ�ֹͣ,�ڻ��˷���ʱ�Զ�����ֹͣ����
    Cursor c_RollAdvice is
        Select Distinct B.������Ա,B.����ʱ��,0 AS ���ͺ�,B.��������,
            0 AS ִ��״̬,Sysdate+NULL AS �״�ʱ��,Sysdate+NULL AS ĩ��ʱ��,
			A.�ϴ�ִ��ʱ��,A.ҽ����Ч,A.������� AS ���,Null AS ����,
			A.����ID,A.��ҳID,A.Ӥ��,A.Ƥ�Խ��
        From ����ҽ����¼ A,����ҽ��״̬ B
        Where A.ID=B.ҽ��ID And (A.ID=ҽ��ID_IN Or A.���ID=ҽ��ID_IN)
            And (Nvl(A.ҽ����Ч,0)=0 And B.�������� Not IN(1,2,3) 
				Or Nvl(A.ҽ����Ч,0)=1 And B.�������� Not IN(1,2,3,8))
        Union ALL
        Select Distinct B.������ AS ������Ա,B.����ʱ�� AS ����ʱ��,
			B.���ͺ�,-NULL as ��������,B.ִ��״̬,B.�״�ʱ��,B.ĩ��ʱ��,
			A.�ϴ�ִ��ʱ��,A.ҽ����Ч,C.���,C.�������� AS ����,
			A.����ID,A.��ҳID,A.Ӥ��,A.Ƥ�Խ��
        From ����ҽ����¼ A,����ҽ������ B,������ĿĿ¼ C
        Where A.ID=B.ҽ��ID And A.������ĿID=C.ID
			And (A.ID=ҽ��ID_IN Or A.���ID=ҽ��ID_IN)
        Order by ����ʱ�� Desc,���ͺ�;
    r_RollAdvice c_RollAdvice%RowType;
    
    --����ҽ��������NO������λ���Ҫ���ʵķ��ü�¼
    --һ��ҽ�������Ƕ���д�˷��ͼ�¼,�ҿ���NO��ͬ(ҩƷ��,�÷��巨��һ����)
    --���ܷ��ͼ�¼�ļƷ�״̬(��������Ʒ�),�з��ü�¼��Ȼ��������
    --����ֻ��۸񸸺�Ϊ�յ�,�Ա�ȡ�������
    --ֻ�ܼ�¼״̬Ϊ1�ķ���,���������ʻ򲿷����ʵļ�¼,���ٴ���
    Cursor c_RollMoney(v_���ͺ� ����ҽ������.���ͺ�%Type) is
        Select A.NO,A.���,A.ִ��״̬
        From ���˷��ü�¼ A,����ҽ����¼ B,����ҽ������ C
        Where C.ҽ��ID=B.ID And C.���ͺ�=v_���ͺ�
            And (B.ID=ҽ��ID_IN Or B.���ID=ҽ��ID_IN)
            And A.ҽ�����=B.ID And A.��¼״̬ IN(0,1)
            And A.NO=C.NO And A.��¼����=C.��¼����
            And A.�۸񸸺� IS NULL
        Order BY A.�շ�ϸĿID;
	
	--����ɾ�������¼
	Cursor c_Case(v_���ͺ� ����ҽ������.���ͺ�%Type) is
		Select ����ID From ����ҽ������ 
		Where ����ID IS Not NULL And ���ͺ�=v_���ͺ� 
			And ҽ��ID IN(Select ID From ����ҽ����¼ Where ID=ҽ��ID_IN OR ���ID=ҽ��ID_IN);
	r_Case c_Case%RowType;

	--���ڴ�������ҽ���Ļ���
	Cursor c_PatiLog(
		v_����ID ���˱䶯��¼.����ID%Type,
		v_��ҳID ���˱䶯��¼.��ҳID%Type) is
		Select * From ���˱䶯��¼ 
		Where ����ID=v_����ID And ��ҳID=v_��ҳID And ��ֹʱ�� IS NULL
		Order by ��ʼʱ�� Desc;
	r_PatiLog c_PatiLog%RowType;

    v_ҽ��״̬      ����ҽ����¼.ҽ��״̬%Type;
    v_����NO        ���˷��ü�¼.NO%Type;
    v_�������      Varchar2(255);
    v_ĩ��ʱ��      ����ҽ������.ĩ��ʱ��%Type;

    v_Count         Number(5);
    v_Temp          Varchar2(255);
    v_��Ա���      ���˷��ü�¼.����Ա���%Type;
    v_��Ա����      ���˷��ü�¼.����Ա����%Type;

    v_Error         Varchar2(255);
    Err_Custom      Exception;
Begin
	Open c_RollAdvice;
    Fetch c_RollAdvice Into r_RollAdvice;
    If c_RollAdvice%RowCount=0 Then
        v_Error:=Nvl(ҽ������_IN,'��ҽ��')||'��ǰû�п��Ի��˵����ݡ�';
        Raise Err_Custom;
    End IF;
	--�������˵���ʱ�ж�
	If ҽ������_IN Is Not Null Then
		If Nvl(r_RollAdvice.��������,0)<>Nvl(��������_IN,0) Then
			v_Error:=Nvl(ҽ������_IN,'��ҽ��')||'�����뵱ǰҽ��һ����ˣ����ܸ�ҽ���Ѿ�ִ��������������';
			Raise Err_Custom;
		End IF;
	End IF;

    If r_RollAdvice.���ͺ�=0 Then
        --����ҽ��״̬����(��ʱ��ؼ���)
        --4-���ϣ�5-������6-��ͣ��7-���ã�8-ֹͣ��9-ȷ��ֹͣ��10-Ƥ�Խ��
        ------------------------------------------------------------------

        --���ֻ���˻ص�У��״̬
        If r_RollAdvice.��������=3 Then
            v_Error:=Nvl(ҽ������_IN,'��ҽ��')||'��ǰ����ͨ��У��״̬�������ٻ��ˡ�';
            Raise Err_Custom;
        End IF;
        
        --��������ʱ�ָ����뵥������
        If r_RollAdvice.��������=4 Then
			Update ���˲�����¼
				Set ������=NULL,��������=NULL
			Where ������=r_RollAdvice.������Ա And ��������=r_RollAdvice.����ʱ��
				And ID IN(Select ����ID From ����ҽ����¼ Where ID=ҽ��ID_IN Or ���ID=ҽ��ID_IN);
        End IF;

        --ɾ��(����ҽ��)�����״̬������¼
        Delete From ����ҽ��״̬ 
        Where ҽ��ID IN (Select ID From ����ҽ����¼ Where ID=ҽ��ID_IN Or ���ID=ҽ��ID_IN) 
            And ����ʱ��=r_RollAdvice.����ʱ��;

        --ȡɾ����Ӧ�ָ���ҽ��״̬
        Select ��������
            Into v_ҽ��״̬ 
        From ����ҽ��״̬ 
        Where ����ʱ��=(Select Max(����ʱ��) From ����ҽ��״̬ Where ҽ��ID=ҽ��ID_IN)
            And ҽ��ID=ҽ��ID_IN;
        
        --�ָ�(����ҽ��)���˺��״̬
        Update ����ҽ����¼ 
            Set ҽ��״̬=v_ҽ��״̬ 
        Where ID=ҽ��ID_IN Or ���ID=ҽ��ID_IN;
        
        --��������Ĵ���
        If r_RollAdvice.��������=8 Then
            --�����ڷ����ջع���ҽ���������ٳ���ֹͣ(����ֹͣ,�ٷ���,�ٻ��˾�������)
            --���ܳ��ڷ����ջ�ʱ��ȫ���ջ�(���ϴ�ִ��ʱ��)
            Select Nvl(Count(*),0) Into v_Count
            From ����ҽ����¼ A,����ҽ������ B
            Where B.ҽ��ID=A.ID And (A.ID=ҽ��ID_IN Or A.���ID=ҽ��ID_IN)
                And B.���ͺ�=(Select Max(���ͺ�) From ����ҽ������ Where ҽ��ID IN(Select ID From ����ҽ����¼ Where ID=ҽ��ID_IN Or ���ID=ҽ��ID_IN))
                And A.ִ����ֹʱ�� IS Not NULL
                And ((A.�ϴ�ִ��ʱ��<B.ĩ��ʱ��) 
                    Or (A.�ϴ�ִ��ʱ�� IS NULL And B.ĩ��ʱ�� IS Not NULL));
            If v_Count>0 Then
                v_Error:=Nvl(ҽ������_IN,'��ҽ��')||'�ѱ����ڷ����ջأ������ٳ���ֹͣ������';
                Raise Err_Custom;
            End if;

            --����ҽ��ֹͣʱ,���ͣ��ҽ����ʱ��
            Update ����ҽ����¼ 
                Set ִ����ֹʱ��=Decode(FLAG_IN,1,ִ����ֹʱ��,NULL),
                    ͣ��ҽ��=NULL,ͣ��ʱ��=NULL
            Where ID=ҽ��ID_IN Or ���ID=ҽ��ID_IN;
		ElsIf r_RollAdvice.��������=9 Then
            --����ҽ��ֹͣʱ,���ͣ��ҽ����ʱ��
            Update ����ҽ����¼ 
				Set ȷ��ͣ��ʱ��=NULL
			Where ID=ҽ��ID_IN Or ���ID=ҽ��ID_IN;
        ElsIf r_RollAdvice.��������=10 Then
            --���˱�עƤ�Խ��,ͬʱɾ�������Ǽ�(+)��(-),���ݼ�¼ʱ��
			Delete From ���˹�����¼ 
			Where ����ID=r_RollAdvice.����ID
				And Nvl(��ҳID,0)=Nvl(r_RollAdvice.��ҳID,0)
				And ��¼ʱ��=r_RollAdvice.����ʱ��;

            Update ����ҽ����¼ 
                Set Ƥ�Խ��=NULL
            Where ID=ҽ��ID_IN Or ���ID=ҽ��ID_IN;
        End IF;
    Else
        --����ҽ������(�Է��ͺŹؼ���)
        ------------------------------------------------------------------
        --�������ջصĳ���ҩƷҽ�����������(���˷��þͶ�����)
		If Nvl(r_RollAdvice.ҽ����Ч,0)=0 Then
			If r_RollAdvice.�ϴ�ִ��ʱ�� IS Not NULL And r_RollAdvice.ĩ��ʱ�� IS Not NULL Then
				If r_RollAdvice.�ϴ�ִ��ʱ��<r_RollAdvice.ĩ��ʱ�� Then
					v_Error:=Nvl(ҽ������_IN,'��ҽ��')||'������ڷ��͵������ѱ��ջأ������ٻ��ˡ�';
					Raise Err_Custom;
				End IF;
			ElsIF r_RollAdvice.�ϴ�ִ��ʱ�� IS NULL And r_RollAdvice.ĩ��ʱ�� IS Not NULL Then
				--�������ܱ�ȫ�������ջ�
				v_Error:=Nvl(ҽ������_IN,'��ҽ��')||'δ�����ͣ����͵������ѱ�ȫ�������ջأ������ٻ��ˡ�';
				Raise Err_Custom;
			End IF;
		End IF;

        If Nvl(r_RollAdvice.ִ��״̬,0) IN(1,3) Then --1-��ȫִ��;3-����ִ��
            v_Error:=Nvl(ҽ������_IN,'��ҽ��')||'������͵������Ѿ�ִ�л�����ִ�У����ܻ��ˡ�';            
            Raise Err_Custom;
        End IF;

        --��ǰ������Ա    
        v_Temp:=zl_Identity;
        v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
        v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
        v_��Ա���:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
        v_��Ա����:=Substr(v_Temp,Instr(v_Temp,',')+1);

        --������ҽ���ķ�������(��һ��ҽ�������в�ͬNO����)
        --���ԭʼ�����ѱ�����(�򲿷�����),���ù��������ж�
        v_����NO:=NULL;v_�������:=NULL;
        For r_RollMoney In c_RollMoney(r_RollAdvice.���ͺ�) Loop
            If Nvl(v_����NO,'��')<>r_RollMoney.NO Then
                If v_������� IS Not NULL And v_����NO IS Not NULL Then
                    v_�������:=Substr(v_�������,2);
                    zl_סԺ���ʼ�¼_Delete(v_����NO,v_�������,v_��Ա���,v_��Ա����);
                End IF;
                v_�������:=NULL;
            End IF;
            v_����NO:=r_RollMoney.NO;
            v_�������:=v_�������||','||r_RollMoney.���;

            If Nvl(r_RollMoney.ִ��״̬,0)<>0 Then
                v_Error:=Nvl(ҽ������_IN,'��ҽ��')||'���͵ķ��õ���"'||r_RollMoney.NO||'"�е������ѱ����ֻ���ȫִ�У����ܻ��ˡ�';
                Raise Err_Custom;
            End IF;
        End Loop;
        If v_������� IS Not NULL And v_����NO IS Not NULL Then
            v_�������:=Substr(v_�������,2);
            zl_סԺ���ʼ�¼_Delete(v_����NO,v_�������,v_��Ա���,v_��Ա����);
        End IF;

		Open c_Case(r_RollAdvice.���ͺ�);--�����ȴ�

        --ɾ�����ͼ�¼(����ҽ����)
        Delete From ����ҽ������ 
        Where ���ͺ�=r_RollAdvice.���ͺ�
            And ҽ��ID IN(Select ID From ����ҽ����¼ Where ID=ҽ��ID_IN OR ���ID=ҽ��ID_IN);
        
        --ɾ����Ӧ�ı��浥
		Fetch c_Case Into r_Case;
		While c_Case%Found Loop
			Delete From ���˲�����¼ Where ID=r_Case.����ID;
			Fetch c_Case Into r_Case;				
		End Loop;
		Close c_Case;

        --���(����ҽ��)�ϴ�ִ��ʱ��(���ϴη��͵�ĩ��ִ��ʱ��)

		--���г���(���������Գ���)����ʱ����д��ĩ��ʱ��
		--��������û�У���ֻ���ܷ�����һ�Ρ�
        v_ĩ��ʱ��:=NULL;
        Begin
			--һ��ҽ���ķ�����ĩʱ����ͬ,һ����ҩ��ȡ��С��
            --ȡ���IDΪNULL��ҽ���ķ��ͼ�¼��ʱ��
			--����ҩ;������ҩ�÷�����δ��д���ͼ�¼
            Select ĩ��ʱ�� Into v_ĩ��ʱ��
            From ����ҽ������
            Where ҽ��ID IN(
					Select ID From ����ҽ����¼ Where ID=ҽ��ID_IN Or ���ID=ҽ��ID_IN)
                And ���ͺ�=(
					Select Max(���ͺ�) From ����ҽ������ Where ҽ��ID IN(
						Select ID From ����ҽ����¼ Where ID=ҽ��ID_IN Or ���ID=ҽ��ID_IN)
						)
				And Rownum=1;
        Exception
            When Others Then NULL;
        End;
		Update ����ҽ����¼ 
			Set �ϴ�ִ��ʱ��=v_ĩ��ʱ��
		Where ID=ҽ��ID_IN Or ���ID=ҽ��ID_IN;

		--������������ʱ��ͬʱ�Զ�����ֹͣ
		If Nvl(r_RollAdvice.ҽ����Ч,0)=1 Then
			--ɾ��(����ҽ��)�����״̬������¼
			Delete From ����ҽ��״̬ 
			Where ҽ��ID IN (Select ID From ����ҽ����¼ Where ID=ҽ��ID_IN Or ���ID=ҽ��ID_IN) 
				And ����ʱ��=r_RollAdvice.����ʱ��;

			--ȡɾ����Ӧ�ָ���ҽ��״̬
			Select �������� Into v_ҽ��״̬ From ����ҽ��״̬ 
			Where ����ʱ��=(Select Max(����ʱ��) From ����ҽ��״̬ Where ҽ��ID=ҽ��ID_IN)
				And ҽ��ID=ҽ��ID_IN;
			
			--�ָ�(����ҽ��)���˺��״̬
			Update ����ҽ����¼ 
				Set ҽ��״̬=v_ҽ��״̬,
					ִ����ֹʱ��=NULL,
					ͣ��ҽ��=NULL,
					ͣ��ʱ��=NULL
			Where ID=ҽ��ID_IN Or ���ID=ҽ��ID_IN;
		End IF;

		--סԺ����ҽ�����ͺ�Ļ���(3-ת��;5-��Ժ;6-תԺ)
		If r_RollAdvice.���='Z' And Instr(',3,5,6,',Nvl(r_RollAdvice.����,'0'))>0 And Nvl(r_RollAdvice.Ӥ��,0)=0 Then
			Open c_PatiLog(r_RollAdvice.����ID,r_RollAdvice.��ҳID);
			Fetch c_PatiLog Into r_PatiLog;
			If c_PatiLog%Found Then
				If r_RollAdvice.����='3' And r_PatiLog.��ʼԭ��=3 And r_PatiLog.��ʼʱ�� Is Null Then
					--ȡ������ת��״̬
					zl_���˱䶯��¼_Undo(r_RollAdvice.����ID,r_RollAdvice.��ҳID,v_��Ա���,v_��Ա����);
				ElsIF r_RollAdvice.����='5' And r_PatiLog.��ʼԭ��=10 Then
					--ȡ������Ԥ��Ժ״̬
					zl_���˱䶯��¼_Undo(r_RollAdvice.����ID,r_RollAdvice.��ҳID,v_��Ա���,v_��Ա����);
				ElsIF r_RollAdvice.����='6' And r_PatiLog.��ʼԭ��=10 Then
					--ȡ������Ԥ��Ժ״̬
					zl_���˱䶯��¼_Undo(r_RollAdvice.����ID,r_RollAdvice.��ҳID,v_��Ա���,v_��Ա����);
				End IF;
			End If;
			Close c_PatiLog;
		End IF;
    End IF;

    Close c_RollAdvice;
Exception
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_����ҽ����¼_����;
/

CREATE OR REPLACE Procedure ZL_����ҽ������_Insert(
--���ܣ���д����ҽ�����ͼ�¼
    ҽ��ID_IN		����ҽ������.ҽ��ID%Type,
    ���ͺ�_IN       ����ҽ������.���ͺ�%Type,
    ��¼����_IN     ����ҽ������.��¼����%Type,
    NO_IN           ����ҽ������.NO%Type,
    ��¼���_IN     ����ҽ������.��¼���%Type,
    ��������_IN     ����ҽ������.��������%Type,
    �״�ʱ��_IN     ����ҽ������.�״�ʱ��%Type,
    ĩ��ʱ��_IN     ����ҽ������.ĩ��ʱ��%Type,
    ����ʱ��_IN     ����ҽ������.����ʱ��%Type,
    ִ��״̬_IN     ����ҽ������.ִ��״̬%Type,
    ִ�в���ID_IN   ����ҽ������.ִ�в���ID%Type,
    �Ʒ�״̬_IN     ����ҽ������.�Ʒ�״̬%Type,
    First_IN        Number:=0
--������First_IN=��ʾ�Ƿ�һ��ҽ���ĵ�һҽ����,�Ա㴦��ҽ���������(���ҩ,�䷽�ĵ�һ��,��Ϊ��ҩ;��,�䷽�巨,�÷�����Ϊ����������)
) IS
    --�������˼�ҽ��(һ��ҽ���е�һ��)�����Ϣ���α�
    Cursor c_Advice is
        Select
            Nvl(A.���ID,A.ID) AS ��ID,A.���,A.����ID,A.�Һŵ�,A.Ӥ��,B.����,C.��������,
            A.�������,A.ҽ��״̬,A.ҽ������,A.����ҽ��,A.��ʼִ��ʱ��,A.ִ��ʱ�䷽��,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ
        From ����ҽ����¼ A,������Ϣ B,������ĿĿ¼ C
        Where A.����ID=B.����ID And A.������ĿID=C.ID And A.ID=ҽ��ID_IN
        Group BY Nvl(A.���ID,A.ID),A.���,A.����ID,A.�Һŵ�,A.Ӥ��,B.����,C.��������,
            A.�������,A.ҽ��״̬,A.ҽ������,A.����ҽ��,A.��ʼִ��ʱ��,A.ִ��ʱ�䷽��,A.Ƶ�ʴ���,A.Ƶ�ʼ��,A.�����λ;
    r_Advice c_Advice%RowType;

	Cursor c_Pati(v_����ID ������Ϣ.����ID%Type) is
		Select * From ������Ϣ Where ����ID=v_����ID;
	r_Pati c_pati%RowType;

    --������ʱ����
    v_Temp			Varchar2(255);
	v_Count			Number;
	v_��������		������ҳ.��������%Type;
    v_��Ա���      ���˷��ü�¼.����Ա���%Type;
    v_��Ա����      ���˷��ü�¼.����Ա����%Type;

	v_Error         Varchar2(255);
    Err_Custom      Exception;
Begin
    --��ǰ������Ա
    v_Temp:=zl_Identity;
    v_Temp:=Substr(v_Temp,Instr(v_Temp,';')+1);
    v_Temp:=Substr(v_Temp,Instr(v_Temp,',')+1);
    v_��Ա���:=Substr(v_Temp,1,Instr(v_Temp,',')-1);
    v_��Ա����:=Substr(v_Temp,Instr(v_Temp,',')+1);

    --��һ��ҽ���ĵ�һ��ʱ����ҽ������
    If Nvl(First_IN,0)=1 Then
        Open c_Advice;
        Fetch c_Advice Into r_Advice;

        --�����������
        ---------------------------------------------------------------------------------------
        IF Nvl(r_Advice.ҽ��״̬,0)<>1 Then
            v_Error:='"'||r_Advice.����||'"��ҽ��"'||r_Advice.ҽ������||'"�Ѿ��������˷��͡�'
                ||CHR(13)||CHR(10)||'�ò��˵�ҽ������ʧ�ܡ������¶�ȡ�����嵥���ԡ�';
            Raise Err_Custom;
        End IF;

        --���ͺ��ҽ������:�������ͺ��Զ�ֹͣ
        ---------------------------------------------------------------------------------------
        Update ����ҽ����¼
            Set ҽ��״̬=8,
                ִ����ֹʱ��=ĩ��ʱ��_IN,--����û��
                ͣ��ʱ��=����ʱ��_IN,--Ҫ��Ϊ����ʱ����ʾ
                ͣ��ҽ��=v_��Ա����--Ҫ��Ϊ��������ʾ,��ͬ��סԺ,����ҽ���޻�ʿ����
        Where ID=r_Advice.��ID Or ���ID=r_Advice.��ID;

        Insert Into ����ҽ��״̬(
            ҽ��ID,��������,������Ա,����ʱ��)
        Select
            ID,8,v_��Ա����,����ʱ��_IN
        From ����ҽ����¼
        Where ID=r_Advice.��ID Or ���ID=r_Advice.��ID;

        --����ҽ���Ĵ���
        ---------------------------------------------------------------------------------------
        If r_Advice.�������='Z' And Nvl(r_Advice.��������,'0')<>'0' And Nvl(r_Advice.Ӥ��,0)=0 Then
            --1-����;2-סԺ;
			If Instr(',1,2,',r_Advice.��������)>0 And ִ�в���ID_IN IS Not NULL Then
				--��������µ�ԤԼ�Ǽǵ�������1.��ǰ��ԤԼ,2.��ǰ����Ժ,3-��Ҫ��ԤԼʱ���ڵ�סԺ��¼
				Select Count(*) Into v_Count From ������ҳ Where ����ID=r_Advice.����ID And Nvl(��ҳID,0)=0;
				If v_Count=0 Then
					Select Count(*) Into v_Count From ������ҳ Where ����ID=r_Advice.����ID And ��Ժ���� IS NULL;
				End IF;
				If v_Count=0 Then
					Select Count(*) Into v_Count From ������ҳ Where ����ID=r_Advice.����ID 
						And (��Ժ����>=r_Advice.��ʼִ��ʱ�� Or ��Ժ����>=r_Advice.��ʼִ��ʱ��);
				End IF;
				If v_Count=0 Then
					If r_Advice.��������='1' Then
						--����ҽ��,��������"��ʼʱ��"���۵��ٴ�ִ�п���
						Begin
							v_��������:=2;
							Select Decode(�������,1,1,2) Into v_�������� From ��������˵�� Where ��������='�ٴ�' And ����ID=ִ�в���ID_IN;
						Exception
							When Others Then Null;
						End;
					ElsIf r_Advice.��������='2' Then
						--סԺҽ��,��������"��ʼʱ��"�Ǽǵ��ٴ�ִ�п���
						v_��������:=0;
					End IF;
					
					Open c_Pati(r_Advice.����ID);
					Fetch c_Pati Into r_Pati;

					zl_��Ժ������ҳ_Insert(1,v_��������,r_Pati.����ID,r_Pati.סԺ��,NULL,r_Pati.����,r_Pati.�Ա�,r_Pati.����,
						r_Pati.�ѱ�,r_Pati.��������,r_Pati.����,r_Pati.����,r_Pati.ѧ��,r_Pati.����״��,r_Pati.ְҵ,r_Pati.���,r_Pati.���֤��,
						r_Pati.�����ص�,r_Pati.��ͥ��ַ,r_Pati.�����ʱ�,r_Pati.��ͥ�绰,r_Pati.��ϵ������,r_Pati.��ϵ�˹�ϵ,r_Pati.��ϵ�˵�ַ,
						r_Pati.��ϵ�˵绰,r_Pati.������λ,r_Pati.��ͬ��λID,r_Pati.��λ�绰,r_Pati.��λ�ʱ�,r_Pati.��λ������,r_Pati.��λ�ʺ�,
						r_Pati.������,r_Pati.������,r_Pati.��������,ִ�в���ID_IN,NULL,NULL,NULL,NULL,NULL,r_Advice.����ҽ��,NULL,r_Advice.��ʼִ��ʱ��,
						NULL,NULL,r_Pati.ҽ�Ƹ��ʽ,NULL,NULL,NULL,NULL,r_Pati.����,v_��Ա���,v_��Ա����,0,NULL);

					Close c_Pati;
				End If;
			End If;
        End IF;

        Close c_Advice;
    End IF;

    --��д���ͼ�¼
    ---------------------------------------------------------------------------------------
    Insert Into ����ҽ������(
        ҽ��ID,���ͺ�,��¼����,NO,��¼���,��������,������,����ʱ��,ִ��״̬,ִ�в���ID,�Ʒ�״̬,�״�ʱ��,ĩ��ʱ��)
    Values(
        ҽ��ID_IN,���ͺ�_IN,��¼����_IN,NO_IN,��¼���_IN,��������_IN,
        v_��Ա����,����ʱ��_IN,ִ��״̬_IN,ִ�в���ID_IN,�Ʒ�״̬_IN,
        �״�ʱ��_IN,ĩ��ʱ��_IN);
Exception
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_����ҽ������_Insert;
/

CREATE OR REPLACE Procedure ZL_����ҽ����¼_�������(
--���ܣ������¿����޸�ҽ��������ҽ���������Ӱ�췢���仯
    ҽ��ID_IN		����ҽ����¼.ID%TYPE:=Null,--��������������ʱ����ʾ����ָ��ҽ�����
    ���_IN			����ҽ����¼.���%TYPE:=Null,
	����ID_IN		����ҽ����¼.����ID%Type:=Null,--��������������ʱ����ʾ������ҽ����Ž�������
	����ID_IN		Varchar2:=Null--��ҳID��Һŵ�,���ַ����ʹ���
) IS
	v_��ҳID		����ҽ����¼.��ҳID%Type;
	v_�Һŵ�		����ҽ����¼.�Һŵ�%Type;
	v_Ӥ��			����ҽ����¼.Ӥ��%Type;

	Cursor c_Advice Is
		Select A.ID,A.Ӥ��
		From ����ҽ����¼ A,����ҽ��״̬ B,����ҽ����¼ C
		Where A.ID=B.ҽ��ID And B.��������=1
			And A.���ID=C.ID(+) And A.����ID=����ID_IN 
			And (A.��ҳID=v_��ҳID Or A.�Һŵ�=v_�Һŵ�)
		Order by Nvl(A.Ӥ��,0),Nvl(C.���,A.���),Nvl(A.���ID,A.ID),A.���,B.����ʱ��;
	
	v_Count			Number;
    v_Error			Varchar2(255);
    Err_Custom		Exception;
Begin
	If ҽ��ID_IN is Not Null Then
		Update ����ҽ����¼ Set ���=���_IN Where ID=ҽ��ID_IN;
	Else
		v_��ҳID:=Null;v_�Һŵ�:=Null;
		Begin
			Select To_Number(����ID_IN) Into v_��ҳID From Dual;
		Exception
			When Others Then Null;
		End;
		If v_��ҳID Is Null Then
			v_�Һŵ�:=����ID_IN;
		End IF;
		
		--�����������
		v_Count:=1;
		For r_Advice In c_Advice Loop
			If Nvl(v_Ӥ��,0)<>Nvl(r_Advice.Ӥ��,0) Then
				v_Count:=1;
			End IF;
			Update ����ҽ����¼ Set ���=v_Count Where ID=r_Advice.ID;
			v_Ӥ��:=r_Advice.Ӥ��;
			v_Count:=v_Count+1;
		End Loop;
	End IF;
Exception
    When Err_Custom Then Raise_application_error(-20101,'[ZLSOFT]'||v_Error||'[ZLSOFT]');
    When OTHERS Then zl_ErrorCenter(SQLCODE,SQLERRM);
End ZL_����ҽ����¼_�������;
/

CREATE OR REPLACE PROCEDURE zl_ҩƷ�������_UPDATE (
    �ڼ�_IN IN �ڼ��.�ڼ�%TYPE
)
IS
        CURSOR C_�ڼ�� 
        IS 
              SELECT ��ʼ����,��ֹ���� 
              FROM �ڼ�� 
              WHERE �ڼ�>=�ڼ�_IN and sysdate>=��ʼ����;

        CURSOR C_ƽ������� (
              V_��ʼ���� DATE,
              V_��ֹ���� DATE)
        IS 
              select D.ҩ��,S.�ⷿID,S.ҩƷID,S.����,decode(sign(S.���),1,S.���/S.���,M.ָ�������/100) as �����
              from (select O.�ⷿID,O.ҩƷID,O.����,
                           nvl(E.��ǰ���,0)-nvl(J.�������,0)-O.������ as ���,
                           nvl(E.��ǰ���,0)-nvl(J.�������,0)-O.������ as ���
                    from (select �ⷿID,ҩƷID,nvl(����,0) as ����,
                                sum(���ϵ��*���۽��) as ������,
                                sum(���ϵ��*���) as ������
                          from ҩƷ�շ���¼ L
                          where ������� between trunc(V_��ʼ����) and trunc(V_��ֹ����)+1-1/24/60/60
                                and (����=6
                                     and exists 
                                         (select 1 
                                          from ��������˵�� C 
                                          where C.����ID=L.�ⷿID
                                                and C.�������� in('��ҩ��','��ҩ��','��ҩ��'))
                                     and NOt exists 
                                         (select 1 
                                          from ��������˵�� C 
                                          where C.����ID=L.�Է�����ID
                                                and C.�������� in('��ҩ��','��ҩ��','��ҩ��','�Ƽ���'))
                                    or ���� between 7 and 11)
                          group by �ⷿID,ҩƷID,nvl(����,0)) O,
                         (select �ⷿID,ҩƷID,nvl(����,0) as ����,
                                 sum(���ϵ��*���۽��) as �������,
                                 sum(���ϵ��*���) as �������
                          from ҩƷ�շ���¼
                          where �������>=trunc(V_��ֹ����)+1
                          group by �ⷿID,ҩƷID,nvl(����,0)) J,
                         (select �ⷿID,ҩƷID,nvl(����,0) as ����,
                                 sum(ʵ�ʽ��) as ��ǰ���,sum(ʵ�ʲ��) as ��ǰ���
                          from ҩƷ���
                          where ����=1
                          group by �ⷿID,ҩƷID,nvl(����,0)) E
                    where O.�ⷿID=J.�ⷿID(+)
                          and O.ҩƷID=J.ҩƷID(+) 
                          and O.����=J.����(+)
                          and O.�ⷿID=E.�ⷿID(+)
                          and O.ҩƷID=E.ҩƷID(+) 
                          and O.����=E.����(+)) S,
                   ҩƷ��� M,
                   (select ����ID,min(decode(��������,'��ҩ��',1,'��ҩ��',1,'��ҩ��',1,2)) as ҩ��
                    from ��������˵�� 
                    where �������� in('��ҩ��','��ҩ��','��ҩ��','��ҩ��','��ҩ��','��ҩ��')
                    group by ����ID) D
              where S.ҩƷID=M.ҩƷID and S.�ⷿID=D.����ID 
              order by D.ҩ��,S.�ⷿID,S.ҩƷID,S.����;

        CURSOR C_ҩƷ�����¼ (
              V_��ʼ���� DATE,
              V_��ֹ���� DATE,
              V_ҩ�� INTEGER,
              V_�ⷿID INTEGER,
              V_ҩƷID INTEGER,
              V_���� INTEGER)
        IS 
            select ID,����,NO,�������,������ID,���ϵ��,
                   �ɱ���,ʵ������*���� as ʵ������,���۽��,���,����,����,Ч��,�Է�����ID
            from ҩƷ�շ���¼ L
            where ������� between trunc(V_��ʼ����) and trunc(V_��ֹ����)+1-1/24/60/60
                  and �ⷿID=V_�ⷿID
                  and ҩƷID=V_ҩƷID
                  and NVL(����,0)=NVL(V_����,0)
                  and (V_ҩ��=1 
                       and ����=6
                       and NOt exists 
                           (select 1 
                            from ��������˵�� C 
                            where C.����ID=L.�Է�����ID
                                  and C.�������� in('��ҩ��','��ҩ��','��ҩ��','�Ƽ���'))
                      or ���� between 7 and 11);

       v_ԭ��� ҩƷ���.ʵ�ʲ��%Type;
       v_�ֲ�� ҩƷ���.ʵ�ʲ��%Type;
       v_�ɱ��� ҩƷ���.�ϴβɹ���%Type;
       v_�Է����ID INTEGER;
       INTDIGIT NUMBER;
Begin
       --��ȡ���С��λ��
       SELECT NVL(����ֵ,ȱʡֵ) INTO INTDIGIT FROM ϵͳ������ WHERE ������='���ý���λ��';

        FOR v_Period IN C_�ڼ�� LOOP
            FOR v_AvgTax IN C_ƽ������� (v_Period.��ʼ����,v_Period.��ֹ����) LOOP
                FOR v_OutRec IN C_ҩƷ�����¼ (v_Period.��ʼ����,v_Period.��ֹ����,v_AvgTax.ҩ��,v_AvgTax.�ⷿID,v_AvgTax.ҩƷID,v_AvgTax.����) LOOP
                    v_ԭ���:=v_OutRec.���;
                    v_�ֲ��:=round(nvl(v_OutRec.���۽��,0)* v_AvgTax.�����,INTDIGIT);
                    IF nvl(v_OutRec.ʵ������,0)=0 THEN
                        v_�ɱ���:=v_OutRec.�ɱ���;
                    ELSE
                        v_�ɱ���:=round((NVL(v_OutRec.���۽��,0)-v_�ֲ��)/v_OutRec.ʵ������,7);
                    END IF;

                    UPDATE ҩƷ�շ���¼
                    SET    ���=v_�ֲ��,
                           �ɱ����=NVL(v_OutRec.���۽��,0)-v_�ֲ��,
                           �ɱ���=v_�ɱ���
                    WHERE ID = v_OutRec.ID;

                    UPDATE ҩƷ���
                    SET ʵ�ʲ�� = nvl(ʵ�ʲ��,0)+(v_�ֲ��-v_ԭ���)*v_OutRec.���ϵ��
                    WHERE �ⷿID=v_AvgTax.�ⷿID
                          and ҩƷID=v_AvgTax.ҩƷID
                          and Nvl(����,0)=NVL(v_AvgTax.����,0)
                          and ����=1;
                    IF SQL%NOTFOUND THEN
                          Insert into ҩƷ���(�ⷿID,ҩƷID,����,����,��������,ʵ������,ʵ�ʽ��,ʵ�ʲ��,�ϴι�Ӧ��ID,�ϴβɹ���,�ϴ�����,�ϴβ���,Ч��)
                          values (v_AvgTax.�ⷿID,v_AvgTax.ҩƷID,v_AvgTax.����,1,0,0,0,(v_�ֲ��-v_ԭ���)*v_OutRec.���ϵ��,null,v_�ɱ���,v_OutRec.����,v_OutRec.����,v_OutRec.Ч��);
                    END IF;

                    UPDATE ҩƷ�շ�����
                    SET ��� = nvl(���,0)+(v_�ֲ��-v_ԭ���)*v_OutRec.���ϵ��
                    WHERE ���� = TRUNC(v_OutRec.�������)
                          and �ⷿID=v_AvgTax.�ⷿID
                          and ҩƷID=v_AvgTax.ҩƷID
                          AND ���ID = v_OutRec.������ID
                          AND ���� = v_OutRec.����;

                    IF SQL%NOTFOUND THEN
                          Insert into ҩƷ�շ�����(����,�ⷿID,ҩƷID,���ID,����,���,���,����)
                          values (TRUNC(v_OutRec.�������),v_AvgTax.�ⷿID,v_AvgTax.ҩƷID,v_OutRec.������ID,0,0,(v_�ֲ��-v_ԭ���)*v_OutRec.���ϵ��,v_OutRec.����);
                    END IF;

                    IF v_OutRec.����=6 THEN
                        UPDATE ҩƷ�շ���¼
                        SET    ���=v_�ֲ��,
                               �ɱ����=NVL(v_OutRec.���۽��,0)-v_�ֲ��,
                               �ɱ���=v_�ɱ���
                        WHERE NO=v_OutRec.NO
                              AND ���� = 6
                              and ҩƷID+0=v_AvgTax.ҩƷID
                              and NVL(����,0)=NVL(v_AvgTax.����,0)
                              and �ⷿID+0=v_OutRec.�Է�����ID
                              and �Է�����ID+0=v_AvgTax.�ⷿID
                              and ���ϵ��=-1*v_OutRec.���ϵ��;
                        IF SQL%NOTFOUND THEN
                            NULL;
                        ELSE
                            SELECT ������ID
                            INTO   v_�Է����ID
                            FROM   ҩƷ�շ���¼
                            WHERE NO=v_OutRec.NO
                                  AND ���� = 6
                                  and ҩƷID+0=v_AvgTax.ҩƷID
                                  and NVL(����,0)=NVL(v_AvgTax.����,0)
                                  and �ⷿID+0=v_OutRec.�Է�����ID
                                  and �Է�����ID+0=v_AvgTax.�ⷿID
                                  and ���ϵ��=-1*v_OutRec.���ϵ��
                                  and ROWNUM<2;

                            UPDATE ҩƷ���
                            SET ʵ�ʲ�� = nvl(ʵ�ʲ��,0)+(v_�ֲ��-v_ԭ���)*v_OutRec.���ϵ��*-1
                            WHERE �ⷿID=v_OutRec.�Է�����ID
                                  and ҩƷID=v_AvgTax.ҩƷID
                                  and NVL(����,0)=NVL(v_AvgTax.����,0)
                                  and ����=1;
                            IF SQL%NOTFOUND THEN
                                  Insert into ҩƷ���(�ⷿID,ҩƷID,����,����,��������,ʵ������,ʵ�ʽ��,ʵ�ʲ��,�ϴι�Ӧ��ID,�ϴβɹ���,�ϴ�����,�ϴβ���,Ч��)
                                  values (v_OutRec.�Է�����ID,v_AvgTax.ҩƷID,v_AvgTax.����,1,0,0,0,(v_�ֲ��-v_ԭ���)*v_OutRec.���ϵ��*-1,null,v_�ɱ���,v_OutRec.����,v_OutRec.����,v_OutRec.Ч��);
                            END IF;

                            UPDATE ҩƷ�շ�����
                            SET ��� = nvl(���,0)+(v_�ֲ��-v_ԭ���)*v_OutRec.���ϵ��*-1
                            WHERE ���� = TRUNC(v_OutRec.�������)
                                  and �ⷿID=v_OutRec.�Է�����ID
                                  and ҩƷID=v_AvgTax.ҩƷID
                                  AND ���ID = v_�Է����ID
                                  AND ���� = v_OutRec.����;

                            IF SQL%NOTFOUND THEN
                                  Insert into ҩƷ�շ�����(����,�ⷿID,ҩƷID,���ID,����,���,���,����)
                                  values (TRUNC(v_OutRec.�������),v_OutRec.�Է�����ID,v_AvgTax.ҩƷID,v_�Է����ID,0,0,(v_�ֲ��-v_ԭ���)*v_OutRec.���ϵ��*-1,v_OutRec.����);
                            END IF;
                        END IF;
                    END IF;
                END LOOP;
                DELETE FROM ҩƷ���
                WHERE �ⷿID=v_AvgTax.�ⷿID
                      and ҩƷID=v_AvgTax.ҩƷID
                      AND nvl(��������,0) = 0
                      AND nvl(ʵ������,0) = 0
                      AND nvl(ʵ�ʽ��,0) = 0
                      AND nvl(ʵ�ʲ��,0) = 0;
                COMMIT;
            END LOOP;
        END LOOP;
EXCEPTION
    WHEN OTHERS THEN
        zl_ErrorCenter (SQLCODE, SQLERRM);
END zl_ҩƷ�������_UPDATE;
/

CREATE OR REPLACE PROCEDURE ZL_ҩƷ�շ���¼_���ŷ�ҩ (
    PARTID_IN IN ҩƷ�շ���¼.�ⷿID%TYPE,
    BILLID_IN IN ҩƷ�շ���¼.ID%TYPE,
    PEOPLE_IN IN ҩƷ�շ���¼.�����%TYPE,
    DATE_IN IN ҩƷ�շ���¼.�������%TYPE,
    ����_IN IN ҩƷ�շ���¼.����%TYPE:=NULL,
    ��ҩ��ʽ_IN IN ҩƷ�շ���¼.��ҩ��ʽ%TYPE:=3,
    ��ҩ��_IN In ҩƷ�շ���¼.������%TYPE:=Null
)
IS
    --ֻ������
    LNG������ID NUMBER (18);
    INT���ϵ�� NUMBER;
    INTִ��״̬ NUMBER;
    INT���� ҩƷ�շ���¼.����%TYPE;
    STRNO ҩƷ�շ���¼.NO%TYPE;
    LNG�ⷿID ҩƷ�շ���¼.�ⷿID%TYPE;
    LNGҩƷID ҩƷ�շ���¼.ҩƷID%TYPE;
    LNG����ID ҩƷ�շ���¼.����ID%TYPE;
    DBL����� ҩƷ���.ʵ�ʽ��%TYPE;
    DBL����� ҩƷ���.ʵ�ʲ��%TYPE;
    DBL����� ҩƷ���.ָ�������%TYPE;
    INTδ���� δ��ҩƷ��¼.δ����%TYPE;
    --��д����
    DBLʵ������ ҩƷ�շ���¼.ʵ������%TYPE;
    DBLʵ�ʽ�� ҩƷ�շ���¼.���۽��%TYPE;
    DBL�ɱ���� ҩƷ�շ���¼.�ɱ����%TYPE;
    DBLʵ�ʲ�� ҩƷ�շ���¼.���%TYPE;
    --2002-07-31����
    --LNGLAST���� ��ҩǰȷ��������(�Ѽ���������)
    INT��¼״̬ ���˷��ü�¼.ִ��״̬%TYPE;
    STRҩ�� VARCHAR2(200);
    DBL�������� ҩƷ�շ���¼.��д����%TYPE;
    LNGLAST���� ҩƷ�շ���¼.����%TYPE;
    LNGCUR���� ҩƷ�շ���¼.����%TYPE;
    STR���� ҩƷ�շ���¼.����%TYPE;
    STRЧ�� ҩƷ�շ���¼.Ч��%TYPE;
    BLN�շ��뷢ҩ���� NUMBER(1);
    V_ERROR VARCHAR2(255);
    intDigit number(1);
    ERR_CUSTOM EXCEPTION ;

	--�Զ���˷���
	intAutoVerify NUMBER(1);
	str����Ա��� ��Ա��.���%TYPE;
	str����Ա���� ��Ա��.����%TYPE;
	int��� ���˷��ü�¼.���%TYPE;
	lng����ID ���˷��ü�¼.����ID%TYPE;
BEGIN
	--ȡ����Ա���������
	SELECT ���,���� INTO str����Ա���,str����Ա����
	FROM ��Ա�� A,�ϻ���Ա�� B
	WHERE A.ID=B.��ԱID AND B.�û���=USER;

	--��ȡ���С��λ��
    SELECT NVL(����ֵ,ȱʡֵ) INTO INTDIGIT FROM ϵͳ������ WHERE ������='���ý���λ��';
	--�жϻ��۵���ҩ���Ƿ��Զ����Ϊ���ʵ�
	SELECT NVL(����ֵ,ȱʡֵ) INTO intAutoVerify FROM ϵͳ������ WHERE ������='ִ�к��Զ���˻��۵�';

    --��ȡ���շ���¼�ĵ��ݡ�ҩƷID���ⷿID,���۽�ʵ��������������ID
    SELECT A.����,A.NO,A.ҩƷID,A.�ⷿID,A.����ID,NVL(A.���۽��,0),NVL(A.ʵ������, 0)*NVL(A.����,1),
        A.������ID,A.���ϵ��,NVL(A.����,0),'['||C.����||']'||C.����
    INTO INT����, STRNO, LNGҩƷID, LNG�ⷿID,LNG����ID,DBLʵ�ʽ��, DBLʵ������,
        LNG������ID,INT���ϵ��,LNGLAST����,STRҩ��
    FROM ҩƷ�շ���¼ A,�շ���ĿĿ¼ C
    WHERE A.ID = BILLID_IN AND A.ҩƷID=C.ID;
    IF NVL(����_IN,0)=0 THEN
        LNGCUR����:=LNGLAST����;
    ELSE
        LNGCUR����:=NVL(����_IN,0);
    END IF ;

    --����Ƿ��Ѿ���д�ⷿ
    BLN�շ��뷢ҩ����:=0;
    IF LNG�ⷿID IS NULL THEN
        BLN�շ��뷢ҩ����:=1;
    END IF ;
    LNG�ⷿID:=PARTID_IN;

    --ȡ����ҩƷ������
    BEGIN
        SELECT �ϴ�����,Ч��,NVL(��������,0)
        INTO STR����,STRЧ��,DBL��������
        FROM ҩƷ���
        WHERE �ⷿID=LNG�ⷿID AND ҩƷID=LNGҩƷID AND ����=1 AND NVL(����,0)=LNGCUR����;
    EXCEPTION
        WHEN OTHERS THEN
            SELECT '','',0 INTO STR����,STRЧ��,DBL�������� FROM DUAL ;
    END ;

    --���������������˳�
    IF LNGCUR����<>NVL(LNGLAST����,0) THEN
        IF DBL��������<DBLʵ������ AND LNGCUR����<>0 THEN
            V_ERROR:=STRҩ��||'�Ŀ����������㣬������ֹ��';
            RAISE ERR_CUSTOM;
        END IF ;
    END IF ;

    --�����ҩƷ�շ���¼�ĳɱ��ۡ��ɱ������۽����(�ȼ������۽��ټ����ۣ�������ɱ����ɱ���)
    --��ȡ��ҩƷ�Ĳ����(��Ϊֻ���ۼۼ������������,�������������,�������¼��㴦ʹ�õ�ǰȷ����������Ϊ��������)
    BEGIN
        SELECT NVL(ʵ�ʽ��, 0) ʵ�ʽ��, NVL(ʵ�ʲ��, 0) ʵ�ʲ�� INTO DBL�����,DBL�����
        FROM ҩƷ���
        WHERE �ⷿID+0 = LNG�ⷿID AND ҩƷID = LNGҩƷID
        AND ����=1 AND NVL(����,0)=LNGCUR���� ;

        IF DBL�����<=0 OR DBL�����<0 THEN
            SELECT NVL (ָ�������, 15) / 100 ָ������� INTO DBL�����
            FROM ҩƷ��� WHERE ҩƷID = LNGҩƷID;
        ELSE
            DBL����� := DBL�����/DBL����� ;
        END IF ;
    EXCEPTION
        WHEN OTHERS THEN
            SELECT NVL (ָ�������, 15) / 100 ָ������� INTO DBL�����
            FROM ҩƷ��� WHERE ҩƷID = LNGҩƷID;
    END ;
    --���
    DBLʵ�ʲ�� := round(DBLʵ�ʽ�� * DBL�����,intDigit);
    --�ɱ����
    DBL�ɱ���� := round(DBLʵ�ʽ�� - DBLʵ�ʲ��,intDigit);

    --����ҩƷ�շ���¼�����۽��ɱ������
    UPDATE ҩƷ�շ���¼
    SET �ⷿID=lng�ⷿID,
        �ɱ��� = round(DBL�ɱ���� / DECODE(DBLʵ������,NULL,1,0,1,DBLʵ������),5),
        �ɱ���� = DBL�ɱ����,
        ��� = DBLʵ�ʲ��,
        ����=LNGCUR����,
        ����=STR����,
        Ч��=STRЧ��,
        �����=PEOPLE_IN,
        �������=DATE_IN,
        ��ҩ��ʽ=��ҩ��ʽ_IN,
        ������=��ҩ��_IN
    WHERE ID = BILLID_IN;
	--�����������
	IF SQL%RowCount=0 Then
		v_Error:='Ҫ��ҩ��ҩƷ��¼"'||STRҩ��||'"�����ڣ�������ֹ��';
		Raise Err_Custom;
	End IF;

    --���������ѷ�ҩ��������ҩ��Ϊ��ҩ��
    UPDATE ҩƷ�շ���¼
    SET ��ҩ�� = PEOPLE_IN
    WHERE NO = STRNO AND ���� = INT���� AND (�ⷿID+0=LNG�ⷿID OR �ⷿID IS NULL) AND ����� IS NOT NULL AND MOD (��¼״̬, 3) = 1;

    --����ԭ���ο��Ŀ�������
    --���·�ҩ���ο��Ŀ��ü�ʵ������
    IF LNGLAST����<>LNGCUR���� THEN
        UPDATE ҩƷ���
        SET ��������=NVL(��������,0)+DBLʵ������
        WHERE �ⷿID+0 = LNG�ⷿID AND ҩƷID = LNGҩƷID AND ���� = 1 AND NVL(����,0)=LNGLAST����;

        UPDATE ҩƷ���
        SET ��������=NVL(��������,0)-DBLʵ������
        WHERE �ⷿID+0 = LNG�ⷿID AND ҩƷID = LNGҩƷID AND ���� = 1 AND NVL(����,0)=LNGCUR����;
    END IF ;

    IF BLN�շ��뷢ҩ����=1 THEN
        UPDATE ҩƷ���
        SET �������� = NVL (��������, 0) - DBLʵ������,
            ʵ������ = NVL (ʵ������, 0) - DBLʵ������,
            ʵ�ʽ�� = NVL (ʵ�ʽ��, 0) - DBLʵ�ʽ��,
            ʵ�ʲ�� = NVL (ʵ�ʲ��, 0) - DBLʵ�ʲ��
        WHERE �ⷿID+0 = LNG�ⷿID AND ҩƷID = LNGҩƷID AND ���� = 1 AND NVL(����,0)=LNGCUR����;
    ELSE
        UPDATE ҩƷ���
        SET ʵ������ = NVL (ʵ������, 0) - DBLʵ������,
            ʵ�ʽ�� = NVL (ʵ�ʽ��, 0) - DBLʵ�ʽ��,
            ʵ�ʲ�� = NVL (ʵ�ʲ��, 0) - DBLʵ�ʲ��
        WHERE �ⷿID+0 = LNG�ⷿID AND ҩƷID = LNGҩƷID AND ���� = 1 AND NVL(����,0)=LNGCUR����;
    END IF ;

    IF SQL%ROWCOUNT = 0 THEN
        IF BLN�շ��뷢ҩ����=1 THEN
            INSERT INTO ҩƷ���
            (�ⷿID,ҩƷID,����,����,��������,ʵ������,ʵ�ʽ��,ʵ�ʲ��)
            VALUES
            (LNG�ⷿID,LNGҩƷID,LNGCUR����,1,0 - DBLʵ������,0 - DBLʵ������,0 - DBLʵ�ʽ��,0 - DBLʵ�ʲ��);
        ELSE
            INSERT INTO ҩƷ���
            (�ⷿID,ҩƷID,����,����,ʵ������,ʵ�ʽ��,ʵ�ʲ��)
            VALUES
            (LNG�ⷿID,LNGҩƷID,LNGCUR����,1,0 - DBLʵ������,0 - DBLʵ�ʽ��,0 - DBLʵ�ʲ��);
        END IF ;
    END IF;

    DELETE ҩƷ���
    WHERE �ⷿID+0 = LNG�ⷿID AND ҩƷID = LNGҩƷID AND ����=1
    AND NVL(��������,0) = 0 AND NVL(ʵ������,0) = 0 AND NVL(ʵ�ʽ��,0) = 0 AND NVL(ʵ�ʲ��,0) = 0;

    --����ҩƷ�շ�����
    UPDATE ҩƷ�շ�����
    SET ���� = NVL (����, 0) + DBLʵ������ * INT���ϵ��,
        ��� = NVL (���, 0) + DBLʵ�ʽ�� * INT���ϵ��,
        ��� = NVL (���, 0) + DBLʵ�ʲ�� * INT���ϵ��
    WHERE �ⷿID+0 = LNG�ⷿID AND ҩƷID+0 = LNGҩƷID AND ���ID+0 = LNG������ID
    AND ���� = TRUNC (DATE_IN) AND ���� = INT����;

    IF SQL%ROWCOUNT = 0 THEN
        INSERT INTO ҩƷ�շ�����
        (����,�ⷿID,ҩƷID,����,���ID,����,���,���)
        VALUES
        (TRUNC (DATE_IN),LNG�ⷿID,LNGҩƷID,INT����,LNG������ID,
        DBLʵ������ * INT���ϵ��,DBLʵ�ʽ�� * INT���ϵ��,DBLʵ�ʲ�� * INT���ϵ��);
    END IF;

    --���²��˷��ü�¼��ִ��״̬(��ִ��)
    SELECT DECODE(SUM(NVL(����,1)*ʵ������),NULL,1,0,1,2) INTO INTִ��״̬
    FROM ҩƷ�շ���¼
    WHERE ����=INT���� AND NO=STRNO AND ����ID=LNG����ID AND ����� IS NULL AND ��¼״̬<>1 AND MOD(��¼״̬,3)<>0;
    UPDATE ���˷��ü�¼
    SET ִ��״̬ = INTִ��״̬
    WHERE ID = LNG����ID;

    --����δ��ҩƷ��¼(���δ����Ϊ����ɾ��)
    SELECT COUNT(*) INTO INTδ����
    FROM ҩƷ�շ���¼
    WHERE ����=INT���� AND (�ⷿID+0=LNG�ⷿID OR �ⷿID IS NULL) AND NO=STRNO AND ����� IS NULL AND NVL(LTRIM(RTRIM(ժҪ)),'С��')<>'�ܷ�';

    IF INTδ���� = 0 THEN
        DELETE δ��ҩƷ��¼ WHERE NO = STRNO AND ���� = INT���� AND (�ⷿID+0=LNG�ⷿID OR �ⷿID IS NULL);
    END IF;

	--������ˣ��ظ����Ҳû�й�ϵ��
	IF intAutoVerify=1 THEN
		SELECT ���,����ID,NO INTO int���,lng����ID,strNO
		FROM ���˷��ü�¼
		WHERE id = (SELECT ����ID FROM ҩƷ�շ���¼ WHERE ID=BILLID_IN);

		zl_סԺ���ʼ�¼_Verify(strNO,str����Ա���,str����Ա����,int���,lng����ID,DATE_IN);
	END IF ;
EXCEPTION
    WHEN ERR_CUSTOM THEN RAISE_APPLICATION_ERROR(-20101,'[ZLSOFT]'||V_ERROR||'[ZLSOFT]');
    WHEN OTHERS THEN ZL_ERRORCENTER (SQLCODE, SQLERRM);
END ZL_ҩƷ�շ���¼_���ŷ�ҩ;
/

-------------------------------------------------------------------------------------------------------
--����ϵͳ�������İ汾��
-------------------------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.12.0' Where ���=100;
--�����汾��
Update zlComponent Set ���汾=10,�ΰ汾=12,���汾=0 Where Upper(����)=Upper('zl9BaseItem')	And ϵͳ=100;
Update zlComponent Set ���汾=10,�ΰ汾=12,���汾=0 Where Upper(����)=Upper('zl9CISBase')		And ϵͳ=100;
Update zlComponent Set ���汾=10,�ΰ汾=12,���汾=0 Where Upper(����)=Upper('zl9Patient')		And ϵͳ=100;
Update zlComponent Set ���汾=10,�ΰ汾=12,���汾=0 Where Upper(����)=Upper('zl9InPatient')   And ϵͳ=100;
Update zlComponent Set ���汾=10,�ΰ汾=12,���汾=0 Where Upper(����)=Upper('zl9RegEvent')	And ϵͳ=100;
Update zlComponent Set ���汾=10,�ΰ汾=12,���汾=0 Where Upper(����)=Upper('zl9OutExse')		And ϵͳ=100;
Update zlComponent Set ���汾=10,�ΰ汾=12,���汾=0 Where Upper(����)=Upper('zl9InExse')		And ϵͳ=100;
Update zlComponent Set ���汾=10,�ΰ汾=12,���汾=0 Where Upper(����)=Upper('zl9CustAcc')		And ϵͳ=100;
Update zlComponent Set ���汾=10,�ΰ汾=12,���汾=0 Where Upper(����)=Upper('zl9CashBill')	And ϵͳ=100;
Update zlComponent Set ���汾=10,�ΰ汾=12,���汾=0 Where Upper(����)=Upper('zl9DrugStore')   And ϵͳ=100;
Update zlComponent Set ���汾=10,�ΰ汾=12,���汾=0 Where Upper(����)=Upper('zl9MediStore')   And ϵͳ=100;
Update zlComponent Set ���汾=10,�ΰ汾=12,���汾=0 Where Upper(����)=Upper('zl9Stuff')		And ϵͳ=100;
Update zlComponent Set ���汾=10,�ΰ汾=12,���汾=0 Where Upper(����)=Upper('zl9Due')			And ϵͳ=100;
Update zlComponent Set ���汾=10,�ΰ汾=12,���汾=0 Where Upper(����)=Upper('zl9Analysis')	And ϵͳ=100;
Update zlComponent Set ���汾=10,�ΰ汾=12,���汾=0 Where Upper(����)=Upper('zl9Ops')			And ϵͳ=100;
Update zlComponent Set ���汾=10,�ΰ汾=12,���汾=0 Where Upper(����)=Upper('zl9Medical')		And ϵͳ=100;
Update zlComponent Set ���汾=10,�ΰ汾=12,���汾=0 Where Upper(����)=Upper('zl9CISWork')		And ϵͳ=100;
Update zlComponent Set ���汾=10,�ΰ汾=12,���汾=0 Where Upper(����)=Upper('zl9PacsWork')	And ϵͳ=100;
Update zlComponent Set ���汾=10,�ΰ汾=12,���汾=0 Where Upper(����)=Upper('zl9LisWork')		And ϵͳ=100;
Update zlComponent Set ���汾=10,�ΰ汾=12,���汾=0 Where Upper(����)=Upper('zl9ImgCapture')	And ϵͳ=100;
--ҽ������
Update zlComponent Set ���汾=9,�ΰ汾=23,���汾=50 Where Upper(����)=Upper('zl9Insure')		And ϵͳ=100;

Commit;