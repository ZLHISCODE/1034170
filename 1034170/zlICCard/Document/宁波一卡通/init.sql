ALTER TABLE һ��ͨĿ¼ MODIFY ҽԺ���� VARCHAR2 (6);
ALTER TABLE һ��ͨĿ¼ ADD CONSTRAINT һ��ͨĿ¼_CK_���� CHECK (���� IN(0,1,2));
CREATE TABLE ���˷�����¼ (
	����ID NUMBER (18),
	�ɿ��� VARCHAR2 (50),
	�ɿ����� NUMBER (3) DEFAULT 2,
	�ɿ�����ҽԺ VARCHAR2 (50),
	�ɿ�����ʱ�� VARCHAR2 (20),
	�ɿ����� VARCHAR2 (20),
	�¿��� VARCHAR2 (50),
	����ʱ�� DATE ,
	�ϴ���־ NUMBER (1) DEFAULT 0);
ALTER TABLE ���˷�����¼ ADD CONSTRAINT ���˷�����¼_PK PRIMARY KEY (����ID,�ɿ���) Using Index Pctfree 0 Tablespace zl9indexhis;

CREATE OR REPLACE PROCEDURE zl_���˷�����¼_����(
	����ID_IN IN ���˷�����¼.����ID%TYPE,
	�¿���_IN IN ���˷�����¼.�¿���%TYPE
)
AS 
BEGIN
	UPDATE ���˷�����¼
	SET �¿���=�¿���_IN,
		����ʱ��= SYSDATE 
	WHERE ����ID=����ID_IN; 
END zl_���˷�����¼_����;
/

CREATE OR REPLACE PROCEDURE zl_���˷�����¼_������(
	����ID_IN IN ���˷�����¼.����ID%TYPE,
	�ɿ���_IN IN ���˷�����¼.�ɿ���%TYPE,
	�¿���_IN IN ���˷�����¼.�¿���%TYPE,
	�ɿ�����ҽԺ_IN IN ���˷�����¼.�ɿ�����ҽԺ%TYPE:=NULL,
	�ɿ�����_IN IN ���˷�����¼.�ɿ�����%TYPE:=2,
	�ɿ�����_IN IN ���˷�����¼.�ɿ�����%TYPE:=NULL,
	�ɿ�����ʱ��_IN IN ���˷�����¼.�ɿ�����ʱ��%TYPE:=NULL
)
AS 
BEGIN
	INSERT INTO ���˷�����¼
	(����ID,�ɿ���,�ɿ�����,�ɿ�����,�ɿ�����ҽԺ,�ɿ�����ʱ��,�¿���,����ʱ��)
	SELECT ����ID_IN,�ɿ���_IN,�ɿ�����_IN,�ɿ�����_IN,NVL(�ɿ�����ҽԺ_IN,ҽԺ����),NVL(�ɿ�����ʱ��_IN,SYSDATE ),�¿���_IN,SYSDATE 
	FROM һ��ͨĿ¼
	WHERE ����='����һ��ͨ';
END zl_���˷�����¼_������;
/

CREATE OR REPLACE PROCEDURE zl_���˷�����¼_�ϴ�(
	����ID_IN IN ���˷�����¼.����ID%TYPE)
IS 
BEGIN
	UPDATE ���˷�����¼ 
	SET �ϴ���־=1
	WHERE ����ID=����ID_IN;
END zl_���˷�����¼_�ϴ�;
/

CREATE OR REPLACE PROCEDURE zl_������Ϣ�ӱ�_Update(
	����ID_IN IN ������Ϣ�ӱ�.����ID%TYPE,
	��Ϣ��_IN IN ������Ϣ�ӱ�.��Ϣ��%TYPE,
	��Ϣֵ_IN IN ������Ϣ�ӱ�.��Ϣֵ%TYPE)
AS 
BEGIN
	UPDATE ������Ϣ�ӱ�
	SET ��Ϣֵ=��Ϣֵ_IN
	WHERE ��Ϣ��=��Ϣ��_IN;
	IF SQL%ROWCOUNT =0 THEN 
		INSERT INTO ������Ϣ�ӱ�(����ID,��Ϣ��,��Ϣֵ) VALUES (����ID_IN,��Ϣ��_IN,��Ϣֵ_IN);
	END IF ;
END zl_������Ϣ�ӱ�_Update;
/

--�����������
INSERT INTO ���㷽ʽ(����,����,����,����,Ӧ�տ�,ȱʡ��־)
SELECT TO_NUMBER(MAX(LPAD(����,2,'0')))+1,'һ��ͨ','YKT',7,0,0
FROM ���㷽ʽ;
INSERT INTO һ��ͨĿ¼(���,����,���㷽ʽ,ҽԺ����,����)
SELECT MAX(���)+1,'����һ��ͨ','һ��ͨ','0001',2
FROM һ��ͨĿ¼;


--Ȩ�޽ű�
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1102,'����','һ��ͨĿ¼',USER,'SELECT');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1102,'����','ְҵ',USER,'SELECT');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1102,'����','���˷�����¼',USER,'SELECT');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1102,'����','������Ϣ�ӱ�',USER,'SELECT');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1102,'����','������ҳ�ӱ�',USER,'SELECT');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1102,'����','zl_������Ϣ_Insert',USER,'EXECUTE');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1102,'����','zl_������Ϣ_Update',USER,'EXECUTE');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1102,'����','zl_������Ϣ_������Ϣ',USER,'EXECUTE');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1102,'����','zl_������Ϣ�ӱ�_Update',USER,'EXECUTE');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1102,'����','zl_���˷�����¼_������',USER,'EXECUTE');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1102,'����','zl_���˷�����¼_�ϴ�',USER,'EXECUTE');

INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1111,'����','���˷�����¼',USER,'SELECT');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1111,'����','������Ϣ�ӱ�',USER,'SELECT');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1111,'����','������ҳ�ӱ�',USER,'SELECT');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1111,'����','zl_������Ϣ_Insert',USER,'EXECUTE');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1111,'����','zl_������Ϣ_Update',USER,'EXECUTE');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1111,'����','zl_������Ϣ_������Ϣ',USER,'EXECUTE');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1111,'����','zl_������Ϣ�ӱ�_Update',USER,'EXECUTE');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1111,'����','zl_���˷�����¼_������',USER,'EXECUTE');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1111,'����','zl_���˷�����¼_�ϴ�',USER,'EXECUTE');

INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1131,'����','һ��ͨĿ¼',USER,'SELECT');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1131,'����','���˷�����¼',USER,'SELECT');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1131,'����','������Ϣ�ӱ�',USER,'SELECT');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1131,'����','zl_������Ϣ_Insert',USER,'EXECUTE');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1131,'����','zl_������Ϣ_Update',USER,'EXECUTE');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1131,'����','zl_������Ϣ_������Ϣ',USER,'EXECUTE');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1131,'����','zl_������Ϣ�ӱ�_Update',USER,'EXECUTE');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1131,'����','zl_���˷�����¼_������',USER,'EXECUTE');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1131,'����','zl_���˷�����¼_�ϴ�',USER,'EXECUTE');

INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1132,'����','һ��ͨĿ¼',USER,'SELECT');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1132,'����','���˷�����¼',USER,'SELECT');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1132,'����','������Ϣ�ӱ�',USER,'SELECT');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1132,'����','zl_������Ϣ_Insert',USER,'EXECUTE');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1132,'����','zl_������Ϣ_Update',USER,'EXECUTE');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1132,'����','zl_������Ϣ_������Ϣ',USER,'EXECUTE');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1132,'����','zl_������Ϣ�ӱ�_Update',USER,'EXECUTE');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1132,'����','zl_���˷�����¼_������',USER,'EXECUTE');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1132,'����','zl_���˷�����¼_�ϴ�',USER,'EXECUTE');

INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1260,'����','һ��ͨĿ¼',USER,'SELECT');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1260,'����','���˷�����¼',USER,'SELECT');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1260,'����','������Ϣ�ӱ�',USER,'SELECT');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1260,'����','������ҳ�ӱ�',USER,'SELECT');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1260,'����','zl_������Ϣ_Insert',USER,'EXECUTE');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1260,'����','zl_������Ϣ_Update',USER,'EXECUTE');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1260,'����','zl_������Ϣ_������Ϣ',USER,'EXECUTE');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1260,'����','zl_������Ϣ�ӱ�_Update',USER,'EXECUTE');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1260,'����','zl_���˷�����¼_������',USER,'EXECUTE');
INSERT INTO zlprogprivs(ϵͳ,���,����,����,������,Ȩ��) VALUES (100,1260,'����','zl_���˷�����¼_�ϴ�',USER,'EXECUTE');








