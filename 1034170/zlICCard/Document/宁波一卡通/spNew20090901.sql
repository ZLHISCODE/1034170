CREATE TABLE ��Ϣ����(
	���� VARCHAR2 (10),
	���к� NUMBER (10),
	FUNCNAME VARCHAR2 (100),
	URL VARCHAR2 (1000),
	���� VARCHAR2 (2000),
	��־ NUMBER (1));
--��־:0-δ����;1-���ڴ���;2-�û�����;3-ʧ��;9-�ɹ�
CREATE TABLE ��Ϣת��(
	���� VARCHAR2 (10),
	���к� NUMBER (10),
	�к� NUMBER (5),
	�������� VARCHAR2 (2000));
CREATE TABLE ��Ϣ����(
	���� VARCHAR2 (10),
	���к� NUMBER (10),
	�к� NUMBER (5),
	�������� VARCHAR2 (2000));
	
ALTER TABLE ��Ϣ���� ADD CONSTRAINT ��Ϣ����_PK PRIMARY KEY (����,���к�);
ALTER TABLE ��Ϣת�� ADD CONSTRAINT ��Ϣת��_PK PRIMARY KEY (����,���к�,�к�);
ALTER TABLE ��Ϣ���� ADD CONSTRAINT ��Ϣ����_PK PRIMARY KEY (����,���к�,�к�);
CREATE SEQUENCE ��Ϣת��_ID MAXVALUE 99999999 START WITH 1;

--�����µ����ݻ����ԭ��¼�ı�־
CREATE OR REPLACE PROCEDURE zl_��Ϣ����_Insert(
	����_IN IN ��Ϣ����.����%TYPE,
	���к�_IN IN ��Ϣ����.���к�%TYPE,
	FUNCNAME_IN IN ��Ϣ����.FUNCNAME%TYPE,
	URL_IN IN ��Ϣ����.URL%TYPE,
	����_IN IN ��Ϣ����.����%TYPE:=NULL,
	��־_IN IN ��Ϣ����.��־%TYPE:=0
)
AS 
BEGIN
	UPDATE ��Ϣ����
	SET FUNCNAME=FUNCNAME_IN,
		URL=URL_IN,
		����=����_IN,
		��־=��־_IN
	WHERE ����=����_IN AND ���к�=���к�_IN;

	IF SQL%ROWCOUNT =0 THEN 
		INSERT INTO ��Ϣ����
		(����,���к�,FUNCNAME,URL,����,��־)
		VALUES 
		(����_IN,���к�_IN,FUNCNAME_IN,URL_IN,����_IN,��־_IN);
	END IF ;
END zl_��Ϣ����_Insert; 
/

--�ɹ���ȡ�������ݺ�ɾ��
CREATE OR REPLACE PROCEDURE zl_��Ϣ����_Delete(
	����_IN IN ��Ϣ����.����%TYPE,
	���к�_IN IN ��Ϣ����.���к�%TYPE
)
AS 
BEGIN
	DELETE ��Ϣ����
	WHERE ����=����_IN AND ���к�=���к�_IN;
END zl_��Ϣ����_Delete; 
/


CREATE OR REPLACE PROCEDURE zl_��Ϣת��_Delete(
	����_IN IN ��Ϣת��.����%TYPE,
	���к�_IN IN ��Ϣת��.���к�%TYPE
)
AS 
BEGIN
	DELETE ��Ϣת��
	WHERE ����=����_IN AND ���к�=���к�_IN;
END zl_��Ϣת��_Delete;
/


CREATE OR REPLACE PROCEDURE zl_��Ϣת��_Insert(
	����_IN IN ��Ϣת��.����%TYPE,
	���к�_IN IN ��Ϣת��.���к�%TYPE,
	�к�_IN IN ��Ϣת��.�к�%TYPE,
	��������_IN IN ��Ϣת��.��������%TYPE
)
AS 
BEGIN
	INSERT INTO ��Ϣת��
	(����,���к�,�к�,��������)
	VALUES 
	(����_IN,���к�_IN,�к�_IN,��������_IN);
END zl_��Ϣת��_Insert;
/


CREATE OR REPLACE PROCEDURE zl_��Ϣ����_Delete(
	����_IN IN ��Ϣ����.����%TYPE,
	���к�_IN IN ��Ϣ����.���к�%TYPE
)
AS 
BEGIN
	DELETE ��Ϣ����
	WHERE ����=����_IN AND ���к�=���к�_IN;
END zl_��Ϣ����_Delete;
/


CREATE OR REPLACE PROCEDURE zl_��Ϣ����_Insert(
	����_IN IN ��Ϣ����.����%TYPE,
	���к�_IN IN ��Ϣ����.���к�%TYPE,
	�к�_IN IN ��Ϣ����.�к�%TYPE,
	��������_IN IN ��Ϣ����.��������%TYPE
)
AS 
BEGIN
	INSERT INTO ��Ϣ����
	(����,���к�,�к�,��������)
	VALUES 
	(����_IN,���к�_IN,�к�_IN,��������_IN);
END zl_��Ϣ����_Insert;
/


CREATE SEQUENCE LOGID_ID START WITH 1;
CONNECT zlhis/@;
ALTER TABLE ������Ϣ ADD һ��ͨ����ʱ�� VARCHAR2 (20);
ALTER TABLE ������Ϣ ADD �������� VARCHAR2 (10);
ALTER TABLE ���˷�����¼ ADD �ɿ�����ʱ�� VARCHAR2 (20);
ALTER TABLE ���˷�����¼ ADD �ɿ����� VARCHAR2 (20);
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

