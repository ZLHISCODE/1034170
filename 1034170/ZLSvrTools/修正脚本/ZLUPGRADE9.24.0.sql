-------------------------------------------------------------------------------
--�ṹ��������
-------------------------------------------------------------------------------
--���������һЩ����
Drop Index zlProgFuncs_IX���
/
Create Index zlProgFuncs_IX_��� ON zlProgFuncs(ϵͳ,���) PCTFREE 5
/
Create Index zlReports_IX_����ID ON zlReports(����ID) PCTFREE 5
/
Create Index zlRPTSubs_IX_����ID ON zlRPTSubs(����ID) PCTFREE 5
/

--����Ҫ�������ظ����ݲ��ܲ���Լ������˸������ڴ���Լ��ǰ
Delete From zlRPTSubs A Where RowID<(Select Max(RowID) From zlRPTSubs B Where A.��ID=B.��ID And A.����ID=B.����ID And A.���=B.��� And A.����=B.����)
/
Alter Table zlRPTSubs ADD CONSTRAINT zlRPTSubs_PK PRIMARY KEY(��ID,����ID) USING INDEX PCTFREE 10
/

--������ģ��Ķ�Ӧ��ϵ
Create Table zlRPTPuts(
    ����ID NUMBER(18),
	ϵͳ NUMBER(5),
	����ID NUMBER(18),
    ���� VARCHAR2(30))
    PCTFREE 10 PCTUSED 85
/
Alter Table zlRPTPuts ADD CONSTRAINT zlRPTPuts_PK PRIMARY KEY(����ID,ϵͳ,����ID) USING INDEX PCTFREE 10
/
Alter Table zlRPTPuts ADD CONSTRAINT zlRPTPuts_FK_����ID FOREIGN KEY(����ID) REFERENCES zlReports(ID) ON DELETE CASCADE
/
Alter Table zlRPTPuts ADD CONSTRAINT zlRPTPuts_FK_ϵͳ FOREIGN KEY(ϵͳ) REFERENCES zlSystems(���) ON DELETE CASCADE
/
Create Index zlRPTPuts_IX_����ID ON zlRPTPuts(����ID) PCTFREE 5
/


--------------------------------------------------------------------------------------------------------------------------------------------------------------------
--���˺�:�ı�����ϴ����ط���
--8337,8338
--
--���ӱ�
Create Table zlClientScheme(
	������	 number(18),
	�������� varchar2(50),
	�������� varchar2(100),
	����վ	 varchar2(50),
	�û���   varchar2(20))
	PCTFREE 5 PCTUSED 90
/
ALTER TABLE zlClientScheme ADD CONSTRAINT zlClientScheme_PK PRIMARY KEY (������) USING INDEX PCTFREE 5
/

ALTER TABLE zlClientScheme ADD CONSTRAINT zlClientScheme_UQ_�������� UNIQUE (��������) USING INDEX PCTFREE 5
/

Create Table zlClientParaSet(
	������	 number(18),
	����վ varchar2(50),
	�û��� varchar2(20),
	�ָ���־ number(2))
	PCTFREE 5 PCTUSED 90
/


ALTER TABLE zlClientParaSet ADD CONSTRAINT zlClientScheme_UQ_����վ UNIQUE (����վ,�û���,������) USING INDEX PCTFREE 5
/
CREATE INDEX zlClientParaSet_IX_�û���  ON zlClientParaSet(�û���)   PCTFREE 5
/


Alter Table zlClientParaSet Add Constraint zlClientParaSet_CK_�ָ���־ Check (�ָ���־ in(0,1,2))
/

Alter Table zlClientParaSet Add Constraint zlClientParaSet_CK_������ Check (������ IS NOT NULL)
/

ALTER TABLE zlClientParaSet ADD CONSTRAINT  zlClientParaSet_FK_������ FOREIGN KEY (������) REFERENCES zlClientScheme(������) ON DELETE CASCADE
/

Create Table zlClientparaList(
	������	 number(18),
	���	 number(18),
	���	 varchar2(20),
	Ŀ¼     varchar2(1000),
	����     varchar2(50),
	��ֵ     varchar2(2000),
	������Դ number(2),
	����˵�� varchar2(50))
	PCTFREE 5 PCTUSED 90
/

ALTER TABLE zlClientparaList ADD CONSTRAINT zlClientparaList_PK PRIMARY KEY (������,���) USING INDEX PCTFREE 5
/

ALTER TABLE zlClientparaList ADD CONSTRAINT  zlClientparaList_FK_������ FOREIGN KEY (������) REFERENCES zlClientScheme(������) ON DELETE CASCADE
/
--------------------------------------------------------------------------------------------------------------------------------------------------------------------


-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--��Ϊ�����ϴ�������
--8337;8338
------------------------------------------------------------------------------------------------------------------------
DELETE zlprogprivs WHERE ���=15 AND ϵͳ IS NULL
/

UPDATE zlprograms SET ����='���ز�������',˵��='��վ����������ϴ������ء�������ָ�����' WHERE  ���=15 AND ϵͳ IS null
/

Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(NULL,15,'�����ϴ�','�ϴ�վ���Ѿ����úõı��ز�����')
/

Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(NULL,15,'��������','�Ե�ǰ���úõķ��������������ء�')
/

Insert Into zlProgFuncs(ϵͳ,���,����,˵��) Values(NULL,15,'������ָ�','�Ա���ע�����Ϣ���б�����ָ���')
/


------------------------------------------------------------------------------------------------------------------------

-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
Create Or Replace Function Zl_To_Number(Input_In In Varchar2) Return Number Is
  n_Output Number;
Begin
  n_Output := To_Number(Input_In);
  Return n_Output;
Exception
  When Others Then
    Return 0;
End Zl_To_Number;
/
----------------------------------------------------------------------------
--������������
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_zlClientParaSet_Restore(
	������_IN	IN zlClientParaSet.������%type,
	����վ_IN	IN zlClientParaSet.����վ%type,
	�û���_IN	IN zlClientParaSet.�û���%TYPE 
)
IS
	mbyt�û�        number(2);
BEGIN
	BEGIN 
		SELECT DECODE (����վ,NULL ,1,0) INTO mbyt�û�
		FROM zlClientParaSet 
		WHERE ������=������_IN AND ����վ IS NULL AND �û���=�û���_IN AND rownum=1;
	EXCEPTION 
		WHEN OTHERS THEN mbyt�û�:=0;
	END ;
	
	--���Ĺ��ò��ֻ�վ�����Ʋ���
	UPDATE zlClientParaSet SET �ָ���־=0 
	WHERE ������=������_IN AND ����վ=����վ_IN  AND (�û��� IS NULL OR �û���=�û���_IN);
	
	IF mbyt�û�=1 THEN 
		--����˽�в���
		UPDATE zlClientParaSet SET �ָ���־=2
		WHERE ������=������_IN AND ����վ =����վ_IN AND �û���=�û���_IN AND  nvl(�ָ���־,0)<>2;
		IF sql%NOTfound THEN 
			--�����¼
			insert into zlClientParaSet(������,����վ,�û���,�ָ���־) VALUES (������_IN,����վ_IN,�û���_IN,2);
		END IF ;
	END IF ;
END zl_zlClientParaSet_Restore;
/

-------------------------------------------------------------------------------
--ͬ��ʺ���Ȩ
-------------------------------------------------------------------------------
--ͬ���
Create Public Synonym zlRPTPuts for zlRPTPuts
/
Create Public Synonym Zl_To_Number For Zl_To_Number
/
--��Ȩ
Grant Select on zlRPTPuts to Public
/
Grant Execute on Zl_To_Number to Public
/
Begin
	For r_User In(Select ������ From zlSystems) Loop
		Execute Immediate 'Grant Select,Insert,Update,Delete on zlRPTPuts to '||r_User.������||' With Grant Option';
	End Loop;
End;
/

--���˺�:�ı�����ϴ�������
--8337,8338
CREATE PUBLIC SYNONYM zlClientScheme for zlClientScheme
/

CREATE PUBLIC SYNONYM zlClientparaList for zlClientparaList
/

CREATE PUBLIC SYNONYM zlClientParaSet for zlClientParaSet
/

CREATE PUBLIC SYNONYM zl_zlClientParaSet_Restore for zl_zlClientParaSet_Restore
/

GRANT SELECT  ON  zlClientScheme  TO PUBLIC 
/

GRANT SELECT  ON  zlClientparaList  TO PUBLIC 
/

GRANT SELECT  ON  zlClientParaSet  TO PUBLIC 
/

GRANT EXECUTE ON  zl_zlClientParaSet_Restore  TO PUBLIC 
/

--Ȩ������
Begin
	--�����:zlclientparas
	BEGIN 
		--ɾ������ͬ���
		Execute Immediate 'drop PUBLIC SYNONYM zlclientparas';
	EXCEPTION 
		WHEN OTHERS THEN null; 
	END ;

	BEGIN 
		--���ݱ������Ҫ��ֻ�ܱ��ݴ��ű�������ɾ�����ű���Ϊ�����Ժ�����һ����ݵĿ��ܡ�
		Execute Immediate 'alter table zlclientparas rename to zlclientparasBAK';
	EXCEPTION 
		WHEN OTHERS THEN null; 
	END ;

	For r_User In(Select ������ From zlSystems) 
	Loop
		BEGIN 
			Execute Immediate 'Grant Select,Insert,Update,Delete on zlClientScheme to '||r_User.������||' With Grant Option';
			Execute Immediate 'Grant Select,Insert,Update,Delete on zlClientparaList to '||r_User.������||' With Grant Option';
			Execute Immediate 'Grant Select,Insert,Update,Delete on zlClientParaSet to '||r_User.������||' With Grant Option';
		EXCEPTION 
			WHEN OTHERS THEN null; 
		END;
	End Loop;

	FOR r_Role IN (Select DISTINCT ��ɫ FROM zlrolegrant)
	LOOP 
		BEGIN 
			Execute Immediate 'Grant Select,Insert,Update,Delete on zlClientScheme to '||r_Role.��ɫ;
			Execute Immediate 'Grant Select,Insert,Update,Delete on zlClientparaList to '||r_Role.��ɫ;
			Execute Immediate 'Grant Select,Insert,Update,Delete on zlClientParaSet to '||r_Role.��ɫ;
		EXCEPTION 
			WHEN OTHERS THEN null;
		END ;
	END LOOP ;
End;
/