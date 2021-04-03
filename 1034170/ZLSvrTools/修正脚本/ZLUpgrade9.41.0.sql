--By:ZQ
Alter table zltools.zlClients add Ԥ��ʱ�� Date
/
Alter table zltools.zlClients add Ԥ����� Number(1)
/
Alter Table zltools.zlClients Add Constraint zlClients_CK_Ԥ����� Check (Ԥ����� IN(0,1))
/
--46211
Drop Table zlPDASynch;
drop table zlStreamTabs;
Drop Procedure Zl_PDASynch_Log;
--41851
Insert Into zlParameters(ID,ϵͳ,ģ��,˽��,����,������,������,����ֵ,ȱʡֵ,����˵��)
Select zlParameters_ID.NEXTVAL, -NULL,-NULL,1,-NULL,24,'��������Զ�����',NULL,NULL,'����������������߶��������л����Զ������������ݿ⣺0-����⣬1-���' From Dual
/

--byZT
CREATE TABLE zlTools.zlUpgradeLog(
	ϵͳ NUMBER(5),
	Ŀ��汾 VARCHAR2(10),
	���� NUMBER(1),
	���� VARCHAR2(200),
	��ע VARCHAR2(200),
	ʱ�� DATE)
	PCTFREE 5 PCTUSED 85
/
ALTER TABLE zlTools.zlUpgradeLog ADD CONSTRAINT  zlUpgradeLog_UQ_ʱ�� Unique (ϵͳ,ʱ��) USING INDEX PCTFREE 5
/
ALTER TABLE zlTools.zlUpgradeLog ADD CONSTRAINT  zlUpgradeLog_FK_ϵͳ FOREIGN KEY (ϵͳ) REFERENCES zlSystems(���) ON DELETE CASCADE
/

Create Public Synonym zlUpgradeLog For zlTools.zlUpgradeLog
/
Grant Select,Insert,Delete,Update On zlTools.zlUpgradeLog To Public
/


CREATE OR REPLACE Function zl_Get_MD5(v_Source Varchar2) Return Varchar2 Is
  Raw_Source Raw(128) := Utl_Raw.Cast_To_Raw(v_Source);
  Raw_Cipher Raw(2048);
  Error_In_Input_Buffer_Length Exception;
Begin
  if v_source is null then 
    return '';
  end if ;
  Sys.Dbms_Obfuscation_Toolkit.Md5(Input => Raw_Source, Checksum => Raw_Cipher);
  Return Rawtohex(Raw_Cipher);
End zl_Get_MD5;
/
