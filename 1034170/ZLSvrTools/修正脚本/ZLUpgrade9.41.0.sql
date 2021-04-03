--By:ZQ
Alter table zltools.zlClients add 预升时点 Date
/
Alter table zltools.zlClients add 预升完成 Number(1)
/
Alter Table zltools.zlClients Add Constraint zlClients_CK_预升完成 Check (预升完成 IN(0,1))
/
--46211
Drop Table zlPDASynch;
drop table zlStreamTabs;
Drop Procedure Zl_PDASynch_Log;
--41851
Insert Into zlParameters(ID,系统,模块,私有,本机,参数号,参数名,参数值,缺省值,参数说明)
Select zlParameters_ID.NEXTVAL, -NULL,-NULL,1,-NULL,24,'网络断网自动重连',NULL,NULL,'允许在网络断网或者多重网络切换后自动重新连接数据库：0-不检测，1-检测' From Dual
/

--byZT
CREATE TABLE zlTools.zlUpgradeLog(
	系统 NUMBER(5),
	目标版本 VARCHAR2(10),
	类型 NUMBER(1),
	内容 VARCHAR2(200),
	备注 VARCHAR2(200),
	时间 DATE)
	PCTFREE 5 PCTUSED 85
/
ALTER TABLE zlTools.zlUpgradeLog ADD CONSTRAINT  zlUpgradeLog_UQ_时间 Unique (系统,时间) USING INDEX PCTFREE 5
/
ALTER TABLE zlTools.zlUpgradeLog ADD CONSTRAINT  zlUpgradeLog_FK_系统 FOREIGN KEY (系统) REFERENCES zlSystems(编号) ON DELETE CASCADE
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
