----------------------------------------------------------------------------------------------------------------
--本脚本支持从ZLHIS+ v10.34.170升级到 v10.34.170
--请以数据空间的所有者登录PLSQL并执行下列脚本
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
---------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------
--结构修正部份
------------------------------------------------------------------------------
--129569:余伟节,2019-12-24,解决三方接口配置
CREATE TABLE 三方接口配置(
  接口名 varchar2(50),
  参数号 Number(3),
  参数名 varchar2(50),
  参数值 varchar2(2000),
  说明 varchar2(200)
  )TABLESPACE zl9BaseItem;

Alter Table 三方接口配置 Add Constraint 三方接口配置_PK Primary Key (接口名,参数号) Using Index Tablespace zl9IndexHis;
Alter Table 三方接口配置 Add Constraint 三方接口配置_UQ_参数名 Unique (接口名,参数名) Using Index Tablespace zl9IndexHis;



------------------------------------------------------------------------------
--数据修正部份
------------------------------------------------------------------------------

--129569:余伟节,2019-12-24,解决三方接口配置
Insert into zlTables(系统,表名,表空间,分类) Values(100,'三方接口配置','ZL9BASEITEM','A1');

-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------


-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--129569:余伟节,2019-12-24,解决三方接口配置
Create Or Replace Procedure Zl_三方接口配置_Update
(
  接口名_In 三方接口配置.接口名%Type,
  参数号_In 三方接口配置.参数号%Type,
  参数名_In 三方接口配置.参数名%Type,
  参数值_In 三方接口配置.参数值%Type,
  说明_In   三方接口配置.说明%Type := Null
) As
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  If 参数值_In Is Null Then
    Delete From 三方接口配置 Where 接口名 = 接口名_In And 参数号 = 参数号_In;
  Else
    Update 三方接口配置 Set 参数值 = 参数值_In, 参数名 = 参数名_In Where 接口名 = 接口名_In And 参数号 = 参数号_In;
    If Sql%RowCount = 0 Then
      Insert Into 三方接口配置
        (接口名, 参数号, 参数名, 参数值, 说明)
      Values
        (接口名_In, 参数号_In, 参数名_In, 参数值_In, 说明_In);
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_三方接口配置_Update;
/

---标准部件信息
EXECUTE Zlfiles_Autoupdate('zl9RegEvent.dll','02F80A59A76B5CE1E93E7EC8A4130780','10.34.170.0014',to_date('2020/4/1 8:58:20','YYYY-MM-DD HH24:MI:SS'),SYSDATE,'1','[APPSOFT]\APPLY','zl9RegEvent','1','挂号部件','1','0','');
EXECUTE Zlfiles_Autoupdate('zlPassInterface.DLL','48733462D056DF56B2EC10B113FADFF0','10.34.170.0015',to_date('2020/4/7 17:00:07','YYYY-MM-DD HH24:MI:SS'),SYSDATE,'0','[PUBLIC]','','1','从ZLCisKernel里提取出来的合理用药接口部件','1','1','');
EXECUTE Zlfiles_Autoupdate('zl9CardSquare.DLL','68EF396A551A6675101405091B7CCD10','10.34.170.0016',to_date('2020/7/1 16:56:11','YYYY-MM-DD HH24:MI:SS'),SYSDATE,'1','[APPSOFT]\APPLY','zl9InExse,zl9BaseItem,zl9CISBase,zl9Peis,zl9PeisManage,zl9Blood','1,21,22,24','一卡通相关部件','1','0','');
EXECUTE Zlfiles_Autoupdate('zl9InPatient.dll','9AF784A7C9CE6A656576C94DF74BCF1E','10.34.170.0016',to_date('2020/7/1 16:55:17','YYYY-MM-DD HH24:MI:SS'),SYSDATE,'1','[APPSOFT]\APPLY','zl9CISJob','1','病人入出管理','1','0','');
EXECUTE Zlfiles_Autoupdate('zl9Patient.dll','B98F59A1A6C919BC2A16A80333E22271','10.34.170.0016',to_date('2020/7/1 16:55:47','YYYY-MM-DD HH24:MI:SS'),SYSDATE,'1','[APPSOFT]\APPLY','zl9Patient','1','病人管理部件','1','0','');
-------------------------------------------------------------------------------

------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.34.170.0016' Where 编号=&n_System;
Commit;