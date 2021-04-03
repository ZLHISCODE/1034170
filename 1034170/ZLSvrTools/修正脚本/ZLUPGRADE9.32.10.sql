-----------------------------------------------------------------
--为配合产品版本号由9.32.0升为9.32.50(VZLHIS10.22.50)
-----------------------------------------------------------------
--12969
Insert Into zlPrograms(序号,标题,说明,系统,部件) Values(16,'数据输出','',NULL,NULL)
/
Insert Into zlProgFuncs(系统,序号,功能) Values(NULL,16,'基本')
/
Insert Into zlProgFuncs(系统,序号,功能) Values(NULL,16,'Excel输出')
/
Insert Into zlProgFuncs(系统,序号,功能) Values(NULL,16,'打印')
/

Insert Into zlRoleGrant
  (系统, 序号, 角色, 功能)
  Select f.系统, f.序号, r.角色, f.功能
  From zlProgFuncs f, (Select Distinct 角色 From zlRoleGrant) r
  Where f.系统 Is Null And f.序号=16
  Minus
  Select 系统, 序号, 角色, 功能 From zlRoleGrant Where 系统 Is Null And 序号=16;