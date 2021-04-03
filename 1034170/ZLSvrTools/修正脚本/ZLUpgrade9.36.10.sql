--27985
Insert Into zlParameters(ID,系统,模块,私有,参数号,参数名,参数值,缺省值,参数说明)
Select zlParameters_ID.NEXTVAL, -NULL,-NULL,1,21,'药品名称显示',NULL,'2','药品名称显示（主界面单据明细、单据输入界面、直接进入的药品选择器时的药品名称显示）：0-显示通用名，1-显示商品名，2-同时显示通用名和商品名' From Dual;

Insert Into zlParameters(ID,系统,模块,私有,参数号,参数名,参数值,缺省值,参数说明)
Select zlParameters_ID.NEXTVAL, -NULL,-NULL,1,22,'输入药品显示',NULL,'0','输入药品显示（通过输入简码方式进入选择器时药品名称的显示）：0-按输入匹配显示，1-固定显示通用名和商品名' From Dual;
