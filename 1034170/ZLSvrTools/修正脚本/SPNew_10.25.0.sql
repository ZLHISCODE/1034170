----10.31.90---》9.42.90
--73596:刘硕,2014-06-18,产品标识调整
alter table zlTools.zlComponent modify 部件 varchar2(50);
alter table zlTools.zlComponent add(注册产品名称  Varchar2(100),注册产品简名 Varchar2(20),注册产品版本 Varchar2(10));



