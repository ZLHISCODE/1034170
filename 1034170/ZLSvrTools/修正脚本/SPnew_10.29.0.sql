
---9.40.40，与10.29.190匹配使用----
--73596:刘硕,2014-06-11,产品标识调整
alter table zlTools.zlComponent modify 部件 varchar2(50);
alter table zlTools.zlComponent add(注册产品名称  Varchar2(100),注册产品简名 Varchar2(20),注册产品版本 Varchar2(10));



