----10.33.30---》9.44.30
--73596:刘硕,2014-06-05,产品标识调整
alter table zlTools.zlComponent modify 部件 varchar2(50);
alter table zlTools.zlComponent add(注册产品名称  Varchar2(100),注册产品简名 Varchar2(20),注册产品版本 Varchar2(10));

--73830:张永康,2014-06-11,历史数据转出性能优化
alter table zldatamove add (重建索引范围 NUMBER(1),重建索引间隔 NUMBER(3));
