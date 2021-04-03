
---9.40.20，与10.29.70匹配使用----


--45546 HIS自动升级 自动注册
Alter Table zlTools.zlFilesUpgrade Add 自动注册 Number(1) DEFAULT 1
/
Alter Table zlTools.zlFilesUpgrade Add Constraint zlFilesUpgrade_CK_自动注册 Check (自动注册 IN(0,1))
/