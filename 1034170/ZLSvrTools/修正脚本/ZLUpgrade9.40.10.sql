---9.40.10，与10.29.60匹配使用----

--44031 HIS自动升级
--添加5字段
Alter Table zlTools.zlFilesUpgrade Add 强制覆盖 Number(1)
/
Alter Table zlTools.zlFilesUpgrade Add Constraint zlFilesUpgrade_CK_强制覆盖 Check (强制覆盖 IN(0,1))
/
