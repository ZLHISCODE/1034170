---9.40.10����10.29.60ƥ��ʹ��----

--44031 HIS�Զ�����
--���5�ֶ�
Alter Table zlTools.zlFilesUpgrade Add ǿ�Ƹ��� Number(1)
/
Alter Table zlTools.zlFilesUpgrade Add Constraint zlFilesUpgrade_CK_ǿ�Ƹ��� Check (ǿ�Ƹ��� IN(0,1))
/
