
---9.40.20����10.29.70ƥ��ʹ��----


--45546 HIS�Զ����� �Զ�ע��
Alter Table zlTools.zlFilesUpgrade Add �Զ�ע�� Number(1) DEFAULT 1
/
Alter Table zlTools.zlFilesUpgrade Add Constraint zlFilesUpgrade_CK_�Զ�ע�� Check (�Զ�ע�� IN(0,1))
/