-----------------------------------------------------------------
--Ϊ��ϲ�Ʒ�汾����9.32.0��Ϊ9.32.50(VZLHIS10.22.50)
-----------------------------------------------------------------
--12969
Insert Into zlPrograms(���,����,˵��,ϵͳ,����) Values(16,'�������','',NULL,NULL)
/
Insert Into zlProgFuncs(ϵͳ,���,����) Values(NULL,16,'����')
/
Insert Into zlProgFuncs(ϵͳ,���,����) Values(NULL,16,'Excel���')
/
Insert Into zlProgFuncs(ϵͳ,���,����) Values(NULL,16,'��ӡ')
/

Insert Into zlRoleGrant
  (ϵͳ, ���, ��ɫ, ����)
  Select f.ϵͳ, f.���, r.��ɫ, f.����
  From zlProgFuncs f, (Select Distinct ��ɫ From zlRoleGrant) r
  Where f.ϵͳ Is Null And f.���=16
  Minus
  Select ϵͳ, ���, ��ɫ, ���� From zlRoleGrant Where ϵͳ Is Null And ���=16;