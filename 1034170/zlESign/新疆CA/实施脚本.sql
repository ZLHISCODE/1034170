delete zlParameters t where t.������=90000 and t.ģ�� is null and t.ϵͳ=100;
Insert Into zlParameters(ID,ϵͳ,ģ��,˽��,����,��Ȩ,�̶�,������,������,����ֵ,ȱʡֵ,����˵��)
Select zlParameters_ID.Nextval,100,-Null,-Null,-Null,-Null,-Null,A.* From (
Select ������,������,����ֵ,ȱʡֵ,����˵�� From zlParameters Where 1 = 0 Union All 
Select 90000,'����ǩ��URL','http://124.117.245.71:18080/webServices/authService|4028e48a39dd529a0139dd5c383d0010','0','��ŵ���ǩ������������·��' From Dual ) A;