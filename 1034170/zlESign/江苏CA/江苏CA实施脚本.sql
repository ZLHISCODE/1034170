delete zlParameters t where t.������=90000 and t.ģ�� is null and t.ϵͳ=100;
Insert Into zlParameters(ID,ϵͳ,ģ��,˽��,����,��Ȩ,�̶�,������,������,����ֵ,ȱʡֵ,����˵��)
Select zlParameters_ID.Nextval,100,-Null,-Null,-Null,-Null,-Null,A.* From (
Select ������,������,����ֵ,ȱʡֵ,����˵�� From zlParameters Where 1 = 0 Union All 
Select 90000,'����ǩ��URL','http://202.102.85.153:8080/HealthWebService.asmx?WSDL','0','��ŵ���ǩ������������·��' From Dual ) A;