delete zlParameters t where t.参数号=90000 and t.模块 is null and t.系统=100;
Insert Into zlParameters(ID,系统,模块,私有,本机,授权,固定,参数号,参数名,参数值,缺省值,参数说明)
Select zlParameters_ID.Nextval,100,-Null,-Null,-Null,-Null,-Null,A.* From (
Select 参数号,参数名,参数值,缺省值,参数说明 From zlParameters Where 1 = 0 Union All 
Select 90000,'电子签名URL','http://202.102.85.153:8080/HealthWebService.asmx?WSDL','0','存放电子签名服务器访问路径' From Dual ) A;