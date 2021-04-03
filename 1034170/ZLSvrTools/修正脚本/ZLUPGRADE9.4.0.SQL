--修改工具表
--注意：以zltools用户登录
--Connect zltools/zltools
grant insert,update,delete on zlXlsDirectory to Public
/
grant insert,update,delete on zlXlsVerify to Public
/

ALTER TABLE zlXlsVerify ADD CONSTRAINT 
	zlXlsVerify_FK_报表号 FOREIGN KEY (系统,报表号) REFERENCES zlXlsDirectory(系统,报表号) ON DELETE CASCADE
/

