--修改工具表
--注意：以zltools用户登录
--Connect zltools/zltools
ALTER TABLE zlOptions ADD CONSTRAINT 
    zlOptions_PK 
    PRIMARY KEY (参数号) 
    USING INDEX 
    PCTFREE 5 
    STORAGE(INITIAL 256 NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/
