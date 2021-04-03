-----------------------------------------------------------------
--为配合产品版本号由9.26升为9.27
--从本次起开始使用升级工具升级 
-----------------------------------------------------------------
ALTER TABLE zlTools.zlRPTConds Drop CONSTRAINT zlRPTConds_PK
/
ALTER TABLE zlTools.zlRPTConds Drop CONSTRAINT zlRPTConds_UQ_条件名称
/
ALTER TABLE zlTools.zlRPTConds ADD CONSTRAINT zlRPTConds_PK PRIMARY KEY (报表ID,条件号,参数名) USING INDEX PCTFREE 5
/
ALTER TABLE zlTools.zlRPTConds ADD CONSTRAINT zlRPTConds_UQ_条件名称 UNIQUE (报表ID,条件名称,参数名) USING INDEX PCTFREE 5
/