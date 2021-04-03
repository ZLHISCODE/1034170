--取消服务器管理工具中的数据归档转移
Delete From zlSvrTools Where 编号='0201'
/

--处理重复的报表参数
Delete From zlRPTPars A Where RowID>(Select Min(RowID) From zlRPTPars B Where B.源ID=A.源ID And B.序号=A.序号)
/
ALTER TABLE zlRPTPars ADD CONSTRAINT zlRPTPars_PK PRIMARY KEY (源ID,序号) USING INDEX PCTFREE 10 STORAGE(INITIAL 32k NEXT 2k PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/