Create Index zlMessages_IX_会话ID ON zlMessages(会话ID)
/
Create Index zlMsgState_IX_用户 ON zlMsgState(用户)
/

--------------------------------------------------------------------------------------
--服务器向客户端自动收集文件
---------------------------------------------------------------------------------------
ALTER TABLE  zlClients add(收集标志 number(1) default 0)
/

Alter Table zlClients Add Constraint zlClients_CK_收集标志 Check (收集标志 in(0,1))
/

INSERT INTO zlRegInfo (项目,行号,内容) VALUES ('收集目录',Null,Null)
/

INSERT INTO zlRegInfo (项目,行号,内容) VALUES ('收集类型',Null,'LOG')
/

insert into zlSvrTools(编号,上级,标题,快键,说明) values ('0309','03','站点文件收集','O',Null)
/
