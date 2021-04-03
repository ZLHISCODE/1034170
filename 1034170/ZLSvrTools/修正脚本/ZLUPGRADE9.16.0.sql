--增加工作站导航台登录数限制,缺省0表示无限制
Alter Table zlClients Add 连接数 Number(1) default 0
/