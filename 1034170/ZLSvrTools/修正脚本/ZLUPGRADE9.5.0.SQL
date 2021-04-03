------------------------------------------------------
--工作站点管理
------------------------------------------------------
Create Table zlClients(
	IP varchar2(20),
	工作站 varchar2(50),
	CPU varchar2(30),
	内存 varchar2(30),
	硬盘 varchar2(30),
	操作系统 varchar2(50),
	部门 varchar2(50),
	用途 varchar2(50),
	说明 varchar2(50),
	升级标志 number(1) default 0,
	禁止使用 number(1) default 0)
	PCTFREE 5
	PCTUSED 90
	STORAGE (INITIAL 8K NEXT 1K PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/
ALTER TABLE zlClients ADD CONSTRAINT 
    zlClients_PK PRIMARY KEY (IP)
    USING INDEX PCTFREE 5
    STORAGE(INITIAL 512 NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED)

/
Alter Table zlClients Add Constraint 
	zlClients_CK_升级标志 Check (升级标志 in(0,1))
/
Alter Table zlClients Add Constraint 
	zlClients_CK_禁止使用 Check (禁止使用 in(0,1))
/

Create Table zlFilesUpgrade(
	序号 varchar2(50),
	文件类型 number(1),
	文件名 varchar2(50),
	版本号 number(18),
	修改日期 Date)
	PCTFREE 5
	PCTUSED 90
	STORAGE (INITIAL 8K NEXT 1K PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/
ALTER TABLE zlFilesUpgrade ADD CONSTRAINT 
    zlFilesUpgrade_PK PRIMARY KEY (序号)
    USING INDEX PCTFREE 5
    STORAGE(INITIAL 512 NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/
Alter Table zlFilesUpgrade Add Constraint 
	zlFilesUpgrade_CK_文件类型 Check (文件类型 in(0,1,2,3))

/

ALTER TABLE zlFilesUpgrade ADD CONSTRAINT 
    zlFilesUpgrade_UQ_文件名 Unique (文件类型,文件名)
    USING INDEX PCTFREE 5
    STORAGE(INITIAL 512 NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/
Create Table zlClientParas(
	序号 number(18),
	IP varchar2(20),
	类别 varchar2(10),
	用户名 varchar2(20),
	目录 varchar2(500),
	键名 varchar2(60),
	键值 varchar2(1000),
	参数来源 number(2) default 0,
	参数说明 varchar2(50),
	修改标志 number(1) default 0)
	PCTFREE 5
	PCTUSED 90
	STORAGE (INITIAL 8K NEXT 1K PCTINCREASE 0 MAXEXTENTS UNLIMITED)

/
Alter Table zlClientParas Add Constraint 
	zlClientParas_CK_参数来源 Check (参数来源 in(0,1))
/
Alter Table zlClientParas Add Constraint 
	zlClientParas_CK_修改标志 Check (修改标志 in(0,1))
/

ALTER TABLE zlClientParas ADD CONSTRAINT 
    zlClientParas_PK PRIMARY KEY (序号,IP)
    USING INDEX PCTFREE 5
    STORAGE(INITIAL 512 NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED)

/
ALTER TABLE zlClientParas ADD CONSTRAINT 
    zlClientParas_FK_上级 FOREIGN KEY (IP) REFERENCES zlClients(IP) ON DELETE CASCADE
/

Create Table zlPeriods(
    期间 VARCHAR2(6),
    开始日期 Date,
    终止日期 Date)
/

ALTER TABLE zlPeriods ADD CONSTRAINT zlPeriods_PK PRIMARY KEY (期间) USING INDEX
/

declare 
	V_cursor PLS_INTEGER;
	v_return PLS_INTEGER;
	v_SQL VARCHAR2(1000);
	CURSOR V_MatList
	IS
		SELECT distinct 所有者
		  FROM zlsystems;

BEGIN
	FOR V_Fld IN V_MatList LOOP
	    v_SQL:='grant select,Insert,UPdate,Delete on zlClients to '|| V_Fld.所有者|| ' With GRANT Option';
	    V_cursor := DBMS_SQL.Open_cursor ();
	    DBMS_SQL.Parse (V_cursor, v_SQL, DBMS_SQL.Native);
	    DBMS_SQL.Close_cursor (V_cursor);
	
	    
	    v_SQL:='grant select,Insert,UPdate,Delete on ZLCLIENTPARAS to '|| V_Fld.所有者|| ' With GRANT Option';
	    V_cursor := DBMS_SQL.Open_cursor ();
	    DBMS_SQL.Parse (V_cursor, v_SQL, DBMS_SQL.Native);
	    DBMS_SQL.Close_cursor (V_cursor);

	    v_SQL:='grant select,Insert,UPdate,Delete on ZLFILESUPGRADE to '|| V_Fld.所有者|| ' With GRANT Option';
	    V_cursor := DBMS_SQL.Open_cursor ();
	    DBMS_SQL.Parse (V_cursor, v_SQL, DBMS_SQL.Native);
	    DBMS_SQL.Close_cursor (V_cursor);

	    v_SQL:='grant select,Insert,UPdate,Delete on zlPeriods to '|| V_Fld.所有者|| ' With GRANT Option';
	    V_cursor := DBMS_SQL.Open_cursor ();
	    DBMS_SQL.Parse (V_cursor, v_SQL, DBMS_SQL.Native);
	    DBMS_SQL.Close_cursor (V_cursor);
	END LOOP;
end;
/


Create public synonym  zlClients  for zlClients
/
Create public synonym  ZLCLIENTPARAS  for ZLCLIENTPARAS 
/
Create public synonym  ZLFILESUPGRADE  for ZLFILESUPGRADE
/
Create public synonym  zlPeriods  for zlPeriods
/
grant select on zlClients to public
/
grant select on ZLCLIENTPARAS to public
/
grant select on ZLFILESUPGRADE to public
/
grant select on zlPeriods to public
/
--在程序中去掉该授权,在此补上,担心有遗漏
grant select on zlDiaryLog to public
/

insert into zlSvrTools(编号,上级,标题,快键,说明) values ('0307','03','站点部件升级','O',Null)
/
insert into zlSvrTools(编号,上级,标题,快键,说明) values ('0308','03','站点运行控制','O',Null)
/

INSERT INTO zlRegInfo (项目,行号,内容) VALUES ('服务器目录',Null,Null)
/
INSERT INTO zlRegInfo (项目,行号,内容) VALUES ('访问用户',Null,Null)
/
INSERT INTO zlRegInfo (项目,行号,内容) VALUES ('访问密码',Null,Null)
/
-----管理工具授权
Insert Into zlPrograms(序号,标题,说明,系统,部件) Values(10,'导航功能清单','',NULL,NULL)
/
Insert Into zlPrograms(序号,标题,说明,系统,部件) Values(11,'字典管理工具','',NULL,NULL)
/
Insert Into zlPrograms(序号,标题,说明,系统,部件) Values(12,'消息收发工具','',NULL,NULL)
/
Insert Into zlPrograms(序号,标题,说明,系统,部件) Values(13,'系统选项设置','',NULL,NULL)
/
Insert Into zlPrograms(序号,标题,说明,系统,部件) Values(14,'EXCEL报表工具','',NULL,NULL)
/
Insert Into zlPrograms(序号,标题,说明,系统,部件) Values(15,'本地参数上传','',NULL,NULL)
/

Insert Into zlProgFuncs(系统,序号,功能) Values(NULL,10,'基本')
/
Insert Into zlProgFuncs(系统,序号,功能) Values(NULL,11,'基本')
/
Insert Into zlProgFuncs(系统,序号,功能) Values(NULL,12,'基本')
/
Insert Into zlProgFuncs(系统,序号,功能) Values(NULL,12,'发送消息')
/
Insert Into zlProgFuncs(系统,序号,功能) Values(NULL,13,'基本')
/
Insert Into zlProgFuncs(系统,序号,功能) Values(NULL,14,'基本')
/
Insert Into zlProgFuncs(系统,序号,功能) Values(NULL,14,'报表增删')
/
Insert Into zlProgFuncs(系统,序号,功能) Values(NULL,14,'报表计算')
/
Insert Into zlProgFuncs(系统,序号,功能) Values(NULL,14,'所有系统')
/
Insert Into zlProgFuncs(系统,序号,功能) Values(NULL,15,'基本')
/

Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(NULL,10,'基本','ZLTOOLS','ZLPROGRAMS','SELECT')
/
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(NULL,11,'基本','ZLTOOLS','ZLBASECODE','SELECT')
/
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(NULL,12,'发送消息','ZLTOOLS','ZLMESSAGES','INSERT')
/
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(NULL,12,'发送消息','ZLTOOLS','ZLMSGSTATE','INSERT')
/
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(NULL,13,'基本','ZLTOOLS','ZLMENUS','INSERT')
/
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(NULL,14,'报表增删','ZLTOOLS','ZLXLSVERIFY','INSERT')
/
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(NULL,14,'报表增删','ZLTOOLS','ZLXLSVERIFY','UPDATE')
/
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(NULL,14,'报表增删','ZLTOOLS','ZLXLSVERIFY','DELETE')
/
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(NULL,14,'报表增删','ZLTOOLS','ZLXLSDIRECTORY','INSERT')
/
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(NULL,14,'报表增删','ZLTOOLS','ZLXLSDIRECTORY','UPDATE')
/
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(NULL,14,'报表增删','ZLTOOLS','ZLXLSDIRECTORY','DELETE')
/
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(NULL,15,'基本','ZLTOOLS','ZLCLIENTPARAS','INSERT')
/
Insert Into zlProgPrivs(系统,序号,功能,所有者,对象,权限) Values(NULL,15,'基本','ZLTOOLS','ZLCLIENTPARAS','DELETE')
/
--期间表数据
insert into zlPeriods(期间,开始日期,终止日期) values ('200301',to_date('2003-01-01','YYYY-MM-DD'),to_date('2003-01-31','YYYY-MM-DD'))
/
insert into zlPeriods(期间,开始日期,终止日期) values ('200302',to_date('2003-02-01','YYYY-MM-DD'),to_date('2003-02-28','YYYY-MM-DD'))
/
insert into zlPeriods(期间,开始日期,终止日期) values ('200303',to_date('2003-03-01','YYYY-MM-DD'),to_date('2003-03-31','YYYY-MM-DD'))
/
insert into zlPeriods(期间,开始日期,终止日期) values ('200304',to_date('2003-04-01','YYYY-MM-DD'),to_date('2003-04-30','YYYY-MM-DD'))
/
insert into zlPeriods(期间,开始日期,终止日期) values ('200305',to_date('2003-05-01','YYYY-MM-DD'),to_date('2003-05-31','YYYY-MM-DD'))
/
insert into zlPeriods(期间,开始日期,终止日期) values ('200306',to_date('2003-06-01','YYYY-MM-DD'),to_date('2003-06-30','YYYY-MM-DD'))
/
insert into zlPeriods(期间,开始日期,终止日期) values ('200307',to_date('2003-07-01','YYYY-MM-DD'),to_date('2003-07-31','YYYY-MM-DD'))
/
insert into zlPeriods(期间,开始日期,终止日期) values ('200308',to_date('2003-08-01','YYYY-MM-DD'),to_date('2003-08-31','YYYY-MM-DD'))
/
insert into zlPeriods(期间,开始日期,终止日期) values ('200309',to_date('2003-09-01','YYYY-MM-DD'),to_date('2003-09-30','YYYY-MM-DD'))
/
insert into zlPeriods(期间,开始日期,终止日期) values ('200310',to_date('2003-10-01','YYYY-MM-DD'),to_date('2003-10-31','YYYY-MM-DD'))
/
insert into zlPeriods(期间,开始日期,终止日期) values ('200311',to_date('2003-11-01','YYYY-MM-DD'),to_date('2003-11-30','YYYY-MM-DD'))
/
insert into zlPeriods(期间,开始日期,终止日期) values ('200312',to_date('2003-12-01','YYYY-MM-DD'),to_date('2003-12-31','YYYY-MM-DD'))
/
insert into zlPeriods(期间,开始日期,终止日期) values ('200401',to_date('2004-01-01','YYYY-MM-DD'),to_date('2004-01-31','YYYY-MM-DD'))
/
insert into zlPeriods(期间,开始日期,终止日期) values ('200402',to_date('2004-02-01','YYYY-MM-DD'),to_date('2004-02-29','YYYY-MM-DD'))
/
insert into zlPeriods(期间,开始日期,终止日期) values ('200403',to_date('2004-03-01','YYYY-MM-DD'),to_date('2004-03-31','YYYY-MM-DD'))
/
insert into zlPeriods(期间,开始日期,终止日期) values ('200404',to_date('2004-04-01','YYYY-MM-DD'),to_date('2004-04-30','YYYY-MM-DD'))
/
insert into zlPeriods(期间,开始日期,终止日期) values ('200405',to_date('2004-05-01','YYYY-MM-DD'),to_date('2004-05-31','YYYY-MM-DD'))
/
insert into zlPeriods(期间,开始日期,终止日期) values ('200406',to_date('2004-06-01','YYYY-MM-DD'),to_date('2004-06-30','YYYY-MM-DD'))
/
insert into zlPeriods(期间,开始日期,终止日期) values ('200407',to_date('2004-07-01','YYYY-MM-DD'),to_date('2004-07-31','YYYY-MM-DD'))
/
insert into zlPeriods(期间,开始日期,终止日期) values ('200408',to_date('2004-08-01','YYYY-MM-DD'),to_date('2004-08-31','YYYY-MM-DD'))
/
insert into zlPeriods(期间,开始日期,终止日期) values ('200409',to_date('2004-09-01','YYYY-MM-DD'),to_date('2004-09-30','YYYY-MM-DD'))
/
insert into zlPeriods(期间,开始日期,终止日期) values ('200410',to_date('2004-10-01','YYYY-MM-DD'),to_date('2004-10-31','YYYY-MM-DD'))
/
insert into zlPeriods(期间,开始日期,终止日期) values ('200411',to_date('2004-11-01','YYYY-MM-DD'),to_date('2004-11-30','YYYY-MM-DD'))
/
insert into zlPeriods(期间,开始日期,终止日期) values ('200412',to_date('2004-12-01','YYYY-MM-DD'),to_date('2004-12-31','YYYY-MM-DD'))
/

------将所有角色缺省具有所有权限
INSERT INTO ZLROLEGRANT(系统,序号,角色,功能)
	SELECT F.系统,F.序号,R.角色,F.功能
	FROM (SELECT DISTINCT 角色 FROM zlrolegrant ) R,zlProgFuncs F
	WHERE F.系统 IS NULL AND F.序号<100 AND NOT EXISTS(SELECT 角色 FROM zlrolegrant T WHERE T.序号=F.序号 AND T.功能=F.功能 AND ROWNUM<2)
/