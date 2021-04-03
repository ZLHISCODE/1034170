-------------------------------------------------------------------------------
--结构修正部份
-------------------------------------------------------------------------------
--补充和修正一些索引
Drop Index zlProgFuncs_IX序号
/
Create Index zlProgFuncs_IX_序号 ON zlProgFuncs(系统,序号) PCTFREE 5
/
Create Index zlReports_IX_程序ID ON zlReports(程序ID) PCTFREE 5
/
Create Index zlRPTSubs_IX_报表ID ON zlRPTSubs(报表ID) PCTFREE 5
/

--因需要先清理重复数据才能补建约束，因此该语句放在创建约束前
Delete From zlRPTSubs A Where RowID<(Select Max(RowID) From zlRPTSubs B Where A.组ID=B.组ID And A.报表ID=B.报表ID And A.序号=B.序号 And A.功能=B.功能)
/
Alter Table zlRPTSubs ADD CONSTRAINT zlRPTSubs_PK PRIMARY KEY(组ID,报表ID) USING INDEX PCTFREE 10
/

--报表与模块的对应关系
Create Table zlRPTPuts(
    报表ID NUMBER(18),
	系统 NUMBER(5),
	程序ID NUMBER(18),
    功能 VARCHAR2(30))
    PCTFREE 10 PCTUSED 85
/
Alter Table zlRPTPuts ADD CONSTRAINT zlRPTPuts_PK PRIMARY KEY(报表ID,系统,程序ID) USING INDEX PCTFREE 10
/
Alter Table zlRPTPuts ADD CONSTRAINT zlRPTPuts_FK_报表ID FOREIGN KEY(报表ID) REFERENCES zlReports(ID) ON DELETE CASCADE
/
Alter Table zlRPTPuts ADD CONSTRAINT zlRPTPuts_FK_系统 FOREIGN KEY(系统) REFERENCES zlSystems(编号) ON DELETE CASCADE
/
Create Index zlRPTPuts_IX_程序ID ON zlRPTPuts(程序ID) PCTFREE 5
/


--------------------------------------------------------------------------------------------------------------------------------------------------------------------
--刘兴宏:改变参数上传下载方案
--8337,8338
--
--增加表
Create Table zlClientScheme(
	方案号	 number(18),
	方案名称 varchar2(50),
	方案描述 varchar2(100),
	工作站	 varchar2(50),
	用户名   varchar2(20))
	PCTFREE 5 PCTUSED 90
/
ALTER TABLE zlClientScheme ADD CONSTRAINT zlClientScheme_PK PRIMARY KEY (方案号) USING INDEX PCTFREE 5
/

ALTER TABLE zlClientScheme ADD CONSTRAINT zlClientScheme_UQ_方案名称 UNIQUE (方案名称) USING INDEX PCTFREE 5
/

Create Table zlClientParaSet(
	方案号	 number(18),
	工作站 varchar2(50),
	用户名 varchar2(20),
	恢复标志 number(2))
	PCTFREE 5 PCTUSED 90
/


ALTER TABLE zlClientParaSet ADD CONSTRAINT zlClientScheme_UQ_工作站 UNIQUE (工作站,用户名,方案号) USING INDEX PCTFREE 5
/
CREATE INDEX zlClientParaSet_IX_用户名  ON zlClientParaSet(用户名)   PCTFREE 5
/


Alter Table zlClientParaSet Add Constraint zlClientParaSet_CK_恢复标志 Check (恢复标志 in(0,1,2))
/

Alter Table zlClientParaSet Add Constraint zlClientParaSet_CK_方案号 Check (方案号 IS NOT NULL)
/

ALTER TABLE zlClientParaSet ADD CONSTRAINT  zlClientParaSet_FK_方案号 FOREIGN KEY (方案号) REFERENCES zlClientScheme(方案号) ON DELETE CASCADE
/

Create Table zlClientparaList(
	方案号	 number(18),
	序号	 number(18),
	类别	 varchar2(20),
	目录     varchar2(1000),
	键名     varchar2(50),
	键值     varchar2(2000),
	参数来源 number(2),
	参数说明 varchar2(50))
	PCTFREE 5 PCTUSED 90
/

ALTER TABLE zlClientparaList ADD CONSTRAINT zlClientparaList_PK PRIMARY KEY (方案号,序号) USING INDEX PCTFREE 5
/

ALTER TABLE zlClientparaList ADD CONSTRAINT  zlClientparaList_FK_方案号 FOREIGN KEY (方案号) REFERENCES zlClientScheme(方案号) ON DELETE CASCADE
/
--------------------------------------------------------------------------------------------------------------------------------------------------------------------


-------------------------------------------------------------------------------
--数据修正部份
-------------------------------------------------------------------------------
--改为参数上传与下载
--8337;8338
------------------------------------------------------------------------------------------------------------------------
DELETE zlprogprivs WHERE 序号=15 AND 系统 IS NULL
/

UPDATE zlprograms SET 标题='本地参数管理',说明='对站点参数进行上传、下载、备份与恢复操作' WHERE  序号=15 AND 系统 IS null
/

Insert Into zlProgFuncs(系统,序号,功能,说明) Values(NULL,15,'参数上传','上传站点已经配置好的本地参数。')
/

Insert Into zlProgFuncs(系统,序号,功能,说明) Values(NULL,15,'参数下载','对当前配置好的方案参数进行下载。')
/

Insert Into zlProgFuncs(系统,序号,功能,说明) Values(NULL,15,'备份与恢复','对本地注册表信息进行备份与恢复。')
/


------------------------------------------------------------------------------------------------------------------------

-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
Create Or Replace Function Zl_To_Number(Input_In In Varchar2) Return Number Is
  n_Output Number;
Begin
  n_Output := To_Number(Input_In);
  Return n_Output;
Exception
  When Others Then
    Return 0;
End Zl_To_Number;
/
----------------------------------------------------------------------------
--方案正常升级
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_zlClientParaSet_Restore(
	方案号_IN	IN zlClientParaSet.方案号%type,
	工作站_IN	IN zlClientParaSet.工作站%type,
	用户名_IN	IN zlClientParaSet.用户名%TYPE 
)
IS
	mbyt用户        number(2);
BEGIN
	BEGIN 
		SELECT DECODE (工作站,NULL ,1,0) INTO mbyt用户
		FROM zlClientParaSet 
		WHERE 方案号=方案号_IN AND 工作站 IS NULL AND 用户名=用户名_IN AND rownum=1;
	EXCEPTION 
		WHEN OTHERS THEN mbyt用户:=0;
	END ;
	
	--更改公用部分或站点限制部分
	UPDATE zlClientParaSet SET 恢复标志=0 
	WHERE 方案号=方案号_IN AND 工作站=工作站_IN  AND (用户名 IS NULL OR 用户名=用户名_IN);
	
	IF mbyt用户=1 THEN 
		--更改私有部分
		UPDATE zlClientParaSet SET 恢复标志=2
		WHERE 方案号=方案号_IN AND 工作站 =工作站_IN AND 用户名=用户名_IN AND  nvl(恢复标志,0)<>2;
		IF sql%NOTfound THEN 
			--插入记录
			insert into zlClientParaSet(方案号,工作站,用户名,恢复标志) VALUES (方案号_IN,工作站_IN,用户名_IN,2);
		END IF ;
	END IF ;
END zl_zlClientParaSet_Restore;
/

-------------------------------------------------------------------------------
--同义词和授权
-------------------------------------------------------------------------------
--同义词
Create Public Synonym zlRPTPuts for zlRPTPuts
/
Create Public Synonym Zl_To_Number For Zl_To_Number
/
--授权
Grant Select on zlRPTPuts to Public
/
Grant Execute on Zl_To_Number to Public
/
Begin
	For r_User In(Select 所有者 From zlSystems) Loop
		Execute Immediate 'Grant Select,Insert,Update,Delete on zlRPTPuts to '||r_User.所有者||' With Grant Option';
	End Loop;
End;
/

--刘兴宏:改变参数上传与下载
--8337,8338
CREATE PUBLIC SYNONYM zlClientScheme for zlClientScheme
/

CREATE PUBLIC SYNONYM zlClientparaList for zlClientparaList
/

CREATE PUBLIC SYNONYM zlClientParaSet for zlClientParaSet
/

CREATE PUBLIC SYNONYM zl_zlClientParaSet_Restore for zl_zlClientParaSet_Restore
/

GRANT SELECT  ON  zlClientScheme  TO PUBLIC 
/

GRANT SELECT  ON  zlClientparaList  TO PUBLIC 
/

GRANT SELECT  ON  zlClientParaSet  TO PUBLIC 
/

GRANT EXECUTE ON  zl_zlClientParaSet_Restore  TO PUBLIC 
/

--权限修正
Begin
	--处理表:zlclientparas
	BEGIN 
		--删除公共同义词
		Execute Immediate 'drop PUBLIC SYNONYM zlclientparas';
	EXCEPTION 
		WHEN OTHERS THEN null; 
	END ;

	BEGIN 
		--备份表：按周韬要求，只能备份此张表，但不能删除这张表，因为可能以后存在找回数据的可能。
		Execute Immediate 'alter table zlclientparas rename to zlclientparasBAK';
	EXCEPTION 
		WHEN OTHERS THEN null; 
	END ;

	For r_User In(Select 所有者 From zlSystems) 
	Loop
		BEGIN 
			Execute Immediate 'Grant Select,Insert,Update,Delete on zlClientScheme to '||r_User.所有者||' With Grant Option';
			Execute Immediate 'Grant Select,Insert,Update,Delete on zlClientparaList to '||r_User.所有者||' With Grant Option';
			Execute Immediate 'Grant Select,Insert,Update,Delete on zlClientParaSet to '||r_User.所有者||' With Grant Option';
		EXCEPTION 
			WHEN OTHERS THEN null; 
		END;
	End Loop;

	FOR r_Role IN (Select DISTINCT 角色 FROM zlrolegrant)
	LOOP 
		BEGIN 
			Execute Immediate 'Grant Select,Insert,Update,Delete on zlClientScheme to '||r_Role.角色;
			Execute Immediate 'Grant Select,Insert,Update,Delete on zlClientparaList to '||r_Role.角色;
			Execute Immediate 'Grant Select,Insert,Update,Delete on zlClientParaSet to '||r_Role.角色;
		EXCEPTION 
			WHEN OTHERS THEN null;
		END ;
	END LOOP ;
End;
/