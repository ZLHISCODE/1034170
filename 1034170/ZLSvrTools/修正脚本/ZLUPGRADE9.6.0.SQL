------------------------------------------------------
--工作站点管理
------------------------------------------------------
--先删除关联
alter table zlClientParas drop  CONSTRAINT zlClientParas_FK_上级
/

--删除主键
alter table zlClientParas drop  CONSTRAINT zlClientParas_PK
/

--增加字段

alter table zlClientParas Add(工作站 varchar2(50))
/

--修改数据
UPDATE  zlClientParas a SET a.工作站=(SELECT 工作站 FROM zlclients b WHERE a.ip=b.ip)
/

--删除工作站为Null的记录
DELETE zlclients WHERE 工作站 IS NULL
/

DELETE zlclientParas WHERE 工作站 IS NULL
/

--删除主键
alter table zlClients drop  CONSTRAINT zlClients_PK
/

alter table zlClients modify(说明 varchar2(300))
/

--增加主键
ALTER TABLE zlClients ADD CONSTRAINT 
    zlClients_PK PRIMARY KEY (工作站)
    USING INDEX PCTFREE 5
    STORAGE(INITIAL 512 NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/

ALTER TABLE zlClientParas ADD CONSTRAINT 
    zlClientParas_PK PRIMARY KEY (序号,工作站)
    USING INDEX PCTFREE 5
    STORAGE(INITIAL 512 NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/

CREATE OR REPLACE PROCEDURE zlExecDDL(
    v_Statement IN VARCHAR2,
    v_ReportErr BOOLEAN DEFAULT TRUE) IS
    v_CursorID INTEGER;
BEGIN
    v_CursorID:=sys.DBMS_SQL.OPEN_CURSOR;
    BEGIN
        sys.DBMS_SQL.PARSE(v_CursorID,v_Statement,sys.DBMS_SQL.NATIVE);
    EXCEPTION
        WHEN OTHERS THEN RAISE;
    END;
    sys.DBMS_SQL.CLOSE_CURSOR(v_CursorID);
EXCEPTION
    WHEN OTHERS THEN
    sys.DBMS_SQL.CLOSE_CURSOR(v_CursorID);
    IF v_ReportErr THEN
        RAISE;
    END IF;
END;
/

----------屏蔽掉对IP字段的删除，以保证9.10.0的客户端在9.6.0的服务器上仍可以登录。
--删除IP
--BEGIN 
--	zlExecDDL('ALTER TABLE zlclientparas drop(ip)',false);	
--EXCEPTION
--	WHEN OTHERS THEN
--	NULL;
--END;
--/

--删除过程
DROP PROCEDURE zlExecDDL
/



