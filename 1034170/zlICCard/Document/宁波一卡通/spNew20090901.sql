CREATE TABLE 消息主表(
	日期 VARCHAR2 (10),
	序列号 NUMBER (10),
	FUNCNAME VARCHAR2 (100),
	URL VARCHAR2 (1000),
	错误 VARCHAR2 (2000),
	标志 NUMBER (1));
--标志:0-未处理;1-正在处理;2-用户放弃;3-失败;9-成功
CREATE TABLE 消息转发(
	日期 VARCHAR2 (10),
	序列号 NUMBER (10),
	行号 NUMBER (5),
	发送数据 VARCHAR2 (2000));
CREATE TABLE 消息接收(
	日期 VARCHAR2 (10),
	序列号 NUMBER (10),
	行号 NUMBER (5),
	接收数据 VARCHAR2 (2000));
	
ALTER TABLE 消息主表 ADD CONSTRAINT 消息主表_PK PRIMARY KEY (日期,序列号);
ALTER TABLE 消息转发 ADD CONSTRAINT 消息转发_PK PRIMARY KEY (日期,序列号,行号);
ALTER TABLE 消息接收 ADD CONSTRAINT 消息接收_PK PRIMARY KEY (日期,序列号,行号);
CREATE SEQUENCE 消息转发_ID MAXVALUE 99999999 START WITH 1;

--插入新的数据或更新原记录的标志
CREATE OR REPLACE PROCEDURE zl_消息主表_Insert(
	日期_IN IN 消息主表.日期%TYPE,
	序列号_IN IN 消息主表.序列号%TYPE,
	FUNCNAME_IN IN 消息主表.FUNCNAME%TYPE,
	URL_IN IN 消息主表.URL%TYPE,
	错误_IN IN 消息主表.错误%TYPE:=NULL,
	标志_IN IN 消息主表.标志%TYPE:=0
)
AS 
BEGIN
	UPDATE 消息主表
	SET FUNCNAME=FUNCNAME_IN,
		URL=URL_IN,
		错误=错误_IN,
		标志=标志_IN
	WHERE 日期=日期_IN AND 序列号=序列号_IN;

	IF SQL%ROWCOUNT =0 THEN 
		INSERT INTO 消息主表
		(日期,序列号,FUNCNAME,URL,错误,标志)
		VALUES 
		(日期_IN,序列号_IN,FUNCNAME_IN,URL_IN,错误_IN,标志_IN);
	END IF ;
END zl_消息主表_Insert; 
/

--成功提取返回数据后删除
CREATE OR REPLACE PROCEDURE zl_消息主表_Delete(
	日期_IN IN 消息主表.日期%TYPE,
	序列号_IN IN 消息主表.序列号%TYPE
)
AS 
BEGIN
	DELETE 消息主表
	WHERE 日期=日期_IN AND 序列号=序列号_IN;
END zl_消息主表_Delete; 
/


CREATE OR REPLACE PROCEDURE zl_消息转发_Delete(
	日期_IN IN 消息转发.日期%TYPE,
	序列号_IN IN 消息转发.序列号%TYPE
)
AS 
BEGIN
	DELETE 消息转发
	WHERE 日期=日期_IN AND 序列号=序列号_IN;
END zl_消息转发_Delete;
/


CREATE OR REPLACE PROCEDURE zl_消息转发_Insert(
	日期_IN IN 消息转发.日期%TYPE,
	序列号_IN IN 消息转发.序列号%TYPE,
	行号_IN IN 消息转发.行号%TYPE,
	发送数据_IN IN 消息转发.发送数据%TYPE
)
AS 
BEGIN
	INSERT INTO 消息转发
	(日期,序列号,行号,发送数据)
	VALUES 
	(日期_IN,序列号_IN,行号_IN,发送数据_IN);
END zl_消息转发_Insert;
/


CREATE OR REPLACE PROCEDURE zl_消息接收_Delete(
	日期_IN IN 消息接收.日期%TYPE,
	序列号_IN IN 消息接收.序列号%TYPE
)
AS 
BEGIN
	DELETE 消息接收
	WHERE 日期=日期_IN AND 序列号=序列号_IN;
END zl_消息接收_Delete;
/


CREATE OR REPLACE PROCEDURE zl_消息接收_Insert(
	日期_IN IN 消息接收.日期%TYPE,
	序列号_IN IN 消息接收.序列号%TYPE,
	行号_IN IN 消息接收.行号%TYPE,
	接收数据_IN IN 消息接收.接收数据%TYPE
)
AS 
BEGIN
	INSERT INTO 消息接收
	(日期,序列号,行号,接收数据)
	VALUES 
	(日期_IN,序列号_IN,行号_IN,接收数据_IN);
END zl_消息接收_Insert;
/


CREATE SEQUENCE LOGID_ID START WITH 1;
CONNECT zlhis/@;
ALTER TABLE 病人信息 ADD 一卡通建档时间 VARCHAR2 (20);
ALTER TABLE 病人信息 ADD 操作类型 VARCHAR2 (10);
ALTER TABLE 病人发卡记录 ADD 旧卡发卡时间 VARCHAR2 (20);
ALTER TABLE 病人发卡记录 ADD 旧卡明码 VARCHAR2 (20);
CREATE OR REPLACE PROCEDURE zl_病人发卡记录_换补卡(
	病人ID_IN IN 病人发卡记录.病人ID%TYPE,
	旧卡号_IN IN 病人发卡记录.旧卡号%TYPE,
	新卡号_IN IN 病人发卡记录.新卡号%TYPE,
	旧卡发卡医院_IN IN 病人发卡记录.旧卡发卡医院%TYPE:=NULL,
	旧卡类型_IN IN 病人发卡记录.旧卡类型%TYPE:=2,
	旧卡明码_IN IN 病人发卡记录.旧卡明码%TYPE:=NULL,
	旧卡发卡时间_IN IN 病人发卡记录.旧卡发卡时间%TYPE:=NULL
)
AS 
BEGIN
	INSERT INTO 病人发卡记录
	(病人ID,旧卡号,旧卡明码,旧卡类型,旧卡发卡医院,旧卡发卡时间,新卡号,发卡时间)
	SELECT 病人ID_IN,旧卡号_IN,旧卡明码_IN,旧卡类型_IN,NVL(旧卡发卡医院_IN,医院编码),NVL(旧卡发卡时间_IN,SYSDATE ),新卡号_IN,SYSDATE 
	FROM 一卡通目录
	WHERE 名称='宁波一卡通';
END zl_病人发卡记录_换补卡;
/

