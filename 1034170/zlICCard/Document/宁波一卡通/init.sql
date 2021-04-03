ALTER TABLE 一卡通目录 MODIFY 医院编码 VARCHAR2 (6);
ALTER TABLE 一卡通目录 ADD CONSTRAINT 一卡通目录_CK_启用 CHECK (启用 IN(0,1,2));
CREATE TABLE 病人发卡记录 (
	病人ID NUMBER (18),
	旧卡号 VARCHAR2 (50),
	旧卡类型 NUMBER (3) DEFAULT 2,
	旧卡发卡医院 VARCHAR2 (50),
	旧卡发卡时间 VARCHAR2 (20),
	旧卡明码 VARCHAR2 (20),
	新卡号 VARCHAR2 (50),
	发卡时间 DATE ,
	上传标志 NUMBER (1) DEFAULT 0);
ALTER TABLE 病人发卡记录 ADD CONSTRAINT 病人发卡记录_PK PRIMARY KEY (病人ID,旧卡号) Using Index Pctfree 0 Tablespace zl9indexhis;

CREATE OR REPLACE PROCEDURE zl_病人发卡记录_发卡(
	病人ID_IN IN 病人发卡记录.病人ID%TYPE,
	新卡号_IN IN 病人发卡记录.新卡号%TYPE
)
AS 
BEGIN
	UPDATE 病人发卡记录
	SET 新卡号=新卡号_IN,
		发卡时间= SYSDATE 
	WHERE 病人ID=病人ID_IN; 
END zl_病人发卡记录_发卡;
/

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

CREATE OR REPLACE PROCEDURE zl_病人发卡记录_上传(
	病人ID_IN IN 病人发卡记录.病人ID%TYPE)
IS 
BEGIN
	UPDATE 病人发卡记录 
	SET 上传标志=1
	WHERE 病人ID=病人ID_IN;
END zl_病人发卡记录_上传;
/

CREATE OR REPLACE PROCEDURE zl_病人信息从表_Update(
	病人ID_IN IN 病人信息从表.病人ID%TYPE,
	信息名_IN IN 病人信息从表.信息名%TYPE,
	信息值_IN IN 病人信息从表.信息值%TYPE)
AS 
BEGIN
	UPDATE 病人信息从表
	SET 信息值=信息值_IN
	WHERE 信息名=信息名_IN;
	IF SQL%ROWCOUNT =0 THEN 
		INSERT INTO 病人信息从表(病人ID,信息名,信息值) VALUES (病人ID_IN,信息名_IN,信息值_IN);
	END IF ;
END zl_病人信息从表_Update;
/

--插入基础数据
INSERT INTO 结算方式(编码,名称,简码,性质,应收款,缺省标志)
SELECT TO_NUMBER(MAX(LPAD(编码,2,'0')))+1,'一卡通','YKT',7,0,0
FROM 结算方式;
INSERT INTO 一卡通目录(编号,名称,结算方式,医院编码,启用)
SELECT MAX(编号)+1,'宁波一卡通','一卡通','0001',2
FROM 一卡通目录;


--权限脚本
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1102,'基本','一卡通目录',USER,'SELECT');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1102,'基本','职业',USER,'SELECT');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1102,'基本','病人发卡记录',USER,'SELECT');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1102,'基本','病人信息从表',USER,'SELECT');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1102,'基本','病案主页从表',USER,'SELECT');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1102,'基本','zl_病人信息_Insert',USER,'EXECUTE');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1102,'基本','zl_病人信息_Update',USER,'EXECUTE');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1102,'基本','zl_病人信息_更新信息',USER,'EXECUTE');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1102,'基本','zl_病人信息从表_Update',USER,'EXECUTE');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1102,'基本','zl_病人发卡记录_换补卡',USER,'EXECUTE');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1102,'基本','zl_病人发卡记录_上传',USER,'EXECUTE');

INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1111,'基本','病人发卡记录',USER,'SELECT');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1111,'基本','病人信息从表',USER,'SELECT');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1111,'基本','病案主页从表',USER,'SELECT');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1111,'基本','zl_病人信息_Insert',USER,'EXECUTE');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1111,'基本','zl_病人信息_Update',USER,'EXECUTE');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1111,'基本','zl_病人信息_更新信息',USER,'EXECUTE');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1111,'基本','zl_病人信息从表_Update',USER,'EXECUTE');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1111,'基本','zl_病人发卡记录_换补卡',USER,'EXECUTE');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1111,'基本','zl_病人发卡记录_上传',USER,'EXECUTE');

INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1131,'基本','一卡通目录',USER,'SELECT');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1131,'基本','病人发卡记录',USER,'SELECT');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1131,'基本','病人信息从表',USER,'SELECT');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1131,'基本','zl_病人信息_Insert',USER,'EXECUTE');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1131,'基本','zl_病人信息_Update',USER,'EXECUTE');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1131,'基本','zl_病人信息_更新信息',USER,'EXECUTE');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1131,'基本','zl_病人信息从表_Update',USER,'EXECUTE');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1131,'基本','zl_病人发卡记录_换补卡',USER,'EXECUTE');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1131,'基本','zl_病人发卡记录_上传',USER,'EXECUTE');

INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1132,'基本','一卡通目录',USER,'SELECT');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1132,'基本','病人发卡记录',USER,'SELECT');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1132,'基本','病人信息从表',USER,'SELECT');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1132,'基本','zl_病人信息_Insert',USER,'EXECUTE');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1132,'基本','zl_病人信息_Update',USER,'EXECUTE');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1132,'基本','zl_病人信息_更新信息',USER,'EXECUTE');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1132,'基本','zl_病人信息从表_Update',USER,'EXECUTE');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1132,'基本','zl_病人发卡记录_换补卡',USER,'EXECUTE');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1132,'基本','zl_病人发卡记录_上传',USER,'EXECUTE');

INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1260,'基本','一卡通目录',USER,'SELECT');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1260,'基本','病人发卡记录',USER,'SELECT');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1260,'基本','病人信息从表',USER,'SELECT');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1260,'基本','病案主页从表',USER,'SELECT');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1260,'基本','zl_病人信息_Insert',USER,'EXECUTE');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1260,'基本','zl_病人信息_Update',USER,'EXECUTE');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1260,'基本','zl_病人信息_更新信息',USER,'EXECUTE');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1260,'基本','zl_病人信息从表_Update',USER,'EXECUTE');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1260,'基本','zl_病人发卡记录_换补卡',USER,'EXECUTE');
INSERT INTO zlprogprivs(系统,序号,功能,对象,所有者,权限) VALUES (100,1260,'基本','zl_病人发卡记录_上传',USER,'EXECUTE');








