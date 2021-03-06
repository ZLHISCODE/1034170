Create Table zlRPTConds(
	报表ID   NUMBER(18),
	条件号   NUMBER(18),
	条件名称 VARCHAR2(200),
	参数名   VARCHAR2(100),
	参数值   VARCHAR2(500),
	CONSTRAINT zlRPTConds_FK_报表ID
        FOREIGN KEY(报表ID) 
        REFERENCES zlReports(ID) 
        ON DELETE CASCADE)
    PCTFREE 10
    PCTUSED 85
    STORAGE (INITIAL 1M NEXT 16K PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/

CREATE OR REPLACE PROCEDURE ZL_RPTCONDS_UPDATE(
    报表ID_IN IN ZLREPORTS.ID%TYPE,
	条件号_IN IN ZLRPTCONDS.条件号%TYPE,
	条件名称_IN IN ZLRPTCONDS.条件名称%TYPE,
	参数_IN IN VARCHAR2:='',						--参数名,参数值,...
	原条件号_IN IN ZLRPTCONDS.条件号%TYPE:=0,		--不为零则删除后插入
	删除_IN IN NUMBER:=0
)
IS
	INTSTART NUMBER(10);
	STR参数名 VARCHAR2(100);
	STR参数值 VARCHAR2(500);
    STR源串 VARCHAR2(2000);
	STR子串 VARCHAR2(2000);
BEGIN
	INTSTART:=1;
    STR源串:=参数_IN;
	
	IF 原条件号_IN<>0 THEN 
		DELETE ZLRPTCONDS WHERE 报表ID=报表ID_IN AND 条件号=原条件号_IN;
	END IF ;
	
	IF 删除_IN=1 THEN 
		DELETE ZLRPTCONDS WHERE 报表ID=报表ID_IN AND 条件号=条件号_IN;
		RETURN ;
	END IF ;

	--分解输入串
    WHILE STR源串 IS NOT NULL LOOP
		--判断是否还需要分解
		IF INSTR(STR源串,'!!',INTSTART)<>0 THEN
			STR子串:=SUBSTR(STR源串,INTSTART,INSTR(STR源串,'!!',INTSTART)-1);
			STR源串:=SUBSTR(STR源串,INSTR(STR源串,'!!',INTSTART)+2);
		ELSE
			STR子串:=SUBSTR(STR源串,INTSTART);
			STR源串:='';
		END IF ;

		STR参数名:=SUBSTR(STR子串,1,INSTR(STR子串,',',1)-1);
		STR参数值:=SUBSTR(STR子串,INSTR(STR子串,',',1)+1);
		INSERT INTO ZLRPTCONDS(报表ID,条件号,条件名称,参数名,参数值)
		VALUES (报表ID_IN,条件号_IN,条件名称_IN,STR参数名,STR参数值);
    END LOOP;

END ZL_RPTCONDS_UPDATE;
/

--公共同义词和权限
Create Public Synonym zlRPTConds for zlRPTConds
/
Create Public Synonym ZL_RPTCONDS_UPDATE for ZL_RPTCONDS_UPDATE
/
Grant Select on zlRPTConds to PUBLIC
/
Grant Execute on ZL_RPTCONDS_UPDATE to PUBLIC
/
