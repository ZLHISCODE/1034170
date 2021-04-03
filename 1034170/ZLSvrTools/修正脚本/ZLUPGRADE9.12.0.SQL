---BEGIN:×Ô¶¯ÌáÐÑ¹¦ÄÜ½Å±¾
Create Table zlNotices(
	ÐòºÅ		NUMBER(5) NOT NULL,
	ÏµÍ³		NUMBER(5),		
	ÌáÐÑÌõ¼þ	VARCHAR2(4000),
	ÌáÐÑÄÚÈÝ	VARCHAR2(250),
	ÌáÐÑ±¨±í	VARCHAR2(50),
	ÌáÐÑÉùÒô	NUMBER(5),
	ÌáÐÑ´°¿Ú	NUMBER(1),
	ÌáÐÑË³Ðò	VARCHAR2(200) DEFAULT '',
	¼ì²éÖÜÆÚ	NUMBER(5),
	ÌáÐÑÖÜÆÚ	NUMBER(5),	
	¿ªÊ¼Ê±¼ä	DATE NOT NULL,
	ÖÕÖ¹Ê±¼ä	DATE)			
	PCTFREE 5
	PCTUSED 90
	STORAGE (INITIAL 512 NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/

Alter Table zlNotices ADD CONSTRAINT 
	zlNotices_PK PRIMARY KEY (ÐòºÅ)
	USING INDEX PCTFREE 5
	STORAGE(INITIAL 256 NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/
Create Table zlNoticeUsr(
	ÌáÐÑÐòºÅ	NUMBER(5) NOT NULL,
	ÌáÐÑ¶ÔÏó	NUMBER(1) DEFAULT 0,	--0-ËùÓÐ;1-Ö¸¶¨ÈËÔ±;2-Ö¸¶¨²¿ÃÅ;3-Ö¸¶¨¹¤×÷Õ¾
	¶ÔÏóÃû³Æ	VARCHAR2(50))			
	PCTFREE 5
	PCTUSED 90
	STORAGE (INITIAL 512 NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/
Create Table zlNoticeRec(
	ÌáÐÑÐòºÅ	NUMBER(5) NOT NULL,	
	ÓÃ»§Ãû		VARCHAR2(30),
	¼ì²éÊ±¼ä	DATE,
	¼ì²é½á¹û	NUMBER(1) DEFAULT 0,	--1±íÊ¾ÓÐÒªÌáÐÑµÄÄÚÈÝ;0±íÊ¾ÎÞÒªÌáÐÑµÄÄÚÈÝ
	ÌáÐÑ±êÖ¾	NUMBER(1) DEFAULT 0,	--1±íÊ¾ÒªÌáÐÑ;0±íÊ¾²»ÌáÐÑ
	ÌáÐÑÊ±¼ä	DATE,
	ÌáÐÑÄÚÈÝ	VARCHAR2(250))			
	PCTFREE 5
	PCTUSED 90
	STORAGE (INITIAL 512 NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/

Alter Table zlNoticeUsr ADD CONSTRAINT 
	zlNoticeUsr_PK PRIMARY KEY (ÌáÐÑÐòºÅ,ÌáÐÑ¶ÔÏó,¶ÔÏóÃû³Æ)
	USING INDEX PCTFREE 5
	STORAGE(INITIAL 256 NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/
Alter Table zlNoticeUsr ADD CONSTRAINT 
	zlNoticeUsr_FK_ÌáÐÑÐòºÅ FOREIGN KEY(ÌáÐÑÐòºÅ) 
	REFERENCES zlNotices(ÐòºÅ) ON DELETE CASCADE
/
Alter Table zlNoticeRec ADD CONSTRAINT 
	zlNoticeRec_PK PRIMARY KEY (ÌáÐÑÐòºÅ,ÓÃ»§Ãû)
	USING INDEX PCTFREE 5
	STORAGE(INITIAL 256 NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/
Alter Table zlNoticeRec ADD CONSTRAINT 
	zlNoticeRec_FK_ÌáÐÑÐòºÅ FOREIGN KEY(ÌáÐÑÐòºÅ) 
	REFERENCES zlNotices(ÐòºÅ) ON DELETE CASCADE
/
---END:×Ô¶¯ÌáÐÑ¹¦ÄÜ½Å±¾


Create Table zlUserRoles(
    ÓÃ»§ varchar2(30),
    ½ÇÉ« varchar2(30),
    ¹ÜÀí number(1))
    PCTFREE 5
    PCTUSED 90
    STORAGE (INITIAL 256K NEXT 8K PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/

ALTER TABLE zlUserRoles ADD CONSTRAINT 
    zlUserRoles_PK UNIQUE (ÓÃ»§,½ÇÉ«)
    USING INDEX PCTFREE 10
    STORAGE(INITIAL 32K NEXT 2K PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/

Alter Table zlReports Add ¶¯Ì¬Ö½ÕÅ NUMBER(1)
/

Alter Table zlReports ADD CONSTRAINT zlReports_CK_¶¯Ì¬Ö½ÕÅ Check(¶¯Ì¬Ö½ÕÅ IN(0,1))
/

--×Ô¶¯ÌáÐÑ¹¦ÄÜ½Å±¾:ÉèÖÃ²Ëµ¥Î»ÖÃ
insert into zlSvrTools(±àºÅ,ÉÏ¼¶,±êÌâ,¿ì¼ü,ËµÃ÷) values ('0504','05','×Ô¶¯ÌáÐÑ','H',Null)
/

----------------------------------------------------------------------------
---  INSERT   for   ZLNOTICES
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_ZLNOTICES_INSERT(
	ÐòºÅ_IN IN ZLNOTICES.ÐòºÅ%TYPE,
	ÏµÍ³_IN IN ZLNOTICES.ÏµÍ³%TYPE,
	ÌáÐÑÌõ¼þ_IN IN ZLNOTICES.ÌáÐÑÌõ¼þ%TYPE,
	ÌáÐÑÄÚÈÝ_IN IN ZLNOTICES.ÌáÐÑÄÚÈÝ%TYPE,
	ÌáÐÑ±¨±í_IN IN ZLNOTICES.ÌáÐÑ±¨±í%TYPE,
	ÌáÐÑÉùÒô_IN IN ZLNOTICES.ÌáÐÑÉùÒô%TYPE,
	ÌáÐÑ´°¿Ú_IN IN ZLNOTICES.ÌáÐÑ´°¿Ú%TYPE,
	¼ì²éÖÜÆÚ_IN IN ZLNOTICES.¼ì²éÖÜÆÚ%TYPE,
	ÌáÐÑÖÜÆÚ_IN IN ZLNOTICES.ÌáÐÑÖÜÆÚ%TYPE,
	¿ªÊ¼Ê±¼ä_IN IN ZLNOTICES.¿ªÊ¼Ê±¼ä%TYPE,
	ÖÕÖ¹Ê±¼ä_IN IN ZLNOTICES.ÖÕÖ¹Ê±¼ä%TYPE,
	ÌáÐÑË³Ðò_IN IN ZLNOTICES.ÌáÐÑË³Ðò%TYPE
)
IS
BEGIN
	Insert Into ZLNOTICES
		(ÐòºÅ,ÏµÍ³,ÌáÐÑÌõ¼þ,ÌáÐÑÄÚÈÝ,ÌáÐÑ±¨±í,ÌáÐÑÉùÒô,ÌáÐÑ´°¿Ú,¼ì²éÖÜÆÚ,ÌáÐÑÖÜÆÚ,¿ªÊ¼Ê±¼ä,ÖÕÖ¹Ê±¼ä,ÌáÐÑË³Ðò)
		VALUES
		(ÐòºÅ_IN,ÏµÍ³_IN,ÌáÐÑÌõ¼þ_IN,ÌáÐÑÄÚÈÝ_IN,ÌáÐÑ±¨±í_IN,ÌáÐÑÉùÒô_IN,ÌáÐÑ´°¿Ú_IN,¼ì²éÖÜÆÚ_IN,ÌáÐÑÖÜÆÚ_IN,¿ªÊ¼Ê±¼ä_IN,ÖÕÖ¹Ê±¼ä_IN,ÌáÐÑË³Ðò_IN);
END ZL_ZLNOTICES_INSERT;
/

----------------------------------------------------------------------------
---  UPDATE   for   ZLNOTICES
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_ZLNOTICES_UPDATE(
	ÐòºÅ_IN IN ZLNOTICES.ÐòºÅ%TYPE,
	ÏµÍ³_IN IN ZLNOTICES.ÏµÍ³%TYPE,
	ÌáÐÑÌõ¼þ_IN IN ZLNOTICES.ÌáÐÑÌõ¼þ%TYPE,
	ÌáÐÑÄÚÈÝ_IN IN ZLNOTICES.ÌáÐÑÄÚÈÝ%TYPE,
	ÌáÐÑ±¨±í_IN IN ZLNOTICES.ÌáÐÑ±¨±í%TYPE,
	ÌáÐÑÉùÒô_IN IN ZLNOTICES.ÌáÐÑÉùÒô%TYPE,
	ÌáÐÑ´°¿Ú_IN IN ZLNOTICES.ÌáÐÑ´°¿Ú%TYPE,
	¼ì²éÖÜÆÚ_IN IN ZLNOTICES.¼ì²éÖÜÆÚ%TYPE,
	ÌáÐÑÖÜÆÚ_IN IN ZLNOTICES.ÌáÐÑÖÜÆÚ%TYPE,
	¿ªÊ¼Ê±¼ä_IN IN ZLNOTICES.¿ªÊ¼Ê±¼ä%TYPE,
	ÖÕÖ¹Ê±¼ä_IN IN ZLNOTICES.ÖÕÖ¹Ê±¼ä%TYPE,
	ÌáÐÑË³Ðò_IN IN ZLNOTICES.ÌáÐÑË³Ðò%TYPE
)
IS
BEGIN
	Update ZLNOTICES
		Set ÐòºÅ=ÐòºÅ_IN,
		    ÏµÍ³=ÏµÍ³_IN,
		    ÌáÐÑÌõ¼þ=ÌáÐÑÌõ¼þ_IN,
		    ÌáÐÑÄÚÈÝ=ÌáÐÑÄÚÈÝ_IN,
		    ÌáÐÑ±¨±í=ÌáÐÑ±¨±í_IN,
		    ÌáÐÑÉùÒô=ÌáÐÑÉùÒô_IN,
		    ÌáÐÑ´°¿Ú=ÌáÐÑ´°¿Ú_IN,
		    ¼ì²éÖÜÆÚ=¼ì²éÖÜÆÚ_IN,
		    ÌáÐÑÖÜÆÚ=ÌáÐÑÖÜÆÚ_IN,
		    ¿ªÊ¼Ê±¼ä=¿ªÊ¼Ê±¼ä_IN,
		    ÖÕÖ¹Ê±¼ä=ÖÕÖ¹Ê±¼ä_IN,
		    ÌáÐÑË³Ðò=ÌáÐÑË³Ðò_IN
		Where  ÐòºÅ=ÐòºÅ_IN;
END ZL_ZLNOTICES_UPDATE;
/

----------------------------------------------------------------------------
---  DELETE   for   ZLNOTICES
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_ZLNOTICES_DELETE(
	ÐòºÅ_IN IN ZLNOTICES.ÐòºÅ%TYPE
)
IS
BEGIN
	Delete From ZLNOTICEUSR Where  ÌáÐÑÐòºÅ=ÐòºÅ_IN;
	Delete From ZLNOTICES Where  ÐòºÅ=ÐòºÅ_IN;
END ZL_ZLNOTICES_DELETE;
/

----------------------------------------------------------------------------
---  INSERT   for   ZLNOTICEUSR
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_ZLNOTICEUSR_INSERT(
	ÌáÐÑÐòºÅ_IN IN ZLNOTICEUSR.ÌáÐÑÐòºÅ%TYPE,
	ÌáÐÑ¶ÔÏó_IN IN ZLNOTICEUSR.ÌáÐÑ¶ÔÏó%TYPE,
	¶ÔÏóÃû³Æ_IN IN ZLNOTICEUSR.¶ÔÏóÃû³Æ%TYPE
)
IS
BEGIN
	Insert Into ZLNOTICEUSR
		(ÌáÐÑÐòºÅ,ÌáÐÑ¶ÔÏó,¶ÔÏóÃû³Æ)
		VALUES
		(ÌáÐÑÐòºÅ_IN,ÌáÐÑ¶ÔÏó_IN,¶ÔÏóÃû³Æ_IN);
END ZL_ZLNOTICEUSR_INSERT;
/
----------------------------------------------------------------------------
---  DELETE   for   ZLNOTICEUSR
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_ZLNOTICEUSR_DELETE(
	ÌáÐÑÐòºÅ_IN IN ZLNOTICEUSR.ÌáÐÑÐòºÅ%TYPE
)
IS
BEGIN
	Delete From ZLNOTICEUSR
		Where  ÌáÐÑÐòºÅ=ÌáÐÑÐòºÅ_IN;
END ZL_ZLNOTICEUSR_DELETE;
/
----------------------------------------------------------------------------
---  ¼ì²é¸üÐÂ
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_ZLNOTICEREC_CHECKNOTICE(
	ÓÃ»§Ãû_IN	IN	zlNoticeUsr.¶ÔÏóÃû³Æ%type,
	²¿ÃÅÃû_IN	IN	zlNoticeUsr.¶ÔÏóÃû³Æ%type:='',
	¹¤×÷Õ¾_IN	IN	zlNoticeUsr.¶ÔÏóÃû³Æ%type:='',
	Æô¶¯¼ì²é_IN	IN	NUMBER:=0)
IS
	Cursor c_Notices IS
		SELECT A.*,B.¼ì²éÊ±¼ä FROM zlNotices A,
			(SELECT ÌáÐÑÐòºÅ,¼ì²éÊ±¼ä FROM zlNoticeRec WHERE ÓÃ»§Ãû=ÓÃ»§Ãû_IN) B 
		WHERE A.ÐòºÅ=B.ÌáÐÑÐòºÅ(+)
			AND ((¼ì²éÖÜÆÚ IS NULL AND 1=Æô¶¯¼ì²é_IN) OR (¼ì²éÖÜÆÚ IS NOT NULL AND 0=Æô¶¯¼ì²é_IN))
			AND A.¿ªÊ¼Ê±¼ä<=SYSDATE AND (A.ÖÕÖ¹Ê±¼ä>=SYSDATE OR A.ÖÕÖ¹Ê±¼ä IS NULL)
			AND (A.ÐòºÅ IN (SELECT ÌáÐÑÐòºÅ FROM zlNoticeUsr
					WHERE (ÌáÐÑ¶ÔÏó = 1 AND ¶ÔÏóÃû³Æ = ÓÃ»§Ãû_IN) 
						OR (ÌáÐÑ¶ÔÏó = 2 AND ¶ÔÏóÃû³Æ = ²¿ÃÅÃû_IN) 
						OR (ÌáÐÑ¶ÔÏó = 3 AND ¶ÔÏóÃû³Æ = ¹¤×÷Õ¾_IN))
			OR A.ÐòºÅ NOT IN (SELECT ÌáÐÑÐòºÅ FROM zlNoticeUsr));

	r_Notice c_Notices%RowType;
	
	v_¼ì²é½á¹û number(1);
	v_ÌáÐÑÄÚÈÝ varchar2(500);
	
	v_¼ì²é·ñ number(1);
	v_CursorID INTEGER;
	v_return INTEGER;
  
	v_Result varchar2(250);
	v_SQL varchar2(4000);

	v_Tmp varchar2(1000);
	v_TmpField varchar2(100);
	v_Pos number(18);
	v_FieldPos number(18);  
	v_FieldType varchar2(50);
	v_Field varchar2(50);
BEGIN
	
	FOR r_Notice In c_Notices Loop
		
		v_¼ì²é·ñ:=0;
				
		--Í¨¹ý¼ì²éÉÏ´Î¼ì²éÊ±¼äÊÇ·ñÎª¿ÕÀ´¼ì²âÊÇ·ñÎªµÚÒ»´Î¼ì²é
		if r_Notice.¼ì²éÊ±¼ä is null then		
			--µÚÒ»´Î¼ì²é,ÐÂÔö¼ÇÂ¼
			insert into zlNoticeRec(ÌáÐÑÐòºÅ,ÓÃ»§Ãû,¼ì²éÊ±¼ä,ÌáÐÑÊ±¼ä,ÌáÐÑÄÚÈÝ) values (r_Notice.ÐòºÅ,ÓÃ»§Ãû_IN,SYSDATE,NULL,NULL);
			v_¼ì²é·ñ:=1;
		else
			--µÚ 2¡¢3¡¢... ´Î¼ì²é
			--µ±Ç°Ê±¼äÊÇ·ñ´óÓÚÉÏ´Î¼ì²éÊ±¼ä¼ÓÉÏÒ»¸ö¼ì²éÖÜÆÚ,Èç¹û´óÓÚÁË£¬Ôò¸üÐÂ¼ì²éÊ±¼ä
			if r_Notice.¼ì²éÖÜÆÚ is null then
				if Æô¶¯¼ì²é_IN=1 then
					update zlNoticeRec set ¼ì²éÊ±¼ä=SYSDATE	where ÌáÐÑÐòºÅ=r_Notice.ÐòºÅ and ÓÃ»§Ãû=ÓÃ»§Ãû_IN;
					v_¼ì²é·ñ:=1;
				end if;
			else
				if SYSDATE>=(r_Notice.¼ì²éÊ±¼ä+r_Notice.¼ì²éÖÜÆÚ/(24*60)) then
					update zlNoticeRec set ¼ì²éÊ±¼ä=SYSDATE	where ÌáÐÑÐòºÅ=r_Notice.ÐòºÅ and ÓÃ»§Ãû=ÓÃ»§Ãû_IN;
					v_¼ì²é·ñ:=1;
				end if;
			end if;
		end if;	
		
		if v_¼ì²é·ñ=1 then
			v_¼ì²é½á¹û:=0;
			v_ÌáÐÑÄÚÈÝ:='';

			--¼ì²éÌáÐÑ		
			if not (r_Notice.ÌáÐÑÌõ¼þ is null) then					
				v_ÌáÐÑÄÚÈÝ:=r_Notice.ÌáÐÑÄÚÈÝ;

				--strTmp¸ñÊ½:Èç'[ÐÕÃû];varchar2|[ÐÔ±ð];date'
				v_Tmp:=r_Notice.ÌáÐÑË³Ðò||'|';
				WHILE not (v_Tmp is null) LOOP

					v_Pos := instr(v_Tmp, '|');								
					v_TmpField:=substr(v_Tmp,1,v_Pos - 1);	
					
					v_FieldPos:=instr(v_TmpField,';');
					v_Field:=substr(v_TmpField,1,v_FieldPos - 1);
					v_FieldType:=trim(Upper(substr(v_TmpField,v_FieldPos+1,100)));

					v_Tmp:=trim(substr(v_Tmp,v_Pos + 1,1000));

					v_Pos:=instr(v_ÌáÐÑÄÚÈÝ,v_Field);

					if v_Pos>0 then
						
						v_Result:=trim(substr(v_Field,2,1000));
						v_Result:=substr(v_Result,1,LENGTH(v_Result)-1);

						if v_FieldType='NUMBER' then
							v_Result:='to_char('||v_Result||')';
						Elsif v_FieldType='DATE' then
							v_Result:='to_char('||v_Result||',''yyyy-mm-dd'')';
						End if;

						v_ÌáÐÑÄÚÈÝ:=trim(substr(v_ÌáÐÑÄÚÈÝ,1,v_Pos - 1)||'''||'||v_Result||'||'''||substr(v_ÌáÐÑÄÚÈÝ,v_Pos + length(v_Field),1000));

					end if;

				END LOOP;
				v_Pos:=instr(Upper(r_Notice.ÌáÐÑÌõ¼þ),' FROM ');
				
				if v_Pos>0 then
					v_SQL:=TRIM('SELECT '''||v_ÌáÐÑÄÚÈÝ||''''||substr(r_Notice.ÌáÐÑÌõ¼þ,v_Pos,4000));

					v_CursorID:=sys.DBMS_SQL.OPEN_CURSOR;
					sys.DBMS_SQL.PARSE(v_CursorID,v_SQL,sys.DBMS_SQL.NATIVE);
					
					dbms_sql.define_column(v_CursorID,1,v_Result,1000);

					v_return :=DBMS_SQL.execute(v_CursorID); 
					
					if DBMS_SQL.FETCH_ROWS(v_CursorID)>0 then
						--¼ì²éºóÓÐÐÂµÄÇé¿ö·¢Éú

						v_¼ì²é½á¹û:=1;	
						dbms_sql.column_value(v_CursorID,1,v_Result);
						v_ÌáÐÑÄÚÈÝ:=trim(v_Result);						
					end if;
				end if;
				
			else
				v_¼ì²é½á¹û:=1;
				v_ÌáÐÑÄÚÈÝ:=r_Notice.ÌáÐÑÄÚÈÝ;
			end if;
			
			update zlNoticeRec set ¼ì²é½á¹û=v_¼ì²é½á¹û,ÌáÐÑÄÚÈÝ=v_ÌáÐÑÄÚÈÝ where ÌáÐÑÐòºÅ=r_Notice.ÐòºÅ and ÓÃ»§Ãû=ÓÃ»§Ãû_IN;
		end if;
	END Loop;

END ZL_ZLNOTICEREC_CHECKNOTICE;
/
----------------------------------------------------------------------------
---  ÌáÐÑ¸üÐÂ
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_ZLNOTICEREC_NOTICE(
	ÓÃ»§Ãû_IN	IN	zlNoticerec.ÓÃ»§Ãû%type,
	Æô¶¯¼ì²é_IN	IN	NUMBER:=0)
IS
	Cursor c_Notices IS
		SELECT B.*,A.ÌáÐÑÖÜÆÚ,A.¼ì²éÖÜÆÚ FROM zlNotices A,
			(SELECT * FROM zlNoticeRec WHERE ÓÃ»§Ãû=ÓÃ»§Ãû_IN) B 
		WHERE A.ÐòºÅ=B.ÌáÐÑÐòºÅ
			AND B.¼ì²é½á¹û=1 
			AND ((Æô¶¯¼ì²é_IN=1 AND ¼ì²éÖÜÆÚ IS NULL) OR (¼ì²éÖÜÆÚ IS NOT NULL AND Æô¶¯¼ì²é_IN=0))
			AND A.¿ªÊ¼Ê±¼ä<=SYSDATE AND (A.ÖÕÖ¹Ê±¼ä>=SYSDATE OR A.ÖÕÖ¹Ê±¼ä IS NULL);

	r_Notice c_Notices%RowType;
  
	v_ÌáÐÑ·ñ number(1);
  
BEGIN
	
	Update zlNoticeRec Set	ÌáÐÑ±êÖ¾=0 where ÓÃ»§Ãû=ÓÃ»§Ãû_IN;
		
	FOR r_Notice In c_Notices Loop
		
		v_ÌáÐÑ·ñ:=0;

		if r_Notice.ÌáÐÑÊ±¼ä is null then
			
			--µÚÒ»´ÎÌáÐÑ
			v_ÌáÐÑ·ñ:=1;
		else
			--µÚ 2¡¢3¡¢... ´ÎÌáÐÑ
			if r_Notice.¼ì²éÖÜÆÚ is null then
				if Æô¶¯¼ì²é_IN=1 then
					v_ÌáÐÑ·ñ:=1;
				end if;
			else
				--µ±Ç°Ê±¼äÊÇ·ñ´óÓÚÉÏ´ÎÌáÐÑÊ±¼ä¼ÓÉÏÒ»¸öÌáÐÑÖÜÆÚ
				if SYSDATE>=(r_Notice.ÌáÐÑÊ±¼ä+r_Notice.ÌáÐÑÖÜÆÚ/(24*60)) then
					v_ÌáÐÑ·ñ:=1;
				end if;
			end if;
		end if;		
		
		if v_ÌáÐÑ·ñ=1 then
			Update zlNoticeRec Set	ÌáÐÑÊ±¼ä=SYSDATE,
						ÌáÐÑ±êÖ¾=v_ÌáÐÑ·ñ
			where ÌáÐÑÐòºÅ=r_Notice.ÌáÐÑÐòºÅ 
				and ÓÃ»§Ãû=ÓÃ»§Ãû_IN;
		end if;

	END Loop;

END ZL_ZLNOTICEREC_NOTICE;
/

Create Or Replace Function zlSpellCode(v_Instr In Varchar2)
	Return Varchar2 Is
	v_Spell   Varchar2(10);
	v_Input   Varchar2(1000);
	v_Bitchar Varchar2(2);
	v_Bitnum  Integer;
	v_Chrnum  Integer;
	v_Stdstr  Varchar2(50) := '°Å²Á´î¶ê·¢¸Á¹þ»÷-¿¦À¬ÂèÄÃÅ¶Å¾ÆÚÈ»ÈöËúÍÚ-ÍÚÎôÑ¹ÔÑ';
	v_Chara   Varchar2(2000) := 'ß¹ï¹åHàÄïÍæXÞßàÈÜtþH×cö°ì\íÁàÉæÈêÓè¨ÙŒø×rèP÷oìaèñâÖÚÏÕYì”ÖOéœõcùgíù“ëˆÛûï§ë@Þîä@áíØtØåB÷öálÛêÝEëJà»âÚéáåÛÖ’÷¡÷éö—úqüÞÖæÁéOá®æñÖ“öËðÆñúòü';
	v_Charb   Varchar2(2000) := 'á±ôÎášØ^÷„ôƒÜØá—ÝÃÝRïT÷Éü–îÙâZÚ•öÑõEå±êþÞãßÂì‹Ù”ívîCÞnÚæÛàîÓô²âkô‘é›ã[ì‡Þkäºß™íDÝòÖræ^ÙèæßìÒöµé–ý_Ýáï’ï–øRødÙ…ìdõÀè˜ãEìsõUètÚéùlãmØÚýã£àfÝKíÕÝíÕRÝ…ä^÷¹öÍêÚßGÙSï¼åQÛÎÛÐÝ™éaàÔìžßJÛMê´éGçaØP÷”æqùSösÝ©Ø°ßÁåþïõÙÂô°Ø„×ß›î¯ßÙÜêáùîéæ¾âØÝÉé[é]åöã¹ÙCÚPääõÏãGésïàŠæÔÞµõIå¨Û‹í@÷ÂèµàˆôÅç@íSí{ÜKÜLôxÚFèEúzú‡ü„í¾ìÔß„æQöýß…öböcØÒÙHíÜøuÛÍâíãêÜÐáŠÞÕçÂérÞgîYÞlÞmÞpÞq×ƒìáè¼ì©÷ÔïRæôûïÚì­ì®ï[Ö€Ù™çSïðïjïkïlïnèsæ»Õ•ål÷§÷B÷Mü‚ý–õ¿ß“ÙÏçÍéÄØhÙeÙfïÙáÙìEè\î éëë÷÷Æ÷ÞôWÙûä‰ÚûêvÙ÷âãuìïžðVÞðÕ@õmìhâÄà£ã\ðGÜ@÷QØÃàRÙñîàâ“ãK÷ˆõÛäcéDõNØmænùPíçè}õËô¤ë¢éÞ×LåÍîßêÎâ˜ÕcðJÞKõ³ß²øGùLûQîÐê³âbÛYà^ð¾ñ£ñ­ñÑñØñÛó÷óëóÙóÖòùñÙñÔñ¹ñ¦ð±ðÇ';
	v_Charc   Varchar2(2000) := 'àêíåßnØ”ÛPï{æî÷õüoôÓè²ÖØ÷û]úIè†Ù‰àÐäîô½ç[Ü³à“üâüá¯ä¹àáè¾âÇã˜ïÊåšæ\âªìxé¶éßÛ‚ïïèdãâæ±îÎâOÙ­Þ{êèäiæ¿åîìøÕSäaâÜäýàšïâàžõðéK×‹èÚÆÝÛÕ~éˆÙæápÖçPêU×€âãåñí]îØöæ½ÝÅãÑè å_é‹öðöKüÜÉéLéMáäæÏä–÷•çL÷lêÆã®ë©äâêÛËÕkíoâ÷ìÌânêËà}ü…ÞCü{ÖšûžÜ‡íºåøÛåÞŠîJÞÓè¡àÁÕ€ÙoÖnÞå·êÚÈÜ•â\ëÖRû‰úmÚ’í×Û{Ù•ö³Úfé´ýYýZÚß×êpèßîõÚWìlîªÚXîdçdçpèKØ©èÇàJÛôîñëóÕ\õ¨ä…õ“ßêí÷àÍæÊÕvø|ùA÷ÎýcüJü[ÜÝÚdÙPßWÚmßgÜ¯õØßtÖsôùãrýXáÜß³âÁÞŒë·à´ï†ÙÑÛLãMë†ßoã‰Ú†ùúuâçÜûô©ã¿ô¾ÛŒê™ï¥ã|Ù±àüã°áOëlöÅÜPá~×‡×‰áhô{ßcéËØŒýiÛ»ØaÚnãIäzërõéúRÜXèÆèúýsýƒØ¡âðç©ØXàsÛUézÕ‘ãÀ÷íÞõà¨àÜõßçÝë°ô­å×ÝŽâ¶îËâAÙiúEêJâëý—Úïé¢é³åNæmîqÝöjùœêÝ»ácåTõžù‡ÙƒÛwõÖåÁÞußOê¡áQÚ}ÝzöºýpèqýwßÚÚeìôÜëÞeâ‘ôÙÞiï“ð@øyÞoú\ú]ÚÙnÜÊèÈæõè®çWäÈçýÕpÙzÙ{Öé¨ëíê£Ýû€û‚û›áÞéãâ§õ¡ÝýÕKÚuÛqõ¾üyõíÜAî•Ùàß¥ïéÜfè‰ìàéÁçJè­Ú~ßýã²ÝÍë¥îxß—ÛZââßuõãáiáÏïóõºûzý€ëâØÈßHï±äSåeå¤ðûö¿óøó×ò¿ò²ñéñåñÝñÒñÎñÃñ¬ó¸ó©ó¤òíòÜòÉñ¡ðîð·æöðËðÈëú';
	v_Chard   Varchar2(2000) := 'ßÕÞÇàªæpÞ…Þ‡æ§âòí³ßQß_ÛQ÷°÷²æ]ÜJèNí^ý‘ý“ß¾Þaá·ß°çªåÊçéÜÜ¤Ü–ÙJÝDõ\øl÷ìÛìOünì^íñÜláGééàîFÙÙü^à¢ÝÌÕQå£ø}ÙœìKÚÔ×[üh×•ÛÊå´í¸ÝÐßTë‹ÚêWØÖß¶âáë®á’÷ô€ê‰þOëIëZìâÜ„ôîï½åuØOàâô£ê­à‡ëQáØíãïëç‹ôÆêÚhàÖïáå~íLçCÙáÝ¶êëì{îEô†ØpûMØµÚ®Û¡êsÛæèÜíÆÝB÷¾öWæ·ÞžÚÐâKé¦íûßfãdíÚßrÖBÛyàÇÛ†áÛîŒîý‚õÚücÚçÛãçèîäâšëŠô¡õõøJõMöôü—õ ùmážîöâyäHëÕ{ä”èSÛìà©Ü¦ÞéÚgëºéPÕ™õÞöøölØêçàôúá”ìwí”ü‡ç–ï}à¤ëëíÖåVîrîûïMäAßËá´ë±õ[üŠöCù…úHÕ‰Ûíá¼ëËÞ“ëØíÏëšÝúêhî×þKâ^àKôYáHékôZäWðLêLô^ô`ôaê^à½á`êAäÂèüë¹ÕiåL÷ò×xØKÚGí~÷Çèoíbíüt×˜Ù€Ü¶ì|åƒæHé²ìÑå‘ÜYîXø‹çŽí¡êŒê íÔí­ïæ×Bç…×míâÜHíïõ»ÜOãçìÀí»ÞšâgîDßqÛvßÍîìâ‡õâõyèIßáç¶ÚrÜoÜ€ôDãõêwêyÛFÛGï˜ùzð¬óýóûóìò½ñõñôñóñ×ñÖñÉñ¼ñ²ñ°ð÷óÎóÆó¼ðãðÛð´';
	v_Chare   Varchar2(2000) := 'åíÞˆÝ­âeï°ÕMä~îPô‰î~ùZù[×Fæ¹ùEêißÀÜÃêqéîÛÑÚÌÜ—ãÕãµÝàØ`ÝQß]ëñïÉß{îOðIØ¬Ö@é‘åŠöùî€ötù˜×†èyý|÷{ÝìÞôíEêzÝ[öÜëXõbøÞWåÇçíîïãsðDßƒÚÙ¦Ù@ÙEð¹ò¦ðÊ';
	v_Charf   Varchar2(2000) := 'áeÛÒéyíÀåzá¦ÞNïcïx÷YâCÞ¬ìÜõìÞÀçxú‹Þxî²ÜèóØœÝGïˆï‰ØÎÚúèÊîÕÚ“â[åpøhöÐô™áÝô³úJåúïwç³ìéìqö­öîöEïyäÇëèã­ì³é¼ôäÕuáôÙMïÐü”çšì]çãÜmâpëƒèûëVôšøXØk÷÷ü‹ØrÞMèMüRüvÙÇö÷å¯÷aããí¿ïLÝ×à•ähØSæ‘çQÛºìbïpüKßôÖSÙºÚRøLøPøiÙˆë€ø]ß‘ß»õÃáKôïïûõÆâaà~ØføWûŸüAüFÙìæÚÜ½ÜÀâöç¦ç¨ÜÞìðî·ÜòÛ®í‚øIíÉÝ³ÙëèõåõÝÊþEãRãVïOøDíhá¥øqÖDÛ~Ý—õHõvíêùfù›ß¼ÞÔàMáœäæÝoôfíëÚâæâØ“ê‚öÖêçÙxÝ•õVÙŽå‡å˜öûövð¥óõò¶òãòðó¾òóòÝðò';
	v_Charg   Varchar2(2000) := 'ê¸Ù¤îÅæÙáåmæØÞÎôpà@ÚëÛòêàëBØdÙWÙ^æYØ¤â}ê®Þ|ÛáãïÜÕôûÞÏøNôvýžä÷ÚséÏß¦÷ ÷hêºí·ç¤äÆÚMêlî¸âGä“æsí°éÀØºízúküŽúê½çÉéÂÞ»æ€Ú¾Û¬ï¯Õaä†ÛÙæüéxømøwÖgøæŠØªàÃÜªë¡ëõéwïÓì‘ík÷ÀÖYÝ‘õsækíuÞPíRöÛÁô´íÑãtßçØ¨ôÞÝ¢âÙûfÙsùˆàQßìç®öáõ†ëÅö¡ÜpýŠýÞÃçîÝ\ì–Ø•ÚCØþçÃâhã^÷¸íxá¸èÛØxÚ¸æÅì°åÜëgêíÙéïÝÔõýÝLÝMôþì±âõYøÝž÷½úXãéÚ¬êôî¹ßEîÜâ’ü‰ØÅëûî­ù]áÄèôêöïÀíådöñöAî™ëÒïNÚoäTïWøŽßÉØÔÚ´ÙÄévêK÷¤êPöŠ÷bÝ„å]ÜIøAÞèäÊØžßkîÂëqæšè…ûX÷}ßÛèæë×Ý_ã üUáîæ£ßžàFé|öÙõqý”ôhôkþIå³âÑØÐê{Ü‰êÐØÛêÁÙF÷¬íW÷Z÷iØ­ÙòçµíÞÝöçõPõ…ÖßÃÛöáÆâuåàþÞâë½Ùåâ£é¤Ý{ðRèJß^óþóôóàóÑòåòäòÁò¼ò´ñøñæñËðáðÙðÀðóð»ð³ð§';
	v_Charh   Varchar2(2000) := 'îþãxàËëÜáVõ°ï™í™ØEõA÷ýÚõêÏìÊäwínØJô_ê\ÝÕâFé\ÞþäIädîhîuÖ›ënå«ú[ôŒÞ†ç¬Ø˜î@ãìÝïàãÞ¶àÆå©×qê»å°î—ö‚Ú­àÀÛÀàAêÂîÁý†Ø€ãFãØ÷…éuûiûîMôçôŸêHíHý[ùŸèYý˜ëaÙRÛÖúQýLìeìfûSìgü\ì•èìçñûaø’ùCÞ¿èUÙêÜŸØFÞ°Ý“åÞZãÈãüÝ¦ØAâvébØDãpìô„äfÞ®ë”ÙäëŸø™üZÚ§é{äUé•é—ô\ýJãô×÷¿æAö\àCááåËàjÜ©Ø_ö×÷õ`÷cìÃéõßüã±Ü ëŒäïëiÖ—àñõúâ©ìÎéÎô–õ­îgì²æLôEö{ù–úCúKä°çúåtöUÙüá²âïìæìïìèà‚øUå×o÷Ÿí_í’÷sûIÕjåkæèîüänÖœçfú†èëÕ–Õ üXõ×øbùJà âµØŽ×’Û¨ä¡ÝÈëfØ}ïÌêaå¾çÙß€ØoæDéIûqÞSêXèG÷ßÝkÛ¼ä½åÕäñöéß§õŒöZödëÁÚòüSáåäÒåØäêéBè«ÖWå–öüÚ‡í‹çuöm÷UúŠÖeæwÚ¶ßÔêÍçõØYëDÝx÷âãÄö™ä§ÜîÞ’ßDõt×eßÜä«Üöí£èíåçà¹çÀê_ÙVÕdÞ¥ÖMî_×M×fçiêTçžìuí}×wîœãÔé’âÆðQÞFý@Ú»äãÕŸïÁØååxß«ß˜îØâ€â·éXØ›þAÖfëoïìàëÞ½èZì[ð©óóóòòºò³ò«ò¥ò¢ñþñüñëñ¥óËó¶ó³ó¨òÂòÀðúð×ðÉð­';
	v_Charj   Varchar2(2000) := 'Ø¢ß´ØÀçáÜ¸í¶ßÒØÞßóåìï|ïúê÷êåõÒøKã‚çÜÙ}Üuì´ÛÔåZëYî¿ÙŠàœëu×Ií‡ù×^çˆÜQíZúaýVèWèiýWûAá§Ø½Ù¥àBþLØCê«éêé®ÝðÚlãšÞªÛeì“ûnÝ‹Ûˆå‰ÞUçgìPúWúnÜeë|ë}Þá÷‚êªáÕ÷äô‡åæÜÁßâä©êéÙÊÛEëHôßÕHõÕö«öÝÕ‚öê÷ÙõJÛ”öaùHýTæ÷õŸ÷DìVö›÷C÷qåÈä¤çìôÂÝçõÊãeïØØjØ†æ‰û“áµÛ£àPí¢ê©îòÛOïäeî]îaø”ùGëÎÙZâ›ê§ÝÑØ]äÕêùégìyÞöçÌÝóØböäøZä’íKû…÷µ÷œùpöx×töúYí[öžè~ídàîèÅõÂíúïµÚÙê¯ôååÀå¿ÖˆôCörûxç‰ç™û{×vû|êðÚÉâVé¥ë¦ëìÚ™éfÙ`ÙÔÕÙvÚ{Û`õÝÖGæIðTæGçZÞYèaèbè{èƒÜüôøçÖ÷šíäí\÷FÖvîŽä®ç­êñánôÝáuÖ˜Ü´æ¯ÜúõÓÙÕöÞõoøŸÞBç€ú„úŒÙ®ÞØäÐë¸Ù]Û]ãqïœáèùa×K÷Rá½Ý^àÝÚŠÞIõ´×_á†ëAà®àµìŒù™ÚàæÝÚ¦ÚµÞ×Þ—èîæ¼ã]ô‚íÙöÚôÉÕmÛdîRæOõ^ï÷ºÕ]ôîÄáŽûvüTÚáÝÀâÛâËéÈèªå\Ö”æ¡Ý£êáßMçÆêîàäÙÚBý„ãþìºÝ¼ëæùXöLù~ù‚ûü û—ÚåØÙëÂÙÓã½îiåòåÉëÖÞŸæºö¦â°ÕeÛVîKìnìoçRØçìçåÄÞ›ïGîyôñãÎà±øF÷ÝôbéNíƒèÑèêÙÖöJûýnúÜÚêÞäé§è¢ôòÚ ï¸öÂÕ‡Ûgä|õLø~÷¶ù‰à`Ý]ÛRÚzÜvé…éÙùVÛžùqúGüŸýAÜìé·é°ö´þFõáýeÚªÜÄßšîÒÙÆêøÛBâ ì«ØeåðõXåáäïZöÄØ‹ÜMõ¶èLä¸ägämïÔæŒùNçîÃïÃäŸèðáúöÁÛ²ëhï…ðCàÙæÞçåáÈèöõûÚbÚ‘ßIØÊÚkâfØãÚÜâ±Þ§ø_ø`àåéÓéQïã×HõêÜBùŠÛÇç~çìßú€ý™ØÜjè‘ÜŠâxãzã—÷÷åå‹õzûŠûŽê}ÞÜðKùQùRùUð¢ðÏðÜðýóÞóÅòÔòÌò»ò±ò¡ñäñÕñÐñÊñÆñÀñ¤ðèðÕóÕóÈóÇñððÔð¯ð¨';
	v_Chark   Varchar2(2000) := 'ßÇØûëÌãlï´é_ç˜ØÜÛîâýê]îøÝÜÝaïÇå|æzêGïaâéæbíèê¬ýÙ©Ý¨Ý|ÝîƒÞRãÛî«êRÜ{ç_÷KØøß’ãÊîÖâ‚é`åêèàîíêûäD÷Šõwõ‘çæéðÚîÝÝVïýâŽî§îWáfîw÷Áá³ã¡ë´æìç¼à¾äÛï¾Õnä˜ØcØ~åoï¬ÕUäLå”çHÙÅáÇÜwÜxåIùyìÜÒíîßµâ@Þ¢údØÚß Ü¥Úœ÷¼õpç«à·ÕFÙ¨ã’ØáÛ¦ßàáöëÚ÷Žà”÷d÷ÅèwÚ²ßßÑÕEÝHÚ¿Ü’ÜœÕNù\ÞÅÚ÷ÛÛæþêÜÙLÝAãkäqà—üYèkã¦êNîåÓàkí–Ø¸à­ÞñêÒî¥î`åžæKÙçÜiõÍíŸÛ“ØÑà°ã´ÝÞÖdçqè^çûï¿÷Õûdõ«åKöïöHù{úAã§ãÍé€éèéîSéŸíAípìHíTôUðâòÒóñóíóØòòòñò¤ñÌñ½ñù';
	v_Charl   Varchar2(2000) := 'ååê¹íÇØÝÞhôFéJö_èníBáÁáâäµà[ßFïªånöDù„üHêãíùÙläþÙ‡îmîsù`ô¥á°ìµïçê@×E×ŽÜ_è|è”íeé­äíî½áYà¥ýœàOÝ¹ïüï¶àHÜqäZæƒãÏÕLéÝõßëáÀï©õ²ç„î‘èáîîã™õuÞLÜ~Øìêbß·ãîí‰÷¦ö˜ðEæÐçÐéÛÙúèDÞ[èhìY÷mýFÚ³ÕC×|èˆûPõªãîLî[åGïKîàÏÜ¨Ûkã¶æêà¬çÊÝñæËØ‚ä‚öâî¾äœÖ‚árÞ¼ß†áëxõ”ç\öPùv÷óèg÷~ûZþGÙµæ²åÎï®ØNýŸä‡å¢õŽõ·÷¯ßŠ÷kß¿ÛÞÜÂìåèÀÙ³èÝÚ\éöÛªáûíÂÝ°à¦ôÏîºõÈö¨äàãWøEë_øtë`ûáBúbû•ÜVÞ]×Þ^ìZ÷uìcÞÆßBöãå¥ì¡ÛšÖ‹æ`×`ôHç ö–çöÝüà˜æ®éçé¬äòåbå€æœönýé£ÞcÜ®õÔÝˆ÷ËôuÝgÕÝvåyÜGàÚå¼â²çÔß|ØIÙ’ÛŽç‚ïmúîÉá‘à€Þ¤éRÞÍßÖÙýä£Þ˜ÛøÞæôóïVõhø•õñôQ÷à÷vßøôÔàëOý á×åàê¥î¬û‹ÞOçl÷ë÷[âÞãÁéÝïCÙUÝþì¢éŠÜCõïÜ\ÜkÞ`àòãöÜßèÚê²û_èùç±ôáÚšÝCâéqÝsë‘ä™ë™õCöìøoûwëëžýhÛ¹öNýgáì`û™ý’êtîIßÊìÖä¯ì¼åÞæòïvïÖûméHöÌæyûˆçBïdçsïiö†úVç¸ï³äÛ‰ìCëwïfôjúwãñÜ×èÐççëÊíÃýˆçXìNýýŽØLÜ[èxì_ûTÛâë]ÚLÙÍà¶ÝäßsÖŒÜ}÷ÃíVáÐïÎçUààß£ÛäãòèÓëÍéñôµâ„öÔô—Þ_èzïB÷|ûRüuûuô”éÖïåæ”çœèuéûê‘äËåÖÙTÝ`äõÚ€Ûjê¤áXä›åhåjè´øšÛÞAçGöIùcùnçeú˜ëªãÌéµé‚úyàLïùëöäXèrèïÙõöÇùFèŽû[á›ï²äsäxàðêÛiÝ†ä—öMÕ“ÞÛîbâ¤ëáé¡ïÝæ ß‰úŸèŒÙÀÜsÙùãøÜýçóÞûäðöÃõiðµðÓóüóöóÒó»ó¹ò÷òëòÛòÈòÃñöñ®ñªñ§ðüðøðìðßðÝðØðÒð½ñïñìñçñÜñÚñÏñÍð¿';
	v_Charm   Varchar2(2000) := 'æÖáïßjæ‹úiö‡è¿éUßéôKö²Ý¤ÙIú”Û½ûœÙuß~ì@ìAî”÷´÷©ôMôNö æžà„Ü¬á£çÏì×ïÜÖ™çNÚøíËâIèšä€äÝØˆêóì¸ÜšáF÷Öå^ùšá¹ã÷ÜâêÄãTë£ÙóÙQà|è£î¦àŽí®Ý®àdáÒäØâ­é¹ïÑäYæ[úBüqä¼ÜzæVüeÚ›ômíi÷ÈÞÑîÍéTéYå{ìËí¯ë‰ÝùÞ«à‘à–ëüíæõ’ô¿ûsìXîŸûLÛÂô»åiãÂöQü€ìDìWÛ_ßäìòâ¨Öi÷ãû†÷çû”ÞÂéSá‚áƒûJáˆØÂåôôÍëßãÚ¢ôéãèåµÚ×à×ü†Ökå²ãæö¼ííäÏëïõ|ìrû ü@üMüIß÷ù‘÷]èÂíðíµç¿åãØ¿ßãøpèf÷xáºçäÜåçëçÅâŒÙ‚ä øsæFãÉãýéhíªüwé}÷ªöšÜøÚ¤àpäéêÔã‘øQî¨õ¤çÑÖ‡ÚÓæÆâÉüN÷áôžÖƒÖ„×OüOéâÜÔï÷Ø{Ýëõöã€ì…ïÒôŽüaõøæŸßèÙ°íøãwÖ\öÊøœüEë¤ãaÛ[ØïãåÛéÜÙîâë‚ãfëŽíJðÅñÇðÌó·ó±ó¡òþòýòúòìòÖòµóúóºò©ñòñ¢';
	v_Charn   Varchar2(2000) := 'ÕyïÕæ“ë~ëÇÞàØvÜ˜Øyâcì„ô›ÜµÞ•áèÍÝÁØ¾åràïà«ßaéªÖQëyôöëîàìôTâÎêÙß­ýQØ«ßÎíÐîóâ®×DçtÛñè§émô[Ú«ðHõƒõàÅâ…äGÛèâõà\îêâ¥ÛCâ‰ØƒÝröòöFûŒýuÙ£ì»ãbëWèXÞ‹êÇíþöÓõRöóùDöTéýÝ‚ÛœÜTØ¥Ûþá|á„ÜàôÁøBæÕëåí±Úíô«êŸà¿ãcÛWÛfÛhåRõææ‡êEÞÁým×‘Übè‡ïDèßÌè_ôVûHØúå¸æ¤âîáðâoìÙ¯ßæÞrÞsáxýP×aæeç×kæÛæååóæÀîÏâSí¤ô¬üQàGÙÐßößSÞùï»ÖZÛåŸð¤ò¨ñ÷ñññÄòïòÍ';
	v_Charo   Varchar2(2000) := 'àÞíMÚ©ê±økÖŽæ–útý{âæñî';
	v_Charp   Varchar2(2000) := 'ÝâèËÙ½Ý‡ßßÝåæWãÝÛAõçÛ˜æoíQãúîGäƒùbè‹ë„äèìQåÌ÷›ý‰ý‹ö„ëãâÒáóÞËÝNìŽûƒüBõ¬êkêŠïÂÙräžàúì·àÎö¬Þ\äÔâñÝJéoàØÜ¡Ý~åAíŠíŽùiôJèmÛsêCØ§ç¢ÚüîëØwâWâtâ”ãYãàèäšåCõBêVêoÜÅèÁÛ¯ÚðÛýØu÷‰î¼ëRô“õQõùùdÜ±âÏØòÛÜã›Õ|øaß¨äÄæÇî¢ê¶úûGêúôæú@æéëÝÙXÕ—õäÚÒÙGÕ›ôØâçÎïgïhôwêQéèî©áoî’àÑæÎë­Ø¯ÜÖçvæ°ØšæÉîlïAé¯êòæ³Ù·îZàZèÒÝZöÒÝƒõGîÇá•áNáwçkÛ¶ÖcØÏîÞãOçêîHïHÞåÙöê·ë¶ê†àÛäõ‹ÙéáTè±å§ïäÙŸçhäßë«ÖEïè×Võëç’ð«ñâñáó¦ó²óáóÍó´óªòçò·ò­ñÈñ±ðå';
	v_Charq   Varchar2(2000) := 'Þ€èçàVÝÂàÒéÊÕƒÛpÖ[ë’õèôtçKù†ØÁÛßáªÜÎêÈä¿Ý½Ú–Ü™âHæëçùç÷ì÷þDèŸí ônôoôëýRÞ­ÛaåW÷’÷¢õšùuù}÷èôGôyö’û˜ßŒá¨Ü»è½ØMç²ôìÖHêMãàÜùÝÝíÓí¬ÝÖÚžáMì—÷ÄÚäÜ·ÙÝá©ã¥Ø@âTâ`ëeí©ãUå¹ûeåºÕßwå½ÖtîvçcùkèBôRôSíaÝ¡îÔÞçÜâjã@ãQäEåXæZübö‘ÛÉëÉã»ç××lècÜÍÜçÙ»èýÝ€ãÞê¨õÄïºïÏäÛ„ïêÛ–æjçIçjæÍéÉÖmôÇìÁíÍàbàzõÎàƒà…ØäÛ^îNçØå æ@Ú‰ÜEÜFèAÜñÚÛã¾÷³éÔ×SÚˆçyíXî˜á ã¸Ú½ê~ÕVímíIÜNæªêüã«ïÆôŠÛoå›ö@çƒôÀÕWîzõÜËâsëdàºäÚì€àßøVéÕÚ_Úcï·äußÄÞììiàWàõÝXÝpöëõ›è[éÑ÷ôÜÜí•Õˆö¥ìmíàõ¼öÆÚöÜäÚ^é±ûjÚ‚öúíFíGöpöqù”÷Gý•áìÙ´ÞåÏá–êäâUÛÏåÙôÃÙgäMábõF÷üõ‰ùj÷AôÜá«Ú°êrìîÕoüLõ@Ú…üDÜ|ôð÷ñö÷OÛ¾ëÔÝ@Þ¡íáøzè³üšÞ¾ë¬áéÜdèŠûYýxÞ‘àTãÖêïé‰üCé˜üzãªçzÚ¹ÜõéúîýÛIÝbãŒÛmêB÷™÷ÜöeýjïEáëî°ç¹íjí¨ã×ãÚÚ|é êIùoåÒð¶òøòéòàòßòÞòÓòÐòËòÇóéóæóäóÜóÌóÀó½òûò°ò¯ñýñûñßñ·ñ³';
	v_Charr   Varchar2(2000) := '÷×ÜÛìüÜ`ôX×j×ŒÜéèãæ¬ëNßvØéâmôã…øžÜóïþÜrØð×šéíâ¿í¥ÜÝØìzì~ígïƒÕJïšÞwê—âJâ~ëÀáõáÉéÅéFægÝPôÛõåÝŠåˆ÷·ökù’íqßï¨ãœønàéå¦Þ¸ø›á}îž÷pàrÞzä²äáçÈÝêøMëÃÜ›Ý‰Þ¨ÜÇèÄî£äJä„écétÙ¼àeö}ö”úUð¦óèòîòÅò¸ò¬ñÅñà';
	v_Chars   Varchar2(2000) := 'ØíìƒØ¦è•ìªëÛâlëMïSàçî|öwÙë§ôLôÖâÌçDédÞúíßærî‹çÒëýöþïbö…÷fÜ£ØÄï¤ëäCÞQçmÖ ïoôOé~ï¡ôÄô‹öèéŒæ|õõßþì¦é„ö®áêßÜÏæ©îÌÛïô®Ü‘áŸé^õÇäúëþõŠê„éWÚ¨ØßÚ]ãˆæóÛ·æÓÖb×iÙ ç—÷­÷W÷XéäõüìØÖ…ôlÛðÙpèlç´ô¹ÝiïYõ}ÜæÛ¿äûâ¦î´ÝfÙdÙhÙÜØÇäÜís÷êÚ·÷“õ˜ù_öYöŸãhö•ß•ßÓïòÚÅ×ŸäÉÕ”îTô×}ëÏé©ävêjê…ê’þJãHå•ü›ù|äÅ×W÷jíòêÉÙKáÓÙ‹ßŸû\âPÝéãAõ§øOø[öõåœöXö‰úPá‡â»ÞyïzìÂÛõÝªßYãJãvöåõZüœ×Rüöˆõ¹âìêêÛéøîæá‹ß±ÚÖÙBÝYâ‹âžï—ßmäKÕœÕžß}ðSÖuáŒö|ýaÞÐô¼á÷ç·æì¯ç£êxæ­Ù¿ÝÄÜ“àgÞóë¨ÛSÛ\Ý”õ_ùeïøÛÓÚHü“÷núž÷tãðëòã_äøØQåfçTùŽùà§ÕXãÅéVäÌþBëpæ×ú{ûtûUç`Õlãßéjí˜ôBÕfÕhåùîåàÊÞ÷ÝôéÃælèpÛÌæùßÐçÁãjïtØËäFïÈäùälæJï\çrúƒýDãáÙîæ¦ìëãôæáÙ¹ï~âLØ|â–ï•âìÚ¡áÂäÁÝ¿áÔã¤ížÕbæànà²äÑâÈì¬ïËágæ}ï`ÛÅàÕî¤Þ´öÕõ‡Ùíä³ÚÕà¼ãºßiûhÝøö¢Úxßpä_ðMÖqÛ‘÷Tú‰â¡Ý´íõî¡å¡ìšëmßUëSÚÇÕrÙwìÝåäçw×\ç›áøÝ¥â¸ïŠöÀé¾æ{úZæ¶êýèøíüàÂôÈÚtõ€ßïæaæiæææ•ßCð£ð¸ðÞððóßóÓóÏóÂóµó°ó§òôòÙóùóâò×òÏòªñêñµðþ';
	v_Chart   Varchar2(2000) := 'õÁîèäâãBÜDõ]÷£öãËßeåÝê`é½ÕwåJìŸêFíOêY×nÜcææÛ¢ìÆõÌöØïUõTÞ·ëÄîÑâØê¼Û°ïÄÕ„á]åUît×TØáv×Zú‚ìþîããgáaêæÙyï¦ôÊïÛÛçMç|íUü‘â¼àoäçëGè©éÌÛ}ÚZõ±æhêOúSàûÙÎéEæ†è’èºï‘ÖzíNíw÷ÒÞä¬ßûìŠá[ì’ä•åcØ»ìýß¯Ø–ï«í«äˆü’ëøß‚Ö`öŒìLäRúeúfç°ç¾ßXÚ„õ®ÖpÛ‡å÷–ø˜î}õ{ùYö[ù•ù—ÜnÜƒÙÃã©åÑßPç‘ÚŒáLìjüVìpî±ãÙø‰êDúcúlãÃéåï›Ùqå`ìtÞÝÙ¬ìöö¶äpì›÷ØöæõæxýföœÕAôÐï¢Úqî\ÝÆÙNÛ@âŸãŽø‡ç“ç”èFï”÷Ñî®ì˜ß‹ÜðæÃÝãéƒöªÖFüžèèîúïFÕPäbîcàÌÙÚÙ¡ÜííÅÚUãPÙ×ãnã~ï ÷‹äüõjâúæBÙï÷»î^üWäŒùWýCÝ±âŠÛTõ©å„ùIùúhú“îÊâQÞƒÜ¢ÝËùrØ‡ÞÒæ˜úoú™î¶åèëPîjîkînôsÛÛƒìÕêÕü`â½ØZëàÜ”ï‚÷ƒôë˜ÙÛØ±×™ï€ô…Ù¢êuÛçãûÞèÞíÈõÉõ¢Û|éÒõDørü˜ö¾üƒâÕùKözèØÚ—ðÃñ»óêóÔó«ó¥òèòÑñíñÓ';
	v_Charw   Varchar2(2000) := 'æ´ü|Øôßœëðící€áËî“ØàæýÜ¹Ø™îBßÝ¸çºëäÝÒçþîµÝnÛläjå†ä[åsÙ–æ~Ú@Øèã¯éþÕsÝy÷ÍÞ‚ÙËåÔêžÚñÝÚìÐÞ±÷˜ögöhàíÛ×àøãíãÇífä¶á¡áÍß`àŒáWå…éõdìSìTì¿çâä¢æ¸ÚÃÚóâ«ôºè¸öÛÕ†ÛcílîQå—õnítï]í|ê¦â¬Ö^åMõKÞEçAìG÷×~ÜZ×ˆÜ^ÞdØnÝ˜Ýœ÷—æ’ö€ö“ãÓâ†ö©ô•øYøjéé”éšü•êZê[ØØãëî‚è·ûlæfúOÝîÞ³ýNÙÁÝ«à¸Ûbë¿á¢ä×íÒö»ý}ÛØÚùÚâEàwÕGÕ_øŒæuöƒàNßíä´ûcùMõˆ÷ùú~ØõåüâÐâèâäåÃêõåqù^ÜRØ£þ@Úãè»ÜÌßAìÉæÄëFì}æðå»Õ`öÈëœýHìFýIúFðÄòêòÚðôðíðÍ';
	v_Charx   Varchar2(2000) := 'ÙâÚÀÛ­ßñÞÉä»ì¤äÀÝ¾ÚTâRôÑôâô¸àqÙÒÕOØgðFæÒðOéØì¨ìäôËåaØGØHØlØ‰ëvõ–ùT×@õµç^ë^êØá@÷ûú è„àEêêÚvÚôÖæˆìI÷žïe÷@çôáãÝßâ|åïÝûãŠìûÖLÖlÛ’÷^Ühâ¾Û§àSâMãÒôªÚiìùÚVëKü_ô]êSìUßÈØBéiïPöyáòèÔêƒíÌê˜åÚè¦ÚYô Ý å’æ_÷ïúTépÕ’ç]ë¯ììôÌÝ²èõÑõ£ã”åßí„åvåwí†õrÛŸ×]úNÜ]÷€æµéeã•ÕtÙtÖPÝá_ûyÚDèvú‘ú’úšÙþáýê“Ú`õÐëUÞºå‚ìÞî‡í`ï@á­ÜÈêˆÕ^ä}åDØRüGçoö±ýEÜ¼à_àlàmç½ÝÙàxæøû‘÷`è‚âÔÛKâÃ÷Ïã}ðAößõaõœí‘÷zí—ç}÷PèÉßØèÕæçç¯åÐû^äìÛXäN÷Ìø{Öyø“újàUáÅÕqÖjÕ[þMÙÉÛÄß¢çÓÖCíPèH×ýšç¥ÙôäÍéÇé¿í…âÝâ³Þ¯åâÛÆÖxíCå¬ý^ýkýKÜaõóâàß”ê¿Ý·âdì§ä\Ü°öÎôgç†êcØ¶ÜŒîˆá…Ö_õSö]ÚêàDè—ê€íÊè™â]ãoã‹ätß©Üôã¬þNÜº×›×œßÝâÓõ÷âÊã–÷ÛæTõxø æ™ïqá¶äåäPçVçnýMíìí¹çï× íšíœôqÕšÖžôzè`ôPÚ¼èòà†ôÚõ¯äªÛÃäÓìãÙ[ã„÷rÞ£ÜŽÚÎÞïÝæêÑìÓÙØÕÖXæM×Xö~×zäöè¯ßxãùìÅîçäÖé¸ãCíÛïàìœïXæ›ÚKÞjíYí´û`õ½ëzú›ÞG÷¨÷LÚÊÚpÖoÛ÷â´Þ¹êÖõ¸á¾âþä­ä±Ü÷à‰öà÷S÷\áßÞ™ÙãßdÙbÞ¦îšèRðªóïóãóÚóÁó­ó¬ó£òáò¹ñãñ¶ðïðçðÂð¼';
	v_Chary   Varchar2(2000) := 'Ñ¹èâè›øfåEø†ùsçŒØóá¬çðíýý\åÂÛëæ«í¼ë²Þëý…âûëÙáÃÝÎäÎëçÛ³æÌáZéŽüiÚ¥ãÆåûÜ¾àIéZééÜîî†û’û}ûšÙðÙ²ÙÈØÉÛ±áDçüßVëCî»÷ÊÜyüdöoùžüfýdýŒüjükî›ýBô|÷úýzüsêÌêšìÍäÙøHõ¦ÚÝ÷ÐøeÖVØÍôeø‘àÙžÜ‚ázú`ÚIÚJ×…á€úŽá‰×—ØVØWãóãZë‡÷±å}ø„êgì¾è–ïráàìÈê–ÖUÝŒåø—ï^ç{öuìRûFÝIðBâóí¦çÛØ²ßºø^Ø³ëÈé÷çòÝUáæßbã“ïuïŸðPôíÖ{Ö|æc÷¥ï_î–öŽèÃáÊø€é™ýoúrì‰ýGê×ú_×Šè€ÞÞîôâXäyæUÚþí“êÊìÇÚËàvà’ØÌÖ]æEædùwìvûEÞvßÞâ¢àcì¥äôãžàæûpát÷ð×búsüpÛÝÚ±âùåÆâÂß×ÜèêÝÞ–íôôýÙOÕBÛDí›ïßzîUîVáÚî{õkÖ–çF×‚û@Þ~îÆÜÓô¯áÞ ârãiøCì½Ý}î‰ÞTýtß®Ø×ØîêdØýß½Ù«á»âøæäÞÈÞÄôàéóã¨ÞÚØ—ê‹ÛüØ[Ø\âNëcÚ˜ÝWâzçËìˆìÚÕxï×ûkûoü]ØæéìÞ²ôèØŠõlÙ“æ„ïîØsìJöGù€ùù‹×g×háyá{èOú^úgÜ²ú…úœ×”ý~êfä¦ë³êŽî÷ê”ê›à³Ü§ãŸìÖNë–éžë íÛóáþâwý‡Û´â¹Õzãyö¸ö¯ý]ýlúÛÈßÅâYâiï‹ëLì‚ï‡Úyë[×ÜáØ·áSõgÝºçøéAàÓÞüÙaè¬ë›øŠâßíŒævúD×súLè]úˆÜ…ú—ûKûWÜãÜþÝÓéºäÞÝöäëÙøÖhå­ÚAÛ«ïIîeçëôì™×Gà¡çßà{Ü­ã¼äVÛÕàaïÞëtçO÷«÷Ó÷Iúxà¯ïJî„öÙ¸÷‘Ûxõ—ákØüßÏû~à›ÞÌèÖÝ¯ÝµÞœà]ßKß[öÏéàâ™÷†ÝjôœÝ’õOßˆØÕÝ¬îðäBë»÷îÙ§àóå¶Þ”ØzáRÕT÷øæúÞ}ê|ßŽæ¥ì£ì¶ô§Ø®ô¨áüÚÄáCâÅÝÇâDêœö§ô~áÎÞíëéè¤êìÕ˜ëkðNôˆÛuÝ›å“Ö~õ‚öVöiú}ûCØñÙ¶àôàöâ×àhèžÕZäoö¹Ø…û‡ýrí²åýâÀêÅîÚþCÚÍßNãÐï„ìÏÝ÷â•îAØ¹ÝhãƒëTßyä`ø\ìÛÖIå[é“øƒø…øˆôráqùO×uÞXçŸìMå÷÷Nú–ûOÜ†ôcôdíóûgøSä‘øxùtüŒüÚOß–ãäë¼Ø’áJâƒö½Ü«ô’à÷éÚÖwÞ@üxæ…ß‡ù úMßRßhÛùæÂÞòè¥îŠë¾Ü‹îáÚ”â_ãXé†é‡éÐÙßå®ügÜSûNý›ûVÚSîfÙšç¡Ü¿êÀàyë…ëµä]êmáñéæâqëEìBýqýyÛ©ã¢àiã³ß\è¹ìÙÙ„ádájíríyíð®ðÐðéðöó¢òöòõòæòâòÕòÊòÄò¾ò£ñÁóîóÛóÄó¿ñ¿ñ¾ñºñ´ñ¯ñ«ñ¨ðùðõðêðàðÖðÎðÁð°';
	v_Charz   Varchar2(2000) := 'Ø´ØÆØÓØëØùÙªÙ¾ÙÌÙÞÚ£Ú¯ÚºÚÁÚÂÚÑÚØÚÚÚÞÚèÚìÚîÛ¤Û¥ÛµÛ¸ÛÚÛúÜÆÜÑÜïÝ§ÝÏÝèÞ©ÞÊÞÙÞêÞøÞýß¡ß¤ßªß¬ß¸ßÆßåßîßðßòßõßùßúàùàýá¤á¿áÌáÑáÖáçâ¯âÍâåâôã·ä¥ä¨ä·ä¾äÃäóåªåÅåéåëæ¢æ¨æÑæÜæàæãæíæïæûç§ç»çÄçÇçÕçÚçÞè°è¶èÌèÎèÏè×èÙèäèåè÷èþé«é»éÆéÍé×éòéôéùéüê¢ê°êµê¾êÃêÞêßêâëÆëÐëÑëÓëÕëÞëêëùì¹ìÄìíìñìóìõìúí§í½íÄíÎíØíÝíéíöî³îÀîÈîÛîùï£ï­ïÅïßïíïñïôïöð¡ð²ðºðÑðäðæðëðññ©ñ¸ñÞñèò§òÆòÎòØó®ó¯óÃóÉóÊóÐóÝóåóçóðô¢ô¦ô±ô¶ô·ôÒôÕôØôãôêôõô÷ôüõ¥õÅõÙõÜõàõîõòõôõþö£ö¤ö·öÉöíöö÷®÷Ú÷æ÷þ';

Begin
	If v_Instr Is Null Or Length(Ltrim(v_Instr)) = 0 Then
		v_Spell := '';
	Else
		v_Input := Upper(v_Instr);
		v_Spell := '';
		For v_Bitnum In 1 .. Length(v_Input) Loop
			v_Bitchar := Substr(v_Input, v_Bitnum, 1);
			If v_Bitchar >= '°¡' And v_Bitchar <= '×ù' Then
				For v_Chrnum In 1 .. Length(v_Stdstr) Loop
					If Substr(v_Stdstr, v_Chrnum, 1) = '-' Then
						Null;
					Elsif v_Bitchar < Substr(v_Stdstr, v_Chrnum, 1) Then
						v_Spell := v_Spell || Chr(64 + v_Chrnum);
						Exit;
					End If;
				End Loop;
				If v_Bitchar >= 'ÔÑ' Then
					v_Spell := v_Spell || 'Z';
				End If;
			Elsif Instr('ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.+-*/', v_Bitchar) > 0 Then
				v_Spell := v_Spell || v_Bitchar;
			Elsif Instr('¢ñ¢ò¢ó¢ô¢õ¢ö¢ø¢ø¢ù', v_Bitchar) > 0 Then
				v_Spell := v_Spell || Chr(Ascii(v_Bitchar) - 41664);
			Elsif Instr('£Á£Â£Ã£Ä£Å£Æ£Ç£È£É£Ê£Ë£Ì£Í£Î£Ï£Ð£Ñ£Ò£Ó£Ô£Õ£Ö£×£Ø£Ù£Ú',v_Bitchar) > 0 Then
				v_Spell := v_Spell || Chr(Ascii(v_Bitchar) - 41856);
			Elsif Instr('¦¡¦Á', v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'A';
			Elsif Instr('¦¢¦Â', v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'B';
			Elsif Instr('¦£¦Ã', v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'G';
			Elsif Instr(v_Chara, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'A';
			Elsif Instr(v_Charb, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'B';
			Elsif Instr(v_Charc, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'C';
			Elsif Instr(v_Chard, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'D';
			Elsif Instr(v_Chare, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'E';
			Elsif Instr(v_Charf, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'F';
			Elsif Instr(v_Charg, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'G';
			Elsif Instr(v_Charh, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'H';
			Elsif Instr(v_Charj, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'J';
			Elsif Instr(v_Chark, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'K';
			Elsif Instr(v_Charl, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'L';
			Elsif Instr(v_Charm, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'M';
			Elsif Instr(v_Charn, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'N';
			Elsif Instr(v_Charo, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'O';
			Elsif Instr(v_Charp, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'P';
			Elsif Instr(v_Charq, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'Q';
			Elsif Instr(v_Charr, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'R';
			Elsif Instr(v_Chars, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'S';
			Elsif Instr(v_Chart, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'T';
			Elsif Instr(v_Charw, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'W';
			Elsif Instr(v_Charx, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'X';
			Elsif Instr(v_Chary, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'Y';
			Elsif Instr(v_Charz, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'Z';
--			Else
--				v_Spell := v_Spell || '_';
			End If;
			Exit When Length(v_Spell) > 9;
		End Loop;
	End If;
	Return(v_Spell);
End;
/

Create Or Replace Function zlWBCode(v_Instr In Varchar2)
	Return Varchar2 Is
	v_Code    Varchar2(10);
	v_Input   Varchar2(1000);
	v_Bitnum  Integer;
	v_a       Varchar2(1200) := '°ª°¬°°°ÅÜØÝÃ°ÐÝò°úÝáÝí÷¹±½Ý©ÜêÝÉ±Í±ÎÞµ±ÞØÒÜÐ²¤±¡²Ë²Ì²Ô²ØÜ³²Ý²ç²èÝÛÝÅÜÉ³¼ÜÝÜûÝ»Üë´ÄÜÊ´ÐÝýÝÍ÷°÷²ß°ÝÌµ´ÝÐÝ¶µÙ¶«ð´¶­ÝúÜ¶Ý­ÜÃÝàÝì·¡·ªÞ¬ÞÀ·¶ØÎ·¼·Æ·Ë·ÒÝ×Ü½ÜÀÜÞÜòÝ³ÝÊ¸ÊÜÕÞ»¸ê¸ï¸ðÝ¢¹¤¹¦¹¥¹§ÞÃ¹®¹¯¹²¹±÷¸¹¶¹½ÝÔðÙØÐºªÝÕÝïÞ¶ºÉºÊÞ¿Þ°Ý¦Þ®ºù»¨»®ÝÈ»Ä»ÆÜîÜöÞ¥»çØå»ò»ñ»óÞ½Ü¸»ùÝðÞªÜÁ¼»Ýç¼ÔÝÑÝó÷µ¼ë¼öÜü½¯½³Ü´Üú½¶½Ú½ä½æ½åÝÀÝ£½ùêî¾¥¾£Ý¼¾¯¾´ÜÚ¾Ï÷¶¾ÕÜì¾ÞÜÄÞ§¾úÝÜ¿±ê¬Ý¨¿Á¿ÖÜÒÞ¢¿àØá¿ï¿ûØÑÝÞÀ³À¶Ý¹ÝõÀÍÀÕÀÙÀòÝñÞ¼ÜÂÀóÝ°Á«ÝüÞ¤ÝþÜßÁâÜ×ÝäÂ«ÂÜÜýÂäÝ¤ò©÷´ÂûÃ¢Ã£Ã§Ã©ÜâÃ¯Ý®ÃÈÝùÞ«ÃÉÞÂÃçðÅÃêÃïÜåÜøÄ¡Ä¢ÜÔÄªÝëÄ³ÜÙÄ¼Ä¹Ä»Ä½ÄºÜµÝÁÄäØ¥ÜàÄõÞÁÅ·Å¹ê±Å¸ÅºÝâÝåÅîÜÅÆ¥ÜÖÆ»Æ¼ØÏÆÎÆÐÆÏÆÑÆßÝÂÆÚÆÛÜÎÆäÝ½ôëÞ­Ü»ÜùÝÝÝÖÜ·Ý¡ÜÍÜçÇ¾Üñ÷³ÇÉÇÌÇÊÇÐÇÑÜËÇÛÇÚÇæéÑÜÜõ¼öÆÚöÜäòËÇøÞ¡Þ¾ÜõÈ§ÈµÜÛÜéÈÇÜóÝØÈÖÈ×ÈÙÈØ÷·ÈãÞ¸ÝêÞ¨ÈïÜÇÈôÈøÉ¢É¯ÜÏÉ»ÉÖÜæÉõÝéÝªÊÀÊ½êÛÝÄÊßÊíÝôË¹Ý¿Þ´ËÕÝøËâÝ´Ý¥ËòÌ¦Þ·ÌÑß¯í«ÌÙÝÆÜðÝãÜíÝ±ÝËÜ¹Ý¸ÝÒÝÚÞ±Î­Î®ÎµÝîÞ³Ý«ÎÔÎ×ÎßÜÌÎôÝ¾ÎõêêÝßÝûÏ»Ý²ÞºÜÈÜ¼ÝÙÏïÏîÏôÐ°Ð¬Þ¯Ð¾Ý·Ð½ÜôÜºÐîÞ£ÝæÑ¥Ñ¦Þ¹Ü÷Þ¦Ñ»ÑÀÑ¿ÑÅåÂÝÎÄèÜ¾Û±Ñà÷±Ò¢Ò©Ò½ÜèÒÃÜÓß®ÒÕÞ²ôèÒðÒñÛ´ÜáÓ¢ÝºÜãÜþÓ«Ó¨Ó©ÓªÝÓÝöÝ¯ÝµÝ¬ÝÇÓóÝ÷ð°Ô·Ü¿ÔÌÔÑÔáÔåØÓÕºÕáÕåÝèÕôÖ¥ÜÆÝ§ÜïÜÑÖø×ÂÝÏÞ©';
	v_b       Varchar2(1200) := '°¢°¯ÚææßÚé±Ýåø³ÂØ©³Ðò¿³Ü³ö³ýÚï´Ïµ¢ñõÚçñô¶¸¶Ó¶é¶ú·ÀÚâ¸½Úë¸ô¹¢¹Âñøº¢º¯ÚòãÄØ½¼Ê½µ½×ÚàæÝÚáÚå¾ÛæÞ¿×ñùÁªÁÉÁÄÁËÁêñöÂ¡Â¤ÂªÂ½ÛÂÃÏÄ°ÚíÄôò¨ñ÷ÅãÚðÆ¸ÚäÛÉÈ¡È¢ÈæÈîÉÂËåËæËîËíËïÌÕôÐÍÓÚñÚóÎÀÚãÚôÏ¶ÏÕÏÞÏÝÚêÑ·ÑôÒ®Ò²ÒõÒþÓçÔºÔÉÕÏÕóÖ°Úìæï×¹×Î×ÓÚîÛ¸×èÚè';
	v_c       Varchar2(1200) := '°ÍÛÎÛÍæô²µ²Îæî²æ³Ò³ÛæõñååÊµ¡µËµþ¶Ôí¡æâ¹Ûº§æè»¶¼¦æ÷¼è½¾ñæ¾¢ØÙ¾±åÉ¾Ô¿¥æìÀÝæêæòÂ¿ÂâÂæÂíÃ¬òúó±î¦Ä²öÊÄÑÄÜæéÆ­æëÆïÛÏÇýÈ°åÒÈáÈþë§É£òªÉ§æóÊ¥Ê»Ë«ÛÌæáææÌ¨Û¢ìÆÍ¨ÍÔÍÕæÄæððÍÏ·æøæçÐÜÑ±ÑéÒÔÒÓæäð®ÓÂÓÁÓÖÓèÔ¦Ô¤Ô¥ðÖÔÊæàÔéæöÖè×¤æíæã';
	v_d       Varchar2(1200) := 'íÁ°­ðÆ°Ù°î°ï°õ±¯±®±²íÕ±¼±Ãí¾íÜ÷Ôì­÷Þíç²¼íå²Þ²êò²³§ñéíº³½í×³Éô©³ø´¡´º´½´À´Åõ¾´æ´èØÈÞÇ´ï´óí¸ìâíãíÆíÚµâµïµúíÖíÏíÔíâõ»í»¶á¶ò¶øð¹íÀ·¯·Çòãì³ôä·Ü·áí¿Ûº·îíÉÞÎÞÏ¸Ðí·íÑ¸û¹¨ð³¹¼¹ÅØÅ¹Ê¹Ë¹èíÞº»ºÄºñºúðÉð­÷ß»Ç»Òåç»Ûñëí¶ØÞ÷Ùí¢ê©¼ï¼îíäñð½¸íÙ¾¤÷Ý¾Â¾ÇØÊØã¿³¿Ä¿ËØÚ¿äÞÅ¿ó¿ø¿ü÷ÕíÇñìñçÀÚÀåÀúÀ÷ÀøíÂÀùÞÆðÓÞÍ÷àÁ×ÁòÁúíÃÁûÂ¢ñïÂµÂëÛ½ÂõíË÷ÖíæÃæñòÄÎÄÍíÐÄëññí¤ñî°ÒÅÕÞËÅáÅéÅðÅöÅøÆÆÆÝÆæÆõÆöíÓÇ£íÍêüÇØíá÷ÜÈ®È·÷×ÈèÈýíßÉ°ÉÝØÇÉéò×Ê¢Ê¯ÊÙÊùË£Ë¬Ë¶ØËñêËÁËéÌ«Ì¬Ì©Ì¼ñíÌ×÷ØÌüíÅíÈÍëÍòÍþíÒÎìÎùÎøÏ®íÌÏÄÏÃÏÌÏáÏõíÊÐÛ÷ÛÐçí¹íÛÑ¹í¼ÑÐÑÙØÉ÷ÊÑáÑâÑÞÑã÷ÐØÍÒ³ØÌì¥Ó²ÞÌÓÈÓÑÓÐÓÒÓôÔ­Ô¸ÔÅÔÒÔÚê°ØÆíÄÕÉíÝÕèíØ×©í½÷Úñè×××à×ó';
	v_e       Varchar2(1200) := '°®°·°¹°ò°û±ª±ìáÙë÷²±²²²É²Ê²ò³¦ëóëí´àëâµ¨õõëëëËëØ¶Çëñ·¾·Êëè·Î·ôæÚ·þÛ®¸¬¸­¸¹æØê®¸Î¸Ø¸ìëõëÅ¹ÉëûëÒë×ë½ëÜºÑ¼¡¼°ëÎëì½º½ÅëæëÂëÖöÂ¾ôëÌ¿èëÚÀ°Àßì¡Á³ì¢ëÊëÍëáÂöÃ²ëüëßÃÙëïåãÄ¤õöõøëÇÄËØ¾ëîÄÔÄåëåÅ§ÅÖëãÅßÅóÅôÅòÆ¢õùëÝÆêëÉÇ»ÇÒëÔëÀÈéëÃëÛÈùëýáêëþÉÅëÏÊ¤õ¹ÊÜëòË¦Ë´Ì¥ëÄÌÅÌÚëøÌóÍÈëàÍÑÍ×ëðëäÍóë¿ÞÉÏ¥ìÞÏØÏÙÐ²ÐÈÐØõ÷ÐëÐüëÙëçÑüÒ£ôíÒ¨ðÎÒ¸ÒÈÒÜÓ¯ëôÓ·ÓÃÓÐëéë¼ÔÂë¾ÔÐÔàëúÕÍëÓëÞÖ«ëÕÖ¬õôëùÖ×ÖâÖðÖúëÆëêëÑ';
	v_f       Varchar2(1200) := '°£ö°ÛûÛê°Ó°ÔÛà±¢êÚê´ØÃ²ª²©ð¾²º²Å²Ã³¡³¬³¯Ûå³Ã³ÇÛôÜ¯³à³áØ¡´£´çÜ¤´÷µÌêëÛæµØµßÛãÛìÜ¦ñó¶¯Ûí¶¼¶Â¶Ñ¶Õ¶â¶þ·»ö­·Ø·â·ò¸°Ûò¸ÉÛá¸ÏÛÙÜª¹¡¹¸ì°åÜêíì±¹Äî­ØÔ¹ç¹æÛö¹ýñüÚõº«º²º¾ºÂîÁºÕÜ©ºøì²»µ»Üí£»ê»ô»÷»øêå¼ªêªö«¼Î½Ì½Ùò¡½Ø½ø¾®¾³ôñ¾Èôò¾ùÛî¿¯¿°¿²¿¼¿À¿Ç¿ÓÜ¥¿å¿é¿îÛÛ¿÷åÓÀ¤À¬ÀÏÀ×Ü¨æËÛÞö¨Ü®ÛøôóÁØÁãÛ¹ÛâÂ¶ÛäÂñö²ÂôÜ¬ë£Ã¹ÛéÄÏôöÛñÛèÄÞéýÛþÅàö¬Ü¡ÅíÅ÷ÅùÛýÜ±ÛÜÆºÆÂÆÒØÁÛßêÈÆðÇ¬Ç½ö¥íàóÀÇóåÏôÃÇ÷ôðÈ¥È¤È´í¨ÈÀÈÍò¬É¥Ü£ØÄö®ÛïÉÊÛðÉâÉùÊ®ÛõÊ¿Ê¾ó§ËªËÂË÷ËúËþÌ®Ì³Ì¹ÌÁÌËèºÌæÌîöªÍÁÜ¢ÛçÍçÎ¤Û×Î¥Î´ö©ÎÓÛØÎÞÎëÎíìäÏ²Ï¼Ï×ö±ÏöÐ¢Ð­Ü°ÐÒÐèÐæÑ©Û÷ÛëÑÎÑßÒ¼ÛÝÛüÜ²Ü§Ûóö¯Ü­ö§ÓêÓòÔªÔ«Ô¬ö½Ü«Ô¶ÛùÔ½ÔÆÔËè¹ÔÕÔÔÔØôõÔöÕÔÕßô÷ÕæÛÚÕðÖ§Ö±ÛúÖ·Ö¾Öóôã×¨ôô×ß';
	v_g       Varchar2(1200) := 'è¨°½éáåÛ°¾ñúòü÷¡æñöË°à°ß±Â±Æ±Ì±íéë±ûÚû²£åÍ²»²Ð²Ïè²²Üè¡ôùë·´Ìè®çýéãè­´õ´ø´ùçéééµ½ôîçèµåçà¶¹¶º¶¾ó¼¶ÙÛÑ¶ñØ¬çí·©çã·ñôï·ó¸¦¸±Ø¤ØªØ¨¸üçî¹åôççñº÷çú»¥»­»¹»·è«ò³çõ»ÝØ¢çáÛÔ¼¬éê¼ÐçìÛ£¼Õê§¼ß¼íê¯ôøèª½úö¦¾²¾Áè¢çå¿ªçæçûØÝÀ´êãÀµÀÅð¿Á§åÎÀíÀôÀöÛªçöéçÁ½ÁÐÁÒÁÑÁÕÁáÁðççè´çóÂêÂóè£ÃµØÂÃðçäçëÄ©éâÄÒØ«è§ÅªÅÃØ§ÚüÅýéèÆ½çêè±ÆÞçùç÷ÇÙÇàÇíÇòè³ÈÚÈðØ¦ÉªÉºéäÊÂÊâÊøËÀËØËÙËöè©ìýÌìãÃéå÷ÑÍÌÍÍéÒÍßÍáÍãÍæçþÍõçâê¦ÎáÎåÎäðÄØ£è¦ÏÂÏÖÐÌÐÏÐÎÐÍçïè¯Ñ³çðÑÇÑÉÛ³ÑÏçüÑêçòÑþÒ»ÒÄéìçøè¬ÓØÓÚì£ÓÛè¤ÓëÓñè¥éæÔÙè¶ÔâÔæÔðÕµè°ÕäÕéÕûÕýÕþÖ³ÖÁÛ¤ÖÂÖé×¸×Á';
	v_h       Varchar2(1200) := 'öµ±ë²·²½²Ç²ÍôÓ²ñêèö³îªí÷³Ý³òöº´Æ´ËõºíñµÉíûµã¶¢¶¦¶½¶ÃíïØ­ÛÖò®»¢ØÀíú½Þ¾¦¾Éö´¾ßåáöÄÛÇ¿¨î«î§¿Ïíî¿ôî¥íùî¬ÁäÂ¬ðµÂ­Â±Â²ÂÇÂ÷ÃÐÃßííÃéíðî¨íøÄ¿ÄÀíþÅ°ÅÎÆ¤î¢î©Æµò­ÆÄê·Æçò¯ÇÆë¬È£êïî£ÉÏÊåË¯Ë²î¤íõî¡íüìþö¶Ì÷Í«Í¹ö»Ï¹Ð©íìÐéÑ£íýÑÛíôö¸ØÕÓÝö¹Õ£íÎÕ°Õ¼Õ½ÕêÕöÖ¹Öõ×¿×Àêßõþö·×Ïö¤íö';
	v_i       Varchar2(1200) := '°Äå±äºääå¨ãê±õ±ô²¨²´²³²»²×äî²âä¹ãâäý³¢³£³¨ë©³±³º³¾³Á³Î³Ø³ü´¾äÈ´ãí³µ­å£µ±µ³ÛÊµÎµÓµáµí¶´äÂ¶Éãçãõ¶ý·¨·ºäÇ·Ð·Úå¯ãã¸¡¸¢äææÙ¸Èãïä÷äÆ¸Û¹µ¹ÁãéäÊ¹à¹â¹öº£º­ººº¹å«ãìå©ºÆå°ºÓºÔãüºéºèÙäö×äïºþä°»¦»¬»´ä¡ä½»ÁäñäÒäê»Ôä§»ãä«»ë»ìäã»î¼¤¼³¼¹ä©¼Ãä¤¼âäÕ½§½¥½¦½­ä®½½äÐ½à½ò½þãþ¾Æ¾Ú¾Ùä¸¾õ¿£¿ÊäÛà·À£äµäþÀ½äíÀÄÀËÁÊÀÔãîÀáÀìå¢Á¤äàÁ°å¥äòÁºÁ»ÁÓä£ÁÜãöÁïä¯Á÷öÌãñÂ©ãòäËÂËäõÂºÂÐÂÙãøÂåäðÂúÂþäÝã÷Ã»äØä¼í¯ãèÃÚãæäÏíµÃìãýäéÄ­Ä®ãåÄ×ÄàÄçÄùÅ¢Å¨Å½ÅÉÅÈÅËãúäèÅÝÅæäÔÅìäÄÆ¯ÆÃÆÅå§ÆÖäßÆÙÆãÆâÆáä¿ãàÆûÆüÇ¢Ç±Ç³äÚÇßÇåÇöÇþÈ¸È¾ÈÜå¦Èêä²äáÈóÈ÷É¬É³ôÄöèäúÉÇÉÑÉÍÉÐÉÙäûÉæäÜÉîÉòäÉÉøäÅÊ¡ÊªÊçãðÊþäøäÌãßË®äùãáãôäÁäÑä³ËÝå¡æ¶êýäâÌ­Ì²Ì¶ÌÀÌÃÌÄäçÌÊÌÌÌÎÌÏä¬ÌÓÌÔØ»ÌéÌíÍ¡äüÍ¿ÍÄãûÍÝÍåÍôãíä¶Î«ä¢Î¼ÎÂãëÎÐÎÖä×ÎÛä´öÈÏ«ä»äÀÏªÏ´ÏÑÏæÏ÷ÏûåÐäìÏýÐ¡Ð¤Ð¹ÐºäÍå¬ÐËÐÚäåäªäÓäöãùäÖÑ§í´ä­ä±Ñ´ÑÄÑÍäÎÑØÑÝäÙãóÑóÑúÒ«ÒºäôÒÊÒçä¦ÒùäÞäëå­Ó¾Ó¿ÓÍÓÎÓÙÓåÓæÔ¡ÓþÔ¨ãäÔ´å®ÔèÔóÔüÕ´å¤Õ¿ÕÄÕÇÕÆÕÓÕ×Õãä¥Ö­ÖÎÖÍÖÞä¨äóä¾×¢äÃ×Çä·åª×Í×Ì×Ò×Õ';
	v_j       Varchar2(1200) := 'êÓ°µ°º°ö±©òùêÎó©²õó¸²ý³©êË³¿òÉó¤³×³æòÜòíµ©ØÖê­µçµûò½¶ê¶ô·äò¶òÝòðòóêºê½¸òò´ò¼òÁ¹Æ¹éêÐêÁòå¹ûòäòÀêÏºµòºê»ò«êÂòÂºçºû»Èó¨»ÎêÍ»×»Þó³â·ó¶ò±òÌ¼á¼à¼øòÔò»½ô¾§¾°¿Åòò¿õêÒòñÀ¥òÒê¹À¯ÀÀòëòÛÀïòÃó¹ÁÀÁ¿ÁÙòÈò÷ÂÝó¡ÂìòýÂüòþòÖêÄÃ°ÃÁòµÃËòìó·ÃáóºÃ÷êÔÃøòïêÙòÍêÇÅ¯ó´ó¦ó²ò·òçóªÆØòÓòàòÞòûòßÇçòÇò°òøÇùòÐó½òéòÅÈÕòîÈäò¸É¹óµÉÎòÙÉßÉêÉöêÉÊ¦Ê±ÊÇÊîÊïÊúÊûË§ó°òÏòôåÝê¼ó¥ó«ÌÞÌâòèòÑÍÉêÕÍÜòêÍíÍúè¸ÎÃÎÏòÚÎîÎúòáó£ó¬êØÏºÏ¾åßÏÍÏÔò¹ó­ÏþÐªÐ«ê¿ÐÇÛÃìãêÑêÖÑÁÑÑêÌòÕê×Ò°Ò·êÊÒÏÒ×òæò¾Ó¬Ó°Ó³Ó¼òÄòöòÊØ®ÓÞòõêÅÓöòâó¢Ô»êÀÔÎÔçê¾òÆó¯ÕÑÕÕÖ©òÎÖëÖû×î×ò';
	v_k       Varchar2(1200) := 'ß¹°¡àÄ°¥°¦àÈàÉà»°È°É°ÏßÂõÀàÔ±ÄßÁ±ÉßÙõÏ±ðà£õÛõËß²²¸àê²ÈàÐàá²äõð³ª³°³³àÁ³Ê³Ñ³ÔßêàÍõØß³³ì³ùõéà¨àÜõß´¨´­´®´µõÖßÚõí´ÚßýõãßÕàª´ôß¾à¢ß¶µ¸àâµÅàÖàÇõÚµðµõµøà©õÞ¶£à¤ßËà½¶Ö¶×ßÍ¶ßõâßá¶åßÀ¶õðÊò¦õì·È·Í·Ôßôß»õÃõÆß¼¸À¸Â¸ÁàÃ¸úßçßì¹¾ßÉ¹ÐßÛ¹ó¹òßÃ¹þàËº°àãàÆº¿ºÅºÇºÈàÀºÙºßºåºíºðºôßü»£»©õ×»½»¼ßÔßÜà¹àëß´ßÒßóõÒê«ßâõÕõÊõÂ¼ùõÝõÓ½ÐàÝà®àµàäà±¾×õá¾à¾á¾éàÙàåõê½ÀßÇ¿§¿¦¿Èà¾¿Ð¿Ô¿Úßµ¿Þ¿çßàßÑà­õÍà°À²À®à¥ßëß·àÏà¬Á¨ß¿à¦õÈõÔàÚßÖõñßøõïÁíßÊ¿©Áüà¶ààÂ·ðØÂÀÂîßéÂðÂïßäà×ö¼ß÷ßãÃùßèÄÄà«àìßÎÄÅÄØàÅÄöà¿õæßÌßæßöàÞÅ¶Å»Å¿Å¾ßßõçÅØÅÜÅÞÅçàØàèÆ¡õäàÑÆ·àÛõëàÒõèÆ÷Ç²ÇºõÄõÎàºàßßÄÈÂõåàéàçÉ¤ßþÉ¶õÇÉÚÉëßÓÊ·ÊÈÊÉà§Ë±Ë³àÊËÔßÐË»ËÃà²àÕà¼ËäËôàÂßïõÁÌ¤Ì£õÌÌ¾÷ÒßûÌßÌäÌãÌçÌøÌýàÌÍÂõÉö¾ÍÙÍÛÎ¨Î¶Î¹ÎÇÎËà¸ÎØÎâßíÎüßñÎûßÈÏÅõÑõÐÏìßØÏùÏøÐ¥õóÐÖßÝÐêÐáÐúÑ«Ñ½ÑÆÑÊÑäßºÒ§Ò­Ò¶ßÞàæß×ÒÅß½ÒØà³Ò÷ßÅàÓÛ«Ó´à¡à¯Ó½Ó»ßÏÓõÓ÷Ô±Ô¾ÔÇßÆÔÛßðÔëÔêßõß¸Õ¦ßîÔûßåßúÖ¨õÅõÜÖ»ÖºõÙÖÐÖÒÖÑõàÖäõîÖößù×Ä×Ù×ãõò×ìßò';
	v_l       Varchar2(1200) := '÷ö°Õ°ìî¯±ß÷õ³µ³ë÷íê¡ê£÷ò¶Úéî¶÷·£î²î··ø¸¨î¸ÛÁéïî¹¹Ì¹ì¹õ¹úºØºÚºäéõàñ»Ø»ûî¿¼­¼ÓåÈôÂ¼×¼Ý¼Üàî½Î½Ï½çéðÀ§î½ÀÛî¾Á¦éöîºÁ¬Á¾ê¥àòéñéûê¤ÂÔàðÂÖÂÞÂßÃóÄ«Ä¬àïÄÐæÕÅÏÅþî¼ÆÔÆèÇ­ÇµèýÇáàõ÷ôÇô÷ñÈ¦éúî°éíÈíéøÊäÊðÊñË¼ËÄÌïî±î®Í¼ÍÅî¶îµéþàíÎ§Î·Î¸Ï½÷ïÛÄÐùÑ¼î»é÷÷ðéóÒò÷îàóàôàöÔ°Ô²Ô¯à÷ÔÝöÉîÀÔþÕ¶Õ·ÕÖéüÕÞéôî³éòéùÖÃÖá×ªê¢×ï';
	v_m       Varchar2(1200) := '°¶°»°¼á±°Ü±´±À÷Â±áì©ì®÷Æ²Æ²áá¯³çàüå×´Í´ÞáÏµ¤áØµÏ÷¾áÛµäµñá´á¼÷Ç¶Ä¶ä¶ç¶ë·«á¦·²···ç·å·ï·ùá¥¸³êçêà¸Ô¸Õ¸Ú÷Àá¸¹º÷½¹ÇáÄØÛáÆàþº¡÷¿á²»Ë»Ï»ßá§¼¸áÕêéáµ¼û¼úá½÷º½íêáØçåÄì«áÈ¾þ¿­ØÜ÷Áá³áÇ÷¼÷ÅêÜ¿ùáÁá°áÀá×Áë÷ÃáÐÂ¸á£á¹Ã±áÒáºÄÚÅÁÅâàúáªÆéá¨Æñ÷Äá©Ç¶ÇÍêäÇúá«È½áÉÈâÉ½É¾ÉÄÉÞáÓì¯ÊêÙîáÂáÔì¬ÙíËèËêÌ¿êæÌùÌûÍ¬Í®÷»áËÍøØèÎ¡àøá¡áÍá¢Ï¿á­áÅá¶á¾á¬ÑÂáÃÑÒÑëÑìáÊêÝáÚÒÙá»Ó¤ó¿ðÐÓÉÓÊáÎÓìÓøáÌÔßÔòàýÔôÔùÕ¸ÕÊÕËáÖá¤êâá¿Ö¡àùÖÄÖÅÖÜëÐÖüò§×¬´±áÑ';
	v_n       Varchar2(1200) := '°Ãã£öÍ±Øã¹±Ù±ÚæÔ±Ü±ÛèµôÅâíë¢éÞ²À²Ñ²Òâü²ãåîâãåñã®âêâ÷³À³Ù³ßâçã¿ã°³óâðãÀâë´Áã²´äâââòµ¬µ°âáµ¼µ¿µëµîµó¶®¶²¶èåíãµ·¢·Éã­·ßâö¸Ä¸Ò¹Ö¹ßº©º·º¶ºÞºãã±âï»³»Å»Ì»Ð»Ö»Úåì¼º¼É¼Â½ì¾¡¾ªã½¾Ó¾Ö¾ç¾ååðâý¿®âé¿¶åêã¡¿ìã¦ã´À¢ã§ÀÁã¶Á¯ãÁðÒÂ¾ÂÅÂÄÂòÂýÃ¦Ã¼ðÌãÂØ¿ÃñÃõíªÄÕÄáâõÄòâîÅ³âæÅÂãÝâñÅüÆ¨ê¶Æ©ÆÁÇ¡ã¥ã»ãÞÇÄã¾ã¸ÇÓã«ÇéÇüãªÉåÉ÷Ê¬Ê­ÊºÊÑÊÕÊéÊèË¢Ë¾ËÈâìã¤ãºÌëã©ÌèÌñâúÍÀÍÎÍïã¯Î©Î²Î¾Î¿ÎÝâèâäÎòÏ§Ï¬Ï°åïåÚÐ¼Ð¸âàÐÄÐÃÐÊÐÔã¬ñãÐôâþÑ¸ÙãñââûâóâùÒÒÒÑÒÔÒäÒìâøôàã¨ÒîÒíã¼ÓÇÓäÓðÔÃã¢ã³ìÙÞÊÔ÷µÔÕ¹ÕúåëâåÖçÊôã·âô';
	v_o       Varchar2(1200) := 'ôÎ±¬±ºìÔìá±þ²Ó²ÚìÌ³´³ã´¶ôÙ´Ö´âµÆ¶ÏìÑìÀ·³ìÜ·Û·à·éíêíë¸âìÊº¸ºæô×ìÃìÎºý»À»Í»â»ðôÝ½ý¾¬¾«¾¼¾æìß¿·¿»¿¾ÀÃÀÓÀàÁ£ôÏÁ¶Á¸ÁÇÁÏôÔåàìÖÂ¦Â¯ì×ÃºìËÃÔÃ×ôÍÅ´ÅÚÆÉìÁôÜÈ²È¼ÈÛôÛôÖÉ¿ØßìØÉÕìÂÊýË¸ìÝÛ°ÌÇÌþìÕÍéìÐÎªì¿ìÉÏ©ôÑÏ¨ôÌÛÆôÚìÓìÅÑÌÑ×ÑæìÍì¾ìÈÒµÚþìÇìÚìÏìÛôØÔãÔäÔîÔïÕ¨Õ³íéÖòìÄ×Æ×ÑôÕ';
	v_p       Varchar2(1200) := '°²°¸°À±¦ñÙ±»ñØñÔñ¹ñÛñÑ±ö²¹²ìñÃìøå·³ÄñÎñÝ³è³õñÒ´©´°åÁìô´Üñ×ñÉå´µ»¶¨ñ¼ñÖ¶îìð¸¤¸£¸»ñË¸î¹¬¹Ñ¹Ó¹Ù¹Úå³º¦º®º±ºÖº×ºêìïå¾»ÂåÕ»í»ö¼Å¼Ä¼ÒñÊñÐåÀå¿½ÑñÆ½ó¾½¾¿ñÕñÀ¾üñäñ½¿ÍñÌ¿Õ¿Ü¿ß¿ã¿í¿úñÜÀÎÀñÁ±ñÍñÏÁÈå¼ÁþÂ»ñÚÂãñÇÃÂìòÚ¢åµÃÜÃÝÃÛå²Ú¤Ä¯ñÄÄþå¸Å©ñÈñáÅÛÆîÆíì÷å¹åºå½ñßÇÏÇÔÇÞÇîñ·ìîÈ¹ìüñÅÈÝÈßñàÈìÈûÈüÉÀÉçÉñÉóÊµìêÊÓÊÒÊØìëËÎËÞåäËüÌ»ñÓìöñ»Í»ÍÊÍàØàÍêÍðÎÑå»ñ¶ìûìùììÏÜÏéÏüÐ´ÐäÐûÑ¨ñ¿ÑçÒ¤ñºñÂÒËÒúÛÈå¶ñ¾ÓîñÁÔ¢Ô£Ô©Û©ÔÖÔ×Õ¬Õ­Õ¯ñÞÕâìõð²Ö®ìóìíÖÏÚ£Öæ×£ñ¸ìú×Ö×Ú×æìñ';
	v_q       Varchar2(1200) := 'ï¹ïÍï§áíîÙöÑîÓ°÷Ùè°ü±¥±«±·±µï¼áùîéöýïÚïð÷§ïÙ±ý²§âÄîà²¬îßîÐ²ÂâÇïÊâªïïîÎ²öïâ²ù²þöðÛË³®îõîñð·âÁï¥Û»³ú³û´¥â¶îË­[´¸´Òâ§ïéï±´íµºï½ïëïáµÒØµÛ¡µéîäöôµöîööø¶¤¶§îû¶µî×¶À¶Æ¶Íïæ¶Û¶àîìï°¶öïÉöù¶ùöÜ¶ûåÇ¶üîï·°·¸·¹îÕöÐöîáôïÐö÷·æ·õÙìÙë¸ºöÖöûîÅ¸Æ¸Ö¸äï¯ïÓ¸õöáö¡¹´¹³¹·¹»õýîÜïÀöñ÷¤¹ÝîÂáî¹ä¹êöÙ÷¬öç¹øâ£îþºÝÙêºïºöºüõúâ©ð×îü»«âµÛ¨ïÌÛ¼öéöü»èâÆïÁîØïì¼¢¼±öÝöêïØîò¼Øöäïµ½¤¼üöÞ½Ç½Æ½È½ÂèîöÚ½âîÄ½ð½õâË¾¨â°¾µ¾Ã¾Ä¾Ñï¸¾äîÒ¾âïÔïÃáúõûâ±ïã¾ûï´îøïÇîÖîíîÝï¾ï¬áö¿ñÀ¡ï¿öïïªïçÀÇï¶ï©îîÀÖ÷¦ÀØÀêöâï®Àð÷¯áûöãÁ­Á´â²ÁÍîÉÁÔÁÛÁåöìÁôåÞÁóïÖï³ïÎöÔÂ³ïåÂÁÂÑï²â¤ÂàïÝáïÂø÷©ïÜÃ¨ÃªÃ®Ã­Ã³â­ïÑÃ¾îÍÃÍÃÌâ¨ÃâÃãÃûÃúâÉïÒîâïÕÄÆâÎîóâ®ÄÙîêâ¥öòöÓöóÄñôÁÄ÷ÄøÄüáðÅ¥îÏï»ÅÙáóïÂîëöÒîÇîÞÆÌÙéïäïè÷¢Ç¥Ç¦îÔÇ®Ç¯Ç·ïºïÏïêÇÂïÆÇÕï·ÇäöëöúáìÛ¾ð¶îýáëÈ»ÈÄâ¿áõï¨ÈñÈúâÌöþÉ«ï¤É±É²ï¡É·îÌ÷­õüÉ×â¦Ê¨öõâ»Ê´öåÊÏÊÎîæß±á÷îåïÈËÇâÈïËöÕö¢â¡áøâ¸ËøîèÌ¡÷£öØîÑïÄîãï¦ïÛâ¼ï«ÌàåÑöæï¢ÌúîúÍ­îÊÍÃâ½ÍÒÍâÎ£â«öÛâ¬ØØÎÒÎÚÚùÎÙÎðÏ¦Ï£Û­ì¤ÎýçôÏ³â¾Û§áòÏÁÏÇÏÊáýÏÚÏââÃÏóèÉÏúâ³åâÐ·Ð¿öÎÐÉÐ×ÐÙâÊÐâîçïà÷¨â´Ñ®öàØ³ëÈ÷¥îôâ¢Ò¿âÂîÆØ×ÒÝï×ïîî÷áþÒøâ¹ÒûÓ¡Ó­ïÞ÷«ÓÌÓËöÏîðÓãáüâÅâÀÓüîÚíóÔ§Ô³Ô¹Ô¿îáÔÈáñï­Õ¡Õ²â¯îÈÕàÕëÕòÕùÕøîÛï£ÖËö£ÖÓïñÖåÖíîùéÍÖýâÍ×¶ïíïÅöö×Þöíïß×ê÷®';
	v_r       Varchar2(1200) := '°¤Þß°¨°±Þî°´°¿ÞÖ°Ç°Æ°Î÷É°Ñêþ°×Þã°Ú°Ý°â°á°ç°è±¨±§±°ðÇÞÕ±÷±øÞð²¦²¥²¯²«²¶²Á²Ù²å²ë²ð²ô²ó³­³¶³¸³·ÞÓ³Å÷Î³Ö³â³é´¤Þõ´§ë°´·ß¥´Ý´ê´é´ì´ë´î´òµ£µ§µªµ²ë®µ·µÄµÖµàµæµôµüð¬Þéë±¶¶¶Ü¶Ý¶Þ¶óÞô·´·µ·Õó¾·ö·÷·ú¸§ÞÔß¦¸×¸Þ¸ã¸éë¡¹°¹Ï¹Ò¹ÕÞè¹Þð§¹íÞâº¤º´Þþº³ð©ºóåË»¤»»ß§»ÊåØ»Óß«¼·Þá¼¼Þö¼ð¼ñÞØð¨½Áë¸½Ó½ÒÞ×½Ý½ï½ü¾¾¾ÐÞä¾Ü¾Ý¾è¾ï¾ñ¾ò¾ðÞÜ¿«¿´¿¸¿¹¿½ë´¿Ø¿Ù¿Û¿æÞñ¿ýÀ¦À©À¨À­À¹À¿ÀÌÀÞ÷ËÁÃÁÌÞæÁàÂ£Â§ß£Â°ÂÈÂÓÂÕÞÛÞû÷ÈÞÑÃèÃòÃþÄ¨Ä´ÞàÄÊß­ÄÓÄâÄéÄêÄíÄìÄóÅ¡Å£Å¤Å²ÞùÅÀÅÄÅÅÅÒÅ×ÅêÅõÅúÅûÛ¯ß¨ë­Æ²Æ´Æ¹ð«ÆÈÆÇÞåÆËñýÆøÆþÇ¤ÞçÇÀÇËÇÜÞìÇâÇèÇðÇñÈªÈ±ÈÁÈÅÈÈÈÓÈàÈöÞúÉ¦É¨ÉÃÉÓÉãÊ§Ê°ÊÆÊÃÊÅÊÄÊÖÞÐÊÚÊãÞóË¤Ë©Þ÷ËºËÑËÓËðËùÌ¢Ì§Ì¯Ì½ÌÂÌÍÌáÞÝÌôÍ¦Í±Í¶ÞÒÍÆÍÐÍÏÍØÍÚÍìÍî÷ÍÎÎÎÕÎæÎèóÁë¯ÏÆ÷ÌÐ®Ð¯ß¢Ð¶ÐÀß©Þïõ½Ñºë²ÞëÑÚÑïÑõÑûÒ¡ÞÞÒ´Ò¾ÒÖÞÚÒóë³ÞüÓµÞíÔ®ÞòÔÀëµÞÙÔÜÔíÔñÞêÔúÕªÞøÕÐÕÒòØÕÛÕÜÕÝß¡ÕñÕõÕüØ´Ö´ôêÞýÖ¸ÖÆÖÊÖ¿êÞÖÀðºÖìÛ¥Öô×¥×¦×§×«×²×¾×½ßª×áß¬ß¤';
	v_s       Varchar2(1200) := 'èñ°Ø°å°ð°ñ°ô±­±¾è¼±ê±òéÄ±úõ³²Ä²Ûè¾²éé¶éß³»é´èßèÇõ¨³È³êéË³÷èÆèú³þ´ªé¢é³´»´¼èÈé¨õ¡´×éÁ´åµµèÜé¦¶¡ôú¶¥¶°èü¶Åé²·®èóèÊé¼·Óèû·Ù·ãèõ¸²¸Å¸Ë¸ÌôûéÏ¸ÜéÀéÂ¸ç¸è¸ñ¸ù¹£èÛ¹¹ôþèô¹×èæ¹ñ¹ð¹÷é¤õ°º¨º¼ºËèìºáéÎõ­èë»±»¸èí»ú¼«é®¼Ï¼ÖèÅ¼ìé¥¼÷½·½Íõ´½ÜéÈ½ûèÑèêé§½ÛéÙé·õ¶Û²èöéÓ¿¬èà¿Â¿Ã¿É¿Ý¿á¿òõ«èéÀ¸À·é­ÀÆõ²èáÀÒéÛõªÀâÀãÀîõ·èÀèÝÀõé¬é£ÁÖéÝèÚèùÁñÁøèÐÂ¥èÓéÖÂ´éµé¡è¿í®Ã¶Ã·é¹Ã¸ÃÊÃÎÃÑÃÞèÂõ¤Ä£Ä¾èÍéªÄðÄûèËÅÊõ¬ÅäÅïèÁØâÆ®Æ°Æ±é¯èÒÆÓÆÜèçéÊÆåè½Ç¹éÉÇÁÇÅéÔéÕé±È¨È©È¶èãéÅèÄÉ­É¼ÉÒé©õ§ÊÁÊàÊáÊõÊöÊ÷Ë¨ËÉËÖËÚËáé¾èøËóé½ÌªñûÌ´éÌõ±ÌÒÌÝõ®èèÍ©ÍªÍ°õ©èÞõ¢ÍÖèØÍ÷Î¦Îàè»Î÷ÎöðªéØõµÏ­èÔõ£ÏàÏëÏðèÕÐ£Ð¨ÐµéÇé¿ÐÑÐÓÐàèòõ¯Ðïé¸õ¸èâéÜõ¦ÑîÑùèÃÒªÒ¬ôýÒÎÓ£éºèÖÓÏÓÜéÚéÐÔÍé«ÔýÕ¤Õ¥Õ»ÕÁÕÈèþèÏèåÕçé»ÕíÖ¦èÙÖ²è×õ¥èÎèäôüÖêéÆèÌÖù×®×µ×Ãè÷×Ø×íé××õ';
	v_t       Varchar2(1200) := '°«á®°Â°Ê°Þ°ã°æô²±¸±¹±¿±Ç±Ëïõ±Êô°±ÒóÙóëó÷óÖ±ü²°²­ô¤²¾²Õô½²ß³¤áä³¹³Æ³Ë³Í³Ì³Óó×óøáÜô¾³î³í³ï³ô´¦ô­´¬´Ñ´¹´ÇáÞ´ØÙà´Ûë¥ïóóÎ´ð´ýóìµ¾µÃµÂô£µÈÙáµÐµÑµÚô¡ëº¶ª¶¬óûë¹¶¿óÆ¶Ìóý¶ã¶æ¶ì·¦·¤·¬·­·±áÝô³óõ·êïû·û¸´ð¥¸Í¸ÑóàØº¸Ý¸å¸æÛ¬ô´¸÷¹ªóôóÑ¹¿êôðÀêö¹Îð»¹Ô¹Üóþ÷ýº½ºÌºÍºÜºâóóááºõóË»²áåóò»É»Õ»à»ýóÇïúê÷»þ»üóÅ¼®¼¾ð¢óÕ¼Ú¼ãêùóÈ¼ò½¢ë¦¼ý½Ãáè½Õ½Ö½î¾¶¾Ì¾Øé°êøóÞêû¿¿¿ÆïýóíóØ¿ê¿ðóñáâô¥ÀºïüÀæÀçÀèÀé÷óÀûóÒóöÁýÂ¨ôµóüëªïùÂÉÂÒÂáöÃÃ«êóÃ´Ã¿ô¿ô»ÃØÃëóúÃô÷ªï÷ë¤ÄµÄÁÄÂð¤ô«ô¬óáÅÇÅÆÅÌÅÍåÌÅñÆ¬êúÆªØ¯êòÛ¶óÍë¶ë«ÆòÆùí¬Ç§Ç¨í©Ç©óéÇÇóæóäóÌÇï÷üÇûáéóÜð¦ÈÉïþÈëóèð£ì¦É¸ô®ÉÔóâô¹ÉàÉäÉíïòÉýÉúÉüóÏÉûíòÊ£Ê¸ÊÊóÂÊÍóßô¼ïøÊòË°Ë½óÓËÒóùËãËëËñÌºÌÉÌØÌðÌòÌõóÔÍ¢Í§Í²Í¸ÍºÍ½ÍÇØ±óêÍùåÔÎ¢Î¯ôºÎºÎÈÎÒÎçåÃêõÎñÎïÏ¢ÎþÏ¤Ï¡ô¸áãÏµÏÈÏÏÏÎóÚÏãÏäÏòóïóãÐ¦Ø¶ÐÆÐÐÐãÐìÑ¡ÑªÑ¬Ñ­áßÑÃÑÓóÛÑÜÑíáàØ²áæåÆÒÆô¯ÒÛØæØ·ë»ÓÔóÄÓíÓùØ¹óîÔÁô¢êÃÔÞÔìåÅóÐô·óåÔõ÷þÕ§Õ±óÉóðð¡Õ÷óÝáçÖªâºÖÈÖÇÖÉïôô±ó®ÖÖÖØÖÛô¦ÖñóÃô¶Öþóç×­ïöóÊ×Ô×â×ë';
	v_u       Varchar2(1200) := '°©°Ìñ£ñ­°ë°ê±±Úý±³±Å±Õ±Ö±Ô±×±æ±ç±èñ¦±ï±îõ¿±ñÙû±ù²¢²¡²¿ê³²îðû²ú²ûÙæãÑ´³³Õà´ñ¡³åñ¬´¯´Ã´É´ÈðË´Î´Õ´áðî´ñµ¥µ¦ð÷ðãµÁµÀôÆµÜµÛµÝñ²µìñ°µòðÛ¶³¶·¶»¶Ë¶ÒãÕ·§ðò·è·ë¸Çðá¸Óí°¸á¸í¸ó¸þðó¹Ø¹ëÙòÙåºÒãØºÛãÈðúÙü»¾»¿ñ¥ãÔ¼²ñ¤¼½ðèðý¼ä¼æ¼åðÏ¼õ¼ôôå½ª½«½¬½±½°½´½»½¼ðÜ½ßôÉ½ê¾»¾·¾º¾¹¾¸ãÎ¾Î¾ÒîÃ¾íèð¾ì¾ö¿¢ãÛãÊðâ¿öÙçãÍÀ«ðøÀ±ñ®À¼À»ãÏðìÀäÁ¢ðÝðßÁ¡Á¹ÁÆÙýÁÝÁèÁçÁöÁùñªðüãÌñ§ÃÀÃÅÃÆãÉÃöñ¢ÄÖðÚÄæÄýÅ±ÅÐÅÑÅÔðåÆ£Æ¦ñ±Æ³Æ¿ÆÊÆÕÆàÇ°Ç¸Ç¼ôÇæªÇ×ÇõåÙñ³ãÖÈ­È¬È¯È³ã×ãÚÈòìªðþððÉÁðÞÉÆÛ·ÉÌÉØÊ×ÊÞÊÝãÅË·éÃÚ¡ñµËÍËÜËìôÈãËÌ±ÌµôÊÌÛÌÜðÃÌêãÙÍ¯Í´Í·ãÇðôÎÁÎÅãÓÎÊðíôËãÒÏÐðïðÂÙþÏÛÏèößÐ§ÐÁÐÂì§ÐßðçÑ¢Ñ¾ðéÑËãÆÑÖÑÕÙðÑåÑòÑñÑøÑ÷í¦Ò±ðêÒßÒæñ´ÒâðùÒãñ¯Òôñ«ñ¨Ó¸ðàéàðöðõãÐÔÄÔÏÔøêµÕ¢ðäñ©Õ¾ÕÂÛµÕÃÕÎÕîÖ£Ö¢ÖÌðëðñðæ×±×°ãÜ×³×´×¼×Å×È×É×Ë×ÊæÜôÒí§×Ü×ð×ñ';
	v_v       Varchar2(1200) := 'æÈæÁåþæ¾æ»°þæ±æ¿æ½æÏ³²æÊçÝæ§´þµ¶µÕæ·¶Ê¶ðæ¹·¢·Áåú÷÷¸¾ê¸ôÞæÅ¹Ãæ£ºÃ»Ù»é¼§¼´¼µåæ¼Ë¼Èôß¼Þ¼é½¨æ¯½¿½Ëæ¼Ëý½ãæ¡æºð¯¾Å¾Ê¾Ë¾ê¾ý¿¤¿Ñ¿ÒØ¸ååÀÑæÐæ²Á¥ÁéÂ¼åÖÂèæÖÃ½ÃÃÃÄÃäÃîæÆÄ·ÄÇÄÈÄÌÄÛÄÝÄïí±æ¤Å«æÛæåÅ¬åóæÀÅ­Å®æÇæÎæ°æÉæ³æÍÈºæ¬ÈÌÈÐÈÑÈçÉ©æ©æÓÛ¿ÉÛÉïÉôÊ¼æ­ÊóË¡æ×åùæ¦ÛÅËààûÌöæÃÍËÍÞæ´ÍèÍñæ¸÷ùåüæÒ÷ûÏ±ôªæµÏÓÐÕÐñÐöÐõÑ°Ñ²æ«æÌåû÷úÑýÒ¦ÒÌÒöÒüçß÷øæ¥ô§ô¨Óéí²åýæÂÔÓçÞæÑÕÙæ¨Öãæ¢';
	v_w       Varchar2(1200) := '°³°Á°Ë°Ö°Û°ä°é°øìÒ±£±¤±¶ÛÐÙÂ±ãÙÏ²®²ÖØ÷²à²íÙ­Øö³¥³«³ÞÙÑ³ðÙ±öÅ´¢´«´´´Ó´Ô´Ùìà´ß´ö´úá·´û´ü÷ìÙÙµ«µ¹µÇµÊµÍµèµùØê¶±¶Î¶íÙ¦·¥ÛÒ·Â·Ö·Ý·ÞÙÇÙº·ð·ü·ý¸«¸©¸ª¸¸¸¶¸·¸µÙ¤¸ë¸ö¹«¹©Øþ¹À¹ÈÙÄ¹ï¹ôº¬ò¥ºÏºÎºÐò¢ºîºò»ª»¯»á»ï»õÙ¥¼¯¼¿ÙÊ¼À¼Ñ¼Û¼Ù¼ó¼þ½£êð½¡ÙÔ½©½¹ÙÕðÔÙ®½Ä½é½è½ñ½öÙÓÙÖ¾ãÙÆ¾ëöÁ¾ó¿¡ØûíèÙ©ØøÙÅÙ¨¿ëÛ¦¿þÀÐØìÀÜÙµÀýÀþÙ³ÀüÁ©Á²ÁÅÁÚÁÞÁæê²ôáÁìÁîÙÍÂÂÂØÂ×ÙÀÃÇÃüÙ°ØïÄÃÄßÙ£ÄãÄîÄúØúÙ¯ÙÐÅ¼Ù½ÅåÅèØòÆ§Æ«Æ¶Ù·Æ¾ÆÍÆóÇªÙÝÙ»ê¨ØäÇÈÇÎÇÖôÀÇÝÇãÙ´È«ÈËØéÈÊØðÈÎí¥ÈÔÈåÙ¼ØíÉ¡É®ÉµÉËî´ÙÜÉáÉìÊ²Ê³Ê¹ÊËÊÌÊÛÙ¿Êæë¨ËÅËÆÙ¹ËËËÊËÌË×ËÛöÀËûÌ°ÌÈÙÎÌåÙÃÙ¬Í£ÙÚÙ¡Ù×ÍµÍ¾ÙÛÙ¢ØôÙËÎ°Î±Î»è·ÎÌÎÍÙÁØõÎéÎêÙâôâÙÒì¨ÏÀÏÉÏñÙÉÐ±ÐÅÐÝÐÞð¼ÐðÙØØóÙ²ÙÈÑðÑöÒ¯ÒÁÒÀÒÇÒÐÒÚØîØýÙ«Ó¶Ù¸ÓÅØüÓÆÓÓÙ§ÓàÓáÓâêìÓßØñÙ¶ÓûðÁÓúÙßÕ®ØëÕÌÕìÖ¶ÖµÖÙÖÚÙªØù×¡ö¿×·Ù¾×ÐÙÌÙÞ×ô×÷×ø×ö';
	v_x       Varchar2(1200) := '°í°óð±±ÁØ°±È±Ï±Ñ±Ðåö±àçÂçÍ²ø³Ú³ñç©´¿´Âçªµ¯µÞ¶Ðç¶·Äç³·Ñ·×·ì¸¥ç¦ç¨åõ¸¿ç¤¸ÙçÉæü¸øç®¹­çÃ¹áçµç¬ºëºì»¡çÙ»º»Ã»æçÀ¼¨¼©çÜ¼¶¼Í¼Ì¼êçÌçÖ½®ç­êñ½Ê½É½Ô½áçÆ¾­åò¾À¾î¾øç¼ç«æþÀÂçÐçÊó»Á·çÔç±ç¸ÂÆÂÌÂÚÂççÏÃÖåôôéÃàÃåç¿çÅçÑÄ¸ÄÉÅ¦àÎç¢çÎç²ç×Ç¿çØÇêç¹ÈÆÈÒÈÞçÈÈõçÒÉ´ÉÉç´ÉÜÉðÉþç·ç£æùË¿ð¸çÁËçËõÌÐç°ç¾Í³åèæýçºÎ¬Î³ÎÆÎãÏ¸ÏËÏÒÏßÏçç½÷Ïç¯çÓç¥ÐåÐ÷ÐøÑ¤çÛÒÉÒÍÒïçËÒÞÒýÓ§ò£Ó±ÓÄÓ×æúå÷ÔµÔ¼ç¡çÕÕÀÕÅçÇÖ¯Ö¼Ö½åéÖÕÖàæûç§×ºçÄç»×Û×Ý×éçÚ';
	v_y       Varchar2(1200) := '°§âÖÚÏâÚ÷é°ù°ý±ÓâØ±â±å±ä±é±óÙ÷Ùñ²ï²÷âÜÚÆ²üêÆÚÈÚß³Ï³ä´²ðÈ´Êµ®ÚÔÚ®µ×ÚÐµêµ÷µý¶©¶Á¶Èí­¶Ø¶ïÚÌ·½Úú·¿·Ã·Åìé·Ì·Ï·í¸®¸¯¸¼¸Ã¸ß¸àÚ¾¸ýâÙÚ¸Ú¬¹ÍÚ´¹ãâÑ¹î¹ù¹üº¥ñþºÁºÀÚ­ÛÀºàÚ§»§ìæìè»°ëÁ»ÑÚ¶÷â»ä»åÚ»¼¥¼£ì´÷ä¼Æ¼Ç¼Á¼çÚÙÚÉ½²Ú¦Úµ½ë½÷âÛ¾©ìºìç¾ÍðÕÚª¾÷ÚÜ÷å¿µ¿ºò¤¿Ì¿Î¿âÚ²Ú¿Ú÷ÀªÀ¾ìµÀÉÀÈÀÊÙúÚ³ÀëìåÁ®æ®ÁµÁ¼ÁÁÁÂÁÎ÷ëâÞÁßÁõì¼Â®Â¹ÂÃëöÂÊÂÏÂÍÂÎèïð½ÙõöÇÂÛÙùÂéÂùÃ¡ÚøÃ¤ì¸ÙóÃ¥ÃÕÃÓ÷ã÷çÃÒÚ×ÃíÃýÚÓ÷áÄ¦Ä¥Ä§Ä±Ä¶Ú«ì»ÅµÚ©ÅÓâÒì·ÅëâÏôæÚÒÆÀÙöÆ×ÆëÆì÷èÆôôìÆýÆúÇ«Ç´ÇÃÚÛÚ½ÇëÇìÚ°Ú¹È¿ÈÃÈÏÚ¨ÉÈÉè÷êÚ·ÚÅÊ«Ê©Ê¶ÊÐÊÔÚÖÊëÛÓÊìÊüË¥Ë­ËµËÏËÐËßÚÕÚÇÌ¸Ì·ÌÆÌÖÍ¤Í¥ÙïâÕÍäÍöÍýÍüÍûÎªÚÃÎ½ÎÄÎÉÎÜâÐÎóÚÀÏ¯ÏåÏêâÔÏíÐ³ÙôÐ»âÝâÓÐíÚ¼ÐòÐóÚÎÐþÐýÚÊÑ¯ÑµÑ¶ÑÈÚ¥ÑÔÑèÚÝÒ¥Ò¹ÚËÒÂÚ±ì½ÒåÒéÒàÒëÒèÞÈÞÄÒêÒáÓ¦âßÓ¥ÙøÓ®Ó¹ÓºÛÕ÷ÓÓÀÓÕì¶ÚÄÓïâ×ÓýÚÍÚÚÕ©Õ«ì¹ÚÞÚ¯ÕØÕÚÚØÕâðÑÕïÖ¤ÚºÖÔÖÝÖßÖïÖîØ¼Ö÷÷æ×¯×»ÚÂÚÑÚÁ×ä×å×ç×ù';
Begin
	If v_Instr Is Null Or Length(Ltrim(v_Instr)) = 0 Then
		v_Code := '';
	Else
		v_Input := Upper(v_Instr);
		v_Code  := '';
		For v_Bitnum In 1 .. Length(v_Input) Loop
            if Instr('ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.+-*/', Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || Substr(v_Input, v_Bitnum, 1);
			Elsif Instr(v_a, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'A';
			Elsif Instr(v_b, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'B';
			Elsif Instr(v_c, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'C';
			Elsif Instr(v_d, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'D';
			Elsif Instr(v_e, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'E';
			Elsif Instr(v_f, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'F';
			Elsif Instr(v_g, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'G';
			Elsif Instr(v_h, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'H';
			Elsif Instr(v_i, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'I';
			Elsif Instr(v_j, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'J';
			Elsif Instr(v_k, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'K';
			Elsif Instr(v_l, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'L';
			Elsif Instr(v_m, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'M';
			Elsif Instr(v_n, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'N';
			Elsif Instr(v_o, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'O';
			Elsif Instr(v_p, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'P';
			Elsif Instr(v_q, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'Q';
			Elsif Instr(v_r, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'R';
			Elsif Instr(v_s, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'S';
			Elsif Instr(v_t, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'T';
			Elsif Instr(v_u, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'U';
			Elsif Instr(v_v, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'V';
			Elsif Instr(v_w, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'W';
			Elsif Instr(v_x, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'X';
			Elsif Instr(v_y, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'Y';
			End If;
			Exit When Length(v_Code) > 9;
		End Loop;
	End If;
	Return(v_Code);
End;
/

--¹«¹²Í¬Òå´Ê
Create Public Synonym zlNotices for zlNotices
/
Create Public Synonym zlNoticeUsr for zlNoticeUsr
/
Create Public Synonym zlNoticeRec for zlNoticeRec
/
Create Public Synonym ZL_ZLNOTICES_INSERT for ZL_ZLNOTICES_INSERT
/
Create Public Synonym ZL_ZLNOTICES_UPDATE for ZL_ZLNOTICES_UPDATE
/
Create Public Synonym ZL_ZLNOTICES_DELETE for ZL_ZLNOTICES_DELETE
/
Create Public Synonym ZL_ZLNOTICEUSR_INSERT for ZL_ZLNOTICEUSR_INSERT
/
Create Public Synonym ZL_ZLNOTICEUSR_DELETE for ZL_ZLNOTICEUSR_DELETE
/
Create Public Synonym ZL_ZLNOTICEREC_CHECKNOTICE for ZL_ZLNOTICEREC_CHECKNOTICE
/
Create Public Synonym ZL_ZLNOTICEREC_NOTICE for ZL_ZLNOTICEREC_NOTICE
/
--È¨ÏÞ
Grant select on zlNotices to PUBLIC
/
Grant select on zlNoticeUsr to PUBLIC
/
Grant select on zlNoticeRec to PUBLIC
/
Grant execute on ZL_ZLNOTICES_INSERT to PUBLIC
/
Grant execute on ZL_ZLNOTICES_UPDATE to PUBLIC
/
Grant execute on ZL_ZLNOTICES_DELETE to PUBLIC
/
Grant execute on ZL_ZLNOTICEUSR_INSERT to PUBLIC
/
Grant execute on ZL_ZLNOTICEUSR_DELETE to PUBLIC
/
Grant execute on ZL_ZLNOTICEREC_CHECKNOTICE to PUBLIC
/
Grant execute on ZL_ZLNOTICEREC_NOTICE to PUBLIC
/

--2005Äê1ÔÂ6ÈÕ²¹³ä
Grant Execute On zlSpellCode to Public
/
Create Public Synonym zlSpellCode For zlSpellCode
/

Grant Execute On zlWBcode to Public
/
Create Public Synonym zlWBcode For zlWBcode
/