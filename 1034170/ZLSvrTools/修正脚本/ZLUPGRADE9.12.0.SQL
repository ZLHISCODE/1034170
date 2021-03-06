---BEGIN:自动提醒功能脚本
Create Table zlNotices(
	序号		NUMBER(5) NOT NULL,
	系统		NUMBER(5),		
	提醒条件	VARCHAR2(4000),
	提醒内容	VARCHAR2(250),
	提醒报表	VARCHAR2(50),
	提醒声音	NUMBER(5),
	提醒窗口	NUMBER(1),
	提醒顺序	VARCHAR2(200) DEFAULT '',
	检查周期	NUMBER(5),
	提醒周期	NUMBER(5),	
	开始时间	DATE NOT NULL,
	终止时间	DATE)			
	PCTFREE 5
	PCTUSED 90
	STORAGE (INITIAL 512 NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/

Alter Table zlNotices ADD CONSTRAINT 
	zlNotices_PK PRIMARY KEY (序号)
	USING INDEX PCTFREE 5
	STORAGE(INITIAL 256 NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/
Create Table zlNoticeUsr(
	提醒序号	NUMBER(5) NOT NULL,
	提醒对象	NUMBER(1) DEFAULT 0,	--0-所有;1-指定人员;2-指定部门;3-指定工作站
	对象名称	VARCHAR2(50))			
	PCTFREE 5
	PCTUSED 90
	STORAGE (INITIAL 512 NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/
Create Table zlNoticeRec(
	提醒序号	NUMBER(5) NOT NULL,	
	用户名		VARCHAR2(30),
	检查时间	DATE,
	检查结果	NUMBER(1) DEFAULT 0,	--1表示有要提醒的内容;0表示无要提醒的内容
	提醒标志	NUMBER(1) DEFAULT 0,	--1表示要提醒;0表示不提醒
	提醒时间	DATE,
	提醒内容	VARCHAR2(250))			
	PCTFREE 5
	PCTUSED 90
	STORAGE (INITIAL 512 NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/

Alter Table zlNoticeUsr ADD CONSTRAINT 
	zlNoticeUsr_PK PRIMARY KEY (提醒序号,提醒对象,对象名称)
	USING INDEX PCTFREE 5
	STORAGE(INITIAL 256 NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/
Alter Table zlNoticeUsr ADD CONSTRAINT 
	zlNoticeUsr_FK_提醒序号 FOREIGN KEY(提醒序号) 
	REFERENCES zlNotices(序号) ON DELETE CASCADE
/
Alter Table zlNoticeRec ADD CONSTRAINT 
	zlNoticeRec_PK PRIMARY KEY (提醒序号,用户名)
	USING INDEX PCTFREE 5
	STORAGE(INITIAL 256 NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/
Alter Table zlNoticeRec ADD CONSTRAINT 
	zlNoticeRec_FK_提醒序号 FOREIGN KEY(提醒序号) 
	REFERENCES zlNotices(序号) ON DELETE CASCADE
/
---END:自动提醒功能脚本


Create Table zlUserRoles(
    用户 varchar2(30),
    角色 varchar2(30),
    管理 number(1))
    PCTFREE 5
    PCTUSED 90
    STORAGE (INITIAL 256K NEXT 8K PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/

ALTER TABLE zlUserRoles ADD CONSTRAINT 
    zlUserRoles_PK UNIQUE (用户,角色)
    USING INDEX PCTFREE 10
    STORAGE(INITIAL 32K NEXT 2K PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/

Alter Table zlReports Add 动态纸张 NUMBER(1)
/

Alter Table zlReports ADD CONSTRAINT zlReports_CK_动态纸张 Check(动态纸张 IN(0,1))
/

--自动提醒功能脚本:设置菜单位置
insert into zlSvrTools(编号,上级,标题,快键,说明) values ('0504','05','自动提醒','H',Null)
/

----------------------------------------------------------------------------
---  INSERT   for   ZLNOTICES
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_ZLNOTICES_INSERT(
	序号_IN IN ZLNOTICES.序号%TYPE,
	系统_IN IN ZLNOTICES.系统%TYPE,
	提醒条件_IN IN ZLNOTICES.提醒条件%TYPE,
	提醒内容_IN IN ZLNOTICES.提醒内容%TYPE,
	提醒报表_IN IN ZLNOTICES.提醒报表%TYPE,
	提醒声音_IN IN ZLNOTICES.提醒声音%TYPE,
	提醒窗口_IN IN ZLNOTICES.提醒窗口%TYPE,
	检查周期_IN IN ZLNOTICES.检查周期%TYPE,
	提醒周期_IN IN ZLNOTICES.提醒周期%TYPE,
	开始时间_IN IN ZLNOTICES.开始时间%TYPE,
	终止时间_IN IN ZLNOTICES.终止时间%TYPE,
	提醒顺序_IN IN ZLNOTICES.提醒顺序%TYPE
)
IS
BEGIN
	Insert Into ZLNOTICES
		(序号,系统,提醒条件,提醒内容,提醒报表,提醒声音,提醒窗口,检查周期,提醒周期,开始时间,终止时间,提醒顺序)
		VALUES
		(序号_IN,系统_IN,提醒条件_IN,提醒内容_IN,提醒报表_IN,提醒声音_IN,提醒窗口_IN,检查周期_IN,提醒周期_IN,开始时间_IN,终止时间_IN,提醒顺序_IN);
END ZL_ZLNOTICES_INSERT;
/

----------------------------------------------------------------------------
---  UPDATE   for   ZLNOTICES
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_ZLNOTICES_UPDATE(
	序号_IN IN ZLNOTICES.序号%TYPE,
	系统_IN IN ZLNOTICES.系统%TYPE,
	提醒条件_IN IN ZLNOTICES.提醒条件%TYPE,
	提醒内容_IN IN ZLNOTICES.提醒内容%TYPE,
	提醒报表_IN IN ZLNOTICES.提醒报表%TYPE,
	提醒声音_IN IN ZLNOTICES.提醒声音%TYPE,
	提醒窗口_IN IN ZLNOTICES.提醒窗口%TYPE,
	检查周期_IN IN ZLNOTICES.检查周期%TYPE,
	提醒周期_IN IN ZLNOTICES.提醒周期%TYPE,
	开始时间_IN IN ZLNOTICES.开始时间%TYPE,
	终止时间_IN IN ZLNOTICES.终止时间%TYPE,
	提醒顺序_IN IN ZLNOTICES.提醒顺序%TYPE
)
IS
BEGIN
	Update ZLNOTICES
		Set 序号=序号_IN,
		    系统=系统_IN,
		    提醒条件=提醒条件_IN,
		    提醒内容=提醒内容_IN,
		    提醒报表=提醒报表_IN,
		    提醒声音=提醒声音_IN,
		    提醒窗口=提醒窗口_IN,
		    检查周期=检查周期_IN,
		    提醒周期=提醒周期_IN,
		    开始时间=开始时间_IN,
		    终止时间=终止时间_IN,
		    提醒顺序=提醒顺序_IN
		Where  序号=序号_IN;
END ZL_ZLNOTICES_UPDATE;
/

----------------------------------------------------------------------------
---  DELETE   for   ZLNOTICES
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_ZLNOTICES_DELETE(
	序号_IN IN ZLNOTICES.序号%TYPE
)
IS
BEGIN
	Delete From ZLNOTICEUSR Where  提醒序号=序号_IN;
	Delete From ZLNOTICES Where  序号=序号_IN;
END ZL_ZLNOTICES_DELETE;
/

----------------------------------------------------------------------------
---  INSERT   for   ZLNOTICEUSR
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_ZLNOTICEUSR_INSERT(
	提醒序号_IN IN ZLNOTICEUSR.提醒序号%TYPE,
	提醒对象_IN IN ZLNOTICEUSR.提醒对象%TYPE,
	对象名称_IN IN ZLNOTICEUSR.对象名称%TYPE
)
IS
BEGIN
	Insert Into ZLNOTICEUSR
		(提醒序号,提醒对象,对象名称)
		VALUES
		(提醒序号_IN,提醒对象_IN,对象名称_IN);
END ZL_ZLNOTICEUSR_INSERT;
/
----------------------------------------------------------------------------
---  DELETE   for   ZLNOTICEUSR
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_ZLNOTICEUSR_DELETE(
	提醒序号_IN IN ZLNOTICEUSR.提醒序号%TYPE
)
IS
BEGIN
	Delete From ZLNOTICEUSR
		Where  提醒序号=提醒序号_IN;
END ZL_ZLNOTICEUSR_DELETE;
/
----------------------------------------------------------------------------
---  检查更新
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_ZLNOTICEREC_CHECKNOTICE(
	用户名_IN	IN	zlNoticeUsr.对象名称%type,
	部门名_IN	IN	zlNoticeUsr.对象名称%type:='',
	工作站_IN	IN	zlNoticeUsr.对象名称%type:='',
	启动检查_IN	IN	NUMBER:=0)
IS
	Cursor c_Notices IS
		SELECT A.*,B.检查时间 FROM zlNotices A,
			(SELECT 提醒序号,检查时间 FROM zlNoticeRec WHERE 用户名=用户名_IN) B 
		WHERE A.序号=B.提醒序号(+)
			AND ((检查周期 IS NULL AND 1=启动检查_IN) OR (检查周期 IS NOT NULL AND 0=启动检查_IN))
			AND A.开始时间<=SYSDATE AND (A.终止时间>=SYSDATE OR A.终止时间 IS NULL)
			AND (A.序号 IN (SELECT 提醒序号 FROM zlNoticeUsr
					WHERE (提醒对象 = 1 AND 对象名称 = 用户名_IN) 
						OR (提醒对象 = 2 AND 对象名称 = 部门名_IN) 
						OR (提醒对象 = 3 AND 对象名称 = 工作站_IN))
			OR A.序号 NOT IN (SELECT 提醒序号 FROM zlNoticeUsr));

	r_Notice c_Notices%RowType;
	
	v_检查结果 number(1);
	v_提醒内容 varchar2(500);
	
	v_检查否 number(1);
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
		
		v_检查否:=0;
				
		--通过检查上次检查时间是否为空来检测是否为第一次检查
		if r_Notice.检查时间 is null then		
			--第一次检查,新增记录
			insert into zlNoticeRec(提醒序号,用户名,检查时间,提醒时间,提醒内容) values (r_Notice.序号,用户名_IN,SYSDATE,NULL,NULL);
			v_检查否:=1;
		else
			--第 2、3、... 次检查
			--当前时间是否大于上次检查时间加上一个检查周期,如果大于了，则更新检查时间
			if r_Notice.检查周期 is null then
				if 启动检查_IN=1 then
					update zlNoticeRec set 检查时间=SYSDATE	where 提醒序号=r_Notice.序号 and 用户名=用户名_IN;
					v_检查否:=1;
				end if;
			else
				if SYSDATE>=(r_Notice.检查时间+r_Notice.检查周期/(24*60)) then
					update zlNoticeRec set 检查时间=SYSDATE	where 提醒序号=r_Notice.序号 and 用户名=用户名_IN;
					v_检查否:=1;
				end if;
			end if;
		end if;	
		
		if v_检查否=1 then
			v_检查结果:=0;
			v_提醒内容:='';

			--检查提醒		
			if not (r_Notice.提醒条件 is null) then					
				v_提醒内容:=r_Notice.提醒内容;

				--strTmp格式:如'[姓名];varchar2|[性别];date'
				v_Tmp:=r_Notice.提醒顺序||'|';
				WHILE not (v_Tmp is null) LOOP

					v_Pos := instr(v_Tmp, '|');								
					v_TmpField:=substr(v_Tmp,1,v_Pos - 1);	
					
					v_FieldPos:=instr(v_TmpField,';');
					v_Field:=substr(v_TmpField,1,v_FieldPos - 1);
					v_FieldType:=trim(Upper(substr(v_TmpField,v_FieldPos+1,100)));

					v_Tmp:=trim(substr(v_Tmp,v_Pos + 1,1000));

					v_Pos:=instr(v_提醒内容,v_Field);

					if v_Pos>0 then
						
						v_Result:=trim(substr(v_Field,2,1000));
						v_Result:=substr(v_Result,1,LENGTH(v_Result)-1);

						if v_FieldType='NUMBER' then
							v_Result:='to_char('||v_Result||')';
						Elsif v_FieldType='DATE' then
							v_Result:='to_char('||v_Result||',''yyyy-mm-dd'')';
						End if;

						v_提醒内容:=trim(substr(v_提醒内容,1,v_Pos - 1)||'''||'||v_Result||'||'''||substr(v_提醒内容,v_Pos + length(v_Field),1000));

					end if;

				END LOOP;
				v_Pos:=instr(Upper(r_Notice.提醒条件),' FROM ');
				
				if v_Pos>0 then
					v_SQL:=TRIM('SELECT '''||v_提醒内容||''''||substr(r_Notice.提醒条件,v_Pos,4000));

					v_CursorID:=sys.DBMS_SQL.OPEN_CURSOR;
					sys.DBMS_SQL.PARSE(v_CursorID,v_SQL,sys.DBMS_SQL.NATIVE);
					
					dbms_sql.define_column(v_CursorID,1,v_Result,1000);

					v_return :=DBMS_SQL.execute(v_CursorID); 
					
					if DBMS_SQL.FETCH_ROWS(v_CursorID)>0 then
						--检查后有新的情况发生

						v_检查结果:=1;	
						dbms_sql.column_value(v_CursorID,1,v_Result);
						v_提醒内容:=trim(v_Result);						
					end if;
				end if;
				
			else
				v_检查结果:=1;
				v_提醒内容:=r_Notice.提醒内容;
			end if;
			
			update zlNoticeRec set 检查结果=v_检查结果,提醒内容=v_提醒内容 where 提醒序号=r_Notice.序号 and 用户名=用户名_IN;
		end if;
	END Loop;

END ZL_ZLNOTICEREC_CHECKNOTICE;
/
----------------------------------------------------------------------------
---  提醒更新
----------------------------------------------------------------------------
CREATE OR REPLACE PROCEDURE ZL_ZLNOTICEREC_NOTICE(
	用户名_IN	IN	zlNoticerec.用户名%type,
	启动检查_IN	IN	NUMBER:=0)
IS
	Cursor c_Notices IS
		SELECT B.*,A.提醒周期,A.检查周期 FROM zlNotices A,
			(SELECT * FROM zlNoticeRec WHERE 用户名=用户名_IN) B 
		WHERE A.序号=B.提醒序号
			AND B.检查结果=1 
			AND ((启动检查_IN=1 AND 检查周期 IS NULL) OR (检查周期 IS NOT NULL AND 启动检查_IN=0))
			AND A.开始时间<=SYSDATE AND (A.终止时间>=SYSDATE OR A.终止时间 IS NULL);

	r_Notice c_Notices%RowType;
  
	v_提醒否 number(1);
  
BEGIN
	
	Update zlNoticeRec Set	提醒标志=0 where 用户名=用户名_IN;
		
	FOR r_Notice In c_Notices Loop
		
		v_提醒否:=0;

		if r_Notice.提醒时间 is null then
			
			--第一次提醒
			v_提醒否:=1;
		else
			--第 2、3、... 次提醒
			if r_Notice.检查周期 is null then
				if 启动检查_IN=1 then
					v_提醒否:=1;
				end if;
			else
				--当前时间是否大于上次提醒时间加上一个提醒周期
				if SYSDATE>=(r_Notice.提醒时间+r_Notice.提醒周期/(24*60)) then
					v_提醒否:=1;
				end if;
			end if;
		end if;		
		
		if v_提醒否=1 then
			Update zlNoticeRec Set	提醒时间=SYSDATE,
						提醒标志=v_提醒否
			where 提醒序号=r_Notice.提醒序号 
				and 用户名=用户名_IN;
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
	v_Stdstr  Varchar2(50) := '芭擦搭蛾发噶哈击-喀垃妈拿哦啪期然撒塌挖-挖昔压匝';
	v_Chara   Varchar2(2000) := '吖锕錒嗄锿鎄捱嗳躷﨟譪霭靄砹嗌嫒暧瑷賹鴱譺鑀鱫靉桉庵谙誝鞌諳闇鮟鵪韽鶕雸埯铵隌揞銨犴豻貋錌黯醠坳軪隞嗷廒獒遨謷鳌鏖鰲鷔鼇拗媪镺岙骜謸鏊鹌聱螯';
	v_Charb   Varchar2(2000) := '岜粑釟豝鲃魞茇釛菝軷颰魃鼥钯鈀跁鲅鮁灞掰捭呗鞁贁韛頒辬阪坂钣舨鈑魬闆鉡靽辦浜邫鞤蒡謗鎊勹孢煲龅闁齙葆飹飽鳵鴇賲靌趵铇鉋靤鮑鑤陂鵯鉳貝邶悖鄁軰碚蓓誖輩鋇鞴鐾贲逩賁锛錛畚坌輽閍嘣鞛逬跰甏镚鏰豍鲾鎞鵖鰏荸匕吡妣秕俾舭貏诐邲畀哔荜狴铋婢庳萆閇閉弼愎貱赑滗跸鉍閟飶鄪嬖薜鮅濞蹕鞞髀璧鄨襞鏎鞸韠躃躄魓贔鐴鷝鷩鼊砭煸邉鍽鳊邊鯾鯿匾貶碥鴘弁忭汴苄釆拚缏閞辡頨辧辪辮辯變灬杓飑髟颮骠麃镖飙飚颷謤贆鏢镳飆飇飈飊鑣婊諘錶鳔鰾鱉鼈龞蹩邠傧缤槟豩賓賔镔豳霦鑌顮殡膑髌鬓鬢冫鋲邴陃禀鈵鉼鞆餅餠摒誁鮩靐饽啵鉢餑蹳鱍孛郣亳钹鈸鉑鲌踣鋍镈鮊豰鎛鵓礴鑮跛簸擘檗譒逋钸晡鈽誧餔轐醭卟鳪鵏鸔钚瓿鈈踄郶鹁瘢癍裱褙褊篦箅筚笾蝙褓裨窆瘭鸨鹎';
	v_Charc   Varchar2(2000) := '嚓礤遪財跴飡骖黪黲粲璨謲伧鸧鶬鑶賶嘈漕艚鏪艹鄵鼜恻岑涔噌杈馇銟锸鍤鎈猹靫槎檫蹅镲鑔汊姹钗釵侪辿觇鋓婵孱禅誗鋋廛潺鄽镡酁躔镵讒鑱谄蒇諂閳冁醦譂鏟闡讇忏羼韂顫伥娼菖阊锠錩閶鲳鯧鼚苌長镸徜嫦鋿鲿鏛鱨昶惝氅鋹怅鬯誯韔怊焯鈔晁鄛鼌轈鼂謿麨車砗屮坼迠頙抻琛嗔諃賝謓迧宸陳谌軙鈂霃諶麎鷐趻碜踸贂龀趂榇齓齔谶讖阷柽铛赪靗瞠赬頳鏳鏿鐺丞枨郕埕铖塍誠酲鋮鯎哧眵嗤媸誺鴟鵄魑齝麶黐茌赿貾遅趍遟墀踟遲謘豉鉹齒彳叱饬迣敕啻飭傺跮鉓雴遫銐趩鶒鷘忡茺舂憧艟蹖隀铳銃俦帱惆酧雔雠躊醻讎讐醜魗遚樗貙齣刍豠趎鉏鋤雛蹰鶵躕杵楮齭齼亍怵绌豖鄐踀閦諔憷黜搋啜嘬踹巛氚舡遄輲舛钏釧賗鶨闖怆龡陲棰槌錘鎚顀輴鰆鶞陙莼醕錞鯙鶉賰踳踔辶辵逴辍酫趠輟龊齪鑡齱呲趀祠茈辝鈶糍辤飺餈鴜辭鶿鷀跐賜苁枞骢璁鏦淙琮誴賨賩謥楱腠辏輳麁麄麤徂殂猝酢蔟誎趗踧蹙鼀蹴蹵顣汆撺镩躥鑹爨榱鏙璀趡啐悴萃毳顇邨踆忖遳蹉醝嵯矬鹾鹺齹脞厝逪锉銼錯澶瘥隹篪笞蚩虿耖皴褫褚裎衩瘳蟾螬螭蝽蜍蛏瘛痤鸱骣鹚鹑膪';
	v_Chard   Varchar2(2000) := '哒耷嗒鎝迏迖妲怛沓逹達跶靼鞑鎉躂鐽韃龖龘呔轪岱甙绐迨玳軑埭軚貸軩鮘鴏黛蹛霴黱靆眈躭酖殚鄲頕儋黕啖萏誕澹鴠贉霮谠譡黨讜凼宕砀菪逿雼趤闣刂叨忉氘釖鱽魛陦﨩隝隯焘軇纛锝鍀豋噔簦戥鄧隥嶝磴镫鐙羝隄趆嘀镝鍉鞮鏑籴荻觌靮頔魡豴鸐氐诋邸阺坻柢砥軧骶鯳娣逓谛釱棣睇遞鉪碲遰諦踶嗲蹎巅顚顛齻踮點阽坫玷钿鈿電簟貂鳭鮉鲷鼦鯛鵰釣铞鈟銱雿調鋽鑃垤喋堞揲趃牒镻諜蹀鲽鰈仃玎酊釘靪頂鼑鐤飣啶腚碇錠顁铥颩銩咚岽氡鮗鼕鯟鶇鶫諌垌峒胨迵胴硐霘蔸阧钭﨣鈄郖鬥酘閗鬦鋀餖闘鬪鬬鬭阇嘟醏闍渎椟牍読錖黩讀豄贕韣髑鑟韇韥黷讟賭芏靯鍍鍴椴煅鍛躖頧鴭鐜怼陮隊碓憝镦譈鐓譵礅蹾盹趸躉沌炖砘逇鈍頓遯踲咄铎鈬踱鮵鐸哚缍趓躱軃鬌沲陊陏跢跥飿鵽瓞簖篼箪蚪聃耵耋褡裰裆窦癫癜瘅笪笃蠹疸疔鸫';
	v_Chare   Varchar2(2000) := '屙迗莪鈋锇誐鋨頟魤額鵝鵞譌婀鵈阨呃苊阸轭垩谔軛阏愕萼豟軶遌腭锷遻頞餓噩諤閼鍔鳄顎鰐鶚讍鑩齶鱷蒽摁鞥陑輀鲕隭鮞鴯轜迩珥铒鉺餌邇趰佴貮貳鸸颚鹗';
	v_Charf   Varchar2(2000) := '醗垡閥砝鍅幡轓颿飜鱕釩蕃燔蹯蘩鐇鷭辺畈軓梵販軬飯飰匚邡枋钫趽鈁錺鴋鲂魴彷舫鶭妃飛绯扉靟霏鲱鯡飝淝腓悱斐榧翡誹狒費镄鼣鐨靅玢躮鈖雰棼隫魵鳻豮鼢鼖豶轒鐼黂黺偾鲼瀵鱝沣砜風葑鄷鋒豐鎽鏠酆靊飌麷唪諷俸赗鳯鳳鴌賵雬鴀邞呋趺酜麸稃跗鈇鄜豧鳺麩麬麱凫孚芙芾怫绂绋苻祓罘茯郛韨鳬砩莩匐桴艴菔﨓鉘鉜颫鳧韍幞鴔諨踾輻鮄鮲黻鵩鶝呒拊郙釡滏輔鬴黼阝驸負陚鲋赙賦輹鮒賻鍑鍢鳆鰒馥篚蚨蜚蝠缶蝮蜉痱';
	v_Charg   Varchar2(2000) := '旮伽钆尜釓錷尕尬魀郂陔垓赅隑豥賅賌鎅丐鈣戤迀坩泔苷酐尴鳱魐秊澉趕橄擀鳡鱤旰矸绀淦贛阬罡釭鋼鎠戆槔睾韟鷎鼛鷱杲缟槁藁鎬诰郜锆誥鋯圪纥閤鴐鴚謌鴿鎶鬲嗝塥搿膈閣镉鞈韐骼諽輵鮯鎘韚轕鞷鰪哿舸硌鉻哏亘艮茛赓鹒賡鶊郠哽绠鲠鯁肱觥躳龏龔廾珙輁鞏貢贑佝缑鈎鉤鞲韝岣枸豿诟媾彀遘雊觏購轱菰觚軱軲酤毂鈲鮕鴣轂鹘鶻汩诂牯罟逧钴鈷鼔嘏臌瞽鵠崮梏牿锢頋錮鲴鯝顧胍颪趏銽颳鴰呱卦诖倌関闗鳏關鰥鱞輨錧躀鳤掼涫貫遦盥雚鏆鑵鸛鱹咣桄胱輄銧黆犷妫邽郌閨鲑鮭龜鬶鬹蘒宄庋匦陒軌晷刿炅貴鳜鞼鱖鱥丨衮绲磙輥鲧鮌鯀謴呙埚崞鈛鍋帼掴虢馘猓椁輠餜鐹過簋篝筻笱蝈蜾蛄蚣虼聒矜袼疳鹳鹄痼鸹鸪皈';
	v_Charh   Varchar2(2000) := '铪鉿嗨胲酼醢餀頇谽魽鼾邗晗焓鋡韓豃鬫闬菡釬閈撖銲鋎頷顄譀雗瀚鶾魧迒绗貥頏沆蒿嚆薅嗥濠譹昊灏顥鰝诃嗬劾郃曷盍龁貈鉌阖鲄閡鹖麧頜翮魺闔鞨齕鶡鑉龢隺賀壑鶴齃靍靎鸖靏黒鞎桁珩鸻鴴鵆蘅鑅訇軣谾薨輷鍧轟闳泓荭谹鈜閎谼鉷鞃魟鋐蕻霐黉霟鴻黌讧閧銾闀闂鬨齁銗糇骺鍭鯸郈後逅鄇堠豞鲎鲘鮜鱟烀轷唿惚軤雽滹雐謼囫斛猢煳槲魱醐頶觳鍸鬍鰗鶘鶦鶮浒琥錿鯱冱岵怙戽祜扈鄠鳸鍙護鳠韄頀鱯鸌誮錵骅铧鋘譁鏵鷨桦諙諣黊踝鴅鵍酄獾貛讙郇洹萑雈貆锾阛寰缳還豲鍰镮鹮轘闤鐶鬟輐奂浣逭漶鲩擐鯇鯶鰀肓隍黃徨湟遑潢锽璜諻鍠鳇趪韹鐄鰉鱑鷬謊鎤诙咴晖珲豗隓輝麾隳鰴洄茴迴逥鮰譭哕浍荟恚桧彗喙缋阓賄誨蕙諱頮譓譮鏸闠鐬靧韢譿顪阍閽馄餛轋鼲诨溷諢锪劐鍃攉邩钬鈥夥閄貨嗀謋雘镬嚯藿鑊靃皓篌篁蚝虺颢颔颌颃顸耠癀笏蠖蟪蟥蚵蚶瘊鹱鹕瓠';
	v_Charj   Varchar2(2000) := '丌叽乩玑芨矶咭剞唧屐飢嵇犄赍跻鳮銈畿賫躸齑墼錤隮羁賷鄿雞譏韲鶏譤鐖躋鞿鷄齎鑇鑙齏鸄岌亟佶郆﨤谻戢殛楫蒺趌銡蕺踖鞊鹡輯蹐鍓轚鏶霵鶺鷑躤雦雧掎鱾戟嵴麂魢彐芰哜洎觊偈跡際暨誋跽霁鲚諅鲫髻鮆蹟鯽鵋齌骥鯚鱀霽鰶鰿鱭迦浃珈袈葭跏鉫镓豭貑鎵麚岬郏郟恝戛铗跲餄鋏頬頰鴶鵊胛賈鉀戋菅豜湔犍間靬搛缣蒹豣鲣鳽鋻鞬麉鞯鳒鵳鰔譼鰜鶼韀鰹鑯韉囝枧趼睑锏谫戬翦謇蹇謭鬋鰎鹸鐗鐧鹻譾鹼牮谏釼楗毽腱跈閒賎僭諓賤趝踐踺諫鍵餞鍳鏩轞鑑鑒鑬鑳茳豇缰鳉礓韁鱂講顜洚绛犟醤糨醬謽艽姣茭跤僬鲛鮫鵁轇鐎鷦鷮佼挢湫敫賋踋鉸餃徼鵤譑鱎峤較噍趭轎醮譥釂階喈嗟鞂鶛卩孑讦诘拮迼桀婕鉣魝碣鲒羯誱踕頡鍻鮚飷骱誡魪钅釒鹶黅卺堇廑馑槿瑾錦謹妗荩赆進缙觐噤賮贐齽泾旌菁腈鵛鯨鶁鶄麖鼱麠阱刭肼儆憬頸弪迳胫逕婧靓獍誩踁頚靚靜鏡冂扃迥逈颎顈赳阄啾鳩鬏鬮镹韮柩桕僦鯦麔齨鷲苴陱掬椐琚趄跔锔雎諊踘鋦鮈鴡鞫鶋郹輂跼趜躹閰橘鵙蹫鵴鶪鼰鼳莒榉榘龃﨔踽齟讵苣邭钜倨犋跙鉅飓豦屦鮔遽鋸颶瞿貗躆醵鐻涓鋑鋗镌鎸鵑鐫蠲锩錈桊狷隽鄄雋飬餋噘孓珏崛桷觖赽趹逫厥趉鈌劂谲獗蕨鴂鴃噱橛镼镢譎蹶蹷鶌矍鐍鐝爝鷢龣貜躩钁軍鈞銁銞鲪麇鍕鮶麏麕陖捃餕鵔鵕鵘稷鹣疖瘕筠笈蛟蛱蚧虮颉皲裾裥袷衿窭瘠痂鹫笳笕笄耩鹪鸠皎';
	v_Chark   Varchar2(2000) := '咔佧胩鉲锎開鐦剀垲恺闿铠蒈輆锴鍇鎧闓颽忾鎎龛戡龕侃莰輡轁顑轗阚瞰闞躿鏮鱇伉邟闶钪鈧閌尻栲铐犒銬鲓鮳鯌珂轲趷钶軻稞鈳瞌頦醘顆髁岢恪氪骒缂嗑溘锞課錁豤貇錹铿誙銵鍞鏗倥崆躻躼錓鵼鞚芤眍叩釦蔻鷇刳郀堀跍骷鮬绔喾誇侉銙蒯郐哙狯脍鲙鄶鱠髋鑧诓邼哐誆軭诳軖軠誑鵟夼邝圹纩贶貺軦鉱鋛鄺黋鑛悝闚顝逵鄈頄馗喹揆暌睽頯鍨鍷夔躨跬頍蹞匮喟愦蒉謉鐀鑎琨锟髡鹍醌錕鲲鯤鵾鶤悃阃閫閸栝頢闊鞟韕霩鞹鬠疴蛞篑箜筘蝌蝰颏裉窠聩';
	v_Charl   Varchar2(2000) := '邋旯砬剌辢鬎镴鯻鑞鞡崃徕涞郲逨铼錸鯠鶆麳赉睐賚濑賴頼顂鵣籁岚斓镧闌譋讕躝鑭钄韊榄漤罱醂啷郎郞莨稂锒郒躴鋃鎯阆誏閬蒗唠崂铹醪鐒顟栳铑銠鮱轑軂仂阞叻泐韷鳓鰳餎嫘缧檑羸鐳轠鑘靁鱩鼺诔誄讄鑸鸓酹銇頛頪錑颣類嘞塄踜愣骊喱缡蓠嫠貍鋫鲡罹錅謧醨藜邌釐離鯏鏫鯬鵹黧鑗鱺鸝礼俚娌逦锂豊裏鋰澧鯉醴鳢邐鱧呖坜苈戾枥俪栎赲轹郦猁砺莅唳粝詈跞雳溧鉝鳨隷鴗隸麗酈鷅麜躒轢讈轣靂鱱靋奁連鲢濂臁蹥謰鎌譧鬑鐮鰱琏蔹鄻娈殓楝潋錬鍊鏈鰊凉椋辌墚踉輬魉魎輌諒輛鍄蹽嘹寮獠缭遼豂賿蹘鐐飉鷯钌釕鄝蓼镽尥咧冽洌迾埒捩趔颲鮤鴷躐鬛鬣鱲啉粼鄰隣隣嶙遴辚瞵麐轔鏻麟鱗廪懔檩顲賃蔺膦閵蹸躏躙躪轥囹泠苓柃瓴鸰棂绫翎跉軨鈴閝輘霊錂霗魿鲮鴒鹷霛霝齢酃鯪齡醽靈麢龗阾領呤熘浏旒遛骝飗镏鹠镠鎏鎦麍鏐飀鐂飅鰡鶹绺锍鋶蹓霤雡飂鬸鷚泷茏栊珑胧砻龍鏧霳龒龓豅躘鑨靇鸗垅隴贚偻喽蒌遱謱軁髅鞻嵝镂鏤噜撸垆泸栌胪轳舻鈩鲈魲轤鑪顱鱸鸕黸鹵魯橹镥鏀鐪鑥辂陸渌逯賂輅漉趢踛辘醁錄録錴璐鴼蹗轆鏕鯥鵦鵱鏴鷺氇闾榈閭鷜郘稆膂鋁鑢栾脔銮鵉鑾鸞釠锊鋝鋢囵陯踚輪錀鯩論捋頱猡脶椤镙鏍邏鸁鑼倮躶蠃泺荦珞摞漯雒鮥鸬鹩簏篥笠蠡蠊蝼螂蜊蛉蛎聆癞癃瘰瘘瘌痨疬疠鹭鹨鸾耧耢耒褴褛裣裢鹂';
	v_Charm   Varchar2(2000) := '嬷犸遤鎷鷌鰢杩閁唛鬕霾荬買鷶劢麥賣邁霡霢顢鞔鳗鬗鬘鰻鏋鄤墁幔缦熳镘謾鏝邙硭釯铓鋩漭貓牦旄軞酕髦錨鶜峁泖茆昴鉚耄袤貿鄚瑁瞀鄮懋莓郿嵋湄猸楣镅鋂鎇鶥黴浼躾鎂黣跊鬽韎魅扪钔門閅鍆焖懑雺甍瞢鄳鄸朦礞鯍艨鹲靀顭鸏勐艋錳懵鯭鼆霥霿踎咪祢猕謎縻麊麋麛蘼镾醾醿鸍釄芈弭敉脒銤冖糸汨宓谧嘧鼏謐宀沔黾眄湎腼鮸靣麪麫麺麵喵鶓鱙杪眇淼缈邈乜咩鴓鑖鱴岷玟苠珉缗鈱賯錉鴖鍲闵泯閔愍黽閩鳘鰵茗冥鄍溟暝銘鳴瞑酩缪謬谟嫫馍麼麽魹謨謩譕麿殁茉秣貃蓦貊銆靺镆魩黙貘鏌哞侔眸鉾謀鍪鴾麰毪鉧踇仫沐坶苜钼雮鉬霂鞪鹋袂鹛蠓蟊蟆蟒螨蝥蜢蛑虻篾蠛颟耱瘼';
	v_Charn   Varchar2(2000) := '誽镎鎿雫肭捺豽軜貀鈉靹魶艿迺釢柰萘鼐錼囡喃遖楠諵難赧腩囔鬞馕曩攮齉孬呶硇铙猱譊鐃垴瑙閙鬧讷餒鮾鯘嗯鈪銰坭怩郳铌猊跜鈮貎輗鲵鯢麑齯伲旎鉨隬鑈迡昵睨鲇鮎鲶鵇鯰辇輦蹨躎廿埝醸釀茑袅鳥嬲脲肀陧臬隉嗫鉩踂踗踙錜蹑鎳闑蘖齧讘躡鑷顳钀咛鑏鬡鸋佞甯妞忸狃鈕靵侬哝農辳醲齈譨鎒鐞譳孥驽弩胬钕釹恧衄黁郍傩喏逽搦锘諾蹃鍩黏颞聍耨衲蝻蛲';
	v_Charo   Varchar2(2000) := '噢鞰讴瓯鴎謳鏂鷗齵怄耦';
	v_Charp   Varchar2(2000) := '葩杷俳輫哌蒎鎃爿跘蹒蹣鎜鞶泮頖鋬鵥鑻雱滂霶逄鳑龎龐鰟脬庖狍匏軳鞄麅麭醅阫陫锫賠錇帔旆辔霈轡湓怦軯閛嘭堋輣錋韸韼鵬鬔鑝踫闏丕纰邳铍豾釽鈚鈹鉟銔噼錃錍魾闢阰芘枇郫陴埤豼鲏罴隦魮鮍貔鵧鼙庀仳圮銢諀鴄擗淠媲睥甓鷿鸊犏翩鶣骈胼賆諚蹁谝貵諞魸剽缥飃飄魒闝殍瞟醥顠嘌嫖氕丿苤鐅姘貧嫔頻顰榀牝娉俜頩郱枰軿鲆輧鮃钋釙酦醱鏺鄱謈叵钷鉕珀頗颒掊裒攴攵陠噗鋪鯆匍酺璞濮镤贌鏷溥氆諩镨譜蹼鐠皤疋襻螃蟛筢笸蟠螵蜱蚍颦袢癖疱';
	v_Charq   Varchar2(2000) := '迉桤郪萋嘁槭諆踦諿霋蹊魌鏚鶈亓圻岐芪耆淇萁跂軝釮骐琦琪祺﨑锜頎鬾鬿綦齊蕲踑錡鲯鳍鯕鵸鶀麒鬐魕鰭麡邔屺芑杞豈绮綮諬闙汔荠葺碛憩葜跒酠鞐髂阡芊佥岍悭谸釺鈆雃愆鉛骞鹐搴諐遷褰謙顅鏲鵮鐱鬜鬝韆荨钤掮軡鈐鉆鉗銭錢鎆黚鰬凵肷慊缱譴鑓芡茜倩椠輤戕戗跄锖锵錆蹌镪蹡鎗鏘鏹嫱樯謒羟炝硗郻鄗跷鄡鄥劁踍頝缲鍫鍬趬蹺蹻鐰荞谯憔鞒樵譙趫鐈鞽顦釥愀诮陗誚韒鞩躈妾挈惬锲魥踥鍥鯜鐑衾誛顉鮼芩鈙雂嗪溱靲噙鳹檎赺赾锓鋟吣揿靑郬圊軽輕鲭鯖鑋檠黥苘頃請謦靘磬跫銎邛茕赹楸鹙趥鳅鞦鞧鰌鰍鶖鱃龝犰俅逎逑釚赇釻巯遒裘賕銶醔鮂鼽鯄鵭鰽糗岖诎阹祛誳麹魼趨麯軀麴黢鰸鱋劬朐軥蕖磲鴝璩鼩蘧氍衢躣鑺鸜齲迲郥阒觑閴麮闃鼁悛鐉诠荃辁铨跧輇銓踡闎鳈鬈鰁齤顴犭畎绻韏悫阕阙趞闋闕鵲逡鸲蝤蜷蜞蜻蜣蛴蛐蛩蚯箝箧箐筌筇罄蠼螓虬虔颀覃襁穹癯';
	v_Charr   Varchar2(2000) := '髯苒禳躟鬤譲讓荛桡娆隢遶亻鈓魜銋鵀荏稔躵仞讱轫饪恁軔葚靭靱韌飪認餁辸陾釰鈤肜狨嵘榕镕鎔軵糅蹂輮鍒鞣鰇鶔韖邚铷銣鴑嚅濡薷鴽醹顬鱬鄏込洳溽缛蓐鳰朊軟輭蕤芮枘睿銳鋭閏閠偌鄀鰙鰯鶸穰箬蝾蚺蚋颥衽襦';
	v_Chars   Varchar2(2000) := '仨靸卅钑飒脎鈒隡颯噻顋鰓賽毵鬖糁馓鏒閐搡磉鎟顙缫臊鳋颾鰠鱢埽啬铯雭銫轖鏼譅飋鬙閪铩裟魦鲨閷鎩鯊鯋唼歃閯霎彡邖芟姗钐埏舢軕釤閊跚潸膻鯅陝閃讪剡赸銏骟鄯嬗謆譱贍鐥鳝鱓鱔殇觞熵謪鬺垧賞鑜绱艄輎颵鮹苕劭潲猞畲輋賒賖佘厍滠韘麝诜鲹鯓鵢鯵鰺鉮鰰邥哂矧谂谉渖諗頣魫讅胂椹鋠阩陞陹﨡鉎鍟鼪鵿渑譝鱦眚晟貹嵊賸邿鸤釶蓍鉇酾鳲鳾鲺鍦鯴鰤鶳釃饣辻飠炻埘莳遈鉐鉽鲥鮖鼫識鼭鰣豕鉂礻贳轼铈釈弑谥貰軾鈰鉃飾適銴諟諡遾餝謚釋鰘齛扌艏狩绶鏉殳纾陎姝倏菽軗鄃摅毹跾踈輸鮛鵨秫塾贖鼡鱪鸀鱰沭腧鉥澍豎錰鏣鶐鶑唰誜闩閂涮﨎雙孀鷞鹴鸘鏯誰氵閖順鬊說説妁铄嗍搠蒴槊鎙鑠厶纟咝缌鉰飔厮銯锶澌鋖鍶颸鐁鷥鼶汜兕姒祀泗驷俟飤釲貄鈻飼忪凇崧淞菘嵩悚頌誦鎹鄋嗖溲馊飕锼醙鎪颼叟嗾瞍薮稣鯂夙涑谡嗉愫遡鹔蔌觫趚遬鋉餗謖蹜鱐鷫狻荽眭睢濉鞖雖遀隨谇誶賥燧邃鐆譢鐩狲荪飧飱隼榫鎨鶽娑挲桫睃嗦羧趖鮻唢鎍鎖鎻鎼鏁逤穑鸶疝痧筮笥笙舐蟮蟀螫螋蛸簌筲蜃蛳颡耜竦瘙';
	v_Chart   Varchar2(2000) := '趿铊溻鉈蹹鮙鳎鰨闼遝遢阘榻誻錔鞜闒鞳闥譶躢骀邰炱跆鲐颱鮐薹肽钛鈦貪昙郯锬談醈錟顃譚貚醰譠鷤忐钽鉭醓赕賧铴羰镗蹚鏜鐋鞺鼞饧鄌溏隚瑭樘踼赯醣鎕闛鶶帑傥镋鎲钂韬飸謟鞱韜饕迯洮啕鞀醄鞉鋾錭鼗忑忒貣铽慝鋱鼟滕邆謄鰧霯銻鷈鷉绨缇遆趧醍謕蹏鍗鳀鴺題鮷鵜鯷鶗鶙躰軆倜悌逖逷鐟趯酟靔黇靝畋阗鴫闐鷆鷏忝殄餂賟錪靦掭佻祧龆鋚鞗髫鲦鯈鎥齠鰷誂粜铫趒頫萜貼跕鉄銕鴩鐡鐢鐵飻餮町鞓邒莛婷葶閮霆諪鼮梃铤颋誔鋌頲嗵仝佟茼砼赨鉖僮鉵銅餇鲖潼鮦恸鍮亠骰頭黈鋵鵚鼵荼鈯跿酴鍎鵌鶟鷋鷵钍釷迌堍菟鵵貒抟鏄鷒鷻疃彖隤頹頺頽魋蹪蹆煺暾黗饨豘豚軘飩鲀魨霕氽乇讬飥魠佗陁坨沱迱柁砣跎酡踻橐鮀鴕鼧鼍鼉庹鵎鰖柝跅鹈窕箨笤螳螗蜩蜓耥裼';
	v_Charw   Varchar2(2000) := '娲鼃佤邷腽韈韤崴顡剜纨芄貦頑邜莞绾脘菀琬畹輓踠鋔鍐鋄錽贃鎫贎罔惘辋誷輞魍迋偎逶隇隈葳煨薇鳂鰃鰄囗圩帏沩闱韋涠帷嵬違鄬醀鍏闈鮠霺霻炜玮洧娓诿隗猥艉韪鲔諉踓韑頠鍡鮪韙颹韡軎猬謂錗鮇轊鏏霨鳚讆躗讏躛辒豱輼轀鳁鎾鰛鰮阌鈫雯魰鳼鴍閺閿闅鼤闦闧刎汶顐璺鹟鎓鶲蓊蕹齆倭莴喔踒肟幄渥硪龌齷圬邬趶釫鄔誈誣鴮鎢鰞郚唔浯鹀鵐鯃鼯鷡仵妩庑忤怃迕牾錻鵡躌兀兀阢杌芴逜焐婺隖靰骛寤誤鋈霚鼿霧齀鶩鹉蜿蜈痿痦鹜';
	v_Charx   Varchar2(2000) := '兮诶郗唏奚浠欷淅菥赥釸粞翕舾鄎僖誒豨餏嬉餙樨歙熹羲錫谿豀豯貕雟鯑鵗譆醯鏭隵曦酅鼷鸂鑴郋觋趘隰謵鎴霫鳛飁鰼玺徙葸鈢屣蓰銑禧諰謑蹝鱚躧饩郄郤釳阋舄趇禊赩隟黖鬩闟霼呷谺閕颬鰕狎柙陜硖陿遐瑕赮魻轄鍜鎋黠鶷閜諕鏬氙祆籼莶铦跹酰銛暹韯鍁鍂韱鮮蹮譣鶱躚鱻娴閑銜誸賢諴輱醎鹹贒鑦鷳鷴鷼冼猃険赻跣險藓鍌燹顕韅顯岘苋陥誢鋧錎豏麲鏾霰鼸芗郷鄉鄊缃葙鄕骧麘鱜鑲庠跭饷飨銄餉鲞鮝鯗響鱶項鐌鱌枭哓枵骁绡逍鸮潇踃銷魈鴞謞鴵鷍郩崤誵謏誟﨧偕勰撷缬諧鞵鐷讗龤绁亵渫榍榭韰廨獬薤邂燮謝鞢瀣齘齥齂躠躞忄邤昕莘鈊歆鋅馨鑫鬵鐔阠囟軐顖釁謃鮏鯹陉郉钘陘硎铏鈃鉶銒鋞擤荇悻﨨芎讻诇咻庥貅馐銝髹鎀鮴鵂鏅飍岫溴銹鏥鏽齅盱砉顼谞須頊魆諝譃魖鑐鬚诩栩鄦糈醑洫勖溆煦賉銊鱮蓿軒谖揎萱暄煊儇諠諼鍹譞鰚讂漩璇選泫炫铉渲楦鉉碹镟鞙颴鏇贙辥鞾泶鸴踅雤鷽轌鳕鱈谑趐謔埙獯薰曛醺峋恂洵浔荀鄩鲟鱏鱘徇迿巽遜賐蕈顨鑂皙箫筱筅罅蟓蟋螅蜥蚬胥穸痫痃鹇鸺';
	v_Chary   Varchar2(2000) := '压桠铔鴉錏鴨鵶鐚伢岈琊睚齖迓垭娅砑氩揠齾恹胭崦菸湮腌鄢嫣醃閹黫讠闫妍芫郔閆閻檐顏顔麙鹽麣兖俨偃厣郾酓琰遃隒罨魇躽黡鰋鶠黤齞龑黬黭顩鼴魘鼹齴黶晏隁焱滟鳫酽谳餍鴈諺赝鬳鴳酀贋軅醶鷃贗贘讌醼鷰釅讞豓豔泱鉠雵鞅鍈鴦阦炀钖飏徉烊陽諹輰鍚鴹颺鐊鰑霷鸉軮養怏恙幺夭吆鴁爻肴轺珧軺徭遙銚飖餆餚繇謠謡鎐鳐颻顤鰩杳崾鴢闄齩鷕靿鼼曜鷂讑鑰揶铘釾鋣鎁邺頁晔烨谒鄓鄴靥謁鍱鎑鵺靨鸈辷咿猗郼欹漪銥噫鹥醫黟譩鷖黳圯诒怡迤饴咦荑贻迻眙酏貽誃跠頉飴遺頤頥嶷顊鮧謻鏔讉鸃迆钇苡舣釔逘鈘鉯鳦旖輢顗轙齮弋刈仡阣佚呓佾峄怿驿奕弈羿轶悒挹貤陭埸豙豛釴隿跇軼鈠缢靾熠誼镒鹝鹢黓劓殪薏翳貖鮨贀鎰镱豷霬鯣鶂鶃鶍譯議醳醷鐿鷁鷊懿鷧鷾讛齸阥洇氤陰铟陻隂喑堙銦鞇諲霒闉霠韾垠狺鈝龂鄞夤誾銀龈霪齗齦鷣廴吲釿鈏飲隠靷飮趛隱讔茚胤酳鮣莺瑛锳嘤撄賏璎霙鴬膺韺鎣鶧譻鶯鑍鷪軈鷹鸎鸚茔荥萦楹滢蓥潆嬴謍瀛贏郢颕頴鐛媵鞕譍唷邕鄘墉慵銿壅郺镛雝鏞鳙饔鱅鷛喁颙顒鰫俑鲬踴鯒醟攸呦麀鄾尢柚莜莸逌郵逰遊鱿猷鈾鲉輏魷輶鮋邎卣莠铕銪牖黝侑囿宥迶貁酭誘鼬纡迃陓邘妤欤於臾禺舁狳谀酑馀萸釪隃雩魚嵛揄腴瑜觎諛雓餘魣踰輿鍝謣鮽鯲鰅鷠鸆伛俣圄圉庾鄅铻語鋙龉貐麌齬聿妪饫昱钰﨏谕逳阈飫煜蓣鈺預毓輍銉隩遹鋊鳿燠諭錥閾鴥鴧鴪魊醧鵒譽轝鐭霱鬻鱊鷸鸒軉鬰鬱眢鹓鳶鋺鴛鵷鼘鼝贠邧沅爰貟酛鈨鼋塬魭圜橼謜轅黿鎱邍鶢鶰逺遠垸媛掾瑗願刖軏钺跀鈅鉞閱閲樾龠瀹黦躍鸑龥鸙赟頵贇纭芸昀鄖雲氲鋆阭狁殒鈗隕霣齫齳郓恽鄆愠運韫熨賱醖醞韗韞韻甬鹦痖瘀螈蝣蝓蜴蜮蛘蚴蚰蚓颍窳箢筵竽罂窨窬窈翊癔瘾瘿瘗瘐痍疣鹬鹞鹆鸢';
	v_Charz   Varchar2(2000) := '卮仄赜仉伫侏倬偬俎冢诏诤诹诼谘谪谮谵阼陟陬郅邾鄣鄹圳埴芷苎茱荮菹蓁蕞奘拶揸搌摭摺撙擢攥吒咂咤哳唣唑啧啭啁帙帻幛峥崽嵫嶂徵獐馔忮怍惴浈洙浞渚涿潴濯迮彘咫姊妯嫜孳驵驺骓骘纣绉缁缒缜缯缵甾璋瓒杼栉柘枳栀桎桢梓棹楂榛槠橥樽轵轸轾辄辎臧甑昃昝贽赀赈肫胄胙胗胝朕腙膣旃炷祉祚祗祯禚恣斫砟砦碡磔黹眦畛罾钊钲铢铮锃锱镞镯锺雉秭稹鸩鸷鹧痄疰痣瘃瘵窀褶耔颛蚱蛭蜇螽蟑竺笊笫笮筝箦箸箴簪籀舯舳舴粢粽糌翥絷趱赭酎酯跖踬踯踵躅躜豸觜觯訾龇錾鲰鲻鳟髭麈齄';

Begin
	If v_Instr Is Null Or Length(Ltrim(v_Instr)) = 0 Then
		v_Spell := '';
	Else
		v_Input := Upper(v_Instr);
		v_Spell := '';
		For v_Bitnum In 1 .. Length(v_Input) Loop
			v_Bitchar := Substr(v_Input, v_Bitnum, 1);
			If v_Bitchar >= '啊' And v_Bitchar <= '座' Then
				For v_Chrnum In 1 .. Length(v_Stdstr) Loop
					If Substr(v_Stdstr, v_Chrnum, 1) = '-' Then
						Null;
					Elsif v_Bitchar < Substr(v_Stdstr, v_Chrnum, 1) Then
						v_Spell := v_Spell || Chr(64 + v_Chrnum);
						Exit;
					End If;
				End Loop;
				If v_Bitchar >= '匝' Then
					v_Spell := v_Spell || 'Z';
				End If;
			Elsif Instr('ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.+-*/', v_Bitchar) > 0 Then
				v_Spell := v_Spell || v_Bitchar;
			Elsif Instr('ⅠⅡⅢⅣⅤⅥⅧⅧⅨ', v_Bitchar) > 0 Then
				v_Spell := v_Spell || Chr(Ascii(v_Bitchar) - 41664);
			Elsif Instr('ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ',v_Bitchar) > 0 Then
				v_Spell := v_Spell || Chr(Ascii(v_Bitchar) - 41856);
			Elsif Instr('Αα', v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'A';
			Elsif Instr('Ββ', v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'B';
			Elsif Instr('Γγ', v_Bitchar) > 0 Then
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
	v_a       Varchar2(1200) := '蔼艾鞍芭茇菝靶蒡苞葆蓓鞴苯荸荜萆蓖蔽薜鞭匾苄菠薄菜蔡苍藏艹草茬茶蒇菖苌臣茌茺莼茈茨苁葱蔟萃靼鞑甙萏荡菪荻蒂东鸫董蔸芏莪苊萼蒽贰藩蕃蘩范匚芳菲匪芬葑芙芾苻茯莩菔甘苷藁戈革葛茛工功攻恭廾巩汞共贡鞲苟菇菰鹳匦邯菡蒿薅荷菏蘅薨荭蕻葫花划萑荒黄茴荟蕙荤劐或获惑藿芨基蒺蕺芰蓟葭荚菅蒹鞯茧荐茳蒋匠艽茭蕉节戒芥藉堇荩靳觐茎荆菁警敬苴鞠鞫菊莒巨苣蕨菌蒈勘戡莰苛恐芤蔻苦蒯匡葵匮蒉莱蓝莨蒗劳勒蕾莉蓠藜苈荔莅莲蔹蓼蔺苓菱茏蒌芦萝荦落荬颟鞔蔓芒茫莽茅茆茂莓萌甍瞢蒙蘼苗鹋藐蔑苠茗摹蘑茉莫蓦某苜募墓幕慕暮艿萘匿廿茑孽蘖欧殴瓯鸥藕葩蒎蓬芘匹苤苹萍叵莆菩葡蒲七萋期欺芪其萁綦蕲芑荠葺葜芊荨芡茜蔷荞鞒巧翘鞘切茄芩芹勤擎檠苘跫銎邛茕蛩区蕖蘧荃颧鹊苒荛惹荏葚戎茸荣蓉鞣茹薷蓐蕤蕊芮若萨散莎芟苫芍苕甚蓍莳世式贳菽蔬薯蒴斯菘薮苏蔌蒜荽荪蓑苔薹萄忒慝藤萜莛葶茼荼菟芄莞菀葳薇苇萎蔚蓊蕹莴卧巫芜芴昔菥熙觋葸蓰匣莶藓苋芗葙巷项萧邪鞋薤芯莘薪荇芎蓄蓿萱靴薛薰荀蕈鸦牙芽雅迓菸蔫芫郾燕鞅尧药医荑颐苡弋艺薏翳茵荫鄞茚英莺茔荥荧莹萤营萦蓥莜莸莠萸芋蓣鸢苑芸蕴匝葬藻赜蘸蔗斟蓁蒸芝芷荮茱苎著茁菹蕞';
	v_b       Varchar2(1200) := '阿隘阪孢陂陛屮陈丞承蚩耻出除陲聪耽聃阽耵陡队堕耳防阝附陔隔耿孤聒孩函隍隳亟际降阶卩孑卺阱聚孓孔聩联辽聊了陵聆隆陇陋陆勐孟陌陧聂颞聍陪陴聘阡凵取娶孺阮陕隋随祟隧孙陶粜陀隈隗卫阢隰隙险限陷陉逊阳耶也阴隐隅院陨障阵职陟骘坠孜子陬鄹阻阼';
	v_c       Varchar2(1200) := '巴畚弁骠驳参骖叉骋驰骢皴迨怠邓叠对怼驸观骇骅欢鸡骥艰骄矜劲刭颈迳驹骏骒垒骊骝驴骡骆马矛蝥蟊瞀牟鍪难能骈骗骐骑巯驱劝逡柔叁毵桑颡骚骟圣驶双厶驷骀台邰炱通驮驼婺骛鹜戏骧骁熊驯验以矣驿甬勇恿又予驭预豫鹬允驵蚤骣骤驻骓驺';
	v_d       Varchar2(1200) := '砹碍鹌百邦帮磅悲碑辈碚奔泵砭碥髟飙鬓礴布礤厕碴虿厂耖砗辰碜成舂厨础春唇蠢磁蹙存磋厝耷达大砀焘磴砥碲碘碉碟碇硐碓礅趸砘夺厄而鸸砝矾非蜚斐翡奋丰砜酆奉砩尬尴感矸硌耕龚鸪辜古嘏故顾硅磙夯耗厚胡鹕瓠鬟磺灰彗慧耠矶剞髻恝戛硷碱礓耩礁碣兢鬏韭厩厥劂砍磕克刳夸夼矿盔奎髡砬耢耒磊厘历厉励砺砾奁鹩尥鬣磷硫龙砻聋垄耧碌码劢迈硭髦礞面耱奈耐硇碾耨恧耦耙耪匏裴砰硼碰砒破戚奇契砌碛牵硗挈秦磲鬈犬确髯辱三磉砂奢厍砷蜃盛石寿戍耍爽硕厮耜肆碎太态泰碳耥套髫厅砼砣碗万威硪戊矽硒袭硖夏厦咸厢硝硎雄髹戌砉碹压砑研奄厣魇厌砚艳雁餍赝页靥欹硬尢尤友有右郁原愿耘砸在臧仄砟丈磔砧碡砖斫髭耔鬃奏左';
	v_e       Varchar2(1200) := '爱胺肮膀胞豹膘豳膑脖膊采彩豺肠塍腠脆脞胆貂腚胨胴肚腭肪肥腓肺肤孚服郛脯腑腹尕戤肝肛胳膈肱股臌胍胱虢胲貉肌及胛腱胶脚腈肼胫雎爵胩胯脍腊肋臁脸膦胧胪脶脉貌朦脒觅腼邈膜貊貘肭乃鼐腩脑腻脲脓胖脬胚朋鹏膨脾貔胼脐肷腔且朐肜乳朊脎腮臊彡膻膳胂胜豕受腧甩舜胎肽膛腾滕腆腿豚脱妥腽脘腕肟奚膝燹县腺胁腥胸貅须悬胭腌腰遥繇舀鹞腋胰臆盈媵臃用有腴爰月刖孕脏膪胀胗朕肢胝脂豸膣肿肘逐助肫腙胙';
	v_f       Varchar2(1200) := '埃霭埯坳坝霸坂雹贲甏孛勃博鹁埠才裁场超朝坼趁城埕墀赤翅亍矗寸埭戴堤觌坻地颠坫垤堞耋动垌都堵堆墩垛二坊霏坟封夫赴垓干坩赶圪塥埂垢彀遘觏毂鼓瞽卦圭规埚过顸邗韩翰壕郝盍赫堠壶觳坏卉恚魂霍击圾赍吉戟霁嘉教劫颉截进井境赳救趄均垲刊堪坎考坷壳坑堀垮块款圹亏逵坤垃老雷塄嫠坜雳墚埒趔霖零酃垅露垆埋霾卖墁耄霉坶南赧垴坭霓辇埝培霈堋彭坯霹埤鼙圮坪坡埔亓圻耆起乾墙謦磬罄求逑裘趋麴去趣却悫壤韧颥丧埽啬霎埏墒垧赦声十埘士示螫霜寺索塌塔坍坛坦塘趟韬替填霆土堍坨顽韦圩违未雯斡圬无坞雾熹喜霞献霰霄孝协馨幸需墟雪埙垭盐堰壹圯埸懿堙垠霪墉雩雨域元垣袁鼋塬远垸越云运韫哉栽载趱增赵者赭真圳震支直埴址志煮翥专趑走';
	v_g       Varchar2(1200) := '瑷敖獒遨熬聱螯鳌骜鏊班斑甭逼碧表殡丙邴玻逋不残蚕璨曹琛豉敕刺璁琮殂璀歹带殆玳殚到纛玷靛玎豆逗毒蠹顿垩恶噩珥珐玢否麸敷甫副丐鬲亘更珙瑰翮珩瑚琥互画还环璜虺珲惠丌玑墼棘殛夹珈郏颊戋歼柬戬豇瑾晋靓静玖琚珏开珂琨剌来赉赖琅鹂璃逦理吏丽郦琏殓两列烈裂琳玲琉珑璐珞玛麦瑁玫芈灭玟珉末殁囊孬瑙弄琶丕邳琵殍平珀璞妻琦琪琴青琼球璩融瑞卅瑟珊殇事殊束死素速琐瑭忑天忝殄餮吞屯橐瓦歪豌玩琬王玮軎吾五武鹉兀瑕下现刑邢形型顼璇殉琊亚焉鄢严琰殃珧瑶一夷殪瑛璎迂于欤盂瑜与玉瑗殒再瓒遭枣责盏璋珍臻整正政殖至郅致珠赘琢';
	v_h       Varchar2(1200) := '龅彪卜步睬餐粲柴觇龀瞠眵齿瞅龊雌此鹾眈瞪睇点盯鼎督睹盹丨壑虍虎乩睑睫睛旧龃具遽瞿矍卡瞰瞌肯眍眶睽睐瞵龄卢鸬颅卤虏虑瞒眯眠眄瞄眇瞑眸目睦睨虐盼皮睥瞟频颦颇攴歧虔瞧氍龋觑睿上叔睡瞬瞍眭睢睃忐龆眺瞳凸龌瞎些盱虚眩睚眼眙龈卣虞龉眨砦瞻占战贞睁止瞩卓桌赀觜龇紫訾眦';
	v_i       Varchar2(1200) := '澳灞浜滗濞汴滨濒波泊渤不沧漕测涔汊潺尝常敞氅潮澈尘沉澄池滁淳淙淬沓淡澹当党凼滴涤滇淀洞渎渡沌沲洱法泛淝沸汾瀵沣浮涪滏尜溉泔澉淦港沟沽汩涫灌光滚海涵汉汗瀚沆濠浩灏河涸泓洪鸿黉鲎滹湖浒沪滑淮洹浣涣漶湟潢辉洄汇浍浑混溷活激汲脊洎济浃尖湔涧渐溅江洚浇湫洁津浸泾酒沮举涓觉浚渴溘喾溃涞濑澜漤滥浪潦涝泐泪漓澧沥溧涟濂潋梁粱劣洌淋泠溜浏流鎏泷漏泸渌滤漉潞滦沦泺洛漯满漫漭泖没湄浼懑汨泌沔湎淼渺泯溟沫漠沐淖泥溺涅泞浓沤派湃潘泮滂泡沛湓澎淠漂泼婆濮浦溥瀑沏柒漆淇汔汽泣洽潜浅溱沁清泅渠雀染溶濡汝洳溽润洒涩沙裟鲨潸汕裳赏尚少潲涉滠深沈渖渗渑省湿淑沭漱澍涮氵水澌汜泗淞溲涑溯濉娑挲溻汰滩潭汤堂棠溏淌烫涛滔洮逃淘鼗涕添汀潼涂湍沱洼湾汪沩涠潍洧渭温汶涡沃渥污浯鋈汐浠淅溪洗涎湘削消逍潇淆小肖泄泻渫瀣兴汹溴洫溆漩泫渲学泶洵浔汛涯淹湮沿演滟泱洋漾耀液漪沂溢洇淫滢潆瀛泳涌油游淤渝渔浴誉渊沅源瀹澡泽渣沾澶湛漳涨掌沼兆浙浈汁治滞洲洙潴渚注涿浊浞濯淄滋滓渍';
	v_j       Varchar2(1200) := '暧暗昂蚌暴蝙晡螬蝉蟾昌畅晁晨蛏螭匙虫蜍蝽旦刂戥电蝶蚪蛾遏蜂蚨蜉蝠蝮旰杲蛤虼蚣蛄蛊归晷炅蝈果蜾蚶晗旱蚝昊颢曷蚵虹蝴蝗蟥晃晖蛔晦蟪夥蠖虮蛱坚监鉴蛟蚧紧晶景颗蝌旷暌蝰昆蛞旯蜡览螂蜊里蛎蠊晾量临蛉蝼螺蟆蚂螨曼蟒蛑昴冒昧虻盟蜢蠓冕蠛明暝螟蝻曩蛲昵暖蟠螃蟛蚍蜱螵曝蛴蜞蜣螓蜻晴蚯虬蝤蛆蛐蠼蜷蚺日蝾蠕蚋晒蟮晌蛸蛇申肾晟师时是暑曙竖墅帅蟀蛳螋遢昙螗螳剔题蜩蜓蜕暾蛙蜿晚旺韪蚊蜗蜈晤晰蜥螅蟋曦虾暇暹贤显蚬蟓晓歇蝎昕星勖煦暄曛蚜蜒晏蛘曜野曳晔蚁易蜴蚓蝇影映蛹蚰蝣蚴禺愚蝓昱遇蜮螈曰昀晕早昃蚱蟑昭照蜘蛭蛛蛀最昨';
	v_k       Varchar2(1200) := '吖啊嗄哎唉嗳嗌嗷叭吧跋呗趵嘣蹦吡鄙哔跸别啵踣跛卟哺嚓踩嘈噌蹭躔唱嘲吵嗔呈逞吃哧嗤踟叱踌躇蹰啜嘬踹川喘串吹踔呲蹴蹿啐蹉哒嗒呆呔啖叨蹈噔蹬嘀嗲踮叼吊跌喋蹀叮啶咚嘟吨蹲咄哆踱哚跺呃鄂鹗颚蹯啡吠吩唪呋趺跗呒咐嘎噶嗝跟哏哽咕呱剐咣贵跪呙哈嗨喊嚆嗥嚎号呵喝嗬嘿哼哄喉吼呼唿唬哗踝唤患咴哕喙嚯叽咭唧跻戢哜跽跏趼践踺跤叫噍喈嗟噤啾咀踽距踞鹃噘噱蹶嚼咔咖喀咳嗑啃吭口叩哭跨哙哐喹跬喟啦喇啷唠叻嘞喱哩呖唳跞踉嘹咧躐啉躏另呤咯咙喽噜路鹭吕骂唛吗嘛咪嘧黾喵咩鸣哞哪喃囔呶呐呢嗯啮嗫蹑咛哝喏噢哦呕趴啪哌蹒咆跑呸喷嘭噼啤蹁嘌品噗蹼嘁蹊器遣呛跄跷嗪噙吣嚷蹂嚅噻嗓唼啥跚哨呻哂史嗜噬唰吮顺嗍嗽咝嘶嗣嗖嗾嗉虽唆嗦唢趿踏蹋跆叹饕啕踢啼蹄嚏跳听嗵吐跎鼍唾哇唯味喂吻嗡喔呜吴唔吸唏嘻呷吓跹跣响哓嚣哮啸躞兄咻嘘嗅喧勋呀哑咽唁吆咬噎叶咿噫咦遗呓邑喑吟吲嘤郢哟唷喁咏踊呦吁喻员跃郧咂咱唣噪躁啧吒咋哳喳咤啁吱跖踯只趾踬中忠盅踵咒躅嘱啭啄踪足躜嘴唑';
	v_l       Varchar2(1200) := '黯罢办畀边黪车畴黜辍辏黩囤轭恩罚畈罘辐辅罡哿轱罟固轨辊国贺黑轰轷囫回畸羁辑加迦袈甲驾架囝轿较界轲困罱累罹力轹詈连辆辚囹轳辂辘略囵轮罗逻皿墨默囡男嬲畔毗罴圃畦黔堑椠轻圊黥囚黢圈辁畎轫软轼输署蜀思四田畋町图团疃畹辋囗围畏胃辖黠勰轩鸭罨轺黟轶因黝囿圄圉园圆辕圜暂錾罾轧斩辗罩辄辙轸畛轵轾置轴转辎罪';
	v_m       Varchar2(1200) := '岸盎凹岜败贝崩髀贬飑飚髌财册岑崇帱遄赐崔嵯丹嶝迪骶巅典雕岽峒髑赌朵剁峨帆幡凡贩风峰凤幅幞赋赙赅冈刚岗骼岣购鹘骨崮刿崞帼骸骺岵凰幌贿岌几嵴觊岬见贱峤骱巾赆冂迥飓崛峻凯剀髁岢崆骷髋贶岿崃岚崂嶙岭髅嵝赂幔峁帽嵋岷内帕赔帔岐崎屺岂髂岍嵌峭赇曲岖冉嵘肉山删赡赊嵊殳赎兕崧嵩飕夙髓岁炭赕贴帖同彤骰崴网罔巍帏帷嵬幄峡岘崤岫峋岈崖崦岩央鸯崾贻嶷屹峄婴罂鹦由邮嵛屿峪崽赃则帻贼赠崭帐账嶂幛赈峥帧帙帜峙周胄贮颛赚幢嵫';
	v_n       Varchar2(1200) := '懊悖鐾必愎辟壁嬖避臂璧襞忭擘檗怖惭惨恻层孱忏羼惝怅怊忱迟尺忡憧惆丑怵憷怆戳悴翠忖怛惮蛋忉导悼惦殿刁懂恫惰屙愕发飞悱愤怫改敢怪惯憨悍憾恨恒惚怙怀慌惶恍恢悔屐己忌悸届尽惊憬居局剧惧屦恺慨忾慷尻恪快悝愦愧悃懒愣怜懔鹨戮屡履买慢忙眉鹛懵乜民悯愍恼尼怩尿忸懦怄怕爿怦劈屁甓譬屏恰悭慊戕悄憔愀怯惬情屈悛慑慎尸虱屎恃收书疏刷司巳忪悚愫屉悌惕恬恸屠臀惋惘惟尾尉慰屋忤怃悟惜犀习屣遐屑懈忄心忻惺性悻胥恤恂迅巽疋恹怏怡乙已以忆异怿羿悒翌翼慵忧愉羽悦恽愠熨奘憎翟展怔咫忮昼属惴怍';
	v_o       Varchar2(1200) := '粑爆焙煸灬炳灿糙焯炒炽炊糍粗粹灯断煅炖烦燔粉粪烽黻黼糕焓焊烘糇烀煳糊焕煌烩火糨烬粳精炯炬爝糠炕烤烂烙类粒粝炼粮燎料粼遴熘娄炉熳煤焖迷米敉糯炮粕炝糗炔燃熔糅糁煽剡熵烧炻数烁燧郯糖烃煺烷煨为炜焐烯粞熄籼燮糈煊炫烟炎焰焱炀烊业邺烨熠煜燠糌糟凿灶燥炸粘黹烛炷灼籽粽';
	v_p       Varchar2(1200) := '安案袄宝褓被褙裨窆褊裱宾补察衩禅宸衬裎褫宠初褚穿窗辶祠窜褡裆宕祷定窦裰额祓袱福富袼割宫寡褂官冠宄害寒罕褐鹤宏祜寰宦逭豁祸寂寄家袷裥謇蹇窖衿襟窘究裾窭军皲窠客裉空寇窟裤宽窥褴牢礼帘裢裣寥寮窿禄褛裸袂寐祢冖宓密幂蜜宀冥寞衲宁甯农袢襻袍祁祈祺骞搴褰襁窍窃寝穷穹祛裙禳衽容冗襦褥塞赛衫社神审实礻视室守祀宋宿邃它袒裼祧窕突褪袜剜完宛窝寤穸禧禊祆宪祥宵写袖宣穴窨宴窑窈衤宜寅廴宥窬宇窳寓裕冤郓灾宰宅窄寨褶这祯鸩之祗祉窒冢宙祝窀禚字宗祖祚';
	v_q       Varchar2(1200) := '锕锿铵犴钯鲅钣镑勹包饱鲍狈钡锛狴铋鳊镖镳鳔镔饼钵饽钹铂钸钚猜馇锸猹镲钗馋镡铲猖鲳鬯钞铛铖鸱饬铳刍锄雏触舛钏璠锤匆猝镩锉错岛锝镫镝狄氐邸甸钿鲷钓铞鲽钉锭铥兜钭独镀锻镦钝多铎锇饿锷鳄儿鲕尔迩饵铒钒犯饭钫鲂鲱狒镄鲼锋孵凫匐负鲋鳆钆钙钢镐锆镉铬鲠觥勾钩狗够觚钴锢鲴鳏馆盥犷逛龟鲑鳜鲧锅猓铪狠訇猴忽狐斛猢鹱铧猾獾郇锾奂鲩鳇昏馄锪钬镬饥急鲚鲫镓铗钾鲣锏饯键鲛角狡饺铰桀鲒解钅金锦馑鲸獍镜久灸狙锔句钜锯镌锩狷觖獗镢钧锎铠锴钪铐钶锞铿狯狂馈锟鲲铼镧狼锒铹铑乐鳓镭狸鲡锂鲤鳢猁鲢镰链獠镣钌猎鳞铃鲮留遛馏镏锍镂鲈鲁镥铝卵锊猡锣镙犸馒鳗镘猫锚卯铆贸猸镅镁钔猛锰猕免勉名铭馍镆钼镎钠馕铙猱馁铌猊鲵鲇鲶鸟袅镊镍狞狃钮钕锘刨狍锫铍鲆钋钷铺匍镤镨鳍钎铅钤钱钳欠锖锵镪锹锲钦锓卿鲭鳅犰劬鸲铨犭然饶饪狨铷锐鳃馓鳋色铯杀刹铩煞钐鳝觞勺猞狮鲺饣蚀鲥氏饰铈弑狩铄锶饲馊锼稣觫狻狲飧锁铊獭鳎鲐钛锬钽铴镗饧铽锑逖鲦铫铁铤铜钍兔饨鸵外危猥鲔猬刎我乌邬钨勿夕希郗欷锡玺铣饩郄狎狭锨鲜猃馅镶饷象枭销獬邂蟹锌鑫猩凶匈馐锈铉镟鳕獯旬鲟爻肴鳐铘猗铱饴钇刈逸镒镱铟狺银夤饮印迎镛鳙犹铀鱿铕鱼狳馀饫狱钰眢鸳猿怨钥钺匀狁锃铡詹獐钊锗针镇争狰钲铮炙觯钟锺皱猪铢橥铸馔锥镯锱鲻邹鲰镞钻鳟';
	v_r       Varchar2(1200) := '挨捱皑氨揞按翱拗扒捌拔魃把掰白捭摆拜扳搬扮拌报抱卑鹎拚摈兵摒拨播帛搏捕擦操插搽拆掺搀抄扯掣撤抻撑魑持斥抽搐搋揣氚捶撺摧搓撮挫措搭打担掸氮挡氘捣的抵掂垫掉迭瓞揲氡抖盾遁掇扼摁反返氛缶扶拂氟抚拊擀缸皋搞搁搿拱瓜挂拐掼罐皈鬼掴氦捍撖撼皓后逅护换擐皇遑挥攉挤掎技搛拣捡挢皎搅敫接揭拮捷斤近揪拘掬拒据捐撅抉掘攫捃揩看扛抗拷氪控抠扣挎揆魁捆扩括拉拦揽捞擂魉撩撂捩拎拢搂撸掳氯掠抡捋摞魅扪描抿摸抹拇捺氖攮挠拟拈年捻撵捏拧牛扭挪搦爬拍排乓抛抨捧批披郫擗氕撇拼乒皤迫魄掊扑颀气掐扦掮抢撬擒揿氢氰丘邱泉缺攘扰热扔揉撒搡搔扫擅捎摄失拾势拭逝誓手扌授抒摅摔拴搠撕搜擞损所挞抬摊探搪掏提掭挑挺捅投抟推托拖拓挖挽皖魍挝握捂舞罅氙掀魈挟携撷卸欣擤揎踅押氩揠掩扬氧邀摇揶掖揖抑挹殷氤撄拥揄援掾岳氲拶攒皂择揸扎摘搌招找蜇折哲蛰摺振挣拯卮执絷摭指制质挚贽掷鸷朱邾拄抓爪拽撰撞拙捉擢揍攥撙';
	v_s       Varchar2(1200) := '桉柏板梆榜棒杯本杓标彬槟柄醭材槽杈查槎檫郴榇柽枨酲橙酬樗橱杵楮楚椽棰槌椿醇枞楱酢醋榱村档柢棣丁酊顶栋椟杜椴樊梵枋榧酚棼焚枫桴覆概杆柑酐橄杠槔槁哥歌格根梗枸构酤梏棺桄柜桂棍椁醢酣杭核桁横槲醐桦槐桓桧机极楫枷贾枧检楗槛椒酵醮杰槿禁柩桕椐桔橘榉醵鄄桷橛楷栲柯棵可枯酷框醌栝栏婪榄榔醪栳酪檑酹棱楞李醴枥栎栗楝椋林檩柃棂榴柳栊楼栌橹麓榈椤杩懋枚梅楣酶檬梦醚棉杪酩模木柰楠酿柠杷攀醅配棚枇剽飘瓢票榀枰朴栖桤槭棋杞枪樯橇桥樵檎楸权醛榷桡榕枘森杉梢椹酾柿枢梳术述树栓松酥粟酸榫桫梭榻酞覃檀樘醣桃梯醍梃桐酮桶酴柁酡椭柝枉桅梧杌西析皙樨醯檄柙酰相想橡枵校楔械榍榭醒杏朽栩醑酗楦醺桠檐酽杨样杳要椰酏椅樱楹柚酉榆橼樾酝楂札栅榨栈樟杖棹柘桢甄榛枕枝栀植枳酯栉桎酎株槠杼柱桩椎酌梓棕醉樽柞';
	v_t       Varchar2(1200) := '矮岙奥笆稗般版舨备惫笨鼻彼秕笔舭币筚箅篦笾秉舶箔簸簿舱艚策长徜彻称乘惩程秤笞篪彳艟愁稠筹臭处舡船囱垂辞徂簇汆篡毳矬笪答待箪稻得德簦等籴敌笛第簟牒丢冬篼牍犊笃短簖躲舵鹅乏筏番翻繁彷舫篚逢稃符复馥竿秆筻睾篙稿告郜舸各躬篝笱箍牯鹄牿刮鸹乖管簋鼾航禾和很衡篌後乎笏徊徨篁簧徽秽积笄嵇犄箕稽笈籍季稷笳稼笺犍笕简舰毽箭矫徼秸街筋径咎矩榘犋筠犒靠科稞箜筘筷筐篑徕籁篮稂梨犁黎篱黧利笠篥笼篓舻簏氇稆律乱箩雒毛牦么每艨艋秘秒篾敏鳘秣毪牡牧穆黏臬衄筢徘牌盘磐逄篷片犏篇丿牝鄱笸攵氆乞迄憩千迁愆签箝乔箧箐筇秋鼽躯衢筌穰壬稔入箬穑歃筛舢稍筲艄舌射身矧升生牲笙甥眚剩矢适舐释筮艏秫黍税私笥艘簌算穗笋毯躺特甜舔条笤廷艇筒透秃徒颓乇箨往逶微委艉魏稳我午迕牾务物息牺悉稀舾徙系先舷衔筅香箱向箫筱笑囟衅行秀徐选血熏循徇衙延筵衍秧徉夭徭迤移舣役劓胤牖釉竽禹御毓箢粤簪昝赞造迮笮舴箦怎齄乍毡笊箴稹征筝徵知夂秩智稚雉舯螽种重舟籀竹竺舳筑箸篆秭笫自租纂';
	v_u       Varchar2(1200) := '癌疤瘢癍半瓣北邶背迸闭敝痹弊辨辩辫瘭憋鳖蹩瘪冫冰并病部瓿差瘥产阐冁阊闯痴啻瘛冲瘳疮疵瓷慈鹚次凑瘁痤瘩单郸瘅疸盗道羝弟帝递癫奠癜凋疔冻斗痘端兑阏阀痱疯冯盖疳赣戆羔疙阁羹痼关闺衮馘阂阖痕闳瘊冱痪豢癀阍疾瘠冀痂瘕间兼煎鹣减剪翦姜将浆奖桨酱交郊疖竭羯疥净痉竞竟靖阄疚疽蠲卷桊眷决竣阚闶疴况夔阃阔瘌辣癞兰阑阆痨冷立疠疬痢凉疗冽凛凌羚瘤六癃瘘闾瘰美门闷闵闽瘼闹疒逆凝疟判叛旁疱疲痞癖瞥瓶剖普凄前歉羌羟妾亲酋遒癯阒拳痊券瘸阕阙闰飒瘙痧闪疝善鄯商韶首兽瘦闩朔槊凇竦送塑遂羧闼瘫痰羰疼誊鹈剃阗童痛头闱痿瘟闻阌问痦羲阋闲痫鹇冼羡翔鲞效辛新歆羞痃癣丫痖阉闫阎颜兖彦羊疡养痒恙冶痍疫益翊意瘗毅癔音瘾瘿痈疣猷瘀瘐阈阅韵曾甑闸痄瘵站章鄣彰瘴疹郑症痔痣瘃疰妆装丬壮状准着兹咨姿资孳粢恣总尊遵';
	v_v       Varchar2(1200) := '嫒媪妣婢婊剥姹婵娼嫦巢媸巛妲逮刀嫡娣妒娥婀发妨妃鼢妇旮艮媾姑妫好毁婚姬即嫉彐妓既暨嫁奸建姣娇剿婕她姐妗婧鸠九臼舅娟君郡垦恳馗邋姥嫘娌隶灵录逯妈嬷媒妹媚娩妙嫫姆那娜奶嫩妮娘肀妞奴孥驽努弩胬怒女媲嫖姘嫔娉嫱群娆忍刃妊如嫂姗嬗劭邵娠婶始姝鼠恕孀妁姒叟肃帑迢婷退娃娲丸婉娓鼯妩嬉鼷媳舄娴嫌姓旭婿絮寻巡娅嫣妍鼹妖姚姨姻尹邕鼬妤臾舁娱聿妪媛杂甾嫜召妯帚姊';
	v_w       Varchar2(1200) := '俺傲八爸佰颁伴傍煲保堡倍坌俾便傧伯仓伧侧岔侪伥偿倡侈傺仇俦雠储传创从丛促爨催傣代岱贷袋黛儋但倒登凳低佃爹仃侗段俄佴伐垡仿分份忿偾俸佛伏俘斧俯釜父付阜傅伽鸽个公供佝估谷倌癸刽含颔合何盒颌侯候华化会伙货佶集伎偈祭佳价假俭件剑牮健僭僵焦僬鹪佼侥介借今仅儆僦俱倨倦隽倔俊佧龛侃伉倥侉侩郐傀佬仂儡俚例俐俪傈俩敛僚邻赁伶瓴翎领令偻侣仑伦倮们命侔仫拿倪伲你念您佞侬傩偶俳佩盆仳僻偏贫俜凭仆企仟佥倩戗劁侨俏侵衾禽倾俅全人亻仁仞任恁仍儒偌仨伞僧傻伤畲佘舍伸什食使仕侍售倏舒毹伺似俟怂耸颂俗僳隼他贪倘傥体倜佻停仝佟僮偷途氽佗佤偎伟伪位璺翁瓮倭仵伍侮兮翕僖歙侠仙像偕斜信休修鸺叙儇伢俨偃佯仰爷伊依仪倚亿仡佚佾佣俑优攸悠佑侑余俞逾觎舆伛俣欲鹆愈龠债仉仗侦侄值仲众侏伫住隹追倬仔偬俎佐作坐做';
	v_x       Varchar2(1200) := '绊绑鸨绷匕比毕毖毙弼编缏缤缠弛绸绌纯绰绐弹缔缎缍纺绯费纷缝弗绂绋艴缚绀纲缟纥给绠弓缑贯绲绗弘红弧缳缓幻绘缋绩缉畿级纪继缄缣缰疆绛犟绞缴皆结缙经弪纠绢绝缂绔纩缆缧缡蠡练缭绫绺缕绿纶络缦弥弭糸绵缅缈缗缪母纳纽辔纰缥绮缱强缲顷绻绕纫绒缛弱缫纱缮绱绍绅绳绶纾纟丝鸶缌绥缩绦绨缇统彖纨绾维纬纹毋细纤弦线乡缃飨绡缬绁绣绪续绚幺疑彝绎缢肄引缨颍颖幽幼纡鬻缘约纭缯绽张缜织旨纸彘终粥纣绉缀缒缁综纵组缵';
	v_y       Varchar2(1200) := '哀庵谙廒鏖谤褒庇庳扁卞变遍斌禀亳诧谗廛谄颤昶谌谶诚充床鹑词诞谠诋底谛店调谍订读度憝敦讹谔方邡房访放扉诽废讽府腐讣该高膏诰庚赓诟诂雇诖广庋诡郭裹亥颃毫豪诃劾亨讧户戽扈话肓谎诙麾讳诲诨讥迹齑麂计记剂肩谫谏讲讦诘诫谨廑京旌扃就鹫讵诀谲麇康亢颏刻课库诓诳邝廓谰斓郎廊朗羸诔离戾廉娈恋良亮谅廖麟廪吝刘旒庐鹿旅膂率孪峦挛栾鸾脔銮论蠃麻蛮谩邙盲旄袤氓谜糜縻麋靡谧庙谬谟麽摩磨魔谋亩讷旎诺讴庞庖旆烹庀翩谝评裒谱齐旗麒启綮讫弃谦谴敲谯诮请庆诎诠瓤让认讪扇设麝诜谂诗施识市试谥孰塾熟庶衰谁说讼诵诉谡谇谈谭唐讨亭庭亠庹弯亡妄忘望为诿谓文紊诬庑误诶席襄详庠享谐亵谢廨庥许诩序畜谖玄旋谑询训讯讶讠言谚谳谣夜谒衣诒旖义议亦译诣奕弈谊裔应膺鹰嬴赢庸雍壅饔永诱於谀语庾育谕谮诈斋旃谵诏肇遮谪这鹧诊证诤衷州诌诛诸丶主麈庄谆诼谘诹卒族诅座';
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

--公共同义词
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
--权限
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

--2005年1月6日补充
Grant Execute On zlSpellCode to Public
/
Create Public Synonym zlSpellCode For zlSpellCode
/

Grant Execute On zlWBcode to Public
/
Create Public Synonym zlWBcode For zlWBcode
/