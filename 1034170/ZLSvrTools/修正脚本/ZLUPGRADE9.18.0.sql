-------------------------------------------------------------------------------
--结构修正部份
-------------------------------------------------------------------------------
--ZLTOOLS后台作业:主键->唯一键
ALTER TABLE zlAutoJobs DROP CONSTRAINT zlAutoJobs_PK
/
ALTER TABLE zlAutoJobs ADD CONSTRAINT 
    zlAutoJobs_PK UNIQUE(系统,类型,序号)
    USING INDEX PCTFREE 5 
    STORAGE(INITIAL 256 NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/

-------------------------------------------------------------------------------
--数据修正部份
-------------------------------------------------------------------------------
--后台作业处理运行和错误日志
Insert Into zlAutoJobs(系统,类型,序号,名称,说明,内容,参数,执行时间,间隔时间) Values(Null,1,1,'日志数据自动处理','对多余的运行日志和错误日志进行清除，以及对异常的运行日志进行标记。','zl_AutoLogProcess',Null,Trunc(Sysdate)+2/24,1)
/
--7048：编译无效对象
insert into zlSvrTools(编号,上级,标题,快键,说明) values ('010301','0103','编译无效对象','P',Null)
/

-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
Create Or Replace Procedure zl_AutoLogProcess as
--功能：
--   1.对多余的运行日志和错误日志进行清除
--   2.对异常的运行日志进行标记
	v_Count		Number;
	v_Limit		Number;
Begin
	--删除多余的运行日志
	Select Count(*) Into v_Count From zlDiaryLog;
	Begin
		Select Nvl(to_Number(参数值),0) Into v_Limit From zlOptions Where 参数号=2;
	Exception
		When Others Then v_Limit:=10000;
	End;
	If v_Count>v_Limit Then
		Delete From zlDiaryLog Where RowID IN(
			Select ID From (
				Select RowID As ID From zlDiaryLog Group By 进入时间,RowID) 
			Where Rownum<v_Count-v_Limit+1);
	End If;
	
	--对异常退出的运行日志记录进行处理
	Update zlDiaryLog 
		Set 退出原因=2,退出时间=Sysdate
	Where 退出原因 is Null And 会话号 Not IN(Select SID+SERIAL# From v$Session Where USER#<>0);

	--删除多余的错误日志
	Select Count(*) Into v_Count From zlErrorLog;
	Begin
		Select Nvl(to_Number(参数值),0) Into v_Limit From zlOptions Where 参数号=4;
	Exception
		When Others Then v_Limit:=10000;
	End;
	If v_Count>v_Limit Then
		Delete From zlErrorLog Where RowID IN(
			Select ID From (
				Select RowID As ID From zlErrorLog Group By 时间,RowID) 
			Where Rownum<v_Count-v_Limit+1);
	End If;
End zl_AutoLogProcess;
/

-------------------------------------------------------
--ZL作业管理
--说明： 该组函数是管理工具的作业管理函数
--清单： 1、zl_JobSubmit：作业提交
--       2、zl_JobRemove：作业删除
--       3、zl_JobRun：作业执行
--       4、zl_JobChange：作业修改
-------------------------------------------------------
CREATE OR REPLACE PROCEDURE zl_JobSubmit (
    Job_system IN INTEGER,
    Job_kind IN INTEGER,
    Job_odd IN INTEGER
)
IS
    V_content VARCHAR2 (200);
    V_parameter VARCHAR2 (200);
    V_paraitem VARCHAR2 (200);
    V_starttime DATE;
    V_cyclekeep INTEGER;
    V_jobnum NUMBER := 0;
    V_what VARCHAR2 (1000);
    V_nextdate DATE;
    V_INterval VARCHAR2 (1000);
BEGIN
    SELECT 内容, 参数, 执行时间, 间隔时间
      INTO V_content, V_parameter, V_starttime, V_cyclekeep
      FROM Zlautojobs
     WHERE Nvl(系统,0) = Nvl(Job_system,0)
        AND 类型 = Job_kind
        AND 序号 = Job_odd;
    V_what := '';

    IF LENGTH (V_parameter) > 0 THEN
        LOOP
            IF INSTR (V_parameter, ';') > 0 THEN
                V_paraitem := SUBSTR (V_parameter, 1, INSTR (V_parameter, ';') - 1);
                V_parameter := SUBSTR (V_parameter, INSTR (V_parameter, ';') + 1);
            ELSE
                V_paraitem := V_parameter;
            END IF;

            V_what :=V_what || ',' || SUBSTR (V_paraitem, INSTR (V_paraitem, ',') + 1);
            EXIT WHEN INSTR (V_parameter, ';') = 0;
        END LOOP;
    END IF;

    IF LENGTH (V_what) <> 0 THEN
        V_what := V_content || '(' || SUBSTR (V_what, 2) || ');';
    ELSE
        V_what := V_content || ';';
    END IF;

    IF TO_CHAR (SYSDATE, 'HH24:MI:SS') >= TO_CHAR (V_starttime, 'HH24:MI:SS') THEN
        V_nextdate :=TO_DATE (TO_CHAR (SYSDATE + 1, 'YYYY-MM-DD') || ' ' ||TO_CHAR (V_starttime, 'HH24:MI:SS'), 'YYYY-MM-DD HH24:MI:SS');
    ELSE
        V_nextdate :=TO_DATE (TO_CHAR (SYSDATE, 'YYYY-MM-DD') || ' ' ||TO_CHAR (V_starttime, 'HH24:MI:SS'), 'YYYY-MM-DD HH24:MI:SS');
    END IF;

    V_INterval :=
      'trunc(Sysdate)+' || V_cyclekeep || '+' || TO_CHAR (V_starttime, 'HH24') ||'/24' || '+' || TO_CHAR (V_starttime, 'MI') ||
          '/(24*60)' || '+' || TO_CHAR (V_starttime, 'SS') || '/(24*60*60)';
    --提交作业
    DBMS_JOB.Submit (V_jobnum, V_what, V_nextdate, V_INterval);

    UPDATE Zlautojobs
        SET 作业号 = V_jobnum
     WHERE Nvl(系统,0) = Nvl(Job_system,0)
        AND 类型 = Job_kind
        AND 序号 = Job_odd;
END;
/

CREATE OR REPLACE PROCEDURE zl_JobRemove(
    Job_system IN INTEGER,
    Job_kind IN INTEGER,
    Job_odd IN INTEGER
)
IS
    V_jobnum NUMBER := 0;
BEGIN
    SELECT 作业号
      INTO V_jobnum
      FROM Zlautojobs
     WHERE Nvl(系统,0) = Nvl(Job_system,0)
        AND 类型 = Job_kind
        AND 序号 = Job_odd;
    --删除作业
    DBMS_JOB.Remove (V_jobnum);

    UPDATE Zlautojobs
        SET 作业号 = NULL
     WHERE Nvl(系统,0) = Nvl(Job_system,0)
        AND 类型 = Job_kind
        AND 序号 = Job_odd;
END;
/

CREATE OR REPLACE PROCEDURE zl_JobRun(
    Job_system IN INTEGER,
    Job_kind IN INTEGER,
    Job_odd IN INTEGER
)
IS
    V_jobnum NUMBER := 0;
BEGIN
    SELECT 作业号
      INTO V_jobnum
      FROM Zlautojobs
     WHERE Nvl(系统,0) = Nvl(Job_system,0)
        AND 类型 = Job_kind
        AND 序号 = Job_odd;
    --执行作业
    DBMS_JOB.Run (V_jobnum);
END;
/

CREATE OR REPLACE PROCEDURE zl_JobChange(
    Job_system IN INTEGER,
    Job_kind IN INTEGER,
    Job_odd IN INTEGER
)
IS
    V_content VARCHAR2 (200);
    V_parameter VARCHAR2 (200);
    V_paraitem VARCHAR2 (200);
    V_starttime DATE;
    V_cyclekeep INTEGER;
    V_jobnum NUMBER := 0;
    V_what VARCHAR2 (1000);
    V_nextdate DATE;
    V_INterval VARCHAR2 (1000);
BEGIN
    SELECT 内容, 参数, 执行时间, 间隔时间, 作业号
      INTO V_content, V_parameter, V_starttime, V_cyclekeep, V_jobnum
      FROM Zlautojobs
     WHERE Nvl(系统,0) = Nvl(Job_system,0)
        AND 类型 = Job_kind
        AND 序号 = Job_odd;
    V_what := '';

    IF LENGTH (V_parameter) > 0 THEN
        LOOP
            IF INSTR (V_parameter, ';') > 0 THEN
                V_paraitem := SUBSTR (V_parameter, 1, INSTR (V_parameter, ';') - 1);
                V_parameter := SUBSTR (V_parameter, INSTR (V_parameter, ';') + 1);
            ELSE
                V_paraitem := V_parameter;
            END IF;

            V_what :=
                     V_what || ',' || SUBSTR (
                                                V_paraitem, INSTR (V_paraitem, ',') + 1
                                            );
            EXIT WHEN INSTR (V_parameter, ';') = 0;
        END LOOP;
    END IF;

    IF LENGTH (V_what) <> 0 THEN
        V_what := V_content || '(' || SUBSTR (V_what, 2) || ');';
    ELSE
        V_what := V_content || ';';
    END IF;

    IF TO_CHAR (SYSDATE, 'HH24:MI:SS') >= TO_CHAR (V_starttime, 'HH24:MI:SS') THEN
        V_nextdate :=
          TO_DATE (
              TO_CHAR (SYSDATE + 1, 'YYYY-MM-DD') || ' ' ||
                  TO_CHAR (V_starttime, 'HH24:MI:SS'), 'YYYY-MM-DD HH24:MI:SS'
          );
    ELSE
        V_nextdate :=
          TO_DATE (
              TO_CHAR (SYSDATE, 'YYYY-MM-DD') || ' ' ||
                  TO_CHAR (V_starttime, 'HH24:MI:SS'), 'YYYY-MM-DD HH24:MI:SS'
          );
    END IF;

    V_INterval :=
      'trunc(Sysdate)+' || V_cyclekeep || '+' || TO_CHAR (V_starttime, 'HH24') ||
          '/24' ||
          '+' ||
          TO_CHAR (V_starttime, 'MI') ||
          '/(24*60)' ||
          '+' ||
          TO_CHAR (V_starttime, 'SS') ||
          '/(24*60*60)';
    --修改作业
    DBMS_JOB.Change (V_jobnum, V_what, V_nextdate, V_INterval);
END;
/


