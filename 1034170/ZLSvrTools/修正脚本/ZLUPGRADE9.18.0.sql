-------------------------------------------------------------------------------
--�ṹ��������
-------------------------------------------------------------------------------
--ZLTOOLS��̨��ҵ:����->Ψһ��
ALTER TABLE zlAutoJobs DROP CONSTRAINT zlAutoJobs_PK
/
ALTER TABLE zlAutoJobs ADD CONSTRAINT 
    zlAutoJobs_PK UNIQUE(ϵͳ,����,���)
    USING INDEX PCTFREE 5 
    STORAGE(INITIAL 256 NEXT 128 PCTINCREASE 0 MAXEXTENTS UNLIMITED)
/

-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--��̨��ҵ�������кʹ�����־
Insert Into zlAutoJobs(ϵͳ,����,���,����,˵��,����,����,ִ��ʱ��,���ʱ��) Values(Null,1,1,'��־�����Զ�����','�Զ����������־�ʹ�����־����������Լ����쳣��������־���б�ǡ�','zl_AutoLogProcess',Null,Trunc(Sysdate)+2/24,1)
/
--7048��������Ч����
insert into zlSvrTools(���,�ϼ�,����,���,˵��) values ('010301','0103','������Ч����','P',Null)
/

-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
Create Or Replace Procedure zl_AutoLogProcess as
--���ܣ�
--   1.�Զ����������־�ʹ�����־�������
--   2.���쳣��������־���б��
	v_Count		Number;
	v_Limit		Number;
Begin
	--ɾ�������������־
	Select Count(*) Into v_Count From zlDiaryLog;
	Begin
		Select Nvl(to_Number(����ֵ),0) Into v_Limit From zlOptions Where ������=2;
	Exception
		When Others Then v_Limit:=10000;
	End;
	If v_Count>v_Limit Then
		Delete From zlDiaryLog Where RowID IN(
			Select ID From (
				Select RowID As ID From zlDiaryLog Group By ����ʱ��,RowID) 
			Where Rownum<v_Count-v_Limit+1);
	End If;
	
	--���쳣�˳���������־��¼���д���
	Update zlDiaryLog 
		Set �˳�ԭ��=2,�˳�ʱ��=Sysdate
	Where �˳�ԭ�� is Null And �Ự�� Not IN(Select SID+SERIAL# From v$Session Where USER#<>0);

	--ɾ������Ĵ�����־
	Select Count(*) Into v_Count From zlErrorLog;
	Begin
		Select Nvl(to_Number(����ֵ),0) Into v_Limit From zlOptions Where ������=4;
	Exception
		When Others Then v_Limit:=10000;
	End;
	If v_Count>v_Limit Then
		Delete From zlErrorLog Where RowID IN(
			Select ID From (
				Select RowID As ID From zlErrorLog Group By ʱ��,RowID) 
			Where Rownum<v_Count-v_Limit+1);
	End If;
End zl_AutoLogProcess;
/

-------------------------------------------------------
--ZL��ҵ����
--˵���� ���麯���ǹ����ߵ���ҵ������
--�嵥�� 1��zl_JobSubmit����ҵ�ύ
--       2��zl_JobRemove����ҵɾ��
--       3��zl_JobRun����ҵִ��
--       4��zl_JobChange����ҵ�޸�
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
    SELECT ����, ����, ִ��ʱ��, ���ʱ��
      INTO V_content, V_parameter, V_starttime, V_cyclekeep
      FROM Zlautojobs
     WHERE Nvl(ϵͳ,0) = Nvl(Job_system,0)
        AND ���� = Job_kind
        AND ��� = Job_odd;
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
    --�ύ��ҵ
    DBMS_JOB.Submit (V_jobnum, V_what, V_nextdate, V_INterval);

    UPDATE Zlautojobs
        SET ��ҵ�� = V_jobnum
     WHERE Nvl(ϵͳ,0) = Nvl(Job_system,0)
        AND ���� = Job_kind
        AND ��� = Job_odd;
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
    SELECT ��ҵ��
      INTO V_jobnum
      FROM Zlautojobs
     WHERE Nvl(ϵͳ,0) = Nvl(Job_system,0)
        AND ���� = Job_kind
        AND ��� = Job_odd;
    --ɾ����ҵ
    DBMS_JOB.Remove (V_jobnum);

    UPDATE Zlautojobs
        SET ��ҵ�� = NULL
     WHERE Nvl(ϵͳ,0) = Nvl(Job_system,0)
        AND ���� = Job_kind
        AND ��� = Job_odd;
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
    SELECT ��ҵ��
      INTO V_jobnum
      FROM Zlautojobs
     WHERE Nvl(ϵͳ,0) = Nvl(Job_system,0)
        AND ���� = Job_kind
        AND ��� = Job_odd;
    --ִ����ҵ
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
    SELECT ����, ����, ִ��ʱ��, ���ʱ��, ��ҵ��
      INTO V_content, V_parameter, V_starttime, V_cyclekeep, V_jobnum
      FROM Zlautojobs
     WHERE Nvl(ϵͳ,0) = Nvl(Job_system,0)
        AND ���� = Job_kind
        AND ��� = Job_odd;
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
    --�޸���ҵ
    DBMS_JOB.Change (V_jobnum, V_what, V_nextdate, V_INterval);
END;
/


