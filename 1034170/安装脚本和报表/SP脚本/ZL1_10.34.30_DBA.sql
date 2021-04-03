--[��������]1
--[�����߰汾��]10.34.30
--���ű�֧�ִ�ZLHIS+ v10.34.20 ������ v10.34.30

Define n_System=100;
-------------------------------------------------------------------------------
--�ṹ��������
-------------------------------------------------------------------------------

--84458:Ϳ����,2015-07-02,PACS�°汨��༭������

Declare
  v_Count Number;
  v_Path Varchar2(255);     
Begin
  Select Substr(File_Name, 1, Decode(Instr(File_Name, '\', -1), 0, Instr(File_Name, '/', -1),Instr(File_Name, '\', -1))) Into v_Path From Dba_Data_Files Where Upper(Tablespace_Name) = Upper('ZL9BASEITEM') And Rownum < 2;
  
  Select Count(*) Into v_Count From DBA_TableSpaces Where Upper(TableSpace_Name)=Upper('zlPacsBaseTab');
  If Nvl(v_Count,0)=0 Then
    Execute Immediate 'Create Tablespace zlPacsBaseTab Datafile '''||v_Path||'zlPacsBaseTab.DBF'' Size 50M REUSE Autoextend On Next 10M Online DEFAULT STORAGE ( INITIAL 512K NEXT 128K MAXEXTENTS UNLIMITED PCTINCREASE 0)';
	End IF;
End;
/


Declare
  v_Count Number;
  v_Path Varchar2(255);     
Begin
  Select Substr(File_Name, 1, Decode(Instr(File_Name, '\', -1), 0, Instr(File_Name, '/', -1),Instr(File_Name, '\', -1))) Into v_Path From Dba_Data_Files Where Upper(Tablespace_Name) = Upper('ZL9BASEITEM') And Rownum < 2;
  
  Select Count(*) Into v_Count From DBA_TableSpaces Where Upper(TableSpace_Name)=Upper('zlPacsBaseIndex');
  If Nvl(v_Count,0)=0 Then
    Execute Immediate 'Create Tablespace zlPacsBaseIndex Datafile '''||v_Path||'zlPacsBaseIndex.DBF'' Size 50M REUSE Autoextend On Next 10M Online DEFAULT STORAGE ( INITIAL 512K NEXT 128K MAXEXTENTS UNLIMITED PCTINCREASE 0)';
	End IF;
End;
/

Declare
  v_Count Number;
  v_Path Varchar2(255);     
Begin
  Select Substr(File_Name, 1, Decode(Instr(File_Name, '\', -1), 0, Instr(File_Name, '/', -1),Instr(File_Name, '\', -1))) Into v_Path From Dba_Data_Files Where Upper(Tablespace_Name) = Upper('ZL9BASEITEM') And Rownum < 2;

  Select Count(*) Into v_Count From DBA_TableSpaces Where Upper(TableSpace_Name)=Upper('zlPacsBizTab');
  If Nvl(v_Count,0)=0 Then
    Execute Immediate 'Create Tablespace zlPacsBizTab Datafile '''||v_Path||'zlPacsBizTab.DBF'' Size 50M REUSE Autoextend On Next 10M Online DEFAULT STORAGE ( INITIAL 512K NEXT 128K MAXEXTENTS UNLIMITED PCTINCREASE 0)';
	End IF;
End;
/

Declare
  v_Count Number;
  v_Path Varchar2(255);     
Begin
  Select Substr(File_Name, 1, Decode(Instr(File_Name, '\', -1), 0, Instr(File_Name, '/', -1),Instr(File_Name, '\', -1))) Into v_Path From Dba_Data_Files Where Upper(Tablespace_Name) = Upper('ZL9BASEITEM') And Rownum < 2;
  
  Select Count(*) Into v_Count From DBA_TableSpaces Where Upper(TableSpace_Name)=Upper('zlPacsBizIndex');
  If Nvl(v_Count,0)=0 Then
    Execute Immediate 'Create Tablespace zlPacsBizIndex Datafile '''||v_Path||'zlPacsBizIndex.DBF'' Size 50M REUSE Autoextend On Next 10M Online DEFAULT STORAGE ( INITIAL 512K NEXT 128K MAXEXTENTS UNLIMITED PCTINCREASE 0)';
	End IF;
End;
/

Declare
  v_Count Number;
  v_Path Varchar2(255);     
Begin
  Select Substr(File_Name, 1, Decode(Instr(File_Name, '\', -1), 0, Instr(File_Name, '/', -1),Instr(File_Name, '\', -1))) Into v_Path From Dba_Data_Files Where Upper(Tablespace_Name) = Upper('ZL9BASEITEM') And Rownum < 2;

  Select Count(*) Into v_Count From DBA_TableSpaces Where Upper(TableSpace_Name)=Upper('zlPacsBizXml');
  If Nvl(v_Count,0)=0 Then
    Execute Immediate 'Create Tablespace zlPacsBizXml Datafile '''||v_Path||'zlPacsBizXml.DBF'' Size 100M REUSE Autoextend On Next 30M Online DEFAULT STORAGE ( INITIAL 512K NEXT 128K MAXEXTENTS UNLIMITED PCTINCREASE 0)';
	End IF;
End;
/



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------





-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------





-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------






-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------





---------------------------------------------------------------------------------------------------
--����ϵͳ�������İ汾��
-------------------------------------------------------------------------------------------------------
--ϵͳ�汾��

--�����汾��
Commit;
