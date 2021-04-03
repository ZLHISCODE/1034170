--[连续升级]1
--[管理工具版本号]10.34.30
--本脚本支持从ZLHIS+ v10.34.20 升级到 v10.34.30

Define n_System=100;
-------------------------------------------------------------------------------
--结构修正部份
-------------------------------------------------------------------------------

--84458:涂建华,2015-07-02,PACS新版报告编辑器处理

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
--数据修正部份
-------------------------------------------------------------------------------





-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------





-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------






-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------





---------------------------------------------------------------------------------------------------
--更改系统及部件的版本号
-------------------------------------------------------------------------------------------------------
--系统版本号

--部件版本号
Commit;
