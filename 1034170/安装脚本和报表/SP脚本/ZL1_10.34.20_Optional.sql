--[连续升级]1
--[管理工具版本号]10.34.0
--本脚本支持从ZLHIS+ v10.34.10 升级到 v10.34.20
--请以系统所有者登录PLSQL并执行下列脚本
--脚本执行后，请手工升级导出报表
Define n_System=100;
-------------------------------------------------------------------------------
------------------------------------------------------------------------------

--65664:张永康,2015-01-14,删除"医嘱执行计价"表中多余的药嘱记录
--10.31.20上的58727增加了表"医嘱执行计价",由于错误的产生了多余的药嘱相关记录（约占96%），该表的记录快速增长，几个月时间可能就会产生上亿条记录。
--直到10.33.0的65664更正了该错误，不再产生药嘱记录。
--这些多余的数据，占用大量磁盘空间，并且将会大大增加索引重建和统计信息收集等维护工作的耗时，也给相关的SQL查询带来性能风险，所以，提供此脚本用于删除该表中的药嘱相关记录。
--由于数据修正比较耗时，所以，请根据当前用户的情况判断是否有必要执行，如果是以下两种情况，则过程中已判断不执行此脚本（因为没有产生多余的药嘱记录）
--1.直接从10.31.20以下的版本，升级到10.33及以上版本
--2.安装时就是从10.33及以上版本开始的
Create Or Replace Procedure Zl_医嘱执行计价_Purge Is
  n_Do     Number(5) := 1;
  v_Vermin Varchar2(20);

  Function Formatver(v_Ver Varchar2) Return Varchar2 Is
    v_Result Varchar2(20);
  Begin
    Select f_List2str(Cast(Collect(Ver) As t_Strlist), '.')
    Into v_Result
    From (Select LPad(Column_Value, 4, '0') Ver From Table(f_Str2list(v_Ver, '.')));
    Return v_Result;
  End;
Begin
  v_Vermin := '0000.0000.0000';
  For Rv In (Select 原始版本, 目标版本 From zlUpGrade Where 系统 = 100 Order By 原始版本) Loop
    If Formatver(Rv.原始版本) < '0010.0031.0020' And Formatver(Rv.目标版本) > '0010.0033.0000' Then
      n_Do := 0;
      Exit;
    End If;
  
    If v_Vermin < Formatver(Rv.原始版本) Then
      v_Vermin := Formatver(Rv.原始版本);
    End If;
  End Loop;
  If v_Vermin > '0010.0033.0000' Then
    n_Do := 0;
  End If;

  If n_Do = 1 Then
    Execute Immediate 'Alter table 医嘱执行计价 rename to 医嘱执行计价_old';
    Execute Immediate 'Create table 医嘱执行计价 nologging tablespace zl9CisRec Initrans 20 as Select a.* From 医嘱执行计价_old A, 病人医嘱记录 B Where a.医嘱id = b.Id And b.诊疗类别 not in(''4'',''5'',''6'',''7'')';
    Execute Immediate 'Alter table 医嘱执行计价 modify 费用性质 default 0';
  
    Execute Immediate 'Alter table 医嘱执行计价_old drop constraint 医嘱执行计价_PK cascade Drop index';
    Execute Immediate 'Alter table 医嘱执行计价_old drop constraint 医嘱执行计价_FK_发送号';
    Execute Immediate 'Alter table 医嘱执行计价_old drop constraint 医嘱执行计价_FK_收费细目ID';
    Execute Immediate 'Drop index 医嘱执行计价_IX_收费细目ID';
    Execute Immediate 'Drop index 医嘱执行计价_IX_待转出';
  
    Execute Immediate 'Create index 医嘱执行计价_PK On 医嘱执行计价(医嘱ID,发送号,要求时间,收费细目ID,费用性质) Pctfree 5 Tablespace zl9Indexcis nologging';
    Execute Immediate 'Alter table 医嘱执行计价 Add Constraint 医嘱执行计价_PK Primary Key (医嘱ID,发送号,要求时间,收费细目ID,费用性质) enable novalidate';
    Execute Immediate 'Alter table 医嘱执行计价 modify constraint 医嘱执行计价_PK validate';
  
    Execute Immediate 'Alter table 医嘱执行计价 Add Constraint 医嘱执行计价_FK_发送号 Foreign Key (医嘱ID,发送号) References 病人医嘱发送(医嘱ID,发送号) On Delete Cascade enable novalidate';
    Execute Immediate 'Alter table 医嘱执行计价 Add Constraint 医嘱执行计价_FK_收费细目id Foreign Key (收费细目id) References 收费项目目录(ID) enable novalidate';
  
    Execute Immediate 'Create index 医嘱执行计价_IX_收费细目id On 医嘱执行计价(收费细目id) Pctfree 5 Tablespace zl9Indexcis';
    Execute Immediate 'Create index 医嘱执行计价_IX_待转出 On 医嘱执行计价(待转出) Tablespace zl9Indexcis';
  
    Execute Immediate 'DROP TABLE 医嘱执行计价_old purge';
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_医嘱执行计价_Purge;
/

-------------------------------------------------------------------------------------------------------
--更改系统及部件的版本号
-------------------------------------------------------------------------------------------------------
--系统版本号
--可选脚本不用更新
--部件版本号
Commit;

