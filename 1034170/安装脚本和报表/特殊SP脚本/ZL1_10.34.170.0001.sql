----------------------------------------------------------------------------------------------------------------
--本脚本支持从ZLHIS+ v10.34.170升级到 v10.34.170
--请以数据空间的所有者登录PLSQL并执行下列脚本
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
---------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------
--结构修正部份
------------------------------------------------------------------------------


------------------------------------------------------------------------------
--数据修正部份
------------------------------------------------------------------------------


-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--报表修正部份
-------------------------------------------------------------------------------


-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------
--138081:焦博,2019-02-27,大医二院反馈的死锁问题
Create Or Replace Procedure Zl1_Autocptpati
(
  Patiid      In Number,
  Pageid      In Number,
  Recalcbdate In 病人变动记录.上次计算时间%Type := Null,
  强制记帐_In In Number := 0
) As
  Modilast Number(1); --是否修正上期自动计费参数
  Period   Varchar2(6); --需要计算的最小期间
Begin
  Begin
    Select 期间 Into Period From 期间表 Where Trunc(Sysdate) Between Trunc(开始日期) And Trunc(终止日期);
  Exception
    When Others Then
      Return;
  End;

  Select Zl_To_Number(zl_GetSysParameter(7)) Into Modilast From Dual;

  If Modilast = 1 Then
    Period := To_Char(Add_Months(To_Date(Period || '05', 'yyyymmdd'), -1), 'yyyymm');
  End If;

  If Recalcbdate Is Not Null Then
    Update 病人变动记录
    Set 上次计算时间 = Null
    Where 病人id = Patiid And 主页id = Pageid And 上次计算时间 >= Recalcbdate;
    Commit;
  End If;

  Zl1_Autocptone(Patiid, Pageid, Period, 0, 强制记帐_In);
End Zl1_Autocptpati;
/

--138081:焦博,2019-02-27,大医二院反馈的死锁问题
Create Or Replace Procedure Zl1_Autocptward
(
  Wardid      In Number,
  Recalcbdate In 病人变动记录.上次计算时间%Type := Null,
  强制记帐_In In Number := 0
) As
  Modilast Number(1); --是否修正上期自动计费参数
  Period   Varchar2(6); --需要计算的最小期间

  Cursor Patitab Is
    Select Distinct 病人id, 主页id
    From 在院病人自动记帐
    Where 病区id = Wardid And Trunc(终止日期) >= (Select Min(开始日期) From 期间表 Where 期间 >= Period);
Begin
  Begin
    Select 期间 Into Period From 期间表 Where Trunc(Sysdate) - 1 Between Trunc(开始日期) And Trunc(终止日期);
  Exception
    When Others Then
      Return;
  End;
  Select zl_GetSysParameter(7) Into Modilast From Dual;

  If Modilast = 1 Then
    Period := To_Char(Add_Months(To_Date(Period || '05', 'yyyymmdd'), -1), 'yyyymm');
  End If;

  For Patifld In Patitab Loop
    If Patifld.病人id Is Not Null And Patifld.主页id Is Not Null Then
      If Recalcbdate Is Not Null Then
        Update 病人变动记录
        Set 上次计算时间 = Null
        Where 病人id = Patifld.病人id And 主页id = Patifld.主页id And 上次计算时间 >= Recalcbdate;
        Commit;
      End If;
      Zl1_Autocptone(Patifld.病人id, Patifld.主页id, Period, 1, 强制记帐_In);
    End If;
  End Loop;
End Zl1_Autocptward;
/



------------------------------------------------------------------------------------
--系统版本号
Update zlSystems Set 版本号='10.34.170.0001' Where 编号=&n_System;
Commit;
