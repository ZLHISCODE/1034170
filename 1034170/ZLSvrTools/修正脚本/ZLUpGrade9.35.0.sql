-----------------------------------------------------------------
--为配合产品版本号由9.34升为9.35
-----------------------------------------------------------------
--适应参数调整(刘兴宏),同时调整了重复的快键
UPDATE zlSvrTools SET 快键='E' WHERE 编号='0307'
/
UPDATE zlSvrTools SET 快键='R' WHERE 编号='0308'
/
UPDATE zlSvrTools SET 快键='A' WHERE 编号='0309'
/
Insert Into zlSvrTools(编号,上级,标题,快键,说明) Values('0310','03','系统参数管理','P',Null)
/

Create Table zlTools.zlParaChangedLog(
    参数ID NUMBER(18),
    序号   NUMBER(18),
    变动说明 VARCHAR2(200),--说明变动情况:比如:私有模块变为公用模块。
    变动内容 VARCHAR2(200),--说明变动字段的变化情况:比如:私有:1-->0,本机:1-->0。
    变动人 VARCHAR2(20),
    变动时间 Date,
    变动原因 varchar2(200))
    PCTFREE 5 PCTUSED 90
    Cache Storage(Buffer_Pool Keep)
/
Alter Table zlTools.zlParaChangedLog Add Constraint zlParaChangedLog_UQ_参数ID Unique(参数ID,序号) Using Index PCTFREE 5
/
Alter Table zlTools.zlParaChangedLog Add Constraint zlParaChangedLog_FK_参数ID Foreign Key (参数ID) References zlTools.zlParameters(ID) On Delete Cascade
/
Create Index zlTools.zlParaChangedLog_IX_变动人 On zlTools.zlParaChangedLog(变动人) PCTFREE 5
/ 

grant Select on zlTools.zlParaChangedLog to Public
/
create public SYNONYM  zlParaChangedLog   FOR zlTools.zlParaChangedLog
/


--********************************************************************************************************
--参数处理调整Begin
--********************************************************************************************************
--先调整数据，以便后面的约束能正确调整。因为药品部分模块的参数号，私有和公共是分开编号的，需改为统一编号。
--先将私有的参数号+100，ZLHIS脚本中再正确调整。程序中模块参数没有使用参数号。
Update zlParameters
Set 参数号 = 参数号 + 100
Where 私有 = 1 And
      (系统, 模块, 参数号) In (Select 系统, 模块, 参数号 From zlParameters Group By 系统, 模块, 参数号 Having Count(*) > 1)
/

--zlParameters
Alter Table zlTools.zlParameters Add(本机 Number(1),授权 Number(1),固定 Number(1))
/
Alter Table zlTools.zlParameters Add Constraint zlParameters_CK_本机 Check (本机 IN(0,1))
/
Alter Table zlTools.zlParameters Add Constraint zlParameters_CK_授权 Check (授权 IN(0,1))
/
Alter Table zlTools.zlParameters Add Constraint zlParameters_CK_固定 Check (固定 IN(0,1))
/
Alter Table zlTools.zlParameters Drop Constraint zlParameters_UQ_参数号
/
Alter Table zlTools.zlParameters Drop Constraint zlParameters_UQ_参数名
/
Alter Table zlTools.zlParameters Add Constraint zlParameters_UQ_参数号 Unique(参数号,模块,系统) Using Index PCTFREE 5
/
Alter Table zlTools.zlParameters Add Constraint zlParameters_UQ_参数名 Unique(参数名,模块,系统) Using Index PCTFREE 5
/

--zlUserParas
Alter Table zlTools.zlUserParas Add 机器名 Varchar2(50)
/
Alter Table zlTools.zlUserParas Drop Constraint zlUserParas_PK
/
Alter Table zlTools.zlUserParas Add Constraint zlUserParas_UQ_参数ID Unique(参数ID,用户名,机器名) Using Index PCTFREE 5
/
Create Index zlTools.zlUserParas_IX_机器名 On zlUserParas(机器名) PCTFREE 5
/


--20952
insert into zlreginfo(项目,内容) values('站点数量',Null);


Create Or Replace Function zlTools.f_Get_Node_Amt Return Number As
  v_Return Number;
Begin
  Begin
    Select 内容 Into v_Return From zlRegInfo Where 项目 = '站点数量';
    If To_Number(v_Return) < 0 Or To_Number(v_Return) > 9 Then
      v_Return := Null;
    End If;
  Exception
    When Others Then
      Null;
  End;
  Return(v_Return);
End f_Get_Node_Amt;
/

Create Public SYNONYM f_Get_Node_Amt FOR zlTools.f_Get_Node_Amt
/
Grant Execute on zlTools.f_Get_Node_Amt to Public
/
alter table ZLTOOLS.ZLPARAMETERS modify 参数值 VARCHAR2(2000)
/
alter table ZLTOOLS.ZLPARAMETERS modify 缺省值 VARCHAR2(2000)
/
alter table ZLTOOLS.zlUserParas modify 参数值 VARCHAR2(2000)
/
--过程
Create Or Replace Procedure zlTools.Zl_Parameters_Update
(
  参数_In   zlParameters.参数名%Type,
  参数值_In zlParameters.参数值%Type,
  系统_In   zlParameters.系统%Type,
  模块_In   zlParameters.模块%Type
  --功能：设置系统参数值，如果是用户私有参数，则用户名以当前的为准
  --参数：
  --      参数_In：必须传入非Null值，以字符形式传入的参数号或参数名,注意参数名不能为数字。
) Is
  v_参数id zlParameters.ID%Type;
  v_私有   zlParameters.私有%Type;
  v_本机   zlParameters.本机%Type;
  v_机器名 zlUserParas.机器名%Type;
Begin
  --确定参数信息
  Begin
    If Zl_To_Number(参数_In) <> 0 Then
      --以参数号为准处理
      Select ID, 私有, 本机, Sys_Context('USERENV', 'TERMINAL')
      Into v_参数id, v_私有, v_本机, v_机器名
      From zlParameters
      Where Nvl(系统, 0) = Nvl(系统_In, 0) And Nvl(模块, 0) = Nvl(模块_In, 0) And 参数号 = Zl_To_Number(参数_In);
    Else
      --以参数名为准处理
      Select ID, 私有, 本机, Sys_Context('USERENV', 'TERMINAL')
      Into v_参数id, v_私有, v_本机, v_机器名
      From zlParameters
      Where Nvl(系统, 0) = Nvl(系统_In, 0) And Nvl(模块, 0) = Nvl(模块_In, 0) And 参数名 = 参数_In;
    End If;
  Exception
    When Others Then
      Return;
  End;

  --更新参数值
  If v_参数id Is Not Null Then
    If Nvl(v_私有, 0) = 0 And Nvl(v_本机, 0) = 0 Then
      Update zlParameters Set 参数值 = 参数值_In Where ID = v_参数id;
    Else
      Update zlUserParas
      Set 参数值 = 参数值_In
      Where 参数id = v_参数id And Nvl(用户名, 'NullUser') = Decode(v_私有, 1, User, 'NullUser') And
            Nvl(机器名, 'NullMachine') = Decode(v_本机, 1, v_机器名, 'NullMachine');
      If Sql%RowCount = 0 Then
        Insert Into zlUserParas
          (参数id, 用户名, 机器名, 参数值)
        Values
          (v_参数id, Decode(v_私有, 1, User, Null), Decode(v_本机, 1, v_机器名, Null), 参数值_In);
      End If;
    End If;
  End If;
End Zl_Parameters_Update;
/

Create Or Replace Procedure zlTools.zl_Parameters_Update_Batch
(
  系统编号_In zlSystems.编号%Type,
  参数列表_In Varchar2
) Is
  --参数列表_IN 参数的填写方式如下："参数号1,参数值1,参数号2,参数值2,"                                            
  n_Pos    Number(5);
  v_Temp   Varchar2(2000);
  v_参数号 zlParameters.参数号%Type;
  v_参数值 zlParameters.参数值%Type;
Begin
  --循环处理
  v_Temp := 参数列表_In;

  While v_Temp Is Not Null Loop
    n_Pos := Instr(v_Temp, ',');
  
    If n_Pos = 0 Then
      v_Temp := '';
    Else
      --得到参数号
      v_参数号 := Substr(v_Temp, 1, n_Pos - 1);
      v_Temp   := Substr(v_Temp, n_Pos + 1);
      --得到参数值
      n_Pos    := Instr(v_Temp, ',');
      v_参数值 := Substr(v_Temp, 1, n_Pos - 1);
      v_Temp   := Substr(v_Temp, n_Pos + 1);
    
      Update zlParameters
      Set 参数值 = v_参数值
      Where 系统 = 系统编号_In And 模块 Is Null And 参数号 = To_Number(v_参数号);
    End If;
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End zl_Parameters_Update_Batch;
/
--********************************************************************************************************
--参数处理调整End
--********************************************************************************************************

--15040
Create Or Replace Procedure Zltools.Zl_Createsynonyms(系统_In In zlProgPrivs.系统%Type) Authid Current_User As
  v_Sql    Varchar2(2000);
  v_所有者 Varchar2(100);
  n_Cnt    Number(5);

  --非当前所有者的对象的私有同义词与当前所有者的对象相同则删除
  Cursor c_Delsyn(v_所有者 Varchar2) Is
    Select Synonym_Name 对象
    From User_Synonyms A
    Where Table_Owner != v_所有者 And Exists
     (Select 1
           From All_Objects B
           Where A.Synonym_Name = B.Object_Name And B.Owner = v_所有者 And
                 B.Object_Type In ('TABLE', 'VIEW', 'SEQUENCE', 'PROCEDURE', 'FUNCTION'));

  --对用户有权访问的对象,属于当前系统所有者的,如果没有公共或私有同义词,则创建私有同义词
  --不仅限于当前模块所访问的对象,因为存在虚拟模块和模块中调用其它模块
  Cursor c_Newsyn(v_所有者 Varchar2) Is
    Select Object_Name 对象, Owner 所有者
    From All_Objects A
    Where Owner = v_所有者 And Instr(Object_Name, 'BIN$') = 0 And
          Object_Type In ('TABLE', 'VIEW', 'SEQUENCE', 'PROCEDURE', 'FUNCTION')
    Minus
    Select Synonym_Name, Table_Owner
    From All_Synonyms C
    Where Table_Owner = v_所有者 And (Owner = User Or Owner = 'PUBLIC');

Begin
  Select Count(Distinct 所有者) Into n_Cnt From Zlsystems;
  If n_Cnt > 1 Then
    Select Upper(所有者) Into v_所有者 From Zlsystems Where 编号 = 系统_In;
    --角色授权已限制所有者不能访问其它系统,所以,系统所有者不用创建私有同义词
    If v_所有者 != User Then
      For c_Syn In c_Delsyn(v_所有者) Loop
        v_Sql := 'Drop Synonym ' || c_Syn.对象;
        Execute Immediate v_Sql;
      End Loop;
    
      For c_Syn In c_Newsyn(v_所有者) Loop
        v_Sql := 'Create Synonym ' || c_Syn.对象 || ' For ' || c_Syn.所有者 || '.' || c_Syn.对象;
        Execute Immediate v_Sql;
      End Loop;
    End If;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Createsynonyms;
/

Create Or Replace Procedure Zltools.Zl_Createpubsynonyms Authid Current_User As
  v_Sql Varchar2(100);

  Cursor c_All Is
    Select Object_Name 对象, Owner 所有者
    From All_Objects A
    Where Owner = User And Instr(Object_Name, 'BIN$') = 0 And
          Object_Type In ('TABLE', 'VIEW', 'SEQUENCE', 'PROCEDURE', 'FUNCTION') And Exists
     (Select 1 From Zlsystems Where Upper(所有者) = User)
    Minus
    Select Synonym_Name, User
    From All_Synonyms
    Where (Table_Owner In (Select Distinct Upper(所有者) From Zlsystems) Or Table_Owner = 'ZLTOOLS') And Owner = 'PUBLIC';
  --如果其它系统有同名同义词,则不创建公共同义词,当用户进入模块时再创建私有同义词  
Begin

  For c_Syn In c_All Loop
    Begin
      v_Sql := 'Create Public Synonym ' || c_Syn.对象 || ' For ' || c_Syn.对象;
      Execute Immediate v_Sql;
    Exception
      When Others Then
        Null;
    End;
  End Loop;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Createpubsynonyms;
/


Create Public SYNONYM Zl_Createpubsynonyms FOR zlTools.Zl_Createpubsynonyms
/
Grant Execute on zlTools.Zl_Createpubsynonyms to Public
/


--15026：自动提醒增加共享报表的调用(Get_noticereport)，2009-01-08　By Fr.Chen
Create Or Replace Package Body b_Expert Is

  -----------------------------------------------------------------------------
  -- 取提醒数据
  -----------------------------------------------------------------------------
  Procedure Get_Notices
  (
    Cursor_Out Out t_Refcur,
    序号_In    In zlNotices.序号%Type,
    系统_In    In zlReports.系统%Type := Null
  ) Is
  Begin
    If Nvl(序号_In, 0) <> 0 Then
      -- frmNoticesEdit.ReadData 使用
      Open Cursor_Out For
        Select A.提醒内容, A.提醒条件, A.提醒报表, A.提醒声音, A.提醒窗口, A.开始时间, A.终止时间, A.检查周期, B.名称 As 报表名称
        From zlNotices A, zlReports B
        Where A.提醒报表 = B.编号(+) And A.序号 = 序号_In;
    Else
      -- cboSystem_Click 使用
      If Nvl(系统_In, 0) = 0 Then
        Open Cursor_Out For
          Select A.序号, A.提醒内容, A.提醒条件, A.提醒报表, A.提醒声音, A.提醒窗口, A.开始时间, A.终止时间, A.检查周期, A.提醒周期, B.名称 As 报表名称
          From zlNotices A, zlReports B
          Where A.提醒报表 = B.编号(+) And A.系统 Is Null;
      Else
        Open Cursor_Out For
          Select A.序号, A.提醒内容, A.提醒条件, A.提醒报表, A.提醒声音, A.提醒窗口, A.开始时间, A.终止时间, A.检查周期, A.提醒周期, B.名称 As 报表名称
          From zlNotices A, zlReports B
          Where A.提醒报表 = B.编号(+) And A.系统 = 系统_In;
      End If;
    End If;
  
  End Get_Notices;

  -----------------------------------------------------------------------------
  -- 取提醒对像数据
  -----------------------------------------------------------------------------
  Procedure Get_Noticeusr
  (
    Cursor_Out  Out t_Refcur,
    提醒对象_In In zlNoticeUsr.提醒对象%Type,
    提醒序号_In In zlNoticeUsr.提醒序号%Type
  ) Is
  Begin
    If Nvl(提醒对象_In, 0) = 0 Then
      Open Cursor_Out For
        Select 1 From zlNoticeUsr Where Rownum < 2 And 提醒序号 = 提醒序号_In;
    Else
      Open Cursor_Out For
        Select 对象名称 From zlNoticeUsr Where 提醒对象 = 提醒对象_In And 提醒序号 = 提醒序号_In;
    End If;
  End Get_Noticeusr;

  -----------------------------------------------------------------------------
  -- 取可以选择的提醒报表
  -----------------------------------------------------------------------------
  Procedure Get_Noticereport
  (
    Cursor_Out Out t_Refcur,
    系统_In    In zlReports.系统%Type
  ) Is
  Begin
    If Nvl(系统_In, 0) = 0 Then
      Open Cursor_Out For
        Select ID, 编号, 名称, 说明
        From zlReports
        Where Not (发布时间 Is Null Or Trunc(发布时间) = To_Date('3000-01-01', 'yyyy-mm-dd')) And
              Nvl(系统, 0) = 0;
    Else
      Open Cursor_Out For
        Select ID, 编号, 名称, 说明
        From zlReports
        Where ((系统 = 系统_In And 编号 Like 'ZL%_REPORT_%') Or 系统 Is Null)  And Not (发布时间 Is Null Or Trunc(发布时间) = To_Date('3000-01-01', 'yyyy-mm-dd'));
    End If;
  End Get_Noticereport;

  -----------------------------------------------------------------------------
  -- 在不同的系统间复制报表
  -----------------------------------------------------------------------------
  Procedure Copy_Report
  (
    系统_In   In zlReports.系统%Type,
    新系统_In In zlReports.系统%Type
  ) Is
    n_Grpid   Number;
    n_Rptid   Number;
    n_Dataid  Number;
    n_Itemid  Number;
    v_Olduser Varchar2(100);
    v_Newuser Varchar2(100);
  
    Function Sub_Owner_Name(Lngsys_In In Number := 0) Return Varchar2 Is
      v_Owner_Name Varchar2(30);
    Begin
      Select Upper(所有者) As 所有者 Into v_Owner_Name From zlSystems Where 编号 = Lngsys_In;
      Return v_Owner_Name;
    End Sub_Owner_Name;
  
  Begin
    Select Nvl(Max(ID), 0) Into n_Grpid From zlRPTGroups;
    Select Nvl(Max(ID), 0) Into n_Rptid From zlReports;
    Select Nvl(Max(ID), 0) Into n_Dataid From zlRPTDatas;
    Select Nvl(Max(ID), 0) Into n_Itemid From zlRPTItems;
    n_Grpid  := n_Grpid + 1;
    n_Rptid  := n_Rptid + 1;
    n_Dataid := n_Dataid + 1;
    n_Itemid := n_Itemid + 1;
  
    v_Olduser := Upper(Sub_Owner_Name(系统_In));
    v_Newuser := Upper(Sub_Owner_Name(新系统_In));
  
    Insert Into zlRPTGroups
      (ID, 编号, 名称, 说明, 系统, 程序id, 发布时间)
      Select ID + n_Grpid, 编号, 名称, 说明, 新系统_In, 程序id, 发布时间 From zlRPTGroups Where 系统 = 系统_In;
  
    Insert Into zlReports
      (ID, 编号, 名称, 说明, 密码, 进纸, 打印机, 票据, 系统, 程序id, 功能, 修改时间, 发布时间)
      Select ID + n_Rptid, 编号, 名称, 说明, 密码, 进纸, 打印机, 票据, 新系统_In, 程序id, 功能, 修改时间, 发布时间
      From zlReports
      Where 系统 = 系统_In;
  
    -- 插入zlRPTSub
    Insert Into zlRPTSubs
      (组id, 报表id, 序号, 功能)
      Select A.组id + n_Grpid, A.报表id + n_Rptid, A.序号, A.功能
      From zlRPTSubs A, zlRPTGroups B
      Where A.组id = B.ID And B.系统 = 系统_In;
  
    -- 插入zlRPTFMTs
    Insert Into zlRPTFMTs
      (报表id, 序号, 说明, W, H, 纸张, 纸向, 动态纸张, 图样)
      Select A.报表id + n_Rptid, A.序号, A.说明, A.W, A.H, A.纸张, A.纸向, A.动态纸张, A.图样
      From zlRPTFMTs A, zlReports B
      Where A.报表id = B.ID And B.系统 = 系统_In;
  
    -- 插入zlRPTItems
    Insert Into zlRPTItems
      (ID, 报表id, 格式号, 名称, 类型, 上级id, 序号, 参照, 性质, 内容, 表头, X, Y, W, H, 行高, 对齐, 自调, 字体, 字号, 粗体, 斜体, 下线, 前景, 背景, 边框, 排序, 格式,
       汇总, 分栏, 网格, 系统)
      Select A.ID + n_Itemid, A.报表id + n_Rptid, A.格式号, A.名称, A.类型, A.上级id + n_Itemid, A.序号, A.参照, A.性质, A.内容, A.表头, A.X,
             A.Y, A.W, A.H, A.行高, A.对齐, A.自调, A.字体, A.字号, A.粗体, A.斜体, A.下线, A.前景, A.背景, A.边框, A.排序, A.格式, A.汇总, A.分栏,
             A.网格, A.系统
      From zlRPTItems A, zlReports B
      Where A.报表id = B.ID And B.系统 = 系统_In;
    -- 插入zlRptDatas
    Insert Into zlRPTDatas
      (ID, 报表id, 名称, 字段, 对象, 类型)
      Select A.ID + n_Dataid, A.报表id + n_Rptid, A.名称, A.字段, A.对象, A.类型
      From zlRPTDatas A, zlReports B
      Where A.报表id = B.ID And B.系统 = 系统_In;
    -- 插入zlRPTSqls
    Insert Into zlRPTSQLs
      (源id, 行号, 内容)
      Select A.源id + n_Dataid, A.行号, A.内容
      From zlRPTSQLs A, zlRPTDatas B, zlReports C
      Where A.源id = B.ID And B.报表id = C.ID And C.系统 = 系统_In;
  
    -- 插入zlRPTPars
    Insert Into zlRPTPars
      (源id, 组名, 序号, 名称, 类型, 缺省值, 格式, 值列表, 分类sql, 明细sql, 分类字段, 明细字段, 对象)
      Select A.源id + n_Dataid, A.组名, A.序号, A.名称, A.类型, A.缺省值, A.格式, A.值列表, A.分类sql, A.明细sql, A.分类字段, A.明细字段, A.对象
      From zlRPTPars A, zlRPTDatas B, zlReports C
      Where A.源id = B.ID And B.报表id = C.ID And C.系统 = 系统_In;
  
    -- zlFunctions数据
    Insert Into zlFunctions
      (系统, 函数号, 函数名, 中文名, 说明)
      Select 新系统_In, 函数号, 函数名, 中文名, 说明 From zlFunctions Where 系统 = 系统_In;
  
    -- zlFuncPars数据
    Insert Into zlFuncPars
      (系统, 函数号, 参数号, 参数名, 中文名, 类型, 缺省值, 值列表, 分类sql, 明细sql, 分类字段, 明细字段, 对象, 组名, 递增否)
      Select 新系统_In, 函数号, 参数号, 参数名, 中文名, 类型, 缺省值, 值列表, 分类sql, 明细sql, 分类字段, 明细字段, 对象, 组名, 递增否
      From zlFuncPars
      Where 系统 = 系统_In;
  
    -- 重新设置数据源对象
    Update zlRPTDatas
    Set 对象 = Replace(对象, v_Olduser || '.', v_Newuser || '.')
    Where ID In (Select A.ID From zlRPTDatas A, zlReports B Where A.报表id = B.ID And B.系统 = 新系统_In);
  
    Update zlRPTPars
    Set 对象 = Replace(对象, v_Olduser || '.', v_Newuser || '.')
    Where 源id In (Select A.ID From zlRPTDatas A, zlReports B Where A.报表id = B.ID And B.系统 = 新系统_In);
  
    Update zlFuncPars Set 对象 = Replace(对象, v_Olduser || '.', v_Newuser || '.') Where 系统 = 新系统_In;
  
    Commit;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Copy_Report;

End b_Expert;
/
--15026：自动提醒增加共享报表的调用(Get_Zlnoticerec)，2009-01-08　By Fr.Chen
Create Or Replace Package Body b_Comfunc Is

  -----------------------------------------------------------------------------
  -- 功能：保存错误日志
  -----------------------------------------------------------------------------
  Procedure Save_Error_Log
  (
    类型_In     In zlErrorLog.类型%Type,
    错误序号_In In zlErrorLog.错误序号%Type,
    错误信息_In In zlErrorLog.错误信息%Type
  ) Is
  Begin
    Insert Into zlErrorLog
      (会话号, 用户名, 工作站, 时间, 类型, 错误序号, 错误信息)
      Select Sid, User, Machine, Sysdate, 类型_In, 错误序号_In, 错误信息_In
      From V$session
      Where Audsid = Userenv('SessionID');
    Commit;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Save_Error_Log;

  -----------------------------------------------------------------------------
  -- 功能：取可用功能
  -----------------------------------------------------------------------------
  Procedure Get_Usable_Function
  (
    Cursor_Out Out t_Refcur,
    部件_In    In zlPrograms.部件%Type
  ) Is
  Begin
    If Nvl(部件_In, '空空') = '空空' Then
      Open Cursor_Out For
        Select Distinct A.序号, A.标题, A.说明
        From zlPrograms A, zlProgFuncs B, Zlregfunc C
        Where A.系统 = B.系统 And A.序号 = B.序号 And Trunc(A.系统 / 100) = C.系统 And A.序号 = C.序号
        Order By A.序号;
    Else
      Open Cursor_Out For
        Select Distinct A.序号, A.标题, A.说明
        From zlPrograms A, zlProgFuncs B, Zlregfunc C
        Where A.系统 = B.系统 And A.序号 = B.序号 And Upper(A.部件) = Upper(部件_In) And Trunc(A.系统 / 100) = C.系统 And
              A.序号 = C.序号
        Order By A.序号;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Usable_Function;

  -----------------------------------------------------------------------------
  -- 功能：取大写金额
  -----------------------------------------------------------------------------    
  Procedure Get_Uppmoney
  (
    Cursor_Out Out t_Refcur,
    金额_In    In Number
  ) Is
  Begin
    Open Cursor_Out For
      Select zlUppMoney(Nvl(金额_In, 0)) As Num From Dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Uppmoney;

  -----------------------------------------------------------------------------
  -- 功能：根据指定的日期、组号、系统判断指定日期的数据是否已转出到后备数据表中
  -----------------------------------------------------------------------------    
  Procedure Get_Datamoved
  (
    Cursor_Out  Out t_Refcur,
    组号_In     In zlDataMove.组号%Type,
    系统_In     In zlDataMove.系统%Type,
    上次日期_In In zlDataMove.上次日期%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select 系统, 组号
      From zlDataMove
      Where 组号 = 组号_In And 系统 = 系统_In And 上次日期 > 上次日期_In And 上次日期 Is Not Null;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Datamoved;

  -----------------------------------------------------------------------------
  -- 功能：取系统所有者
  -----------------------------------------------------------------------------
  Procedure Get_Owner
  (
    Cursor_Out Out t_Refcur,
    编号_In    In zlSystems.编号%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select 所有者 From zlSystems Where 编号 = 编号_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Owner;

  -----------------------------------------------------------------------------
  -- 功能：取简码
  -----------------------------------------------------------------------------
  Procedure Get_Spell_Code
  (
    Cursor_Out Out t_Refcur,
    字符串_In  In Varchar2,
    方式_In    In Number := 0
  ) Is
  Begin
    If Nvl(方式_In, 0) = 0 Then
      Open Cursor_Out For
        Select zlSpellCode(字符串_In) As 简码 From Dual;
    Else
      Open Cursor_Out For
        Select zlWbCode(字符串_In) As 简码 From Dual;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Spell_Code;

  -----------------------------------------------------------------------------
  -- 功能：保存运行日志
  -----------------------------------------------------------------------------
  Procedure Save_Diary_Log
  (
    部件名_In   In zlDiaryLog.部件名%Type,
    窗体名_In   In zlDiaryLog.窗体名%Type,
    工作内容_In In zlDiaryLog.工作内容%Type
  ) Is
  Begin
    Insert Into zlDiaryLog
      (会话号, 用户名, 工作站, 部件名, 窗体名, 工作内容, 进入时间)
      Select Sid + Serial#, User, RTrim(LTrim(Replace(Machine, Chr(0), ''))), 部件名_In, 窗体名_In, 工作内容_In, Sysdate
      From V$session
      Where Audsid = Userenv('SessionID');
    Commit;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Save_Diary_Log;

  -----------------------------------------------------------------------------
  -- 功能：更改运行日志
  -- 调用列表：
  -- clsComLib.SaveWinState
  -----------------------------------------------------------------------------
  Procedure Update_Diary_Log
  (
    部件名_In In zlDiaryLog.部件名%Type,
    窗体名_In In zlDiaryLog.窗体名%Type
  ) Is
    Cursor c_Session Is
      Select Sid + Serial# As 会话号, User As 用户名, RTrim(LTrim(Replace(Machine, Chr(0), ''))) As 工作站
      From V$session
      Where Audsid = Userenv('SessionID');
  Begin
    For r_Tmp In c_Session Loop
      Update zlDiaryLog
      Set 退出原因 = 1, 退出时间 = Sysdate
      Where 退出原因 Is Null And 用户名 = r_Tmp.用户名 And 工作站 = r_Tmp.工作站 And 会话号 = r_Tmp.会话号 And
            部件名 = 部件名_In And 窗体名 = 窗体名_In;
    End Loop;
    Commit;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Update_Diary_Log;

  -----------------------------------------------------------------------------
  -- 功能：取固定发布报表和用户发布报表
  -----------------------------------------------------------------------------
  Procedure Get_Report_Menu
  (
    Cursor_Out Out t_Refcur,
    系统_In    In zlPrograms.系统%Type,
    序号_In    In zlPrograms.序号%Type,
    功能_In    In zlReports.功能%Type,
    编号_In    In zlReports.编号%Type
  ) Is
  Begin
    If Nvl(编号_In, '空空') <> '空空' Then
      Open Cursor_Out For
        Select 标志, 系统, 编号, 名称
        From (Select 1 As 标志, A.系统, A.编号, A.名称
               From zlReports A, zlPrograms B
               Where A.系统 = B.系统 And A.程序id = B.序号 And Not Upper(A.编号) Like '%BILL%' And
                     Upper(B.部件) <> Upper('zl9Report') And B.系统 = 系统_In And B.序号 = 序号_In And
                     Instr(功能_In, ';' || A.功能 || ';') > 0
               Union All
               Select Decode(A.系统, Null, 2, 1) As 标志, A.系统, A.编号, A.名称
               From zlReports A, zlRPTPuts B, zlPrograms C
               Where A.ID = B.报表id And B.系统 = C.系统 And B.程序id = C.序号 And
                     (Not Upper(A.编号) Like '%BILL%' Or A.系统 Is Null) And Instr(功能_In, ';' || B.功能 || ';') > 0 And
                     C.系统 = 系统_In And C.序号 = 序号_In)
        Where Instr(编号_In, ',' || 编号 || ',') = 0
        Order By 标志, 编号;
    Else
      Open Cursor_Out For
        Select 标志, 系统, 编号, 名称
        From (Select 1 As 标志, A.系统, A.编号, A.名称
               From zlReports A, zlPrograms B
               Where A.系统 = B.系统 And A.程序id = B.序号 And Not Upper(A.编号) Like '%BILL%' And
                     Upper(B.部件) <> Upper('zl9Report') And B.系统 = 系统_In And B.序号 = 序号_In And
                     Instr(功能_In, ';' || A.功能 || ';') > 0
               Union All
               Select Decode(A.系统, Null, 2, 1) As 标志, A.系统, A.编号, A.名称
               From zlReports A, zlRPTPuts B, zlPrograms C
               Where A.ID = B.报表id And B.系统 = C.系统 And B.程序id = C.序号 And
                     (Not Upper(A.编号) Like '%BILL%' Or A.系统 Is Null) And Instr(功能_In, ';' || B.功能 || ';') > 0 And
                     C.系统 = 系统_In And C.序号 = 序号_In)
        Order By 标志, 编号;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Report_Menu;

  -----------------------------------------------------------------------------
  -- 功能：取用户提醒信息
  -----------------------------------------------------------------------------
  Procedure Get_Zlnoticerec
  (
    Cursor_Out Out t_Refcur,
    用户名_In  In zlNoticeRec.用户名%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select A.序号, A.系统, C.程序id As 模块,C.系统 As 报表系统, B.提醒内容 As 结果内容, C.名称 As 提醒报表, A.提醒声音, B.检查时间,
             B.已读标志
      From zlNotices A, zlNoticeRec B, (Select * From zlReports Where 发布时间 Is Not Null) C
      Where B.用户名 = 用户名_In And B.提醒标志 > 0 And C.编号(+) = A.提醒报表 And A.序号 = B.提醒序号 And
            B.提醒内容 Is Not Null;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Zlnoticerec;

  -----------------------------------------------------------------------------
  -- 功能：取邮件正文
  -----------------------------------------------------------------------------
  Procedure Get_Zlmessage
  (
    Cursor_Out Out t_Refcur,
    Id_In      In zlMessages.ID%Type,
    类型_In    In zlMsgState.类型%Type,
    用户_In    In zlMsgState.用户%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select A.*, B.删除, B.状态
      From zlMessages A, zlMsgState B
      Where A.ID = B.消息id And B.消息id = Id_In And B.类型 = 类型_In And B.用户 = 用户_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Zlmessage;

  -----------------------------------------------------------------------------
  -- 功能：取邮件内容
  -----------------------------------------------------------------------------
  Procedure Get_Zlmessage
  (
    Cursor_Out Out t_Refcur,
    Id_In      In zlMessages.ID%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select 内容, 背景色 From zlMessages Where ID = Id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Zlmessage;

  -----------------------------------------------------------------------------
  -- 功能：取邮递地址
  -----------------------------------------------------------------------------
  Procedure Get_Zlmsgstate
  (
    Cursor_Out Out t_Refcur,
    消息id_In  In zlMsgState.消息id%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select 类型, 用户, 身份 From zlMsgState Where 消息id = 消息id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Zlmsgstate;

  -----------------------------------------------------------------------------
  -- 功能：删除消息
  -----------------------------------------------------------------------------
  Procedure Delete_Zlmsgstate
  (
    删除_In   In zlMsgState.删除%Type,
    消息id_In In zlMsgState.消息id%Type,
    类型_In   In zlMsgState.类型%Type,
    用户_In   In zlMsgState.用户%Type
  ) Is
    n_总数 Number(10);
    n_数量 Number(10);
  Begin
    If Nvl(删除_In, 0) = 1 Then
      Update zlMsgState Set 删除 = 1 Where 消息id = 消息id_In And 类型 = 类型_In And 用户 = 用户_In;
    Else
      If 类型_In = 0 Then
        -- 对于草稿，把收件人的也一并删除
        Update zlMsgState Set 删除 = 2 Where 消息id = 消息id_In And 用户 = 用户_In;
      Else
        Update zlMsgState Set 删除 = 2 Where 消息id = 消息id_In And 类型 = 类型_In And 用户 = 用户_In;
      End If;
      --  删除指定ID的消息  mnuEditDelete_Click 调用
      Select Count(*) As 总数, Sum(Decode(删除, 2, 1, 0)) As 数量
      Into n_总数, n_数量
      From zlMsgState
      Where 消息id = 消息id_In;
    
      If n_总数 = n_数量 Then
        Delete From zlMessages Where ID = 消息id_In;
      End If;
    End If;
    Commit;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Delete_Zlmsgstate;

  -----------------------------------------------------------------------------
  -- 功能：删除过期消息
  -----------------------------------------------------------------------------
  Procedure Delete_Zlmessage Is
    n_Days Number;
  Begin
    Select Nvl(参数值, 缺省值) Into n_Days From zlOptions Where 参数号 = 5;
    If Nvl(n_Days, 0) > 0 Then
      Delete From zlMessages Where 时间 < Sysdate - n_Days;
      Commit;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Delete_Zlmessage;

  -----------------------------------------------------------------------------
  -- 功能：取邮件列表
  -----------------------------------------------------------------------------
  Procedure Get_Mail_List
  (
    Cursor_Out  Out t_Refcur,
    消息类型_In In Varchar2,
    用户_In     In zlMsgState.用户%Type,
    显示已读_In In Number,
    会话id_In   In zlMessages.会话id%Type
  ) Is
    v_Sql  Varchar2(1000);
    v_已读 Varchar2(100);
    v_类型 Varchar2(100);
  Begin
  
    If Nvl(显示已读_In, 0) = 1 Then
      v_已读 := ' and substr(S.状态,1,1)=''0''';
    Else
      v_已读 := '';
    End If;
  
    If Instr(';草稿;收件箱;已发送消息;已删除消息;相关消息;', ';' || 消息类型_In || ';') <= 0 Then
      v_类型 := '草稿';
    Else
      v_类型 := 消息类型_In;
    End If;
  
    If v_类型 = '草稿' Then
      v_Sql := 'Select M.ID, M.会话id, M.发件人, M.收件人, M.主题, To_Char(M.时间, ''YYYY-MM-DD HH24:MI:SS'') As 时间, S.类型, S.状态 
              From zlMessages M, zlMsgState S
              Where M.ID = S.消息id  and S.删除=0 and S.用户= ''' || 用户_In || ''' And S.类型=0 ' || v_已读;
    End If;
  
    If v_类型 = '收件箱' Then
      v_Sql := 'Select M.ID, M.会话id, M.发件人, M.收件人, M.主题, To_Char(M.时间, ''YYYY-MM-DD HH24:MI:SS'') As 时间, S.类型, S.状态 
              From zlMessages M, zlMsgState S
              Where M.ID = S.消息id  and S.删除=0 and S.用户= ''' || 用户_In || ''' And S.类型=2 ' || v_已读;
    End If;
  
    If v_类型 = '已发送消息' Then
      v_Sql := 'Select M.ID, M.会话id, M.发件人, M.收件人, M.主题, To_Char(M.时间, ''YYYY-MM-DD HH24:MI:SS'') As 时间, S.类型, S.状态 
              From zlMessages M, zlMsgState S
              Where M.ID = S.消息id  and S.删除=0 and S.用户= ''' || 用户_In || ''' And S.类型=1 ' || v_已读;
    End If;
  
    If v_类型 = '已删除消息' Then
      v_Sql := 'Select M.ID, M.会话id, M.发件人, M.收件人, M.主题, To_Char(M.时间, ''YYYY-MM-DD HH24:MI:SS'') As 时间, S.类型, S.状态 
              From zlMessages M, zlMsgState S
              Where M.ID = S.消息id  and S.用户= ''' || 用户_In || ''' And S.删除=1 ' || v_已读;
    End If;
  
    If v_类型 = '相关消息' Then
      v_Sql := 'select M.ID,M.会话ID,M.发件人,M.收件人,M.主题,to_char(M.时间,''YYYY-MM-DD HH24:MI:SS'') as 时间,S.类型,S.状态
         from zlMessages M,zlMsgState S where M.ID=S.消息ID and S.删除<>2 and S.用户= ''' || 用户_In ||
               '''  and M.会话ID=' || 会话id_In;
    End If;
  
    If Nvl(v_Sql, '空空') <> '空空' Then
      Open Cursor_Out For v_Sql;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Mail_List;

  -----------------------------------------------------------------------------
  -- 功能：还原删除的消息
  -----------------------------------------------------------------------------
  Procedure Restore_Zlmsgstate
  (
    消息id_In In zlMsgState.消息id%Type,
    类型_In   In zlMsgState.类型%Type,
    用户_In   In zlMsgState.用户%Type
  ) Is
  Begin
    Update zlMsgState Set 删除 = 0 Where 消息id = 消息id_In And 类型 = 类型_In And 用户 = 用户_In;
    Commit;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Restore_Zlmsgstate;

  -----------------------------------------------------------------------------
  -- 功能：保存消息
  -- 调用列表：
  -- zlApptools.frmMessageEdit.SaveMessage
  -----------------------------------------------------------------------------
  Procedure Save_Zlmessage
  (
    Cursor_Out Out t_Refcur,
    Id_In      In zlMessages.ID%Type,
    会话id_In  In zlMessages.会话id%Type,
    发件人_In  In zlMessages.发件人%Type,
    收件人_In  In zlMessages.收件人%Type,
    主题_In    In zlMessages.主题%Type,
    内容_In    In zlMessages.内容%Type,
    背景色_In  In zlMessages.背景色%Type
  ) Is
    n_Id     zlMessages.ID%Type;
    n_会话id zlMessages.会话id%Type;
  Begin
    If Nvl(Id_In, 0) = 0 Then
      Select Zlmessages_Id.Nextval Into n_Id From Dual;
      n_Id := Nvl(n_Id, 0);
      If Nvl(会话id_In, 0) = 0 Then
        n_会话id := n_Id;
      Else
        n_会话id := 会话id_In;
      End If;
      Insert Into zlMessages
        (ID, 会话id, 发件人, 时间, 收件人, 主题, 内容, 背景色)
      Values
        (n_Id, n_会话id, 发件人_In, Sysdate, 收件人_In, 主题_In, 内容_In, 背景色_In);
      Open Cursor_Out For
        Select n_Id As ID, n_会话id As 会话id From Dual;
    Else
      Update zlMessages
      Set 发件人 = 发件人_In, 时间 = Sysdate, 收件人 = 收件人_In, 主题 = 主题_In, 内容 = 内容_In, 背景色 = 背景色_In
      Where ID = Id_In;
      Open Cursor_Out For
        Select Id_In As ID, 会话id_In As 会话id From Dual;
    End If;
    Commit;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Save_Zlmessage;

  -----------------------------------------------------------------------------
  -- 功能：插入zlMsgstate
  -- 调用列表：
  -- zlApptools.frmMessageEdit.SaveMessage
  -----------------------------------------------------------------------------
  Procedure Insert_Zlmsgstate
  (
    消息id_In In zlMsgState.消息id%Type,
    类型_In   In zlMsgState.类型%Type,
    用户_In   In zlMsgState.用户%Type,
    身份_In   In zlMsgState.身份%Type,
    删除_In   In zlMsgState.删除%Type,
    状态_In   In zlMsgState.状态%Type
  ) Is
  Begin
  
    If 类型_In < 2 Then
      Delete From zlMsgState Where 消息id = 消息id_In;
    End If;
    Insert Into zlMsgState
      (消息id, 类型, 用户, 身份, 删除, 状态)
    Values
      (消息id_In, 类型_In, 用户_In, 身份_In, 删除_In, 状态_In);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Insert_Zlmsgstate;

  -----------------------------------------------------------------------------
  -- 功能：为原件加上答复或转发标志
  -----------------------------------------------------------------------------
  Procedure Update_Zlmsgstate_State
  (
    模式_In   In Number,
    消息id_In In zlMsgState.消息id%Type,
    类型_In   In zlMsgState.类型%Type,
    用户_In   In zlMsgState.用户%Type
  ) Is
  Begin
    If Nvl(模式_In, 0) = 1 Or Nvl(模式_In, 0) = 2 Then
      Update zlMsgState
      Set 状态 = Substr(状态, 1, 1) || '1' || Substr(状态, 3, 2)
      Where 消息id = 消息id_In And 类型 = 类型_In And 用户 = 用户_In;
      Commit;
    End If;
    If Nvl(模式_In, 0) = 3 Then
      Update zlMsgState
      Set 状态 = Substr(状态, 1, 1) || '1' || Substr(状态, 4, 1)
      Where 消息id = 消息id_In And 类型 = 类型_In And 用户 = 用户_In;
      Commit;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Update_Zlmsgstate_State;

  -----------------------------------------------------------------------------
  -- 功能：更新状态和身份
  -----------------------------------------------------------------------------
  Procedure Update_Zlmsgstate_Idtntify
  (
    身份_In   In zlMsgState.身份%Type,
    消息id_In In zlMsgState.消息id%Type,
    类型_In   In zlMsgState.类型%Type,
    用户_In   In zlMsgState.用户%Type
  ) Is
  Begin
    Update zlMsgState
    Set 状态 = '1' || Substr(状态, 2), 身份 = 身份_In
    Where 消息id = 消息id_In And 类型 = 类型_In And 用户 = 用户_In;
    Commit;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Update_Zlmsgstate_Idtntify;

End b_Comfunc;
/


--参数调整相关更改


Create Or Replace Package zlTools.b_Runmana Is
  -----------------------------------------------------------------------------
  -- 作者： 陈东
  -- 创始时间：2006-6-29
  -- 修改人：
  -- 修改时间：
  -- 描述：
  --         主要用于运行管理功能的过程
  -----------------------------------------------------------------------------

  Type t_Refcur Is Ref Cursor;

  -----------------------------------------------------------------------------
  -- 功能：取参数信息
  -- 修改：刘兴宏
  -- 调用列表：
  -- frmParameters
  -----------------------------------------------------------------------------
  Procedure Get_Parameters
  (
    Cursor_Out Out t_Refcur,
    系统_In    In Number
  );

  -----------------------------------------------------------------------------
  -- 功能：根据指定的参数ID取参数信息
  -- 修改：刘兴宏
  -- 调用列表：
  -- frmParameters;frmParaChangeSet
  -----------------------------------------------------------------------------
  Procedure Get_Parameter
  (
    Cursor_Out Out t_Refcur,
    参数id_In  In zlParameters.id%Type
  );

  -----------------------------------------------------------------------------
  -- 功能：取站点或用户的参数信息
  -- 修改：刘兴宏
  -- 调用列表：
  -- frmParameters
  -----------------------------------------------------------------------------
  Procedure Get_Userparameters
  (
    Cursor_Out Out t_Refcur,
    参数id_In  In zlUserParas.参数id%Type,
    intType    IN NUMBER :=0 
    --0-所有参数信息,1-只获取机器名及简码,2-只获取用户名
  );
 
  -----------------------------------------------------------------------------
  -- 功能：取参数修改信息
  -- 修改：刘兴宏
  -- 调用列表：
  -- frmParameters
  -----------------------------------------------------------------------------
  Procedure Get_Parachangedlog
  (
    Cursor_Out Out t_Refcur,
    参数id_In  In Zlparachangedlog.参数id%Type
  );

  -----------------------------------------------------------------------------
  -- 功能：取ZlAutoJob序列号
  -- 调用列表：
  -- frmAutoJobset.cmdok_click
  -----------------------------------------------------------------------------
  Procedure Get_Job_Number
  (
    Cursor_Out Out t_Refcur,
    系统_In    In Number
  );

  -----------------------------------------------------------------------------
  -- 功能：取ZlDataMove描述
  -- 调用列表：
  -- frmAutoJobset.cmdUpdate_Click
  -----------------------------------------------------------------------------
  Procedure Get_Depict
  (
    Cursor_Out Out t_Refcur,
    系统_In    In zlDataMove.系统%Type,
    组号_In    In zlDataMove.组号%Type
  );

  -----------------------------------------------------------------------------
  -- 功能：取zlClients的MAX IP
  -- 调用列表：
  -- frmClientsEdit.InitCard
  -----------------------------------------------------------------------------
  Procedure Get_Client_Maxip(Cur_Out Out t_Refcur);

  -----------------------------------------------------------------------------
  -- 功能：取zlClients的记录
  -- 调用列表：
  -- frmClientsEdit.InitCard、frmClientsParas.LoadClientsInfor、frmClientsUpgrade.LoadClientsInfor
  -- frmFilesSendToServer.LoadClientsInfor
  -----------------------------------------------------------------------------
  Procedure Get_Client
  (
    Cur_Out   Out t_Refcur,
    工作站_In In zlClients.工作站%Type := Null
  );

  -----------------------------------------------------------------------------
  -- 功能：取zlClients的站点
  -- 调用列表：
  -- frmClientsParasSet.InitCard
  -----------------------------------------------------------------------------
  Procedure Get_Client_Station(Cur_Out Out t_Refcur);

  -----------------------------------------------------------------------------
  -- 功能：取方案号
  -- 调用列表：
  -- frmClientsParasSet.InitCard
  -----------------------------------------------------------------------------
  Procedure Get_Project_No(Cur_Out Out t_Refcur);

  -----------------------------------------------------------------------------
  -- 功能：取方案
  -- 调用列表：
  -- frmClientsParasSet.InitCard
  -----------------------------------------------------------------------------
  Procedure Get_Client_Scheme(Cur_Out Out t_Refcur);

  -----------------------------------------------------------------------------
  -- 功能：取恢复信息
  -- 调用列表：
  -- frmClientsParasSet.Load恢复方案、frmClientsParasSet.LoadScremeSet
  -----------------------------------------------------------------------------
  Procedure Get_Resile
  (
    Cur_Out   Out t_Refcur,
    方案号_In In Zlclientparaset.方案号%Type,
    类型_In   In Number := 0
  );

  -----------------------------------------------------------------------------
  -- 功能：取zldataMove数据
  -- 调用列表：
  -- frmDataMove.cmbSystem_Click
  -----------------------------------------------------------------------------
  Procedure Get_Zldatamove
  (
    Cur_Out Out t_Refcur,
    系统_In In zlDataMove.系统%Type
  );

  -----------------------------------------------------------------------------
  -- 功能：取日志数据
  -- 调用列表：
  -- FrmErrLog.RefreshData、FrmRunLog.RefreshData
  -----------------------------------------------------------------------------
  Procedure Get_Log
  (
    Cur_Out     Out t_Refcur,
    日志类型_In In Varchar2,
    Where_In    In Varchar2
  );

  -----------------------------------------------------------------------------
  -- 功能：取日志记录数
  -- 调用列表：
  -- FrmErrLog.DeleteExtra、FrmRunLog.DeleteExtra
  -----------------------------------------------------------------------------
  Procedure Get_Log_Count
  (
    Cur_Out     Out t_Refcur,
    日志类型_In In Varchar2
  );

  -----------------------------------------------------------------------------
  -- 功能：取zlfilesupgradeg数据
  -- 调用列表：
  -- frmFilesSet.intBillInfor
  -----------------------------------------------------------------------------
  Procedure Get_Zlfilesupgrade(Cur_Out Out t_Refcur);

  -----------------------------------------------------------------------------
  -- 功能：取非注册项目
  -- 调用列表：
  -- frmRegist.Form_Load
  -----------------------------------------------------------------------------
  Procedure Get_Not_Regist(Cur_Out Out t_Refcur);

  -----------------------------------------------------------------------------
  -- 功能：取参数值
  -- 调用列表：
  -- FrmRunOption.InitCons
  -----------------------------------------------------------------------------
  Procedure Get_Zloption
  (
    Cur_Out   Out t_Refcur,
    参数号_In In zlOptions.参数号%Type
  );

End b_Runmana;
/

Create Or Replace Package Body zltools.b_Runmana Is

  -----------------------------------------------------------------------------
  -- 功能：取参数信息
  -- 修改：刘兴宏
  -- 调用列表：
  -- frmParameters
  -----------------------------------------------------------------------------
  Procedure Get_Parameters
  (
    Cursor_Out Out t_Refcur,
    系统_In    In Number
  ) Is
  Begin
    If Nvl(系统_In, 0) = 0 Then
      Open Cursor_Out For
        Select A.ID, A.系统, A.模块, A.私有, A.参数号, A.参数名, A.参数值, A.缺省值, A.参数说明, A.本机, A.授权, A.固定,
               B.标题 As 模块名称, zlSpellCode(B.标题) As 模块简码
        From zlParameters A, zlPrograms B
        Where Nvl(A.系统, 0) = 0 And Nvl(A.系统, 0) = B.系统(+) And Nvl(A.模块, 0) = B.序号(+);
    Else
      Open Cursor_Out For
        Select A.ID, A.系统, A.模块, A.私有, A.参数号, A.参数名, A.参数值, A.缺省值, A.参数说明, A.本机, A.授权, A.固定,
               B.标题 As 模块名称, zlSpellCode(B.标题) As 模块简码
        From zlParameters A, zlPrograms B,
             --处理权限部分，只有授权的才能显示
             (Select Distinct F.序号
               From zlProgFuncs F, zlRegFunc R
               Where Trunc(F.系统 / 100) = R.系统 And F.序号 = R.序号 And F.系统 = 系统_In And F.功能 = R.功能 And
                     1 = (Select 1 From Zlregaudit A Where A.项目 = '授权证章')
               Union All
               Select 0 As 序号 From Dual) M
        Where A.系统 = Nvl(系统_In, 0) And Nvl(A.系统, 0) = B.系统(+) And Nvl(A.模块, 0) = B.序号(+) And
              Nvl(A.模块, 0) = M.序号;
    End If;
  End Get_Parameters;

  -----------------------------------------------------------------------------
  -- 功能：根据指定的参数ID取参数信息
  -- 修改：刘兴宏
  -- 调用列表：
  -- frmParameters;frmParaChangeSet
  -----------------------------------------------------------------------------
  Procedure Get_Parameter
  (
    Cursor_Out Out t_Refcur,
    参数id_In  In zlParameters.ID%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select A.ID, A.系统, A.模块, A.私有, A.参数号, A.参数名, A.参数值, A.缺省值, A.参数说明, A.本机, A.授权, A.固定,
             B.标题 As 模块名称, zlSpellCode(B.标题) As 模块简码
      From zlParameters A, zlPrograms B
      Where A.ID = Nvl(参数id_In, 0) And Nvl(A.系统, 0) = B.系统(+) And Nvl(A.模块, 0) = B.序号(+);
  End Get_Parameter;

  -----------------------------------------------------------------------------
  -- 功能：取站点或用户的参数信息
  -- 修改：刘兴宏
  -- 调用列表：
  -- frmParameters
  -----------------------------------------------------------------------------
  Procedure Get_Userparameters
  (
    Cursor_Out Out t_Refcur,
    参数id_In  In zlUserParas.参数id%Type,
    Inttype    In Number := 0
    --0-所有参数信息,1-只获取机器名及简码,2-只获取用户名
  ) Is
    n_私有 zlParameters.私有%Type;
    n_本机 zlParameters.本机%Type;
  Begin
    If Inttype = 0 Then
      Begin
        Select Nvl(A.私有, 0), Nvl(A.本机, 0) Into n_私有, n_本机 From zlParameters A Where ID = Nvl(参数id_In, 0);
      Exception
        When Others Then
          n_私有 := 0;
          n_本机 := 0;
      End;
      If n_本机 = 1 Then
        --分机器
        If n_私有 = 1 Then
          --本机私有模块
          Open Cursor_Out For
            Select 参数id, 用户名, 参数值, 机器名, zlSpellCode(机器名) As 机器名简码
            From zlUserParas
            Where 参数id = Nvl(参数id_In, 0) And 用户名 Is Not Null And 机器名 Is Not Null;
        Else
          --本机公共模块
          Open Cursor_Out For
            Select 参数id, 用户名, 参数值, 机器名, zlSpellCode(机器名) As 机器名简码
            From zlUserParas
            Where 参数id = Nvl(参数id_In, 0) And 用户名 Is Null And 机器名 Is Not Null;
        End If;
      Else
        If n_私有 = 1 Then
          --私有模块或私有全局
          Open Cursor_Out For
            Select 参数id, 用户名, 参数值, 机器名, zlSpellCode(机器名) As 机器名简码
            From zlUserParas
            Where 参数id = Nvl(参数id_In, 0) And 用户名 Is Not Null And 机器名 Is Null;
        Else
          --公共模块和公共全局,不存在相关的数据
          Open Cursor_Out For
            Select 参数id, 用户名, 参数值, 机器名, '' As 机器名简码
            From zlUserParas
            Where 参数id = Nvl(参数id_In, 0) And 1 = 2;
        End If;
      End If;
    Elsif Inttype = 1 Then
      --只获取机器名及简码,
      Open Cursor_Out For
        Select Distinct 机器名, zlSpellCode(机器名) As 机器名简码
        From zlUserParas
        Where 参数id = Nvl(参数id_In, 0) And 机器名 Is Not Null;
    Elsif Inttype = 2 Then
      --只获取用户名
      Open Cursor_Out For
        Select Distinct 用户名 From zlUserParas Where 参数id = Nvl(参数id_In, 0) And 用户名 Is Not Null;
    Else
      Open Cursor_Out For
        Select 参数id, 用户名, 参数值, 机器名, zlSpellCode(机器名) As 机器名简码
        From zlUserParas
        Where 参数id = Nvl(参数id_In, 0);
    End If;
  End Get_Userparameters;

  -----------------------------------------------------------------------------
  -- 功能：取参数修改信息
  -- 修改：刘兴宏
  -- 调用列表：
  -- frmParameters
  -----------------------------------------------------------------------------
  Procedure Get_Parachangedlog
  (
    Cursor_Out Out t_Refcur,
    参数id_In  In Zlparachangedlog.参数id%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select 参数id, 序号, 变动说明, 变动内容, 变动人, 变动时间, 变动原因
      From Zlparachangedlog
      Where 参数id = Nvl(参数id_In, 0);
  
  End;

  -----------------------------------------------------------------------------
  -- 功能：取ZlAutoJob序列号
  -----------------------------------------------------------------------------
  Procedure Get_Job_Number
  (
    Cursor_Out Out t_Refcur,
    系统_In    In Number
  ) Is
  Begin
    Open Cursor_Out For
      Select 序号 + 1 As 序号
      From zlAutoJobs
      Where Nvl(系统, 0) = 系统_In And 类型 = 3 And
            序号 + 1 Not In (Select 序号 From zlAutoJobs Where Nvl(系统, 0) = 系统_In And 类型 = 3);
  End Get_Job_Number;

  -----------------------------------------------------------------------------
  -- 功能：取ZlDataMove描述
  -----------------------------------------------------------------------------
  Procedure Get_Depict
  (
    Cursor_Out Out t_Refcur,
    系统_In    In zlDataMove.系统%Type,
    组号_In    In zlDataMove.组号%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select 转出描述 From zlDataMove Where Nvl(系统, 0) = 系统_In And 组号 = 组号_In;
  End Get_Depict;

  -----------------------------------------------------------------------------
  -- 功能：取zlClients的MAX IP
  -----------------------------------------------------------------------------
  Procedure Get_Client_Maxip(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select Max(Ip) As Ip From zlClients;
  End Get_Client_Maxip;

  -----------------------------------------------------------------------------
  -- 功能：取zlClients的记录
  -----------------------------------------------------------------------------
  Procedure Get_Client
  (
    Cur_Out   Out t_Refcur,
    工作站_In In zlClients.工作站%Type := Null
  ) Is
    v_Sql Varchar2(1000);
  Begin
    If Nvl(工作站_In, '空') = '空' Then
      v_Sql := 'Select a.Ip, a.工作站, a.Cpu, a.内存, a.硬盘, a.操作系统, a.部门, a.用途, a.说明, a.升级标志, a.禁止使用,
							 a.连接数, Decode(b.Terminal, Null, 0, 1) As 状态, a.收集标志,a.升级服务器
				From Zlclients a, (Select Distinct Terminal From V$session) b
				Where Upper(a.工作站) = Upper(b.Terminal(+))
				Order By a.Ip';
      Open Cur_Out For v_Sql;
    Else
      Open Cur_Out For
        Select Ip, 工作站, Cpu, 内存, 硬盘, 操作系统, 部门, 用途, 说明, 升级标志, 收集标志, 禁止使用, 连接数, 升级服务器
        From zlClients
        Where Upper(工作站) = 工作站_In;
    End If;
  End Get_Client;

  -----------------------------------------------------------------------------
  -- 功能：取zlClients的站点
  -----------------------------------------------------------------------------
  Procedure Get_Client_Station(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select Distinct Upper(工作站) || '[' || Ip || ']' As 站点, Upper(工作站) 工作站 From zlClients;
  End Get_Client_Station;

  -----------------------------------------------------------------------------
  -- 功能：取方案号
  -----------------------------------------------------------------------------
  Procedure Get_Project_No(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select 方案号 From Zlclientparaset Where Rownum = 1;
  End Get_Project_No;

  -----------------------------------------------------------------------------
  -- 功能：取方案
  -----------------------------------------------------------------------------
  Procedure Get_Client_Scheme(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select 方案号, 方案号 || '-' || 方案名称 As 方案名称, 方案描述, 工作站, 用户名 From Zlclientscheme;
  End Get_Client_Scheme;

  -----------------------------------------------------------------------------
  -- 功能：取恢复信息
  -----------------------------------------------------------------------------
  Procedure Get_Resile
  (
    Cur_Out   Out t_Refcur,
    方案号_In In Zlclientparaset.方案号%Type,
    类型_In   In Number := 0
  ) Is
  Begin
    If 类型_In = 0 Then
      Open Cur_Out For
        Select Distinct A.工作站 || Decode(M.工作站, Null, ' ', '[' || M.Ip || ']') As 工作站, A.用户名, A.恢复标志,
                        '[' || B.方案号 || ']' || B.方案名称 As 方案名称
        From Zlclientparaset A, Zlclientscheme B, zlClients M
        Where A.方案号 = B.方案号 And A.工作站 = M.工作站(+) And A.方案号 = 方案号_In;
    End If;
  
    If 类型_In = 1 Then
      Open Cur_Out For
        Select Distinct Upper(工作站) 工作站, Min(恢复标志) 恢复标志
        From Zlclientparaset A
        Where A.方案号 = 方案号_In
        Group By 工作站;
    End If;
  
    If 类型_In = 2 Then
      Open Cur_Out For
        Select Distinct Upper(用户名) 用户名, Max(工作站) 工作站, Min(Decode(恢复标志, 2, 0, 恢复标志)) 恢复标志
        From Zlclientparaset A
        Where A.方案号 = 方案号_In
        Group By 用户名
        Order By 用户名;
    End If;
  
  End Get_Resile;

  -----------------------------------------------------------------------------
  -- 功能：取zldataMove数据
  -----------------------------------------------------------------------------
  Procedure Get_Zldatamove
  (
    Cur_Out Out t_Refcur,
    系统_In In zlDataMove.系统%Type
  ) Is
  Begin
    Open Cur_Out For
      Select 组号, 组名, 说明, 日期字段, 转出描述, 上次日期 From zlDataMove Where 系统 = 系统_In Order By 组号;
  End Get_Zldatamove;

  -----------------------------------------------------------------------------
  -- 功能：取日志数据
  -----------------------------------------------------------------------------
  Procedure Get_Log
  (
    Cur_Out     Out t_Refcur,
    日志类型_In In Varchar2,
    Where_In    In Varchar2
  ) Is
    v_Sql Varchar2(1000);
  Begin
    If 日志类型_In = '错误日志' Then
      v_Sql := 'Select 会话号,工作站,用户名,错误序号,错误信息,To_char(时间,''yyyy-MM-dd hh24:mi:ss'') 时间
					 ,Decode(类型,1,''存储过程错误'',2,''数据联结层错误'',''应用程序层错误'') 错误类型
						From ZlErrorLog Where ' || Where_In;
      Open Cur_Out For v_Sql;
    End If;
    If 日志类型_In = '运行日志' Then
      v_Sql := 'Select 会话号,工作站,用户名,部件名,工作内容,To_char(进入时间,''yyyy-MM-dd hh24:mi:ss'') 进入时间
								 ,To_char(退出时间,''yyyy-MM-dd hh24:mi:ss'') 退出时间,Decode(退出原因,1,''正常退出'',''异常退出'') 退出原因
									From ZlDiaryLog Where ' || Where_In;
      Open Cur_Out For v_Sql;
    End If;
  End Get_Log;

  -----------------------------------------------------------------------------
  -- 功能：取日志记录数
  -----------------------------------------------------------------------------
  Procedure Get_Log_Count
  (
    Cur_Out     Out t_Refcur,
    日志类型_In In Varchar2
  ) Is
  Begin
    If 日志类型_In = '错误日志' Then
      Open Cur_Out For
        Select Count(*) 数量
        From zlErrorLog
        Union All
        Select Nvl(To_Number(参数值), 0) From zlOptions Where 参数号 = 4;
    End If;
    If 日志类型_In = '运行日志' Then
      Open Cur_Out For
        Select Count(*) 数量
        From zlDiaryLog
        Union All
        Select Nvl(To_Number(参数值), 0) From zlOptions Where 参数号 = 2;
    
    End If;
  End Get_Log_Count;

  -----------------------------------------------------------------------------
  -- 功能：取zlfilesupgradeg数据
  -----------------------------------------------------------------------------
  Procedure Get_Zlfilesupgrade(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select A.序号, A.文件名, A.版本号, A.修改日期, B.名称 As 说明
      From zlFilesUpgrade A, zlComponent B
      Where Upper(A.文件名) = Upper(B.部件(+))
      Order By A.序号;
  End Get_Zlfilesupgrade;

  -----------------------------------------------------------------------------
  -- 功能：取非注册项目
  -----------------------------------------------------------------------------
  Procedure Get_Not_Regist(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select 项目, 内容
      From zlRegInfo
      Where 项目 Not In ('发行码', '版本号', '服务器目录', '访问用户', '访问密码', '收集目录', '收集类型', '注册码',
             '授权证章', '授权工具', '授权邮戳');
  End Get_Not_Regist;

  -----------------------------------------------------------------------------
  -- 功能：取参数值
  -----------------------------------------------------------------------------
  Procedure Get_Zloption
  (
    Cur_Out   Out t_Refcur,
    参数号_In In zlOptions.参数号%Type
  ) Is
  Begin
    Open Cur_Out For
      Select Nvl(参数值, 缺省值) Option_Value From zlOptions Where 参数号 = 参数号_In;
  End Get_Zloption;

End b_Runmana; 
/


Create Or Replace Procedure zlTools.zl_Parameters_Change
(
  参数id_In   zlParameters.ID%Type,
  私有_In     zlParameters.私有%Type,
  本机_In     zlParameters.本机%Type,
  授权_In     zlParameters.授权%Type,
  变动人_In   Zlparachangedlog.变动人%Type,
  变动原因_In Zlparachangedlog.变动原因%Type
) Is
  v_Temp     Varchar2(200);
  n_模块     zlParameters.模块%Type;
  n_私有     zlParameters.私有%Type;
  n_本机     zlParameters.本机%Type;
  n_授权     zlParameters.授权%Type;
  n_序号     Zlparachangedlog.序号%Type;
  v_变动说明 Zlparachangedlog.变动说明%Type;
  v_变动内容 Zlparachangedlog.变动内容%Type;

  Function Gettype
  (
    模块_In zlParameters.私有%Type,
    私有_In zlParameters.私有%Type,
    本机_In zlParameters.本机%Type
  ) Return Varchar2 Is
  
  Begin
  
    If Nvl(模块_In, 0) = 0 Then
      --不存模块,证明只有两种类型:公共全局和私有全局
      If Nvl(私有_In, 0) = 0 Then
        Return '公共全局';
      End If;
      Return '私有全局';
    End If;
  
    --对模块的处理
    If 本机_In = 0 Then
      --不是本机的情况,只有两种类型:公共模块和私有模块
      If Nvl(私有_In, 0) = 0 Then
        Return '公共模块';
      End If;
      Return '私有模块';
    End If;
    --对本机的模块进行处理也有两种情况:
    If Nvl(私有_In, 0) = 0 Then
      Return '本机公共模块';
    End If;
    Return '本机私有模块';
  Exception
    When Others Then
      Return Null;
  End Gettype;
Begin

  Select Nvl(模块, 0), Nvl(私有, 0), Nvl(本机, 0), Nvl(授权, 0)
  Into n_模块, n_私有, n_本机, n_授权
  From zlParameters
  Where ID = 参数id_In;
  Select Nvl(Max(序号), 0) + 1 Into n_序号 From Zlparachangedlog Where 参数id = 参数id_In;
  --插入数据
  --说明变动说明:比如:私有模块变为公用模块。
  -- 变动内容:说明变动字段的变化情况:比如:私有:1-->0,本机:1-->0
  v_变动说明 := Null;
  v_变动内容 := Null;
  If n_私有 <> Nvl(私有_In, 0) Or n_本机 <> Nvl(本机_In, 0) Then
    --类型发生了改变
    v_Temp     := '从' || Gettype(n_模块, n_私有, n_本机);
    v_Temp     := v_Temp || '变为' || Gettype(n_模块, Nvl(私有_In, 0), Nvl(本机_In, 0));
    v_变动说明 := v_Temp;
    v_Temp     := '';
    If n_私有 <> Nvl(私有_In, 0) Then
      v_Temp := v_Temp || ',私有:' || n_私有 || '-->' || Nvl(私有_In, 0);
    End If;
    If n_私有 <> Nvl(私有_In, 0) Then
      v_Temp := v_Temp || ',本机:' || n_本机 || '-->' || Nvl(本机_In, 0);
    End If;
    v_变动内容 := Substr(v_Temp, 2);
  End If;
  --检查授权发生改变没有
  If n_授权 <> Nvl(授权_In, 0) Then
    If Not v_变动说明 Is Null Then
      v_变动说明 :=v_变动说明|| ',';
    End If;
    If n_授权 = 0 Then
      v_Temp := '不需要授权';
    Else
      v_Temp := '需要授权';
    End If;
    v_变动说明 := Nvl(v_变动说明, '') || '从' || v_Temp || '改为';
    If 授权_In = 0 Then
      v_Temp := '不需要授权';
    Else
      v_Temp := '需要授权';
    End If;
    v_变动说明 := Nvl(v_变动说明, '') || v_Temp;

    If Not v_变动内容 Is Null Then
	v_变动内容:=v_变动内容||',';	
    End If;

    v_变动内容 := Nvl(v_变动内容, '') || '授权:' || n_授权 || '-->' || Nvl(授权_In, 0);
   
  End If;

  Insert Into Zlparachangedlog
    (参数id, 序号, 变动说明, 变动内容, 变动人, 变动时间, 变动原因)
  Values
    (参数id_In, n_序号, v_变动说明, v_变动内容, 变动人_In, Sysdate, 变动原因_In);

  Update zlParameters Set 私有 = Nvl(私有_In, 0), 本机 = Nvl(本机_In, 0), 授权 = Nvl(授权_In, 0) Where ID = 参数id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End  zl_Parameters_Change;
/


grant execute on zlTools.zl_Parameters_Change to Public
/

create public SYNONYM  zl_Parameters_Change   FOR zlTools.zl_Parameters_Change
/
