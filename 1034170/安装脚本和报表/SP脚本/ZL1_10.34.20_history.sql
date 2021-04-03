--[连续升级]1
--[管理工具版本号]10.34.0
--本脚本支持从ZLHIS+ v10.34.10 升级到 v10.34.20
--请以系统所有者登录PLSQL并执行下列脚本
--脚本执行后，请手工升级导出报表
Define n_System=100;
-------------------------------------------------------------------------------
--结构修正部份
------------------------------------------------------------------------------
--82209:刘硕,2015-03-16,手麻记录排序
alter table 病人手麻记录 add 手术次序 number(2);

--82934:冉俊明,2015-04-09,病人预交记录中增加"结算性质"字段，同时修正升级数据
Alter Table 病人预交记录 Add(结算性质 Number(2));

-------------------------------------------------------------------------------
--数据修正部份
------------------------------------------------------------------------------
--82592:冉俊明,2015-04-09,收费财务监控模块性能问题调整预交款数据处理规则，数据修正
--82934:冉俊明,2015-04-09,病人预交记录中增加"结算性质"字段，同时修正升级数据
--耗时说明:该数据修正脚本在15分钟内执行完成，测试环境如下:
--1.硬件环境
--     IBM 3650 M4,cpu E5-2620 2.1GHz,2*6核，32G内存
--     V3700存储,SAS硬盘,10K RPM,Raid 10
--2.软件环境
--     Windows 2008,Oracle 10.2.0.4 64bit
--     日志文件500M/个，Log Buffer设置为500M,PGA为9G,SGA为自动管理，最大25G
--3.数据环境
--     XX医院运行10年的数据
--     住院费用记录1亿条，门诊费用记录2千万条，预交记录1千1百万条
Declare
  --功能：修正使用预交款的记录
  --该游标用于获取使用预交款的结算记录
  Cursor c_结算数据 Is
    Select 结帐id, 操作员编号, 操作员姓名, 收款时间, 缴款组id, 结算性质
    From (With 结算记录 As (Select Distinct 结帐id
                        From 病人预交记录
                        Where 记录性质 In (1, 11) And 结帐id Is Not Null And Nvl(冲预交, 0) <> 0)
           Select /*+ FULL(A)*/ a.Id As 结帐id, Max(a.操作员编号) As 操作员编号, Max(a.操作员姓名) As 操作员姓名, Max(a.收费时间) As 收款时间, Max(a.缴款组id) As 缴款组id,
                  2 As 结算性质
           From 病人结帐记录 A, 结算记录 B
           Where a.Id = b.结帐id
           Group By a.Id
           Union All
           Select /*+ FULL(A)*/ a.结帐id, Max(a.操作员编号), Max(a.操作员姓名), Max(a.登记时间), Max(a.缴款组id), 5
           From 住院费用记录 A, 结算记录 B
           Where a.结帐id = b.结帐id And a.记帐费用 = 0
           Group By a.结帐id
           Union All
           Select /*+ FULL(A)*/ a.结帐id, Max(a.操作员编号), Max(a.操作员姓名), Max(a.登记时间), Max(a.缴款组id),
                  Decode(Mod(Max(记录性质), 10), 1, 3, Mod(Max(记录性质), 10))
           From 门诊费用记录 A, 结算记录 B
           Where a.结帐id = b.结帐id And a.记帐费用 = 0
           Group By a.结帐id);


  Type t_结帐id Is Table Of 病人预交记录.结帐id%Type;
  Type t_收款时间 Is Table Of 病人预交记录.收款时间%Type;
  Type t_操作员编号 Is Table Of 病人预交记录.操作员编号%Type;
  Type t_操作员姓名 Is Table Of 病人预交记录.操作员姓名%Type;
  Type t_缴款组id Is Table Of 病人预交记录.缴款组id%Type;
  Type t_结算性质 Is Table Of 病人预交记录.结算性质%Type;
  c_结帐id     t_结帐id;
  c_收款时间   t_收款时间;
  c_操作员编号 t_操作员编号;
  c_操作员姓名 t_操作员姓名;
  c_缴款组id   t_缴款组id;
  c_结算性质   t_结算性质;
  n_Array_Size Number := 10000; --每批读取一万个结帐ID,多了可能PGA不够
  I            Number(8) := 0; --每修正10万结帐ID提交一次,多了可能Undo不够,少了提交过于频繁
  J            Number(16) := 0;
  v_内容       Zlupgradeconfig.内容%Type;
Begin
  Begin
    Select 内容 Into v_内容 From Zlupgradeconfig Where 项目 = User || '_病人预交记录修正_20150409_1';
  Exception
    When Others Then
      v_内容 := Null;
  End;
  If Nvl(v_内容, 'RJM') = '成功' Then
    --数据已修正成功
    Return;
  End If;

  --备份数据
  If Zl_Checkobject(1, Null, '病人预交记录_20150409_bak') = 0 Then
    Execute Immediate 'Create Table 病人预交记录_20150409_bak As Select * From 病人预交记录';
  End If;

  Update Zlupgradeconfig Set 内容 = Null Where 项目 = User || '_病人预交记录修正_20150409_1';
  If Sql%NotFound Then
    Insert Into Zlupgradeconfig (项目, 内容) Values (User || '_病人预交记录修正_20150409_1', Null);
  End If;

  Open c_结算数据;
  Loop
    Fetch c_结算数据 Bulk Collect
      Into c_结帐id, c_操作员编号, c_操作员姓名, c_收款时间, c_缴款组id, c_结算性质 Limit n_Array_Size;
    Exit When c_结帐id.Count = 0;
  
    --1第二次及之后使用预交款
    Forall K In 1 .. c_结帐id.Count
      Update 病人预交记录 A
      Set 收款时间 = Nvl(c_收款时间(K), 收款时间), 操作员编号 = Nvl(c_操作员编号(K), 操作员编号), 操作员姓名 = Nvl(c_操作员姓名(K), 操作员姓名),
          缴款组id = Nvl(c_缴款组id(K), 缴款组id), 结算性质 = Nvl(c_结算性质(K), 结算性质)
      Where 结帐id = c_结帐id(K) And 记录性质 = 11;
  
    --2第一次使用预交款
    ----2.1新增预交记录
    Forall K In 1 .. c_结帐id.Count
      Insert Into 病人预交记录
        (ID, 记录性质, NO, 实际票号, 记录状态, 病人id, 主页id, 科室id, 缴款单位, 单位开户行, 单位帐号, 摘要, 金额, 结算方式, 结算号码, 收款时间, 操作员编号, 操作员姓名, 冲预交,
         结帐id, 缴款, 找补, 缴款组id, 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 校对标志, 待转出, 结算性质)
        Select 病人预交记录_Id.Nextval, 11, NO, 实际票号, 记录状态, 病人id, 主页id, 科室id, 缴款单位, 单位开户行, 单位帐号, 摘要, Null, 结算方式, 结算号码,
               Nvl(c_收款时间(K), 收款时间), Nvl(c_操作员编号(K), 操作员编号), Nvl(c_操作员姓名(K), 操作员姓名), 冲预交, 结帐id, 缴款, 找补,
               Nvl(c_缴款组id(K), 缴款组id), 预交类别, 卡类别id, 结算卡序号, 卡号, 交易流水号, 交易说明, 合作单位, 结算序号, 校对标志, 待转出, Nvl(c_结算性质(K), 结算性质)
        From 病人预交记录
        Where 结帐id = c_结帐id(K) And 记录性质 = 1 And Nvl(冲预交, 0) <> 0;
  
    ----2.2将原预交记录的冲预交标记为0
    Forall K In 1 .. c_结帐id.Count
      Update 病人预交记录
      Set 冲预交 = 0, 结算性质 = Nvl(c_结算性质(K), 结算性质)
      Where 结帐id = c_结帐id(K) And 记录性质 = 1;
  
    J := J + c_结帐id.Count;
    If I = 10 Then
      Update Zlupgradeconfig Set 内容 = '已处理' || J || '个结帐ID' Where 项目 = User || '_病人预交记录修正_20150409_1';
      Commit;
      I := 0;
    Else
      I := I + 1;
    End If;
  End Loop;
  Update Zlupgradeconfig Set 内容 = '共处理了' || J || '个结帐ID,正在关闭游标' Where 项目 = User || '_病人预交记录修正_20150409_1';
  Commit;
  Close c_结算数据;

  Update Zlupgradeconfig Set 内容 = '成功' Where 项目 = User || '_病人预交记录修正_20150409_1';
  Commit;
End;
/

Declare
  --功能：修正结帐作废记录与预交记录操作员不一致的记录
  --该游标用于获取结帐作废记录与预交记录操作员不一致的结帐作废记录
  Cursor c_结算数据 Is
    Select ID, a.操作员编号, a.操作员姓名, a.收费时间, a.缴款组id
    From 病人结帐记录 A
    Where 记录状态 = 2 And Exists (Select 1 From 病人预交记录 Where 结帐id = a.Id And 操作员姓名 <> a.操作员姓名);

  Type t_结帐id Is Table Of 病人预交记录.结帐id%Type;
  Type t_收款时间 Is Table Of 病人预交记录.收款时间%Type;
  Type t_操作员编号 Is Table Of 病人预交记录.操作员编号%Type;
  Type t_操作员姓名 Is Table Of 病人预交记录.操作员姓名%Type;
  Type t_缴款组id Is Table Of 病人预交记录.缴款组id%Type;
  c_结帐id     t_结帐id;
  c_收款时间   t_收款时间;
  c_操作员编号 t_操作员编号;
  c_操作员姓名 t_操作员姓名;
  c_缴款组id   t_缴款组id;
  n_Array_Size Number := 10000; --每批读取一万个结帐ID,多了可能PGA不够
  I            Number(8) := 0; --每修正10万结帐ID提交一次,多了可能Undo不够,少了提交过于频繁
  v_内容       Zlupgradeconfig.内容%Type;
Begin
  Begin
    Select 内容 Into v_内容 From Zlupgradeconfig Where 项目 = User || '_病人预交记录修正_20150409_2';
  Exception
    When Others Then
      v_内容 := Null;
  End;
  If Nvl(v_内容, 'RJM') = '成功' Then
    --数据已修正成功
    Return;
  End If;

  --备份数据
  If Zl_Checkobject(1, Null, '病人预交记录_20150409_bak') = 0 Then
    Execute Immediate 'Create Table 病人预交记录_20150409_bak As Select * From 病人预交记录';
  End If;

  Update Zlupgradeconfig Set 内容 = Null Where 项目 = User || '_病人预交记录修正_20150409_2';
  If Sql%NotFound Then
    Insert Into Zlupgradeconfig (项目, 内容) Values (User || '_病人预交记录修正_20150409_2', Null);
  End If;

  Open c_结算数据();
  Loop
    Fetch c_结算数据 Bulk Collect
      Into c_结帐id, c_操作员编号, c_操作员姓名, c_收款时间, c_缴款组id Limit n_Array_Size;
    Exit When c_结帐id.Count = 0;
  
    --排除使用预交款记录，因为使用预交款的在前面已修正
    Forall K In 1 .. c_结帐id.Count
      Update 病人预交记录
      Set 收款时间 = Nvl(c_收款时间(K), 收款时间), 操作员编号 = Nvl(c_操作员编号(K), 操作员编号), 操作员姓名 = Nvl(c_操作员姓名(K), 操作员姓名),
          缴款组id = Nvl(c_缴款组id(K), 缴款组id)
      Where 结帐id = c_结帐id(K) And 记录性质 Not In (1, 11);
  
    If I = 10 Then
      Commit;
      I := 0;
    Else
      I := I + 1;
    End If;
  End Loop;
  Close c_结算数据;

  Update Zlupgradeconfig Set 内容 = '成功' Where 项目 = User || '_病人预交记录修正_20150409_2';
  Commit;
End;
/

Declare
  --功能：升级“结算性质”字段
  --预交款填NULL,2-结帐,3-收费,4-挂号,5-就诊卡,6-补充医保结算
  --该游标用于升级结算性质字段,记录性质为1和11的已在前面修正
  Cursor c_结算数据 Is
    Select Rowid, Mod(记录性质, 10) As 结算性质 From 病人预交记录 Where 记录性质 Not In (1, 11);

  Type t_结算性质 Is Table Of 病人预交记录.结算性质%Type;
  c_结算性质 t_结算性质;

  c_Rowid      t_Strlist := t_Strlist();
  n_Array_Size Number := 10000; --每批一万,多了可能PGA不够
  I            Number(8) := 0; --每修正10万条记录提交一次,多了可能Undo不够,少了提交过于频繁
  J            Number(16) := 0;
  v_内容       Zlupgradeconfig.内容%Type;
Begin
  Begin
    Select 内容 Into v_内容 From Zlupgradeconfig Where 项目 = User || '_病人预交记录修正_20150409_3';
  Exception
    When Others Then
      v_内容 := Null;
  End;
  If Nvl(v_内容, 'RJM') = '成功' Then
    --数据已修正成功
    Return;
  End If;

  Update Zlupgradeconfig Set 内容 = Null Where 项目 = User || '_病人预交记录修正_20150409_3';
  If Sql%NotFound Then
    Insert Into Zlupgradeconfig (项目, 内容) Values (User || '_病人预交记录修正_20150409_3', Null);
  End If;

  Open c_结算数据();
  Loop
    Fetch c_结算数据 Bulk Collect
      Into c_Rowid, c_结算性质 Limit n_Array_Size;
    Exit When c_Rowid.Count = 0;
  
    Forall K In 1 .. c_Rowid.Count
      Update 病人预交记录 Set 结算性质 = Nvl(c_结算性质(K), 结算性质) Where Rowid = c_Rowid(K);
  
    J := J + c_Rowid.Count;
    If I = 10 Then
      Update Zlupgradeconfig Set 内容 = '已处理' || J || '个结帐ID' Where 项目 = User || '_病人预交记录修正_20150409_3';
      Commit;
      I := 0;
    Else
      I := I + 1;
    End If;
  End Loop;
  Close c_结算数据;

  Update Zlupgradeconfig Set 内容 = '成功' Where 项目 = User || '_病人预交记录修正_20150409_3';
  Commit;
End;
/



-------------------------------------------------------------------------------
--权限修正部份
-------------------------------------------------------------------------------




-------------------------------------------------------------------------------
--报表修正部分
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--过程修正部份
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------------------------------
--更改历史数据空间系统版本号
-------------------------------------------------------------------------------------------------------
Update zlBakInfo Set 版本号='10.34.20',更新日期=Sysdate Where 系统=&n_System;
Commit;