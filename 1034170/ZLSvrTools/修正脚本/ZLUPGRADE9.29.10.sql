-----------------------------------------------------------------
--为配合产品版本号由9.29升为9.29.10
-----------------------------------------------------------------
--10754
Alter Table zlTools.zlRPTFmts Add(W Number(18),H Number(18),纸张 Number(3),纸向 Number(1),动态纸张 Number(1))
/
ALTER TABLE zlTools.zlRPTFmts ADD CONSTRAINT zlRPTFmts_CK_纸向 Check(纸向 IN(1,2)) PCTFREE 5
/
ALTER TABLE zlTools.zlRPTFmts ADD CONSTRAINT zlRPTFmts_CK_动态纸张 Check(动态纸张 IN(0,1)) PCTFREE 5
/
Begin
	For r_Report IN(Select * From zlReports) Loop
		Update zlRPTFMTs Set W=r_Report.W,H=r_Report.H,纸张=r_Report.纸张,纸向=r_Report.纸向,动态纸张=r_Report.动态纸张 Where 报表ID=r_Report.ID;
	End Loop;
End;
/

Alter Table zlTools.zlReports Drop Column W
/
Alter Table zlTools.zlReports Drop Column H
/
Alter Table zlTools.zlReports Drop Column 纸张
/
Alter Table zlTools.zlReports Drop Column 纸向
/
Alter Table zlTools.zlReports Drop Column 动态纸张
/

Create Or Replace Package Body zlTools.b_Expert Is
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
        Where 编号 Like 'ZL%_REPORT_%' And Not (发布时间 Is Null Or Trunc(发布时间) = To_Date('3000-01-01', 'yyyy-mm-dd')) And
              Nvl(系统, 0) = 0;
    Else
      Open Cursor_Out For
        Select ID, 编号, 名称, 说明
        From zlReports
        Where 编号 Like 'ZL%_REPORT_%' And Not (发布时间 Is Null Or Trunc(发布时间) = To_Date('3000-01-01', 'yyyy-mm-dd')) And
              系统 = 系统_In;
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