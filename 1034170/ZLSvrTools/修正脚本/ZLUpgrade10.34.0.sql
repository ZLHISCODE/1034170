----10.34.0---》9.45.0
--00000:刘硕,2014-10-28,删除ZLRegInfo流状态信息
   Delete From Zlreginfo Where 项目 = '流状态';
--72631:梁唐彬,2014-04-30,报表综合易用性改进
--4.1.  报表关联查询新增表
CREATE TABLE zlTools.zlRPTRelation(
    报表ID NUMBER(18),
    元素ID NUMBER(18),
    关联报表ID	NUMBER(18),
    参数名	VARCHAR2(50),
    参数值来源	VARCHAR2(255)
    )PCTFREE 5
    Cache Storage(Buffer_Pool Keep);
ALTER TABLE zlTools.zlRPTRelation ADD CONSTRAINT zlRPTRelation_PK PRIMARY KEY (报表ID,关联报表ID,元素ID,参数名) USING INDEX PCTFREE 5;
ALTER TABLE zlTools.zlRPTRelation ADD CONSTRAINT zlRPTRelation_FK_元素ID FOREIGN KEY(元素ID) REFERENCES zlRPTItems(ID) ON DELETE CASCADE;
ALTER TABLE zlTools.zlRPTRelation ADD CONSTRAINT zlRPTRelation_FK_关联报表ID FOREIGN KEY(关联报表ID) REFERENCES zlReports(ID) ON DELETE CASCADE;
ALTER TABLE zlTools.zlRPTRelation ADD CONSTRAINT zlRPTRelation_FK_报表ID FOREIGN KEY(报表ID) REFERENCES zlTools.zlReports(ID) ON DELETE CASCADE;
Create Index zlRPTRelation_IX_元素ID ON zlTools.zlRPTRelation(元素ID) PCTFREE 5;
Create Index zlRPTRelation_IX_关联报表ID ON zlTools.zlRPTRelation(关联报表ID) PCTFREE 5;
--4.2.	报表数据源历史记录新增表
CREATE TABLE zlTools.zlRPTSQLsHistory(
    报表ID	NUMBER(18),
    数据源名称	VARCHAR2(20),
    修改人	VARCHAR2(100),
    修改时间	DATE,
    行号	NUMBER(5),
    内容	VARCHAR2(4000)
    )PCTFREE 5;
ALTER TABLE zlTools.zlRPTSQLsHistory ADD CONSTRAINT zlRPTSQLsHistory_PK PRIMARY KEY (报表ID,数据源名称,修改时间,行号) USING INDEX PCTFREE 5;
ALTER TABLE zlTools.zlRPTSQLsHistory ADD CONSTRAINT zlRPTSQLsHistory_FK_报表ID FOREIGN KEY(报表ID) REFERENCES zlTools.zlReports(ID) ON DELETE CASCADE;
--4.3.	Zlreports新增字段：查询开始时间 DATE、查询结束时间 DATE
ALTER TABLE zlTools.zlReports add(禁止开始时间 DATE,禁止结束时间 DATE);
--4.4.	zlrptPars新增字段：锁定 Number(1)
ALTER TABLE zlTools.zlrptPars add(锁定 Number(1));
--4.5.	列特性设置新增表
CREATE TABLE zlTools.zlRPTColProterty(
    报表ID	NUMBER(18),
    元素ID	NUMBER(18),
    条件名称	VARCHAR2(50),
    条件字段	VARCHAR2(255),
    条件关系	VARCHAR2(50),
    条件值	VARCHAR2(255),
    字体颜色	NUMBER(18),
    背景颜色	NUMBER(18),
    是否加粗    NUMBER(1),
    是否整行应用 NUMBER(1)
    )PCTFREE 5
    Cache Storage(Buffer_Pool Keep);
ALTER TABLE zlTools.zlRPTColProterty ADD CONSTRAINT zlRPTColProterty_PK PRIMARY KEY (报表ID,元素ID,条件名称) USING INDEX PCTFREE 5;
ALTER TABLE zlTools.zlRPTColProterty ADD CONSTRAINT zlRPTColProterty_FK_元素ID FOREIGN KEY(元素ID) REFERENCES zlTools.zlRPTItems(ID) ON DELETE CASCADE;
ALTER TABLE zlTools.zlRPTColProterty ADD CONSTRAINT zlRPTColProterty_FK_报表ID FOREIGN KEY(报表ID) REFERENCES zlTools.zlReports(ID) ON DELETE CASCADE;
Create Index zlRPTColProterty_IX_元素ID ON zlTools.zlRPTColProterty(元素ID) PCTFREE 5;

--70590:刘硕,2014-04-30,缺省使用个性化风格
Update zlParameters Set 缺省值 = '1' Where 系统 Is Null And 参数名 = '使用个性化风格';

--74440:刘硕,2014-07-08,模块关联授权
create table ZLTOOLS.ZLModuleRelas
(
系统  Number(5),
模块  Number(18),  
功能  Varchar2(30),
相关系统  Number(5),
相关模块  Number(18),  
相关类型  Number(1), 
相关功能  Varchar2(30),
缺省值    Number(1)
)
tablespace ZLTOOLSTBS;
alter table ZLTOOLS.zlprograms add 性质  Number(1);
alter table ZLTOOLS.ZLModuleRelas Modify 系统  constraint ZLModuleRelas_NN_系统   not  null;
alter table ZLTOOLS.ZLModuleRelas Modify 模块  constraint ZLModuleRelas_NN_模块   not  null;
alter table ZLTOOLS.ZLModuleRelas Modify 相关模块  constraint ZLModuleRelas_NN_相关模块   not  null;
alter table ZLTOOLS.ZLModuleRelas add constraint ZLModuleRelas_UQ_相关模块 Unique(系统,模块,功能,相关系统,相关模块,相关功能) using index tablespace ZLTOOLSTBS;
alter table ZLTOOLS.ZLModuleRelas add constraint ZLModuleRelas_FK_模块 foreign key(系统,模块) references  ZLTOOLS.zlprograms(系统,序号) on delete cascade;

--00000:张永康,2014-03-04,公共函数处理，未登记BUG(2014-6-16更新)
Drop Function zlTools.f_Get_Stream_State;

Create Or Replace Function zlTools.Zl_Checkobject
(
  n_Type        In Number, --1=表,2=字段,3=约束,4=索引
  v_Object_Name In Varchar2,
  v_Table_Name  In Varchar2 := Null --仅当n_Type=2时才需要传入
) Return Number Authid Current_User As
  --功能：以执行者的身份检查指定表的指定对象是否存在
  --返回值：>0表示存在，0表示不存在
  n_Count Number(5);
Begin
  If n_Type = 1 Then
    Select Count(1) Into n_Count From User_Tables Where Table_Name = Upper(v_Table_Name);
  
  Elsif n_Type = 2 Then
    Select Count(1)
    Into n_Count
    From User_Tab_Columns
    Where Table_Name = Upper(v_Table_Name) And Column_Name = Upper(v_Object_Name);
  
  Elsif n_Type = 3 Then
    Select Count(1) Into n_Count From User_Constraints Where Constraint_Name = Upper(v_Object_Name);
  
  Elsif n_Type = 4 Then
    Select Count(1) Into n_Count From User_Indexes Where Index_Name = Upper(v_Object_Name);
  End If;

  Return n_Count;
End Zl_Checkobject;
/


--00000:周韬,2014-03-27,产品特定授权控制改进，未登记问题
Create Or Replace Function zlTools.f_Reg_Menu
(
  Menu_Group_In  In Zlmenus.组别%Type := Null, --本机选择的菜单组别
  System_List_In In Varchar2, --本次会话涉及的应用系统
  Part_List_In   In Varchar2 --以逗号分隔的本机可执行部件列表
) Return t_Menu_Rowset Is
  t_Return t_Menu_Rowset := t_Menu_Rowset();
  t_Middle t_Menu_Rowset := t_Menu_Rowset();

  v_Parts   Varchar2(32767);
  t_Parts   t_Reg_Rowset := t_Reg_Rowset();
  v_Systems Varchar2(32767);
  t_Systems t_Reg_Rowset := t_Reg_Rowset();

Begin
  --变量解析形成类型数组表
  v_Parts := Upper(Part_List_In) || ',';
  While v_Parts Is Not Null Loop
    t_Parts.Extend;
    t_Parts(t_Parts.Count) := t_Reg_Record(Null, Null, Substr(v_Parts, 1, Instr(v_Parts, ',') - 1));
    v_Parts := Trim(Substr(v_Parts, Instr(v_Parts, ',') + 1));
  End Loop;
  t_Parts.Extend;
  t_Parts(t_Parts.Count) := t_Reg_Record(Null, Null, 'ZL9REPORT');
  v_Systems := System_List_In || ',';
  While v_Systems Is Not Null Loop
    t_Systems.Extend;
    t_Systems(t_Systems.Count) := t_Reg_Record(Null, To_Number(Substr(v_Systems, 1, Instr(v_Systems, ',') - 1)), Null);
    v_Systems := Trim(Substr(v_Systems, Instr(v_Systems, ',') + 1));
  End Loop;
  t_Systems.Extend;
  t_Systems(t_Systems.Count) := t_Reg_Record(Null, 0, Null);

  --菜单数据获取：
  Select t_Menu_Record(m.层次, m.Id, m.上级id, m.标题, m.短标题, m.快键, m.说明, m.模块, m.系统, m.图标, p.部件, 0) Bulk Collect
  Into t_Middle
  From (Select Level As 层次, ID, 上级id, 标题, 短标题, 快键, 说明, 模块, 系统, 图标
         From zlMenus
         Where 组别 = Menu_Group_In
         Start With 上级id Is Null
         Connect By Prior ID = 上级id) M,
       (Select Distinct p.系统, p.序号, p.部件
         From zlPrograms P, zlProgFuncs F, zlRegFunc R, zlRPTGroups X, Table(Cast(t_Parts As t_Reg_Rowset)) C,
              Table(Cast(t_Systems As t_Reg_Rowset)) S,
              (Select 1 As 超级 From Sys.Dba_Role_Privs Where Granted_Role = 'DBA' And Grantee = User) A,
              (Select Decode(Count(*), 0, 0, Null, 0, 1) As 编号
                From zlSystems
                Where Upper(所有者) = User
                Union All
                Select 编号
                From zlSystems
                Where Upper(所有者) = User) O,
              (Select Distinct g.系统, g.序号
                From zlRoleGrant G, Sys.Dba_Role_Privs R
                Where g.角色 = r.Granted_Role And r.Grantee = User) G
         Where Nvl(f.系统, 0) = Nvl(p.系统, 0) And f.序号 = p.序号 And Trunc(f.系统 / 100) = r.系统(+) And f.序号 = r.序号(+) And
               f.功能 = r.功能(+) And
               (r.功能 Is Null And f.系统 Is Null Or r.功能 Is Not Null And r.功能 = '基本' Or
                r.功能 Is Not Null And x.程序id Is Not Null Or r.功能 Is Null And (p.序号 Between 10000 And 19999)) And
               p.系统 = x.系统(+) And p.序号 = x.程序id(+) And Upper(p.部件) = c.Text And Nvl(p.系统, 0) = s.Prog And
               p.序号 = p.序号 * a.超级(+) And Nvl(p.系统, 1) = o.编号(+) And Nvl(p.系统, 0) = Nvl(g.系统(+), 0) And p.序号 = g.序号(+) And
               (a.超级 Is Not Null Or o.编号 Is Not Null Or g.序号 Is Not Null)) P
  Where Nvl(m.系统, 0) = Nvl(p.系统(+), 0) And m.模块 = p.序号(+) And (m.模块 Is Null Or m.模块 Is Not Null And p.序号 Is Not Null)
  Order By m.层次 Desc;

  --清理无下级可执行的菜单项目
  For n_Child In 1 .. t_Middle.Count Loop
    If t_Middle(n_Child).部件 Is Not Null Or t_Middle(n_Child).标记 = 1 Then
      t_Return.Extend;
      t_Return(t_Return.Count) := t_Middle(n_Child);
      If t_Middle(n_Child).上级id Is Not Null Then
        For n_Parent In n_Child + 1 .. t_Middle.Count Loop
          If t_Middle(n_Parent).标记 = 0 And t_Middle(n_Parent).Id = t_Middle(n_Child).上级id Then
            t_Middle(n_Parent).标记 := 1;
            Exit;
          End If;
        End Loop;
      End If;
    End If;
  End Loop;

  Return t_Return;
End f_Reg_Menu;
/

--00000:周韬,2014-03-27,产品特定授权控制改进，未登记问题
Create Or Replace Package Body zlTools.b_Popedom Is
  --功能：CopyMenu
  Procedure Copy_Menu
  (
    系统_In   In Zlmenus.系统%Type,
    新系统_In In Zlmenus.系统%Type
  ) Is
    n_Menuid Zlmenus.Id%Type;
  Begin
    Select Max(ID) Into n_Menuid From zlMenus;
    n_Menuid := Nvl(n_Menuid, 0) + 1;
    Insert Into zlMenus
      (组别, ID, 上级id, 标题, 短标题, 快键, 图标, 说明, 模块, 系统)
      Select 组别, ID + n_Menuid, 上级id + n_Menuid, 标题, 短标题, 快键, 图标, 说明, 模块, 新系统_In
      From zlMenus
      Where 系统 = 系统_In;
    Commit;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Copy_Menu;

  --功能：取ZlMenu数据
  Procedure Get_Menu_Tree
  (
    Cursor_Out Out t_Refcur,
    组别_In    In Zlmenus.组别%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select ID, 上级id, 标题, 快键, 说明, 系统, 模块, 短标题, 图标, Level As 级数
      From zlMenus
      Start With 上级id Is Null And 组别 = 组别_In
      Connect By Prior ID = 上级id And 组别 = 组别_In
      Order By Level, ID;
  End Get_Menu_Tree;

  --功能：取ZlMenu数据
  Procedure Get_Menu_Group
  (
    Cursor_Out Out t_Refcur,
    组别_In    In Zlmenus.组别%Type
  ) Is
  Begin
    If 组别_In Is Null Then
      --frmMenu.FillMenuName
      Open Cursor_Out For
        Select Distinct 组别 From zlMenus;
    Else
      --frmMenu.cmdNew_Click
      Open Cursor_Out For
        Select Count(*) As 数量 From zlMenus Where 组别 = 组别_In;
    End If;
  End Get_Menu_Group;

  --功能：取模块
  Procedure Get_Module
  (
    Cursor_Out Out t_Refcur,
    系统_In    In Zlcomponent.系统%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select p.序号, p.标题, c.名称 As 部件
      From zlPrograms P, zlComponent C
      Where Upper(p.部件) = Upper(c.部件) And c.系统 = 系统_In And p.系统 = 系统_In
      Order By c.名称, p.序号;
  End Get_Module;

  --功能：取功能或排列，说明
  Procedure Get_Function
  (
    Cursor_Out Out t_Refcur,
    系统_In    In Zlprogfuncs.系统%Type,
    序号_In    In Zlprogfuncs.序号%Type,
    功能_In    In Zlprogfuncs.功能%Type := Null
  ) Is
  Begin
    If Nvl(功能_In, '空') = '空' Then
      Open Cursor_Out For
        Select 功能 From zlProgFuncs Where 系统 = 系统_In And 序号 = 序号_In Order By Nvl(排列, 0);
    Else
      Open Cursor_Out For
        Select 排列, 说明 From zlProgFuncs Where 系统 = 系统_In And 序号 = 序号_In And 功能 = 功能_In;
    End If;
  End Get_Function;

  --功能：取表权限
  Procedure Get_Impower
  (
    Cursor_Out Out t_Refcur,
    系统_In    In Zlprogprivs.系统%Type,
    序号_In    In Zlprogprivs.序号%Type,
    功能_In    In Zlprogprivs.功能%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select 对象, Sum(Decode(权限, 'SELECT', 1, 0)) As "SELECT", Sum(Decode(权限, 'UPDATE', 1, 0)) As "UPDATE",
             Sum(Decode(权限, 'INSERT', 1, 0)) As "INSERT", Sum(Decode(权限, 'DELETE', 1, 0)) As "DELETE",
             Sum(Decode(权限, 'EXECUTE', 1, 0)) As "EXECUTE"
      From zlProgPrivs
      Where 系统 = 系统_In And 序号 = 序号_In And 功能 = 功能_In
      Group By 对象
      Order By 对象;
  End Get_Impower;

  --功能：得到角色能访问的导航台工具
  Procedure Get_Role_Tools
  (
    Cursor_Out Out t_Refcur,
    角色_In    In Zlrolegrant.角色%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select p.序号, p.标题, p.说明, r.功能
      From zlRoleGrant R, zlPrograms P
      Where r.系统 Is Null And p.序号 = r.序号 And r.角色 = 角色_In And p.系统 Is Null And p.序号 < 100 And p.部件 Is Null
      Order By p.序号;
  End Get_Role_Tools;

  --功能：得到以前的权限
  Procedure Get_Role_Grant
  (
    Curgrand_Out    Out t_Refcur,
    Curprivs_Out    Out t_Refcur,
    Curfuncpars_Out Out t_Refcur,
    角色_In         In Zlrolegrant.角色%Type
  ) Is
  Begin
    Open Curgrand_Out For
      Select Nvl(系统, 0) As 系统, 序号, 功能 From zlRoleGrant Where 角色 = 角色_In;
    Open Curprivs_Out For
      Select Nvl(系统, 0) As 系统, 序号, 功能, 所有者, 权限, 对象 From zlProgPrivs;
    Open Curfuncpars_Out For
      Select p.系统, f.函数名, p.对象
      From zlFuncPars P, zlFunctions F
      Where p.系统 = f.系统 And p.函数号 = f.函数号 And p.对象 Is Not Null;
  End Get_Role_Grant;

  --功能：FillFunc
  Procedure Get_Zlprogfunc
  (
    Cursor_Out Out t_Refcur,
    系统_In    In Zlprogfuncs.系统%Type,
    序号_In    In Zlprogfuncs.序号%Type
  ) Is
  Begin
    If Nvl(系统_In, 0) = 0 Then
      Open Cursor_Out For
        Select 功能, 排列, 说明 From zlProgFuncs Where 系统 Is Null And 序号 = 序号_In And 功能 <> '基本';
    Else
      Open Cursor_Out For
        Select a.功能, a.排列, a.说明
        From zlProgFuncs A, zlRegFunc B
        Where (a.系统 / 100) = b.系统(+) And a.序号 = b.序号(+) And a.功能 = b.功能(+) And
              (b.功能 Is Not Null Or b.功能 Is Null And (a.序号 Between 10000 And 19999)) And a.系统 = 系统_In And a.序号 = 序号_In And
              a.功能 <> '基本';
    End If;
  End Get_Zlprogfunc;

  --功能：是所有角色对应的模块
  Procedure Get_All_Module(Cursor_Out Out t_Refcur) Is
  Begin
    Open Cursor_Out For
      Select a.角色, a.序号, a.功能, b.标题, b.说明
      From zlRoleGrant A, zlPrograms B
      Where a.序号 = b.序号 And Nvl(a.系统, 0) = Nvl(b.系统, 0)
      Order By a.角色, a.序号;
  End Get_All_Module;

End b_Popedom;
/

--00000:周韬,2014-03-27,产品特定授权控制改进，未登记问题
Create Or Replace Package Body zlTools.b_Runmana Is
  --功能：取参数信息
  --frmParameters
  Procedure Get_Parameters
  (
    Cursor_Out Out t_Refcur,
    系统_In    In Number
  ) Is
  Begin
    If Nvl(系统_In, 0) = 0 Then
      Open Cursor_Out For
        Select a.Id, a.系统, a.模块, a.私有, a.参数号, a.参数名, a.参数值, a.缺省值, a.参数说明, a.本机, a.授权, a.固定, b.标题 As 模块名称,
               zlSpellCode(b.标题) As 模块简码
        From zlParameters A, zlPrograms B
        Where Nvl(a.系统, 0) = 0 And Nvl(a.系统, 0) = b.系统(+) And Nvl(a.模块, 0) = b.序号(+);
    Else
      Open Cursor_Out For
        Select a.Id, a.系统, a.模块, a.私有, a.参数号, a.参数名, a.参数值, a.缺省值, a.参数说明, a.本机, a.授权, a.固定, b.标题 As 模块名称,
               zlSpellCode(b.标题) As 模块简码
        From zlParameters A, zlPrograms B,
             --处理权限部分，只有授权的才能显示
             (Select Distinct f.序号
               From zlProgFuncs F, zlRegFunc R
               Where Trunc(f.系统 / 100) = r.系统(+) And f.序号 = r.序号(+) And f.功能 = r.功能(+) And
                     (r.功能 Is Not Null Or r.功能 Is Null And (f.序号 Between 10000 And 19999)) And f.系统 = 系统_In And
                     1 = (Select 1 From Zlregaudit A Where a.项目 = '授权证章')
               Union All
               Select 0 As 序号
               From Dual) M
        Where a.系统 = Nvl(系统_In, 0) And Nvl(a.系统, 0) = b.系统(+) And Nvl(a.模块, 0) = b.序号(+) And Nvl(a.模块, 0) = m.序号;
    End If;
  End Get_Parameters;

  --功能：根据指定的参数ID取参数信息
  --调用列表：frmParameters;frmParaChangeSet
  Procedure Get_Parameter
  (
    Cursor_Out Out t_Refcur,
    参数id_In  In Zlparameters.Id%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select a.Id, a.系统, a.模块, a.私有, a.参数号, a.参数名, a.参数值, a.缺省值, a.参数说明, a.本机, a.授权, a.固定, b.标题 As 模块名称,
             zlSpellCode(b.标题) As 模块简码
      From zlParameters A, zlPrograms B
      Where a.Id = Nvl(参数id_In, 0) And Nvl(a.系统, 0) = b.系统(+) And Nvl(a.模块, 0) = b.序号(+);
  End Get_Parameter;

  --功能：取站点或用户的参数信息
  --调用列表：frmParameters
  Procedure Get_Userparameters
  (
    Cursor_Out Out t_Refcur,
    参数id_In  In Zluserparas.参数id%Type,
    Inttype    In Number := 0
    --0-所有参数信息,1-只获取机器名及简码,2-只获取用户名
  ) Is
    n_私有 Zlparameters.私有%Type;
    n_本机 Zlparameters.本机%Type;
  Begin
    If Inttype = 0 Then
      Begin
        Select Nvl(a.私有, 0), Nvl(a.本机, 0) Into n_私有, n_本机 From zlParameters A Where ID = Nvl(参数id_In, 0);
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

  --功能：取参数修改信息
  --调用列表：frmParameters
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
  --功能：取ZlAutoJob序列号
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

  --功能：取ZlDataMove描述
  Procedure Get_Depict
  (
    Cursor_Out Out t_Refcur,
    系统_In    In Zldatamove.系统%Type,
    组号_In    In Zldatamove.组号%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select 转出描述 From zlDataMove Where Nvl(系统, 0) = 系统_In And 组号 = 组号_In;
  End Get_Depict;

  --功能：取zlClients的MAX IP
  Procedure Get_Client_Maxip(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select Max(Ip) As Ip From zlClients;
  End Get_Client_Maxip;

  --功能：取zlClients的记录
  Procedure Get_Client
  (
    Cur_Out   Out t_Refcur,
    工作站_In In Zlclients.工作站%Type := Null
  ) Is
    v_Sql Varchar2(1000);
  Begin
    If Nvl(工作站_In, '空') = '空' Then
      v_Sql := 'Select a.Ip, a.工作站, a.Cpu, a.内存, a.硬盘, a.操作系统, a.部门, a.用途, a.说明, a.升级标志, a.禁止使用,
                             a.连接数, Decode(b.Terminal, Null, 0, 1) As 状态, a.收集标志,a.升级服务器,a.站点,a.启用视频源
                From Zlclients a, (Select Distinct Terminal From V$session) b
                Where Upper(a.工作站) = Upper(b.Terminal(+))
                Order By a.Ip';
      Open Cur_Out For v_Sql;
    Else
      Open Cur_Out For
        Select Ip, 工作站, Cpu, 内存, 硬盘, 操作系统, 部门, 用途, 说明, 升级标志, 收集标志, 禁止使用, 连接数, 升级服务器, 站点, 启用视频源
        From zlClients
        Where Upper(工作站) = 工作站_In;
    End If;
  End Get_Client;

  --功能：取zlClients的站点
  Procedure Get_Client_Station(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select Distinct Upper(工作站) || '[' || Ip || ']' As 站点, Upper(工作站) 工作站 From zlClients;
  End Get_Client_Station;

  --功能：取方案号
  Procedure Get_Project_No(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select 方案号 From Zlclientparaset Where Rownum = 1;
  End Get_Project_No;

  --功能：取方案
  Procedure Get_Client_Scheme(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select 方案号, 方案号 || '-' || 方案名称 As 方案名称, 方案描述, 工作站, 用户名 From Zlclientscheme;
  End Get_Client_Scheme;

  --功能：取恢复信息
  Procedure Get_Resile
  (
    Cur_Out   Out t_Refcur,
    方案号_In In Zlclientparaset.方案号%Type,
    类型_In   In Number := 0
  ) Is
  Begin
    If 类型_In = 0 Then
      Open Cur_Out For
        Select Distinct a.工作站 || Decode(m.工作站, Null, ' ', '[' || m.Ip || ']') As 工作站, a.用户名, a.恢复标志,
                        '[' || b.方案号 || ']' || b.方案名称 As 方案名称
        From Zlclientparaset A, Zlclientscheme B, zlClients M
        Where a.方案号 = b.方案号 And a.工作站 = m.工作站(+) And a.方案号 = 方案号_In;
    End If;
  
    If 类型_In = 1 Then
      Open Cur_Out For
        Select Distinct Upper(工作站) 工作站, Min(恢复标志) 恢复标志
        From Zlclientparaset A
        Where a.方案号 = 方案号_In
        Group By 工作站;
    End If;
  
    If 类型_In = 2 Then
      Open Cur_Out For
        Select Distinct Upper(用户名) 用户名, Max(工作站) 工作站, Min(Decode(恢复标志, 2, 0, 恢复标志)) 恢复标志
        From Zlclientparaset A
        Where a.方案号 = 方案号_In
        Group By 用户名
        Order By 用户名;
    End If;
  
  End Get_Resile;

  --功能：取zldataMove数据
  Procedure Get_Zldatamove
  (
    Cur_Out Out t_Refcur,
    系统_In In Zldatamove.系统%Type
  ) Is
  Begin
    Open Cur_Out For
      Select 组号, 组名, 说明, 日期字段, 转出描述, 上次日期 From zlDataMove Where 系统 = 系统_In Order By 组号;
  End Get_Zldatamove;

  --功能：取日志数据
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
                     ,Decode(类型,1,''存储过程错误'',2,''数据联结层错误'',3,''应用程序层错误'',''客户端升级错误'') 错误类型
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

  --功能：取日志记录数
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
        Select Nvl(To_Number(参数值), 0)
        From zlOptions
        Where 参数号 = 4;
    End If;
    If 日志类型_In = '运行日志' Then
      Open Cur_Out For
        Select Count(*) 数量
        From zlDiaryLog
        Union All
        Select Nvl(To_Number(参数值), 0)
        From zlOptions
        Where 参数号 = 2;
    
    End If;
  End Get_Log_Count;

  --功能：取zlfilesupgradeg数据
  Procedure Get_Zlfilesupgrade(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select 序号, 文件名, 版本号, 修改日期, 文件说明 As 说明,
             Decode(文件类型, 0, '公共部件', 1, '应用部件', 2, '帮助文件', 3, '其它文件', 4, '三方部件', 5, '系统文件', '') As 类型, 安装路径 As 安装路径,
             Md5 As Md5, 加入日期
      From zlFilesUpgrade
      Order By 序号;
  End Get_Zlfilesupgrade;

  --功能：取非注册项目
  Procedure Get_Not_Regist(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select 项目, 内容
      From zlRegInfo
      Where 项目 Not In ('发行码', '版本号', '服务器目录', '访问用户', '访问密码', '收集目录', '收集类型', '注册码', '授权证章', '授权工具', '授权邮戳');
  End Get_Not_Regist;

  --功能：取参数值
  Procedure Get_Zloption
  (
    Cur_Out   Out t_Refcur,
    参数号_In In Zloptions.参数号%Type
  ) Is
  Begin
    Open Cur_Out For
      Select Nvl(参数值, 缺省值) Option_Value From zlOptions Where 参数号 = 参数号_In;
  End Get_Zloption;

End b_Runmana;
/

--00000:周韬,2014-03-27,产品特定授权控制改进，未登记问题
Create Or Replace Package Body zlTools.b_Comfunc Is
  --功能：保存错误日志
  Procedure Save_Error_Log
  (
    类型_In     In Zlerrorlog.类型%Type,
    错误序号_In In Zlerrorlog.错误序号%Type,
    错误信息_In In Zlerrorlog.错误信息%Type
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
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Save_Error_Log;

  --功能：取可用功能
  Procedure Get_Usable_Function
  (
    Cursor_Out Out t_Refcur,
    部件_In    In Zlprograms.部件%Type
  ) Is
  Begin
    If Nvl(部件_In, '空空') = '空空' Then
      Open Cursor_Out For
        Select Distinct a.序号, a.标题, a.说明
        From zlPrograms A, zlProgFuncs B, zlRegFunc C
        Where a.系统 = b.系统 And a.序号 = b.序号 And Trunc(b.系统 / 100) = c.系统(+) And b.序号 = c.序号(+) And b.功能 = c.功能(+) And
              (c.功能 Is Not Null Or c.功能 Is Null And (a.序号 Between 10000 And 19999))
        Order By a.序号;
    Else
      Open Cursor_Out For
        Select Distinct a.序号, a.标题, a.说明
        From zlPrograms A, zlProgFuncs B, zlRegFunc C
        Where a.系统 = b.系统 And a.序号 = b.序号 And Upper(a.部件) = Upper(部件_In) And Trunc(b.系统 / 100) = c.系统(+) And
              b.序号 = c.序号(+) And b.功能 = c.功能(+) And
              (c.功能 Is Not Null Or c.功能 Is Null And (a.序号 Between 10000 And 19999))
        Order By a.序号;
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Usable_Function;

  --功能：取大写金额    
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
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Uppmoney;

  --功能：根据指定的日期、组号、系统判断指定日期的数据是否已转出到后备数据表中    
  Procedure Get_Datamoved
  (
    Cursor_Out  Out t_Refcur,
    组号_In     In Zldatamove.组号%Type,
    系统_In     In Zldatamove.系统%Type,
    上次日期_In In Zldatamove.上次日期%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select 系统, 组号
      From zlDataMove
      Where 组号 = 组号_In And 系统 = 系统_In And 上次日期 > 上次日期_In And 上次日期 Is Not Null;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Datamoved;

  --功能：取系统所有者
  Procedure Get_Owner
  (
    Cursor_Out Out t_Refcur,
    编号_In    In Zlsystems.编号%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select 所有者 From zlSystems Where 编号 = 编号_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Owner;

  --功能：取简码
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
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Spell_Code;

  --功能：保存运行日志
  Procedure Save_Diary_Log
  (
    部件名_In   In Zldiarylog.部件名%Type,
    窗体名_In   In Zldiarylog.窗体名%Type,
    工作内容_In In Zldiarylog.工作内容%Type
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
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Save_Diary_Log;

  --功能：更改运行日志
  --调用列表：clsComLib.SaveWinState
  Procedure Update_Diary_Log
  (
    部件名_In In Zldiarylog.部件名%Type,
    窗体名_In In Zldiarylog.窗体名%Type
  ) Is
    Cursor c_Session Is
      Select Sid + Serial# As 会话号, User As 用户名, RTrim(LTrim(Replace(Machine, Chr(0), ''))) As 工作站
      From V$session
      Where Audsid = Userenv('SessionID');
  Begin
    For r_Tmp In c_Session Loop
      Update zlDiaryLog
      Set 退出原因 = 1, 退出时间 = Sysdate
      Where 退出原因 Is Null And 用户名 = r_Tmp.用户名 And 工作站 = r_Tmp.工作站 And 会话号 = r_Tmp.会话号 And 部件名 = 部件名_In And 窗体名 = 窗体名_In;
    End Loop;
    Commit;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Update_Diary_Log;

  --功能：取固定发布报表和用户发布报表
  Procedure Get_Report_Menu
  (
    Cursor_Out Out t_Refcur,
    系统_In    In Zlprograms.系统%Type,
    序号_In    In Zlprograms.序号%Type,
    功能_In    In Zlreports.功能%Type,
    编号_In    In Zlreports.编号%Type
  ) Is
  Begin
    If Nvl(编号_In, '空空') <> '空空' Then
      Open Cursor_Out For
        Select 标志, 系统, 编号, 名称
        From (Select 1 As 标志, a.系统, a.编号, a.名称
               From zlReports A, zlPrograms B
               Where a.系统 = b.系统 And a.程序id = b.序号 And Not Upper(a.编号) Like '%BILL%' And
                     Upper(b.部件) <> Upper('zl9Report') And b.系统 = 系统_In And b.序号 = 序号_In And
                     Instr(功能_In, ';' || a.功能 || ';') > 0
               Union All
               Select Decode(a.系统, Null, 2, 1) As 标志, a.系统, a.编号, a.名称
               From zlReports A, zlRPTPuts B, zlPrograms C
               Where a.Id = b.报表id And b.系统 = c.系统 And b.程序id = c.序号 And (Not Upper(a.编号) Like '%BILL%' Or a.系统 Is Null) And
                     Instr(功能_In, ';' || b.功能 || ';') > 0 And c.系统 = 系统_In And c.序号 = 序号_In)
        Where Instr(编号_In, ',' || 编号 || ',') = 0
        Order By 标志, 编号;
    Else
      Open Cursor_Out For
        Select 标志, 系统, 编号, 名称
        From (Select 1 As 标志, a.系统, a.编号, a.名称
               From zlReports A, zlPrograms B
               Where a.系统 = b.系统 And a.程序id = b.序号 And Not Upper(a.编号) Like '%BILL%' And
                     Upper(b.部件) <> Upper('zl9Report') And b.系统 = 系统_In And b.序号 = 序号_In And
                     Instr(功能_In, ';' || a.功能 || ';') > 0
               Union All
               Select Decode(a.系统, Null, 2, 1) As 标志, a.系统, a.编号, a.名称
               From zlReports A, zlRPTPuts B, zlPrograms C
               Where a.Id = b.报表id And b.系统 = c.系统 And b.程序id = c.序号 And (Not Upper(a.编号) Like '%BILL%' Or a.系统 Is Null) And
                     Instr(功能_In, ';' || b.功能 || ';') > 0 And c.系统 = 系统_In And c.序号 = 序号_In)
        Order By 标志, 编号;
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Report_Menu;

  --功能：取用户提醒信息
  Procedure Get_Zlnoticerec
  (
    Cursor_Out Out t_Refcur,
    用户名_In  In Zlnoticerec.用户名%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select a.序号, a.系统, c.程序id As 模块, c.系统 As 报表系统, b.提醒内容 As 结果内容, c.名称 As 提醒报表, a.提醒声音, b.检查时间, b.已读标志
      From zlNotices A, zlNoticeRec B, (Select * From zlReports Where 发布时间 Is Not Null) C
      Where b.用户名 = 用户名_In And b.提醒标志 > 0 And c.编号(+) = a.提醒报表 And a.序号 = b.提醒序号 And b.提醒内容 Is Not Null;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Zlnoticerec;

  --功能：取邮件正文
  Procedure Get_Zlmessage
  (
    Cursor_Out Out t_Refcur,
    Id_In      In Zlmessages.Id%Type,
    类型_In    In Zlmsgstate.类型%Type,
    用户_In    In Zlmsgstate.用户%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select a.*, b.删除, b.状态
      From zlMessages A, zlMsgState B
      Where a.Id = b.消息id And b.消息id = Id_In And b.类型 = 类型_In And b.用户 = 用户_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Zlmessage;

  --功能：取邮件内容
  Procedure Get_Zlmessage
  (
    Cursor_Out Out t_Refcur,
    Id_In      In Zlmessages.Id%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select 内容, 背景色 From zlMessages Where ID = Id_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Zlmessage;

  --功能：取邮递地址
  Procedure Get_Zlmsgstate
  (
    Cursor_Out Out t_Refcur,
    消息id_In  In Zlmsgstate.消息id%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select 类型, 用户, 身份 From zlMsgState Where 消息id = 消息id_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Zlmsgstate;

  --功能：删除消息
  Procedure Delete_Zlmsgstate
  (
    删除_In   In Zlmsgstate.删除%Type,
    消息id_In In Zlmsgstate.消息id%Type,
    类型_In   In Zlmsgstate.类型%Type,
    用户_In   In Zlmsgstate.用户%Type
  ) Is
    n_总数 Number(10);
    n_数量 Number(10);
  Begin
    If Nvl(删除_In, 0) = 1 Then
      Update zlMsgState Set 删除 = 1 Where 消息id = 消息id_In And 类型 = 类型_In And 用户 = 用户_In;
    Else
      If 类型_In = 0 Then
        --对于草稿，把收件人的也一并删除
        Update zlMsgState Set 删除 = 2 Where 消息id = 消息id_In And 用户 = 用户_In;
      Else
        Update zlMsgState Set 删除 = 2 Where 消息id = 消息id_In And 类型 = 类型_In And 用户 = 用户_In;
      End If;
      -- 删除指定ID的消息  mnuEditDelete_Click 调用
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
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Delete_Zlmsgstate;

  --功能：删除过期消息
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
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Delete_Zlmessage;

  --功能：取邮件列表
  Procedure Get_Mail_List
  (
    Cursor_Out  Out t_Refcur,
    消息类型_In In Varchar2,
    用户_In     In Zlmsgstate.用户%Type,
    显示已读_In In Number,
    会话id_In   In Zlmessages.会话id%Type
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
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Mail_List;

  --功能：还原删除的消息
  Procedure Restore_Zlmsgstate
  (
    消息id_In In Zlmsgstate.消息id%Type,
    类型_In   In Zlmsgstate.类型%Type,
    用户_In   In Zlmsgstate.用户%Type
  ) Is
  Begin
    Update zlMsgState Set 删除 = 0 Where 消息id = 消息id_In And 类型 = 类型_In And 用户 = 用户_In;
    Commit;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Restore_Zlmsgstate;

  --功能：保存消息
  --调用列表：zlApptools.frmMessageEdit.SaveMessage
  Procedure Save_Zlmessage
  (
    Cursor_Out Out t_Refcur,
    Id_In      In Zlmessages.Id%Type,
    会话id_In  In Zlmessages.会话id%Type,
    发件人_In  In Zlmessages.发件人%Type,
    收件人_In  In Zlmessages.收件人%Type,
    主题_In    In Zlmessages.主题%Type,
    内容_In    In Zlmessages.内容%Type,
    背景色_In  In Zlmessages.背景色%Type
  ) Is
    n_Id     Zlmessages.Id%Type;
    n_会话id Zlmessages.会话id%Type;
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
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Save_Zlmessage;

  --功能：插入zlMsgstate
  --调用列表：zlApptools.frmMessageEdit.SaveMessage
  Procedure Insert_Zlmsgstate
  (
    消息id_In In Zlmsgstate.消息id%Type,
    类型_In   In Zlmsgstate.类型%Type,
    用户_In   In Zlmsgstate.用户%Type,
    身份_In   In Zlmsgstate.身份%Type,
    删除_In   In Zlmsgstate.删除%Type,
    状态_In   In Zlmsgstate.状态%Type
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
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Insert_Zlmsgstate;

  --功能：为原件加上答复或转发标志
  Procedure Update_Zlmsgstate_State
  (
    模式_In   In Number,
    消息id_In In Zlmsgstate.消息id%Type,
    类型_In   In Zlmsgstate.类型%Type,
    用户_In   In Zlmsgstate.用户%Type
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
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Update_Zlmsgstate_State;

  --功能：更新状态和身份
  Procedure Update_Zlmsgstate_Idtntify
  (
    身份_In   In Zlmsgstate.身份%Type,
    消息id_In In Zlmsgstate.消息id%Type,
    类型_In   In Zlmsgstate.类型%Type,
    用户_In   In Zlmsgstate.用户%Type
  ) Is
  Begin
    Update zlMsgState
    Set 状态 = '1' || Substr(状态, 2), 身份 = 身份_In
    Where 消息id = 消息id_In And 类型 = 类型_In And 用户 = 用户_In;
    Commit;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Update_Zlmsgstate_Idtntify;

End b_Comfunc;
/