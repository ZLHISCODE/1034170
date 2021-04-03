----10.32.0---》9.43.0

--59747:张永康,2013-03-26,历史数据转出性能优化
Alter Table zltools.zlDataMove add 状态 number(1);
Alter Table zltools.zlDataMove add 停用作业号 Varchar2(4000);
Alter Table zltools.zlBakTables add 停用触发器 number(1);

create table zltools.zlDataMovelog
(系统 number(5),
批次 number(8),
序列 number(8),
截止时间 date,
待转出 number(1),
当前进度 varchar2(100),
标记开始时间 date,
标记结束时间 date,
转出开始时间 date,
转出结束时间 date,
重建结束时间 date)
/
ALTER TABLE zltools.zlDataMovelog ADD CONSTRAINT zlDataMovelog_PK PRIMARY KEY (系统,批次) USING INDEX PCTFREE 5
/
ALTER TABLE zltools.zlDataMovelog ADD CONSTRAINT zlDataMovelog_FK_系统 FOREIGN KEY (系统) REFERENCES zlSystems(编号) ON DELETE CASCADE
/
create public synonym zlDataMovelog for zltools.zlDataMovelog;
grant select on zltools.zlDataMovelog to public;

--61786:张永康,2013-05-21,历史数据转出日志
Declare
  v_Sql Varchar2(100);
Begin
  For R In (Select Distinct 所有者 From zlSystems) Loop
    Begin
      v_Sql := 'grant select,insert,update,delete on zlDataMovelog to ' || r.所有者 || ' With GRANT Option';
      Execute Immediate v_Sql;
    Exception
      When Others Then
        Null;
        --所有者可能不存在(系统停用)
    End;
  End Loop;
End;
/


--61766:刘硕,2013-05-20,管理工具系统升级功能改造
create table zltools.Zlbigtables(系统 Number(5),表名 Varchar2(30))
/
alter table ZLTOOLS.Zlbigtables add constraint Zlbigtables_PK primary key (系统, 表名) USING INDEX PCTFREE 5
/

create public synonym Zlbigtables for zltools.Zlbigtables;
grant select on zltools.Zlbigtables to public;

Declare
  v_Sql Varchar2(100);
Begin
  For R In (Select Distinct 所有者 From zlSystems) Loop
    Begin
      v_Sql := 'grant select,insert,update,delete on Zlbigtables to ' || r.所有者 || ' With GRANT Option';
      Execute Immediate v_Sql;
    Exception
      When Others Then
        Null;
        --所有者可能不存在(系统停用)
    End;
  End Loop;
End;
/


--59244:谢荣,2013-03-08,根据授权码读取权限问题。
Create Or Replace Function f_Reg_Menu
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
               f.功能 = r.功能(+) And (r.功能 Is Null And f.系统 Is Null Or r.功能 Is Not Null And r.功能 = '基本' Or
                                   r.功能 Is Not Null And x.程序id Is Not Null) And p.系统 = x.系统(+) And p.序号 = x.程序id(+) And
               Upper(p.部件) = c.Text And Nvl(p.系统, 0) = s.Prog And p.序号 = p.序号 * a.超级(+) And Nvl(p.系统, 1) = o.编号(+) And
               Nvl(p.系统, 0) = Nvl(g.系统(+), 0) And p.序号 = g.序号(+) And
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

--57403:梁唐彬,2012-12-27
Create Table zlTools.zlMgrGrant(
    用户名 varchar2(30),
    功能 varchar2(500))
    PCTFREE 5 PCTUSED 90
    Cache Storage(Buffer_Pool Keep)
/
Create Public Synonym zlMgrGrant For zlTools.zlMgrGrant
/
Grant Select On zlTools.zlMgrGrant To Public
/

ALTER TABLE zlTools.zlMgrGrant ADD CONSTRAINT zlMgrGrant_PK PRIMARY KEY (用户名) USING INDEX PCTFREE 5
/
Insert Into zlTools.zlSvrTools(编号,上级,标题,快键,说明) Values('0404','04','管理工具授权','N',Null)
/

--57426:吴涛,2012-12-28
Insert Into zlSvrTools(编号,上级,标题,快键,说明) Values('0505','05','自定义过程管理','Z',Null)
/
Create Sequence zlTools.zlProcedure_ID start with 1
/
Create Public Synonym zlProcedure_ID For zlTools.zlProcedure_ID
/
Grant Select On zlTools.zlProcedure_ID to Public
/

CREATE TABLE zlTools.zlProcedure(
    ID NUMBER(5),
    类型 NUMBER(5),
    名称 VARCHAR2(50),
    状态 NUMBER(5),
    说明 VARCHAR2(200),
    所有者 VARCHAR2(20),
    修改人员 Varchar2(20),
    修改时间 DATE,
    上次修改人员 VARCHAR2(20),
    上次修改时间 Date)
    PCTFREE 5 PCTUSED 90
    Cache Storage(Buffer_Pool Keep)
/
Create Public Synonym zlProcedure For zlTools.zlProcedure
/
Grant Select On zlTools.zlProcedure to Public
/

ALTER TABLE zlTools.zlProcedure ADD CONSTRAINT zlProcedure_PK PRIMARY KEY (ID) USING INDEX PCTFREE 5
/

CREATE TABLE zlTools.zlProcedureNote(
    过程id NUMBER(5),
    标识 VARCHAR2(50),
    说明 VARCHAR2(4000))
    PCTFREE 5 PCTUSED 90
    Cache Storage(Buffer_Pool Keep)
/
Create Public Synonym zlProcedureNote For zlTools.zlProcedureNote
/
Grant Select On zlTools.zlProcedureNote to Public
/

ALTER TABLE zlTools.zlProcedureNote ADD CONSTRAINT zlProcedureNote_FK_过程id FOREIGN KEY (过程id) REFERENCES zlProcedure(ID)
/

CREATE TABLE zlTools.zlProcedureText(
    过程id NUMBER(5),
    性质 VARCHAR2(50),
    序号 NUMBER(5),
    内容 VARCHAR2(4000))
    PCTFREE 5 PCTUSED 90
    Cache Storage(Buffer_Pool Keep)
/
Create Public Synonym zlProcedureText For zlTools.zlProcedureText
/
Grant Select On zlTools.zlProcedureText to Public
/

ALTER TABLE zlTools.zlProcedureText ADD CONSTRAINT zlProcedureText_FK_过程id FOREIGN KEY (过程id) REFERENCES zlProcedure(ID)
/
Insert Into zlOptions(参数号,参数名,参数值,缺省值,参数说明) Values(8, '连接配置', '','', '搜集更新时连接的数据库配置')
/
Create Or Replace Procedure zlTools.Zl_Zlprocedure_Update
(
  Id_In           In Zlprocedure.Id%Type,
  类型_In         In Zlprocedure.类型%Type,
  名称_In         In Zlprocedure.名称%Type,
  状态_In         In Zlprocedure.状态%Type,
  说明_In         In Zlprocedure.说明%Type := Null,
  所有者_In       In ZLprocedure.所有者%Type :=Null,
  修改人员_In     In Zlprocedure.修改人员%Type := Null,
  修改时间_In     In Zlprocedure.修改时间%Type := Null,
  上次修改人员_In In Zlprocedure.上次修改人员%Type := Null,
  上次修改时间_In In Zlprocedure.上次修改时间%Type := Null
) Is
Begin
  Update Zlprocedure
  Set ID = Id_In, 类型 = 类型_In, 名称 = 名称_In, 状态 = 状态_In, 说明 = 说明_In, 所有者 = 所有者_In, 修改人员 = 修改人员_In, 修改时间 = 修改时间_In, 上次修改人员 = 上次修改人员_In,
      上次修改时间 = 上次修改时间_In
  Where ID = Id_In;
  If Sql%RowCount = 0 Then
    Insert Into Zlprocedure
      (ID, 类型, 名称, 状态, 说明, 所有者, 修改人员, 修改时间, 上次修改人员, 上次修改时间)
    Values
      (Id_In, 类型_In, 名称_In, 状态_In, 说明_In, 所有者_In, 修改人员_In, 修改时间_In, 上次修改人员_In, 上次修改时间_In);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlprocedure_Update;
/
Create Public Synonym Zl_Zlprocedure_Update For zlTools.Zl_Zlprocedure_Update
/
Grant Execute On zlTools.Zl_Zlprocedure_Update to Public
/

Create Or Replace Procedure zlTools.Zl_Zlprocedureconnect_Update
(
  参数名_In In Zloptions.参数名%Type,
  参数值_In In Zloptions.参数值%Type
) Is
Begin
  Update zlOptions Set 参数值 = 参数值_In Where 参数名 = 参数名_In;
  If Sql%RowCount = 0 Then
    Insert Into zlOptions
      (参数号, 参数名, 参数值, 参数说明)
    Values
      (8, 参数名_In, 参数值_In, '搜集更新时连接的数据库配置');
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlprocedureconnect_Update;
/
Create Public Synonym Zl_Zlprocedureconnect_Update For zlTools.Zl_Zlprocedureconnect_Update
/
Grant Execute On zlTools.Zl_Zlprocedureconnect_Update to Public
/

Create Or Replace Procedure zlTools.Zl_Zlprocedure_Delete
(
  Id_In           In Zlprocedure.Id%Type
) Is
Begin
  Delete zlProcedureNote Where 过程ID=Id_In;
  Delete zlProcedureText Where 过程ID=Id_In;
  Delete zlProcedure Where ID = Id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlprocedure_Delete;
/
Create Public Synonym Zl_Zlprocedure_Delete For zlTools.Zl_Zlprocedure_Delete
/
Grant Execute On zlTools.Zl_Zlprocedure_Delete to Public
/

Create Or Replace Procedure zlTools.Zl_zlProcedureNote_Update
(
  过程Id_In       In zlProcedureNote.过程id%Type,
  标识_In         In zlProcedureNote.标识%Type,
  说明_In         In zlProcedureNote.说明%Type
) Is
Begin
  Insert Into Zlprocedurenote (过程id, 标识, 说明) Values (过程id_In, 标识_In, 说明_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_zlProcedureNote_Update;
/
Create Public Synonym Zl_zlProcedureNote_Update For zlTools.Zl_zlProcedureNote_Update
/
Grant Execute On zlTools.Zl_zlProcedureNote_Update to Public
/

Create Or Replace Procedure zlTools.Zl_zlProcedureNote_Delete
(
  过程Id_In       In zlProcedureNote.过程id%Type
) Is
Begin
  Delete From zlProcedureNote Where 过程Id = 过程Id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_zlProcedureNote_Delete;
/
Create Public Synonym Zl_zlProcedureNote_Delete For zlTools.Zl_zlProcedureNote_Delete
/
Grant Execute On zlTools.Zl_zlProcedureNote_Delete to Public
/

CREATE OR REPLACE Procedure zlTools.Zl_Zlproceduretext_Update
(
  过程id_In In Zlproceduretext.过程id%Type,
  性质_In   In Zlproceduretext.性质%Type,
  序号_In   In Zlproceduretext.序号%Type,
  内容_In   In Zlproceduretext.内容%Type
) Is
Begin
  Update Zlproceduretext Set 性质 = 性质_In, 序号 = 序号_In, 内容 = 内容_In Where 过程id = 过程id_In And 性质 = 性质_In And 序号 = 序号_In;
  If Sql%RowCount = 0 Then
    Insert Into Zlproceduretext (过程id, 性质, 序号, 内容) Values (过程id_In, 性质_In, 序号_In, 内容_In);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlproceduretext_Update;
/
Create Public Synonym Zl_Zlproceduretext_Update For zlTools.Zl_Zlproceduretext_Update
/
Grant Execute On zlTools.Zl_Zlproceduretext_Update to Public
/

CREATE OR REPLACE Procedure zlTools.Zl_Zlproceduretext_Delete
(
  过程id_In In Zlproceduretext.过程id%Type,
  性质_In   In Zlproceduretext.性质%Type
) Is
Begin
  Delete From ZLproceduretext Where 过程id = 过程id_In And 性质 = 性质_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlproceduretext_Delete;
/       
Create Public Synonym Zl_Zlproceduretext_Delete For zlTools.Zl_Zlproceduretext_Delete
/
Grant Execute On zlTools.Zl_Zlproceduretext_Delete to Public
/

CREATE OR REPLACE Procedure ZLTOOLS.Zl_Zlproceduretext_Move Is
Begin
  Delete From Zlproceduretext Where 性质 In (1,2);

  Insert Into Zlproceduretext(过程id,性质,序号,内容)
  Select 过程id,Decode(性质,3,1,4,2),序号,内容 From Zlproceduretext Where 性质 In (3,4);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlproceduretext_Move;
/
Create Public Synonym Zl_Zlproceduretext_Move For zlTools.Zl_Zlproceduretext_Move
/
Grant Execute On zlTools.Zl_Zlproceduretext_Move to Public
/
