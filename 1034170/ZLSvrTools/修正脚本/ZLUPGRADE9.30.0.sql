
-------------------------------------------------------------------------------
--刘兴宏:
--问题:9988
alter table zlreginfo modify(项目 varchar2(20))
/

--NULL,表示以前的固定服务器,1表示1号服务器,2表示2号服务器...
alter table zlClients add(升级服务器 number(2))
/


-------------------------------------------------------------------------------
INSERT INTO zlTools.zlRegInfo (项目,行号,内容) VALUES ('站点编号',Null,Null)
/
INSERT INTO zlTools.zlRegInfo (项目,行号,内容) VALUES ('站点类型',Null,Null)
/
INSERT INTO zlTools.zlRegInfo (项目,行号,内容) VALUES ('流状态',Null,Null)
/

Create Table zlTools.zlStreamTabs(
  System_NO  Number(5),
  Table_Name Varchar2(30),
  Dml_Handle Number(1), --是否存在DML句柄过程
  Repeat_Way Number(1), --默认的复制方向：1-本地表;2-主站分发表;3-双向复制表
  Fixation   Number(1)	--复制方向是否固定不可更改
)
/
ALTER TABLE zlTools.zlStreamTabs ADD CONSTRAINT zlStreamTabs_PK PRIMARY KEY (System_NO,Table_Name) USING INDEX PCTFREE 0
/
Alter Table zlTools.zlStreamTabs Add Constraint zlStreamTabs_FK_SYSNO Foreign Key (System_NO) References zlsystems(编号)
/
Create Public Synonym zlStreamTabs For zlTools.zlStreamTabs
/
GRANT SELECT ON zlTools.zlStreamTabs TO PUBLIC 
/
Begin
  For r_User In (Select 所有者 From Zlsystems) Loop
    Execute Immediate 'Grant Select,Insert,Update,Delete on zlTools.zlStreamTabs to ' || r_User.所有者 ||
                      ' With Grant Option';
    Execute Immediate 'Grant Select,Insert,Update,Delete on zlTools.zlStreamTabs to ' || r_User.所有者 ||
                      ' With Grant Option';
  End Loop;
End;
/

Create Or Replace Function zlTools.f_Get_Node_No Return Varchar2 As
  v_Return zlRegInfo.内容%Type;
Begin
  Begin
    Select 内容 Into v_Return From zlRegInfo Where 项目 = '站点编号';
    If To_Number(v_Return) < 0 Or To_Number(v_Return) > 9 Then
      v_Return := Null;
    End If;
  Exception
    When Others Then
      Null;
  End;
  Return(v_Return);
End f_Get_Node_No;
/


Create Or Replace Function zlTools.f_Is_Primary_Node Return Number As
  v_Return    Number;
  v_Node_Type zlRegInfo.内容%Type;
Begin
  Begin
    Select 内容 Into v_Node_Type From zlRegInfo Where 项目 = '站点类型';
  Exception
    When Others Then
      Null;
  End;
  If v_Node_Type = '1' Or v_Node_Type Is Null Then
    v_Return := 1;
  Else
    v_Return := 0;
  End If;
  Return(v_Return);
End f_Is_Primary_Node;
/

Create Or Replace Function Zltools.f_Get_Stream_State Return Number As
  v_Return Number;
  v_State  zlRegInfo.内容%Type;
Begin
  Begin
    Select 内容 Into v_State From zlRegInfo Where 项目 = '流状态';
  Exception
    When Others Then
      Null;
  End;
  If v_State = '1' Or v_State Is Null Then
    v_Return := 1;
  Else
    v_Return := 0;
  End If;
  Return(v_Return);
End f_Get_Stream_State;
/

Grant Execute on zlTools.f_Get_Node_No to Public
/
Grant Execute on zlTools.f_Is_Primary_Node to Public
/
Grant Execute on zlTools.f_Get_Stream_State to Public
/
Create Public Synonym f_Get_Node_No For zlTools.f_Get_Node_No
/
Create Public Synonym f_Is_Primary_Node For zlTools.f_Is_Primary_Node
/
Create Public Synonym f_Get_Stream_State For zlTools.f_Get_Stream_State
/



Create Or Replace Package Body b_Runmana Is
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
