----10.31.0---》9.42.70

--66367:许华锋,2013-10-12,站点控制启用视频源
Declare
  n_Count Number;
Begin

  select count(1) into n_count from all_tab_columns where owner = 'ZLTOOLS' and table_name = 'ZLCLIENTS' and COLUMN_NAME='启用视频源';

  If n_Count = 0 Then
    Execute Immediate 'Alter Table zltools.ZLCLIENTS Add (启用视频源 Number(1))';
  End If;
End;
/ 

--66367:许华锋,2013-10-12,站点控制启用视频源
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
                             a.连接数, Decode(b.Terminal, Null, 0, 1) As 状态, a.收集标志,a.升级服务器,a.站点,a.启用视频源
                From Zlclients a, (Select Distinct Terminal From V$session) b
                Where Upper(a.工作站) = Upper(b.Terminal(+))
                Order By a.Ip';
      Open Cur_Out For v_Sql;
    Else
      Open Cur_Out For
        Select Ip, 工作站, Cpu, 内存, 硬盘, 操作系统, 部门, 用途, 说明, 升级标志, 收集标志, 禁止使用, 连接数, 升级服务器, 站点,启用视频源
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
      Select 序号, 文件名, 版本号, 修改日期, 文件说明 As 说明,Decode(文件类型,0,'公共部件',1,'应用部件',2,'帮助文件',3,'其它文件',4,'三方部件',5,'系统文件','') as 类型,安装路径 as 安装路径,MD5 as MD5,加入日期
      From zlFilesUpgrade Order By 序号;
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




