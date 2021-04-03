----10.33.0---》9.44.0

--66229:祝庆,2013-09-30,
Create Table zltools.zlClientUpdatelog (
   工作站 varchar2(50),
   处理日期 Date,
   内容 varchar2(300))
   PCTFREE 5 PCTUSED 90;

--66229:祝庆,2013-09-30,添加升级情况字段,用于记录最后一次升级情况
Alter Table zltools.zlClients add 升级情况 number(1) default(0);

--65904:梁唐彬,2013-09-17,执行单打印报表系统改进
Alter Table zltools.zlrptitems add (父ID number(18),源ID  number(18),上下间距  number(18),左右间距  number(18),纵向分栏  number(18),横向分栏  number(18),源行号  number(18));
ALTER TABLE zltools.zlRPTITems ADD CONSTRAINT zlRPTItems_FK_父ID FOREIGN KEY(父ID) REFERENCES zltools.zlRPTItems(ID) ON DELETE CASCADE;
CREATE INDEX zlRPTItems_IX_父ID ON zltools.zlRPTItems(父ID) PCTFREE 5  Compress 1;
ALTER TABLE zltools.zlRPTITems ADD CONSTRAINT zlRPTItems_FK_源ID FOREIGN KEY(源ID) REFERENCES zltools.Zlrptdatas(ID) ON DELETE CASCADE;
CREATE INDEX zlRPTItems_IX_源ID ON zltools.zlRPTItems(源ID) PCTFREE 5;

--65714:祝庆,2013-09-011,管理工具添加升级文件模块
Insert Into ZLTOOLS.zlSvrTools(编号,上级,标题,快键,说明) Values('0311','03','升级文件管理','Q',Null);

--65203:刘硕,2013-10-11,管理工具修正之权限修正
Declare
  V_Sql Varchar2(100);
Begin
  For R In (Select Distinct 所有者 From zlSystems) Loop
  
    For R_Table In (Select Column_Value Tabname From Table(F_Str2list('Zlfilesupgrade,zlRegFunc,zlClients'))) Loop
      --由于系统所有者对角色授权使用 With Grant Option 因此取消授权可以取消 角色权限。
      Begin
        V_Sql := 'Revoke select,insert,update,delete on ZLTOOLS.' || R_Table.Tabname || ' from ' || R.所有者;
        Execute Immediate V_Sql;
      Exception
        When Others Then
          Null;
          --所有者可能不存在(系统停用)或者系统所有者没有这些权限或者表不存在
      End;
    
      Begin
        --重新对系统所有者授权
        V_Sql := 'grant select,insert,update,delete on ZLTOOLS.' || R_Table.Tabname || ' to ' || R.所有者 || ' With Grant Option';
        Execute Immediate V_Sql;
      Exception
        When Others Then
          Null;
          --所有者可能不存在(系统停用)或者表不存在
      End;
    End Loop;
  
  End Loop;
  --特殊存储过程权限回收
  For R_Prog In (Select Column_Value Procname
                 From Table(F_Str2list('B_ROLEGROUPMGR,ZL_ZLROLEGRANT_BATCHDELETE,ZL_ZLROLEGRANT_BATCHINSERT'))) Loop
    Begin
      V_Sql := 'Revoke Execute on ZLTOOLS.' || R_Prog.Procname || ' From Public';
      Execute Immediate V_Sql;
    Exception
      When Others Then
        Null;
        --不存在权限或对象不存在
    End;
  End Loop;
End;
/

--65203:刘硕,2013-09-03,管理工具修正
--66367:许华锋,2013-10-12,站点控制启用视频源
Create Or Replace Procedure zltools.Zl_Zlclients_Set
(
  N_Mode_In       Number,
  N_Rowid_In      Varchar2 := Null,
  V_工作站_In     Zlclients.工作站%Type := Null,
  V_Ip_In         Zlclients.Ip%Type := Null,
  V_Cpu_In        Zlclients.Cpu%Type := Null,
  V_内存_In       Zlclients.内存%Type := Null,
  V_硬盘_In       Zlclients.硬盘%Type := Null,
  V_操作系统_In   Zlclients.操作系统%Type := Null,
  V_部门_In       Zlclients.部门%Type := Null,
  V_用途_In       Zlclients.用途%Type := Null,
  V_说明_In       Zlclients.说明%Type := Null,
  N_升级服务器_In Zlclients.升级服务器%Type := Null,
  N_升级标志_In   Zlclients.升级标志%Type := 0,
  N_连接数_In     Zlclients.连接数%Type := 0,
  V_站点_In       Zlclients.站点%Type := Null,
  N_Apply_In      Number := 0,
  V_Ipbegin_In    Varchar2 := Null,
  V_Ipend_In      Varchar2 := Null,
  N_启用视频源    Zlclients.启用视频源%Type := 0
  --功能：新增客户端或站点 或者更新客户端属性
  --应用：1、管理工具：新增或修改站点 （修改时以IP与客户端做判断条件，不需传入N_Rowid_In）
  --      2：应用系统：登录时根据当前登录的客户短来判断是否
  --                   新增站点或修改站点参数（更新时N_Rowid_In需传入）
  --站点设置:0-新增站点，1-更新站点
  --N_Apply_In,站点参数应用范围，0-本站点，1，本部门，2，所有站点，3，固定IP段
  --V_Ipbegin_In,V_Ipend_In:在固定IP断应用时传入,两者在一个IP断上，即前面部分相同
) Is
  N_Pos         Number(3);
  N_Ipbegin_Num Number;
  N_Ipend_Num   Number;
  N_Ip_Num      Number;
  N_Count       Number;

  V_Err Varchar2(500);
  Err_Custom Exception;

  Function Get_Ipnum(V_Ip_Input Varchar2) Return Number Is
    V_Ip_Num  Varchar2(20);
    N_Pos_Tmp Number;
    V_Ip_Tmp  Varchar2(20);
  Begin
    N_Pos_Tmp := Length(V_Ip_Input);
    N_Pos_Tmp := N_Pos_Tmp - Length(Replace(V_Ip_Input, '.', ''));
    If N_Pos_Tmp <> 3 Then
      Return Null;
    Else
      V_Ip_Tmp := V_Ip_Input;
      Loop
        N_Pos_Tmp := Instr(V_Ip_Tmp, '.');
        Exit When(Nvl(N_Pos_Tmp, 0) = 0);
        --将每一断数字转化为3位数
        V_Ip_Num := V_Ip_Num || Trim(To_Char(Substr(V_Ip_Tmp, 1, N_Pos_Tmp - 1), '099'));
        V_Ip_Tmp := Substr(V_Ip_Tmp, N_Pos_Tmp + 1);
      End Loop;
      V_Ip_Num := V_Ip_Num || Trim(To_Char(V_Ip_Tmp, '099'));
      N_Ip_Num := To_Number(Trim(V_Ip_Num));
      Return N_Ip_Num;
    End If;
  End;
Begin
  If N_Mode_In = 0 Then

    Select Count(1) Into N_Count From zlClients Where 工作站 = V_工作站_In;
    If N_Count = 0 Then
      Insert Into ZLTOOLS.zlClients
        (Ip, 工作站, Cpu, 内存, 硬盘, 操作系统, 部门, 用途, 说明, 升级服务器, 升级标志, 连接数, 站点, 启用视频源)
      Values
        (V_Ip_In, V_工作站_In, V_Cpu_In, V_内存_In, V_硬盘_In, V_操作系统_In, V_部门_In, V_用途_In, V_说明_In, N_升级服务器_In, N_升级标志_In,
         N_连接数_In, V_站点_In, N_启用视频源);
    Else
      V_Err := '已经设置了相同IP地址或工作站,不能再设!';
      Raise Err_Custom;
    End If;
  Else
    If N_Rowid_In Is Null Then
      Update ZLTOOLS.zlClients
      Set Cpu = V_Cpu_In, 内存 = V_内存_In, 硬盘 = V_硬盘_In, 操作系统 = V_操作系统_In, 部门 = V_部门_In, 用途 = V_用途_In, 说明 = V_说明_In,
          连接数 = N_连接数_In, 站点 = V_站点_In, 启用视频源=N_启用视频源, 升级服务器 = N_升级服务器_In, 升级标志 = N_升级标志_In
      Where 工作站 = V_工作站_In And Ip = V_Ip_In;
    Else
      Update ZLTOOLS.zlClients
      Set 工作站 = V_工作站_In, Ip = V_Ip_In, Cpu = Decode(Cpu, Null, V_Cpu_In, Cpu), 内存 = Decode(内存, Null, V_内存_In, 内存),
          硬盘 = Decode(硬盘, Null, V_硬盘_In, 硬盘), 操作系统 = Decode(操作系统, Null, V_操作系统_In, 操作系统), 部门 = V_部门_In, 站点 = V_站点_In, 启用视频源=N_启用视频源
      Where Rowid = N_Rowid_In;
    End If;
  End If;
  --本部门
  If N_Apply_In = 1 Then
    Update ZLTOOLS.zlClients
    Set 连接数 = N_连接数_In, 站点 = V_站点_In
    Where Nvl(部门, 'NONE') = Nvl(V_部门_In, 'NONE') And Ip <> V_Ip_In;
  Elsif N_Apply_In = 2 Then
    Update ZLTOOLS.zlClients Set 连接数 = N_连接数_In, 站点 = V_站点_In Where Ip <> V_Ip_In;
  Elsif N_Apply_In = 3 Then
    N_Pos := Length(V_Ipbegin_In);
    N_Pos := N_Pos - Length(Replace(V_Ipbegin_In, '.', ''));
    If N_Pos <> 3 Then
      V_Err := '起始IP格式有误！';
      Raise Err_Custom;
    End If;
    N_Pos := Length(V_Ipend_In);
    N_Pos := N_Pos - Length(Replace(V_Ipend_In, '.', ''));
    If N_Pos <> 3 Then
      V_Err := '结束IP格式有误！';
      Raise Err_Custom;
    End If;

    N_Ipbegin_Num := Get_Ipnum(V_Ipbegin_In);
    N_Ipend_Num   := Get_Ipnum(V_Ipend_In);
    For R_Ip In (Select 工作站, Ip From zlClients) Loop
      N_Ip_Num := Get_Ipnum(R_Ip.Ip);
      If N_Ip_Num >= N_Ipbegin_Num And N_Ip_Num <= N_Ipend_Num Then
        Update ZLTOOLS.zlClients Set 连接数 = N_连接数_In, 站点 = V_站点_In Where 工作站 = R_Ip.工作站 And Ip = R_Ip.Ip;
      End If;
    End Loop;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || V_Err || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlclients_Set;
/

--65203:刘硕,2013-09-03,管理工具修正
Create Or Replace Procedure ZLTOOLS.Zl_Zlclients_Delete
(
  V_工作站_In Zlclients.工作站%Type := Null,
  V_Ip_In     Zlclients.Ip%Type := Null
) Is
Begin
  If Not (V_工作站_In Is Null And V_Ip_In Is Null) Then
    If V_Ip_In Is Null Then
      Delete ZLTOOLS.zlClients Where 工作站 = V_工作站_In;
    Elsif V_工作站_In Is Null Then
      Delete ZLTOOLS.zlClients Where Ip = V_Ip_In;
    Else
      Delete ZLTOOLS.zlClients Where Ip = V_Ip_In And 工作站 = V_工作站_In;
    End If;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlclients_Delete;
/
--65203:刘硕,2013-09-03,管理工具修正
CREATE OR REPLACE Procedure ZLTOOLS.Zl_Zlclients_Control
(  
  N_Mode_In       Number,
  V_工作站_In     Zlclients.工作站%Type := Null,
  V_Ip_In         Zlclients.Ip%Type := Null,
  N_升级标志_In   Zlclients.升级标志%Type := Null,
  N_升级服务器_In Zlclients.升级服务器%Type := Null,
  D_预升时点_In   Zlclients.预升时点%Type := Null,
  N_预升完成_In   Zlclients.预升完成%Type := Null,
  N_Ftp服务器_In  Zlclients.Ftp服务器%Type := Null,
  N_收集标志_In   Zlclients.收集标志%Type := Null,
  N_禁止使用_In   Zlclients.禁止使用%Type := Null,
  V_说明_In       zlclients.说明%Type :=Null
  --对客户端进行控制
  --N_Mode_In：0-禁用或启用客户端(IP做为主要条件）,1-预升级设置,2 -升级信息保存(IP做为主要条件）
  --3-取消预升级标志,4-将所有站点设置为升级,5-部件搜集（设置搜集标志）,6-重置升级状态
) Is
  V_Timeset Varchar2(300);
  V_Err     Varchar2(500);
  Err_Custom Exception;
Begin
  --0-禁用或启用客户端(IP做为主要条件）
  If N_Mode_In = 0 Then
    If V_工作站_In Is Not Null Then
      Update ZLTOOLS.zlClients Set 禁止使用 = N_禁止使用_In Where Ip = V_Ip_In;
    End If;
    --1-预升级设置,不需要传其他参数
  Elsif N_Mode_In = 1 Then
    Select Max(内容) Into V_Timeset From zlRegInfo Where 项目 = '客户端预升级时间点';
    If V_Timeset Is Not Null Then
      For R_Ip In (Select To_Date(Today || ' ' || Date_d, 'yyyy-mm-dd HH24:mi:ss') 预升时点, 工作站, Ip
                   From (Select 工作站, Ip, Rownum Rn_c From zlClients) A,
                        (Select To_Char(Sysdate, 'yyyy-mm-dd') Today, Column_Value Date_d, Rownum Rn_d, Count(1) Over() Sn
                          From Table(F_Str2list(V_Timeset, ','))) B
                   Where Mod(A.Rn_c, Sn) + 1 = Rn_d) Loop

        Update ZLTOOLS.zlClients Set 预升时点 = R_Ip.预升时点 Where 工作站 = R_Ip.工作站 And Ip = R_Ip.Ip;
      End Loop;
    Else
      V_Err := '你尚未进行客户端预升级时间点设置！';
      Raise Err_Custom;
    End If;
    --2 -升级信息保存(IP做为主要条件）
  Elsif N_Mode_In = 2 Then
    If N_Ftp服务器_In Is Null Then
      Update ZLTOOLS.zlClients
      Set 升级标志 = N_升级标志_In, 升级服务器 = N_升级服务器_In, 预升时点 = D_预升时点_In, 预升完成 = N_预升完成_In
      Where Ip = V_Ip_In;

    Else
      Update ZLTOOLS.zlClients
      Set 升级标志 = N_升级标志_In, Ftp服务器 = N_Ftp服务器_In, 预升时点 = D_预升时点_In, 预升完成 = N_预升完成_In
      Where Ip = V_Ip_In;
    End If;
    --3-取消预升级标志
  Elsif N_Mode_In = 3 Then
    Update ZLTOOLS.zlClients Set 预升完成 = N_预升完成_In;
    --4-将所有站点设置为升级
  Elsif N_Mode_In = 4 Then
    Update ZLTOOLS.zlClients Set 升级标志 = N_升级标志_In;
    --5-部件搜集（设置搜集标志）
  Elsif N_Mode_In = 5 Then
    If V_工作站_In Is Null Then
      Update ZLTOOLS.zlClients Set 收集标志 = N_收集标志_In;
    Else
      Update ZLTOOLS.zlClients Set 收集标志 = N_收集标志_In Where 工作站 = V_工作站_In;
    End If;
  Elsif N_Mode_In = 6 Then
    Update ZLTOOLS.zlClients Set 升级情况=0 Where 工作站 = V_工作站_In;
  Elsif N_Mode_In = 7 Then
    --7未升级
    Update ZLTOOLS.zlClients Set 升级情况=1 Where 工作站 = V_工作站_In;
  Elsif N_Mode_In = 8 Then
    --8已升级
    Update ZLTOOLS.zlClients Set 升级情况=2 Where 工作站 = V_工作站_In; 
  Elsif N_Mode_In = 9 Then
    --9修改说明
    Update zltools.zlclients set 说明=V_说明_In where upper(工作站)=upper(V_工作站_In);
  Elsif N_Mode_In = 10 Then
    --10修改说明和收集标志
    Update zltools.zlclients set 说明=V_说明_In,收集标志=0 where upper(工作站)=upper(V_工作站_In);
  Elsif N_Mode_In = 11 Then
    --11修改说明和升级标志
    Update zltools.zlclients set 说明=V_说明_In,升级标志=0 where upper(工作站)=upper(V_工作站_In);
  Elsif N_Mode_In = 12 Then
    Update zltools.zlclients set 说明=V_说明_In,预升完成=0 where upper(工作站)=upper(V_工作站_In);
  Elsif N_Mode_In = 13 Then
    Update zltools.zlclients set 预升完成=1 where upper(工作站)=upper(V_工作站_In);
  Elsif N_Mode_In = 14 Then
    Update zltools.zlclients set 预升时点=Null,预升完成=Null where upper(工作站)=upper(V_工作站_In);
  Elsif N_Mode_In = 15 Then
    Update zltools.zlClients
         Set 升级情况 =1
         Where upper(工作站) = (Select Upper(V_工作站_In)
         From v$Session
         Where AUDSID = UserENV('SessionID'));
  Elsif N_Mode_In = 16 Then
    Update zltools.zlClients
         Set 升级情况 =2
         Where upper(工作站) = (Select Upper(V_工作站_In)
         From v$Session
         Where AUDSID = UserENV('SessionID'));
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || V_Err || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlclients_Control;
/
--65203:刘硕,2013-09-03,管理工具修正
Create Or Replace Procedure ZLTOOLS.Zl_Zlclients_Upgrade
(
  N_Mode_In     Number,
  V_工作站_In   Zlclients.工作站%Type := Null,
  V_说明_In     Zlclients.说明%Type := Null,
  N_升级标志_In Zlclients.升级标志%Type := Null,
  N_收集标志_In Zlclients.收集标志%Type := Null,
  D_预升时点_In Zlclients.预升时点%Type := Null,
  N_预升完成_In Zlclients.预升完成%Type := Null
  --主要是客户端自动升级使用。
  --N_Mode_In:0-更改升级说明,更改升级和收集标记,1-更改站点的预升级完成状态
  --2-客户端为定时升级
) Is
  V_Err Varchar2(500);
  Err_Custom Exception;
Begin
  --0-更改升级说明,更改升级和收集标记
  If N_Mode_In = 0 Then
    If N_收集标志_In Is Null And N_升级标志_In Is Null Then
      Update ZLTOOLS.zlClients Set 说明 = V_说明_In Where Upper(工作站) = Upper(V_工作站_In);
    Elsif N_收集标志_In Is Null Then
      Update ZLTOOLS.zlClients Set 说明 = V_说明_In, 升级标志 = N_升级标志_In Where Upper(工作站) = Upper(V_工作站_In);
    Elsif N_升级标志_In Is Null Then
      Update ZLTOOLS.zlClients Set 说明 = V_说明_In, 收集标志 = N_收集标志_In Where Upper(工作站) = Upper(V_工作站_In);
    End If;
    --1-更改站点的预升级完成状态
  Elsif N_Mode_In = 1 Then
    If V_说明_In Is Null Then
      Update ZLTOOLS.zlClients Set 预升完成 = N_预升完成_In, 说明 = V_说明_In Where Upper(工作站) = Upper(V_工作站_In);
    Else
      Update ZLTOOLS.zlClients Set 预升完成 = N_预升完成_In Where Upper(工作站) = Upper(V_工作站_In);
    End If;
  --2-客户端为定时升级
  Elsif N_Mode_In = 2 Then
    Update ZLTOOLS.zlClients Set 预升时点 = D_预升时点_In, 预升完成 = N_预升完成_In Where Upper(工作站) = Upper(V_工作站_In);
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || V_Err || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlclients_Upgrade;
/

