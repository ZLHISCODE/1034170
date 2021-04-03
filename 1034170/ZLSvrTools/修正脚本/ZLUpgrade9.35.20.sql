Create Or Replace Procedure zlTools.zl_Parameters_Update
(
  参数_In   zlParameters.参数名%Type,
  参数值_In zlParameters.参数值%Type,
  系统_In   zlParameters.系统%Type,
  模块_In   zlParameters.模块%Type,
  权限_IN   Number:=1
  --功能：设置系统参数值，如果是用户私有参数，则用户名以当前的为准
  --参数：
  --      参数_In：必须传入非Null值，以字符形式传入的参数号或参数名,注意参数名不能为数字。
  --      权限_IN：对于要求用权限控制的参数，当前用户是否有权限设置
) Is
  v_参数id zlParameters.ID%Type;
  v_私有   zlParameters.私有%Type;
  v_本机   zlParameters.本机%Type;
  v_授权   zlParameters.授权%Type;
  v_机器名 zlUserParas.机器名%Type;
Begin
  --确定参数信息
  Begin
    If Zl_To_Number(参数_In) <> 0 Then
      --以参数号为准处理
      Select ID, 私有, 本机, 授权, Sys_Context('USERENV', 'TERMINAL')
      Into v_参数id, v_私有, v_本机, v_授权, v_机器名
      From zlParameters
      Where Nvl(系统, 0) = Nvl(系统_In, 0) And Nvl(模块, 0) = Nvl(模块_In, 0) And 参数号 = Zl_To_Number(参数_In);
    Else
      --以参数名为准处理
      Select ID, 私有, 本机, 授权, Sys_Context('USERENV', 'TERMINAL')
      Into v_参数id, v_私有, v_本机, v_授权, v_机器名
      From zlParameters
      Where Nvl(系统, 0) = Nvl(系统_In, 0) And Nvl(模块, 0) = Nvl(模块_In, 0) And 参数名 = 参数_In;
    End If;
  Exception
    When Others Then
      Return;
  End;
  
  --检查权限
  If Nvl(权限_IN, 0) = 0 Then
    If Nvl(系统_In, 0) <> 0 And Nvl(模块_In, 0) = 0 And Nvl(v_私有, 0) = 0 And Nvl(v_本机, 0) = 0 Then
       Return;--公共全局参数,固定需要权限
    Elsif Nvl(系统_In, 0) <> 0 And Nvl(模块_In, 0) <> 0 And Nvl(v_私有, 0) = 0 And Nvl(v_本机, 0) = 0 Then
       Return;--公共模块参数,固定需要权限
    Elsif Nvl(系统_In, 0) <> 0 And Nvl(模块_In, 0) <> 0 And Nvl(v_私有, 0) = 0 And Nvl(v_本机, 0) = 1 And Nvl(v_授权, 0) = 1 Then
       Return;--要授权控制的本机公共模块
    End If;
  End If;
    
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
End zl_Parameters_Update;
/