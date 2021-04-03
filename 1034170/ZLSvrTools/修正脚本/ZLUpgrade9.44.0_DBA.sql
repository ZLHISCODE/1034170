----10.33.0---》9.44.0

--66466:刘硕,2013-10-15,应用系统增加对物化视图的授权
Declare
  V_Sql Varchar2(1000);
Begin
  For R In (Select Distinct 所有者 From zlSystems) Loop
    Begin
      --对系统所有者授权
      V_Sql := 'Grant Create Materialized View, Alter Any Materialized View, Drop Any Materialized View To ' || R.所有者 ||
               ' With Admin Option';
      Execute Immediate V_Sql;
    Exception
      When Others Then
        Null;
        --所有者可能不存在
    End;
  End Loop;
End;
/



