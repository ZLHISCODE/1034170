----10.33.0---��9.44.0

--66466:��˶,2013-10-15,Ӧ��ϵͳ���Ӷ��ﻯ��ͼ����Ȩ
Declare
  V_Sql Varchar2(1000);
Begin
  For R In (Select Distinct ������ From zlSystems) Loop
    Begin
      --��ϵͳ��������Ȩ
      V_Sql := 'Grant Create Materialized View, Alter Any Materialized View, Drop Any Materialized View To ' || R.������ ||
               ' With Admin Option';
      Execute Immediate V_Sql;
    Exception
      When Others Then
        Null;
        --�����߿��ܲ�����
    End;
  End Loop;
End;
/



