Create Or Replace Function zlTools.f_Str2list
(
  Str_In   In Varchar2,
  Split_In In Varchar2 := ','
) Return t_Strlist
  Pipelined As
  v_Str Long;
  P     Number;
  --功能：将由逗号分隔的不带引号的字符序列转换为单列数据表
  --参数：Str_In,如:G0000123,G0000124,G0000125...,Split_In,分隔符,缺省为,号
  --说明：
  --1．当SQL语句中涉及“IN(常量1, 常量2,…) ”子句时使用这种方式以便利用绑定变量。
  --2．使用这两个函数时，需要在SQL语句中加入“/*+ Rule*/”提示，因为Cbo下临时内存表没有统计数据,。
  --3．两种调用示例
  --Select /*+ Rule*/ * From 门诊费用记录 Where NO In (Select * From Table(f_Str2list('A01,A02,A03'));
  --Select /*+ Rule*/ A.* From 门诊费用记录 A, Table(f_Str2list('A01,A02,A03')) B Where A.NO = B.Column_Value;
Begin
  If Str_In Is Null Then
    Return;
  End If;
  v_Str := Str_In || Split_In;
  Loop
    P := Instr(v_Str, Split_In);
    Exit When(Nvl(P, 0) = 0);
    Pipe Row(Trim(Substr(v_Str, 1, P - 1)));
    v_Str := Substr(v_Str, P + 1);
  End Loop;
  Return;
End;
/

Create Or Replace Function zlTools.f_Num2list
(
  Str_In   In Varchar2,
  Split_In In Varchar2 := ','
) Return t_Numlist
  Pipelined As
  v_Str Long;
  P     Number;
Begin
  If Str_In Is Null Then
    Return;
  End If;
  v_Str := Str_In || Split_In;
  Loop
    P := Instr(v_Str, Split_In);
    Exit When(Nvl(P, 0) = 0);
    Pipe Row(To_Number(Trim(Substr(v_Str, 1, P - 1))));
    v_Str := Substr(v_Str, P + 1);
  End Loop;
  Return;
End;
/

Create Or Replace Type zlTools.t_NumObj2 as object (C1 Number,C2 number)
/
Create Or Replace Type zlTools.t_NumList2 as table of zlTools.t_NumObj2
/
Create Or Replace Function zlTools.f_Num2list2
(
  Str_In    In Varchar2,
  Split_In  In Varchar2 := ',',
  SubSplit_In In Varchar2 := ':'
) Return t_NumList2
  Pipelined As
  v_Str   Long;
  P       Number;
  v_Tmp   Varchar2(4000);
  Out_Rec t_NumObj2 := t_NumObj2(Null, Null);
Begin
  If Str_In Is Null Then
    Return;
  End If;
  v_Str := Str_In || Split_In;
  Loop
    P := Instr(v_Str, Split_In);
    Exit When(Nvl(P, 0) = 0);
    v_Tmp      := Trim(Substr(v_Str, 1, P - 1));
    Out_Rec.C1 := To_Number(Substr(v_Tmp, 1, Instr(v_Tmp, SubSplit_In) - 1));
    Out_Rec.C2 := To_Number(Substr(v_Tmp, Instr(v_Tmp, SubSplit_In) + 1));
    Pipe Row(Out_Rec);
    v_Str := Substr(v_Str, P + 1);
  End Loop;
  Return;
End;
/


Create Or Replace Type zlTools.t_StrObj2 as object (C1 Varchar2(4000),C2 Varchar2(4000))
/
Create Or Replace Type zlTools.t_StrList2 as table of zlTools.t_StrObj2
/
Create Or Replace Function zlTools.f_Str2list2
(
  Str_In    In Varchar2,
  Split_In  In Varchar2 := ',',
  SubSplit_In In Varchar2 := ':'
) Return t_StrList2
  Pipelined As
  v_Str   Long;
  P       Number;
  v_Tmp   Varchar2(4000);
  Out_Rec t_StrObj2 := t_StrObj2(Null, Null);
Begin
  If Str_In Is Null Then
    Return;
  End If;
  v_Str := Str_In || Split_In;
  Loop
    P := Instr(v_Str, Split_In);
    Exit When(Nvl(P, 0) = 0);
    v_Tmp      := Trim(Substr(v_Str, 1, P - 1));
    Out_Rec.C1 := trim(Substr(v_Tmp, 1, Instr(v_Tmp, SubSplit_In) - 1));
    Out_Rec.C2 := trim(Substr(v_Tmp, Instr(v_Tmp, SubSplit_In) + 1));
    Pipe Row(Out_Rec);
    v_Str := Substr(v_Str, P + 1);
  End Loop;
  Return;
End;
/

Create Public Synonym f_Str2list2 for zlTools.f_Str2list2; 
Grant execute on zlTools.f_Str2list2 to Public; 
Create Public Synonym f_Num2list2 for zlTools.f_Num2list2; 
Grant execute on zlTools.f_Num2list2 to Public;



--区别9i和10,采取不同的游标（c_Delsyn）优化查询性能
Create Or Replace Procedure zlTools.Zl_Createsynonyms(系统_In In Zlprogprivs.系统%Type) Authid Current_User As
  v_Sql    Varchar2(2000);
  v_所有者 Varchar2(100);
  n_Cnt    Number(5);

  --非当前所有者的对象的私有同义词与当前所有者的对象相同则删除
  Cursor c_Delsyn9(v_所有者 Varchar2) Is
    Select Synonym_Name 对象
    From User_Synonyms A
    Where Table_Owner != v_所有者 And Exists
     (Select 1
           From All_Objects B
           Where a.Synonym_Name = b.Object_Name And b.Owner = v_所有者 And
                 b.Object_Type In ('TABLE', 'VIEW', 'SEQUENCE', 'PROCEDURE', 'FUNCTION'));

  Cursor c_Delsyn10(v_所有者 Varchar2) Is
    Select a.Synonym_Name 对象
    From User_Synonyms A, All_Objects B
    Where a.Table_Owner != v_所有者 And a.Synonym_Name = b.Object_Name And b.Owner = v_所有者 And
          b.Object_Type In ('TABLE', 'VIEW', 'SEQUENCE', 'PROCEDURE', 'FUNCTION');

  --对用户有权访问的对象,属于当前系统所有者的,如果没有公共或私有同义词,则创建私有同义词
  --不仅限于当前模块所访问的对象,因为存在虚拟模块和模块中调用其它模块
  Cursor c_Newsyn(v_所有者 Varchar2) Is
    Select Object_Name 对象, Owner 所有者
    From All_Objects A
    Where Owner = v_所有者 And Instr(Object_Name, 'BIN$') = 0 And
          Object_Type In ('TABLE', 'VIEW', 'SEQUENCE', 'PROCEDURE', 'FUNCTION')
    Minus
    Select Synonym_Name, Table_Owner
    From All_Synonyms C
    Where Table_Owner = v_所有者 And (Owner = User Or Owner = 'PUBLIC');

Begin
  Select Count(Distinct 所有者) Into n_Cnt From Zlsystems;
  If n_Cnt > 1 Then
    Select Upper(所有者) Into v_所有者 From Zlsystems Where 编号 = 系统_In;
    --角色授权已限制所有者不能访问其它系统,所以,系统所有者不用创建私有同义词
    If v_所有者 != User Then
      Begin
        Select Substr(Banner, 6, 2) Into n_Cnt From V$version Where Substr(Banner, 1, 4) = 'CORE';
      Exception
        When Others Then
          n_Cnt := 9;
      End;
      If Nvl(n_Cnt, 10) > 9 Then
        For c_Syn In c_Delsyn10(v_所有者) Loop
          v_Sql := 'Drop Synonym ' || c_Syn.对象;
          Execute Immediate v_Sql;
        End Loop;
      Else
        For c_Syn In c_Delsyn9(v_所有者) Loop
          v_Sql := 'Drop Synonym ' || c_Syn.对象;
          Execute Immediate v_Sql;
        End Loop;
      End If;
    
      For c_Syn In c_Newsyn(v_所有者) Loop
        v_Sql := 'Create Synonym ' || c_Syn.对象 || ' For ' || c_Syn.所有者 || '.' || c_Syn.对象;
        Execute Immediate v_Sql;
      End Loop;
    End If;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Createsynonyms;
/