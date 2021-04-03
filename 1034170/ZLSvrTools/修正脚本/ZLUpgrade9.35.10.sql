
--24274
Create Or Replace Procedure Zltools.Zl_Createsynonyms(系统_In In zlProgPrivs.系统%Type) Authid Current_User As
  v_Sql    Varchar2(2000);
  v_所有者 Varchar2(100);
  n_Cnt    Number(5);

  --非当前所有者的对象的私有同义词与当前所有者的对象相同则删除
  Cursor c_Delsyn(v_所有者 Varchar2) Is
    Select Synonym_Name 对象
    From User_Synonyms A
    Where Table_Owner != v_所有者 And Exists
     (Select 1
           From All_Objects B
           Where A.Synonym_Name = B.Object_Name And B.Owner = v_所有者 And
                 B.Object_Type In ('TABLE', 'VIEW', 'SEQUENCE', 'PROCEDURE', 'FUNCTION'));

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
      For c_Syn In c_Delsyn(v_所有者) Loop
        v_Sql := 'Drop Synonym ' || c_Syn.对象;
        Execute Immediate v_Sql;
      End Loop;
    
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
