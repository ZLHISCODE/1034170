-----------------------------------------------------------------
--为配合产品版本号由9.33.10升为9.33.20
-----------------------------------------------------------------

Create Or Replace Procedure Zl_Createsynonyms
(
  系统_In     In zlProgPrivs.系统%Type,
  序号_In     In zlProgPrivs.序号%Type,
  系统编号_In Varchar2 --以逗号分隔的系统编号,只有报表权限时传入0
) Authid Current_User As
  v_所有者 Varchar2(100);
  v_Sql    Varchar2(2000);

  --oracle 8.17上必须使用显示游标才能用调用者权限访问All_Objects视图
  Cursor c_All Is
    Select Object_Name 对象, Owner 所有者
    From All_Objects A
    Where Owner In (Select Distinct Upper(所有者)
                    From Zlsystems
                    Where (Instr(',' || 系统编号_In || ',', ',' || 编号 || ',') > 0 Or Instr(',' || 系统编号_In || ',', ',0,') > 0) And
                          Upper(所有者) <> User) And Instr(Object_Name, 'BIN$') = 0 And
          Object_Type In ('TABLE', 'VIEW', 'SEQUENCE', 'PROCEDURE', 'FUNCTION') And Not Exists
     (Select 1 From User_Synonyms C Where C.Table_Owner = A.Owner And C.Synonym_Name = A.Object_Name);

  --对模块所访问的对象创建同义词
  Cursor c_Module Is
    Select 对象, 所有者
    From zlProgPrivs A
    Where 系统 = 系统_In And 序号 = 序号_In And Exists
     (Select 1
           From All_Objects B
           Where B.Owner = A.所有者 And B.Object_Name = Upper(A.对象) And
                 B.Object_Type In ('TABLE', 'VIEW', 'SEQUENCE', 'PROCEDURE', 'FUNCTION')) And Not Exists
     (Select 1 From User_Synonyms C Where C.Table_Owner = A.所有者 And C.Synonym_Name = Upper(A.对象));

  Cursor c_User Is
    Select Object_Name 对象, Owner 所有者
    From All_Objects A
    Where Owner = v_所有者 And Object_Name In ('人员表', '部门表', '上机人员表', '部门人员', '人员性质说明') And Not Exists
     (Select 1 From User_Synonyms C Where C.Table_Owner = A.Owner And C.Synonym_Name = A.Object_Name);

Begin
  --a.登录时创建有权限的所有对象的同义词(不能改在进入模块时创建,因为有些模块不是通过导航台进入即可调用)  
  If Nvl(系统_In, 0) = 0 Then
    For c_Syn In c_All Loop
      --可能不同的系统有同名的对象,所以先删除
      Begin
        v_Sql := 'Drop Synonym ' || c_Syn.对象;
        Execute Immediate v_Sql;
      Exception
        When Others Then
          Null;
      End;
      v_Sql := 'Create Synonym ' || c_Syn.对象 || ' For ' || c_Syn.所有者 || '.' || c_Syn.对象;
      Execute Immediate v_Sql;
    End Loop;
  
  Else
    --b.进入模块时创建当前模块需访问的对象的同义词(虽然登录时创建了,但可能是其它所有者的同名对象)
    Select Upper(所有者) Into v_所有者 From Zlsystems Where 编号 = 系统_In;
    If v_所有者 != User Then
      For c_Syn In c_Module Loop
        --可能其它系统有同名的对象,所以先删除
        Begin
          v_Sql := 'Drop Synonym ' || c_Syn.对象;
          Execute Immediate v_Sql;
        Exception
          When Others Then
            Null;
        End;
        v_Sql := 'Create Synonym ' || c_Syn.对象 || ' For ' || c_Syn.所有者 || '.' || c_Syn.对象;
        Execute Immediate v_Sql;
      End Loop;
    
      --虽然登录时创建了人员相关表,但可能不是当前系统(所有者)的,所以要删除重建
      For c_Syn In c_User Loop
        Begin
          v_Sql := 'Drop Synonym ' || c_Syn.对象;
          Execute Immediate v_Sql;
        Exception
          When Others Then
            Null;
        End;
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