-----------------------------------------------------------------
--为配合产品版本号由9.32升为9.33
-----------------------------------------------------------------
--13083、13161
Alter Table zlReports Add 打印方式 Number(1)
/

--14083
Create Or Replace Procedure Zl_Createsynonyms
(
  系统_In     In zlProgPrivs.系统%Type,
  序号_In     In zlProgPrivs.序号%Type,
  系统编号_In Varchar2 --以逗号分隔的系统编号,只有报表权限时传入0
) Authid Current_User As
  v_所有者 Varchar2(100);
  v_Sql    Varchar2(2000);
Begin

  --a.登录时创建非特定系统的对象的同义词
  If Nvl(系统_In, 0) = 0 Then
    For c_Syn In (Select Object_Name 对象, Owner 所有者
                  From All_Objects A
                  Where Owner In (Select Distinct Upper(所有者)
                                  From Zlsystems
                                  Where (Instr(',' || 系统编号_In || ',', ',' || 编号 || ',') > 0 Or
                                        Instr(',' || 系统编号_In || ',', ',0,') > 0) And Upper(所有者) <> User) And
                        Instr(Object_Name, 'BIN$') = 0 And
                        Object_Type In ('TABLE', 'VIEW', 'SEQUENCE', 'PROCEDURE', 'FUNCTION') And Not Exists
                   (Select 1
                         From zlProgPrivs B
                         Where B.系统 Is Not Null And B.所有者 = A.Owner And Upper(B.对象) = A.Object_Name) And Not Exists
                   (Select 1 From User_Synonyms C Where C.Table_Owner = A.Owner And C.Synonym_Name = A.Object_Name)) Loop
      --Exists比in和表连接方式更快
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
  
    --zl9AppTool要用到人员相关表,该部件不属于任何系统
    For c_Syn In (Select Object_Name 对象, Owner 所有者
                  From All_Objects A
                  Where Owner In (Select Distinct Upper(所有者)
                                  From Zlsystems
                                  Where (Instr(',' || 系统编号_In || ',', ',' || 编号 || ',') > 0 Or
                                        Instr(',' || 系统编号_In || ',', ',0,') > 0) And Upper(所有者) <> User) And
                        Object_Name In ('人员表', '部门表', '上机人员表', '部门人员', '人员性质说明') And Not Exists
                   (Select 1 From User_Synonyms C Where C.Table_Owner = A.Owner And C.Synonym_Name = A.Object_Name)) Loop
      --非共享安装,可能有多个人员相关表,所以要先删除再创建
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
    --b.进入模块时创建当前模块需访问的对象的同义词
    Select Upper(所有者) Into v_所有者 From Zlsystems Where 编号 = 系统_In;
    If v_所有者 != User Then
      For c_Syn In (Select Object_Name 对象, Owner 所有者
                    From All_Objects A
                    Where Owner = v_所有者 And Object_Type In ('TABLE', 'VIEW', 'SEQUENCE', 'PROCEDURE', 'FUNCTION') And
                          Exists (Select 1
                           From zlProgPrivs B
                           Where B.系统 = 系统_In And B.序号 = 序号_In And B.所有者 = A.Owner And
                                 Upper(B.对象) = A.Object_Name) And Not Exists
                     (Select 1 From User_Synonyms C Where C.Table_Owner = A.Owner And C.Synonym_Name = A.Object_Name)) Loop
        --Exists比in和表连接方式更快
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
    
      --虽然登录时创建了人员相关表,但可能不是当前系统的,所以要删除重建
      For c_Syn In (Select Object_Name 对象, Owner 所有者
                    From All_Objects A
                    Where Owner = v_所有者 And Object_Name In ('人员表', '部门表', '上机人员表', '部门人员', '人员性质说明') And Not Exists
                     (Select 1 From User_Synonyms C Where C.Table_Owner = A.Owner And C.Synonym_Name = A.Object_Name)) Loop
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

grant Execute on zltools.Zl_Createsynonyms to PUBLIC;
create or replace public synonym zl_CreateSynonyms For zltools.Zl_Createsynonyms;