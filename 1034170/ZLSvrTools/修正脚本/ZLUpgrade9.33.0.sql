-----------------------------------------------------------------
--Ϊ��ϲ�Ʒ�汾����9.32��Ϊ9.33
-----------------------------------------------------------------
--13083��13161
Alter Table zlReports Add ��ӡ��ʽ Number(1)
/

--14083
Create Or Replace Procedure Zl_Createsynonyms
(
  ϵͳ_In     In zlProgPrivs.ϵͳ%Type,
  ���_In     In zlProgPrivs.���%Type,
  ϵͳ���_In Varchar2 --�Զ��ŷָ���ϵͳ���,ֻ�б���Ȩ��ʱ����0
) Authid Current_User As
  v_������ Varchar2(100);
  v_Sql    Varchar2(2000);
Begin

  --a.��¼ʱ�������ض�ϵͳ�Ķ����ͬ���
  If Nvl(ϵͳ_In, 0) = 0 Then
    For c_Syn In (Select Object_Name ����, Owner ������
                  From All_Objects A
                  Where Owner In (Select Distinct Upper(������)
                                  From Zlsystems
                                  Where (Instr(',' || ϵͳ���_In || ',', ',' || ��� || ',') > 0 Or
                                        Instr(',' || ϵͳ���_In || ',', ',0,') > 0) And Upper(������) <> User) And
                        Instr(Object_Name, 'BIN$') = 0 And
                        Object_Type In ('TABLE', 'VIEW', 'SEQUENCE', 'PROCEDURE', 'FUNCTION') And Not Exists
                   (Select 1
                         From zlProgPrivs B
                         Where B.ϵͳ Is Not Null And B.������ = A.Owner And Upper(B.����) = A.Object_Name) And Not Exists
                   (Select 1 From User_Synonyms C Where C.Table_Owner = A.Owner And C.Synonym_Name = A.Object_Name)) Loop
      --Exists��in�ͱ����ӷ�ʽ����
      --���ܲ�ͬ��ϵͳ��ͬ���Ķ���,������ɾ��
      Begin
        v_Sql := 'Drop Synonym ' || c_Syn.����;
        Execute Immediate v_Sql;
      Exception
        When Others Then
          Null;
      End;
      v_Sql := 'Create Synonym ' || c_Syn.���� || ' For ' || c_Syn.������ || '.' || c_Syn.����;
      Execute Immediate v_Sql;
    End Loop;
  
    --zl9AppToolҪ�õ���Ա��ر�,�ò����������κ�ϵͳ
    For c_Syn In (Select Object_Name ����, Owner ������
                  From All_Objects A
                  Where Owner In (Select Distinct Upper(������)
                                  From Zlsystems
                                  Where (Instr(',' || ϵͳ���_In || ',', ',' || ��� || ',') > 0 Or
                                        Instr(',' || ϵͳ���_In || ',', ',0,') > 0) And Upper(������) <> User) And
                        Object_Name In ('��Ա��', '���ű�', '�ϻ���Ա��', '������Ա', '��Ա����˵��') And Not Exists
                   (Select 1 From User_Synonyms C Where C.Table_Owner = A.Owner And C.Synonym_Name = A.Object_Name)) Loop
      --�ǹ���װ,�����ж����Ա��ر�,����Ҫ��ɾ���ٴ���
      Begin
        v_Sql := 'Drop Synonym ' || c_Syn.����;
        Execute Immediate v_Sql;
      Exception
        When Others Then
          Null;
      End;
      v_Sql := 'Create Synonym ' || c_Syn.���� || ' For ' || c_Syn.������ || '.' || c_Syn.����;
      Execute Immediate v_Sql;
    End Loop;
  
  Else
    --b.����ģ��ʱ������ǰģ������ʵĶ����ͬ���
    Select Upper(������) Into v_������ From Zlsystems Where ��� = ϵͳ_In;
    If v_������ != User Then
      For c_Syn In (Select Object_Name ����, Owner ������
                    From All_Objects A
                    Where Owner = v_������ And Object_Type In ('TABLE', 'VIEW', 'SEQUENCE', 'PROCEDURE', 'FUNCTION') And
                          Exists (Select 1
                           From zlProgPrivs B
                           Where B.ϵͳ = ϵͳ_In And B.��� = ���_In And B.������ = A.Owner And
                                 Upper(B.����) = A.Object_Name) And Not Exists
                     (Select 1 From User_Synonyms C Where C.Table_Owner = A.Owner And C.Synonym_Name = A.Object_Name)) Loop
        --Exists��in�ͱ����ӷ�ʽ����
        --��������ϵͳ��ͬ���Ķ���,������ɾ��
        Begin
          v_Sql := 'Drop Synonym ' || c_Syn.����;
          Execute Immediate v_Sql;
        Exception
          When Others Then
            Null;
        End;
        v_Sql := 'Create Synonym ' || c_Syn.���� || ' For ' || c_Syn.������ || '.' || c_Syn.����;
        Execute Immediate v_Sql;
      End Loop;
    
      --��Ȼ��¼ʱ��������Ա��ر�,�����ܲ��ǵ�ǰϵͳ��,����Ҫɾ���ؽ�
      For c_Syn In (Select Object_Name ����, Owner ������
                    From All_Objects A
                    Where Owner = v_������ And Object_Name In ('��Ա��', '���ű�', '�ϻ���Ա��', '������Ա', '��Ա����˵��') And Not Exists
                     (Select 1 From User_Synonyms C Where C.Table_Owner = A.Owner And C.Synonym_Name = A.Object_Name)) Loop
        Begin
          v_Sql := 'Drop Synonym ' || c_Syn.����;
          Execute Immediate v_Sql;
        Exception
          When Others Then
            Null;
        End;
        v_Sql := 'Create Synonym ' || c_Syn.���� || ' For ' || c_Syn.������ || '.' || c_Syn.����;
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