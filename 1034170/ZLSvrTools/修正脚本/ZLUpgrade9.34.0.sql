-----------------------------------------------------------------
--Ϊ��ϲ�Ʒ�汾����9.33.10��Ϊ9.34.0
-----------------------------------------------------------------
Create Or Replace Procedure Zl_Createsynonyms
(
  ϵͳ_In     In zlProgPrivs.ϵͳ%Type,
  ���_In     In zlProgPrivs.���%Type,
  ϵͳ���_In Varchar2 --�Զ��ŷָ���ϵͳ���,ֻ�б���Ȩ��ʱ����0
) Authid Current_User As
  v_������ Varchar2(100);
  v_Sql    Varchar2(2000);

  --oracle 8.17�ϱ���ʹ����ʾ�α�����õ�����Ȩ�޷���All_Objects��ͼ
  Cursor c_All Is
    Select Object_Name ����, Owner ������
    From All_Objects A
    Where Owner In (Select Distinct Upper(������)
                    From Zlsystems
                    Where (Instr(',' || ϵͳ���_In || ',', ',' || ��� || ',') > 0 Or Instr(',' || ϵͳ���_In || ',', ',0,') > 0) And
                          Upper(������) <> User) And Instr(Object_Name, 'BIN$') = 0 And
          Object_Type In ('TABLE', 'VIEW', 'SEQUENCE', 'PROCEDURE', 'FUNCTION') And Not Exists
     (Select 1 From User_Synonyms C Where C.Table_Owner = A.Owner And C.Synonym_Name = A.Object_Name);

  --��ģ�������ʵĶ��󴴽�ͬ���
  Cursor c_Module Is
    Select ����, ������
    From zlProgPrivs A
    Where ϵͳ = ϵͳ_In And ��� = ���_In And Exists
     (Select 1
           From All_Objects B
           Where B.Owner = Upper(A.������) And B.Object_Name = Upper(A.����) And
                 B.Object_Type In ('TABLE', 'VIEW', 'SEQUENCE', 'PROCEDURE', 'FUNCTION')) And Not Exists
     (Select 1 From User_Synonyms C Where C.Table_Owner = Upper(A.������) And C.Synonym_Name = Upper(A.����));
Begin
  --a.��¼ʱ������Ȩ�޵����ж����ͬ���(���ܸ��ڽ���ģ��ʱ����,��Ϊ��Щģ�鲻��ͨ������̨���뼴�ɵ���)  
  If Nvl(ϵͳ_In, 0) = 0 Then
    For c_Syn In c_All Loop
      --���ܲ�ͬ��ϵͳ��ͬ���Ķ���,��ʱ��ɾ��,�ȴ�������һ��,����ģ��ʱ��ɾ��   
      Begin
        v_Sql := 'Create Synonym ' || c_Syn.���� || ' For ' || c_Syn.������ || '.' || c_Syn.����;
        Execute Immediate v_Sql;
      Exception
        When Others Then
          Null;
      End;
    End Loop;
  
  Else
    --b.����ģ��ʱ������ǰģ������ʵĶ����ͬ���(��Ȼ��¼ʱ������,�����������������ߵ�ͬ������)
    Select Upper(������) Into v_������ From Zlsystems Where ��� = ϵͳ_In;
    If v_������ != User Then
      For c_Syn In c_Module Loop
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
    End If;
  End If;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Createsynonyms;
/



