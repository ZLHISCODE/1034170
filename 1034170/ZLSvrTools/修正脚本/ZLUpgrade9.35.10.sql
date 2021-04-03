
--24274
Create Or Replace Procedure Zltools.Zl_Createsynonyms(ϵͳ_In In zlProgPrivs.ϵͳ%Type) Authid Current_User As
  v_Sql    Varchar2(2000);
  v_������ Varchar2(100);
  n_Cnt    Number(5);

  --�ǵ�ǰ�����ߵĶ����˽��ͬ����뵱ǰ�����ߵĶ�����ͬ��ɾ��
  Cursor c_Delsyn(v_������ Varchar2) Is
    Select Synonym_Name ����
    From User_Synonyms A
    Where Table_Owner != v_������ And Exists
     (Select 1
           From All_Objects B
           Where A.Synonym_Name = B.Object_Name And B.Owner = v_������ And
                 B.Object_Type In ('TABLE', 'VIEW', 'SEQUENCE', 'PROCEDURE', 'FUNCTION'));

  --���û���Ȩ���ʵĶ���,���ڵ�ǰϵͳ�����ߵ�,���û�й�����˽��ͬ���,�򴴽�˽��ͬ���
  --�������ڵ�ǰģ�������ʵĶ���,��Ϊ��������ģ���ģ���е�������ģ��
  Cursor c_Newsyn(v_������ Varchar2) Is
    Select Object_Name ����, Owner ������
    From All_Objects A
    Where Owner = v_������ And Instr(Object_Name, 'BIN$') = 0 And
          Object_Type In ('TABLE', 'VIEW', 'SEQUENCE', 'PROCEDURE', 'FUNCTION')
    Minus
    Select Synonym_Name, Table_Owner
    From All_Synonyms C
    Where Table_Owner = v_������ And (Owner = User Or Owner = 'PUBLIC');

Begin
  Select Count(Distinct ������) Into n_Cnt From Zlsystems;
  If n_Cnt > 1 Then
    Select Upper(������) Into v_������ From Zlsystems Where ��� = ϵͳ_In;
    --��ɫ��Ȩ�����������߲��ܷ�������ϵͳ,����,ϵͳ�����߲��ô���˽��ͬ���
    If v_������ != User Then
      For c_Syn In c_Delsyn(v_������) Loop
        v_Sql := 'Drop Synonym ' || c_Syn.����;
        Execute Immediate v_Sql;
      End Loop;
    
      For c_Syn In c_Newsyn(v_������) Loop
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
