Create Or Replace Function zlTools.f_Str2list
(
  Str_In   In Varchar2,
  Split_In In Varchar2 := ','
) Return t_Strlist
  Pipelined As
  v_Str Long;
  P     Number;
  --���ܣ����ɶ��ŷָ��Ĳ������ŵ��ַ�����ת��Ϊ�������ݱ�
  --������Str_In,��:G0000123,G0000124,G0000125...,Split_In,�ָ���,ȱʡΪ,��
  --˵����
  --1����SQL������漰��IN(����1, ����2,��) ���Ӿ�ʱʹ�����ַ�ʽ�Ա����ð󶨱�����
  --2��ʹ������������ʱ����Ҫ��SQL����м��롰/*+ Rule*/����ʾ����ΪCbo����ʱ�ڴ��û��ͳ������,��
  --3�����ֵ���ʾ��
  --Select /*+ Rule*/ * From ������ü�¼ Where NO In (Select * From Table(f_Str2list('A01,A02,A03'));
  --Select /*+ Rule*/ A.* From ������ü�¼ A, Table(f_Str2list('A01,A02,A03')) B Where A.NO = B.Column_Value;
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



--����9i��10,��ȡ��ͬ���α꣨c_Delsyn���Ż���ѯ����
Create Or Replace Procedure zlTools.Zl_Createsynonyms(ϵͳ_In In Zlprogprivs.ϵͳ%Type) Authid Current_User As
  v_Sql    Varchar2(2000);
  v_������ Varchar2(100);
  n_Cnt    Number(5);

  --�ǵ�ǰ�����ߵĶ����˽��ͬ����뵱ǰ�����ߵĶ�����ͬ��ɾ��
  Cursor c_Delsyn9(v_������ Varchar2) Is
    Select Synonym_Name ����
    From User_Synonyms A
    Where Table_Owner != v_������ And Exists
     (Select 1
           From All_Objects B
           Where a.Synonym_Name = b.Object_Name And b.Owner = v_������ And
                 b.Object_Type In ('TABLE', 'VIEW', 'SEQUENCE', 'PROCEDURE', 'FUNCTION'));

  Cursor c_Delsyn10(v_������ Varchar2) Is
    Select a.Synonym_Name ����
    From User_Synonyms A, All_Objects B
    Where a.Table_Owner != v_������ And a.Synonym_Name = b.Object_Name And b.Owner = v_������ And
          b.Object_Type In ('TABLE', 'VIEW', 'SEQUENCE', 'PROCEDURE', 'FUNCTION');

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
      Begin
        Select Substr(Banner, 6, 2) Into n_Cnt From V$version Where Substr(Banner, 1, 4) = 'CORE';
      Exception
        When Others Then
          n_Cnt := 9;
      End;
      If Nvl(n_Cnt, 10) > 9 Then
        For c_Syn In c_Delsyn10(v_������) Loop
          v_Sql := 'Drop Synonym ' || c_Syn.����;
          Execute Immediate v_Sql;
        End Loop;
      Else
        For c_Syn In c_Delsyn9(v_������) Loop
          v_Sql := 'Drop Synonym ' || c_Syn.����;
          Execute Immediate v_Sql;
        End Loop;
      End If;
    
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