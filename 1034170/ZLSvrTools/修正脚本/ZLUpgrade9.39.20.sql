
---9.39.20����10.28.50ƥ��ʹ��----
--38719
Create Or Replace Function Zltools.f_List2str
(
  p_Strlist   In t_Strlist,
  p_Delimiter In Varchar2 Default ',',
  p_Distinct  In Number Default 1
) Return Varchar2 Is
  l_String Long;
  l_Add    Number;
  --���ܣ���һ���б���ת��Ϊһ��ȱʡ�Զ��ŷָ����ַ�����
  --����
  --With Test As
  --(Select a.���� As ����, c.���� As ��Ա
  --From ���ű� A, ������Ա B, ��Ա�� C
  --Where a.Id = b.����id And b.��Աid = c.Id
  --Order By ����, ��Ա)
  --Select ����, f_List2str(Cast(Collect(��Ա) As t_Strlist)) Tt From Test Group By ����

  --��֧��with��ʽ�������ʱ�ڴ���ᱨ��ORA-00932: �������Ͳ�һ��: ӦΪ -, ��ȴ��� -��
  --���磺With Test As (Select '�ڿ�' As ����,'����' As ��Ա From Dual Union All......)
  --  Select ����,f_List2str(cast(COLLECT(��Ա) as t_Strlist)) tt From Test Group By ����
Begin
  If p_Strlist.Count > 0 Then
    For I In p_Strlist.First .. p_Strlist.Last Loop
      l_Add := 0;
      If p_Distinct = 1 Then
        If Instr(',' || l_String || ',', ',' || p_Strlist(I) || ',') = 0 Then
          l_Add := 1;
        End If;
      Else
        l_Add := 1;
      End If;
      If l_Add = 1 Then
        If I != p_Strlist.First Then
          l_String := l_String || p_Delimiter;
        End If;
        l_String := l_String || p_Strlist(I);
      End If;
    End Loop;
  End If;
  Return l_String;
End f_List2str;
/
