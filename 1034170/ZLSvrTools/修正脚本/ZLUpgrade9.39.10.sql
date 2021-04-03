--��10.28.40ƥ��ʹ��---


Create Or Replace Function zlTools.f_List2str
(
  p_Strlist   In t_Strlist,
  p_Delimiter In Varchar2 Default ','
) Return Varchar2 Is
  l_String Long;
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
      If I != p_Strlist.First Then
        l_String := l_String || p_Delimiter;
      End If;
      l_String := l_String || p_Strlist(I);
    End Loop;
  End If;
  Return l_String;
End f_List2str;
/

Create Public Synonym f_List2str for zlTools.f_List2str
/
Grant execute on zlTools.f_List2str to Public
/