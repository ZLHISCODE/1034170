--与10.28.40匹配使用---


Create Or Replace Function zlTools.f_List2str
(
  p_Strlist   In t_Strlist,
  p_Delimiter In Varchar2 Default ','
) Return Varchar2 Is
  l_String Long;
  --功能：将一个列表集合转换为一个缺省以逗号分隔的字符串。
  --例：
  --With Test As
  --(Select a.名称 As 科室, c.姓名 As 人员
  --From 部门表 A, 部门人员 B, 人员表 C
  --Where a.Id = b.部门id And b.人员id = c.Id
  --Order By 科室, 人员)
  --Select 科室, f_List2str(Cast(Collect(人员) As t_Strlist)) Tt From Test Group By 科室

  --不支持with方式构造的临时内存表，会报错：ORA-00932: 数据类型不一致: 应为 -, 但却获得 -。
  --例如：With Test As (Select '内科' As 科室,'张三' As 人员 From Dual Union All......)
  --  Select 科室,f_List2str(cast(COLLECT(人员) as t_Strlist)) tt From Test Group By 科室
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