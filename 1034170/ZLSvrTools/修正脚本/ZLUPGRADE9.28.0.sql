-----------------------------------------------------------------
--为配合产品版本号由9.27升为9.28
-----------------------------------------------------------------
--刘兴宏:主要是解决基本视图的转储问题
Create Table zltools.zlBakTables(
	系统 Number(3),
	表名 Varchar2(30))
/
Alter Table zltools.zlBakTables	Add Constraint zlBakTables_PK Primary Key (系统,表名) USING INDEX PCTFREE 5
/
Alter Table zltools.zlBakTables Add Constraint zlBakTables_FK_系统 Foreign Key (系统) References zlSystems(编号) On Delete Cascade
/


Create Table zltools.zlBakSpaces(
	系统 Number(3),
	编号 Number(18),
	名称 Varchar2(30),
	所有者 Varchar2(30),
	DB连接 Varchar2(128),
	当前 Number(1),
	只读 Number(1))
	PCTFREE 5 PCTUSED 90
/
Alter Table zltools.zlBakSpaces Add Constraint zlBakSpaces_PK Primary Key (系统,编号) USING INDEX PCTFREE 5
/
Alter Table zltools.zlBakSpaces	Add Constraint zlBakSpaces_UQ_名称 Unique (系统,名称) USING INDEX PCTFREE 5
/
Alter Table zltools.zlBakSpaces Add Constraint zlBakSpaces_FK_系统 Foreign Key (系统) References zlSystems(编号) On Delete Cascade
/


CREATE PUBLIC SYNONYM zlBakSpaces for zlTools.zlBakSpaces
/

CREATE PUBLIC SYNONYM zlBakTables for zlTools.zlBakTables
/
 

GRANT SELECT ON zlTools.zlBakSpaces TO PUBLIC 
/

GRANT SELECT ON zlTools.zlBakTables TO PUBLIC 
/


Begin
	For r_User In(Select 所有者 From zlSystems) 
	Loop
			Execute Immediate 'Grant Select,Insert,Update,Delete on zlTools.zlBakTables to '||r_User.所有者||' With Grant Option';
			Execute Immediate 'Grant Select,Insert,Update,Delete on zlTools.zlBakSpaces to '||r_User.所有者||' With Grant Option';
	End Loop;
End;
/



Delete From zlSvrTools Where 编号='0201'
/
Insert Into zlSvrTools(编号,上级,标题,快键,说明) values ('0201','02','数据转移','M',Null)
/




Create Or Replace type zlTools.t_StrList as Table of Varchar2(4000)
/

Create Or Replace Type zlTools.t_NumList as Table of Number
/



Create Or Replace Function zlTools.f_Str2List(Str_In In Varchar2) Return zlTools.t_Strlist As
  v_Str   Long Default Str_In || ',';
  v_Index Number;
  v_List  zlTools.t_Strlist := zlTools.t_Strlist();
  --功能：将由逗号分隔的不带引号的字符序列转换为数据表
  --参数：Str_In如:G0000123,G0000124,G0000125...
	--说明：
  --1．当SQL语句中涉及“IN(常量1, 常量2,…) ”子句时请使用这种方式，如果IN常量是数字类型(如项目ID)，则使用f_Num2List函数，如果是字符类型(如NO)，则使用f_Str2List函数。
  --2．使用这两个函数可以使带这种IN子句的SQL语句利用绑定变量。使用这两个函数时，如果IN子句是对应的索引字段，也同样可以利用索引(如NO IN(…))。
  --3．使用这两个函数时，需要在SQL语句中加入“/*+ Rule*/”提示，以避免CBO下的性能问题。
  --4．两种调用示例，注意在类型名前需要加zlTools前缀：
  --Select /*+ Rule*/ * From 病人费用记录 Where NO In (Select * From Table(Cast(f_Str2list('A01,A02,A03') As zlTools.t_Strlist)));
  --Select /*+ Rule*/ A.* From 病人费用记录 A, Table(Cast(f_Str2list('A01,A02,A03') As zlTools.t_Strlist)) B Where A.NO = B.Column_Value;
Begin
  Loop
    v_Index := Instr(v_Str, ',');
    Exit When(Nvl(v_Index, 0) = 0);
    v_List.Extend;
    v_List(v_List.Count) := Trim(Substr(v_Str, 1, v_Index - 1));
    v_Str := Substr(v_Str, v_Index + 1);
  End Loop;
  Return v_List;
End;
/

Create Or Replace Function zlTools.f_Num2List(Str_In In Varchar2) Return zlTools.t_Numlist As
  v_Str   Long Default Str_In || ',';
  v_Index Number;
  v_List  zlTools.t_Numlist := zlTools.t_Numlist();
  --功能：将由逗号分隔的数字序列转换为数据表
  --参数：Str_In如:73265,73266,73267....
	--说明：
  --1．当SQL语句中涉及“IN(常量1, 常量2,…) ”子句时请使用这种方式，如果IN常量是数字类型(如项目ID)，则使用f_Num2List函数，如果是字符类型(如NO)，则使用f_Str2List函数。
  --2．使用这两个函数可以使带这种IN子句的SQL语句利用绑定变量。使用这两个函数时，如果IN子句是对应的索引字段，也同样可以利用索引(如NO IN(…))。
  --3．使用这两个函数时，需要在SQL语句中加入“/*+ Rule*/”提示，以避免CBO下的性能问题。
  --4．两种调用示例，注意在类型名前需要加zlTools前缀：
  --Select /*+ Rule*/ * From 病人费用记录 Where NO In (Select * From Table(Cast(f_Str2list('A01,A02,A03') As zlTools.t_Strlist)));
  --Select /*+ Rule*/ A.* From 病人费用记录 A, Table(Cast(f_Str2list('A01,A02,A03') As zlTools.t_Strlist)) B Where A.NO = B.Column_Value;
Begin
  Loop
    v_Index := Instr(v_Str, ',');
    Exit When(Nvl(v_Index, 0) = 0);
    v_List.Extend;
    v_List(v_List.Count) := To_Number(Trim(Substr(v_Str, 1, v_Index - 1)));
    v_Str := Substr(v_Str, v_Index + 1);
  End Loop;
  Return v_List;
End;
/

Grant Execute on zlTools.t_StrList to Public
/

Grant Execute on zlTools.t_NumList to Public
/

Grant Execute on zlTools.f_Str2List to Public
/

Grant Execute on zlTools.f_Num2List to Public
/

Create Public Synonym f_Str2List For zlTools.f_Str2List
/

Create Public Synonym f_Num2List For zlTools.f_Num2List
/

-------------------------------------------------------------------------------------
--  陈东(2007-03-01):为了增加对权限关系，缺省值的处理，调整了管理工具的数据结构。
Create table zlTools.zlProgRelas
(
  系统     NUMBER(5) not null,
  序号     NUMBER(18) not null,
  功能     VARCHAR2(30) not null,
  组号     NUMBER(5) not null,
  关系     NUMBER(3),
  主项     NUMBER(1),
  主项关系 NUMBER(1))
  PCTFREE 5 PCTUSED 90
  Cache Storage(Buffer_Pool Keep)
/
Alter Table zlTools.zlProgRelas add Constraint zlProgRelas_PK Primary Key (系统, 序号, 功能, 组号) using index  PCTFREE 5
/
Alter Table zlTools.zlProgRelas add Constraint zlProgRelas_FK_序号 Foreign Key (系统, 序号, 功能) References zlProgFuncs (系统, 序号, 功能) On Delete Cascade
/
Alter Table zlTools.zlProgRelas Add Constraint zlProgRelas_CK_主项 Check (主项 IN(0,1))
/
Alter Table zlTools.zlProgRelas Add Constraint zlProgRelas_CK_主项关系 Check (主项关系 IN(0,1))
/
CREATE PUBLIC SYNONYM zlProgRelas for zlTools.zlProgRelas
/
GRANT SELECT ON zlTools.zlProgRelas TO PUBLIC 
/
Begin
  For r_User In (Select 所有者 From zlTools.zlSystems) Loop
    Execute Immediate 'Grant Select,Insert,Update,Delete on zlTools.zlProgRelas  to ' || r_User.所有者 ||
                      ' With Grant Option';
  End Loop;
End;
/
Alter Table zlTools.zlProgFuncs Add 缺省值 Number(1) Default 1
/
