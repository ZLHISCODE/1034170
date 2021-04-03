-----------------------------------------------------------------
--Ϊ��ϲ�Ʒ�汾����9.27��Ϊ9.28
-----------------------------------------------------------------
--���˺�:��Ҫ�ǽ��������ͼ��ת������
Create Table zltools.zlBakTables(
	ϵͳ Number(3),
	���� Varchar2(30))
/
Alter Table zltools.zlBakTables	Add Constraint zlBakTables_PK Primary Key (ϵͳ,����) USING INDEX PCTFREE 5
/
Alter Table zltools.zlBakTables Add Constraint zlBakTables_FK_ϵͳ Foreign Key (ϵͳ) References zlSystems(���) On Delete Cascade
/


Create Table zltools.zlBakSpaces(
	ϵͳ Number(3),
	��� Number(18),
	���� Varchar2(30),
	������ Varchar2(30),
	DB���� Varchar2(128),
	��ǰ Number(1),
	ֻ�� Number(1))
	PCTFREE 5 PCTUSED 90
/
Alter Table zltools.zlBakSpaces Add Constraint zlBakSpaces_PK Primary Key (ϵͳ,���) USING INDEX PCTFREE 5
/
Alter Table zltools.zlBakSpaces	Add Constraint zlBakSpaces_UQ_���� Unique (ϵͳ,����) USING INDEX PCTFREE 5
/
Alter Table zltools.zlBakSpaces Add Constraint zlBakSpaces_FK_ϵͳ Foreign Key (ϵͳ) References zlSystems(���) On Delete Cascade
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
	For r_User In(Select ������ From zlSystems) 
	Loop
			Execute Immediate 'Grant Select,Insert,Update,Delete on zlTools.zlBakTables to '||r_User.������||' With Grant Option';
			Execute Immediate 'Grant Select,Insert,Update,Delete on zlTools.zlBakSpaces to '||r_User.������||' With Grant Option';
	End Loop;
End;
/



Delete From zlSvrTools Where ���='0201'
/
Insert Into zlSvrTools(���,�ϼ�,����,���,˵��) values ('0201','02','����ת��','M',Null)
/




Create Or Replace type zlTools.t_StrList as Table of Varchar2(4000)
/

Create Or Replace Type zlTools.t_NumList as Table of Number
/



Create Or Replace Function zlTools.f_Str2List(Str_In In Varchar2) Return zlTools.t_Strlist As
  v_Str   Long Default Str_In || ',';
  v_Index Number;
  v_List  zlTools.t_Strlist := zlTools.t_Strlist();
  --���ܣ����ɶ��ŷָ��Ĳ������ŵ��ַ�����ת��Ϊ���ݱ�
  --������Str_In��:G0000123,G0000124,G0000125...
	--˵����
  --1����SQL������漰��IN(����1, ����2,��) ���Ӿ�ʱ��ʹ�����ַ�ʽ�����IN��������������(����ĿID)����ʹ��f_Num2List������������ַ�����(��NO)����ʹ��f_Str2List������
  --2��ʹ����������������ʹ������IN�Ӿ��SQL������ð󶨱�����ʹ������������ʱ�����IN�Ӿ��Ƕ�Ӧ�������ֶΣ�Ҳͬ��������������(��NO IN(��))��
  --3��ʹ������������ʱ����Ҫ��SQL����м��롰/*+ Rule*/����ʾ���Ա���CBO�µ��������⡣
  --4�����ֵ���ʾ����ע����������ǰ��Ҫ��zlToolsǰ׺��
  --Select /*+ Rule*/ * From ���˷��ü�¼ Where NO In (Select * From Table(Cast(f_Str2list('A01,A02,A03') As zlTools.t_Strlist)));
  --Select /*+ Rule*/ A.* From ���˷��ü�¼ A, Table(Cast(f_Str2list('A01,A02,A03') As zlTools.t_Strlist)) B Where A.NO = B.Column_Value;
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
  --���ܣ����ɶ��ŷָ�����������ת��Ϊ���ݱ�
  --������Str_In��:73265,73266,73267....
	--˵����
  --1����SQL������漰��IN(����1, ����2,��) ���Ӿ�ʱ��ʹ�����ַ�ʽ�����IN��������������(����ĿID)����ʹ��f_Num2List������������ַ�����(��NO)����ʹ��f_Str2List������
  --2��ʹ����������������ʹ������IN�Ӿ��SQL������ð󶨱�����ʹ������������ʱ�����IN�Ӿ��Ƕ�Ӧ�������ֶΣ�Ҳͬ��������������(��NO IN(��))��
  --3��ʹ������������ʱ����Ҫ��SQL����м��롰/*+ Rule*/����ʾ���Ա���CBO�µ��������⡣
  --4�����ֵ���ʾ����ע����������ǰ��Ҫ��zlToolsǰ׺��
  --Select /*+ Rule*/ * From ���˷��ü�¼ Where NO In (Select * From Table(Cast(f_Str2list('A01,A02,A03') As zlTools.t_Strlist)));
  --Select /*+ Rule*/ A.* From ���˷��ü�¼ A, Table(Cast(f_Str2list('A01,A02,A03') As zlTools.t_Strlist)) B Where A.NO = B.Column_Value;
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
--  �¶�(2007-03-01):Ϊ�����Ӷ�Ȩ�޹�ϵ��ȱʡֵ�Ĵ��������˹����ߵ����ݽṹ��
Create table zlTools.zlProgRelas
(
  ϵͳ     NUMBER(5) not null,
  ���     NUMBER(18) not null,
  ����     VARCHAR2(30) not null,
  ���     NUMBER(5) not null,
  ��ϵ     NUMBER(3),
  ����     NUMBER(1),
  �����ϵ NUMBER(1))
  PCTFREE 5 PCTUSED 90
  Cache Storage(Buffer_Pool Keep)
/
Alter Table zlTools.zlProgRelas add Constraint zlProgRelas_PK Primary Key (ϵͳ, ���, ����, ���) using index  PCTFREE 5
/
Alter Table zlTools.zlProgRelas add Constraint zlProgRelas_FK_��� Foreign Key (ϵͳ, ���, ����) References zlProgFuncs (ϵͳ, ���, ����) On Delete Cascade
/
Alter Table zlTools.zlProgRelas Add Constraint zlProgRelas_CK_���� Check (���� IN(0,1))
/
Alter Table zlTools.zlProgRelas Add Constraint zlProgRelas_CK_�����ϵ Check (�����ϵ IN(0,1))
/
CREATE PUBLIC SYNONYM zlProgRelas for zlTools.zlProgRelas
/
GRANT SELECT ON zlTools.zlProgRelas TO PUBLIC 
/
Begin
  For r_User In (Select ������ From zlTools.zlSystems) Loop
    Execute Immediate 'Grant Select,Insert,Update,Delete on zlTools.zlProgRelas  to ' || r_User.������ ||
                      ' With Grant Option';
  End Loop;
End;
/
Alter Table zlTools.zlProgFuncs Add ȱʡֵ Number(1) Default 1
/
