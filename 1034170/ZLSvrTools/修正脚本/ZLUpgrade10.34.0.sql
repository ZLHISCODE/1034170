----10.34.0---��9.45.0
--00000:��˶,2014-10-28,ɾ��ZLRegInfo��״̬��Ϣ
   Delete From Zlreginfo Where ��Ŀ = '��״̬';
--72631:���Ʊ�,2014-04-30,�����ۺ������ԸĽ�
--4.1.  ���������ѯ������
CREATE TABLE zlTools.zlRPTRelation(
    ����ID NUMBER(18),
    Ԫ��ID NUMBER(18),
    ��������ID	NUMBER(18),
    ������	VARCHAR2(50),
    ����ֵ��Դ	VARCHAR2(255)
    )PCTFREE 5
    Cache Storage(Buffer_Pool Keep);
ALTER TABLE zlTools.zlRPTRelation ADD CONSTRAINT zlRPTRelation_PK PRIMARY KEY (����ID,��������ID,Ԫ��ID,������) USING INDEX PCTFREE 5;
ALTER TABLE zlTools.zlRPTRelation ADD CONSTRAINT zlRPTRelation_FK_Ԫ��ID FOREIGN KEY(Ԫ��ID) REFERENCES zlRPTItems(ID) ON DELETE CASCADE;
ALTER TABLE zlTools.zlRPTRelation ADD CONSTRAINT zlRPTRelation_FK_��������ID FOREIGN KEY(��������ID) REFERENCES zlReports(ID) ON DELETE CASCADE;
ALTER TABLE zlTools.zlRPTRelation ADD CONSTRAINT zlRPTRelation_FK_����ID FOREIGN KEY(����ID) REFERENCES zlTools.zlReports(ID) ON DELETE CASCADE;
Create Index zlRPTRelation_IX_Ԫ��ID ON zlTools.zlRPTRelation(Ԫ��ID) PCTFREE 5;
Create Index zlRPTRelation_IX_��������ID ON zlTools.zlRPTRelation(��������ID) PCTFREE 5;
--4.2.	��������Դ��ʷ��¼������
CREATE TABLE zlTools.zlRPTSQLsHistory(
    ����ID	NUMBER(18),
    ����Դ����	VARCHAR2(20),
    �޸���	VARCHAR2(100),
    �޸�ʱ��	DATE,
    �к�	NUMBER(5),
    ����	VARCHAR2(4000)
    )PCTFREE 5;
ALTER TABLE zlTools.zlRPTSQLsHistory ADD CONSTRAINT zlRPTSQLsHistory_PK PRIMARY KEY (����ID,����Դ����,�޸�ʱ��,�к�) USING INDEX PCTFREE 5;
ALTER TABLE zlTools.zlRPTSQLsHistory ADD CONSTRAINT zlRPTSQLsHistory_FK_����ID FOREIGN KEY(����ID) REFERENCES zlTools.zlReports(ID) ON DELETE CASCADE;
--4.3.	Zlreports�����ֶΣ���ѯ��ʼʱ�� DATE����ѯ����ʱ�� DATE
ALTER TABLE zlTools.zlReports add(��ֹ��ʼʱ�� DATE,��ֹ����ʱ�� DATE);
--4.4.	zlrptPars�����ֶΣ����� Number(1)
ALTER TABLE zlTools.zlrptPars add(���� Number(1));
--4.5.	����������������
CREATE TABLE zlTools.zlRPTColProterty(
    ����ID	NUMBER(18),
    Ԫ��ID	NUMBER(18),
    ��������	VARCHAR2(50),
    �����ֶ�	VARCHAR2(255),
    ������ϵ	VARCHAR2(50),
    ����ֵ	VARCHAR2(255),
    ������ɫ	NUMBER(18),
    ������ɫ	NUMBER(18),
    �Ƿ�Ӵ�    NUMBER(1),
    �Ƿ�����Ӧ�� NUMBER(1)
    )PCTFREE 5
    Cache Storage(Buffer_Pool Keep);
ALTER TABLE zlTools.zlRPTColProterty ADD CONSTRAINT zlRPTColProterty_PK PRIMARY KEY (����ID,Ԫ��ID,��������) USING INDEX PCTFREE 5;
ALTER TABLE zlTools.zlRPTColProterty ADD CONSTRAINT zlRPTColProterty_FK_Ԫ��ID FOREIGN KEY(Ԫ��ID) REFERENCES zlTools.zlRPTItems(ID) ON DELETE CASCADE;
ALTER TABLE zlTools.zlRPTColProterty ADD CONSTRAINT zlRPTColProterty_FK_����ID FOREIGN KEY(����ID) REFERENCES zlTools.zlReports(ID) ON DELETE CASCADE;
Create Index zlRPTColProterty_IX_Ԫ��ID ON zlTools.zlRPTColProterty(Ԫ��ID) PCTFREE 5;

--70590:��˶,2014-04-30,ȱʡʹ�ø��Ի����
Update zlParameters Set ȱʡֵ = '1' Where ϵͳ Is Null And ������ = 'ʹ�ø��Ի����';

--74440:��˶,2014-07-08,ģ�������Ȩ
create table ZLTOOLS.ZLModuleRelas
(
ϵͳ  Number(5),
ģ��  Number(18),  
����  Varchar2(30),
���ϵͳ  Number(5),
���ģ��  Number(18),  
�������  Number(1), 
��ع���  Varchar2(30),
ȱʡֵ    Number(1)
)
tablespace ZLTOOLSTBS;
alter table ZLTOOLS.zlprograms add ����  Number(1);
alter table ZLTOOLS.ZLModuleRelas Modify ϵͳ  constraint ZLModuleRelas_NN_ϵͳ   not  null;
alter table ZLTOOLS.ZLModuleRelas Modify ģ��  constraint ZLModuleRelas_NN_ģ��   not  null;
alter table ZLTOOLS.ZLModuleRelas Modify ���ģ��  constraint ZLModuleRelas_NN_���ģ��   not  null;
alter table ZLTOOLS.ZLModuleRelas add constraint ZLModuleRelas_UQ_���ģ�� Unique(ϵͳ,ģ��,����,���ϵͳ,���ģ��,��ع���) using index tablespace ZLTOOLSTBS;
alter table ZLTOOLS.ZLModuleRelas add constraint ZLModuleRelas_FK_ģ�� foreign key(ϵͳ,ģ��) references  ZLTOOLS.zlprograms(ϵͳ,���) on delete cascade;

--00000:������,2014-03-04,������������δ�Ǽ�BUG(2014-6-16����)
Drop Function zlTools.f_Get_Stream_State;

Create Or Replace Function zlTools.Zl_Checkobject
(
  n_Type        In Number, --1=��,2=�ֶ�,3=Լ��,4=����
  v_Object_Name In Varchar2,
  v_Table_Name  In Varchar2 := Null --����n_Type=2ʱ����Ҫ����
) Return Number Authid Current_User As
  --���ܣ���ִ���ߵ���ݼ��ָ�����ָ�������Ƿ����
  --����ֵ��>0��ʾ���ڣ�0��ʾ������
  n_Count Number(5);
Begin
  If n_Type = 1 Then
    Select Count(1) Into n_Count From User_Tables Where Table_Name = Upper(v_Table_Name);
  
  Elsif n_Type = 2 Then
    Select Count(1)
    Into n_Count
    From User_Tab_Columns
    Where Table_Name = Upper(v_Table_Name) And Column_Name = Upper(v_Object_Name);
  
  Elsif n_Type = 3 Then
    Select Count(1) Into n_Count From User_Constraints Where Constraint_Name = Upper(v_Object_Name);
  
  Elsif n_Type = 4 Then
    Select Count(1) Into n_Count From User_Indexes Where Index_Name = Upper(v_Object_Name);
  End If;

  Return n_Count;
End Zl_Checkobject;
/


--00000:���,2014-03-27,��Ʒ�ض���Ȩ���ƸĽ���δ�Ǽ�����
Create Or Replace Function zlTools.f_Reg_Menu
(
  Menu_Group_In  In Zlmenus.���%Type := Null, --����ѡ��Ĳ˵����
  System_List_In In Varchar2, --���λỰ�漰��Ӧ��ϵͳ
  Part_List_In   In Varchar2 --�Զ��ŷָ��ı�����ִ�в����б�
) Return t_Menu_Rowset Is
  t_Return t_Menu_Rowset := t_Menu_Rowset();
  t_Middle t_Menu_Rowset := t_Menu_Rowset();

  v_Parts   Varchar2(32767);
  t_Parts   t_Reg_Rowset := t_Reg_Rowset();
  v_Systems Varchar2(32767);
  t_Systems t_Reg_Rowset := t_Reg_Rowset();

Begin
  --���������γ����������
  v_Parts := Upper(Part_List_In) || ',';
  While v_Parts Is Not Null Loop
    t_Parts.Extend;
    t_Parts(t_Parts.Count) := t_Reg_Record(Null, Null, Substr(v_Parts, 1, Instr(v_Parts, ',') - 1));
    v_Parts := Trim(Substr(v_Parts, Instr(v_Parts, ',') + 1));
  End Loop;
  t_Parts.Extend;
  t_Parts(t_Parts.Count) := t_Reg_Record(Null, Null, 'ZL9REPORT');
  v_Systems := System_List_In || ',';
  While v_Systems Is Not Null Loop
    t_Systems.Extend;
    t_Systems(t_Systems.Count) := t_Reg_Record(Null, To_Number(Substr(v_Systems, 1, Instr(v_Systems, ',') - 1)), Null);
    v_Systems := Trim(Substr(v_Systems, Instr(v_Systems, ',') + 1));
  End Loop;
  t_Systems.Extend;
  t_Systems(t_Systems.Count) := t_Reg_Record(Null, 0, Null);

  --�˵����ݻ�ȡ��
  Select t_Menu_Record(m.���, m.Id, m.�ϼ�id, m.����, m.�̱���, m.���, m.˵��, m.ģ��, m.ϵͳ, m.ͼ��, p.����, 0) Bulk Collect
  Into t_Middle
  From (Select Level As ���, ID, �ϼ�id, ����, �̱���, ���, ˵��, ģ��, ϵͳ, ͼ��
         From zlMenus
         Where ��� = Menu_Group_In
         Start With �ϼ�id Is Null
         Connect By Prior ID = �ϼ�id) M,
       (Select Distinct p.ϵͳ, p.���, p.����
         From zlPrograms P, zlProgFuncs F, zlRegFunc R, zlRPTGroups X, Table(Cast(t_Parts As t_Reg_Rowset)) C,
              Table(Cast(t_Systems As t_Reg_Rowset)) S,
              (Select 1 As ���� From Sys.Dba_Role_Privs Where Granted_Role = 'DBA' And Grantee = User) A,
              (Select Decode(Count(*), 0, 0, Null, 0, 1) As ���
                From zlSystems
                Where Upper(������) = User
                Union All
                Select ���
                From zlSystems
                Where Upper(������) = User) O,
              (Select Distinct g.ϵͳ, g.���
                From zlRoleGrant G, Sys.Dba_Role_Privs R
                Where g.��ɫ = r.Granted_Role And r.Grantee = User) G
         Where Nvl(f.ϵͳ, 0) = Nvl(p.ϵͳ, 0) And f.��� = p.��� And Trunc(f.ϵͳ / 100) = r.ϵͳ(+) And f.��� = r.���(+) And
               f.���� = r.����(+) And
               (r.���� Is Null And f.ϵͳ Is Null Or r.���� Is Not Null And r.���� = '����' Or
                r.���� Is Not Null And x.����id Is Not Null Or r.���� Is Null And (p.��� Between 10000 And 19999)) And
               p.ϵͳ = x.ϵͳ(+) And p.��� = x.����id(+) And Upper(p.����) = c.Text And Nvl(p.ϵͳ, 0) = s.Prog And
               p.��� = p.��� * a.����(+) And Nvl(p.ϵͳ, 1) = o.���(+) And Nvl(p.ϵͳ, 0) = Nvl(g.ϵͳ(+), 0) And p.��� = g.���(+) And
               (a.���� Is Not Null Or o.��� Is Not Null Or g.��� Is Not Null)) P
  Where Nvl(m.ϵͳ, 0) = Nvl(p.ϵͳ(+), 0) And m.ģ�� = p.���(+) And (m.ģ�� Is Null Or m.ģ�� Is Not Null And p.��� Is Not Null)
  Order By m.��� Desc;

  --�������¼���ִ�еĲ˵���Ŀ
  For n_Child In 1 .. t_Middle.Count Loop
    If t_Middle(n_Child).���� Is Not Null Or t_Middle(n_Child).��� = 1 Then
      t_Return.Extend;
      t_Return(t_Return.Count) := t_Middle(n_Child);
      If t_Middle(n_Child).�ϼ�id Is Not Null Then
        For n_Parent In n_Child + 1 .. t_Middle.Count Loop
          If t_Middle(n_Parent).��� = 0 And t_Middle(n_Parent).Id = t_Middle(n_Child).�ϼ�id Then
            t_Middle(n_Parent).��� := 1;
            Exit;
          End If;
        End Loop;
      End If;
    End If;
  End Loop;

  Return t_Return;
End f_Reg_Menu;
/

--00000:���,2014-03-27,��Ʒ�ض���Ȩ���ƸĽ���δ�Ǽ�����
Create Or Replace Package Body zlTools.b_Popedom Is
  --���ܣ�CopyMenu
  Procedure Copy_Menu
  (
    ϵͳ_In   In Zlmenus.ϵͳ%Type,
    ��ϵͳ_In In Zlmenus.ϵͳ%Type
  ) Is
    n_Menuid Zlmenus.Id%Type;
  Begin
    Select Max(ID) Into n_Menuid From zlMenus;
    n_Menuid := Nvl(n_Menuid, 0) + 1;
    Insert Into zlMenus
      (���, ID, �ϼ�id, ����, �̱���, ���, ͼ��, ˵��, ģ��, ϵͳ)
      Select ���, ID + n_Menuid, �ϼ�id + n_Menuid, ����, �̱���, ���, ͼ��, ˵��, ģ��, ��ϵͳ_In
      From zlMenus
      Where ϵͳ = ϵͳ_In;
    Commit;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Copy_Menu;

  --���ܣ�ȡZlMenu����
  Procedure Get_Menu_Tree
  (
    Cursor_Out Out t_Refcur,
    ���_In    In Zlmenus.���%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select ID, �ϼ�id, ����, ���, ˵��, ϵͳ, ģ��, �̱���, ͼ��, Level As ����
      From zlMenus
      Start With �ϼ�id Is Null And ��� = ���_In
      Connect By Prior ID = �ϼ�id And ��� = ���_In
      Order By Level, ID;
  End Get_Menu_Tree;

  --���ܣ�ȡZlMenu����
  Procedure Get_Menu_Group
  (
    Cursor_Out Out t_Refcur,
    ���_In    In Zlmenus.���%Type
  ) Is
  Begin
    If ���_In Is Null Then
      --frmMenu.FillMenuName
      Open Cursor_Out For
        Select Distinct ��� From zlMenus;
    Else
      --frmMenu.cmdNew_Click
      Open Cursor_Out For
        Select Count(*) As ���� From zlMenus Where ��� = ���_In;
    End If;
  End Get_Menu_Group;

  --���ܣ�ȡģ��
  Procedure Get_Module
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In Zlcomponent.ϵͳ%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select p.���, p.����, c.���� As ����
      From zlPrograms P, zlComponent C
      Where Upper(p.����) = Upper(c.����) And c.ϵͳ = ϵͳ_In And p.ϵͳ = ϵͳ_In
      Order By c.����, p.���;
  End Get_Module;

  --���ܣ�ȡ���ܻ����У�˵��
  Procedure Get_Function
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In Zlprogfuncs.ϵͳ%Type,
    ���_In    In Zlprogfuncs.���%Type,
    ����_In    In Zlprogfuncs.����%Type := Null
  ) Is
  Begin
    If Nvl(����_In, '��') = '��' Then
      Open Cursor_Out For
        Select ���� From zlProgFuncs Where ϵͳ = ϵͳ_In And ��� = ���_In Order By Nvl(����, 0);
    Else
      Open Cursor_Out For
        Select ����, ˵�� From zlProgFuncs Where ϵͳ = ϵͳ_In And ��� = ���_In And ���� = ����_In;
    End If;
  End Get_Function;

  --���ܣ�ȡ��Ȩ��
  Procedure Get_Impower
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In Zlprogprivs.ϵͳ%Type,
    ���_In    In Zlprogprivs.���%Type,
    ����_In    In Zlprogprivs.����%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select ����, Sum(Decode(Ȩ��, 'SELECT', 1, 0)) As "SELECT", Sum(Decode(Ȩ��, 'UPDATE', 1, 0)) As "UPDATE",
             Sum(Decode(Ȩ��, 'INSERT', 1, 0)) As "INSERT", Sum(Decode(Ȩ��, 'DELETE', 1, 0)) As "DELETE",
             Sum(Decode(Ȩ��, 'EXECUTE', 1, 0)) As "EXECUTE"
      From zlProgPrivs
      Where ϵͳ = ϵͳ_In And ��� = ���_In And ���� = ����_In
      Group By ����
      Order By ����;
  End Get_Impower;

  --���ܣ��õ���ɫ�ܷ��ʵĵ���̨����
  Procedure Get_Role_Tools
  (
    Cursor_Out Out t_Refcur,
    ��ɫ_In    In Zlrolegrant.��ɫ%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select p.���, p.����, p.˵��, r.����
      From zlRoleGrant R, zlPrograms P
      Where r.ϵͳ Is Null And p.��� = r.��� And r.��ɫ = ��ɫ_In And p.ϵͳ Is Null And p.��� < 100 And p.���� Is Null
      Order By p.���;
  End Get_Role_Tools;

  --���ܣ��õ���ǰ��Ȩ��
  Procedure Get_Role_Grant
  (
    Curgrand_Out    Out t_Refcur,
    Curprivs_Out    Out t_Refcur,
    Curfuncpars_Out Out t_Refcur,
    ��ɫ_In         In Zlrolegrant.��ɫ%Type
  ) Is
  Begin
    Open Curgrand_Out For
      Select Nvl(ϵͳ, 0) As ϵͳ, ���, ���� From zlRoleGrant Where ��ɫ = ��ɫ_In;
    Open Curprivs_Out For
      Select Nvl(ϵͳ, 0) As ϵͳ, ���, ����, ������, Ȩ��, ���� From zlProgPrivs;
    Open Curfuncpars_Out For
      Select p.ϵͳ, f.������, p.����
      From zlFuncPars P, zlFunctions F
      Where p.ϵͳ = f.ϵͳ And p.������ = f.������ And p.���� Is Not Null;
  End Get_Role_Grant;

  --���ܣ�FillFunc
  Procedure Get_Zlprogfunc
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In Zlprogfuncs.ϵͳ%Type,
    ���_In    In Zlprogfuncs.���%Type
  ) Is
  Begin
    If Nvl(ϵͳ_In, 0) = 0 Then
      Open Cursor_Out For
        Select ����, ����, ˵�� From zlProgFuncs Where ϵͳ Is Null And ��� = ���_In And ���� <> '����';
    Else
      Open Cursor_Out For
        Select a.����, a.����, a.˵��
        From zlProgFuncs A, zlRegFunc B
        Where (a.ϵͳ / 100) = b.ϵͳ(+) And a.��� = b.���(+) And a.���� = b.����(+) And
              (b.���� Is Not Null Or b.���� Is Null And (a.��� Between 10000 And 19999)) And a.ϵͳ = ϵͳ_In And a.��� = ���_In And
              a.���� <> '����';
    End If;
  End Get_Zlprogfunc;

  --���ܣ������н�ɫ��Ӧ��ģ��
  Procedure Get_All_Module(Cursor_Out Out t_Refcur) Is
  Begin
    Open Cursor_Out For
      Select a.��ɫ, a.���, a.����, b.����, b.˵��
      From zlRoleGrant A, zlPrograms B
      Where a.��� = b.��� And Nvl(a.ϵͳ, 0) = Nvl(b.ϵͳ, 0)
      Order By a.��ɫ, a.���;
  End Get_All_Module;

End b_Popedom;
/

--00000:���,2014-03-27,��Ʒ�ض���Ȩ���ƸĽ���δ�Ǽ�����
Create Or Replace Package Body zlTools.b_Runmana Is
  --���ܣ�ȡ������Ϣ
  --frmParameters
  Procedure Get_Parameters
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In Number
  ) Is
  Begin
    If Nvl(ϵͳ_In, 0) = 0 Then
      Open Cursor_Out For
        Select a.Id, a.ϵͳ, a.ģ��, a.˽��, a.������, a.������, a.����ֵ, a.ȱʡֵ, a.����˵��, a.����, a.��Ȩ, a.�̶�, b.���� As ģ������,
               zlSpellCode(b.����) As ģ�����
        From zlParameters A, zlPrograms B
        Where Nvl(a.ϵͳ, 0) = 0 And Nvl(a.ϵͳ, 0) = b.ϵͳ(+) And Nvl(a.ģ��, 0) = b.���(+);
    Else
      Open Cursor_Out For
        Select a.Id, a.ϵͳ, a.ģ��, a.˽��, a.������, a.������, a.����ֵ, a.ȱʡֵ, a.����˵��, a.����, a.��Ȩ, a.�̶�, b.���� As ģ������,
               zlSpellCode(b.����) As ģ�����
        From zlParameters A, zlPrograms B,
             --����Ȩ�޲��֣�ֻ����Ȩ�Ĳ�����ʾ
             (Select Distinct f.���
               From zlProgFuncs F, zlRegFunc R
               Where Trunc(f.ϵͳ / 100) = r.ϵͳ(+) And f.��� = r.���(+) And f.���� = r.����(+) And
                     (r.���� Is Not Null Or r.���� Is Null And (f.��� Between 10000 And 19999)) And f.ϵͳ = ϵͳ_In And
                     1 = (Select 1 From Zlregaudit A Where a.��Ŀ = '��Ȩ֤��')
               Union All
               Select 0 As ���
               From Dual) M
        Where a.ϵͳ = Nvl(ϵͳ_In, 0) And Nvl(a.ϵͳ, 0) = b.ϵͳ(+) And Nvl(a.ģ��, 0) = b.���(+) And Nvl(a.ģ��, 0) = m.���;
    End If;
  End Get_Parameters;

  --���ܣ�����ָ���Ĳ���IDȡ������Ϣ
  --�����б�frmParameters;frmParaChangeSet
  Procedure Get_Parameter
  (
    Cursor_Out Out t_Refcur,
    ����id_In  In Zlparameters.Id%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select a.Id, a.ϵͳ, a.ģ��, a.˽��, a.������, a.������, a.����ֵ, a.ȱʡֵ, a.����˵��, a.����, a.��Ȩ, a.�̶�, b.���� As ģ������,
             zlSpellCode(b.����) As ģ�����
      From zlParameters A, zlPrograms B
      Where a.Id = Nvl(����id_In, 0) And Nvl(a.ϵͳ, 0) = b.ϵͳ(+) And Nvl(a.ģ��, 0) = b.���(+);
  End Get_Parameter;

  --���ܣ�ȡվ����û��Ĳ�����Ϣ
  --�����б�frmParameters
  Procedure Get_Userparameters
  (
    Cursor_Out Out t_Refcur,
    ����id_In  In Zluserparas.����id%Type,
    Inttype    In Number := 0
    --0-���в�����Ϣ,1-ֻ��ȡ������������,2-ֻ��ȡ�û���
  ) Is
    n_˽�� Zlparameters.˽��%Type;
    n_���� Zlparameters.����%Type;
  Begin
    If Inttype = 0 Then
      Begin
        Select Nvl(a.˽��, 0), Nvl(a.����, 0) Into n_˽��, n_���� From zlParameters A Where ID = Nvl(����id_In, 0);
      Exception
        When Others Then
          n_˽�� := 0;
          n_���� := 0;
      End;
      If n_���� = 1 Then
        --�ֻ���
        If n_˽�� = 1 Then
          --����˽��ģ��
          Open Cursor_Out For
            Select ����id, �û���, ����ֵ, ������, zlSpellCode(������) As ����������
            From zlUserParas
            Where ����id = Nvl(����id_In, 0) And �û��� Is Not Null And ������ Is Not Null;
        Else
          --��������ģ��
          Open Cursor_Out For
            Select ����id, �û���, ����ֵ, ������, zlSpellCode(������) As ����������
            From zlUserParas
            Where ����id = Nvl(����id_In, 0) And �û��� Is Null And ������ Is Not Null;
        End If;
      Else
        If n_˽�� = 1 Then
          --˽��ģ���˽��ȫ��
          Open Cursor_Out For
            Select ����id, �û���, ����ֵ, ������, zlSpellCode(������) As ����������
            From zlUserParas
            Where ����id = Nvl(����id_In, 0) And �û��� Is Not Null And ������ Is Null;
        Else
          --����ģ��͹���ȫ��,��������ص�����
          Open Cursor_Out For
            Select ����id, �û���, ����ֵ, ������, '' As ����������
            From zlUserParas
            Where ����id = Nvl(����id_In, 0) And 1 = 2;
        End If;
      End If;
    Elsif Inttype = 1 Then
      --ֻ��ȡ������������,
      Open Cursor_Out For
        Select Distinct ������, zlSpellCode(������) As ����������
        From zlUserParas
        Where ����id = Nvl(����id_In, 0) And ������ Is Not Null;
    Elsif Inttype = 2 Then
      --ֻ��ȡ�û���
      Open Cursor_Out For
        Select Distinct �û��� From zlUserParas Where ����id = Nvl(����id_In, 0) And �û��� Is Not Null;
    Else
      Open Cursor_Out For
        Select ����id, �û���, ����ֵ, ������, zlSpellCode(������) As ����������
        From zlUserParas
        Where ����id = Nvl(����id_In, 0);
    End If;
  End Get_Userparameters;

  --���ܣ�ȡ�����޸���Ϣ
  --�����б�frmParameters
  Procedure Get_Parachangedlog
  (
    Cursor_Out Out t_Refcur,
    ����id_In  In Zlparachangedlog.����id%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select ����id, ���, �䶯˵��, �䶯����, �䶯��, �䶯ʱ��, �䶯ԭ��
      From Zlparachangedlog
      Where ����id = Nvl(����id_In, 0);
  
  End;
  --���ܣ�ȡZlAutoJob���к�
  Procedure Get_Job_Number
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In Number
  ) Is
  Begin
    Open Cursor_Out For
      Select ��� + 1 As ���
      From zlAutoJobs
      Where Nvl(ϵͳ, 0) = ϵͳ_In And ���� = 3 And
            ��� + 1 Not In (Select ��� From zlAutoJobs Where Nvl(ϵͳ, 0) = ϵͳ_In And ���� = 3);
  End Get_Job_Number;

  --���ܣ�ȡZlDataMove����
  Procedure Get_Depict
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In Zldatamove.ϵͳ%Type,
    ���_In    In Zldatamove.���%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select ת������ From zlDataMove Where Nvl(ϵͳ, 0) = ϵͳ_In And ��� = ���_In;
  End Get_Depict;

  --���ܣ�ȡzlClients��MAX IP
  Procedure Get_Client_Maxip(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select Max(Ip) As Ip From zlClients;
  End Get_Client_Maxip;

  --���ܣ�ȡzlClients�ļ�¼
  Procedure Get_Client
  (
    Cur_Out   Out t_Refcur,
    ����վ_In In Zlclients.����վ%Type := Null
  ) Is
    v_Sql Varchar2(1000);
  Begin
    If Nvl(����վ_In, '��') = '��' Then
      v_Sql := 'Select a.Ip, a.����վ, a.Cpu, a.�ڴ�, a.Ӳ��, a.����ϵͳ, a.����, a.��;, a.˵��, a.������־, a.��ֹʹ��,
                             a.������, Decode(b.Terminal, Null, 0, 1) As ״̬, a.�ռ���־,a.����������,a.վ��,a.������ƵԴ
                From Zlclients a, (Select Distinct Terminal From V$session) b
                Where Upper(a.����վ) = Upper(b.Terminal(+))
                Order By a.Ip';
      Open Cur_Out For v_Sql;
    Else
      Open Cur_Out For
        Select Ip, ����վ, Cpu, �ڴ�, Ӳ��, ����ϵͳ, ����, ��;, ˵��, ������־, �ռ���־, ��ֹʹ��, ������, ����������, վ��, ������ƵԴ
        From zlClients
        Where Upper(����վ) = ����վ_In;
    End If;
  End Get_Client;

  --���ܣ�ȡzlClients��վ��
  Procedure Get_Client_Station(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select Distinct Upper(����վ) || '[' || Ip || ']' As վ��, Upper(����վ) ����վ From zlClients;
  End Get_Client_Station;

  --���ܣ�ȡ������
  Procedure Get_Project_No(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select ������ From Zlclientparaset Where Rownum = 1;
  End Get_Project_No;

  --���ܣ�ȡ����
  Procedure Get_Client_Scheme(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select ������, ������ || '-' || �������� As ��������, ��������, ����վ, �û��� From Zlclientscheme;
  End Get_Client_Scheme;

  --���ܣ�ȡ�ָ���Ϣ
  Procedure Get_Resile
  (
    Cur_Out   Out t_Refcur,
    ������_In In Zlclientparaset.������%Type,
    ����_In   In Number := 0
  ) Is
  Begin
    If ����_In = 0 Then
      Open Cur_Out For
        Select Distinct a.����վ || Decode(m.����վ, Null, ' ', '[' || m.Ip || ']') As ����վ, a.�û���, a.�ָ���־,
                        '[' || b.������ || ']' || b.�������� As ��������
        From Zlclientparaset A, Zlclientscheme B, zlClients M
        Where a.������ = b.������ And a.����վ = m.����վ(+) And a.������ = ������_In;
    End If;
  
    If ����_In = 1 Then
      Open Cur_Out For
        Select Distinct Upper(����վ) ����վ, Min(�ָ���־) �ָ���־
        From Zlclientparaset A
        Where a.������ = ������_In
        Group By ����վ;
    End If;
  
    If ����_In = 2 Then
      Open Cur_Out For
        Select Distinct Upper(�û���) �û���, Max(����վ) ����վ, Min(Decode(�ָ���־, 2, 0, �ָ���־)) �ָ���־
        From Zlclientparaset A
        Where a.������ = ������_In
        Group By �û���
        Order By �û���;
    End If;
  
  End Get_Resile;

  --���ܣ�ȡzldataMove����
  Procedure Get_Zldatamove
  (
    Cur_Out Out t_Refcur,
    ϵͳ_In In Zldatamove.ϵͳ%Type
  ) Is
  Begin
    Open Cur_Out For
      Select ���, ����, ˵��, �����ֶ�, ת������, �ϴ����� From zlDataMove Where ϵͳ = ϵͳ_In Order By ���;
  End Get_Zldatamove;

  --���ܣ�ȡ��־����
  Procedure Get_Log
  (
    Cur_Out     Out t_Refcur,
    ��־����_In In Varchar2,
    Where_In    In Varchar2
  ) Is
    v_Sql Varchar2(1000);
  Begin
    If ��־����_In = '������־' Then
      v_Sql := 'Select �Ự��,����վ,�û���,�������,������Ϣ,To_char(ʱ��,''yyyy-MM-dd hh24:mi:ss'') ʱ��
                     ,Decode(����,1,''�洢���̴���'',2,''������������'',3,''Ӧ�ó�������'',''�ͻ�����������'') ��������
                        From ZlErrorLog Where ' || Where_In;
      Open Cur_Out For v_Sql;
    End If;
    If ��־����_In = '������־' Then
      v_Sql := 'Select �Ự��,����վ,�û���,������,��������,To_char(����ʱ��,''yyyy-MM-dd hh24:mi:ss'') ����ʱ��
                                 ,To_char(�˳�ʱ��,''yyyy-MM-dd hh24:mi:ss'') �˳�ʱ��,Decode(�˳�ԭ��,1,''�����˳�'',''�쳣�˳�'') �˳�ԭ��
                                    From ZlDiaryLog Where ' || Where_In;
      Open Cur_Out For v_Sql;
    End If;
  End Get_Log;

  --���ܣ�ȡ��־��¼��
  Procedure Get_Log_Count
  (
    Cur_Out     Out t_Refcur,
    ��־����_In In Varchar2
  ) Is
  Begin
    If ��־����_In = '������־' Then
      Open Cur_Out For
        Select Count(*) ����
        From zlErrorLog
        Union All
        Select Nvl(To_Number(����ֵ), 0)
        From zlOptions
        Where ������ = 4;
    End If;
    If ��־����_In = '������־' Then
      Open Cur_Out For
        Select Count(*) ����
        From zlDiaryLog
        Union All
        Select Nvl(To_Number(����ֵ), 0)
        From zlOptions
        Where ������ = 2;
    
    End If;
  End Get_Log_Count;

  --���ܣ�ȡzlfilesupgradeg����
  Procedure Get_Zlfilesupgrade(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select ���, �ļ���, �汾��, �޸�����, �ļ�˵�� As ˵��,
             Decode(�ļ�����, 0, '��������', 1, 'Ӧ�ò���', 2, '�����ļ�', 3, '�����ļ�', 4, '��������', 5, 'ϵͳ�ļ�', '') As ����, ��װ·�� As ��װ·��,
             Md5 As Md5, ��������
      From zlFilesUpgrade
      Order By ���;
  End Get_Zlfilesupgrade;

  --���ܣ�ȡ��ע����Ŀ
  Procedure Get_Not_Regist(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select ��Ŀ, ����
      From zlRegInfo
      Where ��Ŀ Not In ('������', '�汾��', '������Ŀ¼', '�����û�', '��������', '�ռ�Ŀ¼', '�ռ�����', 'ע����', '��Ȩ֤��', '��Ȩ����', '��Ȩ�ʴ�');
  End Get_Not_Regist;

  --���ܣ�ȡ����ֵ
  Procedure Get_Zloption
  (
    Cur_Out   Out t_Refcur,
    ������_In In Zloptions.������%Type
  ) Is
  Begin
    Open Cur_Out For
      Select Nvl(����ֵ, ȱʡֵ) Option_Value From zlOptions Where ������ = ������_In;
  End Get_Zloption;

End b_Runmana;
/

--00000:���,2014-03-27,��Ʒ�ض���Ȩ���ƸĽ���δ�Ǽ�����
Create Or Replace Package Body zlTools.b_Comfunc Is
  --���ܣ����������־
  Procedure Save_Error_Log
  (
    ����_In     In Zlerrorlog.����%Type,
    �������_In In Zlerrorlog.�������%Type,
    ������Ϣ_In In Zlerrorlog.������Ϣ%Type
  ) Is
  Begin
    Insert Into zlErrorLog
      (�Ự��, �û���, ����վ, ʱ��, ����, �������, ������Ϣ)
      Select Sid, User, Machine, Sysdate, ����_In, �������_In, ������Ϣ_In
      From V$session
      Where Audsid = Userenv('SessionID');
    Commit;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Save_Error_Log;

  --���ܣ�ȡ���ù���
  Procedure Get_Usable_Function
  (
    Cursor_Out Out t_Refcur,
    ����_In    In Zlprograms.����%Type
  ) Is
  Begin
    If Nvl(����_In, '�տ�') = '�տ�' Then
      Open Cursor_Out For
        Select Distinct a.���, a.����, a.˵��
        From zlPrograms A, zlProgFuncs B, zlRegFunc C
        Where a.ϵͳ = b.ϵͳ And a.��� = b.��� And Trunc(b.ϵͳ / 100) = c.ϵͳ(+) And b.��� = c.���(+) And b.���� = c.����(+) And
              (c.���� Is Not Null Or c.���� Is Null And (a.��� Between 10000 And 19999))
        Order By a.���;
    Else
      Open Cursor_Out For
        Select Distinct a.���, a.����, a.˵��
        From zlPrograms A, zlProgFuncs B, zlRegFunc C
        Where a.ϵͳ = b.ϵͳ And a.��� = b.��� And Upper(a.����) = Upper(����_In) And Trunc(b.ϵͳ / 100) = c.ϵͳ(+) And
              b.��� = c.���(+) And b.���� = c.����(+) And
              (c.���� Is Not Null Or c.���� Is Null And (a.��� Between 10000 And 19999))
        Order By a.���;
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Usable_Function;

  --���ܣ�ȡ��д���    
  Procedure Get_Uppmoney
  (
    Cursor_Out Out t_Refcur,
    ���_In    In Number
  ) Is
  Begin
    Open Cursor_Out For
      Select zlUppMoney(Nvl(���_In, 0)) As Num From Dual;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Uppmoney;

  --���ܣ�����ָ�������ڡ���š�ϵͳ�ж�ָ�����ڵ������Ƿ���ת���������ݱ���    
  Procedure Get_Datamoved
  (
    Cursor_Out  Out t_Refcur,
    ���_In     In Zldatamove.���%Type,
    ϵͳ_In     In Zldatamove.ϵͳ%Type,
    �ϴ�����_In In Zldatamove.�ϴ�����%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select ϵͳ, ���
      From zlDataMove
      Where ��� = ���_In And ϵͳ = ϵͳ_In And �ϴ����� > �ϴ�����_In And �ϴ����� Is Not Null;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Datamoved;

  --���ܣ�ȡϵͳ������
  Procedure Get_Owner
  (
    Cursor_Out Out t_Refcur,
    ���_In    In Zlsystems.���%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select ������ From zlSystems Where ��� = ���_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Owner;

  --���ܣ�ȡ����
  Procedure Get_Spell_Code
  (
    Cursor_Out Out t_Refcur,
    �ַ���_In  In Varchar2,
    ��ʽ_In    In Number := 0
  ) Is
  Begin
    If Nvl(��ʽ_In, 0) = 0 Then
      Open Cursor_Out For
        Select zlSpellCode(�ַ���_In) As ���� From Dual;
    Else
      Open Cursor_Out For
        Select zlWbCode(�ַ���_In) As ���� From Dual;
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Spell_Code;

  --���ܣ�����������־
  Procedure Save_Diary_Log
  (
    ������_In   In Zldiarylog.������%Type,
    ������_In   In Zldiarylog.������%Type,
    ��������_In In Zldiarylog.��������%Type
  ) Is
  Begin
    Insert Into zlDiaryLog
      (�Ự��, �û���, ����վ, ������, ������, ��������, ����ʱ��)
      Select Sid + Serial#, User, RTrim(LTrim(Replace(Machine, Chr(0), ''))), ������_In, ������_In, ��������_In, Sysdate
      From V$session
      Where Audsid = Userenv('SessionID');
    Commit;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Save_Diary_Log;

  --���ܣ�����������־
  --�����б�clsComLib.SaveWinState
  Procedure Update_Diary_Log
  (
    ������_In In Zldiarylog.������%Type,
    ������_In In Zldiarylog.������%Type
  ) Is
    Cursor c_Session Is
      Select Sid + Serial# As �Ự��, User As �û���, RTrim(LTrim(Replace(Machine, Chr(0), ''))) As ����վ
      From V$session
      Where Audsid = Userenv('SessionID');
  Begin
    For r_Tmp In c_Session Loop
      Update zlDiaryLog
      Set �˳�ԭ�� = 1, �˳�ʱ�� = Sysdate
      Where �˳�ԭ�� Is Null And �û��� = r_Tmp.�û��� And ����վ = r_Tmp.����վ And �Ự�� = r_Tmp.�Ự�� And ������ = ������_In And ������ = ������_In;
    End Loop;
    Commit;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Update_Diary_Log;

  --���ܣ�ȡ�̶�����������û���������
  Procedure Get_Report_Menu
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In Zlprograms.ϵͳ%Type,
    ���_In    In Zlprograms.���%Type,
    ����_In    In Zlreports.����%Type,
    ���_In    In Zlreports.���%Type
  ) Is
  Begin
    If Nvl(���_In, '�տ�') <> '�տ�' Then
      Open Cursor_Out For
        Select ��־, ϵͳ, ���, ����
        From (Select 1 As ��־, a.ϵͳ, a.���, a.����
               From zlReports A, zlPrograms B
               Where a.ϵͳ = b.ϵͳ And a.����id = b.��� And Not Upper(a.���) Like '%BILL%' And
                     Upper(b.����) <> Upper('zl9Report') And b.ϵͳ = ϵͳ_In And b.��� = ���_In And
                     Instr(����_In, ';' || a.���� || ';') > 0
               Union All
               Select Decode(a.ϵͳ, Null, 2, 1) As ��־, a.ϵͳ, a.���, a.����
               From zlReports A, zlRPTPuts B, zlPrograms C
               Where a.Id = b.����id And b.ϵͳ = c.ϵͳ And b.����id = c.��� And (Not Upper(a.���) Like '%BILL%' Or a.ϵͳ Is Null) And
                     Instr(����_In, ';' || b.���� || ';') > 0 And c.ϵͳ = ϵͳ_In And c.��� = ���_In)
        Where Instr(���_In, ',' || ��� || ',') = 0
        Order By ��־, ���;
    Else
      Open Cursor_Out For
        Select ��־, ϵͳ, ���, ����
        From (Select 1 As ��־, a.ϵͳ, a.���, a.����
               From zlReports A, zlPrograms B
               Where a.ϵͳ = b.ϵͳ And a.����id = b.��� And Not Upper(a.���) Like '%BILL%' And
                     Upper(b.����) <> Upper('zl9Report') And b.ϵͳ = ϵͳ_In And b.��� = ���_In And
                     Instr(����_In, ';' || a.���� || ';') > 0
               Union All
               Select Decode(a.ϵͳ, Null, 2, 1) As ��־, a.ϵͳ, a.���, a.����
               From zlReports A, zlRPTPuts B, zlPrograms C
               Where a.Id = b.����id And b.ϵͳ = c.ϵͳ And b.����id = c.��� And (Not Upper(a.���) Like '%BILL%' Or a.ϵͳ Is Null) And
                     Instr(����_In, ';' || b.���� || ';') > 0 And c.ϵͳ = ϵͳ_In And c.��� = ���_In)
        Order By ��־, ���;
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Report_Menu;

  --���ܣ�ȡ�û�������Ϣ
  Procedure Get_Zlnoticerec
  (
    Cursor_Out Out t_Refcur,
    �û���_In  In Zlnoticerec.�û���%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select a.���, a.ϵͳ, c.����id As ģ��, c.ϵͳ As ����ϵͳ, b.�������� As �������, c.���� As ���ѱ���, a.��������, b.���ʱ��, b.�Ѷ���־
      From zlNotices A, zlNoticeRec B, (Select * From zlReports Where ����ʱ�� Is Not Null) C
      Where b.�û��� = �û���_In And b.���ѱ�־ > 0 And c.���(+) = a.���ѱ��� And a.��� = b.������� And b.�������� Is Not Null;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Zlnoticerec;

  --���ܣ�ȡ�ʼ�����
  Procedure Get_Zlmessage
  (
    Cursor_Out Out t_Refcur,
    Id_In      In Zlmessages.Id%Type,
    ����_In    In Zlmsgstate.����%Type,
    �û�_In    In Zlmsgstate.�û�%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select a.*, b.ɾ��, b.״̬
      From zlMessages A, zlMsgState B
      Where a.Id = b.��Ϣid And b.��Ϣid = Id_In And b.���� = ����_In And b.�û� = �û�_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Zlmessage;

  --���ܣ�ȡ�ʼ�����
  Procedure Get_Zlmessage
  (
    Cursor_Out Out t_Refcur,
    Id_In      In Zlmessages.Id%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select ����, ����ɫ From zlMessages Where ID = Id_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Zlmessage;

  --���ܣ�ȡ�ʵݵ�ַ
  Procedure Get_Zlmsgstate
  (
    Cursor_Out Out t_Refcur,
    ��Ϣid_In  In Zlmsgstate.��Ϣid%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select ����, �û�, ��� From zlMsgState Where ��Ϣid = ��Ϣid_In;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Zlmsgstate;

  --���ܣ�ɾ����Ϣ
  Procedure Delete_Zlmsgstate
  (
    ɾ��_In   In Zlmsgstate.ɾ��%Type,
    ��Ϣid_In In Zlmsgstate.��Ϣid%Type,
    ����_In   In Zlmsgstate.����%Type,
    �û�_In   In Zlmsgstate.�û�%Type
  ) Is
    n_���� Number(10);
    n_���� Number(10);
  Begin
    If Nvl(ɾ��_In, 0) = 1 Then
      Update zlMsgState Set ɾ�� = 1 Where ��Ϣid = ��Ϣid_In And ���� = ����_In And �û� = �û�_In;
    Else
      If ����_In = 0 Then
        --���ڲݸ壬���ռ��˵�Ҳһ��ɾ��
        Update zlMsgState Set ɾ�� = 2 Where ��Ϣid = ��Ϣid_In And �û� = �û�_In;
      Else
        Update zlMsgState Set ɾ�� = 2 Where ��Ϣid = ��Ϣid_In And ���� = ����_In And �û� = �û�_In;
      End If;
      -- ɾ��ָ��ID����Ϣ  mnuEditDelete_Click ����
      Select Count(*) As ����, Sum(Decode(ɾ��, 2, 1, 0)) As ����
      Into n_����, n_����
      From zlMsgState
      Where ��Ϣid = ��Ϣid_In;
    
      If n_���� = n_���� Then
        Delete From zlMessages Where ID = ��Ϣid_In;
      End If;
    End If;
    Commit;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Delete_Zlmsgstate;

  --���ܣ�ɾ��������Ϣ
  Procedure Delete_Zlmessage Is
    n_Days Number;
  Begin
    Select Nvl(����ֵ, ȱʡֵ) Into n_Days From zlOptions Where ������ = 5;
    If Nvl(n_Days, 0) > 0 Then
      Delete From zlMessages Where ʱ�� < Sysdate - n_Days;
      Commit;
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Delete_Zlmessage;

  --���ܣ�ȡ�ʼ��б�
  Procedure Get_Mail_List
  (
    Cursor_Out  Out t_Refcur,
    ��Ϣ����_In In Varchar2,
    �û�_In     In Zlmsgstate.�û�%Type,
    ��ʾ�Ѷ�_In In Number,
    �Ựid_In   In Zlmessages.�Ựid%Type
  ) Is
    v_Sql  Varchar2(1000);
    v_�Ѷ� Varchar2(100);
    v_���� Varchar2(100);
  Begin
  
    If Nvl(��ʾ�Ѷ�_In, 0) = 1 Then
      v_�Ѷ� := ' and substr(S.״̬,1,1)=''0''';
    Else
      v_�Ѷ� := '';
    End If;
  
    If Instr(';�ݸ�;�ռ���;�ѷ�����Ϣ;��ɾ����Ϣ;�����Ϣ;', ';' || ��Ϣ����_In || ';') <= 0 Then
      v_���� := '�ݸ�';
    Else
      v_���� := ��Ϣ����_In;
    End If;
  
    If v_���� = '�ݸ�' Then
      v_Sql := 'Select M.ID, M.�Ựid, M.������, M.�ռ���, M.����, To_Char(M.ʱ��, ''YYYY-MM-DD HH24:MI:SS'') As ʱ��, S.����, S.״̬ 
              From zlMessages M, zlMsgState S
              Where M.ID = S.��Ϣid  and S.ɾ��=0 and S.�û�= ''' || �û�_In || ''' And S.����=0 ' || v_�Ѷ�;
    End If;
  
    If v_���� = '�ռ���' Then
      v_Sql := 'Select M.ID, M.�Ựid, M.������, M.�ռ���, M.����, To_Char(M.ʱ��, ''YYYY-MM-DD HH24:MI:SS'') As ʱ��, S.����, S.״̬ 
              From zlMessages M, zlMsgState S
              Where M.ID = S.��Ϣid  and S.ɾ��=0 and S.�û�= ''' || �û�_In || ''' And S.����=2 ' || v_�Ѷ�;
    End If;
  
    If v_���� = '�ѷ�����Ϣ' Then
      v_Sql := 'Select M.ID, M.�Ựid, M.������, M.�ռ���, M.����, To_Char(M.ʱ��, ''YYYY-MM-DD HH24:MI:SS'') As ʱ��, S.����, S.״̬ 
              From zlMessages M, zlMsgState S
              Where M.ID = S.��Ϣid  and S.ɾ��=0 and S.�û�= ''' || �û�_In || ''' And S.����=1 ' || v_�Ѷ�;
    End If;
  
    If v_���� = '��ɾ����Ϣ' Then
      v_Sql := 'Select M.ID, M.�Ựid, M.������, M.�ռ���, M.����, To_Char(M.ʱ��, ''YYYY-MM-DD HH24:MI:SS'') As ʱ��, S.����, S.״̬ 
              From zlMessages M, zlMsgState S
              Where M.ID = S.��Ϣid  and S.�û�= ''' || �û�_In || ''' And S.ɾ��=1 ' || v_�Ѷ�;
    End If;
  
    If v_���� = '�����Ϣ' Then
      v_Sql := 'select M.ID,M.�ỰID,M.������,M.�ռ���,M.����,to_char(M.ʱ��,''YYYY-MM-DD HH24:MI:SS'') as ʱ��,S.����,S.״̬
         from zlMessages M,zlMsgState S where M.ID=S.��ϢID and S.ɾ��<>2 and S.�û�= ''' || �û�_In ||
               '''  and M.�ỰID=' || �Ựid_In;
    End If;
  
    If Nvl(v_Sql, '�տ�') <> '�տ�' Then
      Open Cursor_Out For v_Sql;
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Get_Mail_List;

  --���ܣ���ԭɾ������Ϣ
  Procedure Restore_Zlmsgstate
  (
    ��Ϣid_In In Zlmsgstate.��Ϣid%Type,
    ����_In   In Zlmsgstate.����%Type,
    �û�_In   In Zlmsgstate.�û�%Type
  ) Is
  Begin
    Update zlMsgState Set ɾ�� = 0 Where ��Ϣid = ��Ϣid_In And ���� = ����_In And �û� = �û�_In;
    Commit;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Restore_Zlmsgstate;

  --���ܣ�������Ϣ
  --�����б�zlApptools.frmMessageEdit.SaveMessage
  Procedure Save_Zlmessage
  (
    Cursor_Out Out t_Refcur,
    Id_In      In Zlmessages.Id%Type,
    �Ựid_In  In Zlmessages.�Ựid%Type,
    ������_In  In Zlmessages.������%Type,
    �ռ���_In  In Zlmessages.�ռ���%Type,
    ����_In    In Zlmessages.����%Type,
    ����_In    In Zlmessages.����%Type,
    ����ɫ_In  In Zlmessages.����ɫ%Type
  ) Is
    n_Id     Zlmessages.Id%Type;
    n_�Ựid Zlmessages.�Ựid%Type;
  Begin
    If Nvl(Id_In, 0) = 0 Then
      Select Zlmessages_Id.Nextval Into n_Id From Dual;
      n_Id := Nvl(n_Id, 0);
      If Nvl(�Ựid_In, 0) = 0 Then
        n_�Ựid := n_Id;
      Else
        n_�Ựid := �Ựid_In;
      End If;
      Insert Into zlMessages
        (ID, �Ựid, ������, ʱ��, �ռ���, ����, ����, ����ɫ)
      Values
        (n_Id, n_�Ựid, ������_In, Sysdate, �ռ���_In, ����_In, ����_In, ����ɫ_In);
      Open Cursor_Out For
        Select n_Id As ID, n_�Ựid As �Ựid From Dual;
    Else
      Update zlMessages
      Set ������ = ������_In, ʱ�� = Sysdate, �ռ��� = �ռ���_In, ���� = ����_In, ���� = ����_In, ����ɫ = ����ɫ_In
      Where ID = Id_In;
      Open Cursor_Out For
        Select Id_In As ID, �Ựid_In As �Ựid From Dual;
    End If;
    Commit;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Save_Zlmessage;

  --���ܣ�����zlMsgstate
  --�����б�zlApptools.frmMessageEdit.SaveMessage
  Procedure Insert_Zlmsgstate
  (
    ��Ϣid_In In Zlmsgstate.��Ϣid%Type,
    ����_In   In Zlmsgstate.����%Type,
    �û�_In   In Zlmsgstate.�û�%Type,
    ���_In   In Zlmsgstate.���%Type,
    ɾ��_In   In Zlmsgstate.ɾ��%Type,
    ״̬_In   In Zlmsgstate.״̬%Type
  ) Is
  Begin
  
    If ����_In < 2 Then
      Delete From zlMsgState Where ��Ϣid = ��Ϣid_In;
    End If;
    Insert Into zlMsgState
      (��Ϣid, ����, �û�, ���, ɾ��, ״̬)
    Values
      (��Ϣid_In, ����_In, �û�_In, ���_In, ɾ��_In, ״̬_In);
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Insert_Zlmsgstate;

  --���ܣ�Ϊԭ�����ϴ𸴻�ת����־
  Procedure Update_Zlmsgstate_State
  (
    ģʽ_In   In Number,
    ��Ϣid_In In Zlmsgstate.��Ϣid%Type,
    ����_In   In Zlmsgstate.����%Type,
    �û�_In   In Zlmsgstate.�û�%Type
  ) Is
  Begin
    If Nvl(ģʽ_In, 0) = 1 Or Nvl(ģʽ_In, 0) = 2 Then
      Update zlMsgState
      Set ״̬ = Substr(״̬, 1, 1) || '1' || Substr(״̬, 3, 2)
      Where ��Ϣid = ��Ϣid_In And ���� = ����_In And �û� = �û�_In;
      Commit;
    End If;
    If Nvl(ģʽ_In, 0) = 3 Then
      Update zlMsgState
      Set ״̬ = Substr(״̬, 1, 1) || '1' || Substr(״̬, 4, 1)
      Where ��Ϣid = ��Ϣid_In And ���� = ����_In And �û� = �û�_In;
      Commit;
    End If;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Update_Zlmsgstate_State;

  --���ܣ�����״̬�����
  Procedure Update_Zlmsgstate_Idtntify
  (
    ���_In   In Zlmsgstate.���%Type,
    ��Ϣid_In In Zlmsgstate.��Ϣid%Type,
    ����_In   In Zlmsgstate.����%Type,
    �û�_In   In Zlmsgstate.�û�%Type
  ) Is
  Begin
    Update zlMsgState
    Set ״̬ = '1' || Substr(״̬, 2), ��� = ���_In
    Where ��Ϣid = ��Ϣid_In And ���� = ����_In And �û� = �û�_In;
    Commit;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Update_Zlmsgstate_Idtntify;

End b_Comfunc;
/