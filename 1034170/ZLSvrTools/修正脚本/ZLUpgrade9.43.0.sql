----10.32.0---��9.43.0

--59747:������,2013-03-26,��ʷ����ת�������Ż�
Alter Table zltools.zlDataMove add ״̬ number(1);
Alter Table zltools.zlDataMove add ͣ����ҵ�� Varchar2(4000);
Alter Table zltools.zlBakTables add ͣ�ô����� number(1);

create table zltools.zlDataMovelog
(ϵͳ number(5),
���� number(8),
���� number(8),
��ֹʱ�� date,
��ת�� number(1),
��ǰ���� varchar2(100),
��ǿ�ʼʱ�� date,
��ǽ���ʱ�� date,
ת����ʼʱ�� date,
ת������ʱ�� date,
�ؽ�����ʱ�� date)
/
ALTER TABLE zltools.zlDataMovelog ADD CONSTRAINT zlDataMovelog_PK PRIMARY KEY (ϵͳ,����) USING INDEX PCTFREE 5
/
ALTER TABLE zltools.zlDataMovelog ADD CONSTRAINT zlDataMovelog_FK_ϵͳ FOREIGN KEY (ϵͳ) REFERENCES zlSystems(���) ON DELETE CASCADE
/
create public synonym zlDataMovelog for zltools.zlDataMovelog;
grant select on zltools.zlDataMovelog to public;

--61786:������,2013-05-21,��ʷ����ת����־
Declare
  v_Sql Varchar2(100);
Begin
  For R In (Select Distinct ������ From zlSystems) Loop
    Begin
      v_Sql := 'grant select,insert,update,delete on zlDataMovelog to ' || r.������ || ' With GRANT Option';
      Execute Immediate v_Sql;
    Exception
      When Others Then
        Null;
        --�����߿��ܲ�����(ϵͳͣ��)
    End;
  End Loop;
End;
/


--61766:��˶,2013-05-20,������ϵͳ�������ܸ���
create table zltools.Zlbigtables(ϵͳ Number(5),���� Varchar2(30))
/
alter table ZLTOOLS.Zlbigtables add constraint Zlbigtables_PK primary key (ϵͳ, ����) USING INDEX PCTFREE 5
/

create public synonym Zlbigtables for zltools.Zlbigtables;
grant select on zltools.Zlbigtables to public;

Declare
  v_Sql Varchar2(100);
Begin
  For R In (Select Distinct ������ From zlSystems) Loop
    Begin
      v_Sql := 'grant select,insert,update,delete on Zlbigtables to ' || r.������ || ' With GRANT Option';
      Execute Immediate v_Sql;
    Exception
      When Others Then
        Null;
        --�����߿��ܲ�����(ϵͳͣ��)
    End;
  End Loop;
End;
/


--59244:л��,2013-03-08,������Ȩ���ȡȨ�����⡣
Create Or Replace Function f_Reg_Menu
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
               f.���� = r.����(+) And (r.���� Is Null And f.ϵͳ Is Null Or r.���� Is Not Null And r.���� = '����' Or
                                   r.���� Is Not Null And x.����id Is Not Null) And p.ϵͳ = x.ϵͳ(+) And p.��� = x.����id(+) And
               Upper(p.����) = c.Text And Nvl(p.ϵͳ, 0) = s.Prog And p.��� = p.��� * a.����(+) And Nvl(p.ϵͳ, 1) = o.���(+) And
               Nvl(p.ϵͳ, 0) = Nvl(g.ϵͳ(+), 0) And p.��� = g.���(+) And
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

--57403:���Ʊ�,2012-12-27
Create Table zlTools.zlMgrGrant(
    �û��� varchar2(30),
    ���� varchar2(500))
    PCTFREE 5 PCTUSED 90
    Cache Storage(Buffer_Pool Keep)
/
Create Public Synonym zlMgrGrant For zlTools.zlMgrGrant
/
Grant Select On zlTools.zlMgrGrant To Public
/

ALTER TABLE zlTools.zlMgrGrant ADD CONSTRAINT zlMgrGrant_PK PRIMARY KEY (�û���) USING INDEX PCTFREE 5
/
Insert Into zlTools.zlSvrTools(���,�ϼ�,����,���,˵��) Values('0404','04','��������Ȩ','N',Null)
/

--57426:����,2012-12-28
Insert Into zlSvrTools(���,�ϼ�,����,���,˵��) Values('0505','05','�Զ�����̹���','Z',Null)
/
Create Sequence zlTools.zlProcedure_ID start with 1
/
Create Public Synonym zlProcedure_ID For zlTools.zlProcedure_ID
/
Grant Select On zlTools.zlProcedure_ID to Public
/

CREATE TABLE zlTools.zlProcedure(
    ID NUMBER(5),
    ���� NUMBER(5),
    ���� VARCHAR2(50),
    ״̬ NUMBER(5),
    ˵�� VARCHAR2(200),
    ������ VARCHAR2(20),
    �޸���Ա Varchar2(20),
    �޸�ʱ�� DATE,
    �ϴ��޸���Ա VARCHAR2(20),
    �ϴ��޸�ʱ�� Date)
    PCTFREE 5 PCTUSED 90
    Cache Storage(Buffer_Pool Keep)
/
Create Public Synonym zlProcedure For zlTools.zlProcedure
/
Grant Select On zlTools.zlProcedure to Public
/

ALTER TABLE zlTools.zlProcedure ADD CONSTRAINT zlProcedure_PK PRIMARY KEY (ID) USING INDEX PCTFREE 5
/

CREATE TABLE zlTools.zlProcedureNote(
    ����id NUMBER(5),
    ��ʶ VARCHAR2(50),
    ˵�� VARCHAR2(4000))
    PCTFREE 5 PCTUSED 90
    Cache Storage(Buffer_Pool Keep)
/
Create Public Synonym zlProcedureNote For zlTools.zlProcedureNote
/
Grant Select On zlTools.zlProcedureNote to Public
/

ALTER TABLE zlTools.zlProcedureNote ADD CONSTRAINT zlProcedureNote_FK_����id FOREIGN KEY (����id) REFERENCES zlProcedure(ID)
/

CREATE TABLE zlTools.zlProcedureText(
    ����id NUMBER(5),
    ���� VARCHAR2(50),
    ��� NUMBER(5),
    ���� VARCHAR2(4000))
    PCTFREE 5 PCTUSED 90
    Cache Storage(Buffer_Pool Keep)
/
Create Public Synonym zlProcedureText For zlTools.zlProcedureText
/
Grant Select On zlTools.zlProcedureText to Public
/

ALTER TABLE zlTools.zlProcedureText ADD CONSTRAINT zlProcedureText_FK_����id FOREIGN KEY (����id) REFERENCES zlProcedure(ID)
/
Insert Into zlOptions(������,������,����ֵ,ȱʡֵ,����˵��) Values(8, '��������', '','', '�Ѽ�����ʱ���ӵ����ݿ�����')
/
Create Or Replace Procedure zlTools.Zl_Zlprocedure_Update
(
  Id_In           In Zlprocedure.Id%Type,
  ����_In         In Zlprocedure.����%Type,
  ����_In         In Zlprocedure.����%Type,
  ״̬_In         In Zlprocedure.״̬%Type,
  ˵��_In         In Zlprocedure.˵��%Type := Null,
  ������_In       In ZLprocedure.������%Type :=Null,
  �޸���Ա_In     In Zlprocedure.�޸���Ա%Type := Null,
  �޸�ʱ��_In     In Zlprocedure.�޸�ʱ��%Type := Null,
  �ϴ��޸���Ա_In In Zlprocedure.�ϴ��޸���Ա%Type := Null,
  �ϴ��޸�ʱ��_In In Zlprocedure.�ϴ��޸�ʱ��%Type := Null
) Is
Begin
  Update Zlprocedure
  Set ID = Id_In, ���� = ����_In, ���� = ����_In, ״̬ = ״̬_In, ˵�� = ˵��_In, ������ = ������_In, �޸���Ա = �޸���Ա_In, �޸�ʱ�� = �޸�ʱ��_In, �ϴ��޸���Ա = �ϴ��޸���Ա_In,
      �ϴ��޸�ʱ�� = �ϴ��޸�ʱ��_In
  Where ID = Id_In;
  If Sql%RowCount = 0 Then
    Insert Into Zlprocedure
      (ID, ����, ����, ״̬, ˵��, ������, �޸���Ա, �޸�ʱ��, �ϴ��޸���Ա, �ϴ��޸�ʱ��)
    Values
      (Id_In, ����_In, ����_In, ״̬_In, ˵��_In, ������_In, �޸���Ա_In, �޸�ʱ��_In, �ϴ��޸���Ա_In, �ϴ��޸�ʱ��_In);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlprocedure_Update;
/
Create Public Synonym Zl_Zlprocedure_Update For zlTools.Zl_Zlprocedure_Update
/
Grant Execute On zlTools.Zl_Zlprocedure_Update to Public
/

Create Or Replace Procedure zlTools.Zl_Zlprocedureconnect_Update
(
  ������_In In Zloptions.������%Type,
  ����ֵ_In In Zloptions.����ֵ%Type
) Is
Begin
  Update zlOptions Set ����ֵ = ����ֵ_In Where ������ = ������_In;
  If Sql%RowCount = 0 Then
    Insert Into zlOptions
      (������, ������, ����ֵ, ����˵��)
    Values
      (8, ������_In, ����ֵ_In, '�Ѽ�����ʱ���ӵ����ݿ�����');
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlprocedureconnect_Update;
/
Create Public Synonym Zl_Zlprocedureconnect_Update For zlTools.Zl_Zlprocedureconnect_Update
/
Grant Execute On zlTools.Zl_Zlprocedureconnect_Update to Public
/

Create Or Replace Procedure zlTools.Zl_Zlprocedure_Delete
(
  Id_In           In Zlprocedure.Id%Type
) Is
Begin
  Delete zlProcedureNote Where ����ID=Id_In;
  Delete zlProcedureText Where ����ID=Id_In;
  Delete zlProcedure Where ID = Id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlprocedure_Delete;
/
Create Public Synonym Zl_Zlprocedure_Delete For zlTools.Zl_Zlprocedure_Delete
/
Grant Execute On zlTools.Zl_Zlprocedure_Delete to Public
/

Create Or Replace Procedure zlTools.Zl_zlProcedureNote_Update
(
  ����Id_In       In zlProcedureNote.����id%Type,
  ��ʶ_In         In zlProcedureNote.��ʶ%Type,
  ˵��_In         In zlProcedureNote.˵��%Type
) Is
Begin
  Insert Into Zlprocedurenote (����id, ��ʶ, ˵��) Values (����id_In, ��ʶ_In, ˵��_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_zlProcedureNote_Update;
/
Create Public Synonym Zl_zlProcedureNote_Update For zlTools.Zl_zlProcedureNote_Update
/
Grant Execute On zlTools.Zl_zlProcedureNote_Update to Public
/

Create Or Replace Procedure zlTools.Zl_zlProcedureNote_Delete
(
  ����Id_In       In zlProcedureNote.����id%Type
) Is
Begin
  Delete From zlProcedureNote Where ����Id = ����Id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_zlProcedureNote_Delete;
/
Create Public Synonym Zl_zlProcedureNote_Delete For zlTools.Zl_zlProcedureNote_Delete
/
Grant Execute On zlTools.Zl_zlProcedureNote_Delete to Public
/

CREATE OR REPLACE Procedure zlTools.Zl_Zlproceduretext_Update
(
  ����id_In In Zlproceduretext.����id%Type,
  ����_In   In Zlproceduretext.����%Type,
  ���_In   In Zlproceduretext.���%Type,
  ����_In   In Zlproceduretext.����%Type
) Is
Begin
  Update Zlproceduretext Set ���� = ����_In, ��� = ���_In, ���� = ����_In Where ����id = ����id_In And ���� = ����_In And ��� = ���_In;
  If Sql%RowCount = 0 Then
    Insert Into Zlproceduretext (����id, ����, ���, ����) Values (����id_In, ����_In, ���_In, ����_In);
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlproceduretext_Update;
/
Create Public Synonym Zl_Zlproceduretext_Update For zlTools.Zl_Zlproceduretext_Update
/
Grant Execute On zlTools.Zl_Zlproceduretext_Update to Public
/

CREATE OR REPLACE Procedure zlTools.Zl_Zlproceduretext_Delete
(
  ����id_In In Zlproceduretext.����id%Type,
  ����_In   In Zlproceduretext.����%Type
) Is
Begin
  Delete From ZLproceduretext Where ����id = ����id_In And ���� = ����_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlproceduretext_Delete;
/       
Create Public Synonym Zl_Zlproceduretext_Delete For zlTools.Zl_Zlproceduretext_Delete
/
Grant Execute On zlTools.Zl_Zlproceduretext_Delete to Public
/

CREATE OR REPLACE Procedure ZLTOOLS.Zl_Zlproceduretext_Move Is
Begin
  Delete From Zlproceduretext Where ���� In (1,2);

  Insert Into Zlproceduretext(����id,����,���,����)
  Select ����id,Decode(����,3,1,4,2),���,���� From Zlproceduretext Where ���� In (3,4);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlproceduretext_Move;
/
Create Public Synonym Zl_Zlproceduretext_Move For zlTools.Zl_Zlproceduretext_Move
/
Grant Execute On zlTools.Zl_Zlproceduretext_Move to Public
/
