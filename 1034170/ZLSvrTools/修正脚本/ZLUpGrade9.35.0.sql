-----------------------------------------------------------------
--Ϊ��ϲ�Ʒ�汾����9.34��Ϊ9.35
-----------------------------------------------------------------
--��Ӧ��������(���˺�),ͬʱ�������ظ��Ŀ��
UPDATE zlSvrTools SET ���='E' WHERE ���='0307'
/
UPDATE zlSvrTools SET ���='R' WHERE ���='0308'
/
UPDATE zlSvrTools SET ���='A' WHERE ���='0309'
/
Insert Into zlSvrTools(���,�ϼ�,����,���,˵��) Values('0310','03','ϵͳ��������','P',Null)
/

Create Table zlTools.zlParaChangedLog(
    ����ID NUMBER(18),
    ���   NUMBER(18),
    �䶯˵�� VARCHAR2(200),--˵���䶯���:����:˽��ģ���Ϊ����ģ�顣
    �䶯���� VARCHAR2(200),--˵���䶯�ֶεı仯���:����:˽��:1-->0,����:1-->0��
    �䶯�� VARCHAR2(20),
    �䶯ʱ�� Date,
    �䶯ԭ�� varchar2(200))
    PCTFREE 5 PCTUSED 90
    Cache Storage(Buffer_Pool Keep)
/
Alter Table zlTools.zlParaChangedLog Add Constraint zlParaChangedLog_UQ_����ID Unique(����ID,���) Using Index PCTFREE 5
/
Alter Table zlTools.zlParaChangedLog Add Constraint zlParaChangedLog_FK_����ID Foreign Key (����ID) References zlTools.zlParameters(ID) On Delete Cascade
/
Create Index zlTools.zlParaChangedLog_IX_�䶯�� On zlTools.zlParaChangedLog(�䶯��) PCTFREE 5
/ 

grant Select on zlTools.zlParaChangedLog to Public
/
create public SYNONYM  zlParaChangedLog   FOR zlTools.zlParaChangedLog
/


--********************************************************************************************************
--�����������Begin
--********************************************************************************************************
--�ȵ������ݣ��Ա�����Լ������ȷ��������ΪҩƷ����ģ��Ĳ����ţ�˽�к͹����Ƿֿ���ŵģ����Ϊͳһ��š�
--�Ƚ�˽�еĲ�����+100��ZLHIS�ű�������ȷ������������ģ�����û��ʹ�ò����š�
Update zlParameters
Set ������ = ������ + 100
Where ˽�� = 1 And
      (ϵͳ, ģ��, ������) In (Select ϵͳ, ģ��, ������ From zlParameters Group By ϵͳ, ģ��, ������ Having Count(*) > 1)
/

--zlParameters
Alter Table zlTools.zlParameters Add(���� Number(1),��Ȩ Number(1),�̶� Number(1))
/
Alter Table zlTools.zlParameters Add Constraint zlParameters_CK_���� Check (���� IN(0,1))
/
Alter Table zlTools.zlParameters Add Constraint zlParameters_CK_��Ȩ Check (��Ȩ IN(0,1))
/
Alter Table zlTools.zlParameters Add Constraint zlParameters_CK_�̶� Check (�̶� IN(0,1))
/
Alter Table zlTools.zlParameters Drop Constraint zlParameters_UQ_������
/
Alter Table zlTools.zlParameters Drop Constraint zlParameters_UQ_������
/
Alter Table zlTools.zlParameters Add Constraint zlParameters_UQ_������ Unique(������,ģ��,ϵͳ) Using Index PCTFREE 5
/
Alter Table zlTools.zlParameters Add Constraint zlParameters_UQ_������ Unique(������,ģ��,ϵͳ) Using Index PCTFREE 5
/

--zlUserParas
Alter Table zlTools.zlUserParas Add ������ Varchar2(50)
/
Alter Table zlTools.zlUserParas Drop Constraint zlUserParas_PK
/
Alter Table zlTools.zlUserParas Add Constraint zlUserParas_UQ_����ID Unique(����ID,�û���,������) Using Index PCTFREE 5
/
Create Index zlTools.zlUserParas_IX_������ On zlUserParas(������) PCTFREE 5
/


--20952
insert into zlreginfo(��Ŀ,����) values('վ������',Null);


Create Or Replace Function zlTools.f_Get_Node_Amt Return Number As
  v_Return Number;
Begin
  Begin
    Select ���� Into v_Return From zlRegInfo Where ��Ŀ = 'վ������';
    If To_Number(v_Return) < 0 Or To_Number(v_Return) > 9 Then
      v_Return := Null;
    End If;
  Exception
    When Others Then
      Null;
  End;
  Return(v_Return);
End f_Get_Node_Amt;
/

Create Public SYNONYM f_Get_Node_Amt FOR zlTools.f_Get_Node_Amt
/
Grant Execute on zlTools.f_Get_Node_Amt to Public
/
alter table ZLTOOLS.ZLPARAMETERS modify ����ֵ VARCHAR2(2000)
/
alter table ZLTOOLS.ZLPARAMETERS modify ȱʡֵ VARCHAR2(2000)
/
alter table ZLTOOLS.zlUserParas modify ����ֵ VARCHAR2(2000)
/
--����
Create Or Replace Procedure zlTools.Zl_Parameters_Update
(
  ����_In   zlParameters.������%Type,
  ����ֵ_In zlParameters.����ֵ%Type,
  ϵͳ_In   zlParameters.ϵͳ%Type,
  ģ��_In   zlParameters.ģ��%Type
  --���ܣ�����ϵͳ����ֵ��������û�˽�в��������û����Ե�ǰ��Ϊ׼
  --������
  --      ����_In�����봫���Nullֵ�����ַ���ʽ����Ĳ����Ż������,ע�����������Ϊ���֡�
) Is
  v_����id zlParameters.ID%Type;
  v_˽��   zlParameters.˽��%Type;
  v_����   zlParameters.����%Type;
  v_������ zlUserParas.������%Type;
Begin
  --ȷ��������Ϣ
  Begin
    If Zl_To_Number(����_In) <> 0 Then
      --�Բ�����Ϊ׼����
      Select ID, ˽��, ����, Sys_Context('USERENV', 'TERMINAL')
      Into v_����id, v_˽��, v_����, v_������
      From zlParameters
      Where Nvl(ϵͳ, 0) = Nvl(ϵͳ_In, 0) And Nvl(ģ��, 0) = Nvl(ģ��_In, 0) And ������ = Zl_To_Number(����_In);
    Else
      --�Բ�����Ϊ׼����
      Select ID, ˽��, ����, Sys_Context('USERENV', 'TERMINAL')
      Into v_����id, v_˽��, v_����, v_������
      From zlParameters
      Where Nvl(ϵͳ, 0) = Nvl(ϵͳ_In, 0) And Nvl(ģ��, 0) = Nvl(ģ��_In, 0) And ������ = ����_In;
    End If;
  Exception
    When Others Then
      Return;
  End;

  --���²���ֵ
  If v_����id Is Not Null Then
    If Nvl(v_˽��, 0) = 0 And Nvl(v_����, 0) = 0 Then
      Update zlParameters Set ����ֵ = ����ֵ_In Where ID = v_����id;
    Else
      Update zlUserParas
      Set ����ֵ = ����ֵ_In
      Where ����id = v_����id And Nvl(�û���, 'NullUser') = Decode(v_˽��, 1, User, 'NullUser') And
            Nvl(������, 'NullMachine') = Decode(v_����, 1, v_������, 'NullMachine');
      If Sql%RowCount = 0 Then
        Insert Into zlUserParas
          (����id, �û���, ������, ����ֵ)
        Values
          (v_����id, Decode(v_˽��, 1, User, Null), Decode(v_����, 1, v_������, Null), ����ֵ_In);
      End If;
    End If;
  End If;
End Zl_Parameters_Update;
/

Create Or Replace Procedure zlTools.zl_Parameters_Update_Batch
(
  ϵͳ���_In zlSystems.���%Type,
  �����б�_In Varchar2
) Is
  --�����б�_IN ��������д��ʽ���£�"������1,����ֵ1,������2,����ֵ2,"                                            
  n_Pos    Number(5);
  v_Temp   Varchar2(2000);
  v_������ zlParameters.������%Type;
  v_����ֵ zlParameters.����ֵ%Type;
Begin
  --ѭ������
  v_Temp := �����б�_In;

  While v_Temp Is Not Null Loop
    n_Pos := Instr(v_Temp, ',');
  
    If n_Pos = 0 Then
      v_Temp := '';
    Else
      --�õ�������
      v_������ := Substr(v_Temp, 1, n_Pos - 1);
      v_Temp   := Substr(v_Temp, n_Pos + 1);
      --�õ�����ֵ
      n_Pos    := Instr(v_Temp, ',');
      v_����ֵ := Substr(v_Temp, 1, n_Pos - 1);
      v_Temp   := Substr(v_Temp, n_Pos + 1);
    
      Update zlParameters
      Set ����ֵ = v_����ֵ
      Where ϵͳ = ϵͳ���_In And ģ�� Is Null And ������ = To_Number(v_������);
    End If;
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End zl_Parameters_Update_Batch;
/
--********************************************************************************************************
--�����������End
--********************************************************************************************************

--15040
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

Create Or Replace Procedure Zltools.Zl_Createpubsynonyms Authid Current_User As
  v_Sql Varchar2(100);

  Cursor c_All Is
    Select Object_Name ����, Owner ������
    From All_Objects A
    Where Owner = User And Instr(Object_Name, 'BIN$') = 0 And
          Object_Type In ('TABLE', 'VIEW', 'SEQUENCE', 'PROCEDURE', 'FUNCTION') And Exists
     (Select 1 From Zlsystems Where Upper(������) = User)
    Minus
    Select Synonym_Name, User
    From All_Synonyms
    Where (Table_Owner In (Select Distinct Upper(������) From Zlsystems) Or Table_Owner = 'ZLTOOLS') And Owner = 'PUBLIC';
  --�������ϵͳ��ͬ��ͬ���,�򲻴�������ͬ���,���û�����ģ��ʱ�ٴ���˽��ͬ���  
Begin

  For c_Syn In c_All Loop
    Begin
      v_Sql := 'Create Public Synonym ' || c_Syn.���� || ' For ' || c_Syn.����;
      Execute Immediate v_Sql;
    Exception
      When Others Then
        Null;
    End;
  End Loop;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_Createpubsynonyms;
/


Create Public SYNONYM Zl_Createpubsynonyms FOR zlTools.Zl_Createpubsynonyms
/
Grant Execute on zlTools.Zl_Createpubsynonyms to Public
/


--15026���Զ��������ӹ�����ĵ���(Get_noticereport)��2009-01-08��By Fr.Chen
Create Or Replace Package Body b_Expert Is

  -----------------------------------------------------------------------------
  -- ȡ��������
  -----------------------------------------------------------------------------
  Procedure Get_Notices
  (
    Cursor_Out Out t_Refcur,
    ���_In    In zlNotices.���%Type,
    ϵͳ_In    In zlReports.ϵͳ%Type := Null
  ) Is
  Begin
    If Nvl(���_In, 0) <> 0 Then
      -- frmNoticesEdit.ReadData ʹ��
      Open Cursor_Out For
        Select A.��������, A.��������, A.���ѱ���, A.��������, A.���Ѵ���, A.��ʼʱ��, A.��ֹʱ��, A.�������, B.���� As ��������
        From zlNotices A, zlReports B
        Where A.���ѱ��� = B.���(+) And A.��� = ���_In;
    Else
      -- cboSystem_Click ʹ��
      If Nvl(ϵͳ_In, 0) = 0 Then
        Open Cursor_Out For
          Select A.���, A.��������, A.��������, A.���ѱ���, A.��������, A.���Ѵ���, A.��ʼʱ��, A.��ֹʱ��, A.�������, A.��������, B.���� As ��������
          From zlNotices A, zlReports B
          Where A.���ѱ��� = B.���(+) And A.ϵͳ Is Null;
      Else
        Open Cursor_Out For
          Select A.���, A.��������, A.��������, A.���ѱ���, A.��������, A.���Ѵ���, A.��ʼʱ��, A.��ֹʱ��, A.�������, A.��������, B.���� As ��������
          From zlNotices A, zlReports B
          Where A.���ѱ��� = B.���(+) And A.ϵͳ = ϵͳ_In;
      End If;
    End If;
  
  End Get_Notices;

  -----------------------------------------------------------------------------
  -- ȡ���Ѷ�������
  -----------------------------------------------------------------------------
  Procedure Get_Noticeusr
  (
    Cursor_Out  Out t_Refcur,
    ���Ѷ���_In In zlNoticeUsr.���Ѷ���%Type,
    �������_In In zlNoticeUsr.�������%Type
  ) Is
  Begin
    If Nvl(���Ѷ���_In, 0) = 0 Then
      Open Cursor_Out For
        Select 1 From zlNoticeUsr Where Rownum < 2 And ������� = �������_In;
    Else
      Open Cursor_Out For
        Select �������� From zlNoticeUsr Where ���Ѷ��� = ���Ѷ���_In And ������� = �������_In;
    End If;
  End Get_Noticeusr;

  -----------------------------------------------------------------------------
  -- ȡ����ѡ������ѱ���
  -----------------------------------------------------------------------------
  Procedure Get_Noticereport
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In zlReports.ϵͳ%Type
  ) Is
  Begin
    If Nvl(ϵͳ_In, 0) = 0 Then
      Open Cursor_Out For
        Select ID, ���, ����, ˵��
        From zlReports
        Where Not (����ʱ�� Is Null Or Trunc(����ʱ��) = To_Date('3000-01-01', 'yyyy-mm-dd')) And
              Nvl(ϵͳ, 0) = 0;
    Else
      Open Cursor_Out For
        Select ID, ���, ����, ˵��
        From zlReports
        Where ((ϵͳ = ϵͳ_In And ��� Like 'ZL%_REPORT_%') Or ϵͳ Is Null)  And Not (����ʱ�� Is Null Or Trunc(����ʱ��) = To_Date('3000-01-01', 'yyyy-mm-dd'));
    End If;
  End Get_Noticereport;

  -----------------------------------------------------------------------------
  -- �ڲ�ͬ��ϵͳ�临�Ʊ���
  -----------------------------------------------------------------------------
  Procedure Copy_Report
  (
    ϵͳ_In   In zlReports.ϵͳ%Type,
    ��ϵͳ_In In zlReports.ϵͳ%Type
  ) Is
    n_Grpid   Number;
    n_Rptid   Number;
    n_Dataid  Number;
    n_Itemid  Number;
    v_Olduser Varchar2(100);
    v_Newuser Varchar2(100);
  
    Function Sub_Owner_Name(Lngsys_In In Number := 0) Return Varchar2 Is
      v_Owner_Name Varchar2(30);
    Begin
      Select Upper(������) As ������ Into v_Owner_Name From zlSystems Where ��� = Lngsys_In;
      Return v_Owner_Name;
    End Sub_Owner_Name;
  
  Begin
    Select Nvl(Max(ID), 0) Into n_Grpid From zlRPTGroups;
    Select Nvl(Max(ID), 0) Into n_Rptid From zlReports;
    Select Nvl(Max(ID), 0) Into n_Dataid From zlRPTDatas;
    Select Nvl(Max(ID), 0) Into n_Itemid From zlRPTItems;
    n_Grpid  := n_Grpid + 1;
    n_Rptid  := n_Rptid + 1;
    n_Dataid := n_Dataid + 1;
    n_Itemid := n_Itemid + 1;
  
    v_Olduser := Upper(Sub_Owner_Name(ϵͳ_In));
    v_Newuser := Upper(Sub_Owner_Name(��ϵͳ_In));
  
    Insert Into zlRPTGroups
      (ID, ���, ����, ˵��, ϵͳ, ����id, ����ʱ��)
      Select ID + n_Grpid, ���, ����, ˵��, ��ϵͳ_In, ����id, ����ʱ�� From zlRPTGroups Where ϵͳ = ϵͳ_In;
  
    Insert Into zlReports
      (ID, ���, ����, ˵��, ����, ��ֽ, ��ӡ��, Ʊ��, ϵͳ, ����id, ����, �޸�ʱ��, ����ʱ��)
      Select ID + n_Rptid, ���, ����, ˵��, ����, ��ֽ, ��ӡ��, Ʊ��, ��ϵͳ_In, ����id, ����, �޸�ʱ��, ����ʱ��
      From zlReports
      Where ϵͳ = ϵͳ_In;
  
    -- ����zlRPTSub
    Insert Into zlRPTSubs
      (��id, ����id, ���, ����)
      Select A.��id + n_Grpid, A.����id + n_Rptid, A.���, A.����
      From zlRPTSubs A, zlRPTGroups B
      Where A.��id = B.ID And B.ϵͳ = ϵͳ_In;
  
    -- ����zlRPTFMTs
    Insert Into zlRPTFMTs
      (����id, ���, ˵��, W, H, ֽ��, ֽ��, ��ֽ̬��, ͼ��)
      Select A.����id + n_Rptid, A.���, A.˵��, A.W, A.H, A.ֽ��, A.ֽ��, A.��ֽ̬��, A.ͼ��
      From zlRPTFMTs A, zlReports B
      Where A.����id = B.ID And B.ϵͳ = ϵͳ_In;
  
    -- ����zlRPTItems
    Insert Into zlRPTItems
      (ID, ����id, ��ʽ��, ����, ����, �ϼ�id, ���, ����, ����, ����, ��ͷ, X, Y, W, H, �и�, ����, �Ե�, ����, �ֺ�, ����, б��, ����, ǰ��, ����, �߿�, ����, ��ʽ,
       ����, ����, ����, ϵͳ)
      Select A.ID + n_Itemid, A.����id + n_Rptid, A.��ʽ��, A.����, A.����, A.�ϼ�id + n_Itemid, A.���, A.����, A.����, A.����, A.��ͷ, A.X,
             A.Y, A.W, A.H, A.�и�, A.����, A.�Ե�, A.����, A.�ֺ�, A.����, A.б��, A.����, A.ǰ��, A.����, A.�߿�, A.����, A.��ʽ, A.����, A.����,
             A.����, A.ϵͳ
      From zlRPTItems A, zlReports B
      Where A.����id = B.ID And B.ϵͳ = ϵͳ_In;
    -- ����zlRptDatas
    Insert Into zlRPTDatas
      (ID, ����id, ����, �ֶ�, ����, ����)
      Select A.ID + n_Dataid, A.����id + n_Rptid, A.����, A.�ֶ�, A.����, A.����
      From zlRPTDatas A, zlReports B
      Where A.����id = B.ID And B.ϵͳ = ϵͳ_In;
    -- ����zlRPTSqls
    Insert Into zlRPTSQLs
      (Դid, �к�, ����)
      Select A.Դid + n_Dataid, A.�к�, A.����
      From zlRPTSQLs A, zlRPTDatas B, zlReports C
      Where A.Դid = B.ID And B.����id = C.ID And C.ϵͳ = ϵͳ_In;
  
    -- ����zlRPTPars
    Insert Into zlRPTPars
      (Դid, ����, ���, ����, ����, ȱʡֵ, ��ʽ, ֵ�б�, ����sql, ��ϸsql, �����ֶ�, ��ϸ�ֶ�, ����)
      Select A.Դid + n_Dataid, A.����, A.���, A.����, A.����, A.ȱʡֵ, A.��ʽ, A.ֵ�б�, A.����sql, A.��ϸsql, A.�����ֶ�, A.��ϸ�ֶ�, A.����
      From zlRPTPars A, zlRPTDatas B, zlReports C
      Where A.Դid = B.ID And B.����id = C.ID And C.ϵͳ = ϵͳ_In;
  
    -- zlFunctions����
    Insert Into zlFunctions
      (ϵͳ, ������, ������, ������, ˵��)
      Select ��ϵͳ_In, ������, ������, ������, ˵�� From zlFunctions Where ϵͳ = ϵͳ_In;
  
    -- zlFuncPars����
    Insert Into zlFuncPars
      (ϵͳ, ������, ������, ������, ������, ����, ȱʡֵ, ֵ�б�, ����sql, ��ϸsql, �����ֶ�, ��ϸ�ֶ�, ����, ����, ������)
      Select ��ϵͳ_In, ������, ������, ������, ������, ����, ȱʡֵ, ֵ�б�, ����sql, ��ϸsql, �����ֶ�, ��ϸ�ֶ�, ����, ����, ������
      From zlFuncPars
      Where ϵͳ = ϵͳ_In;
  
    -- ������������Դ����
    Update zlRPTDatas
    Set ���� = Replace(����, v_Olduser || '.', v_Newuser || '.')
    Where ID In (Select A.ID From zlRPTDatas A, zlReports B Where A.����id = B.ID And B.ϵͳ = ��ϵͳ_In);
  
    Update zlRPTPars
    Set ���� = Replace(����, v_Olduser || '.', v_Newuser || '.')
    Where Դid In (Select A.ID From zlRPTDatas A, zlReports B Where A.����id = B.ID And B.ϵͳ = ��ϵͳ_In);
  
    Update zlFuncPars Set ���� = Replace(����, v_Olduser || '.', v_Newuser || '.') Where ϵͳ = ��ϵͳ_In;
  
    Commit;
  Exception
    When Others Then
      zl_ErrorCenter(SQLCode, SQLErrM);
  End Copy_Report;

End b_Expert;
/
--15026���Զ��������ӹ�����ĵ���(Get_Zlnoticerec)��2009-01-08��By Fr.Chen
Create Or Replace Package Body b_Comfunc Is

  -----------------------------------------------------------------------------
  -- ���ܣ����������־
  -----------------------------------------------------------------------------
  Procedure Save_Error_Log
  (
    ����_In     In zlErrorLog.����%Type,
    �������_In In zlErrorLog.�������%Type,
    ������Ϣ_In In zlErrorLog.������Ϣ%Type
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
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Save_Error_Log;

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ���ù���
  -----------------------------------------------------------------------------
  Procedure Get_Usable_Function
  (
    Cursor_Out Out t_Refcur,
    ����_In    In zlPrograms.����%Type
  ) Is
  Begin
    If Nvl(����_In, '�տ�') = '�տ�' Then
      Open Cursor_Out For
        Select Distinct A.���, A.����, A.˵��
        From zlPrograms A, zlProgFuncs B, Zlregfunc C
        Where A.ϵͳ = B.ϵͳ And A.��� = B.��� And Trunc(A.ϵͳ / 100) = C.ϵͳ And A.��� = C.���
        Order By A.���;
    Else
      Open Cursor_Out For
        Select Distinct A.���, A.����, A.˵��
        From zlPrograms A, zlProgFuncs B, Zlregfunc C
        Where A.ϵͳ = B.ϵͳ And A.��� = B.��� And Upper(A.����) = Upper(����_In) And Trunc(A.ϵͳ / 100) = C.ϵͳ And
              A.��� = C.���
        Order By A.���;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Usable_Function;

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ��д���
  -----------------------------------------------------------------------------    
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
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Uppmoney;

  -----------------------------------------------------------------------------
  -- ���ܣ�����ָ�������ڡ���š�ϵͳ�ж�ָ�����ڵ������Ƿ���ת���������ݱ���
  -----------------------------------------------------------------------------    
  Procedure Get_Datamoved
  (
    Cursor_Out  Out t_Refcur,
    ���_In     In zlDataMove.���%Type,
    ϵͳ_In     In zlDataMove.ϵͳ%Type,
    �ϴ�����_In In zlDataMove.�ϴ�����%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select ϵͳ, ���
      From zlDataMove
      Where ��� = ���_In And ϵͳ = ϵͳ_In And �ϴ����� > �ϴ�����_In And �ϴ����� Is Not Null;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Datamoved;

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡϵͳ������
  -----------------------------------------------------------------------------
  Procedure Get_Owner
  (
    Cursor_Out Out t_Refcur,
    ���_In    In zlSystems.���%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select ������ From zlSystems Where ��� = ���_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Owner;

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ����
  -----------------------------------------------------------------------------
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
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Spell_Code;

  -----------------------------------------------------------------------------
  -- ���ܣ�����������־
  -----------------------------------------------------------------------------
  Procedure Save_Diary_Log
  (
    ������_In   In zlDiaryLog.������%Type,
    ������_In   In zlDiaryLog.������%Type,
    ��������_In In zlDiaryLog.��������%Type
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
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Save_Diary_Log;

  -----------------------------------------------------------------------------
  -- ���ܣ�����������־
  -- �����б�
  -- clsComLib.SaveWinState
  -----------------------------------------------------------------------------
  Procedure Update_Diary_Log
  (
    ������_In In zlDiaryLog.������%Type,
    ������_In In zlDiaryLog.������%Type
  ) Is
    Cursor c_Session Is
      Select Sid + Serial# As �Ự��, User As �û���, RTrim(LTrim(Replace(Machine, Chr(0), ''))) As ����վ
      From V$session
      Where Audsid = Userenv('SessionID');
  Begin
    For r_Tmp In c_Session Loop
      Update zlDiaryLog
      Set �˳�ԭ�� = 1, �˳�ʱ�� = Sysdate
      Where �˳�ԭ�� Is Null And �û��� = r_Tmp.�û��� And ����վ = r_Tmp.����վ And �Ự�� = r_Tmp.�Ự�� And
            ������ = ������_In And ������ = ������_In;
    End Loop;
    Commit;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Update_Diary_Log;

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ�̶�����������û���������
  -----------------------------------------------------------------------------
  Procedure Get_Report_Menu
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In zlPrograms.ϵͳ%Type,
    ���_In    In zlPrograms.���%Type,
    ����_In    In zlReports.����%Type,
    ���_In    In zlReports.���%Type
  ) Is
  Begin
    If Nvl(���_In, '�տ�') <> '�տ�' Then
      Open Cursor_Out For
        Select ��־, ϵͳ, ���, ����
        From (Select 1 As ��־, A.ϵͳ, A.���, A.����
               From zlReports A, zlPrograms B
               Where A.ϵͳ = B.ϵͳ And A.����id = B.��� And Not Upper(A.���) Like '%BILL%' And
                     Upper(B.����) <> Upper('zl9Report') And B.ϵͳ = ϵͳ_In And B.��� = ���_In And
                     Instr(����_In, ';' || A.���� || ';') > 0
               Union All
               Select Decode(A.ϵͳ, Null, 2, 1) As ��־, A.ϵͳ, A.���, A.����
               From zlReports A, zlRPTPuts B, zlPrograms C
               Where A.ID = B.����id And B.ϵͳ = C.ϵͳ And B.����id = C.��� And
                     (Not Upper(A.���) Like '%BILL%' Or A.ϵͳ Is Null) And Instr(����_In, ';' || B.���� || ';') > 0 And
                     C.ϵͳ = ϵͳ_In And C.��� = ���_In)
        Where Instr(���_In, ',' || ��� || ',') = 0
        Order By ��־, ���;
    Else
      Open Cursor_Out For
        Select ��־, ϵͳ, ���, ����
        From (Select 1 As ��־, A.ϵͳ, A.���, A.����
               From zlReports A, zlPrograms B
               Where A.ϵͳ = B.ϵͳ And A.����id = B.��� And Not Upper(A.���) Like '%BILL%' And
                     Upper(B.����) <> Upper('zl9Report') And B.ϵͳ = ϵͳ_In And B.��� = ���_In And
                     Instr(����_In, ';' || A.���� || ';') > 0
               Union All
               Select Decode(A.ϵͳ, Null, 2, 1) As ��־, A.ϵͳ, A.���, A.����
               From zlReports A, zlRPTPuts B, zlPrograms C
               Where A.ID = B.����id And B.ϵͳ = C.ϵͳ And B.����id = C.��� And
                     (Not Upper(A.���) Like '%BILL%' Or A.ϵͳ Is Null) And Instr(����_In, ';' || B.���� || ';') > 0 And
                     C.ϵͳ = ϵͳ_In And C.��� = ���_In)
        Order By ��־, ���;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Report_Menu;

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ�û�������Ϣ
  -----------------------------------------------------------------------------
  Procedure Get_Zlnoticerec
  (
    Cursor_Out Out t_Refcur,
    �û���_In  In zlNoticeRec.�û���%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select A.���, A.ϵͳ, C.����id As ģ��,C.ϵͳ As ����ϵͳ, B.�������� As �������, C.���� As ���ѱ���, A.��������, B.���ʱ��,
             B.�Ѷ���־
      From zlNotices A, zlNoticeRec B, (Select * From zlReports Where ����ʱ�� Is Not Null) C
      Where B.�û��� = �û���_In And B.���ѱ�־ > 0 And C.���(+) = A.���ѱ��� And A.��� = B.������� And
            B.�������� Is Not Null;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Zlnoticerec;

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ�ʼ�����
  -----------------------------------------------------------------------------
  Procedure Get_Zlmessage
  (
    Cursor_Out Out t_Refcur,
    Id_In      In zlMessages.ID%Type,
    ����_In    In zlMsgState.����%Type,
    �û�_In    In zlMsgState.�û�%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select A.*, B.ɾ��, B.״̬
      From zlMessages A, zlMsgState B
      Where A.ID = B.��Ϣid And B.��Ϣid = Id_In And B.���� = ����_In And B.�û� = �û�_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Zlmessage;

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ�ʼ�����
  -----------------------------------------------------------------------------
  Procedure Get_Zlmessage
  (
    Cursor_Out Out t_Refcur,
    Id_In      In zlMessages.ID%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select ����, ����ɫ From zlMessages Where ID = Id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Zlmessage;

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ�ʵݵ�ַ
  -----------------------------------------------------------------------------
  Procedure Get_Zlmsgstate
  (
    Cursor_Out Out t_Refcur,
    ��Ϣid_In  In zlMsgState.��Ϣid%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select ����, �û�, ��� From zlMsgState Where ��Ϣid = ��Ϣid_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Zlmsgstate;

  -----------------------------------------------------------------------------
  -- ���ܣ�ɾ����Ϣ
  -----------------------------------------------------------------------------
  Procedure Delete_Zlmsgstate
  (
    ɾ��_In   In zlMsgState.ɾ��%Type,
    ��Ϣid_In In zlMsgState.��Ϣid%Type,
    ����_In   In zlMsgState.����%Type,
    �û�_In   In zlMsgState.�û�%Type
  ) Is
    n_���� Number(10);
    n_���� Number(10);
  Begin
    If Nvl(ɾ��_In, 0) = 1 Then
      Update zlMsgState Set ɾ�� = 1 Where ��Ϣid = ��Ϣid_In And ���� = ����_In And �û� = �û�_In;
    Else
      If ����_In = 0 Then
        -- ���ڲݸ壬���ռ��˵�Ҳһ��ɾ��
        Update zlMsgState Set ɾ�� = 2 Where ��Ϣid = ��Ϣid_In And �û� = �û�_In;
      Else
        Update zlMsgState Set ɾ�� = 2 Where ��Ϣid = ��Ϣid_In And ���� = ����_In And �û� = �û�_In;
      End If;
      --  ɾ��ָ��ID����Ϣ  mnuEditDelete_Click ����
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
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Delete_Zlmsgstate;

  -----------------------------------------------------------------------------
  -- ���ܣ�ɾ��������Ϣ
  -----------------------------------------------------------------------------
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
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Delete_Zlmessage;

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ�ʼ��б�
  -----------------------------------------------------------------------------
  Procedure Get_Mail_List
  (
    Cursor_Out  Out t_Refcur,
    ��Ϣ����_In In Varchar2,
    �û�_In     In zlMsgState.�û�%Type,
    ��ʾ�Ѷ�_In In Number,
    �Ựid_In   In zlMessages.�Ựid%Type
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
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Mail_List;

  -----------------------------------------------------------------------------
  -- ���ܣ���ԭɾ������Ϣ
  -----------------------------------------------------------------------------
  Procedure Restore_Zlmsgstate
  (
    ��Ϣid_In In zlMsgState.��Ϣid%Type,
    ����_In   In zlMsgState.����%Type,
    �û�_In   In zlMsgState.�û�%Type
  ) Is
  Begin
    Update zlMsgState Set ɾ�� = 0 Where ��Ϣid = ��Ϣid_In And ���� = ����_In And �û� = �û�_In;
    Commit;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Restore_Zlmsgstate;

  -----------------------------------------------------------------------------
  -- ���ܣ�������Ϣ
  -- �����б�
  -- zlApptools.frmMessageEdit.SaveMessage
  -----------------------------------------------------------------------------
  Procedure Save_Zlmessage
  (
    Cursor_Out Out t_Refcur,
    Id_In      In zlMessages.ID%Type,
    �Ựid_In  In zlMessages.�Ựid%Type,
    ������_In  In zlMessages.������%Type,
    �ռ���_In  In zlMessages.�ռ���%Type,
    ����_In    In zlMessages.����%Type,
    ����_In    In zlMessages.����%Type,
    ����ɫ_In  In zlMessages.����ɫ%Type
  ) Is
    n_Id     zlMessages.ID%Type;
    n_�Ựid zlMessages.�Ựid%Type;
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
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Save_Zlmessage;

  -----------------------------------------------------------------------------
  -- ���ܣ�����zlMsgstate
  -- �����б�
  -- zlApptools.frmMessageEdit.SaveMessage
  -----------------------------------------------------------------------------
  Procedure Insert_Zlmsgstate
  (
    ��Ϣid_In In zlMsgState.��Ϣid%Type,
    ����_In   In zlMsgState.����%Type,
    �û�_In   In zlMsgState.�û�%Type,
    ���_In   In zlMsgState.���%Type,
    ɾ��_In   In zlMsgState.ɾ��%Type,
    ״̬_In   In zlMsgState.״̬%Type
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
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Insert_Zlmsgstate;

  -----------------------------------------------------------------------------
  -- ���ܣ�Ϊԭ�����ϴ𸴻�ת����־
  -----------------------------------------------------------------------------
  Procedure Update_Zlmsgstate_State
  (
    ģʽ_In   In Number,
    ��Ϣid_In In zlMsgState.��Ϣid%Type,
    ����_In   In zlMsgState.����%Type,
    �û�_In   In zlMsgState.�û�%Type
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
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Update_Zlmsgstate_State;

  -----------------------------------------------------------------------------
  -- ���ܣ�����״̬�����
  -----------------------------------------------------------------------------
  Procedure Update_Zlmsgstate_Idtntify
  (
    ���_In   In zlMsgState.���%Type,
    ��Ϣid_In In zlMsgState.��Ϣid%Type,
    ����_In   In zlMsgState.����%Type,
    �û�_In   In zlMsgState.�û�%Type
  ) Is
  Begin
    Update zlMsgState
    Set ״̬ = '1' || Substr(״̬, 2), ��� = ���_In
    Where ��Ϣid = ��Ϣid_In And ���� = ����_In And �û� = �û�_In;
    Commit;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Update_Zlmsgstate_Idtntify;

End b_Comfunc;
/


--����������ظ���


Create Or Replace Package zlTools.b_Runmana Is
  -----------------------------------------------------------------------------
  -- ���ߣ� �¶�
  -- ��ʼʱ�䣺2006-6-29
  -- �޸��ˣ�
  -- �޸�ʱ�䣺
  -- ������
  --         ��Ҫ�������й����ܵĹ���
  -----------------------------------------------------------------------------

  Type t_Refcur Is Ref Cursor;

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ������Ϣ
  -- �޸ģ����˺�
  -- �����б�
  -- frmParameters
  -----------------------------------------------------------------------------
  Procedure Get_Parameters
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In Number
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�����ָ���Ĳ���IDȡ������Ϣ
  -- �޸ģ����˺�
  -- �����б�
  -- frmParameters;frmParaChangeSet
  -----------------------------------------------------------------------------
  Procedure Get_Parameter
  (
    Cursor_Out Out t_Refcur,
    ����id_In  In zlParameters.id%Type
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡվ����û��Ĳ�����Ϣ
  -- �޸ģ����˺�
  -- �����б�
  -- frmParameters
  -----------------------------------------------------------------------------
  Procedure Get_Userparameters
  (
    Cursor_Out Out t_Refcur,
    ����id_In  In zlUserParas.����id%Type,
    intType    IN NUMBER :=0 
    --0-���в�����Ϣ,1-ֻ��ȡ������������,2-ֻ��ȡ�û���
  );
 
  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ�����޸���Ϣ
  -- �޸ģ����˺�
  -- �����б�
  -- frmParameters
  -----------------------------------------------------------------------------
  Procedure Get_Parachangedlog
  (
    Cursor_Out Out t_Refcur,
    ����id_In  In Zlparachangedlog.����id%Type
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡZlAutoJob���к�
  -- �����б�
  -- frmAutoJobset.cmdok_click
  -----------------------------------------------------------------------------
  Procedure Get_Job_Number
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In Number
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡZlDataMove����
  -- �����б�
  -- frmAutoJobset.cmdUpdate_Click
  -----------------------------------------------------------------------------
  Procedure Get_Depict
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In zlDataMove.ϵͳ%Type,
    ���_In    In zlDataMove.���%Type
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡzlClients��MAX IP
  -- �����б�
  -- frmClientsEdit.InitCard
  -----------------------------------------------------------------------------
  Procedure Get_Client_Maxip(Cur_Out Out t_Refcur);

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡzlClients�ļ�¼
  -- �����б�
  -- frmClientsEdit.InitCard��frmClientsParas.LoadClientsInfor��frmClientsUpgrade.LoadClientsInfor
  -- frmFilesSendToServer.LoadClientsInfor
  -----------------------------------------------------------------------------
  Procedure Get_Client
  (
    Cur_Out   Out t_Refcur,
    ����վ_In In zlClients.����վ%Type := Null
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡzlClients��վ��
  -- �����б�
  -- frmClientsParasSet.InitCard
  -----------------------------------------------------------------------------
  Procedure Get_Client_Station(Cur_Out Out t_Refcur);

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ������
  -- �����б�
  -- frmClientsParasSet.InitCard
  -----------------------------------------------------------------------------
  Procedure Get_Project_No(Cur_Out Out t_Refcur);

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ����
  -- �����б�
  -- frmClientsParasSet.InitCard
  -----------------------------------------------------------------------------
  Procedure Get_Client_Scheme(Cur_Out Out t_Refcur);

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ�ָ���Ϣ
  -- �����б�
  -- frmClientsParasSet.Load�ָ�������frmClientsParasSet.LoadScremeSet
  -----------------------------------------------------------------------------
  Procedure Get_Resile
  (
    Cur_Out   Out t_Refcur,
    ������_In In Zlclientparaset.������%Type,
    ����_In   In Number := 0
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡzldataMove����
  -- �����б�
  -- frmDataMove.cmbSystem_Click
  -----------------------------------------------------------------------------
  Procedure Get_Zldatamove
  (
    Cur_Out Out t_Refcur,
    ϵͳ_In In zlDataMove.ϵͳ%Type
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ��־����
  -- �����б�
  -- FrmErrLog.RefreshData��FrmRunLog.RefreshData
  -----------------------------------------------------------------------------
  Procedure Get_Log
  (
    Cur_Out     Out t_Refcur,
    ��־����_In In Varchar2,
    Where_In    In Varchar2
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ��־��¼��
  -- �����б�
  -- FrmErrLog.DeleteExtra��FrmRunLog.DeleteExtra
  -----------------------------------------------------------------------------
  Procedure Get_Log_Count
  (
    Cur_Out     Out t_Refcur,
    ��־����_In In Varchar2
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡzlfilesupgradeg����
  -- �����б�
  -- frmFilesSet.intBillInfor
  -----------------------------------------------------------------------------
  Procedure Get_Zlfilesupgrade(Cur_Out Out t_Refcur);

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ��ע����Ŀ
  -- �����б�
  -- frmRegist.Form_Load
  -----------------------------------------------------------------------------
  Procedure Get_Not_Regist(Cur_Out Out t_Refcur);

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ����ֵ
  -- �����б�
  -- FrmRunOption.InitCons
  -----------------------------------------------------------------------------
  Procedure Get_Zloption
  (
    Cur_Out   Out t_Refcur,
    ������_In In zlOptions.������%Type
  );

End b_Runmana;
/

Create Or Replace Package Body zltools.b_Runmana Is

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ������Ϣ
  -- �޸ģ����˺�
  -- �����б�
  -- frmParameters
  -----------------------------------------------------------------------------
  Procedure Get_Parameters
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In Number
  ) Is
  Begin
    If Nvl(ϵͳ_In, 0) = 0 Then
      Open Cursor_Out For
        Select A.ID, A.ϵͳ, A.ģ��, A.˽��, A.������, A.������, A.����ֵ, A.ȱʡֵ, A.����˵��, A.����, A.��Ȩ, A.�̶�,
               B.���� As ģ������, zlSpellCode(B.����) As ģ�����
        From zlParameters A, zlPrograms B
        Where Nvl(A.ϵͳ, 0) = 0 And Nvl(A.ϵͳ, 0) = B.ϵͳ(+) And Nvl(A.ģ��, 0) = B.���(+);
    Else
      Open Cursor_Out For
        Select A.ID, A.ϵͳ, A.ģ��, A.˽��, A.������, A.������, A.����ֵ, A.ȱʡֵ, A.����˵��, A.����, A.��Ȩ, A.�̶�,
               B.���� As ģ������, zlSpellCode(B.����) As ģ�����
        From zlParameters A, zlPrograms B,
             --����Ȩ�޲��֣�ֻ����Ȩ�Ĳ�����ʾ
             (Select Distinct F.���
               From zlProgFuncs F, zlRegFunc R
               Where Trunc(F.ϵͳ / 100) = R.ϵͳ And F.��� = R.��� And F.ϵͳ = ϵͳ_In And F.���� = R.���� And
                     1 = (Select 1 From Zlregaudit A Where A.��Ŀ = '��Ȩ֤��')
               Union All
               Select 0 As ��� From Dual) M
        Where A.ϵͳ = Nvl(ϵͳ_In, 0) And Nvl(A.ϵͳ, 0) = B.ϵͳ(+) And Nvl(A.ģ��, 0) = B.���(+) And
              Nvl(A.ģ��, 0) = M.���;
    End If;
  End Get_Parameters;

  -----------------------------------------------------------------------------
  -- ���ܣ�����ָ���Ĳ���IDȡ������Ϣ
  -- �޸ģ����˺�
  -- �����б�
  -- frmParameters;frmParaChangeSet
  -----------------------------------------------------------------------------
  Procedure Get_Parameter
  (
    Cursor_Out Out t_Refcur,
    ����id_In  In zlParameters.ID%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select A.ID, A.ϵͳ, A.ģ��, A.˽��, A.������, A.������, A.����ֵ, A.ȱʡֵ, A.����˵��, A.����, A.��Ȩ, A.�̶�,
             B.���� As ģ������, zlSpellCode(B.����) As ģ�����
      From zlParameters A, zlPrograms B
      Where A.ID = Nvl(����id_In, 0) And Nvl(A.ϵͳ, 0) = B.ϵͳ(+) And Nvl(A.ģ��, 0) = B.���(+);
  End Get_Parameter;

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡվ����û��Ĳ�����Ϣ
  -- �޸ģ����˺�
  -- �����б�
  -- frmParameters
  -----------------------------------------------------------------------------
  Procedure Get_Userparameters
  (
    Cursor_Out Out t_Refcur,
    ����id_In  In zlUserParas.����id%Type,
    Inttype    In Number := 0
    --0-���в�����Ϣ,1-ֻ��ȡ������������,2-ֻ��ȡ�û���
  ) Is
    n_˽�� zlParameters.˽��%Type;
    n_���� zlParameters.����%Type;
  Begin
    If Inttype = 0 Then
      Begin
        Select Nvl(A.˽��, 0), Nvl(A.����, 0) Into n_˽��, n_���� From zlParameters A Where ID = Nvl(����id_In, 0);
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

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ�����޸���Ϣ
  -- �޸ģ����˺�
  -- �����б�
  -- frmParameters
  -----------------------------------------------------------------------------
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

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡZlAutoJob���к�
  -----------------------------------------------------------------------------
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

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡZlDataMove����
  -----------------------------------------------------------------------------
  Procedure Get_Depict
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In zlDataMove.ϵͳ%Type,
    ���_In    In zlDataMove.���%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select ת������ From zlDataMove Where Nvl(ϵͳ, 0) = ϵͳ_In And ��� = ���_In;
  End Get_Depict;

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡzlClients��MAX IP
  -----------------------------------------------------------------------------
  Procedure Get_Client_Maxip(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select Max(Ip) As Ip From zlClients;
  End Get_Client_Maxip;

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡzlClients�ļ�¼
  -----------------------------------------------------------------------------
  Procedure Get_Client
  (
    Cur_Out   Out t_Refcur,
    ����վ_In In zlClients.����վ%Type := Null
  ) Is
    v_Sql Varchar2(1000);
  Begin
    If Nvl(����վ_In, '��') = '��' Then
      v_Sql := 'Select a.Ip, a.����վ, a.Cpu, a.�ڴ�, a.Ӳ��, a.����ϵͳ, a.����, a.��;, a.˵��, a.������־, a.��ֹʹ��,
							 a.������, Decode(b.Terminal, Null, 0, 1) As ״̬, a.�ռ���־,a.����������
				From Zlclients a, (Select Distinct Terminal From V$session) b
				Where Upper(a.����վ) = Upper(b.Terminal(+))
				Order By a.Ip';
      Open Cur_Out For v_Sql;
    Else
      Open Cur_Out For
        Select Ip, ����վ, Cpu, �ڴ�, Ӳ��, ����ϵͳ, ����, ��;, ˵��, ������־, �ռ���־, ��ֹʹ��, ������, ����������
        From zlClients
        Where Upper(����վ) = ����վ_In;
    End If;
  End Get_Client;

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡzlClients��վ��
  -----------------------------------------------------------------------------
  Procedure Get_Client_Station(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select Distinct Upper(����վ) || '[' || Ip || ']' As վ��, Upper(����վ) ����վ From zlClients;
  End Get_Client_Station;

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ������
  -----------------------------------------------------------------------------
  Procedure Get_Project_No(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select ������ From Zlclientparaset Where Rownum = 1;
  End Get_Project_No;

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ����
  -----------------------------------------------------------------------------
  Procedure Get_Client_Scheme(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select ������, ������ || '-' || �������� As ��������, ��������, ����վ, �û��� From Zlclientscheme;
  End Get_Client_Scheme;

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ�ָ���Ϣ
  -----------------------------------------------------------------------------
  Procedure Get_Resile
  (
    Cur_Out   Out t_Refcur,
    ������_In In Zlclientparaset.������%Type,
    ����_In   In Number := 0
  ) Is
  Begin
    If ����_In = 0 Then
      Open Cur_Out For
        Select Distinct A.����վ || Decode(M.����վ, Null, ' ', '[' || M.Ip || ']') As ����վ, A.�û���, A.�ָ���־,
                        '[' || B.������ || ']' || B.�������� As ��������
        From Zlclientparaset A, Zlclientscheme B, zlClients M
        Where A.������ = B.������ And A.����վ = M.����վ(+) And A.������ = ������_In;
    End If;
  
    If ����_In = 1 Then
      Open Cur_Out For
        Select Distinct Upper(����վ) ����վ, Min(�ָ���־) �ָ���־
        From Zlclientparaset A
        Where A.������ = ������_In
        Group By ����վ;
    End If;
  
    If ����_In = 2 Then
      Open Cur_Out For
        Select Distinct Upper(�û���) �û���, Max(����վ) ����վ, Min(Decode(�ָ���־, 2, 0, �ָ���־)) �ָ���־
        From Zlclientparaset A
        Where A.������ = ������_In
        Group By �û���
        Order By �û���;
    End If;
  
  End Get_Resile;

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡzldataMove����
  -----------------------------------------------------------------------------
  Procedure Get_Zldatamove
  (
    Cur_Out Out t_Refcur,
    ϵͳ_In In zlDataMove.ϵͳ%Type
  ) Is
  Begin
    Open Cur_Out For
      Select ���, ����, ˵��, �����ֶ�, ת������, �ϴ����� From zlDataMove Where ϵͳ = ϵͳ_In Order By ���;
  End Get_Zldatamove;

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ��־����
  -----------------------------------------------------------------------------
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
					 ,Decode(����,1,''�洢���̴���'',2,''������������'',''Ӧ�ó�������'') ��������
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

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ��־��¼��
  -----------------------------------------------------------------------------
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
        Select Nvl(To_Number(����ֵ), 0) From zlOptions Where ������ = 4;
    End If;
    If ��־����_In = '������־' Then
      Open Cur_Out For
        Select Count(*) ����
        From zlDiaryLog
        Union All
        Select Nvl(To_Number(����ֵ), 0) From zlOptions Where ������ = 2;
    
    End If;
  End Get_Log_Count;

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡzlfilesupgradeg����
  -----------------------------------------------------------------------------
  Procedure Get_Zlfilesupgrade(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select A.���, A.�ļ���, A.�汾��, A.�޸�����, B.���� As ˵��
      From zlFilesUpgrade A, zlComponent B
      Where Upper(A.�ļ���) = Upper(B.����(+))
      Order By A.���;
  End Get_Zlfilesupgrade;

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ��ע����Ŀ
  -----------------------------------------------------------------------------
  Procedure Get_Not_Regist(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select ��Ŀ, ����
      From zlRegInfo
      Where ��Ŀ Not In ('������', '�汾��', '������Ŀ¼', '�����û�', '��������', '�ռ�Ŀ¼', '�ռ�����', 'ע����',
             '��Ȩ֤��', '��Ȩ����', '��Ȩ�ʴ�');
  End Get_Not_Regist;

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ����ֵ
  -----------------------------------------------------------------------------
  Procedure Get_Zloption
  (
    Cur_Out   Out t_Refcur,
    ������_In In zlOptions.������%Type
  ) Is
  Begin
    Open Cur_Out For
      Select Nvl(����ֵ, ȱʡֵ) Option_Value From zlOptions Where ������ = ������_In;
  End Get_Zloption;

End b_Runmana; 
/


Create Or Replace Procedure zlTools.zl_Parameters_Change
(
  ����id_In   zlParameters.ID%Type,
  ˽��_In     zlParameters.˽��%Type,
  ����_In     zlParameters.����%Type,
  ��Ȩ_In     zlParameters.��Ȩ%Type,
  �䶯��_In   Zlparachangedlog.�䶯��%Type,
  �䶯ԭ��_In Zlparachangedlog.�䶯ԭ��%Type
) Is
  v_Temp     Varchar2(200);
  n_ģ��     zlParameters.ģ��%Type;
  n_˽��     zlParameters.˽��%Type;
  n_����     zlParameters.����%Type;
  n_��Ȩ     zlParameters.��Ȩ%Type;
  n_���     Zlparachangedlog.���%Type;
  v_�䶯˵�� Zlparachangedlog.�䶯˵��%Type;
  v_�䶯���� Zlparachangedlog.�䶯����%Type;

  Function Gettype
  (
    ģ��_In zlParameters.˽��%Type,
    ˽��_In zlParameters.˽��%Type,
    ����_In zlParameters.����%Type
  ) Return Varchar2 Is
  
  Begin
  
    If Nvl(ģ��_In, 0) = 0 Then
      --����ģ��,֤��ֻ����������:����ȫ�ֺ�˽��ȫ��
      If Nvl(˽��_In, 0) = 0 Then
        Return '����ȫ��';
      End If;
      Return '˽��ȫ��';
    End If;
  
    --��ģ��Ĵ���
    If ����_In = 0 Then
      --���Ǳ��������,ֻ����������:����ģ���˽��ģ��
      If Nvl(˽��_In, 0) = 0 Then
        Return '����ģ��';
      End If;
      Return '˽��ģ��';
    End If;
    --�Ա�����ģ����д���Ҳ���������:
    If Nvl(˽��_In, 0) = 0 Then
      Return '��������ģ��';
    End If;
    Return '����˽��ģ��';
  Exception
    When Others Then
      Return Null;
  End Gettype;
Begin

  Select Nvl(ģ��, 0), Nvl(˽��, 0), Nvl(����, 0), Nvl(��Ȩ, 0)
  Into n_ģ��, n_˽��, n_����, n_��Ȩ
  From zlParameters
  Where ID = ����id_In;
  Select Nvl(Max(���), 0) + 1 Into n_��� From Zlparachangedlog Where ����id = ����id_In;
  --��������
  --˵���䶯˵��:����:˽��ģ���Ϊ����ģ�顣
  -- �䶯����:˵���䶯�ֶεı仯���:����:˽��:1-->0,����:1-->0
  v_�䶯˵�� := Null;
  v_�䶯���� := Null;
  If n_˽�� <> Nvl(˽��_In, 0) Or n_���� <> Nvl(����_In, 0) Then
    --���ͷ����˸ı�
    v_Temp     := '��' || Gettype(n_ģ��, n_˽��, n_����);
    v_Temp     := v_Temp || '��Ϊ' || Gettype(n_ģ��, Nvl(˽��_In, 0), Nvl(����_In, 0));
    v_�䶯˵�� := v_Temp;
    v_Temp     := '';
    If n_˽�� <> Nvl(˽��_In, 0) Then
      v_Temp := v_Temp || ',˽��:' || n_˽�� || '-->' || Nvl(˽��_In, 0);
    End If;
    If n_˽�� <> Nvl(˽��_In, 0) Then
      v_Temp := v_Temp || ',����:' || n_���� || '-->' || Nvl(����_In, 0);
    End If;
    v_�䶯���� := Substr(v_Temp, 2);
  End If;
  --�����Ȩ�����ı�û��
  If n_��Ȩ <> Nvl(��Ȩ_In, 0) Then
    If Not v_�䶯˵�� Is Null Then
      v_�䶯˵�� :=v_�䶯˵��|| ',';
    End If;
    If n_��Ȩ = 0 Then
      v_Temp := '����Ҫ��Ȩ';
    Else
      v_Temp := '��Ҫ��Ȩ';
    End If;
    v_�䶯˵�� := Nvl(v_�䶯˵��, '') || '��' || v_Temp || '��Ϊ';
    If ��Ȩ_In = 0 Then
      v_Temp := '����Ҫ��Ȩ';
    Else
      v_Temp := '��Ҫ��Ȩ';
    End If;
    v_�䶯˵�� := Nvl(v_�䶯˵��, '') || v_Temp;

    If Not v_�䶯���� Is Null Then
	v_�䶯����:=v_�䶯����||',';	
    End If;

    v_�䶯���� := Nvl(v_�䶯����, '') || '��Ȩ:' || n_��Ȩ || '-->' || Nvl(��Ȩ_In, 0);
   
  End If;

  Insert Into Zlparachangedlog
    (����id, ���, �䶯˵��, �䶯����, �䶯��, �䶯ʱ��, �䶯ԭ��)
  Values
    (����id_In, n_���, v_�䶯˵��, v_�䶯����, �䶯��_In, Sysdate, �䶯ԭ��_In);

  Update zlParameters Set ˽�� = Nvl(˽��_In, 0), ���� = Nvl(����_In, 0), ��Ȩ = Nvl(��Ȩ_In, 0) Where ID = ����id_In;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End  zl_Parameters_Change;
/


grant execute on zlTools.zl_Parameters_Change to Public
/

create public SYNONYM  zl_Parameters_Change   FOR zlTools.zl_Parameters_Change
/
