-----------------------------------------------------------------
--Ϊ��ϲ�Ʒ�汾����9.29��Ϊ9.29.10
-----------------------------------------------------------------
--10754
Alter Table zlTools.zlRPTFmts Add(W Number(18),H Number(18),ֽ�� Number(3),ֽ�� Number(1),��ֽ̬�� Number(1))
/
ALTER TABLE zlTools.zlRPTFmts ADD CONSTRAINT zlRPTFmts_CK_ֽ�� Check(ֽ�� IN(1,2)) PCTFREE 5
/
ALTER TABLE zlTools.zlRPTFmts ADD CONSTRAINT zlRPTFmts_CK_��ֽ̬�� Check(��ֽ̬�� IN(0,1)) PCTFREE 5
/
Begin
	For r_Report IN(Select * From zlReports) Loop
		Update zlRPTFMTs Set W=r_Report.W,H=r_Report.H,ֽ��=r_Report.ֽ��,ֽ��=r_Report.ֽ��,��ֽ̬��=r_Report.��ֽ̬�� Where ����ID=r_Report.ID;
	End Loop;
End;
/

Alter Table zlTools.zlReports Drop Column W
/
Alter Table zlTools.zlReports Drop Column H
/
Alter Table zlTools.zlReports Drop Column ֽ��
/
Alter Table zlTools.zlReports Drop Column ֽ��
/
Alter Table zlTools.zlReports Drop Column ��ֽ̬��
/

Create Or Replace Package Body zlTools.b_Expert Is
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
        Where ��� Like 'ZL%_REPORT_%' And Not (����ʱ�� Is Null Or Trunc(����ʱ��) = To_Date('3000-01-01', 'yyyy-mm-dd')) And
              Nvl(ϵͳ, 0) = 0;
    Else
      Open Cursor_Out For
        Select ID, ���, ����, ˵��
        From zlReports
        Where ��� Like 'ZL%_REPORT_%' And Not (����ʱ�� Is Null Or Trunc(����ʱ��) = To_Date('3000-01-01', 'yyyy-mm-dd')) And
              ϵͳ = ϵͳ_In;
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