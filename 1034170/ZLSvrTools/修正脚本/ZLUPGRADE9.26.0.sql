-----------------------------------------------------------------
--Ϊ��ϲ�Ʒ�汾����9.25��Ϊ9.26
--�ӱ�����ʼʹ�������������� 
-----------------------------------------------------------------
--�µ��Զ�����������־��¼
CREATE TABLE zlTools.zlUpgrade(
	ϵͳ NUMBER(5),
	ԭʼ�汾 VARCHAR2(10),
	Ŀ��汾 VARCHAR2(10),
	��Ǩʱ�� DATE,
	��Ǩ��� NUMBER(1),
	����汾 VARCHAR2(10),
	��ֹ��� VARCHAR2(200))
	PCTFREE 5 PCTUSED 90
/
ALTER TABLE zlTools.zlUpgrade ADD CONSTRAINT 
    zlUpgrade_UQ_��Ǩʱ�� Unique (ϵͳ,��Ǩʱ��)
    USING INDEX PCTFREE 5
/
ALTER TABLE zlTools.zlUpgrade ADD CONSTRAINT
    zlUpgrade_FK_ϵͳ FOREIGN KEY (ϵͳ) REFERENCES zlSystems(���) ON DELETE CASCADE
/
CREATE PUBLIC SYNONYM zlUpgrade for zlTools.zlUpgrade
/
GRANT SELECT ON zlTools.zlUpgrade TO PUBLIC 
/
Begin
	For r_User In(Select ������ From zlSystems) 
	Loop
			Execute Immediate 'Grant Select,Insert,Update,Delete on zlTools.zlUpgrade to '||r_User.������||' With Grant Option';
	End Loop;
End;
/

--����������Ψһ����,���SQLЧ��
Begin
	Begin Execute Immediate 'Drop Index zlTools.zlRPTDatas_IX_����Id'; Exception When Others Then Null; End;
	Begin Execute Immediate 'Drop Index zlTools.zlRPTConds_IX_����Id'; Exception When Others Then Null; End;

	Delete From zlTools.zlRPTDatas A Where RowID<(Select Max(RowID) From zlTools.zlRPTDatas B Where A.����ID=B.����ID And A.����=B.����);
	Delete From zlTools.zlRPTConds A Where RowID<(Select Max(RowID) From zlTools.zlRPTConds B Where A.����ID=B.����ID And A.������=B.������);
	Delete From zlTools.zlRPTConds A Where RowID<(Select Max(RowID) From zlTools.zlRPTConds B Where A.����ID=B.����ID And A.��������=B.��������);
End;
/
ALTER TABLE zlTools.zlRPTDatas ADD CONSTRAINT zlRPTDatas_UQ_���� UNIQUE (����ID,����) USING INDEX PCTFREE 5
/
ALTER TABLE zlTools.zlRPTConds ADD CONSTRAINT zlRPTConds_PK PRIMARY KEY (����ID,������)
/
ALTER TABLE zlTools.zlRPTConds ADD CONSTRAINT zlRPTConds_UQ_�������� UNIQUE (����ID,��������) USING INDEX PCTFREE 5
/

-- ������������9i
-- ɾ����ʹ�õİ�
drop package b_datamana
/

--�¸���,9040
Insert Into zlOptions(������,������,����ֵ,ȱʡֵ,����˵��) Values(7,'���ѷ������',';9999;0',';9999;0','�������ѷ���ķ����������˿ںż�״̬����Ϣ��')
/

-----------------------------------------------------
-- ������ͷ 2006-8-24, 15:41:11 --
-----------------------------------------------------
CREATE OR REPLACE PACKAGE ZLTOOLS.b_Expert IS
  -----------------------------------------------------------------------------
  -- ���ߣ� �¶�
  -- ��ʼʱ�䣺2006-6-29
  -- �޸��ˣ�
  -- �޸�ʱ�䣺
  -- ������
  --         ��Ҫ����ר��ߵĹ���
  -----------------------------------------------------------------------------

  TYPE t_Refcur IS REF CURSOR;

  -----------------------------------------------------------------------------
  -- ȡ��������
  -- �����б� frmNoticesEdit.ReadData��frmNoticeTools.cboSystem_Click
  -----------------------------------------------------------------------------
  PROCEDURE Get_Notices
  (
    Cursor_Out OUT t_Refcur,
    ���_In    IN Zlnotices.���%TYPE,
    ϵͳ_In    IN Zlreports.ϵͳ%TYPE := NULL
  );

  -----------------------------------------------------------------------------
  -- ȡ���Ѷ�������
  -- �����б� frmNoticesEdit.ReadData
  -----------------------------------------------------------------------------
  PROCEDURE Get_Noticeusr
  (
    Cursor_Out  OUT t_Refcur,
    ���Ѷ���_In IN Zlnoticeusr.���Ѷ���%TYPE,
    �������_In IN Zlnoticeusr.�������%TYPE
  );

  -----------------------------------------------------------------------------
  -- ȡ����ѡ������ѱ���
  -- �����б� frmNoticesEdit.cmdOpen_Click
  -----------------------------------------------------------------------------
  PROCEDURE Get_Noticereport
  (
    Cursor_Out OUT t_Refcur,
    ϵͳ_In    IN Zlreports.ϵͳ%TYPE
  );

  -----------------------------------------------------------------------------
  -- �ڲ�ͬ��ϵͳ�临�Ʊ���
  -- �����б�mdlMain.CopyReport
  -----------------------------------------------------------------------------
  PROCEDURE Copy_Report
  (
    ϵͳ_In   IN Zlreports.ϵͳ%TYPE,
    ��ϵͳ_In IN Zlreports.ϵͳ%TYPE
  );

END b_Expert;
/

CREATE OR REPLACE Package ZLTOOLS.b_Loadandunload Is
  -----------------------------------------------------------------------------
  -- ���ߣ� �¶�
  -- ��ʼʱ�䣺2006-6-29
  -- �޸��ˣ�
  -- �޸�ʱ�䣺
  -- ������
  --         ��Ҫ����װж����Ĺ���
  -----------------------------------------------------------------------------

  Type t_Refcur Is Ref Cursor;

   -----------------------------------------------------------------------------
  -- ���ܣ�ȡ��SysFiles����ļ���
  -- �����б�frmAppCheck.cmbSystem_Click��frmClearData.cmbSystem_Click��frmAppScript.cmbSystem_Click
  --           frmAppUpgrade.cmbSystem_Click
  -----------------------------------------------------------------------------
  Procedure Get_Sysfile_Name
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In zlSysFiles.ϵͳ%Type,
    ����_In    In zlSysFiles.����%Type
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ����ϵͳ����
  -- �����б� frmAppStart.cmdFunction_MouseUp
  -----------------------------------------------------------------------------
  Procedure Get_Share_Name
  (
    Cursor_Out Out t_Refcur,
    �����_In  In zlSystems.���%Type
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡOracle�汾��
  -- �����б� frmAppStart.Form_Load��frmImp.Form_Load��frmLoadIn.Form_Load
  -----------------------------------------------------------------------------
  Procedure Get_Oracle_Ver(Cursor_Out Out t_Refcur);
End b_Loadandunload;
/

CREATE OR REPLACE Package ZLTOOLS.b_Popedom Is

  -----------------------------------------------------------------------------
  -- ���ߣ� �¶�
  -- ��ʼʱ�䣺2006-6-29
  -- �޸��ˣ�
  -- �޸�ʱ�䣺
  -- ������
  --         ��Ҫ����Ȩ�޹���Ĺ���
  -----------------------------------------------------------------------------
  Type t_Refcur Is Ref Cursor;

  -----------------------------------------------------------------------------
  -- ���ܣ�CopyMenu
  -- �����б�mdlMain.CopyMenu
  -----------------------------------------------------------------------------
  Procedure Copy_Menu
  (
    ϵͳ_In   In zlMenus.ϵͳ%Type,
    ��ϵͳ_In In zlMenus.ϵͳ%Type
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡZlMenu����
  -- �����б� frmMenu.cmdExp_Click��frmMenu.FillMenu
  -----------------------------------------------------------------------------
  Procedure Get_Menu_Tree
  (
    Cursor_Out Out t_Refcur,
    ���_In    In zlMenus.���%Type
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡZlMenu����
  -- �����б� frmMenu.cmdNew_Click��frmMenu.FillMenuName
  -----------------------------------------------------------------------------
  Procedure Get_Menu_Group
  (
    Cursor_Out Out t_Refcur,
    ���_In    In zlMenus.���%Type
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡģ��
  -- �����б� frmProgPriv.Fillģ��
  -----------------------------------------------------------------------------
  Procedure Get_Module
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In zlComponent.ϵͳ%Type
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ���ܻ����У�˵��
  -- �����б� frmProgPriv.Fill���ܡ�frmProgPriv.Fill��Ȩ��
  -----------------------------------------------------------------------------
  Procedure Get_Function
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In zlProgFuncs.ϵͳ%Type,
    ���_In    In zlProgFuncs.���%Type,
    ����_In    In zlProgFuncs.����%Type := Null
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ��Ȩ��
  -- �����б� frmProgPriv.Fill��Ȩ��
  -----------------------------------------------------------------------------
  Procedure Get_Impower
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In zlProgPrivs.ϵͳ%Type,
    ���_In    In zlProgPrivs.���%Type,
    ����_In    In zlProgPrivs.����%Type
  );

  -----------------------------------------------------------------------------
  -- ���ܣ��õ���ɫ�ܷ��ʵĵ���̨����
  -- �����б� frmRole.FillModule
  -----------------------------------------------------------------------------
  Procedure Get_Role_Tools
  (
    Cursor_Out Out t_Refcur,
    ��ɫ_In    In zlRoleGrant.��ɫ%Type
  );

  -----------------------------------------------------------------------------
  -- ���ܣ��õ���ǰ��Ȩ��
  -- �����б� frmRoleGrant.cmdOK_Click
  -----------------------------------------------------------------------------
  Procedure Get_Role_Grant
  (
    Curgrand_Out    Out t_Refcur,
    Curprivs_Out    Out t_Refcur,
    Curfuncpars_Out Out t_Refcur,
    ��ɫ_In         In zlRoleGrant.��ɫ%Type
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�FillFunc
  -- �����б� frmRoleGrant.FillFunc
  -----------------------------------------------------------------------------
  Procedure Get_Zlprogfunc
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In zlProgFuncs.ϵͳ%Type,
    ���_In    In zlProgFuncs.���%Type
  );

  -----------------------------------------------------------------------------
  -- ���ܣ������н�ɫ��Ӧ��ģ��
  -- �����б� frmUserEdit.UserEdit
  -----------------------------------------------------------------------------
  Procedure Get_All_Module(Cursor_Out Out t_Refcur);

End b_Popedom;
/

CREATE OR REPLACE Package ZLTOOLS.b_Public Is
  -----------------------------------------------------------------------------
  -- ���ߣ� �¶�
  -- ��ʼʱ�䣺2006-6-29
  -- �޸��ˣ�
  -- �޸�ʱ�䣺
  -- ������
  --         ��������
  -----------------------------------------------------------------------------

  Type t_Refcur Is Ref Cursor;

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡϵͳ����
  -- �����б�
  -- mdlMain.CurrentDate
  -- clsDatabase.CurrentDate
  -----------------------------------------------------------------------------
  Procedure Get_Current_Date(Cursor_Out Out t_Refcur);

  -----------------------------------------------------------------------------
  -- ���ܣ�ɾ��������־��������־
  -- �����б�
  -- mdlMain.DeleteAllLog
  -----------------------------------------------------------------------------
  Procedure Delete_All_Log(Runtimelog_In In Number := 0);

  -----------------------------------------------------------------------------
  -- ���ܣ�ɾ����ǰ������־
  -- �����б�
  -- mdlMain.DeleteCurLog
  -----------------------------------------------------------------------------
  Procedure Delete_Diarylog
  (
    �Ự��_In   Number,
    �û���_In   Varchar2,
    ����վ_In   Varchar2,
    ������_In   Varchar2,
    ��������_In Varchar2,
    ����ʱ��_In Date
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�ɾ����ǰ������־
  -- �����б�
  -- mdlMain.DeleteCurLog
  -----------------------------------------------------------------------------
  Procedure Delete_Errorlog
  (
    �Ự��_In   Number,
    �û���_In   Varchar2,
    ����վ_In   Varchar2,
    ����_In     Number,
    �������_In Number,
    ʱ��_In     Date
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡע����
  -- �����б�
  -- mdlMain.Getע����
  -----------------------------------------------------------------------------
  Procedure Get_Regcode(Cursor_Out Out t_Refcur);

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ�汾��
  -- �����б�
  -- mdlMain.UpgradeManager
  -----------------------------------------------------------------------------
  Procedure Get_Ver(Cursor_Out Out t_Refcur);

  -----------------------------------------------------------------------------
  -- ���ܣ����°汾��
  -- �����б�
  -- mdlMain.UpgradeManager
  -----------------------------------------------------------------------------
  Procedure Update_Ver(Verstring_In In Varchar2);

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ��ϵͳ����������
  -- �����б�
  -- frmStatus.cmbsystem_Click��mdlMain.GetOwnerName��mdlMain.cmbSystem_Click
  -- frmAutoJobs.cmbSystem_Click��frmDataMove.cmbSystem_Click ��frmNoticeTools.cboSystem_Click
  -- frmProgPriv.ProgPriv��frmAppScript.cmbSystem_Click
  -----------------------------------------------------------------------------
  Procedure Get_Owner_Name
  (
    Cursor_Out Out t_Refcur,
    ���_In    In zlSystems.���%Type := 0
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡע�������Ϣ
  -- �����б�
  -- frmAbout.GetUnitInfo��frmAutoJobs.From_load��frmClientsUpgrade.InitInfor
  -- frmFilesSet.ShowEdit��frmRegist.From_load��frmAppScript.From_Load
  -- frmFilesSendToServer.InitInfo
  -----------------------------------------------------------------------------
  Procedure Get_Reginfo
  (
    Cursor_Out Out t_Refcur,
    ��Ŀ_In    In zlRegInfo.��Ŀ%Type := Null
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡzlGetSvrToolsg����
  -- �����б�
  -- frmMDIMain.MDIForm_Load
  -----------------------------------------------------------------------------
  Procedure Get_Zlsvrtools(Cursor_Out Out t_Refcur);

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ�Ѱ�װϵͳ�嵥
  -- �����б�
  -- frmAppCheck.Form_Load��frmClearData.Form_Load��frmDataMove.Form_Load
  -- frmImp.FillSystem��frmLoadIn.FillSystem��frmLoadOut.FillSystem
  -- frmMDIMain.mnuFileRemove_Click��frmNoticeTools.Form_Activate��frmRoleGrant.FillSystem
  -- frmAppUpgrade.Form_Load��frmAppScript.Form_Load��frmExp.FillSystem
  -- frmInputTools.from_activate��fromRole.FillSystem��frmAutoJobs.From_load
  -- frmAppstart.sysCreated
  -----------------------------------------------------------------------------
  Procedure Get_Zlsystems
  (
    Cursor_Out Out t_Refcur,
    ������_In  In zlSystems.������%Type := Null
  );

End b_Public;
/

CREATE OR REPLACE Package ZLTOOLS.b_Runmana Is
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
-------------------------------------------------
-- ��������
-------------------------------------------------

CREATE OR REPLACE PACKAGE BODY ZLTOOLS.b_Expert IS

  -----------------------------------------------------------------------------
  -- ȡ��������
  -----------------------------------------------------------------------------
  PROCEDURE Get_Notices
  (
    Cursor_Out OUT t_Refcur,
    ���_In    IN Zlnotices.���%TYPE,
    ϵͳ_In    IN Zlreports.ϵͳ%TYPE := NULL
  ) IS
  BEGIN
    IF Nvl(���_In, 0) <> 0 THEN
      -- frmNoticesEdit.ReadData ʹ��
      OPEN Cursor_Out FOR
        SELECT a.��������, a.��������, a.���ѱ���, a.��������, a.���Ѵ���, a.��ʼʱ��, a.��ֹʱ��, a.�������,
               b.���� AS ��������
        FROM Zlnotices a, Zlreports b
        WHERE a.���ѱ��� = b.���(+) AND a.��� = ���_In;
    ELSE
      -- cboSystem_Click ʹ��
      IF Nvl(ϵͳ_In, 0) = 0 THEN
        OPEN Cursor_Out FOR
          SELECT a.���, a.��������, a.��������, a.���ѱ���, a.��������, a.���Ѵ���, a.��ʼʱ��, a.��ֹʱ��, a.�������,
                 a.��������, b.���� AS ��������
          FROM Zlnotices a, Zlreports b
          WHERE a.���ѱ��� = b.���(+) AND a.ϵͳ IS NULL;
      ELSE
        OPEN Cursor_Out FOR
          SELECT a.���, a.��������, a.��������, a.���ѱ���, a.��������, a.���Ѵ���, a.��ʼʱ��, a.��ֹʱ��, a.�������,
                 a.��������, b.���� AS ��������
          FROM Zlnotices a, Zlreports b
          WHERE a.���ѱ��� = b.���(+) AND a.ϵͳ = ϵͳ_In;
      END IF;
    END IF;

  END Get_Notices;

  -----------------------------------------------------------------------------
  -- ȡ���Ѷ�������
  -----------------------------------------------------------------------------
  PROCEDURE Get_Noticeusr
  (
    Cursor_Out  OUT t_Refcur,
    ���Ѷ���_In IN Zlnoticeusr.���Ѷ���%TYPE,
    �������_In IN Zlnoticeusr.�������%TYPE
  ) IS
  BEGIN
    IF Nvl(���Ѷ���_In, 0) = 0 THEN
      OPEN Cursor_Out FOR
        SELECT 1 FROM Zlnoticeusr WHERE Rownum < 2 AND ������� = �������_In;
    ELSE
      OPEN Cursor_Out FOR
        SELECT �������� FROM Zlnoticeusr WHERE ���Ѷ��� = ���Ѷ���_In AND ������� = �������_In;
    END IF;
  END Get_Noticeusr;

  -----------------------------------------------------------------------------
  -- ȡ����ѡ������ѱ���
  -----------------------------------------------------------------------------
  PROCEDURE Get_Noticereport
  (
    Cursor_Out OUT t_Refcur,
    ϵͳ_In    IN Zlreports.ϵͳ%TYPE
  ) IS
  BEGIN
    IF Nvl(ϵͳ_In, 0) = 0 THEN
      OPEN Cursor_Out FOR
        SELECT Id, ���, ����, ˵��
        FROM Zlreports
        WHERE ��� LIKE 'ZL%_REPORT_%' AND
              NOT (����ʱ�� IS NULL OR Trunc(����ʱ��) = To_Date('3000-01-01', 'yyyy-mm-dd')) AND Nvl(ϵͳ, 0) = 0;
    ELSE
      OPEN Cursor_Out FOR
        SELECT Id, ���, ����, ˵��
        FROM Zlreports
        WHERE ��� LIKE 'ZL%_REPORT_%' AND
              NOT (����ʱ�� IS NULL OR Trunc(����ʱ��) = To_Date('3000-01-01', 'yyyy-mm-dd')) AND ϵͳ = ϵͳ_In;
    END IF;
  END Get_Noticereport;

  -----------------------------------------------------------------------------
  -- �ڲ�ͬ��ϵͳ�临�Ʊ���
  -----------------------------------------------------------------------------
  PROCEDURE Copy_Report
  (
    ϵͳ_In   IN Zlreports.ϵͳ%TYPE,
    ��ϵͳ_In IN Zlreports.ϵͳ%TYPE
  ) IS
    n_Grpid   NUMBER;
    n_Rptid   NUMBER;
    n_Dataid  NUMBER;
    n_Itemid  NUMBER;
    v_Olduser VARCHAR2(100);
    v_Newuser VARCHAR2(100);

    FUNCTION Sub_Owner_Name(Lngsys_In IN NUMBER := 0) RETURN VARCHAR2 IS
      v_Owner_Name VARCHAR2(30);
    BEGIN
      SELECT Upper(������) AS ������ INTO v_Owner_Name FROM Zlsystems WHERE ��� = Lngsys_In;
      RETURN v_Owner_Name;
    END Sub_Owner_Name;

  BEGIN
    SELECT Nvl(MAX(Id), 0) INTO n_Grpid FROM Zlrptgroups;
    SELECT Nvl(MAX(Id), 0) INTO n_Rptid FROM Zlreports;
    SELECT Nvl(MAX(Id), 0) INTO n_Dataid FROM Zlrptdatas;
    SELECT Nvl(MAX(Id), 0) INTO n_Itemid FROM Zlrptitems;
    n_Grpid  := n_Grpid + 1;
    n_Rptid  := n_Rptid + 1;
    n_Dataid := n_Dataid + 1;
    n_Itemid := n_Itemid + 1;

    v_Olduser := Upper(Sub_Owner_Name(ϵͳ_In));
    v_Newuser := Upper(Sub_Owner_Name(��ϵͳ_In));

    INSERT INTO Zlrptgroups
      (Id, ���, ����, ˵��, ϵͳ, ����id, ����ʱ��)
      SELECT Id + n_Grpid, ���, ����, ˵��, ��ϵͳ_In, ����id, ����ʱ�� FROM Zlrptgroups WHERE ϵͳ = ϵͳ_In;

    INSERT INTO Zlreports
      (Id, ���, ����, ˵��, ����, w, h, ֽ��, ֽ��, ��ֽ, ��ӡ��, Ʊ��, ϵͳ, ����id, ����, �޸�ʱ��, ����ʱ��)
      SELECT Id + n_Rptid, ���, ����, ˵��, ����, w, h, ֽ��, ֽ��, ��ֽ, ��ӡ��, Ʊ��, ��ϵͳ_In, ����id, ����,
             �޸�ʱ��, ����ʱ��
      FROM Zlreports
      WHERE ϵͳ = ϵͳ_In;

    -- ����zlRPTSub
    INSERT INTO Zlrptsubs
      (��id, ����id, ���, ����)
      SELECT a.��id + n_Grpid, a.����id + n_Rptid, a.���, a.����
      FROM Zlrptsubs a, Zlrptgroups b
      WHERE a.��id = b.Id AND b.ϵͳ = ϵͳ_In;

    -- ����zlRPTFMTs
    INSERT INTO Zlrptfmts
      (����id, ���, ˵��, ͼ��)
      SELECT a.����id + n_Rptid, a.���, a.˵��, a.ͼ��
      FROM Zlrptfmts a, Zlreports b
      WHERE a.����id = b.Id AND b.ϵͳ = ϵͳ_In;

    -- ����zlRPTItems
    INSERT INTO Zlrptitems
      (Id, ����id, ��ʽ��, ����, ����, �ϼ�id, ���, ����, ����, ����, ��ͷ, x, y, w, h, �и�, ����, �Ե�, ����, �ֺ�,
       ����, б��, ����, ǰ��, ����, �߿�, ����, ��ʽ, ����, ����, ����, ϵͳ)
      SELECT a.Id + n_Itemid, a.����id + n_Rptid, a.��ʽ��, a.����, a.����, a.�ϼ�id + n_Itemid, a.���, a.����, a.����,
             a.����, a.��ͷ, a.x, a.y, a.w, a.h, a.�и�, a.����, a.�Ե�, a.����, a.�ֺ�, a.����, a.б��, a.����, a.ǰ��,
             a.����, a.�߿�, a.����, a.��ʽ, a.����, a.����, a.����, a.ϵͳ
      FROM Zlrptitems a, Zlreports b
      WHERE a.����id = b.Id AND b.ϵͳ = ϵͳ_In;
    -- ����zlRptDatas
    INSERT INTO Zlrptdatas
      (Id, ����id, ����, �ֶ�, ����, ����)
      SELECT a.Id + n_Dataid, a.����id + n_Rptid, a.����, a.�ֶ�, a.����, a.����
      FROM Zlrptdatas a, Zlreports b
      WHERE a.����id = b.Id AND b.ϵͳ = ϵͳ_In;
    -- ����zlRPTSqls
    INSERT INTO Zlrptsqls
      (Դid, �к�, ����)
      SELECT a.Դid + n_Dataid, a.�к�, a.����
      FROM Zlrptsqls a, Zlrptdatas b, Zlreports c
      WHERE a.Դid = b.Id AND b.����id = c.Id AND c.ϵͳ = ϵͳ_In;

    -- ����zlRPTPars
    INSERT INTO Zlrptpars
      (Դid, ����, ���, ����, ����, ȱʡֵ, ��ʽ, ֵ�б�, ����sql, ��ϸsql, �����ֶ�, ��ϸ�ֶ�, ����)
      SELECT a.Դid + n_Dataid, a.����, a.���, a.����, a.����, a.ȱʡֵ, a.��ʽ, a.ֵ�б�, a.����sql, a.��ϸsql,
             a.�����ֶ�, a.��ϸ�ֶ�, a.����
      FROM Zlrptpars a, Zlrptdatas b, Zlreports c
      WHERE a.Դid = b.Id AND b.����id = c.Id AND c.ϵͳ = ϵͳ_In;

    -- zlFunctions����
    INSERT INTO Zlfunctions
      (ϵͳ, ������, ������, ������, ˵��)
      SELECT ��ϵͳ_In, ������, ������, ������, ˵�� FROM Zlfunctions WHERE ϵͳ = ϵͳ_In;

    -- zlFuncPars����
    INSERT INTO Zlfuncpars
      (ϵͳ, ������, ������, ������, ������, ����, ȱʡֵ, ֵ�б�, ����sql, ��ϸsql, �����ֶ�, ��ϸ�ֶ�, ����, ����,
       ������)
      SELECT ��ϵͳ_In, ������, ������, ������, ������, ����, ȱʡֵ, ֵ�б�, ����sql, ��ϸsql, �����ֶ�, ��ϸ�ֶ�, ����,
             ����, ������
      FROM Zlfuncpars
      WHERE ϵͳ = ϵͳ_In;

    -- ������������Դ����
    UPDATE Zlrptdatas
    SET ���� = REPLACE(����, v_Olduser || '.', v_Newuser || '.')
    WHERE Id IN (SELECT a.Id FROM Zlrptdatas a, Zlreports b WHERE a.����id = b.Id AND b.ϵͳ = ��ϵͳ_In);

    UPDATE Zlrptpars
    SET ���� = REPLACE(����, v_Olduser || '.', v_Newuser || '.')
    WHERE Դid IN (SELECT a.Id FROM Zlrptdatas a, Zlreports b WHERE a.����id = b.Id AND b.ϵͳ = ��ϵͳ_In);

    UPDATE Zlfuncpars SET ���� = REPLACE(����, v_Olduser || '.', v_Newuser || '.') WHERE ϵͳ = ��ϵͳ_In;

    COMMIT;
  EXCEPTION
    WHEN OTHERS THEN
      Zl_Errorcenter(SQLCODE, SQLERRM);
  END Copy_Report;

END b_Expert;
/

CREATE OR REPLACE Package Body ZLTOOLS.b_Loadandunload Is

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ��SysFiles����ļ���
  -----------------------------------------------------------------------------
  Procedure Get_Sysfile_Name
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In zlSysFiles.ϵͳ%Type,
    ����_In    In zlSysFiles.����%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select �ļ��� From zlSysFiles Where ϵͳ = ϵͳ_In And ���� = ����_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Sysfile_Name;

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ����ϵͳ����
  -----------------------------------------------------------------------------
  Procedure Get_Share_Name
  (
    Cursor_Out Out t_Refcur,
    �����_In  In zlSystems.���%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select ���� From zlSystems Start With ����� = �����_In Connect By Prior ��� = �����;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Share_Name;

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡOracle�汾��
  -----------------------------------------------------------------------------
  Procedure Get_Oracle_Ver(Cursor_Out Out t_Refcur) Is
  Begin
    Open Cursor_Out For
      Select To_Number(Replace(Substr(Banner, 6, 4), '.', '')) As Oraclever
      From V$version
      Where Substr(Banner, 1, 4) = 'CORE';
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Oracle_Ver;
End b_Loadandunload;
/

CREATE OR REPLACE Package Body ZLTOOLS.b_Popedom Is

  -----------------------------------------------------------------------------
  -- ���ܣ�CopyMenu
  -----------------------------------------------------------------------------
  Procedure Copy_Menu
  (
    ϵͳ_In   In zlMenus.ϵͳ%Type,
    ��ϵͳ_In In zlMenus.ϵͳ%Type
  ) Is
    n_Menuid zlMenus.ID%Type;
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
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Copy_Menu;

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡZlMenu����
  -----------------------------------------------------------------------------
  Procedure Get_Menu_Tree
  (
    Cursor_Out Out t_Refcur,
    ���_In    In zlMenus.���%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select ID, �ϼ�id, ����, ���, ˵��, ϵͳ, ģ��, �̱���, ͼ��, Level As ����
      From zlMenus
      Start With �ϼ�id Is Null And ��� = ���_In
      Connect By Prior ID = �ϼ�id And ��� = ���_In
      Order By Level, ID;
  End Get_Menu_Tree;

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡZlMenu����
  -----------------------------------------------------------------------------
  Procedure Get_Menu_Group
  (
    Cursor_Out Out t_Refcur,
    ���_In    In zlMenus.���%Type
  ) Is
  Begin
    If ���_In Is Null Then
      -- frmMenu.FillMenuName
      Open Cursor_Out For
        Select Distinct ��� From zlMenus;
    Else
      -- frmMenu.cmdNew_Click
      Open Cursor_Out For
        Select Count(*) As ���� From zlMenus Where ��� = ���_In;
    End If;
  End Get_Menu_Group;

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡģ��
  -----------------------------------------------------------------------------
  Procedure Get_Module
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In zlComponent.ϵͳ%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select P.���, P.����, C.���� As ����
      From zlPrograms P, zlComponent C
      Where Upper(P.����) = Upper(C.����) And C.ϵͳ = ϵͳ_In And P.ϵͳ = ϵͳ_In
      Order By C.����, P.���;
  End Get_Module;

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ���ܻ����У�˵��
  -----------------------------------------------------------------------------
  Procedure Get_Function
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In zlProgFuncs.ϵͳ%Type,
    ���_In    In zlProgFuncs.���%Type,
    ����_In    In zlProgFuncs.����%Type := Null
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

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ��Ȩ��
  -----------------------------------------------------------------------------
  Procedure Get_Impower
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In zlProgPrivs.ϵͳ%Type,
    ���_In    In zlProgPrivs.���%Type,
    ����_In    In zlProgPrivs.����%Type
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

  -----------------------------------------------------------------------------
  -- ���ܣ��õ���ɫ�ܷ��ʵĵ���̨����
  -----------------------------------------------------------------------------
  Procedure Get_Role_Tools
  (
    Cursor_Out Out t_Refcur,
    ��ɫ_In    In zlRoleGrant.��ɫ%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select P.���, P.����, P.˵��, R.����
      From zlRoleGrant R, zlPrograms P
      Where R.ϵͳ Is Null And P.��� = R.��� And R.��ɫ = ��ɫ_In And P.ϵͳ Is Null And P.��� < 100 And
            P.���� Is Null
      Order By P.���;
  End Get_Role_Tools;

  -----------------------------------------------------------------------------
  -- ���ܣ��õ���ǰ��Ȩ��
  -----------------------------------------------------------------------------
  Procedure Get_Role_Grant
  (
    Curgrand_Out    Out t_Refcur,
    Curprivs_Out    Out t_Refcur,
    Curfuncpars_Out Out t_Refcur,
    ��ɫ_In         In zlRoleGrant.��ɫ%Type
  ) Is
  Begin
    Open Curgrand_Out For
      Select Nvl(ϵͳ, 0) As ϵͳ, ���, ���� From zlRoleGrant Where ��ɫ = ��ɫ_In;
    Open Curprivs_Out For
      Select Nvl(ϵͳ, 0) As ϵͳ, ���, ����, ������, Ȩ��, ���� From zlProgPrivs;
    Open Curfuncpars_Out For
      Select P.ϵͳ, F.������, P.����
      From zlFuncPars P, zlFunctions F
      Where P.ϵͳ = F.ϵͳ And P.������ = F.������ And P.���� Is Not Null;
  End Get_Role_Grant;

  -----------------------------------------------------------------------------
  -- ���ܣ�FillFunc
  -----------------------------------------------------------------------------
  Procedure Get_Zlprogfunc
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In zlProgFuncs.ϵͳ%Type,
    ���_In    In zlProgFuncs.���%Type
  ) Is
  Begin
    If Nvl(ϵͳ_In, 0) = 0 Then
      Open Cursor_Out For
        Select ����, ����, ˵�� From zlProgFuncs Where ϵͳ Is Null And ��� = ���_In And ���� <> '����';
    Else
      Open Cursor_Out For
        Select A.����, A.����, A.˵��
        From zlProgFuncs A, Zlregfunc B
        Where (A.ϵͳ / 100) = B.ϵͳ And A.��� = B.��� And A.���� = B.���� And A.ϵͳ = ϵͳ_In And A.��� = ���_In And
              A.���� <> '����';
    End If;
  End Get_Zlprogfunc;

  -----------------------------------------------------------------------------
  -- ���ܣ������н�ɫ��Ӧ��ģ��
  -----------------------------------------------------------------------------
  Procedure Get_All_Module(Cursor_Out Out t_Refcur) Is
  Begin
    Open Cursor_Out For
      Select A.��ɫ, A.���, A.����, B.����, B.˵��
      From zlRoleGrant A, zlPrograms B
      Where A.��� = B.��� And Nvl(A.ϵͳ, 0) = Nvl(B.ϵͳ, 0)
      Order By A.��ɫ, A.���;
  End Get_All_Module;

End b_Popedom;
/

CREATE OR REPLACE Package Body ZLTOOLS.b_Public Is

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡϵͳ����
  -----------------------------------------------------------------------------
  Procedure Get_Current_Date(Cursor_Out Out t_Refcur) Is
  Begin
    Open Cursor_Out For
      Select Sysdate As ���� From Dual;
  End Get_Current_Date;

  -----------------------------------------------------------------------------
  -- ���ܣ�ɾ��������־��������־
  -----------------------------------------------------------------------------
  Procedure Delete_All_Log(Runtimelog_In In Number := 0) Is
    n_Count Number;
    n_Loop  Number;
  Begin
    If Runtimelog_In = 1 Then
      Select Count(����ʱ��) Into n_Count From zlDiaryLog;
      If n_Count > 1000 Then
        For n_Loop In 1 .. Ceil(n_Count - 1000) Loop
          Delete zlDiaryLog Where Rownum < 10001;
          Commit;
        End Loop;
      Else
        If n_Count > 0 Then
          Delete zlDiaryLog;
          Commit;
        End If;
      End If;
    Else
      Select Count(ʱ��) Into n_Count From zlErrorLog;
      If n_Count > 1000 Then
        For n_Loop In 1 .. Ceil(n_Count - 1000) Loop
          Delete zlErrorLog Where Rownum < 10001;
          Commit;
        End Loop;
      Else
        If n_Count > 0 Then
          Delete zlErrorLog;
          Commit;
        End If;
      End If;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Delete_All_Log;

  -----------------------------------------------------------------------------
  -- ���ܣ�ɾ����ǰ������־
  -----------------------------------------------------------------------------
  Procedure Delete_Diarylog
  (
    �Ự��_In   Number,
    �û���_In   Varchar2,
    ����վ_In   Varchar2,
    ������_In   Varchar2,
    ��������_In Varchar2,
    ����ʱ��_In Date
  ) Is
  Begin
    Delete zlDiaryLog
    Where �Ự�� = �Ự��_In And �û��� = �û���_In And ����վ = ����վ_In And ������ = ������_In And
          �������� = ��������_In And ����ʱ�� = ����ʱ��_In;
    Commit;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Delete_Diarylog;

  -----------------------------------------------------------------------------
  -- ���ܣ�ɾ����ǰ������־
  -----------------------------------------------------------------------------
  Procedure Delete_Errorlog
  (
    �Ự��_In   Number,
    �û���_In   Varchar2,
    ����վ_In   Varchar2,
    ����_In     Number,
    �������_In Number,
    ʱ��_In     Date
  ) Is
  Begin
    Delete zlErrorLog
    Where �Ự�� = �Ự��_In And �û��� = �û���_In And ����վ = ����վ_In And ���� = ����_In And
          ������� = �������_In And ʱ�� = ʱ��_In;
    Commit;
  
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Delete_Errorlog;

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡע����
  -----------------------------------------------------------------------------
  Procedure Get_Regcode(Cursor_Out Out t_Refcur) Is
  Begin
    Open Cursor_Out For
      Select ���� From zlRegInfo Where ��Ŀ = 'ע����' Or ��Ŀ = '��Ȩ֤��' Order By �к�;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Regcode;

    -----------------------------------------------------------------------------
  -- ���ܣ�ȡ�汾��
  -----------------------------------------------------------------------------
  Procedure Get_Ver(Cursor_Out Out t_Refcur) Is
  Begin
    Open Cursor_Out For
      Select ���� From zlRegInfo Where ��Ŀ = '�汾��';
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Ver;

  -----------------------------------------------------------------------------
  -- ���ܣ����°汾��
  -----------------------------------------------------------------------------
  Procedure Update_Ver(Verstring_In In Varchar2) Is
  Begin
    Update zlRegInfo Set ���� = Verstring_In Where ��Ŀ = '�汾��';
    If Sql%NotFound Then
      Insert Into zlRegInfo (��Ŀ, �к�, ����) Values ('�汾��', 1, Verstring_In);
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Update_Ver;

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ��ϵͳ����������
  -----------------------------------------------------------------------------
  Procedure Get_Owner_Name
  (
    Cursor_Out Out t_Refcur,
    ���_In    In zlSystems.���%Type := 0
  ) Is
  Begin
    Open Cursor_Out For
      Select Upper(������) As ������ From zlSystems Where ��� = ���_In;
  End Get_Owner_Name;

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡע�������Ϣ
  -----------------------------------------------------------------------------
  Procedure Get_Reginfo
  (
    Cursor_Out Out t_Refcur,
    ��Ŀ_In    In zlRegInfo.��Ŀ%Type := Null
  ) Is
  Begin
    If Trim(Nvl(��Ŀ_In, '��')) = '��' Then
      Open Cursor_Out For
        Select * From zlRegInfo;
    Else
      Open Cursor_Out For
        Select ���� From zlRegInfo Where ��Ŀ = ��Ŀ_In Order By �к�;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Reginfo;

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡzlGetSvrToolsg����
  -----------------------------------------------------------------------------
  Procedure Get_Zlsvrtools(Cursor_Out Out t_Refcur) Is
  Begin
    Open Cursor_Out For
      Select * From zlSvrTools Start With �ϼ� Is Null Connect By Prior ��� = �ϼ� Order By Level, ���;
  End Get_Zlsvrtools;

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ�Ѱ�װϵͳ�嵥
  -----------------------------------------------------------------------------
  Procedure Get_Zlsystems
  (
    Cursor_Out Out t_Refcur,
    ������_In  In zlSystems.������%Type := Null
  ) Is
  Begin
    If Nvl(������_In, '��') = '��' Then
      Open Cursor_Out For
        Select * From zlSystems Order By ���;
    Else
      Open Cursor_Out For
        Select * From zlSystems Where Upper(������) = Upper(������_In) Order By ���;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Zlsystems;

End b_Public;
/

CREATE OR REPLACE Package Body ZLTOOLS.b_Runmana Is
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
							 a.������, Decode(b.Terminal, Null, 0, 1) As ״̬, a.�ռ���־
				From Zlclients a, (Select Distinct Terminal From V$session) b
				Where Upper(a.����վ) = Upper(b.Terminal(+))
				Order By a.Ip';
      Open Cur_Out For v_Sql;
    Else
      Open Cur_Out For
        Select Ip, ����վ, Cpu, �ڴ�, Ӳ��, ����ϵͳ, ����, ��;, ˵��, ������־, ��ֹʹ��, ������
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

--8735
Create Or Replace Function zltools.zlSpellCode(v_Instr In Varchar2,v_OutNum In Integer:=10)
	Return Varchar2 Is
	v_Spell   Varchar2(40);
	v_Input   Varchar2(1000);
	v_Bitchar Varchar2(2);
	v_Bitnum  Integer;
	v_Chrnum  Integer;
  v_OutMaxNum Integer;
	v_Stdstr  Varchar2(50) := '�Ų���귢������-��������Ŷž��Ȼ������-����ѹ��';
	v_Chara   Varchar2(2000) := '߹��H�����X�����t�H�c���\���������ٌ���r�P�o�a�������Y��O��c�g��������@���@���t؁�B���l���E�J�������֒�������q�������O���֓��������';
	v_Charb   Varchar2(2000) := '�����^���������R�T�������Zڕ���E��������ٔ�v�C�n���������k����[��k�ߙ�D���r�^����������_�����R�dم�d����E�s�U�t���l�mؐ����f�K�����R݅�^�������G�S��Q����ݙ�a����J�M��G�a�P���q�S�sݩذ����������؄םߛ���������������[�]����C�P�����G�s�����޵�I�ۋ�@��������@�S�{�K�L�x�F�E�z�������߄�Q��߅�b�c���H���u��������������r�g�Y�l�m�p�q׃�������R���������[րٙ�S���j�k�l�n�s�Օ�l���B�M������ߓ�������h�e�f�����E�\����������W������v����u���V���@�m�h����\�G�@�Q���R������K�����c�D�N�m�n�P���}��������L��������c�J�K��߲�G�L�Q����b�Y�^����������������������������';
	v_Charc   Varchar2(2000) := '�����nؔ�P�{�����o���֍���]�I�ى�������[ܳ������������������\��x���ۂ���d������O٭�{���i������S�a�������������K׋������~����p֝�P�U׀�����]����������_����K�����L�M��������L�l����������k�o�����n���}���C�{֚��܇�����ފ�J�����Հ�o�nލ����ܕ�\��R���mڒ���{ٕ���f��Y�Z��׏�p�����W�l��X�d�d�p�Kة���J�������\��������������v�|�A���c�J�[���d�P�W�m�gܯ���t�s���r�X��߳��ތ������L�M��o�چ���u���������ی���|ٱ����O�l���P�~ׇ׉�h�{�c��،�iۻ�a�n�I�z�r���R�X�����s��ء����X�s�U�zՑ������������������ݎ����A�i�E�J���������N�m�qݐ�j���ݻ�c�T����ك�w�����u�O��Q�}�z���p�q�w���e�����e����i��@�y�o�\�]ڝ�n��������W�����p�z�{ց����ݏ����������������K�u�q���y���A���ߥ���f������J��~�������xߗ�Z���u���i�������z�������H��S�e���������������������������������������������';
	v_Chard   Varchar2(2000) := '������pޅއ�����Q�_�Q�����]�J�N�^����߾�a�߰�����܍ܤܖ�J�D�\�l��ې�O�n�^���l�G�����F���^����Q��}ٜ�K���[�hו�������T�ځ�W��߶���������O�I�Z��܄����u�O��������Q�����������h�����~�L�C��ݶ���{�E��p�Mصڮۡ�s�������B���W�ޞ���K����f�d���r�B�y��ۆ���������c���������������J�M�������m����y�H��{��S���ܦ���g��Pՙ�����l��������w�����}������V�r���M�A�����[���C���HՉ�����ޓ��������h���K�^�K�Y�H�k�Z�W�L�L�^�`�a�^��`�A������i�L���x�K�G�~���o�b��tטـܶ�|��H�����Y�X������������B��m���H�����O�����ޚ�g�D�q�v��������y�I����r�o܀�D���w�y�F�G��z�����������������������������������';
	v_Chare   Varchar2(2000) := '��ވݭ�e��M�~�P��~�Z�[�F��E�i�����q������ܗ������`�Q�]�����{�O�Iج�@������t��׆�y�|�{�����E�z�[���X�b���W�������s�D߃ڍ٦�@�E����';
	v_Charf   Varchar2(2000) := '�e���y���z��N�c�x�Y�Cެ�������x���x�܏��؜�G����������ړ�[�p�h���������J���w����q�����E�y����������u���M������]���m�p����V���X�k�����r�M�M�R�v������a����L�����h�S��Qۺ�b�p�K���Sٺ�R�L�P�iو��]ߑ߻���K�������a�~�f�W���A�F����ܽ�������������ۮ��I��ݳ���������E�R�V�O�D�h��q�D�~ݗ�H�v���f��߼���M����o�f������ؓ������xݕ�Vَ�����v���������������';
	v_Charg   Varchar2(2000) := '�٤������m�����p�@�������B�d�W�^�Yؤ�}��|�����������N�v�����s��ߦ���h������M�l��G��s���غ�z�k���������޻�ھ۬��a������x�m�w�g���ت��ܪ����w����k���Yݑ�s�k�u�P�R���������t��ب��ݢ���f�s���Q������������p���������\�ؕ�C�����h�^���x����xڸ������g��ُ�������L�M�����Y��ݞ���X��ڬ����E�����������]����������d���A����N�o�T�W������ڴ���v�K���P���b݄�]�I�A����؞�k���q���X�}�������_��U���ߞ�F�|���q���h�k�I������{܉�������F���W�Z�iح�����݁���P��֏�������u�����������{�R�J�^���������������������������������';
	v_Charh   Varchar2(2000) := '���x�����V�����E�A���������w�n�J�_�\���F�\���I�d�h�u֛�n��[�ކ�ؘ�@������޶����q�����ڭ�����A������؀�F�����u�i���M�����H�H�[���Y���a�R���Q�L�e�f�S�g�\������a���C޿�U��ܟ�Fްݓ��Z����ݦ�A�v�b�D�p���fޮ�������Zڧ�{�U���\�J������A�\�C�����jܩ�_�����`�c�������ܠ����i֗��������������g��L�E�{���C�K����t�U��������������U��o���_��s�I�j�k�����n֜�f����Ֆՠ�X���b�J��؎גۨ����f�}���a���߀�o�D�I�q�S�X�G���kۼ�������ߧ���Z�d�����S���������B��W���ڇ��u�m�U���e�wڶ�������Y�D�x���������ޒ�D�t�e��������������_�V�dޥ�M�_�M�f�i�T��u�}�w�������Q�F�@ڻ��՟�����x߫ߘ�����X؛�A�f�o����޽�Z�[���������������������������������';
	v_Charj   Varchar2(2000) := 'آߴ����ܸ����������|���������K����}�u����Z�Y�ي���u�I����^��Q�Z�a�V�W�i�W�A�ؽ٥�B�L�C�������l�ު�e��n݋ۈ��U�g�P�W�n�e�|�}����������������������E�H���H������Ղ�����J۔�a�H�T�����D�V���C�q������������e���j؆����ۣ�P�����O��e�]�a���G���Z�����]�����g�y�������b���Z��K�������p�x�t���Y�[���~�d�����������������ֈ�C�r�x���{�v�|�����V����ڙ�f�`��Ր�v�{�`���G�I�T�G�Z�Y�a�b�{������������\�F�v������n���uܴ֘����������o���B�����ٮ������]�]�q����a�K�R��^��ڊ�I���_��A���������ڦڵ��ޗ����]��������m�d�R�O�^����]�����v�T������������\֔�ݣ���M������ف�B�����ݼ���X�L�~������������������i������ޟ�����e�V�K�n�o�R������ޛ�G�y������F���b�N��������J���n�����������ڠ���Շ�g�|�L�~�����`�]�R�z�v����V۞�q�G���A�������F���eڪ��ߚ�������B���e���X����Z��؋�M���L��g�m����N������������۲�h��C�������������bڑ�I���k�f�����ާ�_�`�����Q���H���B�����~�������؏�j�܊�x�z�������z�����}���K�Q�R�U�����������������������������������������������';
	v_Chark   Varchar2(2000) := '�������l��_��������]�����a���|�z�G�a���b�����٩ݨ�|ݝ��R����R�{�_�K��ߒ������`���������D���w������ڐ���V�����W�f�w�������������n��c�~�o��U�L��H�����w�x�I�y�����ߵ�@ޢ�d��ߠܥڜ���p���F٨���ۦ�����������d���wڲߝ���E�Hڿܒܜ�N�\�����������L�A�k�q���Y�k��N����k�ظ�������`��K���i���ۓ�������d�q�^������d���K���H�{�A��������S��A�p�H�T�U��������������������';
	v_Charl   Varchar2(2000) := '��������h�F�J�_�n�B������[�F��n�D���H�����l��ه�m�s�`�������@�E׎�_�|��e�����Y����Oݹ����H�q�Z����L������������������u�L�~���b߷��������E���������D�[�h�Y�m�Fڳ�C�|��P����L�[�G�K���ܨ�k����������؂�����ւ�r޼߆��x���\�P�v���g�~�Z�Gٵ�����N����������ߊ�k߿��������ٳ���\��۪����ݰ�����������W�E�_�t�`���B�b���V�]ׁ�^�Z�u�c���B����ۚ֋�`�`�H����������������b���n����cܮ��݈���u�gՏ�v�y�G�������|�Iْێ��m�������ޤ�R�������ޘ�������V�h�����Q���v�������O�����������O�l���[�������C�U�����C���\�k�`����������_�����ښ�C��q�s����C���o�w���h۹�N�g��`�����t�I�����������v���m�H���y���B�d�s�i���V���ۉ�C�w�f�j�w���������������X�N�����L�[�x�_�T���]�L������s֌�}���V�����U��ߣ������������������_�z�B�|�R�u�u���������u��������T�`��ڀ�j��X��h�j���ۍ�A�G�I�c�n�e��������y�L�����X�r�������F��[���s�x����i݆��MՓ���b�������߉������s���������������i��������������������������������������������������������';
	v_Charm   Varchar2(2000) := '�����j��i����U���K��ݤ�I��۽���u�~�@�A������M�N�����ܬ�������֙�N�����I����؈���ܚ�F���^����������T����Q�|�����ݮ�d���������Y�[�B�q��z�V�eڛ�m�i�������T�Y�{������ޫ�������������s�X��L�����i���Q���D�W�_������i�����������S���J����������ڢ������������k������������|�r���@�M�I�����]��������ؿ���p�f�x����������ق��s�F�����h��w�}������ڤ�p������Q�����և�������N����փք�O�O�������{����������a�����ٰ���w�\�����E��a�[������������f��J���������������������������';
	v_Charn   Varchar2(2000) := '�y����~�����vܘ�y�c���ܵޕ�����ؾ�r����a��Q�y�������T����߭�Qث��������D�t����m�[ګ�H��������G�����\����C�؃�r���F���u٣��b�W�Xދ�������R���D�T��݂ۜ�Tإ���|������B������������c�W�f�h�R����E���mב�b��D����_�V�H���������o�ٯ���r�s�x�P�a�e��k�����������S����Q�G�����S����Zہ�������������';
	v_Charo   Varchar2(2000) := '���Mک��k֎��t�{����';
	v_Charp   Varchar2(2000) := '����ٽ݇�����W���A��ۘ�o�Q���G��b�����Q�������������������N����B���k����r���������\�����J�o��ܡ�~�A���i�J�m�s�Cا������w�W�t��Y�����C�B�V�o����ۯ�����u����R���Q���dܱ��������|�aߨ���������G�����@�����X՗�����G՛�������g�h�w�Q����o������د���v�ؚ���l�A����ٷ�Z�Z���Z��݃�G����N�w�k۶�c�����O���H�H���������������T����ٟ�h����E���V�������������������������';
	v_Charq   Varchar2(2000) := 'ހ���V������Ճ�p�[����t�K������������ݽږܙ�H���������D���n�o���Rޭ�a�W�������u�}���G�y����ߌ�ܻ��M����H�M�����������ڞ�M�����ܷ�����@�T�`�e��U��e�Ս�w��t�v�c�k�B�R�S�aݡ����ܝ�j�@�Q�E�X�Z�b����������l�c����ٻ��݀���������ۄ��ۖ�j�I�j�����m�������b�z���������^�N����@ډ�E�F�A����������Sڈ�y�X���ڽ�~�V�m�I�N��������o��@����W�z�����s�d�������V���_�c��u�����i�W���X�p�����[�������Ո���m�����������^��jڂ���F�G�p�q���G����ٴޝ������U�������g�M�b�F�����j�A���ڰ�r���o�L�@څ�D�|�������O۾���@ޡ���z���޾����d��Y�xޑ�T������C��z��zڹ�������I�b��m�B�����e�j�E�����j������|��I�o����������������������������������������������';
	v_Charr   Varchar2(2000) := '�������`�X�j׌������N�v���m���������r��ך����ܐ���z�~�g��J��w��J�~���������F�g�P����݊����k���qߏ���n���޸���}��p�r�z��������M��ܛ݉ި������J��c�tټ�e�}���U�������������';
	v_Chars   Varchar2(2000) := '���ئ�����l�M�S���|�wِ��L�����D�d�����r��������b���fܣ�����C�Q�m֠�o�O�~��������|������������ߍ���������ܑ��^����������Wڨ���]���۷���b�i٠����W�X������օ�l���p�l����i�Y�}��ۿ�����f�d�h�������s��ڷ�����_�Y���h��ߕ������ן��Ք�T���}����v�j���J�H����|���W�j�����K��ًߟ�\�P���A���O�[����X���P���y�z����ݪ�Y�J�v���Z���R����������������߱���B�Y����m�K՜՞�}�S�u��|�a�����������x�ٿ��ܓ�g����S�\ݔ�_�e�����H���n���t�����_���Q�f�T������X���V���B�p���{�t�U�`�l���j��B�f�h�������������l�p���������j�t���F�����l�J�\�r���D�����������ٹ�~�L�|����ڡ����ݿ�����b��n���������g�}�`�����޴������������i�h�����x�p�_�M�qۑ�T���ݴ������m�U�S���r�w�����w�\���ݥ������{�Z������������t�����a�i����C�����������������������������������';
	v_Chart   Varchar2(2000) := '�������B�D�]�������e���`��w�J��F�O�Y�n�c��ۢ�������U�T޷�����؝�۰��Մ�]�U�t�T؍�v�Z�������g�a���y�����ۏ�M�|�U����o���G����}�Z���h�O�S�����E�����z�N�w��ޏ�����[���cػ��߯ؖ�������߂�`���L�R�e�f���Xڄ���pۇ������}�{�Y�[�����n܃������P�ڌ�L�j�V�p������D�c�l������q�`�t��٬�����p��������x�f���A����q�\���N�@�������F�����ߋ����������F�������F�P�b�c����١�����U�P���n�~������j���B�����^�W��W�Cݱ��T����I���h�����Qރܢ���r؇����o������P�j�k�n�s۝ۃ�����`��Z��ܔ�������رי��٢�u����ސ���������|���D�r���������K�z��ڗ�����������������';
	v_Charw   Varchar2(2000) := '��|��ߜ���c��������ؙܹ�Bߐݸ���������n�l�j��[�sٖ�~�@������s�y��ނ�����������ޱ���g�h�����������f�����`���W���d�S�T���������������Ն�c�l�Q��n�t�]�|���^�M�K�E�A�G���~�Z׈�^�d�nݘݜ���������������Y�j������Z�[�������l�f�O��޳�N��ݫ��b���������}����ڏ�E�w�G�_���u���N����c�M�����~���������������q�^�Rأ�@������A�����F�}����`����H�F�I�F������������';
	v_Charx   Varchar2(2000) := '����ۭ��������ݾ�T�R�������q���O�g�F���O��������a�G�H�l؉�v���T�@���^�^���@������E���v��֐��I���e�@�������|��������L�lے�^�h�ۧ�S�M�����i���V�K�_�]�S�U���B�i�P�y������������Y��ݠ��_���T�pՒ�]�����ݲ����������v�w��r۟�]�N�]����e��t�t�Pݍ�_�y�D�v������������`���U޺�����`�@�����^�}�D�R�G�o���Eܼ�_�l�m����x�����`����K�����}�A���a����z��}�P������������^���X�N���{�y���j�U���q�j�[�M����ߢ���C�P�Hא��������������ޯ�����x�C��^�k�K�a����ߔ�ݷ�d��\ܰ���g��cض܌���_�S�]���D������]�o��tߩ����Nܺכל������������T�x����q����P�V�n�M�����נ���q՚֞�z�`�Pڼ����������������[��rޣ܎������������՝�X�M�X�~�z����x����������C������X��K�j�Y��`���z���G���L���p�o���޹����������������S�\��ޙ���d�bަ��R�������������������������';
	v_Chary   Varchar2(2000) := 'ѹ����f�E���s���������\�����������������������۳���Z��iڥ����ܾ�I�Z��������}����ٲ����۱�D���V�C����y�d�o���f�d���j�k��B�|���z�s��������H�������e�V���e����ٞ܂�z�`�I�Jׅ����ח�V�W���Z����}���g���r������U݌����^�{�u�R�F�I�B�����زߺ�^س�������U���b��u��P���{�|�c���_�����������o�r��G���_׊������X�y�U����������v�����]�E�d�w�v�E�v����c�������p�t���b�s�p��ڱ������������ޖ�����O�B�D���z�U�V���{�k֖�Fׂ�@�~�������ޠ�r�i�C��}��T�t߮�����d��߽٫����������������ؗ����[�\�N�cژ�W�z������x���k�o�]����޲��؊�lٓ����s�J�G�������g�h�y�{�O�^�gܲ����ה�~�f��������ܧ���N���������w��۴��z�y�����]�l�������Y�i��L���y�[׍��ط�S�gݺ���A�����a��������v�D�s�L�]��܅���K�W����������������h��A۫�I�e�����G����{ܭ��V���a���t�O�����I�x��J���ٸ���x���k�����~������ݯݵޜ�]�K�[��������j��ݒ�O߈��ݬ���B���٧���ޔ�z�R�T�����}�|ߎ�����خ�������C�����D����~���������՘�k�N��uݛ��~���V�i�}�C��ٶ�������h��Z�o��؅���r����������C���N���������Aع�h��T�y�`�\���I�[��������r�q�O�u�X��M���N���O܆�c�d���g�S��x�t�����Oߖ���ؒ�J���ܫ�������w�@�x�߇���M�R�h���������܋��ڔ�_�X��������g�S�N���V�S�fٚ�ܿ���y���]�m�����q�E�B�q�y۩��i��\���ل�d�j�r�y��������������������������������������������������������';
	v_Charz   Varchar2(2000) := 'ش��������٪پ����ڣگں������������������ۤۥ۵۸����������ݧ����ީ����������ߡߤߪ߬߸���������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������������';

Begin
  If v_OutNum<1 Or v_OutNum>40 Then
     v_OutMaxNum:=10;
  Else
    v_OutMaxNum:=v_OutNum;
  End If;

	If v_Instr Is Null Or Length(Ltrim(v_Instr)) = 0 Then
		v_Spell := '';
	Else
		v_Input := Upper(v_Instr);
		v_Spell := '';
		For v_Bitnum In 1 .. Length(v_Input) Loop
			v_Bitchar := Substr(v_Input, v_Bitnum, 1);
			If v_Bitchar >= '��' And v_Bitchar <= '��' Then
				For v_Chrnum In 1 .. Length(v_Stdstr) Loop
					If Substr(v_Stdstr, v_Chrnum, 1) = '-' Then
						Null;
					Elsif v_Bitchar < Substr(v_Stdstr, v_Chrnum, 1) Then
						v_Spell := v_Spell || Chr(64 + v_Chrnum);
						Exit;
					End If;
				End Loop;
				If v_Bitchar >= '��' Then
					v_Spell := v_Spell || 'Z';
				End If;
			Elsif Instr('ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.+-*/', v_Bitchar) > 0 Then
				v_Spell := v_Spell || v_Bitchar;
			Elsif Instr('���������������', v_Bitchar) > 0 Then
				v_Spell := v_Spell || Chr(Ascii(v_Bitchar) - 41664);
			Elsif Instr('���£ãģţƣǣȣɣʣˣ̣ͣΣϣУѣңӣԣգ֣ףأ٣�',v_Bitchar) > 0 Then
				v_Spell := v_Spell || Chr(Ascii(v_Bitchar) - 41856);
			Elsif Instr('����', v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'A';
			Elsif Instr('����', v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'B';
			Elsif Instr('����', v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'G';
			Elsif Instr(v_Chara, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'A';
			Elsif Instr(v_Charb, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'B';
			Elsif Instr(v_Charc, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'C';
			Elsif Instr(v_Chard, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'D';
			Elsif Instr(v_Chare, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'E';
			Elsif Instr(v_Charf, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'F';
			Elsif Instr(v_Charg, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'G';
			Elsif Instr(v_Charh, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'H';
			Elsif Instr(v_Charj, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'J';
			Elsif Instr(v_Chark, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'K';
			Elsif Instr(v_Charl, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'L';
			Elsif Instr(v_Charm, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'M';
			Elsif Instr(v_Charn, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'N';
			Elsif Instr(v_Charo, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'O';
			Elsif Instr(v_Charp, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'P';
			Elsif Instr(v_Charq, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'Q';
			Elsif Instr(v_Charr, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'R';
			Elsif Instr(v_Chars, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'S';
			Elsif Instr(v_Chart, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'T';
			Elsif Instr(v_Charw, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'W';
			Elsif Instr(v_Charx, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'X';
			Elsif Instr(v_Chary, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'Y';
			Elsif Instr(v_Charz, v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'Z';
--			Else
--				v_Spell := v_Spell || '_';
			End If;
			Exit When Length(v_Spell) > v_OutMaxNum-1;
		End Loop;
	End If;
	Return(v_Spell);
End;
/

Create Or Replace Function zltools.zlWBCode(v_Instr In Varchar2,v_OutNum In Integer:=10)
	Return Varchar2 Is
	v_Code    Varchar2(40);
	v_Input   Varchar2(1000);
	v_Bitnum  Integer;
  v_OutMaxNum Integer;
	v_a       Varchar2(1200) := '�����������ð������������ݩ���ɱͱ�޵�����в����˲̲Բ�ܳ�ݲ�������ɳ�����ݻ����ʴ���������߰�̵���ݶ�ٶ�𴶭��ܶݭ�����췡��ެ�����η��Ʒ˷���ܽ������ݳ�ʸ���޻����ݢ���������ù������������������к�����޶�ɺ�޿ްݦޮ�������ȻĻ�����ޥ�������޽ܸ����ު����������������������ܴ�����ڽ�����ݣ�����ݼ�����ھ����������ާ���ܿ��ݨ������ޢ��������������ݹ������������޼����ݰ����ޤ����������«������ݤ�����âãçé��ïݮ����ޫ����������������ġĢ��Ī��ĳ��ļĹĻĽĺܵ����إ������ŷŹ�Ÿź��������ƥ��ƻƼ����������������������ݽ��ޭܻ������ܷݡ����Ǿ��������������������������������������ޡ޾��ȧȵ����������������������޸��ި��������ɢɯ��ɻ��������ݪ��ʽ����������˹ݿ޴������ݴݥ��̦޷��߯�����������ݱ��ܹݸ����ޱέήε��޳ݫ����������ݾ��������ϻݲ޺��ܼ��������аЬޯоݷн��ܺ��ޣ��ѥѦ޹��ަѻ��ѿ��������ܾ۱����Ңҩҽ������߮��޲������۴��Ӣݺ����ӫӨөӪ����ݯݵݬ�������Էܿ����������պ��������֥��ݧ����������ީ';
	v_b       Varchar2(1200) := '���������������ة���ܳ�����ϵ����������Ӷ�����⸽�����������������ؽ�ʽ�������������޿���������������¡¤ª½����İ�����������Ƹ����ȡȢ��������������������������������϶��������ѷ��ҮҲ������Ժ������ְ����׹������۸����';
	v_c       Varchar2(1200) := '���������������ҳ������ʵ��˵������ۺ��軶�����轾�澢�پ��ɾԿ���������¿������ì����Ĳ��������ƭ��������Ȱ�������ɣ�ɧ��ʥʻ˫������̨ۢ��ͨ����������Ϸ������ѱ�����������������ԦԤԥ������������פ����';
	v_d       Varchar2(1200) := '�����ưٰ����������ձ������������粼��޲�򲳧�����׳�����������������������Ǵ����������ڵ�����������������������������ܷ��ۺ�������ϸ���Ѹ���𳹼���Źʹ˹��޺��ĺ������߻ǻ������������ꩼ�������پ��ݾ¾����㿳�Ŀ��ڿ��ſ��������������������������������������������¢��µ��۽����������������������������������������������������ǣ����������Ȯȷ��������ɰ��������ʢʯ����ˣˬ˶��������̫̬̩̼��������������������������Ϯ�����������������������ѹ���������������������ҳ���Ӳ������������ԭԸ�������������������ש�����������';
	v_e       Varchar2(1200) := '����������������������ɲʲ�������⵨�������ض��񷾷���η��ڷ�ۮ��������ꮸθظ����Ź���������ܺѼ������콺���������¾��̿�������������������ò����������Ĥ��������ؾ��������ŧ������������Ƣ��������ǻ������������������������ʤ������˦˴̥���������������������������ϥ������в����������������ң��Ҩ��Ҹ����ӯ��ӷ����������������������֫��֬������������������';
	v_f       Varchar2(1200) := '��������Ӱ��ౢ����ò���𾲺�Ųó�������ó���ܯ���ء����ܤ��������ص�����ܦ�����¶Ѷն�������ط����������ܪ���������챹���Թ����������������������ܩ��첻�����������弪����ν̽��ؽ���������������������ǿ�ܥ�����ۿ�����������ܨ������ܮ��������۹��¶��������ܬ�ù��������������������ܡ��������ܱ��ƺ������������Ǭǽ����������������ȥȤȴ������ɥܣ��������������ʮ��ʿʾ�˪��������̮̳̹�������������ܢ����Τ��Υδ��������������ϲϼ������ТЭܰ������ѩ��������Ҽ����ܲܧ����ܭ������ԪԫԬ��ܫԶ��Խ���������������������������ֱ֧��ַ־����ר����';
	v_g       Varchar2(1200) := '訰����۰����������˰�߱±Ʊ̱���������Ͳ��в�貲����뷴������譴��������鵽������ඹ�������Ѷ�ج�������󸦸�ؤتب��������������������������آ���Լ������ۣ��ꧼ߼����誽���������忪�������������������������۪����������������������������õ��������ĩ����ث�Ū��ا������ƽ����������������������ئɪɺ��������������������������������������������������������أ����������������ѳ������۳����������һ�������������������������������յ�����������ֳ��ۤ����׸��';
	v_h       Varchar2(1200) := '���벷���ǲ��Ӳ���������ݳ����ƴ���������㶢��������ح��򮻢�����޾������������ǿ������������¬�­±²�����������������Ŀ����Ű��Ƥ��Ƶ����������ȣ�������˯˲������������ͫ͹��ϹЩ����ѣ��������������գ��հռս����ֹ��׿��������������';
	v_i       Varchar2(1200) := '������������������������������������멳��������γس����ȴ���壵����ʵεӵ���¶������������Ƿз���㸡�����ٸ������Ƹ۹������ʹ��������������婺�尺Ӻ�������������䰻������佻��������䧻�䫻��������䩼�䤼��ս�������䮽��н������ƾھ�举�������������������������������������������������������������©��������º������������������û�������������������ĭĮ����������ŢŨŽ��������������������Ư������������������������ǢǱǳ����������ȸȾ������������ɬɳ��������������������������������ʡʪ������������ˮ�������������������̶̭̲���������������������ػ����͡��Ϳ�������������Ϋ�μ���������������ϫ���Ϫϴ��������������СФйк������������������ѧ���Ѵ������������������ҫҺ��������������Ӿӿ����������ԡ��Ԩ��Դ�������մ�տ�������������֭����������ע��������������';
	v_j       Varchar2(1200) := '�Ӱ������������������˳����׳�������ꭵ���������������꽸������ƹ�������������Ϻ�������º�������ͻ׻������̼���������������������������������������������������������������ð�������������������������ů������������������������������������������ɹ�������������ʦʱ����������˧�����������������������������������������������ϺϾ����������ЪЫ���������������������Ұҷ���������ӬӰӳӼ������خ�����������Ի��������������֩����������';
	v_k       Varchar2(1200) := '߹���İ�������໰Ȱɰ������Ա��������ϱ������߲�����������𳪳������ʳѳ�������߳���������ߴ��������������������઴�߾�߶����������ڵ������޶����ྲֶྀ��Ͷ���������������ȷͷ���߻����߼���¸��ø����칾�ɹ��۹���ù��˺����ƺ��źǺ����ٺߺ�����������׻����������ߴ��������������¼����ӽ�������౾��������������ǿ�����྿пԿ�ߵ�޿����������������߷�����߿������������������ʿ������·����������������������������������������������������ŶŻſž����������������ơ����Ʒ����������ǲǺ�����������������ɤ��ɶ��������ʷ�����˱˳������˻����������������̤̣��̾����������������������������Ψζι���������������������������������Х������������ѫѽ������ߺҧҭҶ��������߽���������۫Ӵ��ӽӻ������ԱԾ��������������߸զ��������֨����ֺֻ������������������������������';
	v_l       Varchar2(1200) := '���հ�����������������������������̹�����غں�����ػ�������¼׼ݼ���νϽ��������������������������������������īĬ���������������ǭǵ������������Ȧ���������������˼������ͼ��������ΧηθϽ������Ѽ�����������������԰Բԯ����������նշ�����������������ת���';
	v_m       Varchar2(1200) := '������ᱰܱ����±����ƲƲ�᯳����״ʹ��ϵ��ص����۵�����ǶĶ��뷫ᦷ��������᥸�����Ըո���ḹ�����������������᲻˻ϻ�᧼�����ᵼ����������������Ⱦ���������������ܿ��������������¸��ñ��������������������Ƕ�������Ƚ����ɽɾ����������������������̿������ͬͮ��������Ρ������Ͽ�������������������������Ӥ�������������������������ո����������֡�������������׬����';
	v_n       Varchar2(1200) := '����ͱ�㹱ٱ��Աܱ��������޲��Ѳ�������������������ٳ����㰳��������㲴����򵬵��ᵼ�����󶮶�����㵷���㭷����ĸҹֹߺ������޺���ﻳ�Ż̻лֻ��켺�ɼ½쾡��㽾Ӿ־���������鿶��㡿���������������¾��������æü����ؿ���������������ų����������ƨ�Ʃ��ǡ��������������������ʬʭʺ��������ˢ˾����������������������Ωβξο��������ϧϬϰ����ми�����������������Ѹ������������������������������������������������չ�������������';
	v_o       Varchar2(1200) := '�α���������Ӳ��̳��㴶�ٴִ�ƶ��������ܷ۷��������ʺ��������κ����ͻ���ݽ����������߿�����������������������������¦¯��ú��������Ŵ��������Ȳȼ������ɿ����������˸��۰����������Ϊ���ϩ��Ϩ���������������������ҵ��������������������ըճ������������';
	v_p       Varchar2(1200) := '���������ٱ���������ѱ���������巳����ݳ���Ҵ�������������崵�����ֶ��𸤸����˸�ѹӹٹ�峺������ֺ׺���徻��ջ���żļ�������忽��ƽ󾽾������������̿տܿ߿��������������������»����������ڢ��������ڤį�����ũ����������������������������ȹ������������������������ʵ������������������̻�����ͻ��������������������������д����Ѩ���Ҥ���������������Ԣԣԩ۩����լխկ�������֮������ڣ��ף�����������';
	v_q       Varchar2(1200) := '�����������Ӱ��������������������������ٱ������ಬ���в���������β���������˳���������ۻ��������˭[�������ﱴ�������صۡ���������������������׶��ƶ���۶���ﰶ��������ܶ��Ƕ��ﷰ����������������������븺�����ŸƸָ���Ӹ�����������������������������������������������������������ۨ��ۼ�������������켢�������������ﵽ����޽ǽƽȽ����ڽ��Ľ���˾�ⰾ��þľ�︾��Ҿ���������������������������������������������������������������������������������������������³��������������������èêîíó���þ�����������������������������������������������������ť���������������������������ǥǦ��ǮǯǷ��������������������۾�����Ȼ��������������ɫ�ɱɲ�ɷ���������ʨ���ʴ��������߱������������������������̡��������������������������ͭ���������Σ����������������Ϧϣۭ�����ϳ�ۧ�������������������������зп�������������������Ѯ��س�������ҿ���������������������ӡӭ��������������������������ԧԳԹԿ�������աղ������������������������������������׶����������������';
	v_r       Varchar2(1200) := '���߰�������ְǰư��ɰ�������ڰݰ���豨�������ձ����𲦲����������ٲ�����󳭳������ӳ��γֳ�鴤����밴�ߥ�ݴ�����򵣵�����뮵��ĵֵ��������뱶��ܶݶ޶�������������������ߦ�׸޸��롹��Ϲҹ�������⺤���������˻���ߧ���ػ�߫���Ἴ���������븽ӽ��׽ݽ��������ܾݾ������ܿ���������봿ؿٿۿ������������������������������£§ߣ°��������������������ĨĴ����߭��������������šţŤŲ��������������������ۯߨ�Ʋƴƹ���������������Ǥ������������������Ȫȱ��������������ɦɨ������ʧʰ������������������ˤ˩��˺��������̢̧̯̽����������ͦͱͶ���������������������������������ЮЯߢж��ߩ����Ѻ�����������ҡ��ҴҾ���������ӵ��Ԯ�����������������ժ��������������ߡ������شִ����ָ����ֿ�������ۥ��ץצק׫ײ׾׽ߪ��߬ߤ';
	v_s       Varchar2(1200) := '��ذ��������輱���ı����Ĳ�農���߳���������ȳ��˳����������鳴�������������嵵��馶�����������鲷�����鼷����ٷ������Ÿ˸����ϸ����¸�������۹�����������������������������뻱�������鮼ϼ��ż�饼����������Ƚ�����駽������۲���ӿ���¿ÿɿݿ�����������������������������������������������������¥����´����ö÷�ø������������ģľ���������������������ƮưƱ��������������ǹ�����������Ȩȩȶ������ɭɼ�����������������˨��������������̪��̴������������ͩͪͰ������������Φ������������ϭ������������УШе������������������������������ҪҬ����ӣ����������������դեջ���������������֦��ֲ������������������׮׵������������';
	v_t       Varchar2(1200) := '��᮰°ʰް�����������Ǳ����������������ֱ��������������߳��䳹�Ƴ˳ͳ̳���������������������Ѵ����޴��������δ���쵾�õ�������еѵ���뺶�����빶��ƶ������췦�����������������������͸���غ�ݸ��۬���������ѹ����������Թ��������̺ͺܺ�������˻�����ɻջ�������������ż�����ռڼ����ȼ�림�����սֽ�̾�����������������ؿ������������������������������¨���������������ë��ôÿ�����������������ĵ�����������������������Ƭ��ƪد��۶���������ǧǨ�ǩ�������������������������������ɸ����������������������������ʣʸ��������������˰˽������������̺������������ͧ͢Ͳ͸ͺͽ��ر������΢ί��κ��������������Ϣ��Ϥϡ����ϵ������������������Цض��������ѡѪѬѭ��������������ز������������ط���������ع������������������������էձ�����������֪������������������������������׭����������';
	v_u       Varchar2(1200) := '�������걱�����űձֱԱױ��������������������곲����������Ѵ������񬴯�ôɴ��˴δմ���񵥵���������Ƶܵ۵������۶������˶��շ����������������ع�������غ�������������Լ�񤼽���������ϼ����彪���������������ܽ��ɽ꾻���������ξξ��þ������������������������������������������������������������������������������������ű��������ƣƦ�Ƴƿ������ǰǸǼ������������ȭȬȯȳ�����������������۷������������˷��ڡ�����������̵̱������������ͯʹͷ��������������������������������Ч���������ѢѾ�����������������������ұ������������������Ӹ�����������������բ���վ��۵������֣֢��������ױװ��׳״׼���������������������';
	v_v       Varchar2(1200) := '�������滰�����ϳ�����槴�����淶ʶ�湷�������������Ź�棺ûٻ鼧������˼��߼޼齨毽�����������žʾ˾�����ѿ�ظ�����������¼������ý����������ķ��������������ū����Ŭ����ŭŮ����������Ⱥ���������ɩ���ۿ������ʼ���ˡ���������������������������������ϱ�������������ѰѲ���������Ҧ����������������������������������';
	v_w       Varchar2(1200) := '�����˰ְ۰����ұ��������±��ϲ��������٭���������ѳ�ٱ�Ŵ������ӴԴ���ߴ���ᷴ������ٵ����ǵʵ͵���궱�ζ�٦���ҷ·ַݷ���ٺ�������������������٤��������������Ĺ�����Ϻκ���򻪻�����٥�����ʼ��Ѽۼټ�������Խ�������ٮ�Ľ�������־��ƾ���������٩����٨��ۦ��������ٵ����ٳ�����������������������������������ٰ������٣��������ٯ��żٽ������Ƨƫƶٷƾ����Ǫ��ٻ���������������ٴȫ���������������ټ��ɡɮɵ���������ʲʳʹ������ٿ�������ٹ��������������̰��������٬ͣ��١��͵;��٢����ΰαλ����������������������������б�������������ٲ������ү��������������٫Ӷٸ��������٧������������ٶ��������ծ������ֵֶ����٪��ס��׷پ��������������';
	v_x       Varchar2(1200) := '�����ذ�ȱϱѱ��������Ͳ��ڳ�穴���窵��޶�綷�糷ѷ׷츥������礸�������箹��ù��笺�컡�ٻ��û��������ܼ��ͼ̼����ֽ����ʽɽԽ��ƾ�������������������������������������������������ĸ��Ŧ��������ǿ�����������������ɴ�������������˿�����������ͳ�����άγ����ϸ���������������������Ѥ��������������ӧ�ӱ��������ԵԼ���������ּֽ֯���������׺�����������';
	v_y       Varchar2(1200) := '����������������ر��������������Ʋ������߳ϳ䴲�ȴʵ���ڮ���е�����������ض��̷������÷���̷Ϸ������ø߸�ھ����ڸڬ��ڴ���ѹ������������ڭ����ڧ�����軰����ڶ����ڻ�������ƼǼ������ɽ�ڦڵ����۾�������ڪ�����念���̿ο�ڲڿ���������������ڳ��������������������������®¹���������������������������á��ä���å��������������������ĦĥħıĶګ�ŵک���������������������������������ǫǴ����ڽ����ڰڹȿ����ڨ������ڷ��ʫʩʶ��������������˥˭˵����������̸̷����ͤͥ��������������Ϊ��ν������������ϯ��������г��л������ڼ������������ѯѵѶ��ڥ������ҥҹ����ڱ�������������������Ӧ��ӥ��ӮӹӺ���������������������թի���گ������������֤ں����������ؼ����ׯ׻��������������';
Begin
  If v_OutNum<1 Or v_OutNum>40 Then
     v_OutMaxNum:=10;
  Else
    v_OutMaxNum:=v_OutNum;
  End If;
  
	If v_Instr Is Null Or Length(Ltrim(v_Instr)) = 0 Then
		v_Code := '';
	Else
		v_Input := Upper(v_Instr);
		v_Code  := '';
		For v_Bitnum In 1 .. Length(v_Input) Loop
            if Instr('ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.+-*/', Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || Substr(v_Input, v_Bitnum, 1);
			Elsif Instr(v_a, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'A';
			Elsif Instr(v_b, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'B';
			Elsif Instr(v_c, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'C';
			Elsif Instr(v_d, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'D';
			Elsif Instr(v_e, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'E';
			Elsif Instr(v_f, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'F';
			Elsif Instr(v_g, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'G';
			Elsif Instr(v_h, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'H';
			Elsif Instr(v_i, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'I';
			Elsif Instr(v_j, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'J';
			Elsif Instr(v_k, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'K';
			Elsif Instr(v_l, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'L';
			Elsif Instr(v_m, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'M';
			Elsif Instr(v_n, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'N';
			Elsif Instr(v_o, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'O';
			Elsif Instr(v_p, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'P';
			Elsif Instr(v_q, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'Q';
			Elsif Instr(v_r, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'R';
			Elsif Instr(v_s, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'S';
			Elsif Instr(v_t, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'T';
			Elsif Instr(v_u, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'U';
			Elsif Instr(v_v, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'V';
			Elsif Instr(v_w, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'W';
			Elsif Instr(v_x, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'X';
			Elsif Instr(v_y, Substr(v_Input, v_Bitnum, 1)) > 0 Then
				v_Code := v_Code || 'Y';
			End If;
			Exit When Length(v_Code) > v_OutMaxNum-1;
		End Loop;
	End If;
	Return(v_Code);
End;
/

