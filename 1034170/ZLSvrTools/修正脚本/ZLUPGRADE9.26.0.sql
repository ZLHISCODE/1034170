-----------------------------------------------------------------
--为配合产品版本号由9.25升为9.26
--从本次起开始使用升级工具升级 
-----------------------------------------------------------------
--新的自动升级工具日志记录
CREATE TABLE zlTools.zlUpgrade(
	系统 NUMBER(5),
	原始版本 VARCHAR2(10),
	目标版本 VARCHAR2(10),
	升迁时间 DATE,
	升迁结果 NUMBER(1),
	结果版本 VARCHAR2(10),
	中止语句 VARCHAR2(200))
	PCTFREE 5 PCTUSED 90
/
ALTER TABLE zlTools.zlUpgrade ADD CONSTRAINT 
    zlUpgrade_UQ_升迁时间 Unique (系统,升迁时间)
    USING INDEX PCTFREE 5
/
ALTER TABLE zlTools.zlUpgrade ADD CONSTRAINT
    zlUpgrade_FK_系统 FOREIGN KEY (系统) REFERENCES zlSystems(编号) ON DELETE CASCADE
/
CREATE PUBLIC SYNONYM zlUpgrade for zlTools.zlUpgrade
/
GRANT SELECT ON zlTools.zlUpgrade TO PUBLIC 
/
Begin
	For r_User In(Select 所有者 From zlSystems) 
	Loop
			Execute Immediate 'Grant Select,Insert,Update,Delete on zlTools.zlUpgrade to '||r_User.所有者||' With Grant Option';
	End Loop;
End;
/

--补充主键及唯一索引,提高SQL效率
Begin
	Begin Execute Immediate 'Drop Index zlTools.zlRPTDatas_IX_报表Id'; Exception When Others Then Null; End;
	Begin Execute Immediate 'Drop Index zlTools.zlRPTConds_IX_报表Id'; Exception When Others Then Null; End;

	Delete From zlTools.zlRPTDatas A Where RowID<(Select Max(RowID) From zlTools.zlRPTDatas B Where A.报表ID=B.报表ID And A.名称=B.名称);
	Delete From zlTools.zlRPTConds A Where RowID<(Select Max(RowID) From zlTools.zlRPTConds B Where A.报表ID=B.报表ID And A.条件号=B.条件号);
	Delete From zlTools.zlRPTConds A Where RowID<(Select Max(RowID) From zlTools.zlRPTConds B Where A.报表ID=B.报表ID And A.条件名称=B.条件名称);
End;
/
ALTER TABLE zlTools.zlRPTDatas ADD CONSTRAINT zlRPTDatas_UQ_名称 UNIQUE (报表ID,名称) USING INDEX PCTFREE 5
/
ALTER TABLE zlTools.zlRPTConds ADD CONSTRAINT zlRPTConds_PK PRIMARY KEY (报表ID,条件号)
/
ALTER TABLE zlTools.zlRPTConds ADD CONSTRAINT zlRPTConds_UQ_条件名称 UNIQUE (报表ID,条件名称) USING INDEX PCTFREE 5
/

-- 包调整，兼容9i
-- 删除不使用的包
drop package b_datamana
/

--陈福容,9040
Insert Into zlOptions(参数号,参数名,参数值,缺省值,参数说明) Values(7,'提醒服务参数',';9999;0',';9999;0','用于提醒服务的服务器名、端口号及状态等信息。')
/

-----------------------------------------------------
-- 创建包头 2006-8-24, 15:41:11 --
-----------------------------------------------------
CREATE OR REPLACE PACKAGE ZLTOOLS.b_Expert IS
  -----------------------------------------------------------------------------
  -- 作者： 陈东
  -- 创始时间：2006-6-29
  -- 修改人：
  -- 修改时间：
  -- 描述：
  --         主要用于专项工具的过程
  -----------------------------------------------------------------------------

  TYPE t_Refcur IS REF CURSOR;

  -----------------------------------------------------------------------------
  -- 取提醒数据
  -- 调用列表： frmNoticesEdit.ReadData、frmNoticeTools.cboSystem_Click
  -----------------------------------------------------------------------------
  PROCEDURE Get_Notices
  (
    Cursor_Out OUT t_Refcur,
    序号_In    IN Zlnotices.序号%TYPE,
    系统_In    IN Zlreports.系统%TYPE := NULL
  );

  -----------------------------------------------------------------------------
  -- 取提醒对象数据
  -- 调用列表： frmNoticesEdit.ReadData
  -----------------------------------------------------------------------------
  PROCEDURE Get_Noticeusr
  (
    Cursor_Out  OUT t_Refcur,
    提醒对象_In IN Zlnoticeusr.提醒对象%TYPE,
    提醒序号_In IN Zlnoticeusr.提醒序号%TYPE
  );

  -----------------------------------------------------------------------------
  -- 取可以选择的提醒报表
  -- 调用列表： frmNoticesEdit.cmdOpen_Click
  -----------------------------------------------------------------------------
  PROCEDURE Get_Noticereport
  (
    Cursor_Out OUT t_Refcur,
    系统_In    IN Zlreports.系统%TYPE
  );

  -----------------------------------------------------------------------------
  -- 在不同的系统间复制报表
  -- 调用列表：mdlMain.CopyReport
  -----------------------------------------------------------------------------
  PROCEDURE Copy_Report
  (
    系统_In   IN Zlreports.系统%TYPE,
    新系统_In IN Zlreports.系统%TYPE
  );

END b_Expert;
/

CREATE OR REPLACE Package ZLTOOLS.b_Loadandunload Is
  -----------------------------------------------------------------------------
  -- 作者： 陈东
  -- 创始时间：2006-6-29
  -- 修改人：
  -- 修改时间：
  -- 描述：
  --         主要用于装卸管理的过程
  -----------------------------------------------------------------------------

  Type t_Refcur Is Ref Cursor;

   -----------------------------------------------------------------------------
  -- 功能：取有SysFiles表的文件名
  -- 调用列表：frmAppCheck.cmbSystem_Click、frmClearData.cmbSystem_Click、frmAppScript.cmbSystem_Click
  --           frmAppUpgrade.cmbSystem_Click
  -----------------------------------------------------------------------------
  Procedure Get_Sysfile_Name
  (
    Cursor_Out Out t_Refcur,
    系统_In    In zlSysFiles.系统%Type,
    操作_In    In zlSysFiles.操作%Type
  );

  -----------------------------------------------------------------------------
  -- 功能：取共享系统名称
  -- 调用列表： frmAppStart.cmdFunction_MouseUp
  -----------------------------------------------------------------------------
  Procedure Get_Share_Name
  (
    Cursor_Out Out t_Refcur,
    共享号_In  In zlSystems.编号%Type
  );

  -----------------------------------------------------------------------------
  -- 功能：取Oracle版本号
  -- 调用列表： frmAppStart.Form_Load、frmImp.Form_Load、frmLoadIn.Form_Load
  -----------------------------------------------------------------------------
  Procedure Get_Oracle_Ver(Cursor_Out Out t_Refcur);
End b_Loadandunload;
/

CREATE OR REPLACE Package ZLTOOLS.b_Popedom Is

  -----------------------------------------------------------------------------
  -- 作者： 陈东
  -- 创始时间：2006-6-29
  -- 修改人：
  -- 修改时间：
  -- 描述：
  --         主要用于权限管理的过程
  -----------------------------------------------------------------------------
  Type t_Refcur Is Ref Cursor;

  -----------------------------------------------------------------------------
  -- 功能：CopyMenu
  -- 调用列表：mdlMain.CopyMenu
  -----------------------------------------------------------------------------
  Procedure Copy_Menu
  (
    系统_In   In zlMenus.系统%Type,
    新系统_In In zlMenus.系统%Type
  );

  -----------------------------------------------------------------------------
  -- 功能：取ZlMenu数据
  -- 调用列表： frmMenu.cmdExp_Click、frmMenu.FillMenu
  -----------------------------------------------------------------------------
  Procedure Get_Menu_Tree
  (
    Cursor_Out Out t_Refcur,
    组别_In    In zlMenus.组别%Type
  );

  -----------------------------------------------------------------------------
  -- 功能：取ZlMenu数据
  -- 调用列表： frmMenu.cmdNew_Click、frmMenu.FillMenuName
  -----------------------------------------------------------------------------
  Procedure Get_Menu_Group
  (
    Cursor_Out Out t_Refcur,
    组别_In    In zlMenus.组别%Type
  );

  -----------------------------------------------------------------------------
  -- 功能：取模块
  -- 调用列表： frmProgPriv.Fill模块
  -----------------------------------------------------------------------------
  Procedure Get_Module
  (
    Cursor_Out Out t_Refcur,
    系统_In    In zlComponent.系统%Type
  );

  -----------------------------------------------------------------------------
  -- 功能：取功能或排列，说明
  -- 调用列表： frmProgPriv.Fill功能、frmProgPriv.Fill表权限
  -----------------------------------------------------------------------------
  Procedure Get_Function
  (
    Cursor_Out Out t_Refcur,
    系统_In    In zlProgFuncs.系统%Type,
    序号_In    In zlProgFuncs.序号%Type,
    功能_In    In zlProgFuncs.功能%Type := Null
  );

  -----------------------------------------------------------------------------
  -- 功能：取表权限
  -- 调用列表： frmProgPriv.Fill表权限
  -----------------------------------------------------------------------------
  Procedure Get_Impower
  (
    Cursor_Out Out t_Refcur,
    系统_In    In zlProgPrivs.系统%Type,
    序号_In    In zlProgPrivs.序号%Type,
    功能_In    In zlProgPrivs.功能%Type
  );

  -----------------------------------------------------------------------------
  -- 功能：得到角色能访问的导航台工具
  -- 调用列表： frmRole.FillModule
  -----------------------------------------------------------------------------
  Procedure Get_Role_Tools
  (
    Cursor_Out Out t_Refcur,
    角色_In    In zlRoleGrant.角色%Type
  );

  -----------------------------------------------------------------------------
  -- 功能：得到以前的权限
  -- 调用列表： frmRoleGrant.cmdOK_Click
  -----------------------------------------------------------------------------
  Procedure Get_Role_Grant
  (
    Curgrand_Out    Out t_Refcur,
    Curprivs_Out    Out t_Refcur,
    Curfuncpars_Out Out t_Refcur,
    角色_In         In zlRoleGrant.角色%Type
  );

  -----------------------------------------------------------------------------
  -- 功能：FillFunc
  -- 调用列表： frmRoleGrant.FillFunc
  -----------------------------------------------------------------------------
  Procedure Get_Zlprogfunc
  (
    Cursor_Out Out t_Refcur,
    系统_In    In zlProgFuncs.系统%Type,
    序号_In    In zlProgFuncs.序号%Type
  );

  -----------------------------------------------------------------------------
  -- 功能：是所有角色对应的模块
  -- 调用列表： frmUserEdit.UserEdit
  -----------------------------------------------------------------------------
  Procedure Get_All_Module(Cursor_Out Out t_Refcur);

End b_Popedom;
/

CREATE OR REPLACE Package ZLTOOLS.b_Public Is
  -----------------------------------------------------------------------------
  -- 作者： 陈东
  -- 创始时间：2006-6-29
  -- 修改人：
  -- 修改时间：
  -- 描述：
  --         公共过程
  -----------------------------------------------------------------------------

  Type t_Refcur Is Ref Cursor;

  -----------------------------------------------------------------------------
  -- 功能：取系统日期
  -- 调用列表：
  -- mdlMain.CurrentDate
  -- clsDatabase.CurrentDate
  -----------------------------------------------------------------------------
  Procedure Get_Current_Date(Cursor_Out Out t_Refcur);

  -----------------------------------------------------------------------------
  -- 功能：删除错误日志或运行日志
  -- 调用列表：
  -- mdlMain.DeleteAllLog
  -----------------------------------------------------------------------------
  Procedure Delete_All_Log(Runtimelog_In In Number := 0);

  -----------------------------------------------------------------------------
  -- 功能：删除当前运行日志
  -- 调用列表：
  -- mdlMain.DeleteCurLog
  -----------------------------------------------------------------------------
  Procedure Delete_Diarylog
  (
    会话号_In   Number,
    用户名_In   Varchar2,
    工作站_In   Varchar2,
    部件名_In   Varchar2,
    工作内容_In Varchar2,
    进入时间_In Date
  );

  -----------------------------------------------------------------------------
  -- 功能：删除当前错误日志
  -- 调用列表：
  -- mdlMain.DeleteCurLog
  -----------------------------------------------------------------------------
  Procedure Delete_Errorlog
  (
    会话号_In   Number,
    用户名_In   Varchar2,
    工作站_In   Varchar2,
    类型_In     Number,
    错误序号_In Number,
    时间_In     Date
  );

  -----------------------------------------------------------------------------
  -- 功能：取注册码
  -- 调用列表：
  -- mdlMain.Get注册码
  -----------------------------------------------------------------------------
  Procedure Get_Regcode(Cursor_Out Out t_Refcur);

  -----------------------------------------------------------------------------
  -- 功能：取版本号
  -- 调用列表：
  -- mdlMain.UpgradeManager
  -----------------------------------------------------------------------------
  Procedure Get_Ver(Cursor_Out Out t_Refcur);

  -----------------------------------------------------------------------------
  -- 功能：更新版本号
  -- 调用列表：
  -- mdlMain.UpgradeManager
  -----------------------------------------------------------------------------
  Procedure Update_Ver(Verstring_In In Varchar2);

  -----------------------------------------------------------------------------
  -- 功能：取得系统所有者名称
  -- 调用列表：
  -- frmStatus.cmbsystem_Click、mdlMain.GetOwnerName、mdlMain.cmbSystem_Click
  -- frmAutoJobs.cmbSystem_Click、frmDataMove.cmbSystem_Click 、frmNoticeTools.cboSystem_Click
  -- frmProgPriv.ProgPriv、frmAppScript.cmbSystem_Click
  -----------------------------------------------------------------------------
  Procedure Get_Owner_Name
  (
    Cursor_Out Out t_Refcur,
    编号_In    In zlSystems.编号%Type := 0
  );

  -----------------------------------------------------------------------------
  -- 功能：取注册表中信息
  -- 调用列表：
  -- frmAbout.GetUnitInfo、frmAutoJobs.From_load、frmClientsUpgrade.InitInfor
  -- frmFilesSet.ShowEdit、frmRegist.From_load、frmAppScript.From_Load
  -- frmFilesSendToServer.InitInfo
  -----------------------------------------------------------------------------
  Procedure Get_Reginfo
  (
    Cursor_Out Out t_Refcur,
    项目_In    In zlRegInfo.项目%Type := Null
  );

  -----------------------------------------------------------------------------
  -- 功能：取zlGetSvrToolsg数据
  -- 调用列表：
  -- frmMDIMain.MDIForm_Load
  -----------------------------------------------------------------------------
  Procedure Get_Zlsvrtools(Cursor_Out Out t_Refcur);

  -----------------------------------------------------------------------------
  -- 功能：取已安装系统清单
  -- 调用列表：
  -- frmAppCheck.Form_Load、frmClearData.Form_Load、frmDataMove.Form_Load
  -- frmImp.FillSystem、frmLoadIn.FillSystem、frmLoadOut.FillSystem
  -- frmMDIMain.mnuFileRemove_Click、frmNoticeTools.Form_Activate、frmRoleGrant.FillSystem
  -- frmAppUpgrade.Form_Load、frmAppScript.Form_Load、frmExp.FillSystem
  -- frmInputTools.from_activate、fromRole.FillSystem、frmAutoJobs.From_load
  -- frmAppstart.sysCreated
  -----------------------------------------------------------------------------
  Procedure Get_Zlsystems
  (
    Cursor_Out Out t_Refcur,
    所有者_In  In zlSystems.所有者%Type := Null
  );

End b_Public;
/

CREATE OR REPLACE Package ZLTOOLS.b_Runmana Is
  -----------------------------------------------------------------------------
  -- 作者： 陈东
  -- 创始时间：2006-6-29
  -- 修改人：
  -- 修改时间：
  -- 描述：
  --         主要用于运行管理功能的过程
  -----------------------------------------------------------------------------

  Type t_Refcur Is Ref Cursor;

  -----------------------------------------------------------------------------
  -- 功能：取ZlAutoJob序列号
  -- 调用列表：
  -- frmAutoJobset.cmdok_click
  -----------------------------------------------------------------------------
  Procedure Get_Job_Number
  (
    Cursor_Out Out t_Refcur,
    系统_In    In Number
  );

  -----------------------------------------------------------------------------
  -- 功能：取ZlDataMove描述
  -- 调用列表：
  -- frmAutoJobset.cmdUpdate_Click
  -----------------------------------------------------------------------------
  Procedure Get_Depict
  (
    Cursor_Out Out t_Refcur,
    系统_In    In zlDataMove.系统%Type,
    组号_In    In zlDataMove.组号%Type
  );

  -----------------------------------------------------------------------------
  -- 功能：取zlClients的MAX IP
  -- 调用列表：
  -- frmClientsEdit.InitCard
  -----------------------------------------------------------------------------
  Procedure Get_Client_Maxip(Cur_Out Out t_Refcur);

  -----------------------------------------------------------------------------
  -- 功能：取zlClients的记录
  -- 调用列表：
  -- frmClientsEdit.InitCard、frmClientsParas.LoadClientsInfor、frmClientsUpgrade.LoadClientsInfor
  -- frmFilesSendToServer.LoadClientsInfor
  -----------------------------------------------------------------------------
  Procedure Get_Client
  (
    Cur_Out   Out t_Refcur,
    工作站_In In zlClients.工作站%Type := Null
  );

  -----------------------------------------------------------------------------
  -- 功能：取zlClients的站点
  -- 调用列表：
  -- frmClientsParasSet.InitCard
  -----------------------------------------------------------------------------
  Procedure Get_Client_Station(Cur_Out Out t_Refcur);

  -----------------------------------------------------------------------------
  -- 功能：取方案号
  -- 调用列表：
  -- frmClientsParasSet.InitCard
  -----------------------------------------------------------------------------
  Procedure Get_Project_No(Cur_Out Out t_Refcur);

  -----------------------------------------------------------------------------
  -- 功能：取方案
  -- 调用列表：
  -- frmClientsParasSet.InitCard
  -----------------------------------------------------------------------------
  Procedure Get_Client_Scheme(Cur_Out Out t_Refcur);

  -----------------------------------------------------------------------------
  -- 功能：取恢复信息
  -- 调用列表：
  -- frmClientsParasSet.Load恢复方案、frmClientsParasSet.LoadScremeSet
  -----------------------------------------------------------------------------
  Procedure Get_Resile
  (
    Cur_Out   Out t_Refcur,
    方案号_In In Zlclientparaset.方案号%Type,
    类型_In   In Number := 0
  );

  -----------------------------------------------------------------------------
  -- 功能：取zldataMove数据
  -- 调用列表：
  -- frmDataMove.cmbSystem_Click
  -----------------------------------------------------------------------------
  Procedure Get_Zldatamove
  (
    Cur_Out Out t_Refcur,
    系统_In In zlDataMove.系统%Type
  );

  -----------------------------------------------------------------------------
  -- 功能：取日志数据
  -- 调用列表：
  -- FrmErrLog.RefreshData、FrmRunLog.RefreshData
  -----------------------------------------------------------------------------
  Procedure Get_Log
  (
    Cur_Out     Out t_Refcur,
    日志类型_In In Varchar2,
    Where_In    In Varchar2
  );

  -----------------------------------------------------------------------------
  -- 功能：取日志记录数
  -- 调用列表：
  -- FrmErrLog.DeleteExtra、FrmRunLog.DeleteExtra
  -----------------------------------------------------------------------------
  Procedure Get_Log_Count
  (
    Cur_Out     Out t_Refcur,
    日志类型_In In Varchar2
  );

  -----------------------------------------------------------------------------
  -- 功能：取zlfilesupgradeg数据
  -- 调用列表：
  -- frmFilesSet.intBillInfor
  -----------------------------------------------------------------------------
  Procedure Get_Zlfilesupgrade(Cur_Out Out t_Refcur);

  -----------------------------------------------------------------------------
  -- 功能：取非注册项目
  -- 调用列表：
  -- frmRegist.Form_Load
  -----------------------------------------------------------------------------
  Procedure Get_Not_Regist(Cur_Out Out t_Refcur);

  -----------------------------------------------------------------------------
  -- 功能：取参数值
  -- 调用列表：
  -- FrmRunOption.InitCons
  -----------------------------------------------------------------------------
  Procedure Get_Zloption
  (
    Cur_Out   Out t_Refcur,
    参数号_In In zlOptions.参数号%Type
  );

End b_Runmana;
/
-------------------------------------------------
-- 创建包体
-------------------------------------------------

CREATE OR REPLACE PACKAGE BODY ZLTOOLS.b_Expert IS

  -----------------------------------------------------------------------------
  -- 取提醒数据
  -----------------------------------------------------------------------------
  PROCEDURE Get_Notices
  (
    Cursor_Out OUT t_Refcur,
    序号_In    IN Zlnotices.序号%TYPE,
    系统_In    IN Zlreports.系统%TYPE := NULL
  ) IS
  BEGIN
    IF Nvl(序号_In, 0) <> 0 THEN
      -- frmNoticesEdit.ReadData 使用
      OPEN Cursor_Out FOR
        SELECT a.提醒内容, a.提醒条件, a.提醒报表, a.提醒声音, a.提醒窗口, a.开始时间, a.终止时间, a.检查周期,
               b.名称 AS 报表名称
        FROM Zlnotices a, Zlreports b
        WHERE a.提醒报表 = b.编号(+) AND a.序号 = 序号_In;
    ELSE
      -- cboSystem_Click 使用
      IF Nvl(系统_In, 0) = 0 THEN
        OPEN Cursor_Out FOR
          SELECT a.序号, a.提醒内容, a.提醒条件, a.提醒报表, a.提醒声音, a.提醒窗口, a.开始时间, a.终止时间, a.检查周期,
                 a.提醒周期, b.名称 AS 报表名称
          FROM Zlnotices a, Zlreports b
          WHERE a.提醒报表 = b.编号(+) AND a.系统 IS NULL;
      ELSE
        OPEN Cursor_Out FOR
          SELECT a.序号, a.提醒内容, a.提醒条件, a.提醒报表, a.提醒声音, a.提醒窗口, a.开始时间, a.终止时间, a.检查周期,
                 a.提醒周期, b.名称 AS 报表名称
          FROM Zlnotices a, Zlreports b
          WHERE a.提醒报表 = b.编号(+) AND a.系统 = 系统_In;
      END IF;
    END IF;

  END Get_Notices;

  -----------------------------------------------------------------------------
  -- 取提醒对像数据
  -----------------------------------------------------------------------------
  PROCEDURE Get_Noticeusr
  (
    Cursor_Out  OUT t_Refcur,
    提醒对象_In IN Zlnoticeusr.提醒对象%TYPE,
    提醒序号_In IN Zlnoticeusr.提醒序号%TYPE
  ) IS
  BEGIN
    IF Nvl(提醒对象_In, 0) = 0 THEN
      OPEN Cursor_Out FOR
        SELECT 1 FROM Zlnoticeusr WHERE Rownum < 2 AND 提醒序号 = 提醒序号_In;
    ELSE
      OPEN Cursor_Out FOR
        SELECT 对象名称 FROM Zlnoticeusr WHERE 提醒对象 = 提醒对象_In AND 提醒序号 = 提醒序号_In;
    END IF;
  END Get_Noticeusr;

  -----------------------------------------------------------------------------
  -- 取可以选择的提醒报表
  -----------------------------------------------------------------------------
  PROCEDURE Get_Noticereport
  (
    Cursor_Out OUT t_Refcur,
    系统_In    IN Zlreports.系统%TYPE
  ) IS
  BEGIN
    IF Nvl(系统_In, 0) = 0 THEN
      OPEN Cursor_Out FOR
        SELECT Id, 编号, 名称, 说明
        FROM Zlreports
        WHERE 编号 LIKE 'ZL%_REPORT_%' AND
              NOT (发布时间 IS NULL OR Trunc(发布时间) = To_Date('3000-01-01', 'yyyy-mm-dd')) AND Nvl(系统, 0) = 0;
    ELSE
      OPEN Cursor_Out FOR
        SELECT Id, 编号, 名称, 说明
        FROM Zlreports
        WHERE 编号 LIKE 'ZL%_REPORT_%' AND
              NOT (发布时间 IS NULL OR Trunc(发布时间) = To_Date('3000-01-01', 'yyyy-mm-dd')) AND 系统 = 系统_In;
    END IF;
  END Get_Noticereport;

  -----------------------------------------------------------------------------
  -- 在不同的系统间复制报表
  -----------------------------------------------------------------------------
  PROCEDURE Copy_Report
  (
    系统_In   IN Zlreports.系统%TYPE,
    新系统_In IN Zlreports.系统%TYPE
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
      SELECT Upper(所有者) AS 所有者 INTO v_Owner_Name FROM Zlsystems WHERE 编号 = Lngsys_In;
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

    v_Olduser := Upper(Sub_Owner_Name(系统_In));
    v_Newuser := Upper(Sub_Owner_Name(新系统_In));

    INSERT INTO Zlrptgroups
      (Id, 编号, 名称, 说明, 系统, 程序id, 发布时间)
      SELECT Id + n_Grpid, 编号, 名称, 说明, 新系统_In, 程序id, 发布时间 FROM Zlrptgroups WHERE 系统 = 系统_In;

    INSERT INTO Zlreports
      (Id, 编号, 名称, 说明, 密码, w, h, 纸张, 纸向, 进纸, 打印机, 票据, 系统, 程序id, 功能, 修改时间, 发布时间)
      SELECT Id + n_Rptid, 编号, 名称, 说明, 密码, w, h, 纸张, 纸向, 进纸, 打印机, 票据, 新系统_In, 程序id, 功能,
             修改时间, 发布时间
      FROM Zlreports
      WHERE 系统 = 系统_In;

    -- 插入zlRPTSub
    INSERT INTO Zlrptsubs
      (组id, 报表id, 序号, 功能)
      SELECT a.组id + n_Grpid, a.报表id + n_Rptid, a.序号, a.功能
      FROM Zlrptsubs a, Zlrptgroups b
      WHERE a.组id = b.Id AND b.系统 = 系统_In;

    -- 插入zlRPTFMTs
    INSERT INTO Zlrptfmts
      (报表id, 序号, 说明, 图样)
      SELECT a.报表id + n_Rptid, a.序号, a.说明, a.图样
      FROM Zlrptfmts a, Zlreports b
      WHERE a.报表id = b.Id AND b.系统 = 系统_In;

    -- 插入zlRPTItems
    INSERT INTO Zlrptitems
      (Id, 报表id, 格式号, 名称, 类型, 上级id, 序号, 参照, 性质, 内容, 表头, x, y, w, h, 行高, 对齐, 自调, 字体, 字号,
       粗体, 斜体, 下线, 前景, 背景, 边框, 排序, 格式, 汇总, 分栏, 网格, 系统)
      SELECT a.Id + n_Itemid, a.报表id + n_Rptid, a.格式号, a.名称, a.类型, a.上级id + n_Itemid, a.序号, a.参照, a.性质,
             a.内容, a.表头, a.x, a.y, a.w, a.h, a.行高, a.对齐, a.自调, a.字体, a.字号, a.粗体, a.斜体, a.下线, a.前景,
             a.背景, a.边框, a.排序, a.格式, a.汇总, a.分栏, a.网格, a.系统
      FROM Zlrptitems a, Zlreports b
      WHERE a.报表id = b.Id AND b.系统 = 系统_In;
    -- 插入zlRptDatas
    INSERT INTO Zlrptdatas
      (Id, 报表id, 名称, 字段, 对象, 类型)
      SELECT a.Id + n_Dataid, a.报表id + n_Rptid, a.名称, a.字段, a.对象, a.类型
      FROM Zlrptdatas a, Zlreports b
      WHERE a.报表id = b.Id AND b.系统 = 系统_In;
    -- 插入zlRPTSqls
    INSERT INTO Zlrptsqls
      (源id, 行号, 内容)
      SELECT a.源id + n_Dataid, a.行号, a.内容
      FROM Zlrptsqls a, Zlrptdatas b, Zlreports c
      WHERE a.源id = b.Id AND b.报表id = c.Id AND c.系统 = 系统_In;

    -- 插入zlRPTPars
    INSERT INTO Zlrptpars
      (源id, 组名, 序号, 名称, 类型, 缺省值, 格式, 值列表, 分类sql, 明细sql, 分类字段, 明细字段, 对象)
      SELECT a.源id + n_Dataid, a.组名, a.序号, a.名称, a.类型, a.缺省值, a.格式, a.值列表, a.分类sql, a.明细sql,
             a.分类字段, a.明细字段, a.对象
      FROM Zlrptpars a, Zlrptdatas b, Zlreports c
      WHERE a.源id = b.Id AND b.报表id = c.Id AND c.系统 = 系统_In;

    -- zlFunctions数据
    INSERT INTO Zlfunctions
      (系统, 函数号, 函数名, 中文名, 说明)
      SELECT 新系统_In, 函数号, 函数名, 中文名, 说明 FROM Zlfunctions WHERE 系统 = 系统_In;

    -- zlFuncPars数据
    INSERT INTO Zlfuncpars
      (系统, 函数号, 参数号, 参数名, 中文名, 类型, 缺省值, 值列表, 分类sql, 明细sql, 分类字段, 明细字段, 对象, 组名,
       递增否)
      SELECT 新系统_In, 函数号, 参数号, 参数名, 中文名, 类型, 缺省值, 值列表, 分类sql, 明细sql, 分类字段, 明细字段, 对象,
             组名, 递增否
      FROM Zlfuncpars
      WHERE 系统 = 系统_In;

    -- 重新设置数据源对象
    UPDATE Zlrptdatas
    SET 对象 = REPLACE(对象, v_Olduser || '.', v_Newuser || '.')
    WHERE Id IN (SELECT a.Id FROM Zlrptdatas a, Zlreports b WHERE a.报表id = b.Id AND b.系统 = 新系统_In);

    UPDATE Zlrptpars
    SET 对象 = REPLACE(对象, v_Olduser || '.', v_Newuser || '.')
    WHERE 源id IN (SELECT a.Id FROM Zlrptdatas a, Zlreports b WHERE a.报表id = b.Id AND b.系统 = 新系统_In);

    UPDATE Zlfuncpars SET 对象 = REPLACE(对象, v_Olduser || '.', v_Newuser || '.') WHERE 系统 = 新系统_In;

    COMMIT;
  EXCEPTION
    WHEN OTHERS THEN
      Zl_Errorcenter(SQLCODE, SQLERRM);
  END Copy_Report;

END b_Expert;
/

CREATE OR REPLACE Package Body ZLTOOLS.b_Loadandunload Is

  -----------------------------------------------------------------------------
  -- 功能：取有SysFiles表的文件名
  -----------------------------------------------------------------------------
  Procedure Get_Sysfile_Name
  (
    Cursor_Out Out t_Refcur,
    系统_In    In zlSysFiles.系统%Type,
    操作_In    In zlSysFiles.操作%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select 文件名 From zlSysFiles Where 系统 = 系统_In And 操作 = 操作_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Sysfile_Name;

  -----------------------------------------------------------------------------
  -- 功能：取共享系统名称
  -----------------------------------------------------------------------------
  Procedure Get_Share_Name
  (
    Cursor_Out Out t_Refcur,
    共享号_In  In zlSystems.编号%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select 名称 From zlSystems Start With 共享号 = 共享号_In Connect By Prior 编号 = 共享号;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Share_Name;

  -----------------------------------------------------------------------------
  -- 功能：取Oracle版本号
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
  -- 功能：CopyMenu
  -----------------------------------------------------------------------------
  Procedure Copy_Menu
  (
    系统_In   In zlMenus.系统%Type,
    新系统_In In zlMenus.系统%Type
  ) Is
    n_Menuid zlMenus.ID%Type;
  Begin
    Select Max(ID) Into n_Menuid From zlMenus;
    n_Menuid := Nvl(n_Menuid, 0) + 1;
    Insert Into zlMenus
      (组别, ID, 上级id, 标题, 短标题, 快键, 图标, 说明, 模块, 系统)
      Select 组别, ID + n_Menuid, 上级id + n_Menuid, 标题, 短标题, 快键, 图标, 说明, 模块, 新系统_In
      From zlMenus
      Where 系统 = 系统_In;
    Commit;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Copy_Menu;

  -----------------------------------------------------------------------------
  -- 功能：取ZlMenu数据
  -----------------------------------------------------------------------------
  Procedure Get_Menu_Tree
  (
    Cursor_Out Out t_Refcur,
    组别_In    In zlMenus.组别%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select ID, 上级id, 标题, 快键, 说明, 系统, 模块, 短标题, 图标, Level As 级数
      From zlMenus
      Start With 上级id Is Null And 组别 = 组别_In
      Connect By Prior ID = 上级id And 组别 = 组别_In
      Order By Level, ID;
  End Get_Menu_Tree;

  -----------------------------------------------------------------------------
  -- 功能：取ZlMenu数据
  -----------------------------------------------------------------------------
  Procedure Get_Menu_Group
  (
    Cursor_Out Out t_Refcur,
    组别_In    In zlMenus.组别%Type
  ) Is
  Begin
    If 组别_In Is Null Then
      -- frmMenu.FillMenuName
      Open Cursor_Out For
        Select Distinct 组别 From zlMenus;
    Else
      -- frmMenu.cmdNew_Click
      Open Cursor_Out For
        Select Count(*) As 数量 From zlMenus Where 组别 = 组别_In;
    End If;
  End Get_Menu_Group;

  -----------------------------------------------------------------------------
  -- 功能：取模块
  -----------------------------------------------------------------------------
  Procedure Get_Module
  (
    Cursor_Out Out t_Refcur,
    系统_In    In zlComponent.系统%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select P.序号, P.标题, C.名称 As 部件
      From zlPrograms P, zlComponent C
      Where Upper(P.部件) = Upper(C.部件) And C.系统 = 系统_In And P.系统 = 系统_In
      Order By C.名称, P.序号;
  End Get_Module;

  -----------------------------------------------------------------------------
  -- 功能：取功能或排列，说明
  -----------------------------------------------------------------------------
  Procedure Get_Function
  (
    Cursor_Out Out t_Refcur,
    系统_In    In zlProgFuncs.系统%Type,
    序号_In    In zlProgFuncs.序号%Type,
    功能_In    In zlProgFuncs.功能%Type := Null
  ) Is
  Begin
    If Nvl(功能_In, '空') = '空' Then
      Open Cursor_Out For
        Select 功能 From zlProgFuncs Where 系统 = 系统_In And 序号 = 序号_In Order By Nvl(排列, 0);
    Else
      Open Cursor_Out For
        Select 排列, 说明 From zlProgFuncs Where 系统 = 系统_In And 序号 = 序号_In And 功能 = 功能_In;
    End If;
  End Get_Function;

  -----------------------------------------------------------------------------
  -- 功能：取表权限
  -----------------------------------------------------------------------------
  Procedure Get_Impower
  (
    Cursor_Out Out t_Refcur,
    系统_In    In zlProgPrivs.系统%Type,
    序号_In    In zlProgPrivs.序号%Type,
    功能_In    In zlProgPrivs.功能%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select 对象, Sum(Decode(权限, 'SELECT', 1, 0)) As "SELECT", Sum(Decode(权限, 'UPDATE', 1, 0)) As "UPDATE",
             Sum(Decode(权限, 'INSERT', 1, 0)) As "INSERT", Sum(Decode(权限, 'DELETE', 1, 0)) As "DELETE",
             Sum(Decode(权限, 'EXECUTE', 1, 0)) As "EXECUTE"
      From zlProgPrivs
      Where 系统 = 系统_In And 序号 = 序号_In And 功能 = 功能_In
      Group By 对象
      Order By 对象;
  End Get_Impower;

  -----------------------------------------------------------------------------
  -- 功能：得到角色能访问的导航台工具
  -----------------------------------------------------------------------------
  Procedure Get_Role_Tools
  (
    Cursor_Out Out t_Refcur,
    角色_In    In zlRoleGrant.角色%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select P.序号, P.标题, P.说明, R.功能
      From zlRoleGrant R, zlPrograms P
      Where R.系统 Is Null And P.序号 = R.序号 And R.角色 = 角色_In And P.系统 Is Null And P.序号 < 100 And
            P.部件 Is Null
      Order By P.序号;
  End Get_Role_Tools;

  -----------------------------------------------------------------------------
  -- 功能：得到以前的权限
  -----------------------------------------------------------------------------
  Procedure Get_Role_Grant
  (
    Curgrand_Out    Out t_Refcur,
    Curprivs_Out    Out t_Refcur,
    Curfuncpars_Out Out t_Refcur,
    角色_In         In zlRoleGrant.角色%Type
  ) Is
  Begin
    Open Curgrand_Out For
      Select Nvl(系统, 0) As 系统, 序号, 功能 From zlRoleGrant Where 角色 = 角色_In;
    Open Curprivs_Out For
      Select Nvl(系统, 0) As 系统, 序号, 功能, 所有者, 权限, 对象 From zlProgPrivs;
    Open Curfuncpars_Out For
      Select P.系统, F.函数名, P.对象
      From zlFuncPars P, zlFunctions F
      Where P.系统 = F.系统 And P.函数号 = F.函数号 And P.对象 Is Not Null;
  End Get_Role_Grant;

  -----------------------------------------------------------------------------
  -- 功能：FillFunc
  -----------------------------------------------------------------------------
  Procedure Get_Zlprogfunc
  (
    Cursor_Out Out t_Refcur,
    系统_In    In zlProgFuncs.系统%Type,
    序号_In    In zlProgFuncs.序号%Type
  ) Is
  Begin
    If Nvl(系统_In, 0) = 0 Then
      Open Cursor_Out For
        Select 功能, 排列, 说明 From zlProgFuncs Where 系统 Is Null And 序号 = 序号_In And 功能 <> '基本';
    Else
      Open Cursor_Out For
        Select A.功能, A.排列, A.说明
        From zlProgFuncs A, Zlregfunc B
        Where (A.系统 / 100) = B.系统 And A.序号 = B.序号 And A.功能 = B.功能 And A.系统 = 系统_In And A.序号 = 序号_In And
              A.功能 <> '基本';
    End If;
  End Get_Zlprogfunc;

  -----------------------------------------------------------------------------
  -- 功能：是所有角色对应的模块
  -----------------------------------------------------------------------------
  Procedure Get_All_Module(Cursor_Out Out t_Refcur) Is
  Begin
    Open Cursor_Out For
      Select A.角色, A.序号, A.功能, B.标题, B.说明
      From zlRoleGrant A, zlPrograms B
      Where A.序号 = B.序号 And Nvl(A.系统, 0) = Nvl(B.系统, 0)
      Order By A.角色, A.序号;
  End Get_All_Module;

End b_Popedom;
/

CREATE OR REPLACE Package Body ZLTOOLS.b_Public Is

  -----------------------------------------------------------------------------
  -- 功能：取系统日期
  -----------------------------------------------------------------------------
  Procedure Get_Current_Date(Cursor_Out Out t_Refcur) Is
  Begin
    Open Cursor_Out For
      Select Sysdate As 日期 From Dual;
  End Get_Current_Date;

  -----------------------------------------------------------------------------
  -- 功能：删除错误日志或运行日志
  -----------------------------------------------------------------------------
  Procedure Delete_All_Log(Runtimelog_In In Number := 0) Is
    n_Count Number;
    n_Loop  Number;
  Begin
    If Runtimelog_In = 1 Then
      Select Count(进入时间) Into n_Count From zlDiaryLog;
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
      Select Count(时间) Into n_Count From zlErrorLog;
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
  -- 功能：删除当前运行日志
  -----------------------------------------------------------------------------
  Procedure Delete_Diarylog
  (
    会话号_In   Number,
    用户名_In   Varchar2,
    工作站_In   Varchar2,
    部件名_In   Varchar2,
    工作内容_In Varchar2,
    进入时间_In Date
  ) Is
  Begin
    Delete zlDiaryLog
    Where 会话号 = 会话号_In And 用户名 = 用户名_In And 工作站 = 工作站_In And 部件名 = 部件名_In And
          工作内容 = 工作内容_In And 进入时间 = 进入时间_In;
    Commit;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Delete_Diarylog;

  -----------------------------------------------------------------------------
  -- 功能：删除当前错误日志
  -----------------------------------------------------------------------------
  Procedure Delete_Errorlog
  (
    会话号_In   Number,
    用户名_In   Varchar2,
    工作站_In   Varchar2,
    类型_In     Number,
    错误序号_In Number,
    时间_In     Date
  ) Is
  Begin
    Delete zlErrorLog
    Where 会话号 = 会话号_In And 用户名 = 用户名_In And 工作站 = 工作站_In And 类型 = 类型_In And
          错误序号 = 错误序号_In And 时间 = 时间_In;
    Commit;
  
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Delete_Errorlog;

  -----------------------------------------------------------------------------
  -- 功能：取注册码
  -----------------------------------------------------------------------------
  Procedure Get_Regcode(Cursor_Out Out t_Refcur) Is
  Begin
    Open Cursor_Out For
      Select 内容 From zlRegInfo Where 项目 = '注册码' Or 项目 = '授权证章' Order By 行号;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Regcode;

    -----------------------------------------------------------------------------
  -- 功能：取版本号
  -----------------------------------------------------------------------------
  Procedure Get_Ver(Cursor_Out Out t_Refcur) Is
  Begin
    Open Cursor_Out For
      Select 内容 From zlRegInfo Where 项目 = '版本号';
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Ver;

  -----------------------------------------------------------------------------
  -- 功能：更新版本号
  -----------------------------------------------------------------------------
  Procedure Update_Ver(Verstring_In In Varchar2) Is
  Begin
    Update zlRegInfo Set 内容 = Verstring_In Where 项目 = '版本号';
    If Sql%NotFound Then
      Insert Into zlRegInfo (项目, 行号, 内容) Values ('版本号', 1, Verstring_In);
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Update_Ver;

  -----------------------------------------------------------------------------
  -- 功能：取得系统所有者名称
  -----------------------------------------------------------------------------
  Procedure Get_Owner_Name
  (
    Cursor_Out Out t_Refcur,
    编号_In    In zlSystems.编号%Type := 0
  ) Is
  Begin
    Open Cursor_Out For
      Select Upper(所有者) As 所有者 From zlSystems Where 编号 = 编号_In;
  End Get_Owner_Name;

  -----------------------------------------------------------------------------
  -- 功能：取注册表中信息
  -----------------------------------------------------------------------------
  Procedure Get_Reginfo
  (
    Cursor_Out Out t_Refcur,
    项目_In    In zlRegInfo.项目%Type := Null
  ) Is
  Begin
    If Trim(Nvl(项目_In, '空')) = '空' Then
      Open Cursor_Out For
        Select * From zlRegInfo;
    Else
      Open Cursor_Out For
        Select 内容 From zlRegInfo Where 项目 = 项目_In Order By 行号;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Reginfo;

  -----------------------------------------------------------------------------
  -- 功能：取zlGetSvrToolsg数据
  -----------------------------------------------------------------------------
  Procedure Get_Zlsvrtools(Cursor_Out Out t_Refcur) Is
  Begin
    Open Cursor_Out For
      Select * From zlSvrTools Start With 上级 Is Null Connect By Prior 编号 = 上级 Order By Level, 编号;
  End Get_Zlsvrtools;

  -----------------------------------------------------------------------------
  -- 功能：取已安装系统清单
  -----------------------------------------------------------------------------
  Procedure Get_Zlsystems
  (
    Cursor_Out Out t_Refcur,
    所有者_In  In zlSystems.所有者%Type := Null
  ) Is
  Begin
    If Nvl(所有者_In, '空') = '空' Then
      Open Cursor_Out For
        Select * From zlSystems Order By 编号;
    Else
      Open Cursor_Out For
        Select * From zlSystems Where Upper(所有者) = Upper(所有者_In) Order By 编号;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Zlsystems;

End b_Public;
/

CREATE OR REPLACE Package Body ZLTOOLS.b_Runmana Is
  -----------------------------------------------------------------------------
  -- 功能：取ZlAutoJob序列号
  -----------------------------------------------------------------------------
  Procedure Get_Job_Number
  (
    Cursor_Out Out t_Refcur,
    系统_In    In Number
  ) Is
  Begin
    Open Cursor_Out For
      Select 序号 + 1 As 序号
      From zlAutoJobs
      Where Nvl(系统, 0) = 系统_In And 类型 = 3 And
            序号 + 1 Not In (Select 序号 From zlAutoJobs Where Nvl(系统, 0) = 系统_In And 类型 = 3);
  End Get_Job_Number;

  -----------------------------------------------------------------------------
  -- 功能：取ZlDataMove描述
  -----------------------------------------------------------------------------
  Procedure Get_Depict
  (
    Cursor_Out Out t_Refcur,
    系统_In    In zlDataMove.系统%Type,
    组号_In    In zlDataMove.组号%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select 转出描述 From zlDataMove Where Nvl(系统, 0) = 系统_In And 组号 = 组号_In;
  End Get_Depict;

  -----------------------------------------------------------------------------
  -- 功能：取zlClients的MAX IP
  -----------------------------------------------------------------------------
  Procedure Get_Client_Maxip(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select Max(Ip) As Ip From zlClients;
  End Get_Client_Maxip;

  -----------------------------------------------------------------------------
  -- 功能：取zlClients的记录
  -----------------------------------------------------------------------------
  Procedure Get_Client
  (
    Cur_Out   Out t_Refcur,
    工作站_In In zlClients.工作站%Type := Null
  ) Is
    v_Sql Varchar2(1000);
  Begin
    If Nvl(工作站_In, '空') = '空' Then
      v_Sql := 'Select a.Ip, a.工作站, a.Cpu, a.内存, a.硬盘, a.操作系统, a.部门, a.用途, a.说明, a.升级标志, a.禁止使用,
							 a.连接数, Decode(b.Terminal, Null, 0, 1) As 状态, a.收集标志
				From Zlclients a, (Select Distinct Terminal From V$session) b
				Where Upper(a.工作站) = Upper(b.Terminal(+))
				Order By a.Ip';
      Open Cur_Out For v_Sql;
    Else
      Open Cur_Out For
        Select Ip, 工作站, Cpu, 内存, 硬盘, 操作系统, 部门, 用途, 说明, 升级标志, 禁止使用, 连接数
        From zlClients
        Where Upper(工作站) = 工作站_In;
    End If;
  End Get_Client;

  -----------------------------------------------------------------------------
  -- 功能：取zlClients的站点
  -----------------------------------------------------------------------------
  Procedure Get_Client_Station(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select Distinct Upper(工作站) || '[' || Ip || ']' As 站点, Upper(工作站) 工作站 From zlClients;
  End Get_Client_Station;

  -----------------------------------------------------------------------------
  -- 功能：取方案号
  -----------------------------------------------------------------------------
  Procedure Get_Project_No(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select 方案号 From Zlclientparaset Where Rownum = 1;
  End Get_Project_No;

  -----------------------------------------------------------------------------
  -- 功能：取方案
  -----------------------------------------------------------------------------
  Procedure Get_Client_Scheme(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select 方案号, 方案号 || '-' || 方案名称 As 方案名称, 方案描述, 工作站, 用户名 From Zlclientscheme;
  End Get_Client_Scheme;

  -----------------------------------------------------------------------------
  -- 功能：取恢复信息
  -----------------------------------------------------------------------------
  Procedure Get_Resile
  (
    Cur_Out   Out t_Refcur,
    方案号_In In Zlclientparaset.方案号%Type,
    类型_In   In Number := 0
  ) Is
  Begin
    If 类型_In = 0 Then
      Open Cur_Out For
        Select Distinct A.工作站 || Decode(M.工作站, Null, ' ', '[' || M.Ip || ']') As 工作站, A.用户名, A.恢复标志,
                        '[' || B.方案号 || ']' || B.方案名称 As 方案名称
        From Zlclientparaset A, Zlclientscheme B, zlClients M
        Where A.方案号 = B.方案号 And A.工作站 = M.工作站(+) And A.方案号 = 方案号_In;
    End If;
  
    If 类型_In = 1 Then
      Open Cur_Out For
        Select Distinct Upper(工作站) 工作站, Min(恢复标志) 恢复标志
        From Zlclientparaset A
        Where A.方案号 = 方案号_In
        Group By 工作站;
    End If;
  
    If 类型_In = 2 Then
      Open Cur_Out For
        Select Distinct Upper(用户名) 用户名, Max(工作站) 工作站, Min(Decode(恢复标志, 2, 0, 恢复标志)) 恢复标志
        From Zlclientparaset A
        Where A.方案号 = 方案号_In
        Group By 用户名
        Order By 用户名;
    End If;
  
  End Get_Resile;

  -----------------------------------------------------------------------------
  -- 功能：取zldataMove数据
  -----------------------------------------------------------------------------
  Procedure Get_Zldatamove
  (
    Cur_Out Out t_Refcur,
    系统_In In zlDataMove.系统%Type
  ) Is
  Begin
    Open Cur_Out For
      Select 组号, 组名, 说明, 日期字段, 转出描述, 上次日期 From zlDataMove Where 系统 = 系统_In Order By 组号;
  End Get_Zldatamove;

  -----------------------------------------------------------------------------
  -- 功能：取日志数据
  -----------------------------------------------------------------------------
  Procedure Get_Log
  (
    Cur_Out     Out t_Refcur,
    日志类型_In In Varchar2,
    Where_In    In Varchar2
  ) Is
    v_Sql Varchar2(1000);
  Begin
    If 日志类型_In = '错误日志' Then
      v_Sql := 'Select 会话号,工作站,用户名,错误序号,错误信息,To_char(时间,''yyyy-MM-dd hh24:mi:ss'') 时间
					 ,Decode(类型,1,''存储过程错误'',2,''数据联结层错误'',''应用程序层错误'') 错误类型
						From ZlErrorLog Where ' || Where_In;
      Open Cur_Out For v_Sql;
    End If;
    If 日志类型_In = '运行日志' Then
      v_Sql := 'Select 会话号,工作站,用户名,部件名,工作内容,To_char(进入时间,''yyyy-MM-dd hh24:mi:ss'') 进入时间
								 ,To_char(退出时间,''yyyy-MM-dd hh24:mi:ss'') 退出时间,Decode(退出原因,1,''正常退出'',''异常退出'') 退出原因
									From ZlDiaryLog Where ' || Where_In;
      Open Cur_Out For v_Sql;
    End If;
  End Get_Log;

  -----------------------------------------------------------------------------
  -- 功能：取日志记录数
  -----------------------------------------------------------------------------
  Procedure Get_Log_Count
  (
    Cur_Out     Out t_Refcur,
    日志类型_In In Varchar2
  ) Is
  Begin
    If 日志类型_In = '错误日志' Then
      Open Cur_Out For
        Select Count(*) 数量
        From zlErrorLog
        Union All
        Select Nvl(To_Number(参数值), 0) From zlOptions Where 参数号 = 4;
    End If;
    If 日志类型_In = '运行日志' Then
      Open Cur_Out For
        Select Count(*) 数量
        From zlDiaryLog
        Union All
        Select Nvl(To_Number(参数值), 0) From zlOptions Where 参数号 = 2;
    
    End If;
  End Get_Log_Count;

  -----------------------------------------------------------------------------
  -- 功能：取zlfilesupgradeg数据
  -----------------------------------------------------------------------------
  Procedure Get_Zlfilesupgrade(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select A.序号, A.文件名, A.版本号, A.修改日期, B.名称 As 说明
      From zlFilesUpgrade A, zlComponent B
      Where Upper(A.文件名) = Upper(B.部件(+))
      Order By A.序号;
  End Get_Zlfilesupgrade;

  -----------------------------------------------------------------------------
  -- 功能：取非注册项目
  -----------------------------------------------------------------------------
  Procedure Get_Not_Regist(Cur_Out Out t_Refcur) Is
  Begin
    Open Cur_Out For
      Select 项目, 内容
      From zlRegInfo
      Where 项目 Not In ('发行码', '版本号', '服务器目录', '访问用户', '访问密码', '收集目录', '收集类型', '注册码',
             '授权证章', '授权工具', '授权邮戳');
  End Get_Not_Regist;

  -----------------------------------------------------------------------------
  -- 功能：取参数值
  -----------------------------------------------------------------------------
  Procedure Get_Zloption
  (
    Cur_Out   Out t_Refcur,
    参数号_In In zlOptions.参数号%Type
  ) Is
  Begin
    Open Cur_Out For
      Select Nvl(参数值, 缺省值) Option_Value From zlOptions Where 参数号 = 参数号_In;
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
	v_Stdstr  Varchar2(50) := '芭擦搭蛾发噶哈击-喀垃妈拿哦啪期然撒塌挖-挖昔压匝';
	v_Chara   Varchar2(2000) := '吖锕錒嗄锿鎄捱嗳躷﨟譪霭靄砹嗌嫒暧瑷賹鴱譺鑀鱫靉桉庵谙誝鞌諳闇鮟鵪韽鶕雸埯铵隌揞銨犴豻貋錌黯醠坳軪隞嗷廒獒遨謷鳌鏖鰲鷔鼇拗媪镺岙骜謸鏊鹌聱螯';
	v_Charb   Varchar2(2000) := '岜粑釟豝鲃魞茇釛菝軷颰魃鼥钯鈀跁鲅鮁灞掰捭呗鞁贁韛頒辬阪坂钣舨鈑魬闆鉡靽辦浜邫鞤蒡謗鎊勹孢煲龅闁齙葆飹飽鳵鴇賲靌趵铇鉋靤鮑鑤陂鵯鉳貝邶悖鄁軰碚蓓誖輩鋇鞴鐾贲逩賁锛錛畚坌輽閍嘣鞛逬跰甏镚鏰豍鲾鎞鵖鰏荸匕吡妣秕俾舭貏诐邲畀哔荜狴铋婢庳萆閇閉弼愎貱赑滗跸鉍閟飶鄪嬖薜鮅濞蹕鞞髀璧鄨襞鏎鞸韠躃躄魓贔鐴鷝鷩鼊砭煸邉鍽鳊邊鯾鯿匾貶碥鴘弁忭汴苄釆拚缏閞辡頨辧辪辮辯變灬杓飑髟颮骠麃镖飙飚颷謤贆鏢镳飆飇飈飊鑣婊諘錶鳔鰾鱉鼈龞蹩邠傧缤槟豩賓賔镔豳霦鑌顮殡膑髌鬓鬢冫鋲邴陃禀鈵鉼鞆餅餠摒誁鮩靐饽啵鉢餑蹳鱍孛郣亳钹鈸鉑鲌踣鋍镈鮊豰鎛鵓礴鑮跛簸擘檗譒逋钸晡鈽誧餔轐醭卟鳪鵏鸔钚瓿鈈踄郶鹁瘢癍裱褙褊篦箅筚笾蝙褓裨窆瘭鸨鹎';
	v_Charc   Varchar2(2000) := '嚓礤遪財跴飡骖黪黲粲璨謲伧鸧鶬鑶賶嘈漕艚鏪艹鄵鼜恻岑涔噌杈馇銟锸鍤鎈猹靫槎檫蹅镲鑔汊姹钗釵侪辿觇鋓婵孱禅誗鋋廛潺鄽镡酁躔镵讒鑱谄蒇諂閳冁醦譂鏟闡讇忏羼韂顫伥娼菖阊锠錩閶鲳鯧鼚苌長镸徜嫦鋿鲿鏛鱨昶惝氅鋹怅鬯誯韔怊焯鈔晁鄛鼌轈鼂謿麨車砗屮坼迠頙抻琛嗔諃賝謓迧宸陳谌軙鈂霃諶麎鷐趻碜踸贂龀趂榇齓齔谶讖阷柽铛赪靗瞠赬頳鏳鏿鐺丞枨郕埕铖塍誠酲鋮鯎哧眵嗤媸誺鴟鵄魑齝麶黐茌赿貾遅趍遟墀踟遲謘豉鉹齒彳叱饬迣敕啻飭傺跮鉓雴遫銐趩鶒鷘忡茺舂憧艟蹖隀铳銃俦帱惆酧雔雠躊醻讎讐醜魗遚樗貙齣刍豠趎鉏鋤雛蹰鶵躕杵楮齭齼亍怵绌豖鄐踀閦諔憷黜搋啜嘬踹巛氚舡遄輲舛钏釧賗鶨闖怆龡陲棰槌錘鎚顀輴鰆鶞陙莼醕錞鯙鶉賰踳踔辶辵逴辍酫趠輟龊齪鑡齱呲趀祠茈辝鈶糍辤飺餈鴜辭鶿鷀跐賜苁枞骢璁鏦淙琮誴賨賩謥楱腠辏輳麁麄麤徂殂猝酢蔟誎趗踧蹙鼀蹴蹵顣汆撺镩躥鑹爨榱鏙璀趡啐悴萃毳顇邨踆忖遳蹉醝嵯矬鹾鹺齹脞厝逪锉銼錯澶瘥隹篪笞蚩虿耖皴褫褚裎衩瘳蟾螬螭蝽蜍蛏瘛痤鸱骣鹚鹑膪';
	v_Chard   Varchar2(2000) := '哒耷嗒鎝迏迖妲怛沓逹達跶靼鞑鎉躂鐽韃龖龘呔轪岱甙绐迨玳軑埭軚貸軩鮘鴏黛蹛霴黱靆眈躭酖殚鄲頕儋黕啖萏誕澹鴠贉霮谠譡黨讜凼宕砀菪逿雼趤闣刂叨忉氘釖鱽魛陦﨩隝隯焘軇纛锝鍀豋噔簦戥鄧隥嶝磴镫鐙羝隄趆嘀镝鍉鞮鏑籴荻觌靮頔魡豴鸐氐诋邸阺坻柢砥軧骶鯳娣逓谛釱棣睇遞鉪碲遰諦踶嗲蹎巅顚顛齻踮點阽坫玷钿鈿電簟貂鳭鮉鲷鼦鯛鵰釣铞鈟銱雿調鋽鑃垤喋堞揲趃牒镻諜蹀鲽鰈仃玎酊釘靪頂鼑鐤飣啶腚碇錠顁铥颩銩咚岽氡鮗鼕鯟鶇鶫諌垌峒胨迵胴硐霘蔸阧钭﨣鈄郖鬥酘閗鬦鋀餖闘鬪鬬鬭阇嘟醏闍渎椟牍読錖黩讀豄贕韣髑鑟韇韥黷讟賭芏靯鍍鍴椴煅鍛躖頧鴭鐜怼陮隊碓憝镦譈鐓譵礅蹾盹趸躉沌炖砘逇鈍頓遯踲咄铎鈬踱鮵鐸哚缍趓躱軃鬌沲陊陏跢跥飿鵽瓞簖篼箪蚪聃耵耋褡裰裆窦癫癜瘅笪笃蠹疸疔鸫';
	v_Chare   Varchar2(2000) := '屙迗莪鈋锇誐鋨頟魤額鵝鵞譌婀鵈阨呃苊阸轭垩谔軛阏愕萼豟軶遌腭锷遻頞餓噩諤閼鍔鳄顎鰐鶚讍鑩齶鱷蒽摁鞥陑輀鲕隭鮞鴯轜迩珥铒鉺餌邇趰佴貮貳鸸颚鹗';
	v_Charf   Varchar2(2000) := '醗垡閥砝鍅幡轓颿飜鱕釩蕃燔蹯蘩鐇鷭辺畈軓梵販軬飯飰匚邡枋钫趽鈁錺鴋鲂魴彷舫鶭妃飛绯扉靟霏鲱鯡飝淝腓悱斐榧翡誹狒費镄鼣鐨靅玢躮鈖雰棼隫魵鳻豮鼢鼖豶轒鐼黂黺偾鲼瀵鱝沣砜風葑鄷鋒豐鎽鏠酆靊飌麷唪諷俸赗鳯鳳鴌賵雬鴀邞呋趺酜麸稃跗鈇鄜豧鳺麩麬麱凫孚芙芾怫绂绋苻祓罘茯郛韨鳬砩莩匐桴艴菔﨓鉘鉜颫鳧韍幞鴔諨踾輻鮄鮲黻鵩鶝呒拊郙釡滏輔鬴黼阝驸負陚鲋赙賦輹鮒賻鍑鍢鳆鰒馥篚蚨蜚蝠缶蝮蜉痱';
	v_Charg   Varchar2(2000) := '旮伽钆尜釓錷尕尬魀郂陔垓赅隑豥賅賌鎅丐鈣戤迀坩泔苷酐尴鳱魐秊澉趕橄擀鳡鱤旰矸绀淦贛阬罡釭鋼鎠戆槔睾韟鷎鼛鷱杲缟槁藁鎬诰郜锆誥鋯圪纥閤鴐鴚謌鴿鎶鬲嗝塥搿膈閣镉鞈韐骼諽輵鮯鎘韚轕鞷鰪哿舸硌鉻哏亘艮茛赓鹒賡鶊郠哽绠鲠鯁肱觥躳龏龔廾珙輁鞏貢贑佝缑鈎鉤鞲韝岣枸豿诟媾彀遘雊觏購轱菰觚軱軲酤毂鈲鮕鴣轂鹘鶻汩诂牯罟逧钴鈷鼔嘏臌瞽鵠崮梏牿锢頋錮鲴鯝顧胍颪趏銽颳鴰呱卦诖倌関闗鳏關鰥鱞輨錧躀鳤掼涫貫遦盥雚鏆鑵鸛鱹咣桄胱輄銧黆犷妫邽郌閨鲑鮭龜鬶鬹蘒宄庋匦陒軌晷刿炅貴鳜鞼鱖鱥丨衮绲磙輥鲧鮌鯀謴呙埚崞鈛鍋帼掴虢馘猓椁輠餜鐹過簋篝筻笱蝈蜾蛄蚣虼聒矜袼疳鹳鹄痼鸹鸪皈';
	v_Charh   Varchar2(2000) := '铪鉿嗨胲酼醢餀頇谽魽鼾邗晗焓鋡韓豃鬫闬菡釬閈撖銲鋎頷顄譀雗瀚鶾魧迒绗貥頏沆蒿嚆薅嗥濠譹昊灏顥鰝诃嗬劾郃曷盍龁貈鉌阖鲄閡鹖麧頜翮魺闔鞨齕鶡鑉龢隺賀壑鶴齃靍靎鸖靏黒鞎桁珩鸻鴴鵆蘅鑅訇軣谾薨輷鍧轟闳泓荭谹鈜閎谼鉷鞃魟鋐蕻霐黉霟鴻黌讧閧銾闀闂鬨齁銗糇骺鍭鯸郈後逅鄇堠豞鲎鲘鮜鱟烀轷唿惚軤雽滹雐謼囫斛猢煳槲魱醐頶觳鍸鬍鰗鶘鶦鶮浒琥錿鯱冱岵怙戽祜扈鄠鳸鍙護鳠韄頀鱯鸌誮錵骅铧鋘譁鏵鷨桦諙諣黊踝鴅鵍酄獾貛讙郇洹萑雈貆锾阛寰缳還豲鍰镮鹮轘闤鐶鬟輐奂浣逭漶鲩擐鯇鯶鰀肓隍黃徨湟遑潢锽璜諻鍠鳇趪韹鐄鰉鱑鷬謊鎤诙咴晖珲豗隓輝麾隳鰴洄茴迴逥鮰譭哕浍荟恚桧彗喙缋阓賄誨蕙諱頮譓譮鏸闠鐬靧韢譿顪阍閽馄餛轋鼲诨溷諢锪劐鍃攉邩钬鈥夥閄貨嗀謋雘镬嚯藿鑊靃皓篌篁蚝虺颢颔颌颃顸耠癀笏蠖蟪蟥蚵蚶瘊鹱鹕瓠';
	v_Charj   Varchar2(2000) := '丌叽乩玑芨矶咭剞唧屐飢嵇犄赍跻鳮銈畿賫躸齑墼錤隮羁賷鄿雞譏韲鶏譤鐖躋鞿鷄齎鑇鑙齏鸄岌亟佶郆﨤谻戢殛楫蒺趌銡蕺踖鞊鹡輯蹐鍓轚鏶霵鶺鷑躤雦雧掎鱾戟嵴麂魢彐芰哜洎觊偈跡際暨誋跽霁鲚諅鲫髻鮆蹟鯽鵋齌骥鯚鱀霽鰶鰿鱭迦浃珈袈葭跏鉫镓豭貑鎵麚岬郏郟恝戛铗跲餄鋏頬頰鴶鵊胛賈鉀戋菅豜湔犍間靬搛缣蒹豣鲣鳽鋻鞬麉鞯鳒鵳鰔譼鰜鶼韀鰹鑯韉囝枧趼睑锏谫戬翦謇蹇謭鬋鰎鹸鐗鐧鹻譾鹼牮谏釼楗毽腱跈閒賎僭諓賤趝踐踺諫鍵餞鍳鏩轞鑑鑒鑬鑳茳豇缰鳉礓韁鱂講顜洚绛犟醤糨醬謽艽姣茭跤僬鲛鮫鵁轇鐎鷦鷮佼挢湫敫賋踋鉸餃徼鵤譑鱎峤較噍趭轎醮譥釂階喈嗟鞂鶛卩孑讦诘拮迼桀婕鉣魝碣鲒羯誱踕頡鍻鮚飷骱誡魪钅釒鹶黅卺堇廑馑槿瑾錦謹妗荩赆進缙觐噤賮贐齽泾旌菁腈鵛鯨鶁鶄麖鼱麠阱刭肼儆憬頸弪迳胫逕婧靓獍誩踁頚靚靜鏡冂扃迥逈颎顈赳阄啾鳩鬏鬮镹韮柩桕僦鯦麔齨鷲苴陱掬椐琚趄跔锔雎諊踘鋦鮈鴡鞫鶋郹輂跼趜躹閰橘鵙蹫鵴鶪鼰鼳莒榉榘龃﨔踽齟讵苣邭钜倨犋跙鉅飓豦屦鮔遽鋸颶瞿貗躆醵鐻涓鋑鋗镌鎸鵑鐫蠲锩錈桊狷隽鄄雋飬餋噘孓珏崛桷觖赽趹逫厥趉鈌劂谲獗蕨鴂鴃噱橛镼镢譎蹶蹷鶌矍鐍鐝爝鷢龣貜躩钁軍鈞銁銞鲪麇鍕鮶麏麕陖捃餕鵔鵕鵘稷鹣疖瘕筠笈蛟蛱蚧虮颉皲裾裥袷衿窭瘠痂鹫笳笕笄耩鹪鸠皎';
	v_Chark   Varchar2(2000) := '咔佧胩鉲锎開鐦剀垲恺闿铠蒈輆锴鍇鎧闓颽忾鎎龛戡龕侃莰輡轁顑轗阚瞰闞躿鏮鱇伉邟闶钪鈧閌尻栲铐犒銬鲓鮳鯌珂轲趷钶軻稞鈳瞌頦醘顆髁岢恪氪骒缂嗑溘锞課錁豤貇錹铿誙銵鍞鏗倥崆躻躼錓鵼鞚芤眍叩釦蔻鷇刳郀堀跍骷鮬绔喾誇侉銙蒯郐哙狯脍鲙鄶鱠髋鑧诓邼哐誆軭诳軖軠誑鵟夼邝圹纩贶貺軦鉱鋛鄺黋鑛悝闚顝逵鄈頄馗喹揆暌睽頯鍨鍷夔躨跬頍蹞匮喟愦蒉謉鐀鑎琨锟髡鹍醌錕鲲鯤鵾鶤悃阃閫閸栝頢闊鞟韕霩鞹鬠疴蛞篑箜筘蝌蝰颏裉窠聩';
	v_Charl   Varchar2(2000) := '邋旯砬剌辢鬎镴鯻鑞鞡崃徕涞郲逨铼錸鯠鶆麳赉睐賚濑賴頼顂鵣籁岚斓镧闌譋讕躝鑭钄韊榄漤罱醂啷郎郞莨稂锒郒躴鋃鎯阆誏閬蒗唠崂铹醪鐒顟栳铑銠鮱轑軂仂阞叻泐韷鳓鰳餎嫘缧檑羸鐳轠鑘靁鱩鼺诔誄讄鑸鸓酹銇頛頪錑颣類嘞塄踜愣骊喱缡蓠嫠貍鋫鲡罹錅謧醨藜邌釐離鯏鏫鯬鵹黧鑗鱺鸝礼俚娌逦锂豊裏鋰澧鯉醴鳢邐鱧呖坜苈戾枥俪栎赲轹郦猁砺莅唳粝詈跞雳溧鉝鳨隷鴗隸麗酈鷅麜躒轢讈轣靂鱱靋奁連鲢濂臁蹥謰鎌譧鬑鐮鰱琏蔹鄻娈殓楝潋錬鍊鏈鰊凉椋辌墚踉輬魉魎輌諒輛鍄蹽嘹寮獠缭遼豂賿蹘鐐飉鷯钌釕鄝蓼镽尥咧冽洌迾埒捩趔颲鮤鴷躐鬛鬣鱲啉粼鄰隣隣嶙遴辚瞵麐轔鏻麟鱗廪懔檩顲賃蔺膦閵蹸躏躙躪轥囹泠苓柃瓴鸰棂绫翎跉軨鈴閝輘霊錂霗魿鲮鴒鹷霛霝齢酃鯪齡醽靈麢龗阾領呤熘浏旒遛骝飗镏鹠镠鎏鎦麍鏐飀鐂飅鰡鶹绺锍鋶蹓霤雡飂鬸鷚泷茏栊珑胧砻龍鏧霳龒龓豅躘鑨靇鸗垅隴贚偻喽蒌遱謱軁髅鞻嵝镂鏤噜撸垆泸栌胪轳舻鈩鲈魲轤鑪顱鱸鸕黸鹵魯橹镥鏀鐪鑥辂陸渌逯賂輅漉趢踛辘醁錄録錴璐鴼蹗轆鏕鯥鵦鵱鏴鷺氇闾榈閭鷜郘稆膂鋁鑢栾脔銮鵉鑾鸞釠锊鋝鋢囵陯踚輪錀鯩論捋頱猡脶椤镙鏍邏鸁鑼倮躶蠃泺荦珞摞漯雒鮥鸬鹩簏篥笠蠡蠊蝼螂蜊蛉蛎聆癞癃瘰瘘瘌痨疬疠鹭鹨鸾耧耢耒褴褛裣裢鹂';
	v_Charm   Varchar2(2000) := '嬷犸遤鎷鷌鰢杩閁唛鬕霾荬買鷶劢麥賣邁霡霢顢鞔鳗鬗鬘鰻鏋鄤墁幔缦熳镘謾鏝邙硭釯铓鋩漭貓牦旄軞酕髦錨鶜峁泖茆昴鉚耄袤貿鄚瑁瞀鄮懋莓郿嵋湄猸楣镅鋂鎇鶥黴浼躾鎂黣跊鬽韎魅扪钔門閅鍆焖懑雺甍瞢鄳鄸朦礞鯍艨鹲靀顭鸏勐艋錳懵鯭鼆霥霿踎咪祢猕謎縻麊麋麛蘼镾醾醿鸍釄芈弭敉脒銤冖糸汨宓谧嘧鼏謐宀沔黾眄湎腼鮸靣麪麫麺麵喵鶓鱙杪眇淼缈邈乜咩鴓鑖鱴岷玟苠珉缗鈱賯錉鴖鍲闵泯閔愍黽閩鳘鰵茗冥鄍溟暝銘鳴瞑酩缪謬谟嫫馍麼麽魹謨謩譕麿殁茉秣貃蓦貊銆靺镆魩黙貘鏌哞侔眸鉾謀鍪鴾麰毪鉧踇仫沐坶苜钼雮鉬霂鞪鹋袂鹛蠓蟊蟆蟒螨蝥蜢蛑虻篾蠛颟耱瘼';
	v_Charn   Varchar2(2000) := '誽镎鎿雫肭捺豽軜貀鈉靹魶艿迺釢柰萘鼐錼囡喃遖楠諵難赧腩囔鬞馕曩攮齉孬呶硇铙猱譊鐃垴瑙閙鬧讷餒鮾鯘嗯鈪銰坭怩郳铌猊跜鈮貎輗鲵鯢麑齯伲旎鉨隬鑈迡昵睨鲇鮎鲶鵇鯰辇輦蹨躎廿埝醸釀茑袅鳥嬲脲肀陧臬隉嗫鉩踂踗踙錜蹑鎳闑蘖齧讘躡鑷顳钀咛鑏鬡鸋佞甯妞忸狃鈕靵侬哝農辳醲齈譨鎒鐞譳孥驽弩胬钕釹恧衄黁郍傩喏逽搦锘諾蹃鍩黏颞聍耨衲蝻蛲';
	v_Charo   Varchar2(2000) := '噢鞰讴瓯鴎謳鏂鷗齵怄耦';
	v_Charp   Varchar2(2000) := '葩杷俳輫哌蒎鎃爿跘蹒蹣鎜鞶泮頖鋬鵥鑻雱滂霶逄鳑龎龐鰟脬庖狍匏軳鞄麅麭醅阫陫锫賠錇帔旆辔霈轡湓怦軯閛嘭堋輣錋韸韼鵬鬔鑝踫闏丕纰邳铍豾釽鈚鈹鉟銔噼錃錍魾闢阰芘枇郫陴埤豼鲏罴隦魮鮍貔鵧鼙庀仳圮銢諀鴄擗淠媲睥甓鷿鸊犏翩鶣骈胼賆諚蹁谝貵諞魸剽缥飃飄魒闝殍瞟醥顠嘌嫖氕丿苤鐅姘貧嫔頻顰榀牝娉俜頩郱枰軿鲆輧鮃钋釙酦醱鏺鄱謈叵钷鉕珀頗颒掊裒攴攵陠噗鋪鯆匍酺璞濮镤贌鏷溥氆諩镨譜蹼鐠皤疋襻螃蟛筢笸蟠螵蜱蚍颦袢癖疱';
	v_Charq   Varchar2(2000) := '迉桤郪萋嘁槭諆踦諿霋蹊魌鏚鶈亓圻岐芪耆淇萁跂軝釮骐琦琪祺﨑锜頎鬾鬿綦齊蕲踑錡鲯鳍鯕鵸鶀麒鬐魕鰭麡邔屺芑杞豈绮綮諬闙汔荠葺碛憩葜跒酠鞐髂阡芊佥岍悭谸釺鈆雃愆鉛骞鹐搴諐遷褰謙顅鏲鵮鐱鬜鬝韆荨钤掮軡鈐鉆鉗銭錢鎆黚鰬凵肷慊缱譴鑓芡茜倩椠輤戕戗跄锖锵錆蹌镪蹡鎗鏘鏹嫱樯謒羟炝硗郻鄗跷鄡鄥劁踍頝缲鍫鍬趬蹺蹻鐰荞谯憔鞒樵譙趫鐈鞽顦釥愀诮陗誚韒鞩躈妾挈惬锲魥踥鍥鯜鐑衾誛顉鮼芩鈙雂嗪溱靲噙鳹檎赺赾锓鋟吣揿靑郬圊軽輕鲭鯖鑋檠黥苘頃請謦靘磬跫銎邛茕赹楸鹙趥鳅鞦鞧鰌鰍鶖鱃龝犰俅逎逑釚赇釻巯遒裘賕銶醔鮂鼽鯄鵭鰽糗岖诎阹祛誳麹魼趨麯軀麴黢鰸鱋劬朐軥蕖磲鴝璩鼩蘧氍衢躣鑺鸜齲迲郥阒觑閴麮闃鼁悛鐉诠荃辁铨跧輇銓踡闎鳈鬈鰁齤顴犭畎绻韏悫阕阙趞闋闕鵲逡鸲蝤蜷蜞蜻蜣蛴蛐蛩蚯箝箧箐筌筇罄蠼螓虬虔颀覃襁穹癯';
	v_Charr   Varchar2(2000) := '髯苒禳躟鬤譲讓荛桡娆隢遶亻鈓魜銋鵀荏稔躵仞讱轫饪恁軔葚靭靱韌飪認餁辸陾釰鈤肜狨嵘榕镕鎔軵糅蹂輮鍒鞣鰇鶔韖邚铷銣鴑嚅濡薷鴽醹顬鱬鄏込洳溽缛蓐鳰朊軟輭蕤芮枘睿銳鋭閏閠偌鄀鰙鰯鶸穰箬蝾蚺蚋颥衽襦';
	v_Chars   Varchar2(2000) := '仨靸卅钑飒脎鈒隡颯噻顋鰓賽毵鬖糁馓鏒閐搡磉鎟顙缫臊鳋颾鰠鱢埽啬铯雭銫轖鏼譅飋鬙閪铩裟魦鲨閷鎩鯊鯋唼歃閯霎彡邖芟姗钐埏舢軕釤閊跚潸膻鯅陝閃讪剡赸銏骟鄯嬗謆譱贍鐥鳝鱓鱔殇觞熵謪鬺垧賞鑜绱艄輎颵鮹苕劭潲猞畲輋賒賖佘厍滠韘麝诜鲹鯓鵢鯵鰺鉮鰰邥哂矧谂谉渖諗頣魫讅胂椹鋠阩陞陹﨡鉎鍟鼪鵿渑譝鱦眚晟貹嵊賸邿鸤釶蓍鉇酾鳲鳾鲺鍦鯴鰤鶳釃饣辻飠炻埘莳遈鉐鉽鲥鮖鼫識鼭鰣豕鉂礻贳轼铈釈弑谥貰軾鈰鉃飾適銴諟諡遾餝謚釋鰘齛扌艏狩绶鏉殳纾陎姝倏菽軗鄃摅毹跾踈輸鮛鵨秫塾贖鼡鱪鸀鱰沭腧鉥澍豎錰鏣鶐鶑唰誜闩閂涮﨎雙孀鷞鹴鸘鏯誰氵閖順鬊說説妁铄嗍搠蒴槊鎙鑠厶纟咝缌鉰飔厮銯锶澌鋖鍶颸鐁鷥鼶汜兕姒祀泗驷俟飤釲貄鈻飼忪凇崧淞菘嵩悚頌誦鎹鄋嗖溲馊飕锼醙鎪颼叟嗾瞍薮稣鯂夙涑谡嗉愫遡鹔蔌觫趚遬鋉餗謖蹜鱐鷫狻荽眭睢濉鞖雖遀隨谇誶賥燧邃鐆譢鐩狲荪飧飱隼榫鎨鶽娑挲桫睃嗦羧趖鮻唢鎍鎖鎻鎼鏁逤穑鸶疝痧筮笥笙舐蟮蟀螫螋蛸簌筲蜃蛳颡耜竦瘙';
	v_Chart   Varchar2(2000) := '趿铊溻鉈蹹鮙鳎鰨闼遝遢阘榻誻錔鞜闒鞳闥譶躢骀邰炱跆鲐颱鮐薹肽钛鈦貪昙郯锬談醈錟顃譚貚醰譠鷤忐钽鉭醓赕賧铴羰镗蹚鏜鐋鞺鼞饧鄌溏隚瑭樘踼赯醣鎕闛鶶帑傥镋鎲钂韬飸謟鞱韜饕迯洮啕鞀醄鞉鋾錭鼗忑忒貣铽慝鋱鼟滕邆謄鰧霯銻鷈鷉绨缇遆趧醍謕蹏鍗鳀鴺題鮷鵜鯷鶗鶙躰軆倜悌逖逷鐟趯酟靔黇靝畋阗鴫闐鷆鷏忝殄餂賟錪靦掭佻祧龆鋚鞗髫鲦鯈鎥齠鰷誂粜铫趒頫萜貼跕鉄銕鴩鐡鐢鐵飻餮町鞓邒莛婷葶閮霆諪鼮梃铤颋誔鋌頲嗵仝佟茼砼赨鉖僮鉵銅餇鲖潼鮦恸鍮亠骰頭黈鋵鵚鼵荼鈯跿酴鍎鵌鶟鷋鷵钍釷迌堍菟鵵貒抟鏄鷒鷻疃彖隤頹頺頽魋蹪蹆煺暾黗饨豘豚軘飩鲀魨霕氽乇讬飥魠佗陁坨沱迱柁砣跎酡踻橐鮀鴕鼧鼍鼉庹鵎鰖柝跅鹈窕箨笤螳螗蜩蜓耥裼';
	v_Charw   Varchar2(2000) := '娲鼃佤邷腽韈韤崴顡剜纨芄貦頑邜莞绾脘菀琬畹輓踠鋔鍐鋄錽贃鎫贎罔惘辋誷輞魍迋偎逶隇隈葳煨薇鳂鰃鰄囗圩帏沩闱韋涠帷嵬違鄬醀鍏闈鮠霺霻炜玮洧娓诿隗猥艉韪鲔諉踓韑頠鍡鮪韙颹韡軎猬謂錗鮇轊鏏霨鳚讆躗讏躛辒豱輼轀鳁鎾鰛鰮阌鈫雯魰鳼鴍閺閿闅鼤闦闧刎汶顐璺鹟鎓鶲蓊蕹齆倭莴喔踒肟幄渥硪龌齷圬邬趶釫鄔誈誣鴮鎢鰞郚唔浯鹀鵐鯃鼯鷡仵妩庑忤怃迕牾錻鵡躌兀兀阢杌芴逜焐婺隖靰骛寤誤鋈霚鼿霧齀鶩鹉蜿蜈痿痦鹜';
	v_Charx   Varchar2(2000) := '兮诶郗唏奚浠欷淅菥赥釸粞翕舾鄎僖誒豨餏嬉餙樨歙熹羲錫谿豀豯貕雟鯑鵗譆醯鏭隵曦酅鼷鸂鑴郋觋趘隰謵鎴霫鳛飁鰼玺徙葸鈢屣蓰銑禧諰謑蹝鱚躧饩郄郤釳阋舄趇禊赩隟黖鬩闟霼呷谺閕颬鰕狎柙陜硖陿遐瑕赮魻轄鍜鎋黠鶷閜諕鏬氙祆籼莶铦跹酰銛暹韯鍁鍂韱鮮蹮譣鶱躚鱻娴閑銜誸賢諴輱醎鹹贒鑦鷳鷴鷼冼猃険赻跣險藓鍌燹顕韅顯岘苋陥誢鋧錎豏麲鏾霰鼸芗郷鄉鄊缃葙鄕骧麘鱜鑲庠跭饷飨銄餉鲞鮝鯗響鱶項鐌鱌枭哓枵骁绡逍鸮潇踃銷魈鴞謞鴵鷍郩崤誵謏誟﨧偕勰撷缬諧鞵鐷讗龤绁亵渫榍榭韰廨獬薤邂燮謝鞢瀣齘齥齂躠躞忄邤昕莘鈊歆鋅馨鑫鬵鐔阠囟軐顖釁謃鮏鯹陉郉钘陘硎铏鈃鉶銒鋞擤荇悻﨨芎讻诇咻庥貅馐銝髹鎀鮴鵂鏅飍岫溴銹鏥鏽齅盱砉顼谞須頊魆諝譃魖鑐鬚诩栩鄦糈醑洫勖溆煦賉銊鱮蓿軒谖揎萱暄煊儇諠諼鍹譞鰚讂漩璇選泫炫铉渲楦鉉碹镟鞙颴鏇贙辥鞾泶鸴踅雤鷽轌鳕鱈谑趐謔埙獯薰曛醺峋恂洵浔荀鄩鲟鱏鱘徇迿巽遜賐蕈顨鑂皙箫筱筅罅蟓蟋螅蜥蚬胥穸痫痃鹇鸺';
	v_Chary   Varchar2(2000) := '压桠铔鴉錏鴨鵶鐚伢岈琊睚齖迓垭娅砑氩揠齾恹胭崦菸湮腌鄢嫣醃閹黫讠闫妍芫郔閆閻檐顏顔麙鹽麣兖俨偃厣郾酓琰遃隒罨魇躽黡鰋鶠黤齞龑黬黭顩鼴魘鼹齴黶晏隁焱滟鳫酽谳餍鴈諺赝鬳鴳酀贋軅醶鷃贗贘讌醼鷰釅讞豓豔泱鉠雵鞅鍈鴦阦炀钖飏徉烊陽諹輰鍚鴹颺鐊鰑霷鸉軮養怏恙幺夭吆鴁爻肴轺珧軺徭遙銚飖餆餚繇謠謡鎐鳐颻顤鰩杳崾鴢闄齩鷕靿鼼曜鷂讑鑰揶铘釾鋣鎁邺頁晔烨谒鄓鄴靥謁鍱鎑鵺靨鸈辷咿猗郼欹漪銥噫鹥醫黟譩鷖黳圯诒怡迤饴咦荑贻迻眙酏貽誃跠頉飴遺頤頥嶷顊鮧謻鏔讉鸃迆钇苡舣釔逘鈘鉯鳦旖輢顗轙齮弋刈仡阣佚呓佾峄怿驿奕弈羿轶悒挹貤陭埸豙豛釴隿跇軼鈠缢靾熠誼镒鹝鹢黓劓殪薏翳貖鮨贀鎰镱豷霬鯣鶂鶃鶍譯議醳醷鐿鷁鷊懿鷧鷾讛齸阥洇氤陰铟陻隂喑堙銦鞇諲霒闉霠韾垠狺鈝龂鄞夤誾銀龈霪齗齦鷣廴吲釿鈏飲隠靷飮趛隱讔茚胤酳鮣莺瑛锳嘤撄賏璎霙鴬膺韺鎣鶧譻鶯鑍鷪軈鷹鸎鸚茔荥萦楹滢蓥潆嬴謍瀛贏郢颕頴鐛媵鞕譍唷邕鄘墉慵銿壅郺镛雝鏞鳙饔鱅鷛喁颙顒鰫俑鲬踴鯒醟攸呦麀鄾尢柚莜莸逌郵逰遊鱿猷鈾鲉輏魷輶鮋邎卣莠铕銪牖黝侑囿宥迶貁酭誘鼬纡迃陓邘妤欤於臾禺舁狳谀酑馀萸釪隃雩魚嵛揄腴瑜觎諛雓餘魣踰輿鍝謣鮽鯲鰅鷠鸆伛俣圄圉庾鄅铻語鋙龉貐麌齬聿妪饫昱钰﨏谕逳阈飫煜蓣鈺預毓輍銉隩遹鋊鳿燠諭錥閾鴥鴧鴪魊醧鵒譽轝鐭霱鬻鱊鷸鸒軉鬰鬱眢鹓鳶鋺鴛鵷鼘鼝贠邧沅爰貟酛鈨鼋塬魭圜橼謜轅黿鎱邍鶢鶰逺遠垸媛掾瑗願刖軏钺跀鈅鉞閱閲樾龠瀹黦躍鸑龥鸙赟頵贇纭芸昀鄖雲氲鋆阭狁殒鈗隕霣齫齳郓恽鄆愠運韫熨賱醖醞韗韞韻甬鹦痖瘀螈蝣蝓蜴蜮蛘蚴蚰蚓颍窳箢筵竽罂窨窬窈翊癔瘾瘿瘗瘐痍疣鹬鹞鹆鸢';
	v_Charz   Varchar2(2000) := '卮仄赜仉伫侏倬偬俎冢诏诤诹诼谘谪谮谵阼陟陬郅邾鄣鄹圳埴芷苎茱荮菹蓁蕞奘拶揸搌摭摺撙擢攥吒咂咤哳唣唑啧啭啁帙帻幛峥崽嵫嶂徵獐馔忮怍惴浈洙浞渚涿潴濯迮彘咫姊妯嫜孳驵驺骓骘纣绉缁缒缜缯缵甾璋瓒杼栉柘枳栀桎桢梓棹楂榛槠橥樽轵轸轾辄辎臧甑昃昝贽赀赈肫胄胙胗胝朕腙膣旃炷祉祚祗祯禚恣斫砟砦碡磔黹眦畛罾钊钲铢铮锃锱镞镯锺雉秭稹鸩鸷鹧痄疰痣瘃瘵窀褶耔颛蚱蛭蜇螽蟑竺笊笫笮筝箦箸箴簪籀舯舳舴粢粽糌翥絷趱赭酎酯跖踬踯踵躅躜豸觜觯訾龇錾鲰鲻鳟髭麈齄';

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
			If v_Bitchar >= '啊' And v_Bitchar <= '座' Then
				For v_Chrnum In 1 .. Length(v_Stdstr) Loop
					If Substr(v_Stdstr, v_Chrnum, 1) = '-' Then
						Null;
					Elsif v_Bitchar < Substr(v_Stdstr, v_Chrnum, 1) Then
						v_Spell := v_Spell || Chr(64 + v_Chrnum);
						Exit;
					End If;
				End Loop;
				If v_Bitchar >= '匝' Then
					v_Spell := v_Spell || 'Z';
				End If;
			Elsif Instr('ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.+-*/', v_Bitchar) > 0 Then
				v_Spell := v_Spell || v_Bitchar;
			Elsif Instr('ⅠⅡⅢⅣⅤⅥⅧⅧⅨ', v_Bitchar) > 0 Then
				v_Spell := v_Spell || Chr(Ascii(v_Bitchar) - 41664);
			Elsif Instr('ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ',v_Bitchar) > 0 Then
				v_Spell := v_Spell || Chr(Ascii(v_Bitchar) - 41856);
			Elsif Instr('Αα', v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'A';
			Elsif Instr('Ββ', v_Bitchar) > 0 Then
				v_Spell := v_Spell || 'B';
			Elsif Instr('Γγ', v_Bitchar) > 0 Then
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
	v_a       Varchar2(1200) := '蔼艾鞍芭茇菝靶蒡苞葆蓓鞴苯荸荜萆蓖蔽薜鞭匾苄菠薄菜蔡苍藏艹草茬茶蒇菖苌臣茌茺莼茈茨苁葱蔟萃靼鞑甙萏荡菪荻蒂东鸫董蔸芏莪苊萼蒽贰藩蕃蘩范匚芳菲匪芬葑芙芾苻茯莩菔甘苷藁戈革葛茛工功攻恭廾巩汞共贡鞲苟菇菰鹳匦邯菡蒿薅荷菏蘅薨荭蕻葫花划萑荒黄茴荟蕙荤劐或获惑藿芨基蒺蕺芰蓟葭荚菅蒹鞯茧荐茳蒋匠艽茭蕉节戒芥藉堇荩靳觐茎荆菁警敬苴鞠鞫菊莒巨苣蕨菌蒈勘戡莰苛恐芤蔻苦蒯匡葵匮蒉莱蓝莨蒗劳勒蕾莉蓠藜苈荔莅莲蔹蓼蔺苓菱茏蒌芦萝荦落荬颟鞔蔓芒茫莽茅茆茂莓萌甍瞢蒙蘼苗鹋藐蔑苠茗摹蘑茉莫蓦某苜募墓幕慕暮艿萘匿廿茑孽蘖欧殴瓯鸥藕葩蒎蓬芘匹苤苹萍叵莆菩葡蒲七萋期欺芪其萁綦蕲芑荠葺葜芊荨芡茜蔷荞鞒巧翘鞘切茄芩芹勤擎檠苘跫銎邛茕蛩区蕖蘧荃颧鹊苒荛惹荏葚戎茸荣蓉鞣茹薷蓐蕤蕊芮若萨散莎芟苫芍苕甚蓍莳世式贳菽蔬薯蒴斯菘薮苏蔌蒜荽荪蓑苔薹萄忒慝藤萜莛葶茼荼菟芄莞菀葳薇苇萎蔚蓊蕹莴卧巫芜芴昔菥熙觋葸蓰匣莶藓苋芗葙巷项萧邪鞋薤芯莘薪荇芎蓄蓿萱靴薛薰荀蕈鸦牙芽雅迓菸蔫芫郾燕鞅尧药医荑颐苡弋艺薏翳茵荫鄞茚英莺茔荥荧莹萤营萦蓥莜莸莠萸芋蓣鸢苑芸蕴匝葬藻赜蘸蔗斟蓁蒸芝芷荮茱苎著茁菹蕞';
	v_b       Varchar2(1200) := '阿隘阪孢陂陛屮陈丞承蚩耻出除陲聪耽聃阽耵陡队堕耳防阝附陔隔耿孤聒孩函隍隳亟际降阶卩孑卺阱聚孓孔聩联辽聊了陵聆隆陇陋陆勐孟陌陧聂颞聍陪陴聘阡凵取娶孺阮陕隋随祟隧孙陶粜陀隈隗卫阢隰隙险限陷陉逊阳耶也阴隐隅院陨障阵职陟骘坠孜子陬鄹阻阼';
	v_c       Varchar2(1200) := '巴畚弁骠驳参骖叉骋驰骢皴迨怠邓叠对怼驸观骇骅欢鸡骥艰骄矜劲刭颈迳驹骏骒垒骊骝驴骡骆马矛蝥蟊瞀牟鍪难能骈骗骐骑巯驱劝逡柔叁毵桑颡骚骟圣驶双厶驷骀台邰炱通驮驼婺骛鹜戏骧骁熊驯验以矣驿甬勇恿又予驭预豫鹬允驵蚤骣骤驻骓驺';
	v_d       Varchar2(1200) := '砹碍鹌百邦帮磅悲碑辈碚奔泵砭碥髟飙鬓礴布礤厕碴虿厂耖砗辰碜成舂厨础春唇蠢磁蹙存磋厝耷达大砀焘磴砥碲碘碉碟碇硐碓礅趸砘夺厄而鸸砝矾非蜚斐翡奋丰砜酆奉砩尬尴感矸硌耕龚鸪辜古嘏故顾硅磙夯耗厚胡鹕瓠鬟磺灰彗慧耠矶剞髻恝戛硷碱礓耩礁碣兢鬏韭厩厥劂砍磕克刳夸夼矿盔奎髡砬耢耒磊厘历厉励砺砾奁鹩尥鬣磷硫龙砻聋垄耧碌码劢迈硭髦礞面耱奈耐硇碾耨恧耦耙耪匏裴砰硼碰砒破戚奇契砌碛牵硗挈秦磲鬈犬确髯辱三磉砂奢厍砷蜃盛石寿戍耍爽硕厮耜肆碎太态泰碳耥套髫厅砼砣碗万威硪戊矽硒袭硖夏厦咸厢硝硎雄髹戌砉碹压砑研奄厣魇厌砚艳雁餍赝页靥欹硬尢尤友有右郁原愿耘砸在臧仄砟丈磔砧碡砖斫髭耔鬃奏左';
	v_e       Varchar2(1200) := '爱胺肮膀胞豹膘豳膑脖膊采彩豺肠塍腠脆脞胆貂腚胨胴肚腭肪肥腓肺肤孚服郛脯腑腹尕戤肝肛胳膈肱股臌胍胱虢胲貉肌及胛腱胶脚腈肼胫雎爵胩胯脍腊肋臁脸膦胧胪脶脉貌朦脒觅腼邈膜貊貘肭乃鼐腩脑腻脲脓胖脬胚朋鹏膨脾貔胼脐肷腔且朐肜乳朊脎腮臊彡膻膳胂胜豕受腧甩舜胎肽膛腾滕腆腿豚脱妥腽脘腕肟奚膝燹县腺胁腥胸貅须悬胭腌腰遥繇舀鹞腋胰臆盈媵臃用有腴爰月刖孕脏膪胀胗朕肢胝脂豸膣肿肘逐助肫腙胙';
	v_f       Varchar2(1200) := '埃霭埯坳坝霸坂雹贲甏孛勃博鹁埠才裁场超朝坼趁城埕墀赤翅亍矗寸埭戴堤觌坻地颠坫垤堞耋动垌都堵堆墩垛二坊霏坟封夫赴垓干坩赶圪塥埂垢彀遘觏毂鼓瞽卦圭规埚过顸邗韩翰壕郝盍赫堠壶觳坏卉恚魂霍击圾赍吉戟霁嘉教劫颉截进井境赳救趄均垲刊堪坎考坷壳坑堀垮块款圹亏逵坤垃老雷塄嫠坜雳墚埒趔霖零酃垅露垆埋霾卖墁耄霉坶南赧垴坭霓辇埝培霈堋彭坯霹埤鼙圮坪坡埔亓圻耆起乾墙謦磬罄求逑裘趋麴去趣却悫壤韧颥丧埽啬霎埏墒垧赦声十埘士示螫霜寺索塌塔坍坛坦塘趟韬替填霆土堍坨顽韦圩违未雯斡圬无坞雾熹喜霞献霰霄孝协馨幸需墟雪埙垭盐堰壹圯埸懿堙垠霪墉雩雨域元垣袁鼋塬远垸越云运韫哉栽载趱增赵者赭真圳震支直埴址志煮翥专趑走';
	v_g       Varchar2(1200) := '瑷敖獒遨熬聱螯鳌骜鏊班斑甭逼碧表殡丙邴玻逋不残蚕璨曹琛豉敕刺璁琮殂璀歹带殆玳殚到纛玷靛玎豆逗毒蠹顿垩恶噩珥珐玢否麸敷甫副丐鬲亘更珙瑰翮珩瑚琥互画还环璜虺珲惠丌玑墼棘殛夹珈郏颊戋歼柬戬豇瑾晋靓静玖琚珏开珂琨剌来赉赖琅鹂璃逦理吏丽郦琏殓两列烈裂琳玲琉珑璐珞玛麦瑁玫芈灭玟珉末殁囊孬瑙弄琶丕邳琵殍平珀璞妻琦琪琴青琼球璩融瑞卅瑟珊殇事殊束死素速琐瑭忑天忝殄餮吞屯橐瓦歪豌玩琬王玮軎吾五武鹉兀瑕下现刑邢形型顼璇殉琊亚焉鄢严琰殃珧瑶一夷殪瑛璎迂于欤盂瑜与玉瑗殒再瓒遭枣责盏璋珍臻整正政殖至郅致珠赘琢';
	v_h       Varchar2(1200) := '龅彪卜步睬餐粲柴觇龀瞠眵齿瞅龊雌此鹾眈瞪睇点盯鼎督睹盹丨壑虍虎乩睑睫睛旧龃具遽瞿矍卡瞰瞌肯眍眶睽睐瞵龄卢鸬颅卤虏虑瞒眯眠眄瞄眇瞑眸目睦睨虐盼皮睥瞟频颦颇攴歧虔瞧氍龋觑睿上叔睡瞬瞍眭睢睃忐龆眺瞳凸龌瞎些盱虚眩睚眼眙龈卣虞龉眨砦瞻占战贞睁止瞩卓桌赀觜龇紫訾眦';
	v_i       Varchar2(1200) := '澳灞浜滗濞汴滨濒波泊渤不沧漕测涔汊潺尝常敞氅潮澈尘沉澄池滁淳淙淬沓淡澹当党凼滴涤滇淀洞渎渡沌沲洱法泛淝沸汾瀵沣浮涪滏尜溉泔澉淦港沟沽汩涫灌光滚海涵汉汗瀚沆濠浩灏河涸泓洪鸿黉鲎滹湖浒沪滑淮洹浣涣漶湟潢辉洄汇浍浑混溷活激汲脊洎济浃尖湔涧渐溅江洚浇湫洁津浸泾酒沮举涓觉浚渴溘喾溃涞濑澜漤滥浪潦涝泐泪漓澧沥溧涟濂潋梁粱劣洌淋泠溜浏流鎏泷漏泸渌滤漉潞滦沦泺洛漯满漫漭泖没湄浼懑汨泌沔湎淼渺泯溟沫漠沐淖泥溺涅泞浓沤派湃潘泮滂泡沛湓澎淠漂泼婆濮浦溥瀑沏柒漆淇汔汽泣洽潜浅溱沁清泅渠雀染溶濡汝洳溽润洒涩沙裟鲨潸汕裳赏尚少潲涉滠深沈渖渗渑省湿淑沭漱澍涮氵水澌汜泗淞溲涑溯濉娑挲溻汰滩潭汤堂棠溏淌烫涛滔洮逃淘鼗涕添汀潼涂湍沱洼湾汪沩涠潍洧渭温汶涡沃渥污浯鋈汐浠淅溪洗涎湘削消逍潇淆小肖泄泻渫瀣兴汹溴洫溆漩泫渲学泶洵浔汛涯淹湮沿演滟泱洋漾耀液漪沂溢洇淫滢潆瀛泳涌油游淤渝渔浴誉渊沅源瀹澡泽渣沾澶湛漳涨掌沼兆浙浈汁治滞洲洙潴渚注涿浊浞濯淄滋滓渍';
	v_j       Varchar2(1200) := '暧暗昂蚌暴蝙晡螬蝉蟾昌畅晁晨蛏螭匙虫蜍蝽旦刂戥电蝶蚪蛾遏蜂蚨蜉蝠蝮旰杲蛤虼蚣蛄蛊归晷炅蝈果蜾蚶晗旱蚝昊颢曷蚵虹蝴蝗蟥晃晖蛔晦蟪夥蠖虮蛱坚监鉴蛟蚧紧晶景颗蝌旷暌蝰昆蛞旯蜡览螂蜊里蛎蠊晾量临蛉蝼螺蟆蚂螨曼蟒蛑昴冒昧虻盟蜢蠓冕蠛明暝螟蝻曩蛲昵暖蟠螃蟛蚍蜱螵曝蛴蜞蜣螓蜻晴蚯虬蝤蛆蛐蠼蜷蚺日蝾蠕蚋晒蟮晌蛸蛇申肾晟师时是暑曙竖墅帅蟀蛳螋遢昙螗螳剔题蜩蜓蜕暾蛙蜿晚旺韪蚊蜗蜈晤晰蜥螅蟋曦虾暇暹贤显蚬蟓晓歇蝎昕星勖煦暄曛蚜蜒晏蛘曜野曳晔蚁易蜴蚓蝇影映蛹蚰蝣蚴禺愚蝓昱遇蜮螈曰昀晕早昃蚱蟑昭照蜘蛭蛛蛀最昨';
	v_k       Varchar2(1200) := '吖啊嗄哎唉嗳嗌嗷叭吧跋呗趵嘣蹦吡鄙哔跸别啵踣跛卟哺嚓踩嘈噌蹭躔唱嘲吵嗔呈逞吃哧嗤踟叱踌躇蹰啜嘬踹川喘串吹踔呲蹴蹿啐蹉哒嗒呆呔啖叨蹈噔蹬嘀嗲踮叼吊跌喋蹀叮啶咚嘟吨蹲咄哆踱哚跺呃鄂鹗颚蹯啡吠吩唪呋趺跗呒咐嘎噶嗝跟哏哽咕呱剐咣贵跪呙哈嗨喊嚆嗥嚎号呵喝嗬嘿哼哄喉吼呼唿唬哗踝唤患咴哕喙嚯叽咭唧跻戢哜跽跏趼践踺跤叫噍喈嗟噤啾咀踽距踞鹃噘噱蹶嚼咔咖喀咳嗑啃吭口叩哭跨哙哐喹跬喟啦喇啷唠叻嘞喱哩呖唳跞踉嘹咧躐啉躏另呤咯咙喽噜路鹭吕骂唛吗嘛咪嘧黾喵咩鸣哞哪喃囔呶呐呢嗯啮嗫蹑咛哝喏噢哦呕趴啪哌蹒咆跑呸喷嘭噼啤蹁嘌品噗蹼嘁蹊器遣呛跄跷嗪噙吣嚷蹂嚅噻嗓唼啥跚哨呻哂史嗜噬唰吮顺嗍嗽咝嘶嗣嗖嗾嗉虽唆嗦唢趿踏蹋跆叹饕啕踢啼蹄嚏跳听嗵吐跎鼍唾哇唯味喂吻嗡喔呜吴唔吸唏嘻呷吓跹跣响哓嚣哮啸躞兄咻嘘嗅喧勋呀哑咽唁吆咬噎叶咿噫咦遗呓邑喑吟吲嘤郢哟唷喁咏踊呦吁喻员跃郧咂咱唣噪躁啧吒咋哳喳咤啁吱跖踯只趾踬中忠盅踵咒躅嘱啭啄踪足躜嘴唑';
	v_l       Varchar2(1200) := '黯罢办畀边黪车畴黜辍辏黩囤轭恩罚畈罘辐辅罡哿轱罟固轨辊国贺黑轰轷囫回畸羁辑加迦袈甲驾架囝轿较界轲困罱累罹力轹詈连辆辚囹轳辂辘略囵轮罗逻皿墨默囡男嬲畔毗罴圃畦黔堑椠轻圊黥囚黢圈辁畎轫软轼输署蜀思四田畋町图团疃畹辋囗围畏胃辖黠勰轩鸭罨轺黟轶因黝囿圄圉园圆辕圜暂錾罾轧斩辗罩辄辙轸畛轵轾置轴转辎罪';
	v_m       Varchar2(1200) := '岸盎凹岜败贝崩髀贬飑飚髌财册岑崇帱遄赐崔嵯丹嶝迪骶巅典雕岽峒髑赌朵剁峨帆幡凡贩风峰凤幅幞赋赙赅冈刚岗骼岣购鹘骨崮刿崞帼骸骺岵凰幌贿岌几嵴觊岬见贱峤骱巾赆冂迥飓崛峻凯剀髁岢崆骷髋贶岿崃岚崂嶙岭髅嵝赂幔峁帽嵋岷内帕赔帔岐崎屺岂髂岍嵌峭赇曲岖冉嵘肉山删赡赊嵊殳赎兕崧嵩飕夙髓岁炭赕贴帖同彤骰崴网罔巍帏帷嵬幄峡岘崤岫峋岈崖崦岩央鸯崾贻嶷屹峄婴罂鹦由邮嵛屿峪崽赃则帻贼赠崭帐账嶂幛赈峥帧帙帜峙周胄贮颛赚幢嵫';
	v_n       Varchar2(1200) := '懊悖鐾必愎辟壁嬖避臂璧襞忭擘檗怖惭惨恻层孱忏羼惝怅怊忱迟尺忡憧惆丑怵憷怆戳悴翠忖怛惮蛋忉导悼惦殿刁懂恫惰屙愕发飞悱愤怫改敢怪惯憨悍憾恨恒惚怙怀慌惶恍恢悔屐己忌悸届尽惊憬居局剧惧屦恺慨忾慷尻恪快悝愦愧悃懒愣怜懔鹨戮屡履买慢忙眉鹛懵乜民悯愍恼尼怩尿忸懦怄怕爿怦劈屁甓譬屏恰悭慊戕悄憔愀怯惬情屈悛慑慎尸虱屎恃收书疏刷司巳忪悚愫屉悌惕恬恸屠臀惋惘惟尾尉慰屋忤怃悟惜犀习屣遐屑懈忄心忻惺性悻胥恤恂迅巽疋恹怏怡乙已以忆异怿羿悒翌翼慵忧愉羽悦恽愠熨奘憎翟展怔咫忮昼属惴怍';
	v_o       Varchar2(1200) := '粑爆焙煸灬炳灿糙焯炒炽炊糍粗粹灯断煅炖烦燔粉粪烽黻黼糕焓焊烘糇烀煳糊焕煌烩火糨烬粳精炯炬爝糠炕烤烂烙类粒粝炼粮燎料粼遴熘娄炉熳煤焖迷米敉糯炮粕炝糗炔燃熔糅糁煽剡熵烧炻数烁燧郯糖烃煺烷煨为炜焐烯粞熄籼燮糈煊炫烟炎焰焱炀烊业邺烨熠煜燠糌糟凿灶燥炸粘黹烛炷灼籽粽';
	v_p       Varchar2(1200) := '安案袄宝褓被褙裨窆褊裱宾补察衩禅宸衬裎褫宠初褚穿窗辶祠窜褡裆宕祷定窦裰额祓袱福富袼割宫寡褂官冠宄害寒罕褐鹤宏祜寰宦逭豁祸寂寄家袷裥謇蹇窖衿襟窘究裾窭军皲窠客裉空寇窟裤宽窥褴牢礼帘裢裣寥寮窿禄褛裸袂寐祢冖宓密幂蜜宀冥寞衲宁甯农袢襻袍祁祈祺骞搴褰襁窍窃寝穷穹祛裙禳衽容冗襦褥塞赛衫社神审实礻视室守祀宋宿邃它袒裼祧窕突褪袜剜完宛窝寤穸禧禊祆宪祥宵写袖宣穴窨宴窑窈衤宜寅廴宥窬宇窳寓裕冤郓灾宰宅窄寨褶这祯鸩之祗祉窒冢宙祝窀禚字宗祖祚';
	v_q       Varchar2(1200) := '锕锿铵犴钯鲅钣镑勹包饱鲍狈钡锛狴铋鳊镖镳鳔镔饼钵饽钹铂钸钚猜馇锸猹镲钗馋镡铲猖鲳鬯钞铛铖鸱饬铳刍锄雏触舛钏璠锤匆猝镩锉错岛锝镫镝狄氐邸甸钿鲷钓铞鲽钉锭铥兜钭独镀锻镦钝多铎锇饿锷鳄儿鲕尔迩饵铒钒犯饭钫鲂鲱狒镄鲼锋孵凫匐负鲋鳆钆钙钢镐锆镉铬鲠觥勾钩狗够觚钴锢鲴鳏馆盥犷逛龟鲑鳜鲧锅猓铪狠訇猴忽狐斛猢鹱铧猾獾郇锾奂鲩鳇昏馄锪钬镬饥急鲚鲫镓铗钾鲣锏饯键鲛角狡饺铰桀鲒解钅金锦馑鲸獍镜久灸狙锔句钜锯镌锩狷觖獗镢钧锎铠锴钪铐钶锞铿狯狂馈锟鲲铼镧狼锒铹铑乐鳓镭狸鲡锂鲤鳢猁鲢镰链獠镣钌猎鳞铃鲮留遛馏镏锍镂鲈鲁镥铝卵锊猡锣镙犸馒鳗镘猫锚卯铆贸猸镅镁钔猛锰猕免勉名铭馍镆钼镎钠馕铙猱馁铌猊鲵鲇鲶鸟袅镊镍狞狃钮钕锘刨狍锫铍鲆钋钷铺匍镤镨鳍钎铅钤钱钳欠锖锵镪锹锲钦锓卿鲭鳅犰劬鸲铨犭然饶饪狨铷锐鳃馓鳋色铯杀刹铩煞钐鳝觞勺猞狮鲺饣蚀鲥氏饰铈弑狩铄锶饲馊锼稣觫狻狲飧锁铊獭鳎鲐钛锬钽铴镗饧铽锑逖鲦铫铁铤铜钍兔饨鸵外危猥鲔猬刎我乌邬钨勿夕希郗欷锡玺铣饩郄狎狭锨鲜猃馅镶饷象枭销獬邂蟹锌鑫猩凶匈馐锈铉镟鳕獯旬鲟爻肴鳐铘猗铱饴钇刈逸镒镱铟狺银夤饮印迎镛鳙犹铀鱿铕鱼狳馀饫狱钰眢鸳猿怨钥钺匀狁锃铡詹獐钊锗针镇争狰钲铮炙觯钟锺皱猪铢橥铸馔锥镯锱鲻邹鲰镞钻鳟';
	v_r       Varchar2(1200) := '挨捱皑氨揞按翱拗扒捌拔魃把掰白捭摆拜扳搬扮拌报抱卑鹎拚摈兵摒拨播帛搏捕擦操插搽拆掺搀抄扯掣撤抻撑魑持斥抽搐搋揣氚捶撺摧搓撮挫措搭打担掸氮挡氘捣的抵掂垫掉迭瓞揲氡抖盾遁掇扼摁反返氛缶扶拂氟抚拊擀缸皋搞搁搿拱瓜挂拐掼罐皈鬼掴氦捍撖撼皓后逅护换擐皇遑挥攉挤掎技搛拣捡挢皎搅敫接揭拮捷斤近揪拘掬拒据捐撅抉掘攫捃揩看扛抗拷氪控抠扣挎揆魁捆扩括拉拦揽捞擂魉撩撂捩拎拢搂撸掳氯掠抡捋摞魅扪描抿摸抹拇捺氖攮挠拟拈年捻撵捏拧牛扭挪搦爬拍排乓抛抨捧批披郫擗氕撇拼乒皤迫魄掊扑颀气掐扦掮抢撬擒揿氢氰丘邱泉缺攘扰热扔揉撒搡搔扫擅捎摄失拾势拭逝誓手扌授抒摅摔拴搠撕搜擞损所挞抬摊探搪掏提掭挑挺捅投抟推托拖拓挖挽皖魍挝握捂舞罅氙掀魈挟携撷卸欣擤揎踅押氩揠掩扬氧邀摇揶掖揖抑挹殷氤撄拥揄援掾岳氲拶攒皂择揸扎摘搌招找蜇折哲蛰摺振挣拯卮执絷摭指制质挚贽掷鸷朱邾拄抓爪拽撰撞拙捉擢揍攥撙';
	v_s       Varchar2(1200) := '桉柏板梆榜棒杯本杓标彬槟柄醭材槽杈查槎檫郴榇柽枨酲橙酬樗橱杵楮楚椽棰槌椿醇枞楱酢醋榱村档柢棣丁酊顶栋椟杜椴樊梵枋榧酚棼焚枫桴覆概杆柑酐橄杠槔槁哥歌格根梗枸构酤梏棺桄柜桂棍椁醢酣杭核桁横槲醐桦槐桓桧机极楫枷贾枧检楗槛椒酵醮杰槿禁柩桕椐桔橘榉醵鄄桷橛楷栲柯棵可枯酷框醌栝栏婪榄榔醪栳酪檑酹棱楞李醴枥栎栗楝椋林檩柃棂榴柳栊楼栌橹麓榈椤杩懋枚梅楣酶檬梦醚棉杪酩模木柰楠酿柠杷攀醅配棚枇剽飘瓢票榀枰朴栖桤槭棋杞枪樯橇桥樵檎楸权醛榷桡榕枘森杉梢椹酾柿枢梳术述树栓松酥粟酸榫桫梭榻酞覃檀樘醣桃梯醍梃桐酮桶酴柁酡椭柝枉桅梧杌西析皙樨醯檄柙酰相想橡枵校楔械榍榭醒杏朽栩醑酗楦醺桠檐酽杨样杳要椰酏椅樱楹柚酉榆橼樾酝楂札栅榨栈樟杖棹柘桢甄榛枕枝栀植枳酯栉桎酎株槠杼柱桩椎酌梓棕醉樽柞';
	v_t       Varchar2(1200) := '矮岙奥笆稗般版舨备惫笨鼻彼秕笔舭币筚箅篦笾秉舶箔簸簿舱艚策长徜彻称乘惩程秤笞篪彳艟愁稠筹臭处舡船囱垂辞徂簇汆篡毳矬笪答待箪稻得德簦等籴敌笛第簟牒丢冬篼牍犊笃短簖躲舵鹅乏筏番翻繁彷舫篚逢稃符复馥竿秆筻睾篙稿告郜舸各躬篝笱箍牯鹄牿刮鸹乖管簋鼾航禾和很衡篌後乎笏徊徨篁簧徽秽积笄嵇犄箕稽笈籍季稷笳稼笺犍笕简舰毽箭矫徼秸街筋径咎矩榘犋筠犒靠科稞箜筘筷筐篑徕籁篮稂梨犁黎篱黧利笠篥笼篓舻簏氇稆律乱箩雒毛牦么每艨艋秘秒篾敏鳘秣毪牡牧穆黏臬衄筢徘牌盘磐逄篷片犏篇丿牝鄱笸攵氆乞迄憩千迁愆签箝乔箧箐筇秋鼽躯衢筌穰壬稔入箬穑歃筛舢稍筲艄舌射身矧升生牲笙甥眚剩矢适舐释筮艏秫黍税私笥艘簌算穗笋毯躺特甜舔条笤廷艇筒透秃徒颓乇箨往逶微委艉魏稳我午迕牾务物息牺悉稀舾徙系先舷衔筅香箱向箫筱笑囟衅行秀徐选血熏循徇衙延筵衍秧徉夭徭迤移舣役劓胤牖釉竽禹御毓箢粤簪昝赞造迮笮舴箦怎齄乍毡笊箴稹征筝徵知夂秩智稚雉舯螽种重舟籀竹竺舳筑箸篆秭笫自租纂';
	v_u       Varchar2(1200) := '癌疤瘢癍半瓣北邶背迸闭敝痹弊辨辩辫瘭憋鳖蹩瘪冫冰并病部瓿差瘥产阐冁阊闯痴啻瘛冲瘳疮疵瓷慈鹚次凑瘁痤瘩单郸瘅疸盗道羝弟帝递癫奠癜凋疔冻斗痘端兑阏阀痱疯冯盖疳赣戆羔疙阁羹痼关闺衮馘阂阖痕闳瘊冱痪豢癀阍疾瘠冀痂瘕间兼煎鹣减剪翦姜将浆奖桨酱交郊疖竭羯疥净痉竞竟靖阄疚疽蠲卷桊眷决竣阚闶疴况夔阃阔瘌辣癞兰阑阆痨冷立疠疬痢凉疗冽凛凌羚瘤六癃瘘闾瘰美门闷闵闽瘼闹疒逆凝疟判叛旁疱疲痞癖瞥瓶剖普凄前歉羌羟妾亲酋遒癯阒拳痊券瘸阕阙闰飒瘙痧闪疝善鄯商韶首兽瘦闩朔槊凇竦送塑遂羧闼瘫痰羰疼誊鹈剃阗童痛头闱痿瘟闻阌问痦羲阋闲痫鹇冼羡翔鲞效辛新歆羞痃癣丫痖阉闫阎颜兖彦羊疡养痒恙冶痍疫益翊意瘗毅癔音瘾瘿痈疣猷瘀瘐阈阅韵曾甑闸痄瘵站章鄣彰瘴疹郑症痔痣瘃疰妆装丬壮状准着兹咨姿资孳粢恣总尊遵';
	v_v       Varchar2(1200) := '嫒媪妣婢婊剥姹婵娼嫦巢媸巛妲逮刀嫡娣妒娥婀发妨妃鼢妇旮艮媾姑妫好毁婚姬即嫉彐妓既暨嫁奸建姣娇剿婕她姐妗婧鸠九臼舅娟君郡垦恳馗邋姥嫘娌隶灵录逯妈嬷媒妹媚娩妙嫫姆那娜奶嫩妮娘肀妞奴孥驽努弩胬怒女媲嫖姘嫔娉嫱群娆忍刃妊如嫂姗嬗劭邵娠婶始姝鼠恕孀妁姒叟肃帑迢婷退娃娲丸婉娓鼯妩嬉鼷媳舄娴嫌姓旭婿絮寻巡娅嫣妍鼹妖姚姨姻尹邕鼬妤臾舁娱聿妪媛杂甾嫜召妯帚姊';
	v_w       Varchar2(1200) := '俺傲八爸佰颁伴傍煲保堡倍坌俾便傧伯仓伧侧岔侪伥偿倡侈傺仇俦雠储传创从丛促爨催傣代岱贷袋黛儋但倒登凳低佃爹仃侗段俄佴伐垡仿分份忿偾俸佛伏俘斧俯釜父付阜傅伽鸽个公供佝估谷倌癸刽含颔合何盒颌侯候华化会伙货佶集伎偈祭佳价假俭件剑牮健僭僵焦僬鹪佼侥介借今仅儆僦俱倨倦隽倔俊佧龛侃伉倥侉侩郐傀佬仂儡俚例俐俪傈俩敛僚邻赁伶瓴翎领令偻侣仑伦倮们命侔仫拿倪伲你念您佞侬傩偶俳佩盆仳僻偏贫俜凭仆企仟佥倩戗劁侨俏侵衾禽倾俅全人亻仁仞任恁仍儒偌仨伞僧傻伤畲佘舍伸什食使仕侍售倏舒毹伺似俟怂耸颂俗僳隼他贪倘傥体倜佻停仝佟僮偷途氽佗佤偎伟伪位璺翁瓮倭仵伍侮兮翕僖歙侠仙像偕斜信休修鸺叙儇伢俨偃佯仰爷伊依仪倚亿仡佚佾佣俑优攸悠佑侑余俞逾觎舆伛俣欲鹆愈龠债仉仗侦侄值仲众侏伫住隹追倬仔偬俎佐作坐做';
	v_x       Varchar2(1200) := '绊绑鸨绷匕比毕毖毙弼编缏缤缠弛绸绌纯绰绐弹缔缎缍纺绯费纷缝弗绂绋艴缚绀纲缟纥给绠弓缑贯绲绗弘红弧缳缓幻绘缋绩缉畿级纪继缄缣缰疆绛犟绞缴皆结缙经弪纠绢绝缂绔纩缆缧缡蠡练缭绫绺缕绿纶络缦弥弭糸绵缅缈缗缪母纳纽辔纰缥绮缱强缲顷绻绕纫绒缛弱缫纱缮绱绍绅绳绶纾纟丝鸶缌绥缩绦绨缇统彖纨绾维纬纹毋细纤弦线乡缃飨绡缬绁绣绪续绚幺疑彝绎缢肄引缨颍颖幽幼纡鬻缘约纭缯绽张缜织旨纸彘终粥纣绉缀缒缁综纵组缵';
	v_y       Varchar2(1200) := '哀庵谙廒鏖谤褒庇庳扁卞变遍斌禀亳诧谗廛谄颤昶谌谶诚充床鹑词诞谠诋底谛店调谍订读度憝敦讹谔方邡房访放扉诽废讽府腐讣该高膏诰庚赓诟诂雇诖广庋诡郭裹亥颃毫豪诃劾亨讧户戽扈话肓谎诙麾讳诲诨讥迹齑麂计记剂肩谫谏讲讦诘诫谨廑京旌扃就鹫讵诀谲麇康亢颏刻课库诓诳邝廓谰斓郎廊朗羸诔离戾廉娈恋良亮谅廖麟廪吝刘旒庐鹿旅膂率孪峦挛栾鸾脔銮论蠃麻蛮谩邙盲旄袤氓谜糜縻麋靡谧庙谬谟麽摩磨魔谋亩讷旎诺讴庞庖旆烹庀翩谝评裒谱齐旗麒启綮讫弃谦谴敲谯诮请庆诎诠瓤让认讪扇设麝诜谂诗施识市试谥孰塾熟庶衰谁说讼诵诉谡谇谈谭唐讨亭庭亠庹弯亡妄忘望为诿谓文紊诬庑误诶席襄详庠享谐亵谢廨庥许诩序畜谖玄旋谑询训讯讶讠言谚谳谣夜谒衣诒旖义议亦译诣奕弈谊裔应膺鹰嬴赢庸雍壅饔永诱於谀语庾育谕谮诈斋旃谵诏肇遮谪这鹧诊证诤衷州诌诛诸丶主麈庄谆诼谘诹卒族诅座';
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

