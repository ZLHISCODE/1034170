-----------------------------------------------------------------
--为配合产品版本号由9.28升为9.29
-----------------------------------------------------------------
--10925
Alter TABLE zlTools.zlRPTDatas Add 说明 VARCHAR2(1000)
/

--问题:10926
CREATE TABLE zlTools.zlRoleGroups(
    组名 VARCHAR2(30),
    角色 VARCHAR2(30))
    PCTFREE 5 PCTUSED 90 
    Cache Storage(Buffer_Pool Keep)
/

ALTER TABLE zlTools.zlRoleGroups ADD CONSTRAINT 
    zlRoleGroups_UQ_组名 UNIQUE (组名,角色)
    USING INDEX PCTFREE 5
/

create public synonym zlRoleGroups for zlTools.zlRoleGroups
/
GRANT Select on zlTools.zlRoleGroups to Public
/

Create Or Replace Package zlTools.b_Rolegroupmgr Is
  -----------------------------------------------------------------------------
  -- 作者： 刘兴宏
  -- 创始时间：2007/06/22
  -- 修改人：
  -- 修改时间：
  -- 描述：
  --         主要应用于角色组的处理
  -----------------------------------------------------------------------------

  Type t_Refcur Is Ref Cursor;
  -----------------------------------------------------------------------------
  -- 将角色分配给组
  -- 调用列表： frmRole.cmdAdd_Click
  -----------------------------------------------------------------------------
  Procedure Roletorolegroup
  (
    组名_In In Zlrolegroups.组名%Type,
    角色_In In Zlrolegroups.角色%Type := Null
  );

  -----------------------------------------------------------------------------
  -- 新建组
  -- 调用列表： frmRole.cmdNewGroup_Click
  -----------------------------------------------------------------------------
  Procedure Rolegroup_Add(组名_In In Zlrolegroups.组名%Type);

  -----------------------------------------------------------------------------
  -- 删除组
  -- 调用列表： frmRole[del]
  -----------------------------------------------------------------------------
  Procedure Rolegroup_Delete(组名_In In Zlrolegroups.组名%Type);

  -----------------------------------------------------------------------------
  -- 组更名
  -- 调用列表： frmRole[F2]
  -----------------------------------------------------------------------------
  Procedure Rolegroup_Rename
  (
    组名_Old_In In Zlrolegroups.组名%Type,
    组名_New_In In Zlrolegroups.组名%Type
  );

  -----------------------------------------------------------------------------
  -- 角色删除
  -- 调用列表： frmRole[del]
  -----------------------------------------------------------------------------
  Procedure Role_Delete(角色_In In Zlrolegroups.角色%Type);

  -----------------------------------------------------------------------------
  -- 角色拷贝
  -- 调用列表： frmRole[cmdCopy]
  -----------------------------------------------------------------------------
  Procedure Role_Copy
  (
    源角色_In   In zlRoleGrant.角色%Type,
    目标角色_In In zlRoleGrant.角色%Type
  );

End b_Rolegroupmgr;
/

Create Or Replace Package Body zlTools.b_Rolegroupmgr Is

  -----------------------------------------------------------------------------
  -- 功能：将角色分配给组
  -----------------------------------------------------------------------------
  Procedure Roletorolegroup
  (
    组名_In In Zlrolegroups.组名%Type,
    角色_In In Zlrolegroups.角色%Type := Null
  ) Is
  Begin
    Delete Zlrolegroups Where 角色 = 角色_In;
    If Not 组名_In Is Null Then
    
      Insert Into Zlrolegroups (组名, 角色) Values (组名_In, 角色_In);
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Roletorolegroup;

  -----------------------------------------------------------------------------
  -- 功能：新建组
  -----------------------------------------------------------------------------
  Procedure Rolegroup_Add(组名_In In Zlrolegroups.组名%Type) Is
    v_Err_Msg Varchar2(100);
    Err_Item Exception;
  Begin
    v_Err_Msg := Null;
    Begin
      Select 组名 Into v_Err_Msg From Zlrolegroups Where 组名 = 组名_In And Rownum = 1;
    Exception
      When Others Then
        v_Err_Msg := Null;
    End;
    If v_Err_Msg Is Not Null Then
      v_Err_Msg := '[ZLSOFT]组名为:' || 组名_In || '已经存在,不能增加此组,请检查[ZLSOFT]';
      Raise Err_Item;
    End If;
    Insert Into Zlrolegroups (组名, 角色) Values (组名_In, Null);
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Rolegroup_Add;

  -----------------------------------------------------------------------------
  -- 功能：删除组
  -----------------------------------------------------------------------------
  Procedure Rolegroup_Delete(组名_In In Zlrolegroups.组名%Type) Is
    v_Err_Msg Varchar2(100);
    Err_Item Exception;
  Begin
  
    Delete Zlrolegroups Where 组名 = 组名_In;
    If Sql%NotFound Then
      v_Err_Msg := '[ZLSOFT]组名为:' || 组名_In || '已经被他人删除,请检查[ZLSOFT]';
      Raise Err_Item;
    End If;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Rolegroup_Delete;

  -----------------------------------------------------------------------------
  -- 角色删除
  -- 调用列表： frmRole[del]
  -----------------------------------------------------------------------------
  Procedure Role_Delete(角色_In In Zlrolegroups.角色%Type) Is
    v_Err_Msg Varchar2(100);
    Err_Item Exception;
  Begin
  
    Delete Zlrolegroups Where 角色 = Nvl(角色_In, '|');
    Delete From zlRoleGrant Where 角色 = 角色_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Role_Delete;

  -----------------------------------------------------------------------------
  -- 功能：组更名
  -----------------------------------------------------------------------------
  Procedure Rolegroup_Rename
  (
    组名_Old_In In Zlrolegroups.组名%Type,
    组名_New_In In Zlrolegroups.组名%Type
  ) Is
    v_Err_Msg Varchar2(100);
    Err_Item Exception;
  Begin
    Update Zlrolegroups Set 组名 = 组名_New_In Where 组名 = 组名_Old_In;
    If Sql%NotFound Then
      Insert Into Zlrolegroups (组名, 角色) Values (组名_New_In, Null);
    
      -- v_Err_Msg := '[ZLSOFT]组名为:' || 组名_In || '不存在,可能已经被他人删除或修改,请检查[ZLSOFT]';
      -- Raise Err_Item;
    End If;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Rolegroup_Rename;

  -----------------------------------------------------------------------------
  -- 功能：角色复制
  -----------------------------------------------------------------------------
  Procedure Role_Copy
  (
    源角色_In   In zlRoleGrant.角色%Type,
    目标角色_In In zlRoleGrant.角色%Type
  ) Is
  Begin
    Begin
      Insert Into Zlrolegroups
        (组名, 角色)
        Select 组名, 目标角色_In From Zlrolegroups Where 角色 = 源角色_In;
    Exception
      When Others Then
        Null;
    End;
    Begin
      Insert Into zlRoleGrant
        (系统, 序号, 角色, 功能)
        Select 系统, 序号, 目标角色_In, 功能 From zlRoleGrant Where 角色 = 源角色_In;
    Exception
      When Others Then
        Null;
    End;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Role_Copy;

End b_Rolegroupmgr;
/

create public synonym b_Rolegroupmgr for zlTools.b_Rolegroupmgr
/
GRANT execute on zlTools.b_Rolegroupmgr to Public
/

-------------------------------------------------------------------------------
---- 新版体检要使用新连接,所以加入到正式脚本中 2007-07-06  Beging
-------------------------------------------------------------------------------

-------------------------------------------------------------------------------
-- 创建包头
-------------------------------------------------------------------------------

Create Or Replace Package zltools.b_Comfunc Is
  -----------------------------------------------------------------------------
  -- 作者： 陈东
  -- 创始时间：2006-8-9
  -- 修改人：
  -- 修改时间：
  -- 描述：
  --         主要用于公共部件的过程
  -----------------------------------------------------------------------------  
  Type t_Refcur Is Ref Cursor;

  -----------------------------------------------------------------------------
  -- 功能：保存错误日志
  -- 调用列表：
  -- clsComLib.SaveErrLog
  -----------------------------------------------------------------------------
  Procedure Save_Error_Log
  (
    类型_In     In zlErrorLog.类型%Type,
    错误序号_In In zlErrorLog.错误序号%Type,
    错误信息_In In zlErrorLog.错误信息%Type
  );

  -----------------------------------------------------------------------------
  -- 功能：取可用功能
  -- 调用列表：
  -- clsComLib.ShowAbout
  -----------------------------------------------------------------------------  
  Procedure Get_Usable_Function
  (
    Cursor_Out Out t_Refcur,
    部件_In    In zlPrograms.部件%Type
  );

  -----------------------------------------------------------------------------
  -- 功能：取大写金额
  -- 调用列表：
  -- clsCommFun.UppeMoney
  -----------------------------------------------------------------------------    
  Procedure Get_Uppmoney
  (
    Cursor_Out Out t_Refcur,
    金额_In    In Number
  );

  -----------------------------------------------------------------------------
  -- 功能：根据指定的日期、组号、系统判断指定日期的数据是否已转出到后备数据表中
  -- 调用列表：
  -- clsDatabase.DateMoved
  -----------------------------------------------------------------------------    
  Procedure Get_Datamoved
  (
    Cursor_Out  Out t_Refcur,
    组号_In     In zlDataMove.组号%Type,
    系统_In     In zlDataMove.系统%Type,
    上次日期_In In zlDataMove.上次日期%Type
  );

  -----------------------------------------------------------------------------
  -- 功能：取系统所有者
  -- 调用列表：
  -- clsDatabase.GetOwner
  -----------------------------------------------------------------------------
  Procedure Get_Owner
  (
    Cursor_Out Out t_Refcur,
    编号_In    In zlSystems.编号%Type
  );

  -----------------------------------------------------------------------------
  -- 功能：取简码
  -- 调用列表：
  -- clsCommFun.SpellCode
  -----------------------------------------------------------------------------
  Procedure Get_Spell_Code
  (
    Cursor_Out Out t_Refcur,
    字符串_In  In Varchar2,
    方式_In    In Number := 0
  );

  -----------------------------------------------------------------------------
  -- 功能：保存运行日志
  -- 调用列表：
  -- clsComLib.RestoreWinState
  -----------------------------------------------------------------------------
  Procedure Save_Diary_Log
  (
    部件名_In   In zlDiaryLog.部件名%Type,
    窗体名_In   In zlDiaryLog.窗体名%Type,
    工作内容_In In zlDiaryLog.工作内容%Type
  );

  -----------------------------------------------------------------------------
  -- 功能：更改运行日志
  -- 调用列表：
  -- clsComLib.SaveWinState
  -----------------------------------------------------------------------------
  Procedure Update_Diary_Log
  (
    部件名_In In zlDiaryLog.部件名%Type,
    窗体名_In In zlDiaryLog.窗体名%Type
  );

  -----------------------------------------------------------------------------
  -- 功能：取固定发布报表和用户发布报表
  -- 调用列表：
  -- clsDatabase.ShowReportMenu
  -----------------------------------------------------------------------------
  Procedure Get_Report_Menu
  (
    Cursor_Out Out t_Refcur,
    系统_In    In zlPrograms.系统%Type,
    序号_In    In zlPrograms.序号%Type,
    功能_In    In zlReports.功能%Type,
    编号_In    In zlReports.编号%Type
  );

  -----------------------------------------------------------------------------
  -- 功能：取用户提醒信息
  -- 调用列表：
  -- zlApptools.frmAlert
  -----------------------------------------------------------------------------
  Procedure Get_Zlnoticerec
  (
    Cursor_Out Out t_Refcur,
    用户名_In  In zlNoticeRec.用户名%Type
  );

  -----------------------------------------------------------------------------
  -- 功能：取邮件正文
  -- 调用列表：
  -- zlApptools.frmMessageEdit.LoadMessage
  -----------------------------------------------------------------------------
  Procedure Get_Zlmessage
  (
    Cursor_Out Out t_Refcur,
    Id_In      In zlMessages.ID%Type,
    类型_In    In zlMsgState.类型%Type,
    用户_In    In zlMsgState.用户%Type
  );

  -----------------------------------------------------------------------------
  -- 功能：取邮件内容
  -- 调用列表：
  -- zlApptools.frmMessageManager.FillText
  -- zlApptools.frmMessageRelate.FillText
  -----------------------------------------------------------------------------
  Procedure Get_Zlmessage
  (
    Cursor_Out Out t_Refcur,
    Id_In      In zlMessages.ID%Type
  );

  -----------------------------------------------------------------------------
  -- 功能：取邮递地址
  -- 调用列表：
  -- zlApptools.frmMessageEdit.LoadMessage
  -----------------------------------------------------------------------------
  Procedure Get_Zlmsgstate
  (
    Cursor_Out Out t_Refcur,
    消息id_In  In zlMsgState.消息id%Type
  );

  -----------------------------------------------------------------------------
  -- 功能：删除消息
  -- 调用列表：
  -- zlApptools.frmMessageManager.mnuEditDelete_Click
  -----------------------------------------------------------------------------
  Procedure Delete_Zlmsgstate
  (
    删除_In   In zlMsgState.删除%Type,
    消息id_In In zlMsgState.消息id%Type,
    类型_In   In zlMsgState.类型%Type,
    用户_In   In zlMsgState.用户%Type
  );

  -----------------------------------------------------------------------------
  -- 功能：删除过期消息
  -- 调用列表：
  -- zlApptools.frmMessageManager.DeleteMessage
  -----------------------------------------------------------------------------
  Procedure Delete_Zlmessage;

  -----------------------------------------------------------------------------
  -- 功能：取邮件列表
  -- 调用列表：
  -- zlApptools.frmMessageManager.FillList
  -- zlApptools.frmMessageRelate.FillList
  -----------------------------------------------------------------------------
  Procedure Get_Mail_List
  (
    Cursor_Out  Out t_Refcur,
    消息类型_In In Varchar2,
    用户_In     In zlMsgState.用户%Type,
    显示已读_In In Number,
    会话id_In   In zlMessages.会话id%Type    
  );

  -----------------------------------------------------------------------------
  -- 功能：还原删除的消息
  -- 调用列表：
  -- zlApptools.frmMessageManager.mnuEditRestore_Click
  -----------------------------------------------------------------------------
  Procedure Restore_Zlmsgstate
  (
    消息id_In In zlMsgState.消息id%Type,
    类型_In   In zlMsgState.类型%Type,
    用户_In   In zlMsgState.用户%Type
  );

  -----------------------------------------------------------------------------
  -- 功能：保存主表消息
  -- 调用列表：
  -- zlApptools.frmMessageEdit.SaveMessage
  -----------------------------------------------------------------------------
  Procedure Save_Zlmessage
  (
    Cursor_Out Out t_Refcur,
    Id_In      In zlMessages.ID%Type,
    会话id_In  In zlMessages.会话id%Type,
    发件人_In  In zlMessages.发件人%Type,
    收件人_In  In zlMessages.收件人%Type,
    主题_In    In zlMessages.主题%Type,
    内容_In    In zlMessages.内容%Type,
    背景色_In  In zlMessages.背景色%Type
  );

  -----------------------------------------------------------------------------
  -- 功能：插入zlMsgstate
  -- 调用列表：
  -- zlApptools.frmMessageEdit.SaveMessage
  -----------------------------------------------------------------------------
  Procedure Insert_Zlmsgstate
  (
    消息id_In In zlMsgState.消息id%Type,
    类型_In   In zlMsgState.类型%Type,
    用户_In   In zlMsgState.用户%Type,
    身份_In   In zlMsgState.身份%Type,
    删除_In   In zlMsgState.删除%Type,
    状态_In   In zlMsgState.状态%Type
  );

  -----------------------------------------------------------------------------
  -- 功能：为原件加上答复或转发标志
  -- 调用列表：
  -- zlApptools.frmMessageEdit.SaveMessage
  -----------------------------------------------------------------------------
  Procedure Update_Zlmsgstate_State
  (
    模式_In   In Number,
    消息id_In In zlMsgState.消息id%Type,
    类型_In   In zlMsgState.类型%Type,
    用户_In   In zlMsgState.用户%Type
  );

  -----------------------------------------------------------------------------
  -- 功能：为原件加上答复或转发标志
  -- 调用列表：
  -- zlApptools.frmMessageEdit.LoadMessage
  -----------------------------------------------------------------------------
  Procedure Update_Zlmsgstate_Idtntify
  (
    身份_In   In zlMsgState.身份%Type,
    消息id_In In zlMsgState.消息id%Type,
    类型_In   In zlMsgState.类型%Type,
    用户_In   In zlMsgState.用户%Type
  );

End b_Comfunc;
/

-------------------------------------------------------------------------------
-- 创建包体
-------------------------------------------------------------------------------
Create Or Replace Package Body zltools.b_Comfunc Is

  -----------------------------------------------------------------------------
  -- 功能：保存错误日志
  -----------------------------------------------------------------------------
  Procedure Save_Error_Log
  (
    类型_In     In zlErrorLog.类型%Type,
    错误序号_In In zlErrorLog.错误序号%Type,
    错误信息_In In zlErrorLog.错误信息%Type
  ) Is
  Begin
    Insert Into zlErrorLog
      (会话号, 用户名, 工作站, 时间, 类型, 错误序号, 错误信息)
      Select Sid, User, Machine, Sysdate, 类型_In, 错误序号_In, 错误信息_In
      From V$session
      Where Audsid = Userenv('SessionID');
    Commit;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Save_Error_Log;

  -----------------------------------------------------------------------------
  -- 功能：取可用功能
  -----------------------------------------------------------------------------
  Procedure Get_Usable_Function
  (
    Cursor_Out Out t_Refcur,
    部件_In    In zlPrograms.部件%Type
  ) Is
  Begin
    If Nvl(部件_In, '空空') = '空空' Then
      Open Cursor_Out For
        Select Distinct A.序号, A.标题, A.说明
        From zlPrograms A, zlProgFuncs B, Zlregfunc C
        Where A.系统 = B.系统 And A.序号 = B.序号 And Trunc(A.系统 / 100) = C.系统 And A.序号 = C.序号
        Order By A.序号;
    Else
      Open Cursor_Out For
        Select Distinct A.序号, A.标题, A.说明
        From zlPrograms A, zlProgFuncs B, Zlregfunc C
        Where A.系统 = B.系统 And A.序号 = B.序号 And Upper(A.部件) = Upper(部件_In) And Trunc(A.系统 / 100) = C.系统 And
              A.序号 = C.序号
        Order By A.序号;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Usable_Function;

  -----------------------------------------------------------------------------
  -- 功能：取大写金额
  -----------------------------------------------------------------------------    
  Procedure Get_Uppmoney
  (
    Cursor_Out Out t_Refcur,
    金额_In    In Number
  ) Is
  Begin
    Open Cursor_Out For
      Select zlUppMoney(Nvl(金额_In, 0)) As Num From Dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Uppmoney;

  -----------------------------------------------------------------------------
  -- 功能：根据指定的日期、组号、系统判断指定日期的数据是否已转出到后备数据表中
  -----------------------------------------------------------------------------    
  Procedure Get_Datamoved
  (
    Cursor_Out  Out t_Refcur,
    组号_In     In zlDataMove.组号%Type,
    系统_In     In zlDataMove.系统%Type,
    上次日期_In In zlDataMove.上次日期%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select 系统, 组号
      From zlDataMove
      Where 组号 = 组号_In And 系统 = 系统_In And 上次日期 > 上次日期_In And 上次日期 Is Not Null;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Datamoved;

  -----------------------------------------------------------------------------
  -- 功能：取系统所有者
  -----------------------------------------------------------------------------
  Procedure Get_Owner
  (
    Cursor_Out Out t_Refcur,
    编号_In    In zlSystems.编号%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select 所有者 From zlSystems Where 编号 = 编号_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Owner;

  -----------------------------------------------------------------------------
  -- 功能：取简码
  -----------------------------------------------------------------------------
  Procedure Get_Spell_Code
  (
    Cursor_Out Out t_Refcur,
    字符串_In  In Varchar2,
    方式_In    In Number := 0
  ) Is
  Begin
    If Nvl(方式_In, 0) = 0 Then
      Open Cursor_Out For
        Select zlSpellCode(字符串_In) As 简码 From Dual;
    Else
      Open Cursor_Out For
        Select zlWbCode(字符串_In) As 简码 From Dual;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Spell_Code;

  -----------------------------------------------------------------------------
  -- 功能：保存运行日志
  -----------------------------------------------------------------------------
  Procedure Save_Diary_Log
  (
    部件名_In   In zlDiaryLog.部件名%Type,
    窗体名_In   In zlDiaryLog.窗体名%Type,
    工作内容_In In zlDiaryLog.工作内容%Type
  ) Is
  Begin
    Insert Into zlDiaryLog
      (会话号, 用户名, 工作站, 部件名, 窗体名, 工作内容, 进入时间)
      Select Sid + Serial#, User, RTrim(LTrim(Replace(Machine, Chr(0), ''))), 部件名_In, 窗体名_In, 工作内容_In, Sysdate
      From V$session
      Where Audsid = Userenv('SessionID');
    Commit;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Save_Diary_Log;

  -----------------------------------------------------------------------------
  -- 功能：更改运行日志
  -- 调用列表：
  -- clsComLib.SaveWinState
  -----------------------------------------------------------------------------
  Procedure Update_Diary_Log
  (
    部件名_In In zlDiaryLog.部件名%Type,
    窗体名_In In zlDiaryLog.窗体名%Type
  ) Is
    Cursor c_Session Is
      Select Sid + Serial# As 会话号, User As 用户名, RTrim(LTrim(Replace(Machine, Chr(0), ''))) As 工作站
      From V$session
      Where Audsid = Userenv('SessionID');
  Begin
    For r_Tmp In c_Session Loop
      Update zlDiaryLog
      Set 退出原因 = 1, 退出时间 = Sysdate
      Where 退出原因 Is Null And 用户名 = r_Tmp.用户名 And 工作站 = r_Tmp.工作站 And 会话号 = r_Tmp.会话号 And
            部件名 = 部件名_In And 窗体名 = 窗体名_In;
    End Loop;
    Commit;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Update_Diary_Log;

  -----------------------------------------------------------------------------
  -- 功能：取固定发布报表和用户发布报表
  -----------------------------------------------------------------------------
  Procedure Get_Report_Menu
  (
    Cursor_Out Out t_Refcur,
    系统_In    In zlPrograms.系统%Type,
    序号_In    In zlPrograms.序号%Type,
    功能_In    In zlReports.功能%Type,
    编号_In    In zlReports.编号%Type
  ) Is
  Begin
    If Nvl(编号_In, '空空') <> '空空' Then
      Open Cursor_Out For
        Select 标志, 系统, 编号, 名称
        From (Select 1 As 标志, A.系统, A.编号, A.名称
               From zlReports A, zlPrograms B
               Where A.系统 = B.系统 And A.程序id = B.序号 And Not Upper(A.编号) Like '%BILL%' And
                     Upper(B.部件) <> Upper('zl9Report') And B.系统 = 系统_In And B.序号 = 序号_In And
                     Instr(功能_In, ';' || A.功能 || ';') > 0
               Union All
               Select Decode(A.系统, Null, 2, 1) As 标志, A.系统, A.编号, A.名称
               From zlReports A, zlRPTPuts B, zlPrograms C
               Where A.ID = B.报表id And B.系统 = C.系统 And B.程序id = C.序号 And
                     (Not Upper(A.编号) Like '%BILL%' Or A.系统 Is Null) And Instr(功能_In, ';' || B.功能 || ';') > 0 And
                     C.系统 = 系统_In And C.序号 = 序号_In)
        Where Instr(编号_In, ',' || 编号 || ',') = 0
        Order By 标志, 编号;
    Else
      Open Cursor_Out For
        Select 标志, 系统, 编号, 名称
        From (Select 1 As 标志, A.系统, A.编号, A.名称
               From zlReports A, zlPrograms B
               Where A.系统 = B.系统 And A.程序id = B.序号 And Not Upper(A.编号) Like '%BILL%' And
                     Upper(B.部件) <> Upper('zl9Report') And B.系统 = 系统_In And B.序号 = 序号_In And
                     Instr(功能_In, ';' || A.功能 || ';') > 0
               Union All
               Select Decode(A.系统, Null, 2, 1) As 标志, A.系统, A.编号, A.名称
               From zlReports A, zlRPTPuts B, zlPrograms C
               Where A.ID = B.报表id And B.系统 = C.系统 And B.程序id = C.序号 And
                     (Not Upper(A.编号) Like '%BILL%' Or A.系统 Is Null) And Instr(功能_In, ';' || B.功能 || ';') > 0 And
                     C.系统 = 系统_In And C.序号 = 序号_In)
        Order By 标志, 编号;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Report_Menu;

  -----------------------------------------------------------------------------
  -- 功能：取用户提醒信息
  -----------------------------------------------------------------------------
  Procedure Get_Zlnoticerec
  (
    Cursor_Out Out t_Refcur,
    用户名_In  In zlNoticeRec.用户名%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select A.序号, A.系统, C.程序id As 模块, B.提醒内容 As 结果内容, C.名称 As 提醒报表, A.提醒声音, B.检查时间,
             B.已读标志
      From zlNotices A, zlNoticeRec B, (Select * From zlReports Where 发布时间 Is Not Null) C
      Where B.用户名 = 用户名_In And B.提醒标志 > 0 And C.编号(+) = A.提醒报表 And A.序号 = B.提醒序号 And
            B.提醒内容 Is Not Null;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Zlnoticerec;

  -----------------------------------------------------------------------------
  -- 功能：取邮件正文
  -----------------------------------------------------------------------------
  Procedure Get_Zlmessage
  (
    Cursor_Out Out t_Refcur,
    Id_In      In zlMessages.ID%Type,
    类型_In    In zlMsgState.类型%Type,
    用户_In    In zlMsgState.用户%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select A.*, B.删除, B.状态
      From zlMessages A, zlMsgState B
      Where A.ID = B.消息id And B.消息id = Id_In And B.类型 = 类型_In And B.用户 = 用户_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Zlmessage;

  -----------------------------------------------------------------------------
  -- 功能：取邮件内容
  -----------------------------------------------------------------------------
  Procedure Get_Zlmessage
  (
    Cursor_Out Out t_Refcur,
    Id_In      In zlMessages.ID%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select 内容, 背景色 From zlMessages Where ID = Id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Zlmessage;

  -----------------------------------------------------------------------------
  -- 功能：取邮递地址
  -----------------------------------------------------------------------------
  Procedure Get_Zlmsgstate
  (
    Cursor_Out Out t_Refcur,
    消息id_In  In zlMsgState.消息id%Type
  ) Is
  Begin
    Open Cursor_Out For
      Select 类型, 用户, 身份 From zlMsgState Where 消息id = 消息id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Zlmsgstate;

  -----------------------------------------------------------------------------
  -- 功能：删除消息
  -----------------------------------------------------------------------------
  Procedure Delete_Zlmsgstate
  (
    删除_In   In zlMsgState.删除%Type,
    消息id_In In zlMsgState.消息id%Type,
    类型_In   In zlMsgState.类型%Type,
    用户_In   In zlMsgState.用户%Type
  ) Is
    n_总数 Number(10);
    n_数量 Number(10);
  Begin
    If Nvl(删除_In, 0) = 1 Then
      Update zlMsgState Set 删除 = 1 Where 消息id = 消息id_In And 类型 = 类型_In And 用户 = 用户_In;
    Else
      If 类型_In = 0 Then
        -- 对于草稿，把收件人的也一并删除
        Update zlMsgState Set 删除 = 2 Where 消息id = 消息id_In And 用户 = 用户_In;
      Else
        Update zlMsgState Set 删除 = 2 Where 消息id = 消息id_In And 类型 = 类型_In And 用户 = 用户_In;
      End If;
      --  删除指定ID的消息  mnuEditDelete_Click 调用
      Select Count(*) As 总数, Sum(Decode(删除, 2, 1, 0)) As 数量
      Into n_总数, n_数量
      From zlMsgState
      Where 消息id = 消息id_In;
    
      If n_总数 = n_数量 Then
        Delete From zlMessages Where ID = 消息id_In;
      End If;
    End If;
    Commit;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Delete_Zlmsgstate;

  -----------------------------------------------------------------------------
  -- 功能：删除过期消息
  -----------------------------------------------------------------------------
  Procedure Delete_Zlmessage Is
    n_Days Number;
  Begin
    Select Nvl(参数值, 缺省值) Into n_Days From zlOptions Where 参数号 = 5;
    If Nvl(n_Days, 0) > 0 Then
      Delete From zlMessages Where 时间 < Sysdate - n_Days;
      Commit;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Delete_Zlmessage;

  -----------------------------------------------------------------------------
  -- 功能：取邮件列表
  -----------------------------------------------------------------------------
  Procedure Get_Mail_List
  (
    Cursor_Out  Out t_Refcur,
    消息类型_In In Varchar2,
    用户_In     In zlMsgState.用户%Type,
    显示已读_In In Number,
    会话id_In   In zlMessages.会话id%Type
  ) Is
    v_Sql  Varchar2(1000);
    v_已读 Varchar2(100);
    v_类型 Varchar2(100);
  Begin
  
    If Nvl(显示已读_In, 0) = 1 Then
      v_已读 := ' and substr(S.状态,1,1)=''0''';
    Else
      v_已读 := '';
    End If;
  
    If Instr(';草稿;收件箱;已发送消息;已删除消息;相关消息;', ';' || 消息类型_In || ';') <= 0 Then
      v_类型 := '草稿';
    Else
      v_类型 := 消息类型_In;
    End If;
  
    If v_类型 = '草稿' Then
      v_Sql := 'Select M.ID, M.会话id, M.发件人, M.收件人, M.主题, To_Char(M.时间, ''YYYY-MM-DD HH24:MI:SS'') As 时间, S.类型, S.状态 
              From zlMessages M, zlMsgState S
              Where M.ID = S.消息id  and S.删除=0 and S.用户= ''' || 用户_In || ''' And S.类型=0 ' || v_已读;
    End If;
  
    If v_类型 = '收件箱' Then
      v_Sql := 'Select M.ID, M.会话id, M.发件人, M.收件人, M.主题, To_Char(M.时间, ''YYYY-MM-DD HH24:MI:SS'') As 时间, S.类型, S.状态 
              From zlMessages M, zlMsgState S
              Where M.ID = S.消息id  and S.删除=0 and S.用户= ''' || 用户_In || ''' And S.类型=2 ' || v_已读;
    End If;
  
    If v_类型 = '已发送消息' Then
      v_Sql := 'Select M.ID, M.会话id, M.发件人, M.收件人, M.主题, To_Char(M.时间, ''YYYY-MM-DD HH24:MI:SS'') As 时间, S.类型, S.状态 
              From zlMessages M, zlMsgState S
              Where M.ID = S.消息id  and S.删除=0 and S.用户= ''' || 用户_In || ''' And S.类型=1 ' || v_已读;
    End If;
  
    If v_类型 = '已删除消息' Then
      v_Sql := 'Select M.ID, M.会话id, M.发件人, M.收件人, M.主题, To_Char(M.时间, ''YYYY-MM-DD HH24:MI:SS'') As 时间, S.类型, S.状态 
              From zlMessages M, zlMsgState S
              Where M.ID = S.消息id  and S.用户= ''' || 用户_In || ''' And S.删除=1 ' || v_已读;
    End If;
  
    If v_类型 = '相关消息' Then
      v_Sql := 'select M.ID,M.会话ID,M.发件人,M.收件人,M.主题,to_char(M.时间,''YYYY-MM-DD HH24:MI:SS'') as 时间,S.类型,S.状态
         from zlMessages M,zlMsgState S where M.ID=S.消息ID and S.删除<>2 and S.用户= ''' || 用户_In ||
               '''  and M.会话ID=' || 会话id_In;
    End If;
  
    If Nvl(v_Sql, '空空') <> '空空' Then
      Open Cursor_Out For v_Sql;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Get_Mail_List;

  -----------------------------------------------------------------------------
  -- 功能：还原删除的消息
  -----------------------------------------------------------------------------
  Procedure Restore_Zlmsgstate
  (
    消息id_In In zlMsgState.消息id%Type,
    类型_In   In zlMsgState.类型%Type,
    用户_In   In zlMsgState.用户%Type
  ) Is
  Begin
    Update zlMsgState Set 删除 = 0 Where 消息id = 消息id_In And 类型 = 类型_In And 用户 = 用户_In;
    Commit;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Restore_Zlmsgstate;

  -----------------------------------------------------------------------------
  -- 功能：保存消息
  -- 调用列表：
  -- zlApptools.frmMessageEdit.SaveMessage
  -----------------------------------------------------------------------------
  Procedure Save_Zlmessage
  (
    Cursor_Out Out t_Refcur,
    Id_In      In zlMessages.ID%Type,
    会话id_In  In zlMessages.会话id%Type,
    发件人_In  In zlMessages.发件人%Type,
    收件人_In  In zlMessages.收件人%Type,
    主题_In    In zlMessages.主题%Type,
    内容_In    In zlMessages.内容%Type,
    背景色_In  In zlMessages.背景色%Type
  ) Is
    n_Id     zlMessages.ID%Type;
    n_会话id zlMessages.会话id%Type;
  Begin
    If Nvl(Id_In, 0) = 0 Then
      Select Zlmessages_Id.Nextval Into n_Id From Dual;
      n_Id := Nvl(n_Id, 0);
      If Nvl(会话id_In, 0) = 0 Then
        n_会话id := n_Id;
      Else
        n_会话id := 会话id_In;
      End If;
      Insert Into zlMessages
        (ID, 会话id, 发件人, 时间, 收件人, 主题, 内容, 背景色)
      Values
        (n_Id, n_会话id, 发件人_In, Sysdate, 收件人_In, 主题_In, 内容_In, 背景色_In);
      Open Cursor_Out For
        Select n_Id As ID, n_会话id As 会话id From Dual;
    Else
      Update zlMessages
      Set 发件人 = 发件人_In, 时间 = Sysdate, 收件人 = 收件人_In, 主题 = 主题_In, 内容 = 内容_In, 背景色 = 背景色_In
      Where ID = Id_In;
      Open Cursor_Out For
        Select Id_In As ID, 会话id_In As 会话id From Dual;
    End If;
    Commit;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Save_Zlmessage;

  -----------------------------------------------------------------------------
  -- 功能：插入zlMsgstate
  -- 调用列表：
  -- zlApptools.frmMessageEdit.SaveMessage
  -----------------------------------------------------------------------------
  Procedure Insert_Zlmsgstate
  (
    消息id_In In zlMsgState.消息id%Type,
    类型_In   In zlMsgState.类型%Type,
    用户_In   In zlMsgState.用户%Type,
    身份_In   In zlMsgState.身份%Type,
    删除_In   In zlMsgState.删除%Type,
    状态_In   In zlMsgState.状态%Type
  ) Is
  Begin
  
    If 类型_In < 2 Then
      Delete From zlMsgState Where 消息id = 消息id_In;
    End If;
    Insert Into zlMsgState
      (消息id, 类型, 用户, 身份, 删除, 状态)
    Values
      (消息id_In, 类型_In, 用户_In, 身份_In, 删除_In, 状态_In);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Insert_Zlmsgstate;

  -----------------------------------------------------------------------------
  -- 功能：为原件加上答复或转发标志
  -----------------------------------------------------------------------------
  Procedure Update_Zlmsgstate_State
  (
    模式_In   In Number,
    消息id_In In zlMsgState.消息id%Type,
    类型_In   In zlMsgState.类型%Type,
    用户_In   In zlMsgState.用户%Type
  ) Is
  Begin
    If Nvl(模式_In, 0) = 1 Or Nvl(模式_In, 0) = 2 Then
      Update zlMsgState
      Set 状态 = Substr(状态, 1, 1) || '1' || Substr(状态, 3, 2)
      Where 消息id = 消息id_In And 类型 = 类型_In And 用户 = 用户_In;
      Commit;
    End If;
    If Nvl(模式_In, 0) = 3 Then
      Update zlMsgState
      Set 状态 = Substr(状态, 1, 1) || '1' || Substr(状态, 4, 1)
      Where 消息id = 消息id_In And 类型 = 类型_In And 用户 = 用户_In;
      Commit;
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Update_Zlmsgstate_State;

  -----------------------------------------------------------------------------
  -- 功能：更新状态和身份
  -----------------------------------------------------------------------------
  Procedure Update_Zlmsgstate_Idtntify
  (
    身份_In   In zlMsgState.身份%Type,
    消息id_In In zlMsgState.消息id%Type,
    类型_In   In zlMsgState.类型%Type,
    用户_In   In zlMsgState.用户%Type
  ) Is
  Begin
    Update zlMsgState
    Set 状态 = '1' || Substr(状态, 2), 身份 = 身份_In
    Where 消息id = 消息id_In And 类型 = 类型_In And 用户 = 用户_In;
    Commit;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Update_Zlmsgstate_Idtntify;

End b_Comfunc;
/

-------------------------------------------------------------------------------
-- 授权
-------------------------------------------------------------------------------
grant execute on zltools.b_ComFunc to PUBLIC
/
--- 不创建同义词，是因为要求程序中加所有者前缀来调用。
-------------------------------------------------------------------------------
---- 新版体检要使用新连接,所以加入到正式脚本中 2007-07-06  End
-------------------------------------------------------------------------------