-----------------------------------------------------------------
--Ϊ��ϲ�Ʒ�汾����9.28��Ϊ9.29
-----------------------------------------------------------------
--10925
Alter TABLE zlTools.zlRPTDatas Add ˵�� VARCHAR2(1000)
/

--����:10926
CREATE TABLE zlTools.zlRoleGroups(
    ���� VARCHAR2(30),
    ��ɫ VARCHAR2(30))
    PCTFREE 5 PCTUSED 90 
    Cache Storage(Buffer_Pool Keep)
/

ALTER TABLE zlTools.zlRoleGroups ADD CONSTRAINT 
    zlRoleGroups_UQ_���� UNIQUE (����,��ɫ)
    USING INDEX PCTFREE 5
/

create public synonym zlRoleGroups for zlTools.zlRoleGroups
/
GRANT Select on zlTools.zlRoleGroups to Public
/

Create Or Replace Package zlTools.b_Rolegroupmgr Is
  -----------------------------------------------------------------------------
  -- ���ߣ� ���˺�
  -- ��ʼʱ�䣺2007/06/22
  -- �޸��ˣ�
  -- �޸�ʱ�䣺
  -- ������
  --         ��ҪӦ���ڽ�ɫ��Ĵ���
  -----------------------------------------------------------------------------

  Type t_Refcur Is Ref Cursor;
  -----------------------------------------------------------------------------
  -- ����ɫ�������
  -- �����б� frmRole.cmdAdd_Click
  -----------------------------------------------------------------------------
  Procedure Roletorolegroup
  (
    ����_In In Zlrolegroups.����%Type,
    ��ɫ_In In Zlrolegroups.��ɫ%Type := Null
  );

  -----------------------------------------------------------------------------
  -- �½���
  -- �����б� frmRole.cmdNewGroup_Click
  -----------------------------------------------------------------------------
  Procedure Rolegroup_Add(����_In In Zlrolegroups.����%Type);

  -----------------------------------------------------------------------------
  -- ɾ����
  -- �����б� frmRole[del]
  -----------------------------------------------------------------------------
  Procedure Rolegroup_Delete(����_In In Zlrolegroups.����%Type);

  -----------------------------------------------------------------------------
  -- �����
  -- �����б� frmRole[F2]
  -----------------------------------------------------------------------------
  Procedure Rolegroup_Rename
  (
    ����_Old_In In Zlrolegroups.����%Type,
    ����_New_In In Zlrolegroups.����%Type
  );

  -----------------------------------------------------------------------------
  -- ��ɫɾ��
  -- �����б� frmRole[del]
  -----------------------------------------------------------------------------
  Procedure Role_Delete(��ɫ_In In Zlrolegroups.��ɫ%Type);

  -----------------------------------------------------------------------------
  -- ��ɫ����
  -- �����б� frmRole[cmdCopy]
  -----------------------------------------------------------------------------
  Procedure Role_Copy
  (
    Դ��ɫ_In   In zlRoleGrant.��ɫ%Type,
    Ŀ���ɫ_In In zlRoleGrant.��ɫ%Type
  );

End b_Rolegroupmgr;
/

Create Or Replace Package Body zlTools.b_Rolegroupmgr Is

  -----------------------------------------------------------------------------
  -- ���ܣ�����ɫ�������
  -----------------------------------------------------------------------------
  Procedure Roletorolegroup
  (
    ����_In In Zlrolegroups.����%Type,
    ��ɫ_In In Zlrolegroups.��ɫ%Type := Null
  ) Is
  Begin
    Delete Zlrolegroups Where ��ɫ = ��ɫ_In;
    If Not ����_In Is Null Then
    
      Insert Into Zlrolegroups (����, ��ɫ) Values (����_In, ��ɫ_In);
    End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Roletorolegroup;

  -----------------------------------------------------------------------------
  -- ���ܣ��½���
  -----------------------------------------------------------------------------
  Procedure Rolegroup_Add(����_In In Zlrolegroups.����%Type) Is
    v_Err_Msg Varchar2(100);
    Err_Item Exception;
  Begin
    v_Err_Msg := Null;
    Begin
      Select ���� Into v_Err_Msg From Zlrolegroups Where ���� = ����_In And Rownum = 1;
    Exception
      When Others Then
        v_Err_Msg := Null;
    End;
    If v_Err_Msg Is Not Null Then
      v_Err_Msg := '[ZLSOFT]����Ϊ:' || ����_In || '�Ѿ�����,�������Ӵ���,����[ZLSOFT]';
      Raise Err_Item;
    End If;
    Insert Into Zlrolegroups (����, ��ɫ) Values (����_In, Null);
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Rolegroup_Add;

  -----------------------------------------------------------------------------
  -- ���ܣ�ɾ����
  -----------------------------------------------------------------------------
  Procedure Rolegroup_Delete(����_In In Zlrolegroups.����%Type) Is
    v_Err_Msg Varchar2(100);
    Err_Item Exception;
  Begin
  
    Delete Zlrolegroups Where ���� = ����_In;
    If Sql%NotFound Then
      v_Err_Msg := '[ZLSOFT]����Ϊ:' || ����_In || '�Ѿ�������ɾ��,����[ZLSOFT]';
      Raise Err_Item;
    End If;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Rolegroup_Delete;

  -----------------------------------------------------------------------------
  -- ��ɫɾ��
  -- �����б� frmRole[del]
  -----------------------------------------------------------------------------
  Procedure Role_Delete(��ɫ_In In Zlrolegroups.��ɫ%Type) Is
    v_Err_Msg Varchar2(100);
    Err_Item Exception;
  Begin
  
    Delete Zlrolegroups Where ��ɫ = Nvl(��ɫ_In, '|');
    Delete From zlRoleGrant Where ��ɫ = ��ɫ_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Role_Delete;

  -----------------------------------------------------------------------------
  -- ���ܣ������
  -----------------------------------------------------------------------------
  Procedure Rolegroup_Rename
  (
    ����_Old_In In Zlrolegroups.����%Type,
    ����_New_In In Zlrolegroups.����%Type
  ) Is
    v_Err_Msg Varchar2(100);
    Err_Item Exception;
  Begin
    Update Zlrolegroups Set ���� = ����_New_In Where ���� = ����_Old_In;
    If Sql%NotFound Then
      Insert Into Zlrolegroups (����, ��ɫ) Values (����_New_In, Null);
    
      -- v_Err_Msg := '[ZLSOFT]����Ϊ:' || ����_In || '������,�����Ѿ�������ɾ�����޸�,����[ZLSOFT]';
      -- Raise Err_Item;
    End If;
  Exception
    When Err_Item Then
      Raise_Application_Error(-20101, v_Err_Msg);
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End Rolegroup_Rename;

  -----------------------------------------------------------------------------
  -- ���ܣ���ɫ����
  -----------------------------------------------------------------------------
  Procedure Role_Copy
  (
    Դ��ɫ_In   In zlRoleGrant.��ɫ%Type,
    Ŀ���ɫ_In In zlRoleGrant.��ɫ%Type
  ) Is
  Begin
    Begin
      Insert Into Zlrolegroups
        (����, ��ɫ)
        Select ����, Ŀ���ɫ_In From Zlrolegroups Where ��ɫ = Դ��ɫ_In;
    Exception
      When Others Then
        Null;
    End;
    Begin
      Insert Into zlRoleGrant
        (ϵͳ, ���, ��ɫ, ����)
        Select ϵͳ, ���, Ŀ���ɫ_In, ���� From zlRoleGrant Where ��ɫ = Դ��ɫ_In;
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
---- �°����Ҫʹ��������,���Լ��뵽��ʽ�ű��� 2007-07-06  Beging
-------------------------------------------------------------------------------

-------------------------------------------------------------------------------
-- ������ͷ
-------------------------------------------------------------------------------

Create Or Replace Package zltools.b_Comfunc Is
  -----------------------------------------------------------------------------
  -- ���ߣ� �¶�
  -- ��ʼʱ�䣺2006-8-9
  -- �޸��ˣ�
  -- �޸�ʱ�䣺
  -- ������
  --         ��Ҫ���ڹ��������Ĺ���
  -----------------------------------------------------------------------------  
  Type t_Refcur Is Ref Cursor;

  -----------------------------------------------------------------------------
  -- ���ܣ����������־
  -- �����б�
  -- clsComLib.SaveErrLog
  -----------------------------------------------------------------------------
  Procedure Save_Error_Log
  (
    ����_In     In zlErrorLog.����%Type,
    �������_In In zlErrorLog.�������%Type,
    ������Ϣ_In In zlErrorLog.������Ϣ%Type
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ���ù���
  -- �����б�
  -- clsComLib.ShowAbout
  -----------------------------------------------------------------------------  
  Procedure Get_Usable_Function
  (
    Cursor_Out Out t_Refcur,
    ����_In    In zlPrograms.����%Type
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ��д���
  -- �����б�
  -- clsCommFun.UppeMoney
  -----------------------------------------------------------------------------    
  Procedure Get_Uppmoney
  (
    Cursor_Out Out t_Refcur,
    ���_In    In Number
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�����ָ�������ڡ���š�ϵͳ�ж�ָ�����ڵ������Ƿ���ת���������ݱ���
  -- �����б�
  -- clsDatabase.DateMoved
  -----------------------------------------------------------------------------    
  Procedure Get_Datamoved
  (
    Cursor_Out  Out t_Refcur,
    ���_In     In zlDataMove.���%Type,
    ϵͳ_In     In zlDataMove.ϵͳ%Type,
    �ϴ�����_In In zlDataMove.�ϴ�����%Type
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡϵͳ������
  -- �����б�
  -- clsDatabase.GetOwner
  -----------------------------------------------------------------------------
  Procedure Get_Owner
  (
    Cursor_Out Out t_Refcur,
    ���_In    In zlSystems.���%Type
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ����
  -- �����б�
  -- clsCommFun.SpellCode
  -----------------------------------------------------------------------------
  Procedure Get_Spell_Code
  (
    Cursor_Out Out t_Refcur,
    �ַ���_In  In Varchar2,
    ��ʽ_In    In Number := 0
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�����������־
  -- �����б�
  -- clsComLib.RestoreWinState
  -----------------------------------------------------------------------------
  Procedure Save_Diary_Log
  (
    ������_In   In zlDiaryLog.������%Type,
    ������_In   In zlDiaryLog.������%Type,
    ��������_In In zlDiaryLog.��������%Type
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�����������־
  -- �����б�
  -- clsComLib.SaveWinState
  -----------------------------------------------------------------------------
  Procedure Update_Diary_Log
  (
    ������_In In zlDiaryLog.������%Type,
    ������_In In zlDiaryLog.������%Type
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ�̶�����������û���������
  -- �����б�
  -- clsDatabase.ShowReportMenu
  -----------------------------------------------------------------------------
  Procedure Get_Report_Menu
  (
    Cursor_Out Out t_Refcur,
    ϵͳ_In    In zlPrograms.ϵͳ%Type,
    ���_In    In zlPrograms.���%Type,
    ����_In    In zlReports.����%Type,
    ���_In    In zlReports.���%Type
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ�û�������Ϣ
  -- �����б�
  -- zlApptools.frmAlert
  -----------------------------------------------------------------------------
  Procedure Get_Zlnoticerec
  (
    Cursor_Out Out t_Refcur,
    �û���_In  In zlNoticeRec.�û���%Type
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ�ʼ�����
  -- �����б�
  -- zlApptools.frmMessageEdit.LoadMessage
  -----------------------------------------------------------------------------
  Procedure Get_Zlmessage
  (
    Cursor_Out Out t_Refcur,
    Id_In      In zlMessages.ID%Type,
    ����_In    In zlMsgState.����%Type,
    �û�_In    In zlMsgState.�û�%Type
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ�ʼ�����
  -- �����б�
  -- zlApptools.frmMessageManager.FillText
  -- zlApptools.frmMessageRelate.FillText
  -----------------------------------------------------------------------------
  Procedure Get_Zlmessage
  (
    Cursor_Out Out t_Refcur,
    Id_In      In zlMessages.ID%Type
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ�ʵݵ�ַ
  -- �����б�
  -- zlApptools.frmMessageEdit.LoadMessage
  -----------------------------------------------------------------------------
  Procedure Get_Zlmsgstate
  (
    Cursor_Out Out t_Refcur,
    ��Ϣid_In  In zlMsgState.��Ϣid%Type
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�ɾ����Ϣ
  -- �����б�
  -- zlApptools.frmMessageManager.mnuEditDelete_Click
  -----------------------------------------------------------------------------
  Procedure Delete_Zlmsgstate
  (
    ɾ��_In   In zlMsgState.ɾ��%Type,
    ��Ϣid_In In zlMsgState.��Ϣid%Type,
    ����_In   In zlMsgState.����%Type,
    �û�_In   In zlMsgState.�û�%Type
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�ɾ��������Ϣ
  -- �����б�
  -- zlApptools.frmMessageManager.DeleteMessage
  -----------------------------------------------------------------------------
  Procedure Delete_Zlmessage;

  -----------------------------------------------------------------------------
  -- ���ܣ�ȡ�ʼ��б�
  -- �����б�
  -- zlApptools.frmMessageManager.FillList
  -- zlApptools.frmMessageRelate.FillList
  -----------------------------------------------------------------------------
  Procedure Get_Mail_List
  (
    Cursor_Out  Out t_Refcur,
    ��Ϣ����_In In Varchar2,
    �û�_In     In zlMsgState.�û�%Type,
    ��ʾ�Ѷ�_In In Number,
    �Ựid_In   In zlMessages.�Ựid%Type    
  );

  -----------------------------------------------------------------------------
  -- ���ܣ���ԭɾ������Ϣ
  -- �����б�
  -- zlApptools.frmMessageManager.mnuEditRestore_Click
  -----------------------------------------------------------------------------
  Procedure Restore_Zlmsgstate
  (
    ��Ϣid_In In zlMsgState.��Ϣid%Type,
    ����_In   In zlMsgState.����%Type,
    �û�_In   In zlMsgState.�û�%Type
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�����������Ϣ
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
  );

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
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�Ϊԭ�����ϴ𸴻�ת����־
  -- �����б�
  -- zlApptools.frmMessageEdit.SaveMessage
  -----------------------------------------------------------------------------
  Procedure Update_Zlmsgstate_State
  (
    ģʽ_In   In Number,
    ��Ϣid_In In zlMsgState.��Ϣid%Type,
    ����_In   In zlMsgState.����%Type,
    �û�_In   In zlMsgState.�û�%Type
  );

  -----------------------------------------------------------------------------
  -- ���ܣ�Ϊԭ�����ϴ𸴻�ת����־
  -- �����б�
  -- zlApptools.frmMessageEdit.LoadMessage
  -----------------------------------------------------------------------------
  Procedure Update_Zlmsgstate_Idtntify
  (
    ���_In   In zlMsgState.���%Type,
    ��Ϣid_In In zlMsgState.��Ϣid%Type,
    ����_In   In zlMsgState.����%Type,
    �û�_In   In zlMsgState.�û�%Type
  );

End b_Comfunc;
/

-------------------------------------------------------------------------------
-- ��������
-------------------------------------------------------------------------------
Create Or Replace Package Body zltools.b_Comfunc Is

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
      Select A.���, A.ϵͳ, C.����id As ģ��, B.�������� As �������, C.���� As ���ѱ���, A.��������, B.���ʱ��,
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

-------------------------------------------------------------------------------
-- ��Ȩ
-------------------------------------------------------------------------------
grant execute on zltools.b_ComFunc to PUBLIC
/
--- ������ͬ��ʣ�����ΪҪ������м�������ǰ׺�����á�
-------------------------------------------------------------------------------
---- �°����Ҫʹ��������,���Լ��뵽��ʽ�ű��� 2007-07-06  End
-------------------------------------------------------------------------------