
-------------------------------------------------------------------------------
--���˺�:
--����:9988
alter table zlreginfo modify(��Ŀ varchar2(20))
/

--NULL,��ʾ��ǰ�Ĺ̶�������,1��ʾ1�ŷ�����,2��ʾ2�ŷ�����...
alter table zlClients add(���������� number(2))
/


-------------------------------------------------------------------------------
INSERT INTO zlTools.zlRegInfo (��Ŀ,�к�,����) VALUES ('վ����',Null,Null)
/
INSERT INTO zlTools.zlRegInfo (��Ŀ,�к�,����) VALUES ('վ������',Null,Null)
/
INSERT INTO zlTools.zlRegInfo (��Ŀ,�к�,����) VALUES ('��״̬',Null,Null)
/

Create Table zlTools.zlStreamTabs(
  System_NO  Number(5),
  Table_Name Varchar2(30),
  Dml_Handle Number(1), --�Ƿ����DML�������
  Repeat_Way Number(1), --Ĭ�ϵĸ��Ʒ���1-���ر�;2-��վ�ַ���;3-˫���Ʊ�
  Fixation   Number(1)	--���Ʒ����Ƿ�̶����ɸ���
)
/
ALTER TABLE zlTools.zlStreamTabs ADD CONSTRAINT zlStreamTabs_PK PRIMARY KEY (System_NO,Table_Name) USING INDEX PCTFREE 0
/
Alter Table zlTools.zlStreamTabs Add Constraint zlStreamTabs_FK_SYSNO Foreign Key (System_NO) References zlsystems(���)
/
Create Public Synonym zlStreamTabs For zlTools.zlStreamTabs
/
GRANT SELECT ON zlTools.zlStreamTabs TO PUBLIC 
/
Begin
  For r_User In (Select ������ From Zlsystems) Loop
    Execute Immediate 'Grant Select,Insert,Update,Delete on zlTools.zlStreamTabs to ' || r_User.������ ||
                      ' With Grant Option';
    Execute Immediate 'Grant Select,Insert,Update,Delete on zlTools.zlStreamTabs to ' || r_User.������ ||
                      ' With Grant Option';
  End Loop;
End;
/

Create Or Replace Function zlTools.f_Get_Node_No Return Varchar2 As
  v_Return zlRegInfo.����%Type;
Begin
  Begin
    Select ���� Into v_Return From zlRegInfo Where ��Ŀ = 'վ����';
    If To_Number(v_Return) < 0 Or To_Number(v_Return) > 9 Then
      v_Return := Null;
    End If;
  Exception
    When Others Then
      Null;
  End;
  Return(v_Return);
End f_Get_Node_No;
/


Create Or Replace Function zlTools.f_Is_Primary_Node Return Number As
  v_Return    Number;
  v_Node_Type zlRegInfo.����%Type;
Begin
  Begin
    Select ���� Into v_Node_Type From zlRegInfo Where ��Ŀ = 'վ������';
  Exception
    When Others Then
      Null;
  End;
  If v_Node_Type = '1' Or v_Node_Type Is Null Then
    v_Return := 1;
  Else
    v_Return := 0;
  End If;
  Return(v_Return);
End f_Is_Primary_Node;
/

Create Or Replace Function Zltools.f_Get_Stream_State Return Number As
  v_Return Number;
  v_State  zlRegInfo.����%Type;
Begin
  Begin
    Select ���� Into v_State From zlRegInfo Where ��Ŀ = '��״̬';
  Exception
    When Others Then
      Null;
  End;
  If v_State = '1' Or v_State Is Null Then
    v_Return := 1;
  Else
    v_Return := 0;
  End If;
  Return(v_Return);
End f_Get_Stream_State;
/

Grant Execute on zlTools.f_Get_Node_No to Public
/
Grant Execute on zlTools.f_Is_Primary_Node to Public
/
Grant Execute on zlTools.f_Get_Stream_State to Public
/
Create Public Synonym f_Get_Node_No For zlTools.f_Get_Node_No
/
Create Public Synonym f_Is_Primary_Node For zlTools.f_Is_Primary_Node
/
Create Public Synonym f_Get_Stream_State For zlTools.f_Get_Stream_State
/



Create Or Replace Package Body b_Runmana Is
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
