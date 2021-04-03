--վ��Ŀ¼
Create Table zlTools.zlNodeList(
    ��� Varchar2(1),
    ���� Varchar2(20))
    PCTFREE 5 PCTUSED 90
    Cache Storage(Buffer_Pool Keep)
/
Alter Table zlTools.zlNodeList Add Constraint zlNodeList_PK PRIMARY KEY (���) USING INDEX PCTFREE 5
/
Alter Table zlTools.zlNodeList Add Constraint zlNodeList_UQ_���� Unique (����) USING INDEX PCTFREE 5
/
Create Public Synonym zlNodeList For zlTools.zlNodeList
/
Grant Select On zlTools.zlNodeList To Public
/

Insert Into zlTools.zlNodeList(���,����)
Select '0','վ��0' From Dual Union All
Select '1','վ��1' From Dual Union All
Select '2','վ��2' From Dual Union All
Select '3','վ��3' From Dual Union All
Select '4','վ��4' From Dual Union All
Select '5','վ��5' From Dual Union All
Select '6','վ��6' From Dual Union All
Select '7','վ��7' From Dual Union All
Select '8','վ��8' From Dual Union All
Select '9','վ��9' From Dual
/

-----------------------------------------------------
-- �ͻ�����������ʽ(ף��)
-----------------------------------------------------
--���5�ֶ�
alter table zltools.zlfilesupgrade add ����ϵͳ VARCHAR2(250)
/
alter table zltools.zlfilesupgrade add ҵ�񲿼� VARCHAR2(250)
/
alter table zltools.zlfilesupgrade add ��װ·�� VARCHAR2(250)
/
alter table zltools.zlfilesupgrade add MD5 	VARCHAR2(32)
/
alter table zltools.zlfilesupgrade add �ļ�˵�� VARCHAR2(2000)
/
--���1�ֶ�
alter table zltools.zlclients add FTP������ number(2)
/
--�޸�����
ALTER TABLE zltools.zlfilesupgrade DROP CONSTRAINT ZLFILESUPGRADE_PK
/
ALTER TABLE zltools.zlfilesupgrade ADD (CONSTRAINT ZLFILESUPGRADE_PK PRIMARY KEY(�ļ���))
/
ALTER TABLE zltools.zlfilesupgrade DROP CONSTRAINT zlFilesUpgrade_UQ_�ļ���
/
ALTER TABLE zltools.zlfilesupgrade DROP CONSTRAINT zlFilesUpgrade_CK_�ļ�����
/

--35242
Insert Into zlParameters(ID,ϵͳ,ģ��,˽��,������,������,����ֵ,ȱʡֵ,����˵��)
Select zlParameters_ID.NEXTVAL, -NULL,-NULL,1,23,'����ƥ�䷽ʽ�л�',NULL,'1','�����ڴ��ڽ���Ĺ������л�����ƥ�䷽ʽ��������ʱ����������ʾ�л���ť��' From Dual
/

CREATE OR REPLACE Function Zl_To_Number
( 
  Input_In    In Varchar2, 
  Enhanced_In In Number := 0 
) Return Number Is 
  n_Index  Number; 
  v_Number Varchar2(1000); 
  n_Output Number; 
Begin 
  If Nvl(Enhanced_In, 0) = 0 Then 
    n_Output := To_Number(Input_In); 
  Else 
    n_Index := 0; 
   
    Begin 
      n_Output := To_Number(Input_In); 
    Exception 
      When Others Then 
        n_Index := 1; 
    End; 
   
    If n_Index = 1 Then 
      For n_Index In 1 .. Length(Input_In) Loop 
		if Enhanced_In=1 then 
			--ȡ���ַ�������������
			If Instr('0123456789.-', Substr(Input_In, n_Index, 1)) > 0 Then
			  v_Number := v_Number || Substr(Input_In, n_Index, 1);
			End If;
		else 
			--VAL����
			If Instr('0123456789.-+', Substr(Input_In, n_Index, 1)) > 0 Then 
			   if instr('-+',Substr(Input_In, n_Index, 1))>0 and n_Index>1 then 
				  return to_number(v_Number);             
			   else
				  v_Number := v_Number || Substr(Input_In, n_Index, 1); 
			   end if ;
			else 
			  return to_number(nvl(v_Number,0));
			End If; 
		end if ;
      End Loop; 
     
      n_Output := To_Number(v_Number); 
    End If; 
  End If; 
 
  Return n_Output; 
Exception 
  When Others Then 
    Return 0; 
End Zl_To_Number; 
/

--34943
alter table ZLTOOLS.zlclients add վ�� VARCHAR2(1)
/

-----------------------------------------------------
-- �޸� Procedure Get_Client(ף��)
-----------------------------------------------------
CREATE OR REPLACE Package Body b_Runmana Is
 
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
							 a.������, Decode(b.Terminal, Null, 0, 1) As ״̬, a.�ռ���־,a.����������,a.վ��
				From Zlclients a, (Select Distinct Terminal From V$session) b
				Where Upper(a.����վ) = Upper(b.Terminal(+))
				Order By a.Ip'; 
      Open Cur_Out For v_Sql; 
    Else 
      Open Cur_Out For 
        Select Ip, ����վ, Cpu, �ڴ�, Ӳ��, ����ϵͳ, ����, ��;, ˵��, ������־, �ռ���־, ��ֹʹ��, ������, ���������� , վ��
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
      Select ���, �ļ���, �汾��, �޸�����, �ļ�˵�� As ˵��,�ļ����� as ����,��װ·�� as ��װ·��,MD5 as MD5
      From zlFilesUpgrade Order By ���; 
  End Get_Zlfilesupgrade; 
 
  Procedure Get_Not_Regist(Cur_Out Out t_Refcur) Is 
  Begin 
    Open Cur_Out For 
      Select ��Ŀ, ���� 
      From zlRegInfo 
      Where ��Ŀ Not In ('������', '�汾��', '������Ŀ¼', '�����û�', '��������', '�ռ�Ŀ¼', '�ռ�����', 'ע����', 
             '��Ȩ֤��', '��Ȩ����', '��Ȩ�ʴ�'); 
  End Get_Not_Regist; 
  
  ----------------------------------------------------------------------------- 
  -- ���ܣ�ȡ��ע����Ŀ 
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