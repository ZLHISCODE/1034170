-----------------------------------------------------------------
--Ϊ��ϲ�Ʒ�汾����9.30��Ϊ9.31(VZLHIS10.20.0)
-----------------------------------------------------------------

--�����������ϵͳʱ���� by cfr
Alter Table zlBakTables Modify ϵͳ Number(5);
Alter Table zlBakSpaces Modify ϵͳ Number(5);

Alter Table zlStreamTabs Drop Constraint zlStreamTabs_FK_SYSNO;
Alter Table zlStreamTabs Add Constraint zlStreamTabs_FK_SYSNO Foreign Key (System_NO) References zlsystems(���) ON DELETE CASCADE
/

--����ϵͳ������ض���
Create Sequence zlTools.zlParameters_ID Start With 1
/
Create Table zlTools.zlParameters(
    ID NUMBER(18),
    ϵͳ NUMBER(5),
    ģ�� NUMBER(18),
    ˽�� NUMBER(1),
    ������ NUMBER(5),
    ������ VARCHAR2(100),
		����ֵ VARCHAR2(1000),
		ȱʡֵ VARCHAR2(1000),
		����˵�� VARCHAR2(255))
    PCTFREE 5 PCTUSED 90
    Cache Storage(Buffer_Pool Keep)
/
Alter Table zlTools.zlParameters Add Constraint zlParameters_PK Primary Key(ID) Using Index PCTFREE 5
/
Alter Table zlTools.zlParameters Add Constraint zlParameters_UQ_������ Unique(������,ģ��,ϵͳ,˽��) Using Index PCTFREE 5
/
Alter Table zlTools.zlParameters Add Constraint zlParameters_UQ_������ Unique(������,ģ��,ϵͳ,˽��) Using Index PCTFREE 5
/
Alter Table zlTools.zlParameters Add Constraint zlParameters_CK_˽�� Check (˽�� IN(0,1))
/

Create Table zlTools.zlUserParas(
    ����ID NUMBER(18),
    �û��� VARCHAR2(20),
		����ֵ VARCHAR2(1000))
    PCTFREE 5 PCTUSED 90
    Cache Storage(Buffer_Pool Keep)
/
Alter Table zlTools.zlUserParas Add Constraint zlUserParas_PK Primary Key(����ID,�û���) Using Index PCTFREE 5
/
Alter Table zlTools.zlUserParas Add Constraint zlUserParas_FK_����ID Foreign Key (����ID) References zlParameters(ID) On Delete Cascade
/
Create Index zlTools.zlUserParas_IX_�û��� On zlUserParas(�û���) PCTFREE 5
/

Create Or Replace Procedure zlTools.zl_Parameters_Update
(
	����_In   zlParameters.������%Type,
	����ֵ_In zlParameters.����ֵ%Type,
	ϵͳ_In   zlParameters.ϵͳ%Type,
  ģ��_In   zlParameters.ģ��%Type,
  ˽��_In   zlParameters.˽��%Type
  --���ܣ�����ϵͳ����ֵ��������û�˽�в��������û����Ե�ǰ��Ϊ׼
  --������
  --      ����_In�����봫���Nullֵ�����ַ���ʽ����Ĳ����Ż������,ע�����������Ϊ���֡�
) Is
  v_����id zlParameters.ID%Type;
Begin
  --ȷ������
  Begin
    If Zl_To_Number(����_In) <> 0 Then
      --�Բ�����Ϊ׼����
      Select ID
      Into v_����id
      From zlParameters
      Where Nvl(ϵͳ, 0) = Nvl(ϵͳ_In, 0) And Nvl(ģ��, 0) = Nvl(ģ��_In, 0) And ������ = Zl_To_Number(����_In) And
            Nvl(˽��, 0) = Nvl(˽��_In, 0);
    Else
      --�Բ�����Ϊ׼����
      Select ID
      Into v_����id
      From zlParameters
      Where Nvl(ϵͳ, 0) = Nvl(ϵͳ_In, 0) And Nvl(ģ��, 0) = Nvl(ģ��_In, 0) And ������ = ����_In And Nvl(˽��, 0) = Nvl(˽��_In, 0);
    End If;
  Exception
    When Others Then
      Return;
  End;

  --���²���ֵ
  If Nvl(˽��_In, 0) = 0 Then
    Update zlParameters Set ����ֵ = ����ֵ_In Where ID = v_����id;
  Elsif Nvl(˽��_In, 0) = 1 Then
    Update zlUserParas Set ����ֵ = ����ֵ_In Where �û��� = User And ����id = v_����id;
    If Sql%RowCount = 0 Then
      Insert Into zlUserParas (����id, �û���, ����ֵ) Values (v_����id, User, ����ֵ_In);
    End If;
  End If;
End zl_Parameters_Update;
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
      Where ϵͳ = ϵͳ���_In And ģ�� Is Null And Nvl(˽��, 0) = 0 And ������ = To_Number(v_������);
    End If;
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End zl_Parameters_Update_Batch;
/

--�����ֱ��ز�������
--˽��ȫ�֣�zlAppTools������̨�ȹ�������
Insert Into zlParameters(ID,ϵͳ,ģ��,˽��,������,������,����ֵ,ȱʡֵ,����˵��)
Select Rownum+B.ID,A.* From (
	Select ϵͳ,ģ��,˽��,������,������,����ֵ,ȱʡֵ,����˵�� From zlParameters Where ID=0 Union ALL
	Select -NULL,-NULL,1,1,'�Զ���Ϣͣ��ʱ��',NULL,'3','��¼�Զ���Ϣ������Ϣͣ��ʱ��(��)' From Dual Union ALL
	Select -NULL,-NULL,1,2,'����̨',NULL,'zlBrw','��¼ʹ���������͵ĵ���̨��zlBrw��zlWin��zlMdi' From Dual Union ALL
  Select -NULL,-NULL,1,3,'���ù���ģ��',NULL,NULL,'���ó��õĹ���ģ��' From Dual Union ALL
  Select -NULL,-NULL,1,4,'����ƥ��',NULL,'0','��������ƥ�䷽��0-˫��ƥ�䣬1-����ƥ��' From Dual Union ALL
  Select -NULL,-NULL,1,5,'���뷨',NULL,NULL,'������Ҫ�Զ����������뷨����' From Dual Union ALL
  Select -NULL,-NULL,1,6,'���뷽ʽ',NULL,'0','���ü������ɻ�����ķ�ʽ��0-ƴ����1-���' From Dual Union ALL
  Select -NULL,-NULL,1,7,'�ر�Windows',NULL,'0','�����Ƿ��˳�����ʱ�Զ��ر� Windows' From Dual Union ALL
  Select -NULL,-NULL,1,8,'�ʼ���Ϣ�������',NULL,'30','�����Զ�����ʼ���Ϣ��ʱ����(��)' From Dual Union ALL
  Select -NULL,-NULL,1,9,'��¼����ʼ���Ϣ',NULL,'0','�����Ƿ��¼ʱ����µ��ʼ���Ϣ' From Dual Union ALL
  Select -NULL,-NULL,1,10,'��ʾ�Ѷ��ʼ�',NULL,'0','�������ʼ����������Ƿ���ʾ�Ѷ��ʼ�' From Dual Union ALL
	Select -NULL,-NULL,1,11,'���ʹ��ģ��',NULL,NULL,'��¼���ʹ�õĳ���' From Dual Union ALL
	Select -NULL,-NULL,1,12,'ʹ�ø��Ի����',NULL,'1','�����Ƿ�ʹ�ø��Ի����' From Dual Union ALL   --���˺�:�޺�Ҫ���ΪĬ��ֵ
	Select -NULL,-NULL,1,13,'�����ʼ���Ϣ',NULL,'0','�����Ƿ�����ʼ���Ϣ֪ͨ' From Dual Union ALL
	Select -NULL,-NULL,1,14,'zlBrwFontSize',NULL,'0','��¼Brower��񵼺�̨�����С����С����ֱ�Ϊ��0-9��,1-11��,2-12��' From Dual Union ALL
	Select -NULL,-NULL,1,15,'zlMdiFontColor',NULL,'-1','����MDI��񵼺�̨��������ɫ' From Dual Union ALL
	Select -NULL,-NULL,1,16,'zlMdiBackPic',NULL,NULL,'����MDI��񵼺�̨�ı���ͼƬ�ļ�·��' From Dual Union ALL
	Select -NULL,-NULL,1,17,'zlMdiMenuArray',NULL,'1','����MDI��񵼺�̨�˵����з�ʽ��0-�������У�1-��������' From Dual Union ALL
	Select -NULL,-NULL,1,18,'zlWinFontColor',NULL,'-1','����Windows��񵼺�̨��������ɫ' From Dual Union ALL
	Select -NULL,-NULL,1,19,'zlWinBackPic',NULL,NULL,'����Windows��񵼺�̨�ı���ͼƬ�ļ�·��' From Dual
  ) A,(Select Nvl(Max(ID),0) AS ID From zlParameters) B
/
--������zlParameters������
Select zlParameters_ID.Nextval From zlParameters
/

--�ֱ��ز�����������
-----------------------------------------
--˽��ȫ�֣�zlAppTools������̨�ȹ�������
Declare
	v_������	zlClientScheme.������%Type;
	v_Val zlParameters.����ֵ%Type;
Begin
	--ȡ������Ϊ"zlParaUpdate"�ķ����ţ����û��������������������ת��
	Begin
		Select ������ Into v_������ From zlClientScheme Where Upper(��������)=Upper('zlParaUpdate');
	Exception
		When Others Then Return;
	End;
	
	--���������������ת��
	Begin
		Select ��ֵ Into v_Val From zlClientParaList Where ���='˽��ȫ��' And Ŀ¼='��ʷ��¼' And ����='ϵͳ' And ������=v_������;
		Select v_Val||'|'||��ֵ Into v_Val From zlClientParaList Where ���='˽��ȫ��' And Ŀ¼='��ʷ��¼' And ����='���' And ������=v_������;
		Update zlParameters Set ����ֵ=v_Val Where ϵͳ Is Null And ģ�� Is Null And ˽��=1 And ������='���ʹ��ģ��';
	Exception When Others Then Null; End;
	
	Begin

		Select ��ֵ Into v_Val From zlClientParaList Where ���='˽��ȫ��' And Ŀ¼ Is Null And ����='ʹ�ø��Ի����' And ������=v_������;
		Update zlParameters Set ����ֵ=v_Val Where ϵͳ Is Null And ģ�� Is Null And ˽��=1 And ������='ʹ�ø��Ի����';
	Exception When Others Then Null; End;

	Begin
		Select ��ֵ Into v_Val From zlClientParaList Where ���='˽��ȫ��' And Ŀ¼ Is Null And ����='��Ϣ֪ͨ' And ������=v_������;
		Update zlParameters Set ����ֵ=v_Val Where ϵͳ Is Null And ģ�� Is Null And ˽��=1 And ������='�����ʼ���Ϣ';
	Exception When Others Then Null; End;

	Begin
		Select ��ֵ Into v_Val From zlClientParaList Where ���='����ģ��' And Ŀ¼='BROWER' And ����='ZlBrowerFont' And ������=v_������;
		Update zlParameters Set ����ֵ=v_Val Where ϵͳ Is Null And ģ�� Is Null And ˽��=1 And ������='zlBrwFontSize';
	Exception When Others Then Null; End;

	Begin
		Select ��ֵ Into v_Val From zlClientParaList Where ���='����ģ��' And Ŀ¼='MDI' And ����='�˵����з�ʽ' And ������=v_������;
		Update zlParameters Set ����ֵ=v_Val Where ϵͳ Is Null And ģ�� Is Null And ˽��=1 And ������='zlMdiMenuArray';
	Exception When Others Then Null; End;

	Begin
		Select ��ֵ Into v_Val From zlClientParaList Where ���='����ģ��' And Ŀ¼='MDI' And ����='MDI����ͼƬ' And ������=v_������;
		Update zlParameters Set ����ֵ=v_Val Where ϵͳ Is Null And ģ�� Is Null And ˽��=1 And ������='zlMdiBackPic';
	Exception When Others Then Null; End;

	Begin
		Select ��ֵ Into v_Val From zlClientParaList Where ���='����ģ��' And Ŀ¼='MDI' And ����='����ɫ' And ������=v_������;
		Update zlParameters Set ����ֵ=v_Val Where ϵͳ Is Null And ģ�� Is Null And ˽��=1 And ������='zlMdiFontColor';
	Exception When Others Then Null; End;

	Begin
		Select ��ֵ Into v_Val From zlClientParaList Where ���='����ģ��' And Ŀ¼='WINDOWS' And ����='WIN����ͼƬ' And ������=v_������;
		Update zlParameters Set ����ֵ=v_Val Where ϵͳ Is Null And ģ�� Is Null And ˽��=1 And ������='zlWinBackPic';
	Exception When Others Then Null; End;

	Begin
		Select ��ֵ Into v_Val From zlClientParaList Where ���='����ģ��' And Ŀ¼='WINDOWS' And ����='����ɫ' And ������=v_������;
		Update zlParameters Set ����ֵ=v_Val Where ϵͳ Is Null And ģ�� Is Null And ˽��=1 And ������='zlWinFontColor';
	Exception When Others Then Null; End;

	Begin
		Select ��ֵ Into v_Val From zlClientParaList Where ���='˽��ȫ��' And Ŀ¼ Is Null And ����='������Ϣͣ��ʱ��' And ������=v_������;
		Update zlParameters Set ����ֵ=v_Val Where ϵͳ Is Null And ģ�� Is Null And ˽��=1 And ������='�Զ���Ϣͣ��ʱ��';
	Exception When Others Then Null; End;

	Begin
		Select ��ֵ Into v_Val From zlClientParaList Where ���='˽��ȫ��' And Ŀ¼ Is Null And ����='����̨' And ������=v_������;
		Update zlParameters Set ����ֵ=v_Val Where ϵͳ Is Null And ģ�� Is Null And ˽��=1 And ������='����̨';
	Exception When Others Then Null; End;

	Begin
		Select ��ֵ Into v_Val From zlClientParaList Where ���='˽��ȫ��' And Ŀ¼='���ù���' And ����='ϵͳ' And ������=v_������;
		Select v_Val||'|'||��ֵ Into v_Val From zlClientParaList Where ���='˽��ȫ��' And Ŀ¼='���ù���' And ����='���' And ������=v_������;
		Select v_Val||'|'||��ֵ Into v_Val From zlClientParaList Where ���='˽��ȫ��' And Ŀ¼='���ù���' And ����='ͼ��' And ������=v_������;
		Select v_Val||'|'||��ֵ Into v_Val From zlClientParaList Where ���='˽��ȫ��' And Ŀ¼='���ù���' And ����='����' And ������=v_������;
		Update zlParameters Set ����ֵ=v_Val Where ϵͳ Is Null And ģ�� Is Null And ˽��=1 And ������='���ù���ģ��';
	Exception When Others Then Null; End;

	Begin
		Select ��ֵ Into v_Val From zlClientParaList Where ���='����ģ��' And Ŀ¼='����' And ����='����ƥ��' And ������=v_������;
		Update zlParameters Set ����ֵ=v_Val Where ϵͳ Is Null And ģ�� Is Null And ˽��=1 And ������='����ƥ��';
	Exception When Others Then Null; End;

	Begin
		Select ��ֵ Into v_Val From zlClientParaList Where ���='˽��ȫ��' And Ŀ¼ Is Null And ����='���뷨' And ������=v_������;
		Update zlParameters Set ����ֵ=v_Val Where ϵͳ Is Null And ģ�� Is Null And ˽��=1 And ������='���뷨';
	Exception When Others Then Null; End;

	Begin
		Select ��ֵ Into v_Val From zlClientParaList Where ���='˽��ȫ��' And Ŀ¼ Is Null And ����='��������' And ������=v_������;
		Update zlParameters Set ����ֵ=v_Val Where ϵͳ Is Null And ģ�� Is Null And ˽��=1 And ������='���뷽ʽ';
	Exception When Others Then Null; End;

	Begin
		Select ��ֵ Into v_Val From zlClientParaList Where ���='˽��ȫ��' And Ŀ¼ Is Null And ����='�ر�Windows' And ������=v_������;
		Update zlParameters Set ����ֵ=v_Val Where ϵͳ Is Null And ģ�� Is Null And ˽��=1 And ������='�ر�Windows';
	Exception When Others Then Null; End;

	Begin
		Select ��ֵ Into v_Val From zlClientParaList Where ���='˽��ȫ��' And Ŀ¼ Is Null And ����='֪ͨ�������' And ������=v_������;
		Update zlParameters Set ����ֵ=v_Val Where ϵͳ Is Null And ģ�� Is Null And ˽��=1 And ������='�ʼ���Ϣ�������';
	Exception When Others Then Null; End;

	Begin
		Select ��ֵ Into v_Val From zlClientParaList Where ���='˽��ȫ��' And Ŀ¼ Is Null And ����='��¼ʱ���֪ͨ����Ϣ' And ������=v_������;
		Update zlParameters Set ����ֵ=v_Val Where ϵͳ Is Null And ģ�� Is Null And ˽��=1 And ������='��¼����ʼ���Ϣ';
	Exception When Others Then Null; End;

	Begin
		Select ��ֵ Into v_Val From zlClientParaList Where ���='˽��ģ��' And Ŀ¼='zl9AppTool\frmMessageManager\Menu' And ����='mnuViewShowAll״̬' And ������=v_������;
		Update zlParameters Set ����ֵ=v_Val Where ϵͳ Is Null And ģ�� Is Null And ˽��=1 And ������='��ʾ�Ѷ��ʼ�';
	Exception When Others Then Null; End;
End;
/


--����������Ȩ����
--------------------------------------------------------------------------------------------------
Create Public Synonym zlParameters_ID for zlTools.zlParameters_ID
/
Create Public Synonym zlParameters for zlTools.zlParameters
/
Create Public Synonym zlUserParas for zlTools.zlUserParas
/
Create Public Synonym zl_Parameters_Update for zlTools.zl_Parameters_Update
/
Create Public Synonym zl_Parameters_Update_Batch for zlTools.zl_Parameters_Update_Batch
/
Grant Select On zlTools.zlParameters_ID to Public 
/
Grant Select On zlTools.zlParameters to Public 
/
Grant Select On zlTools.zlUserParas to Public 
/
Grant Execute On zlTools.zl_Parameters_Update to Public 
/
Grant Execute On zlTools.zl_Parameters_Update_Batch to Public 
/
Begin
	For r_User In(Select ������ From zlSystems)
	Loop
			Execute Immediate 'Grant Select,Insert,Update,Delete on zlTools.zlParameters to '||r_User.������||' With Grant Option';
			Execute Immediate 'Grant Select,Insert,Update,Delete on zlTools.zlUserParas to '||r_User.������||' With Grant Option';
			Execute Immediate 'Grant Execute on zlTools.zl_Parameters_Update to '||r_User.������||' With Grant Option';
			Execute Immediate 'Grant Execute on zlTools.zl_Parameters_Update_Batch to '||r_User.������||' With Grant Option';
	End Loop;
End;
/

