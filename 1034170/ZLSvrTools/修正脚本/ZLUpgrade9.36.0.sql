--PDAͬ����־��ض���
Create Table zlTools.zlPDASynch(
    ��� Number(3),
    ��ʶ Varchar2(50),
    ״̬ Number(1),
    ʱ�� TimeStamp)
    PCTFREE 10 PCTUSED 80
Storage(Freelists 4)
/
Alter Table zlTools.zlPDASynch Add Constraint zlPDASynch_PK Primary Key (��ʶ,���,ʱ��) Using Index Pctfree 0
/
Create Index zlTools.zlPDASynch_IX_ʱ�� On zlPDASynch(ʱ��) Pctfree 5
/

Create Or Replace Procedure zlTools.Zl_PDASynch_Log
(
  ���_In zlPDASynch.���%Type,
  ��ʶ_In zlPDASynch.��ʶ%Type,
  ״̬_In zlPDASynch.״̬%Type
) Is
	--״̬��1-���룬2-���£�3-ɾ��
Begin
  Update zlPDASynch Set ״̬ = ״̬_In, ʱ�� = SysTimeStamp Where ��� = ���_In And ��ʶ = ��ʶ_In;
  If Sql%RowCount = 0 Then
    Insert Into zlPDASynch (���, ��ʶ, ״̬, ʱ��) Values (���_In, ��ʶ_In, ״̬_In, SysTimeStamp);
  End If;
End Zl_PDASynch_Log;
/

Create Public Synonym zlPDASynch For zlTools.zlPDASynch;
Create Public Synonym Zl_PDASynch_Log For zlTools.Zl_PDASynch_Log;
Grant Select On zlTools.zlPDASynch To Public;
Grant Execute On zlTools.Zl_PDASynch_Log To Public;

Begin
	For r_Row IN(Select Distinct ������ From zlSystems) Loop
		Execute Immediate 'Grant Select,Insert,Update,Delete On zlTools.zlPDASynch To '||r_Row.������;
	End Loop;
End;
/

--��ǿ������ת������
Create Or Replace Function zlTools.Zl_To_Number
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
        If Instr('0123456789.-', Substr(Input_In, n_Index, 1)) > 0 Then
          v_Number := v_Number || Substr(Input_In, n_Index, 1);
        End If;
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

Create Or Replace Procedure zlTools.zl_Parameters_Update
(
  ����_In   zlParameters.������%Type,
  ����ֵ_In zlParameters.����ֵ%Type,
  ϵͳ_In   zlParameters.ϵͳ%Type,
  ģ��_In   zlParameters.ģ��%Type,
  Ȩ��_IN   Number:=1
  --���ܣ�����ϵͳ����ֵ��������û�˽�в��������û����Ե�ǰ��Ϊ׼
  --������
  --      ����_In�����봫���Nullֵ�����ַ���ʽ����Ĳ����Ż������,ע�����������Ϊ���֡�
  --      Ȩ��_IN������Ҫ����Ȩ�޿��ƵĲ�������ǰ�û��Ƿ���Ȩ������
) Is
  v_����id zlParameters.ID%Type;
  v_˽��   zlParameters.˽��%Type;
  v_����   zlParameters.����%Type;
  v_��Ȩ   zlParameters.��Ȩ%Type;
  v_������ zlUserParas.������%Type;
Begin
  --ȷ��������Ϣ
  Begin
    If Zl_To_Number(����_In) <> 0 Then
      --�Բ�����Ϊ׼����
      Select ID, ˽��, ����, ��Ȩ, Sys_Context('USERENV', 'TERMINAL')
      Into v_����id, v_˽��, v_����, v_��Ȩ, v_������
      From zlParameters
      Where Nvl(ϵͳ, 0) = Nvl(ϵͳ_In, 0) And Nvl(ģ��, 0) = Nvl(ģ��_In, 0) And ������ = Zl_To_Number(����_In);
    Else
      --�Բ�����Ϊ׼����
      Select ID, ˽��, ����, ��Ȩ, Sys_Context('USERENV', 'TERMINAL')
      Into v_����id, v_˽��, v_����, v_��Ȩ, v_������
      From zlParameters
      Where Nvl(ϵͳ, 0) = Nvl(ϵͳ_In, 0) And Nvl(ģ��, 0) = Nvl(ģ��_In, 0) And ������ = ����_In;
    End If;
  Exception
    When Others Then
      Return;
  End;
  
  --���Ȩ��
  If Nvl(Ȩ��_IN, 0) = 0 Then
    If Nvl(ϵͳ_In, 0) <> 0 And Nvl(ģ��_In, 0) = 0 And Nvl(v_˽��, 0) = 0 And Nvl(v_����, 0) = 0 Then
       Return;--����ȫ�ֲ���,�̶���ҪȨ��
    Elsif Nvl(ϵͳ_In, 0) <> 0 And Nvl(ģ��_In, 0) <> 0 And Nvl(v_˽��, 0) = 0 And Nvl(v_����, 0) = 0 Then
       Return;--����ģ�����,�̶���ҪȨ��
    Elsif Nvl(ϵͳ_In, 0) <> 0 And Nvl(ģ��_In, 0) <> 0 And Nvl(v_˽��, 0) = 0 And Nvl(v_����, 0) = 1 And Nvl(v_��Ȩ, 0) = 1 Then
       Return;--Ҫ��Ȩ���Ƶı�������ģ��
    End If;
  End If;
    
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
End zl_Parameters_Update;
/