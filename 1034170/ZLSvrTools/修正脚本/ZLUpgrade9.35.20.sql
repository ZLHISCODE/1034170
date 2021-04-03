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