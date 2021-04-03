----10.33.0---��9.44.0

--66229:ף��,2013-09-30,
Create Table zltools.zlClientUpdatelog (
   ����վ varchar2(50),
   �������� Date,
   ���� varchar2(300))
   PCTFREE 5 PCTUSED 90;

--66229:ף��,2013-09-30,�����������ֶ�,���ڼ�¼���һ���������
Alter Table zltools.zlClients add ������� number(1) default(0);

--65904:���Ʊ�,2013-09-17,ִ�е���ӡ����ϵͳ�Ľ�
Alter Table zltools.zlrptitems add (��ID number(18),ԴID  number(18),���¼��  number(18),���Ҽ��  number(18),�������  number(18),�������  number(18),Դ�к�  number(18));
ALTER TABLE zltools.zlRPTITems ADD CONSTRAINT zlRPTItems_FK_��ID FOREIGN KEY(��ID) REFERENCES zltools.zlRPTItems(ID) ON DELETE CASCADE;
CREATE INDEX zlRPTItems_IX_��ID ON zltools.zlRPTItems(��ID) PCTFREE 5  Compress 1;
ALTER TABLE zltools.zlRPTITems ADD CONSTRAINT zlRPTItems_FK_ԴID FOREIGN KEY(ԴID) REFERENCES zltools.Zlrptdatas(ID) ON DELETE CASCADE;
CREATE INDEX zlRPTItems_IX_ԴID ON zltools.zlRPTItems(ԴID) PCTFREE 5;

--65714:ף��,2013-09-011,��������������ļ�ģ��
Insert Into ZLTOOLS.zlSvrTools(���,�ϼ�,����,���,˵��) Values('0311','03','�����ļ�����','Q',Null);

--65203:��˶,2013-10-11,����������֮Ȩ������
Declare
  V_Sql Varchar2(100);
Begin
  For R In (Select Distinct ������ From zlSystems) Loop
  
    For R_Table In (Select Column_Value Tabname From Table(F_Str2list('Zlfilesupgrade,zlRegFunc,zlClients'))) Loop
      --����ϵͳ�����߶Խ�ɫ��Ȩʹ�� With Grant Option ���ȡ����Ȩ����ȡ�� ��ɫȨ�ޡ�
      Begin
        V_Sql := 'Revoke select,insert,update,delete on ZLTOOLS.' || R_Table.Tabname || ' from ' || R.������;
        Execute Immediate V_Sql;
      Exception
        When Others Then
          Null;
          --�����߿��ܲ�����(ϵͳͣ��)����ϵͳ������û����ЩȨ�޻��߱�����
      End;
    
      Begin
        --���¶�ϵͳ��������Ȩ
        V_Sql := 'grant select,insert,update,delete on ZLTOOLS.' || R_Table.Tabname || ' to ' || R.������ || ' With Grant Option';
        Execute Immediate V_Sql;
      Exception
        When Others Then
          Null;
          --�����߿��ܲ�����(ϵͳͣ��)���߱�����
      End;
    End Loop;
  
  End Loop;
  --����洢����Ȩ�޻���
  For R_Prog In (Select Column_Value Procname
                 From Table(F_Str2list('B_ROLEGROUPMGR,ZL_ZLROLEGRANT_BATCHDELETE,ZL_ZLROLEGRANT_BATCHINSERT'))) Loop
    Begin
      V_Sql := 'Revoke Execute on ZLTOOLS.' || R_Prog.Procname || ' From Public';
      Execute Immediate V_Sql;
    Exception
      When Others Then
        Null;
        --������Ȩ�޻���󲻴���
    End;
  End Loop;
End;
/

--65203:��˶,2013-09-03,����������
--66367:����,2013-10-12,վ�����������ƵԴ
Create Or Replace Procedure zltools.Zl_Zlclients_Set
(
  N_Mode_In       Number,
  N_Rowid_In      Varchar2 := Null,
  V_����վ_In     Zlclients.����վ%Type := Null,
  V_Ip_In         Zlclients.Ip%Type := Null,
  V_Cpu_In        Zlclients.Cpu%Type := Null,
  V_�ڴ�_In       Zlclients.�ڴ�%Type := Null,
  V_Ӳ��_In       Zlclients.Ӳ��%Type := Null,
  V_����ϵͳ_In   Zlclients.����ϵͳ%Type := Null,
  V_����_In       Zlclients.����%Type := Null,
  V_��;_In       Zlclients.��;%Type := Null,
  V_˵��_In       Zlclients.˵��%Type := Null,
  N_����������_In Zlclients.����������%Type := Null,
  N_������־_In   Zlclients.������־%Type := 0,
  N_������_In     Zlclients.������%Type := 0,
  V_վ��_In       Zlclients.վ��%Type := Null,
  N_Apply_In      Number := 0,
  V_Ipbegin_In    Varchar2 := Null,
  V_Ipend_In      Varchar2 := Null,
  N_������ƵԴ    Zlclients.������ƵԴ%Type := 0
  --���ܣ������ͻ��˻�վ�� ���߸��¿ͻ�������
  --Ӧ�ã�1�������ߣ��������޸�վ�� ���޸�ʱ��IP��ͻ������ж����������贫��N_Rowid_In��
  --      2��Ӧ��ϵͳ����¼ʱ���ݵ�ǰ��¼�Ŀͻ������ж��Ƿ�
  --                   ����վ����޸�վ�����������ʱN_Rowid_In�贫�룩
  --վ������:0-����վ�㣬1-����վ��
  --N_Apply_In,վ�����Ӧ�÷�Χ��0-��վ�㣬1�������ţ�2������վ�㣬3���̶�IP��
  --V_Ipbegin_In,V_Ipend_In:�ڹ̶�IP��Ӧ��ʱ����,������һ��IP���ϣ���ǰ�沿����ͬ
) Is
  N_Pos         Number(3);
  N_Ipbegin_Num Number;
  N_Ipend_Num   Number;
  N_Ip_Num      Number;
  N_Count       Number;

  V_Err Varchar2(500);
  Err_Custom Exception;

  Function Get_Ipnum(V_Ip_Input Varchar2) Return Number Is
    V_Ip_Num  Varchar2(20);
    N_Pos_Tmp Number;
    V_Ip_Tmp  Varchar2(20);
  Begin
    N_Pos_Tmp := Length(V_Ip_Input);
    N_Pos_Tmp := N_Pos_Tmp - Length(Replace(V_Ip_Input, '.', ''));
    If N_Pos_Tmp <> 3 Then
      Return Null;
    Else
      V_Ip_Tmp := V_Ip_Input;
      Loop
        N_Pos_Tmp := Instr(V_Ip_Tmp, '.');
        Exit When(Nvl(N_Pos_Tmp, 0) = 0);
        --��ÿһ������ת��Ϊ3λ��
        V_Ip_Num := V_Ip_Num || Trim(To_Char(Substr(V_Ip_Tmp, 1, N_Pos_Tmp - 1), '099'));
        V_Ip_Tmp := Substr(V_Ip_Tmp, N_Pos_Tmp + 1);
      End Loop;
      V_Ip_Num := V_Ip_Num || Trim(To_Char(V_Ip_Tmp, '099'));
      N_Ip_Num := To_Number(Trim(V_Ip_Num));
      Return N_Ip_Num;
    End If;
  End;
Begin
  If N_Mode_In = 0 Then

    Select Count(1) Into N_Count From zlClients Where ����վ = V_����վ_In;
    If N_Count = 0 Then
      Insert Into ZLTOOLS.zlClients
        (Ip, ����վ, Cpu, �ڴ�, Ӳ��, ����ϵͳ, ����, ��;, ˵��, ����������, ������־, ������, վ��, ������ƵԴ)
      Values
        (V_Ip_In, V_����վ_In, V_Cpu_In, V_�ڴ�_In, V_Ӳ��_In, V_����ϵͳ_In, V_����_In, V_��;_In, V_˵��_In, N_����������_In, N_������־_In,
         N_������_In, V_վ��_In, N_������ƵԴ);
    Else
      V_Err := '�Ѿ���������ͬIP��ַ����վ,��������!';
      Raise Err_Custom;
    End If;
  Else
    If N_Rowid_In Is Null Then
      Update ZLTOOLS.zlClients
      Set Cpu = V_Cpu_In, �ڴ� = V_�ڴ�_In, Ӳ�� = V_Ӳ��_In, ����ϵͳ = V_����ϵͳ_In, ���� = V_����_In, ��; = V_��;_In, ˵�� = V_˵��_In,
          ������ = N_������_In, վ�� = V_վ��_In, ������ƵԴ=N_������ƵԴ, ���������� = N_����������_In, ������־ = N_������־_In
      Where ����վ = V_����վ_In And Ip = V_Ip_In;
    Else
      Update ZLTOOLS.zlClients
      Set ����վ = V_����վ_In, Ip = V_Ip_In, Cpu = Decode(Cpu, Null, V_Cpu_In, Cpu), �ڴ� = Decode(�ڴ�, Null, V_�ڴ�_In, �ڴ�),
          Ӳ�� = Decode(Ӳ��, Null, V_Ӳ��_In, Ӳ��), ����ϵͳ = Decode(����ϵͳ, Null, V_����ϵͳ_In, ����ϵͳ), ���� = V_����_In, վ�� = V_վ��_In, ������ƵԴ=N_������ƵԴ
      Where Rowid = N_Rowid_In;
    End If;
  End If;
  --������
  If N_Apply_In = 1 Then
    Update ZLTOOLS.zlClients
    Set ������ = N_������_In, վ�� = V_վ��_In
    Where Nvl(����, 'NONE') = Nvl(V_����_In, 'NONE') And Ip <> V_Ip_In;
  Elsif N_Apply_In = 2 Then
    Update ZLTOOLS.zlClients Set ������ = N_������_In, վ�� = V_վ��_In Where Ip <> V_Ip_In;
  Elsif N_Apply_In = 3 Then
    N_Pos := Length(V_Ipbegin_In);
    N_Pos := N_Pos - Length(Replace(V_Ipbegin_In, '.', ''));
    If N_Pos <> 3 Then
      V_Err := '��ʼIP��ʽ����';
      Raise Err_Custom;
    End If;
    N_Pos := Length(V_Ipend_In);
    N_Pos := N_Pos - Length(Replace(V_Ipend_In, '.', ''));
    If N_Pos <> 3 Then
      V_Err := '����IP��ʽ����';
      Raise Err_Custom;
    End If;

    N_Ipbegin_Num := Get_Ipnum(V_Ipbegin_In);
    N_Ipend_Num   := Get_Ipnum(V_Ipend_In);
    For R_Ip In (Select ����վ, Ip From zlClients) Loop
      N_Ip_Num := Get_Ipnum(R_Ip.Ip);
      If N_Ip_Num >= N_Ipbegin_Num And N_Ip_Num <= N_Ipend_Num Then
        Update ZLTOOLS.zlClients Set ������ = N_������_In, վ�� = V_վ��_In Where ����վ = R_Ip.����վ And Ip = R_Ip.Ip;
      End If;
    End Loop;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || V_Err || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlclients_Set;
/

--65203:��˶,2013-09-03,����������
Create Or Replace Procedure ZLTOOLS.Zl_Zlclients_Delete
(
  V_����վ_In Zlclients.����վ%Type := Null,
  V_Ip_In     Zlclients.Ip%Type := Null
) Is
Begin
  If Not (V_����վ_In Is Null And V_Ip_In Is Null) Then
    If V_Ip_In Is Null Then
      Delete ZLTOOLS.zlClients Where ����վ = V_����վ_In;
    Elsif V_����վ_In Is Null Then
      Delete ZLTOOLS.zlClients Where Ip = V_Ip_In;
    Else
      Delete ZLTOOLS.zlClients Where Ip = V_Ip_In And ����վ = V_����վ_In;
    End If;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlclients_Delete;
/
--65203:��˶,2013-09-03,����������
CREATE OR REPLACE Procedure ZLTOOLS.Zl_Zlclients_Control
(  
  N_Mode_In       Number,
  V_����վ_In     Zlclients.����վ%Type := Null,
  V_Ip_In         Zlclients.Ip%Type := Null,
  N_������־_In   Zlclients.������־%Type := Null,
  N_����������_In Zlclients.����������%Type := Null,
  D_Ԥ��ʱ��_In   Zlclients.Ԥ��ʱ��%Type := Null,
  N_Ԥ�����_In   Zlclients.Ԥ�����%Type := Null,
  N_Ftp������_In  Zlclients.Ftp������%Type := Null,
  N_�ռ���־_In   Zlclients.�ռ���־%Type := Null,
  N_��ֹʹ��_In   Zlclients.��ֹʹ��%Type := Null,
  V_˵��_In       zlclients.˵��%Type :=Null
  --�Կͻ��˽��п���
  --N_Mode_In��0-���û����ÿͻ���(IP��Ϊ��Ҫ������,1-Ԥ��������,2 -������Ϣ����(IP��Ϊ��Ҫ������
  --3-ȡ��Ԥ������־,4-������վ������Ϊ����,5-�����Ѽ��������Ѽ���־��,6-��������״̬
) Is
  V_Timeset Varchar2(300);
  V_Err     Varchar2(500);
  Err_Custom Exception;
Begin
  --0-���û����ÿͻ���(IP��Ϊ��Ҫ������
  If N_Mode_In = 0 Then
    If V_����վ_In Is Not Null Then
      Update ZLTOOLS.zlClients Set ��ֹʹ�� = N_��ֹʹ��_In Where Ip = V_Ip_In;
    End If;
    --1-Ԥ��������,����Ҫ����������
  Elsif N_Mode_In = 1 Then
    Select Max(����) Into V_Timeset From zlRegInfo Where ��Ŀ = '�ͻ���Ԥ����ʱ���';
    If V_Timeset Is Not Null Then
      For R_Ip In (Select To_Date(Today || ' ' || Date_d, 'yyyy-mm-dd HH24:mi:ss') Ԥ��ʱ��, ����վ, Ip
                   From (Select ����վ, Ip, Rownum Rn_c From zlClients) A,
                        (Select To_Char(Sysdate, 'yyyy-mm-dd') Today, Column_Value Date_d, Rownum Rn_d, Count(1) Over() Sn
                          From Table(F_Str2list(V_Timeset, ','))) B
                   Where Mod(A.Rn_c, Sn) + 1 = Rn_d) Loop

        Update ZLTOOLS.zlClients Set Ԥ��ʱ�� = R_Ip.Ԥ��ʱ�� Where ����վ = R_Ip.����վ And Ip = R_Ip.Ip;
      End Loop;
    Else
      V_Err := '����δ���пͻ���Ԥ����ʱ������ã�';
      Raise Err_Custom;
    End If;
    --2 -������Ϣ����(IP��Ϊ��Ҫ������
  Elsif N_Mode_In = 2 Then
    If N_Ftp������_In Is Null Then
      Update ZLTOOLS.zlClients
      Set ������־ = N_������־_In, ���������� = N_����������_In, Ԥ��ʱ�� = D_Ԥ��ʱ��_In, Ԥ����� = N_Ԥ�����_In
      Where Ip = V_Ip_In;

    Else
      Update ZLTOOLS.zlClients
      Set ������־ = N_������־_In, Ftp������ = N_Ftp������_In, Ԥ��ʱ�� = D_Ԥ��ʱ��_In, Ԥ����� = N_Ԥ�����_In
      Where Ip = V_Ip_In;
    End If;
    --3-ȡ��Ԥ������־
  Elsif N_Mode_In = 3 Then
    Update ZLTOOLS.zlClients Set Ԥ����� = N_Ԥ�����_In;
    --4-������վ������Ϊ����
  Elsif N_Mode_In = 4 Then
    Update ZLTOOLS.zlClients Set ������־ = N_������־_In;
    --5-�����Ѽ��������Ѽ���־��
  Elsif N_Mode_In = 5 Then
    If V_����վ_In Is Null Then
      Update ZLTOOLS.zlClients Set �ռ���־ = N_�ռ���־_In;
    Else
      Update ZLTOOLS.zlClients Set �ռ���־ = N_�ռ���־_In Where ����վ = V_����վ_In;
    End If;
  Elsif N_Mode_In = 6 Then
    Update ZLTOOLS.zlClients Set �������=0 Where ����վ = V_����վ_In;
  Elsif N_Mode_In = 7 Then
    --7δ����
    Update ZLTOOLS.zlClients Set �������=1 Where ����վ = V_����վ_In;
  Elsif N_Mode_In = 8 Then
    --8������
    Update ZLTOOLS.zlClients Set �������=2 Where ����վ = V_����վ_In; 
  Elsif N_Mode_In = 9 Then
    --9�޸�˵��
    Update zltools.zlclients set ˵��=V_˵��_In where upper(����վ)=upper(V_����վ_In);
  Elsif N_Mode_In = 10 Then
    --10�޸�˵�����ռ���־
    Update zltools.zlclients set ˵��=V_˵��_In,�ռ���־=0 where upper(����վ)=upper(V_����վ_In);
  Elsif N_Mode_In = 11 Then
    --11�޸�˵����������־
    Update zltools.zlclients set ˵��=V_˵��_In,������־=0 where upper(����վ)=upper(V_����վ_In);
  Elsif N_Mode_In = 12 Then
    Update zltools.zlclients set ˵��=V_˵��_In,Ԥ�����=0 where upper(����վ)=upper(V_����վ_In);
  Elsif N_Mode_In = 13 Then
    Update zltools.zlclients set Ԥ�����=1 where upper(����վ)=upper(V_����վ_In);
  Elsif N_Mode_In = 14 Then
    Update zltools.zlclients set Ԥ��ʱ��=Null,Ԥ�����=Null where upper(����վ)=upper(V_����վ_In);
  Elsif N_Mode_In = 15 Then
    Update zltools.zlClients
         Set ������� =1
         Where upper(����վ) = (Select Upper(V_����վ_In)
         From v$Session
         Where AUDSID = UserENV('SessionID'));
  Elsif N_Mode_In = 16 Then
    Update zltools.zlClients
         Set ������� =2
         Where upper(����վ) = (Select Upper(V_����վ_In)
         From v$Session
         Where AUDSID = UserENV('SessionID'));
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || V_Err || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlclients_Control;
/
--65203:��˶,2013-09-03,����������
Create Or Replace Procedure ZLTOOLS.Zl_Zlclients_Upgrade
(
  N_Mode_In     Number,
  V_����վ_In   Zlclients.����վ%Type := Null,
  V_˵��_In     Zlclients.˵��%Type := Null,
  N_������־_In Zlclients.������־%Type := Null,
  N_�ռ���־_In Zlclients.�ռ���־%Type := Null,
  D_Ԥ��ʱ��_In Zlclients.Ԥ��ʱ��%Type := Null,
  N_Ԥ�����_In Zlclients.Ԥ�����%Type := Null
  --��Ҫ�ǿͻ����Զ�����ʹ�á�
  --N_Mode_In:0-��������˵��,�����������ռ����,1-����վ���Ԥ�������״̬
  --2-�ͻ���Ϊ��ʱ����
) Is
  V_Err Varchar2(500);
  Err_Custom Exception;
Begin
  --0-��������˵��,�����������ռ����
  If N_Mode_In = 0 Then
    If N_�ռ���־_In Is Null And N_������־_In Is Null Then
      Update ZLTOOLS.zlClients Set ˵�� = V_˵��_In Where Upper(����վ) = Upper(V_����վ_In);
    Elsif N_�ռ���־_In Is Null Then
      Update ZLTOOLS.zlClients Set ˵�� = V_˵��_In, ������־ = N_������־_In Where Upper(����վ) = Upper(V_����վ_In);
    Elsif N_������־_In Is Null Then
      Update ZLTOOLS.zlClients Set ˵�� = V_˵��_In, �ռ���־ = N_�ռ���־_In Where Upper(����վ) = Upper(V_����վ_In);
    End If;
    --1-����վ���Ԥ�������״̬
  Elsif N_Mode_In = 1 Then
    If V_˵��_In Is Null Then
      Update ZLTOOLS.zlClients Set Ԥ����� = N_Ԥ�����_In, ˵�� = V_˵��_In Where Upper(����վ) = Upper(V_����վ_In);
    Else
      Update ZLTOOLS.zlClients Set Ԥ����� = N_Ԥ�����_In Where Upper(����վ) = Upper(V_����վ_In);
    End If;
  --2-�ͻ���Ϊ��ʱ����
  Elsif N_Mode_In = 2 Then
    Update ZLTOOLS.zlClients Set Ԥ��ʱ�� = D_Ԥ��ʱ��_In, Ԥ����� = N_Ԥ�����_In Where Upper(����վ) = Upper(V_����վ_In);
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || V_Err || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Zlclients_Upgrade;
/

