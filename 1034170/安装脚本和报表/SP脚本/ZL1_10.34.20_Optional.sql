--[��������]1
--[�����߰汾��]10.34.0
--���ű�֧�ִ�ZLHIS+ v10.34.10 ������ v10.34.20
--����ϵͳ�����ߵ�¼PLSQL��ִ�����нű�
--�ű�ִ�к����ֹ�������������
Define n_System=100;
-------------------------------------------------------------------------------
------------------------------------------------------------------------------

--65664:������,2015-01-14,ɾ��"ҽ��ִ�мƼ�"���ж����ҩ����¼
--10.31.20�ϵ�58727�����˱�"ҽ��ִ�мƼ�",���ڴ���Ĳ����˶����ҩ����ؼ�¼��Լռ96%�����ñ�ļ�¼����������������ʱ����ܾͻ������������¼��
--ֱ��10.33.0��65664�����˸ô��󣬲��ٲ���ҩ����¼��
--��Щ��������ݣ�ռ�ô������̿ռ䣬���ҽ��������������ؽ���ͳ����Ϣ�ռ���ά�������ĺ�ʱ��Ҳ����ص�SQL��ѯ�������ܷ��գ����ԣ��ṩ�˽ű�����ɾ���ñ��е�ҩ����ؼ�¼��
--�������������ȽϺ�ʱ�����ԣ�����ݵ�ǰ�û�������ж��Ƿ��б�Ҫִ�У�����������������������������жϲ�ִ�д˽ű�����Ϊû�в��������ҩ����¼��
--1.ֱ�Ӵ�10.31.20���µİ汾��������10.33�����ϰ汾
--2.��װʱ���Ǵ�10.33�����ϰ汾��ʼ��
Create Or Replace Procedure Zl_ҽ��ִ�мƼ�_Purge Is
  n_Do     Number(5) := 1;
  v_Vermin Varchar2(20);

  Function Formatver(v_Ver Varchar2) Return Varchar2 Is
    v_Result Varchar2(20);
  Begin
    Select f_List2str(Cast(Collect(Ver) As t_Strlist), '.')
    Into v_Result
    From (Select LPad(Column_Value, 4, '0') Ver From Table(f_Str2list(v_Ver, '.')));
    Return v_Result;
  End;
Begin
  v_Vermin := '0000.0000.0000';
  For Rv In (Select ԭʼ�汾, Ŀ��汾 From zlUpGrade Where ϵͳ = 100 Order By ԭʼ�汾) Loop
    If Formatver(Rv.ԭʼ�汾) < '0010.0031.0020' And Formatver(Rv.Ŀ��汾) > '0010.0033.0000' Then
      n_Do := 0;
      Exit;
    End If;
  
    If v_Vermin < Formatver(Rv.ԭʼ�汾) Then
      v_Vermin := Formatver(Rv.ԭʼ�汾);
    End If;
  End Loop;
  If v_Vermin > '0010.0033.0000' Then
    n_Do := 0;
  End If;

  If n_Do = 1 Then
    Execute Immediate 'Alter table ҽ��ִ�мƼ� rename to ҽ��ִ�мƼ�_old';
    Execute Immediate 'Create table ҽ��ִ�мƼ� nologging tablespace zl9CisRec Initrans 20 as Select a.* From ҽ��ִ�мƼ�_old A, ����ҽ����¼ B Where a.ҽ��id = b.Id And b.������� not in(''4'',''5'',''6'',''7'')';
    Execute Immediate 'Alter table ҽ��ִ�мƼ� modify �������� default 0';
  
    Execute Immediate 'Alter table ҽ��ִ�мƼ�_old drop constraint ҽ��ִ�мƼ�_PK cascade Drop index';
    Execute Immediate 'Alter table ҽ��ִ�мƼ�_old drop constraint ҽ��ִ�мƼ�_FK_���ͺ�';
    Execute Immediate 'Alter table ҽ��ִ�мƼ�_old drop constraint ҽ��ִ�мƼ�_FK_�շ�ϸĿID';
    Execute Immediate 'Drop index ҽ��ִ�мƼ�_IX_�շ�ϸĿID';
    Execute Immediate 'Drop index ҽ��ִ�мƼ�_IX_��ת��';
  
    Execute Immediate 'Create index ҽ��ִ�мƼ�_PK On ҽ��ִ�мƼ�(ҽ��ID,���ͺ�,Ҫ��ʱ��,�շ�ϸĿID,��������) Pctfree 5 Tablespace zl9Indexcis nologging';
    Execute Immediate 'Alter table ҽ��ִ�мƼ� Add Constraint ҽ��ִ�мƼ�_PK Primary Key (ҽ��ID,���ͺ�,Ҫ��ʱ��,�շ�ϸĿID,��������) enable novalidate';
    Execute Immediate 'Alter table ҽ��ִ�мƼ� modify constraint ҽ��ִ�мƼ�_PK validate';
  
    Execute Immediate 'Alter table ҽ��ִ�мƼ� Add Constraint ҽ��ִ�мƼ�_FK_���ͺ� Foreign Key (ҽ��ID,���ͺ�) References ����ҽ������(ҽ��ID,���ͺ�) On Delete Cascade enable novalidate';
    Execute Immediate 'Alter table ҽ��ִ�мƼ� Add Constraint ҽ��ִ�мƼ�_FK_�շ�ϸĿid Foreign Key (�շ�ϸĿid) References �շ���ĿĿ¼(ID) enable novalidate';
  
    Execute Immediate 'Create index ҽ��ִ�мƼ�_IX_�շ�ϸĿid On ҽ��ִ�мƼ�(�շ�ϸĿid) Pctfree 5 Tablespace zl9Indexcis';
    Execute Immediate 'Create index ҽ��ִ�мƼ�_IX_��ת�� On ҽ��ִ�мƼ�(��ת��) Tablespace zl9Indexcis';
  
    Execute Immediate 'DROP TABLE ҽ��ִ�мƼ�_old purge';
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҽ��ִ�мƼ�_Purge;
/

-------------------------------------------------------------------------------------------------------
--����ϵͳ�������İ汾��
-------------------------------------------------------------------------------------------------------
--ϵͳ�汾��
--��ѡ�ű����ø���
--�����汾��
Commit;

