----------------------------------------------------------------------------------------------------------------
--���ű�֧�ִ�ZLHIS+ v10.34.170������ v10.34.170
--�������ݿռ�������ߵ�¼PLSQL��ִ�����нű�
Define n_System=100;
----------------------------------------------------------------------------------------------------------------
---------------------------------------------------------------------------------------------------------------
------------------------------------------------------------------------------
--�ṹ��������
------------------------------------------------------------------------------


------------------------------------------------------------------------------
--������������
------------------------------------------------------------------------------


-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------


-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--138081:����,2019-02-27,��ҽ��Ժ��������������
Create Or Replace Procedure Zl1_Autocptpati
(
  Patiid      In Number,
  Pageid      In Number,
  Recalcbdate In ���˱䶯��¼.�ϴμ���ʱ��%Type := Null,
  ǿ�Ƽ���_In In Number := 0
) As
  Modilast Number(1); --�Ƿ����������Զ��ƷѲ���
  Period   Varchar2(6); --��Ҫ�������С�ڼ�
Begin
  Begin
    Select �ڼ� Into Period From �ڼ�� Where Trunc(Sysdate) Between Trunc(��ʼ����) And Trunc(��ֹ����);
  Exception
    When Others Then
      Return;
  End;

  Select Zl_To_Number(zl_GetSysParameter(7)) Into Modilast From Dual;

  If Modilast = 1 Then
    Period := To_Char(Add_Months(To_Date(Period || '05', 'yyyymmdd'), -1), 'yyyymm');
  End If;

  If Recalcbdate Is Not Null Then
    Update ���˱䶯��¼
    Set �ϴμ���ʱ�� = Null
    Where ����id = Patiid And ��ҳid = Pageid And �ϴμ���ʱ�� >= Recalcbdate;
    Commit;
  End If;

  Zl1_Autocptone(Patiid, Pageid, Period, 0, ǿ�Ƽ���_In);
End Zl1_Autocptpati;
/

--138081:����,2019-02-27,��ҽ��Ժ��������������
Create Or Replace Procedure Zl1_Autocptward
(
  Wardid      In Number,
  Recalcbdate In ���˱䶯��¼.�ϴμ���ʱ��%Type := Null,
  ǿ�Ƽ���_In In Number := 0
) As
  Modilast Number(1); --�Ƿ����������Զ��ƷѲ���
  Period   Varchar2(6); --��Ҫ�������С�ڼ�

  Cursor Patitab Is
    Select Distinct ����id, ��ҳid
    From ��Ժ�����Զ�����
    Where ����id = Wardid And Trunc(��ֹ����) >= (Select Min(��ʼ����) From �ڼ�� Where �ڼ� >= Period);
Begin
  Begin
    Select �ڼ� Into Period From �ڼ�� Where Trunc(Sysdate) - 1 Between Trunc(��ʼ����) And Trunc(��ֹ����);
  Exception
    When Others Then
      Return;
  End;
  Select zl_GetSysParameter(7) Into Modilast From Dual;

  If Modilast = 1 Then
    Period := To_Char(Add_Months(To_Date(Period || '05', 'yyyymmdd'), -1), 'yyyymm');
  End If;

  For Patifld In Patitab Loop
    If Patifld.����id Is Not Null And Patifld.��ҳid Is Not Null Then
      If Recalcbdate Is Not Null Then
        Update ���˱䶯��¼
        Set �ϴμ���ʱ�� = Null
        Where ����id = Patifld.����id And ��ҳid = Patifld.��ҳid And �ϴμ���ʱ�� >= Recalcbdate;
        Commit;
      End If;
      Zl1_Autocptone(Patifld.����id, Patifld.��ҳid, Period, 1, ǿ�Ƽ���_In);
    End If;
  End Loop;
End Zl1_Autocptward;
/



------------------------------------------------------------------------------------
--ϵͳ�汾��
Update zlSystems Set �汾��='10.34.170.0001' Where ���=&n_System;
Commit;
