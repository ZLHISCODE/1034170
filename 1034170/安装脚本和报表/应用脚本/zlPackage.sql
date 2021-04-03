--XWPACS�ӿڰ�
create or replace package b_XINWANGInterface is
  Type t_Refcur Is Ref Cursor;
-- ��    �ܣ�PACS״̬�ı���Ϣ
  Procedure PacsStatusChange
  (
    ״̬ID_In In Number,
    ҽ��ID_In In Ӱ�����¼.ҽ��ID%Type, 
    Ӱ�����_In In Ӱ�����¼.Ӱ�����%Type, 
    ����_In In Ӱ�����¼.����%Type,
    ����ʱ��  In date,
    ִ���� In  Varchar2,
    ��Ƭ��С In Varchar2
  );
-- ��    �ܣ�ȡ��ͼ�����
  Procedure PacsUnmatchImage
  (
    ҽ��ID_In In Ӱ�����¼.ҽ��ID%Type 
  );
-- ��    �ܣ���д����ͼ�Ĵ洢�豸
  Procedure PacsSetFTPDeviceNo
  (
    ҽ��ID_In In Ӱ�����¼.ҽ��ID%Type ,
    �豸��_In In Ӱ�����¼.λ��һ%Type
  );
-- ��    �ܣ�����ͼ����
  Procedure UpdateImgCount
  (
    ҽ��ID_In Ӱ�����¼.ҽ��ID%Type,
    ͼ����_In NUMBER
  );  

end b_XINWANGInterface ;
/


create or replace package body b_XINWANGInterface  is

-- ��    �ܣ�PACS״̬�ı���Ϣ
  Procedure PacsStatusChange
  (
    ״̬ID_In In Number,
    ҽ��ID_In In Ӱ�����¼.ҽ��ID%Type, 
    Ӱ�����_In In Ӱ�����¼.Ӱ�����%Type, 
    ����_In In Ӱ�����¼.����%Type,
    ����ʱ��  In date,
    ִ���� In  Varchar2,
    ��Ƭ��С In Varchar2
  ) Is
   Strsql           Varchar2(2000);
    Cursor c_Advice Is
    Select id
    From  ����ҽ����¼ 
    Where Id = ҽ��id_In Or (���id = ҽ��id_In And ������� In ('F', 'G', 'D'));
    
  Begin
       --״̬ID_In:1-ƥ��ɹ�;2-ƥ��ʧ��;3-�¼�飨�յ���һ��ͼ��;4-�յ�ÿһ��ͼ��;5-ɾ�����;6-��Ƭ��ӡ�ɹ�
       If ״̬ID_In = 1 Then 
          --ͼ��ƥ��ɹ�
      
          --��дӰ�����¼��� ���UID���������ڵȣ����ǲ���д���м���ı�,���UID��дҽ��ID
          update Ӱ�����¼ set ���UID= ҽ��ID_In ,�������� = Decode(����ʱ��, Null, Sysdate, ����ʱ��),ͼ��λ��=1 Where ҽ��ID = ҽ��ID_In ;
          
          --����ҽ��ִ��״̬
          For r_Advice In c_Advice Loop
              Update ����ҽ������
                     Set ִ��״̬ = 3, ִ�й��� = Decode(Sign(ִ�й��� - 2), 1, ִ�й���, 3)
                     Where ҽ��id = r_Advice.id;
          End Loop;
       Elsif ״̬ID_In = 2 Then 
          Strsql :='dd';
       Elsif ״̬ID_In = 3 Then 
          -- 3-�¼�飨�յ���һ��ͼ�񣩣���ʱ������
          Strsql :='dd';
       Elsif ״̬ID_In = 4 Then 
          --  4-�յ�ÿһ��ͼ�� ����ʱ������
          Strsql :='dd';
       Elsif ״̬ID_In = 5 Then 
          -- 5-ɾ�����
          -- ɾ��Ӱ�����¼���ж�Ӧ�ļ��UID���������ڵ�
          update Ӱ�����¼ set ���UID=null,λ��һ=null,λ�ö�=null,λ����=null,����ͼ��=null,��������=null
                 where ҽ��ID = ҽ��ID_IN;
       Elsif ״̬ID_In = 6 Then 
          -- 6-��Ƭ��ӡ�ɹ�
          --��¼��Ƭ��С����ӡ�˵�

          --һ��ҽ����ӡһ�Ż��߶��Ž�Ƭ�������ÿ�Ž�Ƭ����һ���̣����IDΪ��
          Insert Into ��Ƭ��ӡ��¼ (ID, ���id, ҽ��id, ��Ƭ��С, ��ӡ��, ��ӡʱ��)
                 Values (��Ƭ��ӡ��¼_Id.Nextval, Null, ҽ��ID_In, ��Ƭ��С, ִ����,Decode(����ʱ��, Null, Sysdate, ����ʱ��));
          Update Ӱ�����¼ Set �Ƿ��ӡ = 1 Where  ҽ��ID =ҽ��ID_In;
       Elsif ״̬ID_In = 7 then
         --���µ��ӽ�Ƭ״̬
          Update Ӱ�����¼ Set �Ƿ���ӽ�Ƭ = 1 Where  ҽ��ID =ҽ��ID_In;        
       End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End PacsStatusChange;
  
-- ��    �ܣ�PACSͼ��ȡ������
  Procedure PacsUnmatchImage
  (
    ҽ��ID_In In Ӱ�����¼.ҽ��ID%Type 
  ) Is
   v_ִ�й���  ����ҽ������.ִ�й���%Type; 
   v_���ͺ�  ����ҽ������.���ͺ�%Type; 
  Begin
       --����Ӱ�����¼���״̬
       update Ӱ�����¼ set ���UID= Null ,�������� = Null,ͼ��λ��=null ,λ��һ =Null Where ҽ��ID = ҽ��ID_In ;
       
       --���� Zl_Ӱ����_State �ı�����̵�״̬
       Select ִ�й���,���ͺ� Into v_ִ�й���,v_���ͺ� From ����ҽ������ Where ҽ��ID = ҽ��ID_In;
       
       --���ִ�й�����3���򽫹����޸ĳ�2
       If v_ִ�й��� = 3 Then 
          Zl_Ӱ����_State(ҽ��ID_In,v_���ͺ�,2);
       End If;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End PacsUnmatchImage;
  
-- ��    �ܣ���д����ͼ�Ĵ洢�豸
  Procedure PacsSetFTPDeviceNo
  (
    ҽ��ID_In In Ӱ�����¼.ҽ��ID%Type,
    �豸��_In In Ӱ�����¼.λ��һ%Type
  ) Is
  Begin
       --����Ӱ�����¼���״̬
       update Ӱ�����¼ set λ��һ= �豸��_In Where ҽ��ID = ҽ��ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End PacsSetFTPDeviceNo;


--RISPACS����ͼ��������Ҫ���õĴ洢����
  procedure UpdateImgCount
  (
    ҽ��ID_IN    Ӱ�����¼.ҽ��ID%Type,
    ͼ����_In    NUMBER
  ) is
  begin
      update Ӱ�����¼ set ͼ������=ͼ����_In where ҽ��ID=ҽ��ID_In;
  Exception
      When Others Then
          Zl_Errorcenter(Sqlcode, Sqlerrm);  
  end UpdateImgCount;

end b_XINWANGInterface;
/
