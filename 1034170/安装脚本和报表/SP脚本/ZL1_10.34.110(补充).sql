--[��������]1
--[�����߰汾��]10.34.110
--���ű�֧�ִ�ZLHIS+ v10.34.100 ������ v10.34.110
--����ϵͳ�����ߵ�¼PLSQL��ִ�����нű�
--�ű�ִ�к����ֹ�������������
Define n_System=100;
-------------------------------------------------------------------------------
--�ṹ��������
-------------------------------------------------------------------------------
--109686:������,2017-05-31,������Ψһ��ȱʧ����
Alter Table ��������Ŀ Add Constraint ��������Ŀ_PK Primary Key (��Ŀ���) Using Index Tablespace zl9Indexcis;

Alter Table �������ÿ��� Add Constraint �������ÿ���_PK Primary Key (��Ŀ���,����ID) Using Index Tablespace zl9Indexcis;

--109412:�ŵ���,2017-05-23,��Ӿ�����ر������Լ��
alter table ��Һ������ҩƷ add constraint ��Һ������ҩƷ_PK primary key(ҩƷid) Using Index Tablespace zl9Indexhis;

alter table ��Һ��ҩ���� add constraint ��Һ��ҩ����_PK primary key(����) Using Index Tablespace zl9Indexhis;

alter table ��ҺҩƷ���ȼ� add constraint ��ҺҩƷ���ȼ�_PK primary key(����id,��ҩ����,Ƶ��) Using Index Tablespace zl9Indexhis;

alter table ��Һ���ȴ�ӡҩƷ add constraint ��Һ���ȴ�ӡҩƷ_PK primary key(ҩƷid) Using Index Tablespace zl9Indexhis;

alter table �����շѷ��� drop constraint �����շѷ���_UQ_��ҩ���� cascade drop index; 

alter table �����շѷ��� add constraint �����շѷ���_PK primary key(��ҩ����,��Ŀid) Using Index Tablespace zl9Indexhis;

--109164:���ϴ�,2017-05-23,���Ӳ������߼�¼���������
declare 
  n_count Number;   
begin   
  --���Ӳ������߼�¼������ǰҪ������ܴ��ڵ��ظ���¼
  n_count := 0;
  For C_���� in (Select ����ID,����ʱ�� From �������߼�¼ group by ����ID,����ʱ�� having count(1) > 1) Loop
      Update �������߼�¼ set ����ʱ�� = ����ʱ�� + RowNum * 1/24/60/60  Where ����ID = C_����.����ID And ����ʱ�� = C_����.����ʱ��;
      n_count := 1;
  end Loop;

  If n_count = 1 Then 
     Commit;
  end if;
end;
/

Alter Table �������߼�¼ add Constraint �������߼�¼_PK Primary Key (����ID,����ʱ��) Using Index Tablespace zl9indexhis;

--103974:���ϴ�,2017-05-19,�Զ����˴Ӳ�����ҳ��ȡ���˻�����Ϣ
Create Or Replace View ��Ժ�����Զ����� As
Select p.����id, p.��ҳid, Nvl(A.����,I.����) as ����, Nvl(A.�Ա�,I.�Ա�) as �Ա�, Nvl(A.����,I.����) as ����, i.סԺ��, a.�ѱ�, p.����id, p.����id, p.����, p.���Ӵ�λ, p.�շ�ϸĿid, p.������Ŀid, 1 As ��־,
       p.�ּ� As ��׼����, p.��ʼ����, p.��ֹ����, p.��ֹ���� - p.��ʼ���� As ����, p.����, p.����ҽʦ, p.���λ�ʿ, p.����Ա���, p.����Ա����
From ������Ϣ I, ������ҳ A,
     (Select b.����id, b.��ҳid, b.����id, b.����id, b.����, b.���Ӵ�λ, p.�շ�ϸĿid, p.������Ŀid, p.�ּ�, b.����ҽʦ, b.���λ�ʿ, b.����Ա���, b.����Ա����,
              Zl_Date_Half(Greatest(Least(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��), Nvl(b.��ֹʱ��, Greatest(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��))),
                                           Greatest(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��))), p.ִ������, Nvl(a.��������, Add_Months(Sysdate, -2)))) As ��ʼ����,
              Zl_Date_Half(Least(Nvl(b.��ֹʱ��, Greatest(b.��ʼʱ��, Sysdate)), Nvl(p.��ֹ����, Sysdate + 30) + 1)) As ��ֹ����, b.����
       From �Զ��Ƽ���Ŀ A,
            (Select a.����id, a.��ҳid, a.��ʼʱ��, a.���Ӵ�λ, a.����id, a.����id, a.����, a.��λ�ȼ�id, 1 As ����, a.���λ�ʿ, a.����ҽʦ, a.��ֹʱ��,
                     a.����Ա���, a.����Ա����, a.�ϴμ���ʱ��
              From ���˱䶯��¼ A, ������Ϣ B
              Where a.��ʼԭ�� <> 10 And a.����id = b.����id And a.��ҳid = b.��ҳid And b.��Ժ = 1
              Union All
              Select b.����id, b.��ҳid, ��ʼʱ��, ���Ӵ�λ, b.����id, b.����id, ����, i.����id As ��λ�ȼ�id, i.�������� As ����, ���λ�ʿ, ����ҽʦ, ��ֹʱ��, ����Ա���,
                     ����Ա����, �ϴμ���ʱ��
              From ���˱䶯��¼ B, �շѴ�����Ŀ I, ������Ϣ C
              Where b.����id = c.����id And b.��ҳid = c.��ҳid And c.��Ժ = 1 And b.��λ�ȼ�id = i.����id And b.��ʼԭ�� <> 10 And i.���д��� > 0) B,
            �շѼ�Ŀ P
       Where a.����id = b.����id And Zl_Date_Half(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��)) <> Zl_Date_Half(Nvl(b.��ֹʱ��, Sysdate)) And p.�ּ� <> 0 And
             a.�����־ = 1 And b.��λ�ȼ�id = p.�շ�ϸĿid And Zl_Date_Half(Nvl(b.��ֹʱ��, Sysdate)) >= Zl_Date_Half(p.ִ������) And
             Zl_Date_Half(b.��ʼʱ��) <= Zl_Date_Half(Nvl(p.��ֹ����, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.��ֹʱ��, Sysdate), Nvl(p.��ֹ����, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.��������, Add_Months(Sysdate, -2)))
       Union All
       Select b.����id, b.��ҳid, b.����id, b.����id, b.����, b.���Ӵ�λ, p.�շ�ϸĿid, p.������Ŀid, p.�ּ�, b.����ҽʦ, b.���λ�ʿ, b.����Ա���, b.����Ա����,
              Zl_Date_Half(Greatest(Least(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��), Nvl(b.��ֹʱ��, Greatest(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��))),
                                           Greatest(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��))), p.ִ������, Nvl(a.��������, Add_Months(Sysdate, -2)))) As ��ʼ����,
              Zl_Date_Half(Least(Nvl(b.��ֹʱ��, Greatest(b.��ʼʱ��, Sysdate)), Nvl(p.��ֹ����, Sysdate + 30) + 1)) As ��ֹ����, b.����
       From �Զ��Ƽ���Ŀ A,
            (Select a.����id, a.��ҳid, ��ʼʱ��, ���Ӵ�λ, a.����id, a.����id, ����, ����ȼ�id, 1 As ����, ���λ�ʿ, ����ҽʦ, ��ֹʱ��, ����Ա���, ����Ա����, �ϴμ���ʱ��
              From ���˱䶯��¼ A, ������Ϣ B
              Where ��ʼԭ�� <> 10 And a.����id = b.����id And a.��ҳid = b.��ҳid And b.��Ժ = 1
              Union All
              Select b.����id, b.��ҳid, ��ʼʱ��, ���Ӵ�λ, b.����id, b.����id, ����, i.����id As ����ȼ�id, i.�������� As ����, ���λ�ʿ, ����ҽʦ, ��ֹʱ��, ����Ա���,
                     ����Ա����, �ϴμ���ʱ��
              From ���˱䶯��¼ B, �շѴ�����Ŀ I, ������Ϣ C
              Where b.����ȼ�id = i.����id And b.����id = c.����id And b.��ҳid = c.��ҳid And c.��Ժ = 1 And b.��ʼԭ�� <> 10 And i.���д��� > 0) B,
            �շѼ�Ŀ P, �շ���ĿĿ¼ C
       Where a.����id = b.����id And b.���Ӵ�λ <> 1 And
             Zl_Date_Half(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��)) <> Zl_Date_Half(Nvl(b.��ֹʱ��, Sysdate)) And p.�ּ� <> 0 And a.�����־ = 2 And
             b.����ȼ�id = p.�շ�ϸĿid And b.����ȼ�id = c.Id And Nvl(c.���㷽ʽ, 0) <> 1 And
             Zl_Date_Half(Nvl(b.��ֹʱ��, Sysdate)) >= Zl_Date_Half(p.ִ������) And
             Zl_Date_Half(b.��ʼʱ��) <= Zl_Date_Half(Nvl(p.��ֹ����, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.��ֹʱ��, Sysdate), Nvl(p.��ֹ����, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.��������, Add_Months(Sysdate, -2)))
       Union All
       Select b.����id, b.��ҳid, b.����id, b.����id, b.����, b.���Ӵ�λ, p.�շ�ϸĿid, p.������Ŀid, p.�ּ�, b.����ҽʦ, b.���λ�ʿ, b.����Ա���, b.����Ա����,
              Zl_Date_Half(Greatest(Least(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��), Nvl(b.��ֹʱ��, Greatest(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��))),
                                           Greatest(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��))), p.ִ������, Nvl(a.��������, Add_Months(Sysdate, -2)))) As ��ʼ����,
              Zl_Date_Half(Least(Nvl(b.��ֹʱ��, Greatest(b.��ʼʱ��, Sysdate)), Nvl(p.��ֹ����, Sysdate + 30) + 1)) As ��ֹ����, a.����
       From (Select ����id, �����־, �շ�ϸĿid, 1 As ����, ��������
              From �Զ��Ƽ���Ŀ
              Union All
              Select ����id, �����־, ����id, i.�������� As ����, ��������
              From �Զ��Ƽ���Ŀ A, �շѴ�����Ŀ I
              Where a.�շ�ϸĿid = i.����id And i.���д��� > 0) A, ���˱䶯��¼ B, �շѼ�Ŀ P, ������Ϣ C
       Where a.����id = b.����id And b.����id = c.����id And b.��ҳid = c.��ҳid And c.��Ժ = 1 And b.���Ӵ�λ <> 1 And b.��ʼԭ�� <> 10 And
             Zl_Date_Half(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��)) <> Zl_Date_Half(Nvl(b.��ֹʱ��, Sysdate)) And p.�ּ� <> 0 And
             a.�շ�ϸĿid = p.�շ�ϸĿid And (a.�����־ = 6 And b.��λ�ȼ�id Is Not Null Or a.�����־ = 7) And
             Zl_Date_Half(Nvl(b.��ֹʱ��, Sysdate)) >= Zl_Date_Half(p.ִ������) And
             Zl_Date_Half(b.��ʼʱ��) <= Zl_Date_Half(Nvl(p.��ֹ����, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.��ֹʱ��, Sysdate), Nvl(p.��ֹ����, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.��������, Add_Months(Sysdate, -2)))) P
Where i.����id = p.����id And a.����id = p.����id And a.��ҳid = p.��ҳid;


Create Or Replace View ��Ժ�����Զ����� As
Select p.����id, p.��ҳid, Nvl(A.����,I.����) as ����, Nvl(A.�Ա�,I.�Ա�) as �Ա�, Nvl(A.����,I.����) as ����, i.סԺ��, a.�ѱ�, p.����id, p.����id, p.����, p.���Ӵ�λ, p.�շ�ϸĿid, p.������Ŀid, 1 As ��־,
       p.�ּ� As ��׼����, p.��ʼ����, p.��ֹ����, p.��ֹ���� - p.��ʼ���� As ����, p.����, p.����ҽʦ, p.���λ�ʿ, p.����Ա���, p.����Ա����
From ������Ϣ I, ������ҳ A,
     (Select b.����id, b.��ҳid, b.����id, b.����id, b.����, b.���Ӵ�λ, p.�շ�ϸĿid, p.������Ŀid, p.�ּ�, b.����ҽʦ, b.���λ�ʿ, b.����Ա���, b.����Ա����,
              Zl_Date_Half(Greatest(Least(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��), Nvl(b.��ֹʱ��, Greatest(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��))),
                                           Greatest(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��))), p.ִ������, Nvl(a.��������, Add_Months(Sysdate, -2)))) As ��ʼ����,
              Zl_Date_Half(Least(Nvl(b.��ֹʱ��, Greatest(b.��ʼʱ��, Sysdate)), Nvl(p.��ֹ����, Sysdate + 30) + 1)) As ��ֹ����, b.����
       From �Զ��Ƽ���Ŀ A,
            (Select ����id, ��ҳid, ��ʼʱ��, ���Ӵ�λ, ����id, ����id, ����, ��λ�ȼ�id, 1 As ����, ���λ�ʿ, ����ҽʦ, ��ֹʱ��, ����Ա���, ����Ա����, �ϴμ���ʱ��
              From ���˱䶯��¼ A
              Where ��ʼԭ�� <> 10
              Union All
              Select ����id, ��ҳid, ��ʼʱ��, ���Ӵ�λ, ����id, ����id, ����, i.����id As ��λ�ȼ�id, i.�������� As ����, ���λ�ʿ, ����ҽʦ, ��ֹʱ��, ����Ա���, ����Ա����,
                     �ϴμ���ʱ��
              From ���˱䶯��¼ B, �շѴ�����Ŀ I
              Where b.��λ�ȼ�id = i.����id And b.��ʼԭ�� <> 10 And i.���д��� > 0) B, �շѼ�Ŀ P
       Where a.����id = b.����id And Zl_Date_Half(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��)) <> Zl_Date_Half(Nvl(b.��ֹʱ��, Sysdate)) And p.�ּ� <> 0 And
             a.�����־ = 1 And b.��λ�ȼ�id = p.�շ�ϸĿid And Zl_Date_Half(Nvl(b.��ֹʱ��, Sysdate)) >= Zl_Date_Half(p.ִ������) And
             Zl_Date_Half(b.��ʼʱ��) <= Zl_Date_Half(Nvl(p.��ֹ����, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.��ֹʱ��, Sysdate), Nvl(p.��ֹ����, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.��������, Add_Months(Sysdate, -2)))
       Union All
       Select b.����id, b.��ҳid, b.����id, b.����id, b.����, b.���Ӵ�λ, p.�շ�ϸĿid, p.������Ŀid, p.�ּ�, b.����ҽʦ, b.���λ�ʿ, b.����Ա���, b.����Ա����,
              Zl_Date_Half(Greatest(Least(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��), Nvl(b.��ֹʱ��, Greatest(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��))),
                                           Greatest(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��))), p.ִ������, Nvl(a.��������, Add_Months(Sysdate, -2)))) As ��ʼ����,
              Zl_Date_Half(Least(Nvl(b.��ֹʱ��, Greatest(b.��ʼʱ��, Sysdate)), Nvl(p.��ֹ����, Sysdate + 30) + 1)) As ��ֹ����, b.����
       From �Զ��Ƽ���Ŀ A,
            (Select ����id, ��ҳid, ��ʼʱ��, ���Ӵ�λ, ����id, ����id, ����, ����ȼ�id, 1 As ����, ���λ�ʿ, ����ҽʦ, ��ֹʱ��, ����Ա���, ����Ա����, �ϴμ���ʱ��
              From ���˱䶯��¼
              Where ��ʼԭ�� <> 10
              Union All
              Select ����id, ��ҳid, ��ʼʱ��, ���Ӵ�λ, ����id, ����id, ����, i.����id As ����ȼ�id, i.�������� As ����, ���λ�ʿ, ����ҽʦ, ��ֹʱ��, ����Ա���, ����Ա����,
                     �ϴμ���ʱ��
              From ���˱䶯��¼ B, �շѴ�����Ŀ I
              Where b.����ȼ�id = i.����id And b.��ʼԭ�� <> 10 And i.���д��� > 0) B, �շѼ�Ŀ P, �շ���ĿĿ¼ C
       Where a.����id = b.����id And b.���Ӵ�λ <> 1 And
             Zl_Date_Half(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��)) <> Zl_Date_Half(Nvl(b.��ֹʱ��, Sysdate)) And p.�ּ� <> 0 And a.�����־ = 2 And
             b.����ȼ�id = p.�շ�ϸĿid And b.����ȼ�id = c.Id And Nvl(c.���㷽ʽ, 0) <> 1 And
             Zl_Date_Half(Nvl(b.��ֹʱ��, Sysdate)) >= Zl_Date_Half(p.ִ������) And
             Zl_Date_Half(b.��ʼʱ��) <= Zl_Date_Half(Nvl(p.��ֹ����, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.��ֹʱ��, Sysdate), Nvl(p.��ֹ����, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.��������, Add_Months(Sysdate, -2)))
       Union All
       Select b.����id, b.��ҳid, b.����id, b.����id, b.����, b.���Ӵ�λ, p.�շ�ϸĿid, p.������Ŀid, p.�ּ�, b.����ҽʦ, b.���λ�ʿ, b.����Ա���, b.����Ա����,
              Zl_Date_Half(Greatest(Least(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��), Nvl(b.��ֹʱ��, Greatest(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��))),
                                           Greatest(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��))), p.ִ������, Nvl(a.��������, Add_Months(Sysdate, -2)))) As ��ʼ����,
              Zl_Date_Half(Least(Nvl(b.��ֹʱ��, Greatest(b.��ʼʱ��, Sysdate)), Nvl(p.��ֹ����, Sysdate + 30) + 1)) As ��ֹ����, a.����
       From (Select ����id, �����־, �շ�ϸĿid, 1 As ����, ��������
              From �Զ��Ƽ���Ŀ
              Union All
              Select ����id, �����־, ����id, i.�������� As ����, ��������
              From �Զ��Ƽ���Ŀ A, �շѴ�����Ŀ I
              Where a.�շ�ϸĿid = i.����id And i.���д��� > 0) A, ���˱䶯��¼ B, �շѼ�Ŀ P
       Where a.����id = b.����id And b.���Ӵ�λ <> 1 And b.��ʼԭ�� <> 10 And
             Zl_Date_Half(Nvl(b.�ϴμ���ʱ��, b.��ʼʱ��)) <> Zl_Date_Half(Nvl(b.��ֹʱ��, Sysdate)) And p.�ּ� <> 0 And
             a.�շ�ϸĿid = p.�շ�ϸĿid And (a.�����־ = 6 And b.��λ�ȼ�id Is Not Null Or a.�����־ = 7) And
             Zl_Date_Half(Nvl(b.��ֹʱ��, Sysdate)) >= Zl_Date_Half(p.ִ������) And
             Zl_Date_Half(b.��ʼʱ��) <= Zl_Date_Half(Nvl(p.��ֹ����, Sysdate) + 1) And
             Zl_Date_Half(Least(Nvl(b.��ֹʱ��, Sysdate), Nvl(p.��ֹ����, Sysdate + 30) + 1)) >=
             Zl_Date_Half(Nvl(a.��������, Add_Months(Sysdate, -2)))) P
Where i.����id = p.����id And a.����id = p.����id And a.��ҳid = p.��ҳid;

--109168:л��,2017-05-16,ҽ�����˵�����ҽ�����˹����������˳��һ�¡�
Alter Table ҽ�����˵��� Drop Constraint ҽ�����˵���_PK Cascade Drop Index;

Alter Table ҽ�����˵��� Add Constraint ҽ�����˵���_PK Primary Key (ҽ����,����,����) Using Index Tablespace zl9Indexhis;

--108762:Ƚ����,2017-05-15,�ٴ����ﰲ��ҽ������ǰ��ʾְ�Ʊ�ʶ��
Alter Table רҵ����ְ�� Add ��ʶ�� Varchar2(5);

--109137:������,2017-05-15,����������ȱʧ
Alter Table ҽ������ԭ�� Add Constraint ҽ������ԭ��_PK Primary Key (����) Using Index Tablespace zl9Indexhis;

Alter Table ҽ������ԭ�� Add Constraint ҽ������ԭ��_UQ_��Ա Unique (��Ա,����,����) Using Index Tablespace zl9Indexhis;

--105165:�ŵ���,2017-04-20,������ҩ������ӸýкŴ��ڵ����д���
alter table ��ҩ���� add �кŴ��� varchar2(10);

--107559:Ƚ����,2017-04-17,������ֹͣ�ﰲ�Ź���
Alter Table �ٴ�����ͣ���¼ Add ʧЧʱ�� Date;

--105791:Ƚ����,2017-04-05,�������ձ��ֶ�������
Alter Table �������ձ� Rename Column ����ԤԼ To ����ԤԼ����;
Alter Table �������ձ� Rename Column ����Һ� To ����Һ�����;

--109164:���ϴ�,2017-05-23,���Ӳ������߼�¼���������
Alter Table �������߼�¼ add Constraint �������߼�¼_FK_����ID Foreign Key (����ID) References ������Ϣ(����ID);

--109421:��ҵ��,2017-05-23,ҩƷ�ӳɷ�������������
Alter Table ҩƷ�ӳɷ��� Add Constraint ҩƷ�ӳɷ���_PK Primary Key (���) Using Index Tablespace zl9Indexhis;



-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--109289:Ƚ����,2017-05-23,ʹ�á�ȫ��������ſ��ơ�����ʱ������������ŵ�û�����÷�ʱ�εİ��ţ�û�����ɶ�Ӧ��ʱ���������
--�������ݣ��ٴ����ﰲ�ţ�����������̫��
Begin
  --1.��������
  --����ʱ�ε���ſ��ƺ����������,��ʼʱ�䡢��ֹʱ����дʱ��εĿ�ʼʱ��ͽ���ʱ��
  For c_���� In (Select a.Id, b.����, Nvl(c.վ��, '-') As վ��
               From �ٴ����ﰲ�� A, �ٴ������Դ B, ���ű� C, �ٴ������ D
               Where a.��Դid = b.Id And b.����id = c.Id And a.����id = d.Id And Nvl(d.�Ű෽ʽ, 0) In (0, 3)) Loop
  
    For c_��¼ In (With c_ʱ��� As
                    (Select ʱ���, ��ʼʱ��, ��ֹʱ��
                    From (Select ʱ���,
                                  To_Date('3000-01-01' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                                  To_Date('3000-01-01' || To_Char(��ֹʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ��ֹʱ��,
                                  Row_Number() Over(Partition By ʱ��� Order By ʱ���, վ�� Asc, ���� Asc) As ���
                           From ʱ���
                           Where Nvl(վ��, c_����.վ��) = c_����.վ�� And Nvl(����, c_����.����) = c_����.����)
                    Where ��� = 1)
                   Select a.Id, a.�޺���,
                          To_Date(To_Char(Sysdate, 'yyyy-mm-dd ') || To_Char(b.��ʼʱ��, 'hh24:mi:ss'),
                                   'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                          To_Date(To_Char(Sysdate, 'yyyy-mm-dd ') || To_Char(b.��ֹʱ��, 'hh24:mi:ss'),
                                  'yyyy-mm-dd hh24:mi:ss') + Case
                            When b.��ֹʱ�� <= b.��ʼʱ�� Then
                             1
                            Else
                             0
                          End As ��ֹʱ��
                   From �ٴ��������� A, c_ʱ��� B
                   Where a.�ϰ�ʱ�� = b.ʱ��� And a.����id = c_����.Id And Nvl(a.�޺���, 0) <> 0 And Nvl(a.�Ƿ���ſ���, 0) = 1 And
                         Nvl(a.�Ƿ��ʱ��, 0) = 0 And Not Exists (Select 1 From �ٴ�����ʱ�� Where ����id = a.Id)) Loop
    
      For I In 1 .. c_��¼.�޺��� Loop
        Insert Into �ٴ�����ʱ��
          (����id, ���, ��ʼʱ��, ��ֹʱ��, ��������, �Ƿ�ԤԼ)
        Values
          (c_��¼.Id, I, c_��¼.��ʼʱ��, c_��¼.��ֹʱ��, 1, 1);
      End Loop;
    End Loop;
  End Loop;

  --2.��¼����
  --����ʱ�ε���ſ��ƺ����������,��ʼʱ�䡢��ֹʱ����дʱ��εĿ�ʼʱ��ͽ���ʱ��
  For c_��¼ In (Select a.Id, a.�޺���, a.��ʼʱ��, a.��ֹʱ��
               From �ٴ������¼ A
               Where a.�������� > Trunc(Sysdate) And Nvl(a.�ѹ���, 0) = 0 And Nvl(a.��Լ��, 0) = 0 And Nvl(a.�޺���, 0) <> 0 And
                     Nvl(a.�Ƿ���ſ���, 0) = 1 And Nvl(a.�Ƿ��ʱ��, 0) = 0 And Not Exists
                (Select 1 From �ٴ�������ſ��� Where ��¼id = a.Id)) Loop
  
    For I In 1 .. c_��¼.�޺��� Loop
      Insert Into �ٴ�������ſ���
        (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ)
      Values
        (c_��¼.Id, I, c_��¼.��ʼʱ��, c_��¼.��ֹʱ��, 1, 1);
    End Loop;
  End Loop;
End;
/

--106248:��͢��,2017-5-15,������ϼ��ICD������
Update zlParameters
Set ����˵��='��Ҫ��Ժ��ϱ��뿪ͷΪC00��D48ʱ,������ϣ�1-������д��2-��ʾ�Ƿ���д��0-����顣'
Where ������ = '������ϼ��' And ģ�� = 1261 And ϵͳ = &n_System;
Update zlParameters
Set ����˵��='��Ҫ��Ժ��ϱ��뿪ͷΪC00��D48ʱ����Ҫ��Ժ��ϵ�ICD���룺1-������д��2-��ʾ�Ƿ���д��0-����顣'
Where ������ = 'ICD������' And ģ�� = 1261 And ϵͳ = &n_System;

--108423:��С��,2017-05-11,����걾�Ǽǣ������걾��ʾ
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1212, 0, 1, 0, 0, 6, '�����걾��ʾ', '0', '0', '����걾�Ǽǣ��ǼǱ걾Ϊ�����걾ʱ�Ƿ���ʾ'
  From Dual;

--106148:������,2017-04-17,�°���ǰ�Һ���ɫ����
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1111, 0, 0, 0, 0, 72, '��ǰ�Һ���ɫ', Null, '0',
         '�°�Һ�ʱ����ǰ�ҺŰ��ŵ�������ɫ��ʾ��'
  From Dual
  Where Not Exists (Select 1 From zlParameters Where ϵͳ = &n_System And ģ�� = 1111 And ������ = '��ǰ�Һ���ɫ');

--89759:���ϴ�,2017-04-14,���ѿ�ˢ���Ƿ�λ�������
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, -Null, -Null, -Null, -Null, -Null, 276, '���ѿ�ˢ�������붨λ�������', '1', '1',
         '������ò��������ѿ�ˢ�����ѵ�ʱ��û�п����������£����Ҳ��λ�������ı���'
  From Dual
  Where Not Exists (Select 1 From zlParameters Where ������ = 276 And Nvl(ģ��, 0) = 0 And Nvl(ϵͳ, 0) = &n_System);

--105443:������,2017-04-06,ԤԼ��Чʱ�������
Update zlParameters
Set ����˵�� = '��ʾԤԼ��ʵ��ԤԼ����ʱ�����Ч��Χ,�Է���Ϊ��λ,0��ʾ������,>0��ʾ��ǰ���յ����Ʒ�����,<0��ʾ�Ӻ���յ����Ʒ�����'
Where ϵͳ = &n_System And ģ�� = 1111 And ������ = 'ԤԼ��Чʱ��';

--104983:Ƚ����,2017-04-05,��첡�˰����ݷֱ��ӡ
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1121, 0, 0, 0, 0, 113, '��첡�˷ֵ��ݴ�ӡ', Null, '0',
         '�ڲ����շѹ����У����Ʊ���ǰ�������ʵ�ʴ�ӡ����Ʊ�š����������ˡ������շ�ÿ�ŵ��ݷֱ��ӡ��ʱ����첡���Ƿ�ÿ���շѵ����зֱ��ӡ��Ʊ��1-��첡�˰�ÿ�ŵ��ݷֱ��ӡ��0��NULL-��첡�˲���ÿ�ŵ��ݷֱ��ӡ'
  From Dual
  Where Not Exists (Select 1 From zlParameters Where ϵͳ = &n_System And ģ�� = 1121 And ������ = '��첡�˷ֵ��ݴ�ӡ');

--107799:������,2017-04-05,ҽ���嵥�����������Ի�������¼
Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1252, 1, 1, 0, 1, 24, 'ҽ��״̬����', Null, '0',
         '����ҽ���嵥ҳǩѡ���¼:0-ҽ��,3-����'
  From Dual
  Where Not Exists
   (Select 1 From zlParameters Where ������ = 'ҽ�����˷�ʽ' And Nvl(ģ��, 0) = 1252 And Nvl(ϵͳ, 0) = &n_System);

Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1252, 1, 1, 0, 1, 59, '����鿴����', Null, '0',
         '����ҽ���嵥ѡ�񱨸�ҳǩʱ�Ĺ���������¼:0-ȫ��,1-���,2-����,3-����'
  From Dual
  Where Not Exists
   (Select 1 From zlParameters Where ������ = '����鿴����' And Nvl(ģ��, 0) = 1252 And Nvl(ϵͳ, 0) = &n_System);

Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1252, 1, 1, 0, 1, 60, '���������Զ�����', Null, '0',
         '����ҽ���嵥�еĹ��������������Ƿ��Զ�����:0-������,1-����'
  From Dual
  Where Not Exists (Select 1
         From zlParameters
         Where ������ = '���������Զ�����' And Nvl(ģ��, 0) = 1252 And Nvl(ϵͳ, 0) = &n_System);

Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1253, 1, 1, 0, 1, 59, '����鿴����', Null, '0',
         'סԺҽ���嵥ѡ�񱨸�ҳǩʱ�Ĺ���������¼:0-ȫ��,1-���,2-����,3-����'
  From Dual
  Where Not Exists
   (Select 1 From zlParameters Where ������ = '����鿴����' And Nvl(ģ��, 0) = 1253 And Nvl(ϵͳ, 0) = &n_System);

Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1253, 1, 1, 0, 1, 60, '���������Զ�����', Null, '0',
         'סԺҽ�����嵥�еĹ��������������Ƿ��Զ�����:0-������,1-����'
  From Dual
  Where Not Exists (Select 1
         From zlParameters
         Where ������ = '���������Զ�����' And Nvl(ģ��, 0) = 1253 And Nvl(ϵͳ, 0) = &n_System);

Insert Into zlParameters
  (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ������, ������, ����ֵ, ȱʡֵ, ����˵��)
  Select Zlparameters_Id.Nextval, &n_System, 1253, 1, 1, 0, 1, 61, 'ҽ����ʾ����', Null, '0', 'סԺҽ���嵥�еĹ�������:0-����ҽ��,1-����ҽ��'
  From Dual
  Where Not Exists
   (Select 1 From zlParameters Where ������ = 'ҽ����ʾ����' And Nvl(ģ��, 0) = 1253 And Nvl(ϵͳ, 0) = &n_System);


--107566:�ƽ�,2017-04-05,ͼ���Ľ���Ϣ����ҽ��ID
Insert Into Ӱ��ͼ����Ϣ��
  (Id, ��ʼ��ַ, ������ַ, Ӣ������, ��������, ���ļ��, Ӣ�ļ��, ����, ��ѡ��, λ��, �������, ʹ�ü���)
Values
  (Ӱ��ͼ����Ϣ��_Id.Nextval, 3, 3, 'cal', 'DBҽ��ID', '[ҽ��ID]', '[OrderID]', -1, 0, 0, 0, 0);


-------------------------------------------------------------------------------
--Ȩ����������
-------------------------------------------------------------------------------
--109070:���ջ�,2017-06-02,���ӷ������ģ��
Insert Into zlPrograms
  (���, ����, ˵��, ϵͳ, ����)
  Select 2228 ���, '�������' ����, '���ڶԷ��Ľ�����˲���' As ˵��, &n_System ϵͳ, 'zl9EmrInterface' ����
  From Dual
  Where Not Exists (Select 1 From zlPrograms Where ��� = 2228 And ����='�������');

Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��,ȱʡֵ)
Select &n_System,2228,A.* From (
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0 Union All
Select '����',0,'������Ϣ',1 From Dual Union All
Select '�������',1,'�Բ�����Ա�������ҺͲ�������˽�з��ĵ����Ȩ��',1 From Dual Union All
Select 'ȫԺ���',2,'�����з��ĵ����Ȩ��',0 From Dual Union All
Select '�����޸�',3,'�Է��ı���,����,˵��,����ԭ��,���÷��ĵ��޸�Ȩ��',0 From Dual Union All
Select '���ݱ༭',4,'�Է������ݽ����޸ĵ�Ȩ��',0 From Dual) A
Where Not Exists (Select 1 From zlProgFuncs Where ��� = 2228 And ����='����');

Insert Into zlMenus(���, ID, �ϼ�id, ����, ˵��, ϵͳ, ģ��, �̱���, ͼ��)
Select A.���,ZlMenus_ID.Nextval,A.ID,B.* From (
Select ���,ID From zlMenus Where ���� = '�����ĵ�����' And ��� = 'ȱʡ' And ϵͳ = &n_System And ģ�� Is Null) A,
(Select ����, ˵��, ϵͳ, ģ��, �̱���, ͼ�� From zlMenus Where 1=0 Union ALL
Select '�������','���ڶԷ��Ľ�����˲���',&n_System,2228,'�������',114 From Dual) B
Where Not Exists (Select 1 From zlMenus Where ģ�� = 2228 And ����='�������');

--100722:�ŵ���,2017-06-01,�������ڸı䲻�ܶ�ȡ���ݵ�����
Insert Into zlProgPrivs(ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1340, '��ɾ��', User, 'zl_��ҩ����_ҵ�����', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1340 And ���� = '��ɾ��' And Upper(����) = Upper('zl_��ҩ����_ҵ�����'));

--106745:������,2017-05-31,�Һż���װ
Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1115, 'ԤԼ�ҺŵǼ�', User, 'Zl_Fun_���˹Һż�¼_Check', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1115 And ���� = 'ԤԼ�ҺŵǼ�' And Upper(����) = Upper('Zl_Fun_���˹Һż�¼_Check'));  

--106745:������,2017-05-31,�Һŷ�װ���
Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1260, '���˹Һ�', User, 'Zl_Fun_���˹Һż�¼_Check', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1260 And ���� = '���˹Һ�' And Upper(����) = Upper('Zl_Fun_���˹Һż�¼_Check'));  
         
Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1260, 'ԤԼ�Һ�', User, 'Zl_Fun_���˹Һż�¼_Check', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1260 And ���� = 'ԤԼ�Һ�' And Upper(����) = Upper('Zl_Fun_���˹Һż�¼_Check'));  

--106745:������,2017-05-31,�Һŷ�װ���
Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1111, '���շѺ�', User, 'Zl_Fun_���˹Һż�¼_Check', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1111 And ���� = '���շѺ�' And Upper(����) = Upper('Zl_Fun_���˹Һż�¼_Check'));  

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1111, '����Ѻ�', User, 'Zl_Fun_���˹Һż�¼_Check', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1111 And ���� = '����Ѻ�' And Upper(����) = Upper('Zl_Fun_���˹Һż�¼_Check'));  

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1111, 'ԤԼ�Һ�', User, 'Zl_Fun_���˹Һż�¼_Check', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1111 And ���� = 'ԤԼ�Һ�' And Upper(����) = Upper('Zl_Fun_���˹Һż�¼_Check'));  

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1111, '����ԤԼ', User, 'Zl_Fun_���˹Һż�¼_Check', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1111 And ���� = '����ԤԼ' And Upper(����) = Upper('Zl_Fun_���˹Һż�¼_Check'));  

--108762:Ƚ����,2017-05-15,�ٴ����ﰲ��ҽ������ǰ��ʾְ�Ʊ�ʶ��
Insert Into zlProgFuncs
  (ϵͳ, ���, ����, ����, ˵��, ȱʡֵ)
  Select &n_System, 1114, 'ְ�Ʊ�ʶ����', 24, '���и�Ȩ��ʱ�����Զ���ʾ���ٴ�������ҽ������ǰ��ҽ��ְ�Ʊ�ʶ���������á�', 0
  From Dual
  Where Not Exists (Select 1 From zlProgFuncs Where ϵͳ = &n_System And ��� = 1114 And ���� = 'ְ�Ʊ�ʶ����');

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1114, '����', User, 'רҵ����ְ��', 'SELECT'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1114 And ���� = '����' And ���� = 'רҵ����ְ��');

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1114, 'ְ�Ʊ�ʶ����', User, 'Zl_רҵ����ְ��_���±�ʶ��', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1114 And ���� = 'ְ�Ʊ�ʶ����' And Upper(����) = Upper('Zl_רҵ����ְ��_���±�ʶ��'));

--108534:���ջ�,2017-5-11,����ȡ���������ģ��
Insert Into zlPrograms
  (���, ����, ˵��, ϵͳ, ����)
  Select 2227 ���, 'ȡ���������' ����, '�����ڲ�����ɺ���Ҫ�ٴ��޸�ʱ������������' As ˵��, &n_System ϵͳ, 'zl9EmrInterface' ����
  From Dual
  Where Not Exists (Select 1 From zlPrograms Where ��� = 2227 And ����='ȡ���������');

Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��,ȱʡֵ)
Select &n_System,2227,A.* From (
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0 Union All
Select '����',0,'������Ϣ',1 From Dual Union All
Select 'ϵͳ����',1,'��ϵͳ�Դ�����ķ���',1 From Dual) A
Where Not Exists (Select 1 From zlProgFuncs Where ��� = 2227 And ����='����');

Insert Into zlMenus(���, ID, �ϼ�id, ����, ˵��, ϵͳ, ģ��, �̱���, ͼ��)
Select A.���,ZlMenus_ID.Nextval,A.ID,B.* From (
Select ���,ID From zlMenus Where ���� = '�ʿ�ϵͳ����' And ��� = 'ȱʡ' And ϵͳ = &n_System And ģ�� Is Null) A,
(Select ����, ˵��, ϵͳ, ģ��, �̱���, ͼ�� From zlMenus Where 1=0 Union ALL
Select 'ȡ���������','�����ڲ�����ɺ���Ҫ�ٴ��޸�ʱ������������',&n_System,2227,'ȡ���������',109 From Dual) B
Where Not Exists (Select 1 From zlMenus Where ģ�� = 2227 And ����='ȡ���������');

--100642:��¶¶,2017-05-11,��������Ϣ����������ɾ����Ȩ�޶��������Էֿ�������Ȩ������
Insert Into zlProgFuncs
  (ϵͳ, ���, ����, ����, ˵��, ȱʡֵ)
  Select &n_System, 1101, '����', 22, '���Ӳ�����Ϣ�Ĳ���Ȩ�ޡ��и�Ȩ��ʱ������Բ�����Ϣ�������Ӳ�����', 1
  From Dual
  Where Not Exists (Select 1 From zlProgFuncs Where ϵͳ = &n_System And ��� = 1101 And ���� = '����');

Insert Into zlProgFuncs
  (ϵͳ, ���, ����, ����, ˵��, ȱʡֵ)
  Select &n_System, 1101, '�޸�', 23, '�޸Ĳ�����Ϣ�Ĳ���Ȩ�ޡ��и�Ȩ��ʱ������Բ�����Ϣ�����޸Ĳ�����', 1
  From Dual
  Where Not Exists (Select 1 From zlProgFuncs Where ϵͳ = &n_System And ��� = 1101 And ���� = '�޸�');

Insert Into zlProgFuncs
  (ϵͳ, ���, ����, ����, ˵��, ȱʡֵ)
  Select &n_System, 1101, 'ɾ��', 24, 'ɾ��������Ϣ�Ĳ���Ȩ�ޡ��и�Ȩ��ʱ������Բ�����Ϣ�����޸Ĳ�����', 1
  From Dual
  Where Not Exists (Select 1 From zlProgFuncs Where ϵͳ = &n_System And ��� = 1101 And ���� = 'ɾ��');

Insert Into zlProgFuncs
  (ϵͳ, ���, ����, ����, ˵��, ȱʡֵ)
  Select &n_System, 1101, '��ͣ', 25, '���ú�ͣ�ò�����Ϣ�Ĳ���Ȩ�ޡ��и�Ȩ��ʱ������Բ�����Ϣ����ͣ�á�ȡ��ͣ�ò�����', 1
  From Dual
  Where Not Exists (Select 1 From zlProgFuncs Where ϵͳ = &n_System And ��� = 1101 And ���� = '��ͣ');

Insert Into Zlprogrelas
  (ϵͳ, ���, ���, ����, ��ϵ, ����, �����ϵ)
  Select &n_System, 1101, 1, '����', 2, 0, 0 From Dual
  Where Not Exists (Select 1 From Zlprogrelas Where ϵͳ = &n_System And ��� = 1101 and ���=1 And ���� = '����');
  
Insert Into Zlprogrelas
  (ϵͳ, ���, ���, ����, ��ϵ, ����, �����ϵ)
  Select &n_System, 1101, 1, '�޸�', 2, 0, 0 From Dual
   Where Not Exists (Select 1 From Zlprogrelas Where ϵͳ = &n_System And ��� = 1101 and ���=1 And ���� = '�޸�');
  
Insert Into Zlprogrelas
  (ϵͳ, ���, ���, ����, ��ϵ, ����, �����ϵ)
  Select &n_System, 1101, 1, 'ɾ��', 2, 0, 0 From Dual
   Where Not Exists (Select 1 From Zlprogrelas Where ϵͳ = &n_System And ��� = 1101 and ���=1 And ���� = 'ɾ��');
  
Insert Into Zlprogrelas
  (ϵͳ, ���, ���, ����, ��ϵ, ����, �����ϵ)
  Select &n_System, 1101, 1, '��ͣ', 2, 0, 0 From Dual
   Where Not Exists (Select 1 From Zlprogrelas Where ϵͳ = &n_System And ��� = 1101 and ���=1 And ���� = '��ͣ');

--107898:Ƚ����,2017-05-08,�����ٴ������Դ���������䡱�������⣬ͬ�����ӡ������Դ���á��ԡ��Ա𡱱�ġ�SELECT��Ȩ��
Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1114, '�����Դ����', User, '�Ա�', 'SELECT'
  From Dual
  Where Not Exists (Select 1 From zlProgPrivs Where ϵͳ = &n_System And ��� = 1114 And ���� = '�����Դ����' And ���� = '�Ա�');

--108825:���ϴ�,2017-05-05,����Ʊ�ݴ�ӡ��Ȩ
Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1802, '����', User, 'Zl_���˹Һ�Ʊ��_Insert', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1802 And ���� = '����' And Upper(����) = Upper('Zl_���˹Һ�Ʊ��_Insert'));

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1803, '����', User, 'Zl_���˹Һ�Ʊ��_Insert', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1803 And ���� = '����' And Upper(����) = Upper('Zl_���˹Һ�Ʊ��_Insert'));

--98580:�ŵ���,2017-05-04,���׷�������Ȩ���ж��ܷ��޸�
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��,ȱʡֵ)
Select &n_System,1054,A.* From (
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0 Union All 
    Select '�޸ĸ��˳��׷���',10,'�и�Ȩ��ʱ������Ա�����޸ĸ��˵ĳ��׷�����',1 From Dual Union All 
    Select '�޸Ŀ��ҳ��׷���',11,'�и�Ȩ��ʱ������Ա�����޸Ŀ��ҵĳ��׷�����',1 From Dual Union All
    Select '�޸�ȫԺ���׷���',12,'�и�Ȩ��ʱ������Ա�����޸�ȫԺ�ĳ��׷�����',1 From Dual Union All
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0) A;

Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��,ȱʡֵ)
Select &n_System,1009,A.* From (
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0 Union All 
    Select '�޸ĸ��˳��׷���',16,'�и�Ȩ��ʱ������Ա�����޸ĸ��˵ĳ��׷�����',1 From Dual Union All 
    Select '�޸Ŀ��ҳ��׷���',17,'�и�Ȩ��ʱ������Ա�����޸Ŀ��ҵĳ��׷�����',1 From Dual Union All
    Select '�޸�ȫԺ���׷���',18,'�и�Ȩ��ʱ������Ա�����޸�ȫԺ�ĳ��׷�����',1 From Dual Union All
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0) A;

--107559:Ƚ����,2017-04-17,������ֹͣ�ﰲ�Ź���
Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1114, 'ͣ������', User, 'Zl_�ٴ�����ͣ��_Stop', 'EXECUTE'
  From Dual
  Where Not Exists (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1114 And ���� = 'ͣ������' And Upper(����) = Upper('Zl_�ٴ�����ͣ��_Stop'));

--106708:Ƚ����,2017-04-07,Υ���淶������
Delete From zlProgPrivs
Where ϵͳ = &n_System And ��� = 1114 And ���� = '���ﰲ��' And Upper(����) = Upper('Zl_Buildregisterfixedrule');

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1114, '���ﰲ��', User, 'Zl_�ٴ������_Addbyfixedrule', 'EXECUTE'
  From Dual
  Where Not Exists
   (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1114 And ���� = '���ﰲ��' And Upper(����) = Upper('Zl_�ٴ������_Addbyfixedrule'));

Delete From zlProgPrivs
Where ϵͳ = &n_System And ��� = 1114 And ���� = '���ﰲ��' And Upper(����) = Upper('Zl_Buildregisterplanbyrecord');

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1114, '���ﰲ��', User, 'Zl_�ٴ������_Addbyrecord', 'EXECUTE'
  From Dual
  Where Not Exists
   (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1114 And ���� = '���ﰲ��' And Upper(����) = Upper('Zl_�ٴ������_Addbyrecord'));

Delete From zlProgPrivs
Where ϵͳ = &n_System And ��� = 1114 And ���� = '���ﰲ��' And Upper(����) = Upper('Zl_Buildregisterplanbytemplet');

Insert Into zlProgPrivs
  (ϵͳ, ���, ����, ������, ����, Ȩ��)
  Select &n_System, 1114, '���ﰲ��', User, 'Zl_�ٴ������_Addbytemplet', 'EXECUTE'
  From Dual
  Where Not Exists
   (Select 1
         From zlProgPrivs
         Where ϵͳ = &n_System And ��� = 1114 And ���� = '���ﰲ��' And Upper(����) = Upper('Zl_�ٴ������_Addbytemplet'));

--105824:��˼��,2017-04-05,��1290 1291 ������ ���淢�� ���ܰ���Zl_Ӱ�񱨸淢�� ִ��Ȩ�� 
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��,ȱʡֵ)
Select &n_System,1290,A.* From (
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0 Union All 
Select '���淢��',39,'��ϱ���ķ���',1 From Dual Union All 
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0) A;

Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��,ȱʡֵ)
Select &n_System,1291,A.* From (
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0 Union All 
Select '���淢��',36,'��ϱ���ķ���',1 From Dual Union All 
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select  &n_System,1290,'���淢��',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All 
Select 'Zl_Ӱ�񱨸淢��','EXECUTE' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

Insert Into zlProgPrivs(ϵͳ,���,����,������,����,Ȩ��)
Select  &n_System,1291,'���淢��',User,A.* From (
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0 Union All 
Select 'Zl_Ӱ�񱨸淢��','EXECUTE' From Dual Union All
Select ����,Ȩ�� From zlProgPrivs Where 1 = 0) A;

--107370:��С��,2017-03-27,�°滤ʿվ-��������������˸ò������в���
Insert Into zlProgFuncs(ϵͳ,���,����,����,˵��,ȱʡֵ)
Select &n_System,1265,A.* From (
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0 Union All
Select '���Ʋ���',9,'ӵ�и�Ȩ��ʱ����������ɹ��˳��ò������в��ˣ��޸�Ȩ��ʱֻ�ܹ��˳���������',1 From Dual Union All
Select ����,����,˵��,ȱʡֵ From zlProgFuncs Where 1 = 0) A;

--107370:��С��,2017-03-27,�°滤ʿվ-��������������˸ò������в���
Insert Into zlRoleGrant
  (ϵͳ, ���, ��ɫ, ����)
  Select Distinct ϵͳ, ���, ��ɫ, '���Ʋ���' ����
  From zlRoleGrant A
  Where ϵͳ = &n_System And ��� = 1265 And ���� = '����';

--99878:Ƚ����,2017-05-27,ͬһ�շ���Ŀ����۸����
Insert Into Zlmodulerelas
  (ϵͳ, ģ��, ����, ���ϵͳ, ���ģ��, �������, ��ع���, ȱʡֵ)
  Select &n_System, 1107, Null, &n_System, 9000, 1, '����', 1 From Dual union all
  Select &n_System, 1111, Null, &n_System, 9000, 1, '����', 1 From Dual union all
  Select &n_System, 1115, Null, &n_System, 9000, 1, '����', 1 From Dual union all
  Select &n_System, 1120, Null, &n_System, 9000, 1, '����', 1 From Dual union all
  Select &n_System, 1121, Null, &n_System, 9000, 1, '����', 1 From Dual union all
  Select &n_System, 1122, Null, &n_System, 9000, 1, '����', 1 From Dual union all
  Select &n_System, 1133, Null, &n_System, 9000, 1, '����', 1 From Dual union all
  Select &n_System, 1134, Null, &n_System, 9000, 1, '����', 1 From Dual union all
  Select &n_System, 1135, Null, &n_System, 9000, 1, '����', 1 From Dual union all
  Select &n_System, 1139, Null, &n_System, 9000, 1, '����', 1 From Dual;




-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------






-------------------------------------------------------------------------------
--������������
-------------------------------------------------------------------------------
--100722:�ŵ���,2017-06-01,�޸ķ�ҩ����ͬ������ҵ������
Create Or Replace Procedure Zl_��ҩ����_ҵ�����
(
  ҩ��id_In In Number,
  �ɴ���_In In Varchar2,
  �´���_In In Varchar2
) Is

  Cursor c_δ������ Is
    Select ����, No, �ⷿid
    From δ��ҩƷ��¼
    Where �������� Between Sysdate - 3 And Sysdate And ��ҩ���� = �ɴ���_In And �ⷿid = ҩ��id_In;

  --ҩ������
  Cursor c_ҩ������ Is
    Select a.����ֵ
    From (Select ������, ����ֵ From Zluserparas Where ����id = 1687) a,
         (Select ������, ����ֵ From Zluserparas Where ����id = 1688) b
    Where a.������ = b.������ And b.����ֵ = ҩ��id_In;

  v_δ������ c_δ������%Rowtype;
  v_ҩ������ c_ҩ������%Rowtype;
Begin
  --���ò���
  Update Zluserparas
  Set ����ֵ = ҩ��id_In || ':' || �´���_In
  Where ����ֵ = ҩ��id_In || ':' || �ɴ���_In And
        ����id In (Select Id From Zlparameters Where ������ In ('��ҩ������', '��ҩ������', '��ҩ������'));

  --ҵ������
  For v_δ������ In c_δ������ Loop
    Update ҩƷ�շ���¼
    Set ��ҩ���� = �´���_In
    Where ���� = v_δ������.���� And No = v_δ������.No And �ⷿid = v_δ������.�ⷿid And ��ҩ���� = �ɴ���_In;
    Update ������ü�¼
    Set ��ҩ���� = �´���_In
    Where No = v_δ������.No And ִ�в���id = v_δ������.�ⷿid And ��ҩ���� = �ɴ���_In;
    Update סԺ���ü�¼
    Set ��ҩ���� = �´���_In
    Where No = v_δ������.No And ִ�в���id = v_δ������.�ⷿid And ��ҩ���� = �ɴ���_In;
  End Loop;

  Update δ��ҩƷ��¼
  Set ��ҩ���� = �´���_In
  Where �������� Between Sysdate - 3 And Sysdate And ��ҩ���� = �ɴ���_In And �ⷿid = ҩ��id_In;

  --ҩƷ����
  Update Zluserparas
  Set ����ֵ = Replace(����ֵ, �ɴ���_In, �´���_In)
  Where ����id = 1687 And ������ In (Select ������ From Zluserparas Where ����ֵ = ҩ��id_In And ����id = 1688);

  --�кŴ���
  Update ��ҩ���� Set �кŴ��� = �´���_In Where ҩ��id = ҩ��id_In And �кŴ��� = �ɴ���_In;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_��ҩ����_ҵ�����;
/

--106745:������,2017-06-01,�Һż���װ
Create Or Replace Function Zl_Fun_���˹Һż�¼_Check
(
  ������ʽ_In   Integer,
  ����id_In     ������ü�¼.����id%Type,
  ����_In       �ҺŰ���.����%Type,
  �����¼id_In �ٴ������¼.Id%Type := Null,
  ����ʱ��_In   ������ü�¼.����ʱ��%Type,
  ר�Һ�_In     Number := 0
) Return Varchar2 As
  --���ܣ��Һ���Ч�Լ��(����ԤԼ;ԤԼ�ҺŲ��ۿ�;ԤԼ�Һſۿ�)
  --���:������ʽ_IN:0-�Һ�(�����տ�ԤԼ),1-ԤԼ,2-ԤԼ����
  --     �Ƿ�Ӻ�_In:�Ƿ�Ӻŵ��ã�0-�ǼӺŵ��ã�1-�Ӻŵ���
  --����:0-���ͨ��
  --     1-�ض��������ʧ�ܣ�ͬʱ���ش�����ʾ�ı�
  --     2-���������µļ��ʧ�ܣ�ͬʱ���ش�����ʾ�ı�
  Err_Item Exception;
  n_����ԤԼ������ Number(18);
  n_��Լ����       Number(18);
  v_Temp           Varchar2(500);
  v_����ԭ��       ���ⲡ��.����ԭ��%Type;
  n_ͬ���޺���     Number;
  n_ͬ����Լ��     Number;
  n_����id         �ҺŰ���.����id%Type;
  n_Count          Number(18);
  n_���˹Һſ����� Number;
  n_ר�ҺŹҺ����� Number;
  n_ר�Һ�ԤԼ���� Number;
  n_ר�Һ�         Number;
  d_��Чʱ��       Date;
  n_�ƻ�id         �ҺŰ��żƻ�.Id%Type;

  Cursor c_Pati(n_����id ������Ϣ.����id%Type) Is
    Select a.����id, a.����, a.�Ա�, a.����, a.סԺ��, a.�����, a.�ѱ�, a.����, c.���� As ���ʽ
    From ������Ϣ A, ҽ�Ƹ��ʽ C
    Where a.����id = n_����id And a.ҽ�Ƹ��ʽ = c.����(+);

  r_Pati c_Pati%RowType;

  Function Zl_����Ա
  (
    Type_In     Integer,
    Splitstr_In Varchar2
  ) Return Varchar2 As
    n_Step Number(18);
    v_Sub  Varchar2(1000);
    --Type_In:0-��ȡȱʡ����ID;1-��ȡ����Ա���;2-��ȡ����Ա����
    -- SplitStr:��ʽΪ:����ID,��������;��ԱID,��Ա���,��Ա����(��Zl_Identity��ȡ��)
  Begin
    If Type_In = 0 Then
      --ȱʡ����
      n_Step := Instr(Splitstr_In, ',');
      v_Sub  := Substr(Splitstr_In, 1, n_Step - 1);
      Return v_Sub;
    End If;
    If Type_In = 1 Then
      --����Ա����
      n_Step := Instr(Splitstr_In, ';');
      v_Sub  := Substr(Splitstr_In, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, 1, n_Step - 1);
      Return v_Sub;
    End If;
    If Type_In = 2 Then
      --����Ա����
      n_Step := Instr(Splitstr_In, ';');
      v_Sub  := Substr(Splitstr_In, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      n_Step := Instr(v_Sub, ',');
      v_Sub  := Substr(v_Sub, n_Step + 1);
      Return v_Sub;
    End If;
  End;

Begin
  --��ⲡ�����
  Open c_Pati(����id_In);
  n_Count := 0;
  Begin
    Fetch c_Pati
      Into r_Pati;
    n_Count := 1;
  Exception
    When Others Then
      n_Count := -1;
  End;
  If n_Count <= 0 Then
    Return '1|����δ�ҵ������ܼ�����';
  End If;
  --ԤԼ��������
  If ������ʽ_In = 1 Then
    Begin
      Select ����ԭ�� Into v_����ԭ�� From ���ⲡ�� Where ����ʱ�� Is Null And ����id = ����id_In And Rownum = 1;
      Return '1|�˲��������ⲡ�������У�ԭ�򣺡�' || v_����ԭ�� || '�����ܼ�����';
    Exception
      When Others Then
        Null;
    End;
  End If;

  --���Һ�ʱ��
  If Trunc(Sysdate) > Trunc(����ʱ��_In) Then
    Return '1|���ܹ���ǰ�ĺ�(' || To_Char(����ʱ��_In, 'yyyy-mm-dd') || ')��';
  End If;

  --����ID,��������;��ԱID,��Ա���,��Ա����
  v_Temp := Zl_Identity(0);
  If Nvl(v_Temp, ' ') = ' ' Then
    Return '1|��ǰ������Աδ���ö�Ӧ����Ա��ϵ,���ܼ�����';
  End If;

  n_ר�Һ� := ר�Һ�_In;
  If �����¼id_In Is Null Then
    Select ����id Into n_����id From �ҺŰ��� Where ���� = ����_In;
  Else
    Select ����id Into n_����id From �ٴ������¼ Where ID = �����¼id_In;
  End If;

  --���ϵͳ����
  v_Temp           := Nvl(zl_GetSysParameter('����ͬ���޹�N����', 1111), '0|0') || '|';
  n_ͬ���޺���     := To_Number(Substr(v_Temp, 1, Instr(v_Temp, '|') - 1));
  n_ͬ����Լ��     := To_Number(Nvl(zl_GetSysParameter('����ͬ����ԼN����', 1111), '0'));
  n_����ԤԼ������ := To_Number(Nvl(zl_GetSysParameter('����ԤԼ������', 1111), '0'));
  n_���˹Һſ����� := To_Number(Nvl(zl_GetSysParameter('���˹Һſ�������', 1111), '0'));
  n_ר�ҺŹҺ����� := To_Number(Nvl(zl_GetSysParameter('ר�ҺŹҺ�����'), '0'));
  n_ר�Һ�ԤԼ���� := To_Number(Nvl(zl_GetSysParameter('ר�Һ�ԤԼ����'), '0'));
  --�Բ������ƽ��м��
  If ������ʽ_In = 1 Then
    If Nvl(n_ͬ����Լ��, 0) <> 0 Or Nvl(n_����ԤԼ������, 0) <> 0 Then
      n_��Լ���� := 0;
      For c_Chkitem In (Select Distinct ִ�в���id
                        From ���˹Һż�¼
                        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 2 And ԤԼʱ�� Between Trunc(����ʱ��_In) And
                              Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ִ�в���id <> n_����id) Loop
        n_��Լ���� := n_��Լ���� + 1;
      End Loop;
      If n_��Լ���� >= Nvl(n_����ԤԼ������, 0) And Nvl(n_����ԤԼ������, 0) > 0 Then
        Return '1|ͬһ�������ͬʱ��ԤԼ[' || Nvl(n_����ԤԼ������, 0) || ']������,������ԤԼ��';
      End If;
    
      Select Count(1)
      Into n_Count
      From ���˹Һż�¼
      Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 2 And ԤԼʱ�� Between Trunc(����ʱ��_In) And
            Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ִ�в���id = n_����id;
      If n_Count >= Nvl(n_ͬ����Լ��, 0) And Nvl(n_ͬ����Լ��, 0) > 0 Then
        Return '1|�ò����Ѿ��ڸÿ���ԤԼ��' || n_Count || '��,������ԤԼ��';
      End If;
    End If;
    If Nvl(n_ר�Һ�ԤԼ����, 0) <> 0 And n_ר�Һ� = 1 Then
      If �����¼id_In Is Null Then
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 2 And ԤԼʱ�� Between Trunc(����ʱ��_In) And
              Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And �ű� = ����_In;
      Else
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 2 And �����¼id = �����¼id_In;
      End If;
      If n_Count >= Nvl(n_ר�Һ�ԤԼ����, 0) And Nvl(n_ר�Һ�ԤԼ����, 0) > 0 Then
        Return '1|�ò����Ѿ���������ԤԼ����,������ԤԼ��';
      End If;
    End If;
  Else
    If (Nvl(n_ͬ���޺���, 0) <> 0 Or Nvl(n_���˹Һſ�����, 0) <> 0) And ������ʽ_In = 0 Then
      n_��Լ���� := 0;
      For c_Chkitem In (Select Distinct ִ�в���id
                        From ���˹Һż�¼
                        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 1 And ����ʱ�� Between Trunc(����ʱ��_In) And
                              Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ִ�в���id <> n_����id) Loop
        n_��Լ���� := n_��Լ���� + 1;
      End Loop;
      If n_��Լ���� >= Nvl(n_���˹Һſ�����, 0) And Nvl(n_���˹Һſ�����, 0) > 0 Then
        Return '1|ͬһ�������ͬʱ�ܹҺ�[' || Nvl(n_���˹Һſ�����, 0) || ']������,�����ٹҺţ�';
      End If;
    
      Select Count(1)
      Into n_Count
      From ���˹Һż�¼
      Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 1 And ����ʱ�� Between Trunc(����ʱ��_In) And
            Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And ִ�в���id = n_����id;
      If n_Count >= Nvl(n_ͬ���޺���, 0) And Nvl(n_ͬ���޺���, 0) > 0 Then
        Return '1|�ò����Ѿ��ڸÿ��ҹҺ���' || n_Count || '��,�����ٹҺţ�';
      End If;
    End If;
  
    If Nvl(n_ר�ҺŹҺ�����, 0) <> 0 And n_ר�Һ� = 1 Then
      If �����¼id_In Is Null Then
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 1 And ����ʱ�� Between Trunc(����ʱ��_In) And
              Trunc(����ʱ��_In) + 1 - 1 / 24 / 60 / 60 And �ű� = ����_In;
      Else
        Select Count(1)
        Into n_Count
        From ���˹Һż�¼
        Where ����id = ����id_In And ��¼״̬ = 1 And ��¼���� = 1 And �����¼id = �����¼id_In;
      End If;
      If n_Count >= Nvl(n_ר�ҺŹҺ�����, 0) And Nvl(n_ר�ҺŹҺ�����, 0) > 0 Then
        Return '1|�ò����Ѿ��������ŹҺ�����,�����ٹҺţ�';
      End If;
    End If;
  End If;

  Return '0-��Դ����';

Exception
  When Others Then
    Return '2-' || SQLErrM;
End Zl_Fun_���˹Һż�¼_Check;
/

--109002:����,2017-05-31,��һ��������������

Create Or Replace Procedure Zl_���˻����ӡ_Update
(
  �ļ�id_In   In ���˻����ӡ.�ļ�id%Type,
  ����ʱ��_In In ���˻����ӡ.����ʱ��%Type,
  ����_In     In ���˻����ӡ.����%Type,
  ɾ��_In     Number := 0,
  ��������_In Number := 0
) Is
  n_Actives   Number;
  n_Rows      Number; --0-����,>0��ʾ�޸� 
  n_Startpage Number; --��ʼҳ 
  n_Startrow  Number; --��ʼ�� 
  n_Endpage   Number; --����ҳ 
  n_Endrow    Number; --������ 
  n_Count     Number; --����ʱ��֮����������� 
  n_Pagerows  Number; --ÿҳ��Ч������ 
  n_Del       Number;
  n_����      ���˻����ӡ.����%Type;
  n_Firstdata Number; --�Ƿ���¼��ĵ�һ������ 
  n_��¼id    ���˻�������.Id%Type;
  n_��¼oldid ���˻����ӡ.��¼id%Type;
  n_��ʽid    ���˻����ļ�.��ʽid%Type;
  d_����ʱ��  ���˻����ӡ.����ʱ��%Type;
  v_Username  ��Ա��.����%Type;

  Err_Item Exception;
  v_Err_Msg Varchar2(1000);
  v_Print   Varchar2(800);
Begin
  n_Del      := ɾ��_In;
  n_����     := ����_In;
  v_Username := Zl_Username;
  Select ��ʽid Into n_��ʽid From ���˻����ļ� Where ID = �ļ�id_In;

  If n_���� = 0 Then
    v_Err_Msg := '��Ч�����в��ܵ����㣬���¼���δ���Ĳ������̣�';
    Raise Err_Item;
  End If;

  Begin
    Select ��¼id, ����, ��ʼҳ��, ��ʼ�к�, ����ҳ��, �����к�
    Into n_��¼oldid, n_Rows, n_Startpage, n_Startrow, n_Endpage, n_Endrow
    From ���˻����ӡ
    Where �ļ�id = �ļ�id_In And ����ʱ�� = ����ʱ��_In;
  Exception
    When Others Then
      n_Rows := 0;
  End;

  --��ȡ�û����ļ���ʽÿҳ��Ч�����У����Ӵ����� 
  Select To_Number(�����ı�)
  Into n_Pagerows
  From �����ļ��ṹ
  Where �������� = '��Ч������' And ��id = (Select ID From �����ļ��ṹ Where �ļ�id = n_��ʽid And ������� = 1 And ��id Is Null);

  Select Count(*) Into n_Count From ���˻������� Where �ļ�id = �ļ�id_In And ����ʱ�� = ����ʱ��_In;

  --�޸�����ʱ,Ҳ����ɾ�� 
  If n_Del = 0 Then
    Begin
      If n_Count = 0 Then
        n_Del := 1;
      End If;
      If n_Count > 1 Then
        v_Err_Msg := '�ڷ���ʱ�䡾' || To_Char(����ʱ��_In, 'YYYY-MM-DD hh24:mi:ss') || '���Ѿ�������Ӧ�����ݣ��������ٴ�¼����޸����ݵ�ʱ��Ϊ�˷���ʱ�䣡';
        Raise Err_Item;
      End If;
    End;
  Elsif n_Del = 1 And n_Count > 0 Then
    n_Del  := 0;
    n_���� := 1;
  End If;

  n_Firstdata := 0;
  If n_Del = 1 Then
    Delete ���˻����ӡ Where �ļ�id = �ļ�id_In And ����ʱ�� = ����ʱ��_In;
    n_Rows := n_Rows * -1;
  Else
    Select ID Into n_��¼id From ���˻������� Where �ļ�id = �ļ�id_In And ����ʱ�� = ����ʱ��_In;
  
    If n_Rows = 0 Then
      --�������д�ӡ���ݼ���Ҫ��������ݣ��������ʼҳ�ţ��кţ�����ҳ�ţ��к� 
      Select Min(����ʱ��) Into d_����ʱ�� From ���˻����ӡ Where �ļ�id = �ļ�id_In And ����ʱ�� > ����ʱ��_In;
      If d_����ʱ�� Is Null Then
        Select Max(����ʱ��) Into d_����ʱ�� From ���˻����ӡ Where �ļ�id = �ļ�id_In And ����ʱ�� < ����ʱ��_In;
        If d_����ʱ�� Is Null Then
          n_Startpage := 1;
          n_Startrow  := 1;
          n_Firstdata := 1;
        Else
          Select ����ҳ��, �����к�
          Into n_Startpage, n_Startrow
          From ���˻����ӡ
          Where �ļ�id = �ļ�id_In And ����ʱ�� = d_����ʱ��;
          n_Startrow := n_Startrow + 1;
        End If;
      Else
        Select ��ʼҳ��, ��ʼ�к�
        Into n_Startpage, n_Startrow
        From ���˻����ӡ
        Where �ļ�id = �ļ�id_In And ����ʱ�� = d_����ʱ��;
      End If;
    
      --У��ҳ��,�к� 
      If n_Startrow > n_Pagerows Then
        n_Startpage := n_Startpage + 1;
        n_Startrow  := n_Startrow - n_Pagerows;
      
        --��ҳʱ���Զ����ݵ�ǰҳ�����ò�����ҳ�Ļ��Ŀ���� 
        Begin
          Select 1 Into n_Actives From ���˻�����Ŀ Where �ļ�id = �ļ�id_In And ҳ�� = n_Startpage And Rownum < 2;
        Exception
          When Others Then
            n_Actives := 0;
        End;
      
        If n_Actives = 0 Then
          Insert Into ���˻�����Ŀ
            (�ļ�id, ҳ��, �к�, ��ͷ����, ���, ��Ŀ���, ��λ, ����Ա, ����ʱ��)
            Select �ļ�id, n_Startpage, �к�, ��ͷ����, ���, ��Ŀ���, ��λ, v_Username, Sysdate
            From ���˻�����Ŀ
            Where �ļ�id = �ļ�id_In And ҳ�� = n_Startpage - 1;
        End If;
      End If;
      n_Endpage := n_Startpage;
      n_Endrow  := n_Startrow + n_���� - 1;
      If n_Endrow > n_Pagerows Then
        --��������������ݳ���һҳ����� 
        n_Endpage := n_Endpage + 1;
        n_Endrow  := n_Endrow - n_Pagerows;
      
        --��ҳʱ���Զ����ݵ�ǰҳ�����ò�����ҳ�Ļ��Ŀ���� 
        Begin
          Select 1 Into n_Actives From ���˻�����Ŀ Where �ļ�id = �ļ�id_In And ҳ�� = n_Endpage And Rownum < 2;
        Exception
          When Others Then
            n_Actives := 0;
        End;
      
        If n_Actives = 0 Then
          Insert Into ���˻�����Ŀ
            (�ļ�id, ҳ��, �к�, ��ͷ����, ���, ��Ŀ���, ��λ, ����Ա, ����ʱ��)
            Select �ļ�id, n_Endpage, �к�, ��ͷ����, ���, ��Ŀ���, ��λ, v_Username, Sysdate
            From ���˻�����Ŀ
            Where �ļ�id = �ļ�id_In And ҳ�� = n_Endpage - 1;
        End If;
      End If;
      --������¼�����ҳ������ 
      If n_Endrow > n_Pagerows Or n_Endpage - n_Startpage > 1 Then
        If ��������_In = 1 Then
          n_����   := n_���� - n_Endrow + n_Pagerows;
          n_Endrow := n_Pagerows;
        Else
          v_Err_Msg := '���ڷ���ʱ�䡾' || To_Char(����ʱ��_In, 'YYYY-MM-DD hh24:mi:ss') || '��¼������ݴ��ڴ���¼������ݲ���������һҳ���ϣ�';
          Raise Err_Item;
        End If;
      End If;
    
      Insert Into ���˻����ӡ
        (��¼id, �ļ�id, ����ʱ��, ����, ��ʼҳ��, ��ʼ�к�, ����ҳ��, �����к�)
      Values
        (n_��¼id, �ļ�id_In, ����ʱ��_In, n_����, n_Startpage, n_Startrow, n_Endpage, n_Endrow);
      --�²�������ݵ��������ǲ�ֵ 
      n_Rows := n_����;
    Else
      --������ԭ�����Ĳ�ֵ 
      n_Rows := n_���� - n_Rows;
      --У��ҳ��,�к� 
      n_Endrow := n_Endrow + n_Rows;
      If n_Endrow <= 0 Then
        n_Endrow  := n_Pagerows + n_Endrow;
        n_Endpage := n_Endpage - 1;
      End If;
      If n_Endrow > n_Pagerows Then
        --��������������ݳ���һҳ����� 
        n_Endpage := n_Endpage + 1;
        n_Endrow  := n_Endrow - n_Pagerows;
      End If;
    
      --������¼�����ҳ������ 
      If n_Endrow > n_Pagerows Or n_Endpage - n_Startpage > 1 Then
        v_Err_Msg := '���ڷ���ʱ�䡾' || To_Char(����ʱ��_In, 'YYYY-MM-DD hh24:mi:ss') || '��¼������ݴ��ڴ���¼������ݲ���������һҳ���ϣ�';
        Raise Err_Item;
      End If;
    
      --���´�ӡ���ݣ���ǰ���ݵĴ�ӡ�����ӡʱ�����ΪNULL��������ݲ����� 
      Update ���˻����ӡ
      Set �ļ�id = �ļ�id_In, ��¼id = n_��¼id, ����ʱ�� = ����ʱ��_In, ���� = n_����, ��ʼҳ�� = n_Startpage, ��ʼ�к� = n_Startrow,
          ����ҳ�� = n_Endpage, �����к� = n_Endrow, �в� = Decode(��ӡ��, Null, 0, n_Rows),
          --ֻ�д�ӡ�������ݲż�¼�в� 
          ��ӡ�� = Null, ��ӡʱ�� = Null
      Where ��¼id = n_��¼oldid;
    End If;
  End If;
  --���в�˳� 
  If n_Rows = 0 Then
    Return;
  End If;

  --֮���Ƿ�������ݣ� 
  Begin
    Select 1 Into n_Count From ���˻����ӡ Where �ļ�id = �ļ�id_In And ����ʱ�� > ����ʱ��_In And Rownum < 2;
  Exception
    When Others Then
      n_Count := 0;
  End;

  If n_Count = 1 Then
    --����֮�����ݵĴ�ӡ������ݣ�����ӡ�����ӡʱ�䣩 
    If n_Rows > 0 Then
      Update ���˻����ӡ
      Set ��ʼҳ�� = ��ʼҳ�� + Decode(Sign(��ʼ�к� + n_Rows - n_Pagerows), 1, 1, 0),
          ����ҳ�� = ����ҳ�� + Decode(Sign(�����к� + n_Rows - n_Pagerows), 1, 1, 0),
          ��ʼ�к� = Decode(Mod(��ʼ�к� + n_Rows, n_Pagerows), 0, n_Pagerows, Mod(��ʼ�к� + n_Rows, n_Pagerows)),
          �����к� = Decode(Mod(�����к� + n_Rows, n_Pagerows), 0, n_Pagerows, Mod(�����к� + n_Rows, n_Pagerows)), ��ӡ�� = Null,
          ��ӡʱ�� = Null
      Where �ļ�id = �ļ�id_In And ����ʱ�� > ����ʱ��_In;
    Else
      --�µ��к�С��1��ҳ��-1 
      --�µ��к�+ÿҳ����Ч�к��ٽ����ж� 
      Update ���˻����ӡ
      Set ��ʼҳ�� = ��ʼҳ�� - Decode(Sign(��ʼ�к� + n_Rows - 1), -1, 1, 0),
          ����ҳ�� = ����ҳ�� - Decode(Sign(�����к� + n_Rows - 1), -1, 1, 0),
          ��ʼ�к� = Decode(Mod(��ʼ�к� + n_Pagerows + n_Rows, n_Pagerows), 0, n_Pagerows,
                         Mod(��ʼ�к� + n_Pagerows + n_Rows, n_Pagerows)),
          �����к� = Decode(Mod(�����к� + n_Pagerows + n_Rows, n_Pagerows), 0, n_Pagerows,
                         Mod(�����к� + n_Pagerows + n_Rows, n_Pagerows)), ��ӡ�� = Null, ��ӡʱ�� = Null
      Where �ļ�id = �ļ�id_In And ����ʱ�� > ����ʱ��_In;
      --����Ӧ������ɾ�������ݲŸ��µģ����Բ������ҳ��Ϊ��ģ�ҳ��Ϊ��Ŀ϶��Ѿ�ɾ���ˡ� 
      --DELETE ���˻����ӡ WHERE ��ʼҳ��=0; 
    End If;
    --������֮��Ĵ�ӡ�����Ƿ����������һҳ���ϣ�����������ֹ�� 
    v_Print := '';
    For r_Print In (Select ����ʱ��, ��ʼҳ��
                    From ���˻����ӡ
                    Where �ļ�id = �ļ�id_In And ����ʱ�� > ����ʱ��_In And ����ҳ�� - ��ʼҳ�� > 1
                    Order By ����ʱ��) Loop
      If Lengthb(v_Print || Chr(13) || Chr(10) || 'ҳ�š�' || r_Print.��ʼҳ�� || '��    ����ʱ�䡾' ||
                 To_Char(r_Print.����ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '��') < 800 Then
        v_Print := v_Print || Chr(13) || Chr(10) || 'ҳ�š�' || r_Print.��ʼҳ�� || '��    ����ʱ�䡾' ||
                   To_Char(r_Print.����ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '��';
      End If;
    End Loop;
    If v_Print Is Not Null Then
      v_Err_Msg := '���ڷ���ʱ�䡾' || To_Char(����ʱ��_In, 'YYYY-MM-DD hh24:mi:ss') || '��¼�������Ӱ���˺�������λ�ã���������������������һҳ���ϣ�';
      v_Err_Msg := v_Err_Msg || v_Print || Chr(13) || Chr(10) || 'Ŀǰ��Ʒ�ݲ�֧�ֶԿ�һҳ���ϵ����ݽ���չʾ�ʹ�ӡ��������ֹ��';
      Raise Err_Item;
    End If;
  End If;
  --���й����ļ���ҳ������ 
  Zl_���˻����ӡ_Batchretrypage(�ļ�id_In, n_Firstdata || ';' || n_Firstdata);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˻����ӡ_Update;
/

--109518:������,2017-05-24,ÿ��ת�����ؽ���ת������
Create Or Replace Procedure Zl1_Datamoveout1
(
  Demoded_In        In Number,
  Optmode_In        In Number := 0,
  Curtime_In        In Number := 1,
  Totaltime_In      In Number := 1,
  Speedmode_In      In Number := 0,
  Disabletrigger_In In Number := 0,
  Disablejob_In     In Number := 0,
  Parallel_In       In Number := 0,
  Sysowner_In       In Varchar2 := Null,
  Peissysowner_In   In Varchar2 := Null,
  Opersysowner_In   In Varchar2 := Null
) As
  --���ܣ���ǲ�ת��n��ǰ�����ݵ���ʷ��ռ� 
  --����:Demoded_in:          ���ת����������ǰ������,������Optmode_InΪ0��1ʱ����Ч 
  --     Optmode_in:           0-��ǲ�ִ��ת��,1-ֻ���б�ǣ�2-ִֻ��ת��(���ѱ�ǵ�) 
  --     Curtime_in,Totaltime_in���������ת��ʱ�ĵ�ǰ�������ܴ����������Ϊ1��ʾһ����ת�� 
  --                �״�ʱ�������߱�����ʷ��Ľṹһ���ԡ����߱���ӱ��Ƿ�ת�������ҽ����������������ת�������÷�ת������������ 
  --                ���һ��ִ�к����ڽ���������ֹ��ָ����õ���������� 
  --     Speedmode_in:        0-����ģʽ��1-����ģʽ���ڿͻ���ͣ��ʱ��ת���ڼ����ת�����������Ψһ�������Լ�����������Լӿ�ɾ�������� 
  --                          ��ʷ���Լ��������������Ӧ�ó���ʱ���У���Ϊ��Ҫ�õ���ʷ������ӣ� 
  --     Disabletrigger_in:   1=ת���ڼ���õ�ǰ�����ߵĴ�������0-������ 
  --     Disablejob_in:       1=ת���ڼ���õ�ǰ�����ߵ��Զ���ҵ��0-������ 
  --     parallel_in:         �ؽ���ǲ�ѯ��������ʱ�Ĳ��жȣ�ȱʡΪ������ִ��
  --     SysOwner_In:         ��׼ϵͳָ��ת����ʷ��ռ�������
  --     PeisSysOwner_In:     ���ϵͳָ��ת����ʷ��ռ�������
  --     OperSysOwner_In:     ����ϵͳָ��ת����ʷ��ռ�������
  --˵����1.���Ҫת�������ݣ����Զ�α�ǣ�Ȼ�����ִ��ת�� 
  --      2.ת��ʱ������zlBakTables�ж���ķ����˳��ת�����ݣ������ύ����; 
  --      3.Ϊ�˱����ѯ��Χ̫�����������⣬��Undo��ռ�����̫�󣬽���ÿ�β�Ҫת��̫�������(����������ʱ�Զ����Ϊÿ�ε���תһ����); 
  d_End        Date;
  n_System     Number(5);
  v_Systems    Varchar2(100);
  n_Peissystem Number(5);
  n_Opersystem Number(5);
  n_Reset      Number(1) := 0;
  v_Sql        Varchar2(4000);
  v_Owner      Varchar2(20);

  v_Pre���      Number(2);
  v_��ǰ����     Number(8);
  v_����         Number(8);
  n_�ؽ�������� Zldatamove.�ؽ��������%Type;
  n_�ؽ�������Χ Zldatamove.�ؽ�������Χ%Type;

  Err_Item Exception;
  v_Err_Msg Varchar2(4000);

  --���ܣ�ת�����ݣ������ɾ���������ύ���� 
  Procedure Movedata
  (
    v_Table    In Varchar2,
    v_��ǰ���� In Varchar2,
    v_Owner    In Varchar2
  ) As
    v_Colstr Varchar2(4000);
  Begin
    Select f_List2str(Cast(Collect(Column_Name) As t_Strlist)) As Colsstr
    Into v_Colstr
    From (Select Column_Name From User_Tab_Columns Where Table_Name = v_Table Order By Column_Id);
  
    v_Sql := 'Insert Into /*+ append */ ' || v_Owner || '.' || v_Table || '(' || v_Colstr || ') Select ' || v_Colstr ||
             ' From ' || v_Table || ' Where ��ת�� = ' || v_��ǰ����;
    Execute Immediate v_Sql;
  
    v_Sql := 'Delete ' || v_Table || ' Where ��ת�� = ' || v_��ǰ����;
    Execute Immediate v_Sql;
    Commit;
    --ÿ�ű��ύһ�Σ�����Undoռ�ù��࣬��ʱ��ҵ���ѯ���ܱ�ora-01555����̫�ɵĴ���
  End Movedata;

  --�����ʷ��� 
  Function Checkvalid(v_Systems In Varchar2) Return Varchar2 Is
    n_ֻ�� Number(3);
    n_״̬ Number(1);
    v_Err  Varchar2(4000);
    v_Tmp1 Varchar2(4000);
    v_Tmp2 Varchar2(4000);
    v_Tmp3 Varchar2(4000);
  Begin
    Select Count(1)
    Into n_ֻ��
    From zlBakSpaces
    Where ϵͳ In (Select Column_Value From Table(f_Num2list(v_Systems))) And
          (������ = Sysowner_In Or ������ = Peissysowner_In Or ������ = Opersysowner_In) And ֻ�� = 1;
  
    If n_ֻ�� > 0 Then
      v_Err := '[ZLSOFT]����ֻ��״̬�ĵ�ǰ��ʷ���ݿռ�,�������ܼ���![ZLSOFT]';
      Return(v_Err);
    End If;
  
    --������飬�����˹�ת���ڼ䣬�Զ���ҵ�ֵ��ñ����� 
    Select Nvl(״̬, 0) Into n_״̬ From zlDataMove Where ϵͳ = n_System And ��� = 1;
    If n_״̬ = 1 Then
      v_Err := '[ZLSOFT]�����û����ڽ���ת���������������������������ֹ�����"zlDataMove.״̬"��ֵΪ��![ZLSOFT]';
      Return(v_Err);
    End If;
    --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- 
    If Optmode_In <> 2 Then
      --������߱���󱸱���ֶ��Ƿ�һ��,�Ա�������ת����һ����ʱ�ű��� 
      For R In (Select ���� From zlBakTables Where ϵͳ In (Select Column_Value From Table(f_Num2list(v_Systems)))) Loop
        v_Tmp1 := '';
        v_Tmp2 := '';
        v_Tmp3 := '';
        For C In (Select *
                  From (Select a.Column_Name, a.Data_Type, a.Data_Precision, b.Column_Name As Bcolumn_Name,
                                b.Data_Type As Bdata_Type, b.Data_Precision As Bdata_Precision
                         From (Select Column_Name, Data_Type,
                                       Decode(Data_Type, 'VARCHAR2', Data_Length, Data_Precision) As Data_Precision
                                From User_Tab_Columns A
                                Where Table_Name = r.����) A,
                              (Select Column_Name, Data_Type,
                                       Decode(Data_Type, 'VARCHAR2', Data_Length, Data_Precision) As Data_Precision
                                From All_Tab_Columns
                                Where Table_Name = r.���� And
                                      Owner In
                                      (Select ������
                                       From zlBakSpaces
                                       Where ϵͳ In (Select Column_Value From Table(f_Num2list(v_Systems))) And
                                             (������ = Sysowner_In Or ������ = Peissysowner_In Or ������ = Opersysowner_In))) B
                         Where a.Column_Name = b.Column_Name(+))
                  Where Bcolumn_Name Is Null Or Data_Type <> Bdata_Type Or Data_Precision > Bdata_Precision) Loop
        
          If c.Bcolumn_Name Is Null Then
            v_Tmp1 := v_Tmp1 || ',' || c.Column_Name || ' ' || c.Data_Type || '(' || c.Data_Precision || ')';
          Elsif c.Data_Type <> c.Bdata_Type Then
            If c.Data_Type = 'DATE' Then
              v_Tmp2 := v_Tmp2 || ',' || c.Column_Name || ' ' || c.Data_Type || ',��ʷ���Ϊ' || c.Bdata_Type;
            Else
              v_Tmp2 := v_Tmp2 || ',' || c.Column_Name || ' ' || c.Data_Type || '(' || c.Data_Precision || '),��ʷ���Ϊ' ||
                        c.Bdata_Type;
            End If;
          Else
            v_Tmp3 := v_Tmp3 || ',' || c.Column_Name || ' ' || c.Data_Type || '(' || c.Data_Precision || '),��ʷ���Ϊ' ||
                      c.Bdata_Precision;
          End If;
        End Loop;
      
        If v_Tmp1 Is Not Null Then
          v_Err := v_Err || Chr(10) || ',ȱ�ֶΣ�' || r.���� || ' ' || v_Tmp1;
        End If;
        If v_Tmp2 Is Not Null Then
          v_Err := v_Err || Chr(10) || ',���Ͳ�ͬ��' || r.���� || ' ' || v_Tmp2;
        End If;
        If v_Tmp3 Is Not Null Then
          v_Err := v_Err || Chr(10) || ',���Ƚ�С��' || r.���� || ' ' || v_Tmp3;
        End If;
      
        If Lengthb(v_Err) > 3000 Then
          v_Err := '[ZLSOFT]�뵽�������ߡ���ִ�С���ʷ��������' || Substr(v_Err, 1, 3000) || '......[ZLSOFT]';
          Return(v_Err);
        End If;
      End Loop;
    
      If v_Err Is Not Null Then
        v_Err := '[ZLSOFT]�뵽�������ߡ���ִ�С���ʷ��������:' || Substr(v_Err, 1, 3000) || '[ZLSOFT]';
        --�ؽ�H����ͼ�Ľű��������ʾ���� 
        --Select 'Create or replace view  ZLHIS.H' || ���� || ' as Select * From ZLBAK1.' || ���� || ';' From Zlbaktables Where ϵͳ In(Select Column_Value From Table(f_num2list(v_Systems))) 
        Return(v_Err);
      End If;
    
      --����������ʷ�����ű�����©����Щ����ʹ�õ�������ӱ�û��ɾ����Ϊ�˱���ת�Ƶ���;ʱ�ű����ȼ��һ�� 
      For P In (Select Constraint_Name
                From (Select Constraint_Name,
                              Row_Number() Over(Partition By Constraint_Name Order By Decode(Constraint_Type, 'P', 0, 1)) Rn
                       From User_Constraints A, zlBakTables B
                       Where b.���� = a.Table_Name And b.ϵͳ In (Select Column_Value From Table(f_Num2list(v_Systems))) And
                             a.Constraint_Type In ('P', 'U'))
                Where Rn = 1) Loop
        For R In (Select a.Table_Name, a.Constraint_Name, a.Delete_Rule
                  From User_Constraints A
                  Where a.r_Constraint_Name = p.Constraint_Name And Not Exists
                   (Select 1
                         From zlBakTables B
                         Where b.���� = a.Table_Name And b.ϵͳ In (Select Column_Value From Table(f_Num2list(v_Systems))))
                  Order By a.r_Constraint_Name) Loop
          v_Err := v_Err || Chr(10) || r.Table_Name || '(' || r.Constraint_Name || ',' || r.Delete_Rule || '->' ||
                   p.Constraint_Name || ')';
          If Lengthb(v_Err) > 2000 Then
            v_Err := '[ZLSOFT]�ӱ�δת��:' || Substr(v_Err, 1, 2000) || '......[ZLSOFT]';
            Return(v_Err);
          End If;
        End Loop;
      End Loop;
    
      If v_Err Is Not Null Then
        v_Err := '[ZLSOFT]�ӱ�δת��:' || Substr(v_Err, 1, 2000) || '[ZLSOFT]';
        Return(v_Err);
      End If;
    End If;
    Return('');
  End Checkvalid;
Begin
  If Optmode_In <> 2 Then
    Select Trunc(Sysdate) - Demoded_In Into d_End From Dual;
  End If;
  v_Owner := Zl_Owner;
  Select ��� Into n_System From zlSystems Where Upper(������) = v_Owner And ��� Like '1%';

  Select Nvl(Min(���), 0) Into n_Peissystem From zlSystems Where Upper(������) = v_Owner And ��� Like '21%';
  Select Nvl(Min(���), 0) Into n_Opersystem From zlSystems Where Upper(������) = v_Owner And ��� Like '24%';

  --1.��ȫ�Լ�� 
  ----------------------------------------------------------------------------------- 
  If Curtime_In = 1 Then
    v_Systems := n_System;
    If n_Peissystem > 0 Then
      v_Systems := v_Systems || ',' || n_Peissystem;
    End If;
    If n_Opersystem > 0 Then
      v_Systems := v_Systems || ',' || n_Opersystem;
    End If;
  
    v_Err_Msg := Checkvalid(v_Systems);
    If v_Err_Msg Is Not Null Then
      Raise Err_Item;
    End If;
  
    --һ���е��״ε���ʱ���ô���������ҵ 
    If Disabletrigger_In = 1 Then
      Zl1_Datamove_Reb(n_System, Speedmode_In, 1, 0);
      If n_Peissystem > 0 Then
        Zl1_Datamove_Reb(n_Peissystem, Speedmode_In, 1, 0);
      End If;
      If n_Opersystem > 0 Then
        Zl1_Datamove_Reb(n_Opersystem, Speedmode_In, 1, 0);
      End If;
    End If;
  
    If Disablejob_In = 1 Then
      Zl1_Datamove_Reb(n_System, Speedmode_In, 2, 0);
      If n_Peissystem > 0 Then
        Zl1_Datamove_Reb(n_Peissystem, Speedmode_In, 2, 0);
      End If;
      If n_Opersystem > 0 Then
        Zl1_Datamove_Reb(n_Opersystem, Speedmode_In, 2, 0);
      End If;
    End If;
  
    Update zlDataMove Set ״̬ = 1 Where ϵͳ = n_System And ��� = 1;
    Commit;
  End If;

  --2.���Ҫת�������� 
  ----------------------------------------------------------------------------------- 
  If Optmode_In <> 2 Then
    --�ϴα��ת�������������б��ת�� 
    Select Nvl(Max(����), 0) Into v_��ǰ���� From Zldatamovelog Where ϵͳ = n_System And ��ת�� = 2;
  
    If v_��ǰ���� = 0 Then
      Select Nvl(Max(����), 0) + 1, Decode(Curtime_In, 1, Nvl(Max(����), 0) + 1, Max(����))
      Into v_��ǰ����, v_����
      From Zldatamovelog
      Where ϵͳ = n_System;
    
      Insert Into Zldatamovelog
        (ϵͳ, ����, ����, ��ֹʱ��, ��ǿ�ʼʱ��, ��ת��, ��ǰ����)
      Values
        (n_System, v_��ǰ����, v_����, d_End, Sysdate, 2, '���ڱ�Ǵ�ת������');
      Commit;
    Else
      Update Zldatamovelog
      Set ��ǿ�ʼʱ�� = Sysdate, ��ǰ���� = '���ڱ�Ǵ�ת������'
      Where ϵͳ = n_System And ���� = v_��ǰ����;
      Commit;
    End If;
  
    Zl1_Datamove_Tag(d_End, v_��ǰ����, n_System);
    If n_Peissystem > 0 Then
      Execute Immediate 'Begin Zl21_Datamove_Tag(:1, :2, :3); End;'
        Using d_End, v_��ǰ����, n_Peissystem;
    End If;
    If n_Opersystem > 0 Then
      Execute Immediate 'Begin Zl24_Datamove_Tag(:1, :2, :3); End;'
        Using d_End, v_��ǰ����, n_Opersystem;
    End If;
  
    Update Zldatamovelog
    Set ��ǽ���ʱ�� = Sysdate, ��ǰ���� = '��Ǵ�ת���������', ��ת�� = 1
    Where ϵͳ = n_System And ���� = v_��ǰ����;
    Commit;
  End If;

  --3.ת�����ݴ��� 
  ----------------------------------------------------------------------------------- 
  If Optmode_In = 1 Then
    If Curtime_In = Totaltime_In Then
      Update zlDataMove Set ״̬ = Null Where ϵͳ = n_System And ��� = 1;
    End If;
    Commit;
  Else
    --����С�����ο�ʼִ��ת�� 
    If Optmode_In = 2 Then
      Select Nvl(Min(����), 0), Max(��ֹʱ��)
      Into v_��ǰ����, d_End
      From Zldatamovelog
      Where ϵͳ = n_System And ��ת�� = 1;
    
      If v_��ǰ���� = 0 Then
        Update zlDataMove Set ״̬ = Null Where ϵͳ = n_System And ��� = 1;
        Return;
      End If;
    End If;
  
    --����Լ�������� 
    If Curtime_In = 1 Then
      Update Zldatamovelog Set ��ǰ���� = '���ڽ���Լ��������' Where ϵͳ = n_System And ���� = v_��ǰ����;
      --Ҫ�Ƚ���Լ��������������Ψһ�������������ú󣬻ᵼ�²�ѯ������������������������Ψһ�����ɾ����Ӧ������ 
      n_Reset := 1;
      Zl1_Datamove_Reb(n_System, Speedmode_In, 3, 0);
      If n_Peissystem > 0 Then
        Zl1_Datamove_Reb(n_Peissystem, Speedmode_In, 3, 0);
      End If;
      If n_Opersystem > 0 Then
        Zl1_Datamove_Reb(n_Opersystem, Speedmode_In, 3, 0);
      End If;
    
      Zl1_Datamove_Reb(n_System, Speedmode_In, 4, 0);
      If n_Peissystem > 0 Then
        Zl1_Datamove_Reb(n_Peissystem, Speedmode_In, 4, 0);
      End If;
      If n_Opersystem > 0 Then
        Zl1_Datamove_Reb(n_Opersystem, Speedmode_In, 4, 0);
      End If;
    End If;
  
    --����ת������ 
    ----------------------------------------------------------------------------------- 
    --�����»��ܱ����˷��û��ܣ�ҩƷ�շ����ܣ�ҩƷ��棬��Ա�ɿ����ȣ���Ȼֻ�Ǹ��µ��ڳ����� 
    --���ǣ����ڲ������ʱ����ǰ�����ݶ�ת���ˣ�����δ����ת������������δת���������º������ʱ���ѯ���ᷢ����������Щ���ڵ����ݷǳ�С������������� 
    --��ʹ����ĳЩ����ԭ����Ҫ���»��ܱ�Ҳ����ͨ�����ܱ���Ĺ��̽������»��ܣ����ԣ�������ת���������������¡� 
  
    --"��ǽ���ʱ��=ת����ʼʱ��"ʱ����¼ 
    If Optmode_In = 2 Then
      Update Zldatamovelog Set ת����ʼʱ�� = Sysdate Where ϵͳ = n_System And ���� = v_��ǰ����;
    End If;
  
    --a.ת����׼������
    For R In (Select ����, ��� From zlBakTables Where ϵͳ = n_System And ֱ��ת�� = 1 Order By ���, ���) Loop
      If Nvl(v_Pre���, -1) <> r.��� Then
        Update Zldatamovelog
        Set ��ǰ���� = '����ת����' || r.��� || '��(' || r.���� || '...)����'
        Where ϵͳ = n_System And ���� = v_��ǰ����;
        Commit;
      End If;
    
      Movedata(r.����, v_��ǰ����, Sysowner_In);
      v_Pre��� := r.���;
    End Loop;
  
    --b.ת���������
    v_Pre��� := -1;
    For R In (Select ����, ��� From zlBakTables Where ϵͳ = n_Peissystem And ֱ��ת�� = 1 Order By ���, ���) Loop
      If Nvl(v_Pre���, -1) <> r.��� Then
        Update Zldatamovelog
        Set ��ǰ���� = '����ת������' || r.��� || '��(' || r.���� || '...)����'
        Where ϵͳ = n_System And ���� = v_��ǰ����;
        Commit;
      End If;
    
      Movedata(r.����, v_��ǰ����, Peissysowner_In);
      v_Pre��� := r.���;
    End Loop;
  
    --c.ת����������
    v_Pre��� := -1;
    For R In (Select ����, ��� From zlBakTables Where ϵͳ = n_Opersystem And ֱ��ת�� = 1 Order By ���, ���) Loop
      If Nvl(v_Pre���, -1) <> r.��� Then
        Update Zldatamovelog
        Set ��ǰ���� = '����ת�������' || r.��� || '��(' || r.���� || '...)����'
        Where ϵͳ = n_System And ���� = v_��ǰ����;
        Commit;
      End If;
    
      Movedata(r.����, v_��ǰ����, Opersysowner_In);
      v_Pre��� := r.���;
    End Loop;
    Commit;
  
    Update ������ҳ Set ��ת�� = Null, ����ת�� = 1 Where ��ת�� = v_��ǰ����;
  
    Update zlDataMove Set �ϴ����� = d_End Where ϵͳ = n_System And ��� = 1;
  
    v_Sql := 'Update ' || Sysowner_In || '.zlBakInfo Set ���ת������ = Sysdate Where ϵͳ = ' || n_System;
    Execute Immediate v_Sql;
  
    If n_Peissystem > 0 Then
      v_Sql := 'Update ' || Peissysowner_In || '.zlBakInfo Set ���ת������ = Sysdate Where ϵͳ = ' || n_Peissystem;
      Execute Immediate v_Sql;
    End If;
  
    If n_Opersystem > 0 Then
      v_Sql := 'Update ' || Opersysowner_In || '.zlBakInfo Set ���ת������ = Sysdate Where ϵͳ = ' || n_Opersystem;
      Execute Immediate v_Sql;
    End If;
  
    Update Zldatamovelog
    Set ת������ʱ�� = Sysdate, ��ת�� = Null, ��ǰ���� = 'ת���������,�����ؽ���ת������'
    Where ϵͳ = n_System And ���� = v_��ǰ����;
    Commit;
  
    If Curtime_In = Totaltime_In Then
      Update zlDataMove
      Set ״̬ = Null, ������������ = Decode(Sign(d_End - ������������), -1, ������������, Null)
      Where ϵͳ = n_System And ��� = 1;
      Commit;
    End If;
  
    --4.�����ؽ���������´α��ת����ѯ���ٶȣ� 
    ----------------------------------------------------------------------------------- 
    --ÿ��ת���Ҫ�ؽ���ת�������������ؽ����׳��ֿ��������������޷��ؽ���ɾ����ORA-08104��
   
    --�������ת����ѯ�����������ɾ����Ŀ��пռ䣬�´α��ת��ʱ���ٷ�Χɨ������ݿ� 
    --���ÿ��ת�����У����ʱ�϶࣬���ԣ��ɸ��ݲ�ѯ�ĺ�ʱ����̬�����������(����ȱʡΪ24��ת�����ؽ�һ��) 
    Select Nvl(�ؽ��������, 0), Nvl(�ؽ�������Χ, 0)
    Into n_�ؽ��������, n_�ؽ�������Χ
    From zlDataMove
    Where ϵͳ = n_System And ��� = 1;
  
    If Mod(Curtime_In, n_�ؽ��������) = 0 And n_�ؽ�������� <> 0 Then
      Zl1_Datamove_Reb(n_System, Speedmode_In, 6, 1, Parallel_In, n_�ؽ�������Χ);
    
      If n_Peissystem > 0 Then
        Zl1_Datamove_Reb(n_Peissystem, Speedmode_In, 6, 1, Parallel_In, n_�ؽ�������Χ);
      End If;
      If n_Opersystem > 0 Then
        Zl1_Datamove_Reb(n_Opersystem, Speedmode_In, 6, 1, Parallel_In, n_�ؽ�������Χ);
      End If;
    End If;
  
    Update Zldatamovelog Set �ؽ�����ʱ�� = Sysdate, ��ǰ���� = '���' Where ϵͳ = n_System And ���� = v_��ǰ����;
    Commit;
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    --���ܲ������ݲ���ɹ���ɾ����������������Ψһ�����������ö�ʧ�� 
    Rollback;
    Update zlDataMove Set ״̬ = Null Where ϵͳ = n_System And ��� = 1;
  
    v_Err_Msg := Substr(SQLErrM, 1, 60);
    If Curtime_In = 1 And n_Reset = 0 Then
      Update Zldatamovelog Set ��ǰ���� = 'ת����ǳ���' || v_Err_Msg Where ϵͳ = n_System And ���� = v_��ǰ����;
    Else
      Update Zldatamovelog
      Set ��ǰ���� = 'ת������' || v_Err_Msg || Substr(v_Sql, 1, 30)
      Where ϵͳ = n_System And ���� = v_��ǰ����;
    End If;
    Commit;
  
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl1_Datamoveout1;
/

--109518:������,2017-05-24,���Ӳ������ݴ��������޸�
Create Or Replace Procedure Zl1_Datamove_Reb
(
  System_In    In Number,
  Speedmode_In In Number,
  Func_In      In Number,
  Enable_In    In Number := 0,
  Parallel_In  In Number := 0,
  Rebscope_In  In Number := 0
) As
  --���ܣ�����ʷ����ת��֮ǰ�����ô��������Զ���ҵ��Լ����������ת��֮��������Щ�����Լ��ؽ���ת���������ջر��ת�����������Ŀռ� 
  --������ 
  --System_In:    Ӧ��ϵͳ���,100=��׼�� 
  --speedmode_in������ת��ģʽ��0-����ģʽ��1-����ģʽ���ڿͻ���ͣ��ʱ��ת���ڼ����ת�����������Ψһ�������Լ�����������Լӿ���ת���ݵ�ɾ�������� 
  --func_in:      1=��������2=�Զ���ҵ��3=Լ����4=������5=�ؽ���ת��������6-�ջر��ת�����������Ŀռ䣬7-�����Ĵ洢�ռ䣨move�������ָ������õ�Լ�������� ,8-�ؽ����ת����ѯ��������������������� 
  --Enable_in:    0-���ã�1=���ã���func_inֵΪ1-4��Ч 
  --rebScope_in:   Func_In=6ʱ��ָ�ؽ������ķ�Χ(0-���ú�����,1-���ú����༰ҽ����,2-ȫ��)��Func_In=7ʱָMove��ķ�Χ(0-���ú����࣬1-ȫ��) 

  v_Sql      Varchar2(4000);
  n_Do       Number(1);
  n_Parallel Number(1);
  v_Tbs      Varchar2(100);
  v_Prompt   Varchar2(100);
  d_Curdate  Date;

  --���ܣ�1.���û���������ת�����������������,����ɾ�������¼ʱ���ӱ�ÿ�м�¼ִ��һ��SQL��ѯ��ɾ�� 
  --      2.���û�����������Ψһ��Լ��������ʱ���Զ�ɾ����Ӧ������������ʱ�Զ������������������ɾ������ 
  --���磺����ҽ������_FK_ҽ��ID�������Щ������ڵı�����δת����δ��zlbaktables���ж��壩��ִ��ǰ���鲢����ת���� 
  Procedure Setconstraintstatus As
    v_Pcol Varchar2(50);
    v_Fcol Varchar2(50);
    v_Del  Varchar2(4000);
  Begin
    --����ʱ���Ƚ�������ת��������������������ٽ���ת��������� 
    If Enable_In = 0 Then
      --1.����ģʽת��ʱ��������ҵ�����ɾ�����������ԣ����ڼ���ɾ����������ô�������������ӱ����ݵ�ɾ������
      If Speedmode_In = 0 Then
        For Rp In (Select Distinct a.Table_Name As Ptable_Name, a.Constraint_Name
                   From User_Constraints A, User_Constraints C, zlBakTables B
                   Where a.Table_Name = b.���� And b.ֱ��ת�� = 1 And b.ϵͳ = System_In And a.Constraint_Type In ('P', 'U') And
                         c.r_Constraint_Name = a.Constraint_Name And c.Constraint_Type = 'R' And
                         c.Delete_Rule = 'CASCADE'
                   Order By a.Table_Name) Loop
        
          Select f_List2str(Cast(Collect(Column_Name Order By Position) As t_Strlist))
          Into v_Pcol
          From User_Cons_Columns
          Where Constraint_Name = Rp.Constraint_Name;
        
          v_Del := '';
          For Rf In (Select b.Table_Name, b.Constraint_Name,
                            f_List2str(Cast(Collect(b.Column_Name Order By b.Position) As t_Strlist)) As r_Col
                     From User_Constraints A, User_Cons_Columns B
                     Where a.r_Constraint_Name = Rp.Constraint_Name And a.Constraint_Name = b.Constraint_Name
                     Group By b.Table_Name, b.Constraint_Name) Loop
            If Instr(v_Pcol, ',') > 0 Then
              v_Del := v_Del || Chr(10) || '        Delete ' || Rf.Table_Name || ' Where (' || Rf.r_Col ||
                       ') in ((:Old.' || Replace(v_Pcol, ',', ',:Old.') || '));';
            Else
              v_Del := v_Del || Chr(10) || '        Delete ' || Rf.Table_Name || ' Where ' || Rf.r_Col || ' = :Old.' ||
                       v_Pcol || ';';
            End If;
          End Loop;
        
          --�Լ���ɾ��������������������ڱ���ֶε����������������ֻ��ɾ������¼ʱ�ż���ɾ���Ӽ�¼
          --���Ҽ����������񣬷�������ora-2099,ora-04091����,��XX�����ݷ����˱仯���������������ܶ���
          Select Max(Column_Name)
          Into v_Fcol
          From User_Cons_Columns A, User_Constraints B
          Where a.Constraint_Name = b.Constraint_Name And b.r_Constraint_Name = b.Table_Name || '_PK' And
                b.r_Constraint_Name = Rp.Constraint_Name;
        
          If v_Fcol Is Not Null Then
            v_Sql := 'Create Or Replace Trigger ' || Rp.Ptable_Name || '_Cascade_Del' || Chr(10) || '    After Delete On ' ||
                     Rp.Ptable_Name || Chr(10) || '    For Each Row' || Chr(10) || 'Declare' || Chr(10) ||
                     ' Pragma Autonomous_Transaction;' || Chr(10) || 'Begin' || Chr(10) ||
                     '    If :Old.��ת�� Is Null And :Old.' || v_Fcol || ' Is Null Then ' || v_Del || Chr(10) || '    Commit;' ||
                     Chr(10) || '    End If; ' || Chr(10) || 'End ' || Rp.Ptable_Name || '_Cascade_Del;';
          Else
            v_Sql := 'Create Or Replace Trigger ' || Rp.Ptable_Name || '_Cascade_Del' || Chr(10) || '    After Delete On ' ||
                     Rp.Ptable_Name || Chr(10) || '    For Each Row' || Chr(10) || 'Begin' || Chr(10) ||
                     '    If :Old.��ת�� Is Null Then ' || v_Del || Chr(10) || '    End If; ' || Chr(10) || 'End ' ||
                     Rp.Ptable_Name || '_Cascade_Del;';
          End If;
        
          Execute Immediate v_Sql;
        End Loop;
      End If;
    
      --2.��������ת�����������������
      For R In (Select c.Table_Name, c.Constraint_Name, a.Table_Name As Ptable_Name
                From User_Constraints A, User_Constraints C, zlBakTables B
                Where a.Table_Name = b.���� And b.ֱ��ת�� = 1 And b.ϵͳ = System_In And a.Constraint_Type In ('P', 'U') And
                      c.r_Constraint_Name = a.Constraint_Name And c.Constraint_Type = 'R' And c.Status = 'ENABLED'
                Order By a.Table_Name) Loop
        v_Sql := 'Alter Table ' || r.Table_Name || ' Disable Constraint ' || r.Constraint_Name;
        Execute Immediate v_Sql;
      End Loop;
    
      --3.����������Ψһ������(����ת��ʱ)
      If Speedmode_In = 1 Then
        --����ɾ������������ʹskip_unusable_indexesΪtrue��Ҳ�޷�ɾ������Unusable״̬��Ψһ�������ı��еļ�¼
        --����ת������е�SQL��ѯ���������(������Ψһ����Ӧ������) 
        For R In (Select a.Table_Name, a.Constraint_Name
                  From User_Constraints A, zlBakTables T, User_Tables B
                  Where a.Table_Name = t.���� And t.ֱ��ת�� = 1 And t.ϵͳ = System_In And a.Status = 'ENABLED' And
                        a.Constraint_Type In ('P', 'U') And a.Table_Name = b.Table_Name And b.Iot_Type Is Null And
                        a.Constraint_Name Not In (Select Upper(������) From Zlbaktableindex Where ϵͳ = System_In)
                  Order By Constraint_Name) Loop
          v_Sql := 'Alter Table ' || r.Table_Name || ' Disable Constraint ' || r.Constraint_Name ||
                   ' Cascade Drop Index';
          Execute Immediate v_Sql;
        End Loop;
      End If;
    Else
      --����ʱ
      --1.������������Ψһ��������������ת����������������� 
      If Speedmode_In = 1 Then
        --���ؽ�������������Լ�����Ա��ؽ�����ʱ���ò���ִ������ʱ�䣬��������Լ��ʱҲ���Բ���novalidate��ʽ 
        For R In (Select d.Table_Name, d.Constraint_Name,
                         f_List2str(Cast(Collect(d.Column_Name Order By d.Position) As t_Strlist)) Colstr
                  From User_Cons_Columns D,
                       (Select a.Table_Name, a.Constraint_Name
                         From User_Constraints A, zlBakTables T
                         Where a.Table_Name = t.���� And t.ֱ��ת�� = 1 And t.ϵͳ = System_In And a.Status = 'DISABLED' And
                               a.Constraint_Type In ('P', 'U')) A
                  Where a.Constraint_Name = d.Constraint_Name And a.Table_Name = d.Table_Name
                  Group By d.Table_Name, d.Constraint_Name
                  Order By Constraint_Name) Loop
        
          Update zlDataMove Set ˵�� = '���ڻָ�Լ��:' || r.Constraint_Name Where ϵͳ = 100 And ��� = 1;
        
          Select Tablespace_Name Into v_Tbs From User_Indexes Where Table_Name = r.Table_Name And Rownum < 2;
        
          --����������Ψһ��ʱ�������Ǳ�ɾ���˵ģ���������Ҫ��Create 
          v_Sql := 'Create Unique Index ' || r.Constraint_Name || ' On ' || r.Table_Name || '(' || r.Colstr ||
                   ') Tablespace ' || v_Tbs || ' Nologging';
          Begin
            Execute Immediate v_Sql;
          Exception
            When Others Then
              Null; --������Щ������Ψһ�����Ǳ���ת���ڼ䱻���õģ�֮ǰ�ʹ��ڲ�Ψһ���ݣ�����Ψһ��������� 
          End;
        
          --���Զ�����Լ���������Ĺ��� 
          v_Sql := 'Alter Table ' || r.Table_Name || ' Enable Novalidate Constraint ' || r.Constraint_Name;
          Execute Immediate v_Sql;
        End Loop;
      End If;
    
      --2.��������ת����������������� 
      For R In (Select c.Table_Name, c.Constraint_Name
                From User_Constraints A, User_Constraints C, zlBakTables B
                Where a.Table_Name = b.���� And b.ֱ��ת�� = 1 And b.ϵͳ = System_In And a.Constraint_Type In ('P', 'U') And
                      c.r_Constraint_Name = a.Constraint_Name And c.Constraint_Type = 'R' And c.Status = 'DISABLED'
                Order By a.Table_Name) Loop
        --Ϊ�˼ӿ��ٶȣ�����novalidate������֤�������� 
        --��������ת����������������zlbaktables�ж����ˣ���û�б�д��Ӧ������ת���ű���δ��֤�����ݿ�����Υ��Լ��������� 
        v_Sql := 'Alter Table ' || r.Table_Name || ' Enable Novalidate Constraint ' || r.Constraint_Name;
        Execute Immediate v_Sql;
      End Loop;
    
      --3.����ģʽת��ʱ��ɾ��֮ǰ�����������������ɾ������Ĵ�����
      If Speedmode_In = 0 Then
        For R In (Select a.Trigger_Name
                  From User_Triggers A, zlBakTables B
                  Where a.Table_Name = b.���� And b.ֱ��ת�� = 1 And b.ϵͳ = System_In And
                        Trigger_Name = Table_Name || '_CASCADE_DEL' And Triggering_Event = 'DELETE') Loop
          v_Sql := 'Drop Trigger ' || r.Trigger_Name;
          Execute Immediate v_Sql;
        End Loop;
      End If;
    End If;
  End Setconstraintstatus;

  --���ܣ�����ģʽʱ����LOB�������������������ģʽʱ������ת�������÷�ת������������(���磺����ҽ���Ƽ�_IX_�շ�ϸĿID) 
  --˵��������������Ϊ�����ɾ�����ݵ����� 
  Procedure Setindexstatus As
  Begin
    If Speedmode_In = 1 Then
      --����ת������е�SQL��ѯ��������� 
      For R In (Select /*+ rule*/
                 a.Index_Name
                From User_Indexes A, zlBakTables T
                Where a.Table_Name = t.���� And t.ֱ��ת�� = 1 And t.ϵͳ = System_In And t.ֱ��ת�� = 1 And
                      a.Index_Name <> a.Table_Name || '_IX_��ת��' And
                      a.Index_Name Not In (Select Upper(������) From Zlbaktableindex Where ϵͳ = System_In) And
                      a.Status = Decode(Enable_In, 0, 'VALID', 'UNUSABLE') And a.Index_Type = 'NORMAL' And Not Exists
                 (Select 1
                       From User_Constraints C
                       Where c.Index_Name = a.Index_Name And c.Constraint_Type In ('P', 'U'))
                Order By Index_Name) Loop
      
        If Enable_In = 0 Then
          v_Sql := 'Alter Index ' || r.Index_Name || ' Unusable';
          Execute Immediate v_Sql;
        Else
          Update zlDataMove Set ˵�� = '�����ؽ�����:' || r.Index_Name Where ϵͳ = 100 And ��� = 1;
        
          v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Nologging';
          Begin
            Execute Immediate v_Sql;
            --�����ؽ��Ƚ������������ؽ�����Ҫ���������������������������ORA-00054: ��Դ��æ, ��ָ���� NOWAIT ��ʽ��ȡ��Դ 
          
          Exception
            When Others Then
              If SQLErrM Like 'ORA-00054%' Then
                v_Sql := Replace(v_Sql, 'Rebuild', 'Rebuild Online');
                Execute Immediate v_Sql;
              End If;
          End;
        End If;
      End Loop;
    Else
      For R In (Select a.Index_Name
                From (Select d.Table_Name, d.Index_Name,
                              f_List2str(Cast(Collect(d.Column_Name Order By d.Column_Position) As t_Strlist)) Colstr
                       From User_Ind_Columns D, zlBakTables T, User_Indexes C
                       Where c.Table_Name = t.���� And t.ֱ��ת�� = 1 And t.ϵͳ = System_In And c.Uniqueness = 'NONUNIQUE' And
                             c.Index_Type = 'NORMAL' And c.Status = Decode(Enable_In, 0, 'VALID', 'UNUSABLE') And
                             c.Index_Name = d.Index_Name And c.Table_Name = d.Table_Name
                       Group By d.Table_Name, d.Index_Name) A,
                     (Select e.Table_Name,
                              f_List2str(Cast(Collect(e.Column_Name Order By e.Position) As t_Strlist)) Colstr
                       From User_Cons_Columns E, User_Constraints F, zlBakTables T, User_Constraints C
                       Where e.Table_Name = t.���� And t.ֱ��ת�� = 1 And t.ϵͳ = System_In And
                             e.Constraint_Name = f.Constraint_Name And f.Constraint_Type = 'R' And
                             c.Constraint_Name = f.r_Constraint_Name And c.Table_Name Not In ('������ҳ', '������Ϣ') And
                             Not Exists
                        (Select 1 From zlBakTables G Where g.���� = c.Table_Name And g.ϵͳ = System_In)
                       Group By e.Table_Name, e.Constraint_Name) B
                Where a.Table_Name = b.Table_Name And a.Colstr = b.Colstr
                Order By Index_Name) Loop
      
        If Enable_In = 0 Then
          --���⴦�������������������ã�������ҩƷĿ¼�޸Ĺ�񣬲���ɿ���Ҫʹ�� 
          If r.Index_Name Not In ('����ҽ����¼_IX_�շ�ϸĿID', 'ҩƷ�շ���¼_IX_ҩƷID', 'ҩƷ�շ���¼_IX_�۸�ID') Then
            v_Sql := 'Alter Index ' || r.Index_Name || ' Unusable';
            Execute Immediate v_Sql;
          End If;
        Else
          Update zlDataMove Set ˵�� = '�����ؽ�����:' || r.Index_Name Where ϵͳ = 100 And ��� = 1;
        
          v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Online Nologging';
          Execute Immediate v_Sql;
          --�����ؽ��Ƚ������������ؽ�����Ҫ���������������������������ORA-00054: ��Դ��æ, ��ָ���� NOWAIT ��ʽ��ȡ��Դ 
        End If;
      End Loop;
    End If;
  End Setindexstatus;

  --���ܣ�ת�������ڼ䣬ͣ��ת�����ϵ����д�������ת�����ٻָ� 
  Procedure Settriggerstatus As
  Begin
    For R In (Select Distinct a.Table_Name, t.ͣ�ô�����
              From User_Triggers A, zlBakTables T
              Where a.Status = Decode(Enable_In, 0, 'ENABLED', 'DISABLED') And a.Table_Name = t.���� And t.ֱ��ת�� = 1 And
                    t.ϵͳ = System_In) Loop
      If Enable_In = 0 Then
        v_Sql := 'Alter Table ' || r.Table_Name || ' DISABLE ALL TRIGGERS';
        Update zlBakTables Set ͣ�ô����� = 1 Where ϵͳ = System_In And ���� = r.Table_Name;
      Elsif Nvl(r.ͣ�ô�����, 0) = 1 Then
        v_Sql := 'Alter Table ' || r.Table_Name || ' ENABLE ALL TRIGGERS';
        Update zlBakTables Set ͣ�ô����� = Null Where ϵͳ = System_In And ���� = r.Table_Name;
      End If;
      Execute Immediate v_Sql;
    End Loop;
    Commit;
  End Settriggerstatus;

  --���ܣ�ת�������ڼ䣬ͣ�õ�ǰ�����ߵ������Զ���ҵ��ת���������� 
  Procedure Setjobstatus As
    v_Jobs Varchar2(4000);
  Begin
    --ͣ�� 
    If Enable_In = 0 Then
      For R In (Select Job From User_Jobs Where Broken = 'N') Loop
        Dbms_Job.Broken(r.Job, True);
        v_Jobs := v_Jobs || ',' || r.Job;
      End Loop;
    
      If v_Jobs Is Not Null Then
        v_Jobs := Substr(v_Jobs, 2);
        Update zlDataMove Set ͣ����ҵ�� = v_Jobs Where ϵͳ = System_In And ��� = 1;
      End If;
    Else
      --���� 
      Select ͣ����ҵ�� Into v_Jobs From zlDataMove Where ϵͳ = System_In And ��� = 1;
      If v_Jobs Is Not Null Then
        For R In (Select Job
                  From User_Jobs
                  Where Broken = 'Y' And Job In (Select Column_Value From Table(f_Num2list(v_Jobs)))) Loop
          Dbms_Job.Broken(r.Job, False);
        End Loop;
        Update zlDataMove Set ͣ����ҵ�� = Null Where ϵͳ = System_In And ��� = 1;
      End If;
    End If;
    --��ҵ���ú�����ύ�������Ч 
    Commit;
  End Setjobstatus;
Begin
  If Parallel_In < 2 Then
    Execute Immediate 'Alter Session DISABLE PARALLEL DDL';
  Else
    If Func_In In (6, 7, 8) Or Func_In In (3, 4) And Enable_In = 1 Then
      --Ϊ�ؽ��������ò���ִ�У�����ͨ��������IO�豸�����ܣ�����̫�ߵĲ��жȷ����ή�����ܣ����и����ܴ洢�豸���ɼӴ��жȣ� 
      --ִ���ؽ���������Զ�Ϊ�������ϲ��ж����ԣ������ȡ������Ӱ�����SQL��ִ�мƻ�(ȫ��ɨ��+���в�ѯ������),�ں���ȡ�������Ĳ��ж� 
      --�ָ����߿��Լ��������ʱ�������ǲ�������ģʽ�������ϲ��У�����̫��
      Execute Immediate 'Alter Session FORCE PARALLEL DDL PARALLEL ' || Parallel_In;
      n_Parallel := 1;
    End If;
  End If;

  --������������ٶȣ�����40%���ϵ�ʱ�䣩
  If Func_In In (6, 7, 8) Or Func_In In (3, 4) And Enable_In = 1 Then
    Execute Immediate 'alter session set workarea_size_policy=MANUAL';
  
    --����ֱ��·��IO�Ĵ�С
    Execute Immediate 'alter session set events ''10351 trace name context forever, level 128''';
    Execute Immediate 'alter session SET db_file_multiblock_read_count=128';
    Execute Immediate 'alter session set "_sort_multiblock_read_count"=128';
    Begin
      --����10G��BUG��sort_area_size��ִ�����βŻ���Ч
      Execute Immediate 'alter session SET sort_area_size=512000000';
      Execute Immediate 'alter session SET sort_area_size=512000000';
    Exception
      When Others Then
        Null; --��������ڴ治��500M��ʧ�������
    End;
    Execute Immediate 'alter session SET db_block_checking=false';
  End If;

  If Func_In In (5, 6, 7, 8) Or Func_In In (3, 4) And Enable_In = 1 Then
    d_Curdate := Sysdate;
  End If;

  If Func_In = 1 Then
    --1.���ô����� 
    Settriggerstatus;
  Elsif Func_In = 2 Then
    --2.�����Զ���ҵ 
    Setjobstatus;
  Elsif Func_In = 3 Then
    --3.����Լ��״̬    
    Setconstraintstatus;
    v_Prompt := '�ָ����õ�Լ��';
  Elsif Func_In = 4 Then
    --4.��������״̬ 
    Setindexstatus;
    v_Prompt := '�ָ����õ�����';
  Elsif Func_In = 5 Then
    --5.�ؽ�"��ת��"����    
    For R In (Select Index_Name
              From (Select b.Index_Name
                     From zlBakTables A, User_Indexes B
                     Where a.���� = b.Table_Name And a.ֱ��ת�� = 1 And a.ϵͳ = System_In And
                           b.Index_Name = b.Table_Name || '_IX_��ת��'
                     Union All
                     Select '������ҳ_IX_��ת��'
                     From Dual
                     Where System_In = 100)
              Order By 1) Loop
      Update zlDataMove Set ˵�� = '�����ؽ���ת������:' || r.Index_Name Where ϵͳ = 100 And ��� = 1;
    
      --��ʱ̫�̣����벢��DDL 
      --����ת��ʱ����ؽ����������������������������������ORA-00054: ��Դ��æ, ��ָ���� NOWAIT ��ʽ��ȡ��Դ 
      --�����ؽ�����̫�������ԣ���ʹ����ת��ģʽҲ���������ؽ�
      v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Nologging';
      Begin
        Execute Immediate v_Sql;
      Exception
        When Others Then
          If SQLErrM Like 'ORA-00054%' Then
            v_Sql := Replace(v_Sql, 'Rebuild', 'Rebuild Online');
            Execute Immediate v_Sql;
          End If;
      End;
    End Loop;
    v_Prompt := '�ؽ���ת������';
  
  Elsif Func_In = 6 Then
    --6.�ؽ����ת����ѯ���õ������������Ա����ؽ�����������һ��Ĳ�ѯʱ�䣩 
    --����ҵ������ý׶��������ؽ���Щ�������Ա���һЩ����Ҫ���ؽ���ʱ    
    For R In (Select b.Index_Name, a.���
              From User_Indexes B, zlBakTables A
              Where a.���� = b.Table_Name And a.ϵͳ = System_In And
                    (b.Table_Name, b.Index_Name) In
                    (Select ����, Upper(������) From Zlbaktableindex Where ϵͳ = System_In)
              Order By Index_Name) Loop
      n_Do := 0;
      If Rebscope_In = 0 Then
        If r.��� < 5 Then
          n_Do := 1; --�����ú����� 
        End If;
      Elsif Rebscope_In = 1 Then
        If r.��� < 5 Or r.��� = 8 Then
          n_Do := 1; --�����ú����ࡢҽ���� 
        End If;
      Else
        n_Do := 1;
      End If;
    
      If n_Do = 1 Then
        Update zlDataMove Set ˵�� = '�����ؽ�����:' || r.Index_Name Where ϵͳ = 100 And ��� = 1;
      
        --v_Sql := 'Alter Index ' || r.Index_Name || ' shrink Space'; 
        --ʹ��shrink��ʽ���ܲ���ִ��,��������ٶȱ�rebuild PARALLEL 8 ��6�� 
        If Speedmode_In = 1 Then
          v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Nologging';
        Else
          v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Online Nologging';
        End If;
        Begin
          Execute Immediate v_Sql;
          --�����ؽ��Ƚ������������ؽ�����Ҫ���������������������������ORA-00054: ��Դ��æ, ��ָ���� NOWAIT ��ʽ��ȡ��Դ 
        
        Exception
          When Others Then
            If SQLErrM Like 'ORA-00054%' Then
              v_Sql := Replace(v_Sql, 'Rebuild', 'Rebuild Online');
              Execute Immediate v_Sql;
            End If;
        End;
      End If;
    End Loop;
    v_Prompt := '�ؽ����ת����������';
  
    --����������
  Elsif Func_In = 7 Then
    --rebScope_in=0,ֻ�������С��5�ľ��ú���������á�ҩƷ��Ʊ�ݣ�������ȫ������     
    For R In (Select a.���� As Table_Name
              From zlBakTables A
              Where a.ֱ��ת�� = 1 And a.ϵͳ = System_In And (��� < Decode(Rebscope_In, 0, 5, 100))
              Order By ���, ���) Loop
    
      Update zlDataMove Set ˵�� = '���������:' || r.Table_Name Where ϵͳ = 100 And ��� = 1;
    
      --����п��еĿռ䣬����Ƶ�������ռ䣬ֻ���������ܾ����ƶ��ļ�β�������ݿ飬�Ա���б�ռ��ļ������� 
      --��ǰ�������˻Ự����ǿ�Ʋ��� 
      v_Sql := 'Alter Table ' || r.Table_Name || ' Move Nologging';
      Execute Immediate v_Sql;
    
      --�����ƶ�Lob���� 
      For L In (Select Column_Name, Tablespace_Name From User_Lobs Where Table_Name = r.Table_Name) Loop
        v_Sql := 'Alter Table ' || r.Table_Name || ' Move Lob(' || l.Column_Name || ') Store as (Tablespace ' ||
                 l.Tablespace_Name || ') Nologging';
        Execute Immediate v_Sql;
      End Loop;
    
      v_Sql := 'Alter Table ' || r.Table_Name || ' Noparallel';
      Execute Immediate v_Sql;
    
      --move�󣬱���ص�������ȫ��ʧЧ����Ҫȫ���ؽ� 
      For S In (Select Index_Name
                From User_Indexes
                Where Table_Name = r.Table_Name And Status = 'UNUSABLE'
                Order By Index_Name) Loop
      
        Update zlDataMove Set ˵�� = '���ڻָ�ʧЧ����:' || s.Index_Name Where ϵͳ = 100 And ��� = 1;
      
        --��ǰ�������˻Ự����ǿ�Ʋ��� 
        v_Sql := 'Alter Index ' || s.Index_Name || ' Rebuild Nologging';
        Execute Immediate v_Sql;
      End Loop;
    End Loop;
    v_Prompt := '���������';
  
    --�ؽ�ת�����ϱ��ת���������������������ת����ɺ��ջؿ��пռ䣩
    --ʧЧ���������ؽ�����Ϊת������е������ؽ�����
  Elsif Func_In = 8 Then
    For R In (Select b.Index_Name, a.���
              From User_Indexes B, zlBakTables A
              Where a.���� = b.Table_Name And a.ϵͳ = System_In And b.Status = 'VALID' And b.Index_Type = 'NORMAL' And
                    b.Index_Name Not Like 'BIN$%' And
                    b.Index_Name Not In (Select Upper(������) From Zlbaktableindex Where ϵͳ = System_In)
              Order By Index_Name) Loop
    
      Update zlDataMove Set ˵�� = '�����ؽ�����:' || r.Index_Name Where ϵͳ = 100 And ��� = 1;
    
      If Speedmode_In = 1 Then
        v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Nologging';
      Else
        v_Sql := 'Alter Index ' || r.Index_Name || ' Rebuild Online Nologging';
      End If;
      Begin
        Execute Immediate v_Sql;
        --�����ؽ��Ƚ������������ؽ�����Ҫ���������������������������ORA-00054: ��Դ��æ, ��ָ���� NOWAIT ��ʽ��ȡ��Դ    
      Exception
        When Others Then
          If SQLErrM Like 'ORA-00054%' Then
            v_Sql := Replace(v_Sql, 'Rebuild', 'Rebuild Online');
            Execute Immediate v_Sql;
          End If;
      End;
    End Loop;
    v_Prompt := '�ؽ����ת���������������';
  End If;

  If Func_In In (5, 6, 7, 8) Or Func_In In (3, 4) And Enable_In = 1 Then
    Update zlDataMove
    Set ˵�� = To_Char(Sysdate, 'mm-dd hh24:mi') || v_Prompt || ':' || Trunc((Sysdate - d_Curdate) * 24 * 60) || '����'
    Where ϵͳ = 100 And ��� = 1;
  End If;

  --ִ���ؽ���������Զ�Ϊ�������ϲ��ж����ԣ������ȡ������Ӱ�����SQL��ִ�мƻ�(ȫ��ɨ��+���в�ѯ������) 
  --------------------------------------------------------------------------------------------------- 
  If n_Parallel = 1 Then
    Execute Immediate 'ALTER Session DISABLE PARALLEL DDL';
  
    For R In (Select Index_Name From User_Indexes Where Degree Not In ('1', '0')) Loop
      v_Sql := 'Alter Index ' || r.Index_Name || ' Noparallel';
      Execute Immediate v_Sql;
    End Loop;
  End If;

  Commit;
  --�����̲����д����������ɵ��ù��̴��� 
End Zl1_Datamove_Reb;
/


--108964:��С��,2017-05-24,��ͬ����βο�ֵ��ȡ
CREATE OR REPLACE Function Zl_Get_Reference
(
  Type_In       In Number, --0=�ο� 1=�ο�ID 2=Σ���ο� 3=Σ���ο����� 4=Σ���ο�����
  ��Ŀid_In     In Number,
  �걾����_In   In Varchar2,
  �Ա�_In       In Number,
  ��������_In   In Date,
  ����id_In     In Number := Null,
  ����_In       In Varchar2 := Null,
  �������id_In In Number := Null
) Return Varchar2 As

  Cursor v_Reference_Type Is
    Select a.Id,
           Trim(To_Char(a.�ο���ֵ, c.��ʽ)) || '��' || Trim(To_Char(a.�ο���ֵ, c.��ʽ)) ||
            Decode(a.�ٴ�����, Null, '', '����', '', 'Ӥ��', '', ' ' || a.�ٴ�����) As ����ο�, b.�������, b.ȡֵ����,
           Trim(To_Char(a.��ʾ����, c.��ʽ)) || '��' || Trim(To_Char(a.��ʾ����, c.��ʽ)) ||
            Decode(a.�ٴ�����, Null, '', '����', '', 'Ӥ��', '', ' ' || a.�ٴ�����) As Σ���ο�, a.��ʾ����, a.��ʾ����, Nvl(b.��ο�, 0) ��ο�
    From ������Ŀ�ο� A, ������Ŀ B,
         (Select '9999990' ||
                   Decode(Max(Nvl(c.С��λ��, -1)), 0, '', -1, '.00', Substr('.000000', 1, 1 + Max(Nvl(c.С��λ��, -1)))) As ��ʽ



           From ����������Ŀ C, ������Ŀ D
           Where d.������Ŀid = ��Ŀid_In And d.������Ŀid = c.��Ŀid(+)) C
    Where a.��Ŀid = ��Ŀid_In And a.��Ŀid = b.������Ŀid;

  v_Return Varchar2(4000);
  v_Sql    Varchar2(4000);

  Type c_Type Is Ref Cursor; --����REF�α�����
  r_Emp v_Reference_Type%RowType; --����һ�������ͱ���
  Cur   c_Type; --����REF�α����͵ı���

  v_������� Number(1);

  v_����     Number(18, 1);
  v_����     Number(18, 1);
  v_����     Number(18, 1);
  v_Сʱ     Number(18, 1);
  v_�������� Date;
  v_Pos      Number(4);
  v_��ο�   Number(4);
  v_Value    Number(18);
  v_Valuerec Varchar2(255);
  v_����     Varchar2(50);
  v_����ο� Varchar2(1000);
  v_�ο�id   Number(18);
  v_Σ���ο� Varchar2(1000);
  v_��ʾ���� Varchar2(1000);
  v_��ʾ���� Varchar2(1000);
  d_Sysdate  Date;

  v_��Ŀid_Bound   ������Ŀ�ο�.��Ŀid%Type;
  v_�걾����_Bound ������Ŀ�ο�.�걾����%Type;
  v_�Ա���1_Bound  ������Ŀ�ο�.�Ա���%Type;
  v_�Ա���2_Bound  ������Ŀ�ο�.�Ա���%Type;
  v_�Ա���3_Bound  ������Ŀ�ο�.�Ա���%Type;
  v_����id_Bound   ������Ŀ�ο�.����id%Type;

  v_���䵥λ��_Bound   ������Ŀ�ο�.���䵥λ%Type;
  v_���䵥λ��_Bound   ������Ŀ�ο�.���䵥λ%Type;
  v_���䵥λСʱ_Bound ������Ŀ�ο�.���䵥λ%Type;
  v_���䵥λ��_Bound   ������Ŀ�ο�.���䵥λ%Type;

  v_���䵥λ��1_Bound   ������Ŀ�ο�.���䵥λ%Type;
  v_���䵥λ��1_Bound   ������Ŀ�ο�.���䵥λ%Type;
  v_���䵥λСʱ1_Bound ������Ŀ�ο�.���䵥λ%Type;
  v_���䵥λ��1_Bound   ������Ŀ�ο�.���䵥λ%Type;

  v_�ٴ�����_Bound   ������Ŀ�ο�.�ٴ�����%Type;
  v_�������id_Bound ������Ŀ�ο�.�������id%Type;
  v_����_1           Varchar2(50);
  v_����_2           Varchar2(50);

  Function Sub_Is_Number(v_In In Varchar2) Return Boolean Is
    n_Tmp Number;
  Begin
    n_Tmp := To_Number(v_In);
    If n_Tmp Is Not Null Then
      Return True;
    Else
      Return False;
    End If;
  Exception
    When Others Then
      Return False;
  End Sub_Is_Number;

  Function Zlsplit
  (
    v_Str       In Varchar2,
    v_Delimiter In Varchar2,
    v_Number    In Number
  ) Return Varchar2 Is
    v_Record     Varchar2(1000);
    v_Currrecord Varchar2(1000);
    v_Currnum    Number;
  Begin
    v_Record  := v_Str || v_Delimiter;
    v_Currnum := 0;
    While v_Record Is Not Null Loop
      v_Currrecord := Substr(v_Record, 1, Instr(v_Record, v_Delimiter) - 1);
      If v_Currnum = v_Number Then
        Return(v_Currrecord);
      End If;

      v_Currnum := v_Currnum + 1;
      v_Record  := Replace(v_Delimiter || v_Record, v_Delimiter || v_Currrecord || v_Delimiter);
    End Loop;

    Return('');
  End Zlsplit;
  Function Zlval(Vstr In Varchar2) Return Number Is
    Result Number(16, 6);
    Intbit Number(8);
    Strnum Varchar(10);
  Begin
    Strnum := '';
    For Intbit In 1 .. 10 Loop
      If Instr('0123456789.', Substr(Vstr, Intbit, 1)) = 0 Then
        Exit;
      End If;
      Strnum := Strnum || Substr(Vstr, Intbit, 1);
      Null;
    End Loop;
    Result := To_Number(Strnum);
    Return(Result);
  End Zlval;

Begin
  d_Sysdate := Sysdate;

  v_Sql := ' Select a.id,Trim(To_Char(A.�ο���ֵ, C.��ʽ)) || ''��'' || Trim(To_Char(A.�ο���ֵ, C.��ʽ)) || ' ||
           ' Decode(A.�ٴ�����, Null, '''', ''����'', '''', ''Ӥ��'','''', '' '' || A.�ٴ�����) As ����ο�, B.�������, B.ȡֵ����, ' ||
           ' Trim(To_Char(A.��ʾ����, C.��ʽ)) || ''��'' || Trim(To_Char(A.��ʾ����, C.��ʽ)) || ' || ' Decode(A.�ٴ�����, Null, '''', ''����'', '''', ''Ӥ��'','''', '' '' || A.�ٴ�����) As Σ���ο�,a.��ʾ����,a.��ʾ����,
             nvl(b.��ο�,0) ��ο� ' || ' From ������Ŀ�ο� A, ������Ŀ B, ' || ' (Select ''9999990'' || ' ||
           ' Decode(Max(Nvl(C.С��λ��, -1)), 0, '''', -1, ''.00'', Substr(''.000000'', 1, 1 + Max(Nvl(C.С��λ��, -1)))) As ��ʽ ' ||
           ' From ����������Ŀ C, ������Ŀ D ' || ' Where D.������ĿID = :��ĿID And D.������ĿID = C.��ĿID(+)) C ' ||
           ' Where A.��ĿID = :��ĿID And A.��ĿID = B.������ĿID ';

  v_��Ŀid_Bound := ��Ŀid_In;

  v_���� := ����_In;
  If v_���� = '��' Then
    v_���� := Null;
  End If;

  If v_���� = '��' Then
    v_���� := Null;
  End If;

  If v_���� = 'Сʱ' Then
    v_���� := Null;
  End If;

  If v_���� = '��' Then
    v_���� := Null;
  End If;

  If Nvl(�걾����_In, '') <> '' Or �걾����_In Is Not Null Then
    v_Sql := v_Sql || ' And A.�걾���� = :�걾���� ';
  Else
    v_Sql := v_Sql || ' And (A.�걾���� = :�걾���� or 1=1) ';
  End If;
  v_�걾����_Bound := �걾����_In;

  If Nvl(�Ա�_In, '') <> '' Or �Ա�_In Is Not Null Then
    --V_Sql := V_Sql || ' And A.�Ա��� = Nvl(' || �Ա�_In || ', 1) ';
    v_Sql := v_Sql || ' And decode(A.�Ա���,null,:�Ա�,0,:�Ա�,A.�Ա���) = Nvl(:�Ա�, 1) ';

  Else
    v_Sql := v_Sql || ' And (decode(A.�Ա���,null,:�Ա�1,0,:�Ա�2,A.�Ա���) = Nvl(:�Ա�3, 1) or 1 = 1) ';
  End If;
  v_�Ա���1_Bound := �Ա�_In;
  v_�Ա���2_Bound := �Ա�_In;
  v_�Ա���3_Bound := �Ա�_In;

  If Nvl(����id_In, '') <> '' Or ����id_In Is Not Null Then
    v_Sql := v_Sql || ' And (A.����id = :����ID Or A.����id Is Null) ';
  Else
    v_Sql := v_Sql || ' And (A.����id = :����ID Or A.����id Is Null or 1=1) ';
  End If;
  v_����id_Bound := ����id_In;

  If Nvl(v_����, '') <> '' Or v_���� Is Not Null Then
    If Instr(v_����, '��') > 0 Or Instr(v_����, '��') > 0 Or Instr(v_����, '��') > 0 Or Instr(v_����, 'Сʱ') > 0 Or
       Sub_Is_Number(v_����) Then
      --��������
      v_�������� := ��������_In;
      v_����_1   := v_����;
      If Instr(v_����_1, '��') > 0 Then
        v_����   := Substr(v_����_1, 1, Instr(v_����_1, '��'));
        v_����_2 := Substr(v_����_1, Instr(v_����_1, '��') + 1);
      Elsif Instr(v_����, '��') > 0 Then
        v_����   := Substr(v_����_1, 1, Instr(v_����_1, '��'));
        v_����_2 := Substr(v_����_1, Instr(v_����_1, '��') + 1);
      Elsif Instr(v_����, '��') > 0 Then
        v_����   := Substr(v_����_1, 1, Instr(v_����_1, '��'));
        v_����_2 := Substr(v_����_1, Instr(v_����_1, '��') + 1);
      Elsif Instr(v_����, 'Сʱ') > 0 Then
        v_����   := Substr(v_����_1, 1, Instr(v_����_1, 'Сʱ') + 1);
        v_����_2 := Substr(v_����_1, Instr(v_����_1, 'Сʱ') + 2);
        If v_���� = '0Сʱ' Or v_���� = '0ʱ' Then
          v_���� := ' ';
        End If;
      End If;
      If v_���� Is Not Null And (v_���� = '����' Or v_���� = 'Ӥ��' Or v_���� = '��') = False Then
        If Substr(v_����, 1, 1) = '*' Then
          v_�������� := Add_Months(d_Sysdate, -216);
        Else
          If Substr(v_����, Length(v_����)) = '��' Then
            v_�������� := Add_Months(d_Sysdate, -1 * Nvl(Zlval(v_����), 0));
          Else
            If Substr(v_����, Length(v_����)) = '��' Then
              v_�������� := d_Sysdate - Nvl(Zlval(v_����), 0);
            Else
              If Substr(v_����, Length(v_����) - 1) = 'Сʱ' Then
                If Nvl(Zlval(v_����), 0) <> 0 Then
                  v_�������� := d_Sysdate - Nvl(Zlval(v_����), 0) / 24;
                End If;
              Else
                v_�������� := Add_Months(d_Sysdate, -12 * Nvl(Zlval(v_����), 0)) - 1;
              End If;
            End If;
          End If;
          If v_����_2 Is Not Null Then
            If Substr(v_����_2, Length(v_����_2)) = '��' Then
              v_�������� := Add_Months(v_��������, -1 * Nvl(Zlval(v_����_2), 0));
            Else
              If Substr(v_����_2, Length(v_����_2)) = '��' Then
                v_�������� := v_�������� - Nvl(Zlval(v_����_2), 0);
              Else
                If Substr(v_����_2, Length(v_����_2) - 1) = 'Сʱ' Then
                  If Nvl(Zlval(v_����_2), 0) <> 0 Then
                    v_�������� := v_�������� - Nvl(Zlval(v_����_2), 0) / 24;
                  End If;
                End If;
              End If;
            End If;
          End If;
        End If;
      End If;
      If Not (v_�������� Is Null) Then
        v_���� := Round(Months_Between(d_Sysdate, v_��������) / 12 ,1);
        v_���� := Round(Months_Between(d_Sysdate, v_��������) ,1);
        v_���� := Round(d_Sysdate - v_�������� ,1);
        v_Сʱ := Round((d_Sysdate - (v_�������� - 1 / 24)) * 24 - 1);
      End If;
      v_Sql := v_Sql || 'And (Decode(A.���䵥λ, ''��'',:��, ''��'',:��,''Сʱ'',:Сʱ,:��) ' ||
               ' Between Nvl(A.��������, -9999) And Nvl(A.��������, 9999) )';
    Else
      v_Sql := v_Sql || 'And (Decode(A.���䵥λ, ''��'',:��, ''��'',:��,''Сʱ'',:Сʱ,:��) ' ||
               ' Between Nvl(A.��������, -9999) And Nvl(A.��������, 9999) or 1=1 )';
    End If;

  Else
    v_Sql := v_Sql || 'And (Decode(A.���䵥λ, ''��'',:��, ''��'',:��,''Сʱ'',:Сʱ,:��) ' ||
             ' Between Nvl(A.��������, -9999) And Nvl(A.��������, 9999) or 1=1 )';
  End If;
  v_���䵥λ��_Bound   := v_����;
  v_���䵥λ��_Bound   := v_����;
  v_���䵥λСʱ_Bound := v_Сʱ;
  v_���䵥λ��_Bound   := v_����;
  If Instr(v_����, '����') > 0 Or Instr(v_����, 'Ӥ��') > 0 Or Instr(v_����, '����') > 0 Then
    --������˺�Ӥ��
    v_Sql := v_Sql || ' And A.�ٴ����� =:����';
  Else
    v_Sql := v_Sql || ' And (A.�ٴ����� =:���� or 1=1)';
    v_Sql := v_Sql || ' And instr(''Ӥ��,����'',nvl(�ٴ�����,'' '')) <= 0  ';
  End If;

  v_�ٴ�����_Bound := Replace(v_����, '����', 'Ӥ��');

  If Nvl(�������id_In, '') <> '' Or �������id_In Is Not Null Then
    v_Sql := v_Sql || ' And (A.�������ID = :�������ID Or nvl(A.�������ID,0) = 0) ';
  Else
    v_Sql := v_Sql || ' And ((A.�������ID = :�������ID Or nvl(A.�������ID,0) = 0) or 1=1) ';
  End If;
  v_�������id_Bound := �������id_In;

  If (Nvl(v_����, '') = '' Or v_���� Is Null) And (��������_In <> '' Or ��������_In Is Not Null) Then
    --���������ڲ�ѯ
    If Not (��������_In Is Null) Then
      v_���� := Round(Months_Between(d_Sysdate, ��������_In) / 12 - 0.5);
      v_���� := Round(Months_Between(d_Sysdate, ��������_In) - 0.5);
      v_���� := Round(d_Sysdate - ��������_In - 0.5);
      v_Сʱ := Round((d_Sysdate - (��������_In - 1 / 24)) * 24 - 1);

      v_Sql := v_Sql || 'And (Decode(A.���䵥λ, ''��'',:��, ''��'',:��,''Сʱ'',:Сʱ,:��) ' ||
               ' Between Nvl(A.��������, -9999) And Nvl(A.��������, 9999) )';
    Else
      v_Sql := v_Sql || 'And (Decode(A.���䵥λ, ''��'',:��, ''��'',:��,''Сʱ'',:Сʱ,:��) ' ||
               ' Between Nvl(A.��������, -9999) And Nvl(A.��������, 9999) or 1=1 )';
    End If;

  Else
    v_Sql := v_Sql || 'And (Decode(A.���䵥λ, ''��'',:��, ''��'',:��,''Сʱ'',:Сʱ,:��) ' ||
             ' Between Nvl(A.��������, -9999) And Nvl(A.��������, 9999) or 1=1 )';
  End If;
  v_���䵥λ��1_Bound   := v_����;
  v_���䵥λ��1_Bound   := v_����;
  v_���䵥λСʱ1_Bound := v_Сʱ;
  v_���䵥λ��1_Bound   := v_����;

  --��������
  v_Sql := v_Sql || ' Order By a.Ĭ�� desc,A.�ٴ����� ';

  If Nvl(�������id_In, '') <> '' Or �������id_In Is Not Null Then
    v_Sql := v_Sql || ' ,a.�������ID  ';
  End If;

  If Nvl(�Ա�_In, '') <> '' Or �Ա�_In Is Not Null Then
    v_Sql := v_Sql || ' ,a.�Ա��� desc  ';
  Else
    v_Sql := v_Sql || ' ,a.�Ա��� ';
  End If;

  v_Sql := v_Sql || ' ,a.id ';

  v_Return := '';
  Open Cur For v_Sql
    Using v_��Ŀid_Bound, v_��Ŀid_Bound, v_�걾����_Bound, v_�Ա���1_Bound, v_�Ա���2_Bound, v_�Ա���3_Bound, v_����id_Bound, v_���䵥λ��_Bound, v_���䵥λ��_Bound, v_���䵥λСʱ_Bound, v_���䵥λ��_Bound, v_�ٴ�����_Bound, v_�������id_Bound, v_���䵥λ��1_Bound, v_���䵥λ��1_Bound, v_���䵥λСʱ1_Bound, v_���䵥λ��1_Bound;

  Loop
    Fetch Cur
      Into r_Emp;
    Exit When Cur%NotFound;
    If Cur%RowCount > 0 Then

      v_������� := r_Emp.�������;
      v_Valuerec := r_Emp.ȡֵ����;
      v_�ο�id   := r_Emp.Id;
      v_��ο�   := r_Emp.��ο�;

      If Nvl(v_Return, '') = '' Or v_Return Is Null Then
        If Type_In = 2 Then
          v_Return := r_Emp.Σ���ο�;
        Else
          v_Return := r_Emp.����ο�;
        End If;
      Else
        If Type_In = 2 Then
          v_Return := v_Return || Chr(13) || Chr(10) || r_Emp.Σ���ο�;
        Else
          If v_��ο� = 1 Then
            v_Return := v_Return || Chr(13) || Chr(10) || r_Emp.����ο�;
          End If;
        End If;
      End If;

      --ֻ���ӵ�һ��ѡ���ľ�ʾ�ο�
      If v_��ʾ���� = '' Or v_��ʾ���� Is Null Then
        v_��ʾ���� := r_Emp.��ʾ����;
      End If;
      If v_��ʾ���� = '' Or v_��ʾ���� Is Null Then
        v_��ʾ���� := r_Emp.��ʾ����;
      End If;
    End If;
  End Loop;

  If v_Return = '' Or v_Return Is Null Then
    Begin
      Select ����ο�, �������, ȡֵ����, ID, Σ���ο�, ��ʾ����, ��ʾ����
      Into v_����ο�, v_�������, v_Valuerec, v_�ο�id, v_Σ���ο�, v_��ʾ����, v_��ʾ����
      From (Select a.Id,
                    Trim(To_Char(a.�ο���ֵ, c.��ʽ)) || '��' || Trim(To_Char(a.�ο���ֵ, c.��ʽ)) ||
                     Decode(a.�ٴ�����, Null, '', '����', '', 'Ӥ��', '', ' ' || a.�ٴ�����) As ����ο�, b.�������, b.ȡֵ����,
                    Trim(To_Char(a.��ʾ����, c.��ʽ)) || '��' || Trim(To_Char(a.��ʾ����, c.��ʽ)) ||
                     Decode(a.�ٴ�����, Null, '', '����', '', 'Ӥ��', '', ' ' || a.�ٴ�����) As Σ���ο�, a.��ʾ����, a.��ʾ����
             From ������Ŀ�ο� A, ������Ŀ B,
                  (Select '9999990' ||
                            Decode(Max(Nvl(c.С��λ��, -1)), 0, '', -1, '.00', Substr('.000000', 1, 1 + Max(Nvl(c.С��λ��, -1)))) As ��ʽ
                    From ����������Ŀ C, ������Ŀ D
                    Where d.������Ŀid = ��Ŀid_In And d.������Ŀid = c.��Ŀid(+)) C
             Where a.��Ŀid = ��Ŀid_In And a.��Ŀid = b.������Ŀid
             Order By a.Ĭ�� Desc, a.�ٴ�����, a.�Ա���)
      Where Rownum = 1;
      If Type_In = 2 Then
        v_Return := v_Σ���ο�;
      Else
        v_Return := v_����ο�;
      End If;
      --ֻ���ӵ�һ��ѡ���ľ�ʾ�ο�
      If v_��ʾ���� = '' Or v_��ʾ���� Is Null Then
        v_��ʾ���� := r_Emp.��ʾ����;
      End If;
      If v_��ʾ���� = '' Or v_��ʾ���� Is Null Then
        v_��ʾ���� := r_Emp.��ʾ����;
      End If;
    Exception
      When Others Then
        v_Return := Null;
    End;
  End If;
  If v_Return <> '' Or v_Return Is Not Null Then

    If v_Return = '��' Then
      v_Return := '';
    Else
      If v_������� = 2 Then
        v_Pos := Instr(v_Return, '��');

        Begin
          Select To_Number(Substr(v_Return, 1, v_Pos - 1)) Into v_Value From Dual;
        Exception
          When Others Then
            v_Value := 0;
        End;
        v_Return := Zlsplit(v_Valuerec, ';', v_Value);
      End If;
    End If;
    If Type_In = 0 Then
      Return v_Return;
    Elsif Type_In = 1 Then
      Return v_�ο�id;
    Elsif Type_In = 2 Then
      Return v_Return;
    Elsif Type_In = 3 Then
      Return v_��ʾ����;
    Elsif Type_In = 4 Then
      Return v_��ʾ����;
    End If;
  End If;
  Close Cur; --�ر��α�
  Return v_Return;
End Zl_Get_Reference;
/

--108762:����,2017-05-23,���ٴ����ﰲ��ҽ������ǰ����ְ�Ʊ�ʶ��
Create Or Replace Procedure Zl_רҵ����ְ��_���±�ʶ��(�����ʶ��_In Varchar2) As
  --���ܣ��޸�ҽ��ְ���ʶ��
  --��ʽ������1,��ʶ��1;����2,��ʶ��2;...
  v_����   רҵ����ְ��.����%Type;
  v_��ʶ�� רҵ����ְ��.��ʶ��%Type;
Begin
  For c_�����ʶ�� In (Select C1 As ����, C2 As ��ʶ�� From Table(f_Str2list2(�����ʶ��_In, ';', ',')) Order By ����) Loop
    v_����   := c_�����ʶ��.����;
    v_��ʶ�� := c_�����ʶ��.��ʶ��;
    Update רҵ����ְ�� Set ��ʶ�� = v_��ʶ�� Where ���� = v_����;
  End Loop;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_רҵ����ְ��_���±�ʶ��;
/

--109314:����,2017-05-18,ȡ��ʧЧ������ʹ��

CREATE OR REPLACE Procedure Zl_������������ʽ_Updatedata
(
  ����id_In In ������������ʽ.����id%Type,
  Id_In     In ������������ʽ.Id%Type := Null
) Is

  v_Content Varchar2(2000);
  v_Xh      Varchar2(4000);

  --ֻ��ȡϵͳ��
  Cursor c_Callboard Is
    Select Id, ����, ����, �к�, λ��, �Ƿ�̶�, �Ƿ�����, ����, ʱ��
    From ������������ʽ
    Where ����id = ����id_In And (Id_In Is Null Or Id = Id_In)
    Order By �к�, λ��;

  Cursor c_Xry Is
    Select f_List2str(Cast(Collect(��Ժ����) As t_Strlist)) As ����Ժ
    From (Select b.��Ժ����
           From ������Ϣ a, ������ҳ b,
                (Select ����id, ��ҳid
                  From ���˱䶯��¼
                  Where ����id = ����id_In And (��ʼԭ�� = 2 Or ��ʼԭ�� = 1) And
                        ��ʼʱ�� Between To_Date(To_Char(Sysdate, 'yyyy-MM-dd') || ' 00:00:00', 'yyyy-MM-dd hh24:mi:ss') And
                        Sysdate
                  Group By ����id, ��ҳid) c, ��Ժ���� d
           Where a.����id = b.����id And a.��ҳID = b.��ҳid And b.����id = c.����id And b.��ҳid = c.��ҳid And a.����id = d.����id And
                 a.��ǰ����id = d.����id And d.����id = ����id_In And Nvl(b.����״̬, 0) <> 5 And b.���ʱ�� Is Null
           Order By b.��Ժ����);
  r_Xry c_Xry%Rowtype;

  Cursor c_Xzy Is
    Select f_List2str(Cast(Collect(��Ժ����) As t_Strlist)) As ��ת��
    From (Select b.��Ժ����
           From ������Ϣ a, ������ҳ b,
                (Select ����id, ��ҳid
                  From ���˱䶯��¼
                  Where ����id = ����id_In And (��ʼԭ�� = 3 Or ��ʼԭ�� = 15) And
                        ��ʼʱ�� Between To_Date(To_Char(Sysdate, 'yyyy-MM-dd') || ' 00:00:00', 'yyyy-MM-dd hh24:mi:ss') And
                        Sysdate
                  Group By ����id, ��ҳid) c, ��Ժ���� d
           Where a.����id = b.����id And a.��ҳID = b.��ҳid And b.����id = c.����id And b.��ҳid = c.��ҳid And a.����id = d.����id And
                 a.��ǰ����id = d.����id And d.����id = ����id_In And Nvl(b.����״̬, 0) <> 5 And b.���ʱ�� Is Null
           Order By b.��Ժ����);
  r_Xzy c_Xzy%Rowtype;

  Cursor c_Yjhl Is
    Select f_List2str(Cast(Collect(��Ժ����) As t_Strlist)) As һ������
    From (Select b.��Ժ����
           From ������Ϣ a, ������ҳ b, �շ���ĿĿ¼ c, ��Ժ���� d
           Where a.����id = b.����id And a.��ҳID = b.��ҳid And b.����ȼ�id = c.Id And a.����id = d.����id And a.��ǰ����id = d.����id And
                 d.����id = ����id_In And
                 (Instr(c.����, 'һ') > 0 Or Instr(c.����, 'I') > 0 Or Instr(c.����, '��') > 0 Or Instr(c.����, '1') > 0) And
                 Nvl(b.����״̬, 0) <> 5 And b.���ʱ�� Is Null
           Order By b.��Ժ����);
  r_Yjhl c_Yjhl%Rowtype;

  Cursor c_Tjhl Is
    Select f_List2str(Cast(Collect(��Ժ����) As t_Strlist)) As �ؼ�����
    From (Select b.��Ժ����
           From ������Ϣ a, ������ҳ b, �շ���ĿĿ¼ c, ��Ժ���� d
           Where a.����id = b.����id And b.����ȼ�id = c.Id And a.����id = d.����id And a.��ǰ����id = d.����id And d.����id = ����id_In And
                 (Instr(c.����, '��') > 0 Or Instr(c.����, '��') > 0) And Nvl(b.����״̬, 0) <> 5 And b.���ʱ�� Is Null
           Order By b.��Ժ����);
  r_Tjhl c_Tjhl%Rowtype;

  Cursor c_Bw Is
    Select f_List2str(Cast(Collect(��Ժ����) As t_Strlist)) As ��Σ
    From (Select b.��Ժ����
           From ������Ϣ a, ������ҳ b, ��Ժ���� d
           Where a.����id = b.����id And a.��ҳID = b.��ҳid And a.����id = d.����id And a.��ǰ����id = d.����id And d.����id = ����id_In And
                 Instr(b.��ǰ����, 'Σ') > 0 And Nvl(b.����״̬, 0) <> 5 And b.���ʱ�� Is Null
           Order By b.��Ժ����);
  r_Bw c_Bw%Rowtype;

  Cursor c_Ycy Is
    Select f_List2str(Cast(Collect(��Ժ����) As t_Strlist)) As Ԥ��Ժ
    From (Select b.��Ժ����
           From ������Ϣ a, ������ҳ b, ��Ժ���� c
           Where a.����id = b.����id And a.��ҳID = b.��ҳid And a.����id = c.����id And a.��ǰ����id = c.����id And c.����id = ����id_In And
                 b.״̬ = 3 And Nvl(b.����״̬, 0) <> 5 And b.���ʱ�� Is Null
           Order By b.��Ժ����);
  r_Ycy c_Ycy%Rowtype;

  Cursor c_Ss Is
    Select f_List2str(Cast(Collect(��Ժ����) As t_Strlist)) As ����
    From (Select Distinct d.��Ժ����
           From ������Ϣ b, ������ҳ d, ����ҽ����¼ a, ������ĿĿ¼ c, ��Ժ���� e
           Where b.����id = d.����id And b.��ҳID = d.��ҳid And d.����id = a.����id And d.��ҳid = a.��ҳid And b.����id = e.����id And
                 b.��ǰ����id = e.����id And e.����id = ����id_In And
                 ((a.ҽ����Ч = 0 And a.ҽ��״̬ In (3, 5, 6, 7, 8, 9) And (a.ִ����ֹʱ�� Is Null Or a.ִ����ֹʱ�� >= Sysdate)) Or
                 (a.ҽ����Ч = 1 And a.ҽ��״̬ In (3, 5, 6, 7, 8))) And
                 a.����ʱ�� Between To_Date(To_Char(Sysdate - 7, 'yyyy-MM-dd') || ' 00:00:00', 'yyyy-MM-dd hh24:mi:ss') And
                 Sysdate And
                 Substr(Nvl(a.�걾��λ, To_Char(��ʼִ��ʱ��, 'YYYY-MM-DD HH24:MI')), 1, 10) = To_Char(Sysdate, 'YYYY-MM-DD') And
                 Nvl(a.Ӥ��, 0) = 0 And a.������Ŀid = c.Id And c.��� = 'F' And Nvl(d.����״̬, 0) <> 5 And d.���ʱ�� Is Null
           Order By d.��Ժ����);
  r_Ss c_Ss%Rowtype;

  Cursor c_Fs Is
    Select f_List2str(Cast(Collect(��Ժ����) As t_Strlist)) As ����
    From (Select Distinct b.��Ժ����
           From ������Ϣ a, ������ҳ b, ��Ժ���� f
           Where a.����id = b.����id And a.��ҳID = b.��ҳid And a.����id = f.����id And a.��ǰ����id = f.����id And f.����id = ����id_In And
                 Nvl(b.����״̬, 0) <> 5 And b.���ʱ�� Is Null And Exists
            (Select c.Id
                  From ���˻����ļ� c, ���˻������� d, ���˻�����ϸ e
                  Where c.Id = d.�ļ�id And d.Id = e.��¼id And e.��¼���� = 1 And e.��Ŀ��� = 1 And
                        Length(Translate(e.��¼����, '-.0123456789' || e.��¼����, '-.0123456789')) = Length(e.��¼����) And
                        Zl_To_Number(e.��¼����) >= 37.2 And e.��ֹ�汾 Is Null And b.����id = c.����id And b.��ҳid = c.��ҳid And
                        Nvl(c.Ӥ��, 0) = 0 And d.����ʱ�� Between Sysdate - 3 And Sysdate)
           Order By b.��Ժ����);
  r_Fs c_Fs%Rowtype;

  Cursor c_Gms Is
    Select f_List2str(Cast(Collect(��Ժ����) As t_Strlist)) As ����ʷ
    From (Select Distinct b.��Ժ����
           From ������Ϣ a, ������ҳ b, ���˹�����¼ c, ��Ժ���� d
           Where a.����id = b.����id And a.��ҳID = b.��ҳid And b.����id = c.����id And a.����id = d.����id And a.��ǰ����id = d.����id And
                 d.����id = ����id_In And Nvl(b.����״̬, 0) <> 5 And b.���ʱ�� Is Null And c.��� = 1 And Not Exists
            (Select ҩ��id
                  From ���˹�����¼
                  Where (Nvl(ҩ��id, 0) = Nvl(c.ҩ��id, 0) Or Nvl(ҩ����, 'Null') = Nvl(c.ҩ����, 'Null')) And Nvl(���, 0) = 0 And
                        ��¼ʱ�� > c.��¼ʱ�� And ����id = c.����id)
           Order By b.��Ժ����);
  r_Gms c_Gms%Rowtype;

  Cursor c_Diy Is
    Select /*+ Rule */
     f_List2str(Cast(Collect(��ǰ����) As t_Strlist)) As �����б�
    From (Select Distinct b.��ǰ����
           From ������Ϣ b, ������ҳ c, ����ҽ����¼ a, ��Ժ���� d, ((Select Column_Value From Table(f_Num2list(v_Xh)))) e
           Where b.����id = c.����id And b.��ҳID = c.��ҳid And c.����id = a.����id And c.��ҳid = a.��ҳid And b.����ID=d.����ID And b.��ǰ����id = d.����id And
                 d.����id = ����id_In And
                 ((a.ҽ����Ч = 0 And a.ҽ��״̬ In (3, 5, 6, 7, 8, 9) And a.��ʼִ��ʱ�� >= b.��Ժʱ�� And
                 (a.ִ����ֹʱ�� Is Null Or a.ִ����ֹʱ�� >= Sysdate)) Or
                 (a.ҽ����Ч = 1 And a.ҽ��״̬ In (3, 5, 6, 7, 8) And a.��ʼִ��ʱ�� Between Sysdate - 1 And Sysdate)) And
                 Nvl(a.Ӥ��, 0) = 0 And a.������Ŀid + 0 = e.Column_Value And Nvl(c.����״̬, 0) <> 5 And c.���ʱ�� Is Null
           Order By b.��ǰ����);
  r_Diy c_Diy%Rowtype;

Begin
  For r_Board In c_Callboard Loop
    v_Content := '';
    If Instr(',����Ժ�б�,��ת���б�,һ�������б�,�ؼ������б�,��Σ�б�,Ԥ��Ժ�б�,�����б�,�����б�,����ʷ�б�,',
             ',' || r_Board.���� || ',') > 0 Then
      --ϵͳ�̶���
      If r_Board.���� = '����Ժ�б�' Then
        Open c_Xry;
        Fetch c_Xry
          Into r_Xry;
        If c_Xry%Rowcount > 0 Then
          v_Content := Nvl(r_Xry.����Ժ, '');
        End If;

        Update ������������ʽ Set ���� = v_Content, ʱ�� = Sysdate Where Id = r_Board.Id;
        Close c_Xry;
      Elsif r_Board.���� = '��ת���б�' Then
        Open c_Xzy;
        Fetch c_Xzy
          Into r_Xzy;
        If c_Xzy%Rowcount > 0 Then
          v_Content := Nvl(r_Xzy.��ת��, '');
        End If;

        Update ������������ʽ Set ���� = v_Content, ʱ�� = Sysdate Where Id = r_Board.Id;
        Close c_Xzy;
      Elsif r_Board.���� = 'һ�������б�' Then
        Open c_Yjhl;
        Fetch c_Yjhl
          Into r_Yjhl;
        If c_Yjhl%Rowcount > 0 Then
          v_Content := Nvl(r_Yjhl.һ������, '');
        End If;

        Update ������������ʽ Set ���� = v_Content, ʱ�� = Sysdate Where Id = r_Board.Id;
        Close c_Yjhl;
      Elsif r_Board.���� = '�ؼ������б�' Then
        Open c_Tjhl;
        Fetch c_Tjhl
          Into r_Tjhl;
        If c_Tjhl%Rowcount > 0 Then
          v_Content := Nvl(r_Tjhl.�ؼ�����, '');
        End If;

        Update ������������ʽ Set ���� = v_Content, ʱ�� = Sysdate Where Id = r_Board.Id;
        Close c_Tjhl;
      Elsif r_Board.���� = '��Σ�б�' Then
        Open c_Bw;
        Fetch c_Bw
          Into r_Bw;
        If c_Bw%Rowcount > 0 Then
          v_Content := Nvl(r_Bw.��Σ, '');
        End If;

        Update ������������ʽ Set ���� = v_Content, ʱ�� = Sysdate Where Id = r_Board.Id;
        Close c_Bw;
      Elsif r_Board.���� = 'Ԥ��Ժ�б�' Then
        Open c_Ycy;
        Fetch c_Ycy
          Into r_Ycy;
        If c_Ycy%Rowcount > 0 Then
          v_Content := Nvl(r_Ycy.Ԥ��Ժ, '');
        End If;

        Update ������������ʽ Set ���� = v_Content, ʱ�� = Sysdate Where Id = r_Board.Id;
        Close c_Ycy;
      Elsif r_Board.���� = '�����б�' Then
        Open c_Ss;
        Fetch c_Ss
          Into r_Ss;
        If c_Ss%Rowcount > 0 Then
          v_Content := Nvl(r_Ss.����, '');
        End If;

        Update ������������ʽ Set ���� = v_Content, ʱ�� = Sysdate Where Id = r_Board.Id;
        Close c_Ss;
      Elsif r_Board.���� = '�����б�' Then
        Open c_Fs;
        Fetch c_Fs
          Into r_Fs;
        If c_Fs%Rowcount > 0 Then
          v_Content := Nvl(r_Fs.����, '');
        End If;

        Update ������������ʽ Set ���� = v_Content, ʱ�� = Sysdate Where Id = r_Board.Id;
        Close c_Fs;
      Else
        Open c_Gms;
        Fetch c_Gms
          Into r_Gms;
        If c_Gms%Rowcount > 0 Then
          v_Content := Nvl(r_Gms.����ʷ, '');
        End If;

        Update ������������ʽ Set ���� = v_Content, ʱ�� = Sysdate Where Id = r_Board.Id;
        Close c_Gms;
      End If;
    Else
      --������Ѱ󶨵���Ŀ
      v_Content := '';
      Begin
        Select f_List2str(Cast(Collect(a.Xh) As t_Strlist))
        Into v_Xh
        From ������������ʽ p, Xmltable('/ITEMLIST/ITEM/XH' Passing p.������Ŀ Columns Xh Varchar2(256) Path '/XH') a
        Where p.Id = r_Board.Id;
      Exception
        When Others Then
          v_Xh := '';
      End;

      If v_Xh Is Not Null Then
        Open c_Diy;
        Fetch c_Diy
          Into r_Diy;
        If c_Diy%Rowcount > 0 Then
          v_Content := Nvl(r_Diy.�����б�, '');
        End If;

        Update ������������ʽ Set ���� = v_Content, ʱ�� = Sysdate Where Id = r_Board.Id;
        Close c_Diy;
      Else
        Update ������������ʽ Set ���� = Null, ʱ�� = Sysdate Where Id = r_Board.Id;
      End If;
    End If;
  End Loop;
Exception
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_������������ʽ_Updatedata;
/

--109214:������,2017-05-18,ȡ��ʧЧ������ʹ��
Create Or Replace Procedure Zl_Third_Getvisitdetails
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  -------------------------------------------------------------------------------------------------- 
  --����:���ݹҺŵ��Ż�ȡ�ôξ�������
  --���:Xml_In: 
  --<IN>
  --    <GHDH>�Һŵ���</GHDH>
  --    <JSKLB>���㿨���</JSKLB>
  --</IN>
  --����:Xml_Out 
  --<OUTPUT>
  --    <DJLIST>  //���Ϊ�ձ�ʾΪ�ҵ�����
  --        <DJ>
  --            <NO>���ݺ�</NO>
  --            <DJLX>��������</DJLX> //1-�շѵ�;4-�Һŵ�
  --            <KDSJ>����ʱ��</KDSJ>
  --            <ZFZT>֧��״̬</ZFZT>    //0δ֧��1��֧��
  --            <SFJSK>�Ƿ���㿨֧��</SFJSK> //���õ����Ƿ�������<JSKLB>������֧����,�Ƿ���1,���򷵻�0
  --            <LX>����</LX> //�Һŵ��̶�Ϊ�Һ�,�������շ�������
  --            <ZXKS>ִ�п���</ZXKS>
  --            <ZXKSID>ִ�п���ID</ZXKSID>
  --            <MXLIST> 
  --                     <MX>
  --                                <JZSJ>����ʱ��</JZSJ>    //�Һ���Ч:yyyy-mm-dd hh24:mi:ss
  --                                <BW>��λ</BW>               //���,����ʱ��Ч
  --                                <XM>��Ŀ����</XM>     //�Һ���Ч:������Ŀ��Ч
  --                                <ZXZT>ִ��״̬</ZXZT> //�Һ�:δ����;�ѽ���;��ɾ���;�շ�:δִ��;��ִ��;����ִ��
  --                                <BG>����״̬</BG>// 1-�ѳ�����;0δ������,���,����ʱ��Ч 
  --                                <BLID>����ID</BLID>  //���<BG>�ֶ�Ϊ1����ֵ��Ϊ��,���,����ʱ��Ч
  --                                <GG>���</GG>                       //ҩƷ��Ч
  --                                <SL>����</SL> //�ǹҺ���Ч
  --                                <DW>��λ</DW> //�ǹҺ���Ч
  --                                <DJ>����</DJ> //�ǹҺ���Ч
  --                                <JE>���</JE>  
  --                     </MX>
  --             </MXLIST>
  --             <DL> //����
  --                        <XH>���</XH>
  --                        <QMRS>ǰ������</QMRS>  //(��Oracle����zl_GetSequenceBeforPerons��ȡ)
  --             </DL>
  --        </DJ>
  --    </DJLIST>
  --    <ERROR><MSG></MSG></ERROR>                      //������󷵻�
  --</OUTPUT>

  -------------------------------------------------------------------------------------------------- 
  v_Err_Msg Varchar2(200);
  Err_Item Exception;
  x_Templet Xmltype; --ģ��XML 

  v_�����   Varchar2(100);
  n_�����id Number(18);
  v_�Һŵ�   Varchar2(10);
  v_�ŶӺ��� Varchar2(10);
  n_Temp     Number(18);

  n_Count Number(18);

  v_Temp       Varchar2(32767); --��ʱXML 
  v_����       Varchar2(32767);
  v_No         Varchar2(50);
  v_Tmp        Varchar2(4000);
  n_Add_Djlist Number(1); --�Ƿ�������DJLIST��;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/GHDH'), Extractvalue(Value(A), 'IN/JSKLB')
  Into v_�Һŵ�, v_�����
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If v_�Һŵ� Is Null Then
    v_Err_Msg := '�����ҵ�ָ���ĹҺŵ���(��ǰ�Һŵ���Ϊ��)';
    Raise Err_Item;
  End If;
  n_Add_Djlist := 0;

  v_Err_Msg := Null;
  If v_����� Is Not Null Then
    Begin
      n_�����id := To_Number(v_�����);
    Exception
      When Others Then
        n_�����id := 0;
    End;
  
    If n_�����id = 0 Then
      Begin
        Select ID, Decode(Nvl(�Ƿ�����, 0), 1, Null, ���� || 'δ����,��������нɷ�!')
        Into n_�����id, v_Err_Msg
        From ҽ�ƿ����
        Where ���� = v_�����;
      Exception
        When Others Then
          v_Err_Msg := '�����:' || v_����� || '������!';
      End;
    
    Else
    
      Begin
        Select ID, Decode(Nvl(�Ƿ�����, 0), 1, Null, ���� || 'δ����,��������нɷ�!')
        Into n_�����id, v_Err_Msg
        From ҽ�ƿ����
        Where ID = n_�����id;
      Exception
        When Others Then
          v_Err_Msg := 'δ�ҵ�ָ���Ľ���֧����Ϣ!';
      End;
    
    End If;
    If Not v_Err_Msg Is Null Then
      Raise Err_Item;
    End If;
  End If;
  --1.��ȡ�Һ�����
  n_Count := 0;
  For c_�Һ� In (Select a.Id, a.No, a.��¼����, a.ִ�в���id, c.���� As ִ�в���, To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��,
                      a.ԤԼʱ��, a.����ʱ��, To_Char(a.����ʱ��, 'yyyy-mm-dd HH24:mi') As ����ʱ��, a.�ű�, a.����, b.���, a.��¼״̬,
                      Decode(Nvl(a.ִ��״̬, 0), 0, '�ȴ�����', 1, '��ɾ���', 2, '���ھ���', -1, 'ȡ������') As ִ��״̬,
                      Decode(Nvl(b.����id, 0), 0, 0, 1) As ֧����־
               From ���˹Һż�¼ A,
                    (Select NO, Max(Nvl(����id, 0)) As ����id, Sum(ʵ�ս��) As ���
                      From ������ü�¼ B
                      Where ��¼���� = 4 And NO = v_�Һŵ�
                      Group By NO) B, ���ű� C
               Where a.No = v_�Һŵ� And a.No = b.No And a.ִ�в���id = c.Id(+)) Loop
    If Nvl(c_�Һ�.��¼״̬, 0) <> 1 Then
      v_Err_Msg := '���ݺ�:' || v_�Һŵ� || '�Ѿ����˺�!';
      Raise Err_Item;
    End If;
    Begin
      Select �ŶӺ��� Into v_�ŶӺ��� From �ŶӽкŶ��� Where ҵ��id = c_�Һ�.Id And Nvl(ҵ������, 0) = 0;
    Exception
      When Others Then
        v_�ŶӺ��� := Null;
    End;
    If v_�ŶӺ��� Is Not Null Then
      --ҵ��id_In ,ҵ������_In �ŶӺ���_In Number := Null
      n_Temp := Zl_Getsequencebeforperons(c_�Һ�.Id, 0, v_�ŶӺ���);
      v_���� := v_���� || '<DL><XH>' || v_�ŶӺ��� || '</XH><QMRS>' || n_Temp || '</QMRS></DL>';
    End If;
    n_Temp := 0;
    If Nvl(n_�����id, 0) <> 0 Then
      Begin
        Select 1
        Into n_Temp
        From ����Ԥ����¼
        Where NO = v_�Һŵ� And ��¼���� = 4 And ��¼״̬ In (1, 3) And �����id = n_�����id And Rownum < 2;
      Exception
        When Others Then
          Null;
      End;
    End If;
    v_Temp := '<NO>' || c_�Һ�.No || '</NO>';
    v_Temp := v_Temp || '<DJLX>' || 4 || '</DJLX>';
    v_Temp := v_Temp || '<KDSJ>' || c_�Һ�.�Ǽ�ʱ�� || '</KDSJ>';
    v_Temp := v_Temp || '<ZFZT>' || c_�Һ�.֧����־ || '</ZFZT>';
    v_Temp := v_Temp || '<SFJSK>' || n_Temp || '</SFJSK>';
    v_Temp := v_Temp || '<LX>�Һ�</LX>';
    v_Temp := v_Temp || '<ZXKS>' || c_�Һ�.ִ�в��� || '</ZXKS>';
    v_Temp := v_Temp || '<ZXKSID>' || c_�Һ�.ִ�в���id || '</ZXKSID>';
    v_Temp := v_Temp || '<MXLIST><MX><JZSJ>' || c_�Һ�.����ʱ�� || '</JZSJ><JE>' || c_�Һ�.��� || '</JE></MX></MXLIST>';
    If v_���� Is Not Null Then
      v_Temp := v_Temp || v_����;
    End If;
  
    If Nvl(n_Add_Djlist, 0) = 0 Then
      --����DJList�ڵ�
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype('<DJLIST></DJLIST>')) Into x_Templet From Dual;
      n_Add_Djlist := 1;
    End If;
    v_Temp := '<DJ>' || v_Temp || '</DJ>';
    Select Appendchildxml(x_Templet, '/OUTPUT/DJLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
    n_Count := n_Count + 1;
  End Loop;
  If Nvl(n_Count, 0) = 0 Then
    v_Err_Msg := 'δ�ҵ�ָ���ĹҺŵ���:' || v_�Һŵ� || '!';
    Raise Err_Item;
  End If;

  --2.��������շѵ�
  v_No := '-_';

  For c_���� In (Select j.ҽ��id, j.���id As ���, j.No, j.���, j.�շ����, i.���� As �շ������, j.ִ�в���id, q.���� As ִ�в���, j.�շ�ϸĿid, m.����,
                      m.���, Max(j.���㵥λ) As ���㵥λ, Decode(Max(j.ִ��״̬), 0, 'δִ��', 1, '��ȫִ��', 2, '����ִ��', '') As ִ��״̬,
                      Max(j.����״̬) As ����״̬, To_Char(Max(j.�Ǽ�ʱ��), 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��, Max(j.����) As ����,
                      Sum(j.����) As ����, Sum(j.ʵ�ս��) As ʵ�ս��
               From (Select a.���id, a.Id As ҽ��id, b.No, b.�շ����, Max(Decode(b.��¼״̬, 0, 0, 1)) As ����״̬, b.����id, b.ִ�в���id,
                             Max(Decode(b.��¼״̬, 2, 0, b.ִ��״̬)) As ִ��״̬,
                             Max(Decode(b.��¼״̬, 2, Null + Sysdate, b.�Ǽ�ʱ��)) As �Ǽ�ʱ��, Nvl(b.�۸񸸺�, b.���) As ���, b.�շ�ϸĿid,
                             b.���㵥λ, Sum(b.��׼����) As ����, Avg(Nvl(b.����, 1) * b.����) As ����, Sum(b.ʵ�ս��) As ʵ�ս��

                      
                      From ������ü�¼ B, ����ҽ����¼ A
                      Where Mod(b.��¼����, 10) = 1 And a.Id = b.ҽ����� And Nvl(b.����״̬, 0) = 0 And a.�Һŵ� = v_�Һŵ�
                      Group By a.���id, a.Id, b.No, b.�շ����, b.����id, b.ִ�в���id, Nvl(b.�۸񸸺�, b.���), b.�շ�ϸĿid, b.���㵥λ) J,
                    �շ���ĿĿ¼ M, ���ű� Q, �շ���Ŀ��� I
               Where j.�շ�ϸĿid = m.Id And j.ִ�в���id = q.Id(+) And j.�շ���� = i.����(+)
               Group By j.ҽ��id, j.���id, j.No, j.���, j.�շ����, i.����, j.ִ�в���id, q.����, j.�շ�ϸĿid, m.����, m.���
               Order By �Ǽ�ʱ�� Desc, NO Desc, �շ����, ���) Loop
    If c_����.No <> v_No Then
      n_Temp := 0;
      --���ݲ�ͬ,������Ľṹ��ͬ
      If Nvl(c_����.����״̬, 0) = 1 Then
        --�Ƿ���㿨֧����
        Begin
          Select 1
          Into n_Temp
          From ����Ԥ����¼ A, ������ü�¼ B
          Where a.����id = b.����id And b.No = c_����.No And Mod(b.��¼����, 10) = 1 And b.��¼״̬ In (1, 3) And a.�����id = n_�����id And
                Rownum < 2;
        Exception
          When Others Then
            n_Temp := 0;
        End;
      End If;
      v_Tmp := Null;
      Begin
        Select f_List2str(Cast(Collect(����) As t_Strlist))
        Into v_Tmp
        From (Select Distinct b.����
               From ������ü�¼ A, �շ���Ŀ��� B
               Where a.�շ���� = b.���� And a.No = c_����.No And a.��¼���� = 1 And a.��¼״̬ In (1, 3));
      Exception
        When Others Then
          Null;
      End;
    
      If Nvl(n_Add_Djlist, 0) = 0 Then
        --����DJList�ڵ�
        Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype('<DJLIST></DJLIST>')) Into x_Templet From Dual;
        n_Add_Djlist := 1;
      End If;
    
      v_No   := c_����.No;
      v_Temp := '<NO>' || c_����.No || '</NO>';
      v_Temp := v_Temp || '<DJLX>' || 1 || '</DJLX>';
      v_Temp := v_Temp || '<KDSJ>' || c_����.�Ǽ�ʱ�� || '</KDSJ>';
      v_Temp := v_Temp || '<ZFZT>' || c_����.����״̬ || '</ZFZT>';
      v_Temp := v_Temp || '<SFJSK>' || n_Temp || '</SFJSK>';
    
      v_Temp := v_Temp || '<LX>' || Nvl(Replace(v_Tmp, ',', '/'), '') || '</LX>';
      v_Temp := v_Temp || '<ZXKS>' || c_����.ִ�в��� || '</ZXKS>';
      v_Temp := v_Temp || '<ZXKSID>' || c_����.ִ�в���id || '</ZXKSID>';
      v_Temp := v_Temp || '<MXLIST></MXLIST>' || Nvl(v_����, '') || '';
      v_Temp := '<DJ NO="' || c_����.No || '">' || v_Temp || '</DJ>';
      Select Appendchildxml(x_Templet, '/OUTPUT/DJLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
    End If;
  
    v_Temp := '<XM>' || Nvl(c_����.����, '') || '</XM>';
    If c_����.�շ���� = 'D' Then
      --����ȡ��λ
      Begin
        Select f_List2str(Cast(Collect(�걾��λ) As t_Strlist))
        Into v_Tmp
        From ����ҽ����¼
        Where ���id = c_����.ҽ��id;
      Exception
        When Others Then
          v_Tmp := Null;
      End;
      v_Temp := v_Temp || '<BW>' || Nvl(v_Tmp, '') || '</BW>';
    Elsif c_����.�շ���� = 'C' Then
      --����
      Begin
        Select Max(Decode(b.���ʱ��, Null, 0, 1))
        Into n_Temp
        From ����ҽ����¼ A, ����걾��¼ B
        Where a.Id = c_����.ҽ��id And a.Id = b.ҽ��id(+) And Exists
         (Select 1 From ����ҽ����¼ Where ���id = c_����.ҽ��id And ������� = 'C');
      Exception
        When Others Then
          n_Temp := 0;
      End;
      v_Temp := v_Temp || '<BG>' || n_Temp || '</BG>';
      If n_Temp = 1 Then
        --ȡ����ID
        Begin
          Select ����id
          Into n_Temp
          From ����ҽ������
          Where ҽ��id = c_����.ҽ��id And Nvl(����id, 0) <> 0 And Rownum < 2;
        Exception
          When Others Then
            n_Temp := Null;
        End;
        v_Temp := v_Temp || '<BLID>' || Nvl(n_Temp, '') || '</BLID>';
      End If;
    End If;
  
    v_Temp := v_Temp || '<GG>' || Nvl(c_����.���, '') || '</GG>';
    v_Temp := v_Temp || '<SL>' || Nvl(c_����.����, 0) || '</SL>';
    v_Temp := v_Temp || '<DW>' || Nvl(c_����.���㵥λ, '') || '</DW>';
    v_Temp := v_Temp || '<DJ>' || Nvl(c_����.����, 0) || '</DJ>';
    v_Temp := v_Temp || '<JE>' || Nvl(c_����.ʵ�ս��, 0) || '</JE>';
    v_Temp := '<MX>' || v_Temp || '</MX>';
    Select Appendchildxml(x_Templet, '/OUTPUT/DJLIST/DJ[@NO="' || v_No || '"]/MXLIST', Xmltype(v_Temp))
    Into x_Templet
    From Dual;
  
  End Loop;
  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getvisitdetails;
/

--109174:���ϴ�,2017-05-16,ɾ��ҽ�ƿ�ͬʱɾ����Ӧ�ض���Ŀ
Create Or Replace Procedure Zl_ҽ�ƿ����_Delete(Id_In In ҽ�ƿ����.ID%Type) Is
  v_Err_Msg Varchar2(200);
  Err_Item Exception;
  n_�Ƿ����� Number;
  n_�Ƿ�̶� Number;
  v_�ض���Ŀ Varchar2(20);
Begin
  Begin
    Select �Ƿ�����, �Ƿ�̶�, �ض���Ŀ Into n_�Ƿ�����, n_�Ƿ�̶�, v_�ض���Ŀ From ҽ�ƿ���� Where ID = Id_In;
  Exception
    When Others Then
      n_�Ƿ����� := -1;
  End;
  If Nvl(n_�Ƿ�����, 0) = -1 Then
    v_Err_Msg := '[ZLSOFT]ҽ�ƿ������ܱ�������ɾ���������ٴ�ɾ��![ZLSOFT]';
    Raise Err_Item;
  End If;
  If Nvl(n_�Ƿ�����, 0) = 0 Then
    v_Err_Msg := '[ZLSOFT]ҽ�ƿ�����Ѿ���ͣ�ã�����ɾ��![ZLSOFT]';
    Raise Err_Item;
  End If;
  If Nvl(n_�Ƿ�̶�, 0) = 1 Then
    v_Err_Msg := '[ZLSOFT]ҽ�ƿ������ϵͳ�̶��ģ�����ɾ��![ZLSOFT]';
    Raise Err_Item;
  End If;

  Delete From ҽ�ƿ���� Where ID = Id_In And Nvl(�Ƿ�����, 0) = 1 And Nvl(�Ƿ�̶�, 0) = 0;
  If Sql%NotFound Then
    v_Err_Msg := '[ZLSOFT]ҽ�ƿ������ܱ�������ɾ���������ٴ�ɾ��![ZLSOFT]';
    Raise Err_Item;
  End If;
  
  IF Not v_�ض���Ŀ is Null then
    Delete From �շ��ض���Ŀ Where �ض���Ŀ = v_�ض���Ŀ;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  When Others Then
    Zl_Errorcenter(Sqlcode, Sqlerrm);
End Zl_ҽ�ƿ����_Delete;
/

--108046:������,2017-05-15,������˹���ȡ����˳�������
Create Or Replace Procedure Zl_������˼�¼_Delete
(
  ����id_In In ������˼�¼.����id%Type,
  ����_In   In ������˼�¼.����%Type
) Is
  v_Err_Msg Varchar2(100);
  Err_Item Exception;
  n_Count Number(4);
  v_No    ������ü�¼.No%Type;
Begin
  Select Count(a.Id), Max(a.No)
  Into n_Count, v_No
  From ������ü�¼ A, (Select Mod(��¼����, 10) As ��¼����, NO, ��� From ������ü�¼ Where ID = ����id_In) B
  Where a.No = b.No And Mod(a.��¼����, 10) = b.��¼���� And a.��� = b.���
  Group By a.No, Mod(a.��¼����, 10), a.���
  Having Sum(a.����) <> 0 Or Sum(a.ʵ�ս��) <> 0;
  If n_Count = 0 Then
    v_Err_Msg := '[ZLSOFT]���ݡ�' || v_No || '�������򲢷�ԭ��,�Ѿ�������ת�����˷�,����ȡ�����![ZLSOFT]';
    Raise Err_Item;
  End If;
  Delete From ������˼�¼ Where ����id = ����id_In And ���� = ����_In;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, v_Err_Msg);
  
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_������˼�¼_Delete;
/

--109116:������,2017-05-15,�Һ�ҽ��У����������
Create Or Replace Procedure Zl_���˽����¼_Update
(
  ����id_In       ����Ԥ����¼.����id%Type,
  ���ս���_In     Varchar2, --"���㷽ʽ|������||....."
  ����_In         Number := 0,
  ȱʡ���㷽ʽ_In Varchar2 := Null,
  ȱʡ��Ԥ��_In   Number := 0, --0-���ֽ�ɿ�,1:ʣ�ڿ����ó�Ԥ��֧��(����Ԥ��),2-ʣ�ڿ����ó�Ԥ��֧��(סԺԤ��)
  �����id_In     ����Ԥ����¼.�����id%Type := Null,
  ���㿨���_In   ����Ԥ����¼.���㿨���%Type := Null,
  ����_In         ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In   ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In     ����Ԥ����¼.����˵��%Type := Null,
  ������λ_In     ����Ԥ����¼.������λ%Type := Null,
  ������ҽ��_In   Number := 0
) As
  --���α�ΪҪɾ�����ɷ��ü�¼�����Ľ����¼

  Cursor c_Del Is
    Select a.Id, a.��¼����, a.��Ԥ��, a.���㷽ʽ, b.����, a.Ԥ�����
    From ����Ԥ����¼ A, ���㷽ʽ B
    Where a.���㷽ʽ = b.���� And a.����id = ����id_In;

  Cursor c_Del_ҽ�� Is
    Select a.Id, a.��¼����, a.��Ԥ��, a.���㷽ʽ, b.����, a.Ԥ�����
    From ����Ԥ����¼ A, ���㷽ʽ B
    Where a.���㷽ʽ = b.���� And b.���� In (3, 4) And a.����id = ����id_In;

  --�����Ϣ
  v_No         ����Ԥ����¼.No%Type;
  v_����id     סԺ���ü�¼.����id%Type;
  v_��ҳid     סԺ���ü�¼.��ҳid%Type;
  v_����ʱ��   סԺ���ü�¼.����ʱ��%Type;
  v_�Ǽ�ʱ��   סԺ���ü�¼.�Ǽ�ʱ��%Type;
  v_����Ա��� סԺ���ü�¼.����Ա���%Type;
  v_����Ա���� סԺ���ü�¼.����Ա����%Type;

  --���ν������
  v_���ϼ� ����Ԥ����¼.��Ԥ��%Type;

  --���ս���
  v_���ս��� Varchar2(255);
  v_��ǰ���� Varchar2(50);
  v_�ֽ���� ����Ԥ����¼.���㷽ʽ%Type;
  v_���㷽ʽ ����Ԥ����¼.���㷽ʽ%Type;
  v_������ ����Ԥ����¼.��Ԥ��%Type;

  v_��¼���� ����Ԥ����¼.��¼����%Type;
  v_ȱʡ     ����Ԥ����¼.���㷽ʽ%Type;

  --�ֱҴ���������
  v_�ֽ���   ����Ԥ����¼.��Ԥ��%Type;
  v_Cashcented ����Ԥ����¼.��Ԥ��%Type;
  v_�����   ����Ԥ����¼.��Ԥ��%Type;
  v_����id     סԺ���ü�¼.Id%Type;
  v_���       סԺ���ü�¼.���%Type;
  v_�շ����   סԺ���ü�¼.�շ����%Type;
  v_�շ�ϸĿid סԺ���ü�¼.�շ�ϸĿid%Type;
  v_������Ŀid סԺ���ü�¼.������Ŀid%Type;
  v_�վݷ�Ŀ   סԺ���ü�¼.�վݷ�Ŀ%Type;
  n_Noexists   Number(3);
  n_ҽ��С��id סԺ���ü�¼.ҽ��С��id%Type;
  n_�������   ����Ԥ����¼.�������%Type;
  n_����״̬   ������ü�¼.����״̬%Type;
  n_Ԥ�����   ����Ԥ����¼.���%Type;
  n_��ǰ���   ����Ԥ����¼.���%Type;
  v_�����     ���㷽ʽ.����%Type;

  --��ʱ����
  Err_Item Exception;
  v_Err_Msg  Varchar2(255);
  n_��id     ����ɿ����.Id%Type;
  n_ִ��״̬ ������ü�¼.ִ��״̬%Type;
Begin
  --���ȱʡ���㷽ʽΪ�գ���ȡ�ֽ���㷽ʽ
  If ȱʡ���㷽ʽ_In Is Null Then
    Begin
      Select ���� Into v_ȱʡ From ���㷽ʽ Where ���� = 1 And Rownum < 2;
    Exception
      When Others Then
        v_ȱʡ := '�ֽ�';
    End;
  Else
    v_ȱʡ := ȱʡ���㷽ʽ_In;
  End If;

  --ȡ�ñ��ν���������Ϣ
  If Nvl(����_In, 0) = 1 Then
    Select NO, ����id, �շ�ʱ��, ����Ա���, ����Ա����, �ɿ���id, 0
    Into v_No, v_����id, v_�Ǽ�ʱ��, v_����Ա���, v_����Ա����, n_��id, n_ִ��״̬
    From ���˽��ʼ�¼
    Where ID = ����id_In;
  Else
    Begin
      n_Noexists := 0;
      Select NO, ����id, �Ǽ�ʱ��, ����Ա���, ����Ա����, �ɿ���id, ִ��״̬, ����״̬
      Into v_No, v_����id, v_�Ǽ�ʱ��, v_����Ա���, v_����Ա����, n_��id, n_ִ��״̬, n_����״̬
      From ������ü�¼
      Where ����id = ����id_In And Rownum < 2;
    Exception
      When Others Then
        n_Noexists := 1;
    End;
    If n_Noexists = 1 Then
      --���ü�¼�����ڣ��Ӳ����¼����
      Select NO, ����id, �Ǽ�ʱ��, ����Ա���, ����Ա����, �ɿ���id, ����״̬
      Into v_No, v_����id, v_�Ǽ�ʱ��, v_����Ա���, v_����Ա����, n_��id, n_����״̬
      From ���ò����¼
      Where ����id = ����id_In And Rownum < 2;
    End If;
    If Nvl(n_����״̬, 0) = 1 Then
      --�쳣����Ϊ��:
      v_ȱʡ := Null;
    End If;
  
    Begin
      --20051027 �¶�
      Select ��¼����
      Into v_��¼����
      From ����Ԥ����¼
      Where ����id = ����id_In And Rownum = 1 And Mod(��¼����, 10) <> 1;
    Exception
      When Others Then
        v_��¼���� := -1;
    End;
    If v_��¼���� = -1 Then
      Begin
        Select Decode(��¼����, 1, 3, 11, 3, 4, 4, ��¼����)
        Into v_��¼����
        From ������ü�¼
        Where ����id = ����id_In And Rownum = 1;
      Exception
        When Others Then
          --�����ǿ���
          Select ��¼���� Into v_��¼���� From סԺ���ü�¼ Where ����id = ����id_In And Rownum = 1;
      End;
    End If;
  End If;

  If Nvl(v_����id, 0) <> 0 And Nvl(����_In, 0) = 1 Then
    Select ��ҳid Into v_��ҳid From ������Ϣ Where ����id = v_����id;
  End If;
  Select ������� Into n_������� From ����Ԥ����¼ Where ����id = ����id_In And Rownum = 1;

  ----���˽ɿ�,Ԥ������,��Ϊû�иĳ�Ԥ����
  --�շ�δ��δ������ɵ�,�����쳣��������,��������Ա�ɿ����
  v_���ϼ� := 0;
  If Nvl(������ҽ��_In, 0) = 0 Then
    For r_Del In c_Del Loop
      If r_Del.��¼���� Not In (1, 11) Then
        If Nvl(n_����״̬, 0) <> 1 Then
          Update ��Ա�ɿ����
          Set ��� = Nvl(���, 0) - r_Del.��Ԥ��
          Where �տ�Ա = v_����Ա���� And ���� = 1 And ���㷽ʽ = r_Del.���㷽ʽ;
        
          If Sql%RowCount = 0 Then
            Insert Into ��Ա�ɿ����
              (�տ�Ա, ���㷽ʽ, ����, ���)
            Values
              (v_����Ա����, r_Del.���㷽ʽ, 1, -1 * r_Del.��Ԥ��);
          End If;
        End If;
        v_���ϼ� := v_���ϼ� + r_Del.��Ԥ��;
        Delete From ����Ԥ����¼ Where ID = r_Del.Id;
      Else
        --����Ƿ��Ԥ��
        If Nvl(ȱʡ��Ԥ��_In, 0) <> 0 Then
          v_���ϼ� := v_���ϼ� + r_Del.��Ԥ��;
          If Nvl(n_����״̬, 0) <> 1 Then
            Update �������
            Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(r_Del.��Ԥ��, 0)
            Where ����id = v_����id And ���� = Nvl(r_Del.Ԥ�����, 2);
            If Sql%NotFound Then
              Insert Into �������
                (����id, ����, Ԥ�����, �������, ����)
              Values
                (v_����id, 1, Nvl(r_Del.��Ԥ��, 0), 0, Nvl(r_Del.Ԥ�����, 2));
            End If;
          End If;
          If r_Del.��¼���� = 1 Then
            Update ����Ԥ����¼ Set ��Ԥ�� = 0 Where ID = r_Del.Id;
          Else
            Delete ����Ԥ����¼ Where ID = r_Del.Id;
          End If;
        End If;
      End If;
    End Loop;
  Else
    For r_Del In c_Del_ҽ�� Loop
      If r_Del.��¼���� Not In (1, 11) Then
        If Nvl(n_����״̬, 0) <> 1 Then
          Update ��Ա�ɿ����
          Set ��� = Nvl(���, 0) - r_Del.��Ԥ��
          Where �տ�Ա = v_����Ա���� And ���� = 1 And ���㷽ʽ = r_Del.���㷽ʽ;
        
          If Sql%RowCount = 0 Then
            Insert Into ��Ա�ɿ����
              (�տ�Ա, ���㷽ʽ, ����, ���)
            Values
              (v_����Ա����, r_Del.���㷽ʽ, 1, -1 * r_Del.��Ԥ��);
          End If;
        End If;
        v_���ϼ� := v_���ϼ� + r_Del.��Ԥ��;
        Delete From ����Ԥ����¼ Where ID = r_Del.Id;
      Else
        --����Ƿ��Ԥ��
        If Nvl(ȱʡ��Ԥ��_In, 0) <> 0 Then
          v_���ϼ� := v_���ϼ� + r_Del.��Ԥ��;
          If Nvl(n_����״̬, 0) <> 1 Then
            Update �������
            Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(r_Del.��Ԥ��, 0)
            Where ����id = v_����id And ���� = Nvl(r_Del.Ԥ�����, 2);
            If Sql%NotFound Then
              Insert Into �������
                (����id, ����, Ԥ�����, �������, ����)
              Values
                (v_����id, 1, Nvl(r_Del.��Ԥ��, 0), 0, Nvl(r_Del.Ԥ�����, 2));
            End If;
          End If;
          If r_Del.��¼���� = 1 Then
            Update ����Ԥ����¼ Set ��Ԥ�� = 0 Where ID = r_Del.Id;
          Else
            Delete ����Ԥ����¼ Where ID = r_Del.Id;
          End If;
        End If;
      End If;
    End Loop;
  End If;

  --------------------------------------------------------------------------------------------------------------
  --------------------------------------------------------------------------------------------------------------
  --����ҽ��֧������
  If ���ս���_In Is Not Null Then
    --�������ս���
    v_���ս��� := ���ս���_In || '||';
    While v_���ս��� Is Not Null Loop
      v_��ǰ���� := Substr(v_���ս���, 1, Instr(v_���ս���, '||') - 1);
    
      v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, '|') - 1);
      v_������ := To_Number(Substr(v_��ǰ����, Instr(v_��ǰ����, '|') + 1));
    
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, ��������)
      Values
        (����Ԥ����¼_Id.Nextval, Decode(����_In, 1, 2, v_��¼����), v_No, 1, v_����id, v_��ҳid, '���ղ���', v_���㷽ʽ, v_�Ǽ�ʱ��, v_����Ա���,
         v_����Ա����, v_������, ����id_In, n_��id, n_�������, Mod(Decode(����_In, 1, 2, v_��¼����), 10));
    
      v_���ϼ� := v_���ϼ� - v_������;
    
      v_���ս��� := Substr(v_���ս���, Instr(v_���ս���, '||') + 2);
    End Loop;
  End If;
  --ʣ�ಿ����Ԥ��
  If Nvl(ȱʡ��Ԥ��_In, 0) <> 0 And v_���ϼ� <> 0 Then
    n_Ԥ����� := v_���ϼ�;
    --�Ƚ�����
    --���������㷽ʽΪ���տ����Ԥ���
    For c_Ԥ�� In (Select a.No, Sum(Nvl(a.���, 0) - Nvl(a.��Ԥ��, 0)) As ���, Nvl(Max(a.����id), 0) As ����id, a.Ԥ�����,
                        Max(Decode(a.��¼����, 1, a.��¼״̬, 1)) As ��¼״̬,
                        Max(Decode(a.��¼����, 1, Decode(a.��¼״̬, 1, a.Id, 3, a.Id, 0), 0)) As ID,
                        Max(Decode(a.��¼����, 1, Decode(a.��¼״̬, 1, a.�տ�ʱ��, 3, a.�տ�ʱ��, Null, Null))) As �տ�ʱ��
                 From ����Ԥ����¼ A
                 Where a.��¼���� In (1, 11) And a.����id = v_����id And Nvl(a.Ԥ�����, 2) = ȱʡ��Ԥ��_In And
                       a.���㷽ʽ Not In (Select ���� From ���㷽ʽ Where ���� = 5)
                 Group By a.No, a.Ԥ�����
                 Having Sum(Nvl(a.���, 0) - Nvl(a.��Ԥ��, 0)) <> 0
                 Order By �տ�ʱ��) Loop
    
      n_��ǰ��� := Case
                  When c_Ԥ��.��� - n_Ԥ����� < 0 Then
                   c_Ԥ��.���
                  Else
                   n_Ԥ�����
                End;
    
      If c_Ԥ��.����id = 0 Then
        --��һ�γ�Ԥ��(����һ�α��Ͻ���ID,��Ԥ�����Ϊ0)
        Update ����Ԥ����¼
        Set ��Ԥ�� = 0, ����id = ����id_In, ������� = n_�������, �������� = Mod(Decode(����_In, 1, 2, v_��¼����), 10)
        Where ID = c_Ԥ��.Id;
      End If;
      --���ϴ�ʣ���
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���,
         ��Ԥ��, ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, ��������)
        Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, Ԥ�����, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�,
               v_�Ǽ�ʱ��, v_����Ա����, v_����Ա���, n_��ǰ���, ����id_In, n_��id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, n_�������,
               Mod(Decode(����_In, 1, 2, v_��¼����), 10)
        From ����Ԥ����¼
        Where NO = c_Ԥ��.No And ��¼״̬ = c_Ԥ��.��¼״̬ And ��¼���� In (1, 11) And Rownum = 1;
    
      --���²���Ԥ�����
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) - n_��ǰ���
      Where ����id = v_����id And ���� = 1 And ���� = Nvl(c_Ԥ��.Ԥ�����, 2);
      --����Ƿ��Ѿ�������
      If c_Ԥ��.��� < n_Ԥ����� Then
        n_Ԥ����� := n_Ԥ����� - c_Ԥ��.���;
      Else
        n_Ԥ����� := 0;
      End If;
    
      If n_Ԥ����� = 0 Then
        Exit;
      End If;
    End Loop;
    If n_Ԥ����� <> 0 Then
      v_Err_Msg := '[ZLSOFT]Ԥ���಻��֧������֧�����,���ܼ���������[ZLSOFT]';
      Raise Err_Item;
    End If;
    Delete From ������� Where ����id = v_����id And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
    v_���ϼ� := n_Ԥ�����;
  End If;

  --ʣ�ಿ��ȫ����ȱʡ���㷽ʽ���㣬(С����Ҳ�����ж��⴦��)
  If v_���ϼ� <> 0 Then
    Update ����Ԥ����¼
    Set ��Ԥ�� = ��Ԥ�� + v_���ϼ�, �����id = �����id_In, ���㿨��� = ���㿨���_In, ���� = ����_In, ������ˮ�� = ������ˮ��_In, ����˵�� = ����˵��_In,
        ������λ = ������λ_In, ������� = n_�������
    
    Where ����id = ����id_In And Nvl(���㷽ʽ, 'LXH_Test') = Nvl(v_ȱʡ, 'LXH_Test') And ��¼���� = Decode(����_In, 1, 2, v_��¼����);
    If Sql%RowCount = 0 Then
      Insert Into ����Ԥ����¼
        (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, �����id, ���㿨���, ����, ������ˮ��,
         ����˵��, ������λ, ��������)
      Values
        (����Ԥ����¼_Id.Nextval, Decode(����_In, 1, 2, v_��¼����), v_No, 1, v_����id, v_��ҳid, '���ս�������', v_ȱʡ, v_�Ǽ�ʱ��, v_����Ա���,
         v_����Ա����, v_���ϼ�, ����id_In, n_��id, n_�������, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, ������λ_In,
         Mod(Decode(����_In, 1, 2, v_��¼����), 10));
    End If;
  
    --�ҺŽ���,�ֱҴ���(���ڹҺŽ���û��Ԥ����,�����ڴ˹����и��ݷֱҴ������������)
    If v_��¼���� = 4 Then
    
      Begin
        Select a.��Ԥ��, a.���㷽ʽ
        Into v_�ֽ���, v_�ֽ����
        From ����Ԥ����¼ A, ���㷽ʽ B
        Where a.���㷽ʽ = b.���� And b.���� = 1 And a.����id = ����id_In And a.��¼���� = 4;
      Exception
        When Others Then
          v_�ֽ��� := 0;
      End;
      If Floor(Abs(v_�ֽ���) * 10) <> Abs(v_�ֽ���) * 10 Then
        --����
        v_Cashcented := Zl_Cent_Money(v_�ֽ���, 1);
        v_�����   := v_�ֽ��� - v_Cashcented;
        If v_����� <> 0 Then
          Begin
            Select ���� Into v_����� From ���㷽ʽ Where ���� = 9;
          Exception
            When Others Then
              v_����� := Null;
          End;
          If v_����� Is Not Null Then
            --10.34֮���������
            Insert Into ����Ԥ����¼
              (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �տ�ʱ��, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, �������, ��������)
            Values
              (����Ԥ����¼_Id.Nextval, Decode(����_In, 1, 2, v_��¼����), v_No, 1, v_����id, v_��ҳid, '����', v_�����, v_�Ǽ�ʱ��, v_����Ա���,
               v_����Ա����, v_�����, ����id_In, n_��id, n_�������, Mod(Decode(����_In, 1, 2, v_��¼����), 10));
            Update ����Ԥ����¼
            Set ��Ԥ�� = v_Cashcented
            Where ����id = ����id_In And ��¼���� = 4 And ���㷽ʽ = v_�ֽ����;
          Else
            --1.����Ԥ����¼(һ�����ڼ�¼)
            Update ����Ԥ����¼
            Set ��Ԥ�� = v_Cashcented
            Where ���㷽ʽ = (Select ���� From ���㷽ʽ Where ���� = 1 And Rownum = 1) And ����id = ����id_In;
          
            --2.���������ü�¼(ע:���㵥λ��¼���Ǻű�,���Բ�ȡ������)
            Begin
              Select a.���, a.Id, c.Id, c.�վݷ�Ŀ
              Into v_�շ����, v_�շ�ϸĿid, v_������Ŀid, v_�վݷ�Ŀ
              From �շ���ĿĿ¼ A, �շѼ�Ŀ B, ������Ŀ C, �շ��ض���Ŀ D
              Where d.�ض���Ŀ = '�����' And d.�շ�ϸĿid = a.Id And a.Id = b.�շ�ϸĿid And b.������Ŀid = c.Id And
                    Sysdate Between b.ִ������ And Nvl(b.��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'));
            Exception
              When Others Then
                v_Err_Msg := '������ȷ��ȡ�շ���������Ϣ�����ȼ�����Ŀ�Ƿ�������ȷ��';
                Raise Err_Item;
            End;
            If Nvl(����_In, 0) = 1 Then
              Select ���˷��ü�¼_Id.Nextval Into v_����id From Dual;
              Select Max(���) + 1, Max(����ʱ��) Into v_���, v_����ʱ�� From סԺ���ü�¼ Where ����id = ����id_In;
              n_ҽ��С��id := Zl_ҽ��С��_Get(0, v_����Ա����, v_����id, v_��ҳid, v_����ʱ��);
            
              Insert Into סԺ���ü�¼
                (ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, �����־, ����id, ��ʶ��, ����, ����, �Ա�, ����, ���˲���id, ���˿���id, �ѱ�, �շ����,
                 �շ�ϸĿid, ���㵥λ, ��ҩ����, ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ���ʷ���, ������, ��������id, ������, ����ʱ��,
                 �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ����id, ���ʽ��, ����Ա���, ����Ա����, �Ƿ��ϴ�, �ɿ���id, ҽ��С��id)
                Select v_����id, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, v_���, Null, Null, �����־, ����id, ��ʶ��, ����, ����, �Ա�, ����, ���˲���id, ���˿���id,
                       �ѱ�, v_�շ����, v_�շ�ϸĿid, ���㵥λ, ��ҩ����, 1, 1, �Ӱ��־, 9, v_������Ŀid, v_�վݷ�Ŀ, v_�����, v_�����, v_�����, ���ʷ���,
                       ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ����id_In, v_�����, ����Ա���, ����Ա����, 1, �ɿ���id,
                       Decode(n_ҽ��С��id, Null, ҽ��С��id, n_ҽ��С��id)
                From סԺ���ü�¼
                Where ����id = ����id_In And Rownum = 1;
            Else
              Select ���˷��ü�¼_Id.Nextval Into v_����id From Dual;
              Select Max(���) + 1 Into v_��� From ������ü�¼ Where ����id = ����id_In;
              Insert Into ������ü�¼
                (ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, �����־, ����id, ��ʶ��, ���ʽ, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid,
                 ���㵥λ, ��ҩ����, ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ���ʷ���, ������, ��������id, ������, ����ʱ��, �Ǽ�ʱ��,
                 ִ�в���id, ִ����, ִ��״̬, ����״̬, ����id, ���ʽ��, ����Ա���, ����Ա����, �Ƿ��ϴ�, �ɿ���id)
                Select v_����id, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, v_���, Null, Null, �����־, ����id, ��ʶ��, ���ʽ, ����, �Ա�, ����, ���˿���id, �ѱ�,
                       v_�շ����, v_�շ�ϸĿid, ���㵥λ, ��ҩ����, 1, 1, �Ӱ��־, 9, v_������Ŀid, v_�վݷ�Ŀ, v_�����, v_�����, v_�����, ���ʷ���, ������,
                       ��������id, ������, ����ʱ��, �Ǽ�ʱ��, ִ�в���id, ִ����, ִ��״̬, ����״̬, ����id_In, v_�����, ����Ա���, ����Ա����, 1, �ɿ���id
                From ������ü�¼
                Where ����id = ����id_In And Rownum = 1;
            End If;
          End If;
          --3.���»��ܱ�
          --ֻ���ܲ��������ı仯.��Ϊ�˱�������������α�
        End If;
      End If;
    End If;
  End If;

  --����ٴ���"��Ա�ɿ����"(û�ж���Ԥ���ǲ���,����"�������"��Ԥ�����ø���)
  For r_Del In c_Del Loop
    If r_Del.��¼���� Not In (1, 11) Then
      If Nvl(n_����״̬, 0) <> 1 Then
        Update ��Ա�ɿ����
        Set ��� = Nvl(���, 0) + r_Del.��Ԥ��
        Where �տ�Ա = v_����Ա���� And ���� = 1 And ���㷽ʽ = r_Del.���㷽ʽ;
        If Sql%RowCount = 0 Then
          Insert Into ��Ա�ɿ����
            (�տ�Ա, ���㷽ʽ, ����, ���)
          Values
            (v_����Ա����, r_Del.���㷽ʽ, 1, r_Del.��Ԥ��);
        End If;
      End If;
    End If;
  End Loop;
  Delete From ��Ա�ɿ���� Where ���� = 1 And �տ�Ա = v_����Ա���� And Nvl(���, 0) = 0;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˽����¼_Update;
/

--89348:������,2017-05-23,ҽ������ӡ
--109076:������,2017-05-12,ҽ������ӡ
Create Or Replace Procedure Zl_����ҽ����ӡ_Insert
(
  ����id_In ����ҽ����¼.����id%Type,
  ��ҳid_In ����ҽ����¼.��ҳid%Type,
  Ӥ��_In   ����ҽ����¼.Ӥ��%Type,
  ��Ч_In   ����ҽ����¼.ҽ����Ч%Type,
  ����_In   Number
  --���ܣ�������û�д�ӡ����ҽ������ ����ҽ����ӡ
  --����������_In������ҽ����һҳ���Դ������
  --      ����_Inҽ���������������ͨ����28�С�
) Is
  n_���       ����ҽ����¼.���%Type;
  n_ҽ��id     ����ҽ����¼.Id%Type;
  n_�������   Number;
  v_Max_Date   Date;
  d_����       Date;
  d_Pdate      Date;
  n_��ҳ��     Number;
  n_���ؿ�     Number;
  n_ת��       Number;
  n_ҳ��       Number;
  n_�к�       Number;
  n_λ��       Number;
  n_��ӡģʽ   Number;
  n_���ҩ��ʽ Number;
  n_Lzzkhy     Number;
  n_Cnt        Number;
  v_Tmp        Varchar2(200);

  --c_Advice ȡ������ӡ��ҽ�����ڴ�ӡ����ʱת��ҽ�������ȡ����������Ҫ�ж��ǲ���Ҫ���ɴ�ӡ��¼
  Cursor c_Advice Is
    Select ҽ��id, ˳��, ��ӡ���, ����ҽ��, ��ҳ
    From (With Printtable As (Select a.Id As ҽ��id, a.��� As ˳��, 0 As ��ӡ���, Null As ����ҽ��,
                                     Decode(a.�������, 'Z', Decode(b.��������, '3', 3, '4', 4, 0), 0) As ��ҳ, a.������Ŀid
                              From ����ҽ����¼ A, ������ĿĿ¼ B
                              Where a.����id = ����id_In And a.��ҳid = ��ҳid_In And Nvl(a.Ӥ��, 0) = Ӥ��_In And a.������Ŀid = b.Id(+) And
                                    (��Ч_In = 0 And (a.ҽ����Ч = 0 Or n_λ�� In (-1, 0, 2) And a.ҽ����Ч = 1 And a.������� = 'Z' And
                                    b.�������� In ('5', '3', '11')) Or
                                    ��Ч_In = 1 And a.ҽ����Ч = 1 And
                                    Not (n_λ�� = 0 And Nvl(a.�������, 'X') = 'Z' And Nvl(b.��������, 'X') In ('5', '3', '11')) Or
                                    ��Ч_In = 1 And a.ҽ����Ч = 1 And n_λ�� = 0 And a.������� = 'Z' And b.�������� = '3') And
                                    a.ҽ��״̬ Not In (-1, 2) And (n_��ӡģʽ = 1 And a.ҽ��״̬ = 1 Or a.ҽ��״̬ <> 1) And
                                    Nvl(a.���δ�ӡ, 0) = 0 And a.��� > n_��� And a.������Դ = 2)
           Select p.ҽ��id, p.˳��, p.��ӡ���, p.����ҽ��, p.��ҳ
           From ����ҽ����¼ L, ������ĿĿ¼ I, Printtable P
           Where l.Id = p.ҽ��id And l.������Ŀid = i.Id And
                 (l.������� Not In ('5', '6', '7', 'E') Or l.������� = 'E' And Nvl(i.��������, '0') Not In ('2', '3') Or
                 i.Id Is Null) And l.���id Is Null
           Union All
           Select p.ҽ��id, p.˳��, p.��ӡ���, p.����ҽ��, p.��ҳ
           From ����ҽ����¼ L, Printtable P
           Where l.Id = p.ҽ��id And l.������� In ('5', '6')
           Union All
           Select p.ҽ��id, p.˳��, p.��ӡ���, p.����ҽ��, p.��ҳ
           From ����ҽ����¼ L, ������ĿĿ¼ I, Printtable P
           Where l.Id = p.ҽ��id And l.������Ŀid = i.Id And l.������� = 'E' And i.�������� = '2' And l.���id Is Null And n_���ҩ��ʽ = 1
           Union All
           Select p.ҽ��id, p.˳��, p.��ӡ���, p.����ҽ��, p.��ҳ
           From Printtable P
           Where p.������Ŀid Is Null
           Order By ˳��);


  Cursor c_Advice_Redo Is
    Select ҽ��id, ˳��, ��ӡ���, ����ҽ��, ��ҳ
    From (With Printtable As (Select a.Id As ҽ��id, a.��� As ˳��, 0 As ��ӡ���, Null As ����ҽ��,
                                     Decode(a.�������, 'Z', Decode(b.��������, '3', 3, '4', 4, 0), 0) As ��ҳ, a.������Ŀid
                              From ����ҽ����¼ A, ������ĿĿ¼ B
                              Where a.����id = ����id_In And a.��ҳid = ��ҳid_In And Nvl(a.Ӥ��, 0) = Ӥ��_In And a.������Ŀid = b.Id(+) And
                                    (��Ч_In = 0 And (a.ҽ����Ч = 0 Or n_λ�� In (-1, 0, 2) And a.ҽ����Ч = 1 And a.������� = 'Z' And
                                    b.�������� In ('5', '3', '11')) Or
                                    ��Ч_In = 1 And a.ҽ����Ч = 1 And
                                    Not (n_λ�� = 0 And Nvl(a.�������, 'X') = 'Z' And Nvl(b.��������, 'X') In ('5', '3', '11'))) And
                                    a.ҽ��״̬ Not In (-1, 2) And (n_��ӡģʽ = 1 And a.ҽ��״̬ = 1 Or a.ҽ��״̬ <> 1) And
                                    Nvl(a.���δ�ӡ, 0) = 0 And a.��� > n_��� And Exists
                               (Select 1 From ����ҽ��״̬ C Where a.Id = c.ҽ��id And c.����ʱ�� >= v_Max_Date) And a.������Դ = 2)
           Select p.ҽ��id, p.˳��, p.��ӡ���, p.����ҽ��, p.��ҳ
           From ����ҽ����¼ L, ������ĿĿ¼ I, Printtable P
           Where l.Id = p.ҽ��id And l.������Ŀid = i.Id And
                 (l.������� Not In ('5', '6', '7', 'E') Or l.������� = 'E' And Nvl(i.��������, '0') Not In ('2', '3') Or
                 i.Id Is Null) And l.���id Is Null
           Union All
           Select p.ҽ��id, p.˳��, p.��ӡ���, p.����ҽ��, p.��ҳ
           From ����ҽ����¼ L, Printtable P
           Where l.Id = p.ҽ��id And l.������� In ('5', '6')
           Union All
           Select p.ҽ��id, p.˳��, p.��ӡ���, p.����ҽ��, p.��ҳ
           From ����ҽ����¼ L, ������ĿĿ¼ I, Printtable P
           Where l.Id = p.ҽ��id And l.������Ŀid = i.Id And l.������� = 'E' And i.�������� = '2' And l.���id Is Null And n_���ҩ��ʽ = 1
           Union All
           Select p.ҽ��id, p.˳��, p.��ӡ���, p.����ҽ��, p.��ҳ
           From Printtable P
           Where p.������Ŀid Is Null
           Order By ˳��);


  --��ȡ��һ�����õ��кź�ҳ��
  Function Getnextpos
  (
    v_ҳ�� ����ҽ����ӡ.ҳ��%Type,
    v_�к� ����ҽ����ӡ.�к�%Type,
    v_���� Number
  ) Return Varchar2 Is
    n_p Number;
    n_r Number;
  Begin
    If v_�к� = 0 Then
      n_p := 1;
      n_r := 1;
    Elsif v_�к� = v_���� Then
      n_p := v_ҳ�� + 1;
      n_r := 1;
    Else
      n_p := v_ҳ��;
      n_r := v_�к� + 1;
    End If;
    Return(n_p || ',' || n_r);
  End;

Begin
  n_λ��       := Zl_To_Number(Nvl(zl_GetSysParameter('ת�ƺͳ�Ժ��ӡ', 1254), 0));
  n_��ӡģʽ   := Zl_To_Number(Nvl(zl_GetSysParameter('ҽ������ӡģʽ', 1253), 0));
  n_���ҩ��ʽ := Zl_To_Number(Nvl(zl_GetSysParameter('ҩƷ�÷�������ӡһ��', 1254), 0));
  n_Lzzkhy     := Zl_To_Number(Nvl(zl_GetSysParameter('������ת�ƻ�ҳ', 1254), 0));
  n_��ҳ��     := Zl_To_Number(Nvl(zl_GetSysParameter('����������ҽ����ҳ��ӡ', 1254), 0));
  n_���ؿ�     := Zl_To_Number(Nvl(zl_GetSysParameter('ת�ƻ�ҳ�������д�ӡ�ؿ�ҽ��', 1254), 0));

  --�ж��ǲ����������ӡҽ��
  If ��Ч_In = 1 Then
    d_���� := To_Date('1900-01-01', 'YYYY-MM-DD');
  Else
    Select ҽ������ʱ�� Into d_���� From ������ҳ Where ����id = ����id_In And ��ҳid = ��ҳid_In;
    If d_���� Is Null Then
      d_���� := To_Date('1900-01-01', 'YYYY-MM-DD');
    End If;
  End If;
  v_Max_Date := d_����;
  Begin
    Select ҽ��id, ��ӡʱ��, ҳ��, �к�
    Into n_ҽ��id, d_Pdate, n_ҳ��, n_�к�
    From (Select ҽ��id, ��ӡʱ��, ҳ��, �к�
           From ����ҽ����ӡ
           Where ����id = ����id_In And ��ҳid = ��ҳid_In And Nvl(Ӥ��, 0) = Ӥ��_In And ��Ч = ��Ч_In And ҽ��id Is Not Null
           Order By ҳ�� Desc, �к� Desc)
    Where Rownum < 2;
  
    Select Nvl(Max(���), 0)
    Into n_���
    From ����ҽ����¼
    Where ID = (Select Nvl(a.���id, a.Id) From ����ҽ����¼ A Where a.Id = n_ҽ��id);
  
    If ��Ч_In = 0 Then
      If d_Pdate Is Not Null Then
        If d_Pdate < d_���� And d_���� <> To_Date('1900-01-01', 'YYYY-MM-DD') Then
          n_������� := 1;
          n_���     := 0;
        End If;
      End If;
    End If;
  Exception
    When Others Then
      n_ҳ�� := 0;
      n_�к� := 0;
      n_��� := 0;
  End;

  If n_ҽ��id Is Not Null Then
    Select Max(b.��������)
    Into v_Tmp
    From ����ҽ����¼ A, ������ĿĿ¼ B
    Where a.������Ŀid = b.Id(+) And a.Id = n_ҽ��id And a.������� = 'Z';
  End If;
  If v_Tmp = '3' Then
    n_Cnt := 3;
  Elsif v_Tmp = '4' Then
    n_Cnt := 4;
  End If;

  v_Tmp  := Getnextpos(n_ҳ��, n_�к�, ����_In);
  n_ҳ�� := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
  n_�к� := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);

  If n_Cnt = 3 And n_Lzzkhy = 1 And ��Ч_In = 1 Then
    --��ʱҽ��ת�ƻ�ҳ
    If n_�к� <> 1 Then
      n_�к� := 1;
      n_ҳ�� := n_ҳ�� + 1;
    End If;
  Elsif ��Ч_In = 0 Then
    --����������ת���ؿ�����Щֻ����ڳ���ҽ����
    --�������
    If n_������� = 1 Then
      If n_��ҳ�� = 1 Then
        If n_�к� <> 1 Then
          n_�к� := 1;
          n_ҳ�� := n_ҳ�� + 1;
        End If;
      End If;
      Insert Into ����ҽ����ӡ
        (ҽ��id, ҳ��, �к�, ����, ����id, ��ҳid, Ӥ��, ��Ч, ��ӡ���, ����ҽ��)
      Values
        (-1 * Null, n_ҳ��, n_�к�, 1, ����id_In, ��ҳid_In, Ӥ��_In, ��Ч_In, 0, Null);
      v_Tmp  := Getnextpos(n_ҳ��, n_�к�, ����_In);
      n_ҳ�� := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
      n_�к� := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
    End If;
  
    --ת�ƻ�ҳ���ؿ�����
    If n_���ؿ� = 1 And n_Cnt = 3 Then
      If n_������� = 1 Then
        --ǰ����������Ͳ���ҳ��
        If n_���ؿ� = 1 Then
          Insert Into ����ҽ����ӡ
            (ҽ��id, ҳ��, �к�, ����, ����id, ��ҳid, Ӥ��, ��Ч, ��ӡ���, ����ҽ��)
          Values
            (-1 * Null, n_ҳ��, n_�к�, 1, ����id_In, ��ҳid_In, Ӥ��_In, ��Ч_In, 0, 1);
          v_Tmp  := Getnextpos(n_ҳ��, n_�к�, ����_In);
          n_ҳ�� := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
          n_�к� := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
        End If;
      Else
        --���ؿ�����
        If n_�к� <> 1 Then
          n_�к� := 1;
          n_ҳ�� := n_ҳ�� + 1;
        End If;
        Insert Into ����ҽ����ӡ
          (ҽ��id, ҳ��, �к�, ����, ����id, ��ҳid, Ӥ��, ��Ч, ��ӡ���, ����ҽ��)
        Values
          (-1 * Null, n_ҳ��, n_�к�, 1, ����id_In, ��ҳid_In, Ӥ��_In, ��Ч_In, 0, 1);
        v_Tmp  := Getnextpos(n_ҳ��, n_�к�, ����_In);
        n_ҳ�� := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
        n_�к� := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
      End If;
    End If;
  
    --ת��ҽ����ҳ
    If Nvl(n_�������, 0) <> 1 And Nvl(n_���ؿ�, 0) <> 1 And n_Cnt = 3 And n_��ҳ�� = 1 Then
      If n_�к� <> 1 Then
        n_�к� := 1;
        n_ҳ�� := n_ҳ�� + 1;
      End If;
    End If;
  
    --����ҽ����ҳ
    If Nvl(n_�������, 0) <> 1 And n_Cnt = 4 And n_��ҳ�� = 1 Then
      If n_�к� <> 1 Then
        n_�к� := 1;
        n_ҳ�� := n_ҳ�� + 1;
      End If;
    End If;
  End If;
  n_ת�� := 0;

  --�����������,��Ҫ��ӡ��ҽ�������ǻ�ҳ��ӡ���ת������
  ---r_Print.��ҳ ������ҽ����ǣ�4������3��ת��
  If v_Max_Date = To_Date('1900-01-01', 'YYYY-MM-DD') Then
    For r_Print In c_Advice Loop
      ----��ҳ���ߴ�ҽ���ؿ�����
      If n_��ҳ�� = 1 And n_ת�� = 1 And ��Ч_In = 0 Then
        If n_���ؿ� = 1 Then
          --���ؿ�����
          If n_�к� <> 1 Then
            n_�к� := 1;
            n_ҳ�� := n_ҳ�� + 1;
          End If;
          Insert Into ����ҽ����ӡ
            (ҽ��id, ҳ��, �к�, ����, ����id, ��ҳid, Ӥ��, ��Ч, ��ӡ���, ����ҽ��)
          Values
            (-1 * Null, n_ҳ��, n_�к�, 1, ����id_In, ��ҳid_In, Ӥ��_In, ��Ч_In, 0, 1);
          v_Tmp  := Getnextpos(n_ҳ��, n_�к�, ����_In);
          n_ҳ�� := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
          n_�к� := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
        Else
          --ֻ�ǵ�����һҳ
          If n_�к� <> 1 Then
            n_�к� := 1;
            n_ҳ�� := n_ҳ�� + 1;
          End If;
        End If;
        n_ת�� := 0;
      End If;
    
      If ��Ч_In = 1 And n_ת�� = 1 And n_Lzzkhy = 1 Then
        If n_�к� <> 1 Then
          n_�к� := 1;
          n_ҳ�� := n_ҳ�� + 1;
        End If;
        n_ת�� := 0;
      End If;
    
      If r_Print.��ҳ = 4 And n_��ҳ�� = 1 Then
        --����ҽ����ҳ
        --����к�Ϊ1˵���Ѿ����µ�һҳ�ĵ�һ��,����ҳ
        If n_�к� <> 1 Then
          n_�к� := 1;
          n_ҳ�� := n_ҳ�� + 1;
        End If;
      End If;
    
      If ��Ч_In = 0 Or ��Ч_In = 1 And (n_λ�� = 2 Or n_λ�� = 1 Or r_Print.��ҳ <> 3) Then
        Insert Into ����ҽ����ӡ
          (ҽ��id, ҳ��, �к�, ����, ����id, ��ҳid, Ӥ��, ��Ч, ��ӡ���, ����ҽ��)
        Values
          (r_Print.ҽ��id, n_ҳ��, n_�к�, 1, ����id_In, ��ҳid_In, Ӥ��_In, ��Ч_In, r_Print.��ӡ���, r_Print.����ҽ��);
        v_Tmp  := Getnextpos(n_ҳ��, n_�к�, ����_In);
        n_ҳ�� := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
        n_�к� := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
      End If;
      --������ת�ƻ�ҳ���ؿ������������һ�ؿ�ҽ����ǣ��˴�һ����ҳ����Ϊת�ƻ�ҳǰҪ�ȴ��ת��ҽ����
      --���ﲻ�������ݣ�ֻ���б�ǣ�����һ��ѭʱ�Ų��롣���ת��ҽ�������һ���ǲ��ô�ӡ�¿������ġ�
      If r_Print.��ҳ = 3 Then
        n_ת�� := 1;
      End If;
    End Loop;
  Else
    For r_Print In c_Advice_Redo Loop
      ----��ҳ���ߴ�ҽ���ؿ�����
      If n_��ҳ�� = 1 And n_ת�� = 1 And ��Ч_In = 0 Then
        If n_���ؿ� = 1 Then
          --���ؿ�����
          If n_�к� <> 1 Then
            n_�к� := 1;
            n_ҳ�� := n_ҳ�� + 1;
          End If;
          Insert Into ����ҽ����ӡ
            (ҽ��id, ҳ��, �к�, ����, ����id, ��ҳid, Ӥ��, ��Ч, ��ӡ���, ����ҽ��)
          Values
            (-1 * Null, n_ҳ��, n_�к�, 1, ����id_In, ��ҳid_In, Ӥ��_In, ��Ч_In, 0, 1);
          v_Tmp  := Getnextpos(n_ҳ��, n_�к�, ����_In);
          n_ҳ�� := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
          n_�к� := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
        Else
          --ֻ�ǵ�����һҳ
          If n_�к� <> 1 Then
            n_�к� := 1;
            n_ҳ�� := n_ҳ�� + 1;
          End If;
        End If;
        n_ת�� := 0;
      End If;
    
      If r_Print.��ҳ = 4 And n_��ҳ�� = 1 Then
        --����ҽ����ҳ
        --����к�Ϊ1˵���Ѿ����µ�һҳ�ĵ�һ��,����ҳ
        If n_�к� <> 1 Then
          n_�к� := 1;
          n_ҳ�� := n_ҳ�� + 1;
        End If;
      End If;
      Insert Into ����ҽ����ӡ
        (ҽ��id, ҳ��, �к�, ����, ����id, ��ҳid, Ӥ��, ��Ч, ��ӡ���, ����ҽ��)
      Values
        (r_Print.ҽ��id, n_ҳ��, n_�к�, 1, ����id_In, ��ҳid_In, Ӥ��_In, ��Ч_In, r_Print.��ӡ���, r_Print.����ҽ��);
      v_Tmp  := Getnextpos(n_ҳ��, n_�к�, ����_In);
      n_ҳ�� := Substr(v_Tmp, 1, Instr(v_Tmp, ',') - 1);
      n_�к� := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
      --������ת�ƻ�ҳ���ؿ������������һ�ؿ�ҽ����ǣ��˴�һ����ҳ����Ϊת�ƻ�ҳǰҪ�ȴ��ת��ҽ��
      --���ﲻ�������ݣ�ֻ���б�ǣ�����һ��ѭʱ�Ų��롣���ת��ҽ�������һ���ǲ��ô�ӡ�¿������ġ�
    
      If r_Print.��ҳ = 3 And ��Ч_In = 0 Then
        n_ת�� := 1;
      End If;
    End Loop;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ҽ����ӡ_Insert;
/


--95807:����,2017-05-12,��¼����������Ŀ����δ��˵����Ϣ¼��

CREATE OR REPLACE Procedure Zl_���˻�������_Update
(
  �ļ�id_In   In ���˻�������.�ļ�id%Type,
  ����ʱ��_In In ���˻�������.����ʱ��%Type,
  ��¼����_In In ���˻�����ϸ.��¼����%Type, --������Ŀ=1��ǩ����¼=5����ǩ��¼=15
  ��Ŀ���_In In ���˻�����ϸ.��Ŀ���%Type, --������Ŀ����ţ��ǻ�����Ŀ�̶�Ϊ0
  ��¼����_In In ���˻�����ϸ.��¼����%Type := Null, --��¼���ݣ��������Ϊ�գ��������ǰ�����ݣ�37��38/37
  ���²�λ_In In ���˻�����ϸ.���²�λ%Type := Null,
  ���˼�¼_In In Number := 1,
  ������Դ_In In ���˻�����ϸ.������Դ%Type := 0,
  ��ǩ_In     In Number := 0,
  ����Ա_In   In ���˻�������.������%Type := Null,
  ��¼���_In In ���˻�����ϸ.��¼���%Type := Null, --���÷������(һ�����ݶ�Ӧ������ͬ��Ŀ����ϸ)
  ������_In In ���˻�����ϸ.������%Type := Null, --���÷������(��¼������Ŀ������������Ŀ���)
  δ��˵��_In In ���˻�����ϸ.δ��˵��%Type := Null --��������洢ҽ��ID:���ͺ�
) Is
  Intins      Number(18);
  Int����     Number(1);
  n_Newid     ���˻�������.Id%Type;
  n_Oldid     ���˻�������.Id%Type;
  n_����      ���˻����ӡ.����%Type;
  n_Mutilbill Number(1);
  n_Syntend   Number(1);
  n_Synchro   Number(1);
  n_δ��˵��  Number(1);
  n_����      Number(1);
  n_Num       Number(18);

  n_�������     ���˻�������.�������%Type;
  v_����id       ���ű�.Id%Type;
  v_������       ��Ա��.����%Type;
  v_��¼��       ��Ա��.����%Type;
  n_�ļ�id       ���˻�������.�ļ�id%Type;
  n_��¼id       ���˻�������.Id%Type;
  n_��ϸid       ���˻�����ϸ.Id%Type;
  n_��Դid       ���˻�����ϸ.��Դid%Type;
  v_������Դ     ���˻�����ϸ.������Դ%Type;
  n_��߰汾     ���˻�����ϸ.��ʼ�汾%Type;
  n_��Ŀ����     �����¼��Ŀ.��Ŀ����%Type;
  n_����id       ���˻����ļ�.����id%Type;
  n_��ҳid       ���˻����ļ�.��ҳid%Type;
  n_Ӥ��         ���˻����ļ�.Ӥ��%Type;
  d_Ӥ����Ժʱ�� ����ҽ����¼.��ʼִ��ʱ��%Type;
  d_�ļ���ʼʱ�� ���˻����ļ�.��ʼʱ��%Type;
  --��ȡ�ò��˵�ǰ��������δ�����Ļ����ļ������ļ���ʼʱ��С�ڵ��ڼ�¼����ʱ����ļ��б�ͬ������ʹ��
  Cursor Cur_Fileformats Is
    Select a.Id As ��ʽid, b.Id As �ļ�id, a.����, a.����, b.Ӥ��
    From �����ļ��б� A, ���˻����ļ� B, ���˻����ļ� C, ���˻������� D
    Where a.���� = 3 And a.���� <> 1 And a.Id = b.��ʽid And b.Id <> c.Id And b.����ʱ�� Is Null And b.��ʼʱ�� <= d.����ʱ�� And
          (a.ͨ�� = 1 Or (a.ͨ�� = 2 And b.����id = c.����id)) And c.����id = b.����id And c.��ҳid = b.��ҳid And c.Ӥ�� = b.Ӥ�� And
          c.Id = d.�ļ�id And d.Id = n_��¼id And c.Id = �ļ�id_In
    Order By a.���;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --ȡ��¼ID
  Int����     := 0;
  n_��¼id    := 0;
  n_Mutilbill := 0;
  n_Syntend   := 0;
  n_δ��˵��  := 0;
  n_����      := 0;

  If ����Ա_In Is Null Then
    v_������ := Zl_Username;
  Else
    v_������ := ����Ա_In;
  End If;

  --����Ƕ�Ӧ��ݻ����ļ�ֵΪ1����ʾ��ͬ�����������ļ������򲻴����ļ�ͬ��
  n_Mutilbill := Zl_To_Number(zl_GetSysParameter('��Ӧ��ݻ����ļ�', 1255));
  --��������ݻ����ļ�֮������ͬ��,���Զ�ͬ��,����ͬ��
  n_Syntend := Zl_To_Number(zl_GetSysParameter('��������ͬ��', 1255));

  Begin
    Select ID, �������
    Into n_��¼id, n_�������
    From ���˻�������
    Where �ļ�id = �ļ�id_In And ����ʱ�� = ����ʱ��_In;
  Exception
    When Others Then
      n_��¼id := 0;
  End;

  --����ǲ��Ǳ��˵ļ�¼
  ---------------------------------------------------------------------------------------------------------------------
  If ���˼�¼_In = 0 And n_��¼id > 0 And ��ǩ_In = 0 Then
    v_��¼�� := '';
    Begin
      Select ��¼��
      Into v_��¼��
      From ���˻�����ϸ
      Where ��¼id = n_��¼id And ��Ŀ��� = ��Ŀ���_In And Nvl(���²�λ, 'TWBW') = Nvl(���²�λ_In, 'TWBW') And
            Nvl(��¼���, 0) = Nvl(��¼���_In, 0) And ��ֹ�汾 Is Null;
    Exception
      When Others Then
        v_��¼�� := '';
    End;
    If v_��¼�� Is Not Null And v_��¼�� <> v_������ Then
      v_Error := '����Ȩ�޸����˵ǼǵĻ������ݣ�';
      Raise Err_Custom;
    End If;
  End If;

  --����Ƿ����
  Select ����id, ��ҳid, Nvl(Ӥ��, 0), ��ʼʱ��
  Into n_����id, n_��ҳid, n_Ӥ��, d_�ļ���ʼʱ��
  From ���˻����ļ�
  Where ID = �ļ�id_In;
  d_Ӥ����Ժʱ�� := Null;
  If n_Ӥ�� <> 0 Then
    Begin
      Select ��ʼִ��ʱ��
      Into d_Ӥ����Ժʱ��
      From ����ҽ����¼ B, ������ĿĿ¼ C
      Where b.������Ŀid + 0 = c.Id And b.ҽ��״̬ = 8 And Nvl(b.Ӥ��, 0) <> 0 And c.��� = 'Z' And
            Instr(',3,5,11,', ',' || c.�������� || ',', 1) > 0 And b.����id = n_����id And b.��ҳid = n_��ҳid And b.Ӥ�� = n_Ӥ��;
    Exception
      When Others Then
        d_Ӥ����Ժʱ�� := Null;
    End;
  End If;
  If d_Ӥ����Ժʱ�� Is Null Then
    v_����id := 0;
    Begin
      Select a.����id
      Into v_����id
      From ���˱䶯��¼ A, ���˻����ļ� B
      Where a.����id Is Not Null And a.����id = b.����id And a.��ҳid = b.��ҳid And b.Id = �ļ�id_In And
            (To_Date(To_Char(����ʱ��_In, 'YYYY-MM-DD HH24:MI') || '59', 'YYYY-MM-DD HH24:MI:SS') >= a.��ʼʱ�� And
            (To_Date(To_Char(����ʱ��_In, 'YYYY-MM-DD HH24:MI') || '00', 'YYYY-MM-DD HH24:MI:SS') < = Nvl(a.��ֹʱ��, Sysdate) Or
            a.��ֹʱ�� Is Null)) And Rownum < 2;
    Exception
      When Others Then
        v_����id := 0;
    End;
    If v_����id = 0 Then
      v_Error := '���ݷ���ʱ�� ' || To_Char(����ʱ��_In, 'YYYY-MM-DD HH24:MI:SS') || ' ���ڲ�����Ч�䶯ʱ�䷶Χ�ڣ����ܲ�����';
      Raise Err_Custom;
    End If;
  Else
    If ����ʱ��_In < d_�ļ���ʼʱ�� Or ����ʱ��_In > d_Ӥ����Ժʱ�� Then
      v_Error := '���ݷ���ʱ�� ' || To_Char(����ʱ��_In, 'YYYY-MM-DD HH24:MI:SS') || ' ���ڲ�����Ч�䶯ʱ�䷶Χ�ڣ����ܲ�����';
      Raise Err_Custom;
    End If;
  End If;

  --���������Դ<>0���˳�
  n_��Դid := 0;
  If n_��¼id > 0 Then
    Begin
      Select ������Դ, Nvl(��Դid, 0)
      Into v_������Դ, n_��Դid
      From ���˻�����ϸ
      Where ��¼id = n_��¼id And Nvl(��Ŀ���, 0) = ��Ŀ���_In And Nvl(���²�λ, 'TWBW') = Nvl(���²�λ_In, 'TWBW') And
            Nvl(��¼���, 0) = Nvl(��¼���_In, 0);
    Exception
      When Others Then
        v_������Դ := 0;
    End;
    If v_������Դ > 0 And n_��Դid > 0 Then
      Return;
    End If;
  End If;

  --ȡ��߰汾
  Select Nvl(Max(Nvl(a.��ʼ�汾, 1)), 0) + 1, Count(b.Id)
  Into n_��߰汾, Intins
  From ���˻�����ϸ A, ���˻������� B
  Where b.Id = n_��¼id And a.��¼id = b.Id And Mod(a.��¼����, 10) = 5;

  --Ŀǰ�Ѿ�ǩ�������ݲ����޸ģ�ֻ������ǩģʽ�½����޸ģ�����ǩ_In=1
  If ��ǩ_In <> 1 And Intins > 0 Then
    v_Error := '����ʱ�� ' || To_Char(����ʱ��_In, 'YYYY-MM-DD HH24:MI:SS') || ' ����Ӧ�������Ѿ�ǩ������ǩ�����ܼ���������' || Chr(13) || Chr(10) ||
               '��������������粢����������ģ���ˢ�º����ԣ�';
    Raise Err_Custom;
  End If;
  Intins := 0;

  --������ʱ,Ҫ������ݣ���ǩ����ʱ���Զ������ǩ�������޸ĵ����ݣ����Դ˴�ֻ�迼����ǩ���ɣ�
  If ��¼����_In Is Null Then
    Begin
      Select ID
      Into n_��ϸid
      From ���˻�����ϸ
      Where ��¼id = n_��¼id And Nvl(��Ŀ���, 0) = ��Ŀ���_In And Nvl(���²�λ, 'TWBW') = Nvl(���²�λ_In, 'TWBW') And
            Nvl(��¼���, 0) = Nvl(��¼���_In, 0) And ��ֹ�汾 Is Null;
    Exception
      --�������˳�
      When Others Then
        Return;
    End;

    --���ҳ��˱���Ҫɾ�������ݣ��Ƿ񻹴�������Ч�����ݣ��������ֻɾ���������ݣ�����ɾ���˷���ʱ���Ӧ���������ݡ�
    Select Count(ID)
    Into Intins
    From ���˻�����ϸ
    Where ��¼id = n_��¼id And Mod(��¼����, 10) <> 5 And ��ֹ�汾 Is Null And ID <> n_��ϸid;
    If Intins = 0 Then
      Delete From ���˻�����ϸ Where ��¼id = n_��¼id;
    Else
      Delete From ���˻�����ϸ Where ID = n_��ϸid;
    End If;

    Delete From ���˻������� A
    Where a.Id = n_��¼id And Not Exists (Select 1 From ���˻�����ϸ B Where b.��¼id = a.Id);

    --�����ɾ��ǩ�����޸Ĳ��������һ������,��Ӧ��ǩ����¼����ֹ�汾��Ϊ��
    Begin
      Select 1
      Into Intins
      From ���˻�����ϸ
      Where ��ʼ�汾 = n_��߰汾 And ��ֹ�汾 Is Null And ��¼���� = 1 And ��¼id = n_��¼id;
    Exception
      When Others Then
        Intins := 0;
    End;
    If Intins = 0 Then
      Update ���˻�����ϸ Set ��ֹ�汾 = Null Where ��¼���� = 5 And ��ʼ�汾 = n_��߰汾 - 1 And ��¼id = n_��¼id;
    End If;
    If Nvl(n_�������, 0) <> 0 Then
      Return;
    End If;

    --############
    --�����������
    --############
    For Rsdel In (Select Distinct ��¼id From ���˻�����ϸ Where ��Դid = n_��ϸid) Loop

      Delete ���˻�����ϸ Where ��Դid = n_��ϸid And ��¼id = Rsdel.��¼id;
      --ɾ����Ӧ�Ĵ�ӡ����
      Begin
        Select Count(*) Into Intins From ���˻�����ϸ Where ��¼id = Rsdel.��¼id;
      Exception
        When Others Then
          Intins := 0;
      End;
      If Intins = 0 Then
        --��ȡ������ݶ�Ӧ���ļ�ID
        Begin
          Select b.Id, a.����
          Into n_�ļ�id, Intins
          From �����ļ��б� A, ���˻����ļ� B, ���˻������� C
          Where a.Id = b.��ʽid And b.Id = c.�ļ�id And c.Id = Rsdel.��¼id;
        Exception
          When Others Then
            n_�ļ�id := 0;
        End;
        Delete ���˻������� Where ID = Rsdel.��¼id;
        If Intins <> -1 Then
          Zl_���˻����ӡ_Update(n_�ļ�id, ����ʱ��_In, 1, 1);
        End If;
      End If;
    End Loop;
  Else
    --���¼�����Ŀ�Ƿ����ڸü�¼��
    Begin
      Select 1
      Into Intins
      From (Select b.��Ŀ���
             From �����ļ��ṹ A, �����¼��Ŀ B
             Where a.Ҫ������ = b.��Ŀ���� And b.��Ŀ��� = ��Ŀ���_In And
                   ��id = (Select b.Id
                          From ���˻����ļ� A, �����ļ��ṹ B
                          Where a.Id = �ļ�id_In And a.��ʽid = b.�ļ�id And b.��id Is Null And b.������� = 4)
             Union
             Select ��Ŀ���
             From �����¼��Ŀ
             Where ��Ŀ���� = 2 And ��Ŀ��� = ��Ŀ���_In);
    Exception
      When Others Then
        Intins := 0;
    End;
    If Intins = 0 Then
      Return;
    End If;
    If n_��¼id = 0 Then
      Select ���˻�������_Id.Nextval Into n_��¼id From Dual;

      Insert Into ���˻�������
        (ID, �ļ�id, ����ʱ��, ���汾, ������, ����ʱ��)
      Values
        (n_��¼id, �ļ�id_In, ����ʱ��_In, n_��߰汾, v_������, Sysdate);
    End If;

    --���뱾�εǼǵĲ��˻�����ϸ
    Update ���˻�����ϸ
    Set ��¼���� = ��¼����_In, ������Դ = ������Դ_In, δ��˵�� = δ��˵��_In, ��¼�� = v_������, ��¼ʱ�� = Sysdate
    Where ��¼id = n_��¼id And ��Ŀ��� = ��Ŀ���_In And Nvl(���²�λ, 'TWBW') = Nvl(���²�λ_In, 'TWBW') And
          Nvl(��¼���, 0) = Nvl(��¼���_In, 0) And ��ʼ�汾 = n_��߰汾 And ��ֹ�汾 Is Null;
    If Sql%RowCount = 0 Then
      Select ���˻�����ϸ_Id.Nextval Into n_��ϸid From Dual;
      Insert Into ���˻�����ϸ
        (ID, ��¼id, ��¼����, ��Ŀ����, ��Ŀid, ������, ��Ŀ���, ��Ŀ����, ��Ŀ����, ��¼����, ��Ŀ��λ, ��¼���, ��¼���, ���²�λ, ������Դ, ����, δ��˵��, ��ʼ�汾, ��ֹ�汾,
         ��¼��, ��¼ʱ��)
        Select n_��ϸid, n_��¼id, ��¼����_In, a.������, a.��Ŀid, ������_In, a.��Ŀ���, Upper(a.��Ŀ����), a.��Ŀ����, ��¼����_In, a.��Ŀ��λ, 0,
               ��¼���_In, ���²�λ_In, ������Դ_In, Nvl(b.����, 0), δ��˵��_In, n_��߰汾, Null, v_������, Sysdate
        From �����¼��Ŀ A, ���˻�����ϸ B
        Where a.��Ŀ��� = b.��Ŀ���(+) And b.��ֹ�汾(+) Is Null And b.��¼id(+) = n_��¼id And a.��Ŀ��� = ��Ŀ���_In And Rownum < 2;
    End If;
    Select ID
    Into n_��ϸid
    From ���˻�����ϸ
    Where ��¼id = n_��¼id And ��Ŀ��� = ��Ŀ���_In And Nvl(���²�λ, 'TWBW') = Nvl(���²�λ_In, 'TWBW') And
          Nvl(��¼���, 0) = Nvl(��¼���_In, 0) And ��ʼ�汾 = n_��߰汾 And ��ֹ�汾 Is Null;
    --��д��ʷ���ݼ�ǩ����¼����ֹ�汾
    Update ���˻�����ϸ
    Set ��ֹ�汾 = n_��߰汾
    Where ��¼id = n_��¼id And ((Mod(��¼����, 10) <> 5 And ��Ŀ��� = ��Ŀ���_In And Nvl(���²�λ, 'TWBW') = Nvl(���²�λ_In, 'TWBW') And
          Nvl(��¼���, 0) = Nvl(��¼���_In, 0)) Or ��¼���� = Decode(��ǩ_In, 1, 15, 5)) And ��ʼ�汾 <= n_��߰汾 - 1 And ��ֹ�汾 Is Null;

    --�����δǩ�����ݣ�����޸Ĳ���Ա��Ϊ�ü�¼�ı����˸���
    If n_��߰汾 = 1 Then
      Update ���˻������� Set ������ = v_������, ����ʱ�� = Sysdate Where ID = n_��¼id;
    End If;

    If Nvl(n_�������, 0) <> 0 Then
      Return;
    End If;

    --############
    --ͬ����������
    --############
    --1\�ȴ������µ���һ������ʼ��ֻ����һ����Ч�����µ��ļ���
    --������±������ͬ����ʱ������ݣ�ʹ������ID
    --CL,2015-12-30,��¼��ͬ��������Ŀ�����µ�
    For Row_Format In Cur_Fileformats Loop
      If Row_Format.���� = -1 Then
        If Row_Format.���� = '1' Then
          Begin
            Select 1, h.��Ŀ����
            Into Intins, n_��Ŀ����
            From (Select To_Char(f.��Ŀ���) As ���, g.��Ŀ����
                   From ���¼�¼��Ŀ F, �����¼��Ŀ G
                   Where f.��Ŀ��� = g.��Ŀ��� And g.��Ŀ���� = 2 And
                         (g.���ÿ��� = 1 Or
                         (g.���ÿ��� = 2 And Exists
                          (Select 1 From �������ÿ��� D Where g.��Ŀ��� = d.��Ŀ��� And d.����id = v_����id))) And Nvl(g.Ӧ�÷�ʽ, 0) <> 0 And
                         (Nvl(g.���ò���, 0) = 0 Or Nvl(g.���ò���, 0) = Decode(Nvl(Row_Format.Ӥ��, 0), 0, 1, 2))
                   Union All
                   Select b.�����ı� As ���, 1 As ��Ŀ����
                   From �����ļ��ṹ A, �����ļ��ṹ B
                   Where a.�ļ�id = Row_Format.��ʽid And a.��id Is Null And a.������� In (2, 3) And b.��id = a.Id) H
            Where Instr(',' || h.��� || ',', ',' || ��Ŀ���_In || ',', 1) > 0;
          Exception
            When Others Then
              Intins := 0;
          End;
        Else
          Begin
            Select 1, g.��Ŀ����
            Into Intins, n_��Ŀ����
            From ���¼�¼��Ŀ F, �����¼��Ŀ G
            Where f.��Ŀ��� = g.��Ŀ��� And Nvl(g.Ӧ�÷�ʽ, 0) <> 0 And g.����ȼ� >= 0 And
                  (Nvl(g.���ò���, 0) = 0 Or Nvl(g.���ò���, 0) = Decode(Nvl(Row_Format.Ӥ��, 0), 0, 1, 2)) And f.��Ŀ��� = ��Ŀ���_In And
                  (g.���ÿ��� = 1 Or (g.���ÿ��� = 2 And Exists
                   (Select 1 From �������ÿ��� D Where g.��Ŀ��� = d.��Ŀ��� And d.����id = v_����id)));
          Exception
            When Others Then
              Intins := 0;
          End;
        End If;

        If Intins > 0 Then
          --LPF,2013-01-23,������Ŀ�Ƿ���Ҫ����ͬ��(������ǰ�Ѿ�ͬ���������ݣ�Ϊ�˱�֤��¼�������µ�����һֱ�������ݴ˺����жϡ�)
          n_Synchro := Zl_Temperatureprogram(�ļ�id_In, v_����id, ��Ŀ���_In, ����ʱ��_In);
          Begin
            Select b.Id
            Into n_Newid
            From ���˻����ļ� A, ���˻������� B
            Where a.Id = Row_Format.�ļ�id And b.�ļ�id = a.Id And b.����ʱ�� = ����ʱ��_In;
          Exception
            When Others Then
              n_Newid := 0;
          End;
          n_Oldid := n_Newid;
          If n_Newid = 0 And n_Synchro = 1 Then
            Select ���˻�������_Id.Nextval Into n_Newid From Dual;
            --�������µ�����¼
            Insert Into ���˻�������
              (ID, �ļ�id, ������, ����ʱ��, ����ʱ��, ���汾)
            Values
              (n_Newid, Row_Format.�ļ�id, v_������, Sysdate, ����ʱ��_In, 1);
          End If;

          Begin
            Select To_Number(��¼����_In) Into n_Num From Dual;
          Exception
            When Invalid_Number Then
              Begin
                Select 1 Into n_���� From ���¼�¼��Ŀ Where ��Ŀ��� = ��Ŀ���_In And ��¼�� = 1;
              Exception
                When Others Then
                  n_���� := 0;
              End;
              Begin
                Select 1 Into n_δ��˵�� From ��������˵�� Where ���� = ��¼����_In;
              Exception
                When Others Then
                  n_δ��˵�� := 0;
              End;
          End;

          If n_Newid > 0 Then
            --����δͬ�������µ�����(��ȻҪ���Ӷ���ѯ)
            Select Count(*)
            Into v_������Դ
            From ���˻�����ϸ
            Where ��¼id = n_Newid And ��Ŀ��� = ��Ŀ���_In And
                  Decode(n_��Ŀ����, 2, Nvl(���²�λ, '��'), Nvl(���²�λ_In, '��')) = Nvl(���²�λ_In, '��');
            If v_������Դ = 0 Then
              --˵����ͬ����ʼ�Ѿ����й����
              If n_Synchro = 1 Then
                --û�м�����Ŀ�Ƿ���Ҫͬ��
                If n_���� = 1 And n_δ��˵�� = 1 Then
                  Insert Into ���˻�����ϸ
                    (ID, ��¼id, ��¼����, ��Ŀ����, ��Ŀid, ��Ŀ���, ��Ŀ����, ��Ŀ����, ��¼����, ��Ŀ��λ, ��¼���, ���²�λ, ������Դ, ��Դid, δ��˵��, ��ʼ�汾, ��ֹ�汾,
                     ��¼��, ��¼ʱ��, ��¼���)
                    Select ���˻�����ϸ_Id.Nextval, n_Newid, b.��¼����, b.��Ŀ����, b.��Ŀid, b.��Ŀ���, b.��Ŀ����, b.��Ŀ����, Null, b.��Ŀ��λ,
                           b.��¼���, b.���²�λ, 1, b.Id, b.��¼����, 1, Null, b.��¼��, Sysdate, 1
                    From (Select ��Ŀ���_In As ��Ŀ���, Nvl(���²�λ_In, '��') As ���²�λ
                           From Dual
                           Minus
                           Select f.��Ŀ���, Decode(Nvl(f.��Ŀ����, 1), 2, Nvl(���²�λ, '��'), Nvl(���²�λ_In, '��'))
                           From ���˻�����ϸ E, �����¼��Ŀ F
                           Where e.��¼id = n_Newid And e.��Ŀ��� = f.��Ŀ���) A, ���˻�����ϸ B
                    Where a.��Ŀ��� = b.��Ŀ��� And b.��¼id = n_��¼id And b.Id = n_��ϸid;
                  If Sql%RowCount > 0 Then
                    Int���� := 1;
                  End If;
                Else
                  Insert Into ���˻�����ϸ
                    (ID, ��¼id, ��¼����, ��Ŀ����, ��Ŀid, ��Ŀ���, ��Ŀ����, ��Ŀ����, ��¼����, ��Ŀ��λ, ��¼���, ���²�λ, ������Դ, ��Դid, ��ʼ�汾, ��ֹ�汾, ��¼��,
                     ��¼ʱ��, ��¼���)
                    Select ���˻�����ϸ_Id.Nextval, n_Newid, b.��¼����, b.��Ŀ����, b.��Ŀid, b.��Ŀ���, b.��Ŀ����, b.��Ŀ����, b.��¼����, b.��Ŀ��λ,
                           b.��¼���, b.���²�λ, 1, b.Id, 1, Null, b.��¼��, Sysdate, 1
                    From (Select ��Ŀ���_In As ��Ŀ���, Nvl(���²�λ_In, '��') As ���²�λ
                           From Dual
                           Minus
                           Select f.��Ŀ���, Decode(Nvl(f.��Ŀ����, 1), 2, Nvl(���²�λ, '��'), Nvl(���²�λ_In, '��'))
                           From ���˻�����ϸ E, �����¼��Ŀ F
                           Where e.��¼id = n_Newid And e.��Ŀ��� = f.��Ŀ���) A, ���˻�����ϸ B
                    Where a.��Ŀ��� = b.��Ŀ��� And b.��¼id = n_��¼id And b.Id = n_��ϸid;
                  If Sql%RowCount > 0 Then
                    Int���� := 1;
                  End If;
                end if;
              End If;
            Else
              If n_���� = 1 And n_δ��˵�� = 1 Then
                Update ���˻�����ϸ
                Set δ��˵�� = ��¼����_In, ��Դid = n_��ϸid, ��¼���� = Null
                Where ��¼id = n_Newid And ��Ŀ��� = ��Ŀ���_In And
                      Decode(n_��Ŀ����, 2, Nvl(���²�λ, '��'), Nvl(���²�λ_In, '��')) = Nvl(���²�λ_In, '��') And ������Դ > 0;
                If Sql%RowCount > 0 Then
                  Int���� := 1;
                End If;
              Else
                Update ���˻�����ϸ
                Set ��¼���� = ��¼����_In, ��Դid = n_��ϸid
                Where ��¼id = n_Newid And ��Ŀ��� = ��Ŀ���_In And
                      Decode(n_��Ŀ����, 2, Nvl(���²�λ, '��'), Nvl(���²�λ_In, '��')) = Nvl(���²�λ_In, '��') And ������Դ > 0;
                If Sql%RowCount > 0 Then
                  Int���� := 1;
                End If;
              End If;
            End If;
          End If;
        End If;
        --2\��ѭ�������¼��
      Else
        If n_Mutilbill = 1 And n_Syntend = 1 Then
          --��ȡ��¼���뵱ǰ��¼�������ص����������ݵĹ̶���Ŀ
          Select Count(*)
          Into Intins
          From (Select b.��Ŀ���
                 From �����ļ��ṹ A, �����¼��Ŀ B
                 Where a.Ҫ������ = b.��Ŀ���� And b.��Ŀ��ʾ In (0, 4, 5) And
                       ��id =
                       (Select ID From �����ļ��ṹ Where �ļ�id = Row_Format.��ʽid And ��id Is Null And ������� = 4)
                 Intersect
                 Select b.��Ŀ���
                 From �����ļ��ṹ A, �����¼��Ŀ B, ���˻����ļ� C, ���˻������� D, ���˻�����ϸ G
                 Where c.Id = d.�ļ�id And a.�ļ�id = c.��ʽid And d.Id = g.��¼id And d.Id = n_��¼id And g.Id = n_��ϸid And
                       b.��Ŀ��� = g.��Ŀ��� And b.��Ŀ��ʾ In (0, 4, 5) And g.��¼���� = 1 And a.Ҫ������ = b.��Ŀ���� And
                       a.��id = (Select ID From �����ļ��ṹ E Where e.�ļ�id = c.��ʽid And ��id Is Null And ������� = 4));

          If Intins > 0 Then
            n_Newid := 0;
            --����ָ���ļ��Ѿ�������ͬ����ʱ������ݣ�ֱ��������ID����
            Begin
              Select c.Id
              Into n_Newid
              From ���˻������� C
              Where c.�ļ�id = Row_Format.�ļ�id And c.����ʱ�� = ����ʱ��_In;
            Exception
              When Others Then
                n_Newid := 0;
            End;

            If n_Newid = 0 Then
              --������¼������¼
              Select ���˻�������_Id.Nextval Into n_Newid From Dual;

              Insert Into ���˻�������
                (ID, �ļ�id, ������, ����ʱ��, ����ʱ��, ���汾)
                Select n_Newid, Row_Format.�ļ�id, c.������, c.����ʱ��, c.����ʱ��, 1
                From ���˻������� C
                Where c.Id = n_��¼id;
            End If;

            If n_Newid > 0 Then
              --����δͬ���ļ�¼������
              Select Count(*) Into v_������Դ From ���˻�����ϸ Where ��¼id = n_Newid And ��Ŀ��� = ��Ŀ���_In;
              If v_������Դ = 0 Then
                Insert Into ���˻�����ϸ
                  (ID, ��¼id, ��¼����, ��Ŀ����, ��Ŀid, ��Ŀ���, ��Ŀ����, ��Ŀ����, ��¼����, ��Ŀ��λ, ��¼���, ���²�λ, ������Դ, ��Դid, δ��˵��, ��ʼ�汾, ��ֹ�汾,
                   ��¼��, ��¼ʱ��)
                  Select ���˻�����ϸ_Id.Nextval, n_Newid, b.��¼����, b.��Ŀ����, b.��Ŀid, b.��Ŀ���, b.��Ŀ����, b.��Ŀ����, b.��¼����, b.��Ŀ��λ,
                         b.��¼���, b.���²�λ, 1, b.Id, b.δ��˵��, 1, Null, b.��¼��, Sysdate
                  From (Select b.��Ŀ���
                         From �����ļ��ṹ A, �����¼��Ŀ B
                         Where a.Ҫ������ = b.��Ŀ���� And b.��Ŀ��ʾ In (0, 4, 5) And
                               ��id = (Select ID
                                      From �����ļ��ṹ
                                      Where �ļ�id = Row_Format.��ʽid And ��id Is Null And ������� = 4)
                         Intersect
                         Select b.��Ŀ���
                         From �����ļ��ṹ A, �����¼��Ŀ B, ���˻����ļ� C, ���˻������� D, ���˻�����ϸ G
                         Where c.Id = d.�ļ�id And a.�ļ�id = c.��ʽid And d.Id = g.��¼id And d.Id = n_��¼id And g.Id = n_��ϸid And
                               b.��Ŀ��� = g.��Ŀ��� And b.��Ŀ��ʾ In (0, 4, 5) And g.��¼���� = 1 And a.Ҫ������ = b.��Ŀ���� And
                               a.��id =
                               (Select ID From �����ļ��ṹ E Where e.�ļ�id = c.��ʽid And ��id Is Null And ������� = 4)) A, ���˻�����ϸ B
                  Where a.��Ŀ��� = b.��Ŀ��� And b.��¼id = n_��¼id And b.Id = n_��ϸid;
                If Sql%RowCount > 0 Then
                  Int���� := 1;
                  --ԭ������Ҫ��
                  Begin
                    Select ���� Into n_���� From ���˻����ӡ Where �ļ�id = Row_Format.�ļ�id And ��¼id = n_Newid;
                  Exception
                    When Others Then
                      n_���� := 1;
                  End;
                  Zl_���˻����ӡ_Update(Row_Format.�ļ�id, ����ʱ��_In, n_����, 0);
                End If;
              Else
                Update ���˻�����ϸ
                Set ��¼���� = ��¼����_In, δ��˵�� = δ��˵��_In, ��Դid = n_��ϸid
                Where ��¼id = n_Newid And ��Ŀ��� = ��Ŀ���_In And ������Դ > 0;
                If Sql%RowCount > 0 Then
                  Int���� := 1;
                End If;
              End If;
            End If;
          End If;
        End If;
      End If;
    End Loop;

    If Int���� = 1 Then
      Update ���˻�����ϸ Set ���� = 1 Where ID = n_��ϸid;
      --����ʷ���ݵĹ��ñ�־����ΪNULL
      Update ���˻�����ϸ Set ���� = Null Where ��¼id = n_��¼id And ��Ŀ��� = ��Ŀ���_In And ID <> n_��ϸid;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���˻�������_Update;
/


--108753:Ƚ����,2017-05-09,�����ڸ���Ʊ�ݷ�������Զ�����Ʊ����ϸ����ʱ��δ�������ܹ淶��cardinality�ؼ��ʵ�����
Create Or Replace Procedure Zl_Custom_Invoice_Autoallot
(
  ��������_In       Number,
  ģ�����_In       Number,
  Ʊ��_In           Ʊ��ʹ����ϸ.Ʊ��%Type,
  ����id_In         Ʊ��ʹ����ϸ.����id%Type,
  ����id_In         ������ü�¼.����id%Type,
  Nos_In            Varchar2,
  ��ʼ��Ʊ��_In     ������ü�¼.ʵ��Ʊ��%Type,
  ʹ����_In         Ʊ��ʹ����ϸ.ʹ����%Type,
  ʹ��ʱ��_In       Ʊ��ʹ����ϸ.ʹ��ʱ��%Type,
  ��Ʊ��_In         In Out Varchar2,
  ��Ʊ����_In       Out Number,
  �����˲���Ʊ��_In Number := 0,
  ��ӡid_In         Ʊ��ʹ����ϸ.��ӡid%Type := Null,
  Print_Nos_In      t_Strlist := Null
) As
  -------------------------------------------------------------------------------------------------------------
  --���ܣ�����Ʊ�ݷ������,�Զ�����Ʊ����ϸ����
  --��Σ�
  --     ��������_In :1-������ӡƱ��;2-����Ʊ��;3-�ش�Ʊ��;4-�˷��ջ�Ʊ�ݲ����·���Ʊ��
  --     ģ�����_IN :0-������ģ�����;1-����ģ�����,ģ�����ʱ����������
  --     Ʊ��_IN     :1-�����շ�;������������Ʊ��
  --     ����ID_IN   :����ID,���Nos�ͷ�Ʊ��_InΪ��ʱ,��ʾ��Ըò��˵�����δ��ӡ��Ʊ�ݽ��д�ӡ
  --     NOs_IN      :���ݺ�,����ö��ŷ���,�����400�ŵ���,��ʽΪ:A00001,A00002.....
  --     �˷�NOs:�˷����漰�ĵ���
  --     ��ʼ��Ʊ��_IN:�ش�Ʊ�ݻ򷢳�Ʊ�ݵ���ʼƱ��;
  --     ��Ʊ��_In   :����Ϊ���,�ö��ŷָ�,����������Ϊ3-�ش�Ʊ��ʱ��4-�˷ѻ���Ʊ����Ч,��ʾ������Ҫ���յ�Ʊ��
  --     ��ӡid_In:�����˲���ʾ��ʱ����������صĴ�ӡID,�����洫��Ĵ�ӡIDΪ׼
  --     �����˲���Ʊ��_In��1-��ʾ�����˲���Ʊ��,���ֽ������
  --     Print_Nos_in:��ǰ�����漰���շѵ��ݺţ���Ҫ�ǿ��Ƴ���varchar2�Ĵ�С���ƣ���Ҫ�ǰ����˲���Ʊʱ����ֳ�������������ͨ�����ϴ���,��Ҫ��Ǹ���ã����δ�ӡ����>3000ʱ��Nos_in����ֵΪ�ա�
  --����:
  --     ��Ʊ��_In   :����Ϊ���,�ö��ŷָ�,����������Ϊ3-�ش�Ʊ��ʱ��4-�˷ѻ���Ʊ����Ч,��ʾ�ش���˷����·�����Ʊ��
  --     ��Ʊ����_IN :���ر����շ�����Ҫ�ķ�Ʊ����
  -------------------------------------------------------------------------------------------------------------
  v_Err_Msg Varchar2(500);
  Err_Item Exception;
  n_�ֵ��ݴ�ӡ Number(3);
  n_ִ�п���   Number(3);
  n_�վݷ�Ŀ   Number(3);
  n_��������   Number(3);
  n_�շ�ϸĿ   Number(3);

  --------------------------------------------------------
  --�����ڲ�Ʊ�ݴ�������ݼ�
  Type Ty_Rec_Bill Is Record(
    Ʊ��     Ʊ�ݴ�ӡ��ϸ.Ʊ��%Type,
    NO       Ʊ�ݴ�ӡ��ϸ.No%Type,
    ���     Ʊ�ݴ�ӡ��ϸ.���%Type,
    ������� Ʊ�ݴ�ӡ��ϸ.����Ʊ�����%Type,
    �޸ı�־ Number(1));
  Type Ty_Tb_Bill Is Table Of Ty_Rec_Bill;
  c_Invoce Ty_Tb_Bill := Ty_Tb_Bill();
  --------------------------------------------------------
  --��Ԫ��1,Ԫ��2,Ԫ��3,Ԫ��4,�ֱ�ͳ�Ƹ����ݵ����
  Type Ty_Rec_No Is Record(
    NO   ������ü�¼.No%Type,
    ��� Varchar2(1000));
  Type Ty_Tb_No Is Table Of Ty_Rec_No;
  c_No Ty_Tb_No := Ty_Tb_No();
  --------------------------------------------------------
  Cursor c_Fact Is
    Select ǰ׺�ı�, ʣ������, ��ʼ����, ��ֹ����, ��ǰ���� From Ʊ�����ü�¼ Where ID = Nvl(����id_In, 0);
  r_Factrow c_Fact%RowType;

  v_Nos        Varchar2(4000);
  v_��Ʊ��     Ʊ�ݴ�ӡ��ϸ.Ʊ��%Type;
  v_��ʼ��Ʊ�� Ʊ�ݴ�ӡ��ϸ.Ʊ��%Type;
  v_��ǰ��Ʊ�� Ʊ�ݴ�ӡ��ϸ.Ʊ��%Type;
  v_����Ʊ�ݺ� Varchar2(4000);
  n_Find       Number(3);

  n_Ԫ��1_Count Number(3);
  n_Ԫ��2_Count Number(3);
  n_Ԫ��3_Count Number(3);
  n_Ԫ��4_Count Number(3);

  v_Ԫ��1    ������ü�¼.No%Type;
  n_Ԫ��2    ������ü�¼.ִ�в���id%Type;
  v_Ԫ��3    ������ü�¼.�վݷ�Ŀ%Type;
  n_Ԫ��4    ������ü�¼.�շ�ϸĿid%Type;
  v_��Ʊ��Ϣ Varchar2(4000);
  n_�����   Number(1);
  n_��ӡid   Ʊ��ʹ����ϸ.��ӡid%Type;
  n_ʹ��id   Ʊ��ʹ����ϸ.Id%Type;
  n_������   Number(18);
  n_������� Number(18);
  r_���ݺ�   t_Strlist := t_Strlist();
  r_������� t_Strlist := t_Strlist();
  l_ʹ��id   t_Numlist := t_Numlist();
  l_������� t_Numlist := t_Numlist();

  v_��ӡ���� Varchar2(4000);
  v_Temp     Varchar2(4000);
  Procedure Invoice_Split_Notgroup
  (
    Print_Nos        t_Strlist,
    ���շ�Ʊ_In      Varchar2,
    ���δ�ӡ��Ʊ_Out In Out Varchar2,
    ���η�Ʊ����_Out In Out Number,
    Invoce_Out       In Out Ty_Tb_Bill
  ) As
    ----------------------------------------------------------------------------
    --���:
    --   �շ��շ�NOs_IN:������Ҫ����ķ�Ʊ���漰�ĵ���,����ö��ŷ���
    --   ���շ�Ʊ_IN-�˷�ʱ��Ч,����ö��ŷ��룬��ʾ������Ҫ���յķ�Ʊ�� 
    --����:
    -- ���δ�ӡ��Ʊ_Out-������Ҫ�ķ�Ʊ��,����ö��ŷ���
    -- ���η�Ʊ����_Out-������Ҫ�ķ�Ʊ��
    -- Invoce_Out:���η��صķ�Ʊ���뵥�ݵĶ�Ӧ��ϵ
    n_Count Number(18);
    n_��ҳ  Number(18);
  
    Cursor Cr_Bill Is
      Select NO As Ԫ��1, ִ�в���id As Ԫ��2, �վݷ�Ŀ As Ԫ��3, NO As Ԫ��4, NO As ����, ���, 0 As ����
      From ������ü�¼
      Where Rownum <= 1;
    c_Bill Cr_Bill%RowType;
    --------------------------------------------------------------------------------------------
    --������ش��������,ȡ��Ӧ�����ݼ�
    Type Ty_������ϸ Is Ref Cursor;
    c_������ϸ Ty_������ϸ; --�α���� 
  
  Begin
    --�����ݷ���Ʊ��
    If ��������_In = 3 Or ��������_In = 4 Then
      --1-������ӡƱ��;2-����Ʊ��;3-�ش�Ʊ��;4-�˷��ջ�Ʊ�ݲ����·���Ʊ��
      Open c_������ϸ For
        With c_���� As
         (Select Decode(n_�ֵ��ݴ�ӡ, 0, '-', a.No) As Ԫ��1, Decode(n_ִ�п���, 0, 0, a.ִ�в���id) As Ԫ��2,
                 Decode(n_�վݷ�Ŀ, 0, '-', a.�վݷ�Ŀ) As Ԫ��3, Decode(n_�շ�ϸĿ, 0, '-', a.No) As Ԫ��4, a.No As ����,
                 Nvl(a.�۸񸸺�, a.���) As ���, Sum(Nvl(a.ʵ�ս��, 0)) As ʵ�ս��
          From ������ü�¼ A,
               (Select /*+cardinality(j,10)*/
                  NO, ���
                 From Ʊ�ݴ�ӡ��ϸ M, Table(f_Str2list(���շ�Ʊ_In)) J
                 Where m.Ʊ�� = j.Column_Value) B
          Where Mod(a.��¼����, 10) = 1 And a.No = b.No And Instr(',' || b.��� || ',', ',' || Nvl(a.�۸񸸺�, a.���) || ',') > 0 And
                Decode(n_�����, 1, Nvl(a.���ӱ�־, 0), 0) <> 9
          Group By a.No, a.ִ�в���id, a.�վݷ�Ŀ, a.�շ�ϸĿid, Nvl(a.�۸񸸺�, a.���)
          Having Sum(Nvl(a.ʵ�ս��, 0)) <> 0)
        Select Ԫ��1, Ԫ��2, Ԫ��3, Ԫ��4, ����, ���, Count(*) As ����
        From c_����
        Group By Ԫ��1, Ԫ��2, Ԫ��3, Ԫ��4, ����, ���
        Order By Ԫ��1, Ԫ��2, Ԫ��3, Ԫ��4, ����, ���;
    Else
      Open c_������ϸ For
        With c_���� As
         (Select /*+cardinality(b,10)*/
           Decode(n_�ֵ��ݴ�ӡ, 0, '-', a.No) As Ԫ��1, Decode(n_ִ�п���, 0, 0, a.ִ�в���id) As Ԫ��2,
           Decode(n_�վݷ�Ŀ, 0, '-', a.�վݷ�Ŀ) As Ԫ��3, Decode(n_�շ�ϸĿ, 0, '-', a.No) As Ԫ��4, a.No As ����,
           Nvl(a.�۸񸸺�, a.���) As ���, Sum(Nvl(a.ʵ�ս��, 0)) As ʵ�ս��
          From ������ü�¼ A, Table(Print_Nos) B
          Where Mod(a.��¼����, 10) = 1 And a.No = b.Column_Value And Decode(n_�����, 1, Nvl(a.���ӱ�־, 0), 0) <> 9
          Group By a.No, a.ִ�в���id, a.�վݷ�Ŀ, a.�շ�ϸĿid, Nvl(a.�۸񸸺�, a.���)
          Having Sum(Nvl(a.ʵ�ս��, 0)) <> 0)
        Select Ԫ��1, Ԫ��2, Ԫ��3, Ԫ��4, ����, ���, Count(*) As ����
        From c_����
        Group By Ԫ��1, Ԫ��2, Ԫ��3, Ԫ��4, ����, ���
        Order By Ԫ��1, Ԫ��2, Ԫ��3, Ԫ��4, ����, ���;
    End If;
  
    v_Ԫ��1          := '+';
    n_Ԫ��2          := 0;
    v_Ԫ��3          := '+';
    n_Ԫ��4          := 0;
    n_Ԫ��1_Count    := 0;
    n_Ԫ��2_Count    := 0;
    n_Ԫ��3_Count    := 0;
    n_Ԫ��4_Count    := 0;
    ���η�Ʊ����_Out := 0;
    If n_�������� <> 0 Then
      n_������� := 1;
    Else
      n_������� := 0;
    End If;
    n_Count := 0;
    c_No.Delete;
    Loop
      Fetch c_������ϸ
        Into c_Bill;
      Exit When c_������ϸ%NotFound;
      n_Count := 1;
    
      n_��ҳ := 0;
      If (v_Ԫ��1 <> c_Bill.Ԫ��1) Or (n_Ԫ��2 <> c_Bill.Ԫ��2 And n_Ԫ��2_Count >= n_ִ�п��� And n_ִ�п��� <> 0) Or
         (v_Ԫ��3 <> c_Bill.Ԫ��3 And n_Ԫ��3_Count >= n_�վݷ�Ŀ And n_�վݷ�Ŀ <> 0) Or (n_Ԫ��4_Count >= n_�շ�ϸĿ And n_�շ�ϸĿ <> 0) Then
      
        If (v_Ԫ��1 <> '+' Or n_Ԫ��2 <> 0 Or v_Ԫ��3 <> '+' Or n_Ԫ��4 <> 0) Then
          n_��ҳ := 1;
        End If;
        n_Ԫ��2_Count := 0;
        n_Ԫ��3_Count := 0;
        n_Ԫ��4_Count := 0;
        n_Ԫ��1_Count := 0;
        v_Ԫ��1       := '+';
        n_Ԫ��2       := 0;
        v_Ԫ��3       := '+';
      End If;
    
      If n_��ҳ = 1 Then
        --��ҳ:���㷢Ʊ�ż���ص�
        For I In 1 .. c_No.Count Loop
          Invoce_Out.Extend;
          Invoce_Out(Invoce_Out.Count).Ʊ�� := v_��Ʊ��;
          Invoce_Out(Invoce_Out.Count).No := c_No(I).No;
          Invoce_Out(Invoce_Out.Count).��� := Case
                                               When Instr(c_No(I).���, ',') > 0 Then
                                                Substr(c_No(I).���, 2)
                                               Else
                                                c_No(I).���
                                             End;
          Invoce_Out(Invoce_Out.Count).������� := n_�������;
        End Loop;
      
        ���η�Ʊ����_Out := ���η�Ʊ����_Out + 1;
        ���δ�ӡ��Ʊ_Out := Nvl(���δ�ӡ��Ʊ_Out, '') || ',' || v_��Ʊ��;
        v_��Ʊ��         := Zl_Incstr(v_��Ʊ��);
        c_No.Delete;
      End If;
      If (v_Ԫ��1 <> c_Bill.Ԫ��1) Then
        n_Ԫ��1_Count := n_Ԫ��1_Count + 1;
        v_Ԫ��1       := c_Bill.Ԫ��1;
      End If;
      If (n_Ԫ��2 <> c_Bill.Ԫ��2) Then
        n_Ԫ��2_Count := n_Ԫ��2_Count + 1;
        n_Ԫ��2       := c_Bill.Ԫ��2;
      End If;
      If (v_Ԫ��3 <> c_Bill.Ԫ��3) Then
        n_Ԫ��3_Count := n_Ԫ��3_Count + 1;
        v_Ԫ��3       := c_Bill.Ԫ��3;
      End If;
      If n_�շ�ϸĿ <> 0 Then
        n_Ԫ��4_Count := n_Ԫ��4_Count + 1;
      End If;
    
      -------------------------------------------
      --���䵥�ݺż����
      n_Find := 0;
      For J In 1 .. c_No.Count Loop
        If c_No(J).No = c_Bill.���� Then
          --���ݺ���ͬ,����źϲ�
          c_No(J).��� := c_No(J).��� || ',' || c_Bill.���;
          n_Find := 1;
          Exit;
        End If;
      End Loop;
      If n_Find = 0 Then
        c_No.Extend;
        c_No(c_No.Count).No := c_Bill.����;
        c_No(c_No.Count).��� := c_No(c_No.Count).��� || ',' || c_Bill.���;
      End If;
    End Loop;
  
    --�Ƿ��з�Ʊ����
    If n_Count >= 1 Then
      --���һ����Ʊ����
      ���η�Ʊ����_Out := ���η�Ʊ����_Out + 1;
      ���δ�ӡ��Ʊ_Out := Nvl(���δ�ӡ��Ʊ_Out, '') || ',' || v_��Ʊ��;
    Else
      ���η�Ʊ����_Out := 0;
      ���δ�ӡ��Ʊ_Out := '';
    End If;
    If c_No.Count <> 0 Then
      For I In 1 .. c_No.Count Loop
        Invoce_Out.Extend;
        Invoce_Out(Invoce_Out.Count).Ʊ�� := v_��Ʊ��;
        Invoce_Out(Invoce_Out.Count).No := c_No(I).No;
        If Instr(c_No(I).���, ',') > 0 Then
          c_No(I).��� := Substr(c_No(I).���, 2);
        End If;
        Invoce_Out(Invoce_Out.Count).��� := c_No(I).���;
        Invoce_Out(Invoce_Out.Count).������� := n_�������;
      End Loop;
      c_No.Delete;
    End If;
    If Instr(Nvl(���δ�ӡ��Ʊ_Out, '-'), ',') > 0 Then
      ���δ�ӡ��Ʊ_Out := Substr(���δ�ӡ��Ʊ_Out, 2);
    End If;
  End Invoice_Split_Notgroup;

Begin
  --����Ʊ������
  If Ʊ��_In <> 1 Then
    --�ݲ�֧������,ֻ֧�������շ�
    Return;
  End If;
  v_��Ʊ�� := ��ʼ��Ʊ��_In;
  v_Nos    := Nos_In;
  -----------------------------------------------------------------------------------------------------------------------------
  --һ����ȡ��Ʊ�������ع���
  --**��ʼ
  --1.ȷ���Ƿ�ֵ��ݷ���Ʊ��,ȱʡ�������ݷֺ�
  n_�ֵ��ݴ�ӡ := 0;
  --2.ȷ���Ƿ�ִ�п��ҷֵ��ݺ�,ȱʡΪ��1��ִ�п��ҷֺ�
  n_ִ�п��� := 1;

  --3.ȷ���Ƿ��վݷ�Ŀ�ֵ��ݺ�,ȱʡΪ��3���վݷ�Ŀ�ֺ�
  n_�վݷ�Ŀ := 3;

  --4.ȷ���Ƿ��շ�ϸĿ�ֵ��ݺ�,ȱʡΪ�����շ�ϸĿ�ֺ�
  n_�շ�ϸĿ := 0;

  --5.�����Ƿ���ҳ����,ȱʡΪ������
  n_�������� := 0;

  v_����Ʊ�ݺ� := ��Ʊ��_In;
  ��Ʊ����_In  := 0;
  --**����
  If Nvl(�����˲���Ʊ��_In, 0) <> 0 Then
    --�����˲���Ʊ��ʱ��ֻ���շѷ�Ŀ��ӡ
    n_ִ�п��� := 0;
  End If;

  -----------------------------------------------------------------------------------------------------------------------------
  --�������з�Ʊ����
  Invoice_Split_Notgroup(Print_Nos_In, ��Ʊ��_In, v_��Ʊ��Ϣ, ��Ʊ����_In, c_Invoce);

  -----------------------------------------------------------------------------------------------------------------------------
  --*****************************************************************************************************************************
  --ע��:
  --���´��룬���������,������Ĵ�������Ҫȷ������������ֵ:һ��v_��Ʊ��Ϣ;����c_Invoce�����е�ֵ
  --  v_��Ʊ��Ϣ:�������漰�ķ�Ʊ��Ϣ,����ö��ŷ���,��ð���������
  --  c_Invoce:Ϊ�������ݣ�Ϊ��Ʊ�ź͵��ݵĶ�Ӧ��ϵ

  ��Ʊ��_In := v_��Ʊ��Ϣ;
  If ģ�����_In = 1 Then
    --ģ�����,ֻ����Ʊ��������ʹ�õ�Ʊ�ݺ�,ֱ���˳�
    Return;
  End If;
  -----------------------------------------------------------------------------------------------------------------------------
  --�ġ��˷�ʱ����Ҫ�ȴ�����շ�Ʊ
  v_��ʼ��Ʊ�� := Null;
  v_��ǰ��Ʊ�� := Null;
  --1-������ӡƱ��;2-����Ʊ��;3-�ش�Ʊ��;4-�˷��ջ�Ʊ�ݲ����·���Ʊ��
  If ��������_In = 3 Or ��������_In = 4 Then
    --�ջ�Ʊ��
    Select ʹ��id Bulk Collect
    Into l_ʹ��id
    From (Select /*+cardinality(j,10)*/
           Distinct b.ʹ��id
           From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ��ϸ B, Table(f_Str2list(v_����Ʊ�ݺ�)) J
           Where a.Id = b.ʹ��id And b.Ʊ�� = j.Column_Value And Nvl(b.Ʊ��, 0) = 1);
  
    --������ռ�¼
    Forall I In 1 .. l_ʹ��id.Count
      Insert Into Ʊ��ʹ����ϸ
        (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ����, ʹ��ʱ��)
        Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, Decode(��������_In, 3, 4, 2), ����id, ��ӡid, ʹ����_In, ʹ��ʱ��_In
        From Ʊ��ʹ����ϸ
        Where ID = l_ʹ��id(I);
    Forall I In 1 .. l_ʹ��id.Count
      Update Ʊ�ݴ�ӡ��ϸ Set �Ƿ���� = 1 Where ʹ��id = l_ʹ��id(I);
  End If;

  If c_Invoce.Count = 0 Then
    --�޷�Ʊ����,��ֱ�ӷ���,�˷�ʱ����ʾֻ�ջ�Ʊ��
    Return;
  End If;

  -----------------------------------------------------------------------------------------------------------------------------
  --�塢���´�������Ʊ��(���˷����·�����Ʊ�ݴ���)
  If ��ʼ��Ʊ��_In Is Null Then
    v_Err_Msg := 'δ������ʼ��Ʊ��,���ܽ���Ʊ�ݷ��䴦��';
    Raise Err_Item;
  End If;

  If Nvl(����id_In, 0) <> 0 Then
    Open c_Fact;
    Fetch c_Fact
      Into r_Factrow;
    If c_Fact%RowCount = 0 Then
      v_Err_Msg := '��Ч��Ʊ���������Σ��޷�����շ�Ʊ�ݷ��������';
      Close c_Fact;
      Raise Err_Item;
    Elsif Nvl(r_Factrow.ʣ������, 0) < ��Ʊ����_In Then
      v_Err_Msg := '��ǰ���ε�ʣ����������' || ��Ʊ����_In || '�ţ��޷�����շ�Ʊ�ݷ��������';
      Close c_Fact;
      Raise Err_Item;
    End If;
  End If;

  --1.ʵ�ʴ���Ʊ����Ϣ
  If Nvl(n_�ֵ��ݴ�ӡ, 0) <> 1 Or Nvl(�����˲���Ʊ��_In, 0) = 1 Then
    --���ֵ��ݴ�ӡʱ,��ʾһ�δ�ӡ,��ӡID���һ��
    n_��ӡid := ��ӡid_In;
    If Nvl(n_��ӡid, 0) = 0 Then
      Select Ʊ�ݴ�ӡ����_Id.Nextval Into n_��ӡid From Dual;
    End If;
  End If;

  ��Ʊ����_In := 0;
  v_��ӡ����  := '';
  For c_Invoce_No In (Select Column_Value As ��Ʊ�� From Table(f_Str2list(v_��Ʊ��Ϣ)) Order By ��Ʊ��) Loop
    --2.���Ʊ�ݷ�Χ�Ƿ���ȷ
    If Nvl(����id_In, 0) <> 0 Then
      If Not (Upper(c_Invoce_No.��Ʊ��) >= Upper(r_Factrow.��ʼ����) And Upper(c_Invoce_No.��Ʊ��) <= Upper(r_Factrow.��ֹ����) And
          Length(c_Invoce_No.��Ʊ��) = Length(r_Factrow.��ֹ����)) Then
        v_Err_Msg := '�õ�����Ҫ��ӡ����Ʊ��,��Ʊ�ݺ�"' || c_Invoce_No.��Ʊ�� || '"����Ʊ�����õĺ��뷶Χ��';
        Close c_Fact;
        Raise Err_Item;
      End If;
    End If;
  
    --3.����Ʊ�ݴ�ӡ��ϸ
    r_���ݺ�.Delete;
    r_�������.Delete;
    l_�������.Delete;
  
    Select Ʊ��ʹ����ϸ_Id.Nextval Into n_ʹ��id From Dual;
  
    n_������� := 0;
    For I In 1 .. c_Invoce.Count Loop
      If c_Invoce(I).Ʊ�� = c_Invoce_No.��Ʊ�� Then
        n_������� := c_Invoce(I).�������;
        Exit;
      End If;
    End Loop;
  
    --�������Ʊ��,�Ա����Ʊ��
    For I In 1 .. c_Invoce.Count Loop
      If c_Invoce(I).������� = n_������� And Nvl(c_Invoce(I).�޸ı�־, 0) = 0 Then
        If n_������� <> 0 Then
          c_Invoce(I).������� := n_ʹ��id;
        End If;
        c_Invoce(I).�޸ı�־ := 1;
      End If;
    End Loop;
  
    For I In 1 .. c_Invoce.Count Loop
      If c_Invoce(I).Ʊ�� = c_Invoce_No.��Ʊ�� Then
        r_���ݺ�.Extend;
        r_���ݺ�(r_���ݺ�.Count) := c_Invoce(I).No;
        r_�������.Extend;
        r_�������(r_�������.Count) := c_Invoce(I).���;
        l_�������.Extend;
        If Nvl(c_Invoce(I).�������, 0) <> 0 Then
          --����Ƿ����������Ʊ��
          n_Find := 0;
          For J In 1 .. c_Invoce.Count Loop
            If c_Invoce(I).������� = c_Invoce(J).������� And c_Invoce(I).Ʊ�� <> c_Invoce(J).Ʊ�� Then
              n_Find := 1;
              Exit;
            End If;
          End Loop;
        
          If n_Find = 0 Then
            l_�������(l_�������.Count) := Null;
            c_Invoce(I).������� := 0;
          Else
            l_�������(l_�������.Count) := c_Invoce(I).�������;
          End If;
        Else
          l_�������(l_�������.Count) := Null;
        End If;
      End If;
    End Loop;
  
    --1.�����Ŵ�ӡ����
    If n_�ֵ��ݴ�ӡ = 1 Then
      --�ֵ��ݴ�ӡ,�谴���ݽ��д���
      --Ʊ�ݴ�ӡ����
      n_Find := 0;
      v_Temp := '';
      For I In 1 .. r_���ݺ�.Count Loop
        v_Temp := v_Temp || ',' || r_���ݺ�(I);
        If Instr(Nvl(v_��ӡ����, '-') || ',', ',' || r_���ݺ�(I) || ',') > 0 Then
          --�Ѿ��ҵ�
          n_Find := 1;
        End If;
      End Loop;
      v_��ӡ���� := v_��ӡ���� || Nvl(v_Temp, '+');
    
      If Nvl(n_Find, 0) = 0 Then
        Select Ʊ�ݴ�ӡ����_Id.Nextval Into n_��ӡid From Dual;
        Forall I In 1 .. r_���ݺ�.Count
          Insert Into Ʊ�ݴ�ӡ����
            (ID, ��������, NO, ��ӡ����)
          Values
            (n_��ӡid, 1, r_���ݺ�(I), Decode(Nvl(�����˲���Ʊ��_In, 0), 1, 1, 0));
        --�Ա����������ü�¼�е�ʵ��Ʊ��
        v_��ʼ��Ʊ�� := c_Invoce_No.��Ʊ��;
        Forall I In 1 .. r_���ݺ�.Count
          Update ������ü�¼ Set ʵ��Ʊ�� = v_��ʼ��Ʊ�� Where Mod(��¼����, 10) = 1 And NO = r_���ݺ�(I);
      End If;
    Else
    
      If v_��ʼ��Ʊ�� Is Null Then
        --�Ա����������ü�¼�е�ʵ��Ʊ��
        v_��ʼ��Ʊ�� := c_Invoce_No.��Ʊ��;
      
        --Ʊ�ݴ�ӡ����
        Insert Into Ʊ�ݴ�ӡ����
          (ID, ��������, NO, ��ӡ����)
          Select n_��ӡid, 1, Column_Value, Decode(Nvl(�����˲���Ʊ��_In, 0), 1, 1, 0) From Table(Print_Nos_In);
      
        Update ������ü�¼
        Set ʵ��Ʊ�� = v_��ʼ��Ʊ��
        Where Mod(��¼����, 10) = 1 And NO In (Select Column_Value From Table(Print_Nos_In));
      End If;
    End If;
    --2.����Ʊ�ݴ�ӡ��ϸ
  
    ��Ʊ����_In := ��Ʊ����_In + 1;
    --����Ʊ��ʹ����ϸ
  
    Insert Into Ʊ��ʹ����ϸ
      (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ����, ʹ��ʱ��)
    Values
      (n_ʹ��id, 1, c_Invoce_No.��Ʊ��, 1, Decode(��������_In, 3, 3, 1), Decode(Nvl(����id_In, 0), 0, Null, ����id_In), n_��ӡid,
       ʹ����_In, ʹ��ʱ��_In);
  
    Forall I In 1 .. r_���ݺ�.Count
      Insert Into Ʊ�ݴ�ӡ��ϸ
        (ʹ��id, Ʊ��, �Ƿ����, NO, Ʊ��, ���, ����Ʊ�����)
      Values
        (n_ʹ��id, 1, 0, r_���ݺ�(I), c_Invoce_No.��Ʊ��, r_�������(I), l_�������(I));
  
    v_��ǰ��Ʊ�� := c_Invoce_No.��Ʊ��;
  End Loop;

  If Nvl(����id_In, 0) <> 0 Then
    Close c_Fact;
    Update Ʊ�����ü�¼
    Set ʹ��ʱ�� = ʹ��ʱ��_In, ��ǰ���� = v_��ǰ��Ʊ��, ʣ������ = Nvl(ʣ������, 0) - ��Ʊ����_In
    Where ID = ����id_In
    Returning ʣ������ Into n_������;
    If n_������ < 0 Then
      v_Err_Msg := '��ǰ���ε�ʣ����������' || ��Ʊ����_In || '�ţ��޷�����շ�Ʊ�ݷ��������';
      Raise Err_Item;
    End If;
  End If;
  --*****************************************************************************************************************************
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Custom_Invoice_Autoallot;
/

--108753:Ƚ����,2017-05-09,�����ڸ���Ʊ�ݷ�������Զ�����Ʊ����ϸ����ʱ��δ�������ܹ淶��cardinality�ؼ��ʵ�����
Create Or Replace Procedure Zl_Invoice_Autoallot
(
  ��������_In   Number,
  ģ�����_In   Number,
  Ʊ��_In       Ʊ��ʹ����ϸ.Ʊ��%Type,
  ����id_In     Ʊ��ʹ����ϸ.����id%Type,
  ����id_In     ������ü�¼.����id%Type,
  Nos_In        Varchar2,
  ��ʼ��Ʊ��_In ������ü�¼.ʵ��Ʊ��%Type,
  ʹ����_In     Ʊ��ʹ����ϸ.ʹ����%Type,
  ʹ��ʱ��_In   Ʊ��ʹ����ϸ.ʹ��ʱ��%Type,
  ��Ʊ��_In     In Out Varchar2,
  ��Ʊ����_In   Out Number,
  ��ӡid_In     Ʊ��ʹ����ϸ.��ӡid%Type := 0
) As
  ---------------------------------------------------------------------------------------------
  --���ܣ�����Ʊ�ݷ������,�Զ�����Ʊ����ϸ����
  --��Σ�
  --     ��������_In :1-������ӡƱ��;2-����Ʊ��;3-�ش�Ʊ��;4-�˷��ջ�Ʊ�ݲ����·���Ʊ��
  --     Ʊ��_IN     :1-�����շ�;������������Ʊ��
  --     ����ID_IN   :����ID,���Nos�ͷ�Ʊ��_InΪ��ʱ,��ʾ��Ըò��˵�����δ��ӡ��Ʊ�ݽ��д�ӡ
  --     NOs_IN      :���ݺ�,����ö��ŷ���,���;��400�ŵ���,��ʽΪ:A00001,A00002.....
  --     ��ʼ��Ʊ��_IN:�ش�Ʊ�ݻ򷢳�Ʊ�ݵ���ʼƱ��;
  --     ��Ʊ��_In   :����Ϊ���,����������Ϊ3-�ش�Ʊ��ʱ,��Ч
  --     ģ�����_IN :0-������ģ�����;1-����ģ�����,ģ�����ʱ����������
  --     ��ӡID_In :��ӡID_In<>0ʱ����ʾ������ʱ��"��ʱƱ�ݴ�ӡ����"����Ӧ��NO��������ӡ����(��Ҫ��������˲���Ʊ���ֽ�����������)
  --����:
  --     ��Ʊ����_IN :���ر����շ�����Ҫ�ķ�Ʊ����
  ---------------------------------------------------------------------------------------------
  v_Err_Msg Varchar2(500);
  Err_Item Exception;
  v_Para       Varchar2(1000);
  v_Temp       Varchar2(32767);
  n_����ģʽ   Number(3);
  n_�ֵ��ݴ�ӡ Number(3);
  n_ִ�п���   Number(3);
  n_�վݷ�Ŀ   Number(3);
  n_��������   Number(3);
  n_�շ�ϸĿ   Number(3);

  ---------------------------------------------------------
  Type Ty_Rec_Splitno Is Record(
    Ԫ��1    Ʊ�ݴ�ӡ��ϸ.No%Type,
    Ԫ��2��  Varchar2(4000),
    Ԫ��3��  Varchar2(4000),
    ������� Number(18));

  Type Ty_Tb_Splitno Is Table Of Ty_Rec_Splitno;
  c_Split_No   Ty_Tb_Splitno := Ty_Tb_Splitno();
  c_Split_��Ŀ Ty_Tb_Splitno := Ty_Tb_Splitno();

  --------------------------------------------------------
  --�����ڲ�Ʊ�ݴ�������ݼ�
  Type Ty_Rec_Bill Is Record(
    Ʊ��     Ʊ�ݴ�ӡ��ϸ.Ʊ��%Type,
    NO       Ʊ�ݴ�ӡ��ϸ.No%Type,
    ���     Ʊ�ݴ�ӡ��ϸ.���%Type,
    ������� Ʊ�ݴ�ӡ��ϸ.����Ʊ�����%Type,
    �޸ı�־ Number(1));
  Type Ty_Tb_Bill Is Table Of Ty_Rec_Bill;
  c_Invoce Ty_Tb_Bill := Ty_Tb_Bill();
  --------------------------------------------------------
  --��Ԫ��1,Ԫ��2,Ԫ��3,Ԫ��4,�ֱ�ͳ�Ƹ����ݵ����
  Type Ty_Rec_No Is Record(
    NO   ������ü�¼.No%Type,
    ��� Varchar2(1000));
  Type Ty_Tb_No Is Table Of Ty_Rec_No;
  c_No Ty_Tb_No := Ty_Tb_No();
  --------------------------------------------------------
  Cursor c_Fact Is
    Select ǰ׺�ı�, ʣ������, ��ʼ����, ��ֹ����, ��ǰ���� From Ʊ�����ü�¼ Where ID = Nvl(����id_In, 0);
  r_Factrow c_Fact%RowType;

  --------------------------------------------------------------------------------------------
  --������ش��������,ȡ��Ӧ�����ݼ�

  v_Nos        Varchar2(32767);
  v_��Ʊ��     Ʊ�ݴ�ӡ��ϸ.Ʊ��%Type;
  v_��ʼ��Ʊ�� Ʊ�ݴ�ӡ��ϸ.Ʊ��%Type;
  v_��ǰ��Ʊ�� Ʊ�ݴ�ӡ��ϸ.Ʊ��%Type;
  v_����Ʊ�ݺ� Varchar2(4000);
  n_Find       Number(3);

  n_Ԫ��1_Count Number(3);
  n_Ԫ��2_Count Number(3);
  n_Ԫ��3_Count Number(3);
  n_Ԫ��4_Count Number(3);

  v_Ԫ��1     ������ü�¼.No%Type;
  n_Ԫ��2     ������ü�¼.ִ�в���id%Type;
  v_Ԫ��3     ������ü�¼.�վݷ�Ŀ%Type;
  n_Ԫ��4     ������ü�¼.�շ�ϸĿid%Type;
  v_��Ʊ��Ϣ  Varchar2(4000);
  n_�����    Number(1);
  n_��ӡid    Ʊ��ʹ����ϸ.��ӡid%Type;
  n_ʹ��id    Ʊ��ʹ����ϸ.Id%Type;
  n_������    Number(18);
  n_�������  Number(18);
  r_���ݺ�    t_Strlist := t_Strlist();
  l_Print_Nos t_Strlist := t_Strlist();
  r_�������  t_Strlist := t_Strlist();
  l_ʹ��id    t_Numlist := t_Numlist();
  l_�������  t_Numlist := t_Numlist();

  v_��ӡ����       Varchar2(4000);
  l_Ԫ��2          t_Numlist := t_Numlist();
  l_Ԫ��3          t_Strlist := t_Strlist();
  v_��ʼ��Ʊ��     Ʊ�����ü�¼.��ʼ����%Type;
  n_�����˲���Ʊ�� Number(2);
  n_��ӡ����       Ʊ�ݴ�ӡ����.��ӡ����%Type;

  -------------------------------------------------------------------------------------------------------------------
  --Invoice_Split_Notgroup:�����з�����ܻ���ҳ����ʱ���ô˹���
  Procedure Invoice_Split_Notgroup
  (
    Print_Nos        t_Strlist,
    ���շ�Ʊ_In      Varchar2,
    ���δ�ӡ��Ʊ_Out Out Varchar2,
    ���η�Ʊ����_Out Out Number
  ) As
    ----------------------------------------------------------------------------
    --���:
    --   �շ��շ�NOs_IN:������Ҫ����ķ�Ʊ���漰�ĵ���,����ö��ŷ���
    --   ���շ�Ʊ_IN-�˷�ʱ��Ч,����ö��ŷ��룬��ʾ������Ҫ���յķ�Ʊ��
    --����:
    -- ���δ�ӡ��Ʊ_Out-������Ҫ�ķ�Ʊ��,����ö��ŷ���
    -- ���η�Ʊ����_Out-������Ҫ�ķ�Ʊ��
    -- �����˷ѵ���_Out-�˷ѻ������漰��NO��,����ö��ŷ���
  
    n_Count Number(18);
    n_��ҳ  Number(18);
  
    Cursor Cr_Bill Is
      Select NO As Ԫ��1, ִ�в���id As Ԫ��2, �վݷ�Ŀ As Ԫ��3, NO As Ԫ��4, NO As ����, ���, 0 As ����
      From ������ü�¼
      Where Rownum <= 1;
    c_Bill Cr_Bill%RowType;
    --------------------------------------------------------------------------------------------
    --������ش��������,ȡ��Ӧ�����ݼ�
    Type Ty_������ϸ Is Ref Cursor;
    c_������ϸ Ty_������ϸ; --�α����
  
  Begin
    --�����ݷ���Ʊ��
    If ��������_In = 3 Or ��������_In = 4 Then
      Open c_������ϸ For
        With c_���� As
         (Select Decode(n_�ֵ��ݴ�ӡ, 0, '-', a.No) As Ԫ��1, Decode(n_ִ�п���, 0, 0, a.ִ�в���id) As Ԫ��2,
                 Decode(n_�վݷ�Ŀ, 0, '-', a.�վݷ�Ŀ) As Ԫ��3, Decode(n_�շ�ϸĿ, 0, '-', a.No) As Ԫ��4, a.No As ����,
                 Nvl(a.�۸񸸺�, a.���) As ���, Sum(Nvl(a.ʵ�ս��, 0)) As ʵ�ս��
          From ������ü�¼ A,
               (Select /*+cardinality(j,10)*/
                  NO, ���
                 From Ʊ�ݴ�ӡ��ϸ M, Table(f_Str2list(���շ�Ʊ_In)) J
                 Where m.Ʊ�� = j.Column_Value) B
          Where Mod(��¼����, 10) = 1 And a.No = b.No And Instr(',' || b.��� || ',', ',' || Nvl(a.�۸񸸺�, a.���) || ',') > 0 And
                Decode(n_�����, 1, Nvl(a.���ӱ�־, 0), 0) <> 9
          Group By a.No, a.ִ�в���id, a.�վݷ�Ŀ, a.�շ�ϸĿid, Nvl(a.�۸񸸺�, a.���)
          Having Sum(Nvl(a.ʵ�ս��, 0)) <> 0)
        Select Ԫ��1, Ԫ��2, Ԫ��3, Ԫ��4, ����, ���, Count(*) As ����
        From c_����
        Group By Ԫ��1, Ԫ��2, Ԫ��3, Ԫ��4, ����, ���
        Order By Ԫ��1, Ԫ��2, Ԫ��3, Ԫ��4, ����, ���;
    Else
      Open c_������ϸ For
        With c_���� As
         (Select /*+cardinality(b,10)*/
           Decode(n_�ֵ��ݴ�ӡ, 0, '-', a.No) As Ԫ��1, Decode(n_ִ�п���, 0, 0, a.ִ�в���id) As Ԫ��2,
           Decode(n_�վݷ�Ŀ, 0, '-', a.�վݷ�Ŀ) As Ԫ��3, Decode(n_�շ�ϸĿ, 0, '-', a.No) As Ԫ��4, a.No As ����,
           Nvl(a.�۸񸸺�, a.���) As ���, Sum(Nvl(a.ʵ�ս��, 0)) As ʵ�ս��
          From ������ü�¼ A, Table(Print_Nos) B
          Where Mod(��¼����, 10) = 1 And a.No = b.Column_Value And Decode(n_�����, 1, Nvl(a.���ӱ�־, 0), 0) <> 9
          Group By a.No, a.ִ�в���id, a.�վݷ�Ŀ, a.�շ�ϸĿid, Nvl(a.�۸񸸺�, a.���)
          Having Sum(Nvl(a.ʵ�ս��, 0)) <> 0)
        Select Ԫ��1, Ԫ��2, Ԫ��3, Ԫ��4, ����, ���, Count(*) As ����
        From c_����
        Group By Ԫ��1, Ԫ��2, Ԫ��3, Ԫ��4, ����, ���
        Order By Ԫ��1, Ԫ��2, Ԫ��3, Ԫ��4, ����, ���;
    End If;
  
    v_Ԫ��1          := '+';
    n_Ԫ��2          := 0;
    v_Ԫ��3          := '+';
    n_Ԫ��4          := 0;
    n_Ԫ��1_Count    := 0;
    n_Ԫ��2_Count    := 0;
    n_Ԫ��3_Count    := 0;
    n_Ԫ��4_Count    := 0;
    ���η�Ʊ����_Out := 0;
    If n_�������� <> 0 Then
      n_������� := 1;
    Else
      n_������� := 0;
    End If;
    n_Count := 0;
    c_No.Delete;
    Loop
      Fetch c_������ϸ
        Into c_Bill;
      Exit When c_������ϸ%NotFound;
      n_Count := 1;
    
      n_��ҳ := 0;
      If (v_Ԫ��1 <> c_Bill.Ԫ��1) Or (n_Ԫ��2 <> c_Bill.Ԫ��2 And n_Ԫ��2_Count >= n_ִ�п��� And n_ִ�п��� <> 0) Or
         (v_Ԫ��3 <> c_Bill.Ԫ��3 And n_Ԫ��3_Count >= n_�վݷ�Ŀ And n_�վݷ�Ŀ <> 0) Or (n_Ԫ��4_Count >= n_�շ�ϸĿ And n_�շ�ϸĿ <> 0) Then
      
        If (v_Ԫ��1 <> '+' Or n_Ԫ��2 <> 0 Or v_Ԫ��3 <> '+' Or n_Ԫ��4 <> 0) Then
          n_��ҳ := 1;
        End If;
        n_Ԫ��2_Count := 0;
        n_Ԫ��3_Count := 0;
        n_Ԫ��4_Count := 0;
        n_Ԫ��1_Count := 0;
        v_Ԫ��1       := '+';
        n_Ԫ��2       := 0;
        v_Ԫ��3       := '+';
      End If;
    
      If n_��ҳ = 1 Then
        --��ҳ:���㷢Ʊ�ż���ص�
        For I In 1 .. c_No.Count Loop
          c_Invoce.Extend;
          c_Invoce(c_Invoce.Count).Ʊ�� := v_��Ʊ��;
          c_Invoce(c_Invoce.Count).No := c_No(I).No;
          c_Invoce(c_Invoce.Count).��� := Case
                                           When Instr(c_No(I).���, ',') > 0 Then
                                            Substr(c_No(I).���, 2)
                                           Else
                                            c_No(I).���
                                         End;
          c_Invoce(c_Invoce.Count).������� := n_�������;
        End Loop;
      
        ���η�Ʊ����_Out := ���η�Ʊ����_Out + 1;
        ���δ�ӡ��Ʊ_Out := Nvl(���δ�ӡ��Ʊ_Out, '') || ',' || v_��Ʊ��;
        v_��Ʊ��         := Zl_Incstr(v_��Ʊ��);
        c_No.Delete;
      End If;
      If (v_Ԫ��1 <> c_Bill.Ԫ��1) Then
        n_Ԫ��1_Count := n_Ԫ��1_Count + 1;
        v_Ԫ��1       := c_Bill.Ԫ��1;
      End If;
      If (n_Ԫ��2 <> c_Bill.Ԫ��2) Then
        n_Ԫ��2_Count := n_Ԫ��2_Count + 1;
        n_Ԫ��2       := c_Bill.Ԫ��2;
      End If;
      If (v_Ԫ��3 <> c_Bill.Ԫ��3) Then
        n_Ԫ��3_Count := n_Ԫ��3_Count + 1;
        v_Ԫ��3       := c_Bill.Ԫ��3;
      End If;
      If n_�շ�ϸĿ <> 0 Then
        n_Ԫ��4_Count := n_Ԫ��4_Count + 1;
      End If;
    
      -------------------------------------------
      --���䵥�ݺż����
      n_Find := 0;
      For J In 1 .. c_No.Count Loop
        If c_No(J).No = c_Bill.���� Then
          --���ݺ���ͬ,����źϲ�
          c_No(J).��� := c_No(J).��� || ',' || c_Bill.���;
          n_Find := 1;
          Exit;
        End If;
      End Loop;
      If n_Find = 0 Then
        c_No.Extend;
        c_No(c_No.Count).No := c_Bill.����;
        c_No(c_No.Count).��� := c_No(c_No.Count).��� || ',' || c_Bill.���;
      End If;
    End Loop;
  
    --�Ƿ��з�Ʊ����
    If n_Count >= 1 Then
      --���һ����Ʊ����
      ���η�Ʊ����_Out := ���η�Ʊ����_Out + 1;
      ���δ�ӡ��Ʊ_Out := Nvl(���δ�ӡ��Ʊ_Out, '') || ',' || v_��Ʊ��;
    Else
      ���η�Ʊ����_Out := 0;
      ���δ�ӡ��Ʊ_Out := '';
    End If;
    If c_No.Count <> 0 Then
      For I In 1 .. c_No.Count Loop
        c_Invoce.Extend;
        c_Invoce(c_Invoce.Count).Ʊ�� := v_��Ʊ��;
        c_Invoce(c_Invoce.Count).No := c_No(I).No;
        If Instr(c_No(I).���, ',') > 0 Then
          c_No(I).��� := Substr(c_No(I).���, 2);
        End If;
        c_Invoce(c_Invoce.Count).��� := c_No(I).���;
        c_Invoce(c_Invoce.Count).������� := n_�������;
      End Loop;
      c_No.Delete;
    End If;
    If Instr(Nvl(���δ�ӡ��Ʊ_Out, '-'), ',') > 0 Then
      ���δ�ӡ��Ʊ_Out := Substr(���δ�ӡ��Ʊ_Out, 2);
    End If;
  End Invoice_Split_Notgroup;
  --����:�����з�����ܻ���ҳ����ʱ���ô˹���
  -------------------------------------------------------------------------------------------------------------------
  --�������
  Procedure Invoice_Split_Group
  (
    Print_Nos        t_Strlist,
    ���շ�Ʊ_In      Varchar2,
    ���δ�ӡ��Ʊ_Out Out Varchar2,
    ���η�Ʊ����_Out Out Number
  ) As
  Begin
    v_Ԫ��1          := '+';
    n_Ԫ��2          := 0;
    v_Ԫ��3          := '+';
    n_Ԫ��4          := 0;
    n_Ԫ��1_Count    := 0;
    n_Ԫ��2_Count    := 0;
    n_Ԫ��3_Count    := 0;
    n_Ԫ��4_Count    := 0;
    ���η�Ʊ����_Out := 0;
  
    c_No.Delete;
    l_Ԫ��2.Delete;
  
    --�����ݷ���Ʊ��
    If ��������_In = 3 Or ��������_In = 4 Then
      --******************************************************************************************************************************
      --�˷Ѻ��ش򰴷�Ʊ�Ŵ���(��ʼ)
      --4.�վݷ�Ŀ+�շ�ϸĿ
      If n_�ֵ��ݴ�ӡ = 0 And n_ִ�п��� = 0 And n_�վݷ�Ŀ <> 0 And n_�շ�ϸĿ <> 0 Then
        v_Ԫ��3 := '+';
        c_Split_��Ŀ.Delete;
        For c_��ҳ In (With c_���� As
                        (Select Decode(n_�ֵ��ݴ�ӡ, 0, '-', a.No) As Ԫ��1, Decode(n_ִ�п���, 0, 0, a.ִ�в���id) As Ԫ��2,
                               Decode(n_�վݷ�Ŀ, 0, '-', a.�վݷ�Ŀ) As Ԫ��3, Decode(n_�շ�ϸĿ, 0, 0, a.�շ�ϸĿid) As Ԫ��4, a.No As ����,
                               Nvl(a.�۸񸸺�, a.���) As ���, Sum(Nvl(a.ʵ�ս��, 0)) As ʵ�ս��
                        From ������ü�¼ A,
                             (Select /*+cardinality(j,10)*/
                                NO, ���
                               From Ʊ�ݴ�ӡ��ϸ M, Table(f_Str2list(���շ�Ʊ_In)) J
                               Where m.Ʊ�� = j.Column_Value) B
                        Where Mod(��¼����, 10) = 1 And a.No = b.No And
                              Instr(',' || b.��� || ',', ',' || Nvl(a.�۸񸸺�, a.���) || ',') > 0 And
                              Decode(n_�����, 1, Nvl(a.���ӱ�־, 0), 0) <> 9
                        Group By a.No, a.ִ�в���id, a.�վݷ�Ŀ, a.�շ�ϸĿid, Nvl(a.�۸񸸺�, a.���)
                        Having Sum(Nvl(a.ʵ�ս��, 0)) <> 0)
                       Select a.Ԫ��3, Count(*) As ���� From c_���� A Group By Ԫ��3 Order By Ԫ��3) Loop
          If (v_Ԫ��3 <> c_��ҳ.Ԫ��3 And n_Ԫ��3_Count >= n_�վݷ�Ŀ And n_�վݷ�Ŀ <> 0) Then
            If v_Ԫ��3 <> '+' Then
              c_Split_��Ŀ.Extend;
              For J In 1 .. l_Ԫ��3.Count Loop
                --���ݺ���ͬ,����źϲ�
                c_Split_��Ŀ(c_Split_��Ŀ.Count).Ԫ��3�� := c_Split_��Ŀ(c_Split_��Ŀ.Count).Ԫ��3�� || ',' || l_Ԫ��3(J);
              End Loop;
              v_Ԫ��3       := '+';
              n_Ԫ��3_Count := 0;
              l_Ԫ��3.Delete;
            End If;
          End If;
          If (v_Ԫ��3 <> c_��ҳ.Ԫ��3) Then
            n_Ԫ��3_Count := n_Ԫ��3_Count + 1;
            v_Ԫ��3       := c_��ҳ.Ԫ��3;
            l_Ԫ��3.Extend;
            l_Ԫ��3(l_Ԫ��3.Count) := v_Ԫ��3;
          End If;
        End Loop;
        If l_Ԫ��3.Count <> 0 Then
          c_Split_��Ŀ.Extend;
          For J In 1 .. l_Ԫ��3.Count Loop
            --���ݺ���ͬ,����źϲ�
            c_Split_��Ŀ(c_Split_��Ŀ.Count).Ԫ��3�� := c_Split_��Ŀ(c_Split_��Ŀ.Count).Ԫ��3�� || ',' || l_Ԫ��3(J);
          End Loop;
        End If;
        n_������� := 0;
        For I In 1 .. c_Split_��Ŀ.Count Loop
          c_No.Delete;
          n_�������    := n_������� + 1;
          n_Ԫ��4_Count := 0;
          For c_��ҳ In (With c_���� As
                          (Select Decode(n_�ֵ��ݴ�ӡ, 0, '-', a.No) As Ԫ��1, Decode(n_ִ�п���, 0, 0, a.ִ�в���id) As Ԫ��2,
                                 Decode(n_�վݷ�Ŀ, 0, '-', a.�վݷ�Ŀ) As Ԫ��3, Decode(n_�շ�ϸĿ, 0, 0, a.�շ�ϸĿid) As Ԫ��4, a.No As ����,
                                 Nvl(a.�۸񸸺�, a.���) As ���, Sum(Nvl(a.ʵ�ս��, 0)) As ʵ�ս��
                          From ������ü�¼ A,
                               (Select /*+cardinality(j,10)*/
                                  NO, ���
                                 From Ʊ�ݴ�ӡ��ϸ M, Table(f_Str2list(���շ�Ʊ_In)) J
                                 Where m.Ʊ�� = j.Column_Value) B
                          Where Mod(��¼����, 10) = 1 And a.No = b.No And
                                Instr(',' || b.��� || ',', ',' || Nvl(a.�۸񸸺�, a.���) || ',') > 0 And
                                Decode(n_�����, 1, Nvl(a.���ӱ�־, 0), 0) <> 9
                          Group By a.No, a.ִ�в���id, a.�վݷ�Ŀ, a.�շ�ϸĿid, Nvl(a.�۸񸸺�, a.���)
                          Having Sum(Nvl(a.ʵ�ս��, 0)) <> 0)
                         Select m.Ԫ��1, Ԫ��2, Ԫ��3, m.Ԫ��4, m.����, m.���, Count(*) As ����
                         From c_���� M
                         Where Instr(',' || c_Split_��Ŀ(I).Ԫ��3�� || ',', ',' || m.Ԫ��3 || ',') > 0
                         Group By m.Ԫ��1, Ԫ��2, m.Ԫ��4, Ԫ��3, m.����, m.���
                         Order By m.Ԫ��1, Ԫ��2, m.Ԫ��4, Ԫ��3, m.����, m.���) Loop
            If n_Ԫ��4_Count >= n_�շ�ϸĿ And n_�շ�ϸĿ <> 0 Then
              --��ҳ:���㷢Ʊ�ż���ص�
              For J In 1 .. c_No.Count Loop
                c_Invoce.Extend;
                c_Invoce(c_Invoce.Count).Ʊ�� := v_��Ʊ��;
                c_Invoce(c_Invoce.Count).No := c_No(J).No;
                c_Invoce(c_Invoce.Count).��� := Case
                                                 When Instr(c_No(J).���, ',') > 0 Then
                                                  Substr(c_No(J).���, 2)
                                                 Else
                                                  c_No(J).���
                                               End;
                c_Invoce(c_Invoce.Count).������� := n_�������;
              End Loop;
              ���η�Ʊ����_Out := ���η�Ʊ����_Out + 1;
              ���δ�ӡ��Ʊ_Out := Nvl(���δ�ӡ��Ʊ_Out, '') || ',' || v_��Ʊ��;
              v_��Ʊ��         := Zl_Incstr(v_��Ʊ��);
              c_No.Delete;
              n_Ԫ��4_Count := 0;
              --��ҳ
            End If;
            n_Ԫ��4_Count := n_Ԫ��4_Count + 1;
            -------------------------------------------
            --���䵥�ݺż����
            n_Find := 0;
            For J In 1 .. c_No.Count Loop
              If c_No(J).No = c_��ҳ.���� Then
                --���ݺ���ͬ,����źϲ�
                c_No(J).��� := c_No(J).��� || ',' || c_��ҳ.���;
                n_Find := 1;
                Exit;
              End If;
            End Loop;
            If n_Find = 0 Then
              c_No.Extend;
              c_No(c_No.Count).No := c_��ҳ.����;
              c_No(c_No.Count).��� := c_No(c_No.Count).��� || ',' || c_��ҳ.���;
            End If;
          End Loop;
          If c_No.Count <> 0 Then
            --��ҳ:���㷢Ʊ�ż���ص�
            For J In 1 .. c_No.Count Loop
              c_Invoce.Extend;
              c_Invoce(c_Invoce.Count).Ʊ�� := v_��Ʊ��;
              c_Invoce(c_Invoce.Count).No := c_No(J).No;
              c_Invoce(c_Invoce.Count).��� := Case
                                               When Instr(c_No(J).���, ',') > 0 Then
                                                Substr(c_No(J).���, 2)
                                               Else
                                                c_No(J).���
                                             End;
              c_Invoce(c_Invoce.Count).������� := n_�������;
            End Loop;
            ���η�Ʊ����_Out := ���η�Ʊ����_Out + 1;
            ���δ�ӡ��Ʊ_Out := Nvl(���δ�ӡ��Ʊ_Out, '') || ',' || v_��Ʊ��;
            v_��Ʊ��         := Zl_Incstr(v_��Ʊ��);
            c_No.Delete;
          End If;
        End Loop;
      End If;
    
      If (n_�ֵ��ݴ�ӡ = 1 Or n_ִ�п��� > 0) And (n_�վݷ�Ŀ <> 0 Or n_�շ�ϸĿ <> 0) Then
        n_Ԫ��2_Count := 0;
        v_Ԫ��1       := '+';
        n_Ԫ��2       := 0;
        c_Split_No.Delete;
        For c_��ҳ In (With c_���� As
                        (Select Decode(n_�ֵ��ݴ�ӡ, 0, '-', a.No) As Ԫ��1, Decode(n_ִ�п���, 0, 0, a.ִ�в���id) As Ԫ��2,
                               Decode(n_�վݷ�Ŀ, 0, '-', a.�վݷ�Ŀ) As Ԫ��3, Decode(n_�շ�ϸĿ, 0, 0, a.�շ�ϸĿid) As Ԫ��4, a.No As ����,
                               Nvl(a.�۸񸸺�, a.���) As ���, Sum(Nvl(a.ʵ�ս��, 0)) As ʵ�ս��
                        From ������ü�¼ A,
                             (Select /*+cardinality(j,10)*/
                                NO, ���
                               From Ʊ�ݴ�ӡ��ϸ M, Table(f_Str2list(���շ�Ʊ_In)) J
                               Where m.Ʊ�� = j.Column_Value) B
                        Where Mod(��¼����, 10) = 1 And a.No = b.No And
                              Instr(',' || b.��� || ',', ',' || Nvl(a.�۸񸸺�, a.���) || ',') > 0 And
                             
                              Decode(n_�����, 1, Nvl(a.���ӱ�־, 0), 0) <> 9
                        Group By a.No, a.ִ�в���id, a.�վݷ�Ŀ, a.�շ�ϸĿid, Nvl(a.�۸񸸺�, a.���)
                        Having Sum(Nvl(a.ʵ�ս��, 0)) <> 0)
                       Select a.Ԫ��1, a.Ԫ��2, b.����, Count(*) As ����
                       From c_���� A, ���ű� B
                       Where a.Ԫ��2 = b.Id(+)
                       Group By Ԫ��1, b.����, Ԫ��2
                       Order By Ԫ��1, b.����, Ԫ��2) Loop
          If (v_Ԫ��1 <> c_��ҳ.Ԫ��1) Or (n_Ԫ��2 <> c_��ҳ.Ԫ��2 And n_Ԫ��2_Count >= n_ִ�п��� And n_ִ�п��� <> 0) Then
            c_Split_No.Extend;
            n_Ԫ��2_Count := 0;
            v_Ԫ��1       := '+';
            n_Ԫ��2       := 0;
          End If;
          If (v_Ԫ��1 <> c_��ҳ.Ԫ��1) Then
            v_Ԫ��1 := c_��ҳ.Ԫ��1;
            c_Split_No(c_Split_No.Count).Ԫ��1 := v_Ԫ��1;
          End If;
          If (n_Ԫ��2 <> c_��ҳ.Ԫ��2) Then
            n_Ԫ��2_Count := n_Ԫ��2_Count + 1;
            n_Ԫ��2 := c_��ҳ.Ԫ��2;
            c_Split_No(c_Split_No.Count).Ԫ��2�� := c_Split_No(c_Split_No.Count).Ԫ��2�� || ',' || n_Ԫ��2;
          End If;
        End Loop;
      End If;
    
      --6.(no Or ִ�п���)+�շ�ϸĿ
      If (n_�ֵ��ݴ�ӡ = 0 Or n_ִ�п��� > 0) And n_�վݷ�Ŀ = 0 And n_�շ�ϸĿ <> 0 Then
      
        For I In 1 .. c_Split_No.Count Loop
          v_Ԫ��3 := '+';
          --ֻ����ҳ���ܵ�,���й������
          n_�������    := n_������� + 1;
          n_Ԫ��4_Count := 0;
          For c_��ҳ In (With c_���� As
                          (Select Decode(n_�ֵ��ݴ�ӡ, 0, '-', a.No) As Ԫ��1, Decode(n_ִ�п���, 0, 0, a.ִ�в���id) As Ԫ��2,
                                 Decode(n_�վݷ�Ŀ, 0, '-', a.�վݷ�Ŀ) As Ԫ��3, Decode(n_�շ�ϸĿ, 0, 0, a.�շ�ϸĿid) As Ԫ��4, a.No As ����,
                                 Nvl(a.�۸񸸺�, a.���) As ���, Sum(Nvl(a.ʵ�ս��, 0)) As ʵ�ս��
                          From ������ü�¼ A,
                               (Select /*+cardinality(j,10)*/
                                  NO, ���
                                 From Ʊ�ݴ�ӡ��ϸ M, Table(f_Str2list(���շ�Ʊ_In)) J
                                 Where m.Ʊ�� = j.Column_Value) B
                          Where Mod(��¼����, 10) = 1 And a.No = b.No And
                                Instr(',' || b.��� || ',', ',' || Nvl(a.�۸񸸺�, a.���) || ',') > 0 And
                               
                                Decode(n_�����, 1, Nvl(a.���ӱ�־, 0), 0) <> 9
                          Group By a.No, a.ִ�в���id, a.�վݷ�Ŀ, a.�շ�ϸĿid, Nvl(a.�۸񸸺�, a.���)
                          Having Sum(Nvl(a.ʵ�ս��, 0)) <> 0)
                         Select Ԫ��1, Ԫ��2, Ԫ��3, a.Ԫ��4, a.����, a.���, Count(*) As ����
                         From c_���� A
                         Where a.Ԫ��1 = c_Split_No(I).Ԫ��1 And
                               Instr(',' || c_Split_No(I).Ԫ��2�� || ',', ',' || a.Ԫ��2 || ',') > 0
                         Group By Ԫ��1, Ԫ��2, Ԫ��4, Ԫ��3, ����, a.���
                         Order By Ԫ��1, Ԫ��2, Ԫ��4, ����, ���) Loop
            If n_Ԫ��4_Count >= n_�շ�ϸĿ And n_�շ�ϸĿ <> 0 Then
              --���䵥��
              If c_No.Count <> 0 Then
                --��ҳ:���㷢Ʊ�ż���ص�
                For J In 1 .. c_No.Count Loop
                  c_Invoce.Extend;
                  c_Invoce(c_Invoce.Count).Ʊ�� := v_��Ʊ��;
                  c_Invoce(c_Invoce.Count).No := c_No(J).No;
                  c_Invoce(c_Invoce.Count).��� := Case
                                                   When Instr(c_No(J).���, ',') > 0 Then
                                                    Substr(c_No(J).���, 2)
                                                   Else
                                                    c_No(J).���
                                                 End;
                  c_Invoce(c_Invoce.Count).������� := n_Ԫ��4_Count;
                End Loop;
                ���η�Ʊ����_Out := ���η�Ʊ����_Out + 1;
                ���δ�ӡ��Ʊ_Out := Nvl(���δ�ӡ��Ʊ_Out, '') || ',' || v_��Ʊ��;
                v_��Ʊ��         := Zl_Incstr(v_��Ʊ��);
                c_No.Delete;
              End If;
              n_Ԫ��4_Count := 0;
            End If;
            n_Ԫ��4_Count := n_Ԫ��4_Count + 1;
          
            -------------------------------------------
            --���䵥�ݺż����
            n_Find := 0;
            For J In 1 .. c_No.Count Loop
              If c_No(J).No = c_��ҳ.���� Then
                --���ݺ���ͬ,����źϲ�
                c_No(J).��� := c_No(J).��� || ',' || c_��ҳ.���;
                n_Find := 1;
                Exit;
              End If;
            End Loop;
            If n_Find = 0 Then
              c_No.Extend;
              c_No(c_No.Count).No := c_��ҳ.����;
              c_No(c_No.Count).��� := c_No(c_No.Count).��� || ',' || c_��ҳ.���;
            End If;
          End Loop;
          --���䵥��
          If c_No.Count <> 0 Then
            --��ҳ:���㷢Ʊ�ż���ص�
            For J In 1 .. c_No.Count Loop
              c_Invoce.Extend;
              c_Invoce(c_Invoce.Count).Ʊ�� := v_��Ʊ��;
              c_Invoce(c_Invoce.Count).No := c_No(J).No;
              c_Invoce(c_Invoce.Count).��� := Case
                                               When Instr(c_No(J).���, ',') > 0 Then
                                                Substr(c_No(J).���, 2)
                                               Else
                                                c_No(J).���
                                             End;
              c_Invoce(c_Invoce.Count).������� := n_Ԫ��4_Count;
            End Loop;
            ���η�Ʊ����_Out := ���η�Ʊ����_Out + 1;
            ���δ�ӡ��Ʊ_Out := Nvl(���δ�ӡ��Ʊ_Out, '') || ',' || v_��Ʊ��;
            v_��Ʊ��         := Zl_Incstr(v_��Ʊ��);
            c_No.Delete;
          End If;
        End Loop;
      End If;
    
      --7.(no Or ִ�п���)+�վݷ�Ŀ+�շ�ϸĿ
      n_������� := 0;
      If (n_�ֵ��ݴ�ӡ = 0 Or n_ִ�п��� > 0) And n_�վݷ�Ŀ <> 0 And n_�շ�ϸĿ <> 0 Then
        c_Split_��Ŀ.Delete;
        For I In 1 .. c_Split_No.Count Loop
        
          n_�������    := n_������� + 1;
          v_Ԫ��3       := '+';
          n_Ԫ��3_Count := 0;
          l_Ԫ��3.Delete;
          For c_��ҳ In (With c_���� As
                          (Select Decode(n_�ֵ��ݴ�ӡ, 0, '-', a.No) As Ԫ��1, Decode(n_ִ�п���, 0, 0, a.ִ�в���id) As Ԫ��2,
                                 Decode(n_�վݷ�Ŀ, 0, '-', a.�վݷ�Ŀ) As Ԫ��3, Decode(n_�շ�ϸĿ, 0, 0, a.�շ�ϸĿid) As Ԫ��4, a.No As ����,
                                 Nvl(a.�۸񸸺�, a.���) As ���, Sum(Nvl(a.ʵ�ս��, 0)) As ʵ�ս��
                          From ������ü�¼ A,
                               (Select /*+cardinality(j,10)*/
                                  NO, ���
                                 From Ʊ�ݴ�ӡ��ϸ M, Table(f_Str2list(���շ�Ʊ_In)) J
                                 Where m.Ʊ�� = j.Column_Value) B
                          Where Mod(��¼����, 10) = 1 And a.No = b.No And
                                Instr(',' || b.��� || ',', ',' || Nvl(a.�۸񸸺�, a.���) || ',') > 0 And
                               
                                Decode(n_�����, 1, Nvl(a.���ӱ�־, 0), 0) <> 9
                          Group By a.No, a.ִ�в���id, a.�վݷ�Ŀ, a.�շ�ϸĿid, Nvl(a.�۸񸸺�, a.���)
                          Having Sum(Nvl(a.ʵ�ս��, 0)) <> 0)
                         Select a.Ԫ��3, Count(*) As ����
                         From c_���� A
                         Where a.Ԫ��1 = c_Split_No(I).Ԫ��1 And
                               Instr(',' || c_Split_No(I).Ԫ��2�� || ',', ',' || a.Ԫ��2 || ',') > 0
                         Group By Ԫ��3
                         Order By Ԫ��3) Loop
          
            If (v_Ԫ��3 <> c_��ҳ.Ԫ��3 And n_Ԫ��3_Count >= n_�վݷ�Ŀ And n_�վݷ�Ŀ <> 0) Then
              If v_Ԫ��3 <> '+' Then
                c_Split_��Ŀ.Extend;
                c_Split_��Ŀ(c_Split_��Ŀ.Count).Ԫ��1 := c_Split_No(I).Ԫ��1;
                c_Split_��Ŀ(c_Split_��Ŀ.Count).Ԫ��2�� := c_Split_No(I).Ԫ��2��;
                c_Split_��Ŀ(c_Split_��Ŀ.Count).������� := n_�������;
                For J In 1 .. l_Ԫ��3.Count Loop
                  --���ݺ���ͬ,����źϲ�
                  c_Split_��Ŀ(c_Split_��Ŀ.Count).Ԫ��3�� := c_Split_��Ŀ(c_Split_��Ŀ.Count).Ԫ��3�� || ',' || l_Ԫ��3(J);
                End Loop;
              End If;
              v_Ԫ��3       := '+';
              n_Ԫ��3_Count := 0;
              l_Ԫ��3.Delete;
            End If;
            If (v_Ԫ��3 <> c_��ҳ.Ԫ��3) Then
              n_Ԫ��3_Count := n_Ԫ��3_Count + 1;
              v_Ԫ��3       := c_��ҳ.Ԫ��3;
              l_Ԫ��3.Extend;
              l_Ԫ��3(l_Ԫ��3.Count) := v_Ԫ��3;
            End If;
          End Loop;
        
          If l_Ԫ��3.Count <> 0 Then
            c_Split_��Ŀ.Extend;
            c_Split_��Ŀ(c_Split_��Ŀ.Count).Ԫ��1 := c_Split_No(I).Ԫ��1;
            c_Split_��Ŀ(c_Split_��Ŀ.Count).Ԫ��2�� := c_Split_No(I).Ԫ��2��;
            c_Split_��Ŀ(c_Split_��Ŀ.Count).������� := n_�������;
            For J In 1 .. l_Ԫ��3.Count Loop
              --���ݺ���ͬ,����źϲ�
              c_Split_��Ŀ(c_Split_��Ŀ.Count).Ԫ��3�� := c_Split_��Ŀ(c_Split_��Ŀ.Count).Ԫ��3�� || ',' || l_Ԫ��3(J);
            End Loop;
          End If;
        End Loop;
      
        For I In 1 .. c_Split_��Ŀ.Count Loop
          c_No.Delete;
          n_Ԫ��4_Count := 0;
          For c_��ҳ In (With c_���� As
                          (Select Decode(n_�ֵ��ݴ�ӡ, 0, '-', a.No) As Ԫ��1, Decode(n_ִ�п���, 0, 0, a.ִ�в���id) As Ԫ��2,
                                 Decode(n_�վݷ�Ŀ, 0, '-', a.�վݷ�Ŀ) As Ԫ��3, Decode(n_�շ�ϸĿ, 0, 0, a.�շ�ϸĿid) As Ԫ��4, a.No As ����,
                                 Nvl(a.�۸񸸺�, a.���) As ���, Sum(Nvl(a.ʵ�ս��, 0)) As ʵ�ս��
                          From ������ü�¼ A,
                               (Select /*+cardinality(j,10)*/
                                  NO, ���
                                 From Ʊ�ݴ�ӡ��ϸ M, Table(f_Str2list(���շ�Ʊ_In)) J
                                 Where m.Ʊ�� = j.Column_Value) B
                          Where Mod(��¼����, 10) = 1 And a.No = b.No And
                                Instr(',' || b.��� || ',', ',' || Nvl(a.�۸񸸺�, a.���) || ',') > 0 And
                               
                                Decode(n_�����, 1, Nvl(a.���ӱ�־, 0), 0) <> 9
                          Group By a.No, a.ִ�в���id, a.�վݷ�Ŀ, a.�շ�ϸĿid, Nvl(a.�۸񸸺�, a.���)
                          Having Sum(Nvl(a.ʵ�ս��, 0)) <> 0)
                         Select Ԫ��1, Ԫ��2, Ԫ��3, a.Ԫ��4, a.����, a.���, Count(*) As ����
                         From c_���� A
                         Where a.Ԫ��1 = c_Split_��Ŀ(I).Ԫ��1 And
                               Instr(',' || c_Split_��Ŀ(I).Ԫ��2�� || ',', ',' || a.Ԫ��2 || ',') > 0 And
                               Instr(',' || c_Split_��Ŀ(I).Ԫ��3�� || ',', ',' || a.Ԫ��3 || ',') > 0
                         Group By Ԫ��1, Ԫ��2, Ԫ��4, Ԫ��3, a.����, a.���
                         Order By Ԫ��1, Ԫ��2, Ԫ��4, Ԫ��3, ����, ���) Loop
            If (n_Ԫ��4_Count >= n_�շ�ϸĿ And n_�շ�ϸĿ <> 0) Then
              --���䵥��
              If c_No.Count <> 0 Then
                --��ҳ:���㷢Ʊ�ż���ص�
                For J In 1 .. c_No.Count Loop
                  c_Invoce.Extend;
                  c_Invoce(c_Invoce.Count).Ʊ�� := v_��Ʊ��;
                  c_Invoce(c_Invoce.Count).No := c_No(J).No;
                  c_Invoce(c_Invoce.Count).��� := Case
                                                   When Instr(c_No(J).���, ',') > 0 Then
                                                    Substr(c_No(J).���, 2)
                                                   Else
                                                    c_No(J).���
                                                 End;
                  c_Invoce(c_Invoce.Count).������� := c_Split_��Ŀ(I).�������;
                End Loop;
                ���η�Ʊ����_Out := ���η�Ʊ����_Out + 1;
                ���δ�ӡ��Ʊ_Out := Nvl(���δ�ӡ��Ʊ_Out, '') || ',' || v_��Ʊ��;
                v_��Ʊ��         := Zl_Incstr(v_��Ʊ��);
                c_No.Delete;
              End If;
              n_Ԫ��4_Count := 0;
            End If;
            n_Ԫ��4_Count := n_Ԫ��4_Count + 1;
            -------------------------------------------
            --���䵥�ݺż����
            n_Find := 0;
            For J In 1 .. c_No.Count Loop
              If c_No(J).No = c_��ҳ.���� Then
                --���ݺ���ͬ,����źϲ�
                c_No(J).��� := c_No(J).��� || ',' || c_��ҳ.���;
                n_Find := 1;
                Exit;
              End If;
            End Loop;
            If n_Find = 0 Then
              c_No.Extend;
              c_No(c_No.Count).No := c_��ҳ.����;
              c_No(c_No.Count).��� := c_No(c_No.Count).��� || ',' || c_��ҳ.���;
            End If;
          End Loop;
        
          --���䵥��
          If c_No.Count <> 0 Then
            --��ҳ:���㷢Ʊ�ż���ص�
            For J In 1 .. c_No.Count Loop
              c_Invoce.Extend;
              c_Invoce(c_Invoce.Count).Ʊ�� := v_��Ʊ��;
              c_Invoce(c_Invoce.Count).No := c_No(J).No;
              c_Invoce(c_Invoce.Count).��� := Case
                                               When Instr(c_No(J).���, ',') > 0 Then
                                                Substr(c_No(J).���, 2)
                                               Else
                                                c_No(J).���
                                             End;
              c_Invoce(c_Invoce.Count).������� := c_Split_��Ŀ(I).�������;
            End Loop;
            ���η�Ʊ����_Out := ���η�Ʊ����_Out + 1;
            ���δ�ӡ��Ʊ_Out := Nvl(���δ�ӡ��Ʊ_Out, '') || ',' || v_��Ʊ��;
            v_��Ʊ��         := Zl_Incstr(v_��Ʊ��);
            c_No.Delete;
          End If;
        End Loop;
      End If;
    
      --�˷Ѻ��ش򰴷�Ʊ�Ŵ���(����)
      --******************************************************************************************************************************
      If Instr(Nvl(���δ�ӡ��Ʊ_Out, '-'), ',') > 0 Then
        ���δ�ӡ��Ʊ_Out := Substr(���δ�ӡ��Ʊ_Out, 2);
      End If;
      Return;
    
    End If;
  
    --******************************************************************************************************************************
    --�����ǰ��������䵥��(��ʼ)
    --4.�վݷ�Ŀ+�շ�ϸĿ
    If n_�ֵ��ݴ�ӡ = 0 And n_ִ�п��� = 0 And n_�վݷ�Ŀ <> 0 And n_�շ�ϸĿ <> 0 Then
      v_Ԫ��3 := '+';
      c_Split_��Ŀ.Delete;
    
      For c_��ҳ In (With c_���� As
                      (Select /*+cardinality(b,10)*/
                       Decode(n_�ֵ��ݴ�ӡ, 0, '-', a.No) As Ԫ��1, Decode(n_ִ�п���, 0, 0, a.ִ�в���id) As Ԫ��2,
                       Decode(n_�վݷ�Ŀ, 0, '-', a.�վݷ�Ŀ) As Ԫ��3, Decode(n_�շ�ϸĿ, 0, 0, a.�շ�ϸĿid) As Ԫ��4, a.No As ����,
                       Nvl(a.�۸񸸺�, a.���) As ���, Sum(Nvl(a.ʵ�ս��, 0)) As ʵ�ս��
                      From ������ü�¼ A, Table(Print_Nos) B
                      Where Mod(��¼����, 10) = 1 And a.No = b.Column_Value And Decode(n_�����, 1, Nvl(a.���ӱ�־, 0), 0) <> 9
                      Group By a.No, a.ִ�в���id, a.�վݷ�Ŀ, a.�շ�ϸĿid, Nvl(a.�۸񸸺�, a.���)
                      Having Sum(Nvl(a.ʵ�ս��, 0)) <> 0)
                     Select a.Ԫ��3, Count(*) As ���� From c_���� A Group By Ԫ��3 Order By Ԫ��3) Loop
        If (v_Ԫ��3 <> c_��ҳ.Ԫ��3 And n_Ԫ��3_Count >= n_�վݷ�Ŀ And n_�վݷ�Ŀ <> 0) Then
          If v_Ԫ��3 <> '+' Then
            c_Split_��Ŀ.Extend;
            For J In 1 .. l_Ԫ��3.Count Loop
              --���ݺ���ͬ,����źϲ�
              c_Split_��Ŀ(c_Split_��Ŀ.Count).Ԫ��3�� := c_Split_��Ŀ(c_Split_��Ŀ.Count).Ԫ��3�� || ',' || l_Ԫ��3(J);
            End Loop;
            v_Ԫ��3       := '+';
            n_Ԫ��3_Count := 0;
            l_Ԫ��3.Delete;
          End If;
        End If;
        If (v_Ԫ��3 <> c_��ҳ.Ԫ��3) Then
          n_Ԫ��3_Count := n_Ԫ��3_Count + 1;
          v_Ԫ��3       := c_��ҳ.Ԫ��3;
          l_Ԫ��3.Extend;
          l_Ԫ��3(l_Ԫ��3.Count) := v_Ԫ��3;
        End If;
      End Loop;
      If l_Ԫ��3.Count <> 0 Then
        c_Split_��Ŀ.Extend;
        For J In 1 .. l_Ԫ��3.Count Loop
          --���ݺ���ͬ,����źϲ�
          c_Split_��Ŀ(c_Split_��Ŀ.Count).Ԫ��3�� := c_Split_��Ŀ(c_Split_��Ŀ.Count).Ԫ��3�� || ',' || l_Ԫ��3(J);
        End Loop;
      End If;
      n_������� := 0;
      For I In 1 .. c_Split_��Ŀ.Count Loop
        c_No.Delete;
        n_�������    := n_������� + 1;
        n_Ԫ��4_Count := 0;
        For c_��ҳ In (With c_���� As
                        (Select /*+cardinality(b,10)*/
                         Decode(n_�ֵ��ݴ�ӡ, 0, '-', a.No) As Ԫ��1, Decode(n_ִ�п���, 0, 0, a.ִ�в���id) As Ԫ��2,
                         Decode(n_�վݷ�Ŀ, 0, '-', a.�վݷ�Ŀ) As Ԫ��3, Decode(n_�շ�ϸĿ, 0, 0, a.�շ�ϸĿid) As Ԫ��4, a.No As ����,
                         Nvl(a.�۸񸸺�, a.���) As ���, Sum(Nvl(a.ʵ�ս��, 0)) As ʵ�ս��
                        From ������ü�¼ A, Table(Print_Nos) B
                        Where Mod(��¼����, 10) = 1 And a.No = b.Column_Value And Decode(n_�����, 1, Nvl(a.���ӱ�־, 0), 0) <> 9
                        Group By a.No, a.ִ�в���id, a.�վݷ�Ŀ, a.�շ�ϸĿid, Nvl(a.�۸񸸺�, a.���)
                        Having Sum(Nvl(a.ʵ�ս��, 0)) <> 0)
                       Select m.Ԫ��1, Ԫ��2, Ԫ��3, m.Ԫ��4, m.����, m.���, Count(*) As ����
                       From c_���� M
                       Where Instr(',' || c_Split_��Ŀ(I).Ԫ��3�� || ',', ',' || m.Ԫ��3 || ',') > 0
                       Group By m.Ԫ��1, Ԫ��2, m.Ԫ��4, Ԫ��3, m.����, m.���
                       Order By m.Ԫ��1, Ԫ��2, m.Ԫ��4, Ԫ��3, m.����, m.���) Loop
          If n_Ԫ��4_Count >= n_�շ�ϸĿ And n_�շ�ϸĿ <> 0 Then
            --��ҳ:���㷢Ʊ�ż���ص�
            For J In 1 .. c_No.Count Loop
              c_Invoce.Extend;
              c_Invoce(c_Invoce.Count).Ʊ�� := v_��Ʊ��;
              c_Invoce(c_Invoce.Count).No := c_No(J).No;
              c_Invoce(c_Invoce.Count).��� := Case
                                               When Instr(c_No(J).���, ',') > 0 Then
                                                Substr(c_No(J).���, 2)
                                               Else
                                                c_No(J).���
                                             End;
              c_Invoce(c_Invoce.Count).������� := n_�������;
            End Loop;
            ���η�Ʊ����_Out := ���η�Ʊ����_Out + 1;
            ���δ�ӡ��Ʊ_Out := Nvl(���δ�ӡ��Ʊ_Out, '') || ',' || v_��Ʊ��;
            v_��Ʊ��         := Zl_Incstr(v_��Ʊ��);
            c_No.Delete;
            n_Ԫ��4_Count := 0;
            --��ҳ
          End If;
          n_Ԫ��4_Count := n_Ԫ��4_Count + 1;
          -------------------------------------------
          --���䵥�ݺż����
          n_Find := 0;
          For J In 1 .. c_No.Count Loop
            If c_No(J).No = c_��ҳ.���� Then
              --���ݺ���ͬ,����źϲ�
              c_No(J).��� := c_No(J).��� || ',' || c_��ҳ.���;
              n_Find := 1;
              Exit;
            End If;
          End Loop;
          If n_Find = 0 Then
            c_No.Extend;
            c_No(c_No.Count).No := c_��ҳ.����;
            c_No(c_No.Count).��� := c_No(c_No.Count).��� || ',' || c_��ҳ.���;
          End If;
        End Loop;
        If c_No.Count <> 0 Then
          --��ҳ:���㷢Ʊ�ż���ص�
          For J In 1 .. c_No.Count Loop
            c_Invoce.Extend;
            c_Invoce(c_Invoce.Count).Ʊ�� := v_��Ʊ��;
            c_Invoce(c_Invoce.Count).No := c_No(J).No;
            c_Invoce(c_Invoce.Count).��� := Case
                                             When Instr(c_No(J).���, ',') > 0 Then
                                              Substr(c_No(J).���, 2)
                                             Else
                                              c_No(J).���
                                           End;
            c_Invoce(c_Invoce.Count).������� := n_�������;
          End Loop;
          ���η�Ʊ����_Out := ���η�Ʊ����_Out + 1;
          ���δ�ӡ��Ʊ_Out := Nvl(���δ�ӡ��Ʊ_Out, '') || ',' || v_��Ʊ��;
          v_��Ʊ��         := Zl_Incstr(v_��Ʊ��);
          c_No.Delete;
        End If;
      End Loop;
    End If;
  
    If (n_�ֵ��ݴ�ӡ = 1 Or n_ִ�п��� > 0) And (n_�վݷ�Ŀ <> 0 Or n_�շ�ϸĿ <> 0) Then
      n_Ԫ��2_Count := 0;
      v_Ԫ��1       := '+';
      n_Ԫ��2       := 0;
      c_Split_No.Delete;
      For c_��ҳ In (With c_���� As
                      (Select /*+cardinality(b,10)*/
                       Decode(n_�ֵ��ݴ�ӡ, 0, '-', a.No) As Ԫ��1, Decode(n_ִ�п���, 0, 0, a.ִ�в���id) As Ԫ��2,
                       Decode(n_�վݷ�Ŀ, 0, '-', a.�վݷ�Ŀ) As Ԫ��3, Decode(n_�շ�ϸĿ, 0, 0, a.�շ�ϸĿid) As Ԫ��4, a.No As ����,
                       Nvl(a.�۸񸸺�, a.���) As ���, Sum(Nvl(a.ʵ�ս��, 0)) As ʵ�ս��
                      From ������ü�¼ A, Table(Print_Nos) B
                      Where Mod(��¼����, 10) = 1 And a.No = b.Column_Value And Decode(n_�����, 1, Nvl(a.���ӱ�־, 0), 0) <> 9
                      Group By a.No, a.ִ�в���id, a.�վݷ�Ŀ, a.�շ�ϸĿid, Nvl(a.�۸񸸺�, a.���)
                      Having Sum(Nvl(a.ʵ�ս��, 0)) <> 0)
                     Select a.Ԫ��1, a.Ԫ��2, b.����, Count(*) As ����
                     From c_���� A, ���ű� B
                     Where a.Ԫ��2 = b.Id(+)
                     Group By Ԫ��1, b.����, Ԫ��2
                     Order By Ԫ��1, b.����, Ԫ��2) Loop
        If (v_Ԫ��1 <> c_��ҳ.Ԫ��1) Or (n_Ԫ��2 <> c_��ҳ.Ԫ��2 And n_Ԫ��2_Count >= n_ִ�п��� And n_ִ�п��� <> 0) Then
          c_Split_No.Extend;
          n_Ԫ��2_Count := 0;
          v_Ԫ��1       := '+';
          n_Ԫ��2       := 0;
        End If;
        If (v_Ԫ��1 <> c_��ҳ.Ԫ��1) Then
          v_Ԫ��1 := c_��ҳ.Ԫ��1;
          c_Split_No(c_Split_No.Count).Ԫ��1 := v_Ԫ��1;
        End If;
        If (n_Ԫ��2 <> c_��ҳ.Ԫ��2) Then
          n_Ԫ��2_Count := n_Ԫ��2_Count + 1;
          n_Ԫ��2 := c_��ҳ.Ԫ��2;
          c_Split_No(c_Split_No.Count).Ԫ��2�� := c_Split_No(c_Split_No.Count).Ԫ��2�� || ',' || n_Ԫ��2;
        End If;
      End Loop;
    End If;
  
    --3.(no Or ִ�п���)+�շ�ϸĿ
    If (n_�ֵ��ݴ�ӡ = 0 Or n_ִ�п��� > 0) And n_�վݷ�Ŀ = 0 And n_�շ�ϸĿ <> 0 Then
    
      For I In 1 .. c_Split_No.Count Loop
        v_Ԫ��3 := '+';
        --ֻ����ҳ���ܵ�,���й������
        n_�������    := Nvl(n_�������, 0) + 1;
        n_Ԫ��4_Count := 0;
        For c_��ҳ In (With c_���� As
                        (Select /*+cardinality(b,10)*/
                         Decode(n_�ֵ��ݴ�ӡ, 0, '-', a.No) As Ԫ��1, Decode(n_ִ�п���, 0, 0, a.ִ�в���id) As Ԫ��2,
                         Decode(n_�վݷ�Ŀ, 0, '-', a.�վݷ�Ŀ) As Ԫ��3, Decode(n_�շ�ϸĿ, 0, 0, a.�շ�ϸĿid) As Ԫ��4, a.No As ����,
                         Nvl(a.�۸񸸺�, a.���) As ���, Sum(Nvl(a.ʵ�ս��, 0)) As ʵ�ս��
                        From ������ü�¼ A, Table(Print_Nos) B
                        Where Mod(��¼����, 10) = 1 And a.No = b.Column_Value And Decode(n_�����, 1, Nvl(a.���ӱ�־, 0), 0) <> 9
                        Group By a.No, a.ִ�в���id, a.�վݷ�Ŀ, a.�շ�ϸĿid, Nvl(a.�۸񸸺�, a.���)
                        Having Sum(Nvl(a.ʵ�ս��, 0)) <> 0)
                       Select Ԫ��1, Ԫ��2, Ԫ��3, a.Ԫ��4, a.����, a.���, Count(*) As ����
                       From c_���� A
                       Where a.Ԫ��1 = c_Split_No(I).Ԫ��1 And
                             Instr(',' || c_Split_No(I).Ԫ��2�� || ',', ',' || a.Ԫ��2 || ',') > 0
                       Group By Ԫ��1, Ԫ��2, Ԫ��4, Ԫ��3, ����, a.���
                       Order By Ԫ��1, Ԫ��2, Ԫ��4, ����, ���) Loop
          If n_Ԫ��4_Count >= n_�շ�ϸĿ And n_�շ�ϸĿ <> 0 Then
            --���䵥��
            If c_No.Count <> 0 Then
              --��ҳ:���㷢Ʊ�ż���ص�
              For J In 1 .. c_No.Count Loop
                c_Invoce.Extend;
                c_Invoce(c_Invoce.Count).Ʊ�� := v_��Ʊ��;
                c_Invoce(c_Invoce.Count).No := c_No(J).No;
                c_Invoce(c_Invoce.Count).��� := Case
                                                 When Instr(c_No(J).���, ',') > 0 Then
                                                  Substr(c_No(J).���, 2)
                                                 Else
                                                  c_No(J).���
                                               End;
                c_Invoce(c_Invoce.Count).������� := n_�������;
              End Loop;
              ���η�Ʊ����_Out := ���η�Ʊ����_Out + 1;
              ���δ�ӡ��Ʊ_Out := Nvl(���δ�ӡ��Ʊ_Out, '') || ',' || v_��Ʊ��;
              v_��Ʊ��         := Zl_Incstr(v_��Ʊ��);
              c_No.Delete;
            End If;
            n_Ԫ��4_Count := 0;
          End If;
          n_Ԫ��4_Count := n_Ԫ��4_Count + 1;
        
          -------------------------------------------
          --���䵥�ݺż����
          n_Find := 0;
          For J In 1 .. c_No.Count Loop
            If c_No(J).No = c_��ҳ.���� Then
              --���ݺ���ͬ,����źϲ�
              c_No(J).��� := c_No(J).��� || ',' || c_��ҳ.���;
              n_Find := 1;
              Exit;
            End If;
          End Loop;
          If n_Find = 0 Then
            c_No.Extend;
            c_No(c_No.Count).No := c_��ҳ.����;
            c_No(c_No.Count).��� := c_No(c_No.Count).��� || ',' || c_��ҳ.���;
          End If;
        End Loop;
        --���䵥��
        If c_No.Count <> 0 Then
          --��ҳ:���㷢Ʊ�ż���ص�
          For J In 1 .. c_No.Count Loop
            c_Invoce.Extend;
            c_Invoce(c_Invoce.Count).Ʊ�� := v_��Ʊ��;
            c_Invoce(c_Invoce.Count).No := c_No(J).No;
            c_Invoce(c_Invoce.Count).��� := Case
                                             When Instr(c_No(J).���, ',') > 0 Then
                                              Substr(c_No(J).���, 2)
                                             Else
                                              c_No(J).���
                                           End;
            c_Invoce(c_Invoce.Count).������� := n_�������;
          End Loop;
          ���η�Ʊ����_Out := ���η�Ʊ����_Out + 1;
          ���δ�ӡ��Ʊ_Out := Nvl(���δ�ӡ��Ʊ_Out, '') || ',' || v_��Ʊ��;
          v_��Ʊ��         := Zl_Incstr(v_��Ʊ��);
          c_No.Delete;
        End If;
      End Loop;
    End If;
  
    --7.(no Or ִ�п���)+�վݷ�Ŀ+�շ�ϸĿ
    n_������� := 0;
    If (n_�ֵ��ݴ�ӡ = 0 Or n_ִ�п��� > 0) And n_�վݷ�Ŀ <> 0 And n_�շ�ϸĿ <> 0 Then
      c_Split_��Ŀ.Delete;
    
      For I In 1 .. c_Split_No.Count Loop
      
        n_�������    := n_������� + 1;
        v_Ԫ��3       := '+';
        n_Ԫ��3_Count := 0;
        l_Ԫ��3.Delete;
        For c_��ҳ In (With c_���� As
                        (Select /*+cardinality(b,10)*/
                         Decode(n_�ֵ��ݴ�ӡ, 0, '-', a.No) As Ԫ��1, Decode(n_ִ�п���, 0, 0, a.ִ�в���id) As Ԫ��2,
                         Decode(n_�վݷ�Ŀ, 0, '-', a.�վݷ�Ŀ) As Ԫ��3, Decode(n_�շ�ϸĿ, 0, 0, a.�շ�ϸĿid) As Ԫ��4, a.No As ����,
                         Nvl(a.�۸񸸺�, a.���) As ���, Sum(Nvl(a.ʵ�ս��, 0)) As ʵ�ս��
                        From ������ü�¼ A, Table(Print_Nos) B
                        Where Mod(��¼����, 10) = 1 And a.No = b.Column_Value And Decode(n_�����, 1, Nvl(a.���ӱ�־, 0), 0) <> 9
                        Group By a.No, a.ִ�в���id, a.�վݷ�Ŀ, a.�շ�ϸĿid, Nvl(a.�۸񸸺�, a.���)
                        Having Sum(Nvl(a.ʵ�ս��, 0)) <> 0)
                       Select a.Ԫ��3, Count(*) As ����
                       From c_���� A
                       Where a.Ԫ��1 = c_Split_No(I).Ԫ��1 And
                             Instr(',' || c_Split_No(I).Ԫ��2�� || ',', ',' || a.Ԫ��2 || ',') > 0
                       Group By Ԫ��3
                       Order By Ԫ��3) Loop
          If (v_Ԫ��3 <> c_��ҳ.Ԫ��3 And n_Ԫ��3_Count >= n_�վݷ�Ŀ And n_�վݷ�Ŀ <> 0) Then
            If v_Ԫ��3 <> '+' Then
              c_Split_��Ŀ.Extend;
              c_Split_��Ŀ(c_Split_��Ŀ.Count).Ԫ��1 := c_Split_No(I).Ԫ��1;
              c_Split_��Ŀ(c_Split_��Ŀ.Count).Ԫ��2�� := c_Split_No(I).Ԫ��2��;
              c_Split_��Ŀ(c_Split_��Ŀ.Count).������� := n_�������;
              For J In 1 .. l_Ԫ��3.Count Loop
                --���ݺ���ͬ,����źϲ�
                c_Split_��Ŀ(c_Split_��Ŀ.Count).Ԫ��3�� := c_Split_��Ŀ(c_Split_��Ŀ.Count).Ԫ��3�� || ',' || l_Ԫ��3(J);
              End Loop;
            End If;
            v_Ԫ��3       := '+';
            n_Ԫ��3_Count := 0;
            l_Ԫ��3.Delete;
          End If;
          If (v_Ԫ��3 <> c_��ҳ.Ԫ��3) Then
            n_Ԫ��3_Count := n_Ԫ��3_Count + 1;
            v_Ԫ��3       := c_��ҳ.Ԫ��3;
            l_Ԫ��3.Extend;
            l_Ԫ��3(l_Ԫ��3.Count) := v_Ԫ��3;
          End If;
        End Loop;
      
        If l_Ԫ��3.Count <> 0 Then
          c_Split_��Ŀ.Extend;
          c_Split_��Ŀ(c_Split_��Ŀ.Count).Ԫ��1 := c_Split_No(I).Ԫ��1;
          c_Split_��Ŀ(c_Split_��Ŀ.Count).Ԫ��2�� := c_Split_No(I).Ԫ��2��;
          c_Split_��Ŀ(c_Split_��Ŀ.Count).������� := n_�������;
          For J In 1 .. l_Ԫ��3.Count Loop
            --���ݺ���ͬ,����źϲ�
            c_Split_��Ŀ(c_Split_��Ŀ.Count).Ԫ��3�� := c_Split_��Ŀ(c_Split_��Ŀ.Count).Ԫ��3�� || ',' || l_Ԫ��3(J);
          End Loop;
        End If;
      End Loop;
    
      For I In 1 .. c_Split_��Ŀ.Count Loop
        c_No.Delete;
        n_Ԫ��4_Count := 0;
        --�շ�ϸĿ,����������,����Ҫ��ִ�п���+�վݷ�Ŀ
        For c_��ҳ In (With c_���� As
                        (Select /*+cardinality(b,10)*/
                         Decode(n_�ֵ��ݴ�ӡ, 0, '-', a.No) As Ԫ��1, Decode(n_ִ�п���, 0, 0, a.ִ�в���id) As Ԫ��2,
                         Decode(n_�վݷ�Ŀ, 0, '-', a.�վݷ�Ŀ) As Ԫ��3, Decode(n_�շ�ϸĿ, 0, 0, a.�շ�ϸĿid) As Ԫ��4, a.No As ����,
                         Nvl(a.�۸񸸺�, a.���) As ���, Sum(Nvl(a.ʵ�ս��, 0)) As ʵ�ս��
                        From ������ü�¼ A, Table(Print_Nos) B
                        Where Mod(��¼����, 10) = 1 And a.No = b.Column_Value And Decode(n_�����, 1, Nvl(a.���ӱ�־, 0), 0) <> 9
                        Group By a.No, a.ִ�в���id, a.�վݷ�Ŀ, a.�շ�ϸĿid, Nvl(a.�۸񸸺�, a.���)
                        Having Sum(Nvl(a.ʵ�ս��, 0)) <> 0)
                       Select Ԫ��1, Ԫ��2, Ԫ��3, a.Ԫ��4, a.����, a.���, Count(*) As ����
                       From c_���� A
                       Where a.Ԫ��1 = c_Split_��Ŀ(I).Ԫ��1 And
                             Instr(',' || c_Split_��Ŀ(I).Ԫ��2�� || ',', ',' || a.Ԫ��2 || ',') > 0 And
                             Instr(',' || c_Split_��Ŀ(I).Ԫ��3�� || ',', ',' || a.Ԫ��3 || ',') > 0
                       Group By Ԫ��1, Ԫ��2, Ԫ��4, Ԫ��3, a.����, a.���
                       Order By Ԫ��1, Ԫ��2, Ԫ��4, Ԫ��3, ����, ���) Loop
          If (n_Ԫ��4_Count >= n_�շ�ϸĿ And n_�շ�ϸĿ <> 0) Then
            --���䵥��
            If c_No.Count <> 0 Then
              --��ҳ:���㷢Ʊ�ż���ص�
              For J In 1 .. c_No.Count Loop
                c_Invoce.Extend;
                c_Invoce(c_Invoce.Count).Ʊ�� := v_��Ʊ��;
                c_Invoce(c_Invoce.Count).No := c_No(J).No;
                c_Invoce(c_Invoce.Count).��� := Case
                                                 When Instr(c_No(J).���, ',') > 0 Then
                                                  Substr(c_No(J).���, 2)
                                                 Else
                                                  c_No(J).���
                                               End;
                c_Invoce(c_Invoce.Count).������� := c_Split_��Ŀ(I).�������;
              End Loop;
              ���η�Ʊ����_Out := ���η�Ʊ����_Out + 1;
              ���δ�ӡ��Ʊ_Out := Nvl(���δ�ӡ��Ʊ_Out, '') || ',' || v_��Ʊ��;
              v_��Ʊ��         := Zl_Incstr(v_��Ʊ��);
              c_No.Delete;
            End If;
            n_Ԫ��4_Count := 0;
          End If;
          n_Ԫ��4_Count := n_Ԫ��4_Count + 1;
          -------------------------------------------
          --���䵥�ݺż����
          n_Find := 0;
          For J In 1 .. c_No.Count Loop
            If c_No(J).No = c_��ҳ.���� Then
              --���ݺ���ͬ,����źϲ�
              c_No(J).��� := c_No(J).��� || ',' || c_��ҳ.���;
              n_Find := 1;
              Exit;
            End If;
          End Loop;
          If n_Find = 0 Then
            c_No.Extend;
            c_No(c_No.Count).No := c_��ҳ.����;
            c_No(c_No.Count).��� := c_No(c_No.Count).��� || ',' || c_��ҳ.���;
          End If;
        End Loop;
        --���䵥��
        If c_No.Count <> 0 Then
          --��ҳ:���㷢Ʊ�ż���ص�
          For J In 1 .. c_No.Count Loop
            c_Invoce.Extend;
            c_Invoce(c_Invoce.Count).Ʊ�� := v_��Ʊ��;
            c_Invoce(c_Invoce.Count).No := c_No(J).No;
            c_Invoce(c_Invoce.Count).��� := Case
                                             When Instr(c_No(J).���, ',') > 0 Then
                                              Substr(c_No(J).���, 2)
                                             Else
                                              c_No(J).���
                                           End;
            c_Invoce(c_Invoce.Count).������� := c_Split_��Ŀ(I).�������;
          End Loop;
          ���η�Ʊ����_Out := ���η�Ʊ����_Out + 1;
          ���δ�ӡ��Ʊ_Out := Nvl(���δ�ӡ��Ʊ_Out, '') || ',' || v_��Ʊ��;
          v_��Ʊ��         := Zl_Incstr(v_��Ʊ��);
          c_No.Delete;
        End If;
      End Loop;
    End If;
    --�������䵥�ݽ���
    --******************************************************************************************************************************
    If Instr(Nvl(���δ�ӡ��Ʊ_Out, '-'), ',') > 0 Then
      ���δ�ӡ��Ʊ_Out := Substr(���δ�ӡ��Ʊ_Out, 2);
    End If;
  End Invoice_Split_Group;
  -------------------------------------------------------------------------------------------------------------------
Begin

  --���ñ�־||NO;ִ�п���(����);�վݷ�Ŀ(��ҳ����,����);�շ�ϸĿ(����)
  v_Para := Nvl(zl_GetSysParameter('Ʊ�ݷ������', 1121), '0||0;0;0,0;0');
  If Instr(v_Para, '||') = 0 Then
    v_Para := v_Para || '||';
  End If;
  v_Temp := Substr(v_Para, 1, Instr(v_Para, '||') - 1);
  If v_Temp Is Null Then
    --������ֵ,����������,ֱ�ӷ���
    Return;
  End If;

  --0-����ʵ�ʴ�ӡ����Ʊ��;1-����Ԥ���������Ʊ��;2-.�����Զ���������Ʊ��
  n_����ģʽ := Zl_To_Number(v_Temp);
  If Nvl(n_����ģʽ, 0) = 0 Then
    --0-����ʵ�ʴ�ӡ����Ʊ��:��ԭ���Ĵ���ʽ����Ʊ��,ֱ���˳�
    Return;
  End If;
  v_Temp       := Nvl(zl_GetSysParameter('����ʹ��Ʊ��', 1121), '0');
  n_�����     := Zl_To_Number(v_Temp);
  v_��ʼ��Ʊ�� := ��ʼ��Ʊ��_In;

  If v_��ʼ��Ʊ�� Is Null Then
    --ģ�����ʱ,���Բ�������ʼ��Ʊ��
    If Nvl(����id_In, 0) <> 0 Then
      Open c_Fact;
      Fetch c_Fact
        Into r_Factrow;
    
      If c_Fact%RowCount <> 0 Then
        If Nvl(r_Factrow.��ǰ����, '-') <> '-' Then
          v_��ʼ��Ʊ�� := Zl_Incstr(r_Factrow.��ǰ����);
        Else
          v_��ʼ��Ʊ�� := r_Factrow.��ʼ����;
        End If;
      End If;
    End If;
    If v_��ʼ��Ʊ�� Is Null Then
      v_��ʼ��Ʊ�� := 'J0000001';
    End If;
  End If;

  v_��Ʊ��   := v_��ʼ��Ʊ��;
  v_��Ʊ��Ϣ := Null;

  n_�����˲���Ʊ�� := 0;
  n_��ӡ����       := Null;
  --�����ݷ���Ʊ��
  If ��������_In = 3 Or ��������_In = 4 Then
    --1-������ӡƱ��;2-����Ʊ��;3-�ش�Ʊ��;4-�˷��ջ�Ʊ�ݲ����·���Ʊ��
    If ��Ʊ��_In Is Null Then
      v_Err_Msg := 'δ����ָ���Ļ���Ʊ��,������' || Case
                     When ��������_In = 1 Then
                      '�ش�Ʊ�ݡ�'
                     Else
                      '����Ʊ�ݡ�'
                   End;
      Raise Err_Item;
    End If;
  
    Select ���ݺ� Bulk Collect
    Into l_Print_Nos
    From (Select /*+cardinality(j,10)*/
           Distinct c.No As ���ݺ�
           From Ʊ�ݴ�ӡ��ϸ A, Ʊ��ʹ����ϸ B, Ʊ�ݴ�ӡ���� C, Table(f_Str2list(��Ʊ��_In)) J
           Where a.ʹ��id = b.Id And b.��ӡid = c.Id And a.Ʊ�� = j.Column_Value
           Order By ���ݺ�);
  
    If l_Print_Nos.Count = 0 Then
      v_Err_Msg := 'δ�ҵ�ָ����Ʊ(' || ��Ʊ��_In || '����Ӧ���շѵ���!';
      Raise Err_Item;
    End If;
  
    Select /*+cardinality(b,10)*/
     Max(��ӡ����)
    Into n_��ӡ����
    From Ʊ�ݴ�ӡ���� A, Table(l_Print_Nos) B
    Where a.No = b.Column_Value And a.�������� = 1;
  
    If Nvl(n_��ӡ����, 0) = 1 Then
      --һ�δ�ӡ�ж�ν���ģ����ʾ��ǰΪ�����˴�ӡ��
      n_�����˲���Ʊ�� := 1;
      n_��ӡ����       := 1;
    End If;
  
  Elsif ��ӡid_In <> 0 Then
    n_�����˲���Ʊ�� := 1;
    n_��ӡ����       := 1;
    Select ���ݺ� Bulk Collect
    Into l_Print_Nos
    From (Select Distinct NO As ���ݺ�
           From ��ʱƱ�ݴ�ӡ���� A
           Where a.Id = ��ӡid_In And Nvl(a.����, 0) = 1
           Order By ���ݺ�);
    If l_Print_Nos.Count = 0 Then
      v_Err_Msg := 'δ�ҵ�������Ҫ����Ʊ�ݵĵ�����Ϣ(��ӡID=' || ��ӡid_In || ')!';
      Raise Err_Item;
    End If;
  
  Else
    Select Column_Value Bulk Collect Into l_Print_Nos From Table(f_Str2list(Nos_In)) J;
    If l_Print_Nos.Count = 0 Then
      v_Err_Msg := 'δ�ҵ�������Ҫ����Ʊ�ݵĵ�����Ϣ(������Ϣ��' || Nvl(Nos_In, '') || ')!';
      Raise Err_Item;
    End If;
  End If;

  v_Nos := Null;
  If n_����ģʽ = 2 Then
    If l_Print_Nos.Count < 3000 Then
      --1.ֻ���Զ���ģʽʱ���Ż��漰���ܴ����û������������������жϣ���Ҫ��Ϊ��Ǹ��
      --2.��ǰ�����ܳ���3000�ŵ��ݣ��������3000�ŵ��ݣ���Ҫ�����Ӧ��Ʊ��,��Ҫ���ð����˲���Ʊ�ݵ����
      For I In 1 .. l_Print_Nos.Count Loop
        v_Nos := Nvl(v_Nos, '') || ',' || l_Print_Nos(I);
      
      End Loop;
      v_Nos := Substr(v_Nos, 2);
    End If;
  
    --�����Զ���������Ʊ��,����:Zl_Custom_Invoice_Autoallot����
    Zl_Custom_Invoice_Autoallot(��������_In, ģ�����_In, Ʊ��_In, ����id_In, ����id_In, v_Nos, ��ʼ��Ʊ��_In, ʹ����_In, ʹ��ʱ��_In, ��Ʊ��_In,
                                ��Ʊ����_In, n_�����˲���Ʊ��, ��ӡid_In, l_Print_Nos);
    Return;
  End If;

  --������ȡ:
  --1.����Ԥ���������Ʊ��
  --   NO;ִ�п���(����);�վݷ�Ŀ(��ҳ����,����);�շ�ϸĿ(����)
  v_Para := Substr(v_Para, Instr(v_Para, '||') + 2);
  If Instr(v_Para, ';') > 0 Then
    --NO:Ʊ���Ƿ񰴵��ݽ��зֱ��ӡ,1��ʾ�����ݴ�ӡ;0-�������ݴ�ӡ
    v_Temp       := Substr(v_Para, 1, Instr(v_Para, ';') - 1);
    n_�ֵ��ݴ�ӡ := Zl_To_Number(v_Temp);
    v_Para       := Substr(v_Para, Instr(v_Para, ';') + 1);
  End If;

  If Instr(v_Para, ';') > 0 Then
    --ִ�п���
    v_Temp     := Nvl(Substr(v_Para, 1, Instr(v_Para, ';') - 1), '0');
    n_ִ�п��� := Zl_To_Number(v_Temp);
    v_Para     := Substr(v_Para, Instr(v_Para, ';') + 1);
  End If;

  If Instr(v_Para, ';') > 0 Then
    --�վݷ�Ŀ
    v_Temp     := Nvl(Substr(v_Para, 1, Instr(v_Para, ';') - 1), '0');
    n_�վݷ�Ŀ := Zl_To_Number(v_Temp);
    v_Para     := Substr(v_Para, Instr(v_Para, ';') + 1);
  End If;

  If Instr(v_Para, ';') > 0 Then
    --�վݷ�Ŀ
    v_Temp     := Nvl(Substr(v_Para, 1, Instr(v_Para, ';') - 1), '0');
    n_�շ�ϸĿ := Zl_To_Number(v_Temp);
    v_Para     := Substr(v_Para, Instr(v_Para, ';') + 1);
  End If;

  If Instr(v_Para, ';') > 0 Then
    --ִ�п���
    v_Temp := Nvl(Substr(v_Para, 1, Instr(v_Para, ';') - 1), '0');
  Else
    v_Temp := Nvl(v_Para, '0');
  End If;
  n_�������� := Zl_To_Number(v_Temp);

  If n_�����˲���Ʊ�� = 1 Then
    --�����ӡID<>0�����,��������㣬��ʾ�����˲���Ʊ����Ʊ�ݽ��Զ����ֵ��ݴ�ӡ����ִ�п��Ҵ�ӡ���վ�ϸĿ��ӡ
    n_�ֵ��ݴ�ӡ := 0;
    n_ִ�п���   := 0;
    n_�շ�ϸĿ   := 0;
  End If;

  v_����Ʊ�ݺ� := ��Ʊ��_In;
  ��Ʊ����_In  := 0;
  --һ����ҳ���ܻ򲻻���
  If n_�������� <> 2 Then
    Invoice_Split_Notgroup(l_Print_Nos, ��Ʊ��_In, v_��Ʊ��Ϣ, ��Ʊ����_In);
  Else
    --�����������
    Invoice_Split_Group(l_Print_Nos, ��Ʊ��_In, v_��Ʊ��Ϣ, ��Ʊ����_In);
  End If;
  ��Ʊ��_In := v_��Ʊ��Ϣ;
  If ģ�����_In = 1 Then
    --ģ�����,ֻ����Ʊ��������ʹ�õ�Ʊ�ݺ�,ֱ���˳�
    Return;
  End If;

  v_��ʼ��Ʊ�� := Null;
  v_��ǰ��Ʊ�� := Null;
  --1-������ӡƱ��;2-����Ʊ��;3-�ش�Ʊ��;4-�˷��ջ�Ʊ�ݲ����·���Ʊ��
  If ��������_In = 3 Or ��������_In = 4 Then
    --�ջ�Ʊ��
    Select ʹ��id Bulk Collect
    Into l_ʹ��id
    From (Select /*+cardinality(j,10)*/
           Distinct b.ʹ��id
           From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ��ϸ B, Table(f_Str2list(v_����Ʊ�ݺ�)) J
           Where a.Id = b.ʹ��id And b.Ʊ�� = j.Column_Value And Nvl(b.Ʊ��, 0) = 1);
  
    --������ռ�¼
    Forall I In 1 .. l_ʹ��id.Count
      Insert Into Ʊ��ʹ����ϸ
        (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ����, ʹ��ʱ��)
        Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, Decode(��������_In, 3, 4, 2), ����id, ��ӡid, ʹ����_In, ʹ��ʱ��_In
        From Ʊ��ʹ����ϸ
        Where ID = l_ʹ��id(I);
    Forall I In 1 .. l_ʹ��id.Count
      Update Ʊ�ݴ�ӡ��ϸ Set �Ƿ���� = 1 Where ʹ��id = l_ʹ��id(I);
  End If;

  If c_Invoce.Count = 0 Then
    --������,ֱ�ӷ���
    Return;
  End If;

  If ��ʼ��Ʊ��_In Is Null Then
    v_Err_Msg := 'δ������ʼ��Ʊ��,���ܽ���Ʊ�ݷ��䴦��';
    Raise Err_Item;
  End If;

  If Nvl(����id_In, 0) <> 0 Then
    Open c_Fact;
    Fetch c_Fact
      Into r_Factrow;
    If c_Fact%RowCount = 0 Then
      v_Err_Msg := '��Ч��Ʊ���������Σ��޷�����շ�Ʊ�ݷ��������';
      Close c_Fact;
      Raise Err_Item;
    Elsif Nvl(r_Factrow.ʣ������, 0) < ��Ʊ����_In Then
      v_Err_Msg := '��ǰ���ε�ʣ����������' || ��Ʊ����_In || '�ţ��޷�����շ�Ʊ�ݷ��������';
      Close c_Fact;
      Raise Err_Item;
    End If;
  End If;

  --ʵ�ʴ���Ʊ����Ϣ
  If Nvl(n_�ֵ��ݴ�ӡ, 0) <> 1 Or Nvl(n_�����˲���Ʊ��, 0) = 1 Then
    --���ֵ��ݴ�ӡʱ,��ʾһ�δ�ӡ,��ӡID���һ��
    n_��ӡid := ��ӡid_In;
    If Nvl(n_��ӡid, 0) = 0 Then
      Select Ʊ�ݴ�ӡ����_Id.Nextval Into n_��ӡid From Dual;
    End If;
  End If;

  ��Ʊ����_In := 0;
  v_��ӡ����  := '';
  For c_Invoce_No In (Select Column_Value As ��Ʊ�� From Table(f_Str2list(v_��Ʊ��Ϣ)) Order By ��Ʊ��) Loop
    --���Ʊ�ݷ�Χ�Ƿ���ȷ
    If Nvl(����id_In, 0) <> 0 Then
      If Not (Upper(c_Invoce_No.��Ʊ��) >= Upper(r_Factrow.��ʼ����) And Upper(c_Invoce_No.��Ʊ��) <= Upper(r_Factrow.��ֹ����) And
          Length(c_Invoce_No.��Ʊ��) = Length(r_Factrow.��ֹ����)) Then
        v_Err_Msg := '�õ�����Ҫ��ӡ����Ʊ��,��Ʊ�ݺ�"' || c_Invoce_No.��Ʊ�� || '"����Ʊ�����õĺ��뷶Χ��';
        Close c_Fact;
        Raise Err_Item;
      End If;
    End If;
  
    --����Ʊ�ݴ�ӡ��ϸ
    r_���ݺ�.Delete;
    r_�������.Delete;
    l_�������.Delete;
  
    Select Ʊ��ʹ����ϸ_Id.Nextval Into n_ʹ��id From Dual;
  
    n_������� := 0;
  
    For I In 1 .. c_Invoce.Count Loop
      If c_Invoce(I).Ʊ�� = c_Invoce_No.��Ʊ�� Then
        n_������� := c_Invoce(I).�������;
        Exit;
      End If;
    End Loop;
    --�������Ʊ��,�Ա����Ʊ��
    For I In 1 .. c_Invoce.Count Loop
      If c_Invoce(I).������� = n_������� And Nvl(c_Invoce(I).�޸ı�־, 0) = 0 Then
        If n_������� <> 0 Then
          c_Invoce(I).������� := n_ʹ��id;
        End If;
        c_Invoce(I).�޸ı�־ := 1;
      End If;
    End Loop;
  
    For I In 1 .. c_Invoce.Count Loop
      If c_Invoce(I).Ʊ�� = c_Invoce_No.��Ʊ�� Then
        r_���ݺ�.Extend;
        r_���ݺ�(r_���ݺ�.Count) := c_Invoce(I).No;
        r_�������.Extend;
        r_�������(r_�������.Count) := c_Invoce(I).���;
        l_�������.Extend;
        If Nvl(c_Invoce(I).�������, 0) <> 0 Then
          --����Ƿ����������Ʊ��
          n_Find := 0;
          For J In 1 .. c_Invoce.Count Loop
            If c_Invoce(I).������� = c_Invoce(J).������� And c_Invoce(I).Ʊ�� <> c_Invoce(J).Ʊ�� Then
              n_Find := 1;
              Exit;
            End If;
          End Loop;
        
          If n_Find = 0 Then
            l_�������(l_�������.Count) := Null;
            c_Invoce(I).������� := 0;
          Else
            l_�������(l_�������.Count) := c_Invoce(I).�������;
          End If;
        Else
          l_�������(l_�������.Count) := Null;
        End If;
      End If;
    End Loop;
  
    --1.�����Ŵ�ӡ����
    If n_�ֵ��ݴ�ӡ = 1 Then
      --�ֵ��ݴ�ӡ,�谴���ݽ��д���
      --Ʊ�ݴ�ӡ����
      n_Find := 0;
      v_Temp := '';
      For I In 1 .. r_���ݺ�.Count Loop
        v_Temp := v_Temp || ',' || r_���ݺ�(I);
        If Instr(Nvl(v_��ӡ����, '-') || ',', ',' || r_���ݺ�(I) || ',') > 0 Then
          --�Ѿ��ҵ�
          n_Find := 1;
        End If;
      End Loop;
      v_��ӡ���� := v_��ӡ���� || Nvl(v_Temp, '+');
    
      If Nvl(n_Find, 0) = 0 Then
        Select Ʊ�ݴ�ӡ����_Id.Nextval Into n_��ӡid From Dual;
        Forall I In 1 .. r_���ݺ�.Count
          Insert Into Ʊ�ݴ�ӡ���� (ID, ��������, NO, ��ӡ����) Values (n_��ӡid, 1, r_���ݺ�(I), n_��ӡ����);
        --�Ա����������ü�¼�е�ʵ��Ʊ��
        v_��ʼ��Ʊ�� := c_Invoce_No.��Ʊ��;
        Forall I In 1 .. r_���ݺ�.Count
          Update ������ü�¼ Set ʵ��Ʊ�� = v_��ʼ��Ʊ�� Where Mod(��¼����, 10) = 1 And NO = r_���ݺ�(I);
      End If;
    Else
    
      If v_��ʼ��Ʊ�� Is Null Then
        --�Ա����������ü�¼�е�ʵ��Ʊ��
        v_��ʼ��Ʊ�� := c_Invoce_No.��Ʊ��;
      
        --Ʊ�ݴ�ӡ����
        Insert Into Ʊ�ݴ�ӡ����
          (ID, ��������, NO, ��ӡ����)
          Select n_��ӡid, 1, Column_Value, n_��ӡ���� From Table(l_Print_Nos);
      
        Update ������ü�¼
        Set ʵ��Ʊ�� = v_��ʼ��Ʊ��
        Where Mod(��¼����, 10) = 1 And NO In (Select Column_Value From Table(l_Print_Nos));
      End If;
    End If;
  
    --2.����Ʊ�ݴ�ӡ��ϸ
  
    ��Ʊ����_In := ��Ʊ����_In + 1;
    --����Ʊ��ʹ����ϸ
  
    Insert Into Ʊ��ʹ����ϸ
      (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ����, ʹ��ʱ��)
    Values
      (n_ʹ��id, 1, c_Invoce_No.��Ʊ��, 1, Decode(��������_In, 3, 3, 1), Decode(Nvl(����id_In, 0), 0, Null, ����id_In), n_��ӡid,
       ʹ����_In, ʹ��ʱ��_In);
  
    Forall I In 1 .. r_���ݺ�.Count
      Insert Into Ʊ�ݴ�ӡ��ϸ
        (ʹ��id, Ʊ��, �Ƿ����, NO, Ʊ��, ���, ����Ʊ�����)
      Values
        (n_ʹ��id, 1, 0, r_���ݺ�(I), c_Invoce_No.��Ʊ��, r_�������(I), l_�������(I));
  
    v_��ǰ��Ʊ�� := c_Invoce_No.��Ʊ��;
  End Loop;

  If Nvl(����id_In, 0) <> 0 Then
    Close c_Fact;
  
    Update Ʊ�����ü�¼
    Set ʹ��ʱ�� = ʹ��ʱ��_In, ��ǰ���� = v_��ǰ��Ʊ��, ʣ������ = Nvl(ʣ������, 0) - ��Ʊ����_In
    Where ID = ����id_In
    Returning ʣ������ Into n_������;
    If n_������ < 0 Then
      v_Err_Msg := '��ǰ���ε�ʣ����������' || ��Ʊ����_In || '�ţ��޷�����շ�Ʊ�ݷ��������';
      Raise Err_Item;
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Invoice_Autoallot;
/

--108706:������,2017-05-08,����תסԺ����������ԭ��סԺԤ����
Create Or Replace Procedure Zl_����תסԺ_����������
(
  No_In         סԺ���ü�¼.No%Type,
  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
  �˷�ʱ��_In   סԺ���ü�¼.����ʱ��%Type,
  �����˷�_In   Number := 0,
  ��Ժ����id_In סԺ���ü�¼.��������id%Type := Null,
  ��ҳid_In     סԺ���ü�¼.��ҳid%Type := Null,
  �����˷�_In   Number := 0,
  ����id_In     ����Ԥ����¼.����id%Type := Null
) As
  v_����ids    Varchar2(3000);
  n_��id       ����ɿ����.Id%Type;
  n_����       Number;
  v_Ԥ��no     ����Ԥ����¼.No%Type;
  n_����id     ������ü�¼.����id%Type;
  v_Nos        Varchar2(3000);
  v_Info       Varchar2(5000);
  v_��ǰ����   Varchar2(3000);
  v_ԭ����ids  Varchar2(5000);
  n_Tempid     ����Ԥ����¼.Id%Type;
  v_��ˮ��     ����Ԥ����¼.������ˮ��%Type;
  v_˵��       ����Ԥ����¼.����˵��%Type;
  n_Ԥ��id     ����Ԥ����¼.Id%Type;
  n_ԭԤ��id   ����Ԥ����¼.Id%Type;
  n_����id     ������Ϣ.����id%Type;
  n_ԭ����id   ����Ԥ����¼.����id%Type;
  n_�������   ����Ԥ����¼.��Ԥ��%Type;
  n_�����     ����Ԥ����¼.�����id%Type;
  n_������     Number;
  n_����ֵ     ��Ա�ɿ����.���%Type;
  v_���㷽ʽ   ����Ԥ����¼.���㷽ʽ%Type;
  n_�����id   ҽ�ƿ����.Id%Type;
  v_����       ����Ԥ����¼.����%Type;
  v_������ˮ�� ����Ԥ����¼.������ˮ��%Type;
  v_����˵��   ����Ԥ����¼.����˵��%Type;
  n_ԭ����     Number;
  Err_Item Exception;
  v_Err_Msg Varchar2(200);
  Procedure Zl_Square_Update
  (
    ����ids_In    Varchar2,
    �ֽ���id_In   ����Ԥ����¼.����id%Type,
    �ɿ���id_In   ����Ԥ����¼.�ɿ���id%Type,
    �˿�ʱ��_In   ����Ԥ����¼.�տ�ʱ��%Type,
    �������_In   ����Ԥ����¼.�������%Type,
    ��������_In   Varchar2 := Null,
    �˷ѽ��_In   ����Ԥ����¼.��Ԥ��%Type := Null,
    ���㿨���_In ����Ԥ����¼.���㿨���%Type := Null
  ) As
    n_��¼״̬ ���˿������¼.��¼״̬%Type;
    n_Ԥ��id   ����Ԥ����¼.Id%Type;
    v_����     ���˿������¼.����%Type;
    n_���ڿ�Ƭ Number;
    d_ͣ������ ���ѿ�Ŀ¼.ͣ������%Type;
    n_������ ���˿������¼.���%Type;
    n_���     ���˿������¼.���%Type;
    n_���     ���ѿ�Ŀ¼.���%Type;
    n_�ӿڱ�� ���˿������¼.�ӿڱ��%Type;
    d_����ʱ�� ���ѿ�Ŀ¼.����ʱ��%Type;
    n_Id       ����Ԥ����¼.Id%Type;
  Begin
    n_Ԥ��id := 0;
  
    --�������ѿ�,���㿨��������Ѿ�������
    For v_У�� In (Select Min(a.Id) As Ԥ��id, c.���ѿ�id, Sum(c.������) As ������, c.�ӿڱ��, c.����, Min(c.���) As ���, Min(c.Id) As ID
                 From ����Ԥ����¼ A, ���˿�������� B, ���˿������¼ C
                 Where a.Id = b.Ԥ��id And a.���㿨��� = ���㿨���_In And b.������id = c.Id And a.��¼���� = 3 And
                       Instr(Nvl(��������_In, '_LXH'), ',' || a.���㷽ʽ || ',') = 0 And
                       a.����id In (Select Column_Value From Table(f_Str2list(����ids_In)))
                 Group By c.���ѿ�id, c.�ӿڱ��, c.����) Loop
    
      If Nvl(v_У��.���ѿ�id, 0) <> 0 Then
        Select Max(��¼״̬)
        Into n_��¼״̬
        From ���˿������¼
        Where �ӿڱ�� = v_У��.�ӿڱ�� And ���ѿ�id = Nvl(v_У��.���ѿ�id, 0) And ���� = v_У��.���� And Nvl(���, 0) = Nvl(v_У��.���, 0);
      Else
        Select Max(��¼״̬)
        Into n_��¼״̬
        From ���˿������¼
        Where �ӿڱ�� = v_У��.�ӿڱ�� And ���ѿ�id Is Null And ���� = v_У��.���� And Nvl(���, 0) = Nvl(v_У��.���, 0);
      End If;
    
      If n_��¼״̬ = 1 Then
        n_��¼״̬ := 2;
      Else
        n_��¼״̬ := n_��¼״̬ + 2;
      End If;
      --����ʱ,ֻ����һ��
      If n_Ԥ��id = 0 Then
        Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id,
           Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, У�Ա�־, �������, ��������)
          Select n_Ԥ��id, ��¼����, NO, 2, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �˿�ʱ��_In, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���_In, ����Ա����_In,
                 -1 * �˷ѽ��_In, �ֽ���id_In, �ɿ���id_In, Ԥ�����, �����id, Nvl(���㿨���, v_У��.�ӿڱ��), ����, ������ˮ��, ����˵��, ������λ, 2, �������_In,
                 ��������
          From ����Ԥ����¼ A
          Where ID = v_У��.Ԥ��id;
      End If;
    
      If Nvl(v_У��.���ѿ�id, 0) <> 0 Then
        --���ѿ�,ֱ���˻ؿ�������
        Begin
          Select ����, 1, ͣ������, (Select Max(���) From ���ѿ�Ŀ¼ B Where a.���� = b.���� And a.�ӿڱ�� = b.�ӿڱ��), ���, ���, �ӿڱ��, ����ʱ��
          Into v_����, n_���ڿ�Ƭ, d_ͣ������, n_������, n_���, n_���, n_�ӿڱ��, d_����ʱ��
          From ���ѿ�Ŀ¼ A
          Where ID = v_У��.���ѿ�id;
        Exception
          When Others Then
            n_���ڿ�Ƭ := 0;
        End;
      
        --ȡ��ͣ��
        If n_���ڿ�Ƭ = 0 Then
          v_Err_Msg := '����Ϊ"' || v_���� || '"�����ѿ�������ɾ�������������øÿ�Ƭ,���飡';
          Raise Err_Item;
        End If;
        If Nvl(n_���, 0) < Nvl(n_������, 0) Then
          v_Err_Msg := '����������ʷ������¼(����Ϊ"' || v_���� || '"),���飡';
          Raise Err_Item;
        End If;
        If Nvl(d_ͣ������, To_Date('3000-01-01', 'yyyy-mm-dd')) < To_Date('3000-01-01', 'yyyy-mm-dd') Then
          v_Err_Msg := '����Ϊ"' || v_���� || '"�����ѿ��Ѿ�������ͣ�ã������ٽ����˷�,���飡';
          Raise Err_Item;
        End If;
      
        If d_����ʱ�� < To_Date('3000-01-01', 'yyyy-mm-dd') Then
          v_Err_Msg := '����Ϊ"' || v_���� || '"�����ѿ��Ѿ����գ������˷�,���飡';
          Raise Err_Item;
        End If;
        Update ���ѿ�Ŀ¼ Set ��� = Nvl(���, 0) + �˷ѽ��_In Where ID = Nvl(v_У��.���ѿ�id, 0);
      End If;
    
      Select ���˿������¼_Id.Nextval Into n_Id From Dual;
      Insert Into ���˿������¼
        (ID, �ӿڱ��, ���ѿ�id, ���, ��¼״̬, ���㷽ʽ, ������, ����, ������ˮ��, ����ʱ��, ��ע, �����־)
        Select n_Id, �ӿڱ��, ���ѿ�id, ���, n_��¼״̬, ���㷽ʽ, -1 * �˷ѽ��_In, ����, ������ˮ��, ����ʱ��, ��ע,
               Decode(���ѿ�id, Null, 0, 0, 0, 1) As ��־
        From ���˿������¼
        Where ID = v_У��.Id;
      Insert Into ���˿�������� (Ԥ��id, ������id) Values (n_Ԥ��id, n_Id);
    
      If n_��¼״̬ <> 2 And n_��¼״̬ <> 1 Then
        Update ���˿������¼ Set ��¼״̬ = 3 Where ID = v_У��.Id;
      End If;
    End Loop;
  End;
Begin
  n_��id := Zl_Get��id(����Ա����_In);
  If ����id_In Is Null Then
    Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
  Else
    n_����id := ����id_In;
  End If;

  Select ����id, ����id
  Into n_ԭ����id, n_����id
  From ������ü�¼
  Where NO = No_In And ��¼���� = 1 And ��¼״̬ In (1, 3) And Rownum < 2;

  For r_����id In (Select Distinct ����id
                 From ������ü�¼
                 Where NO In (Select Distinct NO
                              From ������ü�¼
                              Where ����id In (Select ����id
                                             From ����Ԥ����¼
                                             Where ������� In (Select b.�������
                                                            From ������ü�¼ A, ����Ԥ����¼ B
                                                            Where a.No = No_In And b.������� < 0 And Mod(a.��¼����, 10) = 1 And
                                                                  a.��¼״̬ <> 0 And a.����id = b.����id))) And
                       Mod(��¼����, 10) = 1 And ��¼״̬ <> 0
                 Union
                 Select Distinct ����id
                 From ������ü�¼
                 Where NO In (Select Distinct NO
                              From ������ü�¼
                              Where ����id In (Select a.����id
                                             From ������ü�¼ A, ����Ԥ����¼ B
                                             Where a.No = No_In And b.������� > 0 And Mod(a.��¼����, 10) = 1 And a.��¼״̬ <> 0 And
                                                   a.����id = b.����id))) Loop
    v_ԭ����ids := v_ԭ����ids || ',' || r_����id.����id;
  End Loop;
  v_ԭ����ids := Substr(v_ԭ����ids, 2);

  Begin
    Select ժҪ
    Into v_Info
    From ����Ԥ����¼
    Where ���㷽ʽ Is Null And ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id;
  Exception
    When Others Then
      v_Info := '';
  End;
  --����������Ϣ
  If v_Info Is Not Null Then
    While v_Info Is Not Null Loop
      v_��ǰ���� := Substr(v_Info, 1, Instr(v_Info, '|') - 1);
      n_������   := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1));
    
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
      n_�����   := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1));
    
      v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
      n_������� := -1 * To_Number(v_��ǰ����);
    
      If n_������ = 0 Then
        --���ѿ�
        Select ���㷽ʽ
        Into v_���㷽ʽ
        From ����Ԥ����¼
        Where ����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And ���㿨��� = n_����� And Rownum < 2;
        Zl_Square_Update(v_ԭ����ids, n_����id, n_��id, �˷�ʱ��_In, -1 * n_����id, Null, n_�������, n_�����);
        Update ��Ա�ɿ����
        Set ��� = Nvl(���, 0) - n_�������
        Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_���㷽ʽ
        Returning ��� Into n_����ֵ;
        If Sql%RowCount = 0 Then
          Insert Into ��Ա�ɿ����
            (�տ�Ա, ���㷽ʽ, ����, ���)
          Values
            (����Ա����_In, v_���㷽ʽ, 1, -1 * n_�������);
          n_����ֵ := n_�������;
        End If;
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete From ��Ա�ɿ����
          Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_���㷽ʽ And Nvl(���, 0) = 0;
        End If;
      Else
        --���㿨
        Select ���㷽ʽ, �����id, ����, ������ˮ��, ����˵��
        Into v_���㷽ʽ, n_�����id, v_����, v_������ˮ��, v_����˵��
        From ����Ԥ����¼
        Where ����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And �����id = n_����� And Rownum < 2;
      
        If Nvl(�����˷�_In, 0) = 1 Then
          If �����˷�_In = 0 Then
            v_Err_Msg := '�����޷����ֵ������˻�,�޷������˷�!';
            Raise Err_Item;
          End If;
          Update ����Ԥ����¼
          Set ��Ԥ�� = ��Ԥ�� - n_�������
          Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_���㷽ʽ;
          If Sql%RowCount = 0 Then
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ,
               �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, У�Ա�־, ��������)
            Values
              (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * n_�������, v_���㷽ʽ, Null, �˷�ʱ��_In,
               Null, Null, Null, ����Ա���_In, ����Ա����_In, '', n_��id, n_�����id, Null, v_����, v_������ˮ��, v_����˵��, Null, n_����id,
               -1 * n_����id, 0, 3);
          End If;
          Update ��Ա�ɿ����
          Set ��� = Nvl(���, 0) - n_�������
          Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_���㷽ʽ
          Returning ��� Into n_����ֵ;
          If Sql%RowCount = 0 Then
            Insert Into ��Ա�ɿ����
              (�տ�Ա, ���㷽ʽ, ����, ���)
            Values
              (����Ա����_In, v_���㷽ʽ, 1, -1 * n_�������);
            n_����ֵ := -1 * n_�������;
          End If;
          If Nvl(n_����ֵ, 0) = 0 Then
            Delete From ��Ա�ɿ����
            Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_���㷽ʽ And Nvl(���, 0) = 0;
          End If;
        Else
          Begin
            Select 1 Into n_���� From ҽ�ƿ���� Where ID = n_�����id And �Ƿ����� = 1;
          Exception
            When Others Then
              n_���� := 0;
          End;
        
          If �����˷�_In = 1 Or n_���� = 0 Then
            v_���㷽ʽ := v_���㷽ʽ;
            n_ԭ����   := 1;
          Else
            n_ԭ���� := 0;
            Begin
              Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And ���� Like '%�ֽ�%' And Rownum < 2;
            Exception
              When Others Then
                Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And Rownum < 2;
            End;
          End If;
        
          If �����˷�_In = 0 Then
            If n_ԭ���� = 1 Then
              Select ������ˮ��, ����˵��, ID
              Into v_��ˮ��, v_˵��, n_ԭԤ��id
              From ����Ԥ����¼
              Where ����id = n_ԭ����id And ���㷽ʽ = v_���㷽ʽ And Rownum < 2;
            
              Update ����Ԥ����¼
              Set ��Ԥ�� = ��Ԥ�� - n_�������
              Where ��¼���� = 3 And ��¼״̬ = 2 And �տ�ʱ�� = �˷�ʱ��_In And ����id + 0 = n_����id And ���㷽ʽ = v_���㷽ʽ And ����id = n_����id;
              If Sql%RowCount = 0 Then
                Insert Into ����Ԥ����¼
                  (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                   ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, У�Ա�־, ��������)
                Values
                  (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * n_�������, v_���㷽ʽ, Null, �˷�ʱ��_In,
                   Null, Null, Null, ����Ա���_In, ����Ա����_In, '', n_��id, n_�����id, Null, v_����, v_������ˮ��, v_����˵��, Null, n_����id,
                   -1 * n_����id, 0, 3);
              End If;
            
              v_Ԥ��no := Nextno(11);
              Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
              Insert Into ����Ԥ����¼
                (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ,
                 �ɿ���id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ)
              Values
                (n_Ԥ��id, v_Ԥ��no, Null, 1, 1, n_����id, ��ҳid_In, ��Ժ����id_In, n_�������, v_���㷽ʽ, Null, �˷�ʱ��_In, Null, Null,
                 Null, ����Ա���_In, ����Ա����_In, '����תסԺԤ��', n_��id, 2, n_�����id, Null, v_����, v_��ˮ��, v_˵��, Null);
              Update �������㽻�� Set ����id = n_Ԥ��id Where ����id = n_ԭԤ��id;
            Else
              Update ����Ԥ����¼
              Set ��Ԥ�� = ��Ԥ�� - n_�������
              Where ��¼���� = 3 And ��¼״̬ = 2 And �տ�ʱ�� = �˷�ʱ��_In And ����id + 0 = n_����id And ���㷽ʽ = v_���㷽ʽ And ����id = n_����id;
              If Sql%RowCount = 0 Then
                Insert Into ����Ԥ����¼
                  (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                   ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, У�Ա�־, ��������)
                Values
                  (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * n_�������, v_���㷽ʽ, Null, �˷�ʱ��_In,
                   Null, Null, Null, ����Ա���_In, ����Ա����_In, '', n_��id, Null, Null, Null, Null, Null, Null, n_����id,
                   -1 * n_����id, 0, 3);
              End If;
            
              Update ����Ԥ����¼
              Set ��� = ��� + n_�������
              Where ��¼���� = 1 And ��¼״̬ = 1 And �տ�ʱ�� = �˷�ʱ��_In And ����id + 0 = n_����id And ���㷽ʽ = v_���㷽ʽ;
              If Sql%RowCount = 0 Then
                v_Ԥ��no := Nextno(11);
                Insert Into ����Ԥ����¼
                  (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                   ժҪ, �ɿ���id, Ԥ�����)
                Values
                  (����Ԥ����¼_Id.Nextval, v_Ԥ��no, Null, 1, 1, n_����id, ��ҳid_In, ��Ժ����id_In, n_�������, v_���㷽ʽ, Null, �˷�ʱ��_In,
                   Null, Null, Null, ����Ա���_In, ����Ա����_In, '����תסԺԤ��', n_��id, 2);
              End If;
            End If;
          
            --�������
            Update �������
            Set Ԥ����� = Nvl(Ԥ�����, 0) + n_�������
            Where ���� = 1 And ����id = n_����id And ���� = 2
            Returning Ԥ����� Into n_����ֵ;
            If Sql%RowCount = 0 Then
              Insert Into ������� (����id, ����, ����, Ԥ�����, �������) Values (n_����id, 1, 2, n_�������, 0);
              n_����ֵ := n_�������;
            End If;
            If Nvl(n_����ֵ, 0) = 0 Then
              Delete From �������
              Where ����id = n_����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
            End If;
          End If;
          --4.2�ɿ����ݴ���
          --   ��Ϊû��ʵ���ղ��˵�Ǯ,���Բ�����
          --�����˷��������ԭԤ����¼
          If �����˷�_In = 1 Then
            Update ��Ա�ɿ����
            Set ��� = Nvl(���, 0) - n_�������
            Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_���㷽ʽ
            Returning ��� Into n_����ֵ;
            If Sql%RowCount = 0 Then
              Insert Into ��Ա�ɿ����
                (�տ�Ա, ���㷽ʽ, ����, ���)
              Values
                (����Ա����_In, v_���㷽ʽ, 1, -1 * n_�������);
              n_����ֵ := -1 * n_�������;
            End If;
            If Nvl(n_����ֵ, 0) = 0 Then
              Delete From ��Ա�ɿ����
              Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_���㷽ʽ And Nvl(���, 0) = 0;
            End If;
            Update ����Ԥ����¼
            Set ��Ԥ�� = ��Ԥ�� + (-1 * n_�������)
            Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_���㷽ʽ;
            If Sql%RowCount = 0 Then
              Insert Into ����Ԥ����¼
                (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ,
                 �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, У�Ա�־, ��������)
              Values
                (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * n_�������, v_���㷽ʽ, Null, �˷�ʱ��_In,
                 Null, Null, Null, ����Ա���_In, ����Ա����_In, '', n_��id, n_�����id, Null, v_����, v_������ˮ��, v_����˵��, Null, n_����id,
                 -1 * n_����id, 0, 3);
            End If;
          End If;
        End If;
      End If;
      v_Info := Substr(v_Info, Instr(v_Info, '|') + 1);
    End Loop;
  End If;

  Delete From ����Ԥ����¼ Where ����id = n_����id And ��¼״̬ = 2 And ���㷽ʽ Is Null;
  Update ������ü�¼ Set ����״̬ = 0 Where ����id = n_����id;
  Update ������ü�¼ Set ����״̬ = 0 Where NO = No_In And Mod(��¼����, 10) = 1 And ��¼״̬ = 2;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����תסԺ_����������;
/

--108432:Ƚ����,2017-05-08,�����̶��������ʱ����ȡ����˺�û��ɾ���ɸ���ʱ�������ɵĳ����¼�������������ʱ����ʱ���������
Create Or Replace Procedure Zl_�ٴ����ﰲ��_Publish
(
  Id_In       �ٴ������.Id%Type,
  ������_In   �ٴ������.������%Type := Null,
  ����ʱ��_In �ٴ������.����ʱ��%Type := Null,
  ȡ������_In Number := 0
) As
  --������ȡ����������
  --������
  --        ȡ������_In �Ƿ�ȡ������
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_Count    Number(2);
  n_�Ű෽ʽ �ٴ������.�Ű෽ʽ%Type;
  l_��¼id   t_Numlist := t_Numlist();

  d_��ʼʱ�� �ٴ����ﰲ��.��ʼʱ��%Type;
  d_��ֹʱ�� �ٴ����ﰲ��.��ֹʱ��%Type;

  n_�����ܳ���id �ٴ������.Id%Type;

  Function Get�����ܳ���id(����id_In �ٴ������.Id%Type) Return �ٴ������.Id%Type Is
    ----------------------------------------
    --���ԭ�ܳ����������(����7��)������Ҫ���ҵ���һ�������������
    ----------------------------------------
    n_����id �ٴ������.Id%Type;
    n_���   �ٴ������.���%Type;
    n_�·�   �ٴ������.�·�%Type;
    n_����   �ٴ������.����%Type;
  
    d_��ʼʱ�� �ٴ����ﰲ��.��ʼʱ��%Type;
    d_����ʱ�� �ٴ����ﰲ��.��ֹʱ��%Type;
  
    --�������ڼ��㵱�µ��������Լ�ÿһ�ܵ�ʱ�䷶Χ
    Cursor c_Weekrange(Date_In Date) Is
      Select Rownum As ����, ��ʼ����, ��������
      From (With Month_Range As (Select Trunc(Date_In) As First_Day, Last_Day(Trunc(Date_In)) As Last_Day From Dual)
             Select Decode(To_Char(First_Day, 'day'), '������', First_Day, Null) As ��ʼ����,
                    Decode(To_Char(First_Day, 'day'), '������', First_Day, Null) As ��������
             From Month_Range
             Union All
             Select Decode(Sign(Trunc(First_Day + 7 * Week, 'day') + 1 - First_Day), 1,
                            Trunc(First_Day + 7 * Week, 'day') + 1, First_Day) As ��ʼ����,
                    Decode(Sign(Trunc(First_Day + 7 * Week, 'day') + 7 - Last_Day), 1, Last_Day,
                            Trunc(First_Day + 7 * Week, 'day') + 7) As ��������
             From Month_Range A, (Select Level - 1 As Week From Dual Connect By Level <= 6) B)
             Where ��ʼ���� <= ��������;
  
  
  Begin
    Begin
      Select ���, �·�, ���� Into n_���, n_�·�, n_���� From �ٴ������ Where ID = ����id_In;
    Exception
      When Others Then
        Return 0;
    End;
  
    If n_��� Is Null Or n_�·� Is Null Or n_���� Is Null Then
      Return 0;
    End If;
  
    For r_Weekrange In c_Weekrange(To_Date(n_��� || '-' || n_�·� || '-01', 'yyyy-mm-dd')) Loop
      If r_Weekrange.���� = n_���� Then
        d_��ʼʱ�� := r_Weekrange.��ʼ����;
        d_����ʱ�� := r_Weekrange.��������;
        Exit;
      End If;
    End Loop;
  
    If d_��ʼʱ�� Is Null Or d_����ʱ�� Is Null Then
      Return 0;
    End If;
    If Trunc(d_����ʱ��) - Trunc(d_��ʼʱ��) >= 6 Then
      Return 0;
    End If;
  
    --���ڿ��µģ�������һ��������������
    n_��� := Null;
    n_�·� := Null;
    n_���� := Null;
    If Trunc(d_��ʼʱ�� - 1, 'month') <> Trunc(d_��ʼʱ��, 'month') Then
      --��ǰ�ǵ�һ��,��ȡ��һ������������
      n_��� := To_Number(To_Char(d_��ʼʱ�� - 1, 'yyyy'));
      n_�·� := To_Number(To_Char(d_��ʼʱ�� - 1, 'mm'));
    Elsif Trunc(d_����ʱ�� + 1, 'month') <> Trunc(d_����ʱ��, 'month') Then
      --��ǰ�����һ��,��ȡ��һ������������
      n_��� := To_Number(To_Char(d_����ʱ�� + 1, 'yyyy'));
      n_�·� := To_Number(To_Char(d_����ʱ�� + 1, 'mm'));
      n_���� := 1;
    Else
      Return 0;
    End If;
  
    --��ȡ���µ���һ��������ID
    Begin
      Select ID
      Into n_����id
      From (Select Rownum As �к�, ID
             From �ٴ������
             Where Nvl(�Ű෽ʽ, 0) = 2 And ��� = n_��� And �·� = n_�·� And (n_���� Is Null Or ���� = n_����)
             Order By ���� Desc)
      Where �к� < 2;
    Exception
      When Others Then
        Return 0;
    End;
  
    Return n_����id;
  End;
Begin
  Begin
    Select Nvl(�Ű෽ʽ, 0) Into n_�Ű෽ʽ From �ٴ������ Where ID = Id_In;
  Exception
    When Others Then
      v_Err_Msg := '�������Ϣδ�ҵ���';
      Raise Err_Item;
  End;

  If Nvl(ȡ������_In, 0) = 0 Then
    --��������
    If Nvl(n_�Ű෽ʽ, 0) = 0 Then
      Select Max(1)
      Into n_Count
      From �ٴ����ﰲ�� A, �ٴ��������� B, �ٴ������ C
      Where a.Id = b.����id And a.����id = c.Id And c.�Ű෽ʽ = 0 And c.Id = Id_In And Rownum < 2;
      If Nvl(n_Count, 0) = 0 Then
        v_Err_Msg := '��ǰ���������Ч�İ��ţ����ܷ�����';
        Raise Err_Item;
      End If;
    Else
      If Nvl(n_�Ű෽ʽ, 0) = 2 Then
        n_�����ܳ���id := Get�����ܳ���id(Id_In);
      End If;
      Select Max(1)
      Into n_Count
      From �ٴ����ﰲ�� A, �ٴ������¼ B, �ٴ������ C
      Where a.Id = b.����id And a.����id = c.Id And c.�Ű෽ʽ In (1, 2) And (c.Id = Id_In Or c.Id = n_�����ܳ���id) And Rownum < 2;
      If Nvl(n_Count, 0) = 0 Then
        v_Err_Msg := '��ǰ���������Ч�İ��ţ����ܷ�����';
        Raise Err_Item;
      End If;
    
      Select Max(1)
      Into n_Count
      From �ٴ������¼ A, �ٴ����ﰲ�� B
      Where a.��Դid = b.��Դid And a.�������� Between b.��ʼʱ�� And b.��ֹʱ�� And a.����id <> b.Id And b.����id = Id_In And Rownum < 2;
      If Nvl(n_Count, 0) <> 0 Then
        v_Err_Msg := '��ǰ������еĲ��ֺ�Դ�ڵ�ǰ��������Чʱ�䷶Χ���Ѿ�������Ч�İ��ţ����ܷ�����';
        Raise Err_Item;
      End If;
    End If;
  
    --������ڶ��δ�����İ��ű����������������ڵİ��ţ����밴��С��Чʱ����з���
    Select Max(1)
    Into n_Count
    From (Select ID, ��� || LPad(�·�, 2, '0') || ���� As ����
           From �ٴ������
           Where Nvl(�Ű෽ʽ, 0) = Nvl(n_�Ű෽ʽ, 0) And ������ Is Null) A,
         (Select ID, ��� || LPad(�·�, 2, '0') || ���� As ���� From �ٴ������ Where ID = Id_In) B
    Where a.���� < b.���� And Rownum < 2;
    If Nvl(n_Count, 0) <> 0 Then
      If Nvl(n_�Ű෽ʽ, 0) = 0 Then
        v_Err_Msg := '��ǰ�����ǰ�滹��δ�����Ĺ̶�����������Ƚ��䷢����ɾ������ܷ����ó����';
      Elsif Nvl(n_�Ű෽ʽ, 0) = 1 Then
        v_Err_Msg := '��ǰ�����ǰ�滹��δ�������³���������Ƚ��䷢����ɾ������ܷ����ó����';
      Elsif Nvl(n_�Ű෽ʽ, 0) = 2 Then
        v_Err_Msg := '��ǰ�����ǰ�滹��δ�������ܳ���������Ƚ��䷢����ɾ������ܷ����ó����';
      End If;
      Raise Err_Item;
    End If;
  
    Update �ٴ������ Set ������ = ������_In, ����ʱ�� = ����ʱ��_In Where ID = Id_In;
    Update �ٴ����ﰲ�� Set ����� = ������_In, ���ʱ�� = ����ʱ��_In Where ����id = Id_In;
  
    --ɾ������ʱ�а��ţ����Ǻ�Դ�ѱ�ͣ�õļ�¼
    For c_���� In (Select a.Id
                 From �ٴ����ﰲ�� A, �ٴ������Դ B, ���ű� C, ��Ա�� D, �շ���ĿĿ¼ E
                 Where a.��Դid = b.Id And b.����id = c.Id And a.ҽ��id = d.Id(+) And b.��Ŀid = e.Id And a.����id = Id_In And
                       Not (Nvl(b.�Ƿ�ɾ��, 0) = 0 And (b.����ʱ�� Is Null Or b.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd')) And
                        Nvl(c.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                        Nvl(d.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                        Nvl(e.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd'))) Loop
      Zl_�ٴ����ﰲ��_Delete(c_����.Id, Nvl(n_�Ű෽ʽ, 0));
    End Loop;
  
    If Nvl(n_�Ű෽ʽ, 0) <> 0 Then
      --�°���/�ܰ��Ÿ���ͣ�ﰲ�źͷ����ڼ��յ��������¼�ĳ���/ԤԼ���
      Select ��ʼʱ��, ��ֹʱ�� Into d_��ʼʱ��, d_��ֹʱ�� From �ٴ����ﰲ�� Where ����id = Id_In And Rownum < 2;
      For c_���� In (Select a.Id, a.��Դid, b.����
                   From �ٴ����ﰲ�� A,
                        (Select Trunc(d_��ʼʱ��) + Level - 1 As ����
                          From Dual
                          Connect By Level <= Trunc(d_��ֹʱ��) - Trunc(d_��ʼʱ��) + 1) B
                   Where a.����id = Id_In
                   Order By ��Դid, ����) Loop
      
        Zl_Clinicvisitmodify(c_����.��Դid, c_����.Id, c_����.����, ������_In, ����ʱ��_In);
      End Loop;
    
      --�޸��ٴ������¼�е�"�Ƿ񷢲�"
      Select a.Id Bulk Collect
      Into l_��¼id
      From �ٴ������¼ A, �ٴ����ﰲ�� B
      Where a.����id = b.Id And b.����id = Id_In;
    
      Forall I In 1 .. l_��¼id.Count
        Update �ٴ������¼ Set �Ƿ񷢲� = 1 Where ID = l_��¼id(I);
    End If;
    Return;
  End If;

  --==================================================================================================================
  --ȡ������
  Select Max(1)
  Into n_Count
  From (Select ID, ��� || LPad(�·�, 2, '0') || ���� As ����
         From �ٴ������
         Where Nvl(�Ű෽ʽ, 0) = Nvl(n_�Ű෽ʽ, 0) And ������ Is Not Null) A,
       (Select ID, ��� || LPad(�·�, 2, '0') || ���� As ���� From �ٴ������ Where ID = Id_In) B
  Where a.���� > b.���� And Rownum < 2;
  If Nvl(n_Count, 0) <> 0 Then
    If Nvl(n_�Ű෽ʽ, 0) = 0 Then
      v_Err_Msg := '��ǰ������滹���ѷ����Ĺ̶�����������Ƚ���ȡ�����������ȡ�������ó����';
    Elsif Nvl(n_�Ű෽ʽ, 0) = 1 Then
      v_Err_Msg := '��ǰ������滹���ѷ������³���������Ƚ���ȡ�����������ȡ�������ó����';
    Elsif Nvl(n_�Ű෽ʽ, 0) = 2 Then
      v_Err_Msg := '��ǰ������滹���ѷ������ܳ���������Ƚ���ȡ�����������ȡ�������ó����';
    End If;
    Raise Err_Item;
  End If;

  Select Max(1)
  Into n_Count
  From ���˹Һż�¼ C, �ٴ������¼ A, �ٴ����ﰲ�� B
  Where c.�����¼id = a.Id And a.����id = b.Id And b.����id = Id_In And Rownum < 2;
  If Nvl(n_Count, 0) <> 0 Then
    v_Err_Msg := '��ǰ�����İ����ѱ�ʹ�ã�������ȡ��������';
    Raise Err_Item;
  End If;

  Update �ٴ������ Set ������ = Null, ����ʱ�� = Null Where ID = Id_In;
  If Sql%NotFound Then
    v_Err_Msg := '�������Ϣδ�ҵ���';
    Raise Err_Item;
  End If;
  Update �ٴ����ﰲ�� Set ����� = Null, ���ʱ�� = Null Where ����id = Id_In;

  --�̶�����ȡ������ʱɾ�������¼
  If Nvl(n_�Ű෽ʽ, 0) = 0 Then
    --ɾ�������¼
    Select a.Id Bulk Collect
    Into l_��¼id
    From �ٴ������¼ A, �ٴ����ﰲ�� B
    Where a.����id = b.Id And b.����id = Id_In;
  
    Zl_�ٴ������¼_Batchdelete(l_��¼id);
  Else
    --ɾ�����ݵĳ����¼
    Select a.Id Bulk Collect
    Into l_��¼id
    From �ٴ������¼ A, �ٴ����ﰲ�� B
    Where a.����id = b.Id And b.����id = Id_In And a.���id Is Not Null;
  
    Zl_�ٴ������¼_Batchdelete(l_��¼id);
  
    --�°���/�ܰ������ͣ����Ϣ�����޸��Ƿ񷢲�
    Select a.Id Bulk Collect
    Into l_��¼id
    From �ٴ������¼ A, �ٴ����ﰲ�� B
    Where a.����id = b.Id And b.����id = Id_In;
  
    Forall I In 1 .. l_��¼id.Count
      Delete From �ٴ�����ͣ���¼ Where ��¼id = l_��¼id(I);
  
    --�޸��ٴ������¼�е�"�Ƿ񷢲�"
    Select a.Id Bulk Collect
    Into l_��¼id
    From �ٴ������¼ A, �ٴ����ﰲ�� B
    Where a.����id = b.Id And b.����id = Id_In;
  
    Forall I In 1 .. l_��¼id.Count
      Update �ٴ������¼
      Set ͣ�￪ʼʱ�� = Null, ͣ����ֹʱ�� = Null, ͣ��ԭ�� = Null, �Ƿ񷢲� = 0
      Where ID = l_��¼id(I);
  
    --�ָ��ٴ�������ſ��Ƶ�"�Ƿ�ԤԼ"��"�Ƿ�ͣ��"
    For c_��¼ In (Select a.Id, a.�Ƿ��ʱ��, a.�Ƿ���ſ���
                 From �ٴ������¼ A, �ٴ����ﰲ�� B
                 Where a.����id = b.Id And b.����id = Id_In) Loop
      If Nvl(c_��¼.�Ƿ��ʱ��, 0) = 1 Then
        If Nvl(c_��¼.�Ƿ���ſ���, 0) = 0 Then
          Update �ٴ�������ſ��� Set �Ƿ�ԤԼ = 1 Where ��¼id = c_��¼.Id;
        Else
          Update �ٴ�������ſ��� Set �Ƿ�ԤԼ = Nvl(ԤԼ˳���, 0), �Ƿ�ͣ�� = 0 Where ��¼id = c_��¼.Id;
        End If;
      End If;
    End Loop;
  
    --���ݵĲ��ٻָ�
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ����ﰲ��_Publish;
/

--108432:Ƚ����,2017-05-08,�����̶��������ʱ����ȡ����˺�û��ɾ���ɸ���ʱ�������ɵĳ����¼�������������ʱ����ʱ���������
Create Or Replace Procedure Zl_�ٴ�������ʱ����_Cancel(����id_In In �ٴ����ﰲ��.Id%Type) Is
  Err_Item Exception;
  v_Err_Msg Varchar2(500);

  n_Count  Number(2);
  l_��¼id t_Numlist := t_Numlist();
Begin
  Select Count(1)
  Into n_Count
  From �ٴ������¼ A, ���˹Һż�¼ B
  Where a.Id = b.�����¼id And a.����id = ����id_In And Rownum < 2;
  If n_Count <> 0 Then
    v_Err_Msg := '��ǰ�����Ѵ���ԤԼ�Һ����ݣ�����ȡ����ˣ�';
    Raise Err_Item;
  End If;

  Select Count(1)
  Into n_Count
  From �ٴ����ﰲ�� A, �ٴ����ﰲ�� B, �ٴ������ C
  Where a.��Դid = b.��Դid And a.����id = c.Id And c.�Ű෽ʽ = 0 And a.Id <> b.Id And b.Id = ����id_In And a.�Ǽ�ʱ�� > b.�Ǽ�ʱ�� And
        a.���ʱ�� Is Not Null And Rownum < 2;
  If n_Count <> 0 Then
    v_Err_Msg := '�ú�Դ�ڵ�ǰ����֮�󻹴�������˵İ��ţ��㲻��ȡ����˵�ǰ���ţ�';
    Raise Err_Item;
  End If;

  Update �ٴ����ﰲ�� Set ����� = Null, ���ʱ�� = Null Where ID = ����id_In And ���ʱ�� Is Not Null;
  If Sql%NotFound Then
    v_Err_Msg := '��ǰ�����ѱ�����ȡ����˻�ɾ����������ȡ����ˣ�';
    Raise Err_Item;
  End If;

  --ɾ���ð��������ɵĳ����¼
  Select a.Id Bulk Collect Into l_��¼id From �ٴ������¼ A Where a.����id = ����id_In;
  Zl_�ٴ������¼_Batchdelete(l_��¼id);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ�������ʱ����_Cancel;
/

--108667:Ƚ����,2017-05-08,�����Զ�������۵�ʱ����ͬһ��ҽ������ҩƷҲ��������Ŀ����������Ŀ����ִ��ʱ����
Create Or Replace Procedure Zl_���ﻮ�ۼ�¼_Clear(Day_In Number) As
  --���ܣ��Զ�������۵� 
  --������Day_IN=ɾ�����ۺ󳬹�Day_IN��δ�շѵĵ��� 
  Cursor c_Price Is
    Select Distinct a.No, f_List2str(Cast(Collect(To_Char(a.���)) As t_Strlist)) As ���
    From ������ü�¼ A, δ��ҩƷ��¼ B
    Where a.��¼���� = 1 And a.��¼״̬ = 0 And a.ִ��״̬ Not In (1, 2) And a.������ Is Not Null And a.����Ա���� Is Null And
          b.���� In (8, 24) And Nvl(b.���շ�, 0) = 0 And a.No = b.No And Nvl(a.ִ�в���id, 0) = Nvl(b.�ⷿid, 0) And
          Sysdate - b.�������� >= Day_In
    Group By a.No;
Begin
  For r_Price In c_Price Loop
    If Not r_Price.��� Is Null Then
      Zl_���ﻮ�ۼ�¼_Delete(r_Price.No, r_Price.���, 1);
      Commit;
    End If;
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���ﻮ�ۼ�¼_Clear;
/

--97423:��ΰ��,2017-05-05,��������Ժ����ȡ���Ǽ�ʱ�Ҳ������ݵ�����
Create Or Replace Procedure Zl_��Ժ������ҳ_Delete
(
  ����id_In     ������ҳ.����id%Type,
  ��ҳid_In     ������ҳ.��ҳid%Type,
  ת����_In     Number := 0,
  ���סԺ��_In Number := 0
  --���ܣ�ȡ��������Ժ/ԤԼ�Ǽ�
  --     ��ҳID_IN:Ϊ0ʱ��ʾȡ��ԤԼ�Ǽ�
  --     ת����_IN:��������Ժ�Ǽǲ���תΪסԺ���۲���
  --     ���סԺ��_In:��һ��סԺ�Ĳ���ת����ʱ�Ƿ����סԺ��
) As
  v_��Ժʱ��   ������ҳ.��Ժ����%Type;
  v_��Ժ����   ������ҳ.��Ժ����id%Type;
  v_��Ժʱ��   ������ҳ.��Ժ����%Type;
  v_סԺ��     ������ҳ.סԺ��%Type;
  v_����Ժ     ������ҳ.����Ժ%Type;
  v_��Ժ����id ������ҳ.��Ժ����id%Type;
  n_��������   ������ҳ.��������%Type;
  n_��ҳid     ������ҳ.��ҳid%Type;

  v_Count Number;
  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  Select Nvl(״̬, 0), Nvl(��������, 0)
  Into v_Count, n_��������
  From ������ҳ
  Where ����id = ����id_In And ��ҳid = ��ҳid_In;
  If v_Count <> 1 Then
    v_Error := '�ò����Ѿ����,���Ƚ����˳�������Ժ״̬��';
    Raise Err_Custom;
  End If;

  --ɾ�����Ӳ���ʱ��
  Select ��Ժ����id, ����Ժ Into v_��Ժ����id, v_����Ժ From ������ҳ Where ����id = ����id_In And ��ҳid = ��ҳid_In;
  If v_����Ժ = 0 Then
    Zl_���Ӳ���ʱ��_Delete(����id_In, ��ҳid_In, '��Ժ', v_��Ժ����id);
  Else
    Zl_���Ӳ���ʱ��_Delete(����id_In, ��ҳid_In, '�ٴ���Ժ', v_��Ժ����id);
  End If;

  --��ȡ���һ�β�Ϊ�յ�סԺ��
  Begin
    If ��ҳid_In = 0 Then
      Select סԺ��
      Into v_סԺ��
      From ������ҳ
      Where ����id = ����id_In And
            ��ҳid =
            (Select Max(��ҳid) From ������ҳ Where ����id = ����id_In And Nvl(��ҳid, 0) <> 0 And Nvl(סԺ��, 0) <> 0);
    Else
      Select סԺ��
      Into v_סԺ��
      From ������ҳ
      Where ����id = ����id_In And
            ��ҳid =
            (Select Max(��ҳid) From ������ҳ Where ����id = ����id_In And ��ҳid < ��ҳid_In And Nvl(סԺ��, 0) <> 0);
    End If;
  Exception
    When Others Then
      Null;
  End;

  If ת����_In = 1 And Nvl(��ҳid_In, 0) <> 0 Then
    Update ������ҳ
    Set �������� = 2, סԺ�� = Decode(���סԺ��_In, 1, Null, סԺ��)
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And Nvl(��������, 0) = 0;
  
    --����סԺ����
    Update ������Ϣ Set סԺ���� = Decode(Sign(סԺ���� - 1), 1, סԺ���� - 1, Null) Where ����id = ����id_In;
    If ���סԺ��_In = 1 Then
      Update ������Ϣ Set סԺ�� = v_סԺ�� Where ����id = ����id_In;
    End If;
  Else
    Begin
      Select b.��Ժ����, b.��Ժ����, b.��Ժ����id
      Into v_��Ժʱ��, v_��Ժʱ��, v_��Ժ����
      From ������Ϣ A, ������ҳ B
      Where a.����id = ����id_In And a.����id = b.����id And a.��ҳid = b.��ҳid And Nvl(b.��ҳid, 0) <> 0;
    Exception
      When Others Then
        Null;
    End;
    --����ԤԼ�Ǽǲ��˲����סԺ�ձ�
    If Nvl(��ҳid_In, 0) <> 0 Then
      Select Zl_סԺ�ձ�_Count(v_��Ժ����, v_��Ժʱ��) Into v_Count From Dual;
      If v_Count > 0 Then
        v_Error := '�Ѳ���ҵ��ʱ���ڵ�סԺ�ձ�,���ܰ����ҵ��!';
        Raise Err_Custom;
      End If;
    End If;
    --�������۲����´���Ժ֪ͨ�����������Ч�Ĳ�����ҳ��¼��36549��
    Select Count(*) Into v_Count From ������ҳ Where ����id = ����id_In And ��Ժ���� Is Not Null And ��Ժ���� Is Null;
    If Not v_Count > 1 Then
      v_Count := 0;
      If Nvl(��ҳid_In, 0) <> 0 And Nvl(n_��������, 0) = 0 Then
        v_Count := 1;
      End If;
      --����Ժ����,ȡ����Ժ�Ǽ�ʱ,������Ϣ����Ժʱ��ͳ�Ժʱ��Ӧ�û��˵���һ����Ժ���ںͳ�Ժ����
      If v_����Ժ = 1 Then
        Begin
          Select ��Ժ����, ��Ժ����
          Into v_��Ժʱ��, v_��Ժʱ��
          From ������ҳ
          Where ����id = ����id_In And
                ��ҳid = (Select Max(��ҳid)
                        From ������ҳ
                        Where ����id = ����id_In And ��ҳid < ��ҳid_In And Nvl(סԺ��, 0) <> 0);
        Exception
          When Others Then
            --�쳣������Ϊ������ȡ�������ݵ��쳣���
            Null;
        End;
      End If;    
      Update ������Ϣ
      Set סԺ�� = v_סԺ��, סԺ���� = Decode(v_Count, 0, סԺ����, Decode(Sign(סԺ���� - 1), 1, סԺ���� - 1, Null)), ��ǰ����id = Null,
          ��ǰ����id = Null, ��ǰ���� = Null, ��Ժʱ�� = v_��Ժʱ��, ��Ժʱ�� = v_��Ժʱ��, ������ = Null, ������ = Null, �������� = Null, ��Ժ = Null
      Where ����id = ����id_In;
      Delete From ��Ժ���� Where ����id = ����id_In;
    End If;
    Delete From ���˱䶯��¼ Where ����id = ����id_In And ��ҳid = ��ҳid_In;
    Delete From ������ϼ�¼ Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼��Դ = 2;
  
    --����סԺ�������Ԥ����,��Ϊ�������ｻ��
    Update ����Ԥ����¼ Set ��ҳid = Null Where ����id = ����id_In And ��ҳid = ��ҳid_In;
  
    --���η�����,�ı����﷢��
    Update סԺ���ü�¼ Set ��ҳid = Null Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼���� = 5;
  
    --����סԺ�����з��ü�¼�޽�������ȫ���������򽫶�Ӧ���ü�¼�е�"��ҳID"�����
    v_Count := 0;
    Select Nvl(Count(*), 0)
    Into v_Count
    From סԺ���ü�¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And ���ʷ��� = 1 And ����id Is Not Null;
  
    If v_Count = 0 Then
      Begin
        Select Nvl(Count(*), 0)
        Into v_Count
        From סԺ���ü�¼
        Where ����id = ����id_In And ��ҳid = ��ҳid_In And ���ʷ��� = 1
        Group By NO, ��¼����, ���
        Having Nvl(Sum(ʵ�ս��), 0) <> 0;
      Exception
        When Others Then
          v_Count := 0;
      End;
    
      If v_Count = 0 Then
        Delete ����δ����� Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��� = 0;
        Update סԺ���ü�¼ Set ��ҳid = Null Where ����id = ����id_In And ��ҳid = ��ҳid_In And ���ʷ��� = 1;
      End If;
    End If;
  
    --����סԺ����ҽ����¼��������
    v_Count := 0;
    Select Nvl(Count(*), 0)
    Into v_Count
    From ����ҽ����¼
    Where ����id = ����id_In And ��ҳid = ��ҳid_In And Nvl(ҽ��״̬, 0) <> 4;
    If v_Count = 0 Then
      Delete From ����ҽ����¼ Where ����id = ����id_In And ��ҳid = ��ҳid_In;
    End If;
  
    --���±�,û�н�������ҳ(����ID,��ҳID)�����,��Ϊ����ҳID�����ǹҺ�ID
    Delete From ���˹�����¼ Where ����id = ����id_In And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0);
    Delete From ������ϼ�¼ Where ����id = ����id_In And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0);
    Delete From ������������¼ Where ����id = ����id_In And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0);
    Delete From ���Ӳ�����¼ Where ����id = ����id_In And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0);
    Delete From ���Ӳ�����ӡ Where ����id = ����id_In And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0);
    --�����Ժ�����˾��￨,��ɾ����ʧ��(���˷��ü�¼��ҳID�����Լ��)
    Delete From ������ҳ Where ����id = ����id_In And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0);
    --�޸Ĳ�����Ϣ����ҳID��סԺ����
    Select Max(��ҳid) Into n_��ҳid From ������ҳ Where ����id = ����id_In And Nvl(��ҳid, 0) <> 0;
    Update ������Ϣ Set ��ҳid = n_��ҳid Where ����id = ����id_In;
    If n_��ҳid Is Null Then
      Update ������Ϣ Set סԺ���� = Null Where ����id = ����id_In;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_��Ժ������ҳ_Delete;
/

--108667:Ƚ����,2017-04-28,�����Զ�������۵�ʱ����ͬһ��ҽ������ҩƷҲ��������Ŀ����������Ŀ����ִ��ʱ����
Create Or Replace Procedure Zl_���ﻮ�ۼ�¼_Delete
(
  No_In       ������ü�¼.No%Type,
  ���_In     Varchar2 := Null,
  �Զ����_In Number := 0
) As
  --���ܣ�ɾ��һ�����ﻮ�۵���
  --��Σ�
  --       ���_In����Ҫ��������ҽ��վ���ϵ���ҩƷ
  --      �Զ����_in���Ƿ��Զ�������۵� zl_���ﻮ�ۼ�¼_clear �ڵ���
  --�ù�����ڴ���ҩƷ����������
  Cursor c_Stock Is
    Select ��ҩ��ʽ, �ⷿid, ����, ҩƷid, ʵ������, ����, ���Ч��, ����, ����, Ч��, ID, ��Ʒ����, �ڲ�����, ����id
    From ҩƷ�շ���¼
    Where ���� In (8, 24) And NO = No_In And Mod(��¼״̬, 3) = 1 And ����� Is Null And
          ����id In (Select ID
                   From ������ü�¼
                   Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 0 And �շ���� In ('4', '5', '6', '7') And
                         (Instr(',' || ���_In || ',', ',' || ��� || ',') > 0 Or ���_In Is Null))
    Order By ҩƷid;
  --���α����ڴ�����ü�¼���
  Cursor c_Serial Is
    Select ID, �۸񸸺� From ������ü�¼ Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 0 Order By ���;

  v_ҽ��ids  Varchar2(4000);
  l_ҽ��id   t_Numlist := t_Numlist();
  l_ҩƷ�շ� t_Numlist := t_Numlist();
  v_ҽ��id   ����ҽ����¼.Id%Type;
  l_����id   t_Numlist := t_Numlist();
  n_�������� Number;

  n_����         ������ü�¼.���%Type;
  n_Count        Number;
  n_ҽ����       Number(5);
  n_��ִ��_Count Number;

  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin

  --�Ƿ��Ѿ�ɾ�����շ�
  Select Nvl(Count(ID), 0), Sum(Decode(ҽ�����, Null, 0, 1)), Max(ҽ�����), Sum(Decode(Nvl(ִ��״̬, 0), 1, 1, 2, 1, 0))
  Into n_Count, n_ҽ����, v_ҽ��id, n_��ִ��_Count
  From ������ü�¼
  Where ��¼���� = 1 And ��¼״̬ = 0 And NO = No_In And
        (Instr(',' || ���_In || ',', ',' || Nvl(�۸񸸺�, ���) || ',') > 0 Or ���_In Is Null);

  If n_Count = 0 Then
    If Nvl(�Զ����_In, 0) = 1 Then
      --�Զ�������۵�����ʱ������ֱ���˳�
      Return;
    Else
      v_Err_Msg := 'Ҫɾ���ķ��ü�¼�����ڣ������Ѿ�ɾ�����Ѿ��շѡ�';
      Raise Err_Item;
    End If;
  End If;
  --�Ƿ��Ѿ�ִ��
  If Nvl(n_��ִ��_Count, 0) > 0 Then
    v_Err_Msg := 'Ҫɾ���ķ��ü�¼�а�����ִ�е����ݣ�';
    Raise Err_Item;
  End If;

  --ҽ�����ã��������ִ�е�ҽ��(ע����ִ�е������������,��Ϊ���� ���_IN ����������ý���������)
  --�Զ�������۵�����ʱ������ֻ�ᴫ��ҩƷ���ĵĶ�Ӧ��ţ����Բ��ü��ҽ����
  --������ҽ��������ͬһ��ҽ���м���ҩƷ��Ҳ��������Ŀ����������Ŀ����ִ�л���ִ��ʱ��ҩƷ���ۼ�¼��ɾ������
  If Nvl(�Զ����_In, 0) = 0 Then
    Select Nvl(Count(*), 0)
    Into n_Count
    From ����ҽ������
    Where ִ��״̬ = 3 And (NO, ��¼����, ҽ��id) In
          (Select NO, ��¼����, ҽ�����
                        From ������ü�¼
                        Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 0 And ҽ����� Is Not Null And
                              (Instr(',' || ���_In || ',', ',' || ��� || ',') > 0 Or ���_In Is Null));
    If n_Count > 0 Then
      v_Err_Msg := 'Ҫɾ���ķ����д��ڶ�Ӧ��ҽ������ִ�е����������ɾ����';
      Raise Err_Item;
    End If;
  End If;

  --ҩƷ�������
  --�ȴ���������
  For v_���� In (Select ��ҩ��ʽ, �ⷿid, ����, ҩƷid, ʵ������, ����, ���Ч��, ����, ����, Ч��, ID, ��Ʒ����, �ڲ�����, ����id
               From ҩƷ�շ���¼
               Where ���� = 21 And Mod(��¼״̬, 3) = 1 And ����� Is Null And
                     ����id In (Select ID
                              From ������ü�¼
                              Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 0 And �շ���� = '4' And
                                    (Instr(',' || ���_In || ',', ',' || ��� || ',') > 0 Or ���_In Is Null))
               Order By ҩƷid) Loop
  
    If v_����.�ⷿid Is Not Null Then
      Update ҩƷ���
      Set �������� = Nvl(��������, 0) + Decode(v_����.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(v_����.����, 1) * Nvl(v_����.ʵ������, 0)
      Where �ⷿid = v_����.�ⷿid And ҩƷid = v_����.ҩƷid And Nvl(����, 0) = Nvl(v_����.����, 0) And ���� = 1;
      If Sql%RowCount = 0 Then
        Insert Into ҩƷ���
          (�ⷿid, ҩƷid, ����, ����, Ч��, ��������, �ϴ�����, �ϴβ���, ���Ч��, ��Ʒ����, �ڲ�����)
        Values
          (v_����.�ⷿid, v_����.ҩƷid, 1, v_����.����, v_����.Ч��,
           Decode(v_����.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(v_����.����, 1) * Nvl(v_����.ʵ������, 0), v_����.����, v_����.����, v_����.���Ч��,
           v_����.��Ʒ����, v_����.�ڲ�����);
      End If;
    End If;
    l_ҩƷ�շ�.Extend;
    l_ҩƷ�շ�(l_ҩƷ�շ�.Count) := v_����.Id;
  
    l_����id.Extend;
    l_����id(l_����id.Count) := v_����.����id;
  End Loop;

  For r_Stock In c_Stock Loop
  
    If r_Stock.�ⷿid Is Not Null Then
    
      Select Decode(Count(Column_Value), Null, 0, 0, 0, 1)
      Into n_��������
      From Table(l_����id)
      Where Column_Value = r_Stock.����id;
      If Nvl(n_��������, 0) = 0 Then
        Update ҩƷ���
        Set �������� = Nvl(��������, 0) + Decode(r_Stock.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(r_Stock.����, 1) * Nvl(r_Stock.ʵ������, 0)
        Where �ⷿid = r_Stock.�ⷿid And ҩƷid = r_Stock.ҩƷid And Nvl(����, 0) = Nvl(r_Stock.����, 0) And ���� = 1;
        If Sql%RowCount = 0 Then
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, Ч��, ��������, �ϴ�����, �ϴβ���, ���Ч��, ��Ʒ����, �ڲ�����)
          Values
            (r_Stock.�ⷿid, r_Stock.ҩƷid, 1, r_Stock.����, r_Stock.Ч��,
             Decode(r_Stock.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(r_Stock.����, 1) * Nvl(r_Stock.ʵ������, 0), r_Stock.����, r_Stock.����,
             r_Stock.���Ч��, r_Stock.��Ʒ����, r_Stock.�ڲ�����);
        End If;
      End If;
    End If;
    l_ҩƷ�շ�.Extend;
    l_ҩƷ�շ�(l_ҩƷ�շ�.Count) := r_Stock.Id;
  End Loop;

  --ɾ��ҩƷ�շ���¼
  Forall I In 1 .. l_ҩƷ�շ�.Count
    Delete From ҩƷ�շ���¼ Where ID = l_ҩƷ�շ�(I);

  ------------------------------------------------------------------------------------------------------------------------
  --����ɾδ��ҩƷ��¼
  Delete From δ��ҩƷ��¼ A
  Where NO = No_In And ���� In (8, 24) And Not Exists
   (Select 1
         From ҩƷ�շ���¼
         Where ���� = a.���� And Nvl(�ⷿid, 0) = Nvl(a.�ⷿid, 0) And NO = No_In And Mod(��¼״̬, 3) = 1 And ����� Is Null);
  ------------------------------------------------------------------------------------------------------------------------

  --ɾ������ҽ������(���һ��ɾ��ʱ)
  If ���_In Is Null Then
    --Begin
    --  Select ҽ�����
    --  Into v_ҽ��id
    --  From ������ü�¼
    --  Where ��¼���� = 1 And ��¼״̬ = 0 And NO = No_In And Rownum = 1;
    -- Exception
    --  When Others Then
    --    Null;
    -- End;
  
    If v_ҽ��id Is Not Null Then
      Delete From ����ҽ������ Where ҽ��id = v_ҽ��id And NO = No_In And ��¼���� = 1;
    End If;
  End If;

  If n_ҽ���� > 0 Then
    If n_ҽ���� = 1 Then
      l_ҽ��id.Extend;
      l_ҽ��id(l_ҽ��id.Count) := v_ҽ��id;
    Else
      Select Distinct ҽ����� Bulk Collect
      Into l_ҽ��id
      From ������ü�¼
      Where ��¼���� = 1 And ��¼״̬ = 0 And NO = No_In And ҽ����� Is Not Null And
            (Instr(',' || ���_In || ',', ',' || Nvl(�۸񸸺�, ���) || ',') > 0 Or ���_In Is Null);
    End If;
  End If;

  --������ü�¼
  Delete From ������ü�¼
  Where ��¼���� = 1 And ��¼״̬ = 0 And NO = No_In And
        (Instr(',' || ���_In || ',', ',' || Nvl(�۸񸸺�, ���) || ',') > 0 Or ���_In Is Null);
  If Sql%RowCount = 0 Then
    If Nvl(�Զ����_In, 0) = 1 Then
      --�Զ�������۵�����ʱ������ֱ���˳�
      Return;
    Else
      v_Err_Msg := 'Ҫɾ���ķ��ü�¼�����ڣ������Ѿ�ɾ�����Ѿ��շѡ�';
      Raise Err_Item;
    End If;
  End If;

  If ���_In Is Not Null Then
    --���µ���ʣ����÷��ü�¼�����
    n_Count := 1;
    For r_Serial In c_Serial Loop
      If r_Serial.�۸񸸺� Is Null Then
        n_���� := n_Count;
      End If;
      Update ������ü�¼ Set ��� = n_Count, �۸񸸺� = Decode(�۸񸸺�, Null, Null, n_����) Where ID = r_Serial.Id;
    
      n_Count := n_Count + 1;
    End Loop;
  End If;
  v_ҽ��ids := Null;
  For I In 1 .. l_ҽ��id.Count Loop
    v_ҽ��ids := Nvl(v_ҽ��ids, '') || ',' || l_ҽ��id(I);
  End Loop;
  If v_ҽ��ids Is Not Null Then
    v_ҽ��ids := Substr(v_ҽ��ids, 2);
    --����_In    Integer, --0:����;1-סԺ
    --����_In    Integer, --1-�շѵ�;2-���ʵ�
    --����_In    Integer, --0:ɾ�����۵�;1-�շѻ����;2-�˷ѻ�����
    --No_In      ������ü�¼.No%Type,
    --ҽ��ids_In Varchar2
    Zl_ҽ������_�Ʒ�״̬_Update(0, 1, 0, No_In, v_ҽ��ids);
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���ﻮ�ۼ�¼_Delete;
/

--108251:������,2017-04-27,�������۵��ĹҺŵ��˺Ŵ���
Create Or Replace Procedure Zl_Third_Registdelcheck
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --����:HIS�˺ż��
  --���:Xml_In:
  --<IN>
  --  <GHDH>A000001</GHDH>    //�Һŵ���
  --  <JSKLB>֧����</JSKLB>      //���㿨���
  --  <JCFP>1</JCFP>            //��鷢Ʊ
  --  <GHJE>20</GHJE>            //�ҺŽ��
  --  <LSH>34563</LSH>           //������ˮ��
  --  <JKFS>0</JKFS>             //�ɿʽ,0-�ҺŻ�ԤԼ�ɿ�;1-ԤԼ���ɿ�
  --  <YYFS></YYFS>              //�ɿʽ=1ʱ���룬ԤԼ��ԤԼ��ʽ
  --  <XL></XL>                  //����
  --</IN>

  --����:Xml_Out
  --<OUTPUT>
  -- <ERROR><MSG></MSG></ERROR> //Ϊ�ձ�ʾ���ɹ�
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  v_�����     Varchar2(100);
  v_No         ���˹Һż�¼.No%Type;
  n_�ҺŽ��   ������ü�¼.ʵ�ս��%Type;
  v_����Ա��� ������ü�¼.����Ա���%Type;
  v_����Ա���� ������ü�¼.����Ա����%Type;
  v_���㷽ʽ   ҽ�ƿ����.���㷽ʽ%Type;
  n_ʵ�ս��   ������ü�¼.ʵ�ս��%Type;
  v_������ˮ�� ����Ԥ����¼.������ˮ��%Type;
  n_����       Number(3);
  v_Type       Varchar2(50);
  v_Temp       Varchar2(32767); --��ʱXML
  x_Templet    Xmltype; --ģ��XML
  v_Err_Msg    Varchar2(200);
  n_�ѿ�ҽ��   Number(2);
  n_��鷢Ʊ   Number(3);
  n_�Ƿ��ӡ   Number(3);
  n_�ɿʽ   Number(3);
  n_����       ������Ϣ.����%Type;
  v_ԤԼ��ʽ   ���˹Һż�¼.ԤԼ��ʽ%Type;
  v_�շѵ�     ������ü�¼.No%Type;
  Err_Item Exception;
Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/GHDH'), Extractvalue(Value(A), 'IN/JSKLB'), Extractvalue(Value(A), 'IN/GHJE'),
         Extractvalue(Value(A), 'IN/LSH'), To_Number(Extractvalue(Value(A), 'IN/JCFP')),
         To_Number(Extractvalue(Value(A), 'IN/JKFS')), Extractvalue(Value(A), 'IN/YYFS'),
         To_Number(Extractvalue(Value(A), 'IN/XL'))
  Into v_No, v_�����, n_�ҺŽ��, v_������ˮ��, n_��鷢Ʊ, n_�ɿʽ, v_ԤԼ��ʽ, n_����
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  Select Max(�շѵ�) Into v_�շѵ� From ���˹Һż�¼ Where NO = v_No;

  n_�ɿʽ := Nvl(n_�ɿʽ, 0);

  If v_����� Is Not Null And n_�ɿʽ = 0 Then
    Select Nvl2(Translate(v_�����, '\1234567890', '\'), 'Char', 'Num') Into v_Type From Dual;
    If v_Type = 'Num' Then
      --������ǿ����ID
      Select ���㷽ʽ Into v_���㷽ʽ From ҽ�ƿ���� Where ID = To_Number(v_�����);
    Else
      --������ǿ��������
      Select ���㷽ʽ Into v_���㷽ʽ From ҽ�ƿ���� Where ���� = v_�����;
    End If;
    If Nvl(n_�ɿʽ, 0) = 0 Then
      If Nvl(n_����, 0) = 0 Then
        Select Nvl(Max(1), 0)
        Into n_����
        From ����Ԥ����¼ A,
             (Select Distinct ����id
               From ������ü�¼
               Where NO = v_No And ��¼���� = 4
               Union
               Select Distinct ����id
               From סԺ���ü�¼
               Where NO = v_No And ��¼���� = 5
               Union
               Select Distinct ����id
               From ������ü�¼
               Where NO = v_�շѵ� And ��¼���� = 1) B
        Where a.����id = b.����id And ���㷽ʽ <> v_���㷽ʽ And Mod(��¼����, 10) <> 1 And Rownum < 2;
      Else
        Select Nvl(Max(1), 0)
        Into n_����
        From ����Ԥ����¼ A,
             (Select Distinct ����id
               From ������ü�¼
               Where NO = v_No And ��¼���� = 4
               Union
               Select Distinct ����id
               From סԺ���ü�¼
               Where NO = v_No And ��¼���� = 5
               Union
               Select Distinct ����id
               From ������ü�¼
               Where NO = v_�շѵ� And ��¼���� = 1) B, ���㷽ʽ C
        Where a.����id = b.����id And ���㷽ʽ <> v_���㷽ʽ And Mod(��¼����, 10) <> 1 And a.���㷽ʽ = c.���� And c.���� Not In (3, 4) And
              Rownum < 2;
        If n_���� = 0 Then
          Select Nvl(Max(1), 0)
          Into n_����
          From ���ս����¼ A,
               (Select Distinct ����id
                 From ������ü�¼
                 Where NO = v_No And ��¼���� = 4
                 Union
                 Select Distinct ����id
                 From סԺ���ü�¼
                 Where NO = v_No And ��¼���� = 5
                 Union
                 Select Distinct ����id
                 From ������ü�¼
                 Where NO = v_�շѵ� And ��¼���� = 1) B
          Where a.��¼id = b.����id And ���� <> n_���� And Rownum < 2;
        End If;
      End If;
      If n_���� = 1 Then
        v_Err_Msg := '����ĹҺŵ��ݰ���' || v_���㷽ʽ || '����Ľ��㷽ʽ,�޷��˺�!';
        Raise Err_Item;
      End If;
    Else
      Begin
        Select 1 Into n_���� From ���˹Һż�¼ A Where a.No = v_No And a.ԤԼ��ʽ = v_ԤԼ��ʽ And Rownum < 2;
      Exception
        When Others Then
          n_���� := 0;
      End;
      If n_���� = 0 Then
        v_Err_Msg := '����ĹҺŵ��ݲ���' || v_ԤԼ��ʽ || 'ԤԼ��,�޷��˺�!';
        Raise Err_Item;
      End If;
    End If;
  End If;

  If n_�ɿʽ = 0 Then
    If v_�շѵ� Is Null Then
      Select Sum(ʵ�ս��) Into n_ʵ�ս�� From ������ü�¼ Where NO = v_No And ��¼���� = 4;
    Else
      Select Sum(ʵ�ս��) Into n_ʵ�ս�� From ������ü�¼ Where NO = v_�շѵ� And ��¼���� = 1;
    End If;
    If n_ʵ�ս�� <> n_�ҺŽ�� Then
      v_Err_Msg := '������˿�����ʵ�ʹҺŽ���������!';
      Raise Err_Item;
    End If;
  End If;

  --��������飬�Ѵ��ڲ��������ݵģ������˺�
  Begin
    Select 1
    Into n_����
    From ���ò����¼ A,
         (Select Distinct ����id
           From ������ü�¼
           Where NO = v_No And ��¼���� = 4
           Union
           Select Distinct ����id
           From סԺ���ü�¼
           Where NO = v_No And ��¼���� = 5
           Union
           Select Distinct ����id
           From ������ü�¼
           Where NO = v_�շѵ� And ��¼���� = 1) B
    Where a.�շѽ���id = b.����id And a.��¼���� = 1 And a.���ӱ�־ = 1 And Nvl(a.����״̬, 0) <> 2 And Rownum < 2;
  Exception
    When Others Then
      n_���� := 0;
  End;
  If n_���� = 1 Then
    v_Err_Msg := '����ĹҺŵ����Ѿ������˶��ν���,�޷��˺�!';
    Raise Err_Item;
  End If;
  --ҽ����飬�Ѿ�����ҽ���ģ������˺�
  Begin
    Select Distinct 1 Into n_�ѿ�ҽ�� From ����ҽ����¼ Where �Һŵ� = v_No;
  Exception
    When Others Then
      n_�ѿ�ҽ�� := 0;
  End;
  If n_�ѿ�ҽ�� = 1 Then
    v_Err_Msg := '����ĹҺŵ����Ѿ�����ҽ��,�޷��˺�!';
    Raise Err_Item;
  End If;
  If Nvl(n_��鷢Ʊ, 0) = 1 Then
    Select Max(Decode(a.ʵ��Ʊ��, Null, 0, 1)) Into n_�Ƿ��ӡ From ������ü�¼ A Where NO = v_No And ��¼���� = 4;
    If Nvl(n_�Ƿ��ӡ, 0) = 1 Then
      v_Err_Msg := '�����˺ŵĵ����ѿ���Ʊ,�����˷�!';
      Raise Err_Item;
    End If;
    Select Max(Decode(a.ʵ��Ʊ��, Null, 0, 1))
    Into n_�Ƿ��ӡ
    From ������ü�¼ A
    Where NO = v_�շѵ� And ��¼���� = 1;
    If Nvl(n_�Ƿ��ӡ, 0) = 1 Then
      v_Err_Msg := '�����˺ŵĵ����ѿ���Ʊ,�����˷�!';
      Raise Err_Item;
    End If;
  End If;

  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Registdelcheck;
/

--108251:������,2017-04-27,�������۵��ĹҺŵ�����
Create Or Replace Procedure Zl_Third_Registdel
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --����:HIS�˺�
  --���:Xml_In:
  --<IN>
  --  <GHDH>A000001</GHDH>    //�Һŵ���
  --  <JSKLB>֧����</JSKLB>      //���㿨���
  --  <JCFP>1</JCFP>            //��鷢Ʊ
  --  <GHJE>20</GHJE>            //�ҺŽ��
  --  <LSH>34563</LSH>           //������ˮ��
  --  <JKFS>0</JKFS>             //�ɿʽ,0-�ҺŻ�ԤԼ�ɿ�;1-ԤԼ���ɿ�
  --  <YYFS></YYFS>              //�ɿʽ=1ʱ���룬ԤԼ��ԤԼ��ʽ
  --</IN>

  --����:Xml_Out
  --<OUTPUT>
  -- <CZSJ>����ʱ��</CZSJ>          //HIS�ĵǼ�ʱ��
  -- <YJZID>ԭ����ID</YJZID>
  -- <CXID>����ID</CXID>
  -- <ERROR><MSG></MSG></ERROR> //Ϊ�ձ�ʾȡ���Һųɹ�
  --</OUTPUT>
  --------------------------------------------------------------------------------------------------
  v_�����     Varchar2(100);
  v_No         ���˹Һż�¼.No%Type;
  n_�ҺŽ��   ������ü�¼.ʵ�ս��%Type;
  v_����Ա��� ������ü�¼.����Ա���%Type;
  v_����Ա���� ������ü�¼.����Ա����%Type;
  v_���㷽ʽ   ҽ�ƿ����.���㷽ʽ%Type;
  n_ʵ�ս��   ������ü�¼.ʵ�ս��%Type;
  v_������ˮ�� ����Ԥ����¼.������ˮ��%Type;
  n_����       Number(3);
  v_Type       Varchar2(50);
  v_Temp       Varchar2(32767); --��ʱXML
  x_Templet    Xmltype; --ģ��XML
  v_Err_Msg    Varchar2(200);
  n_�ѿ�ҽ��   Number(2);
  n_��鷢Ʊ   Number(3);
  n_�Ƿ��ӡ   Number(3);
  n_�ɿʽ   Number(3);
  n_����id     ������ü�¼.����id%Type;
  n_����id     ������ü�¼.����id%Type;
  d_�Ǽ�ʱ��   Date;
  v_ԤԼ��ʽ   ���˹Һż�¼.ԤԼ��ʽ%Type;
  v_�շѵ�     ������ü�¼.No%Type;
  n_��¼״̬   ������ü�¼.��¼״̬%Type;
  n_����id     ������ü�¼.����id%Type;
  n_�����id   ҽ�ƿ����.Id%Type;
  v_�˷ѽ���   Varchar2(1000);
  Err_Item Exception;

Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/GHDH'), Extractvalue(Value(A), 'IN/JSKLB'), Extractvalue(Value(A), 'IN/GHJE'),
         Extractvalue(Value(A), 'IN/LSH'), To_Number(Extractvalue(Value(A), 'IN/JCFP')),
         To_Number(Extractvalue(Value(A), 'IN/JKFS')), Extractvalue(Value(A), 'IN/YYFS')
  Into v_No, v_�����, n_�ҺŽ��, v_������ˮ��, n_��鷢Ʊ, n_�ɿʽ, v_ԤԼ��ʽ
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  Select Max(�շѵ�) Into v_�շѵ� From ���˹Һż�¼ Where NO = v_No;

  n_�ɿʽ := Nvl(n_�ɿʽ, 0);

  If n_�ɿʽ = 1 Then
    Begin
      Select 1 Into n_���� From ������ü�¼ Where NO = v_No And ��¼���� = 4 And ����id Is Not Null And Rownum < 2;
      Select 1
      Into n_����
      From ������ü�¼
      Where NO = v_�շѵ� And ��¼���� = 1 And ����id Is Not Null And Rownum < 2;
    Exception
      When Others Then
        n_���� := 0;
    End;
    If n_���� = 1 Then
      v_Err_Msg := '����ĹҺŵ��ݲ���ԤԼ�Һŵ�,�޷�ȡ��ԤԼ!';
      Raise Err_Item;
    End If;
    Begin
      Select 1 Into n_���� From ���˹Һż�¼ A Where a.No = v_No And a.ԤԼ��ʽ = v_ԤԼ��ʽ And Rownum < 2;
    Exception
      When Others Then
        n_���� := 0;
    End;
    If n_���� = 0 Then
      v_Err_Msg := '����ĹҺŵ��ݲ���' || v_ԤԼ��ʽ || 'ԤԼ��,�޷�ȡ��ԤԼ!';
      Raise Err_Item;
    End If;
  End If;

  If v_����� Is Not Null And n_�ɿʽ = 0 Then
    Select Nvl2(Translate(v_�����, '\1234567890', '\'), 'Char', 'Num') Into v_Type From Dual;
    If v_Type = 'Num' Then
      --������ǿ����ID
      Select ���㷽ʽ, ID Into v_���㷽ʽ, n_�����id From ҽ�ƿ���� Where ID = To_Number(v_�����);
    Else
      --������ǿ��������
      Select ���㷽ʽ, ID Into v_���㷽ʽ, n_�����id From ҽ�ƿ���� Where ���� = v_�����;
    End If;
  
    Select Sum(ʵ�ս��) Into n_ʵ�ս�� From ������ü�¼ Where NO = v_No And ��¼���� = 4;
  
    If Nvl(n_�ɿʽ, 0) = 0 Then
      --Ҫ�˵ĵ��ݲ����Ըý��㿨����ģ����ֹ�˺�
      Begin
        Select 1
        Into n_����
        From ����Ԥ����¼ A,
             (Select Distinct ����id
               From ������ü�¼
               Where NO = v_No And ��¼���� = 4
               Union
               Select Distinct ����id
               From סԺ���ü�¼
               Where NO = v_No And ��¼���� = 5
               Union
               Select Distinct ����id
               From ������ü�¼
               Where NO = v_�շѵ� And ��¼���� = 1) B
        Where a.����id = b.����id And ���㷽ʽ = v_���㷽ʽ And Rownum < 2;
      Exception
        When Others Then
          n_���� := 0;
      End;
      If n_���� = 0 Then
        v_Err_Msg := '����ĹҺŵ��ݲ���' || v_���㷽ʽ || '�����,�޷��˺�!';
        Raise Err_Item;
      End If;
    End If;
  End If;

  --��������飬�Ѵ��ڲ��������ݵģ������˺�
  Begin
    Select 1
    Into n_����
    From ���ò����¼ A,
         (Select Distinct ����id
           From ������ü�¼
           Where NO = v_No And ��¼���� = 4
           Union
           Select Distinct ����id
           From סԺ���ü�¼
           Where NO = v_No And ��¼���� = 5
           Union
           Select Distinct ����id
           From ������ü�¼
           Where NO = v_�շѵ� And ��¼���� = 1) B
    Where a.�շѽ���id = b.����id And a.��¼���� = 1 And a.���ӱ�־ = 1 And Nvl(a.����״̬, 0) <> 2 And Rownum < 2;
  Exception
    When Others Then
      n_���� := 0;
  End;
  If n_���� = 1 Then
    v_Err_Msg := '����ĹҺŵ����Ѿ������˶��ν���,�޷��˺�!';
    Raise Err_Item;
  End If;
  --ҽ����飬�Ѿ�����ҽ���ģ������˺�
  Begin
    Select Distinct 1 Into n_�ѿ�ҽ�� From ����ҽ����¼ Where �Һŵ� = v_No;
  Exception
    When Others Then
      n_�ѿ�ҽ�� := 0;
  End;
  If n_�ѿ�ҽ�� = 1 Then
    v_Err_Msg := '����ĹҺŵ����Ѿ�����ҽ��,�޷��˺�!';
    Raise Err_Item;
  End If;
  If Nvl(n_��鷢Ʊ, 0) = 1 Then
    Select Max(Decode(a.ʵ��Ʊ��, Null, 0, 1)) Into n_�Ƿ��ӡ From ������ü�¼ A Where NO = v_No And ��¼���� = 4;
    If Nvl(n_�Ƿ��ӡ, 0) = 1 Then
      v_Err_Msg := '�����˺ŵĵ����ѿ���Ʊ,�����˷�!';
      Raise Err_Item;
    End If;
    Select Max(Decode(a.ʵ��Ʊ��, Null, 0, 1))
    Into n_�Ƿ��ӡ
    From ������ü�¼ A
    Where NO = v_�շѵ� And ��¼���� = 1;
    If Nvl(n_�Ƿ��ӡ, 0) = 1 Then
      v_Err_Msg := '�����˺ŵĵ����ѿ���Ʊ,�����˷�!';
      Raise Err_Item;
    End If;
  End If;
  --��ȡ����Ա��Ϣ
  v_Temp := Zl_Identity(1);
  Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_Temp From Dual;
  Select Substr(v_Temp, 0, Instr(v_Temp, ',') - 1) Into v_����Ա��� From Dual;
  Select Substr(v_Temp, Instr(v_Temp, ',') + 1) Into v_����Ա���� From Dual;
  d_�Ǽ�ʱ�� := Sysdate;

  Zl_���������Һ�_Delete(v_No, v_������ˮ��, '�ƶ�ƽ̨�˺�', d_�Ǽ�ʱ��);

  --ͬ�������۵�
  If v_�շѵ� Is Not Null Then
    Select Max(��¼״̬), Max(����id) Into n_��¼״̬, n_����id From ������ü�¼ Where NO = v_�շѵ� And ��¼���� = 1;
    If n_��¼״̬ = 0 Then
      Zl_���ﻮ�ۼ�¼_Delete(v_�շѵ�);
    End If;
    If n_��¼״̬ = 1 Then
      If v_���㷽ʽ Is Null Then
        v_Err_Msg := '���ιҺŵ����˿�ʧ��,����!';
        Raise Err_Item;
      End If;
      Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
      Zl_�����շѼ�¼_����(v_�շѵ�, v_����Ա���, v_����Ա����, Null, d_�Ǽ�ʱ��, Null, n_����id);
    
      v_�˷ѽ��� := v_���㷽ʽ || '|' || -1 * n_�ҺŽ�� || '|' || ' |' || ' ';
      Zl_�����˷ѽ���_Modify(2, n_����id, n_����id, v_�˷ѽ���, 0, n_�����id, Null, v_������ˮ��, Null, 0, 0, 0, 2);
    End If;
  End If;

  If v_�շѵ� Is Null Then
    Select Max(����id) Into n_����id From ������ü�¼ Where NO = v_No And ��¼���� = 4 And ��¼״̬ = 3;
    Select Max(����id) Into n_����id From ������ü�¼ Where NO = v_No And ��¼���� = 4 And ��¼״̬ = 2;
  Else
    Select Max(����id) Into n_����id From ������ü�¼ Where NO = v_�շѵ� And ��¼���� = 1 And ��¼״̬ = 3;
    Select Max(����id) Into n_����id From ������ü�¼ Where NO = v_�շѵ� And ��¼���� = 1 And ��¼״̬ = 2;
  End If;

  v_Temp := '<CZSJ>' || To_Char(d_�Ǽ�ʱ��, 'YYYY-MM-DD hh24:mi:ss') || '</CZSJ>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<YJZID>' || n_����id || '</YJZID>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
  v_Temp := '<CXID>' || n_����id || '</CXID>';
  Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;

  Xml_Out := x_Templet;
Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Registdel;
/

--108251:������,2017-04-27,�������۵��ĹҺŵ�����
Create Or Replace Procedure Zl_Third_Getvisitinfo
(
  Xml_In  In Xmltype,
  Xml_Out Out Xmltype
) Is
  --------------------------------------------------------------------------------------------------
  --����:���ݹҺŵ��Ż�ȡ�ôξ�������(ҽ��Ϊ��Ҫ��ʾ)
  --���:Xml_In:
  --<IN>
  --    <GHDH>�Һŵ���</GHDH>
  --    <JSKLB>���㿨���</JSKLB>
  --    <MXGL>��ϸ����</MXGL> 0-������,��ϸ�������� 1-����,��ϸ����������,Ĭ��Ϊ1
  --</IN>
  --����:Xml_Out
  --<OUTPUT>
  --  <GH>
  --     <GHDH>�Һŵ���</GHDH> //���β�ѯ�ĹҺŵ���
  --     <YYSJ>ԤԼʱ��</YYSJ> //yyyy-mm-dd hh24:mi:ss
  --     <JZSJ></JZSJ>      //ʵ�ʾ���ʱ��
  --     <DJH></DJH>        //���ݺ�
  --     <JE></JE>          //���
  --     <DJLX></DJLX>      //��������,1-�շѵ���4-�Һŵ�
  --     <KDSJ></KDSJ>      //����ʱ��
  --     <JKFS></JKFS>      //�ɿʽ,0-�ҺŻ�ԤԼ�ɿ�;1-ԤԼ���ɿ�
  --     <ZFZT></ZFZT>  //֧��״̬,0-��֧����1-��֧����2-���˷�
  --     <SFJSK></SFJSK>    //�Ƿ���㿨֧����0-��1-��
  --  </GH>
  --  <YZLIST>
  --     <YZ>                   //ҽ��������HIS����ʾ��������ͬ
  --        <YZID><YZID>        //ҽ��ID��������ҽ��ID
  --        <YZLX><YZLX>        //ҽ������,�紦������顢����
  --         <YZMC></YZMC>        //ҽ������
  --        <ZXKS></ZXKS>       //ִ�п���
  --        <ZXKSID></ZXKSID>   //ִ�п���ID
  --        <FYCK></FYCK>       //��ҩ����
  --        <YZMX>
  --           <MX>
  --              <YZNR></YZNR>        //ҽ������
  --              <ZXZT></ZXZT>        //ҽ��ִ��״̬
  --              <SFFY>�Ƿ�ҩ</SFFY> // 0-�� ��1-��
  --              <GG>���</GG>
  --              <SL>����</SL>
  --              <DW>���㵥λ</DW>
  --              <BZDJ>��׼����</BZDJ>
  --              <YSJE>Ӧ�ս��</YSJE>
  --              <SSJE>ʵ�ս��</SSJE>
  --           </MX>
  --           <MX/>
  --        </YZMX>
  --        <BG></BG>                   //�Ƿ��ѳ����棬�Ƿ�ǩ��
  --        <BGLY></BGLY>               //�Ƿ������Ŀ,1-Ժ����Ŀ��2-�����Ŀ
  --        <BGLYSM></BGLYSM>           //�����Ŀ˵��
  --        <JZBG></JZBG>                //��ֹ��ʾ���档0-����1-��ֹ
  --        <JZTS></JZTS>                 //��ʾ���֡����ڽ�ֹ�鿴�ı��棬�ɷ���������ʾ���˵���Ϣ
  --        <BLID></BLID>              //����ID�����<BG>�ֶ�Ϊ1����ֵ��Ϊ��
  --        <DJLIST>
  --           <DJ>                //���õ�����Ϣ
  --              <DJH></DJH>      //���õ��ݺ�
  --              <DJLX></DJLX>    //��������
  --              <JE></JE>        //�����ܽ��
  --              <KDSJ></KDSJ>    //����ʱ��
  --              <ZFZT></ZFZT>    //֧��״̬,0-��֧����1-��֧����2-���˷�,3-�˷�������,4-���ͨ��,5-���δͨ��
  --              <SHSM></SHSM>    //���˵��,���δͨ��ԭ��
  --              <SFJSK></SFJSK>  //�Ƿ���㿨֧����0-��1-��
  --           </DJ>
  --           <DJ/>
  --        </DJLIST>
  --     </YZ>
  --  </YZLIST>
  --    <ERROR><MSG></MSG></ERROR>                      //������󷵻�
  --</OUTPUT>

  --------------------------------------------------------------------------------------------------
  v_Err_Msg Varchar2(200);
  Err_Item Exception;
  x_Templet Xmltype; --ģ��XML

  v_�����   Varchar2(100);
  n_�����id Number(18);
  v_�Һŵ�   Varchar2(10);
  v_�ŶӺ��� Varchar2(10);
  n_Temp     Number(18);
  v_�������� �ŶӽкŶ���.��������%Type;

  n_Count Number(18);

  v_Temp       Varchar2(32767); --��ʱXML
  v_����       Varchar2(32767);
  v_No         Varchar2(50);
  n_Add_Djlist Number(1); --�Ƿ�������DJLIST��
  n_����       Number(2);
  n_��ҽ��id   Number(18);
  n_����ҽ��   Number(8);
  n_ִ�п���id Number(18);
  v_ִ�п���   Varchar2(50);
  n_�˿���   ����Ԥ����¼.��Ԥ��%Type;
  n_��ϸ����   Number(3);
  n_�˷�״̬   �����˷�����.״̬%Type;
  v_����ԭ��   �����˷�����.����ԭ��%Type;
  v_���ԭ��   �����˷�����.���ԭ��%Type;
  v_��ҩ����   ������ü�¼.��ҩ����%Type;

Begin
  x_Templet := Xmltype('<OUTPUT></OUTPUT>');

  Select Extractvalue(Value(A), 'IN/GHDH'), Extractvalue(Value(A), 'IN/JSKLB'), Extractvalue(Value(A), 'IN/MXGL')
  Into v_�Һŵ�, v_�����, n_��ϸ����
  From Table(Xmlsequence(Extract(Xml_In, 'IN'))) A;

  If v_�Һŵ� Is Null Then
    v_Err_Msg := '�����ҵ�ָ���ĹҺŵ���(��ǰ�Һŵ���Ϊ��)';
    Raise Err_Item;
  End If;
  If n_��ϸ���� Is Null Then
    n_��ϸ���� := 1;
  End If;
  n_Add_Djlist := 0;

  v_Err_Msg := Null;
  If v_����� Is Not Null Then
    Begin
      n_�����id := To_Number(v_�����);
    Exception
      When Others Then
        n_�����id := 0;
    End;
  
    If n_�����id = 0 Then
      Begin
        Select ID, Decode(Nvl(�Ƿ�����, 0), 1, Null, ���� || 'δ����,��������нɷ�!')
        Into n_�����id, v_Err_Msg
        From ҽ�ƿ����
        Where ���� = v_�����;
      Exception
        When Others Then
          v_Err_Msg := '�����:' || v_����� || '������!';
      End;
    
    Else
    
      Begin
        Select ID, Decode(Nvl(�Ƿ�����, 0), 1, Null, ���� || 'δ����,��������нɷ�!')
        Into n_�����id, v_Err_Msg
        From ҽ�ƿ����
        Where ID = n_�����id;
      Exception
        When Others Then
          v_Err_Msg := 'δ�ҵ�ָ���Ľ���֧����Ϣ!';
      End;
    
    End If;
    If Not v_Err_Msg Is Null Then
      Raise Err_Item;
    End If;
  End If;
  n_���� := 4;
  --1.��ȡ�Һ�����
  Begin
    Select �շѵ� Into v_No From ���˹Һż�¼ Where NO = v_�Һŵ�;
  Exception
    When Others Then
      v_No := Null;
  End;

  If v_No Is Not Null Then
    Select Count(*) Into n_Count From ������ü�¼ Where NO = v_No And ��¼���� = 1;
    If n_Count <> 0 Then
      n_���� := 1;
    End If;
  End If;
  If n_���� = 4 Then
    v_No := v_�Һŵ�;
  End If;

  n_Count := 0;
  For c_�Һ� In (Select a.Id, v_No As NO, n_���� As ��¼����, a.ִ�в���id, c.���� As ִ�в���,
                      To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As �Ǽ�ʱ��, To_Char(a.ԤԼʱ��, 'yyyy-mm-dd hh24:mi:ss') As ԤԼʱ��,
                      a.����ʱ��, To_Char(a.����ʱ��, 'yyyy-mm-dd HH24:mi:ss') As ����ʱ��, a.�ű�, a.����, b.���, a.��¼״̬,
                      Decode(Nvl(a.ִ��״̬, 0), 0, '�ȴ�����', 1, '��ɾ���', 2, '���ھ���', -1, 'ȡ������') As ִ��״̬,
                      Decode(Nvl(b.����id, 0), 0, 0, 1) As ֧����־, Decode(Nvl(a.��¼����, 0), 2, 1, 0) As �ɿʽ, b.����id As ����id
               From ���˹Һż�¼ A,
                    (Select Max(Decode(��¼״̬, 0, 0, 2, 0, Nvl(����id, 0))) As ����id, Sum(ʵ�ս��) As ���
                      From ������ü�¼ B
                      Where ��¼���� = n_���� And NO = v_No) B, ���ű� C
               Where a.No = v_�Һŵ� And a.ִ�в���id = c.Id(+)) Loop
  
    If Nvl(c_�Һ�.��¼״̬, 0) <> 1 Then
      v_Err_Msg := '���ݺ�:' || v_�Һŵ� || '�Ѿ����˺�!';
      Raise Err_Item;
    End If;
  
    Begin
      Select �ŶӺ���, ��������
      Into v_�ŶӺ���, v_��������
      From �ŶӽкŶ���
      Where ҵ��id = c_�Һ�.Id And Nvl(ҵ������, 0) = 0;
    Exception
      When Others Then
        v_�ŶӺ��� := Null;
    End;
    If v_�ŶӺ��� Is Not Null Then
      --ҵ��id_In ,ҵ������_In �ŶӺ���_In Number := Null
      n_Temp := Zl_Getsequencebeforperons(c_�Һ�.Id, 0, v_�ŶӺ���, v_��������);
      v_���� := v_���� || '<DL><XH>' || v_�ŶӺ��� || '</XH><QMRS>' || n_Temp || '</QMRS></DL>';
    End If;
    n_Temp := 0;
    If Nvl(n_�����id, 0) <> 0 Then
      Begin
        Select 1
        Into n_Temp
        From ����Ԥ����¼
        Where ����id = c_�Һ�.����id And ��¼���� = 4 And ��¼״̬ In (1, 3) And �����id = n_�����id And Rownum < 2;
      Exception
        When Others Then
          Null;
      End;
    End If;
  
    v_Temp := '<GHDH>' || v_�Һŵ� || '</GHDH>';
    v_Temp := v_Temp || '<DJH>' || c_�Һ�.No || '</DJH>';
    v_Temp := v_Temp || '<YYSJ>' || c_�Һ�.ԤԼʱ�� || '</YYSJ>';
    v_Temp := v_Temp || '<JZSJ>' || c_�Һ�.����ʱ�� || '</JZSJ>';
    v_Temp := v_Temp || '<KDSJ>' || c_�Һ�.�Ǽ�ʱ�� || '</KDSJ>';
    v_Temp := v_Temp || '<JKFS>' || c_�Һ�.�ɿʽ || '</JKFS>';
    v_Temp := v_Temp || '<JE>' || c_�Һ�.��� || '</JE>';
    v_Temp := v_Temp || '<DJLX>' || n_���� || '</DJLX>';
    v_Temp := v_Temp || '<ZFZT>' || c_�Һ�.֧����־ || '</ZFZT>';
    v_Temp := v_Temp || '<SFJSK>' || n_Temp || '</SFJSK>';
    If v_���� Is Not Null Then
      v_Temp := v_Temp || v_����;
    End If;
    v_Temp := '<GH>' || v_Temp || '</GH>';
    Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype(v_Temp)) Into x_Templet From Dual;
    n_Count := n_Count + 1;
  End Loop;

  If Nvl(n_Count, 0) = 0 Then
    v_Err_Msg := 'δ�ҵ�ָ���ĹҺŵ���:' || v_�Һŵ� || '!';
    Raise Err_Item;
  End If;

  --2.�齨ҽ���������������
  n_��ҽ��id := 0;

  For c_ҽ�� In (With ҽ������ As
                  (Select ҽ��id, ���ͺ�, ��¼����, NO, Max(Nvl(ִ��״̬, 0)) As ִ��״̬
                  From (Select b.ҽ��id, b.���ͺ�, b.��¼����, b.No, Nvl(b.ִ��״̬, 0) As ִ��״̬
                         From ����ҽ����¼ A, ����ҽ������ B
                         Where a.�Һŵ� = v_�Һŵ� And a.Id = b.ҽ��id(+)
                         Union All
                         Select b.ҽ��id, b.���ͺ�, b.��¼����, b.No, Nvl(c.ִ��״̬, 0) As ִ��״̬
                         From ����ҽ����¼ A, ����ҽ������ B, ����ҽ������ C
                         Where a.�Һŵ� = v_�Һŵ� And a.Id = b.ҽ��id(+) And b.ҽ��id = c.ҽ��id(+) And b.���ͺ� = c.���ͺ�(+))
                  Group By ҽ��id, ���ͺ�, ��¼����, NO)
                 
                 Select Nvl(a.���id, a.Id) As ��id, Decode(a.���id, Null, 0, 1) As ��ҽ��, a.Id, a.���id, e.��ҩ����,
                        Max(Decode(a.�������, 'E', Decode(q.��������, '2', '����', '4', '����', '6', '����', m.����), m.����)) As ҽ������,
                        a.ִ�п���id, d.���� As ִ�п���, Decode(a.���id, Null, a.ҽ������, Null) As ��ҽ������,
                        Max(Decode(a.�������, '5', 1, '6', 1, '7', 1, 0) * Decode(Nvl(e.ִ��״̬, 0), 1, 1, 3, 1, 0)) As ��ҩ״̬,
                        Decode(a.���id, Null, Null, q.����) As ��ϸҽ������, s.���, (e.���� * e.����) As ����, e.���㵥λ As ��λ,
                        Decode(Nvl(b.ִ��״̬, 0), 0, 'δִ��', 1, '��ȫִ��', 2, '�ܾ�ִ��', 3, '����ִ��', '����ִ��') As ִ��״̬,
                        Max(Decode(p.���ʱ��, Null, Decode(C1.���ʱ��, Null, 0, 1), 1)) As �Ƿ��ѳ�����, c.����id, e.No, e.��¼���� As ��������,
                        Max(e.��׼����) As ��׼����, Sum(e.Ӧ�ս��) As Ӧ�ս��, Sum(e.ʵ�ս��) As ʵ�ս��,
                        To_Char(e.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, Decode(Nvl(e.��¼״̬, 0), 0, 0, 3, 2, 1) As ֧��״̬,
                        a.����id
                 
                 From ����ҽ����¼ A, ҽ������ B, ����ҽ������ C, ���Ӳ�����¼ C1, ���ű� D, ������ü�¼ E, ������Ŀ��� M, ������ĿĿ¼ Q, �շ���ĿĿ¼ S, ����걾��¼ P
                 Where a.Id = b.ҽ��id(+) And a.ִ�п���id = d.Id(+) And c.����id = C1.Id(+) And a.Id = c.ҽ��id(+) And
                       a.Id = p.ҽ��id(+) And b.ҽ��id = e.ҽ�����(+) And e.�շ�ϸĿid = s.Id(+) And b.No = e.No(+) And
                       b.��¼���� = e.��¼����(+) And e.��¼״̬(+) <> 2 And a.�Һŵ� = v_�Һŵ� And a.������� = m.����(+) And
                       a.������Ŀid = q.Id(+) And a.ҽ��״̬ In (3, 8)
                 Group By a.Id, a.Ӥ��, a.���, a.���id, e.��ҩ����, a.�������, a.ִ�п���id, d.����, a.ҽ������, q.����, s.���, e.���� * e.����,
                          e.���㵥λ, Decode(Nvl(b.ִ��״̬, 0), 0, 'δִ��', 1, '��ȫִ��', 2, '�ܾ�ִ��', 3, '����ִ��', '����ִ��'), C1.���ʱ��,
                          Decode(c.����id, Null, 0, 1), c.����id, e.No, e.��¼����, e.�Ǽ�ʱ��, Decode(Nvl(e.��¼״̬, 0), 0, 0, 3, 2, 1),
                          p.���ʱ��, a.����id
                 Order By ��id, ��ҽ��, Nvl(a.Ӥ��, 0), a.���) Loop
    If Nvl(n_Add_Djlist, 0) = 0 Then
      --����DJList�ڵ�
      Select Appendchildxml(x_Templet, '/OUTPUT', Xmltype('<YZLIST></YZLIST>')) Into x_Templet From Dual;
      n_Add_Djlist := 1;
    End If;
  
    If n_��ҽ��id <> Nvl(c_ҽ��.��id, 0) Then
      n_��ҽ��id := Nvl(c_ҽ��.��id, 0);
    
      Zl_Third_Custom_Getdeptinfo(n_��ҽ��id, n_ִ�п���id, v_ִ�п���);
    
      If Nvl(n_ִ�п���id, 0) = 0 Then
        If c_ҽ��.ҽ������ = '����' Then
          --����ҽ������ʾ�ɼ�����
          n_ִ�п���id := c_ҽ��.ִ�п���id;
          v_ִ�п���   := c_ҽ��.ִ�п���;
        Else
          Begin
            Select b.Id, b.����, c.��ҩ����
            Into n_ִ�п���id, v_ִ�п���, v_��ҩ����
            From ����ҽ����¼ A, ���ű� B, ������ü�¼ C
            Where a.Id = c.ҽ����� And a.���id = n_��ҽ��id And a.ִ�п���id = b.Id And Rownum <= 1;
          Exception
            When Others Then
              n_ִ�п���id := c_ҽ��.ִ�п���id;
              v_ִ�п���   := c_ҽ��.ִ�п���;
              v_��ҩ����   := c_ҽ��.��ҩ����;
          End;
        End If;
      End If;
    
      v_Temp := '<YZID>' || n_��ҽ��id || '</YZID>';
      v_Temp := v_Temp || '<YZLX>' || c_ҽ��.ҽ������ || '</YZLX>';
      v_Temp := v_Temp || '<YZMC>' || c_ҽ��.��ҽ������ || '</YZMC>';
      v_Temp := v_Temp || '<ZXKS>' || v_ִ�п��� || '</ZXKS>';
      v_Temp := v_Temp || '<ZXKSID>' || n_ִ�п���id || '</ZXKSID>';
      v_Temp := v_Temp || '<FYCK>' || v_��ҩ���� || '</FYCK>';
      v_Temp := v_Temp || '<BG>' || c_ҽ��.�Ƿ��ѳ����� || '</BG>';
      v_Temp := v_Temp || Zl_Third_Custom_Getrptfrom(n_��ҽ��id);
      v_Temp := v_Temp || Zl_Third_Custom_Rptlimit(c_ҽ��.����id, n_��ҽ��id);
      If Nvl(c_ҽ��.�Ƿ��ѳ�����, 0) = 1 And c_ҽ��.����id Is Not Null Then
        v_Temp := v_Temp || '<BLID>' || c_ҽ��.����id || '</BLID>';
      End If;
      v_Temp := '<YZ ҽ��ID="' || n_��ҽ��id || '">' || v_Temp || '<YZMX></YZMX><DJLIST></DJLIST></YZ>';
      Select Appendchildxml(x_Templet, '/OUTPUT/YZLIST', Xmltype(v_Temp)) Into x_Templet From Dual;
    
      For v_���� In (
                   
                   Select a.No, Mod(a.��¼����, 10) As ��������, To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��,
                           Max(Decode(Nvl(a.��¼״̬, 0), 0, 0, 3, 2, 1)) As ֧��״̬, Sum(a.ʵ�ս��) As ���ݽ��, Max(a.����id) As ���㿨֧��
                   From ������ü�¼ A
                   Where (a.No, a.��¼����) In
                         (Select Distinct q.No, q.��¼����
                          From ����ҽ����¼ M, ����ҽ������ Q
                          Where m.Id = q.ҽ��id(+) And (m.Id = n_��ҽ��id Or m.���id = n_��ҽ��id)
                          Union All
                          Select Distinct q.No, q.��¼����
                          From ����ҽ����¼ M, ����ҽ������ Q
                          Where m.Id = q.ҽ��id(+) And (m.Id = n_��ҽ��id Or m.���id = n_��ҽ��id)) And
                         Nvl(a.��¼״̬, 0) In (0, 1, 3)
                   Group By a.No, Mod(a.��¼����, 10), To_Char(a.�Ǽ�ʱ��, 'yyyy-mm-dd hh24:mi:ss')) Loop
        Begin
          Select 1
          Into n_Temp
          From ����Ԥ����¼ A, ������ü�¼ B
          Where a.����id = b.����id And b.No = v_����.No And Mod(b.��¼����, 10) = 1 And b.��¼״̬ In (1, 3) And a.�����id = n_�����id And
                Rownum < 2;
        Exception
          When Others Then
            n_Temp := 0;
        End;
        Begin
          Select -1 * Sum(���ʽ��)
          Into n_�˿���
          From ������ü�¼ B
          Where b.No = v_����.No And Mod(b.��¼����, 10) = 1 And b.��¼״̬ = 2;
        Exception
          When Others Then
            n_�˿��� := 0;
        End;
        Begin
          Select ״̬, ����ԭ��, ���ԭ��
          Into n_�˷�״̬, v_����ԭ��, v_���ԭ��
          From �����˷�����
          Where NO = v_����.No And Mod(��¼����, 10) = Mod(v_����.��������, 10);
        Exception
          When Others Then
            n_�˷�״̬ := -1;
            v_����ԭ�� := '';
            v_���ԭ�� := '';
        End;
      
        v_Temp := '<DJH>' || v_����.No || '</DJH>';
        v_Temp := v_Temp || '<DJLX>' || v_����.�������� || '</DJLX>';
        v_Temp := v_Temp || '<JE>' || v_����.���ݽ�� || '</JE>';
        v_Temp := v_Temp || '<KDSJ>' || v_����.����ʱ�� || '</KDSJ>';
        If n_�˷�״̬ = -1 Then
          v_Temp := v_Temp || '<ZFZT>' || v_����.֧��״̬ || '</ZFZT>';
        Else
          If n_�˷�״̬ = 0 Then
            v_Temp := v_Temp || '<ZFZT>3</ZFZT>';
          End If;
          If n_�˷�״̬ = 1 Then
            If v_����.֧��״̬ = 2 Then
              v_Temp := v_Temp || '<ZFZT>2</ZFZT>';
            Else
              v_Temp := v_Temp || '<ZFZT>4</ZFZT>';
            End If;
          End If;
          If n_�˷�״̬ = 2 Then
            v_Temp := v_Temp || '<ZFZT>5</ZFZT>';
          End If;
        End If;
      
        If n_�˷�״̬ = -1 Then
          v_Temp := v_Temp || '<SHSM>' || '' || '</SHSM>';
        Else
          If n_�˷�״̬ = 0 Then
            v_Temp := v_Temp || '<SHSM>' || v_����ԭ�� || '</SHSM>';
          End If;
          If n_�˷�״̬ = 1 Then
            v_Temp := v_Temp || '<SHSM>' || v_���ԭ�� || '</SHSM>';
          End If;
          If n_�˷�״̬ = 2 Then
            v_Temp := v_Temp || '<SHSM>' || v_���ԭ�� || '</SHSM>';
          End If;
        End If;
      
        v_Temp := v_Temp || '<YTJE>' || Nvl(n_�˿���, 0) || '</YTJE>';
        v_Temp := v_Temp || '<SFJSK>' || n_Temp || '</SFJSK>';
        v_Temp := '<DJ>' || v_Temp || '</DJ>';
        Select Appendchildxml(x_Templet, '/OUTPUT/YZLIST/YZ[@ҽ��ID="' || n_��ҽ��id || '"]/DJLIST', Xmltype(v_Temp))
        Into x_Templet
        From Dual;
      End Loop;
    End If;
  
    --ֻ��һ����¼��ҽ��������ϸ�����Ӹ���ҽ�����Ի�ȡִ��״̬
    Select Decode(Count(*), 0, 1, 0) Into n_����ҽ�� From ����ҽ����¼ Where ���id = n_��ҽ��id;
    If n_����ҽ�� = 1 Then
      v_Temp := '<YZNR>' || c_ҽ��.��ҽ������ || '</YZNR>';
      v_Temp := v_Temp || '<GG>' || c_ҽ��.��� || '</GG>';
      v_Temp := v_Temp || '<SFFY>' || c_ҽ��.��ҩ״̬ || '</SFFY>';
      v_Temp := v_Temp || '<SL>' || c_ҽ��.���� || '</SL>';
      v_Temp := v_Temp || '<DW>' || c_ҽ��.��λ || '</DW>';
      v_Temp := v_Temp || '<BZDJ>' || Nvl(c_ҽ��.��׼����, 0) || '</BZDJ>';
      v_Temp := v_Temp || '<YSJE>' || Nvl(c_ҽ��.Ӧ�ս��, 0) || '</YSJE>';
      v_Temp := v_Temp || '<SSJE>' || Nvl(c_ҽ��.ʵ�ս��, 0) || '</SSJE>';
      v_Temp := v_Temp || '<ZXZT>' || c_ҽ��.ִ��״̬ || '</ZXZT>';
      v_Temp := '<MX>' || v_Temp || '</MX>';
      Select Appendchildxml(x_Templet, '/OUTPUT/YZLIST/YZ[@ҽ��ID="' || n_��ҽ��id || '"]/YZMX', Xmltype(v_Temp))
      Into x_Templet
      From Dual;
    End If;
  
    If Nvl(c_ҽ��.��ҽ��, 0) = 1 Then
      If n_��ϸ���� = 0 Or (n_��ϸ���� = 1 And c_ҽ��.ҽ������ <> '����') Then
        v_Temp := '<YZNR>' || c_ҽ��.��ϸҽ������ || '</YZNR>';
        v_Temp := v_Temp || '<GG>' || c_ҽ��.��� || '</GG>';
        v_Temp := v_Temp || '<SL>' || c_ҽ��.���� || '</SL>';
        v_Temp := v_Temp || '<DW>' || c_ҽ��.��λ || '</DW>';
        v_Temp := v_Temp || '<SFFY>' || c_ҽ��.��ҩ״̬ || '</SFFY>';
        v_Temp := v_Temp || '<ZXZT>' || c_ҽ��.ִ��״̬ || '</ZXZT>';
        v_Temp := v_Temp || '<BZDJ>' || Nvl(c_ҽ��.��׼����, 0) || '</BZDJ>';
        v_Temp := v_Temp || '<YSJE>' || Nvl(c_ҽ��.Ӧ�ս��, 0) || '</YSJE>';
        v_Temp := v_Temp || '<SSJE>' || Nvl(c_ҽ��.ʵ�ս��, 0) || '</SSJE>';
        v_Temp := '<MX>' || v_Temp || '</MX>';
        Select Appendchildxml(x_Templet, '/OUTPUT/YZLIST/YZ[@ҽ��ID="' || n_��ҽ��id || '"]/YZMX', Xmltype(v_Temp))
        Into x_Templet
        From Dual;
      End If;
    End If;
  
  End Loop;
  Xml_Out := x_Templet;

Exception
  When Err_Item Then
    v_Temp := '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]';
    Raise_Application_Error(-20101, v_Temp);
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Third_Getvisitinfo;
/

--108594:������,2017-04-25,����תסԺ���ŵ����˷�����
Create Or Replace Procedure Zl_����תסԺ_�շ�ת��
(
  No_In         סԺ���ü�¼.No%Type,
  ����Ա���_In סԺ���ü�¼.����Ա���%Type,
  ����Ա����_In סԺ���ü�¼.����Ա����%Type,
  �˷�ʱ��_In   סԺ���ü�¼.����ʱ��%Type,
  �����˷�_In   Number := 0,
  ��Ժ����id_In סԺ���ü�¼.��������id%Type := Null,
  ��ҳid_In     סԺ���ü�¼.��ҳid%Type := Null,
  ���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type := Null,
  ����id_In     ����Ԥ����¼.����id%Type := Null,
  ԭ����id_In   ����Ԥ����¼.����id%Type := Null,
  ����_In     ����Ԥ����¼.��Ԥ��%Type := Null
) As
  --�����˷�_In:0-����תסԺ��������;1-�����˷�ģʽ
  -- �����˷�_InΪ1ʱ:��Ժ����id_In����ҳID_IN���Բ�����
  n_Count      Number(5);
  n_ԭ����id   סԺ���ü�¼.����id%Type;
  n_ʵ�ս��   ������ü�¼.ʵ�ս��%Type;
  n_Ԥ��ʹ�ö� ����Ԥ����¼.��Ԥ��%Type;
  n_ʵ�ʳ���   ����Ԥ����¼.��Ԥ��%Type;
  n_��id       ����ɿ����.Id%Type;
  n_����id     ������Ϣ.����id%Type;
  v_Ԥ��no     ����Ԥ����¼.No%Type;
  n_Ԥ�����   ����Ԥ����¼.��Ԥ��%Type;
  n_��ӡid     Ʊ��ʹ����ϸ.��ӡid%Type;
  n_��������id סԺ���ü�¼.��������id%Type;
  v_������     ������ü�¼.������%Type;
  n_����id     ������ü�¼.����id%Type;
  n_����     ����Ԥ����¼.��Ԥ��%Type;
  v_����     ���㷽ʽ.����%Type;
  n_����ֵ     �������.�������%Type;
  v_���㷽ʽ   ���㷽ʽ.����%Type;
  v_Nos        Varchar2(3000);
  v_����ids    Varchar2(3000);
  v_ԭ����ids  Varchar2(3000);
  n_Tempid     ����Ԥ����¼.Id%Type;
  n_Ԥ��id     ����Ԥ����¼.Id%Type;
  n_ҽ��       Number;
  n_����       Number;
  n_����       Number;
  n_�����˷�   Number;
  n_�˷�����   Number;
  n_�쳣��־   Number;
  n_�������   Number;
  n_����״̬   ������ü�¼.����״̬%Type;

  Err_Item Exception;
  v_Err_Msg Varchar2(200);

  Procedure Zl_Square_Update
  (
    ����ids_In    Varchar2,
    �ֽ���id_In   ����Ԥ����¼.����id%Type,
    �ɿ���id_In   ����Ԥ����¼.�ɿ���id%Type,
    �˿�ʱ��_In   ����Ԥ����¼.�տ�ʱ��%Type,
    �������_In   ����Ԥ����¼.�������%Type,
    ��������_In   Varchar2 := Null,
    �˷ѽ��_In   ����Ԥ����¼.��Ԥ��%Type := Null,
    ���㿨���_In ����Ԥ����¼.���㿨���%Type := Null
  ) As
    n_��¼״̬ ���˿������¼.��¼״̬%Type;
    n_Ԥ��id   ����Ԥ����¼.Id%Type;
    v_����     ���˿������¼.����%Type;
    n_���ڿ�Ƭ Number;
    d_ͣ������ ���ѿ�Ŀ¼.ͣ������%Type;
    n_������ ���˿������¼.���%Type;
    n_���     ���˿������¼.���%Type;
    n_���     ���ѿ�Ŀ¼.���%Type;
    n_�ӿڱ�� ���˿������¼.�ӿڱ��%Type;
    d_����ʱ�� ���ѿ�Ŀ¼.����ʱ��%Type;
    n_Id       ����Ԥ����¼.Id%Type;
  Begin
    n_Ԥ��id := 0;
  
    --�������ѿ�,���㿨��������Ѿ�������
    For v_У�� In (Select Min(a.Id) As Ԥ��id, c.���ѿ�id, Sum(c.������) As ������, c.�ӿڱ��, c.����, Max(c.���) As ���, Max(c.Id) As ID
                 From ����Ԥ����¼ A, ���˿�������� B, ���˿������¼ C
                 Where a.Id = b.Ԥ��id And a.���㿨��� = ���㿨���_In And b.������id = c.Id And a.��¼���� = 3 And
                       Instr(Nvl(��������_In, '_LXH'), ',' || a.���㷽ʽ || ',') = 0 And
                       a.����id In (Select Column_Value From Table(f_Str2list(����ids_In)))
                 Group By c.���ѿ�id, c.�ӿڱ��, c.����) Loop
    
      If Nvl(v_У��.���ѿ�id, 0) <> 0 Then
        Select Max(��¼״̬)
        Into n_��¼״̬
        From ���˿������¼
        Where �ӿڱ�� = v_У��.�ӿڱ�� And ���ѿ�id = Nvl(v_У��.���ѿ�id, 0) And ���� = v_У��.���� And Nvl(���, 0) = Nvl(v_У��.���, 0);
      Else
        Select Max(��¼״̬)
        Into n_��¼״̬
        From ���˿������¼
        Where �ӿڱ�� = v_У��.�ӿڱ�� And ���ѿ�id Is Null And ���� = v_У��.���� And Nvl(���, 0) = Nvl(v_У��.���, 0);
      End If;
    
      If n_��¼״̬ = 1 Then
        n_��¼״̬ := 2;
      Else
        n_��¼״̬ := n_��¼״̬ + 2;
      End If;
      --����ʱ,ֻ����һ��
      If n_Ԥ��id = 0 Then
        Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      
        Insert Into ����Ԥ����¼
          (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id,
           Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, У�Ա�־, �������, ��������)
          Select n_Ԥ��id, ��¼����, NO, 2, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �˿�ʱ��_In, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���_In, ����Ա����_In,
                 -1 * �˷ѽ��_In, �ֽ���id_In, �ɿ���id_In, Ԥ�����, �����id, Nvl(���㿨���, v_У��.�ӿڱ��), ����, ������ˮ��, ����˵��, ������λ, 2, �������_In,
                 ��������
          From ����Ԥ����¼ A
          Where ID = v_У��.Ԥ��id;
      End If;
    
      If Nvl(v_У��.���ѿ�id, 0) <> 0 Then
        --���ѿ�,ֱ���˻ؿ�������
        Begin
          Select ����, 1, ͣ������, (Select Max(���) From ���ѿ�Ŀ¼ B Where a.���� = b.���� And a.�ӿڱ�� = b.�ӿڱ��), ���, ���, �ӿڱ��, ����ʱ��
          Into v_����, n_���ڿ�Ƭ, d_ͣ������, n_������, n_���, n_���, n_�ӿڱ��, d_����ʱ��
          From ���ѿ�Ŀ¼ A
          Where ID = v_У��.���ѿ�id;
        Exception
          When Others Then
            n_���ڿ�Ƭ := 0;
        End;
      
        --ȡ��ͣ��
        If n_���ڿ�Ƭ = 0 Then
          v_Err_Msg := '����Ϊ"' || v_���� || '"�����ѿ�������ɾ�������������øÿ�Ƭ,���飡';
          Raise Err_Item;
        End If;
        If Nvl(n_���, 0) < Nvl(n_������, 0) Then
          v_Err_Msg := '����������ʷ������¼(����Ϊ"' || v_���� || '"),���飡';
          Raise Err_Item;
        End If;
        If Nvl(d_ͣ������, To_Date('3000-01-01', 'yyyy-mm-dd')) < To_Date('3000-01-01', 'yyyy-mm-dd') Then
          v_Err_Msg := '����Ϊ"' || v_���� || '"�����ѿ��Ѿ�������ͣ�ã������ٽ����˷�,���飡';
          Raise Err_Item;
        End If;
      
        If d_����ʱ�� < To_Date('3000-01-01', 'yyyy-mm-dd') Then
          v_Err_Msg := '����Ϊ"' || v_���� || '"�����ѿ��Ѿ����գ������˷�,���飡';
          Raise Err_Item;
        End If;
        Update ���ѿ�Ŀ¼ Set ��� = Nvl(���, 0) + �˷ѽ��_In Where ID = Nvl(v_У��.���ѿ�id, 0);
      End If;
    
      Select ���˿������¼_Id.Nextval Into n_Id From Dual;
      Insert Into ���˿������¼
        (ID, �ӿڱ��, ���ѿ�id, ���, ��¼״̬, ���㷽ʽ, ������, ����, ������ˮ��, ����ʱ��, ��ע, �����־)
        Select n_Id, �ӿڱ��, ���ѿ�id, ���, n_��¼״̬, ���㷽ʽ, -1 * �˷ѽ��_In, ����, ������ˮ��, ����ʱ��, ��ע,
               Decode(���ѿ�id, Null, 0, 0, 0, 1) As ��־
        From ���˿������¼
        Where ID = v_У��.Id;
      Insert Into ���˿�������� (Ԥ��id, ������id) Values (n_Ԥ��id, n_Id);
    
      If n_��¼״̬ <> 2 And n_��¼״̬ <> 1 Then
        Update ���˿������¼ Set ��¼״̬ = 3 Where ID = v_У��.Id;
      End If;
    End Loop;
  End;
Begin
  n_��id := Zl_Get��id(����Ա����_In);
  --����
  Begin
    Select ���� Into v_���� From ���㷽ʽ Where ���� = 9 And Rownum < 2;
  Exception
    When Others Then
      v_Err_Msg := 'û�з��������㷽ʽ�������Ƿ���ȷ���ã�';
      Raise Err_Item;
  End;

  If ԭ����id_In Is Null Then
  
    Select Count(NO), Sum(ʵ�ս��) Into n_Count, n_ʵ�ս�� From ������ü�¼ Where NO = No_In And Mod(��¼����,10) = 1;
    If n_Count = 0 Or n_ʵ�ս�� = 0 Then
      v_Err_Msg := '����' || No_In || '�����շѵ��ݻ��򲢷�ԭ�����˲����˸õ���,����תΪסԺ����.';
      Raise Err_Item;
    End If;
  
    Select ����id, ����id, ��������id, ������
    Into n_ԭ����id, n_����id, n_��������id, v_������
    From ������ü�¼
    Where NO = No_In And Mod(��¼����,10) = 1 And ��¼״̬ In (1, 3) And Rownum < 2;
  
    --1.1���Ϸ��ü�¼
    If ����id_In Is Null Then
      Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
    Else
      n_����id := ����id_In;
    End If;
  
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�, ���˿���id, �շ����, �շ�ϸĿid,
       ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ������, ִ����, ִ��״̬, ִ��ʱ��,
       ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id, ����״̬)
      Select ���˷��ü�¼_Id.Nextval, NO, ʵ��Ʊ��, ��¼����, 2, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�, ���˿���id,
             �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, -1 * ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, -1 * Ӧ�ս��, -1 * ʵ�ս��, ��������id,
             ������, ִ�в���id, ������, ִ����, -1, ִ��ʱ��, ����Ա���_In, ����Ա����_In, ����ʱ��, �˷�ʱ��_In, n_����id, -1 * ���ʽ��, ������Ŀ��, ���մ���id, ͳ����,
             ժҪ, Decode(Nvl(���ӱ�־, 0), 9, 1, 0), ���ձ���, ��������, n_��id, 0
      From ������ü�¼
      Where NO = No_In And Mod(��¼����,10) = 1 And ��¼״̬ = 1;
  
    --Update ������ü�¼ Set ��¼״̬ = 3 Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 1;
  
    --1.2����Ԥ����¼
    --���ϳ�Ԥ������
    For r_����id In (Select Distinct ����id
                   From ������ü�¼
                   Where NO In (Select Distinct NO
                                From ������ü�¼
                                Where ����id In (Select ����id
                                               From ����Ԥ����¼
                                               Where ������� In (Select b.�������
                                                              From ������ü�¼ A, ����Ԥ����¼ B
                                                              Where a.No = No_In And b.������� < 0 And Mod(a.��¼����, 10) = 1 And
                                                                    a.��¼״̬ <> 0 And a.����id = b.����id))) And
                         Mod(��¼����, 10) = 1 And ��¼״̬ <> 0
                   Union
                   Select Distinct ����id
                   From ������ü�¼
                   Where NO In (Select Distinct NO
                                From ������ü�¼
                                Where ����id In (Select a.����id
                                               From ������ü�¼ A, ����Ԥ����¼ B
                                               Where a.No = No_In And b.������� > 0 And Mod(a.��¼����, 10) = 1 And a.��¼״̬ <> 0 And
                                                     a.����id = b.����id)) And Mod(��¼����, 10) = 1 And ��¼״̬ <> 0) Loop
      v_ԭ����ids := v_ԭ����ids || ',' || r_����id.����id;
    End Loop;
    v_ԭ����ids := Substr(v_ԭ����ids, 2);
  
    Begin
      Select 1
      Into n_ҽ��
      From ���ս����¼
      Where ��¼id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And Rownum < 2;
    Exception
      When Others Then
        n_ҽ�� := 0;
    End;
  
    If n_ҽ�� = 1 Then
      Begin
        Select 1
        Into n_����
        From ҽ��������ϸ
        Where NO = No_In And ����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And Rownum < 2;
      Exception
        When Others Then
          v_Err_Msg := '��ǰ����' || No_In || '������ҽ��������ϸ,�޷���������תסԺ!';
          Raise Err_Item;
      End;
    End If;
  
    --ҽ���˿�
    For r_ҽ�� In (Select ����id, NO, ���㷽ʽ, ���, ��ע
                 From ҽ��������ϸ
                 Where NO = No_In And ����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids)))) Loop
      Update ��Ա�ɿ����
      Set ��� = Nvl(���, 0) - r_ҽ��.���
      Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = r_ҽ��.���㷽ʽ
      Returning ��� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ��Ա�ɿ����
          (�տ�Ա, ���㷽ʽ, ����, ���)
        Values
          (����Ա����_In, r_ҽ��.���㷽ʽ, 1, -1 * r_ҽ��.���);
        n_����ֵ := r_ҽ��.���;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From ��Ա�ɿ����
        Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = r_ҽ��.���㷽ʽ And Nvl(���, 0) = 0;
      End If;
    
      Update ����Ԥ����¼
      Set ��Ԥ�� = ��Ԥ�� + (-1 * r_ҽ��.���)
      Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = r_ҽ��.���㷽ʽ;
      If Sql%RowCount = 0 Then
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ,
           �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, У�Ա�־, ��������)
        Values
          (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * r_ҽ��.���, r_ҽ��.���㷽ʽ, Null, �˷�ʱ��_In,
           Null, Null, Null, ����Ա���_In, ����Ա����_In, r_ҽ��.��ע, n_��id, Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id,
           0, 3);
      End If;
    
      Update ����Ԥ����¼
      Set ��¼״̬ = 3
      Where ��¼���� = 3 And ��¼״̬ = 1 And ����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And
            ���㷽ʽ = r_ҽ��.���㷽ʽ;
    
      Update ҽ��������ϸ
      Set ��� = ��� + (-1 * r_ҽ��.���)
      Where NO = No_In And ����id = n_����id And ���㷽ʽ = r_ҽ��.���㷽ʽ;
      If Sql%RowCount = 0 Then
        Insert Into ҽ��������ϸ
          (����id, NO, ���㷽ʽ, ���)
        Values
          (n_����id, No_In, r_ҽ��.���㷽ʽ, -1 * r_ҽ��.���);
      End If;
      n_ʵ�ս�� := n_ʵ�ս�� - r_ҽ��.���;
    End Loop;
  
    Begin
      Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And ���� Like '%�ֽ�%' And Rownum < 2;
    Exception
      When Others Then
        Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And Rownum < 2;
    End;
  
    If n_ʵ�ս�� <> 0 Then
      For r_Prepay In (Select NO, ʵ��Ʊ��, ����id, ��ҳid, ����id, ���㷽ʽ, �������, �ɿλ, ��λ������, ��λ�ʺ�, Sum(��Ԥ��) As ��Ԥ��, �����id, ���㿨���,
                              ����, ������ˮ��, ����˵��, ������λ
                       From ����Ԥ����¼ A
                       Where ��¼���� In (1, 11) And a.����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids)))
                       Group By n_Tempid, NO, ʵ��Ʊ��, ����id, ��ҳid, ����id, ���㷽ʽ, �������, �ɿλ, ��λ������, ��λ�ʺ�, �����id, ���㿨���, ����,
                                ������ˮ��, ����˵��, ������λ) Loop
        If n_ʵ�ս�� <> 0 Then
          If r_Prepay.��Ԥ�� >= n_ʵ�ս�� Then
            Select ����Ԥ����¼_Id.Nextval Into n_Tempid From Dual;
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���,
               ��Ԥ��, ����id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, Ԥ�����, �������, �ɿ���id)
              Select n_Tempid, r_Prepay.No, r_Prepay.ʵ��Ʊ��, 11, 1, r_Prepay.����id, r_Prepay.��ҳid, r_Prepay.����id, Null,
                     r_Prepay.���㷽ʽ, r_Prepay.�������, Null, r_Prepay.�ɿλ, r_Prepay.��λ������, r_Prepay.��λ�ʺ�, �˷�ʱ��_In, ����Ա����_In,
                     ����Ա���_In, -1 * n_ʵ�ս��, n_����id, r_Prepay.�����id, r_Prepay.���㿨���, r_Prepay.����, r_Prepay.������ˮ��,
                     r_Prepay.����˵��, r_Prepay.������λ, 1, -1 * n_����id, n_��id
              From Dual;
            Update �������
            Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(n_ʵ�ս��, 0)
            Where ����id = n_����id And ���� = 1 And ���� = 1
            Returning Ԥ����� Into n_����ֵ;
            If Sql%RowCount = 0 Then
              Insert Into ������� (����id, ����, Ԥ�����, ����) Values (n_����id, 1, n_ʵ�ս��, 1);
              n_����ֵ := n_ʵ�ս��;
            End If;
            If Nvl(n_����ֵ, 0) = 0 Then
              Delete From �������
              Where ����id = n_����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
            End If;
            n_ʵ�ս�� := 0;
          Else
            Select ����Ԥ����¼_Id.Nextval Into n_Tempid From Dual;
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���,
               ��Ԥ��, ����id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, Ԥ�����, �������, �ɿ���id)
              Select n_Tempid, r_Prepay.No, r_Prepay.ʵ��Ʊ��, 11, 1, r_Prepay.����id, r_Prepay.��ҳid, r_Prepay.����id, Null,
                     r_Prepay.���㷽ʽ, r_Prepay.�������, Null, r_Prepay.�ɿλ, r_Prepay.��λ������, r_Prepay.��λ�ʺ�, �˷�ʱ��_In, ����Ա����_In,
                     ����Ա���_In, -1 * r_Prepay.��Ԥ��, n_����id, r_Prepay.�����id, r_Prepay.���㿨���, r_Prepay.����, r_Prepay.������ˮ��,
                     r_Prepay.����˵��, r_Prepay.������λ, 1, -1 * n_����id, n_��id
              From Dual;
            Update �������
            Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(r_Prepay.��Ԥ��, 0)
            Where ����id = n_����id And ���� = 1 And ���� = 1
            Returning Ԥ����� Into n_����ֵ;
            If Sql%RowCount = 0 Then
              Insert Into ������� (����id, ����, Ԥ�����, ����) Values (n_����id, 1, r_Prepay.��Ԥ��, 1);
              n_����ֵ := r_Prepay.��Ԥ��;
            End If;
            If Nvl(n_����ֵ, 0) = 0 Then
              Delete From �������
              Where ����id = n_����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
            End If;
            n_ʵ�ս�� := n_ʵ�ս�� - r_Prepay.��Ԥ��;
          End If;
        End If;
      End Loop;
    End If;
    --2.Ʊ���ջ�
    --������ǰû�д�ӡ,���ջ�
    Select Nvl(Max(ID), 0)
    Into n_��ӡid
    From (Select b.Id
           From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
           Where a.��ӡid = b.Id And a.���� = 1 And a.ԭ�� In (1, 3) And b.�������� = 1 And b.No = No_In
           Order By a.ʹ��ʱ�� Desc)
    Where Rownum < 2;
    If n_��ӡid > 0 Then
      --���ŵ���ѭ������ʱֻ���ջ�һ��
      Select Count(��ӡid) Into n_Count From Ʊ��ʹ����ϸ Where Ʊ�� = 1 And ���� = 2 And ��ӡid = n_��ӡid;
      If n_Count = 0 Then
        Insert Into Ʊ��ʹ����ϸ
          (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����)
          Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, �˷�ʱ��_In, ����Ա����_In
          From Ʊ��ʹ����ϸ
          Where ��ӡid = n_��ӡid And Ʊ�� = 1 And ���� = 1;
      End If;
    End If;
  
    --3.�ɿ����ݴ���(
    --   �����������:
    --    1. ת������ֱ�����ʵ�,��ɿ����ݲ�����;
    --    2. ��ת��,�ٵ������˿���Ʊ,����Ҫ���нɿ����ݴ���
    If Nvl(�����˷�_In, 0) = 1 Then
      For c_Ԥ�� In (Select a.���㷽ʽ, Sum(a.��Ԥ��) As ��Ԥ��, 2 As Ԥ�����, a.�����id, a.���㿨���, a.����, Min(a.������ˮ��) As ������ˮ��,
                          Min(a.����˵��) As ����˵��, Min(a.������λ) As ������λ, b.����
                   From ����Ԥ����¼ A, ���㷽ʽ B
                   Where a.��¼���� = 3 And a.����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And
                         a.���㷽ʽ = b.���� And b.���� In (1, 2, 7, 8) And a.���㷽ʽ Is Not Null
                   Group By a.���㷽ʽ, Ԥ�����, a.�����id, a.���㿨���, a.����, b.����
                   Having Sum(a.��Ԥ��) <> 0
                   Order By a.�����id, ���� Desc) Loop
        If n_ʵ�ս�� <> 0 Then
          Begin
            Select �Ƿ����� Into n_���� From ҽ�ƿ���� Where ID = c_Ԥ��.�����id;
          Exception
            When Others Then
              n_���� := 0;
          End;
          If (c_Ԥ��.���� = 7 Or (c_Ԥ��.���� = 8 And c_Ԥ��.�����id Is Not Null)) And n_���� = 0 Then
            If c_Ԥ��.��Ԥ�� > n_ʵ�ս�� Then
              Update ����Ԥ����¼
              Set ��Ԥ�� = ��Ԥ�� + (-1 * n_ʵ�ս��), ժҪ = ժҪ || '1' || ',' || c_Ԥ��.�����id || ',' || -1 * n_ʵ�ս�� || '|'
              Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ Is Null;
              If Sql%RowCount = 0 Then
                Insert Into ����Ԥ����¼
                  (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                   ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��������, У�Ա�־)
                Values
                  (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * n_ʵ�ս��, Null, Null, �˷�ʱ��_In,
                   Null, Null, Null, ����Ա���_In, ����Ա����_In, '1' || ',' || c_Ԥ��.�����id || ',' || -1 * n_ʵ�ս�� || '|', n_��id,
                   Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 3, 1);
              End If;
              Update ����Ԥ����¼
              Set ��¼״̬ = 3
              Where ��¼���� = 3 And ��¼״̬ = 1 And ����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And
                    ���㷽ʽ = c_Ԥ��.���㷽ʽ;
              n_����״̬ := 1;
              n_ʵ�ս�� := 0;
            Else
              Update ����Ԥ����¼
              Set ��Ԥ�� = ��Ԥ�� + (-1 * c_Ԥ��.��Ԥ��), ժҪ = ժҪ || '1' || ',' || c_Ԥ��.�����id || ',' || -1 * c_Ԥ��.��Ԥ�� || '|'
              Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ Is Null;
              If Sql%RowCount = 0 Then
                Insert Into ����Ԥ����¼
                  (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                   ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��������, У�Ա�־)
                Values
                  (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * c_Ԥ��.��Ԥ��, Null, Null, �˷�ʱ��_In,
                   Null, Null, Null, ����Ա���_In, ����Ա����_In, '1' || ',' || c_Ԥ��.�����id || ',' || -1 * c_Ԥ��.��Ԥ�� || '|', n_��id,
                   Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 3, 1);
              End If;
            
              Update ����Ԥ����¼
              Set ��¼״̬ = 3
              Where ��¼���� = 3 And ��¼״̬ = 1 And ����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And
                    ���㷽ʽ = c_Ԥ��.���㷽ʽ;
              n_����״̬ := 1;
              n_ʵ�ս�� := n_ʵ�ս�� - c_Ԥ��.��Ԥ��;
            End If;
          Else
            n_ʵ�ʳ��� := 0;
            If c_Ԥ��.���� In (3, 4) Or (c_Ԥ��.���� = 8 And c_Ԥ��.���㿨��� Is Not Null) Then
              v_���㷽ʽ := c_Ԥ��.���㷽ʽ;
            Else
              If ���㷽ʽ_In Is Null Then
                Begin
                  Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And ���� Like '%�ֽ�%' And Rownum < 2;
                Exception
                  When Others Then
                    Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And Rownum < 2;
                End;
              Else
                v_���㷽ʽ := ���㷽ʽ_In;
              End If;
            End If;
          
            If c_Ԥ��.���� = 8 And c_Ԥ��.���㿨��� Is Not Null Then
              If n_ʵ�ս�� >= c_Ԥ��.��Ԥ�� Then
                --Zl_Square_Update(v_ԭ����ids, n_����id, n_��id, �˷�ʱ��_In, -1 * n_����id, Null, c_Ԥ��.��Ԥ��, c_Ԥ��.���㿨���);
                Update ����Ԥ����¼
                Set ��Ԥ�� = ��Ԥ�� + (-1 * c_Ԥ��.��Ԥ��), ժҪ = ժҪ || '0' || ',' || c_Ԥ��.���㿨��� || ',' || -1 * c_Ԥ��.��Ԥ�� || '|'
                Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ Is Null;
                If Sql%RowCount = 0 Then
                  Insert Into ����Ԥ����¼
                    (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                     ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��������, У�Ա�־)
                  Values
                    (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * c_Ԥ��.��Ԥ��, Null, Null,
                     �˷�ʱ��_In, Null, Null, Null, ����Ա���_In, ����Ա����_In,
                     '0' || ',' || c_Ԥ��.���㿨��� || ',' || -1 * c_Ԥ��.��Ԥ�� || '|', n_��id, Null, Null, Null, Null, Null, Null,
                     n_����id, -1 * n_����id, 3, 1);
                End If;
                n_����״̬ := 1;
                n_ʵ�ʳ��� := c_Ԥ��.��Ԥ��;
              Else
                --Zl_Square_Update(v_ԭ����ids, n_����id, n_��id, �˷�ʱ��_In, -1 * n_����id, Null, n_ʵ�ս��, c_Ԥ��.���㿨���);
                Update ����Ԥ����¼
                Set ��Ԥ�� = ��Ԥ�� + (-1 * n_ʵ�ս��), ժҪ = ժҪ || '0' || ',' || c_Ԥ��.���㿨��� || ',' || -1 * n_ʵ�ս�� || '|'
                Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ Is Null;
                If Sql%RowCount = 0 Then
                  Insert Into ����Ԥ����¼
                    (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                     ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��������, У�Ա�־)
                  Values
                    (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * n_ʵ�ս��, Null, Null, �˷�ʱ��_In,
                     Null, Null, Null, ����Ա���_In, ����Ա����_In, '0' || ',' || c_Ԥ��.���㿨��� || ',' || -1 * n_ʵ�ս�� || '|', n_��id,
                     Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 3, 1);
                End If;
                n_����״̬ := 1;
                n_ʵ�ʳ��� := n_ʵ�ս��;
              End If;
            Else
              If c_Ԥ��.��Ԥ�� > n_ʵ�ս�� Then
                n_ʵ�ʳ��� := n_ʵ�ս��;
              Else
                n_ʵ�ʳ��� := c_Ԥ��.��Ԥ��;
              End If;
            End If;
          
            If c_Ԥ��.���㿨��� Is Null Then
              Update ��Ա�ɿ����
              Set ��� = Nvl(���, 0) - n_ʵ�ʳ���
              Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_���㷽ʽ
              Returning ��� Into n_����ֵ;
              If Sql%RowCount = 0 Then
                Insert Into ��Ա�ɿ����
                  (�տ�Ա, ���㷽ʽ, ����, ���)
                Values
                  (����Ա����_In, v_���㷽ʽ, 1, -1 * n_ʵ�ʳ���);
                n_����ֵ := n_ʵ�ʳ���;
              End If;
              If Nvl(n_����ֵ, 0) = 0 Then
                Delete From ��Ա�ɿ����
                Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_���㷽ʽ And Nvl(���, 0) = 0;
              End If;
            
              --��ԭԤ����¼
              Update ����Ԥ����¼
              Set ��Ԥ�� = ��Ԥ�� + (-1 * n_ʵ�ʳ���)
              Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_���㷽ʽ;
              If Sql%RowCount = 0 Then
                Insert Into ����Ԥ����¼
                  (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                   ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, У�Ա�־, ��������)
                Values
                  (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * n_ʵ�ʳ���, v_���㷽ʽ, Null, �˷�ʱ��_In,
                   Null, Null, Null, ����Ա���_In, ����Ա����_In, '', n_��id, Null, Null, Null, Null, Null, c_Ԥ��.������λ, n_����id,
                   -1 * n_����id, 0, 3);
              End If;
            End If;
            Update ����Ԥ����¼
            Set ��¼״̬ = 3
            Where ��¼���� = 3 And ��¼״̬ = 1 And ����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And
                  ���㷽ʽ = c_Ԥ��.���㷽ʽ;
            n_ʵ�ս�� := n_ʵ�ս�� - n_ʵ�ʳ���;
          End If;
        End If;
      End Loop;
    
      --���·�����˼�¼
      Update ������˼�¼
      Set ��¼״̬ = 2
      Where ����id In (Select ID From ������ü�¼ Where NO = No_In And ��¼���� = 1 And ��¼״̬ In (1, 3)) And ���� = 1;
      --���������¼
      Update ������ü�¼ Set ��¼״̬ = 3 Where NO = No_In And Mod(��¼����,10) = 1 And ��¼״̬ = 1;
      For r_Clinic In (Select ���, ��������, �۸񸸺�, ����id, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ���ձ���, ��������,
                              ��ҩ����, ����, Sum(����) As ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, Sum(Ӧ�ս��) As Ӧ�ս��,
                              Sum(ʵ�ս��) As ʵ�ս��, Sum(ͳ����) As ͳ����, ��������id, ������, ִ�в���id, ������, Max(���ʵ�id) As ���ʵ�id, ����ʱ��,
                              ʵ��Ʊ��
                       From ������ü�¼
                       Where NO = No_In And Mod(��¼����,10) = 1 And ��¼״̬ In (2, 3) And Nvl(���ӱ�־, 0) Not In (8, 9)
                       Group By ���, ��������, �۸񸸺�, ����id, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��, ���մ���id, ���ձ���,
                                ��������, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, ��������id, ������, ִ�в���id, ������, ����ʱ��, ʵ��Ʊ��
                       Having Sum(����) <> 0) Loop
        Insert Into ������ü�¼
          (ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, �����־, ����id, ��ʶ��, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��,
           ���մ���id, ���ձ���, ��������, ��ҩ����, ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ͳ����, ���ʷ���, ��������id, ������, ����ʱ��,
           �Ǽ�ʱ��, ִ�в���id, ������, ����Ա���, ����Ա����, ���ʵ�id, ժҪ, �ɿ���id, ����id, ���ʽ��, ����״̬)
        Values
          (���˷��ü�¼_Id.Nextval, 1, No_In, r_Clinic.ʵ��Ʊ��, 2, r_Clinic.���, r_Clinic.��������, r_Clinic.�۸񸸺�, 1, r_Clinic.����id,
           '', r_Clinic.����, r_Clinic.�Ա�, r_Clinic.����, r_Clinic.���˿���id, r_Clinic.�ѱ�, r_Clinic.�շ����, r_Clinic.�շ�ϸĿid,
           r_Clinic.���㵥λ, r_Clinic.������Ŀ��, r_Clinic.���մ���id, r_Clinic.���ձ���, r_Clinic.��������, r_Clinic.��ҩ����, r_Clinic.����,
           -1 * r_Clinic.����, r_Clinic.�Ӱ��־, r_Clinic.���ӱ�־, r_Clinic.������Ŀid, r_Clinic.�վݷ�Ŀ, r_Clinic.��׼����,
           -1 * r_Clinic.Ӧ�ս��, -1 * r_Clinic.ʵ�ս��, -1 * r_Clinic.ͳ����, 0, r_Clinic.��������id, r_Clinic.������, r_Clinic.����ʱ��,
           �˷�ʱ��_In, r_Clinic.ִ�в���id, r_Clinic.������, ����Ա���_In, ����Ա����_In, r_Clinic.���ʵ�id, '', n_��id, n_����id,
           -1 * r_Clinic.ʵ�ս��, 0);
      End Loop;
    Else
      --4.�˿�תԤ��(������Ʊ��,�ɲ���Աͨ���ش����)
      For r_Pay In (Select Min(a.Id) As Ԥ��id, a.���㷽ʽ, Sum(a.��Ԥ��) As ��Ԥ��, 2 As Ԥ�����, a.�����id, a.���㿨���, a.����, a.������ˮ��,
                           a.����˵��, a.������λ, b.����
                    From ����Ԥ����¼ A, ���㷽ʽ B
                    Where a.��¼���� = 3 And a.����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And
                          a.���㷽ʽ = b.���� And (b.���� In (1, 2, 7, 8)) And a.���㷽ʽ Is Not Null
                    Group By a.���㷽ʽ, Ԥ�����, a.�����id, a.���㿨���, a.����, b.����, a.������ˮ��, a.����˵��, a.������λ


                    
                    Having Sum(a.��Ԥ��) <> 0
                    Order By a.�����id, ���� Desc) Loop
        --4.1����Ԥ����� (�����ڲ����˷ѵ����)
        --���е���,����������Ԥ�����
        --��Ϊ�տ�������ɿ�,������Ա�ɿ�����ޱ仯
        If n_ʵ�ս�� <> 0 Then
          If r_Pay.���� = 7 Or (r_Pay.���� = 8 And r_Pay.�����id Is Not Null) Then
            If r_Pay.��Ԥ�� > n_ʵ�ս�� Then
              Update ����Ԥ����¼
              Set ��Ԥ�� = ��Ԥ�� + (-1 * n_ʵ�ս��), ժҪ = ժҪ || '1' || ',' || r_Pay.�����id || ',' || -1 * n_ʵ�ս�� || '|'
              Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ Is Null;
              If Sql%RowCount = 0 Then
                Insert Into ����Ԥ����¼
                  (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                   ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��������, У�Ա�־)
                Values
                  (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * n_ʵ�ս��, Null, Null, �˷�ʱ��_In,
                   Null, Null, Null, ����Ա���_In, ����Ա����_In, '1' || ',' || r_Pay.�����id || ',' || -1 * n_ʵ�ս�� || '|', n_��id,
                   Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 3, 1);
              End If;
            
              Update ����Ԥ����¼
              Set ��¼״̬ = 3
              Where ��¼���� = 3 And ��¼״̬ = 1 And ����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And
                    ���㷽ʽ = r_Pay.���㷽ʽ;
              n_����״̬ := 1;
              n_ʵ�ս�� := 0;
            Else
              Update ����Ԥ����¼
              Set ��Ԥ�� = ��Ԥ�� + (-1 * r_Pay.��Ԥ��), ժҪ = ժҪ || '1' || ',' || r_Pay.�����id || ',' || -1 * r_Pay.��Ԥ�� || '|'
              Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ Is Null;
              If Sql%RowCount = 0 Then
                Insert Into ����Ԥ����¼
                  (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                   ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��������, У�Ա�־)
                Values
                  (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * r_Pay.��Ԥ��, Null, Null, �˷�ʱ��_In,
                   Null, Null, Null, ����Ա���_In, ����Ա����_In, '1' || ',' || r_Pay.�����id || ',' || -1 * r_Pay.��Ԥ�� || '|',
                   n_��id, Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 3, 1);
              End If;
            
              Update ����Ԥ����¼
              Set ��¼״̬ = 3
              Where ��¼���� = 3 And ��¼״̬ = 1 And ����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And
                    ���㷽ʽ = r_Pay.���㷽ʽ;
              n_����״̬ := 1;
              n_ʵ�ս�� := n_ʵ�ս�� - r_Pay.��Ԥ��;
            End If;
          Else
            n_ʵ�ʳ��� := 0;
            If r_Pay.���� In (3, 4) Or (r_Pay.���� = 8 And r_Pay.���㿨��� Is Not Null) Then
              v_���㷽ʽ := r_Pay.���㷽ʽ;
            Else
              If ���㷽ʽ_In Is Null Then
                Begin
                  Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And ���� Like '%�ֽ�%' And Rownum < 2;
                Exception
                  When Others Then
                    Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And Rownum < 2;
                End;
              Else
                v_���㷽ʽ := ���㷽ʽ_In;
              End If;
            End If;
          
            If r_Pay.���� = 8 And r_Pay.���㿨��� Is Not Null Then
              If n_ʵ�ս�� >= r_Pay.��Ԥ�� Then
                --Zl_Square_Update(v_ԭ����ids, n_����id, n_��id, �˷�ʱ��_In, -1 * n_����id, Null, r_Pay.��Ԥ��, r_Pay.���㿨���);
                Update ����Ԥ����¼
                Set ��Ԥ�� = ��Ԥ�� + (-1 * r_Pay.��Ԥ��), ժҪ = ժҪ || '0' || ',' || r_Pay.���㿨��� || ',' || -1 * r_Pay.��Ԥ�� || '|'
                Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ Is Null;
                If Sql%RowCount = 0 Then
                  Insert Into ����Ԥ����¼
                    (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                     ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��������, У�Ա�־)
                  Values
                    (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * r_Pay.��Ԥ��, Null, Null,
                     �˷�ʱ��_In, Null, Null, Null, ����Ա���_In, ����Ա����_In,
                     '0' || ',' || r_Pay.���㿨��� || ',' || -1 * r_Pay.��Ԥ�� || '|', n_��id, Null, Null, Null, Null, Null,
                     Null, n_����id, -1 * n_����id, 3, 1);
                End If;
                n_����״̬ := 1;
                n_ʵ�ʳ��� := r_Pay.��Ԥ��;
              Else
                --Zl_Square_Update(v_ԭ����ids, n_����id, n_��id, �˷�ʱ��_In, -1 * n_����id, Null, n_ʵ�ս��, r_Pay.���㿨���);
                Update ����Ԥ����¼
                Set ��Ԥ�� = ��Ԥ�� + (-1 * n_ʵ�ս��), ժҪ = ժҪ || '0' || ',' || r_Pay.���㿨��� || ',' || -1 * n_ʵ�ս�� || '|'
                Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ Is Null;
                If Sql%RowCount = 0 Then
                  Insert Into ����Ԥ����¼
                    (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                     ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��������, У�Ա�־)
                  Values
                    (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * n_ʵ�ս��, Null, Null, �˷�ʱ��_In,
                     Null, Null, Null, ����Ա���_In, ����Ա����_In, '0' || ',' || r_Pay.���㿨��� || ',' || -1 * n_ʵ�ս�� || '|', n_��id,
                     Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 3, 1);
                End If;
                n_����״̬ := 1;
                n_ʵ�ʳ��� := n_ʵ�ս��;
              End If;
            Else
              If r_Pay.��Ԥ�� > n_ʵ�ս�� Then
                n_ʵ�ʳ��� := n_ʵ�ս��;
              Else
                n_ʵ�ʳ��� := r_Pay.��Ԥ��;
              End If;
            End If;
          
            If r_Pay.���� Not In (3, 4, 7, 8) Then
              Update ����Ԥ����¼
              Set ��� = ��� + n_ʵ�ʳ���
              Where ��¼���� = 1 And ��¼״̬ = 1 And �տ�ʱ�� = �˷�ʱ��_In And ����id + 0 = n_����id And ���㷽ʽ = v_���㷽ʽ;
              If Sql%RowCount = 0 Then
                v_Ԥ��no := Nextno(11);
                Insert Into ����Ԥ����¼
                  (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                   ժҪ, �ɿ���id, Ԥ�����)
                Values
                  (����Ԥ����¼_Id.Nextval, v_Ԥ��no, Null, 1, 1, n_����id, ��ҳid_In, ��Ժ����id_In, n_ʵ�ʳ���, v_���㷽ʽ, Null, �˷�ʱ��_In,
                   Null, Null, Null, ����Ա���_In, ����Ա����_In, '����תסԺԤ��', n_��id, r_Pay.Ԥ�����);
              End If;
            
              --�������
              Update �������
              Set Ԥ����� = Nvl(Ԥ�����, 0) + n_ʵ�ʳ���
              Where ���� = 1 And ����id = n_����id And ���� = 2
              Returning Ԥ����� Into n_����ֵ;
              If Sql%RowCount = 0 Then
                Insert Into ������� (����id, ����, ����, Ԥ�����, �������) Values (n_����id, 1, 2, n_ʵ�ʳ���, 0);
                n_����ֵ := n_ʵ�ʳ���;
              End If;
              If Nvl(n_����ֵ, 0) = 0 Then
                Delete From �������
                Where ����id = n_����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
              End If;
            End If;
            --4.2�ɿ����ݴ���
            --   ��Ϊû��ʵ���ղ��˵�Ǯ,���Բ�����
            --�����˷��������ԭԤ����¼
            If r_Pay.���� In (3, 4) Then
              Update ��Ա�ɿ����
              Set ��� = Nvl(���, 0) - n_ʵ�ʳ���
              Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = r_Pay.���㷽ʽ
              Returning ��� Into n_����ֵ;
              If Sql%RowCount = 0 Then
                Insert Into ��Ա�ɿ����
                  (�տ�Ա, ���㷽ʽ, ����, ���)
                Values
                  (����Ա����_In, r_Pay.���㷽ʽ, 1, -1 * n_ʵ�ʳ���);
                n_����ֵ := n_ʵ�ʳ���;
              End If;
              If Nvl(n_����ֵ, 0) = 0 Then
                Delete From ��Ա�ɿ����
                Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = r_Pay.���㷽ʽ And Nvl(���, 0) = 0;
              End If;
            End If;
          
            If r_Pay.���� <> 8 Then
              Update ����Ԥ����¼
              Set ��Ԥ�� = ��Ԥ�� + (-1 * n_ʵ�ʳ���)
              Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_���㷽ʽ;
              If Sql%RowCount = 0 Then
                Insert Into ����Ԥ����¼
                  (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����,
                   ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, У�Ա�־, ��������)
                Values
                  (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * n_ʵ�ʳ���, v_���㷽ʽ, Null, �˷�ʱ��_In,
                   Null, Null, Null, ����Ա���_In, ����Ա����_In, '', n_��id, r_Pay.�����id, r_Pay.���㿨���, r_Pay.����, r_Pay.������ˮ��,
                   r_Pay.����˵��, r_Pay.������λ, n_����id, -1 * n_����id, 0, 3);
              End If;
            End If;
          
            Update ����Ԥ����¼
            Set ��¼״̬ = 3
            Where ��¼���� = 3 And ��¼״̬ = 1 And ����id In (Select Column_Value From Table(f_Str2list(v_ԭ����ids))) And
                  ���㷽ʽ = r_Pay.���㷽ʽ;
            n_ʵ�ս�� := n_ʵ�ս�� - n_ʵ�ʳ���;
          
          End If;
        End If;
      End Loop;
    End If;
  
    If ����_In Is Not Null Then
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ, �ɿ���id,
         �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, У�Ա�־, ��������)
      Values
        (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, ����_In, v_����, Null, �˷�ʱ��_In, Null, Null,
         Null, ����Ա���_In, ����Ա����_In, '', n_��id, Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 0, 3);
    End If;
    Delete From ����Ԥ����¼
    Where ����id = n_����id And ��¼���� = 3 And ��¼״̬ = 2 And ��Ԥ�� = 0 And ���㷽ʽ Is Not Null;
    Delete From ����Ԥ����¼ Where ����id = n_ԭ����id And ժҪ = 'Ԥ����ʱ��¼' And ��¼���� = 3;
    Update ������ü�¼ Set ����״̬ = Nvl(n_����״̬, 0) Where NO = No_In And Mod(��¼����, 10) = 1 And ��¼״̬ = 2;
  Else
    --ҽ��������ת��
    For r_Nos In (Select Distinct a.No
                  From ������ü�¼ A
                  Where Mod(a.��¼����, 10) = 1 And a.��¼״̬ In (1, 3) And a.����id = ԭ����id_In) Loop
      v_Nos := v_Nos || ',' || r_Nos.No;
    End Loop;
    v_Nos := Substr(v_Nos, 2);
  
    For r_����ids In (Select Distinct a.����id
                    From ������ü�¼ A
                    Where a.No In (Select Column_Value From Table(f_Str2list(v_Nos))) And Mod(a.��¼����, 10) = 1 And
                          a.��¼״̬ <> 0) Loop
      v_����ids := v_����ids || ',' || r_����ids.����id;
    End Loop;
    v_����ids := Substr(v_����ids, 2);
    Select Count(a.No), Sum(a.ʵ�ս��)
    Into n_Count, n_ʵ�ս��
    From ������ü�¼ A
    Where a.No In (Select Column_Value From Table(f_Str2list(v_Nos))) And Mod(a.��¼����, 10) = 1;
    If n_Count = 0 Or n_ʵ�ս�� = 0 Then
      v_Err_Msg := '���ν��㲻���շѻ��򲢷�ԭ�����˲����˸ý���,����תΪסԺ����.';
      Raise Err_Item;
    End If;
  
    Select ����id, ����id, ��������id, ������
    Into n_ԭ����id, n_����id, n_��������id, v_������
    From ������ü�¼
    Where ����id = ԭ����id_In And Mod(��¼����, 10) = 1 And ��¼״̬ In (1, 3) And Rownum < 2;
  
    Begin
      Select 1
      Into n_�����˷�
      From ������ü�¼ A
      Where Mod(a.��¼����, 10) = 1 And a.��¼״̬ = 2 And a.����id In (Select Column_Value From Table(f_Str2list(v_����ids))) And
            Rownum < 2;
    Exception
      When Others Then
        n_�����˷� := 0;
    End;
  
    Begin
      Select 0
      Into n_�����˷�
      From ������ü�¼ A
      Where ��¼���� = 11 And a.����id In (Select Column_Value From Table(f_Str2list(v_����ids))) And Rownum < 2;
    Exception
      When Others Then
        Null;
    End;
    Begin
      Select Count(Avg(1))
      Into n_�˷�����
      From ����Ԥ����¼ A
      Where a.��¼���� = 3 And a.��¼״̬ <> 0 And ����id In (Select Column_Value From Table(f_Str2list(v_����ids)))
      Group By a.���㷽ʽ;
    Exception
      When Others Then
        n_�˷����� := 0;
    End;
    --1.1���Ϸ��ü�¼
    If ����id_In Is Null Then
      Select ���˽��ʼ�¼_Id.Nextval Into n_����id From Dual;
    Else
      n_����id := ����id_In;
    End If;
    Insert Into ������ü�¼
      (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�, ���˿���id, �շ����, �շ�ϸĿid,
       ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ������, ִ����, ִ��״̬, ִ��ʱ��,
       ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ����id, ���ʽ��, ������Ŀ��, ���մ���id, ͳ����, ժҪ, �Ƿ��ϴ�, ���ձ���, ��������, �ɿ���id, ����״̬)
      Select ���˷��ü�¼_Id.Nextval, a.No, a.ʵ��Ʊ��, a.��¼����, 2, a.���, a.��������, a.�۸񸸺�, a.����id, a.ҽ�����, a.�����־, a.����, a.�Ա�, a.����,
             a.��ʶ��, a.���ʽ, a.�ѱ�, a.���˿���id, a.�շ����, a.�շ�ϸĿid, a.���㵥λ, a.����, a.��ҩ����, -1 * a.����, a.�Ӱ��־, a.���ӱ�־, a.������Ŀid,
             a.�վݷ�Ŀ, a.���ʷ���, a.��׼����, -1 * a.Ӧ�ս��, -1 * a.ʵ�ս��, a.��������id, a.������, a.ִ�в���id, a.������, a.ִ����, -1, a.ִ��ʱ��,
             ����Ա���_In, ����Ա����_In, a.����ʱ��, �˷�ʱ��_In, n_����id, -1 * a.���ʽ��, a.������Ŀ��, a.���մ���id, a.ͳ����, a.ժҪ,
             Decode(Nvl(a.���ӱ�־, 0), 9, 1, 0), a.���ձ���, a.��������, n_��id, 0
      From ������ü�¼ A
      Where a.No In (Select Column_Value From Table(f_Str2list(v_Nos))) And Mod(a.��¼����, 10) = 1 And a.��¼״̬ = 1;
  
    --����ҽ��
    For r_ҽ�� In (Select ����id, NO, ���㷽ʽ, ���, ��ע
                 From ҽ��������ϸ
                 Where NO In (Select Column_Value From Table(f_Str2list(v_Nos))) And
                       ����id In (Select Column_Value From Table(f_Str2list(v_����ids)))) Loop
      Update ҽ��������ϸ
      Set ��� = ��� + (-1 * r_ҽ��.���)
      Where NO = r_ҽ��.No And ����id = r_ҽ��.����id And ���㷽ʽ = r_ҽ��.���㷽ʽ;
      If Sql%RowCount = 0 Then
        Insert Into ҽ��������ϸ
          (����id, NO, ���㷽ʽ, ���)
        Values
          (r_ҽ��.����id, r_ҽ��.No, r_ҽ��.���㷽ʽ, -1 * r_ҽ��.���);
      End If;
    End Loop;
  
    --Update ������ü�¼ Set ��¼״̬ = 3 Where NO = No_In And ��¼���� = 1 And ��¼״̬ = 1;
    --1.2����Ԥ����¼
    --���ϳ�Ԥ������
    If n_�����˷� = 0 And Nvl(�����˷�_In, 0) = 0 Then
      For r_Prepay In (Select NO, ʵ��Ʊ��, ����id, ��ҳid, ����id, ���㷽ʽ, �������, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, -1 * Sum(��Ԥ��) As ��Ԥ��,
                              �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������
                       From ����Ԥ����¼ A
                       Where ��¼���� In (1, 11) And a.����id In (Select Column_Value From Table(f_Str2list(v_����ids))) And
                             Nvl(��Ԥ��, 0) <> 0
                       Group By n_Tempid, NO, ʵ��Ʊ��, ����id, ��ҳid, ����id, ���㷽ʽ, �������, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, �����id, ���㿨���,
                                ����, ������ˮ��, ����˵��, ������λ, ��������) Loop
        Select ����Ԥ����¼_Id.Nextval Into n_Tempid From Dual;
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
           ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, Ԥ�����, ��������)
          Select n_Tempid, r_Prepay.No, r_Prepay.ʵ��Ʊ��, 11, 1, r_Prepay.����id, r_Prepay.��ҳid, r_Prepay.����id, Null,
                 r_Prepay.���㷽ʽ, r_Prepay.�������, Null, r_Prepay.�ɿλ, r_Prepay.��λ������, r_Prepay.��λ�ʺ�, �˷�ʱ��_In, ����Ա����_In,
                 ����Ա���_In, r_Prepay.��Ԥ��, n_����id, n_��id, r_Prepay.�����id, r_Prepay.���㿨���, r_Prepay.����, r_Prepay.������ˮ��,
                 r_Prepay.����˵��, r_Prepay.������λ, -1 * n_����id, 1, r_Prepay.��������
          From Dual;
      End Loop;
    
      For v_Ԥ�� In (Select ����id, Nvl(Ԥ�����, 2) As Ԥ�����, Nvl(Sum(Nvl(��Ԥ��, 0)), 0) As Ԥ�����
                   From ����Ԥ����¼ A
                   Where ��¼���� In (1, 11) And a.����id In (Select Column_Value From Table(f_Str2list(v_����ids))) And
                         a.����id <> n_����id
                   Group By ����id, Nvl(Ԥ�����, 2)
                   Having Sum(Nvl(��Ԥ��, 0)) <> 0) Loop
      
        Update �������
        Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(v_Ԥ��.Ԥ�����, 0)
        Where ����id = v_Ԥ��.����id And ���� = Nvl(v_Ԥ��.Ԥ�����, 2) And ���� = 1
        Returning Ԥ����� Into n_����ֵ;
        If Sql%RowCount = 0 Then
          Insert Into �������
            (����id, ����, Ԥ�����, ����)
          Values
            (v_Ԥ��.����id, Nvl(v_Ԥ��.Ԥ�����, 2), v_Ԥ��.Ԥ�����, 1);
          n_����ֵ := v_Ԥ��.Ԥ�����;
        End If;
        If Nvl(n_����ֵ, 0) = 0 Then
          Delete From �������
          Where ����id = v_Ԥ��.����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
        End If;
      End Loop;
    Else
      If n_�˷����� = 0 And Nvl(�����˷�_In, 0) = 0 Then
        --ֻʹ����Ԥ����ԭ���˻�Ԥ��
        For r_Prepay In (Select NO, ʵ��Ʊ��, ����id, ��ҳid, ����id, Max(���㷽ʽ) As ���㷽ʽ, �������, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��,
                                -1 * Sum(��Ԥ��) As ��Ԥ��, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ��������
                         From ����Ԥ����¼ A
                         Where ��¼���� In (1, 11) And a.����id In (Select Column_Value From Table(f_Str2list(v_����ids))) And
                               Nvl(��Ԥ��, 0) <> 0
                         Group By n_Tempid, NO, ʵ��Ʊ��, ����id, ��ҳid, ����id, �������, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, �����id, ���㿨���, ����,
                                  ������ˮ��, ����˵��, ������λ, ��������) Loop
          Select ����Ԥ����¼_Id.Nextval Into n_Tempid From Dual;
          Insert Into ����Ԥ����¼
            (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
             ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, Ԥ�����, ��������)
            Select n_Tempid, r_Prepay.No, r_Prepay.ʵ��Ʊ��, 11, 1, r_Prepay.����id, r_Prepay.��ҳid, r_Prepay.����id, Null,
                   r_Prepay.���㷽ʽ, r_Prepay.�������, Null, r_Prepay.�ɿλ, r_Prepay.��λ������, r_Prepay.��λ�ʺ�, �˷�ʱ��_In, ����Ա����_In,
                   ����Ա���_In, r_Prepay.��Ԥ��, n_����id, n_��id, r_Prepay.�����id, r_Prepay.���㿨���, r_Prepay.����, r_Prepay.������ˮ��,
                   r_Prepay.����˵��, r_Prepay.������λ, -1 * n_����id, 1, r_Prepay.��������
            From Dual;
          Select -1 * ��Ԥ�� Into n_Ԥ����� From ����Ԥ����¼ Where ID = n_Tempid;
          Update �������
          Set Ԥ����� = Nvl(Ԥ�����, 0) + Nvl(n_Ԥ�����, 0)
          Where ����id = r_Prepay.����id And ���� = 1 And ���� = 1
          Returning Ԥ����� Into n_����ֵ;
          If Sql%RowCount = 0 Then
            Insert Into ������� (����id, ����, Ԥ�����, ����) Values (n_����id, 1, n_Ԥ�����, 1);
            n_����ֵ := n_Ԥ�����;
          End If;
          If Nvl(n_����ֵ, 0) = 0 Then
            Delete From �������
            Where ����id = r_Prepay.����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
          End If;
        End Loop;
      Else
        Begin
          Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And ���� Like '%�ֽ�%' And Rownum < 2;
        Exception
          When Others Then
            Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And Rownum < 2;
        End;
        Select ����Ԥ����¼_Id.Nextval Into n_Tempid From Dual;
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
           ����id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, ��������)
          Select n_Tempid, Max(NO), Max(ʵ��Ʊ��), 3, 3, ����id, ��ҳid, ����id, Null, v_���㷽ʽ, Max(�������), 'Ԥ����ʱ��¼', Null, Null,
                 Null, Max(�տ�ʱ��), ����Ա����_In, ����Ա���_In, Sum(��Ԥ��), n_ԭ����id, Null, Null, Null, Null, Null, Null,
                 -1 * n_ԭ����id, 3
          From ����Ԥ����¼ A
          Where ��¼���� In (1, 11) And a.����id In (Select Column_Value From Table(f_Str2list(v_����ids))) And
                Nvl(��Ԥ��, 0) <> 0
          Group By n_Tempid, 3, 3, ����id, ��ҳid, ����id, Null, v_���㷽ʽ, 'Ԥ����ʱ��¼', ����Ա����_In, ����Ա���_In, n_ԭ����id;
      End If;
    End If;
  
    --��������ɷѼ�ҽ������
    Insert Into ����Ԥ����¼
      (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����,
       �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, ��������)
      Select ����Ԥ����¼_Id.Nextval, ��¼����, NO, 2, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �˷�ʱ��_In, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���_In, ����Ա����_In,
             0, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, -1 * n_����id, ��������
      From ����Ԥ����¼ A, ���㷽ʽ B
      Where a.��¼���� = 3 And a.��¼״̬ = 1 And a.����id In (Select Column_Value From Table(f_Str2list(v_����ids))) And
            a.���㷽ʽ = b.���� And b.���� Not In (7, 8);
  
    Insert Into ����Ԥ����¼
      (ID, ��¼����, NO, ��¼״̬, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ��Ԥ��, ����id, �ɿ���id, Ԥ�����,
       �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, �������, ��������, У�Ա�־)
      Select ����Ԥ����¼_Id.Nextval, ��¼����, NO, 2, ����id, ��ҳid, ժҪ, ���㷽ʽ, �������, �˷�ʱ��_In, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���_In, ����Ա����_In,
             0, n_����id, n_��id, Ԥ�����, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, -1 * n_����id, ��������, 1
      From ����Ԥ����¼ A, ���㷽ʽ B
      Where a.��¼���� = 3 And a.��¼״̬ = 1 And a.����id In (Select Column_Value From Table(f_Str2list(v_����ids))) And
            a.���㷽ʽ = b.���� And b.���� = 7;
    If Sql%RowCount <> 0 Then
      n_����״̬ := 1;
    End If;
  
    Update ����Ԥ����¼
    Set ��¼״̬ = 3
    Where ��¼���� = 3 And ��¼״̬ = 1 And ����id In (Select Column_Value From Table(f_Str2list(v_����ids)));
  
    --2.Ʊ���ջ�
    --������ǰû�д�ӡ,���ջ�
    For r_Nos In (Select Distinct a.No
                  From ������ü�¼ A
                  Where Mod(a.��¼����, 10) = 1 And a.��¼״̬ In (1, 3) And
                        a.����id In (Select Column_Value From Table(f_Str2list(v_����ids)))) Loop
    
      Select Nvl(Max(ID), 0)
      Into n_��ӡid
      From (Select b.Id
             From Ʊ��ʹ����ϸ A, Ʊ�ݴ�ӡ���� B
             Where a.��ӡid = b.Id And a.���� = 1 And a.ԭ�� In (1, 3) And b.�������� = 1 And b.No = r_Nos.No
             Order By a.ʹ��ʱ�� Desc)
      Where Rownum < 2;
      If n_��ӡid > 0 Then
        --���ŵ���ѭ������ʱֻ���ջ�һ��
        Select Count(��ӡid) Into n_Count From Ʊ��ʹ����ϸ Where Ʊ�� = 1 And ���� = 2 And ��ӡid = n_��ӡid;
        If n_Count = 0 Then
          Insert Into Ʊ��ʹ����ϸ
            (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����)
            Select Ʊ��ʹ����ϸ_Id.Nextval, Ʊ��, ����, 2, 2, ����id, ��ӡid, �˷�ʱ��_In, ����Ա����_In
            From Ʊ��ʹ����ϸ
            Where ��ӡid = n_��ӡid And Ʊ�� = 1 And ���� = 1;
        End If;
      End If;
    End Loop;
  
    --3.�ɿ����ݴ���(
    --   �����������:
    --    1. ת������ֱ�����ʵ�,��ɿ����ݲ�����;
    --    2. ��ת��,�ٵ������˿���Ʊ,����Ҫ���нɿ����ݴ���
    If Nvl(�����˷�_In, 0) = 1 Then
      For c_Ԥ�� In (Select a.���㷽ʽ, Sum(a.��Ԥ��) As ��Ԥ��, 2 As Ԥ�����, a.�����id, a.���㿨���, a.����, Min(a.������ˮ��) As ������ˮ��,
                          Min(a.����˵��) As ����˵��, Min(a.������λ) As ������λ, b.����
                   From ����Ԥ����¼ A, ���㷽ʽ B
                   Where a.��¼���� = 3 And a.��¼״̬ In (2, 3) And
                         a.����id In (Select Column_Value From Table(f_Str2list(v_����ids))) And a.���㷽ʽ = b.���� And
                         b.���� In (1, 2, 3, 4, 7, 8) And a.���㷽ʽ Is Not Null
                   Group By a.���㷽ʽ, Ԥ�����, a.�����id, a.���㿨���, a.����, b.����
                   Having Sum(a.��Ԥ��) <> 0) Loop
        Begin
          Select �Ƿ����� Into n_���� From ҽ�ƿ���� Where ID = c_Ԥ��.�����id;
        Exception
          When Others Then
            n_���� := 0;
        End;
        If (c_Ԥ��.���� = 7 Or (c_Ԥ��.���� = 8 And c_Ԥ��.�����id Is Not Null)) And n_���� = 0 Then
          Update ����Ԥ����¼
          Set ��Ԥ�� = ��Ԥ�� + (-1 * c_Ԥ��.��Ԥ��), ժҪ = ժҪ || '1' || ',' || c_Ԥ��.�����id || ',' || -1 * c_Ԥ��.��Ԥ�� || '|'
          Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ Is Null;
          If Sql%RowCount = 0 Then
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ,
               �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��������, У�Ա�־)
            Values
              (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * c_Ԥ��.��Ԥ��, Null, Null, �˷�ʱ��_In,
               Null, Null, Null, ����Ա���_In, ����Ա����_In, '1' || ',' || c_Ԥ��.�����id || ',' || -1 * c_Ԥ��.��Ԥ�� || '|', n_��id,
               Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 3, 1);
          End If;
          n_����״̬ := 1;
        Else
          If c_Ԥ��.���� In (3, 4) Or (c_Ԥ��.���� = 8 And c_Ԥ��.���㿨��� Is Not Null) Then
            v_���㷽ʽ := c_Ԥ��.���㷽ʽ;
          Else
            If ���㷽ʽ_In Is Null Then
              Begin
                Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And ���� Like '%�ֽ�%' And Rownum < 2;
              Exception
                When Others Then
                  Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And Rownum < 2;
              End;
            Else
              v_���㷽ʽ := ���㷽ʽ_In;
            End If;
          End If;
        
          If c_Ԥ��.���� = 8 And c_Ԥ��.���㿨��� Is Not Null Then
            --Zl_Square_Update(v_����ids, n_����id, n_��id, �˷�ʱ��_In, -1 * n_����id, Null, c_Ԥ��.��Ԥ��, c_Ԥ��.���㿨���);
            Update ����Ԥ����¼
            Set ��Ԥ�� = ��Ԥ�� + (-1 * c_Ԥ��.��Ԥ��), ժҪ = ժҪ || '0' || ',' || c_Ԥ��.���㿨��� || ',' || -1 * c_Ԥ��.��Ԥ�� || '|'
            Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ Is Null;
            If Sql%RowCount = 0 Then
              Insert Into ����Ԥ����¼
                (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ,
                 �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��������, У�Ա�־)
              Values
                (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * c_Ԥ��.��Ԥ��, Null, Null, �˷�ʱ��_In,
                 Null, Null, Null, ����Ա���_In, ����Ա����_In, '0' || ',' || c_Ԥ��.���㿨��� || ',' || -1 * c_Ԥ��.��Ԥ�� || '|', n_��id,
                 Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 3, 1);
            End If;
            n_����״̬ := 1;
          End If;
          If c_Ԥ��.���㿨��� Is Null Then
            Update ��Ա�ɿ����
            Set ��� = Nvl(���, 0) - c_Ԥ��.��Ԥ��
            Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_���㷽ʽ
            Returning ��� Into n_����ֵ;
            If Sql%RowCount = 0 Then
              Insert Into ��Ա�ɿ����
                (�տ�Ա, ���㷽ʽ, ����, ���)
              Values
                (����Ա����_In, v_���㷽ʽ, 1, -1 * c_Ԥ��.��Ԥ��);
              n_����ֵ := c_Ԥ��.��Ԥ��;
            End If;
            If Nvl(n_����ֵ, 0) = 0 Then
              Delete From ��Ա�ɿ����
              Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = v_���㷽ʽ And Nvl(���, 0) = 0;
            End If;
            --�����˷��������ԭԤ����¼
            Update ����Ԥ����¼
            Set ��Ԥ�� = ��Ԥ�� + (-1 * c_Ԥ��.��Ԥ��)
            Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_���㷽ʽ;
            If Sql%RowCount = 0 Then
              Insert Into ����Ԥ����¼
                (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ,
                 �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, У�Ա�־, ��������)
              Values
                (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * c_Ԥ��.��Ԥ��, v_���㷽ʽ, Null, �˷�ʱ��_In,
                 Null, Null, Null, ����Ա���_In, ����Ա����_In, '', n_��id, Null, Null, Null, Null, Null, c_Ԥ��.������λ, n_����id,
                 -1 * n_����id, 0, 3);
            End If;
          End If;
        End If;
      End Loop;
    
      --���·�����˼�¼
      Update ������˼�¼
      Set ��¼״̬ = 2
      Where ����id In (Select a.Id
                     From ������ü�¼ A
                     Where a.No In (Select Column_Value From Table(f_Str2list(v_Nos))) And Mod(a.��¼����, 10) = 1 And
                           a.��¼״̬ In (1, 3)) And ���� = 1;
      --���������¼
      For r_Nos In (Select Distinct NO
                    From ������ü�¼
                    Where Mod(��¼����, 10) = 1 And ��¼״̬ In (1, 3) And
                          ����id In (Select Column_Value From Table(f_Str2list(v_����ids)))) Loop
        Update ������ü�¼ Set ��¼״̬ = 3 Where NO = r_Nos.No And Mod(��¼����, 10) = 1 And ��¼״̬ = 1;
      End Loop;
      For r_Clinic In (Select Min(a.��¼����) As ��¼����, a.No, a.���, a.��������, a.�۸񸸺�, a.����id, a.����, a.�Ա�, a.����, a.���˿���id, a.�ѱ�,
                              a.�շ����, a.�շ�ϸĿid, a.���㵥λ, a.������Ŀ��, a.���մ���id, a.���ձ���, a.��������, a.��ҩ����, a.����, Sum(a.����) As ����,
                              a.�Ӱ��־, a.���ӱ�־, a.������Ŀid, a.�վݷ�Ŀ, a.��׼����, Sum(a.Ӧ�ս��) As Ӧ�ս��, Sum(a.ʵ�ս��) As ʵ�ս��,
                              Sum(a.ͳ����) As ͳ����, a.��������id, a.������, a.ִ�в���id, a.������, Max(a.���ʵ�id) As ���ʵ�id,
                              Max(a.�Ƿ���) As �Ƿ���, a.����ʱ��, Min(a.ʵ��Ʊ��) As ʵ��Ʊ��
                       From ������ü�¼ A
                       Where a.No In (Select Column_Value From Table(f_Str2list(v_Nos))) And Mod(a.��¼����, 10) = 1 And
                             a.��¼״̬ In (2, 3) And Nvl(a.���ӱ�־, 0) Not In (8, 9)
                       Group By a.No, a.���, a.��������, a.�۸񸸺�, a.����id, a.����, a.�Ա�, a.����, a.���˿���id, a.�ѱ�, a.�շ����, a.�շ�ϸĿid,
                                a.���㵥λ, a.������Ŀ��, a.���մ���id, a.���ձ���, a.��������, a.��ҩ����, a.����, a.�Ӱ��־, a.���ӱ�־, a.������Ŀid, a.�վݷ�Ŀ,
                                a.��׼����, a.��������id, a.������, a.ִ�в���id, a.������, a.����ʱ��
                       Having Sum(a.����) <> 0) Loop
        Insert Into ������ü�¼
          (ID, ��¼����, NO, ʵ��Ʊ��, ��¼״̬, ���, ��������, �۸񸸺�, �����־, ����id, ��ʶ��, ����, �Ա�, ����, ���˿���id, �ѱ�, �շ����, �շ�ϸĿid, ���㵥λ, ������Ŀ��,
           ���մ���id, ���ձ���, ��������, ��ҩ����, ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ��׼����, Ӧ�ս��, ʵ�ս��, ͳ����, ���ʷ���, ��������id, ������, ����ʱ��,
           �Ǽ�ʱ��, ִ�в���id, ������, ����Ա���, ����Ա����, ���ʵ�id, ժҪ, �Ƿ���, �ɿ���id, ����id, ���ʽ��, ִ��״̬, ����״̬)
        Values
          (���˷��ü�¼_Id.Nextval, r_Clinic.��¼����, r_Clinic.No, r_Clinic.ʵ��Ʊ��, 2, r_Clinic.���, r_Clinic.��������, r_Clinic.�۸񸸺�,
           1, r_Clinic.����id, '', r_Clinic.����, r_Clinic.�Ա�, r_Clinic.����, r_Clinic.���˿���id, r_Clinic.�ѱ�, r_Clinic.�շ����,
           r_Clinic.�շ�ϸĿid, r_Clinic.���㵥λ, r_Clinic.������Ŀ��, r_Clinic.���մ���id, r_Clinic.���ձ���, r_Clinic.��������, r_Clinic.��ҩ����,
           r_Clinic.����, -1 * r_Clinic.����, r_Clinic.�Ӱ��־, r_Clinic.���ӱ�־, r_Clinic.������Ŀid, r_Clinic.�վݷ�Ŀ, r_Clinic.��׼����,
           -1 * r_Clinic.Ӧ�ս��, -1 * r_Clinic.ʵ�ս��, -1 * r_Clinic.ͳ����, 0, r_Clinic.��������id, r_Clinic.������, r_Clinic.����ʱ��,
           �˷�ʱ��_In, r_Clinic.ִ�в���id, r_Clinic.������, ����Ա���_In, ����Ա����_In, r_Clinic.���ʵ�id, '', r_Clinic.�Ƿ���, n_��id, n_����id,
           -1 * r_Clinic.ʵ�ս��, -1, 0);
      End Loop;
    Else
      --4.�˿�תԤ��(������Ʊ��,�ɲ���Աͨ���ش����)
    
      For r_Pay In (Select Min(a.Id) As Ԥ��id, a.���㷽ʽ, Sum(a.��Ԥ��) As ��Ԥ��, 2 As Ԥ�����, a.�����id, a.���㿨���, a.����, a.������ˮ��,
                           a.����˵��, a.������λ, b.����
                    From ����Ԥ����¼ A, ���㷽ʽ B
                    Where a.��¼���� = 3 And a.��¼״̬ In (2, 3) And
                          a.����id In (Select Column_Value From Table(f_Str2list(v_����ids))) And a.���㷽ʽ = b.���� And
                          b.���� In (1, 2, 3, 4, 7, 8) And a.���㷽ʽ Is Not Null
                    Group By a.���㷽ʽ, Ԥ�����, a.�����id, a.���㿨���, a.����, b.����, a.������ˮ��, a.����˵��, a.������λ


                    
                    Having Sum(a.��Ԥ��) <> 0) Loop
        --4.1����Ԥ����� (�����ڲ����˷ѵ����)
        --���е���,����������Ԥ�����
        --��Ϊ�տ�������ɿ�,������Ա�ɿ�����ޱ仯
        If r_Pay.���� = 7 Or (r_Pay.���� = 8 And r_Pay.�����id Is Not Null) Then
          Update ����Ԥ����¼
          Set ��Ԥ�� = ��Ԥ�� + (-1 * r_Pay.��Ԥ��), ժҪ = ժҪ || '1' || ',' || r_Pay.�����id || ',' || -1 * r_Pay.��Ԥ�� || '|'
          Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ Is Null;
          If Sql%RowCount = 0 Then
            Insert Into ����Ԥ����¼
              (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ,
               �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��������, У�Ա�־)
            Values
              (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * r_Pay.��Ԥ��, Null, Null, �˷�ʱ��_In,
               Null, Null, Null, ����Ա���_In, ����Ա����_In, '1' || ',' || r_Pay.�����id || ',' || -1 * r_Pay.��Ԥ�� || '|', n_��id,
               Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 3, 1);
          End If;
          n_����״̬ := 1;
        Else
          If r_Pay.���� In (3, 4) Or (r_Pay.���� = 8 And r_Pay.���㿨��� Is Not Null) Then
            v_���㷽ʽ := r_Pay.���㷽ʽ;
          Else
            Begin
              Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And ���� Like '%�ֽ�%' And Rownum < 2;
            Exception
              When Others Then
                Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And Rownum < 2;
            End;
          End If;
        
          If r_Pay.���� = 8 Then
            --Zl_Square_Update(v_����ids, n_����id, n_��id, �˷�ʱ��_In, -1 * n_����id, Null, r_Pay.��Ԥ��, r_Pay.���㿨���);
            Update ����Ԥ����¼
            Set ��Ԥ�� = ��Ԥ�� + (-1 * r_Pay.��Ԥ��), ժҪ = ժҪ || '0' || ',' || r_Pay.���㿨��� || ',' || -1 * r_Pay.��Ԥ�� || '|'
            Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ Is Null;
            If Sql%RowCount = 0 Then
              Insert Into ����Ԥ����¼
                (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ,
                 �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, ��������, У�Ա�־)
              Values
                (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * r_Pay.��Ԥ��, Null, Null, �˷�ʱ��_In,
                 Null, Null, Null, ����Ա���_In, ����Ա����_In, '0' || ',' || r_Pay.���㿨��� || ',' || -1 * r_Pay.��Ԥ�� || '|', n_��id,
                 Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 3, 1);
            End If;
            n_����״̬ := 1;
          End If;
          If r_Pay.���� Not In (3, 4, 7, 8) Then
            Update ����Ԥ����¼
            Set ��� = ��� + r_Pay.��Ԥ��
            Where ��¼���� = 1 And ��¼״̬ = 1 And �տ�ʱ�� = �˷�ʱ��_In And ����id + 0 = n_����id And ���㷽ʽ = v_���㷽ʽ;
            If Sql%RowCount = 0 Then
              v_Ԥ��no := Nextno(11);
              Insert Into ����Ԥ����¼
                (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ,
                 �ɿ���id, Ԥ�����)
              Values
                (����Ԥ����¼_Id.Nextval, v_Ԥ��no, Null, 1, 1, n_����id, ��ҳid_In, ��Ժ����id_In, r_Pay.��Ԥ��, v_���㷽ʽ, Null, �˷�ʱ��_In,
                 Null, Null, Null, ����Ա���_In, ����Ա����_In, '����תסԺԤ��', n_��id, r_Pay.Ԥ�����);
            End If;
          
            --�������
            Update �������
            Set Ԥ����� = Nvl(Ԥ�����, 0) + r_Pay.��Ԥ��
            Where ���� = 1 And ����id = n_����id And ���� = 2
            Returning Ԥ����� Into n_����ֵ;
            If Sql%RowCount = 0 Then
              Insert Into ������� (����id, ����, ����, Ԥ�����, �������) Values (n_����id, 1, 2, r_Pay.��Ԥ��, 0);
              n_����ֵ := r_Pay.��Ԥ��;
            End If;
            If Nvl(n_����ֵ, 0) = 0 Then
              Delete From �������
              Where ����id = n_����id And ���� = 1 And Nvl(Ԥ�����, 0) = 0 And Nvl(�������, 0) = 0;
            End If;
          End If;
          --4.2�ɿ����ݴ���
          --   ��Ϊû��ʵ���ղ��˵�Ǯ,���Բ�����
          --�����˷��������ԭԤ����¼
          If r_Pay.���� In (3, 4) Then
            Update ��Ա�ɿ����
            Set ��� = Nvl(���, 0) - r_Pay.��Ԥ��
            Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = r_Pay.���㷽ʽ
            Returning ��� Into n_����ֵ;
            If Sql%RowCount = 0 Then
              Insert Into ��Ա�ɿ����
                (�տ�Ա, ���㷽ʽ, ����, ���)
              Values
                (����Ա����_In, r_Pay.���㷽ʽ, 1, -1 * r_Pay.��Ԥ��);
              n_����ֵ := r_Pay.��Ԥ��;
            End If;
            If Nvl(n_����ֵ, 0) = 0 Then
              Delete From ��Ա�ɿ����
              Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = r_Pay.���㷽ʽ And Nvl(���, 0) = 0;
            End If;
          End If;
        
          If r_Pay.���㿨��� Is Null Then
            Update ����Ԥ����¼
            Set ��Ԥ�� = ��Ԥ�� + (-1 * r_Pay.��Ԥ��)
            Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_���㷽ʽ;
            If Sql%RowCount = 0 Then
              Insert Into ����Ԥ����¼
                (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ,
                 �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, У�Ա�־, ��������)
              Values
                (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, -1 * r_Pay.��Ԥ��, v_���㷽ʽ, Null, �˷�ʱ��_In,
                 Null, Null, Null, ����Ա���_In, ����Ա����_In, '', n_��id, r_Pay.�����id, r_Pay.���㿨���, r_Pay.����, r_Pay.������ˮ��,
                 r_Pay.����˵��, r_Pay.������λ, n_����id, -1 * n_����id, 0, 3);
            End If;
          End If;
        End If;
      End Loop;
    End If;
    If ����_In Is Not Null Then
      Begin
        Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And ���� Like '%�ֽ�%' And Rownum < 2;
      Exception
        When Others Then
          Select ���� Into v_���㷽ʽ From ���㷽ʽ Where ���� = 1 And Rownum < 2;
      End;
      Update ����Ԥ����¼
      Set ��Ԥ�� = ��Ԥ�� - ����_In
      Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_���㷽ʽ;
      Update ����Ԥ����¼
      Set ��Ԥ�� = ��Ԥ�� + ����_In
      Where ��¼���� = 3 And ��¼״̬ = 2 And ����id = n_����id And ���㷽ʽ = v_����;
      If Sql%RowCount = 0 Then
        Insert Into ����Ԥ����¼
          (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ��Ԥ��, ���㷽ʽ, �������, �տ�ʱ��, �ɿλ, ��λ������, ��λ�ʺ�, ����Ա���, ����Ա����, ժҪ,
           �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, ����id, �������, У�Ա�־, ��������)
        Values
          (����Ԥ����¼_Id.Nextval, Null, Null, 3, 2, n_����id, ��ҳid_In, ��Ժ����id_In, ����_In, v_����, Null, �˷�ʱ��_In, Null, Null,
           Null, ����Ա���_In, ����Ա����_In, '', n_��id, Null, Null, Null, Null, Null, Null, n_����id, -1 * n_����id, 0, 3);
      End If;
    End If;
    Delete From ����Ԥ����¼ Where ����id = n_ԭ����id And ժҪ = 'Ԥ����ʱ��¼' And ��¼���� = 3;
    Delete From ����Ԥ����¼
    Where ����id = n_����id And ��¼���� = 3 And ��¼״̬ = 2 And ��Ԥ�� = 0 And ���㷽ʽ Is Not Null;
    Update ������ü�¼
    Set ����״̬ = Nvl(n_����״̬, 0)
    Where NO In (Select Column_Value From Table(f_Str2list(v_Nos))) And Mod(��¼����, 10) = 1 And ��¼״̬ = 2;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����תסԺ_�շ�ת��;
/

--105165:�ŵ���,2017-04-20,������ҩ������ӸýкŴ��ڵ����д���
CREATE OR REPLACE Procedure zl_��ҩ����_Insert (
    ����_IN IN ��ҩ����.����%Type, 
    ����_IN IN ��ҩ����.����%Type, 
    �ϰ�_IN IN ��ҩ����.�ϰ��%Type, 
    ҩ��ID_IN IN ��ҩ����.ҩ��ID%Type, 
    ר��_IN IN ��ҩ����.ר��%Type,
    �кŴ���_IN In ��ҩ����.�кŴ���%type

) 
IS 
    Msg VARCHAR2 (30); 
Begin 
    Insert INTO ��ҩ���� 
                    (����, ����, �ϰ��, ҩ��ID, ר��,�кŴ���) 
          VALUES (����_IN, ����_IN, �ϰ�_IN, ҩ��ID_IN, ר��_IN,�кŴ���_IN); 
Exception 
    When Others Then 
        Zl_ErrorCenter (SQLCODE, SQLERRM); 
End zl_��ҩ����_Insert;
/

--105165:�ŵ���,2017-04-20,������ҩ������ӸýкŴ��ڵ����д���
CREATE OR REPLACE Procedure zl_��ҩ����_UPDATE (
    ����_IN IN ��ҩ����.����%Type,
    ����_IN IN ��ҩ����.����%Type,
    ҩ��ID_IN IN ��ҩ����.ҩ��ID%Type,
    ר��_IN IN ��ҩ����.ר��%Type,
    Old����_IN IN ��ҩ����.����%Type,
    Oldҩ��ID_IN IN ��ҩ����.ҩ��ID%Type,
    �кŴ���_IN In ��ҩ����.�кŴ���%type
)
IS
    Msg VARCHAR2 (30);
Begin
    UPDATE ��ҩ����
        SET ���� = ����_IN,
             ���� = ����_IN,
             ҩ��ID = ҩ��ID_IN,
             ר�� = ר��_IN,
             �кŴ���=�кŴ���_IN
     WHERE ���� = Old����_IN
        AND ҩ��ID = Oldҩ��ID_IN;
Exception
    When Others Then
        Zl_ErrorCenter (SQLCODE, SQLERRM);
End zl_��ҩ����_UPDATE;
/

--105165:�ŵ���,2017-04-20,������ҩ������ӸýкŴ��ڵ����д���
CREATE OR REPLACE Procedure Zl_δ��ҩƷ��¼_����
(
  No_In       ҩƷ�շ���¼.NO%Type,
  ����_In     ҩƷ�շ���¼.����%Type,
  ҩ��id_In   ҩƷ�շ���¼.�ⷿid%Type,
  ��ҩ����_In ҩƷ�շ���¼.��ҩ����%Type,
  ��������_In δ��ҩƷ��¼.��������%Type := Null
) Is
Begin
  If ��������_In Is Null Then
    --��������Ϊ��ʱ������ǰ�ĺ���״̬�ĵ��ݵĺ����������
    Update δ��ҩƷ��¼
    Set �������� = Null
    Where �ⷿid = ҩ��id_In And ���� = ����_In and (��ҩ���� = ��ҩ����_In or ��ҩ���� in(select ���� from ��ҩ���� where �кŴ���=��ҩ����_In)) And NO = No_In  And �Ŷ�״̬ = 3 and �������� between sysdate-3 and sysdate;
  Else
    --�������ݲ�Ϊ��ʱ���Ƚ���ǰ�ĺ���״̬�еĵ�������Ϊ�Ѻ��У��ٽ���ǰ��������Ϊ����״̬������д�������ݺͺ���ʱ��
    --��������ͬһ���ݷ������е����
    Update δ��ҩƷ��¼
    Set �Ŷ�״̬ = 4, �������� = Null
    Where �ⷿid = ҩ��id_In And (��ҩ���� = ��ҩ����_In or ��ҩ���� in(select ���� from ��ҩ���� where �кŴ���=��ҩ����_In)) And �Ŷ�״̬ = 3 and �������� between sysdate-3 and sysdate;

    Update δ��ҩƷ��¼
    Set �Ŷ�״̬ = 3, �������� = ��������_In, ����ʱ�� = Sysdate
    Where �ⷿid = ҩ��id_In And ���� = ����_In And NO = No_In and �������� between sysdate-3 and sysdate;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_δ��ҩƷ��¼_����;
/

--108439:��ΰ��,2017-04-19,סԺ�����´�ҽ��ʱȡ��������ʱ�Ӳ�����ҳ����ȡ
CREATE OR REPLACE Procedure Zl_����ҽ����¼_Insert
(
  Id_In           ����ҽ����¼.Id%Type,
  ���id_In       ����ҽ����¼.���id%Type,
  ���_In         ����ҽ����¼.���%Type,
  ������Դ_In     ����ҽ����¼.������Դ%Type,
  ����id_In       ����ҽ����¼.����id%Type,
  ��ҳid_In       ����ҽ����¼.��ҳid%Type,
  Ӥ��_In         ����ҽ����¼.Ӥ��%Type,
  ҽ��״̬_In     ����ҽ����¼.ҽ��״̬%Type,
  ҽ����Ч_In     ����ҽ����¼.ҽ����Ч%Type,
  �������_In     ����ҽ����¼.�������%Type,
  ������Ŀid_In   ����ҽ����¼.������Ŀid%Type,
  �շ�ϸĿid_In   ����ҽ����¼.�շ�ϸĿid%Type,
  ����_In         ����ҽ����¼.����%Type,
  ��������_In     ����ҽ����¼.��������%Type,
  �ܸ�����_In     ����ҽ����¼.�ܸ�����%Type,
  ҽ������_In     ����ҽ����¼.ҽ������%Type,
  ҽ������_In     ����ҽ����¼.ҽ������%Type,
  �걾��λ_In     ����ҽ����¼.�걾��λ%Type,
  ִ��Ƶ��_In     ����ҽ����¼.ִ��Ƶ��%Type,
  Ƶ�ʴ���_In     ����ҽ����¼.Ƶ�ʴ���%Type,
  Ƶ�ʼ��_In     ����ҽ����¼.Ƶ�ʼ��%Type,
  �����λ_In     ����ҽ����¼.�����λ%Type,
  ִ��ʱ�䷽��_In ����ҽ����¼.ִ��ʱ�䷽��%Type,
  �Ƽ�����_In     ����ҽ����¼.�Ƽ�����%Type,
  ִ�п���id_In   ����ҽ����¼.ִ�п���id%Type,
  ִ������_In     ����ҽ����¼.ִ������%Type,
  ������־_In     ����ҽ����¼.������־%Type,
  ��ʼִ��ʱ��_In ����ҽ����¼.��ʼִ��ʱ��%Type,
  ִ����ֹʱ��_In ����ҽ����¼.ִ����ֹʱ��%Type,
  ���˿���id_In   ����ҽ����¼.���˿���id%Type,
  ��������id_In   ����ҽ����¼.��������id%Type,
  ����ҽ��_In     ����ҽ����¼.����ҽ��%Type,
  ����ʱ��_In     ����ҽ����¼.����ʱ��%Type,
  �Һŵ�_In       ����ҽ����¼.�Һŵ�%Type := Null,
  ǰ��id_In       ����ҽ����¼.ǰ��id%Type := Null,
  ��鷽��_In     ����ҽ����¼.��鷽��%Type := Null,
  ִ�б��_In     ����ҽ����¼.ִ�б��%Type := Null,
  �ɷ����_In     ����ҽ����¼.�ɷ����%Type := Null,
  ժҪ_In         ����ҽ����¼.ժҪ%Type := Null,
  ����Ա����_In   ����ҽ��״̬.������Ա%Type := Null,
  ��Ѽ���_In     ����ҽ����¼.��Ѽ���%Type := Null,
  ��ҩĿ��_In     ����ҽ����¼.��ҩĿ��%Type := Null,
  ��ҩ����_In     ����ҽ����¼.��ҩ����%Type := Null,
  ���״̬_In     ����ҽ����¼.���״̬%Type := Null,
  �������_In     ����ҽ����¼.�������%Type := Null,
  ����˵��_In     ����ҽ����¼.����˵��%Type := Null,
  �״�����_In     ����ҽ����¼.�״�����%Type := Null,
  �䷽id_In       ����ҽ����¼.�䷽id%Type := Null,
  �������_In     ����ҽ����¼.�������%Type := Null,
  �����Ŀid_In   ����ҽ����¼.�����Ŀid%Type := Null,
  Ƥ�Խ��_In     ����ҽ����¼.Ƥ�Խ��%Type := Null
  --���ܣ�ҽ����ʿ�¿�,��¼ҽ��ʱ�²�����ҽ����¼�������������סԺ��
) Is
  v_Temp     Varchar2(255);
  v_��Ա���� ����ҽ��״̬.������Ա%Type;

  v_���� ������Ϣ.����%Type;
  v_�Ա� ������Ϣ.�Ա�%Type;
  v_���� ������Ϣ.����%Type;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  --��ǰ������Ա
  If ����Ա����_In Is Not Null Then
    v_��Ա���� := ����Ա����_In;
  Else
    v_Temp     := Zl_Identity;
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
    v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
    v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  End If;

  If Nvl(��ҳid_In, 0) <> 0 Then
    Select ����, �Ա�, ���� Into v_����, v_�Ա�, v_���� From ������ҳ Where ����id = ����id_In And ��ҳid = ��ҳid_In;
  Else
    Select ����, �Ա�, ���� Into v_����, v_�Ա�, v_���� From ������Ϣ Where ����id = ����id_In;
  End If;

  --����ҽ����¼
  Insert Into ����ҽ����¼
    (ID, ���id, ���, ������Դ, ����id, ��ҳid, ����, �Ա�, ����, Ӥ��, ҽ��״̬, ҽ����Ч, �������, ������Ŀid, �շ�ϸĿid, ����, ��������, �ܸ�����, ҽ������, ҽ������, �걾��λ,
     ��鷽��, ִ�б��, ִ��Ƶ��, Ƶ�ʴ���, Ƶ�ʼ��, �����λ, ִ��ʱ�䷽��, �Ƽ�����, ִ�п���id, ִ������, ������־, �ɷ����, ��ʼִ��ʱ��, ִ����ֹʱ��, ���˿���id, ��������id, ����ҽ��,
     ����ʱ��, �Һŵ�, ǰ��id, ժҪ, ��Ѽ���, ����ʱ��, ��ҩĿ��, ��ҩ����, ���״̬, �������, ����˵��, �״�����, �䷽id, �������, �����Ŀid, Ƥ�Խ��)
  Values
    (Id_In, ���id_In, ���_In, ������Դ_In, ����id_In, ��ҳid_In, v_����, v_�Ա�, v_����, Ӥ��_In, ҽ��״̬_In, ҽ����Ч_In, �������_In, ������Ŀid_In,
     �շ�ϸĿid_In, ����_In, ��������_In, �ܸ�����_In, ҽ������_In, ҽ������_In, �걾��λ_In, ��鷽��_In, ִ�б��_In, ִ��Ƶ��_In, Ƶ�ʴ���_In, Ƶ�ʼ��_In, �����λ_In,
     ִ��ʱ�䷽��_In, �Ƽ�����_In, ִ�п���id_In, ִ������_In, ������־_In, �ɷ����_In, ��ʼִ��ʱ��_In, ִ����ֹʱ��_In, ���˿���id_In, ��������id_In, ����ҽ��_In,
     ����ʱ��_In, �Һŵ�_In, ǰ��id_In, ժҪ_In, ��Ѽ���_In,
     Decode(�������_In, 'F', To_Date(�걾��λ_In, 'yyyy-mm-dd hh24:mi:ss'), 'K', To_Date(�걾��λ_In, 'yyyy-mm-dd hh24:mi:ss'),
             Null), ��ҩĿ��_In, ��ҩ����_In, ���״̬_In, �������_In, ����˵��_In, �״�����_In, �䷽id_In, �������_In, �����Ŀid_In, Ƥ�Խ��_In);

  --����ҽ��״̬
  If ҽ��״̬_In <> -1 Then
    Delete From ����ҽ��״̬ Where ҽ��id = Id_In And �������� = 1;
    If Sql%RowCount <> 0 Then
      v_Error := '��ͬID���¿�ҽ���Ѿ����ڡ�';
      Raise Err_Custom;
    End If;
    --��Ϊ����ͬʱ���¿�->�Զ�У��(סԺҽ������)->�����Զ�ֹͣ(סԺҽ����������ֹͣ),��˷ֱ�-2,-1��
    Insert Into ����ҽ��״̬
      (ҽ��id, ��������, ������Ա, ����ʱ��)
    Values
      (Id_In, 1, v_��Ա����, Sysdate - 2 / 60 / 60 / 24);
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ҽ����¼_Insert;
/

--108340:������,2017-04-18,ִ����ز���֮ǰ���ж��Ƿ��շ�
Create Or Replace Procedure Zl_���鱨�浥_Insert
(
  Id_In   In ����ҽ����¼.Id%Type,
  Type_In In Number -- 0=���� 1=ɾ��
) Is
  --HIS������LIS�ӿ�ʹ��
  v_��ҳid     ����ҽ����¼.��ҳid%Type;
  v_ҽ��id     ����ҽ����¼.Id%Type;
  v_��������id ����ҽ����¼.��������id%Type;
  v_������Դ   ����걾��¼.������Դ%Type;
  v_����id     ����걾��¼.����id%Type;
  v_Ӥ��       ����걾��¼.Ӥ��%Type;
  v_�����ļ�id ��������Ӧ��.�����ļ�id%Type;
  v_�����ļ��� �����ļ��б�.����%Type;
  v_�ļ�id     ���Ӳ�������.�ļ�id%Type;
  v_Temp       Varchar2(255);
  v_��Ա����id ������Ա.����id%Type;
  v_��Ա���   ��Ա��.���%Type;
  v_��Ա����   ��Ա��.����%Type;
  v_No         ����ҽ������.No%Type;
  v_����       ����ҽ������.��¼����%Type;
  v_���       Varchar2(1000);
  v_����       Number;
  v_Error      Varchar2(255);
  Err_Custom Exception;
  n_Par Number;
  --���ҵ�ǰ�걾���������
  Cursor c_Samplequest Is
    Select Distinct ID As ҽ��id From ����ҽ����¼ Where Id_In In (ID, ���id);

  --δ��˵ķ�����(������ҩƷ)
  Cursor c_Verify(v_ҽ��id In Number) Is
    Select Distinct ��¼����, NO, ���, ��¼״̬, �����־
    From סԺ���ü�¼
    Where �շ���� Not In ('5', '6', '7') And ҽ����� + 0 In (Select ID From ����ҽ����¼ Where v_ҽ��id In (ID, ���id)) And ���ʷ��� = 1 And
          ��¼״̬ = 0 And �۸񸸺� Is Null And
          (��¼����, NO) In (Select ��¼����, NO
                         From ����ҽ������
                         Where ҽ��id = v_ҽ��id
                         Union All
                         Select ��¼����, NO
                         From ����ҽ������
                         Where ҽ��id In (Select ID From ����ҽ����¼ Where v_ҽ��id In (ID, ���id)))
    Order By ��¼����, NO, ���;

Begin
  --����Ա��Ϣ:����ID,��������;��ԱID,��Ա���,��Ա����
  v_Temp       := Zl_Identity;
  v_��Ա����id := To_Number(Substr(v_Temp, 1, Instr(v_Temp, ',') - 1));

  v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_��Ա��� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);

  Select Distinct Nvl(b.��ҳid, 0), Nvl(b.���id, 0), Decode(b.������Դ, 2, 2, 4, 4, 1), Nvl(b.����id, 0), Nvl(b.��������id, 0),
                  Nvl(b.Ӥ��, 0)
  Into v_��ҳid, v_ҽ��id, v_������Դ, v_����id, v_��������id, v_Ӥ��
  From ����ҽ����¼ B
  Where b.���id = Id_In;
  If v_������Դ = 1 Then
    --��ҳID�� ���ﲡ����Һ�ID
    Select Nvl(Max(b.Id), 0)
    Into v_��ҳid
    From ���˹Һż�¼ B, ����ҽ����¼ A
    Where a.�Һŵ� = b.No(+) And a.Id = Id_In;
  End If;
  Begin
    Select �����ļ�id, c.����
    Into v_�����ļ�id, v_�����ļ���
    From ����ҽ����¼ A, ��������Ӧ�� B, �����ļ��б� C
    Where a.������Ŀid = b.������Ŀid And b.�����ļ�id = c.Id And a.���id = v_ҽ��id And b.Ӧ�ó��� = v_������Դ And Rownum <= 1;
  Exception
    When Others Then
      Return;
  End;

  If Type_In = 0 Then
    --����Ƿ��շ�
    n_Par := Zl_To_Number(Nvl(zl_GetSysParameter(163), '0'));
    If n_Par = 1 Then
      For r_Samplequest In c_Samplequest Loop
        For r_Verify In c_Verify(r_Samplequest.ҽ��id) Loop
          If r_Verify.��¼״̬ = 0 Then
            If r_Verify.�����־ = 1 Then
              v_Error := '�걾δ�շѣ�������ִ�У�����ϵ����Ա��';
              Raise Err_Custom;
            Elsif r_Verify.�����־ = 2 Then
              v_Error := '�걾δ���ˣ�������ִ�У�����ϵ����Ա��';
              Raise Err_Custom;
            End If;
          End If;
        End Loop;
      End Loop;
    End If;
  
    --����
    --ɾ����ǰ�ı����¼
    Begin
      Select Nvl(����id, 0) Into v_�ļ�id From ����ҽ������ Where ҽ��id = v_ҽ��id And Rownum <= 1;
      If v_�ļ�id > 0 Then
        Delete ���Ӳ�����¼ Where ID = v_�ļ�id;
        Delete ���Ӳ������� Where �ļ�id = v_�ļ�id;
      End If;
    Exception
      When Others Then
        Select ���Ӳ�����¼_Id.Nextval Into v_�ļ�id From Dual;
        --Insert Into ����ҽ������ (ҽ��id, ����id) Values (v_ҽ��id, v_�ļ�id);
    End;
  
    Insert Into ���Ӳ�����¼
      (ID, ������Դ, ����id, ��ҳid, Ӥ��, ����id, ��������, �ļ�id, ��������, ������, ����ʱ��, ������, ����ʱ��, ���汾, ǩ������)
    Values
      (v_�ļ�id, v_������Դ, v_����id, v_��ҳid, v_Ӥ��, v_��������id, 7, v_�����ļ�id, v_�����ļ���, Null, Sysdate, Null, Sysdate, 1, 0);
  
    Insert Into ����ҽ������ (ҽ��id, ����id) Values (v_ҽ��id, v_�ļ�id);
  
    Insert Into ���Ӳ�������
      (ID, �ļ�id, ��ʼ��, ��ֹ��, ��id, �������, ��������, ������, ��������, ��������, �����д�, �����ı�, �Ƿ���)
    Values
      (���Ӳ�������_Id.Nextval, v_�ļ�id, 1, 1, Null, 1, 2, Null, Null, 0, 0, 0, 0);
  
    Update ����ҽ������ Set ִ��״̬ = 1 Where ҽ��id In (Select ID From ����ҽ����¼ Where v_ҽ��id In (ID, ���id));
  
    --2.��鵱ǰ�걾��ص��������ر걾�Ƿ�������
    For r_Samplequest In c_Samplequest Loop
    
      --r_SampleQuest.ҽ��id�����Ѿ����,�����������
    
      --2.����ִ�д���
      If v_���� = 1 Then
        Update ������ü�¼
        Set ִ��״̬ = 1, ִ��ʱ�� = Sysdate, ִ���� = v_��Ա����
        Where �շ���� Not In ('5', '6', '7') And
              (ҽ�����, ��¼����, NO) In
              (Select ҽ��id, ��¼����, NO
               From ����ҽ������
               Where ҽ��id = r_Samplequest.ҽ��id
               Union All
               Select ҽ��id, ��¼����, NO
               From ����ҽ������
               Where ҽ��id In (Select ID From ����ҽ����¼ Where r_Samplequest.ҽ��id In (ID, ���id)));
      Else
        Update סԺ���ü�¼
        Set ִ��״̬ = 1, ִ��ʱ�� = Sysdate, ִ���� = v_��Ա����
        Where �շ���� Not In ('5', '6', '7') And
              (ҽ�����, ��¼����, NO) In
              (Select ҽ��id, ��¼����, NO
               From ����ҽ������
               Where ҽ��id = r_Samplequest.ҽ��id
               Union All
               Select ҽ��id, ��¼����, NO
               From ����ҽ������
               Where ҽ��id In (Select ID From ����ҽ����¼ Where r_Samplequest.ҽ��id In (ID, ���id)));
      End If;
      --3.�Զ���˼���
      For r_Verify In c_Verify(r_Samplequest.ҽ��id) Loop
        If r_Verify.No || ',' || r_Verify.��¼���� <> v_No || ',' || v_���� Then
          If v_��� Is Not Null Then
            If v_���� = 1 Then
              Zl_������ʼ�¼_Verify(v_No, v_��Ա���, v_��Ա����, Substr(v_���, 2));
            Elsif v_���� = 2 Then
              Zl_סԺ���ʼ�¼_Verify(v_No, v_��Ա���, v_��Ա����, Substr(v_���, 2));
            End If;
          End If;
          v_��� := Null;
        End If;
        v_No   := r_Verify.No;
        v_���� := r_Verify.��¼����;
        v_��� := v_��� || ',' || r_Verify.���;
      End Loop;
      If v_��� Is Not Null Then
        If v_���� = 1 Then
          Zl_������ʼ�¼_Verify(v_No, v_��Ա���, v_��Ա����, Substr(v_���, 2));
        Elsif v_���� = 2 Then
          Zl_סԺ���ʼ�¼_Verify(v_No, v_��Ա���, v_��Ա����, Substr(v_���, 2));
        End If;
      End If;
    
    End Loop;
  Else
    --ɾ��
  
    v_���� := 0;
    Select Nvl(����״̬, 0) Into v_���� From ����ҽ������ Where ҽ��id = v_ҽ��id;
    If v_���� = 0 Then
      Select ����id Into v_�ļ�id From ����ҽ������ Where ҽ��id = v_ҽ��id And Rownum <= 1;
      Delete ����ҽ������ Where ҽ��id = v_ҽ��id;
      Delete ���Ӳ�����¼ Where ID = v_�ļ�id;
      Delete ���Ӳ������� Where �ļ�id = v_�ļ�id;
      Update ����ҽ������
      Set ִ��״̬ = 0
      Where ҽ��id In (Select ID From ����ҽ����¼ Where v_ҽ��id In (ID, ���id));
      For r_Samplequest In c_Samplequest Loop
        --2.����ִ�д���
        If v_���� = 1 Then
          Update ������ü�¼
          Set ִ��״̬ = 0, ִ��ʱ�� = Null, ִ���� = Null
          Where �շ���� Not In ('5', '6', '7') And
                (ҽ�����, ��¼����, NO) In
                (Select ҽ��id, ��¼����, NO
                 From ����ҽ������
                 Where ҽ��id = r_Samplequest.ҽ��id
                 Union All
                 Select ҽ��id, ��¼����, NO
                 From ����ҽ������
                 Where ҽ��id In (Select ID From ����ҽ����¼ Where r_Samplequest.ҽ��id In (ID, ���id)));
        Else
          Update סԺ���ü�¼
          Set ִ��״̬ = 0, ִ��ʱ�� = Null, ִ���� = Null
          Where �շ���� Not In ('5', '6', '7') And
                (ҽ�����, ��¼����, NO) In
                (Select ҽ��id, ��¼����, NO
                 From ����ҽ������
                 Where ҽ��id = r_Samplequest.ҽ��id
                 Union All
                 Select ҽ��id, ��¼����, NO
                 From ����ҽ������
                 Where ҽ��id In (Select ID From ����ҽ����¼ Where r_Samplequest.ҽ��id In (ID, ���id)));
        End If;
      End Loop;
    Else
      v_Error := '�ñ����Ѿ���ҽ�����ģ�����ȡ��������ϵҽ����';
      Raise Err_Custom;
    End If;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_���鱨�浥_Insert;
/

--108340:������,2017-04-18,ִ����ز���֮ǰ���ж��Ƿ��շ�
Create Or Replace Procedure Zl_����ҽ������_Sampleinput
(
  ҽ��id      In Varchar2,
  ������_In   In ����ҽ������.������%Type := Null,
  ��������_In In ����ҽ������.��������%Type := 0,
  ��Ա���_In In ��Ա��.���%Type := Null,
  ��Ա����_In In ��Ա��.����%Type := Null,
  �ͼ���_In   In ����ҽ������.�ͼ���%Type := Null
) Is
  --δ��˵ķ�����(������ҩƷ)
  Cursor c_Verify(v_ҽ��id In Number) Is
    Select Distinct ��¼����, NO, ���, ��¼״̬,�����־
    From סԺ���ü�¼
    Where �շ���� Not In ('5', '6', '7') And ҽ����� + 0 = v_ҽ��id And ���ʷ��� = 1 And ��¼״̬ = 0 And �۸񸸺� Is Null And
          (��¼����, NO) In (Select ��¼����, NO
                         From ����ҽ������
                         Where ҽ��id = v_ҽ��id
                         Union All
                         Select ��¼����, NO
                         From ����ҽ������
                         Where ҽ��id In (Select ID From ����ҽ����¼ Where v_ҽ��id In (ID, ���id)))
    Union All
    Select Distinct ��¼����, NO, ���, ��¼״̬, �����־
    From ������ü�¼
    Where �շ���� Not In ('5', '6', '7') And ҽ����� + 0 = v_ҽ��id And ���ʷ��� = 1 And ��¼״̬ = 0 And �۸񸸺� Is Null And
          (��¼����, NO) In (Select ��¼����, NO
                         From ����ҽ������
                         Where ҽ��id = v_ҽ��id
                         Union All
                         Select ��¼����, NO
                         From ����ҽ������
                         Where ҽ��id In (Select ID From ����ҽ����¼ Where v_ҽ��id In (ID, ���id)))
    Order By ��¼����, NO, ���;

  --���ҵ�ǰ�걾���������
  Cursor c_Samplequest(v_ҽ��id In Number) Is
    Select Distinct ID As ҽ��id, ������Դ From ����ҽ����¼ Where v_ҽ��id In (ID, ���id);

  v_ִ�� Number(1);
  v_No   ����ҽ������.No%Type;
  v_���� ����ҽ������.��¼����%Type;
  v_��� Varchar2(1000);

  v_ҽ��id   ����ҽ������.ҽ��id%Type;
  v_���id   ����ҽ����¼.���id%Type;
  v_�������� ����ҽ������.��¼����%Type;
  v_�������� ����ҽ������.��������%Type;
  v_Records  Varchar2(2000);
  v_Currrec  Varchar2(50);
  v_Fields   Varchar2(50);
  v_Count    Number(18);
  v_����id   ����ҽ����¼.����id%Type;
  v_��ҳid   ����ҽ����¼.��ҳid%Type;
  v_�Ƿ��Ժ Number; --0=��Ժ,1=��Ժ
  v_��¼״̬ Number;
  v_������Դ ����ҽ����¼.������Դ%Type;
  v_Date     Date;
  Err_Custom Exception;
  v_Error Varchar2(100);
  n_Par   Number;
Begin
  Select Sysdate Into v_Date From Dual;
  --ִ�к��Զ���˶�Ӧ�ļ��ʻ��۵�(������ҩƷ)
  Select Zl_To_Number(Nvl(zl_GetSysParameter(81), '0')) Into v_ִ�� From Dual;

  v_Records := ҽ��id || '|';

  While v_Records Is Not Null Loop
  
    v_Currrec := Substr(v_Records, 1, Instr(v_Records, '|') - 1);
    v_Fields  := v_Currrec;
    v_ҽ��id  := Substr(v_Fields, 1, Instr(v_Fields, ',') - 1);
    v_���id  := Substr(v_Fields, Instr(v_Fields, ',') + 1);
    If ������_In Is Null Then
      Update ����ҽ������ Set ������ = Null, ����ʱ�� = Null, �������� = Null Where ҽ��id In (v_ҽ��id, v_���id);
      Update ����ҽ������
      Set ִ��״̬ = Decode(��������, Null, 0, 1)
      Where ҽ��id In (Select ID From ����ҽ����¼ Where ID In (v_ҽ��id, v_���id) And ���id Is Null);
      For r_Samplequest In c_Samplequest(v_���id) Loop
        If r_Samplequest.������Դ = 2 Then
          Select Decode(��¼����, 1, 1, Decode(�������, 1, 1, 2))
          Into v_��������
          From ����ҽ������
          Where ҽ��id = r_Samplequest.ҽ��id;
        Else
          v_�������� := 1;
        End If;
        If v_�������� = 2 Then
          --2.����ִ�д���
          Update סԺ���ü�¼
          Set ִ��״̬ = 0, ִ��ʱ�� = Null, ִ���� = ������_In
          Where �շ���� Not In ('5', '6', '7') And
                (ҽ�����, ��¼����, NO) In
                (Select ҽ��id, ��¼����, NO
                 From ����ҽ������
                 Where ҽ��id = r_Samplequest.ҽ��id
                 Union All
                 Select ҽ��id, ��¼����, NO
                 From ����ҽ������
                 Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = r_Samplequest.ҽ��id And ���id Is Not Null)
                 Union All
                 Select ҽ��id, ��¼����, NO
                 From ����ҽ������
                 Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = r_Samplequest.ҽ��id And ���id Is Null) And ������ Is Null);
        Else
          Update ������ü�¼
          Set ִ��״̬ = 0, ִ��ʱ�� = Null, ִ���� = ������_In
          Where �շ���� Not In ('5', '6', '7') And
                (ҽ�����, ��¼����, NO) In
                (Select ҽ��id, ��¼����, NO
                 From ����ҽ������
                 Where ҽ��id = r_Samplequest.ҽ��id
                 Union All
                 Select ҽ��id, ��¼����, NO
                 From ����ҽ������
                 Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = r_Samplequest.ҽ��id And ���id Is Not Null)
                 Union All
                 Select ҽ��id, ��¼����, NO
                 From ����ҽ������
                 Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = r_Samplequest.ҽ��id And ���id Is Null) And ������ Is Null);
        End If;
      End Loop;
    Else
      --�ж��Ƿ��ѳ�Ժ������ѳ�Ժ������ɵǼ�
      Begin
        If v_��ҳid Is Null Then
          Select a.����id, a.��ҳid, a.������Դ
          Into v_����id, v_��ҳid, v_������Դ
          From ����ҽ����¼ A, ������ҳ B
          Where a.����id = b.����id And a.��ҳid = b.��ҳid(+) And a.Id = v_ҽ��id;
        End If;
      Exception
        When Others Then
          v_������Դ := 1;
      End;
      If v_������Դ = 2 Then
        If Nvl(v_��ҳid, 0) > 0 Then
          Select Decode(��Ժ����, Null, 1, 0)
          Into v_�Ƿ��Ժ
          From ������ҳ
          Where ����id = v_����id And ��ҳid = v_��ҳid;
        Else
          v_�Ƿ��Ժ := 0;
        End If;
      
        If v_�Ƿ��Ժ = 0 Then
          --��Ժ�ĲŴ���
          Begin
            Select Nvl(��¼״̬, 0)
            Into v_��¼״̬
            From סԺ���ü�¼
            Where ҽ����� = v_ҽ��id And Nvl(��¼״̬, 0) = 0 And Rownum = 1;
          Exception
            When Others Then
              v_��¼״̬ := 1;
          End;
        
          Select Nvl(��������, 0) Into v_�������� From ����ҽ������ Where ҽ��id = v_ҽ��id;
          If v_�������� = 0 Then
            v_Error := '�����ѳ�Ժ������ɵǼ�!';
            Raise Err_Custom;
          End If;
        End If;
      End If;
    
      --���ҽ���Ƿ��շ�
      n_Par := Zl_To_Number(Nvl(zl_GetSysParameter(163), '0'));
      If n_Par = 1 Then
        For r_Samplequest In c_Samplequest(v_���id) Loop
          For r_Verify In c_Verify(r_Samplequest.ҽ��id) Loop
            If r_Verify.��¼״̬ = 0 Then
              If r_Verify.�����־ = 1 Then
                v_Error := '�걾δ�շѣ�������ִ�У�����ϵ����Ա��';
                Raise Err_Custom;
              Elsif r_Verify.�����־ = 2 Then
                v_Error := '�걾δ���ˣ�������ִ�У�����ϵ����Ա��';
                Raise Err_Custom;
              End If;
            End If;
          End Loop;
        End Loop;
      End If;
    
      Update ����ҽ������
      Set ������ = ������_In, ����ʱ�� = v_Date, �������� = ��������_In, �زɱ걾 = Null, �ͼ��� = �ͼ���_In
      Where ҽ��id In (v_ҽ��id, v_���id);
      Update ����ҽ������
      Set ִ��״̬ = 1
      Where ҽ��id In (Select ID From ����ҽ����¼ Where ID In (v_ҽ��id, v_���id) And ���id Is Null);
      --���ʻ��۵��Ƿ�תΪ���ʵ�
      --2.��鵱ǰ�걾��ص��������ر걾�Ƿ�������
      For r_Samplequest In c_Samplequest(v_���id) Loop
        v_Count := 0;
        --r_SampleQuest.ҽ��id�����Ѿ����,�����������
        If v_Count = 0 Then
          If r_Samplequest.������Դ = 2 Then
            Select Decode(��¼����, 1, 1, Decode(�������, 1, 1, 2))
            Into v_��������
            From ����ҽ������
            Where ҽ��id = r_Samplequest.ҽ��id;
          Else
            v_�������� := 1;
          End If;
          If v_�������� = 2 Then
            --2.����ִ�д���
            Update סԺ���ü�¼
            Set ִ��״̬ = 1, ִ��ʱ�� = Sysdate, ִ���� = ��Ա����_In
            Where �շ���� Not In ('5', '6', '7') And
                  (ҽ�����, ��¼����, NO) In
                  (Select ҽ��id, ��¼����, NO
                   From ����ҽ������
                   Where ҽ��id = r_Samplequest.ҽ��id
                   Union All
                   Select ҽ��id, ��¼����, NO
                   From ����ҽ������
                   Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = r_Samplequest.ҽ��id And ���id Is Not Null)
                   Union All
                   Select ҽ��id, ��¼����, NO
                   From ����ҽ������
                   Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = r_Samplequest.ҽ��id And ���id Is Null) And ������ Is Null);
          Else
            Update ������ü�¼
            Set ִ��״̬ = 1, ִ��ʱ�� = Sysdate, ִ���� = ��Ա����_In
            Where �շ���� Not In ('5', '6', '7') And
                  (ҽ�����, ��¼����, NO) In
                  (Select ҽ��id, ��¼����, NO
                   From ����ҽ������
                   Where ҽ��id = r_Samplequest.ҽ��id
                   Union All
                   Select ҽ��id, ��¼����, NO
                   From ����ҽ������
                   Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = r_Samplequest.ҽ��id And ���id Is Not Null)
                   Union All
                   Select ҽ��id, ��¼����, NO
                   From ����ҽ������
                   Where ҽ��id In (Select ID From ����ҽ����¼ Where ID = r_Samplequest.ҽ��id And ���id Is Null) And ������ Is Null);
          End If;
          --3.�Զ���˼���
          If v_ִ�� = 1 Then
            For r_Verify In c_Verify(r_Samplequest.ҽ��id) Loop
              If r_Verify.No || ',' || r_Verify.��¼���� <> v_No || ',' || v_���� Then
                If v_��� Is Not Null Then
                  If v_�������� = 1 Then
                    Zl_������ʼ�¼_Verify(v_No, ��Ա���_In, ��Ա����_In, Substr(v_���, 2));
                  Elsif v_�������� = 2 Then
                    Zl_סԺ���ʼ�¼_Verify(v_No, ��Ա���_In, ��Ա����_In, Substr(v_���, 2));
                  End If;
                End If;
                v_��� := Null;
              End If;
              v_No   := r_Verify.No;
              v_���� := r_Verify.��¼����;
              v_��� := v_��� || ',' || r_Verify.���;
            End Loop;
            If v_��� Is Not Null Then
              If v_�������� = 1 Then
                Zl_������ʼ�¼_Verify(v_No, ��Ա���_In, ��Ա����_In, Substr(v_���, 2));
              Elsif v_�������� = 2 Then
                Zl_סԺ���ʼ�¼_Verify(v_No, ��Ա���_In, ��Ա����_In, Substr(v_���, 2));
              End If;
            End If;
          End If;
        End If;
      End Loop;
    End If;
    v_Records := Substr('|' || v_Records, Length('|' || v_Currrec || '|') + 1);
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����ҽ������_Sampleinput;
/

--108340:������,2017-04-18,ִ����ز���֮ǰ���ж��Ƿ��շ�
Create Or Replace Procedure Zl_����Ԥ������_�ɼ����
(
  ҽ������_In Varchar2, --���ݰ������ҽ��IDʹ��","�ָ� 
  ��Ա���_In ��Ա��.���%Type := Null,
  ��Ա����_In ��Ա��.����%Type := Null --Null=ȡ������Ϊ��ʱ��ɲɼ� 
) Is
  --���ҵ�ǰ�걾��������� 
  Cursor c_Samplequest(v_ҽ��id In Varchar2) Is
    Select /*+ rule */
    Distinct ID As ҽ��id, ������Դ
    From ����ҽ����¼ A, ����ҽ������ B
    Where a.Id = b.ҽ��id And b.������ Is Null And ���id Is Null And
          a.Id In (Select * From Table(Cast(f_Num2list(v_ҽ��id) As Zltools.t_Numlist)));

  --δ��˵ķ�����(������ҩƷ) 
  Cursor c_Verify(v_ҽ��id In Varchar2) Is
    Select /*+ rule */
    Distinct ��¼����, NO, ���, ��¼״̬, �����־
    From סԺ���ü�¼
    Where �շ���� Not In ('5', '6', '7') And
          ҽ����� + 0 In
          (Select ID
           From ����ҽ����¼
           Where ID In (Select * From Table(Cast(f_Num2list(v_ҽ��id) As Zltools.t_Numlist))) And ���id Is Null) And
          ���ʷ��� = 1 And ��¼״̬ = 0 And �۸񸸺� Is Null And
          (��¼����, NO) In (Select ��¼����, NO
                         From ����ҽ������
                         Where ҽ��id In (Select * From Table(Cast(f_Num2list(v_ҽ��id) As Zltools.t_Numlist)))
                         Union All
                         Select ��¼����, NO
                         From ����ҽ������
                         Where ҽ��id In (Select ID
                                        From ����ҽ����¼
                                        Where ID In (Select * From Table(Cast(f_Num2list(v_ҽ��id) As Zltools.t_Numlist))) And
                                              ���id Is Null) And ������ Is Null)
    Union All
    Select /*+ rule */
    Distinct ��¼����, NO, ���, ��¼״̬, �����־
    From ������ü�¼
    Where �շ���� Not In ('5', '6', '7') And
          ҽ����� + 0 In
          (Select ID
           From ����ҽ����¼
           Where ID In (Select * From Table(Cast(f_Num2list(v_ҽ��id) As Zltools.t_Numlist))) And ���id Is Null) And
          ���ʷ��� = 1 And ��¼״̬ = 0 And �۸񸸺� Is Null And
          (��¼����, NO) In (Select ��¼����, NO
                         From ����ҽ������
                         Where ҽ��id In (Select * From Table(Cast(f_Num2list(v_ҽ��id) As Zltools.t_Numlist)))
                         Union All
                         Select ��¼����, NO
                         From ����ҽ������
                         Where ҽ��id In (Select ID
                                        From ����ҽ����¼
                                        Where ID In (Select * From Table(Cast(f_Num2list(v_ҽ��id) As Zltools.t_Numlist))) And
                                              ���id Is Null) And ������ Is Null)
    Order By ��¼����, NO, ���;

  v_����걾��¼ Number(18);
  v_ִ��״̬     Number(1);
  v_������       Varchar2(50);
  v_Error        Varchar2(100);
  V_ִ��         Number;
  v_No           ����ҽ������.No%Type;
  v_����         ����ҽ������.��¼����%Type;
  v_���         Varchar2(1000);
  Err_Custom Exception;
  n_Par Number;
Begin

  If ��Ա����_In Is Not Null Then
    --���걾�Ƿ񱻺��ջ���� 
    Begin
      Select /*+ rule */
       Nvl(c.Id, 0), b.ִ��״̬, b.������
      Into v_����걾��¼, v_ִ��״̬, v_������
      From ����ҽ����¼ A, ����ҽ������ B, ����걾��¼ C
      Where a.Id = b.ҽ��id And a.���id = c.ҽ��id(+) And
            a.Id In (Select * From Table(Cast(f_Num2list(ҽ������_In) As Zltools.t_Numlist)));
    Exception
      When Others Then
        v_����걾��¼ := 0;
    End;
  
    If v_����걾��¼ <> 0 Then
      v_Error := '�걾�ѱ�����ƺ��ղ�����ɲɼ�!';
      Raise Err_Custom;
    End If;
  
    If v_ִ��״̬ <> 2 And v_������ Is Not Null Then
      v_Error := '�걾�ѱ������ǩ�ղ�����ɲɼ�!';
      Raise Err_Custom;
    End If;
  
    --���ҽ���Ƿ��շ�
    n_Par := Zl_To_Number(Nvl(zl_GetSysParameter(163), '0'));
    If n_Par = 1 Then
      For r_Verify In c_Verify(ҽ������_In) Loop
        If r_Verify.��¼״̬ = 0 Then
          If r_Verify.�����־ = 1 Then
            v_Error := '�걾δ�շѣ�������ִ�У�����ϵ����Ա��';
            Raise Err_Custom;
          Elsif r_Verify.�����־ = 2 Then
            v_Error := '�걾δ���ˣ�������ִ�У�����ϵ����Ա��';
            Raise Err_Custom;
          End If;
        End If;
      End Loop;
    End If;
  
    Update /*+ rule */ ������ռ�¼
    Set �ز��� = ��Ա����_In, �ز�ʱ�� = Sysdate
    Where ҽ��id In (Select * From Table(Cast(f_Num2list(ҽ������_In) As Zltools.t_Numlist)));
  
    --���²ɼ���Ϣ(����Ͳɼ��� 
    Update /*+ rule */ ����ҽ������
    Set ������ = ��Ա����_In, ����ʱ�� = Sysdate, ִ��״̬ = Decode(ִ��״̬, 2, 0, ִ��״̬),
        �زɱ걾 = Decode(Nvl(�زɱ걾, 0), 0, Decode(ִ��״̬, 2, 1, 0), �زɱ걾), ִ��˵�� = Null
    Where ҽ��id In (Select * From Table(Cast(f_Num2list(ҽ������_In) As Zltools.t_Numlist)));
  
    --����ҽ���ͷ��ü�¼ 
    For r_Samplequest In c_Samplequest(ҽ������_In) Loop
      If r_Samplequest.������Դ = 2 Then
        --2.����ִ�д��� 
        Update סԺ���ü�¼
        Set ִ��״̬ = 1, ִ��ʱ�� = Sysdate, ִ���� = ��Ա����_In
        Where �շ���� Not In ('5', '6', '7') And
              (ҽ�����, ��¼����, NO) In (Select ҽ��id, ��¼����, NO
                                   From ����ҽ������
                                   Where ҽ��id = r_Samplequest.ҽ��id
                                   Union All
                                   Select ҽ��id, ��¼����, NO
                                   From ����ҽ������
                                   Where ҽ��id In (Select ID
                                                  From ����ҽ����¼ A, ����ҽ������ B
                                                  Where a.Id = b.ҽ��id And r_Samplequest.ҽ��id In (a.Id) And a.���id Is Null And
                                                        b.ִ��״̬ In (0, 2) And b.������ Is Null));
      Else
        --2.����ִ�д��� 
        Update ������ü�¼
        Set ִ��״̬ = 1, ִ��ʱ�� = Sysdate, ִ���� = ��Ա����_In
        Where �շ���� Not In ('5', '6', '7') And
              (ҽ�����, ��¼����, NO) In (Select ҽ��id, ��¼����, NO
                                   From ����ҽ������
                                   Where ҽ��id = r_Samplequest.ҽ��id
                                   Union All
                                   Select ҽ��id, ��¼����, NO
                                   From ����ҽ������
                                   Where ҽ��id In (Select ID
                                                  From ����ҽ����¼ A, ����ҽ������ B
                                                  Where a.Id = b.ҽ��id And r_Samplequest.ҽ��id In (a.Id) And a.���id Is Null And
                                                        b.ִ��״̬ In (0, 2) And b.������ Is Null));
      End If;
    End Loop;
  
    --����ִ��״̬(ֻ���²ɼ��� 
    Update /*+ rule */ ����ҽ������
    Set ִ��״̬ = 1, ����� = ��Ա����_In, ���ʱ�� = Sysdate
    Where ҽ��id In
          (Select ID
           From ����ҽ����¼
           Where ID In (Select * From Table(Cast(f_Num2list(ҽ������_In) As Zltools.t_Numlist))) And ���id Is Null);
    --ִ�к��Զ���˶�Ӧ�ļ��ʻ��۵�(������ҩƷ)
    Select Zl_To_Number(Nvl(zl_GetSysParameter(81), '0')) Into v_ִ�� From Dual;
    --3.�Զ���˼��� 
    For r_Verify In c_Verify(ҽ������_In) Loop
      If r_Verify.No || ',' || r_Verify.��¼���� <> v_No || ',' || v_���� Then
        If v_��� Is Not Null Then
          If v_���� = 1 Then
            Zl_������ʼ�¼_Verify(v_No, ��Ա���_In, ��Ա����_In, Substr(v_���, 2));
          Elsif v_���� = 2 Then
            Zl_סԺ���ʼ�¼_Verify(v_No, ��Ա���_In, ��Ա����_In, Substr(v_���, 2));
          End If;
        End If;
        v_��� := Null;
      End If;
      v_No   := r_Verify.No;
      v_���� := r_Verify.��¼����;
      v_��� := v_��� || ',' || r_Verify.���;
    End Loop;
    If v_��� Is Not Null Then
      If v_���� = 1 Then
        Zl_������ʼ�¼_Verify(v_No, ��Ա���_In, ��Ա����_In, Substr(v_���, 2));
      Elsif v_���� = 2 Then
        Zl_סԺ���ʼ�¼_Verify(v_No, ��Ա���_In, ��Ա����_In, Substr(v_���, 2));
      End If;
    End If;
  
  Else
    --���걾�Ƿ񱻺��ջ���� 
    Begin
      Select /*+ rule */
       Nvl(c.Id, 0), b.ִ��״̬, b.������
      Into v_����걾��¼, v_ִ��״̬, v_������
      From ����ҽ����¼ A, ����ҽ������ B, ����걾��¼ C
      Where a.Id = b.ҽ��id And a.���id = c.ҽ��id(+) And
            a.Id In (Select * From Table(Cast(f_Num2list(ҽ������_In) As Zltools.t_Numlist)));
    Exception
      When Others Then
        v_����걾��¼ := 0;
    End;
  
    If v_����걾��¼ <> 0 Then
      v_Error := '�걾�ѱ�����ƺ��ղ���ȡ����ɲɼ�!';
      Raise Err_Custom;
    End If;
  
    If v_ִ��״̬ <> 2 And v_������ Is Not Null Then
      v_Error := '�걾�ѱ������ǩ�ղ���ȡ����ɲɼ�!';
      Raise Err_Custom;
    End If;
  
    Update /*+ rule */ ����ҽ������
    Set ������ = Null, ����ʱ�� = Null, ִ��״̬ = 0, ִ��˵�� = Null, ����� = Null, ���ʱ�� = Null
    Where ҽ��id In (Select ID
                   From ����ҽ����¼
                   Where ID In (Select * From Table(Cast(f_Num2list(ҽ������_In) As Zltools.t_Numlist))));
  
    For r_Samplequest In c_Samplequest(ҽ������_In) Loop
    
      If r_Samplequest.������Դ = 2 Then
        --2.����ִ�д��� 
        Update סԺ���ü�¼
        Set ִ��״̬ = 0, ִ��ʱ�� = Null, ִ���� = ��Ա����_In
        Where �շ���� Not In ('5', '6', '7') And
              (ҽ�����, ��¼����, NO) In
              (Select ҽ��id, ��¼����, NO
               From ����ҽ������
               Where ҽ��id = r_Samplequest.ҽ��id
               Union All
               Select ҽ��id, ��¼����, NO
               From ����ҽ������
               Where ҽ��id In (Select ID From ����ҽ����¼ Where r_Samplequest.ҽ��id In (ID) And ���id Is Null) And
                     ִ��״̬ In (0, 2) And ������ Is Null);
      Else
        Update ������ü�¼
        Set ִ��״̬ = 0, ִ��ʱ�� = Null, ִ���� = ��Ա����_In
        Where �շ���� Not In ('5', '6', '7') And
              (ҽ�����, ��¼����, NO) In
              (Select ҽ��id, ��¼����, NO
               From ����ҽ������
               Where ҽ��id = r_Samplequest.ҽ��id
               Union All
               Select ҽ��id, ��¼����, NO
               From ����ҽ������
               Where ҽ��id In (Select ID From ����ҽ����¼ Where r_Samplequest.ҽ��id In (ID) And ���id Is Null) And
                     ִ��״̬ In (0, 2) And ������ Is Null);
      End If;
    End Loop;
  End If;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����Ԥ������_�ɼ����;
/

--108340:������,2017-04-18,ִ����ز���֮ǰ���ж��Ƿ��շ�
Create Or Replace Procedure Zl_����걾��¼_�������
(
  Id_In       ����걾��¼.Id%Type,
  �����_In   ����걾��¼.�����%Type := Null,
  ��Ա���_In ��Ա��.���%Type := Null,
  ��Ա����_In ��Ա��.����%Type := Null
) Is

  --δ��˵ķ�����(������ҩƷ) 
  Cursor c_Verify(v_ҽ��id In Number) Is
    Select Distinct 2 As ��¼����, NO, ���, ��¼״̬, �����־
    From סԺ���ü�¼
    Where �շ���� Not In ('5', '6', '7') And ���ʷ��� = 1 And �۸񸸺� Is Null And
          (��¼����, NO) In (Select ��¼����, NO
                         From ����ҽ������
                         Where ҽ��id = v_ҽ��id
                         Union All
                         Select ��¼����, NO
                         From ����ҽ������
                         Where ҽ��id In (Select ID From ����ҽ����¼ Where v_ҽ��id In (ID, ���id))) And ҽ����� = v_ҽ��id
    Union All
    Select Distinct 1 As ��¼����, NO, ���, ��¼״̬,�����־
    From ������ü�¼
    Where �շ���� Not In ('5', '6', '7') And ���ʷ��� = 1 And �۸񸸺� Is Null And
          (��¼����, NO) In (Select ��¼����, NO
                         From ����ҽ������
                         Where ҽ��id = v_ҽ��id
                         Union All
                         Select ��¼����, NO
                         From ����ҽ������
                         Where ҽ��id In (Select ID From ����ҽ����¼ Where v_ҽ��id In (ID, ���id))) And ҽ����� = v_ҽ��id
    Order By ��¼����, NO, ���;

  --���ҵ�ǰ�걾��������� 
  Cursor c_Samplequest(v_΢���� In Number) Is
    Select Distinct ҽ��id, ������Դ
    From (Select a.ҽ��id, b.������Դ
           From ������Ŀ�ֲ� A, ����걾��¼ B
           Where 0 = v_΢���� And a.�걾id = Id_In And a.ҽ��id Is Not Null And a.�걾id = b.Id
           Union
           Select a.ҽ��id, b.������Դ
           From ������Ŀ�ֲ� A, ����걾��¼ B
           Where 1 = v_΢���� And a.�걾id = Id_In And a.ҽ��id Is Not Null And a.�걾id = b.Id
           Union
           Select b.Id As ҽ��id, a.������Դ
           From ����걾��¼ A, ����ҽ����¼ B
           Where a.Id = Id_In And a.ҽ��id = b.���id);

  Cursor c_Stuff
  (
    v_No     Varchar2,
    v_��ҳid Number
  ) Is
    Select NO, ����, �ⷿid
    From δ��ҩƷ��¼
    Where NO = v_No And ���� In (24, 25, 26) And �ⷿid Is Not Null And Not Exists
     (Select 1 From Dual Where zl_GetSysParameter(Decode(v_��ҳid, Null, 92, 63)) = '1') And Exists
     (Select a.���
           From סԺ���ü�¼ A, �������� B
           Where a.��¼���� = 2 And a.��¼״̬ = 1 And a.No = v_No And a.�շ�ϸĿid = b.����id And b.�������� = 1
           Union All
           Select a.���
           From ������ü�¼ A, �������� B
           Where a.��¼���� = 2 And a.��¼״̬ = 1 And a.No = v_No And a.�շ�ϸĿid = b.����id And b.�������� = 1)
    Order By �ⷿid;

  v_ִ��  Number(1);
  v_No    ����ҽ������.No%Type;
  v_Nonew ����ҽ������.No%Type;
  v_����  ����ҽ������.��¼����%Type;
  v_���  Varchar2(1000);

  v_Count      Number(18);
  v_Counts     Number(18);
  v_΢����걾 Number(1) := 0;
  v_��ҳid     Number(18);
  v_Ӥ��       Number(1);
  v_����       Varchar2(100);
  v_����       Number(18);
  v_Intloop    Number;
  Err_Custom Exception;
  v_Error Varchar2(100);

  n_Par Number;
Begin
  Select Nvl(Ӥ��, 0), ���� Into v_Ӥ��, v_���� From ����걾��¼ Where ID = Id_In;

  --ִ�к��Զ���˶�Ӧ�ļ��ʻ��۵�(������ҩƷ) 
  Select Zl_To_Number(Nvl(zl_GetSysParameter(81), '0')) Into v_ִ�� From Dual;

  v_΢����걾 := 0;
  Begin
    Select 1 Into v_΢����걾 From ����걾��¼ Where ΢����걾 = 1 And ID = Id_In;
  Exception
    When Others Then
      v_΢����걾 := 0;
  End;

  --���ж�ҽ���Ƿ��շ�
  n_Par := Zl_To_Number(Nvl(zl_GetSysParameter(163), '0'));
  If n_Par = 1 Then
    For r_Samplequest In c_Samplequest(v_΢����걾) Loop
      For r_���ҽ�� In (Select ID As ҽ��id From ����ҽ����¼ Where ���id = r_Samplequest.ҽ��id) Loop
        For r_Verify In c_Verify(r_���ҽ��.ҽ��id) Loop
          If r_Verify.��¼״̬ = 0 Then
            If r_Verify.�����־ = 1 Then
              v_Error := '�걾δ�շѣ�������ִ�У�����ϵ����Ա��';
              Raise Err_Custom;
            Elsif r_Verify.�����־ = 2 Then
              v_Error := '�걾δ���ˣ�������ִ�У�����ϵ����Ա��';
              Raise Err_Custom;
            End If;
          End If;
        End Loop;
      End Loop;
    End Loop;
  End If;

  --1.�ñ��걾��״̬������˺�ʱ�� 
  Update ����걾��¼
  Set ����� = Decode(�����_In, Null, ��Ա����_In, �����_In), ���ʱ�� = Sysdate, ����״̬ = 2
  Where ID = Id_In;

  --��¼��˹��� 
  Insert Into ���������¼
    (ID, �걾id, ��������, ����Ա, ����ʱ��)
  Values
    (���������¼_Id.Nextval, Id_In, 0, Decode(�����_In, Null, ��Ա����_In, �����_In), Sysdate);

  --2.��鵱ǰ�걾��ص��������ر걾�Ƿ������� 
  For r_Samplequest In c_Samplequest(v_΢����걾) Loop
  
    v_Count := 0;
  
    If v_΢����걾 = 0 Then
      Begin
        Select Nvl(Count(1), 0)
        Into v_Count
        From ����걾��¼
        Where ����״̬ < 2 And ID In (Select �걾id From ������Ŀ�ֲ� Where ҽ��id = r_Samplequest.ҽ��id);
      Exception
        When Others Then
          v_Count := 0;
      End;
    End If;
  
    --r_SampleQuest.ҽ��id�����Ѿ����,����������� 
    If v_Count = 0 Then
    
      --1.�����뵥��ִ��״̬ 
      Update ����ҽ������
      Set ִ��״̬ = 1, ����� = ��Ա����_In, ���ʱ�� = Sysdate
      Where ҽ��id In (Select ID From ����ҽ����¼ Where r_Samplequest.ҽ��id In (ID, ���id));
    
      Update ����ҽ������
      Set ִ��״̬ = 1, ����� = ��Ա����_In, ���ʱ�� = Sysdate
      Where ҽ��id In (Select ���id
                     From ����ҽ����¼
                     Where ID In (Select ID From ����ҽ����¼ Where r_Samplequest.ҽ��id In (ID, ���id)));
    
      If r_Samplequest.������Դ = 2 Then
        --2.����ִ�д��� 
        Update סԺ���ü�¼
        Set ִ��״̬ = 1, ִ��ʱ�� = Sysdate, ִ���� = ��Ա����_In
        Where �շ���� Not In ('5', '6', '7') And
              (ҽ�����, ��¼����, NO) In
              (Select ҽ��id, ��¼����, NO
               From ����ҽ������
               Where ҽ��id = r_Samplequest.ҽ��id
               Union All
               Select ҽ��id, ��¼����, NO
               From ����ҽ������
               Where ҽ��id In (Select ID From ����ҽ����¼ Where r_Samplequest.ҽ��id In (ID, ���id)));
      Else
        Update ������ü�¼
        Set ִ��״̬ = 1, ִ��ʱ�� = Sysdate, ִ���� = ��Ա����_In
        Where �շ���� Not In ('5', '6', '7') And
              (ҽ�����, ��¼����, NO) In
              (Select ҽ��id, ��¼����, NO
               From ����ҽ������
               Where ҽ��id = r_Samplequest.ҽ��id
               Union All
               Select ҽ��id, ��¼����, NO
               From ����ҽ������
               Where ҽ��id In (Select ID From ����ҽ����¼ Where r_Samplequest.ҽ��id In (ID, ���id)));
      End If;
    
      --3.�Զ���˼��� 
      If v_ִ�� = 1 Then
        Select Count(*) Into v_Counts From ����ҽ����¼ Where ���id = r_Samplequest.ҽ��id;
        If v_Counts > 0 Then
          For r_���ҽ�� In (Select ID As ҽ��id From ����ҽ����¼ Where ���id = r_Samplequest.ҽ��id) Loop
            For r_Verify In c_Verify(r_���ҽ��.ҽ��id) Loop
              If r_Verify.No || ',' || r_Verify.��¼���� <> v_No || ',' || v_���� Then
                If v_��� Is Not Null Then
                  If v_���� = 1 Then
                    Zl_������ʼ�¼_Verify(v_No, ��Ա���_In, ��Ա����_In, Substr(v_���, 2));
                  Elsif v_���� = 2 Then
                    Zl_סԺ���ʼ�¼_Verify(v_No, ��Ա���_In, ��Ա����_In, Substr(v_���, 2));
                  End If;
                End If;
                v_��� := Null;
              End If;
              v_No   := r_Verify.No;
              v_���� := r_Verify.��¼����;
              v_��� := v_��� || ',' || r_Verify.���;
            End Loop;
          End Loop;
        Else
          For r_Verify In c_Verify(r_Samplequest.ҽ��id) Loop
            If r_Verify.No || ',' || r_Verify.��¼���� <> v_No || ',' || v_���� Then
              If v_��� Is Not Null Then
                If v_���� = 1 Then
                  Zl_������ʼ�¼_Verify(v_No, ��Ա���_In, ��Ա����_In, Substr(v_���, 2));
                Elsif v_���� = 2 Then
                  Zl_סԺ���ʼ�¼_Verify(v_No, ��Ա���_In, ��Ա����_In, Substr(v_���, 2));
                End If;
              End If;
              v_��� := Null;
            End If;
            v_No   := r_Verify.No;
            v_���� := r_Verify.��¼����;
            v_��� := v_��� || ',' || r_Verify.���;
          End Loop;
        End If;
        If v_��� Is Not Null Then
          If v_���� = 1 Then
            Zl_������ʼ�¼_Verify(v_No, ��Ա���_In, ��Ա����_In, Substr(v_���, 2));
          Elsif v_���� = 2 Then
            Zl_סԺ���ʼ�¼_Verify(v_No, ��Ա���_In, ��Ա����_In, Substr(v_���, 2));
          End If;
          v_��� := Null;
          --  v_���� := null; 
        End If;
      End If;
    
      --����Լ����ĵ� 
      v_Intloop := 1;
    
      Select ����id Into v_���� From ����걾��¼ Where ID = Id_In;
      For r_�����Լ� In (Select c.����id, c.����
                     From ����ҽ����¼ A, ���鱨����Ŀ B, �����Լ���ϵ C
                     Where a.���id = r_Samplequest.ҽ��id And a.������Ŀid = b.������Ŀid And b.������Ŀid = c.��Ŀid And c.����id = v_����) Loop
        Zl_�����Լ���¼_Insert(r_Samplequest.ҽ��id, v_Intloop, r_�����Լ�.����id, r_�����Լ�.����);
        v_Intloop := v_Intloop + 1;
      End Loop;
      Select Count(*) Into v_Intloop From �����Լ���¼ Where ҽ��id = r_Samplequest.ҽ��id And NO Is Null;
      If v_Intloop > 1 Then
        v_Nonew := Nextno(14);
        Update �����Լ���¼ Set NO = v_Nonew Where ҽ��id = r_Samplequest.ҽ��id;
      End If;
      If v_Nonew Is Not Null Then
      
        Zl_�����Լ���¼_Bill(r_Samplequest.ҽ��id, v_Nonew);
      
        v_��ҳid := Null;
        Select ��ҳid Into v_��ҳid From ����ҽ����¼ A Where ID = r_Samplequest.ҽ��id;
      
        If v_��ҳid Is Null Then
          Zl_������ʼ�¼_Verify(v_Nonew, ��Ա���_In, ��Ա����_In);
        Else
          Zl_סԺ���ʼ�¼_Verify(v_Nonew, ��Ա���_In, ��Ա����_In);
        End If;
      
        --�������û���Զ�����,���Զ�����,���򲻴��� 
        For r_Stuff In c_Stuff(v_Nonew, v_��ҳid) Loop
          Zl_�����շ���¼_��������(r_Stuff.�ⷿid, 25, v_Nonew, ��Ա����_In, ��Ա����_In, ��Ա����_In, 1, Sysdate);
        End Loop;
      End If;
    End If;
  End Loop;
  Begin
    Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
      Using 9, 0 || ',' || Id_In;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����걾��¼_�������;
/

--108340:������,2017-04-18,ִ����ز���֮ǰ���ж��Ƿ��շ�
Create Or Replace Procedure Zl_����걾��¼_�걾����
(
  Id_In         In ����걾��¼.Id%Type,
  ҽ��id_In     In ����걾��¼.ҽ��id%Type,
  ���ҽ��_In   In Varchar2, --���ڸ��¶��ҽ����ִ��״̬ 
  ���Ǳ걾id_In In ����걾��¼.Id%Type := 0, --����ʱָ����һ���걾ʱ����ָ��ı걾 
  �걾���_In   In ����걾��¼.�걾���%Type,
  ����ʱ��_In   In ����걾��¼.����ʱ��%Type,
  ������_In     In ����걾��¼.������%Type,
  ����id_In     In ����걾��¼.����id%Type,
  ����ʱ��_In   In ����걾��¼.����ʱ��%Type,
  �걾��̬_In   In ����걾��¼.�걾��̬%Type,
  ������_In     In ����걾��¼.������%Type := Null,
  ����ʱ��_In   In ����걾��¼.����ʱ��%Type := Null,
  ΢����걾_In In ����걾��¼.΢����걾%Type := Null,
  �걾���_In   In ����걾��¼.�걾���%Type := 0,
  ���鱸ע_In   In ����걾��¼.���鱸ע%Type := Null,
  ����_In       In ����걾��¼.����%Type := Null,
  �Ա�_In       In ����걾��¼.�Ա�%Type := Null,
  ����_In       In ����걾��¼.����%Type := Null,
  No_In         In ����걾��¼.No%Type := Null,
  �걾����_In   In ����걾��¼.�걾����%Type := Null,
  �������id_In In ����걾��¼.�������id%Type := Null,
  ������_In     In ����걾��¼.������%Type := Null,
  ��ʶ��_In     In ����걾��¼.��ʶ��%Type := Null,
  ����_In       In ����걾��¼.����%Type := Null,
  ���˿���_In   In ����걾��¼.���˿���%Type := Null,
  ������Ŀ_In   In ����걾��¼.������Ŀ%Type := Null,
  ��������_In   In ����걾��¼.��������%Type := Null,
  ����id_In     In ����걾��¼.����id%Type := Null,
  ִ�п���_In   In ����걾��¼.ִ�п���id%Type := Null,
  ��Ա���_In   In ��Ա��.���%Type := Null,
  ��Ա����_In   In ��Ա��.����%Type := Null
) Is

  Cursor v_Advice Is
    Select /*+ Rule */
    Distinct a.Id, a.����ʱ��, a.�걾��λ, f.��������, a.ִ�п���id, a.������Ŀid, a.��������id, a.����ҽ��, a.����id, a.������Դ, a.Ӥ��, a.������־ As ����,
             b.�����, b.סԺ��, b.��������, a.�Һŵ�, Decode(c.��ҳid, 0, Null, c.��ҳid) As ��ҳid, d.��������, f.������, f.����ʱ��
    From ����ҽ����¼ A, ����ҽ������ F, ������Ϣ B, ������ҳ C, ������ĿĿ¼ D
    Where a.���id In (Select * From Table(Cast(f_Num2list(���ҽ��_In) As Zltools.t_Numlist))) And a.Id = f.ҽ��id And
          a.����id = b.����id And a.����id = c.����id(+) And a.��ҳid = c.��ҳid(+) And a.������Ŀid = d.Id(+);

  Cursor v_Advice_1 Is
    Select /*+ Rule */
    Distinct b.No As ���ݺ�, a.���id
    From ����ҽ����¼ A, ����ҽ������ B
    Where a.Id = b.ҽ��id And a.���id In (Select * From Table(Cast(f_Num2list(���ҽ��_In) As Zltools.t_Numlist)))
    Union All
    Select /*+ Rule */
    Distinct b.No As ���ݺ�, a.���id
    From ����ҽ����¼ A, ����ҽ������ B
    Where a.Id = b.ҽ��id And a.Id In (Select * From Table(Cast(f_Num2list(���ҽ��_In) As Zltools.t_Numlist)));

  Cursor v_Patient Is
    Select ����id, סԺ��, �����, �������� From ������Ϣ Where ����id = ����id_In;

  --δ��˵ķ�����(������ҩƷ) 
  Cursor c_Verify(v_ҽ��id In Number) Is
    Select /*+ Rule */
    Distinct a.��¼����, a.No, a.���, a.ҽ�����, a.�����־, a.��¼״̬
    From סԺ���ü�¼ A, ����ҽ������ B,
         (Select ID
           From ����ҽ����¼
           Where ID In (Select * From Table(Cast(f_Num2list(���ҽ��_In) As Zltools.t_Numlist)))
           Union All
           Select ID
           From ����ҽ����¼
           Where ���id In (Select * From Table(Cast(f_Num2list(���ҽ��_In) As Zltools.t_Numlist)))) C
    Where a.�շ���� Not In ('5', '6', '7') And a.ҽ����� = c.Id And a.��¼״̬ = 0 And �۸񸸺� Is Null And a.ҽ����� = b.ҽ��id And
          a.��¼���� = b.��¼���� And a.No = b.No And a.���ʷ��� = 1
    Union All
    Select /*+ Rule */
    Distinct a.��¼����, a.No, a.���, a.ҽ�����, a.�����־, a.��¼״̬
    From סԺ���ü�¼ A, ����ҽ������ B,
         (Select ID
           From ����ҽ����¼
           Where ID In (Select * From Table(Cast(f_Num2list(���ҽ��_In) As Zltools.t_Numlist)))
           Union All
           Select ID
           From ����ҽ����¼
           Where ���id In (Select * From Table(Cast(f_Num2list(���ҽ��_In) As Zltools.t_Numlist)))) C
    Where a.�շ���� Not In ('5', '6', '7') And a.ҽ����� = c.Id And a.��¼״̬ = 0 And �۸񸸺� Is Null And a.ҽ����� = b.ҽ��id And
          a.��¼���� = b.��¼���� And a.No = b.No And a.���ʷ��� = 1
    Union All
    Select /*+ Rule */
    Distinct a.��¼����, a.No, a.���, a.ҽ�����, a.�����־, a.��¼״̬
    From ������ü�¼ A, ����ҽ������ B,
         (Select ID
           From ����ҽ����¼
           Where ID In (Select * From Table(Cast(f_Num2list(���ҽ��_In) As Zltools.t_Numlist)))
           Union All
           Select ID
           From ����ҽ����¼
           Where ���id In (Select * From Table(Cast(f_Num2list(���ҽ��_In) As Zltools.t_Numlist)))) C
    Where a.�շ���� Not In ('5', '6', '7') And a.ҽ����� = c.Id And a.��¼״̬ = 0 And �۸񸸺� Is Null And a.ҽ����� = b.ҽ��id And
          a.��¼���� = b.��¼���� And a.No = b.No And a.���ʷ��� = 1
    Union All
    Select /*+ Rule */
    Distinct a.��¼����, a.No, a.���, a.ҽ�����, a.�����־, a.��¼״̬
    From ������ü�¼ A, ����ҽ������ B,
         (Select ID
           From ����ҽ����¼
           Where ID In (Select * From Table(Cast(f_Num2list(���ҽ��_In) As Zltools.t_Numlist)))
           Union All
           Select ID
           From ����ҽ����¼
           Where ���id In (Select * From Table(Cast(f_Num2list(���ҽ��_In) As Zltools.t_Numlist)))) C
    Where a.�շ���� Not In ('5', '6', '7') And a.ҽ����� = c.Id And a.��¼״̬ = 0 And �۸񸸺� Is Null And a.ҽ����� = b.ҽ��id And
          a.��¼���� = b.��¼���� And a.No = b.No And a.���ʷ��� = 1
    Order By ��¼����, NO, ���;

  --���ҵ�ǰ�걾��������� 
  Cursor c_Samplequest(v_΢���� In Number) Is
    Select Distinct ҽ��id, ������Դ
    From (Select Decode(a.ҽ��id, Null, b.ҽ��id, a.ҽ��id) As ҽ��id, b.������Դ
           From ������Ŀ�ֲ� A, ����걾��¼ B
           Where Nvl(v_΢����, 0) = 0 And a.�걾id = b.Id And b.ҽ��id In (Select ҽ��id From ����걾��¼ Where ID = Id_In) And
                 a.ҽ��id Is Not Null
           Union
           Select Decode(a.ҽ��id, Null, b.ҽ��id, a.ҽ��id) As ҽ��id, b.������Դ
           From ������Ŀ�ֲ� A, ����걾��¼ B
           Where 1 = v_΢���� And b.Id = a.�걾id And b.Id = Id_In
           Union
           Select b.Id As ҽ��id, b.������Դ
           From ����걾��¼ A, ����ҽ����¼ B
           Where a.Id = Id_In And a.ҽ��id In (b.Id, b.���id));

  Cursor c_Stuff
  (
    v_No     Varchar2,
    v_��ҳid Number
  ) Is
    Select NO As ���ݺ�, ����, �ⷿid
    From δ��ҩƷ��¼
    Where NO = v_No And ���� In (24, 25, 26) And �ⷿid Is Not Null And Not Exists
     (Select 1 From Dual Where zl_GetSysParameter(Decode(v_��ҳid, Null, 92, 63)) = '1') And Exists
     (Select a.���
           From סԺ���ü�¼ A, �������� B
           Where a.��¼���� = 2 And a.No = v_No And a.�շ�ϸĿid = b.����id And b.�������� = 1
           Union All
           Select a.���
           From ������ü�¼ A, �������� B
           Where a.��¼���� = 2 And a.No = v_No And a.�շ�ϸĿid = b.����id And b.�������� = 1)
    Order By �ⷿid;

  r_Advice   v_Advice%RowType;
  r_Advice_1 v_Advice_1%RowType;
  r_Patient  v_Patient%RowType;

  Err_Custom Exception;
  v_Error Varchar2(1000);
  v_Flag  Number(18);

  v_Temp      Varchar2(255);
  v_Seq       Number;
  v_Union     Number;
  v_Patientid Number;
  v_Itemid    Number;
  v_Count     Number;
  v_ִ��      Number;
  v_No        ����ҽ������.No%Type;
  v_����      ����ҽ������.��¼����%Type;
  v_���      Varchar2(1000);
  v_��ҳid    Number(18);
  v_�����־  סԺ���ü�¼.�����־%Type;
  n_Count     Number;
  v_����      ����ҽ����¼.����%Type;
  v_�Ա�      ����ҽ����¼.�Ա�%Type;
  v_����      ����ҽ����¼.����%Type;
  v_������Դ  ����ҽ����¼.������Դ%Type;
  v_Ӥ��      ����ҽ����¼.Ӥ��%Type;
  v_Ӥ������  ����ҽ����¼.����%Type;
  v_Ӥ���Ա�  ����ҽ����¼.�Ա�%Type;

  n_Par Number;
Begin

  If Nvl(���Ǳ걾id_In, 0) > 0 Then
    Begin
      Select ���� Into v_Temp From ����걾��¼ Where ID = ���Ǳ걾id_In And ���� Is Null;
    Exception
      When Others Then
        v_Error := 'ָ�����ǵı걾�ѱ����ջ���ɾ����������ָ����';
        Raise Err_Custom;
    End;
  End If;

  If Nvl(ҽ��id_In, 0) > 0 Then
    Select ����, �Ա�, ����, ������Դ, Ӥ��
    Into v_����, v_�Ա�, v_����, v_������Դ, v_Ӥ��
    From ����ҽ����¼
    Where ID = ҽ��id_In;
  
    If v_������Դ <> 3 Then
      If Nvl(v_Ӥ��, 0) = 0 Then
        If v_���� <> ����_In Or v_�Ա� <> �Ա�_In Then
          v_Error := '�����������Ա������ҽ���������ܱ��棬������޸Ĳ�����Ϣ���ٽ��б��棡';
          Raise Err_Custom;
        End If;
      Else
        Select b.Ӥ������, b.Ӥ���Ա�
        Into v_Ӥ������, v_Ӥ���Ա�
        From ����ҽ����¼ A, ������������¼ B
        Where a.����id = b.����id And a.��ҳid = b.��ҳid And a.Ӥ�� = b.��� And
              a.���id In (Select * From Table(Cast(f_Num2list(���ҽ��_In) As Zltools.t_Numlist))) And Rownum = 1;
      
        If v_Ӥ������ <> ����_In Or v_Ӥ���Ա� <> �Ա�_In Then
          v_Error := '�����������Ա������ҽ���������ܱ��棬������޸Ĳ�����Ϣ���ٽ��б��棡';
          Raise Err_Custom;
        End If;
      End If;
    End If;
  
    Select Count(ID) Into v_Flag From ����걾��¼ Where ҽ��id = ҽ��id_In And ID <> Id_In;
    If v_Flag > 0 Then
      Select Count(Distinct b.������Ŀid)
      Into v_Flag
      From ����ҽ����¼ A, ���鱨����Ŀ B
      Where a.������Ŀid = b.������Ŀid And a.���id In (Select * From Table(Cast(f_Num2list(���ҽ��_In) As Zltools.t_Numlist)));
    
      Select Count(a.��Ŀid)
      Into n_Count
      From ������Ŀ�ֲ� A
      Where a.ҽ��id In (Select * From Table(Cast(f_Num2list(���ҽ��_In) As Zltools.t_Numlist))) And a.�걾id <> Id_In;
      If (v_Flag - n_Count) <= 0 Then
        v_Error := '��ǰҽ���ѱ����գ������ظ����գ�';
        Raise Err_Custom;
      End If;
    End If;
  End If;

  --�ж�ҽ���Ƿ��շ�
  n_Par := Zl_To_Number(Nvl(zl_GetSysParameter(163), '0'));
  If n_Par = 1 Then
    For r_Advice_1 In v_Advice_1 Loop
      For r_Verify In c_Verify(r_Advice_1.���id) Loop
        If r_Verify.��¼״̬ = 0 Then
          If r_Verify.�����־ = 1 Then
            v_Error := '�걾δ�շѣ�������ִ�У�����ϵ����Ա��';
            Raise Err_Custom;
          Elsif r_Verify.�����־ = 2 Then
            v_Error := '�걾δ���ˣ�������ִ�У�����ϵ����Ա��';
            Raise Err_Custom;
          End If;
        End If;
      End Loop;
    End Loop;
  End If;

  If ҽ��id_In = 0 Then
    Open v_Patient;
    Fetch v_Patient
      Into r_Patient;
  
    If v_Patient%Found Then
      Zl_������Ϣ_�������(r_Patient.����id);
    End If;
  
    Update ����걾��¼
    Set ����ʱ�� = Decode(����ʱ��_In, Null, ����ʱ��, ����ʱ��_In), ������ = Decode(������_In, Null, ������, ������_In), �걾���� = Nvl(�걾����_In, �걾����),
        ����ʱ�� = ����ʱ��_In, ���� = Decode(����_In, Null, ����, ����_In), �Ա� = Decode(�Ա�_In, Null, �Ա�, �Ա�_In),
        ���� = Decode(����_In, Null, ����, ����_In), �������� = Decode(����_In, Null, Null, Zl_Val(����_In)),
        ���䵥λ = Decode(����_In, Null, ���䵥λ,
                       Decode(����_In, Null, Null, '����', '����', 'Ӥ��', 'Ӥ��',
                               Decode(Sign(Instr(����_In, '��')), 1, '��',
                                       Decode(Sign(Instr(����_In, '��')), 1, '��',
                                               Decode(Sign(Instr(����_In, '��')), 1, '��',
                                                       Decode(Sign(Instr(����_In, 'Сʱ')), 1, 'Сʱ', Null)))))),
        �������id = Decode(�������id_In, Null, �������id, �������id_In), ������ = Decode(������_In, Null, ������, ������_In),
        �걾��̬ = Decode(�걾��̬_In, Null, �걾��̬, �걾��̬_In), ��ʶ�� = Decode(��ʶ��_In, Null, ��ʶ��, ��ʶ��_In),
        ���� = Decode(����_In, Null, ����, ����_In), ���˿��� = Decode(���˿���_In, Null, ���˿���, ���˿���_In),
        ������Ŀ = Decode(������Ŀ_In, Null, ������Ŀ, ������Ŀ_In), ����id = Decode(����id_In, Null, ����id, ����id_In),
        ҽ��id = Decode(ҽ��id_In, Null, ҽ��id, 0, ҽ��id, ҽ��id_In)
    Where ID = Id_In;
    If Sql%NotFound Then
      Insert Into ����걾��¼
        (ID, ҽ��id, �걾���, ����ʱ��, ������, �걾����, ������, ����ʱ��, ����״̬, ��������, ����id, ��������, ����ʱ��, �걾��̬, ������, ִ�п���id, ������, ����ʱ��, ΢����걾,
         �걾���, ���鱸ע, �������id, ������, ����, �Ա�, ����, ��������, ���䵥λ, ����id, ������Դ, Ӥ��, NO, �ϲ�id, ��ʶ��, ����, ���˿���, ����, �����, סԺ��, ��������,
         �Һŵ�, ��ҳid, ������Ŀ, ��������, ������, ����ʱ��)
      Values
        (Id_In, Decode(ҽ��id_In, 0, Null, ҽ��id_In), �걾���_In, ����ʱ��_In, ������_In, �걾����_In, ��Ա����_In, ����ʱ��_In, 1, ��������_In,
         Decode(����id_In, 0, Null, ����id_In), Null, Null, �걾��̬_In, 0, ִ�п���_In, ������_In, ����ʱ��_In, ΢����걾_In, �걾���_In, ���鱸ע_In,
         �������id_In, ������_In, ����_In, �Ա�_In, ����_In, Zl_Val(����_In),
         Decode(����_In, Null, Null, '����', '����', 'Ӥ��', 'Ӥ��',
                 Decode(Sign(Instr(����_In, '��')), 1, '��',
                         Decode(Sign(Instr(����_In, '��')), 1, '��',
                                 Decode(Sign(Instr(����_In, '��')), 1, '��', Decode(Sign(Instr(����_In, 'Сʱ')), 1, 'Сʱ', Null))))),
         ����id_In, Decode(r_Patient.סԺ��, Null, Decode(r_Patient.�����, Null, 3, 1), 2), 0, Null, Null, ��ʶ��_In, ����_In,
         ���˿���_In, �걾���_In, r_Patient.�����, r_Patient.סԺ��, r_Patient.��������, Null, Null, ������Ŀ_In, Null, Null, Null);
    End If;
    If Nvl(���Ǳ걾id_In, 0) > 0 Then
      Zl_����걾��¼_Union(Id_In, ���Ǳ걾id_In);
    End If;
    --��¼���պͲ������ 
    Insert Into ���������¼
      (ID, �걾id, ��������, ����Ա, ����ʱ��)
    Values
      (���������¼_Id.Nextval, Id_In, 2, ��Ա����_In, Sysdate);
    Close v_Patient;
  Else
    Open v_Advice;
    Fetch v_Advice
      Into r_Advice;
  
    If v_Advice%Found Then
      Zl_������Ϣ_�������(r_Advice.����id);
    End If;
  
    Update ����걾��¼
    Set ҽ��id = Decode(ҽ��id_In, Null, ҽ��id, 0, ҽ��id, ҽ��id_In), ����ʱ�� = Decode(����ʱ��_In, Null, ����ʱ��, ����ʱ��_In),
        ������ = Decode(������_In, Null, ������, ������_In), �걾��� = Decode(�걾���_In, Null, �걾���, �걾���_In),
        �걾���� = Decode(�걾����_In, Null, Decode(�걾����, Null, r_Advice.�걾��λ, �걾����), �걾����_In),
        ����ʱ�� = Decode(r_Advice.����ʱ��, Null, ����ʱ��, r_Advice.����ʱ��), ������ = Decode(������, Null, ��Ա����_In, ������),
        �������� = Decode(r_Advice.��������, Null, ��������, r_Advice.��������), �������� = Decode(��������_In, Null, ��������, ��������_In),
        ִ�п���id = Decode(ִ�п���_In, Null, ִ�п���id, ִ�п���_In), ������ = Decode(������_In, Null, ������, ������_In),
        ����ʱ�� = Decode(����ʱ��_In, Null, ����ʱ��, ����ʱ��_In), ���鱸ע = Decode(���鱸ע_In, Null, ���鱸ע, ���鱸ע_In),
        �������id = Decode(�������id_In, Null, �������id, �������id_In), ������ = Decode(������_In, Null, ������, ������_In),
        ���� = Decode(����_In, Null, ����, ����_In), �Ա� = Decode(�Ա�_In, Null, �Ա�, �Ա�_In), ���� = Decode(����_In, Null, ����, ����_In),
        �������� = Decode(����_In, Null, ��������, Zl_Val(����_In)),
        ���䵥λ = Decode(����_In, Null, ���䵥λ,
                       Decode(����_In, Null, Null, '����', '����', 'Ӥ��', 'Ӥ��',
                               Decode(Sign(Instr(����_In, '��')), 1, '��',
                                       Decode(Sign(Instr(����_In, '��')), 1, '��',
                                               Decode(Sign(Instr(����_In, '��')), 1, '��',
                                                       Decode(Sign(Instr(����_In, 'Сʱ')), 1, 'Сʱ', Null)))))),
        ����id = Decode(r_Advice.����id, Null, ����id, r_Advice.����id), ������Դ = Decode(r_Advice.������Դ, Null, ������Դ, r_Advice.������Դ),
        Ӥ�� = Decode(r_Advice.Ӥ��, Ӥ��, r_Advice.Ӥ��), NO = Decode(No_In, Null, NO, No_In), �ϲ�id = v_Union,
        �걾��̬ = Decode(�걾��̬_In, Null, �걾��̬, �걾��̬_In), ��ʶ�� = Decode(��ʶ��_In, Null, ��ʶ��, ��ʶ��_In),
        ���� = Decode(����_In, Null, ����, ����_In), ���˿��� = Decode(���˿���_In, Null, ���˿���, ���˿���_In), �걾��� = �걾���_In,
        ����� = r_Advice.�����, סԺ�� = r_Advice.סԺ��, �������� = r_Advice.��������, �Һŵ� = r_Advice.�Һŵ�, ��ҳid = r_Advice.��ҳid,
        ������Ŀ = Decode(������Ŀ_In, Null, ������Ŀ, ������Ŀ_In), �������� = r_Advice.��������, ������ = r_Advice.������, ����ʱ�� = r_Advice.����ʱ��
    Where ID = Id_In;
  
    If Sql%NotFound Then
      Insert Into ����걾��¼
        (ID, ҽ��id, �걾���, ����ʱ��, ������, �걾����, ������, ����ʱ��, ����״̬, ��������, ����id, ��������, ����ʱ��, �걾��̬, ������, ִ�п���id, ������, ����ʱ��, ΢����걾,
         �걾���, ���鱸ע, �������id, ������, ����, �Ա�, ����, ��������, ���䵥λ, ����id, ������Դ, Ӥ��, NO, �ϲ�id, ��ʶ��, ����, ���˿���, ����, �����, סԺ��, ��������,
         �Һŵ�, ��ҳid, ������Ŀ, ��������, ������, ����ʱ��)
      Values
        (Id_In, Decode(ҽ��id_In, 0, Null, ҽ��id_In), �걾���_In, ����ʱ��_In, ������_In, Nvl(�걾����_In, r_Advice.�걾��λ), ��Ա����_In,
         ����ʱ��_In, 1, ��������_In, Decode(����id_In, 0, Null, ����id_In), r_Advice.��������, r_Advice.����ʱ��, �걾��̬_In, 0, ִ�п���_In,
         ������_In, ����ʱ��_In, ΢����걾_In, �걾���_In, ���鱸ע_In, �������id_In, ������_In, ����_In, �Ա�_In, ����_In, Zl_Val(����_In),
         Decode(����_In, Null, Null, '����', '����', 'Ӥ��', 'Ӥ��',
                 Decode(Sign(Instr(����_In, '��')), 1, '��',
                         Decode(Sign(Instr(����_In, '��')), 1, '��',
                                 Decode(Sign(Instr(����_In, '��')), 1, '��', Decode(Sign(Instr(����_In, 'Сʱ')), 1, 'Сʱ', Null))))),
         r_Advice.����id, r_Advice.������Դ, r_Advice.Ӥ��, No_In, v_Union, ��ʶ��_In, ����_In, ���˿���_In, r_Advice.����, r_Advice.�����,
         r_Advice.סԺ��, r_Advice.��������, r_Advice.�Һŵ�, r_Advice.��ҳid, ������Ŀ_In, r_Advice.��������, r_Advice.������, r_Advice.����ʱ��);
    End If;
    If Nvl(���Ǳ걾id_In, 0) > 0 Then
      Zl_����걾��¼_Union(Id_In, ���Ǳ걾id_In);
    End If;
    Insert Into ���������¼
      (ID, �걾id, ��������, ����Ա, ����ʱ��)
    Values
      (���������¼_Id.Nextval, Id_In, 2, ��Ա����_In, Sysdate);
  
    --��������Ŀ��ʱ��д�ϲ�ID 
    Begin
      Select a.Id
      Into v_Union
      From ����걾��¼ A, ����걾��¼ B, ����ҽ����¼ C, ����ϲ����� D, ����ҽ����¼ E
      Where a.����id = b.����id And b.Id = Id_In And a.����״̬ = 1 And Nvl(a.����id, 0) <> 0 And a.ҽ��id = c.���id And
            d.����Ŀid = c.������Ŀid And d.�ϲ���Ŀid = e.������Ŀid And e.Id = r_Advice.Id And Rownum = 1
      Order By a.����ʱ�� Desc;
    Exception
      When Others Then
        v_Union := Null;
    End;
    If Nvl(v_Union, 0) <> 0 Then
      Update ����걾��¼ Set �ϲ�id = v_Union Where (ID = Id_In Or ҽ��id = r_Advice.Id);
    End If;
    --������������Ŀʱ��д�ϲ���Ŀ 
    Begin
      Select a.Id, a.����id, c.����Ŀid
      Into v_Union, v_Patientid, v_Itemid
      From ����걾��¼ A, ����ҽ����¼ B, ����ϲ����� C
      Where a.ҽ��id = b.���id And b.������Ŀid = c.����Ŀid And a.Id = Id_In And Rownum = 1;
    Exception
      When Others Then
        v_Union := Null;
    End;
    If Nvl(v_Union, 0) <> 0 Then
      Update ����걾��¼
      Set �ϲ�id = v_Union
      Where ID In (Select a.Id
                   From ����걾��¼ A, ����ҽ����¼ B, ����ϲ����� C
                   Where a.ҽ��id = b.���id And b.������Ŀid = c.�ϲ���Ŀid And c.����Ŀid = v_Itemid And a.����id = v_Patientid And
                         a.����״̬ = 1);
    End If;
  
    v_Seq := 1;
    Close v_Advice;
    v_Flag := 0;
    Begin
      Select Nvl(Max(1), 0) Into v_Flag From ����������Ŀ Where �걾id = Id_In;
    Exception
      When Others Then
        v_Flag := 0;
    End;
    If v_Flag = 0 Then
      For r_Advice In v_Advice Loop
        Update ����������Ŀ
        Set �걾id = Id_In, ������Ŀid = r_Advice.������Ŀid
        Where �걾id = Id_In And ������Ŀid = r_Advice.������Ŀid;
        If Sql%RowCount = 0 Then
          Insert Into ����������Ŀ (�걾id, ������Ŀid, ���) Values (Id_In, r_Advice.������Ŀid, v_Seq);
        End If;
        v_Seq := v_Seq + 1;
      End Loop;
    End If;
  
  End If;

  --���ݲ������ж��Ƿ��� 
  For r_Advice_1 In v_Advice_1 Loop
    --�������û���Զ�����,���Զ�����,���򲻴��� 
    For r_Stuff In c_Stuff(r_Advice_1.���ݺ�, v_��ҳid) Loop
    
      Zl_�����շ���¼_��������(r_Stuff.�ⷿid, r_Stuff.����, r_Stuff.���ݺ�, ��Ա����_In, ��Ա����_In, ��Ա����_In, 1, Sysdate);
    End Loop;
  End Loop;

  Update /*+ Rule */ ����ҽ������
  Set ִ��״̬ = 3
  Where ִ��״̬ = 0 And
        ҽ��id In (Select ID
                 From ����ҽ����¼
                 Where ID In (Select * From Table(Cast(f_Num2list(���ҽ��_In) As Zltools.t_Numlist)))
                 Union All
                 Select ID
                 From ����ҽ����¼
                 Where ���id In (Select * From Table(Cast(f_Num2list(���ҽ��_In) As Zltools.t_Numlist))));
  --ִ�к��Զ���˶�Ӧ�ļ��ʻ��۵�(������ҩƷ)
  Select Zl_To_Number(Nvl(zl_GetSysParameter(81), '0')) Into v_ִ�� From Dual;
  --2.��鵱ǰ�걾��ص��������ر걾�Ƿ������� 
  For r_Samplequest In c_Samplequest(΢����걾_In) Loop
  
    v_Count := 0;
  
    --r_SampleQuest.ҽ��id�����Ѿ����,����������� 
    If v_Count = 0 Then
    
      If r_Samplequest.������Դ = 2 Then
        Update סԺ���ü�¼
        Set ִ��״̬ = 1, ִ��ʱ�� = Sysdate, ִ���� = ��Ա����_In
        Where �շ���� Not In ('5', '6', '7') And
              (ҽ�����, ��¼����, NO) In
              (Select ҽ��id, ��¼����, NO
               From ����ҽ������
               Where ҽ��id = r_Samplequest.ҽ��id
               Union All
               Select ҽ��id, ��¼����, NO
               From ����ҽ������
               Where ҽ��id In (Select ID From ����ҽ����¼ Where r_Samplequest.ҽ��id In (ID, ���id)));
      Else
        Update ������ü�¼
        Set ִ��״̬ = 1, ִ��ʱ�� = Sysdate, ִ���� = ��Ա����_In
        Where �շ���� Not In ('5', '6', '7') And
              (ҽ�����, ��¼����, NO) In
              (Select ҽ��id, ��¼����, NO
               From ����ҽ������
               Where ҽ��id = r_Samplequest.ҽ��id
               Union All
               Select ҽ��id, ��¼����, NO
               From ����ҽ������
               Where ҽ��id In (Select ID From ����ҽ����¼ Where r_Samplequest.ҽ��id In (ID, ���id)));
      End If;
      --3.�Զ���˼��� 
      If Nvl(v_ִ��, 0) = 1 Then
        For r_Verify In c_Verify(r_Samplequest.ҽ��id) Loop
          If r_Verify.No || ',' || r_Verify.��¼���� <> v_No || ',' || v_���� Then
            If v_��� Is Not Null Then
              If r_Verify.�����־ = 1 Then
                Zl_������ʼ�¼_Verify(v_No, ��Ա���_In, ��Ա����_In, Substr(v_���, 2));
              Elsif r_Verify.�����־ = 2 Then
                Zl_סԺ���ʼ�¼_Verify(v_No, ��Ա���_In, ��Ա����_In, Substr(v_���, 2));
              End If;
            End If;
            v_��� := Null;
          End If;
          v_�����־ := r_Verify.�����־;
          v_No       := r_Verify.No;
          v_����     := r_Verify.��¼����;
          v_���     := v_��� || ',' || r_Verify.���;
        End Loop;
        If v_��� Is Not Null Then
          If v_�����־ = 1 Then
            Zl_������ʼ�¼_Verify(v_No, ��Ա���_In, ��Ա����_In, Substr(v_���, 2));
          Elsif v_�����־ = 2 Then
            Zl_סԺ���ʼ�¼_Verify(v_No, ��Ա���_In, ��Ա����_In, Substr(v_���, 2));
          End If;
        End If;
      End If;
    End If;
  End Loop;

  If Nvl(��������_In, 0) = 1 Then
    Zl_����ҽ����¼_���δ�ӡ(ҽ��id_In, 1);
  End If;

Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_����걾��¼_�걾����;
/

--107559:Ƚ����,2017-04-17,������ֹͣ�ﰲ�Ź���
Create Or Replace Procedure Zl_�ٴ�����ͣ��_Apply
(
  ��������_In Number,
  Id_In       �ٴ�����ͣ���¼.Id%Type,
  ��ʼʱ��_In �ٴ�����ͣ���¼.��ʼʱ��%Type := Null,
  ��ֹʱ��_In �ٴ�����ͣ���¼.��ֹʱ��%Type := Null,
  ͣ��ԭ��_In �ٴ�����ͣ���¼.ͣ��ԭ��%Type := Null,
  ������_In   �ٴ�����ͣ���¼.������%Type := Null,
  ����ʱ��_In �ٴ�����ͣ���¼.����ʱ��%Type := Null,
  �Ǽ���_In   �ٴ�����ͣ���¼.�Ǽ���%Type := Null
) As
  --���ܣ��˷������Լ�ȡ������
  --������
  --        ��������_In��0-���룬else-ȡ������
  --˵����
  n_Count Number;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  If ��������_In = 0 Then
    --����
    If ��ʼʱ��_In <= Sysdate Then
      v_Error := 'ͣ��ʱ��Ŀ�ʼʱ�������ڵ�ǰʱ�䣡';
      Raise Err_Custom;
    End If;
  
    If ��ʼʱ��_In >= ��ֹʱ��_In Then
      v_Error := 'ͣ��ʱ��Ľ���ʱ�������ڿ�ʼʱ�䣡';
      Raise Err_Custom;
    End If;
  
    Select Count(1)
    Into n_Count
    From �ٴ�����ͣ���¼
    Where ��¼id Is Null And Not (��ʼʱ�� > ��ֹʱ��_In Or Nvl(ʧЧʱ��, ��ֹʱ��) < ��ʼʱ��_In) And ������ = ������_In And Rownum < 2;
    If n_Count <> 0 Then
      v_Error := 'ҽ�� ' || ������_In || ' �ڵ�ǰͣ��ʱ�䷶Χ���Ѵ���ͣ�ﰲ�ţ������ظ����룡';
      Raise Err_Custom;
    End If;
  
    Insert Into �ٴ�����ͣ���¼
      (ID, ��ʼʱ��, ��ֹʱ��, ͣ��ԭ��, ������, ����ʱ��, �Ǽ���)
    Values
      (�ٴ�����ͣ���¼_Id.Nextval, ��ʼʱ��_In, ��ֹʱ��_In, ͣ��ԭ��_In, ������_In, ����ʱ��_In, �Ǽ���_In);
  
    Return;
  End If;

  --ȡ������
  Select Count(1) Into n_Count From �ٴ�����ͣ���¼ Where ID = Id_In And ������ Is Not Null;
  If n_Count <> 0 Then
    v_Error := '�������ѱ�����������ȡ�����롣';
    Raise Err_Custom;
  End If;

  Delete �ٴ�����ͣ���¼ Where ID = Id_In;
  If Sql%NotFound Then
    v_Error := '����������ѱ�����ȡ�����룬��ˢ�º�鿴...';
    Raise Err_Custom;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ�����ͣ��_Apply;
/

--107559:Ƚ����,2017-04-17,������ֹͣ�ﰲ�Ź���
Create Or Replace Procedure Zl_�ٴ�����ͣ��_Stop
(
  Id_In       �ٴ�����ͣ���¼.Id%Type,
  ��ֹ��_In   �ٴ�����ͣ���¼.ȡ����%Type,
  ��ֹʱ��_In �ٴ�����ͣ���¼.ʧЧʱ��%Type := Null
) As
  --���ܣ���ֹͣ�ﰲ��
  --������
  --       ��ֹʱ��_In��Null-������ֹ������-�������ֹʱ��
  v_Error Varchar2(255);
  Err_Custom Exception;

  n_Count Number;
Begin
  If ��ֹʱ��_In Is Not Null Then
    If ��ֹʱ��_In < Sysdate Then
      v_Error := '��ֹʱ�������ڵ�ǰʱ�䣡';
      Raise Err_Custom;
    End If;
  End If;

  Select Count(1) Into n_Count From �ٴ�����ͣ���¼ Where ID = Id_In And ��ֹʱ�� < Sysdate;
  If n_Count <> 0 Then
    v_Error := '��ͣ�ﰲ����ʧЧ��������ֹ��';
    Raise Err_Custom;
  End If;

  Select Count(1) Into n_Count From �ٴ�����ͣ���¼ Where ID = Id_In And ʧЧʱ�� Is Not Null;
  If n_Count <> 0 Then
    v_Error := '��ͣ�ﰲ���ѱ���ֹ����������ֹ��';
    Raise Err_Custom;
  End If;

  Update �ٴ�����ͣ���¼
  Set ʧЧʱ�� = Nvl(��ֹʱ��_In, Sysdate), ȡ���� = ��ֹ��_In, ȡ��ʱ�� = Sysdate
  Where ID = Id_In And ������ Is Not Null;
  If Sql%NotFound Then
    v_Error := '��ͣ�ﰲ�Ż�δ������������ֹ��';
    Raise Err_Custom;
  End If;

  For c_��¼ In (Select a.Id, c.����, a.ͣ����ֹʱ��, a.�Ƿ���ſ���, a.�Ƿ��ʱ��
               From �ٴ������¼ A, �ٴ�����ͣ���¼ B, �ٴ������Դ C
               Where ((a.����ҽ������ Is Null And a.ҽ��id Is Not Null And a.ҽ������ = b.������) Or
                     (a.����ҽ������ Is Not Null And a.����ҽ��id Is Not Null And a.����ҽ������ = b.������)) And a.��Դid = c.Id And
                     b.Id = Id_In And (a.��ʼʱ�� Between b.��ʼʱ�� And b.��ֹʱ�� Or a.��ֹʱ�� Between b.��ʼʱ�� And b.��ֹʱ��) And
                     Nvl(a.�Ƿ񷢲�, 0) = 1 And a.ͣ����ֹʱ�� > Nvl(��ֹʱ��_In, Sysdate)) Loop
  
    Update �ٴ������¼
    Set ͣ�￪ʼʱ�� = Case
                   When ͣ�￪ʼʱ�� >= Nvl(��ֹʱ��_In, Sysdate) Then
                    Null
                   Else
                    ͣ�￪ʼʱ��
                 End,
        ͣ����ֹʱ�� = Case
                   When ͣ�￪ʼʱ�� >= Nvl(��ֹʱ��_In, Sysdate) Then
                    Null
                   Else
                    Nvl(��ֹʱ��_In, Sysdate)
                 End
    Where ID = c_��¼.Id;
  
    --����"�ٴ�������ſ���.�Ƿ�ͣ��"Ϊ0
    Update �ٴ�������ſ���
    Set �Ƿ�ͣ�� = 0
    Where ��¼id = c_��¼.Id And Nvl(�Ƿ�ͣ��, 0) = 1 And ��ʼʱ�� Between Nvl(��ֹʱ��_In, Sysdate) And c_��¼.ͣ����ֹʱ�� And
          Nvl(c_��¼.�Ƿ���ſ���, 0) = 1 And Nvl(c_��¼.�Ƿ��ʱ��, 0) = 1;
  
    --��Ϣ����
    -- ͣ������(1-ͣ��,2-ȡ��ͣ��),�����¼ID,ͣ�����
    Begin
      Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
        Using 17, 2 || ',' || c_��¼.Id || ',' || c_��¼.����;
    Exception
      When Others Then
        Null;
    End;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ�����ͣ��_Stop;
/

--108264:������,2017-04-17,���˸��ӷ��ô�������
Create Or Replace Procedure Zl_������ʼ�¼_Delete
(
  No_In         ������ü�¼.No%Type,
  ���_In       Varchar2,
  ����Ա���_In ������ü�¼.����Ա���%Type,
  ����Ա����_In ������ü�¼.����Ա����%Type
) As
  --���ܣ�����һ��������ʵ�����ָ�������
  --��ţ���ʽ��"1,3,5,7,8",Ϊ�ձ�ʾ�������пɳ�����
  --�ù����������ָ��������

  --���α�ΪҪ�˷ѵ��ݵ�����ԭʼ��¼
  Cursor c_Bill(n_��־ Number) Is
    Select a.Id, a.�۸񸸺�, a.���, a.ִ��״̬, a.�շ����, a.ҽ�����, a.����id, a.������Ŀid, a.��������id, a.ִ�в���id, a.���˿���id, a.ʵ�ս��,
           Decode(a.��¼״̬, 0, 1, 0) As ����, j.�������, m.��������
    From ������ü�¼ A, ����ҽ����¼ J, �������� M
    Where a.ҽ����� = j.Id(+) And a.�շ�ϸĿid + 0 = m.����id(+) And a.No = No_In And a.��¼���� = 2 And a.��¼״̬ In (0, 1, 3) And
          a.�����־ = n_��־
    Order By a.�շ�ϸĿid, a.���;

  --���α����ڴ���ҩƷ����������
  --��Ҫ�ܷ��õ�ִ��״̬,��Ϊ���ڴ˲�����
  Cursor c_Stock(n_��־ Number) Is
    Select ID, �ⷿid, ҩƷid, ����, ����, ����, ʵ������, ����, ��ҩ��ʽ, ���Ч��, Ч��, ��Ʒ����, �ڲ�����, ����id
    From ҩƷ�շ���¼
    Where NO = No_In And ���� In (9, 25) And Mod(��¼״̬, 3) = 1 And ����� Is Null And
          ����id In (Select ID
                   From ������ü�¼
                   Where NO = No_In And ��¼���� = 2 And ��¼״̬ In (0, 1, 3) And �շ���� In ('4', '5', '6', '7') And �����־ = n_��־ And
                         (Instr(',' || ���_In || ',', ',' || ��� || ',') > 0 Or ���_In Is Null))
    Order By ҩƷid;

  --���α����ڴ�����ü�¼���
  Cursor c_Serial Is
    Select ���, �۸񸸺� From ������ü�¼ Where NO = No_In And ��¼���� = 2 And ��¼״̬ In (0, 1, 3) Order By ���;
  l_ҩƷ�շ� t_Numlist := t_Numlist();
  l_����     t_Numlist := t_Numlist();
  l_����id   t_Numlist := t_Numlist();
  n_�������� Number;

  v_ҽ��ids Varchar2(4000);

  n_ҽ��id   ����ҽ����¼.Id%Type;
  n_����     ������ü�¼.�۸񸸺�%Type;
  n_�����־ ������ü�¼.�����־%Type;

  --�����˷Ѽ������
  n_ʣ������ Number;
  n_ʣ��Ӧ�� Number;
  n_ʣ��ʵ�� Number;
  n_ʣ��ͳ�� Number;

  n_׼������ Number;
  n_�˷Ѵ��� Number;

  n_Ӧ�ս�� Number;
  n_ʵ�ս�� Number;
  n_ͳ���� Number;

  n_Dec Number;

  n_Count   Number;
  d_Curdate Date;
  Err_Item Exception;
  v_Err_Msg Varchar2(255);
Begin
  --�Ƿ��Ѿ�ȫ����ȫִ��(ֻ�����ŵ��ݵļ��)
  Select Nvl(Count(*), 0), Max(Nvl(�����־, 1))
  Into n_Count, n_�����־
  From ������ü�¼
  Where NO = No_In And ��¼���� = 2 And ��¼״̬ In (0, 1, 3) And Nvl(ִ��״̬, 0) <> 1;
  If n_Count = 0 Then
    v_Err_Msg := '�õ����е���Ŀ�Ѿ�ȫ����ȫִ�У�';
    Raise Err_Item;
  End If;

  If Nvl(n_�����־, 0) = 0 Then
    n_�����־ := 1;
  End If;

  --δ��ȫִ�е���Ŀ�Ƿ���ʣ������(ֻ�����ŵ��ݵļ��)
  Select Nvl(Count(*), 0)
  Into n_Count
  From (Select ���, Sum(����) As ʣ������
         From (Select ��¼״̬, Nvl(�۸񸸺�, ���) As ���, Avg(Nvl(����, 1) * ����) As ����
                From ������ü�¼
                Where NO = No_In And ��¼���� = 2 And �����־ = n_�����־ And
                      Nvl(�۸񸸺�, ���) In
                      (Select Nvl(�۸񸸺�, ���)
                       From ������ü�¼
                       Where NO = No_In And ��¼���� = 2 And �����־ = n_�����־ And ��¼״̬ In (0, 1, 3) And Nvl(ִ��״̬, 0) <> 1)
                Group By ��¼״̬, Nvl(�۸񸸺�, ���))
         Group By ���
         Having Sum(����) <> 0);
  If n_Count = 0 Then
    v_Err_Msg := '�õ�����δ��ȫִ�в�����Ŀʣ������Ϊ��,û�п������ʵķ��ã�';
    Raise Err_Item;
  End If;

  ---------------------------------------------------------------------------------
  --���ñ���
  Select Sysdate Into d_Curdate From Dual;

  --���С��λ��
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_Dec From Dual;

  --ѭ������ÿ�з���(������Ŀ��)
  For r_Bill In c_Bill(n_�����־) Loop
    If Instr(',' || ���_In || ',', ',' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || ',') > 0 Or ���_In Is Null Then
    
      If r_Bill.���� = 0 Then
        If Nvl(r_Bill.ִ��״̬, 0) <> 1 Then
          --��ʣ������,ʣ��Ӧ��,ʣ��ʵ��
          Select Sum(Nvl(����, 1) * ����), Sum(Ӧ�ս��), Sum(ʵ�ս��), Sum(ͳ����)
          Into n_ʣ������, n_ʣ��Ӧ��, n_ʣ��ʵ��, n_ʣ��ͳ��
          From ������ü�¼
          Where NO = No_In And ��¼���� = 2 And ��� = r_Bill.���;
        
          If n_ʣ������ = 0 Then
            If ���_In Is Not Null Then
              v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����Ѿ�ȫ�����ʣ�';
              Raise Err_Item;
            End If;
            --�����δ�޶��к�,ԭʼ�����еĸñ��Ѿ�ȫ������(ִ��״̬=0��һ�ֿ���)
          Else
            --׼������(��ҩƷ��ĿΪʣ������,ԭʼ����)
            If Instr(',4,5,6,7,', r_Bill.�շ����) = 0 Or (r_Bill.�շ���� = '4' And Nvl(r_Bill.��������, 0) = 0) Then
            
              --@@@
              --��ҩƷ����(�Ծ���ҽ��ִ��Ϊ׼���м��)
              --: 1.����ҽ�����͵�,����ҽ��ִ��Ϊ׼(�����ܰ���:���;����;����;������Ѫ)
              --: 2.���ڲ���ҽ�ԼƼ��е��շѷ�ʽΪ:0-������ȡ ��,��֧�ֲ�����;�����������,��ֻ��ȫ��
              --: 3.������ҽ����,����ʣ������Ϊ׼
              n_Count := 0;
              If Instr(',C,D,F,G,K,', ',' || r_Bill.������� || ',') = 0 And r_Bill.������� Is Not Null Then
              
                Select Nvl(Sum(����), 0), Count(*)
                Into n_׼������, n_Count
                From (Select j.ҽ����� As ҽ��id, j.�շ�ϸĿid, Nvl(j.����, 1) * Nvl(j.����, 1) As ����
                       From ������ü�¼ J, ����ҽ����¼ M
                       Where j.ҽ����� = m.Id And j.No = No_In And j.��¼���� = 2 And j.��� = r_Bill.��� And j.��¼״̬ In (1, 3) And
                             Exists
                        (Select 1
                              From ����ҽ������ A
                              Where a.ҽ��id = j.ҽ����� And Nvl(a.ִ��״̬, 0) <> 1 And a.No || '' = No_In) And Exists
                        (Select 1
                              From ����ҽ���Ƽ� A
                              Where a.ҽ��id = j.ҽ����� And a.�շ�ϸĿid = j.�շ�ϸĿid And Nvl(a.�շѷ�ʽ, 0) = 0) And j.�۸񸸺� Is Null And
                             Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And
                             (j.��¼״̬ In (1, 3) And Not Exists
                              (Select 1
                               From ҩƷ�շ���¼
                               Where ����id = j.Id And Instr(',8,9,10,21,24,25,26,', ',' || ���� || ',') > 0) Or
                              j.��¼״̬ = 2 And Not Exists
                              (Select 1 From ҩƷ�շ���¼ Where NO = No_In And ���� In (8, 24) And ҩƷid = j.�շ�ϸĿid))
                       Union All
                       Select a.ҽ��id, a.�շ�ϸĿid, -1 * Nvl(a.����, 1) * Nvl(c.��������, 1) As ����
                       From ����ҽ���Ƽ� A, ����ҽ������ B, ����ҽ��ִ�� C, ������ü�¼ J, ����ҽ����¼ M
                       Where a.ҽ��id = b.ҽ��id And b.ҽ��id = c.ҽ��id And Nvl(a.�շѷ�ʽ, 0) = 0 And b.���ͺ� = c.���ͺ� And
                             a.ҽ��id = m.Id And Nvl(c.ִ�н��, 1) = 1 And Nvl(b.ִ��״̬, 0) <> 1 And a.ҽ��id = j.ҽ����� And
                             a.�շ�ϸĿid = j.�շ�ϸĿid And j.No = No_In And j.��¼���� = 2 And j.��� = r_Bill.��� And
                             j.��¼״̬ In (1, 3) And j.�۸񸸺� Is Null And Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0 And
                             Not Exists
                        (Select 1
                              From ҩƷ�շ���¼
                              Where ����id = j.Id And Instr(',8,9,10,21,24,25,26,', ',' || ���� || ',') > 0) And Not Exists
                        (Select 1 From �������� Where ����id = j.�շ�ϸĿid And Nvl(��������, 0) = 1)
                       Union All
                       Select a.ҽ��id, a.�շ�ϸĿid, 0 As ����
                       From ����ҽ���Ƽ� A, ������ü�¼ J, ����ҽ����¼ M
                       Where a.ҽ��id = m.Id And a.ҽ��id = j.ҽ����� And a.�շ�ϸĿid = j.�շ�ϸĿid And Nvl(a.�շѷ�ʽ, 0) <> 0 And
                             j.No = No_In And j.��¼���� = 2 And Nvl(j.ִ��״̬, 0) = 2 And Not Exists
                        (Select 1 From �������� Where ����id = j.�շ�ϸĿid And Nvl(��������, 0) = 1) And
                             Instr(',C,D,F,G,K,', ',' || m.������� || ',') = 0);
              
              End If;
            
              If Nvl(n_Count, 0) = 0 Then
                n_׼������ := n_ʣ������;
              End If;
            
            Else
              Select Sum(Nvl(����, 1) * ʵ������)
              Into n_׼������
              From ҩƷ�շ���¼
              Where NO = No_In And ���� In (9, 25) And Mod(��¼״̬, 3) = 1 And ����� Is Null And ����id = r_Bill.Id;
            
              --���������õ���������
              If r_Bill.�շ���� = '4' And Nvl(n_׼������, 0) = 0 Then
                n_׼������ := n_ʣ������;
              End If;
            End If;
          
            --����������ü�¼
          
            --�ñ���Ŀ�ڼ�������
            Select Nvl(Max(Abs(ִ��״̬)), 0) + 1
            Into n_�˷Ѵ���
            From ������ü�¼
            Where NO = No_In And ��¼���� = 2 And ��¼״̬ = 2 And ��� = r_Bill.���;
          
            --���=ʣ����*(׼����/ʣ����)
            n_Ӧ�ս�� := Round(n_ʣ��Ӧ�� * (n_׼������ / n_ʣ������), n_Dec);
            n_ʵ�ս�� := Round(n_ʣ��ʵ�� * (n_׼������ / n_ʣ������), n_Dec);
            n_ͳ���� := Round(n_ʣ��ͳ�� * (n_׼������ / n_ʣ������), n_Dec);
          
            --�����˷Ѽ�¼
            Insert Into ������ü�¼
              (ID, NO, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, Ӥ����, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�, ���˿���id, �շ����,
               �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������, ִ�в���id, ������,
               ִ����, ִ��״̬, ִ��ʱ��, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ������Ŀ��, ���մ���id, ͳ����, ���ʵ�id, ժҪ, ���ձ���, �Ƿ���, ����)
              Select ���˷��ü�¼_Id.Nextval, NO, ��¼����, 2, ���, ��������, �۸񸸺�, ����id, ҽ�����, �����־, Ӥ����, ����, �Ա�, ����, ��ʶ��, ���ʽ, �ѱ�,
                     ���˿���id, �շ����, �շ�ϸĿid, ���㵥λ, Decode(Sign(n_׼������ - Nvl(����, 1) * ����), 0, ����, 1), ��ҩ����,
                     Decode(Sign(n_׼������ - Nvl(����, 1) * ����), 0, -1 * ����, -1 * n_׼������), �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���,
                     ��׼����, -1 * n_Ӧ�ս��, -1 * n_ʵ�ս��, ��������id, ������, ִ�в���id, ������, ִ����, -1 * n_�˷Ѵ���, ִ��ʱ��, ����Ա���_In,
                     ����Ա����_In, ����ʱ��, d_Curdate, ������Ŀ��, ���մ���id, -1 * n_ͳ����, ���ʵ�id, ժҪ, ���ձ���, �Ƿ���, ����
              From ������ü�¼
              Where ID = r_Bill.Id;
          
            --��¼����ҽ�����Ѷ�Ӧ��ҽ��ID(����������)
            If n_ҽ��id Is Null And r_Bill.ҽ����� Is Not Null Then
              n_ҽ��id := r_Bill.ҽ�����;
            End If;
          
            --�������
            Update �������
            Set ������� = Nvl(�������, 0) - n_ʵ�ս��
            Where ����id = r_Bill.����id And ���� = 1 And ���� = 1;
            If Sql%RowCount = 0 Then
              Insert Into �������
                (����id, ����, ����, �������, Ԥ�����)
              Values
                (r_Bill.����id, 1, 1, -1 * n_ʵ�ս��, 0);
            End If;
          
            --����δ�����
            Update ����δ�����
            Set ��� = Nvl(���, 0) - n_ʵ�ս��
            Where ����id = r_Bill.����id And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And
                  Nvl(���˿���id, 0) = Nvl(r_Bill.���˿���id, 0) And Nvl(��������id, 0) = Nvl(r_Bill.��������id, 0) And
                  Nvl(ִ�в���id, 0) = Nvl(r_Bill.ִ�в���id, 0) And ������Ŀid + 0 = r_Bill.������Ŀid And ��Դ;�� + 0 = n_�����־;
            If Sql%RowCount = 0 Then
              Insert Into ����δ�����
                (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
              Values
                (r_Bill.����id, Null, Null, r_Bill.���˿���id, r_Bill.��������id, r_Bill.ִ�в���id, r_Bill.������Ŀid, n_�����־,
                 -1 * n_ʵ�ս��);
            End If;
          
            --���ԭ���ü�¼
            --ִ��״̬:ȫ������(׼����=ʣ����)���Ϊ0,������Ϊ1
            Update ������ü�¼
            Set ��¼״̬ = 3, ִ��״̬ = Decode(Sign(n_׼������ - n_ʣ������), 0, 0, 1)
            Where ID = r_Bill.Id;
          End If;
        Else
          If ���_In Is Not Null Then
            v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����Ѿ���ȫִ��,�������ʣ�';
            Raise Err_Item;
          End If;
          --���:û�޶��к�,ԭʼ�����а����Ѿ���ȫִ�е�
        End If;
      End If;
    End If;
  End Loop;

  ---------------------------------------------------------------------------------
  --ҩƷ�������
  ------------------------------------------------------------------------------------------------------------------------
  --�ȴ���������
  For v_���� In (Select ID, �ⷿid, ҩƷid, ����, ����, ����, ʵ������, ����, ��ҩ��ʽ, ���Ч��, Ч��, ��Ʒ����, �ڲ�����, ����id
               From ҩƷ�շ���¼
               Where ���� = 21 And Mod(��¼״̬, 3) = 1 And ����� Is Null And
                     ����id In (Select ID
                              From ������ü�¼
                              Where NO = No_In And ��¼���� = 2 And ��¼״̬ In (0, 1, 3) And �շ���� = '4' And �����־ = n_�����־ And
                                    (Instr(',' || ���_In || ',', ',' || ��� || ',') > 0 Or ���_In Is Null))
               Order By ҩƷid) Loop
    --����ҩƷ���
    If v_����.�ⷿid Is Not Null Then
      Update ҩƷ���
      Set �������� = Nvl(��������, 0) + Decode(v_����.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(v_����.����, 1) * Nvl(v_����.ʵ������, 0)
      Where �ⷿid = v_����.�ⷿid And ҩƷid = v_����.ҩƷid And Nvl(����, 0) = Nvl(v_����.����, 0) And ���� = 1;
      If Sql%RowCount = 0 Then
        Insert Into ҩƷ���
          (�ⷿid, ҩƷid, ����, ����, Ч��, ��������, �ϴ�����, �ϴβ���, ���Ч��, ��Ʒ����, �ڲ�����)
        Values
          (v_����.�ⷿid, v_����.ҩƷid, 1, v_����.����, v_����.Ч��,
           Decode(v_����.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(v_����.����, 1) * Nvl(v_����.ʵ������, 0), v_����.����, v_����.����, v_����.���Ч��,
           v_����.��Ʒ����, v_����.�ڲ�����);
      End If;
    End If;
    l_����id.Extend;
    l_����id(l_����id.Count) := v_����.����id;
    l_ҩƷ�շ�.Extend;
    l_ҩƷ�շ�(l_ҩƷ�շ�.Count) := v_����.Id;
  End Loop;

  For r_Stock In c_Stock(n_�����־) Loop
  
    --����ҩƷ���
    If r_Stock.�ⷿid Is Not Null Then
    
      Select Decode(Count(Column_Value), Null, 0, 0, 0, 1)
      Into n_��������
      From Table(l_����id)
      Where Column_Value = r_Stock.����id;
      If Nvl(n_��������, 0) = 0 Then
        Update ҩƷ���
        Set �������� = Nvl(��������, 0) + Decode(r_Stock.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(r_Stock.����, 1) * Nvl(r_Stock.ʵ������, 0)
        Where �ⷿid = r_Stock.�ⷿid And ҩƷid = r_Stock.ҩƷid And Nvl(����, 0) = Nvl(r_Stock.����, 0) And ���� = 1;
        If Sql%RowCount = 0 Then
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, Ч��, ��������, �ϴ�����, �ϴβ���, ���Ч��, ��Ʒ����, �ڲ�����)
          Values
            (r_Stock.�ⷿid, r_Stock.ҩƷid, 1, r_Stock.����, r_Stock.Ч��,
             Decode(r_Stock.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(r_Stock.����, 1) * Nvl(r_Stock.ʵ������, 0), r_Stock.����, r_Stock.����,
             r_Stock.���Ч��, r_Stock.��Ʒ����, r_Stock.�ڲ�����);
        End If;
      End If;
    End If;
  
    l_ҩƷ�շ�.Extend;
    l_ҩƷ�շ�(l_ҩƷ�շ�.Count) := r_Stock.Id;
  End Loop;

  --ɾ��ҩƷ�շ���¼
  Forall I In 1 .. l_ҩƷ�շ�.Count
    Delete From ҩƷ�շ���¼ Where ID = l_ҩƷ�շ�(I);

  ------------------------------------------------------------------------------------------------------------------------
  --����ɾδ��ҩƷ��¼

  Delete From δ��ҩƷ��¼ A
  Where NO = No_In And ���� In (9, 25) And Not Exists
   (Select 1
         From ҩƷ�շ���¼
         Where ���� = a.���� And Nvl(�ⷿid, 0) = Nvl(a.�ⷿid, 0) And NO = No_In And Mod(��¼״̬, 3) = 1 And ����� Is Null);
  ------------------------------------------------------------------------------------------------------------------------

  ---------------------------------------------------------------------------------
  --����ǻ���,ֱ��ɾ�����ü�¼(ҩƷ�����)
  n_Count   := 0;
  v_ҽ��ids := Null;
  For r_Bill In c_Bill(n_�����־) Loop
    If Instr(',' || ���_In || ',', ',' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || ',') > 0 Or ���_In Is Null Then
      If r_Bill.���� = 1 Then
        If Nvl(r_Bill.ִ��״̬, 0) <> 1 Then
          l_����.Extend;
          l_����(l_����.Count) := r_Bill.Id;
        
          --Delete From ������ü�¼ Where ID = r_Bill.ID;
          n_Count := n_Count + 1; --��¼�Ƿ���ɾ����
        
          If r_Bill.ҽ����� Is Not Null Then
            If Instr(',' || Nvl(v_ҽ��ids, '') || ',', ',' || r_Bill.ҽ����� || ',') = 0 Then
              v_ҽ��ids := Nvl(v_ҽ��ids, '') || ',' || r_Bill.ҽ�����;
            End If;
          
            --��¼����ҽ�����Ѷ�Ӧ��ҽ��ID(����������)
            If n_ҽ��id Is Null Then
              n_ҽ��id := r_Bill.ҽ�����;
            End If;
          End If;
        Else
          If ���_In Is Not Null Then
            v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����Ѿ���ȫִ��,�������ʣ�';
            Raise Err_Item;
          End If;
          --���:û�޶��к�,ԭʼ�����а����Ѿ���ȫִ�е�
        End If;
      End If;
    End If;
  End Loop;

  --ɾ�����ۼ�¼
  Forall I In 1 .. l_����.Count
    Delete From ������ü�¼ Where ID = l_����(I);

  --ɾ��֮����ͳһ�������
  If n_Count > 0 Then
    n_Count := 1;
    For r_Serial In c_Serial Loop
      If r_Serial.�۸񸸺� Is Null Then
        n_���� := n_Count;
      End If;
    
      Update ������ü�¼
      Set ��� = n_Count, �۸񸸺� = Decode(�۸񸸺�, Null, Null, n_����)
      Where NO = No_In And ��¼���� = 2 And ��� = r_Serial.���;
    
      Update ������ü�¼ Set �������� = n_Count Where NO = No_In And ��¼���� = 2 And �������� = r_Serial.���;
    
      n_Count := n_Count + 1;
    End Loop;
  
  End If;

  --���ŵ���ȫ������ʱ��ɾ������ҽ������
  For c_ҽ�� In (Select Distinct ҽ�����
               From ������ü�¼
               Where NO = No_In And ��¼���� = 2 And ��¼״̬ = 3 And ҽ����� Is Not Null) Loop
    Select Nvl(Count(*), 0)
    Into n_Count
    From (Select ���, Sum(����) As ʣ������
           From (Select ��¼״̬, Nvl(�۸񸸺�, ���) As ���, Avg(Nvl(����, 1) * ����) As ����
                  From ������ü�¼
                  Where ��¼���� = 2 And ҽ����� + 0 = c_ҽ��.ҽ����� And NO = No_In
                  Group By ��¼״̬, Nvl(�۸񸸺�, ���))
           Group By ���
           Having Sum(����) <> 0);
  
    If n_Count = 0 Then
      Delete From ����ҽ������ Where ҽ��id = c_ҽ��.ҽ����� And ��¼���� = 2 And NO = No_In;
    End If;
  End Loop;

  If v_ҽ��ids Is Not Null Then
    --ҽ������
    --����_In    Integer:=0, --0:����;1-סԺ
    --����_In    Integer:=1, --1-�շѵ�;2-���ʵ�
    --����_In    Integer:=0, --0:ɾ�����۵�;1-�շѻ����;2-�˷ѻ�����
    --No_In      ������ü�¼.No%Type,
    --ҽ��ids_In Varchar2 := Null
    v_ҽ��ids := Substr(v_ҽ��ids, 2);
    Zl_ҽ������_�Ʒ�״̬_Update(0, 2, 0, No_In, v_ҽ��ids);
  Else
    Zl_ҽ������_�Ʒ�״̬_Update(0, 2, 2, No_In, v_ҽ��ids);
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_������ʼ�¼_Delete;
/

--108264:������,2017-04-17,���˸��ӷ��ô�������
Create Or Replace Procedure Zl_סԺ���ʼ�¼_Delete
(
  No_In           סԺ���ü�¼.No%Type,
  ���_In         Varchar2,
  ����Ա���_In   סԺ���ü�¼.����Ա���%Type,
  ����Ա����_In   סԺ���ü�¼.����Ա����%Type,
  ��¼����_In     סԺ���ü�¼.��¼����%Type := 2,
  ����״̬_In     Number := 0,
  ��Һ��ҩ���_In Number := 1,
  �Ǽ�ʱ��_In     סԺ���ü�¼.�Ǽ�ʱ��%Type := Sysdate
) As
  --���ܣ�����һ��סԺ���ʵ�����ָ�������
  --��ţ���ʽ��"1,3,5,7,8",��"1:2:33456,3:2,5:2,7:2,8:2",ð��ǰ������ֱ�ʾ�к�,�м�����ֱ�ʾ�˵�����,��������ֱ�ʾ��ҩ��¼��ID,Ŀǰ�����������ʱ�Ŵ���
  --      Ϊ�ձ�ʾ�������пɳ�����
  --��¼����:    2-�˹����ʵ�,3-�Զ����ʵ�
  --��Һ��ҩ���:    0-ҽ�����ã������ҩƷ�Ƿ������Һ��ҩ���ģ�1-��ҽ�����ã����ҩƷ�Ƿ������ҩ����
  --�ù����������ָ��������
  --����״̬_In:0-��ʾֱ������;1-��ʾ�������(ͨ����������-->�����������)
  --���α�ΪҪ�˷ѵ��ݵ�����ԭʼ��¼
  Cursor c_Bill Is
    Select ID, �۸񸸺�, ���, ִ��״̬, ��¼����, �շ����, ҽ�����, �շ�ϸĿid, ����id, ��ҳid, ������Ŀid, ��������id, ���˿���id, ִ�в���id, ���˲���id, ����, ����
    From סԺ���ü�¼
    Where NO = No_In And ��¼���� = ��¼����_In And ��¼״̬ In (0, 1, 3) And �����־ = 2
    Order By �շ�ϸĿid, ���;

  --���α����ڴ���ҩƷ����������
  --��Ҫ�ܷ��õ�ִ��״̬,��Ϊ���ڴ˲�����
  Cursor c_Stock(v_���_In Varchar2) Is
    Select ID, ����, NO, �ⷿid, ҩƷid, ����, ��ҩ��ʽ, ����, ʵ������, ���Ч��, Ч��, ����, ����, ��������, ����id, ��Ʒ����, �ڲ�����
    From ҩƷ�շ���¼
    Where NO = No_In And ���� In (9, 10, 25, 26) And Mod(��¼״̬, 3) = 1 And ����� Is Null And
          ����id In (Select ID
                   From סԺ���ü�¼
                   Where NO = No_In And ��¼���� = ��¼����_In And ��¼״̬ In (0, 1, 3) And �շ���� In ('4', '5', '6', '7') And
                         �����־ = 2 And (Instr(',' || v_���_In || ',', ',' || ��� || ',') > 0 Or v_���_In Is Null))
    Order By ҩƷid, �������� Desc;

  r_Stock c_Stock%RowType;
  --���α����ڴ�����ü�¼���
  Cursor c_Serial Is
    Select ���, �۸񸸺�
    From סԺ���ü�¼
    Where NO = No_In And ��¼���� = ��¼����_In And ��¼״̬ In (0, 1, 3)
    Order By ���;

  Cursor Cr_ҩƷ Is
    Select ID, ����, NO, �ⷿid, ҩƷid, ����, ��ҩ��ʽ, 0 As ����, ���Ч��, Ч��, ����, ����, ��������, ����id
    From ҩƷ�շ���¼
    Where Rownum <= 1;
  v_ҩƷ Cr_ҩƷ%RowType;

  v_ҽ��id     ����ҽ����¼.Id%Type;
  n_����       Number;
  v_����       סԺ���ü�¼.�۸񸸺�%Type;
  v_���       Varchar2(2000);
  v_Tmp        Varchar2(4000);
  v_ҽ��ids    Varchar2(4000);
  l_ҩƷ�շ�   t_Numlist := t_Numlist();
  l_����       t_Numlist := t_Numlist();
  l_����id     t_Numlist := t_Numlist();
  n_����       Number;
  n_����ⷿid ҩƷ�շ���¼.�ⷿid%Type;
  n_��������id ҩƷ�շ���¼.Id%Type;
  n_�ⷿid     ҩƷ�շ���¼.�ⷿid%Type;
  n_����ֵ     Number;
  --�����˷Ѽ������
  v_ʣ������ Number;
  v_ʣ��Ӧ�� Number;
  v_ʣ��ʵ�� Number;
  v_ʣ��ͳ�� Number;

  v_׼������ Number;
  v_�˷Ѵ��� Number;
  v_Ӧ�ս�� Number;
  v_ʵ�ս�� Number;
  v_ͳ���� Number;
  n_Temp     Number;
  n_�������� Number;
  v_Dec      Number;
  n_Count    Number;
  v_Curdate  Date;
  Err_Item Exception;
  v_Err_Msg        Varchar2(255);
  n_��������       Number;
  n_����id         ������ҳ.����id%Type;
  n_��ҳid         ������ҳ.��ҳid%Type;
  n_��˱�־       ������ҳ.��˱�־%Type;
  n_סԺ״̬       ������ҳ.״̬%Type;
  n_������˷�ʽ   Number(2);
  n_δ��ƽ�ֹ���� Number(2);
  v_��ҩid         Varchar2(4000);
  Type Ty_ҩƷ Is Ref Cursor;
  c_ҩƷ Ty_ҩƷ; --�α����

Begin
  --�������ʱ,��ҩƷ�ᴫ���кŵ���������
  If Not ���_In Is Null Then
    If Instr(���_In, ':') > 0 Then
      v_Tmp := ���_In || ',';
      While Not v_Tmp Is Null Loop
        v_��� := v_��� || ',' || Substr(v_Tmp, 1, Instr(v_Tmp, ':') - 1);
        If Instr(Substr(v_Tmp, Instr(v_Tmp, ':') + 1, Instr(v_Tmp, ',') - Instr(v_Tmp, ':') - 1), ':') > 0 Then
          v_��ҩid := v_��ҩid || ',' ||
                    Substr(v_Tmp, Instr(v_Tmp, ':', 1, 2) + 1, Instr(v_Tmp, ',') - Instr(v_Tmp, ':', 1, 2) - 1);
        End If;
        v_Tmp := Substr(v_Tmp, Instr(v_Tmp, ',') + 1);
      End Loop;
      v_��� := Substr(v_���, 2);
      If v_��ҩid Is Not Null Then
        v_��ҩid := Substr(v_��ҩid, 2);
      End If;
    Else
      v_��� := ���_In;
    End If;
  End If;

  --�Ƿ��Ѿ�ȫ����ȫִ��(ֻ�����ŵ��ݵļ��)
  Select Nvl(Count(*), 0), Nvl(Max(����id), 0), Nvl(Max(��ҳid), 0)
  Into n_Count, n_����id, n_��ҳid
  From סԺ���ü�¼
  Where NO = No_In And ��¼���� = ��¼����_In And ��¼״̬ In (0, 1, 3) And Nvl(ִ��״̬, 0) <> 1 And �����־ = 2;
  If n_Count = 0 Then
    v_Err_Msg := '�õ����е���Ŀ�Ѿ�ȫ����ȫִ�У�';
    Raise Err_Item;
  End If;

  n_������˷�ʽ   := Nvl(zl_GetSysParameter(185), 0);
  n_δ��ƽ�ֹ���� := Nvl(zl_GetSysParameter(215), 0);
  If n_������˷�ʽ = 1 Or n_δ��ƽ�ֹ���� = 1 Then
  
    Begin
      Select ��˱�־, ״̬ Into n_��˱�־, n_סԺ״̬ From ������ҳ Where ����id = n_����id And ��ҳid = n_��ҳid;
    Exception
      When Others Then
        n_��˱�־ := 0;
        n_סԺ״̬ := 0;
    End;
    If n_δ��ƽ�ֹ���� = 1 And n_סԺ״̬ = 1 Then
      v_Err_Msg := '����δ���,��ֹ�Բ�����ط��õĲ���!';
      Raise Err_Item;
    End If;
  
    If n_������˷�ʽ = 1 Then
    
      If Nvl(n_��˱�־, 0) = 1 Then
        v_Err_Msg := '�ò���Ŀǰ������˷���,���ܽ��з�����ص���!';
        Raise Err_Item;
      End If;
      If Nvl(n_��˱�־, 0) = 2 Then
        v_Err_Msg := '�ò���Ŀǰ�Ѿ�����˷������,���ܽ��з�����ص���!';
        Raise Err_Item;
      End If;
    End If;
  End If;

  --δ��ȫִ�е���Ŀ�Ƿ���ʣ������(ֻ�����ŵ��ݵļ��)
  Select Nvl(Count(*), 0)
  Into n_Count
  From (Select ���, Sum(����) As ʣ������
         From (Select ��¼״̬, Nvl(�۸񸸺�, ���) As ���, Avg(Nvl(����, 1) * ����) As ����
                From סԺ���ü�¼
                Where NO = No_In And ��¼���� = ��¼����_In And �����־ = 2 And
                      Nvl(�۸񸸺�, ���) In
                      (Select Nvl(�۸񸸺�, ���)
                       From סԺ���ü�¼
                       Where NO = No_In And ��¼���� = ��¼����_In And ��¼״̬ In (0, 1, 3) And Nvl(ִ��״̬, 0) <> 1)
                Group By ��¼״̬, Nvl(�۸񸸺�, ���))
         Group By ���
         Having Sum(����) <> 0);
  If n_Count = 0 Then
    v_Err_Msg := '�õ�����δ��ȫִ�в�����Ŀʣ������Ϊ��,û�п������ʵķ��ã�';
    Raise Err_Item;
  End If;

  --ҽ�����ã��������ִ�е�ҽ��(ע����ִ�е������������,��Ϊ���� ���_IN ����������ý���������)
  If Nvl(����״̬_In, 0) <> 1 Then
    --�������������̵ģ������ҽ��ִ��״̬
    Select Nvl(Count(*), 0)
    Into n_Count
    From ����ҽ������
    Where ִ��״̬ = 3 And (NO, ��¼����, ҽ��id) In
          (Select NO, ��¼����, ҽ�����
                        From סԺ���ü�¼
                        Where NO = No_In And ��¼���� = ��¼����_In And ��¼״̬ In (0, 1, 3) And ҽ����� Is Not Null And
                              (Instr(',' || v_��� || ',', ',' || ��� || ',') > 0 Or v_��� Is Null));
    If n_Count > 0 Then
      v_Err_Msg := 'Ҫ���ʵķ����д��ڶ�Ӧ��ҽ������ִ�е�������������ʣ�';
      Raise Err_Item;
    End If;
  End If;

  ---------------------------------------------------------------------------------
  --�ȴ�ҩƷ��Ӧ���ݼ�,��ȷ����ǰ������������,Ϊ�˴������ж�
  --�������α�������ȡ��"����� is Null"��������Ϊ�����ҩ���ܲ������ѷ�
  Open c_Stock(v_���);

  --���ñ���
  Select �Ǽ�ʱ��_In Into v_Curdate From Dual;

  --���С��λ��
  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into v_Dec From Dual;

  For c_��Ŀ���� In (Select a.����
                 From ������Ϣ A, ������ҳ B
                 Where a.����id = b.����id And b.��Ŀ���� Is Not Null And
                       (b.����id, b.��ҳid) In
                       (Select Distinct ����id, ��ҳid
                        From סԺ���ü�¼
                        Where NO = No_In And ��¼���� = ��¼����_In And ��¼״̬ In (0, 1, 3) And �����־ = 2)) Loop
    v_Err_Msg := '���ˡ�' || c_��Ŀ����.���� || '�� �Ѿ���������Ŀ,���ܱ����ʣ�';
    Raise Err_Item;
  End Loop;
  v_ҽ��ids := Null;
  --ѭ������ÿ�з���(������Ŀ��)
  For r_Bill In c_Bill Loop
    --����Ѿ����ڲ�����Ŀ��,���ܽ������ʴ���
    If Instr(',' || v_��� || ',', ',' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || ',') > 0 Or v_��� Is Null Then
      Select Decode(��¼״̬, 0, 1, 0) Into n_���� From סԺ���ü�¼ Where ID = r_Bill.Id;
      If Nvl(r_Bill.ִ��״̬, 0) <> 1 Then
        --��ʣ������,ʣ��Ӧ��,ʣ��ʵ��
        Select Sum(Nvl(����, 1) * ����), Sum(Ӧ�ս��), Sum(ʵ�ս��), Sum(ͳ����)
        Into v_ʣ������, v_ʣ��Ӧ��, v_ʣ��ʵ��, v_ʣ��ͳ��
        From סԺ���ü�¼
        Where NO = No_In And ��¼���� = ��¼����_In And ��� = r_Bill.���;
        n_�������� := 0;
        If v_ʣ������ = 0 Then
          If v_��� Is Not Null Then
            v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����Ѿ�ȫ�����ʣ�';
            Raise Err_Item;
          End If;
          --�����δ�޶��к�,ԭʼ�����еĸñ��Ѿ�ȫ������(ִ��״̬=0��һ�ֿ���)
        Else
        
          If Instr(���_In, ':') > 0 Then
            v_Tmp := ',' || ���_In;
            v_Tmp := Substr(v_Tmp, Instr(v_Tmp, ',' || r_Bill.��� || ':') + Length(',' || r_Bill.��� || ':'));
            v_Tmp := Substr(v_Tmp, 1, Instr(v_Tmp || ',', ',') - 1);
            If Instr(v_Tmp, ':') > 0 Then
              v_Tmp := Substr(v_Tmp, 1, Instr(v_Tmp, ':') - 1);
            End If;
            v_׼������ := v_Tmp;
            n_�������� := 1;
          End If;
        
          --׼������(��ҩƷ��ĿΪʣ������,ԭʼ����)
          If Instr(',4,5,6,7,', r_Bill.�շ����) = 0 Then
            If Instr(���_In, ':') = 0 Or ���_In Is Null Then
              v_׼������ := v_ʣ������;
            End If;
          Else
            --ҽ�������ջ�ʱ,���Ŀ���û�з���,���������ʵ��ǲ�������,����Ҫ�������Ϊ׼
            If Instr(���_In, ':') = 0 Or ���_In Is Null Then
              Select Nvl(Sum(Nvl(����, 1) * ʵ������), 0), Count(*)
              Into v_׼������, n_Count
              From ҩƷ�շ���¼
              Where NO = No_In And ���� In (9, 10, 25, 26) And Mod(��¼״̬, 3) = 1 And ����� Is Null And ����id = r_Bill.Id;
            End If;
          
            --��ʣ��������׼�������������������
            --1.���������õ������޶�Ӧ���շ���¼,��ʱʹ��ʣ������
            --2.��������,��ʱ�ѷ�ҩ����
            If v_׼������ = 0 Then
              If r_Bill.�շ���� = '4' Then
                If n_Count > 0 Then
                  v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����ѷ���,�����Ϻ����˷ѣ�';
                  Raise Err_Item;
                Else
                  v_׼������ := v_ʣ������;
                End If;
              Else
                v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����ѷ�ҩ,����ҩ�����˷ѣ�';
                Raise Err_Item;
              End If;
            End If;
          End If;
        
          --����סԺ���ü�¼
          If Nvl(n_����, 0) = 0 Then
            --����ʱ,ֱ�Ӹ�������,���Բ���黮��������
            --�ñ���Ŀ�ڼ�������
            Select Nvl(Max(Abs(ִ��״̬)), 0) + 1
            Into v_�˷Ѵ���
            From סԺ���ü�¼
            Where NO = No_In And ��¼���� = ��¼����_In And ��¼״̬ = 2 And ��� = r_Bill.��� And �����־ = 2;
          End If;
        
          --���=ʣ����*(׼����/ʣ����)
          v_Ӧ�ս�� := Round(v_ʣ��Ӧ�� * (v_׼������ / v_ʣ������), v_Dec);
          v_ʵ�ս�� := Round(v_ʣ��ʵ�� * (v_׼������ / v_ʣ������), v_Dec);
          v_ͳ���� := Round(v_ʣ��ͳ�� * (v_׼������ / v_ʣ������), v_Dec);
          If Nvl(n_����, 0) = 1 Then
            If Nvl(n_��������, 0) = 0 Then
              l_����.Extend;
              l_����(l_����.Count) := r_Bill.Id;
              n_����ֵ := 0;
            Else
              --��������
              --���۵�,�Ƚ���ص����ݴ������ڲ�����
              n_���� := 0;
              If r_Bill.���� > 1 Then
                --�������ҩ,���ڻ��տ϶��ǻ��յĸ���,�����Ǵ���.���,��Ҫ���׼�������Ƿ������ ��
                If Trunc(v_׼������ / r_Bill.����) <> (v_׼������ / r_Bill.����) Then
                  v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з���Ϊ��ҩ,�밴���������˷ѣ�';
                  Raise Err_Item;
                End If;
                n_���� := Trunc(v_׼������ / r_Bill.����);
                If Nvl(r_Bill.����, 0) - n_���� < 0 Then
                  v_׼������ := r_Bill.����;
                Else
                  v_׼������ := 0;
                End If;
              End If;
              Update סԺ���ü�¼
              Set ���� = ���� - n_����, ���� = ���� - v_׼������, Ӧ�ս�� = Nvl(Ӧ�ս��, 0) - v_Ӧ�ս��, ʵ�ս�� = Nvl(ʵ�ս��, 0) - v_ʵ�ս��,
                  �Ǽ�ʱ�� = v_Curdate, ͳ���� = Nvl(ͳ����, 0) - v_ͳ����
              Where ID = r_Bill.Id
              Returning Nvl(����, 0) * Nvl(����, 0) Into n_����ֵ;
            End If;
            If Nvl(n_����ֵ, 0) <= 0 Then
              l_����.Extend;
              l_����(l_����.Count) := r_Bill.Id;
            End If;
            If r_Bill.ҽ����� Is Not Null Then
              If Instr(',' || Nvl(v_ҽ��ids, '') || ',', ',' || r_Bill.ҽ����� || ',') = 0 Then
                v_ҽ��ids := Nvl(v_ҽ��ids, '') || ',' || r_Bill.ҽ�����;
              End If;
              --��¼����ҽ�����Ѷ�Ӧ��ҽ��ID(����������)
              If v_ҽ��id Is Null Then
                v_ҽ��id := r_Bill.ҽ�����;
              End If;
            End If;
          
          End If;
        
          If Nvl(n_����, 0) = 0 Then
            --����ʱ,ֱ�Ӹ�������,���Բ���黮��������
            --�����˷Ѽ�¼
            Insert Into סԺ���ü�¼
              (ID, NO, ��¼����, ��¼״̬, ���, ��������, �۸񸸺�, ��ҳid, ����id, ҽ�����, �����־, �ಡ�˵�, Ӥ����, ����, �Ա�, ����, ��ʶ��, ����, �ѱ�, ���˲���id,
               ���˿���id, �շ����, �շ�ϸĿid, ���㵥λ, ����, ��ҩ����, ����, �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���, ��׼����, Ӧ�ս��, ʵ�ս��, ��������id, ������,
               ִ�в���id, ������, ִ����, ִ��״̬, ִ��ʱ��, ����Ա���, ����Ա����, ����ʱ��, �Ǽ�ʱ��, ������Ŀ��, ���մ���id, ͳ����, ���ձ���, ���ʵ�id, ժҪ, ��������, �Ƿ���,
               ����, ҽ��С��id)
              Select ���˷��ü�¼_Id.Nextval, NO, ��¼����, 2, ���, ��������, �۸񸸺�, ��ҳid, ����id, ҽ�����, �����־, �ಡ�˵�, Ӥ����, ����, �Ա�, ����, ��ʶ��,
                     ����, �ѱ�, ���˲���id, ���˿���id, �շ����, �շ�ϸĿid, ���㵥λ, Decode(Sign(v_׼������ - Nvl(����, 1) * ����), 0, ����, 1), ��ҩ����,
                     Decode(Sign(v_׼������ - Nvl(����, 1) * ����), 0, -1 * ����, -1 * v_׼������), �Ӱ��־, ���ӱ�־, ������Ŀid, �վݷ�Ŀ, ���ʷ���,
                     ��׼����, -1 * v_Ӧ�ս��, -1 * v_ʵ�ս��, ��������id, ������, ִ�в���id, ������, ִ����, -1 * v_�˷Ѵ���, ִ��ʱ��, ����Ա���_In,
                     ����Ա����_In, ����ʱ��, v_Curdate, ������Ŀ��, ���մ���id, -1 * v_ͳ����, ���ձ���, ���ʵ�id, ժҪ, ��������, �Ƿ���, ����, ҽ��С��id
              From סԺ���ü�¼
              Where ID = r_Bill.Id;
          
            --��¼����ҽ�����Ѷ�Ӧ��ҽ��ID(����������)
            If v_ҽ��id Is Null And r_Bill.ҽ����� Is Not Null Then
              v_ҽ��id := r_Bill.ҽ�����;
            End If;
          
            Update ����������Ŀ
            Set �������� = Nvl(��������, 0) - v_׼������
            Where ����id = r_Bill.����id And ��ҳid = r_Bill.��ҳid And ��Ŀid = r_Bill.�շ�ϸĿid And Nvl(ʹ������, 0) <> 0;
          
            --�������
            Update �������
            Set ������� = Nvl(�������, 0) - v_ʵ�ս��
            Where ����id = r_Bill.����id And ���� = 2 And ���� = 1;
            If Sql%RowCount = 0 Then
              Insert Into �������
                (����id, ����, ����, �������, Ԥ�����)
              Values
                (r_Bill.����id, 2, 1, -1 * v_ʵ�ս��, 0);
            End If;
          
            --����δ�����
            Update ����δ�����
            Set ��� = Nvl(���, 0) - v_ʵ�ս��
            Where ����id = r_Bill.����id And Nvl(��ҳid, 0) = Nvl(r_Bill.��ҳid, 0) And Nvl(���˲���id, 0) = Nvl(r_Bill.���˲���id, 0) And
                  Nvl(���˿���id, 0) = Nvl(r_Bill.���˿���id, 0) And Nvl(��������id, 0) = Nvl(r_Bill.��������id, 0) And
                  Nvl(ִ�в���id, 0) = Nvl(r_Bill.ִ�в���id, 0) And ������Ŀid + 0 = r_Bill.������Ŀid And ��Դ;�� + 0 = 2;
            If Sql%RowCount = 0 Then
              Insert Into ����δ�����
                (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
              Values
                (r_Bill.����id, r_Bill.��ҳid, r_Bill.���˲���id, r_Bill.���˿���id, r_Bill.��������id, r_Bill.ִ�в���id, r_Bill.������Ŀid, 2,
                 -1 * v_ʵ�ս��);
            End If;
          
            --���ԭ���ü�¼
            --ִ��״̬:ȫ������(׼����=ʣ����)���Ϊ0,���򱣳�ԭ״̬
            If Instr(',4,5,6,7,', r_Bill.�շ����) = 0 Then
              --һ�������ҩƷ�����ĵ���Ŀ,�����ڲ������ʵ����,ֻ������������������ʱ,�Ż���ֲ�������,����
              --ִ��״ֻ̬������:0.δִ��;1��ִ��;
              --������������˹����н���ִ��ǿ�Ƹ�Ϊ��2����ִ��,�����Ҫ�ڴ˴���Ϊ1��ִ��.δִ�еĲ���.
              Update סԺ���ü�¼
              Set ��¼״̬ = 3, ִ��״̬ = Decode(Sign(v_׼������ - v_ʣ������), 0, 0, Decode(ִ��״̬, 2, 1, ִ��״̬))
              Where ID = r_Bill.Id;
            Else
              Update סԺ���ü�¼
              Set ��¼״̬ = 3, ִ��״̬ = Decode(Sign(v_׼������ - v_ʣ������), 0, 0, ִ��״̬)
              Where ID = r_Bill.Id;
            End If;
          End If;
        End If;
      Else
        If v_��� Is Not Null Then
          v_Err_Msg := '�����е�' || Nvl(r_Bill.�۸񸸺�, r_Bill.���) || '�з����Ѿ���ȫִ��,�������ʣ�';
          Raise Err_Item;
        End If;
        --���:û�޶��к�,ԭʼ�����а����Ѿ���ȫִ�е�
      End If;
    End If;
  End Loop;

  --��������ҩID,����ҩƷ�Ƿ�����Һ��ҩ����
  If v_��ҩid Is Null And ��Һ��ҩ���_In = 1 Then
    For v_���� In (Select ID
                 From סԺ���ü�¼
                 Where NO = No_In And ��¼���� = ��¼����_In And ��¼״̬ In (0, 1, 3) And �շ���� In ('4', '5', '6', '7') And �����־ = 2 And
                       (Instr(',' || v_��� || ',', ',' || ��� || ',') > 0 Or v_��� Is Null)) Loop
      Begin
        Select Count(1)
        Into n_Count
        From ��Һ��ҩ���� A, ҩƷ�շ���¼ B
        Where a.�շ�id = b.Id And b.����id = v_����.Id And Instr(',8,9,10,21,24,25,26,', ',' || b.���� || ',') > 0;
      Exception
        When Others Then
          n_Count := 0;
      End;
      If n_Count <> 0 Then
        v_Err_Msg := '�����Ѿ�������Һ��ҩ���ĵĴ�����ҩƷ���޷�������ʣ�';
        Raise Err_Item;
      End If;
    End Loop;
  End If;

  n_�������� := 0;
  ---------------------------------------------------------------------------------
  --ҩƷ��ش���:��Ҫ�Ƕ����������Ч.(�����ǲ���)
  For v_���� In (Select ID, ���, �շ����
               From סԺ���ü�¼
               Where NO = No_In And ��¼���� = ��¼����_In And ��¼״̬ In (0, 1, 3) And �շ���� In ('4', '5', '6', '7') And �����־ = 2 And
                     (Instr(',' || v_��� || ',', ',' || ��� || ',') > 0 Or v_��� Is Null)
               Order By �շ�ϸĿid) Loop
    --���ݷ���ID��������صĴ���
    v_׼������ := 0;
    If Instr(���_In, ':') > 0 Then
      v_Tmp := ',' || ���_In;
      v_Tmp := Substr(v_Tmp, Instr(v_Tmp, ',' || v_����.��� || ':') + Length(',' || v_����.��� || ':'));
      v_Tmp := Substr(v_Tmp, 1, Instr(v_Tmp || ',', ',') - 1);
      If Instr(v_Tmp, ':') > 0 Then
        v_Tmp := Substr(v_Tmp, 1, Instr(v_Tmp, ':') - 1);
      End If;
      v_׼������ := v_Tmp;
    End If;
    If v_׼������ <> 0 Then
      n_�������� := 1;
      n_Temp     := 0;
      --------------------------------------------------------------------------------------
      --����Ƿ񱸻���������,��������
      -- a.������ڴ���δ��˵����������Ҳ�������ʱ,ֱ����ԭ���Ļ����ϸ���������������
      -- b.������ڴ���δ��˵�������������ȫ����ʱ,ֱ��ɾ��
      -- c.��洦��:��ԭΪ����ⷿ�Ŀ�������;���ϲ��Ų�����
      -- d.����Ѿ�������,���ʱ�������������ⵥ�Ѿ����,��˾Ͱ����������ת,���ָ������ϲ�����
      n_����ⷿid := Null;
      n_��������id := Null;
      If v_����.�շ���� = '4' Then
        Begin
          Select 1, �ⷿid, ID
          Into n_��������, n_����ⷿid, n_��������id
          From ҩƷ�շ���¼
          Where ����id = v_����.Id And ������� Is Null And ���� = 21 And Rownum = 1;
        Exception
          When Others Then
            n_�������� := 0;
        End;
      Else
        n_�������� := 0;
      End If;
      --------------------------------------------------------------------------------------
      If v_��ҩid Is Not Null Then
        Open c_ҩƷ For
          Select /*+ rule*/
           a.Id, a.����, a.No, a.�ⷿid, a.ҩƷid, a.����, a.��ҩ��ʽ,
           Decode(a.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(a.����, 1) * Nvl(a.ʵ������, 0) As ����, a.���Ч��, a.Ч��, a.����, a.����, a.��������,
           a.����id
          From ҩƷ�շ���¼ A, Table(f_Str2list(v_��ҩid)) B, ��Һ��ҩ���� C
          Where a.No = No_In And a.���� In (9, 10, 25, 26) And Mod(a.��¼״̬, 3) = 1 And a.����� Is Null And a.����id = v_����.Id And
                a.Id = c.�շ�id And c.��¼id = b.Column_Value
          Order By ��������;
      Else
        Open c_ҩƷ For
          Select ID, ����, NO, �ⷿid, ҩƷid, ����, ��ҩ��ʽ, Decode(��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(����, 1) * Nvl(ʵ������, 0) As ����,
                 ���Ч��, Ч��, ����, ����, ��������, ����id
          From ҩƷ�շ���¼
          Where NO = No_In And ���� In (9, 10, 25, 26) And Mod(��¼״̬, 3) = 1 And ����� Is Null And ����id = v_����.Id
          Order By ��������;
      End If;
      Loop
        Fetch c_ҩƷ
          Into v_ҩƷ;
        Exit When c_ҩƷ%NotFound;
        n_Temp := v_ҩƷ.����;
        If v_׼������ >= n_Temp Then
          l_ҩƷ�շ�.Extend;
          l_ҩƷ�շ�(l_ҩƷ�շ�.Count) := v_ҩƷ.Id;
          If Nvl(n_��������id, 0) > 0 Then
            l_ҩƷ�շ�.Extend;
            l_ҩƷ�շ�(l_ҩƷ�շ�.Count) := n_��������id;
          End If;
          v_׼������ := v_׼������ - n_Temp;
        Else
          If v_����.�շ���� = '7' Then
            --��ǰ�е�����Ҫ��
            Update ҩƷ�շ���¼
            Set ���� = 1, ʵ������ = Decode(����, Null, 1, 0, 1, ����) * Nvl(ʵ������, 0) - v_׼������,
                ��д���� = Decode(����, Null, 1, 0, 1, ����) * Nvl(��д����, 0) - v_׼������,
                �ɱ���� =
                 (Decode(����, Null, 1, 0, 1, ����) * Nvl(ʵ������, 0) - v_׼������) * �ɱ���,
                ���۽�� =
                 (Decode(����, Null, 1, 0, 1, ����) * Nvl(ʵ������, 0) - v_׼������) * ���ۼ�,
                ��� = Round((Decode(����, Null, 1, 0, 1, ����) * Nvl(ʵ������, 0) - v_׼������) * ���ۼ� -
                            (Decode(����, Null, 1, 0, 1, ����) * Nvl(ʵ������, 0) - v_׼������) * �ɱ���, 5)
            Where ID = v_ҩƷ.Id;
          Else
            Update ҩƷ�շ���¼
            Set ʵ������ = Nvl(ʵ������, 0) - v_׼������, ��д���� = Nvl(��д����, 0) - v_׼������,
                �ɱ���� =
                 (Nvl(ʵ������, 0) - v_׼������) * �ɱ���,
                ���۽�� =
                 (Nvl(ʵ������, 0) - v_׼������) * ���ۼ�,
                ��� = Round((Nvl(ʵ������, 0) - v_׼������) * ���ۼ� - (Nvl(ʵ������, 0) - v_׼������) * �ɱ���, 5)
            Where ID = v_ҩƷ.Id;
          End If;
          --�����������ⵥ
          If Nvl(n_��������id, 0) <> 0 Then
            If v_����.�շ���� = '7' Then
              Update ҩƷ�շ���¼
              Set ���� = 1, ʵ������ = Decode(����, Null, 1, 0, 1, ����) * Nvl(ʵ������, 0) - v_׼������,
                  ��д���� = Decode(����, Null, 1, 0, 1, ����) * Nvl(ʵ������, 0) - v_׼������,
                  �ɱ���� =
                   (Decode(����, Null, 1, 0, 1, ����) * Nvl(ʵ������, 0) - v_׼������) * �ɱ���,
                  ���۽�� =
                   (Decode(����, Null, 1, 0, 1, ����) * Nvl(ʵ������, 0) - v_׼������) * ���ۼ�,
                  ��� = Round((Decode(����, Null, 1, 0, 1, ����) * Nvl(ʵ������, 0) - v_׼������) * ���ۼ� -
                              (Decode(����, Null, 1, 0, 1, ����) * Nvl(ʵ������, 0) - v_׼������) * �ɱ���, 5)
              Where ID = Nvl(n_��������id, 0);
            Else
              Update ҩƷ�շ���¼
              Set ʵ������ = Nvl(ʵ������, 0) - v_׼������, ��д���� = Nvl(ʵ������, 0) - v_׼������,
                  �ɱ���� =
                   (Nvl(ʵ������, 0) - v_׼������) * �ɱ���,
                  ���۽�� =
                   (Nvl(ʵ������, 0) - v_׼������) * ���ۼ�,
                  ��� = Round((Nvl(ʵ������, 0) - v_׼������) * ���ۼ� - (Nvl(ʵ������, 0) - v_׼������) * �ɱ���, 5)
              Where ID = Nvl(n_��������id, 0);
            End If;
          End If;
          n_Temp     := v_׼������;
          v_׼������ := 0;
        End If;
        If Nvl(n_��������, 0) = 1 Then
          n_�ⷿid := n_����ⷿid;
        Else
          n_�ⷿid := v_ҩƷ.�ⷿid;
        End If;
      
        If n_�ⷿid Is Not Null Then
          Update ҩƷ���
          Set �������� = Nvl(��������, 0) + n_Temp
          Where �ⷿid = n_�ⷿid And ҩƷid = v_ҩƷ.ҩƷid And Nvl(����, 0) = Nvl(v_ҩƷ.����, 0) And ���� = 1;
          If Sql%RowCount = 0 Then
            Insert Into ҩƷ���
              (�ⷿid, ҩƷid, ����, ����, Ч��, ��������, �ϴ�����, �ϴβ���, ���Ч��)
            Values
              (n_�ⷿid, v_ҩƷ.ҩƷid, 1, v_ҩƷ.����, v_ҩƷ.Ч��, n_Temp, v_ҩƷ.����, v_ҩƷ.����, v_ҩƷ.���Ч��);
          End If;
          Delete ҩƷ���
          Where �ⷿid = n_�ⷿid And ҩƷid = v_ҩƷ.ҩƷid And Nvl(����, 0) = Nvl(v_ҩƷ.����, 0) And ���� = 1 And Nvl(��������, 0) = 0 And
                Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0 And Nvl(ʵ�ʲ��, 0) = 0;
        End If;
      
        If Nvl(n_��������, 0) = 1 Then
          Update ҩƷ���
          Set �������� = Nvl(��������, 0) + n_Temp
          Where �ⷿid = v_ҩƷ.�ⷿid And ҩƷid = v_ҩƷ.ҩƷid And Nvl(����, 0) = Nvl(v_ҩƷ.����, 0) And ���� = 1;
          If Sql%RowCount = 0 Then
            Insert Into ҩƷ���
              (�ⷿid, ҩƷid, ����, ����, Ч��, ��������, �ϴ�����, �ϴβ���, ���Ч��)
            Values
              (v_ҩƷ.�ⷿid, v_ҩƷ.ҩƷid, 1, v_ҩƷ.����, v_ҩƷ.Ч��, n_Temp, v_ҩƷ.����, v_ҩƷ.����, v_ҩƷ.���Ч��);
          End If;
        End If;
      
        If v_׼������ = 0 Then
          Exit;
        End If;
      End Loop;
      --���������ĵ�,�����:��Ϊ������Ļ�,������ҩƷ�շ���¼�д���
      If Nvl(v_׼������, 0) <> 0 And Not (v_����.�շ���� = '4' And n_Temp = 0) Then
        --δ�������,��ʾ��ҩƷ�����Ѿ�ִ��.
        v_Err_Msg := 'Ҫ���ʵķ����д����ѷ���ҩƷ�����ģ����ѱ����������ʣ�������ǲ�����������ġ�';
        Raise Err_Item;
      End If;
    End If;
  End Loop;

  If n_�������� = 0 Then
    ------------------------------------------------------------------------------------------------------------------------
    --�ȴ���������
    For v_���� In (Select ID, ����, NO, �ⷿid, ҩƷid, ����, ��ҩ��ʽ, ����, ʵ������, ���Ч��, Ч��, ����, ����, ��������, ����id, ��Ʒ����, �ڲ�����
                 From ҩƷ�շ���¼
                 Where ���� = 21 And Mod(��¼״̬, 3) = 1 And ����� Is Null And
                       ����id In (Select ID
                                From סԺ���ü�¼
                                Where NO = No_In And ��¼���� = ��¼����_In And ��¼״̬ In (0, 1, 3) And �շ���� = '4' And �����־ = 2 And
                                      (Instr(',' || v_��� || ',', ',' || ��� || ',') > 0 Or v_��� Is Null))
                 Order By ҩƷid, �������� Desc) Loop
      --����ҩƷ���
      If v_����.�ⷿid Is Not Null Then
        Update ҩƷ���
        Set �������� = Nvl(��������, 0) + Decode(v_����.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(v_����.����, 1) * Nvl(v_����.ʵ������, 0)
        Where �ⷿid = v_����.�ⷿid And ҩƷid = v_����.ҩƷid And Nvl(����, 0) = Nvl(v_����.����, 0) And ���� = 1;
        If Sql%RowCount = 0 Then
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, Ч��, ��������, �ϴ�����, �ϴβ���, ���Ч��, ��Ʒ����, �ڲ�����)
          Values
            (v_����.�ⷿid, v_����.ҩƷid, 1, v_����.����, v_����.Ч��,
             Decode(v_����.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(v_����.����, 1) * Nvl(v_����.ʵ������, 0), v_����.����, v_����.����, v_����.���Ч��,
             v_����.��Ʒ����, v_����.�ڲ�����);
        End If;
        Delete ҩƷ���
        Where �ⷿid = v_����.�ⷿid And ҩƷid = v_����.ҩƷid And Nvl(����, 0) = Nvl(v_����.����, 0) And ���� = 1 And Nvl(��������, 0) = 0 And
              Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0 And Nvl(ʵ�ʲ��, 0) = 0;
      End If;
      l_����id.Extend;
      l_����id(l_����id.Count) := v_����.����id;
      l_ҩƷ�շ�.Extend;
      l_ҩƷ�շ�(l_ҩƷ�շ�.Count) := v_����.Id;
    End Loop;
  
    --ҩƷ�������
    Fetch c_Stock
      Into r_Stock;
    While c_Stock%Found Loop
    
      --����ҩƷ���
      If r_Stock.�ⷿid Is Not Null Then
      
        Select Decode(Count(Column_Value), Null, 0, 0, 0, 1)
        Into n_��������
        From Table(l_����id)
        Where Column_Value = r_Stock.����id;
      
        Update ҩƷ���
        Set �������� = Nvl(��������, 0) + Decode(r_Stock.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(r_Stock.����, 1) * Nvl(r_Stock.ʵ������, 0)
        Where �ⷿid = r_Stock.�ⷿid And ҩƷid = r_Stock.ҩƷid And Nvl(����, 0) = Nvl(r_Stock.����, 0) And ���� = 1;
        If Sql%RowCount = 0 Then
          Insert Into ҩƷ���
            (�ⷿid, ҩƷid, ����, ����, Ч��, ��������, �ϴ�����, �ϴβ���, ���Ч��)
          Values
            (r_Stock.�ⷿid, r_Stock.ҩƷid, 1, r_Stock.����, r_Stock.Ч��,
             Decode(r_Stock.��ҩ��ʽ, Null, 1, -1, 0, 1) * Nvl(r_Stock.����, 1) * Nvl(r_Stock.ʵ������, 0), r_Stock.����, r_Stock.����,
             r_Stock.���Ч��);
        End If;
        Delete ҩƷ���
        Where �ⷿid = r_Stock.�ⷿid And ҩƷid = r_Stock.ҩƷid And Nvl(����, 0) = Nvl(r_Stock.����, 0) And ���� = 1 And
              Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0 And Nvl(ʵ�ʲ��, 0) = 0;
      End If;
    
      --ɾ��ҩƷ�շ���¼(���ϲ����������:����� Is Null)
      --Delete From ҩƷ�շ���¼ Where ID = r_Stock.ID And ����� Is Null;
    
      l_ҩƷ�շ�.Extend;
      l_ҩƷ�շ�(l_ҩƷ�շ�.Count) := r_Stock.Id;
      Fetch c_Stock
        Into r_Stock;
    End Loop;
    Close c_Stock;
  
    --ɾ��ҩƷ�շ���¼
    Forall I In 1 .. l_ҩƷ�շ�.Count
      Delete From ҩƷ�շ���¼ Where ID = l_ҩƷ�շ�(I) And ����� Is Null;
    If Sql%RowCount <> l_ҩƷ�շ�.Count And l_ҩƷ�շ�.Count <> 0 Then
      v_Err_Msg := 'Ҫ���ʵķ����д����ѷ���ҩƷ�����ģ����ѱ����������ʣ�������ǲ�����������ġ�';
      Raise Err_Item;
    End If;
  Else
    --ɾ��ҩƷ�շ���¼
    Forall I In 1 .. l_ҩƷ�շ�.Count
      Delete From ҩƷ�շ���¼ Where ID = l_ҩƷ�շ�(I) And ����� Is Null;
  End If;
  --δ��ҩƷ��¼
  Delete From δ��ҩƷ��¼ A
  Where NO = No_In And ���� In (9, 10, 25, 26) And Not Exists
   (Select 1
         From ҩƷ�շ���¼
         Where ���� = a.���� And Nvl(�ⷿid, 0) = Nvl(a.�ⷿid, 0) And NO = No_In And Mod(��¼״̬, 3) = 1 And ����� Is Null);

  ---------------------------------------------------------------------------------
  --����ǻ���,ֱ��ɾ�����ü�¼(ҩƷ�����)
  n_Count := l_����.Count;
  --ɾ�����ۼ�¼
  Forall I In 1 .. l_����.Count
    Delete From סԺ���ü�¼ Where ID = l_����(I);

  --ɾ��֮����ͳһ�������
  If n_Count > 0 Then
    n_Count := 1;
    For r_Serial In c_Serial Loop
      If r_Serial.�۸񸸺� Is Null Then
        v_���� := n_Count;
      End If;
    
      Update סԺ���ü�¼
      Set ��� = n_Count, �۸񸸺� = Decode(�۸񸸺�, Null, Null, v_����)
      Where NO = No_In And ��¼���� = ��¼����_In And ��� = r_Serial.���;
    
      Update סԺ���ü�¼
      Set �������� = n_Count
      Where NO = No_In And ��¼���� = ��¼����_In And �������� = r_Serial.���;
    
      n_Count := n_Count + 1;
    End Loop;
  
  End If;

  --���ŵ���ȫ������ʱ��ɾ������ҽ������
  For c_ҽ�� In (Select Distinct ҽ�����
               From סԺ���ü�¼
               Where NO = No_In And ��¼���� = 2 And ��¼״̬ = 3 And ҽ����� Is Not Null) Loop
    Select Nvl(Count(*), 0)
    Into n_Count
    From (Select ���, Sum(����) As ʣ������
           From (Select ��¼״̬, Nvl(�۸񸸺�, ���) As ���, Avg(Nvl(����, 1) * ����) As ����
                  From סԺ���ü�¼
                  Where ��¼���� = 2 And ҽ����� + 0 = c_ҽ��.ҽ����� And NO = No_In
                  Group By ��¼״̬, Nvl(�۸񸸺�, ���))
           Group By ���
           Having Sum(����) <> 0);
  
    If n_Count = 0 Then
      Delete From ����ҽ������ Where ҽ��id = c_ҽ��.ҽ����� And ��¼���� = 2 And NO = No_In;
    End If;
  End Loop;

  If v_ҽ��ids Is Not Null Then
    --ҽ������
    --����_In    Integer:=0, --0:����;1-סԺ
    --����_In    Integer:=1, --1-�շѵ�;2-���ʵ�
    --����_In    Integer:=0, --0:ɾ�����۵�;1-�շѻ����;2-�˷ѻ�����
    --No_In      ������ü�¼.No%Type,
    --ҽ��ids_In Varchar2 := Null
    v_ҽ��ids := Substr(v_ҽ��ids, 2);
    Zl_ҽ������_�Ʒ�״̬_Update(1, 2, 0, No_In, v_ҽ��ids);
  Else
    Zl_ҽ������_�Ʒ�״̬_Update(1, 2, 2, No_In);
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_סԺ���ʼ�¼_Delete;
/

--107321:Ƚ����,2017-04-14,�޸��ޡ����п��ҡ�Ȩ��ʱ�Ĺ��ܿ�������
Create Or Replace Procedure Zl_�ٴ������_Delete
(
  Id_In     �ٴ������.Id%Type,
  ��Աid_In ��Ա��.Id%Type := Null,
  վ��_In   ���ű�.վ��%Type
) As
  --���ܣ�ɾ���ٴ������ 
  --������ 
  --        ��Աid_In ���̶���������Ч����Ϊ0��null��ʾ�ٴ�������Ա��ɾ�� 
  n_Count    Number;
  n_�Ű෽ʽ �ٴ������.�Ű෽ʽ%Type;
  n_����id   �ٴ������.Id%Type;

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  l_��¼id t_Numlist := t_Numlist();
  l_����id t_Numlist := t_Numlist();
Begin
  Begin
    Select 1 Into n_Count From �ٴ������ Where �Ű෽ʽ <> 3 And ������ Is Not Null And ID = Id_In;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If n_Count <> 0 Then
    v_Err_Msg := '��ǰ������ѷ���������ɾ����';
    Raise Err_Item;
  End If;

  Begin
    Select �Ű෽ʽ Into n_�Ű෽ʽ From �ٴ������ Where ID = Id_In;
  Exception
    When Others Then
      v_Err_Msg := '�������Ϣδ�ҵ���';
      Raise Err_Item;
  End;

  --�����Ű����ģ�����ݱ����ڳ����¼�е� 
  If Nvl(n_�Ű෽ʽ, 0) In (0, 3) Then
    --�̶�����/ģ�� 
    --ɾ���ٴ��������� 
    Select b.Id Bulk Collect
    Into l_����id
    From �ٴ����ﰲ�� A, �ٴ��������� B
    Where a.Id = b.����id And a.����id = Id_In;
  
    Forall I In 1 .. l_����id.Count
      Delete From �ٴ�����ʱ�� Where ����id = l_����id(I);
  
    Forall I In 1 .. l_����id.Count
      Delete From �ٴ��������� Where ����id = l_����id(I);
  
    Forall I In 1 .. l_����id.Count
      Delete From �ٴ�����Һſ��� Where ����id = l_����id(I);
  
    Forall I In 1 .. l_����id.Count
      Delete From �ٴ��������� Where ID = l_����id(I);
  
    --ɾ���ٴ����ﰲ�� 
    Delete From �ٴ����ﰲ�� Where ����id = Id_In;
  
    --ɾ���ٴ������ 
    Delete �ٴ������ Where ID = Id_In;
  
    Return;
  End If;

  --======================================================================================================== 
  --�³����/�ܳ���� 
  --�³����/�ܳ����ֻ�ܴ����һ����ʼɾ�� 
  Begin
    Select ID
    Into n_����id
    From (Select a.Id
           From �ٴ������ A, �ٴ����ﰲ�� B, �ٴ������Դ C, ���ű� D
           Where a.�Ű෽ʽ = n_�Ű෽ʽ And a.Id = b.����id(+) And b.��Դid = c.Id(+) And c.����id = d.Id(+)
                --��ǰ��Ա�ɲ����ĺ�Դ 
                 And (Nvl(��Աid_In, 0) = 0 Or (Nvl(c.�Ƿ��ٴ��Ű�, 0) = 1 And Exists
                  (Select 1 From ������Ա Where ����id = c.����id And ��Աid = ��Աid_In)))
                --վ�� 
                 And (d.վ�� Is Null Or d.վ�� = վ��_In)
           Order By a.��� Desc, a.�·� Desc, a.���� Desc)
    Where Rownum < 2;
  Exception
    When Others Then
      n_����id := 0;
  End;
  If Nvl(n_����id, 0) <> 0 And Nvl(n_����id, 0) <> Id_In Then
    v_Err_Msg := '��������һ�������ʼɾ����';
    Raise Err_Item;
  End If;

  If Nvl(��Աid_In, 0) <> 0 Then
    --û��"���п���"Ȩ��
    Select Count(1)
    Into n_Count
    From �ٴ����ﰲ�� A, �ٴ������Դ B
    Where a.��Դid = b.Id And a.����id = Id_In And
          Not (Nvl(b.�Ƿ��ٴ��Ű�, 0) = 1 And Exists (Select 1 From ������Ա Where ����id = b.����id And ��Աid = ��Աid_In)) And
          Rownum < 2;
    If n_Count <> 0 Then
      v_Err_Msg := '��ǰ������к���������Ա�Ѿ��ƶ��İ��ţ�����ɾ����';
      Raise Err_Item;
    End If;
  End If;

  --ɾ���ٴ������¼ 
  Select a.Id Bulk Collect
  Into l_��¼id
  From �ٴ������¼ A, �ٴ����ﰲ�� B, �ٴ������Դ C, ���ű� D
  Where a.����id = b.Id And a.��Դid = c.Id And c.����id = d.Id And b.����id = Id_In
       --��ǰ��Ա�ɲ����ĺ�Դ 
        And (Nvl(��Աid_In, 0) = 0 Or
        (Nvl(c.�Ƿ��ٴ��Ű�, 0) = 1 And Exists (Select 1 From ������Ա Where c.����id = ����id And ��Աid = ��Աid_In)))
       --վ�� 
        And (d.վ�� Is Null Or d.վ�� = վ��_In);

  Zl_�ٴ������¼_Batchdelete(l_��¼id);

  --ɾ���ٴ����ﰲ�� 
  Delete From �ٴ����ﰲ�� A
  Where a.����id = Id_In And Exists
   (Select 1
         From �ٴ������Դ B, ���ű� D
         Where a.��Դid = b.Id And b.����id = d.Id
              --��ǰ��Ա�ɲ����ĺ�Դ 
               And (Nvl(��Աid_In, 0) = 0 Or (Nvl(b.�Ƿ��ٴ��Ű�, 0) = 1 And Exists
                (Select 1 From ������Ա Where b.����id = ����id And ��Աid = ��Աid_In)))
              --վ�� 
               And (d.վ�� Is Null Or d.վ�� = վ��_In));

  --ɾ���ٴ������ 
  Delete �ٴ������ A
  Where a.Id = Id_In And Not Exists (Select 1 From �ٴ����ﰲ�� Where ����id = a.Id And ��Դid Is Not Null);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ������_Delete;
/

--109289:Ƚ����,2017-05-23,ʹ�á�ȫ��������ſ��ơ�����ʱ������������ŵ�û�����÷�ʱ�εİ��ţ�û�����ɶ�Ӧ��ʱ���������
--107321:Ƚ����,2017-04-14,�޸��ޡ����п��ҡ�Ȩ��ʱ�Ĺ��ܿ�������
Create Or Replace Procedure Zl_�ٴ����ﰲ��_��ſ���
(
  ����id_In   �ٴ������.Id%Type,
  ��ſ���_In �ٴ���������.�Ƿ���ſ���%Type,
  վ��_In     ���ű�.վ��%Type := Null,
  ��Աid_In   ��Ա��.Id%Type := 0
) As
  --ȫ��������ſ��ƻ���ȫ��ȡ����ſ���
  --������
  --      ��Աid_In ������0���޸���Ա���ڿ��ҵ����к�Դ���ţ������޸����к�Դ�İ���
  n_Count    Number(2);
  n_�����¼ Number(2);

  Err_Item Exception;
  v_Err_Msg Varchar2(255);

  l_����id t_Numlist := t_Numlist();
  l_��¼id t_Numlist := t_Numlist();

  --���α����ڶ�ȡ�����ٴ����ﰲ�ŵ�ID
  Cursor c_����
  (
    ����id_In �ٴ������.Id%Type,
    ��Աid_In ��Ա��.Id%Type := 0
  ) Is
    Select b.Id
    From �ٴ����ﰲ�� B, �ٴ������Դ C
    Where b.��Դid = c.Id And b.����id = ����id_In And
          (Nvl(��Աid_In, 0) = 0 Or (Nvl(��Աid_In, 0) <> 0 And Nvl(c.�Ƿ��ٴ��Ű�, 0) = 1 And Exists
           (Select 1 From ������Ա Where ����id = c.����id And ��Աid = ��Աid_In))) And Exists
     (Select 1 From ���ű� Where ID = c.����id And (վ��_In Is Null Or (վ�� Is Null Or վ�� = վ��_In)));
Begin
  Select Count(1)
  Into n_Count
  From �ٴ������ A
  Where a.Id = ����id_In And a.������ Is Not Null And a.�Ű෽ʽ <> 3 And Rownum < 2;
  If n_Count <> 0 Then
    v_Err_Msg := '��ǰ������ѷ������������޸ģ�';
    Raise Err_Item;
  End If;

  Select Count(1) Into n_Count From �ٴ������ A Where a.Id = ����id_In And a.�Ű෽ʽ In (1, 2) And Rownum < 2;
  If n_Count <> 0 Then
    n_�����¼ := 1;
  End If;

  Open c_����(����id_In, ��Աid_In);
  Fetch c_���� Bulk Collect
    Into l_����id;
  Close c_����;

  If Nvl(n_�����¼, 0) = 0 Then
    --�ٴ��������ƻ�ģ��
    Forall I In 1 .. l_����id.Count
      Update �ٴ���������
      Set �Ƿ���ſ��� = ��ſ���_In
      Where (�޺��� Is Not Null Or ��Լ�� Is Not Null) And ����id = l_����id(I);
  
    If Nvl(��ſ���_In, 0) = 0 Then
      --ȡ����ſ��ƣ�ɾ���������
      Select /*+cardinality(b,10)*/
       ID Bulk Collect
      Into l_��¼id
      From �ٴ��������� A, Table(l_����id) B
      Where a.����id = b.Column_Value And (a.�޺��� Is Not Null Or a.��Լ�� Is Not Null) And Nvl(a.�Ƿ���ſ���, 0) = 0 And
            Nvl(a.�Ƿ��ʱ��, 0) = 0;
    
      Forall I In 1 .. l_��¼id.Count
        Delete From �ٴ�����ʱ�� Where ����id = l_��¼id(I);
    Else
      --����ʱ�ε���ſ��ƺ����������,��ʼʱ�䡢��ֹʱ����дʱ��εĿ�ʼʱ��ͽ���ʱ��
      For c_���� In (Select /*+cardinality(d,10)*/
                    a.Id, b.����, c.վ��
                   From �ٴ����ﰲ�� A, �ٴ������Դ B, ���ű� C, Table(l_����id) D
                   Where a.��Դid = b.Id And b.����id = c.Id And a.Id = d.Column_Value) Loop
      
        For c_��¼ In (With c_ʱ��� As
                        (Select ʱ���, ��ʼʱ��, ��ֹʱ��
                        From (Select ʱ���,
                                      To_Date('3000-01-01' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                                      To_Date('3000-01-01' || To_Char(��ֹʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') As ��ֹʱ��,
                                      Row_Number() Over(Partition By ʱ��� Order By ʱ���, վ�� Asc, ���� Asc) As ���
                               From ʱ���
                               Where Nvl(վ��, c_����.վ��) = c_����.վ�� And Nvl(����, c_����.����) = c_����.����)
                        Where ��� = 1)
                       Select a.Id, a.�޺���,
                              To_Date(To_Char(Sysdate, 'yyyy-mm-dd ') || To_Char(b.��ʼʱ��, 'hh24:mi:ss'),
                                       'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                              To_Date(To_Char(Sysdate, 'yyyy-mm-dd ') || To_Char(b.��ֹʱ��, 'hh24:mi:ss'),
                                      'yyyy-mm-dd hh24:mi:ss') + Case
                                When b.��ֹʱ�� <= b.��ʼʱ�� Then
                                 1
                                Else
                                 0
                              End As ��ֹʱ��
                       From �ٴ��������� A, c_ʱ��� B
                       Where a.�ϰ�ʱ�� = b.ʱ��� And ����id = c_����.Id And Nvl(�޺���, 0) <> 0 And Nvl(�Ƿ���ſ���, 0) = 1 And
                             Nvl(�Ƿ��ʱ��, 0) = 0 And Not Exists (Select 1 From �ٴ�����ʱ�� Where ����id = a.Id)) Loop
        
          For I In 1 .. c_��¼.�޺��� Loop
            Insert Into �ٴ�����ʱ��
              (����id, ���, ��ʼʱ��, ��ֹʱ��, ��������, �Ƿ�ԤԼ)
            Values
              (c_��¼.Id, I, c_��¼.��ʼʱ��, c_��¼.��ֹʱ��, 1, 1);
          End Loop;
        End Loop;
      End Loop;
    End If;
  Else
    --�ٴ������¼
    Forall I In 1 .. l_����id.Count
      Update �ٴ������¼
      Set �Ƿ���ſ��� = ��ſ���_In
      Where (�޺��� Is Not Null Or ��Լ�� Is Not Null) And ����id = l_����id(I);
  
    If Nvl(��ſ���_In, 0) = 0 Then
      --ȡ����ſ��ƣ�ɾ���������
      Select /*+cardinality(b,10)*/
       a.Id Bulk Collect
      Into l_��¼id
      From �ٴ������¼ A, Table(l_����id) B
      Where a.����id = b.Column_Value And Nvl(a.�޺���, 0) <> 0 And Nvl(a.�Ƿ���ſ���, 0) = 0 And Nvl(a.�Ƿ��ʱ��, 0) = 0;
    
      Forall I In 1 .. l_��¼id.Count
        Delete From �ٴ�������ſ��� Where ��¼id = l_��¼id(I);
    Else
      --����ʱ�ε���ſ��ƺ����������,��ʼʱ�䡢��ֹʱ����дʱ��εĿ�ʼʱ��ͽ���ʱ��
      For c_��¼ In (Select /*+cardinality(b,10)*/
                    a.Id, a.�޺���, a.��ʼʱ��, a.��ֹʱ��
                   From �ٴ������¼ A, Table(l_����id) B
                   Where a.����id = b.Column_Value And Nvl(a.�޺���, 0) <> 0 And Nvl(a.�Ƿ���ſ���, 0) = 1 And Nvl(a.�Ƿ��ʱ��, 0) = 0 And
                         Not Exists (Select 1 From �ٴ�������ſ��� Where ��¼id = a.Id)) Loop
      
        For I In 1 .. c_��¼.�޺��� Loop
          Insert Into �ٴ�������ſ���
            (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ)
          Values
            (c_��¼.Id, I, c_��¼.��ʼʱ��, c_��¼.��ֹʱ��, 1, 1);
        End Loop;
      End Loop;
    End If;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ����ﰲ��_��ſ���;
/

--107321:Ƚ����,2017-04-14,�޸��ޡ����п��ҡ�Ȩ��ʱ�Ĺ��ܿ�������
Create Or Replace Procedure Zl_�ٴ����ﰲ��_Batchdelete
(
  ����id_In   �ٴ������.Id%Type,
  ��Աid_In   ��Ա��.Id%Type := 0,
  վ��_In     ���ű�.վ��%Type := Null,
  ��Դid_In   �ٴ����ﰲ��.��Դid%Type := 0,
  ����id_In   �ٴ����ﰲ��.Id%Type := 0,
  ��ʱ����_In �ٴ����ﰲ��.�Ƿ���ʱ����%Type := 0
) As
  --���ܣ�����ɾ���ٴ����ﰲ�� 
  --������ 
  --      ��Աid_In ������0��ɾ����Ա���ڿ��ҵ����к�Դ���� 
  --      ��Դid_In ������0��ɾ���ú�Դ�����а��� 
  --      ����ID_in ������0��ɾ���ú�Դ�ĵ�ǰ����(һ������ʱ����) 
  --˵���������Աid_In=0�Һ�Դid_In=0 ��ɾ���ó��������к�Դ�����а��� 
  n_Count    Number(8);
  n_�����¼ Number(1);

  Err_Item Exception;
  v_Err_Msg Varchar2(255);

  l_����id t_Numlist := t_Numlist();
  l_��¼id t_Numlist := t_Numlist();
Begin
  If Nvl(��ʱ����_In, 0) = 0 Then
    Begin
      Select 1
      Into n_Count
      From �ٴ������ A
      Where a.Id = ����id_In And a.������ Is Not Null And a.�Ű෽ʽ <> 3 And Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count <> 0 Then
      v_Err_Msg := '��ǰ������ѷ������������޸İ��ţ�';
      Raise Err_Item;
    End If;
  End If;

  Begin
    Select 1 Into n_�����¼ From �ٴ������ A Where a.Id = ����id_In And a.�Ű෽ʽ In (1, 2) And Rownum < 2;
  Exception
    When Others Then
      n_�����¼ := 0;
  End;

  If Nvl(n_�����¼, 0) = 0 Then
    --ɾ���ٴ��������/ģ�� 
    Select a.Id Bulk Collect
    Into l_����id
    From �ٴ��������� A, �ٴ����ﰲ�� B, �ٴ������Դ C, ���ű� D
    Where a.����id = b.Id And b.��Դid = c.Id And c.����id = d.Id And b.����id = ����id_In And
          (
          --ɾ���ó��������к�Դ�����а��� 
           (Nvl(��Դid_In, 0) = 0 And Nvl(��Աid_In, 0) = 0)
          --ɾ���ú�Դ�����а��� 
           Or (Nvl(��Դid_In, 0) <> 0 And b.��Դid = ��Դid_In And Nvl(����id_In, 0) = 0)
          --ɾ���ú�Դ��ѡ���� 
           Or (Nvl(����id_In, 0) <> 0 And b.Id = ����id_In)
          --ɾ����Ա���ڿ��ҵ����к�Դ���� 
           Or (Nvl(��Աid_In, 0) <> 0 And Nvl(c.�Ƿ��ٴ��Ű�, 0) = 1 And Exists
            (Select 1 From ������Ա Where ����id = c.����id And ��Աid = ��Աid_In)))
         --վ�� 
          And (d.վ�� Is Null Or d.վ�� = վ��_In);
  
    Forall I In 1 .. l_����id.Count
      Delete From �ٴ�����Һſ��� Where ����id = l_����id(I);
  
    Forall I In 1 .. l_����id.Count
      Delete From �ٴ�����ʱ�� Where ����id = l_����id(I);
  
    Forall I In 1 .. l_����id.Count
      Delete From �ٴ��������� Where ����id = l_����id(I);
  
    Forall I In 1 .. l_����id.Count
      Delete From �ٴ��������� Where ID = l_����id(I);
  
    --ɾ���ٴ����ﰲ�� 
    For c_���� In (Select b.Id
                 From �ٴ����ﰲ�� B, �ٴ������Դ C, ���ű� D
                 Where b.��Դid = c.Id And c.����id = d.Id And b.����id = ����id_In And
                       (
                       --ɾ���ó��������к�Դ�����а��� 
                        (Nvl(��Դid_In, 0) = 0 And Nvl(��Աid_In, 0) = 0)
                       --ɾ���ú�Դ�����а��� 
                        Or (Nvl(��Դid_In, 0) <> 0 And b.��Դid = ��Դid_In And Nvl(����id_In, 0) = 0)
                       --ɾ���ú�Դ��ѡ���� 
                        Or (Nvl(����id_In, 0) <> 0 And b.Id = ����id_In)
                       --ɾ����Ա���ڿ��ҵ����к�Դ���� 
                        Or (Nvl(��Աid_In, 0) <> 0 And Nvl(c.�Ƿ��ٴ��Ű�, 0) = 1 And Exists
                         (Select 1 From ������Ա Where ����id = c.����id And ��Աid = ��Աid_In)))
                      --վ�� 
                       And (d.վ�� Is Null Or d.վ�� = վ��_In) And Not Exists
                  (Select 1 From �ٴ��������� Where ����id = b.Id)) Loop
      Zl_�ٴ����ﰲ��_Delete(c_����.Id);
    End Loop;
  Else
    --ɾ���ٴ������¼ 
    Select a.Id Bulk Collect
    Into l_��¼id
    From �ٴ������¼ A, �ٴ����ﰲ�� B, �ٴ������Դ C, ���ű� D
    Where a.����id = b.Id And b.��Դid = c.Id And c.����id = d.Id And b.����id = ����id_In And
          (
          --ɾ���ó��������к�Դ�����а��� 
           (Nvl(��Դid_In, 0) = 0 And Nvl(��Աid_In, 0) = 0)
          --ɾ���ú�Դ�����а��� 
           Or (Nvl(��Դid_In, 0) <> 0 And b.��Դid = ��Դid_In And Nvl(����id_In, 0) = 0)
          --ɾ���ú�Դ��ѡ���� 
           Or (Nvl(����id_In, 0) <> 0 And b.Id = ����id_In)
          --ɾ����Ա���ڿ��ҵ����к�Դ���� 
           Or (Nvl(��Աid_In, 0) <> 0 And Nvl(c.�Ƿ��ٴ��Ű�, 0) = 1 And Exists
            (Select 1 From ������Ա Where ����id = c.����id And ��Աid = ��Աid_In)))
         --վ�� 
          And (d.վ�� Is Null Or d.վ�� = վ��_In);
  
    Zl_�ٴ������¼_Batchdelete(l_��¼id);
  
    --ɾ���ٴ����ﰲ�� 
    For c_���� In (Select b.Id
                 From �ٴ����ﰲ�� B, �ٴ������Դ C, ���ű� D
                 Where b.��Դid = c.Id And c.����id = d.Id And b.����id = ����id_In And
                       (
                       --ɾ���ó��������к�Դ�����а��� 
                        (Nvl(��Դid_In, 0) = 0 And Nvl(��Աid_In, 0) = 0)
                       --ɾ���ú�Դ�����а��� 
                        Or (Nvl(��Դid_In, 0) <> 0 And b.��Դid = ��Դid_In And Nvl(����id_In, 0) = 0)
                       --ɾ���ú�Դ��ѡ���� 
                        Or (Nvl(����id_In, 0) <> 0 And b.Id = ����id_In)
                       --ɾ����Ա���ڿ��ҵ����к�Դ���� 
                        Or (Nvl(��Աid_In, 0) <> 0 And Nvl(c.�Ƿ��ٴ��Ű�, 0) = 1 And Exists
                         (Select 1 From ������Ա Where ����id = c.����id And ��Աid = ��Աid_In)))
                      --վ�� 
                       And (d.վ�� Is Null Or d.վ�� = վ��_In) And Not Exists
                  (Select 1 From �ٴ������¼ Where ����id = b.Id)) Loop
      Zl_�ٴ����ﰲ��_Delete(c_����.Id);
    End Loop;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ����ﰲ��_Batchdelete;
/

--108192:��͢��,2017-04-14,���icd����ɾ�����������
Create Or Replace Procedure Zl_������ϼ�¼_Delete
(
  --���ܣ�ɾ��������ϼ�¼ 
  --�������������_IN=Ϊ��ʱ��ʾ��������,����Ϊ�ַ���,��'1,2,3...' 
  --      ���s_In=��Ҫɾ�������ID�� ,��ʽΪ 'ID1,ID2,ID3...'  
  ����id_In   ������ϼ�¼.����id%Type,
  ��ҳid_In   ������ϼ�¼.��ҳid%Type,
  ��¼��Դ_In ������ϼ�¼.��¼��Դ%Type := Null,
  ����id_In   ������ϼ�¼.����id%Type := Null,
  �������_In Varchar2 := Null,
  ���ids_In  Varchar2 := Null
) Is
  V_���ʹ� Varchar2(255);
  V_����   ������ϼ�¼.�������%Type;
Begin
  If �������_In Is Null Then
    If Not ���ids_In Is Null Then
      For Rdiag In (Select /*+ Rule*/
                     ID, ��¼��Դ, �������, ��ϴ���
                    From ������ϼ�¼
                    Where ID In (Select Column_Value From Table(F_Str2list(���ids_In)))
                    Order By ��¼��Դ, �������, ��ϴ���) Loop
        If Rdiag.��¼��Դ = 3 And Rdiag.������� = 2 And Rdiag.��ϴ��� = 1 Then
          Update ������ҳ Set ������ = Null Where ����id = ����id_In And ��ҳid = Nvl(��ҳid_In, 0);
        End If;
        --�����������ǵ�ǰ�����������ͣ���ɾ�����
        If Rdiag.��¼��Դ = ��¼��Դ_In Or ��¼��Դ_In Is Null Then
          Delete From �������ҽ�� Where ���id = Rdiag.Id;
          Delete From ������ϼ�¼
          Where ����id = ����id_In And ��ҳid = ��ҳid_In And ��¼��Դ = Rdiag.��¼��Դ And ������� = Rdiag.������� And ��ϴ��� = Rdiag.��ϴ��� And
                Nvl(�������, 1) = 2;
          Delete From ������ϼ�¼ Where ID = Rdiag.Id;
        End If;
      End Loop;
    Else
      Delete From �������ҽ��
      Where ���id In (Select ID
                     From ������ϼ�¼
                     Where ����id = ����id_In And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0) And (��¼��Դ = ��¼��Դ_In Or ��¼��Դ_In Is Null) And
                           (����id = ����id_In Or ����id_In Is Null));
    
      Delete From ������ϼ�¼
      Where ����id = ����id_In And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0) And (��¼��Դ = ��¼��Դ_In Or ��¼��Դ_In Is Null) And
            (����id = ����id_In Or ����id_In Is Null);
      --ɾ�������ֱ�ʶ 
      If ��¼��Դ_In = 3 Then
        Update ������ҳ Set ������ = Null Where ����id = ����id_In And ��ҳid = Nvl(��ҳid_In, 0);
      End If;
    End If;
  Else
    V_���ʹ� := �������_In || ',';
    While V_���ʹ� Is Not Null Loop
      V_���� := To_Number(Substr(V_���ʹ�, 1, Instr(V_���ʹ�, ',') - 1));
    
      Delete From �������ҽ��
      Where ���id In (Select ID
                     From ������ϼ�¼
                     Where ����id = ����id_In And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0) And (��¼��Դ = ��¼��Դ_In Or ��¼��Դ_In Is Null) And
                           (����id = ����id_In Or ����id_In Is Null) And ������� = V_����);
    
      Delete From ������ϼ�¼
      Where ����id = ����id_In And Nvl(��ҳid, 0) = Nvl(��ҳid_In, 0) And (��¼��Դ = ��¼��Դ_In Or ��¼��Դ_In Is Null) And
            (����id = ����id_In Or ����id_In Is Null) And ������� = V_����;
    
      V_���ʹ� := Substr(V_���ʹ�, Instr(V_���ʹ�, ',') + 1);
    
      --�������Ժ�����ɾ�������ֱ�ʶ 
      If V_���� = 2 And ��¼��Դ_In = 3 Then
        Update ������ҳ Set ������ = Null Where ����id = ����id_In And ��ҳid = Nvl(��ҳid_In, 0);
      End If;
    End Loop;
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_������ϼ�¼_Delete;
/

--106872:������,2017-04-12,ԤԼ���ձ���ժҪ
Create Or Replace Procedure Zl_ԤԼ�ҺŽ���_Insert
(
  No_In            ������ü�¼.No%Type,
  Ʊ�ݺ�_In        ������ü�¼.ʵ��Ʊ��%Type,
  ����id_In        Ʊ��ʹ����ϸ.����id%Type,
  ����id_In        ������ü�¼.����id%Type,
  ����_In          ������ü�¼.��ҩ����%Type,
  ����id_In        ������ü�¼.����id%Type,
  �����_In        ������ü�¼.��ʶ��%Type,
  ����_In          ������ü�¼.����%Type,
  �Ա�_In          ������ü�¼.�Ա�%Type,
  ����_In          ������ü�¼.����%Type,
  ���ʽ_In      ������ü�¼.���ʽ%Type, --���ڴ�Ų��˵�ҽ�Ƹ��ʽ���
  �ѱ�_In          ������ü�¼.�ѱ�%Type,
  ���㷽ʽ_In      ����Ԥ����¼.���㷽ʽ%Type, --�ֽ�Ľ�������
  �ֽ�֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�ֽ�֧�����ݽ��
  Ԥ��֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱʹ�õ�Ԥ�����
  ����֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�����ʻ�֧�����
  ����ʱ��_In      ������ü�¼.����ʱ��%Type,
  ����_In          �Һ����״̬.���%Type,
  ����Ա���_In    ������ü�¼.����Ա���%Type,
  ����Ա����_In    ������ü�¼.����Ա����%Type,
  ���ɶ���_In      Number := 0,
  �Ǽ�ʱ��_In      ������ü�¼.�Ǽ�ʱ��%Type := Null,
  �����id_In      ����Ԥ����¼.�����id%Type := Null,
  ���㿨���_In    ����Ԥ����¼.���㿨���%Type := Null,
  ����_In          ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In    ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In      ����Ԥ����¼.����˵��%Type := Null,
  ����_In          ���˹Һż�¼.����%Type := Null,
  ����ģʽ_In      Number := 0,
  ���ʷ���_In      Number := 0,
  ��Ԥ������ids_In Varchar2 := Null,
  ��������_In      Number := 0,
  ���½������_In  Number := 0, --�Ƿ������Ա��������Ҫ�Ǵ���ͳһ����Ա��¼��̨�����������
  ժҪ_In          ���˹Һż�¼.ժҪ%Type := Null
) As
  --���α������շѳ�Ԥ���Ŀ���Ԥ���б�
  --��ID�������ȳ��ϴ�δ����ġ�
  Cursor c_Deposit
  (
    v_����id        ������Ϣ.����id%Type,
    v_��Ԥ������ids Varchar2
  ) Is
    Select ����id, NO, Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) As ���, Min(��¼״̬) As ��¼״̬, Nvl(Max(����id), 0) As ����id,
           Max(Decode(��¼����, 1, ID, 0) * Decode(��¼״̬, 2, 0, 1)) As ԭԤ��id
    From ����Ԥ����¼
    Where ��¼���� In (1, 11) And ����id In (Select Column_Value From Table(f_Num2list(v_��Ԥ������ids))) And Nvl(Ԥ�����, 2) = 1
     Having Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0
    Group By NO, ����id
    Order By Decode(����id, Nvl(v_����id, 0), 0, 1), ����id, NO;

  v_Err_Msg Varchar2(255);
  Err_Item    Exception;
  Err_Special Exception;

  v_�ֽ�     ���㷽ʽ.����%Type;
  v_�����ʻ� ���㷽ʽ.����%Type;
  v_�������� �ŶӽкŶ���.��������%Type;
  v_�ű�     ������ü�¼.���㵥λ%Type;
  v_����     ������ü�¼.��ҩ����%Type;
  v_�ŶӺ��� �ŶӽкŶ���.�ŶӺ��� %Type;
  v_ԤԼ��ʽ ���˹Һż�¼.ԤԼ��ʽ %Type;

  n_��ӡid        Ʊ�ݴ�ӡ����.Id%Type;
  n_Ԥ�����      ����Ԥ����¼.���%Type;
  n_����ֵ        ����Ԥ����¼.���%Type;
  v_��Ԥ������ids Varchar2(4000);

  n_�Һ�id         ���˹Һż�¼.Id%Type;
  n_����̨ǩ���Ŷ� Number;
  n_��id           ����ɿ����.Id%Type;
  n_Count          Number(18);
  n_�Ŷ�           Number;
  n_�����Ŷ�       Number;
  n_��ǰ���       ����Ԥ����¼.���%Type;
  n_Ԥ��id         ����Ԥ����¼.Id%Type;
  n_���ѿ�id       ���ѿ�Ŀ¼.Id%Type;
  n_���ƿ�         Number;

  d_Date       Date;
  d_ԤԼʱ��   ������ü�¼.����ʱ��%Type;
  d_����ʱ��   Date;
  d_�Ŷ�ʱ��   Date;
  n_ʱ��       Number := 0;
  n_����       Number := 0;
  v_�Ŷ����   �ŶӽкŶ���.�Ŷ����%Type;
  n_����ģʽ   ������Ϣ.����ģʽ%Type;
  n_Ʊ��       Ʊ��ʹ����ϸ.Ʊ��%Type;
  v_���ʽ   ���˹Һż�¼.ҽ�Ƹ��ʽ%Type;
  v_����Ա���� ���˹Һż�¼.������%Type;
  n_����ģʽ   Number := 0;
Begin
  n_��id          := Zl_Get��id(����Ա����_In);
  v_��Ԥ������ids := Nvl(��Ԥ������ids_In, ����id_In);
  n_����ģʽ      := Nvl(zl_GetSysParameter('ԤԼ����ģʽ', 1111), 0);

  --��ȡ���㷽ʽ����
  Begin
    Select ���� Into v_�ֽ� From ���㷽ʽ Where ���� = 1;
  Exception
    When Others Then
      v_�ֽ� := '�ֽ�';
  End;
  Begin
    Select ���� Into v_�����ʻ� From ���㷽ʽ Where ���� = 3;
  Exception
    When Others Then
      v_�����ʻ� := '�����ʻ�';
  End;
  If �Ǽ�ʱ��_In Is Null Then
    Select Sysdate Into d_Date From Dual;
  Else
    d_Date := �Ǽ�ʱ��_In;
  End If;

  --���¹Һ����״̬
  Begin
    Select �ű�, ����, Trunc(����ʱ��), ����ʱ��, ԤԼ��ʽ
    Into v_�ű�, v_����, d_ԤԼʱ��, d_����ʱ��, v_ԤԼ��ʽ
    From ���˹Һż�¼
    Where ��¼���� = 2 And ��¼״̬ = 1 And Rownum = 1 And NO = No_In;
  Exception
    When Others Then
      Select Max(������) Into v_����Ա���� From ���˹Һż�¼ Where ��¼���� = 2 And ��¼״̬ In (1, 3) And NO = No_In;
      If v_����Ա���� Is Null Then
        v_Err_Msg := '��ǰԤԼ�Һŵ��ѱ�ȡ��';
        Raise Err_Item;
      Else
        If v_����Ա���� = ����Ա����_In Then
          v_Err_Msg := '��ǰԤԼ�Һŵ��ѱ�����';
          Raise Err_Special;
        Else
          v_Err_Msg := '��ǰԤԼ�Һŵ��ѱ������˽���';
          Raise Err_Item;
        End If;
      End If;
  End;

  --�ж��Ƿ��ʱ��
  Begin
    Select 1
    Into n_ʱ��
    From Dual
    Where Exists (Select 1
           From �ҺŰ���ʱ�� A, �ҺŰ��� B
           Where a.����id = b.Id And b.���� = v_�ű� And Rownum < 2
           Union All
           Select 1
           From �Һżƻ�ʱ�� C, �ҺŰ��żƻ� D ��
           Where c.�ƻ�id = d.Id And d.���� = v_�ű� And d.��Чʱ�� > Sysdate And Rownum < 2);
  Exception
    When Others Then
      n_ʱ�� := 0;
  End;

  If n_ʱ�� = 0 And ��������_In = 0 Then
    If n_����ģʽ = 0 Then
      If Trunc(����ʱ��_In) = Trunc(Sysdate) Then
        d_����ʱ�� := ����ʱ��_In;
      Else
        d_����ʱ�� := Sysdate;
      End If;
    Else
      d_����ʱ�� := ����ʱ��_In;
    End If;
  Else
    If Not ����ʱ��_In Is Null Then
      d_����ʱ�� := ����ʱ��_In;
    End If;
  End If;
  If Not v_���� Is Null Then
    If ����_In Is Null Then
      Delete �Һ����״̬ Where ���� = v_�ű� And Trunc(����) = Trunc(d_ԤԼʱ��) And ��� = v_����;
    Else
      If Trunc(d_ԤԼʱ��) <> Trunc(Sysdate) And n_����ģʽ = 0 Then
      
        If n_ʱ�� = 0 And ��������_In = 0 Then
          --��ǰ���ջ��ӳٽ���
          Delete �Һ����״̬ Where ���� = v_�ű� And Trunc(����) = Trunc(d_ԤԼʱ��) And ��� = v_����;
          Begin
            Select 1 Into n_���� From �Һ����״̬ Where ���� = v_�ű� And ���� = Trunc(Sysdate) And ��� = v_����;
          Exception
            When Others Then
              n_���� := 0;
          End;
          If n_���� = 0 Then
            Insert Into �Һ����״̬
              (����, ����, ���, ״̬, ����Ա����, �Ǽ�ʱ��)
            Values
              (v_�ű�, Trunc(Sysdate), v_����, 1, ����Ա����_In, Sysdate);
          Else
            --�����ѱ�ʹ�õ����
            Begin
              v_���� := 1;
              Insert Into �Һ����״̬
                (����, ����, ���, ״̬, ����Ա����, �Ǽ�ʱ��)
              Values
                (v_�ű�, Trunc(Sysdate), v_����, 1, ����Ա����_In, Sysdate);
            Exception
              When Others Then
                Select Min(��� + 1)
                Into v_����
                From �Һ����״̬ A
                Where ���� = v_�ű� And ���� = Trunc(Sysdate) And Not Exists
                 (Select 1 From �Һ����״̬ Where ���� = a.���� And ���� = a.���� And ��� = a.��� + 1);
                Insert Into �Һ����״̬
                  (����, ����, ���, ״̬, ����Ա����, �Ǽ�ʱ��)
                Values
                  (v_�ű�, Trunc(Sysdate), v_����, 1, ����Ա����_In, Sysdate);
            End;
          End If;
        Else
          Update �Һ����״̬
          Set ״̬ = 1, �Ǽ�ʱ�� = Sysdate
          Where Trunc(����) = Trunc(d_ԤԼʱ��) And ��� = v_���� And ���� = v_�ű� And ״̬ = 2;
          If Sql% NotFound Then
            Begin
              Insert Into �Һ����״̬
                (����, ����, ���, ״̬, ����Ա����, �Ǽ�ʱ��)
              Values
                (v_�ű�, Trunc(Sysdate), v_����, 1, ����Ա����_In, Sysdate);
            Exception
              When Others Then
                v_Err_Msg := '���' || v_���� || '�ѱ�������ʹ��,������ѡ��һ�����.';
                Raise Err_Item;
            End;
          End If;
        
        End If;
      
      Else
        Update �Һ����״̬
        Set ��� = ����_In, ״̬ = 1, �Ǽ�ʱ�� = Sysdate
        Where ���� = v_�ű� And Trunc(����) = Trunc(d_ԤԼʱ��) And ��� = v_����;
        If Sql%RowCount = 0 Then
          Begin
            Insert Into �Һ����״̬
              (����, ����, ���, ״̬, ����Ա����, �Ǽ�ʱ��)
            Values
              (v_�ű�, Trunc(d_����ʱ��), v_����, 1, ����Ա����_In, Sysdate);
          Exception
            When Others Then
              v_Err_Msg := '���' || v_���� || '�ѱ�������ʹ��,������ѡ��һ�����.';
              Raise Err_Item;
          End;
        End If;
      End If;
    End If;
  Else
    If Not ����_In Is Null Then
      Begin
        Insert Into �Һ����״̬
          (����, ����, ���, ״̬, ����Ա����, �Ǽ�ʱ��)
        Values
          (v_�ű�, Trunc(Sysdate), ����_In, 1, ����Ա����_In, Sysdate);
      Exception
        When Others Then
          v_Err_Msg := '���' || ����_In || '�ѱ�������ʹ��,������ѡ��һ�����.';
          Raise Err_Item;
      End;
      v_���� := ����_In;
    Else
      v_���� := Null;
    End If;
  End If;

  --����������ü�¼
  Update ������ü�¼
  Set ��¼״̬ = 1, ʵ��Ʊ�� = Decode(Nvl(���ʷ���_In, 0), 1, Null, Ʊ�ݺ�_In), ����id = Decode(Nvl(���ʷ���_In, 0), 1, Null, ����id_In),
      ���ʽ�� = Decode(Nvl(���ʷ���_In, 0), 1, Null, ʵ�ս��), ��ҩ���� = ����_In, ����id = ����id_In, ��ʶ�� = �����_In, ���� = ����_In, ���� = ����_In,
      �Ա� = �Ա�_In, ���ʽ = ���ʽ_In, �ѱ� = �ѱ�_In, ����ʱ�� = d_����ʱ��, �Ǽ�ʱ�� = d_Date, ����Ա��� = ����Ա���_In, ����Ա���� = ����Ա����_In,
      �ɿ���id = n_��id, ���ʷ��� = Decode(Nvl(���ʷ���_In, 0), 1, 1, 0), ժҪ = Nvl(ժҪ_In, ժҪ)
  Where ��¼���� = 4 And ��¼״̬ = 0 And NO = No_In;

  --���˹Һż�¼
  Update ���˹Һż�¼
  Set ������ = ����Ա����_In, ����ʱ�� = d_Date, ��¼���� = 1, ����id = ����id_In, ����� = �����_In, ����ʱ�� = d_����ʱ��, ���� = ����_In, �Ա� = �Ա�_In,
      ���� = ����_In, ����Ա��� = ����Ա���_In, ����Ա���� = ����Ա����_In, ���� = Decode(Nvl(����_In, 0), 0, Null, ����_In), ���� = v_����, ���� = ����_In,
      ժҪ = Nvl(ժҪ_In, ժҪ)
  Where ��¼״̬ = 1 And NO = No_In And ��¼���� = 2
  Returning ID Into n_�Һ�id;
  If Sql%NotFound Then
    Begin
      Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
      Begin
        Select ���� Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = ���ʽ_In And Rownum < 2;
      Exception
        When Others Then
          v_���ʽ := Null;
      End;
      Insert Into ���˹Һż�¼
        (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����,
         ժҪ, ����, ԤԼ, ԤԼ��ʽ, ������, ����ʱ��, ԤԼʱ��, ����, ҽ�Ƹ��ʽ)
        Select n_�Һ�id, No_In, 1, 1, ����id_In, �����_In, ����_In, �Ա�_In, ����_In, ���㵥λ, �Ӱ��־, ����_In, Null, ִ�в���id, ִ����, 0, Null,
               �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����, Nvl(ժҪ_In, ժҪ), v_����, 1, Substr(����, 1, 10) As ԤԼ��ʽ, ����Ա����_In,
               Nvl(�Ǽ�ʱ��_In, Sysdate), ����ʱ��, Decode(Nvl(����_In, 0), 0, Null, ����_In), v_���ʽ
        From ������ü�¼
        Where ��¼���� = 4 And ��¼״̬ = 1 And Rownum = 1 And NO = No_In;
    Exception
      When Others Then
        v_Err_Msg := '���ڲ���ԭ��,���ݺ�Ϊ��' || No_In || '���Ĳ���' || ����_In || '�Ѿ�������';
        Raise Err_Item;
    End;
  End If;

  --0-����������;1-��ҽ�������̨�Ŷ�;2-�ȷ���,��ҽ��վ
  If Nvl(���ɶ���_In, 0) <> 0 Then
    n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113));
    If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Then
      For v_�Һ� In (Select ID, ����, ����, ִ����, ִ�в���id, ����ʱ��, �ű�, ���� From ���˹Һż�¼ Where NO = No_In) Loop
      
        Begin
          Select 1,
                 Case
                   When �Ŷ�ʱ�� < Trunc(Sysdate) Then
                    1
                   Else
                    0
                 End
          Into n_�Ŷ�, n_�����Ŷ�
          From �ŶӽкŶ���
          Where ҵ������ = 0 And ҵ��id = v_�Һ�.Id And Rownum <= 1;
        Exception
          When Others Then
            n_�Ŷ� := 0;
        End;
        If n_�Ŷ� = 0 Then
          --��������
          --����ִ�в��š���������
          n_�Һ�id   := v_�Һ�.Id;
          v_�������� := v_�Һ�.ִ�в���id;
          v_�ŶӺ��� := Zlgetnextqueue(v_�Һ�.ִ�в���id, n_�Һ�id, v_�Һ�.�ű� || '|' || v_�Һ�.����);
          v_�Ŷ���� := Zlgetsequencenum(0, n_�Һ�id, 0);
        
          --�Һ�id_In,����_In,����_In,ȱʡ����_In,��չ_In(������)
          d_�Ŷ�ʱ�� := Zl_Get_Queuedate(n_�Һ�id, v_�Һ�.�ű�, v_�Һ�.����, v_�Һ�.����ʱ��);
          --   ��������_In , ҵ������_In, ҵ��id_In,����id_In,�ŶӺ���_In,�Ŷӱ��_In,��������_In,����ID_IN, ����_In, ҽ������_In,
          Zl_�ŶӽкŶ���_Insert(v_��������, 0, n_�Һ�id, v_�Һ�.ִ�в���id, v_�ŶӺ���, Null, ����_In, ����id_In, v_�Һ�.����, v_�Һ�.ִ����, d_�Ŷ�ʱ��,
                           v_ԤԼ��ʽ, Null, v_�Ŷ����);
        Elsif Nvl(n_�����Ŷ�, 0) = 1 Then
          --���¶��к�
          v_�ŶӺ��� := Zlgetnextqueue(v_�Һ�.ִ�в���id, v_�Һ�.Id, v_�Һ�.�ű� || '|' || Nvl(v_�Һ�.����, 0));
          v_�Ŷ���� := Zlgetsequencenum(0, v_�Һ�.Id, 1);
          --�¶�������_IN, ҵ������_In, ҵ��id_In , ����id_In , ��������_In , ����_In, ҽ������_In ,�ŶӺ���_In
          Zl_�ŶӽкŶ���_Update(v_�Һ�.ִ�в���id, 0, v_�Һ�.Id, v_�Һ�.ִ�в���id, v_�Һ�.����, v_�Һ�.����, v_�Һ�.ִ����, v_�ŶӺ���, v_�Ŷ����);
        
        Else
          --�¶�������_IN, ҵ������_In, ҵ��id_In , ����id_In , ��������_In , ����_In, ҽ������_In ,�ŶӺ���_In
          Zl_�ŶӽкŶ���_Update(v_�Һ�.ִ�в���id, 0, v_�Һ�.Id, v_�Һ�.ִ�в���id, v_�Һ�.����, v_�Һ�.����, v_�Һ�.ִ����);
        End If;
        --ԤԼ����ʱ���ı��¼��־
        Update ���˹Һż�¼ Set ��¼��־ = 1 Where ID = n_�Һ�id;
      End Loop;
    End If;
  End If;

  --���ܽ��㵽����Ԥ����¼
  If (Nvl(�ֽ�֧��_In, 0) <> 0 Or (Nvl(�ֽ�֧��_In, 0) = 0 And Nvl(����֧��_In, 0) = 0 And Nvl(Ԥ��֧��_In, 0) = 0)) And
     Nvl(���ʷ���_In, 0) = 0 Then
    Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
    Insert Into ����Ԥ����¼
      (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, �������,
       ��������)
    Values
      (n_Ԥ��id, 4, 1, No_In, ����id_In, Nvl(���㷽ʽ_In, v_�ֽ�), Nvl(�ֽ�֧��_In, 0), d_Date, ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�',
       n_��id, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, ����id_In, 4);
  
    If Nvl(���㿨���_In, 0) <> 0 And Nvl(�ֽ�֧��_In, 0) <> 0 Then
    
      n_���ѿ�id := Null;
      Begin
        Select Nvl(���ƿ�, 0), 1 Into n_���ƿ�, n_Count From �����ѽӿ�Ŀ¼ Where ��� = ���㿨���_In;
      Exception
        When Others Then
          n_Count := 0;
      End;
      If n_Count = 0 Then
        v_Err_Msg := '[ZLSOFT]û�з���ԭ���㿨����Ӧ���,���ܼ���������[ZLSOFT]';
        Raise Err_Item;
      End If;
      If n_���ƿ� = 1 Then
        Select ID
        Into n_���ѿ�id
        From ���ѿ�Ŀ¼
        Where �ӿڱ�� = ���㿨���_In And ���� = ����_In And
              ��� = (Select Max(���) From ���ѿ�Ŀ¼ Where �ӿڱ�� = ���㿨���_In And ���� = ����_In);
      End If;
      Zl_���˿������¼_Insert(���㿨���_In, n_���ѿ�id, ���㷽ʽ_In, �ֽ�֧��_In, ����_In, Null, �Ǽ�ʱ��_In, Null, ����id_In, n_Ԥ��id);
    End If;
  
  End If;

  --���ھ��￨ͨ��Ԥ����Һ�
  If Nvl(Ԥ��֧��_In, 0) <> 0 And Nvl(���ʷ���_In, 0) = 0 Then
    n_Ԥ����� := Ԥ��֧��_In;
    For r_Deposit In c_Deposit(����id_In, v_��Ԥ������ids) Loop
      n_��ǰ��� := Case
                  When r_Deposit.��� - n_Ԥ����� < 0 Then
                   r_Deposit.���
                  Else
                   n_Ԥ�����
                End;
      If r_Deposit.����id = 0 Then
        --��һ�γ�Ԥ��(���Ͻ���ID,���Ϊ0)
        Update ����Ԥ����¼ Set ��Ԥ�� = 0, ����id = ����id_In, �������� = 4 Where ID = r_Deposit.ԭԤ��id;
      End If;
      --���ϴ�ʣ���
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
         ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, Ԥ�����, �������, ��������)
        Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �Ǽ�ʱ��_In,
               ����Ա����_In, ����Ա���_In, n_��ǰ���, ����id_In, n_��id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, Ԥ�����, ����id_In, 4
        From ����Ԥ����¼
        Where NO = r_Deposit.No And ��¼״̬ = r_Deposit.��¼״̬ And ��¼���� In (1, 11) And Rownum = 1;
    
      --���²���Ԥ�����
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) - n_��ǰ���
      Where ����id = r_Deposit.����id And ���� = 1 And ���� = Nvl(1, 2)
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ������� (����id, ����, Ԥ�����, ����) Values (r_Deposit.����id, Nvl(1, 2), -1 * n_��ǰ���, 1);
        n_����ֵ := -1 * n_��ǰ���;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From �������
        Where ����id = r_Deposit.����id And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
      End If;
    
      --����Ƿ��Ѿ�������
      If r_Deposit.��� < n_Ԥ����� Then
        n_Ԥ����� := n_Ԥ����� - r_Deposit.���;
      Else
        n_Ԥ����� := 0;
      End If;
    
      If n_Ԥ����� = 0 Then
        Exit;
      End If;
    End Loop;
  End If;

  --����ҽ���Һ�
  If Nvl(����֧��_In, 0) <> 0 And Nvl(���ʷ���_In, 0) = 0 Then
    Insert Into ����Ԥ����¼
      (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ,
       Ԥ�����, �������, ��������)
    Values
      (����Ԥ����¼_Id.Nextval, 4, 1, No_In, ����id_In, v_�����ʻ�, ����֧��_In, d_Date, ����Ա���_In, ����Ա����_In, ����id_In, 'ҽ���Һ�', n_��id,
       Null, Null, Null, Null, Null, Null, Null, ����id_In, 4);
  End If;

  --��ػ��ܱ�Ĵ���
  --��Ա�ɿ����
  If Nvl(�ֽ�֧��_In, 0) <> 0 And Nvl(���ʷ���_In, 0) = 0 And Nvl(���½������_In, 0) = 0 Then
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) + �ֽ�֧��_In
    Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = Nvl(���㷽ʽ_In, v_�ֽ�)
    Returning ��� Into n_����ֵ;
  
    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ����
        (�տ�Ա, ���㷽ʽ, ����, ���)
      Values
        (����Ա����_In, Nvl(���㷽ʽ_In, v_�ֽ�), 1, �ֽ�֧��_In);
      n_����ֵ := �ֽ�֧��_In;
    
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ��Ա�ɿ����
      Where �տ�Ա = ����Ա����_In And ���� = 1 And ���㷽ʽ = Nvl(���㷽ʽ_In, v_�ֽ�) And Nvl(���, 0) = 0;
    End If;
  End If;

  If Nvl(����֧��_In, 0) <> 0 And Nvl(���ʷ���_In, 0) = 0 And Nvl(���½������_In, 0) = 0 Then
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) + ����֧��_In
    Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = v_�����ʻ�
    Returning ��� Into n_����ֵ;
  
    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ���� (�տ�Ա, ���㷽ʽ, ����, ���) Values (����Ա����_In, v_�����ʻ�, 1, ����֧��_In);
      n_����ֵ := ����֧��_In;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ��Ա�ɿ���� Where �տ�Ա = ����Ա����_In And ���� = 1 And Nvl(���, 0) = 0;
    End If;
  End If;

  --����Ʊ��ʹ�����
  If Ʊ�ݺ�_In Is Not Null And Nvl(���ʷ���_In, 0) = 0 Then
    Select Ʊ�ݴ�ӡ����_Id.Nextval Into n_��ӡid From Dual;
  
    --��ǰƱ�ݵ�Ʊ��
    Select Ʊ�� Into n_Ʊ�� From Ʊ�����ü�¼ Where ID = Nvl(����id_In, 0);
    --����Ʊ��
    Insert Into Ʊ�ݴ�ӡ���� (ID, ��������, NO) Values (n_��ӡid, 4, No_In);
  
    Insert Into Ʊ��ʹ����ϸ
      (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����)
    Values
      (Ʊ��ʹ����ϸ_Id.Nextval, n_Ʊ��, Ʊ�ݺ�_In, 1, 1, ����id_In, n_��ӡid, d_Date, ����Ա����_In);
  
    --״̬�Ķ�
    Update Ʊ�����ü�¼
    Set ��ǰ���� = Ʊ�ݺ�_In, ʣ������ = Decode(Sign(ʣ������ - 1), -1, 0, ʣ������ - 1), ʹ��ʱ�� = d_Date
    Where ID = Nvl(����id_In, 0);
  End If;

  If Nvl(���ʷ���_In, 0) = 1 Then
    --����
    If Nvl(����id_In, 0) = 0 Then
      v_Err_Msg := 'Ҫ��Բ��˵ĹҺŷѽ��м��ʣ������ǽ������˲��ܼ��ʹҺš�';
      Raise Err_Item;
    End If;
    For c_���� In (Select ʵ�ս��, ���˿���id, ��������id, ִ�в���id, ������Ŀid
                 From ������ü�¼
                 Where ��¼���� = 4 And ��¼״̬ = 1 And NO = No_In And Nvl(���ʷ���, 0) = 1) Loop
      --�������
      Update �������
      Set ������� = Nvl(�������, 0) + Nvl(c_����.ʵ�ս��, 0)
      Where ����id = Nvl(����id_In, 0) And ���� = 1 And ���� = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, ����, ����, �������, Ԥ�����)
        Values
          (����id_In, 1, 1, Nvl(c_����.ʵ�ս��, 0), 0);
      End If;
    
      --����δ�����
      Update ����δ�����
      Set ��� = Nvl(���, 0) + Nvl(c_����.ʵ�ս��, 0)
      Where ����id = ����id_In And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���˿���id, 0) = Nvl(c_����.���˿���id, 0) And
            Nvl(��������id, 0) = Nvl(c_����.��������id, 0) And Nvl(ִ�в���id, 0) = Nvl(c_����.ִ�в���id, 0) And ������Ŀid + 0 = c_����.������Ŀid And
            ��Դ;�� + 0 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into ����δ�����
          (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
        Values
          (����id_In, Null, Null, c_����.���˿���id, c_����.��������id, c_����.ִ�в���id, c_����.������Ŀid, 1, Nvl(c_����.ʵ�ս��, 0));
      End If;
    End Loop;
  End If;
  If Nvl(����id_In, 0) <> 0 Then
    n_����ģʽ := 0;
    Update ������Ϣ
    Set ����ʱ�� = d_����ʱ��, ����״̬ = 1, �������� = ����_In
    Where ����id = ����id_In
    Returning Nvl(����ģʽ, 0) Into n_����ģʽ;
    --ȡ����:
    If Nvl(n_����ģʽ, 0) <> Nvl(����ģʽ_In, 0) Then
      --����ģʽ��ȷ��
      If n_����ģʽ = 1 And Nvl(����ģʽ_In, 0) = 0 Then
        --�����Ѿ���"�����ƺ�����",������"�Ƚ�������Ƶ�",�����Ƿ����δ������
        Select Count(1)
        Into n_Count
        From ����δ�����
        Where ����id = ����id_In And (��Դ;�� In (1, 4) Or ��Դ;�� = 3 And Nvl(��ҳid, 0) = 0) And Nvl(���, 0) <> 0 And Rownum < 2;
        If Nvl(n_Count, 0) <> 0 Then
          --����δ�������ݣ������Ƚ���������ִ��
          v_Err_Msg := '��ǰ���˵ľ���ģʽΪ�����ƺ�����Ҵ���δ����ã�����������ò��˵ľ���ģʽ,������ȶ�δ����ý��ʺ��ٹҺŻ򲻵������˵ľ���ģʽ!';
          Raise Err_Item;
        End If;
        --���
        --δ����ҽ��ҵ��ģ�����ʱ�͹Һŵ�,��Ҫ��֤ͬһ�εľ���ģʽ��һ����(�����Ѿ���飬�����ٴ���)
      End If;
      Update ������Ϣ Set ����ģʽ = ����ģʽ_In Where ����id = ����id_In;
    End If;
  End If;

  --���˵�����Ϣ
  If ����id_In Is Not Null Then
    Update ������Ϣ
    Set ������ = Null, ������ = Null, �������� = Null
    Where ����id = ����id_In And Nvl(��Ժ, 0) = 0 And Exists
     (Select 1
           From ���˵�����¼
           Where ����id = ����id_In And ��ҳid Is Not Null And
                 �Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From ���˵�����¼ Where ����id = ����id_In));
    If Sql%RowCount > 0 Then
      Update ���˵�����¼
      Set ����ʱ�� = d_Date
      Where ����id = ����id_In And ��ҳid Is Not Null And Nvl(����ʱ��, d_Date) >= d_Date;
    End If;
  End If;
  --��Ϣ����
  Begin
    Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
      Using 1, n_�Һ�id;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Err_Special Then
    Raise_Application_Error(-20105, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ԤԼ�ҺŽ���_Insert;
/

--106872:������,2017-04-12,ԤԼ���ձ���ժҪ
Create Or Replace Procedure Zl_ԤԼ�ҺŽ���_����_Insert
(
  No_In            ������ü�¼.No%Type,
  Ʊ�ݺ�_In        ������ü�¼.ʵ��Ʊ��%Type,
  ����id_In        Ʊ��ʹ����ϸ.����id%Type,
  ����id_In        ������ü�¼.����id%Type,
  ����_In          ������ü�¼.��ҩ����%Type,
  ����id_In        ������ü�¼.����id%Type,
  �����_In        ������ü�¼.��ʶ��%Type,
  ����_In          ������ü�¼.����%Type,
  �Ա�_In          ������ü�¼.�Ա�%Type,
  ����_In          ������ü�¼.����%Type,
  ���ʽ_In      ������ü�¼.���ʽ%Type, --���ڴ�Ų��˵�ҽ�Ƹ��ʽ���
  �ѱ�_In          ������ü�¼.�ѱ�%Type,
  ���㷽ʽ_In      Varchar2, --�ֽ�Ľ�������
  �ֽ�֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�ֽ�֧�����ݽ��
  Ԥ��֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱʹ�õ�Ԥ�����
  ����֧��_In      ����Ԥ����¼.��Ԥ��%Type, --�Һ�ʱ�����ʻ�֧�����
  ����ʱ��_In      ������ü�¼.����ʱ��%Type,
  ����_In          �Һ����״̬.���%Type,
  ����Ա���_In    ������ü�¼.����Ա���%Type,
  ����Ա����_In    ������ü�¼.����Ա����%Type,
  ���ɶ���_In      Number := 0,
  �Ǽ�ʱ��_In      ������ü�¼.�Ǽ�ʱ��%Type := Null,
  �����id_In      ����Ԥ����¼.�����id%Type := Null,
  ���㿨���_In    ����Ԥ����¼.���㿨���%Type := Null,
  ����_In          ����Ԥ����¼.����%Type := Null,
  ������ˮ��_In    ����Ԥ����¼.������ˮ��%Type := Null,
  ����˵��_In      ����Ԥ����¼.����˵��%Type := Null,
  ����_In          ���˹Һż�¼.����%Type := Null,
  ����ģʽ_In      Number := 0,
  ���ʷ���_In      Number := 0,
  ��Ԥ������ids_In Varchar2 := Null,
  ��������_In      Number := 0,
  ���½������_In  Number := 0, --�Ƿ������Ա��������Ҫ�Ǵ���ͳһ����Ա��¼��̨�����������
  ժҪ_In          ���˹Һż�¼.ժҪ%Type := Null
) As
  --���α������շѳ�Ԥ���Ŀ���Ԥ���б�
  --��ID�������ȳ��ϴ�δ����ġ�
  Cursor c_Deposit
  (
    v_����id        ������Ϣ.����id%Type,
    v_��Ԥ������ids Varchar2
  ) Is
    Select ����id, NO, Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) As ���, Min(��¼״̬) As ��¼״̬, Nvl(Max(����id), 0) As ����id,
           Max(Decode(��¼����, 1, ID, 0) * Decode(��¼״̬, 2, 0, 1)) As ԭԤ��id
    From ����Ԥ����¼
    Where ��¼���� In (1, 11) And ����id In (Select Column_Value From Table(f_Num2list(v_��Ԥ������ids))) And Nvl(Ԥ�����, 2) = 1
     Having Sum(Nvl(���, 0) - Nvl(��Ԥ��, 0)) <> 0
    Group By NO, ����id
    Order By Decode(����id, Nvl(v_����id, 0), 0, 1), ����id, NO;

  v_Err_Msg Varchar2(255);
  Err_Item    Exception;
  Err_Special Exception;
  v_����Ա���� ���˹Һż�¼.������%Type;
  v_�ֽ�       ���㷽ʽ.����%Type;
  v_�����ʻ�   ���㷽ʽ.����%Type;
  v_��������   �ŶӽкŶ���.��������%Type;
  v_�ű�       ������ü�¼.���㵥λ%Type;
  v_����       ������ü�¼.��ҩ����%Type;
  v_�ŶӺ���   �ŶӽкŶ���.�ŶӺ��� %Type;
  v_ԤԼ��ʽ   ���˹Һż�¼.ԤԼ��ʽ %Type;

  n_��ӡid        Ʊ�ݴ�ӡ����.Id%Type;
  n_Ԥ�����      ����Ԥ����¼.���%Type;
  n_����ֵ        ����Ԥ����¼.���%Type;
  v_��Ԥ������ids Varchar2(4000);

  n_�Һ�id         ���˹Һż�¼.Id%Type;
  n_����̨ǩ���Ŷ� Number;
  n_��id           ����ɿ����.Id%Type;
  n_Count          Number(18);
  n_�Ŷ�           Number;
  n_�����Ŷ�       Number;
  n_��ǰ���       ����Ԥ����¼.���%Type;
  n_Ԥ��id         ����Ԥ����¼.Id%Type;
  n_���ѿ�id       ���ѿ�Ŀ¼.Id%Type;
  n_���ƿ�         Number;

  d_Date         Date;
  d_ԤԼʱ��     ������ü�¼.����ʱ��%Type;
  d_����ʱ��     Date;
  d_�Ŷ�ʱ��     Date;
  n_ʱ��         Number := 0;
  n_����         Number := 0;
  v_��������     Varchar2(2000);
  v_��ǰ����     Varchar2(500);
  n_������     ����Ԥ����¼.��Ԥ��%Type;
  v_�������     ����Ԥ����¼.�������%Type;
  v_���㷽ʽ     ����Ԥ����¼.���㷽ʽ%Type;
  n_��������־   Number(3);
  v_�Ŷ����     �ŶӽкŶ���.�Ŷ����%Type;
  n_����ģʽ     ������Ϣ.����ģʽ%Type;
  n_Ʊ��         Ʊ��ʹ����ϸ.Ʊ��%Type;
  v_���ʽ     ���˹Һż�¼.ҽ�Ƹ��ʽ%Type;
  n_����ģʽ     Number := 0;
  n_�����¼id   ���˹Һż�¼.�����¼id%Type;
  n_�³����¼id ���˹Һż�¼.�����¼id%Type;
  n_��Դid       �ٴ������¼.��Դid%Type;
  n_ԤԼ˳���   �ٴ�������ſ���.ԤԼ˳���%Type;
  n_�ɷ�ʱ��     �ٴ������¼.�Ƿ��ʱ��%Type;
  n_����ſ���   �ٴ������¼.�Ƿ���ſ���%Type;
  n_�ɿ���id     �ٴ������¼.����id%Type;
  n_����Ŀid     �ٴ������¼.��Ŀid%Type;
  n_��ҽ��id     �ٴ������¼.ҽ��id%Type;
  n_�Һ�ģʽ     Number(3);
  d_����ʱ��     Date;
  v_Paratemp     Varchar2(500);
  v_Registtemp   Varchar2(500);
  n_���         Number(3);
  n_��ſ���     �ٴ������¼.�Ƿ���ſ���%Type;
  v_���ϰ�ʱ��   �ٴ������¼.�ϰ�ʱ��%Type;
Begin
  n_��id          := Zl_Get��id(����Ա����_In);
  v_��Ԥ������ids := Nvl(��Ԥ������ids_In, ����id_In);
  v_Paratemp      := Nvl(zl_GetSysParameter('�Һ��Ű�ģʽ'), 0);
  n_����ģʽ      := Nvl(zl_GetSysParameter('ԤԼ����ģʽ', 1111), 0);
  n_�Һ�ģʽ      := To_Number(Substr(v_Paratemp, 1, 1));
  If n_�Һ�ģʽ = 1 Then
    Begin
      d_����ʱ�� := To_Date(Substr(v_Paratemp, 3), 'yyyy-mm-dd hh24:mi:ss');
    Exception
      When Others Then
        d_����ʱ�� := Null;
    End;
  End If;

  --��ȡ���㷽ʽ����
  Begin
    Select ���� Into v_�ֽ� From ���㷽ʽ Where ���� = 1;
  Exception
    When Others Then
      v_�ֽ� := '�ֽ�';
  End;
  Begin
    Select ���� Into v_�����ʻ� From ���㷽ʽ Where ���� = 3;
  Exception
    When Others Then
      v_�����ʻ� := '�����ʻ�';
  End;
  If �Ǽ�ʱ��_In Is Null Then
    Select Sysdate Into d_Date From Dual;
  Else
    d_Date := �Ǽ�ʱ��_In;
  End If;

  --���¹Һ����״̬
  Begin
    Select �ű�, ����, Trunc(����ʱ��), ����ʱ��, ԤԼ��ʽ, �����¼id
    Into v_�ű�, v_����, d_ԤԼʱ��, d_����ʱ��, v_ԤԼ��ʽ, n_�����¼id
    From ���˹Һż�¼
    Where ��¼���� = 2 And ��¼״̬ = 1 And Rownum = 1 And NO = No_In;
  Exception
    When Others Then
      Select Max(������) Into v_����Ա���� From ���˹Һż�¼ Where ��¼���� = 2 And ��¼״̬ In (1, 3) And NO = No_In;
      If v_����Ա���� Is Null Then
        v_Err_Msg := '��ǰԤԼ�Һŵ��ѱ�ȡ��';
        Raise Err_Item;
      Else
        If v_����Ա���� = ����Ա����_In Then
          v_Err_Msg := '��ǰԤԼ�Һŵ��ѱ�����';
          Raise Err_Special;
        Else
          v_Err_Msg := '��ǰԤԼ�Һŵ��ѱ������˽���';
          Raise Err_Item;
        End If;
      End If;
  End;

  --�ж��Ƿ��ʱ��
  Select Nvl(�Ƿ��ʱ��, 0), ��Դid, Nvl(�Ƿ���ſ���, 0)
  Into n_ʱ��, n_��Դid, n_��ſ���
  From �ٴ������¼
  Where ID = n_�����¼id;

  If n_ʱ�� = 1 And ��������_In = 0 And n_����ģʽ = 0 Then
    If Trunc(����ʱ��_In) <> Trunc(Sysdate) Then
      v_Err_Msg := '��ʱ�ε�ԤԼ�Һŵ�ֻ�ܵ�����գ�';
      Raise Err_Item;
    End If;
  End If;

  If n_ʱ�� = 0 And ��������_In = 0 Then
    If n_����ģʽ = 0 Then
      If Trunc(����ʱ��_In) = Trunc(Sysdate) Then
        d_����ʱ�� := ����ʱ��_In;
      Else
        d_����ʱ�� := Sysdate;
      End If;
    Else
      d_����ʱ�� := ����ʱ��_In;
    End If;
  Else
    If Not ����ʱ��_In Is Null Then
      d_����ʱ�� := ����ʱ��_In;
    End If;
  End If;

  If d_����ʱ�� Is Not Null Then
    If d_����ʱ�� < d_����ʱ�� Then
      v_Err_Msg := '��ǰԤԼ�Һŵ����ڳ�����Ű�ģʽ���ţ�������' || To_Char(d_����ʱ��, 'yyyy-mm-dd hh24:mi:ss') || '֮ǰ����!';
      Raise Err_Item;
    End If;
  End If;

  If Not v_���� Is Null Then
    If ����_In Is Null Then
      Update �ٴ�������ſ��� Set �Һ�״̬ = 0 Where (��� = v_���� Or ��ע = v_����) And ��¼id = n_�����¼id;
    Else
      If Trunc(d_ԤԼʱ��) <> Trunc(Sysdate) And n_����ģʽ = 0 Then
        If n_ʱ�� = 0 And ��������_In = 0 Then
          --��ǰ���ջ��ӳٽ���
          Update �ٴ�������ſ��� Set �Һ�״̬ = 0 Where ��� = v_���� And ��¼id = n_�����¼id;
        
          Select �Ƿ��ʱ��, �Ƿ���ſ���, ����id, ҽ��id, ��Ŀid, �ϰ�ʱ��
          Into n_�ɷ�ʱ��, n_����ſ���, n_�ɿ���id, n_��ҽ��id, n_����Ŀid, v_���ϰ�ʱ��
          From �ٴ������¼
          Where ID = n_�����¼id;
          Begin
            Select ID
            Into n_�³����¼id
            From �ٴ������¼
            Where ��Դid = n_��Դid And �Ƿ��ʱ�� = n_�ɷ�ʱ�� And �Ƿ���ſ��� = n_����ſ��� And ����id = n_�ɿ���id And
                  Nvl(ҽ��id, 0) = Nvl(n_��ҽ��id, 0) And �ϰ�ʱ�� = v_���ϰ�ʱ�� And Nvl(�Ƿ񷢲�, 0) = 1 And �������� = Trunc(Sysdate) And
                  Rownum < 2;
          Exception
            When Others Then
              v_Err_Msg := '���յ���û�ж�Ӧ�ĳ��ﰲ��,�޷�����!';
              Raise Err_Item;
          End;
        
          Begin
            Select 1
            Into n_����
            From �ٴ�������ſ���
            Where ��¼id = n_�³����¼id And ��� = v_���� And Nvl(�Һ�״̬, 0) = 0;
          Exception
            When Others Then
              n_���� := 0;
          End;
        
          If n_���� = 1 Then
            Update �ٴ�������ſ���
            Set �Һ�״̬ = 1, ����Ա���� = ����Ա����_In
            Where ��¼id = n_�³����¼id And ��� = v_���� And Nvl(�Һ�״̬, 0) = 0;
          Else
            --�����ѱ�ʹ�õ����
            Select Min(���) Into v_���� From �ٴ�������ſ��� Where ��¼id = n_�³����¼id And Nvl(�Һ�״̬, 0) = 0;
            If v_���� Is Null Then
              v_Err_Msg := '���յ���û�п������,�޷�����!';
              Raise Err_Item;
            End If;
            Update �ٴ�������ſ���
            Set �Һ�״̬ = 1, ����Ա���� = ����Ա����_In
            Where ��¼id = n_�³����¼id And ��� = v_���� And Nvl(�Һ�״̬, 0) = 0;
          End If;
        Else
          Select �Ƿ��ʱ��, �Ƿ���ſ���, ����id, ҽ��id, ��Ŀid, �ϰ�ʱ��
          Into n_�ɷ�ʱ��, n_����ſ���, n_�ɿ���id, n_��ҽ��id, n_����Ŀid, v_���ϰ�ʱ��
          From �ٴ������¼
          Where ID = n_�����¼id;
          Begin
            Select ID
            Into n_�³����¼id
            From �ٴ������¼
            Where ��Դid = n_��Դid And �Ƿ��ʱ�� = n_�ɷ�ʱ�� And �Ƿ���ſ��� = n_����ſ��� And ����id = n_�ɿ���id And
                  Nvl(ҽ��id, 0) = Nvl(n_��ҽ��id, 0) And �ϰ�ʱ�� = v_���ϰ�ʱ�� And Nvl(�Ƿ񷢲�, 0) = 1 And �������� = Trunc(Sysdate) And
                  Rownum < 2;
          Exception
            When Others Then
              v_Err_Msg := '���յ���û�ж�Ӧ�ĳ��ﰲ��,�޷�����!';
              Raise Err_Item;
          End;
          Update �ٴ�������ſ���
          Set �Һ�״̬ = 0, ����Ա���� = ����Ա����_In
          Where (��� = v_���� Or ��ע = v_����) And ��¼id = n_�����¼id And Nvl(�Һ�״̬, 0) = 2
          Returning ԤԼ˳��� Into n_ԤԼ˳���;
        
          Update �ٴ�������ſ���
          Set �Һ�״̬ = 1, ����Ա���� = ����Ա����_In, ԤԼ˳��� = n_ԤԼ˳���
          Where ��� = v_���� And ��¼id = n_�³����¼id And Nvl(�Һ�״̬, 0) = 0;
          If Sql% RowCount = 0 Then
            v_Err_Msg := '���յ������' || v_���� || '�ѱ�������ʹ��,�޷�����.';
            Raise Err_Item;
          End If;
        End If;
      Else
        Update �ٴ�������ſ���
        Set �Һ�״̬ = 1, ����Ա���� = ����Ա����_In
        Where (��� = v_���� Or ��ע = v_����) And ��¼id = n_�����¼id;
        If Sql%RowCount = 0 Then
          v_Err_Msg := '���' || v_���� || '�ѱ�������ʹ��,������ѡ��һ�����.';
          Raise Err_Item;
        End If;
      End If;
    End If;
  Else
    If Not ����_In Is Null Then
      If Trunc(d_ԤԼʱ��) <> Trunc(Sysdate) And n_����ģʽ = 0 Then
        Select �Ƿ��ʱ��, �Ƿ���ſ���, ����id, ҽ��id, ��Ŀid, �ϰ�ʱ��
        Into n_�ɷ�ʱ��, n_����ſ���, n_�ɿ���id, n_��ҽ��id, n_����Ŀid, v_���ϰ�ʱ��
        From �ٴ������¼
        Where ID = n_�����¼id;
        Begin
          Select ID
          Into n_�³����¼id
          From �ٴ������¼
          Where ��Դid = n_��Դid And �Ƿ��ʱ�� = n_�ɷ�ʱ�� And �Ƿ���ſ��� = n_����ſ��� And ����id = n_�ɿ���id And
                Nvl(ҽ��id, 0) = Nvl(n_��ҽ��id, 0) And �ϰ�ʱ�� = v_���ϰ�ʱ�� And Nvl(�Ƿ񷢲�, 0) = 1 And �������� = Trunc(Sysdate) And
                Rownum < 2;
        Exception
          When Others Then
            v_Err_Msg := '���յ���û�ж�Ӧ�ĳ��ﰲ��,�޷�����!';
            Raise Err_Item;
        End;
        Update �ٴ�������ſ���
        Set �Һ�״̬ = 0, ����Ա���� = ����Ա����_In
        Where (��� = ����_In Or ��ע = ����_In) And ��¼id = n_�����¼id And Nvl(�Һ�״̬, 0) = 2
        Returning ԤԼ˳��� Into n_ԤԼ˳���;
        Update �ٴ�������ſ���
        Set �Һ�״̬ = 1, ����Ա���� = ����Ա����_In, ԤԼ˳��� = n_ԤԼ˳���
        Where ��� = ����_In And ��¼id = n_�³����¼id And Nvl(�Һ�״̬, 0) = 0;
        If Sql%RowCount = 0 Then
          v_Err_Msg := '���յ������' || ����_In || '�ѱ�������ʹ��,�޷�����.';
          Raise Err_Item;
        End If;
      Else
        Update �ٴ�������ſ���
        Set �Һ�״̬ = 1, ����Ա���� = ����Ա����_In
        Where (��� = ����_In Or ��ע = ����_In) And ��¼id = n_�����¼id;
      
      End If;
      v_���� := ����_In;
    Else
      v_���� := Null;
    End If;
  End If;

  --����������ü�¼
  Update ������ü�¼
  Set ��¼״̬ = 1, ʵ��Ʊ�� = Decode(Nvl(���ʷ���_In, 0), 1, Null, Ʊ�ݺ�_In), ����id = Decode(Nvl(���ʷ���_In, 0), 1, Null, ����id_In),
      ���ʽ�� = Decode(Nvl(���ʷ���_In, 0), 1, Null, ʵ�ս��), ��ҩ���� = ����_In, ����id = ����id_In, ��ʶ�� = �����_In, ���� = ����_In, ���� = ����_In,
      �Ա� = �Ա�_In, ���ʽ = ���ʽ_In, �ѱ� = �ѱ�_In, ����ʱ�� = d_����ʱ��, �Ǽ�ʱ�� = d_Date, ����Ա��� = ����Ա���_In, ����Ա���� = ����Ա����_In,
      �ɿ���id = n_��id, ���ʷ��� = Decode(Nvl(���ʷ���_In, 0), 1, 1, 0), ժҪ = Nvl(ժҪ_In, ժҪ)
  Where ��¼���� = 4 And ��¼״̬ = 0 And NO = No_In;

  v_Registtemp := zl_GetSysParameter('�Һ��Ű�ģʽ');
  If Substr(v_Registtemp, 1, 1) = 1 Then
    Begin
      If To_Date(Substr(v_Registtemp, 3), 'yyyy-mm-dd hh24:mi:ss') > d_����ʱ�� Then
        v_Err_Msg := '����ʱ��' || To_Char(d_����ʱ��, 'yyyy-mm-dd hh24:mi:ss') || 'δ���ó�����Ű�ģʽ,Ŀǰ�޷�����!';
        Raise Err_Item;
      End If;
    Exception
      When Others Then
        Null;
    End;
    Begin
      Select 1
      Into n_���
      From �ٴ������¼
      Where ID = Nvl(n_�³����¼id, n_�����¼id) And d_����ʱ�� Between ͣ�￪ʼʱ�� And ͣ����ֹʱ��;
    Exception
      When Others Then
        n_��� := 0;
    End;
    If n_��� = 1 And Not (n_ʱ�� = 1 And n_��ſ��� = 1) Then
      v_Err_Msg := '����ʱ��' || To_Char(d_����ʱ��, 'yyyy-mm-dd hh24:mi:ss') || '�İ����Ѿ���ͣ��,�޷�����!';
      Raise Err_Item;
    End If;
  End If;

  --���˹Һż�¼
  Update ���˹Һż�¼
  Set ������ = ����Ա����_In, ����ʱ�� = d_Date, ��¼���� = 1, ����id = ����id_In, ����� = �����_In, ����ʱ�� = d_����ʱ��, ���� = ����_In, �Ա� = �Ա�_In,
      ���� = ����_In, ����Ա��� = ����Ա���_In, ����Ա���� = ����Ա����_In, ���� = Decode(Nvl(����_In, 0), 0, Null, ����_In), ���� = v_����, ���� = ����_In,
      �����¼id = Nvl(n_�³����¼id, n_�����¼id), ժҪ = Nvl(ժҪ_In, ժҪ)
  Where ��¼״̬ = 1 And NO = No_In And ��¼���� = 2
  Returning ID Into n_�Һ�id;
  If Sql%NotFound Then
    Begin
      Select ���˹Һż�¼_Id.Nextval Into n_�Һ�id From Dual;
      Begin
        Select ���� Into v_���ʽ From ҽ�Ƹ��ʽ Where ���� = ���ʽ_In And Rownum < 2;
      Exception
        When Others Then
          v_���ʽ := Null;
      End;
      Insert Into ���˹Һż�¼
        (ID, NO, ��¼����, ��¼״̬, ����id, �����, ����, �Ա�, ����, �ű�, ����, ����, ���ӱ�־, ִ�в���id, ִ����, ִ��״̬, ִ��ʱ��, �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����,
         ժҪ, ����, ԤԼ, ԤԼ��ʽ, ������, ����ʱ��, ԤԼʱ��, ����, ҽ�Ƹ��ʽ, �����¼id)
        Select n_�Һ�id, No_In, 1, 1, ����id_In, �����_In, ����_In, �Ա�_In, ����_In, ���㵥λ, �Ӱ��־, ����_In, Null, ִ�в���id, ִ����, 0, Null,
               �Ǽ�ʱ��, ����ʱ��, ����Ա���, ����Ա����, Nvl(ժҪ_In, ժҪ), v_����, 1, Substr(����, 1, 10) As ԤԼ��ʽ, ����Ա����_In,
               Nvl(�Ǽ�ʱ��_In, Sysdate), ����ʱ��, Decode(Nvl(����_In, 0), 0, Null, ����_In), v_���ʽ, Nvl(n_�³����¼id, n_�����¼id)
        From ������ü�¼
        Where ��¼���� = 4 And ��¼״̬ = 1 And Rownum = 1 And NO = No_In;
    Exception
      When Others Then
        v_Err_Msg := '���ڲ���ԭ��,���ݺ�Ϊ��' || No_In || '���Ĳ���' || ����_In || '�Ѿ�������';
        Raise Err_Item;
    End;
  End If;

  --0-����������;1-��ҽ�������̨�Ŷ�;2-�ȷ���,��ҽ��վ
  If Nvl(���ɶ���_In, 0) <> 0 Then
    n_����̨ǩ���Ŷ� := Zl_To_Number(zl_GetSysParameter('����̨ǩ���Ŷ�', 1113));
    If Nvl(n_����̨ǩ���Ŷ�, 0) = 0 Then
      For v_�Һ� In (Select ID, ����, ����, ִ����, ִ�в���id, ����ʱ��, �ű�, ���� From ���˹Һż�¼ Where NO = No_In) Loop
      
        Begin
          Select 1,
                 Case
                   When �Ŷ�ʱ�� < Trunc(Sysdate) Then
                    1
                   Else
                    0
                 End
          Into n_�Ŷ�, n_�����Ŷ�
          From �ŶӽкŶ���
          Where ҵ������ = 0 And ҵ��id = v_�Һ�.Id And Rownum <= 1;
        Exception
          When Others Then
            n_�Ŷ� := 0;
        End;
        If n_�Ŷ� = 0 Then
          --��������
          --����ִ�в��š���������
          n_�Һ�id   := v_�Һ�.Id;
          v_�������� := v_�Һ�.ִ�в���id;
          v_�ŶӺ��� := Zlgetnextqueue(v_�Һ�.ִ�в���id, n_�Һ�id, v_�Һ�.�ű� || '|' || v_�Һ�.����);
          v_�Ŷ���� := Zlgetsequencenum(0, n_�Һ�id, 0);
        
          --�Һ�id_In,����_In,����_In,ȱʡ����_In,��չ_In(������)
          d_�Ŷ�ʱ�� := Zl_Get_Queuedate(n_�Һ�id, v_�Һ�.�ű�, v_�Һ�.����, v_�Һ�.����ʱ��);
          --   ��������_In , ҵ������_In, ҵ��id_In,����id_In,�ŶӺ���_In,�Ŷӱ��_In,��������_In,����ID_IN, ����_In, ҽ������_In,
          Zl_�ŶӽкŶ���_Insert(v_��������, 0, n_�Һ�id, v_�Һ�.ִ�в���id, v_�ŶӺ���, Null, ����_In, ����id_In, v_�Һ�.����, v_�Һ�.ִ����, d_�Ŷ�ʱ��,
                           v_ԤԼ��ʽ, Null, v_�Ŷ����);
        Elsif Nvl(n_�����Ŷ�, 0) = 1 Then
          --���¶��к�
          v_�ŶӺ��� := Zlgetnextqueue(v_�Һ�.ִ�в���id, v_�Һ�.Id, v_�Һ�.�ű� || '|' || Nvl(v_�Һ�.����, 0));
          v_�Ŷ���� := Zlgetsequencenum(0, v_�Һ�.Id, 1);
          --�¶�������_IN, ҵ������_In, ҵ��id_In , ����id_In , ��������_In , ����_In, ҽ������_In ,�ŶӺ���_In
          Zl_�ŶӽкŶ���_Update(v_�Һ�.ִ�в���id, 0, v_�Һ�.Id, v_�Һ�.ִ�в���id, v_�Һ�.����, v_�Һ�.����, v_�Һ�.ִ����, v_�ŶӺ���, v_�Ŷ����);
        
        Else
          --�¶�������_IN, ҵ������_In, ҵ��id_In , ����id_In , ��������_In , ����_In, ҽ������_In ,�ŶӺ���_In
          Zl_�ŶӽкŶ���_Update(v_�Һ�.ִ�в���id, 0, v_�Һ�.Id, v_�Һ�.ִ�в���id, v_�Һ�.����, v_�Һ�.����, v_�Һ�.ִ����);
        End If;
      End Loop;
    End If;
  End If;

  --���ܽ��㵽����Ԥ����¼
  If Nvl(���ʷ���_In, 0) = 0 Then
    If Nvl(�ֽ�֧��_In, 0) = 0 And Nvl(����֧��_In, 0) = 0 And Nvl(Ԥ��֧��_In, 0) = 0 Then
      Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
      Insert Into ����Ԥ����¼
        (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ,
         ��������)
      Values
        (n_Ԥ��id, 4, 1, No_In, Decode(����id_In, 0, Null, ����id_In), v_�ֽ�, 0, �Ǽ�ʱ��_In, ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�',
         n_��id, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, Null, 4);
    End If;
    If Nvl(�ֽ�֧��_In, 0) <> 0 Then
      v_�������� := ���㷽ʽ_In || '|'; --�Կո�ֿ���|��β,û�н�������
      While v_�������� Is Not Null Loop
        v_��ǰ���� := Substr(v_��������, 1, Instr(v_��������, '|') - 1);
        v_���㷽ʽ := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
      
        v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
        n_������ := To_Number(Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1));
      
        v_��ǰ���� := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
        v_������� := Substr(v_��ǰ����, 1, Instr(v_��ǰ����, ',') - 1);
      
        v_��ǰ����   := Substr(v_��ǰ����, Instr(v_��ǰ����, ',') + 1);
        n_��������־ := To_Number(v_��ǰ����);
      
        If n_��������־ = 0 Then
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          Insert Into ����Ԥ����¼
            (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��,
             ������λ, ��������, �������)
          Values
            (n_Ԥ��id, 4, 1, No_In, Decode(����id_In, 0, Null, ����id_In), Nvl(v_���㷽ʽ, v_�ֽ�), Nvl(n_������, 0), �Ǽ�ʱ��_In,
             ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�', n_��id, Null, Null, Null, Null, Null, Null, 4, v_�������);
        Else
          Select ����Ԥ����¼_Id.Nextval Into n_Ԥ��id From Dual;
          Insert Into ����Ԥ����¼
            (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��,
             ������λ, ��������, �������)
          Values
            (n_Ԥ��id, 4, 1, No_In, Decode(����id_In, 0, Null, ����id_In), Nvl(v_���㷽ʽ, v_�ֽ�), Nvl(n_������, 0), �Ǽ�ʱ��_In,
             ����Ա���_In, ����Ա����_In, ����id_In, '�Һ��շ�', n_��id, �����id_In, ���㿨���_In, ����_In, ������ˮ��_In, ����˵��_In, Null, 4, v_�������);
          If Nvl(���㿨���_In, 0) <> 0 Then
            n_���ѿ�id := Null;
            Begin
              Select Nvl(���ƿ�, 0), 1 Into n_���ƿ�, n_Count From �����ѽӿ�Ŀ¼ Where ��� = ���㿨���_In;
            Exception
              When Others Then
                n_Count := 0;
            End;
            If n_Count = 0 Then
              v_Err_Msg := 'û�з���ԭ���㿨����Ӧ���,���ܼ���������';
              Raise Err_Item;
            End If;
            If n_���ƿ� = 1 Then
              Select ID
              Into n_���ѿ�id
              From ���ѿ�Ŀ¼
              Where �ӿڱ�� = ���㿨���_In And ���� = ����_In And
                    ��� = (Select Max(���) From ���ѿ�Ŀ¼ Where �ӿڱ�� = ���㿨���_In And ���� = ����_In);
            End If;
            Zl_���˿������¼_Insert(���㿨���_In, n_���ѿ�id, v_���㷽ʽ, n_������, ����_In, Null, �Ǽ�ʱ��_In, Null, ����id_In, n_Ԥ��id);
          End If;
        End If;
      
        If Nvl(���½������_In, 0) = 0 Then
          Update ��Ա�ɿ����
          Set ��� = Nvl(���, 0) + n_������
          Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = Nvl(v_���㷽ʽ, v_�ֽ�)
          Returning ��� Into n_����ֵ;
        
          If Sql%RowCount = 0 Then
            Insert Into ��Ա�ɿ����
              (�տ�Ա, ���㷽ʽ, ����, ���)
            Values
              (����Ա����_In, Nvl(v_���㷽ʽ, v_�ֽ�), 1, n_������);
            n_����ֵ := n_������;
          End If;
          If Nvl(n_����ֵ, 0) = 0 Then
            Delete From ��Ա�ɿ����
            Where �տ�Ա = ����Ա����_In And ���㷽ʽ = Nvl(v_���㷽ʽ, v_�ֽ�) And ���� = 1 And Nvl(���, 0) = 0;
          End If;
        End If;
      
        v_�������� := Substr(v_��������, Instr(v_��������, '|') + 1);
      End Loop;
    End If;
  End If;

  --���ھ��￨ͨ��Ԥ����Һ�
  If Nvl(Ԥ��֧��_In, 0) <> 0 And Nvl(���ʷ���_In, 0) = 0 Then
    n_Ԥ����� := Ԥ��֧��_In;
    For r_Deposit In c_Deposit(����id_In, v_��Ԥ������ids) Loop
      n_��ǰ��� := Case
                  When r_Deposit.��� - n_Ԥ����� < 0 Then
                   r_Deposit.���
                  Else
                   n_Ԥ�����
                End;
      If r_Deposit.����id = 0 Then
        --��һ�γ�Ԥ��(���Ͻ���ID,���Ϊ0)
        Update ����Ԥ����¼ Set ��Ԥ�� = 0, ����id = ����id_In, �������� = 4 Where ID = r_Deposit.ԭԤ��id;
      End If;
      --���ϴ�ʣ���
      Insert Into ����Ԥ����¼
        (ID, NO, ʵ��Ʊ��, ��¼����, ��¼״̬, ����id, ��ҳid, ����id, ���, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �տ�ʱ��, ����Ա����, ����Ա���, ��Ԥ��,
         ����id, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, Ԥ�����, �������, ��������)
        Select ����Ԥ����¼_Id.Nextval, NO, ʵ��Ʊ��, 11, ��¼״̬, ����id, ��ҳid, ����id, Null, ���㷽ʽ, �������, ժҪ, �ɿλ, ��λ������, ��λ�ʺ�, �Ǽ�ʱ��_In,
               ����Ա����_In, ����Ա���_In, n_��ǰ���, ����id_In, n_��id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ, Ԥ�����, ����id_In, 4
        From ����Ԥ����¼
        Where NO = r_Deposit.No And ��¼״̬ = r_Deposit.��¼״̬ And ��¼���� In (1, 11) And Rownum = 1;
    
      --���²���Ԥ�����
      Update �������
      Set Ԥ����� = Nvl(Ԥ�����, 0) - n_��ǰ���
      Where ����id = r_Deposit.����id And ���� = 1 And ���� = Nvl(1, 2)
      Returning Ԥ����� Into n_����ֵ;
      If Sql%RowCount = 0 Then
        Insert Into ������� (����id, ����, Ԥ�����, ����) Values (r_Deposit.����id, Nvl(1, 2), -1 * n_��ǰ���, 1);
        n_����ֵ := -1 * n_��ǰ���;
      End If;
      If Nvl(n_����ֵ, 0) = 0 Then
        Delete From �������
        Where ����id = r_Deposit.����id And ���� = 1 And Nvl(�������, 0) = 0 And Nvl(Ԥ�����, 0) = 0;
      End If;
    
      --����Ƿ��Ѿ�������
      If r_Deposit.��� < n_Ԥ����� Then
        n_Ԥ����� := n_Ԥ����� - r_Deposit.���;
      Else
        n_Ԥ����� := 0;
      End If;
    
      If n_Ԥ����� = 0 Then
        Exit;
      End If;
    End Loop;
  End If;

  --����ҽ���Һ�
  If Nvl(����֧��_In, 0) <> 0 And Nvl(���ʷ���_In, 0) = 0 Then
    Insert Into ����Ԥ����¼
      (ID, ��¼����, ��¼״̬, NO, ����id, ���㷽ʽ, ��Ԥ��, �տ�ʱ��, ����Ա���, ����Ա����, ����id, ժҪ, �ɿ���id, �����id, ���㿨���, ����, ������ˮ��, ����˵��, ������λ,
       Ԥ�����, �������, ��������)
    Values
      (����Ԥ����¼_Id.Nextval, 4, 1, No_In, ����id_In, v_�����ʻ�, ����֧��_In, d_Date, ����Ա���_In, ����Ա����_In, ����id_In, 'ҽ���Һ�', n_��id,
       Null, Null, Null, Null, Null, Null, Null, ����id_In, 4);
  End If;

  --��ػ��ܱ�Ĵ���
  --��Ա�ɿ����
  If Nvl(����֧��_In, 0) <> 0 And Nvl(���ʷ���_In, 0) = 0 And Nvl(���½������_In, 0) = 0 Then
    Update ��Ա�ɿ����
    Set ��� = Nvl(���, 0) + ����֧��_In
    Where ���� = 1 And �տ�Ա = ����Ա����_In And ���㷽ʽ = v_�����ʻ�
    Returning ��� Into n_����ֵ;
  
    If Sql%RowCount = 0 Then
      Insert Into ��Ա�ɿ���� (�տ�Ա, ���㷽ʽ, ����, ���) Values (����Ա����_In, v_�����ʻ�, 1, ����֧��_In);
      n_����ֵ := ����֧��_In;
    End If;
    If Nvl(n_����ֵ, 0) = 0 Then
      Delete From ��Ա�ɿ���� Where �տ�Ա = ����Ա����_In And ���� = 1 And Nvl(���, 0) = 0;
    End If;
  End If;

  --����Ʊ��ʹ�����
  If Ʊ�ݺ�_In Is Not Null And Nvl(���ʷ���_In, 0) = 0 Then
    Select Ʊ�ݴ�ӡ����_Id.Nextval Into n_��ӡid From Dual;
  
    --��ǰƱ�ݵ�Ʊ��
    Select Ʊ�� Into n_Ʊ�� From Ʊ�����ü�¼ Where ID = Nvl(����id_In, 0);
    --����Ʊ��
    Insert Into Ʊ�ݴ�ӡ���� (ID, ��������, NO) Values (n_��ӡid, 4, No_In);
  
    Insert Into Ʊ��ʹ����ϸ
      (ID, Ʊ��, ����, ����, ԭ��, ����id, ��ӡid, ʹ��ʱ��, ʹ����)
    Values
      (Ʊ��ʹ����ϸ_Id.Nextval, n_Ʊ��, Ʊ�ݺ�_In, 1, 1, ����id_In, n_��ӡid, d_Date, ����Ա����_In);
  
    --״̬�Ķ�
    Update Ʊ�����ü�¼
    Set ��ǰ���� = Ʊ�ݺ�_In, ʣ������ = Decode(Sign(ʣ������ - 1), -1, 0, ʣ������ - 1), ʹ��ʱ�� = d_Date
    Where ID = Nvl(����id_In, 0);
  End If;

  If Nvl(���ʷ���_In, 0) = 1 Then
    --����
    If Nvl(����id_In, 0) = 0 Then
      v_Err_Msg := 'Ҫ��Բ��˵ĹҺŷѽ��м��ʣ������ǽ������˲��ܼ��ʹҺš�';
      Raise Err_Item;
    End If;
    For c_���� In (Select ʵ�ս��, ���˿���id, ��������id, ִ�в���id, ������Ŀid
                 From ������ü�¼
                 Where ��¼���� = 4 And ��¼״̬ = 1 And NO = No_In And Nvl(���ʷ���, 0) = 1) Loop
      --�������
      Update �������
      Set ������� = Nvl(�������, 0) + Nvl(c_����.ʵ�ս��, 0)
      Where ����id = Nvl(����id_In, 0) And ���� = 1 And ���� = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into �������
          (����id, ����, ����, �������, Ԥ�����)
        Values
          (����id_In, 1, 1, Nvl(c_����.ʵ�ս��, 0), 0);
      End If;
    
      --����δ�����
      Update ����δ�����
      Set ��� = Nvl(���, 0) + Nvl(c_����.ʵ�ս��, 0)
      Where ����id = ����id_In And Nvl(��ҳid, 0) = 0 And Nvl(���˲���id, 0) = 0 And Nvl(���˿���id, 0) = Nvl(c_����.���˿���id, 0) And
            Nvl(��������id, 0) = Nvl(c_����.��������id, 0) And Nvl(ִ�в���id, 0) = Nvl(c_����.ִ�в���id, 0) And ������Ŀid + 0 = c_����.������Ŀid And
            ��Դ;�� + 0 = 1;
    
      If Sql%RowCount = 0 Then
        Insert Into ����δ�����
          (����id, ��ҳid, ���˲���id, ���˿���id, ��������id, ִ�в���id, ������Ŀid, ��Դ;��, ���)
        Values
          (����id_In, Null, Null, c_����.���˿���id, c_����.��������id, c_����.ִ�в���id, c_����.������Ŀid, 1, Nvl(c_����.ʵ�ս��, 0));
      End If;
    End Loop;
  End If;
  If Nvl(����id_In, 0) <> 0 Then
    n_����ģʽ := 0;
    Update ������Ϣ
    Set ����ʱ�� = d_����ʱ��, ����״̬ = 1, �������� = ����_In
    Where ����id = ����id_In
    Returning Nvl(����ģʽ, 0) Into n_����ģʽ;
    --ȡ����:
    If Nvl(n_����ģʽ, 0) <> Nvl(����ģʽ_In, 0) Then
      --����ģʽ��ȷ��
      If n_����ģʽ = 1 And Nvl(����ģʽ_In, 0) = 0 Then
        --�����Ѿ���"�����ƺ�����",������"�Ƚ�������Ƶ�",�����Ƿ����δ������
        Select Count(1)
        Into n_Count
        From ����δ�����
        Where ����id = ����id_In And (��Դ;�� In (1, 4) Or ��Դ;�� = 3 And Nvl(��ҳid, 0) = 0) And Nvl(���, 0) <> 0 And Rownum < 2;
        If Nvl(n_Count, 0) <> 0 Then
          --����δ�������ݣ������Ƚ���������ִ��
          v_Err_Msg := '��ǰ���˵ľ���ģʽΪ�����ƺ�����Ҵ���δ����ã�����������ò��˵ľ���ģʽ,������ȶ�δ����ý��ʺ��ٹҺŻ򲻵������˵ľ���ģʽ!';
          Raise Err_Item;
        End If;
        --���
        --δ����ҽ��ҵ��ģ�����ʱ�͹Һŵ�,��Ҫ��֤ͬһ�εľ���ģʽ��һ����(�����Ѿ���飬�����ٴ���)
      End If;
      Update ������Ϣ Set ����ģʽ = ����ģʽ_In Where ����id = ����id_In;
    End If;
  End If;

  --���˵�����Ϣ
  If ����id_In Is Not Null Then
    Update ������Ϣ
    Set ������ = Null, ������ = Null, �������� = Null
    Where ����id = ����id_In And Nvl(��Ժ, 0) = 0 And Exists
     (Select 1
           From ���˵�����¼
           Where ����id = ����id_In And ��ҳid Is Not Null And
                 �Ǽ�ʱ�� = (Select Max(�Ǽ�ʱ��) From ���˵�����¼ Where ����id = ����id_In));
    If Sql%RowCount > 0 Then
      Update ���˵�����¼
      Set ����ʱ�� = d_Date
      Where ����id = ����id_In And ��ҳid Is Not Null And Nvl(����ʱ��, d_Date) >= d_Date;
    End If;
  End If;
  --��Ϣ����
  Begin
    Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
      Using 1, n_�Һ�id;
  Exception
    When Others Then
      Null;
  End;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Err_Special Then
    Raise_Application_Error(-20105, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ԤԼ�ҺŽ���_����_Insert;
/

--108159:���Ʊ�,2017-04-12,ҽ��ִ���޸ĵǼǱ���
CREATE OR REPLACE Procedure Zl_����ҽ��ִ��_Update
( 
  ԭִ��ʱ��_In ����ҽ��ִ��.ִ��ʱ��%Type, 
  ҽ��id_In     ����ҽ��ִ��.ҽ��id%Type, 
  ���ͺ�_In     ����ҽ��ִ��.���ͺ�%Type, 
  Ҫ��ʱ��_In   ����ҽ��ִ��.Ҫ��ʱ��%Type, 
  ��������_In   ����ҽ��ִ��.��������%Type, 
  ִ��ժҪ_In   ����ҽ��ִ��.ִ��ժҪ%Type, 
  ִ����_In     ����ҽ��ִ��.ִ����%Type, 
  ִ��ʱ��_In   ����ҽ��ִ��.ִ��ʱ��%Type, 
  ִ�н��_In   ����ҽ��ִ��.ִ�н��%Type := 1, 
  δִ��ԭ��_In ����ҽ��ִ��.˵��%Type := Null, 
  ����ִ��_In   Number := 0, 
  ����Ա���_In ��Ա��.���%Type := Null, 
  ����Ա����_In ��Ա��.����%Type := Null, 
  ִ�в���id_In ������ü�¼.ִ�в���id%Type := 0 
  --������ҽ��ID_IN=����ִ�е�ҽ��ID���������Ϊ��ʾ�ļ�����Ŀ��ID�� 
  --ִ�в���id_In=������ָ��ִ�в��ŵķ��ã���������0ʱ������ִ�в��� 
) Is 
  --����Ҫִ�е�����¼,�������˸�������,��鲿λ�ļ�¼ 
  --��������,��ҩ�巨,�ɼ�������������,�������ֻ��д�ڵ�һ����Ŀ�ϣ���ִ��״̬��ͬ 
  V_Temp     Varchar2(255); 
  V_��Ա��� ��Ա��.���%Type; 
  V_��Ա���� ��Ա��.����%Type; 
 
  V_��id        ����ҽ����¼.Id%Type; 
  V_�������    ����ҽ����¼.�������%Type; 
  V_ִ�н��old ����ҽ��ִ��.ִ�н��%Type; 
  N_��������old ����ҽ��ִ��.��������%Type; 
 
  V_������Դ ����ҽ����¼.������Դ%Type; 
  V_�������� ����ҽ������.��¼����%Type; 
 
  N_ִ�д��� Number; 
  N_ʣ����� Number; 
  N_ִ��״̬ Number; 
  n_�������� Number;
  n_�������� Number;
  v_Count    Number;
  n_�Ǽ����� Number;
  d_Ҫ��ʱ�� date;
  
  D_�Ǽ�ʱ�� ����ҽ��ִ��.�Ǽ�ʱ��%Type; 
  N_ȡ��ִ�� Number; 
  N_Diffday  Number(18, 3); 
 
  V_Date  Date; 
  V_Error Varchar2(255); 
  Err_Custom Exception; 
Begin 
  --��ǰ������Ա 
  If ����Ա���_In Is Not Null And ����Ա����_In Is Not Null Then 
    V_��Ա��� := ����Ա���_In; 
    V_��Ա���� := ����Ա����_In; 
  Else 
    V_Temp     := Zl_Identity; 
    V_Temp     := Substr(V_Temp, Instr(V_Temp, ';') + 1); 
    V_Temp     := Substr(V_Temp, Instr(V_Temp, ',') + 1); 
    V_��Ա��� := Substr(V_Temp, 1, Instr(V_Temp, ',') - 1); 
    V_��Ա���� := Substr(V_Temp, Instr(V_Temp, ',') + 1); 
  End If; 
 
  Select Sysdate Into V_Date From Dual; 
  Select Nvl(ִ�н��, 1), Nvl(��������, 0), �Ǽ�ʱ�� 
  Into V_ִ�н��old, N_��������old, D_�Ǽ�ʱ�� 
  From ����ҽ��ִ�� 
  Where ҽ��id = ҽ��id_In And ���ͺ� = ���ͺ�_In And ִ��ʱ�� = ԭִ��ʱ��_In; 
  -----ȡ��ִ����Ч�������� 
  Select Zl_To_Number(Nvl(zl_GetSysParameter(220), '999')) Into N_ȡ��ִ�� From Dual; 
  Select V_Date - D_�Ǽ�ʱ�� Into N_Diffday From Dual; 
  --�Ǽ�ʱ�䳬��ȡ��ִ�������ļ�¼���������޸�ҽ��ִ����� 
  If N_Diffday > N_ȡ��ִ�� Then 
    V_Error := 'ҽ��ִ�еǼ�ʱ�䳬����ȡ��ִ����Ч�����������޸�ҽ��ִ�������'; 
    Raise Err_Custom; 
  End If; 
  --����ҽ��ִ�� 
  Update ����ҽ��ִ�� 
  Set Ҫ��ʱ�� = Ҫ��ʱ��_In, �������� = ��������_In, ִ��ժҪ = ִ��ժҪ_In, ִ���� = ִ����_In, ִ��ʱ�� = ִ��ʱ��_In, �Ǽ�ʱ�� = V_Date, �Ǽ��� = V_��Ա����, 
      ִ�н�� = ִ�н��_In, ˵�� = δִ��ԭ��_In 
  Where ҽ��id = ҽ��id_In And ���ͺ� = ���ͺ�_In And ִ��ʱ�� = ԭִ��ʱ��_In; 
  --����ִ�д�������ִ�н���޸ĺ���Ҫ���µ��ݵ�ִ��״̬ 
  If V_ִ�н��old <> ִ�н��_In Or N_��������old <> ��������_In Then 
    Select ������Դ, Nvl(���id, ID), ������� 
    Into V_������Դ, V_��id, V_������� 
    From ����ҽ����¼ 
    Where ID = ҽ��id_In; 
 
    If v_������Դ = 2 Then 
      Select Decode(��¼����, 1, 1, Decode(�������, 1, 1, 2)) 
      Into v_�������� 
      From ����ҽ������ 
      Where ���ͺ� = ���ͺ�_In And ҽ��id = ҽ��id_In; 
    Else 
      v_�������� := 1; 
    End If; 
   
    Select Decode(a.ִ��״̬, 1, a.��������, c.�ǼǴ���), Decode(a.ִ��״̬, 1, 0, a.�������� - c.�ǼǴ���) ,A.��������,C.�ǼǴ���
    Into n_ִ�д���, n_ʣ����� ,n_��������,n_�Ǽ�����
    From ����ҽ������ A, 
         (Select ҽ��id_In ҽ��id, ���ͺ�_In ���ͺ�, Nvl(Sum(b.��������), 0) As �ǼǴ��� 
           From ����ҽ��ִ�� B 
           Where b.ҽ��id = ҽ��id_In And b.���ͺ� = ���ͺ�_In And Nvl(b.ִ�н��, 1) <> 0) C 
    Where a.ҽ��id = c.ҽ��id And a.���ͺ� = c.���ͺ� And a.ҽ��id = ҽ��id_In And a.���ͺ� = ���ͺ�_In; 
   
    --���ȫ��ִ����״̬Ϊ1��δִ��״̬Ϊ0������ִ��״̬Ϊ2 
    Select Decode(N_ʣ�����, 0, 1, Decode(N_ִ�д���, 0, 0, 2)) Into N_ִ��״̬ From Dual; 
    
    --����ҽ��ִ�мƼ�.ִ��״̬
    If n_�������� > 0 Then
      Select Count(distinct Ҫ��ʱ��) Into v_Count From ҽ��ִ�мƼ� Where ҽ��ID = ҽ��ID_IN And ���ͺ� = ���ͺ�_IN;
      If v_Count > 0 Then
        n_�������� := n_�������� / v_Count;
        --��ִ������+�������� �ܹ��ܹ�ִ�ж��ٸ�ʱ���,ȡ�������
        v_Count := ceil((n_�Ǽ����� ) / n_��������);
		If n_�Ǽ����� = 0 Then
			Update ҽ��ִ�мƼ� Set ִ��״̬ = 0 Where ҽ��id = ҽ��ID_IN And ���ͺ� = ���ͺ�_IN And NVL(ִ��״̬,0) <> 2;
		Else
	        --��ȡִ�н���Ҫ��ʱ�� 
	        Select Ҫ��ʱ�� Into d_Ҫ��ʱ��
	        From (Select Ҫ��ʱ��, Rownum As ����
	               From (Select Distinct Ҫ��ʱ�� From ҽ��ִ�мƼ� Where ҽ��id = ҽ��ID_IN And ���ͺ� = ���ͺ�_IN Order By Ҫ��ʱ��))
	        Where ���� = v_Count;
	        
	        If Not d_Ҫ��ʱ�� Is Null Then
	          --�ȼ���Ƿ��Ѿ��˷�
	          Select Max(NVL(ִ��״̬,0)) Into v_Count From ҽ��ִ�мƼ� Where ҽ��id = ҽ��ID_IN And ���ͺ� = ���ͺ�_IN And Ҫ��ʱ�� <= d_Ҫ��ʱ��;
	          If v_Count = 2 Then
	            v_Error := '��ָ����ִ��ʱ��ε�ҽ�������Ѿ����˷ѣ���������ִ�С�'; 
	            Raise Err_Custom; 
	          End If;
	          --���½���Ҫ��ʱ��֮ǰ(��)�ļ�¼ִ��״̬��
	          Update ҽ��ִ�мƼ� Set ִ��״̬ = 1 Where ҽ��id = ҽ��ID_IN And ���ͺ� = ���ͺ�_IN And Ҫ��ʱ�� <= d_Ҫ��ʱ�� And NVL(ִ��״̬,0) <> 2;
	          Update ҽ��ִ�мƼ� Set ִ��״̬ = 0 Where ҽ��id = ҽ��ID_IN And ���ͺ� = ���ͺ�_IN And Ҫ��ʱ�� > d_Ҫ��ʱ�� And NVL(ִ��״̬,0) <> 2;
	        End If;
		End If;
      End If;
    End If;
 
    --ִ�д�����Ϊ0�ͱ��Ϊ����ִ�� 
    If Nvl(����ִ��_In, 0) = 1 Then 
      Update ����ҽ������ 
      Set ִ��״̬ = Decode(N_ִ�д���, 0, 0, 3), ����� = Null, ���ʱ�� = Null 
      Where ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In; 
    Else 
      Update ����ҽ������ 
      Set ִ��״̬ = Decode(N_ִ�д���, 0, 0, 3), ����� = Null, ���ʱ�� = Null 
      Where ִ��״̬ In (0, 3) And ���ͺ� + 0 = ���ͺ�_In And 
            ҽ��id In (Select ID From ����ҽ����¼ Where (ID = V_��id Or ���id = V_��id) And ������� = V_�������); 
    End If; 
 
    If V_�������� = 2 Then 
      If Nvl(����ִ��_In, 0) = 1 Then 
        Update סԺ���ü�¼ A 
        Set ִ��״̬ = N_ִ��״̬, ִ���� = Decode(N_ִ��״̬, 0, Null, ִ����_In), ִ��ʱ�� = Decode(N_ִ��״̬, 0, Null, ִ��ʱ��_In) 
        Where �շ���� Not In ('5', '6', '7') And (ִ�в���id_In = 0 Or A.ִ�в���id = ִ�в���id_In) And Not Exists 
         (Select 1 From �������� Where ����id = A.�շ�ϸĿid And �������� = 1) And A.��¼״̬ In (0, 1, 3) And 
              (ҽ�����, NO, ��¼����) In 
              (Select ҽ��id, NO, ��¼���� 
               From ����ҽ������ 
               Where ִ��״̬ = Decode(N_ִ�д���, 0, 0, 3) And ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In); 
      Else 
        Update סԺ���ü�¼ A 
        Set ִ��״̬ = N_ִ��״̬, ִ���� = Decode(N_ִ��״̬, 0, Null, ִ����_In), ִ��ʱ�� = Decode(N_ִ��״̬, 0, Null, ִ��ʱ��_In) 
        Where �շ���� Not In ('5', '6', '7') And (ִ�в���id_In = 0 Or A.ִ�в���id = ִ�в���id_In) And Not Exists 
         (Select 1 From �������� Where ����id = A.�շ�ϸĿid And �������� = 1) And A.��¼״̬ In (0, 1, 3) And 
              (ҽ�����, NO, ��¼����) In 
              (Select ҽ��id, NO, ��¼���� 
               From ����ҽ������ 
               Where ִ��״̬ = Decode(N_ִ�д���, 0, 0, 3) And ���ͺ� + 0 = ���ͺ�_In And 
                     ҽ��id In 
                     (Select ID From ����ҽ����¼ Where (ID = V_��id Or ���id = V_��id) And ������� = V_�������)); 
      End If; 
    Else 
      If Nvl(����ִ��_In, 0) = 1 Then 
        Update ������ü�¼ A 
        Set ִ��״̬ = N_ִ��״̬, ִ���� = Decode(N_ִ��״̬, 0, Null, ִ����_In), ִ��ʱ�� = Decode(N_ִ��״̬, 0, Null, ִ��ʱ��_In) 
        Where �շ���� Not In ('5', '6', '7') And (ִ�в���id_In = 0 Or A.ִ�в���id = ִ�в���id_In) And Not Exists 
         (Select 1 From �������� Where ����id = A.�շ�ϸĿid And �������� = 1) And A.��¼״̬ In (0, 1, 3) And 
              (ҽ�����, NO, ��¼����) In 
              (Select ҽ��id, NO, ��¼���� 
               From ����ҽ������ 
               Where ִ��״̬ = Decode(N_ִ�д���, 0, 0, 3) And ҽ��id = ҽ��id_In And ���ͺ� + 0 = ���ͺ�_In); 
      Else 
        Update ������ü�¼ A 
        Set ִ��״̬ = N_ִ��״̬, ִ���� = Decode(N_ִ��״̬, 0, Null, ִ����_In), ִ��ʱ�� = Decode(N_ִ��״̬, 0, Null, ִ��ʱ��_In) 
        Where �շ���� Not In ('5', '6', '7') And (ִ�в���id_In = 0 Or A.ִ�в���id = ִ�в���id_In) And Not Exists 
         (Select 1 From �������� Where ����id = A.�շ�ϸĿid And �������� = 1) And A.��¼״̬ In (0, 1, 3) And 
              (ҽ�����, NO, ��¼����) In 
              (Select ҽ��id, NO, ��¼���� 
               From ����ҽ������ 
               Where ִ��״̬ = Decode(N_ִ�д���, 0, 0, 3) And ���ͺ� + 0 = ���ͺ�_In And 
                     ҽ��id In 
                     (Select ID From ����ҽ����¼ Where (ID = V_��id Or ���id = V_��id) And ������� = V_�������)); 
      End If; 
    End If; 
  End If; 
Exception 
  When Err_Custom Then 
    Raise_Application_Error(-20101, '[ZLSOFT]' || V_Error || '[ZLSOFT]'); 
  When Others Then 
    zl_ErrorCenter(SQLCode, SQLErrM); 
End Zl_����ҽ��ִ��_Update;
/

--107950:Ϳ����,2017-04-10,���Ӳ�����ʽ�����еĻ��߻�����Ϣ������ʾ
CREATE OR REPLACE Procedure Zl_������Ϣ_������Ϣ����_PACS
( 
  ����id_In ������Ϣ�䶯.����id%Type, 
  ����id_In Varchar2, --���ﲡ��Ϊ�Һ�ID;סԺ����Ϊ��ҳID;��첡��Ϊ��쵥�� 
  ����_In   ������Ϣ.����%Type, 
  �Ա�_In   ������Ϣ.�Ա�%Type, 
  ����_In   ������Ϣ.����%Type, 
  ����_In   Number,--1-����;2-סԺ;3-��� 
  ˵��_Out  Out ������Ϣ�䶯.˵��%type --���� 
) As 
  Cursor c_AdviceID1 is select a.id as ҽ��id from ����ҽ����¼ a, Ӱ�����¼ b 
                      where a.id=b.ҽ��id and a.�Һŵ�=(select NO from ���˹Һż�¼ where id=to_number(����id_In)) and a.���id is null; --���� 
  Cursor c_AdviceID2 is Select a.id as ҽ��id from ����ҽ����¼ a, Ӱ�����¼ b 
                      where a.id=b.ҽ��id and a.����id=����id_In and a.��ҳid=to_number(����id_In) and a.���id is null;     --סԺ 
  Cursor c_AdviceID3 is Select a.id as ҽ��id from ����ҽ����¼ a where a.�Һŵ�=����id_In and a.����id=����id_In and a.������Դ=4;    --��� 
  Err_Custom Exception; 
  V_Error Varchar2(2000); 
  v_ִ�п��� ���ű�.����%type;  --��ǰҽ����ִ�п��� 
  v_����ִ�п������ Varchar2(2000);--����ҽ��ʱ�����е�ִ�п��� 
  v_��Ŀ���� ������ĿĿ¼.����%type;  --��ǰҽ����Ӧ����Ŀ���� 
  v_ִ����Ŀ������� Varchar2(2000);--����ҽ���ǣ����е���Ŀ���� 
  n_Type Number(1);        --ǩ�����ͣ�1������ǩ����2������ǩ�� 
Begin 
  --���е���ǩ��ʱ�����м�¼�����޸� 
  if ����_In= 1 then  --���� 
    For Row_Cols1 In c_AdviceID1 Loop 
      begin 
        select ǩ������ into n_Type 
        from(Select substr(��������,1,1) as ǩ������ From ���Ӳ������� Where ��������= 8 And �ļ�ID=(Select ����ID From ����ҽ������ Where ҽ��ID= Row_Cols1.ҽ��id) order by ǩ������ desc) 
        where rownum=1; 
 
        select ���� into v_��Ŀ���� from ������ĿĿ¼ where id=(select ������Ŀid from ����ҽ����¼ where id=Row_Cols1.ҽ��id); 
      Exception 
      When Others Then 
        null; 
      end; 
 
      if n_Type is not null then 
        if n_Type=1 then 
          v_ִ����Ŀ�������:=v_ִ����Ŀ�������||'��'||v_��Ŀ����; 
        elsif n_Type=2 then 
          V_Error:='���ˡ�'||����_In||'���� '||v_��Ŀ����||' ��Ŀ�ѽ��й�����ǩ�������ܽ��в�����Ϣ�޸Ĳ�����'; 
          Raise Err_Custom; 
        End if; 
      end if; 
    end loop; 
  elsif ����_In=2 then  --סԺ 
    For Row_Cols2 In c_AdviceID2 Loop 
      begin 
        select ǩ������ into n_Type 
        from(Select substr(��������,1,1) as ǩ������ From ���Ӳ������� Where ��������= 8 And �ļ�ID=(Select ����ID From ����ҽ������ Where ҽ��ID= Row_Cols2.ҽ��id) order by ǩ������ desc) 
        where rownum=1; 
 
        select ���� into v_��Ŀ���� from ������ĿĿ¼ where id=(select ������Ŀid from ����ҽ����¼ where id=Row_Cols2.ҽ��id); 
      Exception 
      When Others Then 
        null; 
      end; 
 
      if n_Type is not null then 
        if n_Type=1 then 
          v_ִ����Ŀ�������:=v_ִ����Ŀ�������||'��'||v_��Ŀ����; 
        elsif n_Type=2 then 
          V_Error:='���ˡ�'||����_In||'���� '||v_��Ŀ����||' ��Ŀ�ѽ��й�����ǩ�������ܽ��в�����Ϣ�޸Ĳ�����'; 
          Raise Err_Custom; 
        End if; 
      end if; 
    end loop; 
  elsif ����_In=3 then  --��� 
    For Row_Cols3 In c_AdviceID3 Loop 
      begin 
        select ǩ������ into n_Type 
        from(Select substr(��������,1,1) as ǩ������ From ���Ӳ������� Where ��������= 8 And �ļ�ID=(Select ����ID From ����ҽ������ Where ҽ��ID= Row_Cols3.ҽ��id) order by ǩ������ desc) 
        where rownum=1; 
 
        select ���� into v_��Ŀ���� from ������ĿĿ¼ where id=(select ������Ŀid from ����ҽ����¼ where id=Row_Cols3.ҽ��id); 
      Exception 
      When Others Then 
        null; 
      end; 
 
      if n_Type is not null then 
        if n_Type=1 then 
          v_ִ����Ŀ�������:=v_ִ����Ŀ�������||'��'||v_��Ŀ����; 
        elsif n_Type=2 then 
          V_Error:='���ˡ�'||����_In||'���� '||v_��Ŀ����||' ��Ŀ�ѽ��й�����ǩ�������ܽ��в�����Ϣ�޸Ĳ�����'; 
          Raise Err_Custom; 
        End if; 
      end if; 
    end loop; 
  end if; 
 
  --�޸���Ϣ 
  if ����_In= 1 then  --���� 
    For Row_Cols1 In c_AdviceID1 Loop 
       Begin
         Update Ӱ�����¼ Set ���� = ����_In, �Ա� = �Ա�_In, ���� = Decode(����_In, Null, ����, ����_In) Where ҽ��id=Row_Cols1.ҽ��id; 
 
         Select ���� Into v_ִ�п��� from ���ű� where id=(select ִ�п���id from Ӱ�����¼ where ҽ��id=Row_Cols1.ҽ��id); 
       Exception 
         When Others Then 
           null; 
       End;

       if nvl(instr(v_����ִ�п������,v_ִ�п���),0)<=0 then 
         v_����ִ�п������:=v_����ִ�п������||','||v_ִ�п���; 
       end if; 
    end loop; 
  elsif ����_In=2 then  --סԺ 
    For Row_Cols2 In c_AdviceID2 Loop 
      Begin
        Update Ӱ�����¼ Set ���� = ����_In, �Ա� = �Ա�_In, ���� = Decode(����_In, Null, ����, ����_In) Where ҽ��id=Row_Cols2.ҽ��id; 
 
        Select ���� Into v_ִ�п��� from ���ű� where id=(select ִ�п���id from Ӱ�����¼ where ҽ��id=Row_Cols2.ҽ��id); 
      Exception 
        When Others Then 
          null; 
      End;

      if nvl(instr(v_����ִ�п������,v_ִ�п���),0)<=0 then 
        v_����ִ�п������:=v_����ִ�п������||','||v_ִ�п���; 
      end if; 
    end loop; 
  elsif ����_In=3 then --��� 
    For Row_Cols3 In c_AdviceID3 Loop 
      Begin
         Update Ӱ�����¼ Set ���� = ����_In, �Ա� = �Ա�_In, ���� = Decode(����_In, Null, ����, ����_In) Where ҽ��id=Row_Cols3.ҽ��id; 
 
         Select ���� Into v_ִ�п��� from ���ű� where id=(select ִ�п���id from Ӱ�����¼ where ҽ��id=Row_Cols3.ҽ��id); 
      Exception 
         When Others Then 
            null; 
      End;
      
      if nvl(instr(v_����ִ�п������,v_ִ�п���),0)<=0 then 
        v_����ִ�п������:=v_����ִ�п������||','||v_ִ�п���; 
      end if; 
    end loop; 
  end if; 
 
  if nvl(v_ִ����Ŀ�������,' ')<>' ' then 
     ˵��_Out:=substr(v_����ִ�п������,2)||':��'||����_In||'���ġ�'|| substr(v_ִ����Ŀ�������,2) ||'����Ӧ��鱨����ǩ������Ҫ�ֹ�������'; 
  else
     ˵��_Out:=substr(v_����ִ�п������,2)||':��'||����_In||'���Ļ�����Ϣ���޸ģ�'; 
  end if;
Exception 
  When Err_Custom Then 
    Raise_Application_Error(-20101, '[ZLSOFT]' || V_Error || '[ZLSOFT]'); 
  When Others Then 
    Zl_Errorcenter(Sqlcode, Sqlerrm); 
End Zl_������Ϣ_������Ϣ����_PACS;
/

--106708:Ƚ����,2017-04-07,Υ���淶������
Drop Procedure Zl_Buildregisterfixedrule;

--106708:Ƚ����,2017-04-07,Υ���淶������
Create Or Replace Procedure Zl_�ٴ������_Addbyfixedrule
(
  Id_In         �ٴ������.Id%Type,
  Newid_In      �ٴ������.Id%Type,
  �������_In   �ٴ������.�������%Type,
  ��ʼʱ��_In   �ٴ����ﰲ��.��ʼʱ��%Type,
  ��ֹʱ��_In   �ٴ����ﰲ��.��ֹʱ��%Type,
  ����Ա����_In �ٴ����ﰲ��.����Ա����%Type := Null,
  �Ǽ�ʱ��_In   �ٴ����ﰲ��.�Ǽ�ʱ��%Type := Null,
  վ��_In       ���ű�.վ��%Type
) As
  -------------------------------------------------------------------------
  --���ܣ��������й̶������������ɳ��µĹ̶������
  -------------------------------------------------------------------------
  n_Count Number;

  n_����id �ٴ������.Id%Type;

  v_����Ա   �ٴ����ﰲ��.����Ա����%Type;
  d_�Ǽ�ʱ�� Date;
  v_Err_Msg  Varchar2(255);
  Err_Item Exception;
Begin
  Begin
    Select 1 Into n_Count From �ٴ������ Where ID = Id_In;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If n_Count = 0 Then
    v_Err_Msg := 'δ����ԭ�������Ϣ��';
    Raise Err_Item;
  End If;

  --����Ƿ�����Ч��Դ
  Begin
    Select 1
    Into n_Count
    From �ٴ������Դ A, ���ű� B, ��Ա�� C, �շ���ĿĿ¼ D
    Where a.����id = b.Id And a.ҽ��id = c.Id(+) And a.��Ŀid = d.Id And a.�Ű෽ʽ = 0 And Nvl(a.�Ƿ�ɾ��, 0) = 0 And
          (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd')) And
          Nvl(b.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
          Nvl(c.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
          Nvl(d.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd')
         --վ��
          And (b.վ�� Is Null Or b.վ�� = վ��_In) And Rownum < 2;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If Nvl(n_Count, 0) = 0 Then
    v_Err_Msg := '��ǰ����������޿ɰ��̶��Ű�ĺ�Դ�����������µĹ̶����ţ�';
    Raise Err_Item;
  End If;

  n_����id := Newid_In;
  If Nvl(n_����id, 0) = 0 Then
    Select �ٴ������_Id.Nextval Into n_����id From Dual;
  End If;

  Insert Into �ٴ������
    (ID, �Ű෽ʽ, �������, ���)
  Values
    (n_����id, 0, �������_In, To_Number(To_Char(��ʼʱ��_In, 'yyyy')));

  d_�Ǽ�ʱ�� := Nvl(�Ǽ�ʱ��_In, Sysdate);
  v_����Ա   := Nvl(����Ա����_In, Zl_Username);

  For c_��Դ In (Select �ٴ����ﰲ��_Id.Nextval As ����id, n_����id As ����id, ԭ����id, ��Դid, ��Ŀid, ҽ��id, ҽ������
               From (Select b.Id As ԭ����id, b.��Դid, c.��Ŀid, c.ҽ��id, c.ҽ������,
                             Row_Number() Over(Partition By c.Id Order By b.��ʼʱ�� Desc) As ���
                      From �ٴ����ﰲ�� B, �ٴ������Դ C, ���ű� D, ��Ա�� E, �շ���ĿĿ¼ F
                      Where b.��Դid = c.Id And c.����id = d.Id And c.ҽ��id = e.Id(+) And c.��Ŀid = f.Id And b.����id = Id_In
                           --��Դ����
                            And c.�Ű෽ʽ = 0 And Nvl(c.�Ƿ�ɾ��, 0) = 0 And
                            (c.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd') Or c.����ʱ�� Is Null) And
                            Nvl(d.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                            Nvl(e.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                            Nvl(f.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd')
                           --վ��
                            And (d.վ�� Is Null Or d.վ�� = վ��_In)) M
               Where ��� = 1) Loop
  
    Insert Into �ٴ����ﰲ��
      (ID, ����id, ��Դid, ��Ŀid, ҽ��id, ҽ������, ��ʼʱ��, ��ֹʱ��, ����Ա����, �Ǽ�ʱ��)
    Values
      (c_��Դ.����id, c_��Դ.����id, c_��Դ.��Դid, c_��Դ.��Ŀid, c_��Դ.ҽ��id, c_��Դ.ҽ������, ��ʼʱ��_In, ��ֹʱ��_In, v_����Ա, d_�Ǽ�ʱ��);
  
    --��������
    For c_���� In (Select ID, ����id, ������Ŀ, �ϰ�ʱ��, �޺���, ��Լ��, �Ƿ���ſ���, �Ƿ��ʱ��, ԤԼ����, ���﷽ʽ, ����id, �Ƿ��ռ
                 From �ٴ���������
                 Where ����id = c_��Դ.ԭ����id) Loop
    
      Zl_�ٴ���������_Copy(c_����.Id, c_��Դ.����id);
    End Loop;
  End Loop;

  --����û�еĳ��ﰲ�ŵĺ�Դ
  For c_��Դ In (Select �ٴ����ﰲ��_Id.Nextval As ����id, n_����id As ����id, a.Id As ��Դid, a.��Ŀid, a.ҽ��id, a.ҽ������
               From �ٴ������Դ A, ���ű� D, ��Ա�� B, �շ���ĿĿ¼ C
               Where a.����id = d.Id And a.ҽ��id = b.Id(+) And a.��Ŀid = c.Id And a.�Ű෽ʽ = 0 And Nvl(a.�Ƿ�ɾ��, 0) = 0 And
                     (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd')) And
                     Nvl(d.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(b.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(c.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd')
                    --վ��
                     And (d.վ�� Is Null Or d.վ�� = վ��_In)
                    
                     And Not Exists (Select 1 From �ٴ����ﰲ�� Where ����id = n_����id And ��Դid = a.Id)) Loop
  
    Insert Into �ٴ����ﰲ��
      (ID, ����id, ��Դid, ��Ŀid, ҽ��id, ҽ������, ��ʼʱ��, ��ֹʱ��, ����Ա����, �Ǽ�ʱ��)
    Values
      (c_��Դ.����id, c_��Դ.����id, c_��Դ.��Դid, c_��Դ.��Ŀid, c_��Դ.ҽ��id, c_��Դ.ҽ������, ��ʼʱ��_In, ��ֹʱ��_In, v_����Ա, d_�Ǽ�ʱ��);
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ������_Addbyfixedrule;
/

--106708:Ƚ����,2017-04-07,Υ���淶������
Drop Procedure Zl_Buildregisterplanbyrecord;

--106708:Ƚ����,2017-04-07,Υ���淶������
Create Or Replace Procedure Zl_�ٴ������_Addbyrecord
(
  ԭ����id_In   �ٴ������.Id%Type,
  �³���id_In   �ٴ������.Id%Type,
  �Ű෽ʽ_In   �ٴ������.�Ű෽ʽ%Type,
  �������_In   �ٴ������.�������%Type,
  ���_In       �ٴ������.���%Type,
  �·�_In       �ٴ������.�·�%Type,
  ����_In       �ٴ������.����%Type,
  ��ʼʱ��_In   �ٴ����ﰲ��.��ʼʱ��%Type,
  ��ֹʱ��_In   �ٴ����ﰲ��.��ֹʱ��%Type,
  ����Ա����_In �ٴ����ﰲ��.����Ա����%Type,
  �Ǽ�ʱ��_In   �ٴ����ﰲ��.�Ǽ�ʱ��%Type,
  վ��_In       ���ű�.վ��%Type,
  ��Աid_In     ��Ա��.Id%Type := Null,
  ɾ������_In   Number := 0
) As
  -------------------------------------------------------------------------
  --���ܣ����ݳ����¼�����µĳ����¼���°���/�ܰ��ţ�
  --������
  --        ��Աid_In ���̶���������Ч����Ϊ0��null��ʾ�ٴ�������Ա�����
  --        ɾ������_In �̶��Ű�תΪ���Ű�/���Ű�ʱ�����ƶ����Ű�/���Ű�ʱ�Ƿ�ɾ���³����ʱ����δʹ�õĳ����¼
  --˵����
  -------------------------------------------------------------------------
  n_Count Number;

  l_��¼id t_Numlist := t_Numlist();
  n_����id �ٴ����ﰲ��.Id%Type;

  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_�����ܳ���id �ٴ������.Id%Type;

  Function Get�����ܳ���id(����id_In �ٴ������.Id%Type) Return �ٴ������.Id%Type Is
    ----------------------------------------
    --���ԭ�ܳ����������(����7��)������Ҫ���ҵ���һ�������������
    ----------------------------------------
    n_����id �ٴ������.Id%Type;
    n_���   �ٴ������.���%Type;
    n_�·�   �ٴ������.�·�%Type;
    n_����   �ٴ������.����%Type;
  
    d_��ʼʱ�� �ٴ����ﰲ��.��ʼʱ��%Type;
    d_����ʱ�� �ٴ����ﰲ��.��ֹʱ��%Type;
  
    --�������ڼ��㵱�µ��������Լ�ÿһ�ܵ�ʱ�䷶Χ
    Cursor c_Weekrange(Date_In Date) Is
      Select Rownum As ����, ��ʼ����, ��������
      From (With Month_Range As (Select Trunc(Date_In) As First_Day, Last_Day(Trunc(Date_In)) As Last_Day From Dual)
             Select Decode(To_Char(First_Day, 'day'), '������', First_Day, Null) As ��ʼ����,
                    Decode(To_Char(First_Day, 'day'), '������', First_Day, Null) As ��������
             From Month_Range
             Union All
             Select Decode(Sign(Trunc(First_Day + 7 * Week, 'day') + 1 - First_Day), 1,
                            Trunc(First_Day + 7 * Week, 'day') + 1, First_Day) As ��ʼ����,
                    Decode(Sign(Trunc(First_Day + 7 * Week, 'day') + 7 - Last_Day), 1, Last_Day,
                            Trunc(First_Day + 7 * Week, 'day') + 7) As ��������
             From Month_Range A, (Select Level - 1 As Week From Dual Connect By Level <= 6) B)
             Where ��ʼ���� <= ��������;
  
  
  Begin
    Begin
      Select ���, �·�, ���� Into n_���, n_�·�, n_���� From �ٴ������ Where ID = ����id_In;
    Exception
      When Others Then
        Return 0;
    End;
  
    If n_��� Is Null Or n_�·� Is Null Or n_���� Is Null Then
      Return 0;
    End If;
  
    For r_Weekrange In c_Weekrange(To_Date(n_��� || '-' || n_�·� || '-01', 'yyyy-mm-dd')) Loop
      If r_Weekrange.���� = n_���� Then
        d_��ʼʱ�� := r_Weekrange.��ʼ����;
        d_����ʱ�� := r_Weekrange.��������;
        Exit;
      End If;
    End Loop;
  
    If d_��ʼʱ�� Is Null Or d_����ʱ�� Is Null Then
      Return 0;
    End If;
    If Trunc(d_����ʱ��) - Trunc(d_��ʼʱ��) >= 6 Then
      Return 0;
    End If;
  
    --���ڿ��µģ�������һ��������������
    n_��� := Null;
    n_�·� := Null;
    n_���� := Null;
    If Trunc(d_��ʼʱ�� - 1, 'month') <> Trunc(d_��ʼʱ��, 'month') Then
      --��ǰ�ǵ�һ��,��ȡ��һ������������
      n_��� := To_Number(To_Char(d_��ʼʱ�� - 1, 'yyyy'));
      n_�·� := To_Number(To_Char(d_��ʼʱ�� - 1, 'mm'));
    Elsif Trunc(d_����ʱ�� + 1, 'month') <> Trunc(d_����ʱ��, 'month') Then
      --��ǰ�����һ��,��ȡ��һ������������
      n_��� := To_Number(To_Char(d_����ʱ�� + 1, 'yyyy'));
      n_�·� := To_Number(To_Char(d_����ʱ�� + 1, 'mm'));
      n_���� := 1;
    Else
      Return 0;
    End If;
  
    --��ȡ���µ���һ��������ID
    Begin
      Select ID
      Into n_����id
      From (Select Rownum As �к�, ID
             From �ٴ������
             Where Nvl(�Ű෽ʽ, 0) = 2 And ��� = n_��� And �·� = n_�·� And (n_���� Is Null Or ���� = n_����)
             Order By ���� Desc)
      Where �к� < 2;
    Exception
      When Others Then
        Return 0;
    End;
  
    Return n_����id;
  End;
Begin
  Begin
    Select 1
    Into n_Count
    From �ٴ������Դ A, ���ű� B, ��Ա�� C, �շ���ĿĿ¼ D
    Where a.����id = b.Id And a.ҽ��id = c.Id(+) And a.��Ŀid = d.Id
         --��Ч��Դ
          And Nvl(a.�Ƿ�ɾ��, 0) = 0 And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd')) And
          Nvl(b.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
          Nvl(c.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
          Nvl(d.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
          (
          --���Ű�
           Nvl(�Ű෽ʽ_In, 0) = 1 And a.�Ű෽ʽ = 1
          --���Ű�
           Or Nvl(�Ű෽ʽ_In, 0) = 2 And
           (
           --��ǰ���������ʱ�䷶Χ�ڲ��������Ű�
            a.�Ű෽ʽ = 2 And Not Exists
            (Select 1
                From �ٴ����ﰲ�� P, �ٴ������ Q
                Where p.����id = q.Id And p.��Դid = a.Id And
                      Not (p.��ֹʱ�� < Trunc(��ʼʱ��_In, 'MONTH') Or p.��ʼʱ�� > Last_Day(��ʼʱ��_In)) And q.�Ű෽ʽ = 1)
           --��ǰ�ѵ���Ϊ�����Ű�,���Ǳ������������Ű࣬����ʣ�µĲ��ֽ��������ܽ����Ű�
            Or a.�Ű෽ʽ = 1 And Exists
            (Select 1
                From �ٴ����ﰲ�� P, �ٴ������ Q
                Where p.����id = q.Id And p.��Դid = a.Id And
                      Not (p.��ֹʱ�� < Trunc(��ʼʱ��_In, 'MONTH') Or p.��ʼʱ�� > Last_Day(��ʼʱ��_In)) And q.�Ű෽ʽ = 2)))
         --��Դ�ڸó����ʱ�䷶Χ���޳����¼
          And Not Exists
     (Select 1
           From �ٴ������¼ O, �ٴ����ﰲ�� P, �ٴ������ Q
           Where o.����id = p.Id And p.����id = q.Id And p.��Դid = a.Id And o.�������� Between ��ʼʱ��_In And ��ֹʱ��_In And
                 (q.�Ű෽ʽ In (1, 2)
                 --ԭ��Ϊ�̶����ﰲ��
                 Or q.�Ű෽ʽ = 0 And (Nvl(ɾ������_In, 0) = 0 Or Nvl(ɾ������_In, 0) = 1 And Exists
                  (Select 1 From ���˹Һż�¼ Where �����¼id = a.Id))))
         --��ǰ��Ա�ɲ����ĺ�Դ
          And (Nvl(��Աid_In, 0) = 0 Or
          (Nvl(a.�Ƿ��ٴ��Ű�, 0) = 1 And Exists (Select 1 From ������Ա Where ����id = a.����id And ��Աid = ��Աid_In)))
         --վ��
          And (b.վ�� Is Null Or b.վ�� = վ��_In) And Rownum < 2;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If n_Count = 0 Then
    If Nvl(�Ű෽ʽ_In, 0) = 1 Then
      v_Err_Msg := '��ǰ����������޿ɰ����Ű�ĺ�Դ�����������µĳ����';
    Else
      v_Err_Msg := '��ǰ����������޿ɰ����Ű�ĺ�Դ�����������µĳ����';
    End If;
    Raise Err_Item;
  End If;

  --��������Ƿ����
  Begin
    Select 1 Into n_Count From �ٴ������ Where ID = �³���id_In;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If Nvl(n_Count, 0) = 0 Then
    Insert Into �ٴ������
      (ID, �Ű෽ʽ, �������, ���, �·�, ����)
    Values
      (�³���id_In, �Ű෽ʽ_In, �������_In, ���_In, �·�_In, ����_In);
  End If;

  --�����ǰ�����ʱ�䷶Χ���޹Һ�����ԤԼ�ĳ����¼(�̶�����)����ɾ���ⲿ�ֳ����¼(��ɾ�������ʱ�ɻָ�)��
  --���޸Ĺ̶����ŵ���ֹʱ�䣬��������ѯ��
  If Nvl(ɾ������_In, 0) = 1 Then
    For c_���� In (Select b.Id As ����id
                 From �ٴ����ﰲ�� B, �ٴ������ C, �ٴ������Դ D
                 Where b.����id = c.Id And b.��Դid = d.Id
                      --��Դ
                       And Nvl(d.�Ƿ�ɾ��, 0) = 0 And (d.����ʱ�� Is Null Or d.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd')) And
                       Nvl(d.�Ű෽ʽ, 0) = �Ű෽ʽ_In
                      --�����б�ʹ���˵ĳ����¼
                       And c.�Ű෽ʽ = 0 And b.��ֹʱ�� >= ��ʼʱ��_In And Not Exists
                  (Select 1
                        From �ٴ������¼ M, ���˹Һż�¼ N
                        Where m.����id = b.Id And m.Id = n.�����¼id And m.�������� >= ��ʼʱ��_In)
                      --��ǰ��Ա�ɲ����ĺ�Դ
                       And (Nvl(��Աid_In, 0) = 0 Or (Nvl(d.�Ƿ��ٴ��Ű�, 0) = 1 And Exists
                        (Select 1 From ������Ա Where ����id = d.����id And ��Աid = ��Աid_In)))) Loop
    
      For c_��¼ In (Select ID As ��¼id From �ٴ������¼ Where ����id = c_����.����id And �������� >= ��ʼʱ��_In) Loop
        l_��¼id.Extend();
        l_��¼id(l_��¼id.Count) := c_��¼.��¼id;
      End Loop;
    End Loop;
  
    Zl_�ٴ������¼_Batchdelete(l_��¼id);
  
  End If;

  --���ԭ�ܳ����������(����7��)������Ҫ���ҵ���һ�������������
  If Nvl(�Ű෽ʽ_In, 0) = 2 Then
    n_�����ܳ���id := Get�����ܳ���id(ԭ����id_In);
  End If;

  For c_��Դ In (Select �³���id_In As ����id, b.Id As ԭ����id, b.��Դid, c.��Ŀid, c.ҽ��id, c.ҽ������
               From �ٴ����ﰲ�� B, �ٴ������Դ C, ���ű� D, ��Ա�� E, �շ���ĿĿ¼ F
               Where b.��Դid = c.Id And c.����id = d.Id And b.ҽ��id = e.Id(+) And c.��Ŀid = f.Id And
                     (b.����id = ԭ����id_In Or b.����id = n_�����ܳ���id)
                    --��Ч��Դ
                     And Nvl(c.�Ƿ�ɾ��, 0) = 0 And (c.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd') Or c.����ʱ�� Is Null) And
                     Nvl(d.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(e.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(f.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     (
                     --���Ű�
                      Nvl(�Ű෽ʽ_In, 0) = 1 And c.�Ű෽ʽ = 1
                     -- ���Ű�
                      Or Nvl(�Ű෽ʽ_In, 0) = 2 And
                      (
                      --��ǰ���������ʱ�䷶Χ�ڲ��������Ű�
                       c.�Ű෽ʽ = 2 And Not Exists
                       (Select 1
                           From �ٴ����ﰲ�� P, �ٴ������ Q
                           Where p.����id = q.Id And p.��Դid = c.Id And
                                 Not (p.��ֹʱ�� < Trunc(��ʼʱ��_In, 'MONTH') Or p.��ʼʱ�� > Last_Day(��ʼʱ��_In)) And q.�Ű෽ʽ = 1)
                      --��ǰ�ѵ���Ϊ�����Ű�,���Ǳ������������Ű࣬����ʣ�µĲ��ֽ��������ܽ����Ű�
                       Or c.�Ű෽ʽ = 1 And Exists
                       (Select 1
                           From �ٴ����ﰲ�� P, �ٴ������ Q
                           Where p.����id = q.Id And p.��Դid = c.Id And
                                 Not (p.��ֹʱ�� < Trunc(��ʼʱ��_In, 'MONTH') Or p.��ʼʱ�� > Last_Day(��ʼʱ��_In)) And q.�Ű෽ʽ = 2)))
                    --��Դ�ڸó����ʱ�䷶Χ���޳����¼
                     And Not Exists
                (Select 1
                      From �ٴ������¼ P
                      Where p.��Դid = c.Id And p.�������� Between ��ʼʱ��_In And ��ֹʱ��_In)
                    --��ǰ��Ա�ɲ����ĺ�Դ
                     And (Nvl(��Աid_In, 0) = 0 Or (Nvl(c.�Ƿ��ٴ��Ű�, 0) = 1 And Exists
                      (Select 1 From ������Ա Where ����id = c.����id And ��Աid = ��Աid_In)))
                    --վ��
                     And (d.վ�� Is Null Or d.վ�� = վ��_In)) Loop
  
    Begin
      Select ID Into n_����id From �ٴ����ﰲ�� Where ����id = c_��Դ.����id And ��Դid = c_��Դ.��Դid;
    Exception
      When Others Then
        n_����id := Null;
    End;
  
    If Nvl(n_����id, 0) = 0 Then
      Select �ٴ����ﰲ��_Id.Nextval Into n_����id From Dual;
    
      Insert Into �ٴ����ﰲ��
        (ID, ����id, ��Դid, ��Ŀid, ҽ��id, ҽ������, ��ʼʱ��, ��ֹʱ��, ����Ա����, �Ǽ�ʱ��)
      Values
        (n_����id, c_��Դ.����id, c_��Դ.��Դid, c_��Դ.��Ŀid, c_��Դ.ҽ��id, c_��Դ.ҽ������, ��ʼʱ��_In, ��ֹʱ��_In, ����Ա����_In, �Ǽ�ʱ��_In);
    End If;
  
    --�����¼
    For c_��¼ In (Select Decode(b.Id, Null, a.Id, b.Id) As ID, c.����
                 From �ٴ������¼ A, �ٴ������¼ B,
                      (Select Trunc(��ʼʱ��_In) + Level - 1 As ����
                        From Dual
                        Connect By Level <= Trunc(��ֹʱ��_In) - Trunc(��ʼʱ��_In) + 1) C
                 Where a.Id = b.���id(+) And a.����id = c_��Դ.ԭ����id And a.���id Is Null And Nvl(a.�Ƿ���ʱ����, 0) = 0
                      --���Ű�
                       And (Nvl(�Ű෽ʽ_In, 0) = 1 And To_Char(a.��������, 'dd') = To_Char(c.����, 'dd')
                       --���Ű�
                       Or Nvl(�Ű෽ʽ_In, 0) = 2 And To_Char(a.��������, 'D') = To_Char(c.����, 'D'))) Loop
      Zl_�ٴ������¼_Copy(c_��¼.Id, n_����id, c_��¼.����, ����Ա����_In, �Ǽ�ʱ��_In);
    End Loop;
  End Loop;

  --����û�еĳ��ﰲ�ŵĺ�Դ
  For c_��Դ In (Select �ٴ����ﰲ��_Id.Nextval As ����id, �³���id_In As ����id, a.Id As ��Դid, a.��Ŀid, a.ҽ��id, a.ҽ������
               From �ٴ������Դ A, ���ű� D, ��Ա�� E, �շ���ĿĿ¼ F
               Where a.����id = d.Id And a.ҽ��id = e.Id(+) And a.��Ŀid = f.Id
                    --��Ч��Դ
                     And Nvl(a.�Ƿ�ɾ��, 0) = 0 And (a.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd') Or a.����ʱ�� Is Null) And
                     Nvl(d.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(e.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(f.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     (
                     --���Ű�
                      Nvl(�Ű෽ʽ_In, 0) = 1 And a.�Ű෽ʽ = 1
                     -- ���Ű�
                      Or Nvl(�Ű෽ʽ_In, 0) = 2 And
                      (
                      --��ǰ���������ʱ�䷶Χ�ڲ��������Ű�
                       a.�Ű෽ʽ = 2 And Not Exists
                       (Select 1
                           From �ٴ����ﰲ�� P, �ٴ������ Q
                           Where p.����id = q.Id And p.��Դid = a.Id And
                                 Not (p.��ֹʱ�� < Trunc(��ʼʱ��_In, 'MONTH') Or p.��ʼʱ�� > Last_Day(��ʼʱ��_In)) And q.�Ű෽ʽ = 1)
                      --��ǰ�ѵ���Ϊ�����Ű�,���Ǳ������������Ű࣬����ʣ�µĲ��ֽ��������ܽ����Ű�
                       Or a.�Ű෽ʽ = 1 And Exists
                       (Select 1
                           From �ٴ����ﰲ�� P, �ٴ������ Q
                           Where p.����id = q.Id And p.��Դid = a.Id And
                                 Not (p.��ֹʱ�� < Trunc(��ʼʱ��_In, 'MONTH') Or p.��ʼʱ�� > Last_Day(��ʼʱ��_In)) And q.�Ű෽ʽ = 2)))
                    --��Դ�ڸó����ʱ�䷶Χ���޳����¼
                     And Not Exists
                (Select 1
                      From �ٴ������¼ P
                      Where p.��Դid = a.Id And p.�������� Between ��ʼʱ��_In And ��ֹʱ��_In)
                    --��ǰ��Ա�ɲ����ĺ�Դ
                     And (Nvl(��Աid_In, 0) = 0 Or (Nvl(a.�Ƿ��ٴ��Ű�, 0) = 1 And Exists
                      (Select 1 From ������Ա Where ����id = a.����id And ��Աid = ��Աid_In)))
                    --վ��
                     And (d.վ�� Is Null Or d.վ�� = վ��_In)
                    
                     And Not Exists (Select 1 From �ٴ����ﰲ�� Where ����id = �³���id_In And ��Դid = a.Id)) Loop
  
    Insert Into �ٴ����ﰲ��
      (ID, ����id, ��Դid, ��Ŀid, ҽ��id, ҽ������, ��ʼʱ��, ��ֹʱ��, ����Ա����, �Ǽ�ʱ��)
    Values
      (c_��Դ.����id, c_��Դ.����id, c_��Դ.��Դid, c_��Դ.��Ŀid, c_��Դ.ҽ��id, c_��Դ.ҽ������, ��ʼʱ��_In, ��ֹʱ��_In, ����Ա����_In, �Ǽ�ʱ��_In);
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ������_Addbyrecord;
/

--106708:Ƚ����,2017-04-07,Υ���淶������
Drop Procedure Zl_Buildregisterplanbytemplet;

--106708:Ƚ����,2017-04-07,Υ���淶������
Create Or Replace Procedure Zl_�ٴ������_Addbytemplet
(
  ģ��id_In   �ٴ������.Id%Type,
  ��Աid_In   ��Ա��.Id%Type,
  ����id_In   �ٴ������.Id%Type,
  �Ű෽ʽ_In �ٴ������.�Ű෽ʽ%Type,
  �������_In �ٴ������.�������%Type,
  ���_In     �ٴ������.���%Type,
  �·�_In     �ٴ������.�·�%Type,
  ����_In     �ٴ������.����%Type,
  ��ʼʱ��_In �ٴ����ﰲ��.��ʼʱ��%Type,
  ��ֹʱ��_In �ٴ����ﰲ��.��ֹʱ��%Type,
  ����Ա_In   �ٴ����ﰲ��.����Ա����%Type,
  �Ǽ�ʱ��_In �ٴ����ﰲ��.�Ǽ�ʱ��%Type,
  վ��_In     ���ű�.վ��%Type,
  ɾ������_In Number := 0
) As
  -------------------------------------------------------------------------
  --����˵��������ģ���Զ������ٴ������¼
  --������
  --        ��Աid_In ���̶���������Ч����Ϊ0��null��ʾ�ٴ�������Ա�����
  --        ɾ������_In �̶��Ű�תΪ���Ű�/���Ű�ʱ�����ƶ����Ű�/���Ű�ʱ�Ƿ�ɾ���³����ʱ����δʹ�õĳ����¼
  --˵����
  -------------------------------------------------------------------------
  Err_Item Exception;
  v_Err_Msg Varchar2(200);
  n_Count   Number(18);

  d_��ѯ���� Date;
  n_��ѯ���� Number;
  v_������Ŀ �ٴ���������.������Ŀ%Type;

  n_�Ƿ���� Number(2);
  d_��ʼʱ�� �ٴ������¼.��ʼʱ��%Type;

  l_��¼id t_Numlist := t_Numlist();

  Procedure Isvisit
  (
    ����id_In       �ٴ����ﰲ��.Id%Type,
    �Ű����_In     �ٴ����ﰲ��.�Ű����%Type,
    ��������_In     �ٴ������¼.��������%Type,
    ��ѯ��ʼʱ��_In �ٴ����ﰲ��.��ʼʱ��%Type,
    ������Ŀ_In     Out �ٴ���������.������Ŀ%Type,
    �Ƿ����_In     Out Number
  ) As
    --�ж��Ƿ�������ȡ������Ŀ
    d_��ѯ���� Date;
    n_��ѯ���� Number;
  Begin
    �Ƿ����_In := 1;
    --��������Ƿ����
    If �Ű����_In = 1 Then
      --�����Ű�
      Select Decode(To_Char(��������_In, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����',
                     Null)
      Into ������Ŀ_In
      From Dual;
      Select Count(1) Into n_Count From �ٴ��������� Where ����id = ����id_In And ������Ŀ = ������Ŀ_In;
      If Nvl(n_Count, 0) = 0 Then
        �Ƿ����_In := 0;
      End If;
    Elsif �Ű����_In = 2 Then
      --�����Ű�
      ������Ŀ_In := '����';
      If Mod(To_Number(To_Char(��������_In, 'dd')), 2) <> 1 Then
        �Ƿ����_In := 0;
      End If;
    Elsif �Ű����_In = 3 Then
      --˫���Ű�
      ������Ŀ_In := '˫��';
      If Mod(To_Number(To_Char(��������_In, 'dd')), 2) <> 0 Then
        �Ƿ����_In := 0;
      End If;
    Elsif �Ű����_In = 4 Or �Ű����_In = 5 Then
      --4-������ѭ,5-��ѭ������
      If �Ű����_In = 4 Then
        d_��ѯ���� := To_Date(To_Char(��������_In, 'yyyy-mm') || To_Char(��ѯ��ʼʱ��_In, '-dd'), 'yyyy-mm-dd');
      Else
        d_��ѯ���� := ��ѯ��ʼʱ��_In;
      End If;
      Begin
        Select To_Number(Substr(������Ŀ, 1, Instr(������Ŀ, '��') - 1))
        Into n_��ѯ����
        From �ٴ���������
        Where ����id = ����id_In And Rownum < 2;
      Exception
        When Others Then
          n_��ѯ���� := 0;
      End;
      If Nvl(n_��ѯ����, 0) > 0 Then
        ������Ŀ_In := n_��ѯ���� || '��';
        If Mod(Trunc(��������_In) - Trunc(d_��ѯ����), n_��ѯ���� + 1) <> 0 Then
          �Ƿ����_In := 0;
        End If;
      End If;
    Elsif �Ű����_In = 6 Then
      --�ض�����
      ������Ŀ_In := To_Number(To_Char(��������_In, 'dd')) || '��';
      Select Count(1) Into n_Count From �ٴ��������� Where ����id = ����id_In And ������Ŀ = ������Ŀ_In;
      If Nvl(n_Count, 0) = 0 Then
        �Ƿ����_In := 0;
      End If;
    End If;
  End;
Begin
  Begin
    Select 1
    Into n_Count
    From �ٴ������Դ A, ���ű� B, ��Ա�� C, �շ���ĿĿ¼ D
    Where a.����id = b.Id And a.ҽ��id = c.Id(+) And a.��Ŀid = d.Id
         --��Ч��Դ
          And Nvl(a.�Ƿ�ɾ��, 0) = 0 And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd')) And
          Nvl(b.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
          Nvl(c.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
          Nvl(d.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
          (
          --���Ű�
           Nvl(�Ű෽ʽ_In, 0) = 1 And a.�Ű෽ʽ = 1
          --���Ű�
           Or Nvl(�Ű෽ʽ_In, 0) = 2 And
           (
           --��ǰ���������ʱ�䷶Χ�ڲ��������Ű�
            a.�Ű෽ʽ = 2 And Not Exists
            (Select 1
                From �ٴ����ﰲ�� P, �ٴ������ Q
                Where p.����id = q.Id And p.��Դid = a.Id And
                      Not (p.��ֹʱ�� < Trunc(��ʼʱ��_In, 'MONTH') Or p.��ʼʱ�� > Last_Day(��ʼʱ��_In)) And q.�Ű෽ʽ = 1)
           --��ǰ�ѵ���Ϊ�����Ű�,���Ǳ������������Ű࣬����ʣ�µĲ��ֽ��������ܽ����Ű�
            Or a.�Ű෽ʽ = 1 And Exists
            (Select 1
                From �ٴ����ﰲ�� P, �ٴ������ Q
                Where p.����id = q.Id And p.��Դid = a.Id And
                      Not (p.��ֹʱ�� < Trunc(��ʼʱ��_In, 'MONTH') Or p.��ʼʱ�� > Last_Day(��ʼʱ��_In)) And q.�Ű෽ʽ = 2)))
         --��Դ�ڸó����ʱ�䷶Χ���޳����¼
          And Not Exists
     (Select 1
           From �ٴ������¼ O, �ٴ����ﰲ�� P, �ٴ������ Q
           Where o.����id = p.Id And p.����id = q.Id And p.��Դid = a.Id And o.�������� Between ��ʼʱ��_In And ��ֹʱ��_In And
                 (q.�Ű෽ʽ In (1, 2)
                 --ԭ��Ϊ�̶����ﰲ��
                 Or q.�Ű෽ʽ = 0 And (Nvl(ɾ������_In, 0) = 0 Or Nvl(ɾ������_In, 0) = 1 And Exists
                  (Select 1 From ���˹Һż�¼ Where �����¼id = a.Id))))
         --��ǰ��Ա�ɲ����ĺ�Դ
          And (Nvl(��Աid_In, 0) = 0 Or
          (Nvl(a.�Ƿ��ٴ��Ű�, 0) = 1 And Exists (Select 1 From ������Ա Where ����id = a.����id And ��Աid = ��Աid_In)))
         --վ��
          And (b.վ�� Is Null Or b.վ�� = վ��_In) And Rownum < 2;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If n_Count = 0 Then
    If Nvl(�Ű෽ʽ_In, 0) = 1 Then
      v_Err_Msg := '��ǰ����������޿ɰ����Ű�ĺ�Դ�����������µĳ����';
    Else
      v_Err_Msg := '��ǰ����������޿ɰ����Ű�ĺ�Դ�����������µĳ����';
    End If;
    Raise Err_Item;
  End If;

  --��������Ƿ����
  Begin
    Select 1 Into n_Count From �ٴ������ Where ID = ����id_In;
  Exception
    When Others Then
      n_Count := 0;
  End;
  If Nvl(n_Count, 0) = 0 Then
    Insert Into �ٴ������
      (ID, �Ű෽ʽ, �������, ���, �·�, ����)
    Values
      (����id_In, �Ű෽ʽ_In, �������_In, ���_In, �·�_In, ����_In);
  End If;

  --�����ǰ�����ʱ�䷶Χ���޹Һ�����ԤԼ�ĳ����¼(�̶�����)����ɾ���ⲿ�ֳ����¼(��ɾ�������ʱ�ɻָ�)��
  --���޸Ĺ̶����ŵ���ֹʱ�䣬��������ѯ��
  If Nvl(ɾ������_In, 0) = 1 Then
    For c_���� In (Select b.Id As ����id
                 From �ٴ����ﰲ�� B, �ٴ������ C, �ٴ������Դ D
                 Where b.����id = c.Id And b.��Դid = d.Id
                      --��Դ
                       And Nvl(d.�Ƿ�ɾ��, 0) = 0 And (d.����ʱ�� Is Null Or d.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd')) And
                       Nvl(d.�Ű෽ʽ, 0) = �Ű෽ʽ_In
                      --�����б�ʹ���˵ĳ����¼
                       And c.�Ű෽ʽ = 0 And b.��ֹʱ�� >= ��ʼʱ��_In And Not Exists
                  (Select 1
                        From �ٴ������¼ M, ���˹Һż�¼ N
                        Where m.����id = b.Id And m.Id = n.�����¼id And m.�������� >= ��ʼʱ��_In)
                      --��ǰ��Ա�ɲ����ĺ�Դ
                       And (Nvl(��Աid_In, 0) = 0 Or (Nvl(d.�Ƿ��ٴ��Ű�, 0) = 1 And Exists
                        (Select 1 From ������Ա Where ����id = d.����id And ��Աid = ��Աid_In)))) Loop
    
      For c_��¼ In (Select ID As ��¼id From �ٴ������¼ Where ����id = c_����.����id And �������� >= ��ʼʱ��_In) Loop
        l_��¼id.Extend();
        l_��¼id(l_��¼id.Count) := c_��¼.��¼id;
      End Loop;
    End Loop;
  
    Zl_�ٴ������¼_Batchdelete(l_��¼id);
  End If;

  For c_��Դ In (Select �ٴ����ﰲ��_Id.Nextval As ����id, ����id_In As ����id, b.Id As ԭ����id, b.��Դid, c.����id, c.��Ŀid, c.ҽ��id, c.ҽ������,
                      b.�Ű����, b.�Ƿ���������, b.�Ƿ����ճ���, b.��ʼʱ��, c.����, Nvl(d.վ��, '-') As վ��
               From �ٴ����ﰲ�� B, �ٴ������Դ C, ���ű� D, ��Ա�� E, �շ���ĿĿ¼ F
               Where b.��Դid = c.Id And c.����id = d.Id And c.ҽ��id = e.Id(+) And c.��Ŀid = f.Id And b.����id = ģ��id_In
                    --��Ч��Դ
                     And Nvl(c.�Ƿ�ɾ��, 0) = 0 And (c.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd') Or c.����ʱ�� Is Null) And
                     Nvl(d.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(e.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(f.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     (
                     --���Ű�
                      Nvl(�Ű෽ʽ_In, 0) = 1 And c.�Ű෽ʽ = 1
                     -- ���Ű�
                      Or Nvl(�Ű෽ʽ_In, 0) = 2 And
                      (
                      --��ǰ���������ʱ�䷶Χ�ڲ��������Ű�
                       c.�Ű෽ʽ = 2 And Not Exists
                       (Select 1
                           From �ٴ����ﰲ�� P, �ٴ������ Q
                           Where p.����id = q.Id And p.��Դid = c.Id And
                                 Not (p.��ֹʱ�� < Trunc(��ʼʱ��_In, 'MONTH') Or p.��ʼʱ�� > Last_Day(��ʼʱ��_In)) And q.�Ű෽ʽ = 1)
                      --��ǰ�ѵ���Ϊ�����Ű�,���Ǳ������������Ű࣬����ʣ�µĲ��ֽ��������ܽ����Ű�
                       Or c.�Ű෽ʽ = 1 And Exists
                       (Select 1
                           From �ٴ����ﰲ�� P, �ٴ������ Q
                           Where p.����id = q.Id And p.��Դid = c.Id And
                                 Not (p.��ֹʱ�� < Trunc(��ʼʱ��_In, 'MONTH') Or p.��ʼʱ�� > Last_Day(��ʼʱ��_In)) And q.�Ű෽ʽ = 2)))
                    --��Դ�ڸó����ʱ�䷶Χ���޳����¼
                     And Not Exists
                (Select 1
                      From �ٴ������¼ P
                      Where p.��Դid = c.Id And p.�������� Between ��ʼʱ��_In And ��ֹʱ��_In)
                    --��ǰ��Ա�ɲ����ĺ�Դ
                     And (Nvl(��Աid_In, 0) = 0 Or (Nvl(c.�Ƿ��ٴ��Ű�, 0) = 1 And Exists
                      (Select 1 From ������Ա Where ����id = c.����id And ��Աid = ��Աid_In)))
                    --վ��
                     And (d.վ�� Is Null Or d.վ�� = վ��_In)) Loop
  
    Insert Into �ٴ����ﰲ��
      (ID, ����id, ��Դid, ��Ŀid, ҽ��id, ҽ������, ��ʼʱ��, ��ֹʱ��, ����Ա����, �Ǽ�ʱ��)
    Values
      (c_��Դ.����id, c_��Դ.����id, c_��Դ.��Դid, c_��Դ.��Ŀid, c_��Դ.ҽ��id, c_��Դ.ҽ������, ��ʼʱ��_In, ��ֹʱ��_In, ����Ա_In, �Ǽ�ʱ��_In);
  
    --�ٴ������¼
    For c_���� In (Select Trunc(��ʼʱ��_In) + Level - 1 As ����,
                        Decode(To_Char(Trunc(��ʼʱ��_In) + Level - 1, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5',
                                '����', '6', '����', '7', '����', Null) As ����
                 From Dual
                 Connect By Level <= Trunc(��ֹʱ��_In) - Trunc(��ʼʱ��_In) + 1) Loop
    
      Isvisit(c_��Դ.ԭ����id, c_��Դ.�Ű����, c_����.����, c_��Դ.��ʼʱ��, v_������Ŀ, n_�Ƿ����);
    
      --�Ƿ����������ղ�����
      --�Ű����:1-�����Ű�;2-�����Ű�;3-˫���Ű�;4-������ѭ;5-��ѭ������;6-�ض�����
      If Instr(',2,3,4,5,', c_��Դ.�Ű����) > 0 And
         (Nvl(c_��Դ.�Ƿ���������, 0) = 0 And c_����.���� = '����' Or Nvl(c_��Դ.�Ƿ����ճ���, 0) = 0 And c_����.���� = '����') Then
        n_�Ƿ���� := 0;
      End If;
    
      If Nvl(n_�Ƿ����, 0) = 1 Then
        For c_��¼ In (With c_ʱ��� As
                        (Select ʱ���, ��ʼʱ��, ��ֹʱ��, ����, վ��, ȱʡʱ��, ��ǰʱ��
                        From (Select ʱ���, ��ʼʱ��, ��ֹʱ��, ����, վ��, ȱʡʱ��, ��ǰʱ��,
                                      Row_Number() Over(Partition By ʱ��� Order By ʱ���, վ�� Asc, ���� Asc) As ���
                               From ʱ���
                               Where Nvl(վ��, c_��Դ.վ��) = c_��Դ.վ�� And Nvl(����, c_��Դ.����) = c_��Դ.����)
                        Where ��� = 1)
                       Select �ٴ������¼_Id.Nextval As ��¼id, m.Id As ����id, m.�ϰ�ʱ��,
                              To_Date(To_Char(c_����.����, 'yyyy-mm-dd ') || To_Char(j.��ʼʱ��, 'hh24:mi:ss'),
                                       'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                              To_Date(To_Char(c_����.����, 'yyyy-mm-dd ') || To_Char(j.��ֹʱ��, 'hh24:mi:ss'),
                                       'yyyy-mm-dd hh24:mi:ss') + Case
                                 When j.��ֹʱ�� <= j.��ʼʱ�� Then
                                  1
                                 Else
                                  0
                               End As ��ֹʱ��,
                              To_Date(To_Char(c_����.����, 'yyyy-mm-dd ') || To_Char(Nvl(j.ȱʡʱ��, j.��ʼʱ��), 'hh24:mi:ss'),
                                       'yyyy-mm-dd hh24:mi:ss') + Case
                                 When j.ȱʡʱ�� < j.��ʼʱ�� Then
                                  1
                                 Else
                                  0
                               End As ȱʡԤԼʱ��,
                              To_Date(To_Char(c_����.����, 'yyyy-mm-dd ') || To_Char(Nvl(j.��ǰʱ��, j.��ʼʱ��), 'hh24:mi:ss'),
                                       'yyyy-mm-dd hh24:mi:ss') + Case
                                 When j.��ʼʱ�� < j.��ǰʱ�� Then
                                  -1
                                 Else
                                  0
                               End As ��ǰ�Һ�ʱ��, m.�޺���, m.��Լ��, m.�Ƿ���ſ���, m.�Ƿ��ʱ��, m.ԤԼ����, a.��Ŀid, a.ҽ��id, a.ҽ������, m.���﷽ʽ,
                              m.����id, m.�Ƿ��ռ
                       From �ٴ����ﰲ�� A, �ٴ��������� M, c_ʱ��� J
                       Where a.Id = m.����id And m.�ϰ�ʱ�� = j.ʱ��� And a.Id = c_��Դ.ԭ����id And m.������Ŀ = v_������Ŀ) Loop
        
          Insert Into �ٴ������¼
            (ID, ����id, ��Դid, ��������, �ϰ�ʱ��, ��ʼʱ��, ��ֹʱ��, ȱʡԤԼʱ��, ��ǰ�Һ�ʱ��, �޺���, ��Լ��, �Ƿ���ſ���, �Ƿ��ʱ��, ԤԼ����, ��Ŀid, ����id, ҽ��id,
             ҽ������, ���﷽ʽ, ����id, �Ǽ���, �Ǽ�ʱ��, �Ƿ��ռ)
          Values
            (c_��¼.��¼id, c_��Դ.����id, c_��Դ.��Դid, c_����.����, c_��¼.�ϰ�ʱ��, c_��¼.��ʼʱ��, c_��¼.��ֹʱ��, c_��¼.ȱʡԤԼʱ��, c_��¼.��ǰ�Һ�ʱ��,
             c_��¼.�޺���, c_��¼.��Լ��, c_��¼.�Ƿ���ſ���, c_��¼.�Ƿ��ʱ��, c_��¼.ԤԼ����, c_��¼.��Ŀid, c_��Դ.����id, c_��¼.ҽ��id, c_��¼.ҽ������,
             c_��¼.���﷽ʽ, c_��¼.����id, ����Ա_In, �Ǽ�ʱ��_In, c_��¼.�Ƿ��ռ);
        
          Begin
            Select ��ʼʱ�� Into d_��ʼʱ�� From �ٴ�����ʱ�� Where ����id = c_��¼.����id And ��� = 1;
          Exception
            When Others Then
              d_��ʼʱ�� := Null;
          End;
          --�����ٴ�������ſ���
          If Nvl(c_��¼.�Ƿ��ʱ��, 0) = 1 And Nvl(c_��¼.�Ƿ���ſ���, 0) = 1 Then
            --��ʱ����������ſ��ƣ�ʹ��"ԤԼ˳���"��¼"�Ƿ�ԤԼ"
            Insert Into �ٴ�������ſ���
              (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, ԤԼ˳���)
              Select c_��¼.��¼id, ���,
                     To_Date(To_Char(c_����.����, 'yyyy-mm-dd ') || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') + Case
                        When Trunc(��ʼʱ��) > Trunc(d_��ʼʱ��) Then
                         1
                        Else
                         0
                      End,
                     To_Date(To_Char(c_����.����, 'yyyy-mm-dd ') || To_Char(��ֹʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') + Case
                        When Trunc(��ֹʱ��) > Trunc(d_��ʼʱ��) Then
                         1
                        Else
                         0
                      End, ��������, �Ƿ�ԤԼ, �Ƿ�ԤԼ
              From �ٴ�����ʱ��
              Where ����id = c_��¼.����id;
          Else
            Insert Into �ٴ�������ſ���
              (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ)
              Select c_��¼.��¼id, ���,
                     To_Date(To_Char(c_����.����, 'yyyy-mm-dd ') || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') + Case
                       When Trunc(��ʼʱ��) > Trunc(d_��ʼʱ��) Then
                        1
                       Else
                        0
                     End,
                     To_Date(To_Char(c_����.����, 'yyyy-mm-dd ') || To_Char(��ֹʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') + Case
                       When Trunc(��ֹʱ��) > Trunc(d_��ʼʱ��) Then
                        1
                       Else
                        0
                     End, ��������, �Ƿ�ԤԼ
              From �ٴ�����ʱ��
              Where ����id = c_��¼.����id;
          End If;
        
          --���������λ�Һſ��Ƽ�¼
          Insert Into �ٴ�����Һſ��Ƽ�¼
            (����, ����, ����, ��¼id, ���, ���Ʒ�ʽ, ����)
            Select ����, ����, ����, c_��¼.��¼id, ���, ���Ʒ�ʽ, ����
            From �ٴ�����Һſ���
            Where ����id = c_��¼.����id;
        
          --�����ٴ��������Ҽ�¼
          Insert Into �ٴ��������Ҽ�¼
            (��¼id, ����id)
            Select c_��¼.��¼id, ����id From �ٴ��������� Where ����id = c_��¼.����id;
        End Loop;
      End If;
    End Loop;
  End Loop;

  --����û�еĳ��ﰲ�ŵĺ�Դ
  For c_��Դ In (Select �ٴ����ﰲ��_Id.Nextval As ����id, ����id_In As ����id, a.Id As ��Դid, a.��Ŀid, a.ҽ��id, a.ҽ������
               From �ٴ������Դ A, ���ű� D, ��Ա�� E, �շ���ĿĿ¼ F
               Where a.����id = d.Id And a.ҽ��id = e.Id(+) And a.��Ŀid = f.Id
                    --��Ч��Դ
                     And Nvl(a.�Ƿ�ɾ��, 0) = 0 And (a.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd') Or a.����ʱ�� Is Null) And
                     Nvl(d.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(e.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(f.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     (
                     --���Ű�
                      Nvl(�Ű෽ʽ_In, 0) = 1 And a.�Ű෽ʽ = 1
                     -- ���Ű�
                      Or Nvl(�Ű෽ʽ_In, 0) = 2 And
                      (
                      --��ǰ���������ʱ�䷶Χ�ڲ��������Ű�
                       a.�Ű෽ʽ = 2 And Not Exists
                       (Select 1
                           From �ٴ����ﰲ�� P, �ٴ������ Q
                           Where p.����id = q.Id And p.��Դid = a.Id And
                                 Not (p.��ֹʱ�� < Trunc(��ʼʱ��_In, 'MONTH') Or p.��ʼʱ�� > Last_Day(��ʼʱ��_In)) And q.�Ű෽ʽ = 1)
                      --��ǰ�ѵ���Ϊ�����Ű�,���Ǳ������������Ű࣬����ʣ�µĲ��ֽ��������ܽ����Ű�
                       Or a.�Ű෽ʽ = 1 And Exists
                       (Select 1
                           From �ٴ����ﰲ�� P, �ٴ������ Q
                           Where p.����id = q.Id And p.��Դid = a.Id And
                                 Not (p.��ֹʱ�� < Trunc(��ʼʱ��_In, 'MONTH') Or p.��ʼʱ�� > Last_Day(��ʼʱ��_In)) And q.�Ű෽ʽ = 2)))
                    --��Դ�ڸó����ʱ�䷶Χ���޳����¼
                     And Not Exists
                (Select 1
                      From �ٴ������¼ P
                      Where p.��Դid = a.Id And p.�������� Between ��ʼʱ��_In And ��ֹʱ��_In)
                    --��ǰ��Ա�ɲ����ĺ�Դ
                     And (Nvl(��Աid_In, 0) = 0 Or (Nvl(a.�Ƿ��ٴ��Ű�, 0) = 1 And Exists
                      (Select 1 From ������Ա Where ����id = a.����id And ��Աid = ��Աid_In)))
                    --վ��
                     And (d.վ�� Is Null Or d.վ�� = վ��_In)
                    
                     And Not Exists (Select 1 From �ٴ����ﰲ�� Where ����id = ����id_In And ��Դid = a.Id)) Loop
  
    Insert Into �ٴ����ﰲ��
      (ID, ����id, ��Դid, ��Ŀid, ҽ��id, ҽ������, ��ʼʱ��, ��ֹʱ��, ����Ա����, �Ǽ�ʱ��)
    Values
      (c_��Դ.����id, c_��Դ.����id, c_��Դ.��Դid, c_��Դ.��Ŀid, c_��Դ.ҽ��id, c_��Դ.ҽ������, ��ʼʱ��_In, ��ֹʱ��_In, ����Ա_In, �Ǽ�ʱ��_In);
  End Loop;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ������_Addbytemplet;
/

--106708:Ƚ����,2017-04-07,Υ���淶������
Create Or Replace Procedure Zl_�ٴ������¼_Stopvisit
(
  ��¼id_In     �ٴ�����ͣ���¼.��¼id%Type,
  ��ʼʱ��_In   �ٴ�����ͣ���¼.��ʼʱ��%Type := Null,
  ��ֹʱ��_In   �ٴ�����ͣ���¼.��ֹʱ��%Type := Null,
  ͣ��ԭ��_In   �ٴ�����ͣ���¼.ͣ��ԭ��%Type := Null,
  ����Ա_In     �ٴ�����ͣ���¼.������%Type := Null,
  ����ʱ��_In   �ٴ�����ͣ���¼.����ʱ��%Type := Null,
  ȡ��ͣ��_In   Number := 0,
  �Ƿ񲻼��_In Number := 0
) As
  --���ܣ�ͣ�����ȡ��ͣ��
  --��Σ�
  --       �Ƿ񲻼��_in ��Ҫ����ͣ��/���ú�Դʱʹ��
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_Count Number;
  d_Cur   Date;

  v_���� �ٴ������Դ.����%Type;
Begin
  If Nvl(ȡ��ͣ��_In, 0) = 0 Then
    --ͣ��
    If Nvl(�Ƿ񲻼��_In, 0) = 0 Then
      Select Count(1) Into n_Count From �ٴ������¼ A Where ID = ��¼id_In And ͣ�￪ʼʱ�� Is Not Null;
      If Nvl(n_Count, 0) <> 0 Then
        v_Err_Msg := '��ǰ�����ѱ����˽�����ͣ���ˢ�����ݺ�鿴��';
        Raise Err_Item;
      End If;
    
      If ��ʼʱ��_In <= Sysdate Then
        v_Err_Msg := 'ͣ��ʱ��Ŀ�ʼʱ��С���˵�ǰʱ�䣬���ܽ���ͣ�������';
        Raise Err_Item;
      End If;
    End If;
  
    Insert Into �ٴ�����ͣ���¼
      (ID, ��¼id, ��ʼʱ��, ��ֹʱ��, ͣ��ԭ��, ������, ����ʱ��, ������, ����ʱ��, �Ǽ���)
      Select �ٴ�����ͣ���¼_Id.Nextval, ��¼id_In, ��ʼʱ��_In, ��ֹʱ��_In, ͣ��ԭ��_In, Nvl(a.ҽ������, ����Ա_In), ����ʱ��_In, ����Ա_In, ����ʱ��_In,
             ����Ա_In
      From �ٴ������¼ A
      Where ID = ��¼id_In;
  
    --����ԭʼ�ٴ������¼
    Select Count(1) Into n_Count From �ٴ������¼ Where ���id = ��¼id_In;
    If Nvl(n_Count, 0) = 0 Then
      For c_��¼ In (Select ID, ����id, To_Date('1900-01-01', 'yyyy-mm-dd') As ��������, �Ǽ���, �Ǽ�ʱ��, �Ƿ񷢲�
                   From �ٴ������¼
                   Where ID = ��¼id_In) Loop
        Zl_�ٴ������¼_Copy(c_��¼.Id, c_��¼.����id, c_��¼.��������, c_��¼.�Ǽ���, c_��¼.�Ǽ�ʱ��, c_��¼.�Ƿ񷢲�, c_��¼.Id);
      End Loop;
    End If;
  
    Update �ٴ������¼
    Set ͣ�￪ʼʱ�� = ��ʼʱ��_In, ͣ����ֹʱ�� = ��ֹʱ��_In, ͣ��ԭ�� = ͣ��ԭ��_In
    Where ID = ��¼id_In;
  
    --����"�ٴ�������ſ���.�Ƿ�ͣ��"Ϊ1
    Update �ٴ�������ſ��� A
    Set �Ƿ�ͣ�� = 1
    Where ��¼id = ��¼id_In And ��ʼʱ�� Between ��ʼʱ��_In And ��ֹʱ��_In And Exists
     (Select 1 From �ٴ������¼ Where ID = a.��¼id And Nvl(�Ƿ���ſ���, 0) = 1 And Nvl(�Ƿ��ʱ��, 0) = 1);
  
    Insert Into ���˷�����Ϣ��¼
      (ID, ֪ͨ����, ��¼id, �Һ�id, ��Դid, ����, ����id, ��Ŀid, ҽ��id, ҽ������, ����id, �Ǽ���, �Ǽ�ʱ��, ֪ͨԭ��)
      Select ���˷�����Ϣ��¼_Id.Nextval, 1, ��¼id_In, �Һ�id, ��Դid, ����, ����id, ��Ŀid, ҽ��id, ҽ������, ����id, ����Ա_In, ����ʱ��_In,
             'ҽ��' || ͣ��ԭ��_In || '����ͣ��'
      From (Select b.Id As �Һ�id, c.Id As ��Դid, c.����, c.����id, a.��Ŀid, a.ҽ��id, a.ҽ������, b.����id
             From �ٴ������¼ A, ���˹Һż�¼ B, �ٴ������Դ C
             Where a.Id = b.�����¼id And a.��Դid = c.Id And b.��¼״̬ = 1 And a.Id = ��¼id_In And
                   (b.��¼���� = 1 And b.����ʱ�� Between a.ͣ�￪ʼʱ�� And a.ͣ����ֹʱ�� Or
                   b.��¼���� = 2 And b.ԤԼʱ�� Between a.ͣ�￪ʼʱ�� And a.ͣ����ֹʱ��));
  
    --��Ϣ����
    -- ͣ������(1-ͣ��,2-ȡ��ͣ��),�����¼ID,ͣ�����
    Begin
      Select b.���� Into v_���� From �ٴ������¼ A, �ٴ������Դ B Where a.��Դid = b.Id And a.Id = ��¼id_In;
      Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
        Using 17, 1 || ',' || ��¼id_In || ',' || v_����;
    Exception
      When Others Then
        Null;
    End;
  Else
    --ȡ��ͣ��
    --���ݼ��
    Select Count(1) Into n_Count From �ٴ������¼ A Where ID = ��¼id_In And ͣ�￪ʼʱ�� Is Null;
    If Nvl(n_Count, 0) <> 0 Then
      If Nvl(�Ƿ񲻼��_In, 0) = 1 Then
        Return;
      End If;
      v_Err_Msg := '��ǰ�����ѱ�����ȡ��ͣ���ˢ�����ݺ�鿴��';
      Raise Err_Item;
    End If;
  
    If Nvl(�Ƿ񲻼��_In, 0) = 0 Then
      Select ͣ����ֹʱ�� Into d_Cur From �ٴ������¼ Where ID = ��¼id_In And ͣ�￪ʼʱ�� Is Not Null;
      If d_Cur <= Sysdate Then
        v_Err_Msg := 'ͣ��ʱ�����ֹʱ��С���˵�ǰʱ�䣬���ܽ���ȡ��ͣ�������';
        Raise Err_Item;
      End If;
      Select Count(1)
      Into n_Count
      From ���˷�����Ϣ��¼
      Where ��¼id = ��¼id_In And ֪ͨ���� = 1 And ������ Is Not Null;
      If n_Count <> 0 Then
        v_Err_Msg := '�ó����¼���ڲ��˷�����Ϣ��¼�����ѱ�����������ȡ��ͣ�������';
        Raise Err_Item;
      End If;
    End If;
  
    Select Count(1)
    Into n_Count
    From (Select ��ʼʱ��, ͣ�￪ʼʱ�� As ��ֹʱ��
           From (Select a.��ʼʱ��, a.��ֹʱ��, a.ͣ�￪ʼʱ��, a.ͣ����ֹʱ��
                  From �ٴ������¼ A, �ٴ������¼ B
                  Where a.��Դid = b.��Դid And a.�������� = b.�������� And b.Id = ��¼id_In And a.Id <> b.Id)
           Where ��ʼʱ�� < ͣ�￪ʼʱ�� And ��ֹʱ�� = ͣ����ֹʱ��
           Union All
           Select ͣ����ֹʱ�� As ��ʼʱ��, ��ֹʱ��
           From (Select a.��ʼʱ��, a.��ֹʱ��, a.ͣ�￪ʼʱ��, a.ͣ����ֹʱ��
                  From �ٴ������¼ A, �ٴ������¼ B
                  Where a.��Դid = b.��Դid And a.�������� = b.�������� And b.Id = ��¼id_In And a.Id <> b.Id)
           Where ��ʼʱ�� = ͣ�￪ʼʱ�� And ��ֹʱ�� > ͣ����ֹʱ��
           Union All
           Select ��ʼʱ��, ͣ�￪ʼʱ�� As ��ֹʱ��
           From (Select a.��ʼʱ��, a.��ֹʱ��, a.ͣ�￪ʼʱ��, a.ͣ����ֹʱ��
                  From �ٴ������¼ A, �ٴ������¼ B
                  Where a.��Դid = b.��Դid And a.�������� = b.�������� And b.Id = ��¼id_In And a.Id <> b.Id)
           Where ��ʼʱ�� < ͣ�￪ʼʱ�� And ��ֹʱ�� > ͣ����ֹʱ��
           Union All
           Select ͣ����ֹʱ�� As ��ʼʱ��, ��ֹʱ��
           From (Select a.��ʼʱ��, a.��ֹʱ��, a.ͣ�￪ʼʱ��, a.ͣ����ֹʱ��
                  From �ٴ������¼ A, �ٴ������¼ B
                  Where a.��Դid = b.��Դid And a.�������� = b.�������� And b.Id = ��¼id_In And a.Id <> b.Id)
           Where ��ʼʱ�� < ͣ�￪ʼʱ�� And ��ֹʱ�� > ͣ����ֹʱ��) M, �ٴ������¼ N
    Where m.��ʼʱ�� < n.��ֹʱ�� And m.��ֹʱ�� > n.��ʼʱ�� And n.Id = ��¼id_In And Rownum < 2;
    If n_Count <> 0 Then
      If Nvl(�Ƿ񲻼��_In, 0) = 1 Then
        Return;
      End If;
      v_Err_Msg := '��ǰ�ϰ�ʱ�ε�ʱ�䷶Χ��ú�Դ����Ŀǰ��Ч���ϰ�ʱ�ε�ʱ�䷶Χ�н��棬�㲻��ȡ��ͣ�';
      Raise Err_Item;
    End If;
  
    Update �ٴ�����ͣ���¼
    Set ȡ���� = ����Ա_In, ȡ��ʱ�� = ����ʱ��_In
    Where ��¼id = ��¼id_In And ����ҽ������ Is Null And ȡ���� Is Null;
  
    Update �ٴ������¼
    Set ͣ�￪ʼʱ�� = Null, ͣ����ֹʱ�� = Null, ͣ��ԭ�� = Null
    Where ID = ��¼id_In And ͣ�￪ʼʱ�� Is Not Null;
  
    --����"�ٴ�������ſ���.�Ƿ�ͣ��"Ϊ0
    Update �ٴ�������ſ��� Set �Ƿ�ͣ�� = 0 Where ��¼id = ��¼id_In And Nvl(�Ƿ�ͣ��, 0) = 1;
  
    Delete ���˷�����Ϣ��¼ Where ��¼id = ��¼id_In And ֪ͨ���� = 1 And ������ Is Null;
  
    --��Ϣ����
    -- ͣ������(1-ͣ��,2-ȡ��ͣ��),�����¼ID,ͣ�����
    Begin
      Select b.���� Into v_���� From �ٴ������¼ A, �ٴ������Դ B Where a.��Դid = b.Id And a.Id = ��¼id_In;
      Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
        Using 17, 2 || ',' || ��¼id_In || ',' || v_����;
    Exception
      When Others Then
        Null;
    End;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ������¼_Stopvisit;
/

--107559:Ƚ����,2017-04-17,������ֹͣ�ﰲ�Ź���
--106712:Ƚ����,2017-04-07,SQL����Ż�
Create Or Replace Procedure Zl_�ٴ�����ͣ��_Audit
(
  ��������_In Number,
  Id_In       �ٴ�����ͣ���¼.Id%Type,
  ������_In   �ٴ�����ͣ���¼.������%Type := Null,
  ����ʱ��_In �ٴ�����ͣ���¼.����ʱ��%Type := Null
) As
  --���ܣ�����ͣ�ﰲ��
  --������
  --       ״̬_In��1-������2-ȡ������
  n_Count Number;

  v_Error Varchar2(255);
  Err_Custom Exception;
Begin
  If Nvl(��������_In, 0) = 1 Then
    --����
    Select Count(1) Into n_Count From �ٴ�����ͣ���¼ Where ID = Id_In And ������ Is Not Null;
    If n_Count <> 0 Then
      v_Error := '�������ѱ������������ٴ�������';
      Raise Err_Custom;
    End If;
  
    Update �ٴ�����ͣ���¼ Set ������ = ������_In, ����ʱ�� = ����ʱ��_In Where ID = Id_In;
    If Sql%NotFound Then
      v_Error := '����������ѱ�ȡ�����룬��ˢ�º�鿴...';
      Raise Err_Custom;
    End If;
  
    --�Գ����¼����ͣ����
    For c_��¼ In (Select a.Id, Greatest(a.��ʼʱ��, b.��ʼʱ��) As ͣ�￪ʼʱ��, Least(a.��ֹʱ��, b.��ֹʱ��) As ͣ����ֹʱ��, b.ͣ��ԭ��, c.����, a.�Ƿ���ſ���,
                        a.�Ƿ��ʱ��
                 From �ٴ������¼ A, �ٴ�����ͣ���¼ B, �ٴ������Դ C
                 Where ((a.����ҽ������ Is Null And a.ҽ��id Is Not Null And a.ҽ������ = b.������) Or
                       (a.����ҽ������ Is Not Null And a.����ҽ��id Is Not Null And a.����ҽ������ = b.������)) And a.��Դid = c.Id And
                       b.Id = Id_In And Not (a.��ʼʱ�� > b.��ֹʱ�� Or a.��ֹʱ�� < b.��ʼʱ��)
                      --ֻ�����ѷ����˵�
                       And Nvl(a.�Ƿ񷢲�, 0) = 1) Loop
    
      Update �ٴ������¼
      Set ͣ�￪ʼʱ�� = c_��¼.ͣ�￪ʼʱ��, ͣ����ֹʱ�� = c_��¼.ͣ����ֹʱ��, ͣ��ԭ�� = c_��¼.ͣ��ԭ��
      Where ID = c_��¼.Id;
    
      --����"�ٴ�������ſ���.�Ƿ�ͣ��"Ϊ1
      Update �ٴ�������ſ��� A
      Set �Ƿ�ͣ�� = 1
      Where ��¼id = c_��¼.Id And ��ʼʱ�� Between c_��¼.ͣ�￪ʼʱ�� And c_��¼.ͣ����ֹʱ�� And Nvl(c_��¼.�Ƿ���ſ���, 0) = 1 And
            Nvl(c_��¼.�Ƿ��ʱ��, 0) = 1;
    
      Insert Into ���˷�����Ϣ��¼
        (ID, ֪ͨ����, ��¼id, �Һ�id, ��Դid, ����, ����id, ��Ŀid, ҽ��id, ҽ������, ����id, �Ǽ���, �Ǽ�ʱ��)
        Select ���˷�����Ϣ��¼_Id.Nextval, 1, a.Id, b.Id, c.Id, c.����, c.����id, a.��Ŀid, a.ҽ��id, a.ҽ������, b.����id, ������_In, ����ʱ��_In
        From �ٴ������¼ A, ���˹Һż�¼ B, �ٴ������Դ C
        Where a.Id = b.�����¼id And a.��Դid = c.Id And b.��¼״̬ = 1 And a.Id = c_��¼.Id And
              (b.��¼���� = 1 And b.����ʱ�� Between a.ͣ�￪ʼʱ�� And a.ͣ����ֹʱ�� Or
              b.��¼���� = 2 And b.ԤԼʱ�� Between a.ͣ�￪ʼʱ�� And a.ͣ����ֹʱ��);
    
      --��Ϣ����
      -- ͣ������(1-ͣ��,2-ȡ��ͣ��),�����¼ID,ͣ�����
      Begin
        Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
          Using 17, 1 || ',' || c_��¼.Id || ',' || c_��¼.����;
      Exception
        When Others Then
          Null;
      End;
    End Loop;
    Return;
  End If;

  --ȡ������
  Select Count(1) Into n_Count From �ٴ�����ͣ���¼ Where ID = Id_In And ��ֹʱ�� < Sysdate;
  If n_Count <> 0 Then
    v_Error := '��ͣ�ﰲ����ʧЧ������ȡ��������';
    Raise Err_Custom;
  End If;

  Select Count(1) Into n_Count From �ٴ�����ͣ���¼ Where ID = Id_In And ʧЧʱ�� Is Not Null;
  If n_Count <> 0 Then
    v_Error := '��ͣ�ﰲ���ѱ���ֹ������ȡ��������';
    Raise Err_Custom;
  End If;

  Select Count(1)
  Into n_Count
  From �ٴ������¼ A, �ٴ�����ͣ���¼ B, ���˷�����Ϣ��¼ C
  Where Nvl(a.����ҽ������, a.ҽ������) = b.������ And Nvl(a.����ҽ��id, a.ҽ��id) Is Not Null And a.Id = c.��¼id And
        (a.��ʼʱ�� Between b.��ʼʱ�� And b.��ֹʱ�� Or a.��ֹʱ�� Between b.��ʼʱ�� And b.��ֹʱ��) And c.������ Is Not Null And b.Id = Id_In;
  If Nvl(n_Count, 0) <> 0 Then
    v_Error := '��ͣ�ﰲ�ŵĲ���ͣ����Ϣ�ѱ���������ȡ��������';
    Raise Err_Custom;
  End If;

  Update �ٴ�����ͣ���¼ Set ������ = Null, ����ʱ�� = Null Where ID = Id_In And ����ʱ�� Is Not Null;
  If Sql%NotFound Then
    v_Error := '�ð��ſ����ѱ�����ȡ����������ˢ�º�鿴...';
    Raise Err_Custom;
  End If;

  For c_��¼ In (Select a.Id, c.����
               From �ٴ������¼ A, �ٴ�����ͣ���¼ B, �ٴ������Դ C
               Where ((a.����ҽ������ Is Null And a.ҽ��id Is Not Null And a.ҽ������ = b.������) Or
                     (a.����ҽ������ Is Not Null And a.����ҽ��id Is Not Null And a.����ҽ������ = b.������)) And a.��Դid = c.Id And
                     b.Id = Id_In And (a.��ʼʱ�� Between b.��ʼʱ�� And b.��ֹʱ�� Or a.��ֹʱ�� Between b.��ʼʱ�� And b.��ֹʱ��) And
                     Nvl(a.�Ƿ񷢲�, 0) = 1) Loop
  
    Update �ٴ������¼ Set ͣ�￪ʼʱ�� = Null, ͣ����ֹʱ�� = Null, ͣ��ԭ�� = Null Where ID = c_��¼.Id;
  
    --����"�ٴ�������ſ���.�Ƿ�ͣ��"Ϊ0
    Update �ٴ�������ſ��� Set �Ƿ�ͣ�� = 0 Where ��¼id = c_��¼.Id And Nvl(�Ƿ�ͣ��, 0) = 1;
  
    Delete ���˷�����Ϣ��¼ Where ��¼id = c_��¼.Id And ֪ͨ���� = 1 And ������ Is Null;
  
    --��Ϣ����
    -- ͣ������(1-ͣ��,2-ȡ��ͣ��),�����¼ID,ͣ�����
    Begin
      Execute Immediate 'Begin ZL_������Ϣ_����(:1,:2); End;'
        Using 17, 2 || ',' || c_��¼.Id || ',' || c_��¼.����;
    Exception
      When Others Then
        Null;
    End;
  End Loop;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ�����ͣ��_Audit;
/

--106556:Ϳ����,2017-04-06,��������ݲ������͵���Ϊxml�ķ�ʽ���д���
--Ӱ�񱨸�������(---���岿��---)***************************************************
CREATE OR REPLACE Package b_PACS_RptCombo Is
  --Create By Hwei;
  --2014/11/25
  Type t_Refcur Is Ref Cursor;

  --1.��  �ܣ����Ӱ�񱨸�����б�
  Procedure p_GetComboList(
    Val Out t_Refcur
	);
  --2.��  �ܣ����Ӱ�񱨸������Ϣ
  Procedure p_AddComboInfo(
    ID_In     In Ӱ�񱨸�����嵥.ID%Type,
    ����_In   In Ӱ�񱨸�����嵥.����%Type,
    ����_In   In Ӱ�񱨸�����嵥.����%Type,
    ˵��_In   In Ӱ�񱨸�����嵥.˵��%Type,
    ����_In   In Ӱ�񱨸�����嵥.����%Type,
    ����_In   In Ӱ�񱨸�����嵥.����%Type,
    ���_In   In Ӱ�񱨸�����嵥.���%Type,
    �༭��_In In Ӱ�񱨸�����嵥.�༭��%Type
	);
  --3.��  ��;�޸�Ӱ�񱨸������Ϣ
  Procedure p_EditComboInfo(
    ID_In     In Ӱ�񱨸�����嵥.ID%Type,
    ����_In   In Ӱ�񱨸�����嵥.����%Type,
    ����_In   In Ӱ�񱨸�����嵥.����%Type,
    ˵��_In   In Ӱ�񱨸�����嵥.˵��%Type,
    ����_In   In Ӱ�񱨸�����嵥.����%Type,
    ����_In   In Ӱ�񱨸�����嵥.����%Type,
    ���_In   In Ӱ�񱨸�����嵥.���%Type,
    �༭��_In In Ӱ�񱨸�����嵥.�༭��%Type
	);
  --4.��  �ܣ�ͨ��IDɾ��Ӱ�񱨸������Ϣ
  Procedure p_DelComboInfo(
    ID_In In Ӱ�񱨸�����嵥.ID%Type
	);
  --5.��  �ܣ�����ID���Ӱ�񱨸������Ϣ
  Procedure p_GetComboInfoByID(
	Val           Out t_Refcur,
	ID_In In Ӱ�񱨸�����嵥.ID%Type
	);
  --6.��  �ܣ����Ӱ�񱨸��������з�����Ϣ
  Procedure p_GetComboAllGroup(
    Val Out t_Refcur
	);
  --7.��  �ܣ����ID��Ӧ��Ӱ�񱨸����Ķ�����Ϣ
  Procedure p_GetComboContent(
	Val           Out t_Refcur,
	ID_In In Ӱ�񱨸�����嵥.ID%Type
	);
  --8.��  �ܣ�����ID��Ӧ��Ӱ�񱨸����Ķ�����Ϣ
  Procedure p_EditComboContent(
	ID_In   In Ӱ�񱨸�����嵥.ID%Type,
	���_In  In Ӱ�񱨸�����嵥.���%Type
	);
  --9.�� �ܣ���ȡ�༭�˶�Ӧ������޸�Ӱ�񱨸������Ϣ
  Procedure p_GetComboInfoByEditor(
	Val           Out t_Refcur,
	�༭��_In In Ӱ�񱨸�����嵥.�༭��%Type
	);
  --10.��  �ܣ�����Ƭ�ε���Ͼ�
  Procedure p_Append_Fragment_Tocombo(
    Text_In In XmlType,
    Id_In   In Ӱ�񱨸�����嵥.ID%Type
	);

  --11.��  �ܣ��޸�Ƭ�ε���Ͼ�
  Procedure p_Update_Combo_Fragment(
    Text_In In XmlType,
    Id_In   In Ӱ�񱨸�����嵥.ID%Type,
    Pid_In  In Varchar2
	);
  --12.��  �ܣ����ݷ���ID��ѯ�ʾ�
  Procedure p_Get_Fragment_By_Typeid(
	Val           Out t_Refcur,
	Id_In In Ӱ�񱨸�����嵥.ID%Type
	);
  --13.��  �ܣ���ȡ��һ������
  Procedure p_Get_ComboNextCode(
    Val Out t_Refcur
	);
end b_PACS_RptCombo;
/

--Ӱ�񱨸�������(---ʵ�ֲ���---)***************************************************
CREATE OR REPLACE Package Body b_PACS_RptCombo Is
  --Create By Hwei;
  --2014/11/25

  --1.��  �ܣ����Ӱ�񱨸�����б�
  Procedure p_GetComboList(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select RawToHex(ID) ID,
             ����,
             ����,
             ˵��,
             ����,
             ����,
             (Nvl(t.���, XmlType('<NULL/>'))).GetClobVal() As ���,
             �༭��,
             ���༭ʱ��
        From Ӱ�񱨸�����嵥 t;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetComboList;

  --2.��  �ܣ����Ӱ�񱨸������Ϣ
  Procedure p_AddComboInfo(
    ID_In     In Ӱ�񱨸�����嵥.ID%Type,
    ����_In   In Ӱ�񱨸�����嵥.����%Type,
    ����_In   In Ӱ�񱨸�����嵥.����%Type,
    ˵��_In   In Ӱ�񱨸�����嵥.˵��%Type,
    ����_In   In Ӱ�񱨸�����嵥.����%Type,
    ����_In   In Ӱ�񱨸�����嵥.����%Type,
    ���_In   In Ӱ�񱨸�����嵥.���%Type,
    �༭��_In In Ӱ�񱨸�����嵥.�༭��%Type
	) As
  Begin
    Insert Into Ӱ�񱨸�����嵥
      (ID, ����, ����, ˵��, ����, ����, ���, �༭��, ���༭ʱ��)
    Values
      (ID_In, ����_In, ����_In, ˵��_In, ����_In, ����_In, ���_In, �༭��_In, Sysdate);
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_AddComboInfo;

  --3.��  ��;�޸�Ӱ�񱨸������Ϣ
  Procedure p_EditComboInfo(
    ID_In     In Ӱ�񱨸�����嵥.ID%Type,
    ����_In   In Ӱ�񱨸�����嵥.����%Type,
    ����_In   In Ӱ�񱨸�����嵥.����%Type,
    ˵��_In   In Ӱ�񱨸�����嵥.˵��%Type,
    ����_In   In Ӱ�񱨸�����嵥.����%Type,
    ����_In   In Ӱ�񱨸�����嵥.����%Type,
    ���_In   In Ӱ�񱨸�����嵥.���%Type,
    �༭��_In In Ӱ�񱨸�����嵥.�༭��%Type
	) As
  Begin
    Update Ӱ�񱨸�����嵥
       set ����         = ����_In,
           ����         = ����_In,
           ˵��         = ˵��_In,
           ����         = ����_In,
           ����         = ����_In,
           ���         = ���_In,
           �༭��       = �༭��_In,
           ���༭ʱ�� = SysDate
     Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_EditComboInfo;

  --4.��  �ܣ�ͨ��IDɾ��Ӱ�񱨸������Ϣ
  Procedure p_DelComboInfo(
    ID_In In Ӱ�񱨸�����嵥.ID%Type
	) As
  Begin
    Delete From Ӱ�񱨸�����嵥 Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_DelComboInfo;

  --5.��  �ܣ�����ID���Ӱ�񱨸������Ϣ
  Procedure p_GetComboInfoByID(
	Val           Out t_Refcur,
	ID_In In Ӱ�񱨸�����嵥.ID%Type
	) As
  Begin
    Open Val For
      Select RawToHex(ID) ID,
             ����,
             ����,
             ˵��,
             ����,
             ����,
             (Nvl(t.���, XmlType('<NULL/>'))).GetClobVal() As ���,
             �༭��,
             ���༭ʱ��
        From Ӱ�񱨸�����嵥 t
       Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetComboInfoByID;

  --6.��  �ܣ����Ӱ�񱨸��������з�����Ϣ
  Procedure p_GetComboAllGroup(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select Distinct ���� From Ӱ�񱨸�����嵥;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetComboAllGroup;

  --7.��  �ܣ����ID��Ӧ��Ӱ�񱨸����Ķ�����Ϣ
  Procedure p_GetComboContent(
	Val           Out t_Refcur,
	ID_In In Ӱ�񱨸�����嵥.ID%Type
	) As
  Begin
    Open Val For
      Select (Nvl(t.���, XmlType('<NULL/>'))).GetClobVal() As ���
        From Ӱ�񱨸�����嵥 t
       Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetComboContent;

  --8.��  �ܣ�����ID��Ӧ��Ӱ�񱨸����Ķ�����Ϣ
  Procedure p_EditComboContent(
    ID_In   In Ӱ�񱨸�����嵥.ID%Type,
    ���_In In Ӱ�񱨸�����嵥.���%Type
	) As
  Begin
    Update Ӱ�񱨸�����嵥 Set ��� = ���_In Where ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_EditComboContent;

  --9.�� �ܣ���ȡ�༭�˶�Ӧ������޸�Ӱ�񱨸������Ϣ
  Procedure p_GetComboInfoByEditor(
	Val           Out t_Refcur,
	�༭��_In In Ӱ�񱨸�����嵥.�༭��%Type
	) AS
  Begin
    Open Val For
      Select RawToHex(ID) ID, �༭��, ���༭ʱ��
        From Ӱ�񱨸�����嵥 t1
       Where Not Exists (Select 1
                From Ӱ�񱨸�����嵥
               Where ���༭ʱ�� > t1.���༭ʱ��)
         And �༭�� = �༭��_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_GetComboInfoByEditor;

  --10.��  �ܣ�����Ƭ�ε���Ͼ�
  Procedure p_Append_Fragment_Tocombo(
    Text_In In XmlType,
	Id_In   In Ӱ�񱨸�����嵥.ID%Type
	) As
  Begin
    Update Ӱ�񱨸�����嵥 A
       Set a.��� = Appendchildxml(a.���, '/root', Text_In)
     Where a.ID = ID_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Append_Fragment_Tocombo;

  --11.��  �ܣ��޸�Ƭ�ε���Ͼ�
  Procedure p_Update_Combo_Fragment(
    Text_In In XmlType,
    Id_In   In Ӱ�񱨸�����嵥.ID%Type,
    Pid_In  In Varchar2
	) As
  Begin
    Update Ӱ�񱨸�����嵥 A
       Set a.��� = Updatexml(a.���,
                            '/root/sentence[@sid="' || Pid_In || '"]',
                            Text_In)
     Where a.ID = Id_In;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Update_Combo_Fragment;

  --12.��  �ܣ����ݷ���ID��ѯ�ʾ�
  Procedure p_Get_Fragment_By_Typeid(
	Val           Out t_Refcur,
	Id_In In Ӱ�񱨸�����嵥.ID%Type
	) As
  Begin
    Open Val For
      Select RawToHex(ID) ID,
             �ϼ�id,
             ����,
             ����,
             ˵��,
             �ڵ�����,
             (Nvl(a.���, XmlType('<NULL/>'))).GetClobVal() As ���,
             ѧ��,
             ��ǩ,
             �Ƿ�˽��,
             ����,
			 (Nvl(a.��Ӧ����, XmlType('<NULL/>'))).GetClobVal() As ��Ӧ����, 
             ���༭ʱ��
        From Ӱ�񱨸�Ƭ���嵥 A
       Where a.�ϼ�id = Id_In
         And a.�ڵ����� <> 0;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_Fragment_By_Typeid;

  --13.��  �ܣ���ȡ��һ������
  Procedure p_Get_ComboNextCode(
    Val Out t_Refcur
	) As
  Begin
    Open Val For
      Select b_pacs_rptpublic.f_Get_Nextcode('Ӱ�񱨸�����嵥') As ����
        From dual;
  Exception
    When Others Then
      Zl_Errorcenter(Sqlcode, Sqlerrm);
  End p_Get_ComboNextCode;
End b_PACS_RptCombo;
/

--107584:Ƚ����,2017-04-06,����������ϰ�ʱ����ŷ���ʱ���ڹ������
Create Or Replace Procedure Zl_�ٴ������Դ����_Modify
(
  Id_In           �ٴ������Դ����.Id%Type,
  ��Դid_In       �ٴ������Դ����.��Դid%Type,
  �ϰ�ʱ��_In     �ٴ������Դ����.�ϰ�ʱ��%Type,
  �޺���_In       �ٴ������Դ����.�޺���%Type,
  ��Լ��_In       �ٴ������Դ����.��Լ��%Type,
  �Ƿ���ſ���_In �ٴ������Դ����.�Ƿ���ſ���%Type,
  �Ƿ��ʱ��_In   �ٴ������Դ����.�Ƿ��ʱ��%Type,
  ԤԼ����_In     �ٴ������Դ����.ԤԼ����%Type,
  �Ƿ��ռ_In     �ٴ������Դ����.�Ƿ��ռ%Type,
  ���﷽ʽ_In     �ٴ������Դ����.���﷽ʽ%Type,
  ����id_In       �ٴ������Դ����.����id%Type,
  ��Դ����_In     Varchar2 := Null,
  ��Դʱ��_In     Varchar2 := Null,
  ��Դ����_In     Varchar2 := Null,
  ɾ����Դ����_In Integer := 0
) As
  --��Դʱ��_IN:���,��ʼʱ��(HH:MM:SS),��ֹʱ(HH:MM:SS)��,����,�Ƿ�ԤԼ|...
  --��Դ����_IN:����id1,����id2,....
  --��Դ����_IN:����,����,����,���Ʒ�ʽ,���,����|
  --ɾ����Դ����_in:1-��������ǰ����ɾ����Դ����,0-��ɾ�����ݣ�ֱ�Ӳ���,-1-��ɾ����Դ����,����������
  v_Err_Msg Varchar2(255);
  Err_Item Exception;
  l_����id t_Numlist := t_Numlist();
  n_Count  Number;

  n_���     �ٴ������Դʱ��.���%Type;
  d_��ʼʱ�� �ٴ������Դʱ��.��ʼʱ��%Type;
  d_��ֹʱ�� �ٴ������Դʱ��.��ֹʱ��%Type;
  n_����     �ٴ������Դʱ��.��������%Type;
  n_�Ƿ�ԤԼ �ٴ������Դʱ��.�Ƿ�ԤԼ%Type;

  n_����     �ٴ������Դ����.����%Type;
  n_����     �ٴ������Դ����.����%Type;
  v_����     �ٴ������Դ����.����%Type;
  n_���Ʒ�ʽ �ٴ������Դ����.���Ʒ�ʽ%Type;
  n_�������� �ٴ������Դ����.����%Type;
Begin
  If Nvl(ɾ����Դ����_In, 0) = 1 Or Nvl(ɾ����Դ����_In, 0) = -1 Then
    Select ID Bulk Collect Into l_����id From �ٴ������Դ���� Where ��Դid = ��Դid_In;
    Forall I In 1 .. l_����id.Count
      Delete �ٴ������Դʱ�� Where ����id = l_����id(I);
  
    Forall I In 1 .. l_����id.Count
      Delete �ٴ������Դ���� Where ����id = l_����id(I);
  
    Forall I In 1 .. l_����id.Count
      Delete �ٴ������Դ���� Where ����id = l_����id(I);
  
    Delete �ٴ������Դ���� Where ��Դid = ��Դid_In;
    Delete From �ٴ������Դ���� Where ��Դid = ��Դid_In;
  
    If Nvl(ɾ����Դ����_In, 0) = -1 Then
      Return;
    End If;
  End If;

  Select Count(1) Into n_Count From �ٴ������Դ���� Where ID = Id_In;
  If n_Count = 0 Then
    Insert Into �ٴ������Դ����
      (ID, ��Դid, �ϰ�ʱ��, �޺���, ��Լ��, �Ƿ���ſ���, �Ƿ��ʱ��, ԤԼ����, �Ƿ��ռ, ���﷽ʽ, ����id)
    Values
      (Id_In, ��Դid_In, �ϰ�ʱ��_In, �޺���_In, ��Լ��_In, �Ƿ���ſ���_In, �Ƿ��ʱ��_In, ԤԼ����_In, �Ƿ��ռ_In, ���﷽ʽ_In, ����id_In);
  
  End If;

  If ��Դʱ��_In Is Not Null Then
    --�����Դȱʡʱ���
    For c_ʱ��μ� In (Select Rownum As ���, Column_Value As ֵ From Table(f_Str2list(��Դʱ��_In, '|'))) Loop
      n_���     := Null;
      n_����     := Null;
      n_�Ƿ�ԤԼ := Null;
      For c_ʱ��� In (Select Rownum As ���, Column_Value As ֵ From Table(f_Str2list(c_ʱ��μ�.ֵ)) Order By ���) Loop
        If c_ʱ���.��� = 1 Then
          n_��� := To_Number(c_ʱ���.ֵ);
        End If;
      
        If c_ʱ���.��� = 2 Then
          d_��ʼʱ�� := To_Date(c_ʱ���.ֵ, 'yyyy-mm-dd hh24:mi:ss');
        End If;
      
        If c_ʱ���.��� = 3 Then
          d_��ֹʱ�� := To_Date(c_ʱ���.ֵ, 'yyyy-mm-dd hh24:mi:ss');
        End If;
      
        If c_ʱ���.��� = 4 Then
          n_���� := To_Number(c_ʱ���.ֵ);
        End If;
      
        If c_ʱ���.��� = 5 Then
          n_�Ƿ�ԤԼ := To_Number(c_ʱ���.ֵ);
        End If;
      
      End Loop;
    
      If Nvl(n_���, 0) <> 0 Then
        Insert Into �ٴ������Դʱ��
          (����id, ���, ��ʼʱ��, ��ֹʱ��, ��������, �Ƿ�ԤԼ)
        Values
          (Id_In, n_���, d_��ʼʱ��, d_��ֹʱ��, n_����, n_�Ƿ�ԤԼ);
      End If;
    End Loop;
  
  End If;

  --�����Դ��ȱʡ����
  --��Դ����_IN:����,����,����,���Ʒ�ʽ,���,����|
  If ��Դ����_In Is Not Null Then
    For c_ʱ��μ� In (Select Rownum As ���, Column_Value As ֵ From Table(f_Str2list(��Դ����_In, '|'))) Loop
      n_����     := Null;
      n_����     := Null;
      v_����     := Null;
      n_���     := Null;
      n_���Ʒ�ʽ := Null;
      n_�������� := Null;
    
      --����,����,����,���Ʒ�ʽ,���,����|
      For c_ʱ��� In (Select Rownum As ���, Column_Value As ֵ From Table(f_Str2list(c_ʱ��μ�.ֵ)) Order By ���) Loop
        If c_ʱ���.��� = 1 Then
          n_���� := To_Number(c_ʱ���.ֵ);
        End If;
      
        If c_ʱ���.��� = 2 Then
          n_���� := To_Number(c_ʱ���.ֵ);
        End If;
      
        If c_ʱ���.��� = 3 Then
          v_���� := c_ʱ���.ֵ;
        End If;
      
        If c_ʱ���.��� = 4 Then
          n_���Ʒ�ʽ := To_Number(c_ʱ���.ֵ);
        End If;
      
        If c_ʱ���.��� = 5 Then
          n_��� := To_Number(c_ʱ���.ֵ);
        End If;
      
        If c_ʱ���.��� = 6 Then
          n_�������� := To_Number(c_ʱ���.ֵ);
        End If;
      
      End Loop;
    
      If v_���� Is Not Null Then
        Insert Into �ٴ������Դ����
          (����id, ����, ����, ����, ���, ���Ʒ�ʽ, ����)
        Values
          (Id_In, n_����, n_����, v_����, n_���, n_���Ʒ�ʽ, n_��������);
      
      End If;
    End Loop;
  End If;
  --�����Դ����
  If ��Դ����_In Is Not Null Then
    Insert Into �ٴ������Դ����
      (����id, ����id)
      Select Id_In As ����id, Column_Value As ����id From Table(f_Num2list(��Դ����_In));
  End If;

Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ������Դ����_Modify;
/

--108432:Ƚ����,2017-05-08,�����̶��������ʱ����ȡ����˺�û��ɾ���ɸ���ʱ�������ɵĳ����¼�������������ʱ����ʱ���������
--107584:Ƚ����,2017-04-06,����������ϰ�ʱ����ŷ���ʱ���ڹ������
Create Or Replace Procedure Zl1_Auto_Buildingregisterplan
(
  �Һ�ʱ��_In In Date := Null,
  ��Դid_In   �ٴ������Դ.Id%Type := Null
) As
  -------------------------------------------------------------------------
  --����˵�����Զ������ٴ������¼
  --          1�����ݺ�Դ�Զ�����ԤԼ���ڵ��ٴ������¼;
  --          2��ԤԼ������ȷ��:��ԴԤԼ����-->ԤԼ��ʽ��������ȡ���)-->ϵͳԤԼ����
  --���:�Һ�ʱ��_IN:NULLʱ���Զ�����;����ֻ���ָ�������Ƿ������˳����¼û��
  --    ��Դid_In:NULLʱ�������к�Դ������ֻ����ָ����Դ
  -------------------------------------------------------------------------
  n_ȱʡԤԼ���� �ٴ������Դ.ԤԼ����%Type;
  v_����Ա����   �ٴ����ﰲ��.����Ա����%Type;
  d_�Ǽ�����     �ٴ����ﰲ��.�Ǽ�ʱ��%Type;
  n_����id       �ٴ����ﰲ��.Id%Type;
  n_��Ŀid       �ٴ����ﰲ��.��Ŀid %Type;

  n_��¼id   �ٴ������¼.Id%Type;
  d_��ǰ���� �ٴ������¼.��������%Type;

  n_�Ƿ���� Number(2);
  l_�̶�ʱ�� t_Strlist := t_Strlist();
  n_Count    Number(18);

  n_��ԤԼ���� Number := 0;
  d_��ʼʱ��   �ٴ������¼.��ʼʱ��%Type;
Begin

  Select Max(ԤԼ����) Into n_ȱʡԤԼ���� From ԤԼ��ʽ;
  If Nvl(n_ȱʡԤԼ����, 0) = 0 Then
    n_ȱʡԤԼ���� := To_Number(Nvl(zl_GetSysParameter('�Һ�����ԤԼ����'), '0'));
  End If;
  If Nvl(n_ȱʡԤԼ����, 0) = 0 Then
    n_ȱʡԤԼ���� := 7;
  End If;

  --�԰���Ϊ��λ,�����������Դ����ʱ�䡱��12:00:00-23:59:59�ڼ�ģ��򿪷�ԤԼ����+1��
  n_��ԤԼ���� := Zl_Fun_Getappointmentdays;

  d_��ǰ����   := Trunc(Nvl(�Һ�ʱ��_In, Sysdate));
  d_�Ǽ�����   := Sysdate;
  v_����Ա���� := Zl_Username;

  --��һ��ѭ������Դ��Ϣ
  For c_��Դ In (Select c.Id, c.����, c.����, c.��Ŀid, c.����id, c.ҽ������,
                      Decode(Nvl(c.ԤԼ����, 0), 0, n_ȱʡԤԼ����, c.ԤԼ����) + n_��ԤԼ���� As ԤԼ����, Nvl(b.վ��, '-') As վ��,
                      Nvl(c.�Ƿ���ջ���, 0) As �Ƿ���ջ���, Nvl(c.���տ���״̬, 0) As ���տ���״̬, Nvl(c.�Ű෽ʽ, 0) As �Ű෽ʽ
               From �ٴ������Դ C, ���ű� B, ��Ա�� A, �շ���ĿĿ¼ D
               Where c.����id = b.Id And c.ҽ��id = a.Id(+) And c.��Ŀid = d.Id And Nvl(c.�Ƿ�ɾ��, 0) = 0 And
                     Nvl(c.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(b.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(a.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     Nvl(d.����ʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) = To_Date('3000-01-01', 'yyyy-mm-dd') And
                     (��Դid_In Is Null Or c.Id = ��Դid_In)
                    --
                     And Exists (Select 1
                      From �ٴ����ﰲ�� M, �ٴ������ N
                      Where m.����id = n.Id And m.��Դid = c.Id And Nvl(n.�Ű෽ʽ, 0) = 0 And n.����ʱ�� Is Not Null And
                            m.���ʱ�� Is Not Null And d_��ǰ���� <= m.��ֹʱ��)) Loop
  
    --��鵱ǰ�������ڵİ��ŵ��շ���Ŀ�Ƿ�Ϊ��Դ�е��շ���Ŀ��������ǣ�����º�Դ�е��շ���Ŀ
    Begin
      Select ��Ŀid
      Into n_��Ŀid
      From (Select a.��Ŀid
             From �ٴ����ﰲ�� A, �ٴ������ B
             Where a.����id = b.Id And a.��Դid = c_��Դ.Id And a.���ʱ�� Is Not Null And d_��ǰ���� Between a.��ʼʱ�� And a.��ֹʱ�� And
                   Nvl(b.�Ű෽ʽ, 0) = 0 And b.����ʱ�� Is Not Null
             Order By a.�Ǽ�ʱ�� Desc)
      Where Rownum < 2;
    Exception
      When Others Then
        n_��Ŀid := Null;
    End;
    If Nvl(n_��Ŀid, 0) <> 0 Then
      If Nvl(c_��Դ.��Ŀid, 0) <> n_��Ŀid Then
        Update �ٴ������Դ Set ��Ŀid = n_��Ŀid Where ID = c_��Դ.Id;
        Commit;
      End If;
    End If;
  
    --�ڶ���ѭ������������
    --��ͷһ�쿪ʼ���ɣ�������ȫ��(8:00-7:59)��0:00-7:59û�г����¼
    For c_���� In (Select m.����,
                        Decode(To_Char(m.����, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7',
                                '����', Null) As ����
                 From (Select Trunc(d_��ǰ����) + ���� As ����
                        From (Select Level - 1 As ���� From Dual Connect By Level <= c_��Դ.ԤԼ���� + 1)
                        Where ��Դid_In Is Not Null
                        Union All
                        Select Trunc(d_��ǰ���� - 1) + ���� As ����
                        From (Select Level - 1 As ���� From Dual Connect By Level <= c_��Դ.ԤԼ���� + 2)
                        Where ��Դid_In Is Null) M) Loop
    
      l_�̶�ʱ�� := t_Strlist();
      --��鵱���Ƿ�����/�ܳ������,���ڣ������ɳ����¼
      Select Count(1)
      Into n_Count
      From �ٴ����ﰲ�� A, �ٴ������ B
      Where a.����id = b.Id And a.��Դid = c_��Դ.Id And c_����.���� Between Trunc(a.��ʼʱ��) And Trunc(a.��ֹʱ��) And
            Nvl(b.�Ű෽ʽ, 0) In (1, 2) And Rownum < 2;
    
      --��ǰ��ԴΪ����/���Ű࣬�ҵ�ǰ����֮ǰ���а���/���Ű�ĳ����¼�Ͳ��ٰ��̶��������ɳ����¼��
      If Nvl(n_Count, 0) = 0 And Nvl(c_��Դ.�Ű෽ʽ, 0) <> 0 Then
        Select Count(1)
        Into n_Count
        From �ٴ����ﰲ�� A, �ٴ������ B
        Where a.����id = b.Id And Nvl(b.�Ű෽ʽ, 0) In (1, 2) And a.��Դid = c_��Դ.Id And a.��ʼʱ�� < c_����.���� And Rownum < 2;
      End If;
    
      If Nvl(n_Count, 0) = 0 Then
        If ��Դid_In Is Null Then
          --���ﰲ��,ȡ���Ǽǵ�һ��
          Begin
            Select ����id
            Into n_����id
            From (Select a.Id As ����id
                   From �ٴ����ﰲ�� A, �ٴ������ B
                   Where a.��Դid = c_��Դ.Id And a.����id = b.Id And Nvl(b.�Ű෽ʽ, 0) = 0 And b.����ʱ�� Is Not Null And
                         a.���ʱ�� Is Not Null And c_����.���� Between a.��ʼʱ�� And a.��ֹʱ��
                   Order By a.�Ǽ�ʱ�� Desc)
            Where Rownum < 2;
          Exception
            When Others Then
              n_����id := 0;
          End;
        Else
          --���ָ���˺�ԴID���϶��Ƿ�������������ʱ�����������ɳ����¼�����Ǽǵ�һ���϶��Ǳ��������ģ�
          --ֻ��Ҫ����������ż��ɣ��������������Чʱ�䷶Χ�ڵľͲ�����
          Begin
            Select ����id
            Into n_����id
            From (Select a.Id As ����id, a.��ʼʱ��, a.��ֹʱ��, Row_Number() Over(Order By a.�Ǽ�ʱ�� Desc) As �к�
                   From �ٴ����ﰲ�� A, �ٴ������ B
                   Where a.��Դid = c_��Դ.Id And a.����id = b.Id And Nvl(b.�Ű෽ʽ, 0) = 0 And b.����ʱ�� Is Not Null And
                         a.���ʱ�� Is Not Null And c_����.���� Between ��ʼʱ�� And ��ֹʱ��)
            Where �к� = 1;
          Exception
            When Others Then
              n_����id := 0;
          End;
        End If;
      
        If Nvl(n_����id, 0) <> 0 Then
          If ��Դid_In Is Null Then
            --ȷ�������Ƿ��г����¼
            Select Count(1)
            Into n_Count
            From �ٴ������¼ A
            Where a.��Դid = c_��Դ.Id And a.�������� = c_����.���� And Rownum < 2;
          
            --1.δָ����ԴID�������������ɳ����¼���г����¼�����ڽ����ٴ���
            If Nvl(n_Count, 0) = 0 Then
              --1.1�޳����¼����������
              n_�Ƿ���� := 1;
            Else
              --1.2�г����¼�����ٴ���
              n_�Ƿ���� := 0;
            End If;
          Else
            --2.ָ���˺�ԴID���϶��Ƿ�������������ʱ�����������ɳ����¼
            n_�Ƿ���� := 1;
            --�����г����¼����Ҫ�����´���
            For c_��¼ In (Select a.����id, a.Id As ��¼id, a.��������, a.�ϰ�ʱ��, a.�Ƿ��ʱ��, a.�Ƿ���ſ���
                         From �ٴ������¼ A
                         Where a.��Դid = c_��Դ.Id And a.�������� = c_����.����) Loop
            
              Select Count(1) Into n_Count From ���˹Һż�¼ Where �����¼id = c_��¼.��¼id;
              If Nvl(n_Count, 0) = 0 Then
                --2.2.1���ʱ�β�����ԤԼ�Һ����ݣ���ɾ����������
                Zl_�ٴ������ϰ�ʱ��_Delete(c_��¼.����id, To_Char(c_��¼.��������, 'yyyy-mm-dd'), 1, c_��¼.�ϰ�ʱ��);
              Else
                --2.2.2���ʱ�δ���ԤԼ�Һ����ݣ���ֻ����������¼�İ���ID����
                Update �ٴ������¼ Set ����id = n_����id Where ID = c_��¼.��¼id;
                l_�̶�ʱ��.Extend();
                l_�̶�ʱ��(l_�̶�ʱ��.Count) := c_��¼.�ϰ�ʱ��;
              End If;
            End Loop;
          End If;
        
          --��������Ƿ����
          If n_�Ƿ���� = 1 Then
            Select Count(1) Into n_Count From �ٴ��������� Where ����id = n_����id And ������Ŀ = c_����.����;
            If Nvl(n_Count, 0) = 0 Then
              n_�Ƿ���� := 0;
            End If;
          End If;
        
          If Nvl(n_�Ƿ����, 0) = 0 Then
            --����������ٴ������¼���������ٴ������¼(ʱ���ΪNULL �Ŀռ�¼)
            Insert Into �ٴ������¼
              (ID, ����id, ��Դid, ��������, �Ǽ���, �Ǽ�ʱ��)
              Select �ٴ������¼_Id.Nextval, n_����id, a.Id As ID, c_����.����, v_����Ա����, d_�Ǽ����� As �Ǽ�ʱ��
              From �ٴ������Դ A, �ٴ����ﰲ�� B
              Where a.Id = b.��Դid And b.Id = n_����id
                   --
                    And Not Exists (Select 1 From �ٴ������¼ Where ��Դid = a.Id And �������� = c_����.����);
          Else
            For c_��¼ In (With c_ʱ��� As
                            (Select ʱ���, ��ʼʱ��, ��ֹʱ��, ����, վ��, ȱʡʱ��, ��ǰʱ��
                            From (Select ʱ���, ��ʼʱ��, ��ֹʱ��, ����, վ��, ȱʡʱ��, ��ǰʱ��,
                                          Row_Number() Over(Partition By ʱ��� Order By ʱ���, վ�� Asc, ���� Asc) As ���
                                   From ʱ���
                                   Where Nvl(վ��, c_��Դ.վ��) = c_��Դ.վ�� And Nvl(����, c_��Դ.����) = c_��Դ.����)
                            Where ��� = 1)
                           Select n_����id As ����id, B1.��Դid, c_����.���� As ��������, m.�ϰ�ʱ��, m.Id As ����id,
                                  To_Date(To_Char(c_����.����, 'yyyy-mm-dd ') || To_Char(j.��ʼʱ��, 'hh24:mi:ss'),
                                           'yyyy-mm-dd hh24:mi:ss') As ��ʼʱ��,
                                  To_Date(To_Char(c_����.����, 'yyyy-mm-dd ') || To_Char(j.��ֹʱ��, 'hh24:mi:ss'),
                                          'yyyy-mm-dd hh24:mi:ss') + Case
                                    When j.��ֹʱ�� <= j.��ʼʱ�� Then
                                     1
                                    Else
                                     0
                                  End As ��ֹʱ��, Null As ͣ�￪ʼʱ��, Null As ͣ����ֹʱ��, Null As ͣ��ԭ��,
                                  To_Date(To_Char(c_����.����, 'yyyy-mm-dd ') || To_Char(Nvl(j.ȱʡʱ��, j.��ʼʱ��), 'hh24:mi:ss'),
                                          'yyyy-mm-dd hh24:mi:ss') + Case
                                    When j.ȱʡʱ�� < j.��ʼʱ�� Then
                                     1
                                    Else
                                     0
                                  End As ȱʡԤԼʱ��,
                                  To_Date(To_Char(c_����.����, 'yyyy-mm-dd ') || To_Char(Nvl(j.��ǰʱ��, j.��ʼʱ��), 'hh24:mi:ss'),
                                          'yyyy-mm-dd hh24:mi:ss') + Case
                                    When j.��ʼʱ�� < j.��ǰʱ�� Then
                                     -1
                                    Else
                                     0
                                  End As ��ǰ�Һ�ʱ��, m.�޺���, 0 As �ѹ���, m.��Լ��, 0 As ��Լ��, 0 As �����ѽ���, m.�Ƿ���ſ���, m.�Ƿ��ʱ��, m.ԤԼ����,
                                  m.�Ƿ��ռ, B1.��Ŀid, B1.ҽ��id, B1.ҽ������, Null As ����ҽ��id, Null As ����ҽ������, m.���﷽ʽ, m.����id,
                                  0 As �Ƿ�����, 0 As �Ƿ���ʱ����, v_����Ա���� As ����Ա����, d_�Ǽ����� As �Ǽ�ʱ��, c_����.���� As ������Ŀ
                           From �ٴ����ﰲ�� B1, �ٴ��������� M, c_ʱ��� J
                           Where B1.Id = n_����id And B1.Id = m.����id And m.������Ŀ = c_����.���� And m.�ϰ�ʱ�� = j.ʱ��� And
                                 To_Date(To_Char(c_����.����, 'yyyy-mm-dd ') || To_Char(j.��ʼʱ��, 'hh24:mi:ss'),
                                         'yyyy-mm-dd hh24:mi:ss') >= B1.��ʼʱ��) Loop
              Begin
                Select 1 Into n_Count From Table(l_�̶�ʱ��) Where Column_Value = c_��¼.�ϰ�ʱ��;
              Exception
                When Others Then
                  n_Count := 0;
              End;
            
              If Nvl(n_Count, 0) = 0 Then
                Select �ٴ������¼_Id.Nextval Into n_��¼id From Dual;
                Insert Into �ٴ������¼
                  (ID, ����id, ��Դid, ��������, �ϰ�ʱ��, ��ʼʱ��, ��ֹʱ��, ͣ�￪ʼʱ��, ͣ����ֹʱ��, ͣ��ԭ��, ȱʡԤԼʱ��, ��ǰ�Һ�ʱ��, �޺���, �ѹ���, ��Լ��, ��Լ��,
                   �����ѽ���, �Ƿ���ſ���, �Ƿ��ʱ��, ԤԼ����, �Ƿ��ռ, ��Ŀid, ����id, ҽ��id, ҽ������, ����ҽ��id, ����ҽ������, ���﷽ʽ, ����id, �Ƿ�����, �Ƿ���ʱ����,
                   �Ǽ���, �Ǽ�ʱ��, �Ƿ񷢲�)
                Values
                  (n_��¼id, c_��¼.����id, c_��¼.��Դid, c_��¼.��������, c_��¼.�ϰ�ʱ��, c_��¼.��ʼʱ��, c_��¼.��ֹʱ��, c_��¼.ͣ�￪ʼʱ��, c_��¼.ͣ����ֹʱ��,
                   c_��¼.ͣ��ԭ��, c_��¼.ȱʡԤԼʱ��, c_��¼.��ǰ�Һ�ʱ��, c_��¼.�޺���, c_��¼.�ѹ���, c_��¼.��Լ��, c_��¼.��Լ��, c_��¼.�����ѽ���, c_��¼.�Ƿ���ſ���,
                   c_��¼.�Ƿ��ʱ��, c_��¼.ԤԼ����, c_��¼.�Ƿ��ռ, c_��¼.��Ŀid, c_��Դ.����id, c_��¼.ҽ��id, c_��¼.ҽ������, c_��¼.����ҽ��id,
                   c_��¼.����ҽ������, c_��¼.���﷽ʽ, c_��¼.����id, c_��¼.�Ƿ�����, c_��¼.�Ƿ���ʱ����, c_��¼.����Ա����, d_�Ǽ�����, 1);
              
                d_��ʼʱ�� := c_��¼.��ʼʱ��;
                --�����ٴ�������ſ���
                If Nvl(c_��¼.�Ƿ��ʱ��, 0) = 1 And Nvl(c_��¼.�Ƿ���ſ���, 0) = 1 Then
                  --��ʱ����������ſ��ƣ�ʹ��"ԤԼ˳���"��¼"�Ƿ�ԤԼ"
                  Insert Into �ٴ�������ſ���
                    (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ, ԤԼ˳���)
                    Select n_��¼id, ���,
                           To_Date(To_Char(c_��¼.��������, 'yyyy-mm-dd ') || To_Char(��ʼʱ��, 'hh24:mi:ss'),
                                    'yyyy-mm-dd hh24:mi:ss') + Case
                              When d_��ʼʱ�� > To_Date(To_Char(c_��¼.��������, 'yyyy-mm-dd ') || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') Then
                               1
                              Else
                               0
                            End,
                           To_Date(To_Char(c_��¼.��������, 'yyyy-mm-dd ') || To_Char(��ֹʱ��, 'hh24:mi:ss'),
                                    'yyyy-mm-dd hh24:mi:ss') + Case
                              When d_��ʼʱ�� >= To_Date(To_Char(c_��¼.��������, 'yyyy-mm-dd ') || To_Char(��ֹʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd hh24:mi:ss') Then
                               1
                              Else
                               0
                            End, ��������, �Ƿ�ԤԼ, �Ƿ�ԤԼ
                    From �ٴ�����ʱ��
                    Where ����id = c_��¼.����id;
                Else
                  Insert Into �ٴ�������ſ���
                    (��¼id, ���, ��ʼʱ��, ��ֹʱ��, ����, �Ƿ�ԤԼ)
                    Select n_��¼id, ���,
                           To_Date(To_Char(c_��¼.��������, 'yyyy-mm-dd ') || To_Char(��ʼʱ��, 'hh24:mi:ss'),
                                   'yyyy-mm-dd hh24:mi:ss') + Case
                             When d_��ʼʱ�� > To_Date(To_Char(c_��¼.��������, 'yyyy-mm-dd ') || To_Char(��ʼʱ��, 'hh24:mi:ss'),
                                                   'yyyy-mm-dd hh24:mi:ss') Then
                              1
                             Else
                              0
                           End,
                           To_Date(To_Char(c_��¼.��������, 'yyyy-mm-dd ') || To_Char(��ֹʱ��, 'hh24:mi:ss'),
                                   'yyyy-mm-dd hh24:mi:ss') + Case
                             When d_��ʼʱ�� >= To_Date(To_Char(c_��¼.��������, 'yyyy-mm-dd ') || To_Char(��ֹʱ��, 'hh24:mi:ss'),
                                                    'yyyy-mm-dd hh24:mi:ss') Then
                              1
                             Else
                              0
                           End, ��������, �Ƿ�ԤԼ
                    From �ٴ�����ʱ��
                    Where ����id = c_��¼.����id;
                End If;
              
                --���������λ�Һſ��Ƽ�¼
                Insert Into �ٴ�����Һſ��Ƽ�¼
                  (����, ����, ����, ��¼id, ���, ���Ʒ�ʽ, ����)
                  Select ����, ����, ����, n_��¼id, ���, ���Ʒ�ʽ, ����
                  From �ٴ�����Һſ���
                  Where ����id = c_��¼.����id;
              
                --�����ٴ��������Ҽ�¼
                Insert Into �ٴ��������Ҽ�¼
                  (��¼id, ����id)
                  Select n_��¼id, ����id From �ٴ��������� Where ����id = c_��¼.����id;
              End If;
            End Loop;
          
            --����ͣ�ﰲ�źͷ����ڼ��յ��������¼�ĳ���/ԤԼ���
            Zl_Clinicvisitmodify(c_��Դ.Id, n_����id, c_����.����, v_����Ա����, d_�Ǽ�����);
          End If;
        End If;
      End If;
      --һ��һ�ύ
      Commit;
    End Loop;
  End Loop;
End Zl1_Auto_Buildingregisterplan;
/

--105791:Ƚ����,2017-04-19,�������ձ��ֶ�������
Create Or Replace Procedure Zl_�������ձ�_Modify
(
  ��������_In     Number,
  ���_In         �������ձ�.���%Type,
  ��������_In     �������ձ�.��������%Type,
  ��ʼ����_In     �������ձ�.��ʼ����%Type,
  ��ֹ����_In     �������ձ�.��ֹ����%Type,
  ��ע_In         �������ձ�.��ע%Type,
  �������_In     Varchar2 := Null,
  ����ԤԼ����_In �������ձ�.����ԤԼ����%Type,
  ����Һ�����_In �������ձ�.����Һ�����%Type
) As
  --�������޸ķ����ڼ���
  --      ��������_In 0-������1-�޸�
  --      �������_In ��ʽ������ʱ��1~ ԭ�ϰ�ʱ��1;����ʱ��2~ ԭ�ϰ�ʱ��2;
  --      ����ԤԼ����_In ����ԤԼ������,��ʽ��yyyy-mm-dd;yyyy-mm-dd;...
  --      ����Һ�����_In ����Һŵ�����,��ʽ��yyyy-mm-dd;yyyy-mm-dd;...
  v_Err_Msg Varchar2(255);
  Err_Item Exception;

  n_Count Number;

  v_������� Varchar2(4000);
  v_��ǰ��Ŀ Varchar2(4000);
  d_��ʼ���� Date;
  d_��ֹ���� Date;
Begin
  If ��������_In = 0 Then
    --����
    Begin
      Select 1
      Into n_Count
      From �������ձ�
      Where ���� = 0 And ��� = ���_In And �������� = ��������_In And Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count > 0 Then
      v_Err_Msg := ���_In || '���Ѵ��ڡ�' || ��������_In || '����';
      Raise Err_Item;
    End If;
  
    Begin
      Select 1
      Into n_Count
      From �ٴ������¼ A
      Where a.�������� >= ��ʼ����_In And a.�ϰ�ʱ�� Is Not Null And Nvl(a.�Ƿ񷢲�, 0) = 1 And Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If Nvl(n_Count, 0) <> 0 Then
      v_Err_Msg := '��ǰ�ڼ��տ�ʼʱ��֮��������Ч�ĳ��ﰲ�ţ�������������';
      Raise Err_Item;
    End If;
  
    Begin
      Select 1
      Into n_Count
      From �ٴ������¼
      Where �������� Between ��ʼ����_In And ��ֹ����_In And Nvl(�Ƿ񷢲�, 0) = 1 And (Nvl(��Լ��, 0) <> 0 Or Nvl(�ѹ���, 0) <> 0) And
            Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count > 0 Then
      v_Err_Msg := '��ǰ�ڼ��յ�ʱ�䷶Χ������ԤԼ�ҺŲ��ˣ��������ã�';
      Raise Err_Item;
    End If;
  
    Begin
      Select 1
      Into n_Count
      From �������ձ�
      Where ���� = 0 And ��ֹ����_In > ��ʼ���� And ��ʼ����_In < ��ֹ���� And Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count > 0 Then
      v_Err_Msg := '��ǰ�ڼ��յ�ʱ�䷶Χ���Ѵ��������ڼ��գ�';
      Raise Err_Item;
    End If;
  
    Insert Into �������ձ�
      (���, ��������, ����, ��ʼ����, ��ֹ����, ��ע, ����ԤԼ����, ����Һ�����)
    Values
      (���_In, ��������_In, 0, ��ʼ����_In, ��ֹ����_In, ��ע_In, ����ԤԼ����_In, ����Һ�����_In);
  
    If �������_In Is Not Null Then
      v_������� := �������_In || ';';
    End If;
    While v_������� Is Not Null Loop
      v_��ǰ��Ŀ := Substr(v_�������, 0, Instr(v_�������, ';') - 1);
      d_��ʼ���� := To_Date(Substr(v_��ǰ��Ŀ, 0, Instr(v_��ǰ��Ŀ, '~') - 1), 'yyyy-mm-dd');
      d_��ֹ���� := To_Date(Substr(v_��ǰ��Ŀ, Instr(v_��ǰ��Ŀ, '~') + 1), 'yyyy-mm-dd');
    
      Insert Into �������ձ�
        (���, ��������, ����, ��ʼ����, ��ֹ����, ��ע)
      Values
        (���_In, ��������_In, 1, d_��ʼ����, d_��ֹ����, Null);
    
      v_������� := Substr(v_�������, Instr(v_�������, ';') + 1);
    End Loop;
  
  Elsif ��������_In = 1 Then
    --�޸�
    Begin
      Select ��ʼ����
      Into d_��ʼ����
      From �������ձ�
      Where ���� = 0 And ��� = ���_In And �������� = ��������_In And Rownum < 2;
    Exception
      When Others Then
        d_��ʼ���� := Null;
    End;
    If d_��ʼ���� Is Null Then
      v_Err_Msg := ���_In || '�겻���ڡ�' || ��������_In || '����';
      Raise Err_Item;
    End If;
  
    Begin
      Select 1
      Into n_Count
      From �ٴ������¼ A
      Where a.�������� >= d_��ʼ���� And a.�ϰ�ʱ�� Is Not Null And Nvl(a.�Ƿ񷢲�, 0) = 1 And Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If Nvl(n_Count, 0) <> 0 Then
      v_Err_Msg := '��ǰ�ڼ��տ�ʼʱ��֮��������Ч�ĳ��ﰲ�ţ������޸ģ�';
      Raise Err_Item;
    End If;
  
    Begin
      Select 1
      Into n_Count
      From �ٴ������¼
      Where �������� Between ��ʼ����_In And ��ֹ����_In And Nvl(�Ƿ񷢲�, 0) = 1 And (Nvl(��Լ��, 0) <> 0 Or Nvl(�ѹ���, 0) <> 0) And
            Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count > 0 Then
      v_Err_Msg := '��ǰ�ڼ��յ�ʱ�䷶Χ������ԤԼ�ҺŲ��ˣ������޸ģ�';
      Raise Err_Item;
    End If;
  
    Begin
      Select 1
      Into n_Count
      From �������ձ�
      Where ���� = 0 And ��ֹ����_In > ��ʼ���� And ��ʼ����_In < ��ֹ���� And Not (��� = ���_In And �������� = ��������_In) And Rownum < 2;
    Exception
      When Others Then
        n_Count := 0;
    End;
    If n_Count > 0 Then
      v_Err_Msg := '��ǰ�ڼ��յ�ʱ�䷶Χ���Ѵ��������ڼ��գ�';
      Raise Err_Item;
    End If;
  
    Update �������ձ�
    Set ��ʼ���� = ��ʼ����_In, ��ֹ���� = ��ֹ����_In, ��ע = ��ע_In, ����ԤԼ���� = ����ԤԼ����_In, ����Һ����� = ����Һ�����_In
    Where ��� = ���_In And Nvl(����, 0) = 0 And �������� = ��������_In;
  
    --��ɾ����������
    Delete From �������ձ� Where ��� = ���_In And Nvl(����, 0) = 1 And �������� = ��������_In;
    If �������_In Is Not Null Then
      v_������� := �������_In || ';';
    End If;
    While v_������� Is Not Null Loop
      v_��ǰ��Ŀ := Substr(v_�������, 0, Instr(v_�������, ';') - 1);
      d_��ʼ���� := To_Date(Substr(v_��ǰ��Ŀ, 0, Instr(v_��ǰ��Ŀ, '~') - 1), 'yyyy-mm-dd');
      d_��ֹ���� := To_Date(Substr(v_��ǰ��Ŀ, Instr(v_��ǰ��Ŀ, '~') + 1), 'yyyy-mm-dd');
    
      Insert Into �������ձ�
        (���, ��������, ����, ��ʼ����, ��ֹ����, ��ע)
      Values
        (���_In, ��������_In, 1, d_��ʼ����, d_��ֹ����, Null);
    
      v_������� := Substr(v_�������, Instr(v_�������, ';') + 1);
    End Loop;
  End If;
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�������ձ�_Modify;
/

--107559:Ƚ����,2017-04-18,������ֹͣ�ﰲ�Ź���
--105791:Ƚ����,2017-04-19,�������ձ��ֶ�������
Create Or Replace Procedure Zl_Clinicvisitmodify
(
  ��Դid_In     In �ٴ������¼.��Դid%Type,
  ����id_In     In �ٴ������¼.��Դid%Type,
  ��������_In   In �ٴ������¼.��������%Type,
  �Ǽ���_In     In �ٴ������¼.�Ǽ���%Type,
  �Ǽ�ʱ��_In   In �ٴ������¼.�Ǽ�ʱ��%Type,
  �Ƿ��ѻ���_In In Number := 0
) As
  --���ܣ�����ͣ�ﰲ�źͷ����ڼ��յ��������¼�ĳ���/ԤԼ���
  --��Σ�
  --     �Ƿ��ѻ���_In ��Ҫ���ڻ��ݺ����ͣ�ﴦ��
  --˵����
  --     �ٴ������Դ.���տ���״̬��0-���ϰ�;1-�ϰ��ҿ���ԤԼ;2-�ϰ൫������ԤԼ;3-�ܽڼ������ÿ���
  --     1-ͣ���ͣ�ﰲ��ʱ�䷶Χ��
  --     2-ͣ��ڷ����ڼ�����
  --       2.1�ٴ������Դ.���տ���״̬=0
  --       2.2�ٴ������Դ.���տ���״̬=3��������������ԤԼ/����Һţ������ϰ�ʱ�β������õ�����ԤԼ/����Һŵ�ʱ�䷶Χ��
  --     3-��ֹԤԼ���ڷ����ڼ����ڣ�
  --       3.1�ٴ������Դ.���տ���״̬=2
  --       3.3�ٴ������Դ.���տ���״̬=3������������ԤԼ/����Һţ��Ҹ��ϰ�ʱ�������õ�����Һŵ�ʱ�䷶Χ�ڣ����������õ�����ԤԼʱ�䷶Χ��
  --     else-��������

  n_���տ���״̬ �ٴ������Դ.���տ���״̬%Type;
  n_�Ƿ���ջ��� �ٴ������Դ.�Ƿ���ջ���%Type;

  d_ԭ�ϰ����� �ٴ������¼.��������%Type;
  d_��������   �ٴ������¼.��������%Type;

  d_ͣ�￪ʼʱ�� �ٴ������¼.ͣ�￪ʼʱ��%Type;
  d_ͣ����ֹʱ�� �ٴ������¼.ͣ����ֹʱ��%Type;
  v_ͣ��ԭ��     �ٴ������¼.ͣ��ԭ��%Type;

  d_���տ�ʼ���� �������ձ�.��ʼ����%Type;
  d_������ֹ���� �������ձ�.��ֹ����%Type;
  v_����ԤԼ     �������ձ�.����ԤԼ����%Type;
  v_����Һ�     �������ձ�.����Һ�����%Type;

  d_ֹͣԤԼ��ʼʱ�� �ٴ������¼.ͣ�￪ʼʱ��%Type;
  d_ֹͣԤԼ��ֹʱ�� �ٴ������¼.ͣ����ֹʱ��%Type;

  n_Count    Number(2);
  n_����ԤԼ Number(2);
  n_����Һ� Number(2);

  Procedure Stopbespeak
  (
    ��¼id_In   In �ٴ������¼.Id%Type,
    ��ʼʱ��_In In �ٴ������¼.��ʼʱ��%Type,
    ��ֹʱ��_In In �ٴ������¼.��ֹʱ��%Type
  ) As
    --���ܣ���ֹԤԼ
    --˵����
    --      ��ʱ������ſ��Ƶģ��޸�"�ٴ�������ſ���.�Ƿ�ԤԼ"����1��Ϊ0��ȡ������ʱ����"ԤԼ˳���"�ָ�
    --      ��ʱ���Ҳ���ſ��Ƶģ��޸�"�ٴ�������ſ���.�Ƿ�ԤԼ"Ϊ0��ȡ������ʱ���ݻָ�Ϊ1
    --      ����ʱ�εģ��ṩ���������ڹҺ�ԤԼʱ���ԤԼʱ���Ƿ��ڲ�����ԤԼ��ʱ�䷶Χ��
  Begin
    Update �ٴ�������ſ��� Set �Ƿ�ԤԼ = 0 Where ��¼id = ��¼id_In And ��ʼʱ�� Between ��ʼʱ��_In And ��ֹʱ��_In;
  End Stopbespeak;

  Procedure Stopvisit
  (
    ��¼id_In       In �ٴ������¼.Id%Type,
    ͣ�￪ʼʱ��_In In �ٴ������¼.ͣ�￪ʼʱ��%Type,
    ͣ����ֹʱ��_In In �ٴ������¼.ͣ����ֹʱ��%Type,
    ͣ��ԭ��_In     In �ٴ������¼.ͣ��ԭ��%Type
  ) As
    --���ܣ�ͣ��
    --˵����
    --     ͬһ�������¼���Դ��ڶ���ͣ���¼���ٴ������¼��ͣ�￪ʼʱ��Ϊ����ͣ���¼����С��ʼʱ�䣬ͣ����ֹʱ��Ϊ����ͣ���¼�������ֹʱ��

  
    d_ͣ�￪ʼʱ�� �ٴ������¼.ͣ�￪ʼʱ��%Type;
    d_ͣ����ֹʱ�� �ٴ������¼.ͣ����ֹʱ��%Type;
    v_ͣ��ԭ��     �ٴ������¼.ͣ��ԭ��%Type;
  Begin
    If ͣ�￪ʼʱ��_In >= ͣ����ֹʱ��_In Then
      Return;
    End If;
  
    --����ͣ���¼
    Insert Into �ٴ�����ͣ���¼
      (ID, ��¼id, ��ʼʱ��, ��ֹʱ��, ͣ��ԭ��, ������, ����ʱ��, ������, ����ʱ��, �Ǽ���)
      Select �ٴ�����ͣ���¼_Id.Nextval, ID, ͣ�￪ʼʱ��_In, ͣ����ֹʱ��_In, ͣ��ԭ��_In, Nvl(ҽ������, �Ǽ���_In), �Ǽ�ʱ��_In, �Ǽ���_In, �Ǽ�ʱ��_In, �Ǽ���_In
      From �ٴ������¼
      Where ID = ��¼id_In;
  
    Begin
      Select Min(a.��ʼʱ��), Max(a.��ֹʱ��), Max(a.ͣ��ԭ��)
      Into d_ͣ�￪ʼʱ��, d_ͣ����ֹʱ��, v_ͣ��ԭ��
      From �ٴ�����ͣ���¼ A
      Where a.��¼id = ��¼id_In And a.ȡ��ʱ�� Is Null;
    Exception
      When Others Then
        d_ͣ�￪ʼʱ�� := Null;
        d_ͣ����ֹʱ�� := Null;
        v_ͣ��ԭ��     := Null;
    End;
  
    Update �ٴ������¼
    Set ͣ�￪ʼʱ�� = d_ͣ�￪ʼʱ��, ͣ����ֹʱ�� = d_ͣ����ֹʱ��, ͣ��ԭ�� = v_ͣ��ԭ��
    Where ID = ��¼id_In;
  
    --����"�ٴ�������ſ���.�Ƿ�ͣ��"Ϊ1
    Update �ٴ�������ſ��� A
    Set �Ƿ�ͣ�� = 1
    Where ��¼id = ��¼id_In And ��ʼʱ�� Between ͣ�￪ʼʱ��_In And ͣ����ֹʱ��_In And Exists
     (Select 1 From �ٴ������¼ Where ID = a.��¼id And Nvl(�Ƿ���ſ���, 0) = 1 And Nvl(�Ƿ��ʱ��, 0) = 1);
  End Stopvisit;

  Procedure Changedaysoff
  (
    ��Դid_In     In �ٴ������¼.��Դid%Type,
    ����id_In     In �ٴ������¼.����id%Type,
    ��������_In   In �ٴ������¼.��������%Type,
    ԭ�ϰ�����_In In �ٴ������¼.��������%Type,
    ��������_In   In �ٴ������¼.��������%Type
  ) As
    --���ܣ����ݴ���
    n_����id �ٴ������¼.����id%Type;
    l_��¼id t_Numlist := t_Numlist();
    n_Count  Number(2);
  Begin
    --1.ǰ��İ��Ż�������
    If ԭ�ϰ�����_In Is Not Null Then
      --1.1.ǰ�������û�а����򲻴���
      Select Count(1)
      Into n_Count
      From �ٴ������¼
      Where ��Դid = ��Դid_In And �������� = ԭ�ϰ�����_In And Rownum < 2;
    
      If Nvl(n_Count, 0) > 0 Then
        --[1]ɾ���������еİ���
        Select ID Bulk Collect Into l_��¼id From �ٴ������¼ Where ��Դid = ��Դid_In And �������� = ��������_In;
        Zl_�ٴ������¼_Batchdelete(l_��¼id);
      
        --[2]���ư���
        For c_���ݼ�¼ In (Select ID, �Ƿ񷢲� From �ٴ������¼ Where ��Դid = ��Դid_In And �������� = ԭ�ϰ�����_In) Loop
          Zl_�ٴ������¼_Copy(c_���ݼ�¼.Id, ����id_In, ��������_In, �Ǽ���_In, �Ǽ�ʱ��_In, c_���ݼ�¼.�Ƿ񷢲�);
        End Loop;
      
        --[3]���¶Խ��ս���ͣ�ﰲ�źͷ����ڼ��յ���
        For c_��¼ In (Select ID From �ٴ������¼ Where ��Դid = ��Դid_In And �������� = ��������_In) Loop
          Zl_Clinicvisitmodify(��Դid_In, ����id_In, ��������_In, �Ǽ���_In, �Ǽ�ʱ��_In, 1);
        End Loop;
      End If;
    End If;
  
    --2.���յİ��Ż���ǰ��
    If ��������_In Is Not Null Then
      --2.1.����û�а����򲻴���
      Select Count(1)
      Into n_Count
      From �ٴ������¼
      Where ��Դid = ��Դid_In And �������� = ��������_In And Rownum < 2;
    
      If Nvl(n_Count, 0) > 0 Then
        --2.2.ǰ����һ��İ����Ѵ���ԤԼ�Һż�¼���滻(��©��)
        Select Count(1)
        Into n_Count
        From �ٴ������¼ A, ���˹Һż�¼ B
        Where a.Id = b.�����¼id And a.��Դid = ��Դid_In And a.�������� = ��������_In And Rownum < 2;
      
        If Nvl(n_Count, 0) = 0 Then
          --[1]��¼ǰ����һ���ԭ����ID,û�оͲ�����
          Begin
            Select ID
            Into n_����id
            From (Select Rownum As Rn, ID
                   From �ٴ����ﰲ��
                   Where ��Դid = ��Դid_In And ��������_In Between ��ʼʱ�� And ��ֹʱ�� And ���ʱ�� Is Not Null
                   Order By �Ǽ�ʱ�� Desc)
            Where Rn < 2;
          Exception
            When Others Then
              n_����id := 0;
          End;
        
          If Nvl(n_����id, 0) <> 0 Then
            --[2]ɾ��ǰ����һ�����еİ���
            Select ID Bulk Collect Into l_��¼id From �ٴ������¼ Where ��Դid = ��Դid_In And �������� = ��������_In;
            Zl_�ٴ������¼_Batchdelete(l_��¼id);
          
            --[3]���ư���
            For c_���ݼ�¼ In (Select ID From �ٴ������¼ Where ��Դid = ��Դid_In And �������� = ��������_In) Loop
              --�϶��Ƿ����˵�
              Zl_�ٴ������¼_Copy(c_���ݼ�¼.Id, n_����id, ��������_In, �Ǽ���_In, �Ǽ�ʱ��_In, 1);
            
            End Loop;
          
            --[4]���¶�ǰ����һ�����ͣ�ﰲ�źͷ����ڼ��յ���
            For c_��¼ In (Select ID From �ٴ������¼ Where ��Դid = ��Դid_In And �������� = ��������_In) Loop
              Zl_Clinicvisitmodify(��Դid_In, ����id_In, ��������_In, �Ǽ���_In, �Ǽ�ʱ��_In, 1);
            End Loop;
          End If;
        End If;
      End If;
    End If;
  End Changedaysoff;
Begin
  Begin
    Select Nvl(b.���տ���״̬, 0), Nvl(b.�Ƿ���ջ���, 0)
    Into n_���տ���״̬, n_�Ƿ���ջ���
    From �ٴ������Դ B
    Where b.Id = ��Դid_In;
  Exception
    When Others Then
      --û���ҵ���Դ��ֱ���˳�
      Return;
  End;

  --================================================================================
  --��1�����ջ��ݴ���
  --˵����ֻ���ú����������ǰ���飬��Ϊ��������ڿ��ܻ�û���ƶ�����
  --================================================================================
  If Nvl(�Ƿ��ѻ���_In, 0) = 0 Then
    --ȷ�������ڼ����Ƿ���Ҫ����
    If Nvl(n_�Ƿ���ջ���, 0) = 1 Then
      --1.ǰ��İ��Ż�������
      Begin
        --��ʼ���ڣ�ԭ����Ϣ��(��������) �� ��ֹ���ڣ�ԭ���ϰ���(����������)
        Select a.��ֹ����
        Into d_ԭ�ϰ�����
        From �������ձ� A
        Where a.���� = 1 And ��������_In = a.��ʼ���� And a.��ֹ���� < ��������_In And Rownum < 2;
      Exception
        When Others Then
          d_ԭ�ϰ����� := Null;
      End;
    
      --2.���յİ��Ż���ǰ��
      Begin
        --��ʼ���ڣ�ԭ����Ϣ��(��������) �� ��ֹ���ڣ�ԭ���ϰ���(����������)
        Select a.��ʼ����
        Into d_��������
        From �������ձ� A
        Where a.���� = 1 And ��������_In = a.��ֹ���� And a.��ʼ���� < ��������_In And Rownum < 2;
      Exception
        When Others Then
          d_�������� := Null;
      End;
    
      Changedaysoff(��Դid_In, ����id_In, ��������_In, d_ԭ�ϰ�����, d_��������);
    End If;
  End If;

  For c_��¼ In (Select ID, ��������, ��ʼʱ��, ��ֹʱ��
               From �ٴ������¼
               Where ��Դid = ��Դid_In And �������� = ��������_In And �ϰ�ʱ�� Is Not Null) Loop
    --================================================================================
    --��2��ͣ�ﰲ��ͣ�ﴦ��
    --================================================================================
    For c_ͣ�� In (Select a.��ʼʱ��, Nvl(a.ʧЧʱ��, a.��ֹʱ��) As ��ֹʱ��, a.ͣ��ԭ��
                 From �ٴ�����ͣ���¼ A, �ٴ������¼ B
                 Where a.������ = b.ҽ������ And a.��¼id Is Null And a.����ʱ�� Is Not Null And b.ҽ��id Is Not Null And
                       b.Id = c_��¼.Id And c_��¼.��ʼʱ�� < Nvl(a.ʧЧʱ��, a.��ֹʱ��) And c_��¼.��ֹʱ�� > a.��ʼʱ��
                 Order By a.����ʱ��) Loop
    
      d_ͣ�￪ʼʱ�� := c_ͣ��.��ʼʱ��;
      d_ͣ����ֹʱ�� := c_ͣ��.��ֹʱ��;
      If d_ͣ�￪ʼʱ�� < c_��¼.��ʼʱ�� Then
        d_ͣ�￪ʼʱ�� := c_��¼.��ʼʱ��;
      End If;
      If d_ͣ����ֹʱ�� > c_��¼.��ֹʱ�� Then
        d_ͣ����ֹʱ�� := c_��¼.��ֹʱ��;
      End If;
      Stopvisit(c_��¼.Id, d_ͣ�￪ʼʱ��, d_ͣ����ֹʱ��, c_ͣ��.ͣ��ԭ��);
    End Loop;
  
    --================================================================================
    --��3�������ڼ���ͣ�Ｐ��ֹԤԼ����
    --================================================================================
    --1.���Һ����ϰ�ʱ��ʱ��Ľڼ��գ��Ե�һ��Ϊ׼����ʼʱ���������򣩣�һ��Ҳֻ��һ��
    Begin
      Select ��ʼ����, ��ֹ����, ��������, ����ԤԼ����, ����Һ�����
      Into d_���տ�ʼ����, d_������ֹ����, v_ͣ��ԭ��, v_����ԤԼ, v_����Һ�
      From (Select a.��ʼ����, a.��ֹ����, a.��������, a.����ԤԼ����, a.����Һ�����
             From �������ձ� A
             Where a.���� = 0 And c_��¼.��ʼʱ�� < a.��ֹ���� And c_��¼.��ֹʱ�� > a.��ʼ����
             Order By a.��ʼ����)
      Where Rownum < 2;
    Exception
      When Others Then
        d_���տ�ʼ���� := Null;
        d_������ֹ���� := Null;
        v_ͣ��ԭ��     := Null;
        v_����ԤԼ     := Null;
        v_����Һ�     := Null;
    End;
  
    If v_ͣ��ԭ�� Is Not Null Then
      --���տ���״̬:0-���ϰ�;1-�ϰ��ҿ���ԤԼ;2-�ϰ൫������ԤԼ;3-�ܽڼ������ÿ���
      If Nvl(n_���տ���״̬, 0) = 0 Then
        --���ϰ࣬ͣ��
        d_ͣ�￪ʼʱ�� := d_���տ�ʼ����;
        d_ͣ����ֹʱ�� := d_������ֹ���� + 1 - 1 / 24 / 60 / 60;
        If d_ͣ�￪ʼʱ�� < c_��¼.��ʼʱ�� Then
          d_ͣ�￪ʼʱ�� := c_��¼.��ʼʱ��;
        End If;
        If d_ͣ����ֹʱ�� > c_��¼.��ֹʱ�� Then
          d_ͣ����ֹʱ�� := c_��¼.��ֹʱ��;
        End If;
        Stopvisit(c_��¼.Id, d_ͣ�￪ʼʱ��, d_ͣ����ֹʱ��, v_ͣ��ԭ��);
      Elsif Nvl(n_���տ���״̬, 0) = 2 Then
        --����Һţ�����ֹԤԼ
        d_ֹͣԤԼ��ʼʱ�� := d_���տ�ʼ����;
        d_ֹͣԤԼ��ֹʱ�� := d_������ֹ���� + 1 - 1 / 24 / 60 / 60;
        If d_ֹͣԤԼ��ʼʱ�� < c_��¼.��ʼʱ�� Then
          d_ֹͣԤԼ��ʼʱ�� := c_��¼.��ʼʱ��;
        End If;
        If d_ֹͣԤԼ��ֹʱ�� > c_��¼.��ֹʱ�� Then
          d_ֹͣԤԼ��ֹʱ�� := c_��¼.��ֹʱ��;
        End If;
        Stopbespeak(c_��¼.Id, d_ֹͣԤԼ��ʼʱ��, d_ֹͣԤԼ��ֹʱ��);
      Elsif Nvl(n_���տ���״̬, 0) = 3 Then
        --û��"����Һ�"�ľ�һ��û��"����ԤԼ"��
        If v_����Һ� Is Not Null Then
          --2.����Ƿ��а����ϰ�ʱ��ʱ���"����Һ�"
          --��Ϊ�ϰ�ʱ�����24Сʱ�����Բ���Ľ��������죬��������һ����������
          n_����Һ� := 0;
          For c_����Һ� In (With ��ʱ�� As
                            (Select To_Date(Column_Value, 'yyyy-mm-dd') As ��ʼʱ��,
                                   To_Date(Column_Value, 'yyyy-mm-dd') + 1 - 1 / 24 / 60 / 60 As ��ֹʱ��
                            From Table(f_Str2list(v_����Һ�, ';'))
                            Where c_��¼.��ʼʱ�� < To_Date(Column_Value, 'yyyy-mm-dd') + 1 - 1 / 24 / 60 / 60 And
                                  c_��¼.��ֹʱ�� > To_Date(Column_Value, 'yyyy-mm-dd')
                            Order By To_Date(Column_Value, 'yyyy-mm-dd'))
                           Select a.��ʼʱ��, Nvl(b.��ֹʱ��, a.��ֹʱ��) As ��ֹʱ��
                           From ��ʱ�� A, ��ʱ�� B
                           Where a.��ֹʱ�� = b.��ʼʱ��(+) - 1 / 24 / 60 / 60 And Rownum < 2) Loop
          
            n_����Һ� := 1;
            n_����ԤԼ := 0;
            --3.����Ƿ��а����ϰ�ʱ��ʱ���"����ԤԼ"
            For c_����ԤԼ In (With ��ʱ�� As
                              (Select To_Date(Column_Value, 'yyyy-mm-dd') As ��ʼʱ��,
                                     To_Date(Column_Value, 'yyyy-mm-dd') + 1 - 1 / 24 / 60 / 60 As ��ֹʱ��
                              From Table(f_Str2list(v_����ԤԼ, ';'))
                              Where c_��¼.��ʼʱ�� < To_Date(Column_Value, 'yyyy-mm-dd') + 1 - 1 / 24 / 60 / 60 And
                                    c_��¼.��ֹʱ�� > To_Date(Column_Value, 'yyyy-mm-dd')
                              Order By To_Date(Column_Value, 'yyyy-mm-dd'))
                             Select a.��ʼʱ��, Nvl(b.��ֹʱ��, a.��ֹʱ��) As ��ֹʱ��
                             From ��ʱ�� A, ��ʱ�� B
                             Where a.��ֹʱ�� = b.��ʼʱ��(+) - 1 / 24 / 60 / 60 And Rownum < 2) Loop
            
              n_����ԤԼ := 1;
              --��"����Һ�","����ԤԼ"ʱ�䷶Χ�ڵĲ���Ҫ����
            
              --���ǰ���Ƿ���Ҫ��ֹԤԼ
              If c_��¼.��ʼʱ�� < c_����ԤԼ.��ʼʱ�� And c_����Һ�.��ʼʱ�� < c_����ԤԼ.��ʼʱ�� Then
                If c_��¼.��ʼʱ�� < c_����Һ�.��ʼʱ�� Then
                  d_ֹͣԤԼ��ʼʱ�� := c_����Һ�.��ʼʱ��;
                Else
                  d_ֹͣԤԼ��ʼʱ�� := c_��¼.��ʼʱ��;
                End If;
                d_ֹͣԤԼ��ֹʱ�� := c_����ԤԼ.��ʼʱ��;
                Stopbespeak(c_��¼.Id, d_ֹͣԤԼ��ʼʱ��, d_ֹͣԤԼ��ֹʱ��);
              End If;
            
              If c_��¼.��ֹʱ�� > c_����ԤԼ.��ֹʱ�� And c_����Һ�.��ֹʱ�� > c_����ԤԼ.��ֹʱ�� Then
                d_ֹͣԤԼ��ʼʱ�� := c_����ԤԼ.��ֹʱ��;
                If c_��¼.��ֹʱ�� > c_����Һ�.��ֹʱ�� Then
                  d_ֹͣԤԼ��ʼʱ�� := c_����Һ�.��ֹʱ��;
                Else
                  d_ֹͣԤԼ��ʼʱ�� := c_��¼.��ֹʱ��;
                End If;
                Stopbespeak(c_��¼.Id, d_ֹͣԤԼ��ʼʱ��, d_ֹͣԤԼ��ֹʱ��);
              End If;
            End Loop;
          
            --����Һţ�����ֹԤԼ
            If Nvl(n_����ԤԼ, 0) = 0 Then
              d_ֹͣԤԼ��ʼʱ�� := c_����Һ�.��ʼʱ��;
              d_ֹͣԤԼ��ֹʱ�� := c_����Һ�.��ֹʱ��;
              If d_ֹͣԤԼ��ʼʱ�� < c_��¼.��ʼʱ�� Then
                d_ֹͣԤԼ��ʼʱ�� := c_��¼.��ʼʱ��;
              End If;
              If d_ֹͣԤԼ��ֹʱ�� > c_��¼.��ֹʱ�� Then
                d_ֹͣԤԼ��ֹʱ�� := c_��¼.��ֹʱ��;
              End If;
              Stopbespeak(c_��¼.Id, d_ֹͣԤԼ��ʼʱ��, d_ֹͣԤԼ��ֹʱ��);
            End If;
          
            --���ǰ���Ƿ���Ҫͣ��
            If c_��¼.��ʼʱ�� < c_����Һ�.��ʼʱ�� And d_���տ�ʼ���� < c_����Һ�.��ʼʱ�� Then
              If c_��¼.��ʼʱ�� < d_���տ�ʼ���� Then
                d_ͣ�￪ʼʱ�� := d_���տ�ʼ����;
              Else
                d_ͣ�￪ʼʱ�� := c_��¼.��ʼʱ��;
              End If;
              d_ͣ����ֹʱ�� := c_����Һ�.��ʼʱ��;
              Stopvisit(c_��¼.Id, d_ͣ�￪ʼʱ��, d_ͣ����ֹʱ��, v_ͣ��ԭ��);
            End If;
          
            If c_��¼.��ֹʱ�� > c_����Һ�.��ֹʱ�� And d_������ֹ���� > c_����Һ�.��ֹʱ�� Then
              d_ͣ�￪ʼʱ�� := c_����Һ�.��ֹʱ��;
              If c_��¼.��ֹʱ�� > d_������ֹ���� Then
                d_ͣ����ֹʱ�� := d_ͣ����ֹʱ��;
              Else
                d_ͣ����ֹʱ�� := c_��¼.��ֹʱ��;
              End If;
              Stopvisit(c_��¼.Id, d_ͣ�￪ʼʱ��, d_ͣ����ֹʱ��, v_ͣ��ԭ��);
            End If;
          End Loop;
        
          --�������õ�����Һ�ʱ�䷶Χ�ڣ�ͣ��
          If Nvl(n_����Һ�, 0) = 0 Then
            d_ͣ�￪ʼʱ�� := d_���տ�ʼ����;
            d_ͣ����ֹʱ�� := d_������ֹ���� + 1 - 1 / 24 / 60 / 60;
            If d_ͣ�￪ʼʱ�� < c_��¼.��ʼʱ�� Then
              d_ͣ�￪ʼʱ�� := c_��¼.��ʼʱ��;
            End If;
            If d_ͣ����ֹʱ�� > c_��¼.��ֹʱ�� Then
              d_ͣ����ֹʱ�� := c_��¼.��ֹʱ��;
            End If;
            Stopvisit(c_��¼.Id, d_ͣ�￪ʼʱ��, d_ͣ����ֹʱ��, v_ͣ��ԭ��);
          End If;
        Else
          --δ��������Һ�/����ԤԼ����ͣ��
          d_ͣ�￪ʼʱ�� := d_���տ�ʼ����;
          d_ͣ����ֹʱ�� := d_������ֹ���� + 1 - 1 / 24 / 60 / 60;
          If d_ͣ�￪ʼʱ�� < c_��¼.��ʼʱ�� Then
            d_ͣ�￪ʼʱ�� := c_��¼.��ʼʱ��;
          End If;
          If d_ͣ����ֹʱ�� > c_��¼.��ֹʱ�� Then
            d_ͣ����ֹʱ�� := c_��¼.��ֹʱ��;
          End If;
          Stopvisit(c_��¼.Id, d_ͣ�￪ʼʱ��, d_ͣ����ֹʱ��, v_ͣ��ԭ��);
        End If;
      End If;
    End If;
  End Loop;
End Zl_Clinicvisitmodify;
/

--105791:Ƚ����,2017-04-19,�������ձ��ֶ�������
Create Or Replace Function Zl_Fun_Get�ٴ�����ԤԼ״̬
(
  ��¼id_In   In �ٴ������¼.Id%Type,
  ԤԼʱ��_In In ���˹Һż�¼.ԤԼʱ��%Type,
  ���_In     �ٴ�������ſ���.���%Type := Null,
  ԤԼ��ʽ_In ԤԼ��ʽ.����%Type := Null,
  ������λ_In �Һź�����λ.����%Type := Null,
  �շ�ԤԼ_In Number := 0
) Return Varchar2 As
  --���ܣ��жϳ����¼��ԤԼʱ���Ƿ��ԤԼ
  --��Σ�
  --���أ�
  --     ��ʽ��ԤԼ״̬|��ʾ��Ϣ���磺"1|ԤԼʱ�䲻�ڵ�ǰ�ϰ�ʱ��ʱ�䷶Χ�ڡ�"
  --     ԤԼ״̬��
  --         0-��ԤԼ
  --         ======================================================
  --         1-����ԤԼ��ԤԼʱ�䲻�ڵ�ǰ�ϰ�ʱ��ʱ�䷶Χ��
  --         2-����ԤԼ����ǰ�ϰ�ʱ�ν�ֹԤԼ
  --         3-����ԤԼ����ǰ�ϰ�ʱ����ԤԼʱ��ʱ��ͣ��
  --         4-����ԤԼ����ǰ�ϰ�ʱ��ʣ���ԤԼ��Ϊ��
  --         ======================================================
  --         5-����ԤԼ����ǰԤԼʱ���ڷ����ڼ���ʱ�䷶Χ�ڣ����ϰ�
  --         6-����ԤԼ����ǰԤԼʱ���ڷ����ڼ���ʱ�䷶Χ�ڣ���ֹԤԼ
  --         7-����ԤԼ����ǰԤԼʱ���ڷ����ڼ��ղ�����ԤԼ��ʱ�䷶Χ��
  --         8-����ԤԼ����ǰԤԼʱ���ڷ����ڼ��ղ�����Һŵ�ʱ�䷶Χ��
  --         9-����ԤԼ����ǰԤԼʱ���ڷ����ڼ���ʱ�䷶Χ�ڣ���ͣ��
  --         ======================================================
  --         10-����ԤԼ����ǰԤԼ��ʽ��ֹԤԼ
  --         11-����ԤԼ����ǰԤԼ��ʽ��ԤԼ������
  --         ======================================================
  --         12-����ԤԼ����ǰ������λ��ֹԤԼ
  --         13-����ԤԼ����ǰ������λ��ԤԼ������
  --         ======================================================
  --         14-����ԤԼ����ǰ��Ž�ֹԤԼ
  --         15-����ԤԼ����ǰ����Ѿ���ʹ��
  --         16-����ԤԼ����ǰ��Ų�����
  --
  n_��Դid         �ٴ������¼.��Դid%Type;
  n_�Ƿ��ʱ��     �ٴ������¼.�Ƿ��ʱ��%Type;
  n_ԤԼ����       �ٴ������¼.ԤԼ����%Type;
  d_ͣ�￪ʼʱ��   �ٴ������¼.ͣ�￪ʼʱ��%Type;
  d_ͣ����ֹʱ��   �ٴ������¼.ͣ����ֹʱ��%Type;
  v_ͣ��ԭ��       �ٴ������¼.ͣ��ԭ��%Type;
  n_��Լ��         �ٴ������¼.��Լ��%Type;
  n_��Լ��         �ٴ������¼.��Լ��%Type;
  n_��ռ           �ٴ������¼.�Ƿ��ռ%Type;
  n_���Ʒ�ʽ       �ٴ�����Һſ��Ƽ�¼.���Ʒ�ʽ%Type;
  n_����           �ٴ�����Һſ��Ƽ�¼.����%Type;
  n_��������       �ٴ�����Һſ��Ƽ�¼.����%Type;
  n_��ſ���       �ٴ������¼.�Ƿ���ſ���%Type;
  v_ԤԼ��ʽ       �ٴ�����Һſ��Ƽ�¼.����%Type;
  n_����           �ٴ�����Һſ��Ƽ�¼.����%Type;
  n_ԤԼ��ʽ��Լ�� �ٴ������¼.��Լ��%Type;
  n_ԤԼ��ʽ��Լ�� �ٴ������¼.��Լ��%Type;
  n_�Һ�״̬       �ٴ�������ſ���.�Һ�״̬%Type;
  n_�Ƿ�ԤԼ       �ٴ�������ſ���.�Ƿ�ԤԼ%Type;

  n_���տ���״̬ �ٴ������Դ.���տ���״̬%Type;

  v_����ԤԼ �������ձ�.����ԤԼ����%Type;
  v_����Һ� �������ձ�.����Һ�����%Type;
  n_Count    Number(2);
  n_��ʹ��   Number(5);
Begin
  Begin
    Select a.��Դid, a.�Ƿ��ʱ��, a.ԤԼ����, a.ͣ�￪ʼʱ��, a.ͣ����ֹʱ��, a.ͣ��ԭ��, Nvl(��Լ��, �޺���), ��Լ��, �Ƿ��ռ, �Ƿ���ſ���
    Into n_��Դid, n_�Ƿ��ʱ��, n_ԤԼ����, d_ͣ�￪ʼʱ��, d_ͣ����ֹʱ��, v_ͣ��ԭ��, n_��Լ��, n_��Լ��, n_��ռ, n_��ſ���
    From �ٴ������¼ A
    Where a.Id = ��¼id_In And ԤԼʱ��_In Between ��ʼʱ�� And ��ֹʱ��;
  Exception
    When Others Then
      Return '1|ԤԼʱ�䲻�ڵ�ǰ�ϰ�ʱ��ʱ�䷶Χ�ڡ�';
  End;

  --ԤԼ��ʽ���
  If ԤԼ��ʽ_In Is Not Null Then
    Begin
      Select ���Ʒ�ʽ
      Into n_���Ʒ�ʽ
      From �ٴ�����Һſ��Ƽ�¼
      Where ���� = 2 And ���� = 1 And ��¼id = ��¼id_In And ���� = ԤԼ��ʽ_In And Rownum < 2;
    Exception
      When Others Then
        Begin
          Select ���Ʒ�ʽ
          Into n_���Ʒ�ʽ
          From �ٴ�����Һſ��Ƽ�¼
          Where ���� = 2 And ���� = 1 And ��¼id = ��¼id_In And Rownum < 2;
        Exception
          When Others Then
            Null;
        End;
    End;
    If n_���Ʒ�ʽ = 0 Then
      Return '10|��ǰԤԼ��ʽ��ֹԤԼ��';
    End If;
    If n_���Ʒ�ʽ = 1 Or n_���Ʒ�ʽ = 2 Then
      Select Nvl(��Լ��, �޺���) Into n_ԤԼ��ʽ��Լ�� From �ٴ������¼ Where ID = ��¼id_In;
      If n_��ռ = 0 Then
        Begin
          Select ����
          Into n_����
          From �ٴ�����Һſ��Ƽ�¼
          Where ���� = 2 And ���� = 1 And ���� = ԤԼ��ʽ_In And ��¼id = ��¼id_In;
        Exception
          When Others Then
            n_���� := Null;
        End;
        If n_���� Is Not Null Then
          If n_���Ʒ�ʽ = 1 Then
            n_ԤԼ��ʽ��Լ�� := Round(n_ԤԼ��ʽ��Լ�� * n_���� / 100);
          Else
            n_ԤԼ��ʽ��Լ�� := n_����;
          End If;
          Select Count(1)
          Into n_ԤԼ��ʽ��Լ��
          From ���˹Һż�¼
          Where �����¼id = ��¼id_In And ��¼״̬ = 1 And ԤԼ��ʽ = ԤԼ��ʽ_In;
          If n_ԤԼ��ʽ��Լ�� >= n_ԤԼ��ʽ��Լ�� Then
            Return '11|��ǰԤԼ��ʽ��ԤԼ�����㡣';
          End If;
        End If;
      Else
        --��������ռ
        Begin
          Select ����
          Into n_����
          From �ٴ�����Һſ��Ƽ�¼
          Where ���� = 2 And ���� = 1 And ���� = ԤԼ��ʽ_In And ��¼id = ��¼id_In;
        Exception
          When Others Then
            n_���� := Null;
        End;
        If n_���� Is Not Null Then
          If n_���Ʒ�ʽ = 1 Then
            n_ԤԼ��ʽ��Լ�� := Round(n_ԤԼ��ʽ��Լ�� * n_���� / 100);
          Else
            n_ԤԼ��ʽ��Լ�� := n_����;
          End If;
          Select Count(1)
          Into n_ԤԼ��ʽ��Լ��
          From ���˹Һż�¼
          Where �����¼id = ��¼id_In And ��¼״̬ = 1 And ԤԼ��ʽ = ԤԼ��ʽ_In;
          If n_ԤԼ��ʽ��Լ�� >= n_ԤԼ��ʽ��Լ�� Then
            Return '11|��ǰԤԼ��ʽ��ԤԼ�����㡣';
          End If;
        Else
          If �շ�ԤԼ_In = 0 Then
            For r_���� In (Select ����, ����, ���� From �ٴ�����Һſ��Ƽ�¼ Where ���� = 1 And ��¼id = ��¼id_In) Loop
              If r_����.���� = 1 Then
                Select Count(1)
                Into n_��ʹ��
                From ���˹Һż�¼
                Where �����¼id = ��¼id_In And ������λ = r_����.���� And ��¼״̬ = 1;
              Else
                Select Count(1)
                Into n_��ʹ��
                From ���˹Һż�¼
                Where �����¼id = ��¼id_In And ԤԼ��ʽ = r_����.���� And ��¼״̬ = 1;
              End If;
              If n_���Ʒ�ʽ = 1 Then
                n_�������� := Nvl(n_��������, 0) + Round(r_����.���� * n_ԤԼ��ʽ��Լ�� / 100) - Nvl(n_��ʹ��, 0);
              Else
                n_�������� := Nvl(n_��������, 0) + r_����.���� - Nvl(n_��ʹ��, 0);
              End If;
            End Loop;
            Select Count(1) Into n_��ʹ�� From ���˹Һż�¼ Where �����¼id = ��¼id_In And ��¼״̬ = 1;
            If n_ԤԼ��ʽ��Լ�� - n_�������� - n_��ʹ�� > 0 Then
              Null;
            Else
              Return '11|��ǰԤԼ��ʽ��ԤԼ�����㡣';
            End If;
          Else
            For r_���� In (Select ����, ����, ����
                         From �ٴ�����Һſ��Ƽ�¼
                         Where ���� = 1 And ���� = 2 And ��¼id = ��¼id_In) Loop
              Select Count(1)
              Into n_��ʹ��
              From ���˹Һż�¼
              Where �����¼id = ��¼id_In And ԤԼ��ʽ = r_����.���� And ��¼״̬ = 1;
              If n_���Ʒ�ʽ = 1 Then
                n_�������� := Nvl(n_��������, 0) + Round(r_����.���� * n_ԤԼ��ʽ��Լ�� / 100) - Nvl(n_��ʹ��, 0);
              Else
                n_�������� := Nvl(n_��������, 0) + r_����.���� - Nvl(n_��ʹ��, 0);
              End If;
            End Loop;
            Select Count(1) Into n_��ʹ�� From ���˹Һż�¼ Where �����¼id = ��¼id_In And ��¼״̬ = 1;
            If n_ԤԼ��ʽ��Լ�� - n_�������� - n_��ʹ�� > 0 Then
              Null;
            Else
              Return '11|��ǰԤԼ��ʽ��ԤԼ�����㡣';
            End If;
          End If;
        End If;
      End If;
    End If;
    If n_���Ʒ�ʽ = 3 Then
      If n_��ſ��� = 1 Then
        If �շ�ԤԼ_In = 0 Then
          Begin
            Select ����, ����, ����
            Into n_ԤԼ��ʽ��Լ��, v_ԤԼ��ʽ, n_����
            From �ٴ�����Һſ��Ƽ�¼
            Where ���� = 1 And ��¼id = ��¼id_In And ��� = ���_In;
          Exception
            When Others Then
              n_ԤԼ��ʽ��Լ�� := Null;
          End;
          If n_ԤԼ��ʽ��Լ�� Is Not Null Then
            If v_ԤԼ��ʽ <> ԤԼ��ʽ_In Or n_���� = 1 Then
              Return '11|��ǰԤԼ��ʽ��ԤԼ�����㡣';
            End If;
            Select Nvl(Max(1), 0)
            Into n_ԤԼ��ʽ��Լ��
            From ���˹Һż�¼
            Where �����¼id = ��¼id_In And ���� = ���_In;
            If n_ԤԼ��ʽ��Լ�� >= n_ԤԼ��ʽ��Լ�� Then
              Return '11|��ǰԤԼ��ʽ��ԤԼ�����㡣';
            End If;
          End If;
        Else
          Begin
            Select ����, ����, ����
            Into n_ԤԼ��ʽ��Լ��, v_ԤԼ��ʽ, n_����
            From �ٴ�����Һſ��Ƽ�¼
            Where ���� = 1 And ���� = 2 And ��¼id = ��¼id_In And ��� = ���_In;
          Exception
            When Others Then
              n_ԤԼ��ʽ��Լ�� := Null;
          End;
          If n_ԤԼ��ʽ��Լ�� Is Not Null Then
            If v_ԤԼ��ʽ <> ԤԼ��ʽ_In Then
              Return '11|��ǰԤԼ��ʽ��ԤԼ�����㡣';
            End If;
            Select Nvl(Max(1), 0)
            Into n_ԤԼ��ʽ��Լ��
            From ���˹Һż�¼
            Where �����¼id = ��¼id_In And ���� = ���_In;
            If n_ԤԼ��ʽ��Լ�� >= n_ԤԼ��ʽ��Լ�� Then
              Return '11|��ǰԤԼ��ʽ��ԤԼ�����㡣';
            End If;
          End If;
        End If;
      Else
        If �շ�ԤԼ_In = 0 Then
          For r_���� In (Select ����, ����, ����
                       From �ٴ�����Һſ��Ƽ�¼
                       Where ���� = 1 And ��¼id = ��¼id_In And ��� = ���_In) Loop
            If r_����.���� <> ԤԼ��ʽ_In Or r_����.���� = 1 Then
              If r_����.���� = 1 Then
                Select Count(1)
                Into n_��ʹ��
                From �ٴ�������ſ��� A, ���˹Һż�¼ B
                Where a.��¼id = ��¼id_In And a.ԤԼ˳��� Is Not Null And Nvl(a.�Һ�״̬, 0) <> 0 And a.��ע = b.���� And
                      b.������λ = r_����.���� And b.��¼״̬ = 1;
              Else
                Select Count(1)
                Into n_��ʹ��
                From �ٴ�������ſ��� A, ���˹Һż�¼ B
                Where a.��¼id = ��¼id_In And a.ԤԼ˳��� Is Not Null And Nvl(a.�Һ�״̬, 0) <> 0 And a.��ע = b.���� And
                      b.ԤԼ��ʽ = r_����.���� And b.��¼״̬ = 1;
              End If;
              n_�������� := Nvl(n_��������, 0) + r_����.���� - Nvl(n_��ʹ��, 0);
            Else
              Select Count(1)
              Into n_ԤԼ��ʽ��Լ��
              From �ٴ�������ſ��� A, ���˹Һż�¼ B
              Where a.��¼id = ��¼id_In And a.ԤԼ˳��� Is Not Null And Nvl(a.�Һ�״̬, 0) <> 0 And a.��ע = b.���� And
                    b.ԤԼ��ʽ = ԤԼ��ʽ_In And b.��¼״̬ = 1;
              If n_ԤԼ��ʽ��Լ�� >= n_ԤԼ��ʽ��Լ�� Then
                Return '11|��ǰԤԼ��ʽ��ԤԼ�����㡣';
              End If;
            End If;
          End Loop;
          Select Count(1)
          Into n_��ʹ��
          From �ٴ�������ſ��� A
          Where a.��¼id = ��¼id_In And a.ԤԼ˳��� Is Not Null And Nvl(a.�Һ�״̬, 0) <> 0 And ��� = ���_In;
          Select Nvl(��Լ��, �޺���) Into n_ԤԼ��ʽ��Լ�� From �ٴ������¼ Where ID = ��¼id_In;
          If n_ԤԼ��ʽ��Լ�� - n_�������� - n_��ʹ�� > 0 Then
            Null;
          Else
            Return '11|��ǰԤԼ��ʽ��ԤԼ�����㡣';
          End If;
        Else
          For r_���� In (Select ����, ����, ����
                       From �ٴ�����Һſ��Ƽ�¼
                       Where ���� = 1 And ���� = 2 And ��¼id = ��¼id_In And ��� = ���_In) Loop
            If r_����.���� <> ԤԼ��ʽ_In Then
              Select Count(1)
              Into n_��ʹ��
              From �ٴ�������ſ��� A, ���˹Һż�¼ B
              Where a.��¼id = ��¼id_In And a.ԤԼ˳��� Is Not Null And Nvl(a.�Һ�״̬, 0) <> 0 And a.��ע = b.���� And
                    b.ԤԼ��ʽ = r_����.���� And b.��¼״̬ = 1;
              n_�������� := Nvl(n_��������, 0) + r_����.���� - Nvl(n_��ʹ��, 0);
            Else
              Select Count(1)
              Into n_ԤԼ��ʽ��Լ��
              From �ٴ�������ſ��� A, ���˹Һż�¼ B
              Where a.��¼id = ��¼id_In And a.ԤԼ˳��� Is Not Null And Nvl(a.�Һ�״̬, 0) <> 0 And a.��ע = b.���� And
                    b.ԤԼ��ʽ = ԤԼ��ʽ_In And b.��¼״̬ = 1;
              If n_ԤԼ��ʽ��Լ�� >= n_ԤԼ��ʽ��Լ�� Then
                Return '11|��ǰԤԼ��ʽ��ԤԼ�����㡣';
              End If;
            End If;
          End Loop;
          Select Count(1)
          Into n_��ʹ��
          From �ٴ�������ſ��� A
          Where a.��¼id = ��¼id_In And a.ԤԼ˳��� Is Not Null And Nvl(a.�Һ�״̬, 0) <> 0 And ��� = ���_In;
          Select Nvl(��Լ��, �޺���) Into n_ԤԼ��ʽ��Լ�� From �ٴ������¼ Where ID = ��¼id_In;
          If n_ԤԼ��ʽ��Լ�� - n_�������� - n_��ʹ�� > 0 Then
            Null;
          Else
            Return '11|��ǰԤԼ��ʽ��ԤԼ�����㡣';
          End If;
        End If;
      End If;
    End If;
  End If;

  --������λ���
  If ������λ_In Is Not Null Then
    Begin
      Select ���Ʒ�ʽ
      Into n_���Ʒ�ʽ
      From �ٴ�����Һſ��Ƽ�¼
      Where ���� = 1 And ���� = 1 And ��¼id = ��¼id_In And ���� = ������λ_In And Rownum < 2;
    Exception
      When Others Then
        Begin
          Select ���Ʒ�ʽ
          Into n_���Ʒ�ʽ
          From �ٴ�����Һſ��Ƽ�¼
          Where ���� = 1 And ���� = 1 And ��¼id = ��¼id_In And Rownum < 2;
        Exception
          When Others Then
            Null;
        End;
    End;
    If n_���Ʒ�ʽ = 0 Then
      Return '12|��ǰ������λ��ֹԤԼ��';
    End If;
    If n_���Ʒ�ʽ = 1 Or n_���Ʒ�ʽ = 2 Then
      Select Nvl(��Լ��, �޺���) Into n_ԤԼ��ʽ��Լ�� From �ٴ������¼ Where ID = ��¼id_In;
      If n_��ռ = 0 Then
        Begin
          Select ����
          Into n_����
          From �ٴ�����Һſ��Ƽ�¼
          Where ���� = 1 And ���� = 1 And ���� = ������λ_In And ��¼id = ��¼id_In;
        Exception
          When Others Then
            n_���� := Null;
        End;
        If n_���� Is Not Null Then
          If n_���Ʒ�ʽ = 1 Then
            n_ԤԼ��ʽ��Լ�� := Round(n_ԤԼ��ʽ��Լ�� * n_���� / 100);
          Else
            n_ԤԼ��ʽ��Լ�� := n_����;
          End If;
          Select Count(1)
          Into n_ԤԼ��ʽ��Լ��
          From ���˹Һż�¼
          Where �����¼id = ��¼id_In And ��¼״̬ = 1 And ������λ = ������λ_In;
          If n_ԤԼ��ʽ��Լ�� >= n_ԤԼ��ʽ��Լ�� Then
            Return '13|��ǰ������λ��ԤԼ�����㡣';
          End If;
        End If;
      Else
        --��������ռ
        Begin
          Select ����
          Into n_����
          From �ٴ�����Һſ��Ƽ�¼
          Where ���� = 1 And ���� = 1 And ���� = ������λ_In And ��¼id = ��¼id_In;
        Exception
          When Others Then
            n_���� := Null;
        End;
        If n_���� Is Not Null Then
          If n_���Ʒ�ʽ = 1 Then
            n_ԤԼ��ʽ��Լ�� := Round(n_ԤԼ��ʽ��Լ�� * n_���� / 100);
          Else
            n_ԤԼ��ʽ��Լ�� := n_����;
          End If;
          Select Count(1)
          Into n_ԤԼ��ʽ��Լ��
          From ���˹Һż�¼
          Where �����¼id = ��¼id_In And ��¼״̬ = 1 And ������λ = ������λ_In;
          If n_ԤԼ��ʽ��Լ�� >= n_ԤԼ��ʽ��Լ�� Then
            Return '13|��ǰ������λ��ԤԼ�����㡣';
          End If;
        Else
          For r_���� In (Select ����, ����, ���� From �ٴ�����Һſ��Ƽ�¼ Where ���� = 1 And ��¼id = ��¼id_In) Loop
            If r_����.���� = 1 Then
              Select Count(1)
              Into n_��ʹ��
              From ���˹Һż�¼
              Where �����¼id = ��¼id_In And ������λ = r_����.���� And ��¼״̬ = 1;
            Else
              Select Count(1)
              Into n_��ʹ��
              From ���˹Һż�¼
              Where �����¼id = ��¼id_In And ԤԼ��ʽ = r_����.���� And ��¼״̬ = 1;
            End If;
            If n_���Ʒ�ʽ = 1 Then
              n_�������� := Nvl(n_��������, 0) + Round(r_����.���� * n_ԤԼ��ʽ��Լ�� / 100) - Nvl(n_��ʹ��, 0);
            Else
              n_�������� := Nvl(n_��������, 0) + r_����.���� - Nvl(n_��ʹ��, 0);
            End If;
          End Loop;
          Select Count(1) Into n_��ʹ�� From ���˹Һż�¼ Where �����¼id = ��¼id_In And ��¼״̬ = 1;
          If n_ԤԼ��ʽ��Լ�� - n_�������� - n_��ʹ�� > 0 Then
            Null;
          Else
            Return '13|��ǰ������λ��ԤԼ�����㡣';
          End If;
        End If;
      End If;
    End If;
    If n_���Ʒ�ʽ = 3 Then
      If n_��ſ��� = 1 Then
        Begin
          Select ����, ����, ����
          Into n_ԤԼ��ʽ��Լ��, v_ԤԼ��ʽ, n_����
          From �ٴ�����Һſ��Ƽ�¼
          Where ���� = 1 And ��¼id = ��¼id_In And ��� = ���_In;
        Exception
          When Others Then
            n_ԤԼ��ʽ��Լ�� := Null;
        End;
        If n_ԤԼ��ʽ��Լ�� Is Not Null Then
          If v_ԤԼ��ʽ <> ������λ_In Or n_���� = 1 Then
            Return '13|��ǰ������λ��ԤԼ�����㡣';
          End If;
          Select Nvl(Max(1), 0)
          Into n_ԤԼ��ʽ��Լ��
          From ���˹Һż�¼
          Where �����¼id = ��¼id_In And ���� = ���_In;
          If n_ԤԼ��ʽ��Լ�� >= n_ԤԼ��ʽ��Լ�� Then
            Return '13|��ǰ������λ��ԤԼ�����㡣';
          End If;
        End If;
      Else
        For r_���� In (Select ����, ����, ����
                     From �ٴ�����Һſ��Ƽ�¼
                     Where ���� = 1 And ��¼id = ��¼id_In And ��� = ���_In) Loop
          If r_����.���� <> ������λ_In Or r_����.���� = 1 Then
            If r_����.���� = 1 Then
              Select Count(1)
              Into n_��ʹ��
              From �ٴ�������ſ��� A, ���˹Һż�¼ B
              Where a.��¼id = ��¼id_In And a.ԤԼ˳��� Is Not Null And Nvl(a.�Һ�״̬, 0) <> 0 And a.��ע = b.���� And
                    b.������λ = r_����.���� And b.��¼״̬ = 1;
            Else
              Select Count(1)
              Into n_��ʹ��
              From �ٴ�������ſ��� A, ���˹Һż�¼ B
              Where a.��¼id = ��¼id_In And a.ԤԼ˳��� Is Not Null And Nvl(a.�Һ�״̬, 0) <> 0 And a.��ע = b.���� And
                    b.ԤԼ��ʽ = r_����.���� And b.��¼״̬ = 1;
            End If;
            n_�������� := Nvl(n_��������, 0) + r_����.���� - Nvl(n_��ʹ��, 0);
          Else
            Select Count(1)
            Into n_ԤԼ��ʽ��Լ��
            From �ٴ�������ſ��� A, ���˹Һż�¼ B
            Where a.��¼id = ��¼id_In And a.ԤԼ˳��� Is Not Null And Nvl(a.�Һ�״̬, 0) <> 0 And a.��ע = b.���� And b.������λ = ������λ_In And
                  b.��¼״̬ = 1;
            If n_ԤԼ��ʽ��Լ�� >= n_ԤԼ��ʽ��Լ�� Then
              Return '13|��ǰ������λ��ԤԼ�����㡣';
            End If;
          End If;
        End Loop;
        Select Count(1)
        Into n_��ʹ��
        From �ٴ�������ſ��� A
        Where a.��¼id = ��¼id_In And a.ԤԼ˳��� Is Not Null And Nvl(a.�Һ�״̬, 0) <> 0 And ��� = ���_In;
        Select Nvl(��Լ��, �޺���) Into n_ԤԼ��ʽ��Լ�� From �ٴ������¼ Where ID = ��¼id_In;
        If n_ԤԼ��ʽ��Լ�� - n_�������� - n_��ʹ�� > 0 Then
          Null;
        Else
          Return '13|��ǰ������λ��ԤԼ�����㡣';
        End If;
      End If;
    End If;
  End If;

  --0-����ԤԼ����;1-�úű��ֹԤԼ;2-����ֹ��������ƽ̨��ԤԼ
  If Nvl(n_ԤԼ����, 0) = 1 Then
    Return '2|��ǰ�ϰ�ʱ�ν�ֹԤԼ��';
  End If;

  If d_ͣ�￪ʼʱ�� Is Not Null And Not (Nvl(n_��ſ���, 0) = 1 And Nvl(n_�Ƿ��ʱ��, 0) = 1) Then
    If ԤԼʱ��_In >= d_ͣ�￪ʼʱ�� And ԤԼʱ��_In <= d_ͣ����ֹʱ�� Then
      Return '3|��ǰ�ϰ�ʱ����ԤԼʱ��ʱ��ͣ�����ԤԼ��';
    End If;
  End If;

  If Nvl(n_��Լ��, 0) > 0 Then
    If Nvl(n_��Լ��, 0) - Nvl(n_��Լ��, 0) <= 0 Then
      Return '4|��ǰ�ϰ�ʱ��ʣ���ԤԼ��Ϊ�㣬���ܼ���ԤԼ��';
    End If;
  End If;

  If Nvl(n_�Ƿ��ʱ��, 0) = 0 Then
    --����ʱ��
    Begin
      Select Nvl(b.���տ���״̬, 0) Into n_���տ���״̬ From �ٴ������Դ B Where b.Id = n_��Դid;
    Exception
      When Others Then
        n_���տ���״̬ := 0;
    End;
  
    --1.���Ұ���ԤԼʱ��Ľڼ���
    Begin
      Select a.����ԤԼ����, a.����Һ�����
      Into v_����ԤԼ, v_����Һ�
      From �������ձ� A
      Where a.���� = 0 And ԤԼʱ��_In Between a.��ʼ���� And a.��ֹ���� + 1 - 1 / 24 / 60 / 60 And Rownum < 2;
    Exception
      When Others Then
        Return '0|����ԤԼ��';
    End;
  
    --���տ���״̬��0-���ϰ�;1-�ϰ��ҿ���ԤԼ;2-�ϰ൫������ԤԼ;3-�ܽڼ������ÿ���
    If Nvl(n_���տ���״̬, 0) = 0 Then
      --���ϰ�Ŀ϶��ǲ���ԤԼ��
      Return '5|��ǰԤԼʱ���ڷ����ڼ���ʱ�䷶Χ�ڣ����ϰࡣ';
    Elsif Nvl(n_���տ���״̬, 0) = 1 Then
      Return '0|����ԤԼ��';
    Elsif Nvl(n_���տ���״̬, 0) = 2 Then
      --�ڽڼ���ʱ�䷶Χ�ڣ�����ԤԼ
      Return '6|��ǰԤԼʱ���ڷ����ڼ���ʱ�䷶Χ�ڣ���ֹԤԼ��';
    Elsif Nvl(n_���տ���״̬, 0) = 3 Then
      --û��"����Һ�"��һ��û��"����ԤԼ"
      If v_����Һ� Is Not Null Then
        --2.����Ƿ��а���ԤԼʱ���"����Һ�"
        Select Max(1)
        Into n_Count
        From Table(f_Str2list(v_����Һ�, ';'))
        Where ԤԼʱ��_In Between To_Date(Column_Value, 'yyyy-mm-dd') And
              To_Date(Column_Value, 'yyyy-mm-dd') + 1 - 1 / 24 / 60 / 60 And Rownum < 2;
      
        If Nvl(n_Count, 0) <> 0 Then
          --3.����Ƿ��а���ԤԼʱ���"����ԤԼ"
          Select Max(1)
          Into n_Count
          From Table(f_Str2list(v_����ԤԼ, ';'))
          Where ԤԼʱ��_In Between To_Date(Column_Value, 'yyyy-mm-dd') And
                To_Date(Column_Value, 'yyyy-mm-dd') + 1 - 1 / 24 / 60 / 60 And Rownum < 2;
        
          If Nvl(n_Count, 0) = 0 Then
            --����"����ԤԼ"ʱ�䷶Χ�ڣ�����ԤԼ
            Return '7|��ǰԤԼʱ���ڷ����ڼ��ղ�����ԤԼ��ʱ�䷶Χ�ڣ�����ԤԼ��';
          Else
            Return '0|����ԤԼ��';
          End If;
        Else
          Return '8|��ǰԤԼʱ���ڷ����ڼ��ղ�����Һŵ�ʱ�䷶Χ�ڣ�����ԤԼ��';
        End If;
      Else
        --û������"����Һ�"/"����ԤԼ"��ʾͣ��϶�����ԤԼ
        Return '9|��ǰԤԼʱ���ڷ����ڼ���ʱ�䷶Χ�ڣ���ͣ�����ԤԼ��';
      End If;
    End If;
  Else
    --��ʱ��
    If Nvl(���_In, 0) <> 0 Then
      Begin
        Select Nvl(�Ƿ�ԤԼ, 0), Nvl(�Һ�״̬, 0)
        Into n_�Ƿ�ԤԼ, n_�Һ�״̬
        From �ٴ�������ſ���
        Where ��¼id = ��¼id_In And ��� = ���_In;
      Exception
        When Others Then
          Return '16|��ǰѡ�����Ų����á�';
      End;
      If n_�Ƿ�ԤԼ = 0 Then
        Return '14|��ǰѡ�����Ž�ֹԤԼ��';
      End If;
      If n_�Һ�״̬ <> 0 Then
        Return '15|��ǰѡ�������Ѿ���ʹ�á�';
      End If;
    End If;
    Return '0|����ԤԼ��';
  End If;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Fun_Get�ٴ�����ԤԼ״̬;
/

--101575:��ΰ��,2017-04-24,���·����Ŀ��Ӧ�ļ̳�ҽ�����ݳ�������

Create Or Replace Procedure Zl_�ٴ�·���汾_Copy
(
  Դ·��id_In     �ٴ�·���汾.·��id%Type,
  Դ�汾��_In     �ٴ�·���汾.�汾��%Type,
  Ŀ��·��id_In   �ٴ�·���汾.·��id%Type,
  Ŀ��汾��_In   �ٴ�·���汾.�汾��%Type,
  Դ��֧id_In     �ٴ�·����֧.Id%Type := Null,
  �Ƿ��֧·��_In Number := Null,
  Ŀ���֧id_In   �ٴ�·����֧.Id%Type := Null
  --���ܣ����Ʋ����µ��ٴ�·���汾
  --������
  --  Դ�汾��_In�����δָ��(0��NULL)����ȡ������Ч�İ汾��
  --  Ŀ�걾��_In�����δָ��(0��NULL)��������µİ汾��
  --  �Ƿ��֧·��_In���༭��֧·��ʱ��������֧����·������·���ṹ,1-�ǣ�0��
  --  Ŀ���֧ID_In:�༭��֧·��ʱ����������֧�Ľṹ��������֧��ID��
) Is
  n_Դ�汾��   �ٴ�·���汾.�汾��%Type;
  n_Ŀ��汾�� �ٴ�·���汾.�汾��%Type;

  n_Advice_New_Id    Number;
  n_Advice_Parent_Id Number;

  n_Step_New_Id    Number;
  n_Step_Parent_Id Number;

  n_Item_New_Id Number;

  n_Eval_New_Id Number;
  n_Eval_Old_Id Number;

  n_Mark_New_Id Number;

  n_Branch_New_Id Number;

  v_Error Varchar2(255);
  Err_Custom Exception;

  n_ǰһ�׶���� �ٴ�·���׶�.���%Type;
  n_��������     �ٴ�·���׶�.��������%Type;
  v_��׼סԺ��   �ٴ�·����֧.��׼סԺ��%Type;
  t_Advice       t_Numlist2 := t_Numlist2(); --����̳е�ҽ��ID C1:=��һ���汾��ҽ��ID;C2:=�����ɵ�ҽ��ID
  --�ٴ�·����֧
  Procedure �ٴ�·����֧_Insert
  (
    Դid_In       Number,
    New_Id_In     Number,
    ·��id_In     Number,
    �汾��_In     Number,
    ����_In       �ٴ�·����֧.����%Type := Null,
    ˵��_In       �ٴ�·����֧.˵��%Type := Null,
    ǰһ�׶�id_In �ٴ�·����֧.ǰһ�׶�id%Type := Null,
    ��׼סԺ��_In �ٴ�·����֧.��׼סԺ��%Type := Null,
    ��׼����_In   �ٴ�·����֧.��׼����%Type := Null
  ) Is
  Begin
    If Nvl(Դid_In, 0) <> 0 Then
      Insert Into �ٴ�·����֧
        (ID, ·��id, �汾��, ����, ˵��, ǰһ�׶�id, ��׼סԺ��, ��׼����, ������, ����ʱ��)
        Select New_Id_In, ·��id_In, �汾��_In, Nvl(����_In, ����), ˵��, ǰһ�׶�id, ��׼סԺ��, ��׼����, Zl_Username, Sysdate
        From �ٴ�·����֧
        Where ID = Դid_In;
    Else
      --����Ǹ�����·�����������׼סԺ�ճ����ˣ��Զ��޸ġ�
      Insert Into �ٴ�·����֧
        (ID, ·��id, �汾��, ����, ˵��, ǰһ�׶�id, ��׼סԺ��, ��׼����, ������, ����ʱ��)
        Select New_Id_In, ·��id_In, �汾��_In, ����_In, ˵��_In, ǰһ�׶�id_In, ��׼סԺ��_In, ��׼����_In, Zl_Username, Sysdate
        From Dual;
    End If;
  
  End;

  --�ٴ�·���׶�
  Procedure �ٴ�·���׶�_Insert
  (
    Դid_In       Number,
    New_Id_In     Number,
    ·��id_In     Number,
    �汾��_In     Number,
    New_��id_In   Number,
    ��֧id_Old_In Number := Null,
    ��֧id_New_In Number := Null
  ) Is
  Begin
    Insert Into �ٴ�·���׶�
      (ID, ·��id, �汾��, ��id, ���, ����, ��ʼ����, ��������, ��־, ˵��, ��֧id)
      Select New_Id_In, ·��id_In, �汾��_In, New_��id_In, ���, ����, ��ʼ����, ��������, ��־, ˵��, ��֧id_New_In
      From �ٴ�·���׶�
      Where ID = Դid_In And Nvl(��֧id, 0) = Nvl(��֧id_Old_In, 0);
  End;
  ---�ٴ�·����Ŀ
  Procedure �ٴ�·����Ŀ_Insert
  (
    Դid_In       Number,
    New_Id_In     Number,
    ·��id_In     Number,
    �汾��_In     Number,
    New_�׶�id_In Number,
    ��֧id_Old_In Number := Null,
    ��֧id_New_In Number := Null
  ) Is
  Begin
    Insert Into �ٴ�·����Ŀ
      (ID, ·��id, �汾��, �׶�id, ����, ��Ŀ���, ��Ŀ����, ����Ҫ��, ִ�з�ʽ, ִ����, ������, ��Ŀ���, ͼ��id, ��֧id)
      Select New_Id_In, ·��id_In, �汾��_In, New_�׶�id_In, ����, ��Ŀ���, ��Ŀ����, ����Ҫ��, ִ�з�ʽ, ִ����, ������, ��Ŀ���, ͼ��id, ��֧id_New_In
      From �ٴ�·����Ŀ
      Where ID = Դid_In And Nvl(��֧id, 0) = Nvl(��֧id_Old_In, 0);
  End;
  --·��ҽ������
  Procedure ·��ҽ������_Insert
  (
    Դid_In       Number,
    New_Id_In     Number,
    New_���id_In Number
  ) Is
  Begin
    Insert Into ·��ҽ������
      (ID, ���id, ���, ��Ч, ������Ŀid, ҽ������, ��������, �ܸ�����, �շ�ϸĿid, �걾��λ, ��鷽��, ִ��Ƶ��, Ƶ�ʴ���, Ƶ�ʼ��, �����λ, ҽ������, ִ������, ִ�п���id, ʱ�䷽��,
       �Ƿ�ȱʡ, �Ƿ�ѡ, �����Ŀid)
      Select New_Id_In, New_���id_In, ���, ��Ч, ������Ŀid, ҽ������, ��������, �ܸ�����, �շ�ϸĿid, �걾��λ, ��鷽��, ִ��Ƶ��, Ƶ�ʴ���, Ƶ�ʼ��, �����λ, ҽ������,
             ִ������, ִ�п���id, ʱ�䷽��, �Ƿ�ȱʡ, �Ƿ�ѡ, �����Ŀid
      From ·��ҽ������
      Where ID = Դid_In;
  End;
  --�ٴ�·��ҽ��
  Procedure �ٴ�·��ҽ��_Inset
  (
    ·����Ŀid_In Number,
    ҽ������id_In Number
  ) Is
  Begin
    Insert Into �ٴ�·��ҽ�� (·����Ŀid, ҽ������id) Values (·����Ŀid_In, ҽ������id_In);
  End;
  --�ٴ�·������
  Procedure �ٴ�·������_Inset
  (
    Դ��Ŀid_In   Number,
    ��Ŀid_New_In Number
  ) Is
  Begin
    Insert Into �ٴ�·������
      (��Ŀid, �ļ�id, ԭ��id, ����, ���)
      Select ��Ŀid_New_In, �ļ�id, ԭ��id, ����, ��� From �ٴ�·������ Where ��Ŀid = Դ��Ŀid_In;
  End;
  ---�ٴ�·������
  Procedure �ٴ�·������_Insert
  (
    Դid_In       Number,
    New_Id_In     Number,
    ·��id_In     Number,
    �汾��_In     Number,
    �׶�id_In     Number,
    ��֧id_Old_In Number := Null,
    ��֧id_New_In Number := Null
  ) Is
  Begin
    Insert Into �ٴ�·������
      (ID, ·��id, �汾��, �׶�id, ��������, ��֧id)
      Select New_Id_In, ·��id_In, �汾��_In, �׶�id_In, ��������, ��֧id_New_In
      From �ٴ�·������
      Where ID = Դid_In And Nvl(��֧id, 0) = Nvl(��֧id_Old_In, 0);
  End;

  Procedure ·������ָ��_Insert
  (
    Դid_In   Number,
    New_Id_In Number,
    ����id_In Number
  ) Is
  Begin
    Insert Into ·������ָ��
      (ID, ����id, ���, ����ָ��, ָ������, ָ����)
      Select New_Id_In, ����id_In, ���, ����ָ��, ָ������, ָ���� From ·������ָ�� Where ID = Դid_In;
  End;
  --·����������
  Procedure ·����������_Insert
  (
    Դ����id_In   Number,
    Դָ��id_In   Number,
    Դ��Ŀid_In   Number,
    New_����id_In Number,
    New_ָ��id_In Number,
    New_��Ŀid_In Number
  ) Is
  Begin
    If Դָ��id_In Is Null Then
      Insert Into ·����������
        (����id, ָ��id, ��Ŀid, ��ϵʽ, ����ֵ, �������)
        Select New_����id_In, New_ָ��id_In, New_��Ŀid_In, ��ϵʽ, ����ֵ, �������
        From ·����������
        Where ����id = Դ����id_In And ָ��id Is Null And ��Ŀid = Դ��Ŀid_In;
    Elsif Դ��Ŀid_In Is Null Then
      Insert Into ·����������
        (����id, ָ��id, ��Ŀid, ��ϵʽ, ����ֵ, �������)
        Select New_����id_In, New_ָ��id_In, New_��Ŀid_In, ��ϵʽ, ����ֵ, �������
        From ·����������
        Where ����id = Դ����id_In And ָ��id = Դָ��id_In And ��Ŀid Is Null;
    End If;
  End;
  --�ٴ�·���׶�
  Procedure �ٴ�·���׶�cascade_Insert
  (
    Դid_In       Number,
    New_Id_In     Number,
    Old·��id_In  Number,
    New·��id_In  Number,
    Old�汾��_In  Number,
    New�汾��_In  Number,
    ��֧id_Old_In Number := Null,
    ��֧id_New_In Number := Null
  ) Is
    n_After   Number(10);
    n_Count   Number(10);
    n_Inherit Number;
    v_Oldid   Varchar2(4000);
    n_Start   Number(10);
    Arr_Id    t_Numlist;
  
  Begin
    ---�ٴ�·������(�׶Σ�ָ��������������
    Select Max(a.Id)
    Into n_Eval_Old_Id
    From �ٴ�·������ A
    Where a.·��id = Old·��id_In And a.�汾�� = Old�汾��_In And a.�׶�id = Դid_In And a.�������� = 2 And
          Nvl(a.��֧id, 0) = Nvl(��֧id_Old_In, 0);
  
    If Nvl(n_Eval_Old_Id, 0) <> 0 Then
      Select �ٴ�·������_Id.Nextval Into n_Eval_New_Id From Dual;
      �ٴ�·������_Insert(n_Eval_Old_Id, n_Eval_New_Id, New·��id_In, New�汾��_In, New_Id_In, ��֧id_Old_In, ��֧id_New_In);
      ---·������ָ��
      For r_·������ָ�� In (Select ID From ·������ָ�� Where ����id = n_Eval_Old_Id) Loop
        Select ·������ָ��_Id.Nextval Into n_Mark_New_Id From Dual;
        ·������ָ��_Insert(r_·������ָ��.Id, n_Mark_New_Id, n_Eval_New_Id);
        ---·����������
        ·����������_Insert(n_Eval_Old_Id, r_·������ָ��.Id, Null, n_Eval_New_Id, n_Mark_New_Id, Null);
      End Loop;
    End If;
    --�ٴ�·����Ŀ
    For r_�ٴ�·����Ŀ In (Select ID
                     From �ٴ�·����Ŀ
                     Where �׶�id = Դid_In And ·��id = Old·��id_In And �汾�� = Old�汾��_In And
                           Nvl(��֧id, 0) = Nvl(��֧id_Old_In, 0)) Loop
    
      Select �ٴ�·����Ŀ_Id.Nextval Into n_Item_New_Id From Dual;
      �ٴ�·����Ŀ_Insert(r_�ٴ�·����Ŀ.Id, n_Item_New_Id, New·��id_In, New�汾��_In, New_Id_In, ��֧id_Old_In, ��֧id_New_In);
      ---�ٴ�·���������׶���������Ŀ������������
      If Nvl(n_Eval_Old_Id, 0) <> 0 Then
        ---·����������
        ·����������_Insert(n_Eval_Old_Id, Null, r_�ٴ�·����Ŀ.Id, n_Eval_New_Id, Null, n_Item_New_Id);
      End If;
      ---�ٴ�·������
      �ٴ�·������_Inset(r_�ٴ�·����Ŀ.Id, n_Item_New_Id);
    
      --·��ҽ������
      For r_�ٴ�·��ҽ�� In (Select b.Id
                       From �ٴ�·��ҽ�� A, ·��ҽ������ B
                       Where a.·����Ŀid = r_�ٴ�·����Ŀ.Id And a.ҽ������id = b.Id And b.���id Is Null) Loop
        --�̳е�ҽ���ж�
        Select Count(1) Into n_Inherit From �ٴ�·��ҽ�� Where ҽ������id = r_�ٴ�·��ҽ��.Id;
        v_Oldid := Null;
        If n_Inherit > 1 Then
          Begin
            Select a.C2 Into v_Oldid From Table(t_Advice) A Where a.C1 = r_�ٴ�·��ҽ��.Id;
          Exception
            When No_Data_Found Then
              v_Oldid := Null;
          End;
        End If;
        If v_Oldid Is Null Then
          ---b.��� > a.��� and b.ID >a.ID --��ȡ��ҽ��ID������ҽ��ID���Ҹ�ҽ����Ŵ�����ҽ����ŵļ�¼��
          Select Count(1)
          Into n_After
          From ·��ҽ������ A
          Where a.���id = r_�ٴ�·��ҽ��.Id And Exists
           (Select 1 From ·��ҽ������ B Where b.Id = r_�ٴ�·��ҽ��.Id And b.��� > a.��� And b.Id > a.Id);
        
          Select Count(1) + 1 Into n_Count From ·��ҽ������ A Where a.���id = r_�ٴ�·��ҽ��.Id;
          Select ·��ҽ������_Id.Nextval Bulk Collect Into Arr_Id From Dual Connect By Rownum <= n_Count;
          If n_After = 0 Then
            n_Advice_Parent_Id := Arr_Id(1);
            n_Start            := 2;
          Else
            n_Advice_Parent_Id := Arr_Id(n_Count);
            n_Start            := 1;
          End If;
        
          ·��ҽ������_Insert(r_�ٴ�·��ҽ��.Id, n_Advice_Parent_Id, Null);
          If n_Inherit > 1 Then
            t_Advice.Extend;
            t_Advice(t_Advice.Count) := t_Numobj2(r_�ٴ�·��ҽ��.Id, n_Advice_Parent_Id);
          End If;
        Else
          n_Advice_Parent_Id := To_Number(v_Oldid);
        End If;
        ---�ٴ�·��ҽ��
        �ٴ�·��ҽ��_Inset(n_Item_New_Id, n_Advice_Parent_Id);
        --·��ҽ��������Ӧ�ӽڵ�
        For r_·��ҽ������ In (Select ID From ·��ҽ������ Where ���id = r_�ٴ�·��ҽ��.Id) Loop
          If v_Oldid Is Null Then
            n_Advice_New_Id := Arr_Id(n_Start);
            n_Start         := n_Start + 1;
          
            ·��ҽ������_Insert(r_·��ҽ������.Id, n_Advice_New_Id, n_Advice_Parent_Id);
            If n_Inherit > 1 Then
              t_Advice.Extend;
              t_Advice(t_Advice.Count) := t_Numobj2(r_·��ҽ������.Id, n_Advice_New_Id);
            End If;
          Else
            --�̳�ҽ����δ�����µ�ID
            If n_Inherit > 1 Then
              Select a.C2 Into n_Advice_New_Id From Table(t_Advice) A Where a.C1 = r_·��ҽ������.Id;
            End If;
          End If;
          ---�ٴ�·��ҽ��
          �ٴ�·��ҽ��_Inset(n_Item_New_Id, n_Advice_New_Id);
        End Loop;
      End Loop;
    End Loop;
  End;
Begin
  --ȷ��Դ·���汾��
  n_Դ�汾�� := Nvl(Դ�汾��_In, 0);
  If n_Դ�汾�� = 0 Then
    Select ���°汾 Into n_Դ�汾�� From �ٴ�·��Ŀ¼ Where ID = Դ·��id_In;
    If Nvl(n_Դ�汾��, 0) = 0 Then
      v_Error := 'Ҫ���Ƶ���Դ�ٴ�·����û�п��õ���Ч�汾��';
      Raise Err_Custom;
    End If;
  End If;

  --ȷ��Ŀ��·���汾��
  n_Ŀ��汾�� := Nvl(Ŀ��汾��_In, 0);
  If n_Ŀ��汾�� = 0 Then
    Select Nvl(Max(�汾��), 0) + 1 Into n_Ŀ��汾�� From �ٴ�·���汾 Where ·��id = Ŀ��·��id_In;
  Else
    If Nvl(�Ƿ��֧·��_In, 0) = 1 Then
      --��������֧����·������ʱ
      --��¼��ǰһ�׶����
      Select Max(a.���)
      Into n_ǰһ�׶����
      From �ٴ�·���׶� A, �ٴ�·����֧ B
      Where a.Id = b.ǰһ�׶�id And b.Id = Nvl(Ŀ���֧id_In, 0);
    
      For r_Ŀ���֧ In (Select * From �ٴ�·����֧ Where ID = Nvl(Ŀ���֧id_In, 0)) Loop
        Zl_�ٴ�·����֧_Delete(Ŀ���֧id_In);
        Select �ٴ�·����֧_Id.Nextval Into n_Branch_New_Id From Dual;
        --��ȷ���Ƿ񳬳���׼סԺ��
        v_��׼סԺ�� := r_Ŀ���֧.��׼סԺ��;
        If Դ��֧id_In = 0 Then
          Select Max(Nvl(��������, ��ʼ����))
          Into n_��������
          From �ٴ�·���׶�
          Where ·��id = Դ·��id_In And �汾�� = n_Ŀ��汾�� And Nvl(��֧id, 0) = Nvl(Դ��֧id_In, 0);
          If Instr(v_��׼סԺ��, '-') > 0 Then
            If Substr(v_��׼סԺ��, Instr(v_��׼סԺ��, '-') + 1) < n_�������� Then
              v_��׼סԺ�� := Substr(v_��׼סԺ��, 1, Instr(v_��׼סԺ��, '-')) || n_��������;
            End If;
          End If;
        End If;
        �ٴ�·����֧_Insert(Դ��֧id_In, n_Branch_New_Id, Ŀ��·��id_In, n_Ŀ��汾��, r_Ŀ���֧.����, r_Ŀ���֧.˵��, r_Ŀ���֧.ǰһ�׶�id, v_��׼סԺ��,
                      r_Ŀ���֧.��׼����);
      End Loop;
    Else
      --������·�����ƻ��������汾��
      Zl_�ٴ�·���汾_Delete(Ŀ��·��id_In, Ŀ��汾��_In);
    End If;
  End If;
  If Nvl(�Ƿ��֧·��_In, 0) <> 1 Then
    --������·�����ƻ��������汾��
    --�ٴ�·���汾
    Insert Into �ٴ�·���汾
      (·��id, �汾��, ��׼סԺ��, ��׼����, �汾˵��, ������, ����ʱ��)
      Select Ŀ��·��id_In, n_Ŀ��汾��, ��׼סԺ��, ��׼����, �汾˵��, Zl_Username, Sysdate
      From �ٴ�·���汾
      Where ·��id = Դ·��id_In And �汾�� = n_Դ�汾��;
    --·����������
    Select Max(ID)
    Into n_Eval_Old_Id
    From �ٴ�·������
    Where ·��id = Դ·��id_In And �汾�� = n_Դ�汾�� And �������� = 1;
    If Nvl(n_Eval_Old_Id, 0) <> 0 Then
      Select �ٴ�·������_Id.Nextval Into n_Eval_New_Id From Dual;
      �ٴ�·������_Insert(n_Eval_Old_Id, n_Eval_New_Id, Ŀ��·��id_In, n_Ŀ��汾��, Null);
      ---·������ָ��
      For r_·������ָ�� In (Select ID From ·������ָ�� Where ����id = n_Eval_Old_Id) Loop
        Select ·������ָ��_Id.Nextval Into n_Mark_New_Id From Dual;
        ·������ָ��_Insert(r_·������ָ��.Id, n_Mark_New_Id, n_Eval_New_Id);
        ---·����������
        ·����������_Insert(n_Eval_Old_Id, r_·������ָ��.Id, Null, n_Eval_New_Id, n_Mark_New_Id, Null);
      End Loop;
    End If;
  Else
    --��������֧����·������ʱ
    Insert Into �ٴ�·������
      (·��id, �汾��, ���, ����, ��֧id)
      Select Ŀ��·��id_In, n_Ŀ��汾��, ���, ����, n_Branch_New_Id
      From �ٴ�·������
      Where ·��id = Դ·��id_In And �汾�� = n_Ŀ��汾�� And Nvl(��֧id, 0) = Nvl(Դ��֧id_In, 0);
  
    For r_�ٴ�·���׶� In (Select ID, ���
                     From �ٴ�·���׶�
                     Where ·��id = Դ·��id_In And �汾�� = n_Ŀ��汾�� And ��id Is Null And Nvl(��֧id, 0) = Nvl(Դ��֧id_In, 0)
                     Order By ���) Loop
      If Nvl(Դ��֧id_In, 0) <> 0 Or r_�ٴ�·���׶�.��� > n_ǰһ�׶���� Then
        --�ٴ�·���׶εĸ����в���
        Select �ٴ�·���׶�_Id.Nextval Into n_Step_Parent_Id From Dual;
        �ٴ�·���׶�_Insert(r_�ٴ�·���׶�.Id, n_Step_Parent_Id, Ŀ��·��id_In, n_Ŀ��汾��, Null, Դ��֧id_In, n_Branch_New_Id);
      
        �ٴ�·���׶�cascade_Insert(r_�ٴ�·���׶�.Id, n_Step_Parent_Id, Դ·��id_In, Ŀ��·��id_In, Դ�汾��_In, n_Ŀ��汾��, Դ��֧id_In,
                             n_Branch_New_Id);
        --�ٴ�·���׶ε��Ӽ���
        For r_�ٴ�·���ӽ׶� In (Select ID
                          From �ٴ�·���׶�
                          Where ·��id = Դ·��id_In And �汾�� = n_Ŀ��汾�� And ��id = r_�ٴ�·���׶�.Id And
                                Nvl(��֧id, 0) = Nvl(Դ��֧id_In, 0)) Loop
          --�����µĽ׶�ID
          Select �ٴ�·���׶�_Id.Nextval Into n_Step_New_Id From Dual;
          �ٴ�·���׶�_Insert(r_�ٴ�·���ӽ׶�.Id, n_Step_New_Id, Ŀ��·��id_In, n_Ŀ��汾��, n_Step_Parent_Id, Դ��֧id_In, n_Branch_New_Id);
        
          �ٴ�·���׶�cascade_Insert(r_�ٴ�·���ӽ׶�.Id, n_Step_New_Id, Դ·��id_In, Ŀ��·��id_In, Դ�汾��_In, n_Ŀ��汾��, Դ��֧id_In,
                               n_Branch_New_Id);
        End Loop;
      End If;
    End Loop;
  End If;

  --�ٴ�·����֧
  If Nvl(Դ��֧id_In, 0) = 0 And Nvl(Ŀ���֧id_In, 0) = 0 Then
    --�����汾ʱ
    For r_�ٴ�·����֧ In (Select ID From �ٴ�·����֧ Where ·��id = Դ·��id_In And �汾�� = n_Դ�汾��) Loop
      Select �ٴ�·����֧_Id.Nextval Into n_Branch_New_Id From Dual;
      �ٴ�·����֧_Insert(r_�ٴ�·����֧.Id, n_Branch_New_Id, Ŀ��·��id_In, n_Ŀ��汾��);
    
      Insert Into �ٴ�·������
        (·��id, �汾��, ���, ����, ��֧id)
        Select Ŀ��·��id_In, n_Ŀ��汾��, ���, ����, n_Branch_New_Id
        From �ٴ�·������
        Where ·��id = Դ·��id_In And �汾�� = n_Դ�汾�� And ��֧id = r_�ٴ�·����֧.Id;
    
      For r_�ٴ�·���׶� In (Select ID
                       From �ٴ�·���׶�
                       Where ·��id = Դ·��id_In And �汾�� = n_Դ�汾�� And ��id Is Null And ��֧id = r_�ٴ�·����֧.Id
                       Order By ���) Loop
        --�ٴ�·���׶εĸ����в���
        Select �ٴ�·���׶�_Id.Nextval Into n_Step_Parent_Id From Dual;
        �ٴ�·���׶�_Insert(r_�ٴ�·���׶�.Id, n_Step_Parent_Id, Ŀ��·��id_In, n_Ŀ��汾��, Null, r_�ٴ�·����֧.Id, n_Branch_New_Id);
      
        �ٴ�·���׶�cascade_Insert(r_�ٴ�·���׶�.Id, n_Step_Parent_Id, Դ·��id_In, Ŀ��·��id_In, Դ�汾��_In, n_Ŀ��汾��, r_�ٴ�·����֧.Id,
                             n_Branch_New_Id);
        --�ٴ�·���׶ε��Ӽ���
        For r_�ٴ�·���ӽ׶� In (Select ID
                          From �ٴ�·���׶�
                          Where ·��id = Դ·��id_In And �汾�� = n_Դ�汾�� And ��id = r_�ٴ�·���׶�.Id And ��֧id = r_�ٴ�·����֧.Id) Loop
          --�����µĽ׶�ID
          Select �ٴ�·���׶�_Id.Nextval Into n_Step_New_Id From Dual;
          �ٴ�·���׶�_Insert(r_�ٴ�·���ӽ׶�.Id, n_Step_New_Id, Ŀ��·��id_In, n_Ŀ��汾��, n_Step_Parent_Id, r_�ٴ�·����֧.Id, n_Branch_New_Id);
        
          �ٴ�·���׶�cascade_Insert(r_�ٴ�·���ӽ׶�.Id, n_Step_New_Id, Դ·��id_In, Ŀ��·��id_In, Դ�汾��_In, n_Ŀ��汾��, r_�ٴ�·����֧.Id,
                               n_Branch_New_Id);
        End Loop;
      End Loop;
    End Loop;
  
  End If;

  If Nvl(�Ƿ��֧·��_In, 0) <> 1 Then
    --������·�����ƻ��������汾��
    --�ٴ�·������
    Insert Into �ٴ�·������
      (·��id, �汾��, ���, ����)
      Select Ŀ��·��id_In, n_Ŀ��汾��, ���, ����
      From �ٴ�·������
      Where ·��id = Դ·��id_In And �汾�� = n_Դ�汾�� And ��֧id Is Null;
  
    --�ٴ�·����Ŀ
    --�ٴ�·��ҽ��
    --·��ҽ������
    --�ٴ�·������
    --�ٴ�·������
    --·������ָ��
    --·����������
  
    For r_�ٴ�·���׶� In (Select ID
                     From �ٴ�·���׶�
                     Where ·��id = Դ·��id_In And �汾�� = n_Դ�汾�� And ��id Is Null And ��֧id Is Null
                     Order By ���) Loop
      --�ٴ�·���׶εĸ����в���
      Select �ٴ�·���׶�_Id.Nextval Into n_Step_Parent_Id From Dual;
      �ٴ�·���׶�_Insert(r_�ٴ�·���׶�.Id, n_Step_Parent_Id, Ŀ��·��id_In, n_Ŀ��汾��, Null);
      If Nvl(Դ��֧id_In, 0) = 0 And Nvl(Ŀ���֧id_In, 0) = 0 Then
        --�����汾ʱ,����ǰһ�׶�ID
        Update �ٴ�·����֧
        Set ǰһ�׶�id = n_Step_Parent_Id
        Where ǰһ�׶�id = r_�ٴ�·���׶�.Id And �汾�� = n_Ŀ��汾��;
      End If;
    
      �ٴ�·���׶�cascade_Insert(r_�ٴ�·���׶�.Id, n_Step_Parent_Id, Դ·��id_In, Ŀ��·��id_In, Դ�汾��_In, n_Ŀ��汾��);
      --�ٴ�·���׶ε��Ӽ���
      For r_�ٴ�·���ӽ׶� In (Select ID
                        From �ٴ�·���׶�
                        Where ·��id = Դ·��id_In And �汾�� = n_Դ�汾�� And ��id = r_�ٴ�·���׶�.Id And ��֧id Is Null) Loop
        --�����µĽ׶�ID
        Select �ٴ�·���׶�_Id.Nextval Into n_Step_New_Id From Dual;
        �ٴ�·���׶�_Insert(r_�ٴ�·���ӽ׶�.Id, n_Step_New_Id, Ŀ��·��id_In, n_Ŀ��汾��, n_Step_Parent_Id);
      
        �ٴ�·���׶�cascade_Insert(r_�ٴ�·���ӽ׶�.Id, n_Step_New_Id, Դ·��id_In, Ŀ��·��id_In, Դ�汾��_In, n_Ŀ��汾��);
      End Loop;
    End Loop;
  End If;
Exception
  When Err_Custom Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Error || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�ٴ�·���汾_Copy;
/

--108821:��ҵ��,2017-05-15,���Ϻ�δ�����д������Ϣ
Create Or Replace Procedure Zl_�����շ���¼_��������
(
  �շ�id_In   In ҩƷ�շ���¼.Id%Type,
  �����_In   In ҩƷ�շ���¼.�����%Type,
  �������_In In ҩƷ�շ���¼.�������%Type,
  ����_In     In ҩƷ���.�ϴ�����%Type := Null,
  Ч��_In     In ҩƷ���.Ч��%Type := Null,
  ����_In     In ҩƷ���.�ϴβ���%Type := Null,
  ��������_In In ҩƷ�շ���¼.ʵ������%Type := Null,
  �Զ�����_In Integer := 0,
  ������_In   In ҩƷ�շ���¼.������%Type := Null
) Is
  Err_Item Exception;
  v_Err_Msg Varchar2(100);
  v_No      ҩƷ�շ���¼.No%Type;

  n_��¼״̬   ҩƷ�շ���¼.��¼״̬%Type;
  n_ִ��״̬   סԺ���ü�¼.ִ��״̬%Type;
  n_��������   Number;
  n_������id Number(18);
  n_����       ҩƷ�շ���¼.����%Type;
  n_�ⷿid     ҩƷ�շ���¼.�ⷿid%Type;
  n_ҩƷid     ҩƷ�շ���¼.ҩƷid%Type;
  n_ʵ������   ҩƷ�շ���¼.ʵ������%Type;
  n_ʵ�ʽ��   ҩƷ�շ���¼.���۽��%Type;
  n_ʵ�ʳɱ�   ҩƷ�շ���¼.�ɱ����%Type;
  n_ʵ�ʲ��   ҩƷ�շ���¼.���%Type;
  n_����id     ҩƷ�շ���¼.����id%Type;
  n_���ۼ�     ҩƷ�շ���¼.���ۼ�%Type;
  n_ʵ������   �շ���ĿĿ¼.�Ƿ���%Type;

  --��������ʱ�������������ʸı��Ĵ���
  n_������       ҩƷ�շ���¼.����%Type;
  n_����         ҩƷ�շ���¼.����%Type;
  n_����         ��������.���÷���%Type;
  n_С��         Number(2);
  n_�ϴι�Ӧ��id ҩƷ���.�ϴι�Ӧ��id%Type;
  n_�ɱ���       ҩƷ�շ���¼.�ɱ���%Type;
  d_�ϴ��������� ҩƷ���.�ϴ���������%Type;
  d_���Ч��     ҩƷ���.���Ч��%Type;
  v_��׼�ĺ�     ҩƷ���.��׼�ĺ�%Type;
  v_����         ҩƷ�շ���¼.����%Type;
  v_����no       סԺ���ü�¼.No%Type;
  v_Temp         Varchar2(255);
  v_��Ա���     ��Ա��.���%Type;
  v_��Ա����     ��Ա��.����%Type;
  n_��ҳid       סԺ���ü�¼.��ҳid%Type;
  n_���         סԺ���ü�¼.���%Type;
  v_������Դ     ����ҽ����¼.������Դ%Type;

  v_����id     ҩƷ�շ���¼.Id%Type;
  v_���no     ҩƷ�շ���¼.No%Type;
  v_������   Number(5) := 0;
  v_ִ��ʱ��   ҩƷ�շ���¼.�������%Type;
  n_ƽ���ɱ��� ҩƷ���.ƽ���ɱ���%Type;
  n_������¼id ҩƷ�շ���¼.Id%Type;
  n_�ƿ�       Number(1) := 0;
  v_��Ʒ����   ҩƷ���.��Ʒ����%Type;
  v_�ڲ�����   ҩƷ���.�ڲ�����%Type;
Begin
  v_Temp     := Zl_Identity;
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ';') + 1);
  v_Temp     := Substr(v_Temp, Instr(v_Temp, ',') + 1);
  v_��Ա��� := Substr(v_Temp, 1, Instr(v_Temp, ',') - 1);
  v_��Ա���� := Substr(v_Temp, Instr(v_Temp, ',') + 1);

  Select Zl_To_Number(Nvl(zl_GetSysParameter(9), '2')) Into n_С�� From Dual;

  If ��������_In Is Not Null Then
    If ��������_In = 0 Then
      Return;
    End If;
  End If;

  --1���жϵ�ǰ�����Ƿ��Ǳ�������
  Begin
    Select ���ܷ�ҩ��
    Into v_����id
    From ҩƷ�շ���¼
    Where ���� = 21 And ������� Is Not Null And
          ���ܷ�ҩ�� =
          (Select Max(a.Id)
           From ҩƷ�շ���¼ A, ҩƷ�շ���¼ B
           Where a.���� = b.���� And a.No = b.No And a.��� = b.��� And b.Id = �շ�id_In And (Mod(a.��¼״̬, 3) = 1 Or a.��¼״̬ = 1)) And
          Rownum = 1;
  Exception
    When Others Then
      v_����id := 0;
  End;

  --��ȡ���շ���¼�ĵ��ݡ�ҩƷID���ⷿID
  Select ����, NO, �ⷿid, ҩƷid, ����id, ������id, ��¼״̬, Nvl(����, 0), ��������, ���Ч��, ��׼�ĺ�, ��ҩ��λid, �ɱ���, ����, ���ۼ�, ��Ʒ����, �ڲ�����
  Into n_����, v_No, n_�ⷿid, n_ҩƷid, n_����id, n_������id, n_��¼״̬, n_����, d_�ϴ���������, d_���Ч��, v_��׼�ĺ�, n_�ϴι�Ӧ��id, n_�ɱ���, v_����,
       n_���ۼ�, v_��Ʒ����, v_�ڲ�����
  From ҩƷ�շ���¼
  Where ID = �շ�id_In;

  --��ȡ�ñʼ�¼ʣ��δ�������������
  --������������δ���������
  Select Sum(Nvl(ʵ������, 0) * Nvl(����, 1)), Sum(Nvl(���۽��, 0)), Sum(Nvl(�ɱ����, 0)), Sum(Nvl(���, 0))
  Into n_ʵ������, n_ʵ�ʽ��, n_ʵ�ʳɱ�, n_ʵ�ʲ��
  From ҩƷ�շ���¼
  Where ����� Is Not Null And NO = v_No And ���� = n_���� And ��� = (Select ��� From ҩƷ�շ���¼ Where ID = �շ�id_In);

  --���������ҩ��Ϊ�㣬��ʾ����ҩ
  If n_ʵ������ = 0 Then
    v_Err_Msg := '�õ����ѱ���������Ա���ϣ���ˢ�º����ԣ�';
    Raise Err_Item;
  End If;

  If Nvl(��������_In, 0) > n_ʵ������ Then
    v_Err_Msg := '�õ����ѱ���������Ա�������ϣ���ˢ�º����ԣ�';
    Raise Err_Item;
  End If;

  --��ȡ�ò��ϵ�ǰ�Ƿ��������Ϣ
  Select Nvl(���÷���, 0) Into n_���� From �������� Where ����id = n_ҩƷid;

  --����ǲ������ϣ������¼������۽����
  n_�������� := 0;
  If Not (��������_In Is Null Or Nvl(��������_In, 0) = n_ʵ������) Then
    n_�������� := 1;
  End If;

  If n_�������� = 1 Then
    n_ʵ�ʽ�� := Round(n_ʵ�ʽ�� * ��������_In / n_ʵ������, n_С��);
    n_ʵ�ʳɱ� := Round(n_ʵ�ʳɱ� * ��������_In / n_ʵ������, n_С��);
    n_ʵ�ʲ�� := Round(n_ʵ�ʲ�� * ��������_In / n_ʵ������, n_С��);
    n_ʵ������ := ��������_In;
  End If;

  --n_����:0-������;1-����;2-ԭ�������ֲ�������������������;3-ԭ���������ַ���������������
  If n_���� = 0 And n_���� <> 0 Then
    --ԭ�������ֲ�������������������
    n_���� := 2;
  Elsif n_���� <> 0 And n_���� = 0 Then
    --ԭ������,�ַ���,�����µ����Σ������²����ķ�ҩ��¼��ʹ��
    n_���� := 3;
  Else
    If n_���� = 0 Then
      n_���� := 0;
    Else
      n_���� := 1;
    End If;
  End If;
  If ����_In Is Not Null Then
    v_���� := ����_In;
  End If;
  --��¼״̬�ĺ��������仯
  --�����ļ�¼״̬        :iif(n_��¼״̬=1,0,1)+1
  --�������ļ�¼״̬        :iif(n_��¼״̬=1,0,1)+2
  --�ȴ����ϵļ�¼״̬    :iif(n_��¼״̬=1,0,1)+3
  Select ҩƷ�շ���¼_Id.Nextval Into n_������¼id From Dual;
  --����������¼
  Insert Into ҩƷ�շ���¼
    (ID, ��¼״̬, ����, NO, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ���Ч��, ����, ��д����, ʵ������, �ɱ���, �ɱ����, ����, ���ۼ�,
     ���۽��, ���, ժҪ, ������, ��������, ��ҩ��, �����, �������, ����id, ����, Ƶ��, �÷�, ��ҩ����, ������, ��ҩ��λid, ��������, ��׼�ĺ�, ��Ʒ����, �ڲ�����)
    Select n_������¼id, n_��¼״̬ + Decode(n_��¼״̬, 1, 0, 1) + 1, n_����, v_No, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ����, ����,
           Ч��, ���Ч��, 1, -n_ʵ������, -n_ʵ������, �ɱ���, -n_ʵ�ʳɱ�, ����, ���ۼ�, -n_ʵ�ʽ��, -n_ʵ�ʲ��, ժҪ, �����_In, �������_In, ��ҩ��, �����_In,
           �������_In, ����id, ����, Ƶ��, �÷�, ��ҩ����, ������_In, ��ҩ��λid, ��������, ��׼�ĺ�, ��Ʒ����, �ڲ�����
    From ҩƷ�շ���¼
    Where ID = �շ�id_In;

  --����ǲ��ֳ�����������Ϊ1��ʵ������Ϊ������ʵ�������Ļ�
  --����������¼�Թ���������
  Select ҩƷ�շ���¼_Id.Nextval Into n_������ From Dual;

  Insert Into ҩƷ�շ���¼
    (ID, ��¼״̬, ����, NO, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ���Ч��, ����, ��д����, ʵ������, �ɱ���, �ɱ����, ����, ���ۼ�,
     ���۽��, ���, ժҪ, ������, ��������, ��ҩ��, �����, �������, ����id, ����, Ƶ��, �÷�, ��ҩ����, ��ҩ��λid, ��������, ��׼�ĺ�, ��Ʒ����, �ڲ�����)
    Select n_������, n_��¼״̬ + Decode(n_��¼״̬, 1, 0, 1) + 3, n_����, v_No, ���, �ⷿid, �Է�����id, ������id, ���ϵ��, ҩƷid,
           Decode(n_����, 1, ����, 3, n_������, Null), Decode(n_����, 3, ����_In, 1, ����, Null), Decode(n_����, 3, ����_In, 1, ����, Null),
           Decode(n_����, 3, Ч��_In, 1, Ч��, Null), ���Ч��, 1, n_ʵ������, n_ʵ������, �ɱ���, n_ʵ�ʳɱ�, ����, ���ۼ�, n_ʵ�ʽ��, n_ʵ�ʲ��, ժҪ, ������,
           ��������, Null, Null, Null, ����id, ����, Ƶ��, �÷�, ��ҩ����, ��ҩ��λid, ��������, ��׼�ĺ�, ��Ʒ����, �ڲ�����
    From ҩƷ�շ���¼
    Where ID = �շ�id_In;

  --���²��˷��ü�¼��ִ��״̬(0-δִ��;1-��ȫִ��;2-����ִ��)
  Select Decode(Sum(Nvl(����, 1) * ʵ������), Null, 0, 0, 0, 2)
  Into n_ִ��״̬
  From ҩƷ�շ���¼
  Where ���� = n_���� And NO = v_No And ����id = n_����id And ����� Is Not Null;

  If n_ִ��״̬ = 0 Then
    Update סԺ���ü�¼ Set ִ��״̬ = n_ִ��״̬, ִ���� = Null, ִ��ʱ�� = Null Where ID = n_����id;
    Update ������ü�¼
    Set ִ��״̬ = n_ִ��״̬, ִ���� = Null, ִ��ʱ�� = Null
    Where NO = v_No And
          ��� = (Select ��� From ������ü�¼ Where ID = (Select ����id From ҩƷ�շ���¼ Where ID = �շ�id_In)) And
          (Mod(��¼����, 10) = 1 Or Mod(��¼����, 10) = 2) And ��¼״̬ <> 2 And ִ�в���id = n_�ⷿid;
  Else
    Update סԺ���ü�¼ Set ִ��״̬ = n_ִ��״̬ Where ID = n_����id;
    Update ������ü�¼
    Set ִ��״̬ = n_ִ��״̬
    Where NO = v_No And
          ��� = (Select ��� From ������ü�¼ Where ID = (Select ����id From ҩƷ�շ���¼ Where ID = �շ�id_In)) And
          (Mod(��¼����, 10) = 1 Or Mod(��¼����, 10) = 2) And ��¼״̬ <> 2 And ִ�в���id = n_�ⷿid;
  End If;

  --����δ��ҩƷ��¼
  Begin
    Insert Into δ��ҩƷ��¼
      (����, NO, ����id, ��ҳid, ����, ���ȼ�, �Է�����id, �ⷿid, ��ҩ����, ��������, ���շ�, ��ҩ��, ��ӡ״̬, δ����)
      Select a.����, a.No, a.����id, a.��ҳid, a.����, Nvl(b.���ȼ�, 0) ���ȼ�, a.�Է�����id, a.�ⷿid, a.��ҩ����, a.��������, a.���շ�, Null, 1, 1
      From (Select b.����, b.No, a.����id, a.��ҳid, a.����, Decode(a.��¼����, 1, Decode(a.����Ա����, Null, 0, 1), 1) ���շ�, b.�Է�����id,
                    b.�ⷿid, b.��ҩ����, b.��������, c.���
             From סԺ���ü�¼ A, ҩƷ�շ���¼ B, ������Ϣ C
             Where b.Id = �շ�id_In And a.Id = b.����id + 0 And a.����id = c.����id(+)
             Union All
             Select b.����, b.No, a.����id, Null As ��ҳid, a.����, Decode(a.��¼����, 1, Decode(a.����Ա����, Null, 0, 1), 1) ���շ�,
                    b.�Է�����id, b.�ⷿid, b.��ҩ����, b.��������, c.���
             From ������ü�¼ A, ҩƷ�շ���¼ B, ������Ϣ C
             Where b.Id = �շ�id_In And a.Id = b.����id + 0 And a.����id = c.����id(+)) A, ��� B
      Where b.����(+) = a.���;
  Exception
    When Others Then
      Null;
  End;

  --�޸�ԭ��¼Ϊ��������¼
  Update ҩƷ�շ���¼ Set ��¼״̬ = n_��¼״̬ + Decode(n_��¼״̬, 1, 0, 1) + 2 Where ID = �շ�id_In;

  --�޸�ҩƷ���(������)
  Select �Ƿ��� Into n_ʵ������ From �շ���ĿĿ¼ Where ID = n_ҩƷid;

  If n_���� <> 3 Then
  
    Update ҩƷ���
    Set ʵ������ = Nvl(ʵ������, 0) + n_ʵ������, ʵ�ʽ�� = Nvl(ʵ�ʽ��, 0) + n_ʵ�ʽ��, ʵ�ʲ�� = Nvl(ʵ�ʲ��, 0) + n_ʵ�ʲ��,
        ���ۼ� = Decode(n_ʵ������, 1, Decode(Nvl(n_����, 0), 0, Null, Decode(Nvl(���ۼ�, 0), 0, n_���ۼ�, ���ۼ�)), Null)
    Where �ⷿid + 0 = n_�ⷿid And ҩƷid = n_ҩƷid And ���� = 1 And Nvl(����, 0) = n_����;
  
    If Sql%RowCount = 0 Then
      Insert Into ҩƷ���
        (�ⷿid, ҩƷid, ����, ����, ʵ������, ʵ�ʽ��, ʵ�ʲ��, Ч��, ���Ч��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴ���������, �ϴβ���, ��׼�ĺ�, ���ۼ�, ƽ���ɱ���, ��Ʒ����,
         �ڲ�����)
      Values
        (n_�ⷿid, n_ҩƷid, Decode(n_����, 2, Null, n_����), 1, n_ʵ������, n_ʵ�ʽ��, n_ʵ�ʲ��, Decode(n_����, 1, Ч��_In, Null), d_���Ч��,
         n_�ϴι�Ӧ��id, n_�ɱ���, Decode(n_����, 1, ����_In, Null), d_�ϴ���������, v_����, v_��׼�ĺ�,
         Decode(n_ʵ������, 1, Decode(Nvl(n_����, 0), 0, Null, n_���ۼ�), Null), n_�ɱ���, v_��Ʒ����, v_�ڲ�����);
    End If;
  Else
    Insert Into ҩƷ���
      (�ⷿid, ҩƷid, ����, ����, ʵ������, ʵ�ʽ��, ʵ�ʲ��, Ч��, ���Ч��, �ϴι�Ӧ��id, �ϴβɹ���, �ϴ�����, �ϴ���������, �ϴβ���, ��׼�ĺ�, ���ۼ�, ƽ���ɱ���, ��Ʒ����, �ڲ�����)
    Values
      (n_�ⷿid, n_ҩƷid, n_������, 1, n_ʵ������, n_ʵ�ʽ��, n_ʵ�ʲ��, Ч��_In, d_���Ч��, n_�ϴι�Ӧ��id, n_�ɱ���, ����_In, d_�ϴ���������, v_����, v_��׼�ĺ�,
       Decode(n_ʵ������, 1, Decode(Nvl(n_������, 0), 0, Null, n_���ۼ�), Null), n_�ɱ���, v_��Ʒ����, v_�ڲ�����);
  End If;

  Delete ҩƷ���
  Where �ⷿid + 0 = n_�ⷿid And ҩƷid = n_ҩƷid And ���� = 1 And Nvl(��������, 0) = 0 And Nvl(ʵ������, 0) = 0 And Nvl(ʵ�ʽ��, 0) = 0 And
        Nvl(ʵ�ʲ��, 0) = 0;

  If �Զ�����_In = 1 And n_���� <> 24 Then
    Begin
      Select ��ҳid, NO, ��� Into n_��ҳid, v_����no, n_��� From סԺ���ü�¼ Where ID = n_����id;
    Exception
      When Others Then
        Begin
          Select Null, NO, ��� Into n_��ҳid, v_����no, n_��� From ������ü�¼ Where ID = n_����id;
        Exception
          When Others Then
            n_��ҳid := Null;
        End;
    End;
    If n_��ҳid Is Null Then
      Zl_������ʼ�¼_Delete(v_����no, n_���, v_��Ա���, v_��Ա����);
    Else
      Zl_סԺ���ʼ�¼_Delete(v_����no, n_���, v_��Ա���, v_��Ա����);
    End If;
  End If;

  --�������Ĵ���
  If v_����id > 0 Then
    --2���Զ���������˵��������ⵥ��
    Begin
      Select 1
      Into n_�ƿ�
      From ҩƷ�շ���¼
      Where ���� = 15 And ������� Is Null And
            ����id In (Select Distinct ����id From ҩƷ�շ���¼ Where NO = v_No And ҩƷid = n_ҩƷid And ���� = n_����);
    Exception
      When Others Then
        n_�ƿ� := 0;
    End;
    If n_�ƿ� <> 0 Then
      For v_������� In (Select 1 �д�, ��¼״̬, NO, ���, ҩƷid
                     From ҩƷ�շ���¼
                     Where ���� = 21 And ������� Is Not Null And ���ܷ�ҩ�� = v_����id) Loop
      
        Zl_������������_Strike(v_�������.�д�, v_�������.��¼״̬, v_�������.No, v_�������.���, v_�������.ҩƷid, ��������_In, �����_In, �������_In, 1);
      End Loop;
    
      --3�������µ��������ⵥ��
      If v_���no Is Null Then
        v_���no := Nextno(74, n_�ⷿid);
      End If;
      v_������ := v_������ + 1;
    
      For v_��� In (Select ������id, �ⷿid, ҩƷid, ����, ��д����, �ɱ���, �ɱ����, ���ۼ�, ���۽��, ���, ����, ����, Ч��, ���Ч��, ժҪ, ����, ��ҩ����
                   From ҩƷ�շ���¼
                   Where ���� = 21 And ������� Is Not Null And ���ܷ�ҩ�� = v_����id) Loop
      
        Zl_������������_Insert(v_���.������id, v_���no, v_������, v_���.�ⷿid, v_���.ҩƷid, v_���.����, v_���.��д����, v_���.�ɱ���, v_���.�ɱ����,
                         v_���.���ۼ�, v_���.���۽��, v_���.���, �����_In, �������_In, v_���.����, v_���.����, v_���.Ч��, v_���.���Ч��, v_���.ժҪ,
                         v_���.����, v_���.��ҩ����);
      
        Update ҩƷ�շ���¼
        Set ����id = n_����id, ���ܷ�ҩ�� = n_������
        Where ���� = 21 And NO = v_���no And ��� = v_������;
      End Loop;
    
      --4��ɾ��δ��˵��⹺��ⵥ�ݣ�������򲻹ܣ�
      Delete ҩƷ�շ���¼
      Where ���� = 15 And ҩƷid = n_ҩƷid And Nvl(����, 0) = n_���� And ����id = n_����id And ������� Is Null;
    End If;
  End If;
  --���������������
  Zl_�����շ���¼_��������(n_������¼id);
Exception
  When Err_Item Then
    Raise_Application_Error(-20101, '[ZLSOFT]' || v_Err_Msg || '[ZLSOFT]');
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_�����շ���¼_��������;
/

--109194:��ҵ��,2017-05-23,�̵㵥��������������
Create Or Replace Procedure Zl_ҩƷ�̵��¼��_Insert
(
  No_In         In ҩƷ�շ���¼.No%Type,
  ���_In       In ҩƷ�շ���¼.���%Type,
  �ⷿid_In     In ҩƷ�շ���¼.�ⷿid%Type,
  ����_In       In ҩƷ�շ���¼.����%Type,
  ������id_In In ҩƷ�շ���¼.������id%Type,
  ���ϵ��_In   In ҩƷ�շ���¼.���ϵ��%Type,
  ҩƷid_In     In ҩƷ�շ���¼.ҩƷid%Type,
  ��������_In   In ҩƷ�շ���¼.��д����%Type,
  ʵ������_In   In ҩƷ�շ���¼.����%Type,
  ������_In     In ҩƷ�շ���¼.ʵ������%Type,
  �ۼ�_In       In ҩƷ�շ���¼.���ۼ�%Type,
  ����_In     In ҩƷ�շ���¼.���۽��%Type,
  ��۲�_In     In ҩƷ�շ���¼.���%Type,
  ������_In     In ҩƷ�շ���¼.������%Type,
  ��������_In   In ҩƷ�շ���¼.��������%Type,
  ժҪ_In       In ҩƷ�շ���¼.ժҪ%Type := Null,
  ����_In       In ҩƷ�շ���¼.����%Type := Null,
  ����_In       In ҩƷ�շ���¼.����%Type := Null,
  Ч��_In       In ҩƷ�շ���¼.Ч��%Type := Null,
  �̵�ʱ��_In   In ҩƷ�շ���¼.Ƶ��%Type := Null,
  �����_In   In ҩƷ�շ���¼.�ɱ���%Type := Null,
  �����_In   In ҩƷ�շ���¼.�ɱ����%Type := Null,
  ��׼�ĺ�_In   In ҩƷ�շ���¼.��׼�ĺ�%Type := Null,
  �ɱ���_In     In ҩƷ�շ���¼.����%Type := Null,
  �ⷿ��λ_In   In ҩƷ�շ���¼.�ⷿ��λ%Type := Null
) Is
  v_���� ҩƷ�շ���¼.����%Type;
Begin
  v_���� := ����_In;
  If v_���� < 0 Then
    v_���� := Zl_Fun_Getbatchnum(ҩƷid_In, ����_In, ����_In, �ɱ���_In, �ۼ�_In, ����_In);
  End If;

  Insert Into ҩƷ�շ���¼
    (ID, ��¼״̬, ����, NO, ���, �ⷿid, ������id, ���ϵ��, ҩƷid, ����, ����, ����, Ч��, ��д����, ����, ʵ������, ���ۼ�, ���۽��, ���, ժҪ, ������, ��������, Ƶ��,
     �ɱ���, �ɱ����, ��׼�ĺ�, ����, �ⷿ��λ)
  Values
    (ҩƷ�շ���¼_Id.Nextval, 1, 14, No_In, ���_In, �ⷿid_In, ������id_In, ���ϵ��_In, ҩƷid_In, v_����, ����_In, ����_In, Ч��_In, ��������_In,
     ʵ������_In, ������_In, �ۼ�_In, ����_In, ��۲�_In, ժҪ_In, ������_In, ��������_In, �̵�ʱ��_In, �����_In, �����_In, ��׼�ĺ�_In, �ɱ���_In, �ⷿ��λ_In);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_ҩƷ�̵��¼��_Insert;
/


--109194:��ҵ��,2017-05-23,�̵㵥��������������
Create Or Replace Function Zl_Fun_Getbatchnum
(
  ҩƷid_In   ҩƷ���Ŷ���.ҩƷid%Type,
  ��������_In ҩƷ���Ŷ���.��������%Type,
  ����_In     ҩƷ���Ŷ���.����%Type,
  �ɱ���_In   ҩƷ���Ŷ���.�ɱ���%Type,
  �ۼ�_In     ҩƷ���Ŷ���.�ۼ�%Type,
  ������_In   ҩƷ���Ŷ���.����%Type
) Return Number Is
  --���ܣ�ҩƷ����������¼ʱ���ݴ��ݹ����Ĳ����Ҷ�Ӧ������
  --����ֵ����ѯ�������Σ��������>0��˵���ҵ�������,�������=0��˵��û���ҵ�
  --������
  --     ��������_in����⴫�ݹ�����������
  --     ����_in�����ʱ¼�������
  --     �ɱ���_in ���ʱ�ĳɱ���
  --     �ۼ�_in  ���ʱ���ۼ�
  --
  n_����     ҩƷ���Ŷ���.����%Type;
  n_ҩ���װ ҩƷ���.ҩ���װ%Type;
  n_�Ƿ��� �շ���ĿĿ¼.�Ƿ���%Type;
  n_Count    Number(1);
Begin
  --ֻ�����������Һ����Ų�Ϊ�յ����
  If ��������_In Is Not Null And ����_In Is Not Null Then
    Begin
      Select ����
      Into n_����
      From ҩƷ���Ŷ���
      Where ҩƷid = ҩƷid_In And Nvl(��������, 'a') = Nvl(��������_In, 'a') And Nvl(����, 'b') = Nvl(����_In, 'b') And �ɱ��� = �ɱ���_In And
            �ۼ� = �ۼ�_In;
    Exception
      When Others Then
        n_���� := ������_In;
      
        If n_���� > 0 Then
          --��������ظ���¼
          Begin
            Select 1
            Into n_Count
            From ҩƷ���Ŷ���
            Where ҩƷid = ҩƷid_In And Nvl(��������, 'a') = Nvl(��������_In, 'a') And Nvl(����, 'b') = Nvl(����_In, 'b') And
                  ���� = n_����;
          Exception
            When Others Then
              n_Count := 0;
          End;
        
          --û���ظ���¼���ܲ���
          If n_Count = 0 Then
            Insert Into ҩƷ���Ŷ���
              (ҩƷid, ��������, ����, ����, �ɱ���, �ۼ�)
            Values
              (ҩƷid_In, ��������_In, ����_In, ������_In, �ɱ���_In, �ۼ�_In);
          End If;
        End If;
    End;
  End If;

  Return(n_����);
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End Zl_Fun_Getbatchnum;
/

---------------------------------------------------------------------------------------------------
--����ϵͳ�������İ汾��
-------------------------------------------------------------------------------------------------------
--104825:�ƽ�,2017-04-07,RIS�ӿ���ҽ�������޸ĳɶ���exe
Insert Into Zlfilesupgrade
  (�ļ�����, �ļ���, �汾��, �޸�����, ����ϵͳ, ҵ�񲿼�, ��װ·��, �ļ�˵��, ǿ�Ƹ���, �Զ�ע��, ��������, ���)
  Select 1, 'ZL9XWINTERFACE.DLL', '', Null, '1,21', 'ZLSVRSTUDIO.EXE,ZLHIS+.EXE,zl9BaseItem.dll,zl9CISJob.dll,zl9PACSWork.dll,zlCISKernel.dll,zl9peimanage.dll', '[APPSOFT]\APPLY', 'XWRIS�ӿڲ���', '1', '1',
         Sysdate, ���
  From Dual a, (Select Max(To_Number(���)) + 1 ��� From Zlfilesupgrade) b
  Where Not Exists (Select 1 From Zlfilesupgrade Where Upper(�ļ���) = 'ZL9XWINTERFACE.DLL');

Insert Into Zlfilesupgrade
  (�ļ�����, �ļ���, �汾��, �޸�����, ����ϵͳ, ҵ�񲿼�, ��װ·��, �ļ�˵��, ǿ�Ƹ���, �Զ�ע��, ��������, ���)
  Select 1, 'ZLSOFTSHOWHISFORMS.EXE', '', Null, '1', 'zl9XWInterface.dll', '[APPSOFT]\APPLY',
         'RIS�鿴HIS�е��Ӳ���������ҽ����סԺҽ���ȹ��ܵĶ���exe����', '1', '0', Sysdate, ���
  From Dual a, (Select Max(To_Number(���)) + 1 ��� From Zlfilesupgrade) b
  Where Not Exists (Select 1 From Zlfilesupgrade Where Upper(�ļ���) = 'ZLSOFTSHOWHISFORMS.EXE');

--ϵͳ�汾��
Update zlSystems Set �汾��='10.34.110' Where ���=&n_System;
--�����汾��
Commit;