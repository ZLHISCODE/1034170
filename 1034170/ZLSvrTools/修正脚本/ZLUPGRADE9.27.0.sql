-----------------------------------------------------------------
--Ϊ��ϲ�Ʒ�汾����9.26��Ϊ9.27
--�ӱ�����ʼʹ�������������� 
-----------------------------------------------------------------
ALTER TABLE zlTools.zlRPTConds Drop CONSTRAINT zlRPTConds_PK
/
ALTER TABLE zlTools.zlRPTConds Drop CONSTRAINT zlRPTConds_UQ_��������
/
ALTER TABLE zlTools.zlRPTConds ADD CONSTRAINT zlRPTConds_PK PRIMARY KEY (����ID,������,������) USING INDEX PCTFREE 5
/
ALTER TABLE zlTools.zlRPTConds ADD CONSTRAINT zlRPTConds_UQ_�������� UNIQUE (����ID,��������,������) USING INDEX PCTFREE 5
/