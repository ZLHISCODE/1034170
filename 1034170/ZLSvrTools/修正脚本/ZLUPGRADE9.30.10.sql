-----------------------------------------------------------------
--Ϊ��ϲ�Ʒ�汾����9.30��Ϊ9.30.10(VZLHIS10.19.40)
-----------------------------------------------------------------
--��������:
--����:12218
BEGIN 
	--�����е�վ��ȡ���ո�
	DELETE zltools.zlclients a WHERE ROWID< (Select Max(Rowid) FROM zlclients WHERE trim(a.����վ)=trim(����վ));

	UPDATE zltools.zlClients SET ����վ=trim(����վ);

	INSERT INTO zltools.zlclients (����վ, Ip, Cpu, �ڴ�, Ӳ��, ����ϵͳ, ����, ��;, ˵��, ����������, ������־, �ռ���־, ��ֹʹ��, ������)
	Select ����վ||rpad(' ',lengthb(����վ)-length(����վ),' '), Ip, Cpu, �ڴ�, Ӳ��, ����ϵͳ, ����, ��;, ˵��, ����������,1 AS  ������־, �ռ���־, ��ֹʹ��, ������
	From zlTools.zlClients
	WHERE lengthb(����վ)-length(����վ)<>0;
END ;
/
