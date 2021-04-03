-----------------------------------------------------------------
--为配合产品版本号由9.30升为9.30.10(VZLHIS10.19.40)
-----------------------------------------------------------------
--数据升级:
--问题:12218
BEGIN 
	--先所有的站点取掉空格
	DELETE zltools.zlclients a WHERE ROWID< (Select Max(Rowid) FROM zlclients WHERE trim(a.工作站)=trim(工作站));

	UPDATE zltools.zlClients SET 工作站=trim(工作站);

	INSERT INTO zltools.zlclients (工作站, Ip, Cpu, 内存, 硬盘, 操作系统, 部门, 用途, 说明, 升级服务器, 升级标志, 收集标志, 禁止使用, 连接数)
	Select 工作站||rpad(' ',lengthb(工作站)-length(工作站),' '), Ip, Cpu, 内存, 硬盘, 操作系统, 部门, 用途, 说明, 升级服务器,1 AS  升级标志, 收集标志, 禁止使用, 连接数
	From zlTools.zlClients
	WHERE lengthb(工作站)-length(工作站)<>0;
END ;
/
