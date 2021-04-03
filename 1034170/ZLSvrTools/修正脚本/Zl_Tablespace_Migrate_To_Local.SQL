	--8022
  --V9.24发布时提供，V9.27时调整
	--功能:将表空间的类型由字典管理转换为本地管理(系统表空间除外)
	--要求:
	--          以SYS用户执行,或授于当前用户相关权限(Dbms_Space_Admin)
	--          Oracle参数文件Compatible参数为8.1.5以上(修改后需重启数据库)
	--说明:
	--          Oracle 8i开始提供本地管理表空间,相对于字典管理表空间,主要有如下好处:
	--          a.减少了递归空间管理(因为申请空间而产生的字典表记录又申请空间)
	--          b.减少了数据字典之间的竞争(UET$和FET$表的并发争用),并避免了由此产生的字典表的回退信息
	--          c.系统自动管理extents大小或采用统一extents大小,不需合并相邻的剩余空间,减少了空间碎片
	--          
	--          在PL/SQL中运行并显示可能发生的出错信息,请先运行SET SERVEROUT ON;
	--          如果要进行反向操作,可以运行DBMS_SPACE_ADMIN.TABLESPACE_MIGRATE_FROM_LOCAL
	Create Or Replace Procedure Zl_Tablespace_Migrate_To_Local As
	Strsql    Varchar2(200);
	v_Version Varchar2(2);
	v_Tmp     Varchar2(200);
  v_Path    Varchar2(255);
Begin
	--1.公共回退段的处理
	--8i的回退段,不支持继续存放在本地管理的临时表空间,改为存放在系统回退表空间RBS中
	Select Substr(Banner, 6, 2) Into v_Version From V$version Where Substr(Banner, 1, 4) = 'CORE';
	If To_Number(v_Version) < 9 Then
		Begin
			Select Segment_Name Into v_Tmp From Dba_Rollback_Segs Where Segment_Name = 'RBS_ZLTOOLS';
		Exception
			When Others Then
				Null;
		End;
		If NOT v_Tmp IS null Then
			Begin
				Strsql := 'Alter Rollback Segment RBS_ZLTOOLS Offline';
				Execute Immediate Strsql;
				Strsql := 'Drop Rollback Segment RBS_ZLTOOLS';
				Execute Immediate Strsql;

				Strsql := 'Create Public Rollback Segment RBS_ZLTOOLS Tablespace RBS';
				Execute Immediate Strsql;
				Strsql := 'Alter Rollback Segment RBS_ZLTOOLS Online';
				Execute Immediate Strsql;
			Exception
				When Others Then
					Dbms_Output.Put_Line('提示:' || Sqlerrm);
			End;
		End If;
	End If;

	--2.转换ZL的所有表空间(除临时表空间外)
	For c_Tp In (Select Tablespace_Name
							 From Dba_Tablespaces
							 Where Contents <> 'TEMPORARY' And Extent_Management <> 'LOCAL' And Status = 'ONLINE' And
										 Tablespace_Name<>'ZLTOOLSTMP'  AND Tablespace_Name Like 'ZL%'
							 Group By Tablespace_Name) Loop
		Dbms_Space_Admin.Tablespace_Migrate_To_Local(c_Tp.Tablespace_Name);
	End Loop;

	--3.临时表空间的特殊转换
	Begin
		Select Tablespace_Name Into v_Tmp From Dba_Tablespaces Where Tablespace_Name = 'ZLTOOLSTMP' And Rownum < 2;
		If Sql%Rowcount > 0 Then
			Strsql := 'DROP TABLESPACE ZLTOOLSTMP INCLUDING CONTENTS';
			Execute Immediate Strsql;
		End If;
	Exception
		When Others Then
			Dbms_Output.Put_Line(Sqlerrm);
	End;
  
  --表空间文件缺省创建在ZLTOOLSTBS的相同目录
  Begin
	  Select Substr(File_Name,1,Instr(File_Name,Tablespace_Name||'.')-1) Into v_Path From DBA_Data_Files Where TableSpace_Name='ZLTOOLSTBS';
    --如果改了管理工具表空间的缺省名称,则以ZL9BASEITEM为准,如果没有ZL9BASEITEM,则以Oracle的缺省数据库文件路径为准
    IF v_Path IS NULL THEN
      Select Substr(File_Name,1,Instr(File_Name,Tablespace_Name||'.')-1) Into v_Path From DBA_Data_Files Where TableSpace_Name='ZL9BASEITEM';
    END IF;
  Exception
		When Others Then
			Null;
	End;  
	Strsql := 'CREATE TEMPORARY TABLESPACE ZLTOOLSTMP TEMPFILE ' || Chr(39) || v_Path || 'ZLTOOLSTMP.DBF' || Chr(39) ||
						' SIZE 100M REUSE AUTOEXTEND ON NEXT 10M EXTENT MANAGEMENT LOCAL UNIFORM SIZE 1M';
	Execute Immediate Strsql;
End Zl_Tablespace_Migrate_To_Local;
/
Execute Zl_Tablespace_Migrate_To_Local;