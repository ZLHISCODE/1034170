-----------------------------------------------------------------
--为配合产品版本号由9.30升为9.31(VZLHIS10.20.0)
-----------------------------------------------------------------

--体检升级到子系统时发现 by cfr
Alter Table zlBakTables Modify 系统 Number(5);
Alter Table zlBakSpaces Modify 系统 Number(5);

Alter Table zlStreamTabs Drop Constraint zlStreamTabs_FK_SYSNO;
Alter Table zlStreamTabs Add Constraint zlStreamTabs_FK_SYSNO Foreign Key (System_NO) References zlsystems(编号) ON DELETE CASCADE
/

--创建系统参数相关对象
Create Sequence zlTools.zlParameters_ID Start With 1
/
Create Table zlTools.zlParameters(
    ID NUMBER(18),
    系统 NUMBER(5),
    模块 NUMBER(18),
    私有 NUMBER(1),
    参数号 NUMBER(5),
    参数名 VARCHAR2(100),
		参数值 VARCHAR2(1000),
		缺省值 VARCHAR2(1000),
		参数说明 VARCHAR2(255))
    PCTFREE 5 PCTUSED 90
    Cache Storage(Buffer_Pool Keep)
/
Alter Table zlTools.zlParameters Add Constraint zlParameters_PK Primary Key(ID) Using Index PCTFREE 5
/
Alter Table zlTools.zlParameters Add Constraint zlParameters_UQ_参数号 Unique(参数号,模块,系统,私有) Using Index PCTFREE 5
/
Alter Table zlTools.zlParameters Add Constraint zlParameters_UQ_参数名 Unique(参数名,模块,系统,私有) Using Index PCTFREE 5
/
Alter Table zlTools.zlParameters Add Constraint zlParameters_CK_私有 Check (私有 IN(0,1))
/

Create Table zlTools.zlUserParas(
    参数ID NUMBER(18),
    用户名 VARCHAR2(20),
		参数值 VARCHAR2(1000))
    PCTFREE 5 PCTUSED 90
    Cache Storage(Buffer_Pool Keep)
/
Alter Table zlTools.zlUserParas Add Constraint zlUserParas_PK Primary Key(参数ID,用户名) Using Index PCTFREE 5
/
Alter Table zlTools.zlUserParas Add Constraint zlUserParas_FK_参数ID Foreign Key (参数ID) References zlParameters(ID) On Delete Cascade
/
Create Index zlTools.zlUserParas_IX_用户名 On zlUserParas(用户名) PCTFREE 5
/

Create Or Replace Procedure zlTools.zl_Parameters_Update
(
	参数_In   zlParameters.参数名%Type,
	参数值_In zlParameters.参数值%Type,
	系统_In   zlParameters.系统%Type,
  模块_In   zlParameters.模块%Type,
  私有_In   zlParameters.私有%Type
  --功能：设置系统参数值，如果是用户私有参数，则用户名以当前的为准
  --参数：
  --      参数_In：必须传入非Null值，以字符形式传入的参数号或参数名,注意参数名不能为数字。
) Is
  v_参数id zlParameters.ID%Type;
Begin
  --确定参数
  Begin
    If Zl_To_Number(参数_In) <> 0 Then
      --以参数号为准处理
      Select ID
      Into v_参数id
      From zlParameters
      Where Nvl(系统, 0) = Nvl(系统_In, 0) And Nvl(模块, 0) = Nvl(模块_In, 0) And 参数号 = Zl_To_Number(参数_In) And
            Nvl(私有, 0) = Nvl(私有_In, 0);
    Else
      --以参数名为准处理
      Select ID
      Into v_参数id
      From zlParameters
      Where Nvl(系统, 0) = Nvl(系统_In, 0) And Nvl(模块, 0) = Nvl(模块_In, 0) And 参数名 = 参数_In And Nvl(私有, 0) = Nvl(私有_In, 0);
    End If;
  Exception
    When Others Then
      Return;
  End;

  --更新参数值
  If Nvl(私有_In, 0) = 0 Then
    Update zlParameters Set 参数值 = 参数值_In Where ID = v_参数id;
  Elsif Nvl(私有_In, 0) = 1 Then
    Update zlUserParas Set 参数值 = 参数值_In Where 用户名 = User And 参数id = v_参数id;
    If Sql%RowCount = 0 Then
      Insert Into zlUserParas (参数id, 用户名, 参数值) Values (v_参数id, User, 参数值_In);
    End If;
  End If;
End zl_Parameters_Update;
/

Create Or Replace Procedure zlTools.zl_Parameters_Update_Batch
(
  系统编号_In zlSystems.编号%Type,
  参数列表_In Varchar2
) Is
  --参数列表_IN 参数的填写方式如下："参数号1,参数值1,参数号2,参数值2,"                                            
  n_Pos    Number(5);
  v_Temp   Varchar2(2000);
  v_参数号 zlParameters.参数号%Type;
  v_参数值 zlParameters.参数值%Type;
Begin
  --循环处理
  v_Temp := 参数列表_In;

  While v_Temp Is Not Null Loop
    n_Pos := Instr(v_Temp, ',');
  
    If n_Pos = 0 Then
      v_Temp := '';
    Else
      --得到参数号
      v_参数号 := Substr(v_Temp, 1, n_Pos - 1);
      v_Temp   := Substr(v_Temp, n_Pos + 1);
      --得到参数值
      n_Pos    := Instr(v_Temp, ',');
      v_参数值 := Substr(v_Temp, 1, n_Pos - 1);
      v_Temp   := Substr(v_Temp, n_Pos + 1);
    
      Update zlParameters
      Set 参数值 = v_参数值
      Where 系统 = 系统编号_In And 模块 Is Null And Nvl(私有, 0) = 0 And 参数号 = To_Number(v_参数号);
    End If;
  End Loop;
Exception
  When Others Then
    zl_ErrorCenter(SQLCode, SQLErrM);
End zl_Parameters_Update_Batch;
/

--插入现本地参数内容
--私有全局：zlAppTools，导航台等公共部分
Insert Into zlParameters(ID,系统,模块,私有,参数号,参数名,参数值,缺省值,参数说明)
Select Rownum+B.ID,A.* From (
	Select 系统,模块,私有,参数号,参数名,参数值,缺省值,参数说明 From zlParameters Where ID=0 Union ALL
	Select -NULL,-NULL,1,1,'自动消息停留时间',NULL,'3','记录自动消息提醒信息停留时间(秒)' From Dual Union ALL
	Select -NULL,-NULL,1,2,'导航台',NULL,'zlBrw','记录使用哪种类型的导航台：zlBrw，zlWin，zlMdi' From Dual Union ALL
  Select -NULL,-NULL,1,3,'常用功能模块',NULL,NULL,'设置常用的功能模块' From Dual Union ALL
  Select -NULL,-NULL,1,4,'输入匹配',NULL,'0','设置输入匹配方向：0-双向匹配，1-从左匹配' From Dual Union ALL
  Select -NULL,-NULL,1,5,'输入法',NULL,NULL,'设置需要自动开启的输入法名称' From Dual Union ALL
  Select -NULL,-NULL,1,6,'简码方式',NULL,'0','设置简码生成或输入的方式：0-拼音，1-五笔' From Dual Union ALL
  Select -NULL,-NULL,1,7,'关闭Windows',NULL,'0','设置是否退出程序时自动关闭 Windows' From Dual Union ALL
  Select -NULL,-NULL,1,8,'邮件消息检查周期',NULL,'30','设置自动检查邮件消息的时间间隔(秒)' From Dual Union ALL
  Select -NULL,-NULL,1,9,'登录检查邮件消息',NULL,'0','设置是否登录时检查新的邮件消息' From Dual Union ALL
  Select -NULL,-NULL,1,10,'显示已读邮件',NULL,'0','设置在邮件管理器中是否显示已读邮件' From Dual Union ALL
	Select -NULL,-NULL,1,11,'最近使用模块',NULL,NULL,'记录最近使用的程序' From Dual Union ALL
	Select -NULL,-NULL,1,12,'使用个性化风格',NULL,'1','设置是否使用个性化风格' From Dual Union ALL   --刘兴宏:罗虹要求改为默认值
	Select -NULL,-NULL,1,13,'接收邮件消息',NULL,'0','设置是否接收邮件消息通知' From Dual Union ALL
	Select -NULL,-NULL,1,14,'zlBrwFontSize',NULL,'0','记录Brower风格导航台字体大小，由小到大分别为：0-9号,1-11号,2-12号' From Dual Union ALL
	Select -NULL,-NULL,1,15,'zlMdiFontColor',NULL,'-1','设置MDI风格导航台的字体颜色' From Dual Union ALL
	Select -NULL,-NULL,1,16,'zlMdiBackPic',NULL,NULL,'设置MDI风格导航台的背景图片文件路径' From Dual Union ALL
	Select -NULL,-NULL,1,17,'zlMdiMenuArray',NULL,'1','设置MDI风格导航台菜单排列方式：0-纵向排列，1-横向排列' From Dual Union ALL
	Select -NULL,-NULL,1,18,'zlWinFontColor',NULL,'-1','设置Windows风格导航台的字体颜色' From Dual Union ALL
	Select -NULL,-NULL,1,19,'zlWinBackPic',NULL,NULL,'设置Windows风格导航台的背景图片文件路径' From Dual
  ) A,(Select Nvl(Max(ID),0) AS ID From zlParameters) B
/
--最后调整zlParameters的序列
Select zlParameters_ID.Nextval From zlParameters
/

--现本地参数设置升级
-----------------------------------------
--私有全局：zlAppTools，导航台等公共部分
Declare
	v_方案号	zlClientScheme.方案号%Type;
	v_Val zlParameters.参数值%Type;
Begin
	--取方案名为"zlParaUpdate"的方案号，如果没有这个方案则不能完成升级转换
	Begin
		Select 方案号 Into v_方案号 From zlClientScheme Where Upper(方案名称)=Upper('zlParaUpdate');
	Exception
		When Others Then Return;
	End;
	
	--逐个参数进行升级转换
	Begin
		Select 键值 Into v_Val From zlClientParaList Where 类别='私有全局' And 目录='历史记录' And 键名='系统' And 方案号=v_方案号;
		Select v_Val||'|'||键值 Into v_Val From zlClientParaList Where 类别='私有全局' And 目录='历史记录' And 键名='序号' And 方案号=v_方案号;
		Update zlParameters Set 参数值=v_Val Where 系统 Is Null And 模块 Is Null And 私有=1 And 参数名='最近使用模块';
	Exception When Others Then Null; End;
	
	Begin

		Select 键值 Into v_Val From zlClientParaList Where 类别='私有全局' And 目录 Is Null And 键名='使用个性化风格' And 方案号=v_方案号;
		Update zlParameters Set 参数值=v_Val Where 系统 Is Null And 模块 Is Null And 私有=1 And 参数名='使用个性化风格';
	Exception When Others Then Null; End;

	Begin
		Select 键值 Into v_Val From zlClientParaList Where 类别='私有全局' And 目录 Is Null And 键名='消息通知' And 方案号=v_方案号;
		Update zlParameters Set 参数值=v_Val Where 系统 Is Null And 模块 Is Null And 私有=1 And 参数名='接收邮件消息';
	Exception When Others Then Null; End;

	Begin
		Select 键值 Into v_Val From zlClientParaList Where 类别='公共模块' And 目录='BROWER' And 键名='ZlBrowerFont' And 方案号=v_方案号;
		Update zlParameters Set 参数值=v_Val Where 系统 Is Null And 模块 Is Null And 私有=1 And 参数名='zlBrwFontSize';
	Exception When Others Then Null; End;

	Begin
		Select 键值 Into v_Val From zlClientParaList Where 类别='公共模块' And 目录='MDI' And 键名='菜单排列方式' And 方案号=v_方案号;
		Update zlParameters Set 参数值=v_Val Where 系统 Is Null And 模块 Is Null And 私有=1 And 参数名='zlMdiMenuArray';
	Exception When Others Then Null; End;

	Begin
		Select 键值 Into v_Val From zlClientParaList Where 类别='公共模块' And 目录='MDI' And 键名='MDI背景图片' And 方案号=v_方案号;
		Update zlParameters Set 参数值=v_Val Where 系统 Is Null And 模块 Is Null And 私有=1 And 参数名='zlMdiBackPic';
	Exception When Others Then Null; End;

	Begin
		Select 键值 Into v_Val From zlClientParaList Where 类别='公共模块' And 目录='MDI' And 键名='字体色' And 方案号=v_方案号;
		Update zlParameters Set 参数值=v_Val Where 系统 Is Null And 模块 Is Null And 私有=1 And 参数名='zlMdiFontColor';
	Exception When Others Then Null; End;

	Begin
		Select 键值 Into v_Val From zlClientParaList Where 类别='公共模块' And 目录='WINDOWS' And 键名='WIN背景图片' And 方案号=v_方案号;
		Update zlParameters Set 参数值=v_Val Where 系统 Is Null And 模块 Is Null And 私有=1 And 参数名='zlWinBackPic';
	Exception When Others Then Null; End;

	Begin
		Select 键值 Into v_Val From zlClientParaList Where 类别='公共模块' And 目录='WINDOWS' And 键名='字体色' And 方案号=v_方案号;
		Update zlParameters Set 参数值=v_Val Where 系统 Is Null And 模块 Is Null And 私有=1 And 参数名='zlWinFontColor';
	Exception When Others Then Null; End;

	Begin
		Select 键值 Into v_Val From zlClientParaList Where 类别='私有全局' And 目录 Is Null And 键名='提醒信息停留时间' And 方案号=v_方案号;
		Update zlParameters Set 参数值=v_Val Where 系统 Is Null And 模块 Is Null And 私有=1 And 参数名='自动消息停留时间';
	Exception When Others Then Null; End;

	Begin
		Select 键值 Into v_Val From zlClientParaList Where 类别='私有全局' And 目录 Is Null And 键名='导航台' And 方案号=v_方案号;
		Update zlParameters Set 参数值=v_Val Where 系统 Is Null And 模块 Is Null And 私有=1 And 参数名='导航台';
	Exception When Others Then Null; End;

	Begin
		Select 键值 Into v_Val From zlClientParaList Where 类别='私有全局' And 目录='常用功能' And 键名='系统' And 方案号=v_方案号;
		Select v_Val||'|'||键值 Into v_Val From zlClientParaList Where 类别='私有全局' And 目录='常用功能' And 键名='序号' And 方案号=v_方案号;
		Select v_Val||'|'||键值 Into v_Val From zlClientParaList Where 类别='私有全局' And 目录='常用功能' And 键名='图标' And 方案号=v_方案号;
		Select v_Val||'|'||键值 Into v_Val From zlClientParaList Where 类别='私有全局' And 目录='常用功能' And 键名='标题' And 方案号=v_方案号;
		Update zlParameters Set 参数值=v_Val Where 系统 Is Null And 模块 Is Null And 私有=1 And 参数名='常用功能模块';
	Exception When Others Then Null; End;

	Begin
		Select 键值 Into v_Val From zlClientParaList Where 类别='公共模块' And 目录='操作' And 键名='输入匹配' And 方案号=v_方案号;
		Update zlParameters Set 参数值=v_Val Where 系统 Is Null And 模块 Is Null And 私有=1 And 参数名='输入匹配';
	Exception When Others Then Null; End;

	Begin
		Select 键值 Into v_Val From zlClientParaList Where 类别='私有全局' And 目录 Is Null And 键名='输入法' And 方案号=v_方案号;
		Update zlParameters Set 参数值=v_Val Where 系统 Is Null And 模块 Is Null And 私有=1 And 参数名='输入法';
	Exception When Others Then Null; End;

	Begin
		Select 键值 Into v_Val From zlClientParaList Where 类别='私有全局' And 目录 Is Null And 键名='简码生成' And 方案号=v_方案号;
		Update zlParameters Set 参数值=v_Val Where 系统 Is Null And 模块 Is Null And 私有=1 And 参数名='简码方式';
	Exception When Others Then Null; End;

	Begin
		Select 键值 Into v_Val From zlClientParaList Where 类别='私有全局' And 目录 Is Null And 键名='关闭Windows' And 方案号=v_方案号;
		Update zlParameters Set 参数值=v_Val Where 系统 Is Null And 模块 Is Null And 私有=1 And 参数名='关闭Windows';
	Exception When Others Then Null; End;

	Begin
		Select 键值 Into v_Val From zlClientParaList Where 类别='私有全局' And 目录 Is Null And 键名='通知检查周期' And 方案号=v_方案号;
		Update zlParameters Set 参数值=v_Val Where 系统 Is Null And 模块 Is Null And 私有=1 And 参数名='邮件消息检查周期';
	Exception When Others Then Null; End;

	Begin
		Select 键值 Into v_Val From zlClientParaList Where 类别='私有全局' And 目录 Is Null And 键名='登录时检查通知新消息' And 方案号=v_方案号;
		Update zlParameters Set 参数值=v_Val Where 系统 Is Null And 模块 Is Null And 私有=1 And 参数名='登录检查邮件消息';
	Exception When Others Then Null; End;

	Begin
		Select 键值 Into v_Val From zlClientParaList Where 类别='私有模块' And 目录='zl9AppTool\frmMessageManager\Menu' And 键名='mnuViewShowAll状态' And 方案号=v_方案号;
		Update zlParameters Set 参数值=v_Val Where 系统 Is Null And 模块 Is Null And 私有=1 And 参数名='显示已读邮件';
	Exception When Others Then Null; End;
End;
/


--参数对象授权处理
--------------------------------------------------------------------------------------------------
Create Public Synonym zlParameters_ID for zlTools.zlParameters_ID
/
Create Public Synonym zlParameters for zlTools.zlParameters
/
Create Public Synonym zlUserParas for zlTools.zlUserParas
/
Create Public Synonym zl_Parameters_Update for zlTools.zl_Parameters_Update
/
Create Public Synonym zl_Parameters_Update_Batch for zlTools.zl_Parameters_Update_Batch
/
Grant Select On zlTools.zlParameters_ID to Public 
/
Grant Select On zlTools.zlParameters to Public 
/
Grant Select On zlTools.zlUserParas to Public 
/
Grant Execute On zlTools.zl_Parameters_Update to Public 
/
Grant Execute On zlTools.zl_Parameters_Update_Batch to Public 
/
Begin
	For r_User In(Select 所有者 From zlSystems)
	Loop
			Execute Immediate 'Grant Select,Insert,Update,Delete on zlTools.zlParameters to '||r_User.所有者||' With Grant Option';
			Execute Immediate 'Grant Select,Insert,Update,Delete on zlTools.zlUserParas to '||r_User.所有者||' With Grant Option';
			Execute Immediate 'Grant Execute on zlTools.zl_Parameters_Update to '||r_User.所有者||' With Grant Option';
			Execute Immediate 'Grant Execute on zlTools.zl_Parameters_Update_Batch to '||r_User.所有者||' With Grant Option';
	End Loop;
End;
/

